'****************************************************************
'
'					Capersville (Bally 1966)
'						by Scottacus
'					   September 2017
'				   Revision October 2017
'
'	Thanks to:
'	HauntFreaks for the excellent graphics work
'	BorgDog for the stripped down Gottlieb EM 4 player VPX table
'	Xenonph for his many hours of testing
'	Loserman76 for his coding help
'	Arngrim for his help with DOF
'
'	Basic DOF config
'		101 Left Flipper, 102 Right Flipper
'		103 Bumper1,  104 bumper2, 105 Bumper3
'		106 Left sling, 107 right sling
'		108 Knocker
'		109 Ball Release
'		110 Right Kicker, 111 Left Kicker
'		112 Ball In Shooter Lane
'		113 CodeZapper Kicker
'		114-116 Yellow Mushrooms
'		117-119 Blue, White & Red Mushrooms
'		121-124 Deep4Caper Capture spaces
'		140 DC4 Gate Lock
'		141 Left OutLane, 142 Right OutLane
'       143 Drain
'	 	150 credit light, 151 knocker Flasher, 152 Zip Flippers
'		153 Chime1-10s, 154 Chime2-100s, 155 Chime3-1000s
'		160 Left Saucer, 161 Right Saucer

'************************************************ Code Flow ***********************************************************
'									 EndGame
'										^
'		Start Game -> New Game -> Check Continue -> Release Ball -> Drain -> Score Bouns -> Advance Player -> Next Ball
'										|																		  |
'									     -------------------------------------------------------------------------
'**********************************************************************************************************************
'	Ball Control Subroutine developed by: rothbauerw
'		Press "c" during play to activate, the arrow keys control the ball
'******************************************************************************

' Thalamus 2018-07-19
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' No special SSF tweaks yet.

Option Explicit
Randomize

' Thalamus 2018-08-09 : Improved directional sounds
' !! NOTE : Table not verified yet !!

' Options
' Volume devided by - lower gets higher sound

Const VolDiv = 2000    ' Lower number, louder ballrolling/collition sound
Const VolCol = 10      ' Ball collition divider ( voldiv/volcol )

' The rest of the values are multipliers
'
'  .5 = lower volume
' 1.5 = higher volume

Const VolBump   = 2    ' Bumpers volume.
Const VolGates  = 1    ' Gates volume.
Const VolMetal  = 1    ' Metals volume.
Const VolRH     = 1    ' Rubber hits volume.
Const VolPo     = 1    ' Rubber posts volume.
Const VolPi     = 1    ' Rubber pins volume.
Const VolTarg   = 1    ' Targets volume.
Const VolKick   = 1    ' Kicker volume.
Const VolSpin   = 1.5  ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.

Const cGameName = "Capersville"
Const Ballsize = 50			'used by ball shadow routine

Dim Balls
Dim Replays
Dim Add1, Add10, Add100, Add1000
Dim HiSc
Dim MaxPlayers
Dim Players
Dim Player
Dim Credit
Dim Score(6)
Dim SReels(6)
Dim State
Dim Tilt
Dim TiltSens
Dim Target(9)
Dim BallinPlay
Dim MatchNumber
dim BallREnabled
dim RStep, LStep
Dim Rep(4)
Dim Rst
Dim EndGame
Dim Bell
Dim i,j, f, ii, Object, Light, x, y
Dim AwardCheck
Dim BStop, DC4Stop(4)
Dim Mush
Dim FreePlay
Dim SeaRay
Dim SeaRayBonus(4)
Dim CodeZapper
Dim DC4
Dim DC4Stage
Dim BallInLane
Dim SaucerCapture
Dim CapturedBalls
Dim DCCapturedBall
Dim BallMass
Dim BallHomeCheck
Dim Button
Dim CodeZapperValue
Dim HSArray, HSiArray
Dim Shift
Dim Launched
Dim CZLaunched
Dim STAT
Dim RelGateHit
Dim CodeZapperGateHit
Dim HSInitial0, HSInitial1, HSInitial2
Dim EnableInitialEntry
Dim HSi,HSx
Dim RoundHS, RoundHSPlayer
Dim BellRing
Dim ShowBallShadow
Dim ScoreMil, Score100K, Score10K, ScoreK, Score100, Score10, ScoreUnit
HSiArray = Array("HSi_0","HSi_1","HSi_2","HSi_3","HSi_4","HSi_5","HSi_6","HSi_7","HSi_8","HSi_9","HSi_10","HSi_11","HSi_12","HSi_13","HSi_14","HSi_15","HSi_16","HSi_17","HSi_18","HSi_19","HSi_20","HSi_21","HSi_22","HSi_23","HSi_24","HSi_25","HSi_26")
HSArray = Array("HS_0","HS_1","HS_2","HS_3","HS_4","HS_5","HS_6","HS_7","HS_8","HS_9","HS_Space","HS_Comma")
BallMass = (Ballsize^3)/125000
Dim BIP
Dim Options
Dim HSUx
Dim ReplayEB
Dim Chime
Dim Language
Dim Lang
Dim STATx
Dim FirstBallLaunched



On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

Sub Table1_init
	LoadEM
	MaxPlayers=4
	Replay(1)= 5800		'Set Value
	Replay(2)= 7200		'Set Value
	Replay(3)= 8700		'Set Value
	SeaRay = 23
	Balls = 5
	Set SReels(1) = ScoreReel1
	Set SReels(2) = ScoreReel2
	Set SReels(3) = ScoreReel3
	Set SReels(4) = ScoreReel4
	Player = 1
	LoadHighScore

'**** Uncomment these lines, run the table then recomment these lines ****
'************** to reset the high scores to lower values *****************
'	For HSUx = 0 to 4
'		HighScore(HSUx) = 2002 - (HSUx*2)
'		Score(HSUx) = 0
'	Next
'*************************************************************************

	If HighScore(0)="" Then HighScore(0)=4500
	If HighScore(1)="" Then HighScore(1)=4000
	If HighScore(2)="" Then HighScore(2)=3500
	If HighScore(3)="" Then HighScore(3)=3000
	If HighScore(4)="" Then HighScore(4)=2500

	If Initial(0,1) = "" Then
		Initial(0,1) = 19: Initial(0,2) = 5: Initial(0,3) = 13
		Initial(1,1) = 1: Initial(1,2) = 1: Initial(1,3) = 1
		Initial(2,1) = 2: Initial(2,2) = 2: Initial(2,3) = 2
		Initial(3,1) = 3: Initial(3,2) = 3: Initial(3,3) = 3
		Initial(4,1) = 4: Initial(4,2) = 4: Initial(4,3) = 4
	End If
	If Credit = "" Then Credit = 0
	If FreePlay = "" Then FreePlay = 1
	If Balls = "" Then Balls = 5
	If ReplayEB = "" Then ReplayEB = 1
	If ShowBallShadow = "" Then ShowBallShadow = 1
	If Chime ="" Then Chime = 0
	If STAT = "" Then STAT = 0
	If Language = "" Then Language = 0
	If MatchNumber = "" Then MatchNumber = 10

	SaveHS

	hstxt.text = HiSc
	UpdatePostIt

	If Language = 0 Then
		tilttxt.text = "TILT"
		For i = 1 to 4
			EVAL("Player" & i).text = "Player " & i
			EVAL("SeaRayTxt" & i).text = "SEA RAY"
		Next
		Credits.text = "Credits"
		CanPlay.text = "Can Play"
		BallnPlay.text ="Ball in Play"
		CodeZapperBox.text = "CODE ZAPPER"
	Else
		tilttxt.text = "GEKIPPT!"
		For i = 1 to 4
			EVAL("Player" & i).text = i & " Spieler"
			EVAL("SeaRayTxt" & i).text = "EXTRA-BONUS"
		Next
		Credits.text = "Kredite"
		CanPlay.text = "Kann Spielen"
		BallnPlay.text ="Kugel im Spiel"
		CodeZapperBox.text = "CODE-RAD"
	End If

	biptext.text = "0"
	Matchtxt.text = MatchNumber
	credittxt.text = Credit
	CodeZapperValue = 100
	CodeZapperTxt.text = CodeZapperValue
	BIP = 0

	For x = 1 to 4
		Score(x) = 0
	Next

	If Language = 0 Then
		Playfield2.visible = False
		wapron.image = "Apron1"
		For Lang = 24 to 29
			EVAL("Light" & Lang).image = "Capersville Slide"
		Next
	Else
		Playfield2.visible = True
		wapron.image = "Apron2"
		For Lang = 24 to 29
			EVAL("Light" & Lang).image = "Capersville German"
		Next
	End If

	If ShowBallShadow = 1 Then
		BallShadowUpdate.enabled = True
	Else
		BallShadowUpdate.enabled = False
	End If

	If B2SOn Then
		Controller.B2SSetCredits Credit
		Controller.B2SSetMatch 34, MatchNumber
		Controller.B2SSetData 35 + Language,1
		Controller.B2SSetData 33 + Language ,1
		Controller.B2SSetBallInPlay 32,0
		Controller.B2SSetData 30 + Language,0
		Controller.B2SSetData 50 + Language, 1
	End If

	If ShowDT = True Then
		For each object in backdropstuff
		Object.visible = 1
		Next
	End If

	If ShowDt = False Then
		For each object in backdropstuff
		Object.visible = 0
		Next
	End If

	For i = 1 to MaxPlayers
		SReels(i).setvalue(score(i))
	Next

	PlaySound "motor"
	Tilt = False
	State = False
	GameState
	DC4 = 0
	SaucerCapture = 0
	CodeZapper = 1
	BlockDC4Gate.isdropped = True
	If FreePlay = 0 Then
		CoinCard.image = Language & "FreeCoin" & Balls & "Ball0"
	Else
		CoinCard.image = Language & "FreeCoin" & Balls & "Ball1"
	End If
	If ReplayEB = 0 Then
		InstructCard.image = "InstructionCardReplay" & Language
	Else
		InstructCard.image = "InstructionCardEB" & Language
	End If
	If STAT = 1 Then
		For STATx = 1 to 3
		EVAL ("BumperCap10pt" & STATx).visible = 	True
		EVAL ("BumperCap" & STATx).visible = False
		Next
	Else
		For STATx = 1 to 3
		EVAL ("BumperCap10pt" & STATx).visible = 	False
		EVAL ("BumperCap" & STATx).visible = True
		Next
	End If

'***********Trough Ball Creation
	TroughSlot1.CreateSizedBallWithMass Ballsize/2, BallMass
	TroughSlot2.CreateSizedBallWithMass Ballsize/2, BallMass
	TroughSlot3.CreateSizedBallWithMass Ballsize/2, BallMass
End Sub

'***********KeyCodes
Sub Table1_KeyDown(ByVal keycode)

	If EnableInitialEntry = True Then EnterIntitals(keycode)

	If keycode=AddCreditKey Then
		PlaySoundAtVol "coinin", drain, 1
		coindelay.enabled=True
    End If

    If keycode=StartGameKey Then
		If EnableInitialEntry = False and OperatorMenu = 0 Then
			If FreePlay = 1 Then StartGame
			If FreePlay = 0 and Credit > 0 and Players < MaxPlayers and FirstBallLaunched = False Then
				If Credit < 1 Then DOF 150, DOFOff
				Credit = Credit - 1
				CreditTxt.text = Credit
				If B2SOn Then Controller.B2SSetCredits Credit
				StartGame
			End If
		End If
	End If

	If keycode = PlungerKey Then
		Plunger.PullBack
		PlaySoundAtVol "plungerpull", Plunger, 1
	End If

  If Tilt = False and State = True Then
	If keycode = LeftFlipperKey Then
		LeftFlipper.RotateToEnd
		LeftFlipper1.RotateToEnd
		PlaySoundAtVol SoundFXDOF("flipperup",101,DOFOn,DOFContactors), LeftFlipper, VolFlip
		PlaySoundAtVol "Buzz", LeftFlipper, VolFlip
	End If

	If keycode = RightFlipperKey Then
		RightFlipper.RotateToEnd
		RightFlipper1.RotateToEnd
		PlaySoundAtVol SoundFXDOF("flipperup",102,DOFOn,DOFContactors), RightFlipper, VolFlip
		PlaySoundAtVol "Buzz1", RightFlipper, VolFlip
	End If

	If keycode = LeftTiltKey Then
		Nudge 90, 2
		CheckTilt
	End If

	If keycode = RightTiltKey Then
		Nudge 270, 2
		CheckTilt
	End If

	If keycode = CenterTiltKey Then
		Nudge 0, 2
		CheckTilt
	End If
  End If

    If keycode = LeftFlipperKey and State = False and OperatorMenu = 0 and EnableInitialEntry = 0 Then
        OperatorMenuTimer.Enabled = true
    End If

    If keycode = LeftFlipperKey and State = False and OperatorMenu = 1 Then
		Options = Options + 1
        If Options = 8 then Options = 0
		TextBox1.text = options
		OptionMenu.visible = true
        playsound "target"
        Select Case (Options)
            Case 0:
                OptionMenu.image = "FreeCoin" & FreePlay
            Case 1:
                OptionMenu.image = "BallNumber" & Balls
			Case 2:
				OptionMenu.image = "ReplayEB" & ReplayEB
            Case 3:
                OptionMenu.image = "BallShadowOption" & ShowBallShadow
			Case 4:
				OptionMenu.image = "Chime" & Chime
			Case 5:
				OptionMenu.image = "Caps" & STAT
			Case 6:
				OptionMenu.image = "Playfield" & Language
			Case 7:
				OptionMenu.image = "SaveExit"
        End Select
    End If

    If keycode = RightFlipperKey and State = False and OperatorMenu = 1 Then
      PlaySound "metalhit2"
      Select Case (Options)
		Case 0:
            If FreePlay = 0 Then
                FreePlay = 1
              Else
                FreePlay = 0
            End If
            OptionMenu.image= "FreeCoin" & FreePlay
				If FreePlay = 0 Then
					CoinCard.image = Language & "FreeCoin" & Balls & "Ball0"
				Else
			CoinCard.image = Language & "FreeCoin" & Balls & "Ball1"
			End If
        Case 1:
            If Balls = 3 Then
                Balls = 5
              Else
                Balls = 3
            End If
            OptionMenu.image = "BallNumber" & Balls
			If FreePlay = 0 Then
				CoinCard.image = Language & "FreeCoin" & Balls & "Ball0"
			Else
				CoinCard.image = Language & "FreeCoin" & Balls & "Ball1"
			End If
		Case 2:
			If ReplayEB = 0 Then
				ReplayEB = 1
				InstructCard.image = "InstructionCardEB" & Language
			Else
				ReplayEB = 0
				InstructCard.image = "InstructionCardReplay" & Language
			End If
			OptionMenu.image = "ReplayEB" & ReplayEB
        Case 3:
            If ShowBallShadow = 0 Then
                ShowBallShadow = 1
              Else
                ShowBallShadow = 0
            End If
			OptionMenu.image = "BallShadowOption" & ShowBallShadow
			If ShowBallShadow = 1 Then
				BallShadowUpdate.enabled = True
			Else
				BallShadowUpdate.enabled = False
			End If
        Case 4:
            If Chime = 0 Then
                Chime= 1
				If B2SOn Then DOF 155,DOFPulse
              Else
                Chime = 0
				Playsound "Bell10"
            End If
			OptionMenu.image = "Chime" & Chime
		Case 5:
			If STAT = 0 Then
				STAT = 1
			Else
				STAT = 0
			End If
			If STAT = 1 Then
				For STATx = 1 to 3
				EVAL ("BumperCap10pt" & STATx).visible = 	True
				EVAL ("BumperCap" & STATx).visible = False
				Next
			Else
				For STATx = 1 to 3
				EVAL ("BumperCap10pt" & STATx).visible = 	False
				EVAL ("BumperCap" & STATx).visible = True
				Next
			End If
			OptionMenu.image = "Caps" & STAT
		Case 6:
			If Language = 0 Then
				Language = 10
				Playfield2.visible = True
				Wapron.image = "Apron2"
				For Lang = 24 to 29
					EVAL("Light" & Lang).image = "Capersville German"
				Next
				CoinCard.image = Language & "FreeCoin" & Balls & "Ball" & FreePlay
				InstructCard.image = "InstructionCardEB" & Language
				If B2SOn Then
					Controller.B2SSetData 50,0
					Controller.B2SSetData 60,1
					Controller.B2SSetTilt 33, 0
					Controller.B2SSetTilt 43, 1
					Controller.B2SSetGameOver 45, 1
					Controller.B2SSetGameOver 35, 0
				End If

				tilttxt.text = "GEKIPPT!"
				For i = 1 to 4
					EVAL("Player" & i).text = i & " Spieler"
					EVAL("SeaRayTxt" & i).text = "EXTRA-BONUS"
				Next
				Credits.text = "Kredite"
				CanPlay.text = "Kann Spielen"
				BallnPlay.text ="Kugel im Spiel"
				CodeZapperBox.text = "CODE-RAD"
				gamov.text = "SPIEL AUS!"
			Else
				Language = 0
				Playfield2.visible = False
				Wapron.image = "Apron1"
				For Lang = 24 to 29
					EVAL("Light" & Lang).image = "Capersville Slide"
				Next
				CoinCard.image = Language & "FreeCoin" & Balls & "Ball" & FreePlay
				InstructCard.image = "InstructionCardEB" & Language
				If B2SOn Then
					Controller.B2SSetData 50,1
					Controller.B2SSetData 60,0
					Controller.B2SSetTilt 33, 1
					Controller.B2SSetTilt 43, 0
					Controller.B2SSetGameOver 45, 0
					Controller.B2SSetGameOver 35, 1
				End If
				tilttxt.text = "TILT"
				For i = 1 to 4
					EVAL("Player" & i).text = "Player " & i
					EVAL("SeaRayTxt" & i).text = "SEA RAY"
				Next
				Credits.text = "Credits"
				CanPlay.text = "Can Play"
				BallnPlay.text ="Ball in Play"
				CodeZapperBox.text = "CODE ZAPPER"
				gamov.text = "GAME OVER"
			End If
			OptionMenu.image = "Playfield" & Language
        Case 7:
            OperatorMenu = 0
            SaveHS
			DynamicUpdatePostIt.enabled = 1
			OptionMenu.image = "FreeCoin" & FreePlay
            OptionMenu.visible = 0
			OptionsMenu.visible = 0
      End Select
		TextBox1.text = options
    End If

	If Keycode = MechanicalTilt Then
		Tilt = True
		If Language = 0 Then
			Tilttxt.text = "TILT"
		Else
			Tilttxt.text = "GEKIPPT!"
		End If
		If B2SOn Then Controller.B2SSetTilt 33 + Language, 1
		Playsound"Tilt"
		UnZipFlippers
		If SaucerCapture = True Then
			BIP = BIP + CapturedBalls
			SaucerCapture = False
			BlockDC4Gate.isdropped = True
			metal_uppergate.rotz = 0
			DC4Light.state = 0
			If B2SOn Then
				DOF 160, DOFOff
				DOF 161, DOFOff
			End If
		End If
		TurnOff
	End If

    If Keycode = 46 Then' C Key
        If contball = 1 Then
            contball = 0
          Else
            contball = 1
        End If
    End If

    If keycode = 48 Then 'B Key
        If bcboost = 1 Then
            bcboost = bcboostmulti
          Else
            bcboost = 1
        End If
    End If

    If keycode = 203 Then Cleft = 1' Left Arrow

    If keycode = 200 Then Cup = 1' Up Arrow

    If keycode = 208 Then Cdown = 1' Down Arrow

    If keycode = 205 Then Cright = 1' Right Arrow

'************************Start Of Test Keys****************************
'	If keycode = 30 Then
'		DynamicUpdatePostIt.enabled = 0
'		Score(1) = 1999
'		Position(1) = 2
'		Score(2) = 1997
'		Position(2) = 1
'		EndGame = 1
'		CheckContinue.enabled = 1
'	End If
'	If keycode = 35 Then WhiteMushroom


'	If keycode = 31 Then UnZipFlippers
'************************End Of Test Keys****************************
End Sub

Sub Table1_KeyUp(ByVal keycode)

	If keycode = PlungerKey Then
		Plunger.Fire
		PlaySoundAtVol "plunger", Plunger, 1
	End If

    If keycode = LeftFlipperKey Then
        OperatorMenuTimer.Enabled = False
    End If

   If Tilt = False and State = True Then
	If keycode = LeftFlipperKey Then
		LeftFlipper.RotateToStart
		LeftFlipper1.RotateToStart
		PlaySoundAtVol SoundFXDOF("flipperdown",101,DOFOff,DOFContactors), LeftFlipper, VolFlip
		StopSound "Buzz"
	End If

	If keycode = RightFlipperKey Then
		RightFlipper.RotateToStart
		RightFlipper1.RotateToStart
		PlaySoundAtVol SoundFXDOF("flipperdown",102,DOFOff,DOFContactors), RightFlipper, VolFlip
		StopSound "Buzz1"
	End If
   End if


   If keycode = 203 then Cleft = 0' Left Arrow

   If keycode = 200 then Cup = 0' Up Arrow

   If keycode = 208 then Cdown = 0' Down Arrow

   If keycode = 205 then Cright = 0' Right Arrow

End Sub

Sub FlipperTimer_timer()
	LFlip.RotY = LeftFlipper.CurrentAngle + 90
	RFlip.RotY = RightFlipper.CurrentAngle +90
	LFlip1.RotY = LeftFlipper.CurrentAngle + 60
	RFlip1.RotY = RightFlipper.CurrentAngle + 120
End Sub

'***********Operator Menu
Dim operatormenu
Sub OperatorMenuTimer_Timer
	Options = 0
    OperatorMenu = 1
    Displayoptions
End Sub

Sub DisplayOptions
	DynamicUpdatePostIt.enabled = 0
	UpdatePostIt
	Options = 0
    OptionsMenu.visible = True
    OptionMenu.visible = True
	OptionMenu.image = "FreeCoin" & FreePlay
End Sub

'***********Start Game
Sub StartGame
	If State = False Then
		BallinPlay = 1
		If B2SOn Then
			Controller.B2SSetCredits Credit
			Controller.B2SSetData 49, 1
			Controller.B2SSetBallinPlay 32, Ballinplay
			Controller.B2SSetPlayerup 30 + Language, 1
			Controller.B2SSetData 31, 1
			Controller.B2SSetGameOver 35 + Language, 0
			Controller.B2SSetData 55 + Language, 1
	'		Controller.B2SStopAnimation "Capersville Logo"
		End If
		Tilt = False
		State = True
		GameState
		DynamicUpdatePostIt.enabled = 0
		UpdatePostIt
		PlaySound "initialize"
		Players = 1
		CanPlayTxt.text = Players
		CodeZapperTxt.text = CodeZapperValue
		CodeZapperBlock.isdropped = False
		DC4Text.text = DC4
		SaucerText.text = SaucerCapture
		For x = 1 to 4
			Score(x) = 0
			If B2SOn Then controller.B2SSetScorePlayer x, Score(x)
			SReels(x).setvalue(0)
		Next
		If B2SOn Then controller.B2SSetScorePlayer 6, 0 'CodeZapper Reels
		Rst=0
		NewGame.enabled = True
	Else If  State = True and Players < MaxPlayers and BallinPlay = 1 Then
		Players = Players + 1
		CanPlayTxt.text = Players
		CreditTxt.text = Credit
		If B2SOn Then
			Controller.B2SSetCredits Credit
			Controller.B2SSetData 31, Players
		End If
		Playsound "cluper"
		End If
	End If
End Sub

'***************New Game
Sub NewGame_timer
	Player = 1
	For i = 1 to MaxPlayers
	    Score(i) = 0
		Rep(i) = 0
	Next
	If B2SOn Then
	  For i = 1 to MaxPlayers
		Controller.B2SSetScorePlayer i, score(i)
	  Next
	End If
    EndGame = 0
	UnZipFlippers
	ShootAgainLight.state = 0
	GameState
	CheckContinue.enabled = True
	NewGame.enabled = False
	For f = 1 to 3
		EVAL("Bumper"&f).hashitevent = 1
	Next
	For i = 1 to 4
		SeaRayBonus(i) = 0
		EVAL("SeaRayBonusTxt"&i).text = SeaRayBonus(i)
	Next
	If B2SOn Then Controller.B2SSetScorePlayer 6, 0 'SeaRay Reels are Player 6
	ResetSeaRay.enabled = True
	CodeZapperBlock.isdropped = False
	CodeZapperLight.state = 0
	Player1.intensityscale = 1.5
End Sub

'*************Check if Game Should Continue
Sub CheckContinue_timer
	If EndGame = 1 Then
		TurnOff
		Match
		State = False
		GameState
		DynamicUpdatePostIt.enabled = 1
		CanPlayTxt.text = 0
		SortScores
		CheckHighScores
		Players = 0
		SaveHS
		FirstBallLaunched = False
		If B2SOn Then
			Controller.B2SSetGameOver 35 + Language,1
			Controller.B2SSetballinplay 32, 0
			Controller.B2SSetPlayerUp 30 + Language, 0
			Controller.B2SSetcanplay 31, 0
			Controller.B2SSetData 49, 0
			Controller.B2SSetData 55 + Language, 0
			Controller.B2SStartAnimation "Capersville Logo"
		End If
		For each Light in GIlights:light.state = 0: Next
		CheckContinue.enabled = False
	Else
		If BallInLane = False and BIP = 0 Then
			ReleaseBall
		End If
	End If
	CheckContinue.enabled = False
End Sub

'***************Drain and Release Ball
Sub Drain_Hit()
	BIP = BIP - 1
	UpdateText
	PlaySoundAtVol "fx_drain", Drain, 1
	If B2SOn Then DOF 143, DOFPulse
	UpdateTrough
	UpdateText
End Sub

Sub ReleaseBall
	If BIP = 0 Then
		PlaysoundAtVol SoundFXDOF("kickerkick",107,DOFPulse,DOFContactors), TroughSlot1, 1
		TroughSlot1.kick 60, 20
		PlaysoundAtVol "metalhit_thin", Plunger, 1
		BIP = BIP + 1
		UpdateText
		UpdateTrough
		Launched = 0
		BIPText.text = BallinPlay
		If B2SOn Then Controller.B2SSetBallinPlay 32, BallinPlay
	End If
End Sub

'**************Advance Players
Sub ScoreBonus
	If ShootAgainLight.state = 0 Then
		If Players = 1 or Player = Players Then
			Player = 1
			UnZipFlippers
			EVAL("Player"&players).intensityscale = 1
			Player1.intensityscale = 1.5
		Else
			UnZipFlippers
			Player = Player + 1
			EVAL("Player" & (player-1)).intensityscale = 1
			EVAL("Player" & player).intensityscale = 1.5
		End If
		If B2SOn Then Controller.B2SSetPlayerup 30 + Language , Player
		NextBall
	Else
		ReleaseBall
	End If
End Sub

'*********************Next Ball
Sub NextBall
    If Tilt = True Then
	  For f = 1 to 3
		EVAL("Bumper"&f).hashitevent = 1
	  Next
      Tilt = False
      TiltTxt.text = " "
		If B2SOn Then
			Controller.B2SSetTilt 33 + Language,0
			Controller.B2SSetData 1, 1
		End If
    End If
	ResetSeaRay.enabled = True
	CodeZapperLight.state = 0
	CodeZapperBlock.isdropped = False
	If Player = 1 Then BallinPlay = BallinPlay + 1
	If BallinPlay > Balls Then
		PlaySound "GameOver"
		EndGame = 1
		CheckContinue.enabled = True
	  Else
		If State = True then
			CheckContinue.enabled = True
		End If
	End If
End Sub

'************Coin Handelers
Sub CoinDelay_timer
	AddCredit
	coindelay.enabled = False
End Sub

Sub AddCredit
	Credit = Credit+1
	DOF 150, DOFOn
	If Credit > 25 Then Credit = 25
	CreditTxt.text = Credit
	If B2SOn Then Controller.B2SSetCredits Credit
End Sub

'************Game State Check
Sub GameState
	If State = False Then
		For each light in GIlights:light.state = 0: Next
		flasher1.visible=0
		flasher2.visible=1
		metal_leftlower.blenddisablelighting = 0
		metal_leftmiddle.blenddisablelighting = 0
		metal_leftupper.blenddisablelighting = 0
		metal_leftlower.blenddisablelighting = 0
		metal_plungerplate.blenddisablelighting = 0
		metal_rightlower.blenddisablelighting = 0
		metal_rightupper.blenddisablelighting = 0
		If Language = 0 Then
			GamOv.text = "GAME OVER"
		Else
			GamOv.text = "SPIEL AUS!"
		End If
		If B2SOn Then Controller.B2SSetGameOver 35 + Language,1
		For i = 1 to 4
			EVAL ("Player" & i).intensityscale = 1
		Next
	Else
		For each Light in GIlights:Light.state = 1: Next
		flasher1.visible=1
		flasher2.visible=0
		metal_leftlower.blenddisablelighting = 0.25
		metal_leftmiddle.blenddisablelighting = 0.25
		metal_leftupper.blenddisablelighting = 0.25
		metal_leftlower.blenddisablelighting = 0.25
		metal_plungerplate.blenddisablelighting = 0.25
		metal_rightlower.blenddisablelighting = 0.25
		metal_rightupper.blenddisablelighting = 0.25
		GamOv.text = ""
		MatchTxt.text = ""
		TiltTxt.text = " "
		If B2SOn Then
			Controller.B2SSetTilt 33 + Language,0
			Controller.B2SSetMatch 34,0
			Controller.B2SSetGameOver 35 + Language,0
		End If
	End If
End Sub

'*************Trough based on nFozzy's
Dim Trough
Sub UpdateTrough()
	 Trough= 0
	UpdateTroughTimer.Interval = 400
	UpdateTroughTimer.Enabled = 1
End Sub

Sub UpdateTroughTimer_Timer()
	If TroughSlot1.BallCntOver = 0 Then TroughSlot2.kick 60, 8
	If TroughSlot2.BallCntOver = 0 Then TroughSlot3.kick 60, 8
	If TroughSlot3.BallCntOver = 0 Then Drain.kick 60, 30
	If Trough > 3 Then
		Me.Enabled = 0
		If BIP = 0 Then ScoreBonus
	End If
	Trough = Trough + 1
End Sub

'*************Ball in Launch Lane
Sub BallHome_hit
	BallREnabled = 1
	If B2SOn Then DOF 112, DOFOn
	RelGateHit = 0
	CodeZapperGateHit = 0
	Set ControlBall = ActiveBall
    contballinplay = True
End Sub

Sub BallHome_unhit
	If B2SOn Then DOF 112, DOFOff
End Sub

'******* For Ball Control Script
Sub EndControl_Hit()
    contballinplay = False
End Sub

'************Check if Ball Out of Launch Lane
Sub BallRel_hit
	If BallREnabled = 1 Then
		ShootAgainLight.state = 0
		BallREnabled = 0
		BallInLane = False
	End If
	me.timerenabled = 1
End Sub

'***************Shooter Lane Gate Animation
Sub BallRel_Timer
Pgate.rotz = Gate.CurrentAngle * .6
End Sub

'***************Match
Sub Match
   y = int(rnd(1)*9)
    MatchNumber = y
		If B2SOn Then
			If MatchNumber = 0 Then
				Controller.B2SSetMatch 10
			Else
				Controller.B2SSetMatch 34,MatchNumber
			End If
		End If
	MatchTxt.text = MatchNumber

	For i = 1 to Players
		MatchScoreTxt.text =  (Score(1) mod 10)
	 	If (MatchNumber) = (Score(i) mod 10) Then
			AddCredit
			If B2SOn Then
				DOF 108, DOFPulse
				DOF 151, DOFPulse
			Else
				PlaySound "Knock"
			End If
	    End If
    Next
End Sub

'***********Code Zapper Kicker
Sub ZapperKicker_hit
	me.timerenabled = 1
End Sub

Sub ZapperKicker_timer
	ZapperKicker.kick 0,40
	If B2SOn Then DOF 113, DOFPulse
	ZapperKicker.timerenabled = 0
End Sub

' Thalamus - maybe a rewrite is needed ?

'********** Bumpers
Sub Bumpers_hit(Index)
	If Tilt = False then AddScore 10
	Select Case (Index)
		Case 0: DOF 103, DOFPulse: PlaysoundAtVol "Bumper1", Bumper1, VolBump
		Case 1: DOF 104, DOFPulse: PlaysoundAtVol "Bumper2", Bumper2, VolBump
		Case 2: DOF 105, DOFPulse: PlaysoundAtVol "Bumper3", Bumper3, VolBump
	End Select
End Sub

'************** Mushroom Bumpers
Sub Mushrooms_Hit(Index)
	Mush = 0
	If Tilt = False Then
		Select Case (Index)
			Case 0: RubberY1.timerenabled = 1: AddScore 100: Playsound"MRCollision": SeaRayLights: CodeZapperCaper: UnZipFlippers: If B2SOn Then DOF 114, DOFPulse
			Case 1: RubberY2.timerenabled = 1: AddScore	100: Playsound"MRCollision":SeaRayLights: CodeZapperCaper: UnZipFlippers: If B2SOn Then DOF 115, DOFPulse
			Case 2: RubberY3.timerenabled = 1: AddScore	100: Playsound"MRCollision": SeaRayLights: CodeZapperCaper: UnZipFlippers: If B2SOn Then DOF 116, DOFPulse
			Case 3: RubberR.timerenabled = 1: AddScore 10: Playsound"MRCollision": ZipFlippers: If B2SOn Then DOF 119, DOFPulse
			Case 4:	RubberW.timerenabled = 1: AddScore 10: Playsound"MRCollision": WhiteMushroom: If B2SOn Then DOF 118, DOFPulse
			Case 5: RubberB.timerenabled = 1: AddScore 10: Playsound"MRCollision": Playsound"GateLock": CodeZapperBlock.isdropped = True: CodeZapperLight.state = 1: If B2SOn Then DOF 117, DOFPulse
		End Select
	End If
End Sub

Sub RubberY1_timer		'***** left yellow mushroom bumper
	Mush = Mush + 1
	Select Case Mush
		Case 1:	YellowCap1.transz=-7
		Case 2: YellowCap1.transz=-12
		Case 3: YellowCap1.transz=-7
		Case 4: YellowCap1.transz=0
		RubberY1.timerenabled=0
	End Select
End Sub

Sub RubberY2_timer		'***** mid yellow mushroom bumper
	Mush = Mush + 1
	Select Case Mush
		Case 1: YellowCap2.transz=-7
		Case 2: YellowCap2.transz=-12
		Case 3: YellowCap2.transz=-7
		Case 4: YellowCap2.transz=0
		RubberY2.timerenabled=0
	End Select
End Sub

Sub RubberY3_timer  	'***** right yellow mushroom bumper
	Mush = Mush+1
	Select Case Mush
		Case 1: YellowCap3.transz=-7
		Case 2: YellowCap3.transz=-12
		Case 3: YellowCap3.transz=-7
		Case 4: YellowCap3.transz=0
		RubberY3.timerenabled=0
	End Select
End Sub

Sub RubberR_timer		'***** Red mushroom bumper
	Mush = Mush + 1
	Select Case Mush
		Case 1:	RedCap.transz=-7
		Case 2: RedCap.transz=-12
		Case 3: RedCap.transz=-7
		Case 4: RedCap.transz=0
		RubberR.timerenabled=0
	End Select
End Sub

Sub RubberW_timer		'***** White mushroom bumper
	Mush = Mush + 1
	Select Case Mush
		Case 1:	WhiteCap.transz=-7
		Case 2: WhiteCap.transz=-12
		Case 3: WhiteCap.transz=-7
		Case 4: WhiteCap.transz=0
		RubberW.timerenabled=0
	End Select
End Sub

Sub RubberB_timer		'***** Blue mushroom bumper
	Mush = Mush + 1
	Select Case Mush
		Case 1:	BlueCap.transz=-7
		Case 2: BlueCap.transz=-12
		Case 3: BlueCap.transz=-7
		Case 4: BlueCap.transz=0
		RubberB.timerenabled=0
	End Select
End Sub

'**************One Point Sling Hits
Sub OnePointSlings_hit(Index)
	If Tilt = False Then AddScore 1
End Sub

'************** Slings
Sub LeftSlingShot_Slingshot
	PlaysoundAtVol SoundFXDOF("left_slingshot",106,DOFPulse,DOFContactors), sling2, 1
    LSling.Visible = 0
    LSling1.Visible = 1
    Sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
	AddScore 10
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3: LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4: LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub RubberRSling_hit
	AddScore 10
	If B2SOn Then DOF 107, DOFPulse
End Sub

'*********White Mushroom Hit
Sub WhiteMushroom
	If SaucerCapture = True Then
		LeftSaucer.timerenabled = 1
		RightSaucer.timerenabled = 1
		BIP = BIP + CapturedBalls
		CapturedBalls = 0
		UpdateText
		SaucerCapture = False
		BlockDC4Gate.isdropped = True
		metal_uppergate.rotz = 0
		DC4Light.state = 0
		If B2SOn Then
			DOF 160, DOFOff
			DOF 161, DOFOff
		End If
	Elseif DC4 = True Then
		Playsound"MetalHit_Thin"
		Select Case (DC4Stage)
		Case 1
			DC4Stop4.isdropped = False
			DC4Stop1.isdropped = True
			If B2SOn Then DOF 121,0:DOF 122,1
		Case 2
			DC4Stop1.isdropped = False
			DC4Stop2.isdropped = True
			If B2SOn Then DOF 122,0:DOF 123,1
		Case 3
			DC4Stop2.isdropped = False
			DC4Stop3.isdropped = True
			If B2SOn Then DOF 123,0:DOF 124,1
		Case 4
			DC4Stop3.isdropped = False
			DC4Stop4.isdropped = True
			If B2SOn Then DOF 124,0
			Deep4CaperEnd
		End Select
		DC4Stage = DC4Stage + 1
	End If
End Sub

'*********Saucer Hits
Sub RightSaucer_Hit
	PlaysoundAtVol"SaucerIn", RightSauser, 1
	If DC4 = True Then
		me.timerenabled = True
	ElseIf Tilt = True Then
		me.timerenabled = True
	Else
		CapturedBalls = CapturedBalls + 1
		BIP = BIP - 1
		UpdateText
		If B2SOn Then DOF 161, DOFOn
		SaucerCaptureCaper
	End If
End Sub

Sub LeftSaucer_Hit
	PlaysoundAtVol"SaucerIn", LeftSaucer, 1
	If DC4 = True Then
		me.timerenabled = True
	ElseIf Tilt = True Then
		me.timerenabled = True
	Else
		CapturedBalls = CapturedBalls + 1
		BIP = BIP - 1
		UpdateText
		If B2SOn Then DOF 160, DOFOn
		SaucerCaptureCaper
	End If
End Sub

Sub LeftSaucer_Timer
	LeftSaucer.Kick 170 + (3 * Rnd()), 12 + ( 3 * Rnd())
	PlaysoundAtVol SoundFXDOF("HoleKick",111,DOFPulse,DOFContactors), LeftSaucer, 1
	CapturedBalls = 0
	UpdateText
	If B2SOn Then DOF 160 , DOFOff
	me.timerenabled = 0
End Sub

Sub RightSaucer_Timer
	RightSaucer.Kick 186 + (3 * Rnd()), 12 + (3 * Rnd())
	PlaysoundAtVol SoundFXDOF("HoleKick",110,DOFPulse,DOFContactors), RightSaucer, 1
	CapturedBalls = 0
	UpdateText
	If B2SOn Then DOF 161, DOFOff
	me.timerenabled = 0
End Sub

'********** Triggers
sub LeftOutLaneTrigger_hit
	If Tilt = False Then
		AddScore 50
		CodeZapperCaper
		If B2SOn Then DOF 141, DOFPulse
	End If
end sub

sub RightOutLaneTrigger_hit
	If Tilt = False then
		AddScore 50
		CodeZapperCaper
		If B2SOn Then DOF 142, DOFPulse
	End If
end sub

Sub TopCenterTrigger_hit
	If Tilt = False Then AddScore 100
End Sub

Sub RedRollOver_hit
	Button = 0
	RedRollOverTimer.enabled = 1
	If Tilt = False Then
		AddScore 10
		CodeZapperCaper
		SeaRayLights
		CodeZapperBlock.isdropped = True
		CodeZapperLight.state = 1
	End If
End Sub

Sub RedRollOverTimer_Timer
	Button = Button +1
	Select Case Button
		Case 1: RedRollOverButton.transz=1
		Case 2: RedRollOverButton.transz=0
		Case 3: RedRollOverButton.transz=1
		Case 4: RedRollOverButton.transz=3
		RedRollOverTimer.enabled = 0
	End Select
End Sub

Sub CodeZapperTrigger_Hit
	If Tilt = False Then
	CZLaunched = CZLaunched + 1
		If CLng(CZLaunched) Mod 2 = 0 Then
			AddScore CodeZapperValue
			PlaySoundAtVol "HoleKick", CodeZapperTrigger, 1
		End If
	End If
	CodeZapperGateHit = 0
End Sub

Sub CodeZapperExit_hit
	me.timerenabled = 1
End Sub

Sub CodeZapperExit_Timer
	Pgate1.rotz = CodeZapperExitGate.currentAngle * .6
End Sub



Sub DC4Trigger1_Hit
	If DC4 = False Then
		BIP = BIP - 1
		metal_uppergate.rotz = -65
		BlockDC4Gate.isdropped = False
		UpdateText
		Deep4CaperStart
	End If
End Sub

Sub BallLaunchedTrigger_hit
	Launched = Launched + 1
	If CLng(Launched) Mod 2 > 0 Then
		If B2SOn Then DOF 109, DOFPulse
	End If
	FirstBallLaunched = True
End Sub

'***********Sea Ray Caper
' The original Capersville table has a block of 4 jacks and 4 plugs that can be connected to select a Special payout at
' any combination of 5-8 SeaRayBonuses.  The score cards on this table are set for a payout at 5 and 8 SeaRayBonuses.  You can simulate the
' original table's plugs by commenting or uncommenting any combination you like of 1288-1291.  The original Capersville also
' had the option of another jack further down the chain that allowed these outputs to either go to the Special or to go to
' a novelty mode that did not payout Specials but rather awarded the CodeZapper value instead.  To set the game to novelty
' mode comment out lines 1288-1292 and uncomment any combination of lines 1292-1295 that you like.
Sub SeaRayLights
	SeaRay = SeaRay + 1
	If SeaRay < 30 Then EVAL("Light"&SeaRay).state = 1
	If SeaRay = 29 Then
		ResetSeaRay.enabled = True
		SeaRayBonus(player) = SeaRayBonus(player) + 1
		PlaySound"ReelSoft"
		If SeaRayBonus(player) = 5 Then Special
'		If SeaRayBonus(player) = 6 Then Special
'		If SeaRayBonus(player) = 7 Then Special
		If SeaRayBonus(player) = 8 Then Special
'		If SeaRayBonus(player) = 5 Then AddScore CodeZapperValue 'Uncomment this an Comment out lines 1282-1285 for "Novelty" mode
'		If SeaRayBonus(player) = 6 Then AddScore CodeZapperValue
'		If SeaRayBonus(player) = 7 Then AddScore CodeZapperValue
'		If SeaRayBonus(player) = 8 Then AddScore CodeZapperValue
		EVAL("SeaRayBonustxt"&player).text = SeaRayBonus(player)
		If B2SOn Then Controller.B2SSetReel (Player+17), SeaRayBonus(Player)
	End If
End Sub

Sub ResetSeaRay_Timer
	For i = 24 to 29
		EVAL("Light"&i).state = 0
	Next
	SeaRay = 23
	ResetSeaRay.enabled = False
End Sub

'*******Code Zapper Caper
Sub CodeZapperCaper
	CodeZapper = CodeZapper + 1
	If CodeZapper = 4 Then CodeZapper = 1
	If B2SOn Then Controller.B2SSetReel 17, CodeZapper - 1
	PlaySound"ReelSoft"
	Select Case CodeZapper
		Case 1: CodeZapperValue = 100
		Case 2: CodeZapperValue = 300
		Case 3: CodeZapperValue = 500
	End Select
	CodeZapperTxt.text = CodeZapperValue
End Sub

'*******Deep 4 Caper
Sub Deep4CaperStart
	DC4 = True
	DC4Stage = 1
	If B2SOn Then
		DOF 121, 1
		DOF 160, 0
		DOF 161, 0
	End If
	Playsound SoundFXDOF("GateLock",140,DOFPulse,DOFcontactors)
	If BIP = 0 Then ShootAgainLight.state = 1
	DC4Text.text = 1
	SaucerText.text = 0
	DC4Light.state = 1
	BlockDC4Gate.isdropped = False
	RightSaucer.timerenabled = 1
	LeftSaucer.timerenabled = 1
	DCCapturedBall = 1
	UpdateText
	ReleaseBall
End Sub

Sub Deep4CaperEnd
	DC4 = False
	DC4Stage = 0
	Playsound"GateLock"
	ShootAgainLight.state = 1
	DC4Text.text = 0
	DC4Light.state = 0
	metal_uppergate.rotz = 0
	BlockDC4Gate.isdropped = True
	BallInLane = True
	RightSaucer.timerenabled = 0
	LeftSaucer.timerenabled = 0
	BIP = BIP + 1
	DCCapturedBall = 0
	UpdateText
End Sub

'*******Saucer Capture Caper
Sub SaucerCaptureCaper
	SaucerCapture = True
	PlaysoundAtVol SoundFXDOF("GateLock",140,DOFPulse,DOFcontactors), metal_uppergate, 1
	metal_uppergate.rotz = -60
	BlockDC4Gate.isdropped = False
	DC4Light.state = 1
	If BIP = 0 Then ShootAgainLight.state = 1
	SaucerText.Text = 1
	DC4Text.Text = 0
	UpdateText
	UpdateTrough
	ReleaseBall
End Sub

'********Zip Flippers
Sub ZipFlippers
	if LeftFlipper.enabled = True Then PlaysoundAtVol SoundFXDOF("FClose",152,DOFPulse,DOFcontactors), LeftFlipper, VolFlip
	LeftFlipper.enabled = False
	RightFlipper.enabled = False
	LeftFlipper1.enabled = True
	RightFlipper1.enabled = True
	LFlip.visible = False
	RFlip.visible = False
	LFlip1.visible = True
	RFlip1.visible = True
End Sub

Sub UnZipFlippers
	If LeftFlipper.enabled = False Then	PlaysoundAtVol SoundFXDOF("FClose",152,DOFPulse,DOFcontactors), LeftFlipper, VolFlip
	LeftFlipper.enabled = True
	RightFlipper.enabled = True
	LeftFlipper1.enabled = False
	RightFlipper1.enabled =	False
	LFlip.visible = True
	RFlip.visible = True
	LFlip1.visible = False
	RFlip1.visible = False
	PlaysoundAtVol SoundFXDOF("FClose",152,DOFPulse,DOFcontactors), LeftFlipper, VolFlip
End Sub

'********Special
Sub Special
	If B2SOn Then
		DOF 108, DOFPulse
		DOF 151, DOFPulse
	Else
		PlaySound "Knock"
	End If
	If ReplayEB = 1 Then
		ShootAgainLight.state = 1
	Else
		AddCredit
	End If
End Sub

'********Update Desktop Text
Sub UpdateText
	BIPtxt.text = BIP
	CapturedBallTxt.text = CapturedBalls
	DCCapturedBallTxt.text = DCCapturedBall
End Sub

'*********Scoring Routine
Sub AddScore(points)
	If Tilt = False Then
		If B2SOn Then Controller.B2SSetScorePlayer Player, Score(Player)
		If Points = 1 Then
			PlaySound "Reel1"
			If Chime = 0 Then
				Playsound  "Bell10"
			Else
				If B2SOn Then DOF 153,DOFPulse
			End If
			TotalUp 1
		End If

		If Points = 10 Then
			Playsound"Reel1"
			If Chime = 0 Then
				Playsound  "Bell100"
			Else
				If B2SOn Then DOF 154,DOFPulse
			End If
			TotalUp 10
		End If

		If Points = 50 Then: BellRing = 5 : BellTimer10.enabled = 1

		If Points = 100 Then
			Playsound"Reel1"
			If Chime = 0 Then
				Playsound  "Bell100"
			Else
				If B2SOn Then DOF 155,DOFPulse
			End If
			TotalUp 100
		End If

		If Points = 300 Then: BellRing = 3 : BellTimer100.enabled = 1
		If Points = 500 Then: BellRing = 5 : BellTimer100.enabled = 1
   End If
End Sub

Dim ReplayX, RepAwarded(5), Replay(7), ReplayText(3)

Sub TotalUp(Points)
	Score(Player) = Score(Player) + Points
	SReels(Player).addvalue(Points)
	If B2SOn Then Controller.B2SSetScorePlayer Player, Score(Player)
	For ReplayX = Rep(Player) +1 to 3
		If Score(Player) >= Replay(ReplayX) Then
			If ReplayEB = 0 Then
				AddCredit
			Else
				ShootAgainLight.state = 1
			End If

			If B2SOn Then
				DOF 108, DOFPulse
				DOF 151, DOFPulse
			Else
				PlaySound "Knock"
			End If
			Rep(Player) = Rep(Player) + 1
		End If
	Next
End Sub

'*************** Bell Timers
Sub BellTimer10_Timer
	If Chime = 0 Then
		Playsound "Bell10"
	Else
		If B2SOn Then DOF 154,DOFPulse
	End If
	Playsound "Reel1"
	BellRing = BellRing - 1
	TotalUp 10
	If BellRing < 1 Then
		BellTimer10.enabled = 0
	End If
End Sub

Sub BellTimer100_Timer
	If Chime = 0 Then
		Playsound  "Bell100"
	Else
		If B2SOn Then DOF 155,DOFPulse
	End If
	Playsound "Reel1"
	BellRing = BellRing - 1
	TotalUp 100
	If BellRing < 1 Then
		BellTimer100.enabled = 0
	End If
End Sub

'*********Tilt Check
Sub CheckTilt
	If Tilttimer.Enabled = True Then
		TiltSens = TiltSens + 1
		If TiltSens = 3 Then
			Tilt = True
			If Language = 0 Then
				tilttxt.text = "TILT"
			Else
				Tilttxt.text = "GEKIPPT"
			End If
			If B2SOn Then Controller.B2SSetTilt 33 + Language,1
			If B2SOn Then Controller.B2SSetdata 1, 0
			Playsound "Tilt"
			UnZipFlippers
			If SaucerCapture = True Then
				BIP = BIP + CapturedBalls
				SaucerCapture = False
				BlockDC4Gate.isdropped = True
				metal_uppergate.rotz = 0
				DC4Light.state = 0
				If B2SOn Then
					DOF 160, DOFOff
					DOF 161, DOFOff
				End If
			End If
			TurnOff
		End If
	Else
	 TiltSens = 0
	 Tilttimer.Enabled = True
	End If
End Sub

Sub Tilttimer_Timer()
	Tilttimer.Enabled = False
End Sub


'***********Ball Shadow Update
Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5)

Sub BallShadowUpdate_timer()
    Dim BOT, b
    BOT = GetBalls
    ' hide shadow of deleted balls
    If UBound(BOT)<(tnob-1) Then
        For b = (UBound(BOT) + 1) to (tnob-1)
            BallShadow(b).visible = 0
        Next
    End If
    ' exit the Sub if no balls on the table
    If UBound(BOT) = -1 Then Exit Sub
    ' render the shadow for each ball
    For b = 0 to UBound(BOT)
        If BOT(b).X < Table1.Width/2 Then
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) + 13
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) - 13
        End If
        ballShadow(b).Y = BOT(b).Y + 10
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
End Sub

'***********Turn Off
Sub TurnOff
	For i= 1 to 3
		EVAL("Bumper"&i).hashitevent = 0
	Next
  	LeftFlipper.RotateToStart
	LeftFlipper1.RotateToStart
	StopSound "Buzz"
	If B2SOn Then DOF 101, DOFOff
	RightFlipper.RotateToStart
	RightFlipper1.RotateToStart
	StopSound "Buzz1"
	If B2SOn Then DOF 102, DOFOff
	If Tilt = True Then
		LeftSaucer.timerenabled = 1
		RightSaucer.timerenabled = 1
	End If
End Sub

'************************************************************************
'                         Ball Control
'************************************************************************

Dim Cup, Cdown, Cleft, Cright, contball, contballinplay, ControlBall, bcboost
Dim bcvel, bcyveloffset, bcboostmulti

bcboost = 1 'Do Not Change - default setting
bcvel = 4 'Controls the speed of the ball movement
bcyveloffset = -0.01 'Offsets the force of gravity to keep the ball from drifting vertically on the table, should be negative
bcboostmulti = 3 'Boost multiplier to ball veloctiy (toggled with the B key)

Sub BallControl_Timer()
    If Contball and ContBallInPlay then
        If Cright = 1 Then
            ControlBall.velx = bcvel*bcboost
          ElseIf Cleft = 1 Then
            ControlBall.velx = - bcvel*bcboost
          Else
            ControlBall.velx=0
        End If
        If Cup = 1 Then
            ControlBall.vely = -bcvel*bcboost
          ElseIf Cdown = 1 Then
            ControlBall.vely = bcvel*bcboost
          Else
            ControlBall.vely= bcyveloffset
        End If
    End If
End Sub

'************************************
' What you need to add to your table
'************************************

' a timer called RollingTimer. With a fast interval, like 10
' one collision sound, in this script is called fx_collide
' as many sound files as max number of balls, with names ending with 0, 1, 2, 3, etc
' for ex. as used in this script: fx_ballrolling0, fx_ballrolling1, fx_ballrolling2, fx_ballrolling3, etc


'******************************************
' Explanation of the rolling sound routine
'******************************************

' sounds are played based on the ball speed and position

' the routine checks first for deleted balls and stops the rolling sound.

' The For loop goes through all the balls on the table and checks for the ball speed and
' if the ball is on the table (height lower than 30) then then it plays the sound
' otherwise the sound is stopped, like when the ball has stopped or is on a ramp or flying.

' The sound is played using the VOL, PAN and PITCH functions, so the volume and pitch of the sound
' will change according to the ball speed, and the PAN function will change the stereo position according
' to the position of the ball on the table.


'**************************************
' Explanation of the collision routine
'**************************************

' The collision is built in VP.
' You only need to add a Sub OnBallBallCollision(ball1, ball2, velocity) and when two balls collide they
' will call this routine. What you add in the sub is up to you. As an example is a simple Playsound with volume and paning
' depending of the speed of the collision.

Sub a_Pins_Hit (idx)
	PlaySound "pinhit_low", 0, Vol(ActiveBall)*VolPi, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub a_Targets_Hit (idx)
	PlaySound "target", 0, Vol(ActiveBall)*VolTarg, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub a_Metals_Thin_Hit (idx)
	PlaySound "metalhit_thin", 0, Vol(ActiveBall)*VolMetal, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub a_Metals_Medium_Hit (idx)
	PlaySound "metalhit_medium", 0, Vol(ActiveBall)*VolMetal, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub a_Metals2_Hit (idx)
	PlaySound "metalhit2", 0, Vol(ActiveBall)*VolMetal, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub a_Gates_Hit (idx)
	PlaySound "gate4", 0, Vol(ActiveBall)*VolGates, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinner_Spin
	PlaySoundAtVol "fx_spinner", Spinner, VolSpin
End Sub

Sub a_Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RubberWheel_hit
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End sub

Sub a_Posts_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 16 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolPo, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 16 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub

Sub LeftFlipper_Collide(parm)
 	RandomSoundFlipper()
End Sub

Sub RightFlipper_Collide(parm)
 	RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub

'***************Bubble Sort
Dim TempScore(2), TempPos(3), Position(5)
Dim BSx, BSy
'Scores are sorted high to low with Position being the player's number
Sub SortScores
	For BSx = 1 to 4
		Position(BSx) = BSx
	Next
	For BSx = 1 to 4
		For BSy = 1 to 3
			If Score(BSy) < Score(BSy+1) Then
				TempScore(1) = Score(BSy+1)
				TempPos(1) = Position(BSy+1)
				Score(BSy+1) = Score(BSy)
				Score(BSy) = TempScore(1)
				Position(BSy+1) = Position(BSy)
				Position(BSy) = TempPos(1)
			End If
		Next
	Next
End Sub

'*************Check for High Scores
Dim HighScore(5), ActiveScore(5), HS, CHx, CHy, CHz, CHix, TempI(4), TempI2(4), Flag
'goes through the 5 high scores one at a time and compares them to the player's scores high to
'if a player's score is higher it marks that postion with ActiveScore(x) and moves all of the other
'	high scores down by one along with the high score's player initials
'	also clears the new high score's initials for entry later
Sub CheckHighScores
	For HS = 1 to 4     							'look at 4 player scores
		For CHy = 0 to 4   					    	'look at all 5 saved high scores
			If Score(HS) > HighScore(CHy) Then
				Flag = Flag + 1						'flag to show how many high scores needs replacing
				TempScore(1) = HighScore(CHy)
				HighScore(CHy) = Score(HS)
				ActiveScore(HS) = CHy				'ActiveScore(x) is the high score being modified with x=1 the largest and x=4 the smallest
				For CHix = 1 to 3					'set initals to blank and make temporary initials = to intials being modifed so they can move down one high score
					TempI(Chix) = Initial(CHy,CHix)
					Initial(CHy,CHix) = 0
				Next

				If CHy < 4 Then						'check if not on lowest high score for overflow error prevention
					For CHz = CHy+1 to 4			'set as high score one more than score being modifed (CHy+1)
						TempScore(2) = HighScore(CHz)	'set a temporaray high score for the high score one higher than the one being modified
						HighScore(CHz) = TempScore(1)	'set this score to the one being moved
						TempScore(1) = TempScore(2)		'reassign TempScore(1) to the next higher high score for the next go around
						For CHix = 1 to 3
							TempI2(CHix) = Initial(CHz,CHix)	'make a new set of temporary initials
						Next
						For CHix = 1 to 3
							Initial(CHz,CHix) = TempI(Chix)		'set the initials to the set being moved
							TempI(CHix) = TempI2(CHix)			'reassign the initials for the next go around
						Next
					Next
				End If
				CHy = 4								'if this loop was accessed set CHy to 4 to get out of the loop
			End If
		Next
	Next
'	Goto Initial Entry
		HSi = 1			'go to the first initial for entry
		HSx = 1			'make the displayed inital be "A"
		If Flag > 0 Then	'Flag 0 when all scores are updated so leave subroutine and reset variables
			ShowScore
			PlayerEntry.visible = 1
			PlayerEntry.image = "Player" & Position(Flag)
			TextBox3.text = ActiveScore(Flag) 'tells which high score is being entered
			TextBox2.text = Flag
			TextBox1.text =  Position(Flag) 'tells which player is entering values
			Initial(ActiveScore(Flag),1) = 1	'make first inital "A"
			For CHy = 2 to 3
				Initial(ActiveScore(Flag),CHy) = 0	'set other two to " "
			Next
			For CHy = 1 to 3
				EVAL("Initial" & CHy).image = HSiArray(Initial(ActiveScore(Flag),CHy))		'display the initals on the tape
			Next
			InitialTimer1.enabled = 1		'flash the first initial
			DynamicUpdatePostIt.enabled = 0		'stop the scrolling intials timer
			If B2SOn Then
				DOF 108,DOFPulse
			Else
				PlaySound "Knock"
			End If
			EnableInitialEntry = True
		End If
End Sub

'************Enter Initials Keycode Subroutine
Dim Initial(6,5)

Sub EnterIntitals(keycode)
		If KeyCode = LeftFlipperKey Then
			HSx = HSx - 1						'HSx is the inital to be displayed A-Z plus " "
			If HSx < 0 Then HSx = 26
			If HSi < 4 Then EVAL("Initial" & HSi).image = HSiArray(HSx)		'HSi is which of the three intials is being modified
			PlaySound "metalhit_thin"
		End If
		If keycode = RightFlipperKey Then
			HSx = HSx + 1
			If HSx > 26 Then HSx = 0
			If HSi < 4 Then EVAL("Initial"& HSi).image = HSiArray(HSx)
			PlaySound "metalhit_thin"
		End If
		If keycode = StartGameKey Then
			If HSi < 3 Then									'if not on the last initial move on to the next intial
				EVAL("Initial" & HSi).image = HSiArray(HSx)	'display the initial
				Initial(ActiveScore(Flag), HSi) = HSx		'save the inital
				EVAL("InitialTimer" & HSi).enabled = 0		'turn that inital's timer off
				EVAL("Initial" & HSi).visible = 1			'make the initial not flash but be turn on
				Initial(ActiveScore(Flag),HSi + 1) = HSx	'move to the next initial and make it the same as the last initial
				EVAL("Initial" & HSi +1).image = HSiArray(HSx)	'display this intial
'				y = 1
				EVAL("InitialTimer" & HSi + 1).enabled = 1	'make the new intial flash
				HSi = HSi + 1								'increment HSi
			Else										'if on the last initial then get ready to exit the subroutine
				Initial3.visible = 1					'make the intial visible
				InitialTimer3.enabled = 0				'shut off the flashing
				Initial(ActiveScore(Flag),3) = HSx		'set last initial
				InitialEntry							'exit subroutine
			End If
		End If
End Sub

'************Update Initials and see if more scores need to be updated
Dim EIx
Sub InitialEntry
	If  Chime = 0 Then
		PlaySound "Bell10"
	Else
		If B2SOn Then DOF 155,DOFPulse
	End If
	Flag = Flag - 1
	TextBox2.text = Flag
	HSi = 1
	If Flag = 0 Then 					'exit high score entry mode and reset variables
		Players = 0
		For EIx = 1 to 4
			ActiveScore(EIx) = 0
			Position(EIx) = 0
		Next
		PlayerEntry.visible = 0
		ScoreUpdate = 0						'go to the highest score
		UpdatePostIt						'display that score
		HighScoreDelay.enabled = 1
	Else
		ShowScore
		PlayerEntry.image = "Player" & Position(Flag)
		TextBox3.text = ActiveScore(Flag) 	'tells which high score is being entered
		TextBox2.text = Flag
		TextBox1.text =  Position(Flag) 	'tells which player is entering values
		Initial(ActiveScore(Flag),1) = 1	'set the first initial to "A"
		For CHy = 2 to 3
			Initial(ActiveScore(Flag),CHy) = 0	'set the other two to " "
		Next
		For CHy = 1 to 3
			EVAL("Initial" & CHy).image = HSiArray(Initial(ActiveScore(Flag),CHy))	'display the intials
		Next
		HSx = 1							'go to the letter "A"
		InitialTimer1.enabled = 1		'flash the first intial
	End If
End Sub

'************Delay to prevent start button push for last initial from starting game Update
Sub HighScoreDelay_timer
	HighScoreDelay.enabled = 0
	EnableInitialEntry = False
	SaveHS
	DynamicUpdatePostIt.enabled = 1		'turn scrolling high score back on
End Sub

'************Flash Initials Timers
Sub InitialTimer1_Timer
	y = y + 1
	If y > 1 Then y = 0
	If y = 0 Then
		Initial1.visible = 1
	Else
		Initial1.visible = 0
	End If
End Sub

Sub InitialTimer2_Timer
	y = y + 1
	If y > 1 Then y = 0
	If y = 0 Then
		Initial2.visible = 1
	Else
		Initial2.visible = 0
	End If
End Sub

Sub InitialTimer3_Timer
	y = y + 1
	If y > 1 Then y = 0
	If y = 0 Then
		Initial3.visible = 1
	Else
		Initial3.visible = 0
	End If
End Sub

'************Save Scores
Sub SaveHS
    savevalue "Capersville", "HighScore(0)", HighScore(0)
    savevalue "Capersville", "HighScore(1)", HighScore(1)
    savevalue "Capersville", "HighScore(2)", HighScore(2)
    savevalue "Capersville", "HighScore(3)", HighScore(3)
    savevalue "Capersville", "HighScore(4)", HighScore(4)
    savevalue "Capersville", "Initial(0,1)", Initial(0,1)
    savevalue "Capersville", "Initial(0,2)", Initial(0,2)
    savevalue "Capersville", "Initial(0,3)", Initial(0,3)
    savevalue "Capersville", "Initial(1,1)", Initial(1,1)
    savevalue "Capersville", "Initial(1,2)", Initial(1,2)
    savevalue "Capersville", "Initial(1,3)", Initial(1,3)
    savevalue "Capersville", "Initial(2,1)", Initial(2,1)
    savevalue "Capersville", "Initial(2,2)", Initial(2,2)
    savevalue "Capersville", "Initial(2,3)", Initial(2,3)
    savevalue "Capersville", "Initial(3,1)", Initial(3,1)
    savevalue "Capersville", "Initial(3,2)", Initial(3,2)
    savevalue "Capersville", "Initial(3,3)", Initial(3,3)
    savevalue "Capersville", "Initial(4,1)", Initial(4,1)
    savevalue "Capersville", "Initial(4,2)", Initial(4,2)
    savevalue "Capersville", "Initial(4,3)", Initial(4,3)
    savevalue "Capersville", "Credit", Credit
    savevalue "Capersville", "FreePlay", FreePlay
	savevalue "Capersville", "Balls", Balls
	savevalue "Capersville", "ReplayEB", ReplayEB
	savevalue "Capersville", "ShowBallShadow", ShowBallShadow
	savevalue "Capersville", "Chime", Chime
	savevalue "Capersville", "STAT", STAT
	savevalue "Capersville", "MatchNumber", MatchNumber
	savevalue "Capersville", "Language", Language
End Sub

'*************Load Scores
Sub LoadHighScore
    dim temp
    temp = LoadValue("Capersville", "HighScore(0)")
    If (temp <> "") then HighScore(0) = CDbl(temp)
    temp = LoadValue("Capersville", "HighScore(1)")
    If (temp <> "") then HighScore(1) = CDbl(temp)
    temp = LoadValue("Capersville", "HighScore(2)")
    If (temp <> "") then HighScore(2) = CDbl(temp)
    temp = LoadValue("Capersville", "HighScore(3)")
    If (temp <> "") then HighScore(3) = CDbl(temp)
    temp = LoadValue("Capersville", "HighScore(4)")
    If (temp <> "") then HighScore(4) = CDbl(temp)
    temp = LoadValue("Capersville", "Initial(0,1)")
    If (temp <> "") then Initial(0,1) = CDbl(temp)
    temp = LoadValue("Capersville", "Initial(0,2)")
    If (temp <> "") then Initial(0,2) = CDbl(temp)
    temp = LoadValue("Capersville", "Initial(0,3)")
    If (temp <> "") then Initial(0,3) = CDbl(temp)
    temp = LoadValue("Capersville", "Initial(1,1)")
    If (temp <> "") then Initial(1,1) = CDbl(temp)
    temp = LoadValue("Capersville", "Initial(1,2)")
    If (temp <> "") then Initial(1,2) = CDbl(temp)
    temp = LoadValue("Capersville", "Initial(1,3)")
    If (temp <> "") then Initial(1,3) = CDbl(temp)
    temp = LoadValue("Capersville", "Initial(2,1)")
    If (temp <> "") then Initial(2,1) = CDbl(temp)
    temp = LoadValue("Capersville", "Initial(2,2)")
    If (temp <> "") then Initial(2,2) = CDbl(temp)
    temp = LoadValue("Capersville", "Initial(2,3)")
    If (temp <> "") then Initial(2,3) = CDbl(temp)
    temp = LoadValue("Capersville", "Initial(3,1)")
    If (temp <> "") then Initial(3,1) = CDbl(temp)
    temp = LoadValue("Capersville", "Initial(3,2)")
    If (temp <> "") then Initial(3,2) = CDbl(temp)
    temp = LoadValue("Capersville", "Initial(3,3)")
    If (temp <> "") then Initial(3,3) = CDbl(temp)
    temp = LoadValue("Capersville", "Initial(4,1)")
    If (temp <> "") then Initial(4,1) = CDbl(temp)
    temp = LoadValue("Capersville", "Initial(4,2)")
    If (temp <> "") then Initial(4,2) = CDbl(temp)
    temp = LoadValue("Capersville", "Initial(4,3)")
    If (temp <> "") then Initial(4,3) = CDbl(temp)
	temp = LoadValue("Capersville", "credit")
    If (temp <> "") then Credit = CDbl(temp)
	temp = LoadValue("Capersville", "FreePlay")
    If (temp <> "") then FreePlay = CDbl(temp)
    temp = LoadValue("Capersville", "balls")
    If (temp <> "") then Balls = CDbl(temp)
    temp = LoadValue("Capersville", "ReplayEB")
    If (temp <> "") then ReplayEB = CDbl(temp)
    temp = LoadValue("Capersville", "ShowBallShadow")
    If (temp <> "") then ShowBallShadow = CDbl(temp)
    temp = LoadValue("Capersville", "Chime")
    If (temp <> "") then Chime = CDbl(temp)
    temp = LoadValue("Capersville", "STAT")
    If (temp <> "") then STAT = CDbl(temp)
    temp = LoadValue("Capersville", "MatchNumber")
    If (temp <> "") then MatchNumber = CDbl(temp)
    temp = LoadValue("Capersville", "Language")
    If (temp <> "") then Language = CDbl(temp)
End Sub

'***************Static Post It Note Update
Dim  HSy
Sub UpdatePostIt
	ScoreMil = Int(HighScore(0)/1000000)
	Score100K = Int( (HighScore(0) - (ScoreMil*1000000) ) / 100000)
	Score10K = Int( (HighScore(0) - (ScoreMil*1000000) - (Score100K*100000) ) / 10000)
	ScoreK = Int( (HighScore(0) - (ScoreMil*1000000) - (Score100K*100000) - (Score10K*10000) ) / 1000)
	Score100 = Int( (HighScore(0) - (ScoreMil*1000000) - (Score100K*100000) - (Score10K*10000) - (ScoreK*1000) ) / 100)
	Score10 = Int( (HighScore(0) - (ScoreMil*1000000) - (Score100K*100000) - (Score10K*10000) - (ScoreK*1000) - (Score100*100) ) / 10)
	ScoreUnit = (HighScore(0) - (ScoreMil*1000000) - (Score100K*100000) - (Score10K*10000) - (ScoreK*1000) - (Score100*100) - (Score10*10) )

	Pscore6.image = HSArray(ScoreMil):If HighScore(0) < 1000000 Then PScore6.image = HSArray(10)
	Pscore5.image = HSArray(Score100K):If HighScore(0) < 100000 Then PScore5.image = HSArray(10)
	PScore4.image = HSArray(Score10K):If HighScore(0) < 10000 Then PScore4.image = HSArray(10)
	PScore3.image = HSArray(ScoreK):If HighScore(0) < 1000 Then PScore3.image = HSArray(10)
	PScore2.image = HSArray(Score100):If HighScore(0) < 100 Then PScore2.image = HSArray(10)
	PScore1.image = HSArray(Score10):If HighScore(0) < 10 Then PScore1.image = HSArray(10)
	PScore0.image = HSArray(ScoreUnit):If HighScore(0) < 1 Then PScore0.image = HSArray(10)
	If HighScore(0) < 1000 Then
		PComma.image = HSArray(10)
	Else
		PComma.image = HSArray(11)
	End If
	If HighScore(0) < 1000000 Then
		PComma1.image = HSArray(10)
	Else
		PComma1.image = HSArray(11)
	End If
	If HighScore(0) >999999 Then Shift = 0:PComma.transx = 0
	If HighScore(0) < 1000000 Then Shift = 1:PComma.transx = -10
	If HighScore(0) < 100000 Then Shift = 2:PComma.transx = -20
	If HighScore(0) < 10000 Then Shift = 3:PComma.transx = -30
	For HSy = 0 to 6
		EVAL("Pscore" & HSy).transx = (-10 * Shift)
	Next
	Initial1.image = HSiArray(Initial(0,1))
	Initial2.image = HSiArray(Initial(0,2))
	Initial3.image = HSiArray(Initial(0,3))
End Sub

'***************Show Current Score
Sub ShowScore
	ScoreMil = Int(HighScore(ActiveScore(Flag))/1000000)
	Score100K = Int( (HighScore(ActiveScore(Flag)) - (ScoreMil*1000000) ) / 100000)
	Score10K = Int( (HighScore(ActiveScore(Flag)) - (ScoreMil*1000000) - (Score100K*100000) ) / 10000)
	ScoreK = Int( (HighScore(ActiveScore(Flag)) - (ScoreMil*1000000) - (Score100K*100000) - (Score10K*10000) ) / 1000)
	Score100 = Int( (HighScore(ActiveScore(Flag)) - (ScoreMil*1000000) - (Score100K*100000) - (Score10K*10000) - (ScoreK*1000) ) / 100)
	Score10 = Int( (HighScore(ActiveScore(Flag)) - (ScoreMil*1000000) - (Score100K*100000) - (Score10K*10000) - (ScoreK*1000) - (Score100*100) ) / 10)
	ScoreUnit = (HighScore(ActiveScore(Flag)) - (ScoreMil*1000000) - (Score100K*100000) - (Score10K*10000) - (ScoreK*1000) - (Score100*100) - (Score10*10) )

	Pscore6.image = HSArray(ScoreMil):If HighScore(ActiveScore(Flag)) < 1000000 Then PScore6.image = HSArray(10)
	Pscore5.image = HSArray(Score100K):If HighScore(ActiveScore(Flag)) < 100000 Then PScore5.image = HSArray(10)
	PScore4.image = HSArray(Score10K):If HighScore(ActiveScore(Flag)) < 10000 Then PScore4.image = HSArray(10)
	PScore3.image = HSArray(ScoreK):If HighScore(ActiveScore(Flag)) < 1000 Then PScore3.image = HSArray(10)
	PScore2.image = HSArray(Score100):If HighScore(ActiveScore(Flag)) < 100 Then PScore2.image = HSArray(10)
	PScore1.image = HSArray(Score10):If HighScore(ActiveScore(Flag)) < 10 Then PScore1.image = HSArray(10)
	PScore0.image = HSArray(ScoreUnit):If HighScore(ActiveScore(Flag)) < 1 Then PScore0.image = HSArray(10)
	If HighScore(ActiveScore(Flag)) < 1000 Then
		PComma.image = HSArray(10)
	Else
		PComma.image = HSArray(11)
	End If
	If HighScore(ActiveScore(Flag)) < 1000000 Then
		PComma1.image = HSArray(10)
	Else
		PComma1.image = HSArray(11)
	End If
	If HighScore(ActiveScore(Flag)) >999999 Then Shift = 0:PComma.transx = 0
	If HighScore(ActiveScore(Flag)) < 1000000 Then Shift = 1:PComma.transx = -10
	If HighScore(ActiveScore(Flag)) < 100000 Then Shift = 2:PComma.transx = -20
	If HighScore(ActiveScore(Flag)) < 10000 Then Shift = 3:PComma.transx = -30
	For HSy = 0 to 6
		EVAL("Pscore" & HSy).transx = (-10 * Shift)
	Next
	Initial1.image = HSiArray(Initial(ActiveScore(Flag),1))
	Initial2.image = HSiArray(Initial(ActiveScore(Flag),2))
	Initial3.image = HSiArray(Initial(ActiveScore(Flag),3))
End Sub


'***************Dynamic Post It Note Update
Dim ScoreUpdate, DHSx
Sub DynamicUpdatePostIt_Timer
	TextBox1.text = ScoreUpdate
	ScoreMil = Int(HighScore(ScoreUpdate)/1000000)
	Score100K = Int( (HighScore(ScoreUpdate) - (ScoreMil*1000000) ) / 100000)
	Score10K = Int( (HighScore(ScoreUpdate) - (ScoreMil*1000000) - (Score100K*100000) ) / 10000)
	ScoreK = Int( (HighScore(ScoreUpdate) - (ScoreMil*1000000) - (Score100K*100000) - (Score10K*10000) ) / 1000)
	Score100 = Int( (HighScore(ScoreUpdate) - (ScoreMil*1000000) - (Score100K*100000) - (Score10K*10000) - (ScoreK*1000) ) / 100)
	Score10 = Int( (HighScore(ScoreUpdate) - (ScoreMil*1000000) - (Score100K*100000) - (Score10K*10000) - (ScoreK*1000) - (Score100*100) ) / 10)
	ScoreUnit = (HighScore(ScoreUpdate) - (ScoreMil*1000000) - (Score100K*100000) - (Score10K*10000) - (ScoreK*1000) - (Score100*100) - (Score10*10) )

	Pscore6.image = HSArray(ScoreMil):If HighScore(ScoreUpdate) < 1000000 Then PScore6.image = HSArray(10)
	Pscore5.image = HSArray(Score100K):If HighScore(ScoreUpdate) < 100000 Then PScore5.image = HSArray(10)
	PScore4.image = HSArray(Score10K):If HighScore(ScoreUpdate) < 10000 Then PScore4.image = HSArray(10)
	PScore3.image = HSArray(ScoreK):If HighScore(ScoreUpdate) < 1000 Then PScore3.image = HSArray(10)
	PScore2.image = HSArray(Score100):If HighScore(ScoreUpdate) < 100 Then PScore2.image = HSArray(10)
	PScore1.image = HSArray(Score10):If HighScore(ScoreUpdate) < 10 Then PScore1.image = HSArray(10)
	PScore0.image = HSArray(ScoreUnit):If HighScore(ScoreUpdate) < 1 Then PScore0.image = HSArray(10)
	If HighScore(ScoreUpdate) < 1000 Then
		PComma.image = HSArray(10)
	Else
		PComma.image = HSArray(11)
	End If
	If HighScore(ScoreUpdate) < 1000000 Then
		PComma1.image = HSArray(10)
	Else
		PComma1.image = HSArray(11)
	End If
	If HighScore(ScoreUpdate) >999999 Then Shift = 0:PComma.transx = 0
	If HighScore(ScoreUpdate) < 1000000 Then Shift = 1:PComma.transx = -10
	If HighScore(ScoreUpdate) < 100000 Then Shift = 2:PComma.transx = -20
	If HighScore(ScoreUpdate) < 10000 Then Shift = 3:PComma.transx = -30
	For DHSx = 0 to 6
		EVAL("Pscore" & DHSx).transx = (-10 * Shift)
	Next
	Initial1.image = HSiArray(Initial(ScoreUpdate,1))
	Initial2.image = HSiArray(Initial(ScoreUpdate,2))
	Initial3.image = HSiArray(Initial(ScoreUpdate,3))
	ScoreUpdate = ScoreUpdate + 1
	If ScoreUpdate = 5 then ScoreUpdate = 0
End Sub

'***************Exit Table
Sub Table1_Exit()
	Savehs
	TurnOff
	If B2SOn Then Controller.stop
End Sub

' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX
' PlaySound sound, 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
' *******************************************************************************************************

' Play a sound, depending on the X,Y position of the table element (especially cool for surround speaker setups, otherwise stereo panning only)
' parameters (defaults): loopcount (1), volume (1), randompitch (0), pitch (0), useexisting (0), restart (1))
' Note that this will not work (currently) for walls/slingshots as these do not feature a simple, single X,Y position

Sub PlayXYSound(soundname, tableobj, loopcount, volume, randompitch, pitch, useexisting, restart)
  PlaySound soundname, loopcount, volume, AudioPan(tableobj), randompitch, pitch, useexisting, restart, AudioFade(tableobj)
End Sub

' Set position as table object (Use object or light but NOT wall) and Vol to 1

Sub PlaySoundAt(soundname, tableobj)
  PlaySound soundname, 1, 1, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'Set all as per ball position & speed.

Sub PlaySoundAtBall(soundname)
  PlaySoundAt soundname, ActiveBall
End Sub

'Set position as table object and Vol manually.

Sub PlaySoundAtVol(sound, tableobj, Vol)
  PlaySound sound, 1, Vol, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'Set all as per ball position & speed, but Vol Multiplier may be used eg; PlaySoundAtBallVol "sound",3

Sub PlaySoundAtBallVol(sound, VolMult)
  PlaySound sound, 0, Vol(ActiveBall) * VolMult, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub

'Set position as bumperX and Vol manually.

Sub PlaySoundAtBumperVol(sound, tableobj, Vol)
  PlaySound sound, 1, Vol, AudioPan(tableobj), 0,0,1, 1, AudioFade(tableobj)
End Sub

'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / table1.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / table1.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
  Vol = Csng(BallVel(ball) ^2 / VolDiv)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
  Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
  BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function


'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 3 ' total number of balls
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub RollingTimer_Timer()
    Dim BOT, b
    BOT = GetBalls

	' stop the sound of deleted balls
    For b = UBound(BOT) + 1 to tnob
        rolling(b) = False
        StopSound("fx_ballrolling" & b)
    Next

	' exit the sub if no balls on the table
    If UBound(BOT) = -1 Then Exit Sub

    For b = 0 to UBound(BOT)
      If BallVel(BOT(b) ) > 1 Then
        rolling(b) = True
        if BOT(b).z < 30 Then ' Ball on playfield
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), AudioPan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        Else ' Ball on raised ramp
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*.5, AudioPan(BOT(b) ), 0, Pitch(BOT(b) )+50000, 1, 0, AudioFade(BOT(b) )
        End If
      Else
        If rolling(b) = True Then
          StopSound("fx_ballrolling" & b)
          rolling(b) = False
        End If
      End If
    Next
End Sub

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / (VolDiv/VolCol), AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


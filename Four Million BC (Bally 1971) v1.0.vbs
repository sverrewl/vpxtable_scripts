'****************************************************************
'
'					  Four Million BC
'						by Scottacus
'					      May 2017
'
'	Basic DOF config
'		101 Left Flipper, 102 Right Flipper
'		103 Left sling, 104 Right sling
'		105 Bumper1,  106 Bumper2, 107 Bumper3
'		109 Knocker
'		122 Right Kicker, 123 Left Kicker
'		124 Drain, 125 Ball Release
'	 	212 credit light
'		214 - 217 Bird Ramp
'		129 Knocker and Kicker Strobe
'		160 Ball In Shooter Lane
'		141 Chime1-10s, 142 Chime2-100s, 143 Chime3-1000s
'
'		Code Flow
'									 EndGame
'										^
'		Start Game -> New Game -> Check Continue -> Release Ball -> Drain -> Score Bouns -> Advance Player -> Next Ball
'										^																		  |
'									EndGame = True <---------------------------------------------------------------

'	Ball Control Subroutine developed by: rothbauerw
'		Press "c" during play to activate, the arrow keys control the ball
'**********************************************************************************************************************

' Thalamus 2018-07-23
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' No special SSF tweaks yet.

Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

Const cGameName = "FourMBC"

Dim Balls
Dim Replays
Dim Add1, Add10, Add100, Add1000
Dim MaxPlayers
Dim Players
Dim Player
Dim Credit
Dim Score(6)
Dim HScore(6)
Dim SReels(6)
Dim State
Dim Tilt
Dim TiltSens
Dim Target(9)
Dim BallinPlay
Dim MatchNumber
Dim BallREnabled
Dim RStep, LStep
Dim Rep(4)
Dim EndGame
Dim Bell
Dim i,j, f, ii, Object, Light, x, y, z
Dim AwardCheck
Dim BStop
Dim FreePlay
Dim BallInLane
Dim Ballsize,BallMass
Dim BallHomeCheck
Dim HSArray
Dim HSiArray
Dim Shift
Dim Launched
Dim RelGateHit
Dim HSInitial0, HSInitial1, HSInitial2
Dim EnableInitialEntry
Dim HSi,HSx
Dim RoundHS, RoundHSPlayer
Dim ShowBallShadow
Dim BellRing
Dim ScoreMil, Score100K, Score10K, ScoreK, Score100, Score10, ScoreUnit
HSArray = Array("HS_0","HS_1","HS_2","HS_3","HS_4","HS_5","HS_6","HS_7","HS_8","HS_9","HS_Space","HS_Comma")
HSiArray = Array("HSi_0","HSi_1","HSi_2","HSi_3","HSi_4","HSi_5","HSi_6","HSi_7","HSi_8","HSi_9","HSi_10","HSi_11","HSi_12","HSi_13","HSi_14","HSi_15","HSi_16","HSi_17","HSi_18","HSi_19","HSi_20","HSi_21","HSi_22","HSi_23","HSi_24","HSi_25","HSi_26")
BallSize = 50
BallMass = (Ballsize^3)/125000
Dim BIP : BIP = 0
Dim Options
Dim HSUx
Dim ReplayEB
Dim FirstBallOut
Dim Chime


Sub Table1_init
	LoadEM
	MaxPlayers=4
	Replay(1) = 30000   'SetValue
	Replay(2) = 38000   'SetValue
	Replay(3) = 46000   'SetValue
	Replay(4) = 54000
	Set SReels(1) = ScoreReel1
	Set SReels(2) = ScoreReel2
	Set SReels(3) = ScoreReel3
	Set SReels(4) = ScoreReel4
	Player=1
	LoadHighScore
	'	For HSUx = 0 to 4
'		HighScore(HSUx) = 2002 - (HSUx*2)
'		Score(HSUx) = 0
'	Next
	If HighScore(0)="" Then HighScore(0)=50000
	If HighScore(1)="" Then HighScore(1)=45000
	If HighScore(2)="" Then HighScore(2)=40000
	If HighScore(3)="" Then HighScore(3)=35000
	If HighScore(4)="" Then HighScore(4)=30000
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
	If Chime ="" Then Chime = False

	SaveHS
	UpdatePostIt
	DynamicUpdatePostIt.enabled = 1
	TiltTxt.text="TILT"
	CreditTxt.text = Credit

	For x = 1 to 4
		Score(x) = 0
	Next

	TarPitLight.state = 1
	VLight = 1
	Volcano1000Light.state = 1
	TarPitStartWall1A.isdropped = False
	TarPitStartWall1B.isdropped = True
	LaunchWall.isdropped = True
	LeftKickerWall0.isdropped = True
	LeftKickerWall1.isdropped = False
	LeftKickerLight.state = 1

	For x = 1 to 3
		EVAL("TarPit" & x & "A").isdropped = False
		EVAL("TarPit" & x & "B").isdropped = False
	Next

	If B2SOn Then
		Controller.B2SSetCredits Credit
		Controller.B2SSetMatch 34, MatchNumber
		Controller.B2SSetGameOver 35,1
		Controller.B2SSetTilt 33,1
		Controller.B2SSetBallInPlay 32,0
		Controller.B2SSetPlayerUp 30,0
		If Credit > 0 Then DOF 212, DOFOn
		If FreePlay = 1 Then DOF 212, DOFOn
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

	If ShowBallShadow = 1 Then
		BallShadowUpdate.enabled = True
	Else
		BallShadowUpdate.enabled = False
	End If

	For i = 1 to MaxPlayers
		SReels(i).setvalue(score(i))
	Next
	Tilt=False
	State = False
	GameState
	If FreePlay = 0 Then
		CoinCard.image = "Card" & Balls & "BallsCoin"
	Else
		CoinCard.image = "Card" & Balls & "BallsFree"
	End If

'***********Trough Ball Creation
	TroughSlot1.CreateSizedBallWithMass Ballsize/2, BallMass
	TroughSlot2.CreateSizedBallWithMass Ballsize/2, BallMass
	TroughSlot3.CreateSizedBallWithMass Ballsize/2, BallMass
End Sub

'***********KeyCodes
Sub Table1_KeyDown(ByVal keycode)

	If EnableInitialEntry = True Then EnterIntitals(keycode)

	If keycode = AddCreditKey Then
		PlaySound "coinin"
		AddCredit
    End If

    If keycode = StartGameKey Then
		If EnableInitialEntry = False and operatormenu = 0 Then
			If FreePlay = 1 and Players < 4 and FirstBallOut = False Then StartGame
			If FreePlay = 0 and Credit > 0 and Players < 4 and FirstBallOut = False Then
				Credit = Credit - 1
				CreditTxt.text = Credit
				If B2SOn Then
					Controller.B2SSetCredits Credit
					If FreePlay = 0 and Credit < 1 Then DOF 212, DOFOff
				End If
				StartGame
			End If
		End If
	End If

	If keycode = PlungerKey Then
		Plunger.PullBack
		PlayFieldSound "plungerpull", 0, Plunger
	End If

  If Tilt = False and State = True Then
	If keycode = LeftFlipperKey and contball = 0 Then
		LeftFlipper.RotateToEnd
		LeftFlipper1.RotateToEnd
		PlayFieldSound "flipperup", 0, LeftFlipper
		If B2SOn Then DOF 101,DOFOn
		PlayFieldSound "Buzz", -1, LeftFlipper
	End If

	If keycode = RightFlipperKey and contball = 0 Then
		RightFlipper.RotateToEnd
		RightFlipper1.RotateToEnd
		PlayFieldSound "flipperup", 0, RightFlipper
		If B2SOn Then DOF 102,DOFOn
		PlayFieldSound "Buzz", -1, RightFlipper
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
    end if

    If keycode = LeftFlipperKey and State = False and OperatorMenu = 1 Then
		Options = Options + 1
        If Options = 5 then Options = 0
'		TextBox1.text = options
		OptionMenu.visible = true
        PlayFieldSound "target", 0, OptionMenu
        Select Case (Options)
            Case 0:
                OptionMenu.image = "FreeCoin" & FreePlay
            Case 1:
                OptionMenu.image = Balls & "Balls"
            Case 2:
                OptionMenu.image = "BallShadow" & ShowBallShadow
			Case 3:
				OptionMenu.image = "Chime" & Chime

			Case 4:
				OptionMenu.image = "SaveExit"
        End Select
    End If

    If keycode = RightFlipperKey and State = False and OperatorMenu = 1 Then
      PlayFieldSound "metalhit2", 0, OptionMenu
      Select Case (Options)
		Case 0:
            If FreePlay = 0 Then
                FreePlay = 1
              Else
                FreePlay = 0
            End If
            OptionMenu.image= "FreeCoin" & FreePlay
				If FreePlay = 0 Then
					CoinCard.image = "Card" & Balls & "BallsCoin"
				Else
			CoinCard.image = "Card" & Balls & "BallsFree"
			End If
			If FreePlay = 0 Then
				If Credit > 0 and B2SOn Then DOF 212, DOFOn
				If Credit < 1 and B2SOn Then DOF 212, DOFOff
			Else
				If B2SOn Then DOF 212, DOFOn
			End If
        Case 1:
            If Balls = 3 Then
                Balls = 5
              Else
                Balls = 3
            End If
            OptionMenu.image = Balls & "Balls"
			If FreePlay = 0 Then
				CoinCard.image = "Card" & Balls & "BallsCoin"
			Else
				CoinCard.image = "Card" & Balls & "BallsFree"
			End If
        Case 2:
            If ShowBallShadow = 0 Then
                ShowBallShadow = 1
              Else
                ShowBallShadow = 0
            End If
			OptionMenu.image = "BallShadow" & ShowBallShadow
			If ShowBallShadow = 1 Then
				BallShadowUpdate.enabled = True
			Else
				BallShadowUpdate.enabled = False
			End If
        Case 3:
            If Chime = 0 Then
                Chime= 1
				If B2SOn Then DOF 142,DOFPulse
              Else
                Chime = 0
				PlayFieldsound "Bell10", 0, SoundPoint13
            End If
			OptionMenu.image = "Chime" & Chime
        Case 4:
            OperatorMenu = 0
            SaveHS
			DynamicUpdatePostIt.enabled = 1
			OptionMenu.image = "FreeCoin" & FreePlay
            OptionMenu.visible = 0
			OptionsMenu.visible = 0
		End Select
    End If

	If Keycode = MechanicalTilt Then
		Tilt = True
		Tilttxt.text = "TILT"
		If B2SOn Then Controller.B2SSetTilt 1
		Playsound"Tilt"
		TurnOff
		UnZipFlippers
	End If

    If keycode = 46 Then' C Key
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
	If keycode = 30 Then ZipFlippers


	If keycode = 31 Then UnZipFlippers

	If keycode = 33 Then VolcanoRubber_Hit


'************************End Of Test Keys****************************
End Sub

Sub Table1_KeyUp(ByVal keycode)

	If keycode = PlungerKey Then
		Plunger.Fire
		PlayFieldSound "Plunger", 0, Plunger
	End If

    If keycode = LeftFlipperKey Then
        OperatorMenuTimer.Enabled = False
    End If

   If Tilt = False and State = True Then
		If keycode = LeftFlipperKey and contball = 0 Then
			LeftFlipper.RotateToStart
			LeftFlipper1.RotateToStart
			PlayFieldSound "flipperdown", 0, LeftFlipper
			If B2SOn Then DOF 101,DOFOff
			StopSound "Buzz"
		End If

		If keycode = RightFlipperKey and contball = 0 Then
			RightFlipper.RotateToStart
			RightFlipper1.RotateToStart
			PlayFieldSound "flipperdown", 0, RightFlipper
			If B2SOn Then DOF 102,DOFOff
			StopSound "Buzz"
		End If
   End If


    If keycode = 203 then Cleft = 0' Left Arrow

    If keycode = 200 then Cup = 0' Up Arrow

    If keycode = 208 then Cdown = 0' Down Arrow

    If keycode = 205 then Cright = 0' Right Arrow
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
			Controller.B2SSetBallinPlay 32, BallinPlay
			Controller.B2SSetPlayerup 30, 1
			Controller.B2SSetCanPlay 31, 1
			Controller.B2SSetGameOver 0
		End If
		DynamicUpdatePostIt.enabled = 0
		UpdatePostIt
		Tilt = False
		State = True
		GameState
		PlayFieldSound "initialize", 0, Bumper1
		Players = 1
		CanPlayTxt.text = Players
		For x = 1 to 4
			Score(x) = 0
			If B2SOn Then controller.B2SSetScorePlayer x, Score(x)
			SReels(x).setvalue(0)
		Next
		NewGame.enabled = True
	Else If State = True and Players < MaxPlayers and BallinPlay = 1 Then
		Players = Players + 1
		CanPlayTxt.text = Players
		CreditTxt.text = Credit
		If B2SOn Then
			Controller.B2SSetCredits Credit
			Controller.B2SSetCanplay 31, Players
		End If
		PlayFieldSound "cluper", 0, Bumper1
		End If
	End If
End Sub

'*********New Game
Sub NewGame_timer
	Player = 1
	For i = 1 to MaxPlayers
	    Score(i) = 0
		Rep(i) = 0
		RepAwarded(1) = 0
	Next
	RoundHS = 0
	RoundHSPlayer = 1
	If B2SOn Then
	  For i = 1 to MaxPlayers
		Controller.B2SSetScorePlayer i, score(i)
	  Next
	End If
    EndGame = 0
	UnZipFlippers
	If B2SOn Then controller.B2SSetShootAgain 36,0
	GameState
	CheckContinue.enabled = 1
	NewGame.enabled = False
	For f = 1 to 3
		EVAL("Bumper"&f).hashitevent = 1
	Next
	Player1.intensityscale = 1.5
	NewGame.enabled = False
End Sub

'**********Check if Game Should Continue
Sub CheckContinue_Timer
	If EndGame = 1 Then
		TurnOff
		StopSound "Buzz1"
		Match
		State = False
		BIPText.text = " "
		GameState
		DynamicUpdatePostIt.enabled = 1
		CanPlayTxt.text = 0
 		SortScores
		CheckHighScores
		FirstBallOut = False
		HsTxt.text = Score(0)
		Players = 0
		For x = 1 to 4
			Rep(x) = 0
			RepAwarded(x) = 0
		Next
		SaveHS
		For x = 1 to 3
			EVAL("BumperLight" & x).state = 0
			bumpers_prim.image = "bumpercaps"
		Next
		If B2SOn Then
			Controller.B2SSetGameOver 35,1
			Controller.B2SSetballinplay 32, 0
			Controller.B2SSetPlayerUp 30, 0
			Controller.B2SSetcanplay 31, 0
			If Credit > 0 Then DOF 212, DOFOn
			If FreePlay = 1 Then DOF 212, DOFOn
		End If
		For each Light in GIlights:light.state = 0:Next
		CheckContinue.enabled = 0
		VolcanoLightTimer.enabled = 0
		StopSound "Buzz1"
	Else
		If BallInLane = False and BIP = 0 Then
			BIPText.text = BallinPlay
			ReleaseBall
			For x = 1 to 3
				EVAL("BumperLight" & x).state = 0
				bumpers_prim.image = "bumpercaps"
			Next
		End If
		CheckContinue.enabled = 0
  End If
End Sub

'***************Drain and Release Ball
Sub Drain_Hit()
	BIP = BIP - 1
	UpdateText
	PlayFieldSound "fx_drain", 0, Drain
	UpdateTrough
	RepAwarded(Player) = 0
	BirdLight.state = 0
	BirdPlasticLight.state = 0
	BirdPlasticLight1.state = 0
	Plastic3k_prim.image="4MBC3kplasticOFF"
	rampmetalcover_prim.image="rampcover"
End Sub

Sub ReleaseBall
	If BIP = 0 Then
		PlayFieldSound "kickerkick", 0, TroughSlot1
		If B2SOn Then DOF 110,DOFPulse
		TroughSlot1.kick 60, 5
		PlayFieldSound "metalhit_thin", 0, TroughSlot1
		BIP = BIP + 1
		UpdateText
		UpdateTrough
		Launched = 0
		BIPText.text = BallinPlay
		If B2SOn Then Controller.B2SSetBallinPlay 32, BallinPlay
		If BallREnabled = 0 then ReleaseBall
	End If
End Sub



'**********Check if Scoring Bonus is True
Sub ScoreBonus
	AdvancePlayers
End Sub

'**********Advance Players
Sub AdvancePlayers
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
	If B2SOn Then Controller.B2SSetPlayerup 30, Player
	NextBall
End Sub

'**********Next Ball
Sub NextBall
    If Tilt = True Then
	  For f = 1 to 3
		EVAL("Bumper"&f).hashitevent = 1
	  Next
      Tilt = False
      TiltTxt.text = " "
		If B2SOn Then
			Controller.B2SSetTilt 33,0
			Controller.B2SSetData 1, 1
		End If
    End If
'	ShootAgainLight.state = 0
	If Player = 1 then BallinPlay = BallinPlay + 1
'	If player = 4 then BallinPlay = 5 'used for match testing

	If BallinPlay > Balls then
		PlayFieldSound "GameOver", 0, Bumper1
		EndGame = 1
		CheckContinue.enabled = 1
	Else
		If State = True Then
			CheckContinue.enabled = 1
		End If
	End If
End Sub

'************Coin Handelers
Sub AddCredit
	Credit = Credit + 1
	If B2SOn Then DOF 212, DOFOn
	If Credit > 16 then Credit = 16
	CreditTxt.text = Credit
	If B2SOn Then
		Controller.B2SSetCredits Credit
		If Credit > 0 Then DOF 212, DOFOn
	End If
End Sub

'************Game State Check
Sub GameState
	If State = False Then
		For each light in GIlights:light.state = 0: Next
		Flasher2.visible = 1
		Flasher1.visible = 0
		rampmetalwalls_prim.blenddisablelighting = 0.05
		rampmetalcover_prim.blenddisablelighting = 0.01
		ramp_prim.blenddisablelighting = 0.05
		outers_prim.image = "outersGIOFF"
		outers_prim.blenddisablelighting = 0.1
		metalupper_prim.blenddisablelighting = 0.05
		metalguideplunger_prim.blenddisablelighting = 0.01
		metalguidelowerleft_prim.blenddisablelighting = 0.05
		metalplates_prim.blenddisablelighting = 0.05
		metalbrackets_prim.blenddisablelighting = 0.05
		yellowcap1.blenddisablelighting = 0.05
		yellowcap2.blenddisablelighting = 0.05
		plastics_prim.blenddisablelighting = 0.1
		plastic3k_prim.blenddisablelighting = 0.1
		bumpers_prim.blenddisablelighting = 0.1
		rflip.blenddisablelighting = 0.05
		rflip1.blenddisablelighting = 0.05
		lflip.blenddisablelighting = 0.05
		lflip1.blenddisablelighting = 0.05
		rflip.image = "rflipGIOFF"
		rflip1.image = "rflipGIOFF"
		lflip.image = "lflipGIOFF"
		lflip1.image = "lflipGIOFF"
		plastics_prim.image="plasticsGIOFF"
		backmetalrail_prim.image="backrailGIOFF"
		GamOv.text="Game Over"
		If B2SOn then Controller.B2SSetGameOver 35,1
		For i = 1 to 4
			EVAL ("Player" & i).intensityscale = 1
		Next
		If B2SOn then Controller.B2SSetData 80,0
		StopSound "Buzz1"
	Else
		For each Light in GIlights:Light.state = 1: Next
		Flasher2.visible = 0
		Flasher1.visible = 1
		rampmetalwalls_prim.blenddisablelighting = 0.2
		rampmetalcover_prim.blenddisablelighting = 0.08
		ramp_prim.blenddisablelighting = 0.2
		outers_prim.image = "outersGION"
		outers_prim.blenddisablelighting = 0.3
		metalupper_prim.blenddisablelighting = 0.2
		metalguideplunger_prim.blenddisablelighting = 0.08
		metalguidelowerleft_prim.blenddisablelighting = 0.08
		metalplates_prim.blenddisablelighting = 0.15
		metalbrackets_prim.blenddisablelighting = 0.2
		yellowcap1.blenddisablelighting = 0.25
		yellowcap2.blenddisablelighting = 0.25
		plastics_prim.blenddisablelighting = 0.6
		plastic3k_prim.blenddisablelighting = 0.3
		bumpers_prim.blenddisablelighting = 0.4
		rflip.blenddisablelighting = 0.3
		rflip1.blenddisablelighting = 0.2
		lflip.blenddisablelighting = 0.3
		lflip1.blenddisablelighting = 0.2
		rflip.image = "rflipGION"
		rflip1.image = "rflipGION"
		lflip.image = "lflipGION"
		lflip1.image = "lflipGION"
		plastics_prim.image="plasticsGION"
		backmetalrail_prim.image="backrailGION"
		GamOv.text=""
		MatchTxt.text= ""
		TiltTxt.text=" "
		If B2SOn Then
			Controller.B2SSetTilt 33,0
			Controller.B2SSetMatch 34,0
			Controller.B2SSetGameOver 35,0
			Controller.B2SSetData 80,1
		End If
		If Volcano = 1 Then
			PlayFieldSound "Buzz1", -1, VolcanoSaucer
			VolcanoLightTimer.enabled = 1

		End If
	End If
End Sub

'*************Trough based on nFozzy's
Dim Trough
Sub UpdateTrough()
	 Trough= 0
	UpdateTroughTimer.Interval = 300
	UpdateTroughTimer.Enabled = 1
End Sub

Sub UpdateTroughTimer_Timer()
	If TroughSlot1.BallCntOver = 0 Then TroughSlot2.kick 60, 8
	If TroughSlot2.BallCntOver = 0 Then TroughSlot3.kick 60, 8
	If TroughSlot3.BallCntOver = 0 Then Drain.kick 60, 30
	If Trough > 4 Then
		Me.Enabled = 0
		If BIP = 0 Then ScoreBonus
	End If
	Trough = Trough + 1
End Sub

'*************Ball in Launch Lane
Sub BallHome_hit
	BallREnabled = 1
	If B2SOn Then DOF 121, DOFOn
	RelGateHit = 0
	Set ControlBall = ActiveBall
    contballinplay = True
End Sub

Sub BallHome_unhit
	If B2SOn Then DOF 121, DOFOff
End Sub

Sub EndControl_Hit()                '******* for ball control script
    contballinplay = false
End Sub
'************Check if Ball Out of Launch Lane
Sub BallsInPlay_hit
	If BallREnabled=1 Then
'		ShootAgainLight.state = 0
		If B2SOn Then controller.B2SSetShootAgain 36,0
		BallREnabled = 0
		BallInLane = False
	End If
	GateTimer.enabled = 1
End Sub

'***************Match
Sub Match
   MatchNumber = int(rnd(1)*10)
		If B2SOn Then
			Controller.B2SSetMatch 34,MatchNumber + 1
		End If
	MatchTxt.text = MatchNumber

	For i = 1 to Players
		MatchScoreTxt.text =  (Score(i) mod 100)
		If (MatchNumber * 10) = (Score(i) mod 100) Then
			AddCredit
			Playsound SoundFXDOF("Knocker",109,DOFPulse,DOFKnocker)
	    End If
    Next
End Sub

'************** Bumpers

Sub Bumpers_hit(Index)
	If Tilt = False Then
		Select Case (Index)
			Case 0: If BumperLight1.state = 0 Then
						AddScore 10
					Else
						AddScore 100
					End If
					PlayFieldSound "fx_bumper1", 0, Bumper1
					If B2SOn Then DOF 105,DOFPulse
			Case 1: If BumperLight2.state = 0 Then
						AddScore 10
					Else
						AddScore 100
					End If
					PlayFieldSound "fx_bumper2", 0, Bumper2
					If B2SOn Then DOF 106,DOFPulse
			Case 2: If BumperLight3.state = 0 Then
						AddScore 10
					Else
						AddScore 100
					End If
					PlayFieldSound "fx_bumper3", 0, Bumper3
					If B2SOn Then DOF 107,DOFPulse
		End Select
	End If
End Sub

'************** Flippers
Sub FlipperTimer_timer()
	LFlip.Rotz = LeftFlipper.CurrentAngle
	RFlip.Rotz = RightFlipper.CurrentAngle
	LFlip1.Rotz = LeftFlipper1.CurrentAngle
	RFlip1.Rotz = RightFlipper1.CurrentAngle
End Sub

'***********Rotate Flipper Shadow
Sub FlipperShadowUpdate_Timer
	FlipperLSh.RotZ = LeftFlipper.CurrentAngle
	FlipperRSh.RotZ = RightFlipper.CurrentAngle
	FlipperLSh1.RotZ = LeftFlipper1.CurrentAngle
	FlipperRSh1.RotZ = RightFlipper1.CurrentAngle
End Sub

'************** Slings
Sub LeftSlingShot_Slingshot
	PlayfieldSound "Slingshot", 0, SoundPoint12

	If B2SOn Then DOF 103,DOFPulse
    LSling.Visible = 0
    LSling1.Visible = 1
    Sling1.Rotx = 22
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
	AddScore 10
	BirdLight.state = 0
	BirdPlasticLight.state = 0
	BirdPlasticLight1.state = 0
	Plastic3k_prim.image="4MBC3kplasticOFF"
	rampmetalcover_prim.image="rampcover"
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3: LSLing1.Visible = 0:LSLing2.Visible = 1:sling1.Rotx = 8
        Case 4: LSLing2.Visible = 0:LSLing.Visible = 1:sling1.Rotx = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
	PlayfieldSound "Slingshot", 0, SoundPoint13
	If B2SOn Then DOF 104,DOFPulse
    RSling.Visible = 0
    RSling1.Visible = 1
    Sling2.RotX = 22
    RStep = 0
    RightSlingShot.TimerEnabled = 1
	AddScore 10
	BirdLight.state = 0
	BirdPlasticLight.state = 0
	BirdPlasticLight1.state = 0
	Plastic3k_prim.image="4MBC3kplasticOFF"
	rampmetalcover_prim.image="rampcover"
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3: RSling1.Visible = 0:RSling2.Visible = 1:Sling2.Rotx = 8
        Case 4: RSling2.Visible = 0:RSling.Visible = 1:Sling2.Rotx = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub


Sub TenPointRubbers_Hit(Idx)
	If Tilt = False Then AddScore 10
	If Idx = 1 or Idx = 5 or Idx = 6 Then
		RubberIndex = Idx + 1
		RubberAnimation.enabled = 1
	End If
	BirdLight.state = 0
	BirdPlasticLight.state = 0
	BirdPlasticLight1.state = 0
	Plastic3k_prim.image="4MBC3kplasticOFF"
	rampmetalcover_prim.image="rampcover"
End Sub

Dim RubberStep, RubberIndex
Sub RubberAnimation_Timer
	RubberStep = RubberStep + 1
	Select Case RubberStep
		Case 1: EVAL("Rubber" & RubberIndex).visible = 0
				EVAL("Rubber" & RubberIndex & "A").visible = 1
		Case 2: EVAL("Rubber" & RubberIndex).visible = 1
				EVAL("Rubber" & RubberIndex & "A").visible = 0
				RubberStep = 0
				RubberAnimation.enabled = 0
	End Select
End Sub

'*************** Triggers
Sub LeftOutLaneTrigger_hit
	If Tilt = False Then
		AddScore 1000
		If B2SOn Then DOF 210, DOFPulse
	End If
End sub

Sub RightOutLaneTrigger_hit
	If Tilt = False Then
		AddScore 1000
		If B2SOn Then DOF 211, DOFPulse
	End If
End sub

Sub BallLaunchedTrigger_hit
	FirstBallOut = True
	If B2SOn Then DOF 212, DOFOff
End Sub

Sub BirdTriggers_Hit(Index)
	If Tilt = False Then

		If BirdLight.State = 1 Then
			If (Index) < 3 Then AddScore 1000
		Else
			AddScore 100
		End If

		Select Case (Index)
			Case 0: BumperLight1.State = 1
					BumperLight1a.State = 1
					If B2SOn Then DOF 214, DOFPulse
					Launchwall.isdropped = False
					LaunchGateLight.state = 0
					If BumperLight2.state = 0 Then Bumpers_prim.image = "bumpers1k"
			Case 1: BumperLight2.State = 1
					BumperLight2a.State = 1
					If B2SOn Then DOF 215, DOFPulse
					If Bumperlight3.state = 0 Then Bumpers_prim.image = "bumpers2k"
			Case 2: BumperLight3.State = 1
					BumperLight3a.State = 1
					If B2SOn Then DOF 216, DOFPulse
					Bumpers_prim.image = "bumperslit"
			Case 3: BirdLight.State = 1
					BirdPlasticLight.State = 1
					BirdPlasticLight1.State = 1
					rampmetalcover_prim.image="rampcover3kON"
					plastic3k_prim.image="4MBC3kplastic"
					If B2SOn Then DOF 217, DOFPulse

		End Select
	End If
End Sub

Sub BirdEndTrigger_Hit
	If Tilt = False Then AddScore 10
	BirdLight.state = 0
	BirdPlasticLight.state = 0
	BirdPlasticLight1.state = 0
	Plastic3k_prim.image="4MBC3kplasticOFF"
	rampmetalcover_prim.image="rampcover"
End Sub

Dim LaneOpen, ButtonNumber
Sub GateTriggers_Hit (Index)
	ButtonNumber = (Index)
	RollOverAnimation.enabled = 1
	Select Case (Index)
		Case 0: LeftKickerLight.state = 0
				LaunchGateLight.state = 1
				Launchwall.isdropped = True
				LeftKickerWall0.isdropped = False
				LeftKickerWall1.isdropped = True
				LaneOpen = True
		Case 1:	LeftKickerLight.state = 1
				LaunchGateLight.state = 0
				Launchwall.isdropped = False
				LeftKickerWall0.isdropped = True
				LeftKickerWall1.isdropped = False
				LaneOpen = False
		Case 2: LeftKickerLight.state = 0
				LaunchGateLight.state = 1
				Launchwall.isdropped = True
				LeftKickerWall0.isdropped = False
				LeftKickerWall1.isdropped = True
				LaneOpen = True
	End Select
End Sub

Sub LaunchExit_Hit
	If LaneOpen = False Then LaunchWall.isdropped = False
End Sub

Sub LaunchExitInside_Hit
	LaunchWall.isdropped = True
End Sub


Sub TarPitTrigger_Hit
	If TarPitStage = 0 Then
		TarPitLight.state = 0
		TarPitAdvanceLight.state = 1
		TarPitStage = 1
		TarPitStartWall1A.isdropped = True
		TarPitStartWall1B.isdropped = False
		metal_uppergate.rotz = 45
		BIP = BIP - 1
		BIPtxt.text = BIP
		ReleaseBall
	End If
End Sub

Sub TarPitEnd_Hit
	TarPitStage = 0
	BIP = BIP + 1
	TarPitLight.state = 1
	TarPitAdvanceLight.state =  0
	TarPitStartWall1A.isdropped = False
	TarPitStartWall1B.isdropped = True
	metal_uppergate.rotz = 0
	For x = 1 to 3
		EVAL("TarPit" & x & "A").isdropped = False
		EVAL("TarPit" & x & "B").isdropped = False
	Next
	BIPtxt.text = BIP
	ZipFlippers
End Sub

Dim Mush
Sub MushroomAnimation_Timer
	Mush = Mush + 1
	Select Case Mush
		Case 1: EVAL("YellowCap" & Cap).transy = 7
		Case 2: EVAL("YellowCap" & Cap).transy = 12
		Case 3: EVAL("YellowCap" & Cap).transy = 7
		Case 4: EVAL("YellowCap" & Cap).transy = 0
				Mush = 0
				MushroomAnimation.enabled = 0
	End Select
End Sub

Dim Button
Sub RollOverAnimation_Timer
	Button = Button +1
	Select Case Button
		Case 1: EVAL ("RollOverButton" & ButtonNumber).transz = -2
		Case 2: EVAL ("RollOverButton" & ButtonNumber).transz = 0
		Case 3: EVAL ("RollOverButton" & ButtonNumber).transz = 1
		Case 4: EVAL ("RollOverButton" & ButtonNumber).transz = 1
				Button = 0
				RollOverAnimation.enabled = 0
	End Select
End Sub

'***************Tar Pit*******************
Dim TarPitStage, Cap
Sub TarPitRubber_Hit
	Cap = 2
	MushroomAnimation.enabled = 1
	If Tilt = False Then
		If TarPitAdvanceLight.state = 1 Then
			AddScore 1000
			If TarPitStage > 3 Then TarPitStage = 3
			EVAL("TarPit" & TarPitStage & "A").isdropped = True
			EVAL("TarPit" & TarPitStage).enabled = 1
			TarPitStage = TarPitStage + 1
		Else
			AddScore 100
		End If
		UnZipFlippers
		If B2SOn Then DOF 301, DOFPulse
		PlayFieldSound "MRCollision", 0, Yellowstem2
	End If
End Sub

Dim TarPitPause(3)
Sub TarPit1_Timer
	TarPitPause(1) = TarPitPause(1) + 1
	Select Case (TarPitPause(1))
		Case 1: Tarpit1A.isdropped = True: TarPitPrimStep1
		Case 2: TarPit1B.isdropped = True: TarPitPrimStep2
		Case 3: TarPit1.enabled = False
				TarPitPause(1) = 0
	End Select

End Sub

Sub TarPit2_Timer
	TarPitPause(2) = TarPitPause(2) + 1

	Select Case (TarPitPause(2))
		Case 1:	Tarpit2A.isdropped = True: TarPitPrimStep1
		Case 2: TarPit2B.isdropped = True: TarPitPrimStep2
		Case 3: TarPit2.enabled = False
				TarPitPause(2) = 0
	End Select
End Sub

Sub TarPit3_Timer
	TarPitPause(3) = TarPitPause(3) + 1
	Select Case (TarPitPause(3))
		Case 1: Tarpit3A.isdropped = True: TarPitPrimStep1
		Case 2: TarPit3B.isdropped = True: TarPitPrimStep2
		Case 3: TarPit3.enabled = False
				TarPitPause(3) = 0
	End Select
End Sub


Sub TarPitPrimStep1
	For x = 1 to 3
		EVAL("BallStop_Prim" & x).ObjRotX = 60
	Next
End Sub

Sub TarPitPrimStep2
	For x = 1 to 3
		EVAL("BallStop_Prim" & x).ObjRotX = -5
	Next
End Sub


'***************Volcano********************
Dim VLight
Sub VolcanoLightTimer_Timer
	Vlight = VLight + 1
	EVAL("Volcano" & (VLight * 1000) & "Light").state = 1
	If VLight > 1 Then
		EVAL("Volcano" & ((VLight - 1) * 1000) & "Light").state = 0
	Else
		Volcano5000Light.state = 0
	End If
	If VLight = 5 Then VLight = 0
End Sub

dim lkickstep

Sub VolcanoRubber_Hit
	Cap = 1
	MushroomAnimation.enabled = 1
	If Tilt = False Then
		If VolcanoLight.state = 1 Then
			BIP = BIP + 1
			VolcanoSaucer.TimerEnabled = 1
			lkickstep = 0
			PlayFieldSound "MRCollision", 0, YellowStem1
			BIPtxt.text = BIP
			AddScore (VLight * 1000)
			VolcanoLightTimer.Enabled = 0
			Volcano = 0
			VolcanoLight.state = 0
			StopSound "Buzz1"
		Else
			AddScore 100
		End If
		If B2SOn Then
			DOF 118, DOFOff
			DOF 300, DOFPulse
		End If
		ZipFlippers
	End If
End Sub

Sub VolcanoSaucer_timer
	StopSound "Buzz1"
    Select Case LKickstep
        Case 1:
			kickarmtop_prim.ObjRotX=-12
			PlayFieldSound "KickerKick", 0, VolcanoSaucer
			VolcanoSaucer.Kick 150,8
        Case 2:kickarmtop_prim.ObjRotX = -45
        Case 3:kickarmtop_prim.ObjRotX = -45
        Case 4:kickarmtop_prim.ObjRotX = -45
        Case 5:kickarmtop_prim.ObjRotX = -45
        Case 6:kickarmtop_prim.ObjRotX = -45
        Case 7:kickarmtop_prim.ObjRotX = -45
        Case 8:kickarmtop_prim.ObjRotX = -45
        Case 9:kickarmtop_prim.ObjRotX = -45
        Case 10:kickarmtop_prim.ObjRotX = -45
        Case 11:kickarmtop_prim.ObjRotX = 24
        Case 12:kickarmtop_prim.ObjRotX = 12
        Case 13:kickarmtop_prim.ObjRotX = 0:VolcanoSaucer.TimerEnabled = 0
    End Select
    LKickstep = LKickstep + 1
End Sub

Dim Volcano
Sub VolcanoSaucer_Hit
	VolcanoLight.state = 1
	PlayFieldSound "Buzz1", -1, VolcanoSaucer
	Volcano = 1
	VolcanoLightTimer.enabled = 1
	BIP = BIP - 1
	ReleaseBall
	If B2SOn Then DOF 118, DOFOn
End Sub


'***************Left OutLane Kicker
Sub LeftOutLaneKicker_Hit
	LeftOutLaneKickerTimer.enabled = 1
End Sub

Sub LeftOutLaneKickerTimer_Timer
	LeftOutLaneKicker.kick 0,30
	If B2SOn Then DOF 213, DOFPulse
	PlayFieldSound "KickerKick", 0, LeftOutLaneKicker
	LeftOutLaneKickerTimer.enabled = 0
End Sub

'***************Target Hit
Sub Targets_Hit(Index)
	If Tilt = False Then
		AddScore 100
		If Index = 0 Then
			ZipFlippers
		Else
			UnZipFlippers
		End If
		If B2SOn Then DOF 210, DOFPulse
	End If
End Sub

'************Zip Flippers
Sub ZipFlippers
	LeftFlipper.enabled = False
	RightFlipper.enabled = False
	LeftFlipper1.enabled = True
	RightFlipper1.enabled = True
	FlipperLSh.visible = False
	FlipperRSh.visible = False
	FlipperLSh1.visible = True
	FlipperRSh1.visible = True
	LFlip.visible = False
	RFlip.visible = False
	LFlip1.visible = True
	RFlip1.visible = True
	ZWall1.isdropped = False
	ZWall2.isdropped = False
	PlayFieldSound "fclose", 0, LeftFlipper1
End Sub

Sub UnZipFlippers
	LeftFlipper.enabled = True
	RightFlipper.enabled = True
	LeftFlipper1.enabled = False
	RightFlipper1.enabled =	False
	FlipperLSh1.visible = False
	FlipperRSh1.visible = False
	FlipperLSh.visible = True
	FlipperRSh.visible = True
	LFlip.visible = True
	RFlip.visible = True
	LFlip1.visible = False
	RFlip1.visible = False
	ZWall1.isdropped = True
	ZWall2.isdropped = True
	StopSound "Buzz"
	PlayFieldSound "fclose", 0, LeftFlipper
End Sub

'**************Special
Sub Special
		AddCredit
		Playsound SoundFXDOF("Knocker",109,DOFPulse,DOFKnocker)
End Sub

'**************Update Desktop Text
Sub UpdateText
	BIPtxt.text = BIP
End Sub

'***************Scoring Routine
Sub AddScore(points)
  If Tilt = False Then
		If Points <100 Then BellRing = (Points / 10): BellTimer10.enabled = 1
		If Points > 99 and Points < 1000 Then BellRing = (Points / 100): BellTimer100.enabled = 1
		If Points > 999 Then BellRing = (Points / 1000): BellTimer1000.enabled = 1
	End If
End Sub

Dim ReplayX, RepAwarded(5), Replay(7), ReplayText(3)

Sub TotalUp(Points)
	Score(Player) = Score(Player) + Points
	SReels(Player).addvalue(Points)
	If B2SOn Then Controller.B2SSetScorePlayer Player, Score(Player)
	For ReplayX = Rep(Player) +1 to 4
		If Score(Player) => Replay(ReplayX) Then
			If RepAwarded(Player) < 1 Then
				If ReplayEB = 0 Then
					AddCredit
				Else
	'				ShootAgainLight.state = 1
					If B2SOn Then controller.B2SSetShootAgain 36,1
				End If
				Rep(Player) = Rep(Player) + 1
			'	RepAwarded(Player) = 1
				Playsound SoundFXDOF("Knocker",109,DOFPulse,DOFKnocker)
			Else
				Rep(Player) = Rep(Player) + 1
			End If
		End If
	Next
End Sub

'*************** Bell Timers
Sub BellTimer10_Timer
	PlaySound "Reels"
	If Chime = False Then
		PlayFieldSound "Bell10", 0, SoundPoint13
	Else
		If B2SOn Then DOF 141,DOFPulse
	End If
	BellRing = BellRing - 1
	TotalUp 10
	If BellRing < 1 Then
		BellTimer10.enabled = 0
	End If
End Sub

Sub BellTimer100_Timer
	PlaySound "Reels"
	If Chime = False Then
		PlayFieldSound  "Bell100", 0, SoundPoint13
	Else
		If B2SOn Then DOF 142,DOFPulse
	End If
	BellRing = BellRing - 1
	TotalUp 100
	If BellRing < 1 Then
		BellTimer100.enabled = 0
	End If
End Sub

Sub BellTimer1000_Timer
	PlaySound "Reels"
	If Chime = False Then
		PlayFieldSound  "Bell1000", 0, SoundPoint13
	Else
		If B2SOn Then DOF 143,DOFPulse
	End If
	BellRing = BellRing - 1
	TotalUp 1000
	If BellRing < 1 Then
		BellTimer1000.enabled = 0
	End If
End Sub


'***************Tilt
Sub CheckTilt
	If Tilttimer.Enabled = True Then
		TiltSens = TiltSens + 1
		If TiltSens = 3 Then
		Tilt = True
		tilttxt.text="TILT"
       	If B2SOn Then Controller.B2SSetTilt 33,1
       	If B2SOn Then Controller.B2SSetdata 1, 0
		PlaySound "tilt"
		TurnOff
		UnZipFlippers
	 End If
	Else
	 TiltSens = 0
	 Tilttimer.Enabled = True
	End If
End Sub

Sub Tilttimer_Timer()
	Tilttimer.Enabled = False
End Sub

'************Shut Down and De-Energize on Tilt
Sub TurnOff
	For i= 1 to 3
		EVAL("Bumper"&i).hashitevent = 0
	Next
  	LeftFlipper.RotateToStart
	StopSound "Buzz"
	If B2SOn Then DOF 101, DOFOff
	RightFlipper.RotateToStart
	If B2SOn Then DOF 102, DOFOff
	UnZipFlippers
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

Sub LightsRandom_Timer()
	Select Case Int(Rnd*2)+1
		Case 1 : DOF 157, 1
		Case 2 : DOF 157, 0
	End Select
	Select Case Int(Rnd*2)+1
		Case 1 : DOF 158, 1
		Case 2 : DOF 158, 0
	End Select
	Select Case Int(Rnd*2)+1
		Case 1 : DOF 159, 1
		Case 2 : DOF 159, 0
	End Select
	Select Case Int(Rnd*2)+1
		Case 1 : DOF 160, 1
		Case 2 : DOF 160, 0
	End Select
End Sub



Dim BotPos, Pos


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

'*************Hit Sound Routines

Sub PlayFieldSound (SoundName, Looper, TableObject)
	PlaySound SoundName, Looper, 1, AudioPan(TableObject), 0, 0, 0, 0, AudioFade(TableObject)
End Sub

Sub a_Posts_Hit(idx)
 	dim finalspeed
  	finalspeed= 2 * (SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely))
 	If finalspeed > 16 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		TextBox1.text = 16
	End if
	If finalspeed >= 6 AND finalspeed <= 16 then
 		RandomSoundRubber()
 	End If
End Sub

Sub a_Targets_Hit (idx)
	PlayFieldSound "target", 0,  EVAL("Target" & idx + 1)
End Sub

Sub a_Metals_Thin_Hit (idx)
	PlaySound "metalhit_thin", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub a_Metals_Medium_Hit (idx)
	PlaySound "metalhit_medium", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub a_Metals2_Hit (idx)
	PlaySound "metalhit2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub a_Gates_Hit (idx)
	PlayFieldSound "gate4", 0,  EVAL("Gate" & idx + 1)
End Sub

Sub a_Rubbers_Hit(idx)
 	dim finalspeed
	If idx  = 0 Then
		RubberIndex = Idx + 1
		RubberAnimation.enabled = 1
	End If
  	finalspeed= 2 * (SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely))
 	If finalspeed > 20 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall)/2, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 4 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "RubberHit1", 0, Vol(ActiveBall) * 2, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "RubberHit2", 0, Vol(ActiveBall) * 2, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "RubberHit3", 0, Vol(ActiveBall) * 2, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
	End Select
End Sub

Sub LeftFlipper_Collide(parm)
 	dim finalspeed
  	finalspeed = (2 * SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely))
	If finalspeed > 5 Then RandomSoundFlipper()
End Sub

Sub LeftFlipper1_Collide(parm)
 	dim finalspeed
  	finalspeed = (2 * SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely))
	If finalspeed > 5 Then RandomSoundFlipper()
End Sub

Sub RightFlipper_Collide(parm)
 	dim finalspeed
  	finalspeed = (2 * SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely))
	If finalspeed > 5 Then RandomSoundFlipper()
End Sub

Sub RightFlipper1_Collide(parm)
 	dim finalspeed
  	finalspeed = (2 * SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely))
	If finalspeed > 5 Then RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
	If Vol(ActiveBall) > .003 Then
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall)*8, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall)*8, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall)*8, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
	End If
End Sub

Dim TempScore(2), TempPos(3), Position(5)
'***************Bubble Sort
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
	For HS = 1 to 4     							'look at all 5 saved high scores
		For CHy = 0 to 4   					    	'look at 4 player scores
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
'			TextBox3.text = ActiveScore(Flag) 'tells which high score is being entered
'			TextBox2.text = Flag
'			TextBox1.text =  Position(Flag) 'tells which player is entering values
			Initial(ActiveScore(Flag),1) = 1	'make first inital "A"
			For CHy = 2 to 3
				Initial(ActiveScore(Flag),CHy) = 0	'set other two to " "
			Next
			For CHy = 1 to 3
				EVAL("Initial" & CHy).image = HSiArray(Initial(ActiveScore(Flag),CHy))		'display the initals on the tape
			Next
			InitialTimer1.enabled = 1		'flash the first initial
			DynamicUpdatePostIt.enabled = 0		'stop the scrolling intials timer
			Playsound SoundFXDOF("Knocker",109,DOFPulse,DOFKnocker)
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
			Else										'if on the last initial then get ready yo exit the subroutine
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
	Playsound SoundFXDOF("Bell10",141,DOFPulse,DOFChimes)
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
'		TextBox3.text = ActiveScore(Flag) 	'tells which high score is being entered
'		TextBox2.text = Flag
'		TextBox1.text =  Position(Flag) 	'tells which player is entering values
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
    savevalue "4MBC", "HighScore(0)", HighScore(0)
    savevalue "4MBC", "HighScore(1)", HighScore(1)
    savevalue "4MBC", "HighScore(2)", HighScore(2)
    savevalue "4MBC", "HighScore(3)", HighScore(3)
    savevalue "4MBC", "HighScore(4)", HighScore(4)
    savevalue "4MBC", "Initial(0,1)", Initial(0,1)
    savevalue "4MBC", "Initial(0,2)", Initial(0,2)
    savevalue "4MBC", "Initial(0,3)", Initial(0,3)
    savevalue "4MBC", "Initial(1,1)", Initial(1,1)
    savevalue "4MBC", "Initial(1,2)", Initial(1,2)
    savevalue "4MBC", "Initial(1,3)", Initial(1,3)
    savevalue "4MBC", "Initial(2,1)", Initial(2,1)
    savevalue "4MBC", "Initial(2,2)", Initial(2,2)
    savevalue "4MBC", "Initial(2,3)", Initial(2,3)
    savevalue "4MBC", "Initial(3,1)", Initial(3,1)
    savevalue "4MBC", "Initial(3,2)", Initial(3,2)
    savevalue "4MBC", "Initial(3,3)", Initial(3,3)
    savevalue "4MBC", "Initial(4,1)", Initial(4,1)
    savevalue "4MBC", "Initial(4,2)", Initial(4,2)
    savevalue "4MBC", "Initial(4,3)", Initial(4,3)
    savevalue "4MBC", "Credit", Credit
    savevalue "4MBC", "FreePlay", FreePlay
	savevalue "4MBC", "Balls", Balls
	savevalue "4MBC", "ReplayEB", ReplayEB
	savevalue "4MBC", "ShowBallShadow", ShowBallShadow
	savevalue "4MBC", "MatchNumber", MatchNumber
	savevalue "4MBC", "Chime", Chime
End Sub

'*************Load Scores
Sub LoadHighScore
    dim temp
    temp = LoadValue("4MBC", "HighScore(0)")
    If (temp <> "") then HighScore(0) = CDbl(temp)
    temp = LoadValue("4MBC", "HighScore(1)")
    If (temp <> "") then HighScore(1) = CDbl(temp)
    temp = LoadValue("4MBC", "HighScore(2)")
    If (temp <> "") then HighScore(2) = CDbl(temp)
    temp = LoadValue("4MBC", "HighScore(3)")
    If (temp <> "") then HighScore(3) = CDbl(temp)
    temp = LoadValue("4MBC", "HighScore(4)")
    If (temp <> "") then HighScore(4) = CDbl(temp)
    temp = LoadValue("4MBC", "Initial(0,1)")
    If (temp <> "") then Initial(0,1) = CDbl(temp)
    temp = LoadValue("4MBC", "Initial(0,2)")
    If (temp <> "") then Initial(0,2) = CDbl(temp)
    temp = LoadValue("4MBC", "Initial(0,3)")
    If (temp <> "") then Initial(0,3) = CDbl(temp)
    temp = LoadValue("4MBC", "Initial(1,1)")
    If (temp <> "") then Initial(1,1) = CDbl(temp)
    temp = LoadValue("4MBC", "Initial(1,2)")
    If (temp <> "") then Initial(1,2) = CDbl(temp)
    temp = LoadValue("4MBC", "Initial(1,3)")
    If (temp <> "") then Initial(1,3) = CDbl(temp)
    temp = LoadValue("4MBC", "Initial(2,1)")
    If (temp <> "") then Initial(2,1) = CDbl(temp)
    temp = LoadValue("4MBC", "Initial(2,2)")
    If (temp <> "") then Initial(2,2) = CDbl(temp)
    temp = LoadValue("4MBC", "Initial(2,3)")
    If (temp <> "") then Initial(2,3) = CDbl(temp)
    temp = LoadValue("4MBC", "Initial(3,1)")
    If (temp <> "") then Initial(3,1) = CDbl(temp)
    temp = LoadValue("4MBC", "Initial(3,2)")
    If (temp <> "") then Initial(3,2) = CDbl(temp)
    temp = LoadValue("4MBC", "Initial(3,3)")
    If (temp <> "") then Initial(3,3) = CDbl(temp)
    temp = LoadValue("4MBC", "Initial(4,1)")
    If (temp <> "") then Initial(4,1) = CDbl(temp)
    temp = LoadValue("4MBC", "Initial(4,2)")
    If (temp <> "") then Initial(4,2) = CDbl(temp)
    temp = LoadValue("4MBC", "Initial(4,3)")
    If (temp <> "") then Initial(4,3) = CDbl(temp)
	temp = LoadValue("4MBC", "credit")
    If (temp <> "") then Credit = CDbl(temp)
	temp = LoadValue("4MBC", "FreePlay")
    If (temp <> "") then FreePlay = CDbl(temp)
    temp = LoadValue("4MBC", "balls")
    If (temp <> "") then Balls = CDbl(temp)
    temp = LoadValue("4MBC", "ReplayEB")
    If (temp <> "") then ReplayEB = CDbl(temp)
    temp = LoadValue("4MBC", "ShowBallShadow")
    If (temp <> "") then ShowBallShadow = CDbl(temp)
    temp = LoadValue("4MBC", "MatchNumber")
    If (temp <> "") then MatchNumber = CDbl(temp)
    temp = LoadValue("4MBC", "Chime")
    If (temp <> "") then Chime = CDbl(temp)
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
	If HighScore(0) > 999999 Then Shift = 0 :PComma.transx = 0
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
	If HighScore(Flag) > 999999 Then Shift = 0 :PComma.transx = 0
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
'	TextBox1.text = ScoreUpdate
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
	If HighScore(ScoreUpdate) > 999999 Then Shift = 0 :PComma.transx = 0
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
' PlaySound sound, 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
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
  PlaySound sound, 1, Vol, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'Set all as per ball position & speed, but Vol Multiplier may be used eg; PlaySoundAtBallVol "sound",3

Sub PlaySoundAtBallVol(sound, VolMult)
  PlaySound sound, 0, Vol(ActiveBall) * VolMult, Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub

'Set position as bumperX and Vol manually.

Sub PlaySoundAtBumperVol(sound, tableobj, Vol)
  PlaySound sound, 1, Vol, Pan(tableobj), 0,0,1, 1, AudioFade(tableobj)
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

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / table1.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioFade(ball) ' Can this be together with the above function ?
  Dim tmp
  tmp = ball.y * 2 / Table1.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
  Vol = Csng(BallVel(ball) ^2 / 2000)
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

Const tnob = 5 ' total number of balls
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

	' play the rolling sound for each ball

    For b = 0 to UBound(BOT)
      If BallVel(BOT(b) ) > 1 Then
        rolling(b) = True
        if BOT(b).z < 30 Then ' Ball on playfield
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        Else ' Ball on raised ramp
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*.5, Pan(BOT(b) ), 0, Pitch(BOT(b) )+50000, 1, 0, AudioFade(BOT(b) )
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
  If Table1.VersionMinor > 3 OR Table1.VersionMajor > 10 Then
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
  Else
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0
  End if
End Sub

' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


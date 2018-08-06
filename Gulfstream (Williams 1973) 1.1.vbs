'****************************************************************
'
'					  Gulfstream (Bally 1973)
'				  Script Code by Scottacus
'
'	Basic DOF config
'		101 Left Flipper, 102 Right Flipper
'		103 Left sling, 104 Right sling
'		105 Bumper1,  106 Bumper2, 107 Bumper3
'		108 - 114 Upper Rollovers
'		115 - 116 InLane Rollovers
'		117 & 118 OutLane Rollovers
'		119 Center  Target
'		120 - 121 Top Targets
'		122 Right Kicker, 123 Left Kicker
'		125 Ball Release
'	 	127 credit light
'		128 Knocker
'		129 - 130 Saucer Kickers
'		131 - 133 Roll Over Buttons
'		153 Chime1-10s, 154 Chime2-100s, 155 Chime3-1000s
'		160 Ball in Launch Lane
'
'		Code Flow
'									 EndGame
'										^
'		Start Game -> New Game -> Check Continue -> Release Ball -> Drain -> Score Bouns -> Advance Players -> Next Ball
'										^													                      |
'									EndGame = True <---------------------------------------------------------------
'***********************************************************************************************************************

' Thalamus 2018-07-23
' This table already had "Positional Sound Playback Functions".
' No special SSF tweaks yet.

Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

Const cGameName = "gulfstream_1973"

Dim Balls
Dim Replays
Dim Add1, Add10, Add100, Add1000
Dim HiSc
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
Dim i,j, f, ii, Object, Light, x, y
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
Dim Button
Dim FlipperArray
Dim Grid
Dim SGrid(10)
Dim BellRing
Dim PinkFlipperColor
Dim HSInitial0, HSInitial1, HSInitial2
Dim EnableInitialEntry
Dim HSi,HSx
Dim RoundHS, RoundHSPlayer
Dim ScoreMil, Score100K, Score10K, ScoreK, Score100, Score10, ScoreUnit
HSArray = Array("HS_0","HS_1","HS_2","HS_3","HS_4","HS_5","HS_6","HS_7","HS_8","HS_9","HS_Space","HS_Comma","HSi_1","HSi_2","HSi_3")
HSiArray = Array("HSi_0","HSi_1","HSi_2","HSi_3","HSi_4","HSi_5","HSi_6","HSi_7","HSi_8","HSi_9","HSi_10","HSi_11","HSi_12","HSi_13","HSi_14","HSi_15","HSi_16","HSi_17","HSi_18","HSi_19","HSi_20","HSi_21","HSi_22","HSi_23","HSi_24","HSi_25","HSi_26")
FlipperArray = Array("FlipperL_pnk_red","FlipperR_pnk_red","FlipperL_wht_red","FlipperR_wht_red")
Dim ShowBallShadow
BallSize = 50
BallMass = (Ballsize^3)/125000
Dim Chime
Dim Options
Dim HSUx
Dim ReplayEB
Dim TicTacToe
Dim ButtonGlow
Dim FirstBallOut

Sub Table1_init
	LoadEM
	MaxPlayers = 1
	Replay(1) = 68000
	Replay(2) = 86000
	Replay(3) = 99000
	Set SReels(1) = ScoreReel1
	Player = 1
	LoadHighScore
	Grid = 5
	Light5.state = 1

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
	If Chime ="" Then Chime = 0
	If PinkFlipperColor = "" Then PinkFlipperColor = 1

	SaveHS
	UpdatePostIt
	tilttxt.text = "TILT"
	credittxt.text = Credit
	For x = 1 to 4
		Score(x) = 0
	Next
	If B2SOn Then
		Controller.B2SSetCredits Credit
		Controller.B2SSetMatch 34, MatchNumber
		Controller.B2SSetGameOver 35,1
		Controller.B2SSetTilt 33,1
		Controller.B2SSetBallInPlay 32,0
		Controller.B2SSetPlayerUp 30,0
		If Credit > 0 Then DOF 127, DOFOn
		If FreePlay = 1 Then DOF 127, DOFOn
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

	If PinkFlipperColor = 1 Then
		RightFlipper.image = FlipperArray(1)
		LeftFlipper.image = FlipperArray(0)
	Else
		RightFlipper.image = FlipperArray(3)
		LeftFlipper.image = FlipperArray(2)
	End If

	If ShowBallShadow = 1 Then
		BallShadowUpdate.enabled = True
	Else
		BallShadowUpdate.enabled = False
	End If

	For i = 1 to MaxPlayers
		SReels(i).setvalue(Score(i))
	Next
	PlaySound "motor"
	Tilt=False
	State = False
	GameState
	InstructCard.image = "InstructionCard" & FreePlay
	CoinCard.image = "CoinCard" & ReplayEB & Balls

'*****************Trough Ball Creation
	Drain.CreateSizedBallWithMass Ballsize/2, BallMass
End Sub

'****************KeyCodes
Sub Table1_KeyDown(ByVal keycode)

	If EnableInitialEntry = True Then EnterIntitals(keycode)

	If keycode=AddCreditKey Then
		PlaySound "coinin"
		AddCredit
    End If

    If keycode=StartGameKey Then
		If EnableInitialEntry = False and OperatorMenu = 0 Then
			If FreePlay = 1 and Players < 1 and FirstBallOut = False Then StartGame
			If FreePlay = 0 and Credit > 0 and FirstBallOut = False and Players < 1 Then
				Credit = Credit - 1
				CreditTxt.text = Credit
				If B2SOn Then
					Controller.B2SSetCredits Credit
					DOF 127, DOFOff
				End If
				StartGame
			End If
		End If
	End If

	If keycode = PlungerKey Then
		Plunger.PullBack
		PlaySound "plungerpull",0,1,0.25,0.25
	End If

  If Tilt = False and State = True Then
	If keycode = LeftFlipperKey Then
		LeftFlipper.RotateToEnd
		PlayFieldSound "flipperup", 0, LeftFlipper
		If B2SOn Then DOF 101,DOFOn
		PlayFieldSound "Buzz", -1, LeftFlipper
	End If

	If keycode = RightFlipperKey Then
		RightFlipper.RotateToEnd
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
  End if

	If Keycode = MechanicalTilt Then
		Tilt = True
		Tilttxt.text = "TILT"
		If B2SOn Then Controller.B2SSetTilt 1
		Playsound"Tilt"
		TurnOff
	End If

    If keycode = LeftFlipperKey and State = False and OperatorMenu = 0 and EnableInitialEntry = 0 Then
        OperatorMenuTimer.Enabled = true
    end if

    If keycode = LeftFlipperKey and State = False and OperatorMenu = 1 Then
		Options = Options + 1
        If Options =  6 Then Options = 0
		OptionMenu.visible = true
        playsound "target"
        Select Case (Options)
            Case 0:
                OptionMenu.image = "FreeCoin" & FreePlay
            Case 1:
                OptionMenu.image = Balls & "Balls"
			Case 2:
				OptionMenu.image = "ReplayEB" & ReplayEB
            Case 3:
                OptionMenu.image = "BallShadow" & ShowBallShadow
			Case 4:
				OptionMenu.image = "Chime" & Chime
'			Case 5:
'				OptionMenu.image = "Flipper" & PinkFlipperColor
			Case 5:
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
			InstructCard.image = "InstructionCard" & FreePlay
			CoinCard.image = "CoinCard" & ReplayEB & Balls
			If FreePlay = 0 Then
				If Credit > 0 and B2SOn Then DOF 127, DOFOn
				If Credit < 1 and B2SOn Then DOF 127, DOFOff
			Else
				If B2SOn Then DOF 127, DOFOn
			End If
        Case 1:
            If Balls = 3 Then
                Balls = 5
              Else
                Balls = 3
            End If
            OptionMenu.image = Balls & "Balls"
			CoinCard.image = "CoinCard" & ReplayEB & Balls
		Case 2:
'			0 = add a ball, 1 = replay, 2 = novelty
			ReplayEB = ReplayEB + 1
			If ReplayEB = 3 Then ReplayEB = 0
			CoinCard.image = "CoinCard" & ReplayEB & Balls
			OptionMenu.image = "ReplayEB" & ReplayEB
        Case 3:
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
            OperatorMenu = 0
            SaveHS
			DynamicUpdatePostIt.enabled = 1
			OptionMenu.image = "FreeCoin" & FreePlay
            OptionMenu.visible = 0
			OptionsMenu.visible = 0
		End Select
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
'	If keycode = 30 Then


'************************End Of Test Keys****************************
End Sub

Sub Table1_KeyUp(ByVal keycode)

	If keycode = PlungerKey Then
		Plunger.Fire
		PlayFieldSound "Plunger", 1, Plunger
	End If

    If keycode = LeftFlipperKey Then
        OperatorMenuTimer.Enabled = False
    End If

   If Tilt = False and State = True Then
		If keycode = LeftFlipperKey Then
			LeftFlipper.RotateToStart
			PlayFieldSound "flipperdown", 0, LeftFlipper
			If B2SOn Then DOF 101,DOFOff
			StopSound "Buzz"
		End If

		If keycode = RightFlipperKey Then
			RightFlipper.RotateToStart
			PlayFieldSound "flipperdown", 0, RightFlipper
			If B2SOn Then DOF 102,DOFOff
			StopSound "Buzz"
		End If
   End if

    If keycode = 203 Then Cleft = 0' Left Arrow

    If keycode = 200 Then Cup = 0' Up Arrow

    If keycode = 208 Then Cdown = 0' Down Arrow

    If keycode = 205 Then Cright = 0' Right Arrow

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

'***********Start Up Game
Sub StartGame
	If State = False Then
		BallinPlay = Balls
		If B2SOn Then
			Controller.B2SSetCredits Credit
			Controller.B2SSetBallinPlay 32, Ballinplay
			Controller.B2SSetPlayerup 30, 1
			Controller.B2SSetCanPlay 31, 1
			Controller.B2SSetGameOver 0
		End If
		DynamicUpdatePostIt.enabled = 0
		UpdatePostIt
		Tilt = False
		State = True
		GameState
		PlaySound "initialize"
		Players = 1
		ShutOffLights
		CanPlayTxt.text = Players
		Score(1) = 0
		If B2SOn then controller.B2SSetScorePlayer 1, Score(1)
		SReels(1).setvalue(0)
		For x = 1 to 9
			EVAL("GridLight" & x).state = 0
			If x < 8 Then EVAL("SpecialLight" & x).state = 0
			SGrid(x) = 0
		Next
		ShutOffLights
		newgame.enabled=true
	Else If  State = True and Players < MaxPlayers and BallinPlay = Balls Then
		Players = Players + 1
		CanPlayTxt.text = Players
		CreditTxt.text = Credit
		If B2SOn then
			Controller.B2SSetCredits Credit
			Controller.B2SSetCanplay 31, Players
		End If
		ShutOffLights
		Playsound "cluper"
		End If
	End If
End Sub

'***************Start a New Game
Sub NewGame_timer
	Player = 1
	For i = 1 to MaxPlayers
	    Score(i) = 0
		Rep(i) = 0
	Next
	RoundHS = 0
	RoundHSPlayer = 1
	If B2SOn Then
	  For i = 1 to MaxPlayers
		Controller.B2SSetScorePlayer i, score(i)
	  Next
	End If
    EndGame = 0
	RoundHS = 0
	RoundHSPlayer = 1
'	ShootAgainLight.state = 0
	If B2SOn Then controller.B2SSetShootAgain 36,0
	GameState
    BIPText.text = BallinPlay
	CheckContinue.enabled = 1
	NewGame.enabled = False
	For f = 1 to 3
		EVAL("Bumper" & f).hashitevent = 1
	Next
	Player1.intensityscale = 1.5
End Sub

'*************Check for Continuing Game
Sub CheckContinue_Timer
	If EndGame = 1 Then
		TurnOff
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
		HsTxt.text = HiSc
		SaveHS
		If B2SOn Then
			Controller.B2SSetGameOver 35,1
			Controller.B2SSetballinplay 32, 0
			Controller.B2SSetPlayerUp 30, 0
			Controller.B2SSetcanplay 31, 0
			If Credit > 0 Then DOF 127, DOFOn
			If FreePlay = 1 Then DOF 127, DOFOn
		End If
		For each Light in GIlights:light.state = 0:Next
		For Each ButtonGlow in Buttons: ButtonGlow.blenddisablelighting = 0: Next
		CheckContinue.enabled = 0
	Else
		If BallInLane = False and BIP = 0 then
		BIPText.text = BallinPlay
		ReleaseBall
	End If
	CheckContinue.enabled = 0
  End If
End Sub


'***************Drain and Release Ball
Dim BIP : BIP = 0

Sub Drain_Hit()
	BIP = BIP - 1
	PlayFieldSound "fx_drain", 0, Drain
	If B2SOn Then DOF 124, DOFPulse
	For each Light in BumperLights:light.state = 0:Next
	SpecialLight.state = 0
	If BIP = 0 Then ScoreBonus
End Sub

Sub ReleaseBall
	If BIP = 0 Then
		PlayFieldSound "kickerkick", 0, Drain
		Drain.kick 62, 62
		If B2SOn Then DOF 125, DOFPulse
		Playsound "metalhit_thin"
		BIP = BIP + 1
		Launched = 0
		If B2SOn Then Controller.B2SSetballinplay 32, BallinPlay
	End If
End Sub

'*******************Check for Bonus Score
Sub ScoreBonus
	AdvancePlayers
End Sub

'**************Advance Players
Sub AdvancePlayers
	If Players = 1 or Player = Players Then
		Player=1
		EVAL("Player" & players).intensityscale = 1
		Player1.intensityscale = 1.5
		NextBall
	Else
		Player = Player + 1
		EVAL("Player" & (player-1)).intensityscale = 1
		EVAL("Player" & player).intensityscale = 1.5
		NextBall
	End If
		If B2SOn Then Controller.B2SSetPlayerup 30, Player
End Sub

'*********************Next Ball
Sub NextBall
    If Tilt = True Then
	  For f = 1 to 3
		EVAL("Bumper" & f).hashitevent = 1
	  Next
      Tilt = False
      TiltTxt.text = " "
		If B2SOn Then
			Controller.B2SSetTilt 33,0
			Controller.B2SSetData 1, 1
		End If
    End If
	If B2SOn Then controller.B2SSetShootAgain 36,0
	If Player = 1 Then BallinPlay = BallinPlay - 1
	If BallinPlay = 0 Then
		PlaySound "GameOver"
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
	If Credit > 25 then Credit = 25
	CreditTxt.text = Credit
	If B2SOn Then
		Controller.B2SSetCredits Credit
		If Credit > 0 Then DOF 127, DOFOn
	End If
End Sub

'************Game State Check
Sub GameState
	If State = False Then
		For each light in GIlights:light.state=0: Next
		For Each ButtonGlow in Buttons: ButtonGlow.blenddisablelighting = 0: Next
		GamOv.text="Game Over"
		If B2SOn then Controller.B2SSetGameOver 35,1
		Player1.intensityscale = 1
		leftmetallow_prim.blenddisablelighting = 0.08
		leftmetalhigh_prim.blenddisablelighting = 0.08
		rightmetallow_prim.blenddisablelighting = 0.08
		rightmetalhigh_prim.blenddisablelighting = 0.08
		aprontop_prim.blenddisablelighting = 0.08
		backmetal_prim.blenddisablelighting = 0
		woodrails_prim.image = "woodrailsGIOFF"
		outers_prim.image = "outers.001GIOFF"
		Table1.ColorGradeImage = "ColorGradeLUT256x16_GIoff"
		flasher1.visible=0
		flasher2.visible=1
	Else
		For each Light in GIlights:Light.state=1: Next
		For Each ButtonGlow in Buttons: ButtonGlow.blenddisablelighting = .1: Next
		GamOv.text=""
		MatchTxt.text= ""
		TiltTxt.text=" "
		If B2SOn then
			Controller.B2SSetTilt 33,0
			Controller.B2SSetMatch 34,0
			Controller.B2SSetGameOver 35,0
		End If
		leftmetallow_prim.blenddisablelighting = 0.5
		leftmetalhigh_prim.blenddisablelighting = 0.5
		rightmetallow_prim.blenddisablelighting = 0.5
		rightmetalhigh_prim.blenddisablelighting = 0.5
		backmetal_prim.blenddisablelighting = 0.5
		aprontop_prim.blenddisablelighting = 0.5
		woodrails_prim.image = "woodrailsGION"
		outers_prim.image = "outers.001"
		Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1"
		flasher1.visible=1
		flasher2.visible=0
	End If
End Sub

'*************Ball in Launch Lane
Sub BallHome_hit
	BallREnabled = 1
	If B2SOn then DOF 160, DOFOn
	Set ControlBall = ActiveBall
    contballinplay = true
End Sub

Sub BallHome_unhit
	If B2SOn Then DOF 160, DOFOff
End Sub

Sub EndControl_Hit()                '******* for ball control script
    contballinplay = false
End Sub

'***************Match
Sub Match
   y = int(rnd(1)*9)
    MatchNumber = y
		If B2SOn Then
			If MatchNumber = 0 Then
				Controller.B2SSetMatch 34, 10
			Else
				Controller.B2SSetMatch 34, MatchNumber
			End If
		End If
	MatchTxt.text = MatchNumber
	If MatchNumber = 0 then Matchtxt.text = "00"

	For i = 1 to Players
		MatchScoreTxt.text =  (Score(i) mod 100)
		If MatchNumber = (Score(i) mod 100) Then
			AddCredit
			If B2SOn Then
				DOF 128, DOFPulse
			Else
				Playsound "knocker"
			End If
		End If
    Next
End Sub

'***********Rotate Flipper Shadow
Sub FlipperShadowUpdate_Timer
	FlipperLSh.RotZ = LeftFlipper.CurrentAngle
	FlipperRSh.RotZ = RightFlipper.CurrentAngle
End Sub

'************** Shut Off Lights
Sub ShutOffLights
	For x = 10 to 13
		EVAL("Light" & x).state = 0
	Next
End Sub

'************** Bumpers
Sub Bumpers_hit(Index)
	If Tilt = False Then
		Select Case (Index)
			Case 0: PlayFieldSound "fx_bumper1", 0, Bumper1
					If B2SOn Then DOF 105,DOFPulse
				If BumperLight1.state = 0 then
					Addscore 10
				Else
					AddScore 100
				End If
			Case 1: PlayFieldSound "fx_bumper2", 0, Bumper2
					If B2SOn Then DOF 106,DOFPulse
				If BumperLight1.state = 0 then
					Addscore 10
				Else
					AddScore 100
				End If
			Case 2: PlayFieldSound "fx_bumper3", 0, Bumper2
					If B2SOn Then DOF 107,DOFPulse
				If BumperLight1.state = 0 then
					Addscore 10
				Else
					AddScore 100
				End If
		End Select
	End If
End Sub

'************** Saucers
Sub Saucers_Hit(Index)
	If Tilt = False Then
		If Light10.state = 0 Then
			AddScore 500
		Else
			AddScore 5000
			SpecialGrid
		End If
	End If
	Select Case Index
		Case 0: LeftSaucerKicker.Timerenabled = 1
		Case 1: RightSaucerKicker.Timerenabled = 1
	End Select
End Sub

'************** Kickers
Sub LeftSaucerKicker_Timer
	LeftSaucerKicker.Kick 92 + (3 * Rnd()), 6
	Playsound SoundFXDOF("HoleKick",129,DOFPulse,DOFContactors)
	me.timerenabled = 0
End Sub

Sub RightSaucerKicker_Timer
	RightSaucerKicker.Kick 268 + ( 3 * Rnd()), 6
	Playsound SoundFXDOF("HoleKick",130,DOFPulse,DOFContactors)
	me.timerenabled = 0
End Sub


'************** Slings
Sub LeftSlingShot_Slingshot
	Playsound SoundFXDOF("left_slingshot",103,DOFPulse,DOFContactors)
    LSling.Visible = 0
    LSling1.Visible = 1
    LeftSling.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
	AddScore 10
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3: LSling1.Visible = 0:LSling2.Visible = 1:LeftSling.TransZ = -10
        Case 4: LSling2.Visible = 0:LSling.Visible = 1:LeftSling.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
	Playsound SoundFXDOF("right_slingshot",104,DOFPulse,DOFContactors)
    RSling.Visible = 0
    RSling1.Visible = 1
    RightSling.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
	AddScore 10
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3: RSling1.Visible = 0:RSling2.Visible = 1:RightSling.TransZ = -10
        Case 4: RSling2.Visible = 0:RSling.Visible = 1:RightSling.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Dim SpecialLanes
'*************** Triggers
Sub Triggers_hit(Index)
	If Tilt = False Then
		AddScore 500
		EVAL ("SpecialLight" & Index + 1).state = 1
		For x = 1 to 7
			If EVAL("SpecialLight" & x).state = 1 then SpecialLanes = SpecialLanes + 1
		Next
		If B2SOn Then DOF (Index +108), DOFPulse
		If SpecialLanes = 7 Then
			Light12.state = 1
			Light13.state = 1
		End If
	End If
	SpecialLanes = 0
End Sub

Sub LeftOutLane_hit
	If Tilt = False Then
		AddScore 1000
		If B2SOn Then DOF 115, DOFPulse
	End If
End Sub

Sub RightOutLane_hit
	If Tilt = False Then
		AddScore 1000
		If B2SOn Then DOF 116, DOFPulse
	End If
End Sub

Sub InLanes_Hit(Index)
	If Tilt = False Then
		If Light12.state = 1 Then
			AddScore 5000
			SpecialCollect
		Else
			AddScore 500
		End If
		If B2SOn Then DOF (Index + 117), DOFPulse
	End If
End Sub

Sub BallLaunchedTrigger_Hit
	FirstBallOut = True
	Launched = Launched + 1
	If CLng(Launched) Mod 2 > 0 Then
		If B2SOn Then DOF 161 ,DOFPulse
	End If
	If B2SOn Then DOF 127, DOFOff
End Sub

Dim ROButton
Sub RollOvers_Hit(Index)
	ROButton = Index
	RollOver.enabled = 1
	If Tilt = False Then
		AddScore 100
		If B2SOn Then DOF (Index + 131), DOFPulse
	End If
End Sub

Sub RollOver_Timer
	Button = Button +1
	Select Case Button
		Case 1: EVAL ("RollOverButton" & ROButton).transz = -1
		Case 2: EVAL ("RollOverButton" & ROButton).transz = -3
		Case 3: EVAL ("RollOverButton" & ROButton).transz = 3
		Case 4: EVAL ("RollOverButton" & ROButton).transz = 1
		Button = -1
		RollOver.enabled = 0
	End Select
End Sub
'**************Targets
Sub Targets_hit(Index)
	If Tilt = False Then
		AddScore 100
		For each Light in BumperLights:light.state = 1:Next
	End If
	If B2SOn Then DOF (Index + 120), DOFPulse
End Sub

Sub TargetGrid_hit
	If Tilt = False Then
		EVAL("GridLight" & LightOn).state = 1
		SGrid(LightOn) = 1
		CheckGrid
		If SpecialLight.state = 1 Then
			AddScore 5000
		Else
			AddScore 10
		End If
		If B2SOn Then DOF 119, DOFPulse
	End If
End Sub


'*************Check Grid
Sub CheckGrid
	For x = 1 to 7 step 3
		If SGrid(x) > 0 Then
			If SGrid(x) = SGrid (x + 1) and SGrid(x) = SGrid (x + 2) Then TicTacToe = True
		End If
	Next

	For x = 1 to 3
		If SGrid(x) > 0 Then
			If SGrid(x) = SGrid(x+3) and SGrid(x) = SGrid(x+6) Then TicTacToe = True
		End If
	Next

	If SGrid(1) > 0 Then
		If SGrid(1) = SGrid(5) and SGrid(1) = SGrid(9) Then TicTacToe = True
	End If

	If SGrid(3) > 0 Then
		If SGrid(3) = SGrid(5) and SGrid(3) = SGrid(7) Then TicTacToe = True
	End If

	If TicTacToe = True Then
		Light10.state = 1: Light11.state = 1: SpecialLight.state = 1
	End If

	TicTacToe = False

	If SGrid(1) = 1 and SGrid(3) = 1 and SGrid(7) =1 and SGrid(9) = 1 Then ReplayCollect

End Sub


'**************Special
Sub SpecialGrid
	If B2SOn Then
		DOF 128, DOFPulse
	Else
		Playsound "knocker"
	End If
	For x = 1 to 9
		SGrid(x) = 0
		EVAL("GridLight" & x).state = 0
	Next
	Light10.State = 0
	Light11.State = 0
	SpecialLight.state = 0
	If ReplayEB = 0 Then BallinPlay = BallinPlay + 1
	If ReplayEB = 1 Then AddCredit
	If ReplayEB = 2 Then AddScore 1000
	If BallinPlay > 10 Then BallinPlay = 10
	If B2SOn Then Controller.B2SSetBallInPlay 32, BallinPlay
End Sub

Sub SpecialCollect
	If B2SOn Then
		DOF 128, DOFPulse
	Else
		Playsound "knocker"
	End If
	Light12.State = 0
	Light13.State = 0
	For x = 1 to 7
		EVAL("SpecialLight" & x).State = 0
	Next
	If ReplayEB = 0 Then BallinPlay = BallinPlay + 1
	If ReplayEB = 1 Then AddCredit
	If ReplayEB = 2 Then AddScore 1000
	If BallinPlay > 10 Then BallinPlay = 10
	If B2SOn Then Controller.B2SSetBallInPlay 32, BallinPlay
End Sub

Sub ReplayCollect
	For x = 10 to 13
		EVAL("Light" & x).State = 0
	Next
	If B2SOn Then
		DOF 128, DOFPulse
	Else
		Playsound "knocker"
	End If
	For x = 1 to 9
		SGrid(x) = 0
		EVAL("GridLight" & x).state = 0
	Next
	AddCredit
End Sub


'************Check if Ball Out of Launch Lane
Sub BallsInPlay_hit
	If BallREnabled = 1 Then
		BallREnabled = 0
		BallInLane = False
	End If
	me.timerenabled = 1
End Sub

'***************Shooter Lane Gate Animation
Sub BallsInPlay_Timer
Pgate.rotz = Gate.CurrentAngle * .6
End Sub

'***************Advance Grid
Dim LightOn
Sub AdvanceGrid
	Grid = Grid + 1
	If Grid = 11 then Grid = 1
	LightOn = Grid
	If Grid <10 Then
		EVAL("Light" & Grid).state = 1
	Else
		Light5.state = 1
		LightOn = 5
	End If

	Select Case Grid
		Case 1: Light5.state = 0
		Case 5: Light4.state = 0: Light9.state = 0
		Case 2: Light1.state = 0
		Case 3: Light2.state = 0
		Case 4: Light3.state = 0
		Case 6: Light5.state = 0
		Case 7: Light6.state = 0
		Case 8: Light7.state = 0
		Case 9: Light8.state = 0
		Case 10: Light9.state = 0
	End Select

End Sub

'***************Scoring Routine
Sub AddScore(points)
  If Tilt = False Then
		If Points = 10 Then: BellRing = 1 : BellTimer10.enabled = 1
		If Points = 100 Then: BellRing = 1 : BellTimer100.enabled = 1
		If Points = 500 Then: BellRing = 5 : BellTimer100.enabled = 1
		If Points = 1000 Then: BellRing = 1 : BellTimer1000.enabled = 1
		If Points = 3000 Then: BellRing = 3 : BellTimer1000.enabled = 1
		If Points = 5000 Then: BellRing = 5 : BellTimer1000.enabled = 1
	End If
End Sub

Dim ReplayX, RepAwarded(5), Replay(7), ReplayText(3)

Sub TotalUp(Points)
	Score(Player) = Score(Player) + Points
	SReels(Player).addvalue(Points)
	If B2SOn Then Controller.B2SSetScorePlayer Player, Score(Player)
	For ReplayX = (Rep(Player) + 1) to 3
		If Score(Player) >= Replay(ReplayX) Then
			If ReplayEB = 0 Then
				AddCredit
			Else
'				ShootAgainLight.state = 1
				If B2SOn Then controller.B2SSetShootAgain 36,1
			End If
			Rep(Player) = Rep(Player) + 1
			If B2SOn Then
				DOF 128, DOFPulse
			Else
				Playsound "knocker"
			End If
		End If
	Next
End Sub

'*************** Bell Timers

Sub BellTimer10_Timer
	If Chime = 0 Then
		Playsound "Bell10"
	Else
		If B2SOn Then DOF 153,DOFPulse
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
		If B2SOn Then DOF 154,DOFPulse
	End If
	Playsound "Reel1"
	AdvanceGrid
	BellRing = BellRing - 1
	TotalUp 100
	If BellRing < 1 Then
		BellTimer100.enabled = 0
	End If
End Sub

Sub BellTimer1000_Timer
	If Chime = 0 Then
		Playsound "Bell1000"
	Else
		If B2SOn Then DOF 155,DOFPulse
	End If
	Playsound "Reel1"
	AdvanceGrid
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
	 End If
	Else
	 TiltSens = 0
	 Tilttimer.Enabled = True
	End If
End Sub

Sub Tilttimer_Timer()
	Tilttimer.Enabled = False
End Sub

'***********Rotate Flipper Shadow
Sub FlipperShadowUpdate_Timer
	FlipperLSh.RotZ = LeftFlipper.CurrentAngle
	FlipperRSh.RotZ = RightFlipper.CurrentAngle
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

'************Shut Down and De-energize for Tilt
Sub TurnOff
	For i= 1 to 3
		EVAL("Bumper" & i).hashitevent = 0
	Next
  	LeftFlipper.RotateToStart
	StopSound "Buzz"
	If B2SOn Then DOF 101, DOFOff
	RightFlipper.RotateToStart
	StopSound "Buzz1"
	If B2SOn Then DOF 102, DOFOff
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

'*************Hit Sound Routines
Sub PlayFieldSound (SoundName, Looper, TableObject)
	PlaySound SoundName, Looper, 1, Pan(TableObject), 0, 0, 0, 0, Fade(TableObject)
End Sub

Sub a_Pins_Hit (idx)
	PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, Fade(ActiveBall)
End Sub

Sub a_Targets_Hit (idx)
	PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, Fade(ActiveBall)
End Sub

Sub a_Metals_Thin_Hit (idx)
	PlaySound "metalhit_thin", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, Fade(ActiveBall)
End Sub

Sub a_Metals_Medium_Hit (idx)
	PlaySound "metalhit_medium", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, Fade(ActiveBall)
End Sub

Sub a_Metals2_Hit (idx)
	PlaySound "metalhit2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, Fade(ActiveBall)
End Sub

Sub a_Gates_Hit (idx)
	PlaySound "Gate", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, Fade(ActiveBall)
End Sub

Sub a_Spinner_Spin
	PlaySound "fx_spinner",0,.25,0,0.25
End Sub

Sub a_Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed = (SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely))
 	If finalspeed > 20 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, Fade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RubberWheel_hit
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, Fade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End sub

Sub a_Posts_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 16 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, Fade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 16 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, Fade(ActiveBall)
		Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, Fade(ActiveBall)
		Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, Fade(ActiveBall)
	End Select
End Sub

Sub LeftFlipper_Collide(parm)
 	dim finalspeed
  	finalspeed = (2 * SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely))
	If finalspeed > 5 Then RandomSoundFlipper()
End Sub

Sub RightFlipper_Collide(parm)
 	dim finalspeed
  	finalspeed = (2 * SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely))
	If finalspeed > 5 Then RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, Fade(ActiveBall)
		Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, Fade(ActiveBall)
		Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, Fade(ActiveBall)
	End Select
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
				DOF 128, DOFPulse
			Else
				Playsound "knocker"
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
	If Chime = 1 Then
		If B2SOn Then DOF 153, DOFPulse
	Else
		Playsound "Bell10"
	End If
	Flag = Flag - 1
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
    savevalue "Gulfstream", "HighScore(0)", HighScore(0)
    savevalue "Gulfstream", "HighScore(1)", HighScore(1)
    savevalue "Gulfstream", "HighScore(2)", HighScore(2)
    savevalue "Gulfstream", "HighScore(3)", HighScore(3)
    savevalue "Gulfstream", "HighScore(4)", HighScore(4)
    savevalue "Gulfstream", "Initial(0,1)", Initial(0,1)
    savevalue "Gulfstream", "Initial(0,2)", Initial(0,2)
    savevalue "Gulfstream", "Initial(0,3)", Initial(0,3)
    savevalue "Gulfstream", "Initial(1,1)", Initial(1,1)
    savevalue "Gulfstream", "Initial(1,2)", Initial(1,2)
    savevalue "Gulfstream", "Initial(1,3)", Initial(1,3)
    savevalue "Gulfstream", "Initial(2,1)", Initial(2,1)
    savevalue "Gulfstream", "Initial(2,2)", Initial(2,2)
    savevalue "Gulfstream", "Initial(2,3)", Initial(2,3)
    savevalue "Gulfstream", "Initial(3,1)", Initial(3,1)
    savevalue "Gulfstream", "Initial(3,2)", Initial(3,2)
    savevalue "Gulfstream", "Initial(3,3)", Initial(3,3)
    savevalue "Gulfstream", "Initial(4,1)", Initial(4,1)
    savevalue "Gulfstream", "Initial(4,2)", Initial(4,2)
    savevalue "Gulfstream", "Initial(4,3)", Initial(4,3)
    savevalue "Gulfstream", "Credit", Credit
    savevalue "Gulfstream", "FreePlay", FreePlay
	savevalue "Gulfstream", "Balls", Balls
	savevalue "Gulfstream", "ReplayEB", ReplayEB
	savevalue "Gulfstream", "ShowBallShadow", ShowBallShadow
	savevalue "Gulfstream", "Chime", Chime
	savevalue "Gulfstream", "PinkFlipperColor", PinkFlipperColor
	savevalue "Gulfstream", "MatchNumber", MatchNumber
End Sub

'*************Load Scores
Sub LoadHighScore
    dim temp
    temp = LoadValue("Gulfstream", "HighScore(0)")
    If (temp <> "") then HighScore(0) = CDbl(temp)
    temp = LoadValue("Gulfstream", "HighScore(1)")
    If (temp <> "") then HighScore(1) = CDbl(temp)
    temp = LoadValue("Gulfstream", "HighScore(2)")
    If (temp <> "") then HighScore(2) = CDbl(temp)
    temp = LoadValue("Gulfstream", "HighScore(3)")
    If (temp <> "") then HighScore(3) = CDbl(temp)
    temp = LoadValue("Gulfstream", "HighScore(4)")
    If (temp <> "") then HighScore(4) = CDbl(temp)
    temp = LoadValue("Gulfstream", "Initial(0,1)")
    If (temp <> "") then Initial(0,1) = CDbl(temp)
    temp = LoadValue("Gulfstream", "Initial(0,2)")
    If (temp <> "") then Initial(0,2) = CDbl(temp)
    temp = LoadValue("Gulfstream", "Initial(0,3)")
    If (temp <> "") then Initial(0,3) = CDbl(temp)
    temp = LoadValue("Gulfstream", "Initial(1,1)")
    If (temp <> "") then Initial(1,1) = CDbl(temp)
    temp = LoadValue("Gulfstream", "Initial(1,2)")
    If (temp <> "") then Initial(1,2) = CDbl(temp)
    temp = LoadValue("Gulfstream", "Initial(1,3)")
    If (temp <> "") then Initial(1,3) = CDbl(temp)
    temp = LoadValue("Gulfstream", "Initial(2,1)")
    If (temp <> "") then Initial(2,1) = CDbl(temp)
    temp = LoadValue("Gulfstream", "Initial(2,2)")
    If (temp <> "") then Initial(2,2) = CDbl(temp)
    temp = LoadValue("Gulfstream", "Initial(2,3)")
    If (temp <> "") then Initial(2,3) = CDbl(temp)
    temp = LoadValue("Gulfstream", "Initial(3,1)")
    If (temp <> "") then Initial(3,1) = CDbl(temp)
    temp = LoadValue("Gulfstream", "Initial(3,2)")
    If (temp <> "") then Initial(3,2) = CDbl(temp)
    temp = LoadValue("Gulfstream", "Initial(3,3)")
    If (temp <> "") then Initial(3,3) = CDbl(temp)
    temp = LoadValue("Gulfstream", "Initial(4,1)")
    If (temp <> "") then Initial(4,1) = CDbl(temp)
    temp = LoadValue("Gulfstream", "Initial(4,2)")
    If (temp <> "") then Initial(4,2) = CDbl(temp)
    temp = LoadValue("Gulfstream", "Initial(4,3)")
    If (temp <> "") then Initial(4,3) = CDbl(temp)
	temp = LoadValue("Gulfstream", "credit")
    If (temp <> "") then Credit = CDbl(temp)
	temp = LoadValue("Gulfstream", "FreePlay")
    If (temp <> "") then FreePlay = CDbl(temp)
    temp = LoadValue("Gulfstream", "balls")
    If (temp <> "") then Balls = CDbl(temp)
    temp = LoadValue("Gulfstream", "ReplayEB")
    If (temp <> "") then ReplayEB = CDbl(temp)
    temp = LoadValue("Gulfstream", "ShowBallShadow")
    If (temp <> "") then ShowBallShadow = CDbl(temp)
    temp = LoadValue("Gulfstream", "Chime")
    If (temp <> "") then Chime = CDbl(temp)
    temp = LoadValue("Gulfstream", "PinkFlipperColor")
    If (temp <> "") then PinkFlipperColor = CDbl(temp)
    temp = LoadValue("Gulfstream", "MatchNumber")
    If (temp <> "") then MatchNumber = CDbl(temp)
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
	SaveHS
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


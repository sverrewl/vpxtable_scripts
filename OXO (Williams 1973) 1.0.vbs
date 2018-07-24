'****************************************************************
'
'					  OXO (Bally 1973)
'				  Script Code by Scottacus
'				   Artwork by HauntFreaks
'					   October 2017
'					   Version  1.0
'
'	Basic DOF config
'		101 Left Flipper, 102 Right Flipper
'		103 Left sling, 104 Right sling
'		105 Bumper1,  106 Bumper2, 107 Bumper3
'		109 Launch Ball
'		108 - 112 Upper X & O Rollovers
'		113 A Target, 114 B Target
'		115 & 116 L&R Mid OX Rollovers
'		117 & 118 L&R InLane OX Rollovers
'		121 Center OX Target
'		122 Right Kicker, 123 Left Kicker
'		124 Drain, 125 Ball Release
'	 	127 credit light
'		128 Knocker, 129 Knocker and Kicker Strobe
'		160 Ball In Shooter Lane
'		153 Chime1-10s, 154 Chime2-100s, 155 Chime3-1000s
'
'		Code Flow
'									 EndGame
'										^
'		Start Game -> New Game -> Check Continue -> Release Ball -> Drain -> Score Bouns -> Advance Players -> Next Ball
'										^													                      |
'									EndGame = True <---------------------------------------------------------------
'***********************************************************************************************************************

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' No special SSF tweaks yet.


Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

Const cGameName = "OXO"

Dim Balls
Dim Replays
Dim Replay1, Replay2, Replay3
Dim ReplayText(3)
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
Dim GateState
Dim Grid(10)
Dim SGrid(10)
Dim Square
Dim LastSquare
Dim TicTacToe
Dim TicTacToeBonus
Dim OXO
Dim TargetA, TargetB
Dim AB
Dim RelGateHit
Dim BellRing
Dim PinkFlipperColor
Dim GridAdvance
Dim GridFill
Dim SpecialGrid
Dim Initial(4)
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

'*************************************Set FreePlay*******************************************

FreePlay = True 'change to False for coin play

'**********************************Select Flipper Color**************************************

PinkFlipperColor = True 'Change to False for standard white rubbers on the flippers

'**********************************Ball Shadow On/Off****************************************

ShowBallShadow = True  'Change to False to hide ball shadow

'********************************************************************************************
Sub Table1_init
	LoadEM
	MaxPlayers=4
	Replay1 = 68000
	Replay2 = 86000
	Replay3 = 99000
	Set SReels(1) = ScoreReel1
	Set SReels(2) = ScoreReel2
	Set SReels(3) = ScoreReel3
	Set SReels(4) = ScoreReel4
	Player = 1
	LoadHighScore
	Balls = 5
	If HiSc = "" Then HiSc=10000
	HSTxt.text=HiSc
	If Initial(1) = "" Then
		Initial(1) = 19: Initial(2) = 5: Initial(3) = 13
	End If
	UpdatePostIt
	UpdateGrid
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

	If PinkFlipperColor = True Then
		RightFlipper.image = FlipperArray(1)
		LeftFlipper.image = FlipperArray(0)
	Else
		RightFlipper.image = FlipperArray(3)
		LeftFlipper.image = FlipperArray(2)
	End If

	If ShowBallShadow = True Then
		BallShadowUpdate.enabled = True
	Else
		BallShadowUpdate.enabled = False
	End If

	For i = 1 to MaxPlayers
		SReels(i).setvalue(Score(i))
	Next

	ClearGrid
	OXO = 1
	AB = False
	TicTacToe = False
	TicTacToeBonus = False
	PlaySound "motor"
	Tilt=False
	State = False
	GameState
	GateState = 1
	LastSquare = 0
	ABGateWallOpen.isdropped = True
	ABGateWallClosed.isdropped = False

'*****************Trough Ball Creation
	Drain.CreateSizedBallWithMass Ballsize/2, BallMass
End Sub

'****************KeyCodes
Sub Table1_KeyDown(ByVal keycode)

	If EnableInitialEntry = True Then EnterIntitals(keycode)

	If keycode=AddCreditKey Then
		PlaySound "coinin"
		coindelay.enabled = True
    End If

    If keycode=StartGameKey Then
		If EnableInitialEntry = False Then
			If FreePlay = True Then StartGame
			If FreePlay = False and Credit > 0 and Players < 4 Then
				Credit = Credit - 1
				CreditTxt.text = Credit
				If B2SOn Then Controller.B2SSetCredits Credit
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
		PlaySound SoundFXDOF("flipperup",101,DOFOn,DOFContactors), 0, .67, -0.05, 0.05
		PlaySound "Buzz", -1,.67, -0.05, 0.05
	End If

	If keycode = RightFlipperKey Then
		RightFlipper.RotateToEnd
		PlaySound SoundFXDOF("flipperup",102,DOFOn,DOFContactors), 0, .67, 0.05, 0.05
		PlaySound "Buzz1", -1,.67, 0.05, 0.05
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

	If Keycode = 20 Then 'The "t" key tilts the machine for machines with mechanical plub bobs
		Tilt = True
		Tilttxt.text = "TILT"
		If B2SOn Then Controller.B2SSetTilt 1
		Playsound"Tilt"
		TurnOff
	End If
End Sub

Sub Table1_KeyUp(ByVal keycode)

	If keycode = PlungerKey Then
		Plunger.Fire
		PlaySound "plunger",0,1,0.25,0.25
	End If

   If Tilt = False and State = True Then
	If keycode = LeftFlipperKey Then
		LeftFlipper.RotateToStart
		PlaySound SoundFXDOF("flipperdown",101,DOFOff,DOFContactors), 0, 1, -0.05, 0.05
		StopSound "Buzz"
	End If

	If keycode = RightFlipperKey Then
		RightFlipper.RotateToStart
		PlaySound SoundFXDOF("flipperdown",102,DOFOff,DOFContactors), 0, 1, 0.05, 0.05
		StopSound "Buzz1"
	End If
   End if
End Sub

'***********Rotate Flipper Shadow
Sub FlipperShadowUpdate_Timer
	FlipperLSh.RotZ = LeftFlipper.CurrentAngle
	FlipperRSh.RotZ = RightFlipper.CurrentAngle
End Sub

'***********Start Up Game
Sub StartGame
	If State = False Then
		BallinPlay=1
		If B2SOn Then
			Controller.B2SSetCredits Credit
			Controller.B2SSetBallinPlay 32, Ballinplay
			Controller.B2SSetPlayerup 30, 1
			Controller.B2SSetCanPlay 31, 1
			Controller.B2SSetGameOver 0
		End If
		Tilt = False
		State = True
		GameState
		PlaySound "initialize"
		Players = 1
		CanPlayTxt.text = Players
		For x = 1 to 4
			Score(x) = 0
			If B2SOn then controller.B2SSetScorePlayer x, Score(x)
			SReels(x).setvalue(0)
		Next
		For x = 1 to 9
			Grid(x) = 0
		Next
		ShutOffLights
		EVAL("OXLight" & OXO).state = 1
		newgame.enabled=true
	Else If  State = True and Players < MaxPlayers and BallinPlay = 1 Then
		Players = Players + 1
		CanPlayTxt.text = Players
		CreditTxt.text = Credit
		If B2SOn then
			Controller.B2SSetCredits Credit
			Controller.B2SSetCanplay 31, Players
		End If
		ShutOffLights
		EVAL("OXLight" & OXO).state = 1
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
	ShootAgainLight.state = 0
	GameState
    BIPText.text = "1"
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
		CanPlayTxt.text = 0
		RoundHS = Score(1)
		RoundHSPlayer = 1
		For i = 2 to MaxPlayers
			If Score(i) > RoundHS Then
				RoundHS = Score(i)
				RoundHSPlayer = i
			End If
		Next
		Players = 0
		If RoundHS > HiSc Then
			HiSc = RoundHS
			HSi = 1
			HSx = 1
			y = 1
			Initial(1) = 1
			For x = 2 to 3
				Initial(x) = 0
			Next
			UpdatePostIt
			InitialTimer1.enabled = 1
			EnableInitialEntry = True
		End If
		HsTxt.text = HiSc
		UpdatePostIt
		SaveHS
		If B2SOn Then
			Controller.B2SSetGameOver 35,1
			Controller.B2SSetballinplay 32, 0
			Controller.B2SSetPlayerUp 30, 0
			Controller.B2SSetcanplay 31, 0
		End If
		For each Light in GIlights:light.state = 0:Next
		CheckContinue.enabled = 0
	Else
		If BallInLane = False and BIP = 0 then
		BIPText.text = BallinPlay
		ReleaseBall
		PlaySound "kickerkick"
	End If
	CheckContinue.enabled = 0
  End If
End Sub

'***************Drain and Release Ball
Dim BIP : BIP = 0

Sub Drain_Hit()
	BIP = BIP - 1
	PlaySound SoundFXDOF("fx_drain",124,DOFPulse,DOFContactors)
	ClearAB
	For each Light in GreenBumperLights:light.state = 0:Next
	For each Light in BlueBumperLights:light.state = 0:Next
	SpecialGrid = False
	ShutOffLights
	If BIP = 0 Then ScoreBonus
End Sub

Sub ReleaseBall
	If BIP = 0 Then
		Drain.kick 62, 62
		If B2SOn Then DOF 125, DOFPulse
		If BallinPlay = 5 Then BonusLight.state = 1
		Playsound "metalhit_thin"
		ClearAB
		BIP = BIP + 1
		Launched = 0
		RotateGateShut
		If B2SOn Then Controller.B2SSetballinplay 32, BallinPlay
	End If
End Sub

'*******************Check for Bonus Score
Sub ScoreBonus
	If TicTacToeBonus = True Then
		For i = 1 to 9
			If Grid(i) > 0 Then
				LastSquare = LastSquare + 1
				SGrid(LastSquare) = 1
				Grid(i) = 0
			End If
		Next
		GridAdvance = False
		ScoreGrid.enabled = True
		Exit Sub
	Else
		ShootAgainLight.state = 0
		For i = 1 to 9
			If Grid(i) > 0 Then
				LastSquare = LastSquare + 1
				SGrid(LastSquare) = 1
				Grid(i) = 0
			End If
		Next
		If BallinPlay = 5 Then ScoreGrid.interval = 600
		GridAdvance = True
		ScoreGrid.enabled = 1
	End If
End Sub

'**************Advance Players
Sub AdvancePlayers
	GridAdvance = False
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
	ShootAgainLight.state = 0
	If Player = 1 Then BallinPlay = BallinPlay + 1
	If BallinPlay = 5 Then BonusLight.state = 1
	If BallinPlay > Balls Then
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
Sub CoinDelay_timer
	AddCredit
	coindelay.enabled = False
End Sub

Sub AddCredit
	Credit = Credit + 1
	If B2SOn Then DOF 127, DOFPulse
	If Credit > 25 then Credit = 25
	CreditTxt.text = Credit
	If B2SOn Then Controller.B2SSetCredits Credit
End Sub

'************Game State Check
Sub GameState
	If State = False Then
		For each light in GIlights:light.state=0: Next
		GamOv.text="Game Over"
		If B2SOn then Controller.B2SSetGameOver 35,1
		For i = 1 to 4
			EVAL ("Player" & i).intensityscale = 1
		Next
	Else
		For each Light in GIlights:Light.state=1: Next
		GamOv.text=""
		MatchTxt.text= ""
		TiltTxt.text=" "
		If B2SOn then
			Controller.B2SSetTilt 33,0
			Controller.B2SSetMatch 34,0
			Controller.B2SSetGameOver 35,0
		End If
	End If
End Sub

'*************Ball in Launch Lane
Sub BallHome_hit
	BallREnabled = 1
	If B2SOn then DOF 160, DOFOn
	ABLight1.state = 0
	ABLight2.state = 0
	RotateGateShut
End Sub

Sub BallHome_unhit
	If B2SOn Then DOF 160, DOFOff
End Sub

'************Check if Ball Out of Launch Lane
Sub BallsInPlay_hit
	If BallREnabled = 1 Then
		If AB = False Then
			ShootAgainLight.state = 0
			TicTacToeBonus = False
		End If
		ClearAB
		BallREnabled = 0
		BallInLane = False
	End If
	GateTimer.enabled = 1
End Sub

'***************Match
Sub Match
   y = int(rnd(1)*9)
    MatchNumber = y * 10
		If B2SOn Then
			If MatchNumber = 0 Then
				Controller.B2SSetMatch 100
			Else
				Controller.B2SSetMatch 34,MatchNumber
			End If
		End If
	MatchTxt.text = MatchNumber
	If MatchNumber = 0 then Matchtxt.text = "00"

	For i = 1 to Players
		MatchScoreTxt.text =  (Score(i) mod 100)
		If MatchNumber = (Score(i) mod 100) Then
			AddCredit
			Playsound SoundFXDOF("Knocker",128,DOFPulse,DOFKnocker)
			If B2SOn Then DOF 151, DOFPulse
	    End If
    Next
End Sub

'***************Shut Off Lights
Sub ShutOffLights
	For x = 1 to 2
		EVAL("EBLight" & x).state = 0
		EVAL("SpecialLight" & x).state = 0
		EVAL("ABLight" & x).state = 0
	Next
	BonusLight.state = 0
End Sub

'***************Bumpers
Sub Bumpers_hit(Index)
	If Tilt = False Then
		Select Case (Index)
			case 0: Playsound SoundFXDOF("FX_Bumper1",105,DOFPulse,DOFContactors)
					If TicTacToe = True Then
						AddScore 100
					Else
						AddScore 10
					End If
					OXOUpdate
			case 1: Playsound SoundFXDOF("FX_Bumper2",106,DOFPulse,DOFContactors)
					If AB = True Then
						AddScore 1000
					Else
						AddScore 100
					End If
			case 2: Playsound SoundFXDOF("FX_Bumper3",107,DOFPulse,DOFContactors)
					If TicTacToe = True Then
						AddScore 100
					Else
						AddScore 10
					End If
					OXOUpdate
		End Select
	End If
End Sub

'***************Saucers
Sub RightSaucerKicker_hit
	If B2SOn Then DOF 123, DOFPulse
	If TicTacToe = True and TicTacToeBonus <1 Then
		TicTacToeBonus = True
		ShootAgainLight.state = 1
		AddScore 5000
	Else
		AddScore 500
	End If
	me.timerenabled = 1
End Sub

Sub LeftSaucerKicker_hit
	If B2SOn Then DOF 122, DOFPulse
	If TicTacToe = True and TicTacToeBonus <1 Then
		TicTacToeBonus = True
		ShootAgainLight.state = 1
		AddScore 5000
	Else
		AddScore 500
	End If
	me.timerenabled = 1
End Sub

'***************Kickers
Sub LeftSaucerKicker_Timer
	LeftSaucerKicker.Kick 92 + (3 * Rnd()), 6
	Playsound SoundFXDOF("HoleKick",129,DOFPulse,DOFContactors)
	me.timerenabled = 0
End Sub

Sub RightSaucerKicker_Timer
	RightSaucerKicker.Kick 268 + ( 3 * Rnd()), 6
	Playsound SoundFXDOF("HoleKick",129,DOFPulse,DOFContactors)
	me.timerenabled = 0
End Sub

'***************Slings
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

'****************Triggers
Sub LeftUpperO_hit
	If Tilt = False Then AddScore 100
	If B2SOn Then DOF 108, DOFPulse
	If Grid(1) = 0 then Grid(1) = 1
	CheckTicTacToe
End Sub

Sub LeftUpperX_hit
	If Tilt = False Then AddScore 100
	If B2SOn Then DOF 109, DOFPulse
	If Grid(1) = 0 Then Grid(1) = 2
	CheckTicTacToe
End Sub

Sub MidUpperO_hit
	If Tilt = False Then AddScore 100
	If B2SOn Then DOF 110, DOFPulse
	If Grid(2) = 0 Then Grid(2) = 1
	CheckTicTacToe
End Sub

Sub MidUpperX_hit
	If Tilt = False Then AddScore 100
	If B2SOn Then DOF 110, DOFPulse
	If Grid(2) = 0 then Grid(2) = 2
	CheckTicTacToe
End Sub

Sub RightUpperO_hit
	If Tilt = False Then AddScore 100
	If B2SOn Then DOF 111, DOFPulse
	If Grid(3) = 0 Then Grid(3) = 1
	CheckTicTacToe
End Sub

Sub RightUpperX_hit
	If Tilt = False Then AddScore 100
	If B2SOn Then DOF 112, DOFPulse
	If Grid(3) = 0 Then Grid(3) = 2
	CheckTicTacToe
End Sub

Sub LeftMidOX_hit
	If Tilt = False Then AddScore 1000
	If B2SOn Then DOF 113, DOFPulse
	If Grid(4) = 0 Then Grid(4) = OXO
	CheckTicTacToe
	If SpecialGrid = True then Special
End Sub

Sub CenterOX_hit
	If Tilt = False Then AddScore 100
	If B2SOn Then DOF 121, DOFPulse
	If Grid(5) = 0 Then Grid(5) = OXO
	CheckTicTacToe
End Sub

Sub RightMidOX_Hit
	If Tilt = False Then AddScore 1000
	If B2SOn Then DOF 114, DOFPulse
	If Grid(6) = 0 Then Grid(6) = OXO
	CheckTicTacToe
	If SpecialGrid = True then Special
End Sub

Sub LeftLowOX_hit
	If Tilt = False Then AddScore 100
	If B2SOn Then DOF 117, DOFPulse
	If Grid(7) = 0 Then Grid(7) = OXO
	CheckTicTacToe
End Sub

Sub MidLowOX_hit
	Button = 0
	me.timerenabled = 1
	If Tilt = False Then
		AddScore 100
		If Grid(8) = 0 Then Grid(8) = OXO
		CheckTicTacToe
	End If
End Sub

Sub MidLowOX_Timer
	Button = Button +1
	Select Case Button
		Case 1: RedRollOverButton.transz=1
		Case 2: RedRollOverButton.transz=0
		Case 3: RedRollOverButton.transz=1
		Case 4: RedRollOverButton.transz=3
		me.timerenabled = 0
	End Select
End Sub

Sub RightLowOX_hit
	If Tilt = False Then AddScore 100
	If B2SOn Then DOF 118, DOFPulse
	If Grid(9) = 0 Then Grid(9) = OXO
	CheckTicTacToe
End Sub

Sub ATarget_hit
	If Tilt = False Then
		If B2SOn Then DOF 119, DOFPulse
		AddScore 100
		TargetA = 1
		ABLight1.state = 1
		CheckAB
	End If
End Sub

Sub BTarget_hit
	If Tilt = False Then
		If B2SOn Then DOF 120, DOFPulse
		AddScore 100
		TargetB = 1
		ABLight2.state = 1
		CheckAB
	End If
End Sub

sub LeftOutLane_hit
	If Tilt = False Then
		AddScore 1000
		CheckTicTacToe
		If B2SOn Then DOF 115, DOFPulse
	End If
end sub

sub RightOutLane_hit
	If Tilt = False Then
		AddScore 1000
		CheckTicTacToe
		If B2SOn Then DOF 116, DOFPulse
	End If
end sub

Sub BallLaunchedTrigger_hit
	Launched = Launched + 1
	If CLng(Launched) Mod 2 > 0 Then
		If B2SOn Then DOF 161 ,DOFPulse
	End If
End Sub

'**************Special
Sub Special
		AddCredit
		Playsound SoundFXDOF("Knocker",128,DOFPulse,DOFKnocker)
		If B2SOn Then DOF 151, DOFPulse
End Sub

'**************Update Grid
Sub UpdateGrid
	For x = 1 to 9
		If Grid(x) = 1 Then
			EVAL("GL_O_a" & x).state = 1
			EVAL("GL_O_b" & x).state = 1
		End If
		If Grid(x) = 2 Then	EVAL("GL_X_" & x).state = 1
	Next
End Sub

'***************Shooter Lane Gate Animation
Sub GateTimer_Timer
	RelGateHit = RelGateHit + 1
	Select Case RelGateHit
		Case 1:	Pgate.Rotz=45
		Case 2: Pgate.Rotz=60
		Case 3: Pgate.Rotz=45
		Case 4: Pgate.Rotz=25
		me.enabled = 0
		RelGateHit = 0
	End Select
End Sub

'***************Rotate AB Gate
Sub RotateGateOpen
	ABDiverterGate.ObjRotZ = 45
	ABGateWallOpen.isDropped = False
	ABGateWallClosed.isdropped = True
End Sub

Sub RotateGateShut
	ABDiverterGate.ObjRotZ = 90
	ABGateWallOpen.isDropped = True
	ABGateWallClosed.isdropped = False
End Sub

'***************Scoring Routine
Sub AddScore(points)
  If Tilt = False Then
		If Points = 10 Then Playsound"Reels" : Playsound SoundFXDOF("Bell10",153,DOFPulse,DOFChimes) : TotalUp 10
		If Points = 100 Then Playsound"Reels" : Playsound SoundFXDOF("Bell100",154,DOFPulse,DOFChimes) : TotalUp 100
		If Points = 500 Then Playsound"Reels5" : BellRing = 5 : BellTimer500.enabled = 1
		If Points = 1000 Then Playsound"Reels" : Playsound SoundFXDOF("Bell1000",155,DOFPulse,DOFChimes) : TotalUp 1000
		If Points = 3000 Then Playsound"Reels3" : BellRing = 3 : BellTimer3000.enabled = 1   'Playsound SoundFXDOF("Bell3000Quick",155,DOFPulse,DOFChimes)
		If Points = 5000 Then Playsound"Reels5" : BellRing = 5 : BellTimer5000.enabled = 1  'Playsound SoundFXDOF("Bell5000",155,DOFPulse,DOFChimes)
	End If
End Sub

Sub TotalUp(Points)
	Score(Player) = Score(Player) + Points
	SReels(Player).addvalue(Points)
	If B2SOn Then Controller.B2SSetScorePlayer Player, Score(Player)
    If Score(Player) => Replay1 and Rep(Player) = 0 Then
		AddCredit
		Rep(Player) = 1
		Playsound SoundFXDOF("Knocker",128,DOFPulse,DOFKnocker)
		If B2SOn Then DOF 151, DOFPulse
	End If
    If Score(Player) => Replay2 and Rep(Player) = 1 Then
		AddCredit
		rep(player)=2
		Playsound SoundFXDOF("Knocker",128,DOFPulse,DOFKnocker)
		If B2SOn Then DOF 151, DOFPulse
    End If
    If Score(Player) => Replay3 and Rep(Player) = 2 Then
		AddCredit
		Rep(Player)=3
		Playsound SoundFXDOF("Knocker",128,DOFPulse,DOFKnocker)
		If B2SOn Then DOF 151, DOFPulse
    End If
End Sub

'*************** Bell Timers
Sub BellTimer500_Timer
	Playsound SoundFXDOF("Bell100",154,DOFPulse,DOFChimes)
	BellRing = BellRing - 1
	TotalUp 100
	If BellRing < 1 Then
		BellTimer500.enabled = 0
	End If
End Sub

Sub BellTimer3000_Timer
	Playsound SoundFXDOF("Bell1000",155,DOFPulse,DOFChimes)
	BellRing = BellRing - 1
	TotalUp 1000
	If BellRing < 1 Then
		BellTimer3000.enabled = 0
	End If
End Sub

Sub BellTimer5000_Timer
	Playsound SoundFXDOF("Bell1000",155,DOFPulse,DOFChimes)
	BellRing = BellRing - 1
	TotalUp 1000
	If BellRing < 1 Then
		BellTimer5000.enabled = 0
	End If
End Sub

'*************** Score Tic Tac Toe Grid
Sub ScoreGrid_Timer
	Square = Square +1
	If SGrid(Square) > 0 Then
		If BallinPlay < 5 Or TicTacToeBonus = True Then AddScore 1000
		If BallinPlay = 5 and TicTacToeBonus = False Then AddScore 3000
		SGrid(Square) = 0
	End If
	If Square = LastSquare Then
		Square = 0
		LastSquare = 0
		ScoreGrid.enabled = 0
		ClearGrid
		UpdateGrid
		GridChoice
		If TicTacToeBonus = True Then ReleaseBall
	End If
End Sub

'*************** Grid Choice
Sub GridChoice
	If GridAdvance = True Then
		AdvancePlayers
	Else
		ReleaseBall
	End If
End Sub

'*************** OXO Update
Sub OXOUpdate
	EVAL("OXLight" & OXO).state = 0
	OXO = OXO + 1
	If OXO = 3 Then OXO = 1
	EVAL("OXLight" & OXO).state = 1
	UpdateGrid
End Sub

'*************** Clear Grid
Sub ClearGrid
	For i = 1 to 9
		Grid(i) = 0
		EVAL("GL_O_a" & i).state = 0
		EVAL("GL_O_b" & i).state = 0
		EVAL("GL_X_" & i).state = 0
	Next
	TicTacToe = False
End Sub

'***************Check Tic Tac Toe
Sub CheckTicTacToe
	For x = 1 to 7 step 3
		If Grid(x) > 0 Then
			If Grid(x) = Grid (x + 1) and Grid(x) = Grid (x + 2) Then TicTacToe = True
		End If
	Next

	For x = 1 to 3
		If Grid(x) > 0 Then
			If Grid(x) = Grid(x+3) and Grid(x) = Grid(x+6) Then TicTacToe = True
		End If
	Next

	If Grid(1) > 0 Then
		If Grid(1) = Grid(5) and Grid(1) = Grid(9) Then TicTacToe = True
	End If

	If Grid(3) > 0 Then
		If Grid(3) = Grid(5) and Grid(3) = Grid(7) Then TicTacToe = True
	End If
	For x = 1 to 9
		If Grid(x) > 0 then GridFill = GridFill + 1
	Next
	If GridFill = 9 then SpecialGrid = True: SpecialLight1.state = 1: SpecialLight2.state = 1
	If TicTacToe = True Then
		EBLight1.state = 1: EBLight2.state = 1
		For each Light in BlueBumperLights:light.state = 1:Next
	End If
	UpdateGrid
	GridFill = 0
End Sub

'***************Clear AB
Sub ClearAB
	TargetA = 0
	TargetB = 0
	AB = 0
End Sub

'***************Check AB
Sub CheckAB
	If TargetA = 1 and TargetB = 1 then AB = True
	If AB = True Then
		RotateGateOpen
		For each Light in GreenBumperLights:light.state = 1:Next
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
Sub a_Pins_Hit (idx)
	PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub a_Targets_Hit (idx)
	PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub a_Metals_Thin_Hit (idx)
	PlaySound "metalhit_thin", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub a_Metals_Medium_Hit (idx)
	PlaySound "metalhit_medium", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub a_Metals2_Hit (idx)
	PlaySound "metalhit2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub a_Gates_Hit (idx)
	PlaySound "gate4", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub a_Spinner_Spin
	PlaySound "fx_spinner",0,.25,0,0.25
End Sub

Sub a_Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RubberWheel_hit
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End sub

Sub a_Posts_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 16 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 16 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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
		Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub

'************Enter Initials
Sub EnterIntitals(keycode)
	If KeyCode = LeftFlipperKey Then
		HSx = HSx - 1
		if HSx < 0 Then HSx = 26
		If HSi < 4 Then EVAL("Initial" & HSi).image = HSiArray(HSx)
	End If
	If keycode = RightFlipperKey Then
		HSx = HSx + 1
		If HSx > 26 Then HSx = 0
		If HSi < 4 Then EVAL("Initial"& HSi).image = HSiArray(HSx)
	End If
		If keycode = StartGameKey Then
			If HSi < 3 Then
				EVAL("Initial" & HSi).image = HSiArray(HSx)
				Initial(HSi) = HSx
				EVAL("InitialTimer" & HSi).enabled = 0
				EVAL("Initial" & HSi).visible = 1
				Initial(HSi + 1) = HSx
				EVAL("Initial" & HSi +1).image = HSiArray(HSx)
				y = 1
				EVAL("InitialTimer" & HSi + 1).enabled = 1
				HSi = HSi + 1
			Else
				Initial3.visible = 1
				InitialTimer3.enabled = 0
				Initial(3) = HSx
				InitialEntry.enabled = 1
				HSi = HSi + 1
			End If
		End If
End Sub

Sub InitialEntry_timer
	SaveHS
	HSi = HSi + 1
	EnableInitialEntry = False
	InitialEntry.enabled = 0
	Players = 0
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
    SaveValue  "OXO", "Credit", Credit
    SaveValue  "OXO", "HIScore", HiSc
    SaveValue  "OXO", "Match", MatchNumber
    SaveValue  "OXO", "Initial1", Initial(1)
    SaveValue  "OXO", "Initial2", Initial(2)
    SaveValue  "OXO", "Initial3", Initial(3)
    SaveValue  "OXO", "Score4", Score(4)
	SaveValue  "OXO", "Balls", Balls
End Sub

'*************Load Scores
Sub LoadHighScore
    dim temp
	temp = LoadValue("OXO", "credit")
    If (temp <> "") then Credit = CDbl(temp)
    temp = LoadValue("OXO", "HiScore")
    If (temp <> "") then HiSc = CDbl(temp)
    temp = LoadValue("OXO", "match")
    If (temp <> "") then MatchNumber = CDbl(temp)
    temp = LoadValue("OXO", "Initial1")
    If (temp <> "") then Initial(1) = CDbl(temp)
    temp = LoadValue("OXO", "Initial2")
    If (temp <> "") then Initial(2) = CDbl(temp)
    temp = LoadValue("OXO", "Initial3")
    If (temp <> "") then Initial(3) = CDbl(temp)
    temp = LoadValue("OXO", "score4")
    If (temp <> "") then score(4) = CDbl(temp)
    temp = LoadValue("OXO", "balls")
    If (temp <> "") then Balls = CDbl(temp)
End Sub

'***************Post It Note Update
Sub UpdatePostIt
	ScoreMil = Int(HiSc/1000000)
	Score100K = Int( (HiSc - (ScoreMil*1000000) ) / 100000)
	Score10K = Int( (HiSc - (ScoreMil*1000000) - (Score100K*100000) ) / 10000)
	ScoreK = Int( (HiSc - (ScoreMil*1000000) - (Score100K*100000) - (Score10K*10000) ) / 1000)
	Score100 = Int( (HiSc - (ScoreMil*1000000) - (Score100K*100000) - (Score10K*10000) - (ScoreK*1000) ) / 100)
	Score10 = Int( (HiSc - (ScoreMil*1000000) - (Score100K*100000) - (Score10K*10000) - (ScoreK*1000) - (Score100*100) ) / 10)
	ScoreUnit = (HiSc - (ScoreMil*1000000) - (Score100K*100000) - (Score10K*10000) - (ScoreK*1000) - (Score100*100) - (Score10*10) )

	Pscore6.image = HSArray(ScoreMil):If HiSc < 1000000 Then PScore6.image = HSArray(10)
	Pscore5.image = HSArray(Score100K):If HiSc < 100000 Then PScore5.image = HSArray(10)
	PScore4.image = HSArray(Score10K):If HiSc < 10000 Then PScore4.image = HSArray(10)
	PScore3.image = HSArray(ScoreK):If HiSc < 1000 Then PScore3.image = HSArray(10)
	PScore2.image = HSArray(Score100):If HiSc < 100 Then PScore2.image = HSArray(10)
	PScore1.image = HSArray(Score10):If HiSc < 10 Then PScore1.image = HSArray(10)
	PScore0.image = HSArray(ScoreUnit):If HiSc < 1 Then PScore0.image = HSArray(10)
	If HiSc < 1000 Then
		PComma.image = HSArray(10)
	Else
		PComma.image = HSArray(11)
	End If
	If HiSc < 1000000 Then
		PComma1.image = HSArray(10)
	Else
		PComma1.image = HSArray(11)
	End If
	If HiSc < 1000000 Then Shift = 1:PComma.transx = -10
	If HiSc < 100000 Then Shift = 2:PComma.transx = -20
	If HiSc < 10000 Then Shift = 3:PComma.transx = -30
	For x = 0 to 6
		EVAL("Pscore" & x).transx = (-10 * Shift)
	Next
	Initial1.image = HSiArray(Initial(1))
	Initial2.image = HSiArray(Initial(2))
	Initial3.image = HSiArray(Initial(3))
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


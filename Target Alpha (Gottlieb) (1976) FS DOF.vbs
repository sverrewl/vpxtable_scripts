'	-------------------------------------------------
'	Based on Gottlieb EM 4 player VPX table blank with options menu by BorgDog, 2016
'	-------------------------------------------------
'
'	Typical Layer usage
'		1 - most stuff
'		2 - triggers
'		3 - under apron walls and ramps
'		4 - options menu
'		5 - hi score sticky
'		6 - GI lighting
'		7 - plastics
'		8 - insert and bumper lighting
'
'	Basic DOF config - matched to loserman76 table + Chimes
'		101 Left Flipper, 102 Right Flipper,
'		103 Left sling, 104 right sling,
'		106 Bumper1, 105 bumper2,
'		107 Drain, 108 Ball Release, 109 ShooterLane
'		127 left round targets, 113-121 triggers
'		123 top left 3 drops, 124 next 4 drops, 125 top right 3 DropSol_On
'		126 lower drop targets,
'		128-131 reset Drops
'		110 Shooter Lane/launch ball, 111 credit light, 112 knocker
'		141 Chime1-10s, 142 Chime2-100s, 143 Chime3-1000s
'
'
'	BUILT IN VPX BY BORGDOG, SCRIPT BASED ON LOSERMAN76 VP9 TABLE
'	-------------------------------------------------
'*
'*        Gottlieb's Target Alpha (1976)
'*        Table build/scripted by Loserman76
'*        Playfield/plastics redrawn by GNance
'*        Original Images by Popotte
'*        DOF Coding by Arngrim
'*        http://www.ipdb.org/machine.cgi?id=2500
'*

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' No special SSF tweaks yet.


Option Explicit
Randomize

Const cGameName = "target_alpha_1976"

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

'* this value adjusts score motor behavior - 0 allows you to continue scoring while the score motor is running - 1 sets score motor to behave more like a real EM
Const ScoreMotorAdjustment=1

'* this is a debug setting to use an older scoring routine vs a newer score routine - don't change this value
Const ScoreAdditionAdjustment=1

'* this controls whether you hear bells (0) or chimes (1) when scoring
Const ChimesOn=1


dim ScoreChecker, CheckAllScores, sortscores(4), sortplayers(4)
Dim TextStr,TextStr2
Dim i, obj, bgpos, kgpos
Dim dooralreadyopen
Dim kgdooralreadyopen
Dim TargetSpecialLit
Dim Points210counter
Dim Points500counter
Dim Points1000counter
Dim Points2000counter
Dim OperatorMenu, options, freeplay
Dim BallsPerGame
Dim InProgress
Dim BallInPlay
Dim CreditsPerCoin
Dim Score100K(4)
Dim Score(4)
Dim ScoreDisplay(4)
Dim HighScorePaid(4)
Dim HighScore
Dim HighScoreReward
Dim BonusMultiplier
Dim Credits
Dim Match
Dim Replay1
Dim Replay2
Dim Replay3
Dim Replay1Paid(4)
Dim Replay2Paid(4)
Dim Replay3Paid(4)
Dim TableTilted
Dim TiltCount
Dim EnableSpecialEB

Dim Replay1Table(15)
Dim Replay2Table(15)
Dim Replay3Table(15)
Dim ReplayTableSet

Dim UpperTargetValue
Dim LowerTargetValue
Dim ExtraBall

Dim BonusBooster
Dim BonusBoosterCounter
Dim BonusCounter
Dim HoleCounter

Dim Ones
Dim Tens
Dim Hundreds
Dim Thousands

Dim Player
Dim Players

Dim rst
Dim bonuscountdown
Dim TempMultiCounter
dim TempPlayerup
dim RotatorTemp

Dim bump1
Dim bump2
Dim bump3

Dim LastChime10
Dim LastChime100
Dim LastChime1000

Dim Score10
Dim Score100

Dim MotorRunning

Dim SpecialEBOption
Dim ReplayLevel
Dim ReplayTableMax

Dim X, xx
Dim LStep, LStep2, RStep, L2Step, R2Step
Dim objekt, light, ballinshooter

Sub TargetAlpha_Init()
	LoadEM
	HideOptions
	Replay1Table(1)=72000
	Replay2Table(1)=84000
	Replay1Table(2)=79000
	Replay2Table(2)=91000
	Replay1Table(3)=85000
	Replay2Table(3)=97000

	OperatorMenu=0
	BallsPerGame=3
	UpperTargetValue=0
	LowerTargetValue=0
	HighScore=0
	MotorRunning=0
	SpecialEBOption=0
	HighScoreReward=3
	ReplayLevel=1
	Credits=0
	freeplay=0
	loadhs
	if HighScore=0 then HighScore=20000
	if HSA1="" then HSA1=25
	if HSA2="" then HSA2=25
	if HSA3="" then HSA3=25
	UpdatePostIt
	TableTilted=false

	Match=int(Rnd*10)*10
	Matchtxt.text = match
	Credittxt.SetValue(Credits)

	CanPlayReel.SetValue(0)
	BIPReel.SetValue(0)

	For each obj in PlayerHUDScores
		obj.state=0
	next

	For i = 1 to 4
		EVAL("EMReel"&i).setvalue score(i)
	next

	Replay1=Replay1Table(ReplayLevel)
	Replay2=Replay2Table(ReplayLevel)
	ReplayCard.image = "ReplayCard"&ReplayLevel
	OptionBalls.image="OptionsBalls"&BallsPerGame
	OptionReplays.image="OptionsReplays"&ReplayLevel
	OptionFreeplay.image="OptionsFreeplay"&freeplay
	BonusCounter=0
	HoleCounter=0
	EnableSpecialEB=0
	if ballspergame=3 then
		InstructCard.image="InstCard3balls"
	  else
		InstructCard.image="InstCard5balls"
	end if
	BumperLight1.state=0
	BumperLight2.state=0
	ExtraBall=false
	TargetSpecialLit = 0
	Points210counter=0
	Points500counter=0
	Points1000counter=0
	Points2000counter=0
	BonusBooster=3
	BonusBoosterCounter=0
	Players=0
	RotatorTemp=1
	InProgress=false

	gamov.text="GAME OVER"

	tilttxt.text="TILT"

 	If B2SOn Then
		For each objekt in Backdropstuff: objekt.visible=false: next
	End If

	startGame.enabled=true

	bump1=1
	bump2=1
	bump3=1
	InitPauser5.enabled=true
	Drain.CreateBall
End Sub

sub startGame_timer
	playsound "poweron"
	lightdelay.enabled=true
	If B2SOn Then
		if Match=0 then
			Controller.B2SSetMatch 100
		else
			Controller.B2SSetMatch Match
		end if
		for i=1 to 4
			Controller.B2SSetScorePlayer i, 0
			Controller.B2SSetScoreRollover i+24, 0
		next

		Controller.B2SSetShootAgain 0
		Controller.B2SSetTilt 0
		Controller.B2SSetCredits Credits
		Controller.B2SSetGameOver 1
	End If
	me.enabled=false
end sub

sub lightdelay_timer
	If Credits > 0 Then DOF 111, 1
	For each light in GIlights:light.state=1:Next
	For each light in BumperLights:light.state=1:Next
	If B2SOn then Controller.B2SSetData 99,1
	me.enabled=false
end sub

Sub OperatorMenuTimer_Timer
	OperatorMenu=1
	Displayoptions
	Options=1
End Sub

Sub DisplayOptions
	OptionsBack.visible = true
	Option1.visible = True
	OptionBalls.visible = True
    OptionReplays.visible = True
	OptionFreeplay.visible = True
End Sub

Sub HideOptions
	for each objekt In OptionMenu
		objekt.visible = false
	next
End Sub

Sub TargetAlpha_exit()
	savehs
	turnoff
	If B2SOn Then Controller.stop
end sub

Sub TargetAlpha_KeyDown(ByVal keycode)

	If keycode = PlungerKey Then
		Plunger.PullBack
		PlaySound "plungerpull",0,1,0.25,0.25
	End If

	if keycode = LeftFlipperKey and InProgress = false  and OperatorMenu=0 then
		OperatorMenuTimer.Enabled = true
	end if

	If keycode=LeftFlipperKey and InProgress= false and OperatorMenu=1 then
		Options=Options+1
		If Options=5 then Options=1
		playsound "target"
		Select Case (Options)
			Case 1:
				Option1.visible=true
				Option4.visible=False
			Case 2:
				Option2.visible=true
				Option1.visible=False
			Case 3:
				Option3.visible=true
				Option2.visible=False
			Case 4:
				Option4.visible=true
				Option3.visible=False
		End Select
	end if

	If keycode=RightFlipperKey and InProgress = false and OperatorMenu=1 then
	  PlaySound "metalhit2"
	  Select Case (Options)
		Case 1:
			if BallsPerGame=3 then
				BallsPerGame=5
				InstructCard.image="InstCard5balls"
			  else
				BallsPerGame=3
				InstructCard.image="InstCard3balls"
			end if
			OptionBalls.image = "OptionsBalls"&BallsPerGame
		Case 2:
			if freeplay=0 Then
				freeplay=1
			  Else
				freeplay=0
			end if
			OptionFreeplay.image="OptionsFreeplay"&freeplay
		Case 3:
			ReplayLevel=ReplayLevel+1
			if ReplayLevel>3 then
				ReplayLevel=1
			end if
			Replay1=Replay1Table(ReplayLevel)
			Replay2=Replay2Table(ReplayLevel)
			OptionReplays.image = "OptionsReplays"&ReplayLevel
			replaycard.image = "ReplayCard"&ReplayLevel
		Case 4:
			OperatorMenu=0
			savehs
			HideOptions
	  End Select
	End If

	If keycode = LeftFlipperKey and InProgress=true and TableTilted=false Then
		LeftFlipper.RotateToEnd
		LeftFlipper2.RotateToEnd
		PlaySound SoundFXDOF("flipperup",101,DOFOn,DOFFlippers), 0, .67, -0.05, 0.05
		PlaySound "Buzz",-1,.05,-0.05, 0.05
		DOF 101,1
	End If

	If keycode = RightFlipperKey  and InProgress=true and TableTilted=false Then
		RightFlipper.RotateToEnd
		RightFlipper2.RotateToEnd
		PlaySound SoundFXDOF("flipperup",102,DOFOn,DOFFlippers), 0, .67, 0.05, 0.05
		PlaySound "Buzz1",-1,.05,0.05,0.05
		DOF 102,1
	End If

	If keycode = LeftTiltKey Then
		Nudge 90, 2
		TiltIt
	End If

	If keycode = RightTiltKey Then
		Nudge 270, 2
		TiltIt
	End If

	If keycode = CenterTiltKey Then
		Nudge 0, 2
		TiltIt
	End If

	If keycode = MechanicalTilt Then
		gametilted
	End If

	If keycode = AddCreditKey then
		If B2SOn Then
			'Controller.B2SSetScore 6,HighScore
		End If
		playsound "coinin"
		AddSpecial2
	end if

   if keycode = StartGameKey and (Credits>0 or freeplay=1) and InProgress=true and Players>0 and Players<4 and BallInPlay<2 And Not HSEnterMode=true  then
		if freeplay=0 then
			Credits=Credits-1
			If Credits = 0 Then DOF 111, 0
			Credittxt.SetValue(Credits)
		end if
		Players=Players+1
		CanPlayReel.SetValue(Players)
		playsound "click"
		If B2SOn Then
			Controller.B2SSetCanPlay Players
			If Players=2 Then
				Controller.B2SSetScoreRolloverPlayer2 0
			End If
			If Players=3 Then
				Controller.B2SSetScoreRolloverPlayer3 0
			End If
			If Players=4 Then
				Controller.B2SSetScoreRolloverPlayer4 0
			End If
			Controller.B2SSetCredits Credits
		End If
    end if

	if keycode=StartGameKey and (Credits>0 or freeplay=1) and InProgress=false and Players=0 And Not HSEnterMode=true  and operatormenu=0 then
		if freeplay=0 then
			Credits=Credits-1
			If Credits = 0 Then DOF 111, 0
			Credittxt.SetValue(Credits)
		end if
		Players=1
		PlasticsOn
		CanPlayReel.SetValue(Players)
		matchtxt.text = ""
		gamov.text=""
		Player=1
		playsound "StartUpSequence"
		TempPlayerUp=Player
		PlayerUpRotator.enabled=true
		rst=0
		BallInPlay=1
		InProgress=true
		resettimer.enabled=true
		BonusMultiplier=1
		If B2SOn Then
			Controller.B2SSetTilt 0
			Controller.B2SSetGameOver 0
			Controller.B2SSetMatch 0
			Controller.B2SSetCredits Credits
			Controller.B2SSetCanPlay 1
			Controller.B2SSetData 81,0
			Controller.B2SSetData 82,0
			Controller.B2SSetData 83,0
			Controller.B2SSetData 84,0
			Controller.B2SSetPlayerUp 1
			Controller.B2SSetData 81,1
			Controller.B2SSetBallInPlay BallInPlay
			Controller.B2SSetScoreRolloverPlayer1 0
		End If
		For each obj in PlayerScores
			obj.ResetToZero
		next
		For each obj in PlayerHuds
			obj.state = 0
		next
		For each obj in PlayerHUDScores
			obj.state=0
		next
		PlayerHuds(Player-1).state =1
		PlayerHUDScores(Player-1).state=1
	end if

    If HSEnterMode Then HighScoreProcessKey(keycode)

End Sub

Sub TargetAlpha_KeyUp(ByVal keycode)

	If keycode = PlungerKey Then
		if ballinshooter=1 then
			playsound "plungerreleaseball",0,1,0.1,.25,0.25
		  else
			playsound "plungerreleasefree",0,1,0.1,.25,0.25
		end if
		Plunger.Fire
	End If

	if keycode = LeftFlipperKey then
		OperatorMenuTimer.Enabled = false
	end if

	If keycode = LeftFlipperKey and InProgress=true and TableTilted=false Then
		LeftFlipper.RotateToStart
		LeftFlipper2.RotateToStart
		PlaySound SoundFXDOF("flipperdown",101,DOFOff,DOFContactors), 0, 1, -0.05, 0.05
		StopSound "Buzz"
		DOF 101,0
	End If

	If keycode = RightFlipperKey and InProgress=true and TableTilted=false Then
		RightFlipper.RotateToStart
        RightFlipper2.RotateToStart
		PlaySound SoundFXDOF("flipperdown",102,DOFOff,DOFContactors), 0, 1, 0.05, 0.05
		StopSound "Buzz1"
		DOF 102,0
	End If

End Sub

Sub Drain_Hit()
	PlaySound "fx_drain"
	DOF 107,2
	Pause4Bonustimer.enabled=1
End Sub

Sub Pause4Bonustimer_timer
	Pause4Bonustimer.enabled=0
	AddBonus

End Sub

'***********************
'     Flipper Logos
'***********************

Sub UpdateFlipperLogos_Timer
	LFlip.RotY = LeftFlipper.CurrentAngle
	RFlip.RotY = RightFlipper.CurrentAngle
	LFlip1.RotY = LeftFlipper2.CurrentAngle-90
	RFlip1.RotY = RightFlipper2.CurrentAngle-90
	Pgate.rotz=(Gate3.currentangle*.75)+25
End Sub


Sub ShooterLane_Hit()
	ballinshooter=1
	DOF 109, 1
End Sub

Sub ShooterLane_Unhit()
	ballinshooter=0
	DOF 109, 0
	DOF 110, 2
End Sub

'***********************
' slingshots
'***********************

 Sub RightSlingShot_Slingshot
	playsound SoundFXDOF("right_slingshot",105,DOFPulse,DOFContactors), 0, 1, 0.05, 0.05
	DOF 106,DOFPulse
	addscore 10
    RSling.Visible = 0
    RSling1.Visible = 1
	slingR.objroty = -15
    RStep = 1
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:slingR.objroty = -7
        Case 4:	slingR.objroty = 0:RSLing2.Visible = 0:RSLing.Visible = 1:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
	playsound SoundFXDOF("left_slingshot",103,DOFPulse,DOFContactors), 0, 1, -0.05, 0.05
	DOF 104,DOFPulse
	addscore 10
    LSling.Visible = 0
    LSling1.Visible = 1
	slingL.objroty = 15
    LStep = 1
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:slingL.objroty = 7
        Case 4:slingL.objroty = 0:LSLing2.Visible = 0:LSLing.Visible = 1:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub


'***********************************

Sub DingwallA_Hit()
	If TableTilted=false then
		AddScore(10)
	End if
	SlingA.visible=0
	SlingA1.visible=1
	dingwalla.uservalue=1
	Me.timerenabled=1
End Sub

sub dingwalla_timer
	select case dingwalla.uservalue
		Case 1: SlingA1.visible=0: SlingA.visible=1
		case 2:	SlingA.visible=0: SlingA2.visible=1
		Case 3: SlingA2.visible=0: SlingA.visible=1: Me.timerenabled=0
	end Select
	dingwalla.uservalue=dingwalla.uservalue+1
end sub

Sub DingwallB_Hit()
	If TableTilted=false then
		AddScore(10)
	End if
	Slingb.visible=0
	Slingb1.visible=1
	dingwallb.uservalue=1
	Me.timerenabled=1
End Sub

sub dingwallb_timer
	select case DingwallB.uservalue
		Case 1: Slingb1.visible=0: SlingB.visible=1
		case 2:	SlingB.visible=0: Slingb2.visible=1
		Case 3: Slingb2.visible=0: SlingB.visible=1: Me.timerenabled=0
	end Select
	dingwallb.uservalue=DingwallB.uservalue+1
end sub

Sub DingwallC_Hit()
	If TableTilted=false then
		AddScore(10)
	End if
	Slingc.visible=0
	Slingc1.visible=1
	dingwallc.uservalue=1
	Me.timerenabled=1
End Sub

sub dingwallc_timer
	select case Dingwallc.uservalue
		Case 1: Slingc1.visible=0: Slingc.visible=1
		case 2:	Slingc.visible=0: Slingc2.visible=1
		Case 3: Slingc2.visible=0: Slingc.visible=1: Me.timerenabled=0
	end Select
	dingwallc.uservalue=Dingwallc.uservalue+1
end sub

Sub DingwallD_Hit()
	If TableTilted=false then
		AddScore(10)
	End if
	Slingd.visible=0
	Slingd1.visible=1
	dingwalld.uservalue=1
	Me.timerenabled=1
End Sub

sub dingwalld_timer
	select case Dingwalld.uservalue
		Case 1: Slingd1.visible=0: Slingd.visible=1
		case 2:	Slingd.visible=0: Slingd2.visible=1
		Case 3: Slingd2.visible=0: Slingd.visible=1: Me.timerenabled=0
	end Select
	dingwalld.uservalue=Dingwalld.uservalue+1
end sub

Sub DingwallE_Hit()
	If TableTilted=false then
		AddScore(10)
	End if
	Slinge.visible=0
	Slinge1.visible=1
	dingwalle.uservalue=1
	Me.timerenabled=1
End Sub

sub dingwalle_timer
	select case Dingwalle.uservalue
		Case 1: Slinge1.visible=0: Slinge.visible=1
		case 2:	Slinge.visible=0: Slinge2.visible=1
		Case 3: Slinge2.visible=0: Slinge.visible=1: Me.timerenabled=0
	end Select
	dingwalle.uservalue=Dingwalle.uservalue+1
end sub

Sub DingwallF_Hit()
	If TableTilted=false then
		AddScore(10)
	End if
	Slingf.visible=0
	Slingf1.visible=1
	dingwallf.uservalue=1
	Me.timerenabled=1
End Sub

sub dingwallf_timer
	select case Dingwallf.uservalue
		Case 1: Slingf1.visible=0: Slingf.visible=1
		case 2:	Slingf.visible=0: Slingf2.visible=1
		Case 3: Slingf2.visible=0: Slingf.visible=1: Me.timerenabled=0
	end Select
	dingwallf.uservalue=Dingwallf.uservalue+1
end sub

Sub DingwallG_Hit()
	If TableTilted=false then
		AddScore(10)
	End if
	Slingg.visible=0
	Slingg1.visible=1
	dingwallg.uservalue=1
	Me.timerenabled=1
End Sub

sub dingwallg_timer
	select case Dingwallg.uservalue
		Case 1: Slingg1.visible=0: Slingg.visible=1
		case 2:	Slingg.visible=0: Slingg2.visible=1
		Case 3: Slingg2.visible=0: Slingg.visible=1: Me.timerenabled=0
	end Select
	dingwallg.uservalue=Dingwallg.uservalue+1
end sub

'***********************
' bumpers
'***********************

Sub Bumper1_Hit
	If TableTilted=false then
		playsound SoundFXDOF("fx_bumper4",107,DOFPulse,DOFContactors)
		DOF 106,DOFPulse
		bump1 = 1
		AddScore(100)
		ToggleBumper
    end if

End Sub

Sub Bumper2_Hit
	If TableTilted=false then
		playsound SoundFXDOF("fx_bumper4",109,DOFPulse,DOFContactors)
		DOF 105,DOFPulse
		bump2 = 1
		AddScore(100)
		ToggleBumper
    end if

End Sub

'***********************
' targets
'***********************

Sub TargetLowerLeft1_Hit()
	If TableTilted=false then
		DOF 127,2
		If BallsPerGame=3 then
			AddScore(1000)
		else
			AddScore(100)
			ToggleBumper
		end if
	end if
end Sub

Sub TargetLowerLeft2_Hit()
	If TableTilted=false then
		DOF 127,2
		If BallsPerGame=3 then
			AddScore(1000)
		else
			AddScore(100)
			ToggleBumper
		end if
	end if
end Sub

'***********************
' rollover triggers
'***********************

Sub TriggerTopA_Hit()
	If TableTilted=false then
		PlaySound "sensor"
		DOF 121,2
		SetMotor(500)

	end if
End Sub

Sub TriggerTopB_Hit()
	If TableTilted=false then
		PlaySound "sensor"
		DOF 120,2
		If UpperRolloverLight1.state=1 then
			SetMotor(5000)
		else
			SetMotor(500)
		end if
	end if
End Sub

Sub TriggerTopC_Hit()
	If TableTilted=false then
		PlaySound "sensor"
		DOF 122,2
		If UpperRolloverLight2.state=1 then
			SetMotor(5000)
		else
			SetMotor(500)
		end if
	end if
End Sub

Sub TriggerLowerRightRollover1_Hit()
	If TableTilted=false then
		PlaySound "sensor"
		DOF 118,2
		If RightRolloverLight1.state=0 then
			SetMotor(500)
		Else
			SetMotor(500)
			ExtraBall=True
			SamePlayerShootsAgain.state=1
			RightRolloverLight1.state=0
			EnableSpecialEB=0
			If B2SOn Then
				Controller.B2SSetShootAgain 1
			end if
		End If

	End if
End Sub

Sub TriggerLowerRightRollover2_Hit()
	If TableTilted=false then
		PlaySound "sensor"
		DOF 119,2
		If RightRolloverLight2.state=0 then
			SetMotor(500)
		Else
			SetMotor(500)
			AddSpecial
			RightRolloverLight2.state=0
			EnableSpecialEB=0
		End If

	End if
End Sub

Sub TriggerLowerLeftRollover_Hit()
	If TableTilted=false then
		PlaySound "sensor"
		DOF 117,2
		If LeftRolloverLight.state=0 then
			SetMotor(500)
		Else
			SetMotor(500)
			ExtraBall=True
			SamePlayerShootsAgain.state=1
			LeftRolloverLight.state=0
			If B2SOn Then
				Controller.B2SSetShootAgain 1
			end if
		end if
	End If
End Sub

Sub TriggerLeftInlane_Hit()
	If TableTilted=false then
		PlaySound "sensor"
		DOF 114,2
		AddScore(1000)
	End If
End Sub

Sub TriggerRightInlane_Hit()
	If TableTilted=false then
		PlaySound "sensor"
		DOF 115,2
		AddScore(1000)
	End If
End Sub

Sub TriggerLeftOutlane_Hit()
	If TableTilted=false then
		PlaySound "sensor"
		DOF 113,2
		if LeftOutlaneLight.state=1 then
			SetMotor(5000)
		Else
			SetMotor(500)
		end if
	End If
End Sub

Sub TriggerRightOutlane_Hit()
	If TableTilted=false then
		PlaySound "sensor"
		DOF 116,2
		if RightOutlaneLight.state=1 then
			SetMotor(5000)
		Else
			SetMotor(500)
		end if
	End If
End Sub

'***********************
' drop targets
'***********************

Sub ScoreUpperDrops()
	if BallsPerGame=3 then
		SetMotor(2000)
	else
		SetMotor(300)
	end if
end sub

Sub ScoreLowerDrops()
	if BallsPerGame=3 then
		if LowerTargetValue=0 then
			SetMotor(2000)
		else
			SetMotor(500)
		end if
	else
		if LowerTargetValue=0 then
			SetMotor(300)
		else
			SetMotor(500)
		end if
	end if
end sub

Sub TopTarget1_dropped()
	If TableTilted=false then
		playsound SoundFXDOF("droptargetdropped",123,DOFPulse,DOFTargets)
		ScoreUpperDrops
		bonuscounter=bonuscounter+1
		TopTarget1.IsDropped = True
		TargetLight1.state=1
		CheckAllDropTargets
	end if
end sub

Sub TopTarget2_dropped()
	If TableTilted=false then
		playsound SoundFXDOF("droptargetdropped",123,DOFPulse,DOFTargets)
		ScoreUpperDrops
		bonuscounter=bonuscounter+1

		TopTarget2.IsDropped = True
		TargetLight2.state=1
		CheckAllDropTargets
	end if
end sub

Sub TopTarget3_dropped()
	If TableTilted=false then
		playsound SoundFXDOF("droptargetdropped",123,DOFPulse,DOFTargets)
		ScoreUpperDrops
		bonuscounter=bonuscounter+1

		TopTarget3.IsDropped = True
		TargetLight3.state=1
		CheckAllDropTargets
	end if
end sub

Sub TopTarget4_dropped()
	If TableTilted=false then
		playsound SoundFXDOF("droptargetdropped",124,DOFPulse,DOFTargets)
		ScoreUpperDrops
		bonuscounter=bonuscounter+1

		TopTarget4.IsDropped = True
		TargetLight4.state=1
		CheckAllDropTargets
	end if
end sub

Sub TopTarget5_dropped()
	If TableTilted=false then
		playsound SoundFXDOF("droptargetdropped",124,DOFPulse,DOFTargets)
		ScoreUpperDrops
		bonuscounter=bonuscounter+1

		TopTarget5.IsDropped = True
		TargetLight5.state=1
		CheckAllDropTargets
	end if
end sub

Sub TopTarget6_dropped()
	If TableTilted=false then
		playsound SoundFXDOF("droptargetdropped",124,DOFPulse,DOFTargets)
		ScoreUpperDrops
		bonuscounter=bonuscounter+1

		TopTarget6.IsDropped = True
		TargetLight6.state=1
		CheckAllDropTargets
	end if
end sub

Sub TopTarget7_dropped()
	If TableTilted=false then
		playsound SoundFXDOF("droptargetdropped",124,DOFPulse,DOFTargets)
		ScoreUpperDrops
		bonuscounter=bonuscounter+1

		TopTarget7.IsDropped = True
		TargetLight7.state=1
		CheckAllDropTargets
	end if
end sub

Sub TopTarget8_dropped()
	If TableTilted=false then
		playsound SoundFXDOF("droptargetdropped",125,DOFPulse,DOFTargets)
		ScoreUpperDrops
		bonuscounter=bonuscounter+1

		TopTarget8.IsDropped = True
		TargetLight8.state=1
		CheckAllDropTargets
	end if
end sub

Sub TopTarget9_dropped()
	If TableTilted=false then
		playsound SoundFXDOF("droptargetdropped",125,DOFPulse,DOFTargets)
		ScoreUpperDrops
		bonuscounter=bonuscounter+1

		TopTarget9.IsDropped = True
		TargetLight9.state=1
		CheckAllDropTargets
	end if
end sub

Sub TopTarget10_dropped()
	If TableTilted=false then
		playsound SoundFXDOF("droptargetdropped",125,DOFPulse,DOFTargets)
		ScoreUpperDrops
		bonuscounter=bonuscounter+1

		TopTarget10.IsDropped = True
		TargetLight10.state=1
		CheckAllDropTargets
	end if
end sub

Sub LowerTarget1_dropped()
	If TableTilted=false then
		playsound SoundFXDOF("droptargetdropped",126,DOFPulse,DOFTargets)
		ScoreLowerDrops
		bonuscounter=bonuscounter+1

		LowerTarget1.IsDropped = True
		LowerTargetLight1.state=1
		CheckAllDropTargets
	end if
end sub

Sub LowerTarget2_dropped()
	If TableTilted=false then
		playsound SoundFXDOF("droptargetdropped",126,DOFPulse,DOFTargets)
		ScoreLowerDrops
		bonuscounter=bonuscounter+1

		LowerTarget2.IsDropped = True
		LowerTargetLight2.state=1
		CheckAllDropTargets
	end if
end sub

Sub LowerTarget3_dropped()
	If TableTilted=false then
		playsound SoundFXDOF("droptargetdropped",126,DOFPulse,DOFTargets)
		ScoreLowerDrops
		bonuscounter=bonuscounter+1

		LowerTarget3.IsDropped = True
		LowerTargetLight3.state=1
		CheckAllDropTargets
	end if
end sub

Sub LowerTarget4_dropped()
	If TableTilted=false then
		playsound SoundFXDOF("droptargetdropped",126,DOFPulse,DOFTargets)
		ScoreLowerDrops
		bonuscounter=bonuscounter+1

		LowerTarget4.IsDropped = True
		LowerTargetLight4.state=1
		CheckAllDropTargets
	end if
end sub

Sub LowerTarget5_dropped()
	If TableTilted=false then
		playsound SoundFXDOF("droptargetdropped",126,DOFPulse,DOFTargets)
		ScoreLowerDrops
		bonuscounter=bonuscounter+1

		LowerTarget5.IsDropped = True
		LowerTargetLight5.state=1
		CheckAllDropTargets
	end if
end sub

Sub CheckAllDropTargets()
	dim topdone
	dim bottomdone
	topdone=0
	bottomdone=0
	for each obj in TopTargetLights
		if obj.state=1 then
			topdone=topdone+1
		end if
	next
	for each obj in LowerTargetLights
		if obj.state=1 then
			bottomdone=bottomdone+1
		end if
	next
	if bottomdone=5 then
		LeftRolloverLight.state=1
	end if
	if topdone=10 then
		EnableSpecialEB=1
		if SpecialEBOption=0 then
			RightRolloverLight1.state=1
			RightRolloverLight2.state=1
		else
			RightRolloverLight1.state=1
			RightRolloverLight2.state=0
		end if
	end if
end sub

Sub AddSpecial()
	PlaySound SoundFXDOF("knock",112,DOFPulse,DOFKnocker)
	Credits=Credits+1
	DOF 111, 1
	if Credits>15 then Credits=15
	If B2SOn Then
		Controller.B2SSetCredits Credits
	End If
	Credittxt.SetValue(Credits)
End Sub

Sub AddSpecial2()
	PlaySound"click"
	Credits=Credits+1
	DOF 111, 1
	if Credits>15 then Credits=15
	If B2SOn Then
		Controller.B2SSetCredits Credits
	End If
	Credittxt.SetValue(Credits)
End Sub


Sub AddBonus()
	bonuscountdown=15
	ScoreBonus.enabled=true
End Sub



Sub ToggleBumper

	if UpperRolloverLight1.state=1 then
		UpperRolloverLight1.state=0
		UpperRolloverLight2.state=1
	else
		UpperRolloverLight1.state=1
		UpperRolloverLight2.state=0
	end if
	if LeftOutlaneLight.state=1 then
		LeftOutlaneLight.state=0
		RightOutlaneLight.state=1
	else
		LeftOutlaneLight.state=1
		RightOutlaneLight.state=0
	end if
	if (SpecialEBOption=1) and (EnableSpecialEB=1) then
		if RightRolloverLight1.state=1 then
			RightRolloverLight1.state=0
			RightRolloverLight2.state=1
		else
			RightRolloverLight1.state=1
			RightRolloverLight2.state=0
		end if
	end if
end sub


Sub ResetBallDrops()
	If (TopTarget1.IsDropped=True or TopTarget2.IsDropped=True or TopTarget3.IsDropped=True) Then DOF 128, 2
	If (TopTarget4.IsDropped=True or TopTarget5.IsDropped=True or TopTarget6.IsDropped=True or TopTarget7.IsDropped=True) Then DOF 129, 2
	If (TopTarget8.IsDropped=True or TopTarget9.IsDropped=True or TopTarget10.IsDropped=True) Then DOF 130, 2
	If (LowerTarget1.IsDropped=True or LowerTarget2.IsDropped=True or LowerTarget3.IsDropped=True or LowerTarget4.IsDropped=True or LowerTarget5.IsDropped=True) Then DOF 131, 2

	for each obj in UpperTargets
		obj.IsDropped=0
	next
	for each obj in LowerTargets
		obj.IsDropped=0
	next

	for each obj in TopTargetLights
		obj.state=0
	next
	for each obj in LowerTargetLights
		obj.state=0
	next
	RightRolloverLight1.state=0
	RightRolloverLight2.state=0
	LeftRolloverLight.state=0
	SamePlayerShootsAgain.state=0
	EnableSpecialEB=0
	DoubleBonus.state=0
	BonusCounter=0
	HoleCounter=0
	BumperLight1.state=1
	BumperLight2.state=1
	BumpersOn
End Sub


Sub LightsOut
	BonusCounter=0
	HoleCounter=0
	BumpersOff
end sub

Sub ResetBalls()
	TempMultiCounter=BallsPerGame-BallInPlay
	ResetBallDrops
	BonusMultiplier=1
	DoubleBonus.state=0
	If (BallsPerGame=BallInPlay) then
		DoubleBonus.state=1
		BonusMultiplier=2
	End If
	TableTilted=false
	tilttxt.text=""
	If B2Son then
		Controller.B2SSetTilt 0
	end if
	PlasticsOn
	Drain.Kick 60, 45, 0
	DOF 108,2
	BIPReel.SetValue(BallInPlay)
End Sub


sub resettimer_timer
    rst=rst+1
    for i=1 to 4
 	If B2SOn Then
		Controller.B2SSetScorePlayer i, 0
	End If
	next
    if rst=20 then
    playsound "StartBall1"
    end if
    if rst=24 then
    newgame
    resettimer.enabled=false
    end if
end sub

sub ScoreBonus_timer
	Dim temprelaycount
	Dim temploopcount
	Dim tempbonuscount
	Dim t
	Dim relayflag
	BumpersOff
	if Bonus(bonuscountdown-1).state=1 then
		SetMotor2(1000*BonusMultiplier)
		ScoreBonus.interval=145
	else
		ScoreBonus.interval=400
	end if
	relayflag=false
	temprelaycount=0
	t=1
	for t = 1 to 5
		if ((bonuscountdown-t)>=0) and (Bonus(bonuscountdown-1).state=0) then
			if (Bonus(bonuscountdown-t).state=0) then
				temprelaycount=temprelaycount+1
			else
				t=5
			end if
		else
			t=5
		end if
	next
	Select Case temprelaycount
		Case 0:
			PlaySound"BonusRelays1"
			bonuscountdown=bonuscountdown-1
		Case 1:
			PlaySound"BonusRelays1"
			bonuscountdown=bonuscountdown-1
		Case 2:
			PlaySound"BonusRelays2"
			bonuscountdown=bonuscountdown-2
		Case 3:
			PlaySound"BonusRelays3"
			bonuscountdown=bonuscountdown-3
		Case 4:
			PlaySound"BonusRelays4"
			bonuscountdown=bonuscountdown-4
		Case 5:
			PlaySound"BonusRelays5"
			bonuscountdown=bonuscountdown-5
	end Select

	if bonuscountdown<=0 then
		ScoreBonus.enabled=false
		ScoreBonus.interval=500
		NextBallDelay.enabled=true
	end if

end sub



sub NextBallDelay_timer()
	NextBallDelay.enabled=false
	nextball

end sub

sub newgame
	InProgress=true
	queuedscore=0
	for i = 1 to 4
		Score(i)=0
		Score100K(1)=0
		HighScorePaid(i)=false
		Replay1Paid(i)=false
		Replay2Paid(i)=false

	next
	If B2SOn Then
		Controller.B2SSetTilt 0
		Controller.B2SSetGameOver 0
		Controller.B2SSetMatch 0
		Controller.B2SSetShootAgain 0
'		Controller.B2SSetScorePlayer1 0
'		Controller.B2SSetScorePlayer2 0
'		Controller.B2SSetScorePlayer3 0
'		Controller.B2SSetScorePlayer4 0
		Controller.B2SSetBallInPlay BallInPlay
	End if
	BumperLight1.state=1
	BumperLight2.state=1
	BumpersOn
	UpperRolloverLight1.state=1
	UpperRolloverLight2.state=0
	LeftOutlaneLight.state=0
	RightOutlaneLight.state=1
	RightRolloverLight1.state=0
	RightRolloverLight2.state=0
	LeftRolloverLight.state=0
	SamePlayerShootsAgain.state=0
	DoubleBonus.state=0
	For each obj in TopTargetLights
		obj.state=0
	next
	For each obj in LowerTargetLights
		obj.state=0
	next
	ResetBalls
End sub

sub nextball
	If B2SOn Then
		Controller.B2SSetTilt 0
		Controller.B2SSetShootAgain 0
	End If
	If ExtraBall=false then
		Player=Player+1
	else
		ExtraBall=false
		SamePlayerShootsAgain.state=0
	end if
	If Player>Players Then
		BallInPlay=BallInPlay+1
		If BallInPlay>BallsPerGame then
			PlaySound("MotorLeer")
			InProgress=false
			If B2SOn Then
				Controller.B2SSetGameOver 1
				Controller.B2SSetPlayerUp 0
				Controller.B2SSetData 81,0
				Controller.B2SSetData 82,0
				Controller.B2SSetData 83,0
				Controller.B2SSetData 84,0
				Controller.B2SSetBallInPlay 0
				Controller.B2SSetCanPlay 0
			End If
			For each obj in PlayerHuds
				obj.State =0
			next
			For each obj in PlayerHUDScores
				obj.state=0
			next
			BIPReel.SetValue(0)
			CanPlayReel.SetValue(0)

			gamov.text="GAME OVER"
			LeftFlipper.RotateToStart
			RightFlipper.RotateToStart
			LightsOut
'			PlasticsOff
			checkmatch
			CheckHighScore
			Players=0
		Else
			Player=1
			If B2SOn Then
				Controller.B2SSetPlayerUp Player
				Controller.B2SSetData 81,0
				Controller.B2SSetData 82,0
				Controller.B2SSetData 83,0
				Controller.B2SSetData 84,0
				Controller.B2SSetData 80+Player,1
				Controller.B2SSetBallInPlay BallInPlay

			End If
			PlaySound("RotateThruPlayers")
			TempPlayerUp=Player
			PlayerUpRotator.enabled=true
			PlayStartBall.enabled=true
			For each obj in PlayerHuds
				obj.State = 0
			next
			For each obj in PlayerHUDScores
				obj.state=0
			next
			PlayerHuds(Player-1).State =1
			PlayerHUDScores(Player-1).state=1
			ResetBalls
		End If
	Else
		If B2SOn Then
			Controller.B2SSetPlayerUp Player
			Controller.B2SSetData 81,0
			Controller.B2SSetData 82,0
			Controller.B2SSetData 83,0
			Controller.B2SSetData 84,0
			Controller.B2SSetData 80+Player,1
			Controller.B2SSetBallInPlay BallInPlay
		End If
		PlaySound("RotateThruPlayers")
		TempPlayerUp=Player
		PlayerUpRotator.enabled=true
		PlayStartBall.enabled=true
		For each obj in PlayerHuds
			obj.State=0
		next
		For each obj in PlayerHUDScores
				obj.state=0
			next
		PlayerHuds(Player-1).State = 1
		PlayerHUDScores(Player-1).state=1
		ResetBalls
	End If

End sub

sub CheckHighScore

  dim hiscstate
	  hiscstate=0

	  for i=1 to players
		if score(i)>highscore then
			highscore=score(i)
			hiscstate=1
		end if
	  next
	  if hiscstate=1 then HighScoreEntryInit()
	  UpdatePostIt
	savehs
end sub


sub checkmatch
	Match=Int(Rnd*10)*10
	matchtxt.text = Match
	If B2SOn Then
		If Match = 0 Then
			Controller.B2SSetMatch 100
		Else
			Controller.B2SSetMatch Match
		End If
	End if
	for i = 1 to Players
		if Match=(Score(i) mod 100) then
			AddSpecial
		end if
	next
end sub

Sub TiltTimer_Timer()
	if TiltCount > 0 then TiltCount = TiltCount - 1
	if TiltCount = 0 then
		TiltTimer.Enabled = False
	end if
end sub

Sub TiltIt()
		TiltCount = TiltCount + 1
		if TiltCount = 3 then
			TableTilted=True
			PlasticsOff
			BumpersOff
			for each obj in Bonus
				obj.state=0
			next
			LeftFlipper.RotateToStart
			RightFlipper.RotateToStart
'			TiltReel.SetValue(1)
			tilttxt.text="TILT"
			If B2Son then
				Controller.B2SSetTilt 1
			end if
		else
			TiltTimer.Interval = 500
			TiltTimer.Enabled = True
		end if

end sub

Sub IncreaseBonus()

	If BonusCounter=10 then

	else
		If BonusCounter>0 then
			Bonus(BonusCounter).state=0
		end if
		BonusCounter=BonusCounter+1
		Bonus(BonusCounter).State=1
'		PlaySound("Score100")
	end if
	if BonusCounter=10 then
		TopLeftTargetLight.state=1
		TopRightTargetLight.state=1
	end if
End Sub


Sub BonusBoost_Timer()
	IncreaseBonus
	BonusBoosterCounter=BonusBoosterCounter-1
	If BonusBoosterCounter=0 then
		BonusBoost.enabled=false
	end if

end sub

Sub CheckForLightSpecial()

	if (TopLightA.state=0) and (TopLightB.state=0) and (TopLightC.state=0) then
		TopRightTargetLight.State=1
		TopLeftTargetLight.State=1
	end if
end sub

Sub PlayStartBall_timer()

	PlayStartBall.enabled=false
	PlaySound("StartBall2-5")
end sub

Sub PlayerUpRotator_timer()
		If RotatorTemp<5 then
			TempPlayerUp=TempPlayerUp+1
			If TempPlayerUp>4 then
				TempPlayerUp=1
			end if
			For each obj in PlayerHuds
				obj.state = 0
			next
			For each obj in PlayerHUDScores
				obj.state=0
			next
			PlayerHuds(TempPlayerUp-1).State = 1
			PlayerHUDScores(TempPlayerUp-1).state=1
			If B2SOn Then
				Controller.B2SSetPlayerUp TempPlayerUp
				Controller.B2SSetData 81,0
				Controller.B2SSetData 82,0
				Controller.B2SSetData 83,0
				Controller.B2SSetData 84,0
				Controller.B2SSetData 80+TempPlayerUp,1
			End If

		else
			if B2SOn then
				Controller.B2SSetPlayerUp Player
				Controller.B2SSetData 81,0
				Controller.B2SSetData 82,0
				Controller.B2SSetData 83,0
				Controller.B2SSetData 84,0
				Controller.B2SSetData 80+Player,1
			end if
			PlayerUpRotator.enabled=false
			RotatorTemp=1
			For each obj in PlayerHuds
				obj.State = 0
			next
			For each obj in PlayerHUDScores
				obj.state=0
			next
			PlayerHuds(Player-1).State = 1
			PlayerHUDScores(Player-1).state=1
		end if
		RotatorTemp=RotatorTemp+1


end sub


sub InitPauser5_timer
		Credittxt.SetValue(Credits)
		InitPauser5.enabled=false
end sub



sub BumpersOff
	For each light in bumperlights
		light.state=0
	next
end sub

sub BumpersOn
	For each light in bumperlights
		light.state=1
	next
end sub

Sub PlasticsOn
	For each light in gilights
		light.state=1
	next
end sub

Sub PlasticsOff
	For each light in gilights
		light.state=0
	next
end sub



'****************************************
'  SCORE MOTOR
'****************************************

ScoreMotorTimer.Enabled = 1
ScoreMotorTimer.Interval = 135 '135
AddScoreTimer.Enabled = 1
AddScoreTimer.Interval = 135

Dim queuedscore
Dim MotorMode
Dim MotorPosition

Sub SetMotor(y)
	Select Case ScoreMotorAdjustment
		Case 0:
			queuedscore=queuedscore+y
		Case 1:
			If MotorRunning<>1 And InProgress=true then
				queuedscore=queuedscore+y
			end if
		end Select
end sub

Sub SetMotor2(x)
	If MotorRunning<>1 And InProgress=true then
		MotorRunning=1

		Select Case x
			Case 10:
				AddScore(10)
				MotorRunning=0
				BumpersOn

			Case 20:
				MotorMode=10
				MotorPosition=2
				BumpersOff
			Case 30:
				MotorMode=10
				MotorPosition=3
				BumpersOff
			Case 40:
				MotorMode=10
				MotorPosition=4
				BumpersOff
			Case 50:
				MotorMode=10
				MotorPosition=5
				BumpersOff
			Case 100:
				AddScore(100)
				MotorRunning=0
				BumpersOn
			Case 200:
				MotorMode=100
				MotorPosition=2
				BumpersOff
			Case 300:
				MotorMode=100
				MotorPosition=3
				BumpersOff
			Case 400:
				MotorMode=100
				MotorPosition=4
				BumpersOff
			Case 500:
				MotorMode=100
				MotorPosition=5
				BumpersOff
			Case 1000:
				AddScore(1000)
				MotorRunning=0
				BumpersOn
			Case 2000:
				MotorMode=1000
				MotorPosition=2
				BumpersOff
			Case 3000:
				MotorMode=1000
				MotorPosition=3
				BumpersOff
			Case 4000:
				MotorMode=1000
				MotorPosition=4
				BumpersOff
			Case 5000:
				MotorMode=1000
				MotorPosition=5
				BumpersOff
		End Select
	End If
End Sub

Sub AddScoreTimer_Timer
	Dim tempscore


	If MotorRunning<>1 And InProgress=true then
		if queuedscore>=5000 then
			tempscore=5000
			queuedscore=queuedscore-5000
			SetMotor2(5000)
			exit sub
		end if
		if queuedscore>=4000 then
			tempscore=4000
			queuedscore=queuedscore-4000
			SetMotor2(4000)
			exit sub
		end if

		if queuedscore>=3000 then
			tempscore=3000
			queuedscore=queuedscore-3000
			SetMotor2(3000)
			exit sub
		end if

		if queuedscore>=2000 then
			tempscore=2000
			queuedscore=queuedscore-2000
			SetMotor2(2000)
			exit sub
		end if

		if queuedscore>=1000 then
			tempscore=1000
			queuedscore=queuedscore-1000
			SetMotor2(1000)
			exit sub
		end if

		if queuedscore>=500 then
			tempscore=500
			queuedscore=queuedscore-500
			SetMotor2(500)
			exit sub
		end if
		if queuedscore>=400 then
			tempscore=400
			queuedscore=queuedscore-400
			SetMotor2(400)
			exit sub
		end if
		if queuedscore>=300 then
			tempscore=300
			queuedscore=queuedscore-300
			SetMotor2(300)
			exit sub
		end if
		if queuedscore>=200 then
			tempscore=200
			queuedscore=queuedscore-200
			SetMotor2(200)
			exit sub
		end if
		if queuedscore>=100 then
			tempscore=100
			queuedscore=queuedscore-100
			SetMotor2(100)
			exit sub
		end if

		if queuedscore>=50 then
			tempscore=50
			queuedscore=queuedscore-50
			SetMotor2(50)
			exit sub
		end if
		if queuedscore>=40 then
			tempscore=40
			queuedscore=queuedscore-40
			SetMotor2(40)
			exit sub
		end if
		if queuedscore>=30 then
			tempscore=30
			queuedscore=queuedscore-30
			SetMotor2(30)
			exit sub
		end if
		if queuedscore>=20 then
			tempscore=20
			queuedscore=queuedscore-20
			SetMotor2(20)
			exit sub
		end if
		if queuedscore>=10 then
			tempscore=10
			queuedscore=queuedscore-10
			SetMotor2(10)
			exit sub
		end if


	End If


end Sub

Sub ScoreMotorTimer_Timer
	If MotorPosition > 0 Then
		Select Case MotorPosition
			Case 5,4,3,2:
				If MotorMode=1000 Then
					AddScore(1000)
				end if
				if MotorMode=100 then
					AddScore(100)
				End If
				if MotorMode=10 then
					AddScore(10)
				End if
				MotorPosition=MotorPosition-1
			Case 1:
				If MotorMode=1000 Then
					AddScore(1000)
				end if
				If MotorMode=100 then
					AddScore(100)
				End If
				if MotorMode=10 then
					AddScore(10)
				End if
				MotorPosition=0:MotorRunning=0:BumpersOn
		End Select
	End If

End Sub


Sub AddScore(x)
	Select Case ScoreAdditionAdjustment
		Case 0:
			AddScore1(x)
		Case 1:
			AddScore2(x)
	end Select

end sub


Sub AddScore1(x)
'	debugtext.text=score
	Select Case x
		Case 1:
			PlayChime(10)
			Score(Player)=Score(Player)+1

		Case 10:
			PlayChime(10)
			Score(Player)=Score(Player)+10
'			debugscore=debugscore+10

		Case 100:
			PlayChime(100)
			Score(Player)=Score(Player)+100
'			debugscore=debugscore+100

		Case 1000:
			PlayChime(1000)
			Score(Player)=Score(Player)+1000
'			debugscore=debugscore+1000
	End Select
	PlayerScores(Player-1).AddValue(x)
	If ScoreDisplay(Player)<100000 then
		ScoreDisplay(Player)=Score(Player)
	Else
		Score100K(Player)=Int(Score(Player)/100000)
		ScoreDisplay(Player)=Score(Player)-100000
	End If
	if Score(Player)=>100000 then
		If B2SOn Then
			If Player=1 Then
				Controller.B2SSetScoreRolloverPlayer1 Score100K(Player)
			End If
			If Player=2 Then
				Controller.B2SSetScoreRolloverPlayer2 Score100K(Player)
			End If

			If Player=3 Then
				Controller.B2SSetScoreRolloverPlayer3 Score100K(Player)
			End If

			If Player=4 Then
				Controller.B2SSetScoreRolloverPlayer4 Score100K(Player)
			End If
		End If
	End If
	If B2SOn Then
		Controller.B2SSetScorePlayer Player, ScoreDisplay(Player)
	End If
	If Score(Player)>Replay1 and Replay1Paid(Player)=false then
		Replay1Paid(Player)=True
		AddSpecial
	End If
	If Score(Player)>Replay2 and Replay2Paid(Player)=false then
		Replay2Paid(Player)=True
		AddSpecial
	End If
'	ScoreText.text=debugscore
End Sub

Sub AddScore2(x)
	Dim OldScore, NewScore, OldTestScore, NewTestScore
    OldScore = Score(Player)

	Select Case x
        Case 1:
            Score(Player)=Score(Player)+1
		Case 10:
			Score(Player)=Score(Player)+10
		Case 100:
			Score(Player)=Score(Player)+100
		Case 1000:
			Score(Player)=Score(Player)+1000
	End Select
	NewScore = Score(Player)
	if Score(Player)=>100000 then
		If B2SOn Then
			If Player=1 Then
				Controller.B2SSetScoreRolloverPlayer1 1
			End If
			If Player=2 Then
				Controller.B2SSetScoreRolloverPlayer2 1
			End If

			If Player=3 Then
				Controller.B2SSetScoreRolloverPlayer3 1
			End If

			If Player=4 Then
				Controller.B2SSetScoreRolloverPlayer4 1
			End If
		End If
	End If

	OldTestScore = OldScore
	NewTestScore = NewScore
	Do
		if OldTestScore < Replay1 and NewTestScore >= Replay1 then
			AddSpecial()
			NewTestScore = 0
		Elseif OldTestScore < Replay2 and NewTestScore >= Replay2 then
			AddSpecial()
			NewTestScore = 0
		End if
		NewTestScore = NewTestScore - 100000
		OldTestScore = OldTestScore - 100000
	Loop While NewTestScore > 0

    OldScore = int(OldScore / 10)	' divide by 10 for games with fixed 0 in 1s position, by 1 for games with real 1s digits
    NewScore = int(NewScore / 10)	' divide by 10 for games with fixed 0 in 1s position, by 1 for games with real 1s digits
	' MsgBox("OldScore="&OldScore&", NewScore="&NewScore&", OldScore Mod 10="&OldScore Mod 10 & ", NewScore % 10="&NewScore Mod 10)

    if (OldScore Mod 10 <> NewScore Mod 10) then
		PlayChime(10)

    end if

    OldScore = int(OldScore / 10)
    NewScore = int(NewScore / 10)
	' MsgBox("OldScore="&OldScore&", NewScore="&NewScore)
    if (OldScore Mod 10 <> NewScore Mod 10) then
		PlayChime(100)

    end if

    OldScore = int(OldScore / 10)
    NewScore = int(NewScore / 10)
	' MsgBox("OldScore="&OldScore&", NewScore="&NewScore)
    if (OldScore Mod 10 <> NewScore Mod 10) then
		PlayChime(1000)

    end if

    OldScore = int(OldScore / 10)
    NewScore = int(NewScore / 10)
	' MsgBox("OldScore="&OldScore&", NewScore="&NewScore)
    if (OldScore Mod 10 <> NewScore Mod 10) then
		PlayChime(1000)
    end if

	If B2SOn Then
		Controller.B2SSetScorePlayer Player, Score(Player)
	End If
'	EMReel1.SetValue Score(Player)
	PlayerScores(Player-1).AddValue(x)
End Sub



Sub PlayChime(x)
	if ChimesOn=0 then
		Select Case x
			Case 10
				If LastChime10=1 Then
					PlaySound "SpinACard_1_10_Point_Bell"
					LastChime10=0
				Else
					PlaySound "SpinACard_1_10_Point_Bell"
					LastChime10=1
				End If
			Case 100
				If LastChime100=1 Then
					PlaySound "SpinACard_100_Point_Bell"
					LastChime100=0
				Else
					PlaySound "SpinACard_100_Point_Bell"
					LastChime100=1
				End If

		End Select
	else
		Select Case x
			Case 10
				If LastChime10=1 Then
					PlaySound SoundFXDOF("SJ_Chime_10a",141,DOFPulse,DOFChimes)
					LastChime10=0
				Else
					PlaySound SoundFXDOF("SJ_Chime_10b",141,DOFPulse,DOFChimes)
					LastChime10=1
				End If
			Case 100
				If LastChime100=1 Then
					PlaySound SoundFXDOF("SJ_Chime_100a",142,DOFPulse,DOFChimes)
					LastChime100=0
				Else
					PlaySound SoundFXDOF("SJ_Chime_100b",142,DOFPulse,DOFChimes)
					LastChime100=1
				End If
			Case 1000
				If LastChime1000=1 Then
					PlaySound SoundFXDOF("SJ_Chime_1000a",143,DOFPulse,DOFChimes)
					LastChime1000=0
				Else
					PlaySound SoundFXDOF("SJ_Chime_1000b",143,DOFPulse,DOFChimes)
					LastChime1000=1
				End If
		End Select
	end if
End Sub


Sub CheckTilt
	If Tilttimer.Enabled = True Then
	 TiltSens = TiltSens + 1
	 if TiltSens = 3 Then GameTilted
	Else
	 TiltSens = 0
	 Tilttimer.Enabled = True
	End If
End Sub

Sub Tilttimer_Timer()
	Tilttimer.Enabled = False
End Sub

Sub GameTilted
	Tilt = True
	tilttxt.text="TILT"
	If B2SOn Then Controller.B2SSetTilt 33,1
    If B2SOn Then Controller.B2ssetdata 1, 0
	playsound "target"
	for each light in GI: light.state=0: Next
	For each light in BumperLights: light.state=0:Next
	turnoff
	if tiltgame=1 then eg=1
End Sub

sub turnoff
	for i=1 to 2
		EVAL("Bumper"&i).hashitevent = 0
	Next
  	LeftFlipper.RotateToStart
	StopSound "Buzz"
	DOF 101, DOFOff
	RightFlipper.RotateToStart
	StopSound "Buzz1"
	DOF 102, DOFOff
end sub

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


sub savehs

    savevalue "TargetAlpha", "credit", credits
    savevalue "TargetAlpha", "hiscore", HighScore
    savevalue "TargetAlpha", "match", match
    savevalue "TargetAlpha", "score1", score(1)
    savevalue "TargetAlpha", "score2", score(2)
    savevalue "TargetAlpha", "score3", score(3)
    savevalue "TargetAlpha", "score4", score(4)
	savevalue "TargetAlpha", "balls", BallsPerGame
	savevalue "TargetAlpha", "replays", ReplayLevel
	savevalue "TargetAlpha", "hsa1", HSA1
	savevalue "TargetAlpha", "hsa2", HSA2
	savevalue "TargetAlpha", "hsa3", HSA3
	savevalue "TargetAlpha", "freeplay", freeplay

end sub

sub loadhs
    dim temp
	temp = LoadValue("TargetAlpha", "credit")
    If (temp <> "") then credits = CDbl(temp)
    temp = LoadValue("TargetAlpha", "hiscore")
    If (temp <> "") then HighScore = CDbl(temp)
    temp = LoadValue("TargetAlpha", "match")
    If (temp <> "") then match = CDbl(temp)
    temp = LoadValue("TargetAlpha", "score1")
    If (temp <> "") then score(1) = CDbl(temp)
    temp = LoadValue("TargetAlpha", "score2")
    If (temp <> "") then score(2) = CDbl(temp)
    temp = LoadValue("TargetAlpha", "score3")
    If (temp <> "") then score(3) = CDbl(temp)
    temp = LoadValue("TargetAlpha", "score4")
    If (temp <> "") then score(4) = CDbl(temp)
    temp = LoadValue("TargetAlpha", "balls")
    If (temp <> "") then BallsPerGame = CDbl(temp)
    temp = LoadValue("TargetAlpha", "replays")
    If (temp <> "") then ReplayLevel = CDbl(temp)
    temp = LoadValue("TargetAlpha", "hsa1")
    If (temp <> "") then HSA1 = CDbl(temp)
    temp = LoadValue("TargetAlpha", "hsa2")
    If (temp <> "") then HSA2 = CDbl(temp)
    temp = LoadValue("TargetAlpha", "hsa3")
    If (temp <> "") then HSA3 = CDbl(temp)
    temp = LoadValue("TargetAlpha", "freeplay")
    If (temp <> "") then freeplay = CDbl(temp)
end sub

Sub TargetAlpha_Exit()
	Savehs
	turnoff
	If B2SOn Then Controller.stop
End Sub

'==========================================================================================================================================
'============================================================= START OF HIGH SCORES ROUTINES =============================================================
'==========================================================================================================================================
'
'ADD LINE TO TABLE_KEYDOWN SUB WITH THE FOLLOWING:    If HSEnterMode Then HighScoreProcessKey(keycode) AFTER THE STARTGAME ENTRY
'ADD And Not HSEnterMode=true TO IF KEYCODE=STARTGAMEKEY
'TO SHOW THE SCORE ON POST-IT ADD LINE AT RELEVENT LOCATION THAT HAS:  UpdatePostIt
'TO INITIATE ADDING INITIALS ADD LINE AT RELEVENT LOCATION THAT HAS:  HighScoreEntryInit()
'ADD THE FOLLOWING LINES TO TABLE_INIT TO SETUP POSTIT
'	if HSA1="" then HSA1=25
'	if HSA2="" then HSA2=25
'	if HSA3="" then HSA3=25
'	UpdatePostIt
'ADD HSA1, HSA2 AND HSA3 TO SAVE AND LOAD VALUES FOR TABLE
'ADD A TIMER NAMED HighScoreFlashTimer WITH INTERVAL 100 TO TABLE
'SET HSSSCOREX BELOW TO WHATEVER VARIABLE YOU USE FOR HIGH SCORE.
'IMPORT POST-IT IMAGES


Dim HSA1, HSA2, HSA3
Dim HSEnterMode, hsLetterFlash, hsEnteredDigits(3), hsCurrentDigit, hsCurrentLetter
Dim HSArray
Dim HSScoreM,HSScore100k, HSScore10k, HSScoreK, HSScore100, HSScore10, HSScore1, HSScorex	'Define 6 different score values for each reel to use
HSArray = Array("Postit0","postit1","postit2","postit3","postit4","postit5","postit6","postit7","postit8","postit9","postitBL","postitCM","Tape")
Const hsFlashDelay = 4

' ***********************************************************
'  HiScore DISPLAY
' ***********************************************************

Sub UpdatePostIt
	dim tempscore
	HSScorex = HighScore
	TempScore = HSScorex
	HSScore1 = 0
	HSScore10 = 0
	HSScore100 = 0
	HSScoreK = 0
	HSScore10k = 0
	HSScore100k = 0
	HSScoreM = 0
	if len(TempScore) > 0 Then
		HSScore1 = cint(right(Tempscore,1))
	end If
	if len(TempScore) > 1 Then
		TempScore = Left(TempScore,len(TempScore)-1)
		HSScore10 = cint(right(Tempscore,1))
	end If
	if len(TempScore) > 1 Then
		TempScore = Left(TempScore,len(TempScore)-1)
		HSScore100 = cint(right(Tempscore,1))
	end If
	if len(TempScore) > 1 Then
		TempScore = Left(TempScore,len(TempScore)-1)
		HSScoreK = cint(right(Tempscore,1))
	end If
	if len(TempScore) > 1 Then
		TempScore = Left(TempScore,len(TempScore)-1)
		HSScore10k = cint(right(Tempscore,1))
	end If
	if len(TempScore) > 1 Then
		TempScore = Left(TempScore,len(TempScore)-1)
		HSScore100k = cint(right(Tempscore,1))
	end If
	if len(TempScore) > 1 Then
		TempScore = Left(TempScore,len(TempScore)-1)
		HSScoreM = cint(right(Tempscore,1))
	end If
	Pscore6.image = HSArray(HSScoreM):If HSScorex<1000000 Then PScore6.image = HSArray(10)
	Pscore5.image = HSArray(HSScore100K):If HSScorex<100000 Then PScore5.image = HSArray(10)
	PScore4.image = HSArray(HSScore10K):If HSScorex<10000 Then PScore4.image = HSArray(10)
	PScore3.image = HSArray(HSScoreK):If HSScorex<1000 Then PScore3.image = HSArray(10)
	PScore2.image = HSArray(HSScore100):If HSScorex<100 Then PScore2.image = HSArray(10)
	PScore1.image = HSArray(HSScore10):If HSScorex<10 Then PScore1.image = HSArray(10)
	PScore0.image = HSArray(HSScore1):If HSScorex<1 Then PScore0.image = HSArray(10)
	if HSScorex<1000 then
		PComma.image = HSArray(10)
	else
		PComma.image = HSArray(11)
	end if
	if HSScorex<1000000 then
		PComma2.image = HSArray(10)
	else
		PComma2.image = HSArray(11)
	end if
'	if showhisc=1 and showhiscnames=1 then
'		for each objekt in hiscname:objekt.visible=1:next
		HSName1.image = ImgFromCode(HSA1, 1)
		HSName2.image = ImgFromCode(HSA2, 2)
		HSName3.image = ImgFromCode(HSA3, 3)
'	  else
'		for each objekt in hiscname:objekt.visible=0:next
'	end if
End Sub

Function ImgFromCode(code, digit)
	Dim Image
	if (HighScoreFlashTimer.Enabled = True and hsLetterFlash = 1 and digit = hsCurrentLetter) then
		Image = "postitBL"
	elseif (code + ASC("A") - 1) >= ASC("A") and (code + ASC("A") - 1) <= ASC("Z") then
		Image = "postit" & chr(code + ASC("A") - 1)
	elseif code = 27 Then
		Image = "PostitLT"
    elseif code = 0 Then
		image = "PostitSP"
    Else
      msgbox("Unknown display code: " & code)
	end if
	ImgFromCode = Image
End Function

Sub HighScoreEntryInit()
	HSA1=0:HSA2=0:HSA3=0
	HSEnterMode = True
	hsCurrentDigit = 0
	hsCurrentLetter = 1:HSA1=1
	HighScoreFlashTimer.Interval = 250
	HighScoreFlashTimer.Enabled = True
	hsLetterFlash = hsFlashDelay
End Sub

Sub HighScoreFlashTimer_Timer()
	hsLetterFlash = hsLetterFlash-1
	UpdatePostIt
	If hsLetterFlash=0 then 'switch back
		hsLetterFlash = hsFlashDelay
	end if
End Sub


' ***********************************************************
'  HiScore ENTER INITIALS
' ***********************************************************

Sub HighScoreProcessKey(keycode)
    If keycode = LeftFlipperKey Then
		hsLetterFlash = hsFlashDelay
		Select Case hsCurrentLetter
			Case 1:
				HSA1=HSA1-1:If HSA1=-1 Then HSA1=26 'no backspace on 1st digit
				UpdatePostIt
			Case 2:
				HSA2=HSA2-1:If HSA2=-1 Then HSA2=27
				UpdatePostIt
			Case 3:
				HSA3=HSA3-1:If HSA3=-1 Then HSA3=27
				UpdatePostIt
		 End Select
    End If

	If keycode = RightFlipperKey Then
		hsLetterFlash = hsFlashDelay
		Select Case hsCurrentLetter
			Case 1:
				HSA1=HSA1+1:If HSA1>26 Then HSA1=0
				UpdatePostIt
			Case 2:
				HSA2=HSA2+1:If HSA2>27 Then HSA2=0
				UpdatePostIt
			Case 3:
				HSA3=HSA3+1:If HSA3>27 Then HSA3=0
				UpdatePostIt
		 End Select
	End If

    If keycode = StartGameKey Then
		Select Case hsCurrentLetter
			Case 1:
				hsCurrentLetter=2 'ok to advance
				HSA2=HSA1 'start at same alphabet spot
'				EMReelHSName1.SetValue HSA1:EMReelHSName2.SetValue HSA2
			Case 2:
				If HSA2=27 Then 'bksp
					HSA2=0
					hsCurrentLetter=1
				Else
					hsCurrentLetter=3 'enter it
					HSA3=HSA2 'start at same alphabet spot
				End If
			Case 3:
				If HSA3=27 Then 'bksp
					HSA3=0
					hsCurrentLetter=2
				Else
					savehs 'enter it
					HighScoreFlashTimer.Enabled = False
					HSEnterMode = False
				End If
		End Select
		UpdatePostIt
    End If
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


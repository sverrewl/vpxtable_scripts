 '*********************************************************************
 '1994 Williams Demolition Man
 '*********************************************************************
 '*********************************************************************
 'Pinball Machine designed by Dennis Nordman
 '*********************************************************************
 '*********************************************************************
 'recreated for Visual Pinball by Knorr
 '*********************************************************************
 '*********************************************************************
 'Playfield and Plastics Redraw by Kiwi
 '*********************************************************************
 '*********************************************************************
 'I would like to give my sincere thanks to
 'Mfuegemann, Freneticamnesic, Toxie and the VPdevs, Clark Kent and
 'Gigalula for always being so friendly, helpful and motivating while
 'building this table.
 '*********************************************************************

 'V1.2
		'new ramps made by flupper (thanks!!!)
		'cleaned up script
		'fewer timers for better performance
		'small changes with the physics
		'claw animation improved by shoopity (thanks for all the explaination!)
		'bug fixes with pballs
		'added missing sounds
		'new enviroment, new lightning

 'V1.1
		'Added controller.vbs
		'Bug Fixes

' Thalamus 2018-07-20
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' No special SSF tweaks yet.


Option Explicit
 Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

 LoadVPM "01560000", "WPC.VBS", 3.36

 '********************
 'Standard definitions
 '********************

 Const cGameName = "dm_h6"
 Const UseSolenoids = 2
 Const UseLamps = 1
 Const SSolenoidOn = "SolOn"
 Const SSolenoidOff = "SolOff"
 Const SFlipperOn = "FlipperUp"
 Const SFlipperOff = "FlipperDown"
 Const SCoin = "Coin5"
 Set GiCallback2 = GetRef("UpdateGI")
 Set LampCallback = GetRef("UpdateMultipleLamps")
 Set MotorCallback = GetRef("RealTimeUpdates")
 BSize = 25.5


 Dim bsTrough, retinascan, bsTopPopper, BottomPopper, bsEject, Oldsmobileball, RedCar, GMUltraliteball, WhiteCar, clawmech, elevatormech, BallinClaw
 Dim ClawSpeed:ClawSpeed = 67



 '************
 ' Table init.
 '************

 Sub Table1_Init
     vpmInit Me
     With Controller
        .GameName = cGameName
          If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
         .Games(cGameName).Settings.Value("rol") = 0
         .HandleKeyboard = 0
         .ShowTitle = 0
         .ShowDMDOnly = 1
         .ShowFrame = 0
         .HandleMechanics = 0
         .Hidden = 0
         '.SetDisplayPosition 0, 0, GetPlayerHWnd 'uncomment this line if you don't see the DMD
          On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
         .Switch(22) = 1 'close coin door
         .Switch(24) = 1 'and keep it close
			PinMAMETimer.Interval = PinMAMEInterval
			PinMAMETimer.Enabled = true
			vpmNudge.TiltSwitch = 14
			vpmNudge.Sensitivity = 2
     End With


 '******
 'Trough
 '******
     Set bsTrough = New cvpmBallStack
     With bsTrough
         .InitSw 0, 31, 32, 33, 34, 35, 0, 0
         .InitKick BallRelease, 90, 10
		 .InitExitSnd SoundFX("BallRelease",DOFContactors), SoundFX("fx_Solenoid",DOFContactors)
         .Balls = 5
'         .IsTrough = 1
     End With

 '***********
 'Retina Scan
 '***********

Set retinascan = New cvpmCaptiveBall
	With retinascan
		.InitCaptive RetinaTrigger, RetinaWall, RetinaKicker, 345
		.ForceTrans = .9
		.MinForce = 2
		.CreateEvents "retinascan"
		.Start
	End With

 '***************
 'Claw & Elevator
 '***************

set clawmech = new cvpmmech
 with clawmech
  .mtype =vpmmechtwodirsol + vpmmechstopend + vpmmechlinear
  .sol1  =19
  .sol2  =20
  .length=Clawspeed
  .steps = 147
  .addsw 25,0,2
  .addsw 26,140,147
  .callback= getRef("UpdateClaw")
  .start
 end with

Set elevatormech = new cvpmMech
	with elevatormech
		.mtype = vpmMechOneSol + vpmMechReverse + vpmMechLinear
		.Sol1 = 18
		.Length = 15
		.steps = 70
		.addsw 67,0,2
		.addsw 74,65,70
		.Callback = getRef("UpdateElevator")
		.Start
	End with


 '***********
 'Top Popper
 '***********

Set bsTopPopper = New cvpmBallStack
	with bsTopPopper
	.InitSaucer sw73, 73, 182, 18
	.InitExitSnd SoundFX("EjectKick",DOFContactors), SoundFX("fx_Solenoid",DOFContactors)
	.KickForceVar = 1
	.KickAngleVar = 0.5
	End With


 '*************
 'Bottom Popper
 '*************

Set BottomPopper = New cvpmBallStack
	With BottomPopper
		.InitSw 0, 76, 0, 0, 0, 0, 0, 0
		.InitKick sw76, 58, 73
		.KickZ = 60
		.InitExitSnd SoundFX("BottomPopper",DOFContactors), SoundFX("fx_Solenoid",DOFContactors)
	End With


 '*****
 'Eject
 '*****

Set bsEject = New cvpmBallStack
	With bsEject
	.InitSaucer sw66, 66, 180, 2
	.InitExitSnd SoundFX("EjectKick",DOFContactors), SoundFX("fx_Solenoid",DOFContactors)
	.KickForceVar = 1
	.KickAngleVar = 0.5
End With

 '************
 'CarCapative
 '************

Set RedCar = New cvpmCaptiveBall
	With RedCar
		.InitCaptive OldsmobileTrigger, OldsmobileWall, OldsmobileKicker, 345
		.ForceTrans = .9
		.MinForce = 3.5
'		.RestSwitch = 71
		.CreateEvents "RedCar"
		.Start
	End With


DiverterOn.isdropped = 1
DiverterOff.isdropped = 0


 '************
 'FS Setting
 '************


 If table1.ShowDT = False then
	LeftSideRail.WidthTop = 0
	LeftSideRail.WidthBottom = 0
	RightSideRail.WidthTop = 0
	RightSideRail.WidthBottom = 0
	Korpus.Size_Y = 1.7
'	Korpus.visible = False
 End if
End Sub


 '******
 'Cars Animation
 '******

'Oldsmobile

	Set Oldsmobileball = CapKicker1.CreateBall
	Oldsmobileball.Visible = False
	CapKicker1.Kick 0,0
	CapKicker1.Enabled = false
	Controller.Switch(71) = 1

Sub OldsmobileTimer_Timer()
	Oldsmobile.x = Oldsmobileball.x:Oldsmobile.y = Oldsmobileball.y
End Sub

'GM Ultralite

	Set GMUltraliteball = CapKicker2.CreateBall
	GMUltraliteball.Visible = False
	CapKicker2.Kick 0,0
	CapKicker2.Enabled = false
	Controller.Switch(72) = 1

Sub GMUltraliteTimer_Timer()
	GMUltralite.x = GMUltraliteball.x:GMUltralite.y = GMUltraliteball.y
End Sub



Sub sw72_Hit:Controller.Switch(72) = 0:End Sub		'GMUltralite Restswitch
Sub sw72_UnHit:Controller.Switch(72) = 1: End Sub

Sub sw71_Hit:Controller.Switch(71) = 0:End Sub		'Oldsmobil Restswitch
Sub sw71_UnHit:Controller.Switch(71) = 1: End Sub

 '*********
 'Solenoids
 '*********

 SolCallback(1) = "SolRelease"
 SolCallback(2) = "BottomPopper.SolOut"
 SolCallback(3) = "AutoPlunge"
 SolCallback(4) = "bsTopPopper.SolOut"
 SolCallback(15) = "DiverterRight"
 SolCallback(7) = "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
 SolCallBack(14) = "bsEject.SolOut"
' SolCallback(18) = "ElevatorMotor"
 SolCallBack(19) = "SolClawMotorLeft"
' SolCallBack(20) = "SolClawMotorRight"
 SolCallback(33) = "ClawMagnetOn"




' SolCallback(9) = "LeftSlingShot"
' SolCallback(10) = "RightSlingShot"
' SolCallback(11) = "LeftJetBumper"
' SolCallBack(12) = "TopSlingShot"
' SolCallback(13) = "RightJetBumper"


'************
' BallRelease
'************


 Sub SolRelease(Enabled)
     If Enabled And bsTrough.Balls > 0 Then
         vpmTimer.PulseSw 36
         bsTrough.ExitSol_On
     End If
 End Sub


'*****************************************
' Claw and Elevator Animation + Ballupdate
'*****************************************



'Sub UpdateClaw(aNewPos,aSpeed,aLastPos)
'	PlaySound SoundFX("ClawMotor",DOFGear)
'	Claw.RotY = -107 +clawmech.position
'	BallP1.RotY = -107 +clawmech.position
'End Sub


Sub UpdateClaw(aNewPos,aSpeed,aLastPos)
	ClawTimer.Enabled = 1
End Sub

Sub SolClawMotorLeft(Enabled)
	If enabled then
		PlaySound SoundFX("Motor",DofGear)
	Else
		StopSound SoundFX("Motor",DofGear)
	End if
End Sub



Sub ClawTimer_Timer()
If clawmech.position-107 > Claw.Roty Then
Claw.RotY = Claw.RotY + (8/ClawSpeed)
If Claw.RotY > clawmech.position-107 Then me.enabled = 0
End If
If clawmech.position-107 < Claw.Roty Then
Claw.RotY = Claw.RotY - (8/ClawSpeed)
If Claw.RotY < clawmech.position-107 Then me.enabled = 0
End If
BallP1.RotY = Claw.RotY
End Sub

Sub UpdateElevator(aNewPos,aSpeed,aLastPos)
	PlaySound SoundFX("ElevatorMotor",DOFGear)
	Elevator.TransY = elevatormech.Position
	BallP.TransY = elevatormech.Position
End Sub

Sub sw74_Hit
	controller.Switch(74) = True
End Sub

Sub ElevatorKicker_Hit()
	me.DestroyBall
	BallP.Visible = True
	ClawKicker1.Enabled = True
	ClawKicker2.Enabled = True
	ClawKicker3.Enabled = True
	ClawKicker4.Enabled = True
	ClawKicker5.Enabled = True
	BallinClaw = True
End Sub


Sub ClawMagnetOn(Enabled)
	If Enabled And BallinClaw = True Then
		vpmTimer.AddTimer 700, "BallP.Visible = False'"
		vpmTimer.AddTImer 700, "BallP1.Visible = True'"
	Else
		ClawOff
	End if
End Sub

Sub ClawOff()
	BallP1.Visible = False
	If BallinClaw = True then
		if Claw.RotY >= 11 Then ClawKicker5.CreateSizedBall(25.5): ClawKicker5.Kick 185, 1
		if Claw.RotY <= 10.99 And Claw.RotY >= -9.99 Then ClawKicker4.CreateSizedBall(25.5): ClawKicker4.Kick 180, 1
		if Claw.RotY <= -10 And Claw.RotY >= -28.99 Then ClawKicker3.CreateSizedBall(25.5): ClawKicker3.Kick 161, 1
		if Claw.RotY <= -29 And Claw.RotY >= -52.99 Then ClawKicker2.CreateSizedBall(25.5): ClawKicker2.Kick 143, 1
		if Claw.Roty <= -53 And Claw.RotY >= -95.99 Then ClawKicker1.CreateSizedBall(25.5): ClawKicker1.Enabled = False: ClawKicker1.Kick 0, 1
	Else
		ClawKicker1.Enabled = False
		ClawKicker2.Enabled = False
		ClawKicker3.Enabled = False
		ClawKicker4.Enabled = False
		ClawKicker5.Enabled = False
		BallinClaw = False
	End if
End Sub







'*****************
' EyeBallPrimitive
'*****************

Dim EyeBallM


Sub EyeballTrigger_Hit()
	Set EyeballM = ActiveBall
	EyeballTimer.Enabled = True
End Sub


Sub EyeBallTimer_Timer()
	EyeballP.x = Eyeballm.x: EyeballP.y = EyeballM.y
	If EyeBallm.VelY <= -1 Then EyeballP.Roty = EyeballP.Roty +30
	If EyeBallm.VelY >= 1 Then EyeballP.Roty = EyeballP.Roty +30
'	If EyeBallm.VelY = 0 Then EyeballP.Roty = EyeballP.Roty 0
End Sub






'**********
' Vuks
'**********


Dim aBall

Sub sw73_Hit()
    PlaySound "HeadquarterHit2", 0, 1, pan(ActiveBall)
    Set aBall = ActiveBall:Me.TimerEnabled = 1
    bsTopPopper.AddBall Me
End Sub

Sub sw73_Timer
    Do While aBall.Z > 0
        aBall.Z = aBall.Z -5
        Exit Sub
    Loop
    Me.TimerEnabled = 0
End Sub


Sub BottomPopperHole_Hit()
	RandomSoundHole
	BottomPopper.AddBall Me
End Sub

Sub BottomPopperHole1_Hit()
	RandomSoundHole
	BottomPopper.AddBall Me
End Sub

Sub ClawRampKicker_Hit()
	RandomSoundHole
	vpmTimer.PulseSw 81
	BottomPopper.AddBall Me
End Sub

Sub sw66_Hit()
	PlaySound ""
	bsEject.AddBall Me
End Sub

Sub Drain_Hit
	bsTrough.AddBall Me
	PlaySound "drain",0,0.5,0
End Sub


'**********
' Keys
'**********

 Sub table1_KeyDown(ByVal Keycode)
	If KeyCode=MechanicalTilt Then
		vpmTimer.PulseSw vpmNudge.TiltSwitch
		Exit Sub
	End if

    If keycode = PlungerKey Then Controller.Switch(11) = 1: Controller.Switch(12) = 1
'    If keycode = LeftMagnaSave Or RightMagnaSave Then Controller.Switch(12) = 1
    If keycode = keyFront Then Controller.Switch(23) = 1
    If vpmKeyDown(keycode) Then Exit Sub
 End Sub

 Sub table1_KeyUp(ByVal Keycode)
     If vpmKeyUp(keycode) Then Exit Sub
     If keycode = PlungerKey Then Controller.Switch(11) = 0: Controller.Switch(12) = 0
'     If keycode = LeftMagnaSave Or RightMagnaSave Then Controller.Switch(12) = 0
     If keycode = keyFront Then Controller.Switch(23) = 0
 End Sub




 '**************
 ' Flipper Subs
 '**************

 SolCallback(sLRFlipper) = "SolRFlipper"
 SolCallback(sLLFlipper) = "SolLFlipper"

 Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFX("FlipperUpLeftBoth",DOFContactors):LeftFlipper.RotateToEnd:LeftFlipper1.RotateToEnd
     Else
         PlaySound SoundFX("FlipperDown",DOFContactors):LeftFlipper.RotateToStart:LeftFlipper1.RotateToStart
     End If
 End Sub

 Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFX("FlipperUpRight",DOFContactors):RightFlipper.RotateToEnd
     Else
         PlaySound SoundFX("FlipperDown",DOFContactors):RightFlipper.RotateToStart
     End If
 End Sub


 '******
 'Plunger
 '******

Dim AP

Sub AutoPlunge(Enabled)
	if enabled then
		AP = True
		Kicker1.Kick 1,50
		PlaySound SoundFX("Plunger",DOFContactors)
	End if
End Sub

	Sub PlungerPTimer()
	if AP = True and PlungerP.TransZ < 45 then PlungerP.TransZ = PlungerP.TransZ +10
	if AP = False and PlungerP.TransZ > 0 then PlungerP.TransZ = PlungerP.TransZ -10
	if PlungerP.TransZ >= 45 then AP = False
End Sub


 '*********
 ' Switches
 '*********


Sub sw27_Hit:Controller.Switch(27) = 1:sw27wire.RotX = 15:PlaySound "metalhit_thin":End Sub		'shooterlane
Sub sw27_UnHit:Controller.Switch(27) = 0:sw27wire.RotX = 0: End Sub
Sub sw15_Hit:Controller.Switch(15) = 1:sw15wire.RotX = 15:PlaySound "metalhit_thin":End Sub 	'left outlane
Sub sw15_UnHit:Controller.Switch(15) = 0:sw15wire.RotX = 0:End Sub
Sub sw16_Hit:Controller.Switch(16) = 1:sw16wire.RotX = 15:PlaySound "metalhit_thin":End Sub 	'left inlane
Sub sw16_UnHit:Controller.Switch(16) = 0:sw16wire.RotX = 0:End Sub
Sub sw17_Hit:Controller.Switch(17) = 1:sw17wire.RotX = 15:PlaySound "metalhit_thin":End Sub 	'right inlane
Sub sw17_UnHit:Controller.Switch(17) = 0:sw17wire.RotX = 0:End Sub
Sub sw18_Hit:Controller.Switch(18) = 1:sw18wire.RotX = 15:PlaySound "metalhit_thin":End Sub 	'right outlane
Sub sw18_UnHit:Controller.Switch(18) = 0:sw18wire.RotX = 0:End Sub
Sub sw63_Hit:Controller.Switch(63) = 1:sw63wire.RotX = 15:PlaySound "metalhit_thin":End Sub 	'leftrollover
Sub sw63_UnHit:Controller.Switch(63) = 0:sw63wire.RotX = 0:End Sub
Sub sw64_Hit:Controller.Switch(64) = 1:sw64wire.RotX = 15:PlaySound "metalhit_thin":End Sub 	'centerrollover
Sub sw64_UnHit:Controller.Switch(64) = 0:sw64wire.RotX = 0:End Sub
Sub sw65_Hit:Controller.Switch(65) = 1:sw65wire.RotX = 15:PlaySound "metalhit_thin":End Sub 	'rightrollover
Sub sw65_UnHit:Controller.Switch(65) = 0:sw65wire.RotX = 0:End Sub
Sub sw48_Hit:Controller.Switch(48) = 1:sw48wire.RotX = 15:PlaySound "metalhit_thin":End Sub 	'Right freeway
Sub sw48_UnHit:Controller.Switch(48) = 0:sw48wire.RotX = 0:End Sub
Sub sw55_Hit:Controller.Switch(55) = 1:sw55wire.RotX = 15:PlaySound "metalhit_thin":End Sub 	'leftloop
Sub sw55_UnHit:Controller.Switch(55) = 0:sw55wire.RotX = 0:End Sub
Sub sw86_Hit:Controller.Switch(86) = 1:End Sub 		'upperleftflippergate
Sub sw86_UnHit:Controller.Switch(86) = 0:End Sub
Sub sw46_Hit:Controller.Switch(46) = 1:End Sub 		'rightrampenter
Sub sw46_UnHit:Controller.Switch(46) = 0:End Sub
Sub sw47_Hit:Controller.Switch(47) = 1:End Sub 		'rightrampexit
Sub sw47_UnHit:Controller.Switch(47) = 0:End Sub
Sub sw75_Hit:Controller.Switch(75) = 1: End Sub 		'elevatorramp
Sub sw75_UnHit:Controller.Switch(75) = 0:End Sub
Sub sw53_Hit:Controller.Switch(53) = 1:End Sub 		'centerramp
Sub sw53_UnHit:Controller.Switch(53) = 0:End Sub
Sub sw51_Hit:Controller.Switch(51) = 1:End Sub 		'leftrampenter
Sub sw51_UnHit:Controller.Switch(51) = 0:End Sub
Sub sw52_Hit:Controller.Switch(52) = 1:End Sub 		'leftrampexit
Sub sw52_UnHit:Controller.Switch(52) = 0:End Sub
Sub sw61_Hit:Controller.Switch(61) = 1:End Sub 		'siderampenter
Sub sw61_UnHit:Controller.Switch(61) = 0:End Sub
Sub sw62_Hit:Controller.Switch(62) = 1:End Sub 		'siderampexit
Sub sw62_UnHit:Controller.Switch(62) = 0:End Sub


Sub sw82_Hit:Controller.Switch(82) = 1:sw82wire.RotX = 75:PlaySound "metalhit_thin":End Sub		'Claw SuperJets
Sub sw82_UnHit:Controller.Switch(82) = 0:sw82wire.RotX = 90: End Sub

Sub sw83_Hit:Controller.Switch(83) = 1:sw83wire.RotX = 70:PlaySound "metalhit_thin":End Sub		'Claw PrisonBreak
Sub sw83_UnHit:Controller.Switch(83) = 0:sw83wire.RotX = 90: End Sub

Sub sw84_Hit:Controller.Switch(84) = 1:sw84wire.RotX = 75:PlaySound "metalhit_thin":End Sub		'Claw Freeze
Sub sw84_UnHit:Controller.Switch(84) = 0:sw84wire.RotX = 90: End Sub

Sub sw85_Hit:Controller.Switch(85) = 1:sw85wire.RotX = 75:PlaySound "metalhit_thin":End Sub		'Claw ACMAG
Sub sw85_UnHit:Controller.Switch(85) = 0:sw85wire.RotX = 90: End Sub




Sub GateTimer()
	upperleftflipperspinnerp.RotX = upperleftflipperspinner.currentangle +90
	rightrampenterspinnerp.RotX = rightrampenterspinner.currentangle +95
	rightrampexitspinnerp.RotX = rightrampexitspinner.currentangle +90
	centerrampspinnerp.RotX = centerrampspinner.currentangle +90
	leftrampenterspinnerp.RotX = leftrampenterspinner.currentangle +90
	leftrampexitspinnerp.RotX = leftrampexitspinner.currentangle +90
	siderampenterspinnerp.RotX = siderampenterspinner.currentangle +95
	siderampexitspinnerp.RotZ = siderampexitspinner.currentangle
	retinagatep.RotX = retinagate.currentangle +90
End Sub



 '*********
 ' Ramps
 '*********


	'RampSounds

Dim SoundBall


Sub WireStartAcmag_Hit()
 Set SoundBall = Activeball 'Ball-assignment
 playsound "WireRamp",-1,1,pan(ActiveBall)                    '-1 = looping, Panning = -1 bis 1, je nach Position
End Sub


Sub WireStartLockFreeze_Hit()
 Set SoundBall = Activeball 'Ball-assignment
 playsound "WireRamp",-1,1,pan(ActiveBall)                    '-1 = looping, Panning = -1 bis 1, je nach Position
End Sub


Sub WireStartSuperJets_Hit()
 Set SoundBall = Activeball 'Ball-assignment
 playsound "WireRamp",-1,1,pan(ActiveBall)                   '-1 = looping, Panning = -1 bis 1, je nach Position
End Sub


Sub WireStartLeftRamp_Hit()
 Set SoundBall = Activeball 'Ball-assignment
 playsound "WireRamp",-1,1,pan(ActiveBall)                    '-1 = looping, Panning = -1 bis 1, je nach Position
End Sub


Sub WireStartRightRamp_Hit()
 Set SoundBall = Activeball 'Ball-assignment
 playsound "WireRamp",-1,1,pan(ActiveBall)                    '-1 = looping, Panning = -1 bis 1, je nach Position
End Sub


Sub WireStartMiddleRamp_Hit()
 Set SoundBall = Activeball 'Ball-assignment
 playsound "WireRamp",-1,1,pan(ActiveBall)                    '-1 = looping, Panning = -1 bis 1, je nach Position
End Sub


Sub WireStartBottomPopper_Hit()
 Set SoundBall = Activeball 'Ball-assignment
 playsound "WireRamp",-1,1,pan(ActiveBall)                    '-1 = looping, Panning = -1 bis 1, je nach Position
End Sub


Sub WireStopLeft_Hit(): StopSound "WireRamp": PlaySound "BallDrop": End Sub

Sub WireStopRight_Hit(): StopSound "WireRamp": PlaySound "BallDrop": End Sub

Sub WireStopRetina_Hit(): StopSound "WireRamp": PlaySound "BallDrop": End Sub

Sub WireStopSuperJets_Hit(): StopSound "WireRamp": PlaySound "BallDrop": End Sub



 '********
 'Diverter
 '********

Sub DiverterRight(Enabled)
	If Enabled then
        DiverterR.rotatetoend
		DiverterOff.isdropped = 1
		DiverterOn.isdropped = 0
		PlaySound SoundFX ("DiverterRight",DOFContactors)
		Else
		DiverterOff.isdropped = 0
		DiverterOn.isdropped = 1
		DiverterR.rotatetostart
		Playsound SoundFX ("DiverterRight",DOFContactors)
	End if
End Sub

 '***************
 ' StandupTargets
 '***************

Sub Standup77_Hit: vpmTimer.pulseSw 77:Standup77p.RotX=Standup77p.RotX +15:Playsound SoundFX("target",DOFContactors):Me.TimerEnabled = 1: End Sub
Sub Standup77_Timer:Standup77p.RotX=Standup77p.RotX -15:Me.TimerEnabled = 0: End Sub

Sub Standup87_Hit: vpmTimer.pulseSw 87:Standup87p.RotX=Standup87p.RotX +15:Playsound SoundFX("target",DOFContactors):Me.TimerEnabled = 1: End Sub
Sub Standup87_Timer:Standup87p.RotX=Standup87p.RotX -15:Me.TimerEnabled = 0: End Sub


Sub Standup78_Hit: vpmTimer.pulseSw 78:Standup78p.RotZ=Standup78p.RotZ +15:Playsound SoundFX("target",DOFContactors):Me.TimerEnabled = 1: End Sub
Sub Standup78_Timer:Standup78p.RotZ=Standup78p.RotZ -15:Me.TimerEnabled = 0: End Sub

Sub Standup38_Hit: vpmTimer.pulseSw 38:Standup38p.RotZ=Standup38p.RotZ +15:Playsound SoundFX("target",DOFContactors):Me.TimerEnabled = 1: End Sub
Sub Standup38_Timer:Standup38p.RotZ=Standup38p.RotZ -15:Me.TimerEnabled = 0: End Sub

Sub Standup57_Hit: vpmTimer.pulseSw 57:Standup57p.RotZ=Standup57p.RotZ +15:Playsound SoundFX("target",DOFContactors):Me.TimerEnabled = 1: End Sub
Sub Standup57_Timer:Standup57p.RotZ=Standup57p.RotZ -15:Me.TimerEnabled = 0: End Sub

Sub Standup58_Hit: vpmTimer.pulseSw 58:Standup58p.RotZ=Standup58p.RotZ +15:Playsound SoundFX("target",DOFContactors):Me.TimerEnabled = 1: End Sub
Sub Standup58_Timer:Standup58p.RotZ=Standup58p.RotZ -15:Me.TimerEnabled = 0: End Sub

Sub Standup56_Hit: vpmTimer.pulseSw 56:Standup56p.RotZ=Standup56p.RotZ +15:Playsound SoundFX("target",DOFContactors):Me.TimerEnabled = 1: End Sub
Sub Standup56_Timer:Standup56p.RotZ=Standup56p.RotZ -15:Me.TimerEnabled = 0: End Sub






 '*********
 ' Bumper
 '*********

Sub leftjetbumper_Hit:vpmTimer.PulseSw 43:Playsound SoundFX ("bumperleft",DOFContactors):Me.TimerEnabled = 1: End Sub
Sub Leftjetbumper_Timer:Me.TimerEnabled = 0:End Sub
Sub rightjetbumper_Hit:vpmTimer.PulseSw 45:Playsound SoundFX ("bumperright",DOFContactors):Me.TimerEnabled = 1: End Sub
Sub rightjetbumper_Timer:Me.TimerEnabled = 0: End Sub



'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, LStep, TStep

Sub RightSlingShot_Slingshot
    PlaySound SoundFX ("right_slingshot",DOFContactors), 0, 1, 0.05, 0.05
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
    vpmTimer.PulseSw 42
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub



Sub LeftSlingShot_Slingshot
    PlaySound SoundFX ("left_slingshot",DOFContactors),0,1,-0.05,0.05
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
    vpmTimer.PulseSw 41
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub



Sub TopSlingShot_Slingshot
	PlaySound SoundFX ("right_slingshot",DOFContactors),0,1,0,0.05
	TSling.Visible = 0
	TSling1.Visible = 1
	Sling3.TransZ = -20
	TStep = 0
	TopSlingshot.TimerEnabled = 1
	vpmTimer.PulseSw 44
End Sub

Sub TopSlingShot_Timer
	Select Case TStep
		Case 3:TSling1.Visible = 0:TSling2.Visible = 1:sling3.TransZ = -10
		Case 4:TSling2.Visible = 0:TSling.Visible = 1:Sling3.TransZ = 0:TopSlingShot.TimerEnabled = 0
	End Select
	TStep = TStep +1
End Sub

Sub UpperRebound_Slingshot
	vpmTimer.PulseSw 54
End Sub

Sub LowerRebound_Slingshot
	vpmTimer.PulseSw 88
End Sub


'*****************************************
 '  JP's Fading Lamps 3.4 VP9 Fading only
 '      Based on PD's Fading Lights
 ' SetLamp 0 is Off
 ' SetLamp 1 is On
 ' LampState(x) current state
 '****************************************

InitLamps

Sub InitLamps()
    Set Lights(11) = l11a
    Set Lights(12) = l12
    Set Lights(13) = l13
    Set Lights(14) = l14
    Set Lights(15) = l15
    Set Lights(16) = l16
    Set Lights(17) = l17
    Set Lights(18) = l18
    Set Lights(21) = l21
    Set Lights(22) = l22
    Set Lights(23) = l23
    Set Lights(24) = l24
    Set Lights(25) = l25
    Set Lights(26) = l26
    Set Lights(27) = l27
    Set Lights(28) = l28
    Set Lights(31) = l31
    Set Lights(32) = l32
    Set Lights(33) = l33
    Set Lights(34) = l34
    Set Lights(35) = l35
    Set Lights(36) = l36
    Set Lights(37) = l37
    Set Lights(38) = l38
    Set Lights(41) = l41
    Set Lights(42) = l42
    Set Lights(43) = l43
    Set Lights(44) = l44
    Set Lights(45) = l45
    Set Lights(46) = l46
    Set Lights(47) = l47
    Set Lights(48) = l48
    Set Lights(51) = l51
    Set Lights(52) = l52
    Set Lights(53) = l53
    Set Lights(54) = l54
    Set Lights(55) = l55
    Set Lights(56) = l56
    Set Lights(57) = l57
    Set Lights(58) = l58
    Set Lights(61) = l61
    Set Lights(62) = l62
    Set Lights(63) = l63
    Set Lights(64) = l64
    Set Lights(65) = l65
    Set Lights(66) = l66
    Set Lights(67) = l67
    Set Lights(68) = l68
    Set Lights(71) = l71
    Set Lights(72) = l72
    Set Lights(73) = l73
    Set Lights(76) = l76
    Set Lights(77) = l77
    Set Lights(78) = l78
    Set Lights(81) = l81
    Set Lights(82) = l82
    Set Lights(83) = l83
    Set Lights(84) = l84
    Set Lights(85) = l85
End Sub


  Sub UpdateMultipleLamps()
		If l11a.state = 1 then l11b.state = 1: else l11b.state = 0
		If l36.state = 1 then targetcars.image = "target1on": Else targetcars.image = "target1"
		If l61.state = 1 then bulbcover61.visible = false: bulbcover61a.visible = true:f61.visible = true: else bulbcover61.visible = true: bulbcover61a.visible = false:f61.visible = false
		If l62.state = 1 then bulbcover62.visible = false: bulbcover62a.visible = true:f62.visible = true: else bulbcover62.visible = true: bulbcover62a.visible = false:f62.visible = false
		If l63.state = 1 then bulbcover63.visible = false: bulbcover63a.visible = true:f63.visible = true: else bulbcover63.visible = true: bulbcover63a.visible = false:f63.visible = false
		If l64.state = 1 then bulbcover64.visible = false: bulbcover64a.visible = true:f64.visible = true: else bulbcover64.visible = true: bulbcover64a.visible = false:f64.visible = false
		If l65.state = 1 then bulbcover65.visible = false: bulbcover65a.visible = true:f65.visible = true: else bulbcover65.visible = true: bulbcover65a.visible = false:f65.visible = false
		If l71.state = 1 then bulbcoverred.visible = false: bulbcoverreda.visible = true:f71.visible = true: else bulbcoverred.visible = true: bulbcoverreda.visible = false:f71.visible = false
		If l72.state = 1 then bulbcoveryellow.visible = false: bulbcoveryellowa.visible = true: f72.visible = true: else bulbcoveryellow.visible = true: bulbcoveryellowa.visible = false: f72.visible = false
		If l73.state = 1 then bulbcoverblue.visible = false: bulbcoverbluea.visible = true: f73.visible = true: else bulbcoverblue.visible = true: bulbcoverbluea.visible = false: f73.visible = false
		If l78.state = 1 then targetretina.image = "target2On": Else targetretina.image = "target2"
		If l82.state = 1 then l82a.state = 1: else l82a.state = 0
		If l83.state = 1 then l83a.state = 1: else l83a.state = 0
		If l81.state = 1 OR l82.state = 1 OR l83.state = 1 OR l82a.state = 1 OR l83a.state = 1 Then Centerlight.state = 1: Else Centerlight.state = 0
End Sub


'RainbowLight

Dim RGBStep, RGBFactor, Red, Green, Blue

RGBStep = 0
RGBFactor = 1
Red = 255
Green = 0
Blue = 0

Sub RGBTimer_timer 'rainbow light color changing
    Select Case RGBStep
        Case 0 'Green
            Green = Green + RGBFactor
            If Green > 255 then
                Green = 255
                RGBStep = 1
            End If
        Case 1 'Red
            Red = Red - RGBFactor
            If Red < 0 then
                Red = 0
                RGBStep = 2
            End If
        Case 2 'Blue
            Blue = Blue + RGBFactor
            If Blue > 255 then
                Blue = 255
                RGBStep = 3
            End If
        Case 3 'Green
            Green = Green - RGBFactor
            If Green < 0 then
                Green = 0
                RGBStep = 4
            End If
        Case 4 'Red
            Red = Red + RGBFactor
            If Red > 255 then
                Red = 255
                RGBStep = 5
            End If
        Case 5 'Blue
            Blue = Blue - RGBFactor
            If Blue < 0 then
                Blue = 0
                RGBStep = 0
            End If
    End Select
    'Light1.color = RGB(Red\10, Green\10, Blue\10)
    l53.colorfull = RGB(Red, Green, Blue)
    l77.colorfull = RGB(Red, Green, Blue)
    l76.colorfull = RGB(Red, Green, Blue)
    l32.colorfull = RGB(Red, Green, Blue)
    l78.colorfull = RGB(Red, Green, Blue)
    l37.colorfull = RGB(Red, Green, Blue)
    l1.colorfull = RGB(Red, Green, Blue)
'    light2.color = RGB(Red, Green, Blue)
'    textbox1.text = Red
'    textbox2.text = Green
'    textbox3.text = Blue
End Sub


 '*********
 'Flasher
 '*********

 SolCallback(17) = "Multi117" 			'ClawFlasher
 SolCallback(21) = "Multi121" 			'JetFlasher + Insert
 SolCallback(22) = "Multi122" 			'SideRampFlasher
 SolCallback(23) = "Multi123" 			'LeftRampUpperFlasher
 SolCallback(24) = "Multi124" 			'LeftRampLowerFlasher
 SolCallback(25) = "Multi125" 			'CarChaseCenterFlasher
 SolCallback(26) = "Multi126" 			'CarChaseLowerFlasher
 SolCallback(27) = "Multi127" 			'RightRampFlasher
 SolCallback(28) = "Multi128" 			'EjectFlasher

 SolCallback(51) = "Multi137" 			'CarChaseUpperFlasher
 SolCallback(52) = "Multi138" 			'LowerReboundFlasher
 SolCallback(53) = "Multi139" 			'EyeBallFlasher
 SolCallback(54) = "Multi140" 			'CenterRampFlasher
 SolCallback(55) = "Multi141" 			'Elevator2Flasher
 SolCallback(56) = "Multi142" 			'Elevator1Flasher
 SolCallback(57) = "Multi143" 			'DiverterFlasher
 SolCallback(58) = "Multi144" 			'RightRampUpperFlasher


Sub Multi117(Enabled)
	If Enabled Then
	f17.Visible = 1
	l117.State = 1
	Else
	f17.Visible = 0
	l117.State = 0
	End if
End Sub

Sub Multi121(Enabled)
	If Enabled Then
	l121.State = 1
	l121a.State = 1
	Else
	l121.State = 0
	l121a.State = 0
	End if
End Sub

Sub Multi122(Enabled)
	If Enabled Then
	l122.State = 1
	Else
	l122.State = 0
	End if
End Sub

Sub Multi124(Enabled)
	If Enabled Then
	l124.State = 1
	Else
	l124.State = 0
	End if
End Sub

Sub Multi125(Enabled)
	If Enabled Then
	l125.State = 1
	Else
	l125.State = 0
	End if
End Sub

Sub Multi126(Enabled)
	If Enabled Then
	l126.State = 1
	l126a.State = 1
	Else
	l126.State = 0
	l126a.State = 0
	End if
End Sub

Sub Multi127(Enabled)
	If Enabled Then
	l127.State = 1
	Else
	l127.State = 0
	End if
End Sub

Sub Multi128(Enabled)
	If Enabled Then
	l128.State = 1
	Else
	l128.State = 0
	End if
End Sub

Sub Multi137(Enabled)
	If Enabled Then
	l137.State = 1
	Else
	l137.State = 0
	End if
End Sub

Sub Multi138(Enabled)
	If Enabled Then
	l138.State = 1
	Else
	l138.State = 0
	End if
End Sub


Sub Multi139(Enabled)
	If Enabled Then
	l139.State = 1
	Else
	l139.State = 0
	End if
End Sub

Sub Multi140(Enabled)
	If Enabled Then
	l140.State = 1
	Else
	l140.State = 0
	End if
End Sub

Sub Multi141(Enabled)
	If Enabled Then
	f41.Visible = 1
	f41a.Visible = 1
	l141.State = 1
	l141a.State = 1
	Else
	f41.Visible = 0
	f41a.Visible = 0
	l141.State = 0
	l141a.State = 0
	End if
End Sub

Sub Multi142(Enabled)
	If Enabled Then
	f42.Visible = 1
	l142.State = 1
	Else
	f42.Visible = 0
	l142.State = 0
	End if
End Sub

Sub Multi143(Enabled)
	If Enabled Then
	l143.State = 1
	Flashercap2.image = "flashercapredON"
	RedS.State = 1
	Else
	l143.State = 0
	Flashercap2.image = "flashercapred"
	RedS.State = 0
	End if
End Sub


Sub Multi144(Enabled)
	If Enabled Then
	l144.State = 1
	Flashercap1.image = "flashercapredON"
	RedS.State = 1
	Else
	l144.State = 0
	Flashercap1.image = "flashercapred"
	RedS.State = 0
	End if
End Sub

Sub Multi123(Enabled)
	If Enabled Then
	l123.State = 1
	Flashercap3.image = "flashercapredON"
	RedS.State = 1
	Else
	l123.State = 0
	Flashercap3.image = "flashercapred"
	RedS.State = 0
	End if
End Sub


 '*********
 'Update GI
 '*********



Dim xx
Dim gistep
gistep = 1 / 8

Sub UpdateGI(no, step)
    If step = 0 OR step = 7 then exit sub
    Select Case no

		'Upper Right String
        Case 1
            For each xx in GIString2:xx.IntensityScale = gistep * step:next

		'Upper Left String
		Case 2
            For each xx in GIString3:xx.IntensityScale = gistep * step:next


		'Lower Right String
		Case 3
            For each xx in GIString4:xx.IntensityScale = gistep * step:next
					if step = 1 then Table1.ColorGradeImage = "-70"
					if step = 2 then Table1.ColorGradeImage = "-60"
					if step = 3 then Table1.ColorGradeImage = "-50"
					if step = 4 then Table1.ColorGradeImage = "-40"
					if step = 5 then Table1.ColorGradeImage = "-30"
					if step = 6 then Table1.ColorGradeImage = "-20"
					if step = 7 then Table1.ColorGradeImage = "-10"
					if step = 8 then Table1.ColorGradeImage = "ColorGradeLUT256x16_ConSat"

		'Lower Left String
		Case 4
            For each xx in GIString5:xx.IntensityScale = gistep * step:next
					if step = 1 then Table1.ColorGradeImage = "-70"
					if step = 2 then Table1.ColorGradeImage = "-60"
					if step = 3 then Table1.ColorGradeImage = "-50"
					if step = 4 then Table1.ColorGradeImage = "-40"
					if step = 5 then Table1.ColorGradeImage = "-30"
					if step = 6 then Table1.ColorGradeImage = "-20"
					if step = 7 then Table1.ColorGradeImage = "-10"
					if step = 8 then Table1.ColorGradeImage = "ColorGradeLUT256x16_ConSat"
    End Select
End Sub

'****************************
'     Realtime Updates
' called by the MotorCallBack
'****************************

Sub RealTimeUpdates
    'flippers
    LeftFlipperP.RotY = LeftFlipper.CurrentAngle
    LeftFlipperP1.RotY = LeftFlipper1.CurrentAngle
    RightFlipperP.RotY = RightFlipper.CurrentAngle
' rolling sound
    RollingSoundUpdate
' Plunger update
    PlungerPTimer
' ramp gate
	GateTimer
' diverter
    DiverterP.RotY = DiverterR.CurrentAngle
End Sub

'******************************
' Diverse Collection Hit Sounds
'******************************


Sub Pins_Hit (idx)
	PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub


Sub Metals_Thin_Hit (idx)
	PlaySound "metalhit_thin", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
	PlaySound "metalhit_medium", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
	PlaySound "metalhit2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
	PlaySound "gate4", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinner_Spin
	PlaySound "fx_spinner",0,.25,0,0.25
End Sub

Sub Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub Posts_Hit(idx)
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
		Case 1 : PlaySound "flip_hit_1", 0, 1, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "flip_hit_2", 0, 1, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "flip_hit_3", 0, 1, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub

Sub RandomSoundHole()
    Select Case Int(Rnd * 3) + 1
        Case 1:PlaySound "Hole1", 0, 0.4, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
        Case 2:PlaySound "Hole2", 0, 0.4, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
        Case 3:PlaySound "Hole4", 0, 0.4, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    End Select
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
'    JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 9 ' total number of balls
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = FALSE
    Next
End Sub

Sub RollingSoundUpdate()
    Dim BOT, b
    BOT = GetBalls

    ' stop the sound of deleted balls
    For b = UBound(BOT) + 1 to tnob
        rolling(b) = FALSE
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


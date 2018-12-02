'Road Kings for VPX by Sliderpoint
Option Explicit
Randomize

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions"
' Changed UseSolenoids=1 to 2
' Table has a non standard "Ball rolling function"
' Thalamus 2018-11-01 : Improved directional sounds
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
Const VolTarg   = 1    ' Targets volume.
Const VolKick   = 1    ' Kicker volume.
Const VolFlip   = 1    ' Flipper volume.


On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM"01210000","S11.VBS",3.1

'standard definitions
Const UseSolenoids=2,UseLamps=0,UseGI=0,SSolenoidOn="solenoid",SSolenoidOff="solenoidoff",SFlipperOn="FlipperUp",SFlipperOff="FlipperDown",SCoin="CoinDrop"
Const UseSync = 0
Const HandleMech = 0
Const sKnocker ="Knocker"

Dim DesktopMode: DesktopMode = Table1.ShowDT

Sub LoadCoreVBS
     On Error Resume Next
     ExecuteGlobal GetTextFile("core.vbs")
     If Err Then MsgBox "Can't open core.vbs"
     On Error Goto 0
End Sub

'table init
Const cGameName = "Rdkng_L4"
Sub Table1_Init
	vpmInit Me
	With Controller
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
		.SplashInfoLine="Road Kings, Williams, 1986"
		.Games(cGameName).Settings.Value("rol") = 0
		.HandleKeyboard=0
		.ShowTitle=0
		.ShowDMDOnly=1
		.HandleMechanics=0
		.DIP(0)=&H00
		.ShowFrame=0
		.Hidden = 0
		On Error Resume Next
		.Run GetPlayerHWnd
		If Err Then MsgBox Err.Description
		On Error Goto 0
	End With

'Main Timer
	PinMAMETimer.Interval=PinMAMEInterval
	PinMAMETimer.Enabled=1

'Other init
	Plunger1.Pullback
	GoRight.IsDropped=1
	Ramp1.Collidable=False
	Controller.Switch(39)=0

'Nudging
    vpmNudge.TiltSwitch = 1
    vpmNudge.Sensitivity = 2

'Trough
	Set bsTrough=New cvpmBallStack
		with bsTrough
			.InitSw 40,41,42,0,0,0,0,0
			.InitKick BallRelease,90,7
			.InitExitSnd SoundFX("ballrel",DOFContactors),SoundFX("solenoid",DOFContactors)
			.IsTrough = True
			.Balls=2
		end with
'kickers
	Set bsLeftLock=New cvpmBallStack
	Set bsLeft=New cvpmBallStack
	Set bsCenter=New cvpmBallStack

'Drop Target
	Set dtDrop=New cvpmDropTarget
	dtDrop.InitDrop DropTarget,30
	dtDrop.InitSnd SoundFX("DropTarget",DOFContactors),SoundFX("resetDrop",DOFContactors)
End Sub

'Pause/exit support
Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub
Sub table1_Exit:Controller.Stop:End Sub

'Balls in locks
 Sub Trigger2_Hit
 	If bsCenter.Balls Then
 		If ActiveBall.VelY<0 Then ActiveBall.VelY=-ActiveBall.VelY/2
		PlaySoundAtVol "ball_ball_hit",Trigger2, 1
	End If
 End Sub

Sub Trigger4_Hit
 	If bsLeft.Balls Then
 		If ActiveBall.VelY<0 Then ActiveBall.VelY=-ActiveBall.VelY/2
		PlaysoundAtVol "ball_ball_hit", Trigger4, 1
 	End If
 End Sub

'Keys
Sub Table1_KeyDown(ByVal keycode)

	If keycode = PlungerKey Then
		Plunger.PullBack
		PlaysoundAtVol "PlungerPull", Plunger, 1
	End If

	If keycode = LeftTiltKey Then
		Nudge 90, 2: PlaySound SoundFX("Nudge_Left",0)
	End If

	If keycode = RightTiltKey Then
		Nudge 270, 2: PlaySound SoundFX("Nudge_Right",0)
	End If

	If keycode = CenterTiltKey Then
		Nudge 0, 2: PlaySound SoundFX("Nudge_Forward",0)
	End If

	If VPMKeyDown(KeyCode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal keycode)

	If keycode = PlungerKey Then
		Plunger.Fire
		PlaySoundAtVol "plungerRel", Plunger, 1
	End If

	If VPMKeyUp(KeyCode) Then Exit Sub
End Sub

' RealTime Updates
Set MotorCallback = GetRef("GameTimer")

	Sub GameTimer
		UpdateFlipperLogos
		UpdateGates
		RollingSound
	End Sub

Dim bsTrough, bsLKick, bsREject, bsAPlunger, x, xx, bump1, bump2, bump3, bump4, target9, target10, target11, target12, target13, target14, target15, target16, target17, bsLeft,bsCenter,dtDrop,bsLeftLock,obj

'Drain
Sub Drain_Hit()   '40/41/42
	ClearBallid
	Drain.DestroyBall
	PlaySoundAtVol "Drain", drain, 1
	bsTrough.addball ME
End Sub

'Flipper Logos
	Sub UpdateFlipperLogos
			Primitive52.objRoty = RightFlipper.currentAngle
			Primitive51.objroty = LeftFlipper.CurrentAngle
	End Sub

'Primitive Gate rotations
	Sub UpdateGates
			GatePrim.RotZ = -(Gate.currentangle +15)
			Gate2Prim.RotZ = -(Gate2.currentangle +15)
			Gate3Prim.RotZ = -(Gate3.currentangle +15)
			Gate4Prim.RotZ = -(Gate4.currentangle)
			Gate5Prim.RotZ = -(Gate5.currentangle)
			Gate6Prim.RotZ = -(Gate6.currentangle -90)
			Gate7Prim.RotZ = -(Spinner7.currentangle)
			Gate8Prim.RotZ = -(Spinner8.currentangle)
			Gate9Prim.RotZ = -(Gate9.currentangle)

	End Sub
'Bumpers
      Sub Bumper1b_Hit::PlaySoundAtVol SoundFX("bumper",DOFContactors),bumper1b,VolBump:bump1 = 1:Me.TimerEnabled = 1:vpmTimer.PulseSw 24:End Sub

       Sub Bumper1b_Timer()
           Select Case bump1
               Case 1:BR1.z = -10:bump1 = 2
               Case 2:BR1.z = -18:bump1 = 3
               Case 3:BR1.z = -32:bump1 = 4
               Case 4:BR1.z = -32:bump1 = 5
               Case 5:BR1.z = -18:bump1 = 6
               Case 6:BR1.z = -10:bump1 = 7
               Case 7:BR1.z = 0:Me.TimerEnabled = 0
           End Select
       End Sub

      Sub Bumper2b_Hit::PlaySoundAtVol SoundFX("bumper",DOFContactors),bumper2b,VolBump:bump2 = 1:Me.TimerEnabled = 1:vpmTimer.PulseSw 26:End Sub

       Sub Bumper2b_Timer()
           Select Case bump2
               Case 1:BR2.z = -10:bump2 = 2
               Case 2:BR2.z = -18:bump2 = 3
               Case 3:BR2.z = -32:bump2 = 4
               Case 4:BR2.z = -32:bump2 = 5
               Case 5:BR2.z = -18:bump2 = 6
               Case 6:BR2.z = -10:bump2 = 7
               Case 7:BR2.z = 0:Me.TimerEnabled = 0
           End Select
       End Sub

      Sub Bumper3b_Hit::PlaySoundAtVol SoundFX("bumper",DOFContactors),bumper3b,volbump:bump3 = 1:Me.TimerEnabled = 1:vpmTimer.PulseSw 25:End Sub

       Sub Bumper3b_Timer()
           Select Case bump3
               Case 1:BR3.z = -10:bump3 = 2
               Case 2:BR3.z = -18:bump3 = 3
               Case 3:BR3.z = -32:bump3 = 4
               Case 4:BR3.z = -32:bump3 = 5
               Case 5:BR3.z = -18:bump3 = 6
               Case 6:BR3.z = -10:bump3 = 7
               Case 7:BR3.z = 0:Me.TimerEnabled = 0
           End Select
		End Sub
	  Sub Bumper4b_Hit::PlaySoundAtVol SoundFX("bumper",DOFContactors),bumper4b,volbump:bump4 = 1:Me.TimerEnabled = 1:vpmTimer.PulseSw 27:End Sub

       Sub Bumper4b_Timer()
           Select Case bump4
               Case 1:BR4.z = -10:bump4 = 2
               Case 2:BR4.z = -18:bump4 = 3
               Case 3:BR4.z = -32:bump4 = 4
               Case 4:BR4.z = -32:bump4 = 5
               Case 5:BR4.z = -18:bump4 = 6
               Case 6:BR4.z = -10:bump4 = 7
               Case 7:BR4.z = 0:Me.TimerEnabled = 0
           End Select
       End Sub
'Targets
	  Sub T9_Hit:PlaySoundAtVol SoundFX("Target",DOFContactors),ActiveBall,VolTarg:Target9 = 1::vpmTimer.PulseSw 9:End Sub
	  Sub T10_Hit:PlaySoundAtVol SoundFX("Target",DOFContactors),ActiveBall,VolTarg:Target10 = 1:vpmTimer.PulseSw 10:End Sub
    Sub T11_Hit:PlaySoundAtVol SoundFX("Target",DOFContactors),ActiveBall,VolTarg:Target11 = 1:vpmTimer.PulseSw 11:End Sub
	  Sub T12_Hit:PlaySoundAtVol SoundFX("Target",DOFContactors),ActiveBall,VolTarg:Target12 = 1:vpmTimer.PulseSw 12:End Sub
	  Sub T13_Hit:PlaySoundAtVol SoundFX("Target",DOFContactors),ActiveBall,VolTarg:Target13 = 1:vpmTimer.PulseSw 13:End Sub
	  Sub T14_Hit:PlaySoundAtVol SoundFX("Target",DOFContactors),ActiveBall,VolTarg:Target14 = 1:vpmTimer.PulseSw 14:End Sub
	  Sub T15_Hit:PlaySoundAtVol SoundFX("Target",DOFContactors),ActiveBall,VolTarg:Target15 = 1:vpmTimer.PulseSw 15:End Sub
	  Sub T16_Hit:PlaySoundAtVol SoundFX("Target",DOFContactors),ActiveBall,VolTarg:Target16 = 1:vpmTimer.PulseSw 16:End Sub
	  Sub T17_Hit:PlaySoundAtVol SoundFX("Target",DOFContactors),ActiveBall,VolTarg:Target17 = 1:vpmTimer.PulseSw 17:End Sub

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingshot_Slingshot()
	PlaySoundAtVol SoundFX("Sling",DOFContactors), sling1, 1
	vpmTimer.PulseSw 44
	RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingshot_Slingshot()
	PlaySoundAtVol SoundFX("Sling",DOFContactors), sling2, 1
	vpmTimer.PulseSw 43
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub


Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

'Solonoids
SolCallback(1)="bsTrough.SolIn"
SolCallback(2)="bsTrough.SolOut"
SolCallback(3)="bsLeft.SolOut"
SolCallback(4)="bsCenter.SolOut"
SolCallBack(5)="RearGI" 'Rear playfield flashers
'SolCallback(6)="vpmSolAutoPlunger Plunger1,1," ' power kicker
SolCallback(6) = "SolKickback"
SolCallback(7)="vpmFlasher Array(Flash7, Flash1, Flash2, Flash3)," 'left bolt
SolCallback(8)="vpmFlasher Array(Flash4, Flash5, Flash6, Flash8)," 'right bolt
SolCallback(9)="vpmSolGate Gate3,true,"
SolCallback(10)="vpmSolGate Gate2,true,"
SolCallback(11)="SolGI"
'SolCallBack(12)= 'Solenoid Select Relay
SolCallback(13)="vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(14)="BoardGI" 'Mid-insert board flashers
'SolCallback(15)="SolBikeFlash" 'bikes flasher backbox
'SolCallback(16)= ' Coin Lock Out relay
SolCallback(17)="vpmSolSound SoundFX(""sling"",DOFContactors)," 'left slingshot
SolCallback(18)="vpmSolSound SoundFX(""sling"",DOFContactors)," 'right slingshot
SolCallback(19)="vpmSolSound SoundFX(""Jet3"",DOFContactors),"
SolCallback(20)="vpmSolSound SoundFX(""Jet3"",DOFContactors),"
SolCallback(21)="vpmSolSound SoundFX(""Jet3"",DOFContactors),"
SolCallback(22)="vpmSolSound SoundFX(""Jet3"",DOFContactors),"
SolCallback(25)="bsLeftLock.SolOut"
SolCallback(26)="SolRampUp"
SolCallback(27)="SolRampDown"
SolCallback(28)="dtDrop.SolDropUp"

Sub SolKickBack(enabled)
    If enabled Then
       Plunger1.Fire
       PlaySoundAtVol SoundFX("popper_ball",DOFContactors), plunger1, 1
    Else
       Plunger1.PullBack
    End If
End Sub

'Flippers
SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFX("flipperup",DOFContactors), LeftFlipper, VolFlip
        LeftFlipper.RotateToEnd
    Else
        PlaySoundAtVol SoundFX("flipperdown",DOFContactors), LeftFlipper, VolFlip
        LeftFlipper.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
        PlaySoundAtVol SoundFX("flipperup",DOFContactors), RightFlipper, VolFlip
        RightFlipper.RotateToEnd
    Else
        PlaySoundAtVol SoundFX("flipperdown",DOFContactors), RightFlipper, VolFlip
        RightFlipper.RotateToStart
    End If
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySound "rubber", 0, parm / 10, -0.1, 0.25
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "rubber", 0, parm / 10, 0.1, 0.25
End Sub

'Rollovers
Sub SW18_Hit:PlaySoundAtVol "RollOver",ActiveBall, 1:Controller.Switch(18)=1:End Sub	'18
Sub SW18_unHit:Controller.Switch(18)=0:End Sub
Sub SW19_Hit:PlaySoundAtVol "RollOver",ActiveBall, 1:Controller.Switch(19)=1:End Sub	'19
Sub SW19_unHit:Controller.Switch(19)=0:End Sub
Sub SW20_Hit:PlaySoundAtVol "RollOver",ActiveBall, 1:Controller.Switch(20)=1:End Sub	'20
Sub SW20_unHit:Controller.Switch(20)=0:End Sub
Sub SW21_Hit:PlaySoundAtVol "RollOver",ActiveBall, 1:Controller.Switch(21)=1:End Sub	'21
Sub SW21_unHit:Controller.Switch(21)=0:End Sub
Sub SW22_Hit:Controller.Switch(22)=1:End Sub	'22
Sub SW22_unHit:Controller.Switch(22)=0:End Sub
'Sub 		'23 Ramp Raise EOS
Sub SW28_Hit:PlaySoundAtVol "RollOver",ActiveBall, 1:Controller.Switch(28)=1:End Sub	'28
Sub SW28_unHit:Controller.Switch(28)=0:End Sub
Sub SW29_Hit:PlaySoundAtVol "RollOver",ActiveBall, 1:Controller.Switch(29)=1:End Sub	'29
Sub SW29_unHit:Controller.Switch(29)=0:End Sub

'DropTarget
Sub DropTarget_Hit:dtDrop.Hit 1:End Sub				'30

Sub Gate6_Hit:Controller.Switch(32)=1:PlaySoundAtVol "GateSmallWire",gate6, VolGates:End Sub	'32
Sub Gate6_unHit:Controller.Switch(32)=0:End Sub
Sub Rubber24_Hit:vpmTimer.PulseSw 33:End Sub		'33 Right Ten Point
Sub LeftLock_Hit:PlaySoundAtVol "Kicker_Enter",leftlock,VolKick:RampLock.Enabled = 1:End Sub	'34
Sub SW35_Hit:Controller.Switch(35)=1:PlaySoundAtVol "gatewire",sw35,1:End Sub	'35
Sub SW35_unHit:Controller.Switch(35)=0:PlaySoundAtVol "gatewire",sw35,1:End Sub
Sub SW36_Hit:Controller.Switch(36)=1:PlaySoundAtVol "gatewire",sw36,1:End Sub	'36
Sub SW36_unHit:Controller.Switch(36)=0:PlaySoundAtVol "gatewire",sw36,1:End Sub
Sub Kicker1_Hit:PlaySoundAtVol "Kicker_Enter",kicker1,VolKick:LeftLockT.Enabled = 1:End Sub		'37
Sub Kicker2_Hit:PlaySoundAtVol "Kicker_Enter",kicker2,VolKick:CenterLock.Enabled = 1:End Sub		'38
Sub SW39_Hit:Controller.Switch(39)=1:End Sub	'39
Sub SW39_unHit:Controller.Switch(39)=0:End Sub

Sub Rubber7_Hit:vpmTimer.PulseSw 45:End Sub	'45 Left Ten Point
'Sub    '46 playfield Tilt
Sub SW54_Hit:Controller.Switch(54)=1:End Sub	'54
Sub SW54_unHit:Controller.Switch(54)=0:End Sub

'Ramp Lock
Sub RampLock_timer()
		bsLeftLock.AddBall 0:
		bsLeftLock.InitSaucer LeftLock,34,0,60
		bsLeftLock.InitExitSnd SoundFX("popper_Ball",DOFContactors),SoundFX("solenoid",DOFContactors)
	RampLock.Enabled = 0
End Sub

'Left Lock
Sub LeftLockT_timer()
		bsLeft.AddBall 0:
		bsLeft.InitSaucer Kicker1,37,239,25
		bsLeft.kickZ = .4
		bsLeft.InitExitSnd SoundFX("popper_Ball",DOFContactors),SoundFX("solenoid",DOFContactors)
	LeftLockT.Enabled = 0
End Sub

'Center Lock
Sub CenterLock_timer()
		bsCenter.AddBall 0:
		bsCenter.InitSaucer Kicker2,38,118,22
		bsCenter.InitExitSnd SoundFX("popper_Ball",DOFContactors),SoundFX("solenoid",DOFContactors)
		popper.isDropped = 0
		popper.timerenabled = 1
	CenterLock.Enabled = 0
End Sub

Sub Popper_timer()
	Popper.isDropped = 1
	popper.timerenabled = 0
end Sub

'Diverter
Sub Gate9_Hit() '31
	vpmTimer.PulseSw 31
	PlaySoundAtVol "GateSmallWire", gate9, VolGates
End Sub

Dim Right1, Left1
Sub Tirgger7_Hit
	If GoRight.IsDropped Then
	ActiveBall.VelX=0.5
	Else
	ActiveBall.VelX=-0.5
	End If
End Sub

Sub Trigger5_Hit::PlaySoundAtVol "Tick",ActiveBall, 1:Right1 = 1:Me.TimerEnabled = 1:End Sub

       Sub Trigger5_Timer()
           Select Case Right1
               Case 1:Diverter.ObjRotZ = -30:Right1 = 2
               Case 2:Diverter.ObjRotZ = -20:Right1 = 3
               Case 3:Diverter.ObjRotZ = -10:Right1 = 4
               Case 4:Diverter.ObjRotZ = 0:Right1 = 5
               Case 5:Diverter.ObjRotZ = 10:Right1 = 6
               Case 6:Diverter.ObjRotZ = 20:Right1 = 7
               Case 7:Diverter.ObjRotZ = 30:Me.TimerEnabled = 0
           End Select
		End Sub


Sub Trigger6_Hit::PlaySoundAtVol "Tick",ActiveBall, 1:Left1 = 1:Me.TimerEnabled = 1:End Sub

       Sub Trigger6_Timer()
           Select Case Left1
               Case 1:Diverter.ObjRotZ = 30:Left1 = 2
               Case 2:Diverter.ObjRotZ = 20:Left1 = 3
               Case 3:Diverter.ObjRotZ = 10:Left1 = 4
               Case 4:Diverter.ObjRotZ = 0:Left1 = 5
               Case 5:Diverter.ObjRotZ = -10:Left1 = 6
               Case 6:Diverter.ObjRotZ = -20:Left1 = 7
               Case 7:Diverter.ObjRotZ = -30:Me.TimerEnabled = 0
           End Select
		End Sub


Sub Trigger7_unHit
	If (Not DivTimer.Enabled) And (Not Trigger7.TimerEnabled) And (Not DivTimer2.Enabled) Then
	DivTimer.Enabled=1
	Else
	Exit Sub
	End If
End Sub

Sub DivTimer_Timer
	Me.Enabled=0
	If GoRight.IsDropped Then
		GoLeft.IsDropped=1
		GoRight.IsDropped=0
	Else
		GoLeft.IsDropped=0
		GoRight.IsDropped=1
	End If
		Gate9.TimerEnabled=1
	End Sub

Sub Trigger7_Timer()
	Trigger7.TimerEnabled=0
	If GoRight.IsDropped Then
		GoLeft.IsDropped=1
		GoRight.IsDropped=0
	Else
		GoLeft.IsDropped=0
		GoRight.IsDropped=1
	End If
	DivTimer2.Enabled=1
End Sub

Sub DivTimer2_Timer
	DivTimer2.Enabled=0
	If GoRight.IsDropped Then
		GoLeft.IsDropped=1
		GoRight.IsDropped=0
	Else
		GoLeft.IsDropped=0
		GoRight.IsDropped=1
	End If
End Sub


'Ramp activation
Dim RampBalls, RampLift, RampDrop

 RampBalls=0
 Sub Trigger1_Hit:RampBalls=RampBalls+1:End Sub
 Sub Trigger1_unHit:RampBalls=RampBalls-1:End Sub

Sub SolRampDown(Enabled)
	If Enabled AND MetalRamp.ObjRotX = 10 Then 'RampDown
		Controller.Switch(23) = 1
 		If RampBalls<1 Then
			Ramp1.Collidable=True
			RampDrop = 1
			RampDownTimer.Enabled = 1
			Trigger1.TimerEnabled = 0
		Else
 			Trigger1.TimerEnabled = 1
 		End If
	End If
End Sub

Sub RampDownTimer_Timer()
		Select Case RampDrop
				Case 1:MetalRamp.ObjRotX = 10:RampArm.RotY = 30:PlaysoundAtVol "RampDown",Gate9Prim,1:RampDrop = 2
				Case 2:MetalRamp.ObjRotX = 7:RampArm.RotY = 15:RampDrop = 3
				Case 3:MetalRamp.ObjRotX = 5:RampArm.RotY = 0:RampDrop = 4
				Case 4:MetalRamp.ObjRotX = 3:RampArm.RotY = -15:RampDrop = 5
				Case 5:MetalRamp.ObjRotX = 0:RampArm.RotY = -30:RampDownTimer.Enabled = 0
		End Select
End Sub

Sub SolRampUp(Enabled)
	If Enabled AND MetalRamp.ObjRotX = 0 Then 'RampUp
		Ramp1.Collidable=False
		RampLift = 1
		RampUpTimer.Enabled = 1
		Controller.Switch(23)= 0
	End If
End Sub

Sub RampUpTimer_Timer()
		Select Case RampLift
				Case 1:MetalRamp.ObjRotX = 0:RampArm.RotY = -30:PlaysoundAtVol "RampUp",Gate9Prim,1:RampLift = 2
				Case 2:MetalRamp.ObjRotX = 3:RampArm.RotY = -15:RampLift = 3
				Case 3:MetalRamp.ObjRotX = 5:RampArm.RotY = 0:RampLift = 4
				Case 4:MetalRamp.ObjRotX = 7:RampArm.RotY = 15:RampLift = 5
				Case 5:MetalRamp.ObjRotX = 10:RampArm.RotY = 30:RampUpTimer.Enabled = 0
		End Select
End Sub

'Only drop ramp if there is no ball under the ramp
 Sub Trigger1_Timer
 	If RampBalls<1 AND MetalRamp.ObjRotX = 9 Then
		Ramp1.Collidable=True
		RampDrop = 1
		RampDownTimer.Enabled = 1
 		Trigger1.TimerEnabled = 0
 	End If
End Sub

'Update GI stuff...
Sub SolGI(Enabled)
	If Enabled Then
			Light1.state = 0
			Light2.state = 0
			Light3.state = 0
			Light4.state = 0
			Light5.state = 0
			Light6.state = 0
			Light7.state = 0
			Light8.state = 0
			Light9.state = 0
			Light10.state = 0
			Light11.state = 0
			Light12.state = 0
			Light13.state = 0
			Light14.state = 0
			Light15.state = 0
			Light16.state = 0
			Light17.state = 0
			Light18.state = 0
			Light19.state = 0
			Light20.state = 0
			Light21.state = 0
			Light22.state = 0
			Light23.state = 0
			Light25.state = 0
			Light28.state = 0
			Light29.state = 0
			Light30.state = 0
			Light31.State = 0
			Light32.State = 0
			Light33.State = 0
			Light34.State = 0
			Light35.State = 0
			Light36.State = 0
			Light37.State = 0
			Light38.State = 0
			Light39.State = 0
			Light40.State = 0
			Light41.State = 0
			Light42.State = 0
			Light43.State = 0
			Light44.State = 0
			Light45.State = 0
			Light46.State = 0
			Light47.State = 0
'			Light48.State = 0
'			Light49.State = 0
			Light50.State = 0
			Light51.State = 0
			Light52.State = 0
			Light53.State = 0
			Light54.State = 0
			Light55.State = 0
			Light56.State = 0
'			Light57.State = 0
			Light58.State = 0
			Light59.state = 0
			Light60.state = 0
			Light61.state = 0
			Light62.state = 0
			Light63.state = 0
			Light64.state = 0
			Light65.state = 0
			Light66.state = 0
			Light67.state = 0
			Light68.state = 0
			Light69.state = 0
			Light70.state = 0
			Light71.state = 0
			Light72.state = 0
			Light73.state = 0
			Light74.state = 0
			Light75.state = 0
			Light76.state = 0
			Light77.state = 0
			Light78.state = 0
			Light79.state = 0
			Light80.state = 0
			Light81.state = 0
			Light82.state = 0
			Light83.state = 0
			Light84.state = 0
			Light85.state = 0
			Light86.state = 0
			Light87.state = 0
			Light88.state = 0
			Light89.state = 0
			Light90.state = 0
			Light91.state = 0
			Light92.state = 0
			Light93.state = 0
			Light94.state = 0
			Light95.state = 0
			Light96.state = 0
			Light97.state = 0
			Light98.state = 0
			Flasher1.Visible = 0
			RightWall.visible = 0
			LeftWall.visible = 0
			Flasher2.Visible = 0


		Else
			Light1.state = 1
			Light2.state = 1
			Light3.state = 1
			Light4.state = 1
			Light5.state = 1
			Light6.state = 1
			Light7.state = 1
			Light8.state = 1
			Light9.state = 1
			Light10.State = 1
			Light11.state = 1
			Light12.state = 1
			Light13.state = 1
			Light14.state = 1
			Light15.state = 1
			Light16.State = 1
			Light17.state = 1
			Light18.state = 1
			Light19.state = 1
			Light20.state = 1
			Light21.state = 1
			Light22.state = 1
			Light23.state = 1
			Light25.state = 1
			Light28.state = 1
			Light29.state = 1
			Light30.state = 1
			Light31.State = 1
			Light32.State = 1
			Light33.State = 1
			Light34.State = 1
			Light35.State = 1
			Light36.State = 1
			Light37.State = 1
			Light38.State = 1
			Light39.State = 1
			Light40.State = 1
			Light41.State = 1
			Light42.State = 1
			Light43.State = 1
			Light44.State = 1
			Light45.State = 1
			Light46.State = 1
			Light47.State = 1
'			Light48.State = 1
'			Light49.State = 1
			Light50.State = 1
			Light51.State = 1
			Light52.State = 1
			Light53.State = 1
			Light54.State = 1
			Light55.State = 1
			Light56.State = 1
'			Light57.State = 1
			Light58.State = 1
			Light59.State = 1
			Light60.state = 1
			Light61.state = 1
			Light62.state = 1
			Light63.state = 1
			Light64.state = 1
			Light65.state = 1
			Light66.state = 1
			Light67.state = 1
			Light68.state = 1
			Light69.state = 1
			Light70.state = 1
			Light71.state = 1
			Light72.state = 1
			Light73.state = 1
			Light74.state = 1
			Light75.state = 1
			Light76.state = 1
			Light77.state = 1
			Light78.state = 1
			Light79.state = 1
			Light80.state = 1
			Light81.state = 1
			Light82.state = 1
			Light83.state = 1
			Light84.state = 1
			Light85.state = 1
			Light86.state = 1
			Light87.state = 1
			Light88.state = 1
			Light89.state = 1
			Light90.state = 1
			Light91.state = 1
			Light92.state = 1
			Light93.state = 1
			Light94.state = 1
			Light95.state = 1
			Light96.state = 1
			Light97.state = 1
			Light98.state = 1
			Flasher1.Visible = 1
			RightWall.Visible = 1
			LeftWall.Visible = 1
			Flasher2.Visible = 1

	End if
End Sub

Sub BoardGI(Enabled)
	If Enabled Then
			Flasher6.Visible=1
		Else
			Flasher6.Visible=0
	End if
End Sub

Sub RearGI(Enabled)
	If Enabled Then
			Light26.State = 1
			Light27.State = 1
			RightWall1.Visible = 1
			LeftWall1.visible = 1
		Else
			Light26.State = 0
			Light27.State = 0
			RightWall1.Visible = 0
			LeftWall1.Visible = 0
	End if
End Sub


Set LampCallback = GetRef("Lamps")
	Sub Lamps
		L1.State = Controller.Lamp(1)
		L2.State = Controller.Lamp(2)
		L3.State = Controller.Lamp(3)
		L4.State = Controller.Lamp(4)
		L5.State = Controller.Lamp(5)
		L6.State = Controller.Lamp(6)
		L7.State = Controller.Lamp(7)
		L8.State = Controller.Lamp(8)
		L9.State = Controller.Lamp(9)
		L10.State = Controller.Lamp(10)
		L11.State = Controller.Lamp(11)
		L12.State = Controller.Lamp(12)
		L13.State = Controller.Lamp(13)
		L14.State = Controller.Lamp(14)
		L15.State = Controller.Lamp(15)
		L16.State = Controller.Lamp(16)
		L17.State = Controller.Lamp(17)
		L18.State = Controller.Lamp(18)
		L19.State = Controller.Lamp(19)
		L20.State = Controller.Lamp(20)
		L21.State = Controller.Lamp(21)
		L22.State = Controller.Lamp(22)
		L23.State = Controller.Lamp(23)
		L24.State = Controller.Lamp(24)
		L25.State = Controller.Lamp(25)
		L26.State = Controller.Lamp(26)
		L27.State = Controller.Lamp(27)
		L28.State = Controller.Lamp(28)
		L29.State = Controller.Lamp(29)
		L30.State = Controller.Lamp(30)
		L31.State = Controller.Lamp(31)
		L32.State = Controller.Lamp(32)
		L33.State = Controller.Lamp(33)
		L34.State = Controller.Lamp(34)
		L35.State = Controller.Lamp(35)
		L36.State = Controller.Lamp(36)
		L37.State = Controller.Lamp(37)
		L38.State = Controller.Lamp(38)
		L39.State = Controller.Lamp(39)
		L40.State = Controller.Lamp(40)
		L41.State = Controller.Lamp(41)
		L42.State = Controller.Lamp(42)
		L43.State = Controller.Lamp(43)
		L44.State = Controller.Lamp(44)
		L45.State = Controller.Lamp(45)
		L46.State = Controller.Lamp(46)
		L47.State = Controller.Lamp(47)
		L48.State = Controller.Lamp(48)
		L49.State = Controller.Lamp(49)
		L50.State = Controller.Lamp(50)
		L51.State = Controller.Lamp(51)
		L52.State = Controller.Lamp(52)
		L53.State = Controller.Lamp(53)
		L54.State = Controller.Lamp(54)
		L55.State = Controller.Lamp(55)

	End Sub

sub FlasherTimer_Timer()
		If L19.State Then
			Flasher3.Visible = 1
		Else
			Flasher3.Visible = 0
		End If

		If L18.State Then
			Flasher4.Visible = 1
		Else
			Flasher4.Visible = 0
		End If

		If L28.State Then
			Flasher5.Visible = 1
		Else
			Flasher5.Visible = 0
		End If

End Sub

'backdrop LCD lettering from Sinbad

Dim ChgLed, iw
Sub Leds_Timer()
	ChgLed = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
	If DesktopMode = True Then
		UpdateLeds
		Controller.hidden = 1
		For each iw in BGLCDs: iw.Visible = True:Next
	Else
		For each iw in BGLCDs: iw.Visible = False:Next
	End If
End Sub

Sub UpdateLeds()
	Dim ii, num, chg, stat, obj
	If Not IsEmpty(ChgLED) Then
		For ii=0 To UBound(chgLED)
			num=chgLED(ii,0):chg=chgLED(ii,1):stat=chgLED(ii,2)
			For Each obj In Digits(num)
				If chg And 1 Then obj.State=stat And 1
				chg=chg\2:stat=stat\2
			Next
		Next
	End If
End Sub

Dim Digits(32)
Digits(0)=Array(A1,A2,A3,A4,A5,A6,A7,A8,A9,A10,A11,A12,A13,A14,A15)
Digits(1)=Array(A16,A17,A18,A19,A20,A21,A22,A23,A24,A25,A26,A27,A28,A29,A30)
Digits(2)=Array(A31,A32,A33,A34,A35,A36,A37,A38,A39,A40,A41,A42,A43,A44,A45)
Digits(3)=Array(A46,A47,A48,A49,A50,A51,A52,A53,A54,A55,A56,A57,A58,A59,A60)
Digits(4)=Array(A61,A62,A63,A64,A65,A66,A67,A68,A69,A70,A71,A72,A73,A74,A75)
Digits(5)=Array(A76,A77,A78,A79,A80,A81,A82,A83,A84,A85,A86,A87,A88,A89,A90)
Digits(6)=Array(A91,A92,A93,A94,A95,A96,A97,A98,A99,A100,A101,A102,A103,A104,A105)
Digits(7)=Array(A106,A107,A108,A109,A110,A111,A112,A113,A114,A115,A116,A117,A118,A119,A120)
Digits(8)=Array(A121,A122,A123,A124,A125,A126,A127,A128,A129,A130,A131,A132,A133,A134,A135)
Digits(9)=Array(A136,A137,A138,A139,A140,A141,A142,A143,A144,A145,A146,A147,A148,A149,A150)
Digits(10)=Array(A151,A152,A153,A154,A155,A156,A157,A158,A159,A160,A161,A162,A163,A164,A165)
Digits(11)=Array(A166,A167,A168,A169,A170,A171,A172,A173,A174,A175,A176,A177,A178,A179,A180)
Digits(12)=Array(A181,A182,A183,A184,A185,A186,A187,A188,A189,A190,A191,A192,A193,A194,A195)
Digits(13)=Array(A196,A197,A198,A199,A200,A201,A202,A203,A204,A205,A206,A207,A208,A209,A210)
Digits(14)=Array(A211,A212,A213,A214,A215,A216,A217,A218)
Digits(15)=Array(A219,A220,A221,A222,A223,A224,A225,A226)
Digits(16)=Array(A227,A228,A229,A230,A231,A232,A233,A234)
Digits(17)=Array(A235,A236,A237,A238,A239,A240,A241,A242)
Digits(18)=Array(A243,A244,A245,A246,A247,A248,A249,A250)
Digits(19)=Array(A251,A252,A253,A254,A255,A256,A257,A258)
Digits(20)=Array(A259,A260,A261,A262,A263,A264,A265,A266)
Digits(21)=Array(A267,A268,A269,A270,A271,A272,A273,A274)
Digits(22)=Array(A275,A276,A277,A278,A279,A280,A281,A282)
Digits(23)=Array(A283,A284,A285,A286,A287,A288,A289,A290)
Digits(24)=Array(A291,A292,A293,A294,A295,A296,A297,A298)
Digits(25)=Array(A299,A300,A301,A302,A303,A304,A305,A306)
Digits(26)=Array(A307,A308,A309,A310,A311,A312,A313,A314)
Digits(27)=Array(A315,A316,A317,A318,A319,A320,A321,A322)
Digits(28)=Array(A339,A344,A345,A341,A343,A342,A340,A346)
Digits(29)=Array(A347,A352,A353,A349,A351,A350,A348,A354)
Digits(30)=Array(A323,A328,A329,A325,A327,A326,A324,A330)
Digits(31)=Array(A331,A336,A337,A333,A335,A334,A332,A338)


' *********************************************************************
' 					Wall, rubber and metal hit sounds
' *********************************************************************

Sub aRubbers_Hit(idx):PlaySound "rubber", 0, 1, pan(ActiveBall)*VolRH, 0.25, AudioFade(ActiveBall):End Sub
Sub aPosts_Hit(idx):PlaySound "rubber_hit_2", 0, Vol(ActiveBall)*VolPo, pan(ActiveBall), 0.25, AudioFade(ActiveBall):End Sub
Sub aMetals_Hit(idx):PlaySound "MetalHit", 0, 1, pan(ActiveBall)*VolMetal, 0.25, AudioFade(ActiveBall):End Sub
Sub Wires_Hit(idx):PlaySound "BallHitWire", 0, 1, pan(ActiveBall)*VolMetal, 0.25, AudioFade(ActiveBall):End Sub


'**************************************
'     Ball Rolling and Hit Sounds
'**************************************
' each ball is assigned an ID from the ball creation sub
' sounds are played based on the balls speed and position
' hit sounds on rubbers and metals are also based on the balls speed and position for a better 3D sound
' this rolling sub will also create a sound when the ball was on the air and hit back the playfield

ReDim rolling(tnopb), airball(tnopb)
Dim b
b = 0

Sub RollingSound()
    b = b + 1
    If b > tnopb Then b = 1
    If BallStatus(b) = 0 Then
        If rolling(b) = True Then
            StopSound "ballrolling"
            rolling(b) = False
            Exit Sub
        Else
            Exit Sub
        End If
    End if
    'If BallVel(CurrentBall(b) ) > 0 AND OnPlayfield(CurrentBall(b) ) Then 'do the rolling sound if the ball is moving and it is on the playfield
    If BallVel(CurrentBall(b) ) > 0 AND (Currentball(b).Z <67)Then 'do the rolling sound if the ball is moving
        If rolling(b) = True then
            PlaySound "ballrolling", -1, Vol(CurrentBall(b) ), Pan(CurrentBall(b) ), 0, Pitch(CurrentBall(b) ), 1, 0, AudioFade(ActiveBall)
        Else
            rolling(b) = True
            If airball(b) = True Then
                PlaySound "ballhit", 0, ABS(CurrentBall(b).Y / table1.height), Pan(CurrentBall(b) ), 0, Pitch(CurrentBall(b) ), 0, 0, AudioFade(ActiveBall)
                airball(b) = False
            End If
            PlaySound "ballrolling", -1, Vol(CurrentBall(b) ), Pan(CurrentBall(b) ), 0, Pitch(CurrentBall(b) ), 1, 0, AudioFade(ActiveBall)
        End If
    Else
        If rolling(b) = True Then
            StopSound "ballrolling"
            rolling(b) = False
        End If
    End If
End Sub

'*************************************************
' destruk's new vpmCreateBall for ball collision
' use it: vpmCreateBall kicker
'*************************************************

Set vpmCreateBall = GetRef("mycreateball")
Function mycreateball(aKicker)
    For cnt = 1 to ubound(ballStatus)
        If ballStatus(cnt) = 0 Then
            If Not IsEmpty(vpmBallImage) Then
                Set CurrentBall(cnt) = aKicker.CreateSizedBall(BSize).Image
            Else
                Set CurrentBall(cnt) = aKicker.CreateSizedBall(BSize)
            End If
            Set mycreateball = aKicker
            currentball(cnt).uservalue = cnt
            ballStatus(cnt) = 1
            ballStatus(0) = ballStatus(0) + 1
            If coff = False Then
                If ballStatus(0) > 1 and XYdata.enabled = False Then XYdata.enabled = True
            End If
            Exit For
        End If
    Next
End Function

'**************************************************************
' vpm Ball Collision based on the code by Steely & Pinball Ken
' added destruk's changes, ball size and height check by koadic
'**************************************************************

Const tnopb = 2 'max nr. of balls
Const nosf = 10  'nr. of sound files

ReDim CurrentBall(tnopb), BallStatus(tnopb)
Dim iball, cnt, coff, errMessage

XYdata.interval = 1
coff = False

For cnt = 0 to ubound(BallStatus):BallStatus(cnt) = 0:Next

' Create ball in kicker and assign a Ball ID used mostly in non-vpm tables
Sub CreateBallID(aKicker)
    For cnt = 1 to ubound(ballStatus)
        If ballStatus(cnt) = 0 Then
            Set CurrentBall(cnt) = aKicker.CreateSizedBall(Bsize)
            currentball(cnt).uservalue = cnt
            ballStatus(cnt) = 1
            ballStatus(0) = ballStatus(0) + 1
            If coff = False Then
                If ballStatus(0) > 1 and XYdata.enabled = False Then XYdata.enabled = True
            End If
            Exit For
        End If
    Next
End Sub

Sub ClearBallID
    On Error Resume Next
    iball = ActiveBall.uservalue
    currentball(iball).UserValue = 0
    If Err Then Msgbox Err.description & vbCrLf & iball
    ballStatus(iBall) = 0
    ballStatus(0) = ballStatus(0) -1
    On Error Goto 0
End Sub

' Ball data collection and B2B Collision detection. jpsalas: added height check
ReDim baX(tnopb, 4), baY(tnopb, 4), baZ(tnopb, 4), bVx(tnopb, 4), bVy(tnopb, 4), TotalVel(tnopb, 4)
Dim cForce, bDistance, xyTime, cFactor, id, id2, id3, B1, B2

Sub XYdata_Timer()
    xyTime = Timer + (XYdata.interval * .001)
    If id2 >= 4 Then id2 = 0
    id2 = id2 + 1
    For id = 1 to ubound(ballStatus)
        If ballStatus(id) = 1 Then
            baX(id, id2) = round(currentball(id).x, 2)
            baY(id, id2) = round(currentball(id).y, 2)
            baZ(id, id2) = round(currentball(id).z, 2)
            bVx(id, id2) = round(currentball(id).velx, 2)
            bVy(id, id2) = round(currentball(id).vely, 2)
            TotalVel(id, id2) = (bVx(id, id2) ^2 + bVy(id, id2) ^2)
            If TotalVel(id, id2) > TotalVel(0, 0) Then TotalVel(0, 0) = int(TotalVel(id, id2) )
        End If
    Next

    id3 = id2:B2 = 2:B1 = 1
    Do
        If ballStatus(B1) = 1 and ballStatus(B2) = 1 Then
            bDistance = int((TotalVel(B1, id3) + TotalVel(B2, id3) ) ^(1.04 * (CurrentBall(B1).radius + CurrentBall(B2).radius) / 50) )
            If((baX(B1, id3) - baX(B2, id3) ) ^2 + (baY(B1, id3) - baY(B2, id3) ) ^2) < (2800 * ((CurrentBall(B1).radius + CurrentBall(B2).radius) / 50) ^2) + bDistance Then
                If ABS(baZ(B1, id3) - baZ(B2, id3) ) < (CurrentBall(B1).radius + CurrentBall(B2).radius) Then collide B1, B2:Exit Sub
            End If
        End If
        B1 = B1 + 1
        If B1 = ubound(ballstatus) Then Exit Do
        If B1 >= B2 then B1 = 1:B2 = B2 + 1
    Loop

    If ballStatus(0) <= 1 Then XYdata.enabled = False

    If XYdata.interval >= 40 Then coff = True:XYdata.enabled = False
    If Timer > xyTime * 3 Then coff = True:XYdata.enabled = False
    If Timer > xyTime Then XYdata.interval = XYdata.interval + 1
End Sub

'Calculate the collision force and play sound
Dim cTime, cb1, cb2, avgBallx, cAngle, bAngle1, bAngle2

Sub Collide(cb1, cb2)
    If TotalVel(0, 0) / 1.8 > cFactor Then cFactor = int(TotalVel(0, 0) / 1.8)
    avgBallx = (bvX(cb2, 1) + bvX(cb2, 2) + bvX(cb2, 3) + bvX(cb2, 4) ) / 4
    If avgBallx < bvX(cb2, id2) + .1 and avgBallx > bvX(cb2, id2) -.1 Then
        If ABS(TotalVel(cb1, id2) - TotalVel(cb2, id2) ) < .000005 Then Exit Sub
    End If
    If Timer < cTime Then Exit Sub
    cTime = Timer + .1
    GetAngle baX(cb1, id3) - baX(cb2, id3), baY(cb1, id3) - baY(cb2, id3), cAngle
    id3 = id3 - 1:If id3 = 0 Then id3 = 4
    GetAngle bVx(cb1, id3), bVy(cb1, id3), bAngle1
    GetAngle bVx(cb2, id3), bVy(cb2, id3), bAngle2
    cForce = Cint((abs(TotalVel(cb1, id3) * Cos(cAngle-bAngle1) ) + abs(TotalVel(cb2, id3) * Cos(cAngle-bAngle2) ) ) )
    If cForce < 4 Then Exit Sub
    cForce = Cint((cForce) / (cFactor / nosf) )
    If cForce > nosf-1 Then cForce = nosf-1
    PlaySound "ball_ball_hit", 0, cForce, Pan(CurrentBall(b1) ), 0, 0, 1, 0, AudioFade(ActiveBall)
End Sub

' Get angle
Dim Xin, Yin, rAngle, Radit, wAngle, Pi
Pi = Round(4 * Atn(1), 6) '3.1415926535897932384626433832795

Sub GetAngle(Xin, Yin, wAngle)
    If Sgn(Xin) = 0 Then
        If Sgn(Yin) = 1 Then rAngle = 3 * Pi / 2 Else rAngle = Pi / 2
        If Sgn(Yin) = 0 Then rAngle = 0
        Else
            rAngle = atn(- Yin / Xin)
    End If
    If sgn(Xin) = -1 Then Radit = Pi Else Radit = 0
    If sgn(Xin) = 1 and sgn(Yin) = 1 Then Radit = 2 * Pi
    wAngle = round((Radit + rAngle), 4)
End Sub

'Extra Sounds

Sub Gate_Hit()
	PlaySoundAtVol "GateFast", gate, VolGates
End Sub

Sub Gate2_Hit()
	PlaySoundAtVol "Gate", gate2, VolGates
End Sub

Sub Gate3_Hit()
	PlaySoundAtVol "Gate", gate3, VolGates
End Sub

Sub Gate4_Hit()
	PlaySoundAtVol "GateFast", gate4, VolGates
End Sub

Sub Trigger8_Hit()
	PlaySoundAtVol "BallHit", ActiveBall, 1
End Sub

Sub Trigger9_Hit()
	PlaySoundAtVol "BallHit", ActiveBall, 1
End Sub

Sub Trigger10_Hit()
	PlaySoundAtVol "top_lane", ActiveBall, 1
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

Sub PlaySoundAtVol(sound, tableobj, Volum)
  PlaySound sound, 1, Volum, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
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
  Vol = Csng(BallVel(ball) ^2 / VolDiv)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
  Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
  BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function


' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


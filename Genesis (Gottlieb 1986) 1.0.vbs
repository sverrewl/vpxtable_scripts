'Genesis - Gottlieb 1986
'VP912 table by jpsalas
'3D VP10 conversion by nFozzy

'Version 1.11

'DOF additions by arngrim

' Thalamus 2018-07-23
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' No special SSF tweaks yet.


Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Dim UseVPMDMD
Dim DesktopMode:DesktopMode = Genesis.ShowDT

If DesktopMode = True Then 'Show Desktop components
  Ramp10.visible = 1
  Ramp5.visible = 1
  Ramp37.visible = 1
  Ramp38.visible = 1
  Prim_Sidewalls.visible=1
Else
  Ramp10.visible = 0
  Ramp5.visible = 0
  Ramp37.visible = 0
  Ramp38.visible = 0
  Prim_Sidewalls.visible=0
End if

UseVPMDMD = 1


dim fullscreendisplay
''===================
'\\\\\\\\\\\\\\\\\\\\\
'fullscreen display options
'2 = fullscreen display (best for single screen setups)
'1 = movable pinmame DMD window
'0 = No fullscreen display (use B2S instead)

fullscreendisplay = 0

'/////////////////////
'====================

dim Dropfix

'Set this to 1 if you are using VP10 1.0 and the drop targets aren't resetting properly
Dropfix = 0



 'setup display
dim xx
If DesktopMode then
	For each xx in Display:xx.X = xx.X - 150: xx.Y = xx.Y - 400: xx.rotX = -55: xx.height = xx.height + 320: Next
	For each xx in Display2:xx.Y = xx.y - 20: xx.X = xx.X - 6: xx.height = xx.height - 30: Next
end if


'Load VPM and scripts

LoadVPM "01560000", "sys80.vbs", 3.36

Const cGameName = "genesis"
Const UseSolenoids = 2
Const UseLamps = 0

'Standard sounds
Const SSolenoidOn = "Solenoid"    'Solenoid activates
Const SSolenoidOff = ""           'Solenoid deactivates
Const SCoin = "Coin"              'Coin inserted

Dim bsTrough, bsArmsLock, bsLegsLock, dtM', plungerIM
Dim x, Balls', bump1, bump2, bump3, bump4
Dim Last12, Current12, Last13, Current13, Last14, Current14

Set MotorCallback = GetRef("RollingUpdate") 'realtime updates - rolling sound

'**********
'Table Init
'**********

Sub Genesis_Init
	vpmInit Me

	With Controller
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
		.SplashInfoLine = "Genesis (Gottlieb 1986)" & vbNewLine & "VP9 table by JPSalas"
'		.Games(cGameName).Settings.Value("dmd_red") = 0
'		.Games(cGameName).Settings.Value("dmd_green") = 128
'		.Games(cGameName).Settings.Value("dmd_blue") = 255
'		.Games(cGameName).Settings.Value("rol") = 0
		.HandleKeyboard = 0
		.ShowTitle = 0
'		.ShowDMDOnly = 1
		.ShowFrame = 0
		.HandleMechanics = 0
		.Hidden = 1
		'.SetDisplayPosition 0,0,GetPlayerHWnd 'if you can't see the DMD then uncomment this line
		On Error Resume Next
		Controller.SolMask(0) = 0
		vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the timer to renable all the solenoids after 2 seconds
		Controller.Run GetPlayerHwnd
		If Err Then MsgBox Err.Description
		On Error Goto 0
	End With

	' Nudging
	vpmNudge.TiltSwitch = 57 'swTilt
	vpmNudge.Sensitivity = 5
	vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, Bumper4, LeftSlingshot, RightSlingshot)

	'Saucers Declaration
	Set bsArmsLock = New cvpmSaucer
	with bsArmsLock
		.InitKicker Armslock, 43, 185, 8, 8	' LeftKickout '160, 5 ,8
		.InitExitVariance 1, 0
		.InitSounds "kicker_enter", SoundFX("Solenoid",DOFContactors), SoundFX("Popper",DOFContactors)
		.createevents "bsArmsLock", Armslock
	end with

	Set bsLegsLock = New cvpmSaucer
	with bsLegsLock
		.InitKicker Legslock, 73, 346, 18, 8	' RightKickout '320, 16, 20
		.InitExitVariance 2, 2
		.InitSounds "kicker_enter", SoundFX("Solenoid",DOFContactors), SoundFX("Popper",DOFContactors)
		.createevents "bsLegsLock", LegsLock
	end with


	Set dtM = New cvpmDropTarget
	dtM.InitDrop Array(dt1, dt2, dt3), Array(41, 51, 61)
	dtM.InitSnd SoundFX("droptarget",DOFContactors), SoundFX("resetdrop",DOFContactors)
	dtM.CreateEvents "dtM"


	'Trough Declaration
	Set bsTrough = New cvpmBallStack
	bsTrough.InitSw 55, 0, 74, 0, 0, 0, 0, 0
	bsTrough.InitKick Ballrelease, 80, 6
	bsTrough.InitExitSnd SoundFX("ballrel",DOFContactors), SoundFX("Solenoid",DOFContactors)
	bsTrough.Balls = 2

    ' New style Trough that didn't quite work out
'    Set bsTrough = New cvpmTrough
'    With bsTrough
'        .size = 2
'        .initSwitches Array(55, 74)
'        .Initexit BallRelease, 80, 6
'        .InitEntrySounds "drain", "Solenoid", "Solenoid"
'        .InitExitSounds "Solenoid", "ballrel"
'        .Balls = 2
'    End With

	'Init Target Walls animation
	RightKick.IsDropped = 1:LeftKick.IsDropped = 1
	RightKick2.IsDropped = 0:LeftKick2.IsDropped = 0

	'Init Robot Lights
	StopRobotLights

	'Other variables
	Last12 = 0
	Current12 = 0
	Last13 = 0
	Current13 = 0
	Last14 = 0
	Current14 = 0

	PinMAMETimer.Interval = PinMAMEInterval
	PinMAMETimer.Enabled = 1
	'StartShake

	SolGI 0:SolLeft 0:SolRight 0

	'display Option
	if not DesktopMode and fullscreendisplay <> 2 then
	For each xx in Display:xx.visible = 0: Next
	Displaytimer.enabled = 0
	end If

End Sub

Sub Genesis_Paused:Controller.Pause = 1:End Sub
Sub Genesis_unPaused:Controller.Pause = 0:End Sub

' keys

Sub Genesis_KeyDown(ByVal keycode)
	If keycode = PlungerKey Then Plunger.Pullback
	If keycode = LeftTiltKey Then PlaySound SoundFX("nudge_left",0)
	If keycode = RightTiltKey Then PlaySound SoundFX("nudge_right",0)
	If keycode = CenterTiltKey Then PlaySound SoundFX("nudge_forward",0)
	If vpmKeyDown(keycode) Then Exit Sub
	If keycode = KeyRules Then Rules
	If keycode=31 then 'test debug
'	kicker1.createball
'	kicker1.kick -10, 35
		'l13.state=1
'	gi1.state = 0
'	Rubber_Straightb10.size_x = 85
'	Rubber_Straightb8.size_x = 85
'	Rubber_Straightb14.size_x = 85
'	rubberanim.enabled = 1
	End If
End Sub


Sub Genesis_KeyUp(ByVal keycode)
	If keycode = PlungerKey Then
		Plunger.Fire
		If(BallinPlunger = 1) then 'the ball is in the plunger lane
			PlaySound "Plunger2"
		else
			PlaySound "Plunger"
		end if
	End If
'	If keycode = LeftTiltKey Then PlaySound "nudge_left"
	If vpmKeyUp(KeyCode) Then Exit Sub
'	If keycode=31 then 'light test
'		l13.state=0
'	End If
End Sub

'***********************
'JP's Alpha Ramp Plunger
'***********************
Dim BallinPlunger

Sub swPlunger_Hit:BallinPlunger = 1:End Sub                            'in this sub you may add a switch, for example Controller.Switch(14) = 1

Sub swPlunger_UnHit:BallinPlunger = 0:End Sub                          'in this sub you may add a switch, for example Controller.Switch(14) = 0


'*******************
'Solenoids Callback
'*******************

'SolCallback(1) = "" 'Varitarget ??
SolCallback(2) = "bsLegsLock.SolOut"
SolCallback(4) = "SolLeft"
SolCallback(5) = "bsArmsLock.SolOut"
SolCallback(6) = "DropDelaysub"
'SolCallback(6) = "dtM.SolDropUp"
SolCallback(7) = "SolRight"
SolCallback(8) = "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
'SolCallback(8) = "VpmSolSound""knocker"","
SolCallback(9) = "SolOuthole"
SolCallback(10) = "SolGI"

'dim Drop targets
dim drop1, drop2, drop3
drop1 = dt1.isdropped
drop2 = dt2.isdropped
drop3 = dt3.isdropped



'Drop Delay
Sub DropDelaysub(enabled)
	If Dropfix = 1 then
		DropDelay.Enabled = 1
	Else
		dtM.DropSol_On
		drop1 = 0
		drop2 = 0
		drop3 = 0
		updateGI
	End If
End Sub

Sub DropDelay_Timer()
	dtM.DropSol_On
	me.interval = 20
	me.enabled = 0
	drop1 = 0
	drop2 = 0
	drop3 = 0
	updateGI
End Sub

'Solenoids Subs

Sub SolGI(Enabled)
	If Enabled Then
'	textbox1.text = Enabled
		gi1.state=1:gi2.state=1:gi3.state=1:gi4.state=1:gi5.state=1:gi6.state=1:gi7.state=1:gi8.state=1:gi9.state=1:gi10.state=1:gi11.state=1:gi12.state=1:gi13.state=1:gi_ambient.state=1:gi15.state=1
		UpdateGi
	Else
'	textbox1.text = Enabled
		gi1.state=0:gi2.state=0:gi3.state=0:gi4.state=0:gi5.state=0:gi6.state=0:gi7.state=0:gi8.state=0:gi9.state=0:gi10.state=0:gi11.state=0:gi12.state=0:gi13.state=0:gi_ambient.state=0:gi15.state=0
		UpdateGi
	End If
End Sub

Sub UpdateGI
gi14.state = gi7.state
if drop1 = 1 then gi14_1.state = gi7.state else gi14_1.state = 0 end if
if drop2 = 1 then gi14_2.state = gi7.state else gi14_2.state = 0 end if
if drop3 = 1 then gi14_3.state = gi7.state else gi14_3.state = 0 end if

end sub


'Sub timer1_timer() 'check DTs Debug
'textbox4.text = dt1.isdropped & " " & dt2.isdropped & " " & dt3.isdropped
'textbox5.text = drop1 & " " & drop2 & " " & drop3
'updategi

'End Sub

Sub SolLeft(Enabled)
	If Enabled Then
		Fl2.state = 2:fl3.state = 2:playsound "lswitch", 0, 0.01, 0, 0.1
'		textbox1.text = "ON"
	Else
		Fl2.state=0:fl3.state=0
'		textbox1.text = "OFF"
	End If
End Sub



Sub SolRight(Enabled)
	If Enabled Then
		fr2.state = 2
		fr3.State = 2
		playsound "lswitch", 0, 0.001, 0
'		textbox2.text = "ON"
	Else
		fr2.state = 0
		fr3.State = 0
'		textbox2.text = "OFF"
	End If
End Sub

Sub SolOuthole(enabled)
	if enabled then
		bsTrough.EntrySol_On
'		bsTrough.ExitSol_On
	end if
End Sub

'*************
'Robots Lights
'*************

Dim RobotLightStep, RobotLightsOn, EndIt

RobotLightStep = 0:RobotLightsOn = 0

Sub StartRobotLights
'	If Robotlightson=1 then
'		Exit Sub
'	End If
'	'light2.state=1
'	RobotLightStep = 0
'    RobotLightsOn = 1
'	RobotLights.Enabled = 1
	ll1.state=2:rl1.state=2:ll2.state=2:rl2.state=2:ll3.state=2:rl3.state=2:ll4.state=2:rl4.state=2:ll5.state=2:rl5.state=2
	cl1.state = 2: cl2.state = 2: cl3.state = 2: cl4.state = 2: cl5.state = 2
End Sub

'Sub maybestoprobotlights	' I think this prevents the lightshow from ending early during the robot reveal sequence
'	If CurrentRot=0 then LightSeqTimer.Enabled=1 End If	'lightseqtimer judges if the lights should be on or not..
'	If CurrentRot<0 then
'	StopRobotLights
'	End If
'end Sub

'the way I scripted this makes my head hurt

Sub StopRobotLights
'	RobotLightStep=65
	'light2.state=0
	ll1.state=0:rl1.state=0:ll2.state=0:rl2.state=0:ll3.state=0:rl3.state=0:ll4.state=0:rl4.state=0:ll5.state=0:rl5.state=0
	cl1.state = 0: cl2.state = 0: cl3.state = 0: cl4.state = 0: cl5.state = 0
'    RobotLightsOn = 0
End Sub

'Sub RobotLights_Timer	'replaced by blink pattern 'interval was 70
'	Select Case RobotLightStep
'		Case 0:Ll1.State=1:Rl1.state=1
'		Case 1:Ll2.state=1:Rl2.state=1
'		Case 2:ll1.state=0:rl1.state=0:ll3.state=1:rl3.state=1
'		Case 3:ll2.state=0:rl2.state=0:ll4.state=1:rl4.state=1
'		Case 4:ll3.state=0:rl3.state=0:ll5.state=1:rl5.state=1
'		Case 5:ll4.state=0:rl4.state=0
'		Case 6:ll5.state=0:rl5.state=0
'		Case 66:ll1.state=0:rl1.state=0:ll2.state=0:rl2.state=0:ll3.state=0:rl3.state=0:ll4.state=0:rl4.state=0:ll5.state=0:rl5.state=0
'		Case 67:RobotLights.Enabled=0:RobotLightStep=1
'	End Select
'	RobotLightStep = RobotLightStep + 1
'	If RobotLightStep = 7 Then RobotLightStep = 0
'End Sub

'****************
' Robot Animation

'****************
dim StartRotation
dim EndRotation
Dim CurrentRot

StartRotation=0
EndRotation=-360

Sub StartRobotAnimation
	CurrentRot=0
	'Light1.State=1 'Light1 + Light 2 are additional ambient lighting
	SpinTimer.Enabled=1
End Sub

Sub SpinTimer_Timer()
	If Currentrot=EndRotation then
		currentrot=StartRotation 'back to 0
		me.Enabled=0
		'Light1.State=0
		Exit Sub
	End If
	'If CurrentRot< -180 then Light1.State=0 End If
	If currentrot> EndRotation then currentrot=currentrot-0.28 End If
	Goldy.roty=CurrentRot
	Goldy2.roty=CurrentRot
End Sub

'**************
' Flipper Subs
'**************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
	Controller.Switch(75) = ABS(enabled)
	If Enabled Then
		'PlaySound "flipperup":LeftFlipper.RotateToEnd
		LeftFlipper.RotateToEnd
		If LeftFlipper.CurrentAngle<80 Then		'If weak flip...
		PlaySound SoundFx("FlipperUp",DOFContactors),0,0.1,-0.05		'Play a Weaker Flip Sound
		Else PlaySound SoundFx("FlipperUp",DOFContactors),0,1,-0.05
		End If
	Else
		PlaySound SoundFx("Flipperdown",DOFContactors),0,0.05,-0.02:LeftFlipper.RotateToStart
	End If
End Sub

Sub SolRFlipper(Enabled)
	Controller.Switch(75) = ABS(enabled)
	If Enabled Then
		'PlaySound "flipperup":RightFlipper.RotateToEnd
		RightFlipper.RotateToEnd
		If RightFlipper.CurrentAngle > (80*-1) Then	'If weak flip...
		PlaySound SoundFx("FlipperUp",DOFContactors),0,0.1,0.05			'Play a Weaker Flip Sound
		Else PlaySound SoundFx("FlipperUp",DOFContactors),0,1,0.05
		End If
	Else
		PlaySound SoundFx("Flipperdown",DOFContactors),0,0.05,0.02:RightFlipper.RotateToStart
	End If
End Sub

'SoundFx("Flipperdown",DOFContactors),

Sub LeftFlipper_Collide(parm)
	PlaySound "rubber_flipper"
End Sub

Sub RightFlipper_Collide(parm)
	PlaySound "rubber_flipper"
End Sub

'Set MotorCallback = GetRef("RealTimeUpdates")

Sub FlipperTimer_Timer()
	UpdateLeftFlipperLogo
	UpdateRightFlipperLogo
End Sub

Sub UpdateLeftFlipperLogo()
	LFLogo.RotY = LeftFlipper.CurrentAngle
End Sub
Sub UpdateRightFlipperLogo()
	RFLogo.RotY = RightFlipper.CurrentAngle
End Sub


'************
'Varitarget
'omg
'************

'Varitarget primitive version:
'range: all the way forward rotX -9 & -128
'all the way back rotx 12 & -20
dim variball

dim Y 'variangle
dim CFA 'FlipperAngle
CFA = Flipper2.currentangle

Y = -9

sub Invari_hit()
	me.timerenabled = 0
	me.timerinterval = 500
	VariChecker.enabled = 1
end sub

sub Invari_unhit()
	me.timerinterval = 500
	me.timerenabled = 1
'	VariChecker.enabled = 0

end sub

sub invari_timer()
	if y < -8.5 then varichecker.enabled = 0: me.timerenabled = 0 end if
end sub


'sub 	set variball = BallcntOver
Sub Varichecker_timer()
	CFA = Flipper2.currentangle
	y = ((7 * CFA) / 36) + (143 / 9)
	Varitarget.rotX = Y
'	textbox2.text = Y
'	textbox3.text = flipper2.currentangle & flipper2.startangle & flipper2.endangle


'I am bad at maths


End sub

dim v1, v2, v3, v4
v1 = 0:v2 = 0:v3=0:v4=0
'Varitimer is 200ms

Sub Varitarget1_Hit
	If ActiveBall.VelY <0 Then
		Controller.Switch(40) = 1
		V1 = 1
		V2 = 0
	end if
End Sub

Sub Varitarget1_UnHit
	If ActiveBall.VelY> 0 Then VariTimer.Enabled = 1
End Sub


Sub Varitarget2_Hit
	If ActiveBall.VelY <0 Then
		Controller.Switch(50) = 1
		V2 = 1
		V3 = 0
	end if
End Sub

Sub Varitarget3_Hit
	If ActiveBall.VelY <0 Then
		Controller.Switch(60) = 1
		V3 = 1
		V4 = 0
	end if
End Sub

Sub Varitarget4_Hit
	Controller.Switch(70) = 1
End Sub

Sub VariTimer_Timer
	If v4 = 0 Then
		Controller.Switch(70) = 0
		V4 = 1
		V3 = 0
		Exit Sub
	End If

	If v3 = 0 Then
		Controller.Switch(60) = 0
		V3 = 1
		V2 = 0
		Exit Sub
	End If

	If v2 = 0 Then
		Controller.Switch(50) = 0
		V2 = 1
		V1 = 0
		Exit Sub
	End If

	Controller.Switch(40) = 0
	VariTimer.Enabled = 0
End Sub


'Triggers

Sub LeftKick_Timer:LeftKick.TimerEnabled = 0:LeftKick2.IsDropped = 0:LeftKick.IsDropped = 1:End Sub

Sub LeftKick2_Slingshot():vpmTimer.PulseSw(45):LeftKick.IsDropped = 0:LeftKick2.IsDropped = 1:LeftKick.TimerEnabled = 1:PlaySound SoundFx("slingshot",DOFContactors),0, 0.8, -0.08, 0.05:End Sub 'PlaySound "name",loopcount,volume,pan,randompitch

Sub RightKick_Timer:RightKick.TimerEnabled = 0:RightKick2.IsDropped = 0:RightKick.IsDropped = 1:End Sub

Sub RightKick2_Slingshot():vpmTimer.PulseSw(65):RightKick.IsDropped = 0:RightKick2.IsDropped = 1:RightKick.TimerEnabled = 1:PlaySound SoundFx("slingshot",DOFContactors),0, 0.8, 0.08, 0.05:End Sub 'PlaySound "name",loopcount,volume,pan,randompitch

Sub LeftLane_Hit():Playsound "sensor":Controller.switch(53) = 1:End Sub
Sub LeftLane_UnHit():Controller.switch(53) = 0:End Sub


Sub Spinner_Spin():Playsound "spinner":End Sub

Sub Toplane1_Hit():Playsound "sensor":Controller.switch(42) = 1:End Sub
Sub Toplane1_UnHit():Controller.switch(42) = 0:End Sub
Sub Toplane2_Hit():Playsound "sensor":Controller.switch(52) = 1:End Sub
Sub Toplane2_UnHit():Controller.switch(52) = 0:End Sub
Sub Toplane3_Hit():Playsound "sensor":Controller.switch(62) = 1:End Sub
Sub Toplane3_UnHit():Controller.switch(62) = 0:End Sub

Sub LeftInLane_Hit():Playsound "sensor":Controller.switch(44) = 1:DOF 101, DOFOn:End Sub
Sub LeftInLane_UnHit():Controller.switch(44) = 0:DOF 101, DOFOff:End Sub
Sub LeftOutlane_Hit():Playsound "sensor":Controller.switch(54) = 1:End Sub
Sub LeftOutlane_UnHit():Controller.switch(54) = 0:End Sub
Sub RightInlane_Hit():Playsound "sensor":Controller.switch(44) = 1:DOF 102, DOFOn:End Sub
Sub RightInlane_UnHit():Controller.switch(44) = 0:DOF 102, DOFOff:End Sub
Sub RightOutlane_Hit():Playsound "sensor":Controller.switch(64) = 1:End Sub
Sub RightOutlane_UnHit():Controller.switch(64) = 0:End Sub

'One Way Switch
Dim TopDown
TopDown=False

Sub OneWaySwitch_Hit()
	OneWayTimer.Enabled=1
	TopDown=True
End Sub

Sub OneWayTimer_Timer()
	TopDown=False
	OneWayTimer.Enabled=0
End Sub

Sub sw63_Hit()
	If TopDown=False then Controller.switch(63) = 1':playsound "Diverter"
End Sub

Sub sw63_UnHit() 'extra switch juice
	me.timerenabled=1
End Sub

Sub sw63_Timer() 'extra switch juice
	Controller.switch(63) = 0
	me.Timerenabled=0
End Sub

'Drop-Targets
Sub dt1_hit():dtM.Hit 1:End Sub
Sub dt1_dropped():drop1 = 1:updategi:End Sub

Sub dt2_hit():dtM.Hit 2:End Sub
Sub dt2_dropped():drop2 = 1:updategi:End Sub

Sub dt3_hit():dtM.Hit 3:End Sub
Sub dt3_dropped():drop3 = 1:updategi:End Sub


'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub LeftSlingShot_Slingshot
	vpmTimer.PulseSw(72)
    PlaySound SoundFXDOF("slingshot",112,DOFPulse,DOFContactors),0,1,-0.05,0.05
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -42
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -25
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0':gi3.State = 1:Gi4.State = 1
    End Select
    LStep = LStep + 1
End Sub

'Right Slingshot

Sub RightSlingShot_Slingshot
	vpmTimer.PulseSw(72)
    PlaySound SoundFXDOF("slingshot",113,DOFPulse,DOFContactors), 0, 1, 0.05, 0.05
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -42
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -25
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0':gi1.State = 1:Gi2.State = 1
    End Select
    RStep = RStep + 1
End Sub

'-==============================

'Bumpers
'Sub Bumper1_Hit():vpmTimer.PulseSw(71):PlaySound "bumper1":End Sub
'
'Sub Bumper2_Hit():vpmTimer.PulseSw(71):PlaySound "bumper2":End Sub
'
'Sub Bumper3_Hit():vpmTimer.PulseSw(71):PlaySound "bumper3":End Sub
'
'Sub Bumper4_Hit():vpmTimer.PulseSw(71):PlaySound "bumper2":End Sub
'

Sub Bumper1_Hit()
vpmTimer.PulseSw(71)
PlaySound SoundFXDOF("Bumper1",103,DOFPulse,DOFContactors)
dim finalspeed
finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
'textbox1.text = finalspeed
If finalspeed > 13.5 and gi7.state = 1 then gi1.state = 2 end if
'If gi7.state = 1 then gi1.state = 2 end if
me.timerenabled = 1
End Sub

Sub Bumper1_Timer()
gi1.state = gi7.state
me.timerenabled = 0
end sub

Sub Bumper2_Hit()
vpmTimer.PulseSw(71)
PlaySound SoundFXDOF("Bumper2",104,DOFPulse,DOFContactors)
End Sub

Sub Bumper3_Hit()
vpmTimer.PulseSw(71)
PlaySound SoundFXDOF("Bumper2",105,DOFPulse,DOFContactors)
dim finalspeed
finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
If finalspeed > 13.5 and gi7.state = 1 then gi3.state = 2 end if
'If gi7.state = 1 then gi3.state = 2 end if
me.timerenabled = 1
End Sub

Sub Bumper3_Timer()
gi3.state = gi7.state
me.timerenabled = 0
end sub

Sub Bumper4_Hit()
vpmTimer.PulseSw(71)
PlaySound SoundFXDOF("Bumper2",106,DOFPulse,DOFContactors)
End Sub


'Outhole

Sub Drain_Hit():Playsound "drain":bsTrough.AddBall Me:End Sub
'Sub Drain_Hit():Playsound "drain":me.destroyball:End Sub	'Debug

'Ramps Top

Sub Ramp1_Hit():PlaySound "Ramp_Hit1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0:End Sub 'PlaySound "name",loopcount,volume,pan,randompitch
Sub Ramp2_Hit():PlaySound "Ramp_Hit1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0:End Sub

Sub Ramp3_Hit():PlaySound "Ramp_Hit1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0:End Sub 'PlaySound "name",loopcount,volume,pan,randompitch
Sub Ramp4_Hit():PlaySound "Ramp_Hit1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0:End Sub

Sub RHelp1_Hit():Playsound "ramp_hit3",0,1,-0.06:DOF 107, DOFPulse:End Sub
Sub RHelp2_Hit():Playsound "ramp_hit3",0,1,0.06:DOF 108, DOFPulse:End Sub

' Holes

Sub ArmsLock_Hit:PlaySound "kicker_enter":vpmTimer.PulseSw 43:End Sub
Sub LegsLock_Hit:PlaySound "kicker_enter":vpmTimer.PulseSw 73:End Sub

'***************
' Special lights
'***************

Set LampCallback = GetRef("UpdateMultipleLamps")

Sub UpdateMultipleLamps

	' Robot animation

	Current12 = l12.State
	if Current12 <> Last12 Then
'		if RobotLightsOn = 1 then
'			StartRobotAnimation
'			StartRobotLights
'		end if
			StartRobotAnimation
	end if
	Last12 = Current12

	' Robot lights

'	Current13 = l13.State
'	if Current13 <> Last13 Then
'		if Current13 = 1 then
'			StartRobotLights
'			LightSeqTimer.Enabled=0
'			LightSeqTimer.Interval=1000
'		else								'StopRobotLights
'		maybestoprobotlights
'		'If CurrentRot<0 then StopRobotLights else LightSeqTimer.enabled=1
'		end if
'	end if
'	Last13 = Current13


	'Check BallTrough
	Current14 = l14.State
	if Current14 <> Last14 Then
		if Current14 = 1 then
			if bsTrough.Balls then bsTrough.ExitSol_On
		end if
	end if
	Last14 = Current14
End Sub

	'Robot Light Sequence Protector
'Sub	LightSeqTimer_Timer()
'	StopRobotLights
'	Me.Enabled=0
'End Sub


'================VP10 Fading Lamps Script

Dim LampState(200), FadingLevel(200)
Dim FlashSpeedUp(200), FlashSpeedDown(200), FlashMin(200), FlashMax(200), FlashLevel(200)

InitLamps()
LampTimer.Interval = 10
LampTimer.Enabled = 1


Sub LampTimer_Timer()
    Dim chgLamp, num, chg, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp) Then
        For ii = 0 To UBound(chgLamp)
            LampState(chgLamp(ii, 0) ) = chgLamp(ii, 1)       'keep the real state in an array
            FadingLevel(chgLamp(ii, 0) ) = chgLamp(ii, 1) + 4 'actual fading step
        Next

    End If

    UpdateLamps
End Sub

Sub UpdateLamps
	NFadeL 2, l2 'FadeL
	NFadeL 3, l3 'FadeL
	NFadeL 5, l5 'FadeL
	NFadeL 6, l6 'FadeL
	NFadeL 7, l7 'FadeL
	NFadeL 8, l8 'FadeL
	NFadeL 9, l9 'FadeL
	NFadeL 10, l10 'FadeL
	NFadeL 11, l11 'FadeL
	NFadeL 12, l12 'start robot animation

'	NFadeL 13, l13 'start robot flash lights
	NFadeLS 13, l13 'start robot flash lights

	NFadeL 14, l14 'check balltrough
	NFadeL 15, l15 'FadeL
	NFadeL 16, l16 'FadeL
	NFadeL 17, l17 'FadeL
	NFadeL 18, l18 'FadeL
	NFadeL 19, l19 'FadeL

'	NFadeL 20, l20 'FadeL
'	NFadeL 21, l21 'FadeL
'	NFadeL 22, l22 'FadeL
'	NFadeL 23, l23 'FadeL
	NFadeL 24, l24 'FadeL

	NFadeLm 20, l20 'FadeL
	NFadeLm 21, l21 'FadeL
	NFadeLm 22, l22 'FadeL
	NFadeLm 23, l23 'FadeL
	NFadeLm 20, l20a 'FadeL
	NFadeLm 21, l21a 'FadeL
	NFadeLm 22, l22a 'FadeL
	NFadeLm 23, l23a 'FadeL

'	NFadeL 25, l25 'FadeL
'	NFadeL 26, l26 'FadeL
'	NFadeL 27, l27 'FadeL
	NFadeLwf2 25, l25, l25F, l25F2 'FadeL
	NFadeLwf2 26, l26, l26F, l26F2 'FadeL
	NFadeLwf2 27, l27, l27F, l27F2 'FadeL

	Flash 28, A_RMS
	Flash 29, AR_MS
	Flash 30, ARM_S
	Flash 31, ARMS_
	Flash 32, B_RAIN
	Flash 33, BR_AIN
	Flash 34, BRA_IN
	Flash 35, BRAI_N
	Flash 36, BRAIN_
	Flash 37, B_ODY
	Flash 38, BO_DY
	Flash 39, BOD_Y
	Flash 40, BODY_
	Flash 41, L_EGS
	Flash 42, LE_GS
	Flash 43, LEG_S
	Flash 44, LEGS_

	NFadeL 45, l45 'FadeL
	NFadeL 46, l46 'FadeL
	NFadeLm 47, l47b 'FadeLm
	NFadeL 47, l47 'FadeL
	NFadeLm 51, l51b'FadeLm
	NFadeL 51, l51 'FadeL
End Sub

Sub InitLamps()
    Dim x
    For x = 0 to 200
        LampState(x) = 0        ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4      ' used to track the fading state
        FlashSpeedUp(x) = 0.2   ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.1 ' slower speed when turning off the flasher
        FlashMax(x) = 1         ' the maximum value when on, usually 1
        FlashMin(x) = 0.05         ' the minimum value when off, usually 0
        FlashLevel(x) = 0       ' the intensity of the flashers, usually from 0 to 1
    Next
End Sub

Sub AllLampsOff
    Dim x
    For x = 0 to 200
        SetLamp x, 0
    Next
End Sub

Sub SetLamp(nr, value)
    If value <> LampState(nr) Then
        LampState(nr) = abs(value)
        FadingLevel(nr) = abs(value) + 4
    End If
End Sub

'Walls

' Lights: used for VP10 standard lights, the fading is handled by VP itself

Sub NFadeL(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.state = 0:FadingLevel(nr) = 0
        Case 5:object.state = 1:FadingLevel(nr) = 1
    End Select
End Sub

'Special - Lights Robot Lights
Sub NFadeLS(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.state = 0:FadingLevel(nr) = 0:LightSeqTimer.enabled = 1
        Case 5:object.state = 1:FadingLevel(nr) = 1:StartRobotLights:LightSeqTimer.interval = 300
    End Select
End Sub

Sub LightSeqTimer_Timer()
	StopRobotLights
	me.enabled = 0
end sub

'LightSeqTimer
'StartRobotLights
'StopRobotLights


Sub NFadeLm(nr, object) ' used for multiple lights
    Select Case FadingLevel(nr)
        Case 4:object.state = 0
        Case 5:object.state = 1
    End Select
End Sub

Sub FadeLm(nr, a, b) 'Old
	Select Case LampState(nr)
		Case 2:b.state = 0
		Case 3:b.state = 1
		Case 4:a.state = 0
		Case 5:b.state = 1
		Case 6:a.state = 1
	End Select
End Sub

'Lights, Ramps & Primitives used as 4 step fading lights
'a,b,c,d are the images used from on to off

Sub FadeObj(nr, object, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 4:object.image = b:FadingLevel(nr) = 6                   'fading to off...
        Case 5:object.image = a:FadingLevel(nr) = 1                   'ON
        Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1             'wait
        Case 9:object.image = c:FadingLevel(nr) = FadingLevel(nr) + 1 'fading...
        Case 10, 11, 12:FadingLevel(nr) = FadingLevel(nr) + 1         'wait
        Case 13:object.image = d:FadingLevel(nr) = 0                  'Off
    End Select
End Sub

Sub FadeObjm(nr, object, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 4:object.image = b
        Case 5:object.image = a
        Case 9:object.image = c
        Case 13:object.image = d
    End Select
End Sub

Sub NFadeObj(nr, object, a, b)
    Select Case FadingLevel(nr)
        Case 4:object.image = b:FadingLevel(nr) = 0 'off
        Case 5:object.image = a:FadingLevel(nr) = 1 'on
    End Select
End Sub

Sub NFadeObjm(nr, object, a, b)
    Select Case FadingLevel(nr)
        Case 4:object.image = b
        Case 5:object.image = a
    End Select
End Sub

' Flasher objects

Sub Flash(nr, object)
    Select Case FadingLevel(nr)
        Case 4 'off
            FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown(nr)
            If FlashLevel(nr) < FlashMin(nr) Then
                FlashLevel(nr) = FlashMin(nr)
                FadingLevel(nr) = 0 'completely off
            End if
            Object.IntensityScale = FlashLevel(nr)
        Case 5 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
            If FlashLevel(nr) > FlashMax(nr) Then
                FlashLevel(nr) = FlashMax(nr)
                FadingLevel(nr) = 1 'completely on
            End if
            Object.IntensityScale = FlashLevel(nr)
    End Select
        '   Object.IntensityScale = 1
End Sub

Sub Flashm(nr, object) 'multiple flashers, it just sets the flashlevel
    Object.IntensityScale = FlashLevel(nr)
         ' Object.IntensityScale = 1
End Sub

Sub NFadeLwF(nr, object1, object2)
    Select Case FadingLevel(nr)
'		Case 0:If object2.intensityscale > 0 then Object2.intensityscale = object2.intensityscale + (object1.fadespeeddown * -1) *2 end if
'		Case 1:If object2.intensityscale < 1 then Object2.intensityscale = object2.intensityscale + (object1.fadespeedup) *2 end if

		Case 0:object2.intensityscale = 0
		Case 1:object2.intensityscale = 1
        Case 4:object1.state = 0:FadingLevel(nr) = 16
        Case 5:object1.state = 1:FadingLevel(nr) = 6':TextBox4.text = object1.fadespeedup	'to 6
'0.1 up, 0.1 down
		Case 6, 7, 8, 9, 10, 11, 12, 13, 14:If object2.intensityscale < 1 then Object2.intensityscale = object2.intensityscale + 0.1 end if:FadingLevel(nr) = FadingLevel(nr) + 1
		Case 15:If object2.intensityscale < 1 then Object2.intensityscale = object2.intensityscale + 0.1 end if:FadingLevel(nr) = 1':TextBox4.text = "Case 11"
		Case 16, 17, 18, 19, 20, 21, 22, 23, 24:If object2.intensityscale > 0 then Object2.intensityscale = object2.intensityscale + -0.1 end if:FadingLevel(nr) = FadingLevel(nr) + 1
		Case 25:If object2.intensityscale > 0 then Object2.intensityscale = object2.intensityscale + -0.1 end if:FadingLevel(nr) = 0':TextBox4.text = "Case 26"
    End Select
End Sub

Sub NFadeLwF2(nr, object1, object2, object3)	'one light two flashers
    Select Case FadingLevel(nr)

		Case 0:object2.intensityscale = 0:object3.intensityscale = object2.intensityscale
		Case 1:object2.intensityscale = 1:object3.intensityscale = object2.intensityscale
        Case 4:object1.state = 0:FadingLevel(nr) = 16
        Case 5:object1.state = 1:FadingLevel(nr) = 6':TextBox4.text = object1.fadespeedup	'to 6
'0.1 up, 0.1 down
		Case 6, 7, 8, 9, 10, 11, 12, 13, 14:If object2.intensityscale < 1 then Object2.intensityscale = object2.intensityscale + 0.1 end if:object3.intensityscale = object2.intensityscale: FadingLevel(nr) = FadingLevel(nr) + 1
		Case 15:If object2.intensityscale < 1 then Object2.intensityscale = object2.intensityscale + 0.1 end if:FadingLevel(nr) = 1:object3.intensityscale = object2.intensityscale':TextBox4.text = "Case 11"
		Case 16, 17, 18, 19, 20, 21, 22, 23, 24:If object2.intensityscale > 0 then Object2.intensityscale = object2.intensityscale + -0.1 end if:object3.intensityscale = object2.intensityscale :FadingLevel(nr) = FadingLevel(nr) + 1
		Case 25:If object2.intensityscale > 0 then Object2.intensityscale = object2.intensityscale + -0.1 end if:FadingLevel(nr) = 0:object3.intensityscale = object2.intensityscale':TextBox4.text = "Case 26"

	end select
End Sub

' Desktop Objects: Reels & texts (you may also use lights on the desktop)

' Reels

Sub FadeR(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.SetValue 1:FadingLevel(nr) = 6                   'fading to off...
        Case 5:object.SetValue 0:FadingLevel(nr) = 1                   'ON
        Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1              'wait
        Case 9:object.SetValue 2:FadingLevel(nr) = FadingLevel(nr) + 1 'fading...
        Case 10, 11, 12:FadingLevel(nr) = FadingLevel(nr) + 1          'wait
        Case 13:object.SetValue 3:FadingLevel(nr) = 0                  'Off
    End Select
End Sub

Sub FadeRm(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.SetValue 1
        Case 5:object.SetValue 0
        Case 9:object.SetValue 2
        Case 3:object.SetValue 3
    End Select
End Sub

'Texts

Sub NFadeT(nr, object, message)
    Select Case FadingLevel(nr)
        Case 4:object.Text = "":FadingLevel(nr) = 0
        Case 5:object.Text = message:FadingLevel(nr) = 1
    End Select
End Sub

Sub NFadeTm(nr, object, b)
    Select Case FadingLevel(nr)
        Case 4:object.Text = ""
        Case 5:object.Text = message
    End Select
End Sub

'===============================================================

' Extra Sounds


Sub PlasticRamps_Hit (idx)
	PlaySound "flip_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
	PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
	PlaySound "metalhit_thin", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
	PlaySound "metalhit_medium", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Hit (idx) 'Inlanes & shooter lane
	PlaySound "metalhit2", 0, Vol(ActiveBall)*2, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
	PlaySound "gate4", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub RubberBands_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RubberSlings_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 10 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RubberPosts_Hit(idx)
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


'Animated rubbers
'Sub Rubber_Straightb8_hit()
'	Rubber_Straightb8.size_x = 90
'	Rubberanim.enabled = 1
'End Sub

'Sub Rubber_Straightb14_hit()
'	Rubber_Straightb14.size_x = 90
'	Rubberanim.enabled = 1
'End Sub

'Sub Rubber_Straightb10_hit()
'	Rubber_Straightb10.size_x = 90
'	Rubberanim.enabled = 1
'End Sub
'Sub Rubber_Straightb5_hit()
'	Rubber_Straightb5.size_x = 90
'	Rubberanim.enabled = 1
'End Sub

'Sub Rubberanim_timer()
'	Rubber_Straightb8.size_x = 100
'	Rubber_Straightb14.size_x = 100
'	Rubber_Straightb10.size_x = 100
'	Rubber_Straightb5.size_x = 100
'	me.enabled = 0
'End Sub

' Eala's rutine
 Dim Digits(40)
 Digits(0)=Array(f01_1, f02_1, f03_1, f04_1, f05_1, f06_1, f07_1, f08_1, f09_1, f10_1, f11_1, f12_1, f13_1, f14_1, f15_1, f16_1)
 Digits(1)=Array(f01_2, f02_2, f03_2, f04_2, f05_2, f06_2, f07_2, f08_2, f09_2, f10_2, f11_2, f12_2, f13_2, f14_2, f15_2, f16_2)
 Digits(2)=Array(f01_3, f02_3, f03_3, f04_3, f05_3, f06_3, f07_3, f08_3, f09_3, f10_3, f11_3, f12_3, f13_3, f14_3, f15_3, f16_3)
 Digits(3)=Array(f01_4, f02_4, f03_4, f04_4, f05_4, f06_4, f07_4, f08_4, f09_4, f10_4, f11_4, f12_4, f13_4, f14_4, f15_4, f16_4)
 Digits(4)=Array(f01_5, f02_5, f03_5, f04_5, f05_5, f06_5, f07_5, f08_5, f09_5, f10_5, f11_5, f12_5, f13_5, f14_5, f15_5, f16_5)
 Digits(5)=Array(f01_6, f02_6, f03_6, f04_6, f05_6, f06_6, f07_6, f08_6, f09_6, f10_6, f11_6, f12_6, f13_6, f14_6, f15_6, f16_6)
 Digits(6)=Array(f01_7, f02_7, f03_7, f04_7, f05_7, f06_7, f07_7, f08_7, f09_7, f10_7, f11_7, f12_7, f13_7, f14_7, f15_7, f16_7)
 Digits(7)=Array(f01_8, f02_8, f03_8, f04_8, f05_8, f06_8, f07_8, f08_8, f09_8, f10_8, f11_8, f12_8, f13_8, f14_8, f15_8, f16_8)
 Digits(8)=Array(f01_9, f02_9, f03_9, f04_9, f05_9, f06_9, f07_9, f08_9, f09_9, f10_9, f11_9, f12_9, f13_9, f14_9, f15_9, f16_9)
 Digits(9)=Array(f01_10, f02_10, f03_10, f04_10, f05_10, f06_10, f07_10, f08_10, f09_10, f10_10, f11_10, f12_10, f13_10, f14_10, f15_10, f16_10)
 Digits(10)=Array(f01_11, f02_11, f03_11, f04_11, f05_11, f06_11, f07_11, f08_11, f09_11, f10_11, f11_11, f12_11, f13_11, f14_11, f15_11, f16_11)
 Digits(11)=Array(f01_12, f02_12, f03_12, f04_12, f05_12, f06_12, f07_12, f08_12, f09_12, f10_12, f11_12, f12_12, f13_12, f14_12, f15_12, f16_12)
 Digits(12)=Array(f01_13, f02_13, f03_13, f04_13, f05_13, f06_13, f07_13, f08_13, f09_13, f10_13, f11_13, f12_13, f13_13, f14_13, f15_13, f16_13)
 Digits(13)=Array(f01_14, f02_14, f03_14, f04_14, f05_14, f06_14, f07_14, f08_14, f09_14, f10_14, f11_14, f12_14, f13_14, f14_14, f15_14, f16_14)
 Digits(14)=Array(f01_15, f02_15, f03_15, f04_15, f05_15, f06_15, f07_15, f08_15, f09_15, f10_15, f11_15, f12_15, f13_15, f14_15, f15_15, f16_15)
 Digits(15)=Array(f01_16, f02_16, f03_16, f04_16, f05_16, f06_16, f07_16, f08_16, f09_16, f10_16, f11_16, f12_16, f13_16, f14_16, f15_16, f16_16)
 Digits(16)=Array(f01_17, f02_17, f03_17, f04_17, f05_17, f06_17, f07_17, f08_17, f09_17, f10_17, f11_17, f12_17, f13_17, f14_17, f15_17, f16_17)
 Digits(17)=Array(f01_18, f02_18, f03_18, f04_18, f05_18, f06_18, f07_18, f08_18, f09_18, f10_18, f11_18, f12_18, f13_18, f14_18, f15_18, f16_18)
 Digits(18)=Array(f01_19, f02_19, f03_19, f04_19, f05_19, f06_19, f07_19, f08_19, f09_19, f10_19, f11_19, f12_19, f13_19, f14_19, f15_19, f16_19)
 Digits(19)=Array(f01_20, f02_20, f03_20, f04_20, f05_20, f06_20, f07_20, f08_20, f09_20, f10_20, f11_20, f12_20, f13_20, f14_20, f15_20, f16_20)
 Digits(20)=Array(f01_21, f02_21, f03_21, f04_21, f05_21, f06_21, f07_21, f08_21, f09_21, f10_21, f11_21, f12_21, f13_21, f14_21, f15_21, f16_21)
 Digits(21)=Array(f01_22, f02_22, f03_22, f04_22, f05_22, f06_22, f07_22, f08_22, f09_22, f10_22, f11_22, f12_22, f13_22, f14_22, f15_22, f16_22)
 Digits(22)=Array(f01_23, f02_23, f03_23, f04_23, f05_23, f06_23, f07_23, f08_23, f09_23, f10_23, f11_23, f12_23, f13_23, f14_23, f15_23, f16_23)
 Digits(23)=Array(f01_24, f02_24, f03_24, f04_24, f05_24, f06_24, f07_24, f08_24, f09_24, f10_24, f11_24, f12_24, f13_24, f14_24, f15_24, f16_24)
 Digits(24)=Array(f01_25, f02_25, f03_25, f04_25, f05_25, f06_25, f07_25, f08_25, f09_25, f10_25, f11_25, f12_25, f13_25, f14_25, f15_25, f16_25)
 Digits(25)=Array(f01_26, f02_26, f03_26, f04_26, f05_26, f06_26, f07_26, f08_26, f09_26, f10_26, f11_26, f12_26, f13_26, f14_26, f15_26, f16_26)
 Digits(26)=Array(f01_27, f02_27, f03_27, f04_27, f05_27, f06_27, f07_27, f08_27, f09_27, f10_27, f11_27, f12_27, f13_27, f14_27, f15_27, f16_27)
 Digits(27)=Array(f01_28, f02_28, f03_28, f04_28, f05_28, f06_28, f07_28, f08_28, f09_28, f10_28, f11_28, f12_28, f13_28, f14_28, f15_28, f16_28)
 Digits(28)=Array(f01_29, f02_29, f03_29, f04_29, f05_29, f06_29, f07_29, f08_29, f09_29, f10_29, f11_29, f12_29, f13_29, f14_29, f15_29, f16_29)
 Digits(29)=Array(f01_30, f02_30, f03_30, f04_30, f05_30, f06_30, f07_30, f08_30, f09_30, f10_30, f11_30, f12_30, f13_30, f14_30, f15_30, f16_30)
 Digits(30)=Array(f01_31, f02_31, f03_31, f04_31, f05_31, f06_31, f07_31, f08_31, f09_31, f10_31, f11_31, f12_31, f13_31, f14_31, f15_31, f16_31)
 Digits(31)=Array(f01_32, f02_32, f03_32, f04_32, f05_32, f06_32, f07_32, f08_32, f09_32, f10_32, f11_32, f12_32, f13_32, f14_32, f15_32, f16_32)
 Digits(32)=Array(f01_33, f02_33, f03_33, f04_33, f05_33, f06_33, f07_33, f08_33, f09_33, f10_33, f11_33, f12_33, f13_33, f14_33, f15_33, f16_33)
 Digits(33)=Array(f01_34, f02_34, f03_34, f04_34, f05_34, f06_34, f07_34, f08_34, f09_34, f10_34, f11_34, f12_34, f13_34, f14_34, f15_34, f16_34)
 Digits(34)=Array(f01_35, f02_35, f03_35, f04_35, f05_35, f06_35, f07_35, f08_35, f09_35, f10_35, f11_35, f12_35, f13_35, f14_35, f15_35, f16_35)
 Digits(35)=Array(f01_36, f02_36, f03_36, f04_36, f05_36, f06_36, f07_36, f08_36, f09_36, f10_36, f11_36, f12_36, f13_36, f14_36, f15_36, f16_36)
 Digits(36)=Array(f01_37, f02_37, f03_37, f04_37, f05_37, f06_37, f07_37, f08_37, f09_37, f10_37, f11_37, f12_37, f13_37, f14_37, f15_37, f16_37)
 Digits(37)=Array(f01_38, f02_38, f03_38, f04_38, f05_38, f06_38, f07_38, f08_38, f09_38, f10_38, f11_38, f12_38, f13_38, f14_38, f15_38, f16_38)
 Digits(38)=Array(f01_39, f02_39, f03_39, f04_39, f05_39, f06_39, f07_39, f08_39, f09_39, f10_39, f11_39, f12_39, f13_39, f14_39, f15_39, f16_39)
 Digits(39)=Array(f01_40, f02_40, f03_40, f04_40, f05_40, f06_40, f07_40, f08_40, f09_40, f10_40, f11_40, f12_40, f13_40, f14_40, f15_40, f16_40)



 Sub Displaytimer_Timer
    Dim ChgLED, ii, jj, num, chg, stat, obj, b, x
    ChgLED=Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
       For ii=0 To UBound(chgLED)
          num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)
          For Each obj In Digits(num)
             If chg And 1 Then obj.visible=stat And 1
'             If chg And 1 Then obj.State=stat And 1
             chg=chg\2 : stat=stat\2
          Next
       Next
    End If
 End Sub



'Gottlieb Genesis
'added by Inkochnito
Sub editDips
	Dim vpmDips:Set vpmDips = New cvpmDips
	With vpmDips
		.AddForm 700, 400, "Genesis - DIP switches"
		.AddFrame 2, 4, 190, "Maximum credits", 49152, Array("8 credits", 0, "10 credits", 32768, "15 credits", &H00004000, "20 credits", 49152)                                                                                  'dip 15&16
		.AddFrame 2, 80, 190, "Coin chute 1 and 2 control", &H00002000, Array("seperate", 0, "same", &H00002000)                                                                                                                  'dip 14
		.AddFrame 2, 126, 190, "Playfield special", &H00200000, Array("replay", 0, "extra ball", &H00200000)                                                                                                                      'dip 22
		.AddFrame 2, 172, 190, "High games to date control", &H00000020, Array("no effect", 0, "reset high games 2-5 on power off", &H00000020)                                                                                   'dip 6
		.AddFrame 2, 218, 190, "Completing drop target sequence", &H00000080, Array("adds a letter to most complete part", 0, "spots a letter to each part", &H00000080)                                                          'dip 8
		.AddFrame 2, 264, 190, "Special lights after", &H40000000, Array("Hitting 'Lifeforce' when flashing", 0, "completing all body parts", &H40000000)                                                                         'dip 31
		.AddFrame 2, 310, 190, "Extra ball after completing", &H80000000, Array("4 body parts during the same ball", 0, "3 body parts during the same ball", &H80000000)                                                          'dip 32
		.AddFrame 205, 4, 190, "High game to date awards", &H00C00000, Array("not displayed and no award", 0, "displayed and no award", &H00800000, "displayed and 2 replays", &H00400000, "displayed and 3 replays", &H00C00000) 'dip 23&24
		.AddFrame 205, 80, 190, "Balls per game", &H01000000, Array("5 balls", 0, "3 balls", &H01000000)                                                                                                                          'dip 25
		.AddFrame 205, 126, 190, "Replay limit", &H04000000, Array("no limit", 0, "one per game", &H04000000)                                                                                                                     'dip 27
		.AddFrame 205, 172, 190, "Novelty", &H08000000, Array("normal", 0, "extra ball and replay scores points", &H08000000)                                                                                                     'dip 28
		.AddFrame 205, 218, 190, "Game mode", &H10000000, Array("replay", 0, "extra ball", &H10000000)                                                                                                                            'dip 29
		.AddFrame 205, 264, 190, "3rd coin chute credits control", &H20000000, Array("no effect", 0, "add 9", &H20000000)                                                                                                         'dip 30
		.AddChk 205, 316, 180, Array("Match feature", &H02000000)                                                                                                                                                                 'dip 26
		.AddChk 205, 331, 190, Array("Attract sound", &H00000040)                                                                                                                                                                 'dip 7
		.AddLabel 50, 360, 300, 20, "After hitting OK, press F3 to reset game with new settings."
		.ViewDips
	End With
End Sub

Set vpmShowDips = GetRef("editDips")

' Rules
Sub Rules()
	Dim Msg(32)
	Msg(0) = "Genesis - Gottlieb 1986" &Chr(10) &Chr(10)
	Msg(1) = ""
	Msg(2) = "SPECIAL: Completing all Body Parts lights LIFEFORCE"
	Msg(3) = "  Hitting the Vari-Target all the way back lights SPECIAL"
	Msg(4) = ""
	Msg(5) = "EXTRA BALL: Completing 3 Body Parts lights EXTRA BALL"
	Msg(6) = "  Completing next Body Part awards EXTRA BALL"
	Msg(7) = ""
	Msg(8) = "SCORING MULTIPLIER: Completing Body Parts when needed"
	Msg(9) = "  advances Scoreing Multiplier"
	Msg(10) = ""
	Msg(11) = "MULTI-MULTIPLIER: Scoring Multiplier is doubled during Multi-Ball play"
	Msg(12) = ""
	Msg(13) = "BODY PARTS LETTERS: Letters awarded bt hitting Vari-target"
	Msg(14) = "  all the way back or by scoring Flashing Targets or Sequences."
	Msg(15) = "  Return Rollovers flash Vari-Target for a period of time."
	Msg(16) = "  Hitting Vari-Target all the way back when flashing"
	Msg(17) = "  awards a letter in all Body Parts."
	Msg(18) = "  Completing the Drop Target Sequence (1-2-3) awards"
	Msg(19) = "  a letter in the most complete Body Part."
	Msg(20) = ""
	Msg(21) = "LIFEFORCE: Expose Robot by hitting Vari-target all the way back"
	Msg(22) = "  when LIFEFORCE is flashing."
	Msg(23) = ""
	Msg(24) = "MULTIBALL: Completing Body Part when needed enables ramp for capture"
	Msg(25) = ""
	Msg(26) = "ENTERING INITIALS: Enter letter by presing Flippers and Credit Button."
	Msg(27) = ""
	Msg(28) = ""

	For X = 1 To 28
		Msg(0) = Msg(0) + Msg(X) &Chr(13)
	Next

	MsgBox Msg(0), , "         Instructions and Rule Card"
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

Sub RollingUpdate()
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

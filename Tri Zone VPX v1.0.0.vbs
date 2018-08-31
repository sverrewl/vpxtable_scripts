'*******************************************************************************************************
'
'									  Tri Zone Williams 1979 v1.0.0
'								 http://www.ipdb.org/machine.cgi?id=2641
'
'											Created by Kiwi
'
'*******************************************************************************************************

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' No special SSF tweaks yet.

Option Explicit
Randomize

'*******************************************************************************************************

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

'********************************************** OPTIONS ************************************************

Dim DmdHidden

 If Table1.ShowDT = True Then
	DmdHidden = 1
Else
	DmdHidden = 0		'Put 1 if you want to have DMD hidden in FS mode
End If

Const cGameName   = "trizn_l1"	'ROM

'********************************************** GiMode

' Whith GiMode = 0 GI lights are always on, whith GiMode = 1 GI lights are on when the game is in play

Const GiMode = 0

'********************************************** Rails Hidden/Visible in FS mode : 0 hidden , 1 visible

Const RailsVisible = 0

'********************************************** Ball

'Const BallSize    = 50

'Const BallMass    = 1			'Mass=(50^3)/125000 ,(BallSize^3)/125000

'******************************************** OPTIONS END **********************************************

LoadVPM "01120100", "s6.vbs", 3.37

Dim bsTrough, dtBank, bsHole, PinPlay

Const UseSolenoids = 1
Const UseLamps = 0
Const UseGI = 0
Const UseSync = 0
Const HandleMech = 0

' Standard Sounds
Const SSolenoidOn = "Solenoid"
Const SSolenoidOff = ""
Const SFlipperOn = ""
Const SFlipperOff = ""
Const SCoin = "coin3"

'************
' Table init.
'************

Sub Table1_Init
	With Controller
		.GameName = cGameName
		.SplashInfoLine = "Tri Zone Williams 1979" & vbNewLine & "VPX table by Kiwi 1.0.0"
		.HandleMechanics = 0
		.HandleKeyboard = 0
		.ShowDMDOnly = 1
		.ShowFrame = 0
		.ShowTitle = 0
		.Hidden = DmdHidden
'		.Games(cGameName).Settings.Value("dmd_pos_x")=0
'		.Games(cGameName).Settings.Value("dmd_pos_y")=0
'		.Games(cGameName).Settings.Value("dmd_width")=400
'		.Games(cGameName).Settings.Value("dmd_height")=92
'		.Games(cGameName).Settings.Value("rol") = 0
'		.Games(cGameName).Settings.Value("sound") = 1
'		.Games(cGameName).Settings.Value("ddraw") = 1
		If Err Then MsgBox Err.Description
	End With
	On Error Goto 0
	Controller.SolMask(0) = 0
	vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the timer to renable all the solenoids after 2 seconds
	Controller.Run

	' Nudging
	vpmNudge.TiltSwitch = 34
	vpmNudge.Sensitivity = 1
	vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingShot, RightSlingShot)

	' Trough
	Set bsTrough = New cvpmBallStack
	With bsTrough
		.InitSw 0, 9, 0, 0, 0, 0, 0, 0
		.InitKick BallRelease, 61, 7
		.InitEntrySnd "Solenoid", "Solenoid"
		.InitExitSnd SoundFX("popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
		.Balls = 1
	End With

	' Z-O-N-E Drop targets
	set dtBank = new cvpmdroptarget
	With dtBank
		.InitDrop Array(sw30, Array(sw16 ,sw16a), Array(sw33 ,sw33a), Array(sw20 ,sw20a)), Array(30, 16, 33, 20)
'		.Initsnd  SoundFX("DROPTARG",DOFContactors), SoundFX("DTResetB",DOFContactors)
		.AllDownSw = 35
	End With

	' Eject Hole
	Set bsHole = New cvpmBallStack
	With bsHole
		.InitSaucer sw25, 25, 170, 12
		.InitExitSnd SoundFX("popper",DOFContactors), SoundFX("popper",DOFContactors)
		.KickForceVar = 3
		.KickAngleVar = 1
		.KickZ=0.75
	End With

	' Main Timer init
	PinMAMETimer.Interval = PinMAMEInterval
	PinMAMETimer.Enabled = 1

' GI Delay Timer
'	If Controller.Lamp(61) Then
	GIDelay.Enabled = 1

	FGi1.IntensityScale = 0
	FGi2.IntensityScale = 0
	FGi3.IntensityScale = 0
	FGi4.IntensityScale = 0
	FGi5.IntensityScale = 0
	FGi6.IntensityScale = 0
	FGi7.IntensityScale = 0
	FGi8.IntensityScale = 0
	FGi10.IntensityScale = 0
	FGi11.IntensityScale = 0
	FGi12.IntensityScale = 0
	FGi13.IntensityScale = 0
	FGi14.IntensityScale = 0
	FGi15.IntensityScale = 0
	FGi20.IntensityScale = 0
	FGi21.IntensityScale = 0
	FGi22.IntensityScale = 0
	FGi23.IntensityScale = 0
	FGi24.IntensityScale = 0
	FGi25.IntensityScale = 0
	FGi26.IntensityScale = 0
	FGi27.IntensityScale = 0

	GiL8DF.Visible=0
	GiL11DF.Visible=0
	GiL13DF.Visible=0

	l50.SetValue 3
	l51.SetValue 3
	l52.SetValue 3
	l53.SetValue 3
	l54.SetValue 3
	l55.SetValue 3
	l57.SetValue 3
	l58.SetValue 3
	l59.SetValue 3
	l60.SetValue 3
	l61.SetValue 3
	l62.SetValue 3
	l63.SetValue 3
	l64.SetValue 3

	If ShowDT=True Then
 		l50.visible=1
 		l51.visible=1
 		l52.visible=1
 		l53.visible=1
 		l54.visible=1
		l55.visible=1
		l57.visible=1
		l58.visible=1
		l59.visible=1
		l60.visible=1
		l61.visible=1
		l62.visible=1
		l63.visible=1
		l64.visible=1
		da1.visible=1
		da2.visible=1
		da3.visible=1
		da4.visible=1
		da5.visible=1
		da6.visible=1
		db1.visible=1
		db2.visible=1
		db3.visible=1
		db4.visible=1
		db5.visible=1
		db6.visible=1
		dc1.visible=1
		dc2.visible=1
		dc3.visible=1
		dc4.visible=1
		dc5.visible=1
		dc6.visible=1
		dd1.visible=1
		dd2.visible=1
		dd3.visible=1
		dd4.visible=1
		dd5.visible=1
		dd6.visible=1
		de1.visible=1
		de2.visible=1
		de3.visible=1
		de4.visible=1

	Else
		RailSX.visible=RailsVisible
		RailDX.visible=RailsVisible
		l38a.TransmissionScale=0
		l39a.TransmissionScale=0
		l40a.TransmissionScale=0
	End If

End Sub

Sub Table1_Exit():Controller.Stop:End Sub
Sub Table1_Paused:Controller.Pause = 1:End Sub
Sub Table1_unPaused:Controller.Pause = 0:End Sub

' GI Init
Sub GIDelay_Timer()
 If GiMode = 0 Then
	SetLamp 150, 1
	SetLamp 151, 1
End If
	LampTimer.Enabled = 1
	GIDelay.Enabled = 0
End Sub

'**********
' Keys
'**********

Sub Table1_KeyDown(ByVal Keycode)
	If KeyCode = LeftFlipperKey And PinPlay=1 Then LeftFlipper.RotateToEnd
	If KeyCode = RightFlipperKey And PinPlay=1 Then RightFlipper.RotateToEnd
	If keycode = LeftTiltKey Then Nudge 90, 4:PlaySound SoundFX("nudge_left",0)
	If keycode = RightTiltKey Then Nudge 270, 4:PlaySound SoundFX("nudge_right",0)
	If keycode = CenterTiltKey Then Nudge 0, 5:PlaySound SoundFX("nudge_forward",0)
	If vpmKeyDown(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Pullback
End Sub

Sub Table1_KeyUp(ByVal Keycode)
	If KeyCode = LeftFlipperKey And PinPlay=1 Then LeftFlipper.RotateToStart
	If KeyCode = RightFlipperKey And PinPlay=1 Then RightFlipper.RotateToStart
	If vpmKeyUp(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Fire:PlaySound "plunger2"
End Sub

'**********
' Switches
'**********

'Slings

Dim LStep, RStep

Sub LeftSlingShot_Slingshot:vpmTimer.PulseSw 26:LeftSling.Visible=1:EmKickerT1L.TransX = -22:LStep=0:Me.TimerEnabled=1:PlaySound SoundFX("slingshot_left",DOFContactors),0,1,0,0:End Sub
Sub LeftSlingShot_Timer
	Select Case LStep
		Case 0:LeftSLing.Visible = 1:EmKickerT1L.TransX = -22
		Case 1: 'pause
		Case 2:LeftSLing.Visible = 0:LeftSLing1.Visible = 1:EmKickerT1L.TransX = -17
		Case 3:LeftSLing1.Visible = 0:LeftSLing2.Visible = 1:EmKickerT1L.TransX = -11
		Case 4:LeftSLing2.Visible = 0:Me.TimerEnabled = 0:EmKickerT1L.TransX = 0
	End Select
	LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot:vpmTimer.PulseSw 24:RightSling.Visible=1:EmKickerT1R.TransX = -22:RStep=0:Me.TimerEnabled = 1:PlaySound SoundFX("slingshot_right",DOFContactors),0,1,0,0:End Sub
Sub RightSlingShot_Timer
	Select Case RStep
		Case 0:RightSLing.Visible = 1:EmKickerT1R.TransX = -22
		Case 1: 'pause
		Case 2:RightSLing.Visible = 0:RightSLing1.Visible = 1:EmKickerT1R.TransX = -17
		Case 3:RightSLing1.Visible = 0:RightSLing2.Visible = 1:EmKickerT1R.TransX = -11
		Case 4:RightSLing2.Visible = 0:Me.TimerEnabled = 0:EmKickerT1R.TransX = 0
	End Select
	RStep = RStep + 1
End Sub

'Rubbers
Sub sw12_Hit:vpmTimer.PulseSw 12:PlaySound "rubber1",0,1,pan(ActiveBall),0:End Sub
Sub sw17_Hit:vpmTimer.PulseSw 17:PlaySound "rubber1",0,1,pan(ActiveBall),0:End Sub
Sub sw21_Hit:vpmTimer.PulseSw 21:PlaySound "rubber1",0,1,pan(ActiveBall),0:End Sub
Sub sw29_Hit:vpmTimer.PulseSw 29:PlaySound "rubber1",0,1,pan(ActiveBall),0:End Sub

'Bumpers
Sub Bumper1_Hit
	If PinPlay=1 Then:vpmTimer.PulseSw 13:Ring1.Z = -28:Me.TimerEnabled = 1:PlaySound SoundFX("Bumper6",DOFContactors),0,1,-0.02,0:End If:End Sub
Sub Bumper1_Timer()
	Ring1.Z = Ring1.Z +2
 If Ring1.Z = 0 Then:Me.TimerEnabled = 0
End Sub

Sub Bumper2_Hit
	If PinPlay=1 Then:vpmTimer.PulseSw 14:Ring2.Z = -28:Me.TimerEnabled = 1:PlaySound SoundFX("Bumper6",DOFContactors),0,1,-0.05,0:End If:End Sub
Sub Bumper2_Timer()
	Ring2.Z = Ring2.Z +2
 If Ring2.Z = 0 Then:Me.TimerEnabled = 0
End Sub

Sub Bumper3_Hit
	If PinPlay=1 Then:vpmTimer.PulseSw 15:Ring3.Z = -28:Me.TimerEnabled = 1:PlaySound SoundFX("Bumper6",DOFContactors),0,1,0,0:End If:End Sub
Sub Bumper3_Timer()
	Ring3.Z = Ring3.Z +2
 If Ring3.Z = 0 Then:Me.TimerEnabled = 0
End Sub

'Spinner
Sub sw31_Spin:vpmTimer.PulseSw 31:PlaySound "spinner":End Sub

'Rollovers
Sub sw10_Hit:  Controller.Switch(10) = 1:Switch10.RotX=15:PlaySound "sensor",0,1,pan(ActiveBall),0:End Sub
Sub sw10_UnHit:Controller.Switch(10) = 0:Switch10.RotX=0:End Sub
Sub sw11_Hit:  Controller.Switch(11) = 1:Switch11.RotX=15:PlaySound "sensor",0,1,pan(ActiveBall),0:End Sub
Sub sw11_UnHit:Controller.Switch(11) = 0:Switch11.RotX=0:End Sub
Sub sw18_Hit:  Controller.Switch(18) = 1:PlaySound "sensor",0,1,pan(ActiveBall),0:Psw18.Z=-3:End Sub
Sub sw18_UnHit:Controller.Switch(18) = 0:Psw18.Z=3:End Sub
Sub sw19_Hit:  Controller.Switch(19) = 1:PlaySound "sensor",0,1,pan(ActiveBall),0:Psw19.Z=-3:End Sub
Sub sw19_UnHit:Controller.Switch(19) = 0:Psw19.Z=3:End Sub
Sub sw22_Hit:  Controller.Switch(22) = 1:Switch22.RotX=15:PlaySound "sensor",0,1,pan(ActiveBall),0:End Sub
Sub sw22_UnHit:Controller.Switch(22) = 0:Switch22.RotX=0:End Sub
Sub sw23_Hit:  Controller.Switch(23) = 1:Switch23.RotX=15:PlaySound "sensor",0,1,pan(ActiveBall),0:End Sub
Sub sw23_UnHit:Controller.Switch(23) = 0:Switch23.RotX=0:End Sub
Sub sw27_Hit:  Controller.Switch(27) = 1:Switch27.RotX=15:PlaySound "sensor",0,1,pan(ActiveBall),0:End Sub
Sub sw27_UnHit:Controller.Switch(27) = 0:Switch27.RotX=0:End Sub
Sub sw28_Hit:  Controller.Switch(28) = 1:Switch28.RotX=15:PlaySound "sensor",0,1,pan(ActiveBall),0:End Sub
Sub sw28_UnHit:Controller.Switch(28) = 0:Switch28.RotX=0:End Sub
Sub sw32_Hit:  Controller.Switch(32) = 1:Switch32.RotX=15:PlaySound "sensor",0,1,pan(ActiveBall),0:End Sub
Sub sw32_UnHit:Controller.Switch(32) = 0:Switch32.RotX=0:End Sub

'Targets
Sub sw30_Hit:dtBank.hit 1:PlaySound SoundFX("DROPTARG",DOFContactors),0,1,-0.07,0:Sw30Shadow:End Sub
Sub sw16_Hit:dtBank.hit 2:PlaySound SoundFX("DROPTARG",DOFContactors),0,1,-0.01,0:Sw16Shadow:End Sub
Sub sw33_Hit:dtBank.hit 3:PlaySound SoundFX("DROPTARG",DOFContactors),0,1,-0.04,0:End Sub
Sub sw20_Hit:dtBank.hit 4:PlaySound SoundFX("DROPTARG",DOFContactors),0,1,0.07,0:Sw20Shadow:End Sub

Sub Sw30Shadow
 If sw30.IsDropped=0 Then
	GiL11DF.Visible=1
End If
End Sub

Sub Sw16Shadow
 If sw16.IsDropped=0 Then
	GiL13DF.Visible=1
End If
End Sub

Sub Sw20Shadow
 If sw20.IsDropped=0 Then
	GiL8DF.Visible=1
End If
End Sub

'Kickers
Sub Drain_Hit:bsTrough.AddBall Me:PlaySound "drain1a",0,1,0,0:End Sub
Sub sw25_Hit:bsHole.AddBall Me:End Sub

'Gates
Sub Gate3_Hit:PlaySound "Gate5",0,1,0.1,0:End Sub
Sub GateTSx1_Hit:PlaySound "Gate5",0,1,-0.2,0:End Sub
Sub GateTDx1_Hit:PlaySound "Gate5",0,1,0.2,0:End Sub

Dim PI
PI = Round(4 * Atn(1), 6) '3.1415926535897932384626433832795

Sub UpdateGates
	GateTSx.RotY = -GateTSx1.CurrentAngle*0.6
	GateTDx.RotY = -GateTDx1.CurrentAngle*0.6
	Gate3D.RotX = -Gate3.CurrentAngle/1.5-30
	LogoL.RotZ = LeftFlipper.CurrentAngle
	LogoR.RotZ = RightFlipper.CurrentAngle
	pSpinnerRod.TransX = sin( (sw31.CurrentAngle+180) * (2*PI/360)) * 9
	pSpinnerRod.TransZ = sin( (sw31.CurrentAngle- 90) * (2*PI/360)) * 9
	pSpinnerRod.RotY = (sin( (sw31.CurrentAngle-180) * (2*PI/360)) * 6) + 13
End Sub

'*********
'Solenoids
'*********
SolCallback(1) = "bsTrough.SolOut"
SolCallback(2) = "bsHoleEject"
SolCallback(3) = "SolReset 1,"
SolCallback(4) = "SolReset 2,"
SolCallback(5) = "SolReset 3,"
SolCallback(6) = "SolReset 4,"
solcallback(14) = "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
solcallback(23) = "SolRun"

Sub bsHoleEject(Enabled)
 If Enabled Then
	bsHole.ExitSol_On
	Arm.TransZ=15
	ArmTimer.Enabled=1
End If
End Sub

Sub ArmTimer_Timer
	Arm.TransZ=0
	ArmTimer.Enabled=0
End Sub

Sub SolReset(No ,Enabled)
	If Enabled Then
		Controller.Switch(35) = False
			dtBank.SolUnHit 1, True
			dtBank.SolUnHit 2, True
			dtBank.SolUnHit 3, True
			dtBank.SolUnHit 4, True
			ShadowsTimer.Enabled=1
	PlaySound SoundFX("DTResetB",DOFContactors),0,1,0,0
	End If
End Sub

Sub ShadowsTimer_Timer
	GiL8DF.Visible=0
	GiL11DF.Visible=0
	GiL13DF.Visible=0
	ShadowsTimer.Enabled=0
End Sub

Sub SolRun(Enabled)
	vpmNudge.SolGameOn Enabled
 If Enabled Then
	PinPlay=1
	LeftSlingShot.Disabled=0
	RightSlingShot.Disabled=0
Else
	PinPlay=0
	LeftFlipper.RotateToStart
	RightFlipper.RotateToStart
	LeftSlingShot.Disabled=1
	RightSlingShot.Disabled=1
End If
 If Enabled And GiMode = 1 Then
	SetLamp 150, 1
	SetLamp 151, 1
End If
 If Enabled+GiMode = 1 Then
	SetLamp 150, 0
	SetLamp 151, 0
End If
End Sub

'**************
' Flipper Subs
'**************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
	If Enabled Then
		PlaySound SoundFX("flipperup_left",DOFContactors),0,1,0,0':LeftFlipper.RotateToEnd
	Else
		PlaySound SoundFX("flipperdown_left",DOFContactors),0,1,0,0':LeftFlipper.RotateToStart
	End If
End Sub

Sub SolRFlipper(Enabled)
	If Enabled Then
		PlaySound SoundFX("flipperup_right",DOFContactors),0,1,0,0':RightFlipper.RotateToEnd
	Else
		PlaySound SoundFX("flipperdown_right",DOFContactors),0,1,0,0':RightFlipper.RotateToStart
	End If
End Sub

Sub LeftFlipper_Collide(parm)
	PlaySound "rubber_flipper", 0, Vol(ActiveBall), pan(ActiveBall)
End Sub

Sub RightFlipper_Collide(parm)
	PlaySound "rubber_flipper", 0, Vol(ActiveBall), pan(ActiveBall)
End Sub

'******************
' RealTime Updates
'******************
Set MotorCallback = GetRef("GameTimer")

Sub GameTimer
	UpdateGates
End Sub

' *********************************************************************
' 					Wall, rubber and metal hit sounds
' *********************************************************************
'Sounds FX

Sub Rubbers_Hit(idx):PlaySound "rubber1", 0, Vol(ActiveBall), pan(ActiveBall), 0.25, AudioFade(ActiveBall):End Sub


'************************************
'          LEDs Display
'************************************

Dim Digits(28)

Set Digits(0) = da1
Set Digits(1) = da2
Set Digits(2) = da3
Set Digits(3) = da4
Set Digits(4) = da5
Set Digits(5) = da6

Set Digits(6) = db1
Set Digits(7) = db2
Set Digits(8) = db3
Set Digits(9) = db4
Set Digits(10) = db5
Set Digits(11) = db6

Set Digits(12) = dc1
Set Digits(13) = dc2
Set Digits(14) = dc3
Set Digits(15) = dc4
Set Digits(16) = dc5
Set Digits(17) = dc6

Set Digits(18) = dd1
Set Digits(19) = dd2
Set Digits(20) = dd3
Set Digits(21) = dd4
Set Digits(22) = dd5
Set Digits(23) = dd6

Set Digits(24) = de1
Set Digits(25) = de2
Set Digits(26) = de3
Set Digits(27) = de4

Sub UPdateLEDs
	On Error Resume Next
	Dim ChgLED, ii, jj, chg, stat
	ChgLED = Controller.ChangedLEDs(&H0000003f, &Hffffffff)
	If Not IsEmpty(ChgLED) Then
		For ii = 0 To UBound(ChgLED)
			chg = chgLED(ii, 1):stat = chgLED(ii, 2)

			Select Case stat
				Case 0:Digits(chgLED(ii, 0) ).SetValue 0    'empty
				Case 63:Digits(chgLED(ii, 0) ).SetValue 1   '0
				Case 6:Digits(chgLED(ii, 0) ).SetValue 2    '1
				Case 91:Digits(chgLED(ii, 0) ).SetValue 3   '2
				Case 79:Digits(chgLED(ii, 0) ).SetValue 4   '3
				Case 102:Digits(chgLED(ii, 0) ).SetValue 5  '4
				Case 109:Digits(chgLED(ii, 0) ).SetValue 6  '5
				Case 124:Digits(chgLED(ii, 0) ).SetValue 7  '6
				Case 125:Digits(chgLED(ii, 0) ).SetValue 7  '6
				Case 252:Digits(chgLED(ii, 0) ).SetValue 7  '6
				Case 7:Digits(chgLED(ii, 0) ).SetValue 8    '7
				Case 127:Digits(chgLED(ii, 0) ).SetValue 9  '8
				Case 103:Digits(chgLED(ii, 0) ).SetValue 10 '9
				Case 111:Digits(chgLED(ii, 0) ).SetValue 10 '9
				Case 231:Digits(chgLED(ii, 0) ).SetValue 10 '9
				Case 128:Digits(chgLED(ii, 0) ).SetValue 0  'empty
				Case 191:Digits(chgLED(ii, 0) ).SetValue 1  '0
				Case 832:Digits(chgLED(ii, 0) ).SetValue 2  '1
				Case 896:Digits(chgLED(ii, 0) ).SetValue 2  '1
				Case 768:Digits(chgLED(ii, 0) ).SetValue 2  '1
				Case 134:Digits(chgLED(ii, 0) ).SetValue 2  '1
				Case 219:Digits(chgLED(ii, 0) ).SetValue 3  '2
				Case 207:Digits(chgLED(ii, 0) ).SetValue 4  '3
				Case 230:Digits(chgLED(ii, 0) ).SetValue 5  '4
				Case 237:Digits(chgLED(ii, 0) ).SetValue 6  '5
				Case 253:Digits(chgLED(ii, 0) ).SetValue 7  '6
				Case 135:Digits(chgLED(ii, 0) ).SetValue 8  '7
				Case 255:Digits(chgLED(ii, 0) ).SetValue 9  '8
				Case 239:Digits(chgLED(ii, 0) ).SetValue 10 '9
			End Select
		Next
	End IF
End Sub

'**********************************
'       JP's VP10 Fading Lamps & Flashers
'       Based on PD's Fading Light System
' SetLamp 0 is Off
' SetLamp 1 is On
' fading for non opacity objects is 4 steps
'***************************************************

Dim LampState(200), FadingLevel(200)
Dim FlashSpeedUp(200), FlashSpeedDown(200), FlashMin(200), FlashMax(200), FlashLevel(200)

InitLamps()             ' turn off the lights and flashers and reset them to the default parameters
LampTimer.Interval = 10 'lamp fading speed
'LampTimer.Enabled = 1

' Lamp & Flasher Timers

Sub LampTimer_Timer()
    Dim chgLamp, num, chg, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp) Then
        For ii = 0 To UBound(chgLamp)
            LampState(chgLamp(ii, 0) ) = chgLamp(ii, 1)       'keep the real state in an array
            FadingLevel(chgLamp(ii, 0) ) = chgLamp(ii, 1) + 4 'actual fading step
        Next
    End If

	UpdateLeds
    UpdateLamps
End Sub

Sub UpdateLamps
	NFadeL 2, l2
	NFadeL 3, l3
	NFadeL 4, l4
	NFadeL 5, l5
	NFadeL 6, l6
	NFadeL 7, l7
	NFadeL 8, l8
	NFadeL 9, l9
	NFadeL 10, l10
	NFadeL 12, l12
	NFadeL 13, l13
	NFadeL 14, l14
	NFadeL 15, l15
	NFadeL 16, l16
	NFadeL 17, l17
	NFadeL 18, l18
	NFadeL 19, l19
	NFadeL 20, l20
	NFadeL 21, l21
	NFadeL 22, l22
	NFadeL 23, l23
	NFadeL 24, l24
	NFadeL 25, l25
	NFadeL 26, l26
	NFadeL 27, l27
	NFadeL 28, l28
	NFadeL 29, l29
	NFadeL 30, l30
	NFadeL 31, l31
	NFadeL 32, l32
	NFadeL 33, l33
	NFadeL 34, l34
	NFadeL 35, l35
	NFadeL 36, l36
	NFadeL 41, l41
	NFadeL 42, l42
	NFadeL 43, l43
	NFadeL 44, l44
	NFadeL 46, l46

	NFadeLm 38, l38
	NFadeL 38, l38a

	NFadeLm 39, l39
	NFadeL 39, l39a

	NFadeLm 40, l40
	NFadeL 40, l40a

	NFadeL 56, l56

   'Backdrop lights

	FadeR 50, l50
	FadeR 51, l51
	FadeR 52, l52
	FadeR 53, l53
	FadeR 57, l57
	FadeR 58, l58
	FadeR 59, l59
	FadeR 60, l60

	FadeR 54, l54
	FadeR 55, l55
	FadeR 61, l61
	FadeR 62, l62
	FadeR 63, l63
	FadeR 64, l64

	NFadeLm 150, GiL1
	NFadeLm 150, GiL2
	NFadeLm 150, GiL3U
	NFadeLm 150, GiL3UL
	NFadeLm 150, GiL3L
	NFadeLm 150, GiL4
	NFadeLm 150, GiL5
	NFadeLm 150, GiL6U
	NFadeLm 150, GiL6UL
	NFadeLm 150, GiL6L
	NFadeLm 150, GiL7U
	NFadeLm 150, GiL7UL
	NFadeLm 150, GiL7D
	NFadeLm 150, GiL8U
	NFadeLm 150, GiL8UL
	NFadeLm 150, GiL8D
	NFadeLm 150, GiL8DF
	NFadeLm 150, GiL9U
	NFadeLm 150, GiL9UL
	NFadeLm 150, GiL9D
	NFadeLm 150, GiL10U
	NFadeLm 150, GiL10UL
	NFadeLm 150, GiL10D
	NFadeLm 150, GiL11U
	NFadeLm 150, GiL11UL
	NFadeLm 150, GiL11D
	NFadeLm 150, GiL11DF
	NFadeLm 150, GiL12U
	NFadeLm 150, GiL12UL
	NFadeLm 150, GiL12D
	NFadeLm 150, GiL13U
	NFadeLm 150, GiL13D
	NFadeLm 150, GiL13T
	NFadeLm 150, GiL13DF
	NFadeLm 150, GiL14U
	NFadeLm 150, GiL14UL
	NFadeLm 150, GiL14D
	NFadeLm 150, GiL15U
	NFadeLm 150, GiL15UL
	NFadeLm 150, GiL15D
	NFadeLm 150, GiL16U
	NFadeLm 150, GiL16UL
	NFadeLm 150, GiL16D
	NFadeLm 150, GiL17U
	NFadeLm 150, GiL17D
	NFadeLm 150, GiL17UL
	NFadeLm 150, GiL18U
	NFadeLm 150, GiL18D
	NFadeLm 150, GiL18UL
	NFadeLm 150, GiL19U
	NFadeLm 150, GiL19D
	NFadeLm 150, GiL19UL
	NFadeLm 150, GiL20U
	NFadeLm 150, GiL20D
	NFadeLm 150, GiL20UL
	NFadeLm 150, GiL21U
	NFadeLm 150, GiL21D
	NFadeLm 150, GiL21UL
	NFadeLm 150, GiL22U
	NFadeLm 150, GiL22D
	NFadeLm 150, GiL22UL
	NFadeLm 150, GiL23U
	NFadeLm 150, GiL23UL
	NFadeLm 150, GiL23D
	NFadeLm 150, GiL24U
	NFadeLm 150, GiL24UL
	NFadeLm 150, GiL24D
	NFadeLm 150, GiL25D
	NFadeLm 150, GiL25U
	NFadeLm 150, GiL26D
	NFadeLm 150, GiL26U
	NFadeLm 150, GiL27D
	NFadeLm 150, GiL27U
	NFadeLm 150, GiL28U
	NFadeLm 150, GiL28UL
	NFadeLm 150, GiL29U
	NFadeLm 150, GiL29UL
	NFadeLm 150, GiL30U
	NFadeLm 150, GiL30UL
	NFadeLm 150, GiL31U
	NFadeLm 150, GiL31UL
	NFadeLm 150, GiL32U
	NFadeLm 150, GiL32UL
	NFadeLm 150, GiR1
	NFadeL 150, GiR2

	Flashm 151, FGi1
	Flashm 151, FGi2
	Flashm 151, FGi3
	Flashm 151, FGi4
	Flashm 151, FGi5
	Flashm 151, FGi6
	Flashm 151, FGi7
	Flashm 151, FGi8
	Flashm 151, FGi10
	Flashm 151, FGi11
	Flashm 151, FGi12
	Flashm 151, FGi13
	Flashm 151, FGi14
	Flashm 151, FGi15
	Flashm 151, FGi20
	Flashm 151, FGi21
	Flashm 151, FGi22
	Flashm 151, FGi23
	Flashm 151, FGi24
	Flashm 151, FGi25
	Flashm 151, FGi26
	Flash 151, FGi27

End Sub

' div lamp subs

Sub InitLamps()
    Dim x
    For x = 0 to 200
        LampState(x) = 0        ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4      ' used to track the fading state
        FlashSpeedUp(x) = 0.1   ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.1' slower speed when turning off the flasher
        FlashMax(x) = 1         ' the maximum value when on, usually 1
        FlashMin(x) = 0         ' the minimum value when off, usually 0
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

' Lights: old method, using 4 images

Sub FadeL(nr, light, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 4:light.image = b:FadingLevel(nr) = 6                   'fading to off...
        Case 5:light.image = a:light.State = 1:FadingLevel(nr) = 1   'ON
        Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1           'wait
        Case 9:light.image = c:FadingLevel(nr) = FadingLevel(nr) + 1 'fading...
        Case 10, 11, 12:FadingLevel(nr) = FadingLevel(nr) + 1        'wait
        Case 13:light.image = d:Light.State = 0:FadingLevel(nr) = 0  'Off
    End Select
End Sub

Sub FadeLm(nr, light, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 4:light.image = b
        Case 5:light.image = a
        Case 9:light.image = c
        Case 13:light.image = d
    End Select
End Sub

' Lights: used for VP10 standard lights, the fading is handled by VP itself

Sub NFadeL(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.state = 0:FadingLevel(nr) = 0
        Case 5:object.state = 1:FadingLevel(nr) = 1
    End Select
End Sub

Sub NFadeLm(nr, object) ' used for multiple lights
    Select Case FadingLevel(nr)
        Case 4:object.state = 0
        Case 5:object.state = 1
    End Select
End Sub

' Ramps & Primitives used as 4 step fading lights
'a,b,c,d are the images used from on to off

Sub FadeObj(nr, object, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 4:object.image = b:FadingLevel(nr) = 6                   'fading to off...
        Case 5:object.image = a:FadingLevel(nr) = 1                   'ON
        Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1            'wait
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
End Sub

Sub Flashm(nr, object) 'multiple flashers, it just sets the flashlevel
    Object.IntensityScale = FlashLevel(nr)
End Sub

' Desktop Objects: Reels & texts (you may also use lights on the desktop)

' Reels

Sub FadeR(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.SetValue 1:FadingLevel(nr) = 6                   'fading to off...
        Case 5:object.SetValue 0:FadingLevel(nr) = 1                   'ON
        Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1             'wait
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

Const tnob = 2 ' total number of balls
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


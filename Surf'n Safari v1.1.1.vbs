'*******************************************************************************************************
'
'					   	    	 Surf'n Safari Premier 1991 VPX v1.0.0
'								http://www.ipdb.org/machine.cgi?id=2461
'
'											Created by Kiwi
'
'*******************************************************************************************************

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' No special SSF tweaks yet.

Option Explicit
Randomize

'*******************************************************************************************************

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

'********************************************** OPTIONS ************************************************

Const cGameName   = "surfnsaf"

Const BallSize    = 50

Const Lumen    = 10

Const BallMass    = 1.025		'Mass=(53.11^3)/125000 ,(BallSize^3)/125000

'************ DMD Visible/Hidden : 0 visible , 1 hidden

Const DMDHidden = 1

'************ Rails and rail lights Hidden/Visible in FS mode : 0 hidden , 1 visible

Const RailsLights = 0

'************ Glasses Color : 1=White ,2=Yellow ,3=Orange ,4=Red ,5=Violet ,6=Green ,7=Blue

Const GlassesColor = 7

'******************************************** OPTIONS END **********************************************


LoadVPM "01560000", "gts3.VBS", 3.26

Dim bsTrough, bump1, bump2, bump3, cdtBank, ddtBank, bsUK, bsLK, mHole, mHole1, mHole2, PinPlay

Const UseSolenoids = 2
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


Const swLCoin = 0
Const swRCoin = 1
Const swCCoin = 2
Const swCoinShuttle = 3
Const swStartButton = 4
Const swTournament = 5
Const swFrontDoor = 6

'************
' Table init.
'************

Sub Table1_Init
	vpmInit me
	With Controller
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
		.SplashInfoLine = "Surf'n Safari Premier (Gottlieb 1991)" & vbNewLine & "VPX table by Kiwi 1.0.0"
		.HandleMechanics = 0
		.HandleKeyboard = 0
		.ShowDMDOnly = 1
		.ShowFrame = 0
		.ShowTitle = 0
		.Hidden = DMDHidden
'		.DoubleSize = 0
'		.Games(cGameName).Settings.Value("dmd_pos_x")=0
'		.Games(cGameName).Settings.Value("dmd_pos_y")=0
'		.Games(cGameName).Settings.Value("dmd_width")=400
'		.Games(cGameName).Settings.Value("dmd_height")=92
'		.Games(cGameName).Settings.Value("rol") = 0
'		.Games(cGameName).Settings.Value("sound") = 1
		.Games(cGameName).Settings.Value("dmd_red")=60
		.Games(cGameName).Settings.Value("dmd_green")=180
		.Games(cGameName).Settings.Value("dmd_blue")=255
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With

	' Nudging
	vpmNudge.TiltSwitch = 151
	vpmNudge.Sensitivity = 2
	vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingShot, RightSlingShot, sw15)

	' Trough
	Set bsTrough = New cvpmBallStack
	With bsTrough
		.InitSw 5, 0, 0, 25, 0, 0, 0, 0
		.InitKick BallRelease, 90, 4
		.InitEntrySnd "Solenoid", "Solenoid"
		.InitExitSnd SoundFX("popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
		.Balls = 3
	End With

	' Top Upkicker
	Set bsUK = New cvpmBallStack
	With bsUK
		.InitSaucer sw21, 21, 0, 32
		.KickZ = 1.56
		.InitExitSnd SoundFX("popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
	End With

	' Left Upkicker
	Set bsLK = New cvpmBallStack
	With bsLK
		.InitSaucer sw35, 35, 0, 35
		.KickZ = 1.56
		.InitExitSnd SoundFX("popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
	End With

	' Drop targets
	set cdtBank = new cvpmdroptarget
	With cdtBank
		.initdrop array(sw16, sw26, sw36), array(16, 26, 36)
'		.initsnd SoundFX("DROPTARG",DOFContactors), SoundFX("DTResetB",DOFContactors)
	End With

	set ddtBank = new cvpmdroptarget
	With ddtBank
		.initdrop array(sw17, sw27, sw37), array(17, 27, 37)
'		.initsnd SoundFX("DROPTARG",DOFContactors), SoundFX("DTResetB",DOFContactors)
	End With

	' Low powered Magnet , Left VUK
    Set mHole = New cvpmMagnet
    With mHole
        .initMagnet MagTrigger, 3
        .X = 77
        .Y = 1004.75
        .Size = 37
        .GrabCenter = 0
        .MagnetOn = 1
        .CreateEvents "mHole"
    End With

	' Low powered Magnet1 , Skill Shot
    Set mHole1 = New cvpmMagnet
    With mHole1
        .initMagnet MagTrigger1, 4
        .X = 475
        .Y = 188
        .Size = 35
        .GrabCenter = 0
        .MagnetOn = 1
        .CreateEvents "mHole1"
    End With

	' Low powered Magnet2 , Whirlpool
    Set mHole2 = New cvpmMagnet
    With mHole2
        .initMagnet MagTrigger2, 1
        .X = 462
        .Y = 405
        .Size = 95
        .GrabCenter = 0
        .MagnetOn = 1
        .CreateEvents "mHole2"
    End With

' Main Timer init
	PinMAMETimer.Interval = PinMAMEInterval
	PinMAMETimer.Enabled = 1

' GI Delay Timer
	GIDelay.Enabled = 1

' Init Droptargets
	Arm.IsDropped=1
	Arm2.IsDropped=1

	If ShowDT=True Then
		RailSx.visible=1
		RailDx.visible=1

		fgit1L.visible=1
		fgit1R.visible=1
		f43a.visible=1
		f50a.visible=1
		f56a.visible=1
		f57a.visible=1
		Bulb10b.visible=1
	Else
		RailSx.visible=RailsLights
		RailDx.visible=RailsLights

		fgit1L.visible=RailsLights
		fgit1R.visible=RailsLights
		f43a.visible=RailsLights
		f50a.visible=RailsLights
		f56a.visible=RailsLights
		f57a.visible=RailsLights
		Bulb10b.visible=RailsLights
	End If
End Sub

Sub Table1_Exit():Controller.Stop:End Sub
Sub Table1_Paused:Controller.Pause = 1:End Sub
Sub Table1_unPaused:Controller.Pause = 0:End Sub

' GI Init
Sub GIDelay_Timer()
	SetLamp 160, 1
	GIDelay.Enabled = 0
End Sub

' Glasses Color
 If GlassesColor=1 Then:Glasses.Image="GlassWhite":End If
 If GlassesColor=2 Then:Glasses.Image="GlassYellow":End If
 If GlassesColor=3 Then:Glasses.Image="GlassOrange":End If
 If GlassesColor=4 Then:Glasses.Image="GlassRed":End If
 If GlassesColor=5 Then:Glasses.Image="GlassViolet":End If
 If GlassesColor=6 Then:Glasses.Image="GlassGreen":End If
 If GlassesColor=7 Then:Glasses.Image="GlassBlue":End If

'**********
' Keys
'**********

Dim BumpersOff
Sub Table1_KeyDown(ByVal KeyCode)
	If KeyCode = LeftFlipperKey And PinPlay=1 Then LeftFlipper.RotateToEnd
	If KeyCode = RightFlipperKey And PinPlay=1 Then RightFlipper.RotateToEnd
	If keycode = PlungerKey Then Plunger.Pullback
	If keycode = LeftTiltKey Then Nudge 90, 4:PlaySound SoundFX("nudge_left",0)
	If keycode = RightTiltKey Then Nudge 270, 4:PlaySound SoundFX("nudge_right",0)
	If keycode = CenterTiltKey Then Nudge 0, 5:PlaySound SoundFX("nudge_forward",0)
	If vpmKeyDown(keycode) Then Exit Sub
    'debug key
    if keycode = "3" then
        setlamp 171, 1
        setlamp 172, 1
        setlamp 175, 1
        setlamp 190, 1
        setlamp 197, 1
        setlamp 198, 1
        setlamp 199, 1
        f151.State = 1
        f151a.State = 1
    end if
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
	If KeyCode = LeftFlipperKey And PinPlay=1 Then LeftFlipper.RotateToStart
	If KeyCode = RightFlipperKey And PinPlay=1 Then RightFlipper.RotateToStart
	If keycode = PlungerKey Then Plunger.Fire:PlaySound "plunger2"
	If vpmKeyUp(keycode) Then Exit Sub
    'debug key
    if keycode = "3" then
        setlamp 171, 0
        setlamp 172, 0
        setlamp 175, 0
        setlamp 190, 0
        setlamp 197, 0
        setlamp 198, 0
        setlamp 199, 0
        f151.State = 0
        f151a.State = 0
    end if
End Sub

'*********
' Switches
'*********

' Slings

Dim LStep, RStep

Sub LeftSlingshot_Slingshot:If PinPlay=1 Then:vpmTimer.PulseSw 13:LeftSling.Visible=1:SxEmKickerT1.TransX=-28:LStep=0:Me.TimerEnabled=1:PlaySound SoundFX("Slingshot",DOFContactors):End If:End Sub
Sub LeftSlingshot_Timer
	Select Case LStep
		Case 0:LeftSling.Visible = 1
		Case 1: 'pause
		Case 2:LeftSling.Visible = 0 :LeftSling1.Visible = 1:SxEmKickerT1.TransX=-23
		Case 3:LeftSling1.Visible = 0:LeftSling2.Visible = 1:SxEmKickerT1.TransX=-18.5
		Case 4:LeftSling2.Visible = 0:Me.TimerEnabled = 0:SxEmKickerT1.TransX=0
	End Select
	LStep = LStep + 1
End Sub

Sub RightSlingshot_Slingshot:If PinPlay=1 Then:vpmTimer.PulseSw 14:RightSling.Visible=1:DxEmKickerT1.TransX=-28:RStep=0:Me.TimerEnabled=1:PlaySound SoundFX("Slingshot",DOFContactors):End If:End Sub
Sub RightSlingshot_Timer
	Select Case RStep
		Case 0:RightSling.Visible = 1
		Case 1: 'pause
		Case 2:RightSling.Visible = 0 :RightSling1.Visible = 1:DxEmKickerT1.TransX=-23
		Case 3:RightSling1.Visible = 0:RightSling2.Visible = 1:DxEmKickerT1.TransX=-18.5
		Case 4:RightSling2.Visible = 0:Me.TimerEnabled = 0:DxEmKickerT1.TransX=0
	End Select
	RStep = RStep + 1
End Sub

' Bumpers
Sub Bumper1_Hit:If PinPlay=1 Then:bump1=1:Me.TimerEnabled=1:vpmTimer.PulseSw 10:PlaySound SoundFX("jet1",DOFContactors),0,1,-0.1,0:End If:End Sub
Sub Bumper1_Timer()
    Select Case bump1
        Case 1:Ring1.Z = 20:bump1 = 2
        Case 2:Ring1.Z = 30:bump1 = 3
        Case 3:Ring1.Z = 40:bump1 = 4
        Case 4:Ring1.Z = 50:Me.TimerEnabled = 0
    End Select
End Sub

Sub Bumper2_Hit:If PinPlay=1 Then:bump2=2:Me.TimerEnabled=1:vpmTimer.PulseSw 11:PlaySound SoundFX("jet1",DOFContactors),0,1,0.1,0:End If:End Sub
Sub Bumper2_Timer()
    Select Case bump2
        Case 1:Ring2.Z = 20:bump2 = 2
        Case 2:Ring2.Z = 30:bump2 = 3
        Case 3:Ring2.Z = 40:bump2 = 4
        Case 4:Ring2.Z = 50:Me.TimerEnabled = 0
    End Select
End Sub

Sub Bumper3_Hit:If PinPlay=1 Then:bump3=1:Me.TimerEnabled=1:vpmTimer.PulseSw 12:PlaySound SoundFX("jet1",DOFContactors),0,1,0,0:End If:End Sub
Sub Bumper3_Timer()
    Select Case bump3
        Case 1:Ring3.Z = 20:bump3 = 2
        Case 2:Ring3.Z = 30:bump3 = 3
        Case 3:Ring3.Z = 40:bump3 = 4
        Case 4:Ring3.Z = 50:Me.TimerEnabled = 0
    End Select
End Sub

' Spinner
Sub sw20_Spin:vpmTimer.PulseSw 20:PlaySound "spinner",0,1,-0.2,0:End Sub

' Eject holes
Sub Drain_Hit:bsTrough.AddBall Me:PlaySound "drain1a":End Sub
Sub sw21_Hit:PlaySound "kicker_enter",0,0.5,0,0:bsUK.AddBall 0:End Sub
Sub sw35_Hit:PlaySound "kicker_enter",0,0.5,-0.3,0:bsLK.AddBall 0:End Sub

' Rollovers
Sub sw23_Hit:  Controller.Switch(23) = 1:PlaySound "sensor", 0, 1, pan(ActiveBall):Psw23.Z=25:End Sub
Sub sw23_UnHit:Controller.Switch(23) = 0:Psw23.Z=40:End Sub
Sub sw24_Hit:  Controller.Switch(24) = 1:PlaySound "sensor", 0, 1, pan(ActiveBall):Psw24.Z=25:End Sub
Sub sw24_UnHit:Controller.Switch(24) = 0:Psw24.Z=40:End Sub
Sub sw30_Hit:  Controller.Switch(30) = 1:PlaySound "sensor", 0, 1, pan(ActiveBall):Psw30.Z=25:End Sub
Sub sw30_UnHit:Controller.Switch(30) = 0:Psw30.Z=40:End Sub
Sub sw31_Hit:  Controller.Switch(31) = 1:PlaySound "sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw31_UnHit:Controller.Switch(31) = 0:End Sub
Sub sw33_Hit:  Controller.Switch(33) = 1:PlaySound "sensor", 0, 1, pan(ActiveBall):Psw33.Z=25:End Sub
Sub sw33_UnHit:Controller.Switch(33) = 0:Psw33.Z=40:End Sub
Sub sw34_Hit:  Controller.Switch(34) = 1:PlaySound "sensor", 0, 1, pan(ActiveBall):Psw34.Z=25:End Sub
Sub sw34_UnHit:Controller.Switch(34) = 0:Psw34.Z=40:End Sub

'Ramp sensors
Sub sw50_Hit:  Controller.Switch(50) = 1:End Sub
Sub sw50_UnHit:Controller.Switch(50) = 0:End Sub
Sub sw51_Hit:  Controller.Switch(51) = 1:End Sub
Sub sw51_UnHit:Controller.Switch(51) = 0:End Sub
Sub sw52_Hit:  Controller.Switch(52) = 1:End Sub
Sub sw52_UnHit:Controller.Switch(52) = 0:End Sub
Sub sw52a_Hit:  Controller.Switch(52) = 1:End Sub
Sub sw52a_UnHit:Controller.Switch(52) = 0:End Sub

' Targets
Sub sw32_Hit:vpmTimer.PulseSw 32:Psw32.TransX=-5:Me.TimerEnabled=1:PlaySound SoundFX("target",DOFContactors),0,1,-0.2,0:End Sub
Sub sw32_Timer:Psw32.TransX=0:Me.TimerEnabled=0:End Sub

' Kicking Targets
Sub sw15_Slingshot:vpmTimer.PulseSw 15:KickingTsw15.ObjRotY=8:Me.TimerEnabled=1:PlaySound SoundFX("target",DOFContactors),0,1,-0.2,0:End Sub
Sub sw15_Timer:KickingTsw15.ObjRotY=0:Me.TimerEnabled=0:End Sub
Sub sw22_Hit:vpmTimer.PulseSw 22:KickingTsw22.ObjRotY=8:Me.TimerEnabled=1:PlaySound SoundFX("target",DOFContactors),0,1,0.2,0:End Sub
Sub sw22_Timer:KickingTsw22.ObjRotY=0:Me.TimerEnabled=0:End Sub

' Droptargets
Sub sw16_Hit():PlaySound SoundFX("DROPTARG",DOFContactors):End Sub
Sub sw16_Dropped:cdtbank.Hit 1: End Sub

Sub sw26_Hit():PlaySound SoundFX("DROPTARG",DOFContactors):End Sub
Sub sw26_Dropped:cdtbank.Hit 2:End Sub

Sub sw36_Hit():PlaySound SoundFX("DROPTARG",DOFContactors):End Sub
Sub sw36_Dropped:cdtbank.Hit 3:End Sub

Sub sw17_Hit():PlaySound SoundFX("DROPTARG",DOFContactors),0,1,0.2
 If Bulb12.State=1 Then
	Bulb12T17.State=1
Else
	Bulb12T17.State=0
End If
	Drop17T=1
End Sub
Sub sw17_Dropped:ddtbank.Hit 1:End Sub

Sub sw27_Hit():PlaySound SoundFX("DROPTARG",DOFContactors),0,1,0.2
 If Bulb12.State=1 Then
	Bulb12T27.State=1
Else
	Bulb12T27.State=0
End If
	Drop27T=1
End Sub
Sub sw27_Dropped:ddtbank.Hit 2:End Sub

Sub sw37_Hit():PlaySound SoundFX("DROPTARG",DOFContactors),0,1,0.2
 If Bulb11.State=1 Then
	Bulb11T37.State=1
Else
	Bulb11T37.State=0
End If
	Drop37T=1
End Sub
Sub sw37_Dropped:ddtbank.Hit 3:End Sub

' Gates
Sub Gate1_Hit:PlaySound "Gate5":End Sub

' Fx Sounds
Sub Trigger1_Hit:PlaySound "WireRolling", 0, 1, pan(ActiveBall), 0, 0, 0, 0:End Sub
Sub Trigger2_Hit:PlaySound "WireRolling", 0, 1, pan(ActiveBall), 0, 0, 0, 0:End Sub
Sub EndStop_Hit :PlaySound "WireHit", 0, 1, pan(ActiveBall), 0, 0, 0, 0:StopSound "WireRolling":End Sub
Sub EndStop1_Hit:PlaySound "WireHit", 0, 1, pan(ActiveBall), 0, 0, 0, 0:StopSound "WireRolling":End Sub

Sub TriggerFlapSx_Hit:PlaySound "FlapHit2", 0, 0.2*(Vol(ActiveBall)), pan(ActiveBall), 0, 0, 0, 0:End Sub
Sub TriggerFlapDx_Hit:PlaySound "FlapHit2", 0, 0.2*(Vol(ActiveBall)), pan(ActiveBall), 0, 0, 0, 0:End Sub
Sub TriggerOVUK_Hit:PlaySound "WoodHitD225",0,0.45,0.2,-2:End Sub
Sub Trigger4_Hit:PlaySound "PlasticJump",0,0.7*(Vol(ActiveBall)),0.3,0,-5000:End Sub
Sub StopVUK2_Hit:PlaySound "WireHit",0,0.8,-0.2,0:End Sub
Sub KickerW_Hit:PlaySound "kicker_enter",0,0.6,0,0:End Sub
Sub TriggerOVUK1_Hit:PlaySound "kicker_enter",0,0.6,-0.2,0:End Sub
Sub TriggerORamp_Hit:PlaySound "kicker_enter",0,0.6,0.2,0:End Sub
Sub CupFX1_Hit():PlaySound "PlasticJump",0,0.7*(Vol(ActiveBall)),0,0,Pitch(ActiveBall):End Sub
Sub CupFX2_Hit():PlaySound "PlasticJump",0,0.7*(Vol(ActiveBall)),0,0,Pitch(ActiveBall):End Sub

' add additional (optional) parameters to PlaySound to increase/decrease the frequency,
' apply all the settings to an already playing sample and choose if to restart this sample from the beginning or not
' PlaySound "name",loopcount,volume,pan,randompitch,pitch,useexisting,restart
' pan ranges from -1.0 (left) over 0.0 (both) to 1.0 (right)
' randompitch ranges from 0.0 (no randomization) to 1.0 (vary between half speed to double speed)
' pitch can be positive or negative and directly adds onto the standard sample frequency
' useexisting is 0 or 1 (if no existing/playing copy of the sound is found, then a new one is created)
' restart is 0 or 1 (only useful if useexisting is 1)

'*********
'Solenoids
'*********

SolCallback(7) = "bsLKBallRelease"
SolCallback(8) = "dtcbank"
SolCallback(9) = "dtdbank"
SolCallback(10) = "bsUK.SolOut"
SolCallback(15) = "setlamp 125,"
SolCallback(16) = "setlamp 150,"	'lamp
SolCallback(17) = "setlamp 197,"
SolCallback(18) = "setlamp 198,"
SolCallback(19) = "setlamp 199,"
SolCallback(20) = "setlamp 190,"
SolCallback(21) = "setlamp 171,"
SolCallback(22) = "setlamp 172,"
SolCallback(23) = "setlamp 122,"
SolCallback(24) = "setlamp 151,"
SolCallback(25) = "setlamp 175,"
SolCallback(28) = "bsTrough.SolOut"
SolCallback(29) = "bsTrough.SolIn"
SolCallback(30) = "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(31) = "GIRelay"
Solcallback(32) = "SolRun"

Sub dtcbank(Enabled)
 If Enabled Then
	cdtbank.DropSol_On
End If
	PlaySound SoundFX("DTResetB",DOFContactors)
End Sub

Sub dtdbank(Enabled)
 If Enabled Then
	ddtbank.DropSol_On
End If
	Drop17T=0
	Drop27T=0
	Drop37T=0
	Bulb12T17.TimerEnabled=1
	PlaySound SoundFX("DTResetB",DOFContactors),0,1,0.2
End Sub

Sub bsLKBallRelease(Enabled)
 If Enabled Then
	bsLK.ExitSol_On
	Arm.IsDropped=0
	Arm.TimerEnabled=1
	Arm2.IsDropped=0
End If
End Sub

Sub Arm_Timer
	Arm.IsDropped=1
	Arm.TimerEnabled=0
	Arm2.IsDropped=1
End Sub

Sub Bulb12T17_Timer()
	Bulb12T17.State=0
	Bulb12T27.State=0
	Bulb11T37.State=0
	Bulb12T17.TimerEnabled=0
End Sub

'**************
' GI
'**************

Dim Drop17T, Drop27T, Drop37T
Sub GIRelay(Enabled)
	Dim GIoffon
	GIoffon = ABS(ABS(Enabled) -1)
	SetLamp 160, GIoffon
 If Drop17T=1 Then
	Bulb12T17.State=GIoffon
End If
 If Drop27T=1 Then
	Bulb12T27.State=GIoffon
End If
 If Drop37T=1 Then
	Bulb11T37.State=GIoffon
End If
End Sub

Sub SolRun(Enabled)
	vpmNudge.SolGameOn Enabled
 If Enabled Then
	PinPlay=1
Else
	PinPlay=0
	LeftFlipper.RotateToStart
	RightFlipper.RotateToStart
End If
End Sub

'**************
' Flipper Subs
'**************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
	Controller.Switch(6) = ABS(enabled)
	If Enabled Then
		PlaySound SoundFX("flipperup_left",DOFContactors)	':LeftFlipper.RotateToEnd
	Else
        PlaySound SoundFX("flipperdown_left",DOFContactors)	':LeftFlipper.RotateToStart
	End If
End Sub

Sub SolRFlipper(Enabled)
	Controller.Switch(7) = ABS(enabled)
	If Enabled Then
		PlaySound SoundFX("flipperup_right",DOFContactors)	':RightFlipper.RotateToEnd
	Else
        PlaySound SoundFX("flipperdown_right",DOFContactors)	':RightFlipper.RotateToStart
	End If
End Sub

Sub LeftFlipper_Collide(parm)
	PlaySound "rubber_flipper",0,1,-0.8,0
End Sub

Sub RightFlipper_Collide(parm)
	PlaySound "rubber_flipper",0,1,0.8,0
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
LampTimer.Enabled = 1

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

    UpdateLamps
End Sub

Sub UpdateLamps
'	FadeL 0, l0, "On", "F66", "F33", "Off"
	NFadeL 0, l0
	NFadeL 1, l1
	NFadeL 2, l2
	NFadeL 3, l3
	NFadeL 4, l4
	NFadeL 5, l5
	NFadeL 6, l6
	NFadeL 7, l7
	NFadeL 10, l10
	NFadeL 11, l11
	NFadeL 12, l12
	NFadeL 13, l13
	NFadeL 14, l14
	NFadeL 15, l15
	NFadeL 16, l16
	NFadeL 17, l17
	NFadeL 20, l20
	NFadeL 21, l21
	NFadeL 22, l22
	NFadeL 23, l23
	NFadeL 24, l24
	NFadeL 25, l25
	NFadeL 26, l26
	NFadeL 27, l27
	NFadeL 30, l30
	NFadeL 31, l31
	NFadeL 32, l32
	NFadeL 33, l33
	NFadeL 34, l34
	NFadeL 35, l35
	NFadeL 36, l36
	NFadeL 37, l37
	NFadeL 40, l40
	NFadeL 41, l41

	NFadeL 67, l67
	NFadeL 70, l70
	NFadeL 71, l71
	NFadeL 72, l72
	NFadeL 73, l73
	NFadeL 74, l74
	NFadeL 75, l75
	NFadeL 76, l76
	NFadeL 80, l80
	NFadeL 81, l81
	NFadeL 82, l82
	NFadeL 83, l83
	NFadeL 84, l84
	NFadeL 85, l85
	NFadeL 86, l86
	NFadeL 90, l90
	NFadeL 91, l91
	NFadeL 92, l92
	NFadeL 93, l93
	NFadeL 94, l94
	NFadeL 95, l95
	NFadeL 96, l96
	NFadeL 100, l100
	NFadeL 101, l101
	NFadeL 102, l102
	NFadeL 103, l103
	NFadeL 104, l104
	NFadeL 105, l105
	NFadeL 106, l106
	NFadeL 110, l110
	NFadeL 111, l111
	NFadeL 112, l112
	NFadeL 113, l113
	NFadeL 114, l114
	NFadeL 115, l115
	NFadeL 116, l116
	NFadeL 122, l122
	NFadeL 125, l125

	Flash 42, f42
	Flashm 43, f43
	Flash 43, f43a
	Flash 44, f44
	Flash 45, f45
	Flash 46, f46
	Flash 47, f47
	Flashm 50, f50
	Flash 50, f50a
	Flash 51, f51
	Flash 52, f52
	Flash 53, f53
	Flash 54, f54
	Flash 55, f55
	Flashm 56, f56
	Flash 56, f56a
	Flashm 57, f57
	Flash 57, f57a
	Flashm 66, f66
	Flash 66, f66a
	NFadeLm 150, fur1
	NFadeLm 150, fur2
	Flashm 150, fur1a
	Flash 150, fur2a
	NFadeLm 151, f151
	NFadeL 151, f151a

	NFadeLm 160, Bulb1
	NFadeLm 160, Bulb1a
	NFadeLm 160, Bulb2
	NFadeLm 160, Bulb2a
	NFadeLm 160, Bulb3
	NFadeLm 160, Bulb3a
	NFadeLm 160, Bulb4
	NFadeLm 160, Bulb4a
	NFadeLm 160, Bulb5
	NFadeLm 160, Bulb5a
	NFadeLm 160, Bulb6
	NFadeLm 160, Bulb6a
	NFadeLm 160, Bulb7
	NFadeLm 160, Bulb7a
	NFadeLm 160, Bulb8
	NFadeLm 160, Bulb8a
	NFadeLm 160, Bulb9
	NFadeLm 160, Bulb9a
	NFadeLm 160, Bulb10
	NFadeLm 160, Bulb10a
	NFadeLm 160, Bulb11
	NFadeLm 160, Bulb11a
	NFadeLm 160, Bulb12
	NFadeLm 160, Bulb12a
	NFadeLm 160, Bulb13
	NFadeLm 160, Bulb13a
	NFadeLm 160, Bulb14
	NFadeLm 160, Bulb14a

	Flashm 160, fgit1
	Flashm 160, fgit1L
	Flashm 160, fgit1R
	Flashm 160, fgit2
	Flashm 160, fgit3
	Flashm 160, fgit4
	Flashm 160, fgit5
	Flashm 160, fgit6
	Flashm 160, fgit7
	Flashm 160, fgit8
	Flashm 160, fgit9
	Flashm 160, fgit10
	Flashm 160, fgit11
	Flashm 160, fgit12
	Flashm 160, fgit13
	Flashm 160, fgit14
	Flash 160, Bulb10b
	Flash 171, f171
	Flash 172, f172
	Flash 175, f175
	Flash 190, f190
	Flash 197, f197
	Flash 198, f198
	Flash 199, f199
End Sub

' div lamp subs

Sub InitLamps()
    Dim x
    For x = 0 to 200
        LampState(x) = 0        ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4      ' used to track the fading state
        FlashSpeedUp(x) = 0.5   ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.07	' slower speed when turning off the flasher
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
            Object.IntensityScale = Lumen*FlashLevel(nr)
        Case 5 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
            If FlashLevel(nr) > FlashMax(nr) Then
                FlashLevel(nr) = FlashMax(nr)
                FadingLevel(nr) = 1 'completely on
            End if
            Object.IntensityScale = Lumen*FlashLevel(nr)
    End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it just sets the flashlevel
    Object.IntensityScale = Lumen*FlashLevel(nr)
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

' *********************************************************************
' 					Wall, rubber and metal hit sounds
' *********************************************************************

Sub Rubbers_Hit(idx):PlaySound "rubber1", 0, Vol(ActiveBall), pan(ActiveBall), 0.25, AudioFade(ActiveBall):End Sub


'******************
' RealTime Updates
'******************
Set MotorCallback = GetRef("GameTimer")

Sub GameTimer
	UpdateFlipperLogos
	TrackSounds
End Sub

Dim PI
PI = Round(4 * Atn(1), 6) '3.1415926535897932384626433832795

Sub UpdateFlipperLogos
	FlipperSx.RotZ = LeftFlipper.CurrentAngle
	FlipperDx.RotZ = RightFlipper.CurrentAngle
	pSpinnerRod.TransY = sin( (sw20.CurrentAngle+180) * (2*PI/360)) * 8
	pSpinnerRod.TransZ = sin( (sw20.CurrentAngle- 90) * (2*PI/360)) * 8
	pSpinnerRod.RotX = sin( sw20.CurrentAngle * (2*PI/360)) * 6
End Sub

'********************************
' Sound Subs from Destruk's table
'********************************

Dim Playing
Playing = 0

'Music & Sound Stuff
Sub TrackSounds
    Dim NewSounds, ii, Snd
    NewSounds = Controller.NewSoundCommands
    If Not IsEmpty(NewSounds) Then
        For ii = 0 To UBound(NewSounds)
            Snd = NewSounds(ii, 0)
            If Snd = 58 Then ShootGently:LegsAnim
'            If Snd = 28 Then Heheheeee
        Next
    End If
End Sub

'Shoot Gently
Dim AnimStep
Sub ShootGently:AnimStep=1:TimerJaw.Enabled=1:End Sub
Sub TimerJaw_Timer()
	Select Case AnimStep
		Case 1:Jaw.ObjRotX = -8:AnimStep=2
		Case 2:Jaw.ObjRotX = 0 :AnimStep=3
		Case 3:Jaw.ObjRotX = -8:AnimStep=4
		Case 4:Jaw.ObjRotX = 0 :AnimStep=5
		Case 5:Jaw.ObjRotX = -8:AnimStep=6
		Case 6:Jaw.ObjRotX = 0 :TimerJaw.Enabled = 0
	End Select
End Sub

Dim AnimStep1
Sub LegsAnim:AnimStep1=1:TimerLegs.Enabled=1:End Sub
Sub TimerLegs_Timer()
	Select Case AnimStep1
		Case 1:LegSx.ObjRotX = -20:LegDx.ObjRotX = -20:AnimStep1=2
		Case 2:LegSx.ObjRotX = -40:LegDx.ObjRotX = -40:AnimStep1=3
		Case 3:LegSx.ObjRotX = -60:LegDx.ObjRotX = -60:AnimStep1=4
		Case 4:LegSx.ObjRotX = -40:LegDx.ObjRotX = -40:AnimStep1=5
		Case 5:LegSx.ObjRotX = -70:LegDx.ObjRotX = -60:AnimStep1=6
		Case 6:LegSx.ObjRotX = -30:LegDx.ObjRotX = -30:AnimStep1=7
		Case 7:LegSx.ObjRotX = 0:LegDx.ObjRotX = 0:TimerLegs.Enabled = 0
	End Select
End Sub

'He he heeee
'Dim AnimStep1
'Sub Heheheeee:AnimStep1=1:TimerJaw.Enabled=1:End Sub
'Sub TimerJaw_Timer
'	Select Case AnimStep1
'		Case 1:Jaw.ObjRotX = -6:AnimStep1=2
'		Case 2:Jaw.ObjRotX = 0 :AnimStep1=3
'		Case 3:Jaw.ObjRotX = -6:AnimStep1=4
'		Case 4:Jaw.ObjRotX = 0 :AnimStep1=5
'		Case 5:Jaw.ObjRotX = -6:AnimStep1=6
'		Case 6:AnimStep1=7 'pause
'		Case 7:AnimStep1=8 'pause
'		Case 8:Jaw.ObjRotX = 0:TimerJaw.Enabled = 0
'	End Select
'End Sub


'006  06-Excellent
'007  07-Extraball
'020  14-Million
'021  15-Advance souvenirs target for extraball
'023  17-The arcade target is lit
'024  18-Lit the birth
'025  19-You wont to pet my monkey
'026  1a-Ho ho ho ho ha ha ha ha
'027  1b-Le target a ?
'028  1c-Nice shoot man
'029  1d-Shoot the rapides to light the whirlpool miilion
'030  1e-The whirlpool million is lit
'033  21-You got absolutly nothing
'034  22-You advance the boomerang
'035  23-You advance the splash
'036  24-You advance the whirlpool
'037  25-You advance the rapides
'038  26-You advance the pipeline
'039  27-Shoot the ball
'040  28-He he heeee
'046  2e-Shoot all lights lit to light super jackpot
'047  2f-Shoot all lights lit to light jackpot
'048  30-Shoot the rapides for super jackpot
'049  31-Shoot the rapides for jackpot
'051  33-Shoot all lights lit for extraball
'052  34-Shoot the spinner
'053  35-Your time is up
'054  36-Hit the three million target
'056  38-Hey man shoot again
'057  39-You shoot to hard
'058  3a-Shoot gently
'059  3b-Shoot the pipeline for multiball
'060  3c-One more for the record
'061  3d-One more for the whirlpool
'062  3e-One more for the boomerang
'063  3f-One more for the splash
'065  41-One more for the rapides
'066  42-One more for the pipeline
'067  43-One more for double
'068  44-you shoot too sofly
'070  46-Extra special
'071  47-Super jackpot
'072  48-super score
'073  49-Three million
'074  4a-Super
'077  4d-Rodney
'078  4e-Two
'079  4f-Three
'080  50-Four
'081  51-Five
'085  55-Got all animals for special
'089  59-Oh noo
'105  69-Another coin please
'107  6b-Jackpot
'109  6d-Feel the power

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

Const tnob = 4 ' total number of balls
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


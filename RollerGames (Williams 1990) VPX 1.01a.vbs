Option Explicit
Randomize

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' No special SSF tweaks yet.

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Dim Ballsize,BallMass
Ballsize = 50
BallMass = (Ballsize^3)/125000


Const UseSolenoids = 2
Const UseLamps = 0
Const UseGI = 0
Const UseSync = 0
Const HandleMech = 0

' Standard Sounds
Const SSolenoidOn = "fx_Solenoid"
Const SSolenoidOff = ""
Const SCoin = "fx_Coin"

Dim DesktopMode: DesktopMode = RollerGames.ShowDT
Dim UseVPMColoredDMD

if DesktopMode = True Then
'   vidrio.visible = 1
'   leftrail.visible = 1
'   rightrail1.visible = 1
'   backrail.visible = 1
 Else
'   vidrio.visible = 0
'   leftrail.visible = 1
'   rightrail1.visible = 0
'   backrail.visible = 0
End If

 LoadVPM "01120100", "S11.VBS", 3.10


'Sub Kicker1_hit:me.kick 342,40 End Sub
'Sub Kicker2_hit:me.kick 354,60 End Sub


' Solenoid Callbacks
' ------------------
SolCallback(1)     		= "bsTrough.SolIn"							      ' ball into outhole
SolCallback(2)    		= "bsTrough.SolOut"							      ' ball into shooter lane
SolCallback(4)			= "bsVuk.SolOut" 								  ' Pit popper
SolCallback(6) 			= "dtDrop.SolDropUp"							  ' Drop "W-A-R" targets
SolCallback(7)          = "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(8) 	        = "vpmSolDiverter LockDiverter,true," 		      ' 8 - Lock Diverter
SolCallback(11) 		= "SolGi"								          ' 11 - GI circuit
SolCallBack(13)         = "bsULock.SolOut"		                          ' UpperKickBack fire
SolCallback(14)         = "solKickback"								      ' Eject balls from lock
SolCallback(20)         = "SolAutoPlunge"
'SolCallBack(18)	    	= "vpmSolgate Array(Gate, Gate3),""fx_diverter"","' 18, Ramp Diverter
SolCallBack(18)	    	= "flipdiverters"' 18, Ramp Diverter



SolCallback(9) = "SolFlasher9"  'Flash 9B
SolCallback(26) = "SolFlasher26"  'Flash 3B
SolCallback(27) = "SolFlasher27"  'Flash 4R
SolCallback(28) = "SolFlasher28"  'Flash 5R
SolCallback(29) = "SolFlasher29"  'Flash 6R
SolCallback(30) = "SolFlasher30"  'Flash 7R
SolCallback(31) = "SolFlasher31"  'Flash 8R



 Const cGameName    = "rollr_l2"

Dim bsTrough, Iman, bsVuk, dtDrop, bslLock, bsULock
Dim x
Sub RollerGames_Init
	vpmInit Me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Black Knight 2000 - Williams 1989" & vbNewLine & "VPX by Flupper"
        .Games(cGameName).Settings.Value("rol") = 0 'set it to 1 to rotate the DMD to the left
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = 0
        On Error Resume Next
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
        On Error Goto 0
    End With

    ' Nudging
    vpmNudge.TiltSwitch = swTilt
  '  vpmNudge.Sensitivity = 1
	vpmNudge.TiltObj = Array(Bumper1b,Bumper2b,Bumper3b,LeftslingShot,RightslingShot)

    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1


    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw  9, 11, 12,13, 0, 0, 0, 0
        .InitKick BallRelease, 90, 6
        .InitEntrySnd "fx_Solenoid", "fx_Solenoid"
        .InitExitSnd SoundFX("fx_ballrel",DOFContactors), SoundFX("fx_Solenoid",DOFContactors)
        .Balls = 3
    End With

	' Lock Section
	Set bslLock = New cvpmBallStack
	With bslLock
	     .initsw 0,53,54,0,0,0,0,0
	     .InitKick KickBack, 0, 55
	     .KickBalls=3
    End With

	Set bsULock = New cvpmBallStack
	With bsULock
	     .InitSaucer UpperKickBack, 14, 0, 50
	    .InitEntrySnd "Solenoid", "Solenoid"
         .InitExitSnd "bumper_retro", "Solenoid"
    End With

	'Setup magnets
	Set Iman = New cvpmMagnet
	With Iman
	     .InitMagnet imanT, 70
	     .Solenoid = 22
	     .CreateEvents "Iman"
    End With

    'vuk
    Set bsVuk = New cvpmBallStack
    With bsVuk
        .InitSaucer Sw16, 16, 0, 32
        .KickZ = 1.56
	    .InitEntrySnd "fx_kicker2", "Solenoid"
        .InitExitSnd "fx_vukout_LAH", "Solenoid"
    End With


    'droptargets
    Set dtDrop = New cvpmDropTarget
    With dtDrop
        .InitDrop Array(Sw49,Sw50,Sw51),Array(49,50,51)
        .initsnd SoundFX("fx_droptarget", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
        .CreateEvents "dtDrop"
    End With
  UpperKickBack1.TransY = 0
End Sub

' Right KickBack Lock
Sub SolKickBack(Enabled)
	If Enabled Then
		bslLock.ExitSol_On
        LaserKickP1.TransY = 90
        Playsound "bumper_retro"
      Else
        LaserKickP1.TransY = 0
        PlaySound "fx_rubber"
	End If
End Sub
Sub KickBack_Hit: bslLock.AddBall Me: End Sub
Sub KickBack_UnHit: Playsound "fx_railShort": End Sub


' Left UpperKickBack
Sub SolUpperKickBack(dummy)
Dim KickBack: KickBack = 1
    bsULock.AddBall 0
	If KickBack = 1 Then
        UpperKickBack1.TransY = 90
        vpmTimer.AddTimer 700, "UpperKickBackRes"
	End If
End Sub
Sub UpperKickBackRes(dummy):UpperKickBack1.TransY = 0:Playsound "fx_rubber" End Sub
Sub UpperKickBack_Hit: vpmTimer.AddTimer 1500, "SolUpperKickBack": End Sub



' LaserKick
Sub SolAutoPlunge(Enabled)
	If Enabled Then
        Playsound "bumper_retro"
		LaserKick.Enabled=True
        LaserKickP.TransY = 100
	Else
		LaserKick.Enabled=False
        vpmTimer.AddTimer 800, "AutoPlungeRes"
	End If
End Sub
Sub AutoPlungeRes(dummy):LaserKickP.TransY = 0:Playsound "fx_rubber" End Sub
Sub LaserKick_Hit: Me.Kick 0,40 End Sub


'Diverters
Sub Flipdiverters(enabled)
	If Enabled Then
		RampDiverter.RotateToEnd
		RampDiverter2.RotateToEnd
		playsound "fx_diverter"
	Else
		RampDiverter.RotateToStart
		RampDiverter2.RotateToStart
		playsound "fx_diverter"
	End If
End Sub


'*********
'Flashers
'*********

Sub SolFlasher9(enabled)
 If enabled Then
    Flasher9.opacity = 100
    Flasher9a.state = 1
  Else
    Flasher9.opacity =  0
    Flasher9a.state = 0
 End If
End Sub


Sub SolFlasher26(enabled)
 If enabled Then
    Flasher3.opacity = 100
    Flasher3a.state = 1
  Else
    Flasher3.opacity = 0
    Flasher3a.state = 0
 End If
End Sub

Sub SolFlasher27(enabled)
 If enabled Then
    Flasher4.opacity = 100
    Flasher4a.state = 1
  Else
    Flasher4.opacity =  0
    Flasher4a.state = 0
 End If
End Sub


Sub SolFlasher28(enabled)
 If enabled Then
    Flasher5.opacity = 100
    Flasher5a.state = 1
  Else
    Flasher5.opacity = 0
    Flasher5a.state = 0
 End If
End Sub

Sub SolFlasher29(enabled)
 If enabled Then
    Flasher6.opacity = 100
    Flasher6a.state = 1
  Else
    Flasher6.opacity =  0
    Flasher6a.state = 0
 End If
End Sub


Sub SolFlasher30(enabled)
 If enabled Then
    Flasher7.opacity = 100
    Flasher7a.state = 1
  Else
    Flasher7.opacity = 0
    Flasher7a.state = 0
 End If
End Sub

Sub SolFlasher31(enabled)
 If enabled Then
    Flasher8.opacity = 100
    Flasher8a.state = 1
  Else
    Flasher8.opacity =  0
    Flasher8a.state = 0
 End If
End Sub






'**********
' Keys
'**********

Sub RollerGames_KeyDown(ByVal Keycode)
    If keycode = PlungerKey Then PlaySound "fx_PlungerPull", 0, 1, 0.1, 0.05:Plunger.Pullback
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub RollerGames_KeyUp(ByVal Keycode)
    If keycode = PlungerKey Then PlaySound "fx_plunger", 0, 1, 0.1, 0.05:Plunger.Fire
    If vpmKeyUp(keycode) Then Exit Sub
End Sub


'*************
' Pause Table
'*************

Sub RollerGames_Paused
End Sub

Sub RollerGames_unPaused
End Sub


'**************
' Flipper Subs
'**************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySound SoundFX("fx_flipperupl", DOFFlippers), 0, 1, -0.1, 0.15
        LeftFlipper.RotateToEnd
    Else
        PlaySound SoundFX("fx_flipperdown", DOFFlippers), 0, 1, -0.1, 0.15
        LeftFlipper.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySound SoundFX("fx_flipperupr", DOFFlippers), 0, 1, 0.1, 0.15
        RightFlipper.RotateToEnd
        RightFlipper1.RotateToEnd
    Else
        PlaySound SoundFX("fx_flipperdown", DOFFlippers), 0, 1, 0.1, 0.15
        RightFlipper.RotateToStart
        RightFlipper1.RotateToStart
    End If
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, -0.1, 0.15
End Sub

Sub Rightflipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, 0.1, 0.15
End Sub

Sub LeftFlipper1_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, -0.1, 0.15
End Sub




Sub UpdateFlipperLogo_Timer
    LFLogo.RotY = LeftFlipper.CurrentAngle
    RFlogo.RotY = RightFlipper.CurrentAngle
    RFlogo1.RotY = RightFlipper1.CurrentAngle
End Sub







'*********
' Switches
'*********

' Slings & div switches

Dim LStep, LStep1, RStep

Sub LeftSlingShot_Slingshot
    PlaySound SoundFX("fx_slingshotl", DOFContactors), 0, 1, -0.05, 0.05
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 33
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 1:LeftSLing4.Visible = 0:LeftSLing3.Visible = 1:Lemk.RotX = 14
        Case 2:LeftSLing3.Visible = 0:LeftSLing2.Visible = 1:Lemk.RotX = 2
        Case 3:LeftSLing2.Visible = 0:Lemk.RotX = -10:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub


Sub RightSlingShot_Slingshot
    PlaySound SoundFX("fx_slingshotr", DOFContactors), 0, 1, 0.05, 0.05
    RightSling4.Visible = 1
    Remk.RotX = 26
    RStep = 0
    vpmTimer.PulseSw 34
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 1:RightSLing4.Visible = 0:RightSLing3.Visible = 1:Remk.RotX = 14
        Case 2:RightSLing3.Visible = 0:RightSLing2.Visible = 1:Remk.RotX = 2
        Case 3:RightSLing2.Visible = 0:Remk.RotX = -10:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

' Bumpers
Sub Bumper1B_Hit:vpmTimer.PulseSw 35:PlaySound SoundFX("fx_bumper1",DOFContactors), 0, 1, -0.1, 0.15:End Sub
Sub Bumper2B_Hit:vpmTimer.PulseSw 36:PlaySound SoundFX("fx_bumper2",DOFContactors), 0, 1, 0, 0.15:End Sub
Sub Bumper3B_Hit:vpmTimer.PulseSw 37:PlaySound SoundFX("fx_bumper3",DOFContactors), 0, 1, 0.1, 0.15:End Sub


' Drain & holes
Sub Drain_Hit:Playsound "fx_drain":bsTrough.AddBall Me:End Sub

Sub sw14_Hit:Controller.Switch(14) = 1:PlaySound "fx_sensor", 0, 1, -0.1, 0.15:End Sub
Sub sw14_UnHit:Controller.Switch(14) = 0:End Sub

Sub Sw16_Hit(): PlaySound "fx_sensor": bsVuk.AddBall 0:End Sub


Sub sw18_Hit:Controller.Switch(18) = 1:PlaySound "fx_sensor", 0, 1, -0.1, 0.15:End Sub
Sub sw18_UnHit:Controller.Switch(18) = 0:End Sub

Sub sw19_Hit:Controller.Switch(19) = 1:PlaySound "fx_sensor", 0, 1, -0.1, 0.15:End Sub
Sub sw19_UnHit:Controller.Switch(19) = 0:End Sub


Sub sw20_Hit:Controller.Switch(20) = 1:PlaySound "fx_sensor", 0, 1, -0.1, 0.15:End Sub
Sub sw20_UnHit:Controller.Switch(20) = 0:End Sub

Sub sw21_Hit:Controller.Switch(21) = 1:PlaySound "fx_sensor", 0, 1, -0.1, 0.15:End Sub
Sub sw21_UnHit:Controller.Switch(21) = 0:End Sub


Sub sw22_Hit:Controller.Switch(22) = 1:PlaySound "fx_sensor", 0, 1, -0.1, 0.15:End Sub
Sub sw22_UnHit:Controller.Switch(22) = 0:End Sub

Sub Sw23_Spin(): vpmTimer.PulseSwitch 23,0,0 :PlaySound "fx_spinner", 0, 1, -0.1, 0.15:End Sub

Sub Sw24_Spin(): vpmTimer.PulseSwitch 24,0,0 :PlaySound "fx_spinner", 0, 1, -0.1, 0.15:End Sub


Sub sw25_Hit:vpmTimer.PulseSw 25:PlaySound "fx_postrubber", 0, 1, -0.1, 0.15:End Sub
Sub sw26_Hit:vpmTimer.PulseSw 26:PlaySound "fx_postrubber", 0, 1, -0.1, 0.15:End Sub
Sub sw27_Hit:vpmTimer.PulseSw 27:PlaySound "fx_postrubber", 0, 1, -0.1, 0.15:End Sub
Sub sw28_Hit:vpmTimer.PulseSw 28:PlaySound "fx_postrubber", 0, 1, -0.1, 0.15:End Sub
Sub sw29_Hit:vpmTimer.PulseSw 29:PlaySound "fx_postrubber", 0, 1, -0.1, 0.15:End Sub


Sub sw30_Hit:vpmTimer.PulseSw 30:PlaySound "fx_postrubber", 0, 1, -0.1, 0.15:End Sub
Sub sw31_Hit:vpmTimer.PulseSw 31:PlaySound "fx_postrubber", 0, 1, -0.1, 0.15:End Sub
Sub sw32_Hit:vpmTimer.PulseSw 32:PlaySound "fx_postrubber", 0, 1, -0.1, 0.15:End Sub

Sub Sw38_Hit():vpmTimer.PulseSwitch(38),0,0:LaserKickP2.TransY = 70:CoilWall1.isdropped = 1:CoilWall2.isdropped = 0 :vpmtimer.AddTimer 800, "CoilAnim '": PlaySound SoundFX("fx_sensor",DOFContactors), 0, 1, 0.1, 0.15:End Sub
Sub CoilAnim():LaserKickP2.TransY = 0 : CoilWall1.isdropped = 0 : CoilWall2.isdropped = 1 End Sub

Sub Sw39_Hit():vpmTimer.PulseSwitch(39),0,0:PlaySound SoundFX("fx_sensor",DOFContactors), 0, 1, 0.1, 0.15:End Sub

Sub sw40_Hit:Controller.Switch(40) = 1:PlaySound "fx_sensor", 0, 1, -0.1, 0.15:End Sub
Sub sw40_UnHit:Controller.Switch(40) = 0:End Sub

Sub sw41_Hit:vpmTimer.PulseSw 41:PlaySound "fx_postrubber", 0, 1, -0.1, 0.15:End Sub
Sub sw42_Hit:vpmTimer.PulseSw 42:PlaySound "fx_postrubber", 0, 1, -0.1, 0.15:End Sub
Sub sw43_Hit:vpmTimer.PulseSw 43:PlaySound "fx_postrubber", 0, 1, -0.1, 0.15:End Sub

Sub sw44_Hit:Controller.Switch(44) = 1:PlaySound "fx_sensor", 0, 1, -0.1, 0.15:End Sub
Sub sw44_UnHit:Controller.Switch(44) = 0:End Sub


Sub sw46_Hit:Controller.Switch(46) = 1:PlaySound "fx_sensor", 0, 1, -0.1, 0.15:End Sub
Sub sw46_UnHit:Controller.Switch(46) = 0:End Sub


Sub sw47_Hit:Controller.Switch(47) = 1:PlaySound "fx_sensor", 0, 1, -0.1, 0.15:End Sub
Sub sw47_UnHit:Controller.Switch(47) = 0:End Sub

Sub sw48_Hit:Controller.Switch(48) = 1:PlaySound "fx_sensor", 0, 1, -0.1, 0.15:End Sub
Sub sw48_UnHit:Controller.Switch(48) = 0:End Sub

Sub sw52_Hit:Controller.Switch(52) = 1:PlaySound "fx_metalhit", 0, 1, -0.1, 0.15:End Sub
Sub sw52_UnHit:Controller.Switch(52) = 0:vpmTimer.AddTimer 400, "RampFX": End Sub

Sub sw56_Hit:Controller.Switch(56) = 1:PlaySound "fx_sensor", 0, 1, -0.1, 0.15:End Sub
Sub sw56_UnHit:Controller.Switch(56) = 0:End Sub

Sub sw49_dropped():dtDrop.Hit 1:End Sub
Sub sw50_dropped():dtDrop.Hit 2:End Sub
Sub sw51_dropped():dtDrop.Hit 3:End Sub







'***************************************************
'       JP's VP10 Fading Lamps & Flashers
'       Based on PD's Fading Light System
' SetLamp 0 is Off
' SetLamp 1 is On
' fading for non opacity objects is 4 steps
'***************************************************

Dim bulb
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

Sub InitLamps()
    Dim x
    For x = 0 to 200
        LampState(x) = 0         ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4       ' used to track the fading state
        FlashSpeedUp(x) = 0.5    ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.35 ' slower speed when turning off the flasher
        FlashMax(x) = 1          ' the maximum value when on, usually 1
        FlashMin(x) = 0          ' the minimum value when off, usually 0
        FlashLevel(x) = 0        ' the intensity of the flashers, usually from 0 to 1
    Next
End Sub



Sub UpdateLamps
'    On Error Resume Next

'	nFadeL 1, Light1
'	nFadeL 2, Light2
'	nFadeL 3, Light3
'	nFadeL 4, Light4
'	nFadeL 5, Light5
'	nFadeL 6, Light6
'	nFadeL 7, Light7
'	nFadeL 8, Light8

	nFadeL 9, Light9
	nFadeL 10, Light10
	nFadeL 11, Light11
	nFadeL 12, Light12
	nFadeL 13, Light13
	nFadeL 14, Light14
	nFadeL 15, Light15
	nFadeL 16, Light16
	nFadeL 17, Light17
	nFadeL 18, Light18
	nFadeL 19, Light19
	nFadeL 20, Light20
	nFadeL 21, Light21
	nFadeL 22, Light22
	nFadeL 23, Light23
	nFadeL 24, Light24
	nFadeL 25, Light25
	nFadeL 26, Light26
	nFadeL 27, Light27
	nFadeL 28, Light28
	nFadeL 29, Light29
	nFadeL 30, Light30
	nFadeL 31, Light31
	nFadeL 32, Light32
	nFadeL 33, Light33
	nFadeL 34, Light34
	nFadeL 35, Light35
	nFadeL 36, Light36
	nFadeL 37, Light37
	nFadeL 38, Light38
	nFadeL 39, Light39
	NFadeLm 40, Light40
	NFadeLm 40, Light40a
	nFadeL 41, Light41
	nFadeL 42, Light42
	nFadeL 43, Light43
	nFadeL 44, Light44
	nFadeL 45, Light45
	nFadeL 46, Light46
	nFadeL 47, Light47
	nFadeL 48, Light48
	nFadeL 49, Light49
	nFadeL 50, Light50
	nFadeL 51, Light51
	nFadeL 52, Light52
	nFadeL 53, Light53
	nFadeL 54, Light54
	nFadeLm 55, Light55
	nFadeL 55, Light67
	nFadeL 56, Light56
	nFadeL 57, Light57
	nFadeL 58, Light58
	nFadeL 59, Light59
	nFadeL 60, Light60
	nFadeL 61, Light61
	nFadeL 62, Light62
	nFadeL 63, Light63
	nFadeL 64, Light64

End Sub




Sub SetLamp(nr, value)
    If value <> LampState(nr) Then
        LampState(nr) = abs(value)
        FadingLevel(nr) = abs(value) + 4
    End If
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

' Walls

Sub FadeWS(nr, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:a.IsDropped = 1:b.IsDropped = 1:c.IsDropped = 1:d.IsDropped = 0:FadingLevel(nr) = 0 'Off
        Case 3:a.IsDropped = 1:b.IsDropped = 0:c.IsDropped = 1:d.IsDropped = 1:FadingLevel(nr) = 2 'fading...
        Case 4:a.IsDropped = 1:b.IsDropped = 1:c.IsDropped = 0:d.IsDropped = 1:FadingLevel(nr) = 3 'fading...
        Case 5:a.IsDropped = 0:b.IsDropped = 1:c.IsDropped = 1:d.IsDropped = 1:FadingLevel(nr) = 1 'ON
    End Select
End Sub

'Lights, Ramps & Primitives used as 4 step fading lights
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
        Case 4:object.image = b:FadingLevel(nr) = 0
        Case 5:object.image = a:FadingLevel(nr) = 1
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

' RGB Leds

Sub RGBLED (object,red,green,blue)
object.color = RGB(0,0,0)
object.colorfull = RGB(2.5*red,2.5*green,2.5*blue)
object.state=1
End Sub

' Modulated Flasher and Lights objects

Sub SetLampMod(nr, value)
    If value > 0 Then
		LampState(nr) = 1
	Else
		LampState(nr) = 0
	End If
	FadingLevel(nr) = value
End Sub

Sub FlashMod(nr, object)
	Object.IntensityScale = FadingLevel(nr)/255
End Sub

Sub LampMod(nr, object)
Object.IntensityScale = FadingLevel(nr)/255
Object.State = LampState(nr)
End Sub




Sub SolGi(enabled)
  If enabled Then
     GiOFF
	Playsound "fx_relay_off"
	RollerGames.ColorGradeImage = "ColorGrade_1"
   Else
     GiON
	Playsound "fx_relay_on"
	RollerGames.ColorGradeImage = "ColorGrade_8"
 End If
End Sub



Sub GiON
    For each x in aGiLights
        x.State = 1
    Next
End Sub

Sub GiOFF
    For each x in aGiLights
        x.State = 0
    Next
End Sub



'****************************************
' Real Time updatess using the GameTimer
'****************************************
'used for all the real time updates

Sub GameTimer_Timer
    RollingUpdate
End Sub
'******************************
' Diverse Collection Hit Sounds
'******************************

Sub aMetals_Hit(idx):PlaySound "fx_MetalHit":End Sub
Sub aRubber_Bands_Hit(idx):PlaySound "fx_rubber":End Sub
Sub aRubber_Posts_Hit(idx):PlaySound "fx_postrubber":End Sub
Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit":End Sub
Sub aGates_Hit(idx):PlaySound "fx_Gate":End Sub
Sub aWoods_Hit(idx):PlaySound "fx_Woodhit":End Sub




' Ramp Soundss
Sub RHelp1_Hit()
    StopSound "fx_railShort"
    PlaySound "balldrop", 0, 1, pan(ActiveBall)
End Sub

Sub RHelp2_Hit()
    PlaySound "balldrop", 0, 1, pan(ActiveBall)
End Sub

Sub RHelp3_Hit()
    StopSound "fx_railShort"
    PlaySound "fx_rampR", 0, 1, pan(ActiveBall)
End Sub

Sub RHelp4_Hit()
    StopSound "rail"
    PlaySound "balldrop", 0, 1, pan(ActiveBall)
End Sub

Sub RampFX(dumy)
    PlaySound "fx_railShort"
End Sub


Sub Trigger1_hit:Playsound "fx_railShort" End Sub

Sub RLFX_hit:PlaySound "fx_rampR", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub

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

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "RollerGames" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / RollerGames.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "RollerGames" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / RollerGames.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "RollerGames" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / RollerGames.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioFade(ball) ' Can this be together with the above function ?
  Dim tmp
  tmp = ball.y * 2 / RollerGames.height-1
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

Const tnob = 20 ' total number of balls
Const lob = 0   'number of locked balls
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub RollingUpdate()
    Dim BOT, b, ballpitch
    BOT = GetBalls

    ' stop the sound of deleted balls
    For b = UBound(BOT) + 1 to tnob
        rolling(b) = False
        StopSound("fx_ballrolling" & b)
    Next

    ' exit the sub if no balls on the table
    If UBound(BOT) = 3 Then Exit Sub 'there are always 4 balls on this table

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
  If RollerGames.VersionMinor > 3 OR RollerGames.VersionMajor > 10 Then
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
  Else
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0
  End if
End Sub


' Thalamus : Exit in a clean and proper way
Sub RollerGames_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


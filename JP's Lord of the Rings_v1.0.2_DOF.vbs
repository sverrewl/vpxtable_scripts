' The Lord of the Rings / IPD No. 4858 / October, 2003 / 4 Players
' VPX version by jpsalas
' DOF commands by arngrim

' Thalamus 2018-08-05
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' Thalamus 2018-08-09 : Improved directional sounds

Option Explicit
Randomize

' Options
' Volume devided by - lower gets higher sound

Const VolDiv = 200

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const BallSize = 50

Dim VarHidden, UseVPMColoredDMD
If Table1.ShowDT = true then
    UseVPMColoredDMD = true
    VarHidden = 1
else
    UseVPMColoredDMD = False
    VarHidden = 0
    TextBox1.Visible = 0
end if

LoadVPM "01560000", "SEGA.VBS", 3.26

'********************
'Standard definitions
'********************

Const UseSolenoids = 2
' Thal : Added because of useSolenoid=2
Const cSingleLFlip = 0
Const cSingleRFlip = 0
Const UseLamps = 0
Const UseGI = 0
Const UseSync = 0
Const HandleMech = 0

' Standard Sounds
Const SSolenoidOn = "fx_Solenoid"
Const SSolenoidOff = ""
Const SCoin = "fx_Coin"

Dim x, i, j, k 'used in loops
Dim bsTrough, bsL, bsR, bsTR, bsTL, vlLock, mRingMagnet, plungerIM

'************
' Table init.
'************

' choose the ROM
Const cGameName = "lotr"    'USA
'Const cGameName = "lotr_fr" 'France
'Const cGameName = "lotr_gr" 'Germany
'Const cGameName = "lotr_it" 'Italy
'Const cGameName = "lotr_sp" 'Spain


Sub Table1_Init
    vpmInit Me
    With Controller
        .GameName = cGameName
        .Games(cGameName).Settings.Value("sound") = 1 'ensure the sound is on
        If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "The Lord of the Rings - Stern 2003" & vbNewLine & "VPX table by JPSalas v1.0.2"
        .Games(cGameName).Settings.Value("rol") = 0
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = VarHidden
        On Error Resume Next
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description

        On Error Goto 0
    End With

    ' Nudging
    vpmNudge.TiltSwitch = 56
    vpmNudge.Sensitivity = 1
    vpmNudge.TiltObj = Array(bumper1, bumper2, bumper3, LeftSlingshot, RightSlingshot)

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0, 14, 13, 12, 11, 0, 0, 0
        .InitKick BallRelease, 90, 3
        .InitEntrySnd "fx_Solenoid", "fx_Solenoid"
        .InitExitSnd SoundFX("fx_ballrel",DOFContactors), SoundFX("fx_Solenoid",DOFContactors)
        .Balls = 4
    End With

    ' Left vuk
    Set bsL = New cvpmBallStack
    With bsL
        .InitSaucer sw9, 9, 0, 30
        .KickZ = 1.56
        .InitExitSnd SoundFX("fx_popper",DOFContactors), SoundFX("fx_Solenoid",DOFContactors)
    End With

    ' Right vuk
    Set bsR = New cvpmBallStack
    With bsR
        .InitSaucer sw30, 30, 0, 30
        .KickZ = 1.56
        .InitExitSnd SoundFX("fx_popper",DOFContactors), SoundFX("fx_Solenoid",DOFContactors)
    End With

    ' Top Right Saucer
    Set bsTR = New cvpmBallStack
    With bsTR
        .InitSaucer sw46, 46, 270, 3
        .KickForceVar = 2
        .InitExitSnd SoundFX("fx_popper",DOFContactors), SoundFX("fx_popper",DOFContactors)
    End With

    ' Top Left vuk
    Set bsTL = New cvpmBallStack
    With bsTL
        .InitSw 0, 41, 0, 0, 0, 0, 0, 0
        .InitKick sw41, 270, 6
        .InitExitSnd SoundFX("fx_popper",DOFContactors), SoundFX("fx_Solenoid",DOFContactors)
        .KickForceVar = 2
    End With

    ' Visible Lock - implements post ball lock
    Set vlLock = New cvpmVLock
    With vlLock
        .InitVLock Array(sw19, sw18, sw17), Array(sw19k, sw18k, sw17k), Array(19, 18, 17)
        .InitSnd "fx_sensor", "fx_sensor"
        .CreateEvents "vlLock"
    End With

    ' Impulse Plunger - used as the autoplunger
    Const IMPowerSetting = 60 ' Plunger Power
    Const IMTime = 0.7        ' Time in seconds for Full Plunge
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swPlunger, IMPowerSetting, IMTime
        .Random 0.3
        .switch 16
        .InitExitSnd "fx_kicker", "fx_kicker"
        .CreateEvents "plungerIM"
    End With

	Set mRingMagnet = New cvpmMagnet
 	With mRingMagnet
		.InitMagnet sw47a, 30
		.GrabCenter = True
 		.solenoid = 6						'Ring Magnet
        .CreateEvents "mRingMagnet"
	End With

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

    ' Div Init
    OrbitPin.IsDropped = 1
    SolBalrog 0
End Sub

Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub

'****************
' Handle Options
'****************

Sub JPLOTRSetOptions
End Sub

'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound SoundFX("fx_nudge",0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound SoundFX("fx_nudge",0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound SoundFX("fx_nudge",0), 0, 1, 0, 0.25
    If keycode = PlungerKey Then PlaySoundAt "fx_PlungerPull", Plunger:Plunger.Pullback
    If KeyDownHandler(KeyCode)Then Exit Sub
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If KeyUpHandler(KeyCode)Then Exit Sub
    If keycode = PlungerKey Then PlaySoundAt "fx_plunger", Plunger:Plunger.Fire
End Sub

'*********
' Switches
'*********

' Slings & div switches

Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    PlaySoundAt SoundFX("fx_slingshot",DOFContactors), Lemk
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 59
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
    PlaySoundAt SoundFX("fx_slingshot",DOFContactors), Remk
    RightSling4.Visible = 1
    Remk.RotX = 26
    RStep = 0
    vpmTimer.PulseSw 62
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
Sub Bumper1_Hit:vpmTimer.PulseSw 49:PlaySoundAt SoundFX("fx_bumper",DOFContactors), Bumper1:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 50:PlaySoundAt SoundFX("fx_bumper",DOFContactors), Bumper2:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 51:PlaySoundAt SoundFX("fx_bumper",DOFContactors), Bumper3:End Sub

' Drain holes, vuks & saucers
Sub Drain_Hit:PlaySoundAt "fx_drain",Drain:bsTrough.AddBall Me:End Sub
Sub sw9_Hit:PlaySoundAt "fx_kicker_enter",sw9:bsL.AddBall 0:End Sub
Sub sw30_Hit:PlaySoundAt "fx_kicker_enter",sw30:bsR.AddBall 0:End Sub
Sub sw41a_Hit:PlaySoundAt "fx_hole-enter",sw41a:bsTL.AddBall Me:End Sub
Sub sw41b_Hit:PlaySoundAt "fx_hole-enter",sw41b:bsTL.AddBall Me:End Sub
Sub sw41c_Hit:PlaySoundAt "fx_hole-enter",sw41c:bsTL.AddBall Me:End Sub
Sub sw41d_Hit:PlaySoundAt "fx_hole-enter",sw41d:bsTL.AddBall Me:End Sub
Sub sw46_Hit:PlaySoundAt "fx_kicker_enter",sw46:bsTR.AddBall 0:End Sub

' Rollovers & Ramp Switches
Sub sw57_Hit:Controller.Switch(57) = 1:PlaySoundAt "fx_sensor", sw57:End Sub
Sub sw57_UnHit:Controller.Switch(57) = 0:End Sub

  Sub sw58_Hit:Controller.Switch(58) = 1:PlaySoundAt "fx_sensor", sw58:End Sub
Sub sw58_UnHit:Controller.Switch(58) = 0:End Sub

Sub sw60_Hit:Controller.Switch(60) = 1:PlaySoundAt "fx_sensor", sw60:End Sub
Sub sw60_Unhit:Controller.Switch(60) = 0:End Sub

Sub sw61_Hit:Controller.Switch(61) = 1:PlaySoundAt "fx_sensor", sw61:End Sub
Sub sw61_Unhit:Controller.Switch(61) = 0:End Sub

Sub sw20_Hit:Controller.Switch(20) = 1:PlaySoundAt "fx_sensor", sw20:End Sub
Sub sw20_UnHit:Controller.Switch(20) = 0:End Sub

Sub sw21_Hit:Controller.Switch(21) = 1:PlaySoundAt "fx_sensor", sw21:End Sub
Sub sw21_Unhit:Controller.Switch(21) = 0:End Sub

Sub sw43_Hit:Controller.Switch(43) = 1:PlaySoundAt "fx_sensor", sw43:End Sub
Sub sw43_Unhit:Controller.Switch(43) = 0:End Sub

Sub sw44_Hit:Controller.Switch(44) = 1:PlaySoundAt "fx_sensor", sw44:End Sub
Sub sw44_Unhit:Controller.Switch(44) = 0:End Sub

Sub sw45_Hit:Controller.Switch(45) = 1:PlaySoundAt "fx_sensor", sw45:End Sub
Sub sw45_Unhit:Controller.Switch(45) = 0:End Sub

Sub sw37_Hit:Controller.Switch(37) = 1:PlaySoundAt "fx_sensor", sw37:End Sub
Sub sw37_Unhit:Controller.Switch(37) = 0:End Sub

Sub sw38_Hit:Controller.Switch(38) = 1:PlaySoundAt "fx_sensor", sw38:End Sub
Sub sw38_Unhit:Controller.Switch(38) = 0:End Sub

Sub sw39_Hit:Controller.Switch(39) = 1:PlaySoundAt "fx_sensor", sw39:End Sub
Sub sw39_Unhit:Controller.Switch(39) = 0:End Sub

Sub sw33_Hit:Controller.Switch(33) = 1:PlaySoundAt "fx_sensor", sw33:End Sub
Sub sw33_Unhit:Controller.Switch(33) = 0:End Sub

Sub sw34_Hit:Controller.Switch(34) = 1:PlaySoundAt "fx_sensor", sw34:End Sub
Sub sw34_Unhit:Controller.Switch(34) = 0:End Sub

Sub sw35_Hit:Controller.Switch(35) = 1:PlaySoundAt "fx_sensor", sw35:End Sub
Sub sw35_Unhit:Controller.Switch(35) = 0:End Sub

Sub sw36_Hit:Controller.Switch(36) = 1:PlaySoundAt "fx_sensor", sw36:End Sub
Sub sw36_Unhit:Controller.Switch(36) = 0:End Sub

Sub sw40_Hit:Controller.Switch(40) = 1:PlaySoundAt "fx_sensor", sw40:End Sub
Sub sw40_Unhit:Controller.Switch(40) = 0:PlaySound "fx_metalrolling", 0, 1, -0.15, 0.25:End Sub

Sub sw42_Hit:Controller.Switch(42) = 1:PlaySoundAt "fx_sensor", sw42:End Sub
Sub sw42_Unhit:Controller.Switch(42) = 0:End Sub

Sub sw24_Hit:Controller.Switch(24) = 1:PlaySoundAt "fx_sensor", sw24:End Sub
Sub sw24_Unhit:Controller.Switch(24) = 0:End Sub

Sub sw22_Hit:Controller.Switch(22) = 1:PlaySoundAt "fx_sensor", sw22:End Sub
Sub sw22_Unhit:Controller.Switch(22) = 0:StopSound "fx_metalrolling":PlaySound "fx_metalrolling", 0, 1, 0.15, 0.25:End Sub

Sub sw48_Hit:Controller.Switch(48) = 1:PlaySoundAt "fx_sensor", sw48:PlaySound "fx_metallrolling", 0, 1, pan(ActiveBall):End Sub
Sub sw48_Unhit:Controller.Switch(48) = 0:ActiveBall.VelX = 3:End Sub

Sub sw16_Hit:Controller.Switch(16) = 1:PlaySoundAt "fx_sensor", sw16:End Sub
Sub sw16_Unhit:Controller.Switch(16) = 0:End Sub

Sub sw25_Hit:Controller.Switch(25) = 1:End Sub
Sub sw25_Unhit:Controller.Switch(25) = 0:End Sub

Sub sw47_Hit:Controller.Switch(47) = 1:End Sub
Sub sw47_Unhit:Controller.Switch(47) = 0:End Sub

' Targets
Sub sw10_Hit:vpmTimer.PulseSw 10:PlaySound SoundFX("fx_target",DOFTargets), 0, 1, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub sw23_Hit:vpmTimer.PulseSw 23:PlaySound SoundFX("fx_target",DOFTargets), 0, 1, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub sw28_Collide(parm):vpmTimer.PulseSw 28:PlaySoundAt SoundFX("fx_target",DOFTargets),sw28:End Sub
Sub sw29_Hit:vpmTimer.PulseSw 29:PlaySound SoundFX("fx_target",DOFTargets), 0, 1, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub sw53_Hit:vpmTimer.PulseSw 53:PlaySound SoundFX("fx_target",DOFTargets), 0, 1, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub sw52_Spin():vpmTimer.PulseSw 52:PlaySoundAt "fx_spinner",sw52:End Sub

' lock post

Sub lock1_Hit:lock.Isdropped = 1:End Sub
Sub lock1_UnHit:lock.Isdropped = 0:End Sub

'*********
'Solenoids
'*********

SolCallBack(1) = "SolRelease"
SolCallBack(2) = "Auto_Plunger"
SolCallback(3) = "bsL.SolOut"
SolCallback(4) = "bsTL.SolOut"
SolCallback(5) = "bsR.SolOut"
' SolCallback(6) = "SolRingMagnet"
SolCallback(7) = "SolTower"
SolCallback(8) = "SolDiv"
' 9 left bumper
'10 right bumper
'11 bottom bumper
'12 not used
SolCallback(13) = "SolOrbit"
SolCallback(14) = "SetLamp 114,"
'15 left flipper
'16 right flipper
'17 left slingshot
'18 right slingshot
SolCallBack(19) = "bsTR.SolOut"
'20 balrog motor relay
SolCallBack(21) = "SolLockRelease"
SolCallBack(22) = "SolBalrog"
SolCallBack(23) = "SetLamp 123,"
SolCallBack(24) = "vpmsolsound SoundFX(""fx_knocker"",DOFKnocker),"
SolCallBack(25) = "SetLamp 125,"
SolCallBack(26) = "SetLamp 126,"
SolCallBack(27) = "SetLamp 127,"
'28 not used
SolCallBack(29) = "SetLamp 129,"
SolCallBack(30) = "SetLamp 130,"
SolCallBack(31) = "SetLamp 131,"
SolCallBack(32) = "SetLamp 132,"

Sub Auto_Plunger(Enabled)
    If Enabled Then
        PlungerIM.AutoFire
    End If
End Sub

Sub SolLockRelease(enabled)
    vlLock.SolExit enabled
End Sub

Sub SolRelease(Enabled)
    If Enabled And bsTrough.Balls > 0 Then
        bsTrough.ExitSol_On
        vpmTimer.PulseSw 15
    End If
End Sub

Sub SolOrbit(Enabled):OrbitPin.IsDropped = NOT Enabled:End Sub

'********************
'     Flippers
'********************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("fx_flipperup",DOFFlippers), LeftFlipper
        LeftFlipper.RotateToEnd
    Else
        PlaySoundAt SoundFX("fx_flipperdown",DOFFlippers), LeftFlipper
        LeftFlipper.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("fx_flipperup",DOFFlippers), RightFlipper
        RightFlipper.RotateToEnd
    Else
        PlaySoundAt SoundFX("fx_flipperdown",DOFFlippers), RightFlipper
        RightFlipper.RotateToStart
    End If
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, -0.1, 0.25
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, 0.1, 0.25
End Sub

'************************
'    Tower animation
'************************

Sub SolTower(Enabled)
    If Enabled Then
        TowerFlipper.RotateToEnd
    Else
        TowerFlipper.RotateToStart
    End If
End Sub

'************************
'   Diverter animation
'************************

Sub SolDiv(Enabled)
    If Enabled Then
        DiverterFlipper.RotateToEnd
    Else
        DiverterFlipper.RotateToStart
    End If
End Sub

'************************
'    Balrog Animation
'************************

Dim BalrogDir
BalrogDir = 0

Sub SolBalrog(Enabled)
    If Enabled Then
        If BalrogDir = 0 Then
            Controller.Switch(32) = 0
            Controller.Switch(31) = 1
            sw28.RotateToStart
            BalrogDir = 1
        Else
            Controller.Switch(31) = 0
            Controller.Switch(32) = 1
            sw28.RotateToEnd
            BalrogDir = 0
        End If
    End If
End Sub

'***************************************************
'       JP's VP10 Fading Lamps & Flashers
'       Based on PD's Fading Light System
' SetLamp 0 is Off
' SetLamp 1 is On
' fading for non opacity objects is 4 steps
'***************************************************

Dim LampState(200), FadingLevel(200)
Dim FlashSpeedUp(200), FlashSpeedDown(200), FlashMin(200), FlashMax(200), FlashLevel(200), FlashRepeat(200)

InitLamps()             ' turn off the lights and flashers and reset them to the default parameters
LampTimer.Interval = 10 'lamp fading speed
LampTimer.Enabled = 1

' Lamp & Flasher Timers

Sub LampTimer_Timer()
    Dim chgLamp, num, chg, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp)Then
        For ii = 0 To UBound(chgLamp)
            LampState(chgLamp(ii, 0)) = chgLamp(ii, 1)       'keep the real state in an array
            FadingLevel(chgLamp(ii, 0)) = chgLamp(ii, 1) + 4 'actual fading step
        Next
    End If
    UpdateLamps
End Sub

Sub UpdateLamps
    NFadeLm 1, l1
    Flashm 1, l1b
    Flash 1, l1a
    NFadeL 2, l2
    NFadeL 3, l3
    NFadeL 4, l4
    NFadeL 5, l5
    NFadeL 6, l6
    NFadeL 7, l7
    NFadeL 8, l8
    NFadeL 9, l9
    NFadeL 10, l10
    NFadeL 11, l11
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
    NFadeLm 23, l23
    FadeObj 23, globe, "globe_on", "globe_a", "globe_b", "globe"

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
    NFadeL 37, l37
    NFadeL 38, l38
    NFadeL 39, l39
    NFadeL 40, l40
    NFadeL 41, l41
    NFadeL 42, l42
    NFadeL 43, l43
    NFadeL 44, l44
    NFadeL 45, l45
    NFadeL 46, l46
    NFadeL 47, l47
    NFadeL 48, l48
    NFadeL 49, l49
    NFadeL 50, l50
    NFadeL 51, l51
    NFadeL 52, l52
    NFadeL 53, l53
    NFadeL 54, l54
    NFadeL 55, l55
    NFadeL 56, l56
    NFadeL 57, l57
    NFadeL 58, l58
    NFadeL 59, l59
    Flashm 60, l60a
    Flash 60, l60
    Flash 61, l61
    Flashm 62, l62a
    Flash 62, l62
    Flash 63, l63
    NFadeLm 64, l64
    Flash 64, l64a
    NFadeLm 65, l65
    Flash 65, l65a
    NFadeLm 66, l66
    Flash 66, l66a
    NFadeLm 67, l67
    Flash 67, l67a
    NFadeLm 68, l68
    Flash 68, l68a
    NFadeLm 69, l69
    Flash 69, l69a
    NFadeLm 70, l70
    Flash 70, l70a
    NFadeLm 71, l71
    Flash 71, l71a
    NFadeLm 72, l72
    Flash 72, l72a
    Flash 73, l73
    Flash 74, l74
    Flash 75, l75
    Flash 76, l76
    Flash 77, l77
    Flash 78, l78
    'NFadeL 79, l79
    'NFadeL 80, l80
    NFadeL 81, l81
    NFadeL 82, l82
    NFadeL 83, l83
    NFadeL 84, l84
    NFadeL 85, l85
    NFadeL 86, l86
    NFadeL 87, l87
    NFadeL 88, l88
    NFadeL 89, l89
    NFadeL 90, l90
    NFadeL 91, l91
    NFadeL 92, l92
    NFadeL 93, l93
    NFadeL 94, l94
    NFadeL 95, l95
    NFadeL 96, l96
    NFadeL 97, l97
    NFadeL 98, l98
    NFadeL 99, l99

    'flashers
    Flash 114, f14
    Flash 123, f23
    Flashm 125, f25a
    Flashm 125, f25b
    Flash 125, f25c
    Flash 126, f26
    Flashm 127, f27a
    Flash 127, f27b
    Flashm 129, f29a
    Flash 129, f29b
    NFadeL 130, f30
    Flash 131, f31
    'NFadeL 132, f32
FadeObj 132, balrog, "balrog_on", "balrog_a", "balrog_b", "balrog"
End Sub

' div lamp subs

Sub InitLamps()
    Dim x
    For x = 0 to 200
        LampState(x) = 0        ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4      ' used to track the fading state
        FlashSpeedUp(x) = 0.2   ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.1 ' slower speed when turning off the flasher
        FlashMax(x) = 1         ' the maximum value when on, usually 1
        FlashMin(x) = 0         ' the minimum value when off, usually 0
        FlashLevel(x) = 0       ' the intensity of the flashers, usually from 0 to 1
        FlashRepeat(x) = 20     ' how many times the flash repeats
    Next
End Sub

Sub AllLampsOff
    Dim x
    For x = 0 to 200
        SetLamp x, 0
    Next
End Sub

Sub SetLamp(nr, value)
    If value <> LampState(nr)Then
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
            FlashLevel(nr) = FlashLevel(nr)- FlashSpeedDown(nr)
            If FlashLevel(nr) < FlashMin(nr)Then
                FlashLevel(nr) = FlashMin(nr)
                FadingLevel(nr) = 0 'completely off
            End if
            Object.IntensityScale = FlashLevel(nr)
        Case 5 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
            If FlashLevel(nr) > FlashMax(nr)Then
                FlashLevel(nr) = FlashMax(nr)
                FadingLevel(nr) = 1 'completely on
            End if
            Object.IntensityScale = FlashLevel(nr)
    End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it doesn't change anything, it just follows the main flasher
    Select Case FadingLevel(nr)
        Case 4, 5
            Object.IntensityScale = FlashLevel(nr)
    End Select
End Sub

Sub FlashBlink(nr, object)
    Select Case FadingLevel(nr)
        Case 4 'off
            FlashLevel(nr) = FlashLevel(nr)- FlashSpeedDown(nr)
            If FlashLevel(nr) < FlashMin(nr)Then
                FlashLevel(nr) = FlashMin(nr)
                FadingLevel(nr) = 0 'completely off
            End if
            Object.IntensityScale = FlashLevel(nr)
            If FadingLevel(nr) = 0 AND FlashRepeat(nr)Then 'repeat the flash
                FlashRepeat(nr) = FlashRepeat(nr)-1
                If FlashRepeat(nr)Then FadingLevel(nr) = 5
            End If
        Case 5 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
            If FlashLevel(nr) > FlashMax(nr)Then
                FlashLevel(nr) = FlashMax(nr)
                FadingLevel(nr) = 1 'completely on
            End if
            Object.IntensityScale = FlashLevel(nr)
            If FadingLevel(nr) = 1 AND FlashRepeat(nr)Then FadingLevel(nr) = 4
    End Select
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

'******************************
' Diverse Collection Hit Sounds
'******************************

Sub aMetal_Wires_Hit(idx):PlaySound "fx_MetalHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aMetal_Lanes_Hit(idx):PlaySound "fx_MetalHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aMetal_Guides_Hit(idx):PlaySound "fx_MetalHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Bands_Hit(idx):PlaySound "fx_rubber_band", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Posts_Hit(idx):PlaySound "fx_rubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Pins_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Thin_Posts_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aWoods_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub

Sub sw_Hit:PlaySound "fx_metalrolling", 0, 1, 0.15, 0.25:End Sub

' Ramp Soundss & Help triggers
Sub REnd1_Hit()
    StopSound "fx_metalrolling"
    PlaySound "fx_ballrampdrop", 0, 1, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub REnd2_Hit()
    StopSound "fx_metalrolling"
    PlaySound "fx_ballrampdrop", 0, 1, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

'******************
' RealTime Updates
'******************

Set MotorCallback = GetRef("RealTimeUpdates")

Sub RealTimeUpdates
    RollingUpdate
    Diverter.RotZ = DiverterFlipper.CurrentAngle
    Balrog.RotZ = sw28.CurrentAngle
    Tower.RotX = TowerFlipper.CurrentAngle
End Sub

'*************************
' GI - needs new vpinmame
'*************************

Set GICallback = GetRef("GIUpdate")

Sub GIUpdate(no, Enabled)
    For each x in aGiLights
        x.State = ABS(Enabled)
    Next
    For each x in aGiFlashers
        x.visible = Enabled
    Next
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
  Vol = Csng(BallVel(ball) ^2 / VolDiv)
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

Const tnob = 5 ' total number of balls in this table is 4, but always use a higher number here because of the timing
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
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / VolDiv, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
  Else
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / VolDiv, Pan(ball1), 0, Pitch(ball1), 0, 0
  End if
End Sub


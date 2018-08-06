' Ripley's Believe or not - Stern 2004
' vpx 1.0 by JPSalas August 2017

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' No special SSF tweaks yet.
' This is a JP table. He often uses walls as switches so I need to be careful of using PlaySoundAt

Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const BallSize = 50

' Language Roms

Const cGameName = "ripleys" 'English
' Const cGameName = "ripleysf" 'French
' Const cGameName = "ripleysg" 'German
' Const cGameName = "ripleysi" 'Italian
' Const cGameName = "ripleysl" 'Spanish

Dim VarHidden, UseVPMColoredDMD
If Table1.ShowDT = true then
    UseVPMColoredDMD = true
    VarHidden = 1
else
    UseVPMColoredDMD = False
    VarHidden = 0
    TextBox1.Visible = 0
    lrail.Visible = 0
    rrail.Visible = 0
end if

LoadVPM "01550000", "SEGA.VBS", 3.26

'********************
'Standard definitions
'********************

Const UseSolenoids = 2
Const UseLamps = 0
Const UseGI = 0
Const UseSync = 0
Const HandleMech = 0

' Standard Sounds
Const SSolenoidOn = "fx_Solenoidon"
Const SSolenoidOff = "fx_SolenoidOff"
Const SCoin = "fx_coin"

Dim bsTrough, mIdolMag, mShrunkenMag, mUp, mDown, bsLock, bsSkill, bsVUK, plungerIM, x

'************
' Table init.
'************

Sub table1_Init
    vpmInit Me
    With Controller
        .GameName = cGameName
        .Games(cGameName).Settings.Value("sound") = 1 '0 = disable  rom sound
        .SplashInfoLine = "Ripleys Believe It Or Not - Stern 2004" & vbNewLine & "VPX table by JPSalas v.1.0"
        .Games(cGameName).Settings.Value("rol") = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .HandleKeyboard = 0
        .Hidden = VarHidden
        .Switch(42) = 1
        .Switch(43) = 1
        On Error Resume Next
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description

        On Error Goto 0
    End With

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

    vpmNudge.TiltSwitch = 56
    vpmNudge.Sensitivity = 1
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, Bumper4, Bumper5, Bumper6, LeftSlingshot, RightSlingshot)

    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0, 14, 13, 12, 11, 0, 0, 0
        .InitKick BallRelease, 100, 3
        .InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .Balls = 4
    End With

    Set mIdolMag = New cvpmMagnet
    With mIdolMag
        .InitMagnet IMagnet, 50
        .Solenoid = 19
        .GrabCenter = 0
    End With

    Set mShrunkenMag = New cvpmMagnet
    With mShrunkenMag
        .InitMagnet SMagnet, 70
        .GrabCenter = 1
    End With

    Set bsLock = New cvpmBallStack
    With bsLock
        .InitSw 0, 46, 45, 44, 0, 0, 0, 0
        .InitKick Lock, 0, 36
        .KickForceVar = 4
        .InitExitSnd SoundFX("fx_popper", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    End With

    Set bsSkill = new cvpmBallStack
    With bsSkill
        .InitSw 0, 29, 0, 0, 0, 0, 0, 0
        .InitKick sw29, 180, 19
        .InitExitSnd "fx_popper", "fx_Solenoid"
    End With

    Set bsVUK = new cvpmBallStack
    With bsVUK
        .InitSw 0, 52, 0, 0, 0, 0, 0, 0
        .InitKick sw52, 0, 36
        .InitExitSnd SoundFX("fx_popper", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .KickZ = 1.5
    End With

    ' Impulse Plunger
    Const IMPowerSetting = 44 ' Plunger Power
    Const IMTime = 0.7        ' Time in seconds for Full Plunge
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swPlunger, IMPowerSetting, IMTime
        .Random 0.3
        .switch 16
        .InitExitSnd "fx_plunger2", "fx_plunger"
        .CreateEvents "plungerIM"
    End With
End Sub

Sub Table1_Paused:Controller.Pause = 1:End Sub
Sub Table1_unPaused:Controller.Pause = 0:End Sub
Sub Table1_Exit:Controller.Stop:End Sub

'****
'Keys
'****

Sub Table1_KeyDown(ByVal keycode)
    If keycode = LeftTiltKey Then Nudge 90, 8:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 8:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 9:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If keycode = PlungerKey Then PlaySound "fx_PlungerPull", 0, 1, 0.1, 0.25:Plunger.Pullback
    If KeyDownHandler(KeyCode)Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal keycode)
    If KeyUpHandler(KeyCode)Then Exit Sub
    If keycode = PlungerKey Then PlaySound "fx_plunger", 0, 1, 0.1, 0.25:Plunger.Fire
End Sub

'*********
'Solenoids
'*********

SolCallBack(1) = "SolTrough"
SolCallBack(2) = "Auto_Plunger"
SolCallBack(3) = "bsVUK.SolOut"
SolCallBack(4) = "vpmSolDiverter TempleDiv,1,"
SolCallBack(5) = "vpmSolDiverter LockDiverter,1,"
' 6, 7, 8, 9, 10 11 pop bumpers
SolCallBack(12) = "bsSkill.SolOut"
SolCallback(13) = "bsLock.SolOut"
SolCallback(20) = "SolUpperMagnet"
SolCallBack(21) = "SolVReset"
SolCallBack(22) = "SetLamp 122,"
SolCallBack(23) = "SolPost"
SolCallBack(24) = "vpmsolsound SoundFX(""fx_knocker"",DOFKnocker),"
SolCallBack(25) = "SetLamp 125,"
SolCallBack(26) = "SetLamp 126,"
SolCallBack(27) = "SetLamp 127,"
SolCallBack(28) = "SetLamp 128,"
SolCallBack(29) = "SetLamp 129,"
SolCallBack(30) = "SetLamp 130,"
SolCallBack(31) = "SetLamp 131,"
SolCallBack(32) = "SetLamp 132,"

'********************
'     Flippers
'********************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySound SoundFX("fx_flipperup", DOFFlippers), 0, 1, -0.1, 0.25
        LeftFlipper.RotateToEnd
    Else
        PlaySound SoundFX("fx_flipperdown", DOFFlippers), 0, 1, -0.1, 0.25
        LeftFlipper.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySound SoundFX("fx_flipperup", DOFFlippers), 0, 1, 0.1, 0.25
        RightFlipper.RotateToEnd
        RightFlipper1.RotateToEnd
    Else
        PlaySound SoundFX("fx_flipperdown", DOFFlippers), 0, 1, 0.1, 0.25
        RightFlipper.RotateToStart
        RightFlipper1.RotateToStart
    End If
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, -0.1, 0.25
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, 0.1, 0.25
End Sub

Sub RightFlipper1_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, 0.1, 0.25
End Sub

'**************
' Solenoid Subs
'**************

Sub Auto_Plunger(Enabled)
    If Enabled Then
        PlungerIM.AutoFire
    End If
End Sub

Sub SolVReset(Enabled)
    If Enabled Then
        VariTimerDown.Enabled = 1
    End If
End Sub

Sub SolPost(Enabled)
    TopPost.IsDropped = NOT Enabled
End Sub

Sub SolTrough(Enabled)
    If Enabled Then
        bsTrough.ExitSol_On
        vpmTimer.PulseSw 15
    End If
End Sub

'**********************
' Shrunken Head Magnet
'**********************

Sub SMagnet_Hit
    Controller.Switch(23) = 1
    mShrunkenMag.AddBall ActiveBall
End Sub

Sub SMagnet_unHit
	Controller.Switch(23) = 0
	mShrunkenMag.RemoveBall ActiveBall
End Sub

Sub SolUpperMagnet(Enabled)
	Dim ball
    If Enabled Then
        mShrunkenMag.MagnetOn = 1
    Else
        mShrunkenMag.MagnetOn = 0
		For Each ball in mShrunkenMag.Balls
			ball.VelY = -19
		Next
    End If
End Sub

' ************************************
' Switches, bumpers, lanes and targets
' ************************************

Sub Drain_Hit::PlaySound "fx_drain":bsTrough.AddBall Me:End Sub

Sub sw9_Hit:vpmTimer.PulseSw 9:PlaySound SoundFX("fx_target", DOFTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw17_Hit:vpmTimer.PulseSw 17:PlaySound SoundFX("fx_target", DOFTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw18_Hit:Controller.Switch(18) = 1:ActiveBall.VelX = - 1:ActiveBall.VelY = 1:PlaySound SoundFX("fx_balldrop", DOFTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw18_unHit:Controller.Switch(18) = 0:End Sub
Sub sw19_Hit:vpmTimer.PulseSw 19:PlaySound SoundFX("fx_target", DOFTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw20_Spin:PlaySound "fx_spinner", 0, 1, -0.01:vpmTimer.PulseSw 20:End Sub
Sub sw21_Spin:PlaySound "fx_spinner", 0, 1, 0.01:vpmTimer.PulseSw 21:End Sub
Sub sw22_Hit:vpmTimer.PulseSw 22:PlaySound SoundFX("fx_target", DOFTargets), 0, 1, pan(ActiveBall):End Sub

Sub IMagnet_Hit:mIdolMag.AddBall ActiveBall:Controller.Switch(24) = 1:End Sub
Sub IMagnet_unHit:mIdolMag.RemoveBall ActiveBall:Controller.Switch(24) = 0:End Sub

'left bumpers
Sub Bumper4_Hit:vpmTimer.PulseSw 25:PlaySound SoundFX("fx_bumper", DOFContactors), 0, 1, 0, 0.15:End Sub
Sub Bumper5_Hit:vpmTimer.PulseSw 26:PlaySound SoundFX("fx_bumper", DOFContactors), 0, 1, 0.15, 0.15:End Sub
Sub Bumper6_Hit:vpmTimer.PulseSw 27:PlaySound SoundFX("fx_bumper", DOFContactors), 0, 1, 0.05, 0.15:End Sub

Sub sw28_Hit:sw28.DestroyBall:PlaySound "fx_hole-enter", 0, 1, -0.005:vpmTimer.PulseSwitch(28), 100, "AddToSkill":End Sub
Sub AddToSkill(swNo):bsSkill.AddBall 0:End Sub

Dim aBall, aZpos

Sub sw29_Hit
    Set aBall = ActiveBall
    PlaySound "fx_hole-enter", 0, 1, -0.005
    aZpos = 35
    Me.TimerInterval = 4
    Me.TimerEnabled = 1
End Sub

Sub sw29_Timer
    aBall.Z = aZpos
    aZpos = aZpos-4
    If aZpos < -30 Then
        Me.TimerEnabled = 0
        Me.DestroyBall
        bsSkill.AddBall 0
    End If
End Sub

Sub sw30_Hit:Controller.Switch(30) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw30_unHit:Controller.Switch(30) = 0:End Sub
Sub sw31_Hit:Controller.Switch(31) = 1:End Sub
Sub sw31_unHit:Controller.Switch(31) = 0:End Sub
Sub sw32_Hit:vpmTimer.PulseSw 32:PlaySound SoundFX("fx_target", DOFTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw32b_Hit:vpmTimer.PulseSw 32:PlaySound SoundFX("fx_target", DOFTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw33_Hit:Controller.Switch(33) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw33_unHit:Controller.Switch(33) = 0:End Sub
Sub sw34_Hit:Controller.Switch(34) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw34_unHit:Controller.Switch(34) = 0:End Sub
Sub sw35_Hit:Controller.Switch(35) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw35_unHit:Controller.Switch(35) = 0:End Sub
Sub sw36_Hit:Controller.Switch(36) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw36_unHit:Controller.Switch(36) = 0:End Sub
Sub sw38_Hit:Controller.Switch(38) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw38_unHit:Controller.Switch(38) = 0:PlaySound "fx_metalrolling", 0, 1, pan(ActiveBall):End Sub
Sub sw39_Hit:Controller.Switch(39) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw39_unHit:Controller.Switch(39) = 0:PlaySound "fx_metalrolling", 0, 1, pan(ActiveBall):End Sub
Sub sw40_Hit:Controller.Switch(40) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw40_unHit:Controller.Switch(40) = 0:PlaySound "fx_metalrolling", 0, 1, pan(ActiveBall):End Sub
Sub Lock_Hit:bsLock.AddBall Me:PlaySound "fx_kicker_enter", 0, 1, 0.01:End Sub
Sub sw47_Hit:Controller.Switch(47) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw47_unHit:Controller.Switch(47) = 0:End Sub
Sub sw48_Hit:Controller.Switch(48) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw48_unHit:Controller.Switch(48) = 0:End Sub

Sub Bumper1_Hit:vpmTimer.PulseSw 49:PlaySound SoundFX("fx_bumper", DOFContactors), 0, 1, 0, 0.15:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 50:PlaySound SoundFX("fx_bumper", DOFContactors), 0, 1, 0.15, 0.15:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 51:PlaySound SoundFX("fx_bumper", DOFContactors), 0, 1, 0.05, 0.15:End Sub

'************
' Varitarget
'************

Sub aVariPos_Hit(idx)
    If ActiveBall.VelY < 0 Then 'ball moves up
        VariTarget.RotY = 14 - idx
    End If
End Sub

Sub sw42_Hit
    If ActiveBall.VelY < 0 Then
        Controller.Switch(42) = 1
        ActiveBall.VelY = ActiveBall.VelY * 0.8
    End If
End Sub

Sub sw43_Hit
    If ActiveBall.VelY < 0 Then
        Controller.Switch(43) = 1
        ActiveBall.VelY = ActiveBall.VelY * 0.8
    End If
End Sub

Sub sw41_Hit
    If ActiveBall.VelY < 0 Then
        Controller.Switch(41) = 1
        ActiveBall.VelY = ActiveBall.VelY * 0.8
    End If
End Sub

Sub VariTimerDown_Timer
    VariTarget.RotY = VariTarget.RotY + 1
    If VariTarget.RotY = -5 Then Controller.Switch(41) = 0
    If VariTarget.RotY = 5 Then Controller.Switch(43) = 0
    If VariTarget.RotY > 15 Then
        VariTarget.RotY = 15
        Controller.Switch(42) = 0
        VariTimerDown.Enabled = 0
    End If
End Sub

'******
' vuk
'******

Sub sw52_Hit():bsVUK.AddBall Me:PlaySound "fx_kicker_enter", 0, 1, 0.01:End Sub

'*******
' lanes
'*******

Sub sw53_Hit:Controller.Switch(53) = 1:End Sub
Sub sw53_unHit:Controller.Switch(53) = 0:End Sub

Sub sw57_Hit:Controller.Switch(57) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw57_UnHit:Controller.Switch(57) = 0:End Sub

Sub sw58_Hit:Controller.Switch(58) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw58_UnHit:Controller.Switch(58) = 0:End Sub

Sub sw61_Hit:Controller.Switch(61) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw61_UnHit:Controller.Switch(61) = 0:End Sub

Sub sw60_Hit:Controller.Switch(60) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw60_UnHit:Controller.Switch(60) = 0:End Sub

'************
' Slingshots
'************

Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    PlaySound SoundFX("fx_slingshot", DOFContactors), 0, 1, -0.05, 0.05
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
    PlaySound SoundFX("fx_slingshot", DOFContactors), 0, 1, 0.05, 0.05
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

'**************
' LED's display
'**************

Dim LED(14)

LED(0) = Array(LED1, LED2, LED3, LED4, LED5, LED6, LED7)
LED(1) = Array(LED8, LED9, LED10, LED11, LED12, LED13, LED14)
LED(2) = Array(LED15, LED16, LED17, LED18, LED19, LED20, LED21)
LED(3) = Array(LED22, LED23, LED24, LED25, LED26, LED27, LED28)
LED(4) = Array(LED29, LED30, LED31, LED32, LED33, LED34, LED35)

LED(5) = Array(LED36, LED37, LED38, LED39, LED40, LED41, LED42)
LED(6) = Array(LED43, LED44, LED45, LED46, LED47, LED48, LED49)
LED(7) = Array(LED50, LED51, LED52, LED53, LED54, LED55, LED56)
LED(8) = Array(LED57, LED58, LED59, LED60, LED61, LED62, LED63)
LED(9) = Array(LED64, LED65, LED66, LED67, LED68, LED69, LED70)

LED(10) = Array(LED71, LED72, LED73, LED74, LED75, LED76, LED77)
LED(11) = Array(LED78, LED79, LED80, LED81, LED82, LED83, LED84)
LED(12) = Array(LED85, LED86, LED87, LED88, LED89, LED90, LED91)
LED(13) = Array(LED92, LED93, LED94, LED95, LED96, LED97, LED98)
LED(14) = Array(LED99, LED100, LED101, LED102, LED103, LED104, LED105)

Sub UpdateLeds()
    Dim ChgLED, ii, num, chg, stat, obj
    ChgLED = Controller.ChangedLEDs(&H00000000, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
        For ii = 0 To UBound(chgLED)
            num = chgLED(ii, 0):chg = chgLED(ii, 1):stat = chgLED(ii, 2)
            For Each obj In LED(num)
                If chg And 1 Then obj.Visible = stat And 1
                chg = chg \ 4:stat = stat \ 4
            Next
        Next
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
    UpdateLeds
    UpdateLamps
    RollingUpdate
End Sub

Sub UpdateLamps
    NFadeL 1, l1
    NFadeL 2, l2
    NFadeL 3, l3
    NFadeL 4, l4
    NFadeL 5, l5
    NFadeL 6, l6
    NFadeL 7, l7
    NFadeLm 8, l8
    Flash 8, l8a
    NFadeLm 9, l9
    Flash 9, l9a
    NFadeL 10, l10
    NFadeL 11, l11
    NFadeL 12, l12
    NFadeL 13, l13
    NFadeL 14, l14
    NFadeLm 15, l15
    Flash 15, l15a
    Flash 16, l16
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
    NFadeLm 33, BumperLight4a
    NFadeL 33, BumperLight4
    NFadeLm 34, BumperLight5a
    NFadeL 34, BumperLight5
    NFadeLm 35, BumperLight6a
    NFadeL 35, BumperLight6
    NFadeL 36, l36
    NFadeL 37, l37
    NFadeL 38, l38
    NFadeL 39, l39
    NFadeL 40, l40
    NFadeLm 41, l41
    Flash 41, l41a
    NFadeLm 42, l42
    Flash 42, l42a
    NFadeLm 43, l43
    Flash 43, l43a
    NFadeLm 44, l44
    Flash 44, l44a
    NFadeLm 45, l45
    Flash 45, l45a
    NFadeLm 46, l46
    Flash 46, l46a
    NFadeLm 47, l47
    Flash 47, l47a
    NFadeLm 48, l48
    Flash 48, l48a
    NFadeLm 49, l49
    Flash 49, l49a
    NFadeL 50, l50
    NFadeL 51, l51
    NFadeLm 52, l52
    Flash 52, l52a
    NFadeL 53, l53
    NFadeL 54, l54

    NFadeL 55, l55
    NFadeLm 56, l56
    Flash 56, l56a
    NFadeL 57, l57
    NFadeL 58, l58
    NFadeL 59, l59
    NFadeLm 60, BumperLight1a
    NFadeL 60, BumperLight1
    NFadeLm 61, BumperLight2a
    NFadeL 61, BumperLight2
    NFadeLm 62, BumperLight3a
    NFadeL 62, BumperLight3
    NFadeLm 63, l63
    Flash 63, l63a
    NFadeLm 64, l64
    Flash 64, l64a
    NFadeL 65, l65
    NFadeL 66, l66
    NFadeL 67, l67
    NFadeL 68, l68
    NFadeLm 69, l69
    Flash 69, l69a
    NFadeL 70, l70
    NFadeL 71, l71
    NFadeL 72, l72
    NFadeLm 73, l73
    Flash 73, l73a
    NFadeLm 74, l74
    Flash 74, l74a
    NFadeLm 75, l75
    Flash 75, l75a
    Flash 76, l76
    Flash 76, l76
    Flash 77, l77
    Flash 78, l78
    'NFadeL 80, l80

    'Flashers
    NFadeLm 122, f22
    Flash 122, f22a
    NFadeLm 125, f25
    Flashm 125, f25a
    Flash 125, f25b
    Flash 126, f26
    Flash 127, f27
    NFadeLm 128, f28a
    Flash 128, f28
    Flash 129, f29
    NFadeLm 130, f30
    Flashm 130, f30a
    Flashm 130, f30c
    Flash 130, f30b
    Flash 131, f31
    Flash 132, f32
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

Sub aMetals_Hit(idx):PlaySound "fx_MetalHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Bands_Hit(idx):PlaySound "fx_rubber_band", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Posts_Hit(idx):PlaySound "fx_rubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Pins_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aWoods_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub

' Ramp Soundss & Help triggers
Sub REnd1_Hit()
    StopSound "fx_metalrolling"
    PlaySound "fx_ballrampdrop", 0, 1, pan(ActiveBall)
End Sub

Sub REnd2_Hit()
    StopSound "fx_metalrolling"
    PlaySound "fx_ballrampdrop", 0, 1, pan(ActiveBall)
End Sub

Sub REnd3_Hit()
    PlaySound "fx_ballrampdrop", 0, 1, pan(ActiveBall)
End Sub

Sub REnd4_Hit()
    PlaySound "fx_ballrampdrop", 0, 1, pan(ActiveBall)
End Sub

Sub RSound1_Hit:PlaySound "fx_metalrolling", 0, 1, pan(ActiveBall):End Sub

'*****************************************
'			FLIPPER SHADOWS
'*****************************************

sub FlipperTimer_Timer()
	FlipperLSh.RotZ = LeftFlipper.currentangle
	FlipperRSh.RotZ = RightFlipper.currentangle

End Sub

'*****************************************
'			BALL SHADOW
'*****************************************
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
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) + 6
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) - 6
        End If
        ballShadow(b).Y = BOT(b).Y + 12
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
End Sub


'*************************
' GI - needs new vpinmame
'*************************

Set GICallback = GetRef("GIUpdate")

Sub GIUpdate(no, Enabled)
    For each x in aGiLights
        x.State = ABS(Enabled)
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
    Dim BOT, b, ballpitch
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


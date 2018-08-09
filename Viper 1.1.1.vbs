' Viper - Stern 1981
' IPD No. 2739 / October, 1981 / 4 Players
' VPX - version by JPSalas 2017, version 1.1.1

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' Thalamus 2018-08-09 : Improved directional sounds

Option Explicit
Randomize

' Options
' Volume devided by - lower gets higher sound

Const VolDiv = 500

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01550000", "stern.vbs", 3.26

Dim bsTrough, dtLeft, dtRight, x

Const cGameName = "viper"

Const UseSolenoids = 2
Const UseLamps = 0
Const UseGI = 0
Const UseSync = 0 'set it to 1 if the table runs too fast
Const HandleMech = 0

Dim VarHidden
If Table1.ShowDT = true then
    VarHidden = 1
    For each x in aReels
        x.Visible = 1
    Next
else
    VarHidden = 0
    For each x in aReels
        x.Visible = 0
    Next
    lrail.Visible = 0
    rrail.Visible = 0
end if

if B2SOn = true then VarHidden = 1

' Standard Sounds
Const SSolenoidOn = "fx_Solenoid"
Const SSolenoidOff = ""
Const SCoin = "fx_Coin"

'************
' Table init.
'************

Sub table1_Init
    vpmInit me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Viper - Stern 1981" & vbNewLine & "VPX table by JPSalas v.1.1.1"
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = VarHidden
        .Games(cGameName).Settings.Value("rol") = 0 '1= rotated display, 0= normal
        '.SetDisplayPosition 0,0, GetPlayerHWnd 'restore dmd window position
        On Error Resume Next
        Controller.SolMask(0) = 0
        vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the Timer to renable all the solenoids after 2 seconds
        Controller.Run GetPlayerHWnd
        On Error Goto 0
    End With

    ' Nudging
    vpmNudge.TiltSwitch = swTilt
    vpmNudge.Sensitivity = 3
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, LeftSlingshot, RightSlingshot)

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 8, 33, 34, 35, 0, 0, 0, 0
        .InitKick BallRelease, 90, 4
        .InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .IsTrough = True
        .Balls = 3
    End With

    ' Left Drop Target Bank
    Set dtLeft = New cvpmDropTarget
    With dtLeft
        .InitDrop Array(sw28, sw29, sw30), Array(28, 29, 30)
        .initsnd SoundFX("", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
    End With

    ' Right Drop Target Bank
    Set dtRight = New cvpmDropTarget
    With dtRight
        .InitDrop Array(sw27, sw26, sw25), Array(27, 26, 25)
        .initsnd SoundFX("", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
    End With

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

    ' Turn on Gi
    GiOn()
    WheelInit()
End Sub

Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub
Sub table1_exit:Controller.stop:End Sub

'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If keycode = PlungerKey Then PlaySound "fx_PlungerPull", 0, 1, 0.1, 0.25:Plunger.Pullback
    If KeyCode = LeftMagnaSave Then WheelDir = -2
    If KeyCode = RightMagnaSave Then Controller.Switch(4) = 1
    If vpmKeyDown(keycode)Then Exit Sub
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If KeyCode = LeftMagnaSave Then WheelDir = 2
    If KeyCode = RightMagnaSave Then Controller.Switch(4) = 0
    If vpmKeyUp(keycode)Then Exit Sub
    If keycode = PlungerKey Then PlaySoundAt "fx_plunger",Plunger:Plunger.Fire
End Sub

'*********
' Switches
'*********

' Slings
Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    PlaySound SoundFX("fx_slingshot", DOFContactors), 0, 1, -0.05, 0.05
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 16
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
    vpmTimer.PulseSw 15
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
Sub Bumper1_Hit:vpmTimer.PulseSw 14:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper1:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 13:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper2:End Sub

' Rollovers
Sub sw12_Hit:Controller.Switch(12) = 1:PlaySoundAt "fx_sensor", sw12:DOF 104, DOFOn:End Sub
Sub sw12_UnHit:Controller.Switch(12) = 0:DOF 104, DOFOff:End Sub

Sub sw36_Hit:Controller.Switch(36) = 1:PlaySoundAt "fx_sensor", sw36:DOF 101, DOFOn:End Sub
Sub sw36_UnHit:Controller.Switch(36) = 0:DOF 101, DOFOff:End Sub

Sub sw36a_Hit:Controller.Switch(36) = 1:PlaySoundAt "fx_sensor", sw36a:DOF 105, DOFOn:End Sub
Sub sw36a_UnHit:Controller.Switch(36) = 0:DOF 105, DOFOff:End Sub

Sub sw37_Hit:Controller.Switch(37) = 1:PlaySoundAt "fx_sensor", sw37:DOF 102, DOFOn:End Sub
Sub sw37_UnHit:Controller.Switch(37) = 0:DOF 102, DOFOff:End Sub

Sub sw37a_Hit:Controller.Switch(37) = 1:PlaySoundAt "fx_sensor", sw37a:DOF 108, DOFOn:End Sub
Sub sw37a_UnHit:Controller.Switch(37) = 0:DOF 108, DOFOff:End Sub

Sub sw9_Hit:Controller.Switch(9) = 1:PlaySoundAt "fx_sensor", sw9:DOF 103, DOFOn:End Sub
Sub sw9_UnHit:Controller.Switch(9) = 0:DOF 103, DOFOff:End Sub

Sub sw9a_Hit:Controller.Switch(9) = 1:PlaySoundAt "fx_sensor", sw9a:DOF 106, DOFOn:End Sub
Sub sw9a_UnHit:Controller.Switch(9) = 0:DOF 106, DOFOff:End Sub

Sub sw9b_Hit:Controller.Switch(9) = 1:PlaySoundAt "fx_sensor", sw9b:DOF 107, DOFOn:End Sub
Sub sw9b_UnHit:Controller.Switch(9) = 0:DOF 107, DOFOff:End Sub

Sub sw11_Hit:Controller.Switch(11) = 1:PlaySoundAt "fx_sensor", sw11:End Sub
Sub sw11_UnHit:Controller.Switch(11) = 0:End Sub

' Droptargets
Sub sw28_Hit:PlaySoundAt SoundFX("fx_droptarget", DOFDropTargets), ActiveBall:End Sub
Sub sw29_Hit:PlaySoundAt SoundFX("fx_droptarget", DOFDropTargets), ActiveBall:End Sub
Sub sw30_Hit:PlaySoundAt SoundFX("fx_droptarget", DOFDropTargets), ActiveBall:End Sub
Sub sw27_Hit:PlaySoundAt SoundFX("fx_droptarget", DOFDropTargets), ActiveBall:End Sub
Sub sw26_Hit:PlaySoundAt SoundFX("fx_droptarget", DOFDropTargets), ActiveBall:End Sub
Sub sw25_Hit:PlaySoundAt SoundFX("fx_droptarget", DOFDropTargets), ActiveBall:End Sub

Sub sw28_Dropped:dtLeft.Hit 1:End Sub
Sub sw29_Dropped:dtLeft.Hit 2:End Sub
Sub sw30_Dropped:dtLeft.Hit 3:End Sub
Sub sw27_Dropped:dtRight.Hit 1:End Sub
Sub sw26_Dropped:dtRight.Hit 2:End Sub
Sub sw25_Dropped:dtRight.Hit 3:End Sub

'Targets
Sub sw14_Hit:vpmTimer.PulseSw 14:PlaySoundAt SoundFX("fx_target", DOFDropTargets), ActiveBall:End Sub
Sub sw17_Hit:vpmTimer.PulseSw 17:PlaySoundAt SoundFX("fx_target", DOFDropTargets), ActiveBall:End Sub
Sub sw18_Hit:vpmTimer.PulseSw 18:PlaySoundAt SoundFX("fx_target", DOFDropTargets), ActiveBall:End Sub
Sub sw19_Hit:vpmTimer.PulseSw 19:PlaySoundAt SoundFX("fx_target", DOFDropTargets), ActiveBall:End Sub
Sub sw20_Hit:vpmTimer.PulseSw 20:PlaySoundAt SoundFX("fx_target", DOFDropTargets), ActiveBall:End Sub
Sub sw21_Hit:vpmTimer.PulseSw 21:PlaySoundAt SoundFX("fx_target", DOFDropTargets), ActiveBall:End Sub
Sub sw22_Hit:vpmTimer.PulseSw 22:PlaySoundAt SoundFX("fx_target", DOFDropTargets), ActiveBall:End Sub
Sub sw23_Hit:vpmTimer.PulseSw 23:PlaySoundAt SoundFX("fx_target", DOFDropTargets), ActiveBall:End Sub

'Spinner
Sub sw5_Spin:vpmTimer.PulseSw 5:Playsound "fx_spinner", sw5:End Sub

' Drain and kickers
Sub Drain_Hit:PlaysoundAt "fx_drain", Drain:bsTrough.AddBall Me:End Sub
Sub sw40_Hit:PlaySoundAt "fx_kicker_enter",sw40:Controller.Switch(40) = 1:End Sub

'Wheel Target Hits

Sub awt1_Hit(idx):vpmTimer.PulseSw 31:tw1.TransY = -113: t1w.TimerEnabled = 1:PlaySoundAt SoundFX("fx_target", DOFDropTargets), ActiveBall:End Sub
Sub t1w_Timer: tw1.TransY = -120: t1w.TimerEnabled = 0: End Sub

Sub awt2_Hit(idx):vpmTimer.PulseSw 32:tw2.TransY = -113: t2w.TimerEnabled = 1:PlaySoundAt SoundFX("fx_target", DOFDropTargets), ActiveBall:End Sub
Sub t2w_Timer: tw2.TransY = -120: t2w.TimerEnabled = 0: End Sub

Sub awt3_Hit(idx):vpmTimer.PulseSw 39:tw3.TransY = -113: t3w.TimerEnabled = 1:PlaySoundAt SoundFX("fx_target", DOFDropTargets), ActiveBall:End Sub
Sub t3w_Timer: tw3.TransY = -120: t3w.TimerEnabled = 0: End Sub

Sub awt4_Hit(idx):vpmTimer.PulseSw 38:tw4.TransY = -113: t4w.TimerEnabled = 1:PlaySoundAt SoundFX("fx_target", DOFDropTargets), ActiveBall:End Sub
Sub t4w_Timer: tw4.TransY = -120: t4w.TimerEnabled = 0: End Sub

Sub awt5_Hit(idx):vpmTimer.PulseSw 31:tw5.TransY = -113: t5w.TimerEnabled = 1:PlaySoundAt SoundFX("fx_target", DOFDropTargets), ActiveBall:End Sub
Sub t5w_Timer: tw5.TransY = -120: t5w.TimerEnabled = 0: End Sub

Sub awt6_Hit(idx):vpmTimer.PulseSw 32:tw6.TransY = -113: t6w.TimerEnabled = 1:PlaySoundAt SoundFX("fx_target", DOFDropTargets), ActiveBall:End Sub
Sub t6w_Timer: tw6.TransY = -120: t6w.TimerEnabled = 0: End Sub

Sub awt7_Hit(idx):vpmTimer.PulseSw 39:tw7.TransY = -113: t7w.TimerEnabled = 1:PlaySoundAt SoundFX("fx_target", DOFDropTargets), ActiveBall:End Sub
Sub t7w_Timer: tw7.TransY = -120: t7w.TimerEnabled = 0: End Sub

'*********
'Solenoids
'*********

SolCallback(5) = "SolFireMissile"
SolCallback(6) = "vpmSolSound SoundFX(""fx_Knocker"",DOFKnocker),"
SolCallback(7) = "dtLeft.SolDropUp"
SolCallback(8) = "dtRight.SolDropUp"
SolCallback(9) = "bsTrough.SolIn"
SolCallback(10) = "bsTrough.SolOut"
SolCallback(17) = "SolBallWalker"
SolCallback(18) = "SolRotorMagnet"
SolCallback(20) = "SolRotorRelay"
SolCallback(19) = "vpmNudge.SolGameOn"

Sub SolFireMissile(Enabled)
    If Enabled Then
        If(Controller.Switch(40))Then
            Controller.Switch(40) = 0
            sw40.Kick WheelAngle, 26
            PlaySound "fx_Kicker"
        End If
    End If
End Sub

Sub SolBallWalker(Enabled)
    If Enabled Then
        Lock1.IsDropped = 1
        Lock2.IsDropped = 1
        Lock3.IsDropped = 1
        PlaySound "fx_SolenoidOn", 1, 0.1, 0.25 ' TODO : Implement object
        ResetLock.Enabled = 1
    End If
End Sub

Sub ResetLock_Timer
    Me.Enabled = 0
    Lock1.IsDropped = 0
    Lock2.IsDropped = 0
    Lock3.IsDropped = 0
    PlaySound "fx_SolenoidOff", 1, 0.1, 0.25 ' TODO : Implement object
End Sub

Sub SolRotorRelay(Enabled)
    If Controller.Switch(40)Then
        WheelTimer.Enabled = 1
    End If
End Sub

Sub SolRotorMagnet(Enabled) 'if the magnet is disabled then shoot the ball
    If Not Enabled AND Controller.Switch(40)Then
        SolFireMissile 1
    End If
End Sub

'**************
' Flipper Subs
'**************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("fx_flipperup", DOFFlippers), LeftFlipper
        LeftFlipper.RotateToEnd
    Else
        PlaySoundAt SoundFX("fx_flipperdown", DOFFlippers), LeftFlipper
        LeftFlipper.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("fx_flipperup", DOFFlippers), RightFlipper
        RightFlipper.RotateToEnd
        RightFlipper1.RotateToEnd
    Else
        PlaySoundAt SoundFX("fx_flipperdown", DOFFlippers), RightFlipper
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

'*****************
'   Gi Effects
'*****************

Dim OldGiState
OldGiState = -1 'start witht he Gi off

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

Sub GiEffect(value) ' value is the duration of the blink
    For each x in aGiLights
        x.Duration 2, 500 * value, 1
    Next
End Sub

Sub GIUpdate
    Dim tmp, obj
    tmp = Getballs
    If UBound(tmp) <> OldGiState Then
        OldGiState = Ubound(tmp)
        If UBound(tmp) = -1 Then
            GiOff
        Else
            GiOn
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
LampTimer.Interval = 10 ' lamp fading speed
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
    If VarHidden Then
        UpdateLeds
    End If
    UpdateLamps
    GIUpdate
    RollingUpdate
End Sub

Sub UpdateLamps()
    'backdrop lights
    If VarHidden Then
        NFadeT 61, l61, "TILT"
        NFadeT 45, l45, "GAME OVER"
        NFadeT 13, l13, "HIGH SCORE"
        NFadeT 63, l63, "MATCH"
        NFadeT 64, l64, "BALL IN PLAY"
    End If
    NFadeL 1, l1
    NFadeL 2, l2
    NFadeL 3, l3
    NFadeL 4, l4
    NFadeL 5, l5
    NFadeL 6, l6
    NFadeL 7, l7
    NFadeL 8, l8
    NFadeL 9, l9
    NFadeLm 10, l10a
    NFadeL 10, l10
    NFadeL 11, l11
    NFadeL 12, l12
    NFadeL 14, l14
    NFadeLm 15, l15a
    NFadeL 15, l15
    'NFadeL 16, l16
    NFadeL 17, l17
    NFadeL 18, l18
    NFadeL 19, l19
    NFadeL 20, l20
    NFadeL 21, l21
    NFadeL 22, l22
    NFadeL 23, l23
    NFadeL 24, l24
    NFadeL 25, l25
    NFadeLm 26, l26a
    NFadeL 26, l26
    NFadeL 27, l27
    NFadeL 28, l28
    NFadeL 29, l29
    NFadeL 30, l30
    NFadeL 31, l31
    'NFadeL 32, l32
    NFadeL 33, l33
    NFadeL 34, l34
    NFadeL 35, l35
    NFadeL 36, l36
    NFadeL 37, l37
    NFadeL 38, l38
    NFadeL 39, l39
    NFadeL 40, l40
    NFadeL 41, l41
    'NFadeL 42, l42
    NFadeL 43, l43
    NFadeL 44, l44
    NFadeL 46, l46
    NFadeL 47, l47
    'NFadeL 48, l48
    NFadeL 50, l50
    NFadeL 51, l51
    NFadeL 52, l52
    NFadeL 53, l53
    NFadeL 54, l54
    NFadeL 55, l55
    NFadeL 56, l56
    NFadeL 57, l57
    'NFadeL 58, l58
    NFadeL 59, l59
    NFadeL 60, l60
    NFadeL 62, l62
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

Sub NFadeTm(nr, object, message)
    Select Case FadingLevel(nr)
        Case 4:object.Text = ""
        Case 5:object.Text = message
    End Select
End Sub

'************************************
'          LEDs Display
'     Based on Scapino's LEDs
'************************************

Dim Digits(32)
Dim Patterns(11)
Dim Patterns2(11)

Patterns(0) = 0     'empty
Patterns(1) = 63    '0
Patterns(2) = 6     '1
Patterns(3) = 91    '2
Patterns(4) = 79    '3
Patterns(5) = 102   '4
Patterns(6) = 109   '5
Patterns(7) = 125   '6
Patterns(8) = 7     '7
Patterns(9) = 127   '8
Patterns(10) = 111  '9

Patterns2(0) = 128  'empty
Patterns2(1) = 191  '0
Patterns2(2) = 134  '1
Patterns2(3) = 219  '2
Patterns2(4) = 207  '3
Patterns2(5) = 230  '4
Patterns2(6) = 237  '5
Patterns2(7) = 253  '6
Patterns2(8) = 135  '7
Patterns2(9) = 255  '8
Patterns2(10) = 239 '9

'Assign 7-digit output to reels
Set Digits(0) = a0
Set Digits(1) = a1
Set Digits(2) = a2
Set Digits(3) = a3
Set Digits(4) = a4
Set Digits(5) = a5
Set Digits(6) = a6

Set Digits(7) = b0
Set Digits(8) = b1
Set Digits(9) = b2
Set Digits(10) = b3
Set Digits(11) = b4
Set Digits(12) = b5
Set Digits(13) = b6

Set Digits(14) = c0
Set Digits(15) = c1
Set Digits(16) = c2
Set Digits(17) = c3
Set Digits(18) = c4
Set Digits(19) = c5
Set Digits(20) = c6

Set Digits(21) = d0
Set Digits(22) = d1
Set Digits(23) = d2
Set Digits(24) = d3
Set Digits(25) = d4
Set Digits(26) = d5
Set Digits(27) = d6

Set Digits(28) = e0
Set Digits(29) = e1
Set Digits(30) = e2
Set Digits(31) = e3

Sub UpdateLeds
    On Error Resume Next
    Dim ChgLED, ii, jj, chg, stat
    ChgLED = Controller.ChangedLEDs(&HFF, &HFFFF)
    If Not IsEmpty(ChgLED)Then
        For ii = 0 To UBound(ChgLED)
            chg = chgLED(ii, 1):stat = chgLED(ii, 2)
            For jj = 0 to 10
                If stat = Patterns(jj)OR stat = Patterns2(jj)then Digits(chgLED(ii, 0)).SetValue jj
            Next
        Next
    End IF
End Sub

'******************************
' Diverse Collection Hit Sounds
'******************************

Sub aMetals_Hit(idx):PlaySound "fx_MetalHit2", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Bands_Hit(idx):PlaySound "fx_rubber_band", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Posts_Hit(idx):PlaySound "fx_rubber_post", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Pins_Hit(idx):PlaySound "fx_rubber_pin", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aWoods_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub

'****************
'Wheel Animation
'****************

dim WheelAngle, WheelDir, WheelCurrPos, WheelOldPos

Sub WheelInit
    Dim i
    WheelCurrPos = 0
    WheelOldPos = 0
    WheelAngle = 0
    WheelDir = 2
    for i = 1 to 179
        aWguide(i).Isdropped = 1
        awp1(i).Isdropped = 1
        awp2(i).Isdropped = 1
        awp3(i).Isdropped = 1
        awp4(i).Isdropped = 1
        awp5(i).Isdropped = 1
        awp6(i).Isdropped = 1
        awt1(i).Isdropped = 1
        awt2(i).Isdropped = 1
        awt3(i).Isdropped = 1
        awt4(i).Isdropped = 1
        awt5(i).Isdropped = 1
        awt6(i).Isdropped = 1
        awt7(i).Isdropped = 1
    Next
End Sub

sub WheelTimer_Timer
    Dim i
    WheelOldPos = WheelCurrPos
    WheelAngle = WheelAngle + WheelDir
    If WheelAngle = 360 Then WheelAngle = 0
    If WheelAngle < 0 Then WheelAngle = 358
    For each i in aWheel
        i.rotz = WheelAngle
    Next
    WheelCurrPos = WheelAngle \ 2

    aWguide(WheelOldPos).Isdropped = 1
    awp1(WheelOldPos).Isdropped = 1
    awp2(WheelOldPos).Isdropped = 1
    awp3(WheelOldPos).Isdropped = 1
    awp4(WheelOldPos).Isdropped = 1
    awp5(WheelOldPos).Isdropped = 1
    awp6(WheelOldPos).Isdropped = 1
    awt1(WheelOldPos).Isdropped = 1
    awt2(WheelOldPos).Isdropped = 1
    awt3(WheelOldPos).Isdropped = 1
    awt4(WheelOldPos).Isdropped = 1
    awt5(WheelOldPos).Isdropped = 1
    awt6(WheelOldPos).Isdropped = 1
    awt7(WheelOldPos).Isdropped = 1

    aWguide(WheelCurrPos).Isdropped = 0
    awp1(WheelCurrPos).Isdropped = 0
    awp2(WheelCurrPos).Isdropped = 0
    awp3(WheelCurrPos).Isdropped = 0
    awp4(WheelCurrPos).Isdropped = 0
    awp5(WheelCurrPos).Isdropped = 0
    awp6(WheelCurrPos).Isdropped = 0
    awt1(WheelCurrPos).Isdropped = 0
    awt2(WheelCurrPos).Isdropped = 0
    awt3(WheelCurrPos).Isdropped = 0
    awt4(WheelCurrPos).Isdropped = 0
    awt5(WheelCurrPos).Isdropped = 0
    awt6(WheelCurrPos).Isdropped = 0
    awt7(WheelCurrPos).Isdropped = 0

    If WheelAngle = 0 AND Controller.Switch(40) = 0 Then ' no ball in viper, then stop the wheel
        WheelTimer.Enabled = 0
    End If
End Sub

' Dip Switches (from Eala's table)
Function DipVal(iSwitch):DipVal = 2 ^(iSwitch - 1):End Function
Sub myShowDips()
    Dim vpmDips:Set vpmDips = New cvpmDips
    With vpmDips
        .AddForm 700, 400, "Viper - DIP switches"
        .AddChk 2, 10, 180, Array("Match feature", DipVal(21))
        .AddChk 2, 25, 115, Array("Credits display", DipVal(20))
        .AddFrame 2, 45, 190, "Maximum credits", &H00060000, Array("10 credits", 0, "15 credits", &H00020000, "25 credits", &H00040000, "40 credits", &H00060000) 'dip 18&19
        .AddFrame 2, 120, 190, "High game to date", 49152, Array("points", 0, "1 free game", &H00004000, "2 free games", 32768, "3 free games", 49152)            'dip 15&16
        .AddFrame 205, 120, 190, "Special limit", DipVal(30), Array("one per game", 0, "one per ball", DipVal(30))
        .AddFrame 2, 195, 190, "Level Pass", DipVal(6), Array("extra ball", 0, "replay", DipVal(6))
        .AddFrame 2, 241, 190, "Level Pass Add-A-Ball", DipVal(13), Array("3", 0, "5", DipVal(13))
        .AddChk 2, 287, 190, Array("Extra Ball Feature", DipVal(22))
        .AddChk 205, 10, 190, Array("Add-A-Ball (Memory)", DipVal(14))
        .AddChk 205, 25, 190, Array("Background sound", DipVal(8))
        .AddFrame 205, 45, 190, "Special award", &HC0000000, Array("no award", 0, "100,000 points", &H40000000, "free ball", &H80000000, "free game", &HC0000000) 'dip 31&32
        .AddFrame 205, 195, 190, "Balls per game", DipVal(7), Array("3 balls", 0, "5 balls", DipVal(7))
        .AddFrame 205, 241, 190, "Extra Ball limit", DipVal(23), Array("one per game", 0, "one per ball", DipVal(23))
        .AddLabel 50, 385, 300, 20, "After hitting OK, press F3 to reset game with new settings."
        .ViewDips
    End With
End Sub
Set vpmShowDips = GetRef("myShowDips")

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
    If UBound(BOT) = lob - 1 Then Exit Sub 'there no extra balls on this table

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


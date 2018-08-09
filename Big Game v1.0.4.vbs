' Big Game -  Stern 1980
' Due to missing pages and errors on the manual, the light and solenoid definitions taken from the script by Destruk and Skalar
' VPX version 1.0.4 jpsalas - 2017

' Thalamus 2018-07-19
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' No special SSF tweaks yet.
' This is a JP table. He often uses walls as switches so I need to be careful of using PlaySoundAt
' Wob 2018-08-08
' Added vpmInit Me for FastFlips Support

Option Explicit

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const BallSize = 50

LoadVPM "01120100", "stern.vbs", 3.02

Dim VarHidden
If Table1.ShowDT = true then
    VarHidden = 1
else
    VarHidden = 0
    For each x in aReels
        x.Visible = 0
    Next
    lrail.Visible = 0
    rrail.Visible = 0
End If

Const cGameName = "biggame"

Dim bsTrough, bsSaucer, dtT, dtR, dtR2
Dim x

' Standard Sounds
Const SSolenoidOn = "fx_Solenoid" 'Solenoid activates
Const SSolenoidOff = ""           'Solenoid deactivates
Const SCoin = "fx_coin"           'Coin inserted

Const UseSolenoids = 2
Const UseLamps = 0
Const UseGI = 0

Sub Table1_Init
	vpmInit Me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Big Game, Stern 1980" & vbNewLine & "VPX table by jpsalas v.1.0.4"
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = VarHidden
    End With
    On Error Goto 0
    Controller.SolMask(0) = 0
    vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the timer to renable all the solenoids after 2 seconds
    Controller.Run

    PinmameTimer.Interval = PinMameInterval
    PinmameTimer.Enabled = 1

    ' trough def
    Set bsTrough = New cvpmBallstack
    bsTrough.InitNoTrough BallRelease, 33, 115, 3
    bsTrough.InitExitSnd "fx_ballrel", "fx_Solenoid"

    Set bsSaucer = New cvpmBallStack
    bsSaucer.InitSaucer sw34, 34, 55, 20
    bsSaucer.InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    bsSaucer.KickForceVar = 4

    ' droptarget def
    Set dtT = New cvpmDropTarget
    dtT.InitDrop Array(sw16, sw15, sw14), Array(16, 15, 14)
    dtT.InitSnd SoundFX("fx_droptarget", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)

    Set dtR = New cvpmDropTarget
    dtR.InitDrop Array(sw24, sw23, sw22), Array(24, 23, 22)
    dtR.InitSnd SoundFX("fx_droptarget", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)

    Set dtR2 = New cvpmDropTarget
    dtR2.InitDrop Array(sw32, sw31, sw30), Array(32, 31, 30)
    dtR2.InitSnd SoundFX("fx_droptarget", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)

    vpmNudge.TiltSwitch = 7
    vpmNudge.Sensitivity = 1
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingshot)
End Sub

Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub

' Keyboard handling

Sub table1_KeyDown(ByVal keycode)
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If keycode = PlungerKey Then PlaySound "fx_PlungerPull", 0, 1, 0.1, 0.25:Plunger.Pullback
    If vpmKeyDown(keycode)Then Exit Sub
End Sub

Sub table1_KeyUp(ByVal keycode)
    If vpmKeyUp(keycode)Then Exit Sub
    If keycode = PlungerKey Then PlaySound "fx_plunger", 0, 1, 0.1, 0.25:Plunger.Fire
End Sub

'**********
'Solenoids
'**********
SolCallback(7) = "dtT.SolDropUp"
SolCallback(8) = "dtR.SolDropUp"
SolCallback(9) = "dtR2.SolDropUp"
SolCallback(6) = "vpmsolsound SoundFX(""fx_knocker"",DOFKnocker),"
SolCallback(11) = "bsTrough.SolOut"
SolCallback(10) = "bsSaucer.SolOut"
SolCallback(19) = "vpmNudge.SolGameOn"

'********************
'     Flippers
'********************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySound SoundFX("fx_flipperup", DOFFlippers), 0, 1, -0.1, 0.25
        LeftFlipper.RotateToEnd
        LeftFlipper1.RotateToEnd
    Else
        PlaySound SoundFX("fx_flipperdown", DOFFlippers), 0, 1, -0.1, 0.25
        LeftFlipper.RotateToStart
        LeftFlipper1.RotateToStart
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

Sub LeftFlipper1_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, -0.1, 0.25
End Sub

Sub RightFlipper1_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, 0.1, 0.25
End Sub

'**********
' Switches
'**********

Sub sw4_Spin:vpmTimer.PulseSw(4):End Sub

Sub sw5_Spin:vpmTimer.PulseSw(5):End Sub

' Bumpers

Sub Bumper1_Hit:vpmTimer.PulseSw 9:PlaySound SoundFX("fx_bumper", DOFContactors), 0, 1, 0.15, 0.15:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 10:PlaySound SoundFX("fx_bumper", DOFContactors), 0, 1, 0.15, 0.15:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 11:PlaySound SoundFX("fx_bumper", DOFContactors), 0, 1, 0.15, 0.15:End Sub

Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    PlaySound SoundFX("fx_slingshot", DOFContactors), 0, 1, -0.05, 0.05
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 12
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
    vpmTimer.PulseSw 13
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

Sub sw37_Hit:vpmTimer.PulseSw(37):End Sub
Sub sw37a_Hit:vpmTimer.PulseSw(37):End Sub

'Droptargets
Sub sw14_Dropped:dtT.Hit 3:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw15_Dropped:dtT.Hit 2:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw16_Dropped:dtT.Hit 1:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw22_Dropped:dtR.Hit 3:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw23_Dropped:dtR.Hit 2:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw24_Dropped:dtR.Hit 1:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw30_Dropped:dtR2.Hit 3:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw31_Dropped:dtR2.Hit 2:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw32_Dropped:dtR2.Hit 1:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub

' Rollovers
Sub sw17_Hit:Controller.Switch(17) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw17_UnHit:Controller.Switch(17) = 0:End Sub
Sub sw18_Hit:Controller.Switch(18) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw18_UnHit:Controller.Switch(18) = 0:End Sub
Sub sw19_Hit:Controller.Switch(19) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw19_UnHit:Controller.Switch(19) = 0:End Sub
Sub sw27_Hit:Controller.Switch(27) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw27_UnHit:Controller.Switch(27) = 0:End Sub
Sub RightInlane_Hit:Controller.Switch(28) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub RightInlane_UnHit:Controller.Switch(28) = 0:End Sub
Sub LeftInlane_Hit:Controller.Switch(29) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub LeftInlane_UnHit:Controller.Switch(29) = 0:End Sub
Sub RightOutlane_Hit:Controller.Switch(35) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub RightOutlane_UnHit:Controller.Switch(35) = 0:End Sub
Sub LeftOutlane_Hit:Controller.Switch(36) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub LeftOutlane_UnHit:Controller.Switch(36) = 0:End Sub

'targets
Sub sw20_Hit:vpmTimer.PulseSw(20):PlaySound SoundFX("fx_target", DOFTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw21_Hit:vpmTimer.PulseSw(21):PlaySound SoundFX("fx_target", DOFTargets), 0, 1, pan(ActiveBall):End Sub

'Rollovers 2
Sub sw25_Hit:L16.State = 0:Controller.Switch(25) = 1:Me.TimerEnabled = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw25_UnHit:Controller.Switch(25) = 0:End Sub
Sub sw25_Timer:Me.TimerEnabled = 0:L16.State = 1:End Sub
Sub sw26_Hit:L30.State = 0:Controller.Switch(26) = 1:Me.TimerEnabled = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw26_UnHit:Controller.Switch(26) = 0:End Sub
Sub sw26_Timer:Me.TimerEnabled = 0:L30.State = 1:End Sub
Sub sw38_Hit:L31.State = 0:Controller.Switch(38) = 1:Me.TimerEnabled = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw38_UnHit:Controller.Switch(38) = 0:End Sub
Sub sw38_Timer:Me.TimerEnabled = 0:L31.State = 1:End Sub
Sub sw39_Hit:L15.State = 0:Controller.Switch(39) = 1:Me.TimerEnabled = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw39_UnHit:Controller.Switch(39) = 0:End Sub
Sub sw39_Timer:Me.TimerEnabled = 0:L15.State = 1:End Sub
Sub sw40_Hit:L14.State = 0:Controller.Switch(40) = 1:Me.TimerEnabled = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw40_UnHit:Controller.Switch(40) = 0:End Sub
Sub sw40_Timer:Me.TimerEnabled = 0:L14.State = 1:End Sub

' Drain & saucers
Sub Drain_Hit:bsTrough.AddBall Me:End Sub
'Sub Drain_Hit:Me.Destroyball:End Sub 'only for debug
Sub sw34_Hit:bsSaucer.AddBall 0:PlaySound "fx_kicker-enter":End Sub

'************************************
'          LEDs Display
'Based on Scapino's 7 digit Reel LEDs
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

'**************
' Gi update
'**************

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

Sub GiEffect
    For each x in aGiLights
        x.Duration 2, 3000, 1
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
    GIUpdate
    RollingUpdate
	ExtraLights
End Sub

Sub UpdateLamps
    NFadeL 1, l1
    NFadeL 2, l2
    NFadeL 3, l3
    NFadeL 4, l4
    NFadeL 5, l5
    NFadeLm 6, l48
    NFadeL 6, l6
    NFadeL 7, l7
    NFadeL 8, l8
    NFadeLm 9, l71
    NFadeL 9, l9
    NFadeLm 10, l67
    NFadeL 10, l10
    NFadeL 11, l11
    NFadeL 12, l12
    NFadeL 14, l14
    NFadeL 15, l15
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
    NFadeLm 25, l70
    NFadeL 25, l25
    NFadeLm 26, l66
    NFadeL 26, l26
    NFadeL 27, l27
    NFadeL 28, l28
    NFadeL 30, l30
    NFadeL 31, l31
    NFadeL 33, l33
    NFadeLm 34, l64
    NFadeL 34, l34
    NFadeL 35, l35
    NFadeL 36, l36
    NFadeL 37, l37
    NFadeL 38, l38
    NFadeL 39, l39
    NFadeL 40, l40
    NFadeLm 41, l69
    NFadeL 41, l41
    NFadeLm 42, l65
    NFadeL 42, l42
    NFadeL 43, l43
    NFadeL 44, l44
    NFadeL 46, l46
    NFadeL 48, l48
    NFadeL 49, l49
    NFadeL 50, l50
    NFadeL 51, l51
    NFadeL 52, l52
    NFadeLm 53, l32
    NFadeL 53, l53
    NFadeL 54, l54
    NFadeL 55, l55
    NFadeL 56, l56
    NFadeLm 57, l68
    NFadeL 57, l57
    NFadeL 58, l58
    NFadeL 59, l59
    NFadeL 60, l60
    NFadeL 62, l62
    NFadeL 64, l64
    NFadeL 72, l72

    ' backdrop lights
    If VarHidden Then
        NFadeT 13, l13, "High Score to Date"
        NFadeTm 29, l29, "Same Player Shoots Again"
		NFadeL 29, l47 'table's shoot again light
        NFadeT 45, l45, "Game Over"
        NFadeT 61, l61, "TILT"
        NFadeT 63, l63, "Match"
	Else
		NFadeL 29, l47 'table's shoot again light
    End If
End Sub

Sub ExtraLights()
    If(LampState(10) = 1 AND LampState(26) = 1 AND LampState(42) = 1)Then
        If LampState(72) = 0 Then SetLamp 72, 1
        Else
            If LampState(72) = 1 Then SetLamp 72, 0
    End If
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

'Stern Big Game by Inkochnito

Sub editDips
    Dim vpmDips:Set vpmDips = New cvpmDips
    With vpmDips
        .AddForm 700, 400, "Big Game - DIP switches"
        .AddChk 2, 10, 180, Array("Match feature", &H00100000)                                                                                                             'dip 21
        .AddChk 205, 10, 115, Array("Credits display", &H00080000)                                                                                                         'dip 20
        .AddFrame 2, 30, 190, "Maximum credits", &H00060000, Array("10 credits", 0, "15 credits", &H00020000, "25 credits", &H00040000, "40 credits", &H00060000)          'dip 18&19
        .AddFrame 2, 104, 190, "Completed line lite", &H03000000, Array("on 2nd ball", 0, "on 3rd ball", &H01000000, "on 4th ball", &H02000000, "on 5th ball", &H03000000) 'dip 25&26
        .AddFrame 2, 178, 190, "Special award", &HC0000000, Array("no award", 0, "100,000 points", &H40000000, "extra ball", &H80000000, "replay", &HC0000000)             'dip 31&32
        .AddFrame 2, 252, 190, "High game to date", 49152, Array("points", 0, "1 free game", &H00004000, "2 free games", 32768, "3 free games", 49152)                     'dip 15&16
        .AddFrame 205, 30, 190, "Balls per game", &H00000040, Array("3 balls", 0, "5 balls", &H00000040)                                                                   'dip 7
        .AddFrame 205, 76, 190, "Extra ball award", &H00010000, Array("no extra ball", 0, "award extra ball", &H00010000)                                                  'dip 17
        .AddFrame 205, 122, 190, "Special replay limit", &H00200000, Array("1 per ball in play", 0, "1 per game", &H00200000)                                              'dip 22
        .AddFrame 205, 168, 190, "Card selection", &H00400000, Array("X,Y,Z,XY,Z,YZ,X,XZ, repaet", 0, "X,Y,Z,X,Y,Z ect.", &H00400000)                                      'dip 23
        .AddFrame 205, 214, 190, "High score feature", &H00000020, Array("extra ball", 0, "replay", &H00000020)                                                            'dip 6
        .AddFrame 205, 260, 190, "Background sound", &H00000080, Array("off", 0, "on", &H00000080)                                                                         'dip 8
        .AddLabel 60, 335, 350, 20, "After hitting OK, press F3 to reset game with new settings."
        .ViewDips                                                                                                                                                          'Dip 14&24 not used
    End With
End Sub
Set vpmShowDips = GetRef("editDips")

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


' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


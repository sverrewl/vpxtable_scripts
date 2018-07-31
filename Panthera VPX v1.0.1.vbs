' Panthera - Gottlieb 1980
' IPD No. 1745 / June, 1980 / 4 Players
' VPX - version by JPSalas 2017, version 1.0.0

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

LoadVPM "01560000", "sys80.VBS", 3.37

'Variables
Dim bsTrough, dtL, dtR, dtT, bsLSaucer, x

Const cGameName = "panther7" '7 digits
'Const cGameName = "panthera"

Const UseSolenoids = 2
Const UseLamps = 0
Const UseGI = 0
Const UseSync = 0 'set it to 1 if the table runs too fast
Const HandleMech = 0

Dim VarHidden
If Table1.ShowDT = True then
    VarHidden = 1
    For each x in aReels
        x.Visible = 1
    Next
Else
    VarHidden = 0
    For each x in aReels
        x.Visible = 0
    Next

    lrail.Visible = 0
    rrail.Visible = 0
End If

if B2SOn = true then VarHidden = 1

' Standard Sounds
Const SSolenoidOn = "fx_Solenoid"
Const SSolenoidOff = ""
Const SCoin = "fx_Coin"

'Table Init
Sub table1_Init
    vpmInit me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Panthera 1980" & vbNewLine & "VPX table by JPSalas"
        '.Games(cGameName).Settings.Value("rol")=0
        '.SetDisplayPosition 0,0, GetPlayerHWnd 'restore dmd window position
        .Games(cGameName).Settings.Value("dmd_red") = 0
        .Games(cGameName).Settings.Value("dmd_green") = 223
        .Games(cGameName).Settings.Value("dmd_blue") = 223
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = VarHidden
        On Error Resume Next
        .SolMask(0) = 0
        vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - Then add the timer to renable all the solenoids after 2 seconds
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
        On Error Goto 0
    End With

    'Nudging
    vpmNudge.TiltSwitch = 57
    vpmNudge.Sensitivity = 4
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, LeftSlingshot, RightSlingShot, RightSlingShot2)

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0, 67, 0, 0, 0, 0, 0, 0
        .InitKick BallRelease, 90, 4
        .InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .Balls = 1
    End With

    'Left Saucer
    Set bsLSaucer = New cvpmBallStack
    With bsLSaucer
        .InitSaucer sw35, 35, 60, 13
        .InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .KickAngleVar = 2
    End With

    'Droptargets
    Set dtL = New cvpmDropTarget
    dtL.InitDrop Array(sw00, sw10, sw20, sw30), Array(00, 10, 20, 30)
    dtL.initsnd SoundFX("", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)

    Set dtT = New cvpmDropTarget
    dtT.InitDrop Array(sw01, sw11, sw21, sw31), Array(01, 11, 21, 31)
    dtT.initsnd SoundFX("", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)

    Set dtR = New cvpmDropTarget
    dtR.InitDrop Array(sw02, sw12, sw22, sw32), Array(02, 12, 22, 32)
    dtR.initsnd SoundFX("", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1
End Sub

'**********
' Keys
'**********

Sub Table1_KeyDown(ByVal KeyCode)
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If keycode = PlungerKey Then PlaySound "fx_PlungerPull", 0, 1, 0.1, 0.25:Plunger.Pullback
    If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
    If vpmKeyUp(keycode) Then Exit Sub
    If keycode = PlungerKey Then PlaySound "fx_plunger", 0, 1, 0.1, 0.25:Plunger.Fire
End Sub

' Slings
Dim LStep, RStep, RStep2

Sub LeftSlingShot_Slingshot
    PlaySound SoundFX("fx_slingshot", DOFContactors), 0, 1, -0.05, 0.05
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 34
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

Sub RightSlingShot2_Slingshot
    PlaySound SoundFX("fx_slingshot", DOFContactors), 0, 1, 0.05, 0.05
    RightSling8.Visible = 1
    Remk2.RotX = 26
    RStep2 = 0
    vpmTimer.PulseSw 34
    RightSlingShot2.TimerEnabled = 1
End Sub

Sub RightSlingShot2_Timer
    Select Case RStep2
        Case 1:RightSLing8.Visible = 0:RightSLing7.Visible = 1:Remk2.RotX = 14
        Case 2:RightSLing7.Visible = 0:RightSLing6.Visible = 1:Remk2.RotX = 2
        Case 3:RightSLing6.Visible = 0:Remk2.RotX = -10:RightSlingShot2.TimerEnabled = 0
    End Select

    RStep2 = RStep2 + 1
End Sub

' Bumpers

Sub Bumper1_Hit:vpmTimer.PulseSw 24:PlaySound SoundFX("fx_bumper", DOFContactors), 0, 1, 0, 0:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 24:PlaySound SoundFX("fx_bumper", DOFContactors), 0, 1, 0, 0:End Sub

' Rollovers
Sub sw03_Hit:Controller.Switch(03) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw03_UnHit:Controller.Switch(03) = 0:End Sub

Sub sw13_Hit:Controller.Switch(13) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw13_UnHit:Controller.Switch(13) = 0:End Sub

Sub sw23_Hit:Controller.Switch(23) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw23_UnHit:Controller.Switch(23) = 0:End Sub

Sub sw33_Hit:Controller.Switch(33) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw33_UnHit:Controller.Switch(33) = 0:End Sub

Sub sw05_Hit:Controller.Switch(05) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw05_UnHit:Controller.Switch(05) = 0:End Sub

Sub sw03b_Hit:Controller.Switch(03) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw03b_UnHit:Controller.Switch(03) = 0:End Sub

Sub sw13b_Hit:Controller.Switch(13) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw13b_UnHit:Controller.Switch(13) = 0:End Sub

Sub sw23b_Hit:Controller.Switch(23) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw23b_UnHit:Controller.Switch(23) = 0:End Sub

Sub sw33b_Hit:Controller.Switch(33) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw33b_UnHit:Controller.Switch(33) = 0:End Sub

Sub sw15_Hit:Controller.Switch(15) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw15_UnHit:Controller.Switch(15) = 0:End Sub

Sub sw15b_Hit:Controller.Switch(15) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw15b_UnHit:Controller.Switch(15) = 0:End Sub

Sub sw25b_Hit:Controller.Switch(25) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw25b_UnHit:Controller.Switch(25) = 0:End Sub

Sub sw14a_Hit:Controller.Switch(14) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):lsw14a.Duration 2, 1000, 0:End Sub
Sub sw14a_UnHit:Controller.Switch(14) = 0:end sub

Sub sw14b_Hit:Controller.Switch(14) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):lsw14b.Duration 2, 1000, 0:End Sub
Sub sw14b_UnHit:Controller.Switch(14) = 0:end sub

'Standup target

Sub sw04_Hit:vpmTimer.PulseSw 4:PlaySound SoundFX("fx_target", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub

'rubbers
Sub sw34a_Hit():vpmTimer.PulseSw 34:PlaySound SoundFX("fx_rubber_band", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw34b_Hit():vpmTimer.PulseSw 34:PlaySound SoundFX("fx_rubber_band", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw34c_Hit():vpmTimer.PulseSw 34:PlaySound SoundFX("fx_rubber_band", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw34d_Hit():vpmTimer.PulseSw 34:PlaySound SoundFX("fx_rubber_band", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw34e_Hit():vpmTimer.PulseSw 34:PlaySound SoundFX("fx_rubber_band", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw34f_Hit():vpmTimer.PulseSw 34:PlaySound SoundFX("fx_rubber_band", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw34g_Hit():vpmTimer.PulseSw 34:PlaySound SoundFX("fx_rubber_band", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub

'Spinner
Sub sw25_Spin():vpmTimer.PulseSw 25:PlaySound "fx_spinner", 0, 1, -0.1:End Sub

'Drop-Targets
Sub sw00_Hit():dtL.hit 1:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw10_Hit():dtL.hit 2:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw20_Hit():dtL.hit 3:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw30_Hit():dtL.hit 4:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub

Sub sw01_Hit():dtT.hit 1:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw11_Hit():dtT.hit 2:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw21_Hit():dtT.hit 3:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw31_Hit():dtT.hit 4:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub

Sub sw02_Hit():dtR.hit 1:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw12_Hit():dtR.hit 2:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw22_Hit():dtR.hit 3:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw32_Hit():dtR.hit 4:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub

' Drain & Holes
Sub Drain_Hit:bsTrough.AddBall Me:PlaySound "fx_drain":End Sub
Sub sw35_Hit:bsLSaucer.AddBall 0:Playsound "fx_kicker_enter", 0, 1, 0, 0.1:End Sub

'****Solenoids

SolCallback(5) = "dtT.soldropup"
SolCallback(2) = "dtL.soldropup"
SolCallback(1) = "dtR.soldropup"
SolCallback(6) = "bsLSaucer.SolOut"
SolCallback(8) = "vpmSolSound SoundFX(""fx_knocker"",DOFKnocker),"
SolCallback(9) = "bsTrough.SolOut"

'**************
' Flipper Subs
'**************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySound SoundFX("fx_flipperup", DOFFlippers), 0, 1, -0.1, 0.02
        LeftFlipper.RotateToEnd
        LeftFlipper1.RotateToEnd
        LeftFlipper2.RotateToEnd
        LeftFlipper3.RotateToEnd
    Else
        PlaySound SoundFX("fx_flipperdown", DOFFlippers), 0, 1, -0.1, 0.02
        LeftFlipper.RotateToStart
        LeftFlipper1.RotateToStart
        LeftFlipper2.RotateToStart
        LeftFlipper3.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySound SoundFX("fx_flipperup", DOFFlippers), 0, 1, 0.1, 0.02
        RightFlipper.RotateToEnd
        RightFlipper1.RotateToEnd
        RightFlipper2.RotateToEnd
        RightFlipper3.RotateToEnd
    Else
        PlaySound SoundFX("fx_flipperdown", DOFFlippers), 0, 1, 0.1, 0.02
        RightFlipper.RotateToStart
        RightFlipper1.RotateToStart
        RightFlipper2.RotateToStart
        RightFlipper3.RotateToStart
    End If
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, -0.1, 0.25
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, 0.1, 0.25
End Sub

'*****************
'   Gi Effects
'*****************

Dim OldGiState
OldGiState = -1 'start witht he Gi off

Sub GiON
    For each x in aGiLights
        GiEffect
    Next
End Sub

Sub GiOFF
    For each x in aGiLights
        x.State = 0
    Next
End Sub

Sub GiEffect
    For each x in aGiLights
        x.Duration 2, 1000, 1
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
    If Not IsEmpty(chgLamp) Then
        For ii = 0 To UBound(chgLamp)
            LampState(chgLamp(ii, 0) ) = chgLamp(ii, 1)       'keep the real state in an array
            FadingLevel(chgLamp(ii, 0) ) = chgLamp(ii, 1) + 4 'actual fading step
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
    NFadeT 1, l1, "TILT"
    NFadeL 3, l3
    NFadeL 4, l4
    NFadeL 5, l5
    NFadeL 6, l6
    NFadeL 7, l7
    'NFadeL 8, l8
    NFadeT 10, l10, "High Score to Date"
    NFadeT 11, l11, "Game Over"
    ' NFadeL 12, l12
    ' NFadeL 13, l13
    ' NFadeL 14, l14
    ' NFadeL 15, l15
    ' NFadeL 16, l16
    ' NFadeL 17, l17
    ' NFadeL 18, l18
    ' NFadeL 19, l19
    ' NFadeL 20, l20
    ' NFadeL 21, l21
    ' NFadeL 22, l22
    ' NFadeL 23, l23
    NFadeL 24, l24
    NFadeL 25, l25
    NFadeL 26, l26
    NFadeL 27, l27
    NFadeL 28, l28
    NFadeLm 29, l29b
    NFadeL 29, l29
    NFadeLm 30, l30b
    NFadeL 30, l30
    NFadeLm 31, l31b
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
    NFadeLm 44, l44d
    NFadeLm 44, l44b
    NFadeL 44, l44
    NFadeLm 45, l45d
    NFadeLm 45, l45b
    NFadeL 45, l45
    NFadeLm 46, l46d
    NFadeLm 46, l46b
    NFadeL 46, l46
    NFadeLm 47, l47d
    NFadeLm 47, l47b
    NFadeL 47, l47
    NFadeLm 48, l48b
    NFadeL 48, l48
    NFadeLm 49, l49b
    NFadeL 49, l49
    NFadeLm 50, l50b
    NFadeL 50, l50
    NFadeLm 51, l51b
    NFadeL 51, l51
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
            FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown(nr)
            If FlashLevel(nr) <FlashMin(nr) Then
                FlashLevel(nr) = FlashMin(nr)
                FadingLevel(nr) = 0 'completely off
            End if
            Object.IntensityScale = FlashLevel(nr)
        Case 5 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
            If FlashLevel(nr)> FlashMax(nr) Then
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
            FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown(nr)
            If FlashLevel(nr) <FlashMin(nr) Then
                FlashLevel(nr) = FlashMin(nr)
                FadingLevel(nr) = 0 'completely off
            End if
            Object.IntensityScale = FlashLevel(nr)
            If FadingLevel(nr) = 0 AND FlashRepeat(nr) Then 'repeat the flash
                FlashRepeat(nr) = FlashRepeat(nr) -1
                If FlashRepeat(nr) Then FadingLevel(nr) = 5
            End If
        Case 5 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
            If FlashLevel(nr)> FlashMax(nr) Then
                FlashLevel(nr) = FlashMax(nr)
                FadingLevel(nr) = 1 'completely on
            End if
            Object.IntensityScale = FlashLevel(nr)
            If FadingLevel(nr) = 1 AND FlashRepeat(nr) Then FadingLevel(nr) = 4
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
'      Gottlieb led Patterns
'************************************

Dim Digits(32)
Dim Patterns(11) 'normal numbers
Dim Patterns2(11) 'numbers with a comma
Dim Patterns3(11) 'the credits and ball in play

Patterns(0) = 0     'empty
Patterns(1) = 63    '0
Patterns(2) = 768   '1
Patterns(3) = 91    '2
Patterns(4) = 79    '3
Patterns(5) = 102   '4
Patterns(6) = 109   '5
Patterns(7) = 124   '6
Patterns(8) = 7     '7
Patterns(9) = 127   '8
Patterns(10) = 103  '9

Patterns2(0) = 128  'empty
Patterns2(1) = 191  '0
Patterns2(2) = 896  '1
Patterns2(3) = 219  '2
Patterns2(4) = 207  '3
Patterns2(5) = 230  '4
Patterns2(6) = 237  '5
Patterns2(7) = 253  '6
Patterns2(8) = 135  '7
Patterns2(9) = 255  '8
Patterns2(10) = 231 '9

Patterns3(0) = 0     'empty
Patterns3(1) = 63    '0
Patterns3(2) = 6     '1
Patterns3(3) = 91    '2
Patterns3(4) = 79    '3
Patterns3(5) = 102   '4
Patterns3(6) = 109   '5
Patterns3(7) = 124   '6
Patterns3(8) = 7     '7
Patterns3(9) = 127   '8
Patterns3(10) = 103  '9

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
    dim oldstat
    Dim ChgLED, ii, jj, chg, stat
    ChgLED = Controller.ChangedLEDs(&HFF, &HFFFF)
    If Not IsEmpty(ChgLED) Then
        For ii = 0 To UBound(ChgLED)
            chg = chgLED(ii, 1):stat = chgLED(ii, 2)
            For jj = 0 to 10
if oldstat <> STAT then
debug.print STAT
oldstat = stat
end if
                If(stat = Patterns(jj) ) OR (stat = Patterns2(jj) ) OR (stat = Patterns3(jj) ) then Digits(chgLED(ii, 0) ).SetValue jj
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
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) + 13
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) - 13
        End If
        ballShadow(b).Y = BOT(b).Y + 10
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
End Sub

'Les transistors inutilisés qui servent pour les lampes alimentent la memoire des DropTargets
Dim N1, O1, N2, O2, N3, O3, N4, O4, N5, O5, N6, O6, N7, O7, N8, O8, N9, O9, N10, O10, N11, O11, N12, O12
N1 = 0:O1 = 0:N2 = 0:O2 = 0:N3 = 0:O3 = 0:N4 = 0:O4 = 0:N5 = 0:O5 = 0:N6 = 0:O6 = 0:N7 = 0:O7 = 0:N8 = 0:O8 = 0:N9 = 0:O9 = 0:N10 = 0:O10 = 0:N11 = 0:O11 = 0:N12 = 0:O12 = 0

Set LampCallback = GetRef("UpdateMultipleLamps")

Sub UpdateMultipleLamps
    '1ere bank
    N1 = Controller.Lamp(12)
    If N1 <> O1 Then
        If N1 Then sw00.isdropped = 1:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, -0.1
        O1 = N1
    End If
    N2 = Controller.Lamp(13)
    If N2 <> O2 Then
        If N2 Then sw10.isdropped = 1:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, -0.1
        O2 = N2
    End If
    N3 = Controller.Lamp(14)
    If N3 <> O3 Then
        If N3 Then sw20.isdropped = 1:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, -0.1
        O3 = N3
    End If
    N4 = Controller.Lamp(15)
    If N4 <> O4 Then
        If N4 Then sw30.isdropped = 1:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, -0.1
        O4 = N4
    End If

    '2eme bank
    N5 = Controller.Lamp(16)
    If N5 <> O5 Then
        If N5 Then sw01.isdropped = 1:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, -0.05
        O5 = N5
    End If
    N6 = Controller.Lamp(17)
    If N6 <> O6 Then
        If N6 Then sw11.isdropped = 1:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, -0.05
        O6 = N6
    End If
    N7 = Controller.Lamp(18)
    If N7 <> O7 Then
        If N7 Then sw21.isdropped = 1:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, -0.05
        O7 = N7
    End If
    N8 = Controller.Lamp(19)
    If N8 <> O8 Then
        If N8 Then sw31.isdropped = 1:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, -0.05
        O8 = N8
    End If

    '3eme bank
    N9 = Controller.Lamp(20)
    If N9 <> O9 Then
        If N9 Then sw02.isdropped = 1:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, 0.1
        O9 = N9
    End If
    N10 = Controller.Lamp(21)
    If N10 <> O10 Then
        If N10 Then sw12.isdropped = 1:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, 0.1
        O10 = N10
    End If
    N11 = Controller.Lamp(22)
    If N11 <> O11 Then
        If N11 Then sw22.isdropped = 1:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, 0.1
        O11 = N11
    End If
    N12 = Controller.Lamp(23)
    If N12 <> O12 Then
        If N12 Then sw32.isdropped = 1:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, 0.1
        O12 = N12
    End If
End Sub

'Gottlieb Panthera
'added by Inkochnito
Sub editDips
    Dim vpmDips:Set vpmDips = New cvpmDips
    With vpmDips
        .AddForm 700, 400, "Panthera - DIP switches"
        .AddFrame 2, 10, 190, "Maximum credits", 49152, Array("8 credits", 0, "10 credits", 32768, "15 credits", &H00004000, "25 credits", 49152)                                                                                  'dip 15&16
        .AddFrame 2, 86, 190, "Coin chute 1 and 2 control", &H00002000, Array("seperate", 0, "same", &H00002000)                                                                                                                   'dip 14
        .AddFrame 2, 132, 190, "Playfield special", &H00200000, Array("replay", 0, "extra ball", &H00200000)                                                                                                                       'dip 22
        .AddFrame 2, 178, 190, "Hole for special", &H80000000, Array("alternating", 0, "stays lit", &H80000000)                                                                                                                    'dip32
        .AddFrame 2, 224, 190, "High score to date awards", &H00C00000, Array("not displayed and no award", 0, "displayed and no award", &H00800000, "displayed and 2 replays", &H00400000, "displayed and 3 replays", &H00C00000) 'dip 23&24
        .AddChk 2, 300, 190, Array("Sound when scoring?", &H01000000)                                                                                                                                                              'dip 25
        .AddChk 2, 315, 190, Array("Replay button tune?", &H02000000)                                                                                                                                                              'dip 26
        .AddChk 2, 330, 190, Array("Coin switch tune?", &H04000000)                                                                                                                                                                'dip 27
        .AddChk 2, 345, 190, Array("Credits displayed?", &H08000000)                                                                                                                                                               'dip 28
        .AddChk 2, 360, 190, Array("Match feature", &H00020000)                                                                                                                                                                    'dip 18
        .AddChk 2, 375, 190, Array("Attract features", &H20000000)                                                                                                                                                                 'dip 30
        .AddChkExtra 2, 390, 190, Array("Background sound off", &H0100)                                                                                                                                                            'S-board dip 1
        .AddFrameExtra 205, 10, 190, "Attract tune", &H0200, Array("no attract tune", 0, "attract tune played every 6 minutes", &H0200)                                                                                            'S-board dip 2
        .AddFrame 205, 56, 190, "Balls per game", &H00010000, Array("5 balls", 0, "3 balls", &H00010000)                                                                                                                           'dip 17
        .AddFrame 205, 102, 190, "Replay limit", &H00040000, Array("no limit", 0, "one per ball", &H00040000)                                                                                                                      'dip 19
        .AddFrame 205, 148, 190, "Novelty", &H00080000, Array("normal game mode", 0, "50,000 points for special/extra ball", &H0080000)                                                                                            'dip 20
        .AddFrame 205, 194, 190, "Game mode", &H00100000, Array("replay", 0, "extra ball", &H00100000)                                                                                                                             'dip 21
        .AddFrame 205, 240, 190, "3rd coin chute credits control", &H00001000, Array("no effect", 0, "add 9", &H00001000)                                                                                                          'dip 13
        .AddFrame 205, 286, 190, "Tilt penalty", &H10000000, Array("game over", 0, "ball in play", &H10000000)                                                                                                                     'dip 29
        .AddFrame 205, 332, 190, "Extra ball target adjust", &H40000000, Array("alternating", 0, "stays lit", &H40000000)                                                                                                          'dip 31
        .AddLabel 50, 420, 300, 20, "After hitting OK, press F3 to reset game with new settings."
    End With

    Dim extra
    extra = Controller.Dip(4) + Controller.Dip(5) * 256
    extra = vpmDips.ViewDipsExtra(extra)
    Controller.Dip(4) = extra And 255
    Controller.Dip(5) = (extra And 65280) \ 256 And 255
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

Const tnob = 5 ' total number of balls
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


' Thalamus : Exit in a clean and proper way
Sub table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


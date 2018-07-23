' VPX Jokerz! by JPSalas

Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName = "jokrz_l6"

LoadVPM "01530000", "S11.VBS", 3.10

Const UseSolenoids = 1
Const UseLamps = 0
Const UseGI = 0
Const UseSync = 0
Const HandleMech = 0

Dim VarHidden
If Table1.ShowDT = true then
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
Const SSolenoidOn = "Solenoid"
Const SSolenoidOff = ""
Const SCoin = "Coin"

dim bsTrough, bsLock, dtTL, dtBL, dtBR, mWheel
dim x, mRDir
Dim RampPos, RampDir

' Init table

Sub table1_Init()
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game: " & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "based on Jokerz! (Williams 1988)" & vbNewLine & "VPX table by JPSalas v1.0.2"
        .Games(cGameName).Settings.Value("rol") = 0
        .HandleMechanics = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .ShowTitle = 0
        .Hidden = VarHidden
        '.SetDisplayPosition 0, 0, GetPlayerHWnd   'uncomment this line If you don't see the vpm window
        On Error Resume Next
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
        On Error Goto 0
    End With

    On Error Goto 0

    ' Nudging
    vpmNudge.TiltSwitch = swTilt
    vpmNudge.Sensitivity = 3
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingshot)

    ' Trough handler
    Set bsTrough = New cvpmBallStack
    bsTrough.InitSw 0, 11, 12, 13, 0, 0, 0, 0
    bsTrough.InitKick BallRelease, 60, 6
    bsTrough.Balls = 3
    bsTrough.InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)

    Set dtTL = New cvpmDropTarget
    dtTL.InitDrop Array(sw41, sw42, sw43), Array(41, 42, 43)
    dtTL.initsnd SoundFX("", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)

    Set dtBL = New cvpmDropTarget
    dtBL.InitDrop Array(sw25, sw26, sw27), Array(25, 26, 27)
    dtBL.initsnd SoundFX("", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)

    Set dtBR = New cvpmDropTarget
    dtBR.InitDrop Array(sw36, sw37, sw38), Array(36, 37, 38)
    dtBR.initsnd SoundFX("", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)

    Set mWheel = new cvpmMech
    With mWheel
        .Length = 200
        .Steps = 34
        .mType = vpmMechStepSol + vpmMechCircle + vpmMechLinear + vpmMechFast
        .Sol1 = 14
        .Sol2 = 13
        .Addsw 59, 0, 6
        .Callback = GetRef("UpdateWheel")
        .Start
    End With

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

    ' Init center ramp
    RampPos = 95
    RampDir = -1
    LowerRamp
End Sub

Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub
Sub table1_exit:Controller.stop:End Sub

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
' Solenoids
'**********

SolCallback(2) = "bsTrough.SolOut"
SolCallback(3) = "SolBLDropTgt"
SolCallback(4) = "SolBRDropTgt"
SolCallback(5) = "SolTLKicker"
SolCallback(6) = "SolTLDropTgt"
SolCallback(7) = "SolBLKicker"
SolCallback(8) = "vpmSolSound SoundFX(""fx_Knocker"",DOFKnocker),"
SolCallback(10) = "SolGi"
SolCallback(16) = "SolTREject"
SolCallback(15) = "SolRamp"

' Flashers
SolCallback(9) = "SetLamp 109,"
SolCallback(11) = "SetLamp 111,"
SolCallback(22) = "SetLamp 122,"
SolCallback(25) = "SetLamp 125,"
SolCallback(26) = "SetLamp 126,"
SolCallback(27) = "SetLamp 127,"
SolCallback(28) = "SetLamp 128,"
SolCallback(29) = "SetLamp 129,"
SolCallback(30) = "SetLamp 130,"
SolCallback(31) = "SetLamp 131,"
SolCallback(32) = "SetLamp 132,"

' Solenoid subs

Sub SolGi(Enabled)
    If Enabled Then
        GIOff
    Else
        GiOn
    End If
End Sub

Sub SolTLDropTgt(Enabled)
    If Enabled Then dtTL.DropSol_On
End Sub

Sub SolBLDropTgt(Enabled)
    If Enabled Then dtBL.DropSol_On
End Sub

Sub SolBRDropTgt(Enabled)
    If Enabled Then dtBR.DropSol_On
End Sub

Sub SolTLKicker(Enabled)
    If Enabled Then
        Controller.Switch(44) = 0
        PlaySound SoundFX("fx_popper", DOFContactors), 0, 1, -0.05, 0.1
        sw44.Kick 20, 26
    End If
End Sub

Sub SolBLKicker(Enabled)
    If Enabled Then
        PlaySound SoundFX("fx_popper", DOFContactors), 0, 1, -0.05, 0.1
        Controller.Switch(46) = 0
        Controller.Switch(45) = 0
        sw46.Kick 40, 26 + RND(1) * 5
        sw45.TimerEnabled = 1
    End If
End Sub

Sub SolTREject(Enabled)
    If Enabled Then
        PlaySound SoundFX("fx_popper", DOFContactors), 0, 1, 0.05, 0.1
        Controller.Switch(47) = 0
        sw47.Kick 320, 26 + RND(1) * 5
    End If
End Sub

Sub Updatewheel(aNewPos, aSpeed, aLastPos)
    Dim tmp
    tmp = aNewPos \ 2
    If tmp> 9 then tmp = tmp -1
    pokercards.SetValue(tmp)
End Sub

'*************
' Center Ramp
'*************

Sub SolRamp(Enabled)
    If Enabled Then
        PlaySound SoundFX("fx_motor", DOFGear), -1
        If mRDir = 0 Then
            LowerRamp
            iRamp.Collidable = False  'bottom ramp - side walls
            iRamp1.Collidable = False ' high ramp
            iKicker1.Enabled = False
            vpmTimer.AddTimer 200, "LowerRamp'"
            vpmTimer.AddTimer 400, "LowerRamp'"
            vpmTimer.AddTimer 600, "LowerRamp'"
            vpmTimer.AddTimer 700, "LowerRamp'"
            vpmTimer.AddTimer 1000, "LowerRamp'"
            vpmTimer.AddTimer 1100, "RampIsDown'"
        Else
            RiseRamp
            vpmTimer.AddTimer 200, "RiseRamp'"
            vpmTimer.AddTimer 400, "RiseRamp'"
            vpmTimer.AddTimer 600, "RiseRamp'"
            vpmTimer.AddTimer 800, "RiseRamp'"
            vpmTimer.AddTimer 1000, "RiseRamp'"
            vpmTimer.AddTimer 1100, "RampIsUp'"
        End If
    End If
End Sub

Sub RiseRamp:vpmtimer.pulsesw 16:RampDir = 1:RampTimer.Enabled = 1:End Sub
Sub RampIsUp
    Controller.Switch(15) = 1
    mRDir = 0
    iRamp.Collidable = True  'bottom ramp - side walls
    iRamp1.Collidable = True ' high ramp
    iKicker1.Enabled = True
End Sub
Sub LowerRamp:vpmtimer.pulsesw 15:RampDir = -1:RampTimer.Enabled = 1:End Sub
Sub RampIsDown:Controller.Switch(16) = 1:mRDir = 1:End Sub

Sub RampTimer_Timer
    RampPos = RampPos + RampDir
    If RampPos> 95 Then
        RampPos = 95
        StopSound "fx_motor"
        RampTimer.Enabled = 0
    End If
    If RampPos <0 Then
        RampPos = 0
        StopSound "fx_motor"
        RampTimer.Enabled = 0
    End If
    CenterRamp.HeightTop = 1 + RampPos
End Sub

Sub iKicker1_Hit:Me.DestroyBall:iKicker2.Createball:iKicker2.Kick 0, 0:End Sub

'**************
' Flipper Subs
'**************

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
    Else
        PlaySound SoundFX("fx_flipperdown", DOFFlippers), 0, 1, 0.1, 0.25
        RightFlipper.RotateToStart
    End If
End Sub

Sub SolRFlipper1(Enabled)
    LowerFlipper = Enabled
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, -0.1, 0.25
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, 0.1, 0.25
End Sub

'*******************
' Switches & Targets
'*******************

' saucers

Sub drain_Hit:PlaySound "fx_drain":Drain.DestroyBall:vpmTimer.PulseSwitch(10), 100, "HandleDrain":End Sub

Sub HandleDrain(dummy):bsTrough.AddBall 0:End Sub

Sub sw44_Hit():Controller.Switch(44) = 1:PlaySound "fx_kicker_enter" End Sub

Sub sw45_Hit():Controller.Switch(45) = 1:sw46.Enabled = True:PlaySound "fx_kicker_enter":End Sub

Sub sw46_Hit():Controller.Switch(46) = 1:sw46.Enabled = False:PlaySound "fx_kicker_enter":End Sub

Sub sw45_Timer:sw45.TimerEnabled = 0:sw45.Kick 40, 26 + RND(1) * 5:End Sub

Sub sw47_Hit():Controller.Switch(47) = 1:PlaySound "fx_kicker_enter" End Sub

' slings

Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    PlaySound SoundFX("fx_slingshot", DOFContactors), 0, 1, -0.05, 0.05
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 63
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
    vpmTimer.PulseSw 64
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

Sub sw53_Hit():vpmtimer.pulsesw 53:Playsound "fx_rubber":End Sub
Sub sw56_Hit():vpmtimer.pulsesw 56:Playsound "fx_rubber":End Sub

' bumpers

Sub Bumper1_Hit:vpmTimer.PulseSw 60:PlaySound SoundFX("fx_bumper", DOFContactors), 0, 1, -0.15, 0.15:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 61:PlaySound SoundFX("fx_bumper", DOFContactors), 0, 1, 0.15, 0.15:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 62:PlaySound SoundFX("fx_bumper", DOFContactors), 0, 1, 0, 0.15:End Sub

' lanes

Sub sw14_Hit():Controller.Switch(14) = 1:Playsound "fx_sensor":End Sub
Sub sw14_UnHit():Controller.Switch(14) = 0:End Sub

Sub sw20_Hit():Controller.Switch(20) = 1:Playsound "fx_sensor":End Sub
Sub sw20_UnHit():Controller.Switch(20) = 0:End Sub

Sub sw21_Hit():Controller.Switch(21) = 1:Playsound "fx_sensor":End Sub
Sub sw21_UnHit():Controller.Switch(21) = 0:End Sub

Sub sw22_Hit():Controller.Switch(22) = 1:Playsound "fx_sensor":End Sub
Sub sw22_UnHit():Controller.Switch(22) = 0:End Sub

Sub sw23_Hit():Controller.Switch(23) = 1:Playsound "fx_sensor":End Sub
Sub sw23_UnHit():Controller.Switch(23) = 0:End Sub

' top lanes

Sub sw17_Hit():Controller.Switch(17) = 1:Playsound "fx_sensor":End Sub
Sub sw17_UnHit():Controller.Switch(17) = 0:End Sub

Sub sw18_Hit():Controller.Switch(18) = 1:Playsound "fx_sensor":End Sub
Sub sw18_UnHit():Controller.Switch(18) = 0:End Sub

Sub sw19_Hit():Controller.Switch(19) = 1:Playsound "fx_sensor":End Sub
Sub sw19_UnHit():Controller.Switch(19) = 0:End Sub

Sub sw28_Hit():Controller.Switch(28) = 1:Playsound "fx_sensor":End Sub
Sub sw28_UnHit():Controller.Switch(28) = 0:End Sub

Sub sw29_Hit():Controller.Switch(29) = 1:Playsound "fx_sensor":End Sub
Sub sw29_UnHit():Controller.Switch(29) = 0:End Sub

Sub sw30_Hit():Controller.Switch(30) = 1:Playsound "fx_sensor":End Sub
Sub sw30_UnHit():Controller.Switch(30) = 0:End Sub

Sub sw31_Hit():Controller.Switch(31) = 1:Playsound "fx_sensor":End Sub
Sub sw31_UnHit():Controller.Switch(31) = 0:End Sub

Sub sw32_Hit():Controller.Switch(32) = 1:Playsound "fx_sensor":End Sub
Sub sw32_UnHit():Controller.Switch(32) = 0:End Sub

' ramps

Sub sw49_Hit():Controller.Switch(49) = 1:Playsound "fx_sensor":End Sub
Sub sw49_UnHit():Controller.Switch(49) = 0:End Sub

Sub sw50_Hit():Controller.Switch(50) = 1:Playsound "fx_sensor":End Sub
Sub sw50_UnHit():Controller.Switch(50) = 0:End Sub

' Droptagets
Sub sw41_Hit:dtTL.Hit 1:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw42_Hit:dtTL.Hit 2:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw43_Hit:dtTL.Hit 3:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw25_Hit:dtBL.Hit 1:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw26_Hit:dtBL.Hit 2:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw27_Hit:dtBL.Hit 3:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw36_Hit:dtBR.Hit 1:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw37_Hit:dtBR.Hit 2:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw38_Hit:dtBR.Hit 3:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub

' targets

Sub sw52_Hit:vpmTimer.PulseSw 52:PlaySound SoundFX("fx_target", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub

Sub sw24_Hit:vpmTimer.PulseSw 24:PlaySound SoundFX("fx_target", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub

' spinners
Sub sw48_Spin():vpmTimer.pulsesw 48:Playsound "fx_spinner", 0, 1, 0, 0.08:End Sub

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
    UpdateLeds
End Sub

Sub UpdateLamps
    NFadeL 1, l1
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
    NFadeLm 34, l34a
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
    NFadeLm 49, l49a
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
    NFadeL 60, l60
    NFadeL 61, l61
    NFadeL 62, l62
    NFadeL 63, l63
    NFadeL 64, l64

    ' flashers

    Flashm 109, f09b
    Flash 109, f09c
    Flash 111, f11
    NFadeLm 122, f09a
    NFadeLm 122, f09
    NFadeLm 122, f22a
    NFadeL 122, f22
    Flash 125, f25
    Flash 126, f26
    Flash 127, f27
    Flash 128, f28
    Flash 129, f29
    Flash 130, f30
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
            If FlashLevel(nr) <FlashMin(nr)Then
                FlashLevel(nr) = FlashMin(nr)
                FadingLevel(nr) = 0 'completely off
            End if
            Object.IntensityScale = FlashLevel(nr)
        Case 5 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
            If FlashLevel(nr)> FlashMax(nr)Then
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
            If FlashLevel(nr) <FlashMin(nr)Then
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
            If FlashLevel(nr)> FlashMax(nr)Then
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

Sub aMetals_Hit(idx):PlaySound "fx_MetalHit2", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aRubber_Bands_Hit(idx):PlaySound "fx_rubber_band", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aRubber_Posts_Hit(idx):PlaySound "fx_rubber_post", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aRubber_Pins_Hit(idx):PlaySound "fx_rubber_pin", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aWoods_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub

' Ramp Sounds
Sub REnd1_Hit()
    ActiveBall.VelX = 0
    ActiveBall.VelY = 0
    ActiveBall.VelZ = 0
    PlaySound "fx_ballrampdrop", 0, 1, pan(ActiveBall)
End Sub

Sub REnd2_Hit()
    ActiveBall.VelX = 0
    ActiveBall.VelY = 0
    ActiveBall.VelZ = 0
    PlaySound "fx_balldrop", 0, 1, pan(ActiveBall)
End Sub

' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 1000)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / table1.width-1
    If tmp> 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10))
    End If
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2)))
End Function

'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 10 ' total number of balls
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
    For b = lob to UBound(BOT)
        If BallVel(BOT(b))> 1 Then
            If BOT(b).z <30 Then
                ballpitch = Pitch(BOT(b))
            Else
                ballpitch = Pitch(BOT(b)) * 100
            End If
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b)), Pan(BOT(b)), 0, ballpitch, 1, 0
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
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0
End Sub

'******************
' RealTime Updates
'******************

Set MotorCallback = GetRef("RealTimeUpdates")

Sub RealTimeUpdates
    RollingUpdate
'GIUpdate
End Sub

'******************
'   GI effects
' independent routine
' it turns on the gi
' when there is a ball
' in play
'******************

Dim OldGiState, GiIsOn
OldGiState = -1 'start witht the Gi off
GiIsOn = False

Sub GIUpdate
    Dim tmp, obj
    tmp = Getballs
    If UBound(tmp) <> OldGiState Then
        OldGiState = Ubound(tmp)
        If UBound(tmp) = 0 Then '-1 is the value returned by Getballs when no balls are on the table, so turn off the gi, since there is a ball in the lower playfield then we increase by 1, so -1 +1 = 0
            GiOff
        Else
            Gion
        End If
    End If
End Sub

Sub GiOn
    Dim bulb
    GiIsOn = True
    For each bulb in aGiLights
        bulb.State = 1
    Next
    Bumper1Light.Visible = 1
    Bumper1Light.Visible = 1
    Bumper1Light.Visible = 1
End Sub

Sub GiOff
    Dim bulb
    GiIsOn = False
    For each bulb in aGiLights
        bulb.State = 0
    Next
    Bumper1Light.Visible = 0
    Bumper1Light.Visible = 0
    Bumper1Light.Visible = 0
End Sub

' LED display
' Based on the Eala's rutine

Dim Digits(32)
Digits(0) = Array(a00, a05, a0c, a0d, a08, a01, a06, a0f, a02, a03, a04, a07, a0b, a0a, a09, a0e)
Digits(1) = Array(a10, a15, a1c, a1d, a18, a11, a16, a1f, a12, a13, a14, a17, a1b, a1a, a19, a1e)
Digits(2) = Array(a20, a25, a2c, a2d, a28, a21, a26, a2f, a22, a23, a24, a27, a2b, a2a, a29, a2e)
Digits(3) = Array(a30, a35, a3c, a3d, a38, a31, a36, a3f, a32, a33, a34, a37, a3b, a3a, a39, a3e)
Digits(4) = Array(a40, a45, a4c, a4d, a48, a41, a46, a4f, a42, a43, a44, a47, a4b, a4a, a49, a4e)
Digits(5) = Array(a50, a55, a5c, a5d, a58, a51, a56, a5f, a52, a53, a54, a57, a5b, a5a, a59, a5e)
Digits(6) = Array(a60, a65, a6c, a6d, a68, a61, a66, a6f, a62, a63, a64, a67, a6b, a6a, a69, a6e)
Digits(7) = Array(a70, a75, a7c, a7d, a78, a71, a76, a7f, a72, a73, a74, a77, a7b, a7a, a79, a7e)
Digits(8) = Array(a80, a85, a8c, a8d, a88, a81, a86, a8f, a82, a83, a84, a87, a8b, a8a, a89, a8e)
Digits(9) = Array(a90, a95, a9c, a9d, a98, a91, a96, a9f, a92, a93, a94, a97, a9b, a9a, a99, a9e)
Digits(10) = Array(aa0, aa5, aac, aad, aa8, aa1, aa6, aaf, aa2, aa3, aa4, aa7, aab, aaa, aa9, aae)
Digits(11) = Array(ab0, ab5, abc, abd, ab8, ab1, ab6, abf, ab2, ab3, ab4, ab7, abb, aba, ab9, abe)
Digits(12) = Array(ac0, ac5, acc, acd, ac8, ac1, ac6, acf, ac2, ac3, ac4, ac7, acb, aca, ac9, ace)
Digits(13) = Array(ad0, ad5, adc, add, ad8, ad1, ad6, adf, ad2, ad3, ad4, ad7, adb, ada, ad9, ade)
Digits(14) = Array(ae0, ae5, aec, aed, ae8, ae1, ae6, aef, ae2, ae3, ae4, ae7, aeb, aea, ae9, aee)
Digits(15) = Array(af0, af5, afc, afd, af8, af1, af6, aff, af2, af3, af4, af7, afb, afa, af9, afe)

Digits(16) = Array(b00, b05, b0c, b0d, b08, b01, b06, b0f, b02, b03, b04, b07, b0b, b0a, b09, b0e)
Digits(17) = Array(b10, b15, b1c, b1d, b18, b11, b16, b1f, b12, b13, b14, b17, b1b, b1a, b19, b1e)
Digits(18) = Array(b20, b25, b2c, b2d, b28, b21, b26, b2f, b22, b23, b24, b27, b2b, b2a, b29, b2e)
Digits(19) = Array(b30, b35, b3c, b3d, b38, b31, b36, b3f, b32, b33, b34, b37, b3b, b3a, b39, b3e)
Digits(20) = Array(b40, b45, b4c, b4d, b48, b41, b46, b4f, b42, b43, b44, b47, b4b, b4a, b49, b4e)
Digits(21) = Array(b50, b55, b5c, b5d, b58, b51, b56, b5f, b52, b53, b54, b57, b5b, b5a, b59, b5e)
Digits(22) = Array(b60, b65, b6c, b6d, b68, b61, b66, b6f, b62, b63, b64, b67, b6b, b6a, b69, b6e)
Digits(23) = Array(b70, b75, b7c, b7d, b78, b71, b76, b7f, b72, b73, b74, b77, b7b, b7a, b79, b7e)
Digits(24) = Array(b80, b85, b8c, b8d, b88, b81, b86, b8f, b82, b83, b84, b87, b8b, b8a, b89, b8e)
Digits(25) = Array(b90, b95, b9c, b9d, b98, b91, b96, b9f, b92, b93, b94, b97, b9b, b9a, b99, b9e)
Digits(26) = Array(ba0, ba5, bac, bad, ba8, ba1, ba6, baf, ba2, ba3, ba4, ba7, bab, baa, ba9, bae)
Digits(27) = Array(bb0, bb5, bbc, bbd, bb8, bb1, bb6, bbf, bb2, bb3, bb4, bb7, bbb, bba, bb9, bbe)
Digits(28) = Array(bc0, bc5, bcc, bcd, bc8, bc1, bc6, bcf, bc2, bc3, bc4, bc7, bcb, bca, bc9, bce)
Digits(29) = Array(bd0, bd5, bdc, bdd, bd8, bd1, bd6, bdf, bd2, bd3, bd4, bd7, bdb, bda, bd9, bde)
Digits(30) = Array(be0, be5, bec, bed, be8, be1, be6, bef, be2, be3, be4, be7, beb, bea, be9, bee)
Digits(31) = Array(bf0, bf5, bfc, bfd, bf8, bf1, bf6, bff, bf2, bf3, bf4, bf7, bfb, bfa, bf9, bfe)

Sub UpdateLeds
    Dim ChgLED, ii, jj, num, chg, stat, obj, b, x
    ChgLED = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
        For ii = 0 To UBound(chgLED)
            num = chgLED(ii, 0):chg = chgLED(ii, 1):stat = chgLED(ii, 2)
            For Each obj In Digits(num)
                If chg And 1 Then obj.State = stat And 1
                chg = chg \ 2:stat = stat \ 2
            Next
        Next
    End If
End Sub
' Lunelle - Taito 1982
'  IPD No. 4591 / 4 Players
' Playfield layout is identical to Williams' 1980 'Alien Poker'.
' VPX - version by JPSalas 2017, version 1.0.2

Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01550000", "Taito.vbs", 3.26

Dim bsTrough, bsLeftSaucer, bsRightSaucer, bsTopSaucer, dtbank, x

Const cGameName = "lunelle"

Const UseSolenoids = 1
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
        .SplashInfoLine = "Lunelle - Taito 1982" & vbNewLine & "VPX table by JPSalas v.1.0.0"
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
    vpmNudge.TiltSwitch = 30
    vpmNudge.Sensitivity = 3
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, Bumper4, LeftSlingshot)

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0, 1, 0, 0, 0, 0, 0, 0
        .InitKick BallRelease, 90, 4
        .InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .Balls = 1
    End With

    ' Drop targets
    set dtbank = new cvpmdroptarget
    dtbank.InitDrop Array(sw22, sw32, sw42, sw52, sw62), Array(22, 32, 42, 52, 62)
    dtbank.initsnd SoundFX("", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)

    ' Saucers
    Set bsLeftSaucer = New cvpmBallStack
    bsLeftSaucer.InitSaucer sw2, 2, 130, 15
    bsLeftSaucer.InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    bsLeftSaucer.KickForceVar = 6

    Set bsRightSaucer = New cvpmBallStack
    bsRightSaucer.InitSaucer sw3, 3, 0, 15
    bsRightSaucer.InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    bsRightSaucer.KickForceVar = 7

    Set bsTopSaucer = New cvpmBallStack
    bsTopSaucer.InitSaucer sw4, 4, 105, 6
    bsTopSaucer.InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    bsTopSaucer.KickForceVar = 4

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

    ' Turn on Gi
    GiOn
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
    If KeyCode = RightFlipperKey Then Controller.Switch(11) = 1
    If vpmKeyDown(keycode)Then Exit Sub
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If KeyCode = RightFlipperKey Then Controller.Switch(11) = 0
    If vpmKeyUp(keycode)Then Exit Sub
    If keycode = PlungerKey Then PlaySound "fx_plunger", 0, 1, 0.1, 0.25:Plunger.Fire
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
    vpmTimer.PulseSw 61
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

' Scoring rubbers
Dim Rub1, Rub2, Rub3, Rub4, Rub5

Sub sw15_Hit:PlaySound "fx_Rubber", 0, 1, pan(ActiveBall):vpmTimer.PulseSw 15:End Sub

Sub sw23_Hit:PlaySound "fx_Rubber", 0, 1, pan(ActiveBall):vpmTimer.PulseSw 23:End Sub

Sub sw53_Hit:PlaySound "fx_Rubber", 0, 1, pan(ActiveBall):vpmTimer.PulseSw 53:Rub1 = 1:sw53_Timer:End Sub
Sub sw53_Timer
    Select Case Rub1
        Case 1:RightSling2.Visible = 0:RightSling3.Visible = 1:sw53.TimerEnabled = 1
        Case 2:RightSling3.Visible = 0:RightSling4.Visible = 1
        Case 3:RightSling4.Visible = 0:RightSling2.Visible = 1:sw53.TimerEnabled = 0
    End Select
    Rub1 = Rub1 + 1
End Sub

Sub sw51_Hit:PlaySound "fx_Rubber", 0, 1, pan(ActiveBall):vpmTimer.PulseSw 51:Rub2 = 1:sw51_Timer:End Sub
Sub sw51_Timer
    Select Case Rub2
        Case 1:rubber2.Visible = 0:rubber3.Visible = 1:sw51.TimerEnabled = 1
        Case 2:rubber3.Visible = 0:rubber22.Visible = 1
        Case 3:rubber22.Visible = 0:rubber2.Visible = 1:sw51.TimerEnabled = 0
    End Select
    Rub2 = Rub2 + 1
End Sub

Sub sw33_Hit:PlaySound "fx_Rubber", 0, 1, pan(ActiveBall):vpmTimer.PulseSw 33:Rub3 = 1:sw33_Timer:End Sub
Sub sw33_Timer
    Select Case Rub3
        Case 1:rubber19.Visible = 0:rubber24.Visible = 1:sw33.TimerEnabled = 1
        Case 2:rubber24.Visible = 0:rubber25.Visible = 1
        Case 3:rubber25.Visible = 0:rubber19.Visible = 1:sw33.TimerEnabled = 0
    End Select
    Rub3 = Rub3 + 1
End Sub

Sub sw43_Hit:PlaySound "fx_Rubber", 0, 1, pan(ActiveBall):vpmTimer.PulseSw 43:Rub4 = 1:sw43_Timer:End Sub
Sub sw43_Timer
    Select Case Rub4
        Case 1:rubber18.Visible = 0:rubber27.Visible = 1:sw43.TimerEnabled = 1
        Case 2:rubber27.Visible = 0:rubber28.Visible = 1
        Case 3:rubber28.Visible = 0:rubber18.Visible = 1:sw43.TimerEnabled = 0
    End Select
    Rub4 = Rub4 + 1
End Sub

Sub sw63_Hit:PlaySound "fx_Rubber", 0, 1, pan(ActiveBall):vpmTimer.PulseSw 63:Rub5 = 1:sw63_Timer:End Sub
Sub sw63_Timer
    Select Case Rub5
        Case 1:rubber29.Visible = 0:rubber30.Visible = 1:sw63.TimerEnabled = 1
        Case 2:rubber30.Visible = 0:rubber31.Visible = 1
        Case 3:rubber31.Visible = 0:rubber29.Visible = 1:sw63.TimerEnabled = 0
    End Select
    Rub5 = Rub5 + 1
End Sub

' Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 24:PlaySound SoundFX("fx_bumper", DOFContactors), 0, 1, 0, -0.05:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 14:PlaySound SoundFX("fx_bumper", DOFContactors), 0, 1, 0, 0:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 34:PlaySound SoundFX("fx_bumper", DOFContactors), 0, 1, 0, 0.05:End Sub
Sub Bumper4_Hit:vpmTimer.PulseSw 73:PlaySound SoundFX("fx_bumper", DOFContactors), 0, 1, 0, 0.02:End Sub

' Drain & Sauders
Sub Drain_Hit:Playsound "fx_drain":bsTrough.AddBall Me:End Sub
Sub sw2_Hit:Playsound "fx_kicker_enter", 0, 1, 0, -0.1:bsLeftSaucer.AddBall 0:End Sub
Sub sw3_Hit:Playsound "fx_kicker_enter", 0, 1, 0, 0.1:bsRightSaucer.AddBall 0:End Sub
Sub sw4_Hit:Playsound "fx_kicker_enter", 0, 1, 0, 0:bsTopSaucer.AddBall 0:End Sub

' Rollovers
Sub sw21_Hit:Controller.Switch(21) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw21_UnHit:Controller.Switch(21) = 0:End Sub

Sub sw31_Hit:Controller.Switch(31) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw31_UnHit:Controller.Switch(31) = 0:End Sub

Sub sw41_Hit:Controller.Switch(41) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw41_UnHit:Controller.Switch(41) = 0:End Sub

Sub sw44_Hit:Controller.Switch(44) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw44_UnHit:Controller.Switch(44) = 0:End Sub

Sub sw54_Hit:Controller.Switch(54) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw54_UnHit:Controller.Switch(54) = 0:End Sub

Sub sw64_Hit:Controller.Switch(64) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw64_UnHit:Controller.Switch(64) = 0:End Sub

Sub sw74_Hit:Controller.Switch(74) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw74_UnHit:Controller.Switch(74) = 0:End Sub

' Droptargets
Sub sw22_Hit:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw32_Hit:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw42_Hit:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw52_Hit:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw62_Hit:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub

Sub sw22_Dropped:dtbank.Hit 1:End Sub
Sub sw32_Dropped:dtbank.Hit 2:End Sub
Sub sw42_Dropped:dtbank.Hit 3:End Sub
Sub sw52_Dropped:dtbank.Hit 4:End Sub
Sub sw62_Dropped:dtbank.Hit 5:End Sub

'Targets
Sub sw5_Hit:vpmTimer.PulseSw 5:PlaySound SoundFX("fx_target", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw12_Hit:vpmTimer.PulseSw 12:PlaySound SoundFX("fx_target", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw13_Hit:vpmTimer.PulseSw 13:PlaySound SoundFX("fx_target", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw71_Hit:vpmTimer.PulseSw 71:PlaySound SoundFX("fx_target", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub

'*********
'Solenoids
'*********

SolCallback(1) = "bsTrough.SolOut"
SolCallback(2) = "SolLeftSaucer"
SolCallback(3) = "SolRightSaucer"
SolCallback(4) = "SolTopSaucer"
SolCallback(5) = "dtbank.SolDropUp"
SolCallback(7) = "dtbank.SolUnhit 1,"
SolCallback(8) = "dtbank.SolUnHit 2,"
SolCallback(9) = "dtbank.SolUnHit 3,"
SolCallback(10) = "dtbank.SolUnhit 4,"
SolCallback(11) = "dtbank.SolUnHit 5,"

SolCallback(13) = "GiEffect" ' jp: unknown solenoid, it fires sometimes, but it can't be a knocker
SolCallback(6) = "GiEffect" '6=Strobe jp:I guess not in use in this table
SolCallback(17) = "SolGi"   '17=relay jp:actually I haven't a clue what this those either :)
SolCallback(18) = "vpmNudge.SolGameOn"

Sub SolGi(enabled)
    If enabled Then
        GiOff
    Else
        GiOn
    End If
End Sub

Dim sw2Pos, sw3Pos, sw4Pos

Sub SolLeftSaucer(enabled)
    If enabled Then
        bsLeftSaucer.ExitSol_On
        sw2pos = 1:sw2.TimerEnabled = 1
    End If
End Sub

Sub sw2_Timer
    Select Case sw2Pos
        Case 1:sw2p.rotx = -30
        Case 2:sw2p.rotx = -20
        Case 3:sw2p.rotx = -10
        Case 4:sw2p.rotx = 0:sw2.TimerEnabled = 0
    End select
    sw2Pos = sw2Pos + 1
End Sub

Sub SolRightSaucer(enabled)
    If enabled Then
        bsRightSaucer.ExitSol_On
        sw3pos = 1:sw3.TimerEnabled = 1
    End If
End Sub

Sub sw3_Timer
    Select Case sw3Pos
        Case 1:sw3p.rotx = 30
        Case 2:sw3p.rotx = 20
        Case 3:sw3p.rotx = 10
        Case 4:sw3p.rotx = 0:sw3.TimerEnabled = 0
    End select
    sw3Pos = sw3Pos + 1
End Sub

Sub SolTopSaucer(enabled)
    If enabled Then
        bsTopSaucer.ExitSol_On
        sw4pos = 1:sw4.TimerEnabled = 1
    End If
End Sub

Sub sw4_Timer
    Select Case sw4Pos
        Case 1:sw4p.roty = 30
        Case 2:sw4p.roty = 20
        Case 3:sw4p.roty = 10
        Case 4:sw4p.roty = 0:sw4.TimerEnabled = 0
    End select
    sw4Pos = sw4Pos + 1
End Sub

'**************
' Flipper Subs
'**************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySound SoundFX("fx_flipperup", DOFFlippers), 0, 1, -0.1, 0.1
        LeftFlipper.RotateToEnd
    Else
        PlaySound SoundFX("fx_flipperdown", DOFFlippers), 0, 1, -0.1, 0.1
        LeftFlipper.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySound SoundFX("fx_flipperup", DOFFlippers), 0, 1, 0.1, 0.1
        RightFlipper.RotateToEnd
        RightFlipper1.RotateToEnd
    Else
        PlaySound SoundFX("fx_flipperdown", DOFFlippers), 0, 1, 0.1, 0.1
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

Sub GiEffect(enabled)
    If enabled Then
        For each x in aGiLights
            x.Duration 2, 1000, 1
        Next
    End If
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
    'GIUpdate
    RollingUpdate
End Sub

Sub UpdateLamps()
    NFadeL 0, l0
    NFadeL 1, l1
    NFadeL 10, l10
    NFadeL 100, l100
    NFadeL 101, l101
    NFadeL 102, l102
    NFadeL 103, l103
    NFadeL 109, l109
    NFadeLm 110, l110a
    NFadeLm 110, l110b
    NFadeL 110, l110c
    NFadeLm 111, l111a
    NFadeLm 111, l111b
    NFadeL 111, l111c
    NFadeLm 112, l112a
    NFadeLm 112, l112b
    NFadeL 112, l112c
    NFadeL 113, l113
    NFadeLm 119, L119a
    NFadeLm 119, L119b
    NFadeL 119, L119c
    NFadeL 11, l11
    NFadeL 12, l12
    NFadeL 120, l120
    NFadeL 121, l121
    NFadeL 122, l122
    'NFadeL 123, l123
    NFadeL 129, l129
    NFadeL 130, l130
    NFadeL 131, l131
    NFadeL 132, l132
    NFadeL 133, l133
    NFadeL 139, l139
    NFadeLm 143, l143
    Flash 143, l143a
    NFadeL 153, l153
    NFadeL 2, l2
    NFadeL 20, l20
    NFadeL 21, l21
    NFadeL 22, l22
    NFadeL 30, l30
    NFadeL 31, l31
    NFadeL 32, l32
    NFadeLm 40, l40
    NFadeLm 40, l40a
    Flashm 40, l40b
    Flash 40, l40c
    NFadeL 41, l41
    NFadeL 42, l42
    NFadeL 50, l50
    NFadeL 51, l51
    NFadeL 52, l52
    NFadeLm 60, l60
    Flash 60, l60a
    NFadeLm 61, l61
    Flash 61, l61a
    NFadeLm 62, l62
    Flash 62, l62a
    NFadeL 70, l70
    NFadeL 71, l71
    NFadeL 72, l72
    NFadeL 79, l79
    NFadeLm 80, l80
    Flash 80, l80a
    'NFadeL 81, l81
    NFadeL 82, l82
    NFadeL 83, l83
    NFadeL 89, l89
    NFadeL 90, l90
    NFadeL 91, l91
    NFadeL 92, l92
    NFadeL 93, l93
    NFadeL 99, l99

    'backdrop lights
    If VarHidden Then
        NFadeL 139, l139
        NFadeL 140, l140
        NFadeL 141, l141
        NFadeL 142, l142
        NFadeL 149, l149
        NFadeL 150, l150
        NFadeL 151, l151
        NFadeL 152, l152
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

Set Digits(6) = b0
Set Digits(7) = b1
Set Digits(8) = b2
Set Digits(9) = b3
Set Digits(10) = b4
Set Digits(11) = b5

Set Digits(12) = c0
Set Digits(13) = c1
Set Digits(14) = c2
Set Digits(15) = c3
Set Digits(16) = c4
Set Digits(17) = c5

Set Digits(18) = d0
Set Digits(19) = d1
Set Digits(20) = d2
Set Digits(21) = d3
Set Digits(22) = d4
Set Digits(23) = d5

Set Digits(24) = e0
Set Digits(25) = e1

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

Sub aMetals_Hit(idx):PlaySound "fx_MetalHit2", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aRubber_Bands_Hit(idx):PlaySound "fx_rubber_band", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aRubber_Posts_Hit(idx):PlaySound "fx_rubber_post", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aRubber_Pins_Hit(idx):PlaySound "fx_rubber_pin", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aWoods_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub

' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / table1.width-1
    If tmp > 0 Then
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

Function AudioFade(ball) 'only on VPX 10.4 and newer
    Dim tmp
    tmp = ball.y * 2 / Table1.height-1
    If tmp > 0 Then
        AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10))
    End If
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
    For b = lob to UBound(BOT)
        If BallVel(BOT(b)) > 1 Then
            If BOT(b).z < 30 Then
                ballpitch = Pitch(BOT(b))
            Else
                ballpitch = Pitch(BOT(b)) + 15000 'increase the pitch on a ramp or elevated surface
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
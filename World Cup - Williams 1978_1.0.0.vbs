' World Cup - Williams
' IPD No. 2810 / March 06, 1978 / 4 Players
' VPX - version by JPSalas 2017, version 1.0.0

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' No special SSF tweaks yet.

Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01550000", "s4.vbs", 3.26

Dim bsTrough, bsLeftSaucer, bsRightSaucer, bsLeftKick, bsRightKick, x

Dim chimes
chimes = FALSE 'change to true if you prefer the chimes instead of the "new" sounds from 1978

Const cGameName = "wldcp_l1"

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
        .SplashInfoLine = "World Cup - Williams 1978" & vbNewLine & "VPX table by JPSalas v.1.0.0"
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
    vpmNudge.TiltSwitch = 15
    vpmNudge.Sensitivity = 3
    vpmNudge.TiltObj = Array(Bumper1, Bumper2)

    ' Trough
    Set bsTrough = New cvpmBallStack
    bsTrough.InitNoTrough BallRelease, 14, 90, 5
    bsTrough.InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    bsTrough.Balls = 1

    ' Saucers
    Set bsLeftSaucer = New cvpmBallStack
    bsLeftSaucer.InitSaucer sw9, 9, 154, 18
    bsLeftSaucer.InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    bsLeftSaucer.KickAngleVar = 2
    bsLeftSaucer.KickForceVar = 1.5

    Set bsRightSaucer = New cvpmBallStack
    bsRightSaucer.InitSaucer sw20, 20, 205, 18
    bsRightSaucer.InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    bsRightSaucer.KickAngleVar = 2
    bsRightSaucer.KickForceVar = 1.5

    Set bsLeftKick = New cvpmBallStack
    bsLeftKick.InitSaucer sw13, 13, 35, 26.5
    bsLeftKick.InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    bsLeftKick.KickAngleVar = 2
    bsLeftKick.KickForceVar = 1.5

    Set bsRightKick = New cvpmBallStack
    bsRightKick.InitSaucer sw16, 16, 325, 26.5
    bsRightKick.InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    bsRightKick.KickAngleVar = 2
    bsRightKick.KickForceVar = 1.5

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1
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
    If vpmKeyDown(keycode)Then Exit Sub
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If vpmKeyUp(keycode)Then Exit Sub
    If keycode = PlungerKey Then PlaySound "fx_plunger", 0, 1, 0.1, 0.25:Plunger.Fire
End Sub

'*********
' Switches
'*********

' Scoring rubbers
Dim Rub1, Rub2, Rub3, Rub4, Rub5, Rub6, Rub7

Sub sw10_Hit:PlaySound "fx_Rubber", 0, 1, pan(ActiveBall):vpmTimer.PulseSw 10:Rub1 = 1:sw10_Timer:End Sub

Sub sw10_Timer
    Select Case Rub1
        Case 1:rubber16.Visible = 0:rubber26.Visible = 1:sw10.TimerEnabled = 1
        Case 2:rubber26.Visible = 0:rubber27.Visible = 1
        Case 3:rubber27.Visible = 0:rubber16.Visible = 1:sw10.TimerEnabled = 0
    End Select
    Rub1 = Rub1 + 1
End Sub

Sub sw12_Hit:PlaySound "fx_Rubber", 0, 1, pan(ActiveBall):vpmTimer.PulseSw 12:Rub2 = 1:sw12_Timer:End Sub
Sub sw12_Timer
    Select Case Rub2
        Case 1:rubber35.Visible = 0:rubber36.Visible = 1:sw12.TimerEnabled = 1
        Case 2:rubber36.Visible = 0:rubber37.Visible = 1
        Case 3:rubber37.Visible = 0:rubber35.Visible = 1:sw12.TimerEnabled = 0
    End Select
    Rub2 = Rub2 + 1
End Sub

Sub sw17_Hit:PlaySound "fx_Rubber", 0, 1, pan(ActiveBall):vpmTimer.PulseSw 17:Rub3 = 1:sw17_Timer:End Sub
Sub sw17_Timer
    Select Case Rub3
        Case 1:rubber29.Visible = 0:rubber6.Visible = 1:sw17.TimerEnabled = 1
        Case 2:rubber6.Visible = 0:rubber7.Visible = 1
        Case 3:rubber7.Visible = 0:rubber29.Visible = 1:sw17.TimerEnabled = 0
    End Select
    Rub3 = Rub3 + 1
End Sub

Sub sw19_Hit:PlaySound "fx_Rubber", 0, 1, pan(ActiveBall):vpmTimer.PulseSw 19:Rub4 = 1:sw19_Timer:End Sub
Sub sw19_Timer
    Select Case Rub4
        Case 1:rubber18.Visible = 0:rubber28.Visible = 1:sw19.TimerEnabled = 1
        Case 2:rubber28.Visible = 0:rubber30.Visible = 1
        Case 3:rubber30.Visible = 0:rubber18.Visible = 1:sw19.TimerEnabled = 0
    End Select
    Rub4 = Rub4 + 1
End Sub

Sub sw29_Hit:PlaySound "fx_Rubber", 0, 1, pan(ActiveBall):vpmTimer.PulseSw 29:Rub5 = 1:sw29_Timer:End Sub
Sub sw29_Timer
    Select Case Rub5
        Case 1:rubber22.Visible = 0:rubber31.Visible = 1:sw29.TimerEnabled = 1
        Case 2:rubber31.Visible = 0:rubber32.Visible = 1
        Case 3:rubber32.Visible = 0:rubber22.Visible = 1:sw29.TimerEnabled = 0
    End Select
    Rub5 = Rub5 + 1
End Sub

Sub sw33_Hit:PlaySound "fx_Rubber", 0, 1, pan(ActiveBall):vpmTimer.PulseSw 33:Rub6 = 1:sw33_Timer:End Sub
Sub sw33_Timer
    Select Case Rub6
        Case 1:rubber21.Visible = 0:rubber33.Visible = 1:sw33.TimerEnabled = 1
        Case 2:rubber33.Visible = 0:rubber34.Visible = 1
        Case 3:rubber34.Visible = 0:rubber21.Visible = 1:sw33.TimerEnabled = 0
    End Select
    Rub6 = Rub6 + 1
End Sub

Sub sw36_Hit:PlaySound "fx_Rubber", 0, 1, pan(ActiveBall):vpmTimer.PulseSw 36:Rub7 = 1:sw36_Timer:End Sub
Sub sw36_Timer
    Select Case Rub7
        Case 1:rubber20.Visible = 0:rubber24.Visible = 1:sw36.TimerEnabled = 1
        Case 2:rubber24.Visible = 0:rubber25.Visible = 1
        Case 3:rubber25.Visible = 0:rubber20.Visible = 1:sw36.TimerEnabled = 0
    End Select
    Rub7 = Rub7 + 1
End Sub

' Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 37:PlaySound SoundFX("fx_bumper", DOFContactors), 0, 1, 0, -0.05:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 38:PlaySound SoundFX("fx_bumper", DOFContactors), 0, 1, 0, 0.02:End Sub

' Drain & Saucers
Sub Drain_Hit:Playsound "fx_drain":bsTrough.AddBall Me:End Sub
Sub sw9_Hit:Playsound "fx_kicker_enter", 0, 1, 0, -0.1:bsLeftSaucer.AddBall 0:End Sub
Sub sw13_Hit:Playsound "fx_kicker_enter", 0, 1, 0, -0.1:bsLeftKick.AddBall 0:End Sub
Sub sw16_Hit:Playsound "fx_kicker_enter", 0, 1, 0, -0.1:bsRightKick.AddBall 0:End Sub
Sub sw20_Hit:Playsound "fx_kicker_enter", 0, 1, 0, -0.1:bsRightSaucer.AddBall 0:End Sub

' Rollovers
Sub sw11_Hit:Controller.Switch(11) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw11_UnHit:Controller.Switch(11) = 0:End Sub

Sub sw18_Hit:Controller.Switch(18) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw18_UnHit:Controller.Switch(18) = 0:End Sub

Sub sw21_Hit:Controller.Switch(21) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw21_UnHit:Controller.Switch(21) = 0:End Sub

Sub sw22_Hit:Controller.Switch(22) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw22_UnHit:Controller.Switch(22) = 0:End Sub

Sub sw23_Hit:Controller.Switch(23) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw23_UnHit:Controller.Switch(23) = 0:End Sub

Sub sw24_Hit:Controller.Switch(24) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw24_UnHit:Controller.Switch(24) = 0:End Sub

Sub sw21a_Hit:Controller.Switch(21) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw21a_UnHit:Controller.Switch(21) = 0:End Sub

Sub sw22a_Hit:Controller.Switch(22) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw22a_UnHit:Controller.Switch(22) = 0:End Sub

Sub sw23a_Hit:Controller.Switch(23) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw23a_UnHit:Controller.Switch(23) = 0:End Sub

Sub sw24a_Hit:Controller.Switch(24) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw24a_UnHit:Controller.Switch(24) = 0:End Sub

Sub sw26_Hit:Controller.Switch(26) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw26_UnHit:Controller.Switch(26) = 0:End Sub

Sub sw28_Hit:Controller.Switch(28) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw28_UnHit:Controller.Switch(28) = 0:End Sub

Sub sw30_Hit:Controller.Switch(30) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw30_UnHit:Controller.Switch(30) = 0:End Sub

Sub sw31_Hit:Controller.Switch(31) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw31_UnHit:Controller.Switch(31) = 0:End Sub

Sub sw32_Hit:Controller.Switch(32) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw32_UnHit:Controller.Switch(32) = 0:End Sub

Sub sw35_Hit:Controller.Switch(35) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw35_UnHit:Controller.Switch(35) = 0:End Sub

'Targets
Sub sw27_Hit
    vpmTimer.PulseSw 27
    PlaySound SoundFX("fx_target", DOFDropTargets), 0, 1, pan(ActiveBall)
End Sub

Sub sw34_Hit
    vpmTimer.PulseSw 34
    PlaySound SoundFX("fx_target", DOFDropTargets), 0, 1, pan(ActiveBall)
End Sub

'Spinner

Sub sw25_Spin:vpmTimer.PulseSw 25:Playsound "fx_spinner", 0, 1, 0, -0.01:End Sub

'*********
'Solenoids
'*********

SolCallback(1) = "SolLeftSaucer"
SolCallback(2) = "SolRightSaucer"
SolCallback(3) = "SolLeftKick"
SolCallback(4) = "SolRightKick"
SolCallback(5) = "bsTrough.SolOut"
'6 not used
SolCallback(7) = "vpmSolSound ""Wc_Chime_Sol7""," 'Bell
'8 not used
If chimes Then
    SolCallback(9) = "vpmSolSound ""Wc_Chime_Sol9"","   'chimes 10 points
    SolCallback(10) = "vpmSolSound ""Wc_Chime_Sol10""," 'chimes 100 points
    SolCallback(11) = "vpmSolSound ""Wc_Chime_Sol11""," 'chimes 1000 points
    SolCallback(12) = "vpmSolSound ""Wc_Chime_Sol12""," 'chimes 10000 points
Else
    SolCallback(9) = "vpmSolSound ""Wc_Synth_Sol9"","   'chimes 10 points
    SolCallback(10) = "vpmSolSound ""Wc_Synth_Sol10""," 'chimes 100 points
    SolCallback(11) = "vpmSolSound ""Wc_Synth_Sol11""," 'chimes 1000 points
    SolCallback(12) = "vpmSolSound ""Wc_Synth_Sol12""," 'chimes 10000 points
End If
SolCallback(13) = "vpmSolSound ""fx_solenoidon"","      'sound alternator
SolCallback(14) = "vpmSolSound ""fx_knocker"","
SolCallback(15) = "vpmSolSound ""Wc_Chime_Sol15"","     'Tilt
' 16 coin lockout

Dim kickerPos ' only one ball so we only need one variable

Sub SolLeftSaucer(enabled)
    If enabled Then
        bsLeftSaucer.ExitSol_On
        kickerPos = 1:sw20_Timer
    End If
End Sub

Sub sw20_Timer
    Select Case kickerPos
        Case 1:sw20p.rotx = 30:sw20.TimerEnabled = 1
        Case 2:sw20p.rotx = 20
        Case 3:sw20p.rotx = 10
        Case 4:sw20p.rotx = 0:sw20.TimerEnabled = 0
    End select
    kickerPos = kickerPos + 1
End Sub

Sub SolRightSaucer(enabled)
    If enabled Then
        bsRightSaucer.ExitSol_On
        kickerPos = 1:sw9_Timer
    End If
End Sub

Sub sw9_Timer
    Select Case kickerPos
        Case 1:sw9p.rotx = 30:sw9.TimerEnabled = 1
        Case 2:sw9p.rotx = 20
        Case 3:sw9p.rotx = 10
        Case 4:sw9p.rotx = 0:sw9.TimerEnabled = 0
    End select
    kickerPos = kickerPos + 1
End Sub

Sub SolLeftKick(enabled)
    If enabled Then
        bsLeftKick.ExitSol_On
        kickerPos = 1:sw13_Timer
    End If
End Sub

Sub sw13_Timer
    Select Case kickerPos
        Case 1:sw13p.rotx = 26:sw13.TimerEnabled = 1
        Case 2:sw13p.rotx = 14
        Case 3:sw13p.rotx = 2
        Case 4:sw13p.rotx = -15:sw13.TimerEnabled = 0
    End select
    kickerPos = kickerPos + 1
End Sub

Sub SolRightKick(enabled)
    If enabled Then
        bsRightKick.ExitSol_On
        kickerPos = 1:sw16_Timer
    End If
End Sub

Sub sw16_Timer
    Select Case kickerPos
        Case 1:sw16p.rotx = 26:sw16.TimerEnabled = 1
        Case 2:sw16p.rotx = 14
        Case 3:sw16p.rotx = 2
        Case 4:sw16p.rotx = -15:sw16.TimerEnabled = 0
    End select
    kickerPos = kickerPos + 1
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
    Else
        PlaySound SoundFX("fx_flipperdown", DOFFlippers), 0, 1, 0.1, 0.1
        RightFlipper.RotateToStart
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
    NFadeLm 1, li1a
    NFadeL 1, li1
    NFadeLm 2, li2a
    NFadeL 2, li2
    NFadeLm 3, li3a
    NFadeL 3, li3
    NFadeLm 4, li4a
    NFadeL 4, li4
    NFadeLm 5, li5a
    NFadeL 5, li5
    NFadeL 6, li6
    NFadeL 7, li7
    NFadeL 8, li8
    NFadeL 9, li9
    NFadeL 10, li10
    NFadeL 11, li11
    NFadeL 12, li12
    NFadeL 13, li13
    NFadeL 14, li14
    NFadeL 15, li15
    NFadeL 16, li16
    NFadeL 17, li17
    NFadeL 18, li18
    NFadeL 19, li19
    NFadeL 20, li20
    NFadeL 21, li21
    NFadeL 22, li22
    NFadeL 23, li23
    NFadeL 24, li24
    NFadeL 25, li25
    NFadeL 26, li26
    NFadeL 27, li27
    NFadeL 28, li28
    NFadeL 29, li29
    NFadeL 30, li30
    NFadeL 31, li31
    NFadeL 32, li32
    NFadeL 33, li33
    NFadeL 34, li34
    NFadeL 35, li35
    NFadeL 36, li36
    NFadeL 37, li37
    NFadeL 38, li38
    NFadeL 39, li39
    NFadeL 40, li40
    NFadeL 41, li41
    NFadeL 42, li42
    NFadeL 43, li43
    NFadeL 44, li44
    NFadeL 45, li45
    NFadeL 46, li46
    NFadeL 47, li47
    NFadeL 48, li48
    NFadeL 49, li49
    NFadeL 56, li56

    'backdrop lights
    If VarHidden Then
        NFadeL 50, li50
        NFadeL 51, li51
        NFadeL 52, li52
        NFadeL 53, li53

        NFadeT 54, li54, "MATCH"
        NFadeT 55, li55, "BALL IN PLAY"
        NFadeL 57, li57
        NFadeL 58, li58
        NFadeL 59, li59
        NFadeL 60, li60
        NFadeT 61, li61, "TILT"
        NFadeT 62, li62, "GAME OVER"
        NFadeT 63, li63, "SAME PLAYER SHOOTS AGAIN"
        NFadeT 64, li64, "HIGH SCORE"
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

'Assign 6-digit output to reels
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
Set Digits(26) = e2
Set Digits(27) = e3

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


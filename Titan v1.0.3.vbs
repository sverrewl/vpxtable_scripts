' Titan - Taito 1981
' IPD No. 5651 / 4 Players
' VPX - version by JPSalas 2018, version 1.0.3
Option Explicit
Randomize


' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Thalamus 2018-11-01 : Improved directional sounds
' !! NOTE : Table not verified yet !!

' Options
' Volume devided by - lower gets higher sound

Const VolDiv = 2000    ' Lower number, louder ballrolling/collition sound
Const VolCol = 10      ' Ball collition divider ( voldiv/volcol )

' The rest of the values are multipliers
'
'  .5 = lower volume
' 1.5 = higher volume

Const VolBump   = 2    ' Bumpers volume.
Const VolRol    = 1    ' Rollovers volume.
Const VolGates  = 1    ' Gates volume.
Const VolMetal  = 1    ' Metals volume.
Const VolRB     = 1    ' Rubber bands volume.
Const VolRH     = 1    ' Rubber hits volume.
Const VolPo     = 1    ' Rubber posts volume.
Const VolPi     = 1    ' Rubber pins volume.
Const VolPlast  = 1    ' Plastics volume.
Const VolTarg   = 1    ' Targets volume.
Const VolWood   = 1    ' Woods volume.
Const VolKick   = 1    ' Kicker volume.
Const VolSpin   = 1.5  ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01550000", "Taito.vbs", 3.26

Dim bsTrough, dtbank1, dtbank2, bsLeftSaucer, bsRightSaucer, x

Const cGameName = "Titan"

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
        .SplashInfoLine = "Titan - Taito 1981" & vbNewLine & "VPX table by JPSalas v.1.0.3"
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
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingshot)

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0, 1, 0, 0, 0, 0, 0, 0
        .InitKick BallRelease, 90, 4
        .InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .Balls = 1
    End With

    ' Saucers
    Set bsLeftSaucer = New cvpmBallStack
    bsLeftSaucer.InitSaucer sw2, 2, 70, 15
    bsLeftSaucer.InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    bsLeftSaucer.KickForceVar = 6

    Set bsRightSaucer = New cvpmBallStack
    bsRightSaucer.InitSaucer sw3, 3, 180, 10
    bsRightSaucer.InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    bsRightSaucer.KickForceVar = 6

    ' Drop targets
    set dtbank1 = new cvpmdroptarget
    dtbank1.InitDrop Array(sw5, sw15, sw25), Array(5, 15, 25)
    dtbank1.initsnd SoundFX("", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)

    set dtbank2 = new cvpmdroptarget
    dtbank2.InitDrop Array(sw4, sw14, sw24, sw34, sw44), Array(4, 14, 24, 34, 44)
    dtbank2.initsnd SoundFX("", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)

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
    If keycode = PlungerKey Then PlaySoundAtVol "fx_PlungerPull",Plunger,1:Plunger.Pullback
    if keycode = rightflipperkey then controller.switch(75) = 1 'Mudas as lampadas de cima
    if keycode = leftflipperkey then controller.switch(74) = 1
    If vpmKeyDown(keycode)Then Exit Sub
End Sub

Sub table1_KeyUp(ByVal Keycode)
    if keycode = rightflipperkey then controller.switch(75) = 0
    if keycode = leftflipperkey then controller.switch(74) = 0
    If vpmKeyUp(keycode)Then Exit Sub
    If keycode = PlungerKey Then PlaySoundAtVol "fx_plunger",Plunger,1:Plunger.Fire
End Sub

'*********
' Switches
'*********

' Slings
Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    PlaySoundAtVol SoundFX("fx_slingshot", DOFContactors), Lemk, 1
    DOF 101, DOFPulse
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 54
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
    PlaySoundAtVol SoundFX("fx_slingshot", DOFContactors), Remk, 1
    DOF 102, DOFPulse
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

' Scoring rubbers
Dim Rub1, Rub2, Rub3, Rub4

Sub sw41_Hit:PlaySoundAtVol "fx_Rubber",ActiveBall,VolRB:vpmTimer.PulseSw 41:Rub1 = 1:sw41_Timer:End Sub

Sub sw41_Timer
    Select Case Rub1
        Case 1:r4.Visible = 0:r6.Visible = 1:sw41.TimerEnabled = 1
        Case 2:r6.Visible = 0:r7.Visible = 1
        Case 3:r7.Visible = 0:r4.Visible = 1:sw41.TimerEnabled = 0
    End Select
    Rub1 = Rub1 + 1
End Sub

Sub sw51_Hit:PlaySoundAtVol "fx_Rubber",ActiveBall,VolRB:vpmTimer.PulseSw 51:Rub2 = 1:sw51_Timer:End Sub
Sub sw51_Timer
    Select Case Rub2
        Case 1:r1.Visible = 0:r8.Visible = 1:sw51.TimerEnabled = 1
        Case 2:r8.Visible = 0:r9.Visible = 1
        Case 3:r9.Visible = 0:r1.Visible = 1:sw51.TimerEnabled = 0
    End Select
    Rub2 = Rub2 + 1
End Sub

Sub sw61_Hit:PlaySoundAtVol "fx_Rubber",ActiveBall,VolRB:vpmTimer.PulseSw 61:Rub3 = 1:sw61_Timer:End Sub
Sub sw61_Timer
    Select Case Rub3
        Case 1:r2.Visible = 0:r10.Visible = 1:sw61.TimerEnabled = 1
        Case 2:r10.Visible = 0:r11.Visible = 1
        Case 3:r11.Visible = 0:r2.Visible = 1:sw61.TimerEnabled = 0
    End Select
    Rub3 = Rub3 + 1
End Sub

Sub sw71_Hit:PlaySoundAtVol "fx_Rubber",ActiveBall,VolRB:vpmTimer.PulseSw 71:Rub4 = 1:sw71_Timer:End Sub
Sub sw71_Timer
    Select Case Rub4
        Case 1:r12.Visible = 0:r13.Visible = 1:sw71.TimerEnabled = 1
        Case 2:r13.Visible = 0:r14.Visible = 1
        Case 3:r14.Visible = 0:r12.Visible = 1:sw71.TimerEnabled = 0
    End Select
    Rub4 = Rub4 + 1
End Sub

' Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 52:PlaySoundAtVol SoundFX("fx_bumper", DOFContactors),Bumper1,VolBump:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 62:PlaySoundAtVol SoundFX("fx_bumper", DOFContactors),Bumper2,VolBump:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 72:PlaySoundAtVol SoundFX("fx_bumper", DOFContactors),Bumper3,VolBump:End Sub

' Drain & Saucers
Sub Drain_Hit:PlaysoundAtVol "fx_drain",drain, 1:bsTrough.AddBall Me:End Sub
Sub sw2_Hit::PlaySoundAtVol "fx_kicker_enter",sw2,VolKick:bsLeftSaucer.AddBall 0:End Sub
Sub sw3_Hit::PlaySoundAtVol "fx_kicker_enter",sw3,VolKick:bsRightSaucer.AddBall 0:End Sub

' Rollovers
Sub sw12_Hit:Controller.Switch(12) = 1:PlaySoundAtVol "fx_sensor",ActiveBall,1:End Sub
Sub sw12_UnHit:Controller.Switch(12) = 0:End Sub

Sub sw22_Hit:Controller.Switch(22) = 1:PlaySoundAtVol "fx_sensor",ActiveBall,1:End Sub
Sub sw22_UnHit:Controller.Switch(22) = 0:End Sub

Sub sw32_Hit:Controller.Switch(32) = 1:PlaySoundAtVol "fx_sensor",ActiveBall,1:End Sub
Sub sw32_UnHit:Controller.Switch(32) = 0:End Sub

Sub sw42_Hit:Controller.Switch(42) = 1:PlaySoundAtVol "fx_sensor",ActiveBall,1:End Sub
Sub sw42_UnHit:Controller.Switch(42) = 0:End Sub

Sub sw13_Hit:Controller.Switch(13) = 1:PlaySoundAtVol "fx_sensor",ActiveBall,1:End Sub
Sub sw13_UnHit:Controller.Switch(13) = 0:End Sub

Sub sw23_Hit:Controller.Switch(23) = 1:PlaySoundAtVol "fx_sensor",ActiveBall,1:End Sub
Sub sw23_UnHit:Controller.Switch(23) = 0:End Sub

Sub sw33_Hit:Controller.Switch(33) = 1:PlaySoundAtVol "fx_sensor",ActiveBall,1:End Sub
Sub sw33_UnHit:Controller.Switch(33) = 0:End Sub

Sub sw53_Hit:Controller.Switch(53) = 1:PlaySoundAtVol "fx_sensor",ActiveBall,1:End Sub
Sub sw53_UnHit:Controller.Switch(53) = 0:End Sub

' Droptargets
Sub sw5_Hit:PlaySoundAtVol SoundFX("fx_droptarget",  DOFDropTargets),ActiveBall,VolTarg:End Sub
Sub sw15_Hit:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets),ActiveBall,VolTarg:End Sub
Sub sw25_Hit:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets),ActiveBall,VolTarg:End Sub
Sub sw4_Hit:PlaySoundAtVol SoundFX("fx_droptarget",  DOFDropTargets),ActiveBall,VolTarg:End Sub
Sub sw14_Hit:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets),ActiveBall,VolTarg:End Sub
Sub sw24_Hit:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets),ActiveBall,VolTarg:End Sub
Sub sw34_Hit:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets),ActiveBall,VolTarg:End Sub
Sub sw44_Hit:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets),ActiveBall,VolTarg:End Sub

Sub sw5_Dropped:dtbank1.Hit 1:End Sub
Sub sw15_Dropped:dtbank1.Hit 2:End Sub
Sub sw25_Dropped:dtbank1.Hit 3:End Sub
Sub sw4_Dropped:dtbank2.Hit 1:End Sub
Sub sw14_Dropped:dtbank2.Hit 2:End Sub
Sub sw24_Dropped:dtbank2.Hit 3:End Sub
Sub sw34_Dropped:dtbank2.Hit 4:End Sub
Sub sw44_Dropped:dtbank2.Hit 5:End Sub

'Targets
Sub sw35_Hit:vpmTimer.PulseSw 35:PlaySoundAtVol SoundFX("fx_target", DOFDropTargets),ActiveBall,VolTarg:End Sub
Sub sw45_Hit:vpmTimer.PulseSw 45:PlaySoundAtVol SoundFX("fx_target", DOFDropTargets),ActiveBall,VolTarg:End Sub
Sub sw55_Hit:vpmTimer.PulseSw 55:PlaySoundAtVol SoundFX("fx_target", DOFDropTargets),ActiveBall,VolTarg:End Sub

'*********
'Solenoids
'*********

SolCallback(1) = "bsTrough.SolOut"
SolCallback(2) = "bsLeftSaucer.SolOut"
SolCallback(3) = "bsRightSaucer.SolOut"
SolCallback(4) = "Sol4" 'backdrop TITAN lights, but I added a Gi effect too.
SolCallback(5) = "dtbank1.SolDropUp"
SolCallback(6) = "dtbank2.SolDropUp"

SolCallback(17) = "SolGi" '17=relay
SolCallback(18) = "vpmNudge.SolGameOn"

Sub SolGi(enabled)
    If enabled Then
        GiOff
    Else
        GiOn
    End If
End Sub

Sub Sol4(enabled)
    Dim a
    a = ABS(enabled)
    'titan flashers
    If VarHidden Then
        t1.State = a
        t2.State = a
        t3.State = a
        t4.State = a
        t5.State = a
  End If
    For each x in aGiLights
        x.state = ABS(a - 1)
    Next
End Sub

'**************
' Flipper Subs
'**************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFX("fx_flipperup", DOFFlippers), LeftFlipper, VolFlip
        LeftFlipper.RotateToEnd
    Else
        PlaySoundAtVol SoundFX("fx_flipperdown", DOFFlippers), LeftFlipper, VolFlip
        LeftFlipper.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFX("fx_flipperup", DOFFlippers), RightFlipper, VolFlip
        RightFlipper.RotateToEnd
    Else
        PlaySoundAtVol SoundFX("fx_flipperdown", DOFFlippers), RightFlipper, VolFlip
        RightFlipper.RotateToStart
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
    GIUpdate
    RollingUpdate
End Sub

Sub UpdateLamps()
    NFadeL 0, li0
    NFadeL 1, li1
    NFadeL 10, li10
    NFadeL 100, li100
    NFadeLm 101, li101
    Flash 101, li101a
    NFadeLm 102, li102
    Flash 102, li102a
    NFadeL 103, li103
    NFadeL 109, li109
    NFadeL 11, li11
    NFadeL 110, li110
    NFadeL 111, li111
    NFadeL 112, li112
    NFadeL 113, li113
    NFadeL 119, li119
    NFadeL 12, li12
    NFadeL 120, li120
    NFadeL 121, li121
    NFadeL 122, li122
    NFadeL 123, li123
    NFadeL 129, li129
    NFadeL 130, li130
    NFadeL 131, li131
    NFadeL 132, li132
    NFadeL 133, li133
    NFadeL 143, li143
    NFadeL 153, li153
    NFadeL 2, li2
    NFadeL 20, li20
    NFadeL 21, li21
    NFadeL 22, li22
    NFadeL 30, li30
    NFadeL 31, li31
    NFadeL 32, li32
    NFadeL 40, li40
    NFadeL 41, li41
    NFadeL 42, li42
    NFadeL 50, li50
    NFadeL 51, li51
    NFadeL 52, li52
    NFadeL 60, li60
    NFadeL 61, li61
    NFadeL 62, li62
    NFadeL 70, li70
    NFadeL 71, li71
    NFadeL 72, li72
    NFadeL 79, li79
    NFadeL 80, li80
    NFadeL 81, li81
    NFadeLm 82, li82
    NFadeLm 82, li82b
    Flashm 82, li82c
    Flash 82, li82a
    NFadeL 83, li83
    NFadeL 89, li89
    NFadeL 90, li90
    NFadeL 91, li91
    NFadeL 92, li92
    NFadeL 93, li93
    NFadeL 99, li99
    ' flasher
    'NFadeLm 190, F2a
    'NFadeLm 190, F2b
    'NFadeL 190, F2

    'backdrop lights
    If VarHidden Then
        NFadeL 139, li139
        NFadeL 140, li140
        NFadeL 141, li141
        NFadeL 142, li142
        NFadeL 149, li149
        NFadeL 150, li150
        NFadeL 151, li151
        NFadeL 152, li152
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

Sub aMetals_Hit(idx):PlaySound "fx_MetalHit2", 0, Vol(ActiveBall)*VolMetal, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Bands_Hit(idx):PlaySound "fx_rubber_band", 0, Vol(ActiveBall)*VolRB, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Posts_Hit(idx):PlaySound "fx_rubber_post", 0, Vol(ActiveBall)*VolPo, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Pins_Hit(idx):PlaySound "fx_rubber_pin", 0, Vol(ActiveBall)*VolPi, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall)*VolPlast, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, Vol(ActiveBall)*VolGates, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aWoods_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall)*VolWood, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
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

Sub PlaySoundAtVol(sound, tableobj, Volum)
  PlaySound sound, 1, Volum, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
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
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / (VolDiv/VolCol), Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub


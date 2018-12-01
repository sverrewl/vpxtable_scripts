' Mr Black - Taito 1982
' IPD No. 4586 / 4 Players
' playfield layout is similar to Williams Electronics, Defender, 1982
' VPX - version by JPSalas 2017, version 1.0.0
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

Dim bsTrough, dtbank1, dtbank2, dtbank3, dtbank4, dtbank5, dtbank6, x

Const cGameName = "mrblack"
' Const cGameName = "mrblack1" 'extra music

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
        .SplashInfoLine = "Mr Black - Taito 1982" & vbNewLine & "VPX table by JPSalas v.1.0.0"
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
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, LeftSlingshot, RightSlingshot)

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0, 1, 11, 21, 0, 0, 0, 0
        .InitKick BallRelease, 90, 4
        .InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .Balls = 3
    End With

    ' Drop targets
    set dtbank1 = new cvpmdroptarget
    dtbank1.InitDrop Array(sw45, sw55, sw65), Array(45, 55, 65)
    dtbank1.initsnd SoundFX("", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)

    set dtbank2 = new cvpmdroptarget
    dtbank2.InitDrop sw72, 72
    dtbank2.initsnd SoundFX("", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)

    set dtbank3 = new cvpmdroptarget
    dtbank3.InitDrop Array(sw4, sw14, sw24, sw34, sw44), Array(4, 14, 24, 34, 44)
    dtbank3.initsnd SoundFX("", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)

    set dtbank4 = new cvpmdroptarget
    dtbank4.InitDrop sw64, 64
    dtbank4.initsnd SoundFX("", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)

    set dtbank5 = new cvpmdroptarget
    dtbank5.InitDrop sw54, 54
    dtbank5.initsnd SoundFX("", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)

    set dtbank6 = new cvpmdroptarget
    dtbank6.InitDrop Array(sw2, sw12, sw22, sw32, sw42), Array(2, 12, 22, 32, 42)
    dtbank6.initsnd SoundFX("", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

    ' Turn on Gi
    GiOn

    Controller.Switch(60) = 1
    Controller.Switch(70) = 1
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
    If KeyCode = RightFlipperKey Then Controller.Switch(41) = 1
    If vpmKeyDown(keycode)Then Exit Sub
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If KeyCode = RightFlipperKey Then Controller.Switch(41) = 0
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
    vpmTimer.PulseSw 53
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
    vpmTimer.PulseSw 53
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

Sub sw13_Hit:PlaySoundAtVol "fx_Rubber",ActiveBall,VolRH:vpmTimer.PulseSw 13:Rub1 = 1:sw13_Timer:End Sub

Sub sw13_Timer
    Select Case Rub1
        Case 1:Rubber7.Visible = 0:Rubber20.Visible = 1:sw13.TimerEnabled = 1
        Case 2:Rubber20.Visible = 0:Rubber21.Visible = 1
        Case 3:Rubber21.Visible = 0:Rubber7.Visible = 1:sw13.TimerEnabled = 0
    End Select
    Rub1 = Rub1 + 1
End Sub

Sub sw23_Hit:PlaySoundAtVol "fx_Rubber",ActiveBall,VolRH:vpmTimer.PulseSw 23:End Sub

Sub sw33_Hit:PlaySoundAtVol "fx_Rubber",ActiveBall,VolRH:vpmTimer.PulseSw 33:Rub2 = 1:sw33_Timer:End Sub
Sub sw33_Timer
    Select Case Rub2
        Case 1:rubber18.Visible = 0:rubber9.Visible = 1:sw33.TimerEnabled = 1
        Case 2:rubber9.Visible = 0:rubber27.Visible = 1
        Case 3:rubber27.Visible = 0:rubber18.Visible = 1:sw33.TimerEnabled = 0
    End Select
    Rub2 = Rub2 + 1
End Sub

Sub sw43_Hit:PlaySoundAtVol "fx_Rubber",ActiveBall,VolRH:vpmTimer.PulseSw 43:Rub4 = 1:sw43_Timer:End Sub
Sub sw43_Timer
    Select Case Rub4
        Case 1:rubber2.Visible = 0:rubber3.Visible = 1:sw43.TimerEnabled = 1
        Case 2:rubber3.Visible = 0:rubber4.Visible = 1
        Case 3:rubber4.Visible = 0:rubber2.Visible = 1:sw43.TimerEnabled = 0
    End Select
    Rub4 = Rub4 + 1
End Sub

' Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 61:PlaySoundAtVol SoundFX("fx_bumper", DOFContactors),Bumper1,VolBump:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 71:PlaySoundAtVol SoundFX("fx_bumper", DOFContactors),Bumper2,VolBump:End Sub

' Drain
Sub Drain_Hit:PlaysoundAtVol "fx_drain",drain,1:bsTrough.AddBall Me:End Sub

' Rollovers
Sub sw3_Hit:Controller.Switch(3) = 1:PlaySoundAtVol "fx_sensor",ActiveBall,1:End Sub
Sub sw3_UnHit:Controller.Switch(3) = 0:End Sub

Sub sw31_Hit:Controller.Switch(31) = 1:PlaySoundAtVol "fx_sensor",ActiveBall,1:End Sub
Sub sw31_UnHit:Controller.Switch(31) = 0:End Sub

Sub sw5_Hit:Controller.Switch(5) = 1:PlaySoundAtVol "fx_sensor",ActiveBall,1:End Sub
Sub sw5_UnHit:Controller.Switch(5) = 0:End Sub

Sub sw15_Hit:Controller.Switch(15) = 1:PlaySoundAtVol "fx_sensor",ActiveBall,1:End Sub
Sub sw15_UnHit:Controller.Switch(15) = 0:End Sub

Sub sw25_Hit:Controller.Switch(25) = 1:PlaySoundAtVol "fx_sensor",ActiveBall,1:End Sub
Sub sw25_UnHit:Controller.Switch(25) = 0:End Sub

Sub sw35_Hit:Controller.Switch(35) = 1:PlaySoundAtVol "fx_sensor",ActiveBall,1:End Sub
Sub sw35_UnHit:Controller.Switch(35) = 0:End Sub

Sub sw45_Hit:Controller.Switch(45) = 1:PlaySoundAtVol "fx_sensor",ActiveBall,1:End Sub
Sub sw45_UnHit:Controller.Switch(45) = 0:End Sub

Sub sw62_Hit:Controller.Switch(62) = 1:PlaySoundAtVol "fx_sensor",ActiveBall,1:End Sub
Sub sw62_UnHit:Controller.Switch(62) = 0:End Sub

Sub sw62Help_Hit:Controller.Switch(62) = 1: sw62Help.TimerEnabled = 1: End Sub 'fix bug?
Sub sw62Help_Timer:Controller.Switch(62) = 0:sw62Help.TimerEnabled = 0:End Sub

Sub sw63_Hit:Controller.Switch(63) = 1:PlaySoundAtVol "fx_sensor",ActiveBall,1:End Sub
Sub sw63_UnHit:Controller.Switch(63) = 0:End Sub

Sub sw73_Hit:Controller.Switch(73) = 1:PlaySoundAtVol "fx_sensor",ActiveBall,1:End Sub
Sub sw73_UnHit:Controller.Switch(73) = 0:End Sub

Sub sw51_Hit:Controller.Switch(51) = 1:PlaySoundAtVol "fx_sensor",ActiveBall,1:End Sub
Sub sw51_UnHit:Controller.Switch(51) = 0:End Sub

Sub sw52_Hit:Controller.Switch(52) = 1:PlaySoundAtVol "fx_sensor",ActiveBall,1:End Sub
Sub sw52_UnHit:Controller.Switch(52) = 0:End Sub

' Droptargets
Sub sw45_Hit:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets),ActiveBall,VolTarg:End Sub
Sub sw55_Hit:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets),ActiveBall,VolTarg:End Sub
Sub sw65_Hit:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets),ActiveBall,VolTarg:End Sub
Sub sw72_Hit:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets),ActiveBall,VolTarg:End Sub
Sub sw4_Hit:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets),ActiveBall,VolTarg:End Sub
Sub sw14_Hit:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets),ActiveBall,VolTarg:End Sub
Sub sw24_Hit:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets),ActiveBall,VolTarg:End Sub
Sub sw34_Hit:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets),ActiveBall,VolTarg:End Sub
Sub sw44_Hit:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets),ActiveBall,VolTarg:End Sub
Sub sw54_Hit:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets),ActiveBall,VolTarg:End Sub
Sub sw64_Hit:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets),ActiveBall,VolTarg:End Sub
Sub sw2_Hit:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets),ActiveBall,VolTarg:End Sub
Sub sw12_Hit:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets),ActiveBall,VolTarg:End Sub
Sub sw22_Hit:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets),ActiveBall,VolTarg:End Sub
Sub sw32_Hit:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets),ActiveBall,VolTarg:End Sub
Sub sw42_Hit:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets),ActiveBall,VolTarg:End Sub

Sub sw45_Dropped:dtbank1.Hit 1:End Sub
Sub sw55_Dropped:dtbank1.Hit 2:End Sub
Sub sw65_Dropped:dtbank1.Hit 3:End Sub
Sub sw72_Dropped:dtbank2.Hit 1:End Sub
Sub sw4_Dropped:dtbank3.Hit 1:End Sub
Sub sw14_Dropped:dtbank3.Hit 2:End Sub
Sub sw24_Dropped:dtbank3.Hit 3:End Sub
Sub sw34_Dropped:dtbank3.Hit 4:End Sub
Sub sw44_Dropped:dtbank3.Hit 5:End Sub
Sub sw64_Dropped:dtbank4.Hit 1:End Sub
Sub sw54_Dropped:dtbank5.Hit 1:End Sub
Sub sw2_Dropped:dtbank6.Hit 1:End Sub
Sub sw12_Dropped:dtbank6.Hit 2:End Sub
Sub sw22_Dropped:dtbank6.Hit 3:End Sub
Sub sw32_Dropped:dtbank6.Hit 4:End Sub
Sub sw42_Dropped:dtbank6.Hit 5:End Sub

'Targets
Sub sw74_Hit:vpmTimer.PulseSw 74:PlaySoundAtVol SoundFX("fx_target", DOFDropTargets),ActiveBall,VolTarg:End Sub

'*********
'Solenoids
'*********

SolCallback(1) = "bsTrough.SolOut"
SolCallback(2) = "SetLamp 190,"
SolCallback(3) = "SolLockRelease"
SolCallback(4) = "dtbank1.SolDropUp"
'SolCallback(5) = ""
SolCallback(6) = "SolGate"
SolCallback(7) = "dtbank6.SolDropUp"
SolCallback(8) = "dtbank3.SolDropUp"
SolCallback(9) = "dtbank2.SolDropUp"
SolCallback(10) = "dtbank4.SolDropUp"
SolCallback(11) = "dtbank5.SolDropUp"

SolCallback(13) = "GiEffect" ' jp: unknown solenoid, it fires sometimes
SolCallback(17) = "SolGi"    '17=relay jp:actually I haven't a clue what this those either :)
SolCallback(18) = "vpmNudge.SolGameOn"

Sub SolLockRelease(Enabled)
    If Enabled Then
        Lock1a.Isdropped = 1
        Lock2a.Isdropped = 1
        Lock3a.Isdropped = 1
        Lock3a.TimerEnabled = 1
        PlaySound "fx_solenoidon", 0, 1, -0.08, 0.1 ' TODO
    End If
End Sub

Sub Lock3a_Timer 'Rise back the locks
    Lock3a.TimerEnabled = 0
    Lock1a.Isdropped = 0
    Lock2a.Isdropped = 0
    Lock3a.Isdropped = 0
    PlaySound "fx_solenoidoff", 0, 1, -0.08, 0.1 ' TODO
End Sub

Sub SolGate(Enabled)
    If Enabled Then
        GateFlipper.RotateToEnd
        PlaySoundAtVol "fx_solenoidon", GateFlipper, 1
    Else
        GateFlipper.RotateToStart
        PlaySoundAtVol "fx_solenoidoff", GateFlipper, 1
    End If
End Sub

Sub SolGi(enabled)
    If enabled Then
        GiOff
    Else
        GiOn
    End If
End Sub

'**************
' Flipper Subs
'**************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFX("fx_flipperup", DOFFlippers), LeftFlipper, VolFlip
        PlaySoundAtVol "fx_flipperup", LeftFlipper1, VolFlip
        LeftFlipper.RotateToEnd
        LeftFlipper1.RotateToEnd
    Else
        PlaySoundAtVol SoundFX("fx_flipperdown", DOFFlippers), LeftFlipper, VolFlip
        PlaySoundAtVol "fx_flipperdown", LeftFlipper1, VolFlip
        LeftFlipper.RotateToStart
        LeftFlipper1.RotateToStart
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

Sub LeftFlipper1_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, -0.1, 0.25
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
    'update the gate
    Gate.RotZ = GateFlipper.CurrentAngle
End Sub

Sub UpdateLamps()
    NFadeL 0, li0
    NFadeL 1, li1
    NFadeLm 10, li10
    Flash 10, li10a
    NFadeL 100, li100
    NFadeL 101, li101
    NFadeL 102, li102
    NFadeL 103, li103
    NFadeL 109, li109
    NFadeLm 11, li11
    Flash 11, li11a
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
    NFadeLm 132, li132a
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
    NFadeLm 70, liBumper1a
    NFadeLm 70, liBumper1b
    NFadeL 70, liBumper1c
    NFadeLm 71, liBumper2a
    NFadeLm 71, liBumper2b
    NFadeL 71, liBumper2c
    NFadeL 72, li72
    NFadeLm 79, li79
    Flash 79, li79a
    NFadeLm 80, li80
    Flash 80, li80a
    NFadeLm 81, li81
    Flash 81, li81a
    NFadeLm 82, li82
    Flash 82, li82a
    NFadeL 83, li83
    NFadeL 89, li89
    NFadeL 90, li90
    NFadeL 91, li91
    NFadeL 92, li92
    NFadeL 93, li93
    NFadeL 99, li99
    ' flasher
    NFadeLm 190, F2a
    NFadeLm 190, F2b
    NFadeL 190, F2

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



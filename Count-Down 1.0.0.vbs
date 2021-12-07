' Gottlieb's Count-Down / IPD No. 573 / April, 1979 / 4 Players
' VPX - version by JPSalas 2018, version 1.0.0
Option Explicit
Randomize

' Thalamus 2018-07-20
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' Thalamus 2018-12-18 : Added FFv2
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
Const VolGates  = 1    ' Gates volume.
Const VolMetal  = 1    ' Metals volume.
Const VolRB     = 1    ' Rubber bands volume.
Const VolPo     = 1    ' Rubber posts volume.
Const VolPi     = 1    ' Rubber pins volume.
Const VolPlast  = 1    ' Plastics volume.
Const VolWood   = 1    ' Woods volume.
Const VolKick   = 1    ' Kicker volume.
Const VolFlip   = 1    ' Flipper volume.

Const BallSize = 50
Const BallMass = 1.1

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01550000", "gts1.vbs", 3.26

' Thalamus - for Fast Flip v2
NoUpperRightFlipper
NoUpperLeftFlipper

Dim bsTrough, dtbank1, dtbank2, dtbank3, dtbank4, bsSaucer, x

Const cGameName = "countdwn"

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
        .SplashInfoLine = "Count-Down - Gottlieb 1979" & vbNewLine & "VPX table by JPSalas v.1.0.0"
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
    vpmNudge.TiltSwitch = 4
    vpmNudge.Sensitivity = 3
    vpmNudge.TiltObj = Array(Bumper1)

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0, 66, 0, 0, 0, 0, 0, 0
        .InitKick BallRelease, 90, 4
        .InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .Balls = 1
    End With

    ' Saucers
    Set bsSaucer = New cvpmBallStack
    bsSaucer.InitSaucer sw41, 41, 172, 8
    bsSaucer.InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    bsSaucer.KickForceVar = 6

    ' Drop targets
    set dtbank1 = new cvpmdroptarget
    dtbank1.InitDrop Array(sw20, sw21, sw23, sw24), Array(20, 21, 23, 24)
    dtbank1.initsnd SoundFX("", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
    dtbank1.CreateEvents "dtbank1"

    set dtbank2 = new cvpmdroptarget
    dtbank2.InitDrop Array(sw30, sw31, sw33, sw34), Array(30, 31, 33, 34)
    dtbank2.initsnd SoundFX("", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
    dtbank2.CreateEvents "dtbank2"

    set dtbank3 = new cvpmdroptarget
    dtbank3.InitDrop Array(sw60, sw61, sw63, sw64), Array(60, 61, 63, 64)
    dtbank3.initsnd SoundFX("", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
    dtbank3.CreateEvents "dtbank3"

    set dtbank4 = new cvpmdroptarget
    dtbank4.InitDrop Array(sw70, sw71, sw73, sw74), Array(70, 71, 73, 74)
    dtbank4.initsnd SoundFX("", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
    dtbank4.CreateEvents "dtbank4"

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
    If keycode = PlungerKey Then PlaySoundAtVol "fx_PlungerPull", plunger, 1:Plunger.Pullback
  If keycode=AddCreditKey then PlaySoundAtVol "fx_coin", drain, 1: vpmTimer.pulseSW (swCoin1): end if
    If vpmKeyDown(keycode)Then Exit Sub
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If vpmKeyUp(keycode)Then Exit Sub
    If keycode = PlungerKey Then PlaySoundAtVol "fx_plunger", plunger, 1:Plunger.Fire
End Sub

'*********
' Switches
'*********

' Rubber & animations
Dim Rub1, Rub2, Rub3, Rub4, Rub5, Rub6, Rub7, Rub8

Sub sw53f_Hit:vpmTimer.PulseSw 53:Rub1 = 1:sw53f_Timer:End Sub

Sub sw53f_Timer
    Select Case Rub1
        Case 1:r8.Visible = 0:r10.Visible = 1:sw53f.TimerEnabled = 1
        Case 2:r10.Visible = 0:r11.Visible = 1
        Case 3:r11.Visible = 0:r8.Visible = 1:sw53f.TimerEnabled = 0
    End Select
    Rub1 = Rub1 + 1
End Sub

Sub sw53e_Hit:vpmTimer.PulseSw 53:Rub2 = 1:sw53e_Timer:End Sub
Sub sw53e_Timer
    Select Case Rub2
        Case 1:r9.Visible = 0:r12.Visible = 1:sw53e.TimerEnabled = 1
        Case 2:r12.Visible = 0:r13.Visible = 1
        Case 3:r13.Visible = 0:r9.Visible = 1:sw53e.TimerEnabled = 0
    End Select
    Rub2 = Rub2 + 1
End Sub

Sub sw53c_Hit:vpmTimer.PulseSw 53:Rub3 = 1:sw53c_Timer:End Sub
Sub sw53c_Timer
    Select Case Rub3
        Case 1:r17.Visible = 0:r16.Visible = 1:sw53c.TimerEnabled = 1
        Case 2:r16.Visible = 0:r18.Visible = 1
        Case 3:r18.Visible = 0:r17.Visible = 1:sw53c.TimerEnabled = 0
    End Select
    Rub3 = Rub3 + 1
End Sub

Sub sw53d_Hit:vpmTimer.PulseSw 53:Rub4 = 1:sw53d_Timer:End Sub
Sub sw53d_Timer
    Select Case Rub4
        Case 1:r5.Visible = 0:r14.Visible = 1:sw53d.TimerEnabled = 1
        Case 2:r14.Visible = 0:r15.Visible = 1
        Case 3:r15.Visible = 0:r5.Visible = 1:sw53d.TimerEnabled = 0
    End Select
    Rub4 = Rub4 + 1
End Sub

Sub sw53h_Hit:vpmTimer.PulseSw 53:Rub5 = 1:sw53h_Timer:End Sub
Sub sw53h_Timer
    Select Case Rub5
        Case 1:r3.Visible = 0:r19.Visible = 1:sw53h.TimerEnabled = 1
        Case 2:r19.Visible = 0:r21.Visible = 1
        Case 3:r21.Visible = 0:r3.Visible = 1:sw53h.TimerEnabled = 0
    End Select
    Rub5 = Rub5 + 1
End Sub

Sub sw53g_Hit:vpmTimer.PulseSw 53:Rub6 = 1:sw53g_Timer:End Sub
Sub sw53g_Timer
    Select Case Rub6
        Case 1:r4.Visible = 0:r20.Visible = 1:sw53g.TimerEnabled = 1
        Case 2:r20.Visible = 0:r22.Visible = 1
        Case 3:r22.Visible = 0:r4.Visible = 1:sw53g.TimerEnabled = 0
    End Select
    Rub6 = Rub6 + 1
End Sub

Sub sw53i_Hit:vpmTimer.PulseSw 53:Rub7 = 1:sw53i_Timer:End Sub
Sub sw53i_Timer
    Select Case Rub7
        Case 1:r1.Visible = 0:r24.Visible = 1:sw53i.TimerEnabled = 1
        Case 2:r24.Visible = 0:r25.Visible = 1
        Case 3:r25.Visible = 0:r1.Visible = 1:sw53i.TimerEnabled = 0
    End Select
    Rub7 = Rub7 + 1
End Sub

Sub sw53j_Hit:vpmTimer.PulseSw 53:Rub8 = 1:sw53j_Timer:End Sub
Sub sw53j_Timer
    Select Case Rub8
        Case 1:r2.Visible = 0:r23.Visible = 1:sw53j.TimerEnabled = 1
        Case 2:r23.Visible = 0:r26.Visible = 1
        Case 3:r26.Visible = 0:r2.Visible = 1:sw53j.TimerEnabled = 0
    End Select
    Rub8 = Rub8 + 1
End Sub

Sub sw53a_Hit:vpmTimer.PulseSw 53:End Sub
Sub sw53b_Hit:vpmTimer.PulseSw 53:End Sub

' Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 51:PlaySoundAtVol SoundFX("fx_bumper", DOFContactors),Bumper1, VolBump:End Sub

' Drain & Saucers
Sub Drain_Hit:PlaysoundAtVol "fx_drain", drain, 1:bsTrough.AddBall Me:End Sub
Sub sw41_Hit::PlaySoundAtVol "fx_kicker_enter", sw41,VolKick:bsSaucer.AddBall 0:End Sub

' Rollovers
Sub sw10_Hit:Controller.Switch(10) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw10_UnHit:Controller.Switch(10) = 0:End Sub

Sub sw11_Hit:Controller.Switch(11) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw11_UnHit:Controller.Switch(11) = 0:End Sub

Sub sw13_Hit:Controller.Switch(13) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw13_UnHit:Controller.Switch(13) = 0:End Sub

Sub sw14_Hit:Controller.Switch(14) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw14_UnHit:Controller.Switch(14) = 0:End Sub

Sub sw40_Hit:Controller.Switch(40) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw40_UnHit:Controller.Switch(40) = 0:End Sub

Sub sw50_Hit:Controller.Switch(50) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw50_UnHit:Controller.Switch(50) = 0:End Sub

Sub sw43_Hit:Controller.Switch(43) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw43_UnHit:Controller.Switch(43) = 0:End Sub

Sub sw44_Hit:Controller.Switch(44) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw44_UnHit:Controller.Switch(44) = 0:End Sub

' Droptargets (only sound effect)
Sub sw20_Hit:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets),ActiveBall, 1:End Sub
Sub sw21_Hit:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets),ActiveBall, 1:End Sub
Sub sw23_Hit:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets),ActiveBall, 1:End Sub
Sub sw24_Hit:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets),ActiveBall, 1:End Sub

Sub sw30_Hit:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets),ActiveBall, 1:End Sub
Sub sw31_Hit:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets),ActiveBall, 1:End Sub
Sub sw33_Hit:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets),ActiveBall, 1:End Sub
Sub sw34_Hit:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets),ActiveBall, 1:End Sub

Sub sw60_Hit:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets),ActiveBall, 1:End Sub
Sub sw61_Hit:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets),ActiveBall, 1:End Sub
Sub sw63_Hit:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets),ActiveBall, 1:End Sub
Sub sw64_Hit:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets),ActiveBall, 1:End Sub

Sub sw70_Hit:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets),ActiveBall, 1:End Sub
Sub sw71_Hit:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets),ActiveBall, 1:End Sub
Sub sw73_Hit:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets),ActiveBall, 1:End Sub
Sub sw74_Hit:PlaySoundAtVol SoundFX("fx_droptarget", DOFDropTargets),ActiveBall, 1:End Sub

'*********
'Solenoids
'*********

SolCallback(1) = "bsTrough.SolOut"
SolCallback(2) = "vpmsolsound SoundFX(""fx_knocker"",DOFKnocker),"
SolCallback(3) = "VpmSolSound""gottlieb_10"","
SolCallback(4) = "VpmSolSound""gottlieb_100"","
SolCallback(5) = "VpmSolSound""gottlieb_1000"","
SolCallback(6) = "bsSaucer.SolOut"
SolCallback(7) = "dtbank2.SolDropUp" 'Green droptargets
SolCallback(8) = "dtbank1.SolDropUp" 'Red droptargets
SolCallback(17) = "vpmNudge.SolGameOn"

' Droptarget banks Yellow and Blue
' And extra backdrop lights
Dim N1, O1, N2, O2
N1 = 0:O1 = 0:N2 = 0:O2 = 0

Set LampCallback = GetRef("UpdateMultipleLamps")
Sub UpdateMultipleLamps
    N1 = Controller.Lamp(17) 'Yellow droptargets
    N2 = Controller.Lamp(18) 'Blue droptargets
    If N1 <> O1 Then
        If N1 Then
            dtbank3.DropSol_On
        End If
        O1 = N1
    End If
    If N2 <> O2 Then
        If N2 Then
            dtbank4.DropSol_On
        End If
        O2 = N2
    End If
    If VarHidden Then
        lix1.State = ABS(Li1.State-1) 'Match
        lix2.State = ABS(Li1.State-1) 'Game Over
    End If
End Sub

'**************
' Flipper Subs
'**************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFX("fx_flipperup", DOFFlippers), LeftFlipper1, VolFlip
        PlaySoundAtVol "fx_flipperup", LeftFlipper3, VolFlip
        LeftFlipper1.RotateToEnd:LeftFlipper2.RotateToEnd:LeftFlipper3.RotateToEnd:LeftFlipper4.RotateToEnd:LeftFlipper5.RotateToEnd
    Else
        PlaySoundAtVol SoundFX("fx_flipperdown", DOFFlippers), LeftFlipper1, VolFlip
        PlaySoundAtVol "fx_flipperdown", LeftFlipper3, VolFlip
        LeftFlipper1.RotateToStart:LeftFlipper2.RotateToStart:LeftFlipper3.RotateToStart:LeftFlipper4.RotateToStart:LeftFlipper5.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFX("fx_flipperup", DOFFlippers), RightFlipper1, VolFlip
        PlaySoundAtVol "fx_flipperup", RightFlipper3, VolFlip
        RightFlipper1.RotateToEnd:RightFlipper2.RotateToEnd:RightFlipper3.RotateToEnd:RightFlipper4.RotateToEnd:RightFlipper5.RotateToEnd
    Else
        PlaySoundAtVol SoundFX("fx_flipperdown", DOFFlippers), RightFlipper1, VolFlip
        PlaySoundAtVol "fx_flipperdown", RightFlipper3, VolFlip
        RightFlipper1.RotateToStart:RightFlipper2.RotateToStart:RightFlipper3.RotateToStart:RightFlipper4.RotateToStart:RightFlipper5.RotateToStart
    End If
End Sub

Sub LeftFlipper1_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, -0.1, 0.25
End Sub

Sub RightFlipper1_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, 0.1, 0.25
End Sub

Sub LeftFlipper2_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, -0.1, 0.25
End Sub

Sub RightFlipper5_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, 0.1, 0.25
End Sub

'*****************
'   Gi Lights
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
        NFadeL 1, li1   'game over
        NFadeLm 2, li2a 'Tilt
        NFadeL 2, li2   'Tilt
        NFadeL 3, li3   'High Game to date
        NFadeLm 4, li4a 'Same Player Shoots
        NFadeLm 4, li4b 'Same Player Shoots
    ' number to match
    ' ball in play
    End If

    NFadeL 4, li4
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
    NFadeL 35, li35
    NFadeL 36, li36
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

'Gottliebs
Patterns(0) = 0    'empty
Patterns(1) = 63   '0
Patterns(2) = 6    '1
Patterns(3) = 91   '2
Patterns(4) = 79   '3
Patterns(5) = 102  '4
Patterns(6) = 109  '5
Patterns(7) = 124  '6
Patterns(8) = 7    '7
Patterns(9) = 127  '8
Patterns(10) = 103 '9

Patterns2(0) = 128  'empty
Patterns2(1) = 191  '0
Patterns2(2) = 768 '134  '1
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

Set Digits(26) = f0
Set Digits(27) = f1

Sub UpdateLeds
    On Error Resume Next
    Dim ChgLED, ii, jj, chg, stat
    ChgLED = Controller.ChangedLEDs(&HFF, &HFFFF)
    If Not IsEmpty(ChgLED)Then
        For ii = 0 To UBound(ChgLED)
            chg = chgLED(ii, 1):stat = chgLED(ii, 2)
            For jj = 0 to 10
                If stat = Patterns(jj)OR stat = Patterns2(jj)then Digits(chgLED(ii, 0)).SetValue jj
        'debug.print stat
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

'Gottlieb System 1 games with 3 tones only (or chimes).
'These are: Charlie's Angels, Cleopatra, Close Encounters, Count Down, Dragon, Joker Poker, Pinball Pool, Sinbad, Solar Ride.
'For the other System 1 games use the file called "Gottlieb System 1 with multi-mode sound.txt"
'Added by Inkochnito

Sub editDips
    Dim vpmDips:Set vpmDips = New cvpmDips
    With vpmDips
        .AddForm 700, 400, "System 1 (3 tones) - DIP switches"
        .AddFrame 205, 0, 190, "Maximum credits", &H00030000, Array("5 credits", 0, "8 credits", &H00020000, "10 credits", &H00010000, "15 credits", &H00030000) 'dip 17&18
        .AddFrame 0, 0, 190, "Coin chute control", &H00040000, Array("seperate", 0, "same", &H00040000)                                                          'dip 19
        .AddFrame 0, 46, 190, "Game mode", &H00000400, Array("extra ball", 0, "replay", &H00000400)                                                              'dip 11
        .AddFrame 0, 92, 190, "High game to date awards", &H00200000, Array("no award", 0, "3 replays", &H00200000)                                              'dip 22
        .AddFrame 0, 138, 190, "Balls per game", &H00000100, Array("5 balls", 0, "3 balls", &H00000100)                                                          'dip 9
        .AddFrame 0, 184, 190, "Tilt effect", &H00000800, Array("game over", 0, "ball in play only", &H00000800)                                                 'dip 12
        .AddChk 205, 80, 190, Array("Match feature", &H00000200)                                                                                                 'dip 10
        .AddChk 205, 95, 190, Array("Credits displayed", &H00001000)                                                                                             'dip 13
        .AddChk 205, 110, 190, Array("Play credit button tune", &H00002000)                                                                                      'dip 14
        .AddChk 205, 125, 190, Array("Play tones when scoring", &H00080000)                                                                                      'dip 20
        .AddChk 205, 140, 190, Array("Play coin switch tune", &H00400000)                                                                                        'dip 23
        .AddChk 205, 155, 190, Array("High game to date displayed", &H00100000)                                                                                  'dip 21
        .AddLabel 50, 240, 300, 20, "After hitting OK, press F3 to reset game with new settings."
        .ViewDips
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


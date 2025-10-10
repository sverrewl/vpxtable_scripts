' Freefall / IPD No. 953 / January, 1981 / 4 Players
' http://www.ipdb.org/machine.cgi?id=953
' VPX8 v6.0.0 by JPSalas 2024
' Many parts of the script inspired/copied from the old table by TAB/Destruk/LuvThatApex/Inkochnito

Option Explicit
Randomize

Const BallSize = 50
Const BallMass = 1

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01200000", "stern.vbs", 3.02

Dim bsTrough, dtDrop5, dtDrop3, plSkyway, plSkyraider
Dim x

Const cGameName = "freefall" ' freefall rom
'Const cGameName = "freefafp" ' freefall freeplay

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
end if

if B2SOn = true then VarHidden = 1

' Standard Sounds
Const SSolenoidOn = "fx_Solenoidon"
Const SSolenoidOff = "fx_SolenoidOff"
Const SCoin = "fx_coin"

Sub table1_Init
    vpmInit me

    With Controller
        .GameName = cGameName
        .SplashInfoLine = "Freefall - Stern 1991" & vbNewLine & "VPX-VPM table by JPSalas v6.0.0"
        .Games(cGameName).Settings.Value("rol") = 0   '1= rotated display, 0= normal
        .Games(cGameName).Settings.Value("sound") = 1 '1 enabled rom sound
        .HandleMechanics = 0
        .HandleKeyboard = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .ShowTitle = 0
        .Hidden = VarHidden
        If Err Then MsgBox Err.Description
    End With
    On Error Goto 0
    Controller.SolMask(0) = 0
    vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the Timer to renable all the solenoids after 2 seconds
    Controller.Run GetPlayerHWnd

    ' Nudging
    vpmNudge.TiltSwitch = 7
    vpmNudge.Sensitivity = 5
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingshot)

    ' Trough & Ball Release
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 8, 33, 34, 35, 0, 0, 0, 0
        .InitKick BallRelease, 180, 10
        .InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .InitEntrySnd "fx_solenoid", "fx_solenoid"
        .IsTrough = True
        .Balls = 3
    End With

    ' 5 Droptargets Left
    Set dtDrop5 = New cvpmDropTarget
    With dtDrop5
        .InitDrop Array(sw21, sw20, sw19, sw18, sw17), Array(21, 20, 19, 18, 17)
        .initsnd SoundFX("fx_droptarget", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
        .CreateEvents "dtDrop5"
    End With

    ' 3 Droptargets Center
    Set dtDrop3 = New cvpmDropTarget
    With dtDrop3
        .InitDrop Array(sw24, sw23, sw22), Array(24, 23, 22)
        .initsnd SoundFX("fx_droptarget", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
        .CreateEvents "dtDrop3"
    End With

    ' Skyway Impulse Plunger
    Const IMPowerSetting = 42 'Plunger Power
    Const IMTime = 0.6        ' Time in seconds for Full Plunge
    Set plSkyway = New cvpmImpulseP
    With plSkyway
        .InitImpulseP sw5, IMPowerSetting, IMTime
        .Random 0.3
        .switch 5
        .InitExitSnd SoundFX("fx_popper", DOFContactors), SoundFX("fx_popper", DOFContactors)
        .CreateEvents "plSkyway"
    End With

    ' Skyraider Impulse Plunger
    Const IMPowerSetting2 = 42 'Plunger Power
    Const IMTime2 = 0.6        ' Time in seconds for Full Plunge
    Set plSkyraider = New cvpmImpulseP
    With plSkyraider
        .InitImpulseP sw4, IMPowerSetting2, IMTime2
        .Random 0.3
        .switch 4
        .InitExitSnd SoundFX("fx_popper", DOFContactors), SoundFX("fx_popper", DOFContactors)
        .CreateEvents "plSkyraider"
    End With

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1
    RealTime.Enabled = 1
End Sub

'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If vpmKeyDown(keycode) Then Exit Sub
    If keycode = PlungerKey Then PlaySoundAt "fx_PlungerPull",Plunger:Plunger.Pullback
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If vpmKeyUp(keycode) Then Exit Sub
    If keycode = PlungerKey Then PlaySoundAt "fx_plunger",Plunger:Plunger.Fire
End Sub

'*********
' Switches
'*********

' Slings
Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    PlaySoundAt SoundFX("fx_slingshot", DOFContactors),Lemk
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 13
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
    PlaySoundAt SoundFX("fx_slingshot", DOFContactors), Remk
    RightSling4.Visible = 1
    Remk.RotX = 26
    RStep = 0
    vpmTimer.PulseSw 12
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

Sub sw11_Hit:vpmTimer.PulseSw 11:End Sub

' Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 16:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper1:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 15:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper2:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 14:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper3:End Sub

' Spinners

Sub Spinner1_Spin:vpmTimer.PulseSw 36:PlaySoundAt "fx_spinner",Spinner1:End Sub

' Eject holes
Sub Drain_Hit:PlaySoundAt "fx_drain", Drain:bsTrough.AddBall Me:End Sub
Sub vuk_hit:vuk.kick 0, 20, 1.56:End Sub

' Rollovers
Sub sw37_Hit:DOF 113, DOFOn:Controller.Switch(37) = 1:PlaySoundAt "fx_sensor", sw37:End Sub
Sub sw37_UnHit:DOF 113, DOFOff:Controller.Switch(37) = 0:End Sub

Sub sw9_Hit:DOF 117, DOFOn:Controller.Switch(9) = 1:PlaySoundAt "fx_sensor", sw9:End Sub
Sub sw9_UnHit:DOF 117, DOFOff:Controller.Switch(9) = 0:End Sub

Sub sw40_Hit:DOF 115, DOFOn:Controller.Switch(40) = 1:PlaySoundAt "fx_sensor", sw40:End Sub
Sub sw40_UnHit:DOF 115, DOFOff:Controller.Switch(40) = 0:End Sub

Sub sw40a_Hit:DOF 116, DOFOn:Controller.Switch(40) = 1:PlaySoundAt "fx_sensor", sw40a:End Sub
Sub sw40a_UnHit:DOF 116, DOFOff:Controller.Switch(40) = 0:End Sub

Sub sw9a_Hit:DOF 118, DOFOn:Controller.Switch(9) = 1:PlaySoundAt "fx_sensor", sw9a:End Sub
Sub sw9a_UnHit:DOF 118, DOFOff:Controller.Switch(9) = 0:End Sub

Sub sw37a_Hit:DOF 114, DOFOn:Controller.Switch(37) = 1:PlaySoundAt "fx_sensor", sw37a:End Sub
Sub sw37a_UnHit:DOF 114, DOFOff:Controller.Switch(37) = 0:End Sub

Sub sw29_Hit:DOF 110, DOFOn:Controller.Switch(29) = 1:PlaySoundAt "fx_sensor", sw29:End Sub
Sub sw29_UnHit:DOF 110, DOFOff:Controller.Switch(29) = 0:End Sub

Sub sw28_Hit:Controller.Switch(28) = 1:PlaySoundAt "fx_sensor", sw28:End Sub
Sub sw28_UnHit:Controller.Switch(28) = 0:End Sub

Sub sw27_Hit:Controller.Switch(27) = 1:PlaySoundAt "fx_sensor", sw27:End Sub
Sub sw27_UnHit:Controller.Switch(27) = 0:End Sub

Sub sw26_Hit:DOF 112, DOFOn:Controller.Switch(26) = 1:PlaySoundAt "fx_sensor", sw26:End Sub
Sub sw26_UnHit:DOF 112, DOFOff:Controller.Switch(26) = 0:End Sub

' Targets

Sub sw30_Hit:vpmTimer.PulseSw 30:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw31_Hit:vpmTimer.PulseSw 31:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw32_Hit:vpmTimer.PulseSw 32:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw35_Hit:vpmTimer.PulseSw 35:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw29b_Hit:vpmTimer.PulseSw 29:PlaySoundAtBall SoundFXDOF("fx_target",111,DOFPulse,DOFTargets):End Sub
Sub sw31b_Hit:vpmTimer.PulseSw 31:PlaySoundAtBall SoundFXDOF("fx_target",111,DOFPulse,DOFTargets):End Sub
Sub sw25_Hit:vpmTimer.PulseSw 25:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw26b_Hit:vpmTimer.PulseSw 26:PlaySoundAtBall SoundFXDOF("fx_target",111,DOFPulse,DOFTargets):End Sub

'********************
'     Solenoids
'********************

'SolCallback(1)="vpmSolSound ""bumper"","
'SolCallback(2)="vpmSolSound ""bumper"","
'SolCallback(3)="vpmSolSound ""bumper"","
'SolCallback(4)="vpmSolSound ""sling"","
'SolCallback(12)="vpmSolSound ""sling"","

SolCallback(6) = "vpmsolsound SoundFX(""fx_knocker"",DOFKnocker),"
SolCallback(7) = "SolSkyway"
SolCallback(8) = "SolBallRelease"
SolCallback(9) = "SolSkyRaider"
SolCallback(10) = "dtDrop5.SolDropUp"
SolCallback(11) = "dtDrop3.SolDropUp"
SolCallback(13) = "SolOutHole"
SolCallback(17) = "SolBallLock"
SolCallback(19) = "vpmNudge.SolGameOn"
SolCallback(20) = "SolDiv"

Dim sw5Step, sw4Step

Sub SolSkyway(Enabled)
    If Enabled Then
        plSkyway.AutoFire
        PLaySoundAt"fx_kicker", Remk2
        sw5Step = 0
        Remk2.RotX = 26
        sw5t.TimerEnabled = 1
    End If
End Sub

Sub sw5t_Timer
    Select Case sw5Step
        Case 1:Remk2.RotX = 14
        Case 2:Remk2.RotX = 2
        Case 3:Remk2.RotX = -10:sw5t.TimerEnabled = 0
    End Select

    sw5Step = sw5Step + 1
End Sub

Sub SolSkyraider(Enabled)
    If Enabled Then
        plSkyraider.AutoFire
        PLaySoundAt"fx_kicker", Remk1
        sw4Step = 0
        Remk1.RotX = 26
        sw4t.TimerEnabled = 1
    End If
End Sub

Sub sw4t_Timer
    Select Case sw4Step
        Case 1:Remk1.RotX = 14
        Case 2:Remk1.RotX = 2
        Case 3:Remk1.RotX = -10:sw4t.TimerEnabled = 0
    End Select

    sw4Step = sw4Step + 1
End Sub

Sub SolBallRelease(Enabled)
    If Enabled Then
        bsTrough.ExitSol_On
    End If
End Sub

Sub SolOutHole(Enabled)
    If Enabled Then
        bsTrough.EntrySol_On
        Drain.DestroyBall
    End If
End Sub

Sub SolDiv(Enabled)
    If Enabled Then
        PlaySoundAt "fx_solenoidon", MGate
        MGate.RotateToEnd
    Else
        PlaySoundAt "fx_solenoidoff", MGate
        MGate.RotateToStart
    End If
End Sub

'*****************
' Lock
'*****************

Sub lock3_Hit:PlaysoundAt "fx_sensor", lock3:Controller.Switch(39) = 1:End Sub

Sub SolBallLock(Enabled)
    If Enabled Then
        Controller.Switch(39) = 0
        lock1a.IsDropped = 1:lock2a.IsDropped = 1:lock3a.IsDropped = 1
        PlaySoundAt SoundFX("fx_solenoidon",DOFContactors), Lock1
        lock1.kick 180, 4:lock2.kick 180, 4:lock3.kick 180, 4
    Else
        lock1a.IsDropped = 0:lock2a.IsDropped = 0:lock3a.IsDropped = 0
        PlaySoundAt SoundFX("fx_solenoidoff",DOFContactors), Lock1
    End If
End Sub

'**************
' Flipper Subs
'**************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Dim LeftUP, RightUP
LeftUP = 0
RightUP = 0

Sub SolLFlipper(Enabled)
    If Enabled Then
        LeftUP = 1:CheckFlippers
        PlaySoundAt SoundFX("fx_flipperup", DOFFlippers), LeftFlipper
        LeftFlipper.RotateToEnd
        LeftFlipperOn = 1
    Else
        LeftUP = 0
        PlaySoundAt SoundFX("fx_flipperdown", DOFFlippers), LeftFlipper
        LeftFlipper.RotateToStart
        LeftFlipperOn = 0
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        RightUP = 1:CheckFlippers
        PlaySoundAt SoundFX("fx_flipperup", DOFFlippers), RightFlipper
        RightFlipper.RotateToEnd
        RightFlipperOn = 1
    Else
        RightUP = 0
        PlaySoundAt SoundFX("fx_flipperdown", DOFFlippers), RightFlipper
        RightFlipper.RotateToStart
        RightFlipperOn = 0
    End If
End Sub

Sub CheckFlippers
    If LeftUP AND RightUP Then
        vpmTimer.PulseSw 10
    End If
End Sub

' flippers hit Sound

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
End Sub

'*********************************************************
' Real Time Flipper adjustments - by JLouLouLou & JPSalas
'        (to enable flipper tricks)
'*********************************************************

Dim FlipperPower
Dim FlipperElasticity
Dim SOSTorque, SOSAngle
Dim FullStrokeEOS_Torque, LiveStrokeEOS_Torque
Dim LeftFlipperOn
Dim RightFlipperOn

Dim LLiveCatchTimer
Dim RLiveCatchTimer
Dim LiveCatchSensivity

FlipperPower = 3600
FlipperElasticity = 0.6
FullStrokeEOS_Torque = 0.6 ' EOS Torque when flipper hold up ( EOS Coil is fully charged. Ampere increase due to flipper can't move or when it pushed back when "On". EOS Coil have more power )
LiveStrokeEOS_Torque = 0.3 ' EOS Torque when flipper rotate to end ( When flipper move, EOS coil have less Ampere due to flipper can freely move. EOS Coil have less power )

LeftFlipper.EOSTorqueAngle = 10
RightFlipper.EOSTorqueAngle = 10

SOSTorque = 0.2
SOSAngle = 6

LiveCatchSensivity = 10

LLiveCatchTimer = 0
RLiveCatchTimer = 0

LeftFlipper.TimerInterval = 1
LeftFlipper.TimerEnabled = 1

Sub LeftFlipper_Timer 'flipper's tricks timer
'Start Of Stroke Flipper Stroke Routine : Start of Stroke for Tap pass and Tap shoot
    If LeftFlipper.CurrentAngle >= LeftFlipper.StartAngle - SOSAngle Then LeftFlipper.Strength = FlipperPower * SOSTorque else LeftFlipper.Strength = FlipperPower : End If

'End Of Stroke Routine : Livecatch and Emply/Full-Charged EOS
  If LeftFlipperOn = 1 Then
    If LeftFlipper.CurrentAngle = LeftFlipper.EndAngle then
      LeftFlipper.EOSTorque = FullStrokeEOS_Torque
      LLiveCatchTimer = LLiveCatchTimer + 1
      If LLiveCatchTimer < LiveCatchSensivity Then
        LeftFlipper.Elasticity = 0
      Else
        LeftFlipper.Elasticity = FlipperElasticity
        LLiveCatchTimer = LiveCatchSensivity
      End If
    End If
  Else
    LeftFlipper.Elasticity = FlipperElasticity
    LeftFlipper.EOSTorque = LiveStrokeEOS_Torque
    LLiveCatchTimer = 0
  End If


'Start Of Stroke Flipper Stroke Routine : Start of Stroke for Tap pass and Tap shoot
    If RightFlipper.CurrentAngle <= RightFlipper.StartAngle + SOSAngle Then RightFlipper.Strength = FlipperPower * SOSTorque else RightFlipper.Strength = FlipperPower : End If

'End Of Stroke Routine : Livecatch and Emply/Full-Charged EOS
  If RightFlipperOn = 1 Then
    If RightFlipper.CurrentAngle = RightFlipper.EndAngle Then
      RightFlipper.EOSTorque = FullStrokeEOS_Torque
      RLiveCatchTimer = RLiveCatchTimer + 1
      If RLiveCatchTimer < LiveCatchSensivity Then
        RightFlipper.Elasticity = 0
      Else
        RightFlipper.Elasticity = FlipperElasticity
        RLiveCatchTimer = LiveCatchSensivity
      End If
    End If
  Else
    RightFlipper.Elasticity = FlipperElasticity
    RightFlipper.EOSTorque = LiveStrokeEOS_Torque
    RLiveCatchTimer = 0
  End If
End Sub

'**********************************************************
'     JP's Lamp Fading for VPX and Vpinmame v4.0
' FadingStep used for all kind of lamps
' FlashLevel used for modulated flashers
' LampState keep the real lamp state in a array
'**********************************************************

Dim LampState(200), FadingStep(200), FlashLevel(200)

InitLamps() ' turn off the lights and flashers and reset them to the default parameters

' vpinmame Lamp & Flasher Timers

Sub LampTimer_Timer()
    Dim chgLamp, num, chg, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp)Then
        For ii = 0 To UBound(chgLamp)
            LampState(chgLamp(ii, 0)) = chgLamp(ii, 1)       'keep the real state in an array
            FadingStep(chgLamp(ii, 0)) = chgLamp(ii, 1)
        Next
    End If
    UpdateLeds
    UpdateLamps
End Sub

Sub UpdateLamps
    Lamp 1, l1
    Lamp 2, l2
    Lamp 3, l3
    Lampm 4, l4b
    Lamp 4, l4
    Lampm 5, l5b
    Lamp 5, l5
    Lampm 6, l6b
    Lamp 6, l6
    Lamp 7, l7
    Lamp 8, l8
    Lampm 9, l9b
    Lamp 9, l9
    Lampm 10, l10b
    Lamp 10, l10
    Lampm 11, l11b  'shoot Again
    Lamp 11, l11
    Lampm 12, l12b
    Lamp 12, l12
    Lamp 13, l13 'Highscore to date
    Lamp 14, l14
    Lamp 15, l15
    'Lamp 16, l16
    Lamp 17, l17
    Lamp 18, l18
    Lamp 19, l19
    Lampm 20, l20
    Lamp 20, l20b
    Lampm 21, l21b
    Lamp 21, l21
    Lamp 22, l22
    Lamp 23, l23
    Lamp 24, l24
    Lampm 25, l25b
    Lamp 25, l25
    Lampm 26, l26b
    Lamp 26, l26
    Lampm 27, l27b
    Lamp 27, l27
    Lamp 28, l28
    Lampm 29, l29b
    Lamp 29, l29
    Lamp 30, l30
    Lamp 31, l31
    'Lamp 32, l32
    Lamp 33, l33
    Lamp 34, l34
    Lamp 35, l35
    Lampm 36, l36b
    Lamp 36, l36
    Lampm 37, l37b
    Lamp 37, l37
    'Lamp 38, l38
    Lamp 39, l39
    Lamp 40, l40
    Lampm 41, l41b
    Lamp 41, l41
    Lampm 42, l42b
    Lamp 42, l42
    Lamp 43, l43
    Lamp 44, l44
    Lampim 45, l45b 'Ball in play
    Lamp 45, l45 'Game Over
    Lamp 46, l46
    Lamp 47, l47
    'Lamp 48, l48
    Lamp 49, l49
    Lamp 50, l50
    Lamp 51, l51
    Lampm 52, l52b
    Lamp 52, l52
    Lamp 53, l53
    Lamp 54, l54
    Lamp 55, l55
    Lamp 56, l56
    Lampm 57, l57b
    Lamp 57, l57
    Lamp 58, l58
    Lamp 59, l59
    Lamp 60, l60
    Lamp 61, l61 'tilt
    Lamp 62, l62
    Lamp 63, l63 'match
    'Lamp 64, l64
End Sub

' div lamp subs

' Normal Lamp & Flasher subs

Sub InitLamps()
    Dim x
    LampTimer.Interval = 40
    LampTimer.Enabled = 1
    For x = 0 to 200
        FadingStep(x) = 0
        FlashLevel(x) = 0
    Next
End Sub

Sub SetLamp(nr, value) ' 0 is off, 1 is on
    FadingStep(nr) = abs(value)
End Sub

' Lights: used for VPX standard lights, the fading is handled by VPX itself, they are here to be able to make them work together with the flashers

Sub Lamp(nr, object)
    Select Case FadingStep(nr)
        Case 1:object.state = 1:FadingStep(nr) = -1
        Case 0:object.state = 0:FadingStep(nr) = -1
    End Select
End Sub

Sub Lampm(nr, object) ' used for multiple lights, it doesn't change the fading state
    Select Case FadingStep(nr)
        Case 1:object.state = 1
        Case 0:object.state = 0
    End Select
End Sub

' Inverted Lights: turn on when the roms turn them off - mostly ball inplay when game over is off

Sub Lampi(nr, object)
    Select Case FadingStep(nr)
        Case 0:object.state = 1:FadingStep(nr) = -1
        Case 1:object.state = 0:FadingStep(nr) = -1
    End Select
End Sub

Sub Lampim(nr, object) ' used for multiple lights, it doesn't change the fading state
    Select Case FadingStep(nr)
        Case 0:object.state = 1
        Case 1:object.state = 0
    End Select
End Sub

' Flashers:  0 starts the fading until it is off

Sub Flash(nr, object)
    Dim tmp
    Select Case FadingStep(nr)
        Case 1:Object.IntensityScale = 1:FadingStep(nr) = -1
        Case 0
            tmp = Object.IntensityScale * 0.85 - 0.01
            If tmp > 0 Then
                Object.IntensityScale = tmp
            Else
                Object.IntensityScale = 0
                FadingStep(nr) = -1
            End If
    End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it doesn't change the fading state
    Dim tmp
    Select Case FadingStep(nr)
        Case 1:Object.IntensityScale = 1
        Case 0
            tmp = Object.IntensityScale * 0.85 - 0.01
            If tmp > 0 Then
                Object.IntensityScale = tmp
            Else
                Object.IntensityScale = 0
            End If
    End Select
End Sub

' Desktop Objects: Reels & texts

' Reels - 4 steps fading
Sub Reel(nr, object)
    Select Case FadingStep(nr)
        Case 1:object.SetValue 1:FadingStep(nr) = -1
        Case 0:object.SetValue 2:FadingStep(nr) = 2
        Case 2:object.SetValue 3:FadingStep(nr) = 3
        Case 3:object.SetValue 0:FadingStep(nr) = -1
    End Select
End Sub

Sub Reelm(nr, object)
    Select Case FadingStep(nr)
        Case 1:object.SetValue 1
        Case 0:object.SetValue 2
        Case 2:object.SetValue 3
        Case 3:object.SetValue 0
    End Select
End Sub

' Reels non fading
Sub NfReel(nr, object)
    Select Case FadingStep(nr)
        Case 1:object.SetValue 1:FadingStep(nr) = -1
        Case 0:object.SetValue 0:FadingStep(nr) = -1
    End Select
End Sub

Sub NfReelm(nr, object)
    Select Case FadingStep(nr)
        Case 1:object.SetValue 1
        Case 0:object.SetValue 0
    End Select
End Sub

'Texts

Sub Text(nr, object, message)
    Select Case FadingStep(nr)
        Case 1:object.Text = message:FadingStep(nr) = -1
        Case 0:object.Text = "":FadingStep(nr) = -1
    End Select
End Sub

Sub Textm(nr, object, message)
    Select Case FadingStep(nr)
        Case 1:object.Text = message
        Case 0:object.Text = ""
    End Select
End Sub

' Modulated Subs for the WPC tables

Sub SetModLamp(nr, level)
    FlashLevel(nr) = level / 150 'lights & flashers
End Sub

Sub LampMod(nr, object)          ' modulated lights used as flashers
    Object.IntensityScale = FlashLevel(nr)
    Object.State = 1             'in case it was off
End Sub

Sub FlashMod(nr, object)         'sets the flashlevel from the SolModCallback
    Object.IntensityScale = FlashLevel(nr)
End Sub

'Walls, flashers, ramps and Primitives used as 4 step fading images
'a,b,c,d are the images used from on to off

Sub FadeObj(nr, object, a, b, c, d)
    Select Case FadingStep(nr)
        Case 1:object.image = a:FadingStep(nr) = -1
        Case 0:object.image = b:FadingStep(nr) = 2
        Case 2:object.image = c:FadingStep(nr) = 3
        Case 3:object.image = d:FadingStep(nr) = -1
    End Select
End Sub

Sub FadeObjm(nr, object, a, b, c, d)
    Select Case FadingStep(nr)
        Case 1:object.image = a
        Case 0:object.image = b
        Case 2:object.image = c
        Case 3:object.image = d
    End Select
End Sub

Sub NFadeObj(nr, object, a, b)
    Select Case FadingStep(nr)
        Case 1:object.image = a:FadingStep(nr) = -1
        Case 0:object.image = b:FadingStep(nr) = -1
    End Select
End Sub

Sub NFadeObjm(nr, object, a, b)
    Select Case FadingStep(nr)
        Case 1:object.image = a
        Case 0:object.image = b
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

'*********************************
' Diverse Collection Hit Sounds
'*********************************

Sub aMetals_Hit(idx):PlaySoundAtBall "fx_MetalHit":End Sub
Sub aMetalWires_Hit(idx):PlaySoundAtBall "fx_MetalWire":End Sub
Sub aRubber_Bands_Hit(idx):PlaySoundAtBall "fx_rubber_band":End Sub
Sub aRubber_LongBands_Hit(idx):PlaySoundAtBall "fx_rubber_longband":End Sub
Sub aRubber_Posts_Hit(idx):PlaySoundAtBall "fx_rubber_post":End Sub
Sub aRubber_Pins_Hit(idx):PlaySoundAtBall "fx_rubber_pin":End Sub
Sub aRubber_Pegs_Hit(idx):PlaySoundAtBall "fx_rubber_peg":End Sub
Sub aPlastics_Hit(idx):PlaySoundAtBall "fx_PlasticHit":End Sub
Sub aGates_Hit(idx):PlaySoundAtBall "fx_Gate":End Sub
Sub aWoods_Hit(idx):PlaySoundAtBall "fx_Woodhit":End Sub

'***************************************************************
'             Supporting Ball & Sound Functions v4.0
'  includes random pitch in PlaySoundAt and PlaySoundAtBall
'***************************************************************

Dim TableWidth, TableHeight

TableWidth = Table1.width
TableHeight = Table1.height

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / TableWidth-1
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
    BallVel = (SQR((ball.VelX ^2) + (ball.VelY ^2)))
End Function

Function AudioFade(ball) 'only on VPX 10.4 and newer
    Dim tmp
    tmp = ball.y * 2 / TableHeight-1
    If tmp > 0 Then
        AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10))
    End If
End Function

Sub PlaySoundAt(soundname, tableobj) 'play sound at X and Y position of an object, mostly bumpers, flippers and other fast objects
    PlaySound soundname, 0, 1, Pan(tableobj), 0.2, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname) ' play a sound at the ball position, like rubbers, targets, metals, plastics
    PlaySound soundname, 0, Vol(ActiveBall), pan(ActiveBall), 0.2, Pitch(ActiveBall) * 10, 0, 0, AudioFade(ActiveBall)
End Sub

Function RndNbr(n) 'returns a random number between 1 and n
    Randomize timer
    RndNbr = Int((n * Rnd) + 1)
End Function

'***********************************************
'   JP's VP10 Rolling Sounds + Ballshadow v4.0
'   uses a collection of shadows, aBallShadow
'***********************************************

Const tnob = 19   'total number of balls
Const lob = 0     'number of locked balls
Const maxvel = 40 'max ball velocity
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub RollingUpdate()
    Dim BOT, b, ballpitch, ballvol, speedfactorx, speedfactory
    BOT = GetBalls

    ' stop the sound of deleted balls
    For b = UBound(BOT) + 1 to tnob
        rolling(b) = False
        StopSound("fx_ballrolling" & b)
        aBallShadow(b).Y = 3000
    Next

    ' exit the sub if no balls on the table
    If UBound(BOT) = lob - 1 Then Exit Sub 'there no extra balls on this table

    ' play the rolling sound for each ball and draw the shadow
    For b = lob to UBound(BOT)
        aBallShadow(b).X = BOT(b).X
        aBallShadow(b).Y = BOT(b).Y
        aBallShadow(b).Height = BOT(b).Z -Ballsize/2

        If BallVel(BOT(b))> 1 Then
            If BOT(b).z <30 Then
                ballpitch = Pitch(BOT(b))
                ballvol = Vol(BOT(b))
            Else
                ballpitch = Pitch(BOT(b)) + 50000 'increase the pitch on a ramp
                ballvol = Vol(BOT(b)) * 3
            End If
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, ballvol, Pan(BOT(b)), 0, ballpitch, 1, 0, AudioFade(BOT(b))
        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If

        ' rothbauerw's Dropping Sounds
        If BOT(b).VelZ <-1 and BOT(b).z <55 and BOT(b).z> 27 Then 'height adjust for ball drop sounds
            PlaySound "fx_balldrop", 0, ABS(BOT(b).velz) / 17, Pan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
        End If

        ' jps ball speed & spin control
            BOT(b).AngMomZ = BOT(b).AngMomZ * 0.95
        If BOT(b).VelX AND BOT(b).VelY <> 0 Then
            speedfactorx = ABS(maxvel / BOT(b).VelX)
            speedfactory = ABS(maxvel / BOT(b).VelY)
            If speedfactorx <1 Then
                BOT(b).VelX = BOT(b).VelX * speedfactorx
                BOT(b).VelY = BOT(b).VelY * speedfactorx
            End If
            If speedfactory <1 Then
                BOT(b).VelX = BOT(b).VelX * speedfactory
                BOT(b).VelY = BOT(b).VelY * speedfactory
            End If
        End If
    Next
End Sub

'*****************************
' Ball 2 Ball Collision Sound
'*****************************

Sub OnBallBallCollision(ball1, ball2, velocity)
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

'******************
' RealTime Updates
'******************

Sub RealTime_Timer
    RollingUpdate
    GIUpdate
End Sub

Sub Mgate_Animate: MgateP.RotZ = Mgate.CurrentAngle: End Sub

Sub GiON
    PlaySound"fx_GiOn"
    For each x in aGiLights
        x.State = 1
    Next
End Sub

Sub GiOFF
    PlaySound"fx_GiOff"
    For each x in aGiLights
        x.State = 0
    Next
End Sub

Dim OldGiState
OldGiState = -1 'start witht he Gi off

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

'********************
'Stern Free Fall
'added by Inkochnito
'********************
Sub editDips
    Dim vpmDips:Set vpmDips = New cvpmDips
    With vpmDips
        .AddForm 500, 400, "Free Fall - DIP switches"
        .AddFrame 0, 0, 190, "Maximum credits", &H00060000, Array("10 credits", 0, "15 credits", &H00020000, "25 credits", &H00040000, "40 credits", &H00060000)      'dip 18&19
        .AddFrame 0, 76, 190, "High game to date", 49152, Array("points", 0, "1 free game", &H00004000, "2 free games", 32768, "3 free games", 49152)                 'dip 15&16
        .AddFrame 0, 152, 190, "Special award", &HC0000000, Array("no award", 0, "100.000 points", &H40000000, "extra ball", &H80000000, "replay", &HC0000000)        'dip 31&32
        .AddFrame 0, 228, 190, "Add-a-ball memory", &H00801000, Array("1 ball only", 0, "3 balls", &H00800000, "5 balls", &H00801000)                                 'dip 13&24
        .AddChk 0, 300, 190, Array("Bonus multiplier in memory", &H10000000)                                                                                          'dip 29
        .AddChk 0, 315, 190, Array("Match feature", &H00100000)                                                                                                       'dip 21
        .AddChk 0, 330, 190, Array("Credits displayed", &H00080000)                                                                                                   'dip 20
        .AddFrame 205, 0, 190, "High score feature", &H00000020, Array("extra ball", 0, "replay", &H00000020)                                                         'dip 6
        .AddFrame 205, 48, 190, "Outlane special lites when", &H00400000, Array("completed card 2 times", 0, "completed card 1 time", &H00400000)                     'dip 23
        .AddFrame 205, 96, 190, "Outlane special lites when", &H00000010, Array("3 ball feature completed 2 times", 0, "3 ball feature completed 1 time", &H00000010) 'dip 5
        .AddFrame 205, 146, 190, "Arrow-card selector", &H00200000, Array("1 arrow on", 0, "2 arrows on", &H00200000)                                                 'dip 22
        .AddFrame 205, 193, 190, "Balls per game", &H00000040, Array("3 balls", 0, "5 balls", &H00000040)                                                             'dip 7
        .AddFrame 205, 242, 190, "Special limit", &H20000000, Array("1 per game", 0, "1 per ball", &H20000000)                                                        'dip 30
        .AddChk 205, 300, 190, Array("Background sound", &H00000080)                                                                                                  'dip 8
        .AddChk 205, 315, 190, Array("Talking feature", &H00008000)                                                                                                   'dip 17
        .AddChk 205, 330, 190, Array("Sky diver lites in memory", &H00002000)                                                                                         'dip 14
        .AddLabel 50, 350, 300, 15, "After hitting OK, press F3 to reset game with new settings."
        .ViewDips
    End With
End Sub
Set vpmShowDips = GetRef("editDips")

'*********************************
' Table Options F12 User Options
'*********************************
' Table1.Option arguments are:
' - option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional array of literal strings

Dim LUTImage

Sub Table1_OptionEvent(ByVal eventId)
    Dim x, y

    'LUT
    LutImage = Table1.Option("Select LUT", 0, 21, 1, 0, 0, Array("Normal 0", "Normal 1", "Normal 2", "Normal 3", "Normal 4", "Normal 5", "Normal 6", "Normal 7", "Normal 8", "Normal 9", "Normal 10", _
        "Warm 0", "Warm 1", "Warm 2", "Warm 3", "Warm 4", "Warm 5", "Warm 6", "Warm 7", "Warm 8", "Warm 9", "Warm 10") )
    UpdateLUT

    ' Cabinet rails
    x = Table1.Option("Cabinet Rails", 0, 1, 1, 1, 0, Array("Hide", "Show") )
    For each y in aRails:y.visible = x:next
End Sub

Sub UpdateLUT
    Select Case LutImage
        Case 0:table1.ColorGradeImage = "LUT0"
        Case 1:table1.ColorGradeImage = "LUT1"
        Case 2:table1.ColorGradeImage = "LUT2"
        Case 3:table1.ColorGradeImage = "LUT3"
        Case 4:table1.ColorGradeImage = "LUT4"
        Case 5:table1.ColorGradeImage = "LUT5"
        Case 6:table1.ColorGradeImage = "LUT6"
        Case 7:table1.ColorGradeImage = "LUT7"
        Case 8:table1.ColorGradeImage = "LUT8"
        Case 9:table1.ColorGradeImage = "LUT9"
        Case 10:table1.ColorGradeImage = "LUT10"
        Case 11:table1.ColorGradeImage = "LUT Warm 0"
        Case 12:table1.ColorGradeImage = "LUT Warm 1"
        Case 13:table1.ColorGradeImage = "LUT Warm 2"
        Case 14:table1.ColorGradeImage = "LUT Warm 3"
        Case 15:table1.ColorGradeImage = "LUT Warm 4"
        Case 16:table1.ColorGradeImage = "LUT Warm 5"
        Case 17:table1.ColorGradeImage = "LUT Warm 6"
        Case 18:table1.ColorGradeImage = "LUT Warm 7"
        Case 19:table1.ColorGradeImage = "LUT Warm 8"
        Case 20:table1.ColorGradeImage = "LUT Warm 9"
        Case 21:table1.ColorGradeImage = "LUT Warm 10"
    End Select
End Sub

' Thalamus : Exit in a clean and proper way
Sub Table1_exit
  Controller.Pause = False
  Controller.Stop
End Sub


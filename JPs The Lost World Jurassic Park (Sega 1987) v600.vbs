' Sega's The Lost World Jurassic Park / IPD No. 4136 / June 06, 1997 / 6 Players
' VPX8 - version by JPSalas 2024, version 6.0.0
'
' This is JPSalas The Lost World Jurassic Park, similar to Sega's nachine but not exactly the same.
' Game play should be similar, playfield graphics should be similar, plastics are original.

Option Explicit
Randomize

Const BallSize = 50
Const BallMass = 1

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Dim vpmhidden, UseVPMColoredDMD
If Table1.ShowDT = true then
    UseVPMColoredDMD = true
    vpmhidden = 1
Else
    UseVPMColoredDMD = False
    vpmhidden = 0
End If

' Use Modulated Flashers
Const UseVPMModSol = True

LoadVPM "03070000", "SEGA.vbs", 3.61 'needs vpinmame 3.7

Dim bsTrough, dtbank, bsScoop, plungerIM, x, mPFMagnet, mOrbitMagnet, mSmagnet

Const cGameName = "jplstw22"

Const UseSolenoids = 2
Const UseLamps = 1
Const UseGI = 0
Const UseSync = 0   'set it to 1 if the table runs too fast
Const HandleMech = 0

' Standard Sounds
Const SSolenoidOn = "fx_SolenoidOn"
Const SSolenoidOff = "fx_SolenoidOff"
Const SCoin = "fx_Coin"

'******************
' Realtime Updates
'******************

Sub RealTime_Timer
    RollingUpdate
End Sub

' flippers top animations

Sub LeftFlipper_Animate: LeftFlipperTop.RotZ = LeftFlipper.CurrentAngle: End Sub
Sub RightFlipper_Animate: RightFlipperTop.RotZ = RightFlipper.CurrentAngle: End Sub

Sub eggf_Animate: egg.Rotx = eggf.CurrentAngle: End Sub
Sub carf_Animate: car.Rotx = carf.CurrentAngle: End Sub

'************
' Table init.
'************

Sub table1_Init
    vpmInit me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "JP's The Lost World Jurassic Park - Sega, 1997" & vbNewLine & "VPX table by JPSalas v6.0.0"
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = vpmhidden
        .Games(cGameName).Settings.Value("rol") = 0   '1= rotated display, 0= normal
        .Games(cGameName).Settings.Value("sound") = 1 '1 enabled rom sound
        '.SetDisplayPosition 0,0, GetPlayerHWnd 'restore dmd window position
        On Error Resume Next
        Controller.SolMask(0) = 0
        vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the Timer to renable all the solenoids after 2 seconds
        Controller.Run GetPlayerHWnd
        On Error Goto 0
    End With

    ' Nudging
    vpmNudge.TiltSwitch = 56
    vpmNudge.Sensitivity = 5
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingshot)

    ' Trough
    Set bsTrough = New cvpmBallStack
    bsTrough.InitSw 0, 15, 14, 13, 12, 0, 0, 0
    bsTrough.InitKick BallRelease, 00, 8
    bsTrough.InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    bsTrough.Balls = 4

    ' Scoop
    Set bsScoop = New cvpmBallStack
    bsScoop.InitSw 0, 46, 0, 0, 0, 0, 0, 0
    bsScoop.InitKick s46, 200, 15
    bsScoop.InitExitSnd SoundFX("fx_popper", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    bsScoop.KickForceVar = 6

    ' Drop targets
    set dtbank = new cvpmdroptarget
    dtbank.InitDrop s29, 29
    dtbank.initsnd SoundFX("fx_Droptarget", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
    dtbank.CreateEvents "dtbank"

    ' Magnets
    Set mPFMagnet = New cvpmMagnet
    mPFMagnet.InitMagnet Magnet1, 20
    'mPFMagnet.Solenoid=14
    mPFMagnet.GrabCenter = True
    mPFMagnet.CreateEvents "mPFMagnet"

    Set mOrbitMagnet = New cvpmMagnet
    mOrbitMagnet.InitMagnet Magnet2, 70
    mOrbitMagnet.Solenoid = 5
    mOrbitMagnet.GrabCenter = True
    mOrbitMagnet.CreateEvents "mOrbitMagnet"

    ' Impulse Plunger
    Const IMPowerSetting = 40 'Plunger Power
    Const IMTime = 0.6        ' Time in seconds for Full Plunge
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swplunger, IMPowerSetting, IMTime
        .Random 0.5
        .switch 16
        .InitExitSnd SoundFX("fx_plunger", DOFContactors), SoundFX("fx_plunger", DOFContactors)
        .CreateEvents "plungerIM"
    End With

    vpmMapLights aLights

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

    RealTime.Enabled = 1
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
    If vpmKeyDown(keycode)Then Exit Sub
    If keycode = PlungerKey Then Controller.Switch(53) = 1
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If vpmKeyUp(keycode)Then Exit Sub
    If keycode = PlungerKey Then Controller.Switch(53) = 0
End Sub

'*********
' Switches
'*********

' Slings
Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    PlaySoundAt SoundFX("fx_slingshot", DOFContactors), Lemk
    DOF 101, DOFPulse
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 59
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 1:LeftSLing4.Visible = 0:LeftSLing3.Visible = 1:Lemk.RotX = 14
        Case 2:LeftSLing3.Visible = 0:LeftSLing2.Visible = 1:Lemk.RotX = 2
        Case 3:LeftSLing2.Visible = 0:Lemk.RotX = -20:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
    PlaySoundAt SoundFX("fx_slingshot", DOFContactors), Remk
    DOF 102, DOFPulse
    RightSling4.Visible = 1
    Remk.RotX = 26
    RStep = 0
    vpmTimer.PulseSw 62
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 1:RightSLing4.Visible = 0:RightSLing3.Visible = 1:Remk.RotX = 14
        Case 2:RightSLing3.Visible = 0:RightSLing2.Visible = 1:Remk.RotX = 2
        Case 3:RightSLing2.Visible = 0:Remk.RotX = -20:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

' Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 49:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper1:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 50:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper2:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 51:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper3:End Sub

' Drain & Saucers
Sub Drain_Hit:PlaysoundAt "fx_drain", Drain:bsTrough.AddBall Me:End Sub

Dim bBall

Sub s46_Hit
    PlaySoundAt "fx_hole_enter", s46
    Set bBall = ActiveBall:Me.TimerEnabled = 1
    bsScoop.AddBall 0
End Sub

Sub s46_Timer
    If bBall.Z> 0 Then
        bBall.Z = bBall.Z -5
        Exit Sub
    End if
    Me.DestroyBall
    Me.TimerEnabled = 0
End Sub

Dim bBall2

Sub s46a_Hit
    PlaySoundAt "fx_hole_enter", s46a
    Set bBall2 = ActiveBall:Me.TimerEnabled = 1
    vpmtimer.addtimer 2000, "bsScoop.AddBall 0 '"
End Sub

Sub s46a_Timer
    Do While bBall2.Z> 0
        bBall2.Z = bBall2.Z -5
        Exit Sub
    Loop
    Me.DestroyBall
    Me.TimerEnabled = 0
End Sub

' Rollovers
Sub s57_Hit:Controller.Switch(57) = 1:PlaySoundAt "fx_sensor", s57:End Sub
Sub s57_UnHit:Controller.Switch(57) = 0:End Sub

Sub s58_Hit:Controller.Switch(58) = 1:PlaySoundAt "fx_sensor", s58:End Sub
Sub s58_UnHit:Controller.Switch(58) = 0:End Sub

Sub s61_Hit:Controller.Switch(61) = 1:PlaySoundAt "fx_sensor", s61:End Sub
Sub s61_UnHit:Controller.Switch(61) = 0:End Sub

Sub s60_Hit:Controller.Switch(60) = 1:PlaySoundAt "fx_sensor", s60:End Sub
Sub s60_UnHit:Controller.Switch(60) = 0:End Sub

Sub s17_Hit:Controller.Switch(17) = 1:PlaySoundAt "fx_sensor", s17:End Sub
Sub s17_UnHit:Controller.Switch(17) = 0:End Sub

Sub s18_Hit:Controller.Switch(18) = 1:PlaySoundAt "fx_sensor", s18:End Sub
Sub s18_UnHit:Controller.Switch(18) = 0:End Sub

Sub s19_Hit:Controller.Switch(19) = 1:PlaySoundAt "fx_sensor", s19:End Sub
Sub s19_UnHit:Controller.Switch(19) = 0:End Sub

Sub s21_Hit:Controller.Switch(21) = 1:PlaySoundAt "fx_sensor", s21:End Sub
Sub s21_UnHit:Controller.Switch(21) = 0:End Sub

Sub s22_Hit:Controller.Switch(22) = 1:PlaySoundAt "fx_sensor", s22:End Sub
Sub s22_UnHit:Controller.Switch(22) = 0:End Sub

Sub s24_Hit:Controller.Switch(24) = 1:PlaySoundAt "fx_sensor", s24:End Sub
Sub s24_UnHit:Controller.Switch(24) = 0:End Sub

Sub s23_Hit:Controller.Switch(23) = 1:PlaySoundAt "fx_sensor", s23:End Sub
Sub s23_UnHit:Controller.Switch(23) = 0:End Sub

Sub s37_Hit:Controller.Switch(37) = 1:PlaySoundAt "fx_sensor", s37:End Sub
Sub s37_UnHit:Controller.Switch(37) = 0:End Sub

Sub s47_Hit:Controller.Switch(47) = 1:PlaySoundAt "fx_sensor", s47:End Sub
Sub s47_UnHit:Controller.Switch(47) = 0:End Sub

Sub s48_Hit:Controller.Switch(48) = 1:PlaySoundAt "fx_sensor", s48:End Sub
Sub s48_UnHit:Controller.Switch(48) = 0:End Sub

' Spinners
Sub Spinner1_Spin:vpmTimer.PulseSw 33:PlaySoundAt "fx_spinner", Spinner1:End Sub
Sub Spinner2_Spin:vpmTimer.PulseSw 34:PlaySoundAt "fx_spinner", Spinner2:End Sub

'Targets
Sub s9_Hit:vpmTimer.PulseSw 9:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub s30_Hit:vpmTimer.PulseSw 30:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub s31_Hit:vpmTimer.PulseSw 31:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub s32_Hit:vpmTimer.PulseSw 32:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub s38_Hit:vpmTimer.PulseSw 38:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub s39_Hit:vpmTimer.PulseSw 39:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub s40_Hit:vpmTimer.PulseSw 40:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub s41_Hit:vpmTimer.PulseSw 41:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub s42_Hit:vpmTimer.PulseSw 42:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub s43_Hit:vpmTimer.PulseSw 43:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub s44_Hit:vpmTimer.PulseSw 44:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub s45_Hit:vpmTimer.PulseSw 45:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub

'*********
'Solenoids
'*********

SolCallback(1) = "bsTrough.SolOut"
SolCallback(2) = "Auto_Plunger"
Solcallback(3) = "dtbank.SolDropUp"
SolCallback(4) = "bsScoop.SolOut"

'SolCallback(5)= top magnet

' SolCallback(6) = "SolShaker"

'SolCallBack(7) ="SolSmagnet" 'snagger up magnet

SolCallback(8) = "vpmSolSound SoundFX(""fx_Knocker"",DOFKnocker),"

'SolCallback(9)="vpmSolSound SoundFX(""jet3""),"
'SolCallback(10)="vpmSolSound SoundFX(""jet3""),"
'SolCallback(11)="vpmSolSound SoundFX(""jet3""),"
'SolCallback(12)="vpmSolSound SoundFX(""Sling""),"
'SolCallback(13)="vpmSolSound SoundFX(""Sling""),"

SolCallback(14) = "SolPFMagnet"

SolCallback(18) = "SolTruckUp"
SolCallback(19) = "SolTruckDown1"
SolCallBack(20) = "dtbank.SolHit 1,"
SolCallback(21) = "SolDinoEgg"

If UseVPMModSol Then
    SolModCallback(17) = "Flasher17"
    SolModCallback(22) = "Flasher22"
    SolModCallback(23) = "Flasher23"
    SolModCallback(25) = "Flasher25"
    SolModCallback(26) = "Flasher26"
    SolModCallback(27) = "Flasher27"
    SolModCallback(28) = "Flasher28"
    SolModCallback(29) = "Flasher29"
    SolModCallback(30) = "Flasher30"
    SolModCallback(31) = "Flasher31"
    SolModCallback(32) = "Flasher32"
    f17l.Fader = 0
    f22.Fader = 0:f22a.Fader = 0
    f23.Fader = 0:f23a.Fader = 0:f23b.Fader = 0
    f25.Fader = 0
    f26.Fader = 0:f26a.Fader = 0
    f27l.Fader = 0
    f28.Fader = 0:f28a.Fader = 0
    f29.Fader = 0:f29a.Fader = 0
    f30l.Fader = 0
    f31.Fader = 0:f31a.Fader = 0
    f32.Fader = 0:f32a.Fader = 0
Else
    SolCallback(17) = "vpmFlasher f17l,"
    SolCallback(22) = "vpmFlasher Array(F22,f22a),"
    SolCallback(23) = "vpmFlasher Array(F23,f23a,f23b),"
    SolCallback(25) = "vpmFlasher f25,"
    SolCallback(26) = "vpmFlasher Array(F26,f26a),"
    SolCallback(27) = "vpmFlasher f27l,"
    SolCallback(28) = "vpmFlasher Array(F28,f28a),"
    SolCallback(29) = "vpmFlasher Array(F29,f29a),"
    SolCallback(30) = "vpmFlasher f30l,"
    SolCallback(31) = "vpmFlasher Array(F31,f31a),"
    SolCallback(32) = "vpmFlasher Array(F32,f32a),"
    f17l.Fader = 2
    f22.Fader = 2:f22a.Fader = 2
    f23.Fader = 2:f23a.Fader = 2:f23b.Fader = 2
    f25.Fader = 2
    f26.Fader = 2:f26a.Fader = 2
    f27l.Fader = 2
    f28.Fader = 2:f28a.Fader = 2
    f29.Fader = 2:f29a.Fader = 2
    f30l.Fader = 2
    f31.Fader = 2:f31a.Fader = 2
    f32.Fader = 2:f32a.Fader = 2
End If

Sub Flasher17(m):m = m / 255:f17l.State = m:End Sub
Sub Flasher22(m):m = m / 255:f22.State = m:f22a.State = m:End Sub
Sub Flasher23(m):m = m / 255:f23.State = m:f23a.State = m:f23b.State = m:End Sub
Sub Flasher25(m):m = m / 255:f25.State = m:End Sub
Sub Flasher26(m):m = m / 255:f26.State = m:f26a.State = m:End Sub
Sub Flasher27(m):m = m / 255:f27l.State = m:End Sub
Sub Flasher28(m):m = m / 255:f28.State = m:f28a.State = m:End Sub
Sub Flasher29(m):m = m / 255:f29.State = m:f29a.State = m:End Sub
Sub Flasher30(m):m = m / 255:f30l.State = m:End Sub
Sub Flasher31(m):m = m / 255:f31.State = m:f31a.State = m:End Sub
Sub Flasher32(m):m = m / 255:f32.State = m:f32a.State = m:End Sub



Sub Auto_Plunger(Enabled)
    If Enabled Then
        PlungerIM.AutoFire
    End If
End Sub

Sub SolShaker(Enabled)
    If Enabled then PlaySoundAt "fx_Shaker", l16
End Sub

Sub SolPFMagnet(enabled)
    If enabled Then
        mPFMagnet.MagnetOn = True
        MagnetKicker.Enabled = True
    Else
        mPFMagnet.MagnetOn = False
        MagnetKicker.Enabled = False
    End If
End Sub

Sub SolSmagnet(Enabled)
    If Enabled then
        MagnetKicker.Kick 356, 20
    End If
End Sub

'Truck up/down
controller.switch(35) = 1

Sub SolTruckUp(Enabled)
    If Enabled Then
        PlaySoundAt "fx_Solenoidon", s46a
        carf.RotatetoEnd
        carwall.IsDropped = 1
        controller.switch(35) = 0
        controller.switch(36) = 1
        vpmtimer.addtimer 1000, "MagnetKicker.Kick 356, 20 '"
    End IF
End Sub

Sub SolTruckDown1(enabled) 'part 1
    If enabled then
        vpmtimer.addtimer 2000, "SolTruckDown '"
    End if
End Sub

Sub SolTruckDown 'part 2
    PlaySoundAt "fx_Solenoidoff", s46a
    carf.RotatetoStart
    carwall.IsDropped = 0
    controller.switch(35) = 1
    controller.switch(36) = 0
End Sub

Sub SolDinoEgg(Enabled)
    If Enabled then
        PlaySoundAt "fx_Solenoidon", s46
        eggf.RotatetoEnd
    Else
        PlaySoundAt "fx_Solenoidoff", s46
        eggf.RotatetoStart
    End If
End Sub

'*******************
' Flipper Subs Rev3
'*******************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("fx_flipperup", DOFContactors), LeftFlipper
        LeftFlipper.RotateToEnd
        LeftFlipperOn = 1
    Else
        PlaySoundAt SoundFX("fx_flipperdown", DOFContactors), LeftFlipper
        LeftFlipper.RotateToStart
        LeftFlipperOn = 0
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("fx_flipperup", DOFContactors), RightFlipper
        RightFlipper.RotateToEnd
        RightFlipperOn = 1
    Else
        PlaySoundAt SoundFX("fx_flipperdown", DOFContactors), RightFlipper
        RightFlipper.RotateToStart
        RightFlipperOn = 0
    End If
End Sub

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

'************************************
' Diverse Collection Hit Sounds v3.0
'************************************

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
Const maxvel = 45 'max ball velocity
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

'*************************
' GI - needs new vpinmame
'*************************

Set GICallback = GetRef("GIUpdate")

Sub GIUpdate(no, Enabled)
If Enabled then
PlaySound"fx_Gion"
Else
PlaySound"fx_GiOff"
End If
    For each x in aGiLights
        x.State = ABS(Enabled)
    Next
End Sub

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

    ' Desktop DMD
    x = Table1.Option("Desktop DMD", 0, 1, 1, 1, 0, Array("Hide", "Show") )
    DesktopDMD.visible = x

    ' Cabinet rails
    x = Table1.Option("Cabinet Rails", 0, 1, 1, 1, 0, Array("Hide", "Show") )
    For each y in aRails:y.visible = x:next

    ' Side Blades
    x = Table1.Option("Side Blades", 0, 1, 1, 1, 0, Array("Hide", "Show") )
    For each y in aSideBlades:y.SideVisible = x:next
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

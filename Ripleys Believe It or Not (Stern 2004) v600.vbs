' Ripley's Believe or not - Stern 2004
' a JPSalas VPX8 table, version 6.0.0

Option Explicit
Randomize

Const Ballsize = 50
Const Ballmass = 1

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

' Language Roms

Const cGameName = "ripleys" 'English
' Const cGameName = "ripleysf" 'French
' Const cGameName = "ripleysg" 'German
' Const cGameName = "ripleysi" 'Italian
' Const cGameName = "ripleysl" 'Spanish

Dim VarHidden, UseVPMColoredDMD
If Table1.ShowDT = true then
    UseVPMColoredDMD = true
    VarHidden = 1
else
    UseVPMColoredDMD = False
    VarHidden = 0
end if

Const UseVPMModSol = True  'needs vpinmame 3.7

LoadVPM "01550000", "SEGA.VBS", 3.26

'********************
'Standard definitions
'********************

Const UseSolenoids = 2
Const UseLamps = 0
Const UseGI = 0
Const UseSync = 0
Const HandleMech = 0

' Standard Sounds
Const SSolenoidOn = "fx_Solenoidon"
Const SSolenoidOff = "fx_SolenoidOff"
Const SCoin = "fx_coin"

Dim bsTrough, mIdolMag, mShrunkenMag, mUp, mDown, bsLock, bsSkill, bsVUK, plungerIM, x

'******************
' RealTime Updates
'******************

Sub RealTime_Timer
    RollingUpdate
End Sub

' flippers top animations

Sub LeftFlipper_Animate:LeftFlipperTop.RotZ = LeftFlipper.CurrentAngle: End Sub
Sub RightFlipper_Animate: RightFlipperTop.RotZ = RightFlipper.CurrentAngle: End Sub
Sub RightFlipper1_Animate: RightFlipperTop1.RotZ = RightFlipper1.CurrentAngle: End Sub

'************
' Table init.
'************

Sub table1_Init
    vpmInit Me
    With Controller
        .GameName = cGameName
        .Games(cGameName).Settings.Value("sound") = 1 '0 = disable  rom sound
        .SplashInfoLine = "Ripleys Believe It Or Not - Stern 2004" & vbNewLine & "VPX table by JPSalas v6.0.0"
        .Games(cGameName).Settings.Value("rol") = 0   '1= rotated display, 0= normal
        .Games(cGameName).Settings.Value("sound") = 1 '1 enabled rom sound
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .HandleKeyboard = 0
        .Hidden = VarHidden
        .Switch(42) = 1
        .Switch(43) = 1
        On Error Resume Next
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
        On Error Goto 0
    End With

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

    vpmNudge.TiltSwitch = 56
    vpmNudge.Sensitivity = 1
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, Bumper4, Bumper5, Bumper6, LeftSlingshot, RightSlingshot)

    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0, 14, 13, 12, 11, 0, 0, 0
        .InitKick BallRelease, 100, 3
        .InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .Balls = 4
    End With

    Set mIdolMag = New cvpmMagnet
    With mIdolMag
        .InitMagnet IMagnet, 50
        .Solenoid = 19
        .GrabCenter = 0
    End With

    Set mShrunkenMag = New cvpmMagnet
    With mShrunkenMag
        .InitMagnet SMagnet, 70
        .GrabCenter = 1
    End With

    Set bsLock = New cvpmBallStack
    With bsLock
        .InitSw 0, 46, 45, 44, 0, 0, 0, 0
        .InitKick Lock, 0, 36
        .KickForceVar = 4
        .InitExitSnd SoundFX("fx_popper", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    End With

    Set bsSkill = new cvpmBallStack
    With bsSkill
        .InitSw 0, 29, 0, 0, 0, 0, 0, 0
        .InitKick sw29a, 180, 16
        .KickZ = 1.12
        .InitExitSnd "fx_popper", "fx_Solenoid"
    End With

    Set bsVUK = new cvpmBallStack
    With bsVUK
        .InitSw 0, 52, 0, 0, 0, 0, 0, 0
        .InitKick sw52, 0, 36
        .InitExitSnd SoundFX("fx_popper", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .KickZ = 1.5
    End With

    ' Impulse Plunger
    Const IMPowerSetting = 44 ' Plunger Power
    Const IMTime = 0.7        ' Time in seconds for Full Plunge
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swPlunger, IMPowerSetting, IMTime
        .Random 0.3
        .switch 16
        .InitExitSnd "fx_plunger2", "fx_plunger"
        .CreateEvents "plungerIM"
    End With

    RealTime.Enabled = 1
End Sub

Sub Table1_Paused:Controller.Pause = 1:End Sub
Sub Table1_unPaused:Controller.Pause = 0:End Sub
Sub Table1_Exit:Controller.Stop:End Sub

'****
'Keys
'****

Sub Table1_KeyDown(ByVal keycode)
    If keycode = LeftTiltKey Then Nudge 90, 8:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 8:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 9:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If KeyDownHandler(KeyCode)Then Exit Sub
    If keycode = PlungerKey Then PlaySoundAt "fx_PlungerPull", Plunger:Plunger.Pullback
End Sub

Sub Table1_KeyUp(ByVal keycode)
    If KeyUpHandler(KeyCode)Then Exit Sub
    If keycode = PlungerKey Then PlaySoundAt "fx_plunger", Plunger:Plunger.Fire
End Sub

'*********
'Solenoids
'*********

SolCallBack(1) = "SolTrough"
SolCallBack(2) = "Auto_Plunger"
SolCallBack(3) = "bsVUK.SolOut"
SolCallBack(4) = "vpmSolDiverter TempleDiv,1,"
SolCallBack(5) = "vpmSolDiverter LockDiverter,1,"
' 6, 7, 8, 9, 10 11 pop bumpers
SolCallBack(12) = "bsSkill.SolOut"
SolCallback(13) = "bsLock.SolOut"
SolCallback(20) = "SolUpperMagnet"
SolCallBack(21) = "SolVReset"

SolCallBack(23) = "SolPost"
SolCallBack(24) = "vpmsolsound SoundFX(""fx_knocker"",DOFKnocker),"

If UseVPMModSol Then
SolModCallback(22) = "Flasher22"
SolModCallback(25) = "Flasher25"
SolModCallback(26) = "Flasher26"
SolModCallback(27) = "Flasher27"
SolModCallback(28) = "Flasher28"
SolModCallback(29) = "Flasher29"
SolModCallback(30) = "Flasher30"
SolModCallback(31) = "Flasher31"
SolModCallback(32) = "Flasher32"
f22.fader = 0
f22a.fader = 0
f25.fader = 0
f25c.fader = 0
f26l.fader = 0
f27l.fader = 0
f28a.fader = 0
f29a.fader = 0
f30.fader = 0
f31l.fader = 0
f32a.fader = 0
Else
SolCallback(22) = "vpmFlasher Array(f22,f22a),"
SolCallback(25) = "vpmFlasher Array(f25,f25c),"
SolCallback(26) = "vpmFlasher f26l,"
SolCallback(27) = "vpmFlasher f27l,"
SolCallback(28) = "vpmFlasher f28a,"
SolCallback(29) = "vpmFlasher f29a,"
SolCallback(30) = "vpmFlasher f30,"
SolCallback(31) = "vpmFlasher f31l,"
SolCallback(32) = "vpmFlasher f32a,"
f22.fader = 2
f22a.fader = 2
f25.fader = 2
f25c.fader = 2
f26l.fader = 2
f27l.fader = 2
f28a.fader = 2
f29a.fader = 2
f30.fader = 2
f31l.fader = 2
f32a.fader = 2
End If

Sub Flasher22(m):m = m /255:f22.State = m:f22a.State = m:End Sub
Sub Flasher25(m):m = m /255:f25.State = m:f25c.State = m:End Sub
Sub Flasher26(m):m = m /255:f26l.State = m:End Sub
Sub Flasher27(m):m = m /255:f27l.State = m:End Sub
Sub Flasher28(m):m = m /255:f28a.State = m:End Sub
Sub Flasher29(m):m = m /255:f29a.State = m:End Sub
Sub Flasher30(m):m = m /255:f30.State = m:End Sub
Sub Flasher31(m):m = m /255:f31l.State = m:End Sub
Sub Flasher32(m):m = m /255:f32a.State = m:End Sub

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
        RightFlipper1.RotateToEnd
        RightFlipperOn = 1
    Else
        PlaySoundAt SoundFX("fx_flipperdown", DOFContactors), RightFlipper
        RightFlipper.RotateToStart
        RightFlipper1.RotateToStart
        RightFlipperOn = 0
    End If
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper1_Collide(parm)
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
FullStrokeEOS_Torque = 0.6  ' EOS Torque when flipper hold up ( EOS Coil is fully charged. Ampere increase due to flipper can't move or when it pushed back when "On". EOS Coil have more power )
LiveStrokeEOS_Torque = 0.3  ' EOS Torque when flipper rotate to end ( When flipper move, EOS coil have less Ampere due to flipper can freely move. EOS Coil have less power )

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

'**************
' Solenoid Subs
'**************

Sub Auto_Plunger(Enabled)
    If Enabled Then
        PlungerIM.AutoFire
    End If
End Sub

Sub SolVReset(Enabled)
    If Enabled Then
        VariTimerDown.Enabled = 1
    End If
End Sub

Sub SolPost(Enabled)
    TopPost.IsDropped = NOT Enabled
End Sub

Sub SolTrough(Enabled)
    If Enabled Then
        bsTrough.ExitSol_On
        vpmTimer.PulseSw 15
    End If
End Sub

'**********************
' Shrunken Head Magnet
'**********************

Sub SMagnet_Hit
    Controller.Switch(23) = 1
    mShrunkenMag.AddBall ActiveBall
End Sub

Sub SMagnet_unHit
  Controller.Switch(23) = 0
  mShrunkenMag.RemoveBall ActiveBall
End Sub

Sub SolUpperMagnet(Enabled)
  Dim ball
    If Enabled Then
        mShrunkenMag.MagnetOn = 1
    Else
        mShrunkenMag.MagnetOn = 0
    For Each ball in mShrunkenMag.Balls
      ball.VelY = -19
    Next
    End If
End Sub

' ************************************
' Switches, bumpers, lanes and targets
' ************************************

Sub Drain_Hit:PlaysoundAt "fx_drain", drain:bsTrough.AddBall Me:End Sub

Sub sw9_Hit:vpmTimer.PulseSw 9:PlaySoundAt SoundFX("fx_target", DOFTargets), sw9:End Sub
Sub sw17_Hit:vpmTimer.PulseSw 17:PlaySoundAt SoundFX("fx_target", DOFTargets), sw17:End Sub
Sub sw18_Hit:Controller.Switch(18) = 1:ActiveBall.VelX = - 1:ActiveBall.VelY = 1:End Sub
Sub sw18_unHit:Controller.Switch(18) = 0:End Sub
Sub sw19_Hit:vpmTimer.PulseSw 19:PlaySoundAt SoundFX("fx_target", DOFTargets), sw19:End Sub
Sub sw20_Spin:PlaySoundAt "fx_spinner", sw20:vpmTimer.PulseSw 20:End Sub
Sub sw21_Spin:PlaySoundAt "fx_spinner", sw21:vpmTimer.PulseSw 21:End Sub
Sub sw22_Hit:vpmTimer.PulseSw 22:PlaySoundAt SoundFX("fx_target", DOFTargets), sw22:End Sub

Sub IMagnet_Hit:mIdolMag.AddBall ActiveBall:Controller.Switch(24) = 1:End Sub
Sub IMagnet_unHit:mIdolMag.RemoveBall ActiveBall:Controller.Switch(24) = 0:End Sub

'left bumpers
Sub Bumper4_Hit:vpmTimer.PulseSw 25:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper4:End Sub
Sub Bumper5_Hit:vpmTimer.PulseSw 26:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper5:End Sub
Sub Bumper6_Hit:vpmTimer.PulseSw 27:PlaySoundAt SoundFX("fx_bumper", DOFContactors),Bumper6:End Sub

Sub sw28_Hit:sw28.DestroyBall:PlaySoundAt "fx_hole_enter", sw28:vpmTimer.PulseSwitch(28), 100, "AddToSkill":End Sub
Sub AddToSkill(swNo):bsSkill.AddBall 0:End Sub
Sub sw29_Hit:PlaySoundAt "fx_hole_enter", sw29:bsSkill.AddBall Me:End Sub

Sub sw30_Hit:Controller.Switch(30) = 1:PlaySoundAt "fx_sensor", sw30:End Sub
Sub sw30_unHit:Controller.Switch(30) = 0:End Sub
Sub sw31_Hit:Controller.Switch(31) = 1:End Sub
Sub sw31_unHit:Controller.Switch(31) = 0:End Sub
Sub sw32_Hit:vpmTimer.PulseSw 32:PlaySoundAt SoundFX("fx_target", DOFTargets), sw32:End Sub
Sub sw32a_Hit:vpmTimer.PulseSw 32:PlaySoundAt SoundFX("fx_target", DOFTargets), sw32a:End Sub
Sub sw33_Hit:Controller.Switch(33) = 1:PlaySoundAt "fx_sensor", sw33:End Sub
Sub sw33_unHit:Controller.Switch(33) = 0:End Sub
Sub sw34_Hit:Controller.Switch(34) = 1:PlaySoundAt "fx_sensor", sw34:End Sub
Sub sw34_unHit:Controller.Switch(34) = 0:End Sub
Sub sw35_Hit:Controller.Switch(35) = 1:PlaySoundAt "fx_sensor", sw35:End Sub
Sub sw35_unHit:Controller.Switch(35) = 0:End Sub
Sub sw36_Hit:Controller.Switch(36) = 1:PlaySoundAt "fx_sensor", sw36:End Sub
Sub sw36_unHit:Controller.Switch(36) = 0:End Sub
Sub sw38_Hit:Controller.Switch(38) = 1:PlaySoundAt "fx_sensor", sw38:End Sub
Sub sw38_unHit:Controller.Switch(38) = 0:End Sub
Sub sw39_Hit:Controller.Switch(39) = 1:PlaySoundAt "fx_sensor", sw39:End Sub
Sub sw39_unHit:Controller.Switch(39) = 0:End Sub
Sub sw40_Hit:Controller.Switch(40) = 1:PlaySoundAt "fx_sensor", sw40:End Sub
Sub sw40_unHit:Controller.Switch(40) = 0:End Sub
Sub Lock_Hit:bsLock.AddBall Me:PlaySoundAt "fx_kicker_enter", lock:End Sub
Sub sw47_Hit:Controller.Switch(47) = 1:PlaySoundAt "fx_sensor", sw47:End Sub
Sub sw47_unHit:Controller.Switch(47) = 0:End Sub
Sub sw48_Hit:Controller.Switch(48) = 1:PlaySoundAt "fx_sensor", sw48:End Sub
Sub sw48_unHit:Controller.Switch(48) = 0:End Sub

Sub Bumper1_Hit:vpmTimer.PulseSw 49:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper1:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 50:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper2:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 51:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper3:End Sub

'************
' Varitarget
'************

Sub sw42_Hit
    If ActiveBall.VelY < 0 Then
        Controller.Switch(42) = 1
        VariTarget.RotY = 8
    End If
End Sub

Sub sw43_Hit
    If ActiveBall.VelY < 0 Then
        Controller.Switch(43) = 1
        VariTarget.RotY = 1
    End If
End Sub

Sub sw41_Hit
    If ActiveBall.VelY < 0 Then
        Controller.Switch(41) = 1
        VariTarget.RotY = -7
    End If
End Sub

Sub VariTimerDown_Timer
    VariTarget.RotY = VariTarget.RotY + 1
    If VariTarget.RotY = -5 Then Controller.Switch(41) = 0
    If VariTarget.RotY = 5 Then Controller.Switch(43) = 0
    If VariTarget.RotY > 15 Then
        VariTarget.RotY = 15
        Controller.Switch(42) = 0
        VariTimerDown.Enabled = 0
    End If
End Sub

'******
' vuk
'******

Sub sw52_Hit():bsVUK.AddBall Me:PlaySoundAt "fx_kicker_enter", sw52:End Sub

'*******
' lanes
'*******

Sub sw53_Hit:Controller.Switch(53) = 1:PlaySoundAt "fx_sensor", sw53:End Sub
Sub sw53_unHit:Controller.Switch(53) = 0:End Sub

Sub sw57_Hit:Controller.Switch(57) = 1:PlaySoundAt "fx_sensor", sw57:End Sub
Sub sw57_UnHit:Controller.Switch(57) = 0:End Sub

Sub sw58_Hit:Controller.Switch(58) = 1:PlaySoundAt "fx_sensor", sw58:End Sub
Sub sw58_UnHit:Controller.Switch(58) = 0:End Sub

Sub sw61_Hit:Controller.Switch(61) = 1:PlaySoundAt "fx_sensor", sw61:End Sub
Sub sw61_UnHit:Controller.Switch(61) = 0:End Sub

Sub sw60_Hit:Controller.Switch(60) = 1:PlaySoundAt "fx_sensor", sw60:End Sub
Sub sw60_UnHit:Controller.Switch(60) = 0:End Sub

'************
' Slingshots
'************

Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    PlaySoundAt SoundFX("fx_slingshot", DOFContactors), Lemk
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
        Case 3:LeftSLing2.Visible = 0:Lemk.RotX = -10:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
    PlaySoundAt SoundFX("fx_slingshot", DOFContactors), Remk
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
        Case 3:RightSLing2.Visible = 0:Remk.RotX = -10:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

'**************
' LED's display
'**************

Dim LED(14)

LED(0) = Array(LED1, LED2, LED3, LED4, LED5, LED6, LED7)
LED(1) = Array(LED8, LED9, LED10, LED11, LED12, LED13, LED14)
LED(2) = Array(LED15, LED16, LED17, LED18, LED19, LED20, LED21)
LED(3) = Array(LED22, LED23, LED24, LED25, LED26, LED27, LED28)
LED(4) = Array(LED29, LED30, LED31, LED32, LED33, LED34, LED35)

LED(5) = Array(LED36, LED37, LED38, LED39, LED40, LED41, LED42)
LED(6) = Array(LED43, LED44, LED45, LED46, LED47, LED48, LED49)
LED(7) = Array(LED50, LED51, LED52, LED53, LED54, LED55, LED56)
LED(8) = Array(LED57, LED58, LED59, LED60, LED61, LED62, LED63)
LED(9) = Array(LED64, LED65, LED66, LED67, LED68, LED69, LED70)

LED(10) = Array(LED71, LED72, LED73, LED74, LED75, LED76, LED77)
LED(11) = Array(LED78, LED79, LED80, LED81, LED82, LED83, LED84)
LED(12) = Array(LED85, LED86, LED87, LED88, LED89, LED90, LED91)
LED(13) = Array(LED92, LED93, LED94, LED95, LED96, LED97, LED98)
LED(14) = Array(LED99, LED100, LED101, LED102, LED103, LED104, LED105)

Sub UpdateLeds()
    Dim ChgLED, ii, num, chg, stat, obj
    ChgLED = Controller.ChangedLEDs(&H00000000, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
        For ii = 0 To UBound(chgLED)
            num = chgLED(ii, 0):chg = chgLED(ii, 1):stat = chgLED(ii, 2)
            For Each obj In LED(num)
                If chg And 1 Then obj.Visible = stat And 1
                chg = chg \ 4:stat = stat \ 4
            Next
        Next
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
    Lamp 4, l4
    Lamp 5, l5
    Lamp 6, l6
    Lamp 7, l7
    Lamp 8, l8
    Lamp 9, l9
    Lamp 10, l10
    Lamp 11, l11
    Lamp 12, l12
    Lamp 13, l13
    Lamp 14, l14
    Lamp 15, l15
    Lamp 16, l16b
    Lamp 17, l17
    Lamp 18, l18
    Lamp 19, l19
    Lamp 20, l20
    Lamp 21, l21
    Lamp 22, l22
    Lamp 23, l23
    Lamp 24, l24
    Lamp 25, l25
    Lamp 26, l26
    Lamp 27, l27
    Lamp 28, l28
    Lamp 29, l29
    Lamp 30, l30
    Lamp 31, l31
    Lamp 32, l32
    Lampm 33, BumperLight4a
    Lamp 33, BumperLight4
    Lampm 34, BumperLight5a
    Lamp 34, BumperLight5
    Lampm 35, BumperLight6a
    Lamp 35, BumperLight6
    Lamp 36, l36
    Lamp 37, l37
    Lamp 38, l38
    Lamp 39, l39
    Lamp 40, l40
    Lampm 41, l41
    Lamp 41, l41b
    Lampm 42, l42
    Lamp 42, l42b
    Lampm 43, l43
    Lamp 43, l43b
    Lampm 44, l44
    Lamp 44, l44b
    Lampm 45, l45
    Lamp 45, l45b
    Lampm 46, l46
    Lamp 46, l46b
    Lampm 47, l47
    Lamp 47, l47b
    Lampm 48, l48
    Lamp 48, l48b
    Lamp 49, l49
    Lamp 50, l50
    Lamp 51, l51
    Lamp 52, l52
    Lamp 53, l53
    Lamp 54, l54
    Lamp 55, l55
    Lampm 56, l56
    Lamp 56, l56b
    Lamp 57, l57
    Lamp 58, l58
    Lamp 59, l59
    Lampm 60, BumperLight1a
    Lamp 60, BumperLight1
    Lampm 61, BumperLight2a
    Lamp 61, BumperLight2
    Lampm 62, BumperLight3a
    Lamp 62, BumperLight3
    Lamp 63, l63
    Lampm 64, l64
    Lamp 64, l64b
    Lamp 65, l65
    Lamp 66, l66
    Lamp 67, l67
    Lamp 68, l68
    Lamp 69, l69
    Lamp 70, l70
    Lamp 71, l71
    Lamp 72, l72
    Lamp 73, l73
    Lamp 74, l74
    Lamp 75, l75
    Lamp 76, l76l
    Lamp 77, l77l
    Lamp 78, l78l
    'Lamp 80, l80
End Sub

' div lamp subs

' Normal Lamp & Flasher subs

Sub InitLamps()
    Dim x
    LampTimer.Interval = 20
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
                ballvol = Vol(BOT(b)) * 5
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
If Enabled Then
PlaySound"fx_gion"
Else
PlaySound"fx_gioff"
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

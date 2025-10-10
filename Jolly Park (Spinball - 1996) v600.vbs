' Spinball's Jolly Park / IPD No. 4618 / 1996 / 4 Players
' VPX8 - version by JPSalas 2024, version 6.0.0
' Thanks to destrUk & TAB for their old table. most vpinmame scripting are taken from their table.

Option Explicit
Randomize

Dim DesktopMode:DesktopMode = Table1.ShowDT
Dim UseVPMColoredDMD

UseVPMColoredDMD = DesktopMode

Const BallSize = 50
Const BallMass = 1

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01530000", "spinball.vbs", 3.10

Dim bsTrough, dtR, bsVUK, bsLock, cbCaptive, cbCaptive1, LeftMag, RightMag, plungerIM, x

Const cGameName = "jolypark"

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

' Standard Sounds
Const SSolenoidOn = "fx_Solenoidon"
Const SSolenoidOff = "fx_Solenoidoff"
Const SCoin = "fx_Coin"

'************
' Table init.
'************

Sub table1_Init
    vpmInit me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Jolly Park - Spinball 1996" & vbNewLine & "VPX table by JPSalas v6.0.0"
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = VarHidden
        .Games(cGameName).Settings.Value("rol") = 0 '1= rotated display, 0= normal
        .Games(cGameName).Settings.Value("sound") = 1 '1 enabled rom sound
        '.SetDisplayPosition 0,0, GetPlayerHWnd 'restore dmd window position
        On Error Resume Next
        Controller.SolMask(0) = 0
        vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the Timer to renable all the solenoids after 2 seconds
        Controller.Run GetPlayerHWnd
        On Error Goto 0
    End With

    ' Nudging
    vpmNudge.TiltSwitch = 33
    vpmNudge.Sensitivity = 3
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, Bumper4, LeftSlingshot)

    ' Drop targets
    set dtR = new cvpmdroptarget
    With dtR
        .initdrop array(sw44, sw43, sw42, sw41), array(44, 43, 42, 41)
        .initsnd SoundFX("fx_droptarget", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
        .CreateEvents "dtR"
    End With

    Set LeftMag = New cvpmMagnet
    With LeftMag
        .InitMagnet Magnet1, 6
        .Solenoid = 9
        .GrabCenter = 0
        .CreateEvents "LeftMag"
    End With

    Set RightMag = New cvpmMagnet
    With RightMag
        .InitMagnet Magnet2, 6
        .Solenoid = 12
        .GrabCenter = 0
        .CreateEvents "RightMag"
    End With

    ' Thalamus, more randomness pls
    Set bsVUK = New cvpmBallStack
    With bsVUK
        .InitSaucer sw50, 50, 0, 55
        .KickZ = 1.5
        .InitExitSnd SoundFX("fx_popper", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .KickForceVar = 3
        .KickAngleVar = 3
    End With

    Set bsLock = New cvpmBallStack
    With bsLock
        .InitSw 0, 67, 0, 0, 0, 0, 0, 0
        .InitKick sw67, 270, 8
        .InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .KickForceVar = 3
        .KickAngleVar = 3
    End With

    Set cbCaptive = New cvpmCaptiveBall
    With cbCaptive
        .InitCaptive captivetrigger, captivewall, captiveKicker, 0
        .Start
        .ForceTrans = .9
        .MinForce = 3.5
        .CreateEvents "cbCaptive"
    End With

    Set cbCaptive1 = New cvpmCaptiveBall
    With cbCaptive1
        .InitCaptive captivetrigger1, captivewall1, captiveKicker1, 0
        .RestSwitch = 60
        .Start
        .ForceTrans = .9
        .MinForce = 3.5
        .CreateEvents "cbCaptive1"
    End With
    captiveKicker1.Destroyball 'as this will be a locked ball
    Controller.Switch(60) = 0

    'Impulse Plunger
    Const IMPowerSetting = 48 ' Plunger Power
    Const IMTime = 1.1        ' Time in seconds for Full Plunge
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP sw74, IMPowerSetting, IMTime
        .Switch 74
        .Random 5
        .InitExitSnd SoundFX("fx_kicker",DOFContactors), "fx_solenoidoff"
        .CreateEvents "plungerIM"
    End With

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

    ' RealTimer timer
    RealTime.Enabled = 1

    ' Init walls
    LeftSling4.IsDropped = 1:LeftSLing3.IsDropped = 1:LeftSLing2.IsDropped = 1

    ' set in 4 balls
    Controller.Switch(73) = 1:Controller.Switch(72) = 1:Controller.Switch(71) = 1:Controller.Switch(70) = 1
End Sub

'*******************************
' RealTime Animations and checks
'*******************************

Sub RealTime_Timer 'primitive object animations
  RollingUpdate
End Sub

Sub Gate5_Animate: g1.rotx = 20 - Gate5.Currentangle / 4.5: End Sub
Sub Gate4_Animate: g2.rotx = 20 - Gate4.Currentangle / 4.5: End Sub
Sub Gate6_Animate: g3.rotx = 20 - Gate6.Currentangle / 4.5: End Sub

Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub
Sub table1_exit:Controller.stop:End Sub

'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
    If KeyCode = keyInsertCoin1 Or KeyCode = keyInsertCoin2 or KeyCode = keyInsertCoin3 Then Controller.Switch(30) = 1
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If KeyCode = LeftMagnaSave Then Controller.Switch(47) = 1  'Left magnet & volume down
    If KeyCode = RightMagnaSave Then Controller.Switch(57) = 1 'Right magnet & volume up
    If KeyCode = PlungerKey Then Controller.Switch(86) = 1
    If KeyCode = KeyFront Then Controller.Switch(75) = 1       'Add extra ball, Keyfront is usually the key nr "2"
    If vpmKeyDown(KeyCode) Then Exit Sub
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If KeyCode = keyInsertCoin1 Or KeyCode = keyInsertCoin2 or KeyCode = keyInsertCoin3 Then Controller.Switch(30) = 0
    If KeyCode = LeftMagnaSave Then Controller.Switch(47) = 0
    If KeyCode = RightMagnaSave Then Controller.Switch(57) = 0
    If KeyCode = PlungerKey Then Controller.Switch(86) = 0
    If KeyCode = KeyFront Then Controller.Switch(75) = 0
    If vpmKeyUp(KeyCode) Then Exit Sub
End Sub

'*********
' Switches
'*********

' Slings
Dim LStep

Sub LeftSlingShot_Slingshot
    PlaySoundAt SoundFX("fx_slingshot", DOFContactors), li67d 'we use a light for the sound position
    LeftSling1.IsDropped = 1
    LeftSling4.IsDropped = 0
    LStep = 0
    vpmTimer.PulseSw 40
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 1:LeftSling4.IsDropped = 1:LeftSLing3.IsDropped = 0
        Case 2:LeftSLing3.IsDropped = 1:LeftSLing2.IsDropped = 0
        Case 3:LeftSLing2.IsDropped = 1:LeftSling1.IsDropped = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

' Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 76:PlaySoundAt SoundFX("fx_bumper", DOFContactors),Bumper1:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 66:PlaySoundAt SoundFX("fx_bumper", DOFContactors),Bumper2:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 46:PlaySoundAt SoundFX("fx_bumper", DOFContactors),Bumper3:End Sub
Sub Bumper4_Hit:vpmTimer.PulseSw 56:PlaySoundAt SoundFX("fx_bumper", DOFContactors),Bumper4:End Sub

Sub bumper2_Animate: YCar.RotY = bumper2.CurrentRingOffset:End Sub
Sub bumper3_Animate: RCar.RotY = - bumper3.CurrentRingOffset:End Sub
Sub bumper4_Animate: GCar.RotY = bumper4.CurrentRingOffset:End Sub

' Vuk
Sub sw50_Hit:PlaySoundAt "fx_kicker_enter", sw50:bsVUK.AddBall 0:End Sub

' Rollovers
Sub sw54_Hit:Controller.Switch(54) = 1:PlaySoundAt "fx_sensor", sw54:DOF 101, DOFOn:End Sub
Sub sw54_UnHit:Controller.Switch(54) = 0:DOF 101, DOFOff:End Sub

Sub sw54a_Hit:Controller.Switch(54) = 1:PlaySoundAt "fx_sensor", sw54a:DOF 102, DOFOn:End Sub
Sub sw54a_UnHit:Controller.Switch(54) = 0:DOF 102, DOFOff:End Sub

Sub sw95_Hit:Controller.Switch(95) = 1:PlaySoundAt "fx_sensor", sw95:End Sub
Sub sw95_UnHit:Controller.Switch(95) = 0:End Sub

Sub sw94_Hit:Controller.Switch(94) = 1:PlaySoundAt "fx_sensor", sw94:End Sub
Sub sw94_UnHit:Controller.Switch(94) = 0:End Sub

Sub sw93_Hit:Controller.Switch(93) = 1:PlaySoundAt "fx_sensor", sw93:End Sub
Sub sw93_UnHit:Controller.Switch(93) = 0:End Sub

Sub sw96_Hit:Controller.Switch(96) = 1:PlaySoundAt "fx_sensor", sw96:End Sub
Sub sw96_UnHit:Controller.Switch(96) = 0:End Sub

Sub sw92_Hit:Controller.Switch(92) = 1:PlaySoundAt "fx_sensor", sw92:End Sub
Sub sw92_UnHit:Controller.Switch(92) = 0:End Sub

Sub sw64_Hit:Controller.Switch(64) = 1:PlaySoundAt "fx_sensor", sw64:End Sub
Sub sw64_UnHit:Controller.Switch(64) = 0:End Sub

Sub sw64a_Hit:Controller.Switch(64) = 1:PlaySoundAt "fx_sensor", sw64a:End Sub
Sub sw64a_UnHit:Controller.Switch(64) = 0:End Sub

Sub sw87_Hit:Controller.Switch(87) = 1:PlaySoundAt "fx_sensor", sw87:End Sub
Sub sw87_UnHit:Controller.Switch(87) = 0:End Sub

Sub sw77_Hit:Controller.Switch(77) = 1:PlaySoundAt "fx_sensor", sw77:End Sub
Sub sw77_UnHit
  Controller.Switch(77) = 0
  If ActiveBall.Vely > 0 then
    ActiveBall.VelX = -1
  End If
End Sub

Sub sw85_Hit:Controller.Switch(85) = 1:PlaySoundAt "fx_sensor", sw85:End Sub
Sub sw85_UnHit:Controller.Switch(85) = 0:End Sub

Sub sw91_Hit:Controller.Switch(91) = 1:PlaySoundAt "fx_sensor", sw91:End Sub
Sub sw91_UnHit:Controller.Switch(91) = 0:End Sub

Sub sw61_Hit:Controller.Switch(61) = 1:PlaySoundAt "fx_sensor", sw61:End Sub
Sub sw61_UnHit:Controller.Switch(61) = 0:End Sub

Sub sw62_Hit:Controller.Switch(62) = 1:PlaySoundAt "fx_sensor", sw62:End Sub
Sub sw62_UnHit:Controller.Switch(62) = 0:End Sub

Sub sw63_Hit:Controller.Switch(63) = 1:PlaySoundAt "fx_sensor", sw63:End Sub
Sub sw63_UnHit:Controller.Switch(63) = 0:End Sub

Sub sw97_Hit:Controller.Switch(97) = 1:PlaySoundAt "fx_sensor", sw97:End Sub
Sub sw97_UnHit:Controller.Switch(97) = 0:End Sub

Sub sw80_Hit:Controller.Switch(80) = 1:PlaySoundAt "fx_sensor", sw80:End Sub
Sub sw80_UnHit:Controller.Switch(80) = 0:End Sub

Sub sw81_Hit:Controller.Switch(81) = 1:PlaySoundAt "fx_sensor", sw81:End Sub
Sub sw81_UnHit:Controller.Switch(81) = 0:End Sub

'Targets
Sub sw51_Hit:vpmTimer.PulseSw 51:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub sw52_Hit:vpmTimer.PulseSw 52:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub sw53_Hit:vpmTimer.PulseSw 53:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub sw90_Hit:vpmTimer.PulseSw 90:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub

'*********
'Solenoids
'*********

SolCallback(3) = "VpmSolSound SoundFX(""fx_knocker"",DOFKnocker),"
SolCallback(4) = "SolTrough"
'SolCallback(6)="vpmSolSound ""Jet3"","
'SolCallback(7)="vpmSolSound ""Jet3"","
SolCallback(8) = "dtR.SolDropUp"
'SolCallback(9)="Left Magnet"
SolCallback(10) = "SolLockOut"
SolCallback(11) = "Auto_Plunger"
'SolCallback(12)="Right Magnet"
SolCallback(13) = "bsVUK.SolOut"
'SolCallback(14)="vpmSolSound ""Sling"","
'SolCallback(15)="vpmSolSound ""Jet3"","
'SolCallback(16)="vpmSolSound ""Jet3"","
SolCallback(17) = "SolRampTog"
SolCallback(20) = "LockRelease"
SolCallback(25) = "vpmNudge.SolGameOn"

' Trough
Dim Balls
Balls = 4

Sub SolTrough(Enabled)
    If Enabled Then
        If Balls> 0 Then
            Controller.Switch(73) = 0:Controller.Switch(72) = 0:Controller.Switch(71) = 0:Controller.Switch(70) = 0
            ExitTrough.Enabled = 1
            BallRelease.CreateBall
            BallRelease.Kick 110, 6
            PlaySoundAt SoundFX("fx_BallRel",DOFContactors),BallRelease
            Balls = Balls-1
        Else
            PlaySoundAt "fx_Solenoid", BallRelease
        End If
    End If
End Sub

Sub ExitTrough_Timer
    ExitTrough.Enabled = 0
    Select Case Balls
        Case 1:Controller.Switch(73) = 1
        Case 2:Controller.Switch(73) = 1:Controller.Switch(72) = 1
        Case 3:Controller.Switch(73) = 1:Controller.Switch(72) = 1:Controller.Switch(71) = 1
    End Select
End Sub

'Drain

Dim BallCheck
BallCheck = 0

Sub Drain_Hit
    PlaySoundAt "fx_drain", Drain
    Drain.DestroyBall
    BallCheck = BallCheck + 1
    EnterTrough.Enabled = 1
End Sub

Sub EnterTrough_Timer
    BallCheck = BallCheck-1
    If BallCheck = 0 Then EnterTrough.Enabled = 0
    Balls = Balls + 1
    Select Case Balls
        Case 1:Controller.Switch(73) = 1
        Case 2:Controller.Switch(73) = 1:Controller.Switch(72) = 1
        Case 3:Controller.Switch(73) = 1:Controller.Switch(72) = 1:Controller.Switch(71) = 1
        Case 4:Controller.Switch(73) = 1:Controller.Switch(72) = 1:Controller.Switch(71) = 1:Controller.Switch(70) = 1
    End Select
End Sub

' Lock

Sub SolLockOut(Enabled)
    If Enabled Then
        sw67.TimerEnabled = 0
        sw67.TimerEnabled = 1
    End If
End Sub

Sub Lock1_Hit:PlaySoundAt "fx_hole", Lock1:bsLock.AddBall Me:End Sub
Sub Lock2_Hit:PlaySoundAt "fx_hole", Lock2:bsLock.AddBall Me:End Sub
Sub sw67_Hit:PlaySoundAt "fx_hole", sw67:bsLock.AddBall Me:End Sub

Sub sw67_Timer
    If bsLock.Balls> 0 Then
        bsLock.ExitSol_On
    Else
        sw67.TimerEnabled = 0
    End If
End Sub

Sub Auto_Plunger(Enabled)
    If Enabled Then
        PlungerIM.AutoFire
    End If
End Sub

' Ramp toggle
Dim RampDir
RampDir = 1 'Right

Sub SolRampTog(Enabled)
    If Enabled Then
        RampDir = 0
    Else
        RampDir = 1
    End If
End Sub

Sub LockKicker_Hit
    If RampDir = 1 And Not Controller.Switch(60) Then
        ActiveBall.VelX = 5
    Else
        ActiveBall.VelX = -5
    End If
End Sub

' 3 ball Lock

Dim LA, LB, LC
LA = 0:LB = 0:LC = 0
Dim LockOpen
LockOpen = 0

'Ball Lock

Sub sw45_Hit:LA = 1:CheckLock 1:End Sub

Sub sw55_Hit:LB = 1:CheckLock 2:End Sub
Sub sw65_Hit:LC = 1:CheckLock 3:End Sub

Sub CheckLock(LNum)
    If LockOpen = 1 Then
        Select Case LNum
            Case 1:sw45.DestroyBall:sw45a.CreateBall:sw45a.Kick 180, 3:vpmTimer.PulseSw 45:LA = 0:PlaySoundAt "fx_balldrop", sw45a
            Case 2:sw55.DestroyBall:sw55a.CreateBall:sw55a.Kick 180, 3:vpmTimer.PulseSw 55:LB = 0:PlaySoundAt "fx_balldrop", sw55a
            Case 3:sw65.DestroyBall:sw65a.CreateBall:sw65a.Kick 180, 3:vpmTimer.PulseSw 65:LC = 0:PlaySoundAt "fx_balldrop", sw65a
        End Select
    Else
        Select Case LNum
            Case 1:Controller.Switch(45) = 1
            Case 2:Controller.Switch(55) = 1
            Case 3:Controller.Switch(65) = 1
        End Select
    End If
End Sub

Sub LockRelease(Enabled)
    If NOT Enabled Then
        PlaySoundAt "fx_Motor", Primitive68
        LockOpen = 1
        If LA = 1 Then
            PlaySoundAt "fx_balldrop", sw45a
            sw45.DestroyBall
            sw45a.CreateBall
            sw45a.Kick 180, 3
            Controller.Switch(45) = 0
            LA = 0
        End If
        If LB = 1 Then
            PlaySoundAt "fx_balldrop", sw55a
            sw55.DestroyBall
            sw55a.CreateBall
            sw55a.Kick 180, 3
            Controller.Switch(55) = 0
            LB = 0
        End If
        If LC = 1 Then
            PlaySoundAt "fx_balldrop", sw65a
            sw65.DestroyBall
            sw65a.CreateBall
            sw65a.Kick 125, 3
            Controller.Switch(65) = 0
            LC = 0
        End If
    Else
        LockOpen = 0
        PlaySoundAt "fx_Motor", Primitive68
    End If
End Sub

'*******************
' Flipper Subs Rev3
'*******************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("fx_flipperup", DOFFlippers), LeftFlipper
        LeftFlipper.RotateToEnd
        LeftFlipper1.RotateToEnd
        LeftFlipperOn = 1
    Else
        PlaySoundAt SoundFX("fx_flipperdown", DOFFlippers), LeftFlipper
        LeftFlipper.RotateToStart
        LeftFlipper1.RotateToStart
        LeftFlipperOn = 0
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("fx_flipperup", DOFFlippers), RightFlipper
        RightFlipper.RotateToEnd
        RightFlipperOn = 1
    Else
        PlaySoundAt SoundFX("fx_flipperdown", DOFFlippers), RightFlipper
        RightFlipper.RotateToStart
        RightFlipperOn = 0
    End If
End Sub

' flippers top animations

Sub LeftFlipper_Animate:leftflip.RotZ = LeftFlipper.CurrentAngle: End Sub
Sub LeftFlipper1_Animate:leftFlip1.RotZ = LeftFlipper1.CurrentAngle: End Sub
Sub RightFlipper_Animate: rightflip.RotZ = RightFlipper.CurrentAngle: End Sub

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
    UpdateLamps
End Sub

Sub UpdateLamps
    ' playfield lights
    Lamp 1, li1
    Lamp 2, li2
    Lamp 3, li3
    Lamp 4, li4
    Lamp 5, li5
    Flash 6, li6
    Flash 7, li7
    'Lamp 8, li8
    Lamp 9, li9
    Lamp 10, li10
    Lamp 11, li11
    Lamp 12, li12
    Lamp 13, li13
    Flash 14, li14
    Lamp 15, li15
    'Lamp 16, li16
    Lamp 17, li17
    Lamp 18, li18
    Lamp 19, li19
    Lamp 20, li20
    Lamp 21, li21
    Lamp 22, li22
    Lamp 23, li23
    'Lamp 24, li24
    Lamp 25, li25
    Lamp 26, li26
    Lamp 27, li27
    Lamp 28, li28
    Lamp 29, li29
    Lamp 30, li30
    'Lamp 31, li31
    'Lamp 32, li32
    Lamp 33, li33
    Lamp 34, li34
    Lamp 35, li35
    Lamp 36, li36
    Lamp 37, li37
    Lampm 38, li38a
    Lampm 38, li38b
    Lamp 38, li38
    'Lamp 39, li39
    'Lamp 40, li40
    Lamp 41, li41
    Lamp 42, li42
    Lamp 43, li43
    Lamp 44, li44
    Lamp 45, li45
    Lamp 46, li46
    Lamp 47, li47
    'Lamp 48, li48
    Lamp 49, li49
    Lamp 50, li50
    Lamp 51, li51
    Lamp 52, li52
    Lamp 53, li53
    'Lamp 54, li54
    'Lamp 55, li55
    'Lamp 56, li56
    Lamp 57, li57
    Lamp 58, li58
    Lamp 59, li59
    Lamp 60, li60
    Lamp 61, li61
    'Lamp 62, li62
    'Lamp 63, li63
    'Lamp 64, li64
    'flashers
    Lampm 73, li73a 'house of terror
    Lampm 73, li73b
    Lampm 73, li73c
    Lampm 73, li73d
    Lamp 73, li73
    Lamp 74, li74   'shooting Gallery
    Lampm 75, li75a ' Magnetic house
    Lamp 75, li75
    Lamp 76, li76   'Train
    Lampm 77, li77a 'bumper
    Lamp 77, li77
    Lampm 78, li78a 'bumper
    Lamp 78, li78
    Lampm 79, li79a 'bumper
    Lamp 79, li79
    Lamp 80, li80   'bumper
    Flash 81, li81  'flash looping
    'extra lamps
    ' Lamp 65  'motor enable
    Lampm 66, li66a 'pf gi 1
    Lampm 66, li66b
    Lampm 66, li66c
    Lampm 66, li66d
    Lampm 66, li66e
    Lampm 66, li66f
    Lampm 66, li66g
    Lampm 66, li66h
    Lampm 66, li66i
    Lampm 66, li66j
    Lampm 66, li66k
    Lampm 66, li66l
    Lampm 66, li66m
    Lampm 66, li66n
    Lampm 66, li66o
    Lamp 66, li66
    Lampm 67, li67a 'pf gi 2
    Lampm 67, li67b
    Lampm 67, li67c
    Lampm 67, li67d
    Lampm 67, li67e
    Lampm 67, li67g
    Lampm 67, li67h
    Lampm 67, li67i
    Lampm 67, li67j
    Lamp 67, li67k
    Lampm 68, li68a 'pf gi 3
    Lampm 68, li68b
    Lampm 68, li68c
    Lamp 68, li68
    Lampm 69, li69a 'pf gi 4
    Lampm 69, li69b
    Lampm 69, li69c
    Lampm 69, li69d
    Lampm 69, li69e
    Lampm 69, li69f
    Lampm 69, li69g
    Lampm 69, li69h
    Lampm 69, li69i
    Lampm 69, li69j
    Lampm 69, li69k
    Lamp 69, li69
    Lampm 70, li70a 'pf gi 5
    Lampm 70, li70b
    Lampm 70, li70c
    Lampm 70, li70d
    Lampm 70, li70e
    Lampm 70, li70f
    Lampm 70, li70g
    Lampm 70, li70j
    Lampm 70, li70i
    Lamp 70, li70
'Lamp 71, li71 'backglass gi 1
'Lamp 72, li72 'backglass gi 2
'Lamp 82, li82 'backglass flash 1
'Lamp 83, li83 'backglass flash 2
'Lamp 84, li84 'backglass flash 3
'Lamp 85, li85 'backglass flash 4
'Lamp 86, li86 'backglass flash 5
End Sub

' div lamp subs

' Normal Lamp & Flasher subs

Sub InitLamps()
    Dim x
    LampTimer.Interval = 10
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
Const maxvel = 42 'max ball velocity
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
                ballpitch = Pitch(BOT(b)) + 25000 'increase the pitch on a ramp
                ballvol = Vol(BOT(b)) * 2
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

'********************
' Wheel animation
'********************

Dim WheelPos
WheelPos = 0

Dim WheelBall(72)

For x = 0 to 71
    WheelBall(x) = 0
Next

Sub WheelEnable_Hit
    WheelTimer.Enabled = 1
End Sub

Sub aWheel_Hit(idx) 'a ball has hit one kicker, set up 1 in the array,
    WheelBall(idx) = 1
End Sub

Sub WheelTimer_Timer
    dim tmp

    ' disable all the kickers
    For each tmp in aWheel
        tmp.Enabled = 0
    Next

    ' turn the wheel 5 degrees
    WheelPos = (WheelPos + 5) Mod 360
    Wheel.Rotz = WheelPos

    ' if there is a ball in the last kicker then kick the ball to the lower floor
    If WheelBall(71) = 1 Then
        vpmTimer.PulseSw 82
        WheelBall(71) = 0
        aWheel(71).DestroyBall
        WheelExit.CreateBall
        WheelExit.Kick 170 + RND * 20, RND * 8
        ' add a little delay and check if there are more balls
        vpmtimer.addtimer 2000 + RND*3000, "CheckWheel '"
    End If

    ' move the balls to the next kickers
    For x = 70 to 0 step -1
        If WheelBall(x) = 1 Then
            WheelBall(x) = 0:aWheel(x).Destroyball
            WheelBall(x + 1) = 1:aWheel(x + 1).CreateBall
        End If
    Next

    ' enable the active kickers that corresponds to the holes
    tmp = Wheelpos \ 5
    aWheel(tmp).Enabled = 1
    tmp = (tmp + 24) Mod 72
    aWheel(tmp).Enabled = 1
    tmp = (tmp + 24) Mod 72
    aWheel(tmp).Enabled = 1
End Sub

' check if there are more balls in the wheel, if not stop the timer

Sub CheckWheel
    Dim tmp
    tmp = 0
    For x = 0 to 71
        tmp = tmp + WheelBall(x)
    Next
    If tmp = 0 Then
        WheelTimer.Enabled = 0
    End If
End Sub

'Spinball Jolly Park
'added by Inkochnito
Sub editDips
    Dim vpmDips:Set vpmDips = New cvpmDips
    With vpmDips
        .AddForm 700, 400, "Jolly Park - DIP switches"
        .AddFrame 0, 0, 190, "Replay threshold", &H000000C0, Array("300 million points", 0, "400 million points", &H00000080, "500 million points", &H00000040, "600 million points", &H000000C0)                                                                                                 'SL1-5&6(Dip 8,7)
        .AddFrame 0, 76, 190, "Handicap value", &H00000300, Array("700 million points", 0, "750 million points", &H00000100, "800 million points", &H00000200, "850 million points", &H00000300)                                                                                                  'SL2-3&4(dip 10,9)
        .AddFrame 0, 152, 190, "Credits per coin", &H0000000F, Array("1 coin - 1 credit && 1 coin - 3 credits", &H00000005, "1 coin - 1 credit && 1 coin - 5 credits", &H00000003, "1 coin - 1 credit && 1 coin - 6 credits", &H00000007, "1 coin - 2 credits && 1 coin - 5 credits", &H0000000F) 'SL1-4,3,2,1(Dip 1-4)
        .AddFrame 0, 228, 190, "Game Mode", &H00000010, Array("normal", 0, "exhibition", &H00000010)                                                                                                                                                                                              'SL1-8(dip 5)
        .AddChk 0, 280, 190, Array("Automatic points adjust enable", &H00080000)                                                                                                                                                                                                                  'SL3-1(dip 20)
        .AddFrame 205, 0, 190, "New ticket time interval", 49152, Array("disabled", 0, "short", 32768, "medium", &H00004000, "long", 49152)                                                                                                                                                       'SL2-5&6(dip 16,15)
        .AddFrame 205, 76, 190, "Repeat time for multiball", &H00C00000, Array("disabled", 0, "short", &H00800000, "medium", &H00400000, "long", &H00C00000)                                                                                                                                      'SL3-5&6(dip 24&23)
        .AddFrame 205, 152, 190, "Loopings needed for extra ball", &H00030000, Array("points (no extra ball)", 0, "1 looping", &H00010000, "2 loopings", &H00020000, "3 loopings", &H00030000)                                                                                                    'SL3-3&4(dip 18&17)
        .AddFrame 205, 228, 190, "Balls per game", &H0000800, Array("3 balls", 0, "5 balls", &H00000800)                                                                                                                                                                                          'SL2-1(dip 12)
        .AddFrame 205, 274, 190, "Special option", &H00200000, Array("easy", 0, "difficult", &H00200000)                                                                                                                                                                                          'SL3-7(dip 22)
        .AddLabel 50, 350, 300, 20, "After hitting OK, press F3 to reset game with new settings."
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

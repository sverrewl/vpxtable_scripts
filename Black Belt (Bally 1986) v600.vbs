'****************************************
'*                    *
'*        BLACK BELT        *
'*                    *
'*      BALLY MIDWAY (1986)     *
'*                    *
'****************************************

'Black Belt Bally Midway June 1986 / 4 Players
'Bally MPU A084-91786-AH06 (6803)
'vpinmame script based on destruk's script

Option Explicit
Randomize

Const BallSize = 50
Const BallMass = 1

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the Controller.vbs file in order to run this table (installed with the VPX package in the scripts folder)"
On Error Goto 0

LoadVPM "01120100", "6803.VBS", 3.1

Const UseSolenoids = 2
Const UseLamps = 0
Const UseSync = 0
Const HandleMech = 0
Const SSolenoidOn = "fx_solenoidon"
Const SSolenoidOff = "fx_solenoidoff"
Const SCoin = "fx_Coin"

Dim DesktopMode:DesktopMode = Table1.ShowDT
Dim x

' Remove desktop items in FS mode
If Table1.ShowDT then
    For each x in aReels
        x.Visible = 1
    Next
Else
    For each x in aReels
        x.Visible = 0
    Next
End If

'************
' Table init.
'************

Dim bsTrough, bsLeft, bsMiddle, bsRight, dtL, dtR
Const cGameName = "blackblt"

Sub Table1_Init
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Black Belt - Bally Midway 1986" & vbnewline & "VPX by fredobiwan & jpsalas v6.0.0"
        .Games(cGameName).Settings.Value("rol") = 0   '1= rotated display, 0= normal
        .Games(cGameName).Settings.Value("sound") = 1 '1 enabled rom sound
        .HandleMechanics = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .ShowTitle = 0
        .Hidden = 1 'DesktopMode
        On Error Resume Next
        .SolMask(0) = 0
        vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the Timer to renable all the solenoids after 2 seconds
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
        On Error Goto 0
    End With

    vpmNudge.TiltSwitch = 14
    vpmNudge.Sensitivity = 5
    vpmNudge.TiltObj = Array(LeftSlingshot, RightSlingshot, sw25, sw26)

    Set dtL = New cvpmDropTarget
    dtL.InitDrop Array(sw42, sw44, sw46), Array(42, 44, 46)
    dtL.InitSnd SoundFX("", DOFContactors), SoundFX("fx_resetdrop", DOFContactors)
    dtL.CreateEvents "dtL"

    Set dtR = New cvpmDropTarget
    dtR.InitDrop Array(sw41, sw43, sw45), Array(41, 43, 45)
    dtR.InitSnd SoundFX("", DOFContactors), SoundFX("fx_resetdrop", DOFContactors)
    dtR.CreateEvents "dtR"

    ' Thalamus - more randomness pls - tempted to add to bsMiddle too - might be a reason why it is commented out.
    Set bsLeft = New cvpmBallStack
    bsLeft.InitSaucer Kicker1, 4, 140, 16
    bsLeft.InitExitSnd SoundFX("fx_Solenoidon", DOFContactors), SoundFX("fx_Solenoidon", DOFContactors)
    bsLeft.KickForceVar = 3
    bsLeft.KickAngleVar = 3

    Set bsMiddle = New cvpmBallStack
    bsMiddle.InitSaucer Kicker2, 2, 185, 17
    bsMiddle.InitExitSnd SoundFX("fx_Solenoidon", DOFContactors), SoundFX("fx_Solenoidon", DOFContactors)
    ' bsMiddle.KickAngleVar=2

    Set bsRight = New cvpmBallStack
    bsRight.InitSaucer Kicker3, 3, 5, 37
    bsRight.InitExitSnd SoundFX("fx_Solenoidon", DOFContactors), SoundFX("fx_Solenoidon", DOFContactors)
    bsRight.KickForceVar = 3
    bsRight.KickAngleVar = 3

    Set bsTrough = New cvpmBallStack
    bsTrough.InitSw 0, 8, 0, 0, 0, 0, 0, 0
    bsTrough.InitKick BallRelease, 90, 8
    bsTrough.InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    bsTrough.Balls = 1

    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

    ' turn on Gi
    vpmtimer.addtimer 1000, "GiOn '"

  ' turn the RealTimeUpdates timer
  RealTimeUpdates.Enabled = 1
End Sub

'******************
' RealTime Updates
'******************

' Solenoid check, only for this table
Dim Phase
Phase = 0

'Set MotorCallback = GetRef("RealTimeUpdates")

Sub RealTimeUpdates_Timer
    RollingUpdate
    If Controller.Lamp(6)And Controller.Lamp(54)Then
        Phase = 1
    Else
        Phase = 0
    End If
End Sub

Sub LeftGate_Animate:LeftGateP.RotZ = LeftGate.CurrentAngle:End Sub
Sub RightGate_Animate: RightGateP.RotZ = RightGate.CurrentAngle:End Sub

'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If KeyDownHandler(KeyCode)Then Exit Sub
    If keycode = PlungerKey Then PlaySoundAt "fx_PlungerPull", Plunger:Plunger.Pullback
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If KeyUpHandler(KeyCode)Then Exit Sub
    If keycode = PlungerKey Then PlaySoundAt "fx_plunger", plunger:Plunger.Fire
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
        LeftFlipper2.RotateToEnd
        LeftFlipperOn = 1
    Else
        PlaySoundAt SoundFX("fx_flipperdown", DOFFlippers), LeftFlipper
        LeftFlipper.RotateToStart
        LeftFlipper2.RotateToStart
        LeftFlipperOn = 0
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("fx_flipperup", DOFFlippers), RightFlipper
        RightFlipper.RotateToEnd
        RightFlipper2.RotateToEnd
        RightFlipperOn = 1
    Else
        PlaySoundAt SoundFX("fx_flipperdown", DOFFlippers), RightFlipper
        RightFlipper.RotateToStart
        RightFlipper2.RotateToStart
        RightFlipperOn = 0
    End If
End Sub

' flippers hit Sound

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub LeftFlipper2_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper2_Collide(parm)
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

'*********
'Solenoids
'*********

'1 bumper left
'2 bumper right
'3 left slingshot
'4 right slingshot
SolCallback(5) = "Sol5" 'flash man
SolCallback(6) = "Sol6" 'flash sun
SolCallback(7) = "SolKickerLeft"
SolCallback(8) = "SolKickerMiddle"
SolCallback(9) = "SolKickerRight"
SolCallback(10) = "SetLamp 110,"
SolCallback(11) = "SetLamp 111,"
SolCallback(12) = "SetLamp 112,"
'13 ??
SolCallback(14) = "bsTrough.SolOut Not"
SolCallback(15) = "vpmSolSound SoundFX(""fx_Knocker"",DOFKnocker),"
SolCallback(16) = "Sol17" '? I use it for a gi effect
SolCallback(17) = "Sol17" '? I use it for a gi effect
SolCallback(18) = "SolLeftGate"
SolCallback(19) = "vpmNudge.SolGameOn"
SolCallback(20) = "SolRightGate"

Sub Sol5(Enabled) 'BLR
    SetLamp 105, Enabled
    If Enabled And Phase = 1 Then
        dtL.DropSol_On
    End If
End Sub

Sub Sol6(Enabled) 'BLR
    SetLamp 106, Enabled
    If Enabled And Phase = 1 Then
        dtR.DropSol_On
    End If
End Sub

Sub SolKickerLeft(Enabled)
    If Enabled And bsLeft.Balls Then
        PlaySoundAt SoundFX("fx_kicker", DOFContactors), kicker1
        bsLeft.ExitSol_On
    End If
End Sub

Sub SolKickerMiddle(Enabled) 'BLR
    SetLamp 108, Enabled
    If Enabled And bsMiddle.Balls Then
        PlaySoundAt SoundFX("fx_kicker", DOFContactors), kicker2
        bsMiddle.ExitSol_On
    End If
End Sub

Dim R2Step

Sub SolKickerRight(Enabled)
    SetLamp 109, Enabled
    If Enabled And bsRight.Balls Then
        PlaySoundAt SoundFX("fx_kicker", DOFContactors), kicker3
        bsRight.ExitSol_On
    R2Step = 0
    Remk2.Rotx = 24
    kicker3.TimerEnabled = 1
    End If
End Sub

Sub kicker3_Timer
    Select Case R2Step
        Case 1:Remk2.RotX = 16
        Case 2:Remk2.RotX = 8
        Case 3:Remk2.RotX = 0:kicker3.TimerEnabled = 0
    End Select
    R2Step = R2Step + 1
End Sub

Sub SolLeftGate(Enabled)
    If Enabled Then
        LeftGate.RotateToEnd
        PlaySoundAt SoundFX("fx_SolenoidOn", DOFContactors), LeftGate
        vpmtimer.addtimer 3000, "LeftGate.RotateToStart:PlaySoundAt SoundFX(""fx_SolenoidOff"", DOFContactors), LeftGate '"
    End If
End Sub

Sub SolRightGate(Enabled)
    If Enabled Then
        RightGate.RotateToEnd
        PlaySoundAt SoundFX("fx_SolenoidOn", DOFContactors), RightGate
        vpmtimer.addtimer 3000, "RightGate.RotateToStart:PlaySoundAt SoundFX(""fx_SolenoidOff"", DOFContactors), RightGate '"
    End If
End Sub

Sub LeftGate_Collide(parm)
    PlaySound "fx_metalhit", 0, parm / 60, pan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightGate_Collide(parm)
    PlaySound "fx_metalhit", 0, parm / 60, pan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub Sol17(Enabled) 'Gi effect
    If Enabled Then
        GiOn
  Else
    GiOff
    End If
End Sub

'Drain & Kickers
Sub Drain_Hit:PlaySoundAt "fx_drain", Drain:bsTrough.AddBall Me:End Sub
Sub Kicker1_Hit:PlaySoundAt "fx_kicker_enter", kicker1:bsLeft.AddBall 0:End Sub
Sub Kicker2_Hit:PlaySoundAt "fx_kicker_enter", kicker2:bsMiddle.AddBall 0:End Sub
Sub Kicker3_Hit:PlaySoundAt "fx_kicker_enter", kicker3:bsRight.AddBall 0:End Sub

'Open Gate Right
Sub sw12_Hit:Controller.Switch(12) = 1:PlaySoundAt "fx_sensor", sw12:End Sub
Sub sw12_UnHit:Controller.Switch(12) = 0:End Sub

'Open Gate Left
Sub sw13_Hit:Controller.Switch(13) = 1:PlaySoundAt "fx_sensor", sw13:End Sub
Sub sw13_UnHit:Controller.Switch(13) = 0:End Sub

'Spinner
Sub sw16_Spin:vpmTimer.PulseSw(16):PlaysoundAt "fx_spinner", sw16:End Sub

'Out Lane Left/Right
Sub sw17_Hit:Controller.Switch(17) = 1:PlaySoundAt "fx_sensor", sw17:End Sub
Sub sw17_UnHit:Controller.Switch(17) = 0:End Sub

Sub sw18_Hit:Controller.Switch(18) = 1:PlaySoundAt "fx_sensor", sw18:End Sub
Sub sw18_UnHit:Controller.Switch(18) = 0:End Sub

'Targets
Sub sw19_Hit:vpmTimer.PulseSw(19):PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw20_Hit:vpmTimer.PulseSw(20):PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw21_Hit:vpmTimer.PulseSw(21):PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw22_Hit:vpmTimer.PulseSw(22):PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw23_Hit:vpmTimer.PulseSw(23):PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw24_Hit:vpmTimer.PulseSw(24):PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

'Bumpers

Sub sw25_Hit:vpmTimer.PulseSw(25):PlaySoundAt SoundFX("fx_bumper", DOFContactors), sw25:End Sub
Sub sw26_Hit:vpmTimer.PulseSw(26):PlaySoundAt SoundFX("fx_bumper", DOFContactors), sw26:End Sub

'Targets
Sub sw33_Hit:vpmTimer.PulseSw(33):PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw34_Hit:vpmTimer.PulseSw(34):PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw35_Hit:vpmTimer.PulseSw(35):PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw36_Hit:vpmTimer.PulseSw(36):PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw37_Hit:vpmTimer.PulseSw(37):PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw38_Hit:vpmTimer.PulseSw(38):PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

'Shop Side
Sub sw29_Hit:Controller.Switch(29) = 1:PlaySoundAt "fx_sensor", sw29:End Sub
Sub sw29_UnHit:Controller.Switch(29) = 0:End Sub

'Right Ramp
Sub sw30_Hit:Controller.Switch(30) = 1:PlaySoundAt "fx_sensor", sw30:End Sub
Sub sw30_UnHit:Controller.Switch(30) = 0:End Sub

'Left Ramp
Sub sw31_Hit:Controller.Switch(31) = 1:PlaySoundAt "fx_sensor", sw31:End Sub
Sub sw31_UnHit:Controller.Switch(31) = 0:End Sub

'Shooter Lane
Sub sw32_Hit:Controller.Switch(32) = 1:PlaySoundAt "fx_sensor", sw32:End Sub
Sub sw32_UnHit:Controller.Switch(32) = 0:End Sub

'Top Arc
Sub sw39_Hit:Controller.Switch(39) = 1:PlaySoundAt "fx_sensor", sw39:End Sub
Sub sw39_UnHit:Controller.Switch(39) = 0:End Sub

'Mushroom

Dim MushroomStep

Sub sw40_hit
    PlaySoundAtBall "fx_passive_bumper"
    vpmTimer.PulseSw(40)
    MushroomStep = 0
    sw40.TimerEnabled = 1
End Sub

Sub sw40_Timer
    Select Case MushroomStep
        Case 1:MushroomPrim.ObjRotX = 4
        Case 2:MushroomPrim.ObjRotX = 8
        Case 3:MushroomPrim.ObjRotX = 4
        Case 4:MushroomPrim.ObjRotX = 0
        Case 5:MushroomPrim.ObjRotX = -4
        Case 6:MushroomPrim.ObjRotX = -8
        Case 7:MushroomPrim.ObjRotX = -4
        Case 8:MushroomPrim.ObjRotX = 0
        Case 9:MushroomPrim.ObjRotX = 4
        Case 10:MushroomPrim.ObjRotX = 0:MushroomStep = 0:sw40.TimerEnabled = 0
    End Select
    MushroomStep = MushroomStep + 1
End Sub

'Droptargets only hit sound
Sub sw41_Hit:PlaySoundAt SoundFX("fx_droptarget", DOFDropTargets), sw41:End Sub
Sub sw42_Hit:PlaySoundAt SoundFX("fx_droptarget", DOFDropTargets), sw42:End Sub
Sub sw43_Hit:PlaySoundAt SoundFX("fx_droptarget", DOFDropTargets), sw43:End Sub
Sub sw44_Hit:PlaySoundAt SoundFX("fx_droptarget", DOFDropTargets), sw44:End Sub
Sub sw45_Hit:PlaySoundAt SoundFX("fx_droptarget", DOFDropTargets), sw45:End Sub
Sub sw46_Hit:PlaySoundAt SoundFX("fx_droptarget", DOFDropTargets), sw46:End Sub

'In Lane Left/Right
Sub sw47_Hit:Controller.Switch(47) = 1:PlaySoundAt "fx_sensor", sw47:End Sub
Sub sw47_UnHit:Controller.Switch(47) = 0:End Sub
Sub sw48_Hit:Controller.Switch(48) = 1:PlaySoundAt "fx_sensor", sw48:End Sub
Sub sw48_UnHit:Controller.Switch(48) = 0:End Sub

' Slings
Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    PlaySoundAt SoundFX("fx_slingshot", DOFContactors), Lemk
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 28
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 1:LeftSling4.Visible = 0:LeftSLing3.Visible = 1:Lemk.RotX = 14
        Case 2:LeftSLing3.Visible = 0:LeftSLing2.Visible = 1:Lemk.RotX = 2
        Case 3:LeftSLing2.Visible = 0:Lemk.RotX = -20:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
    PlaySoundAt SoundFX("fx_slingshot", DOFContactors), Remk
    RightSling4.Visible = 1
    Remk.RotX = 26
    RStep = 0
    vpmTimer.PulseSw 27
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

Sub UpdateLamps()
    ' playfield lights
    Lamp 1, L1
    Lamp 2, L2
    Lamp 3, L3
    Flash 4, L4 'backboard
    Flash 5, L5 'backboard
    'Lamp 6, L6
    'Lamp 7, L7
    'Lamp 8, L8
    Lamp 9, L9
    Lamp 10, L10
    Lamp 11, L11
    Lamp 12, L12
    Lamp 13, L13
    Lamp 14, L14
    Lamp 15, L15
    'Lamp 16, L16
    Lamp 17, L17
    Lamp 18, L18
    Lamp 19, L19
    Flash 20, L20 'backboard
    Flash 21, L21 'backboard
    Lamp 22, L22
    Lamp 23, L23
    'Lamp 24, L24
    Lamp 25, L25
    Lamp 26, L26
    Lamp 27, L27
    Lamp 28, L28
    Lamp 29, L29
    Lamp 30, L30
    Lamp 31, L31
    'Lamp 32, L32
    Lamp 33, L33
    Lamp 34, L34
    Flash 35, L35 'backboard
    Flash 36, L36 'backboard
    Flash 37, L37 'backboard
    'Lamp 38, L38
    'Lamp 39, L39
    Lamp 40, L40
    Lamp 41, L41
    Lamp 42, L42
    Lamp 43, L43
    Lamp 44, L44
    Lamp 45, L45
    Lamp 46, L46
    Lamp 47, L47
    'Lamp 48, L48
    Lamp 49, L49
    Lamp 50, L50
    Lamp 51, L51
    Lamp 52, L52
    Lamp 53, L53
    'Lamp 54, L54
    Lamp 55, L55
    Lamp 56, L56
    Lamp 57, L57
    Lamp 58, L58
    Lamp 59, L59
    Lamp 60, L60
    Lamp 61, L61
    Lamp 62, L62
    Lamp 63, L63
    'Lamp 64, L64
    Lamp 65, L65
    Lamp 66, L66
    Lamp 67, L67
    Lamp 68, L68
    Lamp 69, L69
    Lamp 70, L70
    Lamp 71, L71
    Lamp 72, L72
    Lamp 73, L73
    Lamp 74, L74
    Lamp 75, L75
    Lamp 76, L76
    Lamp 77, L77
    Lamp 78, L78
    Lamp 79, L79
    'Lamp 80, L80
    Lamp 81, L81
    Lamp 82, L82
    Lamp 83, L83
    Lamp 84, L84
    Lamp 85, L85
    Lamp 86, L86
    Lamp 87, L87
    Lamp 88, L88
    'Lamp 89, L89
    Lamp 90, L90
    Lamp 91, L91
    Lamp 92, L92
    Lamp 93, L93
    Lamp 94, L94
    Lamp 95, L95

    'Flashers
    Lampm 105, Fman1
    Lampm 105, Fman2
    Lampm 105, Fman3
    Lampm 105, Fman4
    Lampm 105, Fman5
    Lampm 105, Fman6
    Lamp 105, Fman7
    Lamp 106, BLightSun
    Lampm 108, BLR001
    Lampm 108, BLR002
    Lampm 108, BLR003
    Lampm 108, BLR1
    Lampm 108, BLR2
    Lamp 108, BLR3
    Lampm 109, BLTL001
    Lampm 109, BLTL002
    Lampm 109, BLTL1
    Lampm 109, BLTL2
    Lampm 109, BLTL3
    Lamp 109, BLTL4
    Lampm 110, BLTR001
    Lampm 110, BLTR002
    Lampm 110, BLTR1
    Lampm 110, BLTR2
    Lampm 110, BLTR3
    Lamp 110, BLTR4
    Flash 111, BLBL
    Flash 112, BLBR
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

'******************
' Digital Display
'******************

Dim Digits(28)
Digits(0) = Array(a00, a01, a02, a03, a04, a05, a06, a07, a08)
Digits(1) = Array(a10, a11, a12, a13, a14, a15, a16, a17, a18)
Digits(2) = Array(a20, a21, a22, a23, a24, a25, a26, a27, a28)
Digits(3) = Array(a30, a31, a32, a33, a34, a35, a36, a37, a38)
Digits(4) = Array(a40, a41, a42, a43, a44, a45, a46, a47, a48)
Digits(5) = Array(a50, a51, a52, a53, a54, a55, a56, a57, a58)
Digits(6) = Array(a60, a61, a62, a63, a64, a65, a66, a67, a68)

Digits(7) = Array(b00, b01, b02, b03, b04, b05, b06, b07, b08)
Digits(8) = Array(b10, b11, b12, b13, b14, b15, b16, b17, b18)
Digits(9) = Array(b20, b21, b22, b23, b24, b25, b26, b27, b28)
Digits(10) = Array(b30, b31, b32, b33, b34, b35, b36, b37, b38)
Digits(11) = Array(b40, b41, b42, b43, b44, b45, b46, b47, b48)
Digits(12) = Array(b50, b51, b52, b53, b54, b55, b56, b57, b58)
Digits(13) = Array(b60, b61, b62, b63, b64, b65, b66, b67, b68)

Digits(14) = Array(c00, c01, c02, c03, c04, c05, c06, c07, c08)
Digits(15) = Array(c10, c11, c12, c13, c14, c15, c16, c17, c18)
Digits(16) = Array(c20, c21, c22, c23, c24, c25, c26, c27, c28)
Digits(17) = Array(c30, c31, c32, c33, c34, c35, c36, c37, c38)
Digits(18) = Array(c40, c41, c42, c43, c44, c45, c46, c47, c48)
Digits(19) = Array(c50, c51, c52, c53, c54, c55, c56, c57, c58)
Digits(20) = Array(c60, c61, c62, c63, c64, c65, c66, c67, c68)

Digits(21) = Array(d00, d01, d02, d03, d04, d05, d06, d07, d08)
Digits(22) = Array(d10, d11, d12, d13, d14, d15, d16, d17, d18)
Digits(23) = Array(d20, d21, d22, d23, d24, d25, d26, d27, d28)
Digits(24) = Array(d30, d31, d32, d33, d34, d35, d36, d37, d38)
Digits(25) = Array(d40, d41, d42, d43, d44, d45, d46, d47, d48)
Digits(26) = Array(d50, d51, d52, d53, d54, d55, d56, d57, d58)
Digits(27) = Array(d60, d61, d62, d63, d64, d65, d66, d67, d68)

Sub UpdateLeds()
    Dim chgLED, ii, num, chg, stat, obj
    chgLED = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(chgLED)Then
        If DesktopMode Then
            For ii = 0 To UBound(chgLED)
                num = chgLED(ii, 0):chg = chgLED(ii, 1):stat = chgLED(ii, 2)
                if(num <32)then
                    For Each obj In Digits(num)
                        If chg And 1 Then obj.State = stat And 1
                        chg = chg \ 2:stat = stat \ 2
                    Next
                End If
            Next
        End If
    End If
End Sub

'*******************
' GI Lights
'*******************

Sub GiOn
    Dim bulb
    For each bulb in aGiLights
        bulb.State = 1
    Next
End Sub

Sub GiOff
    Dim bulb
    For each bulb in aGiLights
        bulb.State = 0
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


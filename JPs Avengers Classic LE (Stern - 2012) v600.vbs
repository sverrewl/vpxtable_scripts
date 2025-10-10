' JP's Avengers LE
' Based on Stern's The Avengers (Limited Edition) / IPD No. 5940 / 2012 / 4 Players
' Uses the Avengers LE ROM : avs_170h
' Playfield & plastics redrawn by me & Siggi.
' VPX8 - version by JPSalas 2025, version 6.0.0

Option Explicit
Randomize

Const cSingleRFlip = 0
Const cSingleLFlip = 0
Const BallSize = 50
Const BallMass = 1

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Dim bsTrough, bsSaucer, dtBankThor, dtBankHulk, mHulkMag, plungerIM, vlLock, x

Const cGameName = "avs_170h" 'latest avengers le rom get it from Stern's web site

Const UseSolenoids = 1
Const UseLamps = 1
Const UseGI = 0
Const UseSync = 0 'set it to 1 if the table runs too fast
Const HandleMech = 0

' Use Modulated Flashers?

Const UseVPMModSol = True

Dim VarHidden, UseVPMDMD
If Table1.ShowDT = true then
    UseVPMDMD = True
    VarHidden = 1
else
    UseVPMDMD = False
    VarHidden = 0
end if

LoadVPM "01550000", "SAM.VBS", 3.26

' Standard Sounds
Const SSolenoidOn = "fx_solenoidon"
Const SSolenoidOff = "fx_solenoidoff"
Const SCoin = "fx_Coin"

'************
' Table init.
'************

Sub table1_Init
    vpmInit me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "JP's Avengers Classic" & vbNewLine & "VPX table by JPSalas & Siggi v.6.0.0"
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = VarHidden
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
    vpmNudge.TiltSwitch = -7
    vpmNudge.Sensitivity = 3
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingshot)

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0, 22, 21, 20, 19, 18, 17, 0
        .InitKick BallRelease, 90, 10
        .InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .Balls = 6
    End With

    ' Thalamus, more randomness pls
    'Kickers
    Set bsSaucer = New cvpmBallStack
    bsSaucer.InitSaucer sw62, 62, 25, 16
    bsSaucer.InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    bsSaucer.KickAngleVar = 3
    bsSaucer.KickForceVar = 3

    'droptargets
    Set dtBankHulk = New cvpmDropTarget
    dtBankHulk.InitDrop Array(sw52, sw53, sw54, sw55), Array(52, 53, 54, 55)
    dtBankHulk.InitSnd SoundFX("fx_droptarget", DOFContactors), SoundFX("fx_resetdrop", DOFContactors)
    dtBankHulk.CreateEvents "dtBankHulk"

    Set dtBankThor = New cvpmDropTarget
    dtBankThor.InitDrop Array(sw1, sw2, sw3, sw4), Array(1, 2, 3, 4)
    dtBankThor.InitSnd SoundFX("fx_droptarget", DOFContactors), SoundFX("fx_resetdrop", DOFContactors)
    dtBankThor.CreateEvents "dtBankThor"

    ' Hulk magnet
    Set mHulkMag = New cvpmMagnet
    mHulkMag.InitMagnet HulkMag, 16
    mHulkMag.GrabCenter = False
    mHulkMag.solenoid = 54
    mHulkMag.CreateEvents "mHulkMag"

    ' Visible Lock - implements post ball lock
    Set vlLock = New cvpmVLock2
    With vlLock
        .InitVLock Array(sw49, sw50, sw51), Array(sw49k, sw50k, sw51k), Array(49, 50, 51)
        .InitSnd "fx_sensor", "fx_sensor"
        .CreateEvents "vlLock"
    End With

    ' Impulse Plunger - used as the autoplunger
    Const IMPowerSetting = 62 'Plunger Power
    Const IMTime = 0.6        ' Time in seconds for Full Plunge
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swd15, IMPowerSetting, IMTime
        .Random 0.3
        .switch 86
        .InitExitSnd SoundFX("fx_plunger", DOFContactors), SoundFX("fx_plunger", DOFContactors)
        .CreateEvents "plungerIM"
    End With

    vpmMapLights aLights

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

    ' Loki Lock Init - Optos used so switch is opposite (1=no ball, 0=ball).  Default to 1
    Controller.Switch(51) = 1
    Controller.Switch(50) = 1
    Controller.Switch(49) = 1
    Controller.Switch(8) = 1

    'Fast Flips
    On Error Resume Next
    InitVpmFFlipsSAM
    If Err Then MsgBox "You need the latest sam.vbs in order to run this table, available with vp10.5"
    On Error Goto 0
End Sub

Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub
Sub table1_exit:Controller.stop:End Sub

'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
    If keycode = LeftTiltKey Then Nudge 90, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 7:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If Keycode = StartGameKey Then Controller.Switch(16) = 1
    If vpmKeyDown(keycode)Then Exit Sub
    If keycode = PlungerKey Then PlaySoundAt "fx_PlungerPull", Plunger:Plunger.Pullback
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If Keycode = StartGameKey Then Controller.Switch(16) = 0
    If vpmKeyUp(keycode)Then Exit Sub
    If keycode = PlungerKey Then PlaySoundAt "fx_plunger", Plunger:Plunger.Fire
End Sub

'*********
' Switches
'*********

' Slings
Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    PlaySoundAt SoundFX("fx_slingshot", DOFContactors), Lemk
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 26
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

' Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 30:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper1:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 31:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper2:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 32:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper3:End Sub

' Drain & Saucers
Sub Drain_Hit:PlaySoundAt "fx_drain", Drain:bsTrough.AddBall Me:End Sub
Sub sw62_Hit:PlaySoundAt "fx_kicker_enter", sw62:bsSaucer.AddBall 0:End Sub

' Rollovers
Sub sw10_Hit:Controller.Switch(10) = 1:PlaySoundAt "fx_sensor", sw10:End Sub
Sub sw10_UnHit:Controller.Switch(10) = 0:End Sub

Sub sw11_Hit:Controller.Switch(11) = 1:PlaySoundAt "fx_sensor", sw11:End Sub
Sub sw11_UnHit:Controller.Switch(11) = 0:End Sub

Sub sw12_Hit:Controller.Switch(12) = 1:PlaySoundAt "fx_sensor", sw12:End Sub
Sub sw12_UnHit:Controller.Switch(12) = 0:End Sub

Sub sw23_Hit:Controller.Switch(23) = 1:PlaySoundAt "fx_sensor", sw23:End Sub
Sub sw23_UnHit:Controller.Switch(23) = 0:End Sub

Sub sw24_Hit:Controller.Switch(24) = 1:PlaySoundAt "fx_sensor", sw24:End Sub
Sub sw24_UnHit:Controller.Switch(24) = 0:End Sub

Sub sw25_Hit:Controller.Switch(25) = 1:PlaySoundAt "fx_sensor", sw25:End Sub
Sub sw25_UnHit:Controller.Switch(25) = 0:End Sub

Sub sw29_Hit:Controller.Switch(29) = 1:PlaySoundAt "fx_sensor", sw29:End Sub
Sub sw29_UnHit:Controller.Switch(29) = 0:End Sub

Sub sw28_Hit:Controller.Switch(28) = 1:PlaySoundAt "fx_sensor", sw28:End Sub
Sub sw28_UnHit:Controller.Switch(28) = 0:End Sub

Sub sw33_Hit:Controller.Switch(33) = 1:PlaySoundAt "fx_sensor", sw33:End Sub
Sub sw33_UnHit:Controller.Switch(33) = 0:End Sub

Sub sw43_Hit:Controller.Switch(43) = 1:PlaySoundAt "fx_sensor", sw43:End Sub
Sub sw43_UnHit:Controller.Switch(43) = 0:End Sub

Sub sw47_Hit:Controller.Switch(47) = 1:PlaySoundAt "fx_sensor", sw47:End Sub
Sub sw47_UnHit:Controller.Switch(47) = 0:End Sub

Sub sw48_Hit:Controller.Switch(48) = 1:PlaySoundAt "fx_sensor", sw48:End Sub
Sub sw48_UnHit:Controller.Switch(48) = 0:End Sub

Sub sw58_Hit:Controller.Switch(58) = 1:PlaySoundAt "fx_sensor", sw58:End Sub
Sub sw58_UnHit:Controller.Switch(58) = 0:End Sub

Sub sw57_Hit:Controller.Switch(57) = 1:PlaySoundAt "fx_sensor", sw57:End Sub
Sub sw57_UnHit:Controller.Switch(57) = 0:End Sub

' Spinners
Sub sw44_Spin:vpmTimer.PulseSw 44:PlaySoundAt "fx_spinner", sw44:End Sub

'Targets
Sub sw7_Hit:vpmTimer.Pulsesw 7:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub sw13_Hit:vpmTimer.Pulsesw 13:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub sw14_Hit:vpmTimer.Pulsesw 14:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub sw34_Hit:vpmTimer.Pulsesw 34:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub sw35_Hit:vpmTimer.Pulsesw 35:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub sw36_Hit:vpmTimer.Pulsesw 36:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub sw63_Hit:vpmTimer.Pulsesw 63:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub

'*********
'Solenoids
'*********
SolCallback(1) = "SolTrough"            'trough
SolCallback(2) = "SolAutofire"          'auto plunger
SolCallback(3) = "SolHulkCCW"           'hulk turn ccw
SolCallback(4) = "SolHulkCW"            'hulk turn cw
SolCallback(5) = "bsSaucer.SolOut"      'hulk eject
SolCallback(6) = "dtBankThor.SolDropUp" 'left 4 bank drop reset
SolCallback(7) = "SolLeftOrbit"         'orbit control gate left
SolCallback(8) = "SolShakerMotor"       'shaker motor
'9 bumper left
'10 bumper right
'11 bumper bottom
SolCallback(12) = "SolLockRelease" 'loki lockup
'13 Left Slingshot
'14 RightSLingshot
SolCallback(15) = "SolLFlipper"
SolCallback(16) = "SolRFlipper"
'17 relay (blue)
'22 bridge motor
SolCallback(23) = "SolBridge" 'bridge motor relay
'24 optional coil
SolCallBack(51) = "dtBankHulk.SolDropUp" 'center 4-bank drop reset
SolCallBack(52) = "solRightRampGate"     'ramp control gate (right)
SolCallBack(53) = "solLeftRampGate"      'ramp control gate (left)
'54 hulk magnet
'55 relay (green)
SolCallBack(56) = "SolHulkArms"   'hulk arms
SolCallBack(57) = "SolRightOrbit" 'orbit control gate (right)
'58 relay (red)

SolModCallback(18) = "Flasher18"
SolModCallback(19) = "Flasher19"
SolModCallback(20) = "Flasher20"
SolModCallback(21) = "Flasher21"
SolModCallback(25) = "Flasher25"
SolModCallback(26) = "Flasher26"
SolModCallback(27) = "Flasher27"
SolModCallback(28) = "Flasher28"
SolModCallback(29) = "Flasher29"
SolModCallback(30) = "Flasher30"
SolModCallback(31) = "Flasher31"
SolModCallback(32) = "Flasher32"

Sub Flasher18(m): m = m /255: f18l.State = m: End Sub
Sub Flasher19(m): m = m /255: f19l.State = m: End Sub
Sub Flasher20(m): m = m /255: f20l.State = m: End Sub
Sub Flasher21(m): m = m /255: f21l.State = m: End Sub
Sub Flasher25(m): m = m /255: f25l.State = m: End Sub
Sub Flasher26(m): m = m /255: f26l.State = m: End Sub
Sub Flasher27(m): m = m /255: f27l.State = m: End Sub
Sub Flasher28(m): m = m /255: f28l.State = m: End Sub
Sub Flasher29(m): m = m /255: f29l.State = m: End Sub
Sub Flasher30(m): m = m /255: f30l.State = m: End Sub
Sub Flasher31(m): m = m /255: f31l.State = m: End Sub
Sub Flasher32(m): m = m /255: f32l.State = m: End Sub

Sub SolTrough(Enabled)
    If Enabled Then
        bsTrough.ExitSol_On
        vpmTimer.PulseSw 23
    End If
End Sub

Sub SolAutofire(Enabled)
    If Enabled Then
        PlungerIM.AutoFire
    End If
End Sub

Sub SolHulkCCW(Enabled)
    If Enabled Then
        body.RotateToStart
        arm1.RotateToStart
        arm2.RotateToStart
        PlaySoundAt "fx_solenoidOn", body
    End If
End Sub

Sub SolHulkCW(Enabled)
    If Enabled Then
        body.RotateToEnd
        arm1.RotateToEnd
        arm2.RotateToEnd
        PlaySoundAt "fx_solenoidOff", body
    End If
End Sub

Sub SolLeftOrbit(Enabled)
    If Enabled Then
        OrbitLeft.open = True
        PlaySoundAt "fx_solenoidOn", OrbitLeft
    Else
        OrbitLeft.open = False
        PlaySoundAt "fx_solenoidOff", OrbitLeft
    End If
End Sub

Sub SolShakerMotor(Enabled)
    If enabled Then
        ShakeTimer.Enabled = 1
        Playsound SoundFX("fx_shaker", DOFContactors)
    Else
        ShakeTimer.Enabled = 0
    End If
End Sub

Sub ShakeTimer_Timer()
    Nudge 0, 0.5
    Nudge 90, 0.5
    Nudge 180, 0.5
    Nudge 270, 0.5
End Sub

Sub SolLockRelease(enabled)
    vlLock.SolExit enabled
    lokipost.IsDropped = Enabled
End Sub

Sub SolHulkArms(Enabled)
    If Enabled Then
        arms.RotateToEnd
        arm1.Enabled = 0
        arm2.Enabled = 0
        PlaySoundAt "fx_solenoidOn", arms
    Else
        arms.RotateToStart
        arm1.Enabled = 1
        arm2.Enabled = 1
        PlaySoundAt "fx_solenoidOff", arms
    End If
End Sub

Sub SolBridge(Enabled)
    If Enabled Then
        gateleft2.RotateToEnd
        PlaySoundAt "fx_solenoidOn", gateleft2
        Controller.Switch(8) = 1
        Controller.Switch(9) = 0
    Else
        gateleft2.RotateToStart
        PlaySoundAt "fx_solenoidOff", gateleft2
        Controller.Switch(8) = 0
        Controller.Switch(9) = 1
    End If
End Sub

Sub solRightRampGate(Enabled)
    If Enabled Then
        gateright.RotateToEnd
        PlaySoundAt "fx_solenoidOn", gateright
    Else
        gateright.RotateToStart
        PlaySoundAt "fx_solenoidOff", gateright
    End If
End Sub

Sub solLeftRampGate(Enabled)
    If Enabled Then
        gateleft.RotateToEnd
        PlaySoundAt "fx_solenoidOn", gateleft
    Else
        gateleft.RotateToStart
        PlaySoundAt "fx_solenoidOff", gateleft
    End If
End Sub

Sub SolRightOrbit(Enabled)
    If Enabled Then
        OrbitRight.open = True
        PlaySoundAt "fx_solenoidOn", OrbitRight
    Else
        OrbitRight.open = False
        PlaySoundAt "fx_solenoidOff", OrbitRight
    End If
End Sub

'**************
' Flipper Subs
'**************

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
    PlaySound "fx_rubber_flipper", 0, Vol(ActiveBall), pan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, Vol(ActiveBall), pan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
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
    If LeftFlipper.CurrentAngle >= LeftFlipper.StartAngle - SOSAngle Then LeftFlipper.Strength = FlipperPower * SOSTorque else LeftFlipper.Strength = FlipperPower:End If

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
    If RightFlipper.CurrentAngle <= RightFlipper.StartAngle + SOSAngle Then RightFlipper.Strength = FlipperPower * SOSTorque else RightFlipper.Strength = FlipperPower:End If

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
    RollingTimer.Enabled = 1
End Sub

Sub RollingTimer_Timer()
    Dim BOT, b, ballpitch, ballvol, speedfactorx, speedfactory
    BOT = GetBalls

    ' stop the sound of deleted balls
    For b = UBound(BOT) + 1 to tnob
        rolling(b) = False
        StopSound("fx_ballrolling" & b)
    Next

    ' exit the sub if no balls on the table
    If UBound(BOT) = lob - 1 Then Exit Sub 'there no extra balls on this table

    ' play the rolling sound for each ball and draw the shadow
    For b = lob to UBound(BOT)
        If BallVel(BOT(b)) > 1 Then
            If BOT(b).z < 30 Then
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
        If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
            PlaySound "fx_balldrop", 0, ABS(BOT(b).velz) / 17, Pan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
        End If

        ' jps ball speed & spin control
            BOT(b).AngMomZ = BOT(b).AngMomZ * 0.95
        If BOT(b).VelX AND BOT(b).VelY <> 0 Then
            speedfactorx = ABS(maxvel / BOT(b).VelX)
            speedfactory = ABS(maxvel / BOT(b).VelY)
            If speedfactorx < 1 Then
                BOT(b).VelX = BOT(b).VelX * speedfactorx
                BOT(b).VelY = BOT(b).VelY * speedfactorx
            End If
            If speedfactory < 1 Then
                BOT(b).VelX = BOT(b).VelX * speedfactory
                BOT(b).VelY = BOT(b).VelY * speedfactory
            End If
        End If
    Next
End Sub

'******************
' RealTime Updates
'******************

Sub LeftFlipper_Animate:LeftFlipperTop.RotZ = LeftFlipper.CurrentAngle:End Sub
Sub RightFlipper_Animate:RightFlipperTop.RotZ = RightFlipper.CurrentAngle:End Sub
Sub body_Animate:hulkbody.objRotZ = body.CurrentAngle:hulkarms.objRotZ = body.CurrentAngle:End Sub
Sub arms_Animate:hulkarms.Rotx = arms.CurrentAngle:End Sub
Sub gateleft_Animate:gateleftp.Rotz = gateleft.CurrentAngle:End Sub
Sub gateleft2_Animate:gateleftp2.Rotz = gateleft2.CurrentAngle:End Sub
Sub gateright_Animate:gaterightp.Rotz = gateright.CurrentAngle:End Sub

'*************************
' GI - needs new vpinmame
'*************************

Set GICallback = GetRef("GIUpdate")

Sub GIUpdate(no, Enabled)
    For each x in aGiLights
        x.State = ABS(Enabled)
    Next
End Sub

'****************************************************
'Spinning Tesseract based on the code by sliderpoint
'****************************************************
Dim SpBall(1), sCntrX, sCntrY, Pi, sDegs, sRad, sVel

Pi = Round(4 * Atn(1), 6) '3.14159
sCntrX = CubeB.x
sCntrY = CubeB.y
sRad = sCntrX - 455 'spinner center - left post center
sVel = 0

Set SpBall(0) = SpinKicker.CreateSizedBallWithMass(14, 0.3)
SpBall(0).x = 455
SpBall(0).y = 825
SpBall(0).Visible = False
SpinKicker.Kick 0, 0, 0

Set SpBall(1) = SpinKicker.CreateSizedBallWithMass(14, 0.3)
SpBall(1).x = 569
SpBall(1).y = 825
SpBall(1).Visible = False
SpinKicker.Kick 0, 0, 0

Sub SpinKicker_Timer
    'Immobilize the spinner balls, neg Y velocity offset reflects the timer interval
    SpBall(0).vely = -0.01
    SpBall(1).vely = -0.01
    SpBall(0).velx = 0
    SpBall(1).velx = 0
    SpBall(0).velz = 0.01
    SpBall(1).velz = 0.01
    SpBall(0).z = 25                                      'sPostRadius
    SpBall(1).z = 25                                      'sPostRadius
    sDegs = CubeB.RotZ
    SpBall(0).x = sRad * cos(sDegs * (PI / 180)) + sCntrX 'Place spinner balls to follow 3D CubeB
    SpBall(0).y = sRad * sin(sDegs * (PI / 180)) + sCntrY
    SpBall(1).y = sCntrY -(SpBall(0).y - sCntrY)          'Reverse clone ball(0) movement for ball(1)
    SpBall(1).x = sCntrX -(SpBall(0).x - sCntrX)
    If sVel > 7 Then sVel = 7 End If
    If sVel < -7 Then sVel = -7 End If
    If sVel > .013 Then
        sVel = sVel - .013
    ElseIf sVel < -.013 Then
        sVel = sVel + .013
    Else
        sVel = 0
    End If

    sVel = Round(sVel, 3)
    If CubeB.RotZ <= 0 Then CubeB.RotZ = 360
    If CubeB.RotZ > 360 Then CubeB.RotZ = 1
    CubeB.RotZ = CubeB.RotZ + sVel
    CubePr.RotZ = CubeB.RotZ
    CubePr2.RotZ = CubeB.RotZ
    CubeBase.RotZ = CubeB.RotZ
    LampPost1.X = SpBall(1).x
    LampPost1.Y = SpBall(1).Y
    LampPost2.X = SpBall(0).x
    LampPost2.Y = SpBall(0).Y
End Sub

Dim BallnPlay, SpinBall, RotAdj, PiFilln

Sub OnBallBallCollision(ball1, ball2, velocity)
    ' Determine which ball is which, the Spinner Ball or the Ball-in-play
    If ball1.ID = SpBall(0).ID then  'SpBall(0)'s table ID is 0
        Set SpinBall = SpBall(0)     'or ball1    'Set spinner ball to = ball1
        Set BallnPlay = ball2        'Set ball-in-play to = ball2
        RotAdj = abs(CubeB.RotZ-360) '* Take a sample of the primitive angle and adjust it's reading for calculations *
    ElseIf ball1.ID = SpBall(1).ID then
        Set SpinBall = SpBall(1)
        Set BallnPlay = ball2
        RotAdj = abs(CubeB.RotZ-180)                                         '*
        If abs(CubeB.RotZ-360) < 180 Then RotAdj = abs(CubeB.RotZ-360) + 180 '*
        ElseIf ball2.ID = SpBall(0).ID then
            Set SpinBall = SpBall(0)
            Set BallnPlay = ball1
            RotAdj = abs(CubeB.RotZ-360) '*
        ElseIf ball2.ID = SpBall(1).ID then
            Set SpinBall = SpBall(1)
            Set BallnPlay = ball1
            RotAdj = abs(CubeB.RotZ-180)                                         '*
            If abs(CubeB.RotZ-360) < 180 Then RotAdj = abs(CubeB.RotZ-360) + 180 '*
            Else
                PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
                Exit Sub 'Incase of multi-ball or any other non spinner ball collision, exit sub
    End If
    ' PiFilln accounts for pos/neg values in the collision to provide a proper pos/neg spin velocity
    If SpinBall.X < BallnPlay.X Then PiFilln = Pi Else PiFilln = 0

    ' So basically..... the new spin velocity = old spin velocity +- (spinner angle +- ball collision angle) * collision velocity
    sVel = sVel + sin((RotAdj * Pi / 180)- atn(((SpinBall.Y - BallnPlay.Y) * -1) / (SpinBall.X - BallnPlay.X)) + PiFilln) * velocity / 8
End Sub

' Spinning Tesseract Toy switches
' Use RotatingCubeBall(0).ID to trig below playfield tesseract switches

Sub sw45_Hit
    If ActiveBall.ID = 0 Then vpmtimer.pulsesw 45
End Sub

Sub sw46_Hit
    If ActiveBall.ID = 0 Then vpmtimer.pulsesw 46
End Sub

'*************************************************************
'                        Visible Lock
' Adapted to this table as it uses optops instead of switches
' based on the core.vbs
'*************************************************************

Class cvpmVLock2
    Private mTrig, mKick, mSw(), mSize, mBalls, mGateOpen, mRealForce, mBallSnd, mNoBallSnd
    Public ExitDir, ExitForce, KickForceVar

    Private Sub Class_Initialize
        mBalls = 0:ExitDir = 0:ExitForce = 0:KickForceVar = 0:mGateOpen = False
        vpmTimer.addResetObj Me
    End Sub

    Public Sub InitVLock(aTrig, aKick, aSw)
        Dim ii
        mSize = vpmSetArray(mTrig, aTrig)
        If vpmSetArray(mKick, aKick) <> mSize Then MsgBox "cvpmVLock: Unmatched kick+trig":Exit Sub
        On Error Resume Next
        ReDim mSw(mSize)
        If IsArray(aSw)Then
            For ii = 0 To UBound(aSw):mSw(ii) = aSw(ii):Next
        ElseIf aSw = 0 Or Err Then
            For ii = 0 To mSize:mSw(ii) = mTrig(ii).TimerInterval:Next
        Else
            mSw(0) = aSw
        End If
    End Sub

    Public Sub InitSnd(aBall, aNoBall):mBallSnd = aBall:mNoBallSnd = aNoBall:End Sub
    Public Sub CreateEvents(aName)
        Dim ii
        If Not vpmCheckEvent(aName, Me)Then Exit Sub
        For ii = 0 To mSize
            vpmBuildEvent mTrig(ii), "Hit", aName & ".TrigHit ActiveBall," & ii + 1
            vpmBuildEvent mTrig(ii), "Unhit", aName & ".TrigUnhit ActiveBall," & ii + 1
            vpmBuildEvent mKick(ii), "Hit", aName & ".KickHit " & ii + 1
        Next
    End Sub
    Public Sub SolExit(aEnabled)
        Dim ii
        mGateOpen = aEnabled
        If Not aEnabled Then Exit Sub
        If mBalls > 0 Then PlaySound mBallSnd:Else PlaySound mNoBallSnd:Exit Sub
        For ii = 0 To mBalls-1
            mKick(ii).Enabled = False:If mSw(ii)Then Controller.Switch(mSw(ii)) = False
        Next
        mRealForce = ExitForce + (Rnd - 0.5) * KickForceVar:mKick(0).Kick ExitDir, mRealForce
    End Sub

    Public Sub Reset
        Dim ii:If mBalls = 0 Then Exit Sub
        For ii = 0 To mBalls-1
            If mSw(ii)Then Controller.Switch(mSw(ii)) = True
        Next
    End Sub

    Public Property Get Balls:Balls = mBalls:End Property

    Public Property Let Balls(aBalls)
        Dim ii:mBalls = aBalls
        For ii = 0 To mSize
            If ii >= aBalls Then
                mKick(ii).DestroyBall:If mSw(ii)Then Controller.Switch(mSw(ii)) = False
                Else
                    vpmCreateBall mKick(ii):If mSw(ii)Then Controller.Switch(mSw(ii)) = True
            End If
        Next
    End Property

    Public Sub TrigHit(aBall, aNo)
        aNo = aNo - 1:If mSw(aNo)Then Controller.Switch(mSw(aNo)) = False 'False because it is an opto
        If aBall.VelY < -1 Then Exit Sub                                   ' Allow small upwards speed
        If aNo = mSize Then mBalls = mBalls + 1
        If mBalls > aNo Then mKick(aNo).Enabled = Not mGateOpen
    End Sub

    Public Sub TrigUnhit(aBall, aNo)
        aNo = aNo - 1:If mSw(aNo)Then Controller.Switch(mSw(aNo)) = True 'true because it is en opto and it is on when there is no ball
        If aBall.VelY > -1 Then
            If aNo = 0 Then mBalls = mBalls - 1
            If aNo < mSize Then mKick(aNo + 1).Kick 0, 0
            Else
                If aNo = mSize Then mBalls = mBalls - 1
                If aNo > 0 Then mKick(aNo-1).Kick ExitDir, mRealForce
        End If
    End Sub

    Public Sub KickHit(aNo):mKick(aNo-1).Enabled = False:End Sub
End Class

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

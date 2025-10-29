' Attack from Mars / IPD No. 3781 / December, 1995 / 4 Players
' Based on the table by Bally/Williams 1995
' DOF by arngrim
' VPX 10.8 - version by JPSalas 2024, version 6.0.0

Option Explicit
Randomize

Const Ballsize = 50
Const Ballmass = 1

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const bRotateSaucers = True 'change this if you want the small saucers to rotate or not
Const bRotateBigUFO = True 'the ufo gets out of syc with the dome whan skaing due to the rotation, so best to keep it off

Dim VarHidden, UseVPMColoredDMD
If Table1.ShowDT = true then
    UseVPMColoredDMD = true
    VarHidden = 1
Else
    UseVPMColoredDMD = False
    VarHidden = 0
End If

' Use Modulated Flashers
Const UseVPMModSol = True

LoadVPM "01560000", "WPC.VBS", 3.26

'********************
'Standard definitions
'********************

Const UseSolenoids = 2
Const UseLamps = 1
Const UseSync = 0
Const HandleMech = 0

' special fastflips
Const cSingleLFlip = 0
Const cSingleRFlip = 0

' Standard Sounds
Const SSolenoidOn = "fx_Solenoidon"
Const SSolenoidOff = "fx_solenoidoff"
Const SCoin = "fx_Coin"

' Set GiCallback2 = GetRef("UpdateGI2") 'modulated Gi
Set GiCallback = GetRef("UpdateGI")

Dim bsTrough, bsL, bsR, dtDrop, x, BallFrame, plungerIM, Mech3bank

'******************
' RealTime Updates
'******************

Sub RealTime_Timer
    BigUfoUpdate
    RollingUpdate
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


'************
' Table init.
'************

Const cGameName = "afm_113b" 'arcade rom - with credits
'Const cGameName = "afm_113"  'home rom - free play

Sub Table1_Init
    vpmInit Me
    With Controller
        .GameName = cGameName
  ' NVOffset (2)
        If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Attack from Mars LE" & vbNewLine & "VPX8 table by JPSalas v5.5.2"
        .Games(cGameName).Settings.Value("sound") = 1
        .Games(cGameName).Settings.Value("rol") = 0
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = VarHidden
        On Error Resume Next
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
        On Error Goto 0
        .Switch(22) = 1 'close coin door
        .Switch(24) = 1 'and keep it close
    End With

    ' Nudging
    vpmNudge.TiltSwitch = 14
    vpmNudge.Sensitivity = 1
    vpmNudge.TiltObj = Array(bumper1, bumper2, bumper3, LeftSlingshot, RightSlingshot)

    ' Trough
    Set bsTrough = New cvpmTrough
    With bsTrough
        .size = 4
        '.entrySw = 18
        .initSwitches Array(32, 33, 34, 35)
        .Initexit BallRelease, 80, 6
        .InitExitSounds SoundFX("fx_Solenoid", DOFContactors), SoundFX("fx_ballrel", DOFContactors)
        .Balls = 4
    End With

    ' Droptarget
    Set dtDrop = New cvpmDropTarget
    With dtDrop
        .InitDrop sw77, 77
        .initsnd SoundFX("fx_droptarget", DOFContactors), SoundFX("fx_resetdrop", DOFContactors)
    End With

    ' Left hole
    Set bsL = New cvpmTrough
    With bsL
        .size = 1
        .initSwitches Array(36)
        .Initexit sw36, 180, 0
        .InitExitSounds SoundFX("fx_Solenoid", DOFContactors), SoundFX("fx_popper", DOFContactors)
        .InitExitVariance 3, 2
    End With

    'Right hole
    Set bsR = New cvpmTrough
    With bsR
        .size = 4
        .initSwitches Array(37)
        .Initexit sw37b, 200, 24
        .InitExitSounds SoundFX("fx_Solenoid",DOFContactors), SoundFX("fx_popper",DOFContactors)
        .InitExitVariance 2, 2
        .MaxBallsPerKick = 1
    End With

    '3 Targets Bank
    Set Mech3Bank = new cvpmMech
    With Mech3Bank
        .Sol1 = 24
        .Mtype = vpmMechLinear + vpmMechReverse + vpmMechOneSol
        .Length = 60
        .Steps = 55
        .AddSw 67, 0, 0
        .AddSw 66, 55, 55
        .Callback = GetRef("Update3Bank")
        .Start
    End With

    ' Impulse Plunger
    Const IMPowerSetting = 38 'Plunger Power
    Const IMTime = 0.6        ' Time in seconds for Full Plunge
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swplunger, IMPowerSetting, IMTime
        .Random 0.7
        .switch 18
        .InitExitSnd SoundFX("fx_plunger", DOFContactors), SoundFX("fx_plunger", DOFContactors)
        .CreateEvents "plungerIM"
    End With

    vpmMapLights aLights

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

    ' Init other dropwalls - animations
    UpdateGI 0, 0:UpdateGI 1, 0:UpdateGI 2, 0
    UFORotSpeedSlow
    UfoLed.Enabled = 1
    DivW1.IsDropped = 1

    RealTime.Enabled = 1
    RotateUFO.Enabled = 1

'turn off flashers, small fix for VPX8
f17.state = 2
F25.State = 2
vpmTimer.AddTimer 200, "f17.State = 0: f25.State = 0 '"

End Sub

'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
    If keycode = PlungerKey Then Controller.Switch(11) = 1
    If keycode = LockbarKey Then Controller.Switch(11) = 1:End If
    If keycode = LeftTiltKey Then Nudge 90, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25:aSaucerShake:a3BankShake2
    If keycode = RightTiltKey Then Nudge 270, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25:aSaucerShake:a3BankShake2
    If keycode = CenterTiltKey Then Nudge 0, 7:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25:aSaucerShake:a3BankShake2
    If vpmKeyDown(keycode)Then Exit Sub
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If keycode = PlungerKey Then Controller.Switch(11) = 0
    If keycode = LockbarKey Then Controller.Switch(11) = 0:End If
    If vpmKeyUp(keycode)Then Exit Sub
End Sub

'*********
' Switches
'*********

' Slings & div switches

Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    PlaySoundAt SoundFX("fx_Slingshot", DOFContactors), Lemk
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 51
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
    PlaySoundAt SoundFX("fx_Slingshot", DOFContactors), Remk
    RightSling4.Visible = 1
    Remk.RotX = 26
    RStep = 0
    vpmTimer.PulseSw 52
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

' Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 53:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper1:End Sub

Sub Bumper2_Hit:vpmTimer.PulseSw 54:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper2:End Sub

Sub Bumper3_Hit:vpmTimer.PulseSw 55:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper3:End Sub

' Drain holes, vuks & saucers
Sub Drain_Hit:PlaysoundAt "fx_drain", drain:bsTrough.AddBall Me:End Sub
Sub sw36a_Hit:PlaySoundAt "fx_kicker_enter", sw36a:bsL.AddBall Me:End Sub
Sub sw37a_Hit:PlaySoundAt "fx_hole_enter", sw37a:bsR.AddBall Me:End Sub
Sub sw78_Hit:vpmTimer.PulseSw 78:PlaySoundAt "fx_hole_enter", sw78:bsL.AddBall Me:End Sub
Sub sw37_Hit:PlaySoundAt "fx_hole_enter", sw37:bsR.AddBall Me:End Sub

' Rollovers & Ramp Switches
Sub sw16_Hit:Controller.Switch(16) = 1:PlaySoundAt "fx_sensor", sw16:End Sub
Sub sw16_UnHit:Controller.Switch(16) = 0:End Sub

Sub sw26_Hit:Controller.Switch(26) = 1:PlaySoundAt "fx_sensor", sw26:End Sub
Sub sw26_UnHit:Controller.Switch(26) = 0:End Sub

Sub sw17_Hit:Controller.Switch(17) = 1:PlaySoundAt "fx_sensor", sw17:End Sub
Sub sw17_UnHit:Controller.Switch(17) = 0:End Sub

Sub sw27_Hit:Controller.Switch(27) = 1:PlaySoundAt "fx_sensor", sw27:End Sub
Sub sw27_UnHit:Controller.Switch(27) = 0:End Sub

Sub sw38_Hit:Controller.Switch(38) = 1:PlaySoundAt "fx_sensor", sw38:End Sub
Sub sw38_Unhit:Controller.Switch(38) = 0:End Sub

Sub sw48_Hit:Controller.Switch(48) = 1:PlaySoundAt "fx_sensor", sw48:End Sub
Sub sw48_Unhit:Controller.Switch(48) = 0:End Sub

Sub sw71_Hit:Controller.Switch(71) = 1:PlaySoundAt "fx_sensor", sw71:End Sub
Sub sw71_UnHit:Controller.Switch(71) = 0:End Sub

Sub sw72_Hit:Controller.Switch(72) = 1:PlaySoundAt "fx_sensor", sw72:End Sub
Sub sw72_Unhit:Controller.Switch(72) = 0:End Sub

Sub sw73_Hit:Controller.Switch(73) = 1:PlaySoundAt "fx_sensor", sw73:End Sub
Sub sw73_Unhit:Controller.Switch(73) = 0:End Sub

Sub sw74_Hit:Controller.Switch(74) = 1:PlaySoundAt "fx_sensor", sw74:End Sub
Sub sw74_Unhit:Controller.Switch(74) = 0:End Sub

Sub sw61_Hit:Controller.Switch(61) = 1:End Sub
Sub sw61_Unhit:Controller.Switch(61) = 0:End Sub

Sub sw62_Hit:Controller.Switch(62) = 1:End Sub
Sub sw62_Unhit:Controller.Switch(62) = 0:End Sub

Sub sw63_Hit:Controller.Switch(63) = 1:End Sub
Sub sw63_Unhit:Controller.Switch(63) = 0:End Sub

Sub sw64_Hit:Controller.Switch(64) = 1:End Sub
Sub sw64_Unhit:Controller.Switch(64) = 0:End Sub

Sub sw65_Hit:Controller.Switch(65) = 1:End Sub
Sub sw65_Unhit:Controller.Switch(65) = 0:End Sub

' Targets
Sub sw56_Hit:vpmTimer.PulseSw 56:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

Sub sw57_Hit:vpmTimer.PulseSw 57:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

Sub sw58_Hit:vpmTimer.PulseSw 58:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

Sub sw41_Hit:vpmTimer.PulseSw 41:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

Sub sw42_Hit:vpmTimer.PulseSw 42:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

Sub sw43_Hit:vpmTimer.PulseSw 43:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

Sub sw44_Hit:vpmTimer.PulseSw 44:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

Sub sw75_Hit:vpmTimer.PulseSw 75:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

Sub sw76_Hit:vpmTimer.PulseSw 76:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

' 3 bank
Sub sw45_Hit:vpmTimer.PulseSw 45:a3BankShake:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

Sub sw46_Hit:vpmTimer.PulseSw 46:a3BankShake:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

Sub sw47_Hit:vpmTimer.PulseSw 47:a3BankShake:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

' Droptarget

Sub sw77_Hit:dtDrop.Hit 1:UFORotSpeedFast:End Sub

'*********
'Solenoids
'*********

SolCallback(1) = "Auto_Plunger"
SolCallback(2) = "SolRelease"
SolCallback(3) = "bsL.SolOut"
SolCallback(4) = "Sol4" ' bsR.SolOut
SolCallback(5) = "SolAlien5"
SolCallback(6) = "SolAlien6"
SolCallback(7) = "vpmSolSound SoundFX(""fx_Knocker"",DOFKnocker),"
SolCallback(8) = "SolAlien8"
SolCallback(14) = "SolAlien14"
SolCallBack(15) = "SolUfoShake"
SolCallback(16) = "SolDropTargetUp"
'SolCallback(24) = "SolBank" 'used in the Mech
SolCallback(33) = "vpmSolGate RGate,false,"
SolCallback(34) = "vpmSolGate LGate,false,"
SolCallback(36) = "SolDiv"
SolCallback(43) = "vpmFlasher f30,"

If UseVPMModSol Then
    SolModCallback(17) = "Flasher17"
    SolModCallback(18) = "Flasher18"
    SolModCallback(19) = "Flasher19"
    SolModCallback(20) = "Flasher20"
    SolModCallback(21) = "Flasher21"
    SolModCallback(22) = "Flasher22"
    SolModCallback(23) = "Flasher23"
    SolModCallback(25) = "Flasher25"
    SolModCallback(26) = "Flasher26"
    SolModCallback(27) = "Flasher27"
    SolModCallback(28) = "Flasher28"
    f17.Fader = 0
    f18.Fader = 0
    f19.Fader = 0
    f20.Fader = 0
    f21.Fader = 0
    f22.Fader = 0
    f23.Fader = 0:f23a.Fader = 0
    f25.Fader = 0
    f26.Fader = 0
    f27.Fader = 0
    f28.Fader = 0
Else
    SolCallback(17) = "vpmFlasher f17,"
    SolCallback(18) = "vpmFlasher f18,"
    SolCallback(19) = "vpmFlasher f19,"
    SolCallback(20) = "vpmFlasher f20,"
    SolCallback(21) = "vpmFlasher f21,"
    SolCallback(22) = "vpmFlasher f22,"
    SolCallback(23) = "vpmFlasher Array(F23,f23a)," ' SolUfoFlash"
    SolCallback(25) = "vpmFlasher f25,"
    SolCallback(26) = "vpmFlasher f26,"
    SolCallback(27) = "vpmFlasher f27,"
    SolCallback(28) = "vpmFlasher f28,"
    f17.Fader = 2
    f18.Fader = 2
    f19.Fader = 2
    f20.Fader = 2
    f21.Fader = 2
    f22.Fader = 2
    f23.Fader = 2:f23a.Fader = 2
    f25.Fader = 2
    f26.Fader = 2
    f27.Fader = 2
    f28.Fader = 2
End If

Sub Flasher17(m):m = m / 255:f17.State = m:End Sub
Sub Flasher18(m):m = m / 255:f18.State = m:End Sub
Sub Flasher19(m):m = m / 255:f19.State = m:End Sub
Sub Flasher20(m):m = m / 255:f20.State = m:End Sub
Sub Flasher21(m):m = m / 255:f21.State = m:End Sub
Sub Flasher22(m):m = m / 255:f22.State = m:End Sub
Sub Flasher23(m):m = m / 255:f23.State = m:f23a.State = m:End Sub
Sub Flasher25(m):m = m / 255:f25.State = m:End Sub
Sub Flasher26(m):m = m / 255:f26.State = m:End Sub
Sub Flasher27(m):m = m / 255:f27.State = m:End Sub
Sub Flasher28(m):m = m / 255:f28.State = m:End Sub
Sub Flasher43(m):m = m / 255:f30.State = m:End Sub

Sub SolRelease(Enabled)
    If Enabled And bsTrough.Balls> 0 Then
        vpmTimer.PulseSw 31
        bsTrough.ExitSol_On
    End If
End Sub

Sub Sol4(Enabled) 'bsR
    sw37a.TimerEnabled = 1
End Sub

Sub sw37a_Timer
    If bsR.Balls Then
        bsR.ExitSol_On
    Else
        sw37a.TimerEnabled = 0
    End If
End Sub

Sub Auto_Plunger(Enabled)
    If Enabled Then
        PlungerIM.AutoFire
    End If
End Sub

Sub SolDropTargetUp(Enabled)
    If Enabled Then
        dtDrop.DropSol_On
    End If
End Sub

'****************
' Alien solenoids
'****************

Sub SolAlien5(Enabled)
    If Enabled Then
        Alien5.TransZ = 20
        PlaySoundAt "fx_solenoidon", Alien5
    Else
        Alien5.TransZ = 0
        PlaySoundAt "fx_solenoidoff", Alien5
    End If
End Sub

Sub SolAlien6(Enabled)
    If Enabled Then
        Alien6.TransZ = 20
        PlaySoundAt "fx_solenoidon", Alien6
    Else
        Alien6.TransZ = 0
        PlaySoundAt "fx_solenoidoff", Alien6
    End If
End Sub

Sub SolAlien8(Enabled)
    If Enabled Then
        Alien8.TransZ = 20
        PlaySoundAt "fx_solenoidon", Alien8
    Else
        Alien8.TransZ = 0
        PlaySoundAt "fx_solenoidoff", Alien8
    End If
End Sub

Sub SolAlien14(Enabled)
    If Enabled Then
        Alien14.TransZ = 20
        PlaySoundAt "fx_solenoidon", Alien14
    Else
        Alien14.TransZ = 0
        PlaySoundAt "fx_solenoidoff", Alien14
    End If
End Sub

'***********
' Diverter
'***********

Sub SolDiv(Enabled)
    PlaySoundAt SoundFX("fx_diverter", DOFContactors), DivF
    If Enabled Then
        DivF.RotateToEnd
        DivW1.IsDropped = 0
        DivW2.IsDropped = 1
    Else
        DivF.RotateToStart
        DivW1.IsDropped = 1
        DivW2.IsDropped = 0
    End If
End Sub

Sub DivF_Animate
    DivP.RotZ = DivF.CurrentAngle
End Sub

'*******************
' Flipper Subs v3.0
'*******************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
    If UseSolenoids = 2 then Controller.Switch(swULFlip) = Enabled
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
    if UseSolenoids = 2 then Controller.Switch(swURFlip) = Enabled
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

Sub LeftFlipper_Animate()
  LeftFlipperTop.RotZ = LeftFlipper.CurrentAngle
End Sub

Sub RightFlipper_Animate()
  RightFlipperTop.RotZ =RightFlipper.CurrentAngle
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
End Sub

'******************
'Motor Bank Up Down
'******************

Sub Update3Bank(currpos, currspeed, lastpos)
    If currpos <> lastpos Then
        PlaySound "fx_motor2"
        BackBank.Z = - currpos
        swp45.Z = -(22 + currpos)
        swp46.Z = -(22 + currpos)
        swp47.Z = -(22 + currpos)
    End If
    If currpos> 40 Then
        sw45.Isdropped = 1
        sw46.Isdropped = 1
        sw47.Isdropped = 1
        UFORotSpeedMedium
    End If
    If currpos <10 Then
        sw45.Isdropped = 0
        sw46.Isdropped = 0
        sw47.Isdropped = 0
        UFORotSpeedSlow
    End If
End Sub

'***********
' Update GI
'***********

Sub UpdateGI2(no, step)
    Dim gistep, ii, a
    gistep = step / 8
    Select Case no
        Case 0
            For each ii in aGiLLights
                ii.IntensityScale = gistep
            Next
        Case 1
            For each ii in aGiMLights
                ii.IntensityScale = gistep
            Next
        Case 2 ' also the bumpers er GI
            For each ii in aGiTLights
                ii.IntensityScale = gistep
            Next
    End Select
End Sub

Sub UpdateGI(no, step)
    Dim gistep, ii, a
    gistep = ABS(step)
    Select Case no
        Case 0
            For each ii in aGiLLights
                ii.state = gistep
            Next
        Case 1
            For each ii in aGiMLights
                ii.state = gistep
            Next
        Case 2 ' also the bumpers er GI
            For each ii in aGiTLights
                ii.state = gistep
            Next
    End Select
End Sub

'*************
' 3Bank Shake
'*************

Dim ccBall
Const cMod = .65 'percentage of hit power transfered to the 3 Bank of targets

a3BankInit

Sub a3BankShake
    ccball.velx = activeball.velx * cMod
    ccball.vely = activeball.vely * cMod
    a3BankTimer.enabled = True
    b3BankTimer.enabled = True
End Sub

Sub a3BankShake2 'when nudging
    a3BankTimer.enabled = True
    b3BankTimer.enabled = True
End Sub

Sub a3BankInit
    Set ccBall = hball.createball
    hball.Kick 0, 0
End Sub

Sub a3BankTimer_Timer            'start animation
    Dim x, y
    x = (hball.x - ccball.x) / 4 'reduce the X axis movement
    y = (hball.y - ccball.y) / 2
    backbank.transy = x
    backbank.transx = - y
    swp45.transy = x
    swp45.transx = - y
    swp46.transy = x
    swp46.transx = - y
    swp47.transy = x
    swp47.transx = - y
End Sub

Sub b3BankTimer_Timer 'stop animation
    backbank.transx = 0
    backbank.transy = 0
    swp45.transz = 0
    swp45.transx = 0
    swp46.transz = 0
    swp46.transx = 0
    swp47.transz = 0
    swp47.transx = 0
    a3BankTimer.enabled = False
    b3BankTimer.enabled = False
End Sub

'***************
' Big UFO Shake
'***************
Dim cBall, UFOLedPos
BigUfoInit

Sub BigUfoInit
    UFOLedPos = 0
    Set cBall = ckicker.createball
    ckicker.Kick 0, 0
End Sub

Sub SolUfoShake(Enabled)
    If Enabled Then
        BigUfoShake
    End If
End Sub

Sub BigUfoShake
    cball.velx = -2 + 2 * RND(1)
    cball.vely = -10 + 2 * RND(1)
End Sub

Sub UFOLed_Timer()
    Select Case UfoLedPos
        Case 0,3,6:ufo1.image = "bigufo1"
        Case 1,4,7:ufo1.image = "bigufo2"
        Case 2,5,8:ufo1.image = "bigufo3"
        Case 9,12,15:ufo1.image = "bigufo4"
        Case 10,13,16:ufo1.image = "bigufo5"
        Case 11,14,17:ufo1.image = "bigufo6"
        Case 18,21,24:ufo1.image = "bigufo7"
        Case 19,22,25:ufo1.image = "bigufo8"
        Case 20,23,26:ufo1.image = "bigufo9"
    End Select
UfoLedPos = (UfoLedPos + 1) MOD 26
End Sub

Sub BigUfoUpdate
    Dim a, b
    a = ckicker.y - cball.y
    b = cball.x - ckicker.x

    Ufo1.rotx = - a
    Ufo1.transy = a
    Ufo1.roty = b
    Ufo1d.rotx = - a
    Ufo1d.transy = a
    Ufo1d.roty = b
End Sub

'**********************************
'   Small UFOs Rotation Speed
'**********************************

Sub UFORotSpeedSlow()
    UfoLed.Interval = 400
    RotateUFO.Interval = 300
End Sub

Sub UFORotSpeedMedium()
    UfoLed.Interval = 300
    RotateUFO.Interval =150
End Sub

Sub UFORotSpeedFast()
    UfoLed.Interval = 200
    RotateUFO.Interval = 75
End Sub

Dim UFOSmallPos
UFOSmallPos = 0

Sub RotateUFO_Timer()
If bRotateSaucers Then
    Ufo2.RotZ = (Ufo2.RotZ - 1) MOD 360
    Ufo4.RotZ = (Ufo4.RotZ - 1) MOD 360
    Ufo5.RotZ = (Ufo5.RotZ - 1) MOD 360
    Ufo6.RotZ = (Ufo6.RotZ - 1) MOD 360
    Ufo7.RotZ = (Ufo7.RotZ - 1) MOD 360
    Ufo8.RotZ = (Ufo8.RotZ - 1) MOD 360
End If
If bRotateBigUFO Then
    Ufo1.RotZ = (Ufo1.RotZ - 1) MOD 360
    Ufo1d.RotZ = Ufo1.RotZ
End If

    Select Case UFOSmallPos
        Case 0,8,16:ufo2.image = "saucer1":ufo4.image = "saucer4":ufo5.image = "saucer5":ufo6.image = "saucer6":ufo7.image = "saucer7":ufo8.image = "saucer8"
        Case 1,9,17:ufo2.image = "saucer2":ufo4.image = "saucer5":ufo5.image = "saucer6":ufo6.image = "saucer7":ufo7.image = "saucer8":ufo8.image = "saucer1"
        Case 2,10,18:ufo2.image = "saucer3":ufo4.image = "saucer6":ufo5.image = "saucer7":ufo6.image = "saucer8":ufo7.image = "saucer1":ufo8.image = "saucer2"
        Case 3,11,19:ufo2.image = "saucer4":ufo4.image = "saucer7":ufo5.image = "saucer8":ufo6.image = "saucer1":ufo7.image = "saucer2":ufo8.image = "saucer3"
        Case 4,12,20:ufo2.image = "saucer5":ufo4.image = "saucer8":ufo5.image = "saucer1":ufo6.image = "saucer2":ufo7.image = "saucer3":ufo8.image = "saucer4"
        Case 5,13,21:ufo2.image = "saucer6":ufo4.image = "saucer1":ufo5.image = "saucer2":ufo6.image = "saucer3":ufo7.image = "saucer4":ufo8.image = "saucer5"
        Case 6,14,22:ufo2.image = "saucer7":ufo4.image = "saucer2":ufo5.image = "saucer3":ufo6.image = "saucer4":ufo7.image = "saucer5":ufo8.image = "saucer6"
        Case 7,15,23:ufo2.image = "saucer8":ufo4.image = "saucer3":ufo5.image = "saucer4":ufo6.image = "saucer5":ufo7.image = "saucer6":ufo8.image = "saucer7"

        Case 24,32,40:ufo2.image = "saucer9":ufo4.image = "saucer12":ufo5.image = "saucer13":ufo6.image = "saucer14":ufo7.image = "saucer15":ufo8.image = "saucer16"
        Case 25,33,41:ufo2.image = "saucer10":ufo4.image = "saucer13":ufo5.image = "saucer14":ufo6.image = "saucer15":ufo7.image = "saucer16":ufo8.image = "saucer9"
        Case 26,34,42:ufo2.image = "saucer11":ufo4.image = "saucer14":ufo5.image = "saucer15":ufo6.image = "saucer16":ufo7.image = "saucer9":ufo8.image = "saucer10"
        Case 27,35,43:ufo2.image = "saucer12":ufo4.image = "saucer15":ufo5.image = "saucer16":ufo6.image = "saucer9":ufo7.image = "saucer10":ufo8.image = "saucer11"
        Case 28,36,44:ufo2.image = "saucer13":ufo4.image = "saucer16":ufo5.image = "saucer9":ufo6.image = "saucer10":ufo7.image = "saucer11":ufo8.image = "saucer12"
        Case 29,37,45:ufo2.image = "saucer14":ufo4.image = "saucer9":ufo5.image = "saucer10":ufo6.image = "saucer11":ufo7.image = "saucer12":ufo8.image = "saucer13"
        Case 30,38,46:ufo2.image = "saucer15":ufo4.image = "saucer10":ufo5.image = "saucer11":ufo6.image = "saucer12":ufo7.image = "saucer13":ufo8.image = "saucer14"
        Case 31,39,47:ufo2.image = "saucer16":ufo4.image = "saucer11":ufo5.image = "saucer12":ufo6.image = "saucer13":ufo7.image = "saucer14":ufo8.image = "saucer15"
    End Select
UFOSmallPos = (UFOSmallPos + 1) MOD 47
End Sub

'**********************************************************
' Small shake of small Ufos and other objects when nudging
'**********************************************************

Dim SmallShake:SmallShake = 0

Sub aSaucerShake
    SmallShake = 6
    SaucerShake.Enabled = True
End Sub

Sub SaucerShake_Timer
    ufo2.Rotx = SmallShake
    ufo2a.Rotx = - SmallShake
    ufo7.Rotx = SmallShake
    ufo7a.Rotx = - SmallShake
    ufo8.Rotx = SmallShake
    ufo8a.Rotx = - SmallShake
    ufo6.Rotx = SmallShake
    ufo6a.Rotx = - SmallShake
    ufo5.Rotx = SmallShake
    ufo5a.Rotx = - SmallShake
    ufo4.Rotx = SmallShake
    ufo4a.Rotx = - SmallShake
    alien5.Transz = SmallShake / 2
    alien6.Transz = SmallShake / 2
    alien8.Transz = SmallShake / 2
    alien14.Transz = SmallShake / 2
    If SmallShake = 0 Then SaucerShake.Enabled = False:Exit Sub
    If SmallShake <0 Then
        SmallShake = ABS(SmallShake)- 0.1
    Else
        SmallShake = - SmallShake + 0.1
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

'***********************************
'   JP's VP10 Rolling Sounds
'   JP's Ball Speed Control
'   Rothbauer's dropping sounds
'***********************************

Const tnob = 19   'total number of balls
Const lob = 2     'number of locked balls
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

    ' stop the sound of deleted balls and hide the shadow
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
                ballpitch = Pitch(BOT(b)) + 25000 'increase the pitch on a ramp
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

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
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

' Thalamus : Exit in a clean and proper way
Sub Table1_exit
  Controller.Pause = False
  Controller.Stop
End Sub


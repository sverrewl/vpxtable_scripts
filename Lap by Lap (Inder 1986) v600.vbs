' Lap by Lap / IPD No. 4098 / Inder 1986 / 4 Players
' VPX8 version 6.0.0 by jpsalas
' Script based on Destruk's script

Option Explicit
Randomize

Const BallSize = 50
Const BallMass = 1

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Dim bsTrough, bsR, x, CarPos

Const cGameName = "lapbylap"
Const VarHidden = True 'set it to False if you want to see the vpinmame DMD.

If Table1.ShowDT = true then
For each x in aReels
    x.Visible = True
Next
Else
For each x in aReels
    x.Visible = False
Next
End If

 LoadVPM "01570000", "inder.vbs", 3.27

'********************
'Standard definitions
'********************

Const UseSolenoids = 2
Const UseLamps = 0
Const UseGI = 0
Const UseSync = 0
Const HandleMech = 0

' Standard Sounds
Const SSolenoidOn = "fx_SolenoidOn"
Const SSolenoidOff = "fx_SolenoidOff"
Const SCoin = "fx_Coin"

'************
' Table init.
'************

Sub table1_Init
    vpmInit me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Lap by Lap - Inder 1986" & vbNewLine & "VPX table by JPSalas v6.0.0"
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly =  1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = VarHidden
        .Games(cGameName).Settings.Value("rol") = 0   '1= rotated display, 0= normal
        .Games(cGameName).Settings.Value("sound") = 1 '1 enabled rom sound
        '.SetDisplayPosition 0,0,GetPlayerHWnd 'uncomment if you can't see the dmd
        'On Error Resume Next
        Controller.SolMask(0) = 0
        vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the Timer to renable all the solenoids after 2 seconds
        Controller.Run GetPlayerHWnd
        'On Error Goto 0
    End With

    ' Nudging
    vpmNudge.TiltSwitch = 53
    vpmNudge.Sensitivity = 3
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingshot)

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
         .InitSw 0, 91, 0, 0, 0, 0, 0, 0
        .InitKick BallRelease, 90, 4
        .InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .Balls = 1
    End With

    ' Saucer
     Set bsR = New cvpmBallStack
     With bsR
         .InitSaucer sw70, 70, 130, 4
        .InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .KickForceVar = 2
        .KickAngleVar = 2
    End With

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1
    RealTime.Enabled = 1
    LedsTimer.Enabled = 1

    ' Turn on Gi
    vpmtimer.addtimer 1500, "GiOn '"
End Sub

'******************
' RealTime Updates
'******************

Sub RealTime_Timer
    RollingUpdate
End Sub

'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
    If keycode = LeftTiltKey Then Nudge 90, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If vpmKeyDown(keycode)Then Exit Sub
    If keycode = PlungerKey Then PlaySoundat "fx_PlungerPull", Plunger:Plunger.Pullback
    If keycode = KeyRules Then Rules
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If vpmKeyUp(keycode)Then Exit Sub
    If keycode = PlungerKey Then PlaySoundAt "fx_plunger", Plunger:Plunger.Fire
End Sub


Sub GiOn
    Dim bulb
  PlaySound"fx_gion"
    For each bulb in aGiLights
        bulb.State = 1
    Next
End Sub

Sub GiOff
    Dim bulb
  PlaySound"fx_gioff"
    For each bulb in aGiLights
        bulb.State = 0
    Next
End Sub

'*********
' Switches
'*********

' Slings
Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    PlaySoundAt SoundFX("fx_slingshot", DOFContactors), Lemk
    'DOF 101, DOFPulse
    LeftSling004.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 90
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 1:LeftSLing004.Visible = 0:LeftSLing003.Visible = 1:Lemk.RotX = 14
        Case 2:LeftSLing003.Visible = 0:LeftSLing002.Visible = 1:Lemk.RotX = 2
        Case 3:LeftSLing002.Visible = 0:Lemk.RotX = -20:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
    PlaySoundAt SoundFX("fx_slingshot", DOFContactors), Remk
    'DOF 102, DOFPulse
    RightSling004.Visible = 1
    Remk.RotX = 26
    RStep = 0
    vpmTimer.PulseSw 80
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 1:RightSLing004.Visible = 0:RightSLing003.Visible = 1:Remk.RotX = 14
        Case 2:RightSLing003.Visible = 0:RightSLing002.Visible = 1:Remk.RotX = 2
        Case 3:RightSLing002.Visible = 0:Remk.RotX = -20:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

' Scoring rubbers

Sub rsband002_Hit: vpmTimer.PulseSw 92: End Sub
Sub rsband008_Hit: vpmTimer.PulseSw 92: End Sub
Sub rsband001_Hit: vpmTimer.PulseSw 93: End Sub
Sub rsband007_Hit: vpmTimer.PulseSw 93: End Sub

' Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 76:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper1:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 96:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper2:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 86:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper3:End Sub

 Sub AdvanceCar ()
     If B2SOn Then Controller.B2SSetData CarPos+100, 0
     CarPos = CarPos + 1
     If CarPos> 17 Then CarPos = CarPos - 17 : vpmTimer.PulseSw 61
     Car.SetValue CarPos
     If B2SOn Then Controller.B2SSetData CarPos+100, 1
 End Sub

' Drain & holes
Sub Drain_Hit:PlaysoundAt "fx_drain", Drain:bsTrough.AddBall Me:End Sub
Sub sw70_Hit:PlaysoundAt "fx_kicker_enter", sw70: bsR.AddBall 0:End Sub

' Rollovers
Sub sw62_Hit:Controller.Switch(62) = 1:PlaySoundAt "fx_sensor", sw62:End Sub
Sub sw62_UnHit:Controller.Switch(62) = 0:End Sub

Sub sw97_Hit:Controller.Switch(97) = 1:PlaySoundAt "fx_sensor", sw97:End Sub
Sub sw97_UnHit:Controller.Switch(97) = 0:End Sub

Sub sw77_Hit:Controller.Switch(77) = 1:PlaySoundAt "fx_sensor", sw77:End Sub
Sub sw77_UnHit:Controller.Switch(77) = 0:End Sub

Sub sw67_Hit:Controller.Switch(67) = 1:PlaySoundAt "fx_sensor", sw67:End Sub
Sub sw67_UnHit:Controller.Switch(67) = 0:End Sub

Sub sw87_Hit:Controller.Switch(87) = 1:PlaySoundAt "fx_sensor", sw87:End Sub
Sub sw87_UnHit:Controller.Switch(87) = 0:End Sub

Sub sw63_Hit:Controller.Switch(63) = 1:PlaySoundAt "fx_sensor", sw63:End Sub
Sub sw63_UnHit:Controller.Switch(63) = 0:End Sub

Sub sw83_Hit:Controller.Switch(83) = 1:PlaySoundAt "fx_sensor", sw83:End Sub
Sub sw83_UnHit:Controller.Switch(83) = 0:End Sub

Sub sw82_Hit:Controller.Switch(82) = 1:PlaySoundAt "fx_sensor", sw82:End Sub
Sub sw82_UnHit:Controller.Switch(82) = 0:End Sub

Sub sw81_Hit:Controller.Switch(81) = 1:PlaySoundAt "fx_sensor", sw81:End Sub
Sub sw81_UnHit:Controller.Switch(81) = 0:End Sub

Sub sw71_Hit:Controller.Switch(71) = 1:topsw.visible = 0:PlaySoundAt "fx_sensor", sw71:End Sub
Sub sw71_UnHit:Controller.Switch(71) = 0:topsw.visible = 1:End Sub

'Spinners

Sub Spinner1_Spin():vpmTimer.PulseSw 73:PlaySoundAt "fx_spinner", Spinner1:End Sub
Sub Spinner2_Spin():vpmTimer.PulseSw 72:PlaySoundAt "fx_spinner", Spinner2:End Sub

'Targets
Sub sw60_Hit:vpmTimer.PulseSw 60:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw64_Hit:vpmTimer.PulseSw 64:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw65_Hit:vpmTimer.PulseSw 65:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw74_Hit:vpmTimer.PulseSw 74:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw75_Hit:vpmTimer.PulseSw 75:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw84_Hit:vpmTimer.PulseSw 84:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw85_Hit:vpmTimer.PulseSw 85:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw94_Hit:vpmTimer.PulseSw 94:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw95_Hit:vpmTimer.PulseSw 95:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

'*********
'Solenoids
'*********
SolCallback(4) = "bsTrough.SolOut"
SolCallback(5) = "vpmNudge.SolGameOn"

'*******************
' Flipper Subs Rev3
'*******************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("fx_flipperup", DOFFlippers), LeftFlipper
        LeftFlipper.RotateToEnd
        LeftFlipperOn = 1
    Else
        PlaySoundAt SoundFX("fx_flipperdown", DOFFlippers), LeftFlipper
        LeftFlipper.RotateToStart
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

Sub LeftFlipper_Animate: LeftFlipperTop.RotZ = LeftFlipper.CurrentAngle: End Sub
Sub RightFlipper_Animate: RightFlipperTop.RotZ = RightFlipper.CurrentAngle: End Sub

' flippers hit Sound

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, Vol(ActiveBall), pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, Vol(ActiveBall), pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
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
'     JP's Flasher Fading for VPX and Vpinmame v3.0
'       (Based on Pacdude's Fading Light System)
' This is a fast fading for the Flashers in vpinmame tables
'  just 4 steps, like in Pacdude's original script.
' Included the new Modulated flashers & Lights for WPC
'**********************************************************

 Set LampCallback = GetRef ("UpdateLamps")
Dim NewBonus4, OldBonus4
NewBonus4 = 0
OldBonus4 = 0
' vpinmame Lamp & Flasher Timers

 Sub UpdateLamps ()
     Dim ChgLamp, ii
     ChgLamp = Controller.ChangedLamps
     If Not IsEmpty (ChgLamp) Then
         For ii = 0 To UBound (ChgLamp)
             DisplayLamps chgLamp (ii, 0), chgLamp (ii, 1)
         Next
     End If

     NewBonus4 = ABS (Controller.Lamp (1) ) + ABS (Controller.Lamp (2) ) * 2 + ABS (Controller.Lamp (3) ) * 4 + ABS (Controller.Lamp (4) ) * 8
     If NewBonus4 <> OldBonus4 Then
         Select Case NewBonus4
             Case 0 : L1.State = 1 : L2.State = 0 : L3.State = 0 : L4.State = 0 : L5.State = 0 : L6.State = 0 : L7.State = 0 : L8.State = 0 : L9.State = 0 : L10.State = 0
             Case 1 : L1.State = 0 : L2.State = 1 : L3.State = 0 : L4.State = 0 : L5.State = 0 : L6.State = 0 : L7.State = 0 : L8.State = 0 : L9.State = 0 : L10.State = 0
             Case 2 : L1.State = 0 : L2.State = 0 : L3.State = 1 : L4.State = 0 : L5.State = 0 : L6.State = 0 : L7.State = 0 : L8.State = 0 : L9.State = 0 : L10.State = 0
             Case 3 : L1.State = 0 : L2.State = 0 : L3.State = 0 : L4.State = 1 : L5.State = 0 : L6.State = 0 : L7.State = 0 : L8.State = 0 : L9.State = 0 : L10.State = 0
             Case 4 : L1.State = 0 : L2.State = 0 : L3.State = 0 : L4.State = 0 : L5.State = 1 : L6.State = 0 : L7.State = 0 : L8.State = 0 : L9.State = 0 : L10.State = 0
             Case 5 : L1.State = 0 : L2.State = 0 : L3.State = 0 : L4.State = 0 : L5.State = 0 : L6.State = 1 : L7.State = 0 : L8.State = 0 : L9.State = 0 : L10.State = 0
             Case 6 : L1.State = 0 : L2.State = 0 : L3.State = 0 : L4.State = 0 : L5.State = 0 : L6.State = 0 : L7.State = 1 : L8.State = 0 : L9.State = 0 : L10.State = 0
             Case 7 : L1.State = 0 : L2.State = 0 : L3.State = 0 : L4.State = 0 : L5.State = 0 : L6.State = 0 : L7.State = 0 : L8.State = 1 : L9.State = 0 : L10.State = 0
             Case 8 : L1.State = 0 : L2.State = 0 : L3.State = 0 : L4.State = 0 : L5.State = 0 : L6.State = 0 : L7.State = 0 : L8.State = 0 : L9.State = 1 : L10.State = 0
             Case 9 : L1.State = 0 : L2.State = 0 : L3.State = 0 : L4.State = 0 : L5.State = 0 : L6.State = 0 : L7.State = 0 : L8.State = 0 : L9.State = 0 : L10.State = 1
             Case 15 : L1.State = 0 : L2.State = 0 : L3.State = 0 : L4.State = 0 : L5.State = 0 : L6.State = 0 : L7.State = 0 : L8.State = 0 : L9.State = 0 : L10.State = 0
         End Select

         OldBonus4 = NewBonus4
     End If
 End Sub

 Sub DisplayLamps (idx, Stat)
 Select Case idx
    '1-4=bonus lamps 1 through 10
  Case 5:l11.State = Stat '25,000 Diana Sup
  Case 6:l12.State = Stat '50,000 Diana Sup
  Case 7:l13.State = Stat '100,000 Diana Sup
  Case 8:l14.State = Stat '200,000 Diana Sup
  Case 10:l16.State = Stat '10,000 Diana Sup
  Case 13:l18.State=stat '1st Extra Ball (en Cabeza)
  Case 14:l19.State=stat '2nd Extra Ball (en Cabeza)
  Case 15:l20.State=stat '3rd Extra Ball (en Cabeza)
  Case 16:If Stat=1 Then AdvanceCar 'Motor Car
  Case 18:l21.State = Stat 'Special Diana Sup
  Case 19:l22.State = Stat 'Extra Ball Right Inlane
  Case 20:l23.State = Stat 'Extra Ball Left Inlane
  Case 21:l24.State = Stat 'Special Right Inlane
  Case 22:l25.State = Stat 'Special Left Inlane
  Case 25:l26.State = Stat 'Bonus X2
  Case 26:l27.State = Stat 'Bonus X3
  Case 27:l28.State = Stat 'Bonus X4
  Case 28:l29.State = Stat 'Bonus X5
  Case 29:l30.State = Stat 'Extra Ball Top Left
  Case 30:l31.State = Stat 'Extra Ball Ramp
  Case 31:l32.State = Stat 'Extra Ball Top Right
  Case 32:l33.State=stat 'Game Over (en Cabeza)
  Case 33:l34.State=stat 'Pulsador Partidas (en Trampilla)
 '  Case 34: 'Bobina Monedero (en Trampilla)
  Case 35:l36.State=stat 'Ball In Play
  Case 36:l37.State=stat 'Match
  Case 38:if bsr.balls then bsr.exitsol_on
  Case 39:l38.State=stat ' Handicap
  Case 41:l39.State = Stat: l39b.State = Stat
  Case 42:l40.State = Stat: l40b.State = Stat
  Case 43:l41.State = Stat: l41b.State = Stat
  Case 44:l42.State = Stat: l42b.State = Stat
  Case 45:l43.State = Stat: l43b.State = Stat
  Case 46:l44.State = Stat: l44b.State = Stat
  Case 47:l45.State = Stat: l45b.State = Stat
  Case 48:l46.State = Stat: l46b.State = Stat
 End Select
 End Sub

 '*********************
 'LED's based on Eala's
 '*********************
 Dim Digits (36)
 Digits (0) = Array (a00, a02, a05, a06, a04, a01, a03, a07)
 Digits (1) = Array (a10, a12, a15, a16, a14, a11, a13)
 Digits (2) = Array (a20, a22, a25, a26, a24, a21, a23)
 Digits (3) = Array (a30, a32, a35, a36, a34, a31, a33, a37)
 Digits (4) = Array (a40, a42, a45, a46, a44, a41, a43)
 Digits (5) = Array (a50, a52, a55, a56, a54, a51, a53)
 Digits (6) = Array (a60, a62, a65, a66, a64, a61, a63)

 Digits (7) = Array (b00, b02, b05, b06, b04, b01, b03, b07)
 Digits (8) = Array (b10, b12, b15, b16, b14, b11, b13)
 Digits (9) = Array (b20, b22, b25, b26, b24, b21, b23)
 Digits (10) = Array (b30, b32, b35, b36, b34, b31, b33, b37)
 Digits (11) = Array (b40, b42, b45, b46, b44, b41, b43)
 Digits (12) = Array (b50, b52, b55, b56, b54, b51, b53)
 Digits (13) = Array (b60, b62, b65, b66, b64, b61, b63)

 Digits (14) = Array (e00, e02, e05, e06, e04, e01, e03, e07)
 Digits (15) = Array (e10, e12, e15, e16, e14, e11, e13)
 Digits (16) = Array (e20, e22, e25, e26, e24, e21, e23)
 Digits (17) = Array (e30, e32, e35, e36, e34, e31, e33, e37)
 Digits (18) = Array (e40, e42, e45, e46, e44, e41, e43)
 Digits (19) = Array (e50, e52, e55, e56, e54, e51, e53)
 Digits (20) = Array (e60, e62, e65, e66, e64, e61, e63)

 Digits (21) = Array (f00, f02, f05, f06, f04, f01, f03, f07)
 Digits (22) = Array (f10, f12, f15, f16, f14, f11, f13)
 Digits (23) = Array (f20, f22, f25, f26, f24, f21, f23)
 Digits (24) = Array (f30, f32, f35, f36, f34, f31, f33, f37)
 Digits (25) = Array (f40, f42, f45, f46, f44, f41, f43)
 Digits (26) = Array (f50, f52, f55, f56, f54, f51, f53)
 Digits (27) = Array (f60, f62, f65, f66, f64, f61, f63)

 Digits (28) = Array (c00, c02, c05, c06, c04, c01, c03)
 Digits (29) = Array (c10, c12, c15, c16, c14, c11, c13)

 Digits (30) = Array (d00, d02, d05, d06, d04, d01, d03)
 Digits (31) = Array (d10, d12, d15, d16, d14, d11, d13)

 Digits (34) = Array (g00, g02, g05, g06, g04, g01, g03)
 Digits (35) = Array (g10, g12, g15, g16, g14, g11, g13)

 Digits (32) = Array (h00, h02, h05, h06, h04, h01, h03)
 Digits (33) = Array (h10, h12, h15, h16, h14, h11, h13)

 '********************
 'Update LED's display
 '********************

 Sub LedsTimer_Timer()
     Dim ChgLED, ii, num, chg, stat, obj
     ChgLED = Controller.ChangedLEDs (&Hffffffff, &Hffffffff)
     If Not IsEmpty (ChgLED) Then
         For ii = 0 To UBound (chgLED)
             num = chgLED (ii, 0) : chg = chgLED (ii, 1) : stat = chgLED (ii, 2)
             For Each obj In Digits (num)
                 If chg And 1 Then obj.State = stat And 1
                 chg = chg \ 2 : stat = stat \ 2
             Next
         Next
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
Const maxvel = 36 'max ball velocity
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

'Inder Lap by Lap
'by Inkochnito
Sub editDips
  Dim vpmDips:Set vpmDips=New cvpmDips
  With vpmDips
    .AddForm 700,400,"Lap by Lap - DIP switches"
    .AddFrame 2,0,392,"Credits per coin",&H03000000,Array("1 coin - 1 credit && 1 coin - 4 credits",0,"2 coins - 1 credit (4 coins - 3 credits) && 1 coin - 3 credits",&H03000000)'SL1-G&H (dip 26&25)
    .AddFrame 2,46,190,"Replay threshold",&H30000000,Array("2,300,000 points",0,"2,500,000 points",&H10000000,"2,700,000 points",&H20000000,"2,900,000 points",&H30000000)'SL1-C&D (dip 30&29)
    .AddFrame 2,120,190,"Handicap value",&H00000300,Array("4,000,000 points",0,"4,200,000 points",&H00000100,"4,400,000 points",&H00000200,"4,600,000 points",&H00000300)'SL2-W&X(dip 10&9)
    .AddFrame 2,194,190,"Number of rounds handicap",&H00030000,Array("25",0,"30",&H00010000,"35",&H00020000,"40",&H00030000)'SL2-O&P(dip 18&17)
    .AddFrame 2,268,190,"Extra ball starting sequence",&H000C0000,Array("lights with 10,000",0,"lights with 25,000",&H00040000,"lights with 50,000",&H00080000)'SL2-M&N (dip 20&19)
    .AddFrame 205,46,190,"Balls per game",&H08000000,Array("3 balls",0,"5 balls",&H08000000)'SL1-E (dip 28)
    .AddFrame 205,92,190,"Extra ball when all side targets hit",&H00100000,Array("no",0,"yes",&H00100000)'SL2-L (dip 21)
    .AddFrame 205,138,190,"Side specials",&H00200000,Array("hitting side targets",0,"hitting all side targets",&H00200000)'SL2-K (dip 22)
    .AddFrame 205,184,190,"Targets reset",32768,Array("easy",0,"difficult",32768)'SL3-Q (dip 16)
    .AddFrame 205,230,190,"Ramp extra ball",&H00004000,Array("hitting side targets",0,"hitting all side targets",&H00004000)'SL3-R (dip 15)
    .AddLabel 50,340,300,20,"After hitting OK, press F3 to reset game with new settings."
    .ViewDips
  End With
End Sub
Set vpmShowDips=GetRef("editDips")

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


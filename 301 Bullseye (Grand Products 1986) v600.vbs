'301/Bulls Eye / IPD No. 403 / May, 1986 / 4 Players /  Grand Products Incorporated
'VPX8 table by jpsalas. Version 6.0.0

Option Explicit
Randomize

Const BallSize = 50
Const BallMass = 1

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Dim bsTrough, bsLSaucer,bsMSaucer,bsRSaucer, x

Const cGameName = "bullseye"

Dim VarHidden
If Table1.ShowDT = true then
For each x in aReels
x.visible = 1
Next
    VarHidden = 1
Else
For each x in aReels
x.visible = 0
Next
    VarHidden = 0
End If

if B2SOn = true then VarHidden = 1

LoadVPM "01500000", "BALLY.VBS", 3.2

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
        .SplashInfoLine = "301 Bullseye - Grand Products 1986" & vbNewLine & "VPX table by JPSalas v6.0.0"
    .Games(cGameName).Settings.Value("sound")=1
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = VarHidden
        .Games(cGameName).Settings.Value("rol") = 0
        '.SetDisplayPosition 0,0,GetPlayerHWnd 'uncomment if you can't see the dmd
        On Error Resume Next
        Controller.SolMask(0) = 0
        vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the Timer to renable all the solenoids after 2 seconds
        Controller.Run GetPlayerHWnd
        On Error Goto 0
    End With

    ' Nudging
    vpmNudge.TiltSwitch = 7
    vpmNudge.Sensitivity = 3
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, LeftSlingshot, RightSlingshot)

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0,33,0,0,0,0,0,0
        .InitKick BallRelease, 90, 4
        .InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .Balls = 1
    End With

    ' Saucers
    Set bsRSaucer = New cvpmBallStack
    With bsRSaucer
        .InitSaucer RightHole,1,205,18
        .InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .KickForceVar = 2
        .KickAngleVar = 2
    End With

    Set bsMSaucer = New cvpmBallStack
    With bsMSaucer
        .InitSaucer TopHole,2,195,14
        .InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .KickForceVar = 2
        .KickAngleVar = 2
    End With

    Set bsLSaucer = New cvpmBallStack
    With bsLSaucer
        .InitSaucer LeftHole,3,155,18
        .InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .KickForceVar = 2
        .KickAngleVar = 2
    End With

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1
    RealTime.Enabled = 1

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
    If keycode = CenterTiltKey Then Nudge 0, 8:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If vpmKeyDown(keycode)Then Exit Sub
    If keycode = PlungerKey Then PlaySoundat "fx_PlungerPull", Plunger:Plunger.Pullback
    If keycode = KeyRules Then Rules
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If vpmKeyUp(keycode)Then Exit Sub
    If keycode = PlungerKey Then PlaySoundAt "fx_plunger", Plunger:Plunger.Fire
End Sub

'*************************
' GI - needs new vpinmame
'*************************

Set GICallback = GetRef("GIUpdate")

Sub GIUpdate(no, Enabled)
    If Enabled Then
        GiOn
    Else
        GiOff
    End If
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
    vpmTimer.PulseSw 37
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
    vpmTimer.PulseSw 36
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
Sub SW19a_Hit:vpmTimer.PulseSw 19:End Sub
Sub SW19b_Hit:vpmTimer.PulseSw 19:End Sub
Sub SW19c_Hit:vpmTimer.PulseSw 19:End Sub
Sub SW19d_Hit:vpmTimer.PulseSw 19:End Sub

' Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 38:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper1:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 39:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper2:End Sub

' Drain & holes
Sub Drain_Hit:PlaysoundAt "fx_drain", Drain:bsTrough.AddBall Me:End Sub
Sub LeftHole_Hit:PlaysoundAt "fx_kicker_enter", LeftHole:bsLSaucer.AddBall 0:End Sub
Sub TopHole_Hit:PlaysoundAt "fx_kicker_enter", TopHole:bsMSaucer.AddBall 0:End Sub
Sub RightHole_Hit:PlaysoundAt "fx_kicker_enter", RightHole:bsRSaucer.AddBall 0:End Sub

' Rollovers
Sub sw35_Hit:Controller.Switch(35) = 1:PlaySoundAt "fx_sensor", sw35:End Sub
Sub sw35_UnHit:Controller.Switch(35) = 0:End Sub

Sub sw26_Hit:Controller.Switch(26) = 1:PlaySoundAt "fx_sensor", sw26:End Sub
Sub sw26_UnHit:Controller.Switch(26) = 0:End Sub

Sub sw25_Hit:Controller.Switch(25) = 1:PlaySoundAt "fx_sensor", sw25:End Sub
Sub sw25_UnHit:Controller.Switch(25) = 0:End Sub

Sub sw18_Hit:Controller.Switch(18) = 1:PlaySoundAt "fx_sensor", sw18:End Sub
Sub sw18_UnHit:Controller.Switch(18) = 0:End Sub

Sub sw17_Hit:Controller.Switch(17) = 1:PlaySoundAt "fx_sensor", sw17:End Sub
Sub sw17_UnHit:Controller.Switch(17) = 0:End Sub

Sub sw15_Hit:Controller.Switch(15) = 1:PlaySoundAt "fx_sensor", sw15:End Sub
Sub sw15_UnHit:Controller.Switch(15) = 0:End Sub

Sub sw8_Hit:Controller.Switch(8) = 1:PlaySoundAt "fx_sensor", sw8:End Sub
Sub sw8_UnHit:Controller.Switch(8) = 0:End Sub

Sub sw5_Hit:Controller.Switch(5) = 1:PlaySoundAt "fx_sensor", sw5:End Sub
Sub sw5_UnHit:Controller.Switch(5) = 0:End Sub

Sub sw4_Hit:Controller.Switch(4) = 1:PlaySoundAt "fx_sensor", sw4:End Sub
Sub sw4_UnHit:Controller.Switch(4) = 0:End Sub

'Spinners

Sub spinner001_Spin():vpmTimer.PulseSw 40:PlaySoundAt "fx_spinner", Spinner001:End Sub

'Targets
Sub t32_Hit:vpmTimer.PulseSw 32:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub t31_Hit:vpmTimer.PulseSw 31:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub t30_Hit:vpmTimer.PulseSw 30:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub t29_Hit:vpmTimer.PulseSw 29:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub t28_Hit:vpmTimer.PulseSw 28:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub t20_Hit:vpmTimer.PulseSw 20:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub t21_Hit:vpmTimer.PulseSw 21:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub t22_Hit:vpmTimer.PulseSw 22:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub t23_Hit:vpmTimer.PulseSw 23:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub t24_Hit:vpmTimer.PulseSw 24:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

'*********
'Solenoids
'*********
'SolCallback(14)="vpmSolSound ""sling"","
'SolCallback(15)="vpmSolSound""sling"","
'SolCallback(1)="vpmSolSound ""jet3"","
'SolCallback(2)="vpmSolSound ""jet3"","
SolCallBack(8)="bsLSaucer.SolOut"
SolCallBack(10)="bsMSaucer.SolOut"
SolCallBack(9)="bsRSaucer.SolOut"

SolCallback(7)="bsTrough.SolOut"
SolCallback(6)="vpmSolSound SoundFX(""fx_Knocker"",DOFKnocker),"

SolCallback(19) = "vpmNudge.SolGameOn"

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

Sub LeftFlipper_Animate:LeftFlipperTop.RotZ = LeftFlipper.CurrentAngle: End Sub
Sub RightFlipper_Animate: RightFlipperTop.RotZ = RightFlipper.CurrentAngle: End Sub

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
    Lamp 1, l1    '4 saucer
    Lamp 2, l2    '15 saucer
    Lamp 3, l3    '7 saucer    (was 16 saucer)
    Lamp 4, l4    '16 saucer     (was 8 saucer)
    Lamp 5, l5    '1 game
    Lamp 6, l6    '5 games
    Lamp 7, l7    '9 games
    Lamp 8, l8    'Extra ball left
    Lamp 9, l9    '2X loop     (was 3X loop)
    Lamp 10, l10  'D
    Lamp 11, l11  'Throw again
    Lamp 12, l12  '1 target
    Text 13, l13,   "BALL IN PLAY"    '(was Bumpers)     Not on actual machine
    'Lamp 14     Not used
    'Lamp 15     Not used
    'Lamp 16     Not connected
    Lamp 17, l17  '6 saucer
    Lamp 18, l18  '5 target
    Lamp 19, l19  '8 saucer    (was 11 saucer)
    Lampm 20, l20a  '50 points (2X)    (was only 50 board)
    Lamp 20, l20
    Lamp 21, l21  '2 games
    Lamp 22, l22  '6 games
    Lamp 23, l23  '10 games
    Lamp 24, l24  'Extra ball right
    Lamp 25, l25  '3X loop     (was 2X loop)
    Lamp 26, l26  'A
    TEXT 27, l27,   "MATCH"           'Not on actual machine
    Lamp 28, l28  '18 target
    Text 29, l29, "HIGH SCORE"          'Not on actual machine
    'Lamp 30     Not used
    'Lamp 31     Not used
    'Lamp 32     Not connected
    Lamp 33, l33  '10 saucer
    Lamp 34, l34  '9 target
    Lamp 35, l35  '11 saucer     (was 14 saucer)
    Lamp 36, l36  '3 target
    Lamp 37, l37  '3 games
    Lamp 38, l38  '7 games
    'Lamp 39, l39 'Not used    (was 50 saucer)
    Lamp 40, l40  'Special left
    Lamp 41, l41  '2X active     (was 3X active)
    Lamp 42, l42  'R
    Lamp 43, l43  '2 target
    Lamp 44, l44  '20 target
    TEXT 45, l45, "GAME OVER"
    'Lamp 46     Not used
    'Lamp 47     Not used
    'Lamp 48     Not connected
    Lamp 49, l49  '13 saucer
    Lamp 50, l50  '12 target
    Lamp 51, l51  '14 saucer     (was 7 saucer)
    Lamp 52, l52  '19 target
    Lamp 53, l53  '4 games
    Lamp 54, l54  '8 games
    Lamp 55, l55  'Special center    (was unknown)
    Lamp 56, l56  'Special right
    Lamp 57, l57  '3X active     (was 2X active)
    Lamp 58, l58  'T
    Lamp 59, l59  '17 target
    Lamp 60, l60  'Win
    TEXT 61, l61,   "TILT"
    'Lamp 62     Not used
    'Lamp 63     Not used
    'Lamp 64     Not connected
    Lamp 65, l65  '1 board
    Lamp 66, l66  '5 board
    Lamp 67, l67  '9 board
    Lamp 68, l68  '13 board
    Lamp 69, l69  '17 board
    Lampm 70, l70a  'GI right top          Might be inversed on actual machine
    Lamp 70, l70
    Lampm 71, l71a  'GI right bottom
    Lamp 71, l71
    'Lamp 72     Not connected
    Lamp 81, l81  '2 board
    Lamp 82, l82  '6 board
    Lamp 83, l83  '10 board
    Lamp 84, l84  '14 board
    Lamp 85, l85  '18 board
    Lampm 86, l86a  'GI top left           Might be inversed on actual machine
    Lamp 86, l86
    Lampm 87, l87a  'GI top right
    Lamp 87, l87
    'Lamp 88     Not connected
    Lamp 97, l97  '3 board
    Lamp 98, l98  '7 board
    Lamp 99, l99  '11 board
    Lamp 100, l100  '15 board
    Lamp 101, l101  '19 board
    Lampm 102, l102a    'GI left top           Might be inversed on actual machine
    Lamp 102, l102
    Lampm 103, l103a    'GI left bottom
    Lamp 103, l103
    'Lamp 104    Not connected
    Lamp 113, l113  '4 board
    Lamp 114, l114  '8 board
    Lamp 115, l115  '12 board
    Lamp 116, l116  '16 board
    Lamp 117, l117  '20 board
    'Lamp 118, l118  Not used
    'Lamp 119, l119  Not used
    'Lamp 120    Not connected
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

Sub Lampim(nr, object) ' used for multiple lights, it doesn't change the fading state
    Select Case FadingStep(nr)
        Case 0:object.state = 1
        Case 1:object.state = 0
    End Select
End Sub

Sub Lampi(nr, object)
    Select Case FadingStep(nr)
        Case 0:object.state = 1:FadingStep(nr) = -1
        Case 1:object.state = 0:FadingStep(nr) = -1
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

Sub Texti(nr, object, message)
    Select Case FadingStep(nr)
        Case 0:object.Text = message:FadingStep(nr) = -1
        Case 1:object.Text = "":FadingStep(nr) = -1
    End Select
End Sub

Sub Textim(nr, object, message)
    Select Case FadingStep(nr)
        Case 0:object.Text = message
        Case 1:object.Text = ""
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

Sub Gi(nr, object)
    Select Case FadingStep(nr)
        Case 1:GiOn:FadingStep(nr) = -1
        Case 0:GiOff:FadingStep(nr) = -1
    End Select
End Sub

'************************************
'          LEDs Display
'     Based on Scapino's LEDs
'************************************

Dim Digits(32)

    'Assign 7-digit output to reels
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
    Set Digits(26) = e2
    Set Digits(27) = e3

Sub UpdateLEDs
    On Error Resume Next
    Dim ChgLED, ii, jj, chg, stat
    ChgLED = Controller.ChangedLEDs(&H0000003f, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
        For ii = 0 To UBound(ChgLED)
            chg = chgLED(ii, 1):stat = chgLED(ii, 2)
            Select Case stat
                Case 0:Digits(chgLED(ii, 0)).SetValue 0    'empty
                Case 63:Digits(chgLED(ii, 0)).SetValue 1   '0
                Case 6:Digits(chgLED(ii, 0)).SetValue 2    '1
                Case 91:Digits(chgLED(ii, 0)).SetValue 3   '2
                Case 79:Digits(chgLED(ii, 0)).SetValue 4   '3
                Case 102:Digits(chgLED(ii, 0)).SetValue 5  '4
                Case 109:Digits(chgLED(ii, 0)).SetValue 6  '5
                Case 124:Digits(chgLED(ii, 0)).SetValue 7  '6
                Case 125:Digits(chgLED(ii, 0)).SetValue 7  '6
                Case 252:Digits(chgLED(ii, 0)).SetValue 7  '6
                Case 7:Digits(chgLED(ii, 0)).SetValue 8    '7
                Case 127:Digits(chgLED(ii, 0)).SetValue 9  '8
                Case 103:Digits(chgLED(ii, 0)).SetValue 10 '9
                Case 111:Digits(chgLED(ii, 0)).SetValue 10 '9
                Case 231:Digits(chgLED(ii, 0)).SetValue 10 '9
                Case 128:Digits(chgLED(ii, 0)).SetValue 0  'empty
                Case 191:Digits(chgLED(ii, 0)).SetValue 1  '0
                Case 832:Digits(chgLED(ii, 0)).SetValue 2  '1
                Case 896:Digits(chgLED(ii, 0)).SetValue 2  '1
                Case 768:Digits(chgLED(ii, 0)).SetValue 2  '1
                Case 134:Digits(chgLED(ii, 0)).SetValue 2  '1
                Case 219:Digits(chgLED(ii, 0)).SetValue 3  '2
                Case 207:Digits(chgLED(ii, 0)).SetValue 4  '3
                Case 230:Digits(chgLED(ii, 0)).SetValue 5  '4
                Case 237:Digits(chgLED(ii, 0)).SetValue 6  '5
                Case 253:Digits(chgLED(ii, 0)).SetValue 7  '6
                Case 135:Digits(chgLED(ii, 0)).SetValue 8  '7
                Case 255:Digits(chgLED(ii, 0)).SetValue 9  '8
                Case 239:Digits(chgLED(ii, 0)).SetValue 10 '9
            End Select
        Next
    End If
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
Sub aTargets_Hit(idx):PlaySound SoundFX("fx_target",DOFTargets), 0, Vol(ActiveBall), pan(ActiveBall), 0.2, Pitch(ActiveBall) * 10, 0, 0, AudioFade(ActiveBall):End Sub
Sub aRollovers_Hit(idx):PlaySoundAt "fx_sensor", aRollovers(idx):End Sub

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
Const maxvel = 30 'max ball velocity
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

 'Grand Products 301 Bullseye
 'added by Inkochnito
 'DIP switch 7 inversed according to electronic board schematic
 'Corrected '301 games needed to light extra ball' settings for 8 and 9 games
 Sub editDips
    Dim vpmDips:Set vpmDips=New cvpmDips
    With vpmDips
      .AddForm 700,400,"301 Bullseye - DIP switches"
      .AddChk 7,10,120,Array("Match feature",&H02000000)'dip 26
    .AddChk 205,10,120,Array("Credits displayed",&H01000000)'dip 25
      .AddFrame 2,30,190,"Maximum credits",&H00030000,Array("10 credits",0,"20 credits",&H00020000,"30 credits",&H00010000,"40 credits",&H00030000)'dip 17&18
      .AddFrame 2,106,190,"301 scoring adjustment",&H00000020,Array("score 301 to 0",0,"score 0 to 301",&H00000020)'dip 6
      .AddFrame 2,152,190,"Out/return lanes adjustment",&H00000040,Array("alternating lights",&H00000040,"both lights on",0)'dip 7
      .AddFrame 2,228,190,"301 games needed to light extra ball",&HE0000000,Array ("2",0,"3",&H20000000,"4",&H40000000,"5",&H60000000,"6",&H80000000,"7",&HA0000000,"8",&HC0000000,"9",&HE0000000)'dip 30&31&32
      .AddFrame 205,30,190,"High game to date",&H00006000,Array("no award",0,"1 credit",&H00002000,"2 credits",&H00004000,"3 credits",&H00006000)'dip 14&15
      .AddFrame 205,106,190,"Balls per game",32768,Array("3 balls",0,"5 balls",32768)'dip 16
      .AddFrame 205,152,190,"High score award",&H000C0000,Array("no award",&H000C0000,"novelty",&H00080000,"extra ball",0,"replay",&H00040000)'dip 19&20
      .AddFrame 205,228,190,"301 games needed to light special",&H1C000000,Array ("5",0,"6",&H04000000,"7",&H08000000,"8",&H0C000000,"9",&H10000000,"10",&H14000000,"11",&H18000000,"12",&H1C000000)'dip 27&28&29
      .AddLabel 50,370,300,20,"After hitting OK, press F3 to reset game with new settings."
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


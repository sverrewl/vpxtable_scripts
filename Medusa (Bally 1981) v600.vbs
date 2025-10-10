' Medusa - Bally 1981
' http://www.ipdb.org/machine.cgi?id=1565
' Medusa / IPD No. 1565 / February 04, 1981 / 4 Players
' VPX8 version by JPSalas 2024, version 6.0.0
' Light numbers from the tables by Joe Entropy & RipleYYY and Pacdude.
' Uses the Right Magna saves key to activate the save post (shield)

Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const BallSize = 50
Const BallMass = 1

LoadVPM "01550000", "Bally.vbs", 3.26

Dim bsTrough, bsSaucer, dtRBank, dtTBank
Dim x

Const cGameName = "medusa"

Const UseSolenoids = 0
Const UseLamps = 0
Const UseGI = 0
Const UseSync = 0
Const HandleMech = 0

Dim VarHidden
If Table1.ShowDT = true then
    VarHidden = 1
else
    VarHidden = 0
    For each x in aReels
        x.Visible = 0
    Next
end if

if B2SOn = true then VarHidden = 1

Set MotorCallback = GetRef("UpdateSolenoids")

' Standard Sounds
Const SSolenoidOn = "fx_solenoidon"
Const SSolenoidOff = "fx_solenoidoff"
Const SCoin = "fx_coin"

'************
' Table init.
'************

Sub table1_Init
    vpmInit me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Medusa - Bally 1981" & vbNewLine & "VPX table by JPSalas v6.0.0"
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
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
        On Error Goto 0
    End With

    ' Nudging
    vpmNudge.TiltSwitch = 15
    vpmNudge.Sensitivity = 1
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, Bumper4, LeftSlingshot, RightSlingshot, sw34)

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0, 8, 0, 0, 0, 0, 0, 0
        .InitKick BallRelease, 80, 6
        .InitAddSnd "fx_drain"
        .InitEntrySnd "fx_Solenoid", "fx_Solenoid"
        .InitExitSnd "fx_ballrel", "fx_Solenoid"
        .Balls = 1
        .IsTrough = True
    End With

    ' Saucer
    Set bsSaucer = New cvpmBallStack
    With bsSaucer
        .InitSaucer sw19, 19, 155, 14
        .InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .KickAngleVar = 2
        .KickForceVar = 2
    End With

    ' Right drop targets
    Set dtRBank = New cvpmDropTarget
    With dtRbank
        .initdrop Array(sw1, sw2, sw3, sw4), Array(1, 2, 3, 4)
        .initsnd SoundFX("fx_droptarget", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
        .createEvents "dtRbank"
    End With

    'Top drop targets
    Set dtTBank = New cvpmDropTarget
    With dtTbank
        .initdrop Array(sw48, sw47, sw46, sw45, sw44, sw43, sw42), Array(48, 47, 46, 45, 44, 43, 42)
        .initsnd SoundFX("fx_droptarget", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
        .createEvents "dtTbank"
    End With

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1
    RealTime.Enabled = 1

    ' init shield
    post2.IsDropped = 1
    post2rubber.visible = 0
    Post.Pullback
End Sub

Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub

'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
    If keycode = RightMagnaSave Then Controller.Switch(17) = 1
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If vpmKeyDown(keycode)Then Exit Sub
    If keycode = PlungerKey Then PlaySoundAt "fx_PlungerPull", Plunger:Plunger.Pullback
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If keycode = RightMagnaSave Then Controller.Switch(17) = 0
    If vpmKeyUp(keycode)Then Exit Sub
    If keycode = PlungerKey Then PlaySoundAt "fx_plunger",Plunger:Plunger.Fire
End Sub

'*********
' Switches
'*********

' Slings
Dim LStep, RStep, R2Step

Sub LeftSlingShot_Slingshot
    PlaySoundAt SoundFX("fx_slingshot", DOFContactors), Lemk
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 36
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
    vpmTimer.PulseSw 35
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

Sub sw34_Slingshot
    PlaySoundAt SoundFX("fx_slingshot", DOFContactors), Remk1
    R2Sling4.Visible = 1
    Remk1.RotX = 26
    R2Step = 0
    vpmTimer.PulseSw 34
    sw34.TimerEnabled = 1
End Sub

Sub sw34_Timer
    Select Case R2Step
        Case 1:R2SLing4.Visible = 0:R2SLing3.Visible = 1:Remk1.RotX = 14
        Case 2:R2SLing3.Visible = 0:R2SLing2.Visible = 1:Remk1.RotX = 2
        Case 3:R2SLing2.Visible = 0:Remk1.RotX = -10:sw34.TimerEnabled = 0
    End Select
    R2Step = R2Step + 1
End Sub

' Rubbers
Sub sw18_Hit:vpmTimer.PulseSw 18:End Sub
Sub sw18a_Hit:vpmTimer.PulseSw 18:End Sub
Sub sw18b_Hit:vpmTimer.PulseSw 18:End Sub

' Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 37:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper1:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 38:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper2:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 39:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper3:End Sub
Sub Bumper4_Hit:vpmTimer.PulseSw 40:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper4:End Sub

' Drain holes
Sub Drain_Hit:bsTrough.AddBall Me:End Sub

'Saucer
Sub sw19_Hit:bsSaucer.AddBall 0:End Sub

' Rollovers
Sub sw32_Hit:Controller.Switch(32) = true:PlaySoundAt "fx_sensor",sw32:End Sub
Sub sw32_UnHit:Controller.Switch(32) = false:End Sub
Sub sw32a_Hit:Controller.Switch(32) = true:PlaySoundAt "fx_sensor",sw32a:End Sub
Sub sw32a_UnHit:Controller.Switch(32) = false:End Sub
Sub sw20_Hit():controller.switch(20) = true:PlaySoundAt "fx_sensor",sw20:End Sub
Sub sw20_unHit():controller.switch(20) = false:End Sub
Sub sw21_Hit():controller.switch(21) = true:PlaySoundAt "fx_sensor",sw21:End Sub
Sub sw21_unHit():controller.switch(21) = false:End Sub
Sub sw22_Hit():controller.switch(22) = true:PlaySoundAt "fx_sensor",sw22:End Sub
Sub sw22_unHit():controller.switch(22) = false:End Sub
Sub sw23_Hit():controller.switch(23) = true:PlaySoundAt "fx_sensor",sw23:End Sub
Sub sw23_unHit():controller.switch(23) = false:End Sub
Sub sw24_Hit():controller.switch(24) = true:PlaySoundAt "fx_sensor",sw24:End Sub
Sub sw24_unHit():controller.switch(24) = false:End Sub
Sub sw25_Hit():controller.switch(25) = true:PlaySoundAt "fx_sensor",sw25:End Sub
Sub sw25_unHit():controller.switch(25) = false:End Sub
Sub sw25a_Hit():controller.switch(25) = true:PlaySoundAt "fx_sensor",sw25a:End Sub
Sub sw25a_unHit():controller.switch(25) = false:End Sub
Sub sw26_Hit():controller.switch(26) = true:PlaySoundAt "fx_sensor",sw26:End Sub
Sub sw26_unHit():controller.switch(26) = false:End Sub
Sub sw28_Hit():controller.switch(28) = true:PlaySoundAt "fx_sensor",sw28:End Sub
Sub sw28_unHit():controller.switch(28) = false:End Sub
Sub sw31_Hit():SetLamp 150, 1:controller.switch(31) = true:PlaySoundAt "fx_sensor",sw31:End Sub
Sub sw31_unHit():SetLamp 150, 0:controller.switch(31) = false:End Sub
Sub sw31a_Hit():SetLamp 151, 1:controller.switch(31) = true:PlaySoundAt "fx_sensor",sw31a:End Sub
Sub sw31a_unHit():SetLamp 151, 0:controller.switch(31) = false:End Sub

' Targets
Sub sw12_Hit:vpmTimer.PulseSw 12:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw13_Hit:vpmTimer.PulseSw 13:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw14_Hit:vpmTimer.PulseSw 14:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw29_Hit:vpmTimer.PulseSw 29:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw30_Hit:vpmTimer.PulseSw 30:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw27_Hit:vpmTimer.PulseSw 27:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

'Spinner
Sub sw33_Spin:vpmTimer.PulseSw 33:PlaySoundAt "spinner", sw33:End Sub

'*********
'Solenoids
'*********

Sub UpdateSolenoids
    Dim Changed, ii, solNo
    Changed = Controller.ChangedSolenoids
    If Not IsEmpty(Changed)Then
        For ii = 0 To UBound(Changed)
            solNo = Changed(ii, CHGNO)
            If Controller.Lamp(34)Then
                If SolNo = 1 Or SolNo = 2 Or SolNo = 3 Or SolNo = 4 Or Solno = 5 Or Solno = 6 Or Solno = 8 Then solNo = solNo + 24 '1->25 etc
            End If
            vpmDoSolCallback solNo, Changed(ii, CHGSTATE)
        Next
    End If
End Sub

SolCallback(1) = "vpmSolSound ""fx_knocker"","
SolCallback(2) = "dtRBank.SolDropUp"
SolCallback(3) = "dtTBank.SolDropUp"
SolCallback(7) = "SolShieldPost"
SolCallback(8) = "bsTrough.SolOut"
SolCallback(19) = "RelayAC"
SolCallback(25) = "SolZipOpen"
SolCallback(26) = "SolZipClose"
SolCallback(30) = "dtTBank.SolHit 1,"
SolCallback(6) = "dtTBank.SolHit 2,"
SolCallback(29) = "dtTBank.SolHit 3,"
SolCallback(5) = "dtTBank.SolHit 4,"
SolCallback(28) = "dtTBank.SolHit 5,"
SolCallback(4) = "dtTBank.SolHit 6,"
SolCallback(27) = "dtTBank.SolHit 7,"
SolCallback(32) = "bsSaucer.SolOut"

Sub SolZipOpen(enabled)
    PlaySoundAt "fx_SolenoidOn", LeftFlipper2
    LeftFlipper2.Visible = True:LeftFlipper2.Enabled = True
    RightFlipper2.Visible = True:RightFlipper2.Enabled = True
    LeftFlipper3.Visible = False:LeftFlipper3.Enabled = False
    RightFlipper3.Visible = False:RightFlipper3.Enabled = False
    Controller.Switch(41) = 1
End Sub

Sub SolZipClose(enabled)
    PlaySoundAt "fx_SolenoidOff", LeftFlipper2
    LeftFlipper2.Visible = False:LeftFlipper2.Enabled = False
    RightFlipper2.Visible = False:RightFlipper2.Enabled = False
    LeftFlipper3.Visible = True:LeftFlipper3.Enabled = True
    RightFlipper3.Visible = true:RightFlipper3.Enabled = true
    Controller.Switch(41) = 0
End Sub

Sub SolShieldPost(Enabled)
    If Enabled Then
        PlaySoundAt "fx_solenoidOn", Post
        post1.IsDropped = 1
        post1Rubber.Visible = 0
        post2.IsDropped = 0
        post2Rubber.Visible = 1
        Post.Fire
    Else
        PlaySoundAt "fx_solenoidOff", Post
        post1.IsDropped = 0
        post1Rubber.Visible = 1
        post2.IsDropped = 1
        post2Rubber.Visible = 0
        Post.PullBack
    End If
End Sub

Sub RelayAC(Enabled)
    vpmNudge.SolGameOn Enabled
    If Enabled Then
        GiOn
        SetLamp 152, 1
        SetLamp 153, 1
    Else
        GiOff
        SetLamp 152, 0
        SetLamp 153, 0
    End If
End Sub

'*******************
' Flipper Subs Rev3
'*******************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
    If Enabled Then
        SetLamp 152, 0
        PlaySoundAt SoundFX("fx_flipperup", DOFFlippers), LeftFlipper
        LeftFlipper.RotateToEnd
        LeftFlipper2.RotateToEnd
        LeftFlipper3.RotateToEnd
        LeftFlipperOn = 1
    Else
        SetLamp 152, 1
        PlaySoundAt SoundFX("fx_flipperdown", DOFFlippers), LeftFlipper
        LeftFlipper.RotateToStart
        LeftFlipper2.RotateToStart
        LeftFlipper3.RotateToStart
        LeftFlipperOn = 0
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        SetLamp 153, 0
        PlaySoundAt SoundFX("fx_flipperup", DOFFlippers), RightFlipper
        RightFlipper.RotateToEnd
        RightFlipper2.RotateToEnd
        RightFlipper3.RotateToEnd
        RightFlipperOn = 1
    Else
        SetLamp 153, 1
        PlaySoundAt SoundFX("fx_flipperdown", DOFFlippers), RightFlipper
        RightFlipper.RotateToStart
        RightFlipper2.RotateToStart
        RightFlipper3.RotateToStart
        RightFlipperOn = 0
    End If
End Sub

' flippers hit Sound

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, Vol(ActiveBall), pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, Vol(ActiveBall), pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub LeftFlipper2_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, Vol(ActiveBall), pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper2_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, Vol(ActiveBall), pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub LeftFlipper3_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, Vol(ActiveBall), pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper3_Collide(parm)
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

'*****************
'   Gi Effects
'*****************

Dim OldGiState
OldGiState = -1 'start witht he Gi off

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

Sub GiEffect
    For each x in aGiLights
        x.Duration 2, 2000, 1
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
    Lamp 1, l1
    Lamp 2, l2
    Lamp 3, l3
    Lampm 4, l4_1
    Lamp 4, l4
    Lamp 5, l5
    Lamp 6, l6
    Lamp 7, l7
    Lamp 8, l8
    Lamp 9, l9
    Lamp 10, l10

    Lamp 12, l12

    Lamp 14, l14
    Lamp 15, l15
    Lamp 17, l17
    Lamp 18, l18
    Lamp 19, l19
    Lamp 20, l20
    Lampm 21, l21_1
    Lamp 21, l21
    Lamp 22, l22
    Lamp 23, l23
    Lamp 24, l24
    Lamp 25, l25
    Lamp 26, l26

    Lamp 28, l28

    Lamp 30, l30
    Lamp 31, l31
    Lamp 33, l33
    Lamp 35, l35
    Lamp 36, l36
    Lampm 37, l37_1
    Lamp 37, l37
    Lamp 38, l38
    Lamp 39, l39
    Lamp 40, l40
    Lamp 41, l41
    Lamp 42, l42
    Lamp 43, l43
    Lamp 44, l44

    Lamp 46, l46
    Lampm 47, l47a
    Lampm 47, l47b
    Lampm 47, l47c
    Lampm 47, l47d
    Lampm 47, l47e
    Lamp 47, l47
    Lamp 49, l49
    Lamp 51, l51
    Lamp 52, l52
    Lampm 53, l53_1
    Lamp 53, l53
    Lamp 54, l54
    Lamp 55, l55
    Lamp 56, l56
    Lamp 57, l57
    Lamp 58, l58
    Lamp 59, l59

    Lamp 62, l62
    Lamp 63, l63
    Lampm 70, l70_1
    Lamp 70, l70
    Lampm 71, l71_1
    Lamp 71, l71
    Lampm 86, l86_1
    Lamp 86, l86
    Lamp 87, l87
    Lampm 102, l102_1
    Lamp 102, l102
    Lamp 103, l103
    Lampm 118, l118_1
    Lamp 118, l118
    Lamp 119, l119

    Lamp 65, l65
    Lamp 81, l81
    Lamp 97, l97
    Lamp 113, l113
    Lamp 66, l66
    Lamp 82, l82
    Lamp 98, l98
    Lamp 114, l114
    Lamp 67, l67
    Lamp 83, l83
    Lamp 99, l99
    Lamp 115, l115
    Lamp 68, l68
    Lamp 84, l84
    Lamp 100, l100
    Lamp 116, l116
    Lamp 69, l69
    Lamp 85, l85
    Lamp 101, l101
    Lamp 117, l117
    Lamp 60, l60

    Lamp 150, lleft
    Lamp 151, lright
    Lampm 152, flipperlight1
    Lampm 152, flipperlight2
    Lampm 152, flipperlight3
    Lamp 152, flipperlight4
    Lampm 153, flipperlight5
    Lampm 153, flipperlight6
    Lampm 153, flipperlight7
    Lamp 153, flipperlight8
    ' backdrop lights
        Text 11, l11, "SAME PLAYER SHOOTS AGAIN"
        Text 13, l13, "BALL IN PLAY"
        Text 27, l27, "MATCH"
        Text 29, l29, "HI SCORE TO DATE"
        Text 45, l45, "GAME OVER"
        Text 61, l61, "TILT"
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

'****************
' Real Time timer
'****************

Sub RealTime_Timer
    GIUpdate
    RollingUpdate
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
Const maxvel = 28 'max ball velocity
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

'************************************
'          LEDs Display
'     Based on Scapino's LEDs
'************************************

Dim Digits(38)
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

Digits(32) = Array(a00, a02, a05, a06, a04, a01, a03)
Digits(33) = Array(a10, a12, a15, a16, a14, a11, a13)
Digits(34) = Array(a20, a22, a25, a26, a24, a21, a23)
Digits(35) = Array(a30, a32, a35, a36, a34, a31, a33)
Digits(36) = Array(a40, a42, a45, a46, a44, a41, a43)
Digits(37) = Array(a50, a52, a55, a56, a54, a51, a53)

Sub UpdateLeds
    On Error Resume Next
    Dim ChgLED, ii, jj, chg, stat, num, obj
    ChgLED = Controller.ChangedLEDs(&H0000003f, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
        For ii = 0 To UBound(ChgLED)
            num = chgLED(ii, 0):chg = chgLED(ii, 1):stat = chgLED(ii, 2)
            If num < 32 Then
                For jj = 0 to 10
                    If stat = Patterns(jj)OR stat = Patterns2(jj)then Digits(chgLED(ii, 0)).SetValue jj
                Next
            Else
                For Each obj In Digits(num)
                    If chg And 1 Then obj.State = stat And 1
                    chg = chg \ 2:stat = stat \ 2
                Next
            End If
        Next
    End IF
End Sub

'Bally Medusa
'added by Inkochnito
Sub editDips
    Dim vpmDips:Set vpmDips = New cvpmDips
    With vpmDips
        .AddForm 700, 400, "Medusa - DIP switches"
        .AddChk 0, 5, 120, Array("Match feature", &H08000000)                                                                                                                                                                                                                 'dip 28
        .AddChk 0, 25, 120, Array("Credits displayed", &H04000000)                                                                                                                                                                                                            'dip 27
        .AddFrame 0, 44, 190, "Extra ball match number display adjust", &H00400000, Array("any number flashing will be reset", 0, "any number flashing will for next ball", &H00400000)                                                                                       'dip 23
        .AddFrame 0, 90, 190, "Top Olympus red lites", &H00000020, Array("step back 1 when target is hit", 0, "do not step back", &H00000020)                                                                                                                                 'dip 6
        .AddFrame 0, 136, 190, "Any advanced Colossus bonus lite on", &H00000040, Array("will be reset", 0, "will come back on for next ball", &H00000040)                                                                                                                    'dip 7
        .AddFrame 0, 182, 190, "Any lit left side 2 or 3 arrow", &H00004000, Array("will be reset", 0, "will come back on for next ball", &H00004000)                                                                                                                         'dip 15
        .AddFrame 0, 228, 190, "Medusa special lites with", 32768, Array("80K", 0, "40K and 80K", 32768)                                                                                                                                                                      'dip 16
        .AddFrame 0, 274, 190, "Collect Olympus bonus saucer lite", &H20000000, Array("will be reset", 0, "will come back on for next ball", &H20000000)                                                                                                                      'dip 30
        .AddFrame 0, 320, 395, "Movable flipper timer adjust", &H00000080, Array("closed flippers will open after 10 seconds until next target is hit", 0, "Hitting top targets adds 10 seconds each to keep flippers closed", &H00000080)                                    'dip 8
        .AddFrame 205, 0, 190, "Maximum credits", &H03000000, Array("10 credits", 0, "15 credits", &H01000000, "25 credits", &H02000000, "40 credits", &H03000000)                                                                                                            'dip 25&26
        .AddFrame 205, 76, 190, "Balls per game", &HC0000000, Array("2 balls", &HC0000000, "3 balls", 0, "4 balls", &H80000000, "5 balls", &H40000000)                                                                                                                        'dip 31&32
        .AddFrame 205, 152, 190, "Olympus bonus red lights", &H00100000, Array("1st 5-3, 2nd 5-2, 3dr 5-1, 4th 5-1", 0, "1st 5-3, 2nd 5-2, 3dr 5-2, 4th 5-2", &H00100000, "1st 5-3, 2nd 5-3, 3dr 5-2, 4th 5-2", &H00200000, "1st 5-3, 2nd 5-3, 3dr 5-3, 4th 5-3", &H00300000) 'dip 21&22
        .AddFrame 205, 228, 190, "Medusa bonus from 1 to 19 memory", &H00800000, Array("will be reset", 0, "will come back on for next ball", &H00800000)                                                                                                                     'dip 24
        .AddFrame 205, 274, 190, "Replay limit", &H10000000, Array("1 replay per game", 0, "unlimited replays", &H10000000)                                                                                                                                                   'dip 29
        .AddLabel 30, 390, 350, 15, "Set selftest position 16,17,18 and 19 to 03 for the best gameplay."
        .AddLabel 50, 410, 300, 15, "After hitting OK, press F3 to reset game with new settings."
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


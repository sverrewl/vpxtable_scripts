' Pokerino / IPD No. 1839 / November 28, 1978 / 4 Players
' VPX8 by JPSalas v6.0.0 December 2022

Option Explicit
Randomize

Const BallSize = 50
Const BallMass = 1

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01120100", "s4.vbs", 3.02

Dim bsTrough, dtJ, dtQ, dtS, bsLHole, x, FlipperEnabled

Const cGameName = "pkrno_l1"

Const UseSolenoids = 2
Const UseLamps = 1
Const UseGI = 0
Const UseSync = 0
Const HandleMech = 0

Dim VarHidden
If Table1.ShowDT = true Then
    VarHidden = 1
    For each x in aReels
        x.Visible = 1
    Next
else
    VarHidden = 0
    For each x in aReels
        x.Visible = 0
    Next
End If

If B2SOn = true Then VarHidden = 1

' Standard Sounds
Const SSolenoidOn = "fx_SolenoidOn"
Const SSolenoidOff = "fx_SolenoidOff"
'Const SFlipperOn = "fx_FlipperUp"
'Const SFlipperOff = "fx_FlipperDown"
Const SCoin = "fx_coin"

'************
' Table init.
'************

Sub Table1_Init
    vpmInit me

    With Controller
        .GameName = cGameName
        .SplashInfoLine = "Pokerino, Williams 1978" & vbNewLine & "VPX table by JPSalas v6.0.0"
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
    vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - Then add the timer to renable all the solenoids after 2 seconds
    Controller.Run

    ' Nudging
    vpmNudge.TiltSwitch = 1
    vpmNudge.Sensitivity = 1
    vpmNudge.TiltObj = Array(Bumper001, Bumper002, Bumper003, leftSlingshot, RightSlingshot)

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0, 29, 0, 0, 0, 0, 0, 0
        .InitKick BallRelease, 80, 6
        .InitEntrySnd "fx_Solenoid", "fx_Solenoid"
        .InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .Balls = 1
    End With

    ' J Drop targets
    set dtJ = new cvpmdroptarget
    With dtJ
        .InitDrop Array(sw34, sw33, sw26, sw25), Array(34, 33, 26, 25)
        .initsnd SoundFX("fx_droptarget", DOFContactors), SoundFX("fx_resetdrop", DOFContactors)
        .AllDownSw = 35
        .CreateEvents "dtJ"
    End With

    ' Queen Drop targets 12
    set dtQ = new cvpmdroptarget
    With dtQ
        .InitDrop Array(sw20, sw19, sw18, sw17), Array(20, 19, 18, 17)
        .initsnd SoundFX("fx_droptarget", DOFContactors), SoundFX("fx_resetdrop", DOFContactors)
        .AllDownSw = 21
        .CreateEvents "dtQ"
    End With

    ' Center drop target
    set dtS = new cvpmdroptarget
    With dtS
        .InitDrop Array(sw42), Array(42)
        .initsnd SoundFX("fx_droptarget", DOFContactors), SoundFX("fx_resetdrop", DOFContactors)
        .CreateEvents "dtS"
    End With

    vpmMapLights aLights

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1
    RealTime.Enabled = 1

    ' Turn on Gi
    vpmtimer.addtimer 1500, "GiOn '"
End Sub

'******************
' Realtime Updates
'******************

Sub RealTime_Timer
    GIUpdate
    RollingUpdate
End Sub

'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If vpmKeyDown(keycode)Then Exit Sub
    If keycode = PlungerKey Then PlaySoundat "fx_PlungerPull", Plunger:Plunger.Pullback
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If vpmKeyUp(keycode)Then Exit Sub
    If keycode = PlungerKey Then PlaySoundat "fx_plunger", Plunger:Plunger.Fire
End Sub

'*********
' Switches
'*********

' Slings & Rubbers
' Slings
Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    PlaySoundAt SoundFX("fx_slingshot", DOFContactors), Lemk
    LeftSling004.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 30
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 1:LeftSLing004.Visible = 0:LeftSLing003.Visible = 1:Lemk.RotX = 14
        Case 2:LeftSLing003.Visible = 0:LeftSLing002.Visible = 1:Lemk.RotX = 2
        Case 3:LeftSLing002.Visible = 0:Lemk.RotX = -10:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
    PlaySoundAt SoundFX("fx_slingshot", DOFContactors), Remk
    RightSling004.Visible = 1
    Remk.RotX = 26
    RStep = 0
    vpmTimer.PulseSw 28
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 1:RightSLing004.Visible = 0:RightSLing003.Visible = 1:Remk.RotX = 14
        Case 2:RightSLing003.Visible = 0:RightSLing002.Visible = 1:Remk.RotX = 2
        Case 3:RightSLing002.Visible = 0:Remk.RotX = -10:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

' scoring rubbers
Sub sw11_Hit:vpmTimer.PulseSw 11:End Sub
Sub sw37_Hit:vpmTimer.PulseSw 37:End Sub
Sub sw45_Hit:vpmTimer.PulseSw 45:End Sub
Sub sw32_Hit:vpmTimer.PulseSw 32:End Sub
Sub sw22_Hit:vpmTimer.PulseSw 22:End Sub
Sub sw23_Hit:vpmTimer.PulseSw 23:End Sub
Sub sw46_Hit:vpmTimer.PulseSw 46:End Sub

' Bumpers
Sub Bumper001_Hit:vpmTimer.PulseSw 18:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper001:End Sub
Sub Bumper002_Hit:vpmTimer.PulseSw 19:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper002:End Sub
Sub Bumper003_Hit:vpmTimer.PulseSw 20:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper003:End Sub

' Spinners

Sub sw36_Spin:vpmTimer.PulseSw 17:PlaySoundAt "fx_spinner", sw36:End Sub

' Drain & holes
Sub Drain_Hit:PlaySoundAt "fx_drain", drain:bsTrough.AddBall Me:End Sub

' Rollovers
' oulanes
Sub sw31_Hit:Controller.Switch(31) = 1:PlaySoundAt "fx_sensor", sw31:End Sub
Sub sw31_UnHit:Controller.Switch(31) = 0:End Sub

Sub sw27_Hit:Controller.Switch(27) = 1:PlaySoundAt "fx_sensor", sw27:End Sub
Sub sw27_UnHit:Controller.Switch(27) = 0:End Sub

'stars
Sub sw9_Hit:Controller.Switch(9) = 1:PlaySoundAt "fx_sensor", sw9:l1x.Duration 1, 250, 0:End Sub
Sub sw9_UnHit:Controller.Switch(9) = 0:End Sub

Sub sw10_Hit:Controller.Switch(10) = 1:PlaySoundAt "fx_sensor", sw10:l2x.Duration 1, 250, 0:End Sub
Sub sw10_UnHit:Controller.Switch(10) = 0:End Sub

Sub sw16_Hit:Controller.Switch(16) = 1:PlaySoundAt "fx_sensor", sw16:l7x.Duration 1, 250, 0:End Sub
Sub sw16_UnHit:Controller.Switch(16) = 0:End Sub

Sub sw41_Hit:Controller.Switch(41) = 1:PlaySoundAt "fx_sensor", sw41:l8x.Duration 1, 250, 0:End Sub
Sub sw41_UnHit:Controller.Switch(41) = 0:End Sub

' A
Sub sw12_Hit:Controller.Switch(12) = 1:PlaySoundAt "fx_sensor", sw12:l3x.Duration 1, 250, 0:End Sub
Sub sw12_UnHit:Controller.Switch(12) = 0:End Sub

Sub sw13_Hit:Controller.Switch(13) = 1:PlaySoundAt "fx_sensor", sw13:l4x.Duration 1, 250, 0:End Sub
Sub sw13_UnHit:Controller.Switch(13) = 0:End Sub

Sub sw14_Hit:Controller.Switch(14) = 1:PlaySoundAt "fx_sensor", sw14:l5x.Duration 1, 250, 0:End Sub
Sub sw14_UnHit:Controller.Switch(14) = 0:End Sub

Sub sw15_Hit:Controller.Switch(15) = 1:PlaySoundAt "fx_sensor", sw15:l6x.Duration 1, 250, 0:End Sub
Sub sw15_UnHit:Controller.Switch(15) = 0:End Sub

'*********
'Solenoids
'*********
SolCallback(1) = "bsTrough.SolOut"
SolCallback(2) = "SolResetS"
SolCallback(3) = "SolResetJ1"
SolCallback(4) = "SolResetJ2"
SolCallback(5) = "SolResetQ1"
SolCallback(6) = "SolResetQ2"
SolCallback(7) = "SolPostDown"
SolCallback(8) = "SolPostUp"

Solcallback(14) = "vpmsolsound SoundFX(""fx_knocker"",DOFKnocker),"
Solcallback(23) = "SolRun"

'**************
' Flipper Subs
'**************

SolCallback(sULFlipper) = "SolULFlipper"
SolCallback(sURFlipper) = "SolURFlipper"
SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
If StagedFlipperMod Then
    If Enabled Then
        PlaySoundAt SoundFX("fx_flipperup", DOFFlippers), LeftFlipper
        LeftFlipper.RotateToEnd
        LeftFlipperOn = 1
    Else
        PlaySoundAt SoundFX("fx_flipperdown", DOFFlippers), LeftFlipper
        LeftFlipper.RotateToStart
        LeftFlipperOn = 0
    End If
Else
    If Enabled Then
        PlaySoundAt SoundFX("fx_flipperup", DOFFlippers), LeftFlipper
        LeftFlipper.RotateToEnd
        LeftFlipper001.RotateToEnd
        LeftFlipperOn = 1
    Else
        PlaySoundAt SoundFX("fx_flipperdown", DOFFlippers), LeftFlipper
        LeftFlipper.RotateToStart
        LeftFlipper001.RotateToStart
        LeftFlipperOn = 0
    End If

End If
End Sub

Sub SolRFlipper(Enabled)
If StagedFlipperMod Then
    If Enabled Then
        PlaySoundAt SoundFX("fx_flipperup", DOFFlippers), RightFlipper
        RightFlipper.RotateToEnd
        RightFlipperOn = 1
    Else
        PlaySoundAt SoundFX("fx_flipperdown", DOFFlippers), RightFlipper
        RightFlipper.RotateToStart
        RightFlipperOn = 0
    End If
Else
    If Enabled Then
        PlaySoundAt SoundFX("fx_flipperup", DOFFlippers), RightFlipper
        RightFlipper.RotateToEnd
        RightFlipper001.RotateToEnd
        RightFlipperOn = 1
    Else
        PlaySoundAt SoundFX("fx_flipperdown", DOFFlippers), RightFlipper
        RightFlipper.RotateToStart
        RightFlipper001.RotateToStart
        RightFlipperOn = 0
    End If
End If
End Sub

Sub SolULFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("fx_flipperup", DOFFlippers), LeftFlipper001
        LeftFlipper001.RotateToEnd
    Else
        PlaySoundAt SoundFX("fx_flipperdown", DOFFlippers), LeftFlipper001
        LeftFlipper001.RotateToStart
    End If
End Sub

Sub SolURFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("fx_flipperup", DOFFlippers), RightFlipper001
        RightFlipper001.RotateToEnd
    Else
        PlaySoundAt SoundFX("fx_flipperdown", DOFFlippers), RightFlipper001
        RightFlipper001.RotateToStart
    End If
End Sub

' flippers top animations

Sub LeftFlipper_Animate: LeftFlipperTop.RotZ = LeftFlipper.CurrentAngle: End Sub
Sub RightFlipper_Animate: RightFlipperTop.RotZ = RightFlipper.CurrentAngle: End Sub
Sub LeftFlipper001_Animate: LeftFlipperTop001.RotZ = LeftFlipper001.CurrentAngle: End Sub
Sub RightFlipper001_Animate: RightFlipperTop001.RotZ = RightFlipper001.CurrentAngle: End Sub


Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, Vol(ActiveBall), pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, Vol(ActiveBall), pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub LeftFlipper001_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, Vol(ActiveBall), pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper001_Collide(parm)
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
'Start Of Stroke Flipper001 Stroke Routine : Start of Stroke for Tap pass and Tap shoot
    If LeftFlipper001.CurrentAngle >= LeftFlipper001.StartAngle - SOSAngle Then LeftFlipper001.Strength = FlipperPower * SOSTorque else LeftFlipper001.Strength = FlipperPower : End If

'End Of Stroke Routine : Livecatch and Emply/Full-Charged EOS
  If LeftFlipperOn = 1 Then
    If LeftFlipper001.CurrentAngle = LeftFlipper001.EndAngle then
      LeftFlipper001.EOSTorque = FullStrokeEOS_Torque
      If LLiveCatchTimer < LiveCatchSensivity Then
        LeftFlipper001.Elasticity = 0
      Else
        LeftFlipper001.Elasticity = FlipperElasticity
      End If
    End If
  Else
    LeftFlipper001.Elasticity = FlipperElasticity
    LeftFlipper001.EOSTorque = LiveStrokeEOS_Torque
    LLiveCatchTimer = 0
  End If

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
    If RightFlipper001.CurrentAngle <= RightFlipper001.StartAngle + SOSAngle Then RightFlipper001.Strength = FlipperPower * SOSTorque else RightFlipper001.Strength = FlipperPower : End If

'End Of Stroke Routine : Livecatch and Emply/Full-Charged EOS
  If RightFlipperOn = 1 Then
    If RightFlipper001.CurrentAngle = RightFlipper001.EndAngle Then
      RightFlipper001.EOSTorque = FullStrokeEOS_Torque
      If RLiveCatchTimer < LiveCatchSensivity Then
        RightFlipper001.Elasticity = 0
      Else
        RightFlipper001.Elasticity = FlipperElasticity
        RLiveCatchTimer = LiveCatchSensivity
      End If
    End If
  Else
    RightFlipper001.Elasticity = FlipperElasticity
    RightFlipper001.EOSTorque = LiveStrokeEOS_Torque
    RLiveCatchTimer = 0
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

'Solenoid Subs

Sub SolResetS(Enabled)
    If Enabled Then
        dtS.SolUnHit 1, True
    End If
End Sub

Sub SolResetJ1(Enabled)
    If Enabled Then
        Controller.Switch(35) = False
        dtJ.SolUnHit 1, True
        dtJ.SolUnHit 2, True
    End If
End Sub

Sub SolResetJ2(Enabled)
    If Enabled Then
        Controller.Switch(35) = False
        dtJ.SolUnHit 3, True
        dtJ.SolUnHit 4, True
    End If
End Sub

Sub SolResetQ1(Enabled)
    If Enabled Then
        Controller.Switch(21) = False
        dtQ.SolUnHit 1, True
        dtQ.SolUnHit 2, True
    End If
End Sub

Sub SolResetQ2(Enabled)
    If Enabled Then
        Controller.Switch(21) = False
        dtQ.SolUnHit 3, True
        dtQ.SolUnHit 4, True
    End If
End Sub

Sub SolPostDown(Enabled)
    If Enabled Then
        PlaySoundAt "fx_solenoidOn", KF4
        Post.IsDropped = True
        KC.disabled = True
        KH1.TimerInterval = 100
        KH1.TimerEnabled = True
    End If
End Sub

Sub SolPostUp(Enabled)
    If Enabled Then
        PlaySoundAt "fx_solenoidOff", kf4
        Post.IsDropped = False
        KH1.enabled = True
    End If
End Sub

Sub SolRun(Enabled)
    vpmNudge.SolGameOn Enabled
    FlipperEnabled = Enabled
End Sub

'********************************
'    K i n g s   C h a m b e r
'       from Eala's table
'********************************

const cdircf = 342
const cdirch = 225

dim KFb1, KFb2, KFb3, KFb4
dim KHb1, KHb2, KHb3, KHb4

Dim Ballspeed

 KF1.CreateBall:KFb1 = True
 KF2.CreateBall:KFb2 = True
 KF3.CreateBall:KFb3 = True
 KF4.CreateBall:KFb4 = True

KHb1 = False:KH1.enabled = False
KHb2 = False:KH2.enabled = False
KHb3 = False:KH3.enabled = False
KHb4 = False:KH4.enabled = False

Sub BSpeed_hit
    Ballspeed = BallVel(activeball)
End Sub

Sub KC_Hit
    PlaySoundAt "fx_collide", KF1
    If BallSpeed > 4 Then
        PushKC(Ballspeed * 0.75)
    End If
End Sub

Sub PushKC(f)
    If f < 3 Then Exit Sub
    If f > 15 Then f = 15
    If KFb4 Then KFkick4(f):KF4.enabled = True:Exit Sub:End If
    If KFb3 Then KF4.enabled = False:KFkick3(f):KF3.enabled = True:Exit Sub:End If
    If KFb2 Then KF3.enabled = False:KFkick2(f):KF2.enabled = True:Exit Sub:End If
    If KFb1 Then KF2.enabled = False:KFkick1(f):KF1.enabled = True:Exit Sub:End If
End Sub

Sub KH1_Timer
    If not Post.IsDropped Then
        KH1.TimerEnabled = False
        Exit Sub
    End If

    If KHb1 Then KHkick1:KH2.enabled = False:Exit Sub:End If
    If KHb2 Then KHkick2:KH3.enabled = False:Exit Sub:End If
    If KHb3 Then KHkick3:KH4.enabled = False:Exit Sub:End If
    If KHb4 Then KHkick4:Exit Sub:End If

    If KFb1 and KFb2 and KFb3 and KFb4 Then
        KC.disabled = False
        KH1.TimerEnabled = False
    End If
End Sub

Sub KF1_Hit:KFb1 = True:KF1.enabled = False:KF2.enabled = True:End Sub
Sub KF2_Hit:KFb2 = True:KF2.enabled = False:KF3.enabled = True:End Sub
Sub KF3_Hit:KFb3 = True:KF3.enabled = False:KF4.enabled = True:End Sub
Sub KF4_Hit:KFb4 = True:KF4.enabled = False:End Sub

Sub KH1_Hit:KHb1 = True:KH1.enabled = False:KH2.enabled = True:End Sub
Sub KH2_Hit:KHb2 = True:KH2.enabled = False:KH3.enabled = True:End Sub
Sub KH3_Hit:KHb3 = True:KH3.enabled = False:KH4.enabled = True:End Sub
Sub KH4_Hit:KHb4 = True:KH4.enabled = False:End Sub

Sub KFkick1(f):KF1.Kick cdircf, f:KFb1 = False:End Sub
Sub KFkick2(f):KF2.Kick cdircf, f:KFb2 = False:End Sub
Sub KFkick3(f):KF3.Kick cdircf, f:KFb3 = False:End Sub
Sub KFkick4(f):KF4.Kick cdircf, f:KFb4 = False:End Sub

Sub KHkick1:KH1.Kick cdirch, 1:KHb1 = False:End Sub
Sub KHkick2:KH2.Kick cdirch, 1:KHb2 = False:End Sub
Sub KHkick3:KH3.Kick cdirch, 1:KHb3 = False:End Sub
Sub KHkick4:KH4.Kick cdirch, 1:KHb4 = False:End Sub

'************************************
'          LEDs Display
'************************************

LampTimer.Enabled = 1

Dim Digits(28)

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

Sub LampTimer_Timer
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
Const maxvel = 32 'max ball velocity
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

    ' Exit the Sub If no balls on the table
    If UBound(BOT) = lob - 1 Then Exit Sub 'there no extra balls on this table

    ' play the rolling sound for each ball and draw the shadow
    For b = lob to UBound(BOT)
        aBallShadow(b).X = BOT(b).X
        aBallShadow(b).Y = BOT(b).Y
        aBallShadow(b).Height = BOT(b).Z - Ballsize / 2

        If BallVel(BOT(b)) > 1 Then
            If BOT(b).z < 30 Then
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
        If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
            PlaySound "fx_balldrop", 0, BOT(b).velz ^2 / 1000, Pan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
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

'***********************
' Ball Collision Sound
'***********************

Sub OnBallBallCollision(ball1, ball2, velocity)
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

'******************
' GI routines
'******************

Sub GiON
    For each x in aGiLights
        x.State = 1
    Next
End Sub

Sub GiOFF
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

'*********************************
' Table Options F12 User Options
'*********************************
' Table1.Option arguments are:
' - option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional array of literal strings

Dim LUTImage, StagedFlipperMod

Sub Table1_OptionEvent(ByVal eventId)
    Dim x, y

    'LUT
    LutImage = Table1.Option("Select LUT", 0, 21, 1, 0, 0, Array("Normal 0", "Normal 1", "Normal 2", "Normal 3", "Normal 4", "Normal 5", "Normal 6", "Normal 7", "Normal 8", "Normal 9", "Normal 10", _
        "Warm 0", "Warm 1", "Warm 2", "Warm 3", "Warm 4", "Warm 5", "Warm 6", "Warm 7", "Warm 8", "Warm 9", "Warm 10") )
    UpdateLUT

    ' Cabinet rails
    x = Table1.Option("Cabinet Rails", 0, 1, 1, 1, 0, Array("Hide", "Show") )
    For each y in aRails:y.visible = x:next

    ' Staged Flippers
    StagedFlipperMod = Table1.Option("Staged Flippers (uses Magna Save Keys)", 0, 1, 1, 0, 0, Array("No", "Yes") )
    If StagedFlipperMod = 1 Then
        keyStagedFlipperR = RightMagnaSave
        keyStagedFlipperL = LeftMagnaSave
    End If
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


'Verne's World / IPD No. 4619 / 1996 / 4 Players
'Spinball S.A.L., of Fuenlabrada, Spain (1995-1996)
'VPX.8 table by jpsalas, based on the table by freneticamnesic & gtxjoe
'DIP Switch 12: 3 or 5 balls when selected.

Option Explicit
Randomize

Const BallSize = 50
Const BallMass = 1

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName = "vrnwrld"

Dim VarHidden, UseVPMColoredDMD
If Table1.ShowDT = true then
    UseVPMColoredDMD = true
    VarHidden = 1
Else
    UseVPMColoredDMD = False
    VarHidden = 0
End If

if B2SOn = true then VarHidden = 1

LoadVPM "01210000", "spinball.vbs", 3.1

'********************
'Standard definitions
'********************

Const UseSolenoids = 2
Const UseLamps = 1
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

Dim Mech3bank, bsTrough, bsVUK, bsSVUK, visibleLock, bsTEject, dtRDrop, PlungerIM, bsBallKick, x

Sub table1_Init
    vpmInit me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Verne's World - Spinball 1996" & vbNewLine & "VPX8 table by JPSalas v6.0.0"
        .Games(cGameName).Settings.Value("sound") = 1
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
    vpmNudge.TiltSwitch = swTilt
    vpmNudge.Sensitivity = 3
    vpmNudge.TiltObj = Array(Bumper001, Bumper002, LeftSlingshot)

    ' Impulse Plunger
    Const IMPowerSetting = 40
    Const IMTime = 0.1
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swplunger, IMPowerSetting, IMTime
        .Switch 74
        .Random .2
        .InitExitSnd "fx_popper", "fx_kicker"
        .CreateEvents "plungerIM"
    End With

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0, 73, 72, 71, 70, 0, 0, 0
        .InitKick BallRelease, 90, 6
        .InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .Balls = 4
    End With

    ' Thalamus, more randomness pls
    Set bsVUK = New cvpmBallStack
    With bsVUK
        .InitSaucer sw67, 67, 0, 58
        .InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .InitAddSnd "fx_kicker_enter"
        .KickZ = 1.56
        .KickForceVar = 3
        .KickAngleVar = 3
    End With

    Set bsSVUK = New cvpmBallStack
    With bsSVUK
        .InitSw 0, 50, 0, 0, 0, 0, 0, 0
        .InitKick sw50a, 0, 30
        .KickZ = 1.5
        .InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .InitAddSnd "fx_kicker_enter"
        .KickBalls = 1
        .KickForceVar = 3
        .KickAngleVar = 3
    End With

    ' Drop targets
    Set dtRDrop = New cvpmDropTarget
    With dtRDrop
        .InitDrop Array(sw41, sw42, sw43, sw44), Array(41, 42, 43, 44)
        .InitSnd SoundFX("fx_droptarget", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
        .CreateEvents "dtRDrop"
    End With

    vpmMapLights aLights

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

    ' Turn on Gi
    vpmtimer.addtimer 1500, "GiOn '"

    ' init some VP objects
    volcanodiverter2.IsDropped = 1
    volcanodiverter.IsDropped = 0
    csdiverter.IsDropped = 1
    pulpopost.IsDropped = 0
End Sub

'******************
' RealTime Updates
'******************

' Animations
Sub LeftFlipper_Animate:leftflip.Rotz = LeftFlipper.CurrentAngle:End Sub
Sub RightFlipper_Animate:rightflip.Rotz = RightFlipper.CurrentAngle:End Sub
Sub RightFlipper001_Animate:rightflip001.Rotz = RightFlipper001.CurrentAngle:End Sub
Sub Gate006_Animate:g001.Rotx = 28 - Gate006.Currentangle / 2:End Sub
Sub Gate004_Animate:g002.Rotx = 28 - Gate004.Currentangle / 2:End Sub

Sub OctopusFlipperA_Animate
    OctopusBox.TransY = OctopusFlipperA.CurrentAngle
    Octopus.X = 942 - OctopusFlipperA.CurrentAngle
    Octopus.Y = 128 + OctopusFlipperA.CurrentAngle
End Sub

'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
    If keycode = LeftTiltKey Then Nudge 90, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 8:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If Keycode = LeftFlipperKey then Controller.Switch(133) = 1
    If Keycode = RightFlipperKey then Controller.Switch(131) = 1
    If Keycode = keyInsertCoin2 then vpmTimer.AddTimer 750, "vpmTimer.PulseSw swCoin1'"
    If Keycode = keyInsertCoin3 then vpmTimer.AddTimer 750, "vpmTimer.PulseSw swCoin1'"
    If keycode = keyFront Then Controller.Switch(75) = 1
    If vpmKeyDown(keycode) Then Exit Sub
    If keycode = PlungerKey Then Controller.Switch(86) = 1
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If Keycode = LeftFlipperKey then Controller.Switch(133) = 0
    If Keycode = RightFlipperKey then Controller.Switch(131) = 0
    If keycode = keyFront Then Controller.Switch(75) = 0
    If vpmKeyUp(keycode) Then Exit Sub
    If keycode = PlungerKey Then Controller.Switch(86) = 0
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
    PlaySound "fx_gion"
    For each bulb in aGiLights
        bulb.State = 1
    Next
End Sub

Sub GiOff
    Dim bulb
    PlaySound "fx_gioff"
    For each bulb in aGiLights
        bulb.State = 0
    Next
End Sub

Sub ChangeGiIntensity(factor) 'changes the intensity scale
    Dim bulb
    For each bulb in aGiLights
        bulb.IntensityScale = factor
    Next
End Sub

'*********
' Switches
'*********

' Slings
Dim LStep

Sub LeftSlingShot_Slingshot
    PlaySoundAt SoundFX("fx_slingshot", DOFContactors), Lemk
    LeftSling001.Visible = 0
    LeftSling004.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 40
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 1:LeftSLing004.Visible = 0:LeftSLing003.Visible = 1:Lemk.RotX = 14
        Case 2:LeftSLing003.Visible = 0:LeftSLing002.Visible = 1:Lemk.RotX = 2
        Case 3:LeftSLing002.Visible = 0:LeftSling001.Visible = 1:Lemk.RotX = -20:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

' Scoring rubbers

' Bumpers
Sub Bumper001_Hit:vpmTimer.PulseSw 46:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper001:End Sub
Sub Bumper002_Hit:vpmTimer.PulseSw 56:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper002:End Sub

' Drain & holes
Sub Drain_Hit:PlaysoundAt "fx_drain", Drain:bsTrough.AddBall Me:End Sub
Sub sw67_Hit:bsVUK.AddBall 0:End Sub
Sub sw50_Hit:bsSVUK.AddBall Me:End Sub

' Rollovers
Sub sw94_Hit:Controller.Switch(94) = 1:PlaySoundAt "fx_sensor", sw94:End Sub
Sub sw94_UnHit:Controller.Switch(94) = 0:End Sub

Sub sw62_Hit:Controller.Switch(62) = 1:PlaySoundAt "fx_sensor", sw62:End Sub
Sub sw62_UnHit:Controller.Switch(62) = 0:End Sub

Sub sw63_Hit:Controller.Switch(63) = 1:PlaySoundAt "fx_sensor", sw63:End Sub
Sub sw63_UnHit:Controller.Switch(63) = 0:End Sub

Sub sw94a_Hit:Controller.Switch(94) = 1:PlaySoundAt "fx_sensor", sw94:End Sub
Sub sw94a_UnHit:Controller.Switch(94) = 0:End Sub

Sub sw64_Hit:Controller.Switch(64) = 1:PlaySoundAt "fx_sensor", sw64:End Sub
Sub sw64_UnHit:Controller.Switch(64) = 0:End Sub

Sub sw87_Hit:Controller.Switch(87) = 1:PlaySoundAt "fx_sensor", sw87:End Sub
Sub sw87_UnHit:Controller.Switch(87) = 0:End Sub

Sub sw61_Hit:Controller.Switch(61) = 1:PlaySoundAt "fx_sensor", sw61:End Sub
Sub sw61_UnHit:Controller.Switch(61) = 0:End Sub

Sub sw77_Hit:Controller.Switch(77) = 1:PlaySoundAt "fx_sensor", sw77:End Sub
Sub sw77_UnHit:Controller.Switch(77) = 0:End Sub

Sub sw80_Hit:Controller.Switch(80) = 1:PlaySoundAt "fx_sensor", sw80:End Sub
Sub sw80_UnHit:Controller.Switch(80) = 0:End Sub

Sub sw81_Hit:Controller.Switch(81) = 1:PlaySoundAt "fx_sensor", sw81:End Sub
Sub sw81_UnHit:Controller.Switch(81) = 0:End Sub

Sub sw82_Hit:Controller.Switch(82) = 1:PlaySoundAt "fx_sensor", sw82:End Sub
Sub sw82_UnHit:Controller.Switch(82) = 0:End Sub

Sub sw84_Hit:Controller.Switch(84) = 1:PlaySoundAt "fx_sensor", sw84:sw84.TimerEnabled = 1:End Sub
Sub sw84_Timer:Controller.Switch(84) = 0:sw84.TimerEnabled = 0:End Sub

Sub sw83_Hit:Controller.Switch(83) = 1:PlaySoundAt "fx_sensor", sw83:End Sub
Sub sw83_UnHit:Controller.Switch(83) = 0:End Sub

Sub sw57_Hit
    PlaySoundAt "fx_sensor", sw57
    If DisableSw57 = 0 Then
        Controller.Switch(57) = 1
        DisableSw57 = 0
    End If
End Sub
Sub sw57_UnHit:Controller.Switch(57) = 0:End Sub

Sub sw92_Hit:Controller.Switch(92) = 1:PlaySoundAt "fx_sensor", sw92:End Sub
Sub sw92_UnHit:Controller.Switch(92) = 0:End Sub

Sub sw93_Hit:Controller.Switch(93) = 1:PlaySoundAt "fx_sensor", sw93:End Sub
Sub sw93_UnHit:Controller.Switch(93) = 0:End Sub

Sub sw60_Hit:Controller.Switch(60) = 1:PlaySoundAt "fx_sensor", sw60:End Sub
Sub sw60_UnHit:Controller.Switch(60) = 0:End Sub

Sub sw76_Hit:Controller.Switch(76) = 1:PlaySoundAt "fx_sensor", sw76:End Sub
Sub sw76_UnHit:Controller.Switch(76) = 0:End Sub

'Targets
Sub sw66_Hit:vpmTimer.PulseSw 66:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw51_Hit:vpmTimer.PulseSw 51:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw52_Hit:vpmTimer.PulseSw 52:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw53_Hit:vpmTimer.PulseSw 53:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw54_Hit:vpmTimer.PulseSw 54:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw47_Hit:vpmTimer.PulseSw 47:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw90_Hit:vpmTimer.PulseSw 90:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw91_Hit:vpmTimer.PulseSw 91:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

'*********
'Solenoids
'*********
SolCallback(4) = "bsTrough.Solout" 'salida bola
SolCallback(6) = "PulpoDiverter"
SolCallback(8) = "SolRightBank"    'bancada diana
SolCallback(9) = "VolcEntrada"
SolCallback(10) = "bsVUK.SolOut"
SolCallback(11) = "solAutofire" 'autolaunch 'lanzador bola
SolCallBack(12) = "solCSDiverter"
SolCallback(13) = "bsSVUK.SolOut"
'SolCallback(15) = "SolLeftPop" 'left bumper
'SolCallback(16) = "SolRightPop" 'right bumper
SolCallback(20) = "SolPulpo" 'octopus ATTTTTTTACK
SolCallBack(23) = "corkscrewMotor"
SolCallback(25) = "vpmNudge.SolGameOn"

Sub PulpoDiverter(Enabled)
    If Enabled Then
        pulpopost.isdropped = true
    Else
        pulpopost.isdropped = false
    End If
End Sub

Sub SolPulpo(Enabled)
    If Enabled Then
        OctopusFlipper.RotatetoEnd
        OctopusFlipperA.RotatetoEnd
    Else
        OctopusFlipper.RotatetoStart
        OctopusFlipperA.RotatetoStart
    End If
End Sub

Sub SolRightBank(Enabled)
    If Enabled Then
        dtRDrop.DropSol_On
    End If
End Sub

Sub VolcEntrada(Enabled)
    If Enabled Then
        volcanodiverter.isdropped = true
        volcanodiverter2.isdropped = false
    Else
        volcanodiverter.isdropped = false
        volcanodiverter2.isdropped = true
    End If
End Sub

Sub solAutofire(Enabled)
    If Enabled Then
        PlungerIM.AutoFire
    End If
End Sub

Sub SolCSDiverter(Enabled)
    csdiverter.isdropped = Not Enabled
End Sub

'*************
'  corkscrew
'*************

Dim csball, csball1, csball2, csball3, csball4, csballontop, csballReleased
Dim csball1On, csball2On, csball3On

csball = 0
csball1On = False:csball2On = False:csball3On = False

Sub csKicker_Hit
    cskicker.destroyball
    if csball = 0 then csballontop = 1
    csball = csball + 1
    Controller.Switch(45) = 1
    Select Case csball
        Case 1:csBall1On = True:Set csball1 = corkb1.createball:csball1.z = csLowerSwitchPos:csball1.x = 487:csball1.y = 173:MoveBalloon = 0
        Case 2:csBall2On = True:Set csball2 = corkb2.createball:csball2.z = csLowerSwitchPos:csball2.x = 487:csball2.y = 173
        Case 3:csBall3On = True:Set csball3 = corkb3.createball:csball3.z = csLowerSwitchPos:csball3.x = 487:csball3.y = 173
    End Select
End Sub

Sub corkscrewMotor(enabled)
    corkscrewMotorTimer.enabled = enabled
End Sub
Sub corkscrewMotorTimer_timer
    corkscrew.ObjRotZ = corkscrew.ObjRotZ + 12                                                                'turn screw
    if csball1On AND csball >= 1 then csball1.z = csball1.z + csStepsize:checkballoonswitches csball1, corkb1 'raise ball1
    if csball2On AND csball >= 2 then csball2.z = csball2.z + csStepsize:checkballoonswitches csball2, corkb2 'raise ball2
    if csball3On AND csball >= 3 then csball3.z = csball3.z + csStepsize:checkballoonswitches csball3, corkb3 'raise ball3
    if MoveBalloon = 1 then Balloon.transZ = BalloonBall.z - csMiddleSwitchPos                                'Raise Balloon
end sub

Const csLowerSwitchPos = 0
Const csMiddleSwitchPos = 100
Const csUpperSwitchPos = 200
Const csStepsize = 2.5
Dim DisableSw57, BalloonBall, MoveBalloon

Sub checkBalloonSwitches(ball, kickername)
    Select Case ball.z
        Case csLowerSwitchPos + 2 * csStepsize:Controller.Switch(45) = 0 'Clear Sw45
        Case 0:Controller.Switch(85) = 1
        Case 5:Controller.Switch(85) = 0
        Case csMiddleSwitchPos:Controller.Switch(55) = 1 'Ball at middle sw 55
        Case csMiddleSwitchPos + 2 * csStepsize:Controller.Switch(55) = 0:Set BalloonBall = ball:
            If MoveBalloon = 0 Then MoveBalloon = 1:End If
        Case csUpperSwitchPos:Controller.Switch(65) = 1 'Ball at upper sw 65
        Case csUpperSwitchPos + 2 * csStepsize:         'Ball is at top, kick ball onto ramp
            Controller.Switch(65) = 0
            If csBall1On Then
                csBall1On = False
            ElseIf csBall2On Then
                csBall2On = False
            ElseIf csBall3On Then
                csBall3On = False
            End If
            kickername.kick 180, 1
            MoveBalloonT.Enabled = 1:MoveBalloon = 2
            csballReleased = csballReleased + 1
            if csball = 3 and csballReleased = 1 then DisableSw57 = 1                     'BUGFIX: if multiball and 1st ball released, disable the exit ramp switch once, so motor works
            if csBall = csballReleased then csball = 0:csballReleased = 0:DisableSw57 = 0 'if all balls released, reset ball counts
    End Select
End Sub

Dim MoveBalloonCnt
Sub MoveBalloonT_Timer
    MoveBalloonCnt = MoveBalloonCnt + 1
    Select Case MoveBalloonCnt
        Case 1:Balloon.transZ = 151
        Case 2:Balloon.transZ = 148
        Case 3:Balloon.transZ = 144
        Case 4:Balloon.transZ = 137
        Case 5:Balloon.transZ = 129
        Case 6:Balloon.transZ = 119
        Case 7:Balloon.transZ = 107
        Case 18 Balloon.transZ = 94
        Case 9:Balloon.transZ = 78
        Case 10:Balloon.transZ = 62
        Case 11:Balloon.transZ = 42
        Case 12:Balloon.transZ = 22
        Case 13:Balloon.transZ = 0:MoveBalloonT.Enabled = 0:MoveBalloonCnt = 0
            If csball> 0 Then MoveBalloon = 1:End If
    End Select
End Sub

Sub corkb1_Unhit:Set cs1 = 0:End Sub
Sub corkb2_Unhit:Set cs2 = 0:End Sub
Sub corkb3_Unhit:Set cs3 = 0:End Sub

'*******************
'  Flipper Subs
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
        RightFlipper001.RotateToEnd
        RightFlipperOn = 1
    Else
        PlaySoundAt SoundFX("fx_flipperdown", DOFFlippers), RightFlipper
        RightFlipper.RotateToStart
        RightFlipper001.RotateToStart
        RightFlipperOn = 0
    End If
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper001_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
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
Sub aTargets_Hit(idx):PlaySound SoundFX("fx_target", DOFTargets), 0, Vol(ActiveBall), pan(ActiveBall), 0.2, Pitch(ActiveBall) * 10, 0, 0, AudioFade(ActiveBall):End Sub
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
    If tmp> 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 200
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = (SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

Function AudioFade(ball) 'only on VPX 10.4 and newer
    Dim tmp
    tmp = ball.y * 2 / TableHeight-1
    If tmp> 0 Then
        AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Sub PlaySoundAt(soundname, tableobj) 'play sound at X and Y position of an object, mostly bumpers, flippers and other fast objects
    PlaySound soundname, 0, 1, Pan(tableobj), 0.1, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname) ' play a sound at the ball position, like rubbers, targets, metals, plastics
    PlaySound soundname, 0, Vol(ActiveBall), pan(ActiveBall), 0.2, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Function RndNbr(n) 'returns a random number between 1 and n
    Randomize timer
    RndNbr = Int((n * Rnd) + 1)
End Function

'***********************************************
'   JP's VP10.8 Rolling Sounds
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
        If BallVel(BOT(b) )> 1 Then
            If BOT(b).z < 30 Then
                ballpitch = Pitch(BOT(b) )
                ballvol = Vol(BOT(b) )
            Else
                ballpitch = Pitch(BOT(b) ) + 50000 'increase the pitch on a ramp
                ballvol = Vol(BOT(b) ) * 2
            End If
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, ballvol, Pan(BOT(b) ), 0, ballpitch, 1, 0, AudioFade(BOT(b) )
        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If

        ' rothbauerw's Dropping Sounds
        If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z> 27 Then 'height adjust for ball drop sounds
            PlaySound "fx_balldrop", 0, ABS(BOT(b).velz) / 17, Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
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

'*****************************
' Ball 2 Ball Collision Sound
'*****************************

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


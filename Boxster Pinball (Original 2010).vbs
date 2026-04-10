' Boxster Pinball - original table from 2022
' https://pinballnews.com/learn/boxster/index.html
' Uses the ROM of Sega's Checkpoint and the layout of Sega's TNMT
' VPX8 table, jpsalas 2025, v1.0.0
' Graphics by Flying Dutchman

Option Explicit
Randomize

Const BallSize = 50
Const BallMass = 1

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Dim VarHidden, UseVPMColoredDMD, x
If Table1.ShowDT = true then
    UseVPMColoredDMD = true
    VarHidden = 1 'hide the vpinmame dmd
    For each x in aReels:x.visible = 1:Next
Else
    UseVPMColoredDMD = False
    VarHidden = 0
    For each x in aReels:x.visible = 0:Next
End If

const UseVPMModSol = True

LoadVPM "03060000", "DE.VBS", 3.26

'********************
'Standard definitions
'********************

Const cGameName = "ckpt_a17"
Const UseSolenoids = 1
Const UseLamps = 1
Const UseGI = 0
Const UseSync = 0
Const HandleMech = 0

' Standard Sounds
Const SSolenoidOn = "fx_SolenoidOn"
Const SSolenoidOff = "fx_SolenoidOff"
Const SCoin = "fx_Coin"

Dim bsTrough, plungerIM, bsPit, cbLeft, DtL, DtC, ttS

'************
' Table init.
'************

Sub table1_Init
    vpmInit me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Boxster Pinball" & vbNewLine & "VPX8 table by JPSalas v1.0.0"
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
    vpmNudge.TiltSwitch = 1
    vpmNudge.Sensitivity = 3
    vpmNudge.TiltObj = Array(Bumper001, Bumper002, Bumper003, LeftSlingshot, RightSlingshot)

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 10, 13, 12, 11, 0, 0, 0, 0
        .InitKick BallRelease, 90, 4
        .InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .Balls = 3
        .IsTrough = 1
    End With

    ' Impulse Plunger - used as the autoplunger
    Const IMPowerSetting = 62 'Plunger Power
    Const IMTime = 0.6        ' Time in seconds for Full Plunge
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swPlunger, IMPowerSetting, IMTime
        .Random 0.3
        .switch 23
        .InitExitSnd SoundFX("fx_plunger", DOFContactors), SoundFX("fx_plunger", DOFContactors)
        .CreateEvents "plungerIM"
    End With

    ' Saucer
    Set bsPit = New cvpmBallStack
    With bsPit
        .InitSaucer sw37, 37, 195, 26
        .InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .KickForceVar = 2
        .KickAngleVar = 2
    End With

    ' captive ball
    Set cbLeft = New cvpmCaptiveBall
    With cbLeft
        .InitCaptive CapTrigger1, CapWall1, CapKicker1, 0
        .ForceTrans = .9
        .MinForce = 3.5
        .CreateEvents "cbLeft"
        .Start
    End With

    'drop targets
    Set dtC = New cvpmDropTarget
    dtC.InitDrop Array(SW33, SW34, SW35), Array(33, 34, 35)
    dtC.InitSnd SoundFX("fx_droptarget", DOFContactors), SoundFX("fx_resetdrop", DOFContactors)
    dtC.CreateEvents "dtC"

    Set dtL = New cvpmDropTarget
    dtL.InitDrop sw44, 44
    dtL.InitSnd SoundFX("fx_droptarget", DOFContactors), SoundFX("fx_resetdrop", DOFContactors)
    dtL.CreateEvents "dtL"

    'turn Table
    Set ttS = New cvpmTurnTable
    With ttS
        .InitTurnTable tttrigger, 50
        .SpinCW = True
        .SpinUp = 200
        .SpinDown = 200
        .CreateEvents "ttS"
    End With

    vpmMapLights aLights

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

' Turn on Gi
' vpmtimer.addtimer 1500, "GiOn '"
End Sub

' ***************
' Spinning Disc
' ***************

Dim discAngle, stepAngle, stopDisc

discAngle = 0

Sub SpinDisc
    ttS.MotorOn = 1
    stepAngle = 20.0
    stopDisc = False
    DiscTimer.Enabled = True
    StopDiscTimer.Enabled = 0
    StopDiscTimer.Interval = 3000
    StopDiscTimer.Enabled = 1
End Sub

Sub DiscTimer_Timer()
    discAngle = (discAngle + stepAngle) MOD 360

    ttPrin.RotZ = discAngle

    If stopDisc Then
        stepAngle = stepAngle - 0.1
        If stepAngle <= 0 Then
            DiscTimer.Enabled = False
        End If
    End If
End Sub

Sub StopDiscTimer_Timer
    StopDiscTimer.Enabled = 0
    StopDisc = True
    ttS.MotorOn = 0
End Sub


'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
    If keycode = LeftTiltKey Then Nudge 90, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 8:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If vpmKeyDown(keycode) Then Exit Sub
    If keycode = PlungerKey Then PlaySoundat "fx_PlungerPull", Plunger:Plunger.Pullback
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If vpmKeyUp(keycode) Then Exit Sub
    If keycode = PlungerKey Then PlaySoundAt "fx_plunger", Plunger:Plunger.Fire
End Sub

'*********
'Solenoids
'*********
SolCallback(1) = "bsTrough.SolIn"
SolCallback(2) = "bsTrough.SolOut"
SolCallback(3) = "bsPit.SolOut"
' SolCallback(4)="dtC.SolDropUp"
SolCallback(5) = "dtC.SolDropUp"
SolCallback(6) = "dtL.SolDropUp"
' SolCallback(7)="bsVuk.SolOut"
SolCallback(8) = "solAutofire"
SolCallback(9) = "dtL.SolDropDown"
'SolCallback(10)='Not used
SolCallback(11) = "GIUpdate"
SolCallback(12) = "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"

SolModCallback(13) = "Flasher13"
SolModCallback(14) = "Flasher14"
SolModCallback(15) = "Flasher15"
SolModCallback(16) = "Flasher16"
'SolCallback(17)="vpmSolSound ""fx_bumper1"","
'SolCallback(18)="vpmSolSound ""fx_bumper2"","
'SolCallback(19)="vpmSolSound ""fx_bumper3"","
'SolCallback(20)="vpmSolSound ""right_slingshot"","
'SolCallback(21)="vpmSolSound ""left_slingshot"","
'SolCallback(22)= 'Spinning Disc  Backglass
'SolCallback(23)= 'Not Used
'SolCallback(24)= 'Not Used
SolModCallback(25) = "Flasher25"
SolModCallback(26) = "Flasher26"
SolModCallback(27) = "Flasher27"
SolModCallback(28) = "Flasher28"
SolModCallback(29) = "Flasher29"
SolModCallback(30) = "Flasher30"
SolModCallback(31) = "Flasher31"

' Modulated flashers
Sub Flasher13(m):m = m / 255:f13.State = m:End Sub

Sub Flasher14(m):m = m / 255:f14.State = m:f14a.State = m:End Sub
Sub Flasher15(m):m = m / 255:f15.State = m:f15a.State = m:End Sub
Sub Flasher16(m):m = m / 255:f16.State = m:End Sub
Sub Flasher25(m):m = m / 255:f25.State = m:End Sub
Sub Flasher26(m):m = m / 255:f26.State = m:f26a.State = m:End Sub
Sub Flasher27(m):m = m / 255:f27.State = m:f27b.State = m:End Sub
Sub Flasher28(m):m = m / 255:f28.State = m:f28A.State = m:f28B.State = m:f28C.State = m:f28d.State = m:f28e.State = m:End Sub
Sub Flasher29(m):m = m / 255:f29.State = m:End Sub
Sub Flasher30(m):m = m / 255:f30.State = m:f30a.State = m:f30b.State = m:f30c.State = m:End Sub
Sub Flasher31(m):m = m / 255:f31.State = m:f31a.State = m:f31b.State = m:f31c.State = m:End Sub

'*************************
' GI - needs new vpinmame
'*************************

'Set GICallback = GetRef("GIUpdate")

Sub GIUpdate(Enabled)
    If Enabled Then
        GiOff
    Else
        GiOn
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

'*********
' Switches
'*********

' Slings
Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    PlaySoundAt SoundFX("fx_slingshot", DOFContactors), Lemk
    'DOF 101, DOFPulse
    LeftSLing001.Visible = 0
    LeftSling004.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 21
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 1:LeftSLing004.Visible = 0:LeftSLing003.Visible = 1:Lemk.RotX = 14
        Case 2:LeftSLing003.Visible = 0:LeftSLing002.Visible = 1:Lemk.RotX = 2
        Case 3:LeftSLing002.Visible = 0:LeftSLing001.Visible = 1:Lemk.RotX = -20:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
    PlaySoundAt SoundFX("fx_slingshot", DOFContactors), Remk
    'DOF 102, DOFPulse
    RightSling001.Visible = 0
    RightSling004.Visible = 1
    Remk.RotX = 26
    RStep = 0
    vpmTimer.PulseSw 22
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 1:RightSLing004.Visible = 0:RightSLing003.Visible = 1:Remk.RotX = 14
        Case 2:RightSLing003.Visible = 0:RightSLing002.Visible = 1:Remk.RotX = 2
        Case 3:RightSLing002.Visible = 0:RightSLing001.Visible = 1:Remk.RotX = -20:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

' Scoring rubbers

' Bumpers
Sub Bumper001_Hit:vpmTimer.PulseSw 46:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper001:End Sub
Sub Bumper002_Hit:vpmTimer.PulseSw 48:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper002:End Sub
Sub Bumper003_Hit:vpmTimer.PulseSw 47:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper003:End Sub

' Drain & holes
Sub Drain_Hit:PlaysoundAt "fx_drain", Drain:bsTrough.AddBall Me:End Sub
Sub sw37_Hit:PlaysoundAt "fx_kicker_enter", sw37:bsPit.AddBall 0:End Sub

' Rollovers
Sub sw17_Hit:Controller.Switch(17) = 1:PlaySoundAt "fx_sensor", sw17:End Sub
Sub sw17_UnHit:Controller.Switch(17) = 0:End Sub

Sub sw18_Hit:Controller.Switch(18) = 1:PlaySoundAt "fx_sensor", sw18:End Sub
Sub sw18_UnHit:Controller.Switch(18) = 0:End Sub

Sub sw19_Hit:Controller.Switch(19) = 1:PlaySoundAt "fx_sensor", sw19:End Sub
Sub sw19_UnHit:Controller.Switch(19) = 0:End Sub

Sub sw20_Hit:Controller.Switch(20) = 1:PlaySoundAt "fx_sensor", sw20:End Sub
Sub sw20_UnHit:Controller.Switch(20) = 0:End Sub

Sub sw28_Hit:Controller.Switch(28) = 1:PlaySoundAt "fx_sensor", sw28:End Sub
Sub sw28_UnHit:Controller.Switch(28) = 0:End Sub

Sub sw29_Hit:Controller.Switch(29) = 1:PlaySoundAt "fx_sensor", sw29:End Sub
Sub sw29_UnHit:Controller.Switch(29) = 0:End Sub

'special triggers for this table

Sub sw36_Hit:Controller.Switch(36) = 1:PlaySoundAt "fx_sensor", sw36:End Sub 'really it is a droptarget
Sub sw36_UnHit:Controller.Switch(36) = 0:End Sub

Sub sw24_Hit:Controller.Switch(24) = 1:PlaySoundAt "fx_sensor", sw24:End Sub 'vuk enter, now on right ramp
Sub sw24_UnHit:Controller.Switch(24) = 0:End Sub

Sub sw30_Hit:Controller.Switch(30) = 1:PlaySoundAt "fx_sensor", sw30:End Sub 'vuk exit. now on right ramp
Sub sw30_UnHit:Controller.Switch(30) = 0:End Sub

Sub sw24a_Hit:Controller.Switch(24) = 1:PlaySoundAt "fx_sensor", sw24:End Sub 'vuk enter, now on right ramp
Sub sw24a_UnHit:Controller.Switch(24) = 0:End Sub

Sub sw30a_Hit:Controller.Switch(30) = 1:PlaySoundAt "fx_sensor", sw30:End Sub 'vuk exit. now on right ramp
Sub sw30a_UnHit:Controller.Switch(30) = 0:End Sub

'hot nitro Lights
Sub t001_Hit:l001.Duration 1, 300, 0:End Sub
Sub t002_Hit:l002.Duration 1, 300, 0:End Sub
Sub t003_Hit:l003.Duration 1, 300, 0:End Sub
Sub t004_Hit:l004.Duration 1, 300, 0:End Sub
Sub t005_Hit:l005.Duration 1, 300, 0:End Sub

'Targets
Sub sw41_Hit:vpmTimer.PulseSw 41:PlaySoundAtBall SoundFX("fx_target", DOFTargets):SpinDisc:End Sub
Sub sw42_Hit:vpmTimer.PulseSw 42:PlaySoundAtBall SoundFX("fx_target", DOFTargets):SpinDisc:End Sub
Sub sw43_Hit:vpmTimer.PulseSw 43:PlaySoundAtBall SoundFX("fx_target", DOFTargets):SpinDisc:End Sub
Sub sw23_Hit:vpmTimer.PulseSw 23:PlaySoundAtBall SoundFX("fx_target", DOFTargets):fsw23.Duration 2, 300, 0:End Sub
Sub sw45_Hit:vpmTimer.PulseSw 45:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

'lap targets
Sub sw27_Hit:vpmTimer.PulseSw 27:PlaySoundAtBall SoundFX("fx_target", DOFTargets):SpinDisc:End Sub
Sub sw26_Hit:vpmTimer.PulseSw 26:PlaySoundAtBall SoundFX("fx_target", DOFTargets):SpinDisc:End Sub
Sub sw25_Hit:vpmTimer.PulseSw 25:PlaySoundAtBall SoundFX("fx_target", DOFTargets):SpinDisc:End Sub
Sub sw27a_Hit:vpmTimer.PulseSw 27:PlaySoundAtBall SoundFX("fx_target", DOFTargets):SpinDisc:End Sub
Sub sw26a_Hit:vpmTimer.PulseSw 26:PlaySoundAtBall SoundFX("fx_target", DOFTargets):SpinDisc:End Sub
Sub sw25a_Hit:vpmTimer.PulseSw 25:PlaySoundAtBall SoundFX("fx_target", DOFTargets):SpinDisc:End Sub

'spinner
Sub sw38_Spin:vpmTimer.PulseSw 38:PlaySoundAt "fx:spinner", sw38:End Sub

'Solenoid subs

Sub solAutofire(Enabled)
    If Enabled Then
        PlungerIM.AutoFire
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

'Sub LeftFlipper_Animate:LeftFlipperTop.RotZ = LeftFlipper.CurrentAngle:End Sub
'Sub RightFlipper_Animate:RightFlipperTop.RotZ = RightFlipper.CurrentAngle:End Sub

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
    If LeftFlipper.CurrentAngle >= LeftFlipper.StartAngle - SOSAngle Then LeftFlipper.Strength = FlipperPower * SOSTorque else LeftFlipper.Strength = FlipperPower:End If

    'End Of Stroke Routine : Livecatch and Emply/Full-Charged EOS
    If LeftFlipperOn = 1 Then
        If LeftFlipper.CurrentAngle = LeftFlipper.EndAngle then
            LeftFlipper.EOSTorque = FullStrokeEOS_Torque
            LLiveCatchTimer = LLiveCatchTimer + 1
            If LLiveCatchTimer <LiveCatchSensivity Then
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
            If RLiveCatchTimer <LiveCatchSensivity Then
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
    Pitch = BallVel(ball) * 20
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
'   JP's VP10 Rolling Sounds
'***********************************************

Const tnob = 19   'total number of balls
Const lob = 1     'number of locked balls
Const maxvel = 45 'max ball velocity
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
            If BOT(b).z <30 Then
                ballpitch = Pitch(BOT(b) )
                ballvol = Vol(BOT(b) )
            Else
                ballpitch = Pitch(BOT(b) ) + 30000 'increase the pitch on a ramp
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
        If BOT(b).VelZ <-1 and BOT(b).z <55 and BOT(b).z> 27 Then 'height adjust for ball drop sounds
            PlaySound "fx_balldrop", 0, ABS(BOT(b).velz) / 17, Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
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

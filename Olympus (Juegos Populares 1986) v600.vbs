'Olympus / IPD No. 5140 / 1986 / 4 Players
'Juegos Populares, S.A., of Madrid, Spain (1986)
'a VPX8 table by jpsalas, mfuegeman, akiles50000 and pedator
'vpinmame code by mfuegeman, based on his whiteboard Olympus table.
'graphics by akiles50000
'directb2s by pedator

Option Explicit
Randomize

Const BallSize = 50
Const BallMass = 1

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Dim x

Const cGameName = "olympus"

Dim VarHidden
If Table1.ShowDT = true then
    For each x in aReel
        x.Visible = 1
    Next
    VarHidden = 1
Else
    For each x in aReel
        x.Visible = 0
    Next
    VarHidden = 0
End If

if B2SOn = true then VarHidden = 1

LoadVPM "01560000", "juegos.vbs", 3.2

'********************
'Standard definitions
'********************

Const UseSolenoids = 1
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

Dim bsTrough, LeftDropTargetBank, RightDropTargetBank, Flipperactive, bsHole

Sub table1_Init
    vpmInit me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Olympus - Juegos Populares 1986" & vbNewLine & "VPX table by JPSalas 6.0.0"
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
    vpmNudge.TiltSwitch = 31
    vpmNudge.Sensitivity = 5
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, LeftSlingshot, RightSlingshot)

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0, 25, 0, 0, 0, 0, 0, 0
        .InitKick BallRelease, 90, 4
        .InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .Balls = 1
    End With

    'Kicker Hole
    set bsHole = new cvpmSaucer
    bsHole.InitKicker Hole, 9, 275, 23, 0
    bsHole.InitExitVariance 5, 2
    bsHole.InitSounds SoundFX("fx_kicker_enter", DOFContactors), SoundFX("fx_solenoidOn", DOFContactors), SoundFX("fx_kicker", DOFContactors)

    set LeftDropTargetBank = new cvpmDropTarget
    LeftDropTargetBank.InitDrop Array(LeftDT), Array(5)
    LeftDropTargetBank.InitSnd SoundFX("fx_droptarget", DOFContactors), SoundFX("fx_resetdrop", DOFContactors)
    LeftDropTargetBank.CreateEvents "LeftDropTargetBank"

    set RightDropTargetBank = new cvpmDropTarget
    RightDropTargetBank.InitDrop Array(DT1, DT2, DT3), Array(8, 7, 6)
    RightDropTargetBank.InitSnd SoundFX("fx_droptarget", DOFContactors), SoundFX("fx_resetdrop", DOFContactors)
    RightDropTargetBank.CreateEvents "RightDropTargetBank"

    vpmMapLights aLights

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

Dim L18, OL18:OL18 = 0
Dim L58, OL58:OL58 = 0
Dim L97, OL97:OL97 = 0
Dim L105, OL105:OL105 = 0

Sub Gate001_Animate: Plungergate.Rotz = - Gate001.CurrentAngle / 4: End Sub

Sub RealTime_Timer
    RollingUpdate
    UpdateLEDs
    ' Lights acting like solenoids, code to avoid multiple solenoids after each other
    L18 = Li18.State
    L58 = Li58.State
    L97 = Li97.State
    L105 = Li105.State

    If L18 <> OL18 Then
        If L18 = 1 Then
            RightDropTargetBank.DropSol_On
        End If
    End If
    OL18 = L18

    If L58 <> OL58 Then
        If L58 = 1 Then
            LeftDropTargetBank.DropSol_On
        End If
    End If
    OL58 = L58

    If L97 <> OL97 Then
        If L97 = 1 Then
            SolVariTargetReset True
        End If
    End If
    OL97 = L97

    If L105 <> OL105 Then
        If L105 = 1 Then
            bsHole.ExitSol_On
        End If
    End If
    OL105 = L105
End Sub

'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
    If keycode = LeftFlipperKey AND Flipperactive Then SolLFlipper 1
    If keycode = RightFlipperKey AND Flipperactive Then SolRFlipper 1
    If keycode = LeftTiltKey Then Nudge 90, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 8:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If vpmKeyDown(keycode) Then Exit Sub
    If keycode = PlungerKey Then PlaySoundat "fx_PlungerPull", Plunger:Plunger.Pullback
    If keycode = KeyRules Then Rules
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If keycode = LeftFlipperKey AND Flipperactive Then SolLFlipper 0
    If keycode = RightFlipperKey AND Flipperactive Then SolRFlipper 0
    If vpmKeyUp(keycode) Then Exit Sub
    If keycode = PlungerKey Then PlaySoundAt "fx_plunger", Plunger:Plunger.Fire
End Sub

'*************************
' GI - needs new vpinmame
'*************************

'Set GICallback = GetRef("GIUpdate")

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
    vpmTimer.PulseSw 14
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
    vpmTimer.PulseSw 13
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
Sub rlband009_Hit:vpmtimer.pulsesw(18):End Sub
Sub rlband008_Hit:vpmtimer.pulsesw(18):End Sub
Sub rlband007_Hit:vpmtimer.pulsesw(18):End Sub
Sub rlband006_Hit:vpmtimer.pulsesw(18):End Sub
Sub rlband005_Hit:vpmtimer.pulsesw(18):End Sub

' Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 16:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper1:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 15:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper2:End Sub

' Drain & holes
Sub Drain_Hit:PlaysoundAt "fx_drain", Drain:bsTrough.AddBall Me:End Sub
Sub hole_Hit:PlaysoundAt "fx_kicker_enter", hole:bsHole.AddBall 0:End Sub

' Rollovers
Sub LeftOrbit_Hit:Controller.Switch(1) = 1::PlaySoundAt "fx_sensor", LeftOrbit:End Sub
Sub LeftOrbit_unhit:Controller.Switch(1) = 0:End Sub
Sub RightOrbit_Hit:Controller.Switch(2) = 1::PlaySoundAt "fx_sensor", RightOrbit:End Sub
Sub RightOrbit_unhit:Controller.Switch(2) = 0:End Sub
Sub sw10_Hit:Controller.Switch(10) = 1:PlaySoundAt "fx_sensor", sw10:End Sub
Sub sw10_UnHit:Controller.Switch(10) = 0:End Sub
Sub sw11_Hit:Controller.Switch(11) = 1:PlaySoundAt "fx_sensor", sw11:End Sub
Sub sw11_UnHit:Controller.Switch(11) = 0:End Sub
Sub sw11a_Hit:Controller.Switch(11) = 1:PlaySoundAt "fx_sensor", sw11a:End Sub
Sub sw11a_UnHit:Controller.Switch(11) = 0:End Sub
Sub sw33_Hit:Controller.Switch(33) = 1:PlaySoundAt "fx_sensor", sw33:End Sub
Sub sw33_UnHit:Controller.Switch(33) = 0:End Sub
Sub sw33a_Hit:Controller.Switch(33) = 1:PlaySoundAt "fx_sensor", sw33a:End Sub
Sub sw33a_UnHit:Controller.Switch(33) = 0:End Sub
Sub sw34_Hit:Controller.Switch(34) = 1:PlaySoundAt "fx_sensor", sw34:End Sub
Sub sw34_UnHit:Controller.Switch(34) = 0:End Sub
Sub sw34a_Hit:Controller.Switch(34) = 1:PlaySoundAt "fx_sensor", sw34a:End Sub
Sub sw34a_UnHit:Controller.Switch(34) = 0:End Sub

'Targets
Sub sw3_Hit:vpmTimer.PulseSw 3:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw3a_Hit:vpmTimer.PulseSw 3:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw4_Hit:vpmTimer.PulseSw 4:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw4a_Hit:vpmTimer.PulseSw 4:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw12_Hit:vpmTimer.PulseSw 12:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw12a_Hit:vpmTimer.PulseSw 12:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

'*********
'Solenoids
'*********
SolCallback(7) = "Sol7"              'SolGameOn
SolCallback(8) = "vpmFlasher gi001," 'insert light
SolCallback(13) = "Sol13"            'Right Drop Target Bank
SolCallback(14) = "Sol14"            'Gates open
SolCallback(15) = "vpmSolSound SoundFX(""fx_Knocker"",DOFKnocker),"
SolCallback(16) = "bsTrough.SolOut"

Sub Sol7(enabled)
    VpmNudge.SolGameOn enabled
    Flipperactive = enabled

    if not enabled Then
        LeftFlipper.RotateToStart
        RightFlipper.RotateToStart
        RightFlipper2.RotateToStart
    end if
end sub

Sub Sol8(enabled)
    OLamp.state = abs(enabled)
    OLamp2.state = abs(enabled)
end sub

Sub Sol13(enabled)
    if enabled then
        RightDropTargetBank.DropSol_On
    end if
End Sub

Sub Sol14(enabled)
    Gate003.open = enabled
    Gate004.open = enabled
  If enabled Then
    PGate3O.visible=1
    PGate3C.visible=0
    PGate4O.visible=1
    PGate4C.visible=0
  Else
    PGate3O.visible=0
    PGate3C.visible=1
    PGate4O.visible=0
    PGate4C.visible=1
  End If

End Sub

'*******************
' Flipper Subs Rev3
'*******************

'SolCallback(sLRFlipper) = "SolRFlipper"
'SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
    If Enabled Then
    Controller.Switch(84)=1
        PlaySoundAt SoundFX("fx_flipperup", DOFFlippers), LeftFlipper
        LeftFlipper.RotateToEnd
        LeftFlipperOn = 1
    Else
    Controller.Switch(84)=0
        PlaySoundAt SoundFX("fx_flipperdown", DOFFlippers), LeftFlipper
        LeftFlipper.RotateToStart
        LeftFlipperOn = 0
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
    Controller.Switch(82)=1
        PlaySoundAt SoundFX("fx_flipperup", DOFFlippers), RightFlipper
        RightFlipper.RotateToEnd
        RightFlipper2.RotateToEnd
        RightFlipperOn = 1
    Else
    Controller.Switch(82)=0
        PlaySoundAt SoundFX("fx_flipperdown", DOFFlippers), RightFlipper
        RightFlipper.RotateToStart
        RightFlipper2.RotateToStart
        RightFlipperOn = 0
    End If
End Sub

' flippers top animations

Sub LeftFlipper_Animate: LeftFlipperTop.RotZ = LeftFlipper.CurrentAngle: End Sub
Sub RightFlipper_Animate: RightFlipperTop.RotZ = RightFlipper.CurrentAngle: End Sub
Sub RightFlipper2_Animate: RightFlipperTop2.RotZ = RightFlipper2.CurrentAngle: End Sub

' flippers hit Sound

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, Vol(ActiveBall), pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, Vol(ActiveBall), pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper2_Collide(parm)
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

'*********************
'LED's based on Eala's
'*********************
Dim Digits(32)
Digits(0) = Array(a00, a02, a05, a06, a04, a01, a03, a07)
Digits(1) = Array(a10, a12, a15, a16, a14, a11, a13)
Digits(2) = Array(a20, a22, a25, a26, a24, a21, a23)
Digits(3) = Array(a30, a32, a35, a36, a34, a31, a33, a37)
Digits(4) = Array(a40, a42, a45, a46, a44, a41, a43)
Digits(5) = Array(a50, a52, a55, a56, a54, a51, a53)
Digits(6) = Array(a60, a62, a65, a66, a64, a61, a63)

Digits(7) = Array(b00, b02, b05, b06, b04, b01, b03, b07)
Digits(8) = Array(b10, b12, b15, b16, b14, b11, b13)
Digits(9) = Array(b20, b22, b25, b26, b24, b21, b23)
Digits(10) = Array(b30, b32, b35, b36, b34, b31, b33, b37)
Digits(11) = Array(b40, b42, b45, b46, b44, b41, b43)
Digits(12) = Array(b50, b52, b55, b56, b54, b51, b53)
Digits(13) = Array(b60, b62, b65, b66, b64, b61, b63)

Digits(14) = Array(e00, e02, e05, e06, e04, e01, e03, e07)
Digits(15) = Array(e10, e12, e15, e16, e14, e11, e13)
Digits(16) = Array(e20, e22, e25, e26, e24, e21, e23)
Digits(17) = Array(e30, e32, e35, e36, e34, e31, e33, e37)
Digits(18) = Array(e40, e42, e45, e46, e44, e41, e43)
Digits(19) = Array(e50, e52, e55, e56, e54, e51, e53)
Digits(20) = Array(e60, e62, e65, e66, e64, e61, e63)

Digits(21) = Array(f00, f02, f05, f06, f04, f01, f03, f07)
Digits(22) = Array(f10, f12, f15, f16, f14, f11, f13)
Digits(23) = Array(f20, f22, f25, f26, f24, f21, f23)
Digits(24) = Array(f30, f32, f35, f36, f34, f31, f33, f37)
Digits(25) = Array(f40, f42, f45, f46, f44, f41, f43)
Digits(26) = Array(f50, f52, f55, f56, f54, f51, f53)
Digits(27) = Array(f60, f62, f65, f66, f64, f61, f63)

Digits(28) = Array(c00, c02, c05, c06, c04, c01, c03)
Digits(29) = Array(c10, c12, c15, c16, c14, c11, c13)

Digits(30) = Array(d00, d02, d05, d06, d04, d01, d03)
Digits(31) = Array(d10, d12, d15, d16, d14, d11, d13)

Digits(32) = Array(c20, c22, c25, c26, c24, c21, c23)

'********************
'Update LED's display
'********************

Sub UpdateLeds()
    Dim ChgLED, ii, num, chg, stat, obj
    ChgLED = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED) Then
        For ii = 0 To UBound(chgLED)
            num = chgLED(ii, 0):chg = chgLED(ii, 1):stat = chgLED(ii, 2)
            For Each obj In Digits(num)
                If chg And 1 Then obj.State = stat And 1
                chg = chg \ 2:stat = stat \ 2
            Next
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
                ballpitch = Pitch(BOT(b) ) + 50000 'increase the pitch on a ramp
                ballvol = Vol(BOT(b) ) * 3
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

'************
' Varitarget
'************

Dim vtpos, varipos
vtpos = Array(20, 18, 16, 14, 12, 10, 8, 6, 4, 2, 0, -2, -4, -6, -8, -10, -12, -14, -16, -18, -20)

Sub vt_Hit(idx)
    Dim x
    If ActiveBall.VelY <0 Then
        varipos = idx
        ActiveBall.VelY = ActiveBall.VelY * 0.915
        PlaySound "fx_solenoidoff"
        varitarget.roty = vtpos(idx)
    Else
        If VTTimer.Enabled = False Then
            VTTimer.Interval = 800
            VTTimer.Enabled = True
        End If
    End If
End Sub

Sub SolVariTargetReset(Enabled)
    If Enabled and VariPos> 1 AND VTTimer.Enabled = 0 Then
        VTTimer.Interval = 1500
        VTTimer.Enabled = True
    end if
End Sub

Sub VTTimer_Timer() 'moving down the vari target
    VTTimer.Interval = 35
    varipos = varipos -1
    If varipos <0 Then varipos = 0
    If varipos> 19 Then varipos = 19
    varitarget.roty = vtpos(varipos)
    If varipos = 0 Then VTTimer.Enabled = False
End Sub

'Hit positions:

Sub VT_001_Hit:Controller.Switch(24) = 1:End Sub
Sub VT_001_UnHit:Controller.Switch(24) = 0:End Sub

Sub VT_002_Hit:Controller.Switch(23) = 1:End Sub
Sub VT_002_UnHit:Controller.Switch(23) = 0:End Sub

Sub VT_003_Hit:Controller.Switch(22) = 1:End Sub
Sub VT_003_UnHit:Controller.Switch(22) = 0:End Sub

Sub VT_004_Hit:Controller.Switch(21) = 1:VariWall.Isdropped = 1:End Sub
Sub VT_004_UnHit:Controller.Switch(21) = 0:End Sub

Sub VT_005_Hit:Controller.Switch(20) = 1:End Sub
Sub VT_005_UnHit:Controller.Switch(20) = 0:End Sub

Sub VT_006_Hit:Controller.Switch(19) = 1:VariWall.Isdropped = 0:VariWall.TimerEnabled = 1:End Sub
Sub VT_006_UnHit:Controller.Switch(19) = 0:End Sub

Sub VariWall_Timer: VariWall.Isdropped = 1: VariWall.TimerEnabled = 0: End Sub

'-------------------------------
'------  DIP Switch Menu  ------
'-------------------------------
'by mfuegeman

Sub editDips
  Dim vpmDips:Set vpmDips=New cvpmDips
  With vpmDips
    .AddForm 200,280,"Olympus DIP switches"
    .AddFrame 0, 0,270,"Balls per game",&H00000400,Array("3 balls",0,"5 balls",&H00000400)
    .AddChk   5,50,270,Array("Match feature",32768)
    .AddFrame 0,75,270,"Replay threshold (Extra ball/credit/credit)",&H00000300,Array("3.0M / 3.7M / 4.5M points",&H00000300,"2.5M / 3.2M / 4.0M points",&H00000100,"1.8M / 2.5M / 3.0M points",&H00000200,"1.2M / 1.8M / 2.5M points",0)
    .AddChk   5,155,270,Array("Enable replay extra ball",&H00000008)

    .AddLabel 0,180,280,20,"After hitting OK, press F3 to reset game with new settings."
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

    ' Apron Card Language
    x = Table1.Option("Apron Card Language", 0, 1, 1, 1, 0, Array("Spanish", "English") )
    If x Then Apron.Image = "Plastics" Else Apron.Image = "Plastics-ES"
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


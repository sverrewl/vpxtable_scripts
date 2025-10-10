'Black Velvet / IPD No. 315 / May, 1978 / 4 Players
'Game Plan, Incorporated, of Illinois (1978-1985)
'VPX8 by jpsalas 2025, version 1.0.1
'includes my physycs 4.3.1
'vpinmame script by destruk

Option Explicit
Randomize

Const BallSize = 50
Const BallMass = 1

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Dim bsTrough, bsSaucer, bsSaucer2, bsSaucer3, x

Const cGameName = "blvelvet"

Dim VarHidden 'hide the vpinmame window
VarHidden = 1

LoadVPM "01200000", "GamePlan.vbs", 3.1

'********************
'Standard definitions
'********************

Const UseSolenoids = 2
Const UseLamps = 1
Const UseGI = 0
Const UseSync = 1
Const HandleMech = 0

' Standard Sounds
Const SSolenoidOn = "fx_SolenoidOn"
Const SSolenoidOff = "fx_SolenoidOff"
Const SCoin = "fx_Coin"

'*********
'Solenoids
'*********
SolCallback(3) = "vpmSolSound SoundFX(""fx_Knocker"",DOFKnocker),"
SolCallback(8) = "bsTrough.SolOut"
'SolCallback(9)="vpmSolSound ""lsling"","
'SolCallback(10)="vpmSolSound ""lsling"","
SolCallback(11) = "bsSaucer3.SolOut"
'SolCallback(12)="vpmSolSound ""jet3"","
SolCallback(13) = "bsSaucer2.SolOut"
'SolCallback(14)="vpmSolSound ""jet3"","
SolCallback(15) = "bsSaucer.SolOut"
SolCallback(16) = "vpmNudge.SolGameOn"
'SolCallback(sLLFlipper)="vpmSolFlipper LeftFlipper,Nothing,"
'SolCallback(sLRFlipper)="vpmSolFlipper RightFlipper,Nothing,"
SolCallback(32) = "LuckTime"
'Solenoid 32=Chuck o Luck Motor

'************
' Table init.
'************

Sub table1_Init
    vpmInit me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Black Velvet - Game Plan 1978" & vbNewLine & "VPX8 table by jpsalas 1.0.0"
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

' old default configuration, set up for: extra  ball, 5 ball, but use F6 it is easier to configure
'Controller.Dip(0) = (0*1 + 1*2 + 0*4 + 0*8 + 0*16 + 0*32 + 0*64 + 0*128) '01-08
'Controller.Dip(1) = (0*1 + 0*2 + 0*4 + 0*8 + 0*16 + 0*32 + 0*64 + 1*128) '09-16
'Controller.Dip(2) = (0*1 + 0*2 + 0*4 + 0*8 + 0*16 + 0*32 + 0*64 + 0*128) '17-24
'Controller.Dip(3) = (1*1 + 1*2 + 1*4 + 1*8 + 1*16 + 1*32 + 1*64 + 0*128) '25-32   28 OFF is 3 Balls and ON is 5 balls, 29 OFF is extra ball and ON is replay

' Nudging
    vpmNudge.TiltSwitch = swTilt
    vpmNudge.Sensitivity = 5
    vpmNudge.TiltObj = Array(Bumper001, Bumper002, LeftSlingshot, RightSlingshot)

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0, 11, 0, 0, 0, 0, 0, 0
        .InitKick BallRelease, 90, 4
        .InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .Balls = 1
    End With

    ' Thalamus, more randomness pls
    ' Saucer Top center
    Set bsSaucer = New cvpmBallStack
    With bsSaucer
        .InitSaucer Kicker1, 24, 174, 16
        .InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .KickForceVar = 3
        .KickAngleVar = 3
    End With

    ' Saucer Left
    Set bsSaucer2 = New cvpmBallStack
    With bsSaucer2
        .InitSaucer Kicker2, 12, 120, 16
        .InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .KickForceVar = 3
        .KickAngleVar = 3
    End With

    ' Saucer Right
    Set bsSaucer3 = New cvpmBallStack
    With bsSaucer3
        .InitSaucer Kicker3, 10, 240, 16
        .InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .KickForceVar = 3
        .KickAngleVar = 3
    End With

    vpmMapLights aLights

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

    LedsTimer.Enabled = VarHidden

    ' Turn on Gi
    vpmtimer.addtimer 1500, "GiOn '"
End Sub

Sub PlungerTimer_Timer
    PlungerR.Y = 1465 + (5 * Plunger.Position) -20
end sub

'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If vpmKeyDown(keycode) Then Exit Sub
    If keycode = PlungerKey Then PlaySoundat "fx_PlungerPull", Plunger:Plunger.Pullback
End Sub

Sub table1_KeyUp(ByVal Keycode)
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
Dim LStep, RStep, RStep2

Sub LeftSlingShot_Slingshot
    PlaySoundAt SoundFX("fx_slingshot", DOFContactors), Lemk
    LeftSling004.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 15
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

' Bumpers
Sub Bumper001_Hit:vpmTimer.PulseSw 21:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper001:End Sub
Sub Bumper002_Hit:vpmTimer.PulseSw 22:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper002:End Sub

' Targets

Sub sw9_Slingshot:PlaySoundAtBall "fx_target":vpmTimer.PulseSw 9:End Sub
Sub sw9a_Slingshot:PlaySoundAtBall "fx_target":vpmTimer.PulseSw 9:End Sub

' Drain & holes
Sub Drain_Hit:PlaysoundAt "fx_drain", Drain:bsTrough.AddBall Me:End Sub
Sub kicker1_Hit:PlaysoundAt "fx_kicker_enter", kicker1:bsSaucer.AddBall 0:End Sub
Sub kicker2_Hit:PlaysoundAt "fx_kicker_enter", kicker2:bsSaucer2.AddBall 0:End Sub
Sub kicker3_Hit:PlaysoundAt "fx_kicker_enter", kicker3:bsSaucer3.AddBall 0:End Sub

' Rollovers
Sub sw14_Hit:Controller.Switch(14) = 1:PlaySoundAt "fx_sensor", sw14:End Sub
Sub sw14_UnHit:Controller.Switch(14) = 0:End Sub

Sub sw16_Hit:Controller.Switch(16) = 1:PlaySoundAt "fx_sensor", sw16:End Sub
Sub sw16_UnHit:Controller.Switch(16) = 0:End Sub

Sub sw18_Hit:Controller.Switch(18) = 1:PlaySoundAt "fx_sensor", sw18:End Sub
Sub sw18_UnHit:Controller.Switch(18) = 0:End Sub

Sub sw18a_Hit:Controller.Switch(18) = 1:PlaySoundAt "fx_sensor", sw18a:End Sub
Sub sw18a_UnHit:Controller.Switch(18) = 0:End Sub

Sub sw18b_Hit:Controller.Switch(18) = 1:PlaySoundAt "fx_sensor", sw18b:End Sub
Sub sw18b_UnHit:Controller.Switch(18) = 0:End Sub

Sub sw18c_Hit:Controller.Switch(18) = 1:PlaySoundAt "fx_sensor", sw18c:End Sub
Sub sw18c_UnHit:Controller.Switch(18) = 0:End Sub

Sub sw20_Hit:Controller.Switch(20) = 1:PlaySoundAt "fx_sensor", sw20:End Sub
Sub sw20_UnHit:Controller.Switch(20) = 0:End Sub

Sub sw20a_Hit:Controller.Switch(20) = 1:PlaySoundAt "fx_sensor", sw20a:End Sub
Sub sw20a_UnHit:Controller.Switch(20) = 0:End Sub

Sub sw20b_Hit:Controller.Switch(20) = 1:PlaySoundAt "fx_sensor", sw20b:End Sub
Sub sw20b_UnHit:Controller.Switch(20) = 0:End Sub

Sub sw25_Hit:Controller.Switch(25) = 1:PlaySoundAt "fx_sensor", sw25:End Sub
Sub sw25_UnHit:Controller.Switch(25) = 0:End Sub

Sub sw25a_Hit:Controller.Switch(25) = 1:PlaySoundAt "fx_sensor", sw25a:End Sub
Sub sw25a_UnHit:Controller.Switch(25) = 0:End Sub

'Spinners

Sub Spinner1_Spin():vpmTimer.PulseSw 23:PlaySoundAt "fx_spinner", Spinner1:End Sub
Sub Spinner2_Spin():vpmTimer.PulseSw 17:PlaySoundAt "fx_spinner", Spinner2:End Sub

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
        RightFlipperOn = 1
    Else
        PlaySoundAt SoundFX("fx_flipperdown", DOFFlippers), RightFlipper
        RightFlipper.RotateToStart
        RightFlipperOn = 0
    End If
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper_Collide(parm)
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

'------------------D I S P L A Y  C O D E -----------------------------

Dim Digits(60)
Digits(0) = Array(Light26, Light27, Light29, Light32, Light31, Light30, Light28)
Digits(1) = Array(Light43, Light33, Light35, Light41, Light38, Light36, Light34)
Digits(2) = Array(Light45, Light46, Light48, Light51, Light50, Light49, Light47)
Digits(3) = Array(Light58, Light52, Light54, Light57, Light56, Light55, Light53)
Digits(4) = Array(Light59, Light60, Light62, Light65, Light64, Light63, Light61)
Digits(5) = Array(Light72, Light66, Light68, Light71, Light70, Light69, Light67)
Digits(6) = Array(Light73, Light74, Light76, Light79, Light78, Light77, Light75)
Digits(7) = Array(Light86, Light80, Light82, Light85, Light84, Light83, Light81)
Digits(8) = Array(Light87, Light88, Light90, Light93, Light92, Light91, Light89)
Digits(9) = Array(Light100, Light94, Light96, Light99, Light98, Light97, Light95)
Digits(10) = Array(Light101, Light102, Light104, Light107, Light106, Light105, Light103)
Digits(11) = Array(Light114, Light108, Light110, Light113, Light112, Light111, Light109)
Digits(12) = Array(Light129, Light130, Light132, Light135, Light134, Light133, Light131)
Digits(13) = Array(Light142, Light136, Light138, Light141, Light140, Light139, Light137)
Digits(14) = Array(Light143, Light144, Light146, Light149, Light148, Light147, Light145)
Digits(15) = Array(Light156, Light150, Light152, Light155, Light154, Light153, Light151)
Digits(16) = Array(Light157, Light158, Light160, Light163, Light162, Light161, Light159)
Digits(17) = Array(Light170, Light164, Light166, Light169, Light168, Light167, Light165)
Digits(18) = Array(Light171, Light172, Light174, Light177, Light176, Light175, Light173)
Digits(19) = Array(Light184, Light178, Light180, Light183, Light182, Light181, Light179)
Digits(20) = Array(Light185, Light186, Light188, Light191, Light190, Light189, Light187)
Digits(21) = Array(Light198, Light192, Light194, Light197, Light196, Light195, Light193)
Digits(22) = Array(Light199, Light200, Light202, Light205, Light204, Light203, Light201)
Digits(23) = Array(Light212, Light206, Light208, Light211, Light210, Light209, Light207)
Digits(24) = Array(Light115, Light116, Light118, Light121, Light120, Light119, Light117)
Digits(25) = Array(Light128, Light122, Light124, Light127, Light126, Light125, Light123)
Digits(26) = Array(Light213, Light214, Light216, Light219, Light218, Light217, Light215)
Digits(27) = Array(Light226, Light220, Light222, Light225, Light224, Light223, Light221)

Sub LedsTimer_Timer
    Dim ChgLED, ii, num, chg, stat, obj
    ChgLED = Controller.ChangedLEDs(&H0000003f, &Hffffffff)
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
'   JP's VP10 Rolling Sounds
'***********************************************

Const tnob = 19   'total number of balls
Const lob = 0     'number of locked balls
Const maxvel = 26 'max ball velocity
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
            ballpitch = Pitch(BOT(b) )
            ballvol = Vol(BOT(b) )
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, ballvol, Pan(BOT(b) ), 0, ballpitch, 1, 0, AudioFade(BOT(b) )
        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
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

'Gameplan
'added by Inkochnito
Sub EditDips
    Dim vpmDips:Set vpmDips = New cvpmDips
    With vpmDips
        .AddForm 700, 400, "Black Velvet - DIP switches"
        .AddFrame 2, 5, 150, "Maximum credits", &H07000000, Array("5 credits", 0, "10 credits", &H01000000, "15 credits", &H02000000, "20 credits", &H03000000, "25 credits", &H04000000, "30 credits", &H05000000, "35 credits", &H06000000, "40 credits", &H07000000) 'dip 25&26&27
        .AddFrame 2, 135, 150, "High game to date award", &HC0000000, Array("no award", 0, "1 credit", &H40000000, "2 credits", &H80000000, "3 credits", &HC0000000)                                                                                                    'dip 31&32
        .AddFrame 170, 5, 150, "Special award", &H10000000, Array("extra ball", 0, "replay", &H10000000)                                                                                                                                                                'dip 29
        .AddFrame 170, 51, 150, "Balls per game", &H08000000, Array("3 balls", 0, "5 balls", &H08000000)                                                                                                                                                                'dip 28
        .AddChk 170, 117, 150, Array("Play tunes", 32768)                                                                                                                                                                                                               'dip 16
        .AddChk 170, 137, 150, Array("Match feature", &H20000000)                                                                                                                                                                                                       'dip 30
        .AddLabel 30, 230, 300, 20, "After hitting OK, press F3 to reset game with new settings."
        .ViewDips
    End With
End Sub
Set vpmShowDips = GetRef("EditDips")


'Ruleta

dim ruletapos, WheelPos
WheelPos = 0
ruletapos = 0

Sub timer001_timer
    ruletapos = (ruletapos + 5) MOD 360
    ruletapalos.Rotz = ruletapos
    ruleta.Rotz = ruletapos
    WheelPos = ruletapos \ 30
    if(ruletapos + 15) MOD 30 = 0 Then ruletaf.RotatetoEnd:playsound "clack":ruletaf.TimerEnabled = 1
End Sub

Sub ruletaf_Timer:Me.TimerEnabled = 0:ruletaf.RotatetoStart:End Sub

Sub StopRuleta
    Timer001.Enabled = 0
    ruletapos = WheelPos * 30
    ruletapalos.Rotz = ruletapos: ruleta.Rotz = ruletapos
End Sub

Sub LuckTime(Enabled)
    If Enabled Then
        Select Case WheelPos
            Case 0:Controller.Switch(32) = 0  'Double Bonus
            Case 1:Controller.Switch(28) = 0  '500
            Case 2:Controller.Switch(30) = 0  'Special
            Case 3:Controller.Switch(31) = 0  'Lites Extra
            Case 4:Controller.Switch(29) = 0  '5000
            Case 5:Controller.Switch(27) = 0  '1000
            Case 6:Controller.Switch(32) = 0  'Double Bonus
            Case 7:Controller.Switch(31) = 0  'Lites Extra
            Case 8:Controller.Switch(28) = 0  '500
            Case 9:Controller.Switch(27) = 0  '1000
            Case 10:Controller.Switch(29) = 0 '5000
            Case 11:Controller.Switch(31) = 0 'Lites Extra
        End Select
        Timer001.Enabled = 1
    Else
        StopRuleta
        Select Case WheelPos
            Case 0:Controller.Switch(32) = 1  'Double Bonus
            Case 1:Controller.Switch(28) = 1  '500
            Case 2:Controller.Switch(30) = 1  'Special
            Case 3:Controller.Switch(31) = 1  'Lites Extra
            Case 4:Controller.Switch(29) = 1  '5000
            Case 5:Controller.Switch(27) = 1  '1000
            Case 6:Controller.Switch(32) = 1  'Double Bonus
            Case 7:Controller.Switch(31) = 1  'Lites Extra
            Case 8:Controller.Switch(28) = 1  '500
            Case 9:Controller.Switch(27) = 1  '1000
            Case 10:Controller.Switch(29) = 1 '5000
            Case 11:Controller.Switch(31) = 1 'Lites Extra
        End Select
    End If
End Sub

'*********************************
' Table Options F12 User Options
'*********************************
' Table1.Option arguments are:
' - option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional array of literal strings

Dim LUTImage, chimesON

Sub Table1_OptionEvent(ByVal eventId)
    Dim x, y

    'LUT
    LutImage = Table1.Option("Select LUT", 0, 21, 1, 0, 0, Array("Normal 0", "Normal 1", "Normal 2", "Normal 3", "Normal 4", "Normal 5", "Normal 6", "Normal 7", "Normal 8", "Normal 9", "Normal 10", _
        "Warm 0", "Warm 1", "Warm 2", "Warm 3", "Warm 4", "Warm 5", "Warm 6", "Warm 7", "Warm 8", "Warm 9", "Warm 10") )
    UpdateLUT

    ' Cabinet rails
    x = Table1.Option("Cabinet Rails", 0, 1, 1, 1, 0, Array("Hide", "Show") )
    For each y in aRails:y.visible = x:y.sidevisible = x:next

    ' chimesON
    x = Table1.Option("Play chimes", 0, 1, 1, 1, 0, Array("No", "Yes") )
    If x = 1 Then chimesON = True Else chimesON = False
    If ChimesON then
        SolCallback(1) = "vpmSolSound ""BELL3"","
        SolCallback(2) = "vpmSolSound ""BELL2"","
        SolCallback(6) = "vpmSolSound ""BELL1"","
        SolCallback(7) = "vpmSolSound ""BELL4"","
     Else
        SolCallback(1) = ""
        SolCallback(2) = ""
        SolCallback(6) = ""
        SolCallback(7) = ""
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


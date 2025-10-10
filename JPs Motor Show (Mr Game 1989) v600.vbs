' JP's Motor Show - Layout and ROM from MeGame
' Motor Show / IPD No. 3631 / June, 1989 / 4 Players
' Mr. Game, of Bologna, Italy (1988-1990)
' VPX8 version 6.0.0 by JPSalas
' Used the tables from TAB/destruk as reference for the script (read: copy & paste a lot! :) )

Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const BallSize = 50
Const BallMass = 1

Const cGameName = "motrshow"
Const cCredits = "Mr. Game Motor Show"
Const UseSolenoids = 1, UseLamps = 1, UseGI = 0
Const SSolenoidOn = "fx_solenoid", SSolenoidOff = "fx_solenoidoff"
Const SCoin = "fx_coin", SFlipperOn = "", SFlipperOff = ""

Dim x

LoadVPM "01510000", "mrgame.vbs", 3.1

'**********
' Keys
'**********

Sub Table1_KeyDown(ByVal KeyCode)
    If keycode = LeftTiltKey Then Nudge 90, 4:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 4:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If KeyCode = LeftFlipperKey Then Controller.Switch(115) = 1
    If KeyCode = RightFlipperKey Then Controller.Switch(113) = 1
    If vpmKeyDown(KeyCode) Then Exit Sub
    If keycode = PlungerKey Then Plunger.Pullback
    If keycode = keyFront Then
        Controller.Dip(0) = (1 * 1 + 0 * 2 + 0 * 4 + 0 * 8 + 0 * 16 + 0 * 32 + 0 * 64 + 0 * 128) '01-08
    End If
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
    If KeyCode = LeftFlipperKey Then Controller.Switch(115) = 0
    If KeyCode = RightFlipperKey Then Controller.Switch(113) = 0
    If vpmKeyUp(keycode) Then Exit Sub
    If keycode = PlungerKey Then Plunger.Fire:PlaySound SoundFX("SolOn", DOFContactors)
End Sub

'************************************
' Game timer for real time updates
'************************************

Sub RealTime_Timer
    RollingUpdate
End Sub

'*********************
' Solenoids & Flashers
'*********************

SolCallback(1) = "vpmSolSound SoundFX(""fx_Knocker"",DOFKnocker),"
SolCallback(2) = "vpmFlasher Array(F1,F1a),"
SolCallback(3) = "vpmFlasher Array(F13,f13a),"
SolCallback(4) = "bsTrough.SolOut"
SolCallback(9) = "vpmFlasher F2,"
SolCallback(10) = "vpmFlasher F3,"
SolCallback(11) = "vpmFlasher Array(F4,F4a),"
SolCallback(12) = "vpmFlasher Array(F5,F5a,F5b,F5d),"
SolCallback(13) = "vpmFlasher Array(F6,F6a),"
SolCallback(14) = "SolStopper2"
SolCallback(15) = "vpmFlasher F7,"
SolCallback(17) = "bsTrough.SolIn"
SolCallback(18) = "vpmFlasher Array(F10,F10a),"
SolCallback(19) = "SolStopper"
SolCallback(21) = "vpmFlasher Array(F12,f12a),"
SolCallback(22) = "vpmFlasher Array(F11,f11a),"
SolCallback(23) = "vpmFlasher F8,"
SolCallback(24) = "vpmFlasher Array(F9,F9a,F9b,f9d),"
SolCallback(25) = "vpmNudge.SolGameOn"

'Not used
'SolCallback(6)="vpmSolSound ""sling"","
'SolCallback(7)="vpmSolSound ""sling"","
'SolCallback(8)="vpmSolSound ""jet3"","
'SolCallback(16)="vpmSolSound ""jet3"","
'SolCallback(20)="vpmSolSound ""jet3"","

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

Sub LeftFlipper_Animate:LeftFlipperTop.RotZ = LeftFlipper.CurrentAngle:End Sub
Sub RightFlipper_Animate:RightFlipperTop.RotZ = RightFlipper.CurrentAngle:End Sub

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

' Solenoid subs

Sub SolStopper(Enabled)
    If Enabled Then
        Ballstop1.IsDropped = 0
    Else
        Ballstop1.IsDropped = 1
    End If
End Sub

Sub SolStopper2(Enabled)
    If Enabled Then
        Ballstop2.IsDropped = 0
    Else
        Ballstop2.IsDropped = 1
    End If
End Sub

Sub Table1_Paused:Controller.Pause = 1:End Sub
Sub Table1_UnPaused:Controller.Pause = 0:End Sub

Dim bsTrough

Sub Table1_Exit
End Sub

'************
' Table init.
'************

Sub Table1_Init
    vpmInit Me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
        .Games(cGameName).Settings.Value("rol") = 0   '1= rotated display, 0= normal
        .Games(cGameName).Settings.Value("sound") = 1 '1 enabled rom sound
        .SplashInfoLine = cCredits
        .HandleMechanics = 0
        .HandleKeyboard = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .ShowTitle = 0
        .Run GetPlayerHwnd
        If Err Then MsgBox Err.Description
    End With

    'set video dip 1 to 1/ON to enable collision detection video modes

    Controller.Dip(0) = (1 * 1 + 0 * 2 + 0 * 4 + 0 * 8 + 0 * 16 + 0 * 32 + 0 * 64 + 0 * 128) '01-08
    'Controller.Dip(1) = (0*1 + 0*2 + 0*4 + 0*8 + 0*16 + 0*32 + 0*64 + 0*128) '09-16
    'Controller.Dip(2) = (0*1 + 0*2 + 0*4 + 0*8 + 0*16 + 0*32 + 0*64 + 0*128) '17-24
    'Controller.Dip(3) = (0*1 + 0*2 + 0*4 + 0*8 + 0*16 + 0*32 + 0*64 + 0*128) '25-32

    PinMAMETimer.Interval = PinMAMEInterval:PinMAMETimer.Enabled = 1
    vpmNudge.TiltSwitch = 10:vpmNudge.Sensitivity = 5
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingshot)

    Set bsTrough = New cvpmBallStack
    bsTrough.InitSw 16, 17, 18, 0, 0, 0, 0, 0
    bsTrough.InitKick BallRelease, 115, 8
    bsTrough.InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    bsTrough.Balls = 2

    vpmMapLights AllLamps

    Ballstop1.IsDropped = 1
    Ballstop2.IsDropped = 1

    RealTime.Enabled = 1

    ' Turn on Gi
    vpmtimer.addtimer 1500, "GiOn '"
End Sub

'*********************
' Switches & Triggers
'*********************

Sub Drain_Hit:PlaysoundAt "fx_drain", Drain:bsTrough.AddBall Me:End Sub                                  '16/17/18
Sub sw20_Hit:Controller.Switch(20) = 1:PlaySoundAt "fx_sensor", sw20:End Sub                             '20
Sub sw20_unHit:Controller.Switch(20) = 0:End Sub
Sub sw21_Hit:Controller.Switch(21) = 1:PlaySoundAt "fx_sensor", sw21:End Sub                             '21
Sub sw21_unHit:Controller.Switch(21) = 0:End Sub
Sub sw22_Hit:Controller.Switch(22) = 1:PlaySoundAt "fx_sensor", sw22:End Sub                             '22
Sub sw22_unHit:Controller.Switch(22) = 0:End Sub
Sub sw25_Hit:Controller.Switch(25) = 1:PlaySoundAt "fx_sensor", sw25:End Sub                             '25
Sub sw25_unHit:Controller.Switch(25) = 0:End Sub
Sub sw26_Hit:Controller.Switch(26) = 1:PlaySoundAt "fx_sensor", sw26:End Sub                             '26
Sub sw26_unHit:Controller.Switch(26) = 0:End Sub
Sub sw27_Hit:vpmTimer.PulseSw 27:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub                '27
Sub sw28_Hit:vpmTimer.PulseSw 28:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub                '28
Sub sw29_Hit:vpmTimer.PulseSw 29:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub                '29
Sub sw30_Hit:Controller.Switch(30) = 1:l2.State = 1:PlaySoundAt "fx_sensor", sw30:End Sub                '30
Sub sw30_unHit:Controller.Switch(30) = 0:l2.State = 0:End Sub
Sub sw31_Hit:Controller.Switch(31) = 1:PlaySoundAt "fx_sensor", sw31:End Sub                             '31
Sub sw31_unHit:Controller.Switch(31) = 0:End Sub
Sub sw32_Hit:Controller.Switch(32) = 1:PlaySoundAt "fx_sensor", sw32:End Sub                             '32
Sub sw32_unHit:Controller.Switch(32) = 0:End Sub
Sub sw33_Hit:vpmTimer.PulseSw 33:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub                '33
Sub sw34_Hit:vpmTimer.PulseSw 34:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub                '34
Sub sw35_Hit:vpmTimer.PulseSw 35:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub                '35
Sub sw36_Hit:vpmTimer.PulseSw 36:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub                '36
Sub sw37_Hit:vpmTimer.PulseSw 37:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub                '37
Sub sw38_Hit:vpmTimer.PulseSw 38:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub                '38
Sub sw39_Hit:vpmTimer.PulseSw 39:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub                '39
Sub sw39b_Hit:vpmTimer.PulseSw 39:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 40:PlaySound SoundFX("fx_bumper", DOFContactors), 0, 1, 0, 0.15:End Sub '40
Sub Bumper3_Hit:vpmTimer.PulseSw 41:PlaySound SoundFX("fx_bumper", DOFContactors), 0, 1, 0, 0.15:End Sub '41
Sub Bumper1_Hit:vpmTimer.PulseSw 42:PlaySound SoundFX("fx_bumper", DOFContactors), 0, 1, 0, 0.15:End Sub '42
Sub sw43_Hit:vpmTimer.PulseSw 43:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub                '43
Sub sw44_Hit:vpmTimer.PulseSw 44:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub                '44
Sub sw45_Hit:vpmTimer.PulseSw 45:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub                '45
Sub sw46_Hit:Controller.Switch(46) = 1:PlaySound "fx_metalrolling":End Sub                               '46
Sub sw46_unHit:Controller.Switch(46) = 0:End Sub
Sub sw47_Hit:Controller.Switch(47) = 1:PlaySoundAt "fx_sensor", sw47:End Sub                             '47
Sub sw47_unHit:Controller.Switch(47) = 0:End Sub
Sub sw48_Hit:Controller.Switch(48) = 1:PlaySoundAt "fx_sensor", sw48:End Sub                             '48
Sub sw48_unHit:Controller.Switch(48) = 0:End Sub
Sub sw49_Hit:Controller.Switch(49) = 1:PlaySoundAt "fx_sensor", sw49:End Sub                             '49
Sub sw49_unHit:Controller.Switch(49) = 0:End Sub

'***************
'  Slingshots
'***************
Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    vpmTimer.PulseSw 23
    PlaySoundAt SoundFX("fx_slingshot", DOFContactors), Lemk
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
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
    vpmTimer.PulseSw 24
    PlaySoundAt SoundFX("fx_slingshot", DOFContactors), Remk
    RightSling4.Visible = 1
    Remk.RotX = 26
    RStep = 0
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

'light 5=playfield relay

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
    PlaySound soundname, 0, 1, Pan(tableobj), 0.1, 0, 0, 0, AudioFade(tableobj)
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
Const maxvel = 44 'max ball velocity
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
        aBallShadow(b).Height = BOT(b).Z - Ballsize / 2

        If BallVel(BOT(b) )> 1 Then
            If BOT(b).z <30 Then
                ballpitch = Pitch(BOT(b) )
                ballvol = Vol(BOT(b) )
            Else
                ballpitch = Pitch(BOT(b) ) + 25000 'increase the pitch on a ramp
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

'******************
'   GI effects
' independent routine
' it turns on the gi
' when there is a ball
' in play
'******************

Dim OldGiState, GiIsOn
OldGiState = -1 'start witht the Gi off
GiIsOn = False

Sub GIUpdateTimer_Timer
    Dim tmp, obj
    tmp = Getballs
    If UBound(tmp) <> OldGiState Then
        OldGiState = Ubound(tmp)
        If UBound(tmp) = -1 Then '-1 is the value returned by Getballs when no balls are on the table, so turn off the gi
            GiOff
        Else
            Gion
        End If
    End If
End Sub

Sub GiOn
    Dim bulb
    PlaySound "fx_GiOn"
    GiIsOn = True
    For each bulb in aGiLights
        bulb.State = 1
    Next
End Sub

Sub GiOff
    Dim bulb
    PlaySound "fx_GiOff"
    GiIsOn = False
    For each bulb in aGiLights
        bulb.State = 0
    Next
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

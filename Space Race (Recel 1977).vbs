' Space Race / IPD No. 2257 / April, 1977 / 4 Players
' Recel S. A., of Madrid, Spain (1974-1986)
' VPX8 table by jpsalas, 2026

Option Explicit ' Force explicit variable declaration
Randomize

' core.vbs constants
Const BallSize = 50 ' 50 is the normal size
Const BallMass = 1  ' 1 is the normal ball mass.

'DOF config - (server needs to be updated)
'101 Left Flipper
'102 Right Flipper
'103 Left slingshot
'104 Right Slingshot
'105 left targets
'106 center left target
'107 right targets
'108 center right target
'109 Bumper Left
'110 Bumper Right
'111 Top right target
'112 Top Left Target
'113 reset left droptargets
'114 reset right droptargets
'115
'116
'117
'118
'119
'120 BallRelease
'121
'122 Knocker
'123
'124
'125
'126
'128
'129
'130
'131 Chime
'132 Chime
'133 Chime

' load extra vbs files
LoadCoreFiles

Sub LoadCoreFiles
    On Error Resume Next
    ExecuteGlobal GetTextFile("core.vbs")
    If Err Then MsgBox "Can't open core.vbs"
    On Error Resume Next
    ExecuteGlobal GetTextFile("controller.vbs")
    If Err Then MsgBox "Can't open controller.vbs"
End Sub

' Valores Constants
Const TableName = "spacerace" ' file name to save highscores and other variables
Const cGameName = "spacerace" ' B2S name & DOF config
Const MaxPlayers = 4          ' 1 to 4 can play
Const MaxMultiplier = 3       ' limit bonus multiplier
Const MaxBonus = 10           ' highest bonus count
Const FreePlay = False        ' Free play or coins

' Global variables
Dim PlayersPlayingGame
Dim CurrentPlayer
Dim Credits
Dim Bonus
Dim CBonus
Dim BallsRemaining(4)
Dim BonusMultiplier
Dim PlayfieldMultiplier
Dim ExtraBallsAwards(4)
Dim Special1Awarded(4)
Dim Special2Awarded(4)
Dim Special3Awarded(4)
Dim Special4Awarded(4)
Dim Special1
Dim Special2
Dim Special3
Dim Special4
Dim Score(4)
Dim HighScore
Dim Match
Dim Tilt
Dim TiltSensitivity
Dim Tilted
Dim Add1
Dim Add10
Dim Add100
Dim LastSwicthHit
Dim x

' Control variables
Dim BallsOnPlayfield

' Boolean variables
Dim bAttractMode
Dim bFreePlay
Dim bGameInPlay
Dim bOnTheFirstBall
Dim bExtraBallWonThisBall
Dim bJustStarted
Dim bBallInPlungerLane
Dim bBallSaverActive

' core.vbs variables

' *********************************************************************
'                Common rutines to all the tables
' *********************************************************************

Sub Table1_Init()
    Dim x

    ' Init som objects, like walls, targets
    VPObjects_Init
    LoadEM

    ' load highscore
    Credits = 0
    Loadhs
    vpmTimer.AddTimer 1000, "UpdateStartUp '"

    ' init all the global variables
    bFreePlay = FreePlay
    bAttractMode = False
    bOnTheFirstBall = False
    bGameInPlay = False
    bBallInPlungerLane = False
    BallsOnPlayfield = 0
    Tilt = 0
    TiltSensitivity = 6
    Tilted = False
    Match = 0
    bJustStarted = True
    Add1 = 0
    Add10 = 0
    Add100 = 0
    LastSwicthHit = ""

    ' setup table in game over mode
    EndOfGame

    'turn on GI lights
    vpmtimer.addtimer 1000, "GiOn '"

    ' Remove desktop items in FS mode
    If Table1.ShowDT then
        For each x in aReels
            x.Visible = 1
        Next
    Else
        For each x in aReels
            x.Visible = 0
        Next
    End If
End Sub

Sub UpdateStartup
    ScoreReel1.SetValue HSScore(1)
    CreditsReel.SetValue credits
    If B2SOn then
        Controller.B2SSetScorePlayer 1, HSScore(1)
        Controller.B2SSetCredits Credits
    End If
End Sub

'******
' Keys
'******

Sub Table1_KeyDown(ByVal Keycode)

    If EnteringInitials then
        CollectInitials(keycode)
        Exit Sub
    End If

    ' add coins
    If Keycode = AddCreditKey OR Keycode = AddCreditKey2 Then
        If(Tilted = False) Then
            AddCredits 1
            PlaySoundAt "fx_coin", coinslot
        End If
    End If

    ' plunger
    If keycode = PlungerKey Then
        Plunger.Pullback
        PlaySoundAt "fx_plungerpull", plunger
    End If

    ' tilt keys
    If keycode = LeftTiltKey Then Nudge 90, 8:PlaySound "fx_nudge", 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 8:PlaySound "fx_nudge", 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 9:PlaySound "fx_nudge", 0, 1, 1, 0.25

    ' keys during game

    If bGameInPlay AND NOT Tilted Then
        If keycode = LeftTiltKey Then CheckTilt
        If keycode = RightTiltKey Then CheckTilt
        If keycode = CenterTiltKey Then CheckTilt
        If keycode = MechanicalTilt Then CheckTilt

        If keycode = LeftFlipperKey Then SolLFlipper 1
        If keycode = RightFlipperKey Then SolRFlipper 1

        If keycode = StartGameKey Then
            If((PlayersPlayingGame <MaxPlayers) AND(bOnTheFirstBall = True) ) Then

                If(bFreePlay = True) Then
                    PlayersPlayingGame = PlayersPlayingGame + 1
                'PlayersReel.SetValue, PlayersPlayingGame
                Else
                    If(Credits> 0) then
                        PlayersPlayingGame = PlayersPlayingGame + 1
                        Credits = Credits - 1
                        UpdateCredits
                        UpdateBallInPlay
                    Else
                    ' Not Enough Credits to start a game.
                    'PlaySound "so_nocredits"
                    End If
                End If
            End If
        End If
        Else

            If keycode = StartGameKey Then
                If(bFreePlay = True) Then
                    If(BallsOnPlayfield = 0) Then
                        ResetScores
                        ResetForNewGame()
                    End If
                Else
                    If(Credits> 0) Then
                        If(BallsOnPlayfield = 0) Then
                            Credits = Credits - 1
                            UpdateCredits
                            ResetScores
                            ResetForNewGame()
                        End If
                    Else
                    ' Not Enough Credits to start a game.
                    'PlaySound "so_nocredits"
                    End If
                End If
            End If
    End If ' If (GameInPlay)
    if keycode = "3" then ScoreReel1.addvalue 256
End Sub

Sub Table1_KeyUp(ByVal keycode)

    If EnteringInitials then
        Exit Sub
    End If

    If bGameInPlay AND NOT Tilted Then
        ' teclas de los flipers
        If keycode = LeftFlipperKey Then SolLFlipper 0
        If keycode = RightFlipperKey Then SolRFlipper 0
    End If

    If keycode = PlungerKey Then
        Plunger.Fire
        If bBallInPlungerLane Then
            PlaySoundAt "fx_plunger", plunger
        Else
            PlaySoundAt "fx_plunger_empty", plunger
        End If
    End If
End Sub

'******************
' Table stop/pause
'******************

Sub table1_Paused
End Sub

Sub table1_unPaused
End Sub

Sub table1_Exit
    Savehs
    If B2SOn then
        Controller.Stop
    End If
End Sub

'********************
'     Flippers
'********************

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFXDOF("fx_flipperup", 101, DOFOn, DOFFlippers), LeftFlipper
        LeftFlipper.RotateToEnd
        LeftFlipperOn = 1
    Else
        PlaySoundAt SoundFXDOF("fx_flipperdown", 101, DOFOff, DOFFlippers), LeftFlipper
        LeftFlipper.RotateToStart
        LeftFlipperOn = 0
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFXDOF("fx_flipperup", 102, DOFOn, DOFFlippers), RightFlipper
        RightFlipper.RotateToEnd
        RightFlipperOn = 1
    Else
        PlaySoundAt SoundFXDOF("fx_flipperdown", 102, DOFOff, DOFFlippers), RightFlipper
        RightFlipper.RotateToStart
        RightFlipperOn = 0
    End If
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, Vol(ActiveBall), pan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, Vol(ActiveBall), pan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper2_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, Vol(ActiveBall), pan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub LeftFlipper_Animate
    LeftFlipperTop.Rotz = LeftFlipper.CurrentAngle
End Sub

Sub RightFlipper_Animate
    RightFlipperTop.Rotz = RightFlipper.CurrentAngle
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
Dim RLiveCatchTimer2
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
RLiveCatchTimer2 = 0

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
    If RightFlipper.CurrentAngle <= RightFlipper.StartAngle + SOSAngle Then
        RightFlipper.Strength = FlipperPower * SOSTorque
    else
        RightFlipper.Strength = FlipperPower
    End If

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

'***********
' GI lights
'***********

Sub GiOn 'enciende las luces GI
    Dim bulb
    PlaySound "fx_gion"
    For each bulb in aGiLights
        bulb.State = 1
    Next
End Sub

Sub GiOff 'apaga las luces GI
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

Sub TurnON(Col) 'turn on all the lights in a collection
    Dim i
    For each i in Col
        i.State = 1
    Next
End Sub

Sub TurnOFF(Col) 'turn off all the lights in a collection
    Dim i
    For each i in Col
        i.State = 0
    Next
End Sub

'**************
'    TILT
'**************

Sub CheckTilt
    Tilt = Tilt + TiltSensitivity
    TiltDecreaseTimer.Enabled = True
    If Tilt> 15 Then
        Tilted = True
        TurnON aTilt
        If B2SOn then
            Controller.B2SSetTilt 1
        end if
        DisableTable True
        ' BallsRemaining(CurrentPlayer) = 0 'player looses the game 'mostly on older 1 player games
        TiltRecoveryTimer.Enabled = True 'wait for all the balls to drain
    End If
End Sub

Sub TiltDecreaseTimer_Timer
    If Tilt> 0 Then
        Tilt = Tilt - 0.1
    Else
        TiltDecreaseTimer.Enabled = False
    End If
End Sub

Sub DisableTable(Enabled)
    If Enabled Then
        GiOff
        LeftFlipper.RotateToStart
        RightFlipper.RotateToStart
        Bumper001.Threshold = 100
        Bumper002.Threshold = 100
        LeftSlingshot.Disabled = 1
        RightSlingshot.Disabled = 1
    Else
        GiOn
        Bumper001.Threshold = 1
        Bumper002.Threshold = 1
        LeftSlingshot.Disabled = 0
        RightSlingshot.Disabled = 0
    End If
End Sub

Sub TiltRecoveryTimer_Timer()
    ' all the balls have drained
    If(BallsOnPlayfield = 0) Then
        EndOfBall2()
        TiltRecoveryTimer.Enabled = False
    End If
' otherwise repeat
End Sub

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

'****************************************************
'   JP's VPX Rolling Sounds with ball speed control
'****************************************************

Const tnob = 2    'total number of balls
Const lob = 0     'number of locked balls
Const maxvel = 28 'max ball velocity
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

    ' play the rolling sound for each ball
    For b = lob to UBound(BOT)
        If BallVel(BOT(b) )> 1 Then
            If BOT(b).z <0 Then
                ballpitch = Pitch(BOT(b) ) - 5000 'decrease the pitch under the playfield
                ballvol = Vol(BOT(b) )
            ElseIf BOT(b).z <30 Then
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

        ' dropping sounds
        If BOT(b).VelZ <-1 Then
            'from ramp
            If BOT(b).z <55 and BOT(b).z> 27 Then PlaySound "fx_balldrop", 0, ABS(BOT(b).velz) / 17, Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
            'down a hole
            If BOT(b).z <10 and BOT(b).z> -10 Then PlaySound "fx_hole_enter", 0, ABS(BOT(b).velz) / 17, Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
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

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
    PlaySound "fx_collide", 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
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

'************************************************************************************************************************
' Only for VPX 10.8 and higher.
' FlashForMs will blink light for TotalPeriod(ms) at rate of BlinkPeriod(ms)
' When TotalPeriod done, light or flasher will be set to FinalState value where
' Final State values are:   0=Off, 1=On, 2=Blink, -1 Return to original state
'
' To blink a flasher you need to link it to a light, this will fade the flasher just like the light
'************************************************************************************************************************

Sub FlashForMs(MyLight, TotalPeriod, BlinkPeriod, FinalState)
    If FinalState = -1 Then
        FinalState = MyLight.State
    End If
    MyLight.BlinkInterval = BlinkPeriod
    MyLight.Duration 2, TotalPeriod, FinalState
End Sub

'****************************************
' Init table for a new game
'****************************************

Sub ResetForNewGame()
    'debug.print "ResetForNewGame"
    Dim i

    bGameInPLay = True
    bBallSaverActive = False

    StopAttractMode
    If B2SOn then
        Controller.B2SSetGameOver 0
    end if

    GiOn

    CurrentPlayer = 1
    PlayersPlayingGame = 1
    bOnTheFirstBall = True
    For i = 1 To MaxPlayers
        Score(i) = 0
        ExtraBallsAwards(i) = 0
        Special1Awarded(i) = False
        Special2Awarded(i) = False
        Special3Awarded(i) = False
        Special4Awarded(i) = False
        BallsRemaining(i) = BallsPerGame
    Next
    BonusMultiplier = 1
    Bonus = 0
    CBonus = 0
    UpdateBallInPlay

    Clear_Match

    ' init other variables
    Tilt = 0

    ' init game variables
    Game_Init()

    ' start a music?
    PlaySound "gameStart"
    ' first ball
    vpmtimer.addtimer 2000, "FirstBall '"
End Sub

Sub FirstBall
    'debug.print "FirstBall"
    ' reset table for a new ball, rise droptargets ++
    ResetForNewPlayerBall()
    CreateNewBall()
End Sub

' (Re-)init table for a new ball or player

Sub ResetForNewPlayerBall()
    'debug.print "ResetForNewPlayerBall"
    AddScore 0

    ' reset multiplier to 1x

    ' turn on lights, and variables
    bExtraBallWonThisBall = False
    ResetNewBallVariables
End Sub

' Crete new ball

Sub CreateNewBall()
    'debug.print "CreateNewBall"
    BallRelease.CreateSizedBallWithMass BallSize / 2, BallMass
    BallsOnPlayfield = BallsOnPlayfield + 1
    UpdateBallInPlay
    PlaySoundAt SoundFXDOF("fx_Ballrel", 120, DOFPulse, DOFContactors), BallRelease
    BallRelease.Kick 90, 4
End Sub

' player lost the ball

Sub EndOfBall()

    ' Lost the first ball, now it cannot accept more players
    bOnTheFirstBall = False

    If NOT Tilted Then
        BonusCountTimer.Interval = 200 + 200 * BonusMultiplier
        BonusCountTimer.Enabled = 1
    End If
End Sub

Sub BonusCountTimer_Timer 'when draining
    If Bonus> 0 Then
        Bonus = Bonus -1
        AddScore 100 * BonusMultiplier
        UpdateBonusLights
    Else
        BonusCountTimer.Enabled = 0
        Bonus = 0
        CBonus = 1
        UpdateBonusLights
        vpmtimer.addtimer 1000, "EndOfBall2 '"
    End If
End Sub

Sub UpdateBonusLights
    Select Case Bonus
        Case 0:bl1.State = 0:bl2.State = 0:bl3.State = 0:bl4.State = 0:bl5.State = 0:bl6.State = 0:bl7.State = 0:bl8.State = 0:bl9.State = 0:bl10.State = 0
        Case 1:bl1.State = 1:bl2.State = 0:bl3.State = 0:bl4.State = 0:bl5.State = 0:bl6.State = 0:bl7.State = 0:bl8.State = 0:bl9.State = 0:bl10.State = 0
        Case 2:bl1.State = 0:bl2.State = 1:bl3.State = 0:bl4.State = 0:bl5.State = 0:bl6.State = 0:bl7.State = 0:bl8.State = 0:bl9.State = 0:bl10.State = 0
        Case 3:bl1.State = 0:bl2.State = 0:bl3.State = 1:bl4.State = 0:bl5.State = 0:bl6.State = 0:bl7.State = 0:bl8.State = 0:bl9.State = 0:bl10.State = 0
        Case 4:bl1.State = 0:bl2.State = 0:bl3.State = 0:bl4.State = 1:bl5.State = 0:bl6.State = 0:bl7.State = 0:bl8.State = 0:bl9.State = 0:bl10.State = 0
        Case 5:bl1.State = 0:bl2.State = 0:bl3.State = 0:bl4.State = 0:bl5.State = 1:bl6.State = 0:bl7.State = 0:bl8.State = 0:bl9.State = 0:bl10.State = 0
        Case 6:bl1.State = 0:bl2.State = 0:bl3.State = 0:bl4.State = 0:bl5.State = 0:bl6.State = 1:bl7.State = 0:bl8.State = 0:bl9.State = 0:bl10.State = 0
        Case 7:bl1.State = 0:bl2.State = 0:bl3.State = 0:bl4.State = 0:bl5.State = 0:bl6.State = 0:bl7.State = 1:bl8.State = 0:bl9.State = 0:bl10.State = 0
        Case 8:bl1.State = 0:bl2.State = 0:bl3.State = 0:bl4.State = 0:bl5.State = 0:bl6.State = 0:bl7.State = 0:bl8.State = 1:bl9.State = 0:bl10.State = 0
        Case 9:bl1.State = 0:bl2.State = 0:bl3.State = 0:bl4.State = 0:bl5.State = 0:bl6.State = 0:bl7.State = 0:bl8.State = 0:bl9.State = 1:bl10.State = 0
        Case 10:bl1.State = 0:bl2.State = 0:bl3.State = 0:bl4.State = 0:bl5.State = 0:bl6.State = 0:bl7.State = 0:bl8.State = 0:bl9.State = 0:bl10.State = 1
    End Select
    Select Case CBonus
        Case 0:cl1.State = 0:cl2.State = 0:cl3.State = 0:cl4.State = 0:cl5.State = 0:cl6.State = 0:cl7.State = 0:cl8.State = 0:cl9.State = 0:cl10.State = 0
        Case 1:cl1.State = 1:cl2.State = 0:cl3.State = 0:cl4.State = 0:cl5.State = 0:cl6.State = 0:cl7.State = 0:cl8.State = 0:cl9.State = 0:cl10.State = 0
        Case 2:cl1.State = 0:cl2.State = 1:cl3.State = 0:cl4.State = 0:cl5.State = 0:cl6.State = 0:cl7.State = 0:cl8.State = 0:cl9.State = 0:cl10.State = 0
        Case 3:cl1.State = 0:cl2.State = 0:cl3.State = 1:cl4.State = 0:cl5.State = 0:cl6.State = 0:cl7.State = 0:cl8.State = 0:cl9.State = 0:cl10.State = 0
        Case 4:cl1.State = 0:cl2.State = 0:cl3.State = 0:cl4.State = 1:cl5.State = 0:cl6.State = 0:cl7.State = 0:cl8.State = 0:cl9.State = 0:cl10.State = 0
        Case 5:cl1.State = 0:cl2.State = 0:cl3.State = 0:cl4.State = 0:cl5.State = 1:cl6.State = 0:cl7.State = 0:cl8.State = 0:cl9.State = 0:cl10.State = 0
        Case 6:cl1.State = 0:cl2.State = 0:cl3.State = 0:cl4.State = 0:cl5.State = 0:cl6.State = 1:cl7.State = 0:cl8.State = 0:cl9.State = 0:cl10.State = 0
        Case 7:cl1.State = 0:cl2.State = 0:cl3.State = 0:cl4.State = 0:cl5.State = 0:cl6.State = 0:cl7.State = 1:cl8.State = 0:cl9.State = 0:cl10.State = 0
        Case 8:cl1.State = 0:cl2.State = 0:cl3.State = 0:cl4.State = 0:cl5.State = 0:cl6.State = 0:cl7.State = 0:cl8.State = 1:cl9.State = 0:cl10.State = 0
        Case 9:cl1.State = 0:cl2.State = 0:cl3.State = 0:cl4.State = 0:cl5.State = 0:cl6.State = 0:cl7.State = 0:cl8.State = 0:cl9.State = 1:cl10.State = 0
        Case 10:cl1.State = 0:cl2.State = 0:cl3.State = 0:cl4.State = 0:cl5.State = 0:cl6.State = 0:cl7.State = 0:cl8.State = 0:cl9.State = 0:cl10.State = 1
    End Select
End Sub

' After bonus count go to the next step
'
Sub EndOfBall2()
    'debug.print "EndOfBall2"

    Tilted = False
    Tilt = 0
    TurnOFF aTilt
    If B2SOn then
        Controller.B2SSetTilt 0
    end if
    DisableTable False

    ' win extra ball?
    If(ExtraBallsAwards(CurrentPlayer)> 0) Then
        'debug.print "Extra Ball"

        ' if so then give it
        ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) - 1

        ' turn off light if no more extra balls
        If(ExtraBallsAwards(CurrentPlayer) = 0) Then
            LightShootAgain.State = 0
            If B2SOn then
                Controller.B2SSetShootAgain 0
            end if
        End If

        ' extra ball sound?

        ' reset as in a new ball
        ResetForNewPlayerBall()
        CreateNewBall()
    Else ' no extra ball

        BallsRemaining(CurrentPlayer) = BallsRemaining(CurrentPlayer) - 1

        ' last ball?
        If(BallsRemaining(CurrentPlayer) <= 0) Then
            CheckHighScore()
        End If

        ' this is not the last ball, chack for new player
        EndOfBallComplete()
    End If
End Sub

Sub EndOfBallComplete()
    'debug.print "EndOfBallComplete"
    Dim NextPlayer

    ' other players?
    If(PlayersPlayingGame> 1) Then
        NextPlayer = CurrentPlayer + 1
        ' if it is the last player then go to the first one
        If(NextPlayer> PlayersPlayingGame) Then
            NextPlayer = 1
        End If
    Else
        NextPlayer = CurrentPlayer
    End If

    'debug.print "Next Player = " & NextPlayer

    ' end of game?
    If((BallsRemaining(CurrentPlayer) <= 0) AND(BallsRemaining(NextPlayer) <= 0) ) Then

        ' match if playing with coins
        If bFreePlay = False Then
            Verification_Match
        End If

        ' end of game
        EndOfGame()
    Else
        ' next player
        CurrentPlayer = NextPlayer

        ' update score
        AddScore 0

        ' reset table for new player
        ResetForNewPlayerBall()
        CreateNewBall()
    End If
End Sub

' Called at the end of the game

Sub EndOfGame()
    'debug.print "EndOfGame"
    bGameInPLay = False
    bJustStarted = False
    If B2SOn then
        Controller.B2SSetGameOver 1
        Controller.B2SSetBallInPlay 0
        Controller.B2SSetPlayerUp 0
        Controller.B2SSetCanPlay 0
    end if
    ' turn off flippers
    SolLFlipper 0
    SolRFlipper 0

    ' start the attract mode
    vpmTimer.AddTimer 3000, "StartAttractMode '"
End Sub

' Fuction to calculate the balls left
Function Balls
    Dim tmp
    tmp = BallsPerGame - BallsRemaining(CurrentPlayer) + 1
    If tmp> BallsPerGame Then
        Balls = BallsPerGame
    Else
        Balls = tmp
    End If
End Function

' check the highscore
Sub CheckHighscore
    Dim playertops, si, sj, i, stemp, stempplayers
    For i = 1 to 4
        sortscores(i) = 0
        sortplayers(i) = 0
    Next
    playertops = 0
    For i = 1 to PlayersPlayingGame
        sortscores(i) = Score(i)
        sortplayers(i) = i
    Next
    For si = 1 to PlayersPlayingGame
        For sj = 1 to PlayersPlayingGame-1
            If sortscores(sj)> sortscores(sj + 1) then
                stemp = sortscores(sj + 1)
                stempplayers = sortplayers(sj + 1)
                sortscores(sj + 1) = sortscores(sj)
                sortplayers(sj + 1) = sortplayers(sj)
                sortscores(sj) = stemp
                sortplayers(sj) = stempplayers
            End If
        Next
    Next
    HighScoreTimer.interval = 100
    HighScoreTimer.enabled = True
    ScoreChecker = 4
    CheckAllScores = 1
    NewHighScore sortscores(ScoreChecker), sortplayers(ScoreChecker)
End Sub

'******************
'     Match
'******************

Sub Verification_Match()
    PlaySound "MotorLeer"
    Match = INT(RND(1) * 10) * 10 ' random between 0 and 90
    Display_Match
    If(Score(CurrentPlayer) MOD 100) = Match Then
        PlaySound SoundFXDOF("fx_knocker", 122, DOFPulse, DOFknocker)
        AddCredits 1
    End If
End Sub

Sub Clear_Match()
    Match = 0
    Display_Match
    m0.State = 0
    m0a.State = 0
    If B2SOn then
        Controller.B2SSetMatch 0
    End If
End Sub

Sub Display_Match()
    Update_Match
    If B2SOn then
        If Match = 0 then
            Controller.B2SSetMatch 100
        else
            Controller.B2SSetMatch Match
        end if
    end if
End Sub

Sub Update_Match
    Select Case Match \ 10
        Case 0:m0.State = 1:m1.State = 0:m2.State = 0:m3.State = 0:m4.State = 0:m5.State = 0:m6.State = 0:m7.State = 0:m8.State = 0:m9.State = 0
        Case 1:m0.State = 0:m1.State = 1:m2.State = 0:m3.State = 0:m4.State = 0:m5.State = 0:m6.State = 0:m7.State = 0:m8.State = 0:m9.State = 0
        Case 2:m0.State = 0:m1.State = 0:m2.State = 1:m3.State = 0:m4.State = 0:m5.State = 0:m6.State = 0:m7.State = 0:m8.State = 0:m9.State = 0
        Case 3:m0.State = 0:m1.State = 0:m2.State = 0:m3.State = 1:m4.State = 0:m5.State = 0:m6.State = 0:m7.State = 0:m8.State = 0:m9.State = 0
        Case 4:m0.State = 0:m1.State = 0:m2.State = 0:m3.State = 0:m4.State = 1:m5.State = 0:m6.State = 0:m7.State = 0:m8.State = 0:m9.State = 0
        Case 5:m0.State = 0:m1.State = 0:m2.State = 0:m3.State = 0:m4.State = 0:m5.State = 1:m6.State = 0:m7.State = 0:m8.State = 0:m9.State = 0
        Case 6:m0.State = 0:m1.State = 0:m2.State = 0:m3.State = 0:m4.State = 0:m5.State = 0:m6.State = 1:m7.State = 0:m8.State = 0:m9.State = 0
        Case 7:m0.State = 0:m1.State = 0:m2.State = 0:m3.State = 0:m4.State = 0:m5.State = 0:m6.State = 0:m7.State = 1:m8.State = 0:m9.State = 0
        Case 8:m0.State = 0:m1.State = 0:m2.State = 0:m3.State = 0:m4.State = 0:m5.State = 0:m6.State = 0:m7.State = 0:m8.State = 1:m9.State = 0
        Case 9:m0.State = 0:m1.State = 0:m2.State = 0:m3.State = 0:m4.State = 0:m5.State = 0:m6.State = 0:m7.State = 0:m8.State = 0:m9.State = 1
    End Select
    'extra 0
    Select Case Match \ 10
        Case 0:m0a.State = 1:m1a.State = 0:m2a.State = 0:m3a.State = 0:m4a.State = 0:m5a.State = 0:m6a.State = 0:m7a.State = 0:m8a.State = 0:m9a.State = 0
        Case 1:m0a.State = 0:m1a.State = 1:m2a.State = 0:m3a.State = 0:m4a.State = 0:m5a.State = 0:m6a.State = 0:m7a.State = 0:m8a.State = 0:m9a.State = 0
        Case 2:m0a.State = 0:m1a.State = 0:m2a.State = 1:m3a.State = 0:m4a.State = 0:m5a.State = 0:m6a.State = 0:m7a.State = 0:m8a.State = 0:m9a.State = 0
        Case 3:m0a.State = 0:m1a.State = 0:m2a.State = 0:m3a.State = 1:m4a.State = 0:m5a.State = 0:m6a.State = 0:m7a.State = 0:m8a.State = 0:m9a.State = 0
        Case 4:m0a.State = 0:m1a.State = 0:m2a.State = 0:m3a.State = 0:m4a.State = 1:m5a.State = 0:m6a.State = 0:m7a.State = 0:m8a.State = 0:m9a.State = 0
        Case 5:m0a.State = 0:m1a.State = 0:m2a.State = 0:m3a.State = 0:m4a.State = 0:m5a.State = 1:m6a.State = 0:m7a.State = 0:m8a.State = 0:m9a.State = 0
        Case 6:m0a.State = 0:m1a.State = 0:m2a.State = 0:m3a.State = 0:m4a.State = 0:m5a.State = 0:m6a.State = 1:m7a.State = 0:m8a.State = 0:m9a.State = 0
        Case 7:m0a.State = 0:m1a.State = 0:m2a.State = 0:m3a.State = 0:m4a.State = 0:m5a.State = 0:m6a.State = 0:m7a.State = 1:m8a.State = 0:m9a.State = 0
        Case 8:m0a.State = 0:m1a.State = 0:m2a.State = 0:m3a.State = 0:m4a.State = 0:m5a.State = 0:m6a.State = 0:m7a.State = 0:m8a.State = 1:m9a.State = 0
        Case 9:m0a.State = 0:m1a.State = 0:m2a.State = 0:m3a.State = 0:m4a.State = 0:m5a.State = 0:m6a.State = 0:m7a.State = 0:m8a.State = 0:m9a.State = 1
    End Select
End Sub

' *********************************************************************
'                      Drain / Plunger Functions
' *********************************************************************

Sub Drain_Hit()
    Drain.DestroyBall
    If bGameInPLay = False Then Exit Sub 'don't do anything, just delete the ball
    BallsOnPlayfield = BallsOnPlayfield - 1
    PlaySoundAt "fx_drain", Drain
    'tilted?
    If Tilted Then
        StopEndOfBallMode
    End If
    ' if still playing and not tilted
    If(bGameInPLay = True) AND(Tilted = False) Then

        ' ballsaver?
        If(bBallSaverActive = True) Then
            CreateNewBall()
        Else
            ' last ball?
            If(BallsOnPlayfield = 0) Then
                StopEndOfBallMode
                vpmtimer.addtimer 800, "EndOfBall '"
                Exit Sub
            End If
        End If
    End If
End Sub

Sub swPlungerRest_Hit()
    bBallInPlungerLane = True
End Sub

Sub swPlungerRest_UnHit()
    bBallInPlungerLane = False
End Sub

' ****************************************
'             Score functions
' ****************************************

Sub AddScore(Points)
    If Tilted Then Exit Sub
    Select Case Points
        Case 1
            Score(CurrentPlayer) = Score(CurrentPlayer) + points * 100
            UpdateScore points * 100
            AddCBonus 1
            ' sounds
            If Points = 1 AND(Score(CurrentPlayer) MOD 10) = 0 Then 'new reel 10
                PlaySound SoundFXDOF("10", 132, DOFPulse, DOFChimes)
            Else
                PlaySound SoundFXDOF(Points, 131, DOFPulse, DOFChimes)
            End If
        Case 10, 100
            Score(CurrentPlayer) = Score(CurrentPlayer) + points * 100
            UpdateScore points * 100
            If Points = 10 AND(Score(CurrentPlayer) MOD 100) \ 10 = 0 Then 'new reel 100
                PlaySound SoundFXDOF("100", 132, DOFPulse, DOFChimes)
            Else
                PlaySound SoundFXDOF(Points, 131, DOFPulse, DOFChimes)
            End If
        Case 5
            Add1 = Add1 + 5
            AddScore1Timer.Enabled = TRUE
        Case 50
            Add10 = Add10 + 5
            AddScore10Timer.Enabled = TRUE
        Case 200 'double bonus
            Add100 = Add100 + 2
            AddScore100Timer.Enabled = TRUE
    End Select

    ' check for higher score and specials
    If Score(CurrentPlayer) >= Special1 AND Special1Awarded(CurrentPlayer) = False Then
        AwardSpecial
        Special1Awarded(CurrentPlayer) = True
    End If
    If Score(CurrentPlayer) >= Special2 AND Special2Awarded(CurrentPlayer) = False Then
        AwardSpecial
        Special2Awarded(CurrentPlayer) = True
    End If
End Sub

'************************************
'       Score sound Timers
'************************************

Sub AddScore1Timer_Timer()
    if Add1> 0 then
        AddScore 1
        Add1 = Add1 - 1
    Else
        Me.Enabled = FALSE
    End If
End Sub

Sub AddScore10Timer_Timer()
    if Add10> 0 then
        AddScore 10
        Add10 = Add10 - 1
    Else
        Me.Enabled = FALSE
    End If
End Sub

Sub AddScore100Timer_Timer()
    if Add100> 0 then
        AddScore 100
        Add100 = Add100 - 1
    Else
        Me.Enabled = FALSE
    End If
End Sub

'*******************
'     BONUS
'*******************

Sub AddBonus(bonuspoints) '10000 points
    If(Tilted = False) Then
        Bonus = Bonus + bonuspoints
        If Bonus> MaxBonus Then Bonus = MaxBonus
        UpdateBonusLights
    End if
End Sub

Sub AddCBonus(bonuspoints) 'spaceship bonus, each 10 will add one 10000 bonus
    If(Tilted = False) Then
        CBonus = CBonus + bonuspoints
        If CBonus> 10 Then
            CBonus = CBonus - 10
            AddBonus 1
        End If
        UpdateBonusLights
    End if
End Sub

'**********************************
'        Score EM reels
'**********************************

Sub UpdateScore(playerpoints)
    Select Case CurrentPlayer
        Case 1:ScoreReel1.Addvalue playerpoints
        Case 2:ScoreReel2.Addvalue playerpoints
        Case 3:ScoreReel3.Addvalue playerpoints
        Case 4:ScoreReel4.Addvalue playerpoints
    End Select
    If B2SOn then
        Controller.B2SSetScorePlayer CurrentPlayer, Score(CurrentPlayer)
    end if
End Sub

Sub ResetScores
    ScoreReel1.ResetToZero
    ScoreReel2.ResetToZero
    ScoreReel3.ResetToZero
    ScoreReel4.ResetToZero
    If B2SOn then
        Controller.B2SSetScorePlayer1 0
        Controller.B2SSetScorePlayer2 0
        Controller.B2SSetScorePlayer3 0
        Controller.B2SSetScorePlayer4 0
    end if
End Sub

Sub AddCredits(value) 'limit to 9 credits
    If Credits <9 Then
        Credits = Credits + value
        CreditsReel.SetValue Credits
        UpdateCredits
    end if
End Sub

Sub UpdateCredits
    PlaySound "fx_relay"
    CreditsReel.SetValue credits
    If B2SOn then
        Controller.B2SSetCredits Credits
    end if
End Sub

Sub UpdateBallInPlay 'update backdrop lights
    'Ball in play
    Select Case Balls
        Case 0:bip1.State = 0:bip2.State = 0:bip3.State = 0:bip4.State = 0:bip5.State = 0
        Case 1:bip1.State = 1:bip2.State = 0:bip3.State = 0:bip4.State = 0:bip5.State = 0
        Case 2:bip1.State = 0:bip2.State = 1:bip3.State = 0:bip4.State = 0:bip5.State = 0
        Case 3:bip1.State = 0:bip2.State = 0:bip3.State = 1:bip4.State = 0:bip5.State = 0
        Case 4:bip1.State = 0:bip2.State = 0:bip3.State = 0:bip4.State = 1:bip5.State = 0
        Case 5:bip1.State = 0:bip2.State = 0:bip3.State = 0:bip4.State = 0:bip5.State = 1
    End Select
    If B2SOn then
        Controller.B2SSetBallInPlay Balls
    End If
    ' Update player light
    Select Case CurrentPlayer
        Case 0:pl1.State = 0:pl2.State = 0:pl3.State = 0:pl4.State = 0
        Case 1:pl1.State = 1:pl2.State = 0:pl3.State = 0:pl4.State = 0
        Case 2:pl1.State = 0:pl2.State = 1:pl3.State = 0:pl4.State = 0
        Case 3:pl1.State = 0:pl2.State = 0:pl3.State = 1:pl4.State = 0
        Case 4:pl1.State = 0:pl2.State = 0:pl3.State = 0:pl4.State = 1
    End Select
    If B2SOn then
        Controller.B2SSetPlayerUp CurrentPlayer
    End If
    ' Player can play
    Select Case PlayersPlayingGame
        Case 0:cp1.State = 0:cp2.State = 0:cp3.State = 0:cp4.State = 0
        Case 1:cp1.State = 1:cp2.State = 0:cp3.State = 0:cp4.State = 0
        Case 2:cp1.State = 0:cp2.State = 1:cp3.State = 0:cp4.State = 0
        Case 3:cp1.State = 0:cp2.State = 0:cp3.State = 1:cp4.State = 0
        Case 4:cp1.State = 0:cp2.State = 0:cp3.State = 0:cp4.State = 1
    End Select
    If B2SOn then
        Controller.B2SSetCanPlay PlayersPlayingGame
    end if
End Sub

'*************************
'        Specials
'*************************

Sub AwardExtraBall()
    If NOT bExtraBallWonThisBall Then
        PlaySound SoundFXDOF("fx_knocker", 122, DOFPulse, DOFknocker)
        ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) + 1
        bExtraBallWonThisBall = True
        LightShootAgain.State = 1
        If B2SOn then
            Controller.B2SSetShootAgain 1
        end if
    END If
End Sub

Sub AwardSpecial()
    PlaySound SoundFXDOF("fx_knocker", 122, DOFPulse, DOFknocker)
    AddCredits 1
End Sub

' ********************************
'        Attract Mode
' ********************************
' use the"Blink Pattern" of each light

Sub StartAttractMode()
    If bGameInPlay Then Exit Sub
    Dim x
    bAttractMode = True
    For each x in aLights
        x.State = 2
    Next
    'turn on GameOver light
    TurnON aGameOver

    'update current player and balls
    bip5.State = 0
End Sub

Sub StopAttractMode()
    Dim x
    bAttractMode = False
    TurnOffPlayfieldLights
    ResetScores
    TurnOFF aGameOver
End Sub

'*********************************
'    Load / Save / Highscore
'*********************************

Sub Loadhs
    Dim x
    x = LoadValue(TableName, "HighScore1")
    If(x <> "") Then HSScore(1) = CDbl(x) Else HSScore(1) = 5000 End If
    x = LoadValue(TableName, "HighScore1Name")
    If(x <> "") Then HSName(1) = x Else HSName(1) = "AAA" End If
    x = LoadValue(TableName, "HighScore2")
    If(x <> "") then HSScore(2) = CDbl(x) Else HSScore(2) = 4500 End If
    x = LoadValue(TableName, "HighScore2Name")
    If(x <> "") then HSName(2) = x Else HSName(2) = "BBB" End If
    x = LoadValue(TableName, "HighScore3")
    If(x <> "") then HSScore(3) = CDbl(x) Else HSScore(3) = 4000 End If
    x = LoadValue(TableName, "HighScore3Name")
    If(x <> "") then HSName(3) = x Else HSName(3) = "CCC" End If
    x = LoadValue(TableName, "HighScore4")
    If(x <> "") then HSScore(4) = CDbl(x) Else HSScore(4) = 3500 End If
    x = LoadValue(TableName, "HighScore4Name")
    If(x <> "") then HSName(4) = x Else HSName(4) = "DDD" End If
    x = LoadValue(TableName, "HighScore5")
    If(x <> "") then HSScore(5) = CDbl(x) Else HSScore(5) = 3000 End If
    x = LoadValue(TableName, "HighScore5Name")
    If(x <> "") then HSName(5) = x Else HSName(5) = "EEE" End If
    x = LoadValue(TableName, "Credits")
    If(x <> "") then Credits = CInt(x) Else Credits = 0
End Sub

Sub Savehs
    SaveValue TableName, "HighScore1", HSScore(1)
    SaveValue TableName, "HighScore1Name", HSName(1)
    SaveValue TableName, "HighScore2", HSScore(2)
    SaveValue TableName, "HighScore2Name", HSName(2)
    SaveValue TableName, "HighScore3", HSScore(3)
    SaveValue TableName, "HighScore3Name", HSName(3)
    SaveValue TableName, "HighScore4", HSScore(4)
    SaveValue TableName, "HighScore4Name", HSName(4)
    SaveValue TableName, "HighScore5", HSScore(5)
    SaveValue TableName, "HighScore5Name", HSName(5)
    SaveValue TableName, "Credits", Credits
End Sub

Sub Reseths
    HSName(1) = "AAA"
    HSName(2) = "BBB"
    HSName(3) = "CCC"
    HSName(4) = "DDD"
    HSName(5) = "EEE"
    HSScore(1) = 5000
    HSScore(2) = 4500
    HSScore(3) = 4000
    HSScore(4) = 3500
    HSScore(5) = 3000
    Savehs
End Sub

Sub SortHighscore
    Dim tmp, tmp2, i, j
    For i = 1 to 5
        For j = 1 to 4
            If HSScore(j) <HSScore(j + 1) Then
                tmp = HSScore(j + 1)
                tmp2 = HSName(j + 1)
                HSScore(j + 1) = HSScore(j)
                HSName(j + 1) = HSName(j)
                HSScore(j) = tmp
                HSName(j) = tmp2
            End If
        Next
    Next
End Sub

' ***************************************************
' GNMOD - Multiple High Score Display and Collection  by GNance
' jpsalas: changed ramps by flashers to remove extra shadow
' ***************************************************

Dim EnteringInitials ' Normally zero, set to non-zero to enter initials
EnteringInitials = False
Dim ScoreChecker
ScoreChecker = 0
Dim CheckAllScores
CheckAllScores = 0
Dim sortscores(4)
Dim sortplayers(4)

Dim PlungerPulled
PlungerPulled = 0

Dim SelectedChar   ' character under the "cursor" when entering initials

Dim HSTimerCount   ' Pass counter For HS timer, scores are cycled by the timer
HSTimerCount = 5   ' Timer is initially enabled, it'll wrap from 5 to 1 when it's displayed

Dim InitialString  ' the string holding the player's initials as they're entered

Dim AlphaString    ' A-Z, 0-9, space (_) and backspace (<)
Dim AlphaStringPos ' pointer to AlphaString, move Forward and backward with flipper keys
AlphaString = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_<"

Dim HSNewHigh      ' The new score to be recorded

Dim HSScore(5)     ' High Scores read in from config file
Dim HSName(5)      ' High Score Initials read in from config file

Sub HighScoreTimer_Timer
    If EnteringInitials then
        If HSTimerCount = 1 then
            SetHSLine 3, InitialString & MID(AlphaString, AlphaStringPos, 1)
            HSTimerCount = 2
        Else
            SetHSLine 3, InitialString
            HSTimerCount = 1
        End If
    ElseIf bGameInPlay then
        SetHSLine 1, "HIGH SCORE1"
        SetHSLine 2, HSScore(1)
        SetHSLine 3, HSName(1)
        HSTimerCount = 5 ' set so the highest score will show after the game is over
        HighScoreTimer.enabled = false
    ElseIf CheckAllScores then
        NewHighScore sortscores(ScoreChecker), sortplayers(ScoreChecker)
    Else
        ' cycle through high scores
        HighScoreTimer.interval = 2000
        HSTimerCount = HSTimerCount + 1
        If HsTimerCount> 5 then
            HSTimerCount = 1
        End If
        SetHSLine 1, "HIGH SCORE" + FormatNumber(HSTimerCount, 0)
        SetHSLine 2, HSScore(HSTimerCount)
        SetHSLine 3, HSName(HSTimerCount)
    End If
End Sub

Function GetHSChar(String, Index)
    Dim ThisChar
    Dim FileName
    ThisChar = Mid(String, Index, 1)
    FileName = "PostIt"
    If ThisChar = " " or ThisChar = "" then
        FileName = FileName & "BL"
    ElseIf ThisChar = "<" then
        FileName = FileName & "LT"
    ElseIf ThisChar = "_" then
        FileName = FileName & "SP"
    Else
        FileName = FileName & ThisChar
    End If
    GetHSChar = FileName
End Function

Sub SetHsLine(LineNo, String)
    Dim Letter
    Dim ThisDigit
    Dim ThisChar
    Dim StrLen
    Dim LetterLine
    Dim Index
    Dim StartHSArray
    Dim EndHSArray
    Dim LetterName
    Dim xFor
    StartHSArray = array(0, 1, 12, 22)
    EndHSArray = array(0, 11, 21, 31)
    StrLen = len(string)
    Index = 1

    For xFor = StartHSArray(LineNo) to EndHSArray(LineNo)
        Eval("HS" &xFor).imageA = GetHSChar(String, Index)
        Index = Index + 1
    Next
End Sub

Sub NewHighScore(NewScore, PlayNum)
    If NewScore> HSScore(5) then
        HighScoreTimer.interval = 500
        HSTimerCount = 1
        AlphaStringPos = 1      ' start with first character "A"
        EnteringInitials = true ' intercept the control keys while entering initials
        InitialString = ""      ' initials entered so far, initialize to empty
        SetHSLine 1, "PLAYER " + FormatNumber(PlayNum, 0)
        SetHSLine 2, "ENTER NAME"
        SetHSLine 3, MID(AlphaString, AlphaStringPos, 1)
        HSNewHigh = NewScore
        'Award Special or simply play the knocker sound
        'AwardSpecial
        PlaySound SoundFXDOF("fx_knocker", 122, DOFPulse, DOFknocker)
    End If
    ScoreChecker = ScoreChecker-1
    If ScoreChecker = 0 then
        CheckAllScores = 0
    End If
End Sub

Sub CollectInitials(keycode)
    Dim i
    If keycode = LeftFlipperKey Then
        ' back up to previous character
        AlphaStringPos = AlphaStringPos - 1
        If AlphaStringPos <1 then
            AlphaStringPos = len(AlphaString) ' handle wrap from beginning to End
            If InitialString = "" then
                ' Skip the backspace If there are no characters to backspace over
                AlphaStringPos = AlphaStringPos - 1
            End If
        End If
        SetHSLine 3, InitialString & MID(AlphaString, AlphaStringPos, 1)
        PlaySound "fx_Previous"
    ElseIf keycode = RightFlipperKey Then
        ' advance to Next character
        AlphaStringPos = AlphaStringPos + 1
        If AlphaStringPos> len(AlphaString) or(AlphaStringPos = len(AlphaString) and InitialString = "") then
            ' Skip the backspace If there are no characters to backspace over
            AlphaStringPos = 1
        End If
        SetHSLine 3, InitialString & MID(AlphaString, AlphaStringPos, 1)
        PlaySound "fx_Next"
    ElseIf keycode = StartGameKey or keycode = PlungerKey Then
        SelectedChar = MID(AlphaString, AlphaStringPos, 1)
        If SelectedChar = "_" then
            InitialString = InitialString & " "
            PlaySound("fx_Esc")
        ElseIf SelectedChar = "<" then
            InitialString = MID(InitialString, 1, len(InitialString) - 1)
            If len(InitialString) = 0 then
                ' If there are no more characters to back over, don't leave the < displayed
                AlphaStringPos = 1
            End If
            PlaySound("fx_Esc")
        Else
            InitialString = InitialString & SelectedChar
            PlaySound("fx_Enter")
        End If
        If len(InitialString) <3 then
            SetHSLine 3, InitialString & SelectedChar
        End If
    End If
    If len(InitialString) = 3 then
        ' save the score
        For i = 5 to 1 step -1
            If i = 1 or(HSNewHigh> HSScore(i) and HSNewHigh <= HSScore(i - 1) ) then
                ' Replace the score at this location
                If i <5 then
                    HSScore(i + 1) = HSScore(i)
                    HSName(i + 1) = HSName(i)
                End If
                EnteringInitials = False
                HSScore(i) = HSNewHigh
                HSName(i) = InitialString
                HSTimerCount = 5
                HighScoreTimer_Timer
                HighScoreTimer.interval = 2000
                Exit Sub
            ElseIf i <5 then
                ' move the score in this slot down by 1, it's been exceeded by the new score
                HSScore(i + 1) = HSScore(i)
                HSName(i + 1) = HSName(i)
            End If
        Next
    End If
End Sub
' End GNMOD

'***********************************************************************
' *********************************************************************
'  *********     G A M E  C O D E  S T A R T S  H E R E      *********
' *********************************************************************
'***********************************************************************

Sub VPObjects_Init 'init objects
    TurnOffPlayfieldLights()
End Sub

' Dim all the variables
Dim DTL(5) 'to know when the targets are hit
Dim DTR(5)
Dim TL(5)
Dim TR(5)
Dim DTLCount  'count how many times the bank has been completed
Dim DTRCount

Sub Game_Init 'called at the start of a new game
    Dim x
    'Start music?
    'Init variables?
    For each x in DTL:DTL(x) = 0:Next
    For each x in DTR:DTR(x) = 0:Next
    For each x in TL:DTL(x) = 0:Next
    For each x in TR:DTR(x) = 0:Next
    'Start or init timers
    'Init lights?
    TurnOffPlayfieldLights
End Sub

Sub StopEndOfBallMode     'called when the last ball is drained
End Sub

Sub ResetNewBallVariables 'init variables & lights new ball/player
    TurnOffPlayfieldLights
    Bonus = 0
    CBonus = 1
    UpdateBonusLights
    BonusMultiplier = 1
    ResetBanks
    DTL(0) = 0
    DTR(0) = 0
    DTLCount = 0
    DTRCount = 0
    PlaySound "ReelInitt"
' turn on new ball Lights

End Sub

Sub TurnOffPlayfieldLights
    Dim a
    For each a in aLights
        a.State = 0
    Next
End Sub

' *********************************************************************
'                        Table Object Hit Events
' *********************************************************************
' Any target hit Sub will do this:
' - play a sound
' - do some physical movement
' - add a score, bonus
' - check some variables/modes this trigger is a member of
' - set the "LastSwicthHit" variable in case it is needed later
' *********************************************************************

'***********
' Bumpers
'***********

Sub Bumper001_Hit 'left bumper
    If Tilted Then Exit Sub
    PlaySoundAt SoundFXDOF("fx_bumper", 109, DOFPulse, DOFContactors), Bumper001
    AddScore 1
End Sub

Sub Bumper002_Hit 'right
    If Tilted Then Exit Sub
    PlaySoundAt SoundFXDOF("fx_bumper", 110, DOFPulse, DOFContactors), Bumper002
    AddScore 1
End Sub

'***************
'  Slingshots
'***************
Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    If Tilted Then Exit Sub
    'PlaySoundAt "fx_slingshot", Lemk
    PlaySoundAt SoundFXDOF("fx_slingshot", 103, DOFPulse, DOFcontactors), Lemk
    LeftSling004.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
    ' add points
    AddScore 50
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
    If Tilted Then Exit Sub
    'PlaySoundAt "fx_slingshot", Remk
    PlaySoundAt SoundFXDOF("fx_slingshot", 104, DOFPulse, DOFcontactors), Remk
    RightSling004.Visible = 1
    Remk.RotX = 26
    RStep = 0
    RightSlingShot.TimerEnabled = 1
    ' add points
    AddScore 50
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 1:RightSLing004.Visible = 0:RightSLing003.Visible = 1:Remk.RotX = 14
        Case 2:RightSLing003.Visible = 0:RightSLing002.Visible = 1:Remk.RotX = 2
        Case 3:RightSLing002.Visible = 0:Remk.RotX = -20:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

' 100 point rebounds

Sub rsband003_Hit:AddScore 1:End Sub
Sub rsband008_Hit:AddScore 1:End Sub

' Top Lanes

Sub Trigger009_Hit 'left
    PlaySoundAt "fx_sensor", Trigger009
    If Tilted Then Exit Sub
    If Light015.State Then
        AddScore 50
        AddBonus 1
    Else
        Addscore 5
    End If
    If Light013.State Then AwardExtraBall:Light013.State = 0
End Sub

Sub Trigger010_Hit 'right
    PlaySoundAt "fx_sensor", Trigger010
    If Tilted Then Exit Sub
    If Light016.State Then
        AddScore 50
        AddBonus 1
    Else
        Addscore 5
    End If
    If Light014.State Then AwardExtraBall:Light014.State = 0
End Sub


' Outlanes

Sub Trigger001_Hit 'left outlane
    PlaySoundAt "fx_sensor", Trigger001
    If Tilted Then Exit Sub
    If Light001.State Then
        AwardSpecial
    Else
        AddScore 50
    End If
End Sub

Sub Trigger006_Hit 'right outlane
    PlaySoundAt "fx_sensor", Trigger006
    If Tilted Then Exit Sub
    If Light002.State Then
        AwardSpecial
    Else
        AddScore 50
    End If
End Sub

' Inlanes

Sub Trigger002_Hit
    PlaySoundAt "fx_sensor", Trigger002
    If Tilted Then Exit Sub
    If Light003.State Then
        AddScore 50
        AddBonus 1
    Else
        Addscore 5
    End If
End Sub

Sub Trigger003_Hit
    PlaySoundAt "fx_sensor", Trigger003
    If Tilted Then Exit Sub
    If Light004.State Then
        AddScore 50
        AddBonus 1
    Else
        Addscore 5
    End If
End Sub

Sub Trigger004_Hit
    PlaySoundAt "fx_sensor", Trigger004
    If Tilted Then Exit Sub
    If Light005.State Then
        AddScore 50
        AddBonus 1
    Else
        Addscore 5
    End If
End Sub

Sub Trigger005_Hit
    PlaySoundAt "fx_sensor", Trigger005
    If Tilted Then Exit Sub
    If Light006.State Then
        AddScore 50
        AddBonus 1
    Else
        Addscore 5
    End If
End Sub

' side lane trigger

Sub Trigger007_Hit 'right
    PlaySoundAt "fx_sensor", Trigger007
    If Tilted Then Exit Sub
    If Light008.State Then
        AddScore 50
        AddBonus 1
    Else
        Addscore 5
    End If
End Sub

Sub Trigger008_Hit 'left
    PlaySoundAt "fx_sensor", Trigger008
    If Tilted Then Exit Sub
    If Light007.State Then
        AddScore 50
        AddBonus 1
    Else
        Addscore 5
    End If
End Sub

' Droptargets

' Left bank
Sub dt1_Hit:PlaySoundAt SoundFXDOF("fx_droptarget", 105, DOFPulse, DOFContactors), dt1:End Sub
Sub dt2_Hit:PlaySoundAt SoundFXDOF("fx_droptarget", 105, DOFPulse, DOFContactors), dt2:End Sub
Sub dt3_Hit:PlaySoundAt SoundFXDOF("fx_droptarget", 106, DOFPulse, DOFContactors), dt3:End Sub
Sub dt4_Hit:PlaySoundAt SoundFXDOF("fx_droptarget", 106, DOFPulse, DOFContactors), dt4:End Sub
Sub dt5_Hit:PlaySoundAt SoundFXDOF("fx_droptarget", 107, DOFPulse, DOFContactors), dt5:End Sub

Sub dt1_Dropped
    If Tilted Then Exit Sub
    AddScore 50
    DTL(1) = 1
    Light004.State = 1
    CheckLBank
End Sub

Sub dt2_Dropped
    If Tilted Then Exit Sub
    AddScore 50
    DTL(2) = 1
    Light003.State = 1
    CheckLBank
End Sub

Sub dt3_Dropped
    If Tilted Then Exit Sub
    AddScore 50
    DTL(3) = 1
    Light007.State = 1
    CheckLBank
End Sub

Sub dt4_Dropped
    If Tilted Then Exit Sub
    AddScore 50
    DTL(4) = 1
    Light011.State = 1
    CheckLBank
End Sub

Sub dt5_Dropped
    If Tilted Then Exit Sub
    AddScore 50
    DTL(5) = 1
    Light015.State = 1
    CheckLBank
End Sub

Sub CheckLBank
    'use the DTL(0) as the complete status
    DTL(0) = DTL(1) + DTL(2) + DTL(3) + DTL(4) + DTL(5)
    If DTL(0) = 5 Then
        DTLCount = DTLCount + 1
        If BallsPerGame = 3 Then
            Select Case DTLCount
                Case 1:Light010.State = 1 'double bonus
                Case 2:Light014.State = 1 'extra ball light
                Case 3:Light002.State = 1 'special light
            End Select
        Else                              '5 balls
            Select case DTLCount
                Case 1:vpmTimer.AddTimer 250, "ResetLBank '"
                Case 2:Light010.State = 1 'double bonus
                Case 3:Light014.State = 1 'extra ball light
                Case 4:Light002.State = 1 'special light
            End Select
        End If
    End If
End Sub

'right bank
Sub dt6_Hit:PlaySoundAt SoundFXDOF("fx_droptarget", 105, DOFPulse, DOFContactors), dt6:End Sub
Sub dt7_Hit:PlaySoundAt SoundFXDOF("fx_droptarget", 105, DOFPulse, DOFContactors), dt7:End Sub
Sub dt8_Hit:PlaySoundAt SoundFXDOF("fx_droptarget", 106, DOFPulse, DOFContactors), dt8:End Sub
Sub dt9_Hit:PlaySoundAt SoundFXDOF("fx_droptarget", 106, DOFPulse, DOFContactors), dt9:End Sub
Sub dt10_Hit:PlaySoundAt SoundFXDOF("fx_droptarget", 107, DOFPulse, DOFContactors), dt10:End Sub

Sub dt6_Dropped
    If Tilted Then Exit Sub
    AddScore 50
    DTR(1) = 1
    Light005.State = 1
    CheckRBank
End Sub

Sub dt7_Dropped
    If Tilted Then Exit Sub
    AddScore 50
    DTR(2) = 1
    Light006.State = 1
    CheckRBank
End Sub

Sub dt8_Dropped
    If Tilted Then Exit Sub
    AddScore 50
    DTR(3) = 1
    Light008.State = 1
    CheckRBank
End Sub

Sub dt9_Dropped
    If Tilted Then Exit Sub
    AddScore 50
    DTR(4) = 1
    Light012.State = 1
    CheckRBank
End Sub

Sub dt10_Dropped
    If Tilted Then Exit Sub
    AddScore 50
    DTR(5) = 1
    Light016.State = 1
    CheckRBank
End Sub

Sub CheckRBank
    'use the DTR(0) as the complete status
    DTR(0) = DTR(1) + DTR(2) + DTR(3) + DTR(4) + DTR(5)
    If DTR(0) = 5 Then
        DTRCount = DTRCount + 1
        If BallsPerGame = 3 Then
            Select Case DTRCount
                Case 1:Light010.State = 1 'double bonus
                Case 2:Light014.State = 1 'extra ball light
                Case 3:Light002.State = 1 'special light
            End Select
        Else                              '5 balls
            Select case DTRCount
                Case 1:vpmTimer.AddTimer 250, "ResetRBank '"
                Case 2:Light010.State = 1 'double bonus
                Case 3:Light014.State = 1 'extra ball light
                Case 4:Light002.State = 1 'special light
            End Select
        End If
    End If
End Sub

Sub ResetLBank
    PlaySoundAt SoundFXDOF("fx_resetdrop", 113, DOFPulse, DOFContactors), dt3
    For x = 1 to 5:DTL(x) = 0:Next
    For each x in aDTLLights: x.State = 0: Next
    dt1.IsDropped = 0
    dt2.IsDropped = 0
    dt3.IsDropped = 0
    dt4.IsDropped = 0
    dt5.IsDropped = 0
End Sub

Sub ResetRBank
    PlaySoundAt SoundFXDOF("fx_resetdrop", 114, DOFPulse, DOFContactors), dt8
    For x = 1 to 5:DTR(x) = 0:Next
    For each x in aDTRLights: x.State = 0: Next
    dt6.IsDropped = 0
    dt7.IsDropped = 0
    dt8.IsDropped = 0
    dt9.IsDropped = 0
    dt10.IsDropped = 0
End Sub

Sub ResetBanks
    vpmTimer.AddTimer 250, "ResetLBank '"
    vpmTimer.AddTimer 450, "ResetRBank '"
End Sub

' Targets

Sub T11_Hit 'top left
    PlaySoundAtBall SoundFXDOF("fx_target", 106, DOFPulse, DOFTargets)
    If Tilted Then Exit Sub
    If Light011.State Then
        AddScore 50
        AddBonus 1
    Else
        Addscore 5
    End If
End Sub

Sub T12_Hit 'top right
    PlaySoundAtBall SoundFXDOF("fx_target", 108, DOFPulse, DOFTargets)
    If Tilted Then Exit Sub
    If Light012.State Then
        AddScore 50
        AddBonus 1
    Else
        Addscore 5
    End If
End Sub

' left bank
Sub t1_Hit
    PlaySoundAtBall SoundFXDOF("fx_target", 105, DOFPulse, DOFTargets)
    If Tilted Then Exit Sub
    Addscore 5
End Sub

Sub t2_Hit
    PlaySoundAtBall SoundFXDOF("fx_target", 105, DOFPulse, DOFTargets)
    If Tilted Then Exit Sub
    Addscore 5
End Sub

Sub t3_Hit
    PlaySoundAtBall SoundFXDOF("fx_target", 106, DOFPulse, DOFTargets)
    If Tilted Then Exit Sub
    Addscore 5
    If Light009.State Then
        BonusMultiplier = 2
        DBonusLight.State = 1
        ResetBanks
    End If
End Sub

Sub t4_Hit
    PlaySoundAtBall SoundFXDOF("fx_target", 105, DOFPulse, DOFTargets)
    If Tilted Then Exit Sub
    Addscore 5
End Sub

Sub t5_Hit
    PlaySoundAtBall SoundFXDOF("fx_target", 105, DOFPulse, DOFTargets)
    If Tilted Then Exit Sub
    Addscore 5
End Sub

'right bank
Sub t6_Hit
    PlaySoundAtBall SoundFXDOF("fx_target", 107, DOFPulse, DOFTargets)
    If Tilted Then Exit Sub
    Addscore 5
End Sub

Sub t7_Hit
    PlaySoundAtBall SoundFXDOF("fx_target", 107, DOFPulse, DOFTargets)
    If Tilted Then Exit Sub
    Addscore 5
End Sub

Sub t8_Hit
    PlaySoundAtBall SoundFXDOF("fx_target", 108, DOFPulse, DOFTargets)
    If Tilted Then Exit Sub
    Addscore 5
    If Light010.State Then
        BonusMultiplier = 2
        DBonusLight.State = 1
        ResetBanks
    End If
End Sub

Sub t9_Hit
    PlaySoundAtBall SoundFXDOF("fx_target", 107, DOFPulse, DOFTargets)
    If Tilted Then Exit Sub
    Addscore 5
End Sub

Sub t10_Hit
    PlaySoundAtBall SoundFXDOF("fx_target", 107, DOFPulse, DOFTargets)
    If Tilted Then Exit Sub
    Addscore 5
End Sub


'*********************************
' Table Options F12 User Options
'*********************************
' Table1.Option arguments are:
' - option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional array of literal strings

Dim LUTImage, BallsPerGame

Sub Table1_OptionEvent(ByVal eventId)
    Dim x, y

    'LUT
    LutImage = Table1.Option("Select LUT", 0, 21, 1, 0, 0, Array("Normal 0", "Normal 1", "Normal 2", "Normal 3", "Normal 4", "Normal 5", "Normal 6", "Normal 7", "Normal 8", "Normal 9", "Normal 10", _
        "Warm 0", "Warm 1", "Warm 2", "Warm 3", "Warm 4", "Warm 5", "Warm 6", "Warm 7", "Warm 8", "Warm 9", "Warm 10") )
    UpdateLUT

    ' Cabinet rails
    x = Table1.Option("Cabinet Rails", 0, 1, 1, 1, 0, Array("Hide", "Show") )
    For each y in aRails:y.visible = x:next

    ' Balls per Game
    x = Table1.Option("Balls per Game", 0, 1, 1, 1, 0, Array("3 Balls", "5 Balls") )
    If x = 1 Then BallsPerGame = 5 Else BallsPerGame = 3

    If BallsPerGame = 5 Then
        Special1 = 600000
        Special2 = 720000
    Else
        Special1 = 400000
        Special2 = 520000
    End if
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

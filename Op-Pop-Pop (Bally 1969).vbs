' Op-Pop-Pop / IPD No. 1722 / April, 1969 / 1 Player
' Bally Manufacturing Corporation (1931-1983) [Trade Name: Bally]
' VPX8 table by jpsalas, 2025

Option Explicit ' Force explicit variable declaration
Randomize

' core.vbs constants
Const BallSize = 50 ' 50 is the normal size
Const BallMass = 1  ' 1 is the normal ball mass.

'*******************  P L A Y E R   O P T I O N S  **************************

Const Replay = True ' set it to False to play the Add-A-Ball version

'****************************************************************************


'DOF config - (a config must be created in the server)
'101 Left Flipper
'102 Right Flipper
'103 Left Slingshot
'104 Right Slingshot
'105 left mushroom
'106 center mushroom
'107 right mushroom
'108 kicker
'109 Left Top mushroom
'110 Right Top mushroom
'111
'112
'113
'114
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
'131 Chime - Bell
'132 Chime - Bell
'133 Chime - Bell

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
Const TableName = "op-pop-pop" ' file name to save highscores and other variables
Const cGameName = "op-pop-pop" ' B2S name & DOF config
Const MaxPlayers = 1           ' 1 to 4 can play
Const MaxMultiplier = 1        ' limit bonus multiplier
Const MaxBonus = 10             ' highest bonus count
Const BallsPerGame = 5         ' 3 or 5
Const FreePlay = False         ' Free play or coins

' Global variables
Dim PlayersPlayingGame
Dim CurrentPlayer
Dim Credits
Dim Bonus
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
Dim Add1000
Dim LastSwicthHit

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
    Add1000 = 0
    LastSwicthHit = ""

    If BallsPerGame = 5 Then
        Special1 = 3200
        Special2 = 4000
        Special3 = 4800
        Special4 = 5400
    Else
        Special1 = 3000
        Special2 = 3800
        Special3 = 4600
        Special4 = 5200
    End if

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

    ' Load table color
    LoadLut
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

    If keycode = LeftMagnaSave Then bLutActive = True:SetLUTLine "Color LUT image " & table1.ColorGradeImage
    If keycode = RightMagnaSave AND bLutActive Then NextLUT:End If

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
    If keycode = LeftTiltKey Then Nudge 90, 7:PlaySound "fx_nudge", 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 7:PlaySound "fx_nudge", 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 8:PlaySound "fx_nudge", 0, 1, 1, 0.25

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
End Sub

Sub Table1_KeyUp(ByVal keycode)

    If EnteringInitials then
        Exit Sub
    End If

    If keycode = LeftMagnaSave Then bLutActive = False:HideLUT

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
        LeftFlipper1.RotateToEnd
        LeftFlipperOn = 1
    Else
        PlaySoundAt SoundFXDOF("fx_flipperdown", 101, DOFOff, DOFFlippers), LeftFlipper
        LeftFlipper.RotateToStart
        LeftFlipper1.RotateToStart
        LeftFlipperOn = 0
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFXDOF("fx_flipperup", 102, DOFOn, DOFFlippers), RightFlipper
        RightFlipper.RotateToEnd
        RightFlipper1.RotateToEnd
        RightFlipperOn = 1
    Else
        PlaySoundAt SoundFXDOF("fx_flipperdown", 102, DOFOff, DOFFlippers), RightFlipper
        RightFlipper.RotateToStart
        RightFlipper1.RotateToStart
        RightFlipperOn = 0
    End If
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, Vol(ActiveBall), pan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, Vol(ActiveBall), pan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper1_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, Vol(ActiveBall), pan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
End Sub
Sub LeftFlipper1_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, Vol(ActiveBall), pan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub LeftFlipper_Animate
    LeftFlipperP.Rotz = LeftFlipper.CurrentAngle
    LeftFlipperRubber.Rotz = LeftFlipper.CurrentAngle
End Sub

Sub RightFlipper_Animate
    RightFlipperP.Rotz = RightFlipper.CurrentAngle
    RightFlipperRubber.Rotz = RightFlipper.CurrentAngle
End Sub

Sub LeftFlipper1_Animate
    LeftFlipperP1.Rotz = LeftFlipper1.CurrentAngle
    LeftFlipperRubber1.Rotz = LeftFlipper1.CurrentAngle
End Sub

Sub RightFlipper1_Animate
    RightFlipperP1.Rotz = RightFlipper1.CurrentAngle
    RightFlipperRubber1.Rotz = RightFlipper1.CurrentAngle
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

FlipperPower = 5000
FlipperElasticity = 0.7
FullStrokeEOS_Torque = 0.3 ' EOS Torque when flipper hold up ( EOS Coil is fully charged. Ampere increase due to flipper can't move or when it pushed back when "On". EOS Coil have more power )
LiveStrokeEOS_Torque = 0.2 ' EOS Torque when flipper rotate to end ( When flipper move, EOS coil have less Ampere due to flipper can freely move. EOS Coil have less power )

LeftFlipper.EOSTorqueAngle = 10
RightFlipper.EOSTorqueAngle = 10

SOSTorque = 0.1
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

Dim GiIntensity
GiIntensity = 1               'can be used by the LUT changing to increase the GI lights when the table is darker

Sub ChangeGiIntensity(factor) 'changes the intensity scale
    Dim bulb
    For each bulb in aGiLights
        bulb.IntensityScale = GiIntensity * factor
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
        LeftSlingshot.Disabled = 1
        RightSlingshot.Disabled = 1
    Else
        GiOn
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

Const tnob = 20   'total number of balls
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

    If BonusEnabled Then
        AddScore BonusValue(Bonus)
    End If
    vpmtimer.addtimer 1000, "EndOfBall2 '"
End Sub

Sub UpdateBonusLights
    Select Case Bonus
        Case 0:bl1.State = 0:bl2.State = 0:bl3.State = 0:bl4.State = 0:bl5.State = 0:bl6.State = 0:bl7.State = 0:bl8.State = 0:bl9.State = 0:blx.State = 0
        Case 1:bl1.State = 1:bl2.State = 0:bl3.State = 0:bl4.State = 0:bl5.State = 0:bl6.State = 0:bl7.State = 0:bl8.State = 0:bl9.State = 0:blx.State = 0
        Case 2:bl1.State = 0:bl2.State = 1:bl3.State = 0:bl4.State = 0:bl5.State = 0:bl6.State = 0:bl7.State = 0:bl8.State = 0:bl9.State = 0:blx.State = 0
        Case 3:bl1.State = 0:bl2.State = 0:bl3.State = 1:bl4.State = 0:bl5.State = 0:bl6.State = 0:bl7.State = 0:bl8.State = 0:bl9.State = 0:blx.State = 0
        Case 4:bl1.State = 0:bl2.State = 0:bl3.State = 0:bl4.State = 1:bl5.State = 0:bl6.State = 0:bl7.State = 0:bl8.State = 0:bl9.State = 0:blx.State = 0
        Case 5:bl1.State = 0:bl2.State = 0:bl3.State = 0:bl4.State = 0:bl5.State = 1:bl6.State = 0:bl7.State = 0:bl8.State = 0:bl9.State = 0:blx.State = 0
        Case 6:bl1.State = 0:bl2.State = 0:bl3.State = 0:bl4.State = 0:bl5.State = 0:bl6.State = 1:bl7.State = 0:bl8.State = 0:bl9.State = 0:blx.State = 0
        Case 7:bl1.State = 0:bl2.State = 0:bl3.State = 0:bl4.State = 0:bl5.State = 0:bl6.State = 0:bl7.State = 1:bl8.State = 0:bl9.State = 0:blx.State = 0
        Case 8:bl1.State = 0:bl2.State = 0:bl3.State = 0:bl4.State = 0:bl5.State = 0:bl6.State = 0:bl7.State = 0:bl8.State = 1:bl9.State = 0:blx.State = 0
        Case 9:bl1.State = 0:bl2.State = 0:bl3.State = 0:bl4.State = 0:bl5.State = 0:bl6.State = 0:bl7.State = 0:bl8.State = 0:bl9.State = 1:blx.State = 0
        Case 10:bl1.State = 0:bl2.State = 0:bl3.State = 0:bl4.State = 0:bl5.State = 0:bl6.State = 0:bl7.State = 0:bl8.State = 0:bl9.State = 0:blx.State = 1:Light002.State = 1:Light003.State = 1
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
    DOF 250, DOFPulse

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
    PlaySound "fx_match"
    Match = INT(RND(1) * 10) ' random between 0 and 9
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
    If B2SOn then
        Controller.B2SSetMatch 0
    End If
End Sub

Sub Display_Match()
    Update_Match
    If B2SOn then
        If Match = 0 then
            Controller.B2SSetMatch 10
        else
            Controller.B2SSetMatch Match
        end if
    end if
End Sub

Sub Update_Match
    Select Case Match
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
End Sub

' *********************************************************************
'                      Drain / Plunger Functions
' *********************************************************************

Sub Drain_Hit()
    Drain.DestroyBall
    If bGameInPLay = False Then Exit Sub 'don't do anything, just delete the ball
    BallsOnPlayfield = BallsOnPlayfield - 1
    PlaySoundAt "fx_drain", Drain
    DOF 124, 2
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
            Score(CurrentPlayer) = Score(CurrentPlayer) + points
            UpdateScore points
            ' sounds
            PlaySound SoundFXDOF("10", 131, DOFPulse, DOFChimes)
        Case 10
            Score(CurrentPlayer) = Score(CurrentPlayer) + points
            UpdateScore points
            ' sounds
            PlaySound SoundFXDOF("10", 132, DOFPulse, DOFChimes)
        Case 100
            Score(CurrentPlayer) = Score(CurrentPlayer) + points
            UpdateScore points
            ' sounds
            PlaySound SoundFXDOF("100", 133, DOFPulse, DOFChimes)
        Case 2, 3, 4, 5, 6, 7, 8, 9
            Add1 = Add1 + Points
            AddScore1Timer.Enabled = TRUE
        Case 20, 30, 40, 50, 60, 70, 80, 90
            Add10 = Add10 + Points \ 10
            AddScore10Timer.Enabled = TRUE
        Case 200, 300, 400, 500, 600, 700, 800, 900
            Add100 = Add100 + Points \ 100
            AddScore100Timer.Enabled = TRUE
    End Select

    ' check for higher score and specials
    If Score(CurrentPlayer) >= Special1 AND Special1Awarded(CurrentPlayer) = False Then
        If Replay Then
            AwardSpecial
        Else
            AwardAddaBall
        End If
        Special1Awarded(CurrentPlayer) = True
    End If
    If Score(CurrentPlayer) >= Special2 AND Special2Awarded(CurrentPlayer) = False Then
        If Replay Then
            AwardSpecial
        Else
            AwardAddaBall
        End If
        Special2Awarded(CurrentPlayer) = True
    End If
    If Score(CurrentPlayer) >= Special3 AND Special3Awarded(CurrentPlayer) = False Then
        If Replay Then
            AwardSpecial
        Else
            AwardAddaBall
        End If
        Special3Awarded(CurrentPlayer) = True
    End If
    If Score(CurrentPlayer) >= Special4 AND Special4Awarded(CurrentPlayer) = False Then
        If Replay Then
            AwardSpecial
        Else
            AwardAddaBall
        End If
        Special4Awarded(CurrentPlayer) = True
    End If
End Sub

'************************************
'       Score sound Timers
'************************************

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

Sub AddScore1Timer_Timer()
    if Add1> 0 then
        AddScore 1
        Add1 = Add1 - 1
    Else
        Me.Enabled = FALSE
    End If
End Sub

'*******************
'     BONUS
'*******************

Sub AddBonus(bonuspoints)
    If(Tilted = False) Then
        Bonus = Bonus + bonuspoints
        If Bonus> MaxBonus Then Bonus = MaxBonus
        'PlaySound "Bell"
        UpdateBonusLights
    End if
End Sub

Sub AddBonusMultiplier(multi)
    If(Tilted = False) Then
        BonusMultiplier = BonusMultiplier + multi
        If BonusMultiplier> MaxMultiplier Then
            BonusMultiplier = MaxMultiplier
        End If
        UpdateMultiplierLights
    End if
End Sub

Sub UpdateMultiplierLights
    Select Case BonusMultiplier
        Case 1:bmli2.State = 0:bmli3.State = 0
        Case 2:bmli2.State = 1:bmli3.State = 0
        Case 3:bmli2.State = 0:bmli3.State = 1
    End Select
End Sub

'**********************************
'        Score EM reels
'**********************************

Sub UpdateScore(playerpoints)
    ScoreReel1.Addvalue playerpoints
    If B2SOn then
        Controller.B2SSetScorePlayer CurrentPlayer, Score(CurrentPlayer)
    end if
End Sub

Sub ResetScores
    ScoreReel1.ResetToZero
    If B2SOn then
        Controller.B2SSetScorePlayer1 0
    end if
End Sub

Sub AddCredits(value) 'limit to 25 credits
    If Credits <25 Then
        Credits = Credits + value
        CreditsReel.SetValue Credits
        UpdateCredits
        DOF 200, DOFOn
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
If Replay Then
    Select Case Balls
        Case 0:bip1.State = 0:bip2.State = 0:bip3.State = 0:bip4.State = 0:bip5.State = 0:bip6.State = 0:bip7.State = 0
        Case 1:bip1.State = 1:bip2.State = 0:bip3.State = 0:bip4.State = 0:bip5.State = 0:bip6.State = 0:bip7.State = 0
        Case 2:bip1.State = 0:bip2.State = 1:bip3.State = 0:bip4.State = 0:bip5.State = 0:bip6.State = 0:bip7.State = 0
        Case 3:bip1.State = 0:bip2.State = 0:bip3.State = 1:bip4.State = 0:bip5.State = 0:bip6.State = 0:bip7.State = 0
        Case 4:bip1.State = 0:bip2.State = 0:bip3.State = 0:bip4.State = 1:bip5.State = 0:bip6.State = 0:bip7.State = 0
        Case 5:bip1.State = 0:bip2.State = 0:bip3.State = 0:bip4.State = 0:bip5.State = 1:bip6.State = 0:bip7.State = 0
        Case 6:bip1.State = 0:bip2.State = 0:bip3.State = 0:bip4.State = 0:bip5.State = 0:bip6.State = 1:bip7.State = 0
        Case 7:bip1.State = 0:bip2.State = 0:bip3.State = 0:bip4.State = 0:bip5.State = 0:bip6.State = 0:bip7.State = 1
    End Select
    If B2SOn then
        Controller.B2SSetBallInPlay Balls
    End If
Else 'Add-a-ball
Dim tmp
tmp = 1 + BallsPerGame - Balls
    Select Case tmp
        Case 0:bip1.State = 0:bip2.State = 0:bip3.State = 0:bip4.State = 0:bip5.State = 0:bip6.State = 0:bip7.State = 0
        Case 1:bip1.State = 1:bip2.State = 0:bip3.State = 0:bip4.State = 0:bip5.State = 0:bip6.State = 0:bip7.State = 0
        Case 2:bip1.State = 0:bip2.State = 1:bip3.State = 0:bip4.State = 0:bip5.State = 0:bip6.State = 0:bip7.State = 0
        Case 3:bip1.State = 0:bip2.State = 0:bip3.State = 1:bip4.State = 0:bip5.State = 0:bip6.State = 0:bip7.State = 0
        Case 4:bip1.State = 0:bip2.State = 0:bip3.State = 0:bip4.State = 1:bip5.State = 0:bip6.State = 0:bip7.State = 0
        Case 5:bip1.State = 0:bip2.State = 0:bip3.State = 0:bip4.State = 0:bip5.State = 1:bip6.State = 0:bip7.State = 0
        Case 6:bip1.State = 0:bip2.State = 0:bip3.State = 0:bip4.State = 0:bip5.State = 0:bip6.State = 1:bip7.State = 0
        Case 7:bip1.State = 0:bip2.State = 0:bip3.State = 0:bip4.State = 0:bip5.State = 0:bip6.State = 0:bip7.State = 1
    End Select
    If B2SOn then
        Controller.B2SSetBallInPlay tmp
    End If

End If
' Update player light

' Player can play

End Sub

'*************************
'        Specials
'*************************

Sub AwardExtraBall()
    If NOT bExtraBallWonThisBall Then
        PlaySound SoundFXDOF("fx_knocker", 122, DOFPulse, DOFknocker)
        DOF 230, DOFPulse
        ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) + 1
        bExtraBallWonThisBall = True
        LightShootAgain.State = 1
        If B2SOn then
            Controller.B2SSetShootAgain 1
        end if
    Else
        Addscore 5000
    END If
End Sub

Sub AwardSpecial()
    PlaySound SoundFXDOF("fx_knocker", 122, DOFPulse, DOFknocker)
    DOF 230, DOFPulse
    DOF 200, DOFOn
    AddCredits 1
End Sub

Sub AwardAddaBall()
    If BallsRemaining(CurrentPlayer) <11 Then
        PlaySound SoundFXDOF("fx_knocker", 122, DOFPulse, DOFknocker)
        BallsRemaining(CurrentPlayer) = BallsRemaining(CurrentPlayer) + 1
        If BallsRemaining(CurrentPlayer)> 7 Then BallsRemaining(CurrentPlayer) = 7
        UpdateBallInPlay
    End If
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
    ' Based on Black's Highscore routines
    Dim FileObj
    Dim ScoreFile, TextStr
    Dim temp1
    Dim temp2
    Dim temp3
    Dim temp4
    Dim temp5
    Dim temp6
    Dim temp8
    Dim temp9
    Dim temp10
    Dim temp11
    Dim temp12
    Dim temp13
    Dim temp14
    Dim temp15
    Dim temp16
    Dim temp17

    Set FileObj = CreateObject("Scripting.FileSystemObject")
    If Not FileObj.FolderExists(UserDirectory) then
        Credits = 0
        Exit Sub
    End If
    If Not FileObj.FileExists(UserDirectory & TableName& ".txt") then
        Credits = 0
        Exit Sub
    End If
    Set ScoreFile = FileObj.GetFile(UserDirectory & TableName& ".txt")
    Set TextStr = ScoreFile.OpenAsTextStream(1, 0)
    If(TextStr.AtEndOfStream = True) then
        Exit Sub
    End If
    temp1 = TextStr.ReadLine
    temp2 = textstr.readline

    HighScore = cdbl(temp1)
    If HighScore <1 then
        temp8 = textstr.readline
        temp9 = textstr.readline
        temp10 = textstr.readline
        temp11 = textstr.readline
        temp12 = textstr.readline
        temp13 = textstr.readline
        temp14 = textstr.readline
        temp15 = textstr.readline
        temp16 = textstr.readline
        temp17 = textstr.readline
    End If
    TextStr.Close
    Credits = cdbl(temp2)

    If HighScore <1 then
        HSScore(1) = int(temp8)
        HSScore(2) = int(temp9)
        HSScore(3) = int(temp10)
        HSScore(4) = int(temp11)
        HSScore(5) = int(temp12)

        HSName(1) = temp13
        HSName(2) = temp14
        HSName(3) = temp15
        HSName(4) = temp16
        HSName(5) = temp17
    End If
    Set ScoreFile = Nothing
    Set FileObj = Nothing
    SortHighscore 'added to fix a previous error
End Sub

Sub Savehs
    ' Based on Black's Highscore routines
    Dim FileObj
    Dim ScoreFile
    Dim xx
    Set FileObj = CreateObject("Scripting.FileSystemObject")
    If Not FileObj.FolderExists(UserDirectory) then
        Exit Sub
    End If
    Set ScoreFile = FileObj.CreateTextFile(UserDirectory & TableName& ".txt", True)
    ScoreFile.WriteLine 0
    ScoreFile.WriteLine Credits
    For xx = 1 to 5
        scorefile.writeline HSScore(xx)
    Next
    For xx = 1 to 5
        scorefile.writeline HSName(xx)
    Next
    ScoreFile.Close
    Set ScoreFile = Nothing
    Set FileObj = Nothing
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

' default high scores
HSScore(1) = 3000
HSScore(2) = 2900
HSScore(3) = 2800
HSScore(4) = 2700
HSScore(5) = 2600

HSName(1) = "AAA"
HSName(2) = "ZZZ"
HSName(3) = "XXX"
HSName(4) = "ABC"
HSName(5) = "BBB"

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
        DOF 230, DOFPulse
        DOF 200, DOFOn
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

'************************************
'       LUT - Darkness control
' 10 normal level & 10 warmer levels
'************************************

Dim bLutActive, LUTImage

Sub LoadLUT
    Dim x
    bLutActive = False
    x = LoadValue(cGameName, "LUTImage")
    If(x <> "") Then LUTImage = x Else LUTImage = 0
    UpdateLUT
End Sub

Sub SaveLUT
    SaveValue cGameName, "LUTImage", LUTImage
End Sub

Sub NextLUT:LUTImage = (LUTImage + 1) MOD 22:UpdateLUT:SaveLUT:SetLUTLine "Color LUT image " & table1.ColorGradeImage:End Sub

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

' New LUT postit
Sub SetLUTLine(String)
    Dim Index
    Dim xFor
    Index = 1
    LUBack.imagea = "PostItNote"
    String = CL2(String)
    For xFor = 1 to 40
        Eval("LU" &xFor).imageA = GetHSChar(String, Index)
        Index = Index + 1
    Next
End Sub

Sub HideLUT
    SetLUTLine ""
    LUBack.imagea = "PostitBL"
End Sub

Function CL2(NumString) 'center line
    Dim Temp, TempStr
    If Len(NumString)> 40 Then NumString = Left(NumString, 40)
    Temp = (40 - Len(NumString) ) \ 2
    TempStr = Space(Temp) & NumString & Space(Temp)
    CL2 = TempStr
End Function

'***********************************************************************
' *********************************************************************
'  *********     G A M E  C O D E  S T A R T S  H E R E      *********
' *********************************************************************
'***********************************************************************

Sub VPObjects_Init 'init objects
    TurnOffPlayfieldLights()
End Sub

' Dim all the variables
Dim BonusValue(10)
Dim Zipped
Dim PopBonus
Dim PopBonusValue(6)
Dim BonusEnabled

Sub Game_Init 'called at the start of a new game
    Dim x
    'Start music?
    'Init variables?
    BonusValue(1) = 10
    BonusValue(2) = 50
    BonusValue(3) = 100
    BonusValue(4) = 100
    BonusValue(5) = 200
    BonusValue(6) = 300
    BonusValue(7) = 400
    BonusValue(8) = 500
    BonusValue(9) = 500
    BonusValue(10) = 500 'this is actually the extra ball
    Zipped = 0
    PopBonusValue(1) = 10
    PopBonusValue(2) = 50
    PopBonusValue(3) = 100
    PopBonusValue(4) = 200
    PopBonusValue(5) = 300
    PopBonusValue(6) = 500
    BonusEnabled = False
    'Start or init timers
    'Init lights?
    TurnOffPlayfieldLights
    TurnOFF aBlackLights
    TurnON aRedLights
End Sub

Sub StopEndOfBallMode     'called when the last ball is drained
End Sub

Sub ResetNewBallVariables 'init variables & lights new ball/player
    'TurnOffPlayfieldLights
    Bonus = 1
    UpdateBonusLights
    BonusEnabled = False
    PopBonus = 1
    UpdatePoPBonusLights
    PlaySound "Reelinit"
    UnzipFlippers
    ' turn on/off new ball Lights
    LightShootAgain.State = 0
    blx.State = 0
    Light005.State = 0
    'check if red or black are on and turn off outlane light in case it was On
    If Light006.State then 'red lights are On
        Light003.State = 0
    Else
        Light002.State = 0
    End If
End Sub

Sub TurnOffPlayfieldLights
    Dim a
    For each a in aLights
        a.State = 0
    Next
End Sub

Sub ZipFlippers
    PlaySoundAt "Fx_diverter", LeftFlipper
    LeftFlipper.Enabled = 0
    LeftFlipperRubber.Visible = 0
    LeftFlipperP.Visible = 0
    LeftFlipper1.Enabled = 1
    LeftFlipperRubber1.Visible = 1
    LeftFlipperP1.Visible = 1
    PlaySoundAt "Fx_diverter", RightFlipper
    RightFlipper.Enabled = 0
    RightFlipperRubber.Visible = 0
    RightFlipperP.Visible = 0
    RightFlipper1.Enabled = 1
    RightFlipperRubber1.Visible = 1
    RightFlipperP1.Visible = 1
    Zipped = 1
End Sub

Sub UnzipFlippers
    If Zipped Then
        Zipped = 0
        PlaySoundAt "Fx_diverter", LeftFlipper
        LeftFlipper.Enabled = 1
        LeftFlipperRubber.Visible = 1
        LeftFlipperP.Visible = 1
        LeftFlipper1.Enabled = 0
        LeftFlipperRubber1.Visible = 0
        LeftFlipperP1.Visible = 0
        PlaySoundAt "Fx_diverter", RightFlipper
        RightFlipper.Enabled = 1
        RightFlipperRubber.Visible = 1
        RightFlipperP.Visible = 1
        RightFlipper1.Enabled = 0
        RightFlipperRubber1.Visible = 0
        RightFlipperP1.Visible = 0
    End If
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
' Mushrooms
'***********

Sub mushroom001_Hit 'center - blue
    If Tilted Then Exit Sub
    PlaySoundAt SoundFXDOF("fx_mini bumper", 106, DOFPulse, DOFContactors), mushroom001
    AddScore 10
    AddBonus 1
End Sub

Sub mushroom001_Animate:mushrromP001.TransZ = - mushroom001.CurrentRingOffset / 4:End Sub

Sub mushroom002_Hit 'left - white
    If Tilted Then Exit Sub
    PlaySoundAt SoundFXDOF("fx_mini bumper", 105, DOFPulse, DOFContactors), mushroom002
    AddScore 10
    AddPopBonus 1
    UnZipFlippers
End Sub

Sub mushroom002_Animate:mushrromP002.TransZ = - mushroom002.CurrentRingOffset / 4:End Sub

Sub mushroom004_Hit 'right -red
    If Tilted Then Exit Sub
    PlaySoundAt SoundFXDOF("fx_mini bumper", 107, DOFPulse, DOFContactors), mushroom004
    AddScore 10
    AddPopBonus 1
    UnZipFlippers
    TurnOFF aBlackLights
    TurnON aRedLights
End Sub

Sub mushroom004_Animate:mushrromP004.TransZ = - mushroom004.CurrentRingOffset / 4:End Sub

Sub mushroom005_Hit 'Top left - yellow
    If Tilted Then Exit Sub
    PlaySoundAt SoundFXDOF("fx_mini bumper", 109, DOFPulse, DOFContactors), mushroom005
    AddScore 10
    If Light005.State Then ZipFlippers
End Sub

Sub mushroom005_Animate:mushrromP005.TransZ = - mushroom005.CurrentRingOffset / 4:End Sub

Sub mushroom003_Hit 'Top right -black
    If Tilted Then Exit Sub
    PlaySoundAt SoundFXDOF("fx_mini bumper", 110, DOFPulse, DOFContactors), mushroom003
    AddScore 10
    AddPopBonus 1
    UnZipFlippers
    TurnOff aRedLights
    TurnON aBlackLights
End Sub

Sub mushroom003_Animate:mushrromP003.TransZ = - mushroom003.CurrentRingOffset / 4:End Sub


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
    AddScore 10
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
    AddScore 10
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 1:RightSLing004.Visible = 0:RightSLing003.Visible = 1:Remk.RotX = 14
        Case 2:RightSLing003.Visible = 0:RightSLing002.Visible = 1:Remk.RotX = 2
        Case 3:RightSLing002.Visible = 0:Remk.RotX = -20:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

' 1 point rebounds

Sub rlband005_Hit:AddScore 1:End Sub
Sub rlband007_Hit:AddScore 1:End Sub
Sub rlband004_Hit:AddScore 1:End Sub


' Top Button
Sub button001_Hit 'top
    PlaySoundAt "fx_sensor", button001
    If Tilted Then Exit Sub
    Light005.State = 1
    AddBonus 1
End Sub

' Red buttons - 100 points
Sub button003_Hit
    PlaySoundAt "fx_sensor", button003
    If Tilted Then Exit Sub
    If Light013.State then
        AddBonus 1
        AddScore 100
    End If
End Sub

Sub button005_Hit
    PlaySoundAt "fx_sensor", button005
    If Tilted Then Exit Sub
    If Light015.State then
        AddBonus 1
        AddScore 100
    End If
End Sub

Sub button007_Hit
    PlaySoundAt "fx_sensor", button007
    If Tilted Then Exit Sub
    If Light010.State then
        AddBonus 1
        AddScore 100
    End If
End Sub

Sub button009_Hit
    PlaySoundAt "fx_sensor", button009
    If Tilted Then Exit Sub
    If Light007.State then
        AddBonus 1
        AddScore 100
    End If
End Sub

Sub button011_Hit
    PlaySoundAt "fx_sensor", button011
    If Tilted Then Exit Sub
    If Light006.State then
        AddBonus 1
        AddScore 100
    End If
End Sub

' Black buttons - 10 points

Sub button002_Hit
    PlaySoundAt "fx_sensor", button002
    If Tilted Then Exit Sub
    If Light012.State then
        AddBonus 1
        AddScore 10
    End If
End Sub

Sub button004_Hit
    PlaySoundAt "fx_sensor", button004
    If Tilted Then Exit Sub
    If Light014.State then
        AddBonus 1
        AddScore 10
    End If
End Sub

Sub button008_Hit
    PlaySoundAt "fx_sensor", button008
    If Tilted Then Exit Sub
    If Light009.State then
        AddBonus 1
        AddScore 10
    End If
End Sub

Sub button006_Hit
    PlaySoundAt "fx_sensor", button006
    If Tilted Then Exit Sub
    If Light011.State then
        AddBonus 1
        AddScore 10
    End If
End Sub

Sub button010_Hit
    PlaySoundAt "fx_sensor", button010
    If Tilted Then Exit Sub
    If Light008.State then
        AddBonus 1
        AddScore 10
    End If
End Sub

' Outlanes

Sub Trigger003_Hit 'left outlane
    PlaySoundAt "fx_sensor", Trigger003
    If Tilted Then Exit Sub
    If Light002.State Then
        'Collect Bonus at the drain
        BonusEnabled = True
        If blx.State Then
            AwardExtraBall
        End If
    End If
End Sub

Sub Trigger004_Hit 'right outlane
    PlaySoundAt "fx_sensor", Trigger004
    If Tilted Then Exit Sub
    If Light003.State Then
        'collect Bonus at the drain
        BonusEnabled = True
        If blx.State Then
            AwardExtraBall
        End If
    End If
End Sub

Sub Trigger005_Hit 'Left top outlane
    PlaySoundAt "fx_sensor", Trigger005
    If Tilted Then Exit Sub
    Addscore 100
End Sub

' Hole

Sub Kicker001_Hit 'Right
    Dim tmp
    PlaySoundAt "fx_kicker_enter", Kicker001
    If Tilted Then vpmTimer.AddTimer 200, "PlaySoundAt SoundFXDOF (""fx_kicker"", 108, DOFPulse, DOFContactors), kicker001:kicker001.kick 210,14 '"

    'Count the PopBonus
    Addscore PopBonusValue(PopBonus)
    PopBonus = 1
    UpdatePopBonusLights
    vpmTimer.AddTimer 800, "PlaySoundAt SoundFXDOF (""fx_kicker"", 108, DOFPulse, DOFContactors), kicker001:kicker001.kick 210,14 '"
End Sub

' Pop Score

Sub AddPopBonus(popbonuspoints)
    If(Tilted = False) Then
        PopBonus = PopBonus + popbonuspoints
        If PopBonus> 6 Then PopBonus = 6
        UpdatePopBonusLights
    End if
End Sub

Sub UpdatePopBonusLights
    Select Case PopBonus
        Case 0:ps1.State = 0:ps2.State = 0:ps3.State = 0:ps4.State = 0:ps5.State = 0:ps6.State = 0
        Case 1:ps1.State = 1:ps2.State = 0:ps3.State = 0:ps4.State = 0:ps5.State = 0:ps6.State = 0
        Case 2:ps1.State = 0:ps2.State = 1:ps3.State = 0:ps4.State = 0:ps5.State = 0:ps6.State = 0
        Case 3:ps1.State = 0:ps2.State = 0:ps3.State = 1:ps4.State = 0:ps5.State = 0:ps6.State = 0
        Case 4:ps1.State = 0:ps2.State = 0:ps3.State = 0:ps4.State = 1:ps5.State = 0:ps6.State = 0
        Case 5:ps1.State = 0:ps2.State = 0:ps3.State = 0:ps4.State = 0:ps5.State = 1:ps6.State = 0
        Case 6:ps1.State = 0:ps2.State = 0:ps3.State = 0:ps4.State = 0:ps5.State = 0:ps6.State = 1
    End Select
End Sub

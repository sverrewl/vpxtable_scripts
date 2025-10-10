' Hi-Score / IPD No. 1160 / May, 1967 / 4 Players
' D. Gottlieb & Company (1931-1977) [Trade Name: Gottlieb]
' VPX8 table by jpsalas with graphics by Halen. Version 6.0.0

Option Explicit
Randomize

' core.vbs constants
Const BallSize = 50 ' 50 is the normal size
Const BallMass = 1  ' 1 is the normal ball mass.

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
Const TableName = "hi-score" ' file name to save highscores and other variables
Const cGameName = "hi-score" ' B2S name
Const MaxPlayers = 4           ' 1 to 4 can play
Const MaxMultiplier = 1        ' limit bonus multiplier
Const FreePlay = False         ' Free play or coins
Const Special1 = 4400          ' extra ball or credit
Const Special2 = 5000
Const Special3 = 5600
Const Special4 = 6200

' Global variables
Dim PlayersPlayingGame
Dim CurrentPlayer
Dim Credits
Dim Bonus
Dim BallsRemaining(4)
Dim BonusMultiplier
Dim ExtraBallsAwards(4)
Dim Special1Awarded(4)
Dim Special2Awarded(4)
Dim Special3Awarded(4)
Dim Special4Awarded(4)
Dim Score(4)
Dim HighScore
Dim Match
Dim Tilt
Dim TiltSensitivity
Dim Tilted
Dim Add10
Dim Add100
Dim Add1

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
    ScoreReel1.SetValue HSScore(1)
    If B2SOn then
        Controller.B2SSetScorePlayer 1, HSScore(1)
    End If
    UpdateCredits

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
    Add10 = 0
    Add100 = 0
    Add1 = 0

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
    If credits> 0 then:DOF 123, DOFOn
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

    ' keys during game

    If bGameInPlay AND NOT Tilted Then

    ' tilt keys
    If keycode = LeftTiltKey Then Nudge 90, 8:PlaySound "fx_nudge", 0, 1, -0.1, 0.25:CheckTilt
    If keycode = RightTiltKey Then Nudge 270, 8:PlaySound "fx_nudge", 0, 1, 0.1, 0.25:CheckTilt
    If keycode = CenterTiltKey Then Nudge 0, 9:PlaySound "fx_nudge", 0, 1, 1, 0.25:CheckTilt
    If keycode = MechanicalTilt Then Tilt = 16:CheckTilt

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
                        DOF 123, DOFOff
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
'Controller.Stop
End Sub

'********************
'     Flippers
'********************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFXDOF("fx_flipperup", 101, DOFOn, DOFFlippers), LeftFlipper
        LeftFlipper.RotateToEnd:LeftFlipper001.RotateToEnd
    Else
        PlaySoundAt SoundFXDOF("fx_flipperdown", 101, DOFOff, DOFFlippers), LeftFlipper
        LeftFlipper.RotateToStart:LeftFlipper001.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFXDOF("fx_flipperup", 102, DOFOn, DOFFlippers), RightFlipper
        RightFlipper.RotateToEnd:RightFlipper001.RotateToEnd
    Else
        PlaySoundAt SoundFXDOF("fx_flipperdown", 102, DOFOff, DOFFlippers), RightFlipper
        RightFlipper.RotateToStart:RightFlipper001.RotateToStart
    End If
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, Vol(ActiveBall), pan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, Vol(ActiveBall), pan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub LeftFlipper001_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, Vol(ActiveBall), pan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper001_Collide(parm)
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

Sub LeftFlipper001_Animate
    LeftFlipperP001.Rotz = LeftFlipper001.CurrentAngle
    LeftFlipperRubber001.Rotz = LeftFlipper001.CurrentAngle
End Sub

Sub RightFlipper001_Animate
    RightFlipperP001.Rotz = RightFlipper001.CurrentAngle
    RightFlipperRubber001.Rotz = RightFlipper001.CurrentAngle
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
    If bGameInPlay Then
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
        Bumper003.Threshold = 100
        Bumper004.Threshold = 100
        Bumper005.Threshold = 100
    Else
        GiOn
        Bumper001.Threshold = 2
        Bumper002.Threshold = 1
        Bumper003.Threshold = 1
        Bumper004.Threshold = 1
        Bumper005.Threshold = 1
    End If
End Sub

Sub TiltRecoveryTimer_Timer()
    ' all the balls have drained
    If(BallsOnPlayfield = 0) Then
        EndOfBall()
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
'     (reduced version - only for this table)
'****************************************************

Const tnob = 19   'total number of balls
Const lob = 0     'number of locked balls
Const maxvel = 25 'max ball velocity
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
    Dim BOT, b, speedfactorx, speedfactory
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
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
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
    Bonus = 1
    UpdateBallInPlay

    Clear_Match

    ' init other variables
    Tilt = 0

    ' init game variables
    Game_Init()

    ' start a music?
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
    BallRelease.CreateSizedBallWithMass BallSize / 2, BallMass
    BallsOnPlayfield = BallsOnPlayfield + 1
    UpdateBallInPlay
    PlaySoundAt SoundFXDOF("fx_Ballrel", 120, DOFPulse, DOFContactors), BallRelease
    BallRelease.Kick 90, 4
End Sub

' player lost the ball

Sub EndOfBall()
    'debug.print "EndOfBall"
    'Dim AwardPoints, TotalBonus, ii
    'AwardPoints = 0
    'TotalBonus = 0
    ' Lost the first ball, now it cannot accept more players
    bOnTheFirstBall = False

    'No bonus in this table

    ' add bonus if no tilt
    ' tilt system will take care of the next ball

    'If NOT Tilted Then
    '    BonusCountTimer.Interval = 200 + 50 * BonusMultiplier
    '    BonusCountTimer.Enabled = 1
    'Else
    vpmtimer.addtimer 3000, "EndOfBall2 '"
'End If
End Sub

Sub BonusCountTimer_Timer
'debug.print "BonusCount_Timer"
'If Bonus> 0 Then
'    Bonus = Bonus -1
'    AddScore 1000 * BonusMultiplier
'    UpdateBonusLights
'Else
'    BonusCountTimer.Enabled = 0
'    vpmtimer.addtimer 1000, "EndOfBall2 '"
'End If
End Sub

Sub UpdateBonusLights
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
        'LightShootAgain.State = 0
        'If B2SOn then
        'Controller.B2SSetShootAgain 0
        'end if
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
    PlaySound "fx_match"
    Match = INT(RND(1) * 10) ' random between 1 and 9
    Display_Match
    If(Score(CurrentPlayer) MOD 10) = Match Then
        PlaySound SoundFXDOF("fx_knocker", 122, DOFPulse, DOFknocker)
        AddCredits 1
    End If
End Sub

Sub Clear_Match()
    Match = 0
    Display_Match
    m0.State = 0
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
        Case 1, 10, 100
            Score(CurrentPlayer) = Score(CurrentPlayer) + points
            UpdateScore points
            ' sounds
            If Points = 10 AND(Score(CurrentPlayer) MOD 100) \ 10 = 0 Then 'new reel 100
                PlaySound SoundFXDOF("bell100", 132, DOFPulse, DOFChimes)
            ElseIf Points = 1 AND(Score(CurrentPlayer) MOD 10) = 0 Then    'new reel 10
                PlaySound SoundFXDOF("bell10", 132, DOFPulse, DOFChimes)
            Else
                PlaySound SoundFXDOF("bell" &Points, 131, DOFPulse, DOFChimes)
            End If
        Case 20, 30, 40, 50
            Add10 = Add10 + Points \ 10
            AddScore10Timer.Enabled = TRUE
        Case 200, 300, 400, 500
            Add100 = Add100 + Points \ 100
            AddScore100Timer.Enabled = TRUE
        Case 2, 3, 4, 5
            Add1 = Add1 + Points
            AddScore1Timer.Enabled = TRUE
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
    If Score(CurrentPlayer) >= Special3 AND Special3Awarded(CurrentPlayer) = False Then
        AwardSpecial
        Special3Awarded(CurrentPlayer) = True
    End If
    If Score(CurrentPlayer) >= Special4 AND Special4Awarded(CurrentPlayer) = False Then
        AwardSpecial
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
        If Bonus> 15 Then
            Bonus = 1
        End If
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
    'Case 1:bmli2.State = 1:bmli3.State = 0:bmli4.State = 0
    'Case 2:bmli2.State = 0:bmli3.State = 1:bmli4.State = 0
    'Case 3:bmli2.State = 0:bmli3.State = 0:bmli4.State = 1
    End Select
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
    ScoreReel1.SetValue 1:ScoreReel1.ResetToZero
    ScoreReel2.SetValue 1:ScoreReel2.ResetToZero
    ScoreReel3.SetValue 1:ScoreReel3.ResetToZero
    ScoreReel4.SetValue 1:ScoreReel4.ResetToZero
    If B2SOn then
        Controller.B2SSetScorePlayer1 0
        Controller.B2SSetScoreRolloverPlayer1 0
        Controller.B2SSetScorePlayer2 0
        Controller.B2SSetScoreRolloverPlayer2 0
        Controller.B2SSetScorePlayer3 0
        Controller.B2SSetScoreRolloverPlayer3 0
        Controller.B2SSetScorePlayer4 0
        Controller.B2SSetScoreRolloverPlayer4 0
    end if
End Sub

Sub AddCredits(value) 'limit to 9 credits
    If Credits <9 Then
        Credits = Credits + value
        CreditsReel.SetValue Credits
        UpdateCredits
        DOF 123, DOFOn
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
    Else
    'Addscore 5000
    END If
End Sub

Sub AwardSpecial()
    PlaySound SoundFXDOF("fx_knocker", 122, DOFPulse, DOFknocker)
    AddCredits 1
End Sub

Sub AwardAddaBall()
    If BallsRemaining(CurrentPlayer) <11 Then
        PlaySound SoundFXDOF("fx_knocker", 122, DOFPulse, DOFknocker)
        BallsRemaining(CurrentPlayer) = BallsRemaining(CurrentPlayer) + 1
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
    'For each x in aLights
    '    x.State = 2
    'Next
    'turn on GameOver light
    TurnON aGameOver
    'update current player and balls
    bip5.State = 0
End Sub

Sub StopAttractMode()
    Dim x
    bAttractMode = False
    'TurnOffPlayfieldLights
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
' GNMOD - Multiple High Score Display and Collection
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
HSScore(2) = 2500
HSScore(3) = 2000
HSScore(4) = 1500
HSScore(5) = 1000

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
        AwardSpecial
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
    rkicker001.CreateSizedBallWithMass 18, 1
    DrainBlock.Isdropped = 1
    TurnOffPlayfieldLights()
    A = 0
    B = 0
    C = 0
    ResetSpinnerLights
    bump = 0
    BumperL001.State = 1
    BumperL003.State = 0
    BumperL002.State = 0
    BumperL004.State = 0
    BumperL005.State = 0
    vpmTimer.AddTimer 1500, "UpdateABC '"
End Sub

' Dim all the variables
Dim bump      'rotate counter
Dim A, B, C   'values

Sub Game_Init 'called at the start of a new game
    'Start music?
    'Init variables?
    'Start or init timers
    'Init lights?
End Sub

Sub StopEndOfBallMode     'called when the last ball is drained
End Sub

Sub ResetNewBallVariables 'init variables & lights new ball/player
End Sub

Sub TurnOffPlayfieldLights
    Dim a
    For each a in aLights
        a.State = 0
    Next
    If B2SOn then
        For a = 10 to 18
            Controller.B2SSetData a, 0
        Next
    end if
End Sub

Sub ResetSpinnerLights
    'values
    'Lights
    Light014.State = 1
    Light016.State = 1
    Light015.State = 1
    Light017.State = 1
    Light019.State = 1
    Light020.State = 1
    Light011.State = 1
    Light012.State = 1
    Light013.State = 1
    If B2SOn then
        Controller.B2SSetData 10, 1
        Controller.B2SSetData 11, 1
        Controller.B2SSetData 12, 1
    end if
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

Sub Bumper001_Hit '50 points bumper
    If Tilted Then Exit Sub
    PlaySoundAt "fx_passive_bumper", bumper001
    AddScore 5 + 45 * BumperL001.State
End Sub

Sub Bumper002_Hit 'left bumper
    If Tilted Then Exit Sub
    PlaySoundAt "fx_bumper", Bumper002
    AddScore 1 + 9 * BumperL002.State
End Sub

Sub Bumper003_Hit 'right bumper
    If Tilted Then Exit Sub
    PlaySoundAt "fx_bumper", Bumper003
    AddScore 1 + 9 * BumperL003.State
End Sub

Sub Bumper004_Hit 'lower left bumper
    If Tilted Then Exit Sub
    PlaySoundAt "fx_bumper", Bumper004
    AddScore 1 + 9 * BumperL004.State
End Sub

Sub Bumper005_Hit 'lower right bumper
    If Tilted Then Exit Sub
    PlaySoundAt "fx_bumper", Bumper005
    AddScore 1 + 9 * BumperL005.State
End Sub

' OutLanes

Sub Trigger001_Hit
    PlaySoundAt "fx_sensor", Trigger001
    If Tilted Then Exit Sub
    AddScore 50
    If Light001.State Then StartSpinner
End Sub

Sub Trigger002_Hit
    PlaySoundAt "fx_sensor", Trigger002
    If Tilted Then Exit Sub
    AddScore 50
    If Light002.State Then StartSpinner
End Sub

Sub Trigger003_Hit
    PlaySoundAt "fx_sensor", Trigger003
    If Tilted Then Exit Sub
    AddScore 50
    If Light003.State Then StartSpinner
End Sub

Sub Trigger004_Hit
    PlaySoundAt "fx_sensor", Trigger004
    If Tilted Then Exit Sub
    AddScore 50
    If Light004.State Then StartSpinner
End Sub

' side lanes (C)
Sub Trigger005_Hit
    PlaySoundAt "fx_sensor", Trigger005
    If Tilted Then Exit Sub
    AddScore 50
    If Light020.State Then AdvanceC:Light020.State = 0
End Sub

Sub Trigger006_Hit
    PlaySoundAt "fx_sensor", Trigger006
    If Tilted Then Exit Sub
    AddScore 50
    If Light019.State Then AdvanceC:Light019.State = 0
End Sub

' top lane
Sub Trigger007_Hit
    PlaySoundAt "fx_sensor", Trigger007
    If Tilted Then Exit Sub
    ShootAgain.State = 0
End Sub

' gate
Sub Trigger008_Hit
    If Tilted Then Exit Sub
    Addscore 50
End Sub

' 1 point rebounds

Sub a1PointRubbers_Hit(idx):AddScore 1:RotateBumpers:End Sub
Sub rlband001_Hit:AddScore 1:Light001.State = 0:Light002.State = 1:Light003.State = 0:Light004.State = 1:Light022.State = 0:Light021.State = 1:End Sub
Sub rlband002_Hit:AddScore 1:Light001.State = 1:Light002.State = 0:Light003.State = 1:Light004.State = 0:Light022.State = 1:Light021.State = 0:End Sub

Sub RotateBumpers
    bump = (bump + 1) MOD 7
    Select Case bump
        Case 0:BumperL001.State = 1:BumperL002.State = 0:BumperL003.State = 0:BumperL004.State = 0:BumperL005.State = 0
        Case 1:BumperL001.State = 1:BumperL002.State = 1:BumperL003.State = 0:BumperL004.State = 0:BumperL005.State = 1
        Case 2:BumperL001.State = 1:BumperL002.State = 0:BumperL003.State = 1:BumperL004.State = 1:BumperL005.State = 0
        Case 3:BumperL001.State = 0:BumperL002.State = 0:BumperL003.State = 0:BumperL004.State = 0:BumperL005.State = 0
        Case 4:BumperL001.State = 1:BumperL002.State = 1:BumperL003.State = 0:BumperL004.State = 0:BumperL005.State = 1
        Case 5:BumperL001.State = 1:BumperL002.State = 0:BumperL003.State = 1:BumperL004.State = 1:BumperL005.State = 0
        Case 6:BumperL001.State = 1:BumperL002.State = 1:BumperL003.State = 1:BumperL004.State = 1:BumperL005.State = 1
    End Select
End Sub

'Targets

Sub Target001_Hit 'A
    PlaySoundAtBall "fx_target"
    If Tilted Then Exit Sub
    AddScore 50
    If Light014.State Then
        AdvanceA
        Light014.State = 0
    End If
End Sub

Sub Target004_Hit 'A
    PlaySoundAtBall "fx_target"
    If Tilted Then Exit Sub
    AddScore 50
    If Light016.State Then
        AdvanceA
        Light016.State = 0
    End If
End Sub

Sub Target002_Hit 'B
    PlaySoundAtBall "fx_target"
    If Tilted Then Exit Sub
    AddScore 50
    If Light015.State Then
        AdvanceB
        Light015.State = 0
    End If
End Sub

Sub Target003_Hit 'B
    PlaySoundAtBall "fx_target"
    If Tilted Then Exit Sub
    AddScore 50
    If Light017.State Then
        AdvanceB
        Light017.State = 0
    End If
End Sub

' Holes

Sub Kicker1_Hit
    Dim wait
    PlaySoundAt "fx_hole_enter", Kicker1
    'AddScore
    Addscore 50
    If Light022.State Then
        StartSpinner
        Wait = 8000
    Else
        Wait = 1000
    End If
    vpmTimer.AddTimer wait, "kicker1.kick 15,24: PlaySoundAt""fx_kicker"",kicker1 '"
End Sub

Sub Kicker2_Hit
    Dim wait
    PlaySoundAt "fx_hole_enter", Kicker2
    'AddScore
    Addscore 50
    If Light021.State Then
        StartSpinner
        Wait = 8000
    Else
        Wait = 1000
    End If
    vpmTimer.AddTimer wait, "kicker2.kick -15,24: PlaySoundAt""fx_kicker"",kicker2 '"
End Sub

' Spinner & ABC scoring

Dim RotPos
RotPos = 0

Dim kv(11) 'kickers letter after rotation
' star = 0, A = 1, B = 2, C = 3
kv(0) = Array(0, 1, 2, 1, 2, 1, 3, 1, 2, 1, 3, 2)
kv(1) = Array(2, 0, 1, 2, 1, 2, 1, 3, 1, 2, 1, 3)
kv(2) = Array(3, 2, 0, 1, 2, 1, 2, 1, 3, 1, 2, 1)
kv(3) = Array(1, 3, 2, 0, 1, 2, 1, 2, 1, 3, 1, 2)
kv(4) = Array(2, 1, 3, 2, 0, 1, 2, 1, 2, 1, 3, 1)
kv(5) = Array(1, 2, 1, 3, 2, 0, 1, 2, 1, 2, 1, 3)
kv(6) = Array(3, 1, 2, 1, 3, 2, 0, 1, 2, 1, 2, 1)
kv(7) = Array(1, 3, 1, 2, 1, 3, 2, 0, 1, 2, 1, 2)
kv(8) = Array(2, 1, 3, 1, 2, 1, 3, 2, 0, 1, 2, 1)
kv(9) = Array(1, 2, 1, 3, 1, 2, 1, 3, 2, 0, 1, 2)
kv(10) = Array(2, 1, 2, 1, 3, 1, 2, 1, 3, 2, 0, 1)
kv(11) = Array(1, 2, 1, 2, 1, 3, 1, 2, 1, 3, 2, 0)

Sub StartSpinner
    Dim i
    ' raise the drain block
    DrainBlock.IsDropped = 0
    'disable & remove the ball from the current kicker
    For each i in aRuletaK:i.Enabled = 0:i.DestroyBall:Next
    'start the timer
    SpinnerTimer.Interval = 10
    SpinnerTimer.Enabled = 1
    ' Enabled speedup triggers
    Speedup.Enabled = 1
    Speedup001.Enabled = 1
    Speedup002.Enabled = 1
    Speedup003.Enabled = 1
    ' Create a ball and kick it
    rkicker007.CreateSizedBallWithMass 18, 1
    rkicker007.Kick 270, 20
    ' add timer to stop the SpinnerTimer
    vpmTimer.AddTimer 1500 + RndNbr(10) * 100, "StopSpinner '"
End Sub

Sub SpinnerTimer_Timer
    RotPos = (RotPos + 1) MOD 12
    ruleta.Rotz = Rotpos * 30
    SpinnerTimer.Interval = SpinnerTimer.Interval + 1
End Sub

Sub StopSpinner
    'stop the timer
    SpinnerTimer.Enabled = 0
    ' disable the speed up triggers
    Speedup.Enabled = 0
    Speedup001.Enabled = 0
    Speedup002.Enabled = 0
    Speedup003.Enabled = 0
    ' enable the kickers after a few seconds or so
    vpmTimer.AddTimer 4000 + RndNbr(10) * 10, "Dim i: For each i in aRuletaK: i.Enabled = 1: next '"
End Sub

Sub aRuletaK_Hit(idx)
    Select Case kv(RotPos) (idx)
        Case 0:AwardExtraBall 'Star
        Case 1:               'A
            Select Case A
                Case 0:Addscore 5
                Case 1:Addscore 50
                Case 2:Addscore 500:ResetABC
            End Select
        Case 2: 'B
            Select Case B
                Case 0:Addscore 50
                Case 1:Addscore 500
                Case 2:Addscore 500:Addscore 500:ResetABC
            End Select
        Case 3: 'C
            Select Case C
                Case 0:Addscore 500
                Case 1:Addscore 500:Addscore 500
                Case 2:Addscore 500:Addscore 500:Addscore 500:Addscore 500:ResetABC
            End Select
    End Select
    'drop the drain block if it was up
    DrainBlock.Isdropped = 1
End Sub

Sub Speedup_Hit:ActiveBall.Velx = -20:End Sub
Sub Speedup001_Hit:ActiveBall.Vely = -20:End Sub
Sub Speedup002_Hit:ActiveBall.Velx = 20:End Sub
Sub Speedup003_Hit:ActiveBall.Vely = 20:End Sub

Sub UpdateABC 'Lights
    Select Case A
        Case 0:Light011.State = 1:Light008.State = 0:Light005.State = 0
    If B2SOn then
        Controller.B2SSetData 10, 1
        Controller.B2SSetData 13, 0
        Controller.B2SSetData 16, 0
    end if
        Case 1:Light011.State = 0:Light008.State = 1:Light005.State = 0
    If B2SOn then
        Controller.B2SSetData 10, 0
        Controller.B2SSetData 13, 1
        Controller.B2SSetData 16, 0
    end if
        Case 2:Light011.State = 0:Light008.State = 0:Light005.State = 1
    If B2SOn then
        Controller.B2SSetData 10, 0
        Controller.B2SSetData 13, 0
        Controller.B2SSetData 16, 1
    end if
    End Select
    Select Case B
        Case 0:Light012.State = 1:Light009.State = 0:Light006.State = 0
    If B2SOn then
        Controller.B2SSetData 11, 1
        Controller.B2SSetData 14, 0
        Controller.B2SSetData 17, 0
    end if
        Case 1:Light012.State = 0:Light009.State = 1:Light006.State = 0
    If B2SOn then
        Controller.B2SSetData 11, 0
        Controller.B2SSetData 14, 1
        Controller.B2SSetData 17, 0
    end if
        Case 2:Light012.State = 0:Light009.State = 0:Light006.State = 1
    If B2SOn then
        Controller.B2SSetData 11, 0
        Controller.B2SSetData 14, 0
        Controller.B2SSetData 17, 1
    end if
    End Select
    Select Case C
        Case 0:Light013.State = 1:Light010.State = 0:Light007.State = 0
    If B2SOn then
        Controller.B2SSetData 12, 1
        Controller.B2SSetData 15, 0
        Controller.B2SSetData 18, 0
    end if
        Case 1:Light013.State = 0:Light010.State = 1:Light007.State = 0
    If B2SOn then
        Controller.B2SSetData 12, 0
        Controller.B2SSetData 15, 1
        Controller.B2SSetData 18, 0
    end if
        Case 2:Light013.State = 0:Light010.State = 0:Light007.State = 1
    If B2SOn then
        Controller.B2SSetData 12, 0
        Controller.B2SSetData 15, 0
        Controller.B2SSetData 18, 1
    end if
    End Select
End Sub

Sub AdvanceA
    A = A + 1
    If A> 2 Then A = 2
    UpdateABC
End Sub

Sub AdvanceB
    B = B + 1
    If B> 2 Then B = 2
    UpdateABC
End Sub

Sub AdvanceC
    C = C + 1
    If C> 2 Then C = 2
    UpdateABC
End Sub

Sub ResetABC
    ResetSpinnerLights
    A = 0
    B = 0
    C = 0
    UpdateABC
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

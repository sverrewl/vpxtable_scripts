    ' ***********************************************************************
    '               VISUAL PINBALL X EM Script by JPSalas
    '                  Basic EM script up to 4 players
    '             uses core.vbs for extra functions
    '
    '  Speakeasy / Playmatic / IPD No. 2269 / May, 1977 / 4 Players
    '  VPX8 version by JPSalas, pedator and STAT 2025, version 6.0.0
    ' ***********************************************************************

    ' DOF config - leeoneil
    ' Flippers L/R - 101/102
    ' Slingshots L/R - 109/110
    ' Bumpers (Bumpers Back Left/Back Center/Back Right) - 103/105/104
    ' Ejects-holes (Bumpers Middle Left/Middle Center/Middle Right) - 113/115/114
    '
    ' LED backboard
    ' Flasher Outside Left - 111/216/230/231/232
    ' Flasher left - 106/217/230/231/232
    ' Flasher center - 108/218/230/231/232
    ' Flasher right - 107/219/230/231/232
    ' Flasher Outside Right - 112/221/230/231/232
    '
    ' Strobe - 130
    ' Start button - 131
    ' Undercab complex - 132/200
    '
    ' Knocker - 300

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
    Const TableName = "PlaymaticSpeakeasy" ' file name to save highscores and other variables
    Const cGameName = "PlaymaticSpeakeasy" ' B2S name
    Const MaxPlayers = 4                   ' 1 to 4 can play
    Const MaxMultiplier = 2                ' limit bonus multiplier
    Const FreePlay = False                 ' Free play or coins
    Const Special1 = 490000                ' extra ball or credit
    Const Special2 = 680000
    Const Special3 = 860000

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
    Dim Score(4)
    Dim HighScore
    Dim Match
    Dim Tilt
    Dim TiltSensitivity
    Dim Tilted
    Dim Add10
    Dim Add100
    Dim Add1000
    Dim Add10000
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
        LoadEM

        ' load highscore
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
        CurrentPlayer = 0
        Tilt = 0
        TiltSensitivity = 6
        Tilted = False
        Match = 0
        bJustStarted = True
        Add10 = 0
        Add100 = 0
        Add1000 = 0
        Add10000 = 0

        ' setup table in game over mode
        EndOfGame

        'turn on GI lights
        vpmtimer.addtimer 1500, "GiOn '"
        DOF 132, DOFOn

        ' Init som objects, like walls, targets
        VPObjects_Init

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


    '******
    ' Keys
    '******

    Sub Table1_KeyDown(ByVal Keycode)

        If EnteringInitials then
            CollectInitials(keycode)
            Exit Sub
        End If

        ' add coins
        If Keycode = AddCreditKey Then
            If(Tilted = False)Then
                If Not bFreePlay Then DOF 131, DOFOn
                AddCredits 1
                PlaySound "fx_coin"
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
                If((PlayersPlayingGame <MaxPlayers)AND(bOnTheFirstBall = True))Then

                    If(bFreePlay = True)Then
                        PlayersPlayingGame = PlayersPlayingGame + 1
                        UpdateBallInPlay
                        UpdatePLayersCanPlay
                    Else
                        If(Credits> 0)then
                            PlayersPlayingGame = PlayersPlayingGame + 1
                            Credits = Credits - 1
                            If Credits < 1 Then DOF 131, DOFOff
                            UpdateCredits
                            UpdateBallInPlay
                            UpdatePLayersCanPlay
                        Else
                        ' Not Enough Credits to start a game.
                        'PlaySound "so_nocredits"
                        End If
                    End If
                End If
            End If
            Else

                If keycode = StartGameKey Then
                    If(bFreePlay = True)Then
                        If(BallsOnPlayfield = 0)Then
                            ResetScores
                            ResetForNewGame()
                            UpdatePLayersCanPlay
                        End If
                    Else
                        If(Credits> 0)Then
                            If(BallsOnPlayfield = 0)Then
                                Credits = Credits - 1
                                DOF 131, DOFOff
                                UpdateCredits
                                ResetScores
                                ResetForNewGame()
                                UpdatePLayersCanPlay
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

'*******************
' Flipper Subs Rev3
'*******************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

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

' flippers top animations

Sub LeftFlipper_Animate: LeftFlipperTop.RotZ = LeftFlipper.CurrentAngle: End Sub
Sub RightFlipper_Animate: RightFlipperTop.RotZ = RightFlipper.CurrentAngle: End Sub

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

    '***********
    ' GI lights
    '***********

    Sub GiOn 'enciende las luces GI
        Dim bulb
  PlaySound"fx_gion"
        For each bulb in aGiLights
            bulb.State = 1
        Next
    End Sub

    Sub GiOff 'apaga las luces GI
        Dim bulb
  PlaySound"fx_gioff"
        For each bulb in aGiLights
            bulb.State = 0
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
            TiltReel.SetValue 1
            If B2SOn then
                Controller.B2SSetTilt 1
            end if
            DisableTable True
            ' BallsRemaining(CurrentPlayer) = 0 'player looses the game
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
            Bumper003.Threshold = 100
            LeftSlingshot.Disabled = 1
            RightSlingshot.Disabled = 1
        Else
            GiOn
            Bumper001.Threshold = 1
            Bumper002.Threshold = 1
            Bumper003.Threshold = 1
            LeftSlingshot.Disabled = 0
            RightSlingshot.Disabled = 0
        End If
    End Sub

    Sub TiltRecoveryTimer_Timer()
        ' all the balls have drained
        If(BallsOnPlayfield = 0)Then
            EndOfBall()
            TiltRecoveryTimer.Enabled = False
        End If
    ' otherwise repeat
    End Sub

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


    '***************************************
    '     Collection collision sounds
    '***************************************

    Sub aMetals_Hit(idx):PlaySoundAtBall "fx_MetalHit":End Sub
    Sub aRubber_Bands_Hit(idx):PlaySoundAtBall "fx_rubber_band":End Sub
    Sub aRubber_Posts_Hit(idx):PlaySoundAtBall "fx_rubber_post":End Sub
    Sub aRubber_Pins_Hit(idx):PlaySoundAtBall "fx_rubber_pin":End Sub
    Sub aPlastics_Hit(idx):PlaySoundAtBall "fx_PlasticHit":End Sub
    Sub aGates_Hit(idx):PlaySoundAtBall "fx_Gate":End Sub
    Sub aWoods_Hit(idx):PlaySoundAtBall "fx_Woodhit":End Sub

    '************************************************************************************************************************
    ' Only for VPX 10.2 and higher.
    ' FlashForMs will blink light or a flasher for TotalPeriod(ms) at rate of BlinkPeriod(ms)
    ' When TotalPeriod done, light or flasher will be set to FinalState value where
    ' Final State values are:   0=Off, 1=On, 2=Return to previous State
    '************************************************************************************************************************

    Sub FlashForMs(MyLight, TotalPeriod, BlinkPeriod, FinalState)

        If TypeName(MyLight) = "Light" Then
            If FinalState = 2 Then
                FinalState = MyLight.State
            End If
            MyLight.BlinkInterval = BlinkPeriod
            MyLight.Duration 2, TotalPeriod, FinalState
        ElseIf TypeName(MyLight) = "Flasher" Then
            Dim steps
            steps = Int(TotalPeriod / BlinkPeriod + .5)
            If FinalState = 2 Then
                FinalState = ABS(MyLight.Visible)
            End If
            MyLight.UserValue = steps * 10 + FinalState
            MyLight.TimerInterval = BlinkPeriod
            MyLight.TimerEnabled = 0
            MyLight.TimerEnabled = 1
            ExecuteGlobal "Sub " & MyLight.Name & "_Timer:" & "Dim tmp, steps, fstate:tmp=me.UserValue:fstate = tmp MOD 10:steps= tmp\10 -1:Me.Visible = steps MOD 2:me.UserValue = steps *10 + fstate:If Steps = 0 then Me.Visible = fstate:Me.TimerEnabled=0:End if:End Sub"
        End If
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
        ResetNewBallLights
        ResetNewBallVariables
    End Sub

    ' Crete new ball

    Sub CreateNewBall()
        BallRelease.CreateSizedBallWithMass BallSize / 2, BallMass
        BallsOnPlayfield = BallsOnPlayfield + 1
        UpdateBallInPlay
        PlaySoundAt SoundFXDOF("fx_Ballrel", 110, DOFPulse, DOFContactors), BallRelease
        BallRelease.Kick 90, 4
        ' only for this table
        vpmtimer.addtimer 1000, "ShootAgainL.Visible = 0 '"
    End Sub

    ' player lost the ball

    Sub EndOfBall()
        'debug.print "EndOfBall"
        Dim AwardPoints, TotalBonus, ii
        AwardPoints = 0
        TotalBonus = 0
        ' Lost the first ball, now it cannot accept more players
        bOnTheFirstBall = False

        ' add bonus if no tilt
        ' tilt system will take care of the next ball

        If NOT Tilted Then
            Select Case BonusMultiplier
                Case 1:BonusCountTimer.Interval = 250
                Case 2:BonusCountTimer.Interval = 400
            End Select
            BonusCountTimer.Enabled = 1
        Else
            vpmtimer.addtimer 400, "EndOfBall2 '"
        End If
    End Sub

    Sub BonusCountTimer_Timer
        'debug.print "BonusCount_Timer"
        If Bonus> 0 Then
            Bonus = Bonus -1
            AddScore 10000 * BonusMultiplier
            UpdateBonusLights
        Else
            BonusCountTimer.Enabled = 0
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
    End Sub

    ' After bonus count go to the next step
    '
    Sub EndOfBall2()
        'debug.print "EndOfBall2"

        Tilted = False
        Tilt = 0
        TiltReel.SetValue 0
        If B2SOn then
            Controller.B2SSetTilt 0
        end if
        DisableTable False

        ' win extra ball?
        If(ExtraBallsAwards(CurrentPlayer)> 0)Then
            'debug.print "Extra Ball"

            ' if so then give it
            ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer)- 1

            ' turn off light if no more extra balls
            If(ExtraBallsAwards(CurrentPlayer) = 0)Then
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

            BallsRemaining(CurrentPlayer) = BallsRemaining(CurrentPlayer)- 1

            ' last ball?
            If(BallsRemaining(CurrentPlayer) <= 0)Then
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
        If(PlayersPlayingGame> 1)Then
            NextPlayer = CurrentPlayer + 1
            ' if it is the last player then go to the first one
            If(NextPlayer> PlayersPlayingGame)Then
                NextPlayer = 1
            End If
        Else
            NextPlayer = CurrentPlayer
        End If

        'debug.print "Next Player = " & NextPlayer

        ' end of game?
        If((BallsRemaining(CurrentPlayer) <= 0)AND(BallsRemaining(NextPlayer) <= 0))Then

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
            Controller.B2SSetData 110, 0 'Jackpot
        end if
        ' turn off flippers
        SolLFlipper 0
        SolRFlipper 0
        ' only this table
        ShootAgainL.Visible = 0

        ' start the attract mode
        StartAttractMode
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
                If sortscores(sj)> sortscores(sj + 1)then
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
        Match = INT(RND(1) * 10) * 100 ' random between 10 and 90
        Display_Match
        If(Score(CurrentPlayer)MOD 1000) = Match Then
            PlaySound SoundFXDOF  ("fx_knocker",300,DOFPulse,DOFKnocker)
            DOF 130, DOFPulse
            DOF 131, DOFOn
            AddCredits 1
        End If
    End Sub

    Sub Clear_Match()
        MatchReel.SetValue 0
        If B2SOn then
            Controller.B2SSetMatch 34, 100
        end if
    End Sub

    Sub Display_Match()
        MatchReel.SetValue(Match \ 100)
        If B2SOn then
            If Match = 0 then
                Controller.B2SSetMatch 34, 100
            else
                Controller.B2SSetMatch 34, Match\10
            end if
        end if
    End Sub

    ' *********************************************************************
    '                      Drain / Plunger Functions
    ' *********************************************************************

    Sub Drain_Hit()
        Drain.DestroyBall
        BallsOnPlayfield = BallsOnPlayfield - 1
        PlaySoundAt "fx_drain", Drain
        DOF 231, DOFPulse
        ' turn off Gi lights
        GiOff

        'tilted?
        If Tilted Then
            StopEndOfBallMode
        End If

        ' if still playing and not tilted
        If(bGameInPLay = True)AND(Tilted = False)Then

            ' ballsaver?
            If(bBallSaverActive = True)Then
                CreateNewBall()
            Else
                ' last ball?
                If(BallsOnPlayfield = 0)Then
                    ' only this table
                    'ShootAgainL.Visible = 1
                    StopEndOfBallMode
                    vpmtimer.addtimer 500, "EndOfBall '"
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
            Case 10, 100, 1000, 10000
                Score(CurrentPlayer) = Score(CurrentPlayer) + points
                UpdateScore points
                ' sounds
                If Points = 100 AND(Score(CurrentPlayer)MOD 10000) \ 1000 = 0 Then  'new reel 10000
                    PlaySound SoundFXDOF("10000",137,DOFPulse,DOFChimes)
                ElseIf Points = 10 AND(Score(CurrentPlayer)MOD 1000) \ 100 = 0 Then 'new reel 1000
                    PlaySound SoundFXDOF("1000",136,DOFPulse,DOFChimes)
                Else
                    PlaySound SoundFXDOF(Points,135,DOFPulse,DOFChimes)
                End If
            Case 20, 30, 40, 50
                Add10 = Add10 + Points \ 10
                AddScore10Timer.Enabled = TRUE
            Case 200, 300, 400, 500
                Add100 = Add100 + Points \ 100
                AddScore100Timer.Enabled = TRUE
            Case 2000, 3000, 4000, 5000
                Add1000 = Add1000 + Points \ 1000
                AddScore1000Timer.Enabled = TRUE
            Case 20000, 30000, 40000, 50000
                Add10000 = Add10000 + Points \ 10000
                AddScore10000Timer.Enabled = TRUE
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

    Sub AddScore1000Timer_Timer()
        if Add1000> 0 then
            AddScore 1000
            Add1000 = Add1000 - 1
        Else
            Me.Enabled = FALSE
        End If
    End Sub

    Sub AddScore10000Timer_Timer()
        if Add10000> 0 then
            AddScore 10000
            Add10000 = Add10000 - 1
        Else
            Me.Enabled = FALSE
        End If
    End Sub

    '*******************
    '     BONUS
    '*******************

    Sub AddBonus(bonuspoints)
        If(Tilted = False)Then
            Bonus = Bonus + bonuspoints
            PlaySound "Bell2"
            If Bonus> 10 Then
                Bonus = 10
            End If
            UpdateBonusLights
        End if
    End Sub

    Sub AddBonusMultiplier(multi)
        If(Tilted = False)Then
            BonusMultiplier = BonusMultiplier + multi
            If BonusMultiplier> MaxMultiplier Then
                BonusMultiplier = MaxMultiplier
            End If
            UpdateMultiplierLights
        End if
    End Sub

    Sub UpdateMultiplierLights
        Select Case BonusMultiplier
            Case 1:doublebonusLight.State = 0
            Case 2:doublebonusLight.State = 1
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
            If Score(CurrentPlayer) >= 100000 then
                Controller.B2SSetScoreRollover 24 + CurrentPlayer, 1
            end if
        end if
    End Sub

    Sub ResetScores
        ScoreReel1.ResetToZero
        ScoreReel2.ResetToZero
        ScoreReel3.ResetToZero
        ScoreReel4.ResetToZero
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

    Sub AddCredits(value) 'limit to 15 credits
        If Credits <20 Then
            Credits = Credits + value
            CreditsReel.SetValue Credits
            UpdateCredits
        end if
    End Sub

    Sub UpdateCredits
        'If Credits> 0 Then 'in Bally tables
        '    CreditLight.State = 1
        'Else
        '    CreditLight.State = 0
        'End If
        PlaySound "fx_relay"
        CreditsReel.SetValue credits
        If B2SOn then
            Controller.B2SSetCredits Credits
        end if
    End Sub

    Sub UpdateBallInPlay
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
    End Sub

    Sub UpdatePlayersCanPlay
        Select Case PlayersPlayingGame
            Case 0:cp1.State = 0:cp2.State = 0:cp3.State = 0:cp4.State = 0
            Case 1:cp1.State = 1:cp2.State = 0:cp3.State = 0:cp4.State = 0
            Case 2:cp1.State = 1:cp2.State = 1:cp3.State = 0:cp4.State = 0
            Case 3:cp1.State = 1:cp2.State = 1:cp3.State = 1:cp4.State = 0
            Case 4:cp1.State = 1:cp2.State = 1:cp3.State = 1:cp4.State = 1
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
            PlaySound SoundFXDOF  ("fx_knocker",300,DOFPulse,DOFKnocker)
            DOF 130, DOFPulse
            ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) + 1
            bExtraBallWonThisBall = True
        Else
            Addscore 5000
        END If
    End Sub

    Sub AwardSpecial()
        PlaySound SoundFXDOF  ("fx_knocker",300,DOFPulse,DOFKnocker)
        DOF 130, DOFPulse
        DOF 131, DOFOn
        AddCredits 1
    End Sub

    ' ********************************
    '        Attract Mode
    ' ********************************
    ' use the"Blink Pattern" of each light

    Sub StartAttractMode()
        Dim x
        bAttractMode = True
        For each x in aLights
            x.State = 2
        Next
        GameOverR.SetValue 1
        UpdateBallInPlay
    End Sub

    Sub StopAttractMode()
        Dim x
        bAttractMode = False
        TurnOffPlayfieldLights
        ResetScores
        GameOverR.SetValue 0
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
        If Not FileObj.FolderExists(UserDirectory)then
            Credits = 0
            Exit Sub
        End If
        If Not FileObj.FileExists(UserDirectory & TableName& ".txt")then
            Credits = 0
            Exit Sub
        End If
        Set ScoreFile = FileObj.GetFile(UserDirectory & TableName& ".txt")
        Set TextStr = ScoreFile.OpenAsTextStream(1, 0)
        If(TextStr.AtEndOfStream = True)then
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
        If Credits > 0 Then DOF 131, DOFOn
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
        If Not FileObj.FolderExists(UserDirectory)then
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
                If HSScore(j) <HSScore(j + 1)Then
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

    '*******************
    ' Realtime updates
    '*******************

    Sub GameTimer_Timer
        RollingUpdate
    End Sub

    '***********************************************************************
    ' *********************************************************************
    '  *********     G A M E  C O D E  S T A R T S  H E R E      *********
    ' *********************************************************************
    '***********************************************************************

    Sub VPObjects_Init 'init objects
        TurnOffPlayfieldLights()
        Slot1 = INT(RND * 10)
        Slot2 = INT(RND * 10)
        Slot3 = INT(RND * 10)
        SlotReel1.SetValue Slot1
        SlotReel2.SetValue Slot2
        SlotReel3.SetValue Slot3
        If B2SOn Then
        Controller.B2SSetData 120 + Slot1, 1
        Controller.B2SSetData 140 + Slot2, 1
        Controller.B2SSetData 160 + Slot3, 1
        End If
    End Sub

    ' Dim all the variables
    Dim Slot1
    Dim Slot2
    Dim Slot3

    Sub Game_Init() 'called at the start of a new game
        'Start music?
        'Init variables?
        'Start or init timers
        'Init lights?
        TurnOffPlayfieldLights()
    End Sub

    Sub StopEndOfBallMode()     'called when the last ball is drained
    End Sub

    Sub ResetNewBallVariables() 'init variables new ball/player
        BonusMultiplier = 1
        Bonus = 1:UpdateBonusLights
    End Sub

    Sub ResetNewBallLights() 'init lights for new ball/player
        TurnOffPlayfieldLights
        LightBumper001.State = 1
        LightBumper002.State = 1
        light015.State = 1
    End Sub

    Sub TurnOffPlayfieldLights()
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

    ' Slingshots
    Dim LStep, RStep

    Sub LeftSlingShot_Hit
        If Tilted Then Exit Sub
    AddScore 5000
    LStep = 1
    LeftSlingShot.TimerEnabled = True
    End Sub

    Sub LeftSlingShot_Timer
        Select Case LStep
            Case 1:LeftSLing1.Visible = 0:LeftSLing2.Visible = 1
            Case 2:LeftSLing2.Visible = 0:LeftSLing3.Visible = 1
            Case 3:LeftSLing3.Visible = 0:LeftSLing1.Visible = 1:LeftSlingShot.TimerEnabled = 0
        End Select
        LStep = LStep + 1
    End Sub


    Sub RightSlingShot_Slingshot
        If Tilted Then Exit Sub
        PlaySoundAt SoundFXDOF("fx_slingshot",110,DOFPulse,DOFcontactors), Remk
        DOF 112, DOFPulse
        RightSling4.Visible = 1
        Remk.RotX = 26
        RStep = 0
        RightSlingShot.TimerEnabled = True
        ' add points
        AddScore 100
    ' add effect?
    End Sub

    Sub RightSlingShot_Timer
        Select Case RStep
            Case 1:RightSLing4.Visible = 0:RightSLing3.Visible = 1:Remk.RotX = 14
            Case 2:RightSLing3.Visible = 0:RightSLing2.Visible = 1:Remk.RotX = 2
            Case 3:RightSLing2.Visible = 0:Remk.RotX = -20:RightSlingShot.TimerEnabled = 0
        End Select
        RStep = RStep + 1
    End Sub

    '***********************
    '       Rubbers
    '***********************

    Dim Rub17, Rub15, Rub18, Rub16, Rub11, Rub13

    Sub RubberBand017_Hit
        If Tilted then Exit Sub
        AddScore 100
        Rub17 = 1:RubberBand017_Timer
    End Sub

    Sub RubberBand017_Timer
        Select Case Rub17
            Case 1:r007.Visible = 0:r008.Visible = 1:RubberBand017.TimerEnabled = 1
            Case 2:r008.Visible = 0:r009.Visible = 1
            Case 3:r009.Visible = 0:r007.Visible = 1:RubberBand017.TimerEnabled = 0
        End Select
        Rub17 = Rub17 + 1
    End Sub

    Sub RubberBand015_Hit
        If Tilted then Exit Sub
        AddScore 100
        Rub15 = 1:RubberBand015_Timer
    End Sub

    Sub RubberBand015_Timer
        Select Case Rub15
            Case 1:r010.Visible = 0:r011.Visible = 1:RubberBand015.TimerEnabled = 1
            Case 2:r011.Visible = 0:r012.Visible = 1
            Case 3:r012.Visible = 0:r010.Visible = 1:RubberBand015.TimerEnabled = 0
        End Select
        Rub15 = Rub15 + 1
    End Sub

    Sub RubberBand018_Hit
        If Tilted then Exit Sub
        AddScore 100
        Rub18 = 1:RubberBand018_Timer
    End Sub

    Sub RubberBand018_Timer
        Select Case Rub18
            Case 1:r016.Visible = 0:r017.Visible = 1:RubberBand018.TimerEnabled = 1
            Case 2:r017.Visible = 0:r018.Visible = 1
            Case 3:r018.Visible = 0:r016.Visible = 1:RubberBand018.TimerEnabled = 0
        End Select
        Rub18 = Rub18 + 1
    End Sub

    Sub RubberBand016_Hit
        If Tilted then Exit Sub
        AddScore 100
        Rub16 = 1:RubberBand016_Timer
    End Sub

    Sub RubberBand016_Timer
        Select Case Rub16
            Case 1:r019.Visible = 0:r020.Visible = 1:RubberBand016.TimerEnabled = 1
            Case 2:r020.Visible = 0:r021.Visible = 1
            Case 3:r021.Visible = 0:r019.Visible = 1:RubberBand016.TimerEnabled = 0
        End Select
        Rub16 = Rub16 + 1
    End Sub

    Sub RubberBand011_Hit
        If Tilted then Exit Sub
        AddScore 100
        Rub11 = 1:RubberBand011_Timer
    End Sub

    Sub RubberBand011_Timer
        Select Case Rub11
            Case 1:r022.Visible = 0:r014.Visible = 1:RubberBand011.TimerEnabled = 1
            Case 2:r014.Visible = 0:r015.Visible = 1
            Case 3:r015.Visible = 0:r022.Visible = 1:RubberBand011.TimerEnabled = 0
        End Select
        Rub11 = Rub11 + 1
    End Sub

    Sub RubberBand013_Hit
        If Tilted then Exit Sub
        AddScore 100
        Rub13 = 1:RubberBand013_Timer
    End Sub

    Sub RubberBand013_Timer
        Select Case Rub13
            Case 1:r003.Visible = 0:r023.Visible = 1:RubberBand013.TimerEnabled = 1
            Case 2:r023.Visible = 0:r024.Visible = 1
            Case 3:r024.Visible = 0:r003.Visible = 1:RubberBand013.TimerEnabled = 0
        End Select
        Rub13 = Rub13 + 1
    End Sub

    'rubbers without animation

    'Sub RubberBand020_Hit:Addscore 100:End Sub
    'Sub RubberBand019_Hit:Addscore 100:End Sub
    Sub RubberBand024_Hit:Addscore 100:End Sub
    Sub RubberBand023_Hit:Addscore 100:End Sub
    Sub RubberBand014_Hit:Addscore 100:End Sub
    Sub RubberBand025_Hit:Addscore 100:End Sub

    '*********
    ' Bumpers
    '*********

    Sub Bumper001_Hit
        If Tilted Then Exit Sub
        PlaySoundAt SoundFXDOF ("fx_Bumper2",103,DOFPulse,DOFContactors), bumper001
        DOF 106, DOFPulse
        AddScore 100 + 900 * LightBumper001.State
    End Sub

    Sub Bumper002_Hit
        If Tilted Then Exit Sub
        PlaySoundAt SoundFXDOF ("fx_Bumper2",104,DOFPulse,DOFContactors), bumper002
        DOF 107, DOFPulse
        AddScore 100 + 900 * LightBumper002.State
    End Sub

    Sub Bumper003_Hit
        If Tilted Then Exit Sub
        PlaySoundAt SoundFXDOF ("fx_Bumper2",105,DOFPulse,DOFContactors), bumper003
        DOF 108, DOFPulse
        AddScore 1000 + 9000 * LightBumper003.State
    End Sub

    ' Setas, turn on red bumpers

    Sub ch1_Hit
        If Tilted Then Exit Sub
        DOF 120, DOFPulse
        'PlaySoundAt "fx_passive_bumper", ch1p
        AddScore 100
    End Sub

    Sub ch2_Hit
        If Tilted Then Exit Sub
        DOF 121, DOFPulse
        'PlaySoundAt "fx_passive_bumper", ch2p
        AddScore 100
    End Sub

    '*****************
    '     Lanes
    '*****************

    ' outlanes

    Sub Trigger001_Hit
        DOF 122, DOFPulse
        PlaySoundAt "fx_sensor", Trigger001
        If Tilted Then Exit Sub
        'check?
        If light001.State Then
      AwardSpecial
    Else
      AddScore 10000
    End If
    End Sub

    Sub Trigger003_Hit
        DOF 124, DOFPulse
        PlaySoundAt "fx_sensor", Trigger003
        If Tilted Then Exit Sub
        'check?
        If light002.State Then
      AwardSpecial
    Else
      AddScore 10000
    End If
    End Sub


    ' inlanes

    Sub Trigger002_Hit
        DOF 123, DOFPulse
        PlaySoundAt "fx_sensor", Trigger002
        If Tilted Then Exit Sub
        AddScore 5000 + 25000 * light015.State
    End Sub

    ' DOF backboard lights ball entrance

    Sub Trigger004_Hit
        DOF 230, DOFPulse
    End Sub

    '************************
    '       Targets
    '************************

    Sub Target001_hit
        PlaySoundAtBall SoundFXDOF("fx_target",116,DOFPulse,DOFTargets)
        If Tilted Then Exit Sub
        AddScore 10000
        AddBonus 1
    End Sub

    Sub Target002_hit
        PlaySoundAtBall SoundFXDOF("fx_target",116,DOFPulse,DOFTargets)
        If Tilted Then Exit Sub
        AddScore 10000
        AddBonus 1
    End Sub

    Sub Target003_hit
        PlaySoundAtBall SoundFXDOF("fx_target",117,DOFPulse,DOFTargets)
        If Tilted Then Exit Sub
        AddScore 10000
        AddBonus 1
    End Sub

    Sub Target004_hit
        PlaySoundAtBall SoundFXDOF("fx_target",117,DOFPulse,DOFTargets)
        If Tilted Then Exit Sub
        AddScore 10000
        AddBonus 1
    End Sub

    Sub Target005_hit
        PlaySoundAtBall SoundFXDOF("fx_target",118,DOFPulse,DOFTargets)
        If Tilted Then Exit Sub
        AddScore 10000
        AddBonus 1
    End Sub

    Sub Target006_hit 'on green bumper
        PlaySoundAtBall SoundFXDOF("fx_target",118,DOFPulse,DOFTargets)
        DOF 232, DOFPulse
        If Tilted Then Exit Sub
        AddScore 5000
        LightBumper003.State = 1
        LightBumper001.State = 0
        LightBumper002.State = 0
        light015.State = 0
    End Sub

    Sub Target007_hit
        PlaySoundAtBall SoundFXDOF("fx_target",119,DOFPulse,DOFTargets)
        If Tilted Then Exit Sub
        AddScore 5000 + 25000 * light014.State
    End Sub

    '*******************
    ' Rollover - bonus
    '*******************

    Sub sw001_Hit 'add bonus
        PlaySoundAt "fx_sensor", sw001
        DOF 231, DOFPulse
        If Tilted Then Exit Sub
        AddScore 10000
        AddBonus 1
    End Sub

    Sub sw002_Hit 'off green bumper
        PlaySoundAt "fx_sensor", sw002
        DOF 232, DOFPulse
        If Tilted Then Exit Sub
        AddScore 5000
        LightBumper003.State = 0
        LightBumper001.State = 1
        LightBumper002.State = 1
        light015.State = 1
    End Sub

    '*****************
    '    kickers
    '*****************

    Sub kicker1_Hit 'advance right reel
        PlaySoundAt "fx_kicker_enter", kicker1
            If Tilted Then kicker1.kick 170, 18:Exit Sub
        AdvanceRightReel
        vpmtimer.addtimer 900, "AddScore 1000:DOF 221, DOFPulse:DOF 200, DOFPulse:DOF 114, DOFPulse:PlaySoundAt SoundFX(""fx_kicker"",DOFContactors), kicker1: kicker1.kick 170, 20 '"
    End Sub

    Sub kicker2_Hit 'advance left reel
        PlaySoundAt "fx_kicker_enter", kicker2
        If Tilted Then kicker2.kick 180, 14:Exit Sub
        AdvanceLeftReel
        vpmtimer.addtimer 900, "AddScore 1000:DOF 216, DOFPulse:DOF 200, DOFPulse:DOF 113, DOFPulse:PlaySoundAt SoundFX(""fx_kicker"",DOFContactors), kicker2: kicker2.kick 190, 20 '"
    End Sub

    Sub kicker3_Hit 'advance 3 reels
        PlaySoundAt "fx_kicker_enter", kicker3
        If Tilted Then kicker3.kick 180, 14:Exit Sub
        AdvanceReels
        vpmtimer.addtimer 900, "AddScore 1000:DOF 217, DOFPulse:DOF 200, DOFPulse:DOF 113, DOFPulse:PlaySoundAt SoundFX(""fx_kicker"",DOFContactors), kicker3: kicker3.kick 40, 14 '"
    End Sub

    Sub kicker4_Hit 'advance 3 reels
        PlaySoundAt "fx_kicker_enter", kicker4
        If Tilted Then kicker4.kick 320, 16:Exit Sub
        AdvanceReels
        vpmtimer.addtimer 900, "AddScore 1000:DOF 219, DOFPulse:DOF 200, DOFPulse:DOF 114, DOFPulse:PlaySoundAt SoundFX(""fx_kicker"",DOFContactors), kicker4: kicker4.kick 320, 14 '"
    End Sub

    Sub kicker5_Hit 'advance central reel
        PlaySoundAt "fx_kicker_enter", kicker5
        If Tilted Then kicker5.kick 320, 20:Exit Sub
        AdvanceCentralReel
        vpmtimer.addtimer 900, "AddScore 1000:DOF 218, DOFPulse:DOF 200, DOFPulse:DOF 115, DOFPulse:PlaySoundAt SoundFX(""fx_kicker"",DOFContactors), kicker5: kicker5.kick 320, 20 '"
    End Sub

    '*****************************
    '     Table specific subs
    '*****************************

    Dim AddSlot1, AddSlot2, AddSlot3

    Sub B2SSlotReels
        Dim i
        For i = 120 to 129
            Controller.B2SSetData i, 0
        Next
        For i = 140 to 149
            Controller.B2SSetData i, 0
        Next
        For i = 160 to 169
            Controller.B2SSetData i, 0
        Next
        Controller.B2SSetData 120 + ((Slot1 + AddSlot1)MOD 10), 1
        Controller.B2SSetData 140 + ((Slot2 + AddSlot2)MOD 10), 1
        Controller.B2SSetData 160 + ((Slot3 + AddSlot3)MOD 10), 1
    End Sub

    Sub AdvanceRightReel 'slot 3
        AddSlot3 = INT(RND(1) * 3 + 1)
        Slot3 = (Slot3 + AddSlot3)MOD 10
        AddSlot3 = AddSlot3 + 10:Slot3T.Enabled = 1
        vpmtimer.addtimer 1000, "CheckReels '"
    End Sub

    Sub AdvanceLeftReel 'slot 1
        AddSlot1 = INT(RND(1) * 3 + 1)
        Slot1 = (Slot1 + AddSlot1)MOD 10
        AddSlot1 = AddSlot1 + 10:Slot1T.Enabled = 1
        vpmtimer.addtimer 1000, "CheckReels '"
    End Sub

    Sub AdvanceCentralReel 'slot 2
        AddSlot2 = INT(RND(1) * 3 + 1)
        Slot2 = (Slot2 + AddSlot2)MOD 10
        AddSlot2 = AddSlot2 + 10:Slot2T.Enabled = 1
        vpmtimer.addtimer 1000, "CheckReels '"
    End Sub

    Sub AdvanceReels
        AddSlot1 = INT(RND(1) * 3 + 1)
        Slot1 = (Slot1 + AddSlot1)MOD 10
        AddSlot1 = AddSlot1 + 10:Slot1T.Enabled = 1
        AddSlot2 = INT(RND(1) * 3 + 1)
        Slot2 = (Slot2 + AddSlot2)MOD 10
        AddSlot2 = AddSlot2 + 10:Slot2T.Enabled = 1
        AddSlot3 = INT(RND(1) * 3 + 1)
        Slot3 = (Slot3 + AddSlot3)MOD 10
        AddSlot3 = AddSlot3 + 10:Slot3T.Enabled = 1
        vpmtimer.addtimer 1000, "CheckReels '"
    End Sub

    Sub Slot1T_Timer()
        if AddSlot1> 0 then
            SlotReel1.AddValue 1
            AddSlot1 = AddSlot1 - 1
            If B2SOn Then B2SSlotReels
            Else
                Me.Enabled = FALSE
        End If
    End Sub

    Sub Slot2T_Timer()
        if AddSlot2> 0 then
            SlotReel2.AddValue 1
            AddSlot2 = AddSlot2 - 1
            If B2SOn Then B2SSlotReels
            Else
                Me.Enabled = FALSE
        End If
    End Sub

    Sub Slot3T_Timer()
        if AddSlot3> 0 then
            SlotReel3.AddValue 1
            AddSlot3 = AddSlot3 - 1
            If B2SOn Then B2SSlotReels
            Else
                Me.Enabled = FALSE
        End If
    End Sub

    Sub CheckReels 'check for valid award
        '777
        If(Slot1 = 0 OR Slot1 = 6)AND(Slot2 = 0 OR Slot2 = 3 OR Slot2 = 6)AND(Slot3 = 0 OR Slot3 = 6)Then
            AwardSpecial:AwardSpecial                  '2 credits
            JackpotR.SetValue 1
            If B2SOn Then Controller.B2SSetData 110, 1 'Jackpot
            vpmtimer.addtimer 2000, "JackpotR.SetValue 0:If B2SOn Then Controller.B2SSetData 110, 0 '"
        End if
        'cerezas 'lites special and 30000
        If(Slot1 = 1 OR Slot1 = 9)AND(Slot2 = 1 OR Slot2 = 9)AND(Slot3 = 1 OR Slot3 = 9)Then
            light001.State = 1
            light002.State = 1
            light014.State = 1
        End If
        ' naranjas  'lites special and 30000
        If(Slot1 = 2 OR Slot1 = 4)AND(Slot2 = 2 OR Slot2 = 4)AND(Slot3 = 2 OR Slot3 = 4)Then
            light001.State = 1
            light002.State = 1
            light014.State = 1
        End If
        ' campanas ' double bonus
        If(Slot1 = 3 OR Slot1 = 5 OR Slot1 = 7)AND(Slot2 = 5 OR Slot2 = 7 OR Slot2 = 8)AND(Slot3 = 3 OR Slot3 = 5 OR Slot3 = 7)Then
            doublebonusLight.State = 1
            BonusMultiplier = 2
        End If
        ' cerezas + lemon
        If(Slot1 = 1 OR Slot1 = 9)AND(Slot2 = 1 OR Slot2 = 9)AND(Slot3 = 8)Then
            Addscore 30000
        End If
    End Sub


    ' ***************************************************
    ' GNMOD - Multiple High Score Display and Collection
    ' jpsalas: changed ramps by flashers
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

    ' default high scores, remove this when the scores are available from the config file
    HSScore(1) = 250000
    HSScore(2) = 240000
    HSScore(3) = 230000
    HSScore(4) = 220000
    HSScore(5) = 210000

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

        For xFor = StartHSArray(LineNo)to EndHSArray(LineNo)
            Eval("HS" &xFor).imageA = GetHSChar(String, Index)
            Index = Index + 1
        Next
    End Sub

    Sub NewHighScore(NewScore, PlayNum)
        If NewScore> HSScore(5)then
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
            If AlphaStringPos> len(AlphaString)or(AlphaStringPos = len(AlphaString)and InitialString = "")then
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
                InitialString = MID(InitialString, 1, len(InitialString)- 1)
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
                If i = 1 or(HSNewHigh> HSScore(i)and HSNewHigh <= HSScore(i - 1))then
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
                    PlaySound("fx_Bong")
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

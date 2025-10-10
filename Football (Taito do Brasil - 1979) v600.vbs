' ****************************************************************
'               VISUAL PINBALL X EM Script by JPSalas
'                  Basic EM script up to 4 players
'             uses core.vbs for extra functions
'
'
'  VPX8 - version by JPSalas 2024, version 6.0.0
' ****************************************************************

Option Explicit
Randomize

' DOF config - Foxyt - leeoneil
'
' Option for more lights effects with DOF (Undercab effects on bumpers)
' "True" to activate (False by default)
Const Epileptikdof = False


Const BallSize = 50
Const BallMass = 1

LoadCoreFiles

Sub LoadCoreFiles
    On Error Resume Next
    ExecuteGlobal GetTextFile("core.vbs")
    If Err Then MsgBox "Can't open core.vbs"
    On Error Resume Next
    ExecuteGlobal GetTextFile("controller.vbs")
    If Err Then MsgBox "Can't open controller.vbs"
End Sub

' Constants
Const TableName = "Football_1979_2"
Const cGameName = "Football_1979_2"
Const MaxPlayers = 4
Const Special1 = 620000 ' extra credits
Const Special2 = 790000
Const Special3 = 960000

' Global Variables
Dim PlayersPlayingGame
Dim CurrentPlayer
Dim Credits
Dim Bonus(4)
Dim BonusValue
Dim BonusMultiplier
Dim BonusCount
Dim BallsRemaining(4)
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
Dim Add1, Add10, Add100, Add1000, Add10000

' Controll Variables
Dim BallsOnPlayfield

' Boolean Variables
Dim bAttractMode
Dim bFreePlay
Dim bGameInPlay
Dim bOnTheFirstBall
Dim bExtraBallWonThisBall
Dim bJustStarted
Dim bBallInPlungerLane
Dim bBallSaverActive

Dim x

' core.vbs variables

Sub Table1_Init()
    VPObjects_Init
    LoadEM
    Loadhs
    UpdateCredits
    bFreePlay = False

    ' Init global variables
    bAttractMode = False
    bOnTheFirstBall = False
    bGameInPlay = False
    bBallInPlungerLane = False
    BallsOnPlayfield = 0
    Tilt = 0
    TiltSensitivity = 6
    Tilted = False
    bJustStarted = True
    Add1 = 0
    Add10 = 0
    Add100 = 0
    Add1000 = 0
    Add10000 = 0
    BonusValue = 2000

    EndOfGame

    vpmtimer.addtimer 1500, "GiOn '"

    ' Remove desktop elements
    If Table1.ShowDT = true then
        For each x in aReels
            x.Visible = 1
        Next
    else
        For each x in aReels
            x.Visible = 0
        Next
    end if

    ' Start realtime animations
    GameTimer.Enabled = 1
End Sub

'******
' Keys
'******

Sub Table1_KeyDown(ByVal Keycode)

    If Keycode = AddCreditKey OR Keycode = AddCreditKey2 Then
        If(Tilted = False)Then
            AddCredits 1
            PlaySoundAt "fx_coin", drain
            PlaySound "so_coin"
        End If
    End If

    If keycode = PlungerKey Then
        Plunger.Pullback
        PlaySoundAt "fx_plungerpull", plunger
    End If

    If bGameInPlay AND NOT Tilted Then
        ' teclas de la falta
        If keycode = LeftTiltKey Then Nudge 90, 8:PlaySound "fx_nudge", 0, 1, -0.1, 0.25:CheckTilt
        If keycode = RightTiltKey Then Nudge 270, 8:PlaySound "fx_nudge", 0, 1, 0.1, 0.25:CheckTilt
        If keycode = CenterTiltKey Then Nudge 0, 9:PlaySound "fx_nudge", 0, 1, 1, 0.25:CheckTilt
        If keycode = MechanicalTilt Then CheckTilt

        If keycode = LeftFlipperKey Then SolLFlipper 1
        If keycode = RightFlipperKey Then SolRFlipper 1
        If keycode = StartGameKey Then
            If((PlayersPlayingGame < MaxPlayers)AND(bOnTheFirstBall = True))Then
                If(bFreePlay = True)Then
                    PlayersPlayingGame = PlayersPlayingGame + 1
                    PlaySound "start"
                Else
                    If(Credits > 0)then
                        PlayersPlayingGame = PlayersPlayingGame + 1
                        Credits = Credits - 1
                        UpdateCredits
                        UpdateBallInPlay
                        PlaySound "so_start"
                    Else
                        PlaySound "10000"
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
                        PlaySound "start"
                    End If
                Else
                    If(Credits > 0)Then
                        If(BallsOnPlayfield = 0)Then
                            Credits = Credits - 1
                            UpdateCredits
                            ResetScores
                            ResetForNewGame()
                            PlaySound "start"
                        End If
                    Else

                        PlaySound "10000"
                    End If
                End If
            End If
    End If
End Sub

Sub Table1_KeyUp(ByVal keycode)

    If bGameInPlay AND NOT Tilted Then
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

'*************
' Pause/exit
'*************

Sub table1_Paused
End Sub

Sub table1_unPaused
End Sub

Sub table1_Exit
    Savehs
    If B2SOn Then Controller.Stop
End Sub

'*******************
' Flipper Subs v3.0
'*******************

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

Sub LeftFlipper_Animate:LeftFlipperTop.RotZ = LeftFlipper.CurrentAngle: End Sub
Sub RightFlipper_Animate: RightFlipperTop.RotZ = RightFlipper.CurrentAngle: End Sub

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

'*******************
'    GI lights
'*******************

Sub GiOn 'enciende las luces GI
    Dim bulb
    For each bulb in aGiLights
        bulb.State = 1
        DOF 255, DOFOn
    Next
    If B2SOn Then Controller.B2SSetData 60, 1
End Sub

Sub GiOff 'apaga las luces GI
    Dim bulb
    For each bulb in aGiLights
        bulb.State = 0
        DOF 255, DOFOff
    Next
    If B2SOn Then Controller.B2SSetData 60, 0
End Sub

'**************
'    TILT
'**************

Sub CheckTilt
    Tilt = Tilt + TiltSensitivity
    TiltDecreaseTimer.Enabled = True
    If Tilt > 15 Then
        PlaySound "so_tilt"
        Tilted = True
        TiltL.State = 1
        If B2SOn then
            Controller.B2SSetTilt 1
        end if
        DisableTable True
        TiltRecoveryTimer.Enabled = True
    End If
End Sub

Sub TiltDecreaseTimer_Timer
    If Tilt > 0 Then
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
        Bumper1.Force = 0
        Bumper2.Force = 0
        DOF 101, DOFOff
        DOF 102, DOFOff
    Else
        GiOn
        Bumper1.Force = 10
        Bumper2.Force = 10
    End If
End Sub

Sub TiltRecoveryTimer_Timer()
    If(BallsOnPlayfield = 0)Then
        EndOfBall()
        TiltRecoveryTimer.Enabled = False
    End If
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
                ballvol = Vol(BOT(b)) * 5
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
' Ball to ball hit sound
'*****************************

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
'FlashForMs will blink a light for a total time of"TotalPeriod" every each miliseconds "BlinkPeriod"
'When "TotalPeriod" is finished, the light state will be set to the value of "FinalState"
'"FinalState" can be: 0=off, 1=on, 2=original state
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
' Init the table for a new game
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
        Bonus(i) = 0
    Next
    BonusMultiplier = 1
    UpdateBallInPlay
    Clear_Match
    Tilt = 0
    Game_Init()
    ' start music?
    vpmtimer.addtimer 2000, "FirstBall '"
End Sub

Sub FirstBall
    'debug.print "FirstBall"
    ResetForNewPlayerBall()
    CreateNewBall()
End Sub

Sub ResetForNewPlayerBall()
    'debug.print "ResetForNewPlayerBall"
    AddScore 0
    ' Bonus(CurrentPlayer) = 0
    BonusMultiplier = 1
    UpdateBonusLights
    bExtraBallWonThisBall = False
    ResetNewBallVariables
    ResetNewBallLights
End Sub

Sub CreateNewBall()
    BallRelease.CreateSizedBallWithMass BallSize / 2, BallMass
    BallsOnPlayfield = BallsOnPlayfield + 1
    UpdateBallInPlay
    PlaySoundAt SoundFXDOF ("fx_Ballrel", 104, DOFPulse, DOFContactors), BallRelease
    PlaySound "so_start"
    BallRelease.Kick 90, 4
End Sub

Sub EndOfBall()
    'debug.print "EndOfBall"
    Dim AwardPoints, ii
    AwardPoints = 0
    bOnTheFirstBall = False
    If NOT Tilted Then
        BonusCountTimer.Interval = 200
        CreditLight.State = 2 'taito tables
        BonusCountTimer.Enabled = 1
    Else
        vpmtimer.addtimer 400, "EndOfBall2 '"
    End If
End Sub

Sub BonusCountTimer_Timer
    'debug.print "BonusCount_Timer"
    If BonusCount > 0 Then
        BonusCount = BonusCount -1
        AddScore BonusValue * BonusMultiplier
        PlaySound "so_bonus"
        UpdateBonusLights
    Else
        CreditLight.State = 0 'Taito tables
    BonusCount = Bonus(CurrentPlayer)
        UpdateBonusLights
        BonusCountTimer.Enabled = 0
        vpmtimer.addtimer 1000, "EndOfBall2 '"
    End If
End Sub

Sub EndOfBall2()
    'debug.print "EndOfBall2"
    Tilted = False
    Tilt = 0
    TiltL.State = 0
    If B2SOn then
        Controller.B2SSetTilt 0
    end if
    DisableTable False

    If(ExtraBallsAwards(CurrentPlayer) > 0)Then
        'debug.print "Extra Ball"
        ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer)- 1
        If(ExtraBallsAwards(CurrentPlayer) = 0)Then
            LightShootAgain.State = 0
            If B2SOn then
                Controller.B2SSetShootAgain 0
            end if
        End If

        ' blink a extra ball light? dmd?
        ResetForNewPlayerBall()
        CreateNewBall()
    Else
        BallsRemaining(CurrentPlayer) = BallsRemaining(CurrentPlayer)- 1
        If(BallsRemaining(CurrentPlayer) <= 0)Then
            CheckHighScore()
        End If
        EndOfBallComplete()
    End If
End Sub

Sub EndOfBallComplete()
    'debug.print "EndOfBall Complete"
    Dim NextPlayer
    If(PlayersPlayingGame > 1)Then
        NextPlayer = CurrentPlayer + 1
        If(NextPlayer > PlayersPlayingGame)Then
            NextPlayer = 1
        End If
    Else
        NextPlayer = CurrentPlayer
    End If
    'debug.print "Next Player = " & NextPlayer
    If((BallsRemaining(CurrentPlayer) <= 0)AND(BallsRemaining(NextPlayer) <= 0))Then
        If bFreePlay = False Then
            Verification_Match
        End If
        EndOfGame()
    Else
        CurrentPlayer = NextPlayer
        AddScore 0
        ResetForNewPlayerBall()
        CreateNewBall()
    End If
End Sub

Sub EndOfGame()
    'debug.print "EndOfGame"
    PlaySound "so_gameover"
    bGameInPLay = False
    bJustStarted = False
    If B2SOn then
        Controller.B2SSetGameOver 1
        Controller.B2SSetBallInPlay 0
        Controller.B2SSetPlayerUp 0
        Controller.B2SSetCanPlay 0
    end if
    SolLFlipper 0
    SolRFlipper 0
    StartAttractMode
End Sub

' Balls left?
Function Balls
    Dim tmp
    tmp = BallsPerGame - BallsRemaining(CurrentPlayer) + 1
    If tmp > BallsPerGame Then
        Balls = BallsPerGame
    Else
        Balls = tmp
    End If
End Function

Sub CheckHighscore
    For x = 1 to PlayersPlayingGame
        If Score(x) > Highscore Then
            Highscore = Score(x)
            PlaySound SoundFXDOF ("fx_knocker", 130, DOFPulse, DOFKnocker)
            DOF 230, DOFPulse
            DOF 125, DOFOn
            NewL.State = 1
            If B2SOn then Controller.B2SSetData 59, 1 'STAT Change, New as ID 59 on B2S
            RecordL.State = 1
            If B2SOn then Controller.B2SSetData 60, 1 'STAT Change, Record as ID 60 on B2S
            'HighscoreReel.SetValue Highscore
            AddCredits 1
        End If
    Next
End Sub

'******************
'  Match - Loteria
'******************

Sub Verification_Match()
    'PlaySound "fx_match"
    Match = INT(RND * 10)
    If(Score(CurrentPlayer)MOD 10) = Match Then
        PlaySound SoundFXDOF ("fx_knocker", 130, DOFPulse, DOFKnocker)
            DOF 230, DOFPulse
            DOF 125, DOFOn
        AddCredits 1
        If B2SOn then Controller.B2SSetScore 8, Match 'STAT Change, Match 1 Digit 0-9 Score ID 8 at BG
    End If
End Sub

Sub Clear_Match()
    MatchReel.SetValue 0
    If B2SOn then Controller.B2SSetScore 8, 0 'STAT Change, Match 1 Digit 0-9 Score ID 8 at BG
End Sub

' ******************************
'    Drain / Plunger Functions
' ******************************

Sub Drain_Hit()
    Drain.DestroyBall
    BallsOnPlayfield = BallsOnPlayfield - 1
    PlaySoundAt "fx_drain", Drain
    DOF 256, DOFPulse

    If Tilted Then
        Exit Sub
    End If

    If bGameInPLay = True Then
        If(bBallSaverActive = True)Then
            CreateNewBall()
        Else
            If(BallsOnPlayfield = 0)Then
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
'start a ball save counter?
End Sub

' *******************
'   SCORE functions
' *******************

Sub AddScore(Points)
    If Tilted Then Exit Sub
    Score(CurrentPlayer) = Score(CurrentPlayer) + points
    UpdateScore points

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

'*******************
'     BONUS
'*******************

Sub AddBonus(bonuspoints)
    If(Tilted = False)Then
        Bonus(CurrentPlayer) = Bonus(CurrentPlayer) + bonuspoints
        If Bonus(CurrentPlayer) > 18 then
            Bonus(CurrentPlayer) = 18
        End if
    BonusCount = Bonus(CurrentPlayer)
        LightSeqGoal.UpdateInterval = 1
        LightSeqGoal.Play SeqClockRightOn, 30, 3
        UpdateBonusLights
    End if
End Sub

Sub AddBonusMultiplier(points)
    If(Tilted = False)Then
        li80.State = 1
        BonusMultiplier = BonusMultiplier + points
        If BonusMultiplier = 6 AND li42.State = 0 Then 'lit special
            If li119.State = 0 Then
                li129.State = 1
                li42.State = 1
            End If
        End If
        If BonusMultiplier > 3 Then BonusMultiplier = 5
        UpdateBonusLights
    End if
End Sub

Sub UpdateBonusLights
    Select Case BonusCount
        Case 0:b1.State = 0:b2.State = 0:b3.State = 0:b4.State = 0:b5.State = 0:b6.State = 0:b7.State = 0:b8.State = 0:b9.State = 0:b10.State = 0:b11.State = 0:b12.State = 0:b13.State = 0:b14.State = 0:b15.State = 0:b16.State = 0:b17.State = 0:b18.State = 0
        Case 1:b1.State = 1:b2.State = 0:b3.State = 0:b4.State = 0:b5.State = 0:b6.State = 0:b7.State = 0:b8.State = 0:b9.State = 0:b10.State = 0:b11.State = 0:b12.State = 0:b13.State = 0:b14.State = 0:b15.State = 0:b16.State = 0:b17.State = 0:b18.State = 0
        Case 2:b1.State = 1:b2.State = 1:b3.State = 0:b4.State = 0:b5.State = 0:b6.State = 0:b7.State = 0:b8.State = 0:b9.State = 0:b10.State = 0:b11.State = 0:b12.State = 0:b13.State = 0:b14.State = 0:b15.State = 0:b16.State = 0:b17.State = 0:b18.State = 0
        Case 3:b1.State = 1:b2.State = 1:b3.State = 1:b4.State = 0:b5.State = 0:b6.State = 0:b7.State = 0:b8.State = 0:b9.State = 0:b10.State = 0:b11.State = 0:b12.State = 0:b13.State = 0:b14.State = 0:b15.State = 0:b16.State = 0:b17.State = 0:b18.State = 0
        Case 4:b1.State = 1:b2.State = 1:b3.State = 1:b4.State = 1:b5.State = 0:b6.State = 0:b7.State = 0:b8.State = 0:b9.State = 0:b10.State = 0:b11.State = 0:b12.State = 0:b13.State = 0:b14.State = 0:b15.State = 0:b16.State = 0:b17.State = 0:b18.State = 0
        Case 5:b1.State = 1:b2.State = 1:b3.State = 1:b4.State = 1:b5.State = 1:b6.State = 0:b7.State = 0:b8.State = 0:b9.State = 0:b10.State = 0:b11.State = 0:b12.State = 0:b13.State = 0:b14.State = 0:b15.State = 0:b16.State = 0:b17.State = 0:b18.State = 0
        Case 6:b1.State = 1:b2.State = 1:b3.State = 1:b4.State = 1:b5.State = 1:b6.State = 1:b7.State = 0:b8.State = 0:b9.State = 0:b10.State = 0:b11.State = 0:b12.State = 0:b13.State = 0:b14.State = 0:b15.State = 0:b16.State = 0:b17.State = 0:b18.State = 0
        Case 7:b1.State = 1:b2.State = 1:b3.State = 1:b4.State = 1:b5.State = 1:b6.State = 1:b7.State = 1:b8.State = 0:b9.State = 0:b10.State = 0:b11.State = 0:b12.State = 0:b13.State = 0:b14.State = 0:b15.State = 0:b16.State = 0:b17.State = 0:b18.State = 0
        Case 8:b1.State = 1:b2.State = 1:b3.State = 1:b4.State = 1:b5.State = 1:b6.State = 1:b7.State = 1:b8.State = 1:b9.State = 0:b10.State = 0:b11.State = 0:b12.State = 0:b13.State = 0:b14.State = 0:b15.State = 0:b16.State = 0:b17.State = 0:b18.State = 0
        Case 9:b1.State = 1:b2.State = 1:b3.State = 1:b4.State = 1:b5.State = 1:b6.State = 1:b7.State = 1:b8.State = 1:b9.State = 1:b10.State = 0:b11.State = 0:b12.State = 0:b13.State = 0:b14.State = 0:b15.State = 0:b16.State = 0:b17.State = 0:b18.State = 0
        Case 10:b1.State = 1:b2.State = 1:b3.State = 1:b4.State = 1:b5.State = 1:b6.State = 1:b7.State = 1:b8.State = 1:b9.State = 1:b10.State = 1:b11.State = 0:b12.State = 0:b13.State = 0:b14.State = 0:b15.State = 0:b16.State = 0:b17.State = 0:b18.State = 0
        Case 11:b1.State = 1:b2.State = 1:b3.State = 1:b4.State = 1:b5.State = 1:b6.State = 1:b7.State = 1:b8.State = 1:b9.State = 1:b10.State = 1:b11.State = 1:b12.State = 0:b13.State = 0:b14.State = 0:b15.State = 0:b16.State = 0:b17.State = 0:b18.State = 0
        Case 12:b1.State = 1:b2.State = 1:b3.State = 1:b4.State = 1:b5.State = 1:b6.State = 1:b7.State = 1:b8.State = 1:b9.State = 1:b10.State = 1:b11.State = 1:b12.State = 1:b13.State = 0:b14.State = 0:b15.State = 0:b16.State = 0:b17.State = 0:b18.State = 0
        Case 13:b1.State = 1:b2.State = 1:b3.State = 1:b4.State = 1:b5.State = 1:b6.State = 1:b7.State = 1:b8.State = 1:b9.State = 1:b10.State = 1:b11.State = 1:b12.State = 1:b13.State = 1:b14.State = 0:b15.State = 0:b16.State = 0:b17.State = 0:b18.State = 0
        Case 14:b1.State = 1:b2.State = 1:b3.State = 1:b4.State = 1:b5.State = 1:b6.State = 1:b7.State = 1:b8.State = 1:b9.State = 1:b10.State = 1:b11.State = 1:b12.State = 1:b13.State = 1:b14.State = 1:b15.State = 0:b16.State = 0:b17.State = 0:b18.State = 0
        Case 15:b1.State = 1:b2.State = 1:b3.State = 1:b4.State = 1:b5.State = 1:b6.State = 1:b7.State = 1:b8.State = 1:b9.State = 1:b10.State = 1:b11.State = 1:b12.State = 1:b13.State = 1:b14.State = 1:b15.State = 1:b16.State = 0:b17.State = 0:b18.State = 0
        Case 16:b1.State = 1:b2.State = 1:b3.State = 1:b4.State = 1:b5.State = 1:b6.State = 1:b7.State = 1:b8.State = 1:b9.State = 1:b10.State = 1:b11.State = 1:b12.State = 1:b13.State = 1:b14.State = 1:b15.State = 1:b16.State = 1:b17.State = 0:b18.State = 0
        Case 17:b1.State = 1:b2.State = 1:b3.State = 1:b4.State = 1:b5.State = 1:b6.State = 1:b7.State = 1:b8.State = 1:b9.State = 1:b10.State = 1:b11.State = 1:b12.State = 1:b13.State = 1:b14.State = 1:b15.State = 1:b16.State = 1:b17.State = 1:b18.State = 0
        Case 18:b1.State = 1:b2.State = 1:b3.State = 1:b4.State = 1:b5.State = 1:b6.State = 1:b7.State = 1:b8.State = 1:b9.State = 1:b10.State = 1:b11.State = 1:b12.State = 1:b13.State = 1:b14.State = 1:b15.State = 1:b16.State = 1:b17.State = 1:b18.State = 1
    End Select
    Select Case BonusMultiplier
        Case 1:bm2.State = 0:bm3.State = 0:bm5.State = 0
        Case 2:bm2.State = 1:bm3.State = 0:bm5.State = 0
        Case 3:bm2.State = 0:bm3.State = 1:bm5.State = 0
        Case 5:bm2.State = 0:bm3.State = 0:bm5.State = 1:li121.State = 1:li71.State = 1
    End Select
End Sub

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

' pone todos los marcadores a 0
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
        Controller.B2SSetScoreRolloverPlayer1 0
        Controller.B2SSetScoreRolloverPlayer2 0
        Controller.B2SSetScoreRolloverPlayer3 0
        Controller.B2SSetScoreRolloverPlayer4 0
    end if
End Sub

Sub AddCredits(value)
    If Credits < 9 Then
        Credits = Credits + value
        UpdateCredits
    end if
End Sub

Sub UpdateCredits
    If Credits > 0 Then 'this is for Bally tables
    'CreditLight.State = 1
    DOF 125, DOFOn
    Else
    DOF 125, DOFOff
    'CreditLight.State = 0
    End If
    CreditReel.SetValue Credits
    If B2SOn then Controller.B2SSetScore 6, Credits 'STAT Change, Note: only 1 Digit (0-9)
End Sub

Sub UpdateBallInPlay
    BallInPlayReel.SetValue Balls
    If B2SOn then
        'Controller.B2SSetBallInPlay Balls
        Controller.B2SSetScore 5, Balls 'STAT Change
    end if
    Select Case CurrentPlayer
        Case 1:pl1.State = 1:pl2.State = 0:pl3.State = 0:pl4.State = 0
        Case 2:pl1.State = 0:pl2.State = 1:pl3.State = 0:pl4.State = 0
        Case 3:pl1.State = 0:pl2.State = 0:pl3.State = 1:pl4.State = 0
        Case 4:pl1.State = 0:pl2.State = 0:pl3.State = 0:pl4.State = 1
    End Select
    If B2SOn then
        Controller.B2SSetPlayerUp CurrentPlayer
    end if
    PlayersReel.SetValue PlayersPlayingGame
    If B2SOn then
        'Controller.B2SSetCanPlay PlayersPlayingGame
        Controller.B2SSetScore 7, PlayersPlayingGame 'STAT Change
    end if
End Sub

'*************************
' Special & Extra Ball
'*************************

Sub AwardExtraBall()
    If NOT bExtraBallWonThisBall Then
        PlaySound SoundFXDOF ("fx_knocker", 130, DOFPulse, DOFKnocker)
            DOF 230, DOFPulse
            DOF 125, DOFOn
        PlaySound "so_shootagain"
        ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) + 1
        bExtraBallWonThisBall = True
        LightShootAgain.State = 1
        If B2SOn then
            Controller.B2SSetShootAgain 1
        end if
    END If
End Sub

Sub AwardSpecial()
    PlaySound SoundFXDOF  ("fx_knocker", 130, DOFPulse, DOFKnocker)
        DOF 230, DOFPulse
        DOF 125, DOFOn
    PlaySound "so_special"
    AddCredits 1
End Sub

' ********************************
'        Attract Mode
' ********************************

Sub StartAttractMode()
    bAttractMode = True
    For each x in aLights
        x.State = 2
    Next
    GameOverL.State = 1
    AttractTimer.Enabled = 1
End Sub

Sub StopAttractMode()
    bAttractMode = False
    ResetScores
    For each x in aLights
        x.State = 0
    Next
    GameOverL.State = 0
    NewL.State = 0
    If B2SOn then Controller.B2SSetData 59, 0 'STAT Change, New as ID 59 on B2S
    RecordL.State = 0
    If B2SOn then Controller.B2SSetData 60, 0 'STAT Change, Record as ID 60 on B2S
    AttractTimer.Enabled = 0
End Sub

Dim AttractStep
AttractStep = 0

Sub AttractTimer_Timer 'shows highscore
    Select Case AttractStep
        Case 0
            AttractStep = 1
            ScoreReel1.SetValue Score(1)
            ScoreReel2.SetValue Score(2)
            ScoreReel3.SetValue Score(3)
            ScoreReel4.SetValue Score(4)
            RecordL.State = 0
            If B2SOn then
                Controller.B2SSetScorePlayer1 Score(1)
                Controller.B2SSetScorePlayer2 Score(2)
                Controller.B2SSetScorePlayer3 Score(3)
                Controller.B2SSetScorePlayer4 Score(4)
                Controller.B2SSetData 60, 0 'STAT Change, Record as ID 60 on B2S
            end if
        Case 1
            AttractStep = 0
            ScoreReel1.SetValue Highscore
            ScoreReel2.SetValue Highscore
            ScoreReel3.SetValue Highscore
            ScoreReel4.SetValue Highscore
            RecordL.State = 1
            If B2SOn then
                Controller.B2SSetScorePlayer1 Highscore
                Controller.B2SSetScorePlayer2 Highscore
                Controller.B2SSetScorePlayer3 Highscore
                Controller.B2SSetScorePlayer4 Highscore
                Controller.B2SSetData 60, 1 'STAT Change, Record as ID 60 on B2S
            end if
    End Select
End Sub

'*****************************
'   Load / Save / Highscore
'*****************************

Sub Loadhs
    x = LoadValue(TableName, "HighScore")
    If(x <> "")Then HighScore = CDbl(x)Else HighScore = 0 End If
    x = LoadValue(TableName, "Credits")
    If(x <> "")then Credits = CInt(x)Else Credits = 0 End If
End Sub

Sub Savehs
    SaveValue TableName, "HighScore", HighScore
    SaveValue TableName, "Credits", Credits
End Sub

Sub Reseths
    HighScore = 0
    Savehs
End Sub

'*******************
' Real time updates
'*******************
' rolling sound, flipper tops, animations, gates

Sub GameTimer_Timer
    RollingUpdate
End Sub

'***********************************************************************
' *********************************************************************
'               TABLE SPECIFIC CODE STARTS HERE
' *********************************************************************
'***********************************************************************

Sub VPObjects_Init 'mostly dropwalls
End Sub

' tables own variables
Dim Advances
Dim sw54hits
Dim sw53hits
Dim kickerPos   ' only one ball so we only need one variable for the animations
Dim Goalready

Sub Game_Init() 'called at the start of a new game
    Dim i
    'start music?
    'init variables
    Advances = 0
    sw54hits = 0
    sw53hits = 0
    'statrt a timer?
    'Init some lights?
    TurnOffPlayfieldLights()
End Sub

Sub ResetNewBallVariables()
    Goalready = False
    BonusMultiplier = 1
    Advances = 0
    sw54hits = 0
    sw53hits = 0
End Sub

Sub ResetNewBallLights()
    TurnOffPlayfieldLights()
    UpdateBonusLights
    UpdateAdvances
    ResetGoalLights
    li133.State = 1
    li133a.State = 1
End Sub

Sub TurnOffPlayfieldLights()
    For each x in aLights
        x.State = 0
    Next
End Sub

'******************************************
'           Table Hit events
'
' Each switch/target should do:
' - play a physical sound
' - check for tilt
' - animation?
' - add a score
' - change lights?
' - check for completion of something
'*******************************************

'******************************
'        Ball advances
'******************************

Sub AddAdvance(points)
    If li81.State Then Exit Sub
    Advances = Advances + points
    UpdateAdvances
End Sub

Sub UpdateAdvances
    Select Case Advances
        Case 0:li133.State = 0:li133a.State = 0:li50.State = 0:li50a.State = 0:li51.State = 0:li51a.State = 0:li52.State = 0:li52a.State = 0
        Case 1:li133.State = 1:li133a.State = 1:li50.State = 0:li50a.State = 0:li51.State = 0:li51a.State = 0:li52.State = 0:li52a.State = 0
        Case 2:li133.State = 0:li133a.State = 0:li50.State = 1:li50a.State = 1:li51.State = 0:li51a.State = 0:li52.State = 0:li52a.State = 0
        Case 3:li133.State = 0:li133a.State = 0:li50.State = 0:li50a.State = 0:li51.State = 1:li51a.State = 1:li52.State = 0:li52a.State = 0
        Case 4:li133.State = 0:li133a.State = 0:li50.State = 0:li50a.State = 0:li51.State = 0:li51a.State = 0:li52.State = 1:li52a.State = 1
        Case Else:li133.State = 0:li133a.State = 0:li50.State = 0:li50a.State = 0:li51.State = 0:li51a.State = 0:li52.State = 0:li52a.State = 0
            li81.State = 1
            li100.State = 1
            li99.State = 1
            Advances = 0
    End Select
End Sub

Sub sw12_Hit
    PlaySoundAt "fx_sensor", sw12
    If Tilted Then Exit Sub
    If li133.State Then
    DOF 224, DOFPulse
        AddAdvance 1
        playsound "so_staron"
    Else
        PlaySound "so_staroff"
    End If
End Sub

Sub sw12a_Hit
    PlaySoundAt "fx_sensor", sw12a
    If Tilted Then Exit Sub
    If li133a.State Then
    DOF 220, DOFPulse
        AddAdvance 1
        playsound "so_staron"
    Else
        PlaySound "so_staroff"
    End If
End Sub

Sub sw22_Hit
    PlaySoundAt "fx_sensor", sw22
    If Tilted Then Exit Sub
    If li50.State Then
    DOF 225, DOFPulse
        AddAdvance 1
    End If
End Sub

Sub sw22a_Hit
    PlaySoundAt "fx_sensor", sw22a
    If Tilted Then Exit Sub
    If li50a.State Then
    DOF 221, DOFPulse
        AddAdvance 1
        playsound "so_staron"
    Else
        PlaySound "so_staroff"
    End If
End Sub

Sub sw32_Hit
    PlaySoundAt "fx_sensor", sw32
    If Tilted Then Exit Sub
    If li51.State Then
    DOF 226, DOFPulse
        AddAdvance 1
        playsound "so_staron"
    Else
        PlaySound "so_staroff"
    End If
End Sub

Sub sw32a_Hit
    PlaySoundAt "fx_sensor", sw32a
    If Tilted Then Exit Sub
    If li51a.State Then
    DOF 222, DOFPulse
        AddAdvance 1
        playsound "so_staron"
    Else
        PlaySound "so_staroff"
    End If
End Sub

Sub sw42_Hit
    PlaySoundAt "fx_sensor", sw42
    If Tilted Then Exit Sub
    If li52.State Then
    DOF 227, DOFPulse
        AddAdvance 1
        playsound "so_staron"
    Else
        PlaySound "so_staroff"
    End If
End Sub

Sub sw42a_Hit
    PlaySoundAt "fx_sensor", sw42a
    If Tilted Then Exit Sub
    If li52a.State Then
    DOF 223, DOFPulse
        AddAdvance 1
        playsound "so_staron"
    Else
        PlaySound "so_staroff"
    End If
End Sub

'******************************
'           Rubbers
'******************************

Dim Rub1, Rub2, Rub3, Rub4, Rub5, Rub6, Rub7

Sub sw51_Hit
    If Tilted then Exit Sub
    AddScore 50
    PlaySound "so_middleband"
    Rub1 = 1:sw51_Timer
End Sub

Sub sw51_Timer
    Select Case Rub1
        Case 1:rubber16.Visible = 0:rubber26.Visible = 1:sw51.TimerEnabled = 1
        Case 2:rubber26.Visible = 0:rubber27.Visible = 1
        Case 3:rubber27.Visible = 0:rubber16.Visible = 1:sw51.TimerEnabled = 0
    End Select
    Rub1 = Rub1 + 1
End Sub

Sub sw31_Hit
    If Tilted then Exit Sub
    AddScore 50
    PlaySound "so_lowerband"
    SwapSpecialLights
    Rub2 = 1:sw31_Timer
End Sub

Sub sw31_Timer
    Select Case Rub2
        Case 1:rubber35.Visible = 0:rubber36.Visible = 1:sw31.TimerEnabled = 1
        Case 2:rubber36.Visible = 0:rubber37.Visible = 1
        Case 3:rubber37.Visible = 0:rubber35.Visible = 1:sw31.TimerEnabled = 0
    End Select
    Rub2 = Rub2 + 1
End Sub

Sub sw23_Hit
    If Tilted then Exit Sub
    AddScore 50
    PlaySound "so_slings"
    SwapSpecialLights
    Rub3 = 1:sw23_Timer
End Sub

Sub sw23_Timer
    Select Case Rub3
        Case 1:rubber29.Visible = 0:rubber6.Visible = 1:sw23.TimerEnabled = 1
        Case 2:rubber6.Visible = 0:rubber7.Visible = 1
        Case 3:rubber7.Visible = 0:rubber29.Visible = 1:sw23.TimerEnabled = 0
    End Select
    Rub3 = Rub3 + 1
End Sub

Sub sw33_Hit
    If Tilted then Exit Sub
    AddScore 50
    PlaySound "so_middleband"
    Rub4 = 1:sw33_Timer
End Sub

Sub sw33_Timer
    Select Case Rub4
        Case 1:rubber18.Visible = 0:rubber28.Visible = 1:sw33.TimerEnabled = 1
        Case 2:rubber28.Visible = 0:rubber30.Visible = 1
        Case 3:rubber30.Visible = 0:rubber18.Visible = 1:sw33.TimerEnabled = 0
    End Select
    Rub4 = Rub4 + 1
End Sub

Sub sw5_Hit
    If Tilted then Exit Sub
    AddScore 3
    PlaySound "so_topband"
    Rub5 = 1:sw5_Timer
End Sub

Sub sw5_Timer
    Select Case Rub5
        Case 1:rubber22.Visible = 0:rubber31.Visible = 1:sw5.TimerEnabled = 1
        Case 2:rubber31.Visible = 0:rubber32.Visible = 1
        Case 3:rubber32.Visible = 0:rubber22.Visible = 1:sw5.TimerEnabled = 0
    End Select
    Rub5 = Rub5 + 1
End Sub

Sub sw15_Hit
    If Tilted then Exit Sub
    AddScore 3
    PlaySound "so_topband"
    Rub6 = 1:sw15_Timer
End Sub

Sub sw15_Timer
    Select Case Rub6
        Case 1:rubber21.Visible = 0:rubber33.Visible = 1:sw15.TimerEnabled = 1
        Case 2:rubber33.Visible = 0:rubber34.Visible = 1
        Case 3:rubber34.Visible = 0:rubber21.Visible = 1:sw15.TimerEnabled = 0
    End Select
    Rub6 = Rub6 + 1
End Sub

Sub sw34_Hit
    If Tilted then Exit Sub
    AddScore 100
    PlaySound "so_bandleftbumper"
    AddAdvance 1
    Rub7 = 1:sw34_Timer
End Sub

Sub sw34_Timer
    Select Case Rub7
        Case 1:rubber20.Visible = 0:rubber24.Visible = 1:sw34.TimerEnabled = 1
        Case 2:rubber24.Visible = 0:rubber25.Visible = 1
        Case 3:rubber25.Visible = 0:rubber20.Visible = 1:sw34.TimerEnabled = 0
    End Select
    Rub7 = Rub7 + 1
End Sub

'******************************
'          Bumpers
'******************************

Sub Bumper1_Hit
    If Tilted Then Exit Sub
    PlaySoundAt SoundFXDOF ("fx_Bumper",105,DOFPulse,DOFContactors), bumper1
    DOF 205, DOFPulse
    If Epileptikdof = True Then DOF 201, DOFPulse End If
    PlaySound "so_bumper"
    AddScore 100
End Sub

Sub Bumper2_Hit
    If Tilted Then Exit Sub
    PlaySoundAt SoundFXDOF ("fx_Bumper",107,DOFPulse,DOFContactors), bumper2
    DOF 207, DOFPulse
    If Epileptikdof = True Then DOF 201, DOFPulse End If
    PlaySound "so_bumper"
    AddScore 100
End Sub

'******************************
'         Outlanes
'******************************

Sub sw41_Hit
    PlaySoundAt "fx_sensor", sw41
    If Tilted Then Exit Sub
    DOF 211, DOFPulse
    AddScore 500
    If li119.State Then
        AwardSpecial
        li119.State = 0
        PlaySound "so_special"
    Else
        PlaySound "so_outlane"
    End If
End Sub

Sub sw13_Hit
    PlaySoundAt "fx_sensor", sw13
    If Tilted Then Exit Sub
    DOF 212, DOFPulse
    AddScore 500
    If li129.State Then
        AwardSpecial
        li129.State = 0
        PlaySound "so_special"
    Else
        PlaySound "so_outlane"
    End If
End Sub

'******************************
'        Top lanes
'******************************

Sub sw25_Hit
    PlaySoundAt "fx_sensor", sw25
    If Tilted Then Exit Sub
    DOF 216, DOFPulse
    AddScore 100
    If li153.State Then
        li153.State = 0
        li102.State = 0
        AddAdvance 1
        CheckGoal
    ElseIf li70.State Then
        li70.State = 0
        li120.State = 0
        AddAdvance 1
        CheckGoal
    End If
End Sub

Sub sw35_Hit
    PlaySoundAt "fx_sensor", sw35
    If Tilted Then Exit Sub
    DOF 217, DOFPulse
    If li81.State Then 'goal is lit
        AddScore 2000
        AddBonus 1
        If GoalReady Then
            AddBonusMultiplier 1
            Goalready = False
            vpmtimer.AddTimer 1000, "ResetGoalLights '"
        End If
        TurnOffGoal
        PlaySound "so_goal"
    Else
        AddScore 100
        PlaySound "so_toplanes"
    End If
    If li71.State then
        AwardExtraBall
        li71.State = 0
    End If
End Sub

Sub sw45_Hit
    PlaySoundAt "fx_sensor", sw45
    If Tilted Then Exit Sub
    DOF 218, DOFPulse
    AddScore 100
    If li72.State Then
        li72.State = 0
        li122.State = 0
        AddAdvance 1
        CheckGoal
    Else
        li79.State = 0
        li101.State = 0
        AddAdvance 1
        CheckGoal
    End If
End Sub

'******************************
'        Side lanes
'******************************

Sub sw54_Hit 'left
    PlaySoundAt "fx_sensor", sw45
    If Tilted Then Exit Sub
    DOF 215, DOFPulse
    AddAdvance 5
    AddScore 500
    sw54hits = sw54hits + 1
    AddScore 1000 * li89.State + 1000 * li90.State + 1000 * li91.State + 1000 * li92.State
    Select Case sw54hits
        Case 1:li89.State = 1:PlaySound "so_leftsidelanes"
        Case 2:li90.State = 1:PlaySound "so_leftsidelanes"
        Case 3:li91.State = 1:PlaySound "so_leftsidelanes"
        Case 4:li92.State = 1:PlaySound "so_leftsidelanes"
        Case 5:li89.State = 0:li90.State = 0:li91.State = 0:li92.State = 0:sw54hits = 0:PlaySound "so_resetlanes"
    End Select
End Sub

Sub sw43_Hit 'right to bumper
    PlaySoundAt "fx_sensor", sw43
    If Tilted Then Exit Sub
    DOF 213, DOFPulse
    AddAdvance 5
    Addscore 500
    PlaySound "so_bumperinlane"
End Sub

Sub sw53_Hit 'right
    Dim tmp
    PlaySoundAt "fx_sensor", sw53
    If Tilted Then Exit Sub
    DOF 214, DOFPulse
    AddAdvance 5
    AddScore 500
    sw53hits = sw53hits + 1
    AddScore 1000 * li110.State + 1000 * li111.State + 1000 * li112.State
    Select Case sw53hits
        Case 1:li110.State = 1:PlaySound "so_rightsidelanes"
        Case 2:li111.State = 1:PlaySound "so_rightsidelanes"
        Case 3:li112.State = 1:PlaySound "so_rightsidelanes"
        Case 4:li110.State = 0:li111.State = 0:li112.State = 0:sw53hits = 0::PlaySound "so_resetlanes"
    End Select
End Sub

'******************************
'    Targets & Holes G-O-A-L
'******************************

Sub TargetG_hit
    PlaySoundAtBall SoundFXDOF ("fx_target", 106, DOFPulse, DOFTargets)
    If Tilted Then Exit Sub
    DOF 231, DOFPulse
    Addscore 100
    If li102.State Then
        li102.State = 0
        li153.State = 0
        AddAdvance 1
        CheckGoal
        PlaySound "so_targeton"
    Else
        PlaySound "so_target"
    End If
End Sub

Sub TargetL_hit
    PlaySoundAtBall SoundFXDOF ("fx_target", 106, DOFPulse, DOFTargets)
    If Tilted Then Exit Sub
    DOF 232, DOFPulse
    Addscore 100
    PlaySound "so_target"
    If li101.State Then
        li101.State = 0
        li79.State = 0
        AddAdvance 1
        CheckGoal
        PlaySound "so_targeton"
    Else
        PlaySound "so_target"
    End If
End Sub

Sub kickerO_hit
    PlaySoundAt "fx_kicker_enter", kickerO
    If Not Tilted Then
        DOF 203, DOFPulse
        Addscore 100
        If li120.State Then
            li120.State = 0
            li70.State = 0
            AddAdvance 1
            CheckGoal
            PlaySound "so_targeton"
        Else
            PlaySound "so_target"
        End If
    End If
    kickerPos = 1
    vpmtimer.addtimer 1500, "DOF 230,DOFPulse:DOF 103,DOFPulse:PlaySoundAt SoundFX(""fx_kicker"",DOFContactors), kickerO: kickerO.kick 35+INT(RND*6), 26.5:kickerO_Timer '"
End Sub

Sub kickerO_Timer
    Select Case kickerPos
        Case 1:sw21p.rotx = 26:kickerO.TimerEnabled = 1
        Case 2:sw21p.rotx = 14
        Case 3:sw21p.rotx = 2
        Case 4:sw21p.rotx = -15:kickerO.TimerEnabled = 0
    End select
    kickerPos = kickerPos + 1
End Sub

Sub kickerA_hit
    PlaySoundAt "fx_kicker_enter", kickerA
    If Not Tilted Then
    DOF 204, DOFPulse
        Addscore 100
        If li122.State Then
            li122.State = 0
            li72.State = 0
            AddAdvance 1
            CheckGoal
            PlaySound "so_targeton"
        Else
            PlaySound "so_target"
        End If
    End If
    kickerPos = 1
    vpmtimer.addtimer 1500, "DOF 230,DOFPulse:DOF 104,DOFPulse:PlaySoundAt SoundFX(""fx_kicker"",DOFContactors), kickerA: kickerA.kick 322 +INT(RND*6), 26.5:kickerA_Timer '"
End Sub

Sub kickerA_Timer
    Select Case kickerPos
        Case 1:sw11p.rotx = 26:kickerA.TimerEnabled = 1
        Case 2:sw11p.rotx = 14
        Case 3:sw11p.rotx = 2
        Case 4:sw11p.rotx = -15:kickerA.TimerEnabled = 0
    End select
    kickerPos = kickerPos + 1
End Sub

Sub CheckGoal 'kickers
    If li153.State + li70.State + li72.State + li79.State = 0 Then
        AddScore 2000
        Goalready = True
        li81.State = 1
        li100.State = 1
        li99.State = 1
    End If
End Sub

'******************************
'        Goal Holes
'******************************

Sub sw2_Hit
    PlaySoundAt "fx_kicker_enter", sw2
    If Not Tilted Then
    DOF 208, DOFPulse
        If li100.State Then
            AddScore 2000
            AddBonus 1
            TurnOffGoal
            PlaySound "so_goal"
            If GoalReady Then
                AddBonusMultiplier 1
                Goalready = False
                vpmtimer.AddTimer 1000, "ResetGoalLights '"
            End If
        Else
            AddScore 500
            PlaySound "so_kickerhit"
        End If
    End If
    kickerPos = 1
    vpmtimer.addtimer 1500, "DOF 230,DOFPulse:DOF 110,DOFPulse:PlaySoundAt SoundFX(""fx_kicker"",DOFContactors), sw2: sw2.kick 152+INT(RND*6), 18:sw2_Timer  '"
End Sub

Sub sw2_Timer
    Select Case kickerPos
        Case 1:sw2p.rotx = -30:sw2.TimerEnabled = 1
        Case 2:sw2p.rotx = -20
        Case 3:sw2p.rotx = -10
        Case 4:sw2p.rotx = 0:sw2.TimerEnabled = 0
    End select
    kickerPos = kickerPos + 1
End Sub

Sub sw3_Hit
    PlaySoundAt "fx_kicker_enter", sw3
    If Not Tilted Then
    DOF 210, DOFPulse
        If li100.State Then
            AddScore 2000
            AddBonus 1
            TurnOffGoal
            PlaySound "so_goal"
            If GoalReady Then
                AddBonusMultiplier 1
                Goalready = False
                vpmtimer.AddTimer 1000, "ResetGoalLights '"
            End If
        Else
            AddScore 500
            PlaySound "so_kickerhit"
        End If
    End If
    kickerPos = 1
    vpmtimer.addtimer 1500, "DOF 230,DOFPulse:DOF 110,DOFPulse:PlaySoundAt SoundFX(""fx_kicker"",DOFContactors), sw3: sw3.kick 202+INT(RND*6), 18:sw3_Timer '"
End Sub

Sub sw3_Timer
    Select Case kickerPos
        Case 1:sw3p.rotx = -30:sw3.TimerEnabled = 1
        Case 2:sw3p.rotx = -20
        Case 3:sw3p.rotx = -10
        Case 4:sw3p.rotx = 0:sw3.TimerEnabled = 0
    End select
    kickerPos = kickerPos + 1
End Sub

'*****************************
'          Spinners
'*****************************

Sub sw52_Spin
    PlaySoundAt "fx_spinner", sw52
    PlaySound "so_spinner"
    If Tilted Then Exit Sub
    DOF 219, DOFPulse
    If li80.State Then
        Addscore 100
    Else
        AddScore 3
    End If
End Sub

'******************************
'     Div. routines
'******************************

Sub TurnOffGoal
    li100.State = 0
    li99.State = 0
    li81.State = 0
    AddAdvance 1
End Sub

Sub ResetGoalLights
    For each x in aGoalLights
        x.State = 1
    Next
End Sub

Sub SwapSpecialLights
    Dim tmp
    tmp = li119.State
    li119.State = li129.State
    li129.State = tmp
End Sub

'******************************
'     DOF lights ball entrance
'******************************
Sub TriggerLaunch_Hit
    If Tilted Then Exit Sub
    DOF 260, DOFPulse
End sub

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

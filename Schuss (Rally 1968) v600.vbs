' ***********************************************************************
'               VISUAL PINBALL X EM Script by JPSalas
'                  Basic EM script up to 4 players
'                 uses core.vbs for extra functions
'    Schuss / IPD No. 3541 / July, 1968 / 1 Player
'    by Rally a.k.a. Rally Play Company, of Nice, France (1961-1969)
' ***********************************************************************

' ** Wiesshund **
' adjusted some lamps for 10.8 lighting

Option Explicit
Randomize

' core.vbs constants
Const BallSize = 50  ' 50 is the normal size
Const BallMass = 1   ' 1 is the normal ball mass.

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
Const cGameName = "schuss" ' B2S name
Const MaxPlayers = 1       ' 1 to 4 can play
Const MaxMultiplier = 3    ' limit bonus multiplier
Const FreePlay = False     ' Free play or coins
Const Special1 = 6900      ' extra ball or credit
Const Special2 = 8000
Const Special3 = 9200

'Mechanical Sounds volume
'values are from 1 to 0, being 0.5 about 50%
Const Vol1 = 1 'for the ball hits sounds played with PlaySoundAtBall
Const Vol2 = 1 'for the fixed mechanical sounds played with PlaySoundAt

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
Dim Add1
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
    Add1 = 0
    Add10 = 0
    Add100 = 0

    ' setup table in game over mode
    EndOfGame

    'turn on GI lights
    'vpmtimer.addtimer 500, "GiOn '"

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
    If Keycode = AddCreditKey OR Keycode = AddCreditKey2 Then
        If(Tilted = False) Then
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
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound "fx_nudge", 0, 1, -0.1, 0.25:CheckTilt
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound "fx_nudge", 0, 1, 0.1, 0.25:CheckTilt
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound "fx_nudge", 0, 1, 1, 0.25:CheckTilt
    If keycode = MechanicalTilt Then CheckTilt

    ' keys during game

    If bGameInPlay AND NOT Tilted Then
        If keycode = LeftFlipperKey Then SolLFlipper 1
        If keycode = RightFlipperKey Then SolRFlipper 1

        If keycode = StartGameKey Then
            If((PlayersPlayingGame < MaxPlayers) AND(bOnTheFirstBall = True) ) Then

                If(bFreePlay = True) Then
                    PlayersPlayingGame = PlayersPlayingGame + 1
                    UpdatePlayersPlaying
                'PlaySound "so_fanfare1"
                Else
                    If(Credits > 0) then
                        PlayersPlayingGame = PlayersPlayingGame + 1
                        Credits = Credits - 1
                        UpdateCredits
                        UpdateBallInPlay
                        UpdatePlayersPlaying
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
                    If(Credits > 0) Then
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
        LeftFlipper001.RotateToEnd
    LeftFlipperOn = 1
    Else
        PlaySoundAt SoundFXDOF("fx_flipperdown", 101, DOFOff, DOFFlippers), LeftFlipper
        LeftFlipper.RotateToStart
        LeftFlipper001.RotateToStart
    LeftFlipperOn = 0
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFXDOF("fx_flipperup", 102, DOFOn, DOFFlippers), RightFlipper
        RightFlipper.RotateToEnd
        RightFlipper001.RotateToEnd
    RightFlipperOn = 1
    Else
        PlaySoundAt SoundFXDOF("fx_flipperdown", 102, DOFOff, DOFFlippers), RightFlipper
        RightFlipper.RotateToStart
        RightFlipper001.RotateToStart
    RightFlipperOn = 0
    End If
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

' Upper flippers
'Start Of Stroke Flipper Stroke Routine : Start of Stroke for Tap pass and Tap shoot
    If LeftFlipper001.CurrentAngle >= LeftFlipper001.StartAngle - SOSAngle Then LeftFlipper001.Strength = FlipperPower * SOSTorque else LeftFlipper001.Strength = FlipperPower : End If

'End Of Stroke Routine : Livecatch and Emply/Full-Charged EOS
  If LeftFlipperOn = 1 Then
    If LeftFlipper001.CurrentAngle = LeftFlipper001.EndAngle then
      LeftFlipper001.EOSTorque = FullStrokeEOS_Torque
      LLiveCatchTimer = LLiveCatchTimer + 1
      If LLiveCatchTimer < LiveCatchSensivity Then
        LeftFlipper001.Elasticity = 0
      Else
        LeftFlipper001.Elasticity = FlipperElasticity
        LLiveCatchTimer = LiveCatchSensivity
      End If
    End If
  Else
    LeftFlipper001.Elasticity = FlipperElasticity
    LeftFlipper001.EOSTorque = LiveStrokeEOS_Torque
    LLiveCatchTimer = 0
  End If


'Start Of Stroke Flipper Stroke Routine : Start of Stroke for Tap pass and Tap shoot
    If RightFlipper001.CurrentAngle <= RightFlipper001.StartAngle + SOSAngle Then RightFlipper001.Strength = FlipperPower * SOSTorque else RightFlipper001.Strength = FlipperPower : End If

'End Of Stroke Routine : Livecatch and Emply/Full-Charged EOS
  If RightFlipperOn = 1 Then
    If RightFlipper001.CurrentAngle = RightFlipper001.EndAngle Then
      RightFlipper001.EOSTorque = FullStrokeEOS_Torque
      RLiveCatchTimer = RLiveCatchTimer + 1
      If RLiveCatchTimer < LiveCatchSensivity Then
        RightFlipper001.Elasticity = 0
      Else
        RightFlipper001.Elasticity = FlipperElasticity
        RLiveCatchTimer = LiveCatchSensivity
      End If
    End If
  Else
    RightFlipper001.Elasticity = FlipperElasticity
    RightFlipper001.EOSTorque = LiveStrokeEOS_Torque
    RLiveCatchTimer = 0
  End If
End Sub

' flippers hit Sound

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, Vol1 * parm / 100, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, Vol1 * parm / 100, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub LeftFlipper001_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, Vol1 * parm / 100, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper001_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, Vol1 * parm / 100, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

' Flipper animations
Sub LeftFlipper_Animate:    LeftFlipperTop.RotZ = LeftFlipper.CurrentAngle:End Sub
Sub RightFlipper_Animate:    RightFlipperTop.RotZ = RightFlipper.CurrentAngle:End Sub
Sub LeftFlipper001_Animate:    LeftFlipperTop001.RotZ = LeftFlipper001.CurrentAngle:End Sub
Sub RightFlipper001_Animate:    RightFlipperTop001.RotZ = RightFlipper001.CurrentAngle:End Sub
Sub Gate001_Animate:    Diverter.Rotz = - Gate001.Currentangle:End Sub

'***********
' GI lights
'***********

Sub GiOn 'enciende las luces GI
    Dim bulb
    For each bulb in aGiLights
        bulb.State = 1
    Next
    PlaySound "fx_Gion"
End Sub

Sub GiOff 'apaga las luces GI
    Dim bulb
    For each bulb in aGiLights
        bulb.State = 0
    Next
    PlaySound "fx_Gioff"
End Sub

'**************
'    TILT
'**************

Sub CheckTilt
    Tilt = Tilt + TiltSensitivity
    TiltDecreaseTimer.Enabled = True
    If Tilt > 15 Then
        Tilted = True
        Tiltl.State = 1
        If B2SOn then
            Controller.B2SSetTilt 1
        end if
        DisableTable True
        ' BallsRemaining(CurrentPlayer) = 0 'player looses the game
        TiltRecoveryTimer.Enabled = True 'wait for all the balls to drain
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
        Bumper1.Threshold = 100
        Bumper2.Threshold = 100
        Bumper3.Threshold = 100
        LeftSlingshot.Disabled = 1
        RightSlingshot.Disabled = 1
        DropPost
        PlungerPost.IsDropped = 0
        PlungerPost2.IsDropped = 0
    Else
        GiOn
        Bumper1.Threshold = 1
        Bumper2.Threshold = 1
        Bumper3.Threshold = 1
        LeftSlingshot.Disabled = 0
        RightSlingshot.Disabled = 0
        PlungerPost.IsDropped = 1
        PlungerPost2.IsDropped = 1
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
'             Supporting Ball & Sound Functions v3.0
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
    If tmp > 0 Then
        AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Sub PlaySoundAt(soundname, tableobj) 'play sound at X and Y position of an object, mostly bumpers, flippers and other fast objects
    PlaySound soundname, 0, Vol2, Pan(tableobj), 0.1, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname) ' play a sound at the ball position, like rubbers, targets, metals, plastics
    PlaySound soundname, 0, Vol1 * Vol(ActiveBall), pan(ActiveBall), 0.4, 0, 0, 0, AudioFade(ActiveBall)
End Sub

'***********************************************
'   JP's VP10 Rolling Sounds + Ballshadow v3.0
'***********************************************

Const tnob = 19   'total number of balls, 20 balls, from 0 to 19
Const lob = 0     'number of locked balls
Const maxvel = 24 'max ball velocity
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

    ' stop the sound of deleted balls and hide the shadow
    For b = UBound(BOT) + 1 to tnob
        rolling(b) = False
        StopSound("fx_ballrolling" & b)
    Next

    ' exit the sub if no balls on the table
    If UBound(BOT) = lob - 1 Then Exit Sub 'there no extra balls on this table

    ' play the rolling sound for each ball and draw the shadow
    For b = lob to UBound(BOT)
        If BallVel(BOT(b) ) > 1 Then
            If BOT(b).z < 30 Then
                ballpitch = Pitch(BOT(b) )
                ballvol = Vol(BOT(b) )
            Else
                ballpitch = Pitch(BOT(b) ) + 25000 'increase the pitch on a ramp
                ballvol = Vol(BOT(b) ) * 5
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
        'If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
        '    PlaySound "fx_balldrop", 0, ABS(BOT(b).velz) / 17, Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        'End If

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
    UpdatePlayersPlaying
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
    vpmtimer.addtimer 100, "FirstBall '"
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
    'BallRelease.CreateSizedBallWithMass BallSize / 2, BallMass
    BallsOnPlayfield = BallsOnPlayfield + 1
    UpdateBallInPlay
    UpdateCurrentPlayer
End Sub

' player lost the ball

Sub EndOfBall()
    'debug.print "EndOfBall"
    ' Lost the first ball, now it cannot accept more players
    bOnTheFirstBall = False

    StartBonusCount 'the bonus count does the endofball2
'vpmtimer.addtimer 400, "EndOfBall2 '"
End Sub

' After bonus count go to the next step

Sub EndOfBall2()
    'debug.print "EndOfBall2"

    Tilted = False
    Tilt = 0
    Tiltl.State = 0
    If B2SOn then
        Controller.B2SSetTilt 0
    end if
    DisableTable False

    ' win extra ball?
    If(ExtraBallsAwards(CurrentPlayer) > 0) Then
        'debug.print "Extra Ball"

        ' if so then give it
        ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) - 1

        ' turn off light if no more extra balls
        If(ExtraBallsAwards(CurrentPlayer) = 0) Then
            ShootAgain.State = 0
            If B2SOn then
                Controller.B2SSetShootAgain 0
            End If
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
    If(PlayersPlayingGame > 1) Then
        NextPlayer = CurrentPlayer + 1
        ' if it is the last player then go to the first one
        If(NextPlayer > PlayersPlayingGame) Then
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
    PlaySound "resetscore"
    DisableTable True
    'turn off bumper lights
    li042.State = 0
    li043.State = 0
    li044.State = 0
    bGameInPLay = False
    bJustStarted = False
    GameOverL.State = 1
    If B2SOn then
        Controller.B2SSetGameOver 1
        Controller.B2SSetBallInPlay 0
        Controller.B2SSetPlayerUp 0
        Controller.B2SSetCanPlay 0
    end if
    ' start the attract mode
    StartAttractMode
End Sub

' Function to calculate the balls left
Function Balls
    Dim tmp
    tmp = BallsPerGame - BallsRemaining(CurrentPlayer) + 1
    If tmp > BallsPerGame Then
        Balls = BallsPerGame
    Else
        Balls = tmp
    End If
End Function

'******************
'     Match
'******************

Sub Verification_Match()
    'PlaySound "fx_match"
    Match = INT(RND(1) * 10)
    Display_Match
    If(Score(CurrentPlayer) MOD 10) = Match Then
        PlaySound "fx_knocker"
        AddCredits 1
    End If
End Sub

Sub Clear_Match()
    For each x in aMatchL:x.State = 0:Next
    If B2SOn then
        Controller.B2SSetMatch 0
    end if
End Sub

Sub Display_Match()
    aMatchL(Match).State = 1
    If B2SOn then
        If Match = 0 then
            Controller.B2SSetMatch 10
        Else
            Controller.B2SSetMatch Match
        End If
    end if
End Sub

' *********************************************************************
'                      Drain / Plunger Functions
' *********************************************************************

Sub Drain_Hit()
    'Drain.DestroyBall
    If bGameInPlay = False Then Exit Sub
    BallsOnPlayfield = BallsOnPlayfield - 1
    PlaySoundAt "fx_sensor", Drain

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
                vpmtimer.addtimer 100, "EndOfBall '"
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
    If Tilted or bGameInPlay = False Then Exit Sub
    Select Case Points
        Case 1, 10, 100
            Score(CurrentPlayer) = Score(CurrentPlayer) + points
            UpdateScore points
            ' sounds
            PlaySound "beep" &Points
        Case 20, 30, 40, 50, 60, 70, 80, 90
            Add10 = Add10 + Points \ 10
            AddScore10Timer.Enabled = TRUE
        Case 200, 300, 400, 500, 600, 700, 800, 900
            Add100 = Add100 + Points \ 100
            AddScore100Timer.Enabled = TRUE
        Case 2, 3, 4, 5, 6, 7, 8, 9
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
End Sub

'************************************
'       Score sound Timers
'************************************

Sub AddScore10Timer_Timer()
    if Add10 > 0 then
        AddScore 10
        Add10 = Add10 - 1
    Else
        Me.Enabled = FALSE
    End If
End Sub

Sub AddScore100Timer_Timer()
    if Add100 > 0 then
        AddScore 100
        Add100 = Add100 - 1
    Else
        Me.Enabled = FALSE
    End If
End Sub

Sub AddScore1Timer_Timer()
    if Add1 > 0 then
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
        If Bonus > 5 Then
            Bonus = 5
        End If
        UpdateBonusLights
    End if
End Sub

Sub AddBonusMultiplier(multi) 'not used in this table
    If(Tilted = False) Then
        BonusMultiplier = BonusMultiplier + multi
        If BonusMultiplier > MaxMultiplier Then
            BonusMultiplier = MaxMultiplier
        End If
        UpdateMultiplierLights
    End if
End Sub

Sub UpdateMultiplierLights 'not used in this table
    Select Case BonusMultiplier
    'Case 1:li2.State = 1:li3.State = 0:li4.State = 0
    'Case 2:li2.State = 0:li3.State = 1:li4.State = 0
    'Case 3:li2.State = 0:li3.State = 0:li4.State = 1
    End Select
End Sub

Sub StartBonusCount
    BonusMultiplier = 1 + l001.State + l002.State
    Select Case BonusMultiplier
        Case 1:BonusCountTimer.Interval = 250
        Case 2:BonusCountTimer.Interval = 400
        Case 3:BonusCountTimer.Interval = 650
    End Select
    BonusCountTimer.Enabled = 1
End Sub

Sub BonusCountTimer_Timer
    'debug.print "BonusCount_Timer"
    If Bonus > 0 Then
        Bonus = Bonus -1
        AddScore 100 * BonusMultiplier
        UpdateBonusLights
    Else
        BonusCountTimer.Enabled = 0
        vpmtimer.addtimer 100, "EndOfBall2 '"
    End If
End Sub

Sub UpdateBonusLights
    Select Case Bonus
        Case 0:bl1.State = 0:bl2.State = 0:bl3.State = 0:bl4.State = 0:bl5.State = 0
        Case 1:bl1.State = 1:bl2.State = 0:bl3.State = 0:bl4.State = 0:bl5.State = 0
        Case 2:bl1.State = 1:bl2.State = 1:bl3.State = 0:bl4.State = 0:bl5.State = 0
        Case 3:bl1.State = 1:bl2.State = 1:bl3.State = 1:bl4.State = 0:bl5.State = 0
        Case 4:bl1.State = 1:bl2.State = 1:bl3.State = 1:bl4.State = 1:bl5.State = 0
        Case 5:bl1.State = 1:bl2.State = 1:bl3.State = 1:bl4.State = 1:bl5.State = 1
    End Select
End Sub

'**********************************
'        Score EM reels
'**********************************

Sub UpdateScore(playerpoints)
    Select Case CurrentPlayer
        Case 1:ScoreReel1.Addvalue playerpoints
    'Case 2:ScoreReel2.Addvalue playerpoints
    'Case 3:ScoreReel3.Addvalue playerpoints
    'Case 4:ScoreReel4.Addvalue playerpoints

    End Select
    If B2SOn then
        Controller.B2SSetScorePlayer CurrentPlayer, Score(CurrentPlayer)
    end if
End Sub

Sub ResetScores
    PlaySound "resetscore"
    ScoreReel1.SetValue 0
    'ScoreReel1.ResetToZero
    'ScoreReel2.ResetToZero
    'ScoreReel3.ResetToZero
    'ScoreReel4.ResetToZero
    If B2SOn then
        Controller.B2SSetScorePlayer1 0
        Controller.B2SSetScoreRolloverPlayer1 0
    'Controller.B2SSetScorePlayer2 0
    'Controller.B2SSetScoreRolloverPlayer2 0
    'Controller.B2SSetScorePlayer3 0
    'Controller.B2SSetScoreRolloverPlayer3 0
    'Controller.B2SSetScorePlayer4 0
    'Controller.B2SSetScoreRolloverPlayer4 0
    end if
End Sub

Sub AddCredits(value) 'limit to 15 credits
    If Credits < 9 Then
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
    end if
End Sub

Sub UpdatePlayersPlaying
    Select Case PlayersPlayingGame
    'Case 0:cp1.State = 0:cp2.State = 0:cp3.State = 0:cp4.State = 0
    'Case 1:cp1.State = 1:cp2.State = 0:cp3.State = 0:cp4.State = 0
    'Case 2:cp1.State = 1:cp2.State = 1:cp3.State = 0:cp4.State = 0
    'Case 3:cp1.State = 1:cp2.State = 1:cp3.State = 1:cp4.State = 0
    'Case 4:cp1.State = 1:cp2.State = 1:cp3.State = 1:cp4.State = 1
    End Select
    If B2SOn then
    'Controller.B2SSetPlayersPlaying PlayersPlayingGame
    end if
End Sub

Sub UpdateCurrentPlayer
    Select Case CurrentPlayer
    'Case 0:pl1.State = 0:pl2.State = 0:pl3.State = 0:pl4.State = 0
    'Case 1:pl1.State = 1:pl2.State = 0:pl3.State = 0:pl4.State = 0
    'Case 2:pl1.State = 0:pl2.State = 1:pl3.State = 0:pl4.State = 0
    'Case 3:pl1.State = 0:pl2.State = 0:pl3.State = 1:pl4.State = 0
    'Case 4:pl1.State = 0:pl2.State = 0:pl3.State = 0:pl4.State = 1
    End Select
    If B2SOn then
    'Controller.B2SSetCurrentPlayer CurrentPlayer
    end if
End Sub

'*************************
'        Specials
'*************************

Sub AwardExtraBall()
    If NOT bExtraBallWonThisBall Then
        PlaySound "fx_knocker"
        ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) + 1
        bExtraBallWonThisBall = True
        ShootAgain.State = 1
        If B2SOn then
            Controller.B2SSetShootAgain 1
        End If
    Else
        Addscore 50
    END If
End Sub

Sub AwardSpecial()
    PlaySound "fx_knocker"
    AddCredits 1
End Sub

Sub AwardAddaBall()
    If BallsRemaining(CurrentPlayer) < 11 Then
        PlaySound "fx_knocker"
        BallsRemaining(CurrentPlayer) = BallsRemaining(CurrentPlayer) + 1
        UpdateBallInPlay
    End If
End Sub

' ********************************
'        Attract Mode
' ********************************
' use the"Blink Pattern" of each light

Sub StartAttractMode()
    Dim x
    bAttractMode = True
    TurnOffPlayfieldLights
    GiOff
    GameOverL.State = 1
    StartLightSeq
    PlaySound "fx_motor", -1, 0.001
    'update current player and balls
    bip5.State = 0
End Sub

Sub StopAttractMode()
    Dim x
    bAttractMode = False
    TurnOffPlayfieldLights
    ResetScores
    GameOverL.State = 0
    LightSeqAttract.StopPlay
    StopSound "fx_motor"
End Sub

Sub StartLightSeq()
    'lights sequences
    LightSeqAttract.UpdateInterval = 20
    'LightSeqAttract.Play SeqAllOff
    LightSeqAttract.Play SeqDiagUpRightOn, 25
    LightSeqAttract.Play SeqDiagDownLeftOn, 25
    LightSeqAttract.Play SeqStripe1VertOn, 25
    LightSeqAttract.Play SeqStripe2VertOn, 25
    LightSeqAttract.Play SeqClockRightOn, 180, 2
    LightSeqAttract.Play SeqFanLeftUpOn, 50
    LightSeqAttract.Play SeqFanLeftDownOn, 50
    LightSeqAttract.Play SeqFanRightUpOn, 50
    LightSeqAttract.Play SeqFanRightDownOn, 50
    LightSeqAttract.Play SeqScrewRightOn, 50
End Sub

Sub LightSeqAttract_PlayDone()
    StartLightSeq()
End Sub

'*************************************
'       Highscore Routines
' Based og Black's and GNance's code
'*************************************

Dim EnteringInitials
EnteringInitials = False

Dim SelectedChar   ' character under the "cursor" when entering initials
Dim HSTimerCount:HSTimerCount = 0
Dim InitialString  ' the string holding the player's initials as they're entered
Dim AlphaString:AlphaString = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_<"
Dim AlphaStringPos ' pointer to AlphaString, move Forward and backward with flipper keys
Dim HSNewHigh      ' The new score to be recorded
Dim HSScore(5)     ' High Scores read in from config file
Dim HSName(5)      ' High Score Initials read in from config file

' default high scores
HSScore(1) = 2500
HSScore(2) = 2400
HSScore(3) = 2300
HSScore(4) = 2200
HSScore(5) = 2100

HSName(1) = "AAA"
HSName(2) = "ZZZ"
HSName(3) = "XXX"
HSName(4) = "ABC"
HSName(5) = "BBB"

' Load Highscores

Sub Loadhs
    Dim FileObj
    Dim ScoreFile, TextStr
    Dim check0, x

    Set FileObj = CreateObject("Scripting.FileSystemObject")
    If Not FileObj.FolderExists(UserDirectory) then
        Credits = 0
        Exit Sub
    End If
    If Not FileObj.FileExists(UserDirectory & cGameName& ".txt") then
        Credits = 0
        Exit Sub
    End If
    Set ScoreFile = FileObj.GetFile(UserDirectory & cGameName& ".txt")
    Set TextStr = ScoreFile.OpenAsTextStream(1, 0)
    If(TextStr.AtEndOfStream = True) then
        Exit Sub
    End If
    check0 = Int(textStr.readLine)
    If check0 = 0 then
        Credits = Int(textstr.readline)
        For x = 1 to 5
            HSScore(x) = Int(textstr.readline)
        Next
        For x = 1 to 5
            HSName(x) = textstr.readline
        Next
    End If
    TextStr.Close
    Set ScoreFile = Nothing
    Set FileObj = Nothing
    SortHighscore 'added to fix a previous error
End Sub

' Save Highscores

Sub Savehs
    Dim FileObj
    Dim ScoreFile
    Dim x
    Set FileObj = CreateObject("Scripting.FileSystemObject")
    If Not FileObj.FolderExists(UserDirectory) then
        Exit Sub
    End If
    Set ScoreFile = FileObj.CreateTextFile(UserDirectory & cGameName& ".txt", True)
    ScoreFile.WriteLine 0
    ScoreFile.WriteLine Credits
    For x = 1 to 5
        scorefile.writeline HSScore(x)
    Next
    For x = 1 to 5
        scorefile.writeline HSName(x)
    Next
    ScoreFile.Close
    Set ScoreFile = Nothing
    Set FileObj = Nothing
End Sub

' Sort Highscores

Sub SortHighscore
    Dim tmp, tmp2, i, j
    For i = 1 to 5
        For j = 1 to 4
            If HSScore(j) < HSScore(j + 1) Then
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

' check for new highscore

Sub CheckHighscore
    HighscoreTimer.Enabled = 1
    'If Score(CurrentPlayer)> HSSCore(1)Then
    '   AwardSpecial
    'End If
    If Score(CurrentPlayer) > HSSCore(5) Then
        NewHighScore Score(CurrentPlayer), CurrentPlayer
    End If
End Sub

' Highscore Timer

Sub HighScoreTimer_Timer
    If EnteringInitials then
        SetHSLine 3, InitialString & MID(AlphaString, AlphaStringPos, 1)
    ElseIf bGameInPlay then
        SetHSLine 1, "HIGH SCORE"
        SetHSLine 2, HSScore(1)
        SetHSLine 3, HSName(1)
        Me.Enabled = 0
    Else
        ' cycle through high scores
        SetHSLine 1, "HIGH SCORE" & (HSTimerCount + 1)
        SetHSLine 2, HSScore(HSTimerCount + 1)
        SetHSLine 3, HSName(HSTimerCount + 1)
        HSTimerCount = (HSTimerCount + 1) MOD 5
        If B2SOn then
            Controller.B2SSetScorePlayer 1, HSScore(HSTimerCount + 1)
        End If
    End If
End Sub

' Enter Highscore initials

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
    Dim StrLen
    Dim Index
    Dim StartHSArray
    Dim EndHSArray
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
    If NewScore > HSScore(5) then
        HighScoreTimer.interval = 500
        HighScoreTimer.Enabled = 1
        AlphaStringPos = 1      ' start with first character "A"
        EnteringInitials = true ' intercept the control keys while entering initials
        InitialString = ""      ' initials entered so far, initialize to empty
        SetHSLine 1, "PLAYER " + FormatNumber(PlayNum, 0)
        SetHSLine 2, "ENTER NAME"
        SetHSLine 3, MID(AlphaString, AlphaStringPos, 1)
        HSNewHigh = NewScore
    End If
End Sub

Sub CollectInitials(keycode)
    Dim i
    If keycode = LeftFlipperKey Then
        ' back up to previous character
        AlphaStringPos = AlphaStringPos - 1
        If AlphaStringPos < 1 then
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
        If AlphaStringPos > len(AlphaString) or(AlphaStringPos = len(AlphaString) and InitialString = "") then
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
        If len(InitialString) < 3 then
            SetHSLine 3, InitialString & SelectedChar
        End If
    End If
    If len(InitialString) = 3 then
        EnteringInitials = False
        HSScore(5) = HSNewHigh
        HSName(5) = InitialString
        SortHighScore
        HighScoreTimer.interval = 2000
    End If
End Sub


'***********************************************************************
' *********************************************************************
'  *********     G A M E  C O D E  S T A R T S  H E R E      *********
' *********************************************************************
'***********************************************************************

Sub VPObjects_Init 'init objects
    TurnOffPlayfieldLights()
    DropPost
    BallRelease.CreateSizedBallWithMass BallSize / 2, BallMass
    BallRelease.Kick 180, 0
    ' turn off bumpers & slingshots
    DisableTable True
    'init animations
    Wall031.IsDropped = 1
    Wall032.IsDropped = 1
    Wall033.IsDropped = 1
    Wall034.IsDropped = 1
    Wall035.IsDropped = 1
    Wall036.IsDropped = 1
End Sub

' Dim all the variables

Dim tget1, tget2, tget3, tget4

Sub Game_Init() 'called at the start of a new game
    'Start music?
    'Init variables?
    'Start or init timers
    'Init lights?
    TurnOffPlayfieldLights()
    Bonus = 0
    DropPost
    'turn on top bumpers
    DisableTable False
    li042.State = 1
    li043.State = 1
    li044.State = 0
End Sub

Sub StopEndOfBallMode() 'called when the last ball is drained
    DropPost
End Sub

Sub ResetNewBallVariables() 'init variables new ball/player
    BonusMultiplier = 1
    tget1 = 0
    tget2 = 0
    tget3 = 0
    tget4 = 0
End Sub

Sub ResetNewBallLights() 'init lights for new ball/player
    TurnOffPlayfieldLights()
    li044.State = 0
End Sub

Sub TurnOffPlayfieldLights()
    Dim a
    For each a in aLights
        a.State = 0
    Next
End Sub

Sub DropPost
    l003.State = 0
    Popupwall.IsDropped = 1
    Popup.Z = -32
End Sub

Sub UpPost
    l003.State = 1
    Popupwall.IsDropped = 0
    Popup.Z = 0
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
Sub LeftSlingShot_Slingshot
    If Tilted or bGameInPlay = False Then Exit Sub
    ' add points
    PlaySoundAt "fx_plastichit2", LSling1p
    AddScore 1 + 9 * l006.State
    'animation
    LSling1p.X = 209.8353 - ActiveBall.VelX/2
    LSling1p.Y = 1481.218 - ActiveBall.VelY/2
    LeftSlingshot.TimerEnabled = 1
End Sub

Sub LeftSlingshot_Timer
    LSling1p.X = 209.8353
    LSling1p.Y = 1481.218
    Me.TimerEnabled = 0
End Sub

Sub RightSlingShot_Slingshot
    If Tilted or bGameInPlay = False Then Exit Sub
    ' add points
    PlaySoundAt "fx_plastichit2", RSling1p
    AddScore 1 + 9 * l007.State
    'animation
    RSling1p.X = 669.4285 - ActiveBall.VelX/2
    RSling1p.Y = 1483.827 - ActiveBall.VelY/2
    RightSlingshot.TimerEnabled = 1
End Sub

Sub RightSlingshot_Timer
    RSling1p.X = 669.4285
    RSling1p.Y = 1483.827
    Me.TimerEnabled = 0
End Sub

'**************
'   Rubbers
'**************

Dim r0, r1, r2, r3, r4, r5, r6, r7, r8, r9, r10, r11, r12

Sub rlband013_Hit
    If Tilted or bGameInPlay = False Then Exit Sub
    AddScore 1
    Sling001.Visible = 0
    Sling016.Visible = 1
    Wall014.IsDropped = 1
    Wall033.IsDropped = 0
    r7 = 0
    Me.TimerEnabled = 1
End Sub

Sub rlband013_Timer
    Select case r7
        Case 0:Sling016.Visible = 0:Sling017.Visible = 1
        Case 1:Sling017.Visible = 0:Sling001.Visible = 1:Wall033.IsDropped = 1:Wall014.IsDropped = 0:Me.TimerEnabled = 0
    End Select
    r7 = r7 + 1
End Sub

Sub rlband016_Hit
    If Tilted or bGameInPlay = False Then Exit Sub
    AddScore 1
    Sling002.Visible = 0
    Sling018.Visible = 1
    Wall015.IsDropped = 1
    Wall034.IsDropped = 0
    r8 = 0
    Me.TimerEnabled = 1
End Sub

Sub rlband016_Timer
    Select case r8
        Case 0:Sling018.Visible = 0:Sling019.Visible = 1
        Case 1:Sling019.Visible = 0:Sling002.Visible = 1:Wall034.IsDropped = 1:Wall015.IsDropped = 0:Me.TimerEnabled = 0
    End Select
    r8 = r8 + 1
End Sub

Sub rlband022_Hit
    If Tilted or bGameInPlay = False Then Exit Sub
    AddScore 1
    Sling003.Visible = 0
    Sling009.Visible = 1
    Wall018.IsDropped = 1
    Wall031.IsDropped = 0
    r0 = 0
    Me.TimerEnabled = 1
End Sub

Sub rlband022_Timer
    Select case r0
        Case 0:Sling009.Visible = 0:Sling008.Visible = 1
        Case 1:Sling008.Visible = 0:Sling003.Visible = 1:Wall031.IsDropped = 1:Wall018.IsDropped = 0:Me.TimerEnabled = 0
    End Select
    r0 = r0 + 1
End Sub

Sub rlband023_Hit
    If Tilted or bGameInPlay = False Then Exit Sub
    AddScore 1
    Sling004.Visible = 0
    Sling013.Visible = 1
    Wall020.IsDropped = 1
    Wall032.IsDropped = 0
    r2 = 0
    Me.TimerEnabled = 1
End Sub

Sub rlband023_Timer
    Select case r2
        Case 0:Sling013.Visible = 0:Sling012.Visible = 1
        Case 1:Sling012.Visible = 0:Sling004.Visible = 1:Wall032.IsDropped = 1:Wall020.IsDropped = 0:Me.TimerEnabled = 0
    End Select
    r2 = r2 + 1
End Sub

Sub rlband007_Hit
    If Tilted or bGameInPlay = False Then Exit Sub
    AddScore 1
    Rubber002.Visible = 0
    Rubber027.Visible = 1
    Wall012.IsDropped = 1
    Wall035.IsDropped = 0
    r11 = 0
    Me.TimerEnabled = 1
End Sub

Sub rlband007_Timer
    Select case r11
        Case 0:Rubber027.Visible = 0:Rubber028.Visible = 1
        Case 1:Rubber028.Visible = 0:Rubber002.Visible = 1:Wall035.IsDropped = 1:Wall012.IsDropped = 0:Me.TimerEnabled = 0
    End Select
    r11 = r11 + 1
End Sub

Sub rlband003_Hit
    If Tilted or bGameInPlay = False Then Exit Sub
    AddScore 1
    Rubber003.Visible = 0
    Rubber029.Visible = 1
    Wall004.IsDropped = 1
    Wall036.IsDropped = 0
    r12 = 0
    Me.TimerEnabled = 1
End Sub

Sub rlband003_Timer
    Select case r12
        Case 0:Rubber029.Visible = 0:Rubber030.Visible = 1
        Case 1:Rubber030.Visible = 0:Rubber003.Visible = 1:Wall036.IsDropped = 1:Wall004.IsDropped = 0:Me.TimerEnabled = 0
    End Select
    r12 = r12 + 1
End Sub

'***************************
'   Rubbers animation only
'***************************

Sub rlband012_Hit
    Rubber016.Visible = 0
    Rubber025.Visible = 1
    r10 = 0
    Me.TimerEnabled = 1
End Sub

Sub rlband012_Timer
    Select case r10
        Case 0:Rubber025.Visible = 0:Rubber026.Visible = 1
        Case 1:Rubber026.Visible = 0:Rubber016.Visible = 1:Me.TimerEnabled = 0
    End Select
    r10 = r10 + 1
End Sub

Sub rlband011_Hit
    Rubber015.Visible = 0
    Rubber023.Visible = 1
    r9 = 0
    Me.TimerEnabled = 1
End Sub

Sub rlband011_Timer
    Select case r9
        Case 0:Rubber023.Visible = 0:Rubber024.Visible = 1
        Case 1:Rubber024.Visible = 0:Rubber015.Visible = 1:Me.TimerEnabled = 0
    End Select
    r9 = r9 + 1
End Sub

Sub rlband019_Hit
    Rubber020.Visible = 0
    Rubber021.Visible = 1
    r6 = 0
    Me.TimerEnabled = 1
End Sub

Sub rlband019_Timer
    Select case r6
        Case 0:Rubber021.Visible = 0:Rubber022.Visible = 1
        Case 1:Rubber022.Visible = 0:Rubber020.Visible = 1:Me.TimerEnabled = 0
    End Select
    r6 = r6 + 1
End Sub

Sub rlband024_Hit
    Sling005.Visible = 0
    Sling015.Visible = 1
    r5 = 0
    Me.TimerEnabled = 1
End Sub

Sub rlband024_Timer
    Select case r5
        Case 0:Sling015.Visible = 0:Sling014.Visible = 1
        Case 1:Sling014.Visible = 0:Sling005.Visible = 1:Me.TimerEnabled = 0
    End Select
    r5 = r5 + 1
End Sub

Sub rlband026_Hit
    Sling004.Visible = 0
    Sling011.Visible = 1
    r4 = 0
    Me.TimerEnabled = 1
End Sub

Sub rlband026_Timer
    Select case r4
        Case 0:Sling011.Visible = 0:Sling010.Visible = 1
        Case 1:Sling010.Visible = 0:Sling004.Visible = 1:Me.TimerEnabled = 0
    End Select
    r4 = r4 + 1
End Sub

Sub rlband020_Hit
    Sling003.Visible = 0
    Sling007.Visible = 1
    r1 = 0
    Me.TimerEnabled = 1
End Sub

Sub rlband020_Timer
    Select case r1
        Case 0:Sling007.Visible = 0:Sling006.Visible = 1
        Case 1:Sling006.Visible = 0:Sling003.Visible = 1:Me.TimerEnabled = 0
    End Select
    r1 = r1 + 1
End Sub

'*********
' Bumpers
'*********

Sub Bumper1_Hit
    If Tilted or bGameInPlay = False Then Exit Sub
    PlaySoundAt "fx_plastichit2", bumper1
    AddScore 10
    Bumper1p.X = 295 - ActiveBall.VelX/2
    Bumper1p.Y = 455 - ActiveBall.VelY/2
    Me.TimerEnabled = 1
End Sub

Sub Bumper1_Timer:Bumper1p.X = 295:Bumper1p.Y = 455:Me.TimerEnabled = 0:End Sub

Sub Bumper2_Hit
    If Tilted or bGameInPlay = False Then Exit Sub
    PlaySoundAt "fx_plastichit2", bumper1
    AddScore 10
    Bumper2p.X = 586 - ActiveBall.VelX/2
    Bumper2p.Y = 455 - ActiveBall.VelY/2
    Me.TimerEnabled = 1
End Sub

Sub Bumper2_Timer:Bumper2p.X = 586:Bumper2p.Y = 455:Me.TimerEnabled = 0:End Sub

Sub Bumper3_Hit
    If Tilted or bGameInPlay = False Then Exit Sub
    PlaySoundAt "fx_plastichit2", bumper1
    AddScore 10 + 90 * li044.State
    Bumper3p.X = 439 - ActiveBall.VelX/2
    Bumper3p.Y = 630 - ActiveBall.VelY/2
    Me.TimerEnabled = 1
End Sub

Sub Bumper3_Timer:Bumper3p.X = 439:Bumper3p.Y = 630:Me.TimerEnabled = 0:End Sub

'*****************
'    Rollovers
'*****************

Sub Trigger009_Hit 'top lane, on red bumper
    PlaySoundAt "fx_sensor", Trigger009
    If Tilted or bGameInPlay = False Then Exit Sub
    AddScore 100 + 200 * l026.State
    'red Bumper on
    li044.State = 1
End Sub

Sub Trigger010_Hit 'off red bumper
    PlaySoundAt "fx_sensor", Trigger010
    If Tilted or bGameInPlay = False Then Exit Sub
    AddScore 1
    l023.State = 1
    'red Bumper off
    li044.State = 0
End Sub

Sub Trigger010_UnHit:l023.State = 0:End Sub

Sub Trigger011_Hit 'off red bumper
    PlaySoundAt "fx_sensor", Trigger011
    If Tilted or bGameInPlay = False Then Exit Sub
    AddScore 1
    l021.State = 1
    'red Bumper off
    li044.State = 0
End Sub

Sub Trigger011_UnHit:l021.State = 0:End Sub

Sub Trigger007_Hit
    PlaySoundAt "fx_sensor", Trigger007
    If Tilted or bGameInPlay = False Then Exit Sub
    AddScore 10 + 90 * l020.State
End Sub

Sub Trigger008_Hit
    PlaySoundAt "fx_sensor", Trigger008
    If Tilted or bGameInPlay = False Then Exit Sub
    AddScore 10 + 90 * l024.State
End Sub

Sub Trigger006_Hit
    PlaySoundAt "fx_sensor", Trigger006
    If Tilted or bGameInPlay = False Then Exit Sub
    AddScore 100 + 200 * l019.State
End Sub

Sub Trigger005_Hit
    PlaySoundAt "fx_sensor", Trigger005
    If Tilted or bGameInPlay = False Then Exit Sub
    AddScore 100 + 200 * l018.State
End Sub

Sub Trigger004_Hit
    PlaySoundAt "fx_sensor", Trigger004
    If Tilted or bGameInPlay = False Then Exit Sub
    AddScore 100 + 200 * l017.State
End Sub

Sub Trigger003_Hit
    PlaySoundAt "fx_sensor", Trigger003
    If Tilted or bGameInPlay = False Then Exit Sub
    AddScore 100 + 200 * l016.State
End Sub

Sub Trigger001_Hit
    PlaySoundAt "fx_sensor", Trigger001
    If Tilted or bGameInPlay = False Then Exit Sub
    AddScore 50 + 50 * l005.State
End Sub

Sub Trigger002_Hit
    PlaySoundAt "fx_sensor", Trigger002
    If Tilted or bGameInPlay = False Then Exit Sub
    AddScore 50 + 50 * l008.State
End Sub

Sub Trigger013_Hit
    PlaySoundAt "fx_sensor", Trigger013
End Sub

'************************
'       Targets
'************************

Sub Target001_hit 'blue Lights
    PlaySoundAtBall "fx_target"
    If Tilted or bGameInPlay = False Then Exit Sub
    AddScore 50
    addbonus 1
    DropPost
    'turn on blue Lights
    l014.State = 1
    l018.State = 1
    l016.State = 1
    l020.State = 1
    tget1 = 1:Check4Targets
End Sub

Sub Target002_hit 'green Lights
    PlaySoundAtBall "fx_target"
    If Tilted or bGameInPlay = False Then Exit Sub
    AddScore 50
    addbonus 1
    DropPost
    'turn on green Lights
    l015.State = 1
    l007.State = 1
    l005.State = 1
    l002.State = 1
    tget2 = 1:Check4Targets
End Sub

Sub Target003_hit 'red Lights
    PlaySoundAtBall "fx_target"
    If Tilted or bGameInPlay = False Then Exit Sub
    AddScore 50
    addbonus 1
    UpPost
    'turn on red Lights
    l025.State = 1
    l026.State = 1
    l024.State = 1
    l017.State = 1
    l006.State = 1
    l001.State = 1
    tget3 = 1:Check4Targets
End Sub

Sub Target004_hit 'yellow Lights
    PlaySoundAtBall "fx_target"
    If Tilted or bGameInPlay = False Then Exit Sub
    AddScore 50
    Addbonus 1
    UpPost
    'turn on yellow Lights
    l022.State = 1
    l019.State = 1
    l008.State = 1
    tget4 = 1:Check4Targets
End Sub

Sub Check4Targets
    If tget1 + tget2 + tget3 + tget4 = 4 then
        AwardExtraBall
    End If
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

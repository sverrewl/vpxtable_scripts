' Alice in Wonderland / IPD No. 47 / August, 1948 / 1 Player
' D. Gottlieb & Company (1931-1977) [Trade Name: Gottlieb]
' VPX8 table by jpsalas 2024, version 6.0.0

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
Const TableName = "AliceInWonderland" ' file name to save highscores and other variables
Const cGameName = "AliceInWonderland" ' B2S name
Const MaxPlayers = 1                  ' 1 to 4 can play
Const MaxMultiplier = 2               ' limit bonus multiplier
Const FreePlay = False                ' Free play or coins
Const Special1 = 700000               ' extra ball or credit
Const Special2 = 750000
Const Special3 = 800000
Const Special4 = 850000
Const Special5 = 900000

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
Dim Special5Awarded(4)
Dim Score(4)
Dim HighScore
'Dim Match
Dim Tilt
Dim TiltSensitivity
Dim Tilted
Dim Add10
Dim Add100
Dim Add1000

' Control variables
Dim BallsOnPlayfield
Dim BallsToEject

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
    'Match = 0
    bJustStarted = True
    Add10 = 0
    Add100 = 0
    Add1000 = 0
    BallsToEject = 0

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
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound "fx_nudge", 0, 1, -0.1, 0.25:CheckTilt
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound "fx_nudge", 0, 1, 0.1, 0.25:CheckTilt
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound "fx_nudge", 0, 1, 1, 0.25:CheckTilt
    If keycode = MechanicalTilt Then Tilt = 16:CheckTilt

        ' teclas de los flipers
        If keycode = LeftFlipperKey Then SolLFlipper 1
        If keycode = RightFlipperKey Then SolRFlipper 1

        ' tecla de empezar el juego
        ' si el juego ha empezado saca una bola nueva
        If keycode = StartGameKey AND bAutoBallEject Then
            EjectBall
        End If
        If Keycode = RightMagnaSave And bAutoBallEject = False Then
            If bJustStarted Then
                FirstBall
            Else
                EjectBall
            End If
        End If
    Else ' If (GameInPlay)

        If keycode = StartGameKey Then
            If(bFreePlay = True) Then
                If(BallsOnPlayfield = 0) Then
                    ResetScores
                    ResetForNewGame()
                End If
            Else
                If(Credits> 0) Then
                    DOF 200, DOFOn
                    If(BallsOnPlayfield = 0) Then
                        Credits = Credits - 1
                        UpdateCredits
                        ResetScores
                        ResetForNewGame()
                    End If
                Else
                    ' Not Enough Credits to start a game.
                    'PlaySound "so_nocredits"
                    DOF 200, DOFOff
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
        LeftFlipper.RotateToEnd
        LeftFlipper001.RotateToEnd
        LeftFlipper002.RotateToEnd
    Else
        PlaySoundAt SoundFXDOF("fx_flipperdown", 101, DOFOff, DOFFlippers), LeftFlipper
        LeftFlipper.RotateToStart
        LeftFlipper001.RotateToStart
        LeftFlipper002.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFXDOF("fx_flipperup", 102, DOFOn, DOFFlippers), RightFlipper
        RightFlipper.RotateToEnd
        RightFlipper001.RotateToEnd
        RightFlipper002.RotateToEnd
    Else
        PlaySoundAt SoundFXDOF("fx_flipperdown", 102, DOFOff, DOFFlippers), RightFlipper
        RightFlipper.RotateToStart
        RightFlipper001.RotateToStart
        RightFlipper002.RotateToStart
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

Sub LeftFlipper002_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, Vol(ActiveBall), pan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper002_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, Vol(ActiveBall), pan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
End Sub

' flipper primitive animations vpx8
Sub LeftFlipper_Animate:LeftFlipperRubber.RotZ = LeftFlipper.CurrentAngle:LeftFlipperP.RotZ = LeftFlipper.CurrentAngle:End Sub
Sub RightFlipper_Animate:RightFlipperRubber.RotZ = RightFlipper.CurrentAngle:RightFlipperP.RotZ = RightFlipper.CurrentAngle:End Sub

Sub LeftFlipper001_Animate:LeftFlipperRubber001.RotZ = LeftFlipper001.CurrentAngle:LeftFlipperP001.RotZ = LeftFlipper001.CurrentAngle:End Sub
Sub RightFlipper001_Animate:RightFlipperRubber001.RotZ = RightFlipper001.CurrentAngle:RightFlipperP001.RotZ = RightFlipper001.CurrentAngle:End Sub

Sub LeftFlipper002_Animate:LeftFlipperRubber002.RotZ = LeftFlipper002.CurrentAngle:LeftFlipperP002.RotZ = LeftFlipper002.CurrentAngle:End Sub
Sub RightFlipper002_Animate:RightFlipperRubber002.RotZ = RightFlipper002.CurrentAngle:RightFlipperP002.RotZ = RightFlipper002.CurrentAngle:End Sub

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

'**************
'    TILT
'**************

Sub CheckTilt
    Tilt = Tilt + TiltSensitivity
    TiltDecreaseTimer.Enabled = True
    If Tilt> 15 Then
        Tilted = True
        TurnON TiltReel
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
        TurnOffPlayfieldLights
        GiOff
        LeftFlipper.RotateToStart
        RightFlipper.RotateToStart
        LeftFlipper001.RotateToStart
        RightFlipper001.RotateToStart
        LeftFlipper002.RotateToStart
        RightFlipper002.RotateToStart
    Else
        GiOn
    End If
End Sub

Sub TiltRecoveryTimer_Timer()
    ' all the balls have drained
    If(BallsOnPlayfield = 0) Then
        'End the game with the next line
        BallsRemaining(CurrentPlayer) = 0
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
'       JP's VP10 Rolling Sounds
'***********************************************

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
Sub aBumpers_Hit(idx):PlaySoundAtBall "fx_passive_bumper":End Sub

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
        Special4Awarded(i) = False
        Special5Awarded(i) = False
        BallsRemaining(i) = BallsPerGame
    Next
    BonusMultiplier = 1
    Bonus = 1
    UpdateBallInPlay
    BallsToEject = BallsPerGame

    'Clear_Match

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
    ' ajusta la mesa para una bola nueva, sube las dianas abatibles, etc
    ResetForNewPlayerBall()
    ' crea una bola nueva en la zona del plunger si está en modo automático
    If bAutoBallEject Then
        BallsToEject = BallsToEject - 1
        CreateNewBall()
    End If
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
    balllock3.IsDropped = 0:balllock3.TimerEnabled = 1
    BallRelease.CreateSizedBallWithMass BallSize / 2, BallMass
    BallsOnPlayfield = BallsOnPlayfield + 1
    UpdateBallInPlay
    PlaySoundAt SoundFXDOF("fx_ejectball", 120, DOFPulse, DOFContactors), BallRelease
    BallRelease.Kick 270, 2
End Sub

Sub balllock3_Timer
    balllock3.TimerEnabled = 0
    balllock3.IsDropped = 1
End Sub

' player lost the ball

Sub EndOfBall()
    'debug.print "EndOfBall"
    Dim AwardPoints, TotalBonus, ii
    AwardPoints = 0
    TotalBonus = 0
    ' Lost the first ball, now it cannot accept more players
    bOnTheFirstBall = False

    'No bonus in this table at the end of a ball

    ' add bonus if no tilt
    ' tilt system will take care of the next ball

    'If NOT Tilted Then
    '    BonusCountTimer.Interval = 200 + 50 * BonusMultiplier
    '    BonusCountTimer.Enabled = 1
    'Else
    vpmtimer.addtimer 400, "EndOfBall2 '"
'End If
End Sub

Sub BonusCountTimer_Timer
    'debug.print "BonusCount_Timer"
    If Bonus> 0 Then
        Bonus = Bonus -1
        AddScore 10000 * BonusMultiplier
        PlaySound "fx_relay"
        UpdateBonusLights
    Else
        BonusCountTimer.Enabled = 0
        If BonusMultiplier = 2 Then
            Kicker1.Kick 354, 18
        Else
            Kicker2.kick 175, 8
        End If
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
    TurnOFF TiltReel
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
        'If bFreePlay = False Then
        '    Verification_Match
        'End If

        ' end of game
        EndOfGame()
    Else
        ' next player
        CurrentPlayer = NextPlayer

        ' update score
        AddScore 0

        ' reset table for new player
        ResetForNewPlayerBall()

        ' y sacamos una bola si está en modo automatico
        If bAutoBallEject Then
            BallsToEject = BallsToEject - 1
            CreateNewBall()
        End If
    End If
End Sub

Sub EjectBall 'only for this table
    If BallsToEject> 0 Then
        BallsToEject = BallsToEject - 1
        CreateNewBall()
    End If
End Sub

' Called at the end of the game

Sub EndOfGame()
    'debug.print "EndOfGame"
    bGameInPLay = False
    bJustStarted = False
    'If B2SOn then
    '    Controller.B2SSetGameOver 1
    '    Controller.B2SSetBallInPlay 0
    '    Controller.B2SSetPlayerUp 0
    '    Controller.B2SSetCanPlay 0
    'end if
    ' turn off flippers
    SolLFlipper 0
    SolRFlipper 0

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
'PlaySound "fx_match"
'Match = INT(RND(1) * 10) * 10 ' random between 10 and 90
'Display_Match
'If(Score(CurrentPlayer)MOD 100) = Match Then
'    PlaySound SoundFXDOF("fx_knocker", 122, DOFPulse, DOFknocker)
'    AddCredits 1
'End If
End Sub

Sub Clear_Match()
'    Match = 0
'    Update_Match
'    If B2SOn then
'        Controller.B2SSetMatch Match
'    end if
End Sub

Sub Display_Match()
'    Update_Match
'    If B2SOn then
'        If Match = 0 then
'            Controller.B2SSetMatch 100
'        else
'            Controller.B2SSetMatch Match
'        end if
'    end if
End Sub

Sub Update_Match
'    Select Case Match \ 10
'Case 0:m0.State = 0:m000.state = 0:m1.State = 0:m001.state = 0:m2.State = 0:m002.state = 0:m3.State = 0:m003.state = 0:m4.State = 0:m004.state = 0:m5.State = 0:m005.state = 0:m6.State = 0:m006.state = 0:m7.State = 0:m007.state = 0:m8.State = 0:m008.state = 0:m9.State = 0:m009.state = 0
'Case 1:m0.State = 0:m000.state = 0:m1.State = 1:m001.state = 1:m2.State = 0:m002.state = 0:m3.State = 0:m003.state = 0:m4.State = 0:m004.state = 0:m5.State = 0:m005.state = 0:m6.State = 0:m006.state = 0:m7.State = 0:m007.state = 0:m8.State = 0:m008.state = 0:m9.State = 0:m009.state = 0
'Case 2:m0.State = 0:m000.state = 0:m1.State = 0:m001.state = 0:m2.State = 1:m002.state = 1:m3.State = 0:m003.state = 0:m4.State = 0:m004.state = 0:m5.State = 0:m005.state = 0:m6.State = 0:m006.state = 0:m7.State = 0:m007.state = 0:m8.State = 0:m008.state = 0:m9.State = 0:m009.state = 0
'Case 3:m0.State = 0:m000.state = 0:m1.State = 0:m001.state = 0:m2.State = 0:m002.state = 0:m3.State = 1:m003.state = 1:m4.State = 0:m004.state = 0:m5.State = 0:m005.state = 0:m6.State = 0:m006.state = 0:m7.State = 0:m007.state = 0:m8.State = 0:m008.state = 0:m9.State = 0:m009.state = 0
'Case 4:m0.State = 0:m000.state = 0:m1.State = 0:m001.state = 0:m2.State = 0:m002.state = 0:m3.State = 0:m003.state = 0:m4.State = 1:m004.state = 1:m5.State = 0:m005.state = 0:m6.State = 0:m006.state = 0:m7.State = 0:m007.state = 0:m8.State = 0:m008.state = 0:m9.State = 0:m009.state = 0
'Case 5:m0.State = 0:m000.state = 0:m1.State = 0:m001.state = 0:m2.State = 0:m002.state = 0:m3.State = 0:m003.state = 0:m4.State = 0:m004.state = 0:m5.State = 1:m005.state = 1:m6.State = 0:m006.state = 0:m7.State = 0:m007.state = 0:m8.State = 0:m008.state = 0:m9.State = 0:m009.state = 0
'Case 6:m0.State = 0:m000.state = 0:m1.State = 0:m001.state = 0:m2.State = 0:m002.state = 0:m3.State = 0:m003.state = 0:m4.State = 0:m004.state = 0:m5.State = 0:m005.state = 0:m6.State = 1:m006.state = 1:m7.State = 0:m007.state = 0:m8.State = 0:m008.state = 0:m9.State = 0:m009.state = 0
'Case 7:m0.State = 0:m000.state = 0:m1.State = 0:m001.state = 0:m2.State = 0:m002.state = 0:m3.State = 0:m003.state = 0:m4.State = 0:m004.state = 0:m5.State = 0:m005.state = 0:m6.State = 0:m006.state = 0:m7.State = 1:m007.state = 1:m8.State = 0:m008.state = 0:m9.State = 0:m009.state = 0
'Case 8:m0.State = 0:m000.state = 0:m1.State = 0:m001.state = 0:m2.State = 0:m002.state = 0:m3.State = 0:m003.state = 0:m4.State = 0:m004.state = 0:m5.State = 0:m005.state = 0:m6.State = 0:m006.state = 0:m7.State = 0:m007.state = 0:m8.State = 1:m008.state = 1:m9.State = 0:m009.state = 0
'Case 9:m0.State = 0:m000.state = 0:m1.State = 0:m001.state = 0:m2.State = 0:m002.state = 0:m3.State = 0:m003.state = 0:m4.State = 0:m004.state = 0:m5.State = 0:m005.state = 0:m6.State = 0:m006.state = 0:m7.State = 0:m007.state = 0:m8.State = 0:m008.state = 0:m9.State = 1:m009.state = 1
'    End Select
End Sub

' *********************************************************************
'                      Drain / Plunger Functions
' *********************************************************************

Sub Drain_Hit() 'called from the draintrigger on the last ball
    If bGameInPlay Then
        BallsOnPlayfield = BallsOnPlayfield - 1
        PlaySoundAt "fx_kicker_enter", Drain
        'tilted?
        If Tilted Then
            StopEndOfBallMode
        Else
            EndOfBall
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
        Case 0
            UpdateScore points
        Case 10000
            Score(CurrentPlayer) = Score(CurrentPlayer) + points
            UpdateScore points
            CheckRandomBumperCount
        Case 20000, 30000, 40000, 50000
            Add10 = Add10 + points \ 10000
            AddScore10Timer.Enabled = TRUE
    End Select
    CheckScoreAwards
End Sub

Sub CheckScoreAwards
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
    If Score(CurrentPlayer) >= Special5 AND Special5Awarded(CurrentPlayer) = False Then
        AwardSpecial
        Special5Awarded(CurrentPlayer) = True
    End If
End Sub

'************************************
'       Score sound Timers
'************************************

Sub AddScore10Timer_Timer()
    if Add10> 0 then
        AddScore 10000
        Add10 = Add10 - 1
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
        PlaySound "BellB"
        If Bonus> 10 Then
            Bonus = 10
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
    'Case 1:li2.State = 1:li3.State = 0:li4.State = 0
    'Case 2:li2.State = 0:li3.State = 1:li4.State = 0
    'Case 3:li2.State = 0:li3.State = 0:li4.State = 1
    End Select
End Sub

'***********************************************************************************
'        Score reels - puntuaciones - y actualiza otras luces del backdrop
'***********************************************************************************
'esta es al rutina que actualiza la puntuación del jugador
' solo 1 jugador

Sub ClearB2SScores 'Originally code from STAT
    Dim i
    For i = 1 to 9:Controller.B2SSetData i, 0:Next
    For i = 10 to 90 Step 10:Controller.B2SSetData i, 0:Next
    Controller.B2SSetData 205, 0
End Sub

Sub UpdateScore(playerpoints)
    dim tmp, B2sid

    '10.000 points
    tmp = Score(CurrentPlayer) \ 10000
    Select Case tmp MOD 10
        Case 0:t001.State = 0:t002.State = 0:t003.State = 0:t004.State = 0:t005.State = 0:t006.State = 0:t007.State = 0:t008.State = 0:t009.State = 0
            If B2SOn Then
                Controller.B2SSetData 211, 0
                Controller.B2SSetData 212, 0
                Controller.B2SSetData 213, 0
                Controller.B2SSetData 214, 0
                Controller.B2SSetData 215, 0
                Controller.B2SSetData 216, 0
            End If
        Case 1:t001.State = 1:t002.State = 0:t003.State = 0:t004.State = 0:t005.State = 0:t006.State = 0:t007.State = 0:t008.State = 0:t009.State = 0
            If B2SOn Then
                Controller.B2SSetData 211, 1
                Controller.B2SSetData 212, 0
                Controller.B2SSetData 213, 0
                Controller.B2SSetData 214, 0
                Controller.B2SSetData 215, 0
                Controller.B2SSetData 216, 0
            End If
        Case 2:t001.State = 0:t002.State = 1:t003.State = 0:t004.State = 0:t005.State = 0:t006.State = 0:t007.State = 0:t008.State = 0:t009.State = 0
            If B2SOn Then
                Controller.B2SSetData 211, 1
                Controller.B2SSetData 212, 0
                Controller.B2SSetData 213, 0
                Controller.B2SSetData 214, 0
                Controller.B2SSetData 215, 0
                Controller.B2SSetData 216, 0
            End If
        Case 3:t001.State = 0:t002.State = 0:t003.State = 1:t004.State = 0:t005.State = 0:t006.State = 0:t007.State = 0:t008.State = 0:t009.State = 0
            If B2SOn Then
                Controller.B2SSetData 211, 1
                Controller.B2SSetData 212, 0
                Controller.B2SSetData 213, 0
                Controller.B2SSetData 214, 0
                Controller.B2SSetData 215, 0
                Controller.B2SSetData 216, 0
            End If
        Case 4:t001.State = 0:t002.State = 0:t003.State = 0:t004.State = 1:t005.State = 0:t006.State = 0:t007.State = 0:t008.State = 0:t009.State = 0
            If B2SOn Then
                Controller.B2SSetData 211, 1
                Controller.B2SSetData 212, 0
                Controller.B2SSetData 213, 0
                Controller.B2SSetData 214, 0
                Controller.B2SSetData 215, 0
                Controller.B2SSetData 216, 0
            End If
        Case 5:t001.State = 0:t002.State = 0:t003.State = 0:t004.State = 0:t005.State = 1:t006.State = 0:t007.State = 0:t008.State = 0:t009.State = 0
            If B2SOn Then
                Controller.B2SSetData 211, 0
                Controller.B2SSetData 212, 1
                Controller.B2SSetData 213, 0
                Controller.B2SSetData 214, 0
                Controller.B2SSetData 215, 0
                Controller.B2SSetData 216, 0
            End If
        Case 6:t001.State = 0:t002.State = 0:t003.State = 0:t004.State = 0:t005.State = 0:t006.State = 1:t007.State = 0:t008.State = 0:t009.State = 0
            If B2SOn Then
                Controller.B2SSetData 211, 0
                Controller.B2SSetData 212, 0
                Controller.B2SSetData 213, 1
                Controller.B2SSetData 214, 0
                Controller.B2SSetData 215, 0
                Controller.B2SSetData 216, 0
            End If
        Case 7:t001.State = 0:t002.State = 0:t003.State = 0:t004.State = 0:t005.State = 0:t006.State = 0:t007.State = 1:t008.State = 0:t009.State = 0
            If B2SOn Then
                Controller.B2SSetData 211, 0
                Controller.B2SSetData 212, 0
                Controller.B2SSetData 213, 0
                Controller.B2SSetData 214, 1
                Controller.B2SSetData 215, 0
                Controller.B2SSetData 216, 0
            End If
        Case 8:t001.State = 0:t002.State = 0:t003.State = 0:t004.State = 0:t005.State = 0:t006.State = 0:t007.State = 0:t008.State = 1:t009.State = 0
            If B2SOn Then
                Controller.B2SSetData 211, 0
                Controller.B2SSetData 212, 0
                Controller.B2SSetData 213, 0
                Controller.B2SSetData 214, 0
                Controller.B2SSetData 215, 1
                Controller.B2SSetData 216, 0
            End If
        Case 9:t001.State = 0:t002.State = 0:t003.State = 0:t004.State = 0:t005.State = 0:t006.State = 0:t007.State = 0:t008.State = 0:t009.State = 1
            If B2SOn Then
                Controller.B2SSetData 211, 0
                Controller.B2SSetData 212, 0
                Controller.B2SSetData 213, 0
                Controller.B2SSetData 214, 0
                Controller.B2SSetData 215, 0
                Controller.B2SSetData 216, 1
            End If
    End Select
    If B2SOn Then
        ClearB2SScores
        B2sid = (tmp MOD 10)
        Controller.B2SSetData B2sid, 1
    End If

    '100.000 points
    tmp = Score(CurrentPlayer) \ 100000
    Select Case tmp MOD 10
        Case 0:t010.State = 0:t011.State = 0:t012.State = 0:t013.State = 0:t014.State = 0:t015.State = 0:t016.State = 0:t017.State = 0:t018.State = 0
        Case 1:t010.State = 1:t011.State = 0:t012.State = 0:t013.State = 0:t014.State = 0:t015.State = 0:t016.State = 0:t017.State = 0:t018.State = 0
        Case 2:t010.State = 0:t011.State = 1:t012.State = 0:t013.State = 0:t014.State = 0:t015.State = 0:t016.State = 0:t017.State = 0:t018.State = 0
        Case 3:t010.State = 0:t011.State = 0:t012.State = 1:t013.State = 0:t014.State = 0:t015.State = 0:t016.State = 0:t017.State = 0:t018.State = 0
        Case 4:t010.State = 0:t011.State = 0:t012.State = 0:t013.State = 1:t014.State = 0:t015.State = 0:t016.State = 0:t017.State = 0:t018.State = 0
        Case 5:t010.State = 0:t011.State = 0:t012.State = 0:t013.State = 0:t014.State = 1:t015.State = 0:t016.State = 0:t017.State = 0:t018.State = 0
        Case 6:t010.State = 0:t011.State = 0:t012.State = 0:t013.State = 0:t014.State = 0:t015.State = 1:t016.State = 0:t017.State = 0:t018.State = 0
        Case 7:t010.State = 0:t011.State = 0:t012.State = 0:t013.State = 0:t014.State = 0:t015.State = 0:t016.State = 1:t017.State = 0:t018.State = 0
        Case 8:t010.State = 0:t011.State = 0:t012.State = 0:t013.State = 0:t014.State = 0:t015.State = 0:t016.State = 0:t017.State = 1:t018.State = 0
        Case 9:t010.State = 0:t011.State = 0:t012.State = 0:t013.State = 0:t014.State = 0:t015.State = 0:t016.State = 0:t017.State = 0:t018.State = 1
    End Select
    If B2SOn Then
        B2sid = (tmp MOD 10) * 10
        Controller.B2SSetData B2sid, 1
    End If
    '1 million
    If Score(CurrentPlayer)> 1000000 Then
        TurnON millionReel
        If B2SOn Then
            Controller.B2SSetData 205, 1
        End If
    Else
        TurnOFF millionReel
        If B2SOn Then
            Controller.B2SSetData 205, 0
        End If
    End If
End Sub

' pone todos los marcadores a 0
Sub ResetScores
    'ScoreReel1.ResetToZero
    AddScore 0
End Sub

Sub AddCredits(value)
    If Credits <25 Then
        Credits = Credits + value
        UpdateCredits
    end if
End Sub

Sub UpdateCredits
    Dim i
    PlaySound "fx_relay"
    TurnOFF Credits0
    TurnOFF Credits10
    i = credits MOD 10
    debug.print i
    Credits0(i).State = 1
    i = (Credits \ 10) MOD 10
    debug.print i
    Credits10(i).State = 1
    If B2SOn then Controller.B2SSetMatch Credits + 100 'uses the match B2s name
End Sub

Sub UpdateBallInPlay                                   'actualiza los marcadores de las bolas, el número de jugador y el número total de jugadores
End Sub

Sub TurnON(Col)                                        'turn on all the lights in a collection
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
    PlaySound "BellC"
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
    Dim x
    bAttractMode = True
    For each x in aLights
        x.State = 2
    Next
'GameOverReel.SetValue 1
'update current player and balls
End Sub

Sub StopAttractMode()
    Dim x
    bAttractMode = False
    TurnOffPlayfieldLights
    ResetScores
'GameOverReel.SetValue 0
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

' default high scores, remove this when the scores are available from the config file
HSScore(1) = 180000
HSScore(2) = 160000
HSScore(3) = 140000
HSScore(4) = 120000
HSScore(5) = 100000

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
        PlaySound "fx_knocker" 'but do not give an extra credit
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
    'create the 5 Balls
    balllock.IsDropped = 0
    balllock2.IsDropped = 0
    balllock3.IsDropped = 1
    kicker003.CreateSizedBallWithMass BallSize / 2, BallMass:kicker003.kick 90, 2
    vpmTimer.AddTimer 200, "kicker003.CreateSizedBallWithMass BallSize / 2, BallMass: kicker003.kick 90,2 '"
    vpmTimer.AddTimer 400, "kicker003.CreateSizedBallWithMass BallSize / 2, BallMass: kicker003.kick 90,2 '"
    vpmTimer.AddTimer 600, "kicker003.CreateSizedBallWithMass BallSize / 2, BallMass: kicker003.kick 90,2 '"
    vpmTimer.AddTimer 800, "kicker003.CreateSizedBallWithMass BallSize / 2, BallMass: kicker003.kick 90,2 '"
    TurnOffPlayfieldLights()
End Sub

' Dim all the variables
Dim Special
Dim RedSpecial
Dim RedBumperValue
Dim CenterBumperCounterMax
Dim CenterBumperCounter
Dim PinkBumperCounterMax
Dim PinkBumperCounter

Sub Game_Init() 'called at the start of a new game
    'Start music?
    'Init variables?
    Special = False
    RedSpecial = False
    RedBumperValue = 20000
    'Start or init timers
    PlaySoundAt "fx_releaseballs", drain_dummy
    balllock.IsDropped = 1
    balllock2.IsDropped = 1
    balllock.TimerEnabled = 1
    'Init lights?
    TurnOffPlayfieldLights()
    buttonlight.State = 1
    'turn on red bumpers
    BumperL004.State = 1
    BumperL005.State = 1
    BumperL006.State = 1
    BumperL007.State = 1
    'turn on ALICE lights
    LightA.State = 1
    BumperL002.State = 1
    BumperL002a.State = 1
    BumperL001.State = 1
    BumperL001a.State = 1
    BumperL003.State = 1
    BumperL003a.State = 1
    LightE.State = 1
    'and the ALICE backdrop Lights
    LightA2.State = 1
    LightL2.State = 1
    LightI2.State = 1
    LightC2.State = 1
    LightE2.State = 1
    If B2SOn Then
        Controller.B2SSetData 200, 1
        Controller.B2SSetData 201, 1
        Controller.B2SSetData 202, 1
        Controller.B2SSetData 203, 1
        Controller.B2SSetData 204, 1
        Controller.B2SSetData 211, 0
        Controller.B2SSetData 212, 0
        Controller.B2SSetData 213, 0
        Controller.B2SSetData 214, 0
        Controller.B2SSetData 215, 0
        Controller.B2SSetData 216, 0
    End If
    UpdateRedBumperValue
    ' Initialize Pink and Center Bumper Counters with Random numbers between 10 and 17
    LightCenterBumper.State = 0            ' Turn Off the Bumper
    CenterBumperCounterMax = RndNbr(7) + 2 ' Pick an Initial Random # between 2 and 9
    CenterBumperCounter = 0                ' Reset Counter
    LightA001.State = 0                    ' Turn Off the Pink Triangles
    LightE001.State = 0
    PinkBumperCounterMax = RndNbr(7) + 2   ' Pick an Initial Random # between 2 and 9
    PinkBumperCounter = 0                  ' Reset Counter
End Sub

Sub StopEndOfBallMode()                    'called when the last ball is drained
End Sub

Sub ResetNewBallVariables()                'init variables & lights new ball/player
'TurnOffPlayfieldLights
'BonusMultiplier = 1
'Bonus = 1:UpdateBonusLights
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

' New drain/trough

Sub drain_dummy_Hit
    Me.destroyball
End sub

Sub balllock_Timer
    Me.TimerEnabled = 0
    balllock.IsDropped = 0
    balllock2.IsDropped = 0
End Sub

' Top Bumpers

Sub Bumper002_Hit 'L
    If Tilted Then Exit Sub
    AddBonus 1
    BumperL002.State = 0
    BumperL002a.State = 0
    LightL2.State = 0
    If B2SOn Then
        Controller.B2SSetData 201, 0
    End If
    CheckALICE
    BlinkRedBumpers
End Sub

Sub Bumper001_Hit 'I
    If Tilted Then Exit Sub
    AddBonus 1
    BumperL001.State = 0
    BumperL001a.State = 0
    LightI2.State = 0
    If B2SOn Then
        Controller.B2SSetData 202, 0
    End If
    CheckALICE
End Sub

Sub Bumper003_Hit 'C
    If Tilted Then Exit Sub
    AddBonus 1
    BumperL003.State = 0
    BumperL003a.State = 0
    LightC2.State = 0
    If B2SOn Then
        Controller.B2SSetData 203, 0
    End If
    CheckALICE
    BlinkRedBumpers
End Sub

' Lower Bumpers
Sub Bumper008_Hit
    If Tilted Then Exit Sub
    If Special Then
        AwardSpecial
    Else
        AddScore 10000
        PlaySound "BellA5"
    End If
End Sub

Sub Bumper009_Hit
    If Tilted Then Exit Sub
    If Special Then
        AwardSpecial
    Else
        AddScore 10000
        PlaySound "BellA5"
    End If
    BlinkRedBumpers
End Sub

' Side lanes

Sub Trigger003_Hit 'A
    If Tilted Then Exit Sub
    PLaySoundAt "fx_sensor", Trigger003
    AddBonus 1
    LightA.State = 0
    LightA2.State = 0
    If B2SOn Then
        Controller.B2SSetData 200, 0
    End If
    CheckALICE
    AdvanceRedBumper
    BlinkRedBumpers
    If LightA001.State Then AwardSpecial
End Sub

Sub Trigger004_Hit 'E
    If Tilted Then Exit Sub
    PLaySoundAt "fx_sensor", Trigger004
    AddBonus 1
    LightE.State = 0
    LightE2.State = 0
    If B2SOn Then
        Controller.B2SSetData 204, 0
    End If
    CheckALICE
    AdvanceRedBumper
    BlinkRedBumpers
    If LightE001.State Then AwardSpecial
End Sub

Sub CheckALICE ' check all the ALICE lights are off, then turn on Special
    Dim tmp
    ' Special
    tmp = BumperL001.State + BumperL002.State + BumperL003.State + LightA.State + LightE.State
    If tmp = 0 Then
        Special = True
        'turn on special w/l bumpers
        BumperL008.State = 1
        BumperL009.State = 1
        'double Bonus = special
        dbl1.State = 1
    End If
End Sub

' other lanes

Sub Trigger002_Hit 'Top center
    If Tilted Then Exit Sub
    PLaySoundAt "fx_sensor", Trigger004
    AdvanceRedBumper
    BlinkRedBumpers
End Sub

Sub Trigger001_Hit 'Lower center add bonus
    If Tilted Then Exit Sub
    PLaySoundAt "fx_sensor", Trigger004
    AddBonus 1
End Sub

' RED bumpers

Sub Bumper004_Hit
    If Tilted Then Exit Sub
    If RedSpecial Then
        AwardSpecial
    Else
        AddScore RedBumperValue
        PlaySound "BellA5"
    End If
    BlinkRedBumpers
End Sub

Sub Bumper005_Hit
    If Tilted Then Exit Sub
    If RedSpecial Then
        AwardSpecial
    Else
        AddScore RedBumperValue
        PlaySound "BellA5"
    End If
    BlinkRedBumpers
End Sub

Sub Bumper006_Hit
    If Tilted Then Exit Sub
    If RedSpecial Then
        AwardSpecial
    Else
        AddScore RedBumperValue
        PlaySound "BellA5"
    End If
    BlinkRedBumpers
End Sub

Sub Bumper007_Hit
    If Tilted Then Exit Sub
    If RedSpecial Then
        AwardSpecial
    Else
        AddScore RedBumperValue
        PlaySound "BellA5"
    End If
    BlinkRedBumpers
End Sub

Sub BlinkRedBumpers
    BumperL004.Duration 0, 350, 1
    BumperL005.Duration 0, 350, 1
    BumperL006.Duration 0, 350, 1
    BumperL007.Duration 0, 350, 1
End Sub

Sub AdvanceRedBumper
    RedBumperValue = RedBumperValue + 10000
    If RedBumperValue> 60000 Then RedBumperValue = 60000
    If RedBumperValue = 60000 Then
        RedBumperValue = 60000 'just to enable the special light, but it will not be scored as the Special will be scored
        RedSpecial = True
    End If
    UpdateRedBumperValue
End Sub

Sub UpdateRedBumperValue
    Select Case RedBumperValue
        Case 10000:rbl1.State = 0:rbl2.State = 0:rbl3.State = 0:rbl4.State = 0:rbl5.State = 0
        Case 20000:rbl1.State = 1:rbl2.State = 0:rbl3.State = 0:rbl4.State = 0:rbl5.State = 0
        Case 30000:rbl1.State = 0:rbl2.State = 1:rbl3.State = 0:rbl4.State = 0:rbl5.State = 0
        Case 40000:rbl1.State = 0:rbl2.State = 0:rbl3.State = 1:rbl4.State = 0:rbl5.State = 0
        Case 50000:rbl1.State = 0:rbl2.State = 0:rbl3.State = 0:rbl4.State = 1:rbl5.State = 0
        Case 60000:rbl1.State = 0:rbl2.State = 0:rbl3.State = 0:rbl4.State = 0:rbl5.State = 1
    End Select
End Sub

' Center Bumper

Sub CheckRandomBumperCount                        ' Fake Random/Asynchronous Bumper Specials
    CenterBumperCounter = CenterBumperCounter + 1 ' Bump Counter
    If(CenterBumperCounter >= CenterBumperCounterMax) Then
        LightCenterBumper.State = 1               ' Turn On the Center Bumper
        CenterBumperCounterMax = RndNbr(7) + 10   ' Pick a new Random # between 10 and 17
        CenterBumperCounter = 0                   ' Reset Counter
    Else LightCenterBumper.State = 0              ' Turn off Center Bumper if not at Max Count
    End If
    PinkBumperCounter = PinkBumperCounter + 1     ' Bump Counter
    If(PinkBumperCounter >= PinkBumperCounterMax) Then
        LightA001.State = 1                       ' Turn On the Pink Triangles
        LightE001.State = 1
        PinkBumperCounterMax = RndNbr(7) + 10     ' Pick a new Random # between 10 and 17
        PinkBumperCounter = 0                     ' Reset Counter
    Else
        LightA001.State = 0                       ' Turn Off the Pink Triangles if not at Max Count
        LightE001.State = 0
    End If
End Sub

Sub BumperCenter_Hit
    If Tilted Then Exit Sub
    BlinkRedBumpers
    PlaySound "BellA5"
    If LightCenterBumper.State Then
        AddScore 50000
    Else
        AddScore 10000
    End If
End Sub

' Holes

Sub Kicker1_Hit 'double bonus
    If NOT Tilted Then
        If Special Then
            AwardSpecial
        End If
        BonusMultiplier = 2
        BonusCountTimer.Interval = 300
        BonusCountTimer.Enabled = 1
    Else
        kicker1.Kick 354, 18
    End If
End Sub

Sub Kicker2_Hit 'double bonus
    If NOT Tilted Then
        BonusMultiplier = 1
        BonusCountTimer.Interval = 200
        BonusCountTimer.Enabled = 1
    Else
        kicker2.Kick 170, 8
    End If
End Sub

'*********************************
' Table Options F12 User Options
'*********************************
' Table1.Option arguments are:
' - option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional array of literal strings

Dim LUTImage, BallsPerGame, bAutoBallEject

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

    ' Auto Ball Eject
    x = Table1.Option("Auto Ball Loading", 0, 1, 1, 0, 0, Array("No", "Yes") )
    If x = 1 Then bAutoBallEject = True Else bAutoBallEject = False

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

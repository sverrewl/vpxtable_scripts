' ******************************************************************
'       VPX 10.8 - version by Klodo81 2023, Skylab version 3.0
'
'                     Based on:
'                     JPSalas Basic EM script
'                       JPSalas VPX8 Physics V3.1
'
'          Skylab / IPD No.2202 / June 14, 1974 / 1 Player
'
' ******************************************************************
'
'V1.1 mod Table
' - Added rolling sound
'
'V1.2 mod Table
' - "Balls To Play" on the Backglass counting down rather than up
' - HorseshoeTimer must be disabled by default to have 2/3 balls instead of 0/5 balls in Horseshoe
' - If 100000 Lit, it stay lit until attrack mode stop
'
'V1.3 mod Table
' - Added Playfield Mesh
' - Updated sound files fx_ballrolling2 to 7 as ballrolling1
' - Some others things
'
'V2.0 mod Table
'- Updated for VPX 10.8
'- Added menu F12 with options and LUT
'- Modif lights and GI
'- Some others things
'
'V3.0 mod Table
'- JP's VPX8 Physics V3.1
'- Credit until 25 instead 15
'
' ******************************************************************

Option Explicit
Randomize

' DOF config
'
' Flippers L/R - 101/102
' Slingshots L/R - 103/104
' Bumpers L to R - 105/106
' Triggers - 201 to 208
' Knocker - 110
' Chimes - 301/302/303
' Drain - 250
' Bonus - 210

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

'*******************************************
'  User Options
'*******************************************

' All user options and LUT are accessibles from the F12 menu

'*******************************************
'  Constants and Global Variables
'*******************************************

' Constants
Const TableName = "Skylab_74VPX"   ' file name to save highscores and other variables
Const cGameName = "Skylab1974"   ' B2S name
Const MaxPlayers = 1        ' Only 1 can play
Const MaxMultiplier = 2   ' limit bonus multiplier
Const Special1 = 73000      ' 3 Balls award credit
Const Special2 = 94000      ' 3 Balls award credit
Const Special3 = 137000      ' 5 Balls award credit
Const Special4 = 178000      ' 5 Balls award credit

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
Dim Add1000
Dim Add10000
Dim CBonus
Dim x
Dim LeftBalls, RightBalls

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

    ' load highscores and credits
    Loadhs

  If HSScore(1) >= 100000 Then
        RolloverReel1.setvalue 1
    ScoreReel1.SetValue (HSScore(1) - 100000)
  Else
    ScoreReel1.SetValue HSScore(1)
    End If
    If B2SOn then
        If HSScore(1) >= 100000 then
            Controller.B2SSetScoreRolloverPlayer1 1
      Controller.B2SSetScorePlayer 1, (HSScore(1) - 100000)
    Else
      Controller.B2SSetScorePlayer 1, HSScore(1)
        end if
    end if
    UpdateCredits

    ' init all the global variables
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
    Add1000 = 0
  Add10000 = 0

    ' setup table in game over mode
    SolLFlipper 0
    SolRFlipper 0

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

  'Init Horseshoe
    leftballs=2
    rightballs=3
    KcaptiveL.uservalue=1
    KcaptiveL.timerenabled=1

End Sub

' **************
'   Horseshoe
' **************

Sub KcaptiveL_timer 'create 5 new balls in the horseshoe
        Select Case (me.uservalue)
            Case 1:
                if leftballs>0 then
                    KcaptiveL.createball
                    KcaptiveL.kick 180,1
                end if
                if rightballs>0 then
                    KcaptiveR.createball
                    kcaptiveR.kick 180,1
                end if
            Case 2:
                if leftballs>1 then
                    KcaptiveL.createball
                    KcaptiveL.kick 180,1
                end if
                if rightballs>1 then
                    KcaptiveR.createball
                    kcaptiveR.kick 180,1
                end if
            Case 3:
                if leftballs>2 then
                    KcaptiveL.createball
                    KcaptiveL.kick 180,1
                end if
                if rightballs>2 then
                    KcaptiveR.createball
                    kcaptiveR.kick 180,1
                end if
            Case 4:
                if leftballs>3 then
                    KcaptiveL.createball
                    KcaptiveL.kick 180,1
                end if
                if rightballs>3 then
                    KcaptiveR.createball
                    kcaptiveR.kick 180,1
                end if
            Case 5:
                if leftballs>4 then
                    KcaptiveL.createball
                    KcaptiveL.kick 180,1
                end if
                if rightballs>4 then
                    KcaptiveR.createball
                    kcaptiveR.kick 180,1
                end if
            Case 6:
                me.timerenabled=0
                HorseshoeTimer.enabled=1
        End Select
    me.uservalue=me.uservalue+1
end sub

Sub HorseshoeTimer_timer 'track balls number on leftballs and RightBalls

    select case (TGcaptiveL.BallCntOver+TGcaptiveL1.BallCntOver+TGcaptiveL2.BallCntOver)
        Case 0:
            LeftBalls=0
            RightBalls=5
        Case 1:
            select case (TGcaptiveR.BallCntOver+TGcaptiveR1.BallCntOver+TGcaptiveR2.BallCntOver)
                Case 2:
                    LeftBalls=2
                    RightBalls=3
                Case 3:
                    LeftBalls=1
                    RightBalls=4
            end Select
        case 2:
            LeftBalls=3
            RightBalls=2
        Case 3:
            select case (TGcaptiveR.BallCntOver)
                Case 0:
                    LeftBalls=5
                    RightBalls=0
                Case 1:
                    LeftBalls=4
                    RightBalls=1
            end Select
    end Select

end sub

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
        If(Tilted = False) And (Credits < 25) Then
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
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound "fx_nudge", 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound "fx_nudge", 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound "fx_nudge", 0, 1, 1, 0.25

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
                    UpdateBallsToPlay
          UpdatePlayers
                Else
                    If(Credits> 0) then
                        PlayersPlayingGame = PlayersPlayingGame + 1
                        Credits = Credits - 1
                        UpdateCredits
                        UpdateBallsToPlay
            UpdatePlayers
            Playsound"BallyStartButtonPlayers2-4 plus 10dB"
                    End If
                End If
            End If
        End If

     Else ' (GameInPlay)

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
              Playsound"BallyStartButtonPlayer1+3dB"
                       End If
                  End If
              End If
         End If

     End If
End Sub

Sub Table1_KeyUp(ByVal keycode)

    If EnteringInitials then
        Exit Sub
    End If

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
FullStrokeEOS_Torque = 0.6  ' EOS Torque when flipper hold up ( EOS Coil is fully charged. Ampere increase due to flipper can't move or when it pushed back when "On". EOS Coil have more power )
LiveStrokeEOS_Torque = 0.3  ' EOS Torque when flipper rotate to end ( When flipper move, EOS coil have less Ampere due to flipper can freely move. EOS Coil have less power )

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
'     Flipper
'*******************

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFXDOF ("LeftflipperupH-2dB", 101, DOFOn, DOFFlippers), LeftFlipper
        LeftFlipper.RotateToEnd
    LeftFlipperOn = 1
    Else
        PlaySoundAt SoundFXDOF ("LeftflipperdownH", 101, DOFOff, DOFFlippers), LeftFlipper
        LeftFlipper.RotateToStart
    LeftFlipperOn = 0
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFXDOF ("RightflipperupH-2dB", 102, DOFOn, DOFFlippers), RightFlipper
        RightFlipper.RotateToEnd
        RightFlipperOn = 1
    Else
        PlaySoundAt SoundFXDOF ("RightflipperdownH", 102, DOFOff, DOFFlippers), RightFlipper
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

Sub LeftFlipper_Animate:LeftFlipperTop.Rotz = LeftFlipper.CurrentAngle:End Sub
Sub RightFlipper_Animate:RightFlipperTop.Rotz = RightFlipper.CurrentAngle:End Sub

'*******************
' GI lights
'*******************

Sub GiOn 'GI lights on
    Dim bulb
  PlaySound "fx_gion"
    For each bulb in aGiLights
        bulb.State = 1
    Next
End Sub

Sub GiOff 'GI lights off
    Dim bulb
    PlaySound "fx_gioff"
    For each bulb in aGiLights
        bulb.State = 0
    Next
End Sub

Dim GiIntensity
GiIntensity = 1 'can be used by the LUT changing to increase the GI lights when the table is darker

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
'     TILT
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
        'BallsRemaining(CurrentPlayer) = 0 'player looses the game, mostly on older 1 player games
        TiltRecoveryTimer.Enabled = True 'wait for all the balls to drain
    End If
End Sub

Sub TiltDecreaseTimer_Timer
    ' DecreaseTilt
    If Tilt> 0 Then
        Tilt = Tilt - 0.1
    Else
        TiltDecreaseTimer.Enabled = False
    End If
End Sub

Sub DisableTable(Enabled)
    If Enabled Then
        GiOff
        'Disable slings, bumpers etc
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
    ' all the balls have drained ..
    If(BallsOnPlayfield = 0) Then
        EndOfBall
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

'****************************************************
'   JP's VPX Rolling Sounds with ball speed control
'****************************************************

Const tnob = 10   'total number of balls
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
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

'***************************************
'     Collection collision sounds
'***************************************

Sub aMetals_Hit(idx):PlaySoundAtBall "fx_MetalHit":End Sub
Sub aMetalWires_Hit(idx):PlaySoundAtBall "fx_MetalWire":End Sub
Sub aRubber_Bands_Hit(idx):PlaySoundAtBall "fx_rubber_band":End Sub
Sub aRubber_Posts_Hit(idx):PlaySoundAtBall "fx_rubber_post":End Sub
Sub aRubber_Pins_Hit(idx):PlaySoundAtBall "fx_rubber_pin":End Sub
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
    UpdateBallsToPlay
  UpdatePlayers
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
    ' reset table for a new ball, rise droptargets, etc
    ResetForNewPlayerBall()
    CreateNewBall()
End Sub

' (Re-)init table for a new ball or player

Sub ResetForNewPlayerBall()
    'debug.print "ResetForNewPlayerBall"
  UpdatePlayers
    AddScore 0

    ' reset multiplier to 1x
    BonusMultiplier = 1

    ' turn on lights, and variables
    bExtraBallWonThisBall = False
    ResetNewBallVariables
    ResetNewBallLights
End Sub

' Create new ball

Sub CreateNewBall()
    BallRelease.CreateSizedBallWithMass BallSize / 2, BallMass
    BallsOnPlayfield = BallsOnPlayfield + 1
    UpdateBallsToPlay
  PlaySoundAt SoundFXDOF ("fx_Ballrel", 104, DOFPulse, DOFContactors), BallRelease
  vpmtimer.addtimer 600, "BallRelease.Kick 90, 4 '"
End Sub

' player lost the ball

Sub EndOfBall
    'debug.print "EndOfBall"

    ' Lost the first ball, now it cannot accept more players
    bOnTheFirstBall = False

  ' Bonus Count
' BonusMultiplier = 2 ' For test
    If NOT Tilted Then
        Select Case BonusMultiplier
            Case 1:BonusCountTimer.Interval = 500
      Case 2:BonusCountTimer.Interval = 1000
    End Select
    CBonus = 0
        BonusCountTimer.Enabled = 1
    Else
    vpmtimer.addtimer 500, "EndOfBall2 '"
  Playsound"MotorLeer2"
  End If
End Sub

Sub BonusCountTimer_Timer 'The bonus are count when the ball is lost
    'debug.print "BonusCount_Timer"
    If Bonus> 0 Then
    If BonusMultiplier = 1 Then
      CBonus = CBonus + 1
      If CBonus = 5 Then  BonusCountTimer.Interval = BonusCountTimer.Interval + 100
      If CBonus = 6 Then BonusCountTimer.Interval = BonusCountTimer.Interval - 100
    End If
        AddScore 3000 * BonusMultiplier
    If BonusMultiplier = 2 Then PlaySoundAt "fx_bonus", Bumper001
        Bonus = Bonus -1
        UpdateBonusLights
    Else
        BonusCountTimer.Enabled = 0
        vpmtimer.addtimer 1000, "EndOfBall2 '"
    Playsound"MotorLeer2"
    End If
End Sub

Sub UpdateBonusLights
    Select Case Bonus
        Case 0:BL3K.State = 0:BL6K.State = 0:BL9K.State = 0:BL12K.State = 0:BL15K.State = 0:BL18K.State = 0:BL21K.State = 0:BL24K.State = 0
        Case 1:BL3K.State = 1:BL6K.State = 0:BL9K.State = 0:BL12K.State = 0:BL15K.State = 0:BL18K.State = 0:BL21K.State = 0:BL24K.State = 0
        Case 2:BL3K.State = 1:BL6K.State = 1:BL9K.State = 0:BL12K.State = 0:BL15K.State = 0:BL18K.State = 0:BL21K.State = 0:BL24K.State = 0
        Case 3:BL3K.State = 1:BL6K.State = 1:BL9K.State = 1:BL12K.State = 0:BL15K.State = 0:BL18K.State = 0:BL21K.State = 0:BL24K.State = 0
        Case 4:BL3K.State = 1:BL6K.State = 1:BL9K.State = 1:BL12K.State = 1:BL15K.State = 0:BL18K.State = 0:BL21K.State = 0:BL24K.State = 0
        Case 5:BL3K.State = 1:BL6K.State = 1:BL9K.State = 1:BL12K.State = 1:BL15K.State = 1:BL18K.State = 0:BL21K.State = 0:BL24K.State = 0
        Case 6:BL3K.State = 1:BL6K.State = 1:BL9K.State = 1:BL12K.State = 1:BL15K.State = 1:BL18K.State = 1:BL21K.State = 0:BL24K.State = 0
        Case 7:BL3K.State = 1:BL6K.State = 1:BL9K.State = 1:BL12K.State = 1:BL15K.State = 1:BL18K.State = 1:BL21K.State = 1:BL24K.State = 0
        Case 8:BL3K.State = 1:BL6K.State = 1:BL9K.State = 1:BL12K.State = 1:BL15K.State = 1:BL18K.State = 1:BL21K.State = 1:BL24K.State = 1
    End Select
  If BL24K.State = 1 Then
    If LightLaneL.State = 1 Then
      LightSpecialR.State = 1
    Else
      LightSpecialL.State = 1
    End If
  End If
End Sub

Sub EndOfBall2
    'debug.print "EndOfBall2"

    Tilted = False
    Tilt = 0
    TiltReel.SetValue 0
    If B2SOn then
        Controller.B2SSetTilt 0
    end if
    DisableTable False

  ' no extra ball

    BallsRemaining(CurrentPlayer) = BallsRemaining(CurrentPlayer) - 1

    ' last ball?
    If(BallsRemaining(CurrentPlayer) <= 0) Then
         CheckHighScore()
    End If

    ' this is not the last ball, check for new player
    EndOfBallComplete()

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
    If((BallsRemaining(CurrentPlayer) <= 0) AND (BallsRemaining(NextPlayer) <= 0) ) Then

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
        'AddScore 0

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

    ' turn off flippers
    SolLFlipper 0
    SolRFlipper 0

    ' start the attract mode
    vpmTimer.AddTimer 3000, "StartAttractMode '"
End Sub

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
'      Match
'******************

Sub Verification_Match()
    PlaySound "MotorLeer2"
    Match = INT(RND(1) * 10) * 10 ' random between 00 and 90
    Display_Match
    If(Score(1) MOD 100) = Match Then
        AwardSpecial
    End If
End Sub

Sub Clear_Match()
    MatchReel1.SetValue 0
    MatchReel2.SetValue 0
    If B2SOn then
        Controller.B2SSetMatch 0
    end if
End Sub

Sub Display_Match()
  If (Match \ 10) < 5 Then  MatchReel1.SetValue(Match \ 10) + 1  Else  MatchReel2.SetValue(Match \ 10) + 1
    If B2SOn then
        If Match = 0 then
            Controller.B2SSetMatch 100
        else
            Controller.B2SSetMatch Match
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
  DOF 250, DOFPulse

    'tilted?
    If Tilted Then
        StopEndOfBallMode
    End If

    ' if still playing and not tilted
    If(bGameInPLay = True) AND (Tilted = False) Then

        ' ballsaver?
        If(bBallSaverActive = True) Then
            CreateNewBall()
        Else
            ' last ball?
            If(BallsOnPlayfield = 0) Then
                StopEndOfBallMode
                vpmtimer.addtimer 500, "EndOfBall '"
                Exit Sub
            End If
        End If
    End If
End Sub

' ***************************************
'               Score functions
' ***************************************

Sub AddScore(Points)
    If Tilted Then Exit Sub
    Select Case Points
        Case 10, 100, 1000 , 10000
      If Score(CurrentPlayer) + Points > 199999 Then Exit Sub
            Score(CurrentPlayer) = Score(CurrentPlayer) + points
            UpdateScore points
        ' sounds
      If Points = 100 AND(Score(CurrentPlayer) MOD 1000) \ 100 = 0 Then            ' new reel 1000
                PlaySound SoundFXDOF ("fx_williamschime1000", 303, DOFPulse, DOFChimes)
            ElseIf Points = 10 AND(Score(CurrentPlayer) MOD 100) \ 10 = 0 Then           ' new reel 100
                PlaySound SoundFXDOF ("fx_williamschime100", 302, DOFPulse, DOFChimes)
      ElseIf Points = 10000 Then
                PlaySound SoundFXDOF ("fx_williamschime1000", 303, DOFPulse, DOFChimes)
      ElseIf Points = 1000 Then
                PlaySound SoundFXDOF ("fx_williamschime1000", 303, DOFPulse, DOFChimes)
      ElseIf Points = 100 Then
                PlaySound SoundFXDOF ("fx_williamschime100", 302, DOFPulse, DOFChimes)
      Else
                PlaySound SoundFXDOF ("fx_williamschime10", 301, DOFPulse, DOFChimes)
            End If
        Case 20, 30, 40, 50, 60, 70, 80, 90
            Add10 = Add10 + Points \ 10
            AddScore10Timer.Enabled = TRUE
        Case 200, 300, 400, 500, 600, 700, 800, 900
            Add100 = Add100 + Points \ 100
            AddScore100Timer.Enabled = TRUE
        Case 2000, 3000, 4000, 5000, 6000, 7000, 8000, 9000
            Add1000 = Add1000 + Points \ 1000
            AddScore1000Timer.Enabled = TRUE
        Case 20000, 30000, 40000, 50000
            Add10000 = Add10000 + Points \ 10000
            AddScore10000Timer.Enabled = TRUE
    End Select

    ' check for higher score and specials for 3 Balls
  If BallsPerGame = 3 Then
    If Score(CurrentPlayer) >= Special1 AND Special1Awarded(CurrentPlayer) = False Then
      AwardSpecial
      Special1Awarded(CurrentPlayer) = True
    End If
    If Score(CurrentPlayer) >= Special2 AND Special2Awarded(CurrentPlayer) = False Then
      AwardSpecial
      Special2Awarded(CurrentPlayer) = True
    End If
  End If
   ' check for higher score and specials for 5 Balls
  If BallsPerGame = 5 Then
    If Score(CurrentPlayer) >= Special3 AND Special3Awarded(CurrentPlayer) = False Then
      AwardSpecial
      Special3Awarded(CurrentPlayer) = True
    End If
    If Score(CurrentPlayer) >= Special4 AND Special4Awarded(CurrentPlayer) = False Then
      AwardSpecial
      Special4Awarded(CurrentPlayer) = True
    End If
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
'     Bonus
'*******************

Sub AddBonus
    If Tilted Then Exit Sub
  If Bonus < 8 Then
    Bonus = Bonus + 1
    DOF 210, DOFPulse
    vpmtimer.addtimer 200, "UpdateBonusLights '"
    End If
End Sub

'***********************************************
'               Score EM reels
'***********************************************

Sub UpdateScore(playerpoints)
  ScoreReel1.Addvalue playerpoints
  If Score(CurrentPlayer) >= 100000 Then
        RolloverReel1.setvalue 1
    End If
    If B2SOn then
        Controller.B2SSetScorePlayer CurrentPlayer, Score(CurrentPlayer)
        If Score(CurrentPlayer) >= 100000 then
            Controller.B2SSetScoreRolloverPlayer1 1
        end if
    end if
End Sub

Sub ResetScores
    ScoreReel1.ResetToZero
  RolloverReel1.setvalue 0
    If B2SOn then
        Controller.B2SSetScorePlayer1 0
        Controller.B2SSetScoreRolloverPlayer1 0
    Controller.B2SSetData 81,0
    end if
End Sub

Sub AddCredits(value) 'limit to 25 credits
    If Credits <25 Then
        Credits = Credits + value
        UpdateCredits
    end if
End Sub

Sub UpdateCredits
    If Credits> 0 Then
        CreditLight.State = 1
    Else
        CreditLight.State = 0
    End If
    PlaySound "fx_relay"
    CreditsReel.SetValue credits
    If B2SOn Then
    Controller.B2SSetCredits Credits
    If Credits = 0 then
            Controller.B2SSetdata 20,0
        else
            Controller.B2SSetdata 20,1
        end if
    end if
End Sub

Sub UpdateBallsToPlay
    Select Case(BallsRemaining(CurrentPlayer))
        Case 1:BallInPlayR.SetValue 1
        Case 2:BallInPlayR.SetValue 2
        Case 3:BallInPlayR.SetValue 3
        Case 4:BallInPlayR.SetValue 4
        Case 5:BallInPlayR.SetValue 5
        Case 6:BallInPlayR.SetValue 6
        Case 7:BallInPlayR.SetValue 7
        Case 8:BallInPlayR.SetValue 8
        Case 9:BallInPlayR.SetValue 9
        Case 10:BallInPlayR.SetValue 10
    End Select
    If B2SOn Then
        Controller.B2SSetBallInPlay BallsRemaining(CurrentPlayer)
    End If
End Sub



Sub UpdatePlayers
    If B2SOn Then
    Controller.B2SSetData 81,1
    End If
End Sub

'*************************
'        Specials
'*************************

Sub AwardExtraBall()
  If bExtraBallWonThisBall = False And BallsRemaining(CurrentPlayer) < 10 Then
    'PlaySound SoundfXDOF ("fx_knocker", 110, DOFPulse, DOFKnocker)
    ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) + 1
    BallsRemaining(CurrentPlayer) = BallsRemaining(CurrentPlayer) + 1
    UpdateBallsToPlay
  End If
  bExtraBallWonThisBall = True
End Sub

Sub AwardSpecial()
    PlaySound SoundfXDOF ("fx_knocker", 110, DOFPulse, DOFKnocker)
    AddCredits 1
End Sub

Sub RocketSpecial
  If bExtraBallWonThisBall = False Then
    AwardExtraBall()
    LightRocketL.State = 0 : LightRocketR.State = 0
    Light100.State = 0 : Light1000.State = 1
  End If
End Sub

' ********************************
'        Attract Mode
' ********************************
' use the"Blink Pattern" of each light

Sub StartAttractMode()
  Dim x
  If AttractMode = 1 Then
    bAttractMode = True
    For each x in aLights
      x.State = 2
    Next
    For each x in aSKYLABLights
      x.State = 2
    Next
  End If
    If B2SOn then
        Controller.B2SSetGameOver 1
        Controller.B2SSetBallInPlay 0
    Controller.B2SSetTilt 0
    Controller.B2SSetData 81,1
    end if
    GameOverR.SetValue 1
    BallInPlayR.SetValue 0
  TiltReel.SetValue 0
End Sub

Sub StopAttractMode()
    Dim x
    bAttractMode = False
    TurnOffPlayfieldLights
  TurnOffSKYLABLights
    ResetScores
    GameOverR.SetValue 0
    If B2SOn then
        Controller.B2SSetGameOver 0
    end if
End Sub

'************************************************
'    Load / Save / Highscore
'************************************************

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

'****************************************
' Realtime updates
'****************************************

Sub Gate1_Animate:Pgate1.rotx = Gate1.currentangle*.6:End Sub
Sub Gate2_Animate:Pgate2.rotx = Gate2.currentangle*.6:End Sub

'***********************************************************************
' *********************************************************************
'  *********     G A M E  C O D E  S T A R T S  H E R E      *********
' *********************************************************************
'***********************************************************************

Sub VPObjects_Init 'init objects
    TurnOffPlayfieldLights()
End Sub

' Dim all the variables

Sub Game_Init() 'called at the start of a new game
    'Start music?
    'Init variables?
    'Start or init timers
    'Init lights?
    TurnOffPlayfieldLights()
  TurnOffSKYLABLights()
  TurnOnSKYLABLights()
End Sub

Sub StopEndOfBallMode() 'called when the last ball is drained

End Sub

Sub ResetNewBallVariables() 'init variables new ball/player
  If BallsRemaining(CurrentPlayer) = 1 Then BonusMultiplier = 2
  Bonus = 1
End Sub

Sub ResetNewBallLights()    'init lights for new ball/player
  TurnOffPlayfieldLights()
  TurnOnNewBallLights()
  UpdateBonusLights()
  If BallsRemaining(CurrentPlayer) = 1 Then LightDoubleBonus.State = 1
  If leftballs > 2 Then
    LightRocketL.State = 0 : LightRocketR.State = 1
  Else
    LightRocketL.State = 1 : LightRocketR.State = 0
  End If
End Sub

Sub TurnOnNewBallLights()
    Dim a
    For each a in aNewBallLights
        a.State = 1
    Next
End Sub

Sub TurnOffPlayfieldLights()
    Dim a
    For each a in aLights
        a.State = 0
    Next
End Sub

Sub TurnOnSKYLABLights()
    Dim a
    For each a in aSKYLABLights
        a.State = 1
    Next
End Sub

Sub TurnOffSKYLABLights()
    Dim a
    For each a in aSKYLABLights
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

'**********************
'       Slingshots
'**********************

Dim LStep, RStep

Sub LeftSlingShot_Slingshot   'left slingshot
    If Tilted Then Exit Sub
    PlaySoundAt SoundFXDOF ("fx_slingshot",103,DOFPulse,DOFcontactors), Lemk
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    LeftSlingShot.TimerEnabled = True
    ' add points
    AddScore 10
    ' some effect

End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 1:LeftSLing4.Visible = 0:LeftSLing3.Visible = 1:Lemk.RotX = 14
        Case 2:LeftSLing3.Visible = 0:LeftSLing2.Visible = 1:Lemk.RotX = 2
        Case 3:LeftSLing2.Visible = 0:Lemk.RotX = -20:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot    'right slingshot
    If Tilted Then Exit Sub
    PlaySoundAt SoundFXDOF ("fx_slingshot",104,DOFPulse,DOFcontactors), Remk
    RightSling4.Visible = 1
    Remk.RotX = 26
    RStep = 0
    RightSlingShot.TimerEnabled = True
    ' add points
    AddScore 10
    ' some effect

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

Dim Rub1, Rub2, Rub3, Rub4, Rub5, Rub6

Sub RubberBand001_Hit  'Low left Y
    If Tilted then Exit Sub
  AddScore 1000
  If LightY.State = 1 Then DOF 208, DOFPulse
  LightY.State = 0
  CheckSkylab()
    Rub1 = 1:RubberBand001_Timer
End Sub

Sub RubberBand001_Timer
    Select Case Rub1
        Case 1:r001.Visible = 0:r007.Visible = 1:RubberBand001.TimerEnabled = 1
        Case 2:r007.Visible = 0:r008.Visible = 1
        Case 3:r008.Visible = 0:r001.Visible = 1:RubberBand001.TimerEnabled = 0
    End Select
    Rub1 = Rub1 + 1
End Sub

Sub RubberBand002_Hit   'Middle left K
    If Tilted then Exit Sub
  AddScore 1000
  If LightK.State = 1 Then DOF 208, DOFPulse
  LightK.State = 0
  CheckSkylab()
    Rub2 = 1:RubberBand002_Timer
End Sub

Sub RubberBand002_Timer
    Select Case Rub2
        Case 1:r002.Visible = 0:r009.Visible = 1:RubberBand002.TimerEnabled = 1
        Case 2:r009.Visible = 0:r010.Visible = 1
        Case 3:r010.Visible = 0:r002.Visible = 1:RubberBand002.TimerEnabled = 0
    End Select
    Rub2 = Rub2 + 1
End Sub

Sub RubberBand003_Hit   'Top left S
    If Tilted then Exit Sub
  AddScore 10
  If LightSS.State = 1 Then DOF 208, DOFPulse
  LightSS.State = 0
  CheckSkylab()
    Rub3 = 1:RubberBand003_Timer
End Sub

Sub RubberBand003_Timer
    Select Case Rub3
        Case 1:r003.Visible = 0:r011.Visible = 1:RubberBand003.TimerEnabled = 1
        Case 2:r011.Visible = 0:r012.Visible = 1
        Case 3:r012.Visible = 0:r003.Visible = 1:RubberBand003.TimerEnabled = 0
    End Select
    Rub3 = Rub3 + 1
End Sub

Sub RubberBand004_Hit   'Top right L
    If Tilted then Exit Sub
  AddScore 10
  If LightL.State = 1 Then DOF 208, DOFPulse
  LightL.State = 0
  CheckSkylab()
    Rub4 = 1:RubberBand004_Timer
End Sub

Sub RubberBand004_Timer
    Select Case Rub4
        Case 1:r004.Visible = 0:r018.Visible = 1:RubberBand004.TimerEnabled = 1
        Case 2:r018.Visible = 0:r019.Visible = 1
        Case 3:r019.Visible = 0:r004.Visible = 1:RubberBand004.TimerEnabled = 0
    End Select
    Rub4 = Rub4 + 1
End Sub

Sub RubberBand005_Hit   'Middle right A
    If Tilted then Exit Sub
  AddScore 1000
  If LightA.State = 1 Then DOF 208, DOFPulse
  LightA.State = 0
  CheckSkylab()
    Rub5 = 1:RubberBand005_Timer
End Sub

Sub RubberBand005_Timer
    Select Case Rub5
        Case 1:r005.Visible = 0:r015.Visible = 1:RubberBand005.TimerEnabled = 1
        Case 2:r015.Visible = 0:r016.Visible = 1
        Case 3:r016.Visible = 0:r005.Visible = 1:RubberBand005.TimerEnabled = 0
    End Select
    Rub5 = Rub5 + 1
End Sub

Sub RubberBand006_Hit   'Low right B
    If Tilted then Exit Sub
  AddScore 1000
  If LightB.State = 1 Then DOF 208, DOFPulse
  LightB.State = 0
  CheckSkylab()
    Rub6 = 1:RubberBand006_Timer
End Sub

Sub RubberBand006_Timer
    Select Case Rub6
        Case 1:r006.Visible = 0:r013.Visible = 1:RubberBand006.TimerEnabled = 1
        Case 2:r013.Visible = 0:r014.Visible = 1
        Case 3:r014.Visible = 0:r006.Visible = 1:RubberBand006.TimerEnabled = 0
    End Select
    Rub6 = Rub6 + 1
End Sub

'*********
' Bumpers
'*********

Sub Bumper001_Hit 'left
    If Tilted Then Exit Sub
    PlaySoundAt SoundFXDOF ("fx_Bumper",105,DOFPulse,DOFContactors), bumper001
    AddScore 100
End Sub

Sub Bumper002_Hit ' right
    If Tilted Then Exit Sub
    PlaySoundAt SoundFXDOF ("fx_Bumper",106,DOFPulse,DOFContactors), bumper002
    AddScore 100
End Sub

'*****************
'     Triggers
'*****************

' out lanes

Sub Trigger001_Hit  'Special Left
    PlaySoundAt "fx_sensor", Trigger001
  DOF 204, DOFPulse
    If Tilted Then Exit Sub
  If LightSpecialL.State = 1 Then
    AwardSpecial()
  Else
    Addscore 1000
  End If
End Sub

Sub Trigger002_Hit  'Lane Left
    PlaySoundAt "fx_sensor", Trigger002
  DOF 205, DOFPulse
    If Tilted Then Exit Sub
  If LightLaneL.State = 1 Then
    Addscore 1000
    vpmtimer.addtimer 200, "AddBonus '"
  Else
    Addscore 100
  End If
End Sub

Sub Trigger003_Hit  'Lane Right
    PlaySoundAt "fx_sensor", Trigger003
  DOF 206, DOFPulse
    If Tilted Then Exit Sub
  If LightLaneR.State = 1 Then
    Addscore 1000
    vpmtimer.addtimer 200, "AddBonus '"
  Else
    Addscore 100
  End If
End Sub

Sub Trigger004_Hit  'Special Right
    PlaySoundAt "fx_sensor", Trigger004
  DOF 207, DOFPulse
    If Tilted Then Exit Sub
  If LightSpecialR.State = 1 Then
    AwardSpecial()
  Else
    Addscore 1000
  End If
End Sub

' top lanes

Sub Trigger005_Hit 'Left
    PlaySoundAt "fx_sensor", Trigger005
  DOF 202, DOFPulse
    If Tilted Then Exit Sub
  If LightTopLeft.State = 1 Then
    Addscore 1000
    vpmtimer.addtimer 200, "AddBonus '"
  Else
    Addscore 100
  End If
End Sub

Sub Trigger006_Hit  'Right
    PlaySoundAt "fx_sensor", Trigger006
  DOF 203, DOFPulse
    If Tilted Then Exit Sub
  If LightTopRight.State = 1 Then
    Addscore 1000
    vpmtimer.addtimer 200, "AddBonus '"
  Else
    Addscore 100
  End If
End Sub

' horseshoe

Sub TGcaptiveL3_hit 'left
    If Not tilt Then
        If TGcaptiveR3.timerenabled Then
      PlaySoundAt "fx_sensor", TGcaptiveL3
      DOF 201, DOFPulse
      If Light100.State = 1 Then AddScore 100 Else AddScore 1000
      If LightRocketL.state=1 And (TGcaptiveL.BallCntOver+TGcaptiveL1.BallCntOver+TGcaptiveL2.BallCntOver = 3) Then
        If TGcaptiveR.BallCntOver = 0 Then RocketSpecial()
      End If
    End If
        me.timerenabled=1
    End If
End Sub

Sub TGcaptiveL3_timer
    me.timerenabled=0
End Sub

Sub TGcaptiveR3_hit 'right
    If Not tilt Then
        If TGcaptiveL3.timerenabled Then
      PlaySoundAt "fx_sensor", TGcaptiveR3
      DOF 201, DOFPulse
      If Light100.State = 1 Then AddScore 100 Else AddScore 1000
      If LightRocketR.state=1 And (TGcaptiveR.BallCntOver+TGcaptiveR1.BallCntOver+TGcaptiveR2.BallCntOver = 3) Then
        If TGcaptiveL.BallCntOver = 0 Then RocketSpecial()
      End If
    End If
        me.timerenabled=1
    End If
End Sub

Sub TGcaptiveR3_timer
    me.timerenabled=0
End Sub

'**********************************
'         Rollover Button
'**********************************

Sub sw001_Hit 'left
    PlaySoundAt "fx_sensor", sw001
    If Tilted Then Exit Sub
  Change()
End Sub

Sub sw002_Hit 'Right
    PlaySoundAt "fx_sensor", sw002
    If Tilted Then Exit Sub
  Change()
End Sub

'******************************
'      Extra routines
'******************************

Sub CheckSkylab()
  If LightSS.State=0 And LightK.State=0 And LightY.State=0 And LightL.State=0 And LightA.State=0 And LightB.State=0 Then
    vpmtimer.addtimer 200, "AddBonus '"
    vpmtimer.addtimer 200, "AddBonus '"
  LightSS.State=1 : LightK.State=1 : LightY.State=1 : LightL.State=1 : LightA.State=1 : LightB.State=1
  End If
End Sub

Sub Change()
  If LightTopLeft.State = 1 Then
    LightTopLeft.State = 0 : LightTopRight.State = 1
    LightLaneL.State = 1 : LightLaneR.State = 0
  Else
    LightTopLeft.State = 1 : LightTopRight.State = 0
    LightLaneL.State = 0 : LightLaneR.State = 1
  End If
  If BL24K.State = 1 Then
    If LightSpecialL.State = 1 Then
      LightSpecialL.State = 0 : LightSpecialR.State = 1
    Else
      LightSpecialL.State = 1 : LightSpecialR.State = 0
    End If
  End If
End Sub
'
' ============================================================================================
' GNMOD - Multiple High Score Display and Collection
' jpsalas: changed ramps by flashers
' ============================================================================================

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
HSScore(1) = 40000
HSScore(2) = 35000
HSScore(3) = 30000
HSScore(4) = 25000
HSScore(5) = 20000

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
        PlaySound "sfx_Enter"
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
        PlaySound "sfx_Previous"
    ElseIf keycode = RightFlipperKey Then
        ' advance to Next character
        AlphaStringPos = AlphaStringPos + 1
        If AlphaStringPos> len(AlphaString) or(AlphaStringPos = len(AlphaString) and InitialString = "") then
            ' Skip the backspace If there are no characters to backspace over
            AlphaStringPos = 1
        End If
        SetHSLine 3, InitialString & MID(AlphaString, AlphaStringPos, 1)
        PlaySound "sfx_Next"
    ElseIf keycode = StartGameKey or keycode = PlungerKey Then
        SelectedChar = MID(AlphaString, AlphaStringPos, 1)
        If SelectedChar = "_" then
            InitialString = InitialString & " "
            PlaySound("sfx_Esc")
        ElseIf SelectedChar = "<" then
            InitialString = MID(InitialString, 1, len(InitialString) - 1)
            If len(InitialString) = 0 then
                ' If there are no more characters to back over, don't leave the < displayed
                AlphaStringPos = 1
            End If
            PlaySound("sfx_Esc")
        Else
            InitialString = InitialString & SelectedChar
            PlaySound("sfx_Enter")
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

'************************************************
'  User Options (Example script from VPINWORKSHOP)
'************************************************

Dim LUTImage    ' LUT - Darkness control 10 normal level & 10 warmer levels
Dim BallsPerGame  ' 0 = play 3 balls ; 1 = play 5 balls
Dim AttractMode   ' 0 = no light attract mode ; 1 = light attract mode
Dim FreePlay      ' 0 = coins ; 1 = Free play

' Called when options are tweaked by the player.
' - 0: game has started, good time to load options and adjust accordingly
' - 1: an option has changed
' - 2: options have been reseted
' - 3: player closed the tweak UI, good time to update staticly prerendered parts
' Table1.Option arguments are:
' - option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional array of literal strings

Dim dspTriggered : dspTriggered = False
Sub Table1_OptionEvent(ByVal eventId)
    If eventId = 1 And Not dspTriggered Then dspTriggered = True : DisableStaticPreRendering = True : End If
  dim v

  ' LUT - Darkness control
    LUTImage = Table1.Option("Darkness control", 0, 21, 1, 0, 0, _
    Array("LUT0", "LUT1", "LUT2", "LUT3", "LUT4", "LUT5", "LUT6", "LUT7", "LUT8", "LUT9", "LUT10", _
        "LUT Warm0", "LUT Warm1", "LUT Warm2", "LUT Warm3", "LUT Warm4", "LUT Warm5", "LUT Warm6", "LUT Warm7", "LUT Warm8", "LUT Warm9", "LUT Warm10"))
  ' Balls Per Game
    BallsPerGame = Table1.Option("Balls Per Game", 0, 1, 1, 1, 0, Array("3 Balls","5 Balls"))
  ' Light Attract Mode
    AttractMode = Table1.Option("Attract Mode", 0, 1, 1, 1, 0, Array("No","Yes"))
  ' Free Play
    Freeplay = Table1.Option("Free Play", 0, 1, 1, 0, 0, Array("No","Yes"))

  UpdateOptions

    If eventId = 3 And dspTriggered Then dspTriggered = False : DisableStaticPreRendering = False : End If
End Sub

Sub UpdateOptions

  '*************************
  'Balls Per Game
  '*************************
  ' select Card Instruction on Apron
  If BallsPerGame = 0 Then
    BallsPerGame = 3
    Flasher3balls.visible = 1
    Flasher5balls.visible = 0
  Else
    BallsPerGame = 5
    Flasher3balls.visible = 0
    Flasher5balls.visible = 1
  End If

  '*************************
  'Light Attract Mode
  '*************************
  ' start Attract Mode
  Dim x
  If AttractMode = 1 Then
    StartAttractMode()
  Else
    bAttractMode = False
    TurnOffPlayfieldLights
  End If

  '*************************
  'Freeplay
  '*************************
  ' select Free Play
    If Freeplay = 0 Then bFreePlay = False else bFreePlay = True

  '*************************
  'LUT - Darkness control
  '*************************
  ' select LUT Image
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

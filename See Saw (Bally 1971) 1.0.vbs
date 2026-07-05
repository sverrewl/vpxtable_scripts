' ***********************************************************************
'               VISUAL PINBALL X EM Script
'                  Basic EM script, 1 player only
'             uses core.vbs for extra functions
'
'  A decent approximation of Bally's See Saw (1970).  Scripted by XFL.
'  IPDB #654
'  some info on real table: https://www.ipdb.org/machine.cgi?id=2091
'
' I won't guarantee that I got all of the rules right.  I watched a few videos and read a few things online.
'
' Start of table was Expressway. I also used the top of the gate primitive from the "little joe" table.  I modified some of the Highscore logic from the Space Time table.
'
'  Thanks again to the Visual Pinball community - for making the software, making tutorials, and letting people reuse objects and code.
'
' For what its worth, you have my permission to use or modify anything here.
' I don't need any credit and don't blame me if it doesn't work.
'
'
'
'  2026-06-05  Version 0.0  (initial release - likely full of bugs)
'
' ***********************************************************************



Option Explicit
Randomize

'* * * * * * * * * * * * * * * * * * * * *

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


'* * * * * * * * * * * * * * * * * * * * *

' Constants
Const TableName = "SeeSaw" ' file name to save config
Const cGameName = "SeeSaw" ' B2S name

' see below Const BallsPerGame = 5
Const MaxBallsOnPlayfield = 3
Const TiltSensitivity = 6
Const MaxNumberOfPlayers = 4
Const isFlashingStuff = False

Const PointsToEarnSpecial1 = 40000
Const PointsToEarnSpecial2 = 48000
Const PointsToEarnSpecial3 = 57000
Const PointsToEarnSpecial4 = 66000

' Global variables

Dim Credits
Dim BallsRemaining
Dim BallInPlay
Dim PlayerScore(4)
Dim isFirstBallDone(4)
Dim NumberOfPlayers
Dim CurrentPlayer
Dim TiltMeasure

' Control variables
Dim BallsOnPlayfield

Dim HighScoreAllTime
Dim HighScoreThisGame

' Boolean variables
Dim isGameInPlay
Dim isTilted
Dim isClearbonus



' *********************************************************************
'  Player Input
' *********************************************************************

Sub Table1_KeyDown(ByVal Keycode)

    If HSEnterMode Then
        HighScoreProcessKey(keycode)
        Exit Sub
    End If

    If keycode = PlungerKey Then
        Plunger.Pullback
        PlaySoundAt "fx_plungerpull", plunger
    End If

' tilt/nudge
    If keycode = LeftTiltKey Then Nudge 90, 8:PlaySound "fx_nudge", 0, 1, -0.1, 0.25:CheckTilt
    If keycode = RightTiltKey Then Nudge 270, 8:PlaySound "fx_nudge", 0, 1, 0.1, 0.25:CheckTilt
    If keycode = CenterTiltKey Then Nudge 0, 9:PlaySound "fx_nudge", 0, 1, 1, 0.25:CheckTilt

    If isGameInPlay Then
        If NOT isTilted Then
      If keycode = LeftFlipperKey  Then FlippersUpLeft
      If keycode = RightFlipperKey Then FlippersUpRight
        End If
    End If

    If isGameInPlay = False OR BallInPlay <= 1 Then
        If Keycode = AddCreditKey Then
            PlaySound "fx_coin"
            AddCredit()
        End If
        If keycode = StartGameKey And NumberOfPlayers < MaxNumberOfPlayers Then
            If Credits > 0 Then
                Credits = Credits - 1
                NumberOfPlayers = NumberOfPlayers + 1
                If NumberOfPlayers = 2 Then ReelScore2.setvalue 1111:ReelScore2.ResetToZero
                If NumberOfPlayers = 3 Then ReelScore3.setvalue 1111:ReelScore3.ResetToZero
                If NumberOfPlayers = 4 Then ReelScore4.setvalue 1111:ReelScore4.ResetToZero
                If NOT isGameInPlay Then
                    Game_Init()
                Else
                    UpdateBackglass()
                End If
            Else
                PlaySound "fx_diverter"
            End If
        End If
    End If

End Sub


Sub Table1_KeyUp(ByVal keycode)

    'plunger release
    If keycode = PlungerKey Then
        Plunger.Fire
        If TriggerPlungerLaneBall.BallCntOver() = 0 Then
            PlaySoundAt "fx_plunger_empty", plunger
        Else
            PlaySoundAt "fx_plunger", plunger
        End If
    End If

    If isGameInPlay Then
        If keycode = LeftFlipperKey Then FlippersDownLeft
        If keycode = RightFlipperKey Then FlippersDownRight
    End If

End Sub

Sub AddCredit()
    If Credits < 20 Then Credits = Credits + 1
    UpdateBackglass()
End Sub


'******************
' Table stop/pause
'******************

Sub table1_Paused
End Sub

Sub table1_unPaused
End Sub

' *********************************************************************
' * * * * BEGIN Scripted table objects * * * *
' *********************************************************************

'*******************
'  Flipper Subs
'*******************

Sub FlippersUpLeft()
    If isTilted Then Exit Sub
    If Not isGameInPlay Then Exit Sub
    PlaySoundAt "FX_flipperup", LeftFlipper
    PlayRepeatingSoundAt "buzzL", LeftFlipper 'we use both buzzL and buzzR so that "stopsound" knows which to stop if both flippers are up and one goes down
    If TimerFlipperTricks.Enabled = False Then LeftFlipper.EOSTorque = 0.65
    LeftFlipper.RotateToEnd
    LeftFlipperOn = 1
End Sub

Sub FlippersDownLeft()
    StopSound "buzzL"
    PlaySoundAt "FX_flipperdown", LeftFlipper
    If TimerFlipperTricks.Enabled = False Then LeftFlipper.EOSTorque = 0.15
    LeftFlipper.RotateToStart
    LeftFlipperOn = 0
End Sub

Sub LeftFlipper_Animate
    FlipperShadowL.RotZ = LeftFlipper.currentangle
End Sub

Sub FlippersUpRight()
    If isTilted Then Exit Sub
    If Not isGameInPlay Then Exit Sub
    PlaySoundAt "FX_flipperup", RightFlipper
    PlayRepeatingSoundAt "buzzR", RightFlipper
    If TimerFlipperTricks.Enabled = False Then RightFlipper.EOSTorque = 0.65
    RightFlipper.RotateToEnd
    RightFlipperOn = 1
End Sub

Sub FlippersDownRight()
    StopSound "buzzR"
    PlaySoundAt "FX_flipperdown", RightFlipper
    If TimerFlipperTricks.Enabled = False Then RightFlipper.EOSTorque = 0.15
    RightFlipper.RotateToStart
    RightFlipperOn = 0
End Sub

Sub RightFlipper_Animate
    FlipperShadowR.RotZ = RightFlipper.currentangle
End Sub

'*********************************************************
' <><> BEGIN Real Time Flipper adjustments - by JLouLouLou & JPSalas<><>
'        (to enable flipper tricks)
'*********************************************************
'  Note: this logic is triggered by the timer, disabling the timer turns off the logic

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

TimerFlipperTricks.Interval = 1
'TimerFlipperTricks.TimerEnabled = 1 - see F12 options

Sub TimerFlipperTricks_Timer 'flipper's tricks timer

'Start Of Stroke Flipper Stroke Routine : Start of Stroke for Tap pass and Tap shoot
    If LeftFlipper.CurrentAngle >= LeftFlipper.StartAngle - SOSAngle Then
        LeftFlipper.Strength = FlipperPower * SOSTorque
    Else
        LeftFlipper.Strength = FlipperPower
    End If

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
    If RightFlipper.CurrentAngle <= RightFlipper.StartAngle + SOSAngle Then
        RightFlipper.Strength = FlipperPower * SOSTorque
    Else
        RightFlipper.Strength = FlipperPower
    End If

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

'*********************************************************
' <><> END Real Time Flipper adjustments - by JLouLouLou & JPSalas <><>
'*********************************************************

'************************
'       Targets
'************************

'************************
' Rollover Targets
'************************

'*** Wires

Sub TriggerAdvanceYellowBonus_Hit
    If isTilted Then Exit Sub
    AdvanceYellowBonus()
    AddScore 100
End Sub
Sub TriggerOpenTopGate1_Hit
    If isTilted Then Exit Sub
    OpenTopGate()
    AddScore 1000
End Sub
Sub TriggerAdvanceBlueBonus_Hit
    If isTilted Then Exit Sub
    AdvanceBlueBonus()
    AddScore 100
End Sub
Sub TriggerOpenTopGate2_Hit
    If isTilted Then Exit Sub
    OpenTopGate()
    AddScore 1000
End Sub

Sub TriggerRolloverShooterLane_Hit
debug.print "shooter lane rollover walldropped?=" & InvisibleTopGateBlockerWall.IsDropped
    If InvisibleTopGateBlockerWall.IsDropped=True Then
        AddScore 3000
        CloseTopGate()
    End If
End Sub

Sub TriggerShooterLaneBallRest_Hit
    CloseBottomGate()
End Sub

Sub TriggerRolloverRightOutlane_Hit
    Addscore 1000
End Sub

Sub TriggerRolloverLeftOutlane_Hit
debug.print PLKickerOn.State
  If PLKickerOn.State=1 Then
        LeftAutoPlunger.TimerInterval = 250
        LeftAutoPlunger.TimerEnabled = True
    Else
        Addscore 1000
        InvisibleApRollBallWall.IsDropped=False
        InvisibleApHoldBallWall.IsDropped=True  'just in case so we don't get stuck
    End If
End Sub

Sub TriggerInvisibleLeftOutlane_Hit
' if the autoplunger misses, we want to keep trying otherwise the ball will be stuck
' so, only turn off the autoplunger trigger when the ball is no longer on the TriggerRolloverLeftOutlane
    LeftAutoPlunger.TimerEnabled = False
End Sub

'*** Buttons
Sub TriggerButtonLiteRed_Hit
    If isTilted Then Exit Sub
    LiteRedBumpers()
End Sub
Sub TriggerButtonLiteGreen_Hit
    If isTilted Then Exit Sub
    LiteGreenBumpers()
End Sub

Sub TriggerButtonAdvanceBlue_Hit
    AdvanceBlueBonus()
    Addscore 100
End Sub
Sub TriggerButtonAdvanceYellow_Hit
    AdvanceYellowBonus()
    Addscore 100
End Sub

Sub TriggerButtonUpPostCenter_Hit
    If isTilted Then Exit Sub
    Postup()
End Sub
Sub TriggerButtonDownPostLeft_Hit
    If isTilted Then Exit Sub
    Postdown()
End Sub
Sub TriggerButtonDownPostRight_Hit
    If isTilted Then Exit Sub
    Postdown()
End Sub

'************************
' Standup Hit Targets
'************************

Sub StandupTargetAdvanceYellow_Hit
    If isTilted Then Exit Sub
    AdvanceYellowBonus()
    AddScore 100
End Sub

Sub StandupTargetOpenBottomGate_Hit
    If isTilted Then Exit Sub
    OpenBottomGate()
    AddScore 1000
End Sub

Sub StandupTargetAdvanceBlue_Hit
    If isTilted Then Exit Sub
    AdvanceBlueBonus()
    AddScore 100
End Sub

Sub StandupTargetExtraBall_Hit
    If isTilted Then Exit Sub
    If PLextraBall.State=1 Then
        AwardExtraBall()
    Else
        AddScore 100
    End If
End Sub

Sub StandupTargetYellow10X_Hit
    If isTilted Then Exit Sub
    PLyellowX10.State=1
    AddScore 100
End Sub

Sub StandupTargetBlue10X_Hit
    If isTilted Then Exit Sub
    PLblueX10.State=1
    AddScore 100
End Sub

'************************
' Mushroom Bumper Targets
'************************

Sub TriggerMushYellow_Hit
    If isTilted Then Exit Sub
    AddScore 100
    PlaySoundAt "fx_mini bumper", Me
    AdvanceYellowBonus()
    MushYellowAnimationStep = 0
    Me.TimerEnabled = True
End Sub

Dim MushYellowAnimationStep
Sub TriggerMushYellow_timer()
    Me.TimerInterval=10
    MushYellowAnimationStep = MushYellowAnimationStep + 1
    Select Case MushYellowAnimationStep
        case 1: MushCapYellow.TransZ = 2
        case 2: MushCapYellow.TransZ = 4
        case 3: MushCapYellow.TransZ = 6
        case 4: MushCapYellow.TransZ = 4
        case 5: MushCapYellow.TransZ = 2
        case 6: MushCapYellow.TransZ = 0
        case Else: MushYellowAnimationStep=0:Me.TimerEnabled=False
    End Select
End Sub

Sub TriggerMushBlue_Hit
    If isTilted Then Exit Sub
    AddScore 100
    PlaySoundAt "fx_mini bumper", Me
    AdvanceBlueBonus()
    MushBlueAnimationStep = 0
    Me.TimerEnabled = True
End Sub

Dim MushBlueAnimationStep
Sub TriggerMushBlue_timer()
    Me.TimerInterval=10
    MushBlueAnimationStep = MushBlueAnimationStep + 1
    Select Case MushBlueAnimationStep
        case 1: MushCapBlue.TransZ = 2
        case 2: MushCapBlue.TransZ = 4
        case 3: MushCapBlue.TransZ = 6
        case 4: MushCapBlue.TransZ = 4
        case 5: MushCapBlue.TransZ = 2
        case 6: MushCapBlue.TransZ = 0
        case Else: MushBlueAnimationStep=0:Me.TimerEnabled=False
    End Select
End Sub

'************************
'       Drop Targets
'************************


'************************
'       Spinners
'************************


'************************
'       Kickers
'************************

Sub KickerLeft_Hit
    PlaySoundAt "fx_kicker_enter", Me
    If isTilted Then
        Me.TimerEnabled=True
        Exit Sub
    End If
    TimerAwardYellowBonus.Enabled=True
End Sub

Sub KickerLeft_Timer
    If Me.BallCntOver()=1 Then
        Me.Kick 150, 10
        PlaySoundAt "fx_kicker", Me
    End If
    Me.TimerEnabled=False
End Sub

Sub KickerRight_Hit
    PlaySoundAt "fx_kicker_enter", Me
    If isTilted Then
        Me.TimerEnabled=True
        Exit Sub
    End If
    TimerAwardBlueBonus.Enabled=True
End Sub

Sub KickerRight_Timer
    If Me.BallCntOver()=1 Then
        Me.Kick 210, 10
        PlaySoundAt "fx_kicker", Me
    End If
    Me.TimerEnabled=False
End Sub


'************************
' Scoring Walls are simulating switches behind rubber bands
'************************

Sub ScoringWallTopLeft_Hit
    AddScore 10
End Sub
Sub ScoringWallTopRight_Hit
    AddScore 10
End Sub
Sub ScoringWallCenterLeft_Hit
    AddScore 10
End Sub
Sub ScoringWallCenterRight_Hit
    AddScore 10
End Sub
Sub ScoringWallMidLeft_Hit
    AddScore 10
End Sub
Sub ScoringWallMidRight_Hit
    AddScore 10
End Sub


'*********
' Pop Bumpers (thumper-bumpers)
'*********

'BumperGreenTop     'right  top
'BumperGreenBottom  'left bottom
'BumperRedTop       'left top
'BumperRedBottom    'right  bottom
Sub BumperGreenTop_Hit
    If isTilted Then Exit Sub
    If BLGreenTop1.State=1 Then
        Addscore 100
    Else
        AddScore 10
    End If
    PlaySoundAt "fx_Bumper", Me
    If isFlashingStuff Then
        FlashBumperGreenTopHigh.duration 1, 100, 0
        FlashBumperGreenTopLow.duration 1, 100, 0
    End If
End Sub
Sub BumperGreenBottom_Hit
    If isTilted Then Exit Sub
    If BLGreenBottom1.State=1 Then
        Addscore 100
    Else
        AddScore 10
    End If
    PlaySoundAt "fx_Bumper", Me
    If isFlashingStuff Then
        FlashBumperGreenBottomHigh.duration 1, 100, 0
        FlashBumperGreenBottomLow.duration 1, 100, 0
    End If
End Sub
Sub BumperRedTop_Hit
    If isTilted Then Exit Sub
    If BLRedTop1.State=1 Then
        Addscore 100
    Else
        AddScore 10
    End If
    PlaySoundAt "fx_Bumper", Me
    If isFlashingStuff Then
        FlashBumperRedTopHigh.duration 1, 100, 0
        FlashBumperRedTopLow.duration 1, 100, 0
    End If
End Sub
Sub BumperRedBottom_Hit
    If isTilted Then Exit Sub
    If BLRedBottom1.State=1 Then
        Addscore 100
    Else
        AddScore 10
    End If
    PlaySoundAt "fx_Bumper", Me
    If isFlashingStuff Then
        FlashBumperRedBottomHigh.duration 1, 100, 0
        FlashBumperRedBottomLow.duration 1, 100, 0
    End If
End Sub


' *********************************************************************
'  Drain
' *********************************************************************

Sub Drain_Hit()
    Drain.DestroyBall
    PlaySoundAt "fx_drain", Drain

    'there is no active game
    If NOT isGameInPlay Then Exit Sub

    If BallsOnPlayfield > 0 Then
        BallsOnPlayfield = BallsOnPlayfield - 1
    End If

    Ball_End()
End Sub

'**********************
'       Slingshots
'**********************
' note: The real machine probably doesn't have flashers on the slingshots.

Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    If isTilted Then Exit Sub
    PlaySoundAt "fx_slingshot", Lemk
    LStep = 0
    LeftSlingShot.TimerEnabled = True
    If PLAdvanceYellowBonus.State=1 Then
        AddScore 100
        AdvanceYellowBonus()
    Else
        AddScore 10
    End If
    If isFlashingStuff Then FlashSlingLeft.duration 1, 100, 0
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 0:LeftSling004.Visible = 1:Lemk.RotX = 26
        Case 1:LeftSLing004.Visible = 0:LeftSLing003.Visible = 1:Lemk.RotX = 14
        Case 2:LeftSLing003.Visible = 0:LeftSLing002.Visible = 1:Lemk.RotX = 2
        Case 3:LeftSLing002.Visible = 0:Lemk.RotX = -10:LeftSlingShot.TimerEnabled = False
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
    If isTilted Then Exit Sub
    PlaySoundAt "fx_slingshot", Remk
    RStep = 0
    RightSlingShot.TimerEnabled = True
    If PLAdvanceBlueBonus.State=1 Then
        AddScore 100
        AdvanceBlueBonus()
    Else
        AddScore 10
    End If
    If isFlashingStuff Then FlashSlingRight.duration 1, 100, 0
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 0:RightSling004.Visible = 1:Remk.RotX = 26
        Case 1:RightSLing004.Visible = 0:RightSLing003.Visible = 1:Remk.RotX = 14
        Case 2:RightSLing003.Visible = 0:RightSLing002.Visible = 1:Remk.RotX = 2
        Case 3:RightSLing002.Visible = 0:Remk.RotX = -10:RightSlingShot.TimerEnabled = False
    End Select
    RStep = RStep + 1
End Sub


'***********************
'       Gates
'***********************

Sub InvisibleGateRight_Animate
    PGateTopRight.Rotz = Me.CurrentAngle
End Sub

'***********************
'       Rubbers
'***********************

'***********************
'       AutoPlunger
'***********************
Sub LeftAutoPlunger_Timer
    me.TimerInterval=500
    PlaySoundAt "fx_Slingshot", me
    LeftAutoPlunger.fire
    'leave on in case it misses the ball   Me.TimerEnabled=False
End Sub

'' END Scripted table objects

' *********************************************************************
' * * * * BEGIN game step logic * * * *
' *********************************************************************

'* * * * * * * *
'  table_init = turn on machine
'    game_init = start button - start new game
'      turn_init = start a new turn (the next ball)
'        player_init = next player's turn
'            ball_init = create a new ball in the plunger lane
'            ball_end = destroy a ball
'        player_end = end of player's turn
'      turn_end = end a turn for all players
'    game_end = game over
'  table_exit = turn off machine
'* * * * * * * *


'******************
' Table power on
'******************
Sub Table1_Init()

    LoadEM

    ' Remove desktop background items in FS mode
    Dim item

    For each item in cBackdropItems
        If Table1.ShowDT then
            item.Visible = 1
        Else
            item.Visible = 0
        End If
    Next

    isGameInPlay = False

    UpdateBackglass()

    GiOff()
    TurnOffPlayfieldLights()
    TurnOffBumperLights()
    Postdown()

    ' start the attract mode
    StartAttractMode()

    'wait a bit so that the B2S is loaded first
    vpmtimer.addtimer 1000, "LoadConfig() '"


End Sub


'****************************************
' Init for a new game - start button pressed
'****************************************
Sub Game_Init()
'debug.print "game_init"

    isGameInPlay = True

    UnTiltEnableTable()
    StopAttractMode()
    GiOn()
    PlaySound "StartUpSequence"
    GameInitResetLights()
    GameInitResetVariables()
    GameInitResetPlayfield()

    ' first ball.  delay just a bit to let the score reset sounds finish
    vpmtimer.addtimer 1000, "Turn_Init '"

End Sub

Sub GameInitResetLights
    TurnOffPlayfieldLights()
End Sub

Sub GameInitResetVariables
    PlayerScore(1) = 0
    PlayerScore(2) = 0
    PlayerScore(3) = 0
    PlayerScore(4) = 0
    ReelScore1.setvalue 1111
    ReelScore1.ResetToZero
    UpdateBackglass()
    BallsRemaining = BallsPerGame
    SpecialAwarded=0
End Sub

Sub GameInitResetPlayfield
    CloseTopGate()
    CloseBottomGate()
    PostDown()
End Sub

'****************************************
'  start a new turn (the next ball)
'****************************************
Sub Turn_Init()
debug.print "turn_init"
    TiltMeasure = 0

    TurnInitResetLights
    TurnInitResetVariables
    TurnInitResetPlayfield

    UpdateBackglass()
    vpmtimer.addtimer 250, "Player_Init '"
End Sub

Sub TurnInitResetLights() 'init lights for new ball/player
End Sub

Sub TurnInitResetVariables() 'init variables new ball/player
    CurrentPlayer = 1 'each turn starts with player1
End Sub

Sub TurnInitResetPlayfield()
End Sub

'****************************************
' next player's turn
'****************************************

Sub Player_Init
debug.print "player_init"
    UnTiltEnableTable()
    PlayerInitResetLights
    PlayerInitResetVariables
    PlayerInitResetPlayfield

    UpdateBackglass()

    'using the timer to pause a bit before adding the ball when the turn starts.
    vpmtimer.addtimer 1000, "Ball_Init() '"
End Sub

Sub PlayerInitResetLights() 'init lights for new ball/player
    CheckTenCounterLights()
    PLYellow100.State=1
    PLBlue100.State=1
End Sub

Sub PlayerInitResetVariables() 'init variables new ball/player
End Sub

Sub PlayerInitResetPlayfield()
End Sub

'****************************************
'  Create a new ball in the plunger lane
'****************************************
Sub Ball_Init()
 debug.print "ball_init"
    If BallsOnPlayfield>=MaxBallsOnPlayfield Then Exit Sub

    BallRelease.CreateSizedBallWithMass BallSize / 2, BallMass

    PlaySoundAt "fx_Ballrel", BallRelease
    BallRelease.Kick 90, 4

    BallsOnPlayfield = BallsOnPlayfield + 1
End Sub

Sub BallInitResetLights()
    UpdateBackglass()
End Sub

Sub BallInitResetVariables()
End Sub

Sub BallInitResetPlayfield()
End Sub

'****************************************
'  A ball is no longer in play, but there may still be other balls on the playfield
'****************************************
Dim BallEndStep
Sub Ball_End()
    'debug.print "Ball_End"

    PostDown()
    CloseTopGate()
    CloseBottomGate()
    TurnOffBumperLights()
    BallEndStep=0
    TimerBallEndWait.Enabled=True

End Sub

Sub TimerBallEndWait_Timer()

    If BallEndStep = 0 Then
        isClearbonus=True
        TimerAwardYellowBonus.Enabled=True
        BallEndStep = 1
        Exit Sub
    End If

'wait for yellow bonus clear to finish, then start blue bonus clear
    If BallEndStep = 1 Then
        If TimerAwardYellowBonus.Enabled=True Then Exit Sub
        isClearbonus=True
        TimerAwardBlueBonus.Enabled=True
        BallEndStep = 2
        Exit Sub
    End If

'wait for blue bonus clear to finish, then check shoot again
    If BallEndStep = 2 Then
        If TimerAwardBlueBonus.Enabled=True Then Exit Sub
        isClearbonus=False
        BallEndStep = 3
        Exit Sub
    End If

    'If BallEndStep = 3 Then

    Me.Enabled = False

    If BallsOnPlayfield = 0 Then
        If PLShootAgain.state = 1 Then
            PLShootAgain.state = 0
            UpdateBackglass()
            vpmtimer.addtimer 0010, "Ball_Init() '"
        Else
            vpmtimer.addtimer 0010, "Player_End() '"
        End If
    End If

End Sub



'****************************************
'end the player's turn (no balls left on playfield)
'****************************************
Sub Player_End()

'there is no player end of turn bonus count for this machine

    CurrentPlayer = CurrentPlayer + 1
    If CurrentPlayer > NumberOfPlayers Then
        CurrentPlayer = 1
        vpmtimer.addtimer 0010, "Turn_End() '"
    Else
        vpmtimer.addtimer 0010, "Player_Init() '"
    End If

End Sub


'****************************************
'end the turn (all players have played)
'****************************************
Sub Turn_End()

    BallsRemaining = BallsRemaining - 1

    If BallsRemaining > 0 Then
        vpmtimer.addtimer 0100, "Turn_Init '"
    Else
        vpmtimer.addtimer 0100, "Game_End '"
    End If

End Sub

'****************************************
'end the game
'****************************************
Sub Game_End()

    isGameInPlay = False
    UpdateBackglass()

    Check_For_Match()

    FlippersDownLeft
    FlippersDownRight

    'turn off gi Lights
    GiOff()

    TurnOffPlayfieldLights()
    TurnOffBumperLights()
    LightPlayer1.state=0
    LightPlayer2.state=0
    LightPlayer3.state=0
    LightPlayer4.state=0
    NumberOfPlayers=0
    ' start the attract mode
    vpmtimer.addtimer 1000, "StartAttractMode '"

    HighScoreThisGame=PlayerScore(1)
  If PlayerScore(2) > HighScoreThisGame Then HighScoreThisGame=PlayerScore(2)
    If PlayerScore(3) > HighScoreThisGame Then HighScoreThisGame=PlayerScore(3)
    If PlayerScore(4) > HighScoreThisGame Then HighScoreThisGame=PlayerScore(4)
    If HighScoreThisGame > HighScoreAllTime Then
        HighScoreAllTime = HighScoreThisGame
        If PTape1.Visible Then
            HighScoreEntryInit()
        Else
            HSA1 = 24
            HSA2 = 6
            HSA3 = 12
        End If
    End If

End Sub

'******************
' Table power off
'******************
Sub table1_Exit
    SaveConfig()
    If B2SOn Then Controller.Stop 'stop the B2S controller
End Sub

'**************
'     TILT
'**************

Sub CheckTilt
    TiltMeasure = TiltMeasure + TiltSensitivity
    TiltDecreaseTimer.Enabled = True
    If TiltMeasure > 15 Then
        PlaySound "fx_knocker"
        isTilted = True
        TiltDisableTable()
        UpdateBackglass()
    End If
End Sub

Sub TiltDecreaseTimer_Timer
    ' DecreaseTilt
    If TiltMeasure > 0 Then
        TiltMeasure = TiltMeasure - 0.1
    Else
        TiltDecreaseTimer.Enabled = False
    End If
End Sub

Sub TiltDisableTable
    GiOff()
    FlippersDownLeft()
    FlippersDownRight()
    BumperGreenTop.Threshold = 100
    BumperGreenBottom.Threshold = 100
    BumperRedTop.Threshold = 100
    BumperRedBottom.Threshold = 100
    TurnOffPlayfieldLights()
    LeftSlingshot.Disabled = 1
    RightSlingshot.Disabled = 1
End Sub

Sub UnTiltEnableTable
    isTilted = False
    GiOn()
    BumperGreenTop.Threshold = 1
    BumperGreenBottom.Threshold = 1
    BumperRedTop.Threshold = 1
    BumperRedBottom.Threshold = 1
    LeftSlingshot.Disabled = 0
    RightSlingshot.Disabled = 0
End Sub


'***************************************************************
'             Supporting Ball & Sound Functions v4.0
'  includes random pitch in PlaySoundAt and PlaySoundAtBall
'***************************************************************

Dim TableWidth, TableHeight

TableWidth = Table1.width
TableHeight = Table1.height

'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************
''
''PlaySound(soundname, int loopcount, float volume, float pan, float randompitch, int pitch, bool useexisting, bool restart, float front_rear_fade)''
''           1         2              3             4          5                  6          7                 8             9

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 1000) 'note: adjusted this due to ball speed limit
    'Vol = Csng(BallVel(ball) ^2 / MaxVel ^2) ' impact at max velocity creates sound with max volume (max=1), impact at half max velocity creates sound with 0.25*volume
    'If Vol > 1 Then Vol = 1
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = tableobj.x * 2 / table1.width-1
    If tmp > 0 Then
        AudioPan = Csng(tmp ^10)
    Else
        AudioPan = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioFade(tableobj) 'only on VPX 10.4 and newer
    Dim tmp
    tmp = tableobj.y * 2 / TableHeight-1
    If tmp > 0 Then
        AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10))
    End If
End Function

Sub PlayRepeatingSoundAt(soundname, tableobj) 'play sound at X and Y position of an object, mostly bumpers, flippers and other fast objects
    PlaySound soundname, -1, 1, AudioPan(tableobj), 0.2, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAt(soundname, tableobj) 'play sound at X and Y position of an object, mostly bumpers, flippers and other fast objects
'useexisting=0, restart=0                                  *  *
    PlaySound soundname, 0, 1, AudioPan(tableobj), 0.2, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtUseExisting(soundname, tableobj) 'play sound at X and Y position of an object, mostly bumpers, flippers and other fast objects
'useexisting=1, restart=1                                  *  *
    PlaySound soundname, 0, 1, AudioPan(tableobj), 0.2, 0, 1, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname) ' play a sound at the ball position, like rubbers, targets, metals, plastics
    PlaySound soundname, 0, Vol(ActiveBall), AudioPan(ActiveBall), 0.2, Pitch(ActiveBall) * 10, 0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtBallFullVol(soundname) ' play a sound at the ball position, like rubbers, targets, metals, plastics
    PlaySound soundname, 0, 1, AudioPan(ActiveBall), 0.2, Pitch(ActiveBall) * 10, 0, 0, AudioFade(ActiveBall)
End Sub

Function RndNbr(n) 'returns a random number between 1 and n
    Randomize timer
    RndNbr = Int((n * Rnd) + 1)
End Function

'***********************************
'   JP's VP10 Rolling Sounds v4.0
'   JP's Ball Shadows
'   JP's Ball Speed Control
'   Rothbauer's dropping sounds
'***********************************

Const MaxVel = 35 'max ball velocity
Dim VelAdjFct
Dim isBallRolling(10) 'see cMaxBallsOnPlayfield
'ReDim rolling(MaxBallsOnPlayfield)
InitRolling()

Sub InitRolling
    Dim i
    For i = 0 to (MaxBallsOnPlayfield -1)
        isBallRolling(i) = False
    Next
End Sub

Sub RollingTimer_Timer()  'could also trigger logic from Realtime_Timer

    LostBallCheck()

    Dim BOT, b, ballpitch, ballvol, speedfactorx, speedfactory
    BOT = GetBalls

    ' stop the sound of deleted balls
    If UBound(BOT) < MaxBallsOnPlayfield then
  For b = UBound(BOT) + 1 to MaxBallsOnPlayfield
            isBallRolling(b) = False
      StopSound("fx_ballrolling" & b)
  Next
    End If

    ' exit the sub if no balls on the table
    If UBound(BOT) = -1 or Ubound(BOT) > MaxBallsOnPlayfield Then Exit Sub

    For b = 0 to UBound(BOT)
        ' play the rolling sound for each ball
        If BallVel(BOT(b)) > 1 Then
            If BOT(b).z < 30 Then
                ballpitch = Pitch(BOT(b))
                ballvol = Vol(BOT(b))
            Else
                ballpitch = Pitch(BOT(b)) + 50000 'increase the pitch on a ramp
                ballvol = Vol(BOT(b)) * 5
            End If
            isBallRolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, ballvol, AudioPan(BOT(b)), 0, ballpitch, 1, 0, AudioFade(BOT(b))
        Else
            If isBallRolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                isBallRolling(b) = False
            End If
        End If

        ' play ball drop sounds
        If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
            PlaySound "fx_balldrop" & b, 0, ABS(BOT(b).velz)/17, AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
        End If

        ' jps ball speed & spin control
        BOT(b).AngMomZ = BOT(b).AngMomZ * 0.95

        If isLimitBallVelocity Then
            If BOT(b).VelX <> 0 AND BOT(b).VelY <> 0 Then
                speedfactorX = ABS(maxvel / BOT(b).VelX)
                speedfactorY = ABS(maxvel / BOT(b).VelY)
                If speedfactorX <1 Then
                    BOT(b).VelX = BOT(b).VelX * speedfactorX
                    BOT(b).VelY = BOT(b).VelY * speedfactorX
                End If
                If speedfactorY <1 Then
                    BOT(b).VelX = BOT(b).VelX * speedfactorY
                    BOT(b).VelY = BOT(b).VelY * speedfactorY
                End If
            End If
        End If

''!!!experimental - new max velocity in any direction is constant
'        If BallVel(BOT(b)) > MaxVel Then
'           VelAdjFct = MaxVel / BallVel(BOT(b))
'            BOT(b).VelX = BOT(b).VelX * VelAdjFct
'            BOT(b).VelY = BOT(b).VelY * VelAdjFct
'        End If


    Next
End Sub

'*****************************************
' ninuzzu's BALL SHADOW
'*****************************************

Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5)

Sub BallShadowUpdate_timer()
    Dim BOT, b, c
    BOT = GetBalls

    For b = 0 to UBound(BallShadow)
        If b <= UBound(BOT) Then
            BallShadow(b).X = BOT(b).X + (BOT(b).X - (Table1.Width/2)) * 1.25 / BallSize
            BallShadow(b).Y = BOT(b).Y + 12
          BallShadow(b).Size_X = 5
          BallShadow(b).Size_Y = 5
            If BOT(b).Z > 20 Then
                BallShadow(b).visible = 1
            Else
                BallShadow(b).visible = 0
            End If
      Else
            BallShadow(b).visible = 0
        End If
    Next

End Sub

'**********************
' Lost Ball Check
'**********************
Sub LostBallCheck
    'if any ball is off of the playfield, set the location right in front of the drain so it gets destroyed
    Dim BOT, b
    BOT = GetBalls

    For b = 0 to UBound(BOT)
  'debug.print b & " " & BOT(b).x  & " " & BOT(b).y & " " & BOT(b).z
  if BOT(b).X < 0 or BOT(b).X > Table1.Width  _
            or BOT(b).Y < 0 or BOT(b).Y > Table1.Height _
      or BOT(b).z < -500 _
      or b > MaxBallsOnPlayfield Then
                'debug.print b & " " & BOT(b).X & " " & BOT(b).Y
    BOT(b).x = Drain.x
    BOT(b).y = Drain.y - BOT(b).Radius
    BOT(b).z = BOT(b).Radius
  end if
    Next
End Sub


'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub


'***************************************
'     Collection collision sounds
'***************************************

Sub cHitTarget_Hit (idx)
    PlaySoundAtBall "hittarget"
'debug.print "cHitTarget_Hit " & idx
End Sub

Sub cStandupTargets_Hit (idx)
    PlaySoundAt "target", cStandupTargets(idx)
'debug.print "cStandupTargets_Hit " & idx
End Sub

Sub cRails_Hit (idx)
    PlaySoundAtBall "metalhit_medium"
'debug.print "cRails_Hit " & idx
End Sub

Sub cGates_Hit (idx)
    PlaySoundAtBall "gate4"
'debug.print "cGates_Hit " & idx
End Sub

Sub cRubberPegs_Hit (idx)
    cRubbers_Hit(idx)
'debug.print "cRubberPegs_Hit " & idx
End Sub

Sub cRubberBands_Hit (idx)
    cRubbers_Hit(idx)
'debug.print "cRubberBands_Hit " & idx
End Sub

Sub cRubberBandsLong_Hit (idx)
    cRubbers_Hit(idx)
'debug.print "cRubberBandsLong_Hit " & idx
End Sub

Sub cRubberPins_Hit (idx)
    PlaySoundAtBall "pinhit_low"
'debug.print "cRubberPins_Hit " & idx
End Sub

Sub cRubberPosts_Hit(idx)
'debug.print "cRubberPosts_Hit " & idx
  dim finalspeed
    finalspeed=SQR(activeball.velx ^2 + activeball.vely ^2)
  If finalspeed > 16 then
            PlaySoundAtBallFullVol "fx_rubber2"
'   PlaySound "fx_rubber2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  ElseIf finalspeed >= 6 then
      Select Case Int(Rnd*3)+1
        Case 1 : PlaySoundAtBallFullVol "rubber_hit_1"
        Case 2 : PlaySoundAtBallFullVol "rubber_hit_2"
        Case 3 : PlaySoundAtBallFullVol "rubber_hit_3"
        End Select
  End If
End Sub

Sub cMetalWires_Hit (idx)
    PlaySoundAtBallFullVol "metalhit_medium"
'debug.print "cMetalWires_Hit " & idx
End Sub

Sub cMetals_Hit (idx)
    PlaySoundAtBall "metalhit2"
'debug.print "cMetals_Hit " & idx
End Sub

Sub cRubbers_Hit(idx)
'debug.print "cRubbers_Hit " & idx
    dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
    If finalspeed > 20 then
        PlaySoundAtBall "fx_rubber2"
    ElseIf finalspeed >= 6 then
        Select Case Int(Rnd*3)+1
            Case 1 : PlaySoundAtBall "rubber_hit_1"
      Case 2 : PlaySoundAtBall "rubber_hit_2"
          Case 3 : PlaySoundAtBall "rubber_hit_3"
        End Select
    End If
End Sub

Sub cWoods_Hit (idx)
'debug.print "cWoods_Hit " & idx
    PlaySoundAtBall "fx_woodhit"
End Sub

Sub cPlastics_Hit (idx)
'debug.print "cPlastics_Hit " & idx
    PlaySoundAtBall "fx_plastichit"
End Sub

Sub LeftFlipper_Collide(parm)
    RandomSoundFlipper()
End Sub
Sub RightFlipper_Collide(parm)
    RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
    Select Case Int(Rnd*3)+1
        Case 1 : PlaySoundAtBall "flip_hit_1"
        Case 2 : PlaySoundAtBall "flip_hit_2"
        Case 3 : PlaySoundAtBall "flip_hit_3"
    End Select
End Sub

' ****************************************
'  Score functions
' ****************************************

Dim RemainingPointsToAdd

Sub AddScore(pointsToAdd)
    If isTilted Then Exit Sub

    RemainingPointsToAdd = RemainingPointsToAdd + pointsToAdd
    AddToScoreInParts()
End Sub

Sub AddToScoreInParts

'??? chimes??

    If isTilted Then
        RemainingPointsToAdd = 0
    End If

    If RemainingPointsToAdd = 0 Then Exit Sub

    If RemainingPointsToAdd >= 1000 Then
        If isRingBell100 Then PlaySound "Better100PointBell", 0, 0.5, 0, 0, 0, 1, 1, 0
        AddScoreNoSound 1000
        RemainingPointsToAdd = RemainingPointsToAdd - 1000
    ElseIf RemainingPointsToAdd >= 100 Then
        If isRingBell100 Then PlaySound "Better100PointBell", 0, 0.5, 0, 0, 0, 1, 1, 0
        AddScoreNoSound 100
        RemainingPointsToAdd = RemainingPointsToAdd - 100
    ElseIf RemainingPointsToAdd >= 10 Then
        If isRingBell10 Then PlaySound "Better10PointBell",  0, 0.5, 0, 0, 0, 1, 1, 0
        AddScoreNoSound 10
        RemainingPointsToAdd = RemainingPointsToAdd - 10
        Advance10counter
    ElseIf RemainingPointsToAdd >= 1 Then
'shouldn't happen on this table
        AddScoreNoSound 1
        RemainingPointsToAdd = RemainingPointsToAdd - 1
    End If

    CheckAwardSpecial()
    AdvanceMatchCount()

    If RemainingPointsToAdd > 0 Then
        vpmtimer.addtimer 150, "AddToScoreInParts '"
    End If

End Sub


Sub AddScoreNoSound(pointsToAdd)
    If isTilted Then Exit Sub
    PlayerScore(CurrentPlayer) = PlayerScore(CurrentPlayer) + pointsToAdd
    UpdateBackglass()
End Sub

Dim MatchCounter: MatchCounter = Int(10 * Rnd()) 'pick a random starting number - does not reset for new game
Sub AdvanceMatchCount()
'EM match number generated by a turning counter disk.  Match counter counts down instead of up, creating random-appearing numbers.
'some machines have the 10 numbers in a non-sequential order but from what I have read, Bally machines of this era did not.

    If matchCounter <= 0 Then
        matchCounter = 9
    Else
        MatchCounter = MatchCounter - 1
    End If

End Sub

Dim TenCounter :TenCounter= Int(10*Rnd)
Sub Advance10Counter()

    TenCounter = TenCounter + 1
    If TenCounter > 9 Then TenCounter = 0

    CheckTenCounterLights()
End Sub

Sub CheckTenCounterLights()
    If TenCounter mod 2 = 0 Then
        PLAdvanceYellowBonus.State=1
        PLAdvanceBlueBonus.State=0
    Else
        PLAdvanceYellowBonus.State=0
        PLAdvanceBlueBonus.State=1
    End If

    If TenCounter=0 OR TenCounter=3 OR TenCounter=5 OR TenCounter=8 Then
        PLKickerOn.State=1
        InvisibleApRollBallWall.IsDropped=True
        InvisibleApHoldBallWall.IsDropped=False
    Else
        PLKickerOn.State=0
        InvisibleApRollBallWall.IsDropped=False
        InvisibleApHoldBallWall.IsDropped=True
    End If

    If TenCounter=1 Then
        PLextraBall.State=1
    Else
        PLextraBall.State=0
    End If

'extra ball lite controlled by 10's counter - 1 of the 10
'slingshots are 10 when lit 100 and advance bonus.  toggle on 10 point unit
'left save kicker on 4 of the 10 spot on 10's unit
End Sub

Dim SpecialAwarded : SpecialAwarded=0
Sub CheckAwardSpecial()
    If SpecialAwarded=0 AND PlayerScore(CurrentPlayer) >= PointsToEarnSpecial1 Then
        SpecialAwarded=1
        AwardSpecial()
    End If
    If SpecialAwarded=1 AND PlayerScore(CurrentPlayer) >= PointsToEarnSpecial2 Then
        SpecialAwarded=2
        AwardSpecial()
    End If
    If SpecialAwarded=2 AND PlayerScore(CurrentPlayer) >= PointsToEarnSpecial3 Then
        SpecialAwarded=3
        AwardSpecial()
    End If
    If SpecialAwarded=3 AND PlayerScore(CurrentPlayer) >= PointsToEarnSpecial4 Then
        SpecialAwarded=4
        AwardSpecial()
    End If
End Sub

Sub AwardSpecial
    PlaySound "fx_knocker"
    AddCredit()
    UpdateBackglass()
End Sub

'***********
' Playfield lights
'***********

Sub TurnOnPlayfieldLights()
    Dim bulb
    For each bulb in cGameLights
        bulb.State = 1
    Next
End Sub

Sub TurnOffPlayfieldLights()
    Dim bulb
    For each bulb in cGameLights
        bulb.State = 0
    Next
End Sub

Sub LiteGreenBumpers()
    BLGreenTop1.State=1
    BLGreenTop2.State=1
    BLGreenTop3.State=1
    BLGreenBottom1.State=1
    BLGreenBottom2.State=1
    BLGreenBottom3.State=1
End Sub

Sub LiteRedBumpers()
    BLRedTop1.State=1
    BLRedTop2.State=1
    BLRedTop3.State=1
    BLRedBottom1.State=1
    BLRedBottom2.State=1
    BLRedBottom3.State=1
End Sub

Sub TurnOffBumperLights()
    BLGreenTop1.State=0
    BLGreenTop2.State=0
    BLGreenTop3.State=0
    BLGreenBottom1.State=0
    BLGreenBottom2.State=0
    BLGreenBottom3.State=0
    BLRedTop1.State=0
    BLRedTop2.State=0
    BLRedTop3.State=0
    BLRedBottom1.State=0
    BLRedBottom2.State=0
    BLRedBottom3.State=0
End Sub

'***********
' GI lights
'***********

Sub GiOn
  'debug.print "GiOn"
  PlaySound"fx_gion"
  Dim bulb
    For each bulb in cGiLights
        bulb.State = 1
    Shadows_posts.visible=True
        'debug.print bulb.name
    Next

End Sub

Sub GiOff
  'debug.print "GiOff"
    PlaySound"fx_gioff"

    Dim bulb
    For each bulb in cGiLights
        bulb.State = 0
    Shadows_posts.visible=False
        'debug.print bulb.name
    Next

End Sub

'Dim GiIntensity
'GiIntensity = 1               'can be used by the LUT changing to increase the GI lights when the table is darker'''

'Sub GiIntensity(factor) 'changes the intensity scale
'    Dim bulb
'    For each bulb in cGiLights
'        bulb.IntensityScale = GiIntensity * factor
'    Next
'End Sub

' ********************************
'        Attract Mode
' ********************************
' use the"Blink Pattern" of each light

Sub StartAttractMode()
'no attract mode for this EM machine
End Sub

Sub StopAttractMode()
End Sub

'Sub BackglassAttractTimer_Timer
'not used for this machine
'End Sub

'Sub BackglassAttractTitleOn
'not used for this machine
'     If Not B2SOn Then Exit Sub
'End Sub

'Sub BackglassAttractTitleOff
'not used for this machine
'     If Not B2SOn Then Exit Sub
'End Sub


' ********************************
'        Config load and save
' ********************************


Sub LoadConfig

    PlayerScore(1) = 19710 ' default
    HighScoreAllTime= 0
    HSA1 = 24
    HSA2 = 6
    HSA3 = 12
    Dim FileObj
    Dim ConfigFile, TextStr, TextLine

    Set FileObj = CreateObject("Scripting.FileSystemObject")
    If Not FileObj.FolderExists(UserDirectory) Then Exit Sub
    If Not FileObj.FileExists(UserDirectory & TableName & ".txt") Then Exit Sub
    Set ConfigFile = FileObj.GetFile(UserDirectory & TableName & ".txt")
    Set TextStr = ConfigFile.OpenAsTextStream(1, 0)

    Do until TextStr.AtEndOfStream
        TextLine = TextStr.ReadLine
      If Mid(TextLine,1,16) = "playerscore1    " Then
          On Error Resume Next
          PlayerScore(1) = CInt(Mid(TextLine,17,6))
      End If
      If Mid(TextLine,1,16) = "playerscore2    " Then
          On Error Resume Next
          PlayerScore(2) = CInt(Mid(TextLine,17,6))
      End If
      If Mid(TextLine,1,16) = "playerscore3    " Then
          On Error Resume Next
          PlayerScore(3) = CInt(Mid(TextLine,17,6))
      End If
      If Mid(TextLine,1,16) = "playerscore4    " Then
          On Error Resume Next
          PlayerScore(4) = CInt(Mid(TextLine,17,6))
      End If
        If Mid(TextLine,1,16) =  "highscore       " Then
          On Error Resume Next
            HighScoreAllTime= CInt(Mid(TextLine,17,6))
      End If
        If Mid(TextLine,1,16) =  "hsinitials1     " Then
          On Error Resume Next
            HSA1 = CInt(Mid(TextLine,17,6))
      End If
        If Mid(TextLine,1,16) =  "hsinitials2     " Then
          On Error Resume Next
            HSA2 = CInt(Mid(TextLine,17,6))
      End If
        If Mid(TextLine,1,16) =  "hsinitials3     " Then
          On Error Resume Next
            HSA3 = CInt(Mid(TextLine,17,6))
      End If
    Loop

    TextStr.Close
    Set ConfigFile = Nothing
    Set FileObj = Nothing

    On Error Resume Next

    UpdateBackglass()
    UpdatePostIt()

End Sub

Sub SaveConfig
    Dim FileObj
    Set FileObj = CreateObject("Scripting.FileSystemObject")
    If Not FileObj.FolderExists(UserDirectory)then
        Exit Sub
    End If
  Dim ConfigFile
    Set ConfigFile = FileObj.CreateTextFile(UserDirectory & TableName & ".txt", True)
    ConfigFile.WriteLine "playerscore1    " & PlayerScore(1)
    ConfigFile.WriteLine "playerscore2    " & PlayerScore(2)
    ConfigFile.WriteLine "playerscore3    " & PlayerScore(3)
    ConfigFile.WriteLine "playerscore4    " & PlayerScore(4)
    ConfigFile.WriteLine "highscore       " & HighScoreAllTime
    ConfigFile.WriteLine "hsinitials1     " & HSA1
    ConfigFile.WriteLine "hsinitials2     " & HSA2
    ConfigFile.WriteLine "hsinitials3     " & HSA3
    ConfigFile.Close
    Set ConfigFile = Nothing
    Set FileObj = Nothing
End Sub


'***************************************************************
' Backglass display
'***************************************************************

Sub UpdateBackglass()

    'handles backglass and backdrop display
    ' shoot again || game over || credits || active player || number of players || player turn (aka ball in play) || tilt

    If isGameInPlay Then
        TextGameOver.Text = ""
        TextMatchNumber.Text = ""
        If B2SOn Then
            Controller.B2SSetGameOver 0
            Controller.B2SSetMatch 0
        End If
        LightPlayer1.state=1
        If B2SOn Then Controller.B2SSetCanPlay 1
        If NumberOfPlayers > 1 Then
            LightPlayer2.state=1
            If B2SOn Then Controller.B2SSetCanPlay 2
        End If
        If NumberOfPlayers > 2 Then
            LightPlayer3.state=1
            If B2SOn Then Controller.B2SSetCanPlay 3
        End If
        If NumberOfPlayers > 3 Then
            LightPlayer4.state=1
            If B2SOn Then Controller.B2SSetCanPlay 4
        End If

        LightPlayer1.state=0
        LightPlayer2.state=0
        LightPlayer3.state=0
        LightPlayer4.state=0
        If CurrentPlayer = 1 Then
            LightPlayer1.state=2
            If B2SOn Then Controller.B2SSetPlayerUp 1
        ElseIf CurrentPlayer = 2 Then
            LightPlayer2.state=2
            If B2SOn Then Controller.B2SSetPlayerUp 2
        ElseIf CurrentPlayer = 3 Then
            LightPlayer3.state=2
            If B2SOn Then Controller.B2SSetPlayerUp 3
        ElseIf CurrentPlayer = 4 Then
            LightPlayer4.state=2
            If B2SOn Then Controller.B2SSetPlayerUp 4
        End If
    Else
        TextGameOver.Text = "GAME OVER"
        If MatchCounter <> "" Then
            TextMatchNumber.Text = MatchCounter & "0"
        End If
        If B2SOn Then
            Controller.B2SSetGameOver 1
            If MatchCounter = 0 Then
                Controller.B2SSetMatch 10
            Else
                Controller.B2SSetMatch MatchCounter
            End If
        End If
        If B2SOn Then Controller.B2SSetCanPlay 0
        If B2SOn Then Controller.B2SSetPlayerUp 0
    End If

    ReelCredits.SetValue Credits
    If B2SOn Then Controller.B2SSetCredits Credits

    If isGameInPlay Then
        ballinplay = BallsPerGame - BallsRemaining + 1
        If ballinplay < 0 OR ballinplay > BallsPerGame Then
          ballinplay = 0
            TextBallInPlay.Text = ""
        Else
            TextBallInPlay.Text = "BALL " & ballinplay
        End If

        If B2SOn Then Controller.B2SSetBallInPlay ballinplay
    Else
        If B2SOn Then Controller.B2SSetBallInPlay 0
    End If

    If isTilted Then
        TextTilt.Text = "TILT"
        If B2SOn Then Controller.B2SSetTilt 1
    Else
        TextTilt.Text = ""
        If B2SOn Then Controller.B2SSetTilt 0
    End If

    If CurrentPlayer = 1 Then
        If PlayerScore(1) > 0 Then
            ReelScore1.SetValue PlayerScore(1)
        Else
            ReelScore1.ResetToZero
        End If
    ElseIf CurrentPlayer = 2 Then
        If PlayerScore(2) > 0 Then
            ReelScore2.SetValue PlayerScore(2)
        Else
            ReelScore2.ResetToZero
        End If
    ElseIf CurrentPlayer = 3 Then
        If PlayerScore(3) > 0 Then
            ReelScore3.SetValue PlayerScore(3)
        Else
            ReelScore3.ResetToZero
        End If
    ElseIf CurrentPlayer = 4 Then
        If PlayerScore(4) > 0 Then
            ReelScore4.SetValue PlayerScore(4)
        Else
            ReelScore4.ResetToZero
        End If

    End If

    If B2SOn Then
    Controller.B2SSetScorePlayer CurrentPlayer, PlayerScore(CurrentPlayer)
    If PLShootAgain.state = 1 Then
            Controller.B2SSetShootAgain 1
        Else
            Controller.B2SSetShootAgain 0
        End If
    End If

End Sub


'******************
'     Match
'******************

Sub Check_For_Match()
    If isTilted Then Exit Sub
    PlaySound "fx_match"
    If INT (PlayerScore(1) / 10) MOD 10 = MatchCounter Then
        PlaySound "fx_knocker"
        AddCredit()
    End If
    UpdateBackglass()
End Sub

'***************************************************************************************
' Misc Routines
'***************************************************************************************

Sub AwardExtraBall()
    If PLShootAgain.State=0 Then
    PlaySound "fx_knocker"
    PLShootAgain.State=1
    End If
End Sub


'*******
' Bonus
'*******
Dim BonusAmountYellow
Sub AdvanceYellowBonus()
    If PlYellow1K.State=1 Then
        'Nothing, already at max
    ElseIf PlYellow900.State=1 Then
        PlYellow900.State=0:PlYellow1K.State=1
    ElseIf PlYellow800.State=1 Then
        PlYellow800.State=0:PlYellow900.State=1
    ElseIf PlYellow700.State=1 Then
        PlYellow700.State=0:PlYellow800.State=1
    ElseIf PlYellow600.State=1 Then
        PlYellow600.State=0:PlYellow700.State=1
    ElseIf PlYellow500.State=1 Then
        PlYellow500.State=0:PlYellow600.State=1
    ElseIf PlYellow400.State=1 Then
        PlYellow400.State=0:PlYellow500.State=1
    ElseIf PlYellow300.State=1 Then
        PlYellow300.State=0:PlYellow400.State=1
    ElseIf PlYellow200.State=1 Then
        PlYellow200.State=0:PlYellow300.State=1
    ElseIf PlYellow100.State=1 Then
        PlYellow100.State=0:PlYellow200.State=1
    Else
        PlYellow100.State=1
    End If
End Sub

Sub AdvanceBlueBonus()
    If PlBlue1K.State=1 Then
        'Nothing, already at max
    ElseIf PlBlue900.State=1 Then
        PlBlue900.State=0:PlBlue1K.State=1
    ElseIf PlBlue800.State=1 Then
        PlBlue800.State=0:PlBlue900.State=1
    ElseIf PlBlue700.State=1 Then
        PlBlue700.State=0:PlBlue800.State=1
    ElseIf PlBlue600.State=1 Then
        PlBlue600.State=0:PlBlue700.State=1
    ElseIf PlBlue500.State=1 Then
        PlBlue500.State=0:PlBlue600.State=1
    ElseIf PlBlue400.State=1 Then
        PlBlue400.State=0:PlBlue500.State=1
    ElseIf PlBlue300.State=1 Then
        PlBlue300.State=0:PlBlue400.State=1
    ElseIf PlBlue200.State=1 Then
        PlBlue200.State=0:PlBlue300.State=1
    ElseIf PlBlue100.State=1 Then
        PlBlue100.State=0:PlBlue200.State=1
    Else
        PlBlue100.State=1
    End If
End Sub

Dim ScoreMotorCount

Sub TimerAwardYellowBonus_Timer()

    If isClearBonus Then
        TimerAwardYellowBonus.Interval = 150
    Else
        TimerAwardYellowBonus.Interval = 250
    End If

    ScoreMotorCount = ScoreMotorCount + 1
' pause slightly after 5 adds just like real machine
    If ScoreMotorCount = 6 Then
        ScoreMotorCount = 0
        PlaySound "MotorReset"
        Exit Sub
    End If

    If PLyellow1K.State=1 Then
        PLyellow1K.State=0
        PLyellow900.State=1
    ElseIf PLyellow900.State=1 Then
        PLyellow900.State=0
        PLyellow800.State=1
    ElseIf PLyellow800.State=1 Then
        PLyellow800.State=0
        PLyellow700.State=1
    ElseIf PLyellow700.State=1 Then
        PLyellow700.State=0
        PLyellow600.State=1
    ElseIf PLyellow600.State=1 Then
        PLyellow600.State=0
        PLyellow500.State=1
    ElseIf PLyellow500.State=1 Then
        PLyellow500.State=0
        PLyellow400.State=1
    ElseIf PLyellow400.State=1 Then
        PLyellow400.State=0
        PLyellow300.State=1
    ElseIf PLyellow300.State=1 Then
        PLyellow300.State=0
        PLyellow200.State=1
    ElseIf PLyellow200.State=1 Then
        PLyellow200.State=0
        PLyellow100.State=1
    ElseIf PLyellow100.State=1 Then
        PLyellow100.State=0
    Else
        KickerLeft.TimerEnabled=True
        Me.Enabled=False
        PLyellow100.State=1
        PLyellowX10.State=0
        ScoreMotorCount=0
        Exit Sub
    End If

    If Not isClearBonus Then
        If PLyellowX10.State=1 Then
            AddScore 1000
        Else
            AddScore 100
        End If
    End If
    PlaySound "MotorPulse"

End Sub


Sub TimerAwardBlueBonus_Timer()

debug.print "timerawardbluebonus_timer"
    If isClearBonus Then
        TimerAwardBlueBonus.Interval = 150
    Else
        TimerAwardBlueBonus.Interval = 250
    End If

    ScoreMotorCount = ScoreMotorCount + 1
' pause slightly after 5 adds just like real machine
    If ScoreMotorCount = 6 Then
        ScoreMotorCount = 0
        PlaySound "MotorReset"
        Exit Sub
    End If

    If PLblue1K.State=1 Then
debug.print "1k"
        PLblue1K.State=0
        PLblue900.State=1
    ElseIf PLblue900.State=1 Then
debug.print "900"
        PLblue900.State=0
        PLblue800.State=1
    ElseIf PLblue800.State=1 Then
debug.print "800"
        PLblue800.State=0
        PLblue700.State=1
    ElseIf PLblue700.State=1 Then
debug.print "700"
        PLblue700.State=0
        PLblue600.State=1
    ElseIf PLblue600.State=1 Then
debug.print "600"
        PLblue600.State=0
        PLblue500.State=1
    ElseIf PLblue500.State=1 Then
debug.print "500"
        PLblue500.State=0
        PLblue400.State=1
    ElseIf PLblue400.State=1 Then
debug.print "400"
        PLblue400.State=0
        PLblue300.State=1
    ElseIf PLblue300.State=1 Then
debug.print "300"
        PLblue300.State=0
        PLblue200.State=1
    ElseIf PLblue200.State=1 Then
debug.print "200"
        PLblue200.State=0
        PLblue100.State=1
    ElseIf PLblue100.State=1 Then
debug.print "100"
        PLblue100.State=0
    Else
        KickerRight.TimerEnabled=True
        Me.Enabled=False
        PLblue100.State=1
        PLblueX10.State=0
        ScoreMotorCount=0
        Exit Sub
    End If

    If Not isClearBonus Then
        If PLblueX10.State=1 Then
            AddScore 1000
        Else
            AddScore 100
        End If
    End If
    PlaySound "MotorPulse"

End Sub


''''
' post logic from SpaceTime table
Const PostUpZ = 25
Const PostDownZ = 0

Sub PostUp() 'Play-More Post
  If CenterPost_prim.TransZ=PostDownZ Then
    PlaySoundAt "solenoid", centerpost_prim
  End If

  lp2.state = 1
  lp3.state = 1

  Post.collidable = true

  Centerpost_prim.image="3D_Centerpost"
  Centerpost_prim.BlendDisableLighting = 0.3
  CenterPost_prim.TransZ = PostUpZ
End Sub

Sub PostDown()
  If CenterPost_prim.TransZ=PostUpZ Then
    PlaySoundAt "solenoid", centerpost_prim
  End If

  lp2.state = 0
  lp3.state = 0

  Post.collidable = false

  Centerpost_prim.image="3D_Centerpost"
  Centerpost_prim.BlendDisableLighting = 0
  CenterPost_prim.Transz = PostDownZ
End Sub


Sub OpenTopGate()
  InvisibleTopGateBlockerWall.IsDropped=True
    If PLTopGateOpen.State=1 Then Exit Sub
  If PGateWire.RotX > -90 Then
        PGateWire.RotX = PGateWire.RotX - 2
        vpmtimer.addtimer 35, "OpenTopGate '"
    Else
       PGateWire.RotX = -90
       PlaySoundAt "gate4", PgateWire
       PLTopGateOpen.State=1
    End If
End Sub

Sub CloseTopGate()
  InvisibleTopGateBlockerWall.IsDropped=False
    If PLTopGateOpen.State=0 Then Exit Sub
  If PGateWire.RotX < 0 Then
        PGateWire.RotX = PGateWire.RotX + 2
        vpmtimer.addtimer 35, "CloseTopGate '"
    Else
       PGateWire.RotX = 0
       PlaySoundAt "gate4", PgateWire
       PLTopGateOpen.State=0
    End If
End Sub

Sub OpenBottomGate()
' rotate 30 = open, 90=closed
    WallBottomGateToPlunger.IsDropped = True
    WallBottomGateToDrain.IsDropped = False
    If PLBottomGateOpen.State=1 Then Exit Sub
'Open
    If PBottomGate.RotY > 30 Then
        PBottomGate.RotY = PBottomGate.RotY - 2
        vpmtimer.addtimer 15, "OpenBottomGate '"
    Else
        PBottomGate.RotY = 30
        PlaySoundAt "fx_plastichit", TriggerRolloverRightOutlane
        PLBottomGateOpen.State=1
    End If
End Sub

Sub CloseBottomGate()
    WallBottomGateToPlunger.IsDropped = False
    WallBottomGateToDrain.IsDropped = True
    If PLBottomGateOpen.State=0 Then Exit Sub
'rotate to 90 for Closed
    If PBottomGate.RotY < 90 Then
        PBottomGate.RotY = PBottomGate.RotY + 2
        vpmtimer.addtimer 15, "CloseBottomGate '"
    Else
        PBottomGate.RotY = 90
        PlaySoundAt "fx_plastichit", TriggerRolloverRightOutlane
        PLBottomGateOpen.State=0
    End If
End Sub


'* * * * * * * * * *
' F12 menu options - code based on VPW example table
'* * * * * * * * * *

Dim isRingBell10: isRingBell10 = True
Dim isRingBell100: isRingBell100 = True
Dim ballsPerGame: ballsPerGame = 5
Dim isLimitBallVelocity: isLimitBallVelocity = False


' Called when options are tweaked by the player.
' - 0: game has started, good time to load options and adjust accordingly
' - 1: an option has changed
' - 2: options have been reset
' - 3: player closed the tweak UI, good time to update staticly prerendered parts
' Table1.Option arguments are:
'          1     2              3              4                          5              6                         7
' - option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional arry of literal strings
Dim dspTriggered : dspTriggered = False
Dim ColorLUT : ColorLUT = 1           ' Color desaturation LUTs: 1 to 11, where 1 is normal and 11 is black'n'white
Dim AnswerNumber

Sub Table1_OptionEvent(ByVal eventId)

    If eventId = 1 And Not dspTriggered Then dspTriggered = True : DisableStaticPreRendering = True : End If
    If eventId = 3 And dspTriggered Then dspTriggered = False : DisableStaticPreRendering = False : End If

  ' Color Saturation
    ColorLUT = Table1.Option("Color Saturation", 1, 11, 1, 1, 0, _
    Array("Normal", "Desaturated 10%", "Desaturated 20%", "Desaturated 30%", "Desaturated 40%", "Desaturated 50%", _
        "Desaturated 60%", "Desaturated 70%", "Desaturated 80%", "Desaturated 90%", "Black 'n White"))
  if ColorLUT = 1 Then Table1.ColorGradeImage = ""
  if ColorLUT = 2 Then Table1.ColorGradeImage = "colorgradelut256x16-10"
  if ColorLUT = 3 Then Table1.ColorGradeImage = "colorgradelut256x16-20"
  if ColorLUT = 4 Then Table1.ColorGradeImage = "colorgradelut256x16-30"
  if ColorLUT = 5 Then Table1.ColorGradeImage = "colorgradelut256x16-40"
  if ColorLUT = 6 Then Table1.ColorGradeImage = "colorgradelut256x16-50"
  if ColorLUT = 7 Then Table1.ColorGradeImage = "colorgradelut256x16-60"
  if ColorLUT = 8 Then Table1.ColorGradeImage = "colorgradelut256x16-70"
  if ColorLUT = 9 Then Table1.ColorGradeImage = "colorgradelut256x16-80"
  if ColorLUT = 10 Then Table1.ColorGradeImage = "colorgradelut256x16-90"
  if ColorLUT = 11 Then Table1.ColorGradeImage = "colorgradelut256x16-100"


    AnswerNumber = Table1.Option("Flipper Color", 1, 3, 1, 1, 0, Array("Red", "Yellow", "Black"))
    If AnswerNumber = 1 Then
        MaterialColor "RubberOnFlipper", RGB(138,0,0)
    ElseIf AnswerNumber = 2 Then
        MaterialColor "RubberOnFlipper", RGB(224,192,0)
    ElseIf AnswerNumber = 3 Then
        MaterialColor "RubberOnFlipper", RGB(2,2,2)
    End If

    AnswerNumber = Table1.Option("Flipper Physics", 1, 2, 1, 1, 0, Array("Standard", "Complex"))
    If AnswerNumber = 2 Then
        TimerFlipperTricks.Enabled=True
    Else
        TimerFlipperTricks.Enabled=False
    End If

    AnswerNumber = Table1.Option("Limit Ball Velocity", 1, 2, 1, 1, 0, Array("Yes", "No"))
    If AnswerNumber = 1 Then
        isLimitBallVelocity = True
    Else
        isLimitBallVelocity = False
    End If

    AnswerNumber = Table1.Option("Ring 10 Point Bell", 1, 2, 1, 1, 0, Array("Yes", "No"))
    If AnswerNumber = 1 Then
        isRingBell10 = True
    Else
        isRingBell10 = False
    End If

    AnswerNumber = Table1.Option("Ring 100/1000 Point Bell", 1, 2, 1, 1, 0, Array("Yes", "No"))
    If AnswerNumber = 1 Then
        isRingBell100 = True
    Else
        isRingBell100 = False
    End If

    AnswerNumber = Table1.Option("Balls per game", 1, 2, 1, 2, 0, Array("3", "5"))
    If AnswerNumber = 1 Then
        ballsPerGame=3
    Else
        ballsPerGame=5
    End If

    AnswerNumber = Table1.Option("Show Highscore Tape", 1, 2, 1, 2, 0, Array("Yes", "No"))
    If AnswerNumber = 1 Then
        Dim item
        For each item in cHighscoreStuff
            item.Visible = 1
        Next
    Else
        Dim itemh
        For each itemh in cHighscoreStuff
            itemh.Visible = 0
        Next
    End If

End Sub

'==========================================================================================================================================
'============================================================= START OF HIGH SCORES ROUTINES =============================================================
'==========================================================================================================================================
'
'ADD LINE TO TABLE_KEYDOWN SUB WITH THE FOLLOWING:    If HSEnterMode Then HighScoreProcessKey(keycode) AFTER THE STARTGAME ENTRY
'ADD: And Not HSEnterMode=true TO IF KEYCODE=STARTGAMEKEY
'TO SHOW THE SCORE ON POST-IT ADD LINE AT RELEVENT LOCATION THAT HAS:  UpdatePostIt
'TO INITIATE ADDING INITIALS ADD LINE AT RELEVENT LOCATION THAT HAS:  HighScoreEntryInit()
'ADD THE FOLLOWING LINES TO TABLE_INIT TO SETUP POSTIT
' if HSA1="" then HSA1=25
' if HSA2="" then HSA2=25
' if HSA3="" then HSA3=25
' UpdatePostIt
'ADD HSA1, HSA2 AND HSA3 TO SAVE AND LOAD VALUES FOR TABLE
'ADD A TIMER NAMED HighScoreFlashTimer WITH INTERVAL 100 TO TABLE
'SET HSSSCOREX BELOW TO WHATEVER VARIABLE YOU USE FOR HIGH SCORE.
'ADD OBJECTS TO PLAYFIELD (EASIEST TO JUST COPY FROM THIS TABLE)
'IMPORT POST-IT IMAGES


Dim HSA1, HSA2, HSA3
Dim HSEnterMode, hsLetterFlash, hsEnteredDigits(3), hsCurrentDigit, hsCurrentLetter
Dim HSArray
Dim HSScoreM,HSScore100k, HSScore10k, HSScoreK, HSScore100, HSScore10, HSScore1, HSScorex 'Define 6 different score values for each reel to use
HSArray = Array("Postit0","postit1","postit2","postit3","postit4","postit5","postit6","postit7","postit8","postit9","postitBL","postitCM","Tape")
Const hsFlashDelay = 4

' ***********************************************************
'  HiScore DISPLAY
' ***********************************************************

Sub UpdatePostIt
  dim tempscore
  HSScorex = HighScoreAllTime
  TempScore = HSScorex
  HSScore1 = 0
  HSScore10 = 0
  HSScore100 = 0
  HSScoreK = 0
  HSScore10k = 0
  HSScore100k = 0
  HSScoreM = 0
  if len(TempScore) > 0 Then
    HSScore1 = cint(right(Tempscore,1))
  end If
  if len(TempScore) > 1 Then
    TempScore = Left(TempScore,len(TempScore)-1)
    HSScore10 = cint(right(Tempscore,1))
  end If
  if len(TempScore) > 1 Then
    TempScore = Left(TempScore,len(TempScore)-1)
    HSScore100 = cint(right(Tempscore,1))
  end If
  if len(TempScore) > 1 Then
    TempScore = Left(TempScore,len(TempScore)-1)
    HSScoreK = cint(right(Tempscore,1))
  end If
  if len(TempScore) > 1 Then
    TempScore = Left(TempScore,len(TempScore)-1)
    HSScore10k = cint(right(Tempscore,1))
  end If
  if len(TempScore) > 1 Then
    TempScore = Left(TempScore,len(TempScore)-1)
    HSScore100k = cint(right(Tempscore,1))
  end If
  if len(TempScore) > 1 Then
    TempScore = Left(TempScore,len(TempScore)-1)
    HSScoreM = cint(right(Tempscore,1))
  end If
  Pscore6.image = HSArray(HSScoreM):If HSScorex<1000000 Then PScore6.image = HSArray(10)
  Pscore5.image = HSArray(HSScore100K):If HSScorex<100000 Then PScore5.image = HSArray(10)
  PScore4.image = HSArray(HSScore10K):If HSScorex<10000 Then PScore4.image = HSArray(10)
  PScore3.image = HSArray(HSScoreK):If HSScorex<1000 Then PScore3.image = HSArray(10)
  PScore2.image = HSArray(HSScore100):If HSScorex<100 Then PScore2.image = HSArray(10)
  PScore1.image = HSArray(HSScore10):If HSScorex<10 Then PScore1.image = HSArray(10)
  PScore0.image = HSArray(HSScore1):If HSScorex<1 Then PScore0.image = HSArray(10)
  if HSScorex<1000 then
    PComma.image = HSArray(10)
  else
    PComma.image = HSArray(11)
  end if
  if HSScorex<1000000 then
    PComma2.image = HSArray(10)
  else
    PComma2.image = HSArray(11)
  end if
' if showhisc=1 and showhiscnames=1 then
'   for each objekt in hiscname:objekt.visible=1:next
    HSName1.image = ImgFromCode(HSA1, 1)
    HSName2.image = ImgFromCode(HSA2, 2)
    HSName3.image = ImgFromCode(HSA3, 3)
'   else
'   for each objekt in hiscname:objekt.visible=0:next
' end if
End Sub

Function ImgFromCode(code, digit)
  Dim Image
  if (HighScoreFlashTimer.Enabled = True and hsLetterFlash = 1 and digit = hsCurrentLetter) then
    Image = "postitBL"
  elseif (code + ASC("A") - 1) >= ASC("A") and (code + ASC("A") - 1) <= ASC("Z") then
    Image = "postit" & chr(code + ASC("A") - 1)
  elseif code = 27 Then
    Image = "PostitLT"
    elseif code = 0 Then
    image = "PostitSP"
    Else
      msgbox("Unknown display code: " & code)
  end if
  ImgFromCode = Image
End Function

Sub HighScoreEntryInit()
    playsound "fx_deskbellringing"
  HSA1=0:HSA2=0:HSA3=0
  HSEnterMode = True
  hsCurrentDigit = 0
  hsCurrentLetter = 1:HSA1=1
  HighScoreFlashTimer.Interval = 250
  HighScoreFlashTimer.Enabled = True
  hsLetterFlash = hsFlashDelay
End Sub

Sub HighScoreFlashTimer_Timer()
  hsLetterFlash = hsLetterFlash-1
  UpdatePostIt
  If hsLetterFlash=0 then 'switch back
    hsLetterFlash = hsFlashDelay
  end if
End Sub


' ***********************************************************
'  HiScore ENTER INITIALS
' ***********************************************************

Sub HighScoreProcessKey(keycode)
debug.print keycode
    If keycode = LeftFlipperKey Then
    hsLetterFlash = hsFlashDelay
    Select Case hsCurrentLetter
      Case 1:
        HSA1=HSA1-1:If HSA1=-1 Then HSA1=26 'no backspace on 1st digit
        UpdatePostIt
      Case 2:
        HSA2=HSA2-1:If HSA2=-1 Then HSA2=27
        UpdatePostIt
      Case 3:
        HSA3=HSA3-1:If HSA3=-1 Then HSA3=27
        UpdatePostIt
     End Select
    End If

  If keycode = RightFlipperKey Then
    hsLetterFlash = hsFlashDelay
    Select Case hsCurrentLetter
      Case 1:
        HSA1=HSA1+1:If HSA1>26 Then HSA1=0
        UpdatePostIt
      Case 2:
        HSA2=HSA2+1:If HSA2>27 Then HSA2=0
        UpdatePostIt
      Case 3:
        HSA3=HSA3+1:If HSA3>27 Then HSA3=0
        UpdatePostIt
     End Select
  End If

    If keycode = StartGameKey Then
    Select Case hsCurrentLetter
      Case 1:
                playsound "fx_pencil1"
        hsCurrentLetter=2 'ok to advance
        HSA2=HSA1 'start at same alphabet spot
'       EMReelHSName1.SetValue HSA1:EMReelHSName2.SetValue HSA2
      Case 2:
        If HSA2=27 Then 'bksp
                    playsound "fx_eraser"
          HSA2=0
          hsCurrentLetter=1
        Else
                    playsound "fx_pencil2"
          hsCurrentLetter=3 'enter it
          HSA3=HSA2 'start at same alphabet spot
        End If
      Case 3:
        If HSA3=27 Then 'bksp
                    playsound "fx_eraser"
          HSA3=0
          hsCurrentLetter=2
        Else
                    playsound "fx_pencil3"
                    playsound "fx_simpleding"
          HighScoreFlashTimer.Enabled = False
          HSEnterMode = False
        End If
    End Select
    UpdatePostIt
    End If
End Sub

'*
'*        Williams' Space Mission (1976)
'*        Table scripted by Loserman76
'*
'* 1.0    Dec-07-2021, Loserman76
'*        - original release
'* 1.1    Feb-14-2024, JCalhoun
'*        - Fleep sounds added
'*        - new sound code


option explicit
Randomize
ExecuteGlobal GetTextFile("core.vbs")

On Error Resume Next
ExecuteGlobal GetTextFile("Controller.vbs")
If Err Then MsgBox "Unable To open Controller.vbs. Ensure that it is In the scripts folder."
On Error Goto 0

Const cGameName = "SpaceMission_1976"
Const ShadowFlippersOn = True
Const ShadowBallOn = True
Const ShadowConfigFile = False

Const kDOF_FlipperLeft = 101
Const kDOF_FlipperRight = 102
Const kDOF_BumperBackRight = 105
Const kDOF_BumperBackLeft = 106
Const kDOF_TargetsBase = 107
Const kDOF_BumperMiddleRight = 111
Const kDOF_BumperMiddleLeft = 112
Const kDOF_Shaker = 114
Const kDOF_TriggerBase = 116
Const kDOF_BallRelease = 120
Const kDOF_GateTriggerRelated = 121
Const kDOF_KickerRelated = 126
Const kDOF_Knocker = 128
Const kDOF_StartButton = 130
Const kDOF_ChimeUnitHigh = 141
Const kDOF_ChimeUnitMid = 142
Const kDOF_ChimeUnitLow = 143

Dim Controller  ' B2S
Dim B2SScore  ' B2S Score Displayed

Const HSFileName="SpaceMission_76VPX.txt"
Const B2STableName="SpaceMission_1976"
Const LMEMTableConfig="LMEMTables.txt"
Const LMEMShadowConfig="LMEMShadows.txt"
Dim EnableBallShadow
Dim EnableFlipperShadow

' This value adjusts score motor behavior - 0 allows you To continue scoring while the score motor is running,
' 1 sets score motor to behave more like a real EM
Const ScoreMotorAdjustment = 1

' This is a debug setting to use an older scoring routine vs a newer score routine - don't change this value
Const ScoreAdditionAdjustment = 1

Dim ScoreChecker
Dim CheckAllScores,CheckAllPoints
Dim sortscores(4)
Dim sortpoints(4)
Dim sortplayerpoints(4)
Dim sortplayers(4)
Dim B2SOn   'True/False If want backglass
Dim TextStr,TextStr2
Dim i,xx, LStep, RStep
Dim obj
Dim bgpos
Dim kgpos
Dim dooralreadyopen
Dim kgdooralreadyopen
Dim TargetSpecialLit
Dim Points210counter
Dim Points500counter
Dim Points1000counter
Dim Points2000counter
Dim BallsPerGame
Dim InProgress
Dim BallInPlay
Dim CreditsPerCoin
Dim Score100K(4)
Dim Score(4)
Dim ScoreDisplay(4)
Dim HighScorePaid(4)
Dim HighScore
Dim HighScoreReward
Dim BonusMultiplier
Dim Credits
Dim Match
Dim Replay1
Dim Replay2
Dim Replay3
Dim Replay4
Dim Replay1Paid(4)
Dim Replay2Paid(4)
Dim Replay3Paid(4)
Dim Replay4Paid(4)
Dim BallReplay1
Dim BallReplay2
Dim BallReplay3
Dim ReplayBalls1Paid(4)
Dim ReplayBalls2Paid(4)
Dim ReplayBalls3Paid(4)
Dim TableTilted
Dim TiltCount
Dim debugscore
Dim AltRelayValue
Dim OperatorMenu
Dim BonusBooster
Dim BonusBoosterCounter
Dim BonusCounter
Dim HoleCounter
Dim Ones
Dim Tens
Dim Hundreds
Dim Thousands
Dim Player
Dim Players
Dim rst
Dim bonuscountdown
Dim TempMultiCounter
Dim TempPlayerup
Dim RotatorTemp
Dim GreenTargetsDownCounter
Dim BlueTargetsDownCounter
Dim YellowTargetsDownCounter
Dim bump1
Dim bump2
Dim bump3
Dim bump4
Dim LastChime10
Dim LastChime100
Dim LastChime1000
Dim Score10
Dim Score100
Dim tempbumper
Dim MotorRunning
Dim Replay1Table(35)
Dim Replay2Table(35)
Dim Replay3Table(35)
Dim Replay4Table(35)
Dim ReplayBallsTable1(16)
Dim ReplayBallsTable2(16)
Dim ReplayBallsTable3(16)
Dim ReplayTableSet
Dim ReplayLevel
Dim ReplayBalls
Dim ReplayTableMax
Dim ReplayBallsTableMax
Dim BaseHit, BaseHitCounter, x
Dim LeftTargetFlag, RightTargetFlag
Dim FootballPosition, FootballResetFlag, FootballScore
Dim ImpulseReelCount, ScoreMotorPosition
Dim TempScore, TempAdvance, KickoffFlag, MatchCount
Dim StartGameState, ExtraBallFlag, DoodleBugFlag, AdvanceFlag, SuperAdvanceFlag
Dim KickerCounter, DoodleTargetCount, YellowBumperFlag, GreenBumperFlag, DoodleBGLight, SpecialFlag, SpecialSetting, ScoreMotorStepper, SlingHoldCounter
Dim RightBonusCounter, BonusSide, LeftSpinnerCounter, RightSpinnerCounter, TopTargetCounter, BottomTargetCounter, movetarget, movedirection, movetargetup, movetargetupcount


Sub Table1_Init()
  If Table1.ShowDT = False Then
    For Each obj In DesktopCrap
      obj.visible=False
    Next
  End If

  OperatorMenuBackdrop.image = "PostitBL"
  For XOpt = 1 To MaxOption
    Eval("OperatorOption"&XOpt).image = "PostitBL"
  Next

  For XOpt = 1 To 256
    Eval("Option"&XOpt).image = "PostItBL"
  Next

  LoadEM
  LoadLMEMConfig2
  HideOptions
  SetupReplayTables
  PlasticsOff
  BumpersOff
  OperatorMenu = 0
  StartGameState = 0
  ReplayLevel = 1
  ReplayBalls = 1
  BallsPerGame = 5
  ExtraBallFlag = 0
  EMReel1.ResetToZero
  EMReel2.ResetToZero
  HighScore = 0
  MotorRunning = 0
  BaseHit = 0
  BaseHitCounter = 0
  AdvanceFlag = 1000
  SuperAdvanceFlag = 5
  DoodleBGLight = 1
  HighScoreReward = 1
  BallTextBox.text = ""
  MatchTextBox.text = ""
  SpecialSetting = 1
  Credits = 0
  loadhs
  If HighScore = 0 Then HighScore = 5000

  TableTilted = False
  TiltReel.SetValue(1)

  Match=int(Rnd * 10)
  MatchReel.SetValue((Match) + 1)

  CanPlayReel.SetValue(0)
  BaseRunner1.SetValue(1)
  BaseRunner2.SetValue(1)
  BaseRunner3.SetValue(1)

  OverTheTopReel.SetValue(0)
  RolloverReel1.SetValue(0)
  For Each obj In PlayerHuds
    obj.SetValue(0)
  Next
  For Each obj In PlayerHUDScores
    obj.state = 0
  Next
  If Table1.ShowDT = True Then
    For Each obj In PlayerScores
      obj.visible = 1
    Next
    For Each obj In PlayerScoresOn
      obj.visible = 0
    Next
  End If
  For Each obj In PlayerScores
    obj.ResetToZero
  Next

  KickerCounter = 0

  DoodleTargetCount = 1
  DoodleBugFlag = False
  YellowBumperFlag = False
  GreenBumperFlag = False

  Replay1=Replay1Table(ReplayLevel)
  Replay2=Replay2Table(ReplayLevel)
  Replay3=Replay3Table(ReplayLevel)
  Replay4=Replay4Table(ReplayLevel)
  BallReplay1=ReplayBallsTable1(ReplayBalls)
  BallReplay2=ReplayBallsTable2(ReplayBalls)
  BallReplay3=ReplayBallsTable3(ReplayBalls)

  BonusCounter = 0
  HoleCounter = 0
    bgpos = 6
  kgpos = 0
  FootballPosition= Int(Rnd * 12) + 3
  If B2SOn Then
    For i = 60 To 81
      Controller.B2SSetData i, 0
    Next
    Controller.B2SSetData 60 + FootballPosition + 1, 1
  End If
  FootballReel.setvalue(FootballPosition + 1)
  ImpulseReelCount = Int(Rnd * 4) + 1
  dooralreadyopen = 0
  kgdooralreadyopen = 0

  For i = 0 To 49
    Collection2(i).Visible = 0
    Collection2(i).collidable = False
    Collection2(i).hashitevent = False
  Next
' CenterT50.Visible = 1
  CenterT50.Collidable = True
  CenterT50.Hashitevent = True

  movetarget = 0
  movetargetup = 0
  movetargetupcount = 0
  movedirection = 1
  MovingTargetTimer.enabled = True
  PtargetC.objroty=((movetarget * 0.8) - 20.5) * -1
  RightBonusCounter=1
  RightBonus(RightBonusCounter).state = 1
  InstructCard.image = "IC"
  ReplayCard.image="BC" + FormatNumber(BallsPerGame, 0)

  RefreshReplayCard
  TargetSpecialLit = 0
  Points210counter = 0
  Points500counter = 0
  Points1000counter = 0
  Points2000counter = 0
  BonusBooster = 3
  BonusBoosterCounter = 0
  Players = 0
  RotatorTemp = 1
  InProgress = False

  ScoreText.text=HighScore

  If B2SOn Then
    If Match=0 Then
      Controller.B2SSetMatch 10
    Else
      Controller.B2SSetMatch Match
    End If
    Controller.B2SSetScoreRolloverPlayer1 0
    'Controller.B2SSetScore 4,HighScore
    Controller.B2SSetTilt 1
    Controller.B2SSetCredits Credits
    Controller.B2SSetGameOver 1
  End If
  GameOverReel.setvalue(1)
  CreditsReel.SetValue(Credits)
  If Credits > 0 Then DOF kDOF_StartButton, DOFOn:CreditLight.state = 1
  For i = 1 To 2
      player = i
    If B2SOn Then
      Controller.B2SSetScorePlayer player, 0
    End If
  Next
  bump1 = 1
  bump2 = 1
  bump3 = 1
  bump4 = 1
  InitPauser5.enabled = True
End Sub

Sub Table1_exit()
  savehs
  SaveLMEMConfig
  SaveLMEMConfig2
  If B2SOn Then Controller.Stop
End Sub

Const ReflipAngle = 20

Sub Table1_KeyDown(ByVal keycode)
  ' GNMOD
  If EnteringInitials Then
    CollectInitials(keycode)
    Exit Sub
  End If

  If EnteringOptions Then
    CollectOptions(keycode)
    Exit Sub
  End If

  If keycode = PlungerKey Then
    EMPlayPlungerPullSound Plunger
    Plunger.PullBack
    PlungerPulled = 1
  End If

  If keycode = LeftFlipperKey And InProgress = False Then
    OperatorMenuTimer.Enabled = True
  End If
  ' END GNMOD

  If keycode = LeftFlipperKey And InProgress=True And TableTilted=False Then
    LeftFlipper.RotateToEnd
    If LeftFlipper.currentangle < LeftFlipper.endangle + ReflipAngle Then
      EMPlayLeftFlipperReflipSound LeftFlipper
    Else
      EMPlayLeftFlipperUpAttackSound LeftFlipper
      EMPlayLeftFlipperUpSound LeftFlipper
    End If
    DOF kDOF_FlipperLeft, DOFOn
  End If

  If keycode = RightFlipperKey And InProgress=True And TableTilted=False Then
    RightFlipper.RotateToEnd
    If RightFlipper.currentangle > RightFlipper.endangle - ReflipAngle Then
      EMPlayRightFlipperReflipSound RightFlipper
    Else
      EMPlayRightFlipperUpAttackSound RightFlipper
      EMPlayRightFlipperUpSound RightFlipper
    End If
    DOF kDOF_FlipperRight, DOFON
  End If

  If keycode = LeftTiltKey Then
    Nudge 90, 2
    TiltIt
  End If

  If keycode = RightTiltKey Then
    Nudge 270, 2
    TiltIt
  End If

  If keycode = CenterTiltKey Then
    Nudge 0, 2
    TiltIt
  End If

  If keycode = MechanicalTilt Then
    TiltCount = 2
    TiltIt
  End If

  If keycode = AddCreditKey Or keycode = 4 Then
    If B2SOn Then
      'Controller.B2SSetScorePlayer6 HighScore
    End If
    EMPlayCoinSound
    AddSpecial2
    CreditLight.state=1
  End If

  If keycode = 5 Then
    EMPlayCoinSound
    AddSpecial2
    keycode= StartGameKey
    CreditLight.state=1
  End If

  If keycode = StartGameKey And Credits > 0 And InProgress = True And Players > 0 And Players < 4 And BallInPlay < 2 Then
    Credits = Credits - 1
    If Credits < 1 Then DOF kDOF_StartButton, DOFOff:CreditLight.state = 0
    CreditsReel.SetValue(Credits)
    EMReel7.SetValue(Credits)
    Players=Players + 1
    For Each obj In CanPlayLights
      obj.state = 0
    Next
    CanPlayLights(Players - 1).state = 1

    EMPlayStartButttonSound
    If B2SOn Then
      Controller.B2SSetCanPlay Players
      Controller.B2SSetCredits Credits
    End If
    End If

  If keycode=StartGameKey And Credits>0 And InProgress=False And Players=0 And EnteringOptions = 0 Then
'GNMOD
    OperatorMenuTimer.Enabled = False
'END GNMODthen
    Credits=Credits - 1
    If Credits < 1 Then DOF kDOF_StartButton, 0:CreditLight.state=0
    CreditsReel.SetValue(Credits)
    Players=1

    CanPlay1.state = 1
    MatchReel.SetValue(0)
    Player = 1
    EMPlayStartupSound
    TempPlayerUp=Player
'   PlayerUpRotator.enabled=True
    rst=0
    BallInPlay=1
    InProgress=True
    GameOverReel.setvalue(0)
    resettimer.enabled=True
    RolloverReel1.SetValue(0)
    RolloverReel2.SetValue(0)
    RolloverReel3.SetValue(0)
    RolloverReel4.SetValue(0)
    OverTheTopReel.SetValue(0)
    BonusMultiplier=1
    If B2SOn Then
      Controller.B2SSetTilt 0
      Controller.B2SSetGameOver 0
      Controller.B2SSetMatch 0
      Controller.B2SSetCredits Credits
      'Controller.B2SSetScore 4,HighScore
      Controller.B2SSetCanPlay 1
      Controller.B2SSetPlayerUp 1
'     Controller.B2SSetBallInPlay BallInPlay
      Controller.B2SSetScoreRolloverPlayer1 0
      Controller.B2SSetScoreRolloverPlayer2 0
      Controller.B2SSetScoreRolloverPlayer3 0
      Controller.B2SSetScoreRolloverPlayer4 0
      Controller.B2SSetData 81,1
      Controller.B2SSetData 82,0
      Controller.B2SSetData 83,0
      Controller.B2SSetData 84,0
    End If
    For Each obj In PlayerScores
'     obj.ResetToZero
    Next
    For Each obj In PlayerHuds
      obj.SetValue(0)
    Next
    If Table1.ShowDT = True Then
      For Each obj In PlayerScores
'       obj.ResetToZero
        obj.Visible=True
      Next
      For Each obj In PlayerScoresOn
'       obj.ResetToZero
        obj.Visible=False
      Next
      For Each obj In PlayerHuds
        obj.SetValue(0)
      Next
      For Each obj In PlayerHUDScores
        obj.state=0
      Next
      PlayerHuds(Player-1).SetValue(1)
      PlayerHUDScores(Player-1).state=1
      PlayerHUDScores(Player+3).state=1
      PlayerScores(Player-1).Visible=0
      PlayerScoresOn(Player-1).Visible=1
    End If
  End If
End Sub

Sub Table1_KeyUp(ByVal keycode)
  ' GNMOD
  If EnteringInitials Then
    Exit Sub
  End If

  If keycode = PlungerKey Then
    If PlungerPulled = 0 Then
      Exit Sub
    End If
    EMPlayPlungerReleaseBallSound Plunger
    Plunger.Fire
  End If

  If keycode = LeftFlipperKey Then
    OperatorMenuTimer.Enabled = False
  End If
  ' END GNMOD

  If keycode = LeftFlipperKey And InProgress=True And TableTilted=False Then
    LeftFlipper.RotateToStart
    EMPlayLeftFlipperDownSound LeftFlipper
    DOF kDOF_FlipperLeft,DOFOff
  End If

  If keycode = RightFlipperKey And InProgress=True And TableTilted=False Then
    RightFlipper.RotateToStart
    EMPlayRightFlipperDownSound RightFlipper
    DOF kDOF_FlipperRight, DOFOff
  End If
End Sub

Sub Drain_Hit()
  Drain.DestroyBall
  EMPlayDrainSound Drain
  Pause4Bonustimer.enabled=True
End Sub

Sub Pause4Bonustimer_timer
  Pause4Bonustimer.enabled=0
  AddBonus
  'NextBallDelay.enabled=True
End Sub

Sub CloseGateTrigger_hit
  DOF kDOF_GateTriggerRelated, DOFPulse
End Sub

'***********************
'     Flipper Logos
'***********************

Dim angle

Function PI()
  PI = 4 * Atn(1)
End Function

Sub UpdateFlipperLogos_Timer
' LFlip.ObjRotZ = LeftFlipper.CurrentAngle-90
' RFlip.ObjRotZ = RightFlipper.CurrentAngle+90
' LFlip1.RotZ = LeftFlipper.CurrentAngle
' RFlip1.RotZ = RightFlipper.CurrentAngle
  PGate.Rotz = (Gate.CurrentAngle*.75) + 25
  PGate1.Rotz = (Gate3.CurrentAngle*.75) + 25
  FlipperLSh.RotZ = LeftFlipper.currentangle
  FlipperRSh.RotZ = RightFlipper.currentangle
  Dim SpinnerRadius: SpinnerRadius=7
  SpinnerRod.TransZ = (cos((leftspinner.CurrentAngle + 180) * (PI/180))+1) * SpinnerRadius
  SpinnerRod.TransY = sin((leftspinner.CurrentAngle) * (PI/180)) * -SpinnerRadius
End Sub

'***********************
'     Moving Target
'***********************

Sub MovingTargetTimer_timer
  If movedirection=1 Then
    Collection2(movetarget).visible=0
    Collection2(movetarget).collidable=False
    Collection2(movetarget).hashitevent=False
    movetarget=movetarget + 1
'   Collection2(movetarget).visible=1
    Collection2(movetarget).collidable=True
    Collection2(movetarget).hashitevent=True

    If movetarget > 48 Then
      movedirection=0
    End If
    PtargetC.objroty=((movetarget*.8)-20.5)*-1
    Exit Sub
  Else
    Collection2(movetarget).Visible=0
    Collection2(movetarget).collidable=False
    Collection2(movetarget).hashitevent=False

    movetarget=movetarget-1
'   Collection2(movetarget).visible = 1
    Collection2(movetarget).collidable = True
    Collection2(movetarget).hashitevent = True
    If movetarget < 1 Then
      movedirection = 1
    End If
  End If
  PtargetC.objroty = -((movetarget * 0.8) - 20.5)

  Select Case ExtraBallFlag
    Case 0:
      If RightBonusCounter = 1 Or RightBonusCounter = 4 Or RightBonusCounter=7 Or RightBonusCounter=10 Then
        LeftSpinnerLight.state=1
        If DoubleBonus002.state=1 Or DoubleBonus003.state=1 Or DoubleBonus004.state=1 Then
          UpperLightCenter.state=1
        Else
          UpperLightCenter.state=0
        End If
      Else
        LeftSpinnerLight.state=0
        UpperLightCenter.state=0
      End If
    Case 1:
      If RightBonusCounter=4 Or RightBonusCounter=7 Or RightBonusCounter=10 Then
        LeftSpinnerLight.state=1
        If DoubleBonus002.state=1 Or DoubleBonus003.state=1 Or DoubleBonus004.state=1 Then
          UpperLightCenter.state=1
        Else
          UpperLightCenter.state=0
        End If
      Else
        LeftSpinnerLight.state=0
        UpperLightCenter.state=0
      End If
    Case 2:
      If RightBonusCounter=4 Or RightBonusCounter=10 Then
        LeftSpinnerLight.state=1
        If DoubleBonus002.state=1 Or DoubleBonus003.state=1 Or DoubleBonus004.state=1 Then
          UpperLightCenter.state=1
        Else
          UpperLightCenter.state=0
        End If
      Else
        LeftSpinnerLight.state=0
        UpperLightCenter.state=0
      End If
  End Select
End Sub

'***********************************
' slingshots
'***********************************

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:sling2.rotx = 10:sling1.rotx = 10
        Case 4:sling2.rotx=0:sling1.rotx=0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:sling2.rotx = 10:sling1.rotx = 10
        Case 4:sling2.rotx=0:sling1.rotx=0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

'***********************************
' Walls
'***********************************

Sub RubberWallSwitches_Hit(idx)
  If TableTilted=False Then
    If idx<3 Then
      SetMotor(50)
    Else
      AddScore(10)
    End If
  End If
End Sub

'***********************************
' Spinners
'***********************************

Sub LeftSpinner_Spin
  If TableTilted=False Then
    EMPlaySpinnerSound LeftSpinner
    DOF kDOF_Shaker, DOFPulse
    If LeftSpinnerLight.state=1 Then
      AddScore(1000)
    Else
      AddScore(100)
    End If
  End If
End Sub

'***********************************

Sub Collection2_hit(idx)
  If TableTilted=False Then
    IncreaseBonus
    Select Case RightBonusCounter
      Case 1, 4, 7, 10:
        SetMotor(5000)
      Case 2, 3:
        AddScore(1000)
        EMPlayTargetHitSound
        DoubleBonus002.state=1
        LeftKickerLight.state=1
        RightKickerLight.state=1
        DisplayAltRelay
      Case 5, 6:
        AddScore(1000)
        LeftSlingLight.state=1
        RightSlingLight.state=1
        DoubleBonus003.state=1
        LeftKickerLight.state=1
        RightKickerLight.state=1
        DisplayAltRelay
      Case 8, 9:
        AddScore(1000)
        DoubleBonus004.state=1
        LeftKickerLight.state=1
        RightKickerLight.state=1
        DisplayAltRelay
    End Select
    IncreaseRightBonus
  End If
End Sub

Sub AllTargets_Hit(idx)
  If TableTilted=False Then
    DOF kDOF_TargetsBase + idx, DOFPulse
    IncreaseBonus
    AddScore(AdvanceFlag)
  End If
End Sub

'target001_hit
' target1.transy-5

'***********************************

Sub Bumper1_Hit
  If TableTilted = False Then
    EMPlayTopBumperSound Bumper1
    DOF kDOF_BumperBackLeft, DOFPulse
    bump1 = 1
    AddScore(100)
  End If
End Sub

Sub Bumper2_Hit
  If TableTilted=False Then
    EMPlayTopBumperSound Bumper2
    DOF kDOF_BumperBackRight, DOFPulse
    bump2 = 1
    AddScore(100)
  End If
End Sub

'***********************************

Sub TriggerCollection_Hit(idx)
  If TableTilted=False Then
    EMPlaySensorSound
    DOF kDOF_TriggerBase + idx, DOFPulse
    Select Case idx
      Case 0:
        AddScore(1000)
        If LeftOutlaneLight.state=1 Then AddSpecial
      Case 1:
        If UpperLightLeft.state=1 Then
          Collection2_hit(0)
        Else
          AddScore(100)
        End If
      Case 2:
        If UpperLightRight.state=1 Then
          Collection2_hit(0)
        Else
          AddScore(100)
        End If
      Case 3:
        AddScore(1000)
        If RightOutlaneLight.state=1 Then AddSpecial
      Case 4:
        IncreaseBonus
        AddScore(AdvanceFlag)
        If UpperLightCenter.state=1 Then
          ShootAgainLight.state=1
          ShootAgainReel.Setvalue(1)
          If B2SOn Then
            Controller.B2SSetShootAgain 1
          End If
        End If
      Case 5:
        IncreaseBonus
        AddScore(AdvanceFlag)
      Case 6:
        BonusBoosterCounter=SuperAdvanceFlag
        BonusBoost.enabled=True
    End Select
  End If
End Sub

'*********************************** KICKERS

Dim rkickstep, lkickstep, ktimer, kickFill
Dim kickBalls
Dim kickerBall1

Sub kickBall(kball, kangle, kvel, kvelz, kzlift)
  Dim rangle
  rangle = 3.14 * (kangle - 90) / 180
  kball.z = kball.z + kzlift
  kball.velz = kvelz
  kball.velx = cos(rangle) * kvel
  kball.vely = sin(rangle) * kvel
End Sub

Sub Kicker001_Hit()
  Set kickerBall1 = activeball
  If TableTilted=True Then
    Kicker001.Timerenabled = 1
    Exit Sub
  End If
  EMPlaySaucerLockSound
  KickerHold2.enabled = 1
End Sub

Sub KickerHold2_timer
  If MotorRunning = 0 Then
    SetMotor(500)
    If LeftKickerLight.state = 1 Then
      DoubleBonus001.state = 1
    End If
    KickerHold2.enabled = False
    Kicker001.timerenabled = True
  End If
End Sub

Sub Kicker001_Timer()
  If MotorRunning <> 0 Then Exit Sub
  Kicker001.timerenabled = 0
  If kicker001.ballcntover > 0 Then
    Pkickarm2.rotz=15
    kickBall KickerBall1, 150, 15, 5, 30
    EMPlaySaucerKickSound 1, Kicker001
    DOF kDOF_BumperMiddleRight, DOFPulse
    DOF kDOF_KickerRelated, DOFPulse
    Pkickarm2.rotz=15
    Pkickarm3.rotz=15
    KickerArm2.enabled=True
  End If
End Sub

Sub KickerArm2_timer
  Pkickarm2.rotz = 0
  Pkickarm3.rotz = 0
  KickerArm2.enabled = False
End Sub

Sub kicker002_Hit()
  Set kickerBall1 = activeball
  If TableTilted = True Then
    kicker002.Timerenabled = 1
    Exit Sub
  End If
  EMPlaySaucerLockSound
  KickerHold3.enabled=1
End Sub

Sub KickerHold3_timer
  If MotorRunning = 0 Then
    SetMotor(500)
    If RightKickerLight.state = 1 Then
      DoubleBonus001.state = 1
    End If
    KickerHold3.enabled = False
    kicker002.timerenabled = True
  End If
End Sub

Sub kicker002_Timer()
  If MotorRunning <> 0 Then Exit Sub
  kicker002.timerenabled = 0
  If kicker002.ballcntover > 0 Then
    EMPlaySaucerKickSound 1, Kicker002
    DOF kDOF_BumperMiddleLeft, DOFPulse
    DOF kDOF_KickerRelated, DOFPulse
    kickBall KickerBall1, 210, 15, 5, 30
    Pkickarm2.rotz = 15
    Pkickarm3.rotz = 15
    KickerArm2.enabled = True
  End If
End Sub

Sub LeftSlingKicker_hit
  Set kickerBall1 = activeball
  If TableTilted=True Then
    kickBall KickerBall1, 20, 20, 2, 5
    Exit Sub
  Else
    SlingHoldCounter = 0
    LeftSlingKicker.timerenabled = True
    EMPlaySaucerLockSound
  End If
End Sub

Sub LeftSlingKicker_timer
  If MotorRunning<>0 Then Exit Sub
  SlingHoldCounter=SlingHoldCounter + 1
  Select Case SlingHoldCounter
    Case 1, 2, 3, 4, 5:
      If LeftSlingLight.state=1 Then
        AddScore(100)
      Else
        AddScore(10)
      End If
    Case 7:
    If leftslingkicker.ballcntover > 0 Then
      EMPlaySaucerKickSound 1, LeftSlingKicker
      kickBall KickerBall1, 20, 25, 2, 5
      LeftSlingKicker.timerenabled=False
      sling1.rotx = 20
      sling2.rotx = 20
      LStep = 0
      LeftSlingShot.TimerEnabled = 1
    End If
  End Select
End Sub

Sub RightSlingKicker_hit
  Set kickerBall1 = activeball
  If TableTilted=True Then
    kickBall KickerBall1, 340, 20, 2, 5
    Exit Sub
  Else
    SlingHoldCounter=0
    RightSlingKicker.timerenabled=True
    EMPlaySaucerLockSound
  End If
End Sub

Sub RightSlingKicker_timer
  If MotorRunning<>0 Then Exit Sub
  SlingHoldCounter=SlingHoldCounter + 1
  Select Case SlingHoldCounter
    Case 1,2,3,4,5:
      If RightSlingLight.state=1 Then
        AddScore(100)
      Else
        AddScore(10)
      End If
    Case 7:
    If rightslingkicker.ballcntover > 0 Then
      EMPlaySaucerKickSound 1, RightSlingKicker
      kickBall KickerBall1, 340, 25, 2, 5
      RightSlingKicker.timerenabled=False
      sling1.rotx = 20
      sling2.rotx = 20
      rStep = 0
      RightSlingShot.TimerEnabled = 1
    End If
  End Select
End Sub

'****************************************************

Sub AddSpecial()
  If ExtraBallFlag = 1 Then AddExtraBall:Exit Sub
  If ExtraBallFlag = 2 Then AddExtraBall:Exit Sub
  DOF kDOF_Knocker, DOFPulse
  DOF kDOF_KickerRelated, DOFPulse
  Credits = Credits + 1
  DOF kDOF_StartButton, DOFOn
  If Credits > 15 Then Credits = 15
  If B2SOn Then
    Controller.B2SSetCredits Credits
  End If
  CreditsReel.SetValue(Credits)
End Sub

Sub AddSpecial2()
  EMPlayClickSound
  Credits = Credits + 1
  DOF kDOF_StartButton, DOFOn
  If Credits > 15 Then Credits = 15
  If B2SOn Then
    Controller.B2SSetCredits Credits
  End If
  CreditsReel.SetValue(Credits)
End Sub

Sub AddExtraBall
  DOF kDOF_Knocker, DOFPulse
  DOF kDOF_KickerRelated, DOFPulse
  If ExtraBallFlag = 2 Then AddScore(100000) : Exit Sub
  BallInPlay = BallInPlay + 1
  If BallInPlay > 10 Then BallInPlay = 10
  BallInPlayReel.SetValue(BallInPlay)
  If B2SOn Then
    Controller.B2SSetBallInPlay BallInPlay
  End If
End Sub

Sub openg_timer
  openg.enabled=False
End Sub

Sub closeg_timer
  closeg.enabled=False
End Sub

Sub AddBonus()
  BonusSide=1
  bonuscountdown=bonuscounter
  ScoreMotorStepper=0
  ScoreBonus.enabled=True
End Sub

Sub ToggleOutlane
  '
End Sub

Sub ToggleBumper
  '
End Sub

Sub ToggleAltRelay
  If AltRelayValue=1 Then
    AltRelayValue=0
  Else
    AltRelayValue=1
  End If
  DisplayAltRelay
End Sub

Sub DisplayAltRelay
  If AltRelayValue=0 Then
    UpperLightRight.state=0
    UpperLightLeft.state=0
    LeftOutlaneLight.state=0
    RightOutlaneLight.state=0
    If DoubleBonus002.state=1 AND DoubleBonus003.state=1 AND DoubleBonus004.state=1 Then
      RightOutlaneLight.state=1
    End If
  Else
    UpperLightRight.state=1
    UpperLightLeft.state=1
    LeftOutlaneLight.state=0
    RightOutlaneLight.state=0
    If DoubleBonus002.state=1 AND DoubleBonus003.state=1 AND DoubleBonus004.state=1 Then
      LeftOutlaneLight.state=1
    End If
  End If
End Sub

Sub ResetBallDrops()
  BonusCounter=0
  HoleCounter=0
End Sub

Sub LightsOut
  BonusCounter=0
  HoleCounter=0
End Sub

Sub ResetDrops()
  '
End Sub

Sub LightsOff
  '
End Sub

Sub LightsOn
  '
End Sub

Sub FireAlternatingRelay
  '
End Sub

Sub ToggleArrowRelay
  Select Case BallInPlay
    Case 1,3,5:
    Case 2,4:
  End Select
End Sub

Sub ResetBalls()
  TempMultiCounter=BallsPerGame-BallInPlay
  RightTargetFlag=False
  LeftTargetFlag=False
  BonusMultiplier=1
  YellowBumperFlag=False
  GreenBumperFlag=False
  For Each obj In StarLights
    obj.state=0
  Next
  For x = 1 To 10
    LeftBonus(x).state=0
  Next
  BonusCounter=1
  LeftSpinnerLight.state=0
  LeftSlingLight.state=0
  RightSlingLight.state=0
  LeftKickerLight.state=0
  RightKickerLight.state=0
  LeftBonus(BonusCounter).state=1

  ResetDrops

  TableTilted=False
  TiltReel.SetValue(0)
  BumpersOn
  PlasticsOn
  DisplayAltRelay

  'CreateBallID BallRelease
  Ballrelease.CreateSizedBall 25
  EMPlayBallReleaseSound
    Ballrelease.Kick 40,7
  DOF kDOF_BallRelease, DOFPulse
  'InstructCard.image="IC"+FormatNumber(BallsPerGame,0)+FormatNumber(BallInPlay,0)

  BallInPlayReel.SetValue(BallInPlay)
End Sub

Sub delaykgclose_timer
  delaykgclose.enabled=False
End Sub

Sub resettimer_timer
    rst=rst + 1
  If rst > 1 And rst < 20 Then
    ResetReelsToZero(1)
    ResetReelsToZero(2)
  End If
    If rst = 22 Then
    'playsound "KickerKick"
    End If
    If rst = 24 Then
    newgame
    resettimer.enabled=False
    End If
End Sub

Sub ResetReelsToZero(reelzeroflag)
  Dim d1(5)
  Dim d2(5)
  Dim scorestring1, scorestring2

  If reelzeroflag=1 Then
    scorestring1=CStr(Score(1))
    scorestring2=CStr(Score(2))
    scorestring1=right("00000" & scorestring1,5)
    scorestring2=right("00000" & scorestring2,5)
    For i=0 To 4
      d1(i)=CInt(mid(scorestring1,i+1,1))
      d2(i)=CInt(mid(scorestring2,i+1,1))
    Next
    For i=0 To 4
      If d1(i)>0 Then
        d1(i)=d1(i)+1
        If d1(i)>9 Then d1(i)=0
      End If
      If d2(i)>0 Then
        d2(i)=d2(i)+1
        If d2(i)>9 Then d2(i)=0
      End If
    Next
    Score(1)=(d1(0)*10000) + (d1(1)*1000) + (d1(2)*100) + (d1(3)*10) + d1(4)
    Score(2)=(d2(0)*10000) + (d2(1)*1000) + (d2(2)*100) + (d2(3)*10) + d2(4)
    If B2SOn Then
      Controller.B2SSetScorePlayer 1, Score(1)
      Controller.B2SSetScorePlayer 2, Score(2)
    End If
    PlayerScores(0).SetValue(Score(1))
    PlayerScoresOn(0).SetValue(Score(1))
    PlayerScores(1).SetValue(Score(2))
    PlayerScoresOn(1).SetValue(Score(2))

    scorestring1=CStr(Score(3))
    scorestring2=CStr(Score(4))
    scorestring1=right("00000" & scorestring1, 5)
    scorestring2=right("00000" & scorestring2, 5)
    For i=0 To 4
      d1(i)=CInt(mid(scorestring1, i + 1, 1))
      d2(i)=CInt(mid(scorestring2, i + 1, 1))
    Next
    For i=0 To 4
      If d1(i)>0 Then
        d1(i)=d1(i)+1
        If d1(i)>9 Then d1(i)=0
      End If
      If d2(i)>0 Then
        d2(i)=d2(i)+1
        If d2(i)>9 Then d2(i)=0
      End If
    Next
    Score(3)=(d1(0) * 10000) + (d1(1) * 1000) + (d1(2) * 100) + (d1(3) * 10) + d1(4)
    Score(4)=(d2(0) * 10000) + (d2(1) * 1000) + (d2(2) * 100) + (d2(3) * 10) + d2(4)
    If B2SOn Then
      Controller.B2SSetScorePlayer 3, Score(3)
      Controller.B2SSetScorePlayer 4, Score(4)
    End If
    PlayerScores(2).SetValue(Score(3))
    PlayerScoresOn(2).SetValue(Score(3))
    PlayerScores(3).SetValue(Score(4))
    PlayerScoresOn(3).SetValue(Score(4))
  End If
End Sub

Sub ScoreBonus_timer
  If MotorRunning = 1 Then
    Exit Sub
  End If
  If bonuscountdown < 1 Then
    ScoreBonus.enabled = 0
    NextBallDelay.enabled = True
  Else
    If DoubleBonus001.state = 1 Then
      Select Case ScoreMotorStepper
        Case 0, 1, 3, 4:
          AddScore(1000)
        Case 2, 5:
          EMPlayMotorPauseSound
          If bonuscountdown=20 Then
            LeftBonus(11).state=0
          ElseIf bonuscountdown>10 And bonuscountdown<20 Then
            LeftBonus(bonuscountdown-10).state=0
          Else
            LeftBonus(bonuscountdown).state=0
          End If
          bonuscountdown=bonuscountdown-1
          If bonuscountdown>0 Then
            If bonuscountdown<=10 Then
              LeftBonus(bonuscountdown).state=1
            ElseIf bonuscountdown>10 AND bonuscountdown<20 Then
              LeftBonus(10).state=1
              LeftBonus(bonuscountdown-10).state=1
            Else
              LeftBonus(11).state=1
            End If
          End If
        Case 6:
      End Select
      ScoreMotorStepper=ScoreMotorStepper+1
      If ScoreMotorStepper>6 Then
        ScoreMotorStepper=0
      End If
    Else
      Select Case ScoreMotorStepper
        Case 0,1,2,3,4:
          AddScore(1000)
          If bonuscountdown=20 Then
            LeftBonus(11).state=0
          ElseIf bonuscountdown>10 And bonuscountdown<20 Then
            LeftBonus(bonuscountdown-10).state=0
          Else
            LeftBonus(bonuscountdown).state=0
          End If
          bonuscountdown=bonuscountdown-1
          If bonuscountdown>0 Then
            If bonuscountdown<=10 Then
              LeftBonus(bonuscountdown).state=1
            ElseIf bonuscountdown>10 AND bonuscountdown<20 Then
              LeftBonus(10).state=1
              LeftBonus(bonuscountdown-10).state=1
            Else
              LeftBonus(11).state=1
            End If
          End If
        Case 5:
          EMPlayMotorPauseSound
      End Select
      ScoreMotorStepper=ScoreMotorStepper+1
      If ScoreMotorStepper>5 Then
        ScoreMotorStepper=0
      End If
    End If
  End If
End Sub

Sub AddBonusInGame()
  If BonusSide=1 Then
    bonuscountdown=bonuscounter
  Else
    bonuscountdown=rightbonuscounter
  End If
  ScoreMotorStepper=0
  ScoreBonusInGame.enabled=True
End Sub

Sub ScoreBonusInGame_timer
  If MotorRunning=1 Then
    Exit Sub
  End If
  If bonuscountdown<1 Then
    ScoreBonusInGame.enabled=0
  Else
    If DoubleBonus001.state=0 Then
      Select Case ScoreMotorStepper
        Case 0,1,2,3,4:
          AddScore(1000)
        Case 5:
          EMPlayMotorPauseSound
          If BonusSide=1 Then
            LeftBonus(bonuscountdown).state=0
            LeftSpinnerLight.state=0
          Else
            RightBonus(bonuscountdown).state=0
            RightSpinnerLight.state=0
          End If

          bonuscountdown=bonuscountdown-1
          If bonuscountdown>0 Then
            If BonusSide=1 Then
              LeftBonus(bonuscountdown).state=1
            Else
              RightBonus(bonuscountdown).state=1
            End If

          End If
        Case 6:
      End Select
      ScoreMotorStepper=ScoreMotorStepper+1
      If ScoreMotorStepper>6 Then
        ScoreMotorStepper=0
      End If
    Else
      Select Case ScoreMotorStepper
        Case 0,1,2,3,4:
          AddScore(10000)
          If BonusSide=1 Then
            LeftBonus(bonuscountdown).state=0
            LeftSpinnerLight.state=0
          Else
            RightBonus(bonuscountdown).state=0
            RightSpinnerLight.state=0
          End If
          bonuscountdown=bonuscountdown-1
          If bonuscountdown>0 Then
            If BonusSide=1 Then
              LeftBonus(bonuscountdown).state=1
            Else
              RightBonus(bonuscountdown).state=1
            End If
          End If
        Case 5:
          EMPlayMotorPauseSound
      End Select
      ScoreMotorStepper=ScoreMotorStepper+1
      If ScoreMotorStepper>5 Then
        ScoreMotorStepper=0
      End If
    End If
  End If
End Sub

Sub NextBallDelay_timer()
  If KickoffFlag=True Then Exit Sub
  NextBallDelay.enabled=False
  nextball
End Sub

Sub NewGameHoldTimer_timer
  If KickoffFlag=True Then Exit Sub
  NewGameHoldTimer.enabled=False
  newgame
End Sub

Sub newgame
  InProgress=True
  debugscore=0
  EMReel1.ResetToZero
  EMReel2.ResetToZero
  MatchTextBox.text=""
  'BallTextBox.text=FormatNumber(BallInPlay,0)
  For i = 1 To 2
    Score(i)=0
    Score100K(1)=0
    HighScorePaid(i)=False
    Replay1Paid(i)=False
    Replay2Paid(i)=False
    Replay3Paid(i)=False
    Replay4Paid(i)=False
    ReplayBalls1Paid(i)=False
    ReplayBalls2Paid(i)=False
    ReplayBalls3Paid(i)=False
  Next
  If B2SOn Then
    Controller.B2SSetTilt 0
    Controller.B2SSetGameOver 0
    Controller.B2SSetMatch 0
    Controller.B2SSetScorePlayer1 0
    Controller.B2SSetScorePlayer2 0
'   Controller.B2SSetScorePlayer3 0
'   Controller.B2SSetScorePlayer4 0
    Controller.B2SSetBallInPlay BallInPlay
  End If

  AltRelayValue=0

  BaseHit=0
  BaseHitCounter=0
  BonusCounter=0
  For Each obj In LeftBonus
    obj.state=0
  Next
  For Each obj In LeftSpinnerLights
    obj.state=0
  Next
  For Each obj In RightSpinnerLights
    obj.state=0
  Next
  For Each obj In StarLights
    obj.state=0
  Next
  For Each obj In TargetAdvanceLights
    obj.state=0
  Next

  LeftOutlaneLight.state=0
  RightOutlaneLight.state=0
  LeftSpinnerLight.state=0

  ToggleArrowRelay
  ResetBalls
  EMPlayStartBallSound 0
End Sub

Sub nextball
' If B2SOn Then
'   Controller.B2SSetTilt 0
' End If
  If ShootAgainLight.state = 1 Then
    ShootAgainLight.state = 0
    ShootAgainReel.Setvalue(0)
    If B2SOn Then Controller.B2SSetShootAgain 0
  Else
    Player = Player + 1
  End If
  If Player > Players Then
    BallInPlay = BallInPlay + 1
    If BallInPlay > 5 Then
      EMPlayMotorLeerSound
      InProgress = False
      If B2SOn Then
        Controller.B2SSetGameOver 1
        Controller.B2SSetPlayerUp 0
        Controller.B2SSetBallInPlay 0
        Controller.B2SSetCanPlay 0
        Controller.B2SSetData 81, 0
        Controller.B2SSetData 82, 0
        Controller.B2SSetData 83, 0
        Controller.B2SSetData 84, 0
      End If
      For Each obj In PlayerHuds
        obj.SetValue(0)
      Next
      For Each obj In PlayerHUDScores
        obj.state = 0
      Next
      If Table1.ShowDT = True Then
        For Each obj In PlayerScores
          obj.visible = 1
        Next
        For Each obj In PlayerScoresOn
          obj.visible = 0
        Next
      End If
      'InstructCard.image="IC"

      BallInPlayReel.SetValue(0)
      CanPlayReel.SetValue(0)
      GameOverReel.setvalue(1)
      LeftFlipper.RotateToStart
      RightFlipper.RotateToStart
      LightsOut
      PlasticsOff
      BallTextBox.text = ""
      If TableTilted = False Then
        checkmatch
      End If
      CheckHighScore
      Players=0
      HighScoreTimer.interval = 100
      HighScoreTimer.enabled = True
    Else
      Player = 1
      If B2SOn Then
        Controller.B2SSetPlayerUp Player
        Controller.B2SSetData 81, 0
        Controller.B2SSetData 82, 0
        Controller.B2SSetData 83, 0
        Controller.B2SSetData 84, 0
        Controller.B2SSetData 80 + Player, 1
        Controller.B2SSetBallInPlay BallInPlay
      End If
      EMPlayRotateThroughPlayersSound
      TempPlayerUp = Player
      PlayerUpRotator.enabled = True
      PlayStartBall.enabled = True
      For Each obj In PlayerHuds
        obj.SetValue(0)
      Next
      For Each obj In PlayerHUDScores
        obj.state=0
      Next
      If Table1.ShowDT = True Then
        For Each obj In PlayerScores
          obj.visible=1
        Next
        For Each obj In PlayerScoresOn
          obj.visible=0
        Next
        PlayerHuds(Player-1).SetValue(1)
        PlayerHUDScores(Player-1).state=1
        PlayerHUDScores(Player+3).state=1
        PlayerScores(Player-1).visible=0
        PlayerScoresOn(Player-1).visible=1
'       PlayerHuds(Player).SetValue(1)
'       PlayerHUDScores(Player).state=1
'       PlayerScores(Player).visible=0
'       PlayerScoresOn(Player).visible=1
      End If
      ToggleArrowRelay
      ResetBalls
    End If
  Else
    If B2SOn Then
      Controller.B2SSetPlayerUp Player
      Controller.B2SSetData 81,0
      Controller.B2SSetData 82,0
      Controller.B2SSetData 83,0
      Controller.B2SSetData 84,0
      Controller.B2SSetData 80+Player,1
      Controller.B2SSetBallInPlay BallInPlay
    End If
    EMPlayRotateThroughPlayersSound
    TempPlayerUp=Player
'   PlayerUpRotator.enabled=True
    PlayStartBall.enabled=True
    For Each obj In PlayerHuds
      obj.SetValue(0)
    Next
    For Each obj In PlayerHUDScores
      obj.state=0
    Next
    If Table1.ShowDT = True Then
      For Each obj In PlayerScores
          obj.visible=1
      Next
      For Each obj In PlayerScoresOn
          obj.visible=0
      Next
      PlayerHuds(Player-1).SetValue(1)
      PlayerHUDScores(Player-1).state=1
      PlayerHUDScores(Player+3).state=1
      PlayerScores(Player-1).visible=0
      PlayerScoresOn(Player-1).visible=1
'     PlayerHuds(Player).SetValue(1)
'     PlayerHUDScores(Player).state=1
'     PlayerScores(Player).visible=0
'     PlayerScoresOn(Player).visible=1
    End If
    ResetBalls
  End If
End Sub

Sub CheckHighScore
  Dim playertops
  Dim si
  Dim sj
  Dim stemp
  Dim stempplayers
  For i=1 To 4
    sortscores(i)=0
    sortplayers(i)=0
    sortpoints(i)=0
    sortplayerpoints(i)=0
  Next
  playertops=0
  For i = 1 To Players
    sortscores(i)=Score(i)
    sortplayers(i)=i
  Next
  ScoreChecker=5
  CheckAllScores=1
  CheckAllPoints=0
  NewHighScore sortscores(ScoreChecker-1),sortplayers(ScoreChecker-1)
  savehs
End Sub

Sub checkmatch
  Dim tempmatch
  If ExtraBallFlag>0 Then Exit Sub
  tempmatch=Int(Rnd*10)
  Match=tempmatch
  MatchReel.SetValue(tempmatch+1)
  If Match=0 Then
    MatchTextBox.text="Match: 0"
  Else
    MatchTextBox.text="Match: " + FormatNumber(Match,0)
  End If

  If B2SOn Then
    If Match = 0 Then
      Controller.B2SSetMatch 100
    Else
      Controller.B2SSetMatch Match*10
    End If
  End If
  Match=Match * 10
  For i = 1 To Players
    If Match=(Score(i) mod 100) Then
      AddSpecial
    End If
  Next
End Sub

Sub TiltTimer_Timer()
  If TiltCount > 0 Then TiltCount = TiltCount - 1
  If TiltCount = 0 Then
    TiltTimer.Enabled = False
  End If
End Sub

Sub TiltIt()
    TiltCount = TiltCount + 1
    If TiltCount = 3 Then
      TableTilted=True
      PlasticsOff
      BumpersOff
      LeftFlipper.RotateToStart
      RightFlipper.RotateToStart
      TiltReel.SetValue(1)
      If B2Son Then
        Controller.B2SSetTilt 1
      End If
    Else
      TiltTimer.Interval = 500
      TiltTimer.Enabled = True
    End If
End Sub

Sub IncreaseBonus()
  If BonusCounter>=10 And BonusCounter<20 Then
    LeftBonus(BonusCounter-10).state=0
    BonusCounter=BonusCounter+1
    If BonusCounter<20 Then
      LeftBonus(10).state=1
      LeftBonus(BonusCounter-10).state=1
    Else
      LeftBonus(10).state=0
      LeftBonus(11).state=1
    End If
  ElseIf BonusCounter=20 Then
    Exit Sub
  Else
    If BonusCounter>0 Then
      LeftBonus(BonusCounter).state=0
    End If
    BonusCounter=BonusCounter+1
    LeftBonus(BonusCounter).State=1
'   PlaySound("Score100")
  End If
End Sub

Sub IncreaseRightBonus()
  If RightBonusCounter=10 Then
    RightBonusCounter=1
    RightBonus(RightBonusCounter).state=1
    RightBonus(10).state=0
  Else
    If RightBonusCounter>0 Then
      RightBonus(RightBonusCounter).state=0
    End If
    RightBonusCounter=RightBonusCounter+1
    RightBonus(RightBonusCounter).State=1
    'PlaySound("Score100")
  End If
End Sub

Sub BonusBoost_Timer()
  IncreaseBonus
  AddScore(AdvanceFlag)
  BonusBoosterCounter=BonusBoosterCounter - 1
  If BonusBoosterCounter < 1 Then
    BonusBoost.enabled = False
  End If
End Sub

Sub CheckForLightSpecial()
  '
End Sub

Sub PlayStartBall_timer()
  PlayStartBall.enabled = False
  EMPlayStartBallSound 1
End Sub

Sub PlayerUpRotator_timer()
  Exit Sub
  If RotatorTemp < 5 Then
    TempPlayerUp = TempPlayerUp + 1
    If TempPlayerUp > 4 Then
      TempPlayerUp = 1
    End If
    For Each obj In PlayerHuds
      obj.SetValue(0)
    Next
    For Each obj In PlayerHUDScores
      obj.state = 0
    Next
    If Table1.ShowDT = True Then
      For Each obj In PlayerScores
        obj.visible = 1
      Next
      For Each obj In PlayerScoresOn
        obj.visible=0
      Next
      PlayerHuds(TempPlayerUp-1).SetValue(1)
      PlayerScores(TempPlayerUp-1).visible=0
      PlayerScoresOn(TempPlayerUp-1).visible=1
    End If
    If B2SOn Then
      Controller.B2SSetPlayerUp TempPlayerUp
      Controller.B2SSetData 81,0
      Controller.B2SSetData 82,0
      Controller.B2SSetData 83,0
      Controller.B2SSetData 84,0
      Controller.B2SSetData 80+TempPlayerUp,1
    End If
  Else
    If B2SOn Then
      Controller.B2SSetPlayerUp Player
      Controller.B2SSetData 81,0
      Controller.B2SSetData 82,0
      Controller.B2SSetData 83,0
      Controller.B2SSetData 84,0
      Controller.B2SSetData 80+Player,1
    End If
    PlayerUpRotator.enabled=False
    RotatorTemp=1
    For Each obj In PlayerHuds
      obj.SetValue(0)
    Next
    For Each obj In PlayerHUDScores
      obj.state=0
    Next
    If Table1.ShowDT = True Then
      For Each obj In PlayerScores
        obj.visible=1
      Next
      For Each obj In PlayerScoresOn
        obj.visible=0
      Next
      PlayerHuds(Player-1).SetValue(1)

      PlayerScores(Player-1).visible=0
      PlayerScoresOn(Player-1).visible=1
    End If
  End If
  RotatorTemp=RotatorTemp+1
End Sub

Sub savehs
  ' Based on Black's Highscore routines
  Dim FileObj
  Dim ScoreFile
  Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) Then
    Exit Sub
  End If
  Set ScoreFile=FileObj.CreateTextFile(UserDirectory & HSFileName,True)
    ScoreFile.WriteLine 2
    ScoreFile.WriteLine Credits
    scorefile.writeline BallsPerGame
    ScoreFile.WriteLine ReplayLevel
    ScoreFile.WriteLine SuperAdvanceFlag
    ScoreFile.WriteLine ExtraBallFlag
    ScoreFile.WriteLine AdvanceFlag
    For xx=1 To 5
      scorefile.writeline HSScore(xx)
    Next
    For xx=1 To 5
      scorefile.writeline HSName(xx)
    Next
    For xx=6 To 10
      scorefile.writeline HSScore(xx)
    Next
    For xx=6 To 10
      scorefile.writeline HSName(xx)
    Next
    ScoreFile.Close
  Set ScoreFile=Nothing
  Set FileObj=Nothing
End Sub

Sub loadhs
    ' Based on Black's Highscore routines
  Dim FileObj
  Dim ScoreFile
    Dim temp1
    Dim temp2
  Dim temp3
  Dim temp4
  Dim temp5
  Dim temp6
  Dim temp7
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
  Dim temp18
  Dim temp19
  Dim temp20
  Dim temp21
  Dim temp22
  Dim temp23
  Dim temp24
  Dim temp25
  Dim temp26
  Dim temp27

    Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) Then
    Exit Sub
  End If
  If Not FileObj.FileExists(UserDirectory & HSFileName) Then
    Exit Sub
  End If
  Set ScoreFile=FileObj.GetFile(UserDirectory & HSFileName)
  Set TextStr=ScoreFile.OpenAsTextStream(1,0)
  If (TextStr.AtEndOfStream=True) Then
    Exit Sub
  End If
  temp1=TextStr.ReadLine
  temp2=textstr.readline
  temp3=textstr.readline
  temp4=textstr.readline
  temp5=textstr.readline
  temp6=textstr.readline
  temp7=textstr.readline
  HighScore=cdbl(temp1)
  If HighScore=2 Then
    temp8=textstr.readline
    temp9=textstr.readline
    temp10=textstr.readline
    temp11=textstr.readline
    temp12=textstr.readline
    temp13=textstr.readline
    temp14=textstr.readline
    temp15=textstr.readline
    temp16=textstr.readline
    temp17=textstr.readline
    temp18=textstr.readline
    temp19=textstr.readline
    temp20=textstr.readline
    temp21=textstr.readline
    temp22=textstr.readline
    temp23=textstr.readline
    temp24=textstr.readline
    temp25=textstr.readline
    temp26=textstr.readline
    temp27=textstr.readline
  End If
  TextStr.Close
  If HighScore=2 Then
    Credits=cdbl(temp2)
    BallsPerGame=cdbl(temp3)
    ReplayLevel=cdbl(temp4)
    SuperAdvanceFlag=int(temp5)
    ExtraBallFlag=cdbl(temp6)
    AdvanceFlag= int(temp7)
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

    HSScore(6) = int(temp18)
    HSScore(7) = int(temp19)
    HSScore(8) = int(temp20)
    HSScore(9) = int(temp21)
    HSScore(10) = int(temp22)

    HSName(6) = temp23
    HSName(7) = temp24
    HSName(8) = temp25
    HSName(9) = temp26
    HSName(10) = temp27
  End If
  Set ScoreFile=Nothing
  Set FileObj=Nothing
End Sub

Sub SaveLMEMConfig
  Dim FileObj
  Dim LMConfig
  Dim temp1
  Dim tempb2s
  tempb2s=0
  If B2SOn=True Then
    tempb2s=1
  Else
    tempb2s=0
  End If
  Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) Then
    Exit Sub
  End If
  Set LMConfig=FileObj.CreateTextFile(UserDirectory & LMEMTableConfig,True)
  LMConfig.WriteLine tempb2s
  LMConfig.Close
  Set LMConfig=Nothing
  Set FileObj=Nothing
End Sub

Sub LoadLMEMConfig
  Dim FileObj
  Dim LMConfig
  Dim tempC
  Dim tempb2s

    Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) Then
    Exit Sub
  End If
  If Not FileObj.FileExists(UserDirectory & LMEMTableConfig) Then
    Exit Sub
  End If
  Set LMConfig=FileObj.GetFile(UserDirectory & LMEMTableConfig)
  Set TextStr2=LMConfig.OpenAsTextStream(1,0)
  If (TextStr2.AtEndOfStream=True) Then
    Exit Sub
  End If
  tempC=TextStr2.ReadLine
  TextStr2.Close
  tempb2s=cdbl(tempC)
  If tempb2s=0 Then
    B2SOn=False
  Else
    B2SOn=True
  End If
  Set LMConfig=Nothing
  Set FileObj=Nothing
End Sub

Sub SaveLMEMConfig2
  If ShadowConfigFile=False Then Exit Sub
  Dim FileObj
  Dim LMConfig2
  Dim temp1
  Dim temp2
  Dim tempBS
  Dim tempFS

  If EnableBallShadow=True Then
    tempBS=1
  Else
    tempBS=0
  End If
  If EnableFlipperShadow=True Then
    tempFS=1
  Else
    tempFS=0
  End If

  Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) Then
    Exit Sub
  End If
  Set LMConfig2=FileObj.CreateTextFile(UserDirectory & LMEMShadowConfig,True)
  LMConfig2.WriteLine tempBS
  LMConfig2.WriteLine tempFS
  LMConfig2.Close
  Set LMConfig2=Nothing
  Set FileObj=Nothing
End Sub

Sub LoadLMEMConfig2
  If ShadowConfigFile=False Then
    EnableBallShadow = ShadowBallOn
    BallShadowUpdate.enabled = ShadowBallOn
    EnableFlipperShadow = ShadowFlippersOn
    FlipperLSh.visible = ShadowFlippersOn
    FlipperRSh.visible = ShadowFlippersOn
    Exit Sub
  End If
  Dim FileObj
  Dim LMConfig2
  Dim tempC
  Dim tempD
  Dim tempFS
  Dim tempBS

    Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) Then
    Exit Sub
  End If
  If Not FileObj.FileExists(UserDirectory & LMEMShadowConfig) Then
    Exit Sub
  End If
  Set LMConfig2=FileObj.GetFile(UserDirectory & LMEMShadowConfig)
  Set TextStr2=LMConfig2.OpenAsTextStream(1,0)
  If (TextStr2.AtEndOfStream=True) Then
    Exit Sub
  End If
  tempC=TextStr2.ReadLine
  tempD=TextStr2.Readline
  TextStr2.Close
  tempBS=cdbl(tempC)
  tempFS=cdbl(tempD)
  If tempBS=0 Then
    EnableBallShadow=False
    BallShadowUpdate.enabled=False
  Else
    EnableBallShadow=True
  End If
  If tempFS=0 Then
    EnableFlipperShadow=False
    FlipperLSh.visible=False
    FLipperRSh.visible=False
  Else
    EnableFlipperShadow=True
  End If
  Set LMConfig2=Nothing
  Set FileObj=Nothing
End Sub

Sub DisplayHighScore
  '
End Sub

Sub InitPauser5_timer
  If B2SOn Then
    'Controller.B2SSetScore 4,HighScore
  End If
  DisplayHighScore
  InitPauser5.enabled=False
End Sub

Sub target001_hit
  target1.transy=-4
  target001.timerenabled=1
End Sub

Sub target002_hit
  target2.transy=-4
  target001.timerenabled=1
End Sub

Sub target001_timer
  target1.transy=0
  target2.transy=0
  target001.timerenabled=0
End Sub

Sub BumpersOff
' Bumper1Light.visible=0
' Bumper2Light.visible=0
End Sub

Sub BumpersOn
' Bumper1Light.visible=1
' Bumper2Light.visible=1
End Sub

Sub PlasticsOn
  For Each xx In flashers:xx.State = 1: Next
  playfield_off.visible=0
  For Each xx In layercol1: xx.image = "layer1": Next
  For Each xx In layercol2: xx.image = "layer2": Next
  For Each xx In layercol3: xx.image = "layer3": Next
  For Each xx In layercol4: xx.image = "layer4": Next
  backmetal.image="backmetal"
  toparch.image="archmission"
' toparch.image="archodyssey" 'For odyssey
  bumpcaps.image="bumpcapsmission"
' bumpcaps.image="bumpcapsodyssey" 'For odyssey
End Sub

Sub PlasticsOff
  Dim xx
  For Each xx In flashers:xx.State = 0: Next
  EMStopBuzzSounds
  playfield_off.visible=1
  For Each xx In layercol1: xx.image = "layer1off": Next
  For Each xx In layercol2: xx.image = "layer2off": Next
  For Each xx In layercol3: xx.image = "layer3off": Next
  For Each xx In layercol4: xx.image = "layer4off": Next
  backmetal.image="backmetaloff"
  toparch.image="archmissionoff"
' toparch.image="archodysseyoff"  'For odyssey
  bumpcaps.image="bumpcapsmissionoff"
' bumpcaps.image="bumpcapsodysseyoff" 'For odyssey
End Sub

Sub SetupReplayTables
  Replay1Table(1)=114000
  Replay1Table(2)=116000
  Replay1Table(3)=118000
  Replay1Table(4)=120000
  Replay1Table(5)=122000
  Replay1Table(6)=124000
  Replay1Table(7)=126000
  Replay1Table(8)=128000
  Replay1Table(9)=130000
  Replay1Table(10)=132000
  Replay1Table(11)=134000
  Replay1Table(12)=136000
  Replay1Table(13)=138000
  Replay1Table(14)=140000
  Replay1Table(15)=142000
  Replay1Table(16)=130000
  Replay1Table(17)=132000
  Replay1Table(18)=134000
  Replay1Table(19)=136000
  Replay1Table(20)=138000
  Replay1Table(21)=140000
  Replay1Table(22)=142000
  Replay1Table(23)=144000
  Replay1Table(24)=146000
  Replay1Table(25)=148000
  Replay1Table(26)=150000
  Replay1Table(27)=152000
  Replay1Table(28)=154000
  Replay1Table(29)=156000
  Replay1Table(30)=158000
  Replay1Table(31)=160000
  Replay1Table(32)=162000
  Replay1Table(33)=164000
  Replay1Table(34)=166000
  Replay1Table(35)=168000

  Replay2Table(1)=145000
  Replay2Table(2)=147000
  Replay2Table(3)=149000
  Replay2Table(4)=151000
  Replay2Table(5)=153000
  Replay2Table(6)=155000
  Replay2Table(7)=157000
  Replay2Table(8)=159000
  Replay2Table(9)=161000
  Replay2Table(10)=163000
  Replay2Table(11)=165000
  Replay2Table(12)=167000
  Replay2Table(13)=169000
  Replay2Table(14)=171000
  Replay2Table(15)=173000
  Replay2Table(16)=174000
  Replay2Table(17)=176000
  Replay2Table(18)=178000
  Replay2Table(19)=180000
  Replay2Table(20)=182000
  Replay2Table(21)=184000
  Replay2Table(22)=186000
  Replay2Table(23)=188000
  Replay2Table(24)=190000
  Replay2Table(25)=192000
  Replay2Table(26)=194000
  Replay2Table(27)=196000
  Replay2Table(28)=198000
  Replay2Table(29)=198000
  Replay2Table(30)=197000
  Replay2Table(31)=198000
  Replay2Table(32)=198000
  Replay2Table(33)=198000
  Replay2Table(34)=198000
  Replay2Table(35)=197000

  Replay3Table(1)=176000
  Replay3Table(2)=178000
  Replay3Table(3)=180000
  Replay3Table(4)=182000
  Replay3Table(5)=184000
  Replay3Table(6)=186000
  Replay3Table(7)=188000
  Replay3Table(8)=190000
  Replay3Table(9)=192000
  Replay3Table(10)=194000
  Replay3Table(11)=196000
  Replay3Table(12)=198000
  Replay3Table(13)=197000
  Replay3Table(14)=198000
  Replay3Table(15)=198000
  Replay3Table(16)=999000
  Replay3Table(17)=999000
  Replay3Table(18)=999000
  Replay3Table(19)=999000
  Replay3Table(20)=999000
  Replay3Table(21)=999000
  Replay3Table(22)=999000
  Replay3Table(23)=999000
  Replay3Table(24)=999000
  Replay3Table(25)=999000
  Replay3Table(26)=999000
  Replay3Table(27)=999000
  Replay3Table(28)=999000
  Replay3Table(29)=999000
  Replay3Table(30)=999000
  Replay3Table(31)=999000
  Replay3Table(32)=999000
  Replay3Table(33)=999000
  Replay3Table(34)=999000
  Replay3Table(35)=999000

  Replay4Table(1)=999000
  Replay4Table(2)=999000
  Replay4Table(3)=999000
  Replay4Table(4)=999000
  Replay4Table(5)=999000
  Replay4Table(6)=999000
  Replay4Table(7)=999000
  Replay4Table(8)=999000
  Replay4Table(9)=999000
  Replay4Table(10)=999000
  Replay4Table(11)=999000
  Replay4Table(12)=999000
  Replay4Table(13)=999000
  Replay4Table(14)=999000
  Replay4Table(15)=999000
  Replay4Table(16)=999000
  Replay4Table(17)=999000
  Replay4Table(18)=999000
  Replay4Table(19)=999000
  Replay4Table(20)=999000
  Replay4Table(21)=999000
  Replay4Table(22)=999000
  Replay4Table(23)=999000
  Replay4Table(24)=999000
  Replay4Table(25)=999000
  Replay4Table(26)=999000
  Replay4Table(27)=999000
  Replay4Table(28)=999000
  Replay4Table(29)=999000
  Replay4Table(30)=999000
  Replay4Table(31)=999000
  Replay4Table(32)=999000
  Replay4Table(33)=999000
  Replay4Table(34)=999000
  Replay4Table(35)=999000

  ReplayTableMax=35

  ReplayBallsTable1(1)=99
  ReplayBallsTable1(2)=99
  ReplayBallsTable1(3)=11
  ReplayBallsTable1(4)=11
  ReplayBallsTable1(5)=12
  ReplayBallsTable1(6)=16
  ReplayBallsTable1(7)=17
  ReplayBallsTable1(8)=18
  ReplayBallsTable1(9)=19
  ReplayBallsTable1(10)=19
  ReplayBallsTable1(11)=99
  ReplayBallsTable1(12)=99
  ReplayBallsTable1(13)=99
  ReplayBallsTable1(14)=99
  ReplayBallsTable1(15)=99
  ReplayBallsTable1(16)=99

  ReplayBallsTable2(1)=99
  ReplayBallsTable2(2)=99
  ReplayBallsTable2(3)=15
  ReplayBallsTable2(4)=16
  ReplayBallsTable2(5)=15
  ReplayBallsTable2(6)=22
  ReplayBallsTable2(7)=22
  ReplayBallsTable2(8)=23
  ReplayBallsTable2(9)=23
  ReplayBallsTable2(10)=24
  ReplayBallsTable2(11)=99
  ReplayBallsTable2(12)=99
  ReplayBallsTable2(13)=99
  ReplayBallsTable2(14)=99
  ReplayBallsTable2(15)=99
  ReplayBallsTable2(16)=99

  ReplayBallsTable3(1)=99
  ReplayBallsTable3(2)=99
  ReplayBallsTable3(3)=17
  ReplayBallsTable3(4)=18
  ReplayBallsTable3(5)=17
  ReplayBallsTable3(6)=28
  ReplayBallsTable3(7)=28
  ReplayBallsTable3(8)=29
  ReplayBallsTable3(9)=28
  ReplayBallsTable3(10)=30
  ReplayBallsTable3(11)=99
  ReplayBallsTable3(12)=99
  ReplayBallsTable3(13)=99
  ReplayBallsTable3(14)=99
  ReplayBallsTable3(15)=99
  ReplayBallsTable3(16)=99

  ReplayBallsTableMax=2
End Sub

Sub RefreshReplayCard
  Dim tempst1
  Dim tempst2
  Dim tempst3

  tempst1=FormatNumber(BallsPerGame,0)
  tempst2=FormatNumber(ReplayLevel,0)
  tempst3=FormatNumber(ReplayBalls,0)

  ReplayCard.image="BC" + tempst1
  ReplayCard1.image = "SC" + tempst2
  ReplayCard2.image = "PR"

  Replay1=Replay1Table(ReplayLevel)
  Replay2=Replay2Table(ReplayLevel)
  Replay3=Replay3Table(ReplayLevel)
  Replay4=Replay4Table(ReplayLevel)
  BallReplay1=ReplayBallsTable1(ReplayBalls)
  BallReplay2=ReplayBallsTable2(ReplayBalls)
  BallReplay3=ReplayBallsTable3(ReplayBalls)
End Sub

'****************************************
'  SCORE MOTOR
'****************************************


ScoreMotorTimer.Enabled = 1
ScoreMotorTimer.Interval = 135 '135
AddScoreTimer.Enabled = 1
AddScoreTimer.Interval = 135

Dim queuedscore
Dim MotorMode
Dim MotorPosition

Sub SetMotor(y)
  Select Case ScoreMotorAdjustment
    Case 0:
      queuedscore=queuedscore+y
    Case 1:
      If MotorRunning<>1 And InProgress=True Then
        queuedscore=queuedscore+y
      End If
  End Select
End Sub

Sub SetMotor2(x)
  If MotorRunning<>1 And InProgress=True Then
    MotorRunning=1
    BaseHitCounter=BaseHit
    LightsOff
    Select Case x
      Case 10:
        AddScore(10)
        MotorRunning=0
        LightsOn
      Case 20:
        MotorMode=10
        MotorPosition=2
      Case 30:
        MotorMode=10
        MotorPosition=3
      Case 40:
        MotorMode=10
        MotorPosition=4
      Case 50:
        MotorMode=10
        ToggleAltRelay
        MotorPosition=5
      Case 100:
        MotorMode=100
        MotorPosition=1
      Case 200:
        MotorMode=100
        MotorPosition=2
      Case 300:
        MotorMode=100
        MotorPosition=3
      Case 400:
        MotorMode=100
        MotorPosition=4
      Case 500:
        MotorMode=100
        ToggleAltRelay
        MotorPosition=5
      Case 1000:
        AddScore(1000)
        MotorRunning=0
        LightsOn
      Case 2000:
        MotorMode=1000
        MotorPosition=2
      Case 3000:
        MotorMode=1000
        MotorPosition=3
      Case 4000:
        MotorMode=1000
        MotorPosition=4
      Case 5000:
        MotorMode=1000
        ToggleAltRelay
        MotorPosition=5
      Case 10000:
        AddScore(10000)
        MotorRunning=0
        LightsOn
      Case 50000:
        MotorMode=10000
        MotorPosition=5
    End Select
  End If
End Sub

Sub AddScoreTimer_Timer
  Dim tempscore
  If MotorRunning<>1 And InProgress=True Then
    If queuedscore>=50000 Then
      tempscore=50000
      queuedscore=queuedscore-50000
      SetMotor2(50000)
      Exit Sub
    End If
    If queuedscore>=10000 Then
      tempscore=10000
      queuedscore=queuedscore-10000
      SetMotor2(10000)
      Exit Sub
    End If
    If queuedscore>=5000 Then
      tempscore=5000
      queuedscore=queuedscore-5000
      SetMotor2(5000)
      Exit Sub
    End If
    If queuedscore>=4000 Then
      tempscore=4000
      queuedscore=queuedscore-4000
      SetMotor2(4000)
      Exit Sub
    End If
    If queuedscore>=3000 Then
      tempscore=3000
      queuedscore=queuedscore-3000
      SetMotor2(3000)
      Exit Sub
    End If
    If queuedscore>=2000 Then
      tempscore=2000
      queuedscore=queuedscore-2000
      SetMotor2(2000)
      Exit Sub
    End If
    If queuedscore>=1000 Then
      tempscore=1000
      queuedscore=queuedscore-1000
      SetMotor2(1000)
      Exit Sub
    End If
    If queuedscore>=500 Then
      tempscore=500
      queuedscore=queuedscore-500
      SetMotor2(500)
      Exit Sub
    End If
    If queuedscore>=400 Then
      tempscore=400
      queuedscore=queuedscore-400
      SetMotor2(400)
      Exit Sub
    End If
    If queuedscore>=300 Then
      tempscore=300
      queuedscore=queuedscore-300
      SetMotor2(300)
      Exit Sub
    End If
    If queuedscore>=200 Then
      tempscore=200
      queuedscore=queuedscore-200
      SetMotor2(200)
      Exit Sub
    End If
    If queuedscore>=100 Then
      tempscore=100
      queuedscore=queuedscore-100
      SetMotor2(100)
      Exit Sub
    End If
    If queuedscore>=50 Then
      tempscore=50
      queuedscore=queuedscore-50
      SetMotor2(50)
      Exit Sub
    End If
    If queuedscore>=40 Then
      tempscore=40
      queuedscore=queuedscore-40
      SetMotor2(40)
      Exit Sub
    End If
    If queuedscore>=30 Then
      tempscore=30
      queuedscore=queuedscore-30
      SetMotor2(30)
      Exit Sub
    End If
    If queuedscore>=20 Then
      tempscore=20
      queuedscore=queuedscore-20
      SetMotor2(20)
      Exit Sub
    End If
    If queuedscore>=10 Then
      tempscore=10
      queuedscore=queuedscore-10
      SetMotor2(10)
      Exit Sub
    End If
  End If
End Sub

Sub ScoreMotorTimer_Timer
  If MotorPosition > 0 Then
    Select Case MotorPosition
      Case 5,4,3,2:
        If BaseHitCounter>0 Then
          MoveRunners
          BaseHitCounter=BaseHitCounter-1
        End If
        If MotorMode=10000 Then
          AddScore(10000)
        End If

        If MotorMode=1000 Then
          AddScore(1000)
        End If
        If MotorMode=100 Then
          AddScore(100)
        End If
        If MotorMode=10 Then
          AddScore(10)
        End If
        MotorPosition=MotorPosition-1
      Case 1:
        If MotorMode=10000 Then
          AddScore(10000)
        End If

        If MotorMode=1000 Then
          AddScore(1000)
        End If
        If MotorMode=100 Then
          AddScore(100)
        End If
        If MotorMode=10 Then
          AddScore(10)

        End If
        MotorPosition=0:MotorRunning=0:LightsOn: BaseHit=0: BaseHitCounter=0
    End Select
  End If
End Sub

Sub AddScore(x)
  If TableTilted=True Then Exit Sub
  ImpulseReelCount=ImpulseReelCount+1
  If ImpulseReelCount>4 Then
    ImpulseReelCount=0
  End If
  Select Case ScoreAdditionAdjustment
    Case 0:
      AddScore1(x)
    Case 1:
      AddScore2(x)
  End Select
End Sub

Sub AddScore1(x)
' debugtext.text=score
  Select Case x
    Case 1:
      PlayChime(10)
      Score(Player)=Score(Player)+1
    Case 10:
      PlayChime(100)
      Score(Player)=Score(Player)+10
'     debugscore=debugscore+10
    Case 100:
      PlayChime(1000)
      Score(Player)=Score(Player)+100
'     debugscore=debugscore+100
    Case 1000:
      PlayChime(1000)
      Score(Player)=Score(Player)+1000
'     debugscore=debugscore+1000
  End Select
  PlayerScores(Player-1).AddValue(x)
  PlayerScoresOn(Player-1).AddValue(x)
  If ScoreDisplay(Player)<100000 Then
    ScoreDisplay(Player)=Score(Player)
  Else
    Score100K(Player)=Int(Score(Player)/100000)
    ScoreDisplay(Player)=Score(Player)-100000
  End If
  If Score(Player)=>1000000 Then
    If B2SOn Then
      If Player=1 Then
        Controller.B2SSetScoreRolloverPlayer1 Score100K(Player)
      End If
      If Player=2 Then
        Controller.B2SSetScoreRolloverPlayer2 Score100K(Player)
      End If
      If Player=3 Then
        Controller.B2SSetScoreRolloverPlayer3 Score100K(Player)
      End If
      If Player=4 Then
        Controller.B2SSetScoreRolloverPlayer4 Score100K(Player)
      End If
    End If
  End If
  If B2SOn Then
    Controller.B2SSetScorePlayer Player, ScoreDisplay(Player)
  End If
  If Score(Player)>Replay1 And Replay1Paid(Player)=False Then
    Replay1Paid(Player)=True
    AddSpecial
  End If
  If Score(Player)>Replay2 And Replay2Paid(Player)=False Then
    Replay2Paid(Player)=True
    AddSpecial
  End If
  If Score(Player)>Replay3 And Replay3Paid(Player)=False Then
    Replay3Paid(Player)=True
    AddSpecial
  End If
' ScoreText.text=debugscore
End Sub

Sub AddScore2(x)
  Dim OldScore, NewScore, OldTestScore, NewTestScore
    OldScore = Score(Player)
  Select Case x
        Case 1:
            Score(Player)=Score(Player)+1
    Case 10:
      Score(Player)=Score(Player)+10
    Case 100:
      Score(Player)=Score(Player)+100
    Case 1000:
      Score(Player)=Score(Player)+1000
    Case 10000:
      Score(Player)=Score(Player)+10000
    Case 100000:
      Score(Player)=Score(Player)+100000
  End Select
  NewScore = Score(Player)

  If NewScore>=100000 Then
    EVAL("RolloverReel"&Player).SetValue(1)
    If B2SOn Then
      If Player=1 Then
        Controller.B2SSetScoreRolloverPlayer1 1
      End If
      If Player=2 Then
        Controller.B2SSetScoreRolloverPlayer2 1
      End If

      If Player=3 Then
        Controller.B2SSetScoreRolloverPlayer3 1
      End If

      If Player=4 Then
        Controller.B2SSetScoreRolloverPlayer4 1
      End If
    End If
  End If

  OldTestScore = OldScore
  NewTestScore = NewScore
  Do
    If OldTestScore < Replay1 And NewTestScore >= Replay1 And Replay1Paid(Player)=False Then
      AddSpecial()
      Replay1Paid(Player)=True
      NewTestScore = 0
    ElseIf OldTestScore < Replay2 And NewTestScore >= Replay2 AND Replay2Paid(Player)=False Then
      AddSpecial()
      Replay2Paid(Player)=True
      NewTestScore = 0
    ElseIf OldTestScore < Replay3 And NewTestScore >= Replay3 AND Replay3Paid(Player)=False Then
      AddSpecial()
      Replay3Paid(Player)=True
      NewTestScore = 0
    ElseIf OldTestScore < Replay4 And NewTestScore >= Replay4 AND Replay4Paid(Player)=False Then
      AddSpecial()
      Replay4Paid(Player)=True
      NewTestScore = 0
    End If

    NewTestScore = NewTestScore - 1000000
    OldTestScore = OldTestScore - 1000000
  Loop While NewTestScore > 0

    OldScore = int(OldScore / 10) ' divide by 10 For games with fixed 0 In 1s position, by 1 For games with real 1s digits
    NewScore = int(NewScore / 10) ' divide by 10 For games with fixed 0 In 1s position, by 1 For games with real 1s digits
  ' MsgBox("OldScore="&OldScore&", NewScore="&NewScore&", OldScore Mod 10="&OldScore Mod 10 & ", NewScore % 10="&NewScore Mod 10)

    If (OldScore Mod 10 <> NewScore Mod 10) Then
    PlayChime(10)
    End If

    OldScore = int(OldScore / 10)
    NewScore = int(NewScore / 10)
  ' MsgBox("OldScore="&OldScore&", NewScore="&NewScore)
    If (OldScore Mod 10 <> NewScore Mod 10) Then
    PlayChime(100)
    End If

    OldScore = int(OldScore / 10)
    NewScore = int(NewScore / 10)
  ' MsgBox("OldScore="&OldScore&", NewScore="&NewScore)
    If (OldScore Mod 10 <> NewScore Mod 10) Then
    PlayChime(1000)
    End If

    OldScore = int(OldScore / 10)
    NewScore = int(NewScore / 10)
  ' MsgBox("OldScore="&OldScore&", NewScore="&NewScore)
    If (OldScore Mod 10 <> NewScore Mod 10) Then
    PlayChime(1000)
    End If

    OldScore = int(OldScore / 10)
    NewScore = int(NewScore / 10)
  ' MsgBox("OldScore="&OldScore&", NewScore="&NewScore)
    If (OldScore Mod 10 <> NewScore Mod 10) Then
    PlayChime(1000)
    End If

  If B2SOn Then
    Controller.B2SSetScorePlayer Player, Score(Player)
  End If
' EMReel1.SetValue Score(Player)
  PlayerScores(Player-1).AddValue(x)
  PlayerScoresOn(Player-1).AddValue(x)
End Sub

Sub PlayChime(x)
  If DoodleBugFlag=True Then : Exit Sub
  Select Case x
    Case 10
      EMPlayBellSound 0
      DOF kDOF_ChimeUnitHigh, DOFPulse
    Case 100
      EMPlayBellSound 1
      DOF kDOF_ChimeUnitMid, DOFPulse
    Case 1000
      EMPlayBellSound 2
      DOF kDOF_ChimeUnitLow, DOFPulse
  End Select
End Sub

Sub HideOptions()
  '
End Sub

'*****************************************
' ninuzzu's BALL SHADOW
'*****************************************

Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5)

Sub BallShadowUpdate_timer()
    Dim BOT, b
    BOT = GetBalls
    ' hide shadow of deleted balls
    If UBound(BOT)<(tnob-1) Then
        For b = (UBound(BOT) + 1) To (tnob-1)
            BallShadow(b).visible = 0
        Next
    End If
    ' Exit the Sub If no balls on the table
    If UBound(BOT) = -1 Then Exit Sub
    ' render the shadow For Each ball
    For b = 0 To UBound(BOT)
        If BOT(b).X < Table1.Width/2 Then
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/21)) + 6
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/21)) - 6
        End If
        ballShadow(b).Y = BOT(b).Y + 4
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
End Sub

'*****************************************
' Object sounds
'*****************************************

Sub Plastics_Hit (idx)
' PlaySound "woodhit_low", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Pins_Hit (idx)
  EMPlayPinHitSound
End Sub

Sub Targets_Hit (idx)
' PlaySound "target", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
  EMPlayMetalHitSound
End Sub

Sub Metals_Medium_Hit (idx)
  EMPlayMetalHitSound
End Sub

Sub Metals2_Hit (idx)
  EMPlayMetalHitSound
End Sub

Sub Gates_Hit (idx)
  EMPlayGateHitSound
End Sub

Sub Spinner_Spin
' PlaySound "fx_spinner", 0, .25, AudioPan(Spinner), 0.25, 0, 0, 1, AudioFade(Spinner)
End Sub

Sub Rubbers_Hit(idx)
  EMPlayRubberHitSound
End Sub

Sub Posts_Hit(idx)
  EMPlayPostHitSound
End Sub

Sub LeftFlipper_Collide(parm)
  EMPlayLeftFlipperCollideSound parm
End Sub

Sub RightFlipper_Collide(parm)
  EMPlayRightFlipperCollideSound parm
End Sub

'==================================================================================================
'                        EM Sounds
'==================================================================================================

Const kMasterVolume = 0.8         'value no greater than 1.
Const kPanScalar = 0.2            'smaller value, less "spread" of audio channels
Const kFadeScalar = 0.25          'smaller value, less "spread" of audio channels
Const kMinSoundPanLeft = -1.0
Const kSoundPanLeft = -0.9
Const kSoundPanCenter = 0.0
Const kSoundPanRight = 0.9
Const kMaxSoundPanRight = 1.0
Const kMinSoundFadeNearBackglass = -1.0
Const kSoundFadeNearBackglass = -0.9
Const kSoundFadeCenter = 0.0
Const kSoundFadeNearPlayer = 0.9
Const kMaxSoundFadeNearPlayer = 1.0
Const kDontLoopSound = 0
Const kLoopUntilStopped = -1
Const kNoRandomPitch = 0
Const kNoPitchChange = 0
Const kDontUseExistingSound = 0
Const kUseExistingSound = 1
Const kNoRestartSound = 0
Const kRestartSound = 1

Dim tablewidth, tableheight
tablewidth = 1000             'reasonable defaults,
tableheight = 2000              'but please call EMSoundInit() to get real values assigned

' Pass table object, call before any sound is played
Sub EMSoundInit(tableObject)
  tableheight = tableObject.height
  tablewidth = tableObject.width
End Sub

' Calculates the volume of the sound based on the ball speed
Function EMVolumeForBall(ball)
    EMVolumeForBall = Csng(EMBallVelocity(ball) ^2)
End Function

' Scale pan so there is some bleed-over of left/right channels.
' ex: with headphones you don't want the plunger only in the right ear.
Function EMTransformPan(pan)
  EMTransformPan = pan * kPanScalar
End Function

' Determines pan based on x position on table (-1 is table-left, 0 is table-center, 1 is table-right).
Function EMPanForTableX(x)
    EMPanForTableX = Csng((x * 2 / tablewidth) - 1)
  If EMPanForTableX < kMinSoundPanLeft Then
    EMPanForTableX = kMinSoundPanLeft
  ElseIf EMPanForTableX > kMaxSoundPanRight Then
    EMPanForTableX = kMaxSoundPanRight
  End If
End Function

' Scale fade so there is some bleed-over between front/back channels.
Function EMTransformFade(fade)
  EMTransformFade = fade * kFadeScalar
End Function

' Determines fade based on y position on table (-1 is table-far, 0 is table-center, 1 is table-near).
Function EMFadeForTableY(y)
    EMFadeForTableY = Csng((y * 2 / tableheight) - 1)
  If EMFadeForTableY < kMinSoundFadeNearBackglass Then
    EMFadeForTableY = kMinSoundFadeNearBackglass
  ElseIf EMFadeForTableY > kMaxSoundFadeNearPlayer Then
    EMFadeForTableY = kMaxSoundFadeNearPlayer
  End If
End Function

Sub EMPlaySoundAtVolumePanAndFade (soundName, soundVolume, soundPan, soundFade)
  PlaySound soundName, kDontLoopSound, soundVolume * kMasterVolume, soundPan, kNoRandomPitch, kNoPitchChange, kDontUseExistingSound, kNoRestartSound, soundFade
End Sub

Sub EMPlaySoundAtVolumeForObject (soundName, soundVolume, forObj)
  Dim pan
  pan = EMPanForTableX (forObj.x)
  pan = EMTransformPan (pan)
  Dim fade
  fade = EMFadeForTableY (forObj.y)
  fade = EMTransformFade (fade)
  EMPlaySoundAtVolumePanAndFade soundName, soundVolume, pan, fade
End Sub

Sub EMPlaySoundAtVolumeForActiveBall(soundName, soundVolume)
  Dim pan
  pan = EMPanForTableX (ActiveBall.x)
  pan = EMTransformPan (pan)
  Dim fade
  fade = EMFadeForTableY (ActiveBall.y)
  fade = EMTransformFade (fade)
  EMPlaySoundAtVolumePanAndFade soundName, soundVolume, pan, fade
End Sub

Sub EMPlaySoundExistingAtVolumePanAndFade(soundName, soundVolume, soundPan, soundFade)
  PlaySound soundName, kDontLoopSound, soundVolume * kMasterVolume, soundPan, kNoRandomPitch, kNoPitchChange, kUseExistingSound, kNoRestartSound, soundFade
End Sub

Sub EMPlaySoundExistingAtVolumeForBall(soundName, soundVolume)
  Dim pan
  pan = EMPanForTableX (ActiveBall.x)
  pan = EMTransformPan (pan)
  Dim fade
  fade = EMFadeForTableY (ActiveBall.y)
  fade = EMTransformFade (fade)
  EMPlaySoundExistingAtVolumePanAndFade soundName, soundLevel, pan, fade
End Sub

Sub EMPlaySoundLoopedAtVolumeForObject(soundName, loopCount, soundVolume, forObj)
  Dim pan
  pan = EMPanForTableX (forObj.x)
  pan = EMTransformPan (pan)
  Dim fade
  fade = EMFadeForTableY (forObj.y)
  fade = EMTransformFade (fade)
  PlaySound soundName, loopCount, soundVolume * kMasterVolume, pan, kNoRandomPitch, kNoPitchChange, kDontUseExistingSound, kRestartSound, fade
End Sub

'------------------------------------------------------------------------------------- Ball Rolling

Function EMBallVelocity(ball) 'Calculates the ball speed
    EMBallVelocity = INT(SQR((ball.VelX ^2) + (ball.VelY ^2)))
End Function

Dim RollingSoundScalar
RollingSoundScalar = 0.22

Function EMVolPlayfieldRoll(ball) ' Calculates the roll volume of the sound based on the ball speed
  EMVolPlayfieldRoll = RollingSoundScalar * 0.0005 * Csng(EMBallVelocity(ball) ^3)
End Function

Function EMPitchPlayfieldRoll(ball) ' Calculates the roll pitch of the sound based on the ball speed
  EMPitchPlayfieldRoll = EMBallVelocity(ball) ^2 * 15
End Function

Function EMPitch(ball) ' Calculates the pitch of the sound based on the ball speed
    EMPitch = EMBallVelocity(ball) * 20
End Function

Const tnob = 5    'total number of balls
Const lob = 0   'number of locked balls

ReDim rolling(tnob)
InitRolling

Sub InitRolling
  Dim i
  For i = 0 To tnob
    rolling(i) = False
  Next
End Sub

Sub RollingSoundTimer_Timer()
  Dim BOT, b
  BOT = GetBalls

  ' stop the sound of deleted balls
  For b = UBound(BOT) + 1 To tnob
    rolling(b) = False
    StopSound("BallRoll_" & b)
  Next

  ' exit the Sub If no balls on the table
  If UBound(BOT) = -1 Then Exit Sub

  ' play the rolling sound For each ball
  For b = 0 To UBound(BOT)
    If EMBallVelocity(BOT(b)) > 1 AND BOT(b).z < 30 Then
      rolling(b) = True
      Dim pan
      pan = EMPanForTableX (BOT(b).x)
      pan = EMTransformPan (pan)
      Dim fade
      fade = EMFadeForTableY (BOT(b).y)
      fade = EMTransformFade (fade)
      PlaySound ("BallRoll_" & b), kLoopUntilStopped, EMVolPlayfieldRoll(BOT(b)) * 1.1 * kMasterVolume, pan, kNoRandomPitch, EMPitchPlayfieldRoll(BOT(b)), kUseExistingSound, kNoRestartSound, fade
    Else
      If rolling(b) = True Then
        StopSound("BallRoll_" & b)
        rolling(b) = False
      End If
    End If
  Next
End Sub

'------------------------------------------------------------------------------------- Start Sounds

Dim CoinSoundLevel
CoinSoundLevel = 1              'volume level; range [0, 1]

Sub EMPlayCoinSound()
  EMPlaySoundAtVolumePanAndFade ("Coin_In_" & Int(Rnd * 3) + 1), CoinSoundLevel, kSoundPanCenter, EMTransformFade(kSoundFadeNearPlayer)
End Sub

Dim StartButtonLevel
StartButtonLevel = 1.0            'volume level; range [0, 1]

Sub EMPlayStartButttonSound()
  EMPlaySoundAtVolumePanAndFade "Start_Button", StartButtonLevel, kSoundPanLeft, kSoundFadeNearPlayer
End Sub

Dim StartupSoundLevel
StartupSoundLevel = 1           'volume level; range [0, 1]

Sub EMPlayStartupSound()
  EMPlaySoundAtVolumePanAndFade "StartUpSequence", StartupSoundLevel, kSoundPanCenter, EMTransformFade(kSoundFadeNearBackglass)
End Sub

Dim PlungerPullSoundLevel
PlungerPullSoundLevel = 0.8         'volume level; range [0, 1]

Sub EMPlayPlungerPullSound(plungerObj)
  EMPlaySoundAtVolumeForObject "Plunger_Pull_1", PlungerPullSoundLevel, plungerObj
End Sub

Dim PlungerReleaseSoundLevel
PlungerReleaseSoundLevel = 0.8        'volume level; range [0, 1]

Sub EMPlayPlungerReleaseBallSound(plungerObj)
  EMPlaySoundAtVolumeForObject "Plunger_Release_Ball", PlungerReleaseSoundLevel, plungerObj
End Sub

Dim StartBallSoundLevel
StartBallSoundLevel = 0.4         'volume level; range [0, 1]

Sub EMPlayStartBallSound(which)
  Dim sndName
  If which=0 Then
    sndName = "StartBall1"
  Else
    sndName = "StartBall2-5"
  End If
  EMPlaySoundAtVolumePanAndFade sndName, StartBallSoundLevel, kSoundPanCenter, EMTransformFade(kSoundFadeNearPlayer)
End Sub

Dim BallReleaseSoundLevel
BallReleaseSoundLevel = 1.0         'volume level; range [0, 1]

Sub EMPlayBallReleaseSound()
  EMPlaySoundAtVolumePanAndFade ("BallRelease" & Int(Rnd * 7) + 1), BallReleaseSoundLevel, EMTransformPan(kSoundPanRight), EMTransformFade(kSoundFadeNearPlayer)
End Sub

Dim RotateThroughPlayersSoundLevel
RotateThroughPlayersSoundLevel = 0.8    'volume level; range [0, 1]

Sub EMPlayRotateThroughPlayersSound
  EMPlaySoundAtVolumePanAndFade "RotateThruPlayers", RotateThroughPlayersSoundLevel, kSoundPanCenter, EMTransformFade(kSoundFadeNearPlayer)
End Sub

'----------------------------------------------------------------------------------- Flipper Sounds

Dim FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel
FlipperUpAttackMinimumSoundLevel = 0.010  'volume level; range [0, 1]
FlipperUpAttackMaximumSoundLevel = 0.635  'volume level; range [0, 1]

Dim FlipperUpSoundLevel
FlipperUpSoundLevel = 1.0         'volume level; range [0, 1]

Dim FlipperLeftHitParm, FlipperRightHitParm
FlipperLeftHitParm = FlipperUpSoundLevel  'sound helper; not configurable
FlipperRightHitParm = FlipperUpSoundLevel 'sound helper; not configurable

Dim FlipperBuzzSoundLevel
FlipperBuzzSoundLevel = 0.05        'volume level; range [0, 1]

Sub EMPlayLeftFlipperUpAttackSound (flipperObj)
  Dim soundLevel
  soundLevel = Rnd() * (FlipperUpAttackMaximumSoundLevel-FlipperUpAttackMinimumSoundLevel) + FlipperUpAttackMinimumSoundLevel
  EMPlaySoundAtVolumeForObject SoundFX("Flipper_Attack-L01", DOFFlippers), soundLevel, flipperObj
  EMPlaySoundLoopedAtVolumeForObject "buzzL", kLoopUntilStopped, FlipperBuzzSoundLevel, flipperObj
End Sub

Sub EMPlayRightFlipperUpAttackSound (flipperObj)
  Dim soundLevel
  soundLevel = Rnd() * (FlipperUpAttackMaximumSoundLevel-FlipperUpAttackMinimumSoundLevel) + FlipperUpAttackMinimumSoundLevel
  EMPlaySoundAtVolumeForObject SoundFX("Flipper_Attack-R01", DOFFlippers), soundLevel, flipperObj
  EMPlaySoundLoopedAtVolumeForObject "buzz", kLoopUntilStopped, FlipperBuzzSoundLevel, flipperObj
End Sub

Sub EMPlayLeftFlipperUpSound (flipperObj)
  EMPlaySoundAtVolumeForObject SoundFX("Flipper_L0" & Int(Rnd * 9) + 1, DOFFlippers), FlipperLeftHitParm, flipperObj
End Sub

Sub EMPlayRightFlipperUpSound (flipperObj)
  EMPlaySoundAtVolumeForObject SoundFX("Flipper_R0" & Int(Rnd * 9) + 1, DOFFlippers), FlipperRightHitParm, flipperObj
End Sub

Dim FlipperReflipSoundLevel
FlipperReflipSoundLevel = 0.8       'volume level; range [0, 1]

Sub EMPlayLeftFlipperReflipSound (flipperObj)
  EMPlaySoundAtVolumeForObject SoundFX("Flipper_ReFlip_L0" & Int(Rnd * 3) + 1, DOFFlippers), FlipperReflipSoundLevel, flipperObj
End Sub

Sub EMPlayRightFlipperReflipSound (flipperObj)
  EMPlaySoundAtVolumeForObject SoundFX("Flipper_ReFlip_R0" & Int(Rnd * 3) + 1, DOFFlippers), FlipperReflipSoundLevel, flipperObj
End Sub

Dim FlipperDownSoundLevel
FlipperDownSoundLevel = 0.45        'volume level; range [0, 1]

Sub EMPlayLeftFlipperDownSound(flipperObj)
  EMPlaySoundAtVolumeForObject SoundFX("Flipper_Left_Down_" & Int(Rnd * 7) + 1, DOFFlippers), FlipperDownSoundLevel, flipperObj
  StopSound "buzzL"
End Sub

Sub EMPlayRightFlipperDownSound(flipperObj)
  EMPlaySoundAtVolumeForObject SoundFX("Flipper_Right_Down_" & Int(Rnd * 8) + 1, DOFFlippers), FlipperDownSoundLevel, flipperObj
  StopSound "buzz"
End Sub

Dim RubberFlipperSoundScalar
RubberFlipperSoundScalar = 0.015      'volume multiplier; must not be zero

Sub EMPlayLeftFlipperCollideSound(parm)
  FlipperLeftHitParm = parm / 10
  If FlipperLeftHitParm > 1 Then
    FlipperLeftHitParm = 1
  End If
  FlipperLeftHitParm = FlipperUpSoundLevel * FlipperLeftHitParm
  EMPlaySoundAtVolumeForActiveBall ("Flipper_Rubber_" & Int(Rnd * 7) + 1), parm  * RubberFlipperSoundScalar
End Sub

Sub EMPlayRightFlipperCollideSound(parm)
  FlipperRightHitParm = parm / 10
  If FlipperRightHitParm > 1 Then
    FlipperRightHitParm = 1
  End If
  FlipperRightHitParm = FlipperUpSoundLevel * FlipperRightHitParm
  EMPlaySoundAtVolumeForActiveBall ("Flipper_Rubber_" & Int(Rnd * 7) + 1), parm  * RubberFlipperSoundScalar
End Sub

Sub EMStopBuzzSounds ()
  StopSound "buzz"
  StopSound "buzzL"
End Sub

'--------------------------------------------------------------------------------- Playfield Sounds

Dim MetalImpactSoundScalar
MetalImpactSoundScalar = 0.025

Sub EMPlayMetalHitSound
  EMPlaySoundAtVolumeForActiveBall ("Metal_Touch_" & Int(Rnd * 13) + 1), EMVolumeForBall(ActiveBall) * MetalImpactSoundScalar
End Sub

Dim GateHitSoundLevel
GateHitSoundLevel = 0.1           'volume level; range [0, 1]

Sub EMPlayGateHitSound()
  EMPlaySoundAtVolumeForActiveBall "gate4", GateHitSoundLevel
End Sub

Dim DrainSoundLevel
DrainSoundLevel = 0.8           'volume level; range [0, 1]

Sub EMPlayDrainSound(drainObj)
  EMPlaySoundAtVolumeForObject ("Drain_" & Int(Rnd * 11) + 1), DrainSoundLevel, drainObj
End Sub

Dim RubberStrongSoundScalar, RubberWeakSoundScalar
RubberStrongSoundScalar = 0.011       'volume multiplier; must not be zero
RubberWeakSoundScalar = 0.015       'volume multiplier; must not be zero

Sub EMPlayRubberHitSound()
  Dim finalspeedSquared
  finalspeedSquared = (activeball.velx * activeball.velx) + (activeball.vely * activeball.vely)
  If finalspeedSquared > 25 Then      'strong
    Dim soundIndex
    soundIndex = Int(Rnd * 10) + 1
    If soundIndex > 9 Then
      EMPlaySoundAtVolumeForActiveBall "Rubber_1_Hard", EMVolumeForBall(ActiveBall) * RubberStrongSoundScalar * 0.6
    Else
      EMPlaySoundAtVolumeForActiveBall ("Rubber_Strong_" & Int(soundIndex)), EMVolumeForBall(ActiveBall) * RubberStrongSoundScalar
    End If
  Else                  'weak
    EMPlaySoundAtVolumeForActiveBall ("Rubber_" & Int(Rnd * 9) + 1), EMVolumeForBall(ActiveBall) * RubberWeakSoundScalar
  End If
End Sub

Dim PostSoundScalar
PostSoundScalar = 0.1           'volume level; range [0, 1]

Sub EMPlayPostHitSound()
  Dim finalspeedSquared
    finalspeedSquared = (activeball.velx * activeball.velx) + (activeball.vely * activeball.vely)
  If finalspeedSquared > 256 Then
    EMPlaySoundAtVolumeForActiveBall "fx_rubber2", EMVolumeForBall(ActiveBall) * PostSoundScalar
  ElseIf finalspeedSquared >= 36 Then
    EMPlayRubberHitSound()
  End If
End Sub

Dim SaucerLockSoundLevel
SaucerLockSoundLevel = 0.8          'volume level; range [0, 1]

Sub EMPlaySaucerLockSound()
  EMPlaySoundAtVolumeForActiveBall ("Saucer_Enter_" & Int(Rnd * 2) + 1), SaucerLockSoundLevel
End Sub

Dim SaucerKickSoundLevel
SaucerKickSoundLevel = 0.8          'volume level; range [0, 1]

Sub EMPlaySaucerKickSound(scenario, saucerObj)
  Select Case scenario
    Case 0: EMPlaySoundAtVolumeForObject SoundFX("Saucer_Empty", DOFContactors), SaucerKickSoundLevel, saucerObj
    Case 1: EMPlaySoundAtVolumeForObject SoundFX("Saucer_Kick", DOFContactors), SaucerKickSoundLevel, saucerObj
  End Select
End Sub

Dim BumperSoundScalar
BumperSoundScalar = 6.0           'volume multiplier; must not be zero

Sub EMPlayTopBumperSound(bumperObj)
  EMPlaySoundAtVolumeForObject SoundFX("Bumpers_Top_" & Int(Rnd * 5) + 1,DOFContactors), EMVolumeForBall(ActiveBall) * BumperSoundScalar, bumperObj
End Sub

Dim SensorSoundLevel
SensorSoundLevel = 1.0            'volume level; range [0, 1]

Sub EMPlaySensorSound()
  EMPlaySoundAtVolumeForActiveBall "sensor", SensorSoundLevel
End Sub

Dim TargetSoundScalar
TargetSoundScalar = 0.025         'volume multiplier; must not be zero

Sub EMPlayTargetHitSound()
  Dim finalspeedSquared
  finalspeedSquared = (activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeedSquared > 100 Then
    EMPlaySoundAtVolumeForActiveBall SoundFX("Target_Hit_" & Int(Rnd * 4) + 5, DOFTargets), EMVolumeForBall(ActiveBall) * 0.45 * TargetSoundScalar
  Else
    EMPlaySoundAtVolumeForActiveBall SoundFX("Target_Hit_" & Int(Rnd * 4) + 1, DOFTargets), EMVolumeForBall(ActiveBall) * TargetSoundScalar
  End If
End Sub

Dim PinHitSoundLevel
PinHitSoundLevel = 1.0            'volume level; range [0, 1]

Sub EMPlayPinHitSound()
  EMPlaySoundAtVolumeForActiveBall "pinhit_low", PinHitSoundLevel
End Sub

Dim SpinnerSoundLevel
SpinnerSoundLevel = 1.0           'volume level; range [0, 1]

Sub EMPlaySpinnerSound (spinnerObj)
  EMPlaySoundAtVolumeForObject "Spinner", SpinnerSoundLevel, spinnerObj
End Sub

'------------------------------------------------------------------------------------- Other Sounds

Dim ClickLevel
ClickLevel = 1.0              'volume level; range [0, 1]

Sub EMPlayClickSound()
  EMPlaySoundAtVolumePanAndFade "click", ClickLevel, kSoundPanCenter, kSoundFadeCenter
End Sub

Dim BellLevel
BellLevel = 0.10              'volume level; range [0, 1]

Sub EMPlayBellSound(bellNum)
  Dim sndName
  Select Case bellNum
    Case 0
      EMPlaySoundAtVolumePanAndFade "SpinACard_10_Point_Bell", BellLevel, EMTransformPan(kSoundPanRight), kSoundFadeCenter
    Case 1
      EMPlaySoundAtVolumePanAndFade "SpinACard_100_Point_Bell", BellLevel, EMTransformPan(kSoundPanRight), kSoundFadeCenter
    Case 2
      EMPlaySoundExistingAtVolumePanAndFade "SpinACard_1000_Point_Bell", BellLevel, EMTransformPan(kSoundPanRight), kSoundFadeCenter
  End Select
End Sub

Dim MotorLeerSoundLevel
MotorLeerSoundLevel = 1.0           'volume level; range [0, 1]

Sub EMPlayMotorLeerSound()
  EMPlaySoundAtVolumePanAndFade "MotorLeer", MotorLeerSoundLevel, kSoundPanCenter, EMTransformFade(kSoundFadeNearBackglass)
End Sub

Dim MotorPauseSoundLevel
MotorPauseSoundLevel = 1.0          'volume level; range [0, 1]

Sub EMPlayMotorPauseSound()
  EMPlaySoundAtVolumePanAndFade "MotorPause", MotorPauseSoundLevel, kSoundPanCenter, EMTransformFade(kSoundFadeNearBackglass)
End Sub

'==================================================================================================
'                      End EM Sounds
'==================================================================================================

' ============================================================================================
' GNMOD - Multiple High Score Display And Collection
' ============================================================================================

Dim EnteringInitials    ' Normally zero, Set To non-zero To enter initials
EnteringInitials = 0

Dim PlungerPulled
PlungerPulled = 0

Dim SelectedChar      ' character under the "cursor" when entering initials

Dim HSTimerCount      ' Pass counter For HS timer, scores are cycled by the timer
HSTimerCount = 5      ' Timer is initially enabled, it'll wrap from 5 To 1 when it's displayed

Dim InitialString     ' the string holding the player's initials as they're entered

Dim AlphaString       ' A-Z, 0-9, space (_) And backspace (<)
Dim AlphaStringPos      ' pointer To AlphaString, move forward And backward with flipper keys
AlphaString = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_<"

Dim HSNewHigh       ' The new score To be recorded

Dim HSScore(10)       ' High Scores read In from config file
Dim HSName(10)        ' High Score Initials read In from config file

' default high scores, remove this when the scores are available from the config file
HSScore(1) = 75000
HSScore(2) = 70000
HSScore(3) = 60000
HSScore(4) = 55000
HSScore(5) = 50000

HSName(1) = "AAA"
HSName(2) = "ZZZ"
HSName(3) = "XXX"
HSName(4) = "ABC"
HSName(5) = "BBB"

HSScore(6) = 750000
HSScore(7) = 700000
HSScore(8) = 600000
HSScore(9) = 550000
HSScore(10) = 500000

HSName(6) = "AAA"
HSName(7) = "ZZZ"
HSName(8) = "XXX"
HSName(9) = "ABC"
HSName(10) = "BBB"

Sub HighScoreTimer_Timer
  If EnteringInitials Then
    If HSTimerCount = 1 Then
      SetHSLine 3, InitialString & MID(AlphaString, AlphaStringPos, 1)
      HSTimerCount = 2
    Else
      SetHSLine 3, InitialString
      HSTimerCount = 1
    End If
  ElseIf InProgress Then
    SetHSLine 1, "HIGH SCORE1"
    SetHSLine 2, HSScore(1)
    SetHSLine 3, HSName(1)
    HSTimerCount = 5  ' Set so the highest score will show after the game is over
    HighScoreTimer.enabled=False
  ElseIf CheckAllScores Then
    NewHighScore sortscores(ScoreChecker),sortplayers(ScoreChecker)
  Else
    ' cycle through high scores
    HighScoreTimer.interval=2000
    HSTimerCount = HSTimerCount + 1
    If HsTimerCount > 5 Then
      HSTimerCount = 1
    End If
    SetHSLine 1, "HIGH SCORE"+FormatNumber(HSTimerCount,0)
    SetHSLine 2, HSScore(HSTimerCount)
    SetHSLine 3, HSName(HSTimerCount)
  End If
End Sub

Function GetHSChar(String, Index)
  Dim ThisChar
  Dim FileName
  ThisChar = Mid(String, Index, 1)
  FileName = "PostIt"
  If ThisChar = " " Or ThisChar = "" Then
    FileName = FileName & "BL"
  ElseIf ThisChar = "<" Then
    FileName = FileName & "LT"
  ElseIf ThisChar = "_" Then
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
  Dim xfor
  StartHSArray=array(0,1,12,22)
  EndHSArray=array(0,11,21,31)
  StrLen = len(string)
  Index = 1

  For xfor = StartHSArray(LineNo) To EndHSArray(LineNo)
    Eval("HS"&xfor).image = GetHSChar(String, Index)
    Index = Index + 1
  Next
End Sub

Sub NewHighScore(NewScore, PlayNum)
  If NewScore > HSScore(5) Then
    HighScoreTimer.interval = 500
    HSTimerCount = 1
    AlphaStringPos = 1    ' start with first character "A"
    EnteringInitials = 1  ' intercept the control keys while entering initials
    InitialString = ""    ' initials entered so far, initialize To empty
    SetHSLine 1, "PLAYER "+FormatNumber(PlayNum,0)
    SetHSLine 2, "ENTER NAME"
    SetHSLine 3, MID(AlphaString, AlphaStringPos, 1)
    HSNewHigh = NewScore
    For xx=1 To HighScoreReward
      AddSpecial
    Next
  End If
  ScoreChecker=ScoreChecker-1
  If ScoreChecker=0 Then
    CheckAllScores=0
  End If
End Sub

Sub CollectInitials(keycode)
  If keycode = LeftFlipperKey Then
    ' back up To previous character
    AlphaStringPos = AlphaStringPos - 1
    If AlphaStringPos < 1 Then
      AlphaStringPos = len(AlphaString)   ' handle wrap from beginning To End
      If InitialString = "" Then
        ' Skip the backspace If there are no characters To backspace over
        AlphaStringPos = AlphaStringPos - 1
      End If
    End If
    SetHSLine 3, InitialString & MID(AlphaString, AlphaStringPos, 1)
    PlaySound "droptargetdropped"
  ElseIf keycode = RightFlipperKey Then
    ' advance To Next character
    AlphaStringPos = AlphaStringPos + 1
    If AlphaStringPos > len(AlphaString) Or (AlphaStringPos = len(AlphaString) And InitialString = "") Then
      ' Skip the backspace If there are no characters To backspace over
      AlphaStringPos = 1
    End If
    SetHSLine 3, InitialString & MID(AlphaString, AlphaStringPos, 1)
    PlaySound "droptargetdropped"
  ElseIf keycode = StartGameKey Or keycode = PlungerKey Then
    SelectedChar = MID(AlphaString, AlphaStringPos, 1)
    If SelectedChar = "_" Then
      InitialString = InitialString & " "
      PlaySound("Ding10")
    ElseIf SelectedChar = "<" Then
      InitialString = MID(InitialString, 1, len(InitialString) - 1)
      If len(InitialString) = 0 Then
        ' If there are no more characters To back over, don't leave the < displayed
        AlphaStringPos = 1
      End If
      PlaySound("Ding100")
    Else
      InitialString = InitialString & SelectedChar
      PlaySound("Ding10")
    End If
    If len(InitialString) < 3 Then
      SetHSLine 3, InitialString & SelectedChar
    End If
  End If
  If len(InitialString) = 3 Then
    ' save the score
    For i = 5 To 1 step -1
      If i = 1 Or (HSNewHigh > HSScore(i) And HSNewHigh <= HSScore(i - 1)) Then
        ' Replace the score at this location
        If i < 5 Then
' MsgBox("Moving " & i & " To " & (i + 1))
          HSScore(i + 1) = HSScore(i)
          HSName(i + 1) = HSName(i)
        End If
' MsgBox("Saving initials " & InitialString & " To position " & i)
        EnteringInitials = 0
        HSScore(i) = HSNewHigh
        HSName(i) = InitialString
        HSTimerCount = 5
        HighScoreTimer_Timer
        HighScoreTimer.interval = 2000
        PlaySound("Ding1000")
        Exit Sub
      ElseIf i < 5 Then
        ' move the score In this slot down by 1, it's been exceeded by the new score
' MsgBox("Moving " & i & " To " & (i + 1))
        HSScore(i + 1) = HSScore(i)
        HSName(i + 1) = HSName(i)
      End If
    Next
  End If
End Sub
' END GNMOD

' ============================================================================================
' GNMOD - New Options menu
' ============================================================================================
Dim EnteringOptions
Dim CurrentOption
Dim OptionCHS
Dim MaxOption
Dim OptionHighScorePosition
Dim XOpt
Dim StartingArray
Dim EndingArray

StartingArray=Array(0,1,2,30,33,61,89,117,145,173,201,229)
EndingArray=Array(0,1,29,32,60,88,116,144,172,200,228,256)
EnteringOptions = 0
MaxOption = 9
OptionCHS = 0
OptionHighScorePosition = 0
Const OptionLinesToMark="111111011"
Const OptionLine1="" 'do not use this line
Const OptionLine2="" 'do not use this line
Const OptionLine3="" 'do not use this line
Const OptionLine4="Add Bonus Score"
Const OptionLine5="Super Bonus Adds"
Const OptionLine6="Spinner And Ex. Ball Option"
Const OptionLine7=""
Const OptionLine8="" 'do not use this line
Const OptionLine9="" 'do not use this line

Sub OperatorMenuTimer_Timer
  EnteringOptions = 1
  OperatorMenuTimer.enabled=False
  ShowOperatorMenu
End Sub

Sub ShowOperatorMenu
  OperatorMenuBackdrop.image = "OperatorMenu"
  OptionCHS = 0
  CurrentOption = 1
  DisplayAllOptions
  OperatorOption1.image = "BluePlus"
  SetHighScoreOption
End Sub

Sub DisplayAllOptions
  Dim linecounter
  Dim tempstring
  Dim TempText1
  Dim TempText2
  Dim TempText3
  For linecounter = 1 To MaxOption
    tempstring=Eval("OptionLine"&linecounter)
    Select Case linecounter
      Case 1:
        tempstring=tempstring + FormatNumber(BallsPerGame,0)
        SetOptLine 1,tempstring
      Case 2:
        If Replay3Table(ReplayLevel)=999000 Then
          tempstring = FormatNumber(Replay1Table(ReplayLevel),0) + "/" + FormatNumber(Replay2Table(ReplayLevel),0)
        ElseIf Replay4Table(ReplayLevel)=999000 Then
          tempstring = FormatNumber(Replay1Table(ReplayLevel),0) + "/" + FormatNumber(Replay2Table(ReplayLevel),0) + "/" + FormatNumber(Replay3Table(ReplayLevel),0)
        Else
          tempstring = FormatNumber(Replay1Table(ReplayLevel),0) + "/" + FormatNumber(Replay2Table(ReplayLevel),0) + "/" + FormatNumber(Replay3Table(ReplayLevel),0) + "/" + FormatNumber(Replay4Table(ReplayLevel),0)
        End If
        SetOptLine 2,tempstring
      Case 3:
        If OptionCHS=0 Then
          tempstring = "NO"
        Else
          tempstring = "YES"
        End If
        SetOptLine 3,tempstring
      Case 4:
        SetOptLine 4, tempstring
        tempstring = AdvanceFlag

        SetOptLine 5, tempstring
      Case 5:
        SetOptLine 6, tempstring
        tempstring = SuperAdvanceFlag
        SetOptLine 7, tempstring
      Case 6:
        SetOptLine 8, tempstring
        Select Case ExtraBallFlag
          Case 0:
            tempstring="Liberal"
          Case 1:
            tempstring="Medium"
          Case 2:
            tempstring="Conservative"
        End Select
        SetOptLine 9, tempstring
      Case 7:
        SetOptLine 10, tempstring
        SetOptLine 11, tempstring
      Case 8:
      Case 9:
    End Select
  Next
End Sub

Sub MoveArrow
  Do
    CurrentOption = CurrentOption + 1
    If CurrentOption>Len(OptionLinesToMark) Then
      CurrentOption=1
    End If
  Loop Until Mid(OptionLinesToMark,CurrentOption,1)="1"
End Sub

Sub CollectOptions(ByVal keycode)
  If Keycode = LeftFlipperKey Then
    PlaySound "droptargetdropped"
    For XOpt = 1 To MaxOption
      Eval("OperatorOption"&XOpt).image = "PostitBL"
    Next
    MoveArrow
    If CurrentOption<8 Then
      Eval("OperatorOption"&CurrentOption).image = "BluePlus"
    ElseIf CurrentOption=8 Then
      Eval("OperatorOption"&CurrentOption).image = "GreenCheck"
    Else
      Eval("OperatorOption"&CurrentOption).image = "RedX"
    End If
  ElseIf Keycode = RightFlipperKey Then
    PlaySound "droptargetdropped"
    If CurrentOption = 1 Then
      If BallsPerGame = 3 Then
        BallsPerGame = 5
      Else
        BallsPerGame = 3
      End If
      DisplayAllOptions
    ElseIf CurrentOption = 2 Then
      ReplayLevel=ReplayLevel+1
      If ReplayLevel>ReplayTableMax Then
        ReplayLevel=1
      End If
      DisplayAllOptions
    ElseIf CurrentOption = 3 Then
      If OptionCHS = 0 Then
        OptionCHS = 1
      Else
        OptionCHS = 0
      End If
      DisplayAllOptions
    ElseIf CurrentOption = 4 Then
      If AdvanceFlag = 100 Then
        AdvanceFlag = 1000
      Else
        AdvanceFlag = 100
      End If
      DisplayAllOptions
    ElseIf CurrentOption = 5 Then
      If SuperAdvanceFlag = 5 Then
        SuperAdvanceFlag = 3
      Else
        SuperAdvanceFlag = 5
      End If
      DisplayAllOptions
    ElseIf CurrentOption = 6 Then
      ExtraBallFlag=ExtraBallFlag+1
      If ExtraBallFlag>2 Then ExtraBallFlag=0
      DisplayAllOptions
    ElseIf CurrentOption = 8 Or CurrentOption = 9 Then
      If OptionCHS=1 Then
        HSScore(1) = 75000
        HSScore(2) = 70000
        HSScore(3) = 60000
        HSScore(4) = 55000
        HSScore(5) = 50000

        HSName(1) = "AAA"
        HSName(2) = "ZZZ"
        HSName(3) = "XXX"
        HSName(4) = "ABC"
        HSName(5) = "BBB"

        HSScore(6) = 35
        HSScore(7) = 30
        HSScore(8) = 25
        HSScore(9) = 20
        HSScore(10) = 15

        HSName(6) = "AAA"
        HSName(7) = "ZZZ"
        HSName(8) = "XXX"
        HSName(9) = "ABC"
        HSName(10) = "BBB"
      End If
      If CurrentOption = 8 Then
        savehs
      Else
        loadhs
      End If
      OperatorMenuBackdrop.image = "PostitBL"
      For XOpt = 1 To MaxOption
        Eval("OperatorOption"&XOpt).image = "PostitBL"
      Next
      For XOpt = 1 To 256
        Eval("Option"&XOpt).image = "PostItBL"
      Next
      RefreshReplayCard
      InstructCard.image="IC"
      EnteringOptions = 0
    End If
  End If
End Sub

Sub SetHighScoreOption
  '
End Sub

Function GetOptChar(String, Index)
  Dim ThisChar
  Dim FileName
  ThisChar = Mid(String, Index, 1)
  FileName = "PostIt"
  If ThisChar = " " Or ThisChar = "" Then
    FileName = FileName & "BL"
  ElseIf ThisChar = "<" Then
    FileName = FileName & "LT"
  ElseIf ThisChar = "_" Then
    FileName = FileName & "SP"
  ElseIf ThisChar = "/" Then
    FileName = FileName & "SL"
  ElseIf ThisChar = "," Then
    FileName = FileName & "CM"
  Else
    FileName = FileName & ThisChar
  End If
  GetOptChar = FileName
End Function

Dim LineLengths(22) ' maximum number of lines
Sub SetOptLine(LineNo, String)
  Dim DispLen
    Dim StrLen
  Dim xfor
  Dim Letter
  Dim ThisDigit
  Dim ThisChar
  Dim LetterLine
  Dim Index
  Dim LetterName
  StrLen = len(string)
  Index = 1
  StrLen = len(String)
    DispLen = StrLen
    If (DispLen < LineLengths(LineNo)) Then
        DispLen = LineLengths(LineNo)
    End If
  For xfor = StartingArray(LineNo) To StartingArray(LineNo) + DispLen
    Eval("Option"&xfor).image = GetOptChar(string, Index)
    Index = Index + 1
  Next
  LineLengths(LineNo) = StrLen
End Sub

'******************************************************
'       FLIPPER AND RUBBER CORRECTION
'******************************************************

Dim LFPress, RFPress, EOST, EOSA, EOSTnew, EOSaNew
Dim FStrength, FRampUp, fElasticity, EOSRampUp, SOSRampUp
Dim RFEndAngle, LFEndAngle, LF1EndAngle, RF1EndAngle, LFCount, RFCount, LiveCatch

LFEndAngle = Leftflipper.EndAngle
RFEndAngle = RightFlipper.EndAngle

EOST = leftflipper.eosTorque        'End of Swing Torque
EOSA = leftflipper.eosTorqueAngle   'End of Swing Torque Angle
fStrength = LeftFlipper.strength    'Flipper Strength
fRampUp = LeftFlipper.RampUp      'Flipper Ramp Up
fElasticity = LeftFlipper.elasticity  'Flipper Elasticity
EOStNew = 1.0     'new Flipper Torque
EOSaNew = 0.2   'new FLipper Tprque Angle
EOSRampUp = 1.5   'new EOS Ramp Up weaker at EOS because of the weaker holding coil
SOSRampUp = 8.5   'new SOS Ramp Up strong at start because of the stronger starting coil
LiveCatch = 8   'variable To check elapsed time from

'********Need To have a flipper timer To check For these values
Sub flipperTimer_Timer
' lFlip.rotz = leftflipper.currentangle -121
' rFlip.rotz = rightflipper.currentangle +121
' FlipperLSh.RotZ = LeftFlipper.currentangle
' FlipperRSh.RotZ = RightFlipper.currentangle
'
' lFlip.rotz = leftflipper.CurrentAngle -121 'silver metal flipper obj
' lFlipR.rotz = leftflipper.CurrentAngle -121
' rFlip.rotz = RightFlipper.CurrentAngle +121
' rFlipR.rotz = RightFlipper.CurrentAngle +121
' lFlip001.roty = leftflipper001.CurrentAngle
' rFlip001.roty = RightFlipper001.CurrentAngle
'
''  FlipperLShadow1.RotZ = LeftFlipper1.CurrentAngle
''  FlipperRShadow1.RotZ = RightFlipper1.CurrentAngle

  '--------------Flipper Tricks Section
  'What this code does is swing the flipper fast And make the flipper soft near its EOS To enable live catches.  It resets back To the base Table
  'settings once the flipper reaches the End of swing.  The code also makes the flipper starting ramp up high To simulate the stronger starting
  'coil strength And weaker at its EOS To simulate the weaker hold coil.

  If LeftFlipper.CurrentAngle = LeftFlipper.EndAngle And LFPress = 1 Then   'If the flipper is fully swung And the flipper button is pressed Then...
    LeftFlipper.eosTorqueAngle = EOSaNew  'sets flipper EOS Torque Angle To .2
    LeftFlipper.eosTorque = EOStNew     'sets flipper EOS Torque To 1
    LeftFlipper.RampUp = EOSRampUp      'sets flipper ramp up To 1.5
    If LFCount = 0 Then LFCount = GameTime  'sets the variable LFCount = To the elapsed game time
    If GameTime - LFCount < LiveCatch Then  'If less than 8ms have elasped Then we are In a "Live Catch" scenario
      LeftFlipper.Elasticity = 0.1    'sets flipper elasticity WAY DOWN To allow Live Catches
      If LeftFlipper.EndAngle <> LFEndAngle Then LeftFlipper.EndAngle = LFEndAngle  'Keep the flipper at its EOS And don't let it deflect
    Else
      LeftFlipper.Elasticity = fElasticity  'reset flipper elasticity To the base table setting
    End If
  ElseIf LeftFlipper.CurrentAngle > LeftFlipper.startangle - 0.05  Then   'If the flipper has started its swing, make it swing fast To nearly the End...
    LeftFlipper.RampUp = SOSRampUp        'Set flipper Ramp Up high
    LeftFlipper.EndAngle = LFEndAngle - 3   'swing To within 3 degrees of EOS
    LeftFlipper.Elasticity = fElasticity    'Set the elasticity To the base table elasticity
    LFCount = 0
  ElseIf LeftFlipper.CurrentAngle > LeftFlipper.EndAngle + 0.01 Then  'If the flipper has swung past it's End of swing Then...
    LeftFlipper.eosTorque = EOST      'Set the flipper EOS Torque back To the base table setting
    LeftFlipper.eosTorqueAngle = EOSA   'Set the flipper EOS Torque Angle back To the base table setting
    LeftFlipper.RampUp = fRampUp      'Set the flipper Ramp Up back To the base table setting
    LeftFlipper.Elasticity = fElasticity  'Set the flipper Elasticity back To the base table setting
  End If

  If RightFlipper.CurrentAngle = RightFlipper.EndAngle And RFPress = 1 Then
    RightFlipper.eosTorqueAngle = EOSaNew
    RightFlipper.eosTorque = EOStNew
    RightFlipper.RampUp = EOSRampUp
    If RFCount = 0 Then RFCount = GameTime
    If GameTime - RFCount < LiveCatch Then
      RightFlipper.Elasticity = 0.1
      If RightFlipper.EndAngle <> RFEndAngle Then RightFlipper.EndAngle = RFEndAngle
    Else
      RightFlipper.Elasticity = fElasticity
    End If
  ElseIf RightFlipper.CurrentAngle < RightFlipper.StartAngle + 0.05 Then
    RightFlipper.RampUp = SOSRampUp
    RightFlipper.EndAngle = RFEndAngle + 3
    RightFlipper.Elasticity = fElasticity
    RFCount = 0
  ElseIf RightFlipper.CurrentAngle < RightFlipper.EndAngle - 0.01 Then
    RightFlipper.eosTorque = EOST
    RightFlipper.eosTorqueAngle = EOSA
    RightFlipper.RampUp = fRampUp
    RightFlipper.Elasticity = fElasticity
  End If
End Sub

Dim LF : Set LF = New FlipperPolarity
Dim RF : Set RF = New FlipperPolarity
Dim LF1 : Set LF1 = New FlipperPolarity
Dim RF1 : Set RF1 = New FlipperPolarity

InitPolarity

Sub InitPolarity()
  Dim x, a : a = Array(LF, RF, LF1, RF1)
  For Each x In a
    'safety coefficient (diminishes polarity correction only)
    x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
    x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1

    x.enabled = True
    x.TimeDelay = 69    '*****Important, this variable is an offset For the speed that the ball travels down the table To determine If the flippers have been fired
              'This is needed because the corrections To ball trajectory should only applied If the flippers have been fired And the ball is In the trigger zones.
              'FlipAT is Set To GameTime when the ball enters the flipper trigger zones And If GameTime is less than FlipAT + this time delay Then changes To velocity
              'And trajectory are applied.  If the flipper is fired before the ball enters the trigger zone Then with this delay added To FlipAT the changes
              'To tragectory And velocity will not be applied.  Also If the flipper is In the final 20 degrees changes To ball values will also not be applied.
              '"Faster" tables will need a smaller value while "slower" tables will need a larger value To give the ball more time To get To the flipper.
              'If this value is not Set high enough the Flipper Velocity And Polarity corrections will NEVER be applied.
  Next

  'rf.report "Polarity"
  AddPt "Polarity", 0, 0, -2.7
  AddPt "Polarity", 1, 0.16, -2.7
  AddPt "Polarity", 2, 0.33, -2.7
  AddPt "Polarity", 3, 0.37, -2.7 '4.2
  AddPt "Polarity", 4, 0.41, -2.7
  AddPt "Polarity", 5, 0.45, -2.7 '4.2
  AddPt "Polarity", 6, 0.576,-2.7
  AddPt "Polarity", 7, 0.66, -1.8'-2.1896
  AddPt "Polarity", 8, 0.743, -0.5
  AddPt "Polarity", 9, 0.81, -0.5
  AddPt "Polarity", 10, 0.88, 0

  '"Velocity" Profile
  addpt "Velocity", 0, 0,   1
  addpt "Velocity", 1, 0.16, 1.06
  addpt "Velocity", 2, 0.41,  1.05
  addpt "Velocity", 3, 0.53,  1'0.982
  addpt "Velocity", 4, 0.702, 0.968
  addpt "Velocity", 5, 0.95,  0.968
  addpt "Velocity", 6, 1.03,  0.945

  LF.Object = LeftFlipper
' LF1.Object = LeftFlipper001
  LF.EndPoint = EndPointLp  'you can use just a coordinate, or an object with a .x property. Using a couple of simple primitive objects
' LF1.EndPoint = EndPointLp001
  RF.Object = RightFlipper
' RF1.Object = RightFlipper001
' RF1.EndPoint = EndPointRp001
  RF.EndPoint = EndPointRp
End Sub

Sub AddPt(aStr, idx, aX, aY)  'debugger wrapper For adjusting flipper script In-game
  Dim a : a = Array(LF, RF, LF1, RF1)
  Dim x : For Each x In a
    x.addpoint aStr, idx, aX, aY
  Next
End Sub

Sub TriggerLF_Hit() : LF.Addball activeball : End Sub
Sub TriggerLF_UnHit() :  LF.PolarityCorrect activeball : End Sub
Sub TriggerRF_Hit() : RF.Addball activeball : End Sub
Sub TriggerRF_UnHit() : RF.PolarityCorrect activeball : End Sub
Sub TriggerLF1_Hit() : LF.Addball activeball : End Sub
Sub TriggerLF1_UnHit() :  LF.PolarityCorrect activeball : End Sub
Sub TriggerRF1_Hit() : RF.Addball activeball : End Sub
Sub TriggerRF1_UnHit() : RF.PolarityCorrect activeball : End Sub

'Methods:
'.TimeDelay - Delay before trigger shuts off automatically. Default = 80 (ms)
'.AddPoint - "Polarity", "Velocity", "Ycoef" coordinate points. Use one of these 3 strings, keep coordinates sequential. x = %position on the flipper, y = output
'.Object - Set To flipper reference. Optional.
'.StartPoint - Set start point coord. Unnecessary, If .object is used.

'Called with flipper -
'ProcessBalls - catches ball data.
' - OR -
'.Fire - fires flipper.rotatetoend automatically + processballs. Requires .Object To be Set To the flipper.

'***************This is flipperPolarity's addPoint Sub
Class FlipperPolarity
  Public Enabled
  Private FlipAt  'Timer variable (IE 'flip at 723,530ms...)
  Public TimeDelay  'delay before trigger turns off And polarity is disabled TODO Set time!
  private Flipper, FlipperStart, FlipperEnd, LR, PartialFlipCoef
  Private Balls(20), balldata(20)

  Dim PolarityIn, PolarityOut
  Dim VelocityIn, VelocityOut
  Dim YcoefIn, YcoefOut

  Public Sub Class_Initialize
    Redim PolarityIn(0) : Redim PolarityOut(0) : Redim VelocityIn(0) : Redim VelocityOut(0) : Redim YcoefIn(0) : Redim YcoefOut(0)
    Enabled = True: TimeDelay = 50 : LR = 1:  Dim x : For x = 0 To uBound(balls) : balls(x) = Empty : Set Balldata(x) = new spoofBall: Next
  End Sub

  Public Property let Object(aInput) : Set Flipper = aInput : StartPoint = Flipper.x : End Property
  Public Property Let StartPoint(aInput) : If IsObject(aInput) Then FlipperStart = aInput.x Else FlipperStart = aInput : End If : End Property
  Public Property Get StartPoint : StartPoint = FlipperStart : End Property
  Public Property Let EndPoint(aInput) : If IsObject(aInput) Then FlipperEnd = aInput.x Else FlipperEnd = aInput : End If : End Property
  Public Property Get EndPoint : EndPoint = FlipperEnd : End Property

  Public Sub AddPoint(aChooseArray, aIDX, aX, aY) 'Index #, X position, (In) y Position (out)
    Select Case aChooseArray
      Case "Polarity" : ShuffleArrays PolarityIn, PolarityOut, 1 : PolarityIn(aIDX) = aX : PolarityOut(aIDX) = aY : ShuffleArrays PolarityIn, PolarityOut, 0
      Case "Velocity" : ShuffleArrays VelocityIn, VelocityOut, 1 :VelocityIn(aIDX) = aX : VelocityOut(aIDX) = aY : ShuffleArrays VelocityIn, VelocityOut, 0
      Case "Ycoef" : ShuffleArrays YcoefIn, YcoefOut, 1 :YcoefIn(aIDX) = aX : YcoefOut(aIDX) = aY : ShuffleArrays YcoefIn, YcoefOut, 0
    End Select

  End Sub

'********Triggered by a ball hitting the flipper trigger area
  Public Sub AddBall(aBall) : Dim x :
    For x = 0 To uBound(balls)
      If IsEmpty(balls(x)) Then Set balls(x) = aBall : Exit Sub :End If
    Next
  End Sub

  Private Sub RemoveBall(aBall)
    Dim x : For x = 0 To uBound(balls)
      If TypeName(balls(x) ) = "IBall" Then
        If aBall.ID = Balls(x).ID Then
          balls(x) = Empty
          Balldata(x).Reset
        End If
      End If
    Next
  End Sub

'*********Used To rotate flipper since this is removed from the key down For the flippers
  Public Sub Fire()
    Flipper.RotateToEnd
    processballs
  End Sub

  Public Sub ProcessBalls() 'save data of balls In flipper range
    FlipAt = GameTime
    Dim x : For x = 0 To uBound(balls)
      If Not IsEmpty(balls(x) ) Then balldata(x).Data = balls(x)
    Next
    PartialFlipCoef = ((Flipper.StartAngle - Flipper.CurrentAngle) / (Flipper.StartAngle - Flipper.EndAngle))  '% of flipper swing
    PartialFlipCoef = abs(PartialFlipCoef-1)
    If abs(Flipper.currentAngle - Flipper.EndAngle) < 20 Then 'last 20 degrees of swing is not dealt with
      PartialFlipCoef = 0
    End If
  End Sub

'***********gameTime is a global variable of how long the game has progressed In ms
'***********This function lets the table know If the flipper has been fired
  Private Function FlipperOn()
'   TB.text = gameTime & ":" & (FlipAT + TimeDelay) ' ******MOVE TB into view WHEN THIS FLIPPER FUNCTIONALITY IS ADDED TO A NEW TABLE TO CHECK IF THE TIME DELAY IS LONG ENOUGH*****
    If gameTime < FlipAt + TimeDelay Then FlipperOn = True
  End Function  'Timer shutoff For polaritycorrect

'***********This is turned on when a ball leaves the flipper trigger area
  Public Sub PolarityCorrect(aBall)
    If FlipperOn() Then 'don't run this If the flippers are at rest
'           tb.text = "In"
      Dim tmp, BallPos, x, IDX, Ycoef : Ycoef = 1
      Dim teststr : teststr = "Cutoff"
      tmp = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
      If tmp < 0.1 Then 'If real ball position is behind flipper, Exit Sub To prevent stucks  'Disabled 1.03, I think it's the Mesh that's causing stucks, not this
      End If

      'y safety Exit
      If aBall.VelY > -8 Then 'If ball going down Then remove the ball
        RemoveBall aBall
        Exit Sub
      End If
      'Find balldata. BallPos = % on Flipper
      For x = 0 To uBound(Balls)
        If aBall.id = BallData(x).id AND Not isempty(BallData(x).id) Then
          idx = x
          BallPos = PSlope(BallData(x).x, FlipperStart, 0, FlipperEnd, 1)
          If ballpos > 0.65 Then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)        'find safety coefficient 'ycoef' data
        End If
      Next

      'Velocity correction
      If Not IsEmpty(VelocityIn(0) ) Then
'       tb.text = "Vel corr"
        Dim VelCoef
        If IsEmpty(BallData(idx).id) And aBall.VelY < -12 Then 'If tip hit with no collected data, do vel correction anyway
          If PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1) > 1.1 Then 'adjust plz
            VelCoef = LinearEnvelope(5, VelocityIn, VelocityOut)
            If partialflipcoef < 1 Then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)
            If Enabled Then aBall.Velx = aBall.Velx*VelCoef'VelCoef
            If Enabled Then aBall.Vely = aBall.Vely*VelCoef'VelCoef
          End If
        Else
     :      VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)
          If Enabled Then aBall.Velx = aBall.Velx*VelCoef
          If Enabled Then aBall.Vely = aBall.Vely*VelCoef
        End If
      End If

      'Polarity Correction (optional now)
      If Not IsEmpty(PolarityIn(0) ) Then
        If StartPoint > EndPoint Then LR = -1 'Reverse polarity If left flipper
        Dim AddX : AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR
        If Enabled Then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef)
      End If
    End If
    RemoveBall aBall
  End Sub
End Class

'================================
'Helper Functions

Sub ShuffleArray(ByRef aArray, byVal offset) 'shuffle 1d array
  Dim x, aCount : aCount = 0
  Redim a(uBound(aArray) )
  For x = 0 To uBound(aArray) 'Shuffle objects In a temp array
    If Not IsEmpty(aArray(x) ) Then
      If IsObject(aArray(x)) Then
        Set a(aCount) = aArray(x) 'Set creates an object In VB
      Else
        a(aCount) = aArray(x)
      End If
      aCount = aCount + 1
    End If
  Next
  If offset < 0 Then offset = 0
  Redim aArray(aCount-1+offset) 'Resize original array
  For x = 0 To aCount-1   'Set objects back into original array
    If IsObject(a(x)) Then
      Set aArray(x) = a(x)
    Else
      aArray(x) = a(x)
    End If
  Next
End Sub

'**********Takes In more than one array And passes them To ShuffleArray
Sub ShuffleArrays(aArray1, aArray2, offset)
  ShuffleArray aArray1, offset
  ShuffleArray aArray2, offset
End Sub

'**********Calculate ball speed as hypotenuse of velX/velY triangle
Function BallSpeed(ball) 'Calculates the ball speed
    BallSpeed = SQR(ball.VelX^2 + ball.VelY^2 + ball.VelZ^2)
End Function

'**********Calculates the value of Y For an input x using the slope intercept equation
Function PSlope(Input, X1, Y1, X2, Y2)  'Set up line via two points, no clamping. Input X, output Y
  Dim x, y, b, m : x = input : m = (Y2 - Y1) / (X2 - X1) : b = Y2 - m*X2
  Y = M*x+b
  PSlope = Y
End Function

Class spoofball
  Public X, Y, Z, VelX, VelY, VelZ, ID, Mass, Radius
  Public Property Let Data(aBall)
    With aBall
      x = .x : y = .y : z = .z : velx = .velx : vely = .vely : velz = .velz
      id = .ID : mass = .mass : radius = .radius
    End with
  End Property
  Public Sub Reset()
    x = Empty : y = Empty : z = Empty  : velx = Empty : vely = Empty : velz = Empty
    id = Empty : mass = Empty : radius = Empty
  End Sub
End Class

'****************************************************************************
'PHYSICS DAMPENERS

'These are data mined bounce curves,
'dialed In with the In-game elasticity as much as possible To prevent angle / spin issues.
'Requires tracking ballspeed To calculate COR

Sub dPosts_Hit(idx)
  RubbersD.dampen Activeball
End Sub

Sub dSleeves_Hit(idx)
  SleevesD.Dampen Activeball
End Sub

'*********This sets up the rubbers:
Dim RubbersD
Set RubbersD = new Dampener  'Makes a Dampener Class Object
RubbersD.name = "Rubbers"

'cor bounce curve (linear)
'For best results, try To match In-game velocity as closely as possible To the desired curve
RubbersD.addpoint 0, 0, 0.935 '0.96 'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 1, 3.77, 0.935 '0.96
RubbersD.addpoint 2, 5.76, 0.942 '0.967 'dont take this as gospel. If you can data mine rubber elasticitiy, please help!
RubbersD.addpoint 3, 15.84, 0.874
RubbersD.addpoint 4, 56, 0.64 'there's clamping so interpolate up To 56 at least

Dim SleevesD : Set SleevesD = new Dampener  'this is just rubber but cut down To 85%...
SleevesD.name = "Sleeves"
SleevesD.CopyCoef RubbersD, 0.85

'**********Class For dampener section of nfozzy's code
Class Dampener
  Public Print, debugOn 'tbpOut.text
  Public name, Threshold  'Minimum threshold. Useful For Flippers, which don't have a hit threshold.
  Public ModIn, ModOut
  Private Sub Class_Initialize : Redim ModIn(0) : Redim Modout(0): End Sub

  Public Sub AddPoint(aIdx, aX, aY)
    ShuffleArrays ModIn, ModOut, 1 : ModIn(aIDX) = aX : ModOut(aIDX) = aY : ShuffleArrays ModIn, ModOut, 0
  End Sub

  Public Sub Dampen(aBall)
    If threshold Then If BallSpeed(aBall) < threshold Then Exit Sub End If End If
    Dim RealCOR, DesiredCOR, str, coef
'               Uses the LinearEnvelope function To calculate the correction based upon where it's value sits In relation

'               To the addpoint parameters Set above.  Basically interpolates values between Set points In a linear fashion
    DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )

'                Uses the function BallSpeed's value at the point of impact/the active ball's velocity which is constantly being updated
'               RealCor is always less than 1
    RealCOR = BallSpeed(aBall) / cor.ballvel(aBall.id)

'               Divides the desired CoR by the real COR To make a multiplier To correct velocity In x And y
    coef = desiredcor / realcor
'   tb.text = "Coef = " & coef

'               Applies the coef To x And y velocities
' Thalamus - patched :     aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
    aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
  End Sub

'***********This Sub sets the values For Sleeves (or any other future objects) To 85% (or whatever is passed In) of Posts
  Public Sub CopyCoef(aObj, aCoef) 'alternative addpoints, copy with coef
    Dim x : For x = 0 To uBound(aObj.ModIn)
      addpoint x, aObj.ModIn(x), aObj.ModOut(x)*aCoef
    Next
  End Sub
End Class

'*****************************Generates cor.ballVel For dampener
Sub RDampen_Timer() ' 1 ms timer always on
  CoR.Update
End Sub

'*********CoR is Coefficient of Restitution defined as "how much of the kinetic energy remains For the objects
'To rebound from one another vs. how much is lost as heat, or work done deforming the objects
Dim cor : Set cor = New CoRTracker

Class CoRTracker

  Public ballvel

  Private Sub Class_Initialize : Redim ballvel(0) : End Sub

  Public Sub Update() 'tracks In-ball-velocity
    Dim str, b, allBalls, highestID :
    allBalls = getballs

    For Each b In allballs
      If b.id >= HighestID Then highestID = b.id
    Next

    If uBound(ballvel) < highestID Then Redim ballvel(highestID)  'Set bounds

    For Each b In allballs
      ballvel(b.id) = BallSpeed(b)
    Next
  End Sub
End Class

'********Interpolates the value For areas between the low And upper bounds sent To it
Function LinearEnvelope(xInput, xKeyFrame, yLvl)
  Dim y 'Y output
  Dim L 'Line
  Dim ii : For ii = 1 To uBound(xKeyFrame)  'find active line
    If xInput <= xKeyFrame(ii) Then L = ii : Exit For : End If
  Next
  If xInput > xKeyFrame(uBound(xKeyFrame) ) Then L = uBound(xKeyFrame)  'catch line overrun
  Y = pSlope(xInput, xKeyFrame(L-1), yLvl(L-1), xKeyFrame(L), yLvl(L) )

  'clamp 2.0
  If xInput <= xKeyFrame(lBound(xKeyFrame) ) Then Y = yLvl(lBound(xKeyFrame) )  'Clamp lower
  If xInput >= xKeyFrame(uBound(xKeyFrame) ) Then Y = yLvl(uBound(xKeyFrame) )  'Clamp upper

  LinearEnvelope = Y
End Function

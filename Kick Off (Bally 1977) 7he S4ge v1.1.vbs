'*
'*        Kick Off (Bally 1977)
'*        Author : 7he S4ge
'*
'*        based on the table Amigo (Bally 1974) by Loserman76 & leeoneil
'*        design based on the Future Pinball version by jc144
'*        script based on Amigo table & the VP9 version by Ash (IRPinball)
'*
'* 2023-06 / v1.0 / 7he S4ge / First version
'* 2023-06 / v1.1 / 7he S4ge / Lights have been improved in order to be closer to the real machine (by teisen)
'*                 Correction of a score reel displayed in the cabinet version (by ckpin)
'*

option explicit
Randomize
ExecuteGlobal GetTextFile("core.vbs")

On Error Resume Next
ExecuteGlobal GetTextFile("Controller.vbs")
If Err Then MsgBox "Unable to open Controller.vbs. Ensure that it is in the scripts folder."
On Error Goto 0

'Options
Const activeBlinkingLogoB2S = true ' true or false
Const ShadowFlippersOn = true
Const ShadowBallOn = true
Const ShadowConfigFile = false
' See User tab on the Table1 object for change the Night->Day cycle : I use a night setting for render GI lights better
' Operator menu :
' This table has an operator menu : push the LeftFlipperKey before starting the game for display the menu
' Options saved to User directory\KickOff_77VPX.txt
' Options in the operator menu :
' Balls per game : 3 or 5
' Replay card : define scores for win a credit
' Extra ball light option : define scores for win an extra ball
' Special light : define scores for active special

Dim Controller  ' B2S
Dim B2SScore  ' B2S Score Displayed
Const HSFileName="KickOff_77VPX.txt" ' Filename for High scores and options of the operator menu (in Visual Pinball\User directory)
Const B2STableName="KickOff_1977"
Const LMEMTableConfig="LMEMTables.txt"
Const LMEMShadowConfig="LMEMShadows.txt"
Dim EnableBallShadow
Dim EnableFlipperShadow
Const cGameName = "KickOff_1977"

'* this value adjusts score motor behavior - 0 allows you to continue scoring while the score motor is running - 1 sets score motor to behave more like a real EM
Const ScoreMotorAdjustment=1

'* this is a debug setting to use an older scoring routine vs a newer score routine - don't change this value
Const ScoreAdditionAdjustment=1

dim ScoreChecker
dim CheckAllScores
dim sortscores(4)
dim sortplayers(4)
Dim B2SOn
Dim B2SFrameCounter
Dim BackglassBallFlagColor
Dim TextStr,TextStr2,TiltEndsGame
Dim i
Dim obj
Dim Points210counter
Dim Points500counter
Dim Points1000counter
Dim Points2000counter
Dim BallsPerGame ' 3 or 5
Dim InProgress
Dim BallInPlay
Dim CreditsPerCoin
Dim Score100K(4)
Dim Score(4)
Dim ScoreDisplay(4)
Dim HighScorePaid(4)
Dim HighScore
Dim HighScoreReward
Dim Credits
Dim Match ' 10 to 100 (100=0)
Match = 0
Dim Replay1
Dim Replay2
Dim Replay3
Dim Replay4
Dim Replay1Paid(4)
Dim Replay2Paid(4)
Dim Replay3Paid(4)
Dim Replay4Paid(4)
Dim TableTilted
Dim TiltCount
Dim lowTargetLit 'true if the low target is lit
lowTargetLit = false
Dim toplanesLit 'true if lights of top lanes are lit
toplanesLit = false

Dim OperatorMenu

Dim Ones
Dim Tens
Dim Hundreds
Dim Thousands

Dim CurrentPlayer
Dim NbPlayers

Dim rst
Dim bonuscountdown
Dim TempMultiCounter
dim TempPlayerup
dim RotatorTemp

Dim LastChime10
Dim LastChime100
Dim LastChime1000

Dim Score10
Dim Score100

Dim targettempscore
Dim CardLightOption
Dim HorseshoeCounter
Dim DropTargetCounter

Dim LeftTargetCounter
Dim RightTargetCounter

Dim MotorRunning
Dim Replay1Table(15)
Dim Replay2Table(15)
Dim Replay3Table(15)
Dim Replay4Table(15)
Dim ReplayTableSet
Dim ReplayLevel
Dim ReplayTableMax

Dim BonusSpecialThreshold
Dim AdvanceLightCounter
dim bonustempcounter

dim blinkingLogoB2S

Dim LStep2, xx

Dim ReelCounter
Dim BallCounter
Dim BallReelAddStart(12)
Dim BallReelAddFrames(12)
Dim BallReelDropStart(12)
Dim BallReelDropFrames(12)

Dim EightLit

Dim X
Dim Y
Dim Z
Dim AddABall
Dim AddABallFrames
Dim DropABall
Dim DropABallFrames
Dim CurrentFrame
Dim BonusMotorRunning
Dim QueuedBonusAdds
Dim QueuedBonusDrops

Dim ChimesOn

Dim TempLightTracker

Dim TargetLeftFlag
Dim TargetCenterFlag
Dim TargetRightFlag

Dim SpecialLightsFlag
Dim specialAwarded
specialAwarded = false

Dim ZeroToNineUnit

Dim StarTriggerValue

Dim mHole,mHole2

Dim TargetSequenceCompleted,SpecialLightCounter,HighComplete,LowComplete
Dim ArrowTargetCounter1, ArrowTargetCounter2, ArrowTargetCounter3, ArrowTargetCounter4

Dim Count,Count1,Count2,Count3,Reset,VelX,VelY,BallSpeed,SpecialOption,ScoreMotorStepper
Dim LightsBonusCounter, LightsBallCounter, LightsFieldCounter
LightsBallCounter = -1 ' -1 to 9 (-1 = no lights on, increased with the spinner and others things). Game starts at -1.
LightsBonusCounter = 0 ' 0 to 14 (0 = 1000 points, 14 = 15000 points). Game starts at 1000 points bonus.
LightsFieldCounter = 0 ' 0 to 10 (0 = goal away, 10 = goal home). Game starts at 0.
Dim ArrowLightIsHome ' Toggle arrow light for home/away player. Away light is on for player 2 and 4

Sub Table1_Init()
  If MessagesDebugEnable Then
    MessagesDebugClean()
  End If

  If Table1.ShowDT = false then
    For each obj in DesktopCrap
      obj.visible=False
    next
  End If

  OperatorMenuBackdrop.image = "PostitBL"
  For XOpt = 1 to MaxOption
    Eval("OperatorOption"&XOpt).image = "PostitBL"
  next

  For XOpt = 1 to 256
    Eval("Option"&XOpt).image = "PostItBL"
  next

  LoadEM
  LoadLMEMConfig2

  Count=0
    Count1=0
    Count2=0
  Count3=0
    Reset=0
  ZeroToNineUnit=Int(Rnd*10)

  EightLit=0

  BallCounter=0
  ReelCounter=0
  AddABall=0
  DropABall=0
  HideOptions
  SetupReplayTables
  PlasticsOff
  OperatorMenu=0
  HighScore=0
  lowTargetLit = false
  toplanesLit = false
  MotorRunning=0
  HighScoreReward=1
  ChimesOn=1 ' active Xylophone
  TiltEndsGame=0
  specialAwarded = false
  BackglassBallFlagColor=1

  ' default values, but loaded with loadhs (see User directory\KickOff_77VPX.txt)
  Credits=0
  BallsPerGame=5 '3 or 5
  ReplayLevel=1
  SpecialOption=10 ' 0 = 'No Special Award', 1 = 'Bonus at 10,000'...
  CardLightOption=1 ' 0 = 'No Extra Ball Awards', 1 = 'Bonus at 10k, 12k, 15k'...

  ' Load high scores and options (see operator menu)
  loadhs
  if HighScore=0 then HighScore=5000

  TableTilted=false

  GameOverReel.SetValue(1)
  BallInPlayReel.SetValue(0)
  TiltReel.SetValue(1)

  For each obj in PlayerHuds
    obj.state=0
  next
  For each obj in PlayerScores
    obj.ResetToZero
  next

  Replay1=Replay1Table(ReplayLevel)
  Replay2=Replay2Table(ReplayLevel)
  Replay3=Replay3Table(ReplayLevel)
  Replay4=Replay4Table(ReplayLevel)

  InstructCard.image="IC"

  RefreshReplayCard

  CurrentFrame=0

  AdvanceLightCounter=0

  NbPlayers=0
  RotatorTemp=1
  InProgress=false

  ScoreText.text=HighScore

  If B2SOn Then
    Controller.B2SSetMatch 0
    Controller.B2SSetScoreRolloverPlayer1 0
    Controller.B2SSetScoreRolloverPlayer2 0
    Controller.B2SSetScoreRolloverPlayer3 0
    Controller.B2SSetScoreRolloverPlayer4 0

    'Controller.B2SSetScore 3,HighScore
    Controller.B2SSetTilt 1
    Controller.B2SSetCredits Credits
    Controller.B2SSetGameOver 1
  End If

  for i=1 to 4
    CurrentPlayer = i
    If B2SOn Then
      Controller.B2SSetScorePlayer CurrentPlayer, 0
    End If
  next
  InitPauser5.enabled=true

  If Credits > 0 Then DOF 129, DOFOn
End Sub

Sub Table1_exit()
  'Save High scores
  savehs
  SaveLMEMConfig
  SaveLMEMConfig2
  If B2SOn Then Controller.Stop
end sub

' Blinking logo on B2S
Sub BlinkingLogo_timer
  if B2SOn and activeBlinkingLogoB2S then
    if blinkingLogoB2S=1 then
      Controller.B2SSetData 70,0
      blinkingLogoB2S=0
    else
      Controller.B2SSetData 70,1
      blinkingLogoB2S=1
    end if
  end if
end sub

Sub Table1_KeyDown(ByVal keycode)

  ' GNMOD
  if EnteringInitials then
    CollectInitials(keycode)
    exit sub
  end if


  if EnteringOptions then
    CollectOptions(keycode)
    exit sub
  end if



  If keycode = PlungerKey Then
    Plunger.PullBack
    PlungerPulled = 1
  End If

  ' Operator menu
  if keycode = LeftFlipperKey and InProgress = false then
    OperatorMenuTimer.Enabled = true
  end if
  ' END GNMOD


  If keycode = LeftFlipperKey and InProgress=true and TableTilted=false Then
    LeftFlipper.RotateToEnd
    LeftFlipper2.RotateToEnd
    PlaySoundAtVol SoundFXDOF("FlipperUp",101,DOFOn, DOFFlippers), LeftFlipper, 1
    PlayLoopSoundAtVol "buzzL", LeftFlipper, 1
  End If

  If keycode = RightFlipperKey  and InProgress=true and TableTilted=false Then
    RightFlipper.RotateToEnd
    RightFlipper2.RotateToEnd

    PlaySoundAtVol SoundFXDOF("FlipperUp",102,DOFOn, DOFFlippers), RightFlipper, 1
    PlaySoundAtVol "FlipperUp", RightFlipper2, 1
    PlayLoopSoundAtVol "buzz", RightFlipper, 1

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
    TiltCount=2
    TiltIt
  End If


  If keycode = AddCreditKey or keycode = 4 then
    If B2SOn Then
      'Controller.B2SSetScorePlayer6 HighScore
    End If
    PlaySoundAtVol "coinin", Drain, 1
    AddCredit
  end if

   if keycode = 5 then
    PlaySoundAtVol "coinin", Drain, 1
    AddCredit
    keycode= StartGameKey
  end if


   if keycode = StartGameKey and Credits>0 and InProgress=true and NbPlayers>0 and NbPlayers<4 and BallInPlay<2 then
    Credits=Credits-1
    If Credits <1 Then DOF 129, DOFOff:CreditLight.state=0
    CreditsReel.SetValue(Credits)
    EMReel7.SetValue(Credits)
    NbPlayers=NbPlayers+1
    SetCanPlay NbPlayers,1

    playsound "click"
    If B2SOn Then
      Controller.B2SSetCredits Credits
    End If
    end if

  if keycode=StartGameKey and Credits>0 and InProgress=false and NbPlayers=0 and EnteringOptions = 0 then
'GNMOD
    OperatorMenuTimer.Enabled = false
'END GNMOD
    Credits=Credits-1
    If Credits < 1 Then DOF 129, DOFOff:CreditLight.state=0
    CreditsReel.SetValue(Credits)
    NbPlayers=1
    SetCanPlay NbPlayers,1
    MatchReel.SetValue(0)
    CurrentPlayer=1
    RolloverReel1.SetValue(0)
    RolloverReel2.SetValue(0)
    RolloverReel3.SetValue(0)
    RolloverReel4.SetValue(0)
    playsound "startup_norm"
    TempPlayerUp=CurrentPlayer
'   PlayerUpRotator.enabled=true
    GameOverReel.SetValue(0)
    rst=0
    BallInPlay=1
    InProgress=true
    resettimer.enabled=true

    If B2SOn Then
      Controller.B2SSetTilt 0
      Controller.B2SSetGameOver 0
      Controller.B2SSetMatch 0
      Controller.B2SSetCredits Credits
      'Controller.B2SSetScore 3,HighScore
      Controller.B2SSetPlayerUp 1
      'Controller.B2SSetBallInPlay BallInPlay
      Controller.B2SSetScoreRolloverPlayer1 0
    End If
    If Table1.ShowDT = True then
      For each obj in PlayerScores
'       obj.ResetToZero
        obj.Visible=true
      next
      For each obj in PlayerHuds
        obj.state=0
      next
      PlayerHuds(CurrentPlayer-1).state=1
    end If
  end if
End Sub

Sub Table1_KeyUp(ByVal keycode)

  ' GNMOD
  if EnteringInitials then
    exit sub
  end if

  If keycode = PlungerKey Then

    if PlungerPulled = 0 then
      exit sub
    end if

    PlaySoundAtVol"plungerrelease", Plunger, 1
    Plunger.Fire
  End If

  if keycode = LeftFlipperKey then
    OperatorMenuTimer.Enabled = false
  end if

  ' END GNMOD

  If keycode = LeftFlipperKey and InProgress=true and TableTilted=false Then
    LeftFlipper.RotateToStart
    LeftFlipper2.RotateToStart
    PlaySoundAtVol SoundFXDOF("FlipperDown",101,DOFOff,DOFFlippers), LeftFlipper, 1
    PlaySoundAtVol "FlipperDown", LeftFlipper2, 1
    StopSound "buzzL"

  End If

  If keycode = RightFlipperKey and InProgress=true and TableTilted=false Then
    RightFlipper.RotateToStart
    RightFlipper2.RotateToStart
    PlaySoundAtVol SoundFXDOF("FlipperDown",102,DOFOff,DOFFlippers), RightFlipper, 1
    PlaySoundAtVol "FlipperDown",RightFlipper2, 1
    StopSound "buzz"

  End If
End Sub


' Nb players can play
' state : 0 or 1
Sub SetCanPlay(numberPlayers, state)
  'Reset lights
  for each obj in CanPlayLights
    obj.state = 0
  next
  CanPlayLights(NbPlayers-1).state = state

  If B2SOn Then
    Controller.B2SSetCanPlay numberPlayers
  End If
End Sub

' Ball lost
Sub Drain_Hit()
  Drain.DestroyBall
  PlaySoundAtVol "fx_drain", Drain, 1
  DOF 137, DOFPulse
  Pause4Bonustimer.enabled=true
End Sub

Sub Pause4Bonustimer_timer
  Pause4Bonustimer.enabled=0
  ScoreBonusPoints
End Sub

' Flipper Logos
Sub UpdateFlipperLogos_Timer
  PGate.Rotz = (Gate.CurrentAngle*.75) + 25
  FlipperLSh.RotZ = LeftFlipper.currentangle
  FlipperRSh.RotZ = RightFlipper.currentangle
End Sub

' Walls
Sub RubberWallSwitches_Hit(idx)
  if TableTilted=false then
    AddScore(10)
  end if
end Sub

' Star rollover hit
Sub StarRollover_Hit()
  If TableTilted then Exit Sub
  If toplanesLit = true then
    ' Active low target
    lowTargetLit = true
    TargetLowLight.state = LightStateOn
    TargetLowLight2.state = LightStateOn
  else
    ' Active 1st and 3th top lights
    LightsLane(0).State = LightStateOn
    LightsLane(2).State = LightStateOn
    toplanesLit = true
  end if
End Sub

' Bumper left
Sub Bumper1_Hit
  If TableTilted=false then
    PlaySoundAtVol SoundFXDOF("bumper1",106,DOFPulse,DOFContactors), ActiveBall, 1
    DOF 124, DOFPulse
    AddScore(100)
    end if
End Sub

' Bumper right
Sub Bumper2_Hit
  If TableTilted=false then
    PlaySoundAtVol SoundFXDOF("bumper1",105,DOFPulse,DOFContactors), ActiveBall, 1
    DOF 123, DOFPulse
    AddScore(100)
    end if
End Sub

' Bumper center
Sub Bumper3_Hit
  If TableTilted=false then
    PlaySoundAtVol SoundFXDOF("bumper1",107,DOFPulse,DOFContactors), ActiveBall, 1
    DOF 125, DOFPulse
    AddScore(10) '10 points
    SwitchTopLaneLights
    end if
End Sub

' Trigger rollover lanes
Sub TriggerCollection_Hit(idx)
  DOF 110 + idx, DOFPulse
  If TableTilted=false then
    'The ball can be advanced three positions by either going thru a top lane when lit, or hitting the lower left target when lit.
    Select Case idx
      case 0: 'TriggerTop1
        If LightToplane1.State = LightStateOn then
          ' A video on a the real machine explains that score is different with the number of balls per game
          If BallsPerGame = 3 Then
            AddScore(3000)
          End If
          If BallsPerGame = 5 Then
            AddScore(300)
          End If
          AdvanceField(3)
          for i = 0 to 3
            LightsLane(i).State = LightStateOff
          next
          toplanesLit = false
        Else
          AddScore(500)
          AdvanceBall(5)
        End If
      case 1: 'TriggerTop2
        If LightToplane2.State = LightStateOn then
          ' A video on a the real machine explains that score is different with the number of balls per game
          If BallsPerGame = 3 Then
            AddScore(3000)
          End If
          If BallsPerGame = 5 Then
            AddScore(300)
          End If
          AdvanceField(3)
          for i = 0 to 3
            LightsLane(i).State = LightStateOff
          next
          toplanesLit = false
        Else
          AddScore(500)
          AdvanceBall(5)
        End If
      case 2: 'TriggerTop3
        If LightToplane3.State = LightStateOn then
          ' A video on a the real machine explains that score is different with the number of balls per game
          If BallsPerGame = 3 Then
            AddScore(3000)
          End If
          If BallsPerGame = 5 Then
            AddScore(300)
          End If
          AdvanceField(3)
          for i = 0 to 3
            LightsLane(i).State = LightStateOff
          next
          toplanesLit = false
        Else
          AddScore(500)
          AdvanceBall(5)
        End If
      case 3: 'TriggerTop4
        If LightToplane4.State = LightStateOn then
          ' A video on a the real machine explains that score is different with the number of balls per game
          If BallsPerGame = 3 Then
            AddScore(3000)
          End If
          If BallsPerGame = 5 Then
            AddScore(300)
          End If
          AdvanceField(3)
          for i = 0 to 3
            LightsLane(i).State = LightStateOff
          next
          toplanesLit = false
        Else
          AddScore(500)
          AdvanceBall(5)
        End If
      case 4: 'TriggerLeftOutlane
        AddScore(1000)
      case 5: 'TriggerRightOutlane
        AddScore(1000)
    end select
  end if
end sub

' Drop targets
Sub DropTargets_Dropped(idx)
  If TableTilted=false then
    'DOF 130+idx,DOFPulse
    DOF 111, DOFPulse
    PlaySoundAtVol SoundFX("fx_target", DOFDropTargets), ActiveBall, 1
    AddScore(500)
    ' check if 7 targets are droppped : we get a goal or a special
    Dim nbDropped
    nbDropped = 0
    for each obj in DropTargets
      If obj.IsDropped Then
        nbDropped = nbDropped + 1
      End If
    next
    ' Knocking down all the drop targets will either score a goal or a special.
    ' When the drop targets are lit for a goal teh last drop target down will score teh lit goal value.
    ' The goal value reads 5000 points, the second goal is worth an extra ball and the third goal is worth a special.
    If nbDropped >= 7 then
      if LightSpecialHigh.State = LightStateOn then
        AddSpecial
      end if
      if LightGoal.State = LightStateOn then
        ScoreGoal
        LightSpecialHigh.State = LightStateOn
        LightGoal.State = LightStateOff
      end if
      ResetDropTargets
    End If
    AdvanceBall(5)
  End if
End sub

' Switch top lanes lights
Sub SwitchTopLaneLights
  if toplanesLit = true then
    if LightsLane(0).State = LightStateOn then
      LightsLane(1).State = LightStateOn
      LightsLane(3).State = LightStateOn
      LightsLane(0).State = LightStateOff
      LightsLane(2).State = LightStateOff
    else
      LightsLane(1).State = LightStateOff
      LightsLane(3).State = LightStateOff
      LightsLane(0).State = LightStateOn
      LightsLane(2).State = LightStateOn
    end if
  end if
End Sub

' Spinner
Sub Spinner1_Spin()
  If TableTilted=false and MotorRunning=0 then
    DOF 122, 2
    AddScore(100)
    AdvanceBall(1)
  end if
end Sub

' Active low target if rubber hit
Sub ScoringWall006_Hit()
  lowTargetLit = true
  TargetLowLight.state = LightStateOn
  TargetLowLight2.state = LightStateOn
End Sub

' Target high hit
Sub TargetHigh_Hit()
  DOF 131,DOFPulse
  If TableTilted then Exit Sub
  AddScore(1000)
  ' The ball is also advanced once by hitting the upper left target
  AdvanceField(1)
End Sub

' Target low hit
Sub TargetLow_Hit()
  DOF 132,DOFPulse
  if TableTilted then Exit Sub
  AddScore(500)
  ' Advance ball or field if the target low is lit
  ' The ball can be advanced three positions by either going thru a top lane when lit, or hitting the lower left target when lit.
  if lowTargetLit = false then
    AdvanceField(3)
  else
    AdvanceBall(5)
  end if
End Sub

' Increase ball and lights
Sub AdvanceBall(ballDigit)
  If ballDigit = 0 Then Exit Sub
  If TableTilted then Exit Sub
  dim i,j
  for i = 1 to ballDigit
    If TableTilted then Exit Sub
    ' LightsBallCounter : -1 to 9 (-1 = no lights on)
    LightsBallCounter = LightsBallCounter + 1
    If LightsBallCounter >= 0 and LightsBallCounter < 10 Then
      LightsBall(LightsBallCounter).State = LightStateOn
    End If
    ' Each 10 spins, advance field
    If LightsBallCounter >= 9 then
      ' Increase yards in field
      AdvanceField(1)
      for j = 0 to 9
        LightsBall(j).State = LightStateOff
      next
      LightsBallCounter = -1
    end if
  next
  PlaySound "Light"
End Sub

' Increase ball in field and lights (11 lights)
Sub AdvanceField(fieldDigit)
  if TableTilted then Exit Sub
  dim k, digit
  for k = 1 to fieldDigit
    ' we advance to the left (away direction) or to the right (home direction)
    if ArrowLightIsHome = false then
      digit = (-1)
    else
      digit = 1
    end if
    ' light off
    'LightsField(LightsFieldCounter).State = LightStateOff
    ' Bug with the previous line : with the player 2 the light selected is not correct ! I light off all the lights :
    for each obj in LightsField
      obj.State = LightStateOff
    next
    LightsFieldCounter = LightsFieldCounter + digit
    If LightsFieldCounter < 0 then
      LightsFieldCounter = 0
    End If
    If LightsFieldCounter > 10 then
      LightsFieldCounter = 10
    End If
    ' light on
    PlaySound "Light"
    'LightsField(LightsFieldCounter).State = LightStateOn
    ' Bug with the previous line : with the player 2 the light selected is not correct ! I light on directly the light with his name :
    Select case LightsFieldCounter
      case 0: LightGoalRight.State = LightStateOn
      case 1: Light10Left.State = LightStateOn
      case 2: Light20Left.State = LightStateOn
      case 3: Light30Left.State = LightStateOn
      case 4: Light40Left.State = LightStateOn
      case 5: Light50.State = LightStateOn
      case 6: Light40Right.State = LightStateOn
      case 7: Light30Right.State = LightStateOn
      case 8: Light20Right.State = LightStateOn
      case 9: Light10Right.State = LightStateOn
      case 10: LightGoalRight.State = LightStateOn
    End select

    ' The bonus score advances one step each time the ball reaches a yellow ball position on the soccer field (20, 40 yards...)
    if LightsFieldCounter = 2 or LightsFieldCounter = 4 or LightsFieldCounter = 6 or LightsFieldCounter = 8 or LightsFieldCounter = 10 then
      AdvanceBonus
    end if
    if ArrowLightIsHome = true and LightsFieldCounter = 10 then
      ScoreGoal
    end if
    if ArrowLightIsHome = false and LightsFieldCounter = 0 then
      ScoreGoal
    end if
  next
End Sub

'Increase bonus
sub AdvanceBonus
  'LightsBonusCounter : 0 to 14 (0 = 1000 points, 14 = 15000 points)
  if LightsBonusCounter < 14 then
    DoBonusLight LightsBonusCounter, LightStateOff
    LightsBonusCounter = LightsBonusCounter + 1
    DoBonusLight LightsBonusCounter,LightStateOn
  end if
  ' Special at 15000 points bonus
  if LightsBonusCounter = 14 and not specialAwarded then
    AddSpecial
    specialAwarded = true
  end if
End sub

'Change the state of the indicated bonus light
Sub DoBonusLight(bonusCounter, newState)
  'MessageDebug "DoBonusLight, bonusCounter=" & CStr(bonusCounter)
  'LightsBonusCounter : 0 to 14 (0 = 1000 points, 14 = 15000 points). Game starts at 1000 points bonus.
  if bonusCounter < 10 then
    LightsBonus(bonusCounter).State = newState
  end if
  if bonusCounter >= 10 then
    ' 2 lights are on
    LightsBonus(9).state = newState
    LightsBonus(LightsBonusCounter-10).state = newState
  end if
  Playsound "Light"
End sub

' Add score from ball lights
Sub ScoreGoal
  if TableTilted then Exit Sub
  dim i
  if LightSpecialLow.State = LightStateOn then
    AddSpecial
    LightSpecialLow.State = LightStateOff
  end if
  if LightExtraBall.State = LightStateOn then
    LightShootAgain.State = LightStateOn
    LightExtraBall.State = LightStateOff
    LightSpecialLow.State = LightStateOn
  end if
  if Light5000.State = LightStateOn then
    AddScore (5000)
    Light5000.State = LightStateOff
    LightExtraBall.State = LightStateOn
  end if
  LightsBallCounter = -1
  for i = 0 to 9
    LightsBall(i).State = LightStateOff
  next
  for i = 0 to 10
    LightsField(i).State = LightStateOff
  next
  ' Reset lights on field
  ' LightsFieldCounter : 0 to 10 (0 = goal away, 10 = goal home)
  if ArrowLightIsHome = false then
    LightsField(0).State = LightStateOff
    LightsField(10).State = LightStateOn
    LightsFieldCounter = 10
  end if
  if ArrowLightIsHome = true then
    LightsField(0).State = LightStateOn
    LightsField(10).State = LightStateOff
    LightsFieldCounter = 0
  end if
End Sub

' Add a credit with Special
Sub AddSpecial()
  PlaySound"knocker"
  DOF 140, DOFPulse

  specialAwarded = true
  Credits=Credits+1
  CreditLight.state=1
  DOF 129, DOFOn
  ' Max 15 credits on an EM machine
  if Credits>15 then Credits=15
  If B2SOn Then
    Controller.B2SSetCredits Credits
  End If
  CreditsReel.SetValue(Credits)
End Sub

' Add a credit
Sub AddCredit()
  PlaySound"click"
  Credits=Credits+1
  :CreditLight.state=1
  ' Max 15 credits on an EM machine
  DOF 129, DOFOn
  if Credits>15 then Credits=15
  If B2SOn Then
    Controller.B2SSetCredits Credits
  End If
  CreditsReel.SetValue(Credits)
End Sub

' Advance bonus and lights
sub IncreaseBonusPoints
  'MessageDebug "LightsBonusCounter : " & CStr(LightsBonusCounter)
  if LightsBonusCounter<15 then 'LightsBonusCounter : 0 to 14 (0 = 1000 points, 14 = 15000 points)
    If LightsBonusCounter<10 then
      LightsBonus(LightsBonusCounter).state=0
      LightsBonusCounter=LightsBonusCounter+1
      LightsBonus(LightsBonusCounter).state=1
    else
      LightsBonus(LightsBonusCounter-10).state=0
      LightsBonusCounter=LightsBonusCounter+1
      ' 2 lights are on
      LightsBonus(9).state=1
      LightsBonus(LightsBonusCounter-10).state=1
    end if
    'MessageDebug "LightsBonusCounter after : " & CStr(LightsBonusCounter)
  end if
end sub

' Launch score motor for add bonus on score
sub ScoreBonusPoints
  ScoreMotorStepper=0 ' 0 to 7
  ' CollectBonusPoints is a Timer for simulate a EM score motor
  CollectBonusPoints.interval=135
  CollectBonusPoints.enabled=1
end sub

' Launch score motor for add bonus on score
sub CollectBonusPoints_timer
  'MessageDebug "CollectBonusPoints"
  if MotorRunning=1 then
    exit sub
  end if
  if LightsBonusCounter<1 then
    ' Bonus completly added, next ball launched
    CollectBonusPoints.enabled=0
    NextBallDelay.enabled=true
  else
    ' If light BonusX2 is lit
    If LightBonusX2.state=1 then
      Select case ScoreMotorStepper ' 0 to 7
        case 0,1,3,4:
          AddScore(1000)
        case 2,5:
          PlaySound"BallyClunk"
          ' LightsBonusCounter : 0 to 14 (0 = 1000 points, 14 = 15000 points).
          If LightsBonusCounter>9 then
            LightsBonus(LightsBonusCounter-10).state=0
          else
            LightsBonus(LightsBonusCounter).state=0
          end if
          ' decrease bonus counter
          LightsBonusCounter=LightsBonusCounter-1
          if LightsBonusCounter>=0 then
            If LightsBonusCounter<=9 then
              LightsBonus(LightsBonusCounter).state=1
            else
              LightsBonus(9).state=1
              LightsBonus(LightsBonusCounter-10).state=1
            end if
          end if
      end select
      ScoreMotorStepper=ScoreMotorStepper+1
      If ScoreMotorStepper>7 then
        ScoreMotorStepper=0
      end if
    else
      Select Case ScoreMotorStepper
        case 0,3:
          AddScore(1000)
        case 1,4:
          PlaySound"BallyClunk"
        case 2,5:
          PlaySound"BallyClunk"
          ' LightsBonusCounter : 0 to 14 (0 = 1000 points, 14 = 15000 points).
          If LightsBonusCounter>9 then
            LightsBonus(LightsBonusCounter-10).state=0
          else
            LightsBonus(LightsBonusCounter).state=0
          end if
          ' decrease bonus counter
          LightsBonusCounter=LightsBonusCounter-1
          if LightsBonusCounter>=0 then
            If LightsBonusCounter<=9 then
              LightsBonus(LightsBonusCounter).state=1
            else
              LightsBonus(9).state=1
              LightsBonus(LightsBonusCounter-10).state=1
            end if
          end if
      end select
      ScoreMotorStepper=ScoreMotorStepper+1
      If ScoreMotorStepper>7 then
        ScoreMotorStepper=0
      end if
    end if
  end if
end sub

' Rubbers sounds
Sub aRubber_Posts_Hit(idx):PlaySoundAtBall "fx_rubber_post":End Sub

Sub LightsOut
  StopSound "buzz"
  StopSound "buzzL"
end sub

Sub ResetBalls()
  TempMultiCounter=BallsPerGame-BallInPlay
  TableTilted=false
  TiltReel.SetValue(0)
  If B2Son then
    Controller.B2SSetTilt 0
  end if
  IncreaseBonusPoints
  LightBonusX2.state=0
  If BallInPlay=5 OR BallInPlay=3 then
    LightBonusX2.state=1
  end if

  lowTargetLit = false
  TargetLowLight.state = LightStateOff
  TargetLowLight2.state = LightStateOff
  toplanesLit = false
  specialAwarded = false

  PlasticsOn
  'CreateBallID BallRelease
  Ballrelease.CreateSizedBall 25
    Ballrelease.Kick 40,7
  DOF 128, DOFPulse
  BallInPlayReel.SetValue(BallInPlay)
  InstructCard.image="IC"

  ' Drop targets
  ResetDropTargets

  ' Remark : on a video of the real machine, it seems that the ball counter is not resetted after a ball lost. And with multiple players ?
  LightsBallCounter = -1
  LightsBonusCounter = 0
  if ArrowLightIsHome = false then
    LightsFieldCounter = 10
  else
    LightsFieldCounter = 0
  end if

  ResetLights
End Sub

' Reset drop targets
Sub ResetDropTargets
  for each obj in DropTargets
    obj.IsDropped = false
  next
End Sub


Sub ResetLights
  for each obj in LightsLane
    obj.state = LightStateOff
  next
  for each obj in LightsBall
    obj.state = LightStateOff
  next
  for each obj in LightsField
    obj.state = LightStateOff
  next
  for each obj in LightsBonus
    obj.state = LightStateOff
  next
  for each obj in CanPlayLights
    obj.state = LightStateOff
  next
  LightBonusX2.state = LightStateOff
  LightShootAgain.state = LightStateOff
  LightSpecialHigh.state = LightStateOff
  LightSpecialLow.state = LightStateOff
  If lowTargetLit Then
    TargetLowLight.state = LightStateOn
    TargetLowLight2.state = LightStateOn
  else
    TargetLowLight.state = LightStateOff
    TargetLowLight2.state = LightStateOff
  End if

  'Default lights
  StarRolloverLight.state = LightStateOn
  LightGoal.state = LightStateOn
  Light5000.state = LightStateOn
  LightBonus1.state = LightStateOn

  ' Last ball : bonus X2
  If BallInPlay = BallsPerGame then
    LightBonusX2.state = LightStateOn
  End if

  ' LightsFieldCounter : 0 to 10 (0 = goal away, 10 = goal home)
  if CurrentPlayer = 2 or CurrentPlayer = 4 then
    ArrowLightIsHome = false
    LightsFieldCounter = 10
    LightHome.State = LightStateOff
    LightAway.State = LightStateOn
  else
    ArrowLightIsHome = true
    LightsFieldCounter = 0
    LightHome.State = LightStateOn
    LightAway.State = LightStateOff
  end if
  LightsField(LightsFieldCounter).State = LightStateOn
End Sub

' Laucnh the reset of score reels
sub resettimer_timer
    rst=rst+1
  if rst>1 and rst<12 then
    ResetReelsToZero(1)
    ResetReelsToZero(2)
  end if

    if rst=16 then
    playsound "StartBall1"
    end if
    if rst=18 then
    newgame
    resettimer.enabled=false
    end if
end sub

' Reset of score reels
Sub ResetReelsToZero(reelzeroflag)
  dim d1(5)
  dim d2(5)
  dim scorestring1, scorestring2

  If reelzeroflag=1 then
    scorestring1=CStr(Score(1))
    scorestring2=CStr(Score(2))
    scorestring1=right("00000" & scorestring1,5)
    scorestring2=right("00000" & scorestring2,5)
    for i=0 to 4
      d1(i)=CInt(mid(scorestring1,i+1,1))
      d2(i)=CInt(mid(scorestring2,i+1,1))
    next
    for i=0 to 4
      if d1(i)>0 then
        d1(i)=d1(i)+1
        if d1(i)>9 then d1(i)=0
      end if
      if d2(i)>0 then
        d2(i)=d2(i)+1
        if d2(i)>9 then d2(i)=0
      end if

    next
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
    scorestring1=right("00000" & scorestring1,5)
    scorestring2=right("00000" & scorestring2,5)
    for i=0 to 4
      d1(i)=CInt(mid(scorestring1,i+1,1))
      d2(i)=CInt(mid(scorestring2,i+1,1))
    next
    for i=0 to 4
      if d1(i)>0 then
        d1(i)=d1(i)+1
        if d1(i)>9 then d1(i)=0
      end if
      if d2(i)>0 then
        d2(i)=d2(i)+1
        if d2(i)>9 then d2(i)=0
      end if

    next
    Score(3)=(d1(0)*10000) + (d1(1)*1000) + (d1(2)*100) + (d1(3)*10) + d1(4)
    Score(4)=(d2(0)*10000) + (d2(1)*1000) + (d2(2)*100) + (d2(3)*10) + d2(4)
    If B2SOn Then
      Controller.B2SSetScorePlayer 3, Score(3)
      Controller.B2SSetScorePlayer 4, Score(4)
    End If
    PlayerScores(2).SetValue(Score(3))
    PlayerScoresOn(2).SetValue(Score(3))
    PlayerScores(3).SetValue(Score(4))
    PlayerScoresOn(3).SetValue(Score(4))

  end if

end sub

' Launch the next ball after a delay
sub NextBallDelay_timer()
  NextBallDelay.enabled=false
  nextball
end sub

sub newgame
  InProgress=true
  queuedscore=0
  for i = 1 to 4
    Score(i)=0
    Score100K(1)=0
    HighScorePaid(i)=false
    Replay1Paid(i)=false
    Replay2Paid(i)=false
    Replay3Paid(i)=false
    Replay4Paid(i)=false
  next
  If B2SOn Then
    Controller.B2SSetTilt 0
    Controller.B2SSetGameOver 0
    Controller.B2SSetMatch 0
'   Controller.B2SSetScorePlayer1 0
'   Controller.B2SSetScorePlayer2 0
'   Controller.B2SSetScorePlayer3 0
'   Controller.B2SSetScorePlayer4 0
    Controller.B2SSetBallInPlay BallInPlay
  End if
  for each obj in CardsOn
    obj.state=1
  next
  for each obj in CardsOff
    obj.state=0
  next

  ' LightsBallCounter : -1 to 9 (-1 = no lights on)
  LightsBallCounter = -1
  LightsBonusCounter = 0
  LightsFieldCounter = 0

  LowComplete=false
  HighComplete=false

  ResetBalls 'ResetLights in ResetBalls procedure
End sub

sub nextball
  If LightShootAgain.state=1 then
    LightShootAgain.state=0
    If B2SOn Then
      Controller.B2SSetShootAgain 0
    end if
  else
    CurrentPlayer=CurrentPlayer+1
  end if
  If CurrentPlayer>NbPlayers Then
    BallInPlay=BallInPlay+1
    If BallInPlay>BallsPerGame then
      PlaySound("MotorLeer")
      InProgress=false

      If B2SOn Then
        Controller.B2SSetGameOver 1
        Controller.B2SSetPlayerUp 0
        Controller.B2SSetBallInPlay 0
        SetCanPlay 0,0

      End If
      For each obj in PlayerHuds
        obj.state=0
      next
      If Table1.ShowDT = True then
        For each obj in PlayerScores
          obj.visible=1
        Next

      end If
      GameOverReel.SetValue(1)
      InstructCard.image="IC"

      BallInPlayReel.SetValue(0)
      LeftFlipper.RotateToStart
      RightFlipper.RotateToStart
      LightsOut
      PlasticsOff
      checkmatch
      CheckHighScore
      NbPlayers=0
      HighScoreTimer.interval=100
      HighScoreTimer.enabled=True

    Else
      CurrentPlayer=1
      If B2SOn Then
        Controller.B2SSetPlayerUp CurrentPlayer
        Controller.B2SSetBallInPlay BallInPlay
      End If
'     PlaySound("RotateThruPlayers")
      TempPlayerUp=CurrentPlayer
'     PlayerUpRotator.enabled=true
      PlayStartBall.enabled=true
      For each obj in PlayerHuds
        obj.state=0
      next
      If Table1.ShowDT = True then
        For each obj in PlayerScores
          obj.visible=1
        Next

        PlayerHuds(CurrentPlayer-1).state=1

      end If

      ResetBalls
    End If
  Else
    If B2SOn Then
      Controller.B2SSetPlayerUp CurrentPlayer
      Controller.B2SSetBallInPlay BallInPlay
    End If
'   PlaySound("RotateThruPlayers")
    TempPlayerUp=CurrentPlayer
'   PlayerUpRotator.enabled=true
    PlayStartBall.enabled=true
    For each obj in PlayerHuds
      obj.state=0
    next
    If Table1.ShowDT = True then
      For each obj in PlayerScores
        obj.visible=1
      Next

      PlayerHuds(CurrentPlayer-1).state=1

    end If
    ResetBalls
  End If

End sub

sub CheckHighScore
  Dim playertops
    dim si
  dim sj
  dim stemp
  dim stempplayers
  for i=1 to 4
    sortscores(i)=0
    sortplayers(i)=0
  next
  playertops=0
  for i = 1 to NbPlayers
    sortscores(i)=Score(i)
    sortplayers(i)=i
  next

  for si = 1 to NbPlayers
    for sj = 1 to NbPlayers-1
      if sortscores(sj)>sortscores(sj+1) then
        stemp=sortscores(sj+1)
        stempplayers=sortplayers(sj+1)
        sortscores(sj+1)=sortscores(sj)
        sortplayers(sj+1)=sortplayers(sj)
        sortscores(sj)=stemp
        sortplayers(sj)=stempplayers
      end if
    next
  next
  ScoreChecker=4
  CheckAllScores=1
  NewHighScore sortscores(ScoreChecker),sortplayers(ScoreChecker)
  'Save High scores
  savehs
end sub

' Match at end of the game
sub checkmatch
  if TableTilted=true and TiltEndsGame=1 then
    exit sub
  end if
  Dim tempmatch
  tempmatch=Int(Rnd*10) ' 0 to 9
  Match=tempmatch*10 ' 0 to 90
  MatchReel.SetValue(tempmatch+1)

  If B2SOn Then
    If Match = 0 Then
      Controller.B2SSetMatch 100
    Else
      Controller.B2SSetMatch Match
    End If
  End if
  for i = 1 to NbPlayers
    if Match=(Score(i) mod 100) then
      AddSpecial
    end if
  next
end sub

Sub TiltTimer_Timer()
  if TiltCount > 0 then TiltCount = TiltCount - 1
  if TiltCount = 0 then
    TiltTimer.Enabled = False
  end if
end sub

Sub TiltIt()
    TiltCount = TiltCount + 1
    if TiltCount = 3 then
      TableTilted=True
      If TiltEndsGame=1 then
        BallInPlay=BallsPerGame
        If B2SOn Then
          Controller.B2SSetGameOver 1
          Controller.B2SSetPlayerUp 0
          Controller.B2SSetBallInPlay 0
          SetCanPlay 0,0
        End If
        For each obj in PlayerHuds
          obj.state=0
        next
        GameOverReel.SetValue(1)
        InstructCard.image="IC"

      end if
      PlasticsOff
      LeftFlipper.RotateToStart
      RightFlipper.RotateToStart
      StopSound "buzz"
      StopSound "buzzL"
      TiltReel.SetValue(1)
      If B2Son then
        Controller.B2SSetTilt 1
      end if
    else
      TiltTimer.Interval = 500
      TiltTimer.Enabled = True
    end if

end sub



Sub PlayStartBall_timer()

  PlayStartBall.enabled=false
  PlaySound("StartBall2-5")
end sub

Sub PlayerUpRotator_timer()
    If RotatorTemp<5 then
      TempPlayerUp=TempPlayerUp+1
      If TempPlayerUp>4 then
        TempPlayerUp=1
      end if
      If B2SOn Then
        Controller.B2SSetPlayerUp TempPlayerUp
      End If

    else
      if B2SOn then
        Controller.B2SSetPlayerUp CurrentPlayer
      end if
      PlayerUpRotator.enabled=false
      RotatorTemp=1
    end if
    RotatorTemp=RotatorTemp+1


end sub

'Save High scores
sub savehs
  ' Based on Black's Highscore routines
  Dim FileObj
  Dim ScoreFile
  Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if
  Set ScoreFile=FileObj.CreateTextFile(UserDirectory & HSFileName,True)
    ScoreFile.WriteLine 1
    ScoreFile.WriteLine Credits
    scorefile.writeline BallsPerGame
    scorefile.writeline 0
    scorefile.writeline CardLightOption
    scorefile.writeline SpecialOption
    scorefile.writeline ReplayLevel
    for xx=1 to 5
      scorefile.writeline HSScore(xx)
    next
    for xx=1 to 5
      scorefile.writeline HSName(xx)
    next
    ScoreFile.Close
  Set ScoreFile=Nothing
  Set FileObj=Nothing
end sub

sub SaveLMEMConfig
  Dim FileObj
  Dim LMConfig
  dim temp1
  dim tempb2s
  tempb2s=0
  if B2SOn=true then
    tempb2s=1
  else
    tempb2s=0
  end if
  Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if
  Set LMConfig=FileObj.CreateTextFile(UserDirectory & LMEMTableConfig,True)
  LMConfig.WriteLine tempb2s
  LMConfig.Close
  Set LMConfig=Nothing
  Set FileObj=Nothing

end Sub

sub LoadLMEMConfig
  Dim FileObj
  Dim LMConfig
  dim tempC
  dim tempb2s

    Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if
  If Not FileObj.FileExists(UserDirectory & LMEMTableConfig) then
    Exit Sub
  End if
  Set LMConfig=FileObj.GetFile(UserDirectory & LMEMTableConfig)
  Set TextStr2=LMConfig.OpenAsTextStream(1,0)
  If (TextStr2.AtEndOfStream=True) then
    Exit Sub
  End if
  tempC=TextStr2.ReadLine
  TextStr2.Close
  tempb2s=cdbl(tempC)
  if tempb2s=0 then
    B2SOn=false
  else
    B2SOn=true
  end if
  Set LMConfig=Nothing
  Set FileObj=Nothing
end sub

sub SaveLMEMConfig2
  If ShadowConfigFile=false then exit sub
  Dim FileObj
  Dim LMConfig2
  dim temp1
  dim temp2
  dim tempBS
  dim tempFS

  if EnableBallShadow=true then
    tempBS=1
  else
    tempBS=0
  end if
  if EnableFlipperShadow=true then
    tempFS=1
  else
    tempFS=0
  end if

  Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if
  Set LMConfig2=FileObj.CreateTextFile(UserDirectory & LMEMShadowConfig,True)
  LMConfig2.WriteLine tempBS
  LMConfig2.WriteLine tempFS
  LMConfig2.Close
  Set LMConfig2=Nothing
  Set FileObj=Nothing

end Sub

sub LoadLMEMConfig2
  If ShadowConfigFile=false then
    EnableBallShadow = ShadowBallOn
    BallShadowUpdate.enabled = ShadowBallOn
    EnableFlipperShadow = ShadowFlippersOn
    FlipperLSh.visible = ShadowFlippersOn
    FlipperRSh.visible = ShadowFlippersOn
    exit sub
  end if
  Dim FileObj
  Dim LMConfig2
  dim tempC
  dim tempD
  dim tempFS
  dim tempBS

    Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if
  If Not FileObj.FileExists(UserDirectory & LMEMShadowConfig) then
    Exit Sub
  End if
  Set LMConfig2=FileObj.GetFile(UserDirectory & LMEMShadowConfig)
  Set TextStr2=LMConfig2.OpenAsTextStream(1,0)
  If (TextStr2.AtEndOfStream=True) then
    Exit Sub
  End if
  tempC=TextStr2.ReadLine
  tempD=TextStr2.Readline
  TextStr2.Close
  tempBS=cdbl(tempC)
  tempFS=cdbl(tempD)
  if tempBS=0 then
    EnableBallShadow=false
    BallShadowUpdate.enabled=false
  else
    EnableBallShadow=true
  end if
  if tempFS=0 then
    EnableFlipperShadow=false
    FlipperLSh.visible=false
    FLipperRSh.visible=false
  else
    EnableFlipperShadow=true
  end if
  Set LMConfig2=Nothing
  Set FileObj=Nothing

end sub

' Load high scores
' Path example : C:\Games\VisualPinball\User\
sub loadhs
    ' Based on Black's Highscore routines
  Dim FileObj
  Dim ScoreFile
    dim temp1
    dim temp2
  dim temp3
  dim temp4
  dim temp5
  dim temp6
  dim temp7
  dim temp8
  dim temp9
  dim temp10
  dim temp11
  dim temp12
  dim temp13
  dim temp14
  dim temp15
  dim temp16
  dim temp17


    Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if
  If Not FileObj.FileExists(UserDirectory & HSFileName) then
    Exit Sub
  End if
  Set ScoreFile=FileObj.GetFile(UserDirectory & HSFileName)
  Set TextStr=ScoreFile.OpenAsTextStream(1,0)
    If (TextStr.AtEndOfStream=True) then
      Exit Sub
    End if
    temp1=TextStr.ReadLine
    temp2=textstr.readline
    temp3=textstr.readline
    temp4=textstr.readline
    temp5=textstr.readline
    temp7=textstr.readline
    temp6=textstr.readline
    HighScore=cdbl(temp1)
    if HighScore=1 then

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
    end if
    TextStr.Close
    if HighScore=1 then
      Credits=cdbl(temp2)
      BallsPerGame=cdbl(temp3)
      TiltEndsGame=cdbl(temp4)
      CardLightOption=cdbl(temp5)
      SpecialOption=cdbl(temp7)
      ReplayLevel=cdbl(temp6)

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
    end if
    Set ScoreFile=Nothing
      Set FileObj=Nothing
end sub

Sub DisplayHighScore


end sub

sub InitPauser5_timer
    If B2SOn Then
      'Controller.B2SSetScore 3,HighScore
    End If
    DisplayHighScore
    CreditsReel.SetValue(Credits)
    InitPauser5.enabled=false
end sub

'Plastics lights
Sub PlasticsOn
  For each obj in Flashers
    obj.Visible=1
  Next
End sub

Sub PlasticsOff

  For each obj in Flashers
    obj.Visible=0
  next

end sub

Sub SetupReplayTables
  Replay1Table(1)=74000
  Replay1Table(2)=76000
  Replay1Table(3)=78000
  Replay1Table(4)=80000
  Replay1Table(5)=82000
  Replay1Table(6)=84000
  Replay1Table(7)=122000
  Replay1Table(8)=124000
  Replay1Table(9)=128000
  Replay1Table(10)=132000
  Replay1Table(11)=71000
  Replay1Table(12)=74000
  Replay1Table(13)=78000
  Replay1Table(14)=999000
  Replay1Table(15)=999000


  Replay2Table(1)=86000
  Replay2Table(2)=88000
  Replay2Table(3)=90000
  Replay2Table(4)=92000
  Replay2Table(5)=94000
  Replay2Table(6)=96000
  Replay2Table(7)=140000
  Replay2Table(8)=142000
  Replay2Table(9)=146000
  Replay2Table(10)=150000
  Replay2Table(11)=85000
  Replay2Table(12)=88000
  Replay2Table(13)=92000
  Replay2Table(14)=999000
  Replay2Table(15)=999000

  Replay3Table(1)=98000
  Replay3Table(2)=100000
  Replay3Table(3)=102000
  Replay3Table(4)=104000
  Replay3Table(5)=106000
  Replay3Table(6)=108000
  Replay3Table(7)=158000
  Replay3Table(8)=160000
  Replay3Table(9)=164000
  Replay3Table(10)=168000
  Replay3Table(11)=93000
  Replay3Table(12)=96000
  Replay3Table(13)=999000
  Replay3Table(14)=999000
  Replay3Table(15)=999000

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

  ReplayTableMax=10
end sub

Sub RefreshReplayCard
  Dim tempst1
  Dim tempst2

  tempst1=FormatNumber(BallsPerGame,0)
  tempst2=FormatNumber(ReplayLevel,0)

  BallCard.image = "BC"+tempst1
  Replay1=Replay1Table(ReplayLevel)
  Replay2=Replay2Table(ReplayLevel)
  Replay3=Replay3Table(ReplayLevel)
  Replay4=Replay4Table(ReplayLevel)
  ReplayCard.image = "SC" + tempst2
end sub


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
      If MotorRunning<>1 And InProgress=true then
        queuedscore=queuedscore+y
      end if
    end Select
end sub

Sub SetMotor2(x)
  If MotorRunning<>1 And InProgress=true then
    MotorRunning=1

    Select Case x
      Case 50:
        MotorMode=10
        MotorPosition=5
      Case 100:
        MotorMode=100
        MotorPosition=1

      Case 500:
        MotorMode=100
        MotorPosition=5
      Case 1000:
        MotorMode=1000
        MotorPosition=1
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
        MotorPosition=5
    End Select
  End If
End Sub

Sub AddScoreTimer_Timer
  Dim tempscore


  If MotorRunning<>1 And InProgress=true then
    if queuedscore>=5000 then
      tempscore=5000
      queuedscore=queuedscore-5000
      SetMotor2(5000)
      exit sub
    end if
    if queuedscore>=4000 then
      tempscore=4000
      queuedscore=queuedscore-4000
      SetMotor2(4000)
      exit sub
    end if

    if queuedscore>=3000 then
      tempscore=3000
      queuedscore=queuedscore-3000
      SetMotor2(3000)
      exit sub
    end if

    if queuedscore>=2000 then
      tempscore=2000
      queuedscore=queuedscore-2000
      SetMotor2(2000)
      exit sub
    end if

    if queuedscore>=1000 then
      tempscore=1000
      queuedscore=queuedscore-1000
      SetMotor2(1000)
      exit sub
    end if

    if queuedscore>=500 then
      tempscore=500
      queuedscore=queuedscore-500
      SetMotor2(500)
      exit sub
    end if

    if queuedscore>=100 then
      tempscore=100
      queuedscore=queuedscore-100
      SetMotor2(100)
      exit sub
    end if

    if queuedscore>=50 then
      tempscore=50
      queuedscore=queuedscore-50
      SetMotor2(50)
      exit sub
    end if

  End If


end Sub

Sub ScoreMotorTimer_Timer
  If MotorPosition > 0 Then
    Select Case MotorPosition
      Case 5,4,3,2:
        If MotorMode=1000 Then
          AddScore(1000)
        end if
        if MotorMode=100 then
          AddScore(100)
        End If
        if MotorMode=10 then
          AddScore(10)
        End if
        MotorPosition=MotorPosition-1
      Case 1:
        If MotorMode=1000 Then
          AddScore(1000)
        end if
        If MotorMode=100 then
          AddScore(100)
        End If
        if MotorMode=10 then
          AddScore(10)
        End if
        MotorPosition=0:MotorRunning=0
    End Select
  End If
End Sub


Sub AddScore(x)
  If TableTilted=true then exit sub
  Select Case ScoreAdditionAdjustment
    Case 0:
      AddScore1(x)
    Case 1:
      AddScore2(x)
  end Select
end sub


Sub AddScore1(x)
' debugtext.text=score
  Select Case x
    Case 1:
      PlayChime(10)
      Score(CurrentPlayer)=Score(CurrentPlayer)+1

    Case 10:
      PlayChime(10)
      Score(CurrentPlayer)=Score(CurrentPlayer)+10
'     debugscore=debugscore+10

    Case 100:
      PlayChime(100)
      Score(CurrentPlayer)=Score(CurrentPlayer)+100
'     debugscore=debugscore+100


    Case 1000:
      PlayChime(100)
      Score(CurrentPlayer)=Score(CurrentPlayer)+1000
'     debugscore=debugscore+1000
  End Select
  PlayerScores(CurrentPlayer-1).AddValue(x)

  If ScoreDisplay(CurrentPlayer)<100000 then
    ScoreDisplay(CurrentPlayer)=Score(CurrentPlayer)
  Else
    Score100K(CurrentPlayer)=Int(Score(CurrentPlayer)/100000)
    ScoreDisplay(CurrentPlayer)=Score(CurrentPlayer)-100000
  End If
  if Score(CurrentPlayer)=>100000 then
    EVAL("RolloverReel"&CurrentPlayer).SetValue(1)
    If B2SOn Then
      If CurrentPlayer=1 Then
        Controller.B2SSetScoreRolloverPlayer1 Score100K(CurrentPlayer)
      End If
      If CurrentPlayer=2 Then
        Controller.B2SSetScoreRolloverPlayer2 Score100K(CurrentPlayer)
      End If

      If CurrentPlayer=3 Then
        Controller.B2SSetScoreRolloverPlayer3 Score100K(CurrentPlayer)
      End If

      If CurrentPlayer=4 Then
        Controller.B2SSetScoreRolloverPlayer4 Score100K(CurrentPlayer)
      End If
    End If
  End If
  If B2SOn Then
    Controller.B2SSetScorePlayer CurrentPlayer, ScoreDisplay(CurrentPlayer)
  End If
  If Score(CurrentPlayer)>Replay1 and Replay1Paid(CurrentPlayer)=false then
    Replay1Paid(CurrentPlayer)=True
    AddSpecial
  End If
  If Score(CurrentPlayer)>Replay2 and Replay2Paid(CurrentPlayer)=false then
    Replay2Paid(CurrentPlayer)=True
    AddSpecial
  End If
  If Score(CurrentPlayer)>Replay3 and Replay3Paid(CurrentPlayer)=false then
    Replay3Paid(CurrentPlayer)=True
    AddSpecial
  End If
  If Score(CurrentPlayer)>Replay4 and Replay4Paid(CurrentPlayer)=false then
    Replay4Paid(CurrentPlayer)=True
    AddSpecial
  End If
' ScoreText.text=debugscore
End Sub

Sub AddScore2(x)
  Dim OldScore, NewScore, OldTestScore, NewTestScore
    OldScore = Score(CurrentPlayer)

  Select Case x
        Case 1:
            Score(CurrentPlayer)=Score(CurrentPlayer)+1
    Case 10:
      Score(CurrentPlayer)=Score(CurrentPlayer)+10
    Case 100:
      Score(CurrentPlayer)=Score(CurrentPlayer)+100
    Case 1000:
      Score(CurrentPlayer)=Score(CurrentPlayer)+1000
  End Select
  NewScore = Score(CurrentPlayer)
  If NewScore>=100000 then
    EVAL("RolloverReel"&CurrentPlayer).SetValue(1)
    if B2SOn then
      if CurrentPlayer=1 then Controller.B2SSetScoreRolloverPlayer1 1
      elseif CurrentPlayer=2 then Controller.B2SSetScoreRolloverPlayer2 1
      elseif CurrentPlayer=3 then Controller.B2SSetScoreRolloverPlayer3 1
      elseif CurrentPlayer=4 then Controller.B2SSetScoreRolloverPlayer4 1
    end if
  end if
  OldTestScore = OldScore
  NewTestScore = NewScore
  Do
    if OldTestScore < Replay1 and NewTestScore >= Replay1 AND Replay1Paid(CurrentPlayer)=false then
      AddSpecial()
      Replay1Paid(CurrentPlayer)=true
      NewTestScore = 0
    Elseif OldTestScore < Replay2 and NewTestScore >= Replay2 AND Replay2Paid(CurrentPlayer)=false then
      AddSpecial()
      Replay2Paid(CurrentPlayer)=true
      NewTestScore = 0
    Elseif OldTestScore < Replay3 and NewTestScore >= Replay3 AND Replay3Paid(CurrentPlayer)=false then
      AddSpecial()
      Replay3Paid(CurrentPlayer)=true
      NewTestScore = 0
    Elseif OldTestScore < Replay4 and NewTestScore >= Replay4 AND Replay4Paid(CurrentPlayer)=false then
      AddSpecial()
      Replay4Paid(CurrentPlayer)=true
      NewTestScore = 0
    End if
    NewTestScore = NewTestScore - 100000
    OldTestScore = OldTestScore - 100000
  Loop While NewTestScore > 0

    OldScore = int(OldScore / 10) ' divide by 10 for games with fixed 0 in 1s position, by 1 for games with real 1s digits
    NewScore = int(NewScore / 10) ' divide by 10 for games with fixed 0 in 1s position, by 1 for games with real 1s digits
  ' MsgBox("OldScore="&OldScore&", NewScore="&NewScore&", OldScore Mod 10="&OldScore Mod 10 & ", NewScore % 10="&NewScore Mod 10)

    if (OldScore Mod 10 <> NewScore Mod 10) then
    PlayChime(10)
    end if

    OldScore = int(OldScore / 10)
    NewScore = int(NewScore / 10)
  ' MsgBox("OldScore="&OldScore&", NewScore="&NewScore)
    if (OldScore Mod 10 <> NewScore Mod 10) then
    PlayChime(10)
    end if

    OldScore = int(OldScore / 10)
    NewScore = int(NewScore / 10)
  ' MsgBox("OldScore="&OldScore&", NewScore="&NewScore)
    if (OldScore Mod 10 <> NewScore Mod 10) then
    PlayChime(100)
    end if

    OldScore = int(OldScore / 10)
    NewScore = int(NewScore / 10)
  ' MsgBox("OldScore="&OldScore&", NewScore="&NewScore)
    if (OldScore Mod 10 <> NewScore Mod 10) then
    PlayChime(1000)
    end if

  If B2SOn Then
    Controller.B2SSetScorePlayer CurrentPlayer, Score(CurrentPlayer)
  End If
' EMReel1.SetValue Score(CurrentPlayer)
  PlayerScores(CurrentPlayer-1).AddValue(x)

End Sub



Sub PlayChime(x)
  if ChimesOn=0 then
    Select Case x
      Case 10
        If LastChime10=1 Then
          PlaySound SoundFXDOF("SpinACard_1_10_Point_Bell",141,DOFPulse,DOFChimes)
          LastChime10=0
        Else
          PlaySound SoundFXDOF("SpinACard_1_10_Point_Bell",141,DOFPulse,DOFChimes)
          LastChime10=1
        End If
      Case 100
        If LastChime100=1 Then
          PlaySound SoundFXDOF("SpinACard_100_Point_Bell",142,DOFPulse,DOFChimes)
          LastChime100=0
        Else
          PlaySound SoundFXDOF("SpinACard_100_Point_Bell",142,DOFPulse,DOFChimes)
          LastChime100=1
        End If

    End Select
  else
    Select Case x
      Case 10
        If LastChime10=1 Then
          PlaySound SoundFXDOF("SJ_Chime_10a",141,DOFPulse,DOFChimes)
          LastChime10=0
        Else
          PlaySound SoundFXDOF("SJ_Chime_10b",141,DOFPulse,DOFChimes)
          LastChime10=1
        End If
      Case 100
        If LastChime100=1 Then
          PlaySound SoundFXDOF("SJ_Chime_100a",142,DOFPulse,DOFChimes)
          LastChime100=0
        Else
          PlaySound SoundFXDOF("SJ_Chime_100b",142,DOFPulse,DOFChimes)
          LastChime100=1
        End If
      Case 1000
        If LastChime1000=1 Then
          PlaySound SoundFXDOF("SJ_Chime_1000a",143,DOFPulse,DOFChimes)
          LastChime1000=0
        Else
          PlaySound SoundFXDOF("SJ_Chime_1000b",143,DOFPulse,DOFChimes)
          LastChime1000=1
        End If
    End Select
  end if
End Sub

Sub HideOptions()

end sub

' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX, Rothbauerw, Thalamus and Herweh
' PlaySound sound, 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
' *******************************************************************************************************

' Play a sound, depending on the X,Y position of the table element (especially cool for surround speaker setups, otherwise stereo panning only)
' parameters (defaults): loopcount (1), volume (1), randompitch (0), pitch (0), useexisting (0), restart (1))
' Note that this will not work (currently) for walls/slingshots as these do not feature a simple, single X,Y position

Sub PlayXYSound(soundname, tableobj, loopcount, volume, randompitch, pitch, useexisting, restart)
  PlaySound soundname, loopcount, volume, AudioPan(tableobj), randompitch, pitch, useexisting, restart, AudioFade(tableobj)
End Sub

' Set position as table object (Use object or light but NOT wall) and Vol to 1

Sub PlaySoundAt(soundname, tableobj)
  PlaySound soundname, 1, 1, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

' set position as table object and Vol + RndPitch manually

Sub PlaySoundAtVolPitch(sound, tableobj, Vol, RndPitch)
  PlaySound sound, 1, Vol, AudioPan(tableobj), RndPitch, 0, 0, 1, AudioFade(tableobj)
End Sub

'Set all as per ball position & speed.

Sub PlaySoundAtBall(soundname)
  PlaySoundAt soundname, ActiveBall
End Sub

'Set position as table object and Vol manually.

Sub PlaySoundAtVol(sound, tableobj, Volume)
  PlaySound sound, 1, Volume, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'Set all as per ball position & speed, but Vol Multiplier may be used eg; PlaySoundAtBallVol "sound",3

Sub PlaySoundAtBallVol(sound, VolMult)
  PlaySound sound, 0, Vol(ActiveBall) * VolMult, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtBallAbsVol(sound, VolMult)
  PlaySound sound, 0, VolMult, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub

' requires rampbump1 to 7 in Sound Manager

Sub RandomBump(voladj, freq)
  Dim BumpSnd:BumpSnd= "rampbump" & CStr(Int(Rnd*7)+1)
  PlaySound BumpSnd, 0, Vol(ActiveBall)*voladj, AudioPan(ActiveBall), 0, freq, 0, 1, AudioFade(ActiveBall)
End Sub

' set position as bumperX and Vol manually. Allows rapid repetition/overlaying sound

Sub PlaySoundAtBumperVol(sound, tableobj, Vol)
  PlaySound sound, 1, Vol, AudioPan(tableobj), 0,0,1, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBOTBallZ(sound, BOT)
  PlaySound sound, 0, ABS(BOT.velz)/17, Pan(BOT), 0, Pitch(BOT), 1, 0, AudioFade(BOT)
End Sub

' play a looping sound at a location with volume
Sub PlayLoopSoundAtVol(sound, tableobj, Vol)
  PlaySound sound, -1, Vol, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

'*********************************************************************
'            Supporting Ball, Sound Functions and Math
'*********************************************************************

Function RndNum(min, max)
  RndNum = Int(Rnd() * (max-min + 1) ) + min ' Sets a random number between min and max
End Function

Const Pi = 3.1415927

Function dSin(degrees)
  dsin = sin(degrees * Pi/180)
End Function

Function dCos(degrees)
  dcos = cos(degrees * Pi/180)
End Function

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = tableobj.y * 2 / table1.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = tableobj.x * 2 / table1.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = ball.x * 2 / table1.width-1
  If tmp > 0 Then
    Pan = Csng(tmp ^10)
  Else
    Pan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
  Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function VolMulti(ball,Multiplier) ' Calculates the Volume of the sound based on the ball speed
  VolMulti = Csng(BallVel(ball) ^2 / 150 ) * Multiplier
End Function

Function DVolMulti(ball,Multiplier) ' Calculates the Volume of the sound based on the ball speed
  DVolMulti = Csng(BallVel(ball) ^2 / 150 ) * Multiplier
  debug.print DVolMulti
End Function

Function BallRollVol(ball) ' Calculates the Volume of the sound based on the ball speed
  BallRollVol = Csng(BallVel(ball) ^2 / (80000 - (79900 * Log(RollVol) / Log(100))))
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
  Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
  BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

Function BallVelZ(ball) 'Calculates the ball speed in the -Z
  BallVelZ = INT((ball.VelZ) * -1 )
End Function

Function VolZ(ball) ' Calculates the Volume of the sound based on the ball speed in the Z
  VolZ = Csng(BallVelZ(ball) ^2 / 200)*1.2
End Function

'*** Determines if a Points (px,py) is inside a 4 point polygon A-D in Clockwise/CCW order

Function InRect(px,py,ax,ay,bx,by,cx,cy,dx,dy)
  Dim AB, BC, CD, DA
  AB = (bx*py) - (by*px) - (ax*py) + (ay*px) + (ax*by) - (ay*bx)
  BC = (cx*py) - (cy*px) - (bx*py) + (by*px) + (bx*cy) - (by*cx)
  CD = (dx*py) - (dy*px) - (cx*py) + (cy*px) + (cx*dy) - (cy*dx)
  DA = (ax*py) - (ay*px) - (dx*py) + (dy*px) + (dx*ay) - (dy*ax)

  If (AB <= 0 AND BC <=0 AND CD <= 0 AND DA <= 0) Or (AB >= 0 AND BC >=0 AND CD >= 0 AND DA >= 0) Then
    InRect = True
  Else
    InRect = False
  End If
End Function


'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 5 ' total number of balls
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub RollingSoundTimer_Timer()
    Dim BOT, b
    BOT = GetBalls

  ' stop the sound of deleted balls
    For b = UBound(BOT) + 1 to tnob
        rolling(b) = False
        StopSound("fx_ballrolling" & b)
    Next

  ' exit the sub if no balls on the table
    If UBound(BOT) = -1 Then Exit Sub

  ' play the rolling sound for each ball
    For b = 0 to UBound(BOT)
        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If
    Next
End Sub

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
  PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
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
        For b = (UBound(BOT) + 1) to (tnob-1)
            BallShadow(b).visible = 0
        Next
    End If
    ' exit the Sub if no balls on the table
    If UBound(BOT) = -1 Then Exit Sub
    ' render the shadow for each ball
    For b = 0 to UBound(BOT)
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
  PlaySound "woodhit_low", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Pins_Hit (idx)
  PlaySound "pinhit_low", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
  PlaySound "target", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
  PlaySound "metalhit_thin", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
  PlaySound "metalhit_medium", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
  PlaySound "metalhit2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
  PlaySound "gate4", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinners_Spin (idx)
  PlaySound "fx_spinner", 0, .25, AudioPan(Spinners(idx)), 0.25, 0, 0, 1, AudioFade(Spinners(idx))
End Sub

Sub Rubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 20 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End Sub

Sub Posts_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    RandomSoundRubber()
  End If
End Sub

Sub RandomSoundRubber()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

Sub LeftFlipper_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RightFlipper_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

' Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
Dim RStep, Lstep

' Slingshot right hit
Sub ScoringWall005_Hit()
  ' Animate slingshot
  vpmTimer.PulseSw 43
    PlaySoundAtVol SoundFX("right_slingshot",DOFContactors), ActiveBall, 1
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    ScoringWall005.TimerEnabled = 1

  if TableTilted then Exit Sub
  AddScore(10)
  SwitchTopLaneLights
End Sub

' Slingshot left hit
Sub ScoringWall001_Hit()
  ' Animate slingshot
  'MsgBox "left slingshot"
  vpmTimer.PulseSw 44
    PlaySoundAtVol SoundFX("left_slingshot",DOFContactors), ActiveBall, 1
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    ScoringWall001.TimerEnabled = 1

  if TableTilted then Exit Sub
  AddScore(10)
  SwitchTopLaneLights
End Sub

Sub ScoringWall005_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:ScoringWall005.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub ScoringWall001_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:ScoringWall001.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

' ============================================================================================
' GNMOD - Multiple High Score Display and Collection
' Operator menu, options saved to User directory\KickOff_77VPX.txt
' ============================================================================================
Dim EnteringInitials    ' Normally zero, set to non-zero to enter initials
EnteringInitials = 0

Dim PlungerPulled
PlungerPulled = 0

Dim SelectedChar      ' character under the "cursor" when entering initials

Dim HSTimerCount      ' Pass counter for HS timer, scores are cycled by the timer
HSTimerCount = 5      ' Timer is initially enabled, it'll wrap from 5 to 1 when it's displayed

Dim InitialString     ' the string holding the player's initials as they're entered

Dim AlphaString       ' A-Z, 0-9, space (_) and backspace (<)
Dim AlphaStringPos      ' pointer to AlphaString, move forward and backward with flipper keys
AlphaString = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_<"

Dim HSNewHigh       ' The new score to be recorded

Dim HSScore(5)        ' High Scores read in from config file
Dim HSName(5)       ' High Score Initials read in from config file

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

Sub HighScoreTimer_Timer

  if EnteringInitials then
    if HSTimerCount = 1 then
      SetHSLine 3, InitialString & MID(AlphaString, AlphaStringPos, 1)
      HSTimerCount = 2
    else
      SetHSLine 3, InitialString
      HSTimerCount = 1
    end if
  elseif InProgress then
    SetHSLine 1, "HIGH SCORE1"
    SetHSLine 2, HSScore(1)
    SetHSLine 3, HSName(1)
    HSTimerCount = 5  ' set so the highest score will show after the game is over
    HighScoreTimer.enabled=false
  elseif CheckAllScores then
    NewHighScore sortscores(ScoreChecker),sortplayers(ScoreChecker)

  else
    ' cycle through high scores
    HighScoreTimer.interval=2000
    HSTimerCount = HSTimerCount + 1
    if HsTimerCount > 5 then
      HSTimerCount = 1
    End If
    SetHSLine 1, "HIGH SCORE"+FormatNumber(HSTimerCount,0)
    SetHSLine 2, HSScore(HSTimerCount)
    SetHSLine 3, HSName(HSTimerCount)
  end if
End Sub

Function GetHSChar(String, Index)
  dim ThisChar
  dim FileName
  ThisChar = Mid(String, Index, 1)
  FileName = "PostIt"
  if ThisChar = " " or ThisChar = "" then
    FileName = FileName & "BL"
  elseif ThisChar = "<" then
    FileName = FileName & "LT"
  elseif ThisChar = "_" then
    FileName = FileName & "SP"
  else
    FileName = FileName & ThisChar
  End If
  GetHSChar = FileName
End Function

Sub SetHsLine(LineNo, String)
  dim Letter
  dim ThisDigit
  dim ThisChar
  dim StrLen
  dim LetterLine
  dim Index
  dim StartHSArray
  dim EndHSArray
  dim LetterName
  dim xfor
  StartHSArray=array(0,1,12,22)
  EndHSArray=array(0,11,21,31)
  StrLen = len(string)
  Index = 1

  for xfor = StartHSArray(LineNo) to EndHSArray(LineNo)
    Eval("HS"&xfor).image = GetHSChar(String, Index)
    Index = Index + 1
  next

End Sub

Sub NewHighScore(NewScore, PlayNum)
  if NewScore > HSScore(5) then
    HighScoreTimer.interval = 500
    HSTimerCount = 1
    AlphaStringPos = 1    ' start with first character "A"
    EnteringInitials = 1  ' intercept the control keys while entering initials
    InitialString = ""    ' initials entered so far, initialize to empty
    SetHSLine 1, "PLAYER "+FormatNumber(PlayNum,0)
    SetHSLine 2, "ENTER NAME"
    SetHSLine 3, MID(AlphaString, AlphaStringPos, 1)
    HSNewHigh = NewScore
    For xx=1 to HighScoreReward
      AddSpecial
    next
  End if
  ScoreChecker=ScoreChecker-1
  if ScoreChecker=0 then
    CheckAllScores=0
  end if
End Sub

Sub CollectInitials(keycode)
  If keycode = LeftFlipperKey Then
    ' back up to previous character
    AlphaStringPos = AlphaStringPos - 1
    if AlphaStringPos < 1 then
      AlphaStringPos = len(AlphaString)   ' handle wrap from beginning to end
      if InitialString = "" then
        ' Skip the backspace if there are no characters to backspace over
        AlphaStringPos = AlphaStringPos - 1
      End if
    end if
    SetHSLine 3, InitialString & MID(AlphaString, AlphaStringPos, 1)
    PlaySound "DropTargetDropped"
  elseif keycode = RightFlipperKey Then
    ' advance to next character
    AlphaStringPos = AlphaStringPos + 1
    if AlphaStringPos > len(AlphaString) or (AlphaStringPos = len(AlphaString) and InitialString = "") then
      ' Skip the backspace if there are no characters to backspace over
      AlphaStringPos = 1
    end if
    SetHSLine 3, InitialString & MID(AlphaString, AlphaStringPos, 1)
    PlaySound "DropTargetDropped"
  elseif keycode = StartGameKey or keycode = PlungerKey Then
    SelectedChar = MID(AlphaString, AlphaStringPos, 1)
    if SelectedChar = "_" then
      InitialString = InitialString & " "
      PlaySound("Ding10")
    elseif SelectedChar = "<" then
      InitialString = MID(InitialString, 1, len(InitialString) - 1)
      if len(InitialString) = 0 then
        ' If there are no more characters to back over, don't leave the < displayed
        AlphaStringPos = 1
      end if
      PlaySound("Ding100")
    else
      InitialString = InitialString & SelectedChar
      PlaySound("Ding10")
    end if
    if len(InitialString) < 3 then
      SetHSLine 3, InitialString & SelectedChar
    End If
  End If
  if len(InitialString) = 3 then
    ' save the score
    for i = 5 to 1 step -1
      if i = 1 or (HSNewHigh > HSScore(i) and HSNewHigh <= HSScore(i - 1)) then
        ' Replace the score at this location
        if i < 5 then
' MsgBox("Moving " & i & " to " & (i + 1))
          HSScore(i + 1) = HSScore(i)
          HSName(i + 1) = HSName(i)
        end if
' MsgBox("Saving initials " & InitialString & " to position " & i)
        EnteringInitials = 0
        HSScore(i) = HSNewHigh
        HSName(i) = InitialString
        HSTimerCount = 5
        HighScoreTimer_Timer
        HighScoreTimer.interval = 2000
        PlaySound("Ding1000")
        exit sub
      elseif i < 5 then
        ' move the score in this slot down by 1, it's been exceeded by the new score
' MsgBox("Moving " & i & " to " & (i + 1))
        HSScore(i + 1) = HSScore(i)
        HSName(i + 1) = HSName(i)
      end if
    next
  End If

End Sub

' END GNMOD
' ============================================================================================
' GNMOD - New Options menu
' Operator menu, options saved to User directory\KickOff_77VPX.txt
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
Const OptionLinesToMark="111011011"
Const OptionLine1="" 'do not use this line
Const OptionLine2="" 'do not use this line
Const OptionLine3="" 'do not use this line
Const OptionLine4=""
Const OptionLine5="Extra Ball Light Option"
Const OptionLine6="Special Light Option"
Const OptionLine7=""
Const OptionLine8="" 'do not use this line
Const OptionLine9="" 'do not use this line

Sub OperatorMenuTimer_Timer
  EnteringOptions = 1
  OperatorMenuTimer.enabled=false
  ShowOperatorMenu
end sub

sub ShowOperatorMenu
  OperatorMenuBackdrop.image = "OperatorMenu"

  OptionCHS = 0
  CurrentOption = 1
  DisplayAllOptions
  OperatorOption1.image = "BluePlus"
  SetHighScoreOption

End Sub

Sub DisplayAllOptions
  dim linecounter
  dim tempstring
  For linecounter = 1 to MaxOption
    tempstring=Eval("OptionLine"&linecounter)
    Select Case linecounter
      Case 1:
        tempstring=tempstring + FormatNumber(BallsPerGame,0)
        SetOptLine 1,tempstring
      Case 2:
        if Replay3Table(ReplayLevel)=999000 then
          tempstring = FormatNumber(Replay1Table(ReplayLevel),0) + "/" + FormatNumber(Replay2Table(ReplayLevel),0)
        elseif Replay4Table(ReplayLevel)=999000 then
          tempstring = FormatNumber(Replay1Table(ReplayLevel),0) + "/" + FormatNumber(Replay2Table(ReplayLevel),0) + "/" + FormatNumber(Replay3Table(ReplayLevel),0)
        else
          tempstring = FormatNumber(Replay1Table(ReplayLevel),0) + "/" + FormatNumber(Replay2Table(ReplayLevel),0) + "/" + FormatNumber(Replay3Table(ReplayLevel),0) + "/" + FormatNumber(Replay4Table(ReplayLevel),0)
        end if
        SetOptLine 2,tempstring
      Case 3:
        If OptionCHS=0 then
          tempstring = "NO"
        else
          tempstring = "YES"
        end if
        SetOptLine 3,tempstring
      Case 4:
        SetOptLine 4, tempstring


        SetOptLine 5, tempstring
      Case 5:
        SetOptLine 6, tempstring
        Select Case CardLightOption
          Case 0:
            tempstring = "No Extra Ball Awards"
          Case 1:
            tempstring = "Bonus at 10k, 12k, 15k"
          Case 2:
            tempstring = "Bonus at 11k, 13k, 16k"
          Case 3:
            tempstring = "Bonus at 12k, 14k, 17k"
          Case 4:
            tempstring = "Bonus at 13k, 15k, 18k"
          Case 5:
            tempstring = "Bonus at 14k, 16k, 19k"
        end select

        SetOptLine 7, tempstring

      Case 6:
        SetOptLine 8, tempstring
        Select case SpecialOption
          case 0:
            tempstring = "No Special Award"
          case 1:
            tempstring = "Bonus at 10,000"
          case 2:
            tempstring = "Bonus at 11,000"
          case 3:
            tempstring = "Bonus at 12,000"
          case 4:
            tempstring = "Bonus at 13,000"
          case 5:
            tempstring = "Bonus at 14,000"
          case 6:
            tempstring = "Bonus at 15,000"
          case 7:
            tempstring = "Bonus at 16,000"
          case 8:
            tempstring = "Bonus at 17,000"
          case 9:
            tempstring = "Bonus at 18,000"
          case 10:
            tempstring = "Bonus at 19,000"

        end select

        SetOptLine 9, tempstring

      Case 7:
        SetOptLine 10, tempstring
        SetOptLine 11, tempstring

      Case 8:

      Case 9:


    End Select

  next
end sub

sub MoveArrow
  do
    CurrentOption = CurrentOption + 1
    If CurrentOption>Len(OptionLinesToMark) then
      CurrentOption=1
    end if
  loop until Mid(OptionLinesToMark,CurrentOption,1)="1"
end sub

sub CollectOptions(ByVal keycode)
  if Keycode = LeftFlipperKey then
    PlaySound "DropTargetDropped"
    For XOpt = 1 to MaxOption
      Eval("OperatorOption"&XOpt).image = "PostitBL"
    next
    MoveArrow
    if CurrentOption<8 then
      Eval("OperatorOption"&CurrentOption).image = "BluePlus"
    elseif CurrentOption=8 then
      Eval("OperatorOption"&CurrentOption).image = "GreenCheck"
    else
      Eval("OperatorOption"&CurrentOption).image = "RedX"
    end if

  elseif Keycode = RightFlipperKey then
    PlaySound "DropTargetDropped"
    if CurrentOption = 1 then
      If BallsPerGame = 3 then
        BallsPerGame = 5
      else
        BallsPerGame = 3
      end if
      DisplayAllOptions
    elseif CurrentOption = 2 then
      ReplayLevel=ReplayLevel+1
      If ReplayLevel>ReplayTableMax then
        ReplayLevel=1
      end if
      DisplayAllOptions
    elseif CurrentOption = 3 then
      if OptionCHS = 0 then
        OptionCHS = 1

      else
        OptionCHS = 0

      end if
      DisplayAllOptions


    elseif CurrentOption = 5 then
      CardLightOption=CardLightOption+1
      if CardLightOption>5 then
        CardLightOption=0
      end if
      DisplayAllOptions

    elseif CurrentOption = 6 then
      SpecialOption=SpecialOption+1
      if SpecialOption>10 then
        SpecialOption=0
      end if
      DisplayAllOptions

    elseif CurrentOption = 8 or CurrentOption = 9 then
        if OptionCHS=1 then
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
        end if

        if CurrentOption = 8 then
          savehs
        else
          loadhs
        end if
        OperatorMenuBackdrop.image = "PostitBL"
        For XOpt = 1 to MaxOption
          Eval("OperatorOption"&XOpt).image = "PostitBL"
        next

        For XOpt = 1 to 256
          Eval("Option"&XOpt).image = "PostItBL"
        next
        RefreshReplayCard
        InstructCard.image="IC"
        EnteringOptions = 0

    end if
  end if
End Sub

Sub SetHighScoreOption

End Sub

Function GetOptChar(String, Index)
  dim ThisChar
  dim FileName
  ThisChar = Mid(String, Index, 1)
  FileName = "PostIt"
  if ThisChar = " " or ThisChar = "" then
    FileName = FileName & "BL"
  elseif ThisChar = "<" then
    FileName = FileName & "LT"
  elseif ThisChar = ">" then
    FileName = FileName & "GT"
  elseif ThisChar = "_" then
    FileName = FileName & "SP"
  elseif ThisChar = "/" then
    FileName = FileName & "SL"
  elseif ThisChar = "," then
    FileName = FileName & "CM"
  else
    FileName = FileName & ThisChar
  End If
  GetOptChar = FileName
End Function

dim LineLengths(22) ' maximum number of lines
Sub SetOptLine(LineNo, String)
  Dim DispLen
    Dim StrLen
  dim xfor
  dim Letter
  dim ThisDigit
  dim ThisChar
  dim LetterLine
  dim Index
  dim LetterName
  StrLen = len(string)
  Index = 1

  StrLen = len(String)
    DispLen = StrLen
    if (DispLen < LineLengths(LineNo)) Then
        DispLen = LineLengths(LineNo)
    end If

  for xfor = StartingArray(LineNo) to StartingArray(LineNo) + DispLen
    Eval("Option"&xfor).image = GetOptChar(string, Index)
    Index = Index + 1
  next
  LineLengths(LineNo) = StrLen

End Sub

'*******************************************
' MESSAGES :
' Prerequisites :
' A Timer object named MessagesTimer must exists in the interface !
' A Textbox named MessageDebugText
' mestype : score, (multiball, ballsaved)
' seconds : number of seconds to display
'*******************************************

Dim Messages_PreviousMessage, MessagesDebugEnable
Messages_PreviousMessage = ""
MessagesDebugEnable = False ' TODO : set to False in production (not developer mode)
Sub Message(mes)
  MessagesTimer.Enabled = False
  MessageDebug mes
  MessageDebugText.Text = mes
  Messages_PreviousMessage = mes
End Sub

Sub MessageTimed(mes, seconds)
  MessageDebug mes
  MessagesTimer.Enabled = False
  MessageDebugText.Text = mes
  MessagesTimer.Interval = seconds * 1000 'in milliseconds
  MessagesTimer.Enabled = True
End Sub

Sub MessagesTimer_Init()
  MessagesTimer.Enabled = False
End Sub

Sub MessagesTimer_Timer()
  MessageDebug "end timer message"
  MessagesTimer.Enabled = False
  MessageDebugText.Text = Messages_PreviousMessage
End Sub

' Get path to this .vpx file
Dim Messages_curDir, Messages_fso
Set Messages_fso = CreateObject("Scripting.FileSystemObject")
Messages_curDir = Messages_fso.GetAbsolutePathName(".")
Set Messages_fso=Nothing

' Add message to a file debug.txt
Sub MessageDebug(mes)
  If Not MessagesDebugEnable Then
    Exit Sub
  End If

  Set Messages_fso = CreateObject("Scripting.FileSystemObject")
  Dim debugFile
  Set debugFile = Messages_fso.OpenTextFile(Messages_curDir & "\debug.txt", 8, True)

  Dim MyDate, MyTime, MyTime2
  MyDate = Date
  MyTime = Time
  MyTime2 = MyTime
  Dim dateFull
  dateFull = Date & "_" & MyTime & "_" & Timer
  If Len(dateFull) < 28 Then
    dateFull = dateFull & "0"
  End If

  debugFile.writeline dateFull & "   " & mes ' & vbNewLine
  debugFile.Close
  Set debugFile=Nothing
  Set Messages_fso=Nothing
End Sub

' Clean messages to the file debug.txt
Sub MessagesDebugClean()
  Set Messages_fso = CreateObject("Scripting.FileSystemObject")
  Dim debugFile
  Set debugFile = Messages_fso.OpenTextFile(Messages_curDir & "\debug.txt", 2, True)
  debugFile.Close
  Set debugFile=Nothing
  Set Messages_fso=Nothing
End Sub

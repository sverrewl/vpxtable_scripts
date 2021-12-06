'*
'*        Gottlieb's Spot-A-Card (1960)
'*        Table primary build/scripted by Loserman76
'*
'*
'*

option explicit
Randomize

' Thalamus 2019 March : Improved directional sounds
' !! NOTE : Table not verified yet !!

' Options
' Volume devided by - lower gets higher sound

Const VolDiv = 2000    ' Lower number, louder ballrolling/collition sound
Const VolCol = 10      ' Ball collition divider ( voldiv/volcol )

' The rest of the values are multipliers
'
'  .5 = lower volume
' 1.5 = higher volume

Const VolBump   = 1    ' Bumpers volume.

ExecuteGlobal GetTextFile("core.vbs")

On Error Resume Next
ExecuteGlobal GetTextFile("Controller.vbs")
If Err Then MsgBox "Unable to open Controller.vbs. Ensure that it is in the scripts folder."
On Error Goto 0

Const cGameName = "spotacard_1960"


Const ShadowFlippersOn = true
Const ShadowBallOn = true

Const ShadowConfigFile = false




Dim Controller  ' B2S
Dim B2SScore  ' B2S Score Displayed
Const HSFileName="SpotACard_60VPX.txt"
Const B2STableName="SpotACard_1960"
Const LMEMTableConfig="LMEMTables.txt"
Const LMEMShadowConfig="LMEMShadows.txt"
Dim EnableBallShadow
Dim EnableFlipperShadow

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
Dim TextStr,TextStr2,LaneClear
Dim i
Dim obj
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
Dim TableTilted
Dim TiltCount

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
dim TempPlayerup
dim RotatorTemp

Dim bump1
Dim bump2
Dim bump3
Dim bump4
Dim bump5

Dim LastChime10
Dim LastChime100
Dim LastChime1000

Dim Score10
Dim Score100

Dim targettempscore
Dim SpecialLightCounter
Dim SpecialLightOption
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
Dim TargetLightsOn
Dim AdvanceLightCounter
dim bonustempcounter

dim raceRflag
dim raceAflag
dim raceCflag
dim raceEflag
dim race5flag

Dim LStep, LStep2, RStep, xx

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

Dim TargetSequenceComplete

Dim SpecialLightsFlag

Dim BumperSequenceCompleted,AdvancesCompleted, Motor3EToggle

Dim AdvanceBonusMax, AdvanceSpecialMax

Dim BonusScoreLevel1,BonusScoreLevel2,BonusScoreLevel3,BonusScoreLevel4,BonusScoreLevel5,EndGameCount

Dim AlternatingRelay

Dim ZeroToNineUnit

Dim SlickCounter,ChickCounter,SlickChickFlag,SlickChickComplete,NumberCounter

Dim Kicker1Hold,Kicker2Hold

Dim mHole,mHole2,mHole3,mHole4

Dim GreenBumpersOn, YellowBumpersOn, ArrowCounter, CardSequenceComplete, MatchLit2

Dim WhiteComplete,RedComplete,YellowComplete,GreenComplete,DisableKeysInit,BallLiftOption,BallsPlayed,BallsDrained

Dim SpinPos,Spin,Count,Count1,Count2,Count3,Reset,VelX,VelY,BallSpeed,LitSpinner,BallIndicatorCount,CenterTargetTracker,BallsOnTable,TiltEndsGame

Dim tKickerTCount1,tKickerTCount2,tKickerTCount3,tKickerTCount4
Dim ArrowCounter1,ArrowCounter2,ArrowCounter3,ArrowCounter4
Dim CardCounter

Spin=Array("5","6","7","8","9","10","J","Q","K","A")

Sub Table1_Init()
  DisableKeysInit=1
  LoadEM
  LoadLMEMConfig2
  If Table1.ShowDT = false then
    For each obj in DesktopCrap
      obj.visible=False
    next
    BallsPlayedRamp.image="FSBallsPlayed"
  End If

  OperatorMenuBackdrop.image = "PostitBL"
  For XOpt = 1 to MaxOption
    Eval("OperatorOption"&XOpt).image = "PostitBL"
  next

  For XOpt = 1 to 256
    Eval("Option"&XOpt).image = "PostItBL"
  next

  BallIndicatorCount=0
  InitBallLoad.enabled=true
  For each obj in ArrowLights
    obj.state=0
  next

  Count=0
    Count1=0
    Count2=0
  Count3=0
    Reset=0
  ZeroToNineUnit=Int(Rnd*10)
  Kicker1Hold=0
  ArrowCounter=0
  EightLit=0
  TiltEndsGame=1
  BallCounter=0
  ReelCounter=0
  AddABall=0
  DropABall=0
  BallsOnTable=0
  HideOptions
  SetupReplayTables
  PlasticsOff
  BumpersOff
  OperatorMenu=0
  HighScore=0
  MotorRunning=0
  HighScoreReward=1
  Motor3EToggle=0
  Credits=0
  BallsPerGame=5
  ReplayLevel=1
  AdvanceBonusMax=5
  AdvanceSpecialMax=3
  ChimesOn=0
  BallRelease.enabled=false
  AlternatingRelay=0
  BumperSequenceCompleted=0
  SpecialLightOption=2
  BallLiftOption=1
  BackglassBallFlagColor=1
  loadhs
  if HighScore=0 then HighScore=5000


  TableTilted=false

  Match=int(Rnd*10)
  Do
    MatchLit2=int(Rnd*10)
  loop while MatchLit2=Match

  DisplayMatch

  SpinPos=int(Rnd*10)

' CanPlayReel.SetValue(0)
  GameOverReel.SetValue(1)
  TiltReel.SetValue(1)
  Reel1000.SetValue(0)

  For each obj in PlayerHuds
    obj.SetValue(0)
  next
  For each obj in PlayerScores
    obj.ResetToZero
  next
  for each obj in AllLights
    obj.state=0
  next

  Replay1=Replay1Table(ReplayLevel)
  Replay2=Replay2Table(ReplayLevel)
  Replay3=Replay3Table(ReplayLevel)
  Replay4=Replay4Table(ReplayLevel)


  BonusCounter=0
  HoleCounter=0


  'InstructCard.image="IC"+FormatNumber(BallsPerGame,0)+FormatNumber(AdvanceSpecialMax,0)

  RefreshReplayCard

  CurrentFrame=0

  For each obj in SlickBumpers
    obj.state=0
  next
  For each obj in ChickBumpers
    obj.state=0
  next


  AdvanceLightCounter=0

  Players=0
  RotatorTemp=1
  InProgress=false
  TargetLightsOn=false

  ScoreText.text=HighScore


  If B2SOn Then


    Controller.B2SSetScoreRolloverPlayer1 0
    Controller.B2SSetScoreRolloverPlayer2 0
    Controller.B2SSetScoreRolloverPlayer3 0
    Controller.B2SSetScoreRolloverPlayer4 0
    Controller.B2SSetData 99,0
    Controller.B2SSetScore 3,HighScore
    Controller.B2SSetTilt 1
    Controller.B2SSetCredits Credits
    Controller.B2SSetGameOver 1
  End If

  for i=1 to 4
    player=i
    If B2SOn Then
      Controller.B2SSetScorePlayer player, 0
    End If
  next
  bump1=1
  bump2=1
  bump3=1
  bump4=1
  bump5=1
  InitPauser5.enabled=true
  If B2SOn then
    Controller.B2SSetCredits Credits
  end if
  if Credits > 0 then DOF 144, DOFOn
End Sub

Sub InitBallLoad_timer
  Kicker3.CreateSizedBall 20
  Kicker3.Kick 0,12
  BallIndicatorCount=BallIndicatorCount+1
  if BallIndicatorCount=5 then
    InitBallLoad.enabled=false
    DisableKeysInit=0
  end if
end sub

Sub Kicker2_Hit
  Kicker2.DestroyBall
  BallIndicatorCount=BallIndicatorCount-1
  If BallIndicatorCount=0 then
    Gate2.Collidable=true
  end if
end sub

Sub Table1_exit()
  savehs
  SaveLMEMConfig
  SaveLMEMConfig2
  If B2SOn Then Controller.Stop
end sub


Sub TimerTitle1_timer
  if B2SOn then
    if raceRflag=1 then
      Controller.B2SSetData 11,0
      raceRflag=0
    else
      Controller.B2SSetData 11,1
      raceRflag=1
    end if
  end if
end sub

Sub TimerTitle2_timer
  if B2SOn then
    if raceAflag=1 then
      Controller.B2SSetData 12,0
      raceAflag=0
    else
      Controller.B2SSetData 12,1
      raceAflag=1
    end if
  end if
end sub

Sub TimerTitle3_timer
  if B2SOn then
    if raceCflag=1 then
      Controller.B2SSetData 13,0
      raceCflag=0
    else
      Controller.B2SSetData 13,1
      raceCflag=1
    end if
  end if
end sub

Sub TimerTitle4_timer
  if B2SOn then
    if raceEflag=1 then
      Controller.B2SSetData 14,0
      raceEflag=0
    else
      Controller.B2SSetData 14,1
      raceEflag=1
    end if
  end if
end sub

Sub TimerTitle5_timer
  if B2SOn then
    if race5flag=1 then
      Controller.B2SSetData 15,0
      race5flag=0
    else
      Controller.B2SSetData 15,1
      race5flag=1
    end if
  end if
end sub


Sub Table1_KeyDown(ByVal keycode)
  If DisableKeysInit=1 then
    exit sub
  end if

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

  if keycode = LeftFlipperKey and InProgress = false then
    OperatorMenuTimer.Enabled = true
  end if
  ' END GNMOD


  If keycode = LeftFlipperKey and InProgress=true and TableTilted=false Then
    LeftFlipper.RotateToEnd
    PlaySoundAtLoopVol "buzzL", LeftFlipper, 1
    PlaySoundAtVol SoundFXDOF("FlipperUp",101,DOFOn,DOFContactors), LeftFlipper, 1
  End If

  If keycode = RightFlipperKey  and InProgress=true and TableTilted=false Then
    RightFlipper.RotateToEnd

    PlaySoundAtVol SoundFXDOF("FlipperUp",102,DOFOn,DOFContactors), RightFlipper, 1
    PlaySoundAtLoopVol "buzz",RightFlipper, 1
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
    TiltCount=3
    TiltIt
  End If


  If keycode = AddCreditKey or keycode = 4 then
    If B2SOn Then
      'Controller.B2SSetScorePlayer6 HighScore

    End If

    playsoundAtVol "coinin", drain, 1
    AddSpecial2
  end if

   if keycode = 5 then
    playsoundAtVol "coinin", drain, 1
    AddSpecial2
    keycode= StartGameKey
  end if




  if (keycode=StartGameKey OR keycode=RightMagnaSave) and InProgress=true and BallLiftOption=1 and BallsPlayed<BallsPerGame and LaneClear=0 then
    PlaySoundAtVol SoundFXDOF("BallLifter",145,DOFPulse,DOFContactors), Plunger, 1 'close enough
    BallsPlayed=BallsPlayed+1
    'CreateBallID BallRelease
    Ballrelease.CreateSizedBall 25
    Ballrelease.Kick 270,5
    PlaysoundAtVol "ballrelease", Plunger, 1
    BallsOnTable=BallsOnTable+1
  end if


  if keycode=StartGameKey and Credits>0 and InProgress=false and Players=0 and EnteringOptions = 0 and BallIndicatorCount=5 then
'GNMOD
    OperatorMenuTimer.Enabled = false
'END GNMOD
    DisableKeysInit=1
    Credits=Credits-1
    if Credits <1 then DOF 144, DOFOff
    CreditsReel.SetValue(Credits)
    Gate2.Collidable=false
    Players=1
'   CanPlayReel.SetValue(Players)
    ClearMatch
    Player=1

    If BallLiftOption=2 then
      playsound "startup_norm"
    else
      playsound "startup_manual"
    end if
    TempPlayerUp=Player
'   PlayerUpRotator.enabled=true
    rst=0
    BallInPlay=1
    Light3.state=1
    CenterTargetTracker=1
    InProgress=true
    resettimer.enabled=true
    BonusMultiplier=1
    LaneClear=0
    TableTilted=false
    TiltTimer.Interval = 500
    TiltTimer.Enabled = True
    TimerTitle1.enabled=0
    TimerTitle2.enabled=0
    TimerTitle3.enabled=0
    TimerTitle4.enabled=0
    TimerTitle5.enabled=0
    If B2SOn Then
      Controller.B2SSetTilt 0
      Controller.B2SSetGameOver 0
      Controller.B2SSetMatch 0
      Controller.B2SSetCredits Credits
      '
      Controller.B2SSetCanPlay 1
      Controller.B2SSetPlayerUp 1
      Controller.B2SSetBallInPlay BallInPlay
      Controller.B2SSetScoreRolloverPlayer1 0
      Controller.B2SSetData 99,0
      Controller.B2SSetData 11,0
      Controller.B2SSetData 12,0
      Controller.B2SSetData 13,0
      Controller.B2SSetData 14,0
      Controller.B2SSetData 15,0
      For i = 70 to 84
        Controller.B2SSetData i,0
      next
    End If
    If Table1.ShowDT = True then
      For each obj in PlayerScores
'       obj.ResetToZero
        obj.Visible=false
      next
      For each obj in PlayerScoresOn
'       obj.ResetToZero
        obj.Visible=true
      next

      For each obj in PlayerHuds
        obj.SetValue(0)
      next
      Reel1000.SetValue(0)
      GameOverReel.SetValue(0)
      PlayerHuds(Player-1).SetValue(1)
      PlayerScores(Player-1).visible=0
      PlayerScoresOn(Player-1).visible=1


    end If
    BallsPlayed=0
    BallsDrained=0
    BumperSequenceCompleted=0
  end if



End Sub

Sub Table1_KeyUp(ByVal keycode)
  If DisableKeysInit=1 then
    exit sub
  end if


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
    PlaySoundAtVol SoundFXDOF("FlipperDown",101,DOFOff,DOFContactors), LeftFlipper, 1
    StopSound "buzzL"
  End If

  If keycode = RightFlipperKey and InProgress=true and TableTilted=false Then
    RightFlipper.RotateToStart
        StopSound "buzz"
    PlaySoundAtVol SoundFXDOF("FlipperDown",102,DOFOff,DOFContactors), RightFlipper, 1
  End If

End Sub

Sub ShooterLaneClear_Hit
  LaneClear=1
end sub

Sub ShooterLaneClear_Unhit
  LaneClear=0
end sub

Sub Drain_Hit()

  BallsDrained=BallsDrained+1

  Drain.DestroyBall
  AddScore(10)
  if TableTilted=false then
    CardLights(ZeroToNineUnit).state=1
    If B2SOn then
      Controller.B2SSetData 80+ZeroToNineUnit,1
    end if
    CheckRolloverStates
  end if
  BallsOnTable=BallsOnTable-1
  PlaySoundAtVol "fx_drain", Drain, 1
  DOF 107, DOFPulse
  'Pause4Bonustimer.enabled=true
  Kicker3.CreateSizedBall 20
  Kicker3.Kick 0,12
  BallIndicatorCount=BallIndicatorCount+1
  If BallLiftOption=1 then
    NextBall
  Else
    Pause4Bonustimer.enabled=true
  end if
End Sub

Sub GobbleHole1_Hit()
  BallsDrained=BallsDrained+1
  BallsOnTable=BallsOnTable-1
  GobbleHole1.DestroyBall
  GobbleHole1.timerenabled=true

end sub

Sub GobbleHole1_timer
  If MotorRunning<>0 then
    exit sub
  end if
  GobbleHole1.timerenabled=false
  SetMotor(200)

  PlaySoundAtVol "fx_drain", Drain, 1
  DOF 107, DOFPulse
  'Pause4Bonustimer.enabled=true
  Kicker3.CreateSizedBall 20
  Kicker3.Kick 0,12
  BallIndicatorCount=BallIndicatorCount+1
  If BallLiftOption=1 then
    NextBall
  Else
    Pause4Bonustimer.enabled=true
  end if
End Sub

Sub CloseGateTrigger_Hit()
  DOF 149, DOFPulse
End Sub

Sub Pause4Bonustimer_timer
  Pause4Bonustimer.enabled=0
  NextBallDelay.enabled=true

End Sub

'***********************
'     Flipper Logos
'***********************

Sub UpdateFlipperLogos_Timer
  LFlip.RotZ = LeftFlipper.CurrentAngle
  RFlip.RotZ = RightFlipper.CurrentAngle
  LFlip1.RotZ = LeftFlipper.CurrentAngle
  RFlip1.RotZ = RightFlipper.CurrentAngle

  PGate.Rotz = (Gate.CurrentAngle*.75) + 25
  FlipperLSh.RotZ = LeftFlipper.currentangle
  FlipperRSh.RotZ = RightFlipper.currentangle
End Sub

'***********************
' slingshots
'

Sub RightSlingShot_Slingshot
    PlaySoundAtVol SoundFXDOF("right_slingshot",104,DOFPulse,DOFContactors), sling1, 1
    DOF 106, DOFPulse
  RSling0.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
  If RSlingLight.state=1 then
    AddScore 10
  else
    AddScore 1
  end if
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSling0.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
    PlaySoundAtVol SoundFXDOF("left_slingshot",103,DOFPulse,DOFContactors), sling2, 1
  DOF 105, DOFPulse
    LSling0.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
  If LSlingLight.state=1 then
    AddScore 10
  else
    AddScore 1
  end if
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing0.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

'***********************************
' Walls
'***********************************

Sub RubberWallSwitches_Hit(idx)
  if TableTilted=false then
    Select Case idx
      Case 0,1:
        SetMotor(5)
      Case 2,3:
        AddScore(1)

    end Select
  end if
end Sub


'***********************************
' Bumpers
'***********************************
Sub Bumper1_Hit
  If TableTilted=false then
    PlaySoundAtVol SoundFXDOF("bumper1",112,DOFPulse,DOFContactors), ActiveBall, VolBump
    DOF 113, DOFPulse
    bump3 = 1
    AddScore(1)
    IncreaseArrowCounter
    end if
End Sub

Sub Bumper2_Hit
  If TableTilted=false then
    PlaySoundAtVol SoundFXDOF("bumper1",110,DOFPulse,DOFContactors), ActiveBall, VolBump
    DOF 111, DOFPulse
    bump2 = 1
    If Bumper2Light.state=0 then
      AddScore(1)
    else
      AddScore(10)
    end if
    end if
End Sub

Sub Bumper3_Hit
  If TableTilted=false then
    PlaySoundAtVol SoundFXDOF("bumper1",108,DOFPulse,DOFContactors), ActiveBall, VolBump
    DOF 109, DOFPulse
    bump1 = 1
    If Bumper3Light.state=0 then
      AddScore(1)
    else
      AddScore(10)
    end if
    end if
End Sub


Sub Bumper5_Hit
  If TableTilted=false then
    PlaySoundAtVol SoundFXDOF("bumper1",117,DOFPulse,DOFContactors), ActiveBall, VolBump
    DOF 114, DOFPulse
    AddScore(1)
    end if
End Sub

Sub Bumper6_Hit
  If TableTilted=false then
    PlaySoundAtVol SoundFXDOF("bumper1",119,DOFPulse,DOFContactors), ActiveBall, VolBump
    DOF 115, DOFPulse
    AddScore(1)
    end if
End Sub


'****************
Sub BumperRubber2_Hit
    Bumper6_Hit
End Sub

Sub BumperRubber1_Hit
    Bumper5_Hit
End Sub

'************************************
'  Kickers
'************************************
sub MidKicker1_Hit()
  Dim tempkickscore
  Dim speedx,speedy,finalspeed
  speedx=activeball.velx
  speedy=activeball.vely
  finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  if TableTilted=true then
    MidKicker1.Timerenabled=1
    tKickerTCount1=0
    exit sub
  end if
  if finalspeed>10 then
    MidKicker1.Kick 0,0
    activeball.velx=speedx
    activeball.vely=speedy
  else

    SetMotor(20)

    MidKicker1.timerenabled=1
    CardLights(ArrowCounter1).state=1
    If B2SOn then
      Controller.B2SSetData 80+ArrowCounter1,1
    end if
    If ArrowCounter1=5 then
      TargetLight9_1.state=0
      TargetLight9_2.state=0
    end if
    If ArrowCounter1=6 then
      TargetLight8_1.state=0
      TargetLight8_2.state=0
    end if
    If ArrowCounter1=7 then
      TargetLight7_1.state=0
      TargetLight7_2.state=0
    end if
    If ArrowCounter1=8 then
      TargetLight6_1.state=0
      TargetLight6_2.state=0
    end if
    If ArrowCounter1=10 then
      LowerLight4.state=0
    end if
    If ArrowCounter1=9 then
      LowerLight5.state=0
    end if
    If ArrowCounter1=11 then
      TargetLight3_1.state=0
      TargetLight3_2.state=0
      LowerLight3.state=0
    end if
    If ArrowCounter1=12 then
      TargetLight2_1.state=0
      TargetLight2_2.state=0
      LowerLight2.state=0
    end if
    tKickerTCount1=0
  end if

end sub

Sub MidKicker1_Timer()
  if MotorRunning<>0 then
    exit sub
  end if
  tKickerTCount1=tKickerTCount1+1
  select case tKickerTCount1
  case 1:
  ' mHole.MagnetOn=0
    Pkickarm1.rotz=15
    PlaysoundAtVol SoundFXDOF("saucer",120,DOFPulse,DOFContactors), Pkickarm1, 1
    DOF 121, DOFPulse
    MidKicker1.kick 160,9
  case 2:
    MidKicker1.timerenabled=0
    Pkickarm1.rotz=0
  end Select

end sub


sub MidKicker2_Hit()
  Dim tempkickscore
  Dim speedx,speedy,finalspeed
  speedx=activeball.velx
  speedy=activeball.vely
  finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  if TableTilted=true then
    MidKicker2.Timerenabled=1
    tKickerTCount2=0
    exit sub
  end if
  if finalspeed>10 then
    MidKicker2.Kick 0,0
    activeball.velx=speedx
    activeball.vely=speedy
  else
    SetMotor(20)
    MidKicker2.timerenabled=1
    CardLights(ArrowCounter2).state=1
    If B2SOn then
      Controller.B2SSetData 80+ArrowCounter2,1
    end if
    If ArrowCounter2=5 then
      TargetLight9_1.state=0
      TargetLight9_2.state=0
    end if
    If ArrowCounter2=6 then
      TargetLight8_1.state=0
      TargetLight8_2.state=0
    end if
    If ArrowCounter2=7 then
      TargetLight7_1.state=0
      TargetLight7_2.state=0
    end if
    If ArrowCounter2=8 then
      TargetLight6_1.state=0
      TargetLight6_2.state=0
    end if
    If ArrowCounter2=10 then
      LowerLight4.state=0
    end if
    If ArrowCounter2=9 then
      LowerLight5.state=0
    end if
    If ArrowCounter2=11 then
      TargetLight3_1.state=0
      TargetLight3_2.state=0
      LowerLight3.state=0
    end if
    If ArrowCounter2=12 then
      TargetLight2_1.state=0
      TargetLight2_2.state=0
      LowerLight2.state=0
    end if
    tKickerTCount2=0
  end if

end sub

Sub MidKicker2_Timer()
  if MotorRunning<>0 then
    exit sub
  end if
  tKickerTCount2=tKickerTCount2+1
  select case tKickerTCount2
  case 1:
  ' mHole.MagnetOn=0
    Pkickarm2.rotz=15
    PlaysoundAtVol SoundFXDOF("saucer",120,DOFPulse,DOFContactors), Pkickarm2, 1
    DOF 121, DOFPulse
    MidKicker2.kick 160,9
  case 2:
    MidKicker2.timerenabled=0
    Pkickarm2.rotz=0
  end Select

end sub

sub MidKicker3_Hit()
  Dim tempkickscore
  Dim speedx,speedy,finalspeed
  speedx=activeball.velx
  speedy=activeball.vely
  finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  if TableTilted=true then
    MidKicker3.Timerenabled=1
    tKickerTCount3=0
    exit sub
  end if
  if finalspeed>10 then
    MidKicker3.Kick 0,0
    activeball.velx=speedx
    activeball.vely=speedy
  else

    SetMotor(20)
    MidKicker3.timerenabled=1
    CardLights(ArrowCounter3).state=1
    If B2SOn then
      Controller.B2SSetData 80+ArrowCounter3,1
    end if
    If ArrowCounter3=5 then
      TargetLight9_1.state=0
      TargetLight9_2.state=0
    end if
    If ArrowCounter3=6 then
      TargetLight8_1.state=0
      TargetLight8_2.state=0
    end if
    If ArrowCounter3=7 then
      TargetLight7_1.state=0
      TargetLight7_2.state=0
    end if
    If ArrowCounter3=8 then
      TargetLight6_1.state=0
      TargetLight6_2.state=0
    end if
    If ArrowCounter3=10 then
      LowerLight4.state=0
    end if
    If ArrowCounter3=9 then
      LowerLight5.state=0
    end if
    If ArrowCounter3=11 then
      TargetLight3_1.state=0
      TargetLight3_2.state=0
      LowerLight3.state=0
    end if
    If ArrowCounter3=12 then
      TargetLight2_1.state=0
      TargetLight2_2.state=0
      LowerLight2.state=0
    end if
    tKickerTCount3=0
  end if

end sub

Sub MidKicker3_Timer()
  if MotorRunning<>0 then
    exit sub
  end if
  tKickerTCount3=tKickerTCount3+1
  select case tKickerTCount3
  case 1:
  ' mHole.MagnetOn=0
    Pkickarm3.rotz=15
    PlaysoundAtVol SoundFXDOF("saucer",120,DOFPulse,DOFContactors), Pkickarm3, 1
    DOF 121, DOFPulse
    MidKicker3.kick 160,9
  case 2:
    MidKicker3.timerenabled=0
    Pkickarm3.rotz=0
  end Select
end sub

sub MidKicker4_Hit()
  Dim tempkickscore
  Dim speedx,speedy,finalspeed
  speedx=activeball.velx
  speedy=activeball.vely
  finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  if TableTilted=true then
    MidKicker4.Timerenabled=1
    tKickerTCount4=0
    exit sub
  end if
  if finalspeed>10 then
    MidKicker4.Kick 0,0
    activeball.velx=speedx
    activeball.vely=speedy
  else

    SetMotor(20)
    MidKicker4.timerenabled=1
    CardLights(ArrowCounter4).state=1
    If B2SOn then
      Controller.B2SSetData 80+ArrowCounter4,1
    end if
    If ArrowCounter4=5 then
      TargetLight9_1.state=0
      TargetLight9_2.state=0
    end if
    If ArrowCounter4=6 then
      TargetLight8_1.state=0
      TargetLight8_2.state=0
    end if
    If ArrowCounter4=7 then
      TargetLight7_1.state=0
      TargetLight7_2.state=0
    end if
    If ArrowCounter4=8 then
      TargetLight6_1.state=0
      TargetLight6_2.state=0
    end if
    If ArrowCounter4=10 then
      LowerLight4.state=0
    end if
    If ArrowCounter4=9 then
      LowerLight5.state=0
    end if
    If ArrowCounter4=11 then
      TargetLight3_1.state=0
      TargetLight3_2.state=0
      LowerLight3.state=0
    end if
    If ArrowCounter4=12 then
      TargetLight2_1.state=0
      TargetLight2_2.state=0
      LowerLight2.state=0
    end if
    tKickerTCount4=0
  end if

end sub

Sub MidKicker4_Timer()
  if MotorRunning<>0 then
    exit sub
  end if
  tKickerTCount4=tKickerTCount4+1
  select case tKickerTCount4
  case 1:
  ' mHole.MagnetOn=0
    Pkickarm4.rotz=15
    PlaysoundAtVol SoundFXDOF("saucer",120,DOFPulse,DOFContactors), Pkickarm4, 1
    DOF 121, DOFPulse
    MidKicker4.kick 160,9
  case 2:
    MidKicker4.timerenabled=0
    Pkickarm4.rotz=0
  end Select

end sub

'************************************
'  Targets
'************************************
Sub TargetWall2_Hit
  If TableTilted=false then
    DOF 130, DOFPulse
    AddScore(10)
    CardLights(12).state=1
    TargetNumber2.transX=-5
    TargetWall2.timerenabled=true
    TargetLight2_1.state=0
    TargetLight2_2.state=0
    LowerLight2.state=0
    If B2SOn then
      Controller.B2SSetData 92,1
    end if
  end if
end sub

Sub TargetWall2_timer
  TargetWall2.timerenabled=false
  TargetNumber2.transX=0
end sub

Sub TargetWall3_Hit
  If TableTilted=false then
    DOF 131, DOFPulse
    AddScore(10)
    CardLights(11).state=1
    TargetNumber3.transX=-5
    TargetWall3.timerenabled=true
    TargetLight3_1.state=0
    TargetLight3_2.state=0
    LowerLight3.state=0
    If B2SOn then
      Controller.B2SSetData 91,1
    end if
  end if
end sub

Sub TargetWall3_timer
  TargetWall3.timerenabled=false
  TargetNumber3.transX=0

end sub

Sub TargetWall6_Hit
  If TableTilted=false then
    DOF 133, DOFPulse
    AddScore(10)
    CardLights(8).state=1
    TargetNumber6.transX=-5
    TargetWall6.timerenabled=true
    TargetLight6_1.state=0
    TargetLight6_2.state=0
    If B2SOn then
      Controller.B2SSetData 88,1
    end if
  end if
end sub

Sub TargetWall6_timer
  TargetWall6.timerenabled=false
  TargetNumber6.transX=0
end sub

Sub TargetWall7_Hit
  If TableTilted=false then
    DOF 133, DOFPulse
    AddScore(10)
    CardLights(7).state=1
    TargetNumber7.transX=-5
    TargetWall7.timerenabled=true
    TargetLight7_1.state=0
    TargetLight7_2.state=0
    If B2SOn then
      Controller.B2SSetData 87,1
    end if
  end if
end sub

Sub TargetWall7_timer
  TargetWall7.timerenabled=false
  TargetNumber7.transX=0
end sub

Sub TargetWall8_Hit
  If TableTilted=false then
    DOF 132, DOFPulse
    AddScore(10)
    CardLights(6).state=1
    TargetNumber8.transX=-5
    TargetWall8.timerenabled=true
    TargetLight8_1.state=0
    TargetLight8_2.state=0
    If B2SOn then
      Controller.B2SSetData 86,1
    end if
  end if
end sub

Sub TargetWall8_timer
  TargetWall8.timerenabled=false
  TargetNumber8.transX=0
end sub

Sub TargetWall9_Hit
  If TableTilted=false then
    DOF 132, DOFPulse
    AddScore(10)
    CardLights(5).state=1
    TargetNumber9.transX=-5
    TargetWall9.timerenabled=true
    TargetLight9_1.state=0
    TargetLight9_2.state=0
    If B2SOn then
      Controller.B2SSetData 85,1
    end if
  end if
end sub

Sub TargetWall9_timer
  TargetWall9.timerenabled=false
  TargetNumber9.transX=0
end sub

'************************************
'  Button Trigger
'************************************
Sub ButtonTrigger1_Hit
  if TableTilted=false then
    AddScore(1)
    If SpecialLight.state=1 then
      AddSpecial
    end if
  end if

end sub

'************************************
'  Rollover lanes
'************************************

Sub TriggerCollection_Hit(idx)

  If TableTilted=false then
    DOF 136+idx,DOFPulse
    AddScore(10)
    Select Case idx
      Case 0:

        CardLights(ZeroToNineUnit).state=1
        If B2SOn then
          Controller.B2SSetData 80+ZeroToNineUnit,1
        end if
        if RolloverLight1_1.state=1 then
          Bumper2Light.state=1
          Bumper3Light.state=1
          LSlingLight.state=1
          RSlingLight.state=1
        end if
      Case 1:

        LowerLight2.state=0
        CardLight13.state=1
        TargetLight2_1.state=0
        TargetLight2_2.state=0
        If B2SOn then
          Controller.B2SSetData 92,1
        end if

      Case 2:

        LowerLight3.State=0
        CardLight12.state=1
        TargetLight3_1.state=0
        TargetLight3_2.state=0
        If B2SOn then
          Controller.B2SSetData 91,1
        end if

      Case 3:

        LowerLight4.state=0
        CardLight11.state=1
        If B2SOn then
          Controller.B2SSetData 90,1
        end if

      Case 4:

        LowerLight5.state=0
        CardLight10.state=1
        If B2SOn then
          Controller.B2SSetData 89,1
        end if


    End Select
    CheckRolloverStates
  end if

end Sub

Sub CheckRolloverStates
  If CardSequenceComplete=false then
    CardCounter=0
    For each obj in CardLights
      if obj.state=1 then
        CardCounter=CardCounter+1
      end if
    next
    If CardCounter=13 then
      SpecialLight.state=1
      AddSpecial

      CardSequenceComplete=true
    end if
  end if



end sub

Sub IncreaseArrowCounter

  ArrowCounter1=ArrowCounter1+1
  If ArrowCounter1>12 then
    ArrowCounter1=0
  end if
  ArrowCounter2=ArrowCounter2+1
  If ArrowCounter2>12 then
    ArrowCounter2=0
  end if
  ArrowCounter3=ArrowCounter3+1
  If ArrowCounter3>12 then
    ArrowCounter3=0
  end if
  ArrowCounter4=ArrowCounter4+1
  If ArrowCounter4>12 then
    ArrowCounter4=0
  end if

  For each obj in ArrowLights
    obj.state=0
  next
  ArrowLights(ArrowCounter1).state=1
  ArrowLights(ArrowCounter2).state=1
  ArrowLights(ArrowCounter3).state=1
  ArrowLights(ArrowCounter4).state=1
  If B2SOn then
    for i=60 to 72
      Controller.B2SSetData i,0
    next
    Controller.B2SSetData 60+ArrowCounter1,1
    Controller.B2SSetData 60+ArrowCounter2,1
    Controller.B2SSetData 60+ArrowCounter3,1
    Controller.B2SSetData 60+ArrowCounter4,1
  end if
end sub

'**************************************


Sub AddSpecial()
  PlaySound SoundFXDOF("knocker",142,DOFPulse,DOFContactors)
  DOF 121, DOFPulse
  Credits=Credits+1
  DOF 144, DOFOn
  if Credits>15 then Credits=15
  If B2SOn Then
    Controller.B2SSetCredits Credits
  End If
  CreditsReel.SetValue(Credits)
End Sub

Sub AddSpecial2()
  PlaySound"click"
  Credits=Credits+1
  DOF 144, DOFOn
  if Credits>25 then Credits=25
  If B2SOn Then
    Controller.B2SSetCredits Credits
  End If
  CreditsReel.SetValue(Credits)
End Sub

Sub LightTheBumpers

end sub

Sub AdvanceZeroToNine
  ZeroToNineUnit=ZeroToNineUnit+1
  if ZeroToNineUnit>4 then
    ZeroToNineUnit=0
  end if

  For each obj in RolloverLightsTop
    obj.state=0
  next
  For each obj in OutholeLights
    obj.state=0
  next
  RolloverLightsTop(ZeroToNineUnit).state=1
  OutholeLights(ZeroToNineUnit).state=1
  ToggleAlternatingRelay
end sub

Sub ToggleAlternatingRelay
  AlternatingRelay=AlternatingRelay+1
  if AlternatingRelay>9 then
    AlternatingRelay=0
  end if
  RolloverLight1_1.state=0
  RolloverLight1_2.state=0
  RolloverLight1_3.state=0
  RolloverLight1_4.state=0
  Select case AlternatingRelay
    case 0,3,6,9:
      RolloverLight1_1.state=1
      RolloverLight1_2.state=1
      RolloverLight1_3.state=1
      RolloverLight1_4.state=1
    case 1,4,7:
      if AdvanceSpecialMax=2 or AdvanceSpecialMax=3 then
        RolloverLight1_1.state=1
        RolloverLight1_2.state=1
        RolloverLight1_3.state=1
        RolloverLight1_4.state=1
      end if
    case 2,5,8:
      if AdvanceSpecialMax=3 then
        RolloverLight1_1.state=1
        RolloverLight1_2.state=1
        RolloverLight1_3.state=1
        RolloverLight1_4.state=1
      end if
  end select
end sub

Sub ResetBallDrops
  LSlingLight.state=0
  RSlingLight.state=0

  Bumper5Light.state=1
  Bumper6Light.state=1

  Bumper1Light.state=1
  Bumper2Light.state=0
  Bumper3Light.state=0
  HoleCounter=0
  YellowBumpersOn=false
  GreenBumpersOn=false
  LightTheBumpers
End Sub


Sub LightsOut

  BonusCounter=0
  HoleCounter=0

  StopSound "buzz"
  StopSound "buzzL"

end sub

Sub ResetBalls()

  TempMultiCounter=BallsPerGame-BallInPlay

  ResetBallDrops
  BonusMultiplier=1
  If BallsOnTable<1 then
    TableTilted=false
    TiltReel.SetValue(0)
    If B2Son then
      Controller.B2SSetTilt 0
    end if

    PlasticsOn
    BumpersOn
  end if
  If BallLiftOption=2 then
    PlayStartBall.enabled=true
    BallsPlayed=BallsPlayed+1
    'CreateBallID BallRelease
    Ballrelease.CreateSizedBall 25
    Ballrelease.Kick 270,5
    DOF 145,DOFPulse
    BallsOnTable=BallsOnTable+1
  ' BallInPlayReel.SetValue(BallInPlay)
    'InstructCard.image="IC"+FormatNumber(BallsPerGame,0)+FormatNumber(AdvanceSpecialMax,0)
  end if


End Sub





sub resettimer_timer
    rst=rst+1

  if rst>3 and rst<14 then
    ResetReelsToZero(1)
  end if
    if rst=16 then
    if BallLiftOption=2 then
      playsoundAtVol "StartBall1", Plunger, 1
    end if
    end if
    if rst=18 then
    newgame
    DisableKeysInit=0
    resettimer.enabled=false
    end if
end sub

Sub ResetReelsToZero(reelzeroflag)
  dim d1(5)
  dim d2(5)
  dim scorestring1, scorestring2

  If reelzeroflag=1 then
    scorestring1=CStr(Score(1))
    scorestring1=right("00000" & scorestring1,5)
    for i=0 to 4
      d1(i)=CInt(mid(scorestring1,i+1,1))
    next
    for i=0 to 4
      if d1(i)>0 then
        d1(i)=d1(i)+1
        if d1(i)>9 then d1(i)=0
      end if

    next
    Score(1)=(d1(0)*10000) + (d1(1)*1000) + (d1(2)*100) + (d1(3)*10) + d1(4)
    If B2SOn Then
      Controller.B2SSetScorePlayer 1, Score(1)
    End If
    PlayerScores(0).SetValue(Score(1))
    PlayerScoresOn(0).SetValue(Score(1))

  end if
  If reelzeroflag=2 then
    scorestring1=CStr(Score(2))
    scorestring1=right("00000" & scorestring1,5)
    for i=0 to 4
      d1(i)=CInt(mid(scorestring1,i+1,1))
    next
    for i=0 to 4
      if d1(i)>0 then
        d1(i)=d1(i)+1
        if d1(i)>9 then d1(i)=0
      end if

    next
    Score(2)=(d1(0)*10000) + (d1(1)*1000) + (d1(2)*100) + (d1(3)*10) + d1(4)
    If B2SOn Then
      Controller.B2SSetScorePlayer 2, Score(2)
    End If
    PlayerScores(1).SetValue(Score(2))
    PlayerScoresOn(1).SetValue(Score(2))

  end if

end sub


sub NextBallDelay_timer()
  NextBallDelay.enabled=false
  nextball

end sub

sub newgame
  InProgress=true
  queuedscore=0
  ClearMatch
  for i = 1 to 4
    Score(i)=0
    Score100K(1)=0
    HighScorePaid(i)=false
    Replay1Paid(i)=false
    Replay2Paid(i)=false
    Replay3Paid(i)=false
  next
  If B2SOn Then
    Controller.B2SSetTilt 0
    Controller.B2SSetGameOver 0

'   Controller.B2SSetScorePlayer1 0
'   Controller.B2SSetScorePlayer2 0
'   Controller.B2SSetScorePlayer3 0
'   Controller.B2SSetScorePlayer4 0
    Controller.B2SSetBallInPlay BallInPlay
  End if
  For each obj in RolloverLightsOn
    obj.state=1
  next
  For each obj in CardLights
    obj.state=0
  next
  If B2SOn then
    For i=60 to 92
      Controller.B2SSetData i,0
    next
  end if
  CardSequenceComplete=false
  AlternatingRelay=1
  ZeroToNineUnit=4
  AdvanceZeroToNine
  ArrowCounter1=12
  ArrowCounter2=0
  ArrowCounter3=1
  ArrowCounter4=2
  IncreaseArrowCounter
  SpecialLight.state=0


  LSlingLight.state=0
  RSlingLight.state=0

  Bumper5Light.state=1
  Bumper6Light.state=1

  Bumper1Light.state=1
  Bumper2Light.state=0
  Bumper3Light.state=0

  For each obj in NumberLights
    obj.state=1
  next


  BumperSequenceCompleted=0
  SlickCounter=0
  ChickCounter=0
  NumberCounter=0
  BumpersOn
  BonusCounter=0
  BallCounter=0
  TargetLeftFlag=1
  TargetCenterFlag=1
  TargetRightFlag=1
  TargetSequenceComplete=0
  AdvancesCompleted=0
' IncreaseBonus
' ToggleBumper
  EightLit=1
  ResetBalls
End sub


sub EndOfGame

  InProgress=false

  If B2SOn Then
    Controller.B2SSetGameOver 1
    Controller.B2SSetPlayerUp 0
    Controller.B2SSetBallInPlay 0
    Controller.B2SSetCanPlay 0
  End If
  For each obj in PlayerHuds
    obj.SetValue(0)
  next
  GameOverReel.SetValue(1)
  'InstructCard.image="IC"+FormatNumber(BallsPerGame,0)

'   BallInPlayReel.SetValue(0)
'     CanPlayReel.SetValue(0)
  LeftFlipper.RotateToStart
  RightFlipper.RotateToStart
  LightsOut
  Light3.state=0
  BumpersOff
  PlasticsOff
  checkmatch
  CheckHighScore
  Players=0
  If Table1.ShowDT = True Then
    PlayerScores(Player-1).visible=1
    PlayerScoresOn(Player-1).visible=0
  end If
  TimerTitle1.enabled=1
  TimerTitle2.enabled=1
  TimerTitle3.enabled=1
  TimerTitle4.enabled=1
  TimerTitle5.enabled=1
  HighScoreTimer.interval=100
  HighScoreTimer.enabled=True
end sub

sub EndOfGameTimer_timer
  EndOfGameTimer.enabled=false
  EndOfGame
end sub

sub nextball

  If TiltEndsGame=1 and TableTilted=true and BallsOnTable>0 then
    exit sub
  end if
  If TiltEndsGame=1 and TableTilted=true and BallsOnTable<1 and BallIndicatorCount<5 then
    DisableKeysInit=1
    InitBallLoad.enabled=true
    exit sub
  end if
  Player=Player+1
  If Player>Players Then
    BallInPlay=BallInPlay+1
    If (BallInPlay>BallsPerGame) or (BallsDrained>=BallsPerGame) then
      PlaySound("MotorLeer")


      EndGameCount=0
      EndOfGameTimer.enabled=true
      exit sub
    Else
      Player=1
      If B2SOn Then
        Controller.B2SSetPlayerUp Player
        Controller.B2SSetBallInPlay BallInPlay

      End If
'     PlaySound("RotateThruPlayers")
      TempPlayerUp=Player
'     PlayerUpRotator.enabled=true
'     PlayStartBall.enabled=true
      For each obj in PlayerHuds
        obj.SetValue(0)
      next
      If Table1.ShowDT = True then
        For each obj in PlayerScores
          obj.visible=1
        Next

        PlayerHuds(Player-1).SetValue(1)

      end If

      ResetBalls
    End If
  Else
    If B2SOn Then
      Controller.B2SSetPlayerUp Player
      Controller.B2SSetBallInPlay BallInPlay
    End If
'   PlaySound("RotateThruPlayers")
    TempPlayerUp=Player
'   PlayerUpRotator.enabled=true
'   PlayStartBall.enabled=true
    For each obj in PlayerHuds
      obj.SetValue(0)
    next
    If Table1.ShowDT = True then
      For each obj in PlayerScores
        obj.visible=1
      Next

      PlayerHuds(Player-1).SetValue(1)

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
  for i = 1 to Players
    sortscores(i)=Score(i)
    sortplayers(i)=i
  next

  for si = 1 to Players
    for sj = 1 to Players-1
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
  savehs
end sub


sub checkmatch
  Dim tempmatch
  Dim tempmatch2
  tempmatch=Int(Rnd*10)
  do
    tempmatch2=Int(Rnd*10)
  loop while tempmatch2=tempmatch

  Match=tempmatch
  MatchLit2=tempmatch2
  If TableTilted=True and TiltEndsGame then
    exit sub
  end if
  DisplayMatch


  for i = 1 to Players
    if Match=(Score(i) mod 10) then
      AddSpecial
    end if

  next
end sub

Sub TiltTimer_Timer()
  if InProgress=false then
    exit sub
  end if
  if TiltCount > 0 then TiltCount = TiltCount - 1
  if TiltCount = 0 then
    TiltTimer.Enabled = False
  end if
end sub

Sub TiltIt()
    TiltCount = TiltCount + 1
    if TiltCount > 3 then
      TableTilted=True
      PlasticsOff
      BumpersOff

      LeftFlipper.RotateToStart
      RightFlipper.RotateToStart
      StopSound "buzz"
      StopSound "buzzL"
      TiltReel.SetValue(1)
      If B2Son then
        Controller.B2SSetTilt 1
      end if
      If TiltEndsGame=1 then
        InProgress=false
        If B2SOn Then
          Controller.B2SSetGameOver 1
          Controller.B2SSetPlayerUp 0
          Controller.B2SSetBallInPlay 0
          Controller.B2SSetCanPlay 0
        End If
        For each obj in PlayerHuds
          obj.SetValue(0)
        next
        GameOverReel.SetValue(1)
        'InstructCard.image="IC"+FormatNumber(BallsPerGame,0)

        Players=0
        If BallsOnTable<1 and BallIndicatorCount<5 then
          DisableKeysInit=1
          InitBallLoad.enabled=true

        end if
      end if
    else
      TiltTimer.Interval = 500
      TiltTimer.Enabled = True
    end if

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
        Controller.B2SSetPlayerUp Player
      end if
      PlayerUpRotator.enabled=false
      RotatorTemp=1
    end if
    RotatorTemp=RotatorTemp+1


end sub

sub savehs
  ' Based on Black's Highscore routines
  Dim FileObj
  Dim ScoreFile
  Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if
  Set ScoreFile=FileObj.CreateTextFile(UserDirectory & HSFileName,True)
    ScoreFile.WriteLine 0
    ScoreFile.WriteLine Credits
    scorefile.writeline BallsPerGame
    scorefile.writeline ArrowCounter
    scorefile.writeline AdvanceSpecialMax
    scorefile.writeline BallLiftOption
    scorefile.writeline TiltEndsGame
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
  dim temp18


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
    temp6=textstr.readline
    temp7=textstr.readline
    temp8=textstr.readline

    HighScore=cdbl(temp1)
    if HighScore<1 then

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
    end if
    TextStr.Close
      Credits=cdbl(temp2)
    BallsPerGame=cdbl(temp3)
    ArrowCounter=cdbl(temp4)
    AdvanceSpecialMax=cdbl(temp5)
    BallLiftOption=cdbl(temp6)
    TiltEndsGame=cdbl(temp7)
    ReplayLevel=cdbl(temp8)
    if HighScore<1 then
      HSScore(1) = int(temp9)
      HSScore(2) = int(temp10)
      HSScore(3) = int(temp11)
      HSScore(4) = int(temp12)
      HSScore(5) = int(temp13)

      HSName(1) = temp14
      HSName(2) = temp15
      HSName(3) = temp16
      HSName(4) = temp17
      HSName(5) = temp18
    end if
    Set ScoreFile=Nothing
      Set FileObj=Nothing
end sub

Sub DisplayHighScore


end sub

Sub DisplayMatch
  MatchReels(Match).SetValue(1)
  If B2SOn then
    If Match = 0 then
      Controller.B2SSetMatch 10
    else
      Controller.B2SSetMatch Match
    end if
  end if
' if LowerLight4.state=1 then
'   MatchReels(MatchLit2).SetValue(1)
'   If B2SOn then
'     Controller.B2SSetData MatchLit2+70,1
'   end if
' end if

end sub

Sub ClearMatch
  For each obj in MatchReels
    obj.SetValue(0)
  next
  If B2SOn then
    For i = 70 to 79
      Controller.B2SSetData i,0
    next
  end if
end sub

sub InitPauser5_timer
  If B2SOn then
    Controller.B2SSetCredits Credits
  end if
    DisplayHighScore
    CreditsReel.SetValue(Credits)
    ArrowLights(ArrowCounter).state=1
    InitPauser5.enabled=false
end sub



sub BumpersOff

  Bumper1Light.visible=0
  Bumper2Light.visible=0
  Bumper3Light.visible=0

  Bumper5Light.visible=0
  Bumper6Light.visible=0
  KickLight1.state=0
  KickLight2.state=0
  KickLight3.state=0
  KickLight4.state=0

end sub

sub BumpersOn

  Bumper1Light.visible=1
  Bumper2Light.visible=1
  Bumper3Light.visible=1

  Bumper5Light.visible=1
  Bumper6Light.visible=1
  KickLight1.state=1
  KickLight2.state=1
  KickLight3.state=1
  KickLight4.state=1

end sub

Sub PlasticsOn

  For each obj in Flashers
    obj.visible=1
  next




end sub

Sub PlasticsOff

  For each obj in Flashers
    obj.visible=0
  next

  StopSound "buzz"
  StopSound "buzzL"

end sub

Sub SetupReplayTables

  Replay1Table(1)=900
  Replay1Table(2)=1000
  Replay1Table(3)=1100
  Replay1Table(4)=1200
  Replay1Table(5)=1100
  Replay1Table(6)=1100
  Replay1Table(7)=1200
  Replay1Table(8)=1200
  Replay1Table(9)=1300
  Replay1Table(10)=1300
  Replay1Table(11)=1400
  Replay1Table(12)=1400
  Replay1Table(13)=999000
  Replay1Table(14)=999000
  Replay1Table(15)=999000

  Replay2Table(1)=1100
  Replay2Table(2)=1100
  Replay2Table(3)=1200
  Replay2Table(4)=1300
  Replay2Table(5)=1200
  Replay2Table(6)=1300
  Replay2Table(7)=1300
  Replay2Table(8)=1400
  Replay2Table(9)=1400
  Replay2Table(10)=1500
  Replay2Table(11)=1500
  Replay2Table(12)=1600
  Replay2Table(13)=999000
  Replay2Table(14)=999000
  Replay2Table(15)=999000

  Replay3Table(1)=1200
  Replay3Table(2)=1200
  Replay3Table(3)=1300
  Replay3Table(4)=1400
  Replay3Table(5)=1300
  Replay3Table(6)=1400
  Replay3Table(7)=1400
  Replay3Table(8)=1500
  Replay3Table(9)=1500
  Replay3Table(10)=1600
  Replay3Table(11)=1600
  Replay3Table(12)=1700
  Replay3Table(13)=999000
  Replay3Table(14)=999000
  Replay3Table(15)=999000

  Replay4Table(1)=1300
  Replay4Table(2)=1300
  Replay4Table(3)=1400
  Replay4Table(4)=1500
  Replay4Table(5)=1400
  Replay4Table(6)=1500
  Replay4Table(7)=1500
  Replay4Table(8)=1600
  Replay4Table(9)=1600
  Replay4Table(10)=1700
  Replay4Table(11)=1700
  Replay4Table(12)=1800
  Replay4Table(13)=999000
  Replay4Table(14)=999000
  Replay4Table(15)=999000

  ReplayTableMax=4


end sub

Sub RefreshReplayCard
  Dim tempst1
  Dim tempst2

  tempst1=FormatNumber(BallsPerGame,0)
  tempst2=FormatNumber(ReplayLevel,0)


  InstructCard.image = "SC" + tempst2
  Replay1=Replay1Table(ReplayLevel)
  Replay2=Replay2Table(ReplayLevel)
  Replay3=Replay3Table(ReplayLevel)
  Replay4=Replay4Table(ReplayLevel)
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
      Case 1:
        AddScore(1)
        MotorRunning=0
        BumpersOn

      Case 2:
        MotorMode=1
        MotorPosition=2
        'BumpersOff
      Case 3:
        MotorMode=1
        MotorPosition=3
        'BumpersOff
      Case 4:
        MotorMode=1
        MotorPosition=4
        'BumpersOff
      Case 5:
        MotorMode=1
        MotorPosition=5
        'BumpersOff

      Case 10:
        AddScore(10)
        MotorRunning=0
        BumpersOn

      Case 20:
        MotorMode=10
        MotorPosition=2
        'BumpersOff
      Case 30:
        MotorMode=10
        MotorPosition=3
      Case 40:
        MotorMode=10
        MotorPosition=4
        'BumpersOff
      Case 50:
        MotorMode=10
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
        MotorPosition=5
        'BumpersOff
      Case 1000:
        MotorMode=1000
        MotorPosition=1
      Case 2000:
        MotorMode=1000
        MotorPosition=2
        'BumpersOff
      Case 3000:
        MotorMode=1000
        MotorPosition=3
        'BumpersOff
      Case 4000:
        MotorMode=1000
        MotorPosition=4
        'BumpersOff
      Case 5000:
        MotorMode=1000
        MotorPosition=5
        'BumpersOff
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
    if queuedscore>=40 then
      tempscore=40
      queuedscore=queuedscore-40
      SetMotor2(40)
      exit sub
    end if
    if queuedscore>=30 then
      tempscore=30
      queuedscore=queuedscore-30
      SetMotor2(30)
      exit sub
    end if
    if queuedscore>=20 then
      tempscore=20
      queuedscore=queuedscore-20
      SetMotor2(20)
      exit sub
    end if
    if queuedscore>=10 then
      tempscore=10
      queuedscore=queuedscore-10
      SetMotor2(10)
      exit sub
    end if
    if queuedscore>=5 then
      tempscore=5
      queuedscore=queuedscore-5
      SetMotor2(5)
      exit sub
    end if
    if queuedscore>=4 then
      tempscore=4
      queuedscore=queuedscore-4
      SetMotor2(4)
      exit sub
    end if
    if queuedscore>=3 then
      tempscore=3
      queuedscore=queuedscore-3
      SetMotor2(3)
      exit sub
    end if
    if queuedscore>=2 then
      tempscore=2
      queuedscore=queuedscore-2
      SetMotor2(2)
      exit sub
    end if
    if queuedscore>=1 then
      tempscore=1
      queuedscore=queuedscore-1
      SetMotor2(1)
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
        If MotorMode=1 then
          AddScore(1)
        end if
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
        If MotorMode=1 then
          AddScore(1)
        end if
        MotorPosition=0:MotorRunning=0:BumpersOn
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
      Score(Player)=Score(Player)+1

    Case 10:
      PlayChime(10)
      Score(Player)=Score(Player)+10
'     debugscore=debugscore+10
      ToggleAlternatingRelay
    Case 100:
      PlayChime(100)
      Score(Player)=Score(Player)+100
'     debugscore=debugscore+100


    Case 1000:
      PlayChime(100)
      Score(Player)=Score(Player)+1000
'     debugscore=debugscore+1000
  End Select
  PlayerScores(0).AddValue(x)
  PlayerScoresOn(0).AddValue(x)
  If ScoreDisplay(Player)<100000 then
    ScoreDisplay(Player)=Score(Player)
  Else
    Score100K(Player)=Int(Score(Player)/100000)
    ScoreDisplay(Player)=Score(Player)-100000
  End If
  if Score(Player)=>100000 then
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
  If Score(Player)>Replay1 and Replay1Paid(Player)=false then
    Replay1Paid(Player)=True
    AddSpecial
  End If
  If Score(Player)>Replay2 and Replay2Paid(Player)=false then
    Replay2Paid(Player)=True
    AddSpecial
  End If
  If Score(Player)>Replay3 and Replay3Paid(Player)=false then
    Replay3Paid(Player)=True
    AddSpecial
  End If
  If Score(Player)>Replay4 and Replay4Paid(Player)=false then
    Replay4Paid(Player)=True
    AddSpecial
  End If
' ScoreText.text=debugscore
End Sub

Sub AddScore2(x)
  Dim OldScore, NewScore, OldTestScore, NewTestScore
    OldScore = Score(1)

  Select Case x
        Case 1:
            Score(1)=Score(1)+1
    Case 10:
      Score(1)=Score(1)+10
    Case 100:
      Score(1)=Score(1)+100
    Case 1000:
      Score(1)=Score(1)+1000
  End Select
  NewScore = Score(1)

  OldTestScore = OldScore
  NewTestScore = NewScore
  Do
    if OldTestScore < Replay1 and NewTestScore >= Replay1 then
      AddSpecial()
      NewTestScore = 0
    Elseif OldTestScore < Replay2 and NewTestScore >= Replay2 then
      AddSpecial()
      NewTestScore = 0
    Elseif OldTestScore < Replay3 and NewTestScore >= Replay3 then
      AddSpecial()
      NewTestScore = 0
    Elseif OldTestScore < Replay4 and NewTestScore >= Replay4 then
      AddSpecial()
      NewTestScore = 0
    End if
    NewTestScore = NewTestScore - 100000
    OldTestScore = OldTestScore - 100000
  Loop While NewTestScore > 0

    OldScore = int(OldScore / 1)  ' divide by 10 for games with fixed 0 in 1s position, by 1 for games with real 1s digits
    NewScore = int(NewScore / 1)  ' divide by 10 for games with fixed 0 in 1s position, by 1 for games with real 1s digits
  ' MsgBox("OldScore="&OldScore&", NewScore="&NewScore&", OldScore Mod 10="&OldScore Mod 10 & ", NewScore % 10="&NewScore Mod 10)

    if (OldScore Mod 10 <> NewScore Mod 10) then
    PlayChime(10)
    AdvanceZeroToNine
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
  If Score(1)>999 then
    Reel1000.SetValue(1)
  end if
  If B2SOn Then
    Controller.B2SSetScorePlayer 1, Score(1)
    If Score(1)>999 then
      Controller.B2SSetData 99,1
    end if
  End If
' EMReel1.SetValue Score(Player)
  PlayerScores(0).AddValue(x)
  PlayerScoresOn(0).AddValue(x)
End Sub



Sub PlayChime(x)
  if ChimesOn=0 then
    Select Case x
      Case 10
        If LastChime10=1 Then
          'PlaySound "SpinACard_1_10_Point_Bell"
          PlaySound SoundFXDOF("SpinACard_1_10_Point_Bell",146,DOFPulse,DOFChimes)
          LastChime10=0
        Else
          'PlaySound "SpinACard_1_10_Point_Bell"
          PlaySound SoundFXDOF("SpinACard_1_10_Point_Bell",146,DOFPulse,DOFChimes)
          LastChime10=1
        End If
      Case 100
        If LastChime100=1 Then
          'PlaySound "SpinACard_100_Point_Bell"
          PlaySound SoundFXDOF("SpinACard_100_Point_Bell",147,DOFPulse,DOFChimes)
          LastChime100=0
        Else
          'PlaySound "SpinACard_100_Point_Bell"
          PlaySound SoundFXDOF("SpinACard_100_Point_Bell",147,DOFPulse,DOFChimes)
          LastChime100=1
        End If

    End Select
  else
    Select Case x
      Case 10
        If LastChime10=1 Then
          PlaySound SoundFXDOF("SJ_Chime_10a",146,DOFPulse,DOFChimes)
          LastChime10=0
        Else
          PlaySound SoundFXDOF("SJ_Chime_10b",146,DOFPulse,DOFChimes)
          LastChime10=1
        End If
      Case 100
        If LastChime100=1 Then
          PlaySound SoundFXDOF("SJ_Chime_100a",147,DOFPulse,DOFChimes)
          LastChime100=0
        Else
          PlaySound SoundFXDOF("SJ_Chime_100b",147,DOFPulse,DOFChimes)
          LastChime100=1
        End If
      Case 1000
        If LastChime1000=1 Then
          PlaySound SoundFXDOF("SJ_Chime_1000a",148,DOFPulse,DOFChimes)
          LastChime1000=0
        Else
          PlaySound SoundFXDOF("SJ_Chime_1000b",148,DOFPulse,DOFChimes)
          LastChime1000=1
        End If
    End Select
  end if
End Sub

Sub HideOptions()

end sub

' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX and Rothbauerw
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

'Set all as per ball position & speed.

Sub PlaySoundAtBall(soundname)
  PlaySoundAt soundname, ActiveBall
End Sub

'Set position as table object and Vol manually.

Sub PlaySoundAtVol(sound, tableobj, Volume)
  PlaySound sound, 1, Volume, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLoopVol(sound, tableobj, Volume)
  PlaySound sound, -1, Volume, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'Set all as per ball position & speed, but Vol Multiplier may be used eg; PlaySoundAtBallVol "sound",3

Sub PlaySoundAtBallVol(sound, VolMult)
  PlaySound sound, 0, Vol(ActiveBall) * VolMult, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub

'Set position as bumperX and Vol manually.

Sub PlaySoundAtBumperVol(sound, tableobj, Vol)
  PlaySound sound, 1, Vol, AudioPan(tableobj), 0,0,1, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBOTBallZ(sound, BOT)
    PlaySound sound, 0, ABS(BOT.velz)/17, Pan(BOT), 0, Pitch(BOT), 1, 0, AudioFade(BOT)
End Sub

'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

Function RndNum(min, max)
    RndNum = Int(Rnd() * (max-min + 1) ) + min ' Sets a random number between min and max
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
  Vol = Csng(BallVel(ball) ^2 / VolDiv)
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
  PlaySound("fx_collide"), 0, Csng(velocity) ^2 / (VolCol/VolDiv), AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
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

Sub Spinner_Spin
  PlaySound "fx_spinner", 0, .25, AudioPan(Spinner), 0.25, 0, 0, 1, AudioFade(Spinner)
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

' ============================================================================================
' GNMOD - Multiple High Score Display and Collection
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
HSScore(1) = 800
HSScore(2) = 700
HSScore(3) = 600
HSScore(4) = 500
HSScore(5) = 400

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
    HighScoreReel.SetValue(HSScore(1))
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
    HighScoreReel.SetValue(HSScore(1))
    If B2SOn Then
      Controller.B2SSetScorePlayer 2, HSScore(1)
    End if
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
Const OptionLinesToMark="011011111"
Const OptionLine1="" 'do not use this line
Const OptionLine2="" 'do not use this line
Const OptionLine3="" 'do not use this line
Const OptionLine4=""
Const OptionLine5="Top Rollover Lighting"
Const OptionLine6="Ball Lifter"
Const OptionLine7="Tilt Penalty"
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
  CurrentOption = 2
  DisplayAllOptions
  OperatorOption1.image = "BluePlus"
  SetHighScoreOption

End Sub

Sub DisplayAllOptions
  dim linecounter
  dim tempstring
  For linecounter = 2 to MaxOption
    tempstring=Eval("OptionLine"&linecounter)
    Select Case linecounter
      Case 1:
        tempstring=""
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
        Select Case AdvanceSpecialMax
          Case 1:
            tempstring = "Conservative"
          Case 2:
            tempstring = "Medium"
          Case 3:
            tempstring = "Liberal"
        end select
        SetOptLine 7, tempstring

      Case 6:
        SetOptLine 8, tempstring
        If BallLiftOption=1 then
          tempstring = "Manual use Start Key"
        else
          tempstring = "Automatic"
        end if
        SetOptLine 9, tempstring

      Case 7:
        SetOptLine 10, tempstring
        If TiltEndsGame=0 then
          tempstring="Current balls on table"
        else
          tempstring="Tilt ends game"
        end if
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
        BallsPerGame = 5
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
      AdvanceSpecialMax=AdvanceSpecialMax+1
      If AdvanceSpecialMax>3 then
        AdvanceSpecialMax=1
      end if
      DisplayAllOptions
    elseif CurrentOption = 6 then
      if BallLiftOption=1 then
        BallLiftOption=2
      else
        BallLiftOption=1
      end if
      DisplayAllOptions
    elseif CurrentOption = 7 then
      if TiltEndsGame=1 then
        TiltEndsGame=0
      else
        TiltEndsGame=1
      end if
      DisplayAllOptions
    elseif CurrentOption = 8 or CurrentOption = 9 then
        if OptionCHS=1 then
          HSScore(1) = 800
          HSScore(2) = 700
          HSScore(3) = 600
          HSScore(4) = 500
          HSScore(5) = 400

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
        'InstructCard.image="IC"+FormatNumber(BallsPerGame,0)+FormatNumber(AdvanceSpecialMax,0)
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



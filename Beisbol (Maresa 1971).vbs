'*
'*        Beisbol (Maresa 1971)
'*        Gottlieb's "Baseball" (1970) build/scripted by Loserman76
'*        Playfield/plastics all redrawn by GNance
'*        DOF config By Foxyt
'*      Conversion to Besibol by CactusDude
' Thalamus - tested ok

option explicit
Randomize
ExecuteGlobal GetTextFile("core.vbs")

On Error Resume Next
ExecuteGlobal GetTextFile("Controller.vbs")
If Err Then MsgBox "Unable to open Controller.vbs. Ensure that it is in the scripts folder."
On Error Goto 0

Const cGameName = "Beisbol_1971"

Const ShadowFlippersOn = true
Const ShadowBallOn = true

Const ShadowConfigFile = false

Dim Controller  ' B2S
Dim B2SScore  ' B2S Score Displayed

Const HSFileName="Besibol_71VPX.txt"
Const B2STableName="Beisbol_1971"
Const LMEMTableConfig="LMEMTables.txt"
Const LMEMShadowConfig="LMEMShadows.txt"
Dim EnableBallShadow
Dim EnableFlipperShadow


'* this value adjusts score motor behavior - 0 allows you to continue scoring while the score motor is running - 1 sets score motor to behave more like a real EM
Const ScoreMotorAdjustment=1

'* this is a debug setting to use an older scoring routine vs a newer score routine - don't change this value
Const ScoreAdditionAdjustment=1

dim ScoreChecker
dim CheckAllScores,CheckAllPoints
dim sortscores(4)
dim sortpoints(4)
dim sortplayerpoints(4)
dim sortplayers(4)
Dim B2SOn   'True/False if want backglass
dim ChimesOn
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
dim TempPlayerup
dim RotatorTemp

dim GreenTargetsDownCounter
dim BlueTargetsDownCounter
dim YellowTargetsDownCounter

Dim bump1
Dim bump2
Dim bump3
Dim bump4

Dim LastChime10
Dim LastChime100
Dim LastChime1000

Dim Score10
Dim Score100

dim tempbumper

Dim MotorRunning

Dim Replay1Table(15)
Dim Replay2Table(15)
Dim Replay3Table(15)
Dim Replay4Table(15)
Dim ReplayBallsTable1(16)
Dim ReplayBallsTable2(16)
Dim ReplayBallsTable3(16)

Dim ReplayTableSet
Dim ReplayLevel
Dim ReplayBalls
Dim ReplayTableMax
Dim ReplayBallsTableMax

Dim BaseHit,BaseHitCounter


Sub Table1_Init()
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
  HideOptions
  SetupReplayTables
  PlasticsOff
  BumpersOff
  OperatorMenu=0
  ChimesOn=1
  ReplayLevel=1
  ReplayBalls=1
  BallsPerGame=5
  EMReel1.ResetToZero
  EMReel2.ResetToZero
  HighScore=0
  MotorRunning=0
  BaseHit=0
  BaseHitCounter=0
  HighScoreReward=1
  BallTextBox.text=""
  MatchTextBox.text=""
  Credits=0
  loadhs
  if HighScore=0 then HighScore=5000



  TableTilted=false
  TiltReel.SetValue(1)

  Match=int(Rnd*10)
  MatchReel.SetValue((Match)+1)

  CanPlayReel.SetValue(0)
  BaseRunner1.SetValue(1)
  BaseRunner2.SetValue(1)
  BaseRunner3.SetValue(1)

  For each obj in PlayerHuds
    obj.SetValue(0)
  next
  For each obj in PlayerHUDScores
    obj.state=0
  next
  For each obj in PlayerScores
    obj.ResetToZero
  next

  Replay1=Replay1Table(ReplayLevel)
  Replay2=Replay2Table(ReplayLevel)
  Replay3=Replay3Table(ReplayLevel)
  Replay4=Replay4Table(ReplayLevel)
  BallReplay1=ReplayBallsTable1(ReplayBalls)
  BallReplay2=ReplayBallsTable2(ReplayBalls)
  BallReplay3=ReplayBallsTable3(ReplayBalls)


  BonusCounter=0
  HoleCounter=0
    bgpos=6
  kgpos=0



  dooralreadyopen=0
  kgdooralreadyopen=0

  Bumper1Light.state=0
  Bumper2Light.state=0
  Bumper1Light.visible=1
  Bumper2Light.visible=1


  InstructCard.image="IC"+FormatNumber(BallsPerGame,0)+"0"

  RefreshReplayCard
' ResetBalls
' PlaySound "Plunger"

  TargetSpecialLit = 0
  Points210counter=0
  Points500counter=0
  Points1000counter=0
  Points2000counter=0
  BonusBooster=3
  BonusBoosterCounter=0
  Players=0
  RotatorTemp=1
  InProgress=false

  ScoreText.text=HighScore


  If B2SOn Then


    if Match=0 then
      Controller.B2SSetMatch 10
    else
      Controller.B2SSetMatch Match
    end if
    Controller.B2SSetScoreRolloverPlayer1 0

    'Controller.B2SSetScore 4,HighScore
    Controller.B2SSetTilt 1
    Controller.B2SSetCredits Credits
    Controller.B2SSetGameOver 1
  End If
  CreditsReel.SetValue(Credits)
  If Credits > 0 Then DOF 125, DOFOn
  for i=1 to 2
    player=i
    If B2SOn Then
      Controller.B2SSetScorePlayer player, 0
    End If
  next
  bump1=1
  bump2=1
  bump3=1
  bump4=1

  InitPauser5.enabled=true
End Sub

Sub Table1_exit()
  savehs
  SaveLMEMConfig
  SaveLMEMConfig2
  If B2SOn Then Controller.Stop
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

  if keycode = LeftFlipperKey and InProgress = false then
    OperatorMenuTimer.Enabled = true
  end if
  ' END GNMOD


  If keycode = LeftFlipperKey and InProgress=true and TableTilted=false Then
    LeftFlipper.RotateToEnd
    PlaySoundAtVol SoundFXDOF("FlipperUp",101,DOFOn,DOFFlippers), LeftFlipper, 1
    PlayLoopSoundAtVol "buzzL", LeftFlipper, 1
  End If

  If keycode = RightFlipperKey  and InProgress=true and TableTilted=false Then
    RightFlipper.RotateToEnd
    PlaySoundAtVol SoundFXDOF("FlipperUp",102,DOFOn,DOFFlippers), RightFlipper, 1
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
    AddSpecial2
  end if

   if keycode = 5 then
    PlaySoundAtVol "coinin", Drain, 1
    AddSpecial2
    keycode= StartGameKey
  end if


  if keycode=StartGameKey and Credits>0 and InProgress=false and Players=0 and EnteringOptions = 0 then
'GNMOD
    OperatorMenuTimer.Enabled = false
'END GNMODthen
    Credits=Credits-1
    If Credits < 1 Then DOF 125, DOFOff
    CreditsReel.SetValue(Credits)
    Players=1
    CanPlayReel.SetValue(Players)
    MatchReel.SetValue(0)
    Player=1
    playsound "startup_norm"
    TempPlayerUp=Player
'   PlayerUpRotator.enabled=true
    rst=0
    BallInPlay=1
    InProgress=true
    resettimer.enabled=true
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
    End If
    For each obj in PlayerScores
'     obj.ResetToZero
    next
    For each obj in PlayerHuds
      obj.SetValue(0)
    next
    If Table1.ShowDT = True then
      For each obj in PlayerScores
'       obj.ResetToZero
        obj.Visible=true
      next
      For each obj in PlayerScoresOn
'       obj.ResetToZero
        obj.Visible=false
      next

      For each obj in PlayerHuds
        obj.SetValue(0)
      next
      For each obj in PlayerHUDScores
        obj.state=0
      next
      PlayerHuds(Player-1).SetValue(1)
      PlayerHUDScores(Player-1).state=1
      PlayerScores(Player-1).Visible=0
      PlayerScoresOn(Player-1).Visible=1
      PlayerHuds(Player).SetValue(1)
      PlayerHUDScores(Player).state=1
      PlayerScores(Player).Visible=0
      PlayerScoresOn(Player).Visible=1

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

    PlaySound"plungerrelease"
    Plunger.Fire
  End If

  if keycode = LeftFlipperKey then
    OperatorMenuTimer.Enabled = false
  end if

  ' END GNMOD

  If keycode = LeftFlipperKey and InProgress=true and TableTilted=false Then
    LeftFlipper.RotateToStart
    PlaySoundAtVol SoundFXDOF("FlipperDown",101,DOFOff,DOFFlippers), LeftFlipper, 1
    StopSound "buzzL"
  End If

  If keycode = RightFlipperKey and InProgress=true and TableTilted=false Then
    RightFlipper.RotateToStart
    PlaySoundAtVol SoundFXDOF("FlipperDown",102,DOFOff,DOFFlippers), RightFlipper, 1
        StopSound "buzz"
  End If

End Sub



Sub Drain_Hit()
  Drain.DestroyBall
  PlaySoundAtVol "fx_drain", Drain, 1
  DOF 250, DOFPulse
' AddBonus
  Pause4Bonustimer.enabled=true


End Sub


Sub Pause4Bonustimer_timer
  Pause4Bonustimer.enabled=0
  NextBallDelay.enabled=true

End Sub

Sub TriggerLaunch_hit
  DOF 260, DOFPulse
End Sub

'***********************
'     Flipper Logos
'***********************

Sub UpdateFlipperLogos_Timer
  LFlip.RotZ = LeftFlipper.CurrentAngle
  LMaresa.RotZ = LeftFlipper.CurrentAngle
  RFlip.RotZ = RightFlipper.CurrentAngle
  LFlip1.RotZ = LeftFlipper.CurrentAngle
  RFlip1.RotZ = RightFlipper.CurrentAngle
  RMaresa.RotZ = RightFlipper.CurrentAngle

  PGate.Rotz = (Gate.CurrentAngle*.75) + 25
  FlipperLSh.RotZ = LeftFlipper.currentangle
  FlipperRSh.RotZ = RightFlipper.currentangle
End Sub

'***********************
' slingshots
'

Sub RightSlingShot_Slingshot
    PlaySoundAtVol SoundFXDOF("right_slingshot",104,DOFPulse,DOFContactors), ActiveBall, 1
  DOF 204, DOFPulse
    RSling0.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1

  AddScore(10)

End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSling0.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
    PlaySoundAtVol SoundFXDOF("left_slingshot",103,DOFPulse,DOFContactors), ActiveBall, 1
  DOF 203, DOFPulse
    LSling0.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1

  AddScore(10)


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
    If idx=0 or idx=1 then
      AddScore(10)
    else
      AddScore(1)
    end if
  end if
end Sub


'***********************************

Sub TopTargetLeft_Hit()
  if TableTilted=false then
    DOF 106, DOFPulse
    DOF 201, DOFPulse
    BaseHit=0
    SetMotor(50)
    TargetLightBase21.state=0
    TargetLightBase22.state=0
    TargetLightBase23.state=0
    TargetLightBase24.state=0
    LightBase2.state=1
    If B2SOn then
      Controller.B2SSetData 52,1
    end if
    BaseRunner2.SetValue(1)
    LightTop3.state=0
    BumpersOn
  end if
end sub

Sub TopTargetRight_Hit()
  if TableTilted=false then
    DOF 106, DOFPulse
    DOF 202, DOFPulse
    BaseHit=0
    SetMotor(50)
    TargetLightBase21.state=0
    TargetLightBase22.state=0
    TargetLightBase23.state=0
    TargetLightBase24.state=0
    LightBase2.state=1
    If B2SOn then
      Controller.B2SSetData 52,1
    end if
    BaseRunner2.SetValue(1)
    LightTop3.state=0
    BumpersOn
  end if
end sub

Sub LowerTargetLeft_Hit()
  if TableTilted=false then
    DOF 109, DOFPulse
    DOF 206, DOFPulse
    BaseHit=0
    AddScore(100)
    TargetLightBase3.state=0
    TargetLightBase31.state=0
    LightBase3.state=1
    If B2SOn then
      Controller.B2SSetData 53,1
    end if
    BaseRunner3.SetValue(1)
    LightTop1.state=0
    BumpersOn
  end if
end sub

Sub LowerTargetRight_Hit()
  if TableTilted=false then
    DOF 109, DOFPulse
    DOF 209, DOFPulse
    BaseHit=0
    AddScore(100)
    TargetLightBase1.state=0
    TargetLightBase11.state=0
    LightBase1.state=1
    If B2SOn then
      Controller.B2SSetData 51,1
    end if
    BaseRunner1.SetValue(1)
    LightTop2.state=0
    BumpersOn
  end if
end sub

'***********************************
Sub Bumper1_Hit
  If TableTilted=false then
    PlaySoundAtVol SoundFXDOF("bumper1",105,DOFPulse,DOFContactors), ActiveBall, 1
    DOF 205, DOFPulse
    bump1 = 1

    If Bumper1Light.state=1 then
      AddScore(100)
    else
      AddScore(10)
    End If
    end if

End Sub


Sub Bumper2_Hit
  If TableTilted=false then
    PlaySoundAtVol SoundFXDOF("bumper1",107,DOFPulse,DOFContactors), ActiveBall, 1
    DOF 207, DOFPulse
    bump2 = 1

    If Bumper2Light.state=1 then
      AddScore(100)
    else
      AddScore(10)
    End If
    end if

End Sub


'***********************************


Sub TriggerTop1_Hit()
  If TableTilted=false then
    PlaySoundAtVol "sensor", ActiveBall, 1
    DOF 211, DOFPulse
    BaseHit=0
    SetMotor(50)
    LightBase3.state=1
    LightTop1.state=0
    TargetLightBase3.state=0
    TargetLightBase31.state=0
    If B2SOn then
      Controller.B2SSetData 53,1
    end if
    BaseRunner3.SetValue(1)
    BumpersOn
  end if
End Sub

Sub TriggerTop2_Hit()
  If TableTilted=false then
    DOF 212, DOFPulse
    PlaySoundAtVol "sensor", ActiveBall, 1
    BaseHit=0
    SetMotor(50)
    LightBase1.state=1
    LightTop2.state=0
    TargetLightBase1.state=0
    TargetLightBase11.state=0
    If B2SOn then
      Controller.B2SSetData 51,1
    end if
    BaseRunner1.SetValue(1)
    BumpersOn
  end if
End Sub

Sub TriggerTop3_Hit()
  If TableTilted=false then
    PlaySoundAtVol "sensor", ActiveBall, 1
    DOF 213, DOFPulse
    BaseHit=0
    SetMotor(50)
    LightBase2.state=1
    LightTop3.state=0
    TargetLightBase21.state=0
    TargetLightBase22.state=0
    TargetLightBase23.state=0
    TargetLightBase24.state=0
    If B2SOn then
      Controller.B2SSetData 52,1
    end if
    BaseRunner2.SetValue(1)
    BumpersOn
  end if
End Sub

Sub TriggerLeftRollover_Hit()
  If TableTilted=false then
    PlaySoundAtVol "sensor", ActiveBall, 1
    DOF 216, DOFPulse
    BaseHit=2
    SetMotor(200)

  end if
End Sub

Sub TriggerRightRollover_Hit()
  If TableTilted=false then
    PlaySoundAtVol  "sensor", ActiveBall, 1
    DOF 219, DOFPulse
    BaseHit=3
    SetMotor(300)

  end if
End Sub

Sub TriggerLeftInlane_Hit()
  If TableTilted=false then
    PlaySoundAtVol  "sensor", ActiveBall, 1
    DOF 215, DOFPulse
    BaseHit=1
    SetMotor(100)

  End If
End Sub

Sub TriggerRightInlane_Hit()
  If TableTilted=false then
    PlaySoundAtVol  "sensor", ActiveBall, 1
    DOF 218, DOFPulse
    BaseHit=1
    SetMotor(100)

  End If
End Sub


Sub TriggerLeftOutlane_Hit()
  If TableTilted=false then
    PlaySoundAtVol  "sensor", ActiveBall, 1
    DOF 214, DOFPulse
    BaseHit=4
    SetMotor(500)

  End If
End Sub

Sub TriggerRightOutlane_Hit()
  If TableTilted=false then
    PlaySoundAtVol  "sensor", ActiveBall, 1
    DOF 217, DOFPulse
    BaseHit=4
    SetMotor(500)

  End If
End Sub

Sub MoveRunners
  If LightBase3.state=1 then
    AddRunToScore
    LightBase3.state=0
    LightTop1.state=1
    TargetLightBase3.state=1
    TargetLightBase31.state=1
    If B2SOn then
      Controller.B2SSetData 53,0
    end if
      BaseRunner3.SetValue(0)
  end if
  If LightBase2.state=1 then
    LightBase3.state=1
    LightBase2.state=0
    LightTop3.state=1
    TargetLightBase21.state=1
    TargetLightBase22.state=1
    TargetLightBase23.state=1
    TargetLightBase24.state=1
    LightTop1.state=0
    TargetLightBase3.state=0
    TargetLightBase31.state=0
    If B2SOn then
      Controller.B2SSetData 53,1
      Controller.B2SSetData 52,0
    end if
    BaseRunner3.SetValue(1)
    BaseRunner2.SetValue(0)
  end if
  If LightBase1.state=1 then
    LightBase2.state=1
    LightBase1.state=0
    LightTop2.state=1
    TargetLightBase1.state=1
    TargetLightBase11.state=1
    LightTop3.state=0
    TargetLightBase21.state=0
    TargetLightBase22.state=0
    TargetLightBase23.state=0
    TargetLightBase24.state=0
    If B2SOn then
      Controller.B2SSetData 52,1
      Controller.B2SSetData 51,0
    end if
    BaseRunner2.SetValue(1)
    BaseRunner1.SetValue(0)
  end if


end sub

'****************************************************
' Star Triggers

Sub TriggerHome_Hit
  If TableTilted=true then
    Exit Sub
  end if
  if LightBaseHome.state=1 then
      DOF 230, DOFPulse
    AddScore(10)
  else
    AddScore(1)
  end if
end sub

Sub TriggerBase1_Hit
  If TableTilted=true then
    Exit Sub
  end if
  if LightBase1.state=1 then
      DOF 231, DOFPulse
    AddScore(10)
  else
    AddScore(1)
  end if
end sub

Sub TriggerBase2_Hit
  If TableTilted=true then
    Exit Sub
  end if
  if LightBase2.state=1 then
      DOF 232, DOFPulse
    AddScore(10)
  else
    AddScore(1)
  end if
end sub

Sub TriggerBase3_Hit
  If TableTilted=true then
    Exit Sub
  end if
  if LightBase3.state=1 then
      DOF 233, DOFPulse
    AddScore(10)
  else
    AddScore(1)
  end if
end sub

'****************************************************
' Varitarget

Dim Lvariawarded, Lnewpos, Loldpos
Dim Rvariawarded, Rnewpos, Roldpos
Dim LeftPosition_Vari
Dim RightPosition_Vari
Dim Varimultiplier

Lvariawarded = 0
Lnewpos=0
Loldpos=0


Sub LeftVT_Hit(idx)
  Dim x
  Lnewpos=idx
  If ActiveBall.VelY <0 Then
    DOF 108, DOFOn
    DOF 208, DOFOn
    DOF 129, DOFOn
    ActiveBall.VelY = ActiveBall.VelY * 0.915
    PlaySoundAtVol  "fx_solenoidoff", ActiveBall, 1
'   LeftVari(Loldpos).IsDropped = True
'   LeftVari(Lnewpos).IsDropped = False
    VariTargetL.TransX = -(LeftPosition_Vari*10)
    LeftPosition_Vari = idx +1
  End If
  Loldpos=Lnewpos
End Sub

Sub LeftVT_UnHit(idx)
  If ActiveBall.VelY >0 Then
    DOF 108, DOFOff
    DOF 208, DOFOff
    DOF 129, DOFOff
    TimerLeftVariTarget.Enabled = True
  End If
End Sub

Sub vt18_Hit()
  Dim x
  Lvariawarded = 1
  DOF 109,DOFPulse
  LeftPosition_Vari=17
  BaseHit=4
    SetMotor(500)
End Sub

Sub vt36_Hit()
  Dim x
  Rvariawarded = 1
  DOF 109,DOFPulse
  RightPosition_Vari=17
  BaseHit=4
    SetMotor(500)
End Sub

Sub RightVT_Hit(idx)
  Dim x
  Rnewpos=idx
  If ActiveBall.VelY <0 Then
    DOF 110, DOFOn
    DOF 210, DOFOn
    DOF 129, DOFOn
    ActiveBall.VelY = ActiveBall.VelY * 0.915
    PlaySoundAtVol  "fx_solenoidoff", ActiveBall, 1
'   RightVari(Roldpos).IsDropped = True
'   RightVari(Rnewpos).IsDropped = False
    VariTargetR.TransX = -(RightPosition_Vari*10)
    RightPosition_Vari = idx +1
  End If
  Roldpos=Rnewpos
End Sub

Sub RightVT_UnHit(idx)
  If ActiveBall.VelY >0 Then
    DOF 110, DOFOff
    DOF 210, DOFOff
    DOF 129, DOFOff
    TimerRightVariTarget.Enabled = True
  End If
End Sub


Sub TimerRightVariTarget_Timer()
'        RightVari(RightPosition_Vari).IsDropped = True
        RightPosition_Vari = RightPosition_Vari - 1
    If RightPosition_Vari <0 Then
      RightPosition_Vari = 0
    End If

 '       RightVari(RightPosition_Vari).IsDropped = False
    VariTargetR.TransX = -(RightPosition_Vari*10)
    If RightPosition_Vari = 15 AND Rvariawarded = 0 Then
      SetMotor(300): BaseHit=3:Rvariawarded = 1
    End If
    If RightPosition_Vari = 11 AND Rvariawarded = 0 Then
      SetMotor(200): BaseHit=2:Rvariawarded = 1
    End If
    If RightPosition_Vari = 8 AND Rvariawarded = 0 Then
      SetMotor(200): BaseHit=2:Rvariawarded = 1
    End If
    If RightPosition_Vari = 4 AND Rvariawarded = 0 Then
      SetMotor(100): BaseHit=1:Rvariawarded = 1
    End If
    If RightPosition_Vari = 1 AND Rvariawarded = 0 Then
      SetMotor(100): BaseHit=1:Rvariawarded = 1
    End If
        If RightPosition_Vari = 0 Then
      Rvariawarded = 0
        TimerRightVariTarget.Enabled = False
    End If
End Sub

Sub TimerLeftVariTarget_Timer()
'        LeftVari(LeftPosition_Vari).IsDropped = True
        LeftPosition_Vari = LeftPosition_Vari - 1
    If LeftPosition_Vari <0 Then
      LeftPosition_Vari = 0
    End If

'       LeftVari(LeftPosition_Vari).IsDropped = False
    VariTargetL.TransX = -(LeftPosition_Vari*10)
    If LeftPosition_Vari = 15 AND Lvariawarded = 0 Then
      SetMotor(300): BaseHit=3:Lvariawarded = 1
    End If
    If LeftPosition_Vari = 11 AND Lvariawarded = 0 Then
      SetMotor(200): BaseHit=2:Lvariawarded = 1
    End If
    If LeftPosition_Vari = 8 AND Lvariawarded = 0 Then
      SetMotor(200): BaseHit=2:Lvariawarded = 1
    End If
    If LeftPosition_Vari = 4 AND Lvariawarded = 0 Then
      SetMotor(100): BaseHit=1:Lvariawarded = 1
    End If
    If LeftPosition_Vari = 1 AND Lvariawarded = 0 Then
      SetMotor(100): BaseHit=1:Lvariawarded = 1
    End If
        If LeftPosition_Vari = 0 Then
      Lvariawarded = 0
        TimerLeftVariTarget.Enabled = False
    End If
End Sub

'****************************************************



Sub AddSpecial()
  PlaySound SoundFXDOF("knocker",120,DOFPulse,DOFContactors)
  DOF 121, DOFPulse
  Credits=Credits+1
  DOF 125, DOFOn
  if Credits>15 then Credits=15
  If B2SOn Then
    Controller.B2SSetCredits Credits
  End If
  CreditsReel.SetValue(Credits)
End Sub

Sub AddSpecial2()
  PlaySound"click"
  Credits=Credits+1
  DOF 125, DOFOn
  if Credits>15 then Credits=15
  If B2SOn Then
    Controller.B2SSetCredits Credits
  End If
  CreditsReel.SetValue(Credits)
End Sub

Sub AddBonus()
  bonuscountdown=bonuscounter
  ScoreBonus.enabled=true
End Sub


Sub ToggleOutlane

end sub

Sub ToggleBumper


end sub


Sub ResetBallDrops()



  BonusCounter=0
  HoleCounter=0
End Sub


Sub LightsOut
  BonusCounter=0
  HoleCounter=0

  Bumper1Light.state=0

  Bumper2Light.state=0


end sub

Sub ResetDrops()


End Sub

Sub LightsOff


end sub

Sub LightsOn

end sub


Sub FireAlternatingRelay

end sub


Sub ResetBalls()

  TempMultiCounter=BallsPerGame-BallInPlay



  BonusMultiplier=1

  TableTilted=false
  TiltReel.SetValue(0)

  PlasticsOn

  'CreateBallID BallRelease
  Ballrelease.CreateSizedBall 25
    Ballrelease.Kick 40,7
    DOF 122,DOFPulse
    InstructCard.image="IC"+FormatNumber(BallsPerGame,0)+FormatNumber(BallInPlay,0)


  'BallInPlayReel.SetValue(BallInPlay)

End Sub

sub delaykgclose_timer
  delaykgclose.enabled=false
  closekg.enabled=true

end sub

sub resettimer_timer
    rst=rst+1
    rst=rst+1
  if rst>1 and rst<12 then
    ResetReelsToZero(1)
  end if

    if rst=13 then
    'playsound "StartBall1"
    'playsound SoundFXDOF("StartBall1",122,DOFPulse,DOFContactors)
    end if
    if rst=14 then
    newgame
    resettimer.enabled=false
    end if
end sub

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


sub ScoreBonus_timer

  if (Bonus(bonuscountdown).state=1) and (bonuscountdown>0) then
    AddScore2(1000)
    Bonus(bonuscountdown).state=0
    If BonusMultiplier=1 then
      ScoreBonus.interval=300
    Else
      ScoreBonus.interval=300
    end if
  else
    ScoreBonus.interval=10
  end if
  bonuscountdown=bonuscountdown-1
  if bonuscountdown<=0 then
    ScoreBonus.enabled=false
    ScoreBonus.interval=600
    NextBallDelay.enabled=true
  else
    Bonus(bonuscountdown).state=1
  end if

end sub

sub NextBallDelay_timer()
  NextBallDelay.enabled=false
  nextball

end sub

sub newgame
  InProgress=true
  debugscore=0
  EMReel1.ResetToZero
  EMReel2.ResetToZero
  MatchTextBox.text=""
  BallTextBox.text=FormatNumber(BallInPlay,0)
  for i = 1 to 2
    Score(i)=0
    Score100K(1)=0
    HighScorePaid(i)=false
    Replay1Paid(i)=false
    Replay2Paid(i)=false
    Replay3Paid(i)=false
    Replay4Paid(i)=false
    ReplayBalls1Paid(i)=false
    ReplayBalls2Paid(i)=false
    ReplayBalls3Paid(i)=false

  next
  If B2SOn Then
    Controller.B2SSetTilt 0
    Controller.B2SSetGameOver 0
    Controller.B2SSetMatch 0
    Controller.B2SSetScorePlayer1 0
    Controller.B2SSetScorePlayer2 0
'   Controller.B2SSetScorePlayer3 0
'   Controller.B2SSetScorePlayer4 0
    Controller.B2SSetBallInPlay BallInPlay
  End if

  If BallsPerGame=3 then
    BumpersOn
  else
    BumpersOff
  end if
  AltRelayValue=1
  BaseHit=0
  BaseHitCounter=0
  LightBase1.state=0
  LightBase2.state=0
  LightBase3.state=0
  LightBaseHome.state=1
  LightTop1.state=1
  LightTop2.state=1
  LightTop3.state=1
  TargetLightBase21.state=1
  TargetLightBase22.state=1
  TargetLightBase23.state=1
  TargetLightBase24.state=1
  TargetLightBase1.state=1
  TargetLightBase11.state=1
  TargetLightBase3.state=1
  TargetLightBase31.state=1
  BaseRunner1.SetValue(0)
  BaseRunner2.SetValue(0)
  BaseRunner3.SetValue(0)
  If B2SOn then
    Controller.B2SSetData 51,0
    Controller.B2SSetData 52,0
    Controller.B2SSetData 53,0
  end if
  ResetDrops
  ResetBalls
End sub

sub nextball
' If B2SOn Then
'   Controller.B2SSetTilt 0
' End If
  Player=Player+1
  If Player>Players Then
    BallInPlay=BallInPlay+1
    If BallInPlay>BallsPerGame then
      PlaySound("MotorLeer")
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
      For each obj in PlayerHUDScores
        obj.state=0
      next
      If Table1.ShowDT = True then
        For each obj in PlayerScores
          obj.visible=1
        Next
        For each obj in PlayerScoresOn
          obj.visible=0
        Next
      end If
      InstructCard.image="IC"+FormatNumber(BallsPerGame,0)+"0"


      CanPlayReel.SetValue(0)
      LeftFlipper.RotateToStart
      RightFlipper.RotateToStart
      LightsOut
      PlasticsOff
      BallTextBox.text=""
      if TableTilted=false then
        checkmatch
      end if
      CheckHighScore
      Players=0
      HighScoreTimer.interval=100
      HighScoreTimer.enabled=True
    Else
      Player=1
      If B2SOn Then
        Controller.B2SSetPlayerUp Player
        Controller.B2SSetBallInPlay BallInPlay

      End If
'     PlaySound("RotateThruPlayers")
      TempPlayerUp=Player
      PlayerUpRotator.enabled=true
      PlayStartBall.enabled=true
      For each obj in PlayerHuds
        obj.SetValue(0)
      next
      For each obj in PlayerHUDScores
        obj.state=0
      next
      If Table1.ShowDT = True then
        For each obj in PlayerScores
          obj.visible=1
        Next
        For each obj in PlayerScoresOn
          obj.visible=0
        Next
        PlayerHuds(Player-1).SetValue(1)
        PlayerHUDScores(Player-1).state=1
        PlayerScores(Player-1).visible=0
        PlayerScoresOn(Player-1).visible=1
        PlayerHuds(Player).SetValue(1)
        PlayerHUDScores(Player).state=1
        PlayerScores(Player).visible=0
        PlayerScoresOn(Player).visible=1
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
    PlayStartBall.enabled=true
    For each obj in PlayerHuds
      obj.SetValue(0)
    next
    If Table1.ShowDT = True then
      For each obj in PlayerScores
          obj.visible=1
      Next
      For each obj in PlayerScoresOn
          obj.visible=0
      Next
      PlayerHuds(Player-1).SetValue(1)
      PlayerHUDScores(Player-1).state=1
      PlayerScores(Player-1).visible=0
      PlayerScoresOn(Player-1).visible=1
      PlayerHuds(Player).SetValue(1)
      PlayerHUDScores(Player).state=1
      PlayerScores(Player).visible=0
      PlayerScoresOn(Player).visible=1
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
    sortpoints(i)=0
    sortplayerpoints(i)=0
  next
  playertops=0
  for i = 1 to Players
    sortscores(i)=Score(i)
    sortplayers(i)=i
    sortplayerpoints(i)=i
    sortpoints(i)=Score(i+1)
  next

  ScoreChecker=5
  CheckAllScores=1
  CheckAllPoints=0
  NewHighScore sortscores(ScoreChecker-1),sortplayers(ScoreChecker-1)
  savehs
end sub

sub checkmatch
  Dim tempmatch
  tempmatch=Int(Rnd*10)
  Match=tempmatch
  MatchReel.SetValue(tempmatch+1)
  if Match=0 then
    MatchTextBox.text="Match: 0"
  else
    MatchTextBox.text="Match: " + FormatNumber(Match,0)
  end if

  If B2SOn Then
    If Match = 0 Then
      Controller.B2SSetMatch 10
    Else
      Controller.B2SSetMatch Match
    End If
  End if
  for i = 1 to Players
    if Match=(Score(i) mod 10) then
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
      PlasticsOff
      BumpersOff
            DOF 101, DOFOff
            DOF 102, DOFOff
      LeftFlipper.RotateToStart
      RightFlipper.RotateToStart
      TiltReel.SetValue(1)
      If B2Son then
        Controller.B2SSetTilt 1
      end if
      BallInPlay=BallsPerGame
    else
      TiltTimer.Interval = 500
      TiltTimer.Enabled = True
    end if

end sub

Sub IncreaseBonus()


  If BonusCounter=10 then

  else
    If BonusCounter>0 then
      Bonus(BonusCounter).state=0
    end if
    BonusCounter=BonusCounter+1
    Bonus(BonusCounter).State=1
'   PlaySound("Score100")
  end if
  if BonusCounter=10 then
    TopLeftTargetLight.state=1
    TopRightTargetLight.state=1
  end if
End Sub


Sub BonusBoost_Timer()
  IncreaseBonus
  BonusBoosterCounter=BonusBoosterCounter-1
  If BonusBoosterCounter=0 then
    BonusBoost.enabled=false
  end if

end sub

Sub CheckForLightSpecial()

  if (TopLightA.state=0) and (TopLightB.state=0) and (TopLightC.state=0) then
    TopRightTargetLight.State=1
    TopLeftTargetLight.State=1
  end if
end sub

Sub PlayStartBall_timer()

  PlayStartBall.enabled=false
' PlaySound("StartBall2-5")
'   playsound SoundFXDOF("StartBall2-5",122,DOFPulse,DOFContactors)
end sub

Sub PlayerUpRotator_timer()
    Exit Sub
    If RotatorTemp<5 then
      TempPlayerUp=TempPlayerUp+1
      If TempPlayerUp>4 then
        TempPlayerUp=1
      end if
      For each obj in PlayerHuds
        obj.SetValue(0)
      next
      For each obj in PlayerHUDScores
        obj.state=0
      next
      If Table1.ShowDT = True then
        For each obj in PlayerScores
          obj.visible=1
        Next
        For each obj in PlayerScoresOn
          obj.visible=0
        Next
        PlayerHuds(TempPlayerUp-1).SetValue(1)

        PlayerScores(TempPlayerUp-1).visible=0
        PlayerScoresOn(TempPlayerUp-1).visible=1
      end If
      If B2SOn Then
        Controller.B2SSetPlayerUp TempPlayerUp
        Controller.B2SSetData 81,0
        Controller.B2SSetData 82,0
        Controller.B2SSetData 83,0
        Controller.B2SSetData 84,0
        Controller.B2SSetData 80+TempPlayerUp,1
      End If

    else
      if B2SOn then
        Controller.B2SSetPlayerUp Player
        Controller.B2SSetData 81,0
        Controller.B2SSetData 82,0
        Controller.B2SSetData 83,0
        Controller.B2SSetData 84,0
        Controller.B2SSetData 80+Player,1
      end if
      PlayerUpRotator.enabled=false
      RotatorTemp=1
      For each obj in PlayerHuds
        obj.SetValue(0)
      next
      For each obj in PlayerHUDScores
        obj.state=0
      next
      If Table1.ShowDT = True then
        For each obj in PlayerScores
          obj.visible=1
        Next
        For each obj in PlayerScoresOn
          obj.visible=0
        Next
        PlayerHuds(Player-1).SetValue(1)

        PlayerScores(Player-1).visible=0
        PlayerScoresOn(Player-1).visible=1
      end If
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
    ScoreFile.WriteLine ReplayLevel
    ScoreFile.WriteLine ReplayBalls

    for xx=1 to 5
      scorefile.writeline HSScore(xx)
    next
    for xx=1 to 5
      scorefile.writeline HSName(xx)
    next
    for xx=6 to 10
      scorefile.writeline HSScore(xx)
    next
    for xx=6 to 10
      scorefile.writeline HSName(xx)
    next
    ScoreFile.Close
  Set ScoreFile=Nothing
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
  dim temp19
  dim temp20
  dim temp21
  dim temp22
  dim temp23
  dim temp24
  dim temp25
  dim temp26
  dim temp27

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
    HighScore=cdbl(temp1)
    if HighScore<1 then

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
    end if
    TextStr.Close
      Credits=cdbl(temp2)
    BallsPerGame=cdbl(temp3)
    ReplayLevel=cdbl(temp4)
    ReplayBalls=cdbl(temp5)
    if HighScore<1 then
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

    end if
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

Sub DisplayHighScore

end sub


sub InitPauser5_timer
    If B2SOn Then
      'Controller.B2SSetScore 4,HighScore
    End If
    DisplayHighScore
    InitPauser5.enabled=false
end sub

sub ResetDropsTimer_timer
    ResetDropsTimer.enabled=0
    ResetDrops
end sub

sub BumpersOff
  If BallsPerGame=3 then
    Exit Sub
  end if

  Bumper1Light.state=0


  Bumper2Light.state=0


end sub

sub BumpersOn
  If BallsPerGame=3 then
    Bumper1Light.state=1
    Bumper2Light.state=1
    exit sub
  end if
  If LightBase2.state=1 AND LightBase3.state=1 then
    Bumper1Light.state=1
    Bumper2Light.state=1
  else
    Bumper1Light.state=0
    Bumper2Light.state=0
  end if

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
  Replay1Table(1)=5500
' Replay1Table(1)=3900
  Replay1Table(2)=4000
  Replay1Table(3)=4100
  Replay1Table(4)=4200
  Replay1Table(5)=4300
  Replay1Table(6)=4400
  Replay1Table(7)=4500
  Replay1Table(8)=4700
  Replay1Table(9)=4800
  Replay1Table(10)=4900
  Replay1Table(11)=5000
  Replay1Table(12)=5100
  Replay1Table(13)=5200
  Replay1Table(14)=5300
  Replay1Table(15)=999000

  Replay2Table(1)=6100
' Replay2Table(1)=5300
  Replay2Table(2)=5400
  Replay2Table(3)=5500
  Replay2Table(4)=5600
  Replay2Table(5)=5700
  Replay2Table(6)=5800
  Replay2Table(7)=5900
  Replay2Table(8)=6100
  Replay2Table(9)=6200
  Replay2Table(10)=6300
  Replay2Table(11)=6400
  Replay2Table(12)=6500
  Replay2Table(13)=6600
  Replay2Table(14)=6700
  Replay2Table(15)=999000

  Replay3Table(1)=6700
' Replay3Table(1)=6100
  Replay3Table(2)=6200
  Replay3Table(3)=6300
  Replay3Table(4)=6400
  Replay3Table(5)=6500
  Replay3Table(6)=6600
  Replay3Table(7)=6700
  Replay3Table(8)=6900
  Replay3Table(9)=7000
  Replay3Table(10)=7100
  Replay3Table(11)=7200
  Replay3Table(12)=7300
  Replay3Table(13)=7400
  Replay3Table(14)=7500
  Replay3Table(15)=999000

  Replay4Table(1)=7300
' Replay4Table(1)=999000
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

  ReplayTableMax=14

  ReplayBallsTable1(1)=20
' ReplayBallsTable1(1)=10
  ReplayBallsTable1(2)=10
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

  ReplayBallsTable2(1)=25
' ReplayBallsTable2(1)=14
  ReplayBallsTable2(2)=15
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

  ReplayBallsTable3(1)=30
' ReplayBallsTable3(1)=17
  ReplayBallsTable3(2)=17
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

  ReplayBallsTableMax=10

end sub

Sub RefreshReplayCard
  Dim tempst1
  Dim tempst2
  Dim tempst3

  tempst1=FormatNumber(BallsPerGame,0)
  tempst2=FormatNumber(ReplayLevel,0)
  tempst3=FormatNumber(ReplayBalls,0)

  ReplayCard.image = "SC" + tempst2
  ReplayCard1.image = "BC" + tempst3

  Replay1=Replay1Table(ReplayLevel)
  Replay2=Replay2Table(ReplayLevel)
  Replay3=Replay3Table(ReplayLevel)
  Replay4=Replay4Table(ReplayLevel)
  BallReplay1=ReplayBallsTable1(ReplayBalls)
  BallReplay2=ReplayBallsTable2(ReplayBalls)
  BallReplay3=ReplayBallsTable3(ReplayBalls)
end sub

'****************************************
'  SCORE MOTOR
'****************************************

Sub AddRunToScore
  Score(Player+1) = Score(Player+1) + 1
  If B2SOn Then
    Controller.B2SSetScorePlayer Player+1, Score(Player+1)
  End If
  PlayerScores(Player).AddValue(1)
  PlayerScoresOn(Player).AddValue(1)
  PlaySound SoundFXDOF("SpinACard_100_Point_Bell",142,DOFPulse,DOFChimes)
  If Score(Player+1) = BallReplay1 AND ReplayBalls1Paid(Player)=false THEN
    AddSpecial
    ReplayBalls1Paid(Player)=true
  end if
  If Score(Player+1) = BallReplay2 AND ReplayBalls2Paid(Player)=false THEN
    AddSpecial
    ReplayBalls2Paid(Player)=true
  end if
  If Score(Player+1) = BallReplay3 AND ReplayBalls3Paid(Player)=false THEN
    AddSpecial
    ReplayBalls3Paid(Player)=true
  end if

end sub

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
    if queuedscore>=400 then
      tempscore=400
      queuedscore=queuedscore-400
      SetMotor2(400)
      exit sub
    end if
    if queuedscore>=300 then
      tempscore=300
      queuedscore=queuedscore-300
      SetMotor2(300)
      exit sub
    end if
    if queuedscore>=200 then
      tempscore=200
      queuedscore=queuedscore-200
      SetMotor2(200)
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

  End If


end Sub

Sub ScoreMotorTimer_Timer
  If MotorPosition > 0 Then
    Select Case MotorPosition
      Case 5,4,3,2:
        If BaseHitCounter>0 then
          MoveRunners
          BaseHitCounter=BaseHitCounter-1
        end if
        If MotorMode=1000 Then
          AddScore(1000)
        End if
        if MotorMode=100 then
          AddScore(100)
        End If
        if MotorMode=10 then
          AddScore(10)
        End if
        MotorPosition=MotorPosition-1
      Case 1:
        If BaseHitCounter>0 then
          MoveRunners
          BaseHitCounter=BaseHitCounter-1
        end if
        Select Case BaseHit
          Case 4:
            AddRunToScore
            BumpersOff
          Case 3:
            TargetLightBase3.state=0
            TargetLightBase31.state=0
            LightBase3.state=1
            LightTop1.state=0
            If B2SOn then
              Controller.B2SSetData 53,1
            end if
            BaseRunner3.SetValue(1)
            If LightBase3.state=1 AND LightBase2.state=1 then
              BumpersOn
            else
              BumpersOff
            end if
          Case 2:
            LightBase2.state=1
            LightTop3.state=0
            TargetLightBase21.state=0
            TargetLightBase22.state=0
            TargetLightBase23.state=0
            TargetLightBase24.state=0
            LightTop3.state=0
            If B2SOn then
              Controller.B2SSetData 52,1
            end if
            BaseRunner2.SetValue(1)
            If LightBase3.state=1 AND LightBase2.state=1 then
              BumpersOn
            else
              BumpersOff
            end if
          Case 1:
            TargetLightBase1.state=0
            TargetLightBase11.state=0
            LightBase1.state=1
            LightTop2.state=0
            If B2SOn then
              Controller.B2SSetData 51,1
            end if
            BaseRunner1.SetValue(1)
            If LightBase3.state=1 AND LightBase2.state=1 then
              BumpersOn
            else
              BumpersOff
            end if
        end select
        If MotorMode=1000 Then
          AddScore(1000)
        end if
        if MotorMode=100 then
          AddScore(100)
        End If
        if MotorMode=10 then
          AddScore(10)

        End if
        MotorPosition=0:MotorRunning=0:LightsOn: BaseHit=0: BaseHitCounter=0
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
'
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
  End Select
  NewScore = Score(Player)

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

    end if

    OldScore = int(OldScore / 10)
    NewScore = int(NewScore / 10)
  ' MsgBox("OldScore="&OldScore&", NewScore="&NewScore)
    if (OldScore Mod 10 <> NewScore Mod 10) then
    PlayChime(100)
    FireAlternatingRelay

    end if

    OldScore = int(OldScore / 10)
    NewScore = int(NewScore / 10)
  ' MsgBox("OldScore="&OldScore&", NewScore="&NewScore)
    if (OldScore Mod 10 <> NewScore Mod 10) then
    PlayChime(1000)


    end if

    OldScore = int(OldScore / 10)
    NewScore = int(NewScore / 10)
  ' MsgBox("OldScore="&OldScore&", NewScore="&NewScore)
    if (OldScore Mod 10 <> NewScore Mod 10) then
    PlayChime(1000)
    end if

  If B2SOn Then
    Controller.B2SSetScorePlayer Player, Score(Player)
  End If
' EMReel1.SetValue Score(Player)
  PlayerScores(Player-1).AddValue(x)
  PlayerScoresOn(Player-1).AddValue(x)
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

Sub RollingTimer_Timer()
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
      If BallVel(BOT(b) ) > 1 Then
        rolling(b) = True
        if BOT(b).z < 30 Then ' Ball on playfield
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), AudioPan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
				  ' PlaySound("fx_ballrolling" & b), -1, BallRollVol(BOT(b) )*.8, AudioPan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        Else ' Ball on raised ramp
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*.5, AudioPan(BOT(b) ), 0, Pitch(BOT(b) )+50000, 1, 0, AudioFade(BOT(b) )
				  ' PlaySound("fx_ballrolling" & b), -1, BallRollVol(BOT(b) )*.2, AudioPan(BOT(b) ), 0, Pitch(BOT(b) )+50000, 1, 0, AudioFade(BOT(b) )
        End If
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


Dim HSScore(10)       ' High Scores read in from config file
Dim HSName(10)        ' High Score Initials read in from config file

' default high scores, remove this when the scores are available from the config file
HSScore(1) = 7500
HSScore(2) = 7000
HSScore(3) = 6000
HSScore(4) = 5500
HSScore(5) = 5000

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

Sub HighScoreTimer_Timer

  if EnteringInitials then
    if HSTimerCount = 1 then
      If CheckAllScores=1 then
        SetHSLine 3, InitialString & MID(AlphaString, AlphaStringPos, 1)
      elseIf CheckAllPoints=1 then
        SetHSLine 6, InitialString & MID(AlphaString, AlphaStringPos, 1)
      end if
      HSTimerCount = 2
    else
      If CheckAllScores=1 then
        SetHSLine 3, InitialString
      elseIf CheckAllPoints=1 then
        SetHSLine 6, InitialString
      end if
      HSTimerCount = 1
    end if
    Exit sub
  elseif InProgress then
    SetHSLine 1, "HIGH SCORE1"
    SetHSLine 2, HSScore(1)
    SetHSLine 3, HSName(1)
    SetHSLine 4, "CARRERAS1"
    SetHSLine 5, HSScore(6)
    SetHSLine 6, HSName(6)
    HSTimerCount = 5  ' set so the highest score will show after the game is over
    HighScoreTimer.enabled=false
  elseif CheckAllScores then
    NewHighScore sortscores(ScoreChecker-1),sortplayers(ScoreChecker-1)
  elseif CheckAllPoints then
    NewHighPoints sortpoints(ScoreChecker-1),sortplayerpoints(ScoreChecker-1)

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
    SetHSLine 4, "CARRERAS "+FormatNumber(HSTimerCount,0)
    SetHSLine 5, HSScore(HSTimerCount+5)
    SetHSLine 6, HSName(HSTimerCount+5)
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
  StartHSArray=array(0,1,12,22,32,43,53)
  EndHSArray=array(0,11,21,31,42,52,62)
  StrLen = len(string)
  Index = 1

  for xfor = StartHSArray(LineNo) to EndHSArray(LineNo)
    Eval("HS"&xfor).image = GetHSChar(String, Index)
    Index = Index + 1
  next

End Sub

Sub NewHighScore(NewScore, PlayNum)
  ScoreChecker=ScoreChecker-1
  if ScoreChecker=0 then
    CheckAllScores=0
    ScoreChecker=5
    CheckAllPoints=1
    exit sub
  end if
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

End Sub

Sub NewHighPoints(NewScore, PlayNum)
  ScoreChecker=ScoreChecker-1
  if ScoreChecker=0 then
    CheckAllScores=0
    CheckAllPoints=0
    Exit sub
  end if
  if NewScore > HSScore(10) then
    HighScoreTimer.interval = 500
    HSTimerCount = 1
    AlphaStringPos = 1    ' start with first character "A"
    EnteringInitials = 1  ' intercept the control keys while entering initials
    InitialString = ""    ' initials entered so far, initialize to empty
    SetHSLine 4, "PLAYER "+FormatNumber(PlayNum,0)
    SetHSLine 5, "ENTER NAME"
    SetHSLine 6, MID(AlphaString, AlphaStringPos, 1)
    HSNewHigh = NewScore
    For xx=1 to HighScoreReward
      AddSpecial
    next
  End if

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
    If CheckAllScores=1 then
      SetHSLine 3, InitialString & MID(AlphaString, AlphaStringPos, 1)
    end if
    If CheckAllPoints=1 then
      SetHSLine 6, InitialString & MID(AlphaString, AlphaStringPos, 1)
    end if

    PlaySound "DropTargetDropped"
  elseif keycode = RightFlipperKey Then
    ' advance to next character
    AlphaStringPos = AlphaStringPos + 1
    if AlphaStringPos > len(AlphaString) or (AlphaStringPos = len(AlphaString) and InitialString = "") then
      ' Skip the backspace if there are no characters to backspace over
      AlphaStringPos = 1
    end if
    If CheckAllScores=1 then
      SetHSLine 3, InitialString & MID(AlphaString, AlphaStringPos, 1)
    end if
    If CheckAllPoints=1 then
      SetHSLine 6, InitialString & MID(AlphaString, AlphaStringPos, 1)
    end if
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
      If CheckAllScores=1 then
        SetHSLine 3, InitialString & SelectedChar
      end if
      If CheckAllPoints=1 then
        SetHSLine 6, InitialString & SelectedChar
      end if
    End If
  End If
  if len(InitialString) >= 3 then
    ' save the score
    If CheckAllScores=1 then
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
          HighScoreTimer.enabled=false
          HighScoreTimer.interval = 100
          HighScoreTimer.enabled=true
          SetHSLine 2, HSScore(i)
          SetHSLine 3, HSName(i)
          PlaySound("Ding1000")
          exit sub
        elseif i < 5 then
          ' move the score in this slot down by 1, it's been exceeded by the new score
  ' MsgBox("Moving " & i & " to " & (i + 1))
          HSScore(i + 1) = HSScore(i)
          HSName(i + 1) = HSName(i)
        end if
      next
    end if
    If CheckAllPoints=1 then
      for i = 10 to 6 step -1
        if i = 6 or (HSNewHigh > HSScore(i) and HSNewHigh <= HSScore(i - 1)) then
          ' Replace the score at this location
          if i < 10 then
  ' MsgBox("Moving " & i & " to " & (i + 1))
            HSScore(i + 1) = HSScore(i)
            HSName(i + 1) = HSName(i)
          end if
  ' MsgBox("Saving initials " & InitialString & " to position " & i)
          EnteringInitials = 0
          HSScore(i) = HSNewHigh
          HSName(i) = InitialString
          HSTimerCount = 5
          HighScoreTimer.enabled=false
          HighScoreTimer.interval = 100
          HighScoreTimer.enabled=true
          SetHSLine 5, HSScore(i)
          SetHSLine 6, HSName(i)
          PlaySound("Ding1000")
          exit sub
        elseif i < 10 then
          ' move the score in this slot down by 1, it's been exceeded by the new score
  ' MsgBox("Moving " & i & " to " & (i + 1))
          HSScore(i + 1) = HSScore(i)
          HSName(i + 1) = HSName(i)
        end if
      next
    end if
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
Const OptionLinesToMark="111100011"
Const OptionLine1="" 'do not use this line
Const OptionLine2="" 'do not use this line
Const OptionLine3="" 'do not use this line
Const OptionLine4="Points Scored Replay Settings"
Const OptionLine5=""
Const OptionLine6=""
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
  dim TempText1
  dim TempText2
  dim TempText3
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
        if ReplayBallsTable1(ReplayBalls)>=99 then
          TempText1 = ""
        else
          TempText1 = FormatNumber(ReplayBallsTable1(ReplayBalls),0)
        end if

        if ReplayBallsTable2(ReplayBalls)>=99 then
          TempText2 = ""
        else
          TempText2 = "/" + FormatNumber(ReplayBallsTable2(ReplayBalls),0)
        end if
        if ReplayBallsTable3(ReplayBalls)>=99 then
          TempText3 = ""
        else
          TempText3 = "/" + FormatNumber(ReplayBallsTable3(ReplayBalls),0)
        end if

        tempstring = TempText1 + TempText2 + TempText3

        SetOptLine 5, tempstring
      Case 5:
        SetOptLine 6, tempstring

        SetOptLine 7, tempstring

      Case 6:
        SetOptLine 8, tempstring

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
    elseif CurrentOption = 4 then
      ReplayBalls=ReplayBalls+1
      If ReplayBalls>ReplayBallsTableMax then
        ReplayBalls=1
      end if
      DisplayAllOptions
    elseif CurrentOption = 8 or CurrentOption = 9 then
        if OptionCHS=1 then
          HSScore(1) = 7500
          HSScore(2) = 7000
          HSScore(3) = 6000
          HSScore(4) = 5500
          HSScore(5) = 5000

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
        InstructCard.image="IC"+FormatNumber(BallsPerGame,0)+"0"
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

'FOZZY

Sub ShuffleArray(ByRef aArray, byVal offset) 'shuffle 1d array
  dim x, aCount : aCount = 0
  redim a(uBound(aArray) )
  for x = 0 to uBound(aArray) 'Shuffle objects in a temp array
    if not IsEmpty(aArray(x) ) Then
      if IsObject(aArray(x)) then
        Set a(aCount) = aArray(x)
      Else
        a(aCount) = aArray(x)
      End If
      aCount = aCount + 1
    End If
  Next
  if offset < 0 then offset = 0
  redim aArray(aCount-1+offset) 'Resize original array
  for x = 0 to aCount-1 'set objects back into original array
    if IsObject(a(x)) then
      Set aArray(x) = a(x)
    Else
      aArray(x) = a(x)
    End If
  Next
End Sub

' Used for flipper correction and rubber dampeners
Sub ShuffleArrays(aArray1, aArray2, offset)
  ShuffleArray aArray1, offset
  ShuffleArray aArray2, offset
End Sub

' Used for flipper correction, rubber dampeners, and drop targets
Function BallSpeed(ball) 'Calculates the ball speed
  BallSpeed = SQR(ball.VelX^2 + ball.VelY^2 + ball.VelZ^2)
End Function

' Used for flipper correction and rubber dampeners
Function PSlope(Input, X1, Y1, X2, Y2) 'Set up line via two points, no clamping. Input X, output Y
  dim x, y, b, m : x = input : m = (Y2 - Y1) / (X2 - X1) : b = Y2 - m*X2
  Y = M*x+b
  PSlope = Y
End Function

'******************************************************
' TRACK ALL BALL VELOCITIES
' FOR RUBBER DAMPENER AND DROP TARGETS
'******************************************************

dim cor : set cor = New CoRTracker

Class CoRTracker
  public ballvel, ballvelx, ballvely
  Private Sub Class_Initialize : redim ballvel(0) : redim ballvelx(0): redim ballvely(0) : End Sub

  Public Sub Update() 'tracks in-ball-velocity
    dim str, b, AllBalls, highestID : allBalls = getballs

    for each b in allballs
      if b.id >= HighestID then highestID = b.id

    Next

    if uBound(ballvel) < highestID then redim ballvel(highestID) 'set bounds
    if uBound(ballvelx) < highestID then redim ballvelx(highestID) 'set bounds
    if uBound(ballvely) < highestID then redim ballvely(highestID) 'set bounds

    for each b in allballs
      ballvel(b.id) = BallSpeed(b)
      ballvelx(b.id) = b.velx
      ballvely(b.id) = b.vely
    Next
  End Sub
End Class

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

Sub RDampen_Timer()
Cor.Update
End Sub

'*************************************************************
'PHYSICS DAMPENERS
'*************************************************************

'These are data mined bounce curves,
'dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
'Requires tracking ballspeed to calculate COR


Sub dPosts_Hit(idx)
    RubbersD.dampen Activeball
End Sub

Sub dSleeves_Hit(idx)
  SleevesD.Dampen Activeball
End Sub


dim RubbersD : Set RubbersD = new Dampener 'frubber
RubbersD.name = "Rubbers"
RubbersD.debugOn = False 'shows info in textbox "TBPout"
RubbersD.Print = False 'debug, reports in debugger (in vel, out cor)
'cor bounce curve (linear)
'for best results, try to match in-game velocity as closely as possible to the desired curve
'RubbersD.addpoint 0, 0, 0.935 'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 0, 0, 0.96 'point# (keep sequential),ballspeed, CoR (elasticity)
RubbersD.addpoint 1, 3.77, 0.96
RubbersD.addpoint 2, 5.76, 0.967 'dont take this as gospel. if you can data mine rubber elasticitiy, please help!
RubbersD.addpoint 3, 15.84, 0.874
RubbersD.addpoint 4, 56, 0.64 'there's clamping so interpolate up to 56 at least

dim SleevesD : Set SleevesD = new Dampener 'this is just rubber but cut down to 85%...
SleevesD.name = "Sleeves"
SleevesD.debugOn = False 'shows info in textbox "TBPout"
SleevesD.Print = False 'debug, reports in debugger (invel, out cor)
SleevesD.CopyCoef RubbersD, 0.85

Class Dampener
  Public Print, debugOn 'tbpOut.text
  public name, Threshold 'Minimum threshold. Useful for Flippers, which don't have a hit threshold.
  Public ModIn, ModOut
  Private Sub Class_Initialize : redim ModIn(0) : redim Modout(0): End Sub

  Public Sub AddPoint(aIdx, aX, aY)
    ShuffleArrays ModIn, ModOut, 1 : ModIn(aIDX) = aX : ModOut(aIDX) = aY : ShuffleArrays ModIn, ModOut, 0
    if gametime > 100 then Report
  End Sub

  public sub Dampen(aBall)
    if threshold then if BallSpeed(aBall) < threshold then exit sub end if end if
    dim RealCOR, DesiredCOR, str, coef
    DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
    RealCOR = BallSpeed(aBall) / cor.ballvel(aBall.id)
    coef = desiredcor / realcor
    if debugOn then str = name & " in vel:" & round(cor.ballvel(aBall.id),2 ) & vbnewline & "desired cor: " & round(desiredcor,4) & vbnewline & _
    "actual cor: " & round(realCOR,4) & vbnewline & "ballspeed coef: " & round(coef, 3) & vbnewline
    if Print then debug.print Round(cor.ballvel(aBall.id),2) & ", " & round(desiredcor,3)

    aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
    if debugOn then TBPout.text = str
  End Sub

  Public Sub CopyCoef(aObj, aCoef) 'alternative addpoints, copy with coef
    dim x : for x = 0 to uBound(aObj.ModIn)
      addpoint x, aObj.ModIn(x), aObj.ModOut(x)*aCoef
    Next
  End Sub

  Public Sub Report() 'debug, reports all coordsin tbPL.text
    if not debugOn then exit sub
    dim a1, a2 : a1 = ModIn : a2 = ModOut
    dim str, x : for x = 0 to uBound(a1) : str = str & x & ": " & round(a1(x),4) & ", " & round(a2(x),4) & vbnewline : next
    TBPout.text = str
  End Sub


End Class

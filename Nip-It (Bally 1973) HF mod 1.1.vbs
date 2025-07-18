'*
'*        Bally's Nip-It (1973)
'*        Table primary build/scripted by Loserman!
'*        Playfield/Plastics/lighting/shadows/primitive holes by hauntfreaks
'*        http://www.ipdb.org/machine.cgi?id=1680
'*
'*
'*Disclaimer:
'*Please do not redistribute, repost, repurpose or use the content for public or commercial purposes, images, or software code is prohibited without prior permission from author
'*
'*Reason for disclaimer:
'*do to the total disrespect of the VP builders (tables, b2s, frontend , backend, etc.) and its community, certon talentless, low self esteem, trying to compensate individuals have taken it upon themselves to redistribute our work as if it is there god given right to do so....
'*
'*
'*
'*

option explicit
Randomize
ExecuteGlobal GetTextFile("core.vbs")

On Error Resume Next
ExecuteGlobal GetTextFile("Controller.vbs")
If Err Then MsgBox "Unable to open Controller.vbs. Ensure that it is in the scripts folder."
On Error Goto 0

Const cGameName = "NipIt_1973"

Const ShadowFlippersOn = true
Const ShadowBallOn = true

Const ShadowConfigFile = false




Dim Controller  ' B2S
Dim B2SScore  ' B2S Score Displayed
Const HSFileName="NipIt_73VPX.txt"
Const B2STableName="NipIt_1973"
Const LMEMTableConfig="LMEMTables.txt"
Const LMEMShadowConfig="LMEMShadows.txt"
Dim EnableBallShadow
Dim EnableFlipperShadow

'* this value adjusts whether or not a locked ball remains locked when game is over (for use in next game) or kicked out
'* 0=ball remains locked at game over, 1=ball is kicked out at game over
Const EndGameClearGatorHole=0


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
Dim TextStr,TextStr2
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

Dim bgpos,dooralreadyopen

Dim OperatorMenu

Dim FlipperPosition
Dim LeftFlippersFlipped,RightFlippersFlipped

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

dim ChimesOn

dim raceRflag
dim raceAflag
dim raceCflag
dim raceEflag

Dim LStep, LStep2, RStep, xx,yy

Dim TStep

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

Dim TempLightTracker

Dim TargetLeftFlag
Dim TargetCenterFlag
Dim TargetRightFlag

Dim TargetSequenceComplete

Dim SpecialLightsFlag

Dim AlternatingRelay

Dim ZeroToNineUnit

Dim Kicker1Hold,Kicker2Hold,Kicker3Hold

Dim mHole,mHole2, mHole3

Dim SpinPos,Spin,Count,Count1,Count2,Count3,Reset,VelX,VelY,BallSpeed,LitSpinner

Spin=Array("5","6","7","8","9","10","J","Q","K","A")

Dim SilverBonusCounter

Dim ScoreMotorStepper

Dim FirecrackerLighter

Dim GatorMultiBall
Dim GatorSpot
Dim GatorRunning
Dim GatorPositionCounter

Sub Table1_Init()
  LoadEM
  LoadLMEMConfig2

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

  Mush1.Visible=1
  Mush3.Visible = 0:Mush2.Visible = 0'
  Count=0
    Count1=0
    Count2=0
  Count3=0
    Reset=0
  ZeroToNineUnit=Int(Rnd*10)
  Kicker1Hold=0
  Kicker2Hold=0
  Kicker3Hold=0

  EightLit=0
  ChimesOn=0
  BallCounter=0
  ReelCounter=0
  AddABall=0
  DropABall=0

  SetupReplayTables
  PlasticsOff
  BumpersOff
  OperatorMenu=0
  HighScore=0
  MotorRunning=0
  HighScoreReward=1
  Credits=0
  BallsPerGame=5
  ReplayLevel=1
  GatorMultiBall=False
  AlternatingRelay=1
  SpecialLightOption=2
  BackglassBallFlagColor=1
  loadhs
  if HighScore=0 then HighScore=50000


  TableTilted=false

  Match=int(Rnd*10)
  MatchReel.SetValue((Match)+1)
  Match=Match*10

  SpinPos=int(Rnd*10)

  CanPlayReel.SetValue(0)
  GameOverReel.SetValue(1)

  FlipperPosition=0
  LeftFlippersFlipped=0
  RightFlippersFlipped=0
  ShowFlippers

  For each obj in PlayerHuds
    obj.state=0
  next
  For each obj in PlayerHUDScores
    obj.state=0
  next

  For each obj in PlayerScores
    obj.ResetToZero
  next

  for each obj in BonusLights
    obj.state=0
  next

  for each obj in bottgate
    obj.isdropped=true
    next
  bgpos=6
  bottgate(bgpos).isdropped=false
  dooralreadyopen=0
  primgate.RotY=90


  Replay1=Replay1Table(ReplayLevel)
  Replay2=Replay2Table(ReplayLevel)
  Replay3=Replay3Table(ReplayLevel)
  Replay4=Replay4Table(ReplayLevel)


  BonusCounter=0
  HoleCounter=0


  InstructCard.image="IC"
  BallCard.image="BC"+FormatNumber(BallsPerGame,0)
  RefreshReplayCard

  CurrentFrame=0

  Bumper1Light.state=0
  Bumper2Light.state=0
  Bumper3Light.state=0
  Bumper4Light.state=0

  Kicker5.enabled=false
  Kicker6.enabled=false
  Kicker7.enabled=false

  AdvanceLightCounter=0

  Players=0
  RotatorTemp=1
  InProgress=false
  TargetLightsOn=false

  ScoreText.text=HighScore
  TiltReel.SetValue(1)

  If B2SOn Then

    if Match=0 then
      Controller.B2SSetMatch 100
    else
      Controller.B2SSetMatch Match
    end if
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
  If Credits > 0 Then DOF 129, DOFOn
End Sub

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
    LeftFlippers(FlipperPosition).RotateToEnd
    LeftFlippersFlipped=1
    PlaySoundAtVol SoundFXDOF("FlipperUp",101,DOFOn,DOFFlippers), LeftFlipper, 1
    PlayLoopSoundAtVol "buzzL", LeftFlipper, 1
  End If

  If keycode = RightFlipperKey  and InProgress=true and TableTilted=false Then
    RightFlippers(FlipperPosition).RotateToEnd
    RightFlippersFlipped=1
    PlaySoundAtVol SoundFXDOF("FlipperUp",102,DOFOn,DOFFlippers), RightFlipper, 1
    PlayLoopSoundAtVol "buzz", RightFlipper, 1
  End If

  if keycode = RightMagnaSave and InProgress=true and TableTilted=false then
    If GatorRunning=0 then
      GatorSpot=0
      GatorPositionCounter=15
      GatorWall.IsDropped=true
'     GatorWall1.IsDropped=false

      gatortimerout.enabled=true
      GatorRunning=1
      Playsound "Ballygator"
      Kicker5.enabled=true
      Kicker6.enabled=true
      Kicker7.enabled=true
    end if
  end if

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


   if keycode = StartGameKey and Credits>0 and InProgress=true and Players>0 and Players<4 and BallInPlay<2 then
    Credits=Credits-1
    If Credits < 1 Then DOF 129, DOFOff
    CreditsReel.SetValue(Credits)
    Players=Players+1
    CanPlayReel.SetValue(Players)
    playsound "BallyStartButtonPlayers2-4"
    If B2SOn Then
      Controller.B2SSetCanPlay Players
      If Players=2 Then
        Controller.B2SSetScoreRolloverPlayer2 0
      End If
      If Players=3 Then
        Controller.B2SSetScoreRolloverPlayer3 0
      End If
      If Players=4 Then
        Controller.B2SSetScoreRolloverPlayer4 0
      End If
      Controller.B2SSetCredits Credits
    End If
    end if

  if keycode=StartGameKey and Credits>0 and InProgress=false and Players=0 and GatorMultiBall=False and EnteringOptions = 0 then
'GNMOD
    OperatorMenuTimer.Enabled = false
'END GNMOD
    Credits=Credits-1
    If Credits < 1 Then DOF 129, DOFOff
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
    TimerTitle1.enabled=1
    TimerTitle2.enabled=1
    TimerTitle3.enabled=1
    TimerTitle4.enabled=1
    If B2SOn Then
      Controller.B2SSetTilt 0
      Controller.B2SSetGameOver 0
      Controller.B2SSetMatch 0
      Controller.B2SSetCredits Credits

      Controller.B2SSetCanPlay 1
      Controller.B2SSetPlayerUp 1
      Controller.B2SSetBallInPlay BallInPlay
      Controller.B2SSetScoreRolloverPlayer1 0

      Controller.B2SSetShootAgain 0

    End If
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
        obj.state=0
      next
      For each obj in PlayerHUDScores
        obj.state=0
      next
      PlayerHuds(Player-1).state=1
      PlayerHUDScores(Player-1).state=1
      PlayerScores(Player-1).Visible=0
      PlayerScoresOn(Player-1).Visible=1
    end If
    GameOverReel.SetValue(0)


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
    LeftFlippers(FlipperPosition).RotateToStart
    LeftFlippersFlipped=0
    StopSound "buzzL"
    PlaySoundAtVol SoundFXDOF("FlipperDown",101,DOFOff,DOFFlippers), LeftFlipper, 1
  End If

  If keycode = RightFlipperKey and InProgress=true and TableTilted=false Then
    RightFlippers(FlipperPosition).RotateToStart
    RightFlippersFlipped=0
    StopSound "buzz"
    PlaySoundAtVol SoundFXDOF("FlipperDown",102,DOFOff,DOFFlippers), ActiveBall, 1
  End If

End Sub



Sub Drain_Hit()

  Drain.DestroyBall
  PlaySoundAtVol "fx_drain", Drain, 1
  DOF 115, DOFPulse
  If GatorMultiBall=False then

    Pause4Bonustimer.enabled=true

  else
    GatorMultiBall=False
  end if
End Sub

Sub Trigger0_Hit()
  DOF 130, DOFPulse
End Sub

Sub Pause4Bonustimer_timer
  Pause4Bonustimer.enabled=0
  ScoreSilverBonus


End Sub

'***********************
'     Flipper Logos
'***********************

Sub UpdateFlipperLogos_Timer
  If FlipperPosition=0 then
    FlipperT2Left.ObjRotZ=LeftFlippers(FlipperPosition).CurrentAngle+90
    FlipperT2Right.ObjRotZ=RightFlippers(FlipperPosition).CurrentAngle+90
    FlipperLSh.RotZ = LeftFlippers(FlipperPosition).currentangle
    FlipperRSh.RotZ = RightFlippers(FlipperPosition).currentangle
  else
    FlipperT2Left1.ObjRotZ=LeftFlippers(FlipperPosition).CurrentAngle+90
    FlipperT2Right1.ObjRotZ=RightFlippers(FlipperPosition).CurrentAngle+90
    FlipperLSh1.RotZ = LeftFlippers(FlipperPosition).currentangle
    FlipperRSh1.RotZ = RightFlippers(FlipperPosition).currentangle
    end if
  PGate.Rotz = (Gate.CurrentAngle*.75) + 25

End Sub

'***********************
'     the Gator!
'***********************
Sub gatortimerout_timer
  Primitive5.TransZ = -100
  GatorSpot=-100
  gatortimerout.enabled=false
  Kicker5.enabled=true
  Kicker6.enabled=true
  Kicker7.enabled=true
  gatortimerin.enabled=true
end sub

Sub gatortimerin_timer
  GatorSpot=GatorSpot+10
  GatorPositionCounter=GatorPositionCounter-1
  if GatorSpot>0 then
    GatorSpot=0
  end if
  If GatorPositionCounter<1 then
    GatorWall.IsDropped=false
'   GatorWall1.IsDropped=true
    GatorRunning=0

'   Primitive5.Visible=false
    Kicker5.enabled=false
    Kicker6.enabled=false
    Kicker7.enabled=false
    gatortimerin.enabled=false

  end if
  Primitive5.TransZ = GatorSpot
end sub

Sub Kicker5_Hit()
  if gatortimerin.enabled=true then
    Kicker5.kick 100,15
  else
    Kicker5.kick 275,5
  end if
end sub

Sub Kicker6_Hit()
  if gatortimerin.enabled=true then
    Kicker6.kick 95,15
  else
    Kicker6.kick 275,5
  end if
end sub

Sub Kicker7_Hit()
  if gatortimerin.enabled=true then
    Kicker7.kick 90,15
  else
    Kicker7.kick 275,5
  end if
end sub



'***********************
' slingshots
'
Sub RightSlingShot_Slingshot

  PlaySoundAtVol SoundFXDOF("right_slingshot",104,DOFPulse,DOFContactors), ActiveBall, 1
  DOF 106, DOFPulse
    RSling0.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
  AddScore 10
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
  DOF 105, DOFPulse
    LSling0.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
  AddScore 10
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
    AddScore(10)
  end if
end Sub

'**********************Mushroom Trigger Animation
'****Init


Sub Mush1_hit()
  DOF 131, DOFPulse
  TStep = 0: Mush2.Visible = 1:Mush1.Visible = 0:PlaySoundAtVol "Rubber", ActiveBall, 1
  FlipperPosition=1
  ShowFlippers
  PlaySoundAtVol SoundFXDOF("ZipperFlipper",101,DOFPulse,DOFFlippers), ActiveBall, 1
  DOF 102,DOFPulse
  AddScore(100)
  MushTimer.Enabled = 1
End Sub

Sub MushTimer_Timer()
  Select Case TStep
      Case 0 : Mush3.visible = 1:Mush2.visible = 0
      Case 1 : Mush2.visible = 1:Mush3.visible = 0
      Case 2 : Mush1.visible = 1:Mush2.visible = 0:MushTimer.Enabled = 0
  End Select
  TStep = TStep +1
End Sub

'***********************************
' Bumpers
'***********************************


Sub Bumper1_Hit
  If TableTilted=false then

    PlaySoundAtVol SoundFXDOF("bumper1",107,DOFPulse,DOFContactors), ActiveBall, 1
    DOF 111, DOFPulse
    bump1 = 1
    If Bumper1Light.state=1 then
      AddScore(100)
    else
      AddScore(10)
    end if


    end if

End Sub

Sub Bumper2_Hit
  If TableTilted=false then

    PlaySoundAtVol SoundFXDOF("bumper1",110,DOFPulse,DOFContactors), ActiveBall, 1
    DOF 114, DOFPulse
    bump2 = 1
    If Bumper2Light.state=1 then
      AddScore(100)
    else
      AddScore(10)
    end if

    end if

End Sub

Sub Bumper3_Hit
  If TableTilted=false then

    PlaySoundAtVol SoundFXDOF("bumper1",109,DOFPulse,DOFContactors), ActiveBall, 1
    DOF 113, DOFPulse
    bump3 = 1
    If Bumper3Light.state=1 then
      AddScore(100)
    else
      AddScore(10)
    end if

    end if

End Sub

Sub Bumper4_Hit
  If TableTilted=false then

    PlaySoundAtVol SoundFXDOF("bumper1",108,DOFPulse,DOFContactors), ActiveBall, 1
    DOF 112, DOFPulse
    bump4 = 1
    If Bumper4Light.state=1 then
      AddScore(100)
    else
      AddScore(10)
    end if

    end if

End Sub


'************************************
'  Rollover lanes
'************************************

Sub TriggerCollection_Hit(idx)
  If TableTilted=false then
  DOF 117 + idx, DOFPulse
    Select Case idx
      Case 0:
        AddScore(1000)
        IncreaseSilverBonus
      Case 1:
        SetMotor(500)
      Case 2:
        SetMotor(500)
        If dooralreadyopen=1 then
          PauseGateCloser.enabled=true
        end if
    End Select
  end if

end Sub

'****************************************************



sub Kicker2_Hit()
  Dim speedx,speedy,finalspeed
  speedx=activeball.velx
  speedy=activeball.vely
  finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  if TableTilted=true then
    Kicker2.Timerenabled=1
    exit sub
  end if
  if finalspeed>20 then
    Kicker2.Kick 0,0
    activeball.velx=speedx
    activeball.vely=speedy
  else
    KickerHolder2.enabled=1
    SetMotor(500)
  end if

end sub

Sub KickerHolder2_timer
  If MotorRunning=0 then
    KickerHolder2.enabled=0
    Kicker2.timerenabled=1

  end if
end sub

Sub Kicker2_Timer()

  Kicker2.timerenabled=0
  PlaySoundAtVol SoundFXDOF("saucer",121,DOFPulse,DOFContactors), PkickArm2, 1
  DOF 124, DOFPulse
  Kicker2.kick 120,13
  Pkickarm2.rotz=15
  KickerArm2.enabled=true
end sub

sub KickerArm2_timer
  Pkickarm2.rotz=0
  KickerArm2.enabled=false
end sub

sub Kicker3_Hit()
  Dim speedx,speedy,finalspeed
  speedx=activeball.velx
  speedy=activeball.vely
  finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  if TableTilted=true then
    Kicker3.Timerenabled=1
    exit sub
  end if
  if finalspeed>20 then
    Kicker3.Kick 0,0
    activeball.velx=speedx
    activeball.vely=speedy
  else
' Kicker1.timerinterval=(100)
' Kicker1.timerenabled=1
    Kicker3Hold=1
    KickerHolder3.enabled=1
    TargetLight3.state=1
    TargetLight4.state=1
    If GatorMultiBall=False then
      'CreateBallID BallRelease
      Ballrelease.CreateSizedBall 25
      Ballrelease.Kick 40,7
      DOF 116, DOFPulse
    else
      GatorMultiBall=false
    end if
  end if

end sub

Sub KickerHolder3_timer
  If Kicker3Hold=0 and MotorRunning=0 then
    KickerHolder3.enabled=0
    Kicker3.timerenabled=1
    TargetLight3.state=0
    TargetLight4.state=0
  end if
end sub

Sub Kicker3_Timer()
  Kicker3.timerenabled=0
  PlaySoundAtVol SoundFXDOF("saucer",122,DOFPulse,DOFContactors), ActiveBall, 1
  DOF 124, DOFPulse
  Kicker3.kick 210,10
  Pkickarm3.rotz=15
  KickerArm3.enabled=true
end sub

sub KickerArm3_timer
  Pkickarm3.rotz=0
  KickerArm3.enabled=false
end sub

sub Kicker4_Hit()
  Kicker4.timerenabled=1

end sub

Sub Kicker4_Timer()

  Kicker4.timerenabled=0
  If TableTilted=false then
    AddScore(1000)
  end if
  PlaySoundAtVol SoundFXDOF("saucer",123,DOFPulse,DOFContactors), Kicker4, 1
  DOF 124, DOFPulse
  Kicker4.kick 0,30
end sub

'**************************************
' Stationary Targets
'**************************************

Sub Target1_Hit
  DOF 125, DOFPulse
  If TableTilted=false then
    AddScore(1000)
    IncreaseSilverBonus

  end if

end sub

Sub Target2_Hit
  DOF 125, DOFPulse
  If TableTilted=false then
    AddScore(100)
    If TargetLight2.state=1 then
      if dooralreadyopen=0 then
        openg.enabled=1
      end if
    end if
  end if

end sub

Sub Target3_Hit
  DOF 126, DOFPulse
  If TableTilted=false then
    AddScore(100)
    if TargetLight3.state=1 then
      GatorMultiBall=true
      Kicker3Hold=0
      queuedscore=queuedscore+1000+(1000*SilverBonusCounter)
    end if
  end if
end sub

Sub Target4_Hit
  DOF 127, DOFPulse
  If TableTilted=false then
    AddScore(100)
    if TargetLight4.state=1 then
      GatorMultiBall=true
      Kicker3Hold=0
      queuedscore=queuedscore+1000+(1000*SilverBonusCounter)
    end if
  end if

end sub

'**************************************
' Button Triggers
'**************************************

Sub Trigger1_Hit
  Button001.z=-1.5
  Bumper1Light.state=1
  Bumper4Light.state=1
  BumpersOn
end sub

Sub Trigger2_Hit
  Button002.z=-1.5
  Bumper2Light.state=1
  Bumper3Light.state=1
  BumpersOn
end sub

Sub Trigger3_Hit
  Button003.z=-1.5
  AddScore(10)
end sub

Sub Trigger4_Hit
  Button004.z=-1.5
  AddScore(10)
end sub

Sub Trigger5_Hit
  Button005.z=-1.5
  FlipperPosition=0
  ShowFlippers
  PlaySoundAtVol "ZipperFlipper", Trigger5, 1
end sub

Sub Trigger1_Unhit
  Button001.z=.5
end sub

Sub Trigger2_Unhit
  Button002.z=.5
end sub

Sub Trigger3_Unhit
  Button003.z=.5
end sub

Sub Trigger4_Unhit
  Button004.z=.5
end sub

Sub Trigger5_Unhit
  Button005.z=.5
end sub



'**************************************


Sub AddSpecial()
  PlaySound SoundFXDOF("knocker",128,DOFPulse,DOFKnocker)
  DOF 124, DOFPulse
  Credits=Credits+1
  DOF 129, DOFOn
  if Credits>15 then Credits=15
  If B2SOn Then
    Controller.B2SSetCredits Credits
  End If
  CreditsReel.SetValue(Credits)
End Sub

Sub AddSpecial2()
  PlaySound"click"
  Credits=Credits+1
  DOF 129, DOFOn
  if Credits>15 then Credits=15
  If B2SOn Then
    Controller.B2SSetCredits Credits
  End If
  CreditsReel.SetValue(Credits)
End Sub


sub IncreaseSilverBonus
  if SilverBonusCounter<9 then
    BonusLights(SilverBonusCounter).state=0
    SilverBonusCounter=SilverBonusCounter+1
    BonusLights(SilverBonusCounter).state=1
  end if
end sub

sub IncreaseSilverBonus2
  IncreaseSilverTimer.enabled=true
end sub

Sub IncreaseSilverTimer_timer
  if yy<1 then
    IncreaseSilverTimer.enabled=false

    Exit sub
  end if
  if SilverBonusCounter<9 then
    BonusLights(SilverBonusCounter).state=0
    SilverBonusCounter=SilverBonusCounter+1
    BonusLights(SilverBonusCounter).state=1
  end if
  yy=yy-1
  if yy<2 then

  end if
end sub

sub ScoreSilverBonus
  ScoreMotorStepper=0
  CollectSilverBonus.enabled=1
end sub

sub CollectSilverBonus_timer
  if SilverBonusCounter<0 then
    CollectSilverBonus.enabled=0
    SilverBonusCounter=0
    BonusLights(SilverBonusCounter).state=1
    NextBallDelay.enabled=true

  else
    if ScoreMotorStepper<5 then
      AddScore(1000)
      BonusLights(SilverBonusCounter).state=0
      SilverBonusCounter=SilverBonusCounter-1
      if SilverBonusCounter>=0 then
        BonusLights(SilverBonusCounter).state=1
      end if
      ScoreMotorStepper=ScoreMotorStepper+1
    else
      ScoreMotorStepper=0
    end if
  end if
end sub


sub ResetSilverBonus_timer
  if SilverBonusCounter<0 then
    ResetSilverBonus.enabled=0
    SilverBonusCounter=0
    BonusLights(SilverBonusCounter).state=1

  else
    if ScoreMotorStepper<5 then

      BonusLights(SilverBonusCounter).state=0
      SilverBonusCounter=SilverBonusCounter-1
      if SilverBonusCounter>=0 then
        BonusLights(SilverBonusCounter).state=1
      end if
      ScoreMotorStepper=ScoreMotorStepper+1
    else
      ScoreMotorStepper=0
    end if
  end if
end sub

Sub ShowFlippers
  if FlipperPosition=0 then
    LeftFlippers(0).enabled=true
    RightFlippers(0).enabled=true
    FlipperT2Right.visible=true
    FlipperT2Left.visible=true
    LeftFlippers(1).enabled=false
    RightFlippers(1).enabled=false
    FlipperT2Right1.visible=false
    FlipperT2Left1.visible=false
    If EnableFlipperShadow=true then
      FlipperLSh.visible=true
      FlipperLSh1.visible=false
      FlipperRSh.visible=true
      FlipperRSh1.visible=false
    end if
  else
    LeftFlippers(1).enabled=true
    RightFlippers(1).enabled=true
    FlipperT2Right1.visible=true
    FlipperT2Left1.visible=true
    LeftFlippers(0).enabled=false
    RightFlippers(0).enabled=false
    FlipperT2Right.visible=false
    FlipperT2Left.visible=false
    If EnableFlipperShadow=true then
      FlipperLSh1.visible=true
      FlipperLSh.visible=false
      FlipperRSh1.visible=true
      FlipperRSh.visible=false
    end if
  end if
  if LeftFlippersFlipped=1 then
    LeftFlippers(FlipperPosition).RotateToEnd
  else
    LeftFlippers(FlipperPosition).RotateToStart
  end if
  if RightFlippersFlipped=1 then
    RightFlippers(FlipperPosition).RotateToEnd
  else
    RightFlippers(FlipperPosition).RotateToStart
  end if
end sub


Sub AdvanceZeroToNine
  ZeroToNineUnit=ZeroToNineUnit+1
  if ZeroToNineUnit>9 then
    ZeroToNineUnit=0
  end if
  LightsZeroNine
end sub

Sub LightsZeroNine
  if ZeroToNineUnit=3 then
    TargetLight2.state=1
  else
    TargetLight2.state=0
  end if
end sub

Sub ToggleAlternatingRelay
  AdvanceZeroToNine
  if AlternatingRelay=1 then
    AlternatingRelay=2

  else
    AlternatingRelay=1


  end if
end sub

Sub ToggleRedBumper


  if Bumper2.state=1 then
    Bumper2.state=0
  Else
    Bumper2.state=1
  end if
  BumpersOn

end sub


Sub ResetBallDrops
  FlipperPosition=0
  ShowFlippers
  PlaySoundAtVol "ZipperFlipper", Trigger5, 1
  HoleCounter=0
  ScoreMotorStepper=0
  ResetSilverBonus.enabled=1
  Bumper1Light.state=0
  Bumper2Light.state=0
  Bumper3Light.state=0
  Bumper4Light.state=0
  BumpersOn
  if dooralreadyopen=1 then
    closeg.enabled=true
  end if
End Sub


Sub LightsOut

  BonusCounter=0
  HoleCounter=0
  Bumper1Light.state=0
  Bumper2Light.state=0
  Bumper3Light.state=0
  Bumper4Light.state=0

end sub

Sub ResetBalls()

  TempMultiCounter=BallsPerGame-BallInPlay

' BonusLights(SilverBonusCounter).state=1

  ResetBallDrops

  BonusMultiplier=1
  TableTilted=false
  TiltReel.SetValue(0)
  If B2SOn then
    Controller.B2SSetTilt 0
    Controller.B2SSetData 80+Player,1
  end if
  PlasticsOn
  'CreateBallID BallRelease
  Ballrelease.CreateSizedBall 25
    Ballrelease.Kick 40,7
  DOF 116, DOFPulse
  BallInPlayReel.SetValue(BallInPlay)
' InstructCard.image="IC"+FormatNumber(BallsPerGame,0)+FormatNumber(BallInPlay,0)
' CardLight1.State = ABS(CardLight1.State-1)


End Sub





sub resettimer_timer
    rst=rst+1
  if rst>1 and rst<12 then
    ResetReelsToZero(1)
  end if
    if rst=13 then
    playsound "StartBall1"
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


sub NextBallDelay_timer()
  NextBallDelay.enabled=false
  nextball

end sub

sub newgame
  InProgress=true
  queuedscore=0
  GatorMultiBall=False
  GatorWall.IsDropped=false
' GatorWall1.IsDropped=true
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
  For each obj in BonusLights
    obj.state=0
  next
  if dooralreadyopen=1 then
    closeg.enabled=true
  end if
  RightOutlaneLight.state=0
  TargetLight2.state=0
  TargetLight3.state=0
  TargetLight4.state=0
  if Kicker3Hold=1 then
    TargetLight3.state=1
    TargetLight4.state=1
  end if
  Bumper1Light.state=0
  Bumper2Light.state=0
  Bumper3Light.state=0
  Bumper4Light.state=0
  BumpersOn
  BonusCounter=0
  BallCounter=0
  TargetLeftFlag=1
  TargetCenterFlag=1
  TargetRightFlag=1
  TargetSequenceComplete=0
  LightsZeroNine
' IncreaseBonus
' ToggleBumper
  EightLit=1
  ResetBalls
End sub

sub nextball
  If B2SOn Then
    Controller.B2SSetTilt 0
    Controller.B2SSetData 81,0
    Controller.B2SSetData 82,0
    Controller.B2SSetData 83,0
    Controller.B2SSetData 84,0
  End If
  Player=Player+1

  If Player>Players Then
    BallInPlay=BallInPlay+1
    If BallInPlay>BallsPerGame then
      PlaySound("MotorLeer")
      InProgress=false
      if Kicker3Hold=1 and EndGameClearGatorHole=1 then
        GatorMultiBall=True
        KickerHolder3.enabled=0
        Kicker3Hold=0
        mHole3.MagnetOn=0
        Kicker3.kick 210,5
      end if
      If B2SOn Then
        Controller.B2SSetGameOver 1
        Controller.B2SSetPlayerUp 0
        Controller.B2SSetBallInPlay 0
        Controller.B2SSetCanPlay 0
      End If
      For each obj in PlayerHuds
        obj.state=0
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
      GameOverReel.SetValue(1)
'     InstructCard.image="IC"+FormatNumber(BallsPerGame,0)+"0"
'     CardLight1.State = ABS(CardLight1.State-1)
      BallInPlayReel.SetValue(0)
      CanPlayReel.SetValue(0)
      LeftFlipper.RotateToStart
      RightFlipper.RotateToStart
      FlipperPosition=0
      LeftFlippersFlipped=0
      RightFlippersFlipped=0
      ShowFlippers
      PlaySound "ZipperFlipper"
      LightsOut
      BumpersOff
      PlasticsOff
      checkmatch
      CheckHighScore
      Players=0
      TimerTitle1.enabled=1
      TimerTitle2.enabled=1
      TimerTitle3.enabled=1
      TimerTitle4.enabled=1
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
'     PlayerUpRotator.enabled=true
      PlayStartBall.enabled=true
      For each obj in PlayerHuds
        obj.state=0
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
        PlayerHuds(Player-1).state=1
        PlayerHUDScores(Player-1).state=1
        PlayerScores(Player-1).visible=0
        PlayerScoresOn(Player-1).visible=1
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
      obj.State=0
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
      PlayerHuds(Player-1).state=1
      PlayerHUDScores(Player-1).state=1
      PlayerScores(Player-1).visible=0
      PlayerScoresOn(Player-1).visible=1
    end If
    ResetBalls
  End If

End sub

sub openg_timer
    bottgate(bgpos).isdropped=true
    bgpos=bgpos-1
    bottgate(bgpos).isdropped=false
  primgate.RotY=30+(bgpos*10)
     if bgpos=0 then
    PlaySoundAtVol "postup", primgate, 1
    RightOutlaneLight.state=1
    openg.enabled=false
    dooralreadyopen=1
  end if

 end sub

sub closeg_timer
    closeg.interval=10
    bottgate(bgpos).isdropped=true
    bgpos=bgpos+1
    bottgate(bgpos).isdropped=false
    primgate.RotY=30+(bgpos*10)
  if bgpos=6 then
    RightOutlaneLight.state=0
    closeg.enabled=false
    dooralreadyopen=0
  end if
end sub

sub PauseGateCloser_timer
  PauseGateCloser.enabled=false
  closeg.enabled=true
end sub


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
  tempmatch=Int(Rnd*10)
  Match=tempmatch*10
  MatchReel.SetValue(tempmatch+1)

  If B2SOn Then
    If Match = 0 Then
      Controller.B2SSetMatch 100
    Else
      Controller.B2SSetMatch Match
    End If
  End if
  for i = 1 to Players
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
      PlasticsOff
      BumpersOff
      TiltReel.SetValue(1)
      LeftFlipper.RotateToStart
      RightFlipper.RotateToStart
      FlipperPosition=0
      LeftFlippersFlipped=0
      RightFlippersFlipped=0
      ShowFlippers
      PlaySound "ZipperFlipper"
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
    FlipperLSh1.visible = ShadowFlippersOn
    FlipperRSh1.visible = ShadowFlippersOn

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
    FlipperLSh1.visible=false
    FLipperRSh1.visible=false
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


    HighScore=cdbl(temp1)
    if HighScore<1 then
      temp6=textstr.readline
      temp7=textstr.readline
      temp8=textstr.readline
      temp9=textstr.readline
      temp10=textstr.readline
      temp11=textstr.readline
      temp12=textstr.readline
      temp13=textstr.readline
      temp14=textstr.readline
      temp15=textstr.readline
    end if

    TextStr.Close
      Credits=cdbl(temp2)
    BallsPerGame=cdbl(temp3)
    ReplayLevel=cdbl(temp4)
    if HighScore<1 then
      HSScore(1) = int(temp6)
      HSScore(2) = int(temp7)
      HSScore(3) = int(temp8)
      HSScore(4) = int(temp9)
      HSScore(5) = int(temp10)

      HSName(1) = temp11
      HSName(2) = temp12
      HSName(3) = temp13
      HSName(4) = temp14
      HSName(5) = temp15
    end if
    Set ScoreFile=Nothing
      Set FileObj=Nothing
end sub

Sub DisplayHighScore


end sub

sub InitPauser5_timer

    DisplayHighScore
    CreditsReel.SetValue(Credits)
    InitPauser5.enabled=false
end sub



sub BumpersOff

  Bumper1Light.visible=0


  Bumper2Light.visible=0


  Bumper3Light.visible=0

  Bumper4Light.visible=0

end sub

sub BumpersOn
  Bumper1Light.visible=0
  Bumper2Light.visible=0
  Bumper3Light.visible=0
  Bumper4Light.visible=0
  If Bumper1Light.state=1 then
    Bumper1Light.visible=1
  end if

  if Bumper2Light.state=1 then
    Bumper2Light.visible=1
  end if

  If Bumper3Light.state=1 then
    Bumper3Light.visible=1
  end if

  If Bumper4Light.state=1 then
    Bumper4Light.visible=1
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

  Replay1Table(1)=32000
  Replay1Table(2)=33000
  Replay1Table(3)=34000
  Replay1Table(4)=36000
  Replay1Table(5)=38000
  Replay1Table(6)=40000
  Replay1Table(7)=58000
  Replay1Table(8)=60000
  Replay1Table(9)=62000
  Replay1Table(10)=64000
  Replay1Table(11)=6300
  Replay1Table(12)=6600
  Replay1Table(13)=6800
  Replay1Table(14)=7300
  Replay1Table(15)=999000

  Replay2Table(1)=44000
  Replay2Table(2)=45000
  Replay2Table(3)=46000
  Replay2Table(4)=48000
  Replay2Table(5)=50000
  Replay2Table(6)=52000
  Replay2Table(7)=74000
  Replay2Table(8)=76000
  Replay2Table(9)=78000
  Replay2Table(10)=80000
  Replay2Table(11)=7700
  Replay2Table(12)=8000
  Replay2Table(13)=8200
  Replay2Table(14)=8200
  Replay2Table(15)=999000

  Replay3Table(1)=56000
  Replay3Table(2)=57000
  Replay3Table(3)=58000
  Replay3Table(4)=60000
  Replay3Table(5)=62000
  Replay3Table(6)=64000
  Replay3Table(7)=86000
  Replay3Table(8)=88000
  Replay3Table(9)=90000
  Replay3Table(10)=92000
  Replay3Table(11)=8500
  Replay3Table(12)=8800
  Replay3Table(13)=8900
  Replay3Table(14)=8900
  Replay3Table(15)=999000

  Replay4Table(1)=68000
  Replay4Table(2)=69000
  Replay4Table(3)=70000
  Replay4Table(4)=72000
  Replay4Table(5)=74000
  Replay4Table(6)=76000
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


  ReplayCard.image = "SC" + tempst2
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
      Case 50:
        MotorMode=10
        MotorPosition=5
        BumpersOff
      Case 100:
        MotorMode=100
        MotorPosition=1

      Case 500:
        MotorMode=100
        MotorPosition=5
        BumpersOff
      Case 1000:
        MotorMode=1000
        MotorPosition=1
      Case 2000:
        MotorMode=1000
        MotorPosition=2
        BumpersOff
      Case 3000:
        MotorMode=1000
        MotorPosition=3
        BumpersOff
      Case 4000:
        MotorMode=1000
        MotorPosition=4
        BumpersOff
      Case 5000:
        MotorMode=1000
        MotorPosition=5
        BumpersOff
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
  If Score(Player)>Replay4 and Replay4Paid(Player)=false then
    Replay4Paid(Player)=True
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

    OldScore = int(OldScore / 10) ' divide by 10 for games with fixed 0 in 1s position, by 1 for games with real 1s digits
    NewScore = int(NewScore / 10) ' divide by 10 for games with fixed 0 in 1s position, by 1 for games with real 1s digits
  ' MsgBox("OldScore="&OldScore&", NewScore="&NewScore&", OldScore Mod 10="&OldScore Mod 10 & ", NewScore % 10="&NewScore Mod 10)

    if (OldScore Mod 10 <> NewScore Mod 10) then
    PlayChime(10)
    ToggleAlternatingRelay
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
        PlaySound SoundFXDOF("10a",140,DOFPulse,DOFChimes)
      Case 100
        PlaySound SoundFXDOF("100a",141,DOFPulse,DOFChimes)
      Case 1000
        PlaySound SoundFXDOF("1000a",142,DOFPulse,DOFChimes)
    End Select
  else
    Select Case x
      Case 10
        If LastChime10=1 Then
          PlaySound SoundFXDOF("SJ_Chime_10a",140,DOFPulse,DOFChimes)
          LastChime10=0
        Else
          PlaySound SoundFXDOF("SJ_Chime_10b",140,DOFPulse,DOFChimes)
          LastChime10=1
        End If
      Case 100
        If LastChime100=1 Then
          PlaySound SoundFXDOF("SJ_Chime_100a",141,DOFPulse,DOFChimes)
          LastChime100=0
        Else
          PlaySound SoundFXDOF("SJ_Chime_100b",141,DOFPulse,DOFChimes)
          LastChime100=1
        End If
      Case 1000
        If LastChime1000=1 Then
          PlaySound SoundFXDOF("SJ_Chime_1000a",142,DOFPulse,DOFChimes)
          LastChime1000=0
        Else
          PlaySound SoundFXDOF("SJ_Chime_1000b",142,DOFPulse,DOFChimes)
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
Const OptionLinesToMark="111000011"
Const OptionLine1="" 'do not use this line
Const OptionLine2="" 'do not use this line
Const OptionLine3="" 'do not use this line
Const OptionLine4=""
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
        BallCard.image="BC"+FormatNumber(BallsPerGame,0)
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



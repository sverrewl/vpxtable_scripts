'*
'*        Gottlieb's Capt. Card (1973)
'*        Table primary build/scripted by Loserman76
'*        Table images by GNance
'*        http://www.ipdb.org/machine.cgi?id=1173
'*
'*

option explicit
Randomize
ExecuteGlobal GetTextFile("core.vbs")

On Error Resume Next
ExecuteGlobal GetTextFile("Controller.vbs")
If Err Then MsgBox "Unable to open Controller.vbs. Ensure that it is in the scripts folder."
On Error Goto 0

Const cGameName = "CaptCard_1973"

Const ShadowFlippersOn = true
Const ShadowBallOn = true

Const ShadowConfigFile = false

dim ballsize,ballmass

ballsize=1
ballmass=1

' Thalamus 2020 March : Improved directional sounds
' !! NOTE : Table not verified yet !!

' Options
' Volume devided by - lower gets higher sound

Const VolDiv = 2000    ' Lower number, louder ballrolling/collition sound

' The rest of the values are multipliers
'
'  .5 = lower volume
' 1.5 = higher volume

Const VolFlip   = 1    ' Flipper volume.


Dim Controller  ' B2S
Dim B2SScore  ' B2S Score Displayed
Const HSFileName="CaptCard_73_2VPX.txt"
Const B2STableName="CaptCard_1973"
Const LMEMTableConfig="LMEMTables.txt"
Const LMEMShadowConfig="LMEMShadows.txt"
Dim EnableBallShadow
Dim EnableFlipperShadow

'* this value adjusts score motor behavior - 0 allows you to continue scoring while the score motor is running - 1 sets score motor to behave more like a real EM
Const ScoreMotorAdjustment=1

'* this is a debug setting to use an older scoring routine vs a newer score routine - don't change this value
Const ScoreAdditionAdjustment=1

'* this controls whether you hear bells (0) or chimes (1) when scoring
Const ChimesOn=1

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

Dim TargetSetting
Dim BonusBooster
Dim BonusBoosterCounter
Dim BonusCounter
Dim HoleCounter

Dim ZeroNineCounter
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

dim raceRflag
dim raceAflag
dim raceCflag
dim raceEflag
dim race5flag

Dim LStep, LStep2, RStep, L2Step, R2Step, xx

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

Dim Kicker1Hold,Kicker2Hold

Dim mHole,mHole2, mHole3, mHole4, mHole5

Dim HeartTargetsDownCounter, ClubTargetsDownCounter, SpadeTargetsDownCounter, DiamondTargetsDownCounter

Dim HeartsCompleted,SpadesCompleted

Dim SpinPos,Spin,Count,Count1,Count2,Count3,Reset,VelX,VelY,LitSpinner, LastRotation

Dim HeartComplete
Dim ClubComplete
Dim DiamondComplete
Dim SpadeComplete
Dim tempComplete

Dim Bumper1State
Dim tKickerTCount

Dim GameOn

Dim LHeart,Club,Diamond,Spade, RedSuits,BlackSuits

Spin=Array("5","6","7","8","9","10","J","Q","K","A")

Dim BootCount:BootCount = 0

Sub BootTable_Timer()
  If BootCount = 0 Then
    BootCount = 1
    me.interval = 100
    PlaySoundAt "poweron", Plunger
    If B2SOn Then
      Controller.B2SSetGameOver 1
      Controller.B2SSetData 80,1
    End If

    '*****GI Lights On
    For each xx in GIlights:xx.state = 1: Next
    For each obj in bakery: obj.image= "bakeryon": Next
    For each obj in stars: obj.image = "starson": Next
    flasher1.visible=1
    flasher2.visible=0
    plastics.image="plasticsetcon"
    upperobjects.image="plasticsetcon"
    holes.image="holeson"
    uppermetal.image="plasticedgeson"
    plasticedges.image="plasticedgeson"
    posts1.image="bakeryon"
    rubbers1.image="bakeryon"
    topgate.image="bracketon"
    outers.image="outerson"
    diamonda4.image="dtbanktopleft"
    diamonda3.image="dtbanktopleft"
    diamonda2.image="dtbanktopleft"
    diamonda1.image="dtbanktopleft"
    spadea1.image="spadesproject"
    spadea2.image="spadesproject"
    spadea3.image="spadesproject"
    spadea4.image="spadesproject"
    cluba1.image="clubsrendered"
    cluba2.image="clubsrendered"
    cluba3.image="clubsrendered"
    cluba4.image="clubsrendered"
    hearta1.image="heartsdrops"
    hearta2.image="heartsdrops"
    hearta3.image="heartsdrops"
    hearta4.image="heartsdrops"
    lslingmetal.image="lslingmetal"
    lcentermetal.image="lcentermetal"
    rslingmetal.image="rslingmetal"
    rcentermetal.image="rcentermetal"
    centermetal.image="centermetal"
    lplasticguide.image="lplasticguide"
    rplasticguide.image="rplasticguide"
    screws.image="screwson"
    lsling0.image="lslingon"
    plungeplate.image="plungeplate"
    dropspade1.visible=1
    dropspade2.visible=1
    dropspade3.visible=1
    dropspade4.visible=1
    dropdiamond1.visible=1
    dropdiamond2.visible=1
    dropdiamond3.visible=1
    dropdiamond4.visible=1
    dropclub1.visible=1
    dropclub2.visible=1
    dropclub3.visible=1
    dropclub4.visible=1
    dropheart1.visible=1
    dropheart2.visible=1
    dropheart3.visible=1
    dropheart4.visible=1
    If Credits > 0 then
      apron.Image = "aproncredit"
    Else
      apron.Image = "apronon"
    End If
    GameOn = 1
    me.enabled = False
  End If
End Sub

Sub Table1_Init()
  If Table1.ShowDT = false then
    For each obj in DesktopCrap
      obj.visible=False
    next
  End If

  LoadEM
  LoadLMEMConfig2

  Count=0
    Count1=0
    Count2=0
  Count3=0
    Reset=0
  ZeroToNineUnit=Int(Rnd*10)
  Kicker1Hold=0

  EightLit=0

  BallCounter=0
  ReelCounter=0
  AddABall=0
  DropABall=0
  HideOptions
  SetupReplayTables
  PlasticsOff
  BumpersOff
  OperatorMenu=0
  HighScore=0
  MotorRunning=0
  HighScoreReward=3
  Credits=0
  BallsPerGame=5
  ReplayLevel=1
  TiltEndsGame=1
  TargetSetting=0
  AlternatingRelay=1
  ZeroNineCounter=0
  SpecialLightOption=2
  BackglassBallFlagColor=1
  loadhs
  if HighScore=0 then HighScore=50000


  TableTilted=false

  Match=int(Rnd*10)
  MatchReel.SetValue((Match)+1)
  Match=Match*10
  TiltReel.SetValue(1)
' CanPlayReel.SetValue(0)
  GameOverReel.SetValue(1)

  For each obj in HUDCards
    obj.SetValue(0)
  next

  For each obj in PlayerHuds
    obj.SetValue(0)
  next
  For each obj in PlayerScores
    obj.ResetToZero
  next

  for each obj in bottgate
    obj.isdropped=true
    next

  for each obj in Bonus
    obj.state=0
  next
  for each obj in TargetSpecialLights
    obj.state=0
  next
  for each obj in TargetAdvanceLights
    obj.state=0
  next
  for each obj in RolloverAdvanceLights
    obj.state=0
  next


  For each obj in NumberLights
    obj.state=0
  next

  Replay1=Replay1Table(ReplayLevel)
  Replay2=Replay2Table(ReplayLevel)
  Replay3=Replay3Table(ReplayLevel)
  Replay4=Replay4Table(ReplayLevel)

  BonusCounter=0
  HoleCounter=0


' InstructCard.image="IC"+FormatNumber(BallsPerGame,0)+FormatNumber(TiltEndsGame,0)
'
' RefreshReplayCard

  CurrentFrame=0


  OperatorMenuBackdrop.image = "PostitBL"
  For XOpt = 1 to MaxOption
    Eval("OperatorOption"&XOpt).image = "PostitBL"
  next

  For XOpt = 1 to 256
    Eval("Option"&XOpt).image = "PostItBL"
  next

  AdvanceLightCounter=0

  Players=0
  RotatorTemp=1
  InProgress=false
  TargetLightsOn=false

  HUD100K.SetValue(0)
  ScoreText.text=HighScore


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

    Controller.B2SSetScore 3,HighScore
    Controller.B2SSetTilt 1
    Controller.B2SSetCredits Credits
    Controller.B2SSetGameOver 1
  End If

  for i=1 to 1
    player=i
    If B2SOn Then
      Controller.B2SSetScorePlayer player, 0
    End If
  next
  bump1=1
  bump2=1
  InitPauser5.enabled=true
  If Credits > 0 Then DOF 116, DOFOn
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
      Controller.B2SSetData 10,0
      raceEflag=0
    else
      Controller.B2SSetData 10,1
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
'   LeftFlipper.RotateToEnd
    LF.fire 'LeftFlipper.RotateToEnd
    PlayLoopSoundAtVol "buzzL",LeftFlipper, VolFlip
    PlaySoundAtVol SoundFXDOF("FlipperUp",101,DOFOn,DOFFlippers), LeftFlipper, VolFlip
  End If

  If keycode = RightFlipperKey  and InProgress=true and TableTilted=false Then
'   RightFlipper.RotateToEnd
    RF.fire 'RightFlipper.RotateToEnd
    PlaySoundAtVol SoundFXDOF("FlipperUp",102,DOFOn,DOFFlippers), LeftFlipper, VolFlip
    PlayLoopSoundAtVol "buzz", LeftFlipper, VolFlip
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

    playsound "coinin"
    AddSpecial2
    apron.Image = "aproncredit"
  end if

   if keycode = 5 then
    playsound "coinin"
    AddSpecial2
    apron.Image = "aproncredit"
    keycode= StartGameKey
  end if


  if keycode=StartGameKey and Credits>0 and InProgress=false and Players=0 and EnteringOptions = 0 then
'GNMOD
    OperatorMenuTimer.Enabled = false
'END GNMOD
'   Credits=Credits-1
    If Credits < 1 Then DOF 116, DOFOff
    CreditsReel.SetValue(Credits)
    Players=1
'   CanPlayReel.SetValue(Players)
    MatchReel.SetValue(0)
    Player=1
    playsound "startup_norm"
    TempPlayerUp=Player
'   PlayerUpRotator.enabled=true
    rst=0
    BallInPlay=BallsPerGame
    InProgress=true
    resettimer.enabled=true
    ResetDropsTimer.enabled=true
    BonusMultiplier=1
    HUD100K.SetValue(0)
    TimerTitle1.enabled=0
    TimerTitle2.enabled=0
    TimerTitle3.enabled=0
    TimerTitle4.enabled=0
    If B2SOn Then
      Controller.B2SSetTilt 0
      Controller.B2SSetGameOver 0
      Controller.B2SSetMatch 0
      Controller.B2SSetCredits Credits
      Controller.B2SSetScore 3,HighScore
      Controller.B2SSetCanPlay 1
      Controller.B2SSetPlayerUp 1
      Controller.B2SSetBallInPlay BallInPlay
      Controller.B2SSetScoreRolloverPlayer1 0

      Controller.B2SSetData 11,0
      Controller.B2SSetData 12,0
      Controller.B2SSetData 13,0
      Controller.B2SSetData 10,0
      Controller.B2SSetData 80,1
      for i = 81 to 90
        Controller.B2SSetData i,0
      next
    End If
    For each obj in PlayerScores
      obj.ResetToZero
    next
    For each obj in PlayerHuds
      obj.SetValue(0)
    next

    PlayerHuds(Player-1).SetValue(1)
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

    PlaySoundAtVol "plungerrelease", Plunger, 1
    Plunger.Fire
  End If

  if keycode = LeftFlipperKey then
    OperatorMenuTimer.Enabled = false
  end if

  ' END GNMOD

  If keycode = LeftFlipperKey and InProgress=true and TableTilted=false Then
    LeftFlipper.RotateToStart
    PlaySoundAtVol SoundFXDOF("FlipperDown",101,DOFOff,DOFFlippers), LeftFlipper, VolFlip
    StopSound "buzzL"
  End If

  If keycode = RightFlipperKey and InProgress=true and TableTilted=false Then
    RightFlipper.RotateToStart
        StopSound "buzz"
    PlaySoundAtVol SoundFXDOF("FlipperDown",102,DOFOff,DOFFlippers), RightFlipper, VolFlip
  End If

End Sub



Sub Drain_Hit()

  Drain.DestroyBall
  PlaySoundAtVol "fx_drain", Drain, 1
  DOF 117, DOFPulse
  NewBonusHolder.enabled=true

End Sub

Sub Pause4Bonustimer_timer
  If NewBonusTimer.enabled=0 then
    Pause4Bonustimer.enabled=0
    NextBallDelay.enabled=true
  end if

End Sub

Sub Trigger1_unhit
  DOF 118, DOFPulse
End Sub

Sub NewBonusHolder_timer
  if NewBonusTimer.enabled=0 then
    NewBonusHolder.enabled=0
    NextBallDelay.enabled=true
    If TargetSequenceComplete = 1 then
      ResetDropsTimer.enabled = 1
      TargetSequenceComplete = 0
      for each obj in NumberLights
        obj.state=0
      next

      for each obj in LightSpecial
        obj.state=0
      next
      For each obj in ResetLights
        obj.state=0
      next
      TopLeftLight.state=0
      TopRightLight.state=0
      LeftLight1.state=0
      LeftLight2.state=0
      RightLight1.state=0
      RightLight2.state=0
      LeftInlaneLight.state=0
      RightInlaneLight.state=0
    end if
  end if

end sub

'***********************
'     Flipper Logos
'***********************

Sub UpdateFlipperLogos_Timer
  LFlip.Rotz = LeftFlipper.CurrentAngle
  LFlipr.Rotz = LeftFlipper.CurrentAngle
  RFlip.Rotz = RightFlipper.CurrentAngle
  RFlipr.Rotz = RightFlipper.CurrentAngle
  PGate.Rotz = (Gate.CurrentAngle*.75) + 25
  FlipperLSh.RotZ = LeftFlipper.currentangle
  FlipperRSh.RotZ = RightFlipper.currentangle
    If gameon=1 and Bumper1Light.state=1 then
    bump.Image = "bumpcaplit"
    Else
    If gameon=1 then bump.image="bumpcapon": end if
    End If
    If gameon=1 and Bumper1Light.state=1 then
    bumps1.Image = "bumpcap1lit"
    Else
    If gameon=1 then bumps1.image="bumpcap1on": end if
    End If
    If gameon=1 and Bumper1Light.state=1 then
    bumps2.Image = "bumpbaselit"
    Else
    If gameon=1 then bumps2.image="bumpbase": end if
    End If
    If gameon=1 and Bumper1Light.state=1 then
    bumps3.Image = "bumpcap3lit"
    Else
    If gameon=1 then bumps3.image="bumpcap3on": end if
    End If
    If gameon=1 and Credits > 0 then
    apron.Image = "aproncredit"
    Else
    If gameon=1 then apron.Image = "apronon": end if
    End If
    If bumper1Light.state=1 then flasher3.visible=1: end If
    If bumper1Light.state=0 then flasher3.visible=0: end if
End Sub


'***********************
' slingshots
'

Sub RightSlingShot_Slingshot
  PlaySoundAtVol SoundFXDOF("right_slingshot",104,DOFPulse,DOFContactors), sling1, 1
    RSling0.Visible = 0
    RSling1.Visible = 1
    sling1.Rotx = 30
    RStep = 0
    RightSlingShot.TimerEnabled = 1
  AddScore 10
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.Rotx = 10
        Case 4:RSLing2.Visible = 0:RSling3.Visible = 1:sling1.Rotx = 0
        Case 5:RSLing3.Visible = 0:RSling0.Visible = 1:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  PlaySoundAtVol SoundFXDOF("left_slingshot",103,DOFPulse,DOFContactors), sling2, 1
    LSling0.Visible = 0
    LSling1.Visible = 1
    sling2.RotX = -30
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
  AddScore 10
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.Rotx = -10
        Case 4:LSLing2.Visible = 0:LSLing3.Visible = 1:sling2.Rotx = 0
        Case 5:LSLing3.Visible = 0:LSLing0.Visible = 1:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

'***********************************
' Walls
'***********************************

Sub RubberWallSwitches_Hit(idx)
  if TableTilted=false then
    If BallsPerGame=5 then
      AddScore(100)
    else
      AddScore(1000)
    end if
  end if
end Sub

dim llstep, lrstep, rlstep, rrstep

Sub RubberWallSwitch1_Hit
  PlaySoundAtVol "target", ActiveBall, 1
    lrubber.Visible = 0
    lrubber1.Visible = 1
    llStep = 0
    RubberWallSwitch1a.Enabled = 1
End Sub

Sub RubberWallSwitch1a_Timer
    Select Case llStep
        Case 3:lrubber1.Visible = 0:lrubber2.Visible = 1
        Case 4:lrubber2.Visible = 0:lrubber3.Visible = 1
        Case 5:lrubber3.Visible = 0:lrubber.Visible = 1:RubberWallSwitch1a.Enabled = 0
    End Select
    llStep = llStep + 1
End Sub

Sub RubberWallSwitch4_Hit
  PlaySoundAtVol "target", ActiveBall, 1
    lrrubber.Visible = 0
    lrrubber1.Visible = 1
    lrStep = 0
    RubberWallSwitch4a.Enabled = 1
End Sub

Sub RubberWallSwitch4a_Timer
    Select Case lrStep
        Case 3:lrrubber1.Visible = 0:lrrubber2.Visible = 1
        Case 4:lrrubber2.Visible = 0:lrrubber3.Visible = 1
        Case 5:lrrubber3.Visible = 0:lrrubber.Visible = 1:RubberWallSwitch4a.Enabled = 0
    End Select
    lrStep = lrStep + 1
End Sub

Sub RubberWallSwitch2_Hit
  PlaySoundAtVol "target", ActiveBall, 1
    rrrubber.Visible = 0
    rrrubber1.Visible = 1
    rrStep = 0
    RubberWallSwitch2a.Enabled = 1
End Sub

Sub RubberWallSwitch2a_Timer
    Select Case rrStep
        Case 3:rrrubber1.Visible = 0:rrrubber2.Visible = 1
        Case 4:rrrubber2.Visible = 0:rrrubber3.Visible = 1
        Case 5:rrrubber3.Visible = 0:rrrubber.Visible = 1:RubberWallSwitch2a.Enabled = 0
    End Select
    rrStep = rrStep + 1
End Sub

Sub RubberWallSwitch3_Hit
  PlaySoundAtVol "target", ActiveBall, 1
    rlrubber.Visible = 0
    rlrubber1.Visible = 1
    rlStep = 0
    RubberWallSwitch3a.Enabled = 1
End Sub

Sub RubberWallSwitch3a_Timer
    Select Case rlStep
        Case 3:rlrubber1.Visible = 0:rlrubber2.Visible = 1
        Case 4:rlrubber2.Visible = 0:rlrubber3.Visible = 1
        Case 5:rlrubber3.Visible = 0:rlrubber.Visible = 1:RubberWallSwitch3a.Enabled = 0
    End Select
    rlStep = rlStep + 1
End Sub

'***********************************
' Bumpers
'***********************************

dim bump1step

Sub Bumper1_Hit
  If TableTilted=false then
  bumps3.RotY=-5
  Bump1Step = 0
  Bumper1.timerenabled = 1
  PlaySoundAtVol SoundFXDOF("bumper1",105,DOFPulse,DOFContactors), ActiveBall, 1
    bump1 = 1
    If Bumper1Light.state=1 then
      AddScore(1000)
      ToggleAlternatingRelay
    end if
    end if
End Sub

Sub Bumper1_timer
  Select Case Bump1Step
    Case 3: bumps3.RotY=-3
    Case 4: bumps3.RotY=2
    Case 5: bumps3.RotY=-1
    Case 6: bumps3.RotY=1
    Case 7: bumps3.RotY=0:Bumper1.timerenabled = 0
  End Select
  Bump1Step=Bump1Step + 1
End Sub

'************************************
'  Rollover lanes
'************************************

Sub TriggerCollection_Hit(idx)
' TriggerWires(idx).TopVisible=0
  DOF 200+idx, DOFPulse
  If TableTilted=false then
    Select Case idx
      Case 0:
        If RightInlaneLight.state=1 then
          AddScore(1000)
        else
          AddScore(100)
        end if

      Case 1:
        If LeftOutlaneLight.state=1 then
          AddSpecial
        end if
        ScoreBonus

      Case 2:
        If RightOutlaneLight.state=1 then
          AddSpecial
        end if
        ScoreBonus

      Case 3:
        If LeftInlaneLight.state=1 then
          AddScore(1000)
        else
          AddScore(100)
        end if

      Case 4:
        If TopLeftLight.state=1 then
          AddScore(1000)
        else
          AddScore(100)
        end if

      Case 5:
        If TopRightLight.state=1 then
          AddScore(1000)
        else
          AddScore(100)
        end if

      Case 6:
        If LeftLight1.state=1 then
          AddScore(1000)
        else
          AddScore(100)
        end if

      Case 7:
        If LeftLight2.state=1 then
          AddScore(1000)
        else
          AddScore(100)
        end if

      Case 8:
        If RightLight1.state=1 then
          AddScore(1000)
        else
          AddScore(100)
        end if

      Case 9:
        If RightLight2.state=1 then
          AddScore(1000)
        else
          AddScore(100)
        end if

    End Select

  end if

end Sub

'****************************************************

dim s1step,s2step,s3step,s4step,c1step,c2step,c3step,c4step,d1step,d2step,d3step,d4step,h1step,h2step,h3step,h4step

Sub Spade1_Hit()
  If (Spade1.IsDropped = False) and (TableTilted=false) then
    PlaySoundAtVol SoundFXDOF("droptargetdropped",107,DOFPulse,DOFDropTargets), ActiveBall, 1
    SetMotor(500)
    SpadeTargetsDownCounter=SpadeTargetsDownCounter+1
    SpadeLight1.state=1
    Spade1.IsDropped = True
    Spade1.collidable = False
    dropspade1.visible=0
    SpadeA1.transz=-12
    s1step = 0
    spade1.timerenabled=1
    CheckAllDropTargets
  End If
End Sub

Sub Spade1_Timer
  Select Case s1step
        Case 3:SpadeA1.transz=-24
        Case 4:SpadeA1.transz=-36
    Case 5:SpadeA1.transz=-43:spade1.TimerEnabled = 0
    End Select
    s1Step = s1Step + 1
End Sub

Sub Spade2_Hit()
  If (Spade2.IsDropped = False) and (TableTilted=false) then
    PlaySoundAtVol SoundFXDOF("droptargetdropped",107,DOFPulse,DOFDropTargets), ActiveBall, 1
    SetMotor(500)
    SpadeTargetsDownCounter=SpadeTargetsDownCounter+1
    SpadeLight2.state=1
    Spade2.IsDropped = True
    Spade2.collidable = False
    dropspade2.visible=0
    SpadeA2.transz=-12
    s2step = 0
    spade2.timerenabled=1
    CheckAllDropTargets
  End If
End Sub

Sub Spade2_Timer
  Select Case s2step
        Case 3:SpadeA2.transz=-24
        Case 4:SpadeA2.transz=-36
    Case 5:SpadeA2.transz=-43:spade2.TimerEnabled = 0
    End Select
    s2Step = s2Step + 1
End Sub

Sub Spade3_Hit()
  If (Spade3.IsDropped = False) and (TableTilted=false) then
    PlaySoundAtVol SoundFXDOF("droptargetdropped",107,DOFPulse,DOFDropTargets), ActiveBall, 1
    SetMotor(500)
    SpadeTargetsDownCounter=SpadeTargetsDownCounter+1
    SpadeLight3.state=1
    Spade3.IsDropped = True
    Spade3.collidable = False
    dropspade3.visible=0
    SpadeA3.transz=-12
    s3step = 0
    spade3.timerenabled=1
    CheckAllDropTargets
  End If
End Sub

Sub Spade3_Timer
  Select Case s3step
        Case 3:SpadeA3.transz=-24
        Case 4:SpadeA3.transz=-36
    Case 5:SpadeA3.transz=-43:spade3.TimerEnabled = 0
    End Select
    s3Step = s3Step + 1
End Sub

Sub Spade4_Hit()
  If (Spade4.IsDropped = False) and (TableTilted=false) then
    PlaySoundAtVol SoundFXDOF("droptargetdropped",107,DOFPulse,DOFDropTargets), ActiveBall, 1
    SetMotor(500)
    SpadeTargetsDownCounter=SpadeTargetsDownCounter+1
    SpadeLight4.state=1
    Spade4.IsDropped = True
    Spade4.collidable = False
    dropspade4.visible=0
    SpadeA4.transz=-12
    s4step = 0
    spade4.timerenabled=1
    CheckAllDropTargets
  End If
End Sub

Sub Spade4_Timer
  Select Case s4step
        Case 3:SpadeA4.transz=-24
        Case 4:SpadeA4.transz=-36
    Case 5:SpadeA4.transz=-43:spade4.TimerEnabled = 0
    End Select
    s4Step = s4Step + 1
End Sub

Sub Diamond1_Hit()
  If (Diamond1.IsDropped = False) and (TableTilted=false) then
    PlaySoundAtVol SoundFXDOF("droptargetdropped",106,DOFPulse,DOFDropTargets), ActiveBall, 1
    SetMotor(500)
    DiamondTargetsDownCounter=DiamondTargetsDownCounter+1
    DiamondLight1.state=1
    Diamond1.IsDropped = True
    diamond1.collidable = False
    dropdiamond1.visible=0
    diamondA1.transz=-12
    d1step = 0
    diamond1.timerenabled=1
    CheckAllDropTargets
  End If
End Sub

Sub diamond1_Timer
  Select Case d1step
        Case 3:diamondA1.transz=-24
        Case 4:diamondA1.transz=-36
    Case 5:diamondA1.transz=-43:diamond1.TimerEnabled = 0
    End Select
    d1Step = d1Step + 1
End Sub

Sub Diamond2_Hit()
  If (Diamond2.IsDropped = False) and (TableTilted=false) then
    PlaySoundAtVol SoundFXDOF("droptargetdropped",106,DOFPulse,DOFDropTargets), ActiveBall, 1
    SetMotor(500)
    DiamondTargetsDownCounter=DiamondTargetsDownCounter+1
    DiamondLight2.state=1
    Diamond2.IsDropped = True
    diamond2.collidable = False
    dropdiamond2.visible=0
    diamondA2.transz=-12
    d2step = 0
    diamond2.timerenabled=1
    CheckAllDropTargets
  End If
End Sub

Sub diamond2_Timer
  Select Case d2step
        Case 3:diamondA2.transz=-24
        Case 4:diamondA2.transz=-36
    Case 5:diamondA2.transz=-43:diamond2.TimerEnabled = 0
    End Select
    d2Step = d2Step + 1
End Sub

Sub Diamond3_Hit()
  If (Diamond3.IsDropped = False) and (TableTilted=false) then
    PlaySoundAtVol SoundFXDOF("droptargetdropped",106,DOFPulse,DOFDropTargets), ActiveBall, 1
    SetMotor(500)
    DiamondTargetsDownCounter=DiamondTargetsDownCounter+1
    DiamondLight3.state=1
    Diamond3.IsDropped = True
    diamond3.collidable = False
    dropdiamond3.visible=0
    diamondA3.transz=-12
    d3step = 0
    diamond3.timerenabled=1
    CheckAllDropTargets
  End If
End Sub

Sub diamond3_Timer
  Select Case d3step
        Case 3:diamondA3.transz=-24
        Case 4:diamondA3.transz=-36
    Case 5:diamondA3.transz=-43:diamond3.TimerEnabled = 0
    End Select
    d3Step = d3Step + 1
End Sub

Sub Diamond4_Hit()
  If (Diamond4.IsDropped = False) and (TableTilted=false) then
    PlaySoundAtVol SoundFXDOF("droptargetdropped",106,DOFPulse,DOFDropTargets), ActiveBall, 1
    SetMotor(500)
    DiamondTargetsDownCounter=DiamondTargetsDownCounter+1
    DiamondLight4.state=1
    Diamond4.IsDropped = True
    diamond4.collidable = False
    dropdiamond4.visible=0
    diamondA4.transz=-12
    d4step = 0
    diamond4.timerenabled=1
    CheckAllDropTargets
  End If
End Sub

Sub diamond4_Timer
  Select Case d4step
        Case 3:diamondA4.transz=-24
        Case 4:diamondA4.transz=-36
    Case 5:diamondA4.transz=-43:diamond4.TimerEnabled = 0
    End Select
    d4Step = d4Step + 1
End Sub

Sub Club1_Hit()
  If (Club1.IsDropped = False) and (TableTilted=false) then
    PlaySoundAtVol SoundFXDOF("droptargetdropped",108,DOFPulse,DOFDropTargets), ActiveBall, 1
    SetMotor(500)
    ClubTargetsDownCounter=ClubTargetsDownCounter+1
    ClubLight1.state=1
    Club1.IsDropped = True
    club1.collidable = False
    dropclub1.visible=0
    clubA1.transz=-12
    c1step = 0
    club1.timerenabled=1
    CheckAllDropTargets
  End If
End Sub

Sub club1_Timer
  Select Case c1step
        Case 3:clubA1.transz=-24
        Case 4:clubA1.transz=-36
    Case 5:clubA1.transz=-42:club1.TimerEnabled = 0
    End Select
    c1Step = c1Step + 1
End Sub

Sub Club2_Hit()
  If (Club2.IsDropped = False) and (TableTilted=false) then
    PlaySoundAtVol SoundFXDOF("droptargetdropped",108,DOFPulse,DOFDropTargets), ActiveBall, 1
    SetMotor(500)
    ClubTargetsDownCounter=ClubTargetsDownCounter+1
    ClubLight2.state=1
    Club2.IsDropped = True
    club2.collidable = False
    dropclub2.visible=0
    clubA2.transz=-12
    c2step = 0
    club2.timerenabled=1
    CheckAllDropTargets
  End If
End Sub

Sub club2_Timer
  Select Case c2step
        Case 3:clubA2.transz=-24
        Case 4:clubA2.transz=-36
    Case 5:clubA2.transz=-42:club2.TimerEnabled = 0
    End Select
    c2Step = c2Step + 1
End Sub

Sub Club3_Hit()
  If (Club3.IsDropped = False) and (TableTilted=false) then
    PlaySoundAtVol SoundFXDOF("droptargetdropped",108,DOFPulse,DOFDropTargets), ActiveBall, 1
    SetMotor(500)
    ClubTargetsDownCounter=ClubTargetsDownCounter+1
    ClubLight3.state=1
    Club3.IsDropped = True
    club3.collidable = False
    dropclub3.visible=0
    clubA3.transz=-12
    c3step = 0
    club3.timerenabled=1
    CheckAllDropTargets
  End If
End Sub

Sub club3_Timer
  Select Case c3step
        Case 3:clubA3.transz=-24
        Case 4:clubA3.transz=-36
    Case 5:clubA3.transz=-42:club3.TimerEnabled = 0
    End Select
    c3Step = c3Step + 1
End Sub

Sub Club4_Hit()
  If (Club4.IsDropped = False) and (TableTilted=false) then
    PlaySoundAtVol SoundFXDOF("droptargetdropped",108,DOFPulse,DOFDropTargets), ActiveBall, 1
    SetMotor(500)
    ClubTargetsDownCounter=ClubTargetsDownCounter+1
    ClubLight4.state=1
    Club4.IsDropped = True
    club4.collidable = False
    dropclub4.visible=0
    clubA4.transz=-12
    c4step = 0
    club4.timerenabled=1
    CheckAllDropTargets
  End If
End Sub

Sub club4_Timer
  Select Case c4step
        Case 3:clubA4.transz=-24
        Case 4:clubA4.transz=-36
    Case 5:clubA4.transz=-42:club4.TimerEnabled = 0
    End Select
    c4Step = c4Step + 1
End Sub

Sub Heart1_Hit()
  If (Heart1.IsDropped = False) and (TableTilted=false) then
    PlaySoundAtVol SoundFXDOF("droptargetdropped",109,DOFPulse,DOFDropTargets), ActiveBall, 1
    SetMotor(500)
    HeartTargetsDownCounter=HeartTargetsDownCounter+1
    HeartLight1.state=1
    Heart1.IsDropped = True
    heart1.collidable = False
    dropheart1.visible=0
    heartA1.transz=-12
    h1step = 0
    heart1.timerenabled=1
    CheckAllDropTargets
  End If
End Sub

Sub heart1_Timer
  Select Case h1step
        Case 3:heartA1.transz=-24
        Case 4:heartA1.transz=-36
    Case 5:heartA1.transz=-42:heart1.TimerEnabled = 0
    End Select
    h1Step = h1Step + 1
End Sub

Sub Heart2_Hit()
  If (Heart2.IsDropped = False) and (TableTilted=false) then
    PlaySoundAtVol SoundFXDOF("droptargetdropped",109,DOFPulse,DOFDropTargets), ActiveBall, 1
    SetMotor(500)
    HeartTargetsDownCounter=HeartTargetsDownCounter+1
    HeartLight2.state=1
    Heart2.IsDropped = True
    heart2.collidable = False
    dropheart2.visible=0
    heartA2.transz=-12
    h2step = 0
    heart2.timerenabled=1
    CheckAllDropTargets
  End If
End Sub

Sub heart2_Timer
  Select Case h2step
        Case 3:heartA2.transz=-24
        Case 4:heartA2.transz=-36
    Case 5:heartA2.transz=-42:heart2.TimerEnabled = 0
    End Select
    h2Step = h2Step + 1
End Sub

Sub Heart3_Hit()
  If (Heart3.IsDropped = False) and (TableTilted=false) then
    PlaySoundAtVol SoundFXDOF("droptargetdropped",109,DOFPulse,DOFDropTargets), ActiveBall, 1
    SetMotor(500)
    HeartTargetsDownCounter=HeartTargetsDownCounter+1
    HeartLight3.state=1
    Heart3.IsDropped = True
    heart3.collidable = False
    dropheart3.visible=0
    heartA3.transz=-12
    h3step = 0
    heart3.timerenabled=1
    CheckAllDropTargets
  End If
End Sub

Sub heart3_Timer
  Select Case h3step
        Case 3:heartA3.transz=-24
        Case 4:heartA3.transz=-36
    Case 5:heartA3.transz=-42:heart3.TimerEnabled = 0
    End Select
    h3Step = h3Step + 1
End Sub

Sub Heart4_Hit()
  If (Heart4.IsDropped = False) and (TableTilted=false) then
    PlaySoundAtVol SoundFXDOF("droptargetdropped",109,DOFPulse,DOFDropTargets), ActiveBall, 1
    SetMotor(500)
    HeartTargetsDownCounter=HeartTargetsDownCounter+1
    HeartLight4.state=1
    Heart4.IsDropped = True
    heart4.collidable = False
    dropheart4.visible=0
    heartA4.transz=-12
    h4step = 0
    heart4.timerenabled=1
    CheckAllDropTargets
  End If
End Sub

Sub heart4_Timer
  Select Case h4step
        Case 3:heartA4.transz=-24
        Case 4:heartA4.transz=-36
    Case 5:heartA4.transz=-42:heart4.TimerEnabled = 0
    End Select
    h4Step = h4Step + 1
End Sub



'****************************************************

Sub CheckAllDropTargets

  HeartComplete=0
  ClubComplete=0
  DiamondComplete=0
  SpadeComplete=0
  RedSuits=0
  BlackSuits=0
  If DiamondTargetsDownCounter=4 then
    DiamondComplete=1
    RedSuits=RedSuits+1
    CenterLight1.state=1
    TopLeftLight.state=1
    RightLight2.state=1

  end if
  if ClubTargetsDownCounter=4 then
    ClubComplete=1
    BlackSuits=BlackSuits+1
    CenterLight4.state=1
    LeftInlaneLight.state=1
    RightLight1.state=1

  end if
  if SpadeTargetsDownCounter=4 then
    SpadeComplete=1
    BlackSuits=BlackSuits+1
    CenterLight2.state=1
    LeftLight2.state=1
    TopRightLight.state=1

  end if
  if HeartTargetsDownCounter=4 then
    HeartComplete=1
    RedSuits=RedSuits+1
    CenterLight3.state=1
    LeftLight1.state=1
    RightInlaneLight.state=1


  end if
  tempComplete=DiamondComplete+HeartComplete+ClubComplete+SpadeComplete

  If TempComplete=4 then
    TargetSequenceComplete=1
  end if

  CheckCenterTargetSpecials
end sub

Sub CheckCenterTargetSpecials

    If TargetSequenceComplete=1 then
      LeftOutlaneLight.state=1
      RightOutlaneLight.state=1
      Kick1Light.state=1
      kickcup.image="kickcuplit"
    else
      LeftOutlaneLight.state=0
      RightOutlaneLight.state=0
      Kick1Light.state=0
      kickcup.image="kickcup"
    end if

end sub

'****************************************************

sub Kicker1_Hit()
  Dim speedx,speedy,finalspeed
  speedx=activeball.velx
  speedy=activeball.vely
  finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  if TableTilted=true then
    Kicker1.Timerenabled=1
    exit sub
  end if
  if finalspeed>10 then
    Kicker1.Kick 0,0
    activeball.velx=speedx
    activeball.vely=speedy
  else

    KickerHolder1.enabled=1
    ScoreBonus
  end if

end sub

Sub KickerHolder1_timer
  If NewBonusTimer.enabled=0 then
    KickerHolder1.enabled=0
    Kicker1.timerenabled=1
    tKickerTCount=0

    if Kick1Light.state=1 then
      AddSpecial
    end if
  end if
end sub

Sub Kicker1_Timer()
  tKickerTCount=tKickerTCount+1
  select case tKickerTCount
  case 1:

    Pkickarm.rotz=15
    PlaysoundAtVol SoundFXDOF("saucer",113,DOFPulse,DOFContactors), Kicker1, 1
    DOF 114, DOFPulse
    Kicker1.kick 170,11
  case 2:
    Kicker1.timerenabled=0
    Pkickarm.rotz=0
  end Select
end sub



'**************************************
' Star Trigger
'**************************************

Sub StarTriggerDiamond_hit
  if TableTilted=false then
    if CenterLight1.state=1 then
      AddScore(100)
    else
      AddScore(10)
    end if
  end if
end sub

Sub StarTriggerSpade_hit
  if TableTilted=false then
    if CenterLight2.state=1 then
      AddScore(100)
    else
      AddScore(10)
    end if
  end if
end sub

Sub StarTriggerHeart_hit
  if TableTilted=false then
    if CenterLight3.state=1 then
      AddScore(100)
    else
      AddScore(10)
    end if
  end if
end sub

Sub StarTriggerClub_hit
  if TableTilted=false then
    if CenterLight4.state=1 then
      AddScore(100)
    else
      AddScore(10)
    end if
  end if
end sub

'**************************************


Sub AddSpecial()
  PlaySound SoundFXDOF("knocker",115,DOFPulse,DOFKnocker)
  DOF 114, DOFPulse
  BallInPlay=BallInPlay+1
  If BallInPlay>10 then
    BallInPlay=10
  end if
  If B2SOn Then
    For i=81 to 90
      Controller.B2SSetData i,0
    next
    Controller.B2SSetData 80+BallInPlay,1
  End If
  BallInPlayReel.SetValue(BallInPlay)
  for each obj in BallInPlayLights
    obj.state=0
  next
  BallInPlayLights(0).state=1
  BallInPlayLights(BallInPlay).state=1
End Sub

Sub AddSpecial2()
  PlaySound"click"
  Credits=Credits+1
  DOF 116, DOFOn
  if Credits>15 then Credits=15
  If B2SOn Then
    Controller.B2SSetCredits Credits
  End If
  CreditsReel.SetValue(Credits)
End Sub


Sub ToggleAlternatingRelay
  if NewBonusTimer.enabled=0 then
    If TargetSequenceComplete=1 then
      LeftOutlaneLight.state=1
      RightOutlaneLight.state=1
      Kick1Light.state=1
      kickcup.image="kickcuplit"
    else
      LeftOutlaneLight.state=0
      RightOutlaneLight.state=0
      Kick1Light.state=0
      kickcup.image="kickcup"
    end if

  end if
end sub

Sub ToggleRedBumper



  BumpersOn

end sub


Sub ResetBallDrops

  HoleCounter=0

End Sub


Sub LightsOut
  for each obj in Bonus
    obj.state=0
  next

  BonusCounter=0
  HoleCounter=0
  Bumper1State=0




end sub

Sub ResetBalls()

  TempMultiCounter=BallsPerGame-BallInPlay

  ResetBallDrops
  BonusMultiplier=1
  TableTilted=false
  TiltReel.SetValue(0)
  If B2Son then
    Controller.B2SSetTilt 0
  end if
  PlasticsOn
  'CreateBallID BallRelease
  Ballrelease.CreateSizedBall 25
    Ballrelease.Kick 40,7
  DOF 112, DOFPulse
  BallInPlayReel.SetValue(BallInPlay)
' InstructCard.image="IC"+FormatNumber(BallsPerGame,0)+FormatNumber(TargetSetting,0)+FormatNumber(BallInPlay,0)
' CardLight1.State = ABS(CardLight1.State-1)
  BallInPlayReel.SetValue(BallInPlay)
  for each obj in BallInPlayLights
    obj.state=0
  next
  BallInPlayLights(0).state=1
  BallInPlayLights(BallInPlay).state=1

End Sub

sub ResetDropsTimer_timer
    If NewBonusTimer.enabled = 1 then
      exit sub
    end if
    ResetDropsTimer.enabled=0
    ResetDropsTimer2.enabled=1
    ResetDrops
end sub

sub ResetDropsTimer2_timer

    ResetDropsTimer2.enabled=0
    ResetDrops2
end sub

Sub ResetDrops()

  for each obj in TopDropTargets
    obj.IsDropped=0
    obj.Collidable=1
  next
  PlaySound SoundFXDOF("dropsup2",110,DOFPulse,DOFContactors) ' TODO
  SpadeTargetsDownCounter=0
  DiamondTargetsDownCounter=0
  CenterLight1.state=0
  CenterLight2.state=0
  spadeA1.transz=-6
  spadeA2.transz=-6
  spadeA3.transz=-6
  spadeA4.transz=-6
  diamondA1.transz=-6
  diamondA2.transz=-6
  diamondA3.transz=-6
  diamondA4.transz=-6
  dropspade1.visible=1
  dropspade2.visible=1
  dropspade3.visible=1
  dropspade4.visible=1
  dropdiamond1.visible=1
  dropdiamond2.visible=1
  dropdiamond3.visible=1
  dropdiamond4.visible=1
End Sub

Sub ResetDrops2()

  for each obj in BottomDropTargets
    obj.IsDropped=0
    obj.Collidable=1
  next
  PlaySound SoundFXDOF("dropsup2",111,DOFPulse,DOFContactors) ' TODO
  ClubTargetsDownCounter=0
  HeartTargetsDownCounter=0
  CenterLight3.state=0
  CenterLight4.state=0
  clubA1.transz=-6
  clubA2.transz=-6
  clubA3.transz=-6
  clubA4.transz=-6
  heartA1.transz=-6
  heartA2.transz=-6
  heartA3.transz=-6
  heartA4.transz=-6
  dropclub1.visible=1
  dropclub2.visible=1
  dropclub3.visible=1
  dropclub4.visible=1
  dropheart1.visible=1
  dropheart2.visible=1
  dropheart3.visible=1
  dropheart4.visible=1
End Sub

sub ScoreBonus
  LHeart=1
  Spade=1
  Diamond=1
  Club=1
  LastRotation=1
  NewBonusTimer.enabled=true


end sub

sub OffRollovers
  TopLeftLight.state=0
  TopRightLight.state=0
  LeftLight1.state=0
  LeftLight2.state=0
  RightLight1.state=0
  RightLight2.state=0
  LeftInlaneLight.state=0
  RightInlaneLight.state=0
  LeftOutlaneLight.state=0
  RightOutlaneLight.state=0

end sub


sub NewBonusTimer_timer
  OffRollovers
  BumpersOff
  If Diamond<7 then
    if Diamond<5 then
      if EVAL("DiamondLight"&Diamond).state=1 then
        if CenterLight1.state=1 then
          AddScore(1000)
        else
          AddScore(100)
        end if
      else
        PlaySound"MotorPause"
      end if
    end if
    if Diamond=5 then
      PlaySound"MotorPause"
    end if
    if Diamond=6 then
      PlaySound"HighHandEndCycle"
      CheckAllDropTargets
      BumpersOn
    end if
    Diamond=Diamond+1
    exit sub
  end if

  If Spade<7 then
    if Spade<5 then
      if EVAL("SpadeLight"&Spade).state=1 then
        if CenterLight2.state=1 then
          AddScore(1000)
        else
          AddScore(100)
        end if
      else
        PlaySound"MotorPause"
      end if
    end if
    if Spade=5 then
      PlaySound"MotorPause"
    end if
    if Spade=6 then
      PlaySound"HighHandEndCycle"
      CheckAllDropTargets
      BumpersOn
    end if
    Spade=Spade+1
    exit sub
  end if

  If LHeart<7 then
    if LHeart<5 then
      if EVAL("HeartLight"&LHeart).state=1 then
        if CenterLight3.state=1 then
          AddScore(1000)
        else
          AddScore(100)
        end if
      else
        PlaySound"MotorPause"
      end if
    end if
    if LHeart=5 then
      PlaySound"MotorPause"

    end if
    if LHeart=6 then
      PlaySound"HighHandEndCycle"
      CheckAllDropTargets
      BumpersOn
    end if
    LHeart=LHeart+1
    exit sub
  end if

  If Club<7 then
    if Club<5 then
      if EVAL("ClubLight"&Club).state=1 then
        if CenterLight4.state=1 then
          AddScore(1000)
        else
          AddScore(100)
        end if
      else
        PlaySound"MotorPause"
      end if
    end if
    if Club=5 then
      PlaySound"MotorPause"
    end if
    if Club=6 then
      PlaySound"HighHandEndCycle"
      CheckAllDropTargets
      BumpersOn
    end if
    Club=Club+1
    exit sub
  end if

  If LastRotation<7 then
    if LastRotation<6 then
      PlaySound"MotorPause"
    end if
    if LastRotation=6 then
      PlaySound"HighHandEndCycle"
      CheckAllDropTargets
      BumpersOn
    end if
    LastRotation=LastRotation+1
    exit sub
  end if
  NewBonusTimer.enabled=0
  CheckAllDropTargets
  BumpersOn

end sub

sub resettimer_timer
    rst=rst+1
  if B2SOn then
    Controller.B2SSetScorePlayer 1, 0
  end if

  if rst<=BallsPerGame then
    BallInPlayReel.SetValue(rst)
    for each obj in BallInPlayLights
      obj.state=0
    next
    BallInPlayLights(0).state=1
    BallInPlayLights(rst).state=1
    If B2SOn Then
      For i=81 to 90
        Controller.B2SSetData i,0
      next
      Controller.B2SSetData 80+rst,1
    End If
  end if
    if rst=10 then
    playsound "StartBall1"
    end if
    if rst=12 then
    newgame
    resettimer.enabled=false
    end if
end sub



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
    Controller.B2SSetData 60,0
    Controller.B2SSetData 61,0
    Controller.B2SSetData 62,0
    Controller.B2SSetData 63,0
  End if
  for each obj in NumberLights
    obj.state=0
  next

  for each obj in LightSpecial
    obj.state=0
  next
  For each obj in ResetLights
    obj.state=0
  next
  AlternatingRelay=0
  ZeroNineCounter=0
  Bumper1State=1
  BumpersOn
  BonusCounter=0
  BallCounter=0
  TargetLeftFlag=1
  TargetCenterFlag=1
  TargetRightFlag=1
  TargetSequenceComplete=0
  For each obj in HUDCards
    obj.SetValue(0)
  next
  TopLeftLight.state=0
  TopRightLight.state=0
  LeftLight1.state=0
  LeftLight2.state=0
  RightLight1.state=0
  RightLight2.state=0
  LeftInlaneLight.state=0
  RightInlaneLight.state=0


' IncreaseBonus
' ToggleBumper
  EightLit=1
  ResetBalls
End sub

sub nextball

  Player=Player+1
  If Player>Players Then
    BallInPlay=BallInPlay-1
    If BallInPlay< 1 then
      PlaySound("MotorLeer")
      InProgress=false

      If B2SOn Then
        Controller.B2SSetGameOver 1
        Controller.B2SSetPlayerUp 0
        Controller.B2SSetBallInPlay 0
        Controller.B2SSetCanPlay 0
        For i=80 to 90
          Controller.B2SSetData i,0
        next
      End If
      For each obj in PlayerHuds
        obj.SetValue(0)
      next
      GameOverReel.SetValue(1)
'     InstructCard.image="IC"+FormatNumber(BallsPerGame,0)+FormatNumber(TargetSetting,0)+"0"
'     CardLight1.State = ABS(CardLight1.State-1)
      BallInPlayReel.SetValue(0)
      for each obj in BallInPlayLights
        obj.state=0
      next
'     CanPlayReel.SetValue(0)
      LeftFlipper.RotateToStart
      RightFlipper.RotateToStart
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
      TimerTitle4.enabled=1
      HighScoreTimer.interval=100
      HighScoreTimer.enabled=True
    Else
      Player=1
      If B2SOn Then
        Controller.B2SSetPlayerUp Player
        Controller.B2SSetBallInPlay BallInPlay
        For i=81 to 90
          Controller.B2SSetData i,0
        next
        Controller.B2SSetData 80+BallInPlay,1
      End If
'     PlaySound("RotateThruPlayers")
      TempPlayerUp=Player
'     PlayerUpRotator.enabled=true
      PlayStartBall.enabled=true
      For each obj in PlayerHuds
        obj.SetValue(0)
      next
      PlayerHuds(Player-1).SetValue(1)

      ResetBalls
    End If
  Else
    If B2SOn Then
      Controller.B2SSetPlayerUp Player
      Controller.B2SSetBallInPlay BallInPlay
      For i=81 to 90
        Controller.B2SSetData i,0
      next
      Controller.B2SSetData 80+BallInPlay,1
    End If
'   PlaySound("RotateThruPlayers")
    TempPlayerUp=Player
'   PlayerUpRotator.enabled=true
    PlayStartBall.enabled=true
    For each obj in PlayerHuds
      obj.SetValue(0)
    next
    PlayerHuds(Player-1).SetValue(1)
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
  exit sub
  if TableTilted=true and TiltEndsGame=1 then
    exit sub
  end if
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
    If TableTilted=true then
      exit sub
    end if
    TiltCount = TiltCount + 1
    if TiltCount = 3 then
      TableTilted=True
      TiltReel.SetValue(1)
      If TiltEndsGame=1 then
        If BallInPlay > 1 then
          BallInPlay = BallInPlay - 1
          'BallsToPlayReel.SetValue(BallInPlay)
          for each obj in BallInPlayLights
            obj.state=0
          next
          BallInPlayLights(0).state=1
          BallInPlayLights(BallInPlay).state=1
          if B2SOn then
            For i=81 to 90
              Controller.B2SSetData i,0
            next
            Controller.B2SSetData 80+BallInPlay,1
          end if
        end if
      end if
      PlasticsOff
      BumpersOff
      LeftFlipper.RotateToStart
      RightFlipper.RotateToStart
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
  PlaySound("ballrelease")
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
    scorefile.writeline TiltEndsGame
    scorefile.writeline TargetSetting
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
    temp16=textstr.readline
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
    TiltEndsGame=cdbl(temp4)
    TargetSetting=cdbl(temp5)
    ReplayLevel=cdbl(temp16)
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
    If B2SOn Then
      Controller.B2SSetScore 3,HighScore
    End If
    DisplayHighScore
    CreditsReel.SetValue(Credits)
    InitPauser5.enabled=false
end sub



sub BumpersOff
  Bumper1Light.state=0

  TopLeftLight.visible=0
  TopRightLight.visible=0
  LeftLight1.visible=0
  LeftLight2.visible=0
  RightLight1.visible=0
  RightLight2.visible=0
  LeftInlaneLight.visible=0
  RightInlaneLight.visible=0
  LeftOutlaneLight.visible=0
  RightOutlaneLight.visible=0
end sub

sub BumpersOn


  Bumper1Light.state=0

  If Bumper1State=1 then

    Bumper1Light.state=1
  end if

  TopLeftLight.visible=1
  TopRightLight.visible=1
  LeftLight1.visible=1
  LeftLight2.visible=1
  RightLight1.visible=1
  RightLight2.visible=1
  LeftInlaneLight.visible=1
  RightInlaneLight.visible=1
  LeftOutlaneLight.visible=1
  RightOutlaneLight.visible=1

end sub


Sub PlasticsOn
end sub

Sub PlasticsOff
  StopSound "buzz"
  StopSound "buzzL"
end sub

Sub SetupReplayTables

  Replay1Table(1)=40000
  Replay1Table(2)=60000
  Replay1Table(3)=50000
  Replay1Table(4)=53000
  Replay1Table(5)=59000
  Replay1Table(6)=64000
  Replay1Table(7)=68000
  Replay1Table(8)=58000
  Replay1Table(9)=61000
  Replay1Table(10)=65000
  Replay1Table(11)=68000
  Replay1Table(12)=75000
  Replay1Table(13)=80000
  Replay1Table(14)=85000
  Replay1Table(15)=999000

  Replay2Table(1)=100000
  Replay2Table(2)=120000
  Replay2Table(3)=68000
  Replay2Table(4)=71000
  Replay2Table(5)=77000
  Replay2Table(6)=82000
  Replay2Table(7)=86000
  Replay2Table(8)=76000
  Replay2Table(9)=79000
  Replay2Table(10)=83000
  Replay2Table(11)=86000
  Replay2Table(12)=93000
  Replay2Table(13)=98000
  Replay2Table(14)=99000
  Replay2Table(15)=999000

  Replay3Table(1)=160000
  Replay3Table(2)=180000
  Replay3Table(3)=81000
  Replay3Table(4)=84000
  Replay3Table(5)=90000
  Replay3Table(6)=95000
  Replay3Table(7)=99000
  Replay3Table(8)=999000
  Replay3Table(9)=999000
  Replay3Table(10)=999000
  Replay3Table(11)=999000
  Replay3Table(12)=999000
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

  ReplayTableMax=2


end sub

'Sub RefreshReplayCard
' Dim tempst1
' Dim tempst2
'
' tempst1=FormatNumber(BallsPerGame,0)
' tempst2=FormatNumber(ReplayLevel,0)
'
'
' ReplayCard.image = "SC" + tempst2
' Replay1=Replay1Table(ReplayLevel)
' Replay2=Replay2Table(ReplayLevel)
' Replay3=Replay3Table(ReplayLevel)
' Replay4=Replay4Table(ReplayLevel)
'
'end sub


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
        OffRollovers
        BumpersOff
      Case 100:
        MotorMode=100
        MotorPosition=1

      Case 500:
        MotorMode=100
        MotorPosition=5
        OffRollovers
        BumpersOff
      Case 1000:
        MotorMode=1000
        MotorPosition=1
      Case 2000:
        MotorMode=1000
        MotorPosition=2
        OffRollovers
        BumpersOff
      Case 3000:
        MotorMode=1000
        MotorPosition=3
        OffRollovers
        BumpersOff
      Case 4000:
        MotorMode=1000
        MotorPosition=4
        OffRollovers
        BumpersOff
      Case 5000:
        MotorMode=1000
        MotorPosition=5
        OffRollovers
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
        MotorPosition=0:MotorRunning=0:CheckAllDropTargets:BumpersOn
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
  If ScoreDisplay(Player)<100000 then
    ScoreDisplay(Player)=Score(Player)
  Else
    Score100K(Player)=Int(Score(Player)/100000)
    ScoreDisplay(Player)=Score(Player)-100000
  End If
  if Score(Player)=>100000 then
    HUD100K.SetValue(1)
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
  If NewScore >= 100000 then
    HUD100K.SetValue(1)
    If B2SOn Then
      Controller.B2SSetScoreRolloverPlayer1 1
    End If
  End If
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
    Controller.B2SSetScorePlayer 1, Score(Player)
  End If
' EMReel1.SetValue Score(Player)
  PlayerScores(Player-1).AddValue(x)
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
' Positional Sound Playback Functions by DJRobX, Rothbauerw and Herweh
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


Dim BotPos, Pos

'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 5 ' total number of balls
ReDim rolling(tnob)
InitRolling

ReDim ArchRolling(tnob)
InitArchRolling

Dim ArchHit

Sub LowerArch_Hit
  Archhit = 1
End Sub

Sub NotOnArch_Hit
  ArchHit = 0
End Sub

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub InitArchRolling
  Dim i
  For i = 0 to tnob
    ArchRolling(i) = False
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
        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If

    If BallVel(BOT(b) ) > 1 AND ArchHit =1 Then
      ArchRolling(b) = True
      PlaySound("ArchHitA" & b),   0, (BallVel(BOT(b))/15)^5 * 1, AudioPan(BOT(b)), 0, (BallVel(BOT(b))/40)^7, 1, 0, 0  'Left & Right stereo or Top & Bottom stereo PF Speakers.
      PlaySound("ArchRollA" & b), -1, (BallVel(BOT(b))/30)^5 * 1, AudioPan(BOT(b)), 0, (BallVel(BOT(b))/40)^7, 1, 0, 0  'Left & Right stereo or Top & Bottom stereo PF Speakers.
    Else
      If ArchRolling(b) = True Then
      StopSound("ArchRollA" & b)
      ArchRolling(b) = False
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
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) + 13
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) - 13
        End If
        ballShadow(b).Y = BOT(b).Y + 10
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
End Sub

'************************************
' What you need to add to your table
'************************************

' a timer called RollingTimer. With a fast interval, like 10
' one collision sound, in this script is called fx_collide
' as many sound files as max number of balls, with names ending with 0, 1, 2, 3, etc
' for ex. as used in this script: fx_ballrolling0, fx_ballrolling1, fx_ballrolling2, fx_ballrolling3, etc


'******************************************
' Explanation of the rolling sound routine
'******************************************

' sounds are played based on the ball speed and position

' the routine checks first for deleted balls and stops the rolling sound.

' The For loop goes through all the balls on the table and checks for the ball speed and
' if the ball is on the table (height lower than 30) then then it plays the sound
' otherwise the sound is stopped, like when the ball has stopped or is on a ramp or flying.

' The sound is played using the VOL, PAN and PITCH functions, so the volume and pitch of the sound
' will change according to the ball speed, and the PAN function will change the stereo position according
' to the position of the ball on the table.


'**************************************
' Explanation of the collision routine
'**************************************

' The collision is built in VP.
' You only need to add a Sub OnBallBallCollision(ball1, ball2, velocity) and when two balls collide they
' will call this routine. What you add in the sub is up to you. As an example is a simple Playsound with volume and paning
' depending of the speed of the collision.

Sub PlayFieldSound (SoundName, Looper, TableObject)
  PlaySound SoundName, Looper, 1, AudioPan(TableObject), 0, 0, 0, 0, AudioFade(TableObject)
End Sub

Sub a_Pins_Hit (idx)
  PlaySound "pinhit_low", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub a_Targets_Hit (idx)
  PlaySound "target", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub a_Metals_Thin_Hit (idx)
  PlaySound "metalhit_thin", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub a_Metals_Medium_Hit (idx)
  PlaySound "metalhit_medium", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub a_Metals2_Hit (idx)
  PlaySound "metalhit2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub a_Gates_Hit (idx)
  PlaySound "gate4", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub a_Spinner_Spin
  PlaySound "fx_spinner", 0, .25, AudioPan(Spinner), 0.25, 0, 0, 1, AudioFade(Spinner)
End Sub

Sub a_Rubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 20 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End Sub

Sub a_Posts_Hit(idx)
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

Sub RubberWheel_hit
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 20 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End sub

Sub ApronL_Hit
  PlaySoundAt "ApronHit", ApronL
End Sub

Sub ApronR_Hit
  PlaySoundAt"ApronHit", ApronR
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
    HighScoreReel.SetValue(HSScore(1))
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
    HighScoreReel.SetValue(HSScore(1))
    If B2SOn then
      Controller.B2SSetScorePlayer 2,HSScore(1)
    end if
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
      PlaySound SoundFXDOF("knocker",115,DOFPulse,DOFKnocker)
      DOF 114, DOFPulse
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
Const OptionLinesToMark="111100011"
Const OptionLine1="" 'do not use this line
Const OptionLine2="" 'do not use this line
Const OptionLine3="" 'do not use this line
Const OptionLine4="Tilt Penalty"
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
          tempstring = tempstring + FormatNumber(Replay1Table(ReplayLevel),0) + "/" + FormatNumber(Replay2Table(ReplayLevel),0)
        else
          tempstring = tempstring + FormatNumber(Replay1Table(ReplayLevel),0) + "/" + FormatNumber(Replay2Table(ReplayLevel),0) + "/" + FormatNumber(Replay3Table(ReplayLevel),0)
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
        If TiltEndsGame=0 then
          tempstring="Ball in play only"
        else
          tempstring="Ball in play plus one ball"
        end if
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
      if TiltEndsGame=0 then
        TiltEndsGame=1
      else
        TiltEndsGame=0
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
'       RefreshReplayCard
'       InstructCard.image="IC"+FormatNumber(BallsPerGame,0)+FormatNumber(TiltEndsGame,0)
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

'******************************************************
'   FLIPPER CORRECTION SUPPORTING FUNCTIONS
'******************************************************

dim LF : Set LF = New FlipperPolarity
dim RF : Set RF = New FlipperPolarity

InitPolarity

Sub InitPolarity()
  dim x, a : a = Array(LF, RF)
  for each x in a
    'safety coefficient (diminishes polarity correction only)
    x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
    x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1

    x.enabled = True
    x.TimeDelay = 44
  Next

  '"Polarity" Profile
  AddPt "Polarity", 0, 0, 0
  AddPt "Polarity", 1, 0.368, -3
  AddPt "Polarity", 2, 0.50, -2.5
  AddPt "Polarity", 3, 0.65, -1.3
  AddPt "Polarity", 4, 0.71, -1
  AddPt "Polarity", 5, 0.785,-0.8
  AddPt "Polarity", 6, 1.18, -0
  AddPt "Polarity", 7, 1.2, 0


  '"Velocity" Profile
  addpt "Velocity", 0, 0,   1
  addpt "Velocity", 1, 0.16, 1.06
  addpt "Velocity", 2, 0.41,  1.05
  addpt "Velocity", 3, 0.53,  1'0.982
  addpt "Velocity", 4, 0.702, 0.968
  addpt "Velocity", 5, 0.95,  0.968
  addpt "Velocity", 6, 1.03,  0.945

  LF.Object = LeftFlipper
  LF.EndPoint = EndPointLp  'you can use just a coordinate, or an object with a .x property. Using a couple of simple primitive objects
  RF.Object = RightFlipper
  RF.EndPoint = EndPointRp
End Sub

'Trigger Hit - .AddBall activeball
'Trigger UnHit - .PolarityCorrect activeball

Sub TriggerLF_Hit() : LF.Addball activeball : End Sub
Sub TriggerLF_UnHit() : LF.PolarityCorrect activeball : End Sub
Sub TriggerRF_Hit() : RF.Addball activeball : End Sub
Sub TriggerRF_UnHit() : RF.PolarityCorrect activeball : End Sub

Sub AddPt(aStr, idx, aX, aY)  'debugger wrapper for adjusting flipper script in-game
  dim a : a = Array(LF, RF)
  dim x : for each x in a
    x.addpoint aStr, idx, aX, aY
  Next
End Sub

'Methods:
'.TimeDelay - Delay before trigger shuts off automatically. Default = 80 (ms)
'.AddPoint - "Polarity", "Velocity", "Ycoef" coordinate points. Use one of these 3 strings, keep coordinates sequential. x = %position on the flipper, y = output
'.Object - set to flipper reference. Optional.
'.StartPoint - set start point coord. Unnecessary, if .object is used.

'Called with flipper -
'ProcessBalls - catches ball data.
' - OR -
'.Fire - fires flipper.rotatetoend automatically + processballs. Requires .Object to be set to the flipper.

Class FlipperPolarity
  Public DebugOn, Enabled
  Private FlipAt  'Timer variable (IE 'flip at 723,530ms...)
  Public TimeDelay  'delay before trigger turns off and polarity is disabled TODO set time!
  private Flipper, FlipperStart, FlipperEnd, LR, PartialFlipCoef
  Private Balls(20), balldata(20)

  dim PolarityIn, PolarityOut
  dim VelocityIn, VelocityOut
  dim YcoefIn, YcoefOut
  Public Sub Class_Initialize
    redim PolarityIn(0) : redim PolarityOut(0) : redim VelocityIn(0) : redim VelocityOut(0) : redim YcoefIn(0) : redim YcoefOut(0)
    Enabled = True : TimeDelay = 50 : LR = 1:  dim x : for x = 0 to uBound(balls) : balls(x) = Empty : set Balldata(x) = new SpoofBall : next
  End Sub

  Public Property let Object(aInput) : Set Flipper = aInput : StartPoint = Flipper.x : End Property
  Public Property Let StartPoint(aInput) : if IsObject(aInput) then FlipperStart = aInput.x else FlipperStart = aInput : end if : End Property
  Public Property Get StartPoint : StartPoint = FlipperStart : End Property
  Public Property Let EndPoint(aInput) : if IsObject(aInput) then FlipperEnd = aInput.x else FlipperEnd = aInput : end if : End Property
  Public Property Get EndPoint : EndPoint = FlipperEnd : End Property

  Public Sub AddPoint(aChooseArray, aIDX, aX, aY) 'Index #, X position, (in) y Position (out)
    Select Case aChooseArray
      case "Polarity" : ShuffleArrays PolarityIn, PolarityOut, 1 : PolarityIn(aIDX) = aX : PolarityOut(aIDX) = aY : ShuffleArrays PolarityIn, PolarityOut, 0
      Case "Velocity" : ShuffleArrays VelocityIn, VelocityOut, 1 :VelocityIn(aIDX) = aX : VelocityOut(aIDX) = aY : ShuffleArrays VelocityIn, VelocityOut, 0
      Case "Ycoef" : ShuffleArrays YcoefIn, YcoefOut, 1 :YcoefIn(aIDX) = aX : YcoefOut(aIDX) = aY : ShuffleArrays YcoefIn, YcoefOut, 0
    End Select
    if gametime > 100 then Report aChooseArray
  End Sub

  Public Sub Report(aChooseArray)   'debug, reports all coords in tbPL.text
    if not DebugOn then exit sub
    dim a1, a2 : Select Case aChooseArray
      case "Polarity" : a1 = PolarityIn : a2 = PolarityOut
      Case "Velocity" : a1 = VelocityIn : a2 = VelocityOut
      Case "Ycoef" : a1 = YcoefIn : a2 = YcoefOut
      case else :tbpl.text = "wrong string" : exit sub
    End Select
    dim str, x : for x = 0 to uBound(a1) : str = str & aChooseArray & " x: " & round(a1(x),4) & ", " & round(a2(x),4) & vbnewline : next
    tbpl.text = str
  End Sub

  Public Sub AddBall(aBall) : dim x : for x = 0 to uBound(balls) : if IsEmpty(balls(x)) then set balls(x) = aBall : exit sub :end if : Next  : End Sub

  Private Sub RemoveBall(aBall)
    dim x : for x = 0 to uBound(balls)
      if TypeName(balls(x) ) = "IBall" then
        if aBall.ID = Balls(x).ID Then
          balls(x) = Empty
          Balldata(x).Reset
        End If
      End If
    Next
  End Sub

  Public Sub Fire()
    Flipper.RotateToEnd
    processballs
  End Sub

  Public Property Get Pos 'returns % position a ball. For debug stuff.
    dim x : for x = 0 to uBound(balls)
      if not IsEmpty(balls(x) ) then
        pos = pSlope(Balls(x).x, FlipperStart, 0, FlipperEnd, 1)
      End If
    Next
  End Property

  Public Sub ProcessBalls() 'save data of balls in flipper range
    FlipAt = GameTime
    dim x : for x = 0 to uBound(balls)
      if not IsEmpty(balls(x) ) then
        balldata(x).Data = balls(x)
        if DebugOn then StickL.visible = True : StickL.x = balldata(x).x    'debug TODO
      End If
    Next
    PartialFlipCoef = ((Flipper.StartAngle - Flipper.CurrentAngle) / (Flipper.StartAngle - Flipper.EndAngle))
    PartialFlipCoef = abs(PartialFlipCoef-1)
    if abs(Flipper.currentAngle - Flipper.EndAngle) < 30 Then
      PartialFlipCoef = 0
    End If
  End Sub
  Private Function FlipperOn() : if gameTime < FlipAt+TimeDelay then FlipperOn = True : End If : End Function 'Timer shutoff for polaritycorrect

  Public Sub PolarityCorrect(aBall)
    if FlipperOn() then
      dim tmp, BallPos, x, IDX, Ycoef : Ycoef = 1
      dim teststr : teststr = "Cutoff"
      tmp = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
      if tmp < 0.1 then 'if real ball position is behind flipper, exit Sub to prevent stucks  'Disabled 1.03, I think it's the Mesh that's causing stucks, not this
        if DebugOn then TestStr = "real pos < 0.1 ( " & round(tmp,2) & ")" : tbpl.text = Teststr
        'RemoveBall aBall
        'Exit Sub
      end if

      'y safety Exit
      if aBall.VelY > -8 then 'ball going down
        if DebugOn then teststr = "y velocity: " & round(aBall.vely, 3) & "exit sub" : tbpl.text = teststr
        RemoveBall aBall
        exit Sub
      end if
      'Find balldata. BallPos = % on Flipper
      for x = 0 to uBound(Balls)
        if aBall.id = BallData(x).id AND not isempty(BallData(x).id) then
          idx = x
          BallPos = PSlope(BallData(x).x, FlipperStart, 0, FlipperEnd, 1)
          'TB.TEXT = balldata(x).id & " " & BALLDATA(X).X & VBNEWLINE & FLIPPERSTART & " " & FLIPPEREND
          if ballpos > 0.65 then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)        'find safety coefficient 'ycoef' data
        end if
      Next

      'Velocity correction
      if not IsEmpty(VelocityIn(0) ) then
        Dim VelCoef
        if DebugOn then set tmp = new spoofball : tmp.data = aBall : End If
        if IsEmpty(BallData(idx).id) and aBall.VelY < -12 then 'if tip hit with no collected data, do vel correction anyway
          if PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1) > 1.1 then 'adjust plz
            VelCoef = LinearEnvelope(5, VelocityIn, VelocityOut)
            if partialflipcoef < 1 then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)
            if Enabled then aBall.Velx = aBall.Velx*VelCoef'VelCoef
            if Enabled then aBall.Vely = aBall.Vely*VelCoef'VelCoef
            if DebugOn then teststr = "tip protection" & vbnewline & "velcoef: " & round(velcoef,3) & vbnewline & round(PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1),3) & vbnewline
            'debug.print teststr
          end if
        Else
     :      VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)
          if Enabled then aBall.Velx = aBall.Velx*VelCoef
          if Enabled then aBall.Vely = aBall.Vely*VelCoef
        end if
      End If

      'Polarity Correction (optional now)
      if not IsEmpty(PolarityIn(0) ) then
        If StartPoint > EndPoint then LR = -1 'Reverse polarity if left flipper
        dim AddX : AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR
        if Enabled then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef)
      End If
      'debug
      if DebugOn then
        TestStr = teststr & "%pos:" & round(BallPos,2)
        if IsEmpty(PolarityOut(0) ) then
          teststr = teststr & vbnewline & "(Polarity Disabled)" & vbnewline
        else
          teststr = teststr & "+" & round(1 *(AddX*ycoef*PartialFlipcoef),3)
          if BallPos >= PolarityOut(uBound(PolarityOut) ) then teststr = teststr & "(MAX)" & vbnewline else teststr = teststr & vbnewline end if
          if Ycoef < 1 then teststr = teststr &  "ycoef: " & ycoef & vbnewline
          if PartialFlipcoef < 1 then teststr = teststr & "PartialFlipcoef: " & round(PartialFlipcoef,4) & vbnewline
        end if

        teststr = teststr & vbnewline & "Vel: " & round(BallSpeed(tmp),2) & " -> " & round(ballspeed(aBall),2) & vbnewline
        teststr = teststr & "%" & round(ballspeed(aBall) / BallSpeed(tmp),2)
        tbpl.text = TestSTR
      end if
    Else
      'if DebugOn then tbpl.text = "td" & timedelay
    End If
    RemoveBall aBall
  End Sub
End Class

'================================
'Helper Functions


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
  for x = 0 to aCount-1   'set objects back into original array
    if IsObject(a(x)) then
      Set aArray(x) = a(x)
    Else
      aArray(x) = a(x)
    End If
  Next
End Sub

Sub ShuffleArrays(aArray1, aArray2, offset)
  ShuffleArray aArray1, offset
  ShuffleArray aArray2, offset
End Sub


Function BallSpeed(ball) 'Calculates the ball speed
    BallSpeed = SQR(ball.VelX^2 + ball.VelY^2 + ball.VelZ^2)
End Function

Function PSlope(Input, X1, Y1, X2, Y2)  'Set up line via two points, no clamping. Input X, output Y
  dim x, y, b, m : x = input : m = (Y2 - Y1) / (X2 - X1) : b = Y2 - m*X2
  Y = M*x+b
  PSlope = Y
End Function

Function NullFunctionZ(aEnabled):End Function '1 argument null function placeholder  TODO move me or replac eme

Class spoofball
  Public X, Y, Z, VelX, VelY, VelZ, ID, Mass, Radius
  Public Property Let Data(aBall)
    With aBall
      x = .x : y = .y : z = .z : velx = .velx : vely = .vely : velz = .velz
      id = .ID : mass = .mass : radius = .radius
    end with
  End Property
  Public Sub Reset()
    x = Empty : y = Empty : z = Empty  : velx = Empty : vely = Empty : velz = Empty
    id = Empty : mass = Empty : radius = Empty
  End Sub
End Class


Function LinearEnvelope(xInput, xKeyFrame, yLvl)
  dim y 'Y output
  dim L 'Line
  dim ii : for ii = 1 to uBound(xKeyFrame)  'find active line
    if xInput <= xKeyFrame(ii) then L = ii : exit for : end if
  Next
  if xInput > xKeyFrame(uBound(xKeyFrame) ) then L = uBound(xKeyFrame)  'catch line overrun
  Y = pSlope(xInput, xKeyFrame(L-1), yLvl(L-1), xKeyFrame(L), yLvl(L) )

  'Clamp if on the boundry lines
  'if L=1 and Y < yLvl(LBound(yLvl) ) then Y = yLvl(lBound(yLvl) )
  'if L=uBound(xKeyFrame) and Y > yLvl(uBound(yLvl) ) then Y = yLvl(uBound(yLvl) )
  'clamp 2.0
  if xInput <= xKeyFrame(lBound(xKeyFrame) ) then Y = yLvl(lBound(xKeyFrame) )  'Clamp lower
  if xInput >= xKeyFrame(uBound(xKeyFrame) ) then Y = yLvl(uBound(xKeyFrame) )  'Clamp upper

  LinearEnvelope = Y
End Function

'****************************************************************************
'PHYSICS DAMPENERS

'These are data mined bounce curves,
'dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
'Requires tracking ballspeed to calculate COR


Sub dPosts_Hit(idx)
  RubbersD.dampen Activeball
End Sub

Sub dSleeves_Hit(idx)
  SleevesD.Dampen Activeball
End Sub


dim RubbersD : Set RubbersD = new Dampener  'frubber
RubbersD.name = "Rubbers"
RubbersD.debugOn = False  'shows info in textbox "TBPout"
RubbersD.Print = False  'debug, reports in debugger (in vel, out cor)
'cor bounce curve (linear)
'for best results, try to match in-game velocity as closely as possible to the desired curve
'RubbersD.addpoint 0, 0, 0.935  'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 0, 0, 0.96  'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 1, 3.77, 0.96
RubbersD.addpoint 2, 5.76, 0.967  'dont take this as gospel. if you can data mine rubber elasticitiy, please help!
RubbersD.addpoint 3, 15.84, 0.874
RubbersD.addpoint 4, 56, 0.64 'there's clamping so interpolate up to 56 at least

dim SleevesD : Set SleevesD = new Dampener  'this is just rubber but cut down to 85%...
SleevesD.name = "Sleeves"
SleevesD.debugOn = False  'shows info in textbox "TBPout"
SleevesD.Print = False  'debug, reports in debugger (in vel, out cor)
SleevesD.CopyCoef RubbersD, 0.85

Class Dampener
  Public Print, debugOn 'tbpOut.text
  public name, Threshold  'Minimum threshold. Useful for Flippers, which don't have a hit threshold.
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

' Thalamus - patched : ' Thalamus - patched :     aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
    aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
' Thalamus - patched :     aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
    aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
    if debugOn then TBPout.text = str
  End Sub

  Public Sub CopyCoef(aObj, aCoef) 'alternative addpoints, copy with coef
    dim x : for x = 0 to uBound(aObj.ModIn)
      addpoint x, aObj.ModIn(x), aObj.ModOut(x)*aCoef
    Next
  End Sub


  Public Sub Report()   'debug, reports all coords in tbPL.text
    if not debugOn then exit sub
    dim a1, a2 : a1 = ModIn : a2 = ModOut
    dim str, x : for x = 0 to uBound(a1) : str = str & x & ": " & round(a1(x),4) & ", " & round(a2(x),4) & vbnewline : next
    TBPout.text = str
  End Sub


End Class

'Tracks ball velocity for judging bounce calculations & angle
'apologies to JimmyFingers is this is what his script does. I know his tracks ball velocity too but idk how it works in particular
dim cor : set cor = New CoRTracker
cor.debugOn = False
'cor.update() - put this on a low interval timer
Class CoRTracker
  public DebugOn 'tbpIn.text
  public ballvel

  Private Sub Class_Initialize : redim ballvel(0) : End Sub
  'TODO this would be better if it didn't do the sorting every ms, but instead every time it's pulled for COR stuff
  Public Sub Update() 'tracks in-ball-velocity
    dim str, b, AllBalls, highestID : allBalls = getballs
    if uBound(allballs) < 0 then if DebugOn then str = "no balls" : TBPin.text = str : exit Sub else exit sub end if: end if
    for each b in allballs
      if b.id >= HighestID then highestID = b.id
    Next

    if uBound(ballvel) < highestID then redim ballvel(highestID)  'set bounds

    for each b in allballs
      ballvel(b.id) = BallSpeed(b)
      if DebugOn then
        dim s, bs 'debug spacer, ballspeed
        bs = round(BallSpeed(b),1)
        if bs < 10 then s = " " else s = "" end if
        str = str & b.id & ": " & s & bs & vbnewline
        'str = str & b.id & ": " & s & bs & "z:" & b.z & vbnewline
      end if
    Next
    if DebugOn then str = "ubound ballvels: " & ubound(ballvel) & vbnewline & str : if TBPin.text <> str then TBPin.text = str : end if
  End Sub
End Class

Sub RDampen_Timer()
Cor.Update
End Sub

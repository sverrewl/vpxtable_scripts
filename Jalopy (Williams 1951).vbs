'*
'*        Williams' Jalopy (1951)
'*        VPX Table scripted by Loserman76
'*        Artwork by webby
'*
'*

option explicit
Randomize
ExecuteGlobal GetTextFile("core.vbs")

On Error Resume Next
ExecuteGlobal GetTextFile("Controller.vbs")
If Err Then MsgBox "Unable to open Controller.vbs. Ensure that it is in the scripts folder."
On Error Goto 0

Const cGameName = "Jalopy_1951"


Const ShadowFlippersOn = true
Const ShadowBallOn = false

Const ShadowConfigFile = false




Dim Controller  ' B2S
Dim B2SScore  ' B2S Score Displayed
Const HSFileName="Jalopy_51VPX.txt"
Const B2STableName="Jalopy_1951"
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
Dim bump6

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

Dim GreenBumpersOn, YellowBumpersOn, ArrowCounter, HorsedSequenceComplete, MatchLit2

Dim WhiteComplete,RedComplete,YellowComplete,GreenComplete,DisableKeysInit,BallLiftOption,BallsPlayed,BallsDrained

Dim SpinPos,Spin,Count,Count1,Count2,Count3,Reset,VelX,VelY,BallSpeed,LitSpinner,BallIndicatorCount,CenterTargetTracker,BallsOnTable,TiltEndsGame, WildArray, KickerHoldFlag

Dim tKickerTCount1,tKickerTCount2,tKickerTCount3,tKickerTCount4
Dim ArrowCounter1,ArrowCounter2,ArrowCounter3,ArrowCounter4
Dim HorsedCounter,Player1Tilted,Player2Tilted,NewPoints,JokerCount,Points,PlusPoints,CenterLightCount,BallsExited

Spin=Array("5","6","7","8","9","10","J","Q","K","A")
WildArray = Array(0,0,0,1,1,1,2,2,2,3,3,3,4,4,4,5,5,5,6)
Dim WildBallNumber
Dim WinnerHorse(6)
Dim HorseMovementFlag(6)
Dim PlayerHorse


Sub Table1_Init()
  DisableKeysInit=1
  LoadEM
  LoadLMEMConfig2
  If Table1.ShowDT = false then
    For each obj in DesktopCrap
      obj.visible=False
    next

  End If



  BallIndicatorCount=0
  InitBallLoad.enabled=true


  Points=39
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



  SpinPos=int(Rnd*10)
  HossPointer=int(Rnd*24)
  EVAL("YourHoss00"&HossSelector(HossPointer)).state=1

' CanPlayReel.SetValue(0)
  GameOverReel.SetValue(1)
  TiltReel.SetValue(1)
  Reel1000.SetValue(0)
  CanPlayLight1.state=0
  CanPlayLight2.state=0

  BallInPlayLight.state=0
  GameOverLight.state=1.
  BallInPlayReel.setvalue(0)

  TiltPlayer1.setvalue(1)
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


  CurrentFrame=0

  For i = 1 to 6
    EVAL("HorsePFLight00"&i).state=0
  next
  HossBG1=22
  HossBG2=22
  HossBG3=22
  HossBG4=22
  HossBG5=22
  HossBG6=22

  AdvanceLightCounter=0

  Players=0
  RotatorTemp=1
  InProgress=false
  TargetLightsOn=false
  BallLiftOption=1
  ScoreText.text=HighScore


  If B2SOn Then

    Controller.B2SSetData 222,1
    Controller.B2SSetData 192,1
    Controller.B2SSetData 162,1
    Controller.B2SSetData 132,1
    Controller.B2SSetData 102,1
    Controller.B2SSetData 72,1

    Controller.B2SSetScoreRolloverPlayer1 0
    Controller.B2SSetScoreRolloverPlayer2 0
    Controller.B2SSetScoreRolloverPlayer3 0
    Controller.B2SSetScoreRolloverPlayer4 0
    Controller.B2SSetData 99,0
    Controller.B2SSetScore 3,HighScore
    Controller.B2SSetTilt 41,1
    Controller.B2SSetTilt 42,1
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
  bump6=1
  InitPauser5.enabled=true
  If B2SOn then
    Controller.B2SSetCredits Credits
  end if
end sub

Sub InitBallLoad_timer
  Kicker3.CreateSizedBall 23
  Kicker3.Kick 120,5
  BallIndicatorCount=BallIndicatorCount+1
  if BallIndicatorCount=5 then
    Kicker3.enabled=false
    InitBallLoad.enabled=false

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





Sub Table1_KeyDown(ByVal keycode)
  If DisableKeysInit=1 then
    exit sub
  end if


  If keycode = PlungerKey Then
    Plunger.PullBack

  End If


  If keycode = LeftFlipperKey and TableTilted=false Then
    LeftFlipper.RotateToEnd
    PlaySoundAtVol "FlipperUp", LeftFlipper, 1
    PlayLoopSoundAtVol "buzzL", LeftFlipper, 1
    RightFlipper.RotateToEnd

    PlaySoundAtVol "FlipperUp", LeftFlipper, 1
    PlayLoopSoundAtVol "buzz", LeftFlipper, 1
    RightFlipper.timerenabled=true
  End If

  If keycode = RightFlipperKey and TableTilted=false Then
    LeftFlipper.RotateToEnd
    PlaySoundAtVol "FlipperUp", RightFlipper, 1
    PlayLoopSoundAtVol "buzzL", RightFlipper, 1
    RightFlipper.RotateToEnd

    PlaySoundAtVol "FlipperUp", RightFlipper, 1
    PlayLoopSoundAtVol "buzz", RightFlipper, 1
    RightFlipper.timerenabled=true
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

    playsound "coinin"
    AddSpecial2
  end if

   if keycode = 5 then
    playsound "coinin"
    AddSpecial2
    keycode= StartGameKey
  end if

  if (keycode=30 OR keycode=RightMagnasave) and InProgress=true and BallLiftOption=1 and BallsPlayed<BallsPerGame and LaneClear=0 then
    PlaySound "BallLifter"
    BallsPlayed=BallsPlayed+1
    'CreateBallID BallRelease
    Ballrelease.CreateSizedBall 25
    Ballrelease.Kick 90,5
    Playsound "ballrelease"
    BallsOnTable=BallsOnTable+1
  end if


  if keycode=StartGameKey and Credits>0 and (InProgress=false OR BallsOnTable=5) then

    BallsDrained=0
    DisableKeysInit=1
    Credits=Credits-1

    CreditsReel.SetValue(Credits)

    DrainAll
    ResetHosses
    Players=1
'   CanPlayReel.SetValue(Players)
    drop.image="MetalHole"
    Player=1
    CanPlayLight1.state=1
    GameOverLight.state=0
    If BallLiftOption=2 then
      playsound "BallsAPoppinStartup"
    else
      playsound "startup_manual"
    end if
    TempPlayerUp=Player
'   PlayerUpRotator.enabled=true
    rst=0
    BallInPlay=1
    Light3.state=1
    CenterTargetTracker=1
    Points=Points MOD 10
    InProgress=true
    resettimer.enabled=true
    BonusMultiplier=1
    LaneClear=0
    TableTilted=false
    TiltTimer.Interval = 500
    TiltTimer.Enabled = True
    Score(1)=0
    LightScore


    TiltPlayer1.setvalue(0)
    Player1Tilted=false
    Player2Tilted=false
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
      Controller.B2SSetData 41,0
      Controller.B2SSetData 42,0

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

Sub RightFlipper_Timer
  RightFlipper.timerenabled=false
  LeftFlipper.RotateToStart
  PlaySoundAtVol "FlipperDown", RightFlipper, 1
  StopSound "buzzL"
  RightFlipper.RotateToStart
    StopSound "buzz"
  PlaySoundAtVol "FlipperDown", RightFlipper, 1
end sub

Sub Table1_KeyUp(ByVal keycode)
  If DisableKeysInit=1 then
    exit sub
  end if


  If keycode = PlungerKey Then

    PlaySoundAtVol"plungerrelease", Plunger, 1
    Plunger.Fire
  End If

  If keycode = LeftFlipperKey and TableTilted=false Then
    LeftFlipper.RotateToStart
    PlaySoundAtVol "FlipperDown", LeftFlipper, 1
    StopSound "buzzL"
    RightFlipper.RotateToStart
        StopSound "buzz"
    PlaySoundAtVol "FlipperDown", LeftFlipper, 1
  End If

  If keycode = RightFlipperKey and TableTilted=false Then
    LeftFlipper.RotateToStart
    PlaySoundAtVol "FlipperDown", RightFlipper, 1
    StopSound "buzzL"
    RightFlipper.RotateToStart
        StopSound "buzz"
    PlaySoundAtVol "FlipperDown", RightFlipper, 1
  End If

End Sub

Sub ShooterLaneClear_Hit
  LaneClear=1
end sub

Sub ShooterLaneClear_Unhit
  LaneClear=0
end sub

Sub Drain_hit()
  if DisableKeysInit=1 then
    DisableKeysInit=0
  else
    EndOfGameTimer.enabled=true
  end if
end sub

Sub Drain004_Hit()
  Drain003.enabled=true
end sub

sub Drain003_Hit()
  Drain002.enabled=true
end sub

sub Drain002_Hit()
  Drain001.enabled=true
end sub

sub Drain001_Hit()
  Drain.enabled=true
end sub

Sub DrainAll()
  Drain.DestroyBall
  Drain001.DestroyBall
  Drain002.DestroyBall
  Drain003.DestroyBall
  Drain004.DestroyBall

    BallsDrained=0
    BallsExited=0
    Drain.Enabled=false
    Drain001.Enabled=false
    Drain002.Enabled=false
    Drain003.Enabled=false
    DrainLock.IsDropped=false




  BallsOnTable=0
  'PlaySound "fx_drain"

  'Pause4Bonustimer.enabled=true
  If BallsOnTable<1 then
    BallsOnTable=0
    Pause4BonusTimer.enabled=1
  end if
End Sub



Sub CloseGateTrigger_Hit()

End Sub

Sub Pause4Bonustimer_timer
  Pause4Bonustimer.enabled=0
  NextBallDelay.enabled=true

End Sub

Sub BallExit

end sub

'***********************
'     Hoss Timers
'***********************
dim HorseLoc1, HorseLoc2, HorseLoc3, HorseLoc4, HorseLoc5, HorseLoc6
dim HossBG1, HossBG2, HossBG3, HossBG4, HossBG5, HossBG6
dim HossBGPos1, HossBGPos2, HossBGPos3, HossBGPos4, HossBGPos5, HossBGPos6
dim HOSSSpots1, HOSSSpots2, HOSSSpots3, HOSSSpots4, HOSSSpots5, HOSSSpots6
HossSpots1 = array (946,928.5454545,911.0909091,893.6363636,876.1818182,858.7272727,841.2727273,823.8181818,806.3636364,788.9090909,771.4545455,754)
HossSpots2 = array (945,927.8181818,910.6363636,893.4545455,876.2727273,859.0909091,841.9090909,824.7272727,807.5454545,790.3636364,773.1818182,756)
HossSpots3 = array (944,927.0909091,910.1818182,893.2727273,876.3636364,859.4545455,842.5454545,825.6363636,808.7272727,791.8181818,774.9090909,758)
HossSpots4 = array (943,926.3636364,909.7272727,893.0909091,876.4545455,859.8181818,843.1818182,826.5454545,809.9090909,793.2727273,776.6363636,760)
HossSpots5 = array (942,925.6363636,909.2727273,892.9090909,876.5454545,860.1818182,843.8181818,827.4545455,811.0909091,794.7272727,778.3636364,762)
HossSpots6 = array (941,924.9090909,908.8181818,892.7272727,876.6363636,860.5454545,844.4545455,828.3636364,812.2727273,796.1818182,780.0909091,764)

Dim HossSelector, HossPointer
HossSelector = Array (3,5,1,2,4,3,1,6,2,1,4,5,3,1,6,2,4,1,3,5,1,4,2,6,1)

Sub HOSS006_timer
  dim tempxx
  Select Case HorseMovementFlag(6)
    case 1:  'reset
      tempxx=HOSS006.X
      If tempxx>941 then
        HorseLoc6=0
        HorseMovementFlag(6)=2
        Winner006.state=0
        WinnerHorse(6)=false
        HOSS006.timerenabled=false
      else
        tempxx=tempxx+1
        HossBGPos6=HossBGPos6+1
        HOSS006.X=tempxx
        If HossBGPos6>7 then
          HossBGPos6=0
          If B2SOn then
            Controller.B2SSetData 222,0
            Controller.B2SSetData 200+HossBG6,0
            HossBG6=HossBG6-1
            If HossBG6<0 then HossBG6=0
            Controller.B2SSetData 200+HossBG6,1
          end if
        end if

      end if
    case 2:  'move forward 1 step
      tempxx=HOSS006.X
      If tempxx<HOSSSpots6(HorseLoc6) then
        If HorseLoc6=11 then
          WinnerHorse(6)=true
          Winner006.state=1
          ScoreWinner
          HossBG6=22
          If B2SOn then
            Controller.B2SSetData 221,0
            Controller.B2SSetData 222,1
            Controller.B2SSetData 246,1
          end if
        else
          If B2SOn then
            Controller.B2SSetData 200+HossBG6,0
            HossBG6=HossBG6+1
            If HossBG6>22 then HossBG6=22
            Controller.B2SSetData 200+HossBG6,1
          end if
        end if
        HOSS006.timerenabled=false

      else
        tempxx=tempxx-(16.09091/10)
        HossBGPos6=HossBGPos6+1
        HOSS006.X=tempxx
        If HossBGPos6=4 then

          If B2SOn then

            Controller.B2SSetData 200+HossBG6,0
            HossBG6=HossBG6+1
            If HossBG6>22 then HossBG6=22
            Controller.B2SSetData 200+HossBG6,1
          end if
        End if
      end if
  end select
end sub

Sub HOSS005_timer
  dim tempxx
  Select Case HorseMovementFlag(5)
    case 1:  'reset
      tempxx=HOSS005.X
      If tempxx>942 then
        HorseLoc5=0
        HorseMovementFlag(5)=2
        Winner005.state=0
        WinnerHorse(5)=false
        HOSS005.timerenabled=false
      else
        tempxx=tempxx+1
        HOSS005.X=tempxx
        HossBGPos5=HossBGPos5+1
        If HossBGPos5>7 then
          HossBGPos5=0
          If B2SOn then
            Controller.B2SSetData 192,0
            Controller.B2SSetData 170+HossBG5,0
            HossBG5=HossBG5-1
            If HossBG5<0 then HossBG5=0
            Controller.B2SSetData 170+HossBG5,1
          end if
        end if

      end if
    case 2:  'move forward 1 step
      tempxx=HOSS005.X
      If tempxx<HOSSSpots5(HorseLoc5) then
        If HorseLoc5=11 then
          WinnerHorse(5)=true
          Winner005.state=1
          ScoreWinner
          HossBG5=22
          If B2SOn then
            Controller.B2SSetData 191,0
            Controller.B2SSetData 192,1
            Controller.B2SSetData 245,1
          end if
        else
          If B2SOn then
            Controller.B2SSetData 170+HossBG5,0
            HossBG5=HossBG5+1
            If HossBG5>22 then HossBG5=22
            Controller.B2SSetData 170+HossBG5,1
          end if
        end if
        HOSS005.timerenabled=false
      else
        tempxx=tempxx-(16.36364/10)
        HOSS005.X=tempxx
        HossBGPos5=HossBGPos5+1
        If HossBGPos5=4 then

          If B2SOn then

            Controller.B2SSetData 170+HossBG5,0
            HossBG5=HossBG5+1
            If HossBG5>22 then HossBG5=22
            Controller.B2SSetData 170+HossBG5,1
          end if
        End if
      end if
  end select
end sub

Sub HOSS004_timer
  dim tempxx
  Select Case HorseMovementFlag(4)
    case 1:  'reset
      tempxx=HOSS004.X
      If tempxx>943 then
        HorseLoc4=0
        HorseMovementFlag(4)=2
        Winner004.state=0
        WinnerHorse(4)=false
        HOSS004.timerenabled=false
      else
        tempxx=tempxx+1
        HOSS004.X=tempxx
        HossBGPos4=HossBGPos4+1
        If HossBGPos4>7 then
          HossBGPos4=0
          If B2SOn then
            Controller.B2SSetData 162,0
            Controller.B2SSetData 140+HossBG4,0
            HossBG4=HossBG4-1
            If HossBG4<0 then HossBG4=0
            Controller.B2SSetData 140+HossBG4,1
          end if
        end if
      end if
    case 2:  'move forward 1 step
      tempxx=HOSS004.X
      If tempxx<HOSSSpots4(HorseLoc4) then
        If HorseLoc4=11 then
          WinnerHorse(4)=true
          Winner004.state=1
          ScoreWinner
          HossBG4=22
          If B2SOn then
            Controller.B2SSetData 161,0
            Controller.B2SSetData 162,1
            Controller.B2SSetData 244,1
          end if
        else
          If B2SOn then
            Controller.B2SSetData 140+HossBG4,0
            HossBG4=HossBG4+1
            If HossBG4>22 then HossBG4=22
            Controller.B2SSetData 140+HossBG4,1
          end if
        end if
        HOSS004.timerenabled=false
      else
        tempxx=tempxx-(16.63636/10)
        HOSS004.X=tempxx
        HossBGPos4=HossBGPos4+1
        If HossBGPos4=4 then

          If B2SOn then

            Controller.B2SSetData 140+HossBG4,0
            HossBG4=HossBG4+1
            If HossBG4>22 then HossBG4=22
            Controller.B2SSetData 140+HossBG4,1
          end if
        End if
      end if
  end select
end sub

Sub HOSS003_timer
  dim tempxx
  Select Case HorseMovementFlag(3)
    case 1:  'reset
      tempxx=HOSS003.X
      If tempxx>944 then
        HorseLoc3=0
        HorseMovementFlag(3)=2
        Winner003.state=0
        WinnerHorse(3)=false
        HOSS003.timerenabled=false
      else
        tempxx=tempxx+1
        HOSS003.X=tempxx
        HossBGPos3=HossBGPos3+1
        If HossBGPos3>7 then
          HossBGPos3=0
          If B2SOn then
            Controller.B2SSetData 132,0
            Controller.B2SSetData 110+HossBG3,0
            HossBG3=HossBG3-1
            If HossBG3<0 then HossBG3=0
            Controller.B2SSetData 110+HossBG3,1
          end if
        end if
      end if
    case 2:  'move forward 1 step
      tempxx=HOSS003.X
      If tempxx<HOSSSpots3(HorseLoc3) then
        If HorseLoc3=11 then
          WinnerHorse(3)=true
          Winner003.state=1
          ScoreWinner
          HossBG3=22
          If B2SOn then
            Controller.B2SSetData 131,0
            Controller.B2SSetData 132,1
            Controller.B2SSetData 243,1
          end if
        else
          If B2SOn then
            Controller.B2SSetData 110+HossBG3,0
            HossBG3=HossBG3+1
            If HossBG3>22 then HossBG3=22
            Controller.B2SSetData 110+HossBG3,1
          end if
        end if
        HOSS003.timerenabled=false
      else
        tempxx=tempxx-(16.90909/10)
        HOSS003.X=tempxx
        HossBGPos3=HossBGPos3+1
        If HossBGPos3=4 then

          If B2SOn then

            Controller.B2SSetData 110+HossBG3,0
            HossBG3=HossBG3+1
            If HossBG3>22 then HossBG3=22
            Controller.B2SSetData 110+HossBG3,1
          end if
        End if
      end if
  end select
end sub

Sub HOSS002_timer
  dim tempxx
  Select Case HorseMovementFlag(2)
    case 1:  'reset
      tempxx=HOSS002.X
      If tempxx>945 then
        HorseLoc2=0
        HorseMovementFlag(2)=2
        Winner002.state=0
        WinnerHorse(2)=false
        HOSS002.timerenabled=false
      else
        tempxx=tempxx+1
        HOSS002.X=tempxx
        HossBGPos2=HossBGPos2+1
        If HossBGPos2>7 then
          HossBGPos2=0
          If B2SOn then
            Controller.B2SSetData 102,0
            Controller.B2SSetData 80+HossBG2,0
            HossBG2=HossBG2-1
            If HossBG2<0 then HossBG2=0
            Controller.B2SSetData 80+HossBG2,1
          end if
        end if
      end if
    case 2:  'move forward 1 step
      tempxx=HOSS002.X
      If tempxx<HOSSSpots2(HorseLoc2) then
        If HorseLoc2=11 then
          WinnerHorse(2)=true
          Winner002.state=1
          ScoreWinner
          HossBG2=22
          If B2SOn then
            Controller.B2SSetData 101,0
            Controller.B2SSetData 102,1
            Controller.B2SSetData 242,1
          end if
        else
          If B2SOn then
            Controller.B2SSetData 80+HossBG2,0
            HossBG2=HossBG2+1
            If HossBG2>22 then HossBG2=22
            Controller.B2SSetData 80+HossBG2,1
          end if
        end if
        HOSS002.timerenabled=false
      else
        tempxx=tempxx-(17.18182/10)
        HOSS002.X=tempxx
        HossBGPos2=HossBGPos2+1
        If HossBGPos2=4 then

          If B2SOn then

            Controller.B2SSetData 80+HossBG2,0
            HossBG2=HossBG2+1
            If HossBG2>22 then HossBG2=22
            Controller.B2SSetData 80+HossBG2,1
          end if
        End if
      end if
  end select
end sub

Sub HOSS001_timer
  dim tempxx
  Select Case HorseMovementFlag(1)
    case 1:  'reset
      tempxx=HOSS001.X
      If tempxx>946 then
        HorseLoc1=0
        HorseMovementFlag(1)=2
        Winner001.state=0
        WinnerHorse(1)=false
        HOSS001.timerenabled=false
      else
        tempxx=tempxx+1
        HOSS001.X=tempxx
        HossBGPos1=HossBGPos1+1
        If HossBGPos1>7 then
          HossBGPos1=0
          If B2SOn then
            Controller.B2SSetData 72,0
            Controller.B2SSetData 50+HossBG1,0
            HossBG1=HossBG1-1
            If HossBG1<0 then HossBG1=0
            Controller.B2SSetData 50+HossBG1,1
          end if
        end if
      end if
    case 2:  'move forward 1 step
      tempxx=HOSS001.X
      If tempxx<HOSSSpots1(HorseLoc1) then
        If HorseLoc1=11 then
          WinnerHorse(1)=true
          Winner001.state=1
          ScoreWinner
          HossBG1=22
          If B2SOn then
            Controller.B2SSetData 71,0
            Controller.B2SSetData 72,1
            Controller.B2SSetData 241,1
          end if
        else
          If B2SOn then
            Controller.B2SSetData 50+HossBG1,0
            HossBG1=HossBG1+1
            If HossBG1>22 then HossBG1=22
            Controller.B2SSetData 50+HossBG1,1
          end if
        end if
        HOSS001.timerenabled=false
      else
        tempxx=tempxx-(17.45455/10)
        HOSS001.X=tempxx
        HossBGPos1=HossBGPos1+1
        If HossBGPos1=4 then

          If B2SOn then

            Controller.B2SSetData 50+HossBG1,0
            HossBG1=HossBG1+1
            If HossBG1>22 then HossBG1=22
            Controller.B2SSetData 50+HossBG1,1
          end if
        End if
      end if
  end select
end sub

Sub MoveHoss(HossNumber)

  Dim tempii
  HorseMovementFlag(HossNumber)=2
  for tempii = 1 to 6
    if WinnerHorse(tempii)=true then EXIT SUB
  next
  HossPointer=HossPointer+1
  If EVAL("HOSS00"&HossNumber).timerenabled=true then EXIT SUB
  If HossPointer>24 then HossPointer=0
  Select Case HossNumber
    case 1:
      HorseLoc1=HorseLoc1+1
      If HorseLoc1>11 then HorseLoc1=11
      HossBGPos1=0
    case 2:
      HorseLoc2=HorseLoc2+1
      If HorseLoc2>11 then HorseLoc2=11
      HossBGPos2=0
    case 3:
      HorseLoc3=HorseLoc3+1
      If HorseLoc3>11 then HorseLoc3=11
      HossBGPos3=0
    case 4:
      HorseLoc4=HorseLoc4+1
      If HorseLoc4>11 then HorseLoc4=11
      HossBGPos4=0
    case 5:
      HorseLoc5=HorseLoc5+1
      If HorseLoc5>11 then HorseLoc5=11
      HossBGPos5=0
    case 6:
      HorseLoc6=HorseLoc6+1
      If HorseLoc6>11 then HorseLoc6=11
      HossBGPos6=0
  end select
  EVAL("HOSS00"&HossNumber).timerenabled=true
  PlaySound "HorseMove"
end sub

sub ResetHosses
  for i = 1 to 6
    HorseMovementFlag(i)=1

  next
  HossBGPos1=0
  HossBGPos2=0
  HossBGPos3=0
  HossBGPos4=0
  HossBGPos5=0
  HossBGPos6=0
  PlaySound "MotorLeer",-1
  HOSS006.timerenabled=true
  HOSS005.timerenabled=true
  HOSS004.timerenabled=true
  HOSS003.timerenabled=true
  HOSS002.timerenabled=true
  HOSS001.timerenabled=true

end sub

Sub ScoreWinner
  Select Case BallsPlayed
    case 1:

      If B2SOn then Controller.B2SSetData 20,1: Controller.B2SSetData 5,1
      If WinnerHorse(HossSelector(PlayerHorse))=true then
        AddMultiSpecials=25
        AddSpecialTimer.enabled=true
      end if
    case 2:
      If B2SOn then Controller.B2SSetData 20,1
      If WinnerHorse(HossSelector(PlayerHorse))=true then
        AddMultiSpecials=5
        AddSpecialTimer.enabled=true
      end if
    case 3:
      If B2SOn then Controller.B2SSetData 10,1: Controller.B2SSetData 5,1
      If WinnerHorse(HossSelector(PlayerHorse))=true then
        AddMultiSpecials=3
        AddSpecialTimer.enabled=true
      end if
    case 4:
      If B2SOn then Controller.B2SSetData 10,1
      If WinnerHorse(HossSelector(PlayerHorse))=true then
        AddMultiSpecials=2
        AddSpecialTimer.enabled=true
      end if
    case 5:
      If B2SOn then Controller.B2SSetData 5,1
      If WinnerHorse(HossSelector(PlayerHorse))=true then
        AddSpecial
      end if
  end select

  EndOfGameTimer.enabled=true
end sub

Sub HorseSelection
  dim ix
  HossPointer=HossPointer+1
  If HossPointer>24 then HossPointer=0
  for i = 1 to 6
    EVAL("YourHoss00"&i).state=0
    EVAL("HorsePFLight00"&i).state=0
    if B2SOn then
      For ix = 231 to 236
        Controller.B2SSetData ix,0
        Controller.B2SSetData ix+10,0
      next
    end if
  next
  EVAL("YourHoss00"&HossSelector(HossPointer)).state=1
  EVAL("HorsePFLight00"&HossSelector(HossPointer)).state=1
  If B2SOn then
    Controller.B2SSetData (230+HossSelector(HossPointer)),1
  end if
  PlayerHorse=HossPointer
end sub

Sub HossSelectTimer_timer
  if HOSS006.timerenabled=false AND HOSS005.timerenabled=false AND HOSS004.timerenabled=false AND HOSS003.timerenabled=false AND HOSS002.timerenabled=false AND HOSS001.timerenabled=false then
    HossSelectTimer.enabled=false
    StopSound "MotorLeer"
    PlaySound "HorseResetDone"
    DisableKeysInit=0
    Exit Sub
  end if
  DisableKeysInit=1
  HorseSelection
end sub


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


'***********************************
' Walls
'***********************************

Sub RubberWallSwitches_Hit(idx)
  if TableTilted=false then

    Select Case idx
      Case 0:
        MoveHoss(1)
        MoveHoss(3)
        MoveHoss(5)
        SetMotor(5)
      case 1:
        MoveHoss(2)
        MoveHoss(4)
        MoveHoss(6)
        SetMotor(5)
      case 2,3:
        AddScore(1)
    end Select
  end if
end Sub


'***********************************
' Bumpers
'***********************************
Sub Bumper1_Hit
  If TableTilted=false then

    PlaySoundAtVol "bumper1", ActiveBall, 1

    bump1 = 1
    AddScore(1)
    MoveHoss(1)
    end if

End Sub


Sub Bumper2_Hit
  If TableTilted=false then

    PlaySoundAtVol "bumper1", ActiveBall, 1
    bump2 = 1
    AddScore(1)
    MoveHoss(2)
    end if

End Sub

Sub Bumper3_Hit
  If TableTilted=false then

    PlaySoundAtVol "bumper1", ActiveBall, 1
    bump3 = 1
    AddScore(1)
    MoveHoss(3)
    end if

End Sub


Sub Bumper4_Hit
  If TableTilted=false then
    PlaySoundAtVol "bumper1", ActiveBall, 1
    bump4 = 1
    AddScore(1)
    MoveHoss(4)
    end if

End Sub


Sub Bumper5_Hit
  If TableTilted=false then

    PlaySoundAtVol "bumper1", ActiveBall, 1
    bump5 = 1
    AddScore(1)
    MoveHoss(5)
    end if

End Sub

Sub Bumper6_Hit
  If TableTilted=false then
    PlaySoundAtVol "bumper1", ActiveBall, 1
    bump6 = 1
    AddScore(1)
    MoveHoss(6)

    end if

End Sub


'************************************
'  Rollover lanes
'************************************

Sub TriggerCollection_Hit(idx)

  If TableTilted=false then


    Select Case (idx)
      case 0:
        MoveHoss(1)
        MoveHoss(3)
        MoveHoss(5)
        SetMotor(5)
      case 1:
        HorseSelection
        AddScore(10)
      case 2:
        MoveHoss(2)
        MoveHoss(4)
        MoveHoss(6)
        SetMotor(5)
      case 3:
        MoveHoss(1)
        MoveHoss(2)
        MoveHoss(3)
        MoveHoss(4)
        MoveHoss(5)
        MoveHoss(6)
        SetMotor(5)
    end select
  end if

end Sub


'**************************************
Dim AddMultiSpecials

Sub AddSpecial()
  PlaySound "knocker"

  Credits=Credits+1

  if Credits>25 then Credits=25
  If B2SOn Then
    Controller.B2SSetCredits Credits
  End If
  CreditsReel.SetValue(Credits)
End Sub

Sub AddSpecial2()
  PlaySound"click"
  Credits=Credits+1

  if Credits>25 then Credits=25
  If B2SOn Then
    Controller.B2SSetCredits Credits
  End If
  CreditsReel.SetValue(Credits)
End Sub

Sub AddSpecialTimer_timer
  AddSpecial
  AddMultiSpecials=AddMultiSpecials-1
  if AddMultiSpecials<1 then AddSpecialTimer.enabled=false
end sub



Sub ResetBallDrops



  HoleCounter=0
  YellowBumpersOn=false
  GreenBumpersOn=false

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
    PlayerUpLight1.state=0
    PlayerUpLight2.state=0
    Eval("PlayerUpLight"&Player).state=1
    PlayStartBall.enabled=true
    BallsPlayed=BallsPlayed+1
    'CreateBallID BallRelease
    Ballrelease.CreateSizedBall 23
    Ballrelease.Kick 90,5

    BallsOnTable=BallsOnTable+1
    BallInPlayLight.state=1
    BallInPlayReel.SetValue(BallInPlay)

  end if


End Sub





sub resettimer_timer
    rst=rst+1
  HorseSelection

    if rst=11 then
    if BallLiftOption=2 then
      playsound "StartBall1"
    end if
    end if
    if rst=12 then
    newgame
    DisableKeysInit=0
    drop.image="Metal3"
    resettimer.enabled=false
    HossSelectTimer.enabled=true
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

'   Controller.B2SSetScorePlayer1 0
'   Controller.B2SSetScorePlayer2 0
'   Controller.B2SSetScorePlayer3 0
'   Controller.B2SSetScorePlayer4 0
    Controller.B2SSetBallInPlay BallInPlay
  End if


  Bumper1Light.state=1
  Bumper2Light.state=1
  Bumper3Light.state=1
  Bumper4Light.state=1
  Bumper5Light.state=1
  Bumper6Light.state=1

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
  'PlaySound("MotorLeer")
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


'   BallInPlayReel.SetValue(0)
'     CanPlayReel.SetValue(0)
  LeftFlipper.RotateToStart
  RightFlipper.RotateToStart
  LightsOut
  Light3.state=0
  BumpersOff
  PlasticsOff


  Players=0
  If Table1.ShowDT = True Then
    PlayerScores(Player-1).visible=1
    PlayerScoresOn(Player-1).visible=0
  end If


end sub

sub EndOfGameTimer_timer
  EndOfGameTimer.enabled=false
  EndOfGame
end sub

sub nextball

  If TiltEndsGame=1 and TableTilted=true and BallsOnTable>0 then
    exit sub
  end if
  Player=Player+1
  If Player>Players Then
    BallInPlay=BallInPlay+1
    If (BallInPlay>BallsPerGame) then
      'PlaySound("MotorLeer")


      EndGameCount=0
      EndOfGameTimer.enabled=true
      exit sub
    Else
      Player=1
      If Player1Tilted=True then
        Player=2
      end if
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
    If Player2Tilted=true then
      nextball
      exit sub
    end if
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

      If Player=1 then
        Player1Tilted=true
        TiltPlayer1.SetValue(1)
      else

      end if


      If (Player1Tilted=true) then

        InProgress=false
        'BallInPlay=5
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


        Players=0
        EndOfGame


      end if
    else
      TiltTimer.Interval = 500
      TiltTimer.Enabled = True
    end if

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

    TextStr.Close
      Credits=cdbl(temp2)
    BallsPerGame=cdbl(temp3)
    ArrowCounter=cdbl(temp4)
    AdvanceSpecialMax=cdbl(temp5)
    BallLiftOption=cdbl(temp6)
    TiltEndsGame=cdbl(temp7)
    ReplayLevel=cdbl(temp8)

    Set ScoreFile=Nothing
      Set FileObj=Nothing
end sub


sub InitPauser5_timer
  If B2SOn then
    Controller.B2SSetCredits Credits
  end if

    CreditsReel.SetValue(Credits)

    InitPauser5.enabled=false

end sub



sub BumpersOff

  Bumper1Light.visible=0
  Bumper2Light.visible=0
  Bumper3Light.visible=0
  Bumper4Light.visible=0

  Bumper5Light.visible=0
  Bumper6Light.visible=0



end sub

sub BumpersOn

  Bumper1Light.visible=1
  Bumper2Light.visible=1
  Bumper3Light.visible=1
  Bumper4Light.visible=1

  Bumper5Light.visible=1
  Bumper6Light.visible=1



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
  If AddScoreTimer.enabled=true then exit sub
  QueuedScore=y
  AddScoreTimer.enabled=true
end sub



Sub AddScoreTimer_Timer
  AddScore(1)
  QueuedScore=QueuedScore-1
  If QueuedScore<1 then AddScoreTimer.enabled=false

end Sub

Sub ScoreMotorTimer_Timer
  ScoreMotorTimer.enabled=false
End Sub


Sub AddScore(qqq)
  Score(1)=Score(1) + qqq
  If Score(1)>499 then Score(1)=499
  LightScore
  Exit Sub

end sub

Sub LightScore
  dim ones,tens,hund,tempscore
  for each obj in ScoringLights
    obj.state=0
  next
  ones=Score(1) mod 10
  tens=(Score(1) \ 10) mod 10
  hund=(Score(1)\ 100) mod 10
  If ones=0 then
    for tempscore = 1 to 9
      If B2SOn then Controller.B2SSetData tempscore,0
    next
  end if
  If ones>0 then
    EVAL("Score00"&ones).state=1
    for tempscore = 1 to 9
      If B2SOn then Controller.B2SSetData tempscore,0
    next
    If B2SOn then Controller.B2SSetData ones,1
  end if
  If tens=0 then
    for tempscore = 11 to 19
      If B2SOn then Controller.B2SSetData tempscore,0
    next
  end if
  If tens>0 then
    EVAL("Score0"&tens&"0").state=1
    for tempscore = 11 to 19
      If B2SOn then Controller.B2SSetData tempscore,0
    next
    If B2SOn then Controller.B2SSetData (tens+10),1
  end if
  if hund=0 then
    for tempscore = 21 to 24
      If B2SOn then Controller.B2SSetData tempscore,0
    next
  end if
  if hund>0 then
    tempscore=FormatNumber(hund,0)

    Select case tempscore
      case 1:
        ScoringLights(3).state=1
      case 2:
        ScoringLights(2).state=1
      case 3:
        ScoringLights(1).state=1
      case 4:
        ScoringLights(0).state=1
    end select
    for tempscore = 21 to 24
      If B2SOn then Controller.B2SSetData tempscore,0
    next
    If B2SOn then Controller.B2SSetData (hund+20),1
  end if
end sub


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

Const tnob = 9 ' total number of balls
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
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5,BallShadow6,BallShadow7,BallShadow8,BallShadow9)

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


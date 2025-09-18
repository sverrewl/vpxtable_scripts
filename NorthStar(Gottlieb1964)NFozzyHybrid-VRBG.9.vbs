'*
'*        Gottlieb's North Star (1964)
'*        Table primary build/scripted by Loserman76
'*        Table images from PBecker's VP9 table
'*
'*
'*

option explicit
Randomize
ExecuteGlobal GetTextFile("core.vbs")
SetLocale 1033

On Error Resume Next
ExecuteGlobal GetTextFile("Controller.vbs")
If Err Then MsgBox "Unable to open Controller.vbs. Ensure that it is in the scripts folder."
On Error Goto 0

'****************************
'PLAYER OPTIONS
'****************************
Dim BallRollVolume: BallRollVolume = 0.8  'Level of ball rolling volume. Value between 0 and 1

'*****VRCabinetOptions*************************************
Dim CabArtwork
CabArtwork = 1  '1 = Original (Default), 2 Frostdesign
'****************************************************

'********SCORBITOPTIONS*********************
dim TablesDir : TablesDir = GetTablesFolder
Const     ScorbitAlternateUUID  = 0   ' Force Alternate UUID from Windows Machine and saves it in VPX Users directory (C:\Visual Pinball\User\ScorbitUUID.dat)
Dim ScorbitActive : ScorbitActive = 0
Const myVersion = "1.0.0"
Dim ScorbitQRFolder : ScorbitQRFolder = "NORTHSTARQR"


Const cGameName = "NorthStar_1964"
Const ShadowFlippersOn = true
Const ShadowBallOn = true
Const ShadowConfigFile = false


Dim Controller  ' B2S
Dim B2SScore  ' B2S Score Displayed
Const HSFileName="NorthStar_64VPX.txt"
Const B2STableName="NorthStar_1964"
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
Dim Replay1Table(25)
Dim Replay2Table(25)
Dim Replay3Table(25)
Dim Replay4Table(25)
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

Dim HorseCounter,CowboyCounter,CowboySpins

Dim RotoReplay1,RotoReplay2,RotoReplay3,RotoReplay4,RotoReplayTimerCounter,ABCDComplete,AComplete,BComplete,CComplete,DComplete

Dim CompleteA,CompleteK,CompleteQ,CompleteJ, CompleteClubs,DisableKeysInit,BallLiftOption,BallsPlayed,BallsDrained

Dim SpinPos,Spin,Count,Count1,Count2,Count3,Reset,VelX,VelY,LitSpinner,BallIndicatorCount,CenterTargetTracker,BallsOnTable,TiltEndsGame
Dim KickerLightCounter,BumperLightCounter
Dim tKickerTCount1,tKickerTCount2,tKickerTCount3,tKickerTCount4

Spin=Array("5","6","7","8","9","10","J","Q","K","A")

Dim RotoImagesL1, RotoImagesL2, RotoImagesR1, RotoImagesR2
Dim RotoPosition, RotoFlip

Dim RotoWheel(15)       ' Actual Roto Wheel Number array
Dim RotoPos             ' Index to Left Showing RotoWheel Number
Dim RWAnim        ' Wheel Animation Phase
Dim RotoCycles      ' Number of Digits To Animate on a Spin


RotoImagesL1 = array (4,4,1,1,2,2,7,7,3,3,5,5,6,6,7,7,1,1,2,2,3,3,5,5,1,1,6,6,7,7)
RotoImagesL2 = array (7,7,4,4,1,1,2,2,7,7,3,3,5,5,6,6,7,7,1,1,2,2,3,3,5,5,1,1,6,6)
RotoImagesR1 = array (1,1,6,6,7,7,4,4,1,1,2,2,7,7,3,3,5,5,6,6,7,7,1,1,2,2,3,3,5,5)
RotoImagesR2 = array (5,5,1,1,6,6,7,7,4,4,1,1,2,2,7,7,3,3,5,5,6,6,7,7,1,1,2,2,3,3)

'****************************** Script for VR Rooms  ********************

Dim VRRoom, cabmode, Object, DesktopMode: DesktopMode = Table1.ShowDT
If RenderingMode = 2 Then VRRoom=1 Else VRRoom=0
If Not DesktopMode and VRRoom=0 Then cabmode=1 Else cabmode=0
If cabmode = 1 Then
    for each Object in ColBackdrop : object.visible = 0 : next
    for each Object in ColRoomMinimal : object.visible = 0 : next
    for each Object in ColRoomSphere : object.visible = 0 : next
    for each Object in ColCabinet : object.visible = 0 : next
  for each Object in ColCabinetMode : object.visible = 1 : next
End If



If DesktopMode Then
    for each Object in ColBackdrop : object.visible = 1 : next
    for each Object in ColRoomMinimal : object.visible = 0 : next
    for each Object in ColCabinet : object.visible = 0 : next
    for each Object in ColRoomSphere : object.visible = 0 : next
End If



VRRoom = LoadValue("NorthStar", "V1.0.0")  ' loading the last value saved.  These are saved below each time you make a change with magnasave

Sub Room1()
  for each Object in ColBackdrop : object.visible = 0 : next
  for each Object in ColRoomMinimal : object.visible = 1 : next
  for each Object in ColCabinet : object.visible = 1 : next
  for each Object in ColRoomSphere : object.visible = 0 : next
  If CabArtwork = 1 Then
    VR_Cab_BackWedge.image = "VR_Back_UV_Map"
    VR_Cab.image = "VR_Cab_UV_Map"
  Else
    VR_Cab_BackWedge.image = "VR_Back_UV_Map2"
    VR_Cab.image = "VR_Cab_UV_Map2"
  End If
  SaveValue "NorthStar", "V1.0.0", 0
End Sub

Sub RoomSphere()
  for each Object in ColBackdrop : object.visible = 0 : next
  for each Object in ColRoomMinimal : object.visible = 0 : next
  for each Object in ColCabinet : object.visible = 1 : next
  for each Object in ColRoomSphere : object.visible = 1 : next
  If CabArtwork = 1 Then
    VR_Cab_BackWedge.image = "VR_Back_UV_Map"
    VR_Cab.image = "VR_Cab_UV_Map"
  Else
    VR_Cab_BackWedge.image = "VR_Back_UV_Map2"
    VR_Cab.image = "VR_Cab_UV_Map2"
  End If
  SaveValue "NorthStar", "V1.0.0", 1
End Sub

Sub VRChangeRoom()
  If VRRoom="" Then  ' If this is the first run of the table, there will be no value saved, we default it to 1 here..
    VRRoom = 0
  End if
  If VRRoom = 0 Then
    Room1
  End If
  If VRRoom = 1 Then
    RoomSphere
  End If
End Sub


Sub Table1_Init()
  DisableKeysInit=1
  LoadEM
  LoadLMEMConfig2
  If Table1.ShowDT = false then
    BallsPlayedRamp.image="FSBallsPlayed"
  End If

  OperatorMenuBackdrop.imageA = "PostitBL"
  For XOpt = 1 to MaxOption
    Eval("OperatorOption"&XOpt).imageA = "PostitBL"
  next

  For XOpt = 1 to 256
    Eval("Option"&XOpt).imageA = "PostItBL"
  next


  BallIndicatorCount=0
  InitBallLoad.enabled=true


  Count=0
  Count1=0
  Count2=0
  Count3=0
  Reset=0
  ZeroToNineUnit=Int(Rnd*10)
  Kicker1Hold=0
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
  OperatorMenu=0
  HighScore=0
  MotorRunning=0
  HighScoreReward=1
  Motor3EToggle=0
  Credits=0
  BallsPerGame=5
  ReplayLevel=1
  AdvanceBonusMax=0
  AdvanceSpecialMax=1
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
  MatchReel.SetValue((Match)+1)

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


  Replay1=Replay1Table(ReplayLevel)
  Replay2=Replay2Table(ReplayLevel)
  Replay3=Replay3Table(ReplayLevel)
  Replay4=Replay4Table(ReplayLevel)


  BonusCounter=0
  HoleCounter=0


  InstructCard.image="IC"+FormatNumber(BallsPerGame,0)

  RefreshReplayCard

  CurrentFrame=0

  AdvanceLightCounter=0

  Players=0
  RotatorTemp=1
  InProgress=false
  TargetLightsOn=false

  ScoreText.text=HighScore

  HorseCounter=90
  CowboyCounter=61
  CowboySpins=0

  If B2SOn Then

    if Match=0 then
      Controller.B2SSetMatch 10
    else
      Controller.B2SSetMatch Match
    end if
    Controller.B2SSetScoreRolloverPlayer1 0
    Controller.B2SSetScoreRolloverPlayer2 0
    Controller.B2SSetScoreRolloverPlayer3 0
    Controller.B2SSetScoreRolloverPlayer4 0
    Controller.B2SSetData 98,1
    Controller.B2SSetData 99,0
    Controller.B2SSetScore 3,HighScore
    Controller.B2SSetTilt 1
    Controller.B2SSetCredits Credits
    Controller.B2SSetGameOver 1
    For i = 91 to 93
      Controller.B2SSetData i,0
    next
    For i = 62 to 80
      Controller.B2SSetData i,0
    next
    Controller.B2SSetData 90,1
    Controller.B2SSetData 61,1
  End If

  If RenderingMode=2 Then
    VRRoom = LoadValue("NorthStar", "V1.0.0")
    VRChangeRoom
    VRGameOver.visible = 1
    FlasherMatch
    cred =reels(4, 0)
    reels(4, 0) = 0
    SetDrum -1,0,  0
    SetReel 0,-1,  Credits
    reels(4, 0) = Credits
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
    Controller.B2SSetData 60+SpinPos,1
  end if
' SpinnerLights(SpinPos).state=1
  if Credits > 0 then DOF 135, DOFOn


  StartScorbit()

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

' Change VR-Room with magna-save buttons

  If keycode = rightmagnasave and InProgress=False and not cabmode=1 then
    VRRoom = VRRoom +1
    if VRRoom > 1 then VRRoom = 0
    If Renderingmode = 2 Then
      VRChangeRoom
    End If
  End If

  If keycode = PlungerKey Then
    Plunger.PullBack
    PlungerPulled = 1
  End If

  if keycode = LeftFlipperKey and InProgress = false then
    OperatorMenuTimer.Enabled = true
  end if
  ' END GNMOD


  If keycode = LeftFlipperKey and InProgress=true and TableTilted=false Then
    FlipperActivate LeftFlipper, LFPress
    LeftFlipper.RotateToEnd
    PlaySoundAtVol SoundFXDOF("FlipperUp",101,DOFOn,DOFContactors), LeftFlipper, 1
    PlayLoopSoundAtVol "buzzL", LeftFlipper, 1
  End If

  If keycode = RightFlipperKey  and InProgress=true and TableTilted=false Then
    RightFlipper.RotateToEnd
    FlipperActivate RightFlipper, RFPress
    PlaySoundAtVol SoundFXDOF("FlipperUp",102,DOFOn,DOFContactors), RightFlipper, 1
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
    TiltCount=3
    TiltIt
  End If

  if keycode = 3 then
    AddScore(1)
    AdvanceZeroToNine
  end if

  If keycode = AddCreditKey or keycode = 4 then
    If B2SOn Then
      'Controller.B2SSetScorePlayer6 HighScore

    End If

    PlaySoundAtVol "Coin_In_3", Drain, 1
    AddSpecial2
  end if

   if keycode = 5 then
    PlaySoundAtVol "Coin_In_1", Drain, 1
    AddSpecial2
    keycode= StartGameKey
  end if

  If (keycode=StartGameKey OR keycode=30 OR keycode=RightMagnasave) and InProgress=true and BallLiftOption=1 and BallsPlayed<BallsPerGame and LaneClear=0 then
    PlaySoundAtVol SoundFXDOF("BallLifter",136,DOFPulse,DOFContactors), Plunger, 1
    BallsPlayed=BallsPlayed+1
    'CreateBallID BallRelease
    Ballrelease.CreateSizedBall 25
    Ballrelease.Kick 270,5
    PlaySoundAtVol ("ballrelease"), Ballrelease, 1
    Scorbit_Paired()
    BallsOnTable=BallsOnTable+1
  End if

  if keycode=StartGameKey and Credits>0 and InProgress=false and Players=0 and EnteringOptions = 0 and BallIndicatorCount=5 then
'GNMOD
    OperatorMenuTimer.Enabled = false
'END GNMOD
    DisableKeysInit=1
    Credits=Credits-1
    if Credits < 1 then DOF 135, DOFOff
    CreditsReel.SetValue(Credits)
    Gate2.Collidable=false
    Players=1
'   CanPlayReel.SetValue(Players)
    MatchReel.SetValue(0)
    Player=1
    If ScorbitActive = 1 Then
      Scorbit.StartSession
    End If
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
      Controller.B2SSetScore 3,HighScore
      Controller.B2SSetCanPlay 1
      Controller.B2SSetPlayerUp 1
      Controller.B2SSetBallInPlay BallInPlay
      Controller.B2SSetScoreRolloverPlayer1 0
      Controller.B2SSetData 98,1
      Controller.B2SSetData 99,0
      Controller.B2SSetData 11,0
      Controller.B2SSetData 12,0
      Controller.B2SSetData 13,0
      Controller.B2SSetData 14,0
      Controller.B2SSetData 15,0
    End If
    If Renderingmode = 2 Then
      cred =reels(4, 0)
      reels(4, 0) = 0
      SetDrum -1,0,  0
      SetReel 0,-1,  Credits
      reels(4, 0) = Credits
    End If

    If Table1.ShowDT = True then
      For each obj in PlayerScores
'       obj.ResetToZero
        obj.Visible=true
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
  end if

'******** Copy from this green line to next green line and insert it in the Sub Table1_KeyDown *******
' LUT-Changer
  If Keycode = LeftMagnaSave And InProgress = False Then
        LUTSet = LUTSet  + 1
    if LutSet > 15 then LUTSet = 0
        lutsetsounddir = 1
      If LutToggleSound then
        If lutsetsounddir = 1 And LutSet <> 15 Then
          Playsound "click", 0, 1, 0.1, 0.25
        End If
      If lutsetsounddir = -1 And LutSet <> 15 Then
          Playsound "click", 0, 1, 0.1, 0.25
        End If
      If LutSet = 15 Then
          Playsound "gun", 0, 1, 0.1, 0.25
        End If
        LutSlctr.enabled = true
      SetLUT
      ShowLUT
    End If
  End If

  If keycode = LeftFlipperKey And InProgress = True Then
    VR_CabFlipperLeft.X = VR_CabFlipperLeft.X +10
  End If

  If keycode = RightFlipperKey And InProgress = True Then
    VR_CabFlipperRight.X = VR_CabFlipperRight.X - 10
  End If

  If keycode = PlungerKey Then
    TimerPlunger.Enabled = True
    TimerPlunger2.Enabled = False
  End If

  If keycode = AddCreditKey or keycode = 4 then
    VR_Cab_StartButton.y = VR_Cab_StartButton.y -5
  End If

  If Keycode = StartGameKey Then
    VR_Cab_StartButton.y = VR_Cab_StartButton.y -5
  End If

'**************************************************************


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
    FlipperDeActivate LeftFlipper, LFPress
    LeftFlipper.RotateToStart
    PlaySoundAtVol SoundFXDOF("FlipperDown",101,DOFOff,DOFContactors), LeftFlipper, 1
    StopSound "buzzL"
  End If

  If keycode = RightFlipperKey and InProgress=true and TableTilted=false Then
    FlipperDeActivate RightFlipper, RFPress
    RightFlipper.RotateToStart
    PlaySoundAtVol SoundFXDOF("FlipperDown",102,DOFOff,DOFContactors), RightFlipper, 1
   StopSound "buzz"
  End If

'******** VR Stuff Rajo Joey*******
  If keycode = PlungerKey Then
    TimerPlunger.Enabled = False
    TimerPlunger2.Enabled = True
    VR_Primary_plunger.Y = -100
  end if

  If keycode = LeftFlipperKey And InProgress = True Then
    VR_CabFlipperLeft.X = VR_CabFlipperLeft.X -10
  End If

  If keycode = RightFlipperKey And InProgress = True Then
    VR_CabFlipperRight.X = VR_CabFlipperRight.X +10
  End If

  If Keycode = StartGameKey Then
    VR_Cab_StartButton.y = VR_Cab_StartButton.y +5
  End If

  If keycode = AddCreditKey or keycode = 4 then
    VR_Cab_StartButton.y = VR_Cab_StartButton.y +5
  End If
'**************************************************************

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
  BallsOnTable=BallsOnTable-1
  PlaySoundAtVol "DrainShort1", Drain, 1
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
  DOF 145, DOFPulse
End Sub

Sub Pause4Bonustimer_timer
  Pause4Bonustimer.enabled=0
  NextBallDelay.enabled=true

End Sub

'***********************
'     Flipper Logos
'***********************


Sub UpdateFlipperLogos_Timer
  LFLogo.RotZ = LeftFlipper.CurrentAngle
  RFLogo.RotZ = RightFlipper.CurrentAngle
  LFLogo1.RotZ = LeftFlipper.CurrentAngle
  RFLogo1.RotZ = RightFlipper.CurrentAngle
  PGate.Rotz = (Gate.CurrentAngle*.75) + 25
  FlipperLSh.RotZ = LeftFlipper.currentangle
  FlipperRSh.RotZ = RightFlipper.currentangle
End Sub

'***********************
' slingshots
'

Sub RightSlingShot_Slingshot
  PlaySoundAtVol SoundFXDOF("right_slingshot",105,DOFPulse,DOFContactors), ActiveBall, 1
  DOF 106, DOFPulse
    RSling0.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
  If RightSlingLight.state=1 then
    AddScore(10)
  else
    AddScore(1)
  end if
  AdvanceZeroToNine

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
  DOF 104, DOFPulse
    LSling0.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
  If LeftSlingLight.state=1 then
    AddScore(10)
  else
    AddScore(1)
  end if
  AdvanceZeroToNine

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
        AddScore(1)

    end Select
  end if
end Sub


'***********************************
' Bumpers
'***********************************


Sub Bumper1_Hit
  If TableTilted=false then

    PlaySoundAtVol SoundFXDOF("bumper1",108,DOFPulse,DOFContactors), ActiveBall, 1
    DOF 109, DOFPulse
    bump1 = 1
    if Bumper1Light.state=1 then
      AddScore(10)
    else
      AddScore(1)

    end if
    end if

End Sub

Sub Bumper2_Hit
  If TableTilted=false then

    PlaySoundAtVol SoundFXDOF("bumper1",110,DOFPulse,DOFContactors), ActiveBall, 1
    DOF 111, DOFPulse
    bump2 = 1
    if Bumper2Light.state=1 then
      AddScore(10)
    else
      AddScore(1)

    end if

    end if

End Sub

Sub Bumper3_Hit
  If TableTilted=false then

    PlaySoundAtVol SoundFXDOF("bumper1",112,DOFPulse,DOFContactors), ActiveBall, 1
    DOF 113, DOFPulse
    bump3 = 1
    if Bumper3Light.state=1 then
      AddScore(10)
    else
      AddScore(1)

    end if

    end if

End Sub

Sub Bumper4_Hit
  If TableTilted=false then

    PlaySoundAtVol SoundFXDOF("bumper1",114,DOFPulse,DOFContactors), ActiveBall, 1
    DOF 115, DOFPulse
    bump3 = 1
    if Bumper4Light.state=1 then
      AddScore(10)
    else
      AddScore(1)

    end if

    end if

End Sub


Sub Bumper5_Hit
  If TableTilted=false then

    PlaySoundAtVol SoundFXDOF("bumper1",116,DOFPulse,DOFContactors), ActiveBall, 1
    DOF 117, DOFPulse
    if Bumper5Light.state=1 then
      AddScore(10)
    else
      AddScore(1)

    end if

    end if

End Sub


Sub Bumper6_Hit
  If TableTilted=false then

    PlaySoundAtVol SoundFXDOF("bumper1",118,DOFPulse,DOFContactors), ActiveBall, 1
    DOF 119, DOFPulse
    if Bumper6Light.state=1 then
      AddScore(10)
    else
      AddScore(1)

    end if
    end if

End Sub

Sub Bumper7_Hit
  If TableTilted=false then

    PlaySoundAtVol SoundFXDOF("bumper1",120,DOFPulse,DOFContactors), ActiveBall, 1
    DOF 121, DOFPulse
    bump3 = 1
    if Bumper7Light.state=1 then
      AddScore(10)
    else
      AddScore(1)

    end if

    end if

End Sub

Sub BumperRubber1_hit
  Bumper5_Hit
end sub


Sub BumperRubber2_hit
  Bumper6_Hit
end sub

'****************************************************

sub CenterKicker1_Hit()
  Dim tempkickscore
  Dim speedx,speedy,finalspeed
  speedx=activeball.velx
  speedy=activeball.vely
  finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  if TableTilted=true then
    CenterKicker1.Kick 145,15
    tKickerTCount1=0
    exit sub
  end if
  if finalspeed>10 then
    CenterKicker1.Kick 0,0
    activeball.velx=speedx
    activeball.vely=speedy
  else

    CenterKicker1.timerenabled=1
    tKickerTCount1=0
  end if

end sub

Sub CenterKicker1_Timer()
  if MotorRunning<>0 then
    exit sub
  end if
  tKickerTCount1=tKickerTCount1+1
  select case tKickerTCount1
    case 1:
      Addscore(100)

    case 2:
      If LeftKickLight1.state=1 then
        AddScore(100)
      else
        Playsound "MotorPause"
      end if

    case 3:
      If LeftKickLight2.state=1 then
        AddScore(100)
      else
        Playsound "MotorPause"
      end if

    case 4:
      Playsound "MotorPause"

    case 6:
      if CenterKickLight1.state=1 then
        AddSpecial
      end if
    case 7:
    ' mHole.MagnetOn=0
      Pkickarm1.rotz=15
      PlaySoundAtVol SoundFXDOF("saucer",122,DOFPulse,DOFContactors), Pkickarm1, 1
      DOF 123, DOFPulse
      CenterKicker1.kick 142,15
    case 9:
      CenterKicker1.timerenabled=0
      Pkickarm1.rotz=0
  end Select

end sub


sub CenterKicker2_Hit()
  Dim tempkickscore
  Dim speedx,speedy,finalspeed
  speedx=activeball.velx
  speedy=activeball.vely
  finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  if TableTilted=true then
    CenterKicker2.Kick 215,15
    tKickerTCount2=0
    exit sub
  end if
  if finalspeed>10 then
    CenterKicker2.Kick 0,0
    activeball.velx=speedx
    activeball.vely=speedy
  else

    CenterKicker2.timerenabled=1
    tKickerTCount2=0
  end if

end sub

Sub CenterKicker2_Timer()
  if MotorRunning<>0 then
    exit sub
  end if
  tKickerTCount2=tKickerTCount2+1
  select case tKickerTCount2
    case 1:
      Addscore(100)

    case 2:
      If RightKickLight1.state=1 then
        AddScore(100)
      else
        Playsound "MotorPause"
      end if

    case 3:
      If RightKickLight2.state=1 then
        AddScore(100)
      else
        Playsound "MotorPause"
      end if

    case 4:
      Playsound "MotorPause"



    case 6:
      if CenterKickLight2.state=1 then
        AddSpecial
      end if
    case 7:
    ' mHole.MagnetOn=0
      Pkickarm2.rotz=15
      PlaySoundAtVol SoundFXDOF("saucer",124,DOFPulse,DOFContactors), Pkickarm2, 1
      DOF 125, DOFPulse
      CenterKicker2.kick 217,15
    case 10:
      CenterKicker2.timerenabled=0
      Pkickarm2.rotz=0
  end Select

end sub

'************************************
'  Trigger buttons
'************************************
Sub Trigger1_hit
  Button1.z=-1.5
  If TableTilted=false then
    If ButtonLight1.state=1 then
      AddScore(10)
    else
      AddScore(1)
    end if
  end if
end sub

Sub Trigger1_unhit
  Button1.z=.5
end sub

Sub Trigger2_hit
  Button2.z=-1.5
  If TableTilted=false then
    If ButtonLight2.state=1 then
      AddScore(10)
    else
      AddScore(1)
    end if
  end if
end sub

Sub Trigger2_unhit
  Button2.z=.5
end sub

Sub Trigger3_hit
  Button3.z=-1.5
  If TableTilted=false then
    If ButtonLight3.state=1 then
      AddScore(10)
    else
      AddScore(1)
    end if
  end if
end sub

Sub Trigger3_unhit
  Button3.z=.5
end sub

Sub Trigger4_hit
  Button4.z=-1.5
  If TableTilted=false then
    If UpperSpecialLight.state=1 then
      AddSpecial
    else
      AddScore(1)
    end if
  end if
end sub

Sub Trigger4_unhit
  Button4.z=.5
end sub

'************************************
'  Rollover lanes
'************************************

Sub TriggerCollection_Hit(idx)
  If TableTilted=false then
    DOF 130+idx,DOFPulse
    If idx<5 then BumperLights(idx).state=1:BumperLightsb(idx).state=1

    AllLightsOn(idx).state=0
    SetMotor(5)
    CheckSequences
  end if

end Sub



Sub CheckSequences
  If UpperLight1.state=0 AND UpperLight2.state=0 AND UpperLight3.state=0 AND UpperLight4.state=0 AND UpperLight5.state=0 then
    CompleteA=true
  else
    CompleteA=false
  end if

  If YellowRolloverLight1.state=0 AND YellowRolloverLight2.state=0 AND YellowRolloverLight3.state=0 AND YellowRolloverLight4.state=0 then
    CompleteK=true
  else
    CompleteK=false
  end if

  If GreenRolloverLight1.state=0 AND GreenRolloverLight2.state=0 AND GreenRolloverLight3.state=0 AND GreenRolloverLight4.state=0 then
    CompleteQ=true
  else
    CompleteQ=false
  end if
  If Bumper1Light.State=1 Then Bumper1Lightb.State=1:End If
  If Bumper2Light.State=1 Then Bumper2Lightb.State=1:End If
  If Bumper3Light.State=1 Then Bumper3Lightb.State=1:End If
  If Bumper4Light.State=1 Then Bumper4Lightb.State=1:End If
  If Bumper5Light.State=1 Then Bumper5Lightb.State=1:End If
  If Bumper1Light.State=0 Then Bumper1Lightb.State=0:End If
  If Bumper2Light.State=0 Then Bumper2Lightb.State=0:End If
  If Bumper3Light.State=0 Then Bumper3Lightb.State=0:End If
  If Bumper4Light.State=0 Then Bumper4Lightb.State=0:End If
  If Bumper5Light.State=0 Then Bumper5Lightb.State=0:End If




  CheckLightSequences


end sub

sub CheckLightSequences
  If CompleteK=true then
    CenterKickLight1.state=1
    LeftKickLight1.state=1
    RightKickLight1.state=1
  else
    CenterKickLight1.state=0
    LeftKickLight1.state=0
    RightKickLight1.state=0
  end if

  If CompleteQ=true then
    CenterKickLight2.state=1
    LeftKickLight2.state=1
    RightKickLight2.state=1
  else
    CenterKickLight2.state=0
    LeftKickLight2.state=0
    RightKickLight2.state=0
  end if

  If CompleteA=true AND CompleteK=true AND CompleteQ=true then
    SpecialLight.state=1
  else
    SpecialLight.state=0
  end if
  LeftSlingLight.state=0
  RightSlingLight.state=0
  ButtonLight1.state=0
  ButtonLight3.state=0
  UpperSpecialLight.state=0
  Select case ZeroToNineUnit
    case 0,2,4,6,8:
      ButtonLight1.state=1
      RightSlingLight.state=1
    case 1,3,5,7,9:
      ButtonLight3.state=1
      LeftSlingLight.state=1
  end select
  Select case KickerLightCounter
    case 0:
      If AdvanceSpecialMax=3 AND CompleteA=true then
        UpperSpecialLight.state=1
      end if
    case 1:
      If (AdvanceSpecialMax=2 OR AdvanceSpecialMax=3) AND (CompleteA=true) then
        UpperSpecialLight.state=1
      end if
    case 2:
      If CompleteA=true then
        UpperSpecialLight.state=1
      end if
    case 3:
      If AdvanceSpecialMax=3 AND CompleteK=true then
        UpperSpecialLight.state=1
      end if
    case 4:
      If (AdvanceSpecialMax=2 OR AdvanceSpecialMax=3) AND (CompleteK=true) then
        UpperSpecialLight.state=1
      end if
    case 5:
      If CompleteK=true then
        UpperSpecialLight.state=1
      end if
    case 6:
      If AdvanceSpecialMax=3 AND CompleteQ=true then
        UpperSpecialLight.state=1
      end if
    case 7:
      If (AdvanceSpecialMax=2 OR AdvanceSpecialMax=3) AND (CompleteQ=true) then
        UpperSpecialLight.state=1
      end if
    case 8:
      If CompleteQ=true then
        UpperSpecialLight.state=1
      end if
  end select
End Sub
'


Sub AddSpecial()
  PlaySound SoundFXDOF("knocker",133,DOFPulse,DOFContactors)
  DOF 134, DOFPulse
  Credits=Credits+1
  DOF 135, DOFOn
  if Credits>9 then Credits=9
  If B2SOn Then
    Controller.B2SSetCredits Credits
  End If

  If Renderingmode = 2 Then
    cred =reels(4, 0)
    reels(4, 0) = 0
    SetDrum -1,0,  0

    SetReel 0,-1,  Credits
    reels(4, 0) = Credits
  End If
  CreditsReel.SetValue(Credits)
End Sub

Sub AddSpecial2()
  PlaySound"click"
  Credits=Credits+1
  DOF 135, DOFOn
  if Credits>9 then Credits=9
  If B2SOn Then
    Controller.B2SSetCredits Credits
  End If
  If Renderingmode = 2 Then
    cred =reels(4, 0)
    reels(4, 0) = 0
    SetDrum -1,0,  0

    SetReel 0,-1,  Credits
    reels(4, 0) = Credits
  End If
  CreditsReel.SetValue(Credits)
End Sub



Sub AdvanceZeroToNine
  ZeroToNineUnit=ZeroToNineUnit+1
  if ZeroToNineUnit>9 then
    ZeroToNineUnit=0
  end if

  BumperLightCounter=BumperLightCounter+1
  if BumperLightCounter>2 then
    BumperLightCounter=0
  end if


  LightBumpers
  KickerLightCounter=KickerLightCounter+1
  if KickerLightCounter>8 then
    KickerLightCounter=0
  end if
  CheckLightSequences

end sub

Sub LightBumpers


end sub

Sub ToggleAlternatingRelay

end sub

Sub ResetBallDrops

  HoleCounter=0

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
    DOF 136,DOFPulse
    BallsOnTable=BallsOnTable+1
  ' BallInPlayReel.SetValue(BallInPlay)
    InstructCard.image="IC"+FormatNumber(BallsPerGame,0)
  end if


End Sub





sub resettimer_timer
    rst=rst+1
  if rst>1 and rst<12 then
    ResetReelsToZero(1)
  end if
    if rst=13 then
    if BallLiftOption=2 then
      playsound "StartBall1"
    end if
    end if
    if rst=14 then
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
    If Renderingmode = 2 Then
      EMMODE = 1
      UpdateVRReels 0,0 ,score(0), 0, -1,-1,-1,-1,-1
      UpdateVRReels 1,1 ,score(1), 0, -1,-1,-1,-1,-1
      UpdateVRReels 2,2 ,score(2), 0, -1,-1,-1,-1,-1
      UpdateVRReels 3,3 ,score(3), 0, -1,-1,-1,-1,-1
      UpdateVRReels 4,4 ,score(4), 0, -1,-1,-1,-1,-1
      EMMODE = 0 ' restore EM mode
    End If
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
  for i = 1 to 4
    Score(i)=0
    Score100K(1)=0
    HighScorePaid(i)=false
    Replay1Paid(i)=false
    Replay2Paid(i)=false
    Replay3Paid(i)=false
  next

  If Renderingmode = 2 Then
    VRGameOver.visible = 0
    VRTilt.visible = 0
    VRBG1k.visible = 0
    for each Object in VRBGMatch : object.visible = 0 : next
    UpdateVRReels  0 ,0 ,Score(0), 0, 0,0,0,0,0
    UpdateVRReels  1 ,1 ,Score(0), 0, 0,0,0,0,0
    UpdateVRReels  2 ,2 ,Score(0), 0, 0,0,0,0,0
    UpdateVRReels  3 ,3 ,Score(0), 0, 0,0,0,0,0
  End If
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
  for each obj in AllLightsOn
    obj.state=1
  next
  for each obj in AllLightsOff
    obj.state=0
  next


  CompleteA=false
  CompleteK=false
  CompleteQ=false
  CompleteJ=false
  CompleteClubs=false

  KickerLightCounter=0
  BumperLightCounter=0
  Bumper1Light.state=0
  Bumper2Light.state=0
  Bumper3Light.state=0
  Bumper4Light.state=0
  Bumper5Light.state=0
  Bumper6Light.state=1
  Bumper7Light.state=1
  Bumper1Lightb.state=0
  Bumper2Lightb.state=0
  Bumper3Lightb.state=0
  Bumper4Lightb.state=0
  Bumper5Lightb.state=0

  AlternatingRelay=0
  CheckLightSequences
  BumpersOn
  ABCDComplete=false
  KickerLightCounter=0
  TargetSequenceComplete=0
  AdvancesCompleted=0
  LightBumpers
  SpecialLight.state=0
' IncreaseBonus
' ToggleBumper
  EightLit=1
  ResetBalls
End sub

sub EndOfGameWaiter2_timer
  If MotorRunning<>0 then
    exit sub
  end if


  EndOfGameWaiter2.enabled=false

  EndOfGame
end sub


sub EndOfGame

  InProgress=false

  If B2SOn Then
    Controller.B2SSetGameOver 1
    Controller.B2SSetPlayerUp 0
    Controller.B2SSetBallInPlay 0
    Controller.B2SSetCanPlay 0
  End If
  If Renderingmode = 2 Then
    VRGameOver.visible = 1
  End If
  If Table1.ShowDT = True Then
    PlayerScores(0).visible=1
    PlayerScoresOn(0).visible=0
    For each obj in PlayerHuds
      obj.SetValue(0)
    next
    GameOverReel.SetValue(1)

  end If

  Debug.Print(Score(1))
  If ScorbitActive = 1 Then
    Scorbit.StopSession Score(1),Score(2),Score(3),Score(4),players
  End If
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
  TimerTitle1.enabled=1
  TimerTitle2.enabled=1
  TimerTitle3.enabled=1
  TimerTitle4.enabled=1
  TimerTitle5.enabled=1
  HighScoreTimer.interval=100
  HighScoreTimer.enabled=True

end sub

sub nextball
  If BumperSequenceCompleted=1 then
    BumperSequenceCompleted=0
    LeftSpeshulLight.state=0
    RightSpeshulLight.state=0
    ResetBumperSequence
  end if
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
      EndOfGameWaiter2.enabled=true
      'EndOfGameWaiter2.enabled=true
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
          obj.visible=0
        Next
        For each obj in PlayerScoresOn
          obj.visible=1
        next

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
        obj.visible=0
      Next
      For each obj in PlayerScoresOn
        obj.visible=1
      next

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
  tempmatch=Int(Rnd*10)
  Match=tempmatch
  If TableTilted=True and TiltEndsGame then
    exit sub
  end if
  MatchReel.SetValue(tempmatch+1)

  If B2SOn Then
    If Match = 0 Then
      Controller.B2SSetMatch 10
    Else
      Controller.B2SSetMatch Match
    End If
  End if
  If Renderingmode = 2 Then
    FlasherMatch
  End If
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
      If Renderingmode = 2 Then
        VRTilt.visible = 1
      End If
      If TiltEndsGame=1 then
        InProgress=false
        If B2SOn Then
          Controller.B2SSetGameOver 1
          Controller.B2SSetPlayerUp 0
          Controller.B2SSetBallInPlay 0
          Controller.B2SSetCanPlay 0
        End If
        If Renderingmode = 2 Then
          VRGameOver.visible = 0
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
    scorefile.writeline AdvanceBonusMax
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
    AdvanceBonusMax=cdbl(temp4)
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

sub InitPauser5_timer
    If B2SOn Then
      For i = 90 to 93
        Controller.B2SSetData i,0
      next
      For i = 61 to 80
        Controller.B2SSetData i,0
      next
      Controller.B2SSetData 90,1
      Controller.B2SSetData 61,1
    End If
    DisplayHighScore
    CreditsReel.SetValue(Credits)
    InitPauser5.enabled=false
end sub



sub BumpersOff
  Bumper1Light.Visible=0
  Bumper2Light.Visible=0
  Bumper3Light.Visible=0
  Bumper4Light.Visible=0
  Bumper5Light.Visible=0
  Bumper6Light.Visible=0
  Bumper7Light.Visible=0
  Bumper1Lightb.Visible=0
  Bumper2Lightb.Visible=0
  Bumper3Lightb.Visible=0
  Bumper4Lightb.Visible=0
  Bumper5Lightb.Visible=0
  GILight033.Visible=0
  GILight034.Visible=0
  CenterKickLight1.visible=0
  CenterKickLight2.visible=0
end sub

sub BumpersOn


  Bumper1Light.Visible=1
  Bumper2Light.Visible=1
  Bumper3Light.Visible=1
  Bumper4Light.Visible=1
  Bumper5Light.Visible=1
  Bumper6Light.Visible=1
  Bumper7Light.Visible=1
  Bumper1Lightb.Visible=1
  Bumper2Lightb.Visible=1
  Bumper3Lightb.Visible=1
  Bumper4Lightb.Visible=1
  Bumper5Lightb.Visible=1
  GILight033.Visible=1
  GILight034.Visible=1


  CenterKickLight1.visible=1
  CenterKickLight2.visible=1


end sub

Sub PlasticsOn

  For each obj in Flashers
    obj.Visible=1
  next




end sub

Sub PlasticsOff

  For each obj in Flashers
    obj.Visible=0
  next

  StopSound "buzz"
  StopSound "buzzL"

end sub

Sub SetupReplayTables

  Replay1Table(1)=800
  Replay1Table(2)=900
  Replay1Table(3)=1100
  Replay1Table(4)=1000
  Replay1Table(5)=1000
  Replay1Table(6)=1100
  Replay1Table(7)=1200
  Replay1Table(8)=1200
  Replay1Table(9)=1300
  Replay1Table(10)=1300
  Replay1Table(11)=500
  Replay1Table(12)=500
  Replay1Table(13)=600
  Replay1Table(14)=600
  Replay1Table(15)=600
  Replay1Table(16)=600
  Replay1Table(17)=700
  Replay1Table(18)=800
  Replay1Table(19)=800
  Replay1Table(20)=800

  Replay2Table(1)=1000
  Replay2Table(2)=1100
  Replay2Table(3)=1200
  Replay2Table(4)=1200
  Replay2Table(5)=1300
  Replay2Table(6)=1300
  Replay2Table(7)=1300
  Replay2Table(8)=1400
  Replay2Table(9)=1400
  Replay2Table(10)=1500
  Replay2Table(11)=700
  Replay2Table(12)=700
  Replay2Table(13)=700
  Replay2Table(14)=800
  Replay2Table(15)=800
  Replay2Table(16)=900
  Replay2Table(17)=900
  Replay2Table(18)=1000
  Replay2Table(19)=1100
  Replay2Table(20)=1100

  Replay3Table(1)=1200
  Replay3Table(2)=1300
  Replay3Table(3)=1300
  Replay3Table(4)=1400
  Replay3Table(5)=1400
  Replay3Table(6)=1400
  Replay3Table(7)=1400
  Replay3Table(8)=1500
  Replay3Table(9)=1500
  Replay3Table(10)=1600
  Replay3Table(11)=800
  Replay3Table(12)=900
  Replay3Table(13)=800
  Replay3Table(14)=900
  Replay3Table(15)=1000
  Replay3Table(16)=1100
  Replay3Table(17)=1100
  Replay3Table(18)=1200
  Replay3Table(19)=1200
  Replay3Table(20)=1300

  Replay4Table(1)=1400
  Replay4Table(2)=1500
  Replay4Table(3)=1400
  Replay4Table(4)=1600
  Replay4Table(5)=1500
  Replay4Table(6)=1500
  Replay4Table(7)=1500
  Replay4Table(8)=1600
  Replay4Table(9)=1600
  Replay4Table(10)=1700
  Replay4Table(11)=900
  Replay4Table(12)=1000
  Replay4Table(13)=900
  Replay4Table(14)=1000
  Replay4Table(15)=1200
  Replay4Table(16)=1300
  Replay4Table(17)=1300
  Replay4Table(18)=1400
  Replay4Table(19)=1400
  Replay4Table(20)=1500

  ReplayTableMax=2


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
      Case 1:
        AddScore(1)
        MotorRunning=0
        BumpersOn

      Case 2:
        MotorMode=1
        MotorPosition=2
        BumpersOff
      Case 3:
        MotorMode=1
        MotorPosition=3
        BumpersOff
      Case 4:
        MotorMode=1
        MotorPosition=4
        BumpersOff
      Case 5:
        MotorMode=1
        MotorPosition=5
        BumpersOff

      Case 10:
        AddScore(10)
        MotorRunning=0
        BumpersOn

      Case 20:
        MotorMode=10
        MotorPosition=2
        BumpersOff
      Case 30:
        MotorMode=10
        MotorPosition=3
        BumpersOff
      Case 40:
        MotorMode=10
        MotorPosition=4
        BumpersOff
      Case 50:
        MotorMode=10
        MotorPosition=5
        BumpersOff
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
    AdvanceZeroToNine

    end if

    OldScore = int(OldScore / 10)
    NewScore = int(NewScore / 10)
  ' MsgBox("OldScore="&OldScore&", NewScore="&NewScore)
    if (OldScore Mod 10 <> NewScore Mod 10) then
    PlayChime(100)
    HorseTimer.enabled=true

    end if

    OldScore = int(OldScore / 10)
    NewScore = int(NewScore / 10)
  ' MsgBox("OldScore="&OldScore&", NewScore="&NewScore)
    if (OldScore Mod 10 <> NewScore Mod 10) then
    PlayChime(1000)
    end if
  If Score(1)>999 then
    Reel1000.SetValue(1)
    If Renderingmode = 2 Then VRBG1k.visible = 1
  end if

  If B2SOn Then
    Controller.B2SSetScorePlayer 1, Score(1)
    If Score(1)>999 then
      Controller.B2SSetData 98,0

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
          PlaySound SoundFXDOF("SpinACard_1_10_Point_Bell",137,DOFPulse,DOFChimes)
          LastChime10=0
        Else
          'PlaySound "SpinACard_1_10_Point_Bell"
          PlaySound SoundFXDOF("SpinACard_1_10_Point_Bell",137,DOFPulse,DOFChimes)
          LastChime10=1
        End If
      Case 100
        If LastChime100=1 Then
          'PlaySound "SpinACard_100_Point_Bell"
          PlaySound SoundFXDOF("SpinACard_100_Point_Bell",138,DOFPulse,DOFChimes)
          LastChime100=0
        Else
          'PlaySound "SpinACard_100_Point_Bell"
          PlaySound SoundFXDOF("SpinACard_100_Point_Bell",138,DOFPulse,DOFChimes)
          LastChime100=1
        End If

    End Select

    If Renderingmode = 2 Then
      EMMODE = 1
      UpdateVRReels 0,0 ,score(0), 0, -1,-1,-1,-1,-1
      UpdateVRReels 1,1 ,score(1), 0, -1,-1,-1,-1,-1
      UpdateVRReels 2,2 ,score(2), 0, -1,-1,-1,-1,-1
      UpdateVRReels 3,3 ,score(3), 0, -1,-1,-1,-1,-1
      UpdateVRReels 4,4 ,score(4), 0, -1,-1,-1,-1,-1
      EMMODE = 0 ' restore EM mode
    End If
  else
    Select Case x
      Case 10
        If LastChime10=1 Then
          PlaySound SoundFXDOF("SJ_Chime_10a",137,DOFPulse,DOFChimes)
          LastChime10=0
        Else
          PlaySound SoundFXDOF("SJ_Chime_10b",137,DOFPulse,DOFChimes)
          LastChime10=1
        End If
      Case 100
        If LastChime100=1 Then
          PlaySound SoundFXDOF("SJ_Chime_100a",138,DOFPulse,DOFChimes)
          LastChime100=0
        Else
          PlaySound SoundFXDOF("SJ_Chime_100b",138,DOFPulse,DOFChimes)
          LastChime100=1
        End If
      Case 1000
        If LastChime1000=1 Then
          PlaySound SoundFXDOF("SJ_Chime_1000a",139,DOFPulse,DOFChimes)
          LastChime1000=0
        Else
          PlaySound SoundFXDOF("SJ_Chime_1000b",139,DOFPulse,DOFChimes)
          LastChime1000=1
        End If
    End Select
    If Renderingmode = 2 Then
      EMMODE = 1
      UpdateVRReels 0,0 ,score(0), 0, -1,-1,-1,-1,-1
      UpdateVRReels 1,1 ,score(1), 0, -1,-1,-1,-1,-1
      UpdateVRReels 2,2 ,score(2), 0, -1,-1,-1,-1,-1
      UpdateVRReels 3,3 ,score(3), 0, -1,-1,-1,-1,-1
      UpdateVRReels 4,4 ,score(4), 0, -1,-1,-1,-1,-1
      EMMODE = 0 ' restore EM mode
    End If
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

Sub Metals_Hit (idx)
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

'Sub LeftFlipper_Collide(parm)
'   RandomSoundFlipper()
'End Sub

'Sub RightFlipper_Collide(parm)
'   RandomSoundFlipper()
'End Sub

Sub RandomSoundFlipper()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "Flipper_Rubber_1", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "Flipper_Rubber_2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "Flipper_Rubber_3", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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
HSScore(1) = 1500
HSScore(2) = 1400
HSScore(3) = 1300
HSScore(4) = 1000
HSScore(5) = 900

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
    Eval("HS"&xfor).imageA = GetHSChar(String, Index)
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
Const OptionLinesToMark="011101111"
Const OptionLine1="" 'do not use this line
Const OptionLine2="" 'do not use this line
Const OptionLine3="" 'do not use this line
Const OptionLine4="Special Lighting Option"
Const OptionLine5=""
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
  OperatorMenuBackdrop.imageA = "OperatorMenu"

  OptionCHS = 0
  CurrentOption = 1
  DisplayAllOptions
  OperatorOption1.imageA = "BluePlus"
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
        Select case AdvanceSpecialMax
          case 1:
            tempstring = "Liberal"
          case 2:
            tempstring = "Moderate"
          case 3:
            tempstring = "Conservative"
        end select

        SetOptLine 5, tempstring
      Case 5:
        SetOptLine 6, tempstring

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
      Eval("OperatorOption"&XOpt).imageA = "PostitBL"
    next
    MoveArrow
    if CurrentOption<8 then
      Eval("OperatorOption"&CurrentOption).imageA = "BluePlus"
    elseif CurrentOption=8 then
      Eval("OperatorOption"&CurrentOption).imageA = "GreenCheck"
    else
      Eval("OperatorOption"&CurrentOption).imageA = "RedX"
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

    elseif CurrentOption = 4 then
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
          HSScore(1) = 1400
          HSScore(2) = 1200
          HSScore(3) = 1100
          HSScore(4) = 1000
          HSScore(5) = 800

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
        OperatorMenuBackdrop.imageA = "PostitBL"
        For XOpt = 1 to MaxOption
          Eval("OperatorOption"&XOpt).imageA = "PostitBL"
        next

        For XOpt = 1 to 256
          Eval("Option"&XOpt).imageA = "PostItBL"
        next
        RefreshReplayCard
        InstructCard.image="IC"+FormatNumber(BallsPerGame,0)
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
    Eval("Option"&xfor).imageA = GetOptChar(string, Index)
    Index = Index + 1
  next
  LineLengths(LineNo) = StrLen

End Sub


'*****************************************
'Flipper Polarity 70s to 80s
'*****************************************

dim LF : Set LF = New FlipperPolarity
dim RF : Set RF = New FlipperPolarity

InitPolarity

Sub InitPolarity()
        dim x, a : a = Array(LF, RF)
        for each x in a
                x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1        'disabled
                x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1
                x.enabled = True
                x.TimeDelay = 80
        Next

        AddPt "Polarity", 0, 0, 0
        AddPt "Polarity", 1, 0.05, -2.7
        AddPt "Polarity", 2, 0.33, -2.7
        AddPt "Polarity", 3, 0.37, -2.7
        AddPt "Polarity", 4, 0.41, -2.7
        AddPt "Polarity", 5, 0.45, -2.7
        AddPt "Polarity", 6, 0.576,-2.7
        AddPt "Polarity", 7, 0.66, -1.8
        AddPt "Polarity", 8, 0.743, -0.5
        AddPt "Polarity", 9, 0.81, -0.5
        AddPt "Polarity", 10, 0.88, 0

        addpt "Velocity", 0, 0,         1
        addpt "Velocity", 1, 0.16, 1.06
        addpt "Velocity", 2, 0.41,         1.05
        addpt "Velocity", 3, 0.53,         1'0.982
        addpt "Velocity", 4, 0.702, 0.968
        addpt "Velocity", 5, 0.95,  0.968
        addpt "Velocity", 6, 1.03,         0.945

        LF.Object = LeftFlipper
        LF.EndPoint = EndPointLp
        RF.Object = RightFlipper
        RF.EndPoint = EndPointRp
End Sub

Sub TriggerLF_Hit() : LF.Addball activeball : End Sub
Sub TriggerLF_UnHit() : LF.PolarityCorrect activeball : End Sub
Sub TriggerRF_Hit() : RF.Addball activeball : End Sub
Sub TriggerRF_UnHit() : RF.PolarityCorrect activeball : End Sub


'******************************************************
'                        FLIPPER CORRECTION FUNCTIONS
'******************************************************

Sub AddPt(aStr, idx, aX, aY)        'debugger wrapper for adjusting flipper script in-game
        dim a : a = Array(LF, RF)
        dim x : for each x in a
                x.addpoint aStr, idx, aX, aY
        Next
End Sub

Class FlipperPolarity
        Public DebugOn, Enabled
        Private FlipAt        'Timer variable (IE 'flip at 723,530ms...)
        Public TimeDelay        'delay before trigger turns off and polarity is disabled TODO set time!
        private Flipper, FlipperStart,FlipperEnd, FlipperEndY, LR, PartialFlipCoef
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
        Public Property Let EndPoint(aInput) : FlipperEnd = aInput.x: FlipperEndY = aInput.y: End Property
        Public Property Get EndPoint : EndPoint = FlipperEnd : End Property
        Public Property Get EndPointY: EndPointY = FlipperEndY : End Property

        Public Sub AddPoint(aChooseArray, aIDX, aX, aY) 'Index #, X position, (in) y Position (out)
                Select Case aChooseArray
                        case "Polarity" : ShuffleArrays PolarityIn, PolarityOut, 1 : PolarityIn(aIDX) = aX : PolarityOut(aIDX) = aY : ShuffleArrays PolarityIn, PolarityOut, 0
                        Case "Velocity" : ShuffleArrays VelocityIn, VelocityOut, 1 :VelocityIn(aIDX) = aX : VelocityOut(aIDX) = aY : ShuffleArrays VelocityIn, VelocityOut, 0
                        Case "Ycoef" : ShuffleArrays YcoefIn, YcoefOut, 1 :YcoefIn(aIDX) = aX : YcoefOut(aIDX) = aY : ShuffleArrays YcoefIn, YcoefOut, 0
                End Select
                if gametime > 100 then Report aChooseArray
        End Sub

        Public Sub Report(aChooseArray)         'debug, reports all coords in tbPL.text
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
                        End If
                Next
                PartialFlipCoef = ((Flipper.StartAngle - Flipper.CurrentAngle) / (Flipper.StartAngle - Flipper.EndAngle))
                PartialFlipCoef = abs(PartialFlipCoef-1)
        End Sub
        Private Function FlipperOn() : if gameTime < FlipAt+TimeDelay then FlipperOn = True : End If : End Function        'Timer shutoff for polaritycorrect

        Public Sub PolarityCorrect(aBall)
                if FlipperOn() then
                        dim tmp, BallPos, x, IDX, Ycoef : Ycoef = 1

                        'y safety Exit
                        if aBall.VelY > -8 then 'ball going down
                                RemoveBall aBall
                                exit Sub
                        end if

                        'Find balldata. BallPos = % on Flipper
                        for x = 0 to uBound(Balls)
                                if aBall.id = BallData(x).id AND not isempty(BallData(x).id) then
                                        idx = x
                                        BallPos = PSlope(BallData(x).x, FlipperStart, 0, FlipperEnd, 1)
                                        if ballpos > 0.65 then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)                                'find safety coefficient 'ycoef' data
                                end if
                        Next

                        If BallPos = 0 Then 'no ball data meaning the ball is entering and exiting pretty close to the same position, use current values.
                                BallPos = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
                                if ballpos > 0.65 then  Ycoef = LinearEnvelope(aBall.Y, YcoefIn, YcoefOut)                                                'find safety coefficient 'ycoef' data
                        End If

                        'Velocity correction
                        if not IsEmpty(VelocityIn(0) ) then
                                Dim VelCoef
         :                         VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)

                                if partialflipcoef < 1 then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)

                                if Enabled then aBall.Velx = aBall.Velx*VelCoef
                                if Enabled then aBall.Vely = aBall.Vely*VelCoef
                        End If

                        'Polarity Correction (optional now)
                        if not IsEmpty(PolarityIn(0) ) then
                                If StartPoint > EndPoint then LR = -1        'Reverse polarity if left flipper
                                dim AddX : AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR

                                if Enabled then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef)
                                'playsound "fx_knocker"
                        End If
                End If
                RemoveBall aBall
        End Sub
End Class

'******************************************************
'                FLIPPER POLARITY AND RUBBER DAMPENER
'                        SUPPORTING FUNCTIONS
'******************************************************

' Used for flipper correction and rubber dampeners
Sub ShuffleArray(ByRef aArray, byVal offset) 'shuffle 1d array
        dim x, aCount : aCount = 0
        redim a(uBound(aArray) )
        for x = 0 to uBound(aArray)        'Shuffle objects in a temp array
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
        redim aArray(aCount-1+offset)        'Resize original array
        for x = 0 to aCount-1                'set objects back into original array
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
Function PSlope(Input, X1, Y1, X2, Y2)        'Set up line via two points, no clamping. Input X, output Y
        dim x, y, b, m : x = input : m = (Y2 - Y1) / (X2 - X1) : b = Y2 - m*X2
        Y = M*x+b
        PSlope = Y
End Function

' Used for flipper correction
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

' Used for flipper correction and rubber dampeners
Function LinearEnvelope(xInput, xKeyFrame, yLvl)
        dim y 'Y output
        dim L 'Line
        dim ii : for ii = 1 to uBound(xKeyFrame)        'find active line
                if xInput <= xKeyFrame(ii) then L = ii : exit for : end if
        Next
        if xInput > xKeyFrame(uBound(xKeyFrame) ) then L = uBound(xKeyFrame)        'catch line overrun
        Y = pSlope(xInput, xKeyFrame(L-1), yLvl(L-1), xKeyFrame(L), yLvl(L) )

        if xInput <= xKeyFrame(lBound(xKeyFrame) ) then Y = yLvl(lBound(xKeyFrame) )         'Clamp lower
        if xInput >= xKeyFrame(uBound(xKeyFrame) ) then Y = yLvl(uBound(xKeyFrame) )        'Clamp upper

        LinearEnvelope = Y
End Function



'******************************************************
'                        FLIPPER TRICKS
'******************************************************

RightFlipper.timerinterval=1
Rightflipper.timerenabled=True

sub RightFlipper_timer()
        FlipperTricks LeftFlipper, LFPress, LFCount, LFEndAngle, LFState
        FlipperTricks RightFlipper, RFPress, RFCount, RFEndAngle, RFState
        FlipperNudge RightFlipper, RFEndAngle, RFEOSNudge, LeftFlipper, LFEndAngle
        FlipperNudge LeftFlipper, LFEndAngle, LFEOSNudge,  RightFlipper, RFEndAngle
end sub

Dim LFEOSNudge, RFEOSNudge

Sub FlipperNudge(Flipper1, Endangle1, EOSNudge1, Flipper2, EndAngle2)
        Dim BOT, b

        If Flipper1.currentangle = Endangle1 and EOSNudge1 <> 1 Then
                EOSNudge1 = 1
                'debug.print Flipper1.currentangle &" = "& Endangle1 &"--"& Flipper2.currentangle &" = "& EndAngle2
                If Flipper2.currentangle = EndAngle2 Then
                        BOT = GetBalls
                        For b = 0 to Ubound(BOT)
                                If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper1) Then
                                        'Debug.Print "ball in flip1. exit"
                                         exit Sub
                                end If
                        Next
                        For b = 0 to Ubound(BOT)
                                If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper2) Then
                                        BOT(b).velx = BOT(b).velx / 1.7
                                        BOT(b).vely = BOT(b).vely - 1
                                end If
                        Next
                End If
        Else
                If Flipper1.currentangle <> EndAngle1 then
                        EOSNudge1 = 0
                end if
        End If
End Sub

'*****************
' Maths
'*****************
'Const PI = 3.1415927
'
'Function dSin(degrees)
'        dsin = sin(degrees * Pi/180)
'End Function
'
'Function dCos(degrees)
'        dcos = cos(degrees * Pi/180)
'End Function

'*************************************************
' Check ball distance from Flipper for Rem
'*************************************************

Function Distance(ax,ay,bx,by)
        Distance = SQR((ax - bx)^2 + (ay - by)^2)
End Function

Function DistancePL(px,py,ax,ay,bx,by) ' Distance between a point and a line where point is px,py
        DistancePL = ABS((by - ay)*px - (bx - ax) * py + bx*ay - by*ax)/Distance(ax,ay,bx,by)
End Function

Function Radians(Degrees)
        Radians = Degrees * PI /180
End Function

Function AnglePP(ax,ay,bx,by)
        AnglePP = Atn2((by - ay),(bx - ax))*180/PI
End Function

Function DistanceFromFlipper(ballx, bally, Flipper)
        DistanceFromFlipper = DistancePL(ballx, bally, Flipper.x, Flipper.y, Cos(Radians(Flipper.currentangle+90))+Flipper.x, Sin(Radians(Flipper.currentangle+90))+Flipper.y)
End Function

Function FlipperTrigger(ballx, bally, Flipper)
        Dim DiffAngle
        DiffAngle  = ABS(Flipper.currentangle - AnglePP(Flipper.x, Flipper.y, ballx, bally) - 90)
        If DiffAngle > 180 Then DiffAngle = DiffAngle - 360

        If DistanceFromFlipper(ballx,bally,Flipper) < 48 and DiffAngle <= 90 and Distance(ballx,bally,Flipper.x,Flipper.y) < Flipper.Length Then
                FlipperTrigger = True
        Else
                FlipperTrigger = False
        End If
End Function

' Used for drop targets and stand up targets
Function Atn2(dy, dx)
        dim pi
        pi = 4*Atn(1)

        If dx > 0 Then
                Atn2 = Atn(dy / dx)
        ElseIf dx < 0 Then
                If dy = 0 Then
                        Atn2 = pi
                Else
                        Atn2 = Sgn(dy) * (pi - Atn(Abs(dy / dx)))
                end if
        ElseIf dx = 0 Then
                if dy = 0 Then
                        Atn2 = 0
                else
                        Atn2 = Sgn(dy) * pi / 2
                end if
        End If
End Function

'*************************************************
' End - Check ball distance from Flipper for Rem
'*************************************************

dim LFPress, RFPress, LFCount, RFCount
dim LFState, RFState
dim EOST, EOSA,Frampup, FElasticity,FReturn
dim RFEndAngle, LFEndAngle

EOST = leftflipper.eostorque
EOSA = leftflipper.eostorqueangle
Frampup = LeftFlipper.rampup
FElasticity = LeftFlipper.elasticity
FReturn = LeftFlipper.return
Const EOSTnew = 1
Const EOSAnew = 1
Const EOSRampup = 0
Dim SOSRampup
SOSRampup = 2.5
Const LiveCatch = 16
Const LiveElasticity = 0.45
Const SOSEM = 0.815
Const EOSReturn = 0.055

LFEndAngle = Leftflipper.endangle
RFEndAngle = RightFlipper.endangle

Sub FlipperActivate(Flipper, FlipperPress)
        FlipperPress = 1
        Flipper.Elasticity = FElasticity

        Flipper.eostorque = EOST
        Flipper.eostorqueangle = EOSA
End Sub

Sub FlipperDeactivate(Flipper, FlipperPress)
        FlipperPress = 0
        Flipper.eostorqueangle = EOSA
        Flipper.eostorque = EOST*EOSReturn/FReturn


        If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 0.1 Then
                Dim BOT, b
                BOT = GetBalls

                For b = 0 to UBound(BOT)
                        If Distance(BOT(b).x, BOT(b).y, Flipper.x, Flipper.y) < 55 Then 'check for cradle
                                If BOT(b).vely >= -0.4 Then BOT(b).vely = -0.4
                        End If
                Next
        End If
End Sub

Sub FlipperTricks (Flipper, FlipperPress, FCount, FEndAngle, FState)
        Dim Dir
        Dir = Flipper.startangle/Abs(Flipper.startangle)        '-1 for Right Flipper

        If Abs(Flipper.currentangle) > Abs(Flipper.startangle) - 0.05 Then
                If FState <> 1 Then
                        Flipper.rampup = SOSRampup
                        Flipper.endangle = FEndAngle - 3*Dir
                        Flipper.Elasticity = FElasticity * SOSEM
                        FCount = 0
                        FState = 1
                End If
        ElseIf Abs(Flipper.currentangle) <= Abs(Flipper.endangle) and FlipperPress = 1 then
                if FCount = 0 Then FCount = GameTime

                If FState <> 2 Then
                        Flipper.eostorqueangle = EOSAnew
                        Flipper.eostorque = EOSTnew
                        Flipper.rampup = EOSRampup
                        Flipper.endangle = FEndAngle
                        FState = 2
                End If
        Elseif Abs(Flipper.currentangle) > Abs(Flipper.endangle) + 0.01 and FlipperPress = 1 Then
                If FState <> 3 Then
                        Flipper.eostorque = EOST
                        Flipper.eostorqueangle = EOSA
                        Flipper.rampup = Frampup
                        Flipper.Elasticity = FElasticity
                        FState = 3
                End If

        End If
End Sub

Const LiveDistanceMin = 30  'minimum distance in vp units from flipper base live catch dampening will occur
Const LiveDistanceMax = 114  'maximum distance in vp units from flipper base live catch dampening will occur (tip protection)

Sub CheckLiveCatch(ball, Flipper, FCount, parm) 'Experimental new live catch
        Dim Dir
        Dir = Flipper.startangle/Abs(Flipper.startangle)    '-1 for Right Flipper
        Dim LiveCatchBounce                                                                                                                        'If live catch is not perfect, it won't freeze ball totally
        Dim CatchTime : CatchTime = GameTime - FCount

        if CatchTime <= LiveCatch and parm > 6 and ABS(Flipper.x - ball.x) > LiveDistanceMin and ABS(Flipper.x - ball.x) < LiveDistanceMax Then
                if CatchTime <= LiveCatch*0.5 Then                                                'Perfect catch only when catch time happens in the beginning of the window
                        LiveCatchBounce = 0
                else
                        LiveCatchBounce = Abs((LiveCatch/2) - CatchTime)        'Partial catch when catch happens a bit late
                end If

                If LiveCatchBounce = 0 and ball.velx * Dir > 0 Then ball.velx = 0
                ball.vely = LiveCatchBounce * (32 / LiveCatch) ' Multiplier for inaccuracy bounce
                ball.angmomx= 0
                ball.angmomy= 0
                ball.angmomz= 0
    Else
        If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 1 Then FlippersD.Dampenf Activeball, parm
        End If
End Sub


'**************************************************
'        Flipper Collision Subs
'NOTE: COpy and overwrite collision sound from original collision subs over
'RandomSoundFlipper()' below
'**************************************************'

Sub LeftFlipper_Collide(parm)
    CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
    RandomSoundFlipper() 'Remove this line if Fleep is integrated
    'LeftFlipperCollide parm   'This is the Fleep code
End Sub

Sub RightFlipper_Collide(parm)
    CheckLiveCatch Activeball, RightFlipper, RFCount, parm
    RandomSoundFlipper() 'Remove this line if Fleep is integrated
    'RightFlipperCollide parm  'This is the Fleep code
End Sub


'****************************************************************************
'PHYSICS DAMPENERS
'****************************************************************************

'These are data mined bounce curves,
'dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
'Requires tracking ballspeed to calculate COR


Sub dPosts_Hit(idx)
        RubbersD.dampen Activeball
End Sub

Sub dSleeves_Hit(idx)
        SleevesD.Dampen Activeball
End Sub


dim RubbersD : Set RubbersD = new Dampener        'frubber
RubbersD.name = "Rubbers"
RubbersD.debugOn = False        'shows info in textbox "TBPout"
RubbersD.Print = False        'debug, reports in debugger (in vel, out cor)
'cor bounce curve (linear)
'for best results, try to match in-game velocity as closely as possible to the desired curve
'RubbersD.addpoint 0, 0, 0.935        'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 0, 0, 1.1        'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 1, 3.77, 0.97
RubbersD.addpoint 2, 5.76, 0.967        'dont take this as gospel. if you can data mine rubber elasticitiy, please help!
RubbersD.addpoint 3, 15.84, 0.874
RubbersD.addpoint 4, 56, 0.64        'there's clamping so interpolate up to 56 at least

dim SleevesD : Set SleevesD = new Dampener        'this is just rubber but cut down to 85%...
SleevesD.name = "Sleeves"
SleevesD.debugOn = False        'shows info in textbox "TBPout"
SleevesD.Print = False        'debug, reports in debugger (in vel, out cor)
SleevesD.CopyCoef RubbersD, 0.85

'######################### Add new FlippersD Profile
'#########################    Adjust these values to increase or lessen the elasticity

dim FlippersD : Set FlippersD = new Dampener
FlippersD.name = "Flippers"
FlippersD.debugOn = False
FlippersD.Print = False
FlippersD.addpoint 0, 0, 1.1
FlippersD.addpoint 1, 3.77, 0.99
FlippersD.addpoint 2, 6, 0.99

Class Dampener
        Public Print, debugOn 'tbpOut.text
        public name, Threshold         'Minimum threshold. Useful for Flippers, which don't have a hit threshold.
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
                RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id)+0.0001)
                coef = desiredcor / realcor
                if debugOn then str = name & " in vel:" & round(cor.ballvel(aBall.id),2 ) & vbnewline & "desired cor: " & round(desiredcor,4) & vbnewline & _
                "actual cor: " & round(realCOR,4) & vbnewline & "ballspeed coef: " & round(coef, 3) & vbnewline
                if Print then debug.print Round(cor.ballvel(aBall.id),2) & ", " & round(desiredcor,3)

                aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
                if debugOn then TBPout.text = str
        End Sub

        public sub Dampenf(aBall, parm) 'Rubberizer is handle here
                dim RealCOR, DesiredCOR, str, coef
                DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
                RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id)+0.0001)
                coef = desiredcor / realcor
                If abs(aball.velx) < 2 and aball.vely < 0 and aball.vely > -3.75 then
                        aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
                End If
        End Sub

        Public Sub CopyCoef(aObj, aCoef) 'alternative addpoints, copy with coef
                dim x : for x = 0 to uBound(aObj.ModIn)
                        addpoint x, aObj.ModIn(x), aObj.ModOut(x)*aCoef
                Next
        End Sub


        Public Sub Report()         'debug, reports all coords in tbPL.text
                if not debugOn then exit sub
                dim a1, a2 : a1 = ModIn : a2 = ModOut
                dim str, x : for x = 0 to uBound(a1) : str = str & x & ": " & round(a1(x),4) & ", " & round(a2(x),4) & vbnewline : next
                TBPout.text = str
        End Sub

End Class


'******************************************************
'                TRACK ALL BALL VELOCITIES
'                 FOR RUBBER DAMPENER AND DROP TARGETS
'******************************************************

dim cor : set cor = New CoRTracker

Class CoRTracker
        public ballvel, ballvelx, ballvely

        Private Sub Class_Initialize : redim ballvel(0) : redim ballvelx(0): redim ballvely(0) : End Sub

        Public Sub Update()        'tracks in-ball-velocity
                dim str, b, AllBalls, highestID : allBalls = getballs

                for each b in allballs
                        if b.id >= HighestID then highestID = b.id
                Next

                if uBound(ballvel) < highestID then redim ballvel(highestID)        'set bounds
                if uBound(ballvelx) < highestID then redim ballvelx(highestID)        'set bounds
                if uBound(ballvely) < highestID then redim ballvely(highestID)        'set bounds

                for each b in allballs
                        ballvel(b.id) = BallSpeed(b)
                        ballvelx(b.id) = b.velx
                        ballvely(b.id) = b.vely
                Next
        End Sub
End Class

Sub RDampen_Timer()
       Cor.Update
End Sub

'******** Copy from this green line to to the end of script *******

' VR PLUNGER ANIMATION
'
' Code needed to animate the plunger. If you pull the plunger it will move in VR.
' IMPORTANT: there are two numeric values in the code that define the postion of the plunger and the
' range in which it can move. The first numeric value is the actual y position of the plunger primitive
' and the second is the actual y position + 100 to determine the range in which it can move.
'
' You need to to select the VR_Primary_plunger primitive you copied from the
' template and copy the value of the Y position
' (e.g. 2130) into the code. The value that determines the range of the plunger is always the y
' position + 100 (e.g. 2230).
'

Sub TimerPlunger_Timer

  If VR_Primary_plunger.Y < 0 then
      VR_Primary_plunger.Y = VR_Primary_plunger.Y + 5
  End If
End Sub

Sub TimerPlunger2_Timer
  VR_Primary_plunger.Y = -100 + (5* Plunger.Position) -20
End Sub

' CODE BELOW IS FOR THE VR CLOCK.  This was originally taken from Rascal VP9 clock table

Dim CurrentMinute

Sub ClockTimer_Timer()

    'ClockHands Below
  VR_Clock_Minutes.RotAndTra2 = (Minute(Now())+(Second(Now())/100))*6
  VR_Clock_Hours.RotAndTra2 = Hour(Now())*30+(Minute(Now())/2)
      VR_Clock_Seconds.RotAndTra2 = (Second(Now()))*6
  CurrentMinute=Minute(Now())

End Sub

'LUT (Colour Look Up Table)

'0 = Fleep Natural Dark 1
'1 = Fleep Natural Dark 2
'2 = Fleep Warm Dark
'3 = Fleep Warm Bright
'4 = Fleep Warm Vivid Soft
'5 = Fleep Warm Vivid Hard
'6 = Skitso Natural and Balanced
'7 = Skitso Natural High Contrast
'8 = 3rdaxis Referenced THX Standard
'9 = CalleV Punchy Brightness and Contrast
'10 = HauntFreaks Desaturated
'11 = Tomate Washed Out
'12 = VPW Original 1 to 1
'13 = Bassgeige
'14 = Blacklight
'15 = B&W Comic Book

Dim LUTset, LutToggleSound
LutToggleSound = True
LoadLUT
SetLUT

'LUT selector timer

dim lutsetsounddir
sub LutSlctr_timer
  If lutsetsounddir = 1 And LutSet <> 15 Then
    Playsound "click", 0, 1, 0.1, 0.25
  End If
  If lutsetsounddir = -1 And LutSet <> 15 Then
    Playsound "click", 0, 1, 0.1, 0.25
  End If
  If LutSet = 15 Then
    Playsound "gun", 0, 1, 0.1, 0.25
  End If
  LutSlctr.enabled = False
end sub

'LUT Subs

Sub SetLUT
  Table1.ColorGradeImage = "LUT" & LUTset
end sub

Sub LUTBox_Timer
  LUTBox.TimerEnabled = 0
  LUTBox.Visible = 0
  LUTBack.visible = 0
  VRLutdesc.visible = 0
End Sub

Sub ShowLUT
  LUTBox.visible = 1
  LUTBack.visible = 1
  VRLutdesc.visible = 1

  Select Case LUTSet
        Case 0: LUTBox.text = "Fleep Natural Dark 1": VRLUTdesc.imageA = "LUTcase0"
    Case 1: LUTBox.text = "Fleep Natural Dark 2": VRLUTdesc.imageA = "LUTcase1"
    Case 2: LUTBox.text = "Fleep Warm Dark": VRLUTdesc.imageA = "LUTcase2"
    Case 3: LUTBox.text = "Fleep Warm Bright": VRLUTdesc.imageA = "LUTcase3"
    Case 4: LUTBox.text = "Fleep Warm Vivid Soft": VRLUTdesc.imageA = "LUTcase4"
    Case 5: LUTBox.text = "Fleep Warm Vivid Hard": VRLUTdesc.imageA = "LUTcase5"
    Case 6: LUTBox.text = "Skitso Natural and Balanced": VRLUTdesc.imageA = "LUTcase6"
    Case 7: LUTBox.text = "Skitso Natural High Contrast": VRLUTdesc.imageA = "LUTcase7"
    Case 8: LUTBox.text = "3rdaxis Referenced THX Standard": VRLUTdesc.imageA = "LUTcase8"
    Case 9: LUTBox.text = "CalleV Punchy Brightness and Contrast": VRLUTdesc.imageA = "LUTcase9"
    Case 10: LUTBox.text = "HauntFreaks Desaturated" : VRLUTdesc.imageA = "LUTcase10"
      Case 11: LUTBox.text = "Tomate washed out": VRLUTdesc.imageA = "LUTcase11"
    Case 12: LUTBox.text = "VPW original 1on1": VRLUTdesc.imageA = "LUTcase12"
    Case 13: LUTBox.text = "bassgeige": VRLUTdesc.imageA = "LUTcase13"
    Case 14: LUTBox.text = "blacklight": VRLUTdesc.imageA = "LUTcase14"
    Case 15: LUTBox.text = "B&W Comic Book": VRLUTdesc.imageA = "LUTcase15"
  End Select

  LUTBox.TimerEnabled = 1

End Sub

Sub SaveLUT

  Dim FileObj
  Dim ScoreFile

  Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if

  if LUTset = "" then LUTset = 13 'failsafe

  Set ScoreFile=FileObj.CreateTextFile(UserDirectory & "TableLUT.txt",True) 'Rename the tableLUT
  ScoreFile.WriteLine LUTset
  Set ScoreFile=Nothing
  Set FileObj=Nothing

End Sub

Sub LoadLUT

  Dim FileObj, ScoreFile, TextStr
  dim rLine

  Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    LUTset=13
    Exit Sub
  End if
  If Not FileObj.FileExists(UserDirectory & "TableLUT.txt") then  'Rename the tableLUT
    LUTset=13
    Exit Sub
  End if
  Set ScoreFile=FileObj.GetFile(UserDirectory & "TableLUT.txt")  'Rename the tableLUT
  Set TextStr=ScoreFile.OpenAsTextStream(1,0)
    If (TextStr.AtEndOfStream=True) then
      Exit Sub
    End if
    rLine = TextStr.ReadLine
    If rLine = "" then
      LUTset=13
      Exit Sub
    End if
    LUTset = int (rLine)
    Set ScoreFile = Nothing
      Set FileObj = Nothing
End Sub

'**************************************************************


' ***************************************************************************
'          VR Backglass
' ****************************************************************************


' ***************************************************************************
'          (EM) 1-4 player 5x drums, 1 credit drum CORE CODE
' ****************************************************************************


' ********************* POSITION EM REEL DRUMS ON BACKGLASS *************************

Dim xoff,yoff,zoff,xrot,zscale, xcen,ycen
Dim inx
xoff =472
yoff = 0
zoff =835
xrot = -90

Const USEEMS = 1 ' 1-4 set between 1 to 4 based on number of players

const idx_emp1r1 =0 'player 1
const idx_emp2r1 =5 'player 2
const idx_emp3r1 =10 'player 3
const idx_emp4r1 =15 'player 4
const idx_emp4r6 =20 'credits


Dim BGObjEM(1)
if USEEMS = 1 then
  BGObjEM(0) = Array(emp1r1, emp1r2, emp1r3, emp1r4, emp1r5, _
  Empty,Empty,Empty,Empty,Empty,_
  Empty,Empty,Empty,Empty,Empty,_
  Empty,Empty,Empty,Empty,Empty,_
  emp4r6) ' credits
elseif USEEMS = 2 then
  BGObjEM(0) = Array(emp1r1, emp1r2, emp1r3, emp1r4, emp1r5, _
  emp2r1, emp2r2, emp2r3, emp2r4, emp2r5, _
  Empty,Empty,Empty,Empty,Empty,_
  Empty,Empty,Empty,Empty,Empty,_
  emp4r6) ' credits
elseif USEEMS = 3 then
  BGObjEM(0) = Array(emp1r1, emp1r2, emp1r3, emp1r4, emp1r5, _
  emp2r1, emp2r2, emp2r3, emp2r4, emp2r5, _
  emp3r1, emp3r2, emp3r3, emp3r4, emp3r5, _
  Empty,Empty,Empty,Empty,Empty,_
  emp4r6) ' credits
elseif USEEMS = 4 then
  BGObjEM(0) = Array(emp1r1, emp1r2, emp1r3, emp1r4, emp1r5, _
  emp2r1, emp2r2, emp2r3, emp2r4, emp2r5, _
  emp3r1, emp3r2, emp3r3, emp3r4, emp3r5, _
  emp4r1, emp4r2, emp4r3, emp4r4, emp4r5, _
  emp4r6) ' credits
end If

Sub center_objects_em()
Dim cnt,ii, xx, yy, yfact, xfact, objs
'exit sub
yoff = -150
zscale = 0.0000001
xcen =(960 /2) - (17 / 2)
ycen = (1065 /2 ) + (313 /2)

yfact = -25
xfact = 0

cnt =0
  For Each objs In BGObjEM(0)
  If Not IsEmpty(objs) then
    if objs.name = emp4r6.name then
    yoff = 45 ' credit drum is 60% smaller
    Else
    yoff = 2
    end if

  xx =objs.x

  objs.x = (xoff - xcen) + xx + xfact
  yy = objs.y
  objs.y =yoff

    If yy < 0 then
    yy = yy * -1
    end if

  objs.z = (zoff - ycen) + yy - (yy * zscale) + yfact

  'objs.rotx = xrot
  end if
  cnt = cnt + 1
  Next

end sub



' ********************* UPDATE EM REEL DRUMS CORE LIB *************************

Dim cred,ix, np,npp, reels(5, 7), scores(6,2)

'reset scores to defaults
for np =0 to 5
scores(np,0 ) = 0
scores(np,1 ) = 0
Next

'reset EM drums to defaults
For np =0 to 3
  For  npp =0 to 6
  reels(np, npp) =0 ' default to zero
  Next
Next


Sub SetScore(player, ndx , val)

Dim ncnt

  if player = 5 or player = 6 then
    if val > 0 then
      If(ndx = 0)Then ncnt = val * 10
      If(ndx = 1)Then ncnt = val

      scores(player, 0) = scores(player, 0) + ncnt
    end if
  else
    if val > 0 then

    If(ndx = 0)then ncnt = val * 10000
    If(ndx = 1)then ncnt = val * 1000
    If(ndx = 2)Then ncnt = val * 100
    If(ndx = 3)Then ncnt = val * 10
    If(ndx = 4)Then ncnt = val

    scores(player, 0) = scores(player, 0) + ncnt
    'scores(player, 0) + ncnt

    end if
  end if
End Sub


Sub SetDrum(player, drum , val)
Dim cnt
Dim objs : objs =BGObjEM(0)

  If val = 0 then
    Select case player
    case -1: ' the credit drum
    If Not IsEmpty(objs(idx_emp4r6)) then
    objs(idx_emp4r6).ObjrotX = 0 ' 285
    'cnt =objs(idx_emp4r6).ObjrotX
    end if
    Case 0:
    Select Case drum
        Case 1: If Not IsEmpty(objs(idx_emp1r1)) then: objs(idx_emp1r1).ObjrotX = 0: end if' 283
        Case 2: If Not IsEmpty(objs(idx_emp1r1+1)) then: objs(idx_emp1r1+1).ObjrotX=0: end if
        Case 3: If Not IsEmpty(objs(idx_emp1r1+2)) then: objs(idx_emp1r1+2).ObjrotX=0: end if
        Case 4: If Not IsEmpty(objs(idx_emp1r1+3)) then: objs(idx_emp1r1+3).ObjrotX=0: end if
        Case 5: If Not IsEmpty(objs(idx_emp1r1+4)) then: objs(idx_emp1r1+4).ObjrotX=0: end if
    End Select
    Case 1:
    Select Case drum
        Case 1: If Not IsEmpty(objs(idx_emp2r1)) then: objs(idx_emp2r1).ObjrotX = 0: end if' 283
        Case 2: If Not IsEmpty(objs(idx_emp2r1+1)) then: objs(idx_emp2r1+1).ObjrotX=0: end if
        Case 3: If Not IsEmpty(objs(idx_emp2r1+2)) then: objs(idx_emp2r1+2).ObjrotX=0: end if
        Case 4: If Not IsEmpty(objs(idx_emp2r1+3)) then: objs(idx_emp2r1+3).ObjrotX=0: end if
        Case 5: If Not IsEmpty(objs(idx_emp2r1+4)) then: objs(idx_emp2r1+4).ObjrotX=0: end if
    End Select
    Case 2:
    Select Case drum
        Case 1: If Not IsEmpty(objs(idx_emp3r1)) then: objs(idx_emp3r1).ObjrotX = 0: end if' 283
        Case 2: If Not IsEmpty(objs(idx_emp3r1+1)) then: objs(idx_emp3r1+1).ObjrotX=0: end if
        Case 3: If Not IsEmpty(objs(idx_emp3r1+2)) then: objs(idx_emp3r1+2).ObjrotX=0: end if
        Case 4: If Not IsEmpty(objs(idx_emp3r1+3)) then: objs(idx_emp3r1+3).ObjrotX=0: end if
        Case 5: If Not IsEmpty(objs(idx_emp3r1+4)) then: objs(idx_emp3r1+4).ObjrotX=0: end if
    End Select
    Case 3:
    Select Case drum
        Case 1: If Not IsEmpty(objs(idx_emp4r1)) then: objs(idx_emp4r1).ObjrotX = 0: end if' 283
        Case 2: If Not IsEmpty(objs(idx_emp4r1+1)) then: objs(idx_emp4r1+1).ObjrotX=0: end if
        Case 3: If Not IsEmpty(objs(idx_emp4r1+2)) then: objs(idx_emp4r1+2).ObjrotX=0: end if
        Case 4: If Not IsEmpty(objs(idx_emp4r1+3)) then: objs(idx_emp4r1+3).ObjrotX=0: end if
        Case 5: If Not IsEmpty(objs(idx_emp4r1+4)) then: objs(idx_emp4r1+4).ObjrotX=0: end if
    End Select
  End Select

  else
  Select case player

    Case -1: ' the credit drum
    'emp4r6.ObjrotX = emp4r6.ObjrotX + val
    If Not IsEmpty(objs(idx_emp4r6)) then
    objs(idx_emp4r6).ObjrotX = objs(idx_emp4r6).ObjrotX + val
    end if

    Case 0:
    Select Case drum
        Case 1: If Not IsEmpty(objs(idx_emp1r1)) then: objs(idx_emp1r1).ObjrotX= objs(idx_emp1r1).ObjrotX + val: end if
        Case 2: If Not IsEmpty(objs(idx_emp1r1+1)) then: objs(idx_emp1r1+1).ObjrotX= objs(idx_emp1r1+1).ObjrotX + val: end if
        Case 3: If Not IsEmpty(objs(idx_emp1r1+2)) then: objs(idx_emp1r1+2).ObjrotX= objs(idx_emp1r1+2).ObjrotX + val: end if
        Case 4: If Not IsEmpty(objs(idx_emp1r1+3)) then: objs(idx_emp1r1+3).ObjrotX= objs(idx_emp1r1+3).ObjrotX + val: end if
        Case 5: If Not IsEmpty(objs(idx_emp1r1+4)) then: objs(idx_emp1r1+4).ObjrotX= objs(idx_emp1r1+4).ObjrotX + val: end if
    End Select
    Case 1:
    Select Case drum
        Case 1: If Not IsEmpty(objs(idx_emp2r1)) then: objs(idx_emp2r1).ObjrotX= objs(idx_emp2r1).ObjrotX + val: end if
        Case 2: If Not IsEmpty(objs(idx_emp2r1+1)) then: objs(idx_emp2r1+1).ObjrotX= objs(idx_emp2r1+1).ObjrotX + val: end if
        Case 3: If Not IsEmpty(objs(idx_emp2r1+2)) then: objs(idx_emp2r1+2).ObjrotX= objs(idx_emp2r1+2).ObjrotX + val: end if
        Case 4: If Not IsEmpty(objs(idx_emp2r1+3)) then: objs(idx_emp2r1+3).ObjrotX= objs(idx_emp2r1+3).ObjrotX + val: end if
        Case 5: If Not IsEmpty(objs(idx_emp2r1+4)) then: objs(idx_emp2r1+4).ObjrotX= objs(idx_emp2r1+4).ObjrotX + val: end if
    End Select
    Case 2:
    Select Case drum
        Case 1: If Not IsEmpty(objs(idx_emp3r1)) then: objs(idx_emp3r1).ObjrotX= objs(idx_emp3r1).ObjrotX + val: end if
        Case 2: If Not IsEmpty(objs(idx_emp3r1+1)) then: objs(idx_emp3r1+1).ObjrotX= objs(idx_emp3r1+1).ObjrotX + val: end if
        Case 3: If Not IsEmpty(objs(idx_emp3r1+2)) then: objs(idx_emp3r1+2).ObjrotX= objs(idx_emp3r1+2).ObjrotX + val: end if
        Case 4: If Not IsEmpty(objs(idx_emp3r1+3)) then: objs(idx_emp3r1+3).ObjrotX= objs(idx_emp3r1+3).ObjrotX + val: end if
        Case 5: If Not IsEmpty(objs(idx_emp3r1+4)) then: objs(idx_emp3r1+4).ObjrotX= objs(idx_emp3r1+4).ObjrotX + val: end if
    End Select
    Case 3:
    Select Case drum
        Case 1: If Not IsEmpty(objs(idx_emp4r1)) then: objs(idx_emp4r1).ObjrotX= objs(idx_emp4r1).ObjrotX + val: end if
        Case 2: If Not IsEmpty(objs(idx_emp4r1+1)) then: objs(idx_emp4r1+1).ObjrotX= objs(idx_emp4r1+1).ObjrotX + val: end if
        Case 3: If Not IsEmpty(objs(idx_emp4r1+2)) then: objs(idx_emp4r1+2).ObjrotX= objs(idx_emp4r1+2).ObjrotX + val: end if
        Case 4: If Not IsEmpty(objs(idx_emp4r1+3)) then: objs(idx_emp4r1+3).ObjrotX= objs(idx_emp4r1+3).ObjrotX + val: end if
        Case 5: If Not IsEmpty(objs(idx_emp4r1+4)) then: objs(idx_emp4r1+4).ObjrotX= objs(idx_emp4r1+4).ObjrotX + val: end if
    End Select

  End Select
  end if
End Sub


Sub SetReel(player, drum, val)

'TextBox1.text = "playr:" & player +1 & " drum:" & drum & "val:" & val

Dim  inc , cur, dif, fix, fval

inc = 33.5
fval = -5 ' graphic seam between 5 & 6 fix value, easier to fix here than photoshop

If  (player <= 3) or (drum = -1) then

  If drum = -1 then drum = 0

  cur =reels(player, drum)

  If val <> cur then ' something has changed
  Select Case drum

    Case 0: ' credits drum

      if val > cur then
        dif =val - cur
        fix =0
          If cur < 5 and cur+dif > 5 then
          fix = fix- fval
          end if
        dif = dif * inc

        dif = dif-fix

        SetDrum -1,0,  -dif
      Else
        if val = 0 Then
        SetDrum -1,0,  0' reset the drum to abs. zero
        Else
        dif = 11 - cur
        dif = dif + val

        dif = dif * inc
        dif = dif-fval

        SetDrum -1,0,   -dif
        end if
      end if
    Case 1:
    'TB1.text = val
    if val > cur then
      dif =val - cur
      fix =0
        If cur < 5 and cur+dif > 5 then
        fix = fix- fval
        end if
      dif = dif * inc

      dif = dif-fix

      SetDrum player,drum,  -dif
    Else
      if val = 0 Then
      SetDrum player,drum,  0' reset the drum to abs. zero
      Else
      dif = 11 - cur
      dif = dif + val

      dif = dif * inc
      dif = dif-fval

      SetDrum player,drum,   -dif
      end if

    end if
    reels(player, drum) = val

    Case 2:
    'TB2.text = val

    if val > cur then
      dif =val - cur
      fix =0
        If cur < 5 and cur+dif > 5 then
        fix = fix- fval
        end if
      dif = dif * inc
      dif = dif-fix
      SetDrum player,drum,  -dif
    Else
      if val = 0 Then
      SetDrum player,drum,  0 ' reset the drum to abs. zero
      Else
      dif = 11 - cur
      dif = dif + val
      dif = dif * inc
      dif = dif-fval
      SetDrum player,drum,  -dif
      end if
    end if
    reels(player, drum) = val

    Case 3:
    'TB3.text = val

    if val > cur then
      dif =val - cur
      fix =0
        If cur < 5 and cur+dif > 5 then
        fix = fix- fval
        end if
      dif = dif * inc
      dif = dif-fix

      SetDrum player,drum,  -dif
    Else
      if val = 0 Then
      SetDrum player,drum,  0 ' reset the drum to abs. zero
      Else
      dif = 11 - cur
      dif = dif + val
      dif = dif * inc
      dif = dif-fval
      SetDrum player,drum,  -dif
      end if

    end if
    reels(player, drum) = val

    Case 4:
    'TB4.text = val

    if val > cur then
      dif =val - cur
      fix =0
        If cur < 5 and cur+dif > 5 then
        fix = fix- fval
        end if
      dif = dif * inc
      dif = dif-fix
      SetDrum player,drum,  -dif
    Else
      if val = 0 Then
      SetDrum player,drum,  0 ' reset the drum to abs. zero
      Else
      dif = 11 - cur
      dif = dif + val
      dif = dif * inc
      dif = dif-fval
      SetDrum player,drum,  -dif
      end if

    end if
    reels(player, drum) = val

    Case 5:
    'TB5.text = val

    if val > cur then
      dif =val - cur
      fix =0
        If cur < 5 and cur+dif > 5 then
        fix = fix- fval
        end if
      dif = dif * inc
      dif = dif-fix
      SetDrum player,drum,  -dif
    Else
      if val = 0 Then
      SetDrum player,drum,  0 ' reset the drum to abs. zero
      Else
      dif = 11 - cur
      dif = dif + val
      dif = dif * inc
      dif = dif-fval
      SetDrum player,drum,  -dif
      end if

    end if
    reels(player, drum) = val
   End Select

  end if
end if
End Sub

Dim EMMODE: EMMODE = 0
'Dim Score10,Score100,Score1000,Score10000,Score100000
Dim Score1000,Score10000,Score100000, ActivePLayer
Dim nplayer,playr,value,curscr,curplayr

' EMMODE = 1
'ex: UpdateVRReels 0-3, 0-3, 0-199999, n/a, n/a, n/a, n/a, n/a, n/a
' EMMODE = 0
'ex: UpdateVRReels 0-3, 0-3, n/a, 0-1,0-99999 ,0-9999, 0-999, 0-99, 0-9

Sub UpdateVRReels (Player,nReels ,nScore, n100K, Score10000 ,Score1000,Score100,Score10,Score1)

' to-do find out if player is one or zero based, if 1 based subtract 1.
value =nScore'Score(Player)
  nplayer = Player -1

  curscr = value
  curplayr = nplayer


scores(0,1) = scores(0,0)
  scores(0,0) = 0
  scores(1,1) = scores(1,0)
  scores(1,0) = 0
  scores(2,1) = scores(2,0)
  scores(2,0) = 0
  scores(3,1) = scores(3,0)
  scores(3,0) = 0

  For  ix =0 to 6
    reels(0, ix) =0
    reels(1, ix) =0
    reels(2, ix) =0
    reels(3, ix) =0
  Next

  For  ix =0 to 4

  SetDrum ix, 1 , 0
  SetDrum ix, 2 , 0
  SetDrum ix, 3 , 0
  SetDrum ix, 4 , 0
  SetDrum ix, 5 , 0

  Next

  For playr =0 to nReels

    if EMMODE = 0 then
    If (ActivePLayer) = playr Then
    nplayer = playr

    SetReel nplayer, 1 , Score10000 : SetScore nplayer,0,Score10000
    SetReel nplayer, 2 , Score1000 : SetScore nplayer,1,Score1000
    SetReel nplayer, 3 , Score100 : SetScore nplayer,2,Score100
    SetReel nplayer, 4 , Score10 : SetScore nplayer,3,Score10
    SetReel nplayer, 5 , 0 : SetScore nplayer,4,0 ' assumes ones position is always zero

    else
    nplayer = playr
    value =scores(nplayer, 1)


  ' do ten thousands
    if(value >= 90000)  then:  SetReel nplayer, 1 , 9 : SetScore nplayer,0,9 : value = value - 90000: end if
    if(value >= 80000)  then:  SetReel nplayer, 1 , 8 : SetScore nplayer,0,8 : value = value - 80000: end if
    if(value >= 70000)  then:  SetReel nplayer, 1 , 7 : SetScore nplayer,0,7 : value = value - 70000: end if
    if(value >= 60000)  then:  SetReel nplayer, 1 , 6 : SetScore nplayer,0,6 : value = value - 60000: end if
    if(value >= 50000)  then:  SetReel nplayer, 1 , 5 : SetScore nplayer,0,5 : value = value - 50000: end if
    if(value >= 40000)  then:  SetReel nplayer, 1 , 4 : SetScore nplayer,0,4 : value = value - 40000: end if
    if(value >= 30000)  then:  SetReel nplayer, 1 , 3 : SetScore nplayer,0,3 : value = value - 30000: end if
    if(value >= 20000)  then:  SetReel nplayer, 1 , 2 : SetScore nplayer,0,2 : value = value - 20000: end if
    if(value >= 10000)  then:  SetReel nplayer, 1 , 1 : SetScore nplayer,0,1 : value = value - 10000: end if


  ' do thousands
    if(value >= 9000)  then:  SetReel nplayer, 2 , 9 : SetScore nplayer,1,9 : value = value - 9000: end if
    if(value >= 8000)  then:  SetReel nplayer, 2 , 8 : SetScore nplayer,1,8 : value = value - 8000: end if
    if(value >= 7000)  then:  SetReel nplayer, 2 , 7 : SetScore nplayer,1,7 : value = value - 7000: end if
    if(value >= 6000)  then:  SetReel nplayer, 2 , 6 : SetScore nplayer,1,6 : value = value - 6000: end if
    if(value >= 5000)  then:  SetReel nplayer, 2 , 5 : SetScore nplayer,1,5 : value = value - 5000: end if
    if(value >= 4000)  then:  SetReel nplayer, 2 , 4 : SetScore nplayer,1,4 : value = value - 4000: end if
    if(value >= 3000)  then:  SetReel nplayer, 2 , 3 : SetScore nplayer,1,3 : value = value - 3000: end if
    if(value >= 2000)  then:  SetReel nplayer, 2 , 2 : SetScore nplayer,1,2 : value = value - 2000: end if
    if(value >= 1000)  then:  SetReel nplayer, 2 , 1 : SetScore nplayer,1,1 : value = value - 1000: end if

    'do hundreds

    if(value >= 900) then: SetReel nplayer, 3 , 9 : SetScore nplayer,2,9 : value = value - 900: end if
    if(value >= 800) then: SetReel nplayer, 3 , 8 : SetScore nplayer,2,8 : value = value - 800: end if
    if(value >= 700) then: SetReel nplayer, 3 , 7 : SetScore nplayer,2,7 : value = value - 700: end if
    if(value >= 600) then: SetReel nplayer, 3 , 6 : SetScore nplayer,2,6 : value = value - 600: end if
    if(value >= 500) then: SetReel nplayer, 3 , 5 : SetScore nplayer,2,5 : value = value - 500: end if
    if(value >= 400) then: SetReel nplayer, 3 , 4 : SetScore nplayer,2,4 : value = value - 400: end if
    if(value >= 300) then: SetReel nplayer, 3 , 3 : SetScore nplayer,2,3 : value = value - 300: end if
    if(value >= 200) then: SetReel nplayer, 3 , 2 : SetScore nplayer,2,2 : value = value - 200: end if
    if(value >= 100) then: SetReel nplayer, 3 , 1 : SetScore nplayer,2,1 : value = value - 100: end if

    'do tens
    if(value >= 90) then: SetReel nplayer, 4 , 9 : SetScore nplayer,3,9 : value = value - 90: end if
    if(value >= 80) then: SetReel nplayer, 4 , 8 : SetScore nplayer,3,8 : value = value - 80: end if
    if(value >= 70) then: SetReel nplayer, 4 , 7 : SetScore nplayer,3,7 : value = value - 70: end if
    if(value >= 60) then: SetReel nplayer, 4 , 6 : SetScore nplayer,3,6 : value = value - 60: end if
    if(value >= 50) then: SetReel nplayer, 4 , 5 : SetScore nplayer,3,5 : value = value - 50: end if
    if(value >= 40) then: SetReel nplayer, 4 , 4 : SetScore nplayer,3,4 : value = value - 40: end if
    if(value >= 30) then: SetReel nplayer, 4 , 3 : SetScore nplayer,3,3 : value = value - 30: end if
    if(value >= 20) then: SetReel nplayer, 4 , 2 : SetScore nplayer,3,2 : value = value - 20: end if
    if(value >= 10) then: SetReel nplayer, 4 , 1 : SetScore nplayer,3,1 : value = value - 10: end if

    'do ones
    if(value >= 9) then: SetReel nplayer, 5 , 9 : SetScore nplayer,4,9 : value = value - 9: end if
    if(value >= 8) then: SetReel nplayer, 5 , 8 : SetScore nplayer,4,8 : value = value - 8: end if
    if(value >= 7) then: SetReel nplayer, 5 , 7 : SetScore nplayer,4,7 : value = value - 7: end if
    if(value >= 6) then: SetReel nplayer, 5 , 6 : SetScore nplayer,4,6 : value = value - 6: end if
    if(value >= 5) then: SetReel nplayer, 5 , 5 : SetScore nplayer,4,5 : value = value - 5: end if
    if(value >= 4) then: SetReel nplayer, 5 , 4 : SetScore nplayer,4,4 : value = value - 4: end if
    if(value >= 3) then: SetReel nplayer, 5 , 3 : SetScore nplayer,4,3 : value = value - 3: end if
    if(value >= 2) then: SetReel nplayer, 5 , 2 : SetScore nplayer,4,2 : value = value - 2: end if
    if(value >= 1) then: SetReel nplayer, 5 , 1 : SetScore nplayer,4,1 : value = value - 1: end if

    end if
    Else
      If curplayr = playr Then
      nplayer = curplayr
      value = curscr
      else
      value =scores(playr, 1) ' store score
      nplayer = playr
      end if

    scores(playr, 0)  = 0 ' reset score

    if(value >= 100000) then

      'if nplayer = 0 then: FL100K1.visible = 1
      'if nplayer = 1 then: FL100K2.visible = 1
      'if nplayer = 2 then: FL100K3.visible = 1
      'if nplayer = 3 then: FL100K4.visible = 1

    value = value - 100000

    end if


  ' do ten thousands
    if(value >= 90000)  then:  SetReel nplayer, 1 , 9 : SetScore nplayer,0,9 : value = value - 90000: end if
    if(value >= 80000)  then:  SetReel nplayer, 1 , 8 : SetScore nplayer,0,8 : value = value - 80000: end if
    if(value >= 70000)  then:  SetReel nplayer, 1 , 7 : SetScore nplayer,0,7 : value = value - 70000: end if
    if(value >= 60000)  then:  SetReel nplayer, 1 , 6 : SetScore nplayer,0,6 : value = value - 60000: end if
    if(value >= 50000)  then:  SetReel nplayer, 1 , 5 : SetScore nplayer,0,5 : value = value - 50000: end if
    if(value >= 40000)  then:  SetReel nplayer, 1 , 4 : SetScore nplayer,0,4 : value = value - 40000: end if
    if(value >= 30000)  then:  SetReel nplayer, 1 , 3 : SetScore nplayer,0,3 : value = value - 30000: end if
    if(value >= 20000)  then:  SetReel nplayer, 1 , 2 : SetScore nplayer,0,2 : value = value - 20000: end if
    if(value >= 10000)  then:  SetReel nplayer, 1 , 1 : SetScore nplayer,0,1 : value = value - 10000: end if


  ' do thousands
    if(value >= 9000)  then:  SetReel nplayer, 2 , 9 : SetScore nplayer,1,9 : value = value - 9000: end if
    if(value >= 8000)  then:  SetReel nplayer, 2 , 8 : SetScore nplayer,1,8 : value = value - 8000: end if
    if(value >= 7000)  then:  SetReel nplayer, 2 , 7 : SetScore nplayer,1,7 : value = value - 7000: end if
    if(value >= 6000)  then:  SetReel nplayer, 2 , 6 : SetScore nplayer,1,6 : value = value - 6000: end if
    if(value >= 5000)  then:  SetReel nplayer, 2 , 5 : SetScore nplayer,1,5 : value = value - 5000: end if
    if(value >= 4000)  then:  SetReel nplayer, 2 , 4 : SetScore nplayer,1,4 : value = value - 4000: end if
    if(value >= 3000)  then:  SetReel nplayer, 2 , 3 : SetScore nplayer,1,3 : value = value - 3000: end if
    if(value >= 2000)  then:  SetReel nplayer, 2 , 2 : SetScore nplayer,1,2 : value = value - 2000: end if
    if(value >= 1000)  then:  SetReel nplayer, 2 , 1 : SetScore nplayer,1,1 : value = value - 1000: end if

    'do hundreds

    if(value >= 900) then: SetReel nplayer, 3 , 9 : SetScore nplayer,2,9 : value = value - 900: end if
    if(value >= 800) then: SetReel nplayer, 3 , 8 : SetScore nplayer,2,8 : value = value - 800: end if
    if(value >= 700) then: SetReel nplayer, 3 , 7 : SetScore nplayer,2,7 : value = value - 700: end if
    if(value >= 600) then: SetReel nplayer, 3 , 6 : SetScore nplayer,2,6 : value = value - 600: end if
    if(value >= 500) then: SetReel nplayer, 3 , 5 : SetScore nplayer,2,5 : value = value - 500: end if
    if(value >= 400) then: SetReel nplayer, 3 , 4 : SetScore nplayer,2,4 : value = value - 400: end if
    if(value >= 300) then: SetReel nplayer, 3 , 3 : SetScore nplayer,2,3 : value = value - 300: end if
    if(value >= 200) then: SetReel nplayer, 3 , 2 : SetScore nplayer,2,2 : value = value - 200: end if
    if(value >= 100) then: SetReel nplayer, 3 , 1 : SetScore nplayer,2,1 : value = value - 100: end if

    'do tens
    if(value >= 90) then: SetReel nplayer, 4 , 9 : SetScore nplayer,3,9 : value = value - 90: end if
    if(value >= 80) then: SetReel nplayer, 4 , 8 : SetScore nplayer,3,8 : value = value - 80: end if
    if(value >= 70) then: SetReel nplayer, 4 , 7 : SetScore nplayer,3,7 : value = value - 70: end if
    if(value >= 60) then: SetReel nplayer, 4 , 6 : SetScore nplayer,3,6 : value = value - 60: end if
    if(value >= 50) then: SetReel nplayer, 4 , 5 : SetScore nplayer,3,5 : value = value - 50: end if
    if(value >= 40) then: SetReel nplayer, 4 , 4 : SetScore nplayer,3,4 : value = value - 40: end if
    if(value >= 30) then: SetReel nplayer, 4 , 3 : SetScore nplayer,3,3 : value = value - 30: end if
    if(value >= 20) then: SetReel nplayer, 4 , 2 : SetScore nplayer,3,2 : value = value - 20: end if
    if(value >= 10) then: SetReel nplayer, 4 , 1 : SetScore nplayer,3,1 : value = value - 10: end if

    'do ones
    if(value >= 9) then: SetReel nplayer, 5 , 9 : SetScore nplayer,4,9 : value = value - 9: end if
    if(value >= 8) then: SetReel nplayer, 5 , 8 : SetScore nplayer,4,8 : value = value - 8: end if
    if(value >= 7) then: SetReel nplayer, 5 , 7 : SetScore nplayer,4,7 : value = value - 7: end if
    if(value >= 6) then: SetReel nplayer, 5 , 6 : SetScore nplayer,4,6 : value = value - 6: end if
    if(value >= 5) then: SetReel nplayer, 5 , 5 : SetScore nplayer,4,5 : value = value - 5: end if
    if(value >= 4) then: SetReel nplayer, 5 , 4 : SetScore nplayer,4,4 : value = value - 4: end if
    if(value >= 3) then: SetReel nplayer, 5 , 3 : SetScore nplayer,4,3 : value = value - 3: end if
    if(value >= 2) then: SetReel nplayer, 5 , 2 : SetScore nplayer,4,2 : value = value - 2: end if
    if(value >= 1) then: SetReel nplayer, 5 , 1 : SetScore nplayer,4,1 : value = value - 1: end if


    end if

  Next
End Sub


'******************************************************
'*******     VR Backglass Lighting    *******
'******************************************************

Dim BGObj

If Renderingmode = 2 Then
  SetBackglass
End If

Sub SetBackglass()

  For Each BGObj In VRBackglass
    BGObj.x = BGobj.x
    BGObj.height = - BGObj.y + 170
    BGObj.y =80 'adjusts the distance from the backglass towards the user
  Next

End Sub


Sub FlasherMatch
  If Match = 10 Then FlM00.visible = 1 : FlM00A.visible = 1 : Else : FlM00.visible = 0 : FlM00A.visible = 0 : End If
  If Match = 1 Then FlM10.visible = 1 : FlM10A.visible = 1 : Else : FlM10.visible = 0 : FlM10A.visible = 0 : End If
  If Match = 2 Then FlM20.visible = 1 : FlM20A.visible = 1 : Else : FlM20.visible = 0 : FlM20A.visible = 0 : End If
  If Match = 3 Then FlM30.visible = 1 : FlM30A.visible = 1 : Else : FlM30.visible = 0 : FlM30A.visible = 0 : End If
  If Match = 4 Then FlM40.visible = 1 : FlM40A.visible = 1 : Else : FlM40.visible = 0 : FlM40A.visible = 0 : End If
  If Match = 5 Then FlM50.visible = 1 : FlM50A.visible = 1 : Else : FlM50.visible = 0 : FlM50A.visible = 0 : End If
  If Match = 6 Then FlM60.visible = 1 : FlM60A.visible = 1 : Else : FlM60.visible = 0 : FlM60A.visible = 0 : End If
  If Match = 7 Then FlM70.visible = 1 : FlM70A.visible = 1 : Else : FlM70.visible = 0 : FlM70A.visible = 0 : End If
  If Match = 8 Then FlM80.visible = 1 : FlM80A.visible = 1 : Else : FlM80.visible = 0 : FlM80A.visible = 0 : End If
  If Match = 9 Then FlM90.visible = 1 : FlM90A.visible = 1 : Else : FlM90.visible = 0 : FlM90A.visible = 0 : End If
End Sub

'Sub FlasherBalls
' If BallInPlay = 1 Then FlBIP1.visible = 1 : FlBIP1A.visible = 1 Else FlBIP1.visible = 0 : FlBIP1A.visible = 0 End If
' If BallInPlay = 2 Then FlBIP2.visible = 1 : FlBIP2A.visible = 1 Else FlBIP2.visible = 0 : FlBIP2A.visible = 0 End If
' If BallInPlay = 3 Then FlBIP3.visible = 1 : FlBIP3A.visible = 1 Else FlBIP3.visible = 0 : FlBIP3A.visible = 0 End If
' If BallInPlay = 4 Then FlBIP4.visible = 1 : FlBIP4A.visible = 1 Else FlBIP4.visible = 0 : FlBIP4A.visible = 0 End If
' If BallInPlay = 5 Then FlBIP5.visible = 1 : FlBIP5A.visible = 1 Else FlBIP5.visible = 0 : FlBIP5A.visible = 0 End If
'End Sub

'Sub FlasherPlayers
' If Players = 1 Then FlPl1.visible = 1 Else FlPl1.visible = 0 End If
' If Players = 2 Then FlPl2.visible = 1 Else FlPl2.visible = 0 End If
' If Players = 3 Then FlPl3.visible = 1 Else FlPl3.visible = 0 End If
' If Players = 4 Then FlPl4.visible = 1 Else FlPl4.visible = 0 End If
'End Sub

'Sub FlasherCurrentPlayer
' If Player = 1 Then : for each Object in VRBGPlayer1 : object.visible = 1 : Next: Else : for each Object in VRBGPlayer1 : object.visible = 0 :Next
' If Player = 2 Then : for each Object in VRBGPlayer2 : object.visible = 1 : Next: Else : for each Object in VRBGPlayer2 : object.visible = 0 :Next
' If Player = 3 Then : for each Object in VRBGPlayer3 : object.visible = 1 : Next: Else : for each Object in VRBGPlayer3 : object.visible = 0 :Next
' If Player = 4 Then : for each Object in VRBGPlayer4 : object.visible = 1 : Next: Else : for each Object in VRBGPlayer4 : object.visible = 0 :Next
'End Sub

'Sub Flasher100k
' VRBG1k.visible = 1
'End Sub




'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X  X
'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/
'  SCORBIT Interface
' To Use:
' 1) Define a timer tmrScorbit
' 2) Call DoInit at the end of PupInit or in Table Init if you are nto using pup with the appropriate parameters
'   if Scorbit.DoInit(389, "PupOverlays", "1.0.0", "GRWvz-MP37P") then
'     tmrScorbit.Interval=2000
'     tmrScorbit.UserValue = 0
'     tmrScorbit.Enabled=True
'   End if
' 3) Customize helper functions below for different events if you want or make your own
' 4) Call
'   StartSession - When a game starts
'   StopSession - When the game is over
'   SendUpdate - When Score Changes
'   SetGameMode - When different game events happen like starting a mode, MB etc.  (ScorbitBuildGameModes helper function shows you how)
' 5) Drop the binaries sQRCode.exe and sToken.exe in your Pup Root so we can create session kets and QRCodes.
' 6) Callbacks
'   Scorbit_Paired      - Called when machine is successfully paired.  Hide QRCode and play a sound
'   Scorbit_PlayerClaimed - Called when player is claimed.  Hide QRCode, play a sound and display name
'
'
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' TABLE CUSTOMIZATION START HERE

Sub StartScorbit
  If IsNull(Scorbit) Then
    If ScorbitActive = 1 then
      ScorbitExesCheck ' Check the exe's are in the tables folder.
      If ScorbitActive = 1 then ' check again as the check exes may have disabled scorbit
        Set Scorbit = New ScorbitIF
        If Scorbit.DoInit(3282, ScorbitQRFolder, myVersion, "GRV7Q-MJNez-VPIN") then  ' Prod
          tmrScorbit.Interval=2000
          tmrScorbit.UserValue = 0
          tmrScorbit.Enabled=True
          Scorbit.UploadLog = 1
        End If
      End If
    End If
  End If
End Sub

Sub Scorbit_NeedsPairing()                ' Scorbit callback when new machine needs pairing
  LoadTexture "", TablesDir & "\" & ScorbitQRFolder & "\QRcode.png"
  ScorbitFlasher.ImageA = "QRcode"
  ScorbitFlasher.Visible = True
End Sub

Sub Scorbit_Paired()
  ScorbitFlasher.Visible = False
  If ScorbitActive = 1 then
    ScorbitExesCheck ' Check the exe's are in the tables folder.
    If ScorbitActive = 1 then ' check again as the check exes may have disabled scorbit
      If InProgress And BallsPlayed = 1 Then
        LoadTexture "", TablesDir & "\" & ScorbitQRFolder & "\QRclaim.png"
        ScorbitFlasher.ImageA = "QRclaim"
        ScorbitFlasher.Visible = True
      End If
    End If
  End If
End Sub

Sub Scorbit_PlayerClaimed(PlayerNum, PlayerName)  ' Scorbit callback when QR Is Claimed

End Sub

Sub ScorbitClaimQR(bShow)           '  Show QRCode on first ball for users to claim this position

End Sub

Sub ScorbitBuildGameModes()   ' Custom function to build the game modes for better stats
  'dim GameModeStr
  'Scorbit.SetGameMode(GameModeStr)"NA{green}:The Hunt
End Sub

Sub Scorbit_LOGUpload(state)  ' Callback during the log creation process.  0=Creating Log, 1=Uploading Log, 2=Done
  Select Case state
    case 0:
      'Debug.print "CREATING LOG"
    case 1:
      'Debug.print "Uploading LOG"
    case 2:
      'Debug.print "LOG Complete"
  End Select
End Sub
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' TABLE CUSTOMIZATION END HERE - NO NEED TO EDIT BELOW THIS LINE

dim Scorbit : Scorbit = Null
' Workaround - Call get a reference to Member Function
Sub tmrScorbit_Timer()                ' Timer to send heartbeat
  Scorbit.DoTimer(tmrScorbit.UserValue)
  tmrScorbit.UserValue=tmrScorbit.UserValue+1
  if tmrScorbit.UserValue>5 then tmrScorbit.UserValue=0
End Sub
Function ScorbitIF_Callback()
  Scorbit.Callback()
End Function
Class ScorbitIF

  Public bSessionActive
  Public bNeedsPairing
  Private bUploadLog
  Private bActive
  Private LOGFILE(10000000)
  Private LogIdx

  Private bProduction

  Private TypeLib
  Private MyMac
  Private Serial
  Private MyUUID
  Private TableVersion

  Private SessionUUID
  Private SessionSeq
  Private SessionTimeStart
  Private bRunAsynch
  Private bWaitResp
  Private GameMode
  Private GameModeOrig    ' Non escaped version for log
  Private VenueMachineID
  Private CachedPlayerNames(4)
  Private SaveCurrentPlayer

  Public bEnabled
  Private sToken
  Private machineID
  Private dirQRCode
  Private opdbID
  Private wsh

  Private objXmlHttpMain
  Private objXmlHttpMainAsync
  Private fso
  Private Domain

  Public Sub Class_Initialize()
    bActive="false"
    bSessionActive=False
    bEnabled=False
  End Sub

  Property Let UploadLog(bValue)
    bUploadLog = bValue
  End Property

  Sub DoTimer(bInterval)  ' 2 second interval
    dim holdScores(4)
    dim i

    if bInterval=0 then
      SendHeartbeat()
    elseif bRunAsynch then ' Game in play
      Scorbit.SendUpdate Score(1), Score(2), Score(3), Score(4), BallsPlayed, player, players
    End if
  End Sub

  Function GetName(PlayerNum) ' Return Parsed Players name
    if PlayerNum<1 or PlayerNum>4 then
      GetName=""
    else
      GetName=CachedPlayerNames(PlayerNum)
    End if
  End Function

  Function DoInit(MyMachineID, Directory_PupQRCode, Version, opdb)
    dim Nad
    Dim EndPoint
    Dim resultStr
    Dim UUIDParts
    Dim UUIDFile

    bProduction=1
'   bProduction=0
    SaveCurrentPlayer=0
    VenueMachineID=""
    bWaitResp=False
    bRunAsynch=False
    DoInit=False
    opdbID=opdb
    dirQrCode=Directory_PupQRCode
    MachineID=MyMachineID
    TableVersion=version
    bNeedsPairing=False
    if bProduction then
      domain = "api.scorbit.io"
    else
      domain = "staging.scorbit.io"
      domain = "scorbit-api-staging.herokuapp.com"
    End if
    'msgbox "production: " & bProduction
    Set fso = CreateObject("Scripting.FileSystemObject")
    dim objLocator:Set objLocator = CreateObject("WbemScripting.SWbemLocator")
    Dim objService:Set objService = objLocator.ConnectServer(".", "root\cimv2")
    Set objXmlHttpMain = CreateObject("Msxml2.ServerXMLHTTP")
    Set objXmlHttpMainAsync = CreateObject("Microsoft.XMLHTTP")
    objXmlHttpMain.onreadystatechange = GetRef("ScorbitIF_Callback")
    Set wsh = CreateObject("WScript.Shell")

    ' Get Mac for Serial Number
    dim Nads: set Nads = objService.ExecQuery("Select * from Win32_NetworkAdapter where physicaladapter=true")
    for each Nad in Nads
      if not isnull(Nad.MACAddress) then
        'msgbox "Using MAC Addresses:" & Nad.MACAddress & " From Adapter:" & Nad.description
        MyMac=replace(Nad.MACAddress, ":", "")
        Exit For
      End if
    Next
    Serial=eval("&H" & mid(MyMac, 5))
'   Serial=123456
    debug.print "Serial: " & Serial
    serial = serial + MachineID
    debug.print "New Serial with machine id: " & Serial

    ' Get System UUID
    set Nads = objService.ExecQuery("SELECT * FROM Win32_ComputerSystemProduct")
    for each Nad in Nads
      'msgbox "Using UUID:" & Nad.UUID
      MyUUID=Nad.UUID
      Exit For
    Next

    if MyUUID="" then
      MsgBox "SCORBIT - Can get UUID, Disabling."
      Exit Function
    elseif MyUUID="03000200-0400-0500-0006-000700080009" or ScorbitAlternateUUID then
      If fso.FolderExists(UserDirectory) then
        If fso.FileExists(UserDirectory & "ScorbitUUID.dat") then
          Set UUIDFile = fso.OpenTextFile(UserDirectory & "ScorbitUUID.dat",1)
          MyUUID = UUIDFile.ReadLine()
          UUIDFile.Close
          Set UUIDFile = Nothing
        Else
          MyUUID=GUID()
          Set UUIDFile=fso.CreateTextFile(UserDirectory & "ScorbitUUID.dat",True)
          UUIDFile.WriteLine MyUUID
          UUIDFile.Close
          Set UUIDFile=Nothing
    End if
      End if
    End if

    ' Clean UUID
    UUIDParts=split(MyUUID, "-")
    'msgbox UUIDParts(0)
    MyUUID=LCASE(Hex(eval("&h" & UUIDParts(0))+MyMachineID)) & UUIDParts(1) &  UUIDParts(2) &  UUIDParts(3) & UUIDParts(4)     ' Add MachineID to UUID
    'msgbox UUIDParts(0)
    MyUUID=LPad(MyUUID, 32, "0")
'   MyUUID=Replace(MyUUID, "-",  "")
    Debug.print "MyUUID:" & MyUUID


' Debug
'   myUUID="adc12b19a3504453a7414e722f58737f"
'   Serial="123456778"

    'create own folder for table QRCodes TablesDir & "\" & dirQrCode
    If Not fso.FolderExists(TablesDir & "\" & dirQrCode) then
      fso.CreateFolder(TablesDir & "\" & dirQrCode)
    end if

    ' Authenticate and get our token
    if getStoken() then
      bEnabled=True
'     SendHeartbeat
      DoInit=True
    End if
  End Function



  Sub Callback()
    Dim ResponseStr
    Dim i
    Dim Parts
    Dim Parts2
    Dim Parts3
    if bEnabled=False then Exit Sub

    if bWaitResp and objXmlHttpMain.readystate=4 then
      'Debug.print "CALLBACK: " & objXmlHttpMain.Status & " " & objXmlHttpMain.readystate
      if objXmlHttpMain.Status=200 and objXmlHttpMain.readystate = 4 then
        ResponseStr=objXmlHttpMain.responseText
        'Debug.print "RESPONSE: " & ResponseStr

        ' Parse Name
        if CachedPlayerNames(SaveCurrentPlayer-1)="" then  ' Player doesnt have a name
          if instr(1, ResponseStr, "cached_display_name") <> 0 Then ' There are names in the result
            Parts=Split(ResponseStr,",{")             ' split it
            if ubound(Parts)>=SaveCurrentPlayer-1 then        ' Make sure they are enough avail
              if instr(1, Parts(SaveCurrentPlayer-1), "cached_display_name")<>0 then  ' See if mine has a name
                CachedPlayerNames(SaveCurrentPlayer-1)=GetJSONValue(Parts(SaveCurrentPlayer-1), "cached_display_name")    ' Get my name
                CachedPlayerNames(SaveCurrentPlayer-1)=Replace(CachedPlayerNames(SaveCurrentPlayer-1), """", "")
                Scorbit_PlayerClaimed SaveCurrentPlayer, CachedPlayerNames(SaveCurrentPlayer-1)
                'Debug.print "Player Claim:" & SaveCurrentPlayer & " " & CachedPlayerNames(SaveCurrentPlayer-1)
              End if
            End if
          End if
        else                            ' Check for unclaim
          if instr(1, ResponseStr, """player"":null")<>0 Then ' Someone doesnt have a name
            Parts=Split(ResponseStr,"[")            ' split it
'Debug.print "Parts:" & Parts(1)
            Parts2=Split(Parts(1),"}")              ' split it
            for i = 0 to Ubound(Parts2)
'Debug.print "Parts2:" & Parts2(i)
            if instr(1, Parts2(i), """player"":null")<>0 Then
                CachedPlayerNames(i)=""
              End if
            Next
          End if
        End if
      End if
      bWaitResp=False
    End if
  End Sub



  Public Sub StartSession()
    if bEnabled=False then Exit Sub
'msgbox  "Scorbit Start Session"
    CachedPlayerNames(0)=""
    CachedPlayerNames(1)=""
    CachedPlayerNames(2)=""
    CachedPlayerNames(3)=""
    bRunAsynch=True
    bActive="true"
    bSessionActive=True
    SessionSeq=0
    SessionUUID=GUID()
    SessionTimeStart=GameTime
    LogIdx=0
    SendUpdate 0, 0, 0, 0, 1, 1, 1
  End Sub

  Public Sub StopSession(P1Score, P2Score, P3Score, P4Score, NumberPlayers)
    StopSession2 P1Score, P2Score, P3Score, P4Score, NumberPlayers, False
  End Sub

  Public Sub StopSession2(P1Score, P2Score, P3Score, P4Score, NumberPlayers, bCancel)
    Dim i
    dim objFile
    if bEnabled=False then Exit Sub
    if bSessionActive=False then Exit Sub
Debug.print "Scorbit Stop Session"

    bRunAsynch=False
    bActive="false"
    bSessionActive=False
    SendUpdate P1Score, P2Score, P3Score, P4Score, -1, -1, NumberPlayers
'   SendHeartbeat

    if bUploadLog and LogIdx<>0 and bCancel=False then
      Debug.print "Creating Scorbit Log: Size" & LogIdx
      Scorbit_LOGUpload(0)
'     Set objFile = fso.CreateTextFile(puplayer.getroot&"\" & cGameName & "\sGameLog.csv")
      Set objFile = fso.CreateTextFile(TablesDir & "\" & dirQRCode & "\sGameLog.csv")
      For i = 0 to LogIdx-1
        objFile.Writeline LOGFILE(i)
      Next
      objFile.Close
      LogIdx=0
      Scorbit_LOGUpload(1)
'     pvPostFile "https://" & domain & "/api/session_log/", puplayer.getroot&"\" & cGameName & "\sGameLog.csv", False
      pvPostFile "https://" & domain & "/api/session_log/", TablesDir & "\" & dirQRCode & "\sGameLog.csv", False
      Scorbit_LOGUpload(2)
      on error resume next
'     fso.DeleteFile(puplayer.getroot&"\" & cGameName & "\sGameLog.csv")
      fso.DeleteFile(TablesDir & "\" & dirQRCode & "\sGameLog.csv")
      on error goto 0
    End if

  End Sub

  Public Sub SetGameMode(GameModeStr)
    GameModeOrig=GameModeStr
    GameMode=GameModeStr
    GameMode=Replace(GameMode, ":", "%3a")
    GameMode=Replace(GameMode, ";", "%3b")
    GameMode=Replace(GameMode, " ", "%20")
    GameMode=Replace(GameMode, "{", "%7B")
    GameMode=Replace(GameMode, "}", "%7D")
  End sub

  Public Sub SendUpdate(P1Score, P2Score, P3Score, P4Score, CurrentBall, CurrentPlayer, NumberPlayers)
    SendUpdateAsynch P1Score, P2Score, P3Score, P4Score, CurrentBall, CurrentPlayer, NumberPlayers, bRunAsynch
  End Sub

  Public Sub SendUpdateAsynch(P1Score, P2Score, P3Score, P4Score, CurrentBall, CurrentPlayer, NumberPlayers, bAsynch)
    dim i
    Dim PostData
    Dim resultStr
    dim LogScores(4)

    if bUploadLog then
      if NumberPlayers>=1 then LogScores(0)=P1Score
      if NumberPlayers>=2 then LogScores(1)=P2Score
      if NumberPlayers>=3 then LogScores(2)=P3Score
      if NumberPlayers>=4 then LogScores(3)=P4Score
      LOGFILE(LogIdx)=DateDiff("S", "1/1/1970", Now()) & "," & LogScores(0) & "," & LogScores(1) & "," & LogScores(2) & "," & LogScores(3) & ",,," &  CurrentPlayer & "," & CurrentBall & ",""" & GameModeOrig & """"
      LogIdx=LogIdx+1
    End if

    if bEnabled=False then Exit Sub
    if bWaitResp then exit sub ' Drop message until we get our next response
'   debug.print "Current players: " & CurrentPlayer
    SaveCurrentPlayer=CurrentPlayer
'   PostData = "session_uuid=" & SessionUUID & "&session_time=" & DateDiff("S", "1/1/1970", Now()) & _
'         "&session_sequence=" & SessionSeq & "&active=" & bActive
    PostData = "session_uuid=" & SessionUUID & "&session_time=" & GameTime-SessionTimeStart+1 & _
          "&session_sequence=" & SessionSeq & "&active=" & bActive

    SessionSeq=SessionSeq+1
    if NumberPlayers > 0 then
      for i = 0 to NumberPlayers-1
        PostData = PostData & "&current_p" & i+1 & "_score="
        if i <= NumberPlayers-1 then
                    if i = 0 then PostData = PostData & P1Score
                    if i = 1 then PostData = PostData & P2Score
                    if i = 2 then PostData = PostData & P3Score
                    if i = 3 then PostData = PostData & P4Score
        else
          PostData = PostData & "-1"
        End if
      Next

      PostData = PostData & "&current_ball=" & CurrentBall & "&current_player=" & CurrentPlayer
      if GameMode<>"" then PostData=PostData & "&game_modes=" & GameMode
    End if
    resultStr = PostMsg("https://" & domain, "/api/entry/", PostData, bAsynch)
    if resultStr<>"" then Debug.print "SendUpdate Resp:" & resultStr
  End Sub

' PRIVATE BELOW
  Private Function LPad(StringToPad, Length, CharacterToPad)
    Dim x : x = 0
    If Length > Len(StringToPad) Then x = Length - len(StringToPad)
    LPad = String(x, CharacterToPad) & StringToPad
  End Function

  Private Function GUID()
    Dim TypeLib
    Set TypeLib = CreateObject("Scriptlet.TypeLib")
    GUID = Mid(TypeLib.Guid, 2, 36)

'   Set wsh = CreateObject("WScript.Shell")
'   Set fso = CreateObject("Scripting.FileSystemObject")
'
'   dim rc
'   dim result
'   dim objFileToRead
'   Dim sessionID:sessionID=puplayer.getroot&"\" & cGameName & "\sessionID.txt"
'
'   on error resume next
'   fso.DeleteFile(sessionID)
'   On error goto 0
'
'   rc = wsh.Run("powershell -Command ""(New-Guid).Guid"" | out-file -encoding ascii " & sessionID, 0, True)
'   if FileExists(sessionID) and rc=0 then
'     Set objFileToRead = fso.OpenTextFile(sessionID,1)
'     result = objFileToRead.ReadLine()
'     objFileToRead.Close
'     GUID=result
'   else
'     MsgBox "Cant Create SessionUUID through powershell. Disabling Scorbit"
'     bEnabled=False
'   End if

  End Function

  Private Function GetJSONValue(JSONStr, key)
    dim i
    Dim tmpStrs,tmpStrs2
    if Instr(1, JSONStr, key)<>0 then
      tmpStrs=split(JSONStr,",")
      for i = 0 to ubound(tmpStrs)
        if instr(1, tmpStrs(i), key)<>0 then
          tmpStrs2=split(tmpStrs(i),":")
          GetJSONValue=tmpStrs2(1)
          exit for
        End if
      Next
    End if
  End Function

  Private Sub SendHeartbeat()
    Dim resultStr
    dim TmpStr
    Dim Command
    Dim rc
'   Dim QRFile:QRFile=puplayer.getroot&"\" & cGameName & "\" & dirQrCode
    Dim QRFile:QRFile=TablesDir & "\" & dirQrCode
    if bEnabled=False then Exit Sub
    resultStr = GetMsgHdr("https://" & domain, "/api/heartbeat/", "Authorization", "SToken " & sToken)
'Debug.print "Heartbeat Resp:" & resultStr
    If VenueMachineID="" then

      if resultStr<>"" and Instr(resultStr, """unpaired"":true")=0 then   ' We Paired
        bNeedsPairing=False
        debug.print "Paired"
        Scorbit_Paired()
      else
        debug.print "Needs Pairing"
        bNeedsPairing=True
        Scorbit_NeedsPairing()
'       if not FScorbitQRIcon.visible then showQRPairImage
      End if

      TmpStr=GetJSONValue(resultStr, "venuemachine_id")
      if TmpStr<>"" then
        VenueMachineID=TmpStr
'Debug.print "VenueMachineID=" & VenueMachineID
'       Command = puplayer.getroot&"\" & cGameName & "\sQRCode.exe " & VenueMachineID & " " & opdbID & " " & QRFile
        debug.print "RUN sqrcode"
        Command = """" & TablesDir & "\sQRCode.exe"" " & VenueMachineID & " " & opdbID & " """ & QRFile & """"
'       msgbox Command
        rc = wsh.Run(Command, 0, False)
      End if
    End if
  End Sub

  Private Function getStoken()
    Dim result
    Dim results
'   dim wsh
    Dim tmpUUID:tmpUUID="adc12b19a3504453a7414e722f58736b"
    Dim tmpVendor:tmpVendor="vscorbitron"
    Dim tmpSerial:tmpSerial="999990104"
'   Dim QRFile:QRFile=puplayer.getroot&"\" & cGameName & "\" & dirQrCode
    Dim QRFile:QRFile=TablesDir & "\" & dirQrCode
'   Dim sTokenFile:sTokenFile=puplayer.getroot&"\" & cGameName & "\sToken.dat"
    Dim sTokenFile:sTokenFile=QRFile & "\sToken.dat"

    ' Set everything up
    tmpUUID=MyUUID
    tmpVendor="vpin"
    tmpSerial=Serial

    on error resume next
    fso.DeleteFile("""" & sTokenFile & """")
    On error goto 0

    ' get sToken and generate QRCode
'   Set wsh = CreateObject("WScript.Shell")
    Dim waitOnReturn: waitOnReturn = True
    Dim windowStyle: windowStyle = 0
    Dim Command
    Dim rc
    Dim objFileToRead

'   msgbox """" & " 55"

'   Command = puplayer.getroot&"\" & cGameName & "\sToken.exe " & tmpUUID & " " & tmpVendor & " " &  tmpSerial & " " & MachineID & " " & QRFile & " " & sTokenFile & " " & domain
    debug.print "RUN sToken"
    Command = """" & TablesDir & "\sToken.exe"" " & tmpUUID & " " & tmpVendor & " " &  tmpSerial & " " & MachineID & " """ & QRFile & """ """ & sTokenFile & """ " & domain
'msgbox "RUNNING:" & Command
    rc = wsh.Run(Command, windowStyle, waitOnReturn)
'msgbox "Return:" & rc
'   if FileExists(puplayer.getroot&"\" & cGameName & "\sToken.dat") and rc=0 then
'   msgbox """" & TablesDir & "\sToken.dat"""
    if FileExists(sTokenFile) and rc=0 then
'     Set objFileToRead = fso.OpenTextFile(puplayer.getroot&"\" & cGameName & "\sToken.dat",1)
'     msgbox """" & TablesDir & "\sToken.dat"""
      Set objFileToRead = fso.OpenTextFile(sTokenFile,1)
      result = objFileToRead.ReadLine()
      objFileToRead.Close
      Set objFileToRead = Nothing
'msgbox result

      if Instr(1, result, "Invalid timestamp")<> 0 then
        MsgBox "Scorbit Timestamp Error: Please make sure the time on your system is exact"
        getStoken=False
      elseif Instr(1, result, "Internal Server error")<> 0 then
        MsgBox "Scorbit Internal Server error ??"
        getStoken=False
      elseif Instr(1, result, ":")<>0 then
        results=split(result, ":")
        sToken=results(1)
        sToken=mid(sToken, 3, len(sToken)-4)
Debug.print "Got TOKEN:" & sToken
        getStoken=True
      Else
Debug.print "ERROR:" & result
        getStoken=False
      End if
    else
'msgbox "ERROR No File:" & rc
    End if

  End Function

  private Function FileExists(FilePath)
    If fso.FileExists(FilePath) Then
      FileExists=CBool(1)
    Else
      FileExists=CBool(0)
    End If
  End Function

  Private Function GetMsg(URLBase, endpoint)
    GetMsg = GetMsgHdr(URLBase, endpoint, "", "")
  End Function

  Private Function GetMsgHdr(URLBase, endpoint, Hdr1, Hdr1Val)
    Dim Url
    Url = URLBase + endpoint & "?session_active=" & bActive
'Debug.print "Url:" & Url  & "  Async=" & bRunAsynch
    objXmlHttpMain.open "GET", Url, bRunAsynch
'   objXmlHttpMain.setRequestHeader "Content-Type", "text/xml"
    objXmlHttpMain.setRequestHeader "Cache-Control", "no-cache"
    if Hdr1<> "" then objXmlHttpMain.setRequestHeader Hdr1, Hdr1Val

'   on error resume next
      err.clear
      objXmlHttpMain.send ""
      if err.number=-2147012867 then
        MsgBox "Multiplayer Server is down (" & err.number & ") " & Err.Description
        bEnabled=False
      elseif err.number <> 0 then
        debug.print "Server error: (" & err.number & ") " & Err.Description
      End if
      if bRunAsynch=False then
          Debug.print "Status: " & objXmlHttpMain.status
          If objXmlHttpMain.status = 200 Then
            GetMsgHdr = objXmlHttpMain.responseText
        Else
            GetMsgHdr=""
          End if
      Else
        bWaitResp=True
        GetMsgHdr=""
      End if
'   On error goto 0

  End Function

  Private Function PostMsg(URLBase, endpoint, PostData, bAsynch)
    Dim Url

    Url = URLBase + endpoint
'debug.print "PostMSg:" & Url & " " & PostData

    objXmlHttpMain.open "POST",Url, bAsynch
    objXmlHttpMain.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    objXmlHttpMain.setRequestHeader "Content-Length", Len(PostData)
    objXmlHttpMain.setRequestHeader "Cache-Control", "no-cache"
    objXmlHttpMain.setRequestHeader "Authorization", "SToken " & sToken
    if bAsynch then bWaitResp=True

    on error resume next
      objXmlHttpMain.send PostData
      if err.number=-2147012867 then
        MsgBox "Multiplayer Server is down (" & err.number & ") " & Err.Description
        bEnabled=False
      elseif err.number <> 0 then
        'debug.print "Multiplayer Server error (" & err.number & ") " & Err.Description
      End if
      If objXmlHttpMain.status = 200 Then
        PostMsg = objXmlHttpMain.responseText
      else
        PostMsg="ERROR: " & objXmlHttpMain.status & " >" & objXmlHttpMain.responseText & "<"
      End if
    On error goto 0
  End Function

  Private Function pvPostFile(sUrl, sFileName, bAsync)
Debug.print "Posting File " & sUrl & " " & sFileName & " " & bAsync & " File:" & Mid(sFileName, InStrRev(sFileName, "\") + 1)
    Dim STR_BOUNDARY:STR_BOUNDARY  = GUID()
    Dim nFile
    Dim baBuffer()
    Dim sPostData
    Dim Response

    '--- read file
    Set nFile = fso.GetFile(sFileName)
    With nFile.OpenAsTextStream()
      sPostData = .Read(nFile.Size)
      .Close
    End With
'   fso.Open sFileName For Binary Access Read As nFile
'   If LOF(nFile) > 0 Then
'     ReDim baBuffer(0 To LOF(nFile) - 1) As Byte
'     Get nFile, , baBuffer
'     sPostData = StrConv(baBuffer, vbUnicode)
'   End If
'   Close nFile

    '--- prepare body
    sPostData = "--" & STR_BOUNDARY & vbCrLf & _
      "Content-Disposition: form-data; name=""uuid""" & vbCrLf & vbCrLf & _
      SessionUUID & vbcrlf & _
      "--" & STR_BOUNDARY & vbCrLf & _
      "Content-Disposition: form-data; name=""log_file""; filename=""" & SessionUUID & ".csv""" & vbCrLf & _
      "Content-Type: application/octet-stream" & vbCrLf & vbCrLf & _
      sPostData & vbCrLf & _
      "--" & STR_BOUNDARY & "--"

'Debug.print "POSTDATA: " & sPostData & vbcrlf

    '--- post
    With objXmlHttpMain
      .Open "POST", sUrl, bAsync
      .SetRequestHeader "Content-Type", "multipart/form-data; boundary=" & STR_BOUNDARY
      .SetRequestHeader "Authorization", "SToken " & sToken
      .Send sPostData ' pvToByteArray(sPostData)
      If Not bAsync Then
        Response= .ResponseText
        pvPostFile = Response
Debug.print "Upload Response: " & Response
      End If
    End With

  End Function

  Private Function pvToByteArray(sText)
    pvToByteArray = StrConv(sText, 128)   ' vbFromUnicode
  End Function

End Class
'  END SCORBIT
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


''QRView support by iaakki
Function GetTablesFolder()
    Dim GTF
    Set GTF = CreateObject("Scripting.FileSystemObject")
    GetTablesFolder= GTF.GetParentFolderName(userdirectory) & "\Tables"
    set GTF = nothing
End Function

' Checks that all needed binaries are available in correct place..
Sub ScorbitExesCheck
  dim fso
  Set fso = CreateObject("Scripting.FileSystemObject")

  If fso.FileExists(TablesDir & "\sToken.exe") then
'   msgbox "Stoken.exe found at: " & TablesDir & "\sToken.exe"
  else
    msgbox "Stoken.exe NOT found at: " & TablesDir & "\sToken.exe Disabling Scorbit for now."
    Scorbitactive = 0
    SaveValue cGameName, "SCORBIT", ScorbitActive
  end if

  If fso.FileExists(TablesDir & "\sQRCode.exe") then
'   msgbox "sQRCode.exe found at: " & TablesDir & "\sQRCode.exe"
  else
    msgbox "sQRCode.exe NOT found at: " & TablesDir & "\sQRCode.exe Disabling Scorbit for now."
    Scorbitactive = 0
    SaveValue cGameName, "SCORBIT", ScorbitActive
  end if
end sub

'*        Williams' Magic City (1967)



option explicit
Randomize



' *******   User Options   ********

Const RoomBrightness =  0.2875        'Room brightness - Value between 0 and 1 (0=Dark ... 1=Bright)

Const DynamicBallShadowsOn = 1    '0 = no dynamic ball shadow ("triangles" near slings and such), 1 = enable dynamic ball shadow
Const AmbientBallShadowOn = 0   '0 = Static shadow under ball ("flasher" image, like JP's)
                  '1 = Moving ball shadow ("primitive" object, like ninuzzu's) - This is the only one that shows up on the pf when in ramps and fades when close to lights!
                  '2 = flasher image shadow, but it moves like ninuzzu's
' ******* END User Options ********



Const Ballsize = 50
Const BallMass = 1

ExecuteGlobal GetTextFile("core.vbs")

On Error Resume Next
ExecuteGlobal GetTextFile("Controller.vbs")
If Err Then MsgBox "Unable to open Controller.vbs. Ensure that it is in the scripts folder."
On Error Goto 0

Const cGameName = "MagicCity_1967"

'
'Const ShadowFlippersOn = true
'Const ShadowBallOn = true
'
'Const ShadowConfigFile = false
'
Dim Controller  ' B2S
Dim B2SScore  ' B2S Score Displayed

Const HSFileName="MagicCity_67VPX.txt"
Const B2STableName="MagicCity_1967"
Const LMEMTableConfig="LMEMTables.txt"
Const LMEMShadowConfig="LMEMShadows.txt"
Dim EnableBallShadow
Dim EnableFlipperShadow

Dim tablewidth: tablewidth = Table1.width
Dim tableheight: tableheight = Table1.height


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

Dim AltRelayValue, ZeroNineUnit

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

Dim Replay1Table(20)
Dim Replay2Table(20)
Dim Replay3Table(20)
Dim Replay4Table(20)
Dim ReplayBallsTable1(16)
Dim ReplayBallsTable2(16)
Dim ReplayBallsTable3(16)

Dim ReplayTableSet
Dim ReplayLevel
Dim ReplayBalls
Dim ReplayTableMax
Dim ReplayBallsTableMax

Dim BaseHit,BaseHitCounter

Dim LeftTargetFlag,RightTargetFlag

Dim FootballPosition,FootballResetFlag,FootballScore

Dim ImpulseReelCount,ScoreMotorPosition

Dim TempScore, TempAdvance, KickoffFlag, AdvanceRelaySwitch

Dim StartGameState, SpecialsFlag

Dim KickerCounter

Dim MagicCityCompleted

Sub Table1_Init()

  vpmTimer.AddTimer 1000, "WarmUpDone '"

  ChangeRoomBrightness RoomBrightness

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
' LoadLMEMConfig2
  HideOptions
  SetupReplayTables
  PlasticsOff
  BumpersOff
  OperatorMenu=0
  StartGameState=0
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
  Bumper1.Threshold = 1
  Bumper2.Threshold = 1
  Bumper3.Threshold = 1
  Bumper4.Threshold = 1
  LeftSlingShot.Disabled = False
  RightSlingShot.Disabled = False

  TiltReel.SetValue(1)
  ZeroNineUnit=int(Rnd*10)
  LightZeroNine
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

  Bumper1Light.state=1



  ReplayCard.image="BC"+FormatNumber(BallsPerGame,0)

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
  GameOverReel.setvalue(1)
  CreditsReel.SetValue(Credits)
  If Credits > 0 Then DOF 232, 1
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
  LeftSlingShot.TimerEnabled = 1
  RightSlingShot.TimerEnabled = 1

  BumpersOff
  PlasticsOff
  InsertsOff

End Sub

Sub Table1_exit()
  savehs
  SaveLMEMConfig
' SaveLMEMConfig2
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
    PlaySoundAtVol SoundFXDOF("FlipperUp",201,DOFOn,DOFContactors), LeftFlipper, 1
    PlayLoopSoundAtVol "buzzL", LeftFlipper, 1

  End If

  If keycode = RightFlipperKey  and InProgress=true and TableTilted=false Then
    RightFlipper.RotateToEnd
    PlaySoundAtVol SoundFXDOF("FlipperUp",202,DOFOn,DOFContactors), RightFlipper, 1
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

' if keycode = 7 then
'   Player=1
'   AdvanceFootball
' end if
'
' if keycode =8 then
'   ImpulseReelCount=int(rnd*5)
'   ResetFootballMain
' end if


  if keycode=StartGameKey and Credits>0 and InProgress=false and Players=0 and EnteringOptions = 0 then
'GNMOD
    OperatorMenuTimer.Enabled = false
'END GNMODthen
    Credits=Credits-1
    If Credits < 1 Then DOF 232, 0
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
    GameOverReel.setvalue(0)
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

    PlaySoundAtVol"plungerrelease", Plunger, 1
    Plunger.Fire
  End If

  if keycode = LeftFlipperKey then
    OperatorMenuTimer.Enabled = false
  end if

  ' END GNMOD

  If keycode = LeftFlipperKey and InProgress=true and TableTilted=false Then
    LeftFlipper.RotateToStart
    PlaySoundAtVol SoundFXDOF("FlipperDown",201,DOFOff,DOFContactors), LeftFlipper, 1
    StopSound "buzzL"
  End If

  If keycode = RightFlipperKey and InProgress=true and TableTilted=false Then
    RightFlipper.RotateToStart
    StopSound "buzz"
    PlaySoundAtVol SoundFXDOF("FlipperDown",202,DOFOff,DOFContactors), RightFlipper, 1
  End If

End Sub



Sub Drain_Hit()
  Drain.DestroyBall
  PlaySoundAtVol "fx_drain", Drain, 1
  DOF 207,2
' AddBonus

  Pause4Bonustimer.enabled=true


End Sub


Sub Pause4Bonustimer_timer
  Pause4Bonustimer.enabled=0

  NextBallDelay.enabled=true

End Sub


Sub CloseGateTrigger_Hit()

  DOF 238, DOFPulse
End Sub

''***********************
''     Flipper Logos
''***********************
'
'Sub UpdateFlipperLogos_Timer
''    LFlip.ObjRotZ = LeftFlipper.CurrentAngle-126
''    RFlip.ObjRotZ = RightFlipper.CurrentAngle+126
''    LFlipR.ObjRotZ = LeftFlipper.CurrentAngle-126
''    RFlipR.ObjRotZ = RightFlipper.CurrentAngle+126
''  PGate.Rotz = (Gate.CurrentAngle*.75) + 25
' FlipperLSh.RotZ = LeftFlipper.currentangle
' FlipperRSh.RotZ = RightFlipper.currentangle
'End Sub

'***********************
' slingshots
'

Sub RightSlingShot_Slingshot
  RS.VelocityCorrect(activeball)
    PlaySoundAtVol SoundFXDOF("right_slingshot",205,DOFPulse,DOFContactors), ActiveBall, 1
  DOF 206,2
    sling1.TransZ = -20
    RStep = 0
  RightSlingShot_Timer
    RightSlingShot.TimerEnabled = 1
  AddScore(1)
  PlaySound SoundFXDOF("1_Point_Bell",234,DOFPulse,DOFChimes)
End Sub

Sub RightSlingShot_Timer
  Dim bl
  Dim x0, x1, x2: x0 = False: x1 = False: x2 = True
    Select Case RStep
    Case 2: x0 = False: x1 = True:  x2 = False:  sling1.TransZ = -10
    Case 3: x0 = True:  x1 = False: x2 = False:  sling1.TransZ = 0:  RightSlingShot.TimerEnabled = 0
    End Select
  'VLM movable script
  For each bl in rsling_bl : bl.visible = x0 : Next
  For each bl in rsling002_bl : bl.visible = x1 : Next
  For each bl in rsling001_bl : bl.visible = x2 : Next
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  LS.VelocityCorrect(activeball)
    PlaySoundAtVol SoundFXDOF("left_slingshot",203,DOFPulse,DOFContactors), ActiveBall, 1
  DOF 204,2
    sling2.TransZ = -20
    LStep = 0
  LeftSlingShot_Timer
    LeftSlingShot.TimerEnabled = 1
  AddScore(1)
  PlaySound SoundFXDOF("1_Point_Bell",234,DOFPulse,DOFChimes)
End Sub

Sub LeftSlingShot_Timer
  Dim bl
  Dim x0, x1, x2: x0 = False: x1 = False: x2 = True
    Select Case LStep
    Case 2: x0 = False: x1 = True:  x2 = False:  sling2.TransZ = -10
    Case 3: x0 = True:  x1 = False: x2 = False:  sling2.TransZ = 0:  LeftSlingShot.TimerEnabled = 0
    End Select
  'VLM movable script
  For each bl in lsling_bl : bl.visible = x0 : Next
  For each bl in lsling002_bl : bl.visible = x1 : Next
  For each bl in lsling001_bl : bl.visible = x2 : Next
    LStep = LStep + 1
End Sub

'***********************************
' Walls
'***********************************

Sub RubberwallSwitches_Hit(idx)
  if TableTilted=false then
    AddScore(1)
    PlaySound SoundFXDOF("1_Point_Bell",234,DOFPulse,DOFChimes)

  end if
end Sub





Sub Bumper1_Hit
  If TableTilted=false then
    PlaySoundAtVol SoundFXDOF("topbumper_hit",208,DOFPulse,DOFContactors), ActiveBall, 1
    DOF 209,2
    bump1 = 1

    If Bumper1Light.state=1 Then
      AddScore(100)
    else
      AddScore(10)
    End If

    'VLM movable script
    dim bl
    For each bl in Pop_Bumper_WPC_003_BL : bl.transz = -20 : Next
    For each bl in Circle_016_BL
      bl.roty = skirtAY(me,Activeball)
      bl.rotx = skirtAX(me,Activeball)
    Next

    me.TimerEnabled = 1
    end if

End Sub

Sub Bumper1_Timer
  me.TimerEnabled = 0

  'VLM movable script
  dim bl
  For each bl in Pop_Bumper_WPC_003_BL : bl.transz = 0 : Next
  For each bl in Circle_016_BL
    bl.roty = 0
    bl.rotx = 0
  Next

End Sub


Sub Bumper2_Hit
  If TableTilted=false then

    'Bumper3.PlayHit()
    PlaySoundAtVol SoundFXDOF("rightbumper_hit",210,DOFPulse,DOFContactors), ActiveBall, 1
    DOF 211,2
    bump2 = 1

    If Bumper2Light.state=1 then
      AddScore(10)
    else
      AddScore(1)
      PlaySound SoundFXDOF("1_Point_Bell",234,DOFPulse,DOFChimes)
    End If

    'VLM movable script
    dim bl
    For each bl in Pop_Bumper_WPC_002_BL : bl.transz = -20 : Next
    For each bl in Circle_015_BL
      bl.roty = skirtAY(me,Activeball)
      bl.rotx = skirtAX(me,Activeball)
    Next

    me.TimerEnabled = 1
    end if
End Sub

Sub Bumper2_Timer
  me.TimerEnabled = 0

  'VLM movable script
  dim bl
  For each bl in Pop_Bumper_WPC_002_BL : bl.transz = 0 : Next
  For each bl in Circle_015_BL
    bl.roty = 0
    bl.rotx = 0
  Next
End Sub


Sub Bumper3_Hit
  If TableTilted=false then
    'Bumper2.PlayHit()
    PlaySoundAtVol SoundFXDOF("leftbumper_hit",208,DOFPulse,DOFContactors), ActiveBall, 1
    DOF 209,2
    bump1 = 1

    If Bumper2Light.state=1 then
      AddScore(10)
    else
      AddScore(1)
      PlaySound SoundFXDOF("1_Point_Bell",234,DOFPulse,DOFChimes)
    End If

    'VLM movable script
    dim bl
    For each bl in Pop_Bumper_WPC_002_BL : bl.transz = -20 : Next
    For each bl in Circle_037_BL
      bl.roty = skirtAY(me,Activeball)
      bl.rotx = skirtAX(me,Activeball)
    Next

    me.TimerEnabled = 1
    end if
End Sub

Sub Bumper3_Timer
  me.TimerEnabled = 0

  'VLM movable script
  dim bl
  For each bl in Pop_Bumper_WPC_002_BL : bl.transz = 0 : Next
  For each bl in Circle_037_BL
    bl.roty = 0
    bl.rotx = 0
  Next
End Sub


Sub Bumper4_Hit
  If TableTilted=false then
    PlaySoundAtVol SoundFXDOF("bumper1",208,DOFPulse,DOFContactors), ActiveBall, 1
    DOF 209,2
    bump1 = 1

    If Bumper2Light.state=1 then
      AddScore(10)
    else
      AddScore(1)
      PlaySound SoundFXDOF("1_Point_Bell",234,DOFPulse,DOFChimes)
    End If
    ToggleZeroNine

    'VLM movable script
    dim bl
    For each bl in Pop_Bumper_WPC_005_BL : bl.transz = -20 : Next
    For each bl in Circle_019_BL
      bl.roty = skirtAY(me,Activeball)
      bl.rotx = skirtAX(me,Activeball)
    Next

    me.TimerEnabled = 1
    end if
End Sub

Sub Bumper4_Timer
  me.TimerEnabled = 0

  'VLM movable script
  dim bl
  For each bl in Pop_Bumper_WPC_005_BL : bl.transz = 0 : Next
  For each bl in Circle_019_BL
    bl.roty = 0
    bl.rotx = 0
  Next
End Sub


'**
'            SKIRT ANIMATION FUNCTIONS
'**
' NOTE: set bumper object timer to around 150-175 in order to be able
' to actually see the animation, adjust to your liking

'Const PI = 3.1415926
Const SkirtTilt=3        'angle of skirt tilting in degrees

Function SkirtAX(bumper, bumperball)
    skirtAX=cos(skirtA(bumper,bumperball))*(SkirtTilt)        'x component of angle
    if (bumper.y<bumperball.y) then    skirtAX=skirtAX-1    'adjust for ball hit bottom half
End Function

Function SkirtAY(bumper, bumperball)
    skirtAY=sin(skirtA(bumper,bumperball))*(SkirtTilt)        'y component of angle
    if (bumper.x>bumperball.x) then    skirtAY=skirtAY-1    'adjust for ball hit left half
End Function

Function SkirtA(bumper, bumperball)
    dim hitx, hity, dx, dy
    hitx=bumperball.x
    hity=bumperball.y

    dy=Abs(hity-bumper.y)                    'y offset ball at hit to center of bumper
    if dy=0 then dy=0.0000001
    dx=Abs(hitx-bumper.x)                    'x offset ball at hit to center of bumper
    skirtA=(atn(dx/dy)) '/(PI/180)            'angle in radians to ball from center of Bumper1
End Function



'***********************************

Sub Trigger1_Hit()
  If TableTilted=false then
    PlaySoundAtVol "sensor", ActiveBall, 1
    SetMotor(50)
    MagicCity2Lights(ZeroNineUnit).state=1

    Bumper2Light.state=1
    'Bumper3Light.state=1
    'Bumper4Light.state=1
    if ZeroNineUnit<5 then
      MagicCityLights(ZeroNineUnit).state=0
    elseif ZeroNineUnit=5 then
      Bumper1Light.state=1
      Special002.state=1
      Special003.state=1
    else
      MagicCityLights(ZeroNineUnit-1).state=0
    end if
    CheckMagicCity
    ToggleZeroNine
  end if

' Button001.z=-1.5

  'VLM movable script
  dim bl
  For each bl in Sphere_007_BL : bl.transz = -2 : Next

End Sub


Sub Trigger1_Unhit
  'Button001.z=.5

  'VLM movable script
  dim bl
  For each bl in Sphere_007_BL : bl.transz = 0 : Next
end sub


Sub Trigger001_hit
  Update_Wires 9, true
  If TableTilted=false then
    PlaySoundAtVol "sensor", ActiveBall, 1
    AddScore(100)
    MagicCityLight001.state=0
    MagicCity2001.state=1
    CheckMagicCity
  end if
end sub

Sub Trigger001_unhit: Update_Wires 9, false: End Sub


Sub Trigger002_hit
  Update_Wires 1, true
  If TableTilted=false then
    PlaySoundAtVol "sensor", ActiveBall, 1
    AddScore(100)
    MagicCityLight002.state=0
    MagicCity2002.state=1
    CheckMagicCity
  end if
end sub

Sub Trigger002_unhit: Update_Wires 1, false : End Sub


Sub Trigger003_hit
  Update_Wires 2, true
  If TableTilted=false then
    PlaySoundAtVol "sensor", ActiveBall, 1
    AddScore(100)
    MagicCityLight003.state=0
    MagicCity2003.state=1
    CheckMagicCity
  end if
end sub

Sub Trigger003_unhit: Update_Wires 2, false : End Sub


Sub Trigger004_hit
  Update_Wires 10, true
  If TableTilted=false then
    PlaySoundAtVol "sensor", ActiveBall, 1
    AddScore(100)
    MagicCityLight004.state=0
    MagicCity2004.state=1
    CheckMagicCity
  end if
end sub

Sub Trigger004_unhit: Update_Wires 10, false : End Sub


Sub Trigger005_hit
  Update_Wires 11, true
  If TableTilted=false then
    PlaySoundAtVol "sensor", ActiveBall, 1
    AddScore(100)
    MagicCityLight005.state=0
    MagicCity2005.state=1
    CheckMagicCity
  end if
end sub

Sub Trigger005_unhit: Update_Wires 11, false : End Sub


Sub TriggerLeftOutlane_Hit()
  Update_Wires 0, true
  If TableTilted=false then
    PlaySoundAtVol "sensor", ActiveBall, 1
    DOF 228,2
    AddScore(100)
  End If
End Sub

Sub TriggerLeftOutlane_UnHit: Update_Wires 0, true : End Sub

Sub TriggerLeftOutlane2_Hit()
  Update_Wires 3, true
  If TableTilted=false then
    PlaySoundAtVol "sensor", ActiveBall, 1
    DOF 228,2
    If Special001.state=1 then
      AddSpecial
    end if
  End If
End Sub

Sub TriggerLeftOutlane2_UnHit: Update_Wires 3, true : End Sub


Sub TriggerRightOutlane_Hit()
  Update_Wires 5, true
  If TableTilted=false then
    PlaySoundAtVol "sensor", ActiveBall, 1
    DOF 229,2
    AddScore(100)
  End If
End Sub

Sub TriggerRightOutlane_UnHit: Update_Wires 5, true : End Sub


Sub TriggerRightOutlane2_Hit()
  Update_Wires 4, true
  If TableTilted=false then
    PlaySoundAtVol "sensor", ActiveBall, 1
    DOF 229,2
    If Special004.state=1 then
      AddSpecial
    end if
  End If
End Sub

Sub TriggerRightOutlane2_UnHit: Update_Wires 4, true : End Sub

Sub TriggerLeftCenterOutlane_Hit()
  Update_Wires 8, true
  If TableTilted=false then
    PlaySoundAtVol "sensor", ActiveBall, 1
    DOF 229,2
    If Special002.state=1 then
      AddSpecial
    else
      AddScore(100)
    end if

  End If
End Sub

Sub TriggerLeftCenterOutlane_UnHit: Update_Wires 8, true : End Sub

Sub TriggerCenterOutlane_Hit()
  Update_Wires 7, true
  If TableTilted=false then
    PlaySoundAtVol "sensor", ActiveBall, 1
    DOF 229,2
    AddScore(100)

  End If
End Sub

Sub TriggerCenterOutlane_UnHit: Update_Wires 7, true : End Sub

Sub TriggerRightCenterOutlane_Hit()
  Update_Wires 6, true
  If TableTilted=false then
    PlaySoundAtVol "sensor", ActiveBall, 1
    DOF 229,2
    If Special003.state=1 then
      AddSpecial
    else
      AddScore(100)
    end if

  End If
End Sub

Sub TriggerRightCenterOutlane_UnHit: Update_Wires 6, true : End Sub

Sub CheckMagicCity
  dim tempCounter
  tempCounter=0
  for each obj in MagicCityLights
    if obj.state=0 then tempCounter=tempCounter+1
  next
  MagicCityCompleted=false
  if tempCounter>8 then
    MagicCityCompleted=true

    LightAltRelay
  end if
end sub

Sub Check123

end sub




'****************************************************
Sub AddSpecial()
  PlaySound SoundFXDOF("knocker",230,DOFPulse,DOFContactors)
  DOF 231,2
  Credits=Credits+1
  DOF 232,1
  if Credits>15 then Credits=15
  If B2SOn Then
    Controller.B2SSetCredits Credits
  End If
  CreditsReel.SetValue(Credits)
End Sub

Sub AddSpecial2()
  PlaySound"click"
  Credits=Credits+1
  DOF 232,1
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


sub ToggleAltRelay
  if AltRelayValue=1 then
    AltRelayValue=0

  else
    AltRelayValue=1
  end if
  LightAltRelay
end sub

Sub LightAltRelay
  Special001.state=0
  Special004.state=0
  if AltRelayValue=0 then
    If MagicCityCompleted=true then
      Special001.state=1
    end if
  else
    If MagicCityCompleted=true then
      Special004.state=1
    end if

  end if

end sub

Sub ToggleZeroNine
  ZeroNineUnit=ZeroNineUnit+1
  If ZeroNineUnit>9 then ZeroNineUnit=0
  LightZeroNine
end sub

Sub LightZeroNine
  For each obj in AdvanceLights
    obj.state=0
  next
  AdvanceLights(ZeroNineUnit).state=1
end sub


Sub ResetBallDrops()



  BonusCounter=0
  HoleCounter=0
End Sub


Sub LightsOut
  BonusCounter=0
  HoleCounter=0



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

  Bumper1Light.state=0
  Bumper2Light.state=0
  'Bumper3Light.state=0
  'Bumper4Light.state=0
  MagicCity2006.state=0
  Special002.state=0
  Special003.state=0
  RightTargetFlag=false
  LeftTargetFlag=false
  BonusMultiplier=1


  TableTilted=false
  Bumper1.Threshold = 1
  Bumper2.Threshold = 1
  Bumper3.Threshold = 1
  Bumper4.Threshold = 1
  LeftSlingShot.Disabled = False
  RightSlingShot.Disabled = False
  TiltReel.SetValue(0)

  PlasticsOn

  'CreateBallID BallRelease
  Ballrelease.CreateSizedBall 25
    Ballrelease.Kick 40,7
  DOF 233,2
  'InstructCard.image="IC"+FormatNumber(BallsPerGame,0)+FormatNumber(BallInPlay,0)


  BallInPlayReel.SetValue(BallInPlay)

End Sub

sub delaykgclose_timer
  delaykgclose.enabled=false
  closeg.enabled=true

end sub


 sub openg_timer
    bottgate(bgpos).isdropped=true
    bgpos=bgpos-1
    bottgate(bgpos).isdropped=false
  primgate.RotY=30+(bgpos*10)
     if bgpos=0 then
    PlaySoundAtVol "postup", bottgate, 1
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
    closeg.enabled=false
    dooralreadyopen=0
  end if
end sub

sub resettimer_timer

    rst=rst+1
  if rst>1 and rst<20 then
    ResetReelsToZero(1)
  end if

    if rst=13 then
    'playsound "StartBall1"
    end if
    if rst>=24 then
    newgame
    'newgame
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
  If KickoffFlag=true then exit sub
  NextBallDelay.enabled=false
  nextball

end sub

Sub NewGameHoldTimer_timer
  if KickoffFlag=true then exit sub
  NewGameHoldTimer.enabled=false
  newgame
end sub

sub newgame
  InProgress=true
  debugscore=0
  EMReel1.ResetToZero
  EMReel2.ResetToZero
  MatchTextBox.text=""
  'BallTextBox.text=FormatNumber(BallInPlay,0)
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

  AltRelayValue=1
  BaseHit=0
  BaseHitCounter=0
  SpecialsFlag=false
  Bumper1Light.state=0
  Bumper2Light.state=0
  'Bumper3Light.state=0
  'Bumper4Light.state=0

  BumpersOn
  For each obj in MagicCityLights
    obj.state=1
  next
  For each obj in MagicCity2Lights
    obj.state=0
  next
  for each obj in AdvanceLights
    obj.state=0
  next
  LightTop4.state=1
  Special001.state=0
  Special002.state=0
  Special003.state=0
  Special004.state=0
  LightZeroNine

  ResetDrops
  ResetBalls
  playsound "StartBall1"
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
      If MagicCityCompleted=true then AddSpecial

      BallInPlayReel.SetValue(0)
      CanPlayReel.SetValue(0)
      GameOverReel.setvalue(1)
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
      Bumper1.Threshold = 1000
      Bumper2.Threshold = 1000
      Bumper3.Threshold = 1000
      Bumper4.Threshold = 1000
      LeftSlingShot.Disabled = True
      RightSlingShot.Disabled = True

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



' --- is this dead code? -apophis
' yes  -bord

'Sub IncreaseBonus()
'
'
' If BonusCounter=10 then
'
' else
'   If BonusCounter>0 then
'     Bonus(BonusCounter).state=0
'   end if
'   BonusCounter=BonusCounter+1
'   Bonus(BonusCounter).State=1
''    PlaySound("Score100")
' end if
' if BonusCounter=10 then
'   TopLeftTargetLight.state=1
'   TopRightTargetLight.state=1
' end if
'End Sub
'
'
'Sub BonusBoost_Timer()
' IncreaseBonus
' BonusBoosterCounter=BonusBoosterCounter-1
' If BonusBoosterCounter=0 then
'   BonusBoost.enabled=false
' end if
'
'end sub
'
'Sub CheckForLightSpecial()
'
' if (TopLightA.state=0) and (TopLightB.state=0) and (TopLightC.state=0) then
'   TopRightTargetLight.State=1
'   TopLeftTargetLight.State=1
' end if
'end sub




Sub PlayStartBall_timer()

  PlayStartBall.enabled=false
  PlaySound("StartBall2-5")
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
    ScoreFile.WriteLine 1
    ScoreFile.WriteLine Credits
    scorefile.writeline BallsPerGame
    ScoreFile.WriteLine ReplayLevel
    ScoreFile.WriteLine ReplayBalls
    ScoreFile.WriteLine StartGameState

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
  dim temp6
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
    if HighScore=1 then

      Credits=cdbl(temp2)
      BallsPerGame=cdbl(temp3)
      ReplayLevel=cdbl(temp4)
      ReplayBalls=cdbl(temp5)
      StartGameState=cdbl(temp6)
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

'sub SaveLMEMConfig2
' If ShadowConfigFile=false then exit sub
' Dim FileObj
' Dim LMConfig2
' dim temp1
' dim temp2
' dim tempBS
' dim tempFS
'
' if EnableBallShadow=true then
'   tempBS=1
' else
'   tempBS=0
' end if
' if EnableFlipperShadow=true then
'   tempFS=1
' else
'   tempFS=0
' end if
'
' Set FileObj=CreateObject("Scripting.FileSystemObject")
' If Not FileObj.FolderExists(UserDirectory) then
'   Exit Sub
' End if
' Set LMConfig2=FileObj.CreateTextFile(UserDirectory & LMEMShadowConfig,True)
' LMConfig2.WriteLine tempBS
' LMConfig2.WriteLine tempFS
' LMConfig2.Close
' Set LMConfig2=Nothing
' Set FileObj=Nothing
'
'end Sub
'
'sub LoadLMEMConfig2
' If ShadowConfigFile=false then
'   EnableBallShadow = ShadowBallOn
'   BallShadowUpdate.enabled = ShadowBallOn
'   EnableFlipperShadow = ShadowFlippersOn
'   FlipperLSh.visible = ShadowFlippersOn
'   FlipperRSh.visible = ShadowFlippersOn
'   exit sub
' end if
' Dim FileObj
' Dim LMConfig2
' dim tempC
' dim tempD
' dim tempFS
' dim tempBS
'
'    Set FileObj=CreateObject("Scripting.FileSystemObject")
' If Not FileObj.FolderExists(UserDirectory) then
'   Exit Sub
' End if
' If Not FileObj.FileExists(UserDirectory & LMEMShadowConfig) then
'   Exit Sub
' End if
' Set LMConfig2=FileObj.GetFile(UserDirectory & LMEMShadowConfig)
' Set TextStr2=LMConfig2.OpenAsTextStream(1,0)
' If (TextStr2.AtEndOfStream=True) then
'   Exit Sub
' End if
' tempC=TextStr2.ReadLine
' tempD=TextStr2.Readline
' TextStr2.Close
' tempBS=cdbl(tempC)
' tempFS=cdbl(tempD)
' if tempBS=0 then
'   EnableBallShadow=false
'   BallShadowUpdate.enabled=false
' else
'   EnableBallShadow=true
' end if
' if tempFS=0 then
'   EnableFlipperShadow=false
'   FlipperLSh.visible=false
'   FLipperRSh.visible=false
' else
'   EnableFlipperShadow=true
' end if
' Set LMConfig2=Nothing
' Set FileObj=Nothing
'
'end sub

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
  Bumper1Light.state=0
  Bumper2Light.state=0
  'Bumper1Light.visible=0
  'Bumper2Light.visible=0
  'Bumper3Light.visible=0
  'Bumper4Light.visible=0
end sub

sub BumpersOn
  Bumper1Light.state=1
  Bumper2Light.state=1
  'Bumper1Light.visible=1
  'Bumper2Light.visible=1
  'Bumper3Light.visible=1
  'Bumper4Light.visible=1
end sub

Sub InsertsOff
  For each obj in MagicCityLights: obj.state=0 : next
  For each obj in MagicCity2Lights: obj.state=0 : next
  For each obj in AdvanceLights: obj.state=0 : next
  Special001.state=0
  Special002.state=0
  Special003.state=0
  Special004.state=0
End Sub


Sub PlasticsOn
  Lampz.State(100) = 1
' For each obj in Flashers
'   obj.state=1
' next
end sub

Sub PlasticsOff
  Lampz.State(100) = 0
' For each obj in Flashers
'   obj.state=0
' next
  StopSound "buzz"
  StopSound "buzzL"
end sub

Sub SetupReplayTables

  Replay1Table(1)=1800
  Replay1Table(2)=2000
  Replay1Table(3)=2700
  Replay1Table(4)=1800
  Replay1Table(5)=2100
  Replay1Table(6)=2200
  Replay1Table(7)=2300
  Replay1Table(8)=2400
  Replay1Table(9)=2500
  Replay1Table(10)=2600
  Replay1Table(11)=2700
  Replay1Table(12)=2800
  Replay1Table(13)=2900
  Replay1Table(14)=3000
  Replay1Table(15)=3800
  Replay1Table(16)=3900
  Replay1Table(17)=4000

  Replay2Table(1)=2200
  Replay2Table(2)=2400
  Replay2Table(3)=3200
  Replay2Table(4)=2400
  Replay2Table(5)=3200
  Replay2Table(6)=3300
  Replay2Table(7)=3400
  Replay2Table(8)=3500
  Replay2Table(9)=3600
  Replay2Table(10)=3700
  Replay2Table(11)=3800
  Replay2Table(12)=3900
  Replay2Table(13)=4100
  Replay2Table(14)=4100
  Replay2Table(15)=5100
  Replay2Table(16)=5100
  Replay2Table(17)=5100

  Replay3Table(1)=2700
  Replay3Table(2)=3100
  Replay3Table(3)=4100
  Replay3Table(4)=3000
  Replay3Table(5)=4600
  Replay3Table(6)=4600
  Replay3Table(7)=4600
  Replay3Table(8)=4600
  Replay3Table(9)=4700
  Replay3Table(10)=4800
  Replay3Table(11)=4900
  Replay3Table(12)=5100
  Replay3Table(13)=5200
  Replay3Table(14)=5200
  Replay3Table(15)=6000
  Replay3Table(16)=6000
  Replay3Table(17)=6200

  Replay4Table(1)=3300
  Replay4Table(2)=3500
  Replay4Table(3)=4600
  Replay4Table(4)=3600
  Replay4Table(5)=5700
  Replay4Table(6)=5700
  Replay4Table(7)=5700
  Replay4Table(8)=5700
  Replay4Table(9)=5800
  Replay4Table(10)=5900
  Replay4Table(11)=6000
  Replay4Table(12)=6200
  Replay4Table(13)=6300
  Replay4Table(14)=6300
  Replay4Table(15)=7300
  Replay4Table(16)=7300
  Replay4Table(17)=7300

  ReplayTableMax=3

  ReplayBallsTable1(1)=20
  ReplayBallsTable1(2)=26
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

end sub

Sub RefreshReplayCard
  Dim tempst1
  Dim tempst2
  Dim tempst3

  tempst1=FormatNumber(BallsPerGame,0)
  tempst2=FormatNumber(ReplayLevel,0)
  tempst3=FormatNumber(ReplayBalls,0)

  ReplayCard.image="IC"
  ReplayCard1.image = "SC" + tempst2
  ReplayCard2.image = "BC" + tempst1

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
  PlayerScores(Player).SetValue(Score(Player+1))
  PlayerScoresOn(Player).SetValue(Score(Player+1))
  'PlaySound SoundFXDOF("SpinACard_100_Point_Bell",237,DOFPulse,DOFChimes)
  If Score(Player+1) >= BallReplay1 THEN
    SpecialsFlag=true
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
            TargetLightBase3.state=0         'There are several missing objects in this sub. -apophis
            TargetLightBase31.state=0   'appears to be detritus from a baseball game. Jeff left stuff behind often.  -bord
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
  ImpulseReelCount=ImpulseReelCount+1
  if ImpulseReelCount>4 then
    ImpulseReelCount=0
  end if
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
    'PlayChime(10)
    ToggleAltRelay
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
    PlayChime(100)


    end if

    OldScore = int(OldScore / 10)
    NewScore = int(NewScore / 10)
  ' MsgBox("OldScore="&OldScore&", NewScore="&NewScore)
    if (OldScore Mod 10 <> NewScore Mod 10) then
    'PlayChime(1000)
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
            Case 1
                If LastChime10=1 Then
                    PlaySound SoundFXDOF("1_Point_Bell",234,DOFPulse,DOFChimes)
                    LastChime10=0
                Else
                    PlaySound SoundFXDOF("1_Point_Bell",234,DOFPulse,DOFChimes)
                    LastChime10=1
                End If
            Case 10,100
                If LastChime100=1 Then
                    PlaySound SoundFXDOF("SpinACard_100_Point_Bell",235,DOFPulse,DOFChimes)
                    LastChime100=0
                Else
                    PlaySound SoundFXDOF("SpinACard_100_Point_Bell",235,DOFPulse,DOFChimes)
                    LastChime100=1
                End If

        End Select
    else
        Select Case x
            Case 10
                If LastChime10=1 Then
                    PlaySound SoundFXDOF("1_Point_Bell",234,DOFPulse,DOFChimes)
                    LastChime10=0
                Else
                    PlaySound SoundFXDOF("1_Point_Bell",234,DOFPulse,DOFChimes)
                    LastChime10=1
                End If
            Case 100
                If LastChime100=1 Then
                    PlaySound SoundFXDOF("SpinACard_100_Point_Bell",235,DOFPulse,DOFChimes)
                    LastChime100=0
                Else
                    PlaySound SoundFXDOF("SpinACard_100_Point_Bell",235,DOFPulse,DOFChimes)
                    LastChime100=1
                End If
            Case 1000
                If LastChime1000=1 Then
                    PlaySound SoundFXDOF("",236,DOFPulse,DOFChimes)
                    LastChime1000=0
                Else
                    PlaySound SoundFXDOF("",236,DOFPulse,DOFChimes)
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

Const tnob = 1 ' total number of balls
Const lob = 0
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
    ' Comment the next line if you are not implementing Dyanmic Ball Shadows
    If AmbientBallShadowOn = 0 Then BallShadowA(b).visible = 0
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

    ' "Static" Ball Shadows
    ' Comment the next If block, if you are not implementing the Dynamic Ball Shadows
    If AmbientBallShadowOn = 0 Then
      If BOT(b).Z > 30 Then
        BallShadowA(b).height = BOT(b).z - BallSize / 4   'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping the ramp
      Else
        BallShadowA(b).height = 0.1
      End If
      BallShadowA(b).Y = BOT(b).Y + offsetY
      BallShadowA(b).X = BOT(b).X + offsetX
      BallShadowA(b).visible = 1
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
HSScore(1) = 3500
HSScore(2) = 3000
HSScore(3) = 2800
HSScore(4) = 2500
HSScore(5) = 2000

HSName(1) = "AAA"
HSName(2) = "ZZZ"
HSName(3) = "XXX"
HSName(4) = "ABC"
HSName(5) = "BBB"

HSScore(6) = 25
HSScore(7) = 20
HSScore(8) = 18
HSScore(9) = 14
HSScore(10) = 10

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
    SetHSLine 4, "TOP POINTS1"
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
    SetHSLine 4, "TOP POINTS"+FormatNumber(HSTimerCount,0)
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
Const OptionLinesToMark="111001011"
Const OptionLine1="" 'do not use this line
Const OptionLine2="" 'do not use this line
Const OptionLine3="" 'do not use this line
Const OptionLine4=""
Const OptionLine5=""
Const OptionLine6="S and P lights"
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


        SetOptLine 5, tempstring
      Case 5:
        SetOptLine 6, tempstring

        SetOptLine 7, tempstring

      Case 6:
        SetOptLine 8, tempstring
        Select case StartGameState
          case 0:
            tempstring = "Liberal"
          case 1:
            tempstring = "Moderate"
          case 2:
            tempstring = "Conservative"
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

    elseif CurrentOption = 6 then
      StartGameState=StartGameState+1
      if StartGameState>2 then StartGameState=0
      DisplayAllOptions
    elseif CurrentOption = 8 or CurrentOption = 9 then
        if OptionCHS=1 then
          HSScore(1) = 3500
          HSScore(2) = 3000
          HSScore(3) = 2800
          HSScore(4) = 2500
          HSScore(5) = 2000

          HSName(1) = "AAA"
          HSName(2) = "ZZZ"
          HSName(3) = "XXX"
          HSName(4) = "ABC"
          HSName(5) = "BBB"

          HSScore(6) = 25
          HSScore(7) = 20
          HSScore(8) = 18
          HSScore(9) = 14
          HSScore(10) = 10

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
'****  LAMPZ by nFozzy
'******************************************************

Dim NullFader : set NullFader = new NullFadingObject
Dim Lampz : Set Lampz = New LampFader
InitLampsNF               ' Setup lamp assignments
LampTimer.Interval = -1
LampTimer.Enabled = 1

Sub LampTimer_Timer()
    'apophis - Use the InPlayState of the 1st light in the Lampz.obj array to set the Lampz.state
    dim idx : for idx = 0 to 150
        if Lampz.IsLight(idx) then
            if IsArray(Lampz.obj(idx)) then
                dim tmp : tmp = Lampz.obj(idx)
                Lampz.state(idx) = tmp(0).GetInPlayStateBool
                'debug.print tmp(0).name & " " &  tmp(0).GetInPlayStateBool & " " & tmp(0).IntensityScale  & vbnewline
            Else
                Lampz.state(idx) = Lampz.obj(idx).GetInPlayStateBool
                'debug.print Lampz.obj(idx).name & " " &  Lampz.obj(idx).GetInPlayStateBool & " " & Lampz.obj(idx).IntensityScale  & vbnewline
            end if
        end if
    Next
  Lampz.Update2 'update (fading logic only)
End Sub

Sub InitLampsNF()

  'Filtering (comment out to disable)
  Lampz.Filter = "LampFilter" 'Puts all lamp intensityscale output (no callbacks) through this function before updating

  'Adjust fading speeds (1 / full MS fading time)
  dim x : for x = 0 to 150 : Lampz.FadeSpeedUp(x) = 1/40 : Lampz.FadeSpeedDown(x) = 1/120 : next
  Lampz.FadeSpeedUp(100) = 1/150: Lampz.FadeSpeedDown(100) = 1/400 'GI

  'GI is set to Lampz id 100 in this table
  Lampz.MassAssign(100) = ColToArray(GI)
  Lampz.State(100) = 0

  'Control lights placed into the lightStateCol collection must be in the correct order per the Lampz MassAssign index
  ' Note: control lights must be first object in MassAssign array, so do this before running the LampzHelper sub
  Dim L, idx
    idx=1 'starting at 1 to save 0 for GI
    for each L in lightStateCol
    Lampz.MassAssign(idx) = L
    idx=idx+1
    next

  'Assign all lightmaps after the control light assignments
  LampzHelper

  'Turn off all lamps on startup
  Lampz.Init  'This just turns state of any lamps to 1

  'Immediate update to turn on GI, turn off lamps
  Lampz.Update

End Sub


Sub SetLampMod(id, val)
  'Debug.print "> " & id & " => " & val
  Lampz.state(id) = val
End Sub

'Helper functions

Function ColtoArray(aDict)  'converts a collection to an indexed array. Indexes will come out random probably.
  redim a(999)
  dim count : count = 0
  dim x  : for each x in aDict : set a(Count) = x : count = count + 1 : Next
  redim preserve a(count-1) : ColtoArray = a
End Function

Sub DisableLighting(pri, DLintensity, ByVal aLvl) 'cp's script  DLintensity = disabled lighting intesity
  if Lampz.UseFunction then aLvl = Lampz.FilterOut(aLvl)  'Callbacks don't get this filter automatically
  pri.blenddisablelighting = aLvl * DLintensity
End Sub

'Sub UpdateLightMap(lightmap, intensity, ByVal aLvl)
'   if Lampz.UseFunction then aLvl = Lampz.FilterOut(aLvl)  'Callbacks don't get this filter automatically
'   lightmap.Opacity = aLvl * intensity
'End Sub


'====================
'Class jungle nf
'====================

'No-op object instead of adding more conditionals to the main loop
'It also prevents errors if empty lamp numbers are called, and it's only one object
'should be g2g?

Class NullFadingObject : Public Property Let IntensityScale(input) : : End Property : End Class

'version 0.11 - Mass Assign, Changed modulate style
'version 0.12 - Update2 (single -1 timer update) update method for core.vbs
'Version 0.12a - Filter can now be accessed via 'FilterOut'
'Version 0.12b - Changed MassAssign from a sub to an indexed property (new syntax: lampfader.MassAssign(15) = Light1 )
'Version 0.13 - No longer requires setlocale. Callback() can be assigned multiple times per index
' Note: if using multiple 'LampFader' objects, set the 'name' variable to avoid conflicts with callbacks
'Version 0.14 - Updated to support modulated signals - Niwak
'Version 0.15 - Added IsLight property - apophis

Class LampFader
  Public IsLight(150)
  Public FadeSpeedDown(150), FadeSpeedUp(150)
  Private Lock(150), Loaded(150), OnOff(150)
  Public UseFunction
  Private cFilter
  Public UseCallback(150), cCallback(150)
  Public Lvl(150), Obj(150)
  Private Mult(150)
  Public FrameTime
  Private InitFrame
  Public Name

  Sub Class_Initialize()
    InitFrame = 0
    dim x : for x = 0 to uBound(OnOff)  'Set up fade speeds
      FadeSpeedDown(x) = 1/100  'fade speed down
      FadeSpeedUp(x) = 1/80   'Fade speed up
      UseFunction = False
      lvl(x) = 0
      OnOff(x) = 0
      Lock(x) = True : Loaded(x) = False
      Mult(x) = 1
      IsLight(x) = False
    Next
    Name = "LampFaderNF" 'NEEDS TO BE CHANGED IF THERE'S MULTIPLE OF THESE OBJECTS, OTHERWISE CALLBACKS WILL INTERFERE WITH EACH OTHER!!
    for x = 0 to uBound(OnOff)    'clear out empty obj
      if IsEmpty(obj(x) ) then Set Obj(x) = NullFader' : Loaded(x) = True
    Next
  End Sub

  Public Property Get Locked(idx) : Locked = Lock(idx) : End Property   ''debug.print Lampz.Locked(100) 'debug
  Public Property Get state(idx) : state = OnOff(idx) : end Property
  Public Property Let Filter(String) : Set cFilter = GetRef(String) : UseFunction = True : End Property
  Public Function FilterOut(aInput) : if UseFunction Then FilterOut = cFilter(aInput) Else FilterOut = aInput End If : End Function
  'Public Property Let Callback(idx, String) : cCallback(idx) = String : UseCallBack(idx) = True : End Property
  Public Property Let Callback(idx, String)
    UseCallBack(idx) = True
    'cCallback(idx) = String 'old execute method
    'New method: build wrapper subs using ExecuteGlobal, then call them
    cCallback(idx) = cCallback(idx) & "___" & String  'multiple strings dilineated by 3x _

    dim tmp : tmp = Split(cCallback(idx), "___")

    dim str, x : for x = 0 to uBound(tmp) 'build proc contents
      'If Not tmp(x)="" then str = str & "  " & tmp(x) & " aLVL" & "  '" & x & vbnewline  'more verbose
      If Not tmp(x)="" then str = str & tmp(x) & " aLVL:"
    Next
    'msgbox "Sub " & name & idx & "(aLvl):" & str & "End Sub"
    dim out : out = "Sub " & name & idx & "(aLvl):" & str & "End Sub"
    ExecuteGlobal Out

  End Property

  Public Property Let state(ByVal idx, input) 'Major update path
    if TypeName(input) <> "Double" and typename(input) <> "Integer"  and typename(input) <> "Long" then
      If input Then
        input = 1
      Else
        input = 0
      End If
    End If
    if Input <> OnOff(idx) then  'discard redundant updates
      OnOff(idx) = input
      Lock(idx) = False
      Loaded(idx) = False
    End If
  End Property

  'Mass assign, Builds arrays where necessary
  'Sub MassAssign(aIdx, aInput)
  Public Property Let MassAssign(aIdx, aInput)
    If typename(obj(aIdx)) = "NullFadingObject" Then 'if empty, use Set
      if IsArray(aInput) then
        obj(aIdx) = aInput
      Else
        Set obj(aIdx) = aInput
        if typename(aInput) = "Light" then IsLight(aIdx) = True
      end if
    Else
      Obj(aIdx) = AppendArray(obj(aIdx), aInput)
    end if
  end Property

  Sub SetLamp(aIdx, aOn) : state(aIdx) = aOn : End Sub  'Solenoid Handler

  Public Sub TurnOnStates() 'If obj contains any light objects, set their states to 1 (Fading is our job!)
    dim debugstr
    dim idx : for idx = 0 to uBound(obj)
      if IsArray(obj(idx)) then
        'debugstr = debugstr & "array found at " & idx & "..."
        dim x, tmp : tmp = obj(idx) 'set tmp to array in order to access it
        for x = 0 to uBound(tmp)
          if typename(tmp(x)) = "Light" then DisableState tmp(x)' : debugstr = debugstr & tmp(x).name & " state'd" & vbnewline
          tmp(x).intensityscale = 0.001 ' this can prevent init stuttering
        Next
      Else
        if typename(obj(idx)) = "Light" then DisableState obj(idx)' : debugstr = debugstr & obj(idx).name & " state'd (not array)" & vbnewline
        obj(idx).intensityscale = 0.001 ' this can prevent init stuttering
      end if
    Next
    ''debug.print debugstr
  End Sub
  Private Sub DisableState(ByRef aObj) : aObj.FadeSpeedUp = 1000 : aObj.State = 1 : End Sub 'turn state to 1

  Public Sub Init() 'Just runs TurnOnStates right now
    TurnOnStates
  End Sub

  Public Property Let Modulate(aIdx, aCoef) : Mult(aIdx) = aCoef : Lock(aIdx) = False : Loaded(aIdx) = False: End Property
  Public Property Get Modulate(aIdx) : Modulate = Mult(aIdx) : End Property

  Public Sub Update1()   'Handle all boolean numeric fading. If done fading, Lock(x) = True. Update on a '1' interval Timer!
    dim x : for x = 0 to uBound(OnOff)
      if not Lock(x) then 'and not Loaded(x) then
        if OnOff(x) > 0 then 'Fade Up
          Lvl(x) = Lvl(x) + FadeSpeedUp(x)
          if Lvl(x) >= OnOff(x) then Lvl(x) = OnOff(x) : Lock(x) = True
        else 'fade down
          Lvl(x) = Lvl(x) - FadeSpeedDown(x)
          if Lvl(x) <= 0 then Lvl(x) = 0 : Lock(x) = True
        end if
      end if
    Next
  End Sub

  Public Sub Update2()   'Both updates on -1 timer (Lowest latency, but less accurate fading at 60fps vsync)
    FrameTime = gametime - InitFrame : InitFrame = GameTime 'Calculate frametime
    dim x : for x = 0 to uBound(OnOff)
      if not Lock(x) then 'and not Loaded(x) then
        if OnOff(x) > 0 then 'Fade Up
          Lvl(x) = Lvl(x) + FadeSpeedUp(x) * FrameTime
          if Lvl(x) >= OnOff(x) then Lvl(x) = OnOff(x) : Lock(x) = True
        else 'fade down
          Lvl(x) = Lvl(x) - FadeSpeedDown(x) * FrameTime
          if Lvl(x) <= 0 then Lvl(x) = 0 : Lock(x) = True
        end if
      end if
    Next
    Update
  End Sub

  Public Sub Update() 'Handle object updates. Update on a -1 Timer! If done fading, loaded(x) = True
    dim x,xx, aLvl : for x = 0 to uBound(OnOff)
      if not Loaded(x) then
        aLvl = Lvl(x)*Mult(x)
        if IsArray(obj(x) ) Then  'if array
          If UseFunction then
            for each xx in obj(x) : xx.IntensityScale = cFilter(aLvl) : Next
          Else
            for each xx in obj(x) : xx.IntensityScale = aLvl : Next
          End If
        else            'if single lamp or flasher
          If UseFunction then
            obj(x).Intensityscale = cFilter(aLvl)
          Else
            obj(x).Intensityscale = aLvl
          End If
        end if
        'if TypeName(lvl(x)) <> "Double" and typename(lvl(x)) <> "Integer" and typename(lvl(x)) <> "Long" then msgbox "uhh " & 2 & " = " & lvl(x)
        'If UseCallBack(x) then execute cCallback(x) & " " & (Lvl(x)) 'Callback
        If UseCallBack(x) then Proc name & x,aLvl 'Proc
        If Lock(x) Then
          if Lvl(x) = OnOff(x) or Lvl(x) = 0 then Loaded(x) = True  'finished fading
        end if
      end if
    Next
  End Sub
End Class

'Lamp Filter
Function LampFilter(aLvl)
  LampFilter = aLvl^1.6 'exponential curve?
End Function

'Helper functions
Sub Proc(string, Callback)  'proc using a string and one argument
  'On Error Resume Next
  dim p : Set P = GetRef(String)
  P Callback
  If err.number = 13 then  msgbox "Proc error! No such procedure: " & vbnewline & string
  if err.number = 424 then msgbox "Proc error! No such Object"
End Sub

Function AppendArray(ByVal aArray, aInput)  'append one value, object, or Array onto the end of a 1 dimensional array
  if IsArray(aInput) then 'Input is an array...
    dim tmp : tmp = aArray
    If not IsArray(aArray) Then 'if not array, create an array
      tmp = aInput
    Else            'Append existing array with aInput array
      Redim Preserve tmp(uBound(aArray) + uBound(aInput)+1) 'If existing array, increase bounds by uBound of incoming array
      dim x : for x = 0 to uBound(aInput)
        if isObject(aInput(x)) then
          Set tmp(x+uBound(aArray)+1 ) = aInput(x)
        Else
          tmp(x+uBound(aArray)+1 ) = aInput(x)
        End If
      Next
      AppendArray = tmp  'return new array
    End If
  Else 'Input is NOT an array...
    If not IsArray(aArray) Then 'if not array, create an array
      aArray = Array(aArray, aInput)
    Else
      Redim Preserve aArray(uBound(aArray)+1) 'If array, increase bounds by 1
      if isObject(aInput) then
        Set aArray(uBound(aArray)) = aInput
      Else
        aArray(uBound(aArray)) = aInput
      End If
    End If
    AppendArray = aArray 'return new array
  End If
End Function


'******************************************************
'****  END LAMPZ
'******************************************************



' ===============================================================
' ZVLM       Virtual Pinball X Light Mapper generated code
'
' This file provide default implementation and template to add bakemap
' & lightmap synchronization for position and lighting.
'
' Lines ending with a comment starting with "' VLM." are meant to
' be copy/pasted ONLY ONCE, since the toolkit will take care of
' updating them directly in your table script, each time an
' export is made.


' ===============================================================
' The following code NEEDS to be copy/pasted to hide the elements
' that are placed to avoid stutters when VPX loads all the nestmaps
' to the GPU. It also NEEDS the following line to be added to the.
' table init function:

' vpmTimer.AddTimer 1000, "WarmUpDone '"

Sub WarmUpDone
  VLM_Warmup_Nestmap_0.Visible = False
  VLM_Warmup_Nestmap_1.Visible = False
  VLM_Warmup_Nestmap_2.Visible = False
End Sub



' ===============================================================
' The following code can be copy/pasted to have premade array for
' movable objects:
' - _LM suffixed arrays contains the lightmaps
' - _BM suffixed arrays contains the bakemap
' - _BL suffixed arrays contains both the bakemap & the lightmaps
Dim CenterTarget_LM: CenterTarget_LM=Array(CenterTarget_LM_Gi_L100) ' VLM.Array;LM;CenterTarget
Dim CenterTarget_BM: CenterTarget_BM=Array(CenterTarget_BM_World) ' VLM.Array;BM;CenterTarget
Dim CenterTarget_BL: CenterTarget_BL=Array(CenterTarget_BM_World, CenterTarget_LM_Gi_L100) ' VLM.Array;BL;CenterTarget
Dim Cherry_Switch_000_LM: Cherry_Switch_000_LM=Array(Cherry_Switch_000_LM_inserts_L2) ' VLM.Array;LM;Cherry_Switch_000
Dim Cherry_Switch_000_BM: Cherry_Switch_000_BM=Array(Cherry_Switch_000_BM_World) ' VLM.Array;BM;Cherry_Switch_000
Dim Cherry_Switch_000_BL: Cherry_Switch_000_BL=Array(Cherry_Switch_000_BM_World, Cherry_Switch_000_LM_inserts_L2) ' VLM.Array;BL;Cherry_Switch_000
Dim Cherry_Switch_001_LM: Cherry_Switch_001_LM=Array(Cherry_Switch_001_LM_Gi_L100, Cherry_Switch_001_LM_inserts_L0) ' VLM.Array;LM;Cherry_Switch_001
Dim Cherry_Switch_001_BM: Cherry_Switch_001_BM=Array(Cherry_Switch_001_BM_World) ' VLM.Array;BM;Cherry_Switch_001
Dim Cherry_Switch_001_BL: Cherry_Switch_001_BL=Array(Cherry_Switch_001_BM_World, Cherry_Switch_001_LM_Gi_L100, Cherry_Switch_001_LM_inserts_L0) ' VLM.Array;BL;Cherry_Switch_001
Dim Cherry_Switch_002_LM: Cherry_Switch_002_LM=Array(Cherry_Switch_002_LM_Gi_L100, Cherry_Switch_002_LM_inserts_L0) ' VLM.Array;LM;Cherry_Switch_002
Dim Cherry_Switch_002_BM: Cherry_Switch_002_BM=Array(Cherry_Switch_002_BM_World) ' VLM.Array;BM;Cherry_Switch_002
Dim Cherry_Switch_002_BL: Cherry_Switch_002_BL=Array(Cherry_Switch_002_BM_World, Cherry_Switch_002_LM_Gi_L100, Cherry_Switch_002_LM_inserts_L0) ' VLM.Array;BL;Cherry_Switch_002
Dim Cherry_Switch_003_LM: Cherry_Switch_003_LM=Array(Cherry_Switch_003_LM_Gi_L100, Cherry_Switch_003_LM_inserts_L3) ' VLM.Array;LM;Cherry_Switch_003
Dim Cherry_Switch_003_BM: Cherry_Switch_003_BM=Array(Cherry_Switch_003_BM_World) ' VLM.Array;BM;Cherry_Switch_003
Dim Cherry_Switch_003_BL: Cherry_Switch_003_BL=Array(Cherry_Switch_003_BM_World, Cherry_Switch_003_LM_Gi_L100, Cherry_Switch_003_LM_inserts_L3) ' VLM.Array;BL;Cherry_Switch_003
Dim Cherry_Switch_004_LM: Cherry_Switch_004_LM=Array(Cherry_Switch_004_LM_Gi_L100, Cherry_Switch_004_LM_inserts_L3) ' VLM.Array;LM;Cherry_Switch_004
Dim Cherry_Switch_004_BM: Cherry_Switch_004_BM=Array(Cherry_Switch_004_BM_World) ' VLM.Array;BM;Cherry_Switch_004
Dim Cherry_Switch_004_BL: Cherry_Switch_004_BL=Array(Cherry_Switch_004_BM_World, Cherry_Switch_004_LM_Gi_L100, Cherry_Switch_004_LM_inserts_L3) ' VLM.Array;BL;Cherry_Switch_004
Dim Cherry_Switch_005_LM: Cherry_Switch_005_LM=Array(Cherry_Switch_005_LM_inserts_L3) ' VLM.Array;LM;Cherry_Switch_005
Dim Cherry_Switch_005_BM: Cherry_Switch_005_BM=Array(Cherry_Switch_005_BM_World) ' VLM.Array;BM;Cherry_Switch_005
Dim Cherry_Switch_005_BL: Cherry_Switch_005_BL=Array(Cherry_Switch_005_BM_World, Cherry_Switch_005_LM_inserts_L3) ' VLM.Array;BL;Cherry_Switch_005
Dim Cherry_Switch_006_LM: Cherry_Switch_006_LM=Array(Cherry_Switch_006_LM_inserts_L3) ' VLM.Array;LM;Cherry_Switch_006
Dim Cherry_Switch_006_BM: Cherry_Switch_006_BM=Array(Cherry_Switch_006_BM_World) ' VLM.Array;BM;Cherry_Switch_006
Dim Cherry_Switch_006_BL: Cherry_Switch_006_BL=Array(Cherry_Switch_006_BM_World, Cherry_Switch_006_LM_inserts_L3) ' VLM.Array;BL;Cherry_Switch_006
Dim Cherry_Switch_007_LM: Cherry_Switch_007_LM=Array(Cherry_Switch_007_LM_inserts_L3, Cherry_Switch_007_LM_inserts_L3) ' VLM.Array;LM;Cherry_Switch_007
Dim Cherry_Switch_007_BM: Cherry_Switch_007_BM=Array(Cherry_Switch_007_BM_World) ' VLM.Array;BM;Cherry_Switch_007
Dim Cherry_Switch_007_BL: Cherry_Switch_007_BL=Array(Cherry_Switch_007_BM_World, Cherry_Switch_007_LM_inserts_L3, Cherry_Switch_007_LM_inserts_L3) ' VLM.Array;BL;Cherry_Switch_007
Dim Cherry_Switch_008_LM: Cherry_Switch_008_LM=Array(Cherry_Switch_008_LM_inserts_L3) ' VLM.Array;LM;Cherry_Switch_008
Dim Cherry_Switch_008_BM: Cherry_Switch_008_BM=Array(Cherry_Switch_008_BM_World) ' VLM.Array;BM;Cherry_Switch_008
Dim Cherry_Switch_008_BL: Cherry_Switch_008_BL=Array(Cherry_Switch_008_BM_World, Cherry_Switch_008_LM_inserts_L3) ' VLM.Array;BL;Cherry_Switch_008
Dim Cherry_Switch_009_LM: Cherry_Switch_009_LM=Array(Cherry_Switch_009_LM_Gi_L100, Cherry_Switch_009_LM_inserts_L0) ' VLM.Array;LM;Cherry_Switch_009
Dim Cherry_Switch_009_BM: Cherry_Switch_009_BM=Array(Cherry_Switch_009_BM_World) ' VLM.Array;BM;Cherry_Switch_009
Dim Cherry_Switch_009_BL: Cherry_Switch_009_BL=Array(Cherry_Switch_009_BM_World, Cherry_Switch_009_LM_Gi_L100, Cherry_Switch_009_LM_inserts_L0) ' VLM.Array;BL;Cherry_Switch_009
Dim Cherry_Switch_010_LM: Cherry_Switch_010_LM=Array(Cherry_Switch_010_LM_Gi_L100, Cherry_Switch_010_LM_inserts_L0) ' VLM.Array;LM;Cherry_Switch_010
Dim Cherry_Switch_010_BM: Cherry_Switch_010_BM=Array(Cherry_Switch_010_BM_World) ' VLM.Array;BM;Cherry_Switch_010
Dim Cherry_Switch_010_BL: Cherry_Switch_010_BL=Array(Cherry_Switch_010_BM_World, Cherry_Switch_010_LM_Gi_L100, Cherry_Switch_010_LM_inserts_L0) ' VLM.Array;BL;Cherry_Switch_010
Dim Cherry_Switch_011_LM: Cherry_Switch_011_LM=Array(Cherry_Switch_011_LM_Gi_L100, Cherry_Switch_011_LM_inserts_L0) ' VLM.Array;LM;Cherry_Switch_011
Dim Cherry_Switch_011_BM: Cherry_Switch_011_BM=Array(Cherry_Switch_011_BM_World) ' VLM.Array;BM;Cherry_Switch_011
Dim Cherry_Switch_011_BL: Cherry_Switch_011_BL=Array(Cherry_Switch_011_BM_World, Cherry_Switch_011_LM_Gi_L100, Cherry_Switch_011_LM_inserts_L0) ' VLM.Array;BL;Cherry_Switch_011
Dim Circle_015_LM: Circle_015_LM=Array(Circle_015_LM_Gi_L100, Circle_015_LM_inserts_L06, Circle_015_LM_inserts_L35) ' VLM.Array;LM;Circle_015
Dim Circle_015_BM: Circle_015_BM=Array(Circle_015_BM_World) ' VLM.Array;BM;Circle_015
Dim Circle_015_BL: Circle_015_BL=Array(Circle_015_BM_World, Circle_015_LM_Gi_L100, Circle_015_LM_inserts_L06, Circle_015_LM_inserts_L35) ' VLM.Array;BL;Circle_015
Dim Circle_016_LM: Circle_016_LM=Array(Circle_016_LM_Gi_L100, Circle_016_LM_inserts_L34) ' VLM.Array;LM;Circle_016
Dim Circle_016_BM: Circle_016_BM=Array(Circle_016_BM_World) ' VLM.Array;BM;Circle_016
Dim Circle_016_BL: Circle_016_BL=Array(Circle_016_BM_World, Circle_016_LM_Gi_L100, Circle_016_LM_inserts_L34) ' VLM.Array;BL;Circle_016
Dim Circle_019_LM: Circle_019_LM=Array(Circle_019_LM_Gi_L100, Circle_019_LM_inserts_L31, Circle_019_LM_inserts_L32, Circle_019_LM_inserts_L35) ' VLM.Array;LM;Circle_019
Dim Circle_019_BM: Circle_019_BM=Array(Circle_019_BM_World) ' VLM.Array;BM;Circle_019
Dim Circle_019_BL: Circle_019_BL=Array(Circle_019_BM_World, Circle_019_LM_Gi_L100, Circle_019_LM_inserts_L31, Circle_019_LM_inserts_L32, Circle_019_LM_inserts_L35) ' VLM.Array;BL;Circle_019
Dim Circle_037_LM: Circle_037_LM=Array(Circle_037_LM_Gi_L100, Circle_037_LM_inserts_L07, Circle_037_LM_inserts_L35) ' VLM.Array;LM;Circle_037
Dim Circle_037_BM: Circle_037_BM=Array(Circle_037_BM_World) ' VLM.Array;BM;Circle_037
Dim Circle_037_BL: Circle_037_BL=Array(Circle_037_BM_World, Circle_037_LM_Gi_L100, Circle_037_LM_inserts_L07, Circle_037_LM_inserts_L35) ' VLM.Array;BL;Circle_037
Dim LFlip_001_LM: LFlip_001_LM=Array(LFlip_001_LM_Gi_L100, LFlip_001_LM_inserts_L30, LFlip_001_LM_inserts_L31, LFlip_001_LM_inserts_L34, LFlip_001_LM_inserts_L35) ' VLM.Array;LM;LFlip_001
Dim LFlip_001_BM: LFlip_001_BM=Array(LFlip_001_BM_World) ' VLM.Array;BM;LFlip_001
Dim LFlip_001_BL: LFlip_001_BL=Array(LFlip_001_BM_World, LFlip_001_LM_Gi_L100, LFlip_001_LM_inserts_L30, LFlip_001_LM_inserts_L31, LFlip_001_LM_inserts_L34, LFlip_001_LM_inserts_L35) ' VLM.Array;BL;LFlip_001
Dim LFlipR_LM: LFlipR_LM=Array(LFlipR_LM_Gi_L100, LFlipR_LM_inserts_L30, LFlipR_LM_inserts_L31, LFlipR_LM_inserts_L35) ' VLM.Array;LM;LFlipR
Dim LFlipR_BM: LFlipR_BM=Array(LFlipR_BM_World) ' VLM.Array;BM;LFlipR
Dim LFlipR_BL: LFlipR_BL=Array(LFlipR_BM_World, LFlipR_LM_Gi_L100, LFlipR_LM_inserts_L30, LFlipR_LM_inserts_L31, LFlipR_LM_inserts_L35) ' VLM.Array;BL;LFlipR
Dim MidTargetLeft_LM: MidTargetLeft_LM=Array(MidTargetLeft_LM_Gi_L100, MidTargetLeft_LM_inserts_L08) ' VLM.Array;LM;MidTargetLeft
Dim MidTargetLeft_BM: MidTargetLeft_BM=Array(MidTargetLeft_BM_World) ' VLM.Array;BM;MidTargetLeft
Dim MidTargetLeft_BL: MidTargetLeft_BL=Array(MidTargetLeft_BM_World, MidTargetLeft_LM_Gi_L100, MidTargetLeft_LM_inserts_L08) ' VLM.Array;BL;MidTargetLeft
Dim MidTargetRight_LM: MidTargetRight_LM=Array(MidTargetRight_LM_Gi_L100, MidTargetRight_LM_inserts_L09) ' VLM.Array;LM;MidTargetRight
Dim MidTargetRight_BM: MidTargetRight_BM=Array(MidTargetRight_BM_World) ' VLM.Array;BM;MidTargetRight
Dim MidTargetRight_BL: MidTargetRight_BL=Array(MidTargetRight_BM_World, MidTargetRight_LM_Gi_L100, MidTargetRight_LM_inserts_L09) ' VLM.Array;BL;MidTargetRight
Dim Pop_Bumper_WPC_002_LM: Pop_Bumper_WPC_002_LM=Array(Pop_Bumper_WPC_002_LM_Gi_L100, Pop_Bumper_WPC_002_LM_inserts_L) ' VLM.Array;LM;Pop_Bumper_WPC_002
Dim Pop_Bumper_WPC_002_BM: Pop_Bumper_WPC_002_BM=Array(Pop_Bumper_WPC_002_BM_World) ' VLM.Array;BM;Pop_Bumper_WPC_002
Dim Pop_Bumper_WPC_002_BL: Pop_Bumper_WPC_002_BL=Array(Pop_Bumper_WPC_002_BM_World, Pop_Bumper_WPC_002_LM_Gi_L100, Pop_Bumper_WPC_002_LM_inserts_L) ' VLM.Array;BL;Pop_Bumper_WPC_002
Dim Pop_Bumper_WPC_003_LM: Pop_Bumper_WPC_003_LM=Array(Pop_Bumper_WPC_003_LM_Gi_L100, Pop_Bumper_WPC_003_LM_inserts_L) ' VLM.Array;LM;Pop_Bumper_WPC_003
Dim Pop_Bumper_WPC_003_BM: Pop_Bumper_WPC_003_BM=Array(Pop_Bumper_WPC_003_BM_World) ' VLM.Array;BM;Pop_Bumper_WPC_003
Dim Pop_Bumper_WPC_003_BL: Pop_Bumper_WPC_003_BL=Array(Pop_Bumper_WPC_003_BM_World, Pop_Bumper_WPC_003_LM_Gi_L100, Pop_Bumper_WPC_003_LM_inserts_L) ' VLM.Array;BL;Pop_Bumper_WPC_003
Dim Pop_Bumper_WPC_005_LM: Pop_Bumper_WPC_005_LM=Array(Pop_Bumper_WPC_005_LM_Gi_L100, Pop_Bumper_WPC_005_LM_inserts_L) ' VLM.Array;LM;Pop_Bumper_WPC_005
Dim Pop_Bumper_WPC_005_BM: Pop_Bumper_WPC_005_BM=Array(Pop_Bumper_WPC_005_BM_World) ' VLM.Array;BM;Pop_Bumper_WPC_005
Dim Pop_Bumper_WPC_005_BL: Pop_Bumper_WPC_005_BL=Array(Pop_Bumper_WPC_005_BM_World, Pop_Bumper_WPC_005_LM_Gi_L100, Pop_Bumper_WPC_005_LM_inserts_L) ' VLM.Array;BL;Pop_Bumper_WPC_005
Dim RFlip_LM: RFlip_LM=Array(RFlip_LM_Gi_L100, RFlip_LM_inserts_L32, RFlip_LM_inserts_L33, RFlip_LM_inserts_L34, RFlip_LM_inserts_L35) ' VLM.Array;LM;RFlip
Dim RFlip_BM: RFlip_BM=Array(RFlip_BM_World) ' VLM.Array;BM;RFlip
Dim RFlip_BL: RFlip_BL=Array(RFlip_BM_World, RFlip_LM_Gi_L100, RFlip_LM_inserts_L32, RFlip_LM_inserts_L33, RFlip_LM_inserts_L34, RFlip_LM_inserts_L35) ' VLM.Array;BL;RFlip
Dim RFlipR_LM: RFlipR_LM=Array(RFlipR_LM_Gi_L100, RFlipR_LM_inserts_L32, RFlipR_LM_inserts_L33, RFlipR_LM_inserts_L35) ' VLM.Array;LM;RFlipR
Dim RFlipR_BM: RFlipR_BM=Array(RFlipR_BM_World) ' VLM.Array;BM;RFlipR
Dim RFlipR_BL: RFlipR_BL=Array(RFlipR_BM_World, RFlipR_LM_Gi_L100, RFlipR_LM_inserts_L32, RFlipR_LM_inserts_L33, RFlipR_LM_inserts_L35) ' VLM.Array;BL;RFlipR
Dim Sphere_007_LM: Sphere_007_LM=Array(Sphere_007_LM_Gi_L100, Sphere_007_LM_inserts_L02, Sphere_007_LM_inserts_L03, Sphere_007_LM_inserts_L04) ' VLM.Array;LM;Sphere_007
Dim Sphere_007_BM: Sphere_007_BM=Array(Sphere_007_BM_World) ' VLM.Array;BM;Sphere_007
Dim Sphere_007_BL: Sphere_007_BL=Array(Sphere_007_BM_World, Sphere_007_LM_Gi_L100, Sphere_007_LM_inserts_L02, Sphere_007_LM_inserts_L03, Sphere_007_LM_inserts_L04) ' VLM.Array;BL;Sphere_007
Dim TopTargetLeft_001_LM: TopTargetLeft_001_LM=Array(TopTargetLeft_001_LM_Gi_L100, TopTargetLeft_001_LM_inserts_L0) ' VLM.Array;LM;TopTargetLeft_001
Dim TopTargetLeft_001_BM: TopTargetLeft_001_BM=Array(TopTargetLeft_001_BM_World) ' VLM.Array;BM;TopTargetLeft_001
Dim TopTargetLeft_001_BL: TopTargetLeft_001_BL=Array(TopTargetLeft_001_BM_World, TopTargetLeft_001_LM_Gi_L100, TopTargetLeft_001_LM_inserts_L0) ' VLM.Array;BL;TopTargetLeft_001
Dim TopTargetRight_001_LM: TopTargetRight_001_LM=Array(TopTargetRight_001_LM_Gi_L100, TopTargetRight_001_LM_inserts_L) ' VLM.Array;LM;TopTargetRight_001
Dim TopTargetRight_001_BM: TopTargetRight_001_BM=Array(TopTargetRight_001_BM_World) ' VLM.Array;BM;TopTargetRight_001
Dim TopTargetRight_001_BL: TopTargetRight_001_BL=Array(TopTargetRight_001_BM_World, TopTargetRight_001_LM_Gi_L100, TopTargetRight_001_LM_inserts_L) ' VLM.Array;BL;TopTargetRight_001
Dim lockdown_LM: lockdown_LM=Array() ' VLM.Array;LM;lockdown
Dim lockdown_BM: lockdown_BM=Array(lockdown_BM_World) ' VLM.Array;BM;lockdown
Dim lockdown_BL: lockdown_BL=Array(lockdown_BM_World) ' VLM.Array;BL;lockdown
Dim lsling_LM: lsling_LM=Array(lsling_LM_Gi_L100, lsling_LM_inserts_L20, lsling_LM_inserts_L21, lsling_LM_inserts_L30, lsling_LM_inserts_L31) ' VLM.Array;LM;lsling
Dim lsling_BM: lsling_BM=Array(lsling_BM_World) ' VLM.Array;BM;lsling
Dim lsling_BL: lsling_BL=Array(lsling_BM_World, lsling_LM_Gi_L100, lsling_LM_inserts_L20, lsling_LM_inserts_L21, lsling_LM_inserts_L30, lsling_LM_inserts_L31) ' VLM.Array;BL;lsling
Dim lsling001_LM: lsling001_LM=Array(lsling001_LM_Gi_L100, lsling001_LM_inserts_L20, lsling001_LM_inserts_L21, lsling001_LM_inserts_L30, lsling001_LM_inserts_L31) ' VLM.Array;LM;lsling001
Dim lsling001_BM: lsling001_BM=Array(lsling001_BM_World) ' VLM.Array;BM;lsling001
Dim lsling001_BL: lsling001_BL=Array(lsling001_BM_World, lsling001_LM_Gi_L100, lsling001_LM_inserts_L20, lsling001_LM_inserts_L21, lsling001_LM_inserts_L30, lsling001_LM_inserts_L31) ' VLM.Array;BL;lsling001
Dim lsling002_LM: lsling002_LM=Array(lsling002_LM_Gi_L100, lsling002_LM_inserts_L20, lsling002_LM_inserts_L21, lsling002_LM_inserts_L30, lsling002_LM_inserts_L31) ' VLM.Array;LM;lsling002
Dim lsling002_BM: lsling002_BM=Array(lsling002_BM_World) ' VLM.Array;BM;lsling002
Dim lsling002_BL: lsling002_BL=Array(lsling002_BM_World, lsling002_LM_Gi_L100, lsling002_LM_inserts_L20, lsling002_LM_inserts_L21, lsling002_LM_inserts_L30, lsling002_LM_inserts_L31) ' VLM.Array;BL;lsling002
Dim pgate_LM: pgate_LM=Array(pgate_LM_Gi_L100) ' VLM.Array;LM;pgate
Dim pgate_BM: pgate_BM=Array(pgate_BM_World) ' VLM.Array;BM;pgate
Dim pgate_BL: pgate_BL=Array(pgate_BM_World, pgate_LM_Gi_L100) ' VLM.Array;BL;pgate
Dim rsling_LM: rsling_LM=Array(rsling_LM_Gi_L100, rsling_LM_inserts_L28, rsling_LM_inserts_L29, rsling_LM_inserts_L32, rsling_LM_inserts_L33) ' VLM.Array;LM;rsling
Dim rsling_BM: rsling_BM=Array(rsling_BM_World) ' VLM.Array;BM;rsling
Dim rsling_BL: rsling_BL=Array(rsling_BM_World, rsling_LM_Gi_L100, rsling_LM_inserts_L28, rsling_LM_inserts_L29, rsling_LM_inserts_L32, rsling_LM_inserts_L33) ' VLM.Array;BL;rsling
Dim rsling001_LM: rsling001_LM=Array(rsling001_LM_Gi_L100, rsling001_LM_inserts_L28, rsling001_LM_inserts_L29, rsling001_LM_inserts_L32, rsling001_LM_inserts_L33) ' VLM.Array;LM;rsling001
Dim rsling001_BM: rsling001_BM=Array(rsling001_BM_World) ' VLM.Array;BM;rsling001
Dim rsling001_BL: rsling001_BL=Array(rsling001_BM_World, rsling001_LM_Gi_L100, rsling001_LM_inserts_L28, rsling001_LM_inserts_L29, rsling001_LM_inserts_L32, rsling001_LM_inserts_L33) ' VLM.Array;BL;rsling001
Dim rsling002_LM: rsling002_LM=Array(rsling002_LM_Gi_L100, rsling002_LM_inserts_L28, rsling002_LM_inserts_L29, rsling002_LM_inserts_L32, rsling002_LM_inserts_L33) ' VLM.Array;LM;rsling002
Dim rsling002_BM: rsling002_BM=Array(rsling002_BM_World) ' VLM.Array;BM;rsling002
Dim rsling002_BL: rsling002_BL=Array(rsling002_BM_World, rsling002_LM_Gi_L100, rsling002_LM_inserts_L28, rsling002_LM_inserts_L29, rsling002_LM_inserts_L32, rsling002_LM_inserts_L33) ' VLM.Array;BL;rsling002
Dim siderails_LM: siderails_LM=Array(siderails_LM_Gi_L100, siderails_LM_inserts_L30, siderails_LM_inserts_L31) ' VLM.Array;LM;siderails
Dim siderails_BM: siderails_BM=Array(siderails_BM_World) ' VLM.Array;BM;siderails
Dim siderails_BL: siderails_BL=Array(siderails_BM_World, siderails_LM_Gi_L100, siderails_LM_inserts_L30, siderails_LM_inserts_L31) ' VLM.Array;BL;siderails



Dim BM_World: BM_World = Array(playfield_mesh1,Parts_BM_World,Overlay_BM_World,CenterTarget_BM_World,Cherry_Switch_000_BM_World,Cherry_Switch_001_BM_World,Cherry_Switch_002_BM_World,Cherry_Switch_003_BM_World,Cherry_Switch_004_BM_World,Cherry_Switch_005_BM_World,Cherry_Switch_006_BM_World,Cherry_Switch_007_BM_World,Cherry_Switch_008_BM_World,Cherry_Switch_009_BM_World,Cherry_Switch_010_BM_World,Cherry_Switch_011_BM_World,Circle_015_BM_World,Circle_016_BM_World,Circle_019_BM_World,Circle_037_BM_World,LFlip_001_BM_World,LFlipR_BM_World,MidTargetLeft_BM_World,MidTargetRight_BM_World,Pop_Bumper_WPC_002_BM_World,Pop_Bumper_WPC_003_BM_World,Pop_Bumper_WPC_005_BM_World,RFlip_BM_World,RFlipR_BM_World,Sphere_007_BM_World,TopTargetLeft_001_BM_World,TopTargetRight_001_BM_World,lockdown_BM_World,lsling_BM_World,lsling001_BM_World,lsling002_BM_World,pgate_BM_World,rsling_BM_World,rsling001_BM_World,rsling002_BM_World,siderails_BM_World)




' ===============================================================
' The following code can be copy/pasted if using Lampz fading system
' It links each Lampz lamp/flasher to the corresponding light and lightmap

Sub UpdateLightMap(lightmap, intensity, ByVal aLvl)
   if Lampz.UseFunction then aLvl = Lampz.FilterOut(aLvl) 'Callbacks don't get this filter automatically
   lightmap.Opacity = aLvl * intensity
End Sub

Sub LampzHelper
  ' Lampz.MassAssign(100) = L100 ' VLM.Lampz;Gi-L100
  Lampz.Callback(100) = "UpdateLightMap Playfield_LM_Gi_L100, 100.0, " ' VLM.Lampz;Gi-L100
  Lampz.Callback(100) = "UpdateLightMap Overlay_LM_Gi_L100, 100.0, " ' VLM.Lampz;Gi-L100
  Lampz.Callback(100) = "UpdateLightMap CenterTarget_LM_Gi_L100, 100.0, " ' VLM.Lampz;Gi-L100
  Lampz.Callback(100) = "UpdateLightMap Cherry_Switch_001_LM_Gi_L100, 100.0, " ' VLM.Lampz;Gi-L100
  Lampz.Callback(100) = "UpdateLightMap Cherry_Switch_002_LM_Gi_L100, 100.0, " ' VLM.Lampz;Gi-L100
  Lampz.Callback(100) = "UpdateLightMap Cherry_Switch_003_LM_Gi_L100, 100.0, " ' VLM.Lampz;Gi-L100
  Lampz.Callback(100) = "UpdateLightMap Cherry_Switch_004_LM_Gi_L100, 100.0, " ' VLM.Lampz;Gi-L100
  Lampz.Callback(100) = "UpdateLightMap Cherry_Switch_009_LM_Gi_L100, 100.0, " ' VLM.Lampz;Gi-L100
  Lampz.Callback(100) = "UpdateLightMap Cherry_Switch_010_LM_Gi_L100, 100.0, " ' VLM.Lampz;Gi-L100
  Lampz.Callback(100) = "UpdateLightMap Cherry_Switch_011_LM_Gi_L100, 100.0, " ' VLM.Lampz;Gi-L100
  Lampz.Callback(100) = "UpdateLightMap Circle_015_LM_Gi_L100, 100.0, " ' VLM.Lampz;Gi-L100
  Lampz.Callback(100) = "UpdateLightMap Circle_016_LM_Gi_L100, 100.0, " ' VLM.Lampz;Gi-L100
  Lampz.Callback(100) = "UpdateLightMap Circle_019_LM_Gi_L100, 100.0, " ' VLM.Lampz;Gi-L100
  Lampz.Callback(100) = "UpdateLightMap Circle_037_LM_Gi_L100, 100.0, " ' VLM.Lampz;Gi-L100
  Lampz.Callback(100) = "UpdateLightMap LFlip_001_LM_Gi_L100, 100.0, " ' VLM.Lampz;Gi-L100
  Lampz.Callback(100) = "UpdateLightMap LFlipR_LM_Gi_L100, 100.0, " ' VLM.Lampz;Gi-L100
  Lampz.Callback(100) = "UpdateLightMap MidTargetLeft_LM_Gi_L100, 100.0, " ' VLM.Lampz;Gi-L100
  Lampz.Callback(100) = "UpdateLightMap MidTargetRight_LM_Gi_L100, 100.0, " ' VLM.Lampz;Gi-L100
  Lampz.Callback(100) = "UpdateLightMap Pop_Bumper_WPC_002_LM_Gi_L100, 100.0, " ' VLM.Lampz;Gi-L100
  Lampz.Callback(100) = "UpdateLightMap Pop_Bumper_WPC_003_LM_Gi_L100, 100.0, " ' VLM.Lampz;Gi-L100
  Lampz.Callback(100) = "UpdateLightMap Pop_Bumper_WPC_005_LM_Gi_L100, 100.0, " ' VLM.Lampz;Gi-L100
  Lampz.Callback(100) = "UpdateLightMap RFlip_LM_Gi_L100, 100.0, " ' VLM.Lampz;Gi-L100
  Lampz.Callback(100) = "UpdateLightMap RFlipR_LM_Gi_L100, 100.0, " ' VLM.Lampz;Gi-L100
  Lampz.Callback(100) = "UpdateLightMap Sphere_007_LM_Gi_L100, 100.0, " ' VLM.Lampz;Gi-L100
  Lampz.Callback(100) = "UpdateLightMap TopTargetLeft_001_LM_Gi_L100, 100.0, " ' VLM.Lampz;Gi-L100
  Lampz.Callback(100) = "UpdateLightMap TopTargetRight_001_LM_Gi_L100, 100.0, " ' VLM.Lampz;Gi-L100
  Lampz.Callback(100) = "UpdateLightMap lsling_LM_Gi_L100, 100.0, " ' VLM.Lampz;Gi-L100
  Lampz.Callback(100) = "UpdateLightMap lsling001_LM_Gi_L100, 100.0, " ' VLM.Lampz;Gi-L100
  Lampz.Callback(100) = "UpdateLightMap lsling002_LM_Gi_L100, 100.0, " ' VLM.Lampz;Gi-L100
  Lampz.Callback(100) = "UpdateLightMap pgate_LM_Gi_L100, 100.0, " ' VLM.Lampz;Gi-L100
  Lampz.Callback(100) = "UpdateLightMap rsling_LM_Gi_L100, 100.0, " ' VLM.Lampz;Gi-L100
  Lampz.Callback(100) = "UpdateLightMap rsling001_LM_Gi_L100, 100.0, " ' VLM.Lampz;Gi-L100
  Lampz.Callback(100) = "UpdateLightMap rsling002_LM_Gi_L100, 100.0, " ' VLM.Lampz;Gi-L100
  Lampz.Callback(100) = "UpdateLightMap siderails_LM_Gi_L100, 100.0, " ' VLM.Lampz;Gi-L100
  ' Lampz.MassAssign(01) = L01 ' VLM.Lampz;inserts-L01
  Lampz.Callback(01) = "UpdateLightMap Playfield_LM_inserts_L01, 100.0, " ' VLM.Lampz;inserts-L01
  Lampz.Callback(01) = "UpdateLightMap Overlay_LM_inserts_L01, 100.0, " ' VLM.Lampz;inserts-L01
  Lampz.Callback(01) = "UpdateLightMap Cherry_Switch_009_LM_inserts_L0, 100.0, " ' VLM.Lampz;inserts-L01
  ' Lampz.MassAssign(02) = L02 ' VLM.Lampz;inserts-L02
  Lampz.Callback(02) = "UpdateLightMap Playfield_LM_inserts_L02, 100.0, " ' VLM.Lampz;inserts-L02
  Lampz.Callback(02) = "UpdateLightMap Cherry_Switch_001_LM_inserts_L0, 100.0, " ' VLM.Lampz;inserts-L02
  Lampz.Callback(02) = "UpdateLightMap Sphere_007_LM_inserts_L02, 100.0, " ' VLM.Lampz;inserts-L02
  ' Lampz.MassAssign(03) = L03 ' VLM.Lampz;inserts-L03
  Lampz.Callback(03) = "UpdateLightMap Playfield_LM_inserts_L03, 100.0, " ' VLM.Lampz;inserts-L03
  Lampz.Callback(03) = "UpdateLightMap Cherry_Switch_002_LM_inserts_L0, 100.0, " ' VLM.Lampz;inserts-L03
  Lampz.Callback(03) = "UpdateLightMap Sphere_007_LM_inserts_L03, 100.0, " ' VLM.Lampz;inserts-L03
  ' Lampz.MassAssign(04) = L04 ' VLM.Lampz;inserts-L04
  Lampz.Callback(04) = "UpdateLightMap Playfield_LM_inserts_L04, 100.0, " ' VLM.Lampz;inserts-L04
  Lampz.Callback(04) = "UpdateLightMap Cherry_Switch_010_LM_inserts_L0, 100.0, " ' VLM.Lampz;inserts-L04
  Lampz.Callback(04) = "UpdateLightMap Sphere_007_LM_inserts_L04, 100.0, " ' VLM.Lampz;inserts-L04
  ' Lampz.MassAssign(05) = L05 ' VLM.Lampz;inserts-L05
  Lampz.Callback(05) = "UpdateLightMap Playfield_LM_inserts_L05, 100.0, " ' VLM.Lampz;inserts-L05
  Lampz.Callback(05) = "UpdateLightMap Overlay_LM_inserts_L05, 100.0, " ' VLM.Lampz;inserts-L05
  Lampz.Callback(05) = "UpdateLightMap Cherry_Switch_011_LM_inserts_L0, 100.0, " ' VLM.Lampz;inserts-L05
  ' Lampz.MassAssign(06) = L06 ' VLM.Lampz;inserts-L06
  Lampz.Callback(06) = "UpdateLightMap Playfield_LM_inserts_L06, 100.0, " ' VLM.Lampz;inserts-L06
  Lampz.Callback(06) = "UpdateLightMap Overlay_LM_inserts_L06, 100.0, " ' VLM.Lampz;inserts-L06
  Lampz.Callback(06) = "UpdateLightMap Circle_015_LM_inserts_L06, 100.0, " ' VLM.Lampz;inserts-L06
  Lampz.Callback(06) = "UpdateLightMap TopTargetLeft_001_LM_inserts_L0, 100.0, " ' VLM.Lampz;inserts-L06
  ' Lampz.MassAssign(07) = L07 ' VLM.Lampz;inserts-L07
  Lampz.Callback(07) = "UpdateLightMap Playfield_LM_inserts_L07, 100.0, " ' VLM.Lampz;inserts-L07
  Lampz.Callback(07) = "UpdateLightMap Overlay_LM_inserts_L07, 100.0, " ' VLM.Lampz;inserts-L07
  Lampz.Callback(07) = "UpdateLightMap Circle_037_LM_inserts_L07, 100.0, " ' VLM.Lampz;inserts-L07
  Lampz.Callback(07) = "UpdateLightMap TopTargetRight_001_LM_inserts_L, 100.0, " ' VLM.Lampz;inserts-L07
  ' Lampz.MassAssign(08) = L08 ' VLM.Lampz;inserts-L08
  Lampz.Callback(08) = "UpdateLightMap Playfield_LM_inserts_L08, 100.0, " ' VLM.Lampz;inserts-L08
  Lampz.Callback(08) = "UpdateLightMap Overlay_LM_inserts_L08, 100.0, " ' VLM.Lampz;inserts-L08
  Lampz.Callback(08) = "UpdateLightMap MidTargetLeft_LM_inserts_L08, 100.0, " ' VLM.Lampz;inserts-L08
  ' Lampz.MassAssign(09) = L09 ' VLM.Lampz;inserts-L09
  Lampz.Callback(09) = "UpdateLightMap Playfield_LM_inserts_L09, 100.0, " ' VLM.Lampz;inserts-L09
  Lampz.Callback(09) = "UpdateLightMap Overlay_LM_inserts_L09, 100.0, " ' VLM.Lampz;inserts-L09
  Lampz.Callback(09) = "UpdateLightMap MidTargetRight_LM_inserts_L09, 100.0, " ' VLM.Lampz;inserts-L09
  ' Lampz.MassAssign(10) = L10 ' VLM.Lampz;inserts-L10
  Lampz.Callback(10) = "UpdateLightMap Playfield_LM_inserts_L10, 100.0, " ' VLM.Lampz;inserts-L10
  ' Lampz.MassAssign(11) = L11 ' VLM.Lampz;inserts-L11
  Lampz.Callback(11) = "UpdateLightMap Playfield_LM_inserts_L11, 100.0, " ' VLM.Lampz;inserts-L11
  ' Lampz.MassAssign(12) = L12 ' VLM.Lampz;inserts-L12
  Lampz.Callback(12) = "UpdateLightMap Playfield_LM_inserts_L12, 100.0, " ' VLM.Lampz;inserts-L12
  ' Lampz.MassAssign(13) = L13 ' VLM.Lampz;inserts-L13
  Lampz.Callback(13) = "UpdateLightMap Playfield_LM_inserts_L13, 100.0, " ' VLM.Lampz;inserts-L13
  ' Lampz.MassAssign(14) = L14 ' VLM.Lampz;inserts-L14
  Lampz.Callback(14) = "UpdateLightMap Playfield_LM_inserts_L14, 100.0, " ' VLM.Lampz;inserts-L14
  ' Lampz.MassAssign(15) = L15 ' VLM.Lampz;inserts-L15
  Lampz.Callback(15) = "UpdateLightMap Playfield_LM_inserts_L15, 100.0, " ' VLM.Lampz;inserts-L15
  ' Lampz.MassAssign(16) = L16 ' VLM.Lampz;inserts-L16
  Lampz.Callback(16) = "UpdateLightMap Playfield_LM_inserts_L16, 100.0, " ' VLM.Lampz;inserts-L16
  ' Lampz.MassAssign(17) = L17 ' VLM.Lampz;inserts-L17
  Lampz.Callback(17) = "UpdateLightMap Playfield_LM_inserts_L17, 100.0, " ' VLM.Lampz;inserts-L17
  ' Lampz.MassAssign(18) = L18 ' VLM.Lampz;inserts-L18
  Lampz.Callback(18) = "UpdateLightMap Playfield_LM_inserts_L18, 100.0, " ' VLM.Lampz;inserts-L18
  ' Lampz.MassAssign(19) = L19 ' VLM.Lampz;inserts-L19
  Lampz.Callback(19) = "UpdateLightMap Playfield_LM_inserts_L19, 100.0, " ' VLM.Lampz;inserts-L19
  ' Lampz.MassAssign(20) = L20 ' VLM.Lampz;inserts-L20
  Lampz.Callback(20) = "UpdateLightMap Playfield_LM_inserts_L20, 100.0, " ' VLM.Lampz;inserts-L20
  Lampz.Callback(20) = "UpdateLightMap Cherry_Switch_000_LM_inserts_L2, 100.0, " ' VLM.Lampz;inserts-L20
  Lampz.Callback(20) = "UpdateLightMap lsling_LM_inserts_L20, 100.0, " ' VLM.Lampz;inserts-L20
  Lampz.Callback(20) = "UpdateLightMap lsling001_LM_inserts_L20, 100.0, " ' VLM.Lampz;inserts-L20
  Lampz.Callback(20) = "UpdateLightMap lsling002_LM_inserts_L20, 100.0, " ' VLM.Lampz;inserts-L20
  ' Lampz.MassAssign(21) = L21 ' VLM.Lampz;inserts-L21
  Lampz.Callback(21) = "UpdateLightMap Playfield_LM_inserts_L21, 100.0, " ' VLM.Lampz;inserts-L21
  Lampz.Callback(21) = "UpdateLightMap lsling_LM_inserts_L21, 100.0, " ' VLM.Lampz;inserts-L21
  Lampz.Callback(21) = "UpdateLightMap lsling001_LM_inserts_L21, 100.0, " ' VLM.Lampz;inserts-L21
  Lampz.Callback(21) = "UpdateLightMap lsling002_LM_inserts_L21, 100.0, " ' VLM.Lampz;inserts-L21
  ' Lampz.MassAssign(22) = L22 ' VLM.Lampz;inserts-L22
  Lampz.Callback(22) = "UpdateLightMap Playfield_LM_inserts_L22, 100.0, " ' VLM.Lampz;inserts-L22
  ' Lampz.MassAssign(23) = L23 ' VLM.Lampz;inserts-L23
  Lampz.Callback(23) = "UpdateLightMap Playfield_LM_inserts_L23, 100.0, " ' VLM.Lampz;inserts-L23
  ' Lampz.MassAssign(24) = L24 ' VLM.Lampz;inserts-L24
  Lampz.Callback(24) = "UpdateLightMap Playfield_LM_inserts_L24, 100.0, " ' VLM.Lampz;inserts-L24
  ' Lampz.MassAssign(25) = L25 ' VLM.Lampz;inserts-L25
  Lampz.Callback(25) = "UpdateLightMap Playfield_LM_inserts_L25, 100.0, " ' VLM.Lampz;inserts-L25
  ' Lampz.MassAssign(26) = L26 ' VLM.Lampz;inserts-L26
  Lampz.Callback(26) = "UpdateLightMap Playfield_LM_inserts_L26, 100.0, " ' VLM.Lampz;inserts-L26
  ' Lampz.MassAssign(27) = L27 ' VLM.Lampz;inserts-L27
  Lampz.Callback(27) = "UpdateLightMap Playfield_LM_inserts_L27, 100.0, " ' VLM.Lampz;inserts-L27
  ' Lampz.MassAssign(28) = L28 ' VLM.Lampz;inserts-L28
  Lampz.Callback(28) = "UpdateLightMap Playfield_LM_inserts_L28, 100.0, " ' VLM.Lampz;inserts-L28
  Lampz.Callback(28) = "UpdateLightMap rsling_LM_inserts_L28, 100.0, " ' VLM.Lampz;inserts-L28
  Lampz.Callback(28) = "UpdateLightMap rsling001_LM_inserts_L28, 100.0, " ' VLM.Lampz;inserts-L28
  Lampz.Callback(28) = "UpdateLightMap rsling002_LM_inserts_L28, 100.0, " ' VLM.Lampz;inserts-L28
  ' Lampz.MassAssign(29) = L29 ' VLM.Lampz;inserts-L29
  Lampz.Callback(29) = "UpdateLightMap Playfield_LM_inserts_L29, 100.0, " ' VLM.Lampz;inserts-L29
  Lampz.Callback(29) = "UpdateLightMap Overlay_LM_inserts_L29, 100.0, " ' VLM.Lampz;inserts-L29
  Lampz.Callback(29) = "UpdateLightMap rsling_LM_inserts_L29, 100.0, " ' VLM.Lampz;inserts-L29
  Lampz.Callback(29) = "UpdateLightMap rsling001_LM_inserts_L29, 100.0, " ' VLM.Lampz;inserts-L29
  Lampz.Callback(29) = "UpdateLightMap rsling002_LM_inserts_L29, 100.0, " ' VLM.Lampz;inserts-L29
  ' Lampz.MassAssign(30) = L30 ' VLM.Lampz;inserts-L30
  Lampz.Callback(30) = "UpdateLightMap Playfield_LM_inserts_L30, 100.0, " ' VLM.Lampz;inserts-L30
  Lampz.Callback(30) = "UpdateLightMap Overlay_LM_inserts_L30, 100.0, " ' VLM.Lampz;inserts-L30
  Lampz.Callback(30) = "UpdateLightMap Cherry_Switch_003_LM_inserts_L3, 100.0, " ' VLM.Lampz;inserts-L30
  Lampz.Callback(30) = "UpdateLightMap LFlip_001_LM_inserts_L30, 100.0, " ' VLM.Lampz;inserts-L30
  Lampz.Callback(30) = "UpdateLightMap LFlipR_LM_inserts_L30, 100.0, " ' VLM.Lampz;inserts-L30
  Lampz.Callback(30) = "UpdateLightMap lsling_LM_inserts_L30, 100.0, " ' VLM.Lampz;inserts-L30
  Lampz.Callback(30) = "UpdateLightMap lsling001_LM_inserts_L30, 100.0, " ' VLM.Lampz;inserts-L30
  Lampz.Callback(30) = "UpdateLightMap lsling002_LM_inserts_L30, 100.0, " ' VLM.Lampz;inserts-L30
  Lampz.Callback(30) = "UpdateLightMap siderails_LM_inserts_L30, 100.0, " ' VLM.Lampz;inserts-L30
  ' Lampz.MassAssign(31) = L31 ' VLM.Lampz;inserts-L31
  Lampz.Callback(31) = "UpdateLightMap Playfield_LM_inserts_L31, 100.0, " ' VLM.Lampz;inserts-L31
  Lampz.Callback(31) = "UpdateLightMap Overlay_LM_inserts_L31, 100.0, " ' VLM.Lampz;inserts-L31
  Lampz.Callback(31) = "UpdateLightMap Cherry_Switch_007_LM_inserts_L3, 100.0, " ' VLM.Lampz;inserts-L31
  Lampz.Callback(31) = "UpdateLightMap Cherry_Switch_008_LM_inserts_L3, 100.0, " ' VLM.Lampz;inserts-L31
  Lampz.Callback(31) = "UpdateLightMap Circle_019_LM_inserts_L31, 100.0, " ' VLM.Lampz;inserts-L31
  Lampz.Callback(31) = "UpdateLightMap LFlip_001_LM_inserts_L31, 100.0, " ' VLM.Lampz;inserts-L31
  Lampz.Callback(31) = "UpdateLightMap LFlipR_LM_inserts_L31, 100.0, " ' VLM.Lampz;inserts-L31
  Lampz.Callback(31) = "UpdateLightMap lsling_LM_inserts_L31, 100.0, " ' VLM.Lampz;inserts-L31
  Lampz.Callback(31) = "UpdateLightMap lsling001_LM_inserts_L31, 100.0, " ' VLM.Lampz;inserts-L31
  Lampz.Callback(31) = "UpdateLightMap lsling002_LM_inserts_L31, 100.0, " ' VLM.Lampz;inserts-L31
  Lampz.Callback(31) = "UpdateLightMap siderails_LM_inserts_L31, 100.0, " ' VLM.Lampz;inserts-L31
  ' Lampz.MassAssign(32) = L32 ' VLM.Lampz;inserts-L32
  Lampz.Callback(32) = "UpdateLightMap Playfield_LM_inserts_L32, 100.0, " ' VLM.Lampz;inserts-L32
  Lampz.Callback(32) = "UpdateLightMap Overlay_LM_inserts_L32, 100.0, " ' VLM.Lampz;inserts-L32
  Lampz.Callback(32) = "UpdateLightMap Cherry_Switch_006_LM_inserts_L3, 100.0, " ' VLM.Lampz;inserts-L32
  Lampz.Callback(32) = "UpdateLightMap Cherry_Switch_007_LM_inserts_L3, 100.0, " ' VLM.Lampz;inserts-L32
  Lampz.Callback(32) = "UpdateLightMap Circle_019_LM_inserts_L32, 100.0, " ' VLM.Lampz;inserts-L32
  Lampz.Callback(32) = "UpdateLightMap RFlip_LM_inserts_L32, 100.0, " ' VLM.Lampz;inserts-L32
  Lampz.Callback(32) = "UpdateLightMap RFlipR_LM_inserts_L32, 100.0, " ' VLM.Lampz;inserts-L32
  Lampz.Callback(32) = "UpdateLightMap rsling_LM_inserts_L32, 100.0, " ' VLM.Lampz;inserts-L32
  Lampz.Callback(32) = "UpdateLightMap rsling001_LM_inserts_L32, 100.0, " ' VLM.Lampz;inserts-L32
  Lampz.Callback(32) = "UpdateLightMap rsling002_LM_inserts_L32, 100.0, " ' VLM.Lampz;inserts-L32
  ' Lampz.MassAssign(33) = L33 ' VLM.Lampz;inserts-L33
  Lampz.Callback(33) = "UpdateLightMap Playfield_LM_inserts_L33, 100.0, " ' VLM.Lampz;inserts-L33
  Lampz.Callback(33) = "UpdateLightMap Overlay_LM_inserts_L33, 100.0, " ' VLM.Lampz;inserts-L33
  Lampz.Callback(33) = "UpdateLightMap Cherry_Switch_004_LM_inserts_L3, 100.0, " ' VLM.Lampz;inserts-L33
  Lampz.Callback(33) = "UpdateLightMap Cherry_Switch_005_LM_inserts_L3, 100.0, " ' VLM.Lampz;inserts-L33
  Lampz.Callback(33) = "UpdateLightMap RFlip_LM_inserts_L33, 100.0, " ' VLM.Lampz;inserts-L33
  Lampz.Callback(33) = "UpdateLightMap RFlipR_LM_inserts_L33, 100.0, " ' VLM.Lampz;inserts-L33
  Lampz.Callback(33) = "UpdateLightMap rsling_LM_inserts_L33, 100.0, " ' VLM.Lampz;inserts-L33
  Lampz.Callback(33) = "UpdateLightMap rsling001_LM_inserts_L33, 100.0, " ' VLM.Lampz;inserts-L33
  Lampz.Callback(33) = "UpdateLightMap rsling002_LM_inserts_L33, 100.0, " ' VLM.Lampz;inserts-L33
  ' Lampz.MassAssign(34) = L34 ' VLM.Lampz;inserts-L34
  Lampz.Callback(34) = "UpdateLightMap Playfield_LM_inserts_L34, 100.0, " ' VLM.Lampz;inserts-L34
  Lampz.Callback(34) = "UpdateLightMap Overlay_LM_inserts_L34, 100.0, " ' VLM.Lampz;inserts-L34
  Lampz.Callback(34) = "UpdateLightMap Circle_016_LM_inserts_L34, 100.0, " ' VLM.Lampz;inserts-L34
  Lampz.Callback(34) = "UpdateLightMap LFlip_001_LM_inserts_L34, 100.0, " ' VLM.Lampz;inserts-L34
  Lampz.Callback(34) = "UpdateLightMap Pop_Bumper_WPC_003_LM_inserts_L, 100.0, " ' VLM.Lampz;inserts-L34
  Lampz.Callback(34) = "UpdateLightMap RFlip_LM_inserts_L34, 100.0, " ' VLM.Lampz;inserts-L34
  ' Lampz.MassAssign(35) = L35 ' VLM.Lampz;inserts-L35
  Lampz.Callback(35) = "UpdateLightMap Playfield_LM_inserts_L35, 100.0, " ' VLM.Lampz;inserts-L35
  Lampz.Callback(35) = "UpdateLightMap Circle_015_LM_inserts_L35, 100.0, " ' VLM.Lampz;inserts-L35
  Lampz.Callback(35) = "UpdateLightMap Circle_019_LM_inserts_L35, 100.0, " ' VLM.Lampz;inserts-L35
  Lampz.Callback(35) = "UpdateLightMap Circle_037_LM_inserts_L35, 100.0, " ' VLM.Lampz;inserts-L35
  Lampz.Callback(35) = "UpdateLightMap LFlip_001_LM_inserts_L35, 100.0, " ' VLM.Lampz;inserts-L35
  Lampz.Callback(35) = "UpdateLightMap LFlipR_LM_inserts_L35, 100.0, " ' VLM.Lampz;inserts-L35
  Lampz.Callback(35) = "UpdateLightMap Pop_Bumper_WPC_002_LM_inserts_L, 100.0, " ' VLM.Lampz;inserts-L35
  Lampz.Callback(35) = "UpdateLightMap Pop_Bumper_WPC_005_LM_inserts_L, 100.0, " ' VLM.Lampz;inserts-L35
  Lampz.Callback(35) = "UpdateLightMap RFlip_LM_inserts_L35, 100.0, " ' VLM.Lampz;inserts-L35
  Lampz.Callback(35) = "UpdateLightMap RFlipR_LM_inserts_L35, 100.0, " ' VLM.Lampz;inserts-L35
End Sub


' ===============================================================
' The following code can serve as a base for movable position synchronization.
' You will need to adapt the part of the transform you want to synchronize
' and the source on which you want it to be synchronized.

'VLM movable script
'Sub CabinetModeHelper
' If CabinetMode Then
'   For each bl in lockdown_BL : bl.visible = 0 : Next
'   For each bl in siderails_BL : bl.visible = 0 : Next
' Else
'   For each bl in lockdown_BL : bl.visible = 1 : Next
'   For each bl in siderails_BL : bl.visible = 1 : Nex
' End If
'End Sub


'VLM movable script
Sub MovableHelper
  dim bl

  'Update Flippers
  dim lfa: lfa = -LeftFlipper.CurrentAngle
  For each bl in LFlip_001_BL : bl.objRotZ = lfa : Next
  For each bl in LFlipR_BL : bl.objRotZ = lfa : Next
  FlipperLSh.RotZ = -lfa

  dim rfa: rfa = -RightFlipper.CurrentAngle
  For each bl in RFlip_BL : bl.objRotZ = rfa : Next
  For each bl in RFlipR_BL : bl.objRotZ = rfa : Next
  FlipperRSh.RotZ = -rfa


  'Update Gates
  dim ga
  ga = Gate.currentAngle
  For each bl in pgate_BL : bl.rotx = ga : Next
End Sub


'VLM movable script
Sub STMovableHelper
  'Update Standup Targets
  dim ty, lm

  ty = TopTargetLeft_001_BM_World.transy
  For each lm in TopTargetLeft_001_lm : lm.transy = ty : Next

  ty = TopTargetRight_001_BM_World.transy
  For each lm in TopTargetRight_001_lm : lm.transy = ty : Next

  ty = MidTargetLeft_BM_World.transy
  For each lm in MidTargetLeft_lm : lm.transy = ty : Next

  ty = CenterTarget_BM_World.transy
  For each lm in CenterTarget_lm : lm.transy = ty : Next

  ty = MidTargetRight_BM_World.transy
  For each lm in MidTargetRight_lm : lm.transy = ty : Next

End Sub


'VLM movable script
Sub Update_Wires(wire, pushed)
  dim bl
  Dim z : If pushed Then z = -14 Else z = 0
  Select Case wire
    Case 0: For each bl in Cherry_Switch_000_BL : bl.transz = z : Next
    Case 1: For each bl in Cherry_Switch_001_BL : bl.transz = z : Next
    Case 2: For each bl in Cherry_Switch_002_BL : bl.transz = z : Next
    Case 3: For each bl in Cherry_Switch_003_BL : bl.transz = z : Next
    Case 4: For each bl in Cherry_Switch_004_BL : bl.transz = z : Next
    Case 5: For each bl in Cherry_Switch_005_BL : bl.transz = z : Next
    Case 6: For each bl in Cherry_Switch_006_BL : bl.transz = z : Next
    Case 7: For each bl in Cherry_Switch_007_BL : bl.transz = z : Next
    Case 8: For each bl in Cherry_Switch_008_BL : bl.transz = z : Next
    Case 9: For each bl in Cherry_Switch_009_BL : bl.transz = z : Next
    Case 10: For each bl in Cherry_Switch_010_BL : bl.transz = z : Next
    Case 11: For each bl in Cherry_Switch_011_BL : bl.transz = z : Next
  End Select
End Sub



'******************************************************
'   STAND-UP TARGET INITIALIZATION
'******************************************************

'Define a variable for each stand-up target
Dim ST1, ST2, ST3, ST4, ST5

'Set array with stand-up target objects
'
'StandupTargetvar = Array(primary, prim, swtich)
'   primary:      vp target to determine target hit
' prim:       primitive target used for visuals and animation
'             IMPORTANT!!!
'             transy must be used to offset the target animation
' switch:       ROM switch number
' animate:      Arrary slot for handling the animation instrucitons, set to 0
'
'You will also need to add a secondary hit object for each stand up (name sw11o, sw12o, and sw13o on the example Table1)
'these are inclined primitives to simulate hitting a bent target and should provide so z velocity on high speed impacts

ST1 = Array(TopTargetLeft, TopTargetLeft_001_BM_World,1,0)
ST2 = Array(TopTargetRight, TopTargetRight_001_BM_World,2,0)
ST3 = Array(MidTargetLeft, MidTargetLeft_BM_World,3,0)
ST4 = Array(CenterTarget, CenterTarget_BM_World,4,0)
ST5 = Array(MidTargetRight, MidTargetRight_BM_World,5,0)

'Add all the Stand-up Target Arrays to Stand-up Target Animation Array
' STAnimationArray = Array(ST1, ST2, ....)
Dim STArray
STArray = Array (ST1, ST2, ST3, ST4, ST5)

'Configure the behavior of Stand-up Targets
Const STAnimStep =  1.5     'vpunits per animation step (control return to Start)
Const STMaxOffset = 9       'max vp units target moves when hit

Const STMass = 0.2        'Mass of the Stand-up Target (between 0 and 1), higher values provide more resistance

'******************************************************
'       STAND-UP TARGETS FUNCTIONS
'******************************************************

Sub STHit(switch)
  Dim i
  i = STArrayID(switch)

  'PlayTargetSound
  STArray(i)(3) =  STCheckHit(Activeball,STArray(i)(0))

  If STArray(i)(3) <> 0 Then
    DTBallPhysics Activeball, STArray(i)(0).orientation, STMass
  End If
  DoSTAnim
End Sub

Function STArrayID(switch)
  Dim i
  For i = 0 to uBound(STArray)
    If STArray(i)(2) = switch Then STArrayID = i:Exit Function
  Next
End Function

'Check if target is hit on it's face
Function STCheckHit(aBall, target)
  dim bangle, bangleafter, rangle, rangle2, perpvel, perpvelafter, paravel, paravelafter
  rangle = (target.orientation - 90) * 3.1416 / 180
  bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))
  bangleafter = Atn2(aBall.vely,aball.velx)

  perpvel = cor.BallVel(aball.id) * cos(bangle-rangle)
  paravel = cor.BallVel(aball.id) * sin(bangle-rangle)

  perpvelafter = BallSpeed(aBall) * cos(bangleafter - rangle)
  paravelafter = BallSpeed(aBall) * sin(bangleafter - rangle)

  If perpvel > 0 and  perpvelafter <= 0 Then
    STCheckHit = 1
  ElseIf perpvel > 0 and ((paravel > 0 and paravelafter > 0) or (paravel < 0 and paravelafter < 0)) Then
    STCheckHit = 1
  Else
    STCheckHit = 0
  End If
End Function

Sub DoSTAnim()
  Dim i
  For i=0 to Ubound(STArray)
    STArray(i)(3) = STAnimate(STArray(i)(0),STArray(i)(1),STArray(i)(2),STArray(i)(3))
  Next
End Sub

Function STAnimate(primary, prim, switch,  animate)
  Dim animtime

  STAnimate = animate

  if animate = 0  Then
    primary.uservalue = 0
    STAnimate = 0
    Exit Function
  Elseif primary.uservalue = 0 then
    primary.uservalue = gametime
  end if

  animtime = gametime - primary.uservalue

  If animate = 1 Then
    primary.collidable = 0
    prim.transy = -STMaxOffset
    STAction switch
    STAnimate = 2
    Exit Function
  elseif animate = 2 Then
    prim.transy = prim.transy + STAnimStep
    If prim.transy >= 0 Then
      prim.transy = 0
      primary.collidable = 1
      STAnimate = 0
      Exit Function
    Else
      STAnimate = 2
    End If
  End If
End Function

sub DTBallPhysics(aBall, angle, mass)
  dim rangle,bangle,calc1, calc2, calc3
  rangle = (angle - 90) * 3.1416 / 180
  bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))

  calc1 = cor.BallVel(aball.id) * cos(bangle - rangle) * (aball.mass - mass) / (aball.mass + mass)
  calc2 = cor.BallVel(aball.id) * sin(bangle - rangle) * cos(rangle + 4*Atn(1)/2)
  calc3 = cor.BallVel(aball.id) * sin(bangle - rangle) * sin(rangle + 4*Atn(1)/2)

  aBall.velx = calc1 * cos(rangle) + calc2
  aBall.vely = calc1 * sin(rangle) + calc3
End Sub




'***********************************

Sub TopTargetLeft_Hit :  STHit 1 : End Sub
Sub TopTargetRight_Hit : STHit 2 : End Sub
Sub MidTargetLeft_Hit :  STHit 3 : End Sub
Sub CenterTarget_Hit :   STHit 4 : End Sub
Sub MidTargetRight_Hit : STHit 5 : End Sub


'***********************************


Sub STAction(Switch)
  if TableTilted=false then
    Select Case Switch
      Case 1:
        DOF 216,2
        AddScore(10)
        If MagicCityLight006.state=1 then
          MagicCityLight006.state=0
          MagicCity2007.state=1
          CheckMagicCity
        end if

      Case 2:
        DOF 217,2
        AddScore(10)
        If MagicCityLight007.state=1 then
          MagicCityLight007.state=0
          MagicCity2008.state=1
          CheckMagicCity
        end if

      Case 3:
        DOF 216,2
        AddScore(10)
        If MagicCityLight008.state=1 then
          MagicCityLight008.state=0
          MagicCity2009.state=1
          CheckMagicCity
        end if

      Case 4:
        DOF 216,2
        SetMotor(50)
        MagicCity2Lights(ZeroNineUnit).state=1

        Bumper2Light.state=1
        'Bumper3Light.state=1
        'Bumper4Light.state=1
        if ZeroNineUnit<5 then
          MagicCityLights(ZeroNineUnit).state=0
        elseif ZeroNineUnit=5 then
          Bumper1Light.state=1
          Special002.state=1
          Special003.state=1
        else
          MagicCityLights(ZeroNineUnit-1).state=0
        end if
        CheckMagicCity
        ToggleZeroNine

      Case 5:
    DOF 216,2
      AddScore(10)
      If MagicCityLight009.state=1 then
        MagicCityLight009.state=0
        MagicCity2010.state=1
        CheckMagicCity
      end if

    End Select
  End If
End Sub




'******************************************************
' Flippers Polarity (Select appropriate sub based on era)
'******************************************************

Dim LF
Set LF = New FlipperPolarity
Dim RF
Set RF = New FlipperPolarity

InitPolarity


'*******************************************
' Late 70's to early 80's

Sub InitPolarity()
   dim x, a : a = Array(LF, RF)
  for each x in a
    x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
    x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
    x.enabled = True
    x.TimeDelay = 80
    x.DebugOn=False ' prints some info in debugger

    x.AddPt "Polarity", 0, 0, 0
    x.AddPt "Polarity", 1, 0.05, - 2.7
    x.AddPt "Polarity", 2, 0.33, - 2.7
    x.AddPt "Polarity", 3, 0.37, - 2.7
    x.AddPt "Polarity", 4, 0.41, - 2.7
    x.AddPt "Polarity", 5, 0.45, - 2.7
    x.AddPt "Polarity", 6, 0.576, - 2.7
    x.AddPt "Polarity", 7, 0.66, - 1.8
    x.AddPt "Polarity", 8, 0.743, - 0.5
    x.AddPt "Polarity", 9, 0.81, - 0.5
    x.AddPt "Polarity", 10, 0.88, 0

    x.AddPt "Velocity", 0, 0, 1
    x.AddPt "Velocity", 1, 0.16, 1.06
    x.AddPt "Velocity", 2, 0.41, 1.05
    x.AddPt "Velocity", 3, 0.53, 1 '0.982
    x.AddPt "Velocity", 4, 0.702, 0.968
    x.AddPt "Velocity", 5, 0.95,  0.968
    x.AddPt "Velocity", 6, 1.03, 0.945
  Next

  ' SetObjects arguments: 1: name of object 2: flipper object: 3: Trigger object around flipper
    LF.SetObjects "LF", LeftFlipper, TriggerLF
    RF.SetObjects "RF", RightFlipper, TriggerRF
End Sub


'******************************************************
'  FLIPPER CORRECTION FUNCTIONS
'******************************************************

' modified 2023 by nFozzy
' Removed need for 'endpoint' objects
' Added 'createvents' type thing for TriggerLF / TriggerRF triggers.
' Removed AddPt function which complicated setup imo
' made DebugOn do something (prints some stuff in debugger)
'   Otherwise it should function exactly the same as before

Class FlipperPolarity
  Public DebugOn, Enabled
  Private FlipAt    'Timer variable (IE 'flip at 723,530ms...)
  Public TimeDelay    'delay before trigger turns off and polarity is disabled
  private Flipper, FlipperStart, FlipperEnd, FlipperEndY, LR, PartialFlipCoef
  Private Balls(20), balldata(20)
  private Name

  dim PolarityIn, PolarityOut
  dim VelocityIn, VelocityOut
  dim YcoefIn, YcoefOut
  Public Sub Class_Initialize
    redim PolarityIn(0) : redim PolarityOut(0) : redim VelocityIn(0) : redim VelocityOut(0) : redim YcoefIn(0) : redim YcoefOut(0)
    Enabled = True : TimeDelay = 50 : LR = 1:  dim x : for x = 0 to uBound(balls) : balls(x) = Empty : set Balldata(x) = new SpoofBall : next
  End Sub

  Public Sub SetObjects(aName, aFlipper, aTrigger)

    if typename(aName) <> "String" then msgbox "FlipperPolarity: .SetObjects error: first argument must be a string (and name of Object). Found:" & typename(aName) end if
    if typename(aFlipper) <> "Flipper" then msgbox "FlipperPolarity: .SetObjects error: second argument must be a flipper. Found:" & typename(aFlipper) end if
    if typename(aTrigger) <> "Trigger" then msgbox "FlipperPolarity: .SetObjects error: third argument must be a trigger. Found:" & typename(aTrigger) end if
    if aFlipper.EndAngle > aFlipper.StartAngle then LR = -1 Else LR = 1 End If
    Name = aName
    Set Flipper = aFlipper : FlipperStart = aFlipper.x
    FlipperEnd = Flipper.Length * sin((Flipper.StartAngle / 57.295779513082320876798154814105)) + Flipper.X ' big floats for degree to rad conversion
    FlipperEndY = Flipper.Length * cos(Flipper.StartAngle / 57.295779513082320876798154814105)*-1 + Flipper.Y

    dim str : str = "sub " & aTrigger.name & "_Hit() : " & aName & ".AddBall ActiveBall : End Sub'"
    ExecuteGlobal(str)
    str = "sub " & aTrigger.name & "_UnHit() : " & aName & ".PolarityCorrect ActiveBall : End Sub'"
    ExecuteGlobal(str)

  End Sub

  Public Property Let EndPoint(aInput) :  : End Property ' Legacy: just no op

  Public Sub AddPt(aChooseArray, aIDX, aX, aY) 'Index #, X position, (in) y Position (out)
    Select Case aChooseArray
      case "Polarity" : ShuffleArrays PolarityIn, PolarityOut, 1 : PolarityIn(aIDX) = aX : PolarityOut(aIDX) = aY : ShuffleArrays PolarityIn, PolarityOut, 0
      Case "Velocity" : ShuffleArrays VelocityIn, VelocityOut, 1 :VelocityIn(aIDX) = aX : VelocityOut(aIDX) = aY : ShuffleArrays VelocityIn, VelocityOut, 0
      Case "Ycoef" : ShuffleArrays YcoefIn, YcoefOut, 1 :YcoefIn(aIDX) = aX : YcoefOut(aIDX) = aY : ShuffleArrays YcoefIn, YcoefOut, 0
    End Select
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
  Private Function FlipperOn() : if gameTime < FlipAt+TimeDelay then FlipperOn = True : End If : End Function   'Timer shutoff for polaritycorrect

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
          if ballpos > 0.65 then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)                'find safety coefficient 'ycoef' data
        end if
      Next

      If BallPos = 0 Then 'no ball data meaning the ball is entering and exiting pretty close to the same position, use current values.
        BallPos = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
        if ballpos > 0.65 then  Ycoef = LinearEnvelope(aBall.Y, YcoefIn, YcoefOut)                        'find safety coefficient 'ycoef' data
      End If

      'Velocity correction
      if not IsEmpty(VelocityIn(0) ) then
        Dim VelCoef
        VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)

        if partialflipcoef < 1 then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)

        if Enabled then aBall.Velx = aBall.Velx*VelCoef
        if Enabled then aBall.Vely = aBall.Vely*VelCoef
      End If

      'Polarity Correction (optional now)
      if not IsEmpty(PolarityIn(0) ) then
        dim AddX : AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR

        if Enabled then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef)
      End If
      if DebugOn then debug.print "PolarityCorrect" & " " & Name & " @ " & gametime & " " & Round(BallPos*100) & "%" & " AddX:" & Round(AddX,2) & " Vel%:" & Round(VelCoef*100)
    End If
    RemoveBall aBall
  End Sub
End Class

'******************************************************
'  FLIPPER POLARITY AND RUBBER DAMPENER SUPPORTING FUNCTIONS
'******************************************************

' Used for flipper correction and rubber dampeners
Sub ShuffleArray(ByRef aArray, byVal offset) 'shuffle 1d array
  dim x, aCount : aCount = 0
  redim a(uBound(aArray) )
  for x = 0 to uBound(aArray)   'Shuffle objects in a temp array
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
  redim aArray(aCount-1+offset)   'Resize original array
  for x = 0 to aCount-1       'set objects back into original array
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
Function PSlope(Input, X1, Y1, X2, Y2)    'Set up line via two points, no clamping. Input X, output Y
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
  dim ii : for ii = 1 to uBound(xKeyFrame)    'find active line
    if xInput <= xKeyFrame(ii) then L = ii : exit for : end if
  Next
  if xInput > xKeyFrame(uBound(xKeyFrame) ) then L = uBound(xKeyFrame)    'catch line overrun
  Y = pSlope(xInput, xKeyFrame(L-1), yLvl(L-1), xKeyFrame(L), yLvl(L) )

  if xInput <= xKeyFrame(lBound(xKeyFrame) ) then Y = yLvl(lBound(xKeyFrame) )     'Clamp lower
  if xInput >= xKeyFrame(uBound(xKeyFrame) ) then Y = yLvl(uBound(xKeyFrame) )    'Clamp upper

  LinearEnvelope = Y
End Function

'******************************************************
'  FLIPPER TRICKS
'******************************************************

RightFlipper.timerinterval = 1
Rightflipper.timerenabled = True

Sub RightFlipper_timer()
  FlipperTricks LeftFlipper, LFPress, LFCount, LFEndAngle, LFState
  FlipperTricks RightFlipper, RFPress, RFCount, RFEndAngle, RFState
  FlipperNudge RightFlipper, RFEndAngle, RFEOSNudge, LeftFlipper, LFEndAngle
  FlipperNudge LeftFlipper, LFEndAngle, LFEOSNudge,  RightFlipper, RFEndAngle
End Sub

Dim LFEOSNudge, RFEOSNudge

Sub FlipperNudge(Flipper1, Endangle1, EOSNudge1, Flipper2, EndAngle2)
  Dim b
  Dim BOT
  BOT = GetBalls

  If Flipper1.currentangle = Endangle1 And EOSNudge1 <> 1 Then
    EOSNudge1 = 1
    '   debug.print Flipper1.currentangle &" = "& Endangle1 &"--"& Flipper2.currentangle &" = "& EndAngle2
    If Flipper2.currentangle = EndAngle2 Then
      For b = 0 To UBound(BOT)
        If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper1) Then
          'Debug.Print "ball in flip1. exit"
          Exit Sub
        End If
      Next
      For b = 0 To UBound(BOT)
        If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper2) Then
          BOT(b).velx = BOT(b).velx / 1.3
          BOT(b).vely = BOT(b).vely - 0.5
        End If
      Next
    End If
  Else
    If Abs(Flipper1.currentangle) > Abs(EndAngle1) + 30 Then EOSNudge1 = 0
  End If
End Sub

'*****************
' Maths
'*****************

Dim PI
PI = 4 * Atn(1)

Function dSin(degrees)
  dsin = Sin(degrees * Pi / 180)
End Function

Function dCos(degrees)
  dcos = Cos(degrees * Pi / 180)
End Function

Function Atn2(dy, dx)
  If dx > 0 Then
    Atn2 = Atn(dy / dx)
  ElseIf dx < 0 Then
    If dy = 0 Then
      Atn2 = pi
    Else
      Atn2 = Sgn(dy) * (pi - Atn(Abs(dy / dx)))
    End If
  ElseIf dx = 0 Then
    If dy = 0 Then
      Atn2 = 0
    Else
      Atn2 = Sgn(dy) * pi / 2
    End If
  End If
End Function

Function max(a,b)
  If a > b Then
    max = a
  Else
    max = b
  End If
End Function

' Used for drop targets
Function InRect(px,py,ax,ay,bx,by,cx,cy,dx,dy) 'Determines if a Points (px,py) is inside a 4 point polygon A-D in Clockwise/CCW order
  Dim AB, BC, CD, DA
  AB = (bx * py) - (by * px) - (ax * py) + (ay * px) + (ax * by) - (ay * bx)
  BC = (cx * py) - (cy * px) - (bx * py) + (by * px) + (bx * cy) - (by * cx)
  CD = (dx * py) - (dy * px) - (cx * py) + (cy * px) + (cx * dy) - (cy * dx)
  DA = (ax * py) - (ay * px) - (dx * py) + (dy * px) + (dx * ay) - (dy * ax)

  If (AB <= 0 And BC <= 0 And CD <= 0 And DA <= 0) Or (AB >= 0 And BC >= 0 And CD >= 0 And DA >= 0) Then
    InRect = True
  Else
    InRect = False
  End If
End Function

Function InRotRect(ballx,bally,px,py,angle,ax,ay,bx,by,cx,cy,dx,dy)
  Dim rax,ray,rbx,rby,rcx,rcy,rdx,rdy
  Dim rotxy
  rotxy = RotPoint(ax,ay,angle)
  rax = rotxy(0) + px
  ray = rotxy(1) + py
  rotxy = RotPoint(bx,by,angle)
  rbx = rotxy(0) + px
  rby = rotxy(1) + py
  rotxy = RotPoint(cx,cy,angle)
  rcx = rotxy(0) + px
  rcy = rotxy(1) + py
  rotxy = RotPoint(dx,dy,angle)
  rdx = rotxy(0) + px
  rdy = rotxy(1) + py

  InRotRect = InRect(ballx,bally,rax,ray,rbx,rby,rcx,rcy,rdx,rdy)
End Function

Function RotPoint(x,y,angle)
  Dim rx, ry
  rx = x * dCos(angle) - y * dSin(angle)
  ry = x * dSin(angle) + y * dCos(angle)
  RotPoint = Array(rx,ry)
End Function


'*************************************************
'  Check ball distance from Flipper for Rem
'*************************************************

Function Distance(ax,ay,bx,by)
  Distance = Sqr((ax - bx) ^ 2 + (ay - by) ^ 2)
End Function

Function DistancePL(px,py,ax,ay,bx,by) 'Distance between a point and a line where point is px,py
  DistancePL = Abs((by - ay) * px - (bx - ax) * py + bx * ay - by * ax) / Distance(ax,ay,bx,by)
End Function

Function Radians(Degrees)
  Radians = Degrees * PI / 180
End Function

Function AnglePP(ax,ay,bx,by)
  AnglePP = Atn2((by - ay),(bx - ax)) * 180 / PI
End Function

Function DistanceFromFlipper(ballx, bally, Flipper)
  DistanceFromFlipper = DistancePL(ballx, bally, Flipper.x, Flipper.y, Cos(Radians(Flipper.currentangle + 90)) + Flipper.x, Sin(Radians(Flipper.currentangle + 90)) + Flipper.y)
End Function

Function FlipperTrigger(ballx, bally, Flipper)
  Dim DiffAngle
  DiffAngle = Abs(Flipper.currentangle - AnglePP(Flipper.x, Flipper.y, ballx, bally) - 90)
  If DiffAngle > 180 Then DiffAngle = DiffAngle - 360

  If DistanceFromFlipper(ballx,bally,Flipper) < 48 And DiffAngle <= 90 And Distance(ballx,bally,Flipper.x,Flipper.y) < Flipper.Length Then
    FlipperTrigger = True
  Else
    FlipperTrigger = False
  End If
End Function

'*************************************************
'  End - Check ball distance from Flipper for Rem
'*************************************************

Dim LFPress, RFPress, LFCount, RFCount
Dim LFState, RFState
Dim EOST, EOSA,Frampup, FElasticity,FReturn
Dim RFEndAngle, LFEndAngle

Const FlipperCoilRampupMode = 0 '0 = fast, 1 = medium, 2 = slow (tap passes should work)

LFState = 1
RFState = 1
EOST = leftflipper.eostorque
EOSA = leftflipper.eostorqueangle
Frampup = LeftFlipper.rampup
FElasticity = LeftFlipper.elasticity
FReturn = LeftFlipper.return
Const EOSTnew = 1 'EM's to late 80's
Const EOSAnew = 1
Const EOSRampup = 0
Dim SOSRampup
Select Case FlipperCoilRampupMode
  Case 0
  SOSRampup = 2.5
  Case 1
  SOSRampup = 6
  Case 2
  SOSRampup = 8.5
End Select

Const LiveCatch = 16
Const LiveElasticity = 0.45
Const SOSEM = 0.815
Const EOSReturn = 0.055  'EM's


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
  Flipper.eostorque = EOST * EOSReturn / FReturn

  If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 0.1 Then
    Dim b, BOT
    BOT = GetBalls

    For b = 0 To UBound(BOT)
      If Distance(BOT(b).x, BOT(b).y, Flipper.x, Flipper.y) < 55 Then 'check for cradle
        If BOT(b).vely >= - 0.4 Then BOT(b).vely =  - 0.4
      End If
    Next
  End If
End Sub

Sub FlipperTricks (Flipper, FlipperPress, FCount, FEndAngle, FState)
  Dim Dir
  Dir = Flipper.startangle / Abs(Flipper.startangle) '-1 for Right Flipper

  If Abs(Flipper.currentangle) > Abs(Flipper.startangle) - 0.05 Then
    If FState <> 1 Then
      Flipper.rampup = SOSRampup
      Flipper.endangle = FEndAngle - 3 * Dir
      Flipper.Elasticity = FElasticity * SOSEM
      FCount = 0
      FState = 1
    End If
  ElseIf Abs(Flipper.currentangle) <= Abs(Flipper.endangle) And FlipperPress = 1 Then
    If FCount = 0 Then FCount = GameTime

    If FState <> 2 Then
      Flipper.eostorqueangle = EOSAnew
      Flipper.eostorque = EOSTnew
      Flipper.rampup = EOSRampup
      Flipper.endangle = FEndAngle
      FState = 2
    End If
  ElseIf Abs(Flipper.currentangle) > Abs(Flipper.endangle) + 0.01 And FlipperPress = 1 Then
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
Const LiveDistanceMax = 114 'maximum distance in vp units from flipper base live catch dampening will occur (tip protection)

Sub CheckLiveCatch(ball, Flipper, FCount, parm) 'Experimental new live catch
  Dim Dir
  Dir = Flipper.startangle / Abs(Flipper.startangle)  '-1 for Right Flipper
  Dim LiveCatchBounce                                                           'If live catch is not perfect, it won't freeze ball totally
  Dim CatchTime
  CatchTime = GameTime - FCount

  If CatchTime <= LiveCatch And parm > 6 And Abs(Flipper.x - ball.x) > LiveDistanceMin And Abs(Flipper.x - ball.x) < LiveDistanceMax Then
    If CatchTime <= LiveCatch * 0.5 Then                        'Perfect catch only when catch time happens in the beginning of the window
      LiveCatchBounce = 0
    Else
      LiveCatchBounce = Abs((LiveCatch / 2) - CatchTime)    'Partial catch when catch happens a bit late
    End If

    If LiveCatchBounce = 0 And ball.velx * Dir > 0 Then ball.velx = 0
    ball.vely = LiveCatchBounce * (32 / LiveCatch) ' Multiplier for inaccuracy bounce
    ball.angmomx = 0
    ball.angmomy = 0
    ball.angmomz = 0
  Else
    If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 1 Then FlippersD.Dampenf Activeball, parm
  End If
End Sub

'******************************************************
'****  END FLIPPER CORRECTIONS
'******************************************************





'******************************************************
'   ZDMP:  RUBBER  DAMPENERS
'******************************************************
' These are data mined bounce curves,
' dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
' Requires tracking ballspeed to calculate COR

Sub dPosts_Hit(idx)
  RubbersD.dampen Activeball
End Sub

Sub dSleeves_Hit(idx)
  SleevesD.Dampen Activeball
End Sub

Dim RubbersD        'frubber
Set RubbersD = New Dampener
RubbersD.name = "Rubbers"
RubbersD.debugOn = False  'shows info in textbox "TBPout"
RubbersD.Print = False    'debug, reports in debugger (in vel, out cor); cor bounce curve (linear)

'for best results, try to match in-game velocity as closely as possible to the desired curve
'   RubbersD.addpoint 0, 0, 0.935   'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 0, 0, 1.1    'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 1, 3.77, 0.97
RubbersD.addpoint 2, 5.76, 0.967  'dont take this as gospel. if you can data mine rubber elasticitiy, please help!
RubbersD.addpoint 3, 15.84, 0.874
RubbersD.addpoint 4, 56, 0.64    'there's clamping so interpolate up to 56 at least

Dim SleevesD  'this is just rubber but cut down to 85%...
Set SleevesD = New Dampener
SleevesD.name = "Sleeves"
SleevesD.debugOn = False  'shows info in textbox "TBPout"
SleevesD.Print = False    'debug, reports in debugger (in vel, out cor)
SleevesD.CopyCoef RubbersD, 0.85

'######################### Add new FlippersD Profile
'######################### Adjust these values to increase or lessen the elasticity

Dim FlippersD
Set FlippersD = New Dampener
FlippersD.name = "Flippers"
FlippersD.debugOn = False
FlippersD.Print = False
FlippersD.addpoint 0, 0, 1.1
FlippersD.addpoint 1, 3.77, 0.99
FlippersD.addpoint 2, 6, 0.99

Class Dampener
  Public Print, debugOn   'tbpOut.text
  Public name, Threshold  'Minimum threshold. Useful for Flippers, which don't have a hit threshold.
  Public ModIn, ModOut
  Private Sub Class_Initialize
    ReDim ModIn(0)
    ReDim Modout(0)
  End Sub

  Public Sub AddPoint(aIdx, aX, aY)
    ShuffleArrays ModIn, ModOut, 1
    ModIn(aIDX) = aX
    ModOut(aIDX) = aY
    ShuffleArrays ModIn, ModOut, 0
    If gametime > 100 Then Report
  End Sub

  Public Sub Dampen(aBall)
    If threshold Then
      If BallSpeed(aBall) < threshold Then Exit Sub
    End If
    Dim RealCOR, DesiredCOR, str, coef
    DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
    RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id) + 0.0001)
    coef = desiredcor / realcor
    If debugOn Then str = name & " in vel:" & Round(cor.ballvel(aBall.id),2 ) & vbNewLine & "desired cor: " & Round(desiredcor,4) & vbNewLine & _
    "actual cor: " & Round(realCOR,4) & vbNewLine & "ballspeed coef: " & Round(coef, 3) & vbNewLine
    If Print Then Debug.print Round(cor.ballvel(aBall.id),2) & ", " & Round(desiredcor,3)

    aBall.velx = aBall.velx * coef
    aBall.vely = aBall.vely * coef
    If debugOn Then TBPout.text = str
  End Sub

  Public Sub Dampenf(aBall, parm) 'Rubberizer is handle here
    Dim RealCOR, DesiredCOR, str, coef
    DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
    RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id) + 0.0001)
    coef = desiredcor / realcor
    If Abs(aball.velx) < 2 And aball.vely < 0 And aball.vely >  - 3.75 Then
      aBall.velx = aBall.velx * coef
      aBall.vely = aBall.vely * coef
    End If
  End Sub

  Public Sub CopyCoef(aObj, aCoef) 'alternative addpoints, copy with coef
    Dim x
    For x = 0 To UBound(aObj.ModIn)
      addpoint x, aObj.ModIn(x), aObj.ModOut(x) * aCoef
    Next
  End Sub

  Public Sub Report() 'debug, reports all coords in tbPL.text
    If Not debugOn Then Exit Sub
    Dim a1, a2
    a1 = ModIn
    a2 = ModOut
    Dim str, x
    For x = 0 To UBound(a1)
      str = str & x & ": " & Round(a1(x),4) & ", " & Round(a2(x),4) & vbNewLine
    Next
    TBPout.text = str
  End Sub
End Class

'******************************************************
'  TRACK ALL BALL VELOCITIES
'  FOR RUBBER DAMPENER AND DROP TARGETS
'******************************************************

Dim cor
Set cor = New CoRTracker

Class CoRTracker
  Public ballvel, ballvelx, ballvely

  Private Sub Class_Initialize
    ReDim ballvel(0)
    ReDim ballvelx(0)
    ReDim ballvely(0)
  End Sub

  Public Sub Update() 'tracks in-ball-velocity
    Dim str, b, AllBalls, highestID
    allBalls = getballs

    For Each b In allballs
      If b.id >= HighestID Then highestID = b.id
    Next

    If UBound(ballvel) < highestID Then ReDim ballvel(highestID)  'set bounds
    If UBound(ballvelx) < highestID Then ReDim ballvelx(highestID)  'set bounds
    If UBound(ballvely) < highestID Then ReDim ballvely(highestID)  'set bounds

    For Each b In allballs
      ballvel(b.id) = BallSpeed(b)
      ballvelx(b.id) = b.velx
      ballvely(b.id) = b.vely
    Next
  End Sub
End Class

' Note, cor.update must be called in a 10 ms timer. The example table uses the GameTimer for this purpose, but sometimes a dedicated timer call RDampen is used.
'
'Sub RDampen_Timer
' Cor.Update
'End Sub

'******************************************************
'****  END PHYSICS DAMPENERS
'******************************************************



'******************************************************
' ZSSC: SLINGSHOT CORRECTION FUNCTIONS by apophis
'******************************************************
' To add these slingshot corrections:
'  - On the table, add the endpoint primitives that define the two ends of the Slingshot
'  - Initialize the SlingshotCorrection objects in InitSlingCorrection
'  - Call the .VelocityCorrect methods from the respective _Slingshot event sub

Dim LS
Set LS = New SlingshotCorrection
Dim RS
Set RS = New SlingshotCorrection

InitSlingCorrection

Sub InitSlingCorrection
  LS.Object = LeftSlingshot
  LS.EndPoint1 = EndPoint1LS
  LS.EndPoint2 = EndPoint2LS

  RS.Object = RightSlingshot
  RS.EndPoint1 = EndPoint1RS
  RS.EndPoint2 = EndPoint2RS

  'Slingshot angle corrections (pt, BallPos in %, Angle in deg)
  ' These values are best guesses. Retune them if needed based on specific table research.
  AddSlingsPt 0, 0.00, - 4
  AddSlingsPt 1, 0.45, - 7
  AddSlingsPt 2, 0.48,  0
  AddSlingsPt 3, 0.52,  0
  AddSlingsPt 4, 0.55,  7
  AddSlingsPt 5, 1.00,  4
End Sub

Sub AddSlingsPt(idx, aX, aY)    'debugger wrapper for adjusting flipper script in-game
  Dim a
  a = Array(LS, RS)
  Dim x
  For Each x In a
    x.addpoint idx, aX, aY
  Next
End Sub

Class SlingshotCorrection
  Public DebugOn, Enabled
  Private Slingshot, SlingX1, SlingX2, SlingY1, SlingY2

  Public ModIn, ModOut

  Private Sub Class_Initialize
    ReDim ModIn(0)
    ReDim Modout(0)
    Enabled = True
  End Sub

  Public Property Let Object(aInput)
    Set Slingshot = aInput
  End Property

  Public Property Let EndPoint1(aInput)
    SlingX1 = aInput.x
    SlingY1 = aInput.y
  End Property

  Public Property Let EndPoint2(aInput)
    SlingX2 = aInput.x
    SlingY2 = aInput.y
  End Property

  Public Sub AddPoint(aIdx, aX, aY)
    ShuffleArrays ModIn, ModOut, 1
    ModIn(aIDX) = aX
    ModOut(aIDX) = aY
    ShuffleArrays ModIn, ModOut, 0
    If gametime > 100 Then Report
  End Sub

  Public Sub Report() 'debug, reports all coords in tbPL.text
    If Not debugOn Then Exit Sub
    Dim a1, a2
    a1 = ModIn
    a2 = ModOut
    Dim str, x
    For x = 0 To UBound(a1)
      str = str & x & ": " & Round(a1(x),4) & ", " & Round(a2(x),4) & vbNewLine
    Next
    TBPout.text = str
  End Sub


  Public Sub VelocityCorrect(aBall)
    Dim BallPos, XL, XR, YL, YR

    'Assign right and left end points
    If SlingX1 < SlingX2 Then
      XL = SlingX1
      YL = SlingY1
      XR = SlingX2
      YR = SlingY2
    Else
      XL = SlingX2
      YL = SlingY2
      XR = SlingX1
      YR = SlingY1
    End If

    'Find BallPos = % on Slingshot
    If Not IsEmpty(aBall.id) Then
      If Abs(XR - XL) > Abs(YR - YL) Then
        BallPos = PSlope(aBall.x, XL, 0, XR, 1)
      Else
        BallPos = PSlope(aBall.y, YL, 0, YR, 1)
      End If
      If BallPos < 0 Then BallPos = 0
      If BallPos > 1 Then BallPos = 1
    End If

    'Velocity angle correction
    If Not IsEmpty(ModIn(0) ) Then
      Dim Angle, RotVxVy
      Angle = LinearEnvelope(BallPos, ModIn, ModOut)
      '   debug.print " BallPos=" & BallPos &" Angle=" & Angle
      '   debug.print " BEFORE: aBall.Velx=" & aBall.Velx &" aBall.Vely" & aBall.Vely
      RotVxVy = RotPoint(aBall.Velx,aBall.Vely,Angle)
      If Enabled Then aBall.Velx = RotVxVy(0)
      If Enabled Then aBall.Vely = RotVxVy(1)
      '   debug.print " AFTER: aBall.Velx=" & aBall.Velx &" aBall.Vely" & aBall.Vely
      '   debug.print " "
    End If
  End Sub
End Class






'***************************************************************
' ZSHA: VPW DYNAMIC BALL SHADOWS by Iakki, Apophis, and Wylte
'***************************************************************


'Ambient (Room light source)
Const AmbientBSFactor = 0.9  '0 to 1, higher is darker
Const AmbientMovement = 1    '1+ higher means more movement as the ball moves left and right
Const offsetX = 0        'Offset x position under ball (These are if you want to change where the "room" light is for calculating the shadow position,)
Const offsetY = 5        'Offset y position under ball (^^for example 5,5 if the light is in the back left corner)

'Dynamic (Table light sources)
Const DynamicBSFactor = 0.99  '0 to 1, higher is darker
Const Wideness = 20      'Sets how wide the dynamic ball shadows can get (20 +5 thinness is technically most accurate for lights at z ~25 hitting a 50 unit ball)
Const Thinness = 5        'Sets minimum as ball moves away from source

' *** Trim or extend these to match the number of balls/primitives/flashers on the table!  (will throw errors if there aren't enough objects)
Dim objrtx1(1), objrtx2(1)
Dim objBallShadow(1)
Dim OnPF(1)
Dim BallShadowA
BallShadowA = Array (BallShadowA0,BallShadowA1)
Dim DSSources(30), numberofsources', DSGISide(30) 'Adapted for TZ with GI left / GI right

'' *** The Shadow Dictionary
'Dim bsDict
'Set bsDict = New cvpmDictionary
'Const bsNone = "None"
'Const bsWire = "Wire"
'Const bsRamp = "Ramp"
'Const bsRampClear = "Clear"

'Initialization
DynamicBSInit

Sub DynamicBSInit()
  Dim iii, source

  'Prepare the shadow objects before play begins
  For iii = 0 To tnob - 1
    Set objrtx1(iii) = Eval("RtxBallShadow" & iii)
    objrtx1(iii).material = "RtxBallShadow" & iii
    objrtx1(iii).z = 1 + iii / 1000 + 0.01  'Separate z for layering without clipping
    objrtx1(iii).visible = 0

    Set objrtx2(iii) = Eval("RtxBall2Shadow" & iii)
    objrtx2(iii).material = "RtxBallShadow2_" & iii
    objrtx2(iii).z = 1 + iii / 1000 + 0.02
    objrtx2(iii).visible = 0

    Set objBallShadow(iii) = Eval("BallShadow" & iii)
    objBallShadow(iii).material = "BallShadow" & iii
    UpdateMaterial objBallShadow(iii).material,1,0,0,0,0,0,AmbientBSFactor,RGB(0,0,0),0,0,False,True,0,0,0,0
    objBallShadow(iii).Z = 1 + iii / 1000 + 0.04
    objBallShadow(iii).visible = 0

    BallShadowA(iii).Opacity = 100 * AmbientBSFactor
    BallShadowA(iii).visible = 0
  Next

  iii = 0

  For Each Source In DynamicSources
    DSSources(iii) = Array(Source.x, Source.y)
    '   If Instr(Source.name , "Left") > 0 Then DSGISide(iii) = 0 Else DSGISide(iii) = 1  'Adapted for TZ with GI left / GI right
    iii = iii + 1
  Next
  numberofsources = iii
End Sub

'Sub BallOnPlayfieldNow(onPlayfield, ballNum) 'Whether a ball is currently on the playfield. Only update certain things once, save some cycles
' If onPlayfield Then
'   OnPF(ballNum) = True
'   bsRampOff BOT(ballNum).ID
'   '   debug.print "Back on PF"
'   UpdateMaterial objBallShadow(ballNum).material,1,0,0,0,0,0,AmbientBSFactor,RGB(0,0,0),0,0,False,True,0,0,0,0
'   objBallShadow(ballNum).size_x = 5
'   objBallShadow(ballNum).size_y = 4.5
'   objBallShadow(ballNum).visible = 1
'   BallShadowA(ballNum).visible = 0
'   BallShadowA(ballNum).Opacity = 100 * AmbientBSFactor
' Else
'   OnPF(ballNum) = False
'   '   debug.print "Leaving PF"
' End If
'End Sub

Sub DynamicBSUpdate
  Dim falloff 'Max distance to light sources, can be changed dynamically if you have a reason
  falloff = 150
  Dim ShadowOpacity1, ShadowOpacity2
  Dim s, LSd, iii
  Dim dist1, dist2, src1, src2
  Dim bsRampType
  Dim BOT: BOT=getballs 'Uncomment if you're destroying balls - Not recommended! #SaveTheBalls

  'Hide shadow of deleted balls
  For s = UBound(BOT) + 1 To tnob - 1
    objrtx1(s).visible = 0
    objrtx2(s).visible = 0
    objBallShadow(s).visible = 0
    BallShadowA(s).visible = 0
  Next

  If UBound(BOT) < lob Then Exit Sub 'No balls in play, exit

  'The Magic happens now
  For s = lob To UBound(BOT)
    ' *** Normal "ambient light" ball shadow
    'Layered from top to bottom. If you had an upper pf at for example 80 units and ramps even above that, your Elseif segments would be z>110; z<=110 And z>100; z<=100 And z>30; z<=30 And z>20; Else (under 20)

    'Primitive shadow on playfield, flasher shadow in ramps
    If AmbientBallShadowOn = 1 Then

      '** Above the playfield
      If BOT(s).Z > 30 Then
''        If OnPF(s) Then BallOnPlayfieldNow False, s   'One-time update
'       bsRampType = getBsRampType(BOT(s).id)
'       '   debug.print bsRampType
'
'       If Not bsRampType = bsRamp Then 'Primitive visible on PF
'         objBallShadow(s).visible = 1
'         objBallShadow(s).X = (s).X + (BOT(s).X - (tablewidth / 2)) / (Ballsize / AmbientMovement) + offsetX
'         objBallShadow(s).Y = BOT(s).Y + offsetY
'         objBallShadow(s).size_x = 5 * ((BOT(s).Z + BallSize) / 80) 'Shadow gets larger and more diffuse as it moves up
'         objBallShadow(s).size_y = 4.5 * ((BOT(s).Z + BallSize) / 80)
'         UpdateMaterial objBallShadow(s).material,1,0,0,0,0,0,AmbientBSFactor * (30 / (BOT(s).Z)),RGB(0,0,0),0,0,False,True,0,0,0,0
'       Else 'Opaque, no primitive below
'         objBallShadow(s).visible = 0
'       End If
'
'       If bsRampType = bsRampClear Or bsRampType = bsRamp Then 'Flasher visible on opaque ramp
'         BallShadowA(s).visible = 1
'         BallShadowA(s).X = BOT(s).X + offsetX
'         BallShadowA(s).Y = BOT(s).Y + offsetY + BallSize / 10
'         BallShadowA(s).height = BOT(s).z - BallSize / 4 + s / 1000 'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping the ramp
'         If bsRampType = bsRampClear Then BallShadowA(s).Opacity = 50 * AmbientBSFactor
'       ElseIf bsRampType = bsWire Or bsRampType = bsNone Then 'Turn it off on wires or falling out of a ramp
'         BallShadowA(s).visible = 0
'       End If
'
        '** On pf, primitive only
      ElseIf BOT(s).Z <= 30 And BOT(s).Z > 20 Then
        objBallShadow(s).visible = 1
'       If Not OnPF(s) Then BallOnPlayfieldNow True, s
        objBallShadow(s).X = BOT(s).X + (BOT(s).X - (tablewidth / 2)) / (Ballsize / AmbientMovement) + offsetX
        objBallShadow(s).Y = BOT(s).Y + offsetY
        '   objBallShadow(s).Z = BOT(s).Z + s/1000 + 0.04   'Uncomment (and adjust If/Elseif height logic) if you want the primitive shadow on an upper/split pf

        '** Under pf, flasher shadow only
      Else
'       If OnPF(s) Then BallOnPlayfieldNow False, s
        objBallShadow(s).visible = 0
        BallShadowA(s).visible = 1
        BallShadowA(s).X = BOT(s).X + offsetX
        BallShadowA(s).Y = BOT(s).Y + offsetY
        BallShadowA(s).height = BOT(s).z - BallSize / 4 + s / 1000
      End If

      'Flasher shadow everywhere
    ElseIf AmbientBallShadowOn = 2 Then
      If BOT(s).Z > 30 Then 'In a ramp
        BallShadowA(s).X = BOT(s).X + offsetX
        BallShadowA(s).Y = BOT(s).Y + offsetY + BallSize / 10
        BallShadowA(s).height = BOT(s).z - BallSize / 4 + s / 1000 'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping the ramp
      ElseIf BOT(s).Z <= 30 And BOT(s).Z > 20 Then 'On pf
        BallShadowA(s).visible = 1
        BallShadowA(s).X = BOT(s).X + (BOT(s).X - (tablewidth / 2)) / (Ballsize / AmbientMovement) + offsetX
        BallShadowA(s).Y = BOT(s).Y + offsetY
        BallShadowA(s).height = 1.04 + s / 1000
      Else 'Under pf
        BallShadowA(s).X = BOT(s).X + offsetX
        BallShadowA(s).Y = BOT(s).Y + offsetY
        BallShadowA(s).height = BOT(s).z - BallSize / 4 + s / 1000
      End If
    End If

    ' *** Dynamic shadows
    If DynamicBallShadowsOn Then
      If BOT(s).Z < 30 And BOT(s).X < 780 Then 'Parameters for where the shadows can show, here they are not visible above the table (no upper pf) or in the plunger lane
        dist1 = falloff
        dist2 = falloff

        For iii = 0 To numberofsources - 1 'Search the 2 nearest influencing lights
          LSd = Distance(BOT(s).x, BOT(s).y, DSSources(iii)(0), DSSources(iii)(1)) 'Calculating the Linear distance to the Source
          If LSd < falloff Then'And gilvl > 0 Then
            '   If LSd < dist2 And ((DSGISide(iii) = 0 And Lampz.State(100)>0) Or (DSGISide(iii) = 1 And Lampz.State(104)>0)) Then  'Adapted for TZ with GI left / GI right
            dist2 = dist1
            dist1 = LSd
            src2 = src1
            src1 = iii
          End If
        Next
        ShadowOpacity1 = 0
        If dist1 < falloff Then
          objrtx1(s).visible = 1
          objrtx1(s).X = BOT(s).X
          objrtx1(s).Y = BOT(s).Y
          '   objrtx1(s).Z = BOT(s).Z - 25 + s/1000 + 0.01 'Uncomment if you want to add shadows to an upper/lower pf
          objrtx1(s).rotz = AnglePP(DSSources(src1)(0), DSSources(src1)(1), BOT(s).X, BOT(s).Y) + 90
          ShadowOpacity1 = 1 - dist1 / falloff
          objrtx1(s).size_y = Wideness * ShadowOpacity1 + Thinness
          UpdateMaterial objrtx1(s).material,1,0,0,0,0,0,ShadowOpacity1 * DynamicBSFactor ^ 3,RGB(0,0,0),0,0,False,True,0,0,0,0
        Else
          objrtx1(s).visible = 0
        End If
        ShadowOpacity2 = 0
        If dist2 < falloff Then
          objrtx2(s).visible = 1
          objrtx2(s).X = BOT(s).X
          objrtx2(s).Y = BOT(s).Y + offsetY
          '   objrtx2(s).Z = BOT(s).Z - 25 + s/1000 + 0.02 'Uncomment if you want to add shadows to an upper/lower pf
          objrtx2(s).rotz = AnglePP(DSSources(src2)(0), DSSources(src2)(1), BOT(s).X, BOT(s).Y) + 90
          ShadowOpacity2 = 1 - dist2 / falloff
          objrtx2(s).size_y = Wideness * ShadowOpacity2 + Thinness
          UpdateMaterial objrtx2(s).material,1,0,0,0,0,0,ShadowOpacity2 * DynamicBSFactor ^ 3,RGB(0,0,0),0,0,False,True,0,0,0,0
        Else
          objrtx2(s).visible = 0
        End If
        If AmbientBallShadowOn = 1 Then
          'Fades the ambient shadow (primitive only) when it's close to a light
          UpdateMaterial objBallShadow(s).material,1,0,0,0,0,0,AmbientBSFactor * (1 - max(ShadowOpacity1, ShadowOpacity2)),RGB(0,0,0),0,0,False,True,0,0,0,0
        Else
          BallShadowA(s).Opacity = 100 * AmbientBSFactor * (1 - max(ShadowOpacity1, ShadowOpacity2))
        End If
      Else 'Hide dynamic shadows everywhere else, just in case
        objrtx2(s).visible = 0
        objrtx1(s).visible = 0
      End If
    End If
  Next
End Sub


'****************************************************************
'****  END VPW DYNAMIC BALL SHADOWS by Iakki, Apophis, and Wylte
'****************************************************************







'******************************************************
'  Timers
'******************************************************


Sub FrameTimer_Timer  'This is a timer with -1 interval. If it already exists in your table, then use the existing timer.
  MovableHelper
  If DynamicBallShadowsOn Or AmbientBallShadowOn Then DynamicBSUpdate 'update ball shadows
End Sub

Sub GameTimer_Timer  'This is a timer with 10 ms interval. If it already exists in your table, then use the existing timer.
  Cor.Update
  DoSTAnim
  STMovableHelper
End Sub


Sub ChangeRoomBrightness(level)
  ' Lighting level
  If level>1 Then level=1
  If level<0 Then level=0
  dim v: v = Int(level * 200 + 55)
  Dim x: For Each x in BM_World: x.Color = RGB(v, v, v): Next
End Sub



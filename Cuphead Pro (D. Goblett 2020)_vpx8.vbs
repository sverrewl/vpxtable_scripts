'*
'*      Cuphead Pro (D. Goblett & Co 2020)
'*    By Onevox, Scottacus and bord. Core script by Loserman76
'*      "The Bees Knees" for helping the dev: Loserman76 (base table and script), BorgDog (playfield elements), Thalamus (SSF), Xenonph (Music), cyberpez (enhancement)
'*      Adapted from the "Cuphead" videogame by StudioMDHR



option explicit
Randomize

ExecuteGlobal GetTextFile("core.vbs")

On Error Resume Next
ExecuteGlobal GetTextFile("Controller.vbs")
If Err Then MsgBox "Unable to open Controller.vbs. Ensure that it is in the scripts folder."
On Error Goto 0

Const cGameName = "cuphead"

Const ShadowFlippersOn = true
Const ShadowBallOn = true

Const ShadowConfigFile = false

Dim Controller  'B2S
Dim B2SScore  ' B2S Score Displayed
Const HSFileName="CupheadPro.txt"
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

'* This controls whether music is played (1) or not (0) during game time
Dim MusicOn
MusicOn = 1

'* This sets the music volume for all songs
Dim MusicVolume
MusicVolume = 1

'* Set to 1 for Free Play or 0 for Coin PlayChime
Dim FreePlay
FreePlay = 1

'* Set to 1 for Double Bonus on last ball 0
Dim LastBallDoubleBonus
LastBallDoubleBonus = 1

'* SoulLight Shifting Set to 0 for flippers and 1 for CupHead/Mugman Rubbers
Dim ShiftControl
ShiftControl = 0

'* Colored Balls For Multiball
Dim ColoredBall
ColoredBall = 1

'* Flipper Length Set to '2' for 2" Flippers and '3' for 3" Flippers
Dim FlipperLength
FlipperLength = 2

'* SOUL light reset - Set to 1 to keep SOUL lights lit from ball to ball
Dim SoulLightReset
SoulLightReset = 1
Dim SoulLightArray(4,4)

'* Plunger Sounds - Set to 1 to enable mechanical plunger sounds or 0 to not enable these sounds
Dim PlungerSound
PlungerSound = 1

dim ScoreChecker
dim CheckAllScores
dim sortscores(4)
dim sortplayers(4)
Dim TextStr,TextStr2
Dim i,xx,LStep,RStep

Dim obj
Dim bgpos
Dim dooralreadyopen
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
Dim Replay1Paid(4)
Dim Replay2Paid(4)
Dim Replay3Paid(4)
Dim TableTilted
Dim TiltCount
Dim OperatorMenu

Dim ExtraBallSetting
Dim BonusBooster
Dim BonusBoosterCounter
Dim BonusCounter
Dim HoleCounter

Dim AdvanceLightCounter
Dim ExtraBall

Dim Ones
Dim Tens
Dim Hundreds
Dim Thousands

Dim Player
Dim Players

Dim AlternateRelay

'Dim CheckRollovers
'Dim CheckMultiplierBonus

Dim LightSequenceCounter
Dim LightSequence2Counter
Dim SoulBonusCounter
Dim BonusMultiplierCounter
Dim ScoreMotorStepper

Dim rst
Dim bonuscountdown
Dim TempMultiCounter
dim TempPlayerup
dim RotatorTemp

Dim bump1
Dim bump2
Dim bump3

Dim LastChime10
Dim LastChime100
Dim LastChime1000

Dim Score10
Dim Score100

Dim MotorRunning
Dim Replay1Table(15)
Dim Replay2Table(15)
Dim Replay3Table(15)
Dim ReplayTableSet
Dim ReplayLevel
Dim ReplayTableMax
Dim ScoreMotorClicks


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
  HighScore=0
  MotorRunning=0
  HighScoreReward=3
  BallsPerGame=5
  BonusBooster=0
  ReplayLevel=1
  ExtraBallSetting=1
  Credits=0
  loadhs
  if HighScore=0 then HighScore=50000

  Replay1=80000
  Replay2=125000
  Replay3=140000
  TableTilted=false

  Match=int(Rnd*10)*10
  MatchReel.SetValue((Match/10)+1)
  GameOverReel.SetValue(1)
  TiltReel.SetValue(1)
  CanPlayReel.SetValue(0)
  BallInPlayReel.SetValue(0)

' For each obj in PlayerHuds
'   obj.SetValue(0)
' next
  For each obj in PlayerScoresOn
    obj.ResetToZero
  next
  For each obj in PlayerScores
    obj.ResetToZero
  next
  for each obj in StarLights
    obj.state=0
  next
  for each obj in BonusX
    obj.state=0
  next
  for each obj in TopRolloverLights
    obj.state=0
  next
' for each obj in LeftKickerLights
'   obj.state=0
' next
' for each obj in RightKickerLights
'   obj.state=0
' next
  For x = 1 to 4
    PerditionEntry(x) = 1
  Next
  FireLight1.state=0
  FireLight2.state=0
  FireLight3.state=0
  LightSequenceCounter=0
  LightSequence2Counter=0
  SnakeEyesLight.state=2

  If FlipperLength = 3 Then
    For each obj in inch2
      obj.visible = False
      FlipperLSh.visible = False
      FlipperRSh.visible = False
    next
    For each obj in inch2col
      obj.collidable = False
    next
    For each obj in inch3
      obj.visible = True
      FlipperLSh001.visible = True
      FlipperRSh001.visible = True
    next
    For each obj in inch3col
      obj.collidable = True
    next
    LeftFlipper.Enabled = 0
    RightFlipper.Enabled = 0
    LeftFlipper001.Enabled = 1
    RightFlipper001.Enabled = 1
  Else
    For each obj in inch2
      obj.visible = True
      FlipperLSh.visible = True
      FlipperRSh.visible = True
    next
    For each obj in inch2col
      obj.collidable = True
    next
    For each obj in inch3
      obj.visible = False
      FlipperLSh001.visible = False
      FlipperRSh001.visible = False
    next
    For each obj in inch3col
      obj.collidable = False
    next
    LeftFlipper.Enabled = 1
    RightFlipper.Enabled = 1
    LeftFlipper001.Enabled = 0
    RightFlipper001.Enabled = 0
    End If

  Replay1=Replay1Table(ReplayLevel)
  Replay2=Replay2Table(ReplayLevel)
  Replay3=Replay3Table(ReplayLevel)

  SoulBonusCounter=0
  BonusCounter=0
  HoleCounter=0
    bgpos=6
  AdvanceLightCounter=0

  ShadowMain.visible=False
  FlipperLSh.visible=False
  FlipperRSh.visible=False
  FlipperLSh001.visible=False
  FlipperRSh001.visible=False

  for each obj in ShadowCup
    obj.visible=False
  Next
  for each obj in ShadowMug
    obj.visible=False
  Next

  dooralreadyopen=0
  InstructCard.image="IC_"+FormatNumber(BallsPerGame,0)

  RefreshReplayCard

' ResetBalls
' PlaySound "Plunger"

  TargetSpecialLit = 0
  Points210counter=0
  Points500counter=0
  Points1000counter=0
  Points2000counter=0

  BonusBoosterCounter=0
  Players=0
  RotatorTemp=1
  InProgress=false
  ExtraBall=false
  AlternateRelay=1


'***************************

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
    'Controller.B2SSetScore 6,HighScore
    Controller.B2SSetTilt 0
    Controller.B2SSetCredits Credits
    Controller.B2SSetGameOver 1
    Controller.B2SSetShootAgain 0
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
'   PlungerPulled = 1
    PlaySound"CHPlungerpull"

  End If

  if keycode = LeftFlipperKey and InProgress = false then
    OperatorMenuTimer.Enabled = true
  end if
  ' END GNMOD

  If keycode = LeftFlipperKey and InProgress=true and TableTilted=false and BotchedMB = 0 and contball = 0 Then
    LFPress = 1
    If FlipperLength = 2 Then
      lf.fire
    Else
      lf1.fire
    End If
    PlaySoundAt SoundFXDOF("Flipper",101,DOFOn,DOFFlippers), LeftFlipper
    PlayLoopSoundAtVol "buzzL", LeftFlipper, 1
    If ShiftControl = 0 Then SoulLightsLeft
  End If

  If keycode = RightFlipperKey and InProgress=true and TableTilted=false and BotchedMB = 0 and contball = 0 Then
    RFPress = 1
    If FlipperLength = 2 Then
      rf.fire
    Else
      rf1.fire
    End If
    PlaySoundAt SoundFXDOF("Flipper",102,DOFOn,DOFFlippers), RightFlipper
    PlayLoopSoundAtVol "buzz", RightFlipper, 1
    If ShiftControl = 0 Then SoulLightsRight
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

    PlaySoundAt "coinin", Drain
    playsound "CHCoinin"
    AddSpecial2
  end if

   if keycode = 5 then

    PlaySoundAt "coinin", Drain
    playsound "CHCoinin"
    AddSpecial2
    keycode= StartGameKey
  end if


   if keycode = StartGameKey and InProgress=true and Players>0 and Players<4 and BallInPlay<2 then
    If FreePlay = 0 and Credits <1 Then Exit Sub
    If Credits > 0 Then Credits = Credits - 1
    If Credits < 1 Then DOF 126, DOFOff
    CreditsReel.SetValue(Credits)
    Players=Players+1
    CanPlayReel.SetValue(Players)
    playsound "click":PB=1:If MusicOn = 1 Then s01.enabled=true
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

  if keycode=StartGameKey and InProgress=false and Players=0 and EnteringOptions = 0 then
'GNMOD
    OperatorMenuTimer.Enabled = false
'END GNMOD
    If FreePlay = 0 and Credits <1 Then Exit Sub
    If Credits > 0 Then Credits = Credits - 1
    If Credits < 1 Then DOF 126, DOFOff
    CreditsReel.SetValue(Credits)
    Players=1
    CanPlayReel.SetValue(Players)
    MatchReel.SetValue(0)
    GameOverReel.SetValue(0)
    Player=1
    playsound "StartUpSequence":EndMusic:If MusicOn = 1 Then s01.enabled=True:AA=0
    TempPlayerUp=Player
    PlayerUpRotator.enabled=true
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
      'Controller.B2SSetScore 6,HighScore
      Controller.B2SSetCanPlay 1
      Controller.B2SSetPlayerUp 1
      Controller.B2SSetBallInPlay BallInPlay
      Controller.B2SSetScoreRolloverPlayer1 0
    End If
    If Table1.ShowDT = True then
      For each obj in PlayerScores
        obj.ResetToZero
        obj.Visible=true
      next
      For each obj in PlayerScoresOn
        obj.ResetToZero
        obj.Visible=false
      next

'     For each obj in PlayerHuds
'       obj.SetValue(0)
'     next
'     For each obj in PlayerHUDScores
'       obj.state=0
'     next
'     PlayerHuds(Player-1).SetValue(1)
'     PlayerHUDScores(Player-1).state=1
      PlayerScores(Player-1).Visible=0
      PlayerScoresOn(Player-1).Visible=1
    end If
  end if

    If keycode = 46 Then' C Key
        If contball = 1 Then
            contball = 0
          Else
            contball = 1
        End If
    End If

    If keycode = 48 Then 'B Key
        If bcboost = 1 Then
            bcboost = bcboostmulti
          Else
            bcboost = 1
        End If
    End If

    If keycode = 203 Then cLeft = 1' Left Arrow

    If keycode = 200 Then cUp = 1' Up Arrow

    If keycode = 208 Then cDown = 1' Down Arrow

    If keycode = 205 Then cRight = 1' Right Arrow

    If keycode = 52 Then Zup = 1' Period

  If Keycode = 30 Then 'a' key
    If FlipperLength = 2 Then
      FlipperLength = 3
      For each obj in inch2
        obj.visible = False
        FlipperLSh.visible = False
        FlipperRSh.visible = False
      next
      For each obj in inch2col
        obj.collidable = False
      next
      For each obj in inch3
        obj.visible = True
        FlipperLSh001.visible = True
        FlipperRSh001.visible = True
      next
      For each obj in inch3col
        obj.collidable = True
      next
      LeftFlipper.Enabled = 0
      RightFlipper.Enabled = 0
      LeftFlipper001.Enabled = 1
      RightFlipper001.Enabled = 1
    Else
      FlipperLength = 2
      For each obj in inch2
        obj.visible = True
        FlipperLSh.visible = True
        FlipperRSh.visible = True
      next
      For each obj in inch2col
        obj.collidable = True
      next
      For each obj in inch3
        obj.visible = False
        FlipperLSh001.visible = False
        FlipperRSh001.visible = False
      next
      For each obj in inch3col
        obj.collidable = False
      next
      LeftFlipper.Enabled = 1
      RightFlipper.Enabled = 1
      LeftFlipper001.Enabled = 0
      RightFlipper001.Enabled = 0
    End If

  End If

  If Keycode = 31 Then 's' key

    GIReset
    ColorChange = ColorChange + 1
    Tb.text = "Color" & ColorChange
    If ColorChange = 1 Then
      For each obj in Flashers
        obj.color=RGB(255, 0, 0)
        obj.colorfull = RGB(255, 0, 0)
      next
    Else
      ColorChange = 0
      For each obj in Flashers
        obj.color=RGB(0, 0, 255)
        obj.colorfull=RGB(0, 0, 255)
      next
    End If

  End If

End Sub

Dim ColorChange

Sub Table1_KeyUp(ByVal keycode)

  ' GNMOD
  if EnteringInitials then
    exit sub
  end if

  If keycode = PlungerKey Then

'   if PlungerPulled = 0 then
'     exit sub
'   end if
  ' PlaySoundAt "plungerrelease", Plunger
    PlaySoundAt "CHPlungerRelease", Plunger
    Plunger.Fire
  End If

  if keycode = LeftFlipperKey then
    OperatorMenuTimer.Enabled = false
  end if

  ' END GNMOD

  If keycode = LeftFlipperKey and InProgress=true and TableTilted=false Then
    lfpress = 0
    LeftFlipper.eosTorqueAngle = EOSA
    LeftFlipper.eosTorque = EOST
    LeftFlipper.RotateToStart
    LeftFlipper001.RotateToStart
    PlaySoundAt SoundFXDOF("FlipperDown",101,DOFOff,DOFFlippers), LeftFlipper
    StopSound "buzzL"
  End If

  If keycode = RightFlipperKey and InProgress=true and TableTilted=false Then
    rfpress = 0
    RightFlipper.eosTorqueAngle = EOSA
    RightFlipper.eosTorque = EOST
    RightFlipper.RotateToStart
    RightFlipper001.RotateToStart
    PlaySoundAt SoundFXDOF("FlipperDown",102,DOFOff,DOFFlippers), RightFlipper
    StopSound "buzz"
  End If

    If keycode = 203 then cLeft = 0' Left Arrow

    If keycode = 200 then cUp = 0' Up Arrow

    If keycode = 208 then cDown = 0' Down Arrow

    If keycode = 205 then cRight = 0' Right Arrow

    If keycode = 52 Then Zup = 0' Period

End Sub

Dim BotchedMB
Sub Drain_Hit()
  FireLight1.state=0
  FireLight2.state=0
  FireLight3.state=0
  PlaySoundAt SoundFXDOF("fx_drain",122,DOFPulse,DOFContactors), Drain
  If MultiballOn = 1 Then
    EndMusic
    Drain.DestroyBall
    BotchedMB = 1
    MultiballOn = 0
    PlaySound "CHPerditionKicker"
    PlasticsOff
    BumpersOff
    for each obj in StarLights
      obj.state=0
    next
    bumpercap2.image="bcap-100wl-unlit devilBW"
    LeftFlipper.RotateToStart
    RightFlipper.RotateToStart
    Perditionpost.transz= -37
    PerditionWall.collidable = 0
    dim  BOT, b
    BOT = GetBalls
    For b = 0 to ubound(BOT)
      If BOT(b).x  > 790 and BOT(b).y < 800 Then Set KickerBall1 = BOT(b)
    Next
    Pkickarm1.rotz=15
    KickerArm1.enabled=true
    kickBall KickerBall1, 0, 45, 5, 30
    PlaySoundAt SoundFXDOF("saucer",131,DOFPulse,DOFcontactors), Pkickarm1
    Exit Sub
  End If
  If MultiballReleased = 1 Then
    MultiballReleased = 0
    PerditionWall.collidable = 0
    perditionpost.transz = -37
    Drain.DestroyBall
    GIReset
    Exit Sub
  End If
  EndMusic:PlaySound "CHDrain":AA=1:PB=0:If MusicOn = 1 Then m01.enabled=false:m02.enabled=false
  Drain.DestroyBall
  For x = 1 to 3
    EVAL("PerditionLight" & x).state = 0
  Next
  SnakeEyesLight.state=2
  Pause4Bonustimer.enabled=1
  BotchedMB = 0
End Sub

'************************Music Timers and Triggers

Sub BallReleaseGate_Hit()
    PB=1
    If AA=1 Then
    If MusicOn = 1 Then m01.enabled=True
    End If
End Sub

Sub m01_Timer
    Dim za
    za = INT(10 * RND(1) )
    Select Case za
    Case 0:PlayMusic"CupMusic1.mp3", MusicVolume:m02.enabled=true:m02.interval=239000
    Case 1:PlayMusic"CupMusic2.mp3", MusicVolume:m02.enabled=true:m02.interval=230700
    Case 2:PlayMusic"CupMusic3.mp3", MusicVolume:m02.enabled=true:m02.interval=204800
    Case 3:PlayMusic"CupMusic4.mp3", MusicVolume:m02.enabled=true:m02.interval=223300
    Case 4:PlayMusic"CupMusic5.mp3", MusicVolume:m02.enabled=true:m02.interval=240200
    Case 5:PlayMusic"CupMusic6.mp3", MusicVolume:m02.enabled=true:m02.interval=239200
    Case 6:PlayMusic"CupMusic7.mp3", MusicVolume:m02.enabled=true:m02.interval=226000
    Case 7:PlayMusic"CupMusic8.mp3", MusicVolume:m02.enabled=true:m02.interval=200500
    Case 8:PlayMusic"CupMusic9.mp3", MusicVolume:m02.enabled=true:m02.interval=257400
    Case 9:PlayMusic"CupMusic10.mp3", MusicVolume:m02.enabled=true:m02.interval=252800
    End Select
m01.enabled=false
End sub

Sub m02_Timer
    Dim zb
    zb = INT(10 * RND(1) )
    Select Case zb
    Case 0:PlayMusic"CupMusic1.mp3", MusicVolume:m01.enabled=true:m01.interval=239000
    Case 1:PlayMusic"CupMusic2.mp3", MusicVolume:m01.enabled=true:m01.interval=230700
    Case 2:PlayMusic"CupMusic3.mp3", MusicVolume:m01.enabled=true:m01.interval=204800
    Case 3:PlayMusic"CupMusic4.mp3", MusicVolume:m01.enabled=true:m01.interval=223300
    Case 4:PlayMusic"CupMusic5.mp3", MusicVolume:m01.enabled=true:m01.interval=240200
    Case 5:PlayMusic"CupMusic6.mp3", MusicVolume:m01.enabled=true:m01.interval=239200
    Case 6:PlayMusic"CupMusic7.mp3", MusicVolume:m01.enabled=true:m01.interval=226000
    Case 7:PlayMusic"CupMusic8.mp3", MusicVolume:m01.enabled=true:m01.interval=200500
    Case 8:PlayMusic"CupMusic9.mp3", MusicVolume:m01.enabled=true:m01.interval=257400
    Case 9:PlayMusic"CupMusic10.mp3", MusicVolume:m01.enabled=true:m01.interval=252800
    End Select
m02.enabled=false
End sub

Sub s01_Timer
    If PB=0 Then
     Dim zc
     zc = INT(4 * RND(1) )
     Select Case zc
     Case 0:PlaySound"CHballstart1":m01.enabled=true:m01.interval=2000
     Case 1:PlaySound"CHballstart2":m01.enabled=true:m01.interval=1818
     Case 2:PlaySound"CHballstart3":m01.enabled=true:m01.interval=1824
     Case 3:PlaySound"CHballstart4":m01.enabled=true:m01.interval=1704
     End Select
     s01.enabled=false
     End If
    If PB=1 Then
     Dim zd
     zd = INT(4 * RND(1) )
     Select Case zd
     Case 0:PlaySound"CHballstart1"
     Case 1:PlaySound"CHballstart2"
     Case 2:PlaySound"CHballstart3"
     Case 3:PlaySound"CHballstart4"
     End Select
     s01.enabled=false
     End If
End sub

Dim AA
Dim PB
PB=0

'************************Bonus Timers

Sub Pause4Bonustimer_timer
  Pause4Bonustimer.enabled=0
  ScoreSoulBonus
End Sub

Sub NewBonusHolder_timer
  if NewBonusTimer.enabled=0 then
    NewBonusHolder.enabled=0
    NextBallDelay.enabled=true
  end if

end sub

'***********************
'     Flipper Logos
'***********************

Sub UpdateFlipperLogos_Timer
  LFlip.Rotz = LeftFlipper.CurrentAngle -121
  RFlip.Rotz = RightFlipper.CurrentAngle +121
  LFlipr.Rotz = LeftFlipper.CurrentAngle -121
  RFlipr.Rotz = RightFlipper.CurrentAngle +121
  PGate.Rotz = (Gate.CurrentAngle*.75) + 25
  FlipperLSh001.RotZ = LeftFlipper001.currentangle
  FlipperRSh001.RotZ = RightFlipper001.currentangle
  FlipperLSh.RotZ = LeftFlipper.currentangle
  FlipperRSh.RotZ = RightFlipper.currentangle
End Sub


'*********************** slingshots


Sub RightSlingShot_Slingshot
    PlaySoundAt SoundFXDOF("right_slingshot",104,DOFPulse,DOFContactors), sling1
    RSling0.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -11
    RStep = 0
    RightSlingShot.TimerEnabled = 1
  If RightSlingLight.state=1 Then
    AddScore(100)
  Else
    AddScore(10)
  End If
  LeftSlingLight.state = 1
  RightSlingLight.state = 0
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -1
        Case 4:RSLing2.Visible = 0:RSling0.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
    PlaySoundAt SoundFXDOF("left_slingshot",103,DOFPulse,DOFContactors), sling2
    LSling0.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -11
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
    If LeftSlingLight.state=1 Then
    AddScore(100)
  Else
    AddScore(10)
  End If
  LeftSlingLight.state = 0
  RightSlingLight.state = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -1
        Case 4:LSLing2.Visible = 0:LSLing0.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

'************************************ POP Bumpers

Sub Bumper1_Hit
  If TableTilted=false then
    PlaySoundAt SoundFXDOF("bumper1",106,DOFPulse,DOFcontactors), Bumper1
    bump1 = 1
    If Bumper1Light.state = 1 Then
      AddScore(100)
    Else
      AddScore(10)
    End If

    end if

End Sub

Sub Bumper2_Hit
  If TableTilted=false then
    PlaySoundAt SoundFXDOF("bumper1",107,DOFPulse,DOFcontactors), Bumper2
    bump3 = 1
    If Bumper2Light.state = 1 Then
      AddScore(100)
    Else
      AddScore(10)
    End If
    AlternateRelayFire
    If MultiballOn = 1 Then
      bumpercap2.image="bcap-100wl-unlit devilBW"
      Perditionpost.transz= -37
      PerditionWall.collidable = 0
      For x = 1 to 3
        EVAL("PerditionLight" & x).state = 0
        EVAL("FireLight" & x).state = 0
      Next
      dim  BOT, b
      BOT = GetBalls
      For b = 0 to ubound(BOT)
        If BOT(b).x  > 790 and BOT(b).y < 800 Then Set KickerBall1 = BOT(b)
      Next
      Pkickarm1.rotz=15
      KickerArm1.enabled=true
      kickBall KickerBall1, 0, 45, 5, 30
      PostTimer.enabled = 1
      PlaySoundAt SoundFXDOF("saucer",131,DOFPulse,DOFcontactors), Pkickarm1
      PlaySound "CHPerditionKicker"
      MultiballReleased = 1
      MultiballOn = 0
      Bumper2Light.state = 1
      Bumper1Light.state = 1
      Bumper3Light.state = 1
      GIReset
    End If
    end if
End Sub

Sub PostTimer_Timer
  PerditionWall.collidable = 1
  perditionpost.transz = 0
  PostTimer.enabled = 0
End Sub

Sub Bumper3_Hit
  If TableTilted=false then
    PlaySoundAt SoundFXDOF("bumper1",108,DOFPulse,DOFcontactors), Bumper3
    bump2 = 1
    If Bumper3Light.state = 1 Then
      AddScore(100)
    Else
      AddScore(10)
    End If
    end if

End Sub

'**************************************** ROLLOVER TRIGGERS

Sub TriggerTopA_Hit()
  If TableTilted=false then
    PlaySoundAt "sensor", ActiveBall
    PlaySound "CHBounce"
    DOF 120, DOFPulse
    SetMotor(100)
    SoulLight1.state=1
    CheckRollovers
    Special7
    If BonusMultiplier < 4 Then SoulLightArray(Player,1) = 1
  end if
End Sub

Sub TriggerTopB_Hit()
  If TableTilted=false then
    DOF 119, DOFPulse
    PlaySoundAt "sensor", ActiveBall
    PlaySound "CHBounce"
    SetMotor(100)
    SoulLight2.state=1
    CheckRollovers
    Special7
    If BonusMultiplier < 4 Then SoulLightArray(Player,2) = 1
  end if
End Sub

Sub TriggerTopC_Hit()
  If TableTilted=false then
    PlaySoundAt "sensor", ActiveBall
    PlaySound "CHBounce"
    DOF 121, DOFPulse
    SetMotor(100)
    SoulLight3.state=1
    CheckRollovers
    Special7
    If BonusMultiplier < 4 Then SoulLightArray(Player,3) = 1
  end if
End Sub

Sub TriggerTopD_Hit()
  If TableTilted=false then
    PlaySoundAt "sensor", ActiveBall
    PlaySound "CHBounce"
    DOF 118, DOFPulse
    SetMotor(100)
    SoulLight4.state=1
    CheckRollovers
    Special7
    If BonusMultiplier < 4 Then SoulLightArray(Player,4) = 1
  end if
End Sub


Sub CheckRollovers
  if (SoulLight1.state=1) and (SoulLight2.state=1) and (SoulLight3.state=1) and (SoulLight4.state=1) then
    Addscore(1000)
    BonusMultiplier=BonusMultiplier+1
    For x = 1 to 4
      EVAL("SoulLight" & x).state = 0
      SoulLightArray(Player,x) = 0
    Next
  end If
  MultiplierBonus
end sub

Sub MultiplierBonus
  if BonusMultiplier=2 Then
    Bonus2X.state=1
  end If
  if BonusMultiplier=3 Then
    Bonus3X.state=1
    Bonus2X.state=0
  end If
  if BonusMultiplier=4 Then
    Bonus4X.state=1
    Bonus3X.state=0
    SoulLight1.state=1
    SoulLight2.state=1
    SoulLight3.state=1
    SoulLight4.state=1
  end If
  if BonusMultiplier>4 Then
    BonusMultiplier=4
    Bonus4X.state=1
    SoulLight1.state=1
    SoulLight2.state=1
    SoulLight3.state=1
    SoulLight4.state=1
  end If
End Sub

'********************************LOWER ROLLOVERS

Sub TriggerUpperRightRollover_Hit()
  If TableTilted=false then
    DOF 133, DOFPulse
    PlaySoundAt "sensor", ActiveBall
    PlaySound "CHRollover"
    AddScore(100)
    PerditionEntry(Player) = PerditionEntry(Player) + 1
    If PerditionEntry(Player) > 6 Then PerditionEntry(Player) = 1
    If PerditionEntry(Player) mod 2 = 0 Then EVAL("PerditionLight" & (PerditionEntry(Player)/2)).state = 1
  End if
  If PerditionEntry(Player) = 6 Then
    FireLight1.state=2
    FireLight2.state=2
    FireLight3.state=2
  End If
'  If PerditionCounter > 5 Then
'   FireLight1.state=1
'   FireLight2.state=0
'   FireLight3.state=0
'   PerditionCounter=0
'   for each obj in TopRolloverLights
'     obj.state=0
'   Next
'   for each obj in RolloverLightsLeft
'     obj.state=0
'   Next
'   for each obj in RolloverLightsRight
'     obj.state=0
'   Next
'   for each obj in StandUpsDash
'     obj.state=0
'   Next
'   for each obj in DashLights
'     obj.state=0
'   Next
'   for each obj in LoopLightsLeft
'     obj.state=0
'   Next
'   for each obj in LoopLightsRight
'     obj.state=0
'   Next
'   for each obj in StarLights
'     obj.state=1
'   Next
'   RightExtraBallLight.state=0
' End If
End Sub

Sub TriggerLowerRightRollover_Hit()
  If TableTilted=false then
    DOF 134, DOFPulse
    PlaySoundAt "sensor", ActiveBall
    PlaySound "CHRollover"
    SetMotor(50)
  End If
End Sub


Sub TriggerLeftInlane_Hit()
  If TableTilted=false then
    DOF 128, DOFPulse
    PlaySoundAt "sensor", ActiveBall
    PlaySound "CHInlanes"
    If LeftInlaneLight.state=1 then
      SetMotor(500)
      LeftInlaneLight.state=0
    Else
      SetMotor(50)
    End If
  End If
End Sub


Sub TriggerRightInlane_Hit()
  If TableTilted=false then
    DOF 129, DOFPulse
    PlaySoundAt "sensor", ActiveBall
    PlaySound "CHInlanes"
    If RightInlaneLight.state=1 then
      SetMotor(500)
      RightInlaneLight.state=0
    Else
      SetMotor(50)
    End If
  End If
End Sub

Sub TriggerLeftOutlane_Hit()
  If TableTilted=false then
    DOF 127, DOFPulse
    PlaySoundAt "sensor", ActiveBall
    PlaySound "CHOutlanes"
    If LeftOutlaneLight.state=1 then
      SetMotor(1000)
      LeftOutlaneLight.state=0
    Else
      SetMotor(100)
    End If
  End If
End Sub

Sub TriggerRightOutlane_Hit()
  If TableTilted=false then
    DOF 130, DOFPulse
    PlaySoundAt "sensor", ActiveBall
    PlaySound "CHOutlanes"
    If LeftOutlaneLight.state=1 then
      SetMotor(1000)
      RightOutlaneLight.state=0
    Else
      SetMotor(100)
    End If
  End If
End Sub

Sub Trigger1_Hit
  BotchedMB = 0
  If MultiballOn Then MugBallLight.state = 2
  Set controlBall = ActiveBall
    contBallInPlay = True
End Sub

Sub Trigger1_Unhit
  MugBallLight.state = 0
  Trigger1.timerEnabled = 0
  PlungerPlaying = 0
End Sub

'*******************************STAR ROLLOVERS

Sub UpperLeftStar_Hit()
  if TableTilted=false then
    DOF 135, 2
    If UpperLeftStarLight.state=2 then
      SetMotor(500)
      LeftOutlaneLight.state=1
      RightOutlaneLight.state=1
      PlaySound"CHStars"
    Else
      SetMotor(50)
    End If
  End If
end sub

Sub LowerLeftStar_Hit()
  if TableTilted=false then
    DOF 136, 2
    If LowerLeftStarLight.state=2 then
      SetMotor(500)
      PlaySound"CHStars"
    Else
      SetMotor(50)
    end if
  end if
  If SpecialLight.state=2 Then
    Credits=Credits+1
    PlaySound SoundFXDOF("knocker",117, DOFPulse, DOFKnocker)
    PlaySound "CHAlright"
    SpecialLight.state=0
  End If
end sub

Sub RightStar_Hit()
  if TableTilted=false then
    DOF 137, 2
    If RightStarLight.state=2 then
      RightOutlaneLight.state=1
      LeftOutlaneLight.state=1
      SetMotor(500)
      PlaySound"CHStars"
    Else
      SetMotor(50)
    End If
  end if
end sub

'*********************************** Special 7 Light1

Sub Special7()
  if (SpecialLight.state=2) and (SoulLight1.state=1) and (SoulLight2.state=1) and (SoulLight3.state=1) and (SoulLight4.state=1) then
    PlaySound "CH7Special"
    Credits=Credits+1
    PlaySound SoundFXDOF("knocker",117, DOFPulse, DOFKnocker)
    SpecialLight.state=0
  end If
End Sub


'*********************************** LoopLight Sequences

Sub LightSequence_timer
' For each obj in LoopLightsLeft
'   obj.state=0
' next
' LightSequenceCounter=LightSequenceCounter+1
' If LowerLeftStarLight.state=2 Then
'   If LightSequenceCounter > 5 then LightSequenceCounter=0
'     LoopLightsLeft(LightSequenceCounter).state=1
' End If
End Sub

Sub LightSequence2_timer
' For each obj in LoopLightsRight
'   obj.state=0
' next
' LightSequence2Counter=LightSequence2Counter+1
' If RightExtraBallLight.state=1 Then
'   If LightSequence2Counter > 2 then LightSequence2Counter=0
'     LoopLightsRight(LightSequence2Counter).state=1
' End If
End Sub

'*********************************** KICKERS

Dim rkickstep, lkickstep, ktimer, kickFill, rKickerHit, lKickerHit


Dim kickBalls
Dim kickerBall1
Sub kickBall(kball, kangle, kvel, kvelz, kzlift)
  dim rangle
  rangle = 3.14 * (kangle - 90) / 180
  kball.z = kball.z + kzlift
  kball.velz = kvelz
  kball.velx = cos(rangle)*kvel
  kball.vely = sin(rangle)*kvel
End Sub


'****************************************************
' Kickers
'****************************************************
Dim PerditionEntry(4)
Sub Kicker1_Hit
' set kickerBall1 = activeball
  If TableTilted=false then
    If PerditionEntry(Player) < 6 Then
      Kicker1Hold.enabled=true
    Elseif PerditionEntry(Player) = 6 Then
'     Kicker1Hold.enabled=False
'     PerditionFire.enabled=True
      SetMotor(300)
      PlaySound "CHPerditionKicker"
      If ColoredBall = 1 Then
        Dim BOT
        Bot = GetBalls
        BOT(0).color = RGB(255, 33, 0)
      End If
      Multiball.enabled = 1
    End If
  Else
    set kickerBall1 = activeball
    Pkickarm1.rotz=15
    KickerArm1.enabled=true
    kickBall KickerBall1, 1, 45, 5, 30
  End If
end sub

Sub PerditionFire_timer
  PerditionFire.enabled=False
' Kicker1.kick 3,45
  if kicker1.ballcntover > 0 Then
    Pkickarm1.rotz=15
    KickerArm1.enabled=true
    kickBall KickerBall1, 1, 45, 5, 30
  PlaySoundAt SoundFXDOF("saucer",131,DOFPulse,DOFcontactors), Pkickarm1
  DOF 123, DOFPulse
  Else
  end if
End Sub


Sub Kicker1Hold_timer()
  dim leftboosttemp
' If MotorRunning<>0 then
'   exit sub
' end if
  Kicker1Hold.enabled=false
  leftboosttemp=0
  if PerditionLight1.state=1 then
    leftboosttemp=leftboosttemp+1
  end if
  if PerditionLight2.state=1 then
    leftboosttemp=leftboosttemp+1
  end if
  if PerditionLight3.state=1 then
    leftboosttemp=leftboosttemp+1
  end if
  BonusBoosterCounter=leftboosttemp
  if BonusBoosterCounter>0 then
    BonusBoost.enabled=true
  end if
  If TableTilted=false then
    SetMotor(300)
  end if
' Kicker1.TimerInterval=1000
  Kicker1.TimerEnabled=true
end sub

Dim MultiballOn, MultiballReleased
Sub Multiball_Timer
  Ballrelease.CreateSizedBall 25
  MultiballOn = 1
  If ColoredBall =  1 Then
    DIM BOT
    Bot = GetBalls
    BOT(1).color = RGB(100, 150, 255)
  End If
    Ballrelease.Kick 40,7
    ' Thalamus added next
    PlaysoundAt "ballrelease", Plunger
  DOF 112, DOFPulse
  perditionpost.transz = 0
  PerditionWall.collidable = 1
  Multiball.enabled = 0
  bumpercap2.image="bcap-100wl-unlit devil"
  Bumper2Light.state = 2
  Bumper1Light.state = 0
  Bumper3Light.state = 0
  GISet
  PerditionEntry(Player) = 0
End Sub


Sub Kicker1_Timer()
  if kicker1.ballcntover > 0 Then
    dim  BOT, b
    BOT = GetBalls
    For b = 0 to ubound(BOT)
      If BOT(b).x  > 790 and BOT(b).y < 800 Then Set KickerBall1 = BOT(b)
    Next
    Pkickarm1.rotz=15
    KickerArm1.enabled=true
    kickBall KickerBall1, 0, 45, 5, 30
    PlaySoundAt SoundFXDOF("saucer",131,DOFPulse,DOFcontactors), Pkickarm1
  End if
  If TableTilted = True Then
    Pkickarm1.rotz=15
    KickerArm1.enabled=true
    kickBall KickerBAll1, 0, 45, 5, 30
  End If
' DOF 118, DOFPulse
  Kicker1.TimerEnabled = False
end sub

sub KickerArm1_timer
  Pkickarm1.rotz=0
  KickerArm1.enabled=false
end sub

Sub Kicker2_Hit
  set kickerBall1 = activeball
  If TableTilted=false then
    Kicker2Hold.enabled=true
    If RightExtraBallLight.state=1 then
      PlaySound "CHExtraLifeAward"
      ExtraBall=True
      SamePlayerShootsAgain.state=1
      RightExtraBallLight.state=0
      If B2SOn Then
        Controller.B2SSetShootAgain 1
      End if
    end if
  else
    Kicker2.TimerEnabled=true
  end if
end sub

Sub Kicker2Hold_timer()
  dim rightboosttemp
  If MotorRunning<>0 then
    exit sub
  end if
  Kicker2Hold.enabled=false
  rightboosttemp=0
  if PerditionLight1.state=1 then
    rightboosttemp=rightboosttemp+1
  end if
  if PerditionLight2.state=1 then
    rightboosttemp=rightboosttemp+1
  end if
  if PerditionLight3.state=1 then
    rightboosttemp=rightboosttemp+1
  end if

  BonusBoosterCounter=rightboosttemp
  if BonusBoosterCounter>0 then
    BonusBoost.enabled=true
  end if
  If TableTilted=false then
    SetMotor(1000 + (1000 * rightboosttemp))
  end if
  Kicker2.TimerInterval=1000
  Kicker2.TimerEnabled=true
end sub

Sub Kicker2_Timer()
  if kicker2.ballcntover > 0 Then
    Pkickarm2.rotz=15
    KickerArm2.enabled=true
    kickBall KickerBall1, 320, 20, 5, 30
    PlaySoundAt SoundFXDOF("saucer",132,DOFPulse,DOFcontactors), Pkickarm2
    Kicker2.TimerEnabled=false
  Else
    Kicker2.TimerEnabled=false
  End if
' DOF 124, DOFPulse
end sub

sub KickerArm2_timer
  Pkickarm2.rotz=0
  KickerArm2.enabled=false
end sub

Sub Kicker3_Hit()
  set kickerBall1 = activeball
    Kicker3.TIMERINTERVAL = 1000
    Kicker3.TIMERENABLED = TRUE
    if TableTilted=false then
      if RightDashLight.state=1 then
        SetMotor(500)
      else
        SetMotor(100)
      end if
    end if
End Sub

Dim RandomKicker
Sub Kicker3_Timer()
  RandomKicker = int(RND * 3)
' tb.text = RandomKicker
  if kicker3.ballcntover > 0 Then
    kickBall KickerBall1, (307 + RandomKicker), 32.5, 5, 25
    PlaySoundAt SoundFXDOF("saucer",123,DOFPulse,DOFcontactors), Kicker3
    PlaySound "CHLowerKickers"
    Kicker3.TIMERENABLED = FALSE
  Else
    Kicker3.TIMERENABLED = FALSE
  End if
End Sub

Sub Kicker4_Hit()
  set kickerBall1 = activeball
    Kicker4.TIMERINTERVAL = 500
    Kicker4.TIMERENABLED = TRUE
    If TableTilted=false then
      if LeftDashLight.state=1 Then
        SetMotor(500)
      else
        SetMotor(100)
      end if
    end if
End Sub

Sub Kicker4_Timer()
  if kicker4.ballcntover > 0 Then
    kickBall KickerBall1, 57, 32, 5, 22
    PlaySoundAt SoundFXDOF("saucer",124,DOFPulse,DOFcontactors), Kicker4
    PlaySound "CHLowerKickers"
    Kicker4.TIMERENABLED = FALSE
  Else
    Kicker4.TIMERENABLED = FALSE
  End if
end Sub


'****************  DROP Targets CUP & Mug

Dim zMultiplier: zMultiplier = 2.2

Dim DTR(3), DTL(3)


Sub DT1_hit
  DTL(1) = 1
  PlaySoundAt "droptargetdropped", DT1
  DOF 113, DOFPulse
  addscore(100)
  DTcheck
  ShadowCup1.visible=False
  activeball.velz = activeball.velz*zMultiplier
end sub

Sub DT2_hit
  DTL(2) = 1
  PlaySoundAt "droptargetdropped", DT2
  DOF 114, DOFPulse
  addscore(100)
  DTcheck
  ShadowCup2.visible=False
  activeball.velz = activeball.velz*zMultiplier
end sub

Sub DT3_hit
  DTL(3) = 1
  PlaySoundAt "droptargetdropped", DT3
  DOF 115, DOFPulse
  addscore(100)
  DTcheck
  ShadowCup3.visible=False
  activeball.velz = activeball.velz*zMultiplier
end sub

Sub DT4_dropped
  DTR(1) = 1
  PlaySoundAt "droptargetdropped", DT4
  DOF 109, DOFPulse
  addscore(100)
  DTcheck2
  ShadowMug1.visible=False
  activeball.velz = activeball.velz*zMultiplier
end sub

Sub DT5_dropped
  DTR(2) = 1
  PlaySoundAt "droptargetdropped", DT5
  DOF 110, DOFPulse
  addscore(100)
  DTcheck2
  ShadowMug2.visible=False
  activeball.velz = activeball.velz*zMultiplier
end sub

Sub DT6_dropped
  DTR(3) = 1
  PlaySoundAt "droptargetdropped", DT6
  DOF 111, DOFPulse
  addscore(100)
  DTcheck2
  ShadowMug3.visible=False
  activeball.velz = activeball.velz*zMultiplier
end sub

Sub DTcheck
  if DTL(1) = 1 and DTL(2) = 1 and DTL(3) = 1 then
    DT1.timerenabled = 1
  End If
end sub

Sub DTcheck2
  if DTR(1) = 1 and DTR(2) = 1 and DTR(3) = 1 Then DT2.timerenabled=1
end sub

Sub DT1_timer
  playsoundat "droptargetreset", DT1
  for i = 1 to 3
    DTL(i) = 0
    EVAL("DT" & i).isdropped = False
  next
  DT1.timerenabled=0
  IncreaseSoulBonus
  for each obj in ShadowCup
    obj.visible=True
  Next
End Sub

Sub DT2_timer
  playsoundat "droptargetreset", DT2
  for i = 1 to 3
    EVAL("DT"&i+3).isdropped=False
    DTR(i) = 0
  next
  DT2.timerenabled=0
  IncreaseSoulBonus
  for each obj in ShadowMug
    obj.visible=True
  Next
End Sub

'************************ Cuphead Rubber
Sub phys_bands001_hit
  AddScore(10)
  If ShiftControl = 1 Then SoulLightsLeft
End Sub

'************************ Mugman Rubber
Sub phys_bands002_hit
  AddScore(10)
  If ShiftControl = 1 Then SoulLightsRight
End Sub


'************************ STANDUP Targets

Sub StandUpRight_hit()
  PlaySoundAtBall "target"
  PlaySound "CHStandups"
  DOF 125, DOFPulse
  AddScore(10)
  RightInlaneLight.state=1
  LeftInlaneLight.state=1
  RightStarLight.state=2
  UpperLeftStarLight.state=2
  LowerLeftStarLight.state=2
end sub

Sub StandUpLeft1_hit()
  PlaySoundAtBall "target"
  PlaySound "CHStandups"
  DOF 138, DOFPulse
  AddScore(10)
  LeftStandupLight1.state=1
  CheckStandups
end sub

Sub StandUpLeft2_hit()
  PlaySoundAtBall "target"
  PlaySound "CHStandups"
  DOF 139, DOFPulse
  AddScore(10)
  LeftStandupLight2.state=1
  CheckStandups
end sub

Sub StandUpLeft3_hit()
  PlaySoundAtBall "target"
  PlaySound "CHStandups"
  DOF 140, DOFPulse
  AddScore(10)
  LeftStandupLight3.state=1
  CheckStandups
end sub

Sub CheckStandups
  if LeftStandupLight1.state=1 and LeftStandupLight2.state=1 and LeftStandupLight3.state=1 then
    LeftDashLight.state=1
    RightDashLight.state=1
  end if
end sub

'********************** Spinner1

' Thalamus : This sub is used twice - this means ... this one IS NOT USED
' I merge the scoring down to the other sub
' sub Spinner1_Spin()
'     AddScore(10)
' end sub

'********************** Relay

Sub AlternateRelayFire
  RightExtraBallLight.state=0

' If AlternateRelay = 1 then
'   PerditionLight1.state=1
'   PerditionLight2.state=1
'   PerditionLight3.state=1
'   AlternateRelay=2
'   Exit Sub
'
' Else
'   PerditionLight1.state=0
'   PerditionLight2.state=0
'   PerditionLight3.state=0
'   AlternateRelay=1
' end if
end sub

Sub AddSpecial()
  PlaySound SoundFXDOF("knocker",117, DOFPulse, DOFKnocker)
  Credits=Credits+1
  if Credits>15 then Credits=15
  If B2SOn Then
    Controller.B2SSetCredits Credits
  End If
  CreditsReel.SetValue(Credits)
End Sub

Sub AddSpecial2()
  PlaySound"click"
  Credits=Credits+1
  DOF 105, DOFOn
  if Credits>15 then Credits=15
  If B2SOn Then
    Controller.B2SSetCredits Credits
  End If
  CreditsReel.SetValue(Credits)
End Sub

Sub ResetBallDrops()

  RightExtraBallLight.state=0
  SnakeEyesLight.state=0
  SoulBonusCounter=0
  BonusMultiplierCounter=0
  BonusCounter=0
  SpecialLight.state=0
  for each obj in BonusX
    obj.state=0
  Next
  for each obj in SoulBonus
    obj.state=0
  Next
  for each obj in TopRolloverLights
    obj.state=0
  Next
' for each obj in LeftKickerLights
'   obj.state=0
' Next
' for each obj in RightKickerLights
'   obj.state=0
' Next
  for each obj in StandUpsDash
    obj.state=0
  Next
  for each obj in DashLights
    obj.state=0
  Next
  for each obj in StarLights
    obj.state=0
  Next
  for each obj in LoopLightsLeft
  Next
  for each obj in LoopLightsRight
    obj.state=0
  Next
  for each obj in RolloverLightsLeft
    obj.state=0
  Next
  for each obj in RolloverLightsRight
    obj.state=0
  Next
  for each obj in SlingLights
    obj.state=0
  Next
  for each obj in DropTargetsCUP
    obj.IsDropped=0
    DOF 116, DOFPulse
  Next
  For x = 1 to 3
    DTR(x) = 0
    DTL(x) = 0
  Next
  for each obj in DropTargetsMUG
    obj.IsDropped=0
    DOF 116, DOFPulse
  Next
End Sub


Sub ResetBalls()
  TempMultiCounter=BallsPerGame-BallInPlay
  ResetBallDrops
  BonusMultiplier=1
  For x = 1 to 3
    If PerditionEntry(Player) >= (x * 2) Then
      EVAL("PerditionLight" & x).state = 1
    Else
      EVAL("PerditionLight" & x).state = 0
    End If
  Next
  If LastBallDoubleBonus = 1 Then
    If BallInPlay = BallsPerGame Then BonusMultiplier = 2: Bonus2X.state = 1
  End If
  Bumper1Light.State=1
  Bumper2Light.State=1
  Bumper3Light.State=1
  SnakeEyesLight.state=0
  If SoulLightReset =  1 Then
    For x = 1 to 4
      If SoulLightArray(Player,x) = 1 Then
        EVAL("SoulLight" & x).state = 1
      Else
        EVAL("SoulLight" & x).state = 0
      End If
    Next
  End If
  BumpersOn
  for each obj in StarLights
    obj.state=1
  Next
  SoulBonusCounter=0
  SoulBonus(SoulBonusCounter).state=1
  TableTilted=false
  TiltReel.SetValue(0)
  If B2Son then
    Controller.B2SSetTilt 0
  end if
  PlasticsOn
  Ballrelease.CreateSizedBall 25
    Ballrelease.Kick 40,7
    ' Thalamus added next
    PlaysoundAt "ballrelease", Plunger
  DOF 112, DOFPulse
  BallInPlayReel.SetValue(BallInPlay)
End Sub

'**************************Soul Lights Right
Dim x, y
Dim TempSoulR(5), TempSoulL(5)
Sub SoulLightsRight
  For x = 1 to 4
    If EVAL("SoulLight" & x).state = 1 Then TempSoulR(x + 1) = 1
  Next
  For x = 1 to 4
    If TempSoulR(x) = 1 Then
      EVAL("SoulLight" & x).state = 1
      SoulLightArray(Player,x) = 1
    Else
      EVAL("SoulLight" & x).state = 0
      SoulLightArray(Player,x) = 0
    End If
  Next
  If TempSoulR(5) = 1 Then
    SoulLight1.state = 1
    SoulLightArray(Player,1) = 1
  Else
    SoulLight1.state = 0
    SoulLightArray(Player,1) = 0
  End If
  For x = 0 to 5
    TempSoulR(x) = 0
  Next
End Sub

'**************************Soul Lights Left
Sub SoulLightsLeft
  For x = 1 to 4
    If EVAL("SoulLight" & x).state = 1 Then TempSoulL(x - 1) = 1
  Next
  For x = 1 to 4
    If TempSoulL(x) = 1 Then
      EVAL("SoulLight" & x).state = 1
      SoulLightArray(Player,x) = 1
    Else
      EVAL("SoulLight" & x).state = 0
      SoulLightArray(Player,x) = 0
    End If
  Next
  If TempSoulL(0) = 1 Then
    SoulLight4.state = 1
    SoulLightArray(Player,4) = 1
  Else
    SoulLight4.state = 0
    SoulLightArray(Player,4) = 0
  End If
  For x = 0 to 5
    TempSoulL(x) = 0
  Next
End Sub

'**************************NEW SCORE Bonus

Sub IncreaseSoulBonus
  if SoulBonusCounter<15 then
    If SoulBonusCounter<10 then
      SoulBonus(SoulBonusCounter).state=0
      SoulBonusCounter=SoulBonusCounter+1
      SoulBonus(SoulBonusCounter).state=1
    elseif SoulBonusCounter>9 then
      SoulBonus(SoulBonusCounter-10).state=0
      SoulBonusCounter=SoulBonusCounter+1
      SoulBonus(10).state=1
      SoulBonus(SoulBonusCounter-10).state=1
    end if
  end if
  If SoulBonusCounter > 4 Then
    RightExtraBallLight.state=1
    PlaySound "CHExtraLifeActive"
  end If
  If SoulBonusCounter = 7 Then
    SpecialLight.state = 2
  End If
  SoulSounds
End Sub

sub ScoreBonusMultiplier

  If Bonus2X.state=0 Then
    ScoreMotorStepper=0
  else
    ScoreMotorStepper=0
  end if
  CollectBonusMultiplier.enabled=1
  If Bonus3X.state=0 Then
    ScoreMotorStepper=0
  else
    ScoreMotorStepper=0
  end if
  CollectBonusMultiplier.enabled=1
  If Bonus4X.state=0 Then
    ScoreMotorStepper=0
  else
    ScoreMotorStepper=0
  end if
  CollectBonusMultiplier.enabled=1
end sub

sub CollectBonusMultiplier_timer
  if MotorRunning=1 then
    exit sub
  end if
  if SoulBonusCounter<1 then
    CollectBonusMultiplier.enabled=0
    SoulBonusCounter=1
    SoulBonus(SoulBonusCounter).state=1
  else
    if ScoreMotorStepper<5 then
    If TableTilted=false then
        SetMotor(5000)
      end if
      If Bonus2X.state=0 then
        if TableTilted=false then
          SetMotor(5000)
        end if
      else
        if TableTilted=false then
          AddScore(10000)
        end if
      end if
      SoulBonus(SoulBonusCounter).state=0
      SoulBonusCounter=SoulBonusCounter-1
      if SoulBonusCounter>=0 then
        SoulBonus(SoulBonusCounter).state=1
      end if
      ScoreMotorStepper=ScoreMotorStepper+1
    else
      If Bonus2X.state=0 then
        ScoreMotorStepper=0
      else
        ScoreMotorStepper=0
      end if
      If Bonus3X.state=0 then
        ScoreMotorStepper=0
      else
        ScoreMotorStepper=0
      end if
      If Bonus4X.state=0 then
        ScoreMotorStepper=0
      else
        ScoreMotorStepper=0
      end if
    end if
  end if
end sub

sub ScoreSoulBonus
  ScoreMotorStepper=0
  CollectSoulBonus.interval=135
  CollectSoulBonus.enabled=1
end sub

sub CollectSoulBonus_timer
  if MotorRunning=1 then
    exit sub
  end if
  if SoulBonusCounter<1 then
    CollectSoulBonus.enabled=0
    NextBallDelay.enabled=true
  else
    If Bonus2X.state=1 then
      Select case ScoreMotorStepper
        case 0,1,3,4:
          AddScore(1000)

        case 2,5:
          PlaySound"pinhit_low"
          If SoulBonusCounter>10 then
            SoulBonus(SoulBonusCounter-10).state=0
          else
            SoulBonus(SoulBonusCounter).state=0
          end if

          SoulBonusCounter=SoulBonusCounter-1
          if SoulBonusCounter>=0 then
            If SoulBonusCounter<=10 then
              SoulBonus(SoulBonusCounter).state=1
            elseif SoulBonusCounter>10 then
              SoulBonus(10).state=1
              SoulBonus(SoulBonusCounter-10).state=1
            end if
          end if

      end select
      ScoreMotorStepper=ScoreMotorStepper+1
      If ScoreMotorStepper>5 then
        ScoreMotorStepper=0
      end if
    elseif Bonus3x.state=1 then
      Select case ScoreMotorStepper
        case 0,1,2,4,5,6:
          AddScore(1000)

        case 3,7:
          PlaySound"pinhit_low"
          If SoulBonusCounter>10 then
            SoulBonus(SoulBonusCounter-10).state=0
          else
            SoulBonus(SoulBonusCounter).state=0
          end if

          SoulBonusCounter=SoulBonusCounter-1
          if SoulBonusCounter>=0 then
            If SoulBonusCounter<=10 then
              SoulBonus(SoulBonusCounter).state=1
            elseif SoulBonusCounter>10 then
              SoulBonus(10).state=1
              SoulBonus(SoulBonusCounter-10).state=1
            end if
          end if

      end select
      ScoreMotorStepper=ScoreMotorStepper+1
      If ScoreMotorStepper>7 then
        ScoreMotorStepper=0
      end if
    elseif Bonus4x.state=1 then
      Select case ScoreMotorStepper
        case 0,1,2,3,5,6,7,8:
          AddScore(1000)

        case 4,9:
          PlaySound"pinhit_low"
          If SoulBonusCounter>10 then
            SoulBonus(SoulBonusCounter-10).state=0
          else
            SoulBonus(SoulBonusCounter).state=0

          end if

          SoulBonusCounter=SoulBonusCounter-1
          if SoulBonusCounter>=0 then
            If SoulBonusCounter<=10 then
              SoulBonus(SoulBonusCounter).state=1
            elseif SoulBonusCounter>10 then
              SoulBonus(10).state=1
              SoulBonus(SoulBonusCounter-10).state=1
            end if
          end if
      end select

      ScoreMotorStepper=ScoreMotorStepper+1
      If ScoreMotorStepper>9 then
        ScoreMotorStepper=0
      end if
    else
      Select Case ScoreMotorStepper
        case 0,1,2,3,4:
          AddScore(1000)
          If SoulBonusCounter>10 then
            SoulBonus(SoulBonusCounter-10).state=0
          else
            SoulBonus(SoulBonusCounter).state=0

          end if
          SoulBonusCounter=SoulBonusCounter-1
          if SoulBonusCounter>=0 then
            If SoulBonusCounter<=10 then
              SoulBonus(SoulBonusCounter).state=1
            elseif SoulBonusCounter>10 then
              SoulBonus(10).state=1
              SoulBonus(SoulBonusCounter-10).state=1
            end if
          end if
        case 5:
          PlaySound"pinhit_low"

      end select
      ScoreMotorStepper=ScoreMotorStepper+1
      If ScoreMotorStepper>7 then
        ScoreMotorStepper=0
      end if
    end if
  end if
end sub

'********************************* SOUL SOUNDS

Sub SoulSounds
  If Bonus1.state=1 Then
    PlaySound "SoulMusic1"
  Elseif Bonus2.state=1 Then
    PlaySound "SoulMusic2"
  Elseif Bonus3.state=1 Then
    PlaySound "SoulMusic3"
  Elseif Bonus4.state=1 Then
    PlaySound "SoulMusic4"
  Elseif Bonus5.state=1 Then
    PlaySound "SoulMusic5"
  Elseif Bonus6.state=1 Then
    PlaySound "SoulMusic6"
  Elseif Bonus7.state=1 Then
    PlaySound "SoulMusic7"
  Elseif Bonus8.state=1 Then
    PlaySound "SoulMusic8"
  Elseif Bonus9.state=1 Then
    PlaySound "SoulMusic9"
  Elseif Bonus10.state=1 Then
    PlaySound "SoulMusic10"
  End If
End Sub

'********************************GI Light Subs
Sub GISet
  For each obj in Flashers
    obj.color = RGB(255,0,0)
    obj.colorfull = RGB(255,0,0)
  Next

  For each obj in Flashers2
    obj.color = RGB(255,0,0)
  Next
End Sub

Sub GIReset
  For each obj in Flashers
    obj.color = RGB(255,128,0)
    obj.colorfull = RGB(255,197,143)
    obj.state = 1
  Next

  For each obj in Flashers2
    obj.color = RGB(128,0,0)
    obj.colorfull = RGB(255,0,0)
    obj.state = 1
  Next
End Sub

Sub GIOff
  For each obj in Flashers
    obj.state = 0
  Next
  For each obj in Flashers2
    obj.state = 0
  Next
End Sub


'*********************************

sub resettimer_timer
    rst=rst+1
    for i=1 to 4
  If B2SOn Then
    Controller.B2SSetScorePlayer i, 0
  End If
  next
    if rst=20 then
    playsound "StartBall1"
    end if
    if rst=24 then
    newgame
    resettimer.enabled=false
    end if
end sub

'***********************************

Sub NextBallDelay_timer()
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
    Controller.B2SSetShootAgain 0
  End if
  Bumper1Light.state=1
  Bumper2Light.state=1
  Bumper3Light.state=1
  AlternateRelay=1
  BumpersOn
  For x = 1 to 4
    PerditionEntry(x) = 1
    For y = 1 to 4
      SoulLightArray(x,y) = 0
    Next
  Next
  ResetBalls
  For x = 1 to 3
    EVAL("PerditionLight" & x).state = 0
  Next
End sub

sub nextball
  If B2SOn Then
    Controller.B2SSetTilt 0
    Controller.B2SSetShootAgain 0
  End If
  If Bonus4X.state = 1 Then
    For x = 1 to 4
      SoulLightArray(Player, x) = 0
    Next
  End If
  If ExtraBall=false then
    Player=Player+1
  else
    ExtraBall=false
    SamePlayerShootsAgain.state=0
  end if
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
'     For each obj in PlayerHuds
'       obj.SetValue(0)
'     next
'     For each obj in PlayerHUDScores
'       obj.state=0
'     next
      If Table1.ShowDT = True then
        For each obj in PlayerScores
          obj.visible=1
        Next
        For each obj in PlayerScoresOn
          obj.visible=0
        Next
      end If
      BallInPlayReel.SetValue(0)
      CanPlayReel.SetValue(0)
      GameOverReel.SetValue(1)
      LeftFlipper.RotateToStart
      RightFlipper.RotateToStart
      LightsOut
      PlasticsOff
      BumpersOff
      for each obj in StarLights
        obj.state=0
      next
      checkmatch
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
      PlaySound("RotateThruPlayers")
      TempPlayerUp=Player
      PlayerUpRotator.enabled=true
      PlayStartBall.enabled=true
'     For each obj in PlayerHuds
'       obj.SetValue(0)
'     next
'     For each obj in PlayerHUDScores
'       obj.state=0
'     next
      If Table1.ShowDT = True then
        For each obj in PlayerScores
          obj.visible=1
        Next
        For each obj in PlayerScoresOn
          obj.visible=0
        Next
'       PlayerHuds(Player-1).SetValue(1)
'       PlayerHUDScores(Player-1).state=1
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
    PlaySound("RotateThruPlayers")
    TempPlayerUp=Player
    PlayerUpRotator.enabled=true
    PlayStartBall.enabled=true
'   For each obj in PlayerHuds
'     obj.SetValue(0)
'   next
'   For each obj in PlayerHUDScores
'       obj.state=0
'     next
    If Table1.ShowDT = True then
      For each obj in PlayerScores
          obj.visible=1
      Next
      For each obj in PlayerScoresOn
          obj.visible=0
      Next
'     PlayerHuds(Player-1).SetValue(1)
'     PlayerHUDScores(Player-1).state=1
      PlayerScores(Player-1).visible=0
      PlayerScoresOn(Player-1).visible=1
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
      TiltReel.SetValue(1)
      PlasticsOff
      BumpersOff
      for each obj in StarLights
        obj.state=0
      next
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

Sub BonusBoost_Timer()
' IncreaseBonus
  BonusBoosterCounter=BonusBoosterCounter-1
  If BonusBoosterCounter=0 then
    BonusBoost.enabled=false
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
'     For each obj in PlayerHuds
'       obj.SetValue(0)
'     next
'     For each obj in PlayerHUDScores
'       obj.state=0
'     next
      If Table1.ShowDT = True then
        For each obj in PlayerScores
          obj.visible=1
        Next
        For each obj in PlayerScoresOn
          obj.visible=0
        Next
'       PlayerHuds(TempPlayerUp-1).SetValue(1)
'       PlayerHUDScores(TempPlayerUp-1).state=1
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
'     For each obj in PlayerHuds
'       obj.SetValue(0)
'     next
'     For each obj in PlayerHUDScores
'       obj.state=0
'     next
      If Table1.ShowDT = True then
        For each obj in PlayerScores
          obj.visible=1
        Next
        For each obj in PlayerScoresOn
          obj.visible=0
        Next
'       PlayerHuds(Player-1).SetValue(1)
'       PlayerHUDScores(Player-1).state=1
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
    ScoreFile.WriteLine
    ScoreFile.WriteLine MusicVolume
    ScoreFile.WriteLine Credits
    scorefile.writeline BallsPerGame
    scorefile.writeline ExtraBallSetting
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
    temp6=textstr.readline
    if HighScore<1 then
      temp7=textstr.readline
      temp8=textstr.readline
      temp9=textstr.readline
      temp10=textstr.readline
      temp11=textstr.readline
      temp12=textstr.readline
      temp13=textstr.readline
      temp14=textstr.readline
      temp15=textstr.readline
      temp16=textstr.readline
    end if

    TextStr.Close
    MusicVolume = cdbl(temp2)
      Credits=cdbl(temp3)
    If Credits > 0 Then DOF 126, DOFOn
    BallsPerGame=cdbl(temp4)
    ExtraBallSetting=cdbl(temp5)
    ReplayLevel=cdbl(temp6)

    if HighScore<1 then
      HSScore(1) = int(temp7)
      HSScore(2) = int(temp8)
      HSScore(3) = int(temp9)
      HSScore(4) = int(temp10)
      HSScore(5) = int(temp11)

      HSName(1) = temp12
      HSName(2) = temp13
      HSName(3) = temp14
      HSName(4) = temp15
      HSName(5) = temp16
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
  if B2SOn then
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
    FlipperLSh001.visible = ShadowFlippersOn
    FlipperRSh001.visible = ShadowFlippersOn
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
    FlipperLSh001.visible=false
    FlipperRSh001.visible=false
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
      'Controller.B2SSetScore 6,HighScore
    End If
    DisplayHighScore
    CreditsReel.SetValue(Credits)
    InitPauser5.enabled=false
end sub

sub LightsOut

end sub

sub ToggleBumper

end sub

sub BumpersOff
  Bumper1Light.Visible=false
  Bumper2Light.Visible=false
  Bumper3Light.Visible=false
end sub

sub BumpersOn
  Bumper1Light.Visible=True
  Bumper2Light.Visible=True
  Bumper3Light.Visible=True
end Sub


Sub PlasticsOn

  GIReset
  Light3.state=1
  Light2.state=1
  ShadowMain.visible=True
  for each obj in ShadowCup
    obj.visible=True
  Next
  for each obj in ShadowMug
    obj.visible=True
  Next
  If FlipperLength = 2 Then
    FlipperLSh.visible=True
    FlipperRSh.visible=True
    FlipperLSh001.visible=False
    FlipperRSh001.visible=False
  Else
    FlipperLSh.visible=False
    FlipperRSh.visible=False
    FlipperLSh001.visible=True
    FlipperRSh001.visible=True
  End If
end sub

Sub PlasticsOff
  GIOff
  Light3.state=0
  Light2.state=0
  StopSound "buzzL"
  StopSound "buzz"
    If MusicOn = 1 and BotchedMB = 0 Then PlayMusic"CupAttract.mp3", MusicVolume
  ShadowMain.visible=False
  for each obj in ShadowCup
    obj.visible=False
  Next
  for each obj in ShadowMug
    obj.visible=False
  Next
  FlipperLSh.visible=False
  FlipperRSh.visible=False
  FlipperLSh001.visible=False
  FlipperRSh001.visible=False
end sub

Sub SetupReplayTables

  Replay1Table(1)=65000
  Replay1Table(2)=68000
  Replay1Table(3)=70000
  Replay1Table(4)=72000
  Replay1Table(5)=74000
  Replay1Table(6)=76000
  Replay1Table(7)=78000
  Replay1Table(8)=80000
  Replay1Table(9)=83000
  Replay1Table(10)=85000
  Replay1Table(11)=87000
  Replay1Table(12)=90000
  Replay1Table(13)=999000
  Replay1Table(14)=999000
  Replay1Table(15)=999000

  Replay2Table(1)=77000
  Replay2Table(2)=80000
  Replay2Table(3)=82000
  Replay2Table(4)=84000
  Replay2Table(5)=86000
  Replay2Table(6)=88000
  Replay2Table(7)=90000
  Replay2Table(8)=92000
  Replay2Table(9)=95000
  Replay2Table(10)=97000
  Replay2Table(11)=99000
  Replay2Table(12)=999000
  Replay2Table(13)=999000
  Replay2Table(14)=999000
  Replay2Table(15)=999000

  Replay3Table(1)=999000
  Replay3Table(2)=999000
  Replay3Table(3)=999000
  Replay3Table(4)=999000
  Replay3Table(5)=999000
  Replay3Table(6)=999000
  Replay3Table(7)=999000
  Replay3Table(8)=999000
  Replay3Table(9)=999000
  Replay3Table(10)=999000
  Replay3Table(11)=999000
  Replay3Table(12)=999000
  Replay3Table(13)=999000
  Replay3Table(14)=999000
  Replay3Table(15)=999000

  ReplayTableMax=12

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
        AddScore(100)
        MotorRunning=0
        BumpersOn
      Case 200:
        MotorMode=100
        MotorPosition=2
        BumpersOff
      Case 300:
        MotorMode=100
        MotorPosition=3
        BumpersOff
      Case 400:
        MotorMode=100
        MotorPosition=4
        BumpersOff
      Case 500:
        MotorMode=100
        MotorPosition=5
        BumpersOff
      Case 1000:
        AddScore(1000)
        MotorRunning=0
        BumpersOn
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
  If BotchedMB = 1 Then Exit Sub
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
  if Score(Player)=>100000 then
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

'*********************************************************************
'                 Positional Sound Playback Functions
'*********************************************************************

' Play a sound, depending on the X,Y position of the table element (especially cool for surround speaker setups, otherwise stereo panning only)
' parameters (defaults): loopcount (1), volume (1), randompitch (0), pitch (0), useexisting (0), restart (1))
' Note that this will not work (currently) for walls/slingshots as these do not feature a simple, single X,Y position
Sub PlayXYSound(soundname, tableobj, loopcount, volume, randompitch, pitch, useexisting, restart)
  PlaySound soundname, loopcount, volume, AudioPan(tableobj), randompitch, pitch, useexisting, restart, AudioFade(tableobj)
End Sub

' Similar subroutines that are less complicated to use (e.g. simply use standard parameters for the PlaySound call)
Sub PlaySoundAt(soundname, tableobj)
    PlaySound soundname, 1, 1, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname)
    PlaySoundAt soundname, ActiveBall
End Sub

Sub PlayLoopSoundAtVol(sound, tableobj, Vol)
  PlaySound sound, -1, Vol, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBOTBallZ(sound, BOT)
    PlaySound sound, 0, ABS(BOT.velz)/17, AudioPan(BOT), 0, Pitch(BOT), 1, 0, AudioFade(BOT)
End Sub

'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
    tmp = tableobj.y * 2 / table1.height-1
    If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
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

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
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
    ' Thalamus  - added next three lines for effect after popper - Roth's ball jump
    If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
      PlaySoundAtBOTBallZ "fx_ball_drop" & b, BOT(b)
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

Sub Spinner1_Spin
  PlaySound "fx_spinner", 0, .25, AudioPan(Spinner1), 0.25, 0, 0, 1, AudioFade(Spinner1)
  PlaySound "CHWhip"
  AddScore(10)
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

'Dim PlungerPulled
'PlungerPulled = 0

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

HSName(1) = "CUPHEAD"
HSName(2) = "MUGMAN"
HSName(3) = "DEVIL"
HSName(4) = "KINGDICE"
HSName(5) = "ELDER"

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
    EnteringInitials = 1:If MusicOn = 1 Then PlayMusic"CupHighScore.mp3" , MusicVolume' intercept the control keys while entering initials
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
Const OptionLinesToMark="111100011"
Const OptionLine1="" 'do not use this line
Const OptionLine2="" 'do not use this line
Const OptionLine3="" 'do not use this line
Const OptionLine4="Extra Ball Light"
Const OptionLine5=""
Const OptionLine6=""
Const OptionLine7=""
Const OptionLine8="" 'do not use this line
Const OptionLine9="" 'do not use this line

Sub OperatorMenuTimer_Timer
  OperatorMenuBackdrop.image = "OperatorMenu"
  EnteringOptions = 1
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
        If Replay2Table(ReplayLevel)=999000 then
          tempstring=FormatNumber(Replay1Table(ReplayLevel),0)
        elseif Replay3Table(ReplayLevel)=999000 then
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
        If ExtraBallSetting=1 then
          tempstring = "Use Alternating Relay"
        else
          tempstring = "Do Not Use Alternating Relay"
        end if
        SetOptLine 5, tempstring
      Case 5:
        SetOptLine 6, tempstring
          Tempstring = "Press MagnaSave Buttons"
        SetOptLine 7, tempstring
      Case 6:
        SetOptLine 8, tempstring
          TempString = "to change Music Volume"
        SetOptLine 9, tempstring

      Case 7:
        SetOptLine 10, tempstring
          tempstring = FormatNumber(MusicVolume)
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
      If ExtraBallSetting=1 then
        ExtraBallSetting=2
      else
        ExtraBallSetting=1
      end if
      DisplayAllOptions
    elseif CurrentOption = 8 or CurrentOption = 9 then
        if OptionCHS=1 then
          HSScore(1) = 75000
          HSScore(2) = 70000
          HSScore(3) = 60000
          HSScore(4) = 55000
          HSScore(5) = 50000

          HSName(1) = "CUPHEAD"
          HSName(2) = "MUGMAN"
          HSName(3) = "DEVIL"
          HSName(4) = "KINGDICE"
          HSName(5) = "ELDER"
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
        InstructCard.image="IC_"+FormatNumber(BallsPerGame,0)
        EnteringOptions = 0
    end if

  Elseif keycode = LeftMagnaSave Then
        MusicVolume = MusicVolume - .1
        If MusicVolume < 0 Then MusicVolume = 0
        PlayMusic "CupMusic5.mp3" , MusicVolume
        DisplayAllOptions

  Elseif keycode =RightMagnaSave Then
        MusicVolume = MusicVolume + .1
        If MusicVolume > 1 Then MusicVolume = 1
        PlayMusic "CupMusic5.mp3" , MusicVolume
        DisplayAllOptions
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

Sub SetOptLine(LineNo, String)
  dim xfor
  dim Letter
  dim ThisDigit
  dim ThisChar
  dim StrLen
  dim LetterLine
  dim Index
  dim LetterName
  StrLen = len(string)
  Index = 1

  for xfor = StartingArray(LineNo) to EndingArray(LineNo)
    Eval("Option"&xfor).image = GetOptChar(string, Index)
    Index = Index + 1
  next
End Sub


'******************************************************
'       FLIPPER AND RUBBER CORRECTION
'******************************************************
dim LFPress, RFPress, EOST, EOSA, EOSTnew, EOSaNew,Plunge
dim FStrength, FRampUp, fElasticity, EOSRampUp, SOSRampUp
dim RFEndAngle, LFEndAngle, LF1EndAngle, RF1EndAngle, LFCount, RFCount, LiveCatch

LFEndAngle = Leftflipper.EndAngle
LF1Endangle = LeftFlipper001.EndAngle
RFEndAngle = RightFlipper.EndAngle
RF1EndAngle = RightFlipper001.EndAngle

EOST = leftflipper.eosTorque        'End of Swing Torque
EOSA = leftflipper.eosTorqueAngle   'End of Swing Torque Angle
fStrength = LeftFlipper.strength    'Flipper Strength
fRampUp = LeftFlipper.RampUp      'Flipper Ramp Up
fElasticity = LeftFlipper.elasticity  'Flipper Elasticity
EOStNew = 1.0     'new Flipper Torque
EOSaNew = 0.2   'new FLipper Tprque Angle
EOSRampUp = 1.5   'new EOS Ramp Up weaker at EOS because of the weaker holding coil
SOSRampUp = 8.5   'new SOS Ramp Up strong at start because of the stronger starting coil
LiveCatch = 8   'variable to check elapsed time from

'********Need to have a flipper timer to check for these values
Dim PlungerCheck, PlungerCheck1, PlungerLoop, PlungerPlaying, PlungerPause
Dim PlungerMax, PlungerMin
Sub flipperTimer_Timer
  lFlip.rotz = leftflipper.currentangle -121
  rFlip.rotz = rightflipper.currentangle +121
  FlipperLSh001.RotZ = LeftFlipper001.currentangle
  FlipperRSh001.RotZ = RightFlipper001.currentangle
  FlipperLSh.RotZ = LeftFlipper.currentangle
  FlipperRSh.RotZ = RightFlipper.currentangle

  lFlip.rotz = leftflipper.CurrentAngle -121 'silver metal flipper obj
  lFlipR.rotz = leftflipper.CurrentAngle -121
  rFlip.rotz = RightFlipper.CurrentAngle +121
  rFlipR.rotz = RightFlipper.CurrentAngle +121
  lFlip001.roty = leftflipper001.CurrentAngle
  rFlip001.roty = RightFlipper001.CurrentAngle


' FlipperLShadow1.RotZ = LeftFlipper1.CurrentAngle
' FlipperRShadow1.RotZ = RightFlipper1.CurrentAngle

  '--------------Flipper Tricks Section
  'What this code does is swing the flipper fast and make the flipper soft near its EOS to enable live catches.  It resets back to the base Table
  'settings once the flipper reaches the end of swing.  The code also makes the flipper starting ramp up high to simulate the stronger starting
  'coil strength and weaker at its EOS to simulate the weaker hold coil.

  If LeftFlipper.CurrentAngle = LeftFlipper.EndAngle and LFPress = 1 Then   'If the flipper is fully swung and the flipper button is pressed then...
    LeftFlipper.eosTorqueAngle = EOSaNew  'sets flipper EOS Torque Angle to .2
    LeftFlipper.eosTorque = EOStNew     'sets flipper EOS Torque to 1
    LeftFlipper.RampUp = EOSRampUp      'sets flipper ramp up to 1.5
'   If LFCount = 0 Then LFCount = GameTime  'sets the variable LFCount = to the elapsed game time
    If GameTime - LFCount < LiveCatch Then  'if less than 8ms have elasped then we are in a "Live Catch" scenario
      TextBox1.text = "Catch"
      LeftFlipper.Elasticity = 0.1    'sets flipper elasticity WAY DOWN to allow Live Catches
      If LeftFlipper.EndAngle <> LFEndAngle Then LeftFlipper.EndAngle = LFEndAngle  'Keep the flipper at its EOS and don't let it deflect
    Else
      LeftFlipper.Elasticity = fElasticity  'reset flipper elasticity to the base table setting
      TextBox1.text = fElasticity
    End If
  Elseif LeftFlipper.CurrentAngle > LeftFlipper.startangle - 0.05  Then   'If the flipper has started its swing, make it swing fast to nearly the end...
    LeftFlipper.RampUp = SOSRampUp        'set flipper Ramp Up high
    LeftFlipper.EndAngle = LFEndAngle - 3   'swing to within 3 degrees of EOS
    LeftFlipper.Elasticity = fElasticity    'Set the elasticity to the base table elasticity
    LFCount = 0                 'Sets LF Count = 0 which would override the LFCount = GameTime set in hitting the flipper trigger zone (why??)
  Elseif LeftFlipper.CurrentAngle > LeftFlipper.EndAngle + 0.01 Then  'If the flipper has swung past it's end of swing then...
    LeftFlipper.eosTorque = EOST      'set the flipper EOS Torque back to the base table setting
    LeftFlipper.eosTorqueAngle = EOSA   'set the flipper EOS Torque Angle back to the base table setting
    LeftFlipper.RampUp = fRampUp      'set the flipper Ramp Up back to the base table setting
    LeftFlipper.Elasticity = fElasticity  'set the flipper Elasticity back to the base table setting
  End If

  If RightFlipper.CurrentAngle = RightFlipper.EndAngle and RFPress = 1 Then
    RightFlipper.eosTorqueAngle = EOSaNew
    RightFlipper.eosTorque = EOStNew
    RightFlipper.RampUp = EOSRampUp
'   If RFCount = 0 Then RFCount = GameTime
    If GameTime - RFCount < LiveCatch Then
      RightFlipper.Elasticity = 0.1
      If RightFlipper.EndAngle <> RFEndAngle Then RightFlipper.EndAngle = RFEndAngle
    Else
      RightFlipper.Elasticity = fElasticity
    End If
  Elseif RightFlipper.CurrentAngle < RightFlipper.StartAngle + 0.05 Then
    RightFlipper.RampUp = SOSRampUp
    RightFlipper.EndAngle = RFEndAngle + 3
    RightFlipper.Elasticity = fElasticity
    RFCount = 0
  Elseif RightFlipper.CurrentAngle < RightFlipper.EndAngle - 0.01 Then
    RightFlipper.eosTorque = EOST
    RightFlipper.eosTorqueAngle = EOSA
    RightFlipper.RampUp = fRampUp
    RightFlipper.Elasticity = fElasticity
  End If

  If LeftFlipper001.CurrentAngle = LeftFlipper001.EndAngle and LFPress = 1 Then
    LeftFlipper001.eosTorqueAngle = EOSaNew
    LeftFlipper001.eosTorque = EOStNew
    LeftFlipper001.RampUp = EOSRampUp
'   If LFCount = 0 Then LFCount = GameTime
    If GameTime - LFCount < LiveCatch Then
      LeftFlipper001.Elasticity = 0.1
      If LeftFlipper001.EndAngle <> LF1EndAngle Then LeftFlipper001.EndAngle = LF1EndAngle
    Else
      LeftFlipper001.Elasticity = fElasticity
    End If
  Elseif LeftFlipper001.CurrentAngle > LeftFlipper001.StartAngle - 0.05  Then
    LeftFlipper001.RampUp = SOSRampUp
    LeftFlipper001.EndAngle = LF1EndAngle - 3
    LeftFlipper001.Elasticity = fElasticity
    LFCount = 0
  Elseif LeftFlipper001.CurrentAngle > LeftFlipper001.EndAngle + 0.01 Then
    LeftFlipper001.eosTorque = EOST
    LeftFlipper001.eosTorqueAngle = EOSA
    LeftFlipper001.RampUp = fRampUp
    LeftFlipper001.Elasticity = fElasticity
  End If

  If RightFlipper001.CurrentAngle = RightFlipper.EndAngle and RFPress = 1 Then
    RightFlipper001.eosTorqueAngle = EOSaNew
    RightFlipper001.eosTorque = EOStNew
    RightFlipper001.RampUp = EOSRampUp
'   If RFCount = 0 Then RFCount = GameTime
    If GameTime - RFCount < LiveCatch Then
      RightFlipper001.Elasticity = 0.1
      If RightFlipper001.EndAngle <> RF1EndAngle Then RightFlipper001.EndAngle = RF1EndAngle
    Else
      RightFlipper001.Elasticity = fElasticity
    End if
  Elseif RightFlipper001.CurrentAngle < RightFlipper001.StartAngle + 0.05 Then
    RightFlipper001.RampUp = SOSRampUp
    RightFlipper001.EndAngle = RF1EndAngle + 3
    RightFlipper001.Elasticity = fElasticity
    RFCount = 0
  Elseif RightFlipper001.CurrentAngle < RightFlipper001.EndAngle - 0.01 Then
    RightFlipper001.eosTorque = EOST
    RightFlipper001.eosTorqueAngle = EOSA
    RightFlipper001.RampUp = fRampUp
    RightFlipper001.Elasticity = fElasticity
  End If

  If PlungerSound=1 and Plunger.position>=8 and ShowDT=False Then
    If Plunge = 0 then PlaySoundAt "CHPlungerPull", Plunger
    Plunge = 1
  End If
  If PlungerSound=1 and Plunger.position<=3 and ShowDT=False Then
    If Plunge = 1 then PlaySoundAt "CHPlungerRelease", Plunger
    Plunge = 0
  End If
TB.text = "plunger.position=" & round(plunger.position,2)

' PlungerCheck = plunger.Position()
' If PlungerSound = 1 Then
'   TB.text = "Max = " & PlungerMax & " Min = " & PlungerMin
'   TB1.text = plungerLoop
'   If PlungerLoop = 0 Then PlungerCheck1 = PlungerCheck
'   If PlungerLoop = 1 And PlungerCheck > PlungerCheck1 + .5 and PlungerCheck > 4 And PlungerPlaying = 0 And ShowDT = False then
'     PlaySoundAt "CHPlungerPull", Plunger
'     PlungerMax = PlungerCheck
'     PlungerPlaying = 1
'   ElseIf PlungerLoop = 1 And PlungerCheck1 > PlungerCheck + .1 and PlungerPlaying = 1 And ShowDT = False Then
'     PlaySoundAt "CHPlungerRelease", Plunger
'     PlungerMin = PlungerCheck
'     PlungerPlaying = 0
'     PlungerPause = 1
'   End If
'   PlungerLoop = PlungerLoop + 1
'   If PlungerPause = 0 Then
'     If PlungerLoop = 2 Then PlungerLoop = 0
'   Else
'     If PlungerLoop = 18 Then PlungerLoop = 0: PlungerPause = 0
'
'   End If
' End If
End Sub

dim LF : Set LF = New FlipperPolarity
dim RF : Set RF = New FlipperPolarity
dim LF1 : Set LF1 = New FlipperPolarity
dim RF1 : Set RF1 = New FlipperPolarity

InitPolarity

Sub InitPolarity()
  dim x, a : a = Array(LF, RF, LF1, RF1)
  for each x in a
    'safety coefficient (diminishes polarity correction only)
    x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
    x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1

    x.enabled = True
    x.TimeDelay = 69    '*****Important, this variable is an offset for the speed that the ball travels down the table to determine if the flippers have been fired
              'This is needed because the corrections to ball trajectory should only applied if the flippers have been fired and the ball is in the trigger zones.
              'FlipAT is set to GameTime when the ball enters the flipper trigger zones and if GameTime is less than FlipAT + this time delay then changes to velocity
              'and trajectory are applied.  If the flipper is fired before the ball enters the trigger zone then with this delay added to FlipAT the changes
              'to tragectory and velocity will not be applied.  Also if the flipper is in the final 20 degrees, changes to ball values will also not be applied.
              '"Faster" tables will need a smaller value while "slower" tables will need a larger value to give the ball more time to get to the flipper.
              'If this value is not set high enough the Flipper Velocity and Polarity corrections will NEVER be applied.
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
  LF1.Object = LeftFlipper001
  LF.EndPoint = EndPointLp  'you can use just a coordinate, or an object with a .x property. Using a couple of simple primitive objects
  LF1.EndPoint = EndPointLp001
  RF.Object = RightFlipper
  RF1.Object = RightFlipper001
  RF1.EndPoint = EndPointRp001
  RF.EndPoint = EndPointRp
End Sub

Sub AddPt(aStr, idx, aX, aY)  'debugger wrapper for adjusting flipper script in-game
  dim a : a = Array(LF, RF, LF1, RF1)
  dim x : for each x in a
    x.addpoint aStr, idx, aX, aY
  Next
End Sub

Sub TriggerLF_Hit()
  LFCount = gameTime
  If FlipperLength = 3 Then LF.Addball activeball
End Sub
Sub TriggerLF_UnHit()
  If FlipperLength = 3 Then LF.PolarityCorrect activeball
End Sub
Sub TriggerRF_Hit()
  RFCount = gameTime
  If FlipperLength = 3 Then RF.Addball activeball
End Sub
Sub TriggerRF_UnHit()
  If FlipperLength = 3 Then RF.PolarityCorrect activeball
End Sub
Sub TriggerLF1_Hit()
  If FlipperLength = 2 Then LF1.Addball activeball
End Sub
Sub TriggerLF1_UnHit() :
  If FlipperLength = 2 Then LF1.PolarityCorrect activeball
End Sub
Sub TriggerRF1_Hit()
  If FlipperLength = 2 Then RF1.Addball activeball
End Sub
Sub TriggerRF1_UnHit()
  If FlipperLength = 2 Then RF1.PolarityCorrect activeball
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

'***************This is flipperPolarity's addPoint Sub
Class FlipperPolarity
  Public Enabled
  Private FlipAt  'Timer variable (IE 'flip at 723,530ms...)
  Public TimeDelay  'delay before trigger turns off and polarity is disabled TODO set time!
  private Flipper, FlipperStart, FlipperEnd, LR, PartialFlipCoef
  Private Balls(20), balldata(20)

  dim PolarityIn, PolarityOut
  dim VelocityIn, VelocityOut
  dim YcoefIn, YcoefOut

  Public Sub Class_Initialize
    redim PolarityIn(0) : redim PolarityOut(0) : redim VelocityIn(0) : redim VelocityOut(0) : redim YcoefIn(0) : redim YcoefOut(0)
    Enabled = True: TimeDelay = 50 : LR = 1:  dim x : for x = 0 to uBound(balls) : balls(x) = Empty : set Balldata(x) = new spoofBall: next
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

  End Sub

'********Triggered by a ball hitting the flipper trigger area
  Public Sub AddBall(aBall) : dim x :
    for x = 0 to uBound(balls)
      if IsEmpty(balls(x)) then set balls(x) = aBall : exit sub :end if
    Next
  End Sub

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

'*********Used to rotate flipper since this is removed from the key down for the flippers
  Public Sub Fire()
    Flipper.RotateToEnd
    processballs
    FlipperOn
  End Sub

  Public Sub ProcessBalls() 'save data of balls in flipper range
    FlipAt = GameTime
    dim x : for x = 0 to uBound(balls)
      if not IsEmpty(balls(x) ) then balldata(x).Data = balls(x)
    Next
    PartialFlipCoef = ((Flipper.StartAngle - Flipper.CurrentAngle) / (Flipper.StartAngle - Flipper.EndAngle))  '% of flipper swing
    PartialFlipCoef = abs(PartialFlipCoef-1) 'corrects for negative flipper angles
    if abs(Flipper.currentAngle - Flipper.EndAngle) < 20 Then 'last 20 degrees of swing is not dealt with
      PartialFlipCoef = 0
    End If
  End Sub

'***********gameTime is a global variable of how long the game has progressed in ms
'***********This function lets the table know if the flipper has been fired
  Private Function FlipperOn()
'   TB.text = gameTime & ":" & (FlipAT + TimeDelay) & ":" & LFCount ' ******MOVE TB into view WHEN THIS FLIPPER FUNCTIONALITY IS ADDED TO A NEW TABLE TO CHECK IF THE TIME DELAY IS LONG ENOUGH*****
    if gameTime < FlipAt + TimeDelay then FlipperOn = True
  End Function  'Timer shutoff for polaritycorrect

'***********This is turned on when a ball leaves the flipper trigger area
  Public Sub PolarityCorrect(aBall)
    if FlipperOn() then 'don't run this if the flippers are at rest
'           tb.text = "In"
      dim tmp, BallPos, x, IDX, Ycoef : Ycoef = 1
      dim teststr : teststr = "Cutoff"
      tmp = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
      if tmp < 0.1 then 'if real ball position is behind flipper, exit Sub to prevent stucks  'Disabled 1.03, I think it's the Mesh that's causing stucks, not this
      end if

      'y safety Exit
      if aBall.VelY > -8 then 'if ball going down then remove the ball
        RemoveBall aBall
        exit Sub
      end if
      'Find balldata. BallPos = % on Flipper
      for x = 0 to uBound(Balls)
        if aBall.id = BallData(x).id AND not isempty(BallData(x).id) then
          idx = x
          BallPos = PSlope(BallData(x).x, FlipperStart, 0, FlipperEnd, 1)
          if ballpos > 0.65 then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)        'find safety coefficient 'ycoef' data
        end if
      Next

      'Velocity correction
      if not IsEmpty(VelocityIn(0) ) then
'       tb.text = "Vel corr"
        Dim VelCoef
        if IsEmpty(BallData(idx).id) and aBall.VelY < -12 then 'if tip hit with no collected data, do vel correction anyway
          if PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1) > 1.1 then 'adjust plz
            VelCoef = LinearEnvelope(5, VelocityIn, VelocityOut)
            if partialflipcoef < 1 then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)
            if Enabled then aBall.Velx = aBall.Velx*VelCoef'VelCoef
            if Enabled then aBall.Vely = aBall.Vely*VelCoef'VelCoef
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
        Set a(aCount) = aArray(x) 'Set creates an object in VB
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

'**********Takes in more than one array and passes them to ShuffleArray
Sub ShuffleArrays(aArray1, aArray2, offset)
  ShuffleArray aArray1, offset
  ShuffleArray aArray2, offset
End Sub

'**********Calculate ball speed as hypotenuse of velX/velY triangle
Function BallSpeed(ball) 'Calculates the ball speed
    BallSpeed = SQR(ball.VelX^2 + ball.VelY^2 + ball.VelZ^2)
End Function

'**********Calculates the value of Y for an input x using the slope intercept equation
Function PSlope(Input, X1, Y1, X2, Y2)  'Set up line via two points, no clamping. Input X, output Y
  dim x, y, b, m : x = input : m = (Y2 - Y1) / (X2 - X1) : b = Y2 - m*X2
  Y = M*x+b
  PSlope = Y
End Function

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

'*********This sets up the rubbers:
dim RubbersD
Set RubbersD = new Dampener  'Makes a Dampener Class Object
RubbersD.name = "Rubbers"

'cor bounce curve (linear)
'for best results, try to match in-game velocity as closely as possible to the desired curve
RubbersD.addpoint 0, 0, 0.935 '0.96 'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 1, 3.77, 0.935 '0.96
RubbersD.addpoint 2, 5.76, 0.942 '0.967 'dont take this as gospel. if you can data mine rubber elasticitiy, please help!
RubbersD.addpoint 3, 15.84, 0.874
RubbersD.addpoint 4, 56, 0.64 'there's clamping so interpolate up to 56 at least

dim SleevesD : Set SleevesD = new Dampener  'this is just rubber but cut down to 85%...
SleevesD.name = "Sleeves"
SleevesD.CopyCoef RubbersD, 0.85

'**********Class for dampener section of nfozzy's code
Class Dampener
  Public Print, debugOn 'tbpOut.text
  public name, Threshold  'Minimum threshold. Useful for Flippers, which don't have a hit threshold.
  Public ModIn, ModOut
  Private Sub Class_Initialize : redim ModIn(0) : redim Modout(0): End Sub

  Public Sub AddPoint(aIdx, aX, aY)
    ShuffleArrays ModIn, ModOut, 1 : ModIn(aIDX) = aX : ModOut(aIDX) = aY : ShuffleArrays ModIn, ModOut, 0
  End Sub

  public sub Dampen(aBall)
    if threshold then if BallSpeed(aBall) < threshold then exit sub end if end if
    dim RealCOR, DesiredCOR, str, coef
'               Uses the LinearEnvelope function to calculate the correction based upon where it's value sits in relation

'               to the addpoint parameters set above.  Basically interpolates values between set points in a linear fashion
    DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )

'                Uses the function BallSpeed's value at the point of impact/the active ball's velocity which is constantly being updated
'               RealCor is always less than 1
    RealCOR = BallSpeed(aBall) / cor.ballvel(aBall.id)

'               Divides the desired CoR by the real COR to make a multiplier to correct velocity in x and y
    coef = desiredcor / realcor
'   tb.text = "Coef = " & coef

'               Applies the coef to x and y velocities
' Thalamus - patched :     aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
    aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
  End Sub

'***********This Sub sets the values for Sleeves (or any other future objects) to 85% (or whatever is passed in) of Posts
  Public Sub CopyCoef(aObj, aCoef) 'alternative addpoints, copy with coef
    dim x : for x = 0 to uBound(aObj.ModIn)
      addpoint x, aObj.ModIn(x), aObj.ModOut(x)*aCoef
    Next
  End Sub
End Class

'*****************************Generates cor.ballVel for dampener
Sub RDampen_Timer() ' 1 ms timer always on
  CoR.Update
End Sub

'*********CoR is Coefficient of Restitution defined as "how much of the kinetic energy remains for the objects
'to rebound from one another vs. how much is lost as heat, or work done deforming the objects
dim cor : set cor = New CoRTracker

Class CoRTracker

  public ballvel

  Private Sub Class_Initialize : redim ballvel(0) : End Sub

  Public Sub Update() 'tracks in-ball-velocity
    dim str, b, allBalls, highestID :
    allBalls = getballs

    for each b in allballs
      if b.id >= HighestID then highestID = b.id
    Next

    if uBound(ballvel) < highestID then redim ballvel(highestID)  'set bounds

    for each b in allballs
      ballvel(b.id) = BallSpeed(b)
    Next
  End Sub
End Class

'********Interpolates the value for areas between the low and upper bounds sent to it
Function LinearEnvelope(xInput, xKeyFrame, yLvl)
  dim y 'Y output
  dim L 'Line
  dim ii : for ii = 1 to uBound(xKeyFrame)  'find active line
    if xInput <= xKeyFrame(ii) then L = ii : exit for : end if
  Next
  if xInput > xKeyFrame(uBound(xKeyFrame) ) then L = uBound(xKeyFrame)  'catch line overrun
  Y = pSlope(xInput, xKeyFrame(L-1), yLvl(L-1), xKeyFrame(L), yLvl(L) )

  'clamp 2.0
  if xInput <= xKeyFrame(lBound(xKeyFrame) ) then Y = yLvl(lBound(xKeyFrame) )  'Clamp lower
  if xInput >= xKeyFrame(uBound(xKeyFrame) ) then Y = yLvl(uBound(xKeyFrame) )  'Clamp upper

  LinearEnvelope = Y
End Function

'************************************************************************
'                         Ball Control
'************************************************************************

Dim Cup, Cdown, Cleft, Cright, Zup, contball, contballinplay, ControlBall, bcboost
Dim bcvel, bcyveloffset, bcboostmulti

bcboost = 1 'Do Not Change - default setting
bcvel = 4 'Controls the speed of the ball movement
bcyveloffset = -0.1 'Offsets the force of gravity to keep the ball from drifting vertically on the table, should be negative
bcboostmulti = 3 'Boost multiplier to ball veloctiy (toggled with the B key)

Sub BallControl_Timer()
    If Contball and ContBallInPlay then
        If Cright = 1 Then
            ControlBall.velx = bcvel*bcboost
          ElseIf Cleft = 1 Then
            ControlBall.velx = -bcvel*bcboost
          Else
            ControlBall.velx=0
        End If
        If Cup = 1 Then
            ControlBall.vely = -bcvel*bcboost
          ElseIf Cdown = 1 Then
            ControlBall.vely = bcvel*bcboost
          Else
            ControlBall.vely = bcyveloffset
        End If
        If Zup = 1 Then
            ControlBall.velz = bcvel*bcboost
    Else
      ControlBall.velz = -bcvel*bcboost
        End If
    End If
End Sub

'******* for ball control script
Sub endControl_Hit()
    contBallInPlay = False
End Sub


Sub Table1_MusicDone()

End Sub

Sub Table1_Paused()

End Sub

Sub Table1_UnPaused()

End Sub

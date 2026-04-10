'****************************************************************
'
'        Gulfstream (Williams 1973)
'          Script by Scottacus
'             v 2.06
'              August 2022
'
'   DOF config
'   101 Left Flipper, 102 Right Flipper
'   103 Left sling, 104 Right sling
'   105 Bumper1,  106 Bumper2, 107 Bumper3
'   108 - 114 Upper Rollovers
'   115 - 116 InLane Rollovers
'   117 & 118 OutLane Rollovers
'   119 Center  Target
'   120 - 121 Top Targets
'   122 Right Kicker, 123 Left Kicker
'   125 Ball Release
'   127 credit light
'   128 Knocker
'   131 - 133 Roll Over Buttons
'   153 Chime1-10s, 154 Chime2-100s, 155 Chime3-1000s
'   160 Ball in Launch Lane
'
'   Code Flow
'                  EndGame
'                   ^
'   Start Game -> New Game -> Check Continue -> Release Ball -> Drain -> Score Bouns -> Advance Player -> Next Ball
'                   ^                                     |
'                 EndGame = True <---------------------------------------------------------------

' Ball Control Subroutine developed by: rothbauerw
'   Press "c" during play to activate, the arrow keys control the ball
'**********************************************************************************************************************

Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

Const cGameName = "gulfstream_1973"
Const cOptions = "gulfStream_1973_v2.06.txt"
Const HSFileName = "Gulfstream (Williams 1973)"

Dim vrOption
'*************************** VR Room****************************************************
'The following two lines of code are two ways of controling which mode the table should
'be run in (Cab/DT, VR Minimal, VR Room). With the first line uncommented, the software
'will attempt to determine if you are running VPVR and if you have set it up to VR Enable.
'If you have not done this then it will run the table in DT/Cab mode.  Should you experience
'problems with this, just comment out the line "GetAskkToTurnOn", uncomment the following
'line of code and change the value of vrOption to the desired setting.

'GetAskToTurnOn
vrOption = 0 ' 0 - VR Room off, 1 - VR Cab On, 2- VR Cab and Room On

'*************************** Glass Scratches *******************************************
Const VRGlassScratches = 0  ' set this to 1 to turn on VR glass scratches
If vrOption > 0 and VRGlassScratches > 0 then GlassImpurities.visible = true

'*************************** PinballY Settings *****************************************
'Set this variable to 1 to save a PinballY High Score file to your Tables Folder
'this will let the Pinball Y front end display the high scores when searching for tables
'0 = No PinballY High Scores, 1 = Save PinballY High Scores
Const cPinballY = 1

'************ New Ball Shadow Code *******************************************************************************
Const DynamicBallShadowsOn = 1    '0 = no dynamic ball shadows, 1 = enable dynamic ball shadow
Const AmbientBallShadowOn = 1       '0 = no regular ball shadow   1 = regular ball shadow on

'**************************** Auto Detect VR Mode **************************************
Sub GetAskToTurnOn
  Dim objShell, AskToTurnOn, AutoVR, entry
  AskToTurnOn = "HKCU\Software\Visual Pinball\VP10\PlayerVR\AskToTurnOn"

  Set objShell = CreateObject("WScript.Shell")

  On Error Resume Next

  AutoVr = objShell.RegRead(AskToTurnOn)

  If Err.Number <> 0 Then
    Err.Clear
    vrOption = 0
  Else
    Err.Clear
    If AutoVR = 0 Then
      vrOption = 2
    Else
      vrOption = 0
    End If
  End If
End Sub

'***********************************************************************************************

Dim tablewidth: tablewidth = Table1.width
Dim tableheight: tableheight = Table1.height

Sub FrameTimer_Timer()
  FlipperVisualUpdate       'update flipper shadows and primitives
  If DynamicBallShadowsOn Or AmbientBallShadowOn Then DynamicBSUpdate 'update ball shadows
End Sub

' The game timer interval is 10 ms
Sub GameTimer_Timer()
  Cor.Update            'update ball tracking (this sometimes goes in the RDampen_Timer sub)
' RollingUpdate         'update rolling sounds
' DoDTAnim            'handle drop target animations
' DoSTAnim            'handle stand up target animations
End Sub

'****** End New ball Shadow Code..  More at bottom of script ****************************************************************************

' RuleSet = 1 ' uncomment to collect S with I and P with C with any of the rollovers (Liberal/Conservative jacks in backbox)
  CornerSetting = 3 'Change this to any number from 1 to 3 for pay out if getting all four corners without a tictactoe, this is very hard to do
'***************************************************************************************

Dim balls
Dim replays
Dim maxPlayers
Dim players
Dim player
Dim credit
Dim score(6)
Dim hScore(6)
Dim sReel(4)
Dim state
Dim tilt
Dim matchNumber
Dim i,j, f, ii, Object, Light, x, y, z, obj
Dim freePlay
Dim ballsize,BallMass
ballSize = 50
ballMass = (Ballsize^3)/125000
Dim BIP : BIP = 0
Dim options
Dim chime
Dim pfOption
Dim replayEB
Dim RuleSet
Dim lutValue

'Desktop reel pans,  Valid values: "Lpan", "Mpan" and "Rpan"
Const dtcrReel = "Lpan"
Const dts1Reel = "Rpan"
Const dts2Reel = "Rpan"
Const dts3Reel = "Rpan"
Const dts4Reel = "Rpan"

'Backglass reel pans,  Valid values: "Lpan", "Mpan" and "Rpan"
Const bgcrReel = "Rpan"
Const bgs1Reel = "Lpan"
Const bgs2Reel = "Rpan"
Const bgs3Reel = "Lpan"
Const bgs4Reel = "Rpan"

Sub Table1_init
  LoadEM
  maxPlayers = 1

  For x = 1 to maxPlayers
    Set sReel(x) = EVAL("scoreReel" & x)
  Next

  Player=1
  LoadHighScore

  ballsInPlay = 0

  If highScore(0) = "" Then highScore(0) = 100000
  If highScore(1) = "" Then highScore(1) = 80000
  If highScore(2) = "" Then highScore(2) = 60000
  If highScore(3) = "" Then highScore(3) = 50000
  If HighScore(4) = "" Then highScore(4) = 45000
  If matchNumber = "" Then matchNumber = 4
  If ShowDT = True Then pfOption = 1
  If pfOption = "" Then pfOption = 1
  If initial(0,1) = "" Then
    initial(0,1) = 19: initial(0,2) = 5: initial(0,3) = 13
    initial(1,1) = 1: initial(1,2) = 1: initial(1,3) = 1
    initial(2,1) = 2: initial(2,2) = 2: initial(2,3) = 2
    initial(3,1) = 3: initial(3,2) = 3: initial(3,3) = 3
    initial(4,1) = 4: initial(4,2) = 4: initial(4,3) = 4
  End If
  If credit = "" Then credit = 0
  If freePlay = "" Then freePlay = 1
  If balls = "" Then balls = 5
  If chime = "" Then chime = 0
  If replayEB = "" Then replayEB = 0
  If lutValue = "" Then lutValue = 0

  replaySettings

  firstBallOut = 0
  updatePostIt
  dynamicUpdatePostIt.enabled = 1
  TiltReel.setValue(1)
  If vrOption > 0 Then vrTilt.visible = 1
  If vrOption > 0 Then vrCredit
  creditReel.setValue (credit)

  If ShowDT = True Then
    For each object in backdropstuff
    Object.visible = 1
    Next
  End If

  If ShowDt = False Then
    For each object in backdropstuff
    Object.visible = 0
    Next
  End If

  Table1.ColorGradeImage = "LUT" & lutValue

  tilt = False
  state = False
  gameState
  bootTable.enabled = 1

  overflowReel.setValue(0)

  If freePlay = 1 Then
    Coincard.image = "CoinCard" & replayEB & balls
  Else
    Coincard.image = "CoinCard" & replayEB & balls
  End If

  InstructCard.image = "instructionCard" & replayEB
  Drain.CreateSizedBallWithMass Ballsize/2, BallMass

End Sub

'***********KeyCodes
Dim enableInitialEntry, firstBallOut
Sub Table1_KeyDown(ByVal keycode)

  If enableInitialEntry = True Then enterIntitals(keycode)

  If keycode = addCreditKey Then
    playFieldSound "coinin",0,Drain,1
    addCredits
    End If

    If keycode = startGameKey Then
    StartButton.y = 1885
    If enableInitialEntry = False and operatormenu = 0 and backGlassOn = 1 Then
      If freePlay = 1 and players < MaxPlayers and firstBallOut = 0 Then startGame
      If freePlay = 0 and credit > 0 and players < MaxPlayers and firstBallOut = 0 Then
        credit = credit - 1
        creditReel.setValue (credit)
        If B2SOn Then controller.B2SSetCredits credit
        If showDT = False Then PlayReelSound "Reel5", bgcrReel Else PlayReelSound "Reel5", dtcrReel
        If credit < 1 Then
'         creditLight.image = "layer4off"
          If B2SOn Then DOF 127, DOFOff
          CreditLight1.state = 0
        End If
        If B2SOn Then
          If freePlay = 0 and credit < 1 Then DOF 127, DOFOff
        End If
        startGame
      End If
    End If
  End If

  If keycode = rightFlipperKey Then
    If vrOption > 0 Then VRFlipperButtonRight.X = 2086 - 8
  End If

  If keycode = leftFlipperKey and contball = 0 Then
    If vrOption > 0 Then VRFlipperButtonLeft.X = 2122 + 8
  End If

  If keycode = PlungerKey Then
    plunger.PullBack
    playFieldSound "plungerpull", 0, plunger, 1
    If vrOption > 0 Then
      TimerVRPlunger.Enabled = True
      TimerVRPlunger1.Enabled = False
      PinCab_Shooter.Y = -444
    End If
  End If

  If tilt = False and state = True Then
  If keycode = leftFlipperKey and contball = 0 Then
    FlipperActivate LeftFlipper, LFPress
    lf.fire
    playFieldSound "FlipUpL", 0, leftFlipper, 1
    If B2SOn Then DOF 101,DOFOn
    playFieldSound "FlipBuzzL", -1, leftFlipper, 1
  End If

  If keycode = RightFlipperKey and contball = 0 Then
    FlipperActivate RightFlipper, RFPress
    rf.fire
    playFieldSound "FlipUpR", 0, RightFlipper,1
    If B2SOn Then DOF 102,DOFOn
    playFieldSound "FlipBuzzR", -1, RightFlipper,1
  End If

  If keycode = leftTiltKey Then
    Nudge 90, 2
    checkTilt
  End If

  If keycode = rightTiltKey Then
    Nudge 270, 2
    checkTilt
  End If

  If keycode = centerTiltKey Then
    Nudge 0, 2
    checkTilt
  End If
  End If

  If keycode = RightMagnaSave and state = False Then
    lutText.visible = 1
        playFieldSound "metalHitHigh", 0, SoundPointScoreMotor, 0.2
    lutValue = lutValue + 1
    If lutValue > 9 Then lutValue = 0
    setLUT
  End If

  If keycode = LeftMagnaSave and state = False Then
    lutText.visible = 1
        playFieldSound "metalHitHigh", 0, SoundPointScoreMotor, 0.2
    lutValue = lutValue - 1
    If lutValue < 0 Then lutValue = 9
    setLUT
  End If

    If keycode = leftFlipperKey and state = False and operatorMenu = 0 and enableInitialEntry = 0 Then
        operatorMenuTimer.Enabled = true
    End If

     If keycode = leftFlipperKey and state = False and operatorMenu = 1 Then
    options = options + 1
    If vrOption = 0 Then If options = 4 Then options = 5
    If showDt = True Then If options = 5 Then options = 6 'skips non DT options
        If options = 7 Then options = 0
    optionMenu.visible = True
        playFieldSound "target", 0, SoundPointScoreMotor, .2
        Select Case (Options)
            Case 0:
                optionMenu.image = "FreeCoin" & freePlay
            Case 1:
                optionMenu.image = balls & "Balls"
      Case 2:
        OptionMenu.image = "replayEB" & replayEB
      Case 3:
        If showDT = False Then optionMenu.image = "UnderCab"
        If showDT = True Then OptionMenu.visible = False
        optionMenu.image = "UnderCab"
        optionMenu1.visible = 1
        optionMenu1.image = "Sound" & pfOption
        optionMenu2.visible = 1
        optionMenu2.image = "SoundChange"
        Select Case (pfOption)
          Case 1: speaker1.visible = 1: speaker2.visible = 1: speaker3.visible = 0: speaker4.visible = 0
          Case 2: speaker5.visible = 1: speaker6.visible = 1: speaker1.visible = 0: speaker2.visible = 0
          Case 3: speaker1.visible = 1: speaker2.visible = 1: speaker3.visible = 1: speaker4.visible = 1: speaker5.visible = 0: speaker6.visible = 0
        End Select
      Case 4:
        For x = 1 to 6
          EVAL("Speaker" & x).visible = 0
        Next
        optionMenu1.visible = 0
        OptionMenu2.visible = 0
        optionMenu.image = "VR" & vrOption
      Case 5:
        optionMenu2.visible = 0
        For x = 1 to 6
          EVAL("Speaker" & x).visible = 0
        Next
        optionMenu1.visible = 1
        optionMenu1.image = "DOF"
        optionMenu.image = "Chime" & chime
      Case 6:
        optionMenu2.visible = 0
        For x = 1 to 6
          EVAL("Speaker" & x).visible = 0
        Next
        optionMenu1.visible = 0
        optionMenu.image = "SaveExit"
    End Select
    End If

     If keycode = RightFlipperKey and state = False and operatorMenu = 1 Then
        playFieldSound "metalHitHigh", 0, SoundPointScoreMotor, 0.2
      Select Case (options)
    Case 0:
            If freePlay = 0 Then
                freePlay = 1
'       creditLight.image = "layer4"
        CreditLight1.state = True
        If B2SOn Then DOF 127, DOFOn
              Else
                freePlay = 0
        If credit < 1 Then
'         creditLight.image = "layer4off"
          If B2SOn Then DOF 127, DOFOff
          CreditLight1.state = 0
        End If
        If credit > 0 Then
'         creditLight.image = "layer4"
          If B2SOn Then DOF 127, DOFOn
          CreditLight1.state = 1
        End If
            End If
      If freePlay = 1 Then
        Coincard.image = "CoinCard" & replayEB & balls
      Else
        Coincard.image = "CoinCard" & replayEB & balls
      End If
      optionMenu.image = "FreeCoin" & freePlay
      InstructCard.image = "instructionCard" & replayEB
        Case 1:
            If Balls = 3 Then
                Balls = 5
              Else
                Balls = 3
            End If
      replaySettings
            OptionMenu.image = Balls & "Balls"
      Coincard.image = "CoinCard" & replayEB & balls
    Case 2:
'     0 = add a ball, 1 = replay, 2 = novelty
      replayEB = replayEB + 1
      If replayEB = 1 Then replayEB = 2
      If replayEB = 3 Then replayEB = 0
      OptionMenu.image = "replayEB" & replayEB
      InstructCard.image = "instructionCard" & replayEB
      Coincard.image = "CoinCard" & replayEB & balls
            If Balls = 5 Then
        Replay(0) = 200000
        Replay(1) = 400000
        Replay(2) = 600000
              Else
        Replay(0) = 100000
        Replay(1) = 300000
        Replay(2) = 600000
            End If
    Case 3:
      pfOption = pfOption + 1
      optionMenu1.visible = 1
      If showDt = True and pfOption = 2 Then pfOption = 3
      If pfOption = 4 Then pfOption = 1
      optionMenu1.image = "Sound" & pfOption
      Select Case (pfOption)
        Case 1: speaker1.visible = 1: speaker2.visible = 1: speaker3.visible = 0: speaker4.visible = 0
        Case 2: speaker5.visible = 1: speaker6.visible = 1: speaker1.visible = 0: speaker2.visible = 0
        Case 3: speaker1.visible = 1: speaker2.visible = 1: speaker3.visible = 1: speaker4.visible = 1: speaker5.visible = 0: speaker6.visible = 0
      End Select
    Case 4:
      If vrOption = 1 Then
        vrOption = 2
        For Each object in vrRoom: object.visible = 1: Next
      Else
        vrOption = 1
        For Each object in vrRoom: object.visible = 0: Next
      End If
      optionMenu.image = "VR" & vrOption
        Case 5:
            If chime = 0 Then
                chime= 1
        If B2SOn Then DOF 442,DOFPulse
              Else
                chime = 0
        pts10
            End If
      optionMenu.image = "Chime" & chime
        Case 6:
            operatorMenu = 0
            saveHighScore
      dynamicUpdatePostIt.enabled = 1
      optionMenu.image = "FreeCoin" & freePlay
            optionMenu1.visible = 0
      optionMenu.visible = 0
      optionsMenu.visible = 0
      replaySettings
    End Select
    End If


  If Keycode = mechanicalTilt Then
    tilt = True
    TiltReel.setValue(1)
    If vrOption > 0 Then vrTilt.visible = 1
    If B2SOn Then controller.B2SSetTilt 1
    turnOff
  End If

    If keycode = 46 Then' C Key
       stopSound "FlipBuzzLA"
       stopSound "FlipBuzzLB"
       stopSound "FlipBuzzLC"
       stopSound "FlipBuzzLD"
       stopSound "FlipBuzzRA"
       stopSound "FlipBuzzRB"
       stopSound "FlipBuzzRC"
       stopSound "FlipBuzzRD"
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

'************************Start Of Test Keys****************************
  If keycode = 30 Then
  End If

  If Keycode = 31 Then
  End If

  If keycode = 33 Then  '"f"
  End If
'************************End Of Test Keys****************************
End Sub

dim TestChr1, TestChr2, TestChr3, TestChr

Sub Table1_KeyUp(ByVal keycode)

  If keycode = rightFlipperKey Then
    If vrOption > 0 Then VRFlipperButtonRight.X = 2086
  End If

  If keycode = leftFlipperKey and contball = 0 Then
    If vrOption > 0 Then VRFlipperButtonLeft.X = 2122
  End If

  If keycode = StartGameKey Then
    StartButton.y = 1890
  End If

  If keycode = plungerKey Then
    If vrOption > 0 Then
      TimerVRPlunger.Enabled = False
      TimerVRPlunger1.Enabled = True
      PinCab_Shooter.Y = -444
    End If
    plunger.Fire
    playFieldSound "plungerfire", 0, plunger, 1
  End If

    If keycode = leftFlipperKey Then
        operatorMenuTimer.Enabled = False
    End If

   If tilt = False and state = True Then
    If keycode = leftFlipperKey and contball = 0 Then
      FlipperDeActivate LeftFlipper, LFPress
      LeftFlipper.eosTorqueAngle = EOSA
      LeftFlipper.eosTorque = EOST
      LeftFlipper.RotateToStart
      playFieldSound "FlipDownL", 0, leftFlipper, 1
      If B2SOn Then DOF 101,DOFOff
      stopSound "FlipBuzzLA"
      stopSound "FlipBuzzLB"
      stopSound "FlipBuzzLC"
      stopSound "FlipBuzzLD"
    End If

    If keycode = RightFlipperKey and contball = 0 Then
      FlipperDeActivate RightFlipper, RFPress
      RightFlipper.eosTorqueAngle = EOSA
      RightFlipper.eosTorque = EOST
      RightFlipper.rotateToStart
      playFieldSound "FlipDownR", 0, RightFlipper, 1
      If B2SOn Then DOF 102,DOFOff
      stopSound "FlipBuzzRA"
      stopSound "FlipBuzzRB"
      stopSound "FlipBuzzRC"
      stopSound "FlipBuzzRD"
    End If
   End If

    If keycode = 203 then cLeft = 0' Left Arrow

    If keycode = 200 then cUp = 0' Up Arrow

    If keycode = 208 then cDown = 0' Down Arrow

    If keycode = 205 then cRight = 0' Right Arrow

    If keycode = 52 Then Zup = 0' Period
End Sub

'************** Table Boot
Dim backGlassOn
Dim bootCount:bootCount = 0
Sub bootTable_Timer
  bootCount = bootCount + 1
  If bootCount = 1 Then
    If vrOption > 0 Then vrBackGlass.image = "gsOn"
    For each Light in GIlights:light.state = 0:Next
    If B2SOn Then
      Controller.B2SSetGameOver 35,1
      Controller.B2SSetTilt 33,1
      Controller.B2SSetBallInPlay 32,0
      controller.B2SSetCredits credit
      controller.B2SSetData 80,1
      If Credit > 0 Then DOF 127, DOFOn
      If FreePlay = 1 Then DOF 127, DOFOn
    End If
    If freePlay = 0 and credit = 0 Then
      creditlight1.state = 0
    Else
      creditlight1.state = 1
    End If

    For x = 1 to maxPlayers
      If B2SOn Then controller.B2SSetScorePlayer x, score(x)
      sReel(x).setvalue(score(x))
    Next
    Light5.state = 1
    lightOn = 5
    For each Light in GIlights:Light.state = 1: Next
    backGlassOn = 1
    me.enabled = False
  End If
  dim xx
  For each xx in filaments: xx.blenddisablelighting = 120: Next
  For each xx in bulbs: xx.blenddisablelighting = 12: Next

End Sub

'***********Replay Settings
Sub replaySettings
    If Balls = 5 Then
    Replay(0) = 53000
    Replay(1) = 75000
    Replay(2) = 97000
    Else
    Replay(0) = 45000
    Replay(1) = 67000
    Replay(2) = 89000
    End If
End Sub

'***********Operator Menu
Dim operatormenu
Sub operatorMenuTimer_Timer
    If optionMenu.visible = False Then playFieldSound "target", 0, SoundPointScoreMotor, .2
  options = 0
    operatorMenu = 1
  dynamicUpdatePostIt.enabled = 0
  updatePostIt
  options = 0
    optionsMenu.visible = True
    optionMenu.visible = True
  optionMenu.image = "FreeCoin" & freePlay
End Sub

'***********Start Game
Dim ballInPlay, bonusScore
Grid = 5
Sub startGame
  If state = False Then
    ballInPlay = balls
    If B2SOn Then
      controller.B2SSetGameOver 0
    End If
    bipDisplay
    dynamicUpdatePostIt.enabled = 0
    updatePostIt
    tilt = False
    state = True
    TicTacToe = False
    gameState
    players = 1
    for x = 1 to maxPlayers
      score(x) = (score(x) mod 100000)
    Next
    For x = 1 to 9
      EVAL("GridLight" & x).state = 0
      If x < 8 Then EVAL("SpecialLight" & x).state = 0
    Next
    SpecialLight.state = 0
    light10.state = 0
    light11.state = 0
    light12.state = 0
    light13.state = 0
    corners = 0
    EVAL("Light" & Grid).state = 1
    newGame
    If vrOption > 0 Then
      vrBip
      vrCredit
      vrGameOver.visible = 0
    End If
  Else If state = True and players < maxPlayers and ballinPlay = 1 Then
      players = players + 1
      If vrOption > 0 Then vrCredit
    End If
  End If
End Sub

'*********New Game
Dim endGame
Sub newGame
  player = 1
    endGame = 0
  If B2SOn Then
    For x = 91 to 95
      controller.B2SSetData x, 0
      controller.B2SSetData 25, 0
    Next
  End If
  overflowReel.setValue(0)
  If vrOption > 0 Then
    vrMatchBg.image = "m0"
    vrOverflow.visible = 0
  End If
  gameState
  resetTable
  resetReel.enabled = True
End Sub

'**********Check if Game Should Continue
Dim relBall, rep(4)
Sub checkContinue
  If endGame = 1 Then
    turnOff
    stopSound "FlipBuzzLA"
    stopSound "FlipBuzzLB"
    stopSound "FlipBuzzLC"
    stopSound "FlipBuzzLD"
    stopSound "FlipBuzzRA"
    stopSound "FlipBuzzRB"
    stopSound "FlipBuzzRC"
    stopSound "FlipBuzzRD"
    state = False
    BIPReel.setValue(0)
    gameState
    DynamicUpdatePostIt.enabled = 1
    sortScores
    checkHighScores
    firstBallOut = 0
    players = 0
    For x = 1 to 4
      rep(x) = 0
      repAwarded(x) = 0
    Next
    saveHighScore
    GameOverReel.setValue(1)
    If B2SOn Then
      controller.B2SSetGameOver 35,1
      controller.B2SSetballinplay 32, 0
    End If
    If vrOption > 0 Then
      vrBip
      vrGameOver.visible = 1
      vrBipBg.image = "bip0"
    End If
  Else
    relBall = 1   'variable to tell score motor routine that this is a ball release run
    scoreMotor5.enabled = 1
    If vrOption > 0 Then vrBip
    resetTable
  End If
End Sub

'***************Drain and Release Ball
Dim drainActive
Sub drain_Hit()
  playFieldSound "Drain", 0, drain, 0.001
  RightSlingShot.collidable = 1
  LeftSlingShot.collidable = 1
  repAwarded(Player) = 0
  drainActive = 1
' drain.DestroyBall 'this is used for multiball tables to avoid timing issues with balls moving in the trough
  ballsInPlay = ballsInPlay - 1
  For each Light in BumperLights:light.state = 0: Next
  If ballsInPlay = 0 Then scoreBonus
End Sub

Dim launched, BallsInPlay
Sub ReleaseBall
  playFieldSound "FastKickIntoLaunchLane", 0, Drain, 0.5
  If B2SOn Then DOF 125,DOFPulse
  ballsInPlay = 1
  drain.kick 70,10
  launched = 0
  bipDisplay
  drainActive = 0
End Sub

'**********Check if Scoring Bonus is True
Dim bonusFlag
Sub scoreBonus
  advancePlayers
End Sub

'**********Display Ball In Play
Sub bipDisplay
  If ballInPlay > 10 Then BallInPlay = 10
  bipReel.setValue(ballInPlay)
  If B2SOn Then controller.B2SSetBallinPlay 32, BallinPlay
  If vrOption > 0 Then vrBip
End Sub


'**********Advance Players
Sub advancePlayers
  BonusFlag = 0
  nextBall
End Sub

'**********Next Ball
Sub nextBall
    If Tilt = True Then
    tilt = False
    tiltReel.setValue(0)
    If vrOption > 0 Then vrTilt.visible = 0
    If B2SOn Then
      controller.B2SSetTilt 33,0
    End If
    End If
  resetTable
  If Player = 1 then ballinPlay = ballinPlay - 1
' If player = 4 then BallinPlay = 5 'used for match testing

  If ballinPlay = 0 then
    endGame = 1
    checkContinue
  Else
    If state = True Then checkContinue
  End If
End Sub

'************Game State Check
  dim xx

Sub gameState
  If state = False Then
    If vrOption > 0 Then vrGameOver.visible = 1
    GameOverReel.setValue(1)
    If B2SOn then controller.B2SSetGameOver 35,1
    stopSound "FlipBuzzLA"
    stopSound "FlipBuzzLB"
    stopSound "FlipBuzzLC"
    stopSound "FlipBuzzLD"
    stopSound "FlipBuzzRA"
    stopSound "FlipBuzzRB"
    stopSound "FlipBuzzRC"
    stopSound "FlipBuzzRD"
  Else
    GameOverReel.setValue(0)
    If vrOption > 0 Then vrGameOver.visible = 0
    TiltReel.setValue(0)
    If vrOption > 0 Then vrTilt.visible = 0
    If B2SOn Then
      controller.B2SSetTilt 33,0
      controller.B2SSetGameOver 35,0
    End If
  End If
End Sub

'*************Ball in Launch Lane on Plunger Tip
Dim ballREnabled, relGateHit
Sub ballHome_hit
  ballREnabled = 1
  relGateHit = 0
  Set controlBall = ActiveBall
    contBallInPlay = True
End Sub

'*************Ball off of Plunger Tip
Sub ballHome_unhit
End Sub

'******* for ball control script
Sub EndControl_Hit()
    contBallinPlay = False
End Sub

'************Check if Ball Out of Launch Lane
Dim ballInLane
Sub ballsInPlay_hit
  If ballREnabled = 1 Then
    ballREnabled = 0
    ballInLane = False
  End If
  firstBallOut = 1
End Sub

'**********Reset Table
Sub resetTable
    For f = 1 to 3
    EVAL("Bumper" & f).hashitevent = 1
  Next
  bumpersOff
End Sub

'************** Bumpers
dim bump1step, bump2step, bump3step
Sub Bumpers_hit(Index)
  If ballControlBlock = True Then Exit Sub
  If Tilt = False Then
    Select Case (Index)
      Case 0: PlayFieldSound "PopBump", 0, bumper1, 1
          If B2SOn Then DOF 105,DOFPulse
        If BumperLight1.state = 0 then
          Addscore 10
        Else
          AddScore 100
        End If
      Case 1: PlayFieldSound "PopBump", 0, bumper2, 1
          If B2SOn Then DOF 106,DOFPulse
        If BumperLight1.state = 0 then
          Addscore 10
        Else
          AddScore 100
        End If
      Case 2: PlayFieldSound "PopBump", 0, bumper3, 1
          If B2SOn Then DOF 107,DOFPulse
        If BumperLight1.state = 0 then
          Addscore 10
        Else
          AddScore 100
        End If
    End Select
  End If
End Sub

Sub bumpersOn
  bumpermesh1.image = "render2uvlit"
  bumpermesh2.image = "render2uvlit"
  bumpermesh3.image = "render2uvlit"
  bumperskirt1.image = "render2uvlit"
  bumperskirt2.image = "render2uvlit"
  bumperskirt3.image = "render2uvlit"
End Sub

Sub bumpersOff
  bumpermesh1.image = "render2uv"
  bumpermesh2.image = "render2uv"
  bumpermesh3.image = "render2uv"
  bumperskirt1.image = "render2uv"
  bumperskirt2.image = "render2uv"
  bumperskirt3.image = "render2uv"
End Sub

Dim SpecialLanes
'*************** Triggers
Sub Triggers_hit(Index)
  If Tilt = False Then
    If Index < 7 Then
      AddScore 500
      EVAL ("SpecialLight" & (Index + 1)).state = 1
      If RuleSet = 1 Then
        If Index = 0 or Index = 4 Then
          SpecialLight1.state = 1
          SpecialLight5.state = 1
        End If
        If Index = 1 or Index = 3 Then
          SpecialLight2.state = 1
          SpecialLight4.state = 1
        End If
      End If
      For x = 1 to 7
        If EVAL("SpecialLight" & x).state = 1 then SpecialLanes = SpecialLanes + 1
      Next
      If B2SOn Then DOF (Index +108), DOFPulse
      If SpecialLanes = 7 Then
        Light12.state = 1
        Light13.state = 1
      End If
    End If
    Select Case Index
      Case 7: AddScore 1000: If B2SOn Then DOF 115, DOFPulse
      Case 8: InLanes: If B2SOn Then DOF 118, DOFPulse
      Case 9: InLanes: If B2SOn Then DOF 117, DOFPulse
      Case 10: AddScore 1000: If B2SOn Then DOF 116, DOFPulse
    End Select
  End If
  SpecialLanes = 0
  wireNumber = index
  wireAnimation.enabled = 1
End Sub

Sub InLanes
  If Tilt = False Then
    If Light12.state = 1 Then
      AddScore 5000
      inLaneSpecial
    Else
      AddScore 500
    End If
  End If
End Sub

'************** Wire RollOver Animation
Dim wire, wireNumber
Sub wireAnimation_Timer
  wire = wire + 1
  Select Case wire
    Case 1: EVAL ("wire" & wireNumber + 1).transz = -10
    Case 2: EVAL ("wire" & wireNumber + 1).transz = -4
    Case 3: EVAL ("wire" & wireNumber + 1).transz = -1
    Case 4: EVAL ("wire" & wireNumber + 1).transz = 0
        wire = 0
        wireAnimation.enabled = 0
  End Select
End Sub

'************** Rollover Buttons
Dim ROButton
Sub RollOvers_Hit(Index)
  ROButton = Index
  RollOverAnimation.enabled = 1
  If Tilt = False Then
    AddScore 100
    If B2SOn Then DOF (Index + 131), DOFPulse
  End If
End Sub

Dim Button
Sub RollOverAnimation_Timer
  Button = Button +1
  Select Case Button
    Case 1: EVAL ("RollOverButton" & ROButton).transz = -1
    Case 2: EVAL ("RollOverButton" & ROButton).transz = -3
    Case 3: EVAL ("RollOverButton" & ROButton).transz = 3
    Case 4: EVAL ("RollOverButton" & ROButton).transz = 1
    Button = -1
    RollOverAnimation.enabled = 0
  End Select
End Sub

'*************** Targets
Dim sideSpecial, TicTacToe, suIndex
Sub targets_hit(index)
  standUpAnimation.enabled = 1
  If ballControlBlock = True Then Exit Sub
  scoreMotorLoop = 0
  suIndex = index + 1
  If tilt = False Then
    Select Case Index
      Case 0:
        For each Light in BumperLights:light.state = 1: Next
        bumpersOn
        If B2SOn Then DOF 121, DOFPulse
        AddScore 100
      Case 1:
        For each Light in BumperLights:light.state = 1: Next
        bumpersOn
        If B2SOn Then DOF 120, DOFPulse
        AddScore 100
      Case 2:
        If TicTacToe = False Then
          EVAL("gridLight" & LightOn).state = 1
          CheckGrid
        End If
        If B2SOn Then DOF 119, DOFPulse
        If SpecialLight.state = 1 Then
          AddScore 5000
        Else
          addScore 1000
        End If
    End Select
  End If
End Sub

Dim standUpBend
Sub standUpAnimation_timer
  standUpBend = standUpBend + 1
  Select Case StandUpBend
    Case 1: EVAL ("TargetMesh" & suIndex).rotx = -5
    Case 2: EVAL ("TargetMesh" & suIndex).rotx = -5
    Case 3: EVAL ("TargetMesh" & suIndex).rotx = 5
    Case 4: EVAL ("TargetMesh" & suIndex).rotx = 5
        standUpBend = 0
        standUpAnimation.enabled = 0
  End Select
End Sub


'*************Check Grid
Dim NewTicTacToe, Grid
Sub CheckGrid
  For x = 1 to 7 step 3
    If EVAL("gridLight" & x).state = 1 Then
      If EVAL("gridLight" & (x + 1)).state = 1 And EVAL("gridLight" & (x + 2)).state = 1 Then TicTacToe = True
    End If
  Next
  For x = 1 to 3
    If EVAL("gridLight" & x).state = 1 Then
      If EVAL("gridLight" & (x+ 3)).state = 1  and EVAL("gridLight" & (x + 6)).state = 1 Then TicTacToe = True
    End If
  Next
  If gridLight5.state = 1 Then
    If gridLight1.state = 1 and gridlight9.state = 1 Then TicTacToe = True
  End If
  If gridlight5.state = 1 Then
    If gridLight3.state = 1 and gridLight7.state = 1 Then TicTacToe = True
  End If
  If TicTacToe = True Then
    Light10.state = 1: Light11.state = 1: SpecialLight.state = 1
    For x = 1 to 9
      EVAL("Light" & x).state = 0
    Next
    If NewTicTacToe = False Then
      NewTicTacToe = True
      AddScore 5000
    End If
  End If
  If gridLight1.state = 1 and gridLight2.state = 0 and gridLight3.state = 1 and gridLight4.state = 0 and gridLight5.state = 0 and gridLight6.state = 0 and gridLight7.state = 1 and gridLight8.state = 0 and gridLight9.state = 1 Then CornerCollect.enabled = 1
End Sub

'**************Corner Collect
Dim Corners, CornerSetting, CornerLoop
Sub CornerCollect_Timer
  If Corners = 0 Then
    If replayEB = 0 Then BallinPlay = BallinPlay + 1: If vrOption > 0 Then vrBip
    If replayEB = 1 Then credit = credit + 1: If credit > 15 Then credit = 15: creditReel.setValue (credit): If B2SOn Then controller.B2SSetCredits credit
    If replayEB = 2 Then addScore 10000 * 1
    If BallinPlay > 10 Then BallinPlay = 10
    bipDisplay
    knock
  End If
  CornerLoop = CornerLoop + 1
  If CornerLoop = CornerSetting Then
    CornerLoop = 0
    Corners = 1
    CornerCollect.enabled = 0
  End If
End Sub

Sub addCredits
  addCredit = 1
  If B2SOn Then DOF 127, DOFOn
  CreditLight1.state = 1
  scoreMotor5.enabled = 1
End Sub

'**************Tic Tac Toe
Sub SpecialGrid
  For x = 10 to 11
    EVAL("Light" & x).State = 0
  Next
  For x = 1 to 9
    EVAL("GridLight" & x).state = 0
  Next
  SpecialLight.state = 0
  NewTicTacToe = False
  TicTacToe = False
  gridSpecial
  AdvanceGrid
End Sub


Sub AddABall
  BallinPlay = BallinPlay + 1
  If BallinPlay > 10 Then BallinPlay = 10
  bipDisplay
  If vrOption > 0 Then vrBip
End Sub

Dim LightOn
Sub AdvanceGrid
  Grid = Grid + 1
  If Grid > 10 Then Grid = 1
  LightOn = Grid
  If Grid <10 Then
    EVAL("Light" & Grid).state = 1
  Else
    Light5.state = 1
    LightOn = 5
  End If

  Select Case Grid
    Case 1: Light5.state = 0
    Case 2: Light1.state = 0
    Case 3: Light2.state = 0
    Case 4: Light3.state = 0
    Case 5: Light4.state = 0
    Case 6: Light5.state = 0
    Case 7: Light6.state = 0
    Case 8: Light7.state = 0
    Case 9: Light8.state = 0
    Case 10: Light9.state = 0
  End Select

End Sub

'************ Kickers
Dim kickStep,  kickerBall
Sub kickBall(kball, kangle, kvel, kvelz, kzlift)
  dim rangle
  rangle = 3.14 * (kangle - 90) / 180
  kball.z = kball.z + kzlift
  kball.velz = kvelz
  kball.velx = cos(rangle)*kvel
  kball.vely = sin(rangle)*kvel
End Sub

kickStep = 1
Sub Kickers_hit(Index)
  settleTimer.enabled = 1
End Sub

Dim saucerCount
Sub settleTimer_Timer
  saucerCount = saucerCount + 1
  If SaucerCount > 1 Then
    If kickerL.ballCntOver > 0 Or kickerR.ballCntOver > 0 Then kickerActive
    saucerCount = 0
    settleTimer.enabled = 0
  End If
End Sub

Dim SaucerFlag
Sub kickerActive
  set kickerBall = activeBall
  If kickerL.ballCntOver > 0 Then
    playFieldSound "SaucerIn", 0, kickerL, 1
  Else
    playFieldSound "SaucerIn", 0, kickerR, 1
  End If
  scoreMotorLoop = 0
  kickStep = 0
  If tilt = False Then
    If light11.state = 0 Then
      addScore 500
    Else
      addScore 5000
      SpecialGrid
    End If
  End If
  saucerFlag = 1
  scoremotor.enabled = 1
End Sub

Sub kickerTimer_timer
    Select Case kickStep
        Case 7: If kickerL.ballCntOver > 0 Or kickerR.ballCntOver > 0 Then
          For Each obj in kickerArm: obj.ObjRotX = 12: Next
          If kickerL.ballCntOver > 0 Then
            kickBall kickerBall, 142 - INT(RND*6), 13 + INT(RND*2), 5, 30
            PlayFieldSound "SaucerKick",0, kickerL, 1
            If B2SOn Then DOF 122, DOFPulse
          Else
            kickBall kickerBall, 230 - INT(RND*6), 13 + INT(RND*2), 5, 30
            PlayFieldSound "SaucerKick",0, kickerR, 1
            If B2SOn Then DOF 123, DOFPulse
          End If
        End If
        Case 8: For Each obj in kickerArm: obj.ObjRotX = -45: Next
        Case 9: For Each obj in kickerArm: obj.ObjRotX = -45: Next
        Case 10: For Each obj in kickerArm: obj.ObjRotX = 24: Next
        Case 11: For Each obj in kickerArm: obj.ObjRotX = 12: Next
        Case 12: For Each obj in kickerArm: obj.ObjRotX = 0: Next: kickerTimer.Enabled = 0
    End Select
    kickStep = kickStep + 1
End Sub

''**********Sling Shot Animations
'' Rstep and Lstep  are the variables that increment the animation
''****************
Dim lStep, rStep
Sub rightSlingShot_Slingshot
  If ballControlBlock = True Then Exit Sub
  If Tilt = False Then
    playfieldSound "SlingShot", 0, SoundPoint13, 1
    If B2SOn Then DOF 104, DOFPulse
    rSling.Visible = 0
    rSling1.Visible = 1
    sling1.Rotx = 10
    rStep = 0
    rightSlingShot.TimerEnabled = 1
    addScore(10)
  End If
End Sub

Sub rightSlingShot_Timer
    Select Case rStep
        Case 3:rSLing1.Visible = 0:rSLing2.Visible = 1:sling1.Rotx = 0
        Case 4:rSLing2.Visible = 0:rSLing.Visible = 1:rightSlingShot.TimerEnabled = 0
    End Select
    rStep = rStep + 1
End Sub
'
Sub leftSlingShot_Slingshot
  If ballControlBlock = True Then Exit Sub
  If Tilt = False Then
    playfieldSound "SlingShot", 0, SoundPoint12, 1
    If B2SOn Then DOF 103, DOFPulse
    lSling.Visible = 0
    lSling1.Visible = 1
    sling2.Rotx = 10
    lStep = 0
    leftSlingShot.TimerEnabled = 1
    addScore(10)
  End If
End Sub

Sub leftSlingShot_Timer
    Select Case lStep
        Case 3:lSLing1.Visible = 0:lSLing2.Visible = 1:sling2.Rotx = 0
        Case 4:lSLing2.Visible = 0:lSLing.Visible = 1:leftSlingShot.TimerEnabled = 0
    End Select
    lStep = lStep + 1
End Sub

Sub knock
  If B2SOn = True Then DOF 128, DOFPulse Else playSound "knocker"
End Sub

Dim matchOutPut, initialize
'***************Match Williams Style
Sub match
  Select Case matchNumber
    Case 0: matchOutPut = 70
    Case 1: matchOutPut = 10
    Case 2: matchOutPut = 60
    Case 3: matchOutPut = 00
    Case 4: matchOutPut = 40
    Case 5: matchOutPut = 90
    Case 6: matchOutPut = 50
    Case 7: matchOutPut = 20
    Case 8: matchOutPut = 80
    Case 9: matchOutPut = 30
  End Select
  If B2SOn Then
    If matchNumber = 3 Then
      controller.B2SSetMatch 100
    Else
      controller.B2SSetMatch 34, matchOutPut
    End If
  End If
  matchTxt.text = matchOutPut
  If vrOption > 0 Then vrMatch

  If initialize = 0 Then
    For x = 1 to players
      If matchOutPut = (score(x) mod 100) Then
        increaseCredits
        knock
        If B2SOn Then DOF 151, DOFPulse
      End If
    Next

  End If
End Sub

'***************Score Motor Run one full rotation
Dim scoreMotorCount, addCredit
Sub scoremotor5_Timer
  scoreMotorCount = scoreMotorCount + 1
  playFieldSound "ScoreMotorSingleFire", 0, SoundPointScoreMotor, 0. 'need to set location of score motor under the PF
  If scoremotorCount = 5 Then
    If relBall = 1 Then
      relBall = 0
      releaseBall
    End If
    If addCredit = 1 Then
      increaseCredits
      addCredit = 0
    End If
    scoreMotorCount = 0
    scoreMotor5.enabled = 0
  End If
End Sub

Sub increaseCredits
  credit = credit + 1
  If credit > 15 Then credit = 15
  If showDT = False Then PlayReelSound "Reel5", bgcrReel Else PlayReelSound "Reel5", dtcrReel
  CreditLight1.state = 1
  creditReel.setValue (credit)
  If vrOption > 0 Then vrCredit
  If B2SOn Then
    controller.B2SSetCredits credit
    If credit > 0 Then DOF 127, DOFOn
  End If
End Sub

'**************Score Motor Routine
Dim bellRing, scoreMotorLoop, kickerUp, KickBonus, motorOn
Sub ScoreMotor_timer
  scoreMotorLoop = scoreMotorLoop + 1
  If scoreMotorLoop < 6 Then playFieldSound "ScoreMotorSingleFire", 0, SoundPointScoreMotor, 0.2
  motorOn = 1

'  These Flags are passed by scores with multiple of 10, 100 or 1000
  If flag = 1 or flag10 = 1 or flag100 = 1 or flag1000 = 1 or SaucerFlag = 1 Then
    Select Case scoreMotorLoop
      Case 1: totalUp Point
      Case 2: If bellRing > 1 Then totalUp point
      Case 3: If bellRing > 2 Then totalUp point
      Case 4: If bellRing > 3 Then totalUp point
      Case 5: If bellRing > 4 Then totalUp point
          If SaucerFlag = 1 Then kickerTimer.enabled = 1
      Case 6: scoreMotorLoop = 0
          flag1 = 0
          flag10 = 0
          flag100 = 0
          flag1000 = 0
          SaucerFlag = 0
          motorOn = 0
          scoreMotor.enabled = 0 'originally had If drainActive = 0 Then scoreMotor.Enabeled = 0 but this casued a final ball drain error in PG
    End Select
  End If

' BonusFlag lets the sub know that the bonus score is being paid out
' The bonus is paid out on the 2nd and 4th positions of the score motor, the 6th position is the index reel notch
  If bonusFlag = 1 Then
    Select Case scoreMotorLoop
      Case 1:
      Case 2: If bonusScore > 0 Then
            If tilt = False Then totalUp 1000
            If doubleBonus.State = 0 Then
              bonusScore = bonusScore - 1
              If bonusScore > 0 Then
                If bonusScore < 10  Then
                  If tilt = False Then EVAL("bonus" & BonusScore).State =  1
                  EVAL("bonus" & BonusScore + 1).State = 0
                End If
                If bonusScore = 10 Then bonus1.State = 0
                If bonusScore > 10 Then
                  If tilt = False Then EVAL("bonus" & BonusScore - 10).state = 1
                  EVAL("bonus" & BonusScore - 9).state = 0
                End If
              End If
              If bonusScore = 0 Then bonus1.state = 0
            End If
          End If
          If bonusScore > 7 and bonusScore < 15 and tilt = False Then AlternateEB
          If bonusScore = 7 then extraBall1.state = 0
      Case 3:
      Case 4: If bonusScore > 0 Then
            If tilt = False Then totalUp 1000
            bonusScore = bonusScore -1
              If bonusScore > 0 Then
                If bonusScore < 10 Then
                  If tilt = False Then EVAL("bonus" & BonusScore).State =  1
                  EVAL("bonus" & BonusScore + 1).State = 0
                End If
                If bonusScore = 10 Then bonus1.State = 0
                If bonusScore > 10 Then
                  If tilt = False Then EVAL("bonus" & BonusScore - 10).state = 1
                  EVAL("bonus" & BonusScore - 9).state = 0
                End If
              End If
              If bonusScore = 0 Then bonus1.state = 0
          End If
          If bonusScore > 7 and bonusScore < 15 and tilt = False Then AlternateEB
          If bonusScore = 7 then extraBall1.state = 0
      Case 5:
      Case 6: scoreMotorLoop = 0
        If bonusScore < 1 Then
          scoreBonus
          motorOn = 0
          scoreMotor.enabled = 0
        End If
    End Select
  End If
End Sub

'***************Scoring Routine
Dim flag1, flag10, flag100, flag1000, flag10000, fl, point, point10

Sub pts10
  If pts10Timer.enabled = False Then
    pts10Timer.enabled = True
    playFieldSound "Chime10", 0, soundPoint13, 1
  End If
End Sub

Sub pts100
  If pts100Timer.enabled = False Then
    pts100Timer.enabled = True
    playFieldSound "Chime100", 0, soundPoint13, 1
  End If
End Sub

Sub pts1000
  If pts1000Timer.enabled = False Then
    pts1000Timer.enabled = True
    playFieldSound "Chime1000", 0, soundPoint13, 1
  End If
End Sub

Dim pts10Block, pts100Block, pts1000Block, intvl
intvl = 40
pts10Timer.interval = intvl
pts100Timer.interval = intvl
pts1000Timer.interval = intvl

Sub pts10Timer_Timer
  pts10Block = pts10Block + 1
  If pts10Block = 1 Then
    pts10Block = 0
    pts10Timer.enabled = False
  End If
End Sub

Sub pts100Timer_Timer
  pts100Block = pts100Block + 1
  If pts100Block = 1 Then
    pts100Block = 0
    pts100Timer.enabled = False
  End If
End Sub

Sub pts1000Timer_Timer
  pts1000Block = pts1000Block + 1
  If pts1000Block = 1 Then
    pts1000Block = 0
    pts1000Timer.enabled = False
  End If
End Sub

Sub ballControlBlockTimer_Timer
  ballControlBlock = False
  ballControlBlockTimer.enabled = 0
End Sub

Dim ballControlBlock
Sub addScore(points)
  If ballControlBlock = True Then Exit Sub
  If contball = 1 Then
    ballControlBlock = True
    BallControlBlockTimer.enabled = 1
  End If
  reelDone(player) = 0

    If points > 9 and points <100 Then
'     Number Matching, decrement the match unit for each 10 point score roll up from 0 to 9
      point10 =  (points / 10)
      If matchNumber >= point10 Then
        matchNumber = matchNumber - point10
      Else
        point10 = point10 - matchNumber
        matchNumber = 10 - point10
      End If

      If points > 10 Then
        bellRing = (points / 10)
        point = 10: flag10 = 1: scoreMotor.enabled = 1
      Else
        If motorOn = 0 Then totalUp 10

      End If
      Exit Sub
    End If

    If points > 99 and points < 1000 Then
      If points > 100 Then
        bellRing = (points / 100)
        point = 100: flag100 = 1: scoreMotor.enabled = 1
      Else
        If motorOn = 0 Then totalUp 100
      End If
      Exit Sub
    End If

    If Points > 999 and points < 10000 Then
      If points > 1000 Then
        bellRing = (points / 1000)
        flag1000 = 1: point = 1000: scoreMotor.enabled = 1
      Else
        If motorOn = 0 Then totalUp 1000
      End If
      Exit Sub
    End If

    If Points > 9999 Then
      If points > 10000 Then
        bellRing = (points / 1000)
        flag10000 = 1: point = 10000: scoreMotor.enabled = 1
      Else
        If motorOn = 0 Then totalUp 10000
      End If
    End If
End Sub

Dim replayX,  replay(7), repAwarded(5), block
Sub totalUp(points)
  If tilt = True Then Exit Sub
  If pts10Timer.enabled = False and points = 10 Or pts100Timer.enabled = False and points = 100 Or pts1000Timer.enabled = False and points = 1000 Then block = 0 Else block = 1

  If B2SOn And showDT = False Then
    If Player = 1 and block = 0 Then PlayReelSound "Reel1", bgs1Reel
  End If

  If showDT = True Then
    If Player = 1 and block = 0 Then PlayReelSound "Reel1", dts1Reel
  End If

  If flag10 = 1 or points = 10 Then
    If chime = 0 Then
      pts10
    Else
      If B2SOn Then DOF 142,DOFPulse
    End If
    If Score(Player) mod 100 = 0 and TicTacToe = False Then AdvanceGrid
  End If

  If flag100 = 1 or points = 100 Then
    If chime = 0 Then
      pts100
    Else
      If B2SOn Then DOF 143,DOFPulse
    End If
    If TicTacToe = False Then AdvanceGrid
  End If

  If flag1000 = 1 or points = 1000 Then
    If chime = 0 Then
      pts1000
    Else
      If B2SOn Then DOF 143,DOFPulse
    End If
  End If

  If flag10000 = 1 or points = 10000 Then
    If chime = 0 Then
      pts1000
    Else
      If B2SOn Then DOF 143,DOFPulse
    End If
  End If

  If bonusFlag = 1 Then
    If chime = 0 Then
      pts1000
    Else
      If B2SOn Then DOF 143,DOFPulse
    End If
  End If

  If block = 0 Then score(Player) = score(player) + points
  If block = 0 Then sReel(Player).addvalue(points)
  If vrOption > 0 Then vrScore(player)

  If score(player) >= 100000 Then
    overflowReel.setValue(1)
    If B2SOn Then controller.B2SSetData 25, 1
    If vrOption > 0 Then vrOverflow.visible = 1
  End If

  If B2SOn Then controller.B2SSetScorePlayer player, score(player)

  If score(player) > replay(rep(player)) or score(player) = replay(rep(player)) Then
    If rep(player) > 2 Then Exit Sub
    Special
  End If
End Sub

Sub Special
  If replayEB = 0 Then BallinPlay = BallinPlay + 1: If vrOption > 0 Then vrBip
  If replayEB = 1 Then increaseCredits
  If replayEB = 2 Then score(1) = score(1) + 10000

  sReel(Player).addvalue(10000)
  If vrOption > 0 Then vrScore(player)

  If score(player) >= 100000 Then
    overflowReel.setValue(1)
    If B2SOn Then controller.B2SSetData 25, 1
    If vrOption > 0 Then vrOverflow.visible = 1
  End If

  rep(player) = rep(player) + 1
  knock
  bipDisplay
End Sub

Sub gridSpecial
  If replayEB = 0 Then BallinPlay = BallinPlay + 1: If vrOption > 0 Then vrBip
  If replayEB = 1 Then increaseCredits
  If replayEB = 2 Then addScore 10000
  knock
  bipDisplay
  End Sub

Sub inLaneSpecial
  Light12.State = 0
  Light13.State = 0
  For x = 1 to 7
    EVAL("SpecialLight" & x).state = 0
  Next
  If replayEB = 0 Then BallinPlay = BallinPlay + 1: If vrOption > 0 Then vrBip
  If replayEB = 1 Then increaseCredits
  If replayEB = 2 Then addScore 10000
  knock
  bipDisplay
End Sub

'***************Tilt
Dim tiltSens, tiltPenalty
'**** Set tiltPenalty; 0 = loose current ball / 1 = end game
tiltPenalty = 0

Sub checkTilt
  If tilttimer.enabled = True Then
    tiltSens = tiltSens + 1
    If tiltSens = 3 Then
    tilt = True
    TiltReel.setValue(1)
    If vrOption > 0 Then vrTilt.visible = 1
        If B2SOn Then controller.B2SSetTilt 33,1
        If B2SOn Then controller.B2SSetdata 1, 0
    turnOff
   End If
  Else
   tiltSens = 0
   tilttimer.enabled = True
  End If
End Sub

Sub tilttimer_Timer()
  tilttimer.enabled = False
End Sub

'************Reset Reels

'This Sub looks at each individual digit in each players score and sets them in an array RScore.  If the value is >0 and <9
'then the players score is increased by one times the position value of that digit (ie 1 * 1000 for the 1000's digit)
'If the value of the digit is 9 then it subtracts 9 times the postion value of that digit (ie 9*100 for the 100's digit)
'so that the score is not rolled over and the next digit in line gets incremented as well (ie 9 in the 10's positon gets
'incremented so the 100's position rolls up by one as well since 90 -> 100).  Lastly the RScore array values get incremented
'by one to get ready for the next pass.

Dim rScore(4,5), resetLoop, test, playerTest, resetFlag, reelFlag, reelStop, reelDone(4), SJWreelDone(4)
Sub countUp
  For playerTest = 1 to maxPlayers
    For test = 0 to 4
      rScore(playerTest,test) = Int(score(playerTest)/10^test) mod 10
    Next
  Next

  For playerTest = 1 to maxPlayers
    For x = 0 to 4
      If rscore(playerTest, x) > 0 And rscore(playerTest, x) < 9 Then score(playerTest) = score(playerTest) + 10^x
      If rScore(playerTest, x) = 9 Then score(playerTest) = score(playerTest) - (9 * 10^x)
      If rScore(playerTest, x) > 0 Then rScore(playerTest, x) = rScore(playerTest, x) + 1
      If rScore(playerTest, x) = 10 Then rScore(playerTest, x) = 0
    Next
  Next

  If score(1) = 0 Then
    reelFlag = 1
    For i = 1 to maxPlayers
      score(i) = 0
      rep(i) = 0
      repAwarded(i) = 0
    Next
  End If
End Sub

'This Sub sets each B2S reel or Desktop reels to their new values and then plays the score motor sound each time and the
'reel sounds only if the reels are being stepped

Sub updateReels
  For playerTest = 1 to maxPlayers
    If B2SOn and showDT = False and reelDone(playerTest) = 0 and reelStop = 0 Then
      controller.B2SSetScorePlayer playerTest, score(playerTest)
      If reelDone(1) = 0 and playerTest = 1 Then PlayReelSound "Reel1", bgs1Reel
      If score(playerTest) = 0 Then reelDone(playerTest) = 1
    End If

    If showDT = True and reelDone(playerTest) = 0 and reelStop = 0 Then
      sReel(playerTest).setvalue (score(playerTest))
      If reelDone(1) = 0 and playerTest = 1 Then playReelSound "Reel1", dts1Reel
      If score(playerTest) = 0 Then reelDone(playerTest) = 1
    End If

    If vrOption > 0 Then vrScore(playerTest)
  Next
  playfieldSound "scoreMotorSingleFire", 0, soundPointScoreMotor, 0.2
  If reelFlag = 1 Then reelStop = 1

End Sub

'This Timer runs a loop that calls the CountUp and UpdateReels routines to step the reels up five times and Then
'check to see if they are all at zero during a two loop pause and then step them the rest of the way to zero

Dim testFlag
Sub resetReel_Timer
  For x = 1 to 4
    score(x) = (score(x) Mod 100000)
  Next
  resetLoop = resetLoop + 1
  If resetLoop = 1 and score(1) = 0 and score(2) = 0 and score(3) = 0 and score(4) = 0 Then
    resetLoop = 0
    If testFlag = 0 Then releaseBall
    testFlag = 0
    resetReel.enabled = 0
    Exit Sub
  End If
  Select Case resetLoop
    Case 1: countUp: updateReels
    Case 2: countUp: updateReels
    Case 3: countUp: updateReels
    Case 4: countUp: updateReels
    Case 5: countUp: updateReels
    Case 6: If reelStop = 1 Then
          resetLoop = 0
          reelFlag = 0
          reelStop = 0
          If testFlag = 0 Then releaseBall
          testFlag = 0
          resetReel.enabled = 0
          Exit Sub
        End If

    Case 7:
    Case 8: countUp: updateReels
    Case 9: countUp: updateReels
    Case 10: countUp: updateReels
    Case 11: countUp: updateReels
    Case 12: countUp: updateReels:
      resetLoop = 0
      reelFlag = 0
      reelStop = 0
      If testFlag = 0 Then releaseBall
      testFlag = 0
      resetReel.enabled = 0
      Exit Sub
  End Select
End Sub

'************************************************Post It Note Section**************************************************************************
'***************Static Post It Note Update
Dim  hsY, shift, scoreMil, score100K, score10K, scoreK, score100, score10, scoreUnit
Dim hsInitial0, hsInitial1, hsInitial2
Dim hsArray: hsArray = Array("HS_0","HS_1","HS_2","HS_3","HS_4","HS_5","HS_6","HS_7","HS_8","HS_9","HS_Space","HS_Comma")
Dim hsiArray: hsIArray = Array("HSi_0","HSi_1","HSi_2","HSi_3","HSi_4","HSi_5","HSi_6","HSi_7","HSi_8","HSi_9","HSi_10","HSi_11","HSi_12","HSi_13","HSi_14","HSi_15","HSi_16","HSi_17","HSi_18","HSi_19","HSi_20","HSi_21","HSi_22","HSi_23","HSi_24","HSi_25","HSi_26")

Sub updatePostIt
  scoreMil = Int(highScore(0)/1000000)
  score100K = Int( (highScore(0) - (scoreMil*1000000) ) / 100000)
  score10K = Int( (highScore(0) - (scoreMil*1000000) - (score100K*100000) ) / 10000)
  scoreK = Int( (highScore(0) - (scoreMil*1000000) - (score100K*100000) - (score10K*10000) ) / 1000)
  score100 = Int( (highScore(0) - (scoreMil*1000000) - (score100K*100000) - (score10K*10000) - (scoreK*1000) ) / 100)
  score10 = Int( (highScore(0) - (scoreMil*1000000) - (score100K*100000) - (score10K*10000) - (scoreK*1000) - (score100*100) ) / 10)
  scoreUnit = (highScore(0) - (scoreMil*1000000) - (score100K*100000) - (score10K*10000) - (scoreK*1000) - (score100*100) - (score10*10) )

  pScore6.image = hsArray(scoreMil):If highScore(0) < 1000000 Then pScore6.image = hsArray(10)
  pScore5.image = hsArray(score100K):If highScore(0) < 100000 Then pScore5.image = hsArray(10)
  pScore4.image = hsArray(score10K):If highScore(0) < 10000 Then pScore4.image = hsArray(10)
  pScore3.image = hsArray(scoreK):If highScore(0) < 1000 Then pScore3.image = hsArray(10)
  pScore2.image = hsArray(score100):If highScore(0) < 100 Then pScore2.image = hsArray(10)
  pScore1.image = hsArray(score10):If highScore(0) < 10 Then pScore1.image = hsArray(10)
  pScore0.image = hsArray(scoreUnit):If highScore(0) < 1 Then pScore0.image = hsArray(10)
  If highScore(0) < 1000 Then
    PComma.image = hsArray(10)
  Else
    pComma.image = hsArray(11)
  End If
  If highScore(0) < 1000000 Then
    pComma1.image = hsArray(10)
  Else
    pComma1.image = hsArray(11)
  End If
  If highScore(0) > 999999 Then shift = 0 :pComma.transx = 0
  If highScore(0) < 1000000 Then shift = 1:pComma.transx = -10
  If highScore(0) < 100000 Then shift = 2:pComma.transx = -20
  If highScore(0) < 10000 Then shift = 3:pComma.transx = -30
  For hsY = 0 to 6
    EVAL("Pscore" & hsY).transx = (-10 * shift)
  Next
  initial1.image = hsIArray(initial(0,1))
  initial2.image = hsIArray(initial(0,2))
  initial3.image = hsIArray(initial(0,3))
End Sub

'***************Show Current Score
Sub showScore
  scoreMil = Int(highScore(activeScore(flag))/1000000)
  score100K = Int( (highScore(activeScore(flag)) - (scoreMil*1000000) ) / 100000)
  score10K = Int( (highScore(activeScore(flag)) - (scoreMil*1000000) - (score100K*100000) ) / 10000)
  scoreK = Int( (highScore(activeScore(flag)) - (scoreMil*1000000) - (score100K*100000) - (score10K*10000) ) / 1000)
  score100 = Int( (highScore(activeScore(flag)) - (scoreMil*1000000) - (score100K*100000) - (score10K*10000) - (scoreK*1000) ) / 100)
  score10 = Int( (highScore(activeScore(flag)) - (scoreMil*1000000) - (score100K*100000) - (score10K*10000) - (scoreK*1000) - (score100*100) ) / 10)
  scoreUnit = (highScore(activeScore(flag)) - (scoreMil*1000000) - (score100K*100000) - (score10K*10000) - (scoreK*1000) - (score100*100) - (score10*10) )

  pScore6.image = hsArray(scoreMil):If highScore(activeScore(flag)) < 1000000 Then pScore6.image = hsArray(10)
  pScore5.image = hsArray(score100K):If highScore(activeScore(flag)) < 100000 Then pScore5.image = hsArray(10)
  pScore4.image = hsArray(score10K):If highScore(activeScore(flag)) < 10000 Then pScore4.image = hsArray(10)
  pScore3.image = hsArray(scoreK):If highScore(activeScore(flag)) < 1000 Then pScore3.image = hsArray(10)
  pScore2.image = hsArray(score100):If highScore(activeScore(flag)) < 100 Then pScore2.image = hsArray(10)
  pScore1.image = hsArray(score10):If highScore(activeScore(flag)) < 10 Then pScore1.image = hsArray(10)
  pScore0.image = hsArray(scoreUnit):If highScore(activeScore(flag)) < 1 Then pScore0.image = hsArray(10)
  If highScore(activeScore(flag)) < 1000 Then
    pComma.image = hsArray(10)
  Else
    pComma.image = hsArray(11)
  End If
  If highScore(activeScore(flag)) < 1000000 Then
    pComma1.image = hsArray(10)
  Else
    pComma1.image = hsArray(11)
  End If
  If highScore(flag) > 999999 Then shift = 0 :pComma.transx = 0
  If highScore(activeScore(flag)) < 1000000 Then shift = 1:pComma.transx = -10
  If highScore(activeScore(flag)) < 100000 Then shift = 2:pComma.transx = -20
  If highScore(activeScore(flag)) < 10000 Then shift = 3:pComma.transx = -30
  For HSy = 0 to 6
    EVAL("Pscore" & hsY).transx = (-10 * shift)
  Next
  initial1.image = hsIArray(initial(activeScore(flag),1))
  initial2.image = hsIArray(initial(activeScore(flag),2))
  initial3.image = hsIArray(initial(activeScore(flag),3))
End Sub

'***************Dynamic Post It Note Update
Dim scoreUpdate, dHSx
Sub dynamicUpdatePostIt_Timer
  scoreMil = Int(highScore(scoreUpdate)/1000000)
  score100K = Int( (highScore(ScoreUpdate) - (scoreMil*1000000) ) / 100000)
  score10K = Int( (highScore(scoreUpdate) - (ScoreMil*1000000) - (Score100K*100000) ) / 10000)
  scoreK = Int( (highScore(scoreUpdate) - (scoreMil*1000000) - (score100K*100000) - (score10K*10000) ) / 1000)
  score100 = Int( (highScore(ScoreUpdate) - (scoreMil*1000000) - (score100K*100000) - (score10K*10000) - (scoreK*1000) ) / 100)
  score10 = Int( (highScore(ScoreUpdate) - (scoreMil*1000000) - (score100K*100000) - (score10K*10000) - (scoreK*1000) - (score100*100) ) / 10)
  scoreUnit = (highScore(ScoreUpdate) - (scoreMil*1000000) - (score100K*100000) - (score10K*10000) - (scoreK*1000) - (score100*100) - (score10*10) )

  pScore6.image = hsArray(ScoreMil):If highScore(scoreUpdate) < 1000000 Then pScore6.image = hsArray(10)
  pScore5.image = hsArray(Score100K):If highScore(scoreUpdate) < 100000 Then pScore5.image = hsArray(10)
  pScore4.image = hsArray(Score10K):If highScore(scoreUpdate) < 10000 Then pScore4.image = hsArray(10)
  pScore3.image = hsArray(ScoreK):If highScore(scoreUpdate) < 1000 Then pScore3.image = hsArray(10)
  pScore2.image = hsArray(Score100):If highScore(scoreUpdate) < 100 Then pScore2.image = hsArray(10)
  pScore1.image = hsArray(Score10):If highScore(scoreUpdate) < 10 Then pScore1.image = hsArray(10)
  pScore0.image = hsArray(ScoreUnit):If highScore(scoreUpdate) < 1 Then pScore0.image = hsArray(10)
  If highScore(scoreUpdate) < 1000 Then
    pComma.image = hsArray(10)
  Else
    pComma.image = hsArray(11)
  End If
  If highScore(scoreUpdate) < 1000000 Then
    pComma1.image = hsArray(10)
  Else
    pComma1.image = hsArray(11)
  End If
  If highScore(scoreUpdate) > 999999 Then shift = 0 :pComma.transx = 0
  If highScore(scoreUpdate) < 1000000 Then shift = 1:pComma.transx = -10
  If highScore(scoreUpdate) < 100000 Then shift = 2:pComma.transx = -20
  If highScore(scoreUpdate) < 10000 Then shift = 3:pComma.transx = -30
  For dHSx = 0 to 6
    EVAL("Pscore" & dHSx).transx = (-10 * shift)
  Next
  initial1.image = hsIArray(initial(scoreUpdate,1))
  initial2.image = hsIArray(initial(scoreUpdate,2))
  initial3.image = hsIArray(initial(scoreUpdate,3))
  scoreUpdate = scoreUpdate + 1
  If scoreUpdate = 5 then scoreUpdate = 0
End Sub

'***************Bubble Sort
Dim tempScore(2), tempPos(3), position(5)
Dim bSx, bSy
'Scores are sorted high to low with Position being the player's number
Sub sortScores
  For bSx = 1 to 4
    position(bSx) = bSx
  Next
  For bSx = 1 to 4
    For bSy = 1 to 3
      If score(bSy) < score(bSy+1) Then
        tempScore(1) = score(bSy+1)
        tempPos(1) = position(bSy+1)
        score(bSy+1) = score(bSy)
        score(bSy) = tempScore(1)
        position(bSy+1) = position(BSy)
        position(bSy) = tempPos(1)
      End If
    Next
  Next
End Sub

'*************Check for High Scores

Dim highScore(5), activeScore(5), hs, chX, chY, chZ, chIX, tempI(4), tempI2(4), flag, hsI, hsX
'goes through the 5 high scores one at a time and compares them to the player's scores high to
'if a player's score is higher it marks that postion with ActiveScore(x) and moves all of the other
' high scores down by one along with the high score's player initials
' also clears the new high score's initials for entry later
Sub checkHighScores
  For hs = 1 to maxPlayers              'look at 4 player scores
    For chY = 0 to 4                'look at all 5 saved high scores
      If score(hs) > highScore(chY) Then
        flag = flag + 1           'flag to show how many high scores needs replacing
        tempScore(1) = highScore(chY)
        highScore(chY) = score(hs)
        activeScore(hs) = chY       'ActiveScore(x) is the high score being modified with x=1 the largest and x=4 the smallest
        For chIX = 1 to 3         'set initals to blank and make temporary initials = to intials being modifed so they can move down one high score
          tempI(chIX) = initial(chY,chIX)
          initial(chY,chIX) = 0
        Next

        If chY < 4 Then           'check if not on lowest high score for overflow error prevention
          For chZ = chY+1 to 4      'set as high score one more than score being modifed (CHy+1)
            tempScore(2) = highScore(chZ) 'set a temporaray high score for the high score one higher than the one being modified
            highScore(chZ) = tempScore(1) 'set this score to the one being moved
            tempScore(1) = tempScore(2)   'reassign TempScore(1) to the next higher high score for the next go around
            For chIX = 1 to 3
              tempI2(chIX) = initial(chZ,chIX)  'make a new set of temporary initials
            Next
            For chIX = 1 to 3
              initial(chZ,chIX) = tempI(chIX)   'set the initials to the set being moved
              tempI(chIX) = tempI2(chIX)      'reassign the initials for the next go around
            Next
          Next
        End If
        chY = 4               'if this loop was accessed set CHy to 4 to get out of the loop
      End If
    Next
  Next
' Goto Initial Entry
    hsI = 1     'go to the first initial for entry
    hsX = 1     'make the displayed inital be "A"
    If flag > 0 Then  'Flag 0 when all scores are updated so leave subroutine and reset variables
      showScore
      playerEntry.visible = 1
      playerEntry.image = "Player" & position(Flag)
      initial(activeScore(flag),1) = 1  'make first inital "A"
      For chY = 2 to 3
        initial(activeScore(flag),chY) = 0  'set other two to " "
      Next
      For chY = 1 to 3
        EVAL("Initial" & chY).image = hsIArray(initial(activeScore(flag),chY))    'display the initals on the tape
      Next
      initialTimer1.enabled = 1   'flash the first initial
      dynamicUpdatePostIt.enabled = 0   'stop the scrolling intials timer
      knock
      enableInitialEntry = True
    End If
End Sub

'************Enter Initials Keycode Subroutine
Dim initial(6,5), initialsDone
Sub enterIntitals(keycode)
    If keyCode = leftFlipperKey Then
      hsX = hsX - 1           'HSx is the inital to be displayed A-Z plus " "
      If hsX < 0 Then hsX = 26
      If hsI < 4 Then EVAL("Initial" & hsI).image = hsIArray(hsX)   'HSi is which of the three intials is being modified
      playSound "metalHitHigh"
    End If
    If keycode = RightFlipperKey Then
      hsX = hsX + 1
      If hsX > 26 Then hsX = 0
      If hsI < 4 Then EVAL("Initial"& hsI).image = hsIArray(hsX)
      playSound "metalHitHigh"
    End If
    If keycode = startGameKey and initialsDone = 0 Then
      If hsI < 3 Then                 'if not on the last initial move on to the next intial
        EVAL("Initial" & hsI).image = hsIArray(hsX) 'display the initial
        initial(activeScore(flag), hsI) = hsX   'save the inital
        playSound "metalHitMedium"
        EVAL("InitialTimer" & hsI).enabled = 0    'turn that inital's timer off
        EVAL("Initial" & hsI).visible = 1     'make the initial not flash but be turn on
        initial(activeScore(flag),hsI + 1) = hsX  'move to the next initial and make it the same as the last initial
        EVAL("Initial" & hsI +1).image = hsIArray(hsX)  'display this intial
'       y = 1
        EVAL("InitialTimer" & hsI + 1).enabled = 1  'make the new intial flash
        hsI = hsI + 1               'increment HSi
      Else                    'if on the last initial then get ready yo exit the subroutine
        initial3.visible = 1          'make the intial visible
        playSound "metalHitMedium"
        initialTimer3.enabled = 0       'shut off the flashing
        initial(activeScore(flag),3) = hsX    'set last initial
        initialEntry              'exit subroutine
      End If
    End If
End Sub

'************Update Initials and see if more scores need to be updated
Dim eIX
Sub initialEntry
  playsound SoundFXDOF("Chime10",141,DOFPulse,DOFChimes)
  flag = flag - 1
  hsI = 1
  If flag < 0 Then flag = 0: Exit Sub
  If flag = 0 Then          'exit high score entry mode and reset variables
    initialsDone = 1        'prevents changes in intials while the highScoreDelay timer waits to finish
    players = 0
    For eIX = 1 to 4
      activeScore(eIX) = 0
      position(eIX) = 0
    Next
    For eIX = 1 to 3
      EVAL("InitialTimer" & eIX).enabled = 0
    Next
    playerEntry.visible = 0
    scoreUpdate = 0           'go to the highest score
    updatePostIt            'display that score
    highScoreDelay.enabled = 1
  Else
    showScore
    playerEntry.image = "Player" & position(flag)
    initial(activeScore(flag),1) = 1  'set the first initial to "A"
    For chY = 2 to 3
      initial(activeScore(flag),chY) = 0  'set the other two to " "
    Next
    For chY = 1 to 3
      EVAL("Initial" & chY).image = hsIArray(initial(activeScore(flag),chY))  'display the intials
    Next
    hsX = 1             'go to the letter "A"
    initialTimer1.enabled = 1   'flash the first intial
  End If
End Sub

'************Delay to prevent start button push for last initial from starting game Update
Sub highScoreDelay_timer
  highScoreDelay.enabled = 0
  enableInitialEntry = False
  initialsDone = 0
  saveHighScore
  For eIX = 1 to 3
    EVAL("InitialTimer" & eIX).enabled = 0
  Next
  dynamicUpdatePostIt.enabled = 1   'turn scrolling high score back on
End Sub

'************Flash Initials Timers
Sub initialTimer1_Timer
  y = y + 1
  If y > 1 Then y = 0
  If y = 0 Then
    initial1.visible = 1
  Else
    initial1.visible = 0
  End If
End Sub

Sub initialTimer2_Timer
  y = y + 1
  If y > 1 Then y = 0
  If y = 0 Then
    initial2.visible = 1
  Else
    initial2.visible = 0
  End If
End Sub

Sub initialTimer3_Timer
  y = y + 1
  If y > 1 Then y = 0
  If y = 0 Then
    initial3.visible = 1
  Else
    initial3.visible = 0
  End If
End Sub


'************************************************** Set Table LUT Value *****************************************************
Sub setLUT
  lut.enabled = 1
  Select Case lutValue
    Case 0: lutText.text = "1 to 1"
    Case 1: lutText.text = "ConSat"
    Case 2: lutText.text = "Onevox"
    Case 3: lutText.text = "OVDark"
    Case 4: lutText.text = "Mandolin"
    Case 5: lutText.text = "BassGeige1"
    Case 6: lutText.text = "OnevoxLite2"
    Case 7: lutText.text = "OnevoxLite1"
    Case 8: lutText.text = "1 to 1 Lite 2"
    Case 9: lutText.text = "1 to 1 Lite 1"
  End Select
  Table1.ColorGradeImage = "LUT" & lutValue
End Sub

Sub lut_Timer
  lutText.visible = 0
  lut.enabled = 0
  saveHighScore
End Sub

'************************************************* File Writing Section *****************************************************

'*************Load Scores
Sub loadHighScore
  Dim fileObj
  Dim scoreFile
  Dim temp(40)
  Dim textStr

  dim hiInitTemp(3)
  dim hiInit(5)

    Set fileObj = CreateObject("Scripting.FileSystemObject")
  If Not fileObj.FolderExists(UserDirectory) Then
    Exit Sub
  End If
  If Not fileObj.FileExists(UserDirectory & cOptions) Then
    Exit Sub
  End If
  Set scoreFile = fileObj.GetFile(UserDirectory & cOptions)
  Set textStr = scoreFile.OpenAsTextStream(1,0)
    If (textStr.AtEndOfStream = True) Then
      Exit Sub
    End If

    For x = 1 to 33
      temp(x) = textStr.readLine
    Next
    TextStr.Close
    For x = 0 to 4
      highScore(x) = cdbl (temp(x+1))
    Next
    For x = 0 to 4
      hiInit(x) = (temp(x + 6))
    Next
    i = 10
    For x = 0 to 4
      For y = 1 to 3
        i = i + 1
        initial(x,y) = cdbl (temp(i))
      Next
    Next
    credit = CInt(Right(temp(26),1))
    freePlay = CInt(Right(temp(27),1))
    balls = CInt(Right(temp(28),1))
    matchNumber = CInt(Right(temp(29),1))
    chime = CInt(Right(temp(30),1))
    pfOption = CInt(Right(temp(31),1))
    lutValue = CInt(Right(temp(32),1))
    replayEB = CInt(Right(temp(33),1))
    Set scoreFile = Nothing
      Set fileObj = Nothing
End Sub

'************Save Scores
Sub saveHighScore
Dim hiInit(5)
Dim hiInitTemp(5)
Dim FolderPath
  For x = 0 to 4
    For y = 1 to 3
      hiInitTemp(y) = chr(initial(x,y) + 64)
    Next
    hiInit(x) = hiInitTemp(1) + hiInitTemp(2) + hiInitTemp(3)
  Next
  Dim fileObj
  Dim scoreFile
  Set fileObj = createObject("Scripting.FileSystemObject")
  If Not fileObj.folderExists(userDirectory) Then
    Exit Sub
  End If
  Set scoreFile = fileObj.createTextFile(userDirectory & cOptions,True)

    For x = 0 to 4
      scoreFile.writeLine highScore(x)
    Next
    For x = 0 to 4
      scoreFile.writeLine hiInit(x)
    Next
    For x = 0 to 4
      For y = 1 to 3
        scoreFile.writeLine initial(x,y)
      Next
    Next
    scoreFile.WriteLine "Credits: " & credit
    scorefile.writeline "FreePlay (0 = coin, 1 = free): " & freePlay
    scoreFile.WriteLine "Balls: " & balls
    scoreFile.WriteLine "Match Number: " & matchNumber
    scoreFile.WriteLine "Chime (0 = sound file, 1 = DOF chime): " & chime
    scoreFile.WriteLine "pfOption (1 = L/R Stereo, 2 = U/D Stereo, 3 = quad): " & pfOption
    scoreFile.WriteLine "LUT: " & lutValue
    scoreFile.WriteLine "ReplayEB (0 = Add-A-Ball, 1 = Replay, 2 = Novelty): " & replayEB
    scoreFile.Close
  Set scoreFile = Nothing
  Set fileObj = Nothing

'This section of code writes a file in the User Folder of VisualPinball that contains the High Score data for PinballY.
'PinballY can read this data and display the high scores on the DMD during game selection mode in PinballY.

  Set FileObj = CreateObject("Scripting.FileSystemObject")

  If cPinballY <> 1 Then Exit Sub

  If Not FileObj.FolderExists(UserDirectory) Then
    Exit Sub
  End If

  FolderPath = FileObj.GetParentFolderName(UserDirectory)

  If cPinballY = 1 Then
    Set ScoreFile = FileObj.CreateTextFile(UserDirectory & hsFileName & ".PinballYHighScores",True)
  End If

  For x = 0 to 4
    ScoreFile.WriteLine HighScore(x)
    ScoreFile.WriteLine HiInit(x)
  Next
  ScoreFile.Close
  Set ScoreFile = Nothing
  Set FileObj = Nothing

End Sub

'************Shut Down and De-Energize on Tilt
Sub turnOff
  For i = 1 to 3
    EVAL("Bumper" & i).hasHitEvent = 0
  Next
    leftFlipper.RotateToStart
  stopSound "flipBuzzLA"
  stopSound "flipBuzzLB"
  stopSound "flipBuzzLC"
  stopSound "flipBuzzLD"
  stopSound "flipBuzzRA"
  stopSound "flipBuzzRB"
  stopSound "flipBuzzRC"
  stopSound "flipBuzzRD"
  If B2SOn Then DOF 101, DOFOff
  rightFlipper.RotateToStart
  If B2SOn Then DOF 102, DOFOff
  LeftSlingShot.collidable = 0
  RightSlingShot.collidable = 0
End Sub

Dim logo, oldLogo
Sub LogoLights_Timer
  logo = Int(Rnd*3)+1
  If logo = oldLogo Then logo = Int(Rnd*3)+1
  If B2SOn Then
    DOF 81, DOFOn
    DOF 82, DOFOn
    DOF 83, DOFOn
    Select Case Int(Rnd*3)+1
      Case 1 : DOF 81, DOFOff
      Case 2 : DOF 82, DOFOff
      Case 3 : DOF 83, DOFOff
    End Select
  End If
  oldLogo = logo
End Sub


'************VR Subs
Sub vrCredit
  vrCreditReel.objRotX = credit * 18
End Sub

Dim vrTen, vrHundred, vr1k, vr10k, reelUp
Sub vrScore(reelUp)
  vrTen = (INT (score(reelUp)/10) )Mod 10
  vrHundred = (INT(score(reelUp)/100)) Mod 10
  vr1k =(INT (score(reelUp)/1000) )Mod 10
  vr10k = (INT (score(reelUp)/10000) )Mod 10

  Select Case reelUp
    Case 1: vrScoreReel110.objRotX = vrTen * 36
        vrScoreReel1100.objRotX = vrHundred * 36
        vrScoreReel11k.objRotX = vr1k * 36
        vrScoreReel110k.objRotX = vr10k * 36
    Case 2: vrScoreReel210.objRotX = vrTen * 36
        vrScoreReel2100.objRotX = vrHundred * 36
        vrScoreReel21k.objRotX = vr1k * 36
        vrScoreReel210k.objRotX = vr10k * 36
    Case 3: vrScoreReel310.objRotX = vrTen * 36
        vrScoreReel3100.objRotX = vrHundred * 36
        vrScoreReel31k.objRotX = vr1k * 36
        vrScoreReel310k.objRotX = vr10k * 36
    Case 4: vrScoreReel410.objRotX = vrTen * 36
        vrScoreReel4100.objRotX = vrHundred * 36
        vrScoreReel41k.objRotX = vr1k * 36
        vrScoreReel410k.objRotX = vr10k * 36
  End Select
End Sub

Sub vrBip
  Select Case ballInPlay
    Case 0: vrBipBg.image = "bip0"
    Case 1: vrBipBg.image = "bip1"
    Case 2: vrBipBg.image = "bip2"
    Case 3: vrBipBg.image = "bip3"
    Case 4: vrBipBg.image = "bip4"
    Case 5: vrBipBg.image = "bip5"
    Case 6: vrBipBg.image = "bip6"
    Case 7: vrBipBg.image = "bip7"
    Case 8: vrBipBg.image = "bip8"
    Case 9: vrBipBg.image = "bip9"
    Case 10: vrBipBg.image = "bip10"
  End Select
End Sub

Sub vrMatch
  'Williams Style Match
  Select Case matchNumber
    Case 0: vrMatchBg.image = "m7"
    Case 1: vrMatchBg.image = "m1"
    Case 2: vrMatchBg.image = "m6"
    Case 3: vrMatchBg.image = "m10"
    Case 4: vrMatchBg.image = "m4"
    Case 5: vrMatchBg.image = "m9"
    Case 6: vrMatchBg.image = "m6"
    Case 7: vrMatchBg.image = "m5"
    Case 8: vrMatchBg.image = "m2"
    Case 9: vrMatchBg.image = "m8"
    Case 10: vrMatchBg.image = "m3"
  End Select
End Sub

Sub vrCanPlay
  Select Case players
    Case 0: vrcp.image = "cp0"
    Case 1: vrcp.image = "cp1"
    Case 2: vrcp.image = "cp2"
    Case 3: vrcp.image = "cp3"
    Case 4: vrcp.image = "cp4"
  End Select
End Sub

Sub vrPlayerUp
  Select Case player
    Case 1: vrup.image = "up1"
    Case 2: vrup.image = "up2"
    Case 3: vrUp.image = "up3"
    Case 4: vrUp.image = "up4"
  End Select
End Sub

'*******************************************
' VR Room / VR Cabinet
'*******************************************

DIM VRThings
If vrOption = 1 Then
  for each VRThings in vrCab: VRThings.visible = 1: Next
  for each VRThings in vrRoom: VRThings.visible = 0: Next
  for each VRThings in cabParts: VRThings.visible = 0: Next
Elseif vrOption = 2 Then
  for each VRThings in vrCab: VRThings.visible = 1: Next
  for each VRThings in vrRoom: VRThings.visible = 1: Next
  for each VRThings in cabParts: VRThings.visible = 0: Next
Else
  for each VRThings in vrCab: VRThings.visible = 0: Next
  for each VRThings in vrRoom: VRThings.visible = 0: Next
  tape.image = "tapeCab"
End if



'*****************************************************Supporting Code Written By Others*************************************

'*********************************************
'  VR Plunger Code From Flash by Bord and Roth
'*********************************************

Sub TimerVRPlunger_Timer
  If PinCab_Shooter.Y < -344 then
       PinCab_Shooter.Y = PinCab_Shooter.Y + 5
  End If
End Sub

Sub TimerVRPlunger1_Timer
  PinCab_Shooter.Y = -444 + (5* Plunger.Position) -20
End Sub

'************************************************************************
'                         Ball Control - 3 Axis
'************************************************************************

Dim Cup, Cdown, Cleft, Cright, Zup, contball, contballinplay, ControlBall, bcboost
Dim bcvel, bcyveloffset, bcboostmulti

bcboost = 1 'Do Not Change - default setting
bcvel = 4 'Controls the speed of the ball movement
bcyveloffset = -.014 'Offsets the force of gravity to keep the ball from drifting vertically on the table, should be negative
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

' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Function AudioPan(TableObj) 'Calculates the pan for a TableObj based on the X position on the table. "table1" is the name of the table.  New AudioPan algorithm for accurate stereo pan positioning.
    Dim tmp
    If PFOption=1 Then tmp = TableObj.x * 2 / table1.width-1
  If PFOption=2 Then tmp = TableObj.y * 2 / table1.height-1
  If tmp < 0 Then
    AudioPan = -((0.8745898957*(ABS(tmp)^12.78313661)) + (0.1264569796*(ABS(tmp)^1.000771219)))
  Else
    AudioPan = (0.8745898957*(ABS(tmp)^12.78313661)) + (0.1264569796*(ABS(tmp)^1.000771219))
  End If
End Function

Function xGain(TableObj)
'xGain algorithm calculates a PlaySound Volume parameter multiplier to provide a Constant Power "pan".
'PFOption=1:  xGain = 1 at PF Left, xGain = 0.32931 (-3dB for PlaySound's volume parameter) at PF Center and xGain = 1 at PF Right.  Used for Left & Right stereo PF Speakers.
'PFOption=2:  xGain = 1 at PF Top, xGain = 0.32931 (-3dB for PlaySound's volume parameter) at PF Center and xGain = 1 at PF Bottom.  Used for Top & Bottom stereo PF Speakers.
  Dim tmp, PI
    If PFOption=1 Then tmp = TableObj.x * 2 / table1.width-1
  If PFOption=2 Then tmp = TableObj.y * 2 / table1.height-1
  PI = 4 * ATN(1)
  If tmp < 0 Then
  xGain = 0.3293074856*EXP(-0.9652695455*tmp^3 - 2.452909811*tmp^2 - 2.597701999*tmp)
  Else
  xGain = 0.3293074856*EXP(-0.9652695455*-tmp^3 - 2.452909811*-tmp^2 - 2.597701999*-tmp)
  End If
' TB1.text = "xGain=" & Round(xGain,4)
End Function

Function XVol(tableobj)
'XVol algorithm calculates a PlaySound Volume parameter multiplier for a tableobj based on its X table position to provide a Constant Power "pan".
'XVol = 1 at PF Left, XVol = 0.32931 (-3dB for PlaySound's volume parameter) at PF Center and XVol = 0 at PF Right
Dim tmpx
  If PFOption = 3 Then
    tmpx = tableobj.x * 2 / table1.width-1
    XVol = 0.3293074856*EXP(-0.9652695455*tmpx^3 - 2.452909811*tmpx^2 - 2.597701999*tmpx)
  End If
' TB1.text = "xVol=" & Round(xVol,4)
End Function

Function YVol(tableobj)
'YVol algorithm calculates a PlaySound Volume parameter multiplier for a tableobj based on its Y table position to provide a Constant Power "fade".
'YVol = 1 at PF Top, YVol = 0.32931 (-3dB for PlaySound's volume parameter) at PF Center and YVol = 0 at PF Bottom
Dim tmpy
  If PFOption = 3 Then
    tmpy = tableobj.y * 2 / table1.height-1
    YVol = 0.3293074856*EXP(-0.9652695455*tmpy^3 - 2.452909811*tmpy^2 - 2.597701999*tmpy)
  End If
' TB2.text = "yVol=" & Round(yVol,4)
End Function

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speedb
  Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

'******************************************************
'      JP's VP10 Rolling Sounds - Modified by Whirlwind
'******************************************************

'******************************************
' Explanation of the rolling sound routine
'******************************************

' ball rolling sounds are played based on the ball speed and position
' the routine checks first for deleted balls and stops the rolling sound.
' The For loop goes through all the balls on the table and checks for the ball speed and
' if the ball is on the table (height lower than 30) then then it plays the sound
' otherwise the sound is stopped.

'New algorithms added to make sounds for TopArch Hits, Arch Rolls, ball bounces and glass hits.
'For stereo, xGain is a Playsound volume multiplier that provides a Constant Power pan.
'For quad, multiple PlaySound commands are launched together that are panned and faded to their maximum extents where PlaySound's PAN and FADE have the least error.
'XVol and YVol are Playsound volume multipliers that provide a Constant Power "pan" and "fade".
'Subtracting XVol or YVol from 1 yields an inverse response.

Const tnob = 1 ' total number of balls

ReDim rolling(tnob)
InitRolling

ReDim ArchRolling(tnob)
InitArchRolling

Dim ArchHit
Sub TopArch_Hit
  ArchHit = 1
  ArchTimer.Enabled = True
End Sub

Dim archCount
'The ArchHit sound is played and.the ArchTimer is enabled upon the first hit of the arch by the ball
'ArchTimer is enabled for 15ms after which ArchTimer2 is enabled for 2 seconds
'While ArchTimer2 is enabled it "locks-out" the ArchHit sound from being played again until after 2 seconds have elapsed
Sub ArchTimer_Timer
  archCount = archCount + 1
  If archCount = 1 Then
    archCount = 0
    ArchTimer.enabled = False
    If ArchTimer2.enabled = False Then ArchTimer2.enabled = True
  End If
End Sub

Dim archCount2
Sub ArchTimer2_Timer
  archCount2 = archCount2 + 1
  If archCount2 = 1 Then
    archCount2 = 0
    ArchTimer2.enabled = False
  End If
End Sub

Sub NotOnArch_Hit
  ArchHit = 0
End Sub

Sub NotOnArch2_Hit
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
  Dim BOT, b, paSub
  BOT = GetBalls
  paSub=35000 'Playsound pitch adder for subway/ramp rolling ball sound

'TB.text="BOT(b).Z  " & formatnumber(BOT(b).Z,1)
'TB.text="BOT(b).VelZ  " & formatnumber(BOT(b).VelZ,1)
'TB.text="GLASS  " & formatnumber((BOT(b).Y * ((GHT-GHB)/PFL)) + (GHB*50/1.0625) - (BallSize/2),1)
'TB.text = "ArchTimer.enabled=" & ArchTimer.enabled & "    ArchTimer2.enabled=" & ArchTimer2.enabled & "    ArchHit=" & ArchHit

  ' stop the sound of deleted balls
  For b = UBound(BOT) + 1 to tnob
    rolling(b) = False
    StopSound("BallrollingA" & b)
    StopSound("BallrollingB" & b)
    StopSound("BallrollingC" & b)
    StopSound("BallrollingD" & b)
  Next

  ' exit the sub if no balls on the table
  If UBound(BOT) = -1 Then Exit Sub

  ' play the rolling sound for each ball
  For b = 0 to UBound(BOT)
' TB.text = "Ball0.z=" & (BOT(0).z) 'Use to verify ball.z (currently -3) when in Adrian's new saucer primitives.


' Ball Rolling sounds
'**********************
' Ball=50 units=1.0625".  One unit = 0.02125"  Ball.z is ball center.
' A ball in Adrian's saucer has a Z of -3.  Use <-5 for subway sounds.
  If PFOption = 1 or PFOption = 2 Then
    If BallVel(BOT(b)) > 1 AND BOT(b).z > 10 and BOT(b).z <26 Then  'Ball on playfield
      rolling(b) = True
      PlaySound("BallrollingA" & b), -1, Vol(BOT(b)) * 0.2 * xGain(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, 0 'Left & Right stereo or Top & Bottom stereo PF Speakers.
    ElseIf BallVel(BOT(b)) > 1 AND BOT(b).z < -5 Then 'Ball on subway
      PlaySound("BallrollingA" & b), -1, Vol(BOT(b)) * 0.2 * xGain(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b))+paSub, 1, 0, 0 'Left & Right stereo or Top & Bottom stereo PF Speakers.
    ElseIf rolling(b) = True Then
      StopSound("BallrollingA" & b)
      rolling(b) = False
    End If
  End If

  If PFOption = 3 Then
    If BallVel(BOT(b)) > 1 AND BOT(b).z > 10 and BOT(b).z < 26 Then 'Ball on playfield
      rolling(b) = True
      PlaySound("BallrollingA" & b), -1, Vol(BOT(b)) * 0.2 *    XVol(BOT(b))  *     YVol(BOT(b)), -1, 0, Pitch(BOT(b)), 1, 0, -1  'Top Left PF Speaker
      PlaySound("BallrollingB" & b), -1, Vol(BOT(b)) * 0.2 * (1-XVol(BOT(b))) *     YVol(BOT(b)),  1, 0, Pitch(BOT(b)), 1, 0, -1  'Top Right PF Speaker
      PlaySound("BallrollingC" & b), -1, Vol(BOT(b)) * 0.2 *    XVol(BOT(b))  * (1-YVol(BOT(b))), -1, 0, Pitch(BOT(b)), 1, 0,  1  'Bottom Left PF Speaker
      PlaySound("BallrollingD" & b), -1, Vol(BOT(b)) * 0.2 * (1-XVol(BOT(b))) * (1-YVol(BOT(b))),  1, 0, Pitch(BOT(b)), 1, 0,  1  'Bottom Right PF Speaker
    ElseIf BallVel(BOT(b)) > 1 AND BOT(b).z < -5 Then 'Ball on subway
      PlaySound("BallrollingA" & b), -1, Vol(BOT(b)) * 0.2 *    XVol(BOT(b))  *     YVol(BOT(b)), -1, 0, Pitch(BOT(b))+paSub, 1, 0, -1  'Top Left PF Speaker
      PlaySound("BallrollingB" & b), -1, Vol(BOT(b)) * 0.2 * (1-XVol(BOT(b))) *     YVol(BOT(b)),  1, 0, Pitch(BOT(b))+paSub, 1, 0, -1  'Top Right PF Speaker
      PlaySound("BallrollingC" & b), -1, Vol(BOT(b)) * 0.2 *    XVol(BOT(b))  * (1-YVol(BOT(b))), -1, 0, Pitch(BOT(b))+paSub, 1, 0,  1  'Bottom Left PF Speaker
      PlaySound("BallrollingD" & b), -1, Vol(BOT(b)) * 0.2 * (1-XVol(BOT(b))) * (1-YVol(BOT(b))),  1, 0, Pitch(BOT(b))+paSub, 1, 0,  1  'Bottom Right PF Speaker
    ElseIf rolling(b) = True Then
      StopSound("BallrollingA" & b)   'Top Left PF Speaker
      StopSound("BallrollingB" & b)   'Top Right PF Speaker
      StopSound("BallrollingC" & b)   'Bottom Left PF Speaker
      StopSound("BallrollingD" & b)   'Bottom Right PF Speaker
      rolling(b) = False
    End If
  End If

' Arch Hit and Arch Rolling sounds
'***********************************
  If PFOption = 1 or PFOption = 2 Then
    If BallVel(BOT(b)) > 1 And ArchHit =1 Then
      If ArchTimer2.enabled = 0 Then
        PlaySound("ArchHit" & b), 0, (BallVel(BOT(b))/32)^5 * xGain(BOT(b)), AudioPan(BOT(b)), 0, (BallVel(BOT(b))/40)^7, 0, 0, 0 'Left & Right stereo or Top & Bottom stereo PF Speakers.
      End If
      ArchRolling(b) = True
      PlaySound("ArchRollA" & b), -1, (BallVel(BOT(b))/40)^5 * xGain(BOT(b)), AudioPan(BOT(b)), 0, (BallVel(BOT(b))/40)^7, 1, 0, 0  'Left & Right stereo or Top & Bottom stereo PF Speakers.
    Else
      If ArchRolling(b) = True Then
      StopSound("ArchRollA" & b)
      ArchRolling(b) = False
      End If
    End If
' If ArchTimer2.enabled = 0 And ArchHit = 1 Then TB.text = "ArchHit vol=" & round((BallVel(BOT(b))/32)^5 * 1 * xGain(BOT(b)),4) 'Keep below 1.
  End If

  If PFOption = 3 Then
    If BallVel(BOT(b)) > 1 And ArchHit =1 Then
      If ArchTimer2.enabled = 0 Then
        PlaySound("ArchHit" & b),   0, (BallVel(BOT(b))/32)^5 *    XVol(BOT(b))  *     YVol(BOT(b)),  -1, 0, (BallVel(BOT(b))/40)^5, 0, 0, -1 'Top Left PF Speaker
        PlaySound("ArchHit" & b),   0, (BallVel(BOT(b))/32)^5 * (1-XVol(BOT(b))) *     YVol(BOT(b)),   1, 0, (BallVel(BOT(b))/40)^5, 0, 0, -1 'Top Right PF Speaker
        PlaySound("ArchHit" & b),   0, (BallVel(BOT(b))/32)^5 *    XVol(BOT(b))  *  (1-YVol(BOT(b))), -1, 0, (BallVel(BOT(b))/40)^5, 0, 0,  1 'Bottom Left PF Speaker
        PlaySound("ArchHit" & b),   0, (BallVel(BOT(b))/32)^5 * (1-XVol(BOT(b))) *  (1-YVol(BOT(b))),  1, 0, (BallVel(BOT(b))/40)^5, 0, 0,  1 'Bottom Right PF Speaker
      End If
      ArchRolling(b) = True
      PlaySound("ArchRollA" & b), -1, (BallVel(BOT(b))/40)^5 *    XVol(BOT(b))  *     YVol(BOT(b)),  -1, 0, (BallVel(BOT(b))/40)^5, 1, 0, -1  'Top Left PF Speaker
      PlaySound("ArchRollB" & b), -1, (BallVel(BOT(b))/40)^5 * (1-XVol(BOT(b))) *     YVol(BOT(b)),   1, 0, (BallVel(BOT(b))/40)^5, 1, 0, -1  'Top Right PF Speaker
      PlaySound("ArchRollC" & b), -1, (BallVel(BOT(b))/40)^5 *    XVol(BOT(b))  *  (1-YVol(BOT(b))), -1, 0, (BallVel(BOT(b))/40)^5, 1, 0,  1  'Bottom Left PF Speaker
      PlaySound("ArchRollD" & b), -1, (BallVel(BOT(b))/40)^5 * (1-XVol(BOT(b))) *  (1-YVol(BOT(b))),  1, 0, (BallVel(BOT(b))/40)^5, 1, 0,  1  'Bottom Right PF Speaker
    Else
      If ArchRolling(b) = True Then
      StopSound("ArchRollA" & b)  'Top Left PF Speaker
      StopSound("ArchRollB" & b)  'Top Right PF Speaker
      StopSound("ArchRollC" & b)  'Bottom Left PF Speaker
      StopSound("ArchRollD" & b)  'Bottom Right PF Speaker
      ArchRolling(b) = False
      End If
    End If
  End If

' Ball drop sounds
'*******************
'Four intensities of ball bounce sound files ranging from 1 to 4 bounces.  The number of bounces increases as the ball's downward Z velocity increases.
'A BOT(b).VelZ < -2 eliminates nuisance ball bounce sounds.

  If PFOption = 1 or PFOption = 2 Then
    If BOT(b).VelZ > -4 And BOT(b).VelZ < -2 And BOT(b).Z > 24 And BallinPlay => 1 Then
      PlaySound "BallDrop1" & b, 0, ABS(BOT(b).VelZ)/600 * xGain(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 0, 0, 0 'Left & Right stereo or Top & Bottom stereo PF Speakers.
    ElseIf BOT(b).VelZ > -8 And BOT(b).VelZ < -4 And BOT(b).Z > 24 And BallinPlay => 1 Then
      PlaySound "BallDrop2" & b, 0, ABS(BOT(b).VelZ)/600 * xGain(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 0, 0, 0 'Left & Right stereo or Top & Bottom stereo PF Speakers.
    ElseIf BOT(b).VelZ > -12 And BOT(b).VelZ < -8 And BOT(b).Z > 24 And BallinPlay => 1 Then
      PlaySound "BallDrop3" & b, 0, ABS(BOT(b).VelZ)/600 * xGain(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 0, 0, 0 'Left & Right stereo or Top & Bottom stereo PF Speakers.
    ElseIf BOT(b).VelZ < -12 And BOT(b).Z > 24 And BallinPlay => 1 Then
      PlaySound "BallDrop4" & b, 0, ABS(BOT(b).VelZ)/600 * xGain(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 0, 0, 0 'Left & Right stereo or Top & Bottom stereo PF Speakers.
    End If
  End If

  If PFOption = 3 Then
    If BOT(b).VelZ > -4 And BOT(b).VelZ < -2 And BOT(b).Z > 24 And BallinPlay => 1 Then
      PlaySound "BallDrop1" & b, 0, ABS(BOT(b).VelZ)/600 *    XVol(BOT(b))  *     YVol(BOT(b)), -1, 0, Pitch(BOT(b)), 1, 0, -1  'Top Left PF Speaker
      PlaySound "BallDrop1" & b, 0, ABS(BOT(b).VelZ)/600 * (1-XVol(BOT(b))) *     YVol(BOT(b)),  1, 0, Pitch(BOT(b)), 1, 0, -1  'Top Right PF Speaker
      PlaySound "BallDrop1" & b, 0, ABS(BOT(b).VelZ)/600 *    XVol(BOT(b))  * (1-YVol(BOT(b))), -1, 0, Pitch(BOT(b)), 1, 0,  1  'Bottom Left PF Speaker
      PlaySound "BallDrop1" & b, 0, ABS(BOT(b).VelZ)/600 * (1-XVol(BOT(b))) * (1-YVol(BOT(b))),  1, 0, Pitch(BOT(b)), 1, 0,  1  'Bottom Right PF Speaker
    ElseIf BOT(b).VelZ > -8 And BOT(b).VelZ < -4 And BOT(b).Z > 24 And BallinPlay => 1 Then
      PlaySound "BallDrop2" & b, 0, ABS(BOT(b).VelZ)/600 *    XVol(BOT(b))  *     YVol(BOT(b)), -1, 0, Pitch(BOT(b)), 1, 0, -1  'Top Left PF Speaker
      PlaySound "BallDrop2" & b, 0, ABS(BOT(b).VelZ)/600 * (1-XVol(BOT(b))) *     YVol(BOT(b)),  1, 0, Pitch(BOT(b)), 1, 0, -1  'Top Right PF Speaker
      PlaySound "BallDrop2" & b, 0, ABS(BOT(b).VelZ)/600 *    XVol(BOT(b))  * (1-YVol(BOT(b))), -1, 0, Pitch(BOT(b)), 1, 0,  1  'Bottom Left PF Speaker
      PlaySound "BallDrop2" & b, 0, ABS(BOT(b).VelZ)/600 * (1-XVol(BOT(b))) * (1-YVol(BOT(b))),  1, 0, Pitch(BOT(b)), 1, 0,  1  'Bottom Right PF Speaker
    ElseIf BOT(b).VelZ > -12 And BOT(b).VelZ < -8 And BOT(b).Z > 24 And BallinPlay => 1 Then
      PlaySound "BallDrop3" & b, 0, ABS(BOT(b).VelZ)/600 *    XVol(BOT(b))  *     YVol(BOT(b)), -1, 0, Pitch(BOT(b)), 1, 0, -1  'Top Left PF Speaker
      PlaySound "BallDrop3" & b, 0, ABS(BOT(b).VelZ)/600 * (1-XVol(BOT(b))) *     YVol(BOT(b)),  1, 0, Pitch(BOT(b)), 1, 0, -1  'Top Right PF Speaker
      PlaySound "BallDrop3" & b, 0, ABS(BOT(b).VelZ)/600 *    XVol(BOT(b))  * (1-YVol(BOT(b))), -1, 0, Pitch(BOT(b)), 1, 0,  1  'Bottom Left PF Speaker
      PlaySound "BallDrop3" & b, 0, ABS(BOT(b).VelZ)/600 * (1-XVol(BOT(b))) * (1-YVol(BOT(b))),  1, 0, Pitch(BOT(b)), 1, 0,  1  'Bottom Right PF Speaker
    ElseIf BOT(b).VelZ < -12 And BOT(b).Z > 24 And BallinPlay => 1 Then
      PlaySound "BallDrop4" & b, 0, ABS(BOT(b).VelZ)/600 *    XVol(BOT(b))  *     YVol(BOT(b)), -1, 0, Pitch(BOT(b)), 1, 0, -1  'Top Left PF Speaker
      PlaySound "BallDrop4" & b, 0, ABS(BOT(b).VelZ)/600 * (1-XVol(BOT(b))) *     YVol(BOT(b)),  1, 0, Pitch(BOT(b)), 1, 0, -1  'Top Right PF Speaker
      PlaySound "BallDrop4" & b, 0, ABS(BOT(b).VelZ)/600 *    XVol(BOT(b))  * (1-YVol(BOT(b))), -1, 0, Pitch(BOT(b)), 1, 0,  1  'Bottom Left PF Speaker
      PlaySound "BallDrop4" & b, 0, ABS(BOT(b).VelZ)/600 * (1-XVol(BOT(b))) * (1-YVol(BOT(b))),  1, 0, Pitch(BOT(b)), 1, 0,  1  'Bottom Right PF Speaker
    End If
  End If
'TB.text="BOT(b).VelZ  " & round(BOT(b).VelZ,1)


' Glass hit sounds
'*******************
' Ball=50 units=1.0625".  Ball.z is ball center.  Balls are physically limited by the table's Top Glass Height property which is always parallel to the playfield
' To ensure a ball can go high enough to trigger a glass hit the max ball.z is 25 units below Top Glass Height-5
' Change GHB below if the glass is not parallel to the playfield

  Dim GHT, GHB, PFL
  GHT = (table1.glassheight-5)*1.0625/50  'Glass height at top of real playfield in inches
  GHB = (table1.glassheight-5)*1.0625/50  'Glass height at bottom of real playfield in inches
  PFL = table1.height*1.0625/50   'Length of real playfield in inches
' TB.text = "GHT=" & GHT & """" & "  GHB=" & GHT & """" & "  PFL=" & formatnumber(PFL,5) & """"

  If PFOption = 1 or PFOption = 2 Then
    If BOT(b).Z > (BOT(b).Y * ((GHT-GHB)/PFL)) + (GHB*50/1.0625) - BallSize/2 And BallinPlay => 1 Then
      PlaySound "GlassHit" & b, 0, ABS(BOT(b).VelZ)/30 * xGain(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 0, 0, 0 'Left & Right stereo or Top & Bottom stereo PF Speakers.
    End If
  End If

  If PFOption = 3 Then
    If BOT(b).Z > (BOT(b).Y * ((GHT-GHB)/PFL)) + (GHB*50/1.0625) - Ballsize/2 And BallinPlay => 1 Then
      PlaySound "GlassHit" & b, 0, ABS(BOT(b).VelZ)/30 *    XVol(BOT(b))  *     YVol(BOT(b)), -1, 0, Pitch(BOT(b)), 0, 0, -1  'Top Left PF Speaker
      PlaySound "GlassHit" & b, 0, ABS(BOT(b).VelZ)/30 * (1-XVol(BOT(b))) *     YVol(BOT(b)),  1, 0, Pitch(BOT(b)), 0, 0, -1  'Top Right PF Speaker
      PlaySound "GlassHit" & b, 0, ABS(BOT(b).VelZ)/30 *    XVol(BOT(b))  * (1-YVol(BOT(b))), -1, 0, Pitch(BOT(b)), 0, 0,  1  'Bottom Left PF Speaker
      PlaySound "GlassHit" & b, 0, ABS(BOT(b).VelZ)/30 * (1-XVol(BOT(b))) * (1-YVol(BOT(b))),  1, 0, Pitch(BOT(b)), 0, 0,  1  'Bottom Right PF Speaker
    End If
  End If
  Next
End Sub

'*************Hit Sound Routines
'Eliminated the Hit Subs extra velocity criteria since the PlayFieldSoundAB command already incorporates the ball's velocity.

Sub aMetalPins_Hit(idx)
  PlayFieldSoundAB "metalPinHit", 0, 1
End Sub

Sub aTargets_Hit(idx)
  PlayFieldSoundAB "target", 0, 1
End Sub

Sub aMetalsHigh_Hit(idx)
  PlayFieldSoundAB "metalHitHigh", 0, 1
End Sub

Sub aMetalsMedium_Hit(idx)
  PlayFieldSoundAB "metalHitMedium", 0, 1
End Sub

Sub aMetalsLow_Hit(idx)
  PlayFieldSoundAB "metalHitLow", 0, 1
End Sub

Sub aGates_Hit(idx)
  PlayFieldSoundAB "gate", 0, 1
End Sub

Sub aRubberBands_Hit(idx)
  If BallinPlay > 0 Then  'Eliminates the thump of Trough Ball Creation balls hitting walls 9 and 14 during table initiation
  PlayFieldSoundAB "rubberBand", 0, 1
  End If
End Sub

Sub RubberWheel_hit
  PlayFieldSoundAB "rubberWheel", 0, 1
End Sub

Sub aRubberPosts_Hit(idx)
  PlayFieldSoundAB "rubberPost", 0, 1
End Sub

Sub aWoods_Hit(idx)
  PlayFieldSoundAB "wood", 0, 1
End Sub

Sub LeftFlipper_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RightFlipper_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
  Select Case Int(Rnd*3)+1
    Case 1 : PlayFieldSoundAB "flip_hit_1", 0, 1
    Case 2 : PlayFieldSoundAB "flip_hit_2", 0, 1
    Case 3 : PlayFieldSoundAB "flip_hit_3", 0, 1
  End Select
End Sub

Sub ApronWall_Hit
  Dim Volume
  If ActiveBall.vely < 0 Then Volume = abs(ActiveBall.vely) / 1 Else Volume = ActiveBall.vely / 30  'The first bounce is -vely subsequent bounces are +vely
    If ActiveBall.z > 24 Then
      If PFOption = 1 Or PFOption = 2 Then
        PlaySound "ApronHit", 0, Volume * xGain(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, 0  'Left & Right stereo or Top & Bottom stereo PF Speakers.
      End If
      If PFOption = 3 Then
        PlaySound "ApronHit", 0, Volume *    XVol(ActiveBall)  *    YVol(ActiveBall),  -1, 0, Pitch(ActiveBall), 0, 0, -1 'Top Left PF Speaker
        PlaySound "ApronHit", 0, Volume * (1-XVol(ActiveBall)) *    YVol(ActiveBall),   1, 0, Pitch(ActiveBall), 0, 0, -1 'Top Right PF Speaker
        PlaySound "ApronHit", 0, Volume *    XVol(ActiveBall)  * (1-YVol(ActiveBall)), -1, 0, Pitch(ActiveBall), 0, 0,  1 'Bottom Left PF Speaker
        PlaySound "ApronHit", 0, Volume * (1-XVol(ActiveBall)) * (1-YVol(ActiveBall)),  1, 0, Pitch(ActiveBall), 0, 0,  1 'Bottom Right PF Speaker
      End If
    End If
End Sub

'**********************
' Ball Collision Sound
'**********************

'**************************************
' Explanation of the collision routine
'**************************************

' The collision is built in VP.
' You only need to add a Sub OnBallBallCollision(ball1, ball2, velocity) and when two balls collide they
' will call this routine.

'New algorithm for BallBallCollision
'For stereo, xGain is a Playsound volume multiplier that provides a Constant Power pan.
'For quad, multiple PlaySound commands are launched together that are panned and faded to their maximum extents where PlaySound's PAN and FADE have the least error.
'XVol and YVol are Playsound volume multipliers that provide a Constant Power "pan" and "fade".
'Subtracting XVol or YVol from 1 yields an inverse response.

Sub OnBallBallCollision(ball1, ball2, velocity)
  If PFOption = 1 or PFOption = 2 Then
    PlaySound "BBcollide", 0, (Csng(velocity) ^2 / 2000) * xGain(ball1), AudioPan(ball1), 0, Pitch(ball1), 0, 0, 0  'Left & Right stereo or Top & Bottom stereo PF Speakers.
  End If
  If PFOption = 3 Then
    PlaySound "BBcollide", 0, (Csng(velocity) ^2 / 2000) *    XVol(ball1)  *    YVol(ball1),  -1, 0, Pitch(ball1), 0, 0, -1 'Top Left Playfield Speaker
    PlaySound "BBcollide", 0, (Csng(velocity) ^2 / 2000) * (1-XVol(ball1)) *    YVol(ball1),   1, 0, Pitch(ball1), 0, 0, -1 'Top Right Playfield Speaker
    PlaySound "BBcollide", 0, (Csng(velocity) ^2 / 2000) *    XVol(ball1)  * (1-YVol(ball1)), -1, 0, Pitch(ball1), 0, 0,  1 'Bottom Left Playfield Speaker
    PlaySound "BBcollide", 0, (Csng(velocity) ^2 / 2000) * (1-XVol(ball1)) * (1-YVol(ball1)),  1, 0, Pitch(ball1), 0, 0,  1 'Bottom Right Playfield Speaker
  End If
End Sub

Sub PlayFieldSound (SoundName, Looper, TableObject, VolMult)
'Plays the sound of a table object at the table object's coordinates.
'For stereo, xGain is a Playsound volume multiplier that provides a Constant Power pan.
'For quad, multiple PlaySound commands are launched together that are panned and faded to their maximum extents where PlaySound's PAN and FADE have the least error.
'XVol and YVol are Playsound volume multipliers that provide a Constant Power "pan" and "fade".
'Subtracting XVol or YVol from 1 yields an inverse response.

  If PFOption = 1 Or PFOption = 2 Then
    If Looper = -1 Then
      PlaySound SoundName&"A", Looper, VolMult * xGain(TableObject), AudioPan(TableObject), 0, 0, 0, 0, 0 'Left & Right stereo or Top & Bottom stereo PF Speakers.
    End If
    If Looper = 0 Then
      PlaySound SoundName, Looper, VolMult * xGain(TableObject), AudioPan(TableObject), 0, 0, 0, 0, 0 'Left & Right stereo or Top & Bottom stereo PF Speakers.
    End If
  End If
  If PFOption = 3 Then
    If Looper = -1 Then
      PlaySound SoundName&"A", Looper, VolMult *    XVol(TableObject)  *    YVol(TableObject),  -1, 0, 0, 0, 0, -1  'Top Left PF Speaker
      PlaySound SoundName&"B", Looper, VolMult * (1-XVol(TableObject)) *    YVol(TableObject),   1, 0, 0, 0, 0, -1  'Top Right PF Speaker
      PlaySound SoundName&"C", Looper, VolMult *    XVol(TableObject)  * (1-YVol(TableObject)), -1, 0, 0, 0, 0,  1  'Bottom Left PF Speaker
      PlaySound SoundName&"D", Looper, VolMult * (1-XVol(TableObject)) * (1-YVol(TableObject)),  1, 0, 0, 0, 0,  1  'Bottom Right PF Speaker
    End If
    If Looper = 0 Then
      PlaySound SoundName, Looper, VolMult *    XVol(TableObject)  *    YVol(TableObject),  -1, 0, 0, 0, 0, -1  'Top Left PF Speaker
      PlaySound SoundName, Looper, VolMult * (1-XVol(TableObject)) *    YVol(TableObject),   1, 0, 0, 0, 0, -1  'Top Right PF Speaker
      PlaySound SoundName, Looper, VolMult *    XVol(TableObject)  * (1-YVol(TableObject)), -1, 0, 0, 0, 0,  1  'Bottom Left PF Speaker
      PlaySound SoundName, Looper, VolMult * (1-XVol(TableObject)) * (1-YVol(TableObject)),  1, 0, 0, 0, 0,  1  'Bottom Right PF Speaker
    End If
  End If
End Sub

Sub PlayFieldSoundAB (SoundName, Looper, VolMult)
'Plays the sound of a table object at the Active Ball's location.
'For stereo, xGain is a Playsound volume multiplier that provides a Constant Power pan.
'For quad, multiple PlaySound commands are launched together that are panned and faded to their maximum extents where PlaySound's PAN and FADE have the least error.
'XVol and YVol are Playsound volume multipliers that provide a Constant Power "pan" and "fade".
'Subtracting XVol or YVol from 1 yields an inverse response.

  If PFOption = 1 Or PFOption = 2 Then
    PlaySound SoundName, Looper, VolMult * Vol(ActiveBall) * xGain(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, 0 'Left & Right stereo or Top & Bottom stereo PF Speakers.
  End If
  If PFOption = 3 Then
    PlaySound SoundName, Looper, VolMult * Vol(ActiveBall) *    XVol(ActiveBall)  *    YVol(ActiveBall),  -1, 0, Pitch(ActiveBall), 0, 0, -1  'Top Left PF Speaker
    PlaySound SoundName, Looper, VolMult * Vol(ActiveBall) * (1-XVol(ActiveBall)) *    YVol(ActiveBall),   1, 0, Pitch(ActiveBall), 0, 0, -1  'Top Right PF Speaker
    PlaySound SoundName, Looper, VolMult * Vol(ActiveBall) *    XVol(ActiveBall)  * (1-YVol(ActiveBall)), -1, 0, Pitch(ActiveBall), 0, 0,  1  'Bottom Left PF Speaker
    PlaySound SoundName, Looper, VolMult * Vol(ActiveBall) * (1-XVol(ActiveBall)) * (1-YVol(ActiveBall)),  1, 0, Pitch(ActiveBall), 0, 0,  1  'Bottom Right PF Speaker
  End If
End Sub

Sub PlayReelSound (SoundName, Pan)
Dim ReelVolAdj
ReelVolAdj = 0.2
'Provides a Constant Power Pan for the backglass reel sound volume to match the playfield's Constant Power Pan response
  If showDT = False Then  '-3dB for desktop mode
    If Pan = "Lpan" Then PlaySound SoundName, 0, ReelVolAdj * 1.00, -0.12, 0, 0, 0, 0, 0  'Panned 3/4 Left at 0dB * ReelVolAdj
    If Pan = "Mpan" Then PlaySound SoundName, 0, ReelVolAdj * 0.33,  0.00, 0, 0, 0, 0, 0  'Panned Middle at -3dB * ReelVolAdj
    If Pan = "Rpan" Then PlaySound SoundName, 0, ReelVolAdj * 1.00,  0.12, 0, 0, 0, 0, 0  'Panned 3/4 Right at 0dB * ReelVolAdj
  Else
    If Pan = "Lpan" Then PlaySound SoundName, 0, ReelVolAdj * 0.33, -0.12, 0, 0, 0, 0, 0  'Panned 3/4 Left at -3dB * ReelVolAdj
    If Pan = "Mpan" Then PlaySound SoundName, 0, ReelVolAdj * 0.11,  0.00, 0, 0, 0, 0, 0  'Panned Middle at -6dB * ReelVolAdj
    If Pan = "Rpan" Then PlaySound SoundName, 0, ReelVolAdj * 0.33,  0.12, 0, 0, 0, 0, 0  'Panned 3/4 Right at -3dB * ReelVolAdj
  End If
End Sub

Sub Table1_Exit
  If B2sOn Then controller.stop
  saveHighScore
End Sub


'******************************************************
'****  GNEREAL ADVICE ON PHYSICS
'******************************************************
'
' It's advised that flipper corrections, dampeners, and general physics settings should all be updated per these
' examples as all of these improvements work together to provide a realistic physics simulation.
'
' Tutorial videos provided by Bord
' Flippers:   https://www.youtube.com/watch?v=FWvM9_CdVHw
' Dampeners:  https://www.youtube.com/watch?v=tqsxx48C6Pg
' Physics:    https://www.youtube.com/watch?v=UcRMG-2svvE
'
'
' Note: BallMass must be set to 1. BallSize should be set to 50 (in other words the ball radius is 25)
'
' Recommended Table Physics Settings
' | Gravity Constant             | 0.97      |
' | Playfield Friction           | 0.15-0.25 |
' | Playfield Elasticity         | 0.25      |
' | Playfield Elasticity Falloff | 0         |
' | Playfield Scatter            | 0         |
' | Default Element Scatter      | 2         |
'
' Bumpers
' | Force         | 9.5-10.5 |
' | Hit Threshold | 1.6-2    |
' | Scatter Angle | 2        |
'
' Slingshots
' | Hit Threshold      | 2    |
' | Slingshot Force    | 4-5  |
' | Slingshot Theshold | 2-3  |
' | Elasticity         | 0.85 |
' | Friction           | 0.8  |
' | Scatter Angle      | 1    |



'******************************************************
'****  FLIPPER CORRECTIONS by nFozzy
'******************************************************
'
' There are several steps for taking advantage of nFozzy's flipper solution.  At a high level we'll need the following:
' 1. flippers with specific physics settings
' 2. custom triggers for each flipper (TriggerLF, TriggerRF)
' 3. an object or point to tell the script where the tip of the flipper is at rest (EndPointLp, EndPointRp)
' 4. and, special scripting
'
' A common mistake is incorrect flipper length.  A 3-inch flipper with rubbers will be about 3.125 inches long.
' This translates to about 147 vp units.  Therefore, the flipper start radius + the flipper length + the flipper end
' radius should  equal approximately 147 vp units. Another common mistake is is that sometimes the right flipper
' angle was set with a large postive value (like 238 or something). It should be using negative value (like -122).
'
' The following settings are a solid starting point for various eras of pinballs.
' |                    | EM's           | late 70's to mid 80's | mid 80's to early 90's | mid 90's and later |
' | ------------------ | -------------- | --------------------- | ---------------------- | ------------------ |
' | Mass               | 1              | 1                     | 1                      | 1                  |
' | Strength           | 500-1000 (750) | 1400-1600 (1500)      | 2000-2600              | 3200-3300 (3250)   |
' | Elasticity         | 0.88           | 0.88                  | 0.88                   | 0.88               |
' | Elasticity Falloff | 0.15           | 0.15                  | 0.15                   | 0.15               |
' | Fricition          | 0.8-0.9        | 0.9                   | 0.9                    | 0.9                |
' | Return Strength    | 0.11           | 0.09                  | 0.07                   | 0.055              |
' | Coil Ramp Up       | 2.5            | 2.5                   | 2.5                    | 2.5                |
' | Scatter Angle      | 0              | 0                     | 0                      | 0                  |
' | EOS Torque         | 0.3            | 0.3                   | 0.275                  | 0.275              |
' | EOS Torque Angle   | 4              | 4                     | 6                      | 6                  |
'


'******************************************************
' Flippers Polarity (Select appropriate sub based on era)
'******************************************************

dim LF : Set LF = New FlipperPolarity
dim RF : Set RF = New FlipperPolarity

InitPolarity

'
''*******************************************
'' Late 70's to early 80's
'
Sub InitPolarity()
        dim x, a : a = Array(LF, RF)
        for each x in a
                x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1        'disabled
                x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1
                x.enabled = True
                x.TimeDelay = 80  '*****Important, this variable is an offset for the speed that the ball travels down the table to determine if the flippers have been fired
              'This is needed because the corrections to ball trajectory should only applied if the flippers have been fired and the ball is in the trigger zones.
              'FlipAT is set to GameTime when the ball enters the flipper trigger zones and if GameTime is less than FlipAT + this time delay then changes to velocity
              'and trajectory are applied.  If the flipper is fired before the ball enters the trigger zone then with this delay added to FlipAT the changes
              'to tragectory and velocity will not be applied.  Also if the flipper is in the final 20 degrees changes to ball values will also not be applied.
              '"Faster" tables will need a smaller value while "slower" tables will need a larger value to give the ball more time to get to the flipper.
              'If this value is not set high enough the Flipper Velocity and Polarity corrections will NEVER be applied.

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
        LF.EndPoint = EndPointLp  'you can use just a coordinate, or an object with a .x property. Using a couple of simple primitive objects
        RF.Object = RightFlipper
        RF.EndPoint = EndPointRp
End Sub



''*******************************************
'' Mid 80's
'
'Sub InitPolarity()
'        dim x, a : a = Array(LF, RF)
'        for each x in a
'                x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1        'disabled
'                x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1
'                x.enabled = True
'                x.TimeDelay = 80
'        Next
'
'        AddPt "Polarity", 0, 0, 0
'        AddPt "Polarity", 1, 0.05, -3.7
'        AddPt "Polarity", 2, 0.33, -3.7
'        AddPt "Polarity", 3, 0.37, -3.7
'        AddPt "Polarity", 4, 0.41, -3.7
'        AddPt "Polarity", 5, 0.45, -3.7
'        AddPt "Polarity", 6, 0.576,-3.7
'        AddPt "Polarity", 7, 0.66, -2.3
'        AddPt "Polarity", 8, 0.743, -1.5
'        AddPt "Polarity", 9, 0.81, -1
'        AddPt "Polarity", 10, 0.88, 0
'
'        addpt "Velocity", 0, 0,         1
'        addpt "Velocity", 1, 0.16, 1.06
'        addpt "Velocity", 2, 0.41,         1.05
'        addpt "Velocity", 3, 0.53,         1'0.982
'        addpt "Velocity", 4, 0.702, 0.968
'        addpt "Velocity", 5, 0.95,  0.968
'        addpt "Velocity", 6, 1.03,         0.945
'
'        LF.Object = LeftFlipper
'        LF.EndPoint = EndPointLp
'        RF.Object = RightFlipper
'        RF.EndPoint = EndPointRp
'End Sub
'
'


'*******************************************
'  Late 80's early 90's
'
'Sub InitPolarity()
' dim x, a : a = Array(LF, RF)
' for each x in a
'   x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1        'disabled
'   x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1
'   x.enabled = True
'   x.TimeDelay = 60
' Next
'
' AddPt "Polarity", 0, 0, 0
' AddPt "Polarity", 1, 0.05, -5
' AddPt "Polarity", 2, 0.4, -5
' AddPt "Polarity", 3, 0.6, -4.5
' AddPt "Polarity", 4, 0.65, -4.0
' AddPt "Polarity", 5, 0.7, -3.5
' AddPt "Polarity", 6, 0.75, -3.0
' AddPt "Polarity", 7, 0.8, -2.5
' AddPt "Polarity", 8, 0.85, -2.0
' AddPt "Polarity", 9, 0.9,-1.5
' AddPt "Polarity", 10, 0.95, -1.0
' AddPt "Polarity", 11, 1, -0.5
' AddPt "Polarity", 12, 1.1, 0
' AddPt "Polarity", 13, 1.3, 0
'
' addpt "Velocity", 0, 0,         1
' addpt "Velocity", 1, 0.16, 1.06
' addpt "Velocity", 2, 0.41,         1.05
' addpt "Velocity", 3, 0.53,         1'0.982
' addpt "Velocity", 4, 0.702, 0.968
' addpt "Velocity", 5, 0.95,  0.968
' addpt "Velocity", 6, 1.03,         0.945
'
' LF.Object = LeftFlipper
' LF.EndPoint = EndPointLp
' RF.Object = RightFlipper
' RF.EndPoint = EndPointRp
'End Sub
'
'
'
'
'*******************************************
'' Early 90's and after
'
'Sub InitPolarity()
'        dim x, a : a = Array(LF, RF)
'        for each x in a
'                x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1        'disabled
'                x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1
'                x.enabled = True
'                x.TimeDelay = 60
'        Next
'
'        AddPt "Polarity", 0, 0, 0
'        AddPt "Polarity", 1, 0.05, -5.5
'        AddPt "Polarity", 2, 0.4, -5.5
'        AddPt "Polarity", 3, 0.6, -5.0
'        AddPt "Polarity", 4, 0.65, -4.5
'        AddPt "Polarity", 5, 0.7, -4.0
'        AddPt "Polarity", 6, 0.75, -3.5
'        AddPt "Polarity", 7, 0.8, -3.0
'        AddPt "Polarity", 8, 0.85, -2.5
'        AddPt "Polarity", 9, 0.9,-2.0
'        AddPt "Polarity", 10, 0.95, -1.5
'        AddPt "Polarity", 11, 1, -1.0
'        AddPt "Polarity", 12, 1.05, -0.5
'        AddPt "Polarity", 13, 1.1, 0
'        AddPt "Polarity", 14, 1.3, 0
'
'        addpt "Velocity", 0, 0,         1
'        addpt "Velocity", 1, 0.16, 1.06
'        addpt "Velocity", 2, 0.41,         1.05
'        addpt "Velocity", 3, 0.53,         1'0.982
'        addpt "Velocity", 4, 0.702, 0.968
'        addpt "Velocity", 5, 0.95,  0.968
'        addpt "Velocity", 6, 1.03,         0.945
'
'        LF.Object = LeftFlipper
'        LF.EndPoint = EndPointLp
'        RF.Object = RightFlipper
'        RF.EndPoint = EndPointRp
'End Sub


' Flipper trigger hit subs
Sub TriggerLF_Hit() : LF.Addball activeball : End Sub
Sub TriggerLF_UnHit() : LF.PolarityCorrect activeball : End Sub
Sub TriggerRF_Hit() : RF.Addball activeball : End Sub
Sub TriggerRF_UnHit() : RF.PolarityCorrect activeball : End Sub

'Methods:
'.TimeDelay - Delay before trigger shuts off automatically. Default = 80 (ms)
'.AddPoint - "Polarity", "Velocity", "Ycoef" coordinate points. Use one of these 3 strings, keep coordinates sequential. x = %position on the flipper, y = output
'.Object - set to flipper reference. Optional.
'.StartPoint - set start point coord. Unnecessary, if .object is used.

'Called with flipper -
'ProcessBalls - catches ball data.
' - OR -
'.Fire - fires flipper.rotatetoend automatically + processballs. Requires .Object to be set to the flipper.


'******************************************************
'  FLIPPER CORRECTION FUNCTIONS
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

'********Triggered by a ball hitting the flipper trigger area
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

'*********Used to rotate flipper since this is removed from the key down for the flippers
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
    PartialFlipCoef = ((Flipper.StartAngle - Flipper.CurrentAngle) / (Flipper.StartAngle - Flipper.EndAngle))   '% of flipper swing
    PartialFlipCoef = abs(PartialFlipCoef-1)
  End Sub

'***********gameTime is a global variable of how long the game has progressed in ms
'***********This function lets the table know if the flipper has been fired
  Private Function FlipperOn() : if gameTime < FlipAt+TimeDelay then FlipperOn = True : End If : End Function        'Timer shutoff for polaritycorrect

'***********This is turned on when a ball leaves the flipper trigger area
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
        VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)

        if partialflipcoef < 1 then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)

        if Enabled then aBall.Velx = aBall.Velx*VelCoef
        if Enabled then aBall.Vely = aBall.Vely*VelCoef
      End If

      'Polarity Correction (optional now)
      if not IsEmpty(PolarityIn(0) ) then
        If StartPoint > EndPoint then LR = -1        'Reverse polarity if left flipper
        dim AddX : AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR

        if Enabled then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef)
      End If
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
  for x = 0 to uBound(aArray)        'Shuffle objects in a temp array
    if not IsEmpty(aArray(x) ) Then
      if IsObject(aArray(x)) then
        Set a(aCount) = aArray(x)   'Set creates an object in VB
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
'**********Takes in more than one array and passes them to ShuffleArray
Sub ShuffleArrays(aArray1, aArray2, offset)
  ShuffleArray aArray1, offset
  ShuffleArray aArray2, offset
End Sub

' Used for flipper correction, rubber dampeners, and drop targets
'**********Calculate ball speed as hypotenuse of velX/velY triangle
Function BallSpeed(ball) 'Calculates the ball speed
  BallSpeed = SQR(ball.VelX^2 + ball.VelY^2 + ball.VelZ^2)
End Function

' Used for flipper correction and rubber dampeners
'**********Calculates the value of Y for an input x using the slope intercept equation
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
'********Interpolates the value for areas between the low and upper bounds sent to it
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
'  FLIPPER TRICKS
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
  Dim b, BOT
  BOT = GetBalls

  If Flipper1.currentangle = Endangle1 and EOSNudge1 <> 1 Then
    EOSNudge1 = 1
    debug.print Flipper1.currentangle &" = "& Endangle1 &"--"& Flipper2.currentangle &" = "& EndAngle2
    If Flipper2.currentangle = EndAngle2 Then
      For b = 0 to Ubound(BOT)
        If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper1) Then
          Debug.Print "ball in flip1. exit"
          exit Sub
        end If
      Next
      For b = 0 to Ubound(BOT)
        If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper2) Then
          BOT(b).velx = BOT(b).velx / 1.3
          BOT(b).vely = BOT(b).vely - 0.5
        end If
      Next
    End If
  Else
    If Abs(Flipper1.currentangle) > Abs(EndAngle1) + 30 then EOSNudge1 = 0
  End If
End Sub

'*****************
' Maths
'*****************
Dim PI: PI = 4*Atn(1)

Function dSin(degrees)
  dsin = sin(degrees * Pi/180)
End Function

Function dCos(degrees)
  dcos = cos(degrees * Pi/180)
End Function

Function Atn2(dy, dx)
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
'  Check ball distance from Flipper for Rem
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


'*************************************************
'  End - Check ball distance from Flipper for Rem
'*************************************************

dim LFPress, RFPress, LFCount, RFCount
dim LFState, RFState
dim EOST, EOSA,Frampup, FElasticity,FReturn
dim RFEndAngle, LFEndAngle

Const FlipperCoilRampupMode = 1     '0 = fast, 1 = medium, 2 = slow (tap passes should work)

LFState = 1
RFState = 1
EOST = leftflipper.eostorque      'End of Swing Torque
EOSA = leftflipper.eostorqueangle   'End of Swing Torque Angle
Frampup = LeftFlipper.rampup      'Flipper Stregth Ramp Up
FElasticity = LeftFlipper.elasticity  'Flipper Elasticity
FReturn = LeftFlipper.return      'Flipper Return Strength
'Const EOSTnew = 1 'EM's to late 80's
Const EOSTnew = 0.8 '90's and later
Const EOSAnew = 1
Const EOSRampup = 0
Dim SOSRampup
Select Case FlipperCoilRampupMode     'determines strength of coil field at start of swing
  Case 0:
    SOSRampup = 2.5
  Case 1:
    SOSRampup = 6
  Case 2:
    SOSRampup = 8.5
End Select

Const LiveCatch = 16          'variable to check elapsed time
Const LiveElasticity = 0.45
Const SOSEM = 0.815
Const EOSReturn = 0.055  'EM's
'Const EOSReturn = 0.045  'late 70's to mid 80's
'Const EOSReturn = 0.035  'mid 80's to early 90's
'Const EOSReturn = 0.025  'mid 90's and later

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
    Dim b, BOT
    BOT = GetBalls

    For b = 0 to UBound(BOT)
      If Distance(BOT(b).x, BOT(b).y, Flipper.x, Flipper.y) < 55 Then 'check for cradle
        If BOT(b).vely >= -0.4 Then BOT(b).vely = -0.4
      End If
    Next
  End If
End Sub

Sub FlipperTricks (Flipper, FlipperPress, FCount, FEndAngle, FState)
  'What this code does is swing the flipper fast and make the flipper soft near its EOS to enable live catches.  It resets back to the base Table
  'settings once the flipper reaches the end of swing.  The code also makes the flipper starting ramp up high to simulate the stronger starting
  'coil strength and weaker at its EOS to simulate the weaker hold coil.

  Dim Dir
  Dir = Flipper.startangle/Abs(Flipper.startangle)        '-1 for Right Flipper

  If Abs(Flipper.currentangle) > Abs(Flipper.startangle) - 0.05 Then  'If the flipper has started its swing, make it swing fast to nearly the end...
    If FState <> 1 Then
      Flipper.rampup = SOSRampup                  'set flipper Ramp Up high
      Flipper.endangle = FEndAngle - 3*Dir            'swing to within 3 degrees of EOS
      Flipper.Elasticity = FElasticity * SOSEM          'Set the elasticity to the base table elasticity
      FCount = 0
      FState = 1
    End If
  ElseIf Abs(Flipper.currentangle) <= Abs(Flipper.endangle) and FlipperPress = 1 then   'If the flipper is fully swung and the flipper button is pressed then
    if FCount = 0 Then FCount = GameTime        'notes the Game Time to see if a live catch is possible

    If FState <> 2 Then
      Flipper.eostorqueangle = EOSAnew        'sets flipper EOS Torque Angle to .2
      Flipper.eostorque = EOSTnew           'sets flipper EOS Torque to 1
      Flipper.rampup = EOSRampup
      Flipper.endangle = FEndAngle
      FState = 2
    End If
  Elseif Abs(Flipper.currentangle) > Abs(Flipper.endangle) + 0.01 and FlipperPress = 1 Then 'If the flipper has swung past it's end of swing then..
    If FState <> 3 Then
      Flipper.eostorque = EOST                        'set the flipper EOS Torque back to the base table setting
      Flipper.eostorqueangle = EOSA                 'set the flipper EOS Torque Angle back to the base table setting
      Flipper.rampup = Frampup                    'set the flipper Ramp Up back to the base table setting
      Flipper.Elasticity = FElasticity                'set the flipper Elasticity back to the base table setting
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


'******************************************************
'****  END FLIPPER CORRECTIONS
'******************************************************

'******************************************************
'****  PHYSICS DAMPENERS
'******************************************************
'
' These are data mined bounce curves,
' dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
' Requires tracking ballspeed to calculate COR



Sub dPosts_Hit(idx)
  RubbersD.dampen Activeball
  TargetBouncer Activeball, 1
End Sub

Sub dSleeves_Hit(idx)
  SleevesD.Dampen Activeball
  TargetBouncer Activeball, 0.7
End Sub

'*********This sets up the rubbers:
dim RubbersD : Set RubbersD = new Dampener     'Makes a Dampener Class Object
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
'            Uses the LinearEnvelope function to calculate the correction based upon where it's value sits in relation
'            to the addpoint parameters set above.  Basically interpolates values between set points in a linear fashion
    DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
'            Uses the function BallSpeed's value at the point of impact/the active ball's velocity which is constantly being updated
'        RealCor is always less than 1
    RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id)+0.0001)
'                Divides the desired CoR by the real COR to make a multiplier to correct velocity in x and y
    coef = desiredcor / realcor
    if debugOn then str = name & " in vel:" & round(cor.ballvel(aBall.id),2 ) & vbnewline & "desired cor: " & round(desiredcor,4) & vbnewline & _
    "actual cor: " & round(realCOR,4) & vbnewline & "ballspeed coef: " & round(coef, 3) & vbnewline
    if Print then debug.print Round(cor.ballvel(aBall.id),2) & ", " & round(desiredcor,3)
'                 Applies the coef to x and y velocities
' Thalamus - patched :     aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
    aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
    if debugOn then TBPout.text = str
  End Sub

  public sub Dampenf(aBall, parm) 'Rubberizer is handled here
    dim RealCOR, DesiredCOR, str, coef
    DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
    RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id)+0.0001)
    coef = desiredcor / realcor
    If abs(aball.velx) < 2 and aball.vely < 0 and aball.vely > -3.75 then
' Thalamus - patched :       aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
    aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
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
'  TRACK ALL BALL VELOCITIES
'  FOR RUBBER DAMPENER AND DROP TARGETS
'******************************************************
'*********CoR is Coefficient of Restitution defined as "how much of the kinetic energy remains for the objects
'to rebound from one another vs. how much is lost as heat, or work done deforming the objects
dim cor : set cor = New CoRTracker

Class CoRTracker
  public ballvel, ballvelx, ballvely

  Private Sub Class_Initialize : redim ballvel(0) : redim ballvelx(0): redim ballvely(0) : End Sub

  Public Sub Update() 'tracks in-ball-velocity
    dim str, b, AllBalls, highestID : allBalls = getballs

    for each b in allballs
      if b.id >= HighestID then highestID = b.id
    Next

    if uBound(ballvel) < highestID then redim ballvel(highestID)  'set bounds
    if uBound(ballvelx) < highestID then redim ballvelx(highestID)  'set bounds
    if uBound(ballvely) < highestID then redim ballvely(highestID)  'set bounds

    for each b in allballs
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



' ****** Put these at the bottom of your table script if you transfer the clock and beer to your room *******

' ***** Beer Bubble Code - Rawd *****
Sub BeerTimer_Timer()

Randomize(21)
BeerBubble1.z = BeerBubble1.z + Rnd(1)*0.5
if BeerBubble1.z > -771 then BeerBubble1.z = -955
BeerBubble2.z = BeerBubble2.z + Rnd(1)*1
if BeerBubble2.z > -768 then BeerBubble2.z = -955
BeerBubble3.z = BeerBubble3.z + Rnd(1)*1
if BeerBubble3.z > -768 then BeerBubble3.z = -955
BeerBubble4.z = BeerBubble4.z + Rnd(1)*0.75
if BeerBubble4.z > -774 then BeerBubble4.z = -955
BeerBubble5.z = BeerBubble5.z + Rnd(1)*1
if BeerBubble5.z > -771 then BeerBubble5.z = -955
BeerBubble6.z = BeerBubble6.z + Rnd(1)*1
if BeerBubble6.z > -774 then BeerBubble6.z = -955
BeerBubble7.z = BeerBubble7.z + Rnd(1)*0.8
if BeerBubble7.z > -768 then BeerBubble7.z = -955
BeerBubble8.z = BeerBubble8.z + Rnd(1)*1
if BeerBubble8.z > -771 then BeerBubble8.z = -955
End Sub



' ***************** VR Clock code below - THANKS RASCAL ******************
Dim CurrentMinute ' for VR clock
' VR Clock code below....
Sub ClockTimer_Timer()
  Pminutes.RotAndTra2 = (Minute(Now())+(Second(Now())/100))*6
  Phours.RotAndTra2 = Hour(Now())*30+(Minute(Now())/2)
    Pseconds.RotAndTra2 = (Second(Now()))*6
  CurrentMinute=Minute(Now())
End Sub
 ' ********************** END CLOCK CODE   *********************************


'***************************************************************
'****  VPW DYNAMIC BALL SHADOWS by Iakki, Apophis, and Wylte
'***************************************************************

' This subroutine updates the flipper shadows and visual primitives
Sub FlipperVisualUpdate
' lFlip.rotz = leftflipper.currentangle
' rFlip.rotz = rightflipper.currentangle
  pGate.rotx = Gate.CurrentAngle * 0.6
  FlipperLSh.RotZ = LeftFlipper.CurrentAngle
  FlipperRSh.RotZ = RightFlipper.CurrentAngle
End Sub

'****** INSTRUCTIONS please read ******

'****** Part A:  Table Elements ******
'
' Import the "bsrtx7" and "ballshadow" images
' Import the shadow materials file (3 sets included) (you can also export the 3 sets from this table to create the same file)
' Copy in the BallShadowA flasher set and the sets of primitives named BallShadow#, RtxBallShadow#, and RtxBall2Shadow#
' * with at least as many objects each as there can be balls, including locked balls
' Ensure you have a timer with a -1 interval that is always running

' Create a collection called DynamicSources that includes all light sources you want to cast ball shadows
'***These must be organized in order, so that lights that intersect on the table are adjacent in the collection***
'***If there are more than 3 lights that overlap in a playable area, exclude the less important lights***
' This is because the code will only project two shadows if they are coming from lights that are consecutive in the collection, and more than 3 will cause "jumping" between which shadows are drawn
' The easiest way to keep track of this is to start with the group on the right slingshot and move anticlockwise around the table
' For example, if you use 6 lights: A & B on the left slingshot and C & D on the right, with E near A&B and F next to C&D, your collection would look like EBACDF
'
'G        H                     ^ E
'                             ^ B
' A    C                        ^ A
'  B    D     your collection should look like  ^ G   because E&B, B&A, etc. intersect; but B&D or E&F do not
'  E      F                       ^ H
'                             ^ C
'                             ^ D
'                             ^ F
'   When selecting them, you'd shift+click in this order^^^^^

'****** End Part A:  Table Elements ******


'****** Part B:  Code and Functions ******

' *** Timer sub
' The "DynamicBSUpdate" sub should be called by a timer with an interval of -1 (framerate)
'Sub FrameTimer_Timer()
' If DynamicBallShadowsOn Or AmbientBallShadowOn Then DynamicBSUpdate 'update ball shadows
'End Sub

' *** These are usually defined elsewhere (ballrolling), but activate here if necessary
'Const tnob = 10 ' total number of balls
Const lob = 0 'locked balls on start; might need some fiddling depending on how your locked balls are done
dim gilvl:gilvl = 1
'Dim tablewidth: tablewidth = Table1.width
'Dim tableheight: tableheight = Table1.height

' *** User Options - Uncomment here or move to top
'----- Shadow Options -----
'Const DynamicBallShadowsOn = 1   '0 = no dynamic ball shadow ("triangles" near slings and such), 1 = enable dynamic ball shadow
'Const AmbientBallShadowOn = 1    '0 = Static shadow under ball ("flasher" image, like JP's)
'                 '1 = Moving ball shadow ("primitive" object, like ninuzzu's) - This is the only one that shows up on the pf when in ramps and fades when close to lights!
'                 '2 = flasher image shadow, but it moves like ninuzzu's

Const fovY          = 0   'Offset y position under ball to account for layback or inclination (more pronounced need further back)
Const DynamicBSFactor     = 0.95  '0 to 1, higher is darker
Const AmbientBSFactor     = 0.7 '0 to 1, higher is darker
Const AmbientMovement   = 2   '1 to 4, higher means more movement as the ball moves left and right
Const Wideness        = 20  'Sets how wide the dynamic ball shadows can get (20 +5 thinness should be most realistic for a 50 unit ball)
Const Thinness        = 5   'Sets minimum as ball moves away from source

' *** Required Functions, enable these if they are not already present elswhere in your table
Function DistanceFast(x, y)
  dim ratio, ax, ay
  ax = abs(x)         'Get absolute value of each vector
  ay = abs(y)
  ratio = 1 / max(ax, ay)   'Create a ratio
  ratio = ratio * (1.29289 - (ax + ay) * ratio * 0.29289)
  if ratio > 0 then     'Quickly determine if it's worth using
    DistanceFast = 1/ratio
  Else
    DistanceFast = 0
  End if
end Function

Function max(a,b)
  if a > b then
    max = a
  Else
    max = b
  end if
end Function

'Dim PI: PI = 4*Atn(1)

'Function Atn2(dy, dx)
' If dx > 0 Then
'   Atn2 = Atn(dy / dx)
' ElseIf dx < 0 Then
'   If dy = 0 Then
'     Atn2 = pi
'   Else
'     Atn2 = Sgn(dy) * (pi - Atn(Abs(dy / dx)))
'   end if
' ElseIf dx = 0 Then
'   if dy = 0 Then
'     Atn2 = 0
'   else
'     Atn2 = Sgn(dy) * pi / 2
'   end if
' End If
'End Function
'
'Function AnglePP(ax,ay,bx,by)
' AnglePP = Atn2((by - ay),(bx - ax))*180/PI
'End Function

'****** End Part B:  Code and Functions ******


'****** Part C:  The Magic ******
Dim sourcenames, currentShadowCount, DSSources(30), numberofsources, numberofsources_hold
sourcenames = Array ("","","","","","","","","","","","")
currentShadowCount = Array (0,0,0,0,0,0,0,0,0,0,0,0)

' *** Trim or extend these to match the number of balls/primitives/flashers on the table!
dim objrtx1(12), objrtx2(12)
dim objBallShadow(12)
Dim BallShadowA
BallShadowA = Array (BallShadowA0,BallShadowA1,BallShadowA2,BallShadowA3,BallShadowA4)

DynamicBSInit

sub DynamicBSInit()
  Dim iii, source

  for iii = 0 to tnob                 'Prepares the shadow objects before play begins
    Set objrtx1(iii) = Eval("RtxBallShadow" & iii)
    objrtx1(iii).material = "RtxBallShadow" & iii
    objrtx1(iii).z = iii/1000 + 0.01
    objrtx1(iii).visible = 0

    Set objrtx2(iii) = Eval("RtxBall2Shadow" & iii)
    objrtx2(iii).material = "RtxBallShadow2_" & iii
    objrtx2(iii).z = (iii)/1000 + 0.02
    objrtx2(iii).visible = 0

    currentShadowCount(iii) = 0

    Set objBallShadow(iii) = Eval("BallShadow" & iii)
    objBallShadow(iii).material = "BallShadow" & iii
    UpdateMaterial objBallShadow(iii).material,1,0,0,0,0,0,AmbientBSFactor,RGB(0,0,0),0,0,False,True,0,0,0,0
    objBallShadow(iii).Z = iii/1000 + 0.04
    objBallShadow(iii).visible = 0

    BallShadowA(iii).Opacity = 100*AmbientBSFactor
    BallShadowA(iii).visible = 0
  Next

  iii = 0

  For Each Source in DynamicSources
    DSSources(iii) = Array(Source.x, Source.y)
    iii = iii + 1
  Next
  numberofsources = iii
  numberofsources_hold = iii
end sub


Sub DynamicBSUpdate
  Dim falloff:  falloff = 150     'Max distance to light sources, can be changed if you have a reason
  Dim ShadowOpacity, ShadowOpacity2
  Dim s, Source, LSd, currentMat, AnotherSource, BOT, iii
  BOT = GetBalls

  'Hide shadow of deleted balls
  For s = UBound(BOT) + 1 to tnob
    objrtx1(s).visible = 0
    objrtx2(s).visible = 0
    objBallShadow(s).visible = 0
    BallShadowA(s).visible = 0
  Next

  If UBound(BOT) < lob Then Exit Sub    'No balls in play, exit

'The Magic happens now
  For s = lob to UBound(BOT)

' *** Normal "ambient light" ball shadow
  'Layered from top to bottom. If you had an upper pf at for example 80 and ramps even above that, your segments would be z>110; z<=110 And z>100; z<=100 And z>30; z<=30 And z>20; Else invisible

    If AmbientBallShadowOn = 1 Then     'Primitive shadow on playfield, flasher shadow in ramps
      If BOT(s).Z > 30 Then             'The flasher follows the ball up ramps while the primitive is on the pf
        If BOT(s).X < tablewidth/2 Then
          objBallShadow(s).X = ((BOT(s).X) - (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) + 5
        Else
          objBallShadow(s).X = ((BOT(s).X) + (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) - 5
        End If
        objBallShadow(s).Y = BOT(s).Y + BallSize/10 + fovY
        objBallShadow(s).visible = 1

        BallShadowA(s).X = BOT(s).X
        BallShadowA(s).Y = BOT(s).Y + BallSize/5 + fovY
        BallShadowA(s).height=BOT(s).z - BallSize/4   'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping
        BallShadowA(s).visible = 1
      Elseif BOT(s).Z <= 30 And BOT(s).Z > 20 Then  'On pf, primitive only
        objBallShadow(s).visible = 1
        If BOT(s).X < tablewidth/2 Then
          objBallShadow(s).X = ((BOT(s).X) - (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) + 5
        Else
          objBallShadow(s).X = ((BOT(s).X) + (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) - 5
        End If
        objBallShadow(s).Y = BOT(s).Y + fovY
        BallShadowA(s).visible = 0
      Else                      'Under pf, no shadows
        objBallShadow(s).visible = 0
        BallShadowA(s).visible = 0
      end if

    Elseif AmbientBallShadowOn = 2 Then   'Flasher shadow everywhere
      If BOT(s).Z > 30 Then             'In a ramp
        BallShadowA(s).X = BOT(s).X
        BallShadowA(s).Y = BOT(s).Y + BallSize/5 + fovY
        BallShadowA(s).height=BOT(s).z - BallSize/4   'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping
        BallShadowA(s).visible = 1
      Elseif BOT(s).Z <= 30 And BOT(s).Z > 20 Then  'On pf
        BallShadowA(s).visible = 1
        If BOT(s).X < tablewidth/2 Then
          BallShadowA(s).X = ((BOT(s).X) - (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) + 5
        Else
          BallShadowA(s).X = ((BOT(s).X) + (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) - 5
        End If
        BallShadowA(s).Y = BOT(s).Y + Ballsize/10 + fovY
        BallShadowA(s).height=BOT(s).z - BallSize/2 + 5
      Else                      'Under pf
        BallShadowA(s).visible = 0
      End If
    End If

' *** Dynamic shadows
    If DynamicBallShadowsOn Then
      If BOT(s).Z < 30 Then 'And BOT(s).Y < (TableHeight - 200) Then 'Or BOT(s).Z > 105 Then    'Defining when and where (on the table) you can have dynamic shadows
        For iii = 0 to numberofsources - 1
          LSd=DistanceFast((BOT(s).x-DSSources(iii)(0)),(BOT(s).y-DSSources(iii)(1))) 'Calculating the Linear distance to the Source
          If LSd < falloff And gilvl > 0 Then               'If the ball is within the falloff range of a light and light is on (we will set numberofsources to 0 when GI is off)
            currentShadowCount(s) = currentShadowCount(s) + 1   'Within range of 1 or 2
            if currentShadowCount(s) = 1 Then           '1 dynamic shadow source
              sourcenames(s) = iii
              currentMat = objrtx1(s).material
              objrtx2(s).visible = 0 : objrtx1(s).visible = 1 : objrtx1(s).X = BOT(s).X : objrtx1(s).Y = BOT(s).Y + fovY
  '           objrtx1(s).Z = BOT(s).Z - 25 + s/1000 + 0.01            'Uncomment if you want to add shadows to an upper/lower pf
              objrtx1(s).rotz = AnglePP(DSSources(iii)(0), DSSources(iii)(1), BOT(s).X, BOT(s).Y) + 90
              ShadowOpacity = (falloff-LSd)/falloff                 'Sets opacity/darkness of shadow by distance to light
              objrtx1(s).size_y = Wideness*ShadowOpacity+Thinness           'Scales shape of shadow with distance/opacity
              UpdateMaterial currentMat,1,0,0,0,0,0,ShadowOpacity*DynamicBSFactor^2,RGB(0,0,0),0,0,False,True,0,0,0,0
              If AmbientBallShadowOn = 1 Then
                currentMat = objBallShadow(s).material                  'Brightens the ambient primitive when it's close to a light
                UpdateMaterial currentMat,1,0,0,0,0,0,AmbientBSFactor*(1-ShadowOpacity),RGB(0,0,0),0,0,False,True,0,0,0,0
              Else
                BallShadowA(s).Opacity = 100*AmbientBSFactor*(1-ShadowOpacity)
              End If

            Elseif currentShadowCount(s) = 2 Then
                                  'Same logic as 1 shadow, but twice
              currentMat = objrtx1(s).material
              AnotherSource = sourcenames(s)
              objrtx1(s).visible = 1 : objrtx1(s).X = BOT(s).X : objrtx1(s).Y = BOT(s).Y + fovY
  '           objrtx1(s).Z = BOT(s).Z - 25 + s/1000 + 0.01              'Uncomment if you want to add shadows to an upper/lower pf
              objrtx1(s).rotz = AnglePP(DSSources(AnotherSource)(0),DSSources(AnotherSource)(1), BOT(s).X, BOT(s).Y) + 90
              ShadowOpacity = (falloff-DistanceFast((BOT(s).x-DSSources(AnotherSource)(0)),(BOT(s).y-DSSources(AnotherSource)(1))))/falloff
              objrtx1(s).size_y = Wideness*ShadowOpacity+Thinness
              UpdateMaterial currentMat,1,0,0,0,0,0,ShadowOpacity*DynamicBSFactor^3,RGB(0,0,0),0,0,False,True,0,0,0,0

              currentMat = objrtx2(s).material
              objrtx2(s).visible = 1 : objrtx2(s).X = BOT(s).X : objrtx2(s).Y = BOT(s).Y + fovY
  '           objrtx2(s).Z = BOT(s).Z - 25 + s/1000 + 0.02              'Uncomment if you want to add shadows to an upper/lower pf
              objrtx2(s).rotz = AnglePP(DSSources(iii)(0), DSSources(iii)(1), BOT(s).X, BOT(s).Y) + 90
              ShadowOpacity2 = (falloff-LSd)/falloff
              objrtx2(s).size_y = Wideness*ShadowOpacity2+Thinness
              UpdateMaterial currentMat,1,0,0,0,0,0,ShadowOpacity2*DynamicBSFactor^3,RGB(0,0,0),0,0,False,True,0,0,0,0
              If AmbientBallShadowOn = 1 Then
                currentMat = objBallShadow(s).material                  'Brightens the ambient primitive when it's close to a light
                UpdateMaterial currentMat,1,0,0,0,0,0,AmbientBSFactor*(1-max(ShadowOpacity,ShadowOpacity2)),RGB(0,0,0),0,0,False,True,0,0,0,0
              Else
                BallShadowA(s).Opacity = 100*AmbientBSFactor*(1-max(ShadowOpacity,ShadowOpacity2))
              End If
            end if
          Else
            currentShadowCount(s) = 0
            BallShadowA(s).Opacity = 100*AmbientBSFactor
          End If
        Next
      Else                  'Hide dynamic shadows everywhere else
        objrtx2(s).visible = 0 : objrtx1(s).visible = 0
      End If
    End If
  Next
End Sub
'****************************************************************
'****  END VPW DYNAMIC BALL SHADOWS by Iakki, Apophis, and Wylte
'****************************************************************

'******************************************************
' VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************

Const TargetBouncerEnabled = 1    '0 = normal standup targets, 1 = bouncy targets, 2 = orig TargetBouncer
Const TargetBouncerFactor = 0.7   'Level of bounces. Recommmended value of 0.7 when TargetBouncerEnabled=1, and 1.1 when TargetBouncerEnabled=2

sub TargetBouncer(aBall,defvalue)  'use defvalue to manipulate the bounce of the ball/target interaction
    dim zMultiplier, vel, vratio
    if TargetBouncerEnabled = 1 and aball.z < 30 then
        'debug.print "velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz
        vel = BallSpeed(aBall)
        if aBall.velx = 0 then vratio = 1 else vratio = aBall.vely/aBall.velx
        Select Case Int(Rnd * 6) + 1
            Case 1: zMultiplier = 0.2*defvalue
      Case 2: zMultiplier = 0.25*defvalue
            Case 3: zMultiplier = 0.3*defvalue
      Case 4: zMultiplier = 0.4*defvalue
            Case 5: zMultiplier = 0.45*defvalue
            Case 6: zMultiplier = 0.5*defvalue
        End Select
        aBall.velz = abs(vel * zMultiplier * TargetBouncerFactor)
        aBall.velx = sgn(aBall.velx) * sqr(abs((vel^2 - aBall.velz^2)/(1+vratio^2)))
        aBall.vely = aBall.velx * vratio
        'debug.print "---> velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz
        'debug.print "conservation check: " & BallSpeed(aBall)/vel
    elseif TargetBouncerEnabled = 2 and aball.z < 30 then
    'debug.print "velz: " & activeball.velz
    Select Case Int(Rnd * 4) + 1
      Case 1: zMultiplier = defvalue+1.1
      Case 2: zMultiplier = defvalue+1.05
      Case 3: zMultiplier = defvalue+0.7
      Case 4: zMultiplier = defvalue+0.3
    End Select
    aBall.velz = aBall.velz * zMultiplier * TargetBouncerFactor
    'debug.print "----> velz: " & activeball.velz
    'debug.print "conservation check: " & BallSpeed(aBall)/vel
  end if
end sub

Sub TargetBounce_Hit
  TargetBouncer activeball, 1
End Sub

Sub ShadowHide_Hit()
Shadowblock.visible = false
End Sub

Sub ShadowShow_Hit()
Shadowblock.visible = true
End Sub





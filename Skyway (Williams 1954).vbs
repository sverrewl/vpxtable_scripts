'****************************************************************
'
'          Skyway (Williams 1954)
'          Script by Scottacus
'              v 2.02
'              June 2022
'
' Basic DOF config
'   80  BackGlass On Image
'   101 Left Flipper
'   102 Start Button
'   105 Bumper1, 106 Bumper2, 107 DeadBumper
'   104 Right Sling
'   120 - 123 Kickers
'   130 - 137 Left Lane Wire Rollovers
'   138 - 140 Lane RollOvers
'   150 - 152 Button RollOvers
'   154 Small Bell, 155 Large Bell
'   128 Knocker
'   Code Flow
'                  EndGame
'                   ^
'   Start Game -> New Game -> Check Continue -> Release Ball -> Drain -> Score Bouns -> Advance Player -> Next Ball
'                   ^                                     |
'                 EndGame = True <---------------------------------------------------------------
'
' Note Score Motor Location for sound routines is at shootAgain Light

' Ball Control Subroutine developed by: rothbauerw
'   Press "c" during play to activate, the arrow keys control the ball

Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the Controller.vbs file in order to run this table (installed with the VPX package in the scripts folder)"
On Error Goto 0

Const cGameName = "Skyway_1954"
Const cOptions = "Skyway_1954_2.02.txt"
Const hsFileName = "Skyway (Williams 1954)"

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

'************************** New Ball Shadow Code ***************************************
Const DynamicBallShadowsOn = 1 '0 = no dynamic ball shadows, 1 = enable dynamic ball shadow
Const AmbientBallShadowOn = 1  '0 = no regular ball shadow   1 = regular ball shadow on

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

Dim tablewidth: tablewidth = Table1.width
Dim tableheight: tableheight = Table1.height

'****************************************************************************************************************************************

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
Dim i,j, f, ii, Object, Light, x, y, z
Dim freePlay
Dim ballsize,BallMass
ballSize = 50
ballMass = (Ballsize^3)/125000
Dim BIP : BIP = 0
Dim options
Dim chime
Dim onBumper
Dim pfOption
Dim lutValue

'BackGlass reel pans,  Valid values: "Lpan", "Mpan" and "Rpan" (Quotes needed, not case sensitive)
  Const bgcrReel = "Rpan"

'Desktop reel pans,  Valid values: "Lpan", "Mpan" and "Rpan" (Quotes needed, not case sensitive)
  Const dtcrReel = "Rpan"

Sub Table1_init
  LoadEM
  maxPlayers = 1
  player = 1
  loadHighScore
  If highScore(0) = "" Then highScore(0) = 500
  If highScore(1) = "" Then highScore(1) = 450
  If highScore(2) = "" Then highScore(2) = 400
  If highScore(3) = "" Then highScore(3) = 350
  If HighScore(4) = "" Then highScore(4) = 300
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
  If onBumper = "" Then onBumper = 1
  If vrOption = "" Then vrOption = 0
  If lutValue = "" Then lutValue = 0

  updatePostIt
  dynamicUpdatePostIt.enabled = 1
  TiltReel.setValue(1)
  If vrOption > 0 Then
    vrTilt.visible = 1
    vrCredit
  End If
  CreditReel.setvalue(credit)
  DTReel.setValue(0)
  Table1.ColorGradeImage = "LUT" & lutValue

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

  BootTable.enabled = 1
  replaySettings
  tilt = False
  state = False
  gameState
  ballsToPlay = 5
  laneDrain = 0

  For x = 1 to balls
    EVAL("Kicker" & x).createball
    EVAL("Kicker" & x).kick 0,0
    EVAL("Kicker" & x).enabled = False
  Next
End Sub


'***********KeyCodes
Dim enableInitialEntry, flipperPress
Sub Table1_KeyDown(ByVal keycode)

  If enableInitialEntry = True Then enterIntitals(keycode)

  If keycode = addCreditKey Then
    playFieldSound "coinin",0,Drain3,1
    addCredit
    End If

  If keycode = startGameKey and state = false and enableInitialEntry = False and vrOption > 0 Then VR_cab_buttonstart.y = -5

    If keycode = startGameKey Then
    If enableInitialEntry = False and operatormenu = 0 and backGlassOn = 1 and ballsToPlay + ballsLifted = 5 Then
      If freePlay = 1 and players < 1 Then startGame
      If freePlay = 0 and credit > 0 and players < 1 Then
        credit = credit - 1
        If showDT = False Then playReelSound "Reel5", bgcrReel Else playReelSound "Reel5", dtcrReel
        creditReel.setvalue(credit)
        If B2SOn Then
          If freeplay = 0 Then controller.B2SSetCredits credit
        End If
        startGame
      End If
    End If
  End If

  If keycode = rightFlipperKey Then
    If vrOption > 0 Then VR_cab_buttonRight.X = -8
  End If

  If keycode = leftFlipperKey Then
    If vrOption > 0 Then VR_cab_buttonLeft.X = 8
  End If


  If keycode = PlungerKey Then
    plunger.PullBack
    playFieldSound "plungerpull", 0, plunger, 1
    If vrOption > 0 Then
      TimerVRPlunger.Enabled = True
      TimerVRPlunger1.Enabled = False
'     VR_cab_plunger.Y = -128
    End If
  End If

  If Keycode = startGameKey and contBall = 0 and lifter = 0 and destroyedBalls > 4 and cycleSaucers.enabled = 0 and enableInitialEntry = False and state = True and ready = True Then
    ballLiftTimer.enabled = 1
    If vrOption > 0 Then
      VR_cab_balllift.Y = -30
    End If
  End If

  If Keycode = rightMagnaSave and contBall = 0 and lifter = 0 and destroyedBalls > 4 and cycleSaucers.enabled = 0 and enableInitialEntry = False and state = True and ready = True Then
    ballLiftTimer.enabled = 1
    If vrOption > 0 Then
      VR_cab_balllift.Y = -30
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

  If tilt = False Then
  If keycode = leftFlipperKey and contball = 0 and operatormenu = 0 and enableInitialEntry = False and state = True Then
    playFieldSound SoundFX("FlipUpL",DOFFlippers), 0, leftFlipper, 1
    flipperBack = 0
    impulseFlipper.enabled = 1
  End If

  If keycode = rightFlipperKey and contball = 0  and operatormenu = 0 and enableInitialEntry = False and state = True Then
    playFieldSound SoundFX("FlipUpL",DOFFlippers), 0, leftFlipper, 1
    flipperBack = 0
    impulseFlipper.enabled = 1
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

    If keycode = leftFlipperKey and state = False and operatorMenu = 0 and enableInitialEntry = 0 Then
        operatorMenuTimer.Enabled = true
    End If

    If keycode = leftFlipperKey and state = False and operatorMenu = 1 Then
    options = options + 1
    If vrOption = 0 Then If options = 2 Then options = 3
    If showDt = True Then If options = 3 Then options = 4 'skips non DT options
        If options = 5 Then options = 0
    optionMenu.visible = True
        playFieldSound "target", 0, SoundPointScoreMotor, 1.5
        Select Case (Options)
            Case 0:
                optionMenu.image = "FreeCoin" & freePlay
      Case 1:
        If showDT = False Then optionMenu.image = "UnderCab"
        If showDT = True Then OptionMenu.visible = False
        optionMenu1.visible = 1
        optionMenu1.image = "Sound" & pfOption
        optionMenu2.visible = 1
        optionMenu2.image = "SoundChange"
        Select Case (pfOption)
          Case 1: speaker1.visible = 1: speaker2.visible = 1: speaker3.visible = 0: speaker4.visible = 0
          Case 2: speaker5.visible = 1: speaker6.visible = 1: speaker1.visible = 0: speaker2.visible = 0
          Case 3: speaker1.visible = 1: speaker2.visible = 1: speaker3.visible = 1: speaker4.visible = 1: speaker5.visible = 0: speaker6.visible = 0
        End Select
      Case 2:
        For x = 1 to 6
          EVAL("Speaker" & x).visible = 0
        Next
        optionMenu1.visible = 0
        OptionMenu2.visible = 0
        optionMenu.image = "VR" & vrOption
      Case 3:
        optionMenu2.visible = 0
        For x = 1 to 6
          EVAL("Speaker" & x).visible = 0
        Next
        optionMenu1.image = "DOF"
        optionMenu.image = "Chime" & chime
      Case 4:
        For x = 1 to 6
          EVAL("Speaker" & x).visible = 0
        Next
        optionMenu1.visible = 0
        optionMenu.image = "SaveExit"
        optionMenu2.visible = 0
    End Select
    End If

    If keycode = RightFlipperKey and state = False and operatorMenu = 1 Then
        playFieldSound "metalHitHigh", 0, SoundPointScoreMotor, 0.2
      Select Case (options)
    Case 0:
            If freePlay = 0 Then
                freePlay = 1
        If B2SOn Then DOF 102, DOFOn
              Else
                freePlay = 0
        If credit < 1 Then If B2SOn Then DOF 102, DOFOff
            End If
            optionMenu.image= "FreeCoin" & freePlay
    Case 1:
      pfOption = pfOption + 1
      If showDt = True and pfOption = 2 Then pfOption = 3
      If pfOption = 4 Then pfOption = 1
      optionMenu1.visible = 1
      optionMenu1.image = "Sound" & pfOption
      Select Case (pfOption)
        Case 1: speaker1.visible = 1: speaker2.visible = 1: speaker3.visible = 0: speaker4.visible = 0
        Case 2: speaker5.visible = 1: speaker6.visible = 1: speaker1.visible = 0: speaker2.visible = 0
        Case 3: speaker1.visible = 1: speaker2.visible = 1: speaker3.visible = 1: speaker4.visible = 1: speaker5.visible = 0: speaker6.visible = 0
      End Select
    Case 2:
      If vrOption = 1 Then
        vrOption = 2
        For Each object in vrRoom: object.visible = 1: Next
      Else
        vrOption = 1
        For Each object in vrRoom: object.visible = 0: Next
      End If
      optionMenu.image = "VR" & vrOption
        Case 3:
            If chime = 0 Then
                chime= 1
        If B2SOn Then DOF 155, DOFPulse
              Else
                chime = 0
        pts10
            End If
      optionMenu.image = "Chime" & chime
        Case 4:
            operatorMenu = 0
            saveHighScore
      dynamicUpdatePostIt.enabled = 1
      optionMenu.image = "FreeCoin" & freePlay
            optionMenu1.visible = 0
      optionMenu.visible = 0
      optionsMenu.visible = 0
    End Select
    End If

  If Keycode = mechanicalTilt Then
    tilt = True
    TiltReel.setValue(1)
    DTReel.setValue(0)
    inserts.image = "UV1_off"
    If vrOption > 0 Then vrTilt.visible = 1
    If B2SOn Then
      controller.B2SSetTilt 1
      controller.B2SSetData 80, 0
    End If
    turnOff
    endGame = 1
    checkContinue
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
  If Keycode = 30 Then 'a' key = KnockOff Button
    If credit > 0 Then credit = credit - 1
    creditReel.setvalue(credit)
    If B2SOn Then controller.B2SSetCredits credit
  End If

  If Keycode= 31 Then 's' key
    x=x+1
    if x=10 then x=0
    reel100k.setValue(x)
  End If

  If Keycode = 33 Then 'f' key
    x=x+1

    if x=10 then x=0
    milReel.setValue(x)
  End If
'************************End Of Test Keys****************************
End Sub

Dim flipperBack
Sub impulseFlipper_timer
    If flipperPress = 0 Then LeftFlipper.rotateToEnd: DOF 101,DOFOn
    If flipperPress = 1 Then LeftFlipper.rotateToStart: DOF 101,DOFOff
    If flipperPress = 1 and flipperBack = 0 Then playFieldSound SoundFX("FlipDownL",DOFFlippers), 0, leftFlipper, 1: flipperBack = 1
    If LeftFlipper.currentAngle = LeftFlipper.endAngle Then flipperPress = flipperPress + 1
    If LeftFlipper.currentAngle = LeftFlipper.startAngle Then flipperPress = flipperPress + 1
    If flipperPress > 1 Then flipperPress = 2
End Sub

dim TestChr1, TestChr2, TestChr3, TestChr


Sub Table1_KeyUp(ByVal keycode)

  If keycode = rightFlipperKey Then
    If vrOption > 0 Then VR_cab_buttonRight.X = 0
  End If

  If keycode = leftFlipperKey and contball = 0 Then
    If vrOption > 0 Then VR_cab_buttonLeft.X = 0
  End If

    If keycode = startGameKey and state = False Then
    VR_cab_buttonstart.y = 0
  End If

  If Keycode = startGameKey and contBall = 0 and lifter = 0 and destroyedBalls > 4 and cycleSaucers.enabled = 0 and enableInitialEntry = False and state = True and ready = True Then
    If vrOption > 0 Then
      VR_cab_balllift.Y = 0
    End If
  Else
    VR_cab_buttonstart.y = 0
  End If

  If Keycode = rightMagnaSave and contBall = 0 and lifter = 0 and destroyedBalls > 4 and cycleSaucers.enabled = 0 and enableInitialEntry = False and state = True and ready = True Then
    If vrOption > 0 Then
      VR_cab_balllift.Y = 0
    End If
  End If

  If keycode = plungerKey Then
    plunger.Fire
    playFieldSound "PlungerFire", 0, plunger, 1
    If vrOption > 0 Then
      TimerVRPlunger.Enabled = False
      TimerVRPlunger1.Enabled = True
      VR_cab_plunger.Y = -128
    End If
  End If

    If keycode = leftFlipperKey Then
        operatorMenuTimer.Enabled = False
    End If

   If tilt = False Then
    If keycode = leftFlipperKey and contball = 0 and operatormenu = 0 Then
      LeftFlipper.rotateToStart
      impulseFlipper.enabled = 0
      flipperPress = 0
    End If

    If keycode = rightFlipperKey and contball = 0 and operatormenu = 0 Then
      LeftFlipper.rotateToStart
      impulseFlipper.enabled = 0
      flipperPress = 0
    End If
   End If

    If keycode = 203 then cLeft = 0' Left Arrow

    If keycode = 200 then cUp = 0' Up Arrow

    If keycode = 208 then cDown = 0' Down Arrow

    If keycode = 205 then cRight = 0' Right Arrow

    If keycode = 52 Then Zup = 0' Period
End Sub

'************** Table Boot
Dim backGlassOn, gameStart
Dim bootCount:bootCount = 0
Sub bootTable_Timer()
  bootCount = bootCount + 1

  If bootCount = 1 Then
    If vrOption > 0 Then vrBackGlass.image = "skyWayOn"
    UV1.image = "UV1"
    rsling.image = "UV1"
    inserts.image = "UV1"
    pf_off.visible = 0
    dim xx
    For each xx in GI: xx.State = 1: Next
    gameStart = 1
    If B2SOn Then
      controller.B2SSetCredits credit
      controller.B2SSetTilt 33,1
      controller.B2SSetData 80, 1 'turns the backglass image on
      If freePlay = 1 Then DOF 102, DOFOn
      If credit > 0 Then DOF 102, DOFOn
    End If
    DTReel.setValue(1)
    arrow = 1
    lightROG1.state = 1
    backGlassOn = 1
    me.enabled = False
  End If
End Sub

'***********Replay Settings
Sub replaySettings
  replay(1) = 5000000
  replay(2) = 6000000
  replay(3) = 7000000
  replay(4) = 8000000
  replay(5) = 9000000
End Sub

Sub addCredit
  credit = credit + 1
  If credit > 25 Then credit = 25
  If showDT = False Then playReelSound "Reel5", bgcrReel Else playReelSound "Reel5", dtcrReel
  creditReel.setvalue(credit)
  If vrOption > 0 Then vrCredit
  If B2SOn Then
    controller.B2SSetCredits credit
    DOF 102, DOFOn
  End If
End Sub

'***********Operator Menu
Dim operatormenu
Sub operatorMenuTimer_Timer
    If optionMenu.visible = False Then playFieldSound "target", 0, SoundPointScoreMotor, 1.5
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
Dim bonusScore, ballsLifted
Sub startGame
    For x = 0 to 5
      outOfPlay(x) = 0
    Next
    DTReel.setValue(1)
    If B2SOn Then DOF 102, DOFOff
    ballsToPlay = 5
    ballsLifted = 0
    If B2SOn Then controller.B2SSetData 80, 1


    dynamicUpdatePostIt.enabled = 0
        UV1.image = "UV1"
        rsling.image = "UV1"
        pf_off.visible = 0
    inserts.image = "UV1"
        dim xx
        For each xx in GI: xx.State = 1: Next
    updatePostIt
    If tilt = True Then
      tilt = False
      destroyedBalls = 5
    End If
    state = True
    gameState
    players = 1
    cycleSaucers.enabled = 1
    newGame
    If vrOption > 0 Then
      vrCredit
      vr100k.image = "vr100k0"
      vr10k.image = "vr10k0"
      vrMillion.image = "vrMillion0"
    End If
End Sub

'***********Timer to delay lift of first ball until one ball is destroyed on game start
'Dim starter, starterCount
'Sub gameStart_Timer
' starterCount = starterCount + 1
' If starterCount = 4 Then
'   starter = 1
'   starterCount = 0
'   gameStart.enabled = 0
' End If
'End Sub

Dim TD, ready
Sub TroughDrain_timer
  TD = TD + 1
  Select Case TD
    Case 1: blockerWall.collidable = 0
        phys_draindynamic.collidable = 0
        phys_drainstatic.collidable = 0
        drain_hinge.rotx=90
    Case 2: blockerWall.collidable = 1
        phys_draindynamic.collidable = 1
        phys_drainstatic.collidable = 1
        drain_hinge.rotx=0
        TD = 0
        Ready = True
'       If destroyedBalls < 5 Then destroyedBalls = 5 'fail safe for game start up
        TroughDrain.enabled = 0
  End Select
End Sub

'*********New Game
Dim endGame, roundHS
Sub newGame
  player = 1
    endGame = 0
  For i = 1 to 2
    EVAL("Bumper" & i).hasHitEvent = 1
  Next
  roundHS = 0
  gameState
  For x = 1 to 4
    score(x) = 0
  Next
  resetBackglass
  resetLights
  cycleSaucers.enabled = 1
  TroughDrain.enabled = 1
End Sub

'**********Check if Game Should Continue
Dim relBall, rep(5)
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
    destroyedBalls = 0
    state = False
    gameState
    DynamicUpdatePostIt.enabled = 1
    ready = False
    players = 0
    rep(1) = 0
    laneSpecialCheck = 0
    saveHighScore
    endGameDelay.enabled = 1
    If B2SOn Then
      If freePlay = 1 or credit > 0 Then
        If B2SOn Then DOF 102, DOFOn
      End If
    End If
  Else
    relBall = 1   'variable to tell score motor routine that this is a ball release run
  End If
End Sub

Sub endGameDelay_timer
  If ScoreMotor.enabled = 0 Then
    sortScores
    endGameDelay.enabled = 0
  End If
End Sub


'***************Drain and Release Ball
Dim drainActive, destroyedBalls
Sub drain_Hit()
  playFieldSound "Drain", 0, drain, 0.001
  drainActive = 1
  drain.DestroyBall 'this is used for multiball tables to avoid timing issues with balls moving in the trough
  destroyedBalls = destroyedBalls + 1
End Sub


'**********Check if Scoring Bonus is True
Dim BonusFlag
Sub ScoreBonus
  bonusScore = 0
  If BonusScore = 0 Then AdvancePlayers 'And ShootAgain.state = 0
End Sub

'**********Advance Players
Sub advancePlayers
  nextBall
End Sub

'**********Next Ball
Sub nextBall
  If ballsToPlay = 0 then
    endGame = 1
    checkContinue
  Else
    If state = True Then checkContinue
  End If
End Sub

'************Game State Check
Sub gameState
    dim xx
  If state = False Then
    stopSound "FlipBuzzLA"
    stopSound "FlipBuzzLB"
    stopSound "FlipBuzzLC"
    stopSound "FlipBuzzLD"
    stopSound "FlipBuzzRA"
    stopSound "FlipBuzzRB"
    stopSound "FlipBuzzRC"
    stopSound "FlipBuzzRD"
  Else
    TiltReel.setValue(0)
    If vrOption > 0 Then vrTilt.visible = 0
    If B2SOn Then
      controller.B2SSetTilt 33,0
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
Sub BallControlOff_Hit()
    contBallInPlay = False
End Sub

Dim outOfPlay(5), ballDraining
Sub BallControlOff1_Hit()
    contBallInPlay = False
  nextBall
End Sub


Sub phys_subway_hit
  contballinplay = False
End Sub

'***********Ball Lift Speed Limiter to Prevent Loss of Balls
Dim lifter
lifter = 0
Sub ballLiftTimer_Timer
  lifter = lifter + 1
  If lifter = 1 Then
    If ballsToPlay > 0 and ballsLifted < 5 Then
      ballsLifted = ballsLifted + 1
      playFieldSound "BallLift" , 0, Plunger, .05
      ballLifter.CreateBall
      ballLifter.kick 80, 12
    End If
  End If
  If lifter = 2 Then
    lifter = 0
    ballLiftTimer.enabled = 0
  End If
End Sub

'************Check if Ball Out of Launch Lane
Dim ballInLane
Sub ballsInPlay_hit
  If ballREnabled=1 Then
    ballREnabled = 0
    ballInLane = False
  End If
End Sub

'************** Reset Table
Sub resetTable
' gate1.collidable = 1: gates004.visible = 1: gates003.visible = 0
End Sub

'************** Bumpers and Skirt Animation
dim bump1step, bump2step, bump3step
Sub bumpers_Hit(Index)

  If tilt = False Then
    Select Case (Index)
      Case 0: addScore 10000
          EVAL ("skirt" & (index + 1)).enabled = 1
          bumper2.playHit
          PlayFieldSound SoundFX("PopBump",DOFContactors), 0, bumper1, 1
          If B2SOn Then
            DOF 105, DOFPulse
            DOF 106, DOFPulse
          End If
      Case 1: addScore 10000
          bumper1.playHit
          EVAL ("skirt" & (index + 1)).enabled = 1
          playFieldSound SoundFX("PopBump",DOFContactors), 0, bumper2, 1
          If B2SOn Then
            DOF 105, DOFPulse
            DOF 106, DOFPulse
          End If
      Case 2: ArrowUnit
          addScore 10000
          If B2SOn Then DOF 107, DOFPulse
    End Select
  End If
End Sub

Sub Skirt1_timer
  Select Case bump1Step
    Case 3: bumpskirt1.RotY=-3
    Case 4: bumpskirt1.RotY=2
    Case 5: bumpskirt1.RotY=-1
    Case 6: bumpskirt1.RotY=1
    Case 7: bumpskirt1.RotY=0: bump1Step = 0: Skirt1.enabled = 0
  End Select
  bump1Step = bump1Step + 1
End Sub

Sub Skirt2_timer
  Select Case bump2Step
    Case 3: bumpskirt2.RotY=-3
    Case 4: bumpskirt2.RotY=2
    Case 5: bumpskirt2.RotY=-1
    Case 6: bumpskirt2.RotY=1
    Case 7: bumpskirt2.RotY=0:bump2Step = 0: Skirt2.enabled = 0
  End Select
  bump2Step = bump2Step + 1
End Sub

'************** Triggers
Dim ballsToPlay, payOut, laneDrain
Sub Triggers_Hit (Index)
  If Index < 13 Then
      If B2SOn Then DOF (130 + Index), DOFPulse
  End If

  Select Case Index
    Case 0: lightRO1.state = 0: laneCheck: addScore 100000: If L119a.state = 1 Then special
    Case 1: lightRO2.state = 0: laneCheck: addScore 100000: If L119a.state = 1 Then special
    Case 2: lightRO3.state = 0: laneCheck: addScore 100000: If L119a.state = 1 Then special
    Case 3: lightRO4.state = 0: laneCheck: addScore 100000: If L119a.state = 1 Then special
    Case 4: lightRO5.state = 0: laneCheck: addScore 100000: If L119a.state = 1 Then special
    Case 5: lightRO6.state = 0: laneCheck: addScore 100000: If L119a.state = 1 Then special
    Case 6: lightRO7.state = 0: laneCheck: addScore 100000: If L119a.state = 1 Then special
    Case 7: lightRO8.state = 0: laneCheck: addScore 100000: If L119a.state = 1 Then special
    Case 8: addScore 500000
        saucerKick.enabled = 1
        AdvanceBalls.enabled = 1
        If saucer4.ballCntOver = 0 Then ballsToPlay = ballsToPlay - 1
    Case 9: addScore 100000
        laneDrain = 1
        ArrowUnit
        ballsToPlay = ballsToPlay - 1
        If WL5.state = 1 Then special
    Case 10: addScore 100000
         laneDrain = 1
         ArrowUnit
         ballsToPlay = ballsToPlay - 1
         If WL6.state = 1 Then special
    Case 11: ArrowUnit
    Case 12: addScore 500000: AdvanceBalls.enabled = 1
         If saucer4.ballCntOver = 0 Then ballsToPlay = ballsToPlay - 1
         nextBall
    Case 13:
    Case 14:
    Case 15:
    Case 16:
    Case 17: saucer5Kick.enabled = 1
        If WL3.state = 0 Then
          addScore 500000
        Else
          payOut = 1: scoreMotor.enabled = 1
        End If
    Case 18: If destroyedBalls > 4 Then
          ballsLifted = ballsLifted - 1
          WL4.state = 1
          WL3.state = 1
         End If
    Case 19: If B2SOn Then DOF 150, DOFPulse
         If WL1.state = 1 Then
          addScore 100000:
         Else
          addScore 10000
         End If
    Case 20: If B2SOn Then DOF 151, DOFPulse
         If WL2.state = 1 Then
          addScore 100000:
         Else
          addScore 10000
         End If
    Case 21: If B2SOn Then DOF 152, DOFPulse
         addScore 100000
         spotArrow
         If WL4.state = 1 Then advanceBonus
    Case 22: ballsToPlay = ballsToPlay - 1
  End Select
End Sub

Dim freeBall
Sub Saucer4_hit
  freeBall = 1
End Sub

Sub scoringswitches1pt_hit
  ArrowUnit
End Sub

Dim laneSpecialCheck
Sub laneCheck
  If lightRO1.state = 0 And lightRO2.state = 0 And lightRO3.state = 0 And lightRO4.state = 0 Then
    WL1.state = 1
  End If
  If lightRO5.state = 0 And lightRO6.state = 0 And lightRO7.state = 0 And lightRO8.state = 0 Then
    WL2.state = 1
  End If
  If WL1.state = 1 And WL2.state = 1 And laneSpecialCheck = 0 Then
    laneSpecial
    WL5.state = 1: WL6.state = 1: L119a.state = 1:UV2.image="uv2lit"
  End If
End Sub

Sub resetLights
  For x = 1 to 6
    EVAL ("WL" & x).state = 0
  Next
  For x = 1 to 8
    EVAL("LightRO" & x).state = 1
  Next
  L119a.state = 0
  uv2.image="uv2"
  For x = 1 to 11
    EVAL( "LightAdvance" & x).state = 0
  Next
End Sub

Dim ASB
Sub advanceBonus
  ASB = ASB + 1
  For x = 1 to 11
    EVAL ("LightAdvance" & x).state = 0
  Next
  If ASB < 11 Then
    EVAL ("LightAdvance" & ASB).state = 1
  ElseIf ASB > 10 And ASB < 20 Then
    LightAdvance10.state = 1
    EVAL ("LightAdvance" & ASB-10).state = 1
  Else
    ASB = 20
    LightAdvance11.state = 1
  End If
End Sub


Dim Arrow
Sub ArrowUnit
  If laneDrain = 0 Then addScore 10000
  laneDrain = 0
  Arrow = Arrow + 1
  If Arrow > 8 Then Arrow = 1
  playFieldSound "Stepper",0, bumper1, 1
  For x = 1 to 8
    EVAL ("lightROG" & x).state = 0
  Next
  Select Case Arrow
    Case 1: lightROG1.state = 1
    Case 2: lightROG8.state = 1
    Case 3: lightROG3.state = 1
    Case 4: lightROG7.state = 1
    Case 5: lightROG2.state = 1
    Case 6: lightROG4.state = 1
    Case 7: lightROG6.state = 1
    Case 8: lightROG5.state = 1
  End Select
End Sub

Sub spotArrow
  For x = 1 to 8
    If EVAL ("LightROG" & x).state = 1 Then
      EVAL ("LightRO" & x).state = 0
    End If
  Next
  laneCheck
End Sub

Sub AdvanceBalls_timer
  If saucer1.BallCntOver Or saucer2.BallCntOver Or saucer3.BallCntOver Or saucer4.BallCntOver Then kickSaucerLaneBall
  AdvanceBalls.enabled = 0
End Sub

'************** RollOvers
Dim buttonNumber
Sub rollOver_Hit (Index)
  buttonNumber = (Index)
  rollOverAnimation.enabled = 1
End Sub

Dim button
Sub rollOverAnimation_Timer
  button = button + 1
  Select Case button
    Case 1: EVAL ("RObutton" & buttonNumber + 1).transz = -2
    Case 2: EVAL ("RObutton" & buttonNumber + 1).transz = 0
    Case 3: EVAL ("RObutton" & buttonNumber + 1).transz = 1
    Case 4: EVAL ("RObutton" & buttonNumber + 1).transz = 1
        button = 0
        rollOverAnimation.enabled = 0
  End Select
End Sub


'************ Kickers
Sub Kickers_hit(Index)
  Select Case Index:
    Case 0: wireRampKicker.timerenabled = 1
    Case 1: leftlaneKicker.timerenabled = 1
  End Select
End Sub

Dim rampDelay
Sub wirerampkicker_timer
  rampDelay = rampDelay + 1
  Select Case rampDelay
    Case 10: wireRampKicker.kick 0,30
         playFieldSound "SaucerKick",0, wireRampKicker, 1
         If B2SOn Then DOF 121, DOFPulse
         sling002.rotx = 24
    Case 11: sling002.rotx = 0
         rampDelay = 0
         wireRampkicker.timerenabled=0
  End Select
End Sub

Dim laneDelay
Sub leftlanekicker_timer
  laneDelay = laneDelay + 1
  Select Case laneDelay
    Case 10: leftLaneKicker.kick 0, 31 + (RND * 6)
         playFieldSound "SaucerKick",0, leftLaneKicker, 1
         If B2SOn Then DOF 120, DOFPulse
         sling001.rotx = 24
    Case 11: sling001.rotx = 0
         laneDelay = 0
         leftlanekicker.timerenabled = 0
  End Select
End Sub

Dim saucerDelay
Sub saucer5Kick_timer
  saucerDelay = saucerDelay + 1

  If saucerDelay = 2 Then
    kickSaucer5
    saucerDelay = 0
    saucer5Kick.enabled = 0
  End If
End Sub

Sub KickSaucer5
  dim rangle, BOT, b, saucerScore
  saucerKick1.enabled = 1
  BOT = GetBalls
  playFieldSound "SaucerKick",0, Saucer5, 1
  If B2SOn Then DOF 122, DOFPulse
    For b = 0 to uBound(BOT)
      If BOT(b).x > 465 and BOT(b).x < 485 and BOT(b).y > 718 and BOT(b).y < 738 Then
        rangle = 0
        BOT(b).z = BOT(b).z + 30
        BOT(b).velz = 8
        BOT(b).vely = 6
        BOT(b).velX = -1
      End If
    Next
End Sub

Dim saucerCount
Sub cycleSaucers_timer
  saucerCount = saucerCount + 1
  saucerKick.enabled = 1
  kickSaucerLaneBall
  playFieldSound "SaucerKick",0, Saucer3, 1
  If B2SOn Then DOF 123, DOFPulse
  If saucerCount > 4 Then
    saucerCount = 0
    ballsLifted = 0
    cycleSaucers.enabled = 0

  End If
End Sub


Sub kickSaucerLaneBall()
  dim rangle, BOT, b, saucerScore
  BOT = GetBalls
    For b = 0 to uBound(BOT)
      If BOT(b).x > 796 and BOT(b).x < 816 and BOT(b).y > 1073 and BOT(b).y < 1093 Then
        BOT(b).z = BOT(b).z + 30
        BOT(b).velz = 8
        BOT(b).vely = 4
      End If
      If BOT(b).x > 796 and BOT(b).x < 816 and BOT(b).y > 1208 and BOT(b).y < 1228 Then
        BOT(b).z = BOT(b).z + 30
        BOT(b).velz = 8
        BOT(b).vely = 4
      End If
      If BOT(b).x > 796 and BOT(b).x < 816 and BOT(b).y > 1344 and BOT(b).y < 1364 Then
        BOT(b).z = BOT(b).z + 30
        BOT(b).velz = 8
        BOT(b).vely = 4
      End If
    If BOT(b).x > 796 and BOT(b).x < 816 and BOT(b).y > 1483 and BOT(b).y < 1503 Then
        BOT(b).z = BOT(b).z + 30
        BOT(b).velz = 8
        BOT(b).vely = 4
      End If
    Next
End Sub

dim kickerActive
Sub kicker_hit
  kickerActive = 1
End Sub

dim kickStep1, kickStep2

Sub saucerKick_timer
  Select Case kickStep1
        Case 3:pKickArm1.rotz=15: pKickArm2.rotz=15: pKickArm3.rotz=15: pKickArm4.rotz=15
        Case 4:pKickArm1.rotz=15: pKickArm2.rotz=15: pKickArm3.rotz=15: pKickArm4.rotz=15
        Case 5:pKickArm1.rotz=15: pKickArm2.rotz=15: pKickArm3.rotz=15: pKickArm4.rotz=15
        Case 6:pKickArm1.rotz=8: pKickArm2.rotz=8: pKickArm3.rotz=8: pKickArm4.rotz=8
        Case 7:pKickArm1.rotz=3: pKickArm2.rotz=3: pKickArm3.rotz=3: pKickArm4.rotz=3
    Case 8:pKickArm1.rotz=0: pKickArm2.rotz=0: pKickArm3.rotz=0: pKickArm4.rotz=0: kickStep1 = 0: saucerKick.Enabled = 0
    End Select
   kickstep1 = kickstep1 + 1
End Sub

Sub saucerKick1_timer
  Select Case kickStep2
        Case 3:pKickArm5.rotz=15
        Case 4:pKickArm5.rotz=15
        Case 5:pKickArm5.rotz=15
        Case 6:pKickArm5.rotz=8
        Case 7:pKickArm5.rotz=3
    Case 8:pKickArm5.rotz=0: kickStep1 = 0: saucerKick1.Enabled = 0
    End Select
   kickstep2 = kickstep2 + 1
End Sub

''**********Sling Shot Animations
'' Rstep and Lstep  are the variables that increment the animation
''****************
Dim lStep, rStep
Sub rightSlingShot_Slingshot
  ArrowUnit
  playfieldSound SoundFX("SlingShot",DOFContactors), 0, SoundPoint13, 1
  If B2SOn Then DOF 104, DOFPulse
'    addScore(10)
    rSling.Visible = 0
    rSling2.Visible = 1
    sling003.Rotx = 16
    rStep = 0
    rightSlingShot.TimerEnabled = 1
End Sub

Sub rightSlingShot_Timer
    Select Case rStep
        Case 3:rSLing2.Visible = 0:rSLing3.Visible = 1:sling003.Rotx = 0
        Case 4:rSLing3.Visible = 0:rSLing.Visible = 1:rightSlingShot.TimerEnabled = 0
    End Select
    rStep = rStep + 1
End Sub

'***************Scoring Routine
Dim flag10k, flag100k, point
Sub pts1
  playFieldSound "Chime1", 0, soundPoint13, 1
End Sub

Sub pts10
  playFieldSound "Chime10", 0, soundPoint13, 1
End Sub

Sub pts100
  playFieldSound "Chime100", 0, soundPoint13, 1
End Sub

Sub pts1000
  playFieldSound "Chime1000", 0, soundPoint13, 1
End Sub

Sub addScore(points)
  If tilt = False Then
    If points <99999 Then
      bellRing = (points/10000)
      If bellRing = 1 Then
        score(1) = score(1) + 10000
        scoreLightCalc
        If chime = 0 Then
          pts10
        Else
          If B2SOn Then DOF 155, DOFPulse
        End If
        checkReplay
      End If
      If bellRing > 1 Then point = 10000: flag10k = 1: scoreMotorLoop = 0: scoreMotor.enabled = 1
      Exit Sub
    End If

    If points > 99999 Then
      bellRing = (points / 100000)
      If bellRing = 1 Then
        score(1) = score(1) + 100000
        scoreLightCalc
        If chime = 0 Then
          pts100
        Else
          If B2SOn Then DOF 155,DOFPulse
        End If
        checkReplay
      End If
      If bellRing > 1 Then point = 100000: flag100k = 1: scoreMotorLoop = 0: scoreMotor.enabled = 1
      Exit Sub
    End If
  End If
End Sub

Sub ScoreLightCalc
  tenK = (score(1) \ 10000) mod 10
  hundredK = (score(1) \ 100000) mod 10
  million = (score(1) \ 1000000) mod 10
  reel10k.setValue(tenK)
  reel100k.setValue(hundredK)
  milReel.setValue(million)
  If vrOption > 0 Then
    vr10k.image = "vr10K" & tenK
    vr100k.image = "vr100K" & hundredK
    vrMillion.image = "vrMillion" & million
  End If
  If B2SOn Then
    resetBackglass
    controller.B2SSetData (60 + tenK), 1
    controller.B2SSetData (70 + hundredK), 1
    controller.B2SSetData (80 + million), 1
  End If
  TB5.text = score(1)
End Sub

Sub resetBackglass
  For x = 1 to 9
    reel10k.setValue(0)
    reel100k.setValue(0)
    milReel.setValue(0)
    If B2SOn Then
      controller.B2SSetData (60 + x), 0
      controller.B2SSetData (70 + x), 0
      controller.B2SSetData (80 + x), 0
    End If
  Next
End Sub

Dim  tenK, hundredK, million
Sub totalUp(points)
  If flag10k = 1 Then
    If chime = 0 Then
      pts10
    Else
      If B2SOn Then DOF 154, DOFPulse
    End If
  End If

  If flag100k = 1 Then
    If chime = 0 Then
      pts100
    Else
      If B2SOn Then DOF 155, DOFPulse
    End If
  End If

  score(1) = score(1) + points
  scoreLightCalc
  checkReplay
End Sub

Sub checkReplay
  For replayX = rep(1) + 1 to 5
    If score(1) => replay(replayX) Then
      ScoreSpecial
      rep(1) = rep(1) + 1
    End If
  Next
End Sub

'****************Score Motor Run Timer
Dim bellRing, scoreMotorLoop, kickDelay, scoreLoop
Sub scoreMotor_timer
  scoreMotorLoop = scoreMotorLoop + 1
  If scoreMotorLoop < 6 Then playFieldSound "ScoreMotorSingleFire", 0, SoundPointScoreMotor, .3

'  These Flags are passed by scores with multiple of 10, 100 or 1000
  If  flag10k = 1 or flag100k = 1 Then
    Select Case scoreMotorLoop
      Case 1: totalUp point
      Case 2: totalUp point
      Case 3: If bellRing > 2 Then totalUp point
      Case 4: If bellRing > 3 Then totalUp point
      Case 5: If bellRing > 4 Then totalUp point
      Case 6: scoreMotorLoop = 0
          flag10k = 0
          flag100k = 0
          scoremotor.enabled = 0
    End Select
  End If

  If payOut = 1 Then
    Select Case scoreMotorLoop
      Case 1: If ASB > 0 Then checkASB
      Case 2: If ASB > 0 Then checkASB
      Case 3: If ASB > 0 Then checkASB
      Case 4: If ASB > 0 Then checkASB
      Case 5: If ASB > 0 Then checkASB
      Case 6: scoreMotorLoop = 0
          If ASB < 1 Then scoremotor.enabled = 0: payOut = 0
    End Select
  End If
End Sub

Sub checkASB
  If ASB < 1 Then Exit Sub
  addCredit
  If ASB < 11 Then
    EVAL ("LightAdvance" & ASB).state = 0
    ASB = ASB - 1
    If ASB > 0 Then EVAL ("LightAdvance" & ASB).state = 1
  End If
  If ASB > 10 And ASB < 20 Then
    EVAL ("LightAdvance" & ASB - 10).state = 0
    ASB = ASB - 1
    If ASB > 10 Then EVAL("LightAdvance" & ASB - 10).state = 1
  End If
  If ASB = 20 Then
    LightAdvance11.state = 0
    LightAdvance10.state = 1
    LightAdvance9.state = 1
    ASB = ASB - 1
  End If
End Sub

Sub scoreSpecial
  For x = 1 to (rep(1) + 1)
    addCredit
    playsound SoundFXDOF("Knocker",128,DOFPulse,DOFKnocker)
  Next
End Sub

Sub special
  addCredit
  playsound SoundFXDOF("Knocker",128,DOFPulse,DOFKnocker)
End Sub

Sub laneSpecial
  addCredit
  laneSpecialCheck = 1
  playsound SoundFXDOF("Knocker",128,DOFPulse,DOFKnocker)
End Sub


Dim replayX,  replay(7), repAwarded(5)  ', reelVol, reelLpan, reelRpan
'reelVol = 0.3    'Volume value for reels.  0.1 a good value
'reelLpan = -0.12 'Pan value for left reels.  -0.12 is 3/4 pan left
'reelRpan = 0.12  'Pan value for right reels.  0.12 is 3/4 pan right

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
    DTReel.setValue(0)
        UV1.image = "UV1_off"
        rsling.image = "UV1_off"
    inserts.image = "UV1_off"
        pf_off.visible=1
        dim xx
        For each xx in GI:xx.State = 0: Next
        If B2SOn Then
      controller.B2SSetTilt 33,1
      controller.B2SSetdata 1, 0
      controller.B2SSetData 80, 0
    End If
    turnOff
    endGame = 1
    checkContinue
   End If
  Else
   tiltSens = 0
   tilttimer.enabled = True
  End If
End Sub

Sub tilttimer_Timer()
  tilttimer.enabled = False
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
  TB6.text = score(1) & " " & highScore(activeScore(flag))
  scoreUnit = (highScore(activeScore(flag)) \ 1) mod 10
  score10 = (highScore(activeScore(flag)) \ 10) mod 10
  score100 = (highScore(activeScore(flag)) \ 100) mod 10
  scoreK = (highScore(activeScore(flag)) \ 1000) mod 10
  score10K = (highScore(activeScore(flag)) \ 10000) mod 10
  score100K = (highScore(activeScore(flag)) \ 100000) mod 10
  scoreMil = (highScore(activeScore(flag)) \ 1000000) mod 10

' scoreMil = Int(highScore(activeScore(flag))/1000000)
' score100K = Int( (highScore(activeScore(flag)) - (scoreMil*1000000) ) / 100000)
' score10K = Int( (highScore(activeScore(flag)) - (scoreMil*1000000) - (score100K*100000) ) / 10000)
' scoreK = Int( (highScore(activeScore(flag)) - (scoreMil*1000000) - (score100K*100000) - (score10K*10000) ) / 1000)
' score100 = Int( (highScore(activeScore(flag)) - (scoreMil*1000000) - (score100K*100000) - (score10K*10000) - (scoreK*1000) ) / 100)
' score10 = Int( (highScore(activeScore(flag)) - (scoreMil*1000000) - (score100K*100000) - (score10K*10000) - (scoreK*1000) - (score100*100) ) / 10)
' scoreUnit = (highScore(activeScore(flag)) - (scoreMil*1000000) - (score100K*100000) - (score10K*10000) - (scoreK*1000) - (score100*100) - (score10*10) )

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
  If highScore(activeScore(flag)) > 999999 Then shift = 0 :pComma.transx = 0
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
  position(1) = 1
  checkHighScores
' For bSx = 1 to 4
'   position(bSx) = bSx
' Next
' For bSx = 1 to 4
'   For bSy = 1 to 1
'     If score(bSy) < score(bSy+1) Then
'       tempScore(1) = score(bSy+1)
'       tempPos(1) = position(bSy+1)
'       score(bSy+1) = score(bSy)
'       score(bSy) = tempScore(1)
'       position(bSy+1) = position(BSy)
'       position(bSy) = tempPos(1)
'     End If
'   Next
' Next
End Sub

'*************Check for High Scores
Dim highScore(5), activeScore(5), hs, chX, chY, chZ, chIX, tempI(4), tempI2(4), flag, hsI, hsX
'goes through the 5 high scores one at a time and compares them to the player's scores high to low
'if a player's score is higher, it marks that postion with ActiveScore(x) and moves all of the other
'high scores down by one along with the high score's player initials
'it also clears the new high score's initials for entry later
Sub checkHighScores
  tb1.text = score(1)
' For hs = 1 to 1                   'look at 4 player scores
    For chY = 0 to 4                  'look at all 5 saved high scores
      If score(1) > highScore(chY) Then
        flag = 1            'flag to show how many high scores needs replacing
        tempScore(1) = highScore(chY)
        highScore(chY) = score(1)
        activeScore(1) = chY        'ActiveScore(x) is the high score being modified with x=0 the largest and x=4 the smallest
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
' Next
' Goto Initial Entry
    hsI = 1     'go to the first initial for entry
    hsX = 1     'make the displayed inital be "A"
    If flag > 0 Then  'Flag 0 when all scores are updated so leave subroutine and reset variables
      showScore
      playerEntry.visible = 1
      playerEntry.image = "Player1"
'     TextBox3.text = ActiveScore(Flag) 'tells which high score is being entered
      TB2.text = "flag " & Flag
'     TextBox1.text =  Position(Flag) 'tells which player is entering values
      initial(activeScore(flag),1) = 1  'make first inital "A"
      For chY = 2 to 3
        initial(activeScore(flag),chY) = 0  'set other two to " "
      Next
      For chY = 1 to 3
        EVAL("Initial" & chY).image = hsIArray(initial(activeScore(flag),chY))    'display the initals on the tape
      Next
      initialTimer1.enabled = 1   'flash the first initial
      dynamicUpdatePostIt.enabled = 0   'stop the scrolling intials timer
      playsound SoundFXDOF("Knocker",128,DOFPulse,DOFKnocker)
      enableInitialEntry = True
      Exit Sub
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
        TB.text = initial(activeScore(flag),2) & " " & initial(activeScore(flag),3)
        initialEntry              'exit subroutine
      End If

    End If
End Sub

'************Update Initials and see if more scores need to be updated
Dim eIX
Sub initialEntry
  pts10
  flag = flag - 1
' TextBox2.text = Flag
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
'   TextBox3.text = ActiveScore(Flag)   'tells which high score is being entered
'   TextBox2.text = Flag
'   TextBox1.text =  Position(Flag)   'tells which player is entering values
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

'**************************************************File Writing Section******************************************************
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

'**************************************************File Writing Section******************************************************

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
    vrOption = CInt(Right(temp(33),1))
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
    scoreFile.WriteLine "vrOption (0 = DT/Cab, 1 = VR Min, 2 = VR Room): " & vrOption
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

  If cPinballY = 1 Then Set ScoreFile = FileObj.CreateTextFile(UserDirectory & hsFileName & ".PinballYHighScores",True)


  For x = 0 to 4
    ScoreFile.WriteLine HighScore(x)
    ScoreFile.WriteLine HiInit(x)
  Next
  ScoreFile.Close
  Set ScoreFile = Nothing
  Set FileObj = Nothing

End Sub

'************VR Subs
Sub vrCredit
  vrCreditReel.objRotX = credit * 22.5
End Sub

'*******************************************
' VR Room / VR Cabinet
'*******************************************

DIM VRThings
If vrOption = 1 Then
  for each VRThings in vrCab: VRThings.visible = 1: Next
  for each VRThings in vrRoom: VRThings.visible = 0: Next
  For each object in backdropstuff: Object.visible = 0: Next

ElseIf vrOption = 2 Then
  for each VRThings in vrCab: VRThings.visible = 1: Next
  for each VRThings in vrRoom: VRThings.visible = 1: Next
  For each object in backdropstuff: Object.visible = 0: Next

Else
  for each VRThings in vrCab: VRThings.visible = 0: Next
  for each VRThings in vrRoom: VRThings.visible = 0: Next
  tape.image = "tapeCab"
End if

Sub turnOff
  For i = 1 to 2
    EVAL("Bumper" & i).hasHitEvent = 0
  Next
  If tiltPenalty = 1 then ballsToPlay = 0
    leftFlipper.RotateToStart
    stopSound "FlipBuzzLA"
    stopSound "FlipBuzzLB"
    stopSound "FlipBuzzLC"
    stopSound "FlipBuzzLD"
    stopSound "FlipBuzzRA"
    stopSound "FlipBuzzRB"
    stopSound "FlipBuzzRC"
    stopSound "FlipBuzzRD"
  DOF 101, DOFOff
  DOF 111, DOFOff
  DOF 112, DOFOff
' BonusScore = 0
End Sub

'*****************************************************Supporting Code Written By Others*************************************
'*********************************************
'  VR Plunger Code From Flash by Bord and Roth
'*********************************************

Sub TimerVRPlunger_Timer
  If VR_cab_plunger.Y < -28 then
       VR_cab_plunger.Y = VR_cab_plunger.Y + 5
  End If
End Sub

Sub TimerVRPlunger1_Timer
  VR_cab_plunger.Y = -128 + (5 * Plunger.Position) -20
End Sub

Sub TimerVRBallLift_Timer
  If VR_cab_balllift.Y > -20 then
       VR_cab_balllift.Y = VR_cab_balllift.Y - 5
  End If
End Sub


'************************************************************************
'                         Ball Control
'************************************************************************

Dim Cup, Cdown, Cleft, Cright, Zup, contball, contballinplay, ControlBall, bcboost
Dim bcvel, bcyveloffset, bcboostmulti

bcboost = 1 'Do Not Change - default setting
bcvel = 4 'Controls the speed of the ball movement
bcyveloffset = 0 'Offsets the force of gravity to keep the ball from drifting vertically on the table, should be negative
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

Sub LightsRandom_Timer()
  Select Case Int(Rnd*2)+1
    Case 1 : DOF 157, 1
    Case 2 : DOF 157, 0
  End Select
  Select Case Int(Rnd*2)+1
    Case 1 : DOF 158, 1
    Case 2 : DOF 158, 0
  End Select
  Select Case Int(Rnd*2)+1
    Case 1 : DOF 159, 1
    Case 2 : DOF 159, 0
  End Select
  Select Case Int(Rnd*2)+1
    Case 1 : DOF 160, 1
    Case 2 : DOF 160, 0
  End Select
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
  Dim tmp
    If PFOption = 1 Then tmp = TableObj.x * 2 / table1.width-1
  If PFOption = 2 Then tmp = TableObj.y * 2 / table1.height-1
  If tmp < 0 Then
  xGain = 0.3293074856*EXP(-0.9652695455*tmp^3 - 2.452909811*tmp^2 - 2.597701999*tmp)
  Else
  xGain = 0.3293074856*EXP(-0.9652695455*-tmp^3 - 2.452909811*-tmp^2 - 2.597701999*-tmp)
  End If
End Function

Function XVol(tableobj)
'XVol algorithm calculates a PlaySound Volume parameter multiplier for a tableobj based on its X table position to provide a Constant Power "pan".
'XVol = 1 at PF Left, XVol = 0.32931 (-3dB for PlaySound's volume parameter) at PF Center and XVol = 0 at PF Right
Dim tmpx
  If PFOption = 3 Then
    tmpx = tableobj.x * 2 / table1.width-1
    XVol = 0.3293074856*EXP(-0.9652695455*tmpx^3 - 2.452909811*tmpx^2 - 2.597701999*tmpx)
  End If
End Function

Function YVol(tableobj)
'YVol algorithm calculates a PlaySound Volume parameter multiplier for a tableobj based on its Y table position to provide a Constant Power "fade".
'YVol = 1 at PF Top, YVol = 0.32931 (-3dB for PlaySound's volume parameter) at PF Center and YVol = 0 at PF Bottom
Dim tmpy
  If PFOption = 3 Then
    tmpy = tableobj.y * 2 / table1.height-1
    YVol = 0.3293074856*EXP(-0.9652695455*tmpy^3 - 2.452909811*tmpy^2 - 2.597701999*tmpy)
  End If
End Function

'*********************************************************************************

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
'      JP's VP10 Rolling Sounds - Modified by Whirlwind
'*****************************************

'******************************************
' Explanation of the rolling sound routine
'******************************************

' ball rolling sounds are played based on the ball speed and position
' the routine checks first for deleted balls and stops the rolling sound.
' The For loop goes through all the balls on the table and checks for the ball speed and
' if the ball is on the table (height lower than 30) then then it plays the sound
' otherwise the sound is stopped.

'New algorithms added to make sounds for TopArch Hits, TopArch Rolls, ball bounces and glass hits.
'For stereo, xGain is a Playsound volume multiplier that provides a Constant Power pan.
'For quad, multiple PlaySound commands are launched together that are panned and faded to their maximum extents where PlaySound's PAN and FADE have the least error.
'XVol and YVol are Playsound volume multipliers that provide a Constant Power "pan" and "fade".
'Subtracting XVol or YVol from 1 yeilds an inverse response.

Const tnob = 4 ' total number of balls

'Change GHT, GHB and PFL values based upon the real pinball table dimensions.  Values are used by the GlassHit code.
'To ensure ball can go high enough to trigger glass hit, make Table Options/Dimensions & Slope/Top Glass Height equal to (GHT*50/1.0625) + 5
Const GHT = 3 'Glass height in inches at top of real playfield
Const GHB = 3 'Glass height in inches at bottom of real playfield
Const PFL = 40  'Real playfield length in inches

ReDim rolling(tnob)
InitRolling

Sub InitRolling
  Dim i
  For i = 0 to tnob
    rolling(i) = False
  Next
End Sub

Sub RollingTimer_Timer()
  Dim BOT, b, pa
  BOT = GetBalls
  pa=35000  'Playsound pitch adder for subway and wire rolling ball sound

'TextBox1.text="BOT(b).Z  " & formatnumber(BOT(b).Z,1)
'TextBox2.text="BOT(b).VelZ  " & formatnumber(BOT(b).VelZ,1)
'TextBox3.text="GLASS  " & formatnumber((BOT(b).Y * ((GHT-GHB)/PFL)) + (GHB*50/1.0625) - (BallSize/2),1)

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

''  Ball Rolling sounds
''**********************
' Ball=50 units=1.0625".  One unit = 0.02125"  Ball.z is ball center.

  If PFOption = 1 Or PFOption = 2 Then
    If BallVel(BOT(b) ) > 1 And BOT(b).z < 26 And BOT(b).z > 0 Then
      rolling(b) = True
      PlaySound("BallrollingA" & b), -1, Vol(BOT(b)) * 0.2 * xGain(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, 0 'Left & Right stereo or Top & Bottom stereo PF Speakers.
    ElseIf BallVel(BOT(b) ) > 1 And BOT(b).z > 30 Then
      PlaySound("BallrollingA" & b), -1, Vol(BOT(b)) * 0.2 * xGain(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b))+pa, 1, 0, 0  'Left & Right stereo or Top & Bottom stereo PF Speakers.
    ElseIf rolling(b) = True Then
      StopSound("BallrollingA" & b)
      rolling(b) = False
    End If
  End If

  If PFOption = 3 Then
    If BallVel(BOT(b) ) > 1 And BOT(b).z < 26 And BOT(b).z > 0 Then
      rolling(b) = True
      PlaySound("BallrollingA" & b), -1, Vol(BOT(b)) * 0.2 *    XVol(BOT(b))  *     YVol(BOT(b)), -1, 0, Pitch(BOT(b)), 1, 0, -1  'Top Left PF Speaker
      PlaySound("BallrollingB" & b), -1, Vol(BOT(b)) * 0.2 * (1-XVol(BOT(b))) *     YVol(BOT(b)),  1, 0, Pitch(BOT(b)), 1, 0, -1  'Top Right PF Speaker
      PlaySound("BallrollingC" & b), -1, Vol(BOT(b)) * 0.2 *    XVol(BOT(b))  * (1-YVol(BOT(b))), -1, 0, Pitch(BOT(b)), 1, 0,  1  'Bottom Left PF Speaker
      PlaySound("BallrollingD" & b), -1, Vol(BOT(b)) * 0.2 * (1-XVol(BOT(b))) * (1-YVol(BOT(b))),  1, 0, Pitch(BOT(b)), 1, 0,  1  'Bottom Right PF Speaker
    ElseIf BallVel(BOT(b) ) > 1 And BOT(b).z > 30 Then
      PlaySound("BallrollingA" & b), -1, Vol(BOT(b)) * 0.2 *    XVol(BOT(b))  *     YVol(BOT(b)), -1, 0, Pitch(BOT(b))+pa, 1, 0, -1 'Top Left PF Speaker
      PlaySound("BallrollingB" & b), -1, Vol(BOT(b)) * 0.2 * (1-XVol(BOT(b))) *     YVol(BOT(b)),  1, 0, Pitch(BOT(b))+pa, 1, 0, -1 'Top Right PF Speaker
      PlaySound("BallrollingC" & b), -1, Vol(BOT(b)) * 0.2 *    XVol(BOT(b))  * (1-YVol(BOT(b))), -1, 0, Pitch(BOT(b))+pa, 1, 0,  1 'Bottom Left PF Speaker
      PlaySound("BallrollingD" & b), -1, Vol(BOT(b)) * 0.2 * (1-XVol(BOT(b))) * (1-YVol(BOT(b))),  1, 0, Pitch(BOT(b))+pa, 1, 0,  1 'Bottom Right PF Speaker
    ElseIf rolling(b) = True Then
      StopSound("BallrollingA" & b)   'Top Left PF Speaker
      StopSound("BallrollingB" & b)   'Top Right PF Speaker
      StopSound("BallrollingC" & b)   'Bottom Left PF Speaker
      StopSound("BallrollingD" & b)   'Bottom Right PF Speaker
      rolling(b) = False
    End If
  End If

' Ball drop sounds
'*******************
'Four intensities of ball bounce sound files ranging from 1 to 4 bounces.  The number of bounces increases as the ball's downward Z velocity increases.
'A BOT(b).VelZ < -2 eliminates nuisance ball bounce sounds.

  If PFOption = 1 or PFOption = 2 Then
    If BOT(b).VelZ > -4 And BOT(b).VelZ < -2 And BOT(b).Z > 24 And ballsToPlay => 1 Then
      PlaySound "BallDrop1" & b, 0, ABS(BOT(b).VelZ)/600 * xGain(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 0, 0, 0 'Left & Right stereo or Top & Bottom stereo PF Speakers.
    ElseIf BOT(b).VelZ > -8 And BOT(b).VelZ < -4 And BOT(b).Z > 24 And ballsToPlay => 1 Then
      PlaySound "BallDrop2" & b, 0, ABS(BOT(b).VelZ)/600 * xGain(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 0, 0, 0 'Left & Right stereo or Top & Bottom stereo PF Speakers.
    ElseIf BOT(b).VelZ > -12 And BOT(b).VelZ < -8 And BOT(b).Z > 24 And ballsToPlay => 1 Then
      PlaySound "BallDrop3" & b, 0, ABS(BOT(b).VelZ)/600 * xGain(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 0, 0, 0 'Left & Right stereo or Top & Bottom stereo PF Speakers.
    ElseIf BOT(b).VelZ < -12 And BOT(b).Z > 24 And ballsToPlay => 1 Then
      PlaySound "BallDrop4" & b, 0, ABS(BOT(b).VelZ)/600 * xGain(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 0, 0, 0 'Left & Right stereo or Top & Bottom stereo PF Speakers.
    End If
  End If

  If PFOption = 3 Then
    If BOT(b).VelZ > -4 And BOT(b).VelZ < -2 And BOT(b).Z > 24 And ballsToPlay => 1 Then
      PlaySound "BallDrop1" & b, 0, ABS(BOT(b).VelZ)/600 *    XVol(BOT(b))  *     YVol(BOT(b)), -1, 0, Pitch(BOT(b)), 1, 0, -1  'Top Left PF Speaker
      PlaySound "BallDrop1" & b, 0, ABS(BOT(b).VelZ)/600 * (1-XVol(BOT(b))) *     YVol(BOT(b)),  1, 0, Pitch(BOT(b)), 1, 0, -1  'Top Right PF Speaker
      PlaySound "BallDrop1" & b, 0, ABS(BOT(b).VelZ)/600 *    XVol(BOT(b))  * (1-YVol(BOT(b))), -1, 0, Pitch(BOT(b)), 1, 0,  1  'Bottom Left PF Speaker
      PlaySound "BallDrop1" & b, 0, ABS(BOT(b).VelZ)/600 * (1-XVol(BOT(b))) * (1-YVol(BOT(b))),  1, 0, Pitch(BOT(b)), 1, 0,  1  'Bottom Right PF Speaker
    ElseIf BOT(b).VelZ > -8 And BOT(b).VelZ < -4 And BOT(b).Z > 24 And ballsToPlay => 1 Then
      PlaySound "BallDrop2" & b, 0, ABS(BOT(b).VelZ)/600 *    XVol(BOT(b))  *     YVol(BOT(b)), -1, 0, Pitch(BOT(b)), 1, 0, -1  'Top Left PF Speaker
      PlaySound "BallDrop2" & b, 0, ABS(BOT(b).VelZ)/600 * (1-XVol(BOT(b))) *     YVol(BOT(b)),  1, 0, Pitch(BOT(b)), 1, 0, -1  'Top Right PF Speaker
      PlaySound "BallDrop2" & b, 0, ABS(BOT(b).VelZ)/600 *    XVol(BOT(b))  * (1-YVol(BOT(b))), -1, 0, Pitch(BOT(b)), 1, 0,  1  'Bottom Left PF Speaker
      PlaySound "BallDrop2" & b, 0, ABS(BOT(b).VelZ)/600 * (1-XVol(BOT(b))) * (1-YVol(BOT(b))),  1, 0, Pitch(BOT(b)), 1, 0,  1  'Bottom Right PF Speaker
    ElseIf BOT(b).VelZ > -12 And BOT(b).VelZ < -8 And BOT(b).Z > 24 And ballsToPlay => 1 Then
      PlaySound "BallDrop3" & b, 0, ABS(BOT(b).VelZ)/600 *    XVol(BOT(b))  *     YVol(BOT(b)), -1, 0, Pitch(BOT(b)), 1, 0, -1  'Top Left PF Speaker
      PlaySound "BallDrop3" & b, 0, ABS(BOT(b).VelZ)/600 * (1-XVol(BOT(b))) *     YVol(BOT(b)),  1, 0, Pitch(BOT(b)), 1, 0, -1  'Top Right PF Speaker
      PlaySound "BallDrop3" & b, 0, ABS(BOT(b).VelZ)/600 *    XVol(BOT(b))  * (1-YVol(BOT(b))), -1, 0, Pitch(BOT(b)), 1, 0,  1  'Bottom Left PF Speaker
      PlaySound "BallDrop3" & b, 0, ABS(BOT(b).VelZ)/600 * (1-XVol(BOT(b))) * (1-YVol(BOT(b))),  1, 0, Pitch(BOT(b)), 1, 0,  1  'Bottom Right PF Speaker
    ElseIf BOT(b).VelZ < -12 And BOT(b).Z > 24 And ballsToPlay => 1 Then
      PlaySound "BallDrop4" & b, 0, ABS(BOT(b).VelZ)/600 *    XVol(BOT(b))  *     YVol(BOT(b)), -1, 0, Pitch(BOT(b)), 1, 0, -1  'Top Left PF Speaker
      PlaySound "BallDrop4" & b, 0, ABS(BOT(b).VelZ)/600 * (1-XVol(BOT(b))) *     YVol(BOT(b)),  1, 0, Pitch(BOT(b)), 1, 0, -1  'Top Right PF Speaker
      PlaySound "BallDrop4" & b, 0, ABS(BOT(b).VelZ)/600 *    XVol(BOT(b))  * (1-YVol(BOT(b))), -1, 0, Pitch(BOT(b)), 1, 0,  1  'Bottom Left PF Speaker
      PlaySound "BallDrop4" & b, 0, ABS(BOT(b).VelZ)/600 * (1-XVol(BOT(b))) * (1-YVol(BOT(b))),  1, 0, Pitch(BOT(b)), 1, 0,  1  'Bottom Right PF Speaker
    End If
  End If

' Glass hit sounds
'*******************
' Ball=50 units=1.0625".  Ball.z is ball center.  Balls are physically limited by Top Glass Height.  Max ball.z is 25 units below Top Glass Height.
' To ensure ball can go high enough to trigger glass hit, make Table Options/Dimensions & Slope/Top Glass Height equal to (GHT*50/1.0625) + 5

  If PFOption = 1 or PFOption = 2 Then
    If BOT(b).Z > (BOT(b).Y * ((GHT-GHB)/PFL)) + (GHB*50/1.0625) - BallSize/2 And ballsToPlay => 1 Then
      PlaySound "GlassHit" & b, 0, ABS(BOT(b).VelZ)/30 * xGain(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 0, 0, 0 'Left & Right stereo or Top & Bottom stereo PF Speakers.
    End If
  End If

  If PFOption = 3 Then
    If BOT(b).Z > (BOT(b).Y * ((GHT-GHB)/PFL)) + (GHB*50/1.0625) - Ballsize/2 And ballsToPlay => 1 Then
      PlaySound "GlassHit" & b, 0, ABS(BOT(b).VelZ)/30 *    XVol(BOT(b))  *     YVol(BOT(b)), -1, 0, Pitch(BOT(b)), 0, 0, -1  'Top Left PF Speaker
      PlaySound "GlassHit" & b, 0, ABS(BOT(b).VelZ)/30 * (1-XVol(BOT(b))) *     YVol(BOT(b)),  1, 0, Pitch(BOT(b)), 0, 0, -1  'Top Right PF Speaker
      PlaySound "GlassHit" & b, 0, ABS(BOT(b).VelZ)/30 *    XVol(BOT(b))  * (1-YVol(BOT(b))), -1, 0, Pitch(BOT(b)), 0, 0,  1  'Bottom Left PF Speaker
      PlaySound "GlassHit" & b, 0, ABS(BOT(b).VelZ)/30 * (1-XVol(BOT(b))) * (1-YVol(BOT(b))),  1, 0, Pitch(BOT(b)), 0, 0,  1  'Bottom Right PF Speaker
    End If
  End If
  Next
End Sub

'*************Hit Sound Routines
'Eliminated the Hit Subs extra velocity criteria since the PlayFieldSoundAB command already incorporates the ball’s velocity.

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
' If BallinPlay > 0 Then  'Eliminates the thump of Trough Ball Creation balls hitting walls 9 and 14 during table initiation
  PlayFieldSoundAB "rubberBand", 0, 1
' End If
End Sub

Sub aRubberWheel_hit(idx)
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

Sub ApronWalls_Hit  'Uses only Y velocity to capture ball vertical bounces.  ^3 gives faster volume decay of the ball bouncing off the apron repeatedly.
  Dim Volume
  If ActiveBall.vely < 0 Then Volume = abs(ActiveBall.vely) / 1 Else Volume = ActiveBall.vely / 30  'The first bounce is -vely subsequent bounces are +vely
  Volume = ABS(ActiveBall.vely ^3) / 50
  If ActiveBall.z > 24 And ActiveBall.x > 150 Then
    If PFOption = 1 Or PFOption = 2 Then
      PlaySound "woodHit", 0, Volume * xGain(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, 0 'Left & Right stereo or Top & Bottom stereo PF Speakers.
    End If
    If PFOption = 3 Then
      PlaySound "woodHit", 0, Volume *    XVol(ActiveBall)  *    YVol(ActiveBall),  -1, 0, Pitch(ActiveBall), 0, 0, -1  'Top Left PF Speaker
      PlaySound "woodHit", 0, Volume * (1-XVol(ActiveBall)) *    YVol(ActiveBall),   1, 0, Pitch(ActiveBall), 0, 0, -1  'Top Right PF Speaker
      PlaySound "woodHit", 0, Volume *    XVol(ActiveBall)  * (1-YVol(ActiveBall)), -1, 0, Pitch(ActiveBall), 0, 0,  1  'Bottom Left PF Speaker
      PlaySound "woodHit", 0, Volume * (1-XVol(ActiveBall)) * (1-YVol(ActiveBall)),  1, 0, Pitch(ActiveBall), 0, 0,  1  'Bottom Right PF Speaker

    End If
  End If
'Tb2.text= "Volume=" & round(Volume,2) & "     Adjust Volume devisor to keep Volume around 0.6 when dropping balls on the apron using ball control"
End Sub

Sub Saucers_Hit(idx)
  If PFOption = 1 Or PFOption = 2 Then
    PlaySound "metalHitMedium", 0, 1 * xGain(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall)-11025, 0, 0, 0 'Left & Right stereo or Top & Bottom stereo PF Speakers.
  End If
  If PFOption = 3 Then
    PlaySound "metalHitMedium", 0, 1 *    XVol(ActiveBall)  *    YVol(ActiveBall),  -1, 0, Pitch(ActiveBall)-11025, 0, 0, -1  'Top Left PF Speaker
    PlaySound "metalHitMedium", 0, 1 * (1-XVol(ActiveBall)) *    YVol(ActiveBall),   1, 0, Pitch(ActiveBall)-11025, 0, 0, -1  'Top Right PF Speaker
    PlaySound "metalHitMedium", 0, 1 *    XVol(ActiveBall)  * (1-YVol(ActiveBall)), -1, 0, Pitch(ActiveBall)-11025, 0, 0,  1  'Bottom Left PF Speaker
    PlaySound "metalHitMedium", 0, 1 * (1-XVol(ActiveBall)) * (1-YVol(ActiveBall)),  1, 0, Pitch(ActiveBall)-11025, 0, 0,  1  'Bottom Right PF Speaker
  End If
End Sub

'**********************
' Ball Collision Sound
'*********************,

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
'Subtracting XVol or YVol from 1 yeilds an inverse response.

Sub OnBallBallCollision(ball1, ball2, velocity)
  If gameStart = 0 Then Exit Sub
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
'Subtracting XVol or YVol from 1 yeilds an inverse response.

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
'Subtracting XVol or YVol from 1 yeilds an inverse response.

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
    If Ucase(Pan) = "LPAN" Then PlaySound SoundName, 0, ReelVolAdj * 1.00, -0.12, 0, 0, 1, 1, 0 'Panned 3/4 Left at 0dB * ReelVolAdj
    If Ucase(Pan) = "MPAN" Then PlaySound SoundName, 0, ReelVolAdj * 0.33,  0.00, 0, 0, 1, 1, 0 'Panned Middle at -3dB * ReelVolAdj
    If Ucase(Pan) = "RPAN" Then PlaySound SoundName, 0, ReelVolAdj * 1.00,  0.12, 0, 0, 1, 1, 0 'Panned 3/4 Right at 0dB * ReelVolAdj
  Else
    If Ucase(Pan) = "LPAN" Then PlaySound SoundName, 0, ReelVolAdj * 0.33, -0.12, 0, 0, 1, 1, 0 'Panned 3/4 Left at -3dB * ReelVolAdj
    If Ucase(Pan) = "MPAN" Then PlaySound SoundName, 0, ReelVolAdj * 0.11,  0.00, 0, 0, 1, 1, 0 'Panned Middle at -6dB * ReelVolAdj
    If Ucase(Pan) = "RPAN" Then PlaySound SoundName, 0, ReelVolAdj * 0.33,  0.12, 0, 0, 1, 1, 0 'Panned 3/4 Right at -3dB * ReelVolAdj
  End If
End Sub

'********Need to have a flipper timer to check for these values

Sub flipperTimer_Timer
  Pgate002.rotx = -Gate002.CurrentAngle*0.5
' tb4.text = "ball draining = " & ballDraining
' tb.text = "BtP = " & ballsToPlay
' tb1.text = "ballsLifted = " & ballsLifted
' tb2.text = "BallsDestroyed = " & destroyedBalls
  lFlip.rotz = leftflipper.CurrentAngle -121 'silver metal flipper obj
  FlipperLShadow.RotZ = LeftFlipper.CurrentAngle
  If DynamicBallShadowsOn Or AmbientBallShadowOn Then DynamicBSUpdate 'update ball shadows
End Sub

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
'RubbersD.addpoint 0, 0, 0.935  'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 0, 0, 1.12 '0.97'0.935 '0.96  'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 1, 3.77, 0.999 '0.97 '0.935 '0.96
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

'   TB.text = coef

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


' *** These are usually defined elsewhere (ballrolling), but activate here if necessary
Const lob = 0 'locked balls on start; might need some fiddling depending on how your locked balls are done
dim gilvl:gilvl = 1

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

Dim PI: PI = 4*Atn(1)

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

Function AnglePP(ax,ay,bx,by)
  AnglePP = Atn2((by - ay),(bx - ax))*180/PI
End Function

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

Sub DynamicBSInit()
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
End Sub


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

Sub Table1_Exit
  If B2sOn Then controller.stop
  saveHighScore
End Sub

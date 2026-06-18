'****************************************************************
'
'            Nudgy (Bally 1947)
'             v1.05
'         Script Code by Scottacus
'             August 2022
'
' Basic DOF config
'   104 Right side ball advance
'   127 credit light
'   128 Knocker, 129 Knocker and Kicker Strobe
'   141 Chime/Bell
'
'   Code Flow
'
'   Start Game -> Check Continue ->  BallDrain -
'             ^           |
'           EndGame = False <-----------
'***********************************************************************************************************************

Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

Const cGameName = "Nudgy_1947"
Const cOptions = "Nudgy_1947_1.05.txt"
Const hsFileName = "Nudgy (Bally 1947)"

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
'Const VRGlassScratches = 0  ' set this to 1 to turn on VR glass scratches
'If vrOption > 0 and VRGlassScratches > 0 Then GlassImpurities.visible = True

'*************************** PinballY Settings *****************************************
'Set this variable to 1 to save a PinballY High Score file to your Tables Folder
'this will let the Pinball Y front end display the high scores when searching for tables
'0 = No PinballY High Scores, 1 = Save PinballY High Scores
Const cPinballY = 1

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
Dim maxPlayers
Dim players
Dim player
Dim credit
Dim score(6)
Dim hScore(6)
Dim state
Dim tilt
Dim matchNumber
Dim i,j, f, ii, Object, Light, x, y, z
Dim freePlay
Dim ballsize,BallMass
Dim hsInitial0, hsInitial1, hsInitial2
Dim hsArray: hsArray = Array("HS_0","HS_1","HS_2","HS_3","HS_4","HS_5","HS_6","HS_7","HS_8","HS_9","HS_Space","HS_Comma")
Dim hsiArray: hsIArray = Array("HSi_0","HSi_1","HSi_2","HSi_3","HSi_4","HSi_5","HSi_6","HSi_7","HSi_8","HSi_9","HSi_10","HSi_11","HSi_12","HSi_13","HSi_14","HSi_15","HSi_16","HSi_17","HSi_18","HSi_19","HSi_20","HSi_21","HSi_22","HSi_23","HSi_24","HSi_25","HSi_26")
ballSize = 50
ballMass = 1
Dim options
Dim replayEB
Dim chime
Dim onBumper
Dim pfOption
Dim lutValue

dim ballsout

Sub Table1_init
  LoadEM
  maxPlayers = 1
  player = 1
  loadHighScore
  If highScore(0)="" Then highScore(0)=200000
  If highScore(1)="" Then highScore(1)=150000
  If highScore(2)="" Then highScore(2)=120000
  If highScore(3)="" Then highScore(3)=110000
  If highScore(4)="" Then highScore(4)=95000
  If matchNumber = "" Then matchNumber = 4
  If showDT = True Then pfOption = 1
  If pfOption = "" Then pfOption = 1
  score(1) = 0
  If initial(0,1) = "" Then
    initial(0,1) = 2: initial(0,2) = 18: initial(0,3) = 4
    initial(1,1) = 1: initial(1,2) = 1: initial(1,3) = 1
    initial(2,1) = 2: initial(2,2) = 2: initial(2,3) = 2
    initial(3,1) = 3: initial(3,2) = 3: initial(3,3) = 3
    initial(4,1) = 4: initial(4,2) = 4: initial(4,3) = 4
  End If
  If credit = "" Then credit = 0
  If freePlay = "" Then freePlay = 1
  If balls = "" Then balls = 5
  If chime = "" Then chime = 0
  If vrOption = "" Then vrOption = 0
  If lutValue = "" Then lutValue = 0

  If FreePlay = 1 Then
    If B2SOn Then DOF 127, DOFOn
  End If

  move = 1

  If vrOption > 0 Then vrBackglass.image = "nudgyOff"

' This turns off collidability for all layers but the primary at table start up
  For x = 2 to 5
    For each obj in EVAL("NudgePF" & x)
      obj.collidable = 0
    Next
  Next

  contball = 0

  For x = 1 to balls
    EVAL("Kicker" & x).createball
    EVAL("Kicker" & x).kick 180, 3
    EVAL("Kicker" & x).enabled = False
  Next
  ballsout = 1
  firstBallOut = 0
  updatePostIt
  dynamicUpdatePostIt.enabled = 1


  ballShadowUpdate.enabled = True

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

  tilt = False
  state = False
  gameState

End Sub

'****************KeyCodes
Dim enableInitialEntry, firstBallOut
Sub Table1_KeyDown(ByVal keycode)

  If enableInitialEntry = True Then enterInitials(keycode)

  If keycode = addCreditKey Then
    playFieldSound "coinin",0,Drain,1
    Credit = 1
    End If

    If keycode = startGameKey Then
    If enableInitialEntry = False and operatormenu = 0 and backGlassOn = 1 Then
      If freePlay = 1 and state = 0 Then startGame
      If freePlay = 0 and credit > 0 and players = 0 and firstBallOut = 0 Then
        If B2SOn Then DOF 127, DOFOff
        startGame
      End If
    End If
  End If

  If keycode = startGameKey and state = false and enableInitialEntry = False and vrOption > 0 Then VR_cab_buttonstart.y = -5

  If Keycode = startGameKey and contBall = 0 and lifter = 0 and starter = 1 and state = True and enableInitialEntry = False Then
    ballLiftTimer.enabled = 1
    If vrOption > 0 Then
      VR_cab_balllift.Y = -30
    End If
  End If

  If Keycode = rightMagnaSave and contBall = 0 and lifter = 0 and state = True Then
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

  If keycode = PlungerKey Then
    plunger.PullBack
    playFieldSound "plungerpull", 0, plunger, 1
    If vrOption > 0 Then
      TimerVRPlunger.Enabled = True
      TimerVRPlunger1.Enabled = False
'     VR_cab_plunger.Y = -128
    End If
  End If

  If tilt = False and state = True Then
  If keycode = leftFlipperKey and contball = 0 Then
    nudgeUpTimer.enabled = 1
    nudgeDownTimer.enabled = 0
  End If

  If keycode = rightFlipperKey and contball = 0 Then
    nudgeUpTimer.enabled = 1
    nudgeDownTimer.enabled = 0
  End If

  End if

    If keycode = leftFlipperKey and state = False and operatorMenu = 0 and enableInitialEntry =  False Then
        operatorMenuTimer.Enabled = true
    end if

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
        If showDT = True Then optionMenu.visible = False
        OptionMenu1.visible = 1
        optionMenu1.image = "Sound" & pfOption
        optionMenu2.visible = 1
        optionMenu2.image = "SoundChange"
        Select Case (pfOption)
          Case 1: speaker1.visible = 1: speaker2.visible = 1: speaker3.visible = 0: speaker4.visible = 0
          Case 2: speaker5.visible = 1: speaker6.visible = 1: speaker1.visible = 0: speaker2.visible = 0
          Case 3: speaker1.visible = 1: speaker2.visible = 1: speaker3.visible = 1: speaker4.visible = 1: speaker5.visible = 0: speaker6.visible = 0
        End Select
      Case 2:
        OptionMenu.visible = 1
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

    If keycode = rightFlipperKey and state = False and operatorMenu = 1 Then
      playFieldSound "metalhit2", 0, Speaker5, 1.5
      Select Case (Options)
    Case 0:
            If freePlay = 0 Then
                freePlay = 1
              Else
                freePlay = 0
            End If
            OptionMenu.image= "FreeCoin" & freePlay
      If freePlay = 0 Then
        If credit > 0 and B2SOn Then DOF 127, DOFOn
        If credit < 1 and B2SOn Then DOF 127, DOFOff
      Else
        If B2SOn Then DOF 127, DOFOn
      End If
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
        For Each object in vrCab: object.visible = 1: Next
      Else
        vrOption = 1
        For Each object in vrRoom: object.visible = 0: Next
      End If
      optionMenu.image = "VR" & vrOption
        Case 3:
            If chime = 0 Then
                chime= 1
        If B2SOn Then DOF 141, DOFPulse
              Else
                chime = 0
        playSound "Bell10"
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


    If keycode = 46 Then' C Key
        If contBall = 1 Then
            contBall = 0
          Else
            contBall = 1
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
  If keycode = 30 Then ' "a" key
  End If

  If keycode = 31 Then ' "s" key
  End If

'************************End Of Test Keys****************************
End Sub

Sub Table1_KeyUp(ByVal keycode)

  If keycode = plungerKey Then
    plunger.Fire
    playFieldSound "PlungerFire", 0, plunger, 1
  End If

    If keycode = startGameKey and state = False Then
    VR_cab_buttonstart.y = 0
  End If

  If Keycode = startGameKey and contBall = 0 and state = True Then
    If vrOption > 0 Then
      VR_cab_balllift.Y = 0
    End If
  End If

  If Keycode = rightMagnaSave and contBall = 0 and state = True Then
    If vrOption > 0 Then
      VR_cab_balllift.Y = 0
    End If
  End If

    If keycode = leftFlipperKey Then
        operatorMenuTimer.Enabled = False
    End If

   If tilt = False and state = True Then
    If keycode = leftFlipperKey and contball = 0 Then
      nudgeUpTimer.enabled = 0
      nudgeDownTimer.enabled = 1
    End If

    If keycode = rightFlipperKey and contball = 0 Then
      nudgeUpTimer.enabled = 0
      nudgeDownTimer.enabled = 1
    End If
   End If

  If keycode = rightMagnaSave Then playFieldSound "BallLiftDown", 0, Plunger, .7


    If keycode = 203 Then cLeft = 0' Left Arrow

    If keycode = 200 Then cUp = 0' Up Arrow

    If keycode = 208 Then cDown = 0' Down Arrow

    If keycode = 205 Then cRight = 0' Right Arrow

  If keycode = 52 Then Zup = 0' Period

  If keycode = 30 Then
    nudgeUpTimer.enabled = 0
    nudgeDownTimer.enabled = 1
  End If

End Sub


Dim move, moveMinus, movePlus, obj
Sub nudgeUpTimer_timer
  move = move + 1
  If move < 1 Then move = 1
  If move > 5 Then move = 5
  If move > 1 Then
    meshes_visibletroughdrop.visible = 0
  Else
    meshes_visibletroughdrop.visible = 1
  End If
  Dim BOT, b
  BOT = GetBalls
  For b = 0 to uBound(BOT)
    If BOT(b).z < 0 Then BOT(b).vely = 5
  Next
  For each obj in EVAL("NudgePF" & move)
    obj.collidable = 1
  Next
  moveMinus = move - 1
  For Each obj in EVAL("NudgePF" & moveMinus)
    obj.collidable = 0
  Next
  EVAL("meshes_visible" & move).visible = 1
  EVAL("playfield_mesh" & move).visible = 1
  EVAL("meshes_diamond_visible" & move).visible = 1
  EVAL("meshes_lite_visible" & move).visible = 1
  EVAL("meshes_special_visible" & move).visible = 1
  EVAL("meshes_GI_visible" & move).visible = 1
  EVAL("meshes3_visible" & move).visible = 1
  EVAL("meshes2_visible" & move).visible = 1
  EVAL("meshes2_visible" & moveMinus).visible = 0
  EVAL("meshes3_visible" & moveMinus).visible = 0
  EVAL("meshes_GI_visible" & moveMinus).visible = 0
  EVAL("meshes_special_visible" & moveMinus).visible = 0
  EVAL("meshes_lite_visible" & moveMinus).visible = 0
  EVAL("meshes_diamond_visible" & moveMinus).visible = 0
  EVAL("meshes_visible" & moveMinus).visible = 0
  EVAL("playfield_mesh" & moveMinus).visible = 0
  NudgyHandle.rotz = 280 - (move * 10)

  If move = 5 Then
    nudgeUpTimer.enabled = 0
    If kickDelay = 1 Then
      kickBall 9, 8, 30
      kickDelay = 0
    End If
  End If
End Sub

Sub nudgeDownTimer_timer
  move = move - 1
  If move < 1 Then move = 1
  If move > 5 Then move = 5
  For each obj in EVAL("NudgePF" & move)
    obj.collidable = 1
  Next
  movePlus = move + 1
  If move > 1 Then
    meshes_visibletroughdrop.visible = 0
  Else
    meshes_visibletroughdrop.visible = 1
  End If
  Dim BOT, b
  BOT = GetBalls
  For b = 0 to uBound(BOT)
    If BOT(b).z < 0 Then BOT(b).vely = -5
  Next
  For Each obj in EVAL("NudgePF" & movePlus)
    obj.collidable = 0
  Next
  NudgyHandle.rotz = 280 - (move * 10)
  EVAL("meshes_visible" & move).visible= 1
  EVAL("playfield_mesh" & move).visible = 1
  EVAL("meshes_diamond_visible" & move).visible = 1
  EVAL("meshes_lite_visible" & move).visible = 1
  EVAL("meshes_special_visible" & move).visible = 1
  EVAL("meshes_GI_visible" & move).visible = 1
  EVAL("meshes3_visible" & move).visible = 1
  EVAL("meshes2_visible" & move).visible = 1
  EVAL("meshes2_visible" & movePlus).visible = 0
  EVAL("meshes3_visible" & movePlus).visible = 0
  EVAL("meshes_GI_visible" & movePlus).visible = 0
  EVAL("meshes_special_visible" & movePlus).visible = 0
  EVAL("meshes_lite_visible" & movePlus).visible = 0
  EVAL("meshes_diamond_visible" & movePlus).visible = 0
  EVAL("meshes_visible" & movePlus).visible = 0
  EVAL("playfield_mesh" & movePlus).visible = 0
  If move = 1 then nudgeDownTimer.enabled = 0
End Sub


'************** Table Boot
Dim backGlassOn
Dim bootCount:bootCount = 0
Sub bootTable_Timer
  score(1) = 0
  bootCount = bootCount + 1
  If bootCount = 1 Then
    If B2SOn Then
      If FreePlay = 1 Then DOF 127, DOFOn
      controller.B2SSetData 80, 1
    End If
    backGlassOn = 1
    bootCount = 0
    GIOn
    bootTable.enabled = False
    DTReel.setValue(1)
    If vrOption > 0 Then vrBackglass.image = "nudgyOn"
  End If
End Sub

'***********Operator Menu
Dim operatormenu

Sub operatorMenuTimer_Timer
  options = 0
    operatorMenu = 1
  dynamicUpdatePostIt.enabled = 0
  updatePostIt
  options = 0
    playFieldSound "target", 0, SoundPointScoreMotor, 1.5
    optionsMenu.visible = True
    optionMenu.visible = True
  optionMenu.image = "FreeCoin" & freePlay
End Sub

'***********Timer to delay lift of first ball until one ball is destroyed on game start
Dim starter, starterCount
Sub gameStart_Timer
  starterCount = starterCount + 1
  tb.text = "count = " & starterCount
  Select Case starterCount
    Case 1: starter = 1
        score(1) = 0
    Case 5: meshes_visibletroughdrop.rotx = 0
        starterCount = 0
        drainwall.collidable = 1
        gameStart.enabled = 0
  End Select
End Sub

'***********Start Game
Dim ballInPlay, bonusScore
Sub startGame
  If state = False Then
    ballInPlay = 0
    dynamicUpdatePostIt.enabled = 0
    updatePostIt
    tilt = False
    state = True
    players = 1
    DiamondLight1.state = 0
    DiamondLight2.state = 0
    TriangleLight001.state = 0
    TriangleLight002.state = 0
    DBLightsOn
    liteState = 0
    diamondState = 0
    light001.state = 0
    meshes_visibletroughdrop.rotx = 90
    drainwall.collidable = 0
    newGame
  End If
End Sub

'*********New Game
Dim endGame, roundHS
Sub newGame
  player = 1
    endGame = 0
  roundHS = 0
  gameState
  resetBackglass
  resetDt
  gameStart.enabled = 1
End Sub

'*************Check for Continuing Game
Sub checkContinue
  If endGame = 1 Then
    turnOff
    starter = 0
    state = False
    gameState
    dynamicUpdatePostIt.enabled = 1
    sortScores
    checkHighScores
    firstBallOut = 0
    players = 0
    saveHighScore
  End If
' NOTE because this is a ball lift table there is no Else to eject the next ball
End Sub

'***************Drain and Release Ball
Dim ballDrain, archHit
Sub drain_Hit
  If firstBallOut = True Then
'   playsound "CHLowerkickers"
    archHit = 0
    scoreMotorLoop = 0
    ballDrain = ballDrain + 1
    If ballDrain > ballInPlay Then ballDrain = ballInPlay
    If ballDrain = balls Then endGame = 1
    checkContinue
    Tilt = False
  End If
End Sub

Dim ballsKilled
Sub drainKill_hit()
  ballsKilled = ballsKilled + 1
  drainKill.destroyBall
End Sub

'**************Destroy Balls to Start New Game
Sub ballKiller_Hit()
  ballKiller.destroyball
End Sub


'***********Ball Lift Speed Limiter to Prevent Loss of Balls
Dim lifter, bip
lifter = 0
Sub ballLiftTimer_Timer
  lifter = lifter + 1
  If lifter = 1 Then
    If ballinPlay < balls Then
      playFieldSound "BallLift" , 0, Plunger, .7
      ballRelease.CreateBall
      ballRelease.kick 90, 5
      ballinPlay = ballinPlay + 1
      bip = ballInPlay - 1
    End If
  End If
  If lifter = 2 Then
    lifter = 0
    ballLiftTimer.enabled = 0
  End If
End Sub
'************Game State Check
Sub gameState
  If state = 0 Then
    ballDrain = 0
  Else
    tilt = False
  End If
End Sub

'*************Ball in Launch Lane on Plunger Tip
Dim ballREnabled
Sub ballHome_hit
  ArchHit = 0
  BallREnabled = 1
  Set ControlBall = ActiveBall
    contBallInPlay = True
End Sub

'******* For ball control script
Sub endControl_Hit()
    contBallInPlay = False
End Sub

'************Check if Ball Out of Launch Lane
Sub ballsInPlay_hit
  If BallREnabled = 1 Then
    BallREnabled = 0
  End If
  FirstBallOut = True
End Sub

'************** Bumpers
Sub bumpers_hit(index)
  addScore 5000
End Sub

Sub SpecialBumpers_hit(index)
  addScore 5000
  TriangleLight001.state = 1
  TriangleLight002.state = 1
  LiteOn
End Sub

'************* Diamonds
Dim doubleScore
Sub diamonds_hit(index)
  addScore 5000
  If DiamondLight1.state = 1 Then
    doubleScore = 1
    light001.state = 1
    TB.text = "Double Score"
  End If
End Sub

'************* Triggers
Sub triggers_hit(index)
  addScore 25000
  If TriangleLight001.state = 1 Then
    DiamondLight1.state = 1
    DiamondLight2.state = 1
    DiamondsOn
  End If
End Sub

'************* Saucers
Dim inSaucer, triggerCount
Sub saucers_hit(index)
' playFieldSound "CHCarHorn", 0, saucers(index), 1
  EVAL("SaucerTimer" & Index + 1).enabled = 1
  TB2.text = "saucer = " & (index+1)
End Sub

Dim saucerCount1, saucerCount2, saucerCount3, saucercount4
Sub SaucerTimer1_Timer
  TB3.text = "Timer 1 On"
  Dim BOT, b
  BOT = GetBalls
  saucerCount1 = saucerCount1 + 1
  If saucerCount1 > 1 Then
    For b = 0 to uBound(BOT)
      If BOT(b).x > 180 and BOT(b).x < 231 and BOT(b).y > 1397 and BOT(b).y < 1475 Then addScore 25000
      TB3.text = "Timer Off"
    Next
    inSaucer = 1
    saucerCount1 = 0
    SaucerTimer1.enabled = 0
  End If
End Sub

Sub SaucerTimer2_Timer
  TB3.text = "Timer 2 On"
  Dim BOT, b
  BOT = GetBalls
  saucerCount2 = saucerCount2 + 1
  If saucerCount2 > 1 Then
    For b = 0 to uBound(BOT)
      If BOT(b).x > 430 and BOT(b).x < 471 and BOT(b).y > 1498 and BOT(b).y < 1570 Then addScore 25000
      TB3.text = "Timer Off"
    Next
    inSaucer = 2
    saucerCount2 = 0
    SaucerTimer2.enabled = 0
  End If
End Sub

Sub SaucerTimer3_Timer
  TB3.text = "Timer 3 On"
  Dim BOT, b
  BOT = GetBalls
  saucerCount3 = saucerCount3 + 1
  If saucerCount3 > 1 Then
    For b = 0 to uBound(BOT)
      If BOT(b).x > 680 and BOT(b).x < 712 and BOT(b).y > 1400 and BOT(b).y < 1470 Then addScore 25000
      TB3.text = "Timer Off"
    Next
    inSaucer = 3
    saucerCount3 = 0
    SaucerTimer3.enabled = 0
  End If
End Sub

Sub SaucerTimer4_Timer
  TB3.text = "Timer On"
  Dim BOT, b
  BOT = GetBalls
  saucerCount4 = saucerCount4 + 1
  If saucerCount4 > 1 Then
    For b = 0 to uBound(BOT)
      If BOT(b).x > 430 and BOT(b).x < 473 and BOT(b).y > 789 and BOT(b).y < 871 Then addScore 25000
      TB3.text = "Timer Off"
    Next
    inSaucer = 4
    saucerCount4 = 0
    SaucerTimer4.enabled = 0
  End If
End Sub

dim kickStep1, kickStep2

'************ Kickers
'These check to see if there is a ball in the x/y location and kicks the balls

Sub kickBall(kvel, kvelz, kzlift)
  dim rangle, BOT, b
  BOT = GetBalls
  Select Case inSaucer
    Case 1: For b = 0 to uBound(BOT)
          If BOT(b).x > 190 and BOT(b).x < 221 and BOT(b).y > 1407 and BOT(b).y < 1465 Then
            rangle = 3.14 * (90/180)
            BOT(b).z = BOT(b).z + kzlift
            BOT(b).velz = kvelz
            BOT(b).velx = cos(rangle)*kvel
            BOT(b).vely = sin(rangle)*kvel
            If B2SOn Then DOF 104, DOFPulse
          End If
        Next
    Case 2: For b = 0 to uBound(BOT)
          If BOT(b).x > 440 and BOT(b).x < 461 and BOT(b).y > 1502 and BOT(b).y < 1560 Then
            rangle = 3.14 * (90/180)
            BOT(b).z = BOT(b).z + kzlift
            BOT(b).velz = kvelz
            BOT(b).velx = cos(rangle)*kvel
            BOT(b).vely = sin(rangle)*kvel
            If B2SOn Then DOF 104, DOFPulse
          End If
        Next
    Case 3: For b = 0 to uBound(BOT)
          If BOT(b).x > 680 and BOT(b).x < 701 and BOT(b).y > 1407 and BOT(b).y < 1462 Then
            rangle = 3.14 * (90/180)
            BOT(b).z = BOT(b).z + kzlift
            BOT(b).velz = kvelz
            BOT(b).velx = cos(rangle)*kvel
            BOT(b).vely = sin(rangle)*kvel
            If B2SOn Then DOF 104, DOFPulse
          End If
        Next
    Case 4: For b = 0 to uBound(BOT)
          If BOT(b).x > 440 and BOT(b).x < 463 and BOT(b).y > 799 and BOT(b).y < 861 Then
            rangle = 3.14 * (90/180)
            BOT(b).z = BOT(b).z + kzlift
            BOT(b).velz = kvelz
            BOT(b).velx = cos(rangle)*kvel
            BOT(b).vely = sin(rangle)*kvel
            If B2SOn Then DOF 104, DOFPulse
          End If
        Next
  End Select
End Sub

Dim diamondState
Sub DiamondsOn
  diamondState = 1
  meshes_diamond_visible1.image = "render2Lit"
  meshes_diamond_visible2.image = "render2Lit"
  meshes_diamond_visible3.image = "render2Lit"
  meshes_diamond_visible4.image = "render2Lit"
  meshes_diamond_visible5.image = "render2Lit"
End Sub

Sub DiamondsOff
  meshes_diamond_visible1.image = "render2Off"
  meshes_diamond_visible2.image = "render2Off"
  meshes_diamond_visible3.image = "render2Off"
  meshes_diamond_visible4.image = "render2Off"
  meshes_diamond_visible5.image = "render2Off"
End Sub

Dim liteState
Sub LiteOn
  liteState = 1
  meshes_lite_visible1.image = "render2Lit"
  meshes_lite_visible2.image = "render2Lit"
  meshes_lite_visible3.image = "render2Lit"
  meshes_lite_visible4.image = "render2Lit"
  meshes_lite_visible5.image = "render2Lit"
End Sub

Sub LiteOff
  meshes_lite_visible1.image = "render2Off"
  meshes_lite_visible2.image = "render2Off"
  meshes_lite_visible3.image = "render2Off"
  meshes_lite_visible4.image = "render2Off"
  meshes_lite_visible5.image = "render2Off"
End Sub

Sub GIOn
  tb.text = "ON"
  meshes_gi_visible1.image = "render2Lit"
  meshes_gi_visible2.image = "render2Lit"
  meshes_gi_visible3.image = "render2Lit"
  meshes_gi_visible4.image = "render2Lit"
  meshes_gi_visible5.image = "render2Lit"
  meshes_special_visible1.image = "render2Lit"
  meshes_special_visible2.image = "render2Lit"
  meshes_special_visible3.image = "render2Lit"
  meshes_special_visible4.image = "render2Lit"
  meshes_special_visible5.image = "render2Lit"
  meshes_visible1.image = "render1"
  meshes2_visible1.image = "render1"
  meshes_visible2.image = "render1"
  meshes2_visible2.image = "render1"
  meshes_visible3.image = "render1"
  meshes2_visible3.image = "render1"
  meshes_visible4.image = "render1"
  meshes2_visible4.image = "render1"
  meshes_visible5.image = "render1"
  meshes2_visible5.image = "render1"
' meshes_static.image = "render1"
  playfield_mesh.image = "NudgyPF4"
  meshes_visibletroughdrop.image ="render1"
End Sub

Sub GIOff
  meshes_gi_visible1.image = "render2Off"
  meshes_gi_visible2.image = "render2Off"
  meshes_gi_visible3.image = "render2Off"
  meshes_gi_visible4.image = "render2Off"
  meshes_gi_visible5.image = "render2Off"
  meshes_special_visible1.image = "render2Off"
  meshes_special_visible2.image = "render2Off"
  meshes_special_visible3.image = "render2Off"
  meshes_special_visible4.image = "render2Off"
  meshes_special_visible5.image = "render2Off"
  meshes_visible1.image = "render1Off"
  meshes2_visible1.image = "render1Off"
  meshes_visible2.image = "render1Off"
  meshes2_visible2.image = "render1Off"
  meshes_visible3.image = "render1Off"
  meshes2_visible3.image = "render1Off"
  meshes_visible4.image = "render1Off"
  meshes2_visible4.image = "render1Off"
  meshes_visible5.image = "render1Off"
  meshes2_visible5.image = "render1Off"
' meshes_static.image = "render1Off"
  playfield_mesh.image = "NudgyPF4Off"
  meshes_visibletroughdrop.image ="render1Off"
End Sub

Dim lightsOff
Sub DBLightsOff
  Dim xx
  lightsOff = 1
  For Each xx in DeadBumperLights
    xx.state = 0
  Next
End Sub

Sub DBLightsOn
  Dim xx
  lightsOff = 0
  For Each xx in DeadBumperLights
    xx.state = 1
  Next
End Sub

'****************Score Motor Run Timer
Dim bellRing, scoreMotorLoop, kickDelay, scoreLoop, motorOn
Sub scoreMotor_timer
  motorOn = 1
  scoreMotorLoop = scoreMotorLoop + 1
  If scoreMotorLoop < 6 Then playFieldSound "ScoreMotorSingleFire", 0, SoundPointScoreMotor, .3

'  These Flags are passed by scores with multiple of 10, 100 or 1000
  If  flag1000 = 1 or flag5000 = 1 Then
    Select Case scoreMotorLoop
      Case 1: totalUp point: DBLightsOff: GIOff: DiamondsOff: LiteOff
      Case 2: totalUp point
      Case 3: If bellRing > 2 Then totalUp point
      Case 4: If bellRing > 3 Then totalUp point
      Case 5: If bellRing > 4 Then totalUp point
          If Flag5000 = 1 and nudgeUpTimer.enabled = 0 Then
            kickBall 9, 8, 30
          Else
            kickDelay = 1
          End If
      Case 6: If doubleScore = 0 Then
            scoreMotorLoop = 0
            flag1000 = 0
            flag5000 = 0
            DBLightsOn
            GIOn
            If diamondState = 1 Then DiamondsOn
            If liteState = 1 Then LiteOn
            motorOn = 0
            scoremotor.enabled = 0
          ElseIf doubleScore = 1 and scoreLoop = 0 Then
            scoreMotorLoop = 0
            scoreLoop = 1 ' allows second go through for double score
            If diamondState = 1 Then DiamondsOn
            If liteState = 1 Then LiteOn
          ElseIf doubleScore = 1 and scoreLoop = 1 Then
            scoreMotorLoop = 0
            flag1000 = 0
            flag5000 = 0
            scoreLoop = 0
            DBLightsOn
            GIOn
            If diamondState = 1 Then DiamondsOn
            If liteState = 1 Then LiteOn
            motorOn = 0
            scoremotor.enabled = 0
          End If
    End Select
  End If
End Sub

Sub resetBackglass
  For x = 1 to 9
    If B2SOn Then
      controller.B2SSetData (60 + x), 0
      controller.B2SSetData (70 + x), 0
    End If
  Next
End Sub

Sub resetDT
  Reel5k.setValue(0)
  Reel10k.setValue(0)
  Reel100k.setValue(0)
  vr5k.image = "5k0"
  vr10k.image = "10k0"
  vr100k.image = "100k0"
End Sub

'***************Scoring Routine
Dim flag1000, flag5000, point, point1, ones
Sub addScore(points)
  If tilt = False and lightsOff = 0 and motorOn = 0 Then

    If points = 5000 Then
      bellRing = 5
      point = 1000: flag1000 = 1: scoreMotorLoop = 0: scoreMotor.enabled = 1
    End If

    If points > 5000 Then
      bellRing = (points / 5000)
      point = 5000: flag5000 = 1: scoreMotorLoop = 0: scoreMotor.enabled = 1
    End If


  End If
End Sub

Dim fiveK, tenK, hundredK
Sub totalUp(points)
  If flag1000 = 1 Then
    If chime = 0 Then
      playSound "Bell10"
    Else
      If B2SOn Then DOF 141,DOFPulse
    End If
  End If

  If flag5000 = 1 Then
    If chime = 0 Then
      playSound "Bell10"
    Else
      If B2SOn Then DOF 141,DOFPulse
    End If
  End If
  resetBackglass

  score(player) = score(player) + points
  fiveK = score(1) mod 10000
  tenK = (score(1) \ 10000) mod 10
  hundredK = score(1) \ 100000 mod 10

  If fiveK > 0 Then
    Reel5k.setValue(1)
    If B2SOn Then controller.B2SSetData 50, 1
    vr5k.image = "5k"
  Else
    Reel5k.setValue(0)
    If B2SOn Then controller.B2SSetData 50, 0
    vr5k.image = "5k0"
  End If

  If tenK > 0 And tenK < 10 Then
    Reel10k.setValue(tenK)
    If B2SOn Then controller.B2SSetData (60 + tenK), 1
    vr10k.image = ("10k" & tenK)
  Else
    Reel10k.setValue(0)
    vr10k.image = "10k0"
  End If

  If hundredK > 0 And hundredK < 10 Then
    Reel100k.setValue(hundredK)
    If B2SOn Then controller.B2SSetData (70 + hundredK), 1
    vr100k.image = ("100k" & hundredK)
  Else
    Reel100k.setValue(0)
    vr100k.image = "100k0"
  End If
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

'************************************************Post It Note Section**************************************************************************
'***************Static Post It Note Update
Dim  hsY, shift, scoreMil, score100K, score10K, scoreK, score100, score10, scoreUnit
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
  flag = 0
  For hs = 1 to maxPlayers              'look at 4 player scores
    For chY = 0 to 4                  'look at all 5 saved high scores
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
'     playerEntry.visible = 1
'     playerEntry.image = "Player" & position(Flag)
'     TextBox3.text = ActiveScore(Flag) 'tells which high score is being entered
'     TextBox2.text = Flag
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
    End If
End Sub


'************Enter Initials Keycode Subroutine
Dim initial(6,5), initialsDone
Sub enterInitials(keycode)
    If keyCode = leftFlipperKey Then
      hsX = hsX - 1           'HSx is the inital to be displayed A-Z plus " "
      If hsX < 0 Then hsX = 26
      If hsI < 4 Then EVAL("Initial" & hsI).image = hsIArray(hsX)   'HSi is which of the three intials is being modified
      playSound "metalhit_thin"
    End If
    If keycode = rightFlipperKey Then
      hsX = hsX + 1
      If hsX > 26 Then hsX = 0
      If hsI < 4 Then EVAL("Initial"& hsI).image = hsIArray(hsX)
      playSound "metalhit_thin"
    End If
    If keycode = startGameKey and initialsDone = 0 Then
      If hsI < 3 Then                 'if not on the last initial move on to the next intial
        EVAL("Initial" & hsI).image = hsIArray(hsX) 'display the initial
        initial(activeScore(flag), hsI) = hsX   'save the inital
        playSound "metalhit_medium"
        EVAL("InitialTimer" & hsI).enabled = 0    'turn that inital's timer off
        EVAL("Initial" & hsI).visible = 1     'make the initial not flash but be turn on
        initial(activeScore(flag),hsI + 1) = hsX  'move to the next initial and make it the same as the last initial
        EVAL("Initial" & hsI +1).image = hsIArray(hsX)  'display this intial
'       y = 1
        EVAL("InitialTimer" & hsI + 1).enabled = 1  'make the new intial flash
        hsI = hsI + 1               'increment HSi
      Else                    'if on the last initial then get ready yo exit the subroutine
'       enableInitialEntry = False
'       highScoreDelay.enabled = 1
        initial3.visible = 1          'make the intial visible
        playSound "metalhit_medium"
        initialTimer3.enabled = 0       'shut off the flashing
        initial(activeScore(flag),3) = hsX    'set last initial
        initialEntry              'exit subroutine
      End If
    End If
End Sub

'************Update Initials and see if more scores need to be updated
Dim eIX
Sub initialEntry
  playsound SoundFXDOF("Chime100",141,DOFPulse,DOFChimes)
  flag = flag - 1
  hsI = 1
  If flag < 0 Then flag = 0: Exit Sub
  If flag = 0 Then          'exit high score entry mode and reset variables
    initialsDone = 1
    players = 0
    For eIX = 1 to 4
      activeScore(eIX) = 0
      position(eIX) = 0
    Next
    For eIX = 1 to 3
      EVAL("InitialTimer" & eIX).enabled = 0
    Next
'   playerEntry.visible = 0
    scoreUpdate = 0           'go to the highest score
    updatePostIt            'display that score
    highScoreDelay.enabled = 1
  Else
    showScore
'   playerEntry.image = "Player" & position(flag)
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
Dim hsCount
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

    For x = 1 to 32
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
  If tiltPenalty = 1 then ballInPlay = Balls
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

'*****************************************
'     BALL SHADOW
'*****************************************
Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5,BallShadow6)

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
            BallShadow(b).X = ((BOT(b).X) + ((BOT(b).X - (Table1.Width/2))/7))
        Else
            BallShadow(b).X = ((BOT(b).X) + ((BOT(b).X - (Table1.Width/2))/7))
        End If
        ballShadow(b).Y = BOT(b).Y + 10
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
End Sub

'************************************************************************
'                         Ball Control - 3 Axis
'************************************************************************

Dim Cup, Cdown, Cleft, Cright, Zup, contball, contballinplay, ControlBall, bcboost
Dim bcvel, bcyveloffset, bcboostmulti

bcboost = 1 'Do Not Change - default setting
bcvel = 4 'Controls the speed of the ball movement
bcyveloffset = -0.014 'Offsets the force of gravity to keep the ball from drifting vertically on the table, should be negative
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

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
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
'Subtracting XVol or YVol from 1 yeilds an inverse response.

Const tnob = 5 ' total number of balls

'Change GHT, GHB and PFL values based upon the real pinball table dimensions.  Values are used by the GlassHit code.
Const GHT = 2 'Glass height in inches at top of real playfield
Const GHB = 2 'Glass height in inches at bottom of real playfield
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
  Dim BOT, b, paSub
  BOT = GetBalls
  paSub=35000 'Playsound pitch adder for subway rolling ball sound

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
    ElseIf rolling(b) = True OR BallVel(BOT(b)) < 0.1 AND BOT(b).z < -5 Then
'   ElseIf rolling(b) = True Then
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
    ElseIf rolling(b) = True OR BallVel(BOT(b)) < 0.1 AND BOT(b).z < -5 Then
'   ElseIf rolling(b) = True Then
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
' Ball=50 units=1.0625".  Ball.z is ball center.  Balls are physically limited by Top Glass Height.  Max ball.z is 25 units below Top Glass Height.
' To ensure ball can go high enough to trigger glass hit, make Table Options/Dimensions & Slope/Top Glass Height equal to (GHT*50/1.0625) + 5

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
'The Hit Subs use PlayFieldSoundAB that incorporates the ball’s velocity.

Sub aRubberPins_Hit(idx)
  PlayFieldSoundAB "pinhit_low", 0, 1
End Sub

Sub aRubberPosts_Hit(idx)
  PlayFieldSoundAB "pinhit_low", 0, 1
End Sub

Sub aMushroom_Hit(idx)
  PlayFieldSoundAB "pinhit_low", 0, 1
End Sub

Sub aTargets_Hit(idx)
  PlayFieldSoundAB "target", 0, 1
End Sub

Sub aMetalsThin_Hit(idx)
  PlayFieldSoundAB "metalhit_thin", 0, 1
End Sub

Sub aMetalsMedium_Hit(idx)
  PlayFieldSoundAB "metalhit_medium", 0, 1
End Sub

Sub aMetals2_Hit(idx)
  PlayFieldSoundAB "metalhit2", 0, 1
End Sub

Sub aGates_Hit(idx)
  PlayFieldSoundAB "gate4", 0, 1
End Sub

Sub aRubberBands_Hit(idx)
  If BallinPlay > 0 Then
  PlayFieldSoundAB "rubber_hit_2", 0, 0.1
  End If
End Sub

Sub RubberWheel_hit
  PlayFieldSoundAB "rubber_hit_1", 0, 0.5
End sub

Sub aPosts_Hit(idx)
  PlayFieldSoundAB "rubber_hit_3", 0, 1
End Sub

Sub aWoods_Hit(idx)
  PlayFieldSoundAB "Wood", 0, 1
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

    Sub ApronWalls_Hit
  Dim Volume
  If ActiveBall.vely < 0 Then Volume = abs(ActiveBall.vely) / 1 Else Volume = ActiveBall.vely / 30  'The first bounce is -vely subsequent bounces are +vely
'   TextBox1.text = "Volume = " & Volume
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

'New algorithm for OnBallBallCollision
'For stereo, xGain is a Playsound volume multiplier that provides a Constant Power pan.
'For quad, multiple PlaySound commands are launched together that are panned and faded to their maximum extents where PlaySound's PAN and FADE have the least error.
'XVol and YVol are Playsound volume multipliers that provide a Constant Power "pan" and "fade".
'Subtracting XVol or YVol from 1 yeilds an inverse response.

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


'******************************************************
'       FLIPPER AND RUBBER CORRECTION
'******************************************************

dim LFPress, RFPress, EOST, EOSA, EOSTnew, EOSAnew
dim FStrength, Frampup, FElasticity, EOSRampup, SOSRampup
dim RFEndAngle, LFEndAngle, LFCount, RFCount, LiveCatch


'********Need to have a flipper timer to check for these values
Sub flipperTimer_Timer
' TB.text = "Drain = " & ballDrain
  ScoreBox.text = score(1)
' TB2.Text = "tCount = " & triggerCount
' If contBallInPlay = True Then
'   Dim BOT, b
'   BOT = GetBalls
'   TB2.text = "x = " & BOT(0).x
'   TB3.text = "Y = " & BOT(0).y
' End If
End Sub


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

Sub Table1_Exit
  If B2SOn Then controller.stop
  saveHighScore
End Sub

'***************** CODE BELOW IS FOR THE VR CLOCK.  This was originally taken from Rascal VP9 clock table *******************************
'*****************************************************************************************************************************************

Dim CurrentMinute ' for VR clock

Sub ClockTimer_Timer()

    'ClockHands Below *********************************************************
  VR_ClockMinutes.RotAndTra2 = (Minute(Now())+(Second(Now())/100))*6
  VR_ClockHours.RotAndTra2 = Hour(Now())*30+(Minute(Now())/2)
  VR_ClockSeconds.RotAndTra2 = (Second(Now()))*6
  CurrentMinute=Minute(Now())

End Sub

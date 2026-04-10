'****************************************************************
'
'          Nags (Williams 1960)
'          Script by Scottacus
'              v 1.00
'              April 2025
'
'   DOF config
'   401 Left Flipper, 402 Right Flipper
'   405 -409 Top InLanes
'   410 -415 OutLanes
'   420 Bumpers
'   427 StartButton
'   428 Knocker
'   429 Knocker and Kicker Strobe
'   441 Chime/Bell 10 pts
'   442 Chime/Bell Race Over
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

Const cGameName = "nags_1960"
Const cOptions = "Nags_Williams_v.1.00.txt"
Const HSFileName="Nags (Williams 1960)"

'*************************** Set Flex DMD **********************************************
DMD_init

'*************************** Ball Lift Tap *********************************************
Dim ballLiftTap 'This will determine if the right magna/start button needs to be held in to lift
ballLiftTap = 1 'Setting to 0 means hold button to lift, 1 just tap to lift

'*************************** Tilt Ends Game ********************************************
Dim tiltEndsGame  ' Set this to True if you want the game to play like the schematic
tiltEndsGame = False

'*************************** PinballY Settings *****************************************
'Set this variable to 1 to save a PinballY High Score file to your Tables Folder
'this will let the Pinball Y front end display the high scores when searching for tables
'0 = No PinballY High Scores, 1 = Save PinballY High Scores
Const cPinballY = 1

Dim vrOption
'*************************** VR Room****************************************************
Dim VR_Room:
If RenderingMode = 2 Then
  vrOption = 2
Else
  vrOption = 0
End If

Const VRGlassScratches = 0  ' set this to 1 to turn on VR glass scratches
If vrOption > 0 and VRGlassScratches > 0 then GlassImpurities.visible = true

'***********************************************************************************************

Dim tablewidth: tablewidth = Table1.width
Dim tableheight: tableheight = Table1.height

Sub FrameTimer_Timer()
  FlipperVisualUpdate       'update flipper shadows and primitives
  AnimateBumperSkirts
' If DynamicBallShadowsOn Or AmbientBallShadowOn Then DynamicBSUpdate 'update ball shadows
End Sub

dim gateAngle
' The game timer interval is 10 ms
Sub GameTimer_Timer()
  Cor.Update            'update ball tracking (this sometimes goes in the RDampen_Timer sub)
' RollingUpdate         'update rolling sounds
End Sub

dim BP

' This subroutine updates the flipper shadows and visual primitives
Sub FlipperVisualUpdate
  For each BP in BP_pflflip : BP.rotz=leftflipper.currentangle : Next
  For each BP in BP_pfrflip : BP.rotz=rightflipper.currentangle : Next
  For each BP in BP_pflrub : BP.rotz=leftflipper.currentangle : Next
  For each BP in BP_pfrrub : BP.rotz=rightflipper.currentangle : Next
  FlipperLSh001.RotZ = LeftFlipper.currentangle
  FlipperRSh001.RotZ = RightFlipper.currentangle
End Sub

'****** End New ball Shadow Code..  More at bottom of script ****************************************************************************

Dim balls
balls = 5
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
Dim i,j, f, ii, Object, Light, x, y, z
Dim freePlay
Dim ballsize,BallMass
ballSize = 50
ballMass = (Ballsize^3)/125000
Dim BIP : BIP = 0
Dim options
Dim chime
Dim pfOption
Dim lutValue
Dim gBall0, gBall1, gBall2, gBall3, gBall4
Dim gBOT
Dim B2SMethod
Dim B2SVersion

'Desktop reel pans,  Valid values: "Lpan", "Mpan" and "Rpan"
Const dtcrReel = "Rpan"
Const dts1Reel = "Lpan"
Const dts2Reel = "Rpan"
Const dts3Reel = "Lpan"
Const dts4Reel = "Rpan"

'Backglass reel pans,  Valid values: "Lpan", "Mpan" and "Rpan"
Const bgcrReel = "Rpan"
Const bgs1Reel = "Mpan"
Const bgs2Reel = "Rpan"
Const bgs3Reel = "Lpan"
Const bgs4Reel = "Rpan"

Sub Table1_init
  LoadEM
  maxPlayers=1
  LoadHighScore
  If vrOption > 0 Then
    dtReel.setValue(0)
    If vrOption = 2 Then
      For Each object in vrRoom: object.visible = 1: Next
    Else
      For Each object in vrRoom: object.visible = 0: Next
    End If
  Else
     If showDT = True And vrOption = 0 Then dtReel.setValue(1)
  End If

  For x = 1 to maxPlayers
    Set sReel(x) = EVAL("scoreReel" & x)
  Next

  Player=1

  selectorCount =  Int(Rnd*12)

  If highScore(0) = "" Then highScore(0) = 500
  If highScore(1) = "" Then highScore(1) = 400
  If highScore(2) = "" Then highScore(2) = 300
  If highScore(3) = "" Then highScore(3) = 250
  If HighScore(4) = "" Then highScore(4) = 200
  If showDT = True Then pfOption = 1
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
  If lutValue = "" Then lutValue = 0

  replaySettings
  Table1.ColorGradeImage = "LUT" & lutValue
  updatePostIt
  dynamicUpdatePostIt.enabled = 1
  yellowHatReel.setValue(2)
  CreditReel.setvalue(credit)
  If vrOption > 0 Then vrCredit

  If ShowDT = True and vrOption = 0 Then
    For each object in backdropstuff
      object.visible = 1
    Next
  Else
    For each object in backdropstuff
      object.visible = 0
      table1.BackdropImage_DT = ""
    Next
  End If

  For Each xx in controlLights
    xx.visible = 0
  Next

  tilt = False
  state = False
  gameState
  bootTable.enabled = 1

  CreateTurnTable


'***********Trough Ball Creation
    'Ball initializations need for physical trough
  Set gBall0 = KickerTop0.CreateSizedballWithMass(Ballsize / 2,Ballmass)
  Set gBall1 = KickerTop1.CreateSizedballWithMass(Ballsize / 2,Ballmass)
  Set gBall2 = KickerTop2.CreateSizedballWithMass(Ballsize / 2,Ballmass)
  Set gBall3 = KickerTop3.CreateSizedballWithMass(Ballsize / 2,Ballmass)
  Set gBall4 = KickerTop4.CreateSizedballWithMass(Ballsize / 2,Ballmass)
    gBOT = Array(gBall0, gBall1, gBall2, gBall3, gBall4)

  troughwall1.collidable = 1
  wallState = 1
  troughwall2.collidable = 0

  for each xx in topKickers
    xx.enabled = 0
    xx.kickz 70,1,0,-25
  Next
End Sub

Dim liftCount
Sub liftTap_timer
  If liftCount > 0 Then
    liftPress = 0
    liftCount = 0
    liftTap.Enabled = 0
  End If
  liftCount = liftCount + 1
End Sub

'***********KeyCodes
Dim enableInitialEntry, liftPress
Sub Table1_KeyDown(ByVal keycode)

  If enableInitialEntry = True Then enterIntitals(keycode)

  If keycode = addCreditKey Then
    playFieldSound "coinin",0,KickerTop4,1
    addCredit = 1
    If B2SOn Then DOF 427, DOFOn
    increaseCredits
    End If

    If keycode = startGameKey Then
    cabBM_b_start.y = 2063 - 5
    If (enableInitialEntry = False and operatormenu = 0 and backGlassOn = 1 and endgame = 1) or tilt = True Then
      If freePlay = 1 Then startGame
      If freePlay = 0 and credit > 0  Then
        credit = credit - 1
        If showDT = False Then PlayReelSound "Reel5", bgcrReel Else PlayReelSound "Reel5", dtcrReel
        creditReel.setvalue(credit)
        If credit < 1 Then
          If B2SOn Then DOF 427, DOFOff
        End If
        If B2SOn Then
          If freeplay = 0 Then controller.B2SSetCredits credit
          If freePlay = 0 and credit < 1 Then DOF 427, DOFOff
        End If
        startGame
      End If
    End If
  End If


  If keycode = rightFlipperKey Then
    If vrOption > 0 Then cabBM_b_rflip.X = 477 - 8
  End If

  If keycode = leftFlipperKey and contball = 0 Then
    If vrOption > 0 Then cabBM_b_lflip.X = 477 + 8
  End If

  If keycode = PlungerKey Then
    plunger.PullBack
    playFieldSound "plungerpull", 0, plunger, 1
    If vrOption > 0 Then
      TimerVRPlunger.Enabled = True
      TimerVRPlunger1.Enabled = False
      cabBM_b_plunge.Y = 2128
    End If
  End If

  If tilt = False and state = True Then
  If keycode = leftFlipperKey and contball = 0 and flippersOn = True Then
    FlipperActivate LeftFlipper, LFPress
    lf.fire
    playFieldSound "FlipUpL", 0, leftFlipper, 1
    If B2SOn Then DOF 401,DOFOn
    playFieldSound "FlipBuzzL", -1, leftFlipper, 1
  End If

  If keycode = RightFlipperKey and contball = 0 and flippersOn = True Then
    FlipperActivate RightFlipper, RFPress
    rf.fire
    playFieldSound "FlipUpR", 0, RightFlipper,1
    If B2SOn Then DOF 402,DOFOn
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

  If Keycode = startGameKey and contBall = 0  and not isEmpty(liftBall) and enableInitialEntry = False Then
    ballLift
    If ballLiftTap = 1 Then liftTap.Enabled = 1
    liftPress = 1
  End If

  If Keycode = rightMagnaSave and contBall = 0 Then
    ballLift
    If ballLiftTap = 1 Then liftTap.Enabled = 1
    liftPress = 1
  End If

  If keycode = RightMagnaSave and state = false Then
  End If

  If keycode = LeftMagnaSave and state = false Then
  End If

    If keycode = leftFlipperKey and state = False and operatorMenu = 0 and enableInitialEntry = 0 Then
        operatorMenuTimer.Enabled = true
    End If

    If keycode = leftFlipperKey and state = False and operatorMenu = 1 Then
    options = options + 1
    If vrOption = 0 Then If options = 3 Then options = 4
    If showDt = True Then If options = 4 Then options = 5 'skips non DT options
        If options = 6 Then options = 0
    optionMenu.visible = True
        playFieldSound "target", 0, SoundPointScoreMotor, .2
        Select Case (Options)
            Case 0:
                optionMenu.image = "FreeCoin" & freePlay
      Case 1:
        If showDT = False or vrOption > 0 Then optionMenu.image = "UnderCab"
        If showDT = True And vrOption = 0 Then OptionMenu.visible = False
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
        OptionMenu1.image = "currentLUT"
        OptionMenu.image = "LUT" & (lutValue+1) & "m"
        For x = 1 to 6
          EVAL("Speaker" & x).visible = 0
        Next
      Case 3:
        optionMenu1.visible = 0
        OptionMenu2.visible = 0
        optionMenu.image = "VR" & vrOption
      Case 4:
        optionMenu2.visible = 0
        For x = 1 to 6
          EVAL("Speaker" & x).visible = 0
        Next
        optionMenu1.image = "DOF"
        optionMenu.image = "Chime" & chime
      Case 5:
        For x = 1 to 6
          EVAL("Speaker" & x).visible = 0
        Next
        optionMenu1.visible = 0
        optionMenu.image = "SaveExit"
        optionMenu2.visible = 0
    End Select
    End If

    If keycode = RightFlipperKey and state = False and operatorMenu = 1 Then
        playFieldSound "metalHitMedium", 0, SoundPointScoreMotor, 0.2
      Select Case (options)
    Case 0:
            If freePlay = 0 Then
                freePlay = 1
        If B2SOn Then DOF 427, DOFOn
              Else
                freePlay = 0
        If credit < 1 Then
          If B2SOn Then DOF 427, DOFOff
        End If
        If credit > 0 Then
          If B2SOn Then DOF 427, DOFOn
        End If
            End If
      optionMenu.image = "FreeCoin" & freePlay
    Case 1:
      optionMenu1.visible = 1
      pfOption = pfOption + 1
      If showDt = True and pfOption = 2 Then pfOption = 3
      If pfOption = 4 Then pfOption = 1
      optionMenu1.image = "Sound" & pfOption
      Select Case (pfOption)
        Case 1: speaker1.visible = 1: speaker2.visible = 1: speaker3.visible = 0: speaker4.visible = 0
        Case 2: speaker5.visible = 1: speaker6.visible = 1: speaker1.visible = 0: speaker2.visible = 0
        Case 3: speaker1.visible = 1: speaker2.visible = 1: speaker3.visible = 1: speaker4.visible = 1: speaker5.visible = 0: speaker6.visible = 0
      End Select
    Case 2:
      playFieldSound "metalHitHigh", 0, SoundPointScoreMotor, 0.2
      lutValue = lutValue + 1
      If lutValue > 9 Then lutValue = 0
      OptionMenu.image = "LUT" & (lutValue+1) & "m"
      setLUT
    Case 3:
      If vrOption = 1 Then
        vrOption = 2
        For Each object in vrRoom: object.visible = 1: Next
      Else
        vrOption = 1
        For Each object in vrRoom: object.visible = 0: Next
      End If
      optionMenu.image = "VR" & vrOption
        Case 4:
            If chime = 0 Then
                chime= 1
        If B2SOn Then DOF 441,DOFPulse
              Else
                chime = 0
        pts10
            End If
      optionMenu.image = "Chime" & chime
        Case 5:
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
    flippersOn = False
    yellowHatReel.setValue(2)
    If B2SOn Then controller.B2SSetTilt 1
    tiltLight.state = 1
    turnOff
  End If

    If keycode = 46 Then' C Key
       flipOff 2
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
  If keycode = 30 Then  '"a
    tiltLight.state = 1
  End if

  If Keycode = 31 Then  '"s"
    tiltLight.state = 0
  End If

  If keycode = 33 Then  '"f"
    controller.B2SSetData 18,1
  End If
'************************End Of Test Keys****************************
End Sub
Dim hSlot, yHat
dim TestChr1, TestChr2, TestChr3, TestChr

Sub Table1_KeyUp(ByVal keycode)

  If keycode = rightFlipperKey Then
    If vrOption > 0 Then cabBM_b_rflip.X = 477
  End If

  If keycode = leftFlipperKey and contball = 0 Then
    If vrOption > 0 Then cabBM_b_lflip.X = 477
  End If

  If keycode = StartGameKey Then
    If liftPress = 1 Then playFieldSound "BallLiftDown" , 0, Plunger, .7
    If ballLiftTap = 0 Then liftpress = 0       'discontinue ball lift
    cabBM_b_lift.y = 2049
    cabBM_b_start.y = 2063
  End If

  If Keycode = rightMagnaSave Then
    If liftPress = 1 Then playFieldSound "BallLiftDown" , 0, Plunger, .7
    cabBM_b_lift.y = 2049
    If ballLiftTap = 0 Then liftpress = 0       'discontinue ball lift
  End If

  If keycode = plungerKey Then
    If vrOption > 0 Then
      TimerVRPlunger.Enabled = False
      TimerVRPlunger1.Enabled = True
      cabBM_b_plunge.Y = 2128
    End If
    plunger.Fire
    playFieldSound "PlungerFire", 0, plunger, 1
  End If

    If keycode = leftFlipperKey Then
        operatorMenuTimer.Enabled = False
    End If

   If tilt = False and state = True Then
    If keycode = leftFlipperKey and contball = 0 and flippersOn = True Then
      FlipperDeActivate LeftFlipper, LFPress
      lfPress = 0
      LeftFlipper.eosTorqueAngle = EOSA
      LeftFlipper.eosTorque = EOST
      LeftFlipper.RotateToStart
      playFieldSound "FlipDownL", 0, leftFlipper, 1
      If B2SOn Then DOF 401,DOFOff
      flipOff 0
    End If

    If keycode = RightFlipperKey and contball = 0 and flippersOn = True Then
      FlipperDeActivate RightFlipper, RFPress
      rfpress = 0
      RightFlipper.eosTorqueAngle = EOSA
      RightFlipper.eosTorque = EOST
      RightFlipper.rotateToStart
      playFieldSound "FlipDownR", 0, RightFlipper, 1
      If B2SOn Then DOF 402,DOFOff
      flipOff 1
    End If
   End If

    If keycode = 203 then cLeft = 0' Left Arrow

    If keycode = 200 then cUp = 0' Up Arrow

    If keycode = 208 then cDown = 0' Down Arrow

    If keycode = 205 then cRight = 0' Right Arrow

    If keycode = 52 Then Zup = 0' Period

'************************Start Of Test Keys****************************
  If keycode = 30 Then  '"a"
  End if
'************************End Of Test Keys****************************
End Sub

Dim optionCount
Sub optionNoticeTimer_timer
  optionCount = optionCount + 1
  If optionCount > 2 Then
    optionNoticeTimer.Enabled = 0
    OptionNoticePrim.visible = 0
  End If
End Sub

'************** Table Boot
Dim backGlassOn
Dim bootCount:bootCount = 0
Sub bootTable_Timer
  bootCount = bootCount + 1
  If bootCount = 1 Then
    For each light in bipLights: light.state = 0: Next
    gameOverLight.state = 1
    tiltLight.state = 1
    If B2SOn Then
      '*************************** B2S Horse Method Choice ***********************************
      On Error Resume Next
      B2SVersion = controller.B2SBuildVersion
      If Err.Number <> 0 Then
        ' Handle the error
        tb.text = "Old version of B2S"
        B2SMethod = 0
        Err.Clear
      ElseIf B2SVersion > 20103.0 Then
        tb.text = B2SVersion
        B2SMethod = 1
      Else
        B2SMethod = 0
        tb.text = "Not newest version of B2S"
      End If
      'B2SMethod is set to 0 for exe multiple horses or 1 for B2SSetPos(id,xPos,yPos) single horses
      controller.B2SSetCredits Credit
      controller.B2SSetGameOver 35,1
      controller.B2SSetTilt 33,1
      controller.B2SSetBallInPlay 32,0
      controller.B2SSetData 10,1
      If Credit > 0 Then DOF 427, DOFOn
      If FreePlay = 1 Then DOF 427, DOFOn
    End If

    For Each light in bumperLights: light.state = 1: Next
    For x = 1 to maxPlayers
      If B2SOn Then controller.B2SSetScorePlayer x, score(x)
      sReel(x).setvalue(score(x))
    Next

    For each light in giLights: light.state = 1: Next
    backGlassOn = 1
    randomizeHosses
    me.enabled = False
  End If
End Sub


Dim posit(6), fSlot(6)
Sub randomizeHosses
  Dim k, j, hoss, hossVLM
  k = 0
  j = 5
  For Each hoss in hosses
    f = Int(Rnd * -700)
    For Each hossVLM in allHosses(k)
      hossVLM.transX = f
    Next
    posit(k) = Int(((hoss.transX/-800)*30) + 40 + (j*30))
    If B2SOn Then
      If B2SMethod = 0 Then
        controller.B2SSetData posit(k),1
      Else
        Select Case k:
          Case 0: controller.B2SSetData 190,1
              controller.B2SSetPos 190, 1639+((bgBM_hoss001.transX/800)*1429), 1336
          Case 1: controller.B2SSetData 160,1: controller.B2SSetPos 160, 1618+((bgBM_hoss002.transX/800)*1383), 1298
          Case 2: controller.B2SSetData 130,1: controller.B2SSetPos 130, 1602+((bgBM_hoss003.transX/800)*1351), 1275
          Case 3: controller.B2SSetData 100,1: controller.B2SSetPos 100, 1592+((bgBM_hoss004.transX/800)*1332), 1269
          Case 4: controller.B2SSetData 70,1: controller.B2SSetPos 70, 1566+((bgBM_hoss005.transX/800)*1279), 1210
          Case 5: controller.B2SSetData 40,1: controller.B2SSetPos 40, 1550+((bgBM_hoss006.transX/800)*1247), 1215
        End Select
      End If
    End If
    k = k+1
    j = j-1
  Next
End Sub

'***********Replay Settings
Sub replaySettings
  replay(0) = 600
  replay(1) = 700
  replay(2) = 800
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
Sub startGame
  If state = False Then
    raceOver = 0
    startUpLoop = 1
    scoreMotor.enabled = 1
    ballInPlay = 0
    If B2SOn Then
      controller.B2SSetCredits credit
      controller.B2SSetBallinPlay 32, ballInPlay
      controller.B2SSetGameOver 0
      controller.B2SSetTilt 0
    End If
    gameOverLight.state = 0
    tiltLight.state = 0
    RedHatReel.setValue(ballInPlay)
    dynamicUpdatePostIt.enabled = 0
    updatePostIt
    state = True
    gameState
    players = 1
    resetAllHorses
    resetReel.enabled = True
    If vrOption > 0 Then vrCredit
    resetGame.Enabled = 1
  End If
End Sub

Dim gameReady
Sub resetGame_Timer
  If resetOn = 0 Then
    newGame
    resetGame.Enabled = 0
  End If
End Sub

'*********New Game
Dim endGame: endgame = 1
Sub newGame
  player = 1
    endGame = 0
  gameState
  changeUnit
  newWin = 0
  For Each light in raceOverInLights: light.state = 0: Next
  For each light in bipLights: light.state = 0: Next
  For Each light in doorLights: light.state = 0: Next
  door1to3Reel.setValue(0)
  door4to6Reel.setValue(0)
  troughwall1.collidable = 0              ' Release balls in trough
  wallState = 0
  If tilt = False Then troughwall2.collidable = 0
  If tilt = True Then tilt = False
  tiltLight.state = 0
  yellowHatReel.setValue(0)
  If B2SOn Then
    controller.B2SSetTilt 33,0
    For x = 21 to 26
      controller.B2SSetData x, 0
    Next
    controller.B2SSetData 18, 0
  End If
End Sub

'**********Check if Game Should Continue
Dim relBall, rep(4)
Sub checkContinue
  If endGame = 1 Then
    turnOff
    flipOff 2
    state = False
    gameState
    yellowHatReel.setValue(1)
    DynamicUpdatePostIt.enabled = 1
    sortScores
    checkHighScores
    players = 0
    saveHighScore
    flippersOn = False
    If B2SOn Then controller.B2SSetGameOver 35,1
    gameOverLight.state = 1
  Else
    'relBall = 1    'variable to tell score motor routine that this is a ball release run
    'scoreMotor5.enabled = 1
  End If
End Sub

Dim ballOver4
Sub troughSwitch_Hit            ' this mimics the switch in the lower ball trough that must be active when an in-lane gate is hit to allow the playfield to energize
  ballOver4 = 1
End Sub

Dim ballInPlay, flippersOn, wallState
Sub gateSwitches_Hit            ' there is a gate switch activated by the top in-lanes that turns on the table and also increments the ball count unit
  troughwall1.collidable = 1
  wallState = 1

  If ballOver4 = 1 or resetOn = 0 Then flippersOn = True: ' if there is no ball over the trough switch or the table is resetting then it will not energize the playfield
  If flippersOn = True and raceOver = 0 Then
    ballInPlayLight.state = 1
    ballInPlay = ballInPlay + 1
    If raceOver = 0 Then RedHatReel.setValue(ballInPlay)
    If B2SOn Then
      controller.B2SSetBallinPlay 32, BallinPlay
      controller.B2SSetData 17,1
    End If
    For Each light in bipLights: light.state = 0: Next
    EVAL("bip00" & ballInPlay).state = 1
  End If
End Sub

'*************** Ball Lift
Sub ballLift
  liftpress = 1
  playFieldSound "BallLift1" , 0, Plunger, .7
  troughTimer.enabled = 1                         ' turn on the delay timer for advancing the balls in the lower kickers
  cabBM_b_lift.y = 2000
End Sub

troughtimer.interval = 10
Dim liftZ, liftBall, liftSpeed
liftSpeed = 10

Sub troughTimer_Timer
  If liftpress = 1 Then
    if not isEmpty(liftBall) Then
      liftZ = liftZ + liftSpeed
      If liftBall.z >= 50 Then
        liftBall.x = liftBall.x - liftSpeed/2
        If liftBall.x < 912 Then
          BallLiftLower.kickz 270, 5, 0, -25
          liftBall = Empty
        End If
      Else
        liftBall.z = liftBall.z + liftSpeed
      End If
    End If
  Else
    liftZ = liftZ - liftSpeed*2
    if not isEmpty(liftBall) Then
      If liftBall.z >= 50 and liftBall.x < BallLiftLower.x Then
        liftBall.x = liftBall.x + liftSpeed
        playFieldSound "BallLift2" , 0, Plunger, .7
      Else
        liftBall.z = liftBall.z - liftSpeed*2
      End If
    End If

    If liftZ <= 0 then
      liftz = 0
      If isEmpty(liftBall) Then
        troughwall2.collidable = 0
      Else
        liftBall.z = -135.68        'prevent z from getting offset from start point
      End If
      troughTimer.enabled = 0
    End If
  End If
End Sub

Sub BallLiftLower_hit         ' This is the kicker just under the banana track to the launch lane
  set liftBall = activeball
  troughwall2.collidable = 1
End Sub

Dim TH4, TH5

Sub TroughHit4_Hit
  TH4 = 1
  CheckGameOver
End Sub

Sub TroughHit4_UnHit
  TH4 = 0
End Sub

Sub TroughHit5_Hit
  TH5 = 1
  CheckGameOver
End Sub

Sub TroughHit5_UnHit
  TH5 = 0
End Sub

Sub CheckGameOver
  If TH4 = 1 and TH5 = 1 Then
    endGame = 1
    checkContinue
  End If
End Sub

'**********Check if Scoring Bonus is True
Sub scoreBonus
  advancePlayers
End Sub


'**********Advance Players
Sub advancePlayers
  nextBall
End Sub

'**********Next Ball
Sub nextBall
  tilt = False
  yellowHatReel.setValue(0)
  tiltLight.state = 0
  If B2SOn Then controller.B2SSetTilt 33,0
End Sub

'************Game State Check
  dim xx

Sub gameState
  If state = False Then
    yellowHatReel.setValue(1)
    If B2SOn then controller.B2SSetGameOver 35,1
    gameOverLight.state = 1
    flipOff 2
  Else
    yellowHatReel.setValue(0)
    If B2SOn Then
      controller.B2SSetTilt 33,0
      controller.B2SSetGameOver 35,0
    End If
    gameOverLight.state = 0
    tiltLight.state = 0
    yellowHatReel.setValue(0)
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


'******* for ball control script
Sub EndControl_Hit()
    contBallinPlay = False
  If tilt = True and tiltEndsGame = False Then nextBall
End Sub

'************** 10's Rubbers
Sub TensRubbers_Hit
  addScore 10
End Sub

'*************** Triggers
Sub triggers_hit(index)
  If tilt = False and flippersOn = True Then
    Select Case (index)
      Case 0: If UpperLight0.state = 0 Then
            addScore 10
          Else
            addScore 50
          End If
      Case 1:   runOdds: addScore 10
      Case 2:  selectorUnit
      Case 3:   runEvens: addScore 10
      Case 4  :If UpperLight1.state = 0 Then
            addScore 10
          Else
            addScore 50
          End If
      Case 5: horse1Flag = 1
          If motorOn = 0 Then scoreMotor.enabled = 1
          If horsesRunning = 0 Then startHorses
          checkHoss(1)
      Case 6: horse2Flag = 1
          If motorOn = 0 Then scoreMotor.enabled = 1
          If horsesRunning = 0 Then startHorses
          checkHoss(2)
      Case 7: horse3Flag = 1
          If motorOn = 0 Then scoreMotor.enabled = 1
          If horsesRunning = 0 Then startHorses
          checkHoss(3)
      Case 8: horse4Flag = 1
          If motorOn = 0 Then scoreMotor.enabled = 1
          If horsesRunning = 0 Then startHorses
          checkHoss(4)
      Case 9: horse5Flag = 1
          If motorOn = 0 Then scoreMotor.enabled = 1
          If horsesRunning = 0 Then startHorses
          checkHoss(5)
      Case 10:horse6Flag = 1
          If motorOn = 0 Then scoreMotor.enabled = 1
          If horsesRunning = 0 Then startHorses
          checkHoss(6)
    End Select
    If B2SOn = True Then DOF (405+index), DOFPulse
  End If
End Sub

Sub checkHoss(tHit)
  If  tHit = yourHoss Then
    addScore 50
  Else
    addScore 10
  End If
End Sub

'************** Wire RollOvers
Sub wireRollOvers_hit(index)

  wireNumber = index
  wireAnimation.enabled = 1
End Sub

'************** Wire RollOver Animation
Dim wire, wireNumber
Sub wireAnimation_Timer
  wire = wire + 1
  Select Case wire
    Case 1: EVAL ("pfBM_wire00" & wireNumber + 1).transz = -10
    Case 2: EVAL ("pfBM_wire00" & wireNumber + 1).transz = -4
    Case 3: EVAL ("pfBM_wire00" & wireNumber + 1).transz = -1
    Case 4: EVAL ("pfBM_wire00" & wireNumber + 1).transz = 0
        wire = 0
        wireAnimation.enabled = 0
  End Select
End Sub

Dim changeCount
Sub changeUnit                  ' the change unit is a stepper that toggles between the two 50 point lights at the top of the table
  changeCount = changeCount + 1
  If changeCount > 1 Then
    changeCount = 0
    UpperLight0.state = 1
    UpperLight1.state = 0
  Else
    UpperLight0.state = 0
    UpperLight1.state = 1
  End If
End Sub

' ************** Horse Race Section
Dim yourHoss, selectorCount
Sub selectorUnit                ' the selector unit is a stepper to determine the player's horse number
  If raceOver = 0 Then
    selectorCount = selectorCount + 1
    If selectorCount  > 11 Then selectorCount = 0
    Select Case selectorCount
      Case 0:  yourHoss = 1
      Case 1:  yourHoss = 4
      Case 2:  yourHoss = 6
      Case 3:  yourHoss = 5
      Case 4:  yourHoss = 2
      Case 5:  yourHoss = 3
      Case 6:  yourHoss = 6
      Case 7:  yourHoss = 1
      Case 8:  yourHoss = 3
      Case 9:  yourHoss = 4
      Case 10: yourHoss = 2
      Case 11: yourHoss = 5
    End Select
    If yourHoss < 4 Then
      horse1to3Reel.setValue(yourHoss)
      horse4to6Reel.setValue(0)
    Else
      horse1to3Reel.setValue(0)
      horse4to6Reel.setValue(yourHoss-3)
    End If
    For Each xx in yourHossLights: xx.state = 0: Next
    EVAL("h00" & yourHoss).state = 1
    If B2SOn Then
      For x =11 to 16
        controller.B2SSetData x,0
      Next
      controller.B2SSetData yourHoss+10, 1
    End If
    For Each light in lowerLaneLights: light.state = 0: Next
    For Each light in bumperLights: light.state = 0: Next
    EVAL("b00" & yourHoss).state = 1
    EVAL("l00" & yourHoss).state = 1
    EVAL("h00" & yourHoss).state = 1
  End If
End Sub

Dim hHit
Sub bumpHossScore(hHit)
  If ballControlBlock = True Then Exit Sub
  If hHit = yourHoss Then
    totalUp 10
  Else
    totalUp 1
  End If
  For Each xx in bumperLights: xx.state = 1: Next
  bumperFlicker.enabled = 1
End Sub

Dim flickerCount
Sub bumperFlicker_Timer
  flickerCount = flickerCount + 1
  If flickerCount > 1 Then
    For Each xx in bumperLights: xx.state = 0: Next
    EVAL("b00" & yourHoss).state = 1
    flickerCount = 0
    bumperFlicker.enabled = 0
  End If
End Sub

' ************** Run Horses
Dim raceOver, runHorseCount, gTimeStart, gTimeEnd, gTimeRun, gTimeMax, gTimeMin, gTimeSum, gTimeRunNum, gTimeAvg
Dim pPosition, hPos1, hPos2, hPos3, hPos4, hPos5, hPos6, horsesRunning
gTimeMin = 2000
Sub runHorses_Timer
  If horsesRunning = 0 Or raceOver = 1 Then
    gTimeEnd = gameTime
    gTimeRun = gTimeEnd - gTimeStart
    gTimeSum = gTimeSum + gTimeRun
    gTimeRunNum = gTimeRunNum + 1
    If gTimeRun > gTimeMax Then gTimeMax = gTimeRun
    If gTimeRun < gTimeMin Then gTimeMin = gTimeRun
    gTimeAvg = Int(gTimeSum/gTimeRunNum)
    'tb.text = "Time: " & gTimeRun & " Max: " & gTimeMax & " Min: " & gTimeMin & " Avg: " & gTimeAvg
    runHorses.Enabled = 0
    stopSound "ACHorseMotor"
    Exit Sub
  End If
  If horse1Flag = 1 Then
    For Each xx in BP_bghoss001: xx.transX = xx.transX - 4: Next
    hPos1 = Int((bgBM_hoss001.transX/-800)*30) + 190
    If hPos1 > 219 Then hPos1 = 219
  End If
  If horse2Flag = 1 Then
    For Each xx in BP_bghoss002: xx.transX = xx.transX - 4: Next
    hPos2 = Int((bgBM_hoss002.transX/-800)*30) + 160
    If hPos2 > 189 Then hPos2 = 189
  End If
  If horse3Flag = 1 Then
    For Each xx in BP_bghoss003: xx.transX = xx.transX - 4: Next
    hPos3 = Int((bgBM_hoss003.transX/-800)*30) + 130
    If hPos3 > 159 Then hPos4 = 159
  End If
  If horse4Flag = 1 Then
    For Each xx in BP_bghoss004: xx.transX = xx.transX - 4: Next
    hPos4 = Int((bgBM_hoss004.transX/-800)*30) + 100
    If hPos4 > 129 Then hPos4 = 129
  End If
  If horse5Flag = 1 Then
    For Each xx in BP_bghoss005: xx.transX = xx.transX - 4: Next
    hPos5 = Int((bgBM_hoss005.transX/-800)*30) + 70
    If hPos5 > 99 Then hPos5 = 99
  End If
  If horse6Flag = 1 Then
    For Each xx in BP_bghoss006: xx.transX = xx.transX - 4: Next
    hPos6 = Int((bgBM_hoss006.transX/-800)*30) + 40
    If hPos6 > 69 Then hPos6 = 69
  End If
  If B2SOn Then updateB2SHorses 1
  checkWinner
  runHorseCount = runHorseCount + 1
End Sub

Dim replayFlag, checkWinnerCount, newWin
replayFlag = 0
Sub checkWinner
  checkWinnerCount = checkWinnerCount + 1
  Dim horseNumber, winner
  winner = 0
  horseNumber = 0
  For Each xx in horsePrims
    horseNumber = horseNumber + 1
    If xx.transX < -800 Then
      raceOver = 1
      winner = horseNumber
      If B2SOn Then controller.B2SSetData 17,0
    End If
  Next
  If raceOver = 1 Then
    ballInPlayLight.state = 0
    redHatReel.setValue(ballInPlay+5)
    If B2SOn Then
      controller.B2SSetData 20+winner, 1
      controller.B2SSetData 18, 1
      If chime = 1 Then DOF 442, DOFPulse
    End If
    Eval("d00" & winner).state = 1
    For Each light in raceOverInLights: light.state = 1: Next
    If chime = 0 and newWin = 0 Then playFieldSound "BellWin", 0, soundPoint13, 1
    newWin = 1
    If winner < 4 Then
      door1to3Reel.setValue(winner)
    Else
      door4to6Reel.setValue(winner - 3)
    End If
  End If
  If winner = yourHoss Then
    If ballInPlay = 1 Then replayFlag = 10
    If ballInPlay = 2 Then replayFlag = 5
    If ballInPlay = 3 Then replayFlag = 2
    If ballInPlay = 4 Then replayFlag = 1
    If ballInPlay = 5 Then replayFlag = 1
    scoreMotor.enabled = 1
  End If
End Sub


Sub runOdds
  horse1Flag = 1
  horse3Flag = 1
  horse5Flag = 1
  If motorOn = 0 Then scoreMotor.enabled = 1
  If horsesRunning = 0 Then startHorses
End Sub

Sub runEvens
  horse2Flag = 1
  horse4Flag = 1
  horse6Flag = 1
  If motorOn = 0 Then scoreMotor.enabled = 1
  If horsesRunning = 0 Then startHorses
End Sub

Sub startHorses
  gTimeStart = gameTime
  horsesRunning = 1
  runHorses.Enabled = 1
  playSound "ACHorseMotor", -1, 0.001, 0.12, 0, 0, 1, 1, 0
End Sub

'******* Return Horses
' Note these timers are all a part of the returnHorses collection and get turned on that way
' also Bord's horses are not set at the start of the race track so a transX of 285 is the start of the track

Dim resetOn


Dim horseRun, horse1Flag, horse2Flag, horse3Flag, horse4Flag, horse5Flag, horse6Flag
Sub resetAllHorses
  horse1Flag = 1
  horse2Flag = 1
  horse3Flag = 1
  horse4Flag = 1
  horse5Flag = 1
  horse6Flag = 1
  horsesReset = False
  returnHorses.Enabled = 1
  playSound "ACHorseMotor", -1, 0.001, 0.12, 0, 0, 1, 1, 0
End Sub

Dim horsesReset
Sub returnHorses_timer
  If horse1Flag = 0 And horse2Flag = 0 And horse3Flag = 0 And horse4Flag = 0 and horse5Flag = 0 And horse6Flag = 0 Then
    horsesReset = True
    stopSound "ACHorseMotor"
    returnHorses.Enabled = 0
  End If
  If horse1Flag =  1 Then
    For Each xx in BP_bghoss001: xx.transX = xx.transX + 5: Next
    hPos1 = Int((bgBM_hoss001.transX/-800)*30) + 190
    If hPos1 < 191 Then
      horse1Flag = 0
      hPos1 = 191
    End If
  End If
  If horse2Flag =  1 Then
    For Each xx in BP_bghoss002: xx.transX = xx.transX + 5: Next
    hPos2 = Int((bgBM_hoss002.transX/-800)*30) + 160
    If hPos2 < 161 Then
      horse2Flag = 0
      hPos2 = 161
    End If
  End If
  If horse3Flag =  1 Then
    For Each xx in BP_bghoss003: xx.transX = xx.transX + 5: Next
    hPos3 = Int((bgBM_hoss003.transX/-800)*30) + 130
    If hPos3 < 131 Then
      horse3Flag = 0
      hPos3 = 131
    End If
  End If
  If horse4Flag =  1 Then
    For Each xx in BP_bghoss004: xx.transX = xx.transX + 5: Next
    hPos4 = Int((bgBM_hoss004.transX/-800)*30) + 100
    If hPos4 < 101 Then
      horse4Flag = 0
      hPos4 = 101
    End If
  End If
  If horse5Flag =  1 Then
    For Each xx in BP_bghoss005: xx.transX = xx.transX + 5: Next
    hPos5 = Int((bgBM_hoss005.transX/-800)*30) + 70
    If hPos5 < 71 Then
      horse5Flag = 0
      hPos5 = 71
    End If
  End If
  If horse6Flag =  1 Then
    For Each xx in BP_bghoss006: xx.transX = xx.transX + 5: Next
    hPos6 = Int((bgBM_hoss006.transX/-800)*30) + 40
    If hPos6 < 41 Then
      horse6Flag = 0
      hPos6 = 40
    End If
  End If

  If B2SOn Then upDateB2SHorses 0

End Sub




Dim horse, direction
Sub updateB2SHorses(direction)
  If B2SMethod = 0 Then
    If direction = 0 Then
      controller.B2SSetData hPos1+1, 0
      controller.B2SSetData hPos1, 1
      controller.B2SSetData hPos2+1, 0
      controller.B2SSetData hPos2, 1
      controller.B2SSetData hPos3+1, 0
      controller.B2SSetData hPos3, 1
      controller.B2SSetData hPos4+1, 0
      controller.B2SSetData hPos4, 1
      controller.B2SSetData hPos5+1, 0
      controller.B2SSetData hPos5, 1
      controller.B2SSetData hPos6+1, 0
      controller.B2SSetData hPos6, 1
    Else
      controller.B2SSetData hPos1-1, 0
      controller.B2SSetData hPos1, 1
      controller.B2SSetData hPos2-1, 0
      controller.B2SSetData hPos2, 1
      controller.B2SSetData hPos3-1, 0
      controller.B2SSetData hPos3, 1
      controller.B2SSetData hPos4-1, 0
      controller.B2SSetData hPos4, 1
      controller.B2SSetData hPos5-1, 0
      controller.B2SSetData hPos5, 1
      controller.B2SSetData hPos6-1, 0
      controller.B2SSetData hPos6, 1
    End If
  Else
    tb.text = "in reset"
    controller.B2SSetPos 190, 1639+((bgBM_hoss001.transX/800)*1429), 1336
    controller.B2SSetPos 160, 1618+((bgBM_hoss002.transX/800)*1383), 1298
    controller.B2SSetPos 130, 1602+((bgBM_hoss003.transX/800)*1351), 1275
    controller.B2SSetPos 100, 1592+((bgBM_hoss004.transX/800)*1332), 1269
    controller.B2SSetPos 70, 1566+((bgBM_hoss005.transX/800)*1279), 1210
    controller.B2SSetPos 40, 1550+((bgBM_hoss006.transX/800)*1247), 1215
  End If
End Sub

'**************Special
Sub special
  increaseCredits
  knock
  If sideSpecial = 1 Then
    specialReset
    sideSpecial = 0
  End If
End Sub

Sub knock
  If B2SOn = True Then DOF 428, DOFPulse Else playSound "knocker"
End Sub


'***************Score Motor Run one full rotation
Dim scoreMotorCount, addCredit
Sub scoremotor5_Timer
  scoreMotorCount = scoreMotorCount + 1
  playFieldSound "ScoreMotorSingleFire", 0, SoundPointScoreMotor, .02 'need to set location of score motor under the PF
  If scoremotorCount = 5 Then
    If addCredit = 1 Then
      addCredit = 0
      increaseCredits
    End If
    scoreMotorCount = 0
    scoreMotor5.enabled = 0
  End If
End Sub

Sub increaseCredits
  credit = credit + 1
  If showDT = False Then PlayReelSound "Reel5", bgcrReel Else PlayReelSound "Reel5", dtcrReel
  If B2SOn Then DOF 427, DOFOn
  If credit > 15 then credit = 15
  CreditReel.setValue (credit)
  If vrOption > 0 Then vrCredit
  If B2SOn Then
    controller.B2SSetCredits credit
    If credit > 0 Then DOF 427, DOFOn
  End If
End Sub

'**************Score Motor Routine
Dim bellRing, scoreMotorLoop,  motorOn, startUpLoop
Sub scoreMotor_timer
  scoreMotorLoop = scoreMotorLoop + 1

  If scoreMotorLoop < 6 Then
    playFieldSound "ScoreMotorSingleFire", 0, SoundPointScoreMotor, .02 '0.2
    motorOn = 1
  Else
    motorOn = 0
  End If

  If startUpLoop > 0  And scoreMotorLoop < 6 Then
    selectorUnit
  End If

'  These Flags are passed by scores with multiple of 10, 100 or 1000
  If flag = 1 or flag10 = 1 or flag100 = 1 Then
    Select Case scoreMotorLoop
      Case 1: totalUp Point:
      Case 2: If bellRing > 1 Then totalUp point
      Case 3: If bellRing > 2 Then totalUp Point: 'If point = 1000 Then advanceBonus
      Case 4: If bellRing > 3 Then totalUp point
      Case 5: changeUnit: If bellRing > 4 Then totalUp point: 'If point = 1000 Then advanceBonus
      Case 6: flag1 = 0
          flag10 = 0
          flag100 = 0
          If replayFlag < 1 And startUpLoop < 1 Then
            motorOn = 0
            scoreMotorLoop = 0
            scoreMotor.enabled = 0
            horse1Flag = 0
            horse2Flag = 0
            horse3Flag = 0
            horse4Flag = 0
            horse5Flag = 0
            horse6Flag = 0
            horsesRunning = 0
          End If
    End Select
  Else
    If scoreMotorLoop > 5 Then
      startUpLoop = startUpLoop - 1
      If replayFlag < 1 And startUpLoop < 1 And horsesReset = True Then
        motorOn = 0
        scoreMotorLoop = 0
        horse1Flag = 0
        horse2Flag = 0
        horse3Flag = 0
        horse4Flag = 0
        horse5Flag = 0
        horse6Flag = 0
        horsesRunning = 0
        scoreMotor.enabled = 0
      Else
        scoreMotorLoop = 0
      End If
    End If
  End If

  If replayFlag > 0 Then
    Select Case scoreMotorLoop
      Case 1: If replayFlag > 0 Then increaseCredits: replayFlag = replayFlag - 1
      Case 2: If replayFlag > 1 Then increaseCredits: replayFlag = replayFlag - 1
      Case 3: If replayFlag > 2 Then increaseCredits: replayFlag = replayFlag - 1
      Case 4: If replayFlag > 3 Then increaseCredits: replayFlag = replayFlag - 1
      Case 5: If replayFlag > 4 Then increaseCredits: replayFlag = replayFlag - 1
      Case 6: If replayFlag < 1 Then
            motorOn = 0
            scoreMotor.enabled = 0
            scoreMotorLoop = 0
           Else
            scoreMotorLoop = 0
           End If
    End Select
  End If

End Sub

'***************Scoring Routine
Dim flag1, flag10, flag100, flag1000, point, point10
Sub pts10
  If pts10Timer.enabled = False Then
    pts10Timer.enabled = True
    playFieldSound "Bell10", 0, soundPoint13, 1
  End If
End Sub

Sub pts100
  If pts100Timer.enabled = False Then
    pts100Timer.enabled = True
    playFieldSound "Bell100", 0, soundPoint13, 1
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
  If tilt = False Then
    If points > 9 and points <100 Then
      point10 =  (points / 10)

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

    If Points > 999 Then
      If points > 1000 Then
        bellRing = (points / 1000)
        point = 1000: flag1000 = 1: scoreMotor.enabled = 1
      Else
        If motorOn = 0 Then totalUp 1000
      End If
    End If
  End If
End Sub

Dim replayX, replay(7), repAwarded(5), block
Sub totalUp(points)
  If tilt = True Then Exit Sub
  block = 0

  If B2SOn And showDT = False Then
    If Player = 1 and block = 0 Then PlayReelSound "Reel1", bgs1Reel
    If Player = 2 and block = 0 Then PlayReelSound "Reel2", bgs2Reel
    If Player = 3 and block = 0 Then PlayReelSound "Reel3", bgs3Reel
    If Player = 4 and block = 0 Then PlayReelSound "Reel4", bgs4Reel
  End If

  If showDT = True Then
    If Player = 1 and block = 0 Then PlayReelSound "Reel1", dts1Reel
    If Player = 2 and block = 0 Then PlayReelSound "Reel2", dts2Reel
    If Player = 3 and block = 0 Then PlayReelSound "Reel3", dts3Reel
    If Player = 4 and block = 0 Then PlayReelSound "Reel4", dts4Reel
  End If

  If flag10 = 1 or points = 10 Then
    If chime = 0 Then
      pts10
    Else
      If B2SOn Then DOF 441,DOFPulse
    End If
  End If

  If flag100 = 1 or points = 100 Then
    If chime = 0 Then
      pts100
    Else
      If B2SOn Then DOF 441,DOFPulse
    End If
  End If

  If block = 0 Then score(Player) = score(player) + points
  If block = 0 Then sReel(Player).addvalue(points)
  If vrOption > 0 Then vrScore
  vrScore

  If B2SOn Then controller.B2SSetScorePlayer player, score(player)

  If score(player) > replay(rep(player)) or score(player) = replay(rep(player)) Then
    If rep(player) > 3 Then Exit Sub
    increaseCredits
    rep(player) = rep(player) + 1
    knock
  End If
End Sub

'***************Tilt
Dim tiltSens

Sub checkTilt
  If tilttimer.enabled = True Then
    tiltSens = tiltSens + 1
    If tiltSens = 3 Then
      tilt = True
      FlippersOn = False
      yellowHatReel.setValue(2)
      tiltLight.state = 1
      If B2SOn Then controller.B2SSetTilt 33,1
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

  If score(1) = 0 and score(2) = 0 and score(3) = 0 and score(4) = 0  Then
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
    For x = 1 to maxPlayers
      If score(x) = 0 Then reelDone(x) = 1: resetOn = 0
    Next
    If B2SOn Then
      controller.B2SSetScorePlayer playerTest, score(playerTest)
      sReel(playerTest).setvalue (score(playerTest))
    Else
      sReel(playerTest).setvalue (score(playerTest))
    End If

    If reelDone(playerTest) = 0 and reelStop = 0 Then
      resetOn = 1
      If showDT = False Then
        If reelDone(1) = 0 and playerTest = 1 Then PlayReelSound "Reel1", bgs1Reel
        If reelDone(2) = 0 and playerTest = 2 Then PlayReelSound "Reel2", bgs2Reel
        If reelDone(3) = 0 and playerTest = 3 Then PlayReelSound "Reel3", bgs3Reel
        If reelDone(4) = 0 and playerTest = 4 Then PlayReelSound "Reel4", bgs4Reel
      Else
        If reelDone(1) = 0 and playerTest = 1 Then PlayReelSound "Reel1", dts1Reel
        If reelDone(2) = 0 and playerTest = 2 Then PlayReelSound "Reel2", dts2Reel
        If reelDone(3) = 0 and playerTest = 3 Then PlayReelSound "Reel3", dts3Reel
        If reelDone(4) = 0 and playerTest = 4 Then PlayReelSound "Reel4", dts4Reel
      End If
    End If

    If score(playerTest) = 0 Then reelDone(playerTest) = 1
    If vrOption > 0 Then
      vrScore
    End If
    vrScore
  Next
  If reelFlag = 1 Then reelStop = 1
End Sub

'This Timer runs a loop that calls the CountUp and UpdateReels routines to step the reels up five times and Then
'check to see if they are all at zero during a two loop pause and then step them the rest of the way to zero

Dim testFlag
Sub resetReel_Timer
  For x = 1 to maxPlayers
    score(x) = (score(x) Mod 100000)
  Next
  resetLoop = resetLoop + 1
  If resetLoop = 1 and score(1) = 0 and score(2) = 0 and score(3) = 0 and score(4) = 0 Then
    resetLoop = 0
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

'************************************************** Set Table LUT Value *****************************************************
Sub setLUT
  lut.enabled = 1
  Table1.ColorGradeImage = "LUT" & lutValue
End Sub

Sub lut_Timer
  lutText.visible = 0
  lut.enabled = 0
  saveHighScore
End Sub

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

    For x = 1 to 31
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
    chime = CInt(Right(temp(28),1))
    pfOption = CInt(Right(temp(29),1))
    lutValue = CInt(Right (temp(30),1))
    If vrOption > 0 Then
      vrOption = CInt(Right (temp(31),1))
    End If
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
    scoreFile.WriteLine "Chime (0 = sound file, 1 = DOF chime): " & chime
    scoreFile.WriteLine "pfOption (1 = L/R Stereo, 2 = U/D Stereo, 3 = quad): " & pfOption
    scoreFile.WriteLine "LUT: " & lutValue
    If vrOption > 0 Then
      scoreFile.WriteLine "vrOption: " & vrOption
    Else
      scoreFile.WriteLine "vrOption: 2"
    End If
    scoreFile.WriteLine "----- Log Section -----"
    scoreFile.WriteLine "B2SVersion: " & B2SVersion
    scoreFile.WriteLine "B2SMethod: " & B2SMethod
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
    leftFlipper.RotateToStart
  flipOff 2
  If B2SOn Then DOF 401, DOFOff
  rightFlipper.RotateToStart
  If B2SOn Then DOF 402, DOFOff
End Sub

Dim side
Sub flipOff (side)
  If side = 0 or side = 2 Then
      stopSound "FlipBuzzLA"
      stopSound "FlipBuzzLB"
      stopSound "FlipBuzzLC"
      stopSound "FlipBuzzLD"
  End If
  If side = 1 or side = 2 Then
      stopSound "FlipBuzzRA"
      stopSound "FlipBuzzRB"
      stopSound "FlipBuzzRC"
      stopSound "FlipBuzzRD"
  End If
End Sub

' ************************************* VR Reels *****************************************
Sub vrCredit
  vrCreditReel.objRotX = credit * 18
End Sub

Dim vrOne, vrTen, vrHundred, reelUp
Sub vrScore
  vrOne =   score(1) Mod 10
  vrTen =   (INT (score(1)/10) )Mod 10
  vrHundred = (INT(score(1)/100)) Mod 10

  vrScoreReel_1.objRotX = vrOne * 36
  vrScoreReel_10.objRotX = vrTen * 36
  vrScoreReel_100.objRotX = vrHundred * 36

End Sub


'*******************************************
' VR Room / VR Cabinet
'*******************************************

DIM VRThings
If vrOption = 1 Then
' for each VRThings in vr: VRThings.visible = 1: Next
' for each VRThings in vrCab: VRThings.visible = 1: Next
' for each VRThings in vrRoom: VRThings.visible = 0: Next
Elseif vrOption = 2 Then
' for each VRThings in vrCab: VRThings.visible = 1: Next
' for each VRThings in vrRoom: VRThings.visible = 1: Next
Else
' for each VRThings in vr: VRThings.visible = 0: Next
' for each VRThings in vrCab: VRThings.visible = 0: Next
' for each VRThings in vrRoom: VRThings.visible = 0: Next
  tape.image = "tapeCab"
End if

'If vrOption > 0 and VRGlassScratches = 1 then GlassImpurities.visible = true

'*****************************************************Supporting Code Written By Others*************************************

'*********************************************
'  VR Plunger Code From Flash by Bord and Roth
'*********************************************

Sub TimerVRPlunger_Timer
  If cabBM_b_plunge.Y < 2210 then
       cabBM_b_plunge.Y = cabBM_b_plunge.Y + 5
  End If
End Sub

Sub TimerVRPlunger1_Timer
  cabBM_b_plunge.Y = 2128 + (5* Plunger.Position) -20
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

Const tnob =  5' total number of balls

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
  Dim b, paSub
  paSub=35000 'Playsound pitch adder for subway rolling ball sound

  ' play the rolling sound for each ball
  For b = 0 to 4  'UBound(BOT)

  ' Ball Rolling sounds
  '**********************
  ' Ball=50 units=1.0625".  One unit = 0.02125"  Ball.z is ball center.
  ' A ball in Adrian's saucer has a Z of -3.  Use <-5 for subway sounds.
    If PFOption = 1 or PFOption = 2 Then
      If BallVel(gBOT(b)) > 1 AND gBOT(b).z > 10 and gBOT(b).z <26 Then 'Ball on playfield
        rolling(b) = True
        PlaySound("BallrollingA" & b), -1, Vol(gBOT(b)) * 0.2 * xGain(gBOT(b)), AudioPan(gBOT(b)), 0, Pitch(gBOT(b)), 1, 0, 0 'Left & Right stereo or Top & Bottom stereo PF Speakers.
      ElseIf BallVel(gBOT(b)) > 1 AND gBOT(b).z < -5 Then 'Ball on subway
        PlaySound("BallrollingA" & b), -1, Vol(gBOT(b)) * 0.2 * xGain(gBOT(b)), AudioPan(gBOT(b)), 0, Pitch(gBOT(b))+paSub, 1, 0, 0 'Left & Right stereo or Top & Bottom stereo PF Speakers.
        rolling(b) = True
      ElseIf rolling(b) = True Then
        StopSound("BallrollingA" & b)
        rolling(b) = False
      End If
    End If

    If PFOption = 3 Then
      If BallVel(gBOT(b)) > 1 AND gBOT(b).z > 10 and gBOT(b).z < 26 Then  'Ball on playfield
        rolling(b) = True
        PlaySound("BallrollingA" & b), -1, Vol(gBOT(b)) * 0.2 *    XVol(gBOT(b))  *     YVol(gBOT(b)), -1, 0, Pitch(gBOT(b)), 1, 0, -1  'Top Left PF Speaker
        PlaySound("BallrollingB" & b), -1, Vol(gBOT(b)) * 0.2 * (1-XVol(gBOT(b))) *     YVol(gBOT(b)),  1, 0, Pitch(gBOT(b)), 1, 0, -1  'Top Right PF Speaker
        PlaySound("BallrollingC" & b), -1, Vol(gBOT(b)) * 0.2 *    XVol(gBOT(b))  * (1-YVol(gBOT(b))), -1, 0, Pitch(gBOT(b)), 1, 0,  1  'Bottom Left PF Speaker
        PlaySound("BallrollingD" & b), -1, Vol(gBOT(b)) * 0.2 * (1-XVol(gBOT(b))) * (1-YVol(gBOT(b))),  1, 0, Pitch(gBOT(b)), 1, 0,  1  'Bottom Right PF Speaker
      ElseIf BallVel(gBOT(b)) > 1 AND gBOT(b).z < -5 Then 'Ball on subway
        rolling(b) = True
        PlaySound("BallrollingA" & b), -1, Vol(gBOT(b)) * 0.2 *    XVol(gBOT(b))  *     YVol(gBOT(b)), -1, 0, Pitch(gBOT(b))+paSub, 1, 0, -1  'Top Left PF Speaker
        PlaySound("BallrollingB" & b), -1, Vol(gBOT(b)) * 0.2 * (1-XVol(gBOT(b))) *     YVol(gBOT(b)),  1, 0, Pitch(gBOT(b))+paSub, 1, 0, -1  'Top Right PF Speaker
        PlaySound("BallrollingC" & b), -1, Vol(gBOT(b)) * 0.2 *    XVol(gBOT(b))  * (1-YVol(gBOT(b))), -1, 0, Pitch(gBOT(b))+paSub, 1, 0,  1  'Bottom Left PF Speaker
        PlaySound("BallrollingD" & b), -1, Vol(gBOT(b)) * 0.2 * (1-XVol(gBOT(b))) * (1-YVol(gBOT(b))),  1, 0, Pitch(gBOT(b))+paSub, 1, 0,  1  'Bottom Right PF Speaker
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
      If gBOT(b).VelZ > -4 And gBOT(b).VelZ < -2 And gBOT(b).Z > 27 Then
        PlaySound "BallDrop1" & b, 0, ABS(gBOT(b).VelZ)/600 * xGain(gBOT(b)), AudioPan(gBOT(b)), 0, Pitch(gBOT(b)), 0, 0, 0 'Left & Right stereo or Top & Bottom stereo PF Speakers.
      ElseIf gBOT(b).VelZ > -8 And gBOT(b).VelZ < -4 And gBOT(b).Z > 27  Then
        PlaySound "BallDrop2" & b, 0, ABS(gBOT(b).VelZ)/600 * xGain(gBOT(b)), AudioPan(gBOT(b)), 0, Pitch(gBOT(b)), 0, 0, 0 'Left & Right stereo or Top & Bottom stereo PF Speakers.
      ElseIf gBOT(b).VelZ > -12 And gBOT(b).VelZ < -8 And gBOT(b).Z > 27  Then
        PlaySound "BallDrop3" & b, 0, ABS(gBOT(b).VelZ)/600 * xGain(gBOT(b)), AudioPan(gBOT(b)), 0, Pitch(gBOT(b)), 0, 0, 0 'Left & Right stereo or Top & Bottom stereo PF Speakers.
      ElseIf gBOT(b).VelZ < -12 And gBOT(b).Z > 27  Then
        PlaySound "BallDrop4" & b, 0, ABS(gBOT(b).VelZ)/600 * xGain(gBOT(b)), AudioPan(gBOT(b)), 0, Pitch(gBOT(b)), 0, 0, 0 'Left & Right stereo or Top & Bottom stereo PF Speakers.
      End If
    End If

    If PFOption = 3 Then
      If gBOT(b).VelZ > -4 And gBOT(b).VelZ < -2 And gBOT(b).Z > 27  Then
        PlaySound "BallDrop1" & b, 0, ABS(gBOT(b).VelZ)/600 *    XVol(gBOT(b))  *     YVol(gBOT(b)), -1, 0, Pitch(gBOT(b)), 1, 0, -1  'Top Left PF Speaker
        PlaySound "BallDrop1" & b, 0, ABS(gBOT(b).VelZ)/600 * (1-XVol(gBOT(b))) *     YVol(gBOT(b)),  1, 0, Pitch(gBOT(b)), 1, 0, -1  'Top Right PF Speaker
        PlaySound "BallDrop1" & b, 0, ABS(gBOT(b).VelZ)/600 *    XVol(gBOT(b))  * (1-YVol(gBOT(b))), -1, 0, Pitch(gBOT(b)), 1, 0,  1  'Bottom Left PF Speaker
        PlaySound "BallDrop1" & b, 0, ABS(gBOT(b).VelZ)/600 * (1-XVol(gBOT(b))) * (1-YVol(gBOT(b))),  1, 0, Pitch(gBOT(b)), 1, 0,  1  'Bottom Right PF Speaker
      ElseIf gBOT(b).VelZ > -8 And gBOT(b).VelZ < -4 And gBOT(b).Z > 27  Then
        PlaySound "BallDrop2" & b, 0, ABS(gBOT(b).VelZ)/600 *    XVol(gBOT(b))  *     YVol(gBOT(b)), -1, 0, Pitch(gBOT(b)), 1, 0, -1  'Top Left PF Speaker
        PlaySound "BallDrop2" & b, 0, ABS(gBOT(b).VelZ)/600 * (1-XVol(gBOT(b))) *     YVol(gBOT(b)),  1, 0, Pitch(gBOT(b)), 1, 0, -1  'Top Right PF Speaker
        PlaySound "BallDrop2" & b, 0, ABS(gBOT(b).VelZ)/600 *    XVol(gBOT(b))  * (1-YVol(gBOT(b))), -1, 0, Pitch(gBOT(b)), 1, 0,  1  'Bottom Left PF Speaker
        PlaySound "BallDrop2" & b, 0, ABS(gBOT(b).VelZ)/600 * (1-XVol(gBOT(b))) * (1-YVol(gBOT(b))),  1, 0, Pitch(gBOT(b)), 1, 0,  1  'Bottom Right PF Speaker
      ElseIf gBOT(b).VelZ > -12 And gBOT(b).VelZ < -8 And gBOT(b).Z > 27  Then
        PlaySound "BallDrop3" & b, 0, ABS(gBOT(b).VelZ)/600 *    XVol(gBOT(b))  *     YVol(gBOT(b)), -1, 0, Pitch(gBOT(b)), 1, 0, -1  'Top Left PF Speaker
        PlaySound "BallDrop3" & b, 0, ABS(gBOT(b).VelZ)/600 * (1-XVol(gBOT(b))) *     YVol(gBOT(b)),  1, 0, Pitch(gBOT(b)), 1, 0, -1  'Top Right PF Speaker
        PlaySound "BallDrop3" & b, 0, ABS(gBOT(b).VelZ)/600 *    XVol(gBOT(b))  * (1-YVol(gBOT(b))), -1, 0, Pitch(gBOT(b)), 1, 0,  1  'Bottom Left PF Speaker
        PlaySound "BallDrop3" & b, 0, ABS(gBOT(b).VelZ)/600 * (1-XVol(gBOT(b))) * (1-YVol(gBOT(b))),  1, 0, Pitch(gBOT(b)), 1, 0,  1  'Bottom Right PF Speaker
      ElseIf gBOT(b).VelZ < -12 And gBOT(b).Z > 27  Then
        PlaySound "BallDrop4" & b, 0, ABS(gBOT(b).VelZ)/600 *    XVol(gBOT(b))  *     YVol(gBOT(b)), -1, 0, Pitch(gBOT(b)), 1, 0, -1  'Top Left PF Speaker
        PlaySound "BallDrop4" & b, 0, ABS(gBOT(b).VelZ)/600 * (1-XVol(gBOT(b))) *     YVol(gBOT(b)),  1, 0, Pitch(gBOT(b)), 1, 0, -1  'Top Right PF Speaker
        PlaySound "BallDrop4" & b, 0, ABS(gBOT(b).VelZ)/600 *    XVol(gBOT(b))  * (1-YVol(gBOT(b))), -1, 0, Pitch(gBOT(b)), 1, 0,  1  'Bottom Left PF Speaker
        PlaySound "BallDrop4" & b, 0, ABS(gBOT(b).VelZ)/600 * (1-XVol(gBOT(b))) * (1-YVol(gBOT(b))),  1, 0, Pitch(gBOT(b)), 1, 0,  1  'Bottom Right PF Speaker
      End If
    End If

  ' Glass hit sounds
  '*******************
  ' Ball=50 units=1.0625".  Ball.z is ball center.  Balls are physically limited by the table's Top Glass Height property which is always parallel to the playfield
  ' To ensure a ball can go high enough to trigger a glass hit the max ball.z is 25 units below Top Glass Height-5
  ' Change GHB below if the glass is not parallel to the playfield

    Dim GHT, GHB, PFL
    GHT = (table1.glassheight-5)*1.0625/50  'Glass height at top of real playfield in inches
    GHB = (table1.glassheight-5)*1.0625/50  'Glass height at bottom of real playfield in inches
    PFL = table1.height*1.0625/50   'Length of real playfield in inches

    If PFOption = 1 or PFOption = 2 Then
      If gBOT(b).Z > (gBOT(b).Y * ((GHT-GHB)/PFL)) + (GHB*50/1.0625) - BallSize/2  Then
        PlaySound "GlassHit" & b, 0, ABS(gBOT(b).VelZ)/30 * xGain(gBOT(b)), AudioPan(gBOT(b)), 0, Pitch(gBOT(b)), 0, 0, 0 'Left & Right stereo or Top & Bottom stereo PF Speakers.
      End If
    End If

    If PFOption = 3 Then
      If gBOT(b).Z > (gBOT(b).Y * ((GHT-GHB)/PFL)) + (GHB*50/1.0625) - Ballsize/2  Then
        PlaySound "GlassHit" & b, 0, ABS(gBOT(b).VelZ)/30 *    XVol(gBOT(b))  *     YVol(gBOT(b)), -1, 0, Pitch(gBOT(b)), 0, 0, -1  'Top Left PF Speaker
        PlaySound "GlassHit" & b, 0, ABS(gBOT(b).VelZ)/30 * (1-XVol(gBOT(b))) *     YVol(gBOT(b)),  1, 0, Pitch(gBOT(b)), 0, 0, -1  'Top Right PF Speaker
        PlaySound "GlassHit" & b, 0, ABS(gBOT(b).VelZ)/30 *    XVol(gBOT(b))  * (1-YVol(gBOT(b))), -1, 0, Pitch(gBOT(b)), 0, 0,  1  'Bottom Left PF Speaker
        PlaySound "GlassHit" & b, 0, ABS(gBOT(b).VelZ)/30 * (1-XVol(gBOT(b))) * (1-YVol(gBOT(b))),  1, 0, Pitch(gBOT(b)), 0, 0,  1  'Bottom Right PF Speaker
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

Sub aRubberWheel_hit(idx)
  PlayFieldSoundAB "rubberWheel", 0, 1
End Sub

Sub aRubberPosts_Hit(idx)
  PlayFieldSoundAB "rubberPost", 0, 1
End Sub

Sub aWoods_Hit(idx)
  If activeBall.Y > 550 Then
    PlayFieldSoundAB "wood", 0, 1
  End If
End Sub

Sub phys_plastic_hit
  PlayFieldSoundAB "fx_plastichit", 0, 0.01
End Sub

Sub LeftFlipper_collide(parm)
  RandomSoundFlipper()
End Sub

Sub LeftFlipper1_collide(parm)
  RandomSoundFlipper()
End Sub

Sub LeftFlipper2_collide(parm)
  RandomSoundFlipper()
End Sub

Sub RightFlipper_collide(parm)
  RandomSoundFlipper()
End Sub

Sub RightFlipper1_collide(parm)
  RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
  Select Case Int(Rnd*3)+1
    Case 1 : PlayFieldSoundAB "flip_hit_1", 0, 1
    Case 2 : PlayFieldSoundAB "flip_hit_2", 0, 1
    Case 3 : PlayFieldSoundAB "flip_hit_3", 0, 1
  End Select
End Sub

Sub physApron_Hit
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

'New algorithm for BallBallCollision
'For stereo, xGain is a Playsound volume multiplier that provides a Constant Power pan.
'For quad, multiple PlaySound commands are launched together that are panned and faded to their maximum extents where PlaySound's PAN and FADE have the least error.
'XVol and YVol are Playsound volume multipliers that provide a Constant Power "pan" and "fade".
'Subtracting XVol or YVol from 1 yields an inverse response.

'*******NOTE MOVED TO BILL'S Ball on BALL COLLISION METHOD ************

'Sub OnBallBallCollision(ball1, ball2, velocity)
' If PFOption = 1 or PFOption = 2 Then
'   PlaySound "BBcollide", 0, (Csng(velocity) ^2 / 2000) * xGain(ball1), AudioPan(ball1), 0, Pitch(ball1), 0, 0, 0  'Left & Right stereo or Top & Bottom stereo PF Speakers.
' End If
' If PFOption = 3 Then
'   PlaySound "BBcollide", 0, (Csng(velocity) ^2 / 2000) *    XVol(ball1)  *    YVol(ball1),  -1, 0, Pitch(ball1), 0, 0, -1 'Top Left Playfield Speaker
'   PlaySound "BBcollide", 0, (Csng(velocity) ^2 / 2000) * (1-XVol(ball1)) *    YVol(ball1),   1, 0, Pitch(ball1), 0, 0, -1 'Top Right Playfield Speaker
'   PlaySound "BBcollide", 0, (Csng(velocity) ^2 / 2000) *    XVol(ball1)  * (1-YVol(ball1)), -1, 0, Pitch(ball1), 0, 0,  1 'Bottom Left Playfield Speaker
'   PlaySound "BBcollide", 0, (Csng(velocity) ^2 / 2000) * (1-XVol(ball1)) * (1-YVol(ball1)),  1, 0, Pitch(ball1), 0, 0,  1 'Bottom Right Playfield Speaker
' End If
'End Sub

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
    If Ucase(Pan) = "LPAN" Then PlaySound SoundName, 0, ReelVolAdj * 1.00, -0.12, 0, 0, 1, 1, 0 'Panned 3/4 Left at 0dB * ReelVolAdj
    If Ucase(Pan) = "MPAN" Then PlaySound SoundName, 0, ReelVolAdj * 0.33,  0.00, 0, 0, 1, 1, 0 'Panned Middle at -3dB * ReelVolAdj
    If Ucase(Pan) = "RPAN" Then PlaySound SoundName, 0, ReelVolAdj * 1.00,  0.12, 0, 0, 1, 1, 0 'Panned 3/4 Right at 0dB * ReelVolAdj
  Else
    If Ucase(Pan) = "LPAN" Then PlaySound SoundName, 0, ReelVolAdj * 0.33, -0.12, 0, 0, 1, 1, 0 'Panned 3/4 Left at -3dB * ReelVolAdj
    If Ucase(Pan) = "MPAN" Then PlaySound SoundName, 0, ReelVolAdj * 0.11,  0.00, 0, 0, 1, 1, 0 'Panned Middle at -6dB * ReelVolAdj
    If Ucase(Pan) = "RPAN" Then PlaySound SoundName, 0, ReelVolAdj * 0.33,  0.12, 0, 0, 1, 1, 0 'Panned 3/4 Right at -3dB * ReelVolAdj
  End If
End Sub

Sub Table1_Exit
  If B2SOn Then controller.stop
  If flexError = False and showDt = False Then FlexDMD.Run = False
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
    If FlipperOn() then
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
    'debug.print Flipper1.currentangle &" = "& Endangle1 &"--"& Flipper2.currentangle &" = "& EndAngle2
    If Flipper2.currentangle = EndAngle2 Then
      For b = 0 to Ubound(BOT)
        If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper1) Then
          'Debug.Print "ball in flip1. exit"
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
'Dim PI: PI = 4*Atn(1)

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
    SOSRampup = 0 'was 2.5
  Case 1:
    SOSRampup = 0 'was 6
  Case 2:
    SOSRampup = 0 'was 8.5
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
' Thalamus - patched : ' Thalamus - patched :     aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
    aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
' Thalamus - patched :     aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
    aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
    if debugOn then TBPout.text = str
  End Sub

  public sub Dampenf(aBall, parm) 'Rubberizer is handled here
    dim RealCOR, DesiredCOR, str, coef
    DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
    RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id)+0.0001)
    coef = desiredcor / realcor
    If abs(aball.velx) < 2 and aball.vely < 0 and aball.vely > -3.75 then
' Thalamus - patched : ' Thalamus - patched :       aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
    aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
' Thalamus - patched :     aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
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

Sub ShadowHide_Hit()
Shadowblock.visible = false
End Sub

Sub ShadowShow_Hit()
Shadowblock.visible = true
End Sub

' *************************************************************************
'   JP's Reduced Display Driver Functions (based on script by Black)
' *************************************************************************

Dim FlexDMD, DMDScene, FlexDMDHighQuality, flexError
FlexDMDHighQuality = False
flexError = False

Sub DMD_Init()
    On Error Resume Next
        Set FlexDMD = CreateObject("FlexDMD.FlexDMD")
    If Err.Number <> 0 Then flexError = True
  On Error GoTo 0
        If flexError = False and showDT = False Then
            If FlexDMDHighQuality Then
                FlexDMD.TableFile = table1.Filename & ".vpx"
                FlexDMD.RenderMode = 2
                FlexDMD.Width = 256
                FlexDMD.Height = 64
                FlexDMD.Clear = True
                FlexDMD.GameName = cGameName
                FlexDMD.Run = True
                Set DMDScene = FlexDMD.NewGroup("Scene")
                DMDScene.AddActor FlexDMD.NewImage("Back", "VPX.nags")
                DMDScene.GetImage("Back").SetSize FlexDMD.Width, FlexDMD.Height
                FlexDMD.LockRenderThread
                FlexDMD.Stage.AddActor DMDScene
                FlexDMD.UnlockRenderThread
            Else
                FlexDMD.TableFile = Table1.Filename & ".vpx"  'Needed
                FlexDMD.RenderMode = 2
                FlexDMD.Width = 128
                FlexDMD.Height = 32
                FlexDMD.Clear = True
                FlexDMD.GameName = cGameName
                FlexDMD.Run = True
                Set DMDScene = FlexDMD.NewGroup("Scene")
                DMDScene.AddActor FlexDMD.NewImage("Back", "VPX.nags")
                DMDScene.GetImage("Back").SetSize FlexDMD.Width, FlexDMD.Height
                FlexDMD.LockRenderThread
                FlexDMD.Stage.AddActor DMDScene
                FlexDMD.UnlockRenderThread
            End If
        End If
End Sub



'******************************************************
'         Turn Table Bumpers
' This section was written by Roth and is something that
' I wasn't sure was possible.  The only reason you are
' playng this table is because of Bill's code.
'******************************************************

Dim BBall1, BBall2, BBall3, BBall4, BBall5, BBall6, Skirts
Dim MSize, MMass
MSize = 45
MMass = 0.01

Sub CreateTurnTable

  SpinnerBallTimer.enabled = True
  PlayFieldSound "ACMotor", -1, TurnTable, 0.0001

  Set BBall1 = BKicker1.CreatesizedballwithMass(MSize,MMass)
  BBall1.visible = False
  BKicker1.Enabled = 0

  Set BBall2 = BKicker2.CreatesizedballwithMass(MSize,MMass)
  BBall2.visible = False
  BKicker2.Enabled = 0

  Set BBall3 = BKicker3.CreatesizedballwithMass(MSize,MMass)
  BBall3.visible = False
  BKicker3.Enabled = 0

  Set BBall4 = BKicker4.CreatesizedballwithMass(MSize,MMass)
  BBall4.visible = False
  BKicker4.Enabled = 0

  Set BBall5 = BKicker5.CreatesizedballwithMass(MSize,MMass)
  BBall5.visible = False
  BKicker5.Enabled = 0

  Set BBall6 = BKicker6.CreatesizedballwithMass(MSize,MMass)
  BBall6.visible = False
  BKicker6.Enabled = 0

  Skirts = Array(BBall1, BBall2, BBall3, BBall4, BBall5, BBall6)

End Sub

Dim discPosition, discSpinSpeed
dim spinAngle, degAngle, spinAngle2, degAngle2, startAngle
dim spinAngle3, DegAngle3, spinAngle4, degAngle4, spinAngle5, degAngle5, spinAngle6, degAngle6
dim discX, discY
dim bballheight

bballheight = 45

startAngle = 0
discX = BM_bump.x     '425.4256
discY = BM_Bump.y     '853.2974

Const cDiscRadius = 212.5   '200
discSpinSpeed = -18       'I think this is equivalent to degrees per second

Sub SpinnerBallTimer_Timer()
  Dim BP, BP_ang

  discPosition = discPosition + discSpinSpeed * Me.Interval / 1000

  Do While discPosition < 0 : discPosition = discPosition + 360 : Loop
  Do While discPosition > 360 : discPosition = discPosition - 360 : Loop

  degAngle  = startAngle + discPosition
  degAngle2 = degAngle + 60
  degAngle3 = degAngle2 + 60
  degAngle4 = degAngle3 + 60
  degAngle5 = degAngle4 + 60
  degAngle6 = degAngle5 + 60

  BP_Ang = degAngle - 150

  spinAngle  = PI * (degAngle) / 180
  spinAngle2 = PI * (degAngle2) / 180
  spinAngle3 = PI * (degAngle3) / 180
  spinAngle4 = PI * (degAngle4) / 180
  spinAngle5 = PI * (degAngle5) / 180
  spinAngle6 = PI * (degAngle6) / 180

  For each BP in BP_screws: BP.rotz = BP_Ang : Next
  For each BP in BP_bump: BP.rotz = BP_Ang : Next
  For each BP in BP_Playfield2: BP.rotz = BP_Ang : Next

  BBall1.x = discX + (cDiscRadius * Cos(spinAngle))
  BBall1.y = discY + (cDiscRadius * Sin(spinAngle))
  BBall1.z = BBallHeight

  For each BP in BP_skirt1:BP.ObjRotz = BP_Ang:Next
  For each BP in BP_ring1:BP.ObjRotz = BP_Ang:Next

  BBall2.x = discX + (cDiscRadius * Cos(spinAngle2))
  BBall2.y = discY + (cDiscRadius * Sin(spinAngle2))
  BBall2.z = BBallHeight

  For each BP in BP_skirt2:BP.ObjRotz = BP_Ang:Next
  For each BP in BP_ring2:BP.ObjRotz = BP_Ang:Next

  BBall3.x = discX + (cDiscRadius * Cos(spinAngle3))
  BBall3.y = discY + (cDiscRadius * Sin(spinAngle3))
  BBall3.z = BBallHeight

  For each BP in BP_skirt3:BP.ObjRotz = BP_Ang:Next
  For each BP in BP_ring3:BP.ObjRotz = BP_Ang:Next

  BBall4.x = discX + (cDiscRadius * Cos(spinAngle4))
  BBall4.y = discY + (cDiscRadius * Sin(spinAngle4))
  BBall4.z = BBallHeight

  For each BP in BP_skirt4:BP.ObjRotz = BP_Ang:Next
  For each BP in BP_ring4:BP.ObjRotz = BP_Ang:Next

  BBall5.x = discX + (cDiscRadius * Cos(spinAngle5))
  BBall5.y = discY + (cDiscRadius * Sin(spinAngle5))
  BBall5.z = BBallHeight

  For each BP in BP_skirt5:BP.ObjRotz = BP_Ang:Next
  For each BP in BP_ring5:BP.ObjRotz = BP_Ang:Next

  BBall6.x = discX + (cDiscRadius * Cos(spinAngle6))
  BBall6.y = discY + (cDiscRadius * Sin(spinAngle6))
  BBall6.z = BBallHeight

  For each BP in BP_skirt6:BP.ObjRotz = BP_Ang:Next
  For each BP in BP_ring6:BP.ObjRotz = BP_Ang:Next

End Sub

sub bumpert001_timer    '****************part of bumper skirt animation
  Dim z, BP
  z = BM_ring1.transz + 5
  If z >= 0 Then
    z = 0
    me.enabled=0
  End If

  For Each BP in BP_ring1 : BP.transz = z: Next
end sub

sub bumpert002_timer    '****************part of bumper skirt animation
  Dim z, BP
  z = BM_ring2.transz + 5
  If z >= 0 Then
    z = 0
    me.enabled=0
  End If

  For Each BP in BP_ring2 : BP.transz = z: Next
end sub

sub bumpert003_timer    '****************part of bumper skirt animation
  Dim z, BP
  z = BM_ring3.transz + 5
  If z >= 0 Then
    z = 0
    me.enabled=0
  End If

  For Each BP in BP_ring3 : BP.transz = z: Next
end sub

sub bumpert004_timer    '****************part of bumper skirt animation
  Dim z, BP
  z = BM_ring4.transz + 5
  If z >= 0 Then
    z = 0
    me.enabled=0
  End If

  For Each BP in BP_ring4 : BP.transz = z: Next

end sub

sub bumpert005_timer    '****************part of bumper skirt animation
  Dim z, BP
  z = BM_ring5.transz + 5
  If z >= 0 Then
    z = 0
    me.enabled=0
  End If

  For Each BP in BP_ring5 : BP.transz = z: Next
end sub

sub bumpert006_timer    '****************part of bumper skirt animation
  Dim z, BP
  z = BM_ring6.transz + 5
  If z >= 0 Then
    z = 0
    me.enabled=0
  End If

  For Each BP in BP_ring6 : BP.transz = z: Next
end sub

Sub AnimateBumperSkirts
  dim r, g, s, x, y, b, tz
  ' Animate Bumper switch (experimental)
  For r = 0 To 5
    g = 10000.
    For s = 0 to UBound(gBOT)
      x = Skirts(r).x - gBOT(s).x
      y = Skirts(r).y - gBOT(s).y
      b = x * x + y * y
      If b < g Then g = b
    Next
    tz = 0
    If g < 80 * 80 Then
      tz = -4
    End If
    If r = 0 Then For Each x in BP_skirt1: x.transZ = tz:Next
    If r = 1 Then For Each x in BP_skirt2: x.transZ = tz:Next
    If r = 2 Then For Each x in BP_skirt3: x.transZ = tz:Next
    If r = 3 Then For Each x in BP_skirt4: x.transZ = tz:Next
    If r = 4 Then For Each x in BP_skirt5: x.transZ = tz:Next
    If r = 5 Then For Each x in BP_skirt6: x.transZ = tz:Next
  Next
End Sub

'*************** needed for bumper skirt animation
'*************** NOTE: set bumper object timer to around 150-175 in order to be able
'***************       to actually see the animaation, adjust to your liking

Const PI = 3.1415926
Const SkirtTilt=5   'angle of skirt tilting in degrees

Function SkirtAX(bumper, bumperball)

  skirtAX=cos(skirtA(bumper,bumperball))*(SkirtTilt)    'x component of angle
  if (bumper.y<bumperball.y) then skirtAX=skirtAX*-1  'adjust for ball hit bottom half

End Function

Function SkirtAY(bumper, bumperball)

  skirtAY=sin(skirtA(bumper,bumperball))*(SkirtTilt)    'y component of angle
  if (bumper.x>bumperball.x) then skirtAY=skirtAY*-1  'adjust for ball hit left half

End Function

Function SkirtA(bumper, bumperball)
  dim hitx, hity, dx, dy
  hitx=bumperball.x
  hity=bumperball.y

  dy=Abs(hity-bumper.y)         'y offset ball at hit to center of bumper
  if dy=0 then dy=0.0000001
  dx=Abs(hitx-bumper.x)         'x offset ball at hit to center of bumper
  skirtA=(atn(dx/dy)) '/(PI/180)      'angle in radians to ball from center of Bumper1
End Function


'********************************************
' Ball Collision, spinner collision and Sound
'********************************************

dim bxvel, byvel, hitball

Sub OnBallBallCollision(ball1, ball2, velocity)
  dim collAngle,bvelx,bvely,whichBall, whichBall2, BP
  If ((ball1.radius > 40) or (ball2.radius > 40)) and tilt = False then
    If ball1.radius > 40 Then
      collAngle = GetCollisionAngle(ball1.x,ball1.y,ball2.x,ball2.y)
      set hitball = ball2
      playfieldSound "popBump", 0, ball2, 1
      If ball1.x = BBall1.x and ball1.y = BBall1.y Then
        set whichball = BBall1
        whichBall2 = 1
      ElseIf ball1.x = BBall2.x and ball1.y = BBall2.y Then
        set whichball = BBall2
        whichBall2 = 2
      ElseIf ball1.x = BBall3.x and ball1.y = BBall3.y Then
        set whichball = BBall3
        whichBall2 = 3
      ElseIf ball1.x = BBall4.x and ball1.y = BBall4.y Then
        set whichball = BBall4
        whichBall2 = 4
      ElseIf ball1.x = BBall5.x and ball1.y = BBall5.y Then
        set whichball = BBall5
        whichBall2 = 5
      Else
        set whichball = BBall6
        whichBall2 = 6
      End If
    else
      collAngle = GetCollisionAngle(ball2.x,ball2.y,ball1.x,ball1.y)
      set hitball = ball1
      playfieldSound  "popBump", 0, ball1, 1
      If ball2.x = BBall1.x and ball2.y = BBall1.y Then
        set whichball = BBall1
        whichBall2 = 1
      ElseIf ball2.x = BBall2.x and ball2.y = BBall2.y Then
        set whichball = BBall2
        whichBall2 = 2
      ElseIf ball2.x = BBall3.x and ball2.y = BBall3.y Then
        set whichball = BBall3
        whichBall2 = 3
      ElseIf ball2.x = BBall4.x and ball2.y = BBall4.y Then
        set whichball = BBall4
        whichBall2 = 4
      ElseIf ball2.x = BBall5.x and ball2.y = BBall5.y Then
        set whichball = BBall5
        whichBall2 = 5
      Else
        set whichball = BBall6
        whichBall2 = 6
      End If
    End If

    'debug.print collAngle*180/PI & " " & hitball.velx &  " " &  hitball.vely

    dim tempvel

    tempvel = BallVel(hitball) * 0.6

    If tempvel < 10 Then tempvel = 10

    bxvel = tempvel * cos(collangle+PI)
    byvel = tempvel * sin(collAngle+PI)

    bumpit.enabled = 1

    'PlayFieldSoundAB "popBump", 0, 1
    If B2SOn Then DOF 420, DOFPulse


    If whichBall2 = 1 Then
      For Each BP in BP_ring1 : BP.transz = -20: Next
      bumpert001.enabled = 1
      If tilt = False Then
        horse1Flag = 1
        bumpHossScore(1)
        If horsesRunning = 0 Then startHorses
        If motorOn = 0 Then scoreMotor.enabled = 1
      End If
    ElseIf whichBall2 = 2 Then
      For Each BP in BP_ring2 : BP.transz = -20: Next
      bumpert002.enabled = 1
      If tilt = False Then
        horse2Flag = 1
        bumpHossScore(2)
        If horsesRunning = 0 Then startHorses
        If motorOn = 0 Then scoreMotor.enabled = 1
      End If
    ElseIf whichBall2 = 3 Then
      For Each BP in BP_ring3 : BP.transz = -20: Next
      bumpert003.enabled = 1
      If tilt = False Then
        horse3Flag = 1
        bumpHossScore(3)
        If horsesRunning = 0 Then startHorses
        If motorOn = 0 Then scoreMotor.enabled = 1
      End If
    ElseIf whichBall2 = 4 Then
      For Each BP in BP_ring4 : BP.transz = -20: Next
      bumpert004.enabled = 1
      If tilt = False Then
        horse4Flag = 1
        bumpHossScore(4)
        If horsesRunning = 0 Then startHorses
        If motorOn = 0 Then scoreMotor.enabled = 1
      End If
    ElseIf whichBall2 = 5 Then
      For Each BP in BP_ring5 : BP.transz = -20: Next
      bumpert005.enabled = 1
      If tilt = False Then
        horse5Flag = 1
        bumpHossScore(5)
        If horsesRunning = 0 Then startHorses
        If motorOn = 0 Then scoreMotor.enabled = 1
      End If
    Else
      For Each BP in BP_ring6 : BP.transz = -20: Next
      bumpert006.enabled = 1
      If tilt = False Then
        horse6Flag = 1
        bumpHossScore(6)
        If horsesRunning = 0 Then startHorses
        If motorOn = 0 Then scoreMotor.enabled = 1
      End If
    End If
  Else
    If ball1.z < -30 Then
      Exit Sub
    End If
    If velocity > 1 and PFOption = 1 or PFOption = 2 Then
      PlaySound "rubber_hit_3", 0, (Csng(velocity) ^2 / 2000) * xGain(ball1), AudioPan(ball1), 0, Pitch(ball1), 0, 0, 0 'Left & Right stereo or Top & Bottom stereo PF Speakers.
    End If
    If velocity > 1 and PFOption = 3 Then
      PlaySound "rubber_hit_3", 0, (Csng(velocity) ^2 / 2000) *    XVol(ball1)  *    YVol(ball1),  -1, 0, Pitch(ball1), 0, 0, -1  'Top Left Playfield Speaker
      PlaySound "rubber_hit_3", 0, (Csng(velocity) ^2 / 2000) * (1-XVol(ball1)) *    YVol(ball1),   1, 0, Pitch(ball1), 0, 0, -1  'Top Right Playfield Speaker
      PlaySound "rubber_hit_3", 0, (Csng(velocity) ^2 / 2000) *    XVol(ball1)  * (1-YVol(ball1)), -1, 0, Pitch(ball1), 0, 0,  1  'Bottom Left Playfield Speaker
      PlaySound "rubber_hit_3", 0, (Csng(velocity) ^2 / 2000) * (1-XVol(ball1)) * (1-YVol(ball1)),  1, 0, Pitch(ball1), 0, 0,  1  'Bottom Right Playfield Speaker
    End If
  End If
  If ((ball1.radius < 40) and (ball2.radius < 40)) Then
    If PFOption = 1 or PFOption = 2 Then
      PlaySound "BBcollide", 0, (Csng(velocity) ^2 / 2000) * xGain(ball1), AudioPan(ball1), 0, Pitch(ball1), 0, 0, 0  'Left & Right stereo or Top & Bottom stereo PF Speakers.
    End If
    If PFOption = 3 Then
      PlaySound "BBcollide", 0, (Csng(velocity) ^2 / 2000) *    XVol(ball1)  *    YVol(ball1),  -1, 0, Pitch(ball1), 0, 0, -1 'Top Left Playfield Speaker
      PlaySound "BBcollide", 0, (Csng(velocity) ^2 / 2000) * (1-XVol(ball1)) *    YVol(ball1),   1, 0, Pitch(ball1), 0, 0, -1 'Top Right Playfield Speaker
      PlaySound "BBcollide", 0, (Csng(velocity) ^2 / 2000) *    XVol(ball1)  * (1-YVol(ball1)), -1, 0, Pitch(ball1), 0, 0,  1 'Bottom Left Playfield Speaker
      PlaySound "BBcollide", 0, (Csng(velocity) ^2 / 2000) * (1-XVol(ball1)) * (1-YVol(ball1)),  1, 0, Pitch(ball1), 0, 0,  1 'Bottom Right Playfield Speaker
    End If
  End If
End Sub

sub bumpit_timer
  hitball.velx = bxvel
  hitball.vely = byvel
  me.enabled = 0
end sub

Function GetCollisionAngle(ax, ay, bx, by)
  Dim ang
  Dim collisionV:Set collisionV = new jVector
  collisionV.SetXY ax - bx, ay - by
  GetCollisionAngle = collisionV.ang
End Function

Function NormAngle(angle)
  NormAngle = angle
  Do While NormAngle>2 * pi
    NormAngle = NormAngle - 2 * pi
  Loop
  Do While NormAngle <0
    NormAngle = NormAngle + 2 * pi
  Loop
End Function

Class jVector
  Private m_mag, m_ang, pi

  Sub Class_Initialize
    m_mag = CDbl(0)
    m_ang = CDbl(0)
    pi = CDbl(3.14159265358979323846)
  End Sub

  Public Function add(anothervector)
    Dim tx, ty, theta
    If TypeName(anothervector) = "jVector" then
      Set add = new jVector
      add.SetXY x + anothervector.x, y + anothervector.y
    End If
  End Function

  Public Function multiply(scalar)
    Set multiply = new jVector
    multiply.SetXY x * scalar, y * scalar
  End Function

  Sub ShiftAxes(theta)
    ang = ang - theta
  end Sub

  Sub SetXY(tx, ty)

    if tx = 0 And ty = 0 Then
      ang = 0
    elseif tx = 0 And ty <0 then
      ang = - pi / 180 ' -90 degrees
    elseif tx = 0 And ty>0 then
      ang = pi / 180   ' 90 degrees
    else
      ang = atn(ty / tx)
      if tx <0 then ang = ang + pi ' Add 180 deg if in quadrant 2 or 3
    End if

    mag = sqr(tx ^2 + ty ^2)
  End Sub

  Property Let mag(nmag)
    m_mag = nmag
  End Property

  Property Get mag
    mag = m_mag
  End Property

  Property Let ang(nang)
    m_ang = nang
    Do While m_ang>2 * pi
      m_ang = m_ang - 2 * pi
    Loop
    Do While m_ang <0
      m_ang = m_ang + 2 * pi
    Loop
  End Property

  Property Get ang
    Do While m_ang>2 * pi
      m_ang = m_ang - 2 * pi
    Loop
    Do While m_ang <0
      m_ang = m_ang + 2 * pi
    Loop
    ang = m_ang
  End Property

  Property Get x
    x = m_mag * cos(ang)
  End Property

  Property Get y
    y = m_mag * sin(ang)
  End Property

  Property Get dump
    dump = "vector "
    Select Case CInt(ang + pi / 8)
      case 0, 8:dump = dump & "->"
      case 1:dump = dump & "/'"
      case 2:dump = dump & "/\"
      case 3:dump = dump & "'\"
      case 4:dump = dump & "<-"
      case 5:dump = dump & ":/"
      case 6:dump = dump & "\/"
      case 7:dump = dump & "\:"
    End Select

    dump = dump & " mag:" & CLng(mag * 10) / 10 & ", ang:" & CLng(ang * 180 / pi) & ", x:" & CLng(x * 10) / 10 & ", y:" & CLng(y * 10) / 10
  End Property
End Class

' VLM Bumper Arrays - Start
' Arrays per baked part
Dim BP_Playfield2: BP_Playfield2=Array(BM_Playfield2, LM_bump_lights_B001_Playfield, LM_bump_lights_B002_Playfield, LM_bump_lights_B003_Playfield, LM_bump_lights_B004_Playfield, LM_bump_lights_B005_Playfield)
Dim BP_Screws: BP_Screws=Array(BM_Screws, LM_bump_lights_B001_Screws, LM_bump_lights_B002_Screws, LM_bump_lights_B003_Screws, LM_bump_lights_B004_Screws, LM_bump_lights_B005_Screws, LM_bump_lights_B006_Screws)
Dim BP_bump: BP_bump=Array(BM_bump, LM_bump_lights_B001_bump, LM_bump_lights_B002_bump, LM_bump_lights_B003_bump, LM_bump_lights_B004_bump, LM_bump_lights_B005_bump, LM_bump_lights_B006_bump)
Dim BP_ring1: BP_ring1=Array(BM_ring1, LM_bump_lights_B001_ring1, LM_bump_lights_B005_ring1, LM_bump_lights_B006_ring1)
Dim BP_ring2: BP_ring2=Array(BM_ring2, LM_bump_lights_B001_ring2, LM_bump_lights_B002_ring2, LM_bump_lights_B004_ring2)
Dim BP_ring3: BP_ring3=Array(BM_ring3, LM_bump_lights_B002_ring3, LM_bump_lights_B003_ring3, LM_bump_lights_B004_ring3)
Dim BP_ring4: BP_ring4=Array(BM_ring4, LM_bump_lights_B002_ring4, LM_bump_lights_B003_ring4, LM_bump_lights_B004_ring4)
Dim BP_ring5: BP_ring5=Array(BM_ring5, LM_bump_lights_B001_ring5, LM_bump_lights_B002_ring5, LM_bump_lights_B003_ring5, LM_bump_lights_B004_ring5, LM_bump_lights_B005_ring5)
Dim BP_ring6: BP_ring6=Array(BM_ring6, LM_bump_lights_B001_ring6, LM_bump_lights_B002_ring6, LM_bump_lights_B004_ring6, LM_bump_lights_B005_ring6, LM_bump_lights_B006_ring6)
Dim BP_skirt1: BP_skirt1=Array(BM_skirt1, LM_bump_lights_B001_skirt1)
Dim BP_skirt2: BP_skirt2=Array(BM_skirt2, LM_bump_lights_B002_skirt2)
Dim BP_skirt3: BP_skirt3=Array(BM_skirt3, LM_bump_lights_B003_skirt3)
Dim BP_skirt4: BP_skirt4=Array(BM_skirt4, LM_bump_lights_B004_skirt4)
Dim BP_skirt5: BP_skirt5=Array(BM_skirt5, LM_bump_lights_B005_skirt5)
Dim BP_skirt6: BP_skirt6=Array(BM_skirt6, LM_bump_lights_B006_skirt6)
' Arrays per lighting scenario
Dim BL_World2: BL_World2=Array(BM_Playfield2, BM_Screws, BM_bump, BM_ring1, BM_ring2, BM_ring3, BM_ring4, BM_ring5, BM_ring6, BM_skirt1, BM_skirt2, BM_skirt3, BM_skirt4, BM_skirt5, BM_skirt6)
Dim BL_bump_lights_B001: BL_bump_lights_B001=Array(LM_bump_lights_B001_Playfield, LM_bump_lights_B001_Screws, LM_bump_lights_B001_bump, LM_bump_lights_B001_ring1, LM_bump_lights_B001_ring2, LM_bump_lights_B001_ring5, LM_bump_lights_B001_ring6, LM_bump_lights_B001_skirt1)
Dim BL_bump_lights_B002: BL_bump_lights_B002=Array(LM_bump_lights_B002_Playfield, LM_bump_lights_B002_Screws, LM_bump_lights_B002_bump, LM_bump_lights_B002_ring2, LM_bump_lights_B002_ring3, LM_bump_lights_B002_ring4, LM_bump_lights_B002_ring5, LM_bump_lights_B002_ring6, LM_bump_lights_B002_skirt2)
Dim BL_bump_lights_B003: BL_bump_lights_B003=Array(LM_bump_lights_B003_Playfield, LM_bump_lights_B003_Screws, LM_bump_lights_B003_bump, LM_bump_lights_B003_ring3, LM_bump_lights_B003_ring4, LM_bump_lights_B003_ring5, LM_bump_lights_B003_skirt3)
Dim BL_bump_lights_B004: BL_bump_lights_B004=Array(LM_bump_lights_B004_Playfield, LM_bump_lights_B004_Screws, LM_bump_lights_B004_bump, LM_bump_lights_B004_ring2, LM_bump_lights_B004_ring3, LM_bump_lights_B004_ring4, LM_bump_lights_B004_ring5, LM_bump_lights_B004_ring6, LM_bump_lights_B004_skirt4)
Dim BL_bump_lights_B005: BL_bump_lights_B005=Array(LM_bump_lights_B005_Playfield, LM_bump_lights_B005_Screws, LM_bump_lights_B005_bump, LM_bump_lights_B005_ring1, LM_bump_lights_B005_ring5, LM_bump_lights_B005_ring6, LM_bump_lights_B005_skirt5)
Dim BL_bump_lights_B006: BL_bump_lights_B006=Array(LM_bump_lights_B006_Screws, LM_bump_lights_B006_bump, LM_bump_lights_B006_ring1, LM_bump_lights_B006_ring6, LM_bump_lights_B006_skirt6)
' Global arrays
Dim BG_Bakemap2: BG_Bakemap2=Array(BM_Playfield2, BM_Screws, BM_bump, BM_ring1, BM_ring2, BM_ring3, BM_ring4, BM_ring5, BM_ring6, BM_skirt1, BM_skirt2, BM_skirt3, BM_skirt4, BM_skirt5, BM_skirt6)
Dim BG_Lightmap2: BG_Lightmap2=Array(LM_bump_lights_B001_Playfield, LM_bump_lights_B001_Screws, LM_bump_lights_B001_bump, LM_bump_lights_B001_ring1, LM_bump_lights_B001_ring2, LM_bump_lights_B001_ring5, LM_bump_lights_B001_ring6, LM_bump_lights_B001_skirt1, LM_bump_lights_B002_Playfield, LM_bump_lights_B002_Screws, LM_bump_lights_B002_bump, LM_bump_lights_B002_ring2, LM_bump_lights_B002_ring3, LM_bump_lights_B002_ring4, LM_bump_lights_B002_ring5, LM_bump_lights_B002_ring6, LM_bump_lights_B002_skirt2, LM_bump_lights_B003_Playfield, LM_bump_lights_B003_Screws, LM_bump_lights_B003_bump, LM_bump_lights_B003_ring3, LM_bump_lights_B003_ring4, LM_bump_lights_B003_ring5, LM_bump_lights_B003_skirt3, LM_bump_lights_B004_Playfield, LM_bump_lights_B004_Screws, LM_bump_lights_B004_bump, LM_bump_lights_B004_ring2, LM_bump_lights_B004_ring3, LM_bump_lights_B004_ring4, LM_bump_lights_B004_ring5, LM_bump_lights_B004_ring6, LM_bump_lights_B004_skirt4, LM_bump_lights_B005_Playfield, LM_bump_lights_B005_Screws, _
  LM_bump_lights_B005_bump, LM_bump_lights_B005_ring1, LM_bump_lights_B005_ring5, LM_bump_lights_B005_ring6, LM_bump_lights_B005_skirt5, LM_bump_lights_B006_Screws, LM_bump_lights_B006_bump, LM_bump_lights_B006_ring1, LM_bump_lights_B006_ring6, LM_bump_lights_B006_skirt6)
Dim BG_All2: BG_All2=Array(BM_Playfield2, BM_Screws, BM_bump, BM_ring1, BM_ring2, BM_ring3, BM_ring4, BM_ring5, BM_ring6, BM_skirt1, BM_skirt2, BM_skirt3, BM_skirt4, BM_skirt5, BM_skirt6, LM_bump_lights_B001_Playfield, LM_bump_lights_B001_Screws, LM_bump_lights_B001_bump, LM_bump_lights_B001_ring1, LM_bump_lights_B001_ring2, LM_bump_lights_B001_ring5, LM_bump_lights_B001_ring6, LM_bump_lights_B001_skirt1, LM_bump_lights_B002_Playfield, LM_bump_lights_B002_Screws, LM_bump_lights_B002_bump, LM_bump_lights_B002_ring2, LM_bump_lights_B002_ring3, LM_bump_lights_B002_ring4, LM_bump_lights_B002_ring5, LM_bump_lights_B002_ring6, LM_bump_lights_B002_skirt2, LM_bump_lights_B003_Playfield, LM_bump_lights_B003_Screws, LM_bump_lights_B003_bump, LM_bump_lights_B003_ring3, LM_bump_lights_B003_ring4, LM_bump_lights_B003_ring5, LM_bump_lights_B003_skirt3, LM_bump_lights_B004_Playfield, LM_bump_lights_B004_Screws, LM_bump_lights_B004_bump, LM_bump_lights_B004_ring2, LM_bump_lights_B004_ring3, LM_bump_lights_B004_ring4, _
  LM_bump_lights_B004_ring5, LM_bump_lights_B004_ring6, LM_bump_lights_B004_skirt4, LM_bump_lights_B005_Playfield, LM_bump_lights_B005_Screws, LM_bump_lights_B005_bump, LM_bump_lights_B005_ring1, LM_bump_lights_B005_ring5, LM_bump_lights_B005_ring6, LM_bump_lights_B005_skirt5, LM_bump_lights_B006_Screws, LM_bump_lights_B006_bump, LM_bump_lights_B006_ring1, LM_bump_lights_B006_ring6, LM_bump_lights_B006_skirt6)
' VLM Bumper Arrays - End

' VLM bg Arrays - Start
' Arrays per baked part
Dim BP_bgbg: BP_bgbg=Array(bgBM_bg, bgLM_l_d006_bg, bgLM_l_d003_bg, bgLM_l_h003_bg, bgLM_l_h001_bg, bgLM_l_h004_bg, bgLM_l_d004_bg, bgLM_l_h005_bg, bgLM_l_d005_bg, bgLM_l_h006_bg, bgLM_GI_bg, bgLM_l_raceOverIn_bg, bgLM_l_bip002_bg, bgLM_l_bip003_bg, bgLM_l_bip004_bg, bgLM_l_bip005_bg, bgLM_l_bip001_bg, bgLM_l_balls_bg, bgLM_l_tiltLight_bg, bgLM_l_gameOverLight_bg, bgLM_l_ballInPlayLight_bg, bgLM_l_d001_bg, bgLM_l_d002_bg, bgLM_l_h002_bg)
Dim BP_bghoss001: BP_bghoss001=Array(bgBM_hoss001, bgLM_GI_hoss001)
Dim BP_bghoss002: BP_bghoss002=Array(bgBM_hoss002)
Dim BP_bghoss003: BP_bghoss003=Array(bgBM_hoss003, bgLM_GI_hoss003)
Dim BP_bghoss004: BP_bghoss004=Array(bgBM_hoss004, bgLM_GI_hoss004)
Dim BP_bghoss005: BP_bghoss005=Array(bgBM_hoss005, bgLM_GI_hoss005)
Dim BP_bghoss006: BP_bghoss006=Array(bgBM_hoss006, bgLM_GI_hoss006)
Dim allHosses: allHosses = Array(BP_bghoss001, BP_bghoss002, BP_bghoss003, BP_bghoss004, BP_bghoss005, BP_bghoss006)
Dim BP_bgplayfield: BP_bgplayfield=Array(bgBM_playfield, bgLM_l_d006_playfield, bgLM_l_d003_playfield, bgLM_l_h003_playfield, bgLM_l_h001_playfield, bgLM_l_h004_playfield, bgLM_l_d004_playfield, bgLM_l_h005_playfield, bgLM_l_d005_playfield, bgLM_l_h006_playfield, bgLM_GI_playfield, bgLM_l_raceOverIn_playfield, bgLM_l_bip002_playfield, bgLM_l_bip003_playfield, bgLM_l_bip004_playfield, bgLM_l_bip005_playfield, bgLM_l_bip001_playfield, bgLM_l_balls_playfield, bgLM_l_tiltLight_playfield, bgLM_l_gameOverLight_playfield, bgLM_l_ballInPlayLight_playfiel, bgLM_l_d001_playfield, bgLM_l_d002_playfield, bgLM_l_h002_playfield)
' Arrays per lighting scenario
Dim BL_bgGI: BL_bgGI=Array(bgLM_GI_bg, bgLM_GI_hoss001, bgLM_GI_hoss003, bgLM_GI_hoss004, bgLM_GI_hoss005, bgLM_GI_hoss006, bgLM_GI_playfield)
Dim BL_bgWorld: BL_bgWorld=Array(bgBM_bg, bgBM_hoss001, bgBM_hoss003, bgBM_hoss004, bgBM_hoss005, bgBM_hoss006, bgBM_playfield)
Dim BL_bgl_ballInPlayLight: BL_bgl_ballInPlayLight=Array(bgLM_l_ballInPlayLight_bg, bgLM_l_ballInPlayLight_playfiel)
Dim BL_bgl_balls: BL_bgl_balls=Array(bgLM_l_balls_bg, bgLM_l_balls_playfield)
Dim BL_bgl_bip001: BL_bgl_bip001=Array(bgLM_l_bip001_bg, bgLM_l_bip001_playfield)
Dim BL_bgl_bip002: BL_bgl_bip002=Array(bgLM_l_bip002_bg, bgLM_l_bip002_playfield)
Dim BL_bgl_bip003: BL_bgl_bip003=Array(bgLM_l_bip003_bg, bgLM_l_bip003_playfield)
Dim BL_bgl_bip004: BL_bgl_bip004=Array(bgLM_l_bip004_bg, bgLM_l_bip004_playfield)
Dim BL_bgl_bip005: BL_bgl_bip005=Array(bgLM_l_bip005_bg, bgLM_l_bip005_playfield)
Dim BL_bgl_d001: BL_bgl_d001=Array(bgLM_l_d001_bg, bgLM_l_d001_playfield)
Dim BL_bgl_d002: BL_bgl_d002=Array(bgLM_l_d002_bg, bgLM_l_d002_playfield)
Dim BL_bgl_d003: BL_bgl_d003=Array(bgLM_l_d003_bg, bgLM_l_d003_playfield)
Dim BL_bgl_d004: BL_bgl_d004=Array(bgLM_l_d004_bg, bgLM_l_d004_playfield)
Dim BL_bgl_d005: BL_bgl_d005=Array(bgLM_l_d005_bg, bgLM_l_d005_playfield)
Dim BL_bgl_d006: BL_bgl_d006=Array(bgLM_l_d006_bg, bgLM_l_d006_playfield)
Dim BL_bgl_gameOverLight: BL_bgl_gameOverLight=Array(bgLM_l_gameOverLight_bg, bgLM_l_gameOverLight_playfield)
Dim BL_bgl_h001: BL_bgl_h001=Array(bgLM_l_h001_bg, bgLM_l_h001_playfield)
Dim BL_bgl_h002: BL_bgl_h002=Array(bgLM_l_h002_bg, bgLM_l_h002_playfield)
Dim BL_bgl_h003: BL_bgl_h003=Array(bgLM_l_h003_bg, bgLM_l_h003_playfield)
Dim BL_bgl_h004: BL_bgl_h004=Array(bgLM_l_h004_bg, bgLM_l_h004_playfield)
Dim BL_bgl_h005: BL_bgl_h005=Array(bgLM_l_h005_bg, bgLM_l_h005_playfield)
Dim BL_bgl_h006: BL_bgl_h006=Array(bgLM_l_h006_bg, bgLM_l_h006_playfield)
Dim BL_bgl_raceOverIn: BL_bgl_raceOverIn=Array(bgLM_l_raceOverIn_bg, bgLM_l_raceOverIn_playfield)
Dim BL_bgl_tiltLight: BL_bgl_tiltLight=Array(bgLM_l_tiltLight_bg, bgLM_l_tiltLight_playfield)
' Global arrays
Dim BGbg_Bakemap: BGbg_Bakemap=Array(bgBM_bg, bgBM_hoss001, bgBM_hoss003, bgBM_hoss004, bgBM_hoss005, bgBM_hoss006, bgBM_playfield)
Dim BGbg_Lightmap: BGbg_Lightmap=Array(bgLM_GI_bg, bgLM_GI_hoss001, bgLM_GI_hoss003, bgLM_GI_hoss004, bgLM_GI_hoss005, bgLM_GI_hoss006, bgLM_GI_playfield, bgLM_l_ballInPlayLight_bg, bgLM_l_ballInPlayLight_playfiel, bgLM_l_balls_bg, bgLM_l_balls_playfield, bgLM_l_bip001_bg, bgLM_l_bip001_playfield, bgLM_l_bip002_bg, bgLM_l_bip002_playfield, bgLM_l_bip003_bg, bgLM_l_bip003_playfield, bgLM_l_bip004_bg, bgLM_l_bip004_playfield, bgLM_l_bip005_bg, bgLM_l_bip005_playfield, bgLM_l_d001_bg, bgLM_l_d001_playfield, bgLM_l_d002_bg, bgLM_l_d002_playfield, bgLM_l_d003_bg, bgLM_l_d003_playfield, bgLM_l_d004_bg, bgLM_l_d004_playfield, bgLM_l_d005_bg, bgLM_l_d005_playfield, bgLM_l_d006_bg, bgLM_l_d006_playfield, bgLM_l_gameOverLight_bg, bgLM_l_gameOverLight_playfield, bgLM_l_h001_bg, bgLM_l_h001_playfield, bgLM_l_h002_bg, bgLM_l_h002_playfield, bgLM_l_h003_bg, bgLM_l_h003_playfield, bgLM_l_h004_bg, bgLM_l_h004_playfield, bgLM_l_h005_bg, bgLM_l_h005_playfield, bgLM_l_h006_bg, bgLM_l_h006_playfield, bgLM_l_raceOverIn_bg, bgLM_l_raceOverIn_playfield, bgLM_l_tiltLight_bg, bgLM_l_tiltLight_playfield)
Dim BGbg_All: BGbg_All=Array(bgBM_bg, bgBM_hoss001, bgBM_hoss003, bgBM_hoss004, bgBM_hoss005, bgBM_hoss006, bgBM_playfield, bgLM_GI_bg, bgLM_GI_hoss001, bgLM_GI_hoss003, bgLM_GI_hoss004, bgLM_GI_hoss005, bgLM_GI_hoss006, bgLM_GI_playfield, bgLM_l_ballInPlayLight_bg, bgLM_l_ballInPlayLight_playfiel, bgLM_l_balls_bg, bgLM_l_balls_playfield, bgLM_l_bip001_bg, bgLM_l_bip001_playfield, bgLM_l_bip002_bg, bgLM_l_bip002_playfield, bgLM_l_bip003_bg, bgLM_l_bip003_playfield, bgLM_l_bip004_bg, bgLM_l_bip004_playfield, bgLM_l_bip005_bg, bgLM_l_bip005_playfield, bgLM_l_d001_bg, bgLM_l_d001_playfield, bgLM_l_d002_bg, bgLM_l_d002_playfield, bgLM_l_d003_bg, bgLM_l_d003_playfield, bgLM_l_d004_bg, bgLM_l_d004_playfield, bgLM_l_d005_bg, bgLM_l_d005_playfield, bgLM_l_d006_bg, bgLM_l_d006_playfield, _
  bgLM_l_gameOverLight_bg, bgLM_l_gameOverLight_playfield, bgLM_l_h001_bg, bgLM_l_h001_playfield, bgLM_l_h002_bg, bgLM_l_h002_playfield, bgLM_l_h003_bg, bgLM_l_h003_playfield, bgLM_l_h004_bg, bgLM_l_h004_playfield, bgLM_l_h005_bg, bgLM_l_h005_playfield, bgLM_l_h006_bg, bgLM_l_h006_playfield, bgLM_l_raceOverIn_bg, bgLM_l_raceOverIn_playfield, bgLM_l_tiltLight_bg, bgLM_l_tiltLight_playfield)
' VLM bg Arrays - End

' VLM cab Arrays - Start
' Arrays per baked part
Dim BP_cabb_lflip: BP_cabb_lflip=Array(cabBM_b_lflip)
Dim BP_cabb_lift: BP_cabb_lift=Array(cabBM_b_lift)
Dim BP_cabb_plunge: BP_cabb_plunge=Array(cabBM_b_plunge)
Dim BP_cabb_rflip: BP_cabb_rflip=Array(cabBM_b_rflip)
Dim BP_cabb_start: BP_cabb_start=Array(cabBM_b_start)
Dim BP_cabcabmetal: BP_cabcabmetal=Array(cabBM_cabmetal)
Dim BP_cabcabpaint: BP_cabcabpaint=Array(cabBM_cabpaint)
' Arrays per lighting scenario
Dim BL_cabWorld: BL_cabWorld=Array(cabBM_b_lflip, cabBM_b_lift, cabBM_b_plunge, cabBM_b_rflip, cabBM_b_start, cabBM_cabmetal, cabBM_cabpaint)
' Global arrays
Dim BGcab_Bakemap: BGcab_Bakemap=Array(cabBM_b_lflip, cabBM_b_lift, cabBM_b_plunge, cabBM_b_rflip, cabBM_b_start, cabBM_cabmetal, cabBM_cabpaint)
Dim BGcab_Lightmap: BGcab_Lightmap=Array()
Dim BGcab_All: BGcab_All=Array(cabBM_b_lflip, cabBM_b_lift, cabBM_b_plunge, cabBM_b_rflip, cabBM_b_start, cabBM_cabmetal, cabBM_cabpaint)
' VLM cab Arrays - End

' VLM pf Arrays - Start
' Arrays per baked part
Dim BP_pfPlayfield: BP_pfPlayfield=Array(pfBM_Playfield, pfLM_GI_GILight27_Playfield, pfLM_GI_GILight25_Playfield, pfLM_GI_GILight23_Playfield, pfLM_GI_GILight13_Playfield, pfLM_GI_GILight5_Playfield, pfLM_GI_GILight3_Playfield, pfLM_GI_GILight12_Playfield, pfLM_GI_GILight4_Playfield, pfLM_Ins_UpperLight0_Playfield, pfLM_Ins_L001_Playfield, pfLM_Ins_L002_Playfield, pfLM_Ins_L003_Playfield, pfLM_Ins_L004_Playfield, pfLM_Ins_L005_Playfield, pfLM_Ins_L006_Playfield, pfLM_GI_GILight6_Playfield, pfLM_GI_GILight7_Playfield, pfLM_GI_GILight11_Playfield, pfLM_GI_GILight14_Playfield, pfLM_GI_GILight15_Playfield, pfLM_GI_GILight24_Playfield, pfLM_GI_GILight26_Playfield, pfLM_Ins_UpperLight1_Playfield, pfLM_GI_GILight10_Playfield)
Dim BP_pfbulbs: BP_pfbulbs=Array(pfBM_bulbs, pfLM_GI_GILight27_bulbs, pfLM_GI_GILight25_bulbs, pfLM_GI_GILight001_bulbs, pfLM_GI_GILight23_bulbs, pfLM_GI_GILight13_bulbs, pfLM_GI_GILight5_bulbs, pfLM_GI_GILight3_bulbs, pfLM_GI_GILight12_bulbs, pfLM_GI_GILight4_bulbs, pfLM_Ins_L004_bulbs, pfLM_GI_GILight6_bulbs, pfLM_GI_GILight7_bulbs, pfLM_GI_GILight11_bulbs, pfLM_GI_GILight14_bulbs, pfLM_GI_GILight15_bulbs, pfLM_GI_GILight24_bulbs, pfLM_GI_GILight26_bulbs, pfLM_GI_GILight10_bulbs)
Dim BP_pfgate001: BP_pfgate001=Array(pfBM_gate001, pfLM_GI_GILight3_gate001, pfLM_GI_GILight4_gate001)
Dim BP_pfgate002: BP_pfgate002=Array(pfBM_gate002, pfLM_GI_GILight13_gate002, pfLM_GI_GILight5_gate002, pfLM_GI_GILight4_gate002)
Dim BP_pfgate003: BP_pfgate003=Array(pfBM_gate003, pfLM_GI_GILight5_gate003, pfLM_GI_GILight6_gate003)
Dim BP_pfgate004: BP_pfgate004=Array(pfBM_gate004, pfLM_GI_GILight6_gate004, pfLM_GI_GILight7_gate004)
Dim BP_pfgate005: BP_pfgate005=Array(pfBM_gate005, pfLM_GI_GILight6_gate005, pfLM_GI_GILight7_gate005, pfLM_GI_GILight10_gate005)
Dim BP_pflflip: BP_pflflip=Array(pfBM_lflip, pfLM_GI_GILight27_lflip, pfLM_GI_GILight25_lflip, pfLM_GI_GILight001_lflip, pfLM_GI_GILight13_lflip, pfLM_GI_GILight4_lflip, pfLM_Ins_L003_lflip, pfLM_Ins_L004_lflip, pfLM_GI_GILight11_lflip, pfLM_GI_GILight14_lflip, pfLM_GI_GILight15_lflip, pfLM_GI_GILight26_lflip)
Dim BP_pflrub: BP_pflrub=Array(pfBM_lrub, pfLM_GI_GILight27_lrub, pfLM_GI_GILight25_lrub, pfLM_GI_GILight001_lrub, pfLM_GI_GILight13_lrub, pfLM_GI_GILight4_lrub, pfLM_Ins_L002_lrub, pfLM_Ins_L003_lrub, pfLM_GI_GILight15_lrub, pfLM_GI_GILight26_lrub)
Dim BP_pfparts: BP_pfparts=Array(pfBM_parts, pfLM_GI_GILight27_parts, pfLM_GI_GILight25_parts, pfLM_GI_GILight001_parts, pfLM_GI_GILight23_parts, pfLM_GI_GILight13_parts, pfLM_GI_GILight5_parts, pfLM_GI_GILight3_parts, pfLM_GI_GILight12_parts, pfLM_GI_GILight4_parts, pfLM_Ins_UpperLight0_parts, pfLM_Ins_L001_parts, pfLM_Ins_L002_parts, pfLM_Ins_L003_parts, pfLM_Ins_L004_parts, pfLM_Ins_L005_parts, pfLM_Ins_L006_parts, pfLM_GI_GILight6_parts, pfLM_GI_GILight7_parts, pfLM_GI_GILight11_parts, pfLM_GI_GILight14_parts, pfLM_GI_GILight15_parts, pfLM_GI_GILight24_parts, pfLM_GI_GILight26_parts, pfLM_Ins_UpperLight1_parts, pfLM_GI_GILight10_parts)
Dim BP_pfplastics: BP_pfplastics=Array(pfBM_plastics, pfLM_GI_GILight23_plastics, pfLM_GI_GILight13_plastics, pfLM_GI_GILight3_plastics, pfLM_GI_GILight12_plastics, pfLM_GI_GILight4_plastics, pfLM_GI_GILight7_plastics, pfLM_GI_GILight11_plastics, pfLM_GI_GILight14_plastics, pfLM_GI_GILight15_plastics, pfLM_Ins_UpperLight1_plastics, pfLM_GI_GILight10_plastics)
Dim BP_pfrflip: BP_pfrflip=Array(pfBM_rflip, pfLM_GI_GILight27_rflip, pfLM_GI_GILight25_rflip, pfLM_GI_GILight001_rflip, pfLM_GI_GILight23_rflip, pfLM_GI_GILight13_rflip, pfLM_GI_GILight12_rflip, pfLM_GI_GILight4_rflip, pfLM_Ins_L004_rflip, pfLM_Ins_L005_rflip, pfLM_GI_GILight14_rflip, pfLM_GI_GILight24_rflip, pfLM_GI_GILight26_rflip)
Dim BP_pfrrub: BP_pfrrub=Array(pfBM_rrub, pfLM_GI_GILight27_rrub, pfLM_GI_GILight25_rrub, pfLM_GI_GILight001_rrub, pfLM_GI_GILight23_rrub, pfLM_GI_GILight4_rrub, pfLM_Ins_L004_rrub, pfLM_Ins_L005_rrub, pfLM_GI_GILight24_rrub, pfLM_GI_GILight26_rrub)
Dim BP_pfwire001: BP_pfwire001=Array(pfBM_wire001, pfLM_GI_GILight5_wire001, pfLM_GI_GILight3_wire001, pfLM_GI_GILight4_wire001, pfLM_Ins_UpperLight0_wire001)
Dim BP_pfwire002: BP_pfwire002=Array(pfBM_wire002, pfLM_GI_GILight5_wire002, pfLM_GI_GILight3_wire002, pfLM_GI_GILight4_wire002, pfLM_Ins_UpperLight0_wire002, pfLM_GI_GILight6_wire002)
Dim BP_pfwire003: BP_pfwire003=Array(pfBM_wire003, pfLM_GI_GILight5_wire003, pfLM_GI_GILight6_wire003, pfLM_GI_GILight7_wire003)
Dim BP_pfwire004: BP_pfwire004=Array(pfBM_wire004, pfLM_GI_GILight6_wire004, pfLM_GI_GILight7_wire004, pfLM_GI_GILight10_wire004)
Dim BP_pfwire005: BP_pfwire005=Array(pfBM_wire005, pfLM_GI_GILight7_wire005, pfLM_Ins_UpperLight1_wire005, pfLM_GI_GILight10_wire005)
Dim BP_pfwire006: BP_pfwire006=Array(pfBM_wire006, pfLM_Ins_L001_wire006)
Dim BP_pfwire007: BP_pfwire007=Array(pfBM_wire007, pfLM_Ins_L002_wire007)
Dim BP_pfwire008: BP_pfwire008=Array(pfBM_wire008, pfLM_GI_GILight25_wire008, pfLM_Ins_L002_wire008, pfLM_Ins_L003_wire008)
Dim BP_pfwire009: BP_pfwire009=Array(pfBM_wire009, pfLM_Ins_L004_wire009, pfLM_Ins_L005_wire009)
Dim BP_pfwire010: BP_pfwire010=Array(pfBM_wire0010, pfLM_Ins_L005_wire010, pfLM_GI_GILight26_wire010)
Dim BP_pfwire011: BP_pfwire011=Array(pfBM_wire0011, pfLM_Ins_L006_wire011)
' Arrays per lighting scenario
Dim BL_pfGI_GILight001: BL_pfGI_GILight001=Array(pfLM_GI_GILight001_bulbs, pfLM_GI_GILight001_lflip, pfLM_GI_GILight001_lrub, pfLM_GI_GILight001_parts, pfLM_GI_GILight001_rflip, pfLM_GI_GILight001_rrub)
Dim BL_pfGI_GILight10: BL_pfGI_GILight10=Array(pfLM_GI_GILight10_Playfield, pfLM_GI_GILight10_bulbs, pfLM_GI_GILight10_gate005, pfLM_GI_GILight10_parts, pfLM_GI_GILight10_plastics, pfLM_GI_GILight10_wire004, pfLM_GI_GILight10_wire005)
Dim BL_pfGI_GILight11: BL_pfGI_GILight11=Array(pfLM_GI_GILight11_Playfield, pfLM_GI_GILight11_bulbs, pfLM_GI_GILight11_lflip, pfLM_GI_GILight11_parts, pfLM_GI_GILight11_plastics)
Dim BL_pfGI_GILight12: BL_pfGI_GILight12=Array(pfLM_GI_GILight12_Playfield, pfLM_GI_GILight12_bulbs, pfLM_GI_GILight12_parts, pfLM_GI_GILight12_plastics, pfLM_GI_GILight12_rflip)
Dim BL_pfGI_GILight13: BL_pfGI_GILight13=Array(pfLM_GI_GILight13_Playfield, pfLM_GI_GILight13_bulbs, pfLM_GI_GILight13_gate002, pfLM_GI_GILight13_lflip, pfLM_GI_GILight13_lrub, pfLM_GI_GILight13_parts, pfLM_GI_GILight13_plastics, pfLM_GI_GILight13_rflip)
Dim BL_pfGI_GILight14: BL_pfGI_GILight14=Array(pfLM_GI_GILight14_Playfield, pfLM_GI_GILight14_bulbs, pfLM_GI_GILight14_lflip, pfLM_GI_GILight14_parts, pfLM_GI_GILight14_plastics, pfLM_GI_GILight14_rflip)
Dim BL_pfGI_GILight15: BL_pfGI_GILight15=Array(pfLM_GI_GILight15_Playfield, pfLM_GI_GILight15_bulbs, pfLM_GI_GILight15_lflip, pfLM_GI_GILight15_lrub, pfLM_GI_GILight15_parts, pfLM_GI_GILight15_plastics)
Dim BL_pfGI_GILight23: BL_pfGI_GILight23=Array(pfLM_GI_GILight23_Playfield, pfLM_GI_GILight23_bulbs, pfLM_GI_GILight23_parts, pfLM_GI_GILight23_plastics, pfLM_GI_GILight23_rflip, pfLM_GI_GILight23_rrub)
Dim BL_pfGI_GILight24: BL_pfGI_GILight24=Array(pfLM_GI_GILight24_Playfield, pfLM_GI_GILight24_bulbs, pfLM_GI_GILight24_parts, pfLM_GI_GILight24_rflip, pfLM_GI_GILight24_rrub)
Dim BL_pfGI_GILight25: BL_pfGI_GILight25=Array(pfLM_GI_GILight25_Playfield, pfLM_GI_GILight25_bulbs, pfLM_GI_GILight25_lflip, pfLM_GI_GILight25_lrub, pfLM_GI_GILight25_parts, pfLM_GI_GILight25_rflip, pfLM_GI_GILight25_rrub, pfLM_GI_GILight25_wire008)
Dim BL_pfGI_GILight26: BL_pfGI_GILight26=Array(pfLM_GI_GILight26_Playfield, pfLM_GI_GILight26_bulbs, pfLM_GI_GILight26_lflip, pfLM_GI_GILight26_lrub, pfLM_GI_GILight26_parts, pfLM_GI_GILight26_rflip, pfLM_GI_GILight26_rrub, pfLM_GI_GILight26_wire010)
Dim BL_pfGI_GILight27: BL_pfGI_GILight27=Array(pfLM_GI_GILight27_Playfield, pfLM_GI_GILight27_bulbs, pfLM_GI_GILight27_lflip, pfLM_GI_GILight27_lrub, pfLM_GI_GILight27_parts, pfLM_GI_GILight27_rflip, pfLM_GI_GILight27_rrub)
Dim BL_pfGI_GILight3: BL_pfGI_GILight3=Array(pfLM_GI_GILight3_Playfield, pfLM_GI_GILight3_bulbs, pfLM_GI_GILight3_gate001, pfLM_GI_GILight3_parts, pfLM_GI_GILight3_plastics, pfLM_GI_GILight3_wire001, pfLM_GI_GILight3_wire002)
Dim BL_pfGI_GILight4: BL_pfGI_GILight4=Array(pfLM_GI_GILight4_Playfield, pfLM_GI_GILight4_bulbs, pfLM_GI_GILight4_gate001, pfLM_GI_GILight4_gate002, pfLM_GI_GILight4_lflip, pfLM_GI_GILight4_lrub, pfLM_GI_GILight4_parts, pfLM_GI_GILight4_plastics, pfLM_GI_GILight4_rflip, pfLM_GI_GILight4_rrub, pfLM_GI_GILight4_wire001, pfLM_GI_GILight4_wire002)
Dim BL_pfGI_GILight5: BL_pfGI_GILight5=Array(pfLM_GI_GILight5_Playfield, pfLM_GI_GILight5_bulbs, pfLM_GI_GILight5_gate002, pfLM_GI_GILight5_gate003, pfLM_GI_GILight5_parts, pfLM_GI_GILight5_wire001, pfLM_GI_GILight5_wire002, pfLM_GI_GILight5_wire003)
Dim BL_pfGI_GILight6: BL_pfGI_GILight6=Array(pfLM_GI_GILight6_Playfield, pfLM_GI_GILight6_bulbs, pfLM_GI_GILight6_gate003, pfLM_GI_GILight6_gate004, pfLM_GI_GILight6_gate005, pfLM_GI_GILight6_parts, pfLM_GI_GILight6_wire002, pfLM_GI_GILight6_wire003, pfLM_GI_GILight6_wire004)
Dim BL_pfGI_GILight7: BL_pfGI_GILight7=Array(pfLM_GI_GILight7_Playfield, pfLM_GI_GILight7_bulbs, pfLM_GI_GILight7_gate004, pfLM_GI_GILight7_gate005, pfLM_GI_GILight7_parts, pfLM_GI_GILight7_plastics, pfLM_GI_GILight7_wire003, pfLM_GI_GILight7_wire004, pfLM_GI_GILight7_wire005)
Dim BL_pfIns_L001: BL_pfIns_L001=Array(pfLM_Ins_L001_Playfield, pfLM_Ins_L001_parts, pfLM_Ins_L001_wire006)
Dim BL_pfIns_L002: BL_pfIns_L002=Array(pfLM_Ins_L002_Playfield, pfLM_Ins_L002_lrub, pfLM_Ins_L002_parts, pfLM_Ins_L002_wire007, pfLM_Ins_L002_wire008)
Dim BL_pfIns_L003: BL_pfIns_L003=Array(pfLM_Ins_L003_Playfield, pfLM_Ins_L003_lflip, pfLM_Ins_L003_lrub, pfLM_Ins_L003_parts, pfLM_Ins_L003_wire008)
Dim BL_pfIns_L004: BL_pfIns_L004=Array(pfLM_Ins_L004_Playfield, pfLM_Ins_L004_bulbs, pfLM_Ins_L004_lflip, pfLM_Ins_L004_parts, pfLM_Ins_L004_rflip, pfLM_Ins_L004_rrub, pfLM_Ins_L004_wire009)
Dim BL_pfIns_L005: BL_pfIns_L005=Array(pfLM_Ins_L005_Playfield, pfLM_Ins_L005_parts, pfLM_Ins_L005_rflip, pfLM_Ins_L005_rrub, pfLM_Ins_L005_wire009, pfLM_Ins_L005_wire010)
Dim BL_pfIns_L006: BL_pfIns_L006=Array(pfLM_Ins_L006_Playfield, pfLM_Ins_L006_parts, pfLM_Ins_L006_wire011)
Dim BL_pfIns_UpperLight0: BL_pfIns_UpperLight0=Array(pfLM_Ins_UpperLight0_Playfield, pfLM_Ins_UpperLight0_parts, pfLM_Ins_UpperLight0_wire001, pfLM_Ins_UpperLight0_wire002)
Dim BL_pfIns_UpperLight1: BL_pfIns_UpperLight1=Array(pfLM_Ins_UpperLight1_Playfield, pfLM_Ins_UpperLight1_parts, pfLM_Ins_UpperLight1_plastics, pfLM_Ins_UpperLight1_wire005)
Dim BL_pfWorld: BL_pfWorld=Array(pfBM_Playfield, pfBM_bulbs, pfBM_gate001, pfBM_gate002, pfBM_gate003, pfBM_gate004, pfBM_gate005, pfBM_lflip, pfBM_lrub, pfBM_parts, pfBM_plastics, pfBM_rflip, pfBM_rrub, pfBM_wire001, pfBM_wire002, pfBM_wire003, pfBM_wire004, pfBM_wire005, pfBM_wire006, pfBM_wire007, pfBM_wire008, pfBM_wire009, pfBM_wire0010, pfBM_wire0011)
' Global arrays
Dim BGpf_Bakemap: BGpf_Bakemap=Array(pfBM_Playfield, pfBM_bulbs, pfBM_gate001, pfBM_gate002, pfBM_gate003, pfBM_gate004, pfBM_gate005, pfBM_lflip, pfBM_lrub, pfBM_parts, pfBM_plastics, pfBM_rflip, pfBM_rrub, pfBM_wire001, pfBM_wire002, pfBM_wire003, pfBM_wire004, pfBM_wire005, pfBM_wire006, pfBM_wire007, pfBM_wire008, pfBM_wire009, pfBM_wire0010, pfBM_wire0011)
Dim BGpf_Lightmap: BGpf_Lightmap=Array(pfLM_GI_GILight001_bulbs, pfLM_GI_GILight001_lflip, pfLM_GI_GILight001_lrub, pfLM_GI_GILight001_parts, pfLM_GI_GILight001_rflip, pfLM_GI_GILight001_rrub, pfLM_GI_GILight10_Playfield, pfLM_GI_GILight10_bulbs, pfLM_GI_GILight10_gate005, pfLM_GI_GILight10_parts, pfLM_GI_GILight10_plastics, pfLM_GI_GILight10_wire004, pfLM_GI_GILight10_wire005, pfLM_GI_GILight11_Playfield, pfLM_GI_GILight11_bulbs, pfLM_GI_GILight11_lflip, pfLM_GI_GILight11_parts, pfLM_GI_GILight11_plastics, pfLM_GI_GILight12_Playfield, pfLM_GI_GILight12_bulbs, pfLM_GI_GILight12_parts, pfLM_GI_GILight12_plastics, pfLM_GI_GILight12_rflip, pfLM_GI_GILight13_Playfield, pfLM_GI_GILight13_bulbs, pfLM_GI_GILight13_gate002, pfLM_GI_GILight13_lflip, pfLM_GI_GILight13_lrub, pfLM_GI_GILight13_parts, pfLM_GI_GILight13_plastics, pfLM_GI_GILight13_rflip, pfLM_GI_GILight14_Playfield, pfLM_GI_GILight14_bulbs, pfLM_GI_GILight14_lflip, pfLM_GI_GILight14_parts, pfLM_GI_GILight14_plastics, pfLM_GI_GILight14_rflip, _
  pfLM_GI_GILight15_Playfield, pfLM_GI_GILight15_bulbs, pfLM_GI_GILight15_lflip, pfLM_GI_GILight15_lrub, pfLM_GI_GILight15_parts, pfLM_GI_GILight15_plastics, pfLM_GI_GILight23_Playfield, pfLM_GI_GILight23_bulbs, pfLM_GI_GILight23_parts, pfLM_GI_GILight23_plastics, pfLM_GI_GILight23_rflip, pfLM_GI_GILight23_rrub, pfLM_GI_GILight24_Playfield, pfLM_GI_GILight24_bulbs, pfLM_GI_GILight24_parts, pfLM_GI_GILight24_rflip, pfLM_GI_GILight24_rrub, pfLM_GI_GILight25_Playfield, pfLM_GI_GILight25_bulbs, pfLM_GI_GILight25_lflip, pfLM_GI_GILight25_lrub, pfLM_GI_GILight25_parts, pfLM_GI_GILight25_rflip, pfLM_GI_GILight25_rrub, pfLM_GI_GILight25_wire008, pfLM_GI_GILight26_Playfield, pfLM_GI_GILight26_bulbs, pfLM_GI_GILight26_lflip, pfLM_GI_GILight26_lrub, pfLM_GI_GILight26_parts, pfLM_GI_GILight26_rflip, pfLM_GI_GILight26_rrub, pfLM_GI_GILight26_wire010, pfLM_GI_GILight27_Playfield, pfLM_GI_GILight27_bulbs, pfLM_GI_GILight27_lflip, pfLM_GI_GILight27_lrub, pfLM_GI_GILight27_parts, pfLM_GI_GILight27_rflip, pfLM_GI_GILight27_rrub, _
  pfLM_GI_GILight3_Playfield, pfLM_GI_GILight3_bulbs, pfLM_GI_GILight3_gate001, pfLM_GI_GILight3_parts, pfLM_GI_GILight3_plastics, pfLM_GI_GILight3_wire001, pfLM_GI_GILight3_wire002, pfLM_GI_GILight4_Playfield, pfLM_GI_GILight4_bulbs, pfLM_GI_GILight4_gate001, pfLM_GI_GILight4_gate002, pfLM_GI_GILight4_lflip, pfLM_GI_GILight4_lrub, pfLM_GI_GILight4_parts, pfLM_GI_GILight4_plastics, pfLM_GI_GILight4_rflip, pfLM_GI_GILight4_rrub, pfLM_GI_GILight4_wire001, pfLM_GI_GILight4_wire002, pfLM_GI_GILight5_Playfield, pfLM_GI_GILight5_bulbs, pfLM_GI_GILight5_gate002, pfLM_GI_GILight5_gate003, pfLM_GI_GILight5_parts, pfLM_GI_GILight5_wire001, pfLM_GI_GILight5_wire002, pfLM_GI_GILight5_wire003, pfLM_GI_GILight6_Playfield, pfLM_GI_GILight6_bulbs, pfLM_GI_GILight6_gate003, pfLM_GI_GILight6_gate004, pfLM_GI_GILight6_gate005, pfLM_GI_GILight6_parts, pfLM_GI_GILight6_wire002, pfLM_GI_GILight6_wire003, pfLM_GI_GILight6_wire004, pfLM_GI_GILight7_Playfield, pfLM_GI_GILight7_bulbs, pfLM_GI_GILight7_gate004, pfLM_GI_GILight7_gate005, _
  pfLM_GI_GILight7_parts, pfLM_GI_GILight7_plastics, pfLM_GI_GILight7_wire003, pfLM_GI_GILight7_wire004, pfLM_GI_GILight7_wire005, pfLM_Ins_L001_Playfield, pfLM_Ins_L001_parts, pfLM_Ins_L001_wire006, pfLM_Ins_L002_Playfield, pfLM_Ins_L002_lrub, pfLM_Ins_L002_parts, pfLM_Ins_L002_wire007, pfLM_Ins_L002_wire008, pfLM_Ins_L003_Playfield, pfLM_Ins_L003_lflip, pfLM_Ins_L003_lrub, pfLM_Ins_L003_parts, pfLM_Ins_L003_wire008, pfLM_Ins_L004_Playfield, pfLM_Ins_L004_bulbs, pfLM_Ins_L004_lflip, pfLM_Ins_L004_parts, pfLM_Ins_L004_rflip, pfLM_Ins_L004_rrub, pfLM_Ins_L004_wire009, pfLM_Ins_L005_Playfield, pfLM_Ins_L005_parts, pfLM_Ins_L005_rflip, pfLM_Ins_L005_rrub, pfLM_Ins_L005_wire009, pfLM_Ins_L005_wire010, pfLM_Ins_L006_Playfield, pfLM_Ins_L006_parts, pfLM_Ins_L006_wire011, pfLM_Ins_UpperLight0_Playfield, pfLM_Ins_UpperLight0_parts, pfLM_Ins_UpperLight0_wire001, pfLM_Ins_UpperLight0_wire002, pfLM_Ins_UpperLight1_Playfield, pfLM_Ins_UpperLight1_parts, pfLM_Ins_UpperLight1_plastics, pfLM_Ins_UpperLight1_wire005)
Dim BGpf_All: BGpf_All=Array(pfBM_Playfield, pfBM_bulbs, pfBM_gate001, pfBM_gate002, pfBM_gate003, pfBM_gate004, pfBM_gate005, pfBM_lflip, pfBM_lrub, pfBM_parts, pfBM_plastics, pfBM_rflip, pfBM_rrub, pfBM_wire001, pfBM_wire002, pfBM_wire003, pfBM_wire004, pfBM_wire005, pfBM_wire006, pfBM_wire007, pfBM_wire008, pfBM_wire009, pfBM_wire0010, pfBM_wire0011, pfLM_GI_GILight001_bulbs, pfLM_GI_GILight001_lflip, pfLM_GI_GILight001_lrub, pfLM_GI_GILight001_parts, pfLM_GI_GILight001_rflip, pfLM_GI_GILight001_rrub, pfLM_GI_GILight10_Playfield, pfLM_GI_GILight10_bulbs, pfLM_GI_GILight10_gate005, pfLM_GI_GILight10_parts, pfLM_GI_GILight10_plastics, pfLM_GI_GILight10_wire004, pfLM_GI_GILight10_wire005, pfLM_GI_GILight11_Playfield, pfLM_GI_GILight11_bulbs, pfLM_GI_GILight11_lflip, pfLM_GI_GILight11_parts, pfLM_GI_GILight11_plastics, pfLM_GI_GILight12_Playfield, pfLM_GI_GILight12_bulbs, pfLM_GI_GILight12_parts, pfLM_GI_GILight12_plastics, pfLM_GI_GILight12_rflip, pfLM_GI_GILight13_Playfield, pfLM_GI_GILight13_bulbs, _
  pfLM_GI_GILight13_gate002, pfLM_GI_GILight13_lflip, pfLM_GI_GILight13_lrub, pfLM_GI_GILight13_parts, pfLM_GI_GILight13_plastics, pfLM_GI_GILight13_rflip, pfLM_GI_GILight14_Playfield, pfLM_GI_GILight14_bulbs, pfLM_GI_GILight14_lflip, pfLM_GI_GILight14_parts, pfLM_GI_GILight14_plastics, pfLM_GI_GILight14_rflip, pfLM_GI_GILight15_Playfield, pfLM_GI_GILight15_bulbs, pfLM_GI_GILight15_lflip, pfLM_GI_GILight15_lrub, pfLM_GI_GILight15_parts, pfLM_GI_GILight15_plastics, pfLM_GI_GILight23_Playfield, pfLM_GI_GILight23_bulbs, pfLM_GI_GILight23_parts, pfLM_GI_GILight23_plastics, pfLM_GI_GILight23_rflip, pfLM_GI_GILight23_rrub, pfLM_GI_GILight24_Playfield, pfLM_GI_GILight24_bulbs, pfLM_GI_GILight24_parts, pfLM_GI_GILight24_rflip, pfLM_GI_GILight24_rrub, pfLM_GI_GILight25_Playfield, pfLM_GI_GILight25_bulbs, pfLM_GI_GILight25_lflip, pfLM_GI_GILight25_lrub, pfLM_GI_GILight25_parts, pfLM_GI_GILight25_rflip, pfLM_GI_GILight25_rrub, pfLM_GI_GILight25_wire008, pfLM_GI_GILight26_Playfield, pfLM_GI_GILight26_bulbs, _
  pfLM_GI_GILight26_lflip, pfLM_GI_GILight26_lrub, pfLM_GI_GILight26_parts, pfLM_GI_GILight26_rflip, pfLM_GI_GILight26_rrub, pfLM_GI_GILight26_wire010, pfLM_GI_GILight27_Playfield, pfLM_GI_GILight27_bulbs, pfLM_GI_GILight27_lflip, pfLM_GI_GILight27_lrub, pfLM_GI_GILight27_parts, pfLM_GI_GILight27_rflip, pfLM_GI_GILight27_rrub, pfLM_GI_GILight3_Playfield, pfLM_GI_GILight3_bulbs, pfLM_GI_GILight3_gate001, pfLM_GI_GILight3_parts, pfLM_GI_GILight3_plastics, pfLM_GI_GILight3_wire001, pfLM_GI_GILight3_wire002, pfLM_GI_GILight4_Playfield, pfLM_GI_GILight4_bulbs, pfLM_GI_GILight4_gate001, pfLM_GI_GILight4_gate002, pfLM_GI_GILight4_lflip, pfLM_GI_GILight4_lrub, pfLM_GI_GILight4_parts, pfLM_GI_GILight4_plastics, pfLM_GI_GILight4_rflip, pfLM_GI_GILight4_rrub, pfLM_GI_GILight4_wire001, pfLM_GI_GILight4_wire002, pfLM_GI_GILight5_Playfield, pfLM_GI_GILight5_bulbs, pfLM_GI_GILight5_gate002, pfLM_GI_GILight5_gate003, pfLM_GI_GILight5_parts, pfLM_GI_GILight5_wire001, pfLM_GI_GILight5_wire002, pfLM_GI_GILight5_wire003, _
  pfLM_GI_GILight6_Playfield, pfLM_GI_GILight6_bulbs, pfLM_GI_GILight6_gate003, pfLM_GI_GILight6_gate004, pfLM_GI_GILight6_gate005, pfLM_GI_GILight6_parts, pfLM_GI_GILight6_wire002, pfLM_GI_GILight6_wire003, pfLM_GI_GILight6_wire004, pfLM_GI_GILight7_Playfield, pfLM_GI_GILight7_bulbs, pfLM_GI_GILight7_gate004, pfLM_GI_GILight7_gate005, pfLM_GI_GILight7_parts, pfLM_GI_GILight7_plastics, pfLM_GI_GILight7_wire003, pfLM_GI_GILight7_wire004, pfLM_GI_GILight7_wire005, pfLM_Ins_L001_Playfield, pfLM_Ins_L001_parts, pfLM_Ins_L001_wire006, pfLM_Ins_L002_Playfield, pfLM_Ins_L002_lrub, pfLM_Ins_L002_parts, pfLM_Ins_L002_wire007, pfLM_Ins_L002_wire008, pfLM_Ins_L003_Playfield, pfLM_Ins_L003_lflip, pfLM_Ins_L003_lrub, pfLM_Ins_L003_parts, pfLM_Ins_L003_wire008, pfLM_Ins_L004_Playfield, pfLM_Ins_L004_bulbs, pfLM_Ins_L004_lflip, pfLM_Ins_L004_parts, pfLM_Ins_L004_rflip, pfLM_Ins_L004_rrub, pfLM_Ins_L004_wire009, pfLM_Ins_L005_Playfield, pfLM_Ins_L005_parts, pfLM_Ins_L005_rflip, pfLM_Ins_L005_rrub, pfLM_Ins_L005_wire009, _
  pfLM_Ins_L005_wire010, pfLM_Ins_L006_Playfield, pfLM_Ins_L006_parts, pfLM_Ins_L006_wire011, pfLM_Ins_UpperLight0_Playfield, pfLM_Ins_UpperLight0_parts, pfLM_Ins_UpperLight0_wire001, pfLM_Ins_UpperLight0_wire002, pfLM_Ins_UpperLight1_Playfield, pfLM_Ins_UpperLight1_parts, pfLM_Ins_UpperLight1_plastics, pfLM_Ins_UpperLight1_wire005)
' VLM pf Arrays - End



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



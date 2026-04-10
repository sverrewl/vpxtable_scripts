'****************************************************************
'
'         Rockmakers (Bally 1968)
'          Script by Scottacus
'             v 0.60
'           November 2023
'
' Basic DOF config
'   90  BackGlass On Image
'   101 Left Flipper, 102 Right Flipper
'   103 Left Sling, 104 Right Sling
'   105 Bumper1, 106 Bumper2, 107 Bumper3
'   125 Ball Release
'   142 Chime 10 & 100
'   170 - 173 Mushrooms
'   199 credit light
'   128 Knocker
'
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

Const cGameName = "Rockmakers_1968"
Const cOptions = "Rockmakers(Bally1968)v0.60.txt"
Const hsFileName = "Rockmakers (Bally 1968)"

Dim replayLevel, lowRaR
'********************** User Settings For High Scores and Bonus*************************
' Choose one value for the High Scores
' replayLevel = 0 | replayLevel = 1 | replayLevel = 2
'3 Ball   5 Ball  | 3 Ball  5 Ball  | 3 ball  5 ball
' 500   1700  |  700  3700  |  900  5700
' 800   2900  |  900  4700  | 1000  6700
' 900   2000  | 1000  5700  | 1100  7700
'1000   2100  | 1100  6700  | 1200  8700

replayLevel = 2

' Set a low Rock-A-Rock value from 3-8 for payout of either Add-A-Ball or Replay
' Note this will also have a second value of 9 added automatically
lowRaR = 6

Dim vrOption
'*************************** VR Room****************************************************
'The following two lines of code are two ways of controling which mode the table should
'be run in (Cab/DT, VR Minimal, VR Room). With the first line uncommented, the software
'will attempt to determine if you are running VPVR and if you have set it up to VR Enable.
'If you have not done this then it will run the table in DT/Cab mode.  Should you experience
'problems with this, just comment out the line "GetAskkToTurnOn", uncomment the following
'line of code and change the value of vrOption to the desired setting.

'GetAskToTurnOn
vrOption = 2 ' 0 - VR Room off, 1 - VR Cab On, 2- VR Cab and Room On

'*************************** Glass Scratches *******************************************
Const VRGlassScratches = 0  ' set this to 1 to turn on VR glass scratches
If vrOption > 0 and VRGlassScratches > 0 then GlassImpurities.visible = true

'*************************** PinballY Settings *****************************************
'Set this variable to 1 to save a PinballY High Score file to your Tables Folder
'this will let the Pinball Y front end display the high scores when searching for tables
'0 = No PinballY High Scores, 1 = Save PinballY High Scores
Const cPinballY = 1

'**************************** Enable Flex DMD **************************************
UseFlexDMD = False  '  Change to True to use a FlexDMD

Dim tablewidth: tablewidth = Table1.width
Dim tableheight: tableheight = Table1.height

Sub frameTimer_Timer()
  FlipperVisualUpdate       'update flipper shadows and primitives
End Sub

' The game timer interval is 10 ms
Sub gameTimer_Timer()
  Cor.Update            'update ball tracking (this sometimes goes in the RDampen_Timer sub)
End Sub

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
Dim RaRReel(4)
Dim state
Dim tilt
Dim matchNumber
Dim i,j, f, ii, Object, Light, x, y, z, xx
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
Dim powerSetting : powerSetting = 0
Dim difficultySetting : difficultySetting = 0

'Desktop reel pans,  Valid values: "Lpan", "Mpan" and "Rpan"
Const dtcrReel = "Lpan"
Const dts1Reel = "Rpan"
Const dts2Reel = "Rpan"
Const dts3Reel = "Rpan"
Const dts4Reel = "Rpan"
Const dtRaRReel = "Rpan"

'Backglass reel pans,  Valid values: "Lpan", "Mpan" and "Rpan"
Const bgcrReel = "Mpan"
Const bgs1Reel = "Lpan"
Const bgs2Reel = "Rpan"
Const bgs3Reel = "Lpan"
Const bgs4Reel = "Rpan"
Const bgRaRReel = "Rpan"

Sub Table1_init
  LoadEM
  maxPlayers = 4

  If vrOption > 0 Then
    dtReel.setValue(0)
    BM_rails.visible = 0
  Else
    dtReel.setValue(1)
    BM_rails.visible = 1
  End If

  For x = 1 to maxPlayers
    Set sReel(x) = EVAL("scoreReel" & x)
    Set RaRReel(x) = EVAL("dtRaRReel" & x)
  Next

  player = 1
  loadHighScore

  ballsInPlay = 0

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
  If difficultySetting = "" Then difficultySetting = 0
  If balls = "" Then balls = 5
  If chime = "" Then chime = 0
  If onBumper = "" Then onBumper = 1
  If lutValue = "" Then lutValue = 0

  replaySettings
  Table1.ColorGradeImage = "LUT" & lutValue

  If UseFlexDMD = True Then
    DMD_Init
  End If

  dim bp

  firstBallOut = 0
  updatePostIt
  dynamicUpdatePostIt.enabled = 1
  TiltReel.setValue(1)
  If vrOption > 0 Then vrTilt.visible = 1
  CreditReel.setvalue(credit)

  For each BP in BP_german_version : BP.visible = 0: Next

  If ShowDT = True Then
    For each object in backdropstuff
    Object.visible = 1
    Next
    'VR_cab_siderails.visible = 0
    'vr_cab.visible = 0
    BM_rails.visible=0
  End If

  If ShowDt = False Then
    For each object in backdropstuff
    Object.visible = 0
    Next
  End If

  for x = 1 to maxPlayers
    RaRreelDone(x) = 0
    reelDone(x) = 0
  Next

  BootTable.enabled = 1

  reelStop = 0

  MatchReel.setValue(MatchNumber) 'Need to set to change if 1 point table

  tilt = False
  state = False
  gameState

  For each bp in BP_flipL_001 : bp.visible=1 : Next
  For each bp in BP_flipL_002 : bp.visible=0 : Next
  For each bp in BP_flipR_001 : bp.visible=1 : Next
  For each bp in BP_flipR_002 : bp.visible=0 : Next

  'Initialize slings
  RStep = 0:RightSlingShot.Timerenabled=True
  LStep = 0:LeftSlingShot.Timerenabled=True
  L1Step = 0:LeftSlingShotUpper.Timerenabled=True

  If difficultySetting = 1 Then
    For each BP in BP_DifficultyA : BP.visible = 1: Next
    For each BP in BP_DifficultyB : BP.visible = 0: Next
    phys_postsdiffA.collidable = 1
  Else
    For each BP in BP_DifficultyB : BP.visible = 1: Next
    For each BP in BP_DifficultyA : BP.visible = 0: Next
    phys_postsdiffA.collidable = 0
  End If

End Sub

'***********KeyCodes
Dim enableInitialEntry, firstBallOut, replayValue
replayValue = 0
Sub Table1_KeyDown(ByVal keycode)
  dim BP

  If enableInitialEntry = True Then enterIntitals(keycode)

  If keycode = addCreditKey Then
    playFieldSound "coinin",0,Drain,1
    addCredit = 1
    scoreMotor5.enabled = 1
    End If

    If keycode = startGameKey Then
    StartButton.y = 1989.833 - 5
    If enableInitialEntry = False and operatormenu = 0 and backGlassOn = 1 Then
      If freePlay = 1 and players < 4 and firstBallOut = 0 Then startGame
      If freePlay = 0 and credit > 0 and players < 4 and firstBallOut = 0 Then
        credit = credit - 1
        If showDT = False Then PlayReelSound "Reel5", bgcrReel Else PlayReelSound "Reel5", dtcrReel
        creditReel.setvalue(credit)
        If B2SOn Then
          If freeplay = 0 Then controller.B2SSetCredits credit
          If freePlay = 0 and credit < 1 Then DOF 199, DOFOff
        End If
        startGame
      End If
    End If
  End If

  If keycode = rightFlipperKey Then
    If vrOption > 0 Then VRFlipperButtonRight.X = 2114 - 8
  End If

  If keycode = leftFlipperKey and contball = 0 Then
    If vrOption > 0 Then VRFlipperButtonLeft.X = 2105 + 8
  End If

  If keycode = PlungerKey Then
    plunger.PullBack
    playFieldSound "plungerpull", 0, plunger, 1
    If vrOption > 0 Then
      TimerVRPlunger.Enabled = True
      TimerVRPlunger1.Enabled = False
      PinCab_Shooter.Y = -484.9259
    End If
  End If

  If tilt = False and state = True Then
  If keycode = leftFlipperKey and contball = 0 Then
    FlipperActivate LeftFlipper, LFPress
    LeftFlipper2.rotateToEnd
    lf.fire
    If Zipped = 0 Then
      lf.fire
    Else
      lf1.fire
    End If
    playFieldSound "FlipUpL", 0, leftFlipper, 1
    If B2SOn Then DOF 101,DOFOn
    playFieldSound "FlipBuzzL", -1, leftFlipper, 1
  End If

  If keycode = RightFlipperKey and contball = 0 Then
    FlipperActivate RightFlipper, RFPress
    rf.fire
    If Zipped = 0 Then
      rf.fire
    Else
      rf1.fire
    End If
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
    If vrOption = 0 Then If options = 6 Then options = 7
    If showDt = True Then If options = 7 Then options = 8 'skips non DT options
        If options = 9 Then options = 0
    optionMenu.visible = True
        playFieldSound "target", 0, SoundPointScoreMotor, 1.5
        Select Case (Options)
            Case 0:
                optionMenu.image = "FreeCoin" & freePlay
            Case 1:
                optionMenu.image = balls & "Balls"
      Case 2:
                optionMenu.image = "replayEB" & replayValue
      Case 3:
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
      Case 4:
        For x = 1 to 6
          EVAL("Speaker" & x).visible = 0
        Next
        optionMenu1.visible = 0
        OptionMenu2.visible = 0
        optionMenu.image = "difficulty" & difficultySetting
      Case 5:
        optionMenu.image = "coilTap" & powerSetting
      Case 6:
        optionMenu.image = "VR" & vrOption
      Case 7:
        optionMenu2.visible = 0
        For x = 1 to 6
          EVAL("Speaker" & x).visible = 0
        Next
        optionMenu1.image = "DOF"
        optionMenu.image = "Chime" & chime
      Case 8:
        For x = 1 to 6
          EVAL("Speaker" & x).visible = 0
        Next
        optionMenu1.visible = 0
        optionMenu2.visible = 0
        optionMenu.image = "SaveExit"
    End Select
    End If

    If keycode = RightFlipperKey and state = False and operatorMenu = 1 Then
        playFieldSound "metalHitHigh", 0, SoundPointScoreMotor, 0.2
      Select Case (options)
    Case 0:
            If freePlay = 0 Then
                freePlay = 1
        If B2SOn Then DOF 199, DOFOn
              Else
                freePlay = 0
            End If
            optionMenu.image= "FreeCoin" & freePlay
      If freePlay = 0 Then
        If credit > 0 and B2SOn Then DOF 199, DOFOn
        If credit < 1 and B2SOn Then DOF 199, DOFOff
      Else
        If B2SOn Then DOF 199, DOFOn
      End If
        Case 1:
            If balls = 3 Then
                balls = 5
              Else
                balls = 3
            End If
      replaySettings
            optionMenu.image = balls & "Balls"
        Case 2:
            replayValue = replayValue + 1
      If replayValue > 1 Then replayValue = 0
      replaySettings
            optionMenu.image = "replayEB" & replayValue
    Case 3:
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
    Case 4:
      difficultySetting = difficultySetting + 1
      If difficultySetting > 1 Then difficultySetting = 0
      optionMenu.image = "difficulty" & difficultySetting
      If difficultySetting = 1 Then
        For each BP in BP_DifficultyA : BP.visible = 1: Next
        For each BP in BP_DifficultyB : BP.visible = 0: Next
        phys_postsdiffA.collidable = 1
      Else
        For each BP in BP_DifficultyB : BP.visible = 1: Next
        For each BP in BP_DifficultyA : BP.visible = 0: Next
        phys_postsdiffA.collidable = 0
      End If
    Case 5:
      powerSetting = powerSetting + 1
      If powerSetting > 1 Then powerSetting = 0
      optionMenu.image = "coilTap" & powerSetting
      If powerSetting = 1 Then
        bumper1.force = 6
        bumper2.force = 6
        bumper3.force = 6
        LeftSlingShot.collidable = 0
        RightSlingShot.collidable = 0
        LeftSlingShotUpper.collidable = 0
        LeftSlingShotLP.collidable = 1
        RightSlingShotLP.collidable = 1
        LeftSlingShotUpperLP.collidable = 1
      Else
        bumper1.force = 9
        bumper2.force = 9
        bumper3.force = 9
        LeftSlingShot.collidable = 1
        RightSlingShot.collidable = 1
        LeftSlingShotUpper.collidable = 1
        LeftSlingShotLP.collidable = 0
        RightSlingShotLP.collidable = 0
        LeftSlingShotUpperLP.collidable = 0
      End If

    Case 6:
      If vrOption = 1 Then
        vrOption = 2
        For Each object in vrRoom: object.visible = 1: Next
      Else
        vrOption = 1
        For Each object in vrRoom: object.visible = 0: Next
      End If
      optionMenu.image = "VR" & vrOption
        Case 7:
            If chime = 0 Then
                chime= 1
        If B2SOn Then DOF 142,DOFPulse
              Else
                chime = 0
        pts10
            End If
      optionMenu.image = "Chime" & chime
        Case 8:
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
  If Keycode = 30 Then 'a' key

    replayLevel = replayLevel + 1
    if replayLevel > 6 Then replayLevel = 1
    if replayLevel < 4 Then
      Balls = 3
    Else
      Balls = 5
    End If
    replaySettings
    tb.text = replayLevel
  End If

  If Keycode= 31 Then 's' key
    lowRaR = lowRaR+ 1
    If lowRaR > 9 Then lowRaR = 3
    tb.text = lowRaR
    replaySettings
  End If

  If Keycode = 33 Then 'f' key
    If Balls = 3 Then
      Balls = 5
    Else
      Balls = 3
    End If
    replaySettings
  End If

  If Keycode = 34 Then 'g' key
    replayValue = replayValue + 1
    If replayValue > 1 Then replayValue = 0
    replaySettings
  End If
'************************End Of Test Keys****************************
End Sub

dim TestChr1, TestChr2, TestChr3, TestChr, cardValue

Sub Table1_KeyUp(ByVal keycode)

  If keycode = rightFlipperKey Then
    If vrOption > 0 Then VRFlipperButtonRight.X = 2114
  End If

  If keycode = leftFlipperKey and contball = 0 Then
    LeftFlipper2.rotateToStart
    If vrOption > 0 Then VRFlipperButtonLeft.X = 2105
  End If

  If keycode = StartGameKey Then
    StartButton.y = 1989.833
  End If

  If keycode = plungerKey Then
    If vrOption > 0 Then
      TimerVRPlunger.Enabled = False
      TimerVRPlunger1.Enabled = True
      PinCab_Shooter.Y = -484.9259
    End If
    plunger.Fire
    playFieldSound "PlungerFire", 0, plunger, 1
  End If


    If keycode = leftFlipperKey Then
        operatorMenuTimer.Enabled = False
    End If

   If tilt = False and state = True Then
    If keycode = leftFlipperKey and contball = 0 Then
      FlipperDeActivate LeftFlipper, LFPress
      lfpress = 0
      LeftFlipper.eosTorqueAngle = EOSA
      LeftFlipper.eosTorque = EOST
      LeftFlipper.RotateToStart
      LeftFlipper1.RotateToStart
      LeftFlipper2.RotateToStart
      playFieldSound "FlipDownL", 0, leftFlipper, 1
      If B2SOn Then DOF 101,DOFOff
      stopSound "FlipBuzzLA"
      stopSound "FlipBuzzLB"
      stopSound "FlipBuzzLC"
      stopSound "FlipBuzzLD"
    End If

    If keycode = RightFlipperKey and contball = 0 Then
      FlipperDeActivate RightFlipper, RFPress
      rfpress = 0
      RightFlipper.eosTorqueAngle = EOSA
      RightFlipper.eosTorque = EOST
      RightFlipper.rotateToStart
      RightFlipper1.RotateToStart
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
Sub bootTable_Timer()
  bootCount = bootCount + 1
  If bootCount = 1 Then
    If vrOption > 0 Then vrBackGlass.image = "rmOn"
    '*****GI Lights On
    For each xx in GI:xx.State = 1: Next
    If B2SOn Then
      controller.B2SSetCredits credit
      controller.B2SSetMatch 34, matchNumber
      controller.B2SSetTilt 33,1
      controller.B2SSetData 80, 1 'turns the backglass image on
      for x = 0 to 3
        controller.B2SSetData 61+x,0
      next
      If credit > 0 Then DOF 199, DOFOn
      If freePlay = 1 Then DOF 199, DOFOn
    End If

    For x = 1 to maxPlayers
      If B2SOn Then controller.B2SSetScorePlayer x, score(x)
      RaRReel(x).setvalue(RaRBonus(x))
      If B2SOn Then Controller.B2SSetReel (x+16), RaRBonus(x)

      sReel(x).setvalue(score(x))
    Next

    backGlassOn = 1
    me.enabled = False
  End If
End Sub

'***********Replay Settings
Dim hsCardValue
Sub replaySettings
  If balls = 5 Then
    If replayLevel = 0  Then
      replay(1) = 1700: replay(2) = 1900: replay(3) = 2000: replay(4) = 2100
      hsCardValue = 4
    Elseif replayLevel = 1 Then
      replay(1) = 3700: replay(2) = 4700: replay(3) = 5700: replay(4) = 6700
      hsCardValue = 5
    Else
      replay(1) = 4400: replay(2) = 5600: replay(3) = 6200: replay(4) = 6800
      hsCardValue = 6
    End If
  Else
    If replayLevel = 0 Then
      replay(1) = 500: replay(2) = 800: replay(3) = 900: replay(4) = 1000
      hsCardValue = 1
    Elseif replayLevel = 1 Then
      replay(1) = 700: replay(2) = 900: replay(3) = 1000: replay(4) = 1100
      hsCardValue = 2
    Else
      replay(1) = 2400: replay(2) = 3600: replay(3) = 4200: replay(4) = 4800
      hsCardValue = 3
    End If
  End If
  RaRReplay(1) = lowRaR
  RaRReplay(2) = 9
  For Each xx in CardFlashers: xx.visible = 0: Next
  EVAL("FlasherHS" & hsCardValue).visible = 1
  FlasherMatch.visible = 1
  If replayValue = 0 Then
    If lowRaR < 9 Then
      FlasherAAB1.visible = 1
    Else
      FlasherAAB2.visible = 1
    End If
  Else
    If lowRaR < 9 Then
      Flasherreplay1.visible = 1
    Else
      Flasherreplay2.visible = 1
    End If
  End If
  If Balls = 3 Then
    flasherballs1.visible = 1
  Else
    flasherballs2.visible = 1
  End If
  Select Case lowRaR
    Case 3: FlasherRaR3Base.visible = 1
    Case 4: FlasherRaR4Base.visible = 1
    Case 5: FlasherRaR5Base.visible = 1
    Case 6: FlasherRaR6Base.visible = 1
    Case 7: FlasherRaR7Base.visible = 1
    Case 8: FlasherRaR8Base.visible = 1
  End Select
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
Dim ballInPlay, bonusScore
Sub startGame
  If state = False Then
    newGame
    ballInPlay = 1
    If B2SOn Then
      controller.B2SSetCredits credit
      controller.B2SSetBallinPlay 32, ballInPlay
      controller.B2SSetCanPlay 31, 1
      controller.B2SSetData 61, 1
      controller.B2SSetGameOver 0
    End If
    UP1.setValue(1)
    BIPReel.setValue(ballInPlay)
    dynamicUpdatePostIt.enabled = 0
    updatePostIt
    tilt = False
    state = True
    gameState
    players = 1
    CanPlay1.setValue(1)
    for x = 1 to maxPlayers
      score(x) = (score(x) mod 100000)
    Next
    If vrOption > 0 Then
      vrBipSet
      vrCredit
      vrCp.image = "cp1"
      vrGameOver.visible = 0
      vrup.image = "up1"
    End If
  Else If state = True and players < maxPlayers and ballinPlay = 1 Then
    players = players + 1
    creditReel.setvalue(credit)
    If vrOption > 0 Then
      vrCredit
      vrCanPlay
    End If
    If B2SOn Then
      controller.B2SSetCredits credit
      controller.B2SSetCanPlay 31, players
    End If
    EVAL("CanPlay" & players).setValue(1)
    EVAL("CanPlay" & players - 1).setValue(0)
    End If
  End If
End Sub

'*********New Game
Dim endGame
Sub newGame
  player = 1
    endGame = 0
  gameState
  For x = 1 to maxPlayers
    EVAL("CanPlay" & x).setValue(0)
    EVAL("UP" & x).SetValue(0)
  Next
  If B2SOn Then
    For x = 1 to players
      controller.B2SSetData 60+x, 0
      controller.B2SSetCanPlay 31, 0
    Next
  End If
  CanPlay1.setValue(1)
  For f = 1 to 3
    EVAL("Bumper" & f).hashitevent = 1
  Next
  If vrOption > 0 Then vrMatch.image = "m0"
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
    match
    UnzipFlippers
    For x = 1 to players
      EVAL("CanPlay" & x).setValue(1)
      EVAL("UP" & x).SetValue(1)
    Next
    state = False
    BIPReel.setValue(0)

    gameState
    DynamicUpdatePostIt.enabled = 1
    sortScores
    checkHighScores
    firstBallOut = 0

    For x = 1 to maxPlayers
      rep(x) = 0
      repAwarded(x) = 0
    Next
    saveHighScore

    If B2SOn Then
      controller.B2SSetGameOver 35,1
      controller.B2SSetballinplay 32, 0
      For x = 1 to players
        controller.B2SSetData 60+x, 1
        controller.B2SSetCanPlay 31, x
      Next
      If credit > 0 Then DOF 199, DOFOn
      If freePlay = 1 Then DOF 199, DOFOn
    End If
    BIPReel.setValue(0)
    players = 0
    If vrOption > 0 Then
      vrGameOver.visible = 1
      vrBip.image = "bip0"
      vrCp.image = "cp0"
      vrUp.image = "up0"
    End If
  Else
    relBall = 1   'variable to tell score motor routine that this is a ball release run
    scoreMotor5.enabled = 1
    BIPReel.setValue(ballinPlay)
    FreeBallGateClose
    UnzipFlippers
    If vrOption > 0 Then vrBipSet
  End If
End Sub

'***************Drain and Release Ball
Dim drainActive
Sub drain_Hit()
  repAwarded(Player) = 0
  drainActive = 1
  drain.DestroyBall 'this is used for multiball tables to avoid timing issues with balls moving in the trough
  scoreBonus
End Sub

Dim launched, ballsInPlay
Sub releaseBall
  drain.CreateSizedBallWithMass Ballsize/2, BallMass  'used for multiball tables (see above)
  ballsInPlay = ballsInPlay + 1
  playFieldSound "FastKickIntoLaunchLane", 0, drain, 0.5
  If B2SOn Then DOF 125, DOFPulse
  drain.kick 60,20    'use only this and no create/destroy ball on single ball tables
  launched = 0
  If B2SOn Then Controller.B2SSetBallinPlay 32, ballinPlay
  drainActive = 0
  BIPReel.setValue(ballInPlay)
  If vrOption > 0 Then vrBipSet
End Sub

'**********Check if Scoring Bonus is True
Dim bonusFlag
Sub scoreBonus
  If LShootAgain.state = 0 Then
    rockArockReset.enabled = 1
    L100a.state = 0: L100b.state = 0: L200.state = 0
    AdvancePlayers
  Else
    releaseBall
  End If
End Sub

'**********Advance Players
Sub advancePlayers
   If players = 1 or player = players Then
    player = 1
    EVAL("UP" & players).SetValue(0)
    Up1.SetValue(1)
    If B2SOn Then
      controller.B2SSetData 61, 1
      controller.B2SSetData 60+players, 0
    End If
   Else
    player = player + 1
    EVAL ("Up" & (player - 1)).setValue(0)
    EVAL ("Up" & player).setValue(1)
    If B2SOn Then
      controller.B2SSetData 60+player, 1
      controller.B2SSetData 60+player-1, 0
    End If
  End If


  If vrOption > 0 Then vrPlayerUp
  nextBall
End Sub

'**********Next Ball
Sub nextBall
  LeftSlingShot.collidable = 1
  LeftSlingShotUpper.collidable = 1
  RightSlingShot.collidable = 1
    If tilt = True Then
    For f = 1 to 3  'set to number of bumpers
      EVAL("Bumper" & f).hashitevent = 1
    Next
    tilt = False
    TiltReel.setValue(0)
    If vrOption > 0 Then vrTilt.visible = 0
    If B2SOn Then
      controller.B2SSetTilt 33,0
      controller.B2SSetData 1, 1
    End If
    End If
  resetTable
  If Player = 1 Then ballinPlay = ballinPlay + 1

  If ballinPlay > balls then
    endGame = 1
    checkContinue
  Else
    If state = True Then checkContinue
  End If
End Sub

'************Game State Check
Sub gameState
  If state = False Then
    GameOverReel.setValue(1)
    If B2SOn Then
      controller.B2SSetGameOver 35,1
      DOF 121, DOFOff
      DOF 122, DOFOff
    End If
    If vrOption > 0 Then vrGameOver.visible = 1
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
    MatchReel.setValue(0)
    TiltReel.setValue(0)
    tiltReel.SetValue(0)
    If vrOption > 0 Then vrTilt.visible = 0
    If B2SOn Then
      controller.B2SSetTilt 33,0
      controller.B2SSetMatch 34,0
      controller.B2SSetGameOver 35,0
    End If
  End If
End Sub

'*************Ball in Launch Lane on Plunger Tip
Dim ballREnabled, relGateHit
Sub ballHome_hit
  ballREnabled = 1
  shootAgainFlag = 0
  If B2SOn Then DOF 165, DOFOn
  relGateHit = 0
  Set controlBall = ActiveBall
    contBallInPlay = True
End Sub

'*************Ball off of Plunger Tip
Sub ballHome_unhit
  If B2SOn Then DOF 165, DOFOff
End Sub

'******* for ball control script
Sub endControl_Hit()
    contBallInPlay = False
End Sub

'************Check if Ball Out of Launch Lane
Dim ballInLane
Sub launchLaneExit_hit
  If ballREnabled=1 Then
    ballREnabled = 0
    ballInLane = False
  End If
  firstBallOut = 1
End Sub

Dim shootAgainFlag
Sub shootAgainOn
  shootAgainFlag = 1
  LshootAgain.state = 1
  If B2SOn Then controller.B2SSetShootAgain 36,1
  freeBallCount = 0
End Sub

'************** Reset Table
Sub resetTable
  L200.state = 0
  L100a.state = 0
  L100b.state = 0
  FreeBallGateClose
  freeBallCount = 0
End Sub

'************** Bumpers and Skirt Animation
dim bump1step, bump2step, bump3step
Sub bumpers_Hit(Index)
  If ballControlBlock = True Then Exit Sub
  EVAL ("skirt" & (index + 1)).enabled = 1
  If tilt = False Then
    Select Case (Index)
      Case 0: PlayFieldSound "PopBump", 0, bumper1, 1
          If B2SOn Then DOF 105,DOFPulse
          addScore 10
      Case 1: playFieldSound "PopBump", 0, bumper2, 1
          If B2SOn Then DOF 106,DOFPulse
          addScore 1
      Case 2: playFieldSound "PopBump", 0, bumper3, 1
          If B2SOn Then DOF 107,DOFPulse
          addScore 1
    End Select
  End If
End Sub

Sub skirt1_timer
    Dim RotY, BP
    Select Case bump1Step
        Case 3: RotY=-3
        Case 4: RotY=2
        Case 5: RotY=-1
        Case 6: RotY=1
        Case 7: RotY=0: bump1Step = 0: Skirt1.enabled = 0
    End Select
    bump1Step = bump1Step + 1
    For Each BP in BP_Skirt1 : BP.RotY= RotY: Next
End Sub

Sub skirt2_timer
    Dim RotY, BP
    Select Case bump2Step
        Case 3: RotY=-3
        Case 4: RotY=2
        Case 5: RotY=-1
        Case 6: RotY=1
        Case 7: RotY=0: bump2Step = 0: Skirt2.enabled = 0
    End Select
    bump2Step = bump2Step + 1
    For Each BP in BP_Skirt2 : BP.RotY= RotY: Next
End Sub

Sub skirt3_timer
    Dim RotY, BP
    Select Case bump3Step
        Case 3: RotY=-3
        Case 4: RotY=2
        Case 5: RotY=-1
        Case 6: RotY=1
        Case 7: RotY=0: bump3Step = 0: Skirt3.enabled = 0
    End Select
    bump3Step = bump3Step + 1
    For Each BP in BP_Skirt3 : BP.RotY= RotY: Next
End Sub


'************** Wire RollOvers
Dim wire, wireNumber
Sub wireAnimation_Timer
  wire = wire + 1
  Select Case wire
    Case 1: EVAL ("BM_wire_00" & wireNumber).transz = -10
    Case 2: EVAL ("BM_wire_00" & wireNumber).transz = -4
    Case 3: EVAL ("BM_wire_00" & wireNumber).transz = -1
    Case 4: EVAL ("BM_wire_00" & wireNumber).transz = 0
        wire = 0
        wireAnimation.enabled = 0
  End Select
End Sub

Dim button, buttonHit
Sub buttonAnimation_Timer
  button = button +1
  Select Case button
    Case 1: EVAL("BM_sw_00" & buttonHit).transz=1
    Case 2: EVAL("BM_sw_00" & buttonHit).transz=0
    Case 3: EVAL("BM_sw_00" & buttonHit).transz=1
    Case 4: EVAL("BM_sw_00" & buttonHit).transz=3
          button = 0
          buttonAnimation.enabled = 0
  End Select
End Sub

sub swLoutlane_hit
  dim bp
  for each bp in bp_sw_outleft_001 : bp.transz = -8 : Next
end Sub

sub swLoutlane_unhit
  dim bp
  for each bp in bp_sw_outleft_001 : bp.transz = 0 : Next
end sub

sub swRoutlane_hit
  dim bp
  for each bp in bp_sw_outright_001 : bp.transz = -8 : Next
end Sub

sub swRoutlane_unhit
  dim bp
  for each bp in bp_sw_outright_001 : bp.transz = 0 : Next
end sub

'************** Triggers
Dim s100, s200
Sub triggers_Hit (Index)
  tb.text = Index
  Select Case Index
    Case 0: pillScore(Lrar001.state): Lrar001.state = 1:  rockmakers : playsound "relay"
    Case 1: pillScore(Lrar002.state): Lrar002.state = 1: rockmakers : playsound "relay"
    Case 2: pillScore(Lrar003.state): Lrar003.state = 1: rockmakers : playsound "relay"
    Case 3: pillScore(Lrar004.state): Lrar004.state = 1: rockmakers : playsound "relay"
    Case 4: pillScore(Lrar005.state): Lrar005.state = 1: rockmakers : playsound "relay"
    Case 5: pillScore(Lrar006.state): Lrar006.state = 1: rockmakers : playsound "relay"
    Case 6: pillScore(Lrar007.state): Lrar007.state = 1: rockmakers : playsound "relay"
    Case 7: pillScore(Lrar008.state): Lrar008.state = 1: rockmakers : playsound "relay"
    Case 8: pillScore(Lrar009.state): Lrar009.state = 1: rockmakers : playsound "relay"
    Case 9: pillScore(Lrar0010.state): Lrar0010.state = 1: rockmakers : playsound "relay"
    Case 10: addScore 1 : playsound "relay"                         ' one point rubbers
    Case 11: L100a.state = 1: L100b.state = 1: s100 = 1 : playsound "relay" ' top switch hit from shooter lane
            buttonHit = 1
            buttonAnimation.Enabled =1
    Case 12: L200.state = 1: s200 = 1 : playsound "relay"       ' 200 point light rubber
    Case 13: topLanes(1)                        ' left 100
            buttonHit = 4
            buttonAnimation.Enabled =1
    Case 14: topLanes(2)                        ' top 100 of 200 pt lane
            buttonHit = 5
            buttonAnimation.Enabled =1
    Case 15: topLanes(2)                        ' lower 100 of 200 pt lane
            buttonHit = 6
            buttonAnimation.Enabled =1
    Case 16: topLanes(1)                        ' right 100
            buttonHit = 7
            buttonAnimation.Enabled =1
    Case 17: FreeBallGateCounter                  ' top fbg switch
            buttonHit = 2
            buttonAnimation.Enabled =1
    Case 18: FreeBallGateCounter                  ' middle pf fbg switch
            buttonHit = 8
            buttonAnimation.Enabled =1
    Case 19: FreeBallGateCounter                  ' lower fbg switch
            buttonHit = 9
            buttonAnimation.Enabled =1
    Case 20: addScore 50: wireNumber = 1: wireAnimation.enabled = 1 'left outlane
    Case 21: addScore 50:                                 'right outlne
            wireNumber = 2: wireAnimation.Enabled = 1
            If freeBallCount > 1 and Zipped = 0 Then ShootAgainOn
    Case 22: L100a.state = 1: L100b.state = 1: s100 = 1 'upper ball gate to inlanes
            buttonHit = 3
            buttonAnimation.Enabled =1
  End Select
End Sub

Sub pillScore(lightState)
  If lightState = 1 Then
    addScore 1
  Else
    addScore 10
  End If
End Sub

Sub topLanes(lane)
  If lane = 1 Then
    If s100 = 1 Then
      addScore 100
    Else
      addScore 10
    End If
  End If
  If lane = 2 Then
    If s200 = 1 Then
      addscore 100
    Else
      addScore 10
    End If
  End If
End Sub

Dim ROCK, MAKERS
Sub rockMakers
  If Lrar001.state = 1 and Lrar002.state = 1 and Lrar003.state = 1 and Lrar004.state = 1 and ROCK = 0 Then
    ROCK = 1
    rockArock
    playsound "relay"
  End If
  If Lrar005.state = 1 and Lrar006.state = 1 and Lrar007.state = 1 and Lrar008.state = 1 and Lrar009.state = 1 and Lrar0010.state = 1 and MAKERS = 0 Then
    MAKERS = 1
    rockArock
    playsound "relay"
  End If
End Sub

Dim rarCount
Sub rockArock
  RaRBonus(player) = RaRBonus(player) + 1
  If RaRBonus(player) > 9 Then RaRBonus(player) = 9
  If vrOption > 0 Then vrRaR(player)
  If B2SOn Then Controller.B2SSetReel (Player+16), RaRBonus(Player)
  If showDT = False Then PlayReelSound "Reel5", bgRaRReel Else PlayReelSound "Reel5", dtRaRReel
  RaRReel(player).addvalue(1)
  rarCount = rarCount + 1
  checkRaRReplay
  If rarCount > 1 Then
    rarCount = 0
    rockArockReset.enabled  = 1
  End If
End Sub

Dim rarLightCount
Sub rockArockReset_Timer
  rarLightCount = 0
  ROCK = 0
  MAKERS = 0
  For i = 1 to 10
    rarLightCount = rarLightCount + 1
    If EVAL("Lrar00" & rarLightCount).state = 1 Then
      EVAL("Lrar00" & rarLightCount).state = 0
      tb.text = rarLightCount
      i = 10
    End If
  Next
  If rarLightCount > 9 Then rockArockReset.enabled = 0
End Sub


'************* Free Ball Gate
Dim freeBallCount
Sub freeBallGateCounter
  If Zipped = 1 Then
    If freeBallCount = 0 Then freeBallCount = 1
    Lopen.state = 1
  Else
        Lopen.state = 1
    freeBallCount = freeBallCount + 1
    If freeBallCount > 1 Then
      FreeBallGateOpen
      Lgate1b.state = 1
      Lgate1a.state = 1
    End If
    If freeBallCount>2 Then freeBallCount = 2
  End If
End Sub

Sub freeBallGateOpen
  If Tilt = False Then
    bottomgate.rotateToEnd
  End If
End Sub

Sub freeBallGateClose
  Lgate1a.state = 0
  Lgate1b.state = 0
  Lopen.state = 0
  bottomgate.rotateToStart
End Sub

'************** RollOvers
'Dim buttonNumber
'Sub rollOver_Hit (Index)
' buttonNumber = (Index)
' rollOverAnimation.enabled = 1
' If B2SOn Then DOF (160 + Index), DOFPulse
' Select Case (Index)
'   Case 0: addScore (1)
'   Case 1: addScore (1)
' End Select
'End Sub

'Dim button
'Sub rollOverAnimation_Timer
' button = button + 1
' Select Case Button
'   Case 1:
'   Case 2:
'   Case 3:
'   Case 4:
' End Select
'End Sub

'************** Mushrooms
Dim mushroomNumber, mush1step, mush2step, mush3step
Sub mushrooms_Hit (Index)
  If ballControlBlock = True Then Exit Sub
  EVAL ("mushroom" & (index + 1) & "Timer").enabled = 1
  If Tilt = False Then
    Select Case (Index)
      Case 0: zipFlippers: addScore 10: L100a.state = 0: L100b.state = 0: s100 = 0: L200.state = 0: s200 = 0
      Case 1: UnzipFlippers: addScore 100
      Case 2: UnzipFlippers: addScore 100
    End Select
  End If
End Sub

Sub mushroom1Timer_timer
    Dim Transz, BP
    Select Case mush1Step
        Case 1: transz=6
        Case 3: transz=3
        Case 5: transz=2
        Case 6: transz=1
        Case 7: transz=0: mush1Step = 0: mushroom1Timer.enabled = 0
    End Select
    mush1Step = mush1Step + 1
    For Each BP in BP_mush1 : BP.transz= transz: Next
End Sub

Sub mushroom2Timer_timer
    Dim Transz, BP
    Select Case mush2Step
        Case 1: transz=6
        Case 3: transz=3
        Case 5: transz=2
        Case 6: transz=1
        Case 7: transz=0: mush2Step = 0: mushroom2Timer.enabled = 0
    End Select
    mush2Step = mush2Step + 1
    For Each BP in BP_mush2 : BP.transz= transz: Next
End Sub

Sub mushroom3Timer_timer
    Dim Transz, BP
    Select Case mush3Step
        Case 1: transz=6
        Case 3: transz=3
        Case 5: transz=2
        Case 6: transz=1
        Case 7: transz=0: mush3Step = 0: mushroom3Timer.enabled = 0
    End Select
    mush3Step = mush3Step + 1
    For Each BP in BP_mush3 : BP.transz= transz: Next
End Sub


Sub knock
  If B2SOn = True Then DOF 128, DOFPulse Else playSound "knocker"
End Sub

Dim replayRaR, repRaR(4), RaRReplay(4)
Sub checkRaRReplay
  For replayRaR = (repRaR(player) + 1) to 2
    If RaRBonus(player) => RaRReplay(replayRaR) Then
      If replayValue = 0 Then
        shootAgainOn
      Else
        increaseCredits
      End If
      repRaR(player) = repRaR(Player) + 1
      knock
    End If
  Next
End Sub

' *******************Slings
Dim LStep, L1Step, RStep

Sub leftSlingShot_Slingshot
  If ballControlBlock = True Then Exit Sub
  playfieldSound "SlingShot", 0, SoundPoint12, 1
  If B2SOn Then DOF 103, DOFPulse
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
  addScore(10)
End Sub

Sub LeftSlingShotLP_Slingshot
  If ballControlBlock = True Then Exit Sub
  playfieldSound "SlingShot", 0, SoundPoint12, 1
  If B2SOn Then DOF 103, DOFPulse
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
  addScore(10)
End Sub

Sub leftSlingShot_Timer
  Dim BP
  Dim v1, v2, v3, v4, x
  v1 = True: v2 = False: v3 = False: v4 = True: x = -30
    Select Case LStep
        Case 2:v1 = True: v2 = False: v3 = True:  v4 = False: x = -20
        Case 3:v1 = True: v2 = True:  v3 = False: v4 = False: x = -10
        Case 4:v1 = True:  v2 = False: v3 = False: v4 = False: x = 0: LeftSlingShot.TimerEnabled = 0
    End Select

  For Each BP in BP_Lsling : BP.Visible = v1: Next
  For Each BP in BP_Lsling2 : BP.Visible = v2: Next
  For Each BP in BP_Lsling1 : BP.Visible = v3: Next
  For Each BP in BP_Lsling2 : BP.Visible = v4: Next

    LStep = LStep + 1
End Sub

Sub leftSlingShotUpper_Slingshot
  If ballControlBlock = True Then Exit Sub
  playfieldSound "SlingShot", 0, SoundPoint12, 1
  If B2SOn Then DOF 103, DOFPulse
    L1Step = 0
    LeftSlingShotUpper.TimerEnabled = 1
  addScore(1)
End Sub

Sub LeftSlingShotUpperLP_Slingshot
  If ballControlBlock = True Then Exit Sub
  playfieldSound "SlingShot", 0, SoundPoint12, 1
  If B2SOn Then DOF 103, DOFPulse
    L1Step = 0
    LeftSlingShotUpper.TimerEnabled = 1
  addScore(1)
End Sub

Sub leftSlingShotUpper_Timer
  Dim BP
  Dim v1, v2, v3, x
  v1 = True: v2 = False: v3 = True: x = -30
    Select Case L1Step
        Case 2:v1 = True: v2 = True:  v3 = False: x = -20
        Case 3:v1 = True:  v2 = False: v3 = False: x = 0: LeftSlingShotUpper.TimerEnabled = 0
    End Select

  For Each BP in BP_Circle_031 : BP.Visible = v1: Next
  For Each BP in BP_Circle_024 : BP.Visible = v2: Next
  For Each BP in BP_Circle_019 : BP.Visible = v3: Next

    L1Step = L1Step + 1
End Sub

Sub rightSlingShot_Slingshot
  If ballControlBlock = True Then Exit Sub
  playfieldSound "SlingShot", 0, SoundPoint13, 1
  If B2SOn Then DOF 104, DOFPulse
    RStep = 0
    RightSlingShot.TimerEnabled = 1
    addScore(10)
End Sub

Sub RightSlingShotLP_Slingshot
  If ballControlBlock = True Then Exit Sub
  playfieldSound "SlingShot", 0, SoundPoint13, 1
  If B2SOn Then DOF 104, DOFPulse
    RStep = 0
    RightSlingShot.TimerEnabled = 1
    addScore(10)
End Sub

Sub rightSlingShot_Timer
  Dim BP
  Dim v1, v2, v3, v4, x
  v1 = True: v2 = False: v3 = False: v4 = True: x = -30
    Select Case RStep
        Case 2:v1 = True: v2 = False: v3 = True:  v4 = False: x = -20
        Case 3:v1 = True: v2 = True:  v3 = False: v4 = False: x = -10
        Case 4:v1 = True:  v2 = False: v3 = False: v4 = False: x = 0: RightSlingShot.TimerEnabled = 0
    End Select

  For Each BP in BP_Rsling : BP.Visible = v1: Next
  For Each BP in BP_Rsling2 : BP.Visible = v2: Next
  For Each BP in BP_Rsling1 : BP.Visible = v3: Next
  For Each BP in BP_Rsling2 : BP.Visible = v4: Next

    RStep = RStep + 1
End Sub

'************ Zip Flippers
Dim Zipped
Sub zipFlippers
  dim bp
  Lopen.state = 0
  Lgate1a.state = 0
  Lgate1b.state = 0
  FreeBallGateClose
  If Zipped = 0 Then playFieldSound "metalHitLow", 0, SoundPointZip, 1
  Zipped = 1
  For each bp in BP_flipL_001 : bp.visible=0 : Next
  For each bp in BP_flipL_002 : bp.visible=1 : Next
  For each bp in BP_flipR_001 : bp.visible=0 : Next
  For each bp in BP_flipR_002 : bp.visible=1 : Next

  LeftFlipper.Enabled = 0
  RightFlipper.Enabled = 0
  LeftFlipper1.Enabled = 1
  RightFlipper1.Enabled = 1
  flipperRShadow.visible = 0
  flipperLShadow.visible = 0
  flipperRShadow1.visible = 1
  flipperLShadow1.visible = 1
End Sub

Sub unzipFlippers
  dim bp
  If Zipped = 1 Then playFieldSound "metalHitLow", 0, SoundPointZip, 1
  Zipped = 0
  If freeBallCount = 1 Then Lopen.state = 1
  If freeBallCount > 1 Then
    Lopen.state = 1
    Lgate1a.state = 1
    Lgate1b.state = 1
    freeBallGateOpen
  End If
  LeftFlipper.Enabled = 1
  RightFlipper.Enabled = 1
  For each bp in BP_flipL_001 : bp.visible=1 : Next
  For each bp in BP_flipL_002 : bp.visible=0 : Next
  For each bp in BP_flipR_001 : bp.visible=1 : Next
  For each bp in BP_flipR_002 : bp.visible=0 : Next
  LeftFlipper1.Enabled = 0
  RightFlipper1.Enabled = 0
  flipperRShadow.visible = 1
  flipperLShadow.visible = 1
  flipperRShadow1.visible = 0
  flipperLShadow1.visible = 0
End Sub


'***************Score Motor Run one full rotation
Dim scoreMotorCount, addCredit
Sub scoremotor5_Timer
  scoreMotorCount = scoreMotorCount + 1
  playFieldSound "ScoreMotorSingleFire", 0, SoundPointScoreMotor, 0. 'need to set location of score motor under the PF
  If scoremotorCount = 5 Then
    FreeBallGateClose
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
  If showDT = False Then PlayReelSound "Reel5", bgcrReel Else PlayReelSound "Reel5", dtcrReel
  If B2SOn Then DOF 199, DOFOn
  If credit > 15 then credit = 15
  CreditReel.setValue (credit)
  If vrOption > 0 Then vrCredit
  If B2SOn Then
    controller.B2SSetCredits credit
    If credit > 0 Then DOF 199, DOFOn
  End If
End Sub

'****************Score Motor Run Timer
Dim bellRing, scoreMotorLoop, kickerUp, KickBonus, motorOn
Sub scoreMotor_timer
  scoreMotorLoop = scoreMotorLoop + 1
  If scoreMotorLoop < 6 Then playFieldSound "ScoreMotorSingleFire", 0, SoundPointScoreMotor, 0.2
  motorOn = 1

  If scoreMotorLoop > 6 Then
    scoreMotorLoop = 0
    motorOn = 0
    scoreMotor.Enabled = 0
    Exit Sub
  End If

'  These Flags are passed by scores with multiple of 10, 100 or 1000
  If flag = 1 or flag10 = 1 or flag100 = 1 or flag1000 = 1 Then
    Select Case scoreMotorLoop
      Case 1: totalUp Point:' If point = 1000 Then advanceBonus 'this is for tables with a 1000 point bonus tree
      Case 2: If bellRing > 1 Then totalUp point
      Case 3: If bellRing > 2 Then totalUp Point: 'If point = 1000 Then advanceBonus
      Case 4: If bellRing > 3 Then totalUp point
      Case 5: If bellRing > 4 Then totalUp point: 'If point = 1000 Then advanceBonus
      Case 6: scoreMotorLoop = 0
          FreeBallGateClose
          flag1 = 0
          flag10 = 0
          flag100 = 0
          flag1000 = 0
          motorOn = 0
          scoreMotor.enabled = 0 'originally had If drainActive = 0 Then scoreMotor.Enabeled = 0 but this casued a final ball drain error in PG
    End Select
  End If
End Sub

'***************Scoring Routine
Dim flag1, flag10, flag100, flag1000, point, point10
Sub pts1
  If pts1Timer.enabled = False Then
    pts1Timer.enabled = True
  End If
End Sub

Sub pts10
  If pts10Timer.enabled = False Then
    pts10Timer.enabled = True
    stopsound "Bell10"
    playFieldSound "Bell10", 0, soundPoint13, 1
  End If
End Sub

Sub pts100
  If pts100Timer.enabled = False Then
    pts100Timer.enabled = True
    stopsound "bell10"
    playFieldSound "Bell10", 0, soundPoint13, 1
  End If
End Sub

Dim pts1Block, pts10Block, pts100Block, intvl
intvl = 40
pts1Timer.interval = intvl
pts10Timer.interval = intvl
pts100Timer.interval = intvl

Sub pts1Timer_Timer
  pts1Block = pts1Block + 1
  If pts1Block = 1 Then
    pts1Block = 0
    pts1Timer.enabled = False
  End If
End Sub

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

    If points < 9 Then
'     Number Matching, decrement the match unit for each 1 point score roll over from 10 to 1
      matchNumber = matchNumber - 1
      If matchNumber < 1 Then matchNumber = 10
      If points > 1 Then
        bellRing = (points)
        point = 1: flag1 = 1: scoreMotor.enabled = 1
      Else
        totalUp 1
      End If
      Exit Sub
    End If

    If points > 9 and points <100 Then
      If shootAgainFlag = 0 Then LshootAgain.state = 0
      If B2SOn Then controller.B2SSetShootAgain 36,0
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
        totalUp 10
      End If
      Exit Sub
    End If

    If points > 99 and points < 1000 Then
      If shootAgainFlag = 0 Then LshootAgain.state = 0
      If B2SOn Then controller.B2SSetShootAgain 36,0
      If points > 100 Then
        bellRing = (points / 100)
        point = 100: flag100 = 1: scoreMotor.enabled = 1
      Else
        totalUp 100
      End If
      Exit Sub
    End If

    If Points > 999 Then
      If points > 1000 Then
        bellRing = (points / 1000)
        point = 1000: flag1000 = 1: scoreMotor.enabled = 1
      Else
        totalUp 1000
      End If
    End If
  End If
End Sub

Dim replayX, replay(7), repAwarded(5), block
Sub totalUp(points)

  If pts1Timer.enabled = False and points = 1 Or pts10Timer.enabled = False and points = 10 Or pts100Timer.enabled = False and points = 100 Then block = 0 Else block = 1

  If B2SOn And showDT = False Then
    If Player = 1 and block = 0 Then PlayReelSound "Reel1", bgs1Reel
    If Player = 2 and block = 0 Then PlayReelSound "Reel2", bgs2Reel
    If Player = 3 and block = 0 Then PlayReelSound "Reel3", bgs3Reel
    If Player = 4 and block = 0 Then PlayReelSound "Reel4", bgs4Reel
  Else
    If Player = 1 and block = 0 Then PlayReelSound "Reel1", dts1Reel
    If Player = 2 and block = 0 Then PlayReelSound "Reel2", dts2Reel
    If Player = 3 and block = 0 Then PlayReelSound "Reel3", dts3Reel
    If Player = 4 and block = 0 Then PlayReelSound "Reel4", dts4Reel
  End If

  If flag1 = 1 or points = 1 Then
    If chime = 0 Then pts1
  End If

  If flag10 = 1 or points = 10 Then
    If chime = 0 Then
      pts10
    Else
      If B2SOn Then DOF 142,DOFPulse
    End If
  End If

  If flag100 = 1 or points = 100 Then
    If chime = 0 Then
      pts100
    Else
      If B2SOn Then DOF 142,DOFPulse
    End If
  End If

  If block = 0 Then score(player) = score(player) + points
  If block = 0 Then sReel(player).addvalue(points)
  If vrOption > 0 Then vrScore(player)

  If B2SOn Then controller.B2SSetScorePlayer player, score(player)

  For replayX = (rep(player) + 1) to 4
    If score(player) => replay(replayX) Then
      If replayValue = 0 Then
        shootAgainOn
      Else
        increaseCredits
      End If
      rep(player) = rep(Player) + 1
      knock
    End If
  Next
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
        If B2SOn Then
      controller.B2SSetTilt 33,1
      controller.B2SSetdata 1, 0
    End If
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

'***************Match
Sub match
' matchNumber = ((int(rnd(1)*10)) * 10) + 10
  If matchNumber = 0 Then matchNumber = 10
  MatchReel.setValue(matchNumber) 'need to change if 1 point table

  If B2SOn Then controller.B2SSetMatch 34, matchNumber
  If vrOption > 0 Then vrMatchSet

  For i = 1 to players
    If (matchNumber * 1) = (score(i) mod 10) Then 'need to set this to match 10's or 100's
      increaseCredits
      knock
      End If
    Next
End Sub

'***************************************************Reset Reels Section*******************************************************

'This Sub looks at each individual digit in each players score and sets them in an array RScore.  If the value is >0 and <9
'then the players score is increased by one times the position value of that digit (ie 1 * 1000 for the 1000's digit)
'If the value of the digit is 9 then it subtracts 9 times the postion value of that digit (ie 9*100 for the 100's digit)
'so that the score is not rolled over and the next digit in line gets incremented as well (ie 9 in the 10's positon gets
'incremented so the 100's position rolls up by one as well since 90 -> 100).  Lastly the RScore array values get incremented
'by one to get ready for the next pass.

Dim rScore(4,5), resetLoop, test, playerTest, resetFlag, reelFlag, reelStop, reelDone(4), RaRreelDone(4), RaRBonus(4)
Sub countUp
  For playerTest = 1 to maxPlayers
    For test = 0 to 4
      rScore(playerTest,test) = Int(score(playerTest)/10^test) mod 10
    Next
    If RaRBonus(playerTest) <> 0 Then RaRBonus(playerTest) = RaRBonus(playerTest) + 1
    If RaRBonus(playerTest) = 10 Then RaRBonus(playerTest) = 0
  Next

  For playerTest = 1 to maxPlayers
    For x = 0 to 4
      If rscore(playerTest, x) > 0 And rscore(playerTest, x) < 9 Then score(playerTest) = score(playerTest) + 10^x
      If rScore(playerTest, x) = 9 Then score(playerTest) = score(playerTest) - (9 * 10^x)
      If rScore(playerTest, x) > 0 Then rScore(playerTest, x) = rScore(playerTest, x) + 1
      If rScore(playerTest, x) = 10 Then rScore(playerTest, x) = 0
    Next
  Next

  If score(1) = 0 and score(2) = 0 and score(3) = 0 and score(4) = 0 and RaRBonus(1) = 0 and RaRBonus(2) = 0 and RaRBonus(3) = 0 and RaRBonus(4) = 0 Then
    reelFlag = 1
    For i = 1 to maxPlayers
      score(i) = 0
      rep(i) = 0
      repRaR(i) = 0
      repAwarded(i) = 0
    Next
  End If
End Sub

'This Sub sets each B2S reel or Desktop reels to their new values and then plays the score motor sound each time and the
'reel sounds only if the reels are being stepped

Sub updateReels
' tb.text = reelDone(2)
  For playerTest = 1 to maxPlayers
    For x = 1 to maxPlayers
      If score(x) = 0 Then reelDone(x) = 1
    Next
    If B2SOn Then
      controller.B2SSetScorePlayer playerTest, score(playerTest)
      controller.B2SSetReel(playerTest+16), RaRBonus(playerTest)
    Else
      sReel(playerTest).setvalue (score(playerTest))
      RaRReel(playerTest).setvalue (RaRBonus(playerTest))
    End If

    If reelDone(playerTest) = 0 and reelStop = 0 Then
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
      vrScore(playerTest)
      vrRaR(playerTest)
    End If


    If RaRreelDone(playerTest) = 0 Then
      If showDT = False Then
        If RaRBonus(playerTest) = 0 Then RaRreelDone(playerTest) = 1
        If RaRreelDone(playerTest) = 0 Then PlayReelSound "Reel5", bgRaRReel
      Else
        If RaRBonus(playerTest) = 0 Then RaRreelDone(playerTest) = 1
        If RaRreelDone(playerTest) = 0 Then PlayReelSound "Reel5", dtRaRReel
      End If
    End If

  Next
  playfieldSound "ScoreMotorSingleFire", 0, SoundPointScoreMotor, 0.2
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

Dim highScore(5)
Dim activeScore(5), hs, chX, chY, chZ, chIX, tempI(4), tempI2(4), flag, hsI, hsX
'goes through the 5 high scores one at a time and compares them to the player's scores high to
'if a player's score is higher it marks that postion with ActiveScore(x) and moves all of the other
' high scores down by one along with the high score's player initials
' also clears the new high score's initials for entry later
Sub checkHighScores
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

    For x = 1 to 35
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
    powerSetting = CInt(Right(temp(33),1))
    replayValue = CInt(Right(temp(34),1))
    difficultySetting = CInt(Right(temp(35),1))
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
    scoreFile.WriteLine "Power Setting: " & powerSetting
    scoreFile.WriteLine "Replay Value: " & replayValue
    scoreFile.WriteLine "Difficulty: " & difficultySetting
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

Sub turnOff
  For i = 1 to 3
    EVAL("Bumper" & i).hasHitEvent = 0
  Next
  If tiltPenalty = 1 then ballInPlay = Balls
    leftFlipper.RotateToStart
  stopSound "FlipBuzzLA"
  stopSound "FlipBuzzLB"
  stopSound "FlipBuzzLC"
  stopSound "FlipBuzzLD"
  stopSound "FlipBuzzRA"
  stopSound "FlipBuzzRB"
  stopSound "FlipBuzzRC"
  stopSound "FlipBuzzRD"
  If B2SOn Then DOF 101, DOFOff
  If B2SOn Then DOF 111, DOFOff
  RightFlipper.RotateToStart
  If B2SOn Then DOF 102, DOFOff
  If B2SOn Then DOF 112, DOFOff
  LshootAgain.state = 0
  FreeBallGateClose
  UnzipFlippers
  LeftSlingShot.collidable = 0
  RightSlingShot.collidable = 0
End Sub

'************VR Subs
Sub vrCredit
  vrCreditReel.objRotX = credit * 22.5
End Sub
Dim vrOne, stringOne, vrTen, vrHundred, vr1k, vr10k, resetReelFlag, nowUp

Sub vrScore(nowUp)
  vrTen = (INT (score(nowUp)/10) )Mod 10
  vrHundred = (INT(score(nowUp)/100)) Mod 10
  vr1k =(INT (score(nowUp)/1000) )Mod 10
  stringOne = (Cstr(score(nowUp)))
  stringOne = (Right(trim(stringOne), 1))
  vrOne = (Cint(stringOne))


  Select Case nowUp
    Case 1: vrScoreReel11.objRotX = vrOne * 36
        vrScoreReel110.objRotX = vrTen * 36
        vrScoreReel1100.objRotX = vrHundred * 36
        vrScoreReel11k.objRotX = vr1k * 36
    Case 2: vrScoreReel21.objRotX = vrOne * 36
        vrScoreReel210.objRotX = vrTen * 36
        vrScoreReel2100.objRotX = vrHundred * 36
        vrScoreReel21k.objRotX = vr1k * 36
    Case 3: vrScoreReel31.objRotX = vrOne * 36
        vrScoreReel310.objRotX = vrTen * 36
        vrScoreReel3100.objRotX = vrHundred * 36
        vrScoreReel31k.objRotX = vr1k * 36
    Case 4: vrScoreReel41.objRotX = vrOne * 36
        vrScoreReel410.objRotX = vrTen * 36
        vrScoreReel4100.objRotX = vrHundred * 36
        vrScoreReel41k.objRotX = vr1k * 36
  End Select
End Sub

Dim RaRUp
Sub vrRaR(RaRup)
  Select Case RaRUp
    Case 1: vrRaR1.objRotX = RaRBonus(1) * 36
    Case 2: vrRaR2.objRotX = RaRBonus(2) * 36
    Case 3: vrRaR3.objRotX = RaRBonus(3) * 36
    Case 4: vrRaR4.objRotX = RaRBonus(4) * 36
  End Select
End Sub

Sub vrBipSet
  Select Case ballInPlay
    Case 0: vrBip.image = "bip0"
    Case 1: vrBip.image = "bip1"
    Case 2: vrBip.image = "bip2"
    Case 3: vrBip.image = "bip3"
    Case 4: vrbip.image = "bip4"
    Case 5: vrBip.image = "bip5"
  End Select
End Sub

Sub vrMatchSet
  Select Case matchNumber
    Case 0: vrMatch.image = "m0"
    Case 1: vrMatch.image = "m1"
    Case 2: vrMatch.image = "m2"
    Case 3: vrMatch.image = "m3"
    Case 4: vrMatch.image = "m4"
    Case 5: vrMatch.image = "m5"
    Case 6: vrMatch.image = "m6"
    Case 7: vrMatch.image = "m7"
    Case 8: vrMatch.image = "m8"
    Case 9: vrMatch.image = "m9"
    Case 10: vrMatch.image = "m10"
  End Select
End Sub

Sub vrCanPlay
  Select Case players
    Case 0: vrCp.image = "cp0"
    Case 1: vrCp.image = "cp1"
    Case 2: vrCp.image = "cp2"
    Case 3: vrCp.image = "cp3"
    Case 4: vrCp.image = "cp4"
  End Select
End Sub

Sub vrPlayerUp
  Select Case player
    Case 1: vrUp.image = "up1"
    Case 2: vrUp.image = "up2"
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
  For each object in backdropstuff: Object.visible = 0: Next
  lockdown.visible = 0
ElseIf vrOption = 2 Then
  for each VRThings in vrCab: VRThings.visible = 1: Next
  for each VRThings in vrRoom: VRThings.visible = 1: Next
  For each object in backdropstuff: Object.visible = 0: Next
  lockdown.visible = 0
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
  If PinCab_Shooter.Y < -339 then
       PinCab_Shooter.Y = PinCab_Shooter.Y + 5
  End If
End Sub

Sub TimerVRPlunger1_Timer
  PinCab_Shooter.Y = -484.9259 + (5 * Plunger.Position) -20
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
  Dim tmp
    If PFOption=1 Then tmp = TableObj.x * 2 / table1.width-1
  If PFOption=2 Then tmp = TableObj.y * 2 / table1.height-1
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

Const tnob = 2 ' total number of balls

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
        PlaySound("ArchHit" & b), 0, (BallVel(BOT(b))/15)^5 * xGain(BOT(b)), AudioPan(BOT(b)), 0, (BallVel(BOT(b))/10)^5, 0, 0, 0 'Left & Right stereo or Top & Bottom stereo PF Speakers.
      End If
      ArchRolling(b) = True
      PlaySound("ArchRollA" & b), -1, (BallVel(BOT(b))/40)^5 * xGain(BOT(b)), AudioPan(BOT(b)), 0, (BallVel(BOT(b))/40)^5, 1, 0, 0  'Left & Right stereo or Top & Bottom stereo PF Speakers.
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
        PlaySound("ArchHit" & b),   0, (BallVel(BOT(b))/15)^5 *    XVol(BOT(b))  *     YVol(BOT(b)),  -1, 0, (BallVel(BOT(b))/10)^5, 0, 0, -1 'Top Left PF Speaker
        PlaySound("ArchHit" & b),   0, (BallVel(BOT(b))/15)^5 * (1-XVol(BOT(b))) *     YVol(BOT(b)),   1, 0, (BallVel(BOT(b))/10)^5, 0, 0, -1 'Top Right PF Speaker
        PlaySound("ArchHit" & b),   0, (BallVel(BOT(b))/15)^5 *    XVol(BOT(b))  *  (1-YVol(BOT(b))), -1, 0, (BallVel(BOT(b))/10)^5, 0, 0,  1 'Bottom Left PF Speaker
        PlaySound("ArchHit" & b),   0, (BallVel(BOT(b))/15)^5 * (1-XVol(BOT(b))) *  (1-YVol(BOT(b))),  1, 0, (BallVel(BOT(b))/10)^5, 0, 0,  1 'Bottom Right PF Speaker
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
    If BOT(b).VelZ > -4 And BOT(b).VelZ < -2 And BOT(b).Z > 27 And BallinPlay => 1 Then
      PlaySound "BallDrop1" & b, 0, ABS(BOT(b).VelZ)/600 * xGain(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 0, 0, 0 'Left & Right stereo or Top & Bottom stereo PF Speakers.
    ElseIf BOT(b).VelZ > -8 And BOT(b).VelZ < -4 And BOT(b).Z > 27 And BallinPlay => 1 Then
      PlaySound "BallDrop2" & b, 0, ABS(BOT(b).VelZ)/600 * xGain(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 0, 0, 0 'Left & Right stereo or Top & Bottom stereo PF Speakers.
    ElseIf BOT(b).VelZ > -12 And BOT(b).VelZ < -8 And BOT(b).Z > 27 And BallinPlay => 1 Then
      PlaySound "BallDrop3" & b, 0, ABS(BOT(b).VelZ)/600 * xGain(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 0, 0, 0 'Left & Right stereo or Top & Bottom stereo PF Speakers.
    ElseIf BOT(b).VelZ < -12 And BOT(b).Z > 27 And BallinPlay => 1 Then
      PlaySound "BallDrop4" & b, 0, ABS(BOT(b).VelZ)/600 * xGain(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 0, 0, 0 'Left & Right stereo or Top & Bottom stereo PF Speakers.
    End If
  End If

  If PFOption = 3 Then
    If BOT(b).VelZ > -4 And BOT(b).VelZ < -2 And BOT(b).Z > 27 And BallinPlay => 1 Then
      PlaySound "BallDrop1" & b, 0, ABS(BOT(b).VelZ)/600 *    XVol(BOT(b))  *     YVol(BOT(b)), -1, 0, Pitch(BOT(b)), 1, 0, -1  'Top Left PF Speaker
      PlaySound "BallDrop1" & b, 0, ABS(BOT(b).VelZ)/600 * (1-XVol(BOT(b))) *     YVol(BOT(b)),  1, 0, Pitch(BOT(b)), 1, 0, -1  'Top Right PF Speaker
      PlaySound "BallDrop1" & b, 0, ABS(BOT(b).VelZ)/600 *    XVol(BOT(b))  * (1-YVol(BOT(b))), -1, 0, Pitch(BOT(b)), 1, 0,  1  'Bottom Left PF Speaker
      PlaySound "BallDrop1" & b, 0, ABS(BOT(b).VelZ)/600 * (1-XVol(BOT(b))) * (1-YVol(BOT(b))),  1, 0, Pitch(BOT(b)), 1, 0,  1  'Bottom Right PF Speaker
    ElseIf BOT(b).VelZ > -8 And BOT(b).VelZ < -4 And BOT(b).Z > 27 And BallinPlay => 1 Then
      PlaySound "BallDrop2" & b, 0, ABS(BOT(b).VelZ)/600 *    XVol(BOT(b))  *     YVol(BOT(b)), -1, 0, Pitch(BOT(b)), 1, 0, -1  'Top Left PF Speaker
      PlaySound "BallDrop2" & b, 0, ABS(BOT(b).VelZ)/600 * (1-XVol(BOT(b))) *     YVol(BOT(b)),  1, 0, Pitch(BOT(b)), 1, 0, -1  'Top Right PF Speaker
      PlaySound "BallDrop2" & b, 0, ABS(BOT(b).VelZ)/600 *    XVol(BOT(b))  * (1-YVol(BOT(b))), -1, 0, Pitch(BOT(b)), 1, 0,  1  'Bottom Left PF Speaker
      PlaySound "BallDrop2" & b, 0, ABS(BOT(b).VelZ)/600 * (1-XVol(BOT(b))) * (1-YVol(BOT(b))),  1, 0, Pitch(BOT(b)), 1, 0,  1  'Bottom Right PF Speaker
    ElseIf BOT(b).VelZ > -12 And BOT(b).VelZ < -8 And BOT(b).Z > 27 And BallinPlay => 1 Then
      PlaySound "BallDrop3" & b, 0, ABS(BOT(b).VelZ)/600 *    XVol(BOT(b))  *     YVol(BOT(b)), -1, 0, Pitch(BOT(b)), 1, 0, -1  'Top Left PF Speaker
      PlaySound "BallDrop3" & b, 0, ABS(BOT(b).VelZ)/600 * (1-XVol(BOT(b))) *     YVol(BOT(b)),  1, 0, Pitch(BOT(b)), 1, 0, -1  'Top Right PF Speaker
      PlaySound "BallDrop3" & b, 0, ABS(BOT(b).VelZ)/600 *    XVol(BOT(b))  * (1-YVol(BOT(b))), -1, 0, Pitch(BOT(b)), 1, 0,  1  'Bottom Left PF Speaker
      PlaySound "BallDrop3" & b, 0, ABS(BOT(b).VelZ)/600 * (1-XVol(BOT(b))) * (1-YVol(BOT(b))),  1, 0, Pitch(BOT(b)), 1, 0,  1  'Bottom Right PF Speaker
    ElseIf BOT(b).VelZ < -12 And BOT(b).Z > 27 And BallinPlay => 1 Then
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

Sub phys_plastic_hit  'dBm NEW plastics sub
  PlayFieldSoundAB "fx_plastichit", 0, 0.01
End Sub

Sub LeftFlipper_collide(parm) 'dBm corrected was "_collidable"
  RandomSoundFlipper()
End Sub

Sub LeftFlipper1_collide(parm)  'dBm NEW flipper sub added
  RandomSoundFlipper()
End Sub

Sub LeftFlipper2_collide(parm)  'dBm NEW flipper sub added
  RandomSoundFlipper()
End Sub

Sub RightFlipper_collide(parm)  'dBm corrected was "_collidable"
  RandomSoundFlipper()
End Sub

Sub RightFlipper1_collide(parm) 'dBm NEW flipper sub added
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
' You only need to add a Sub OnBallBallCollision(ball1, ball2, velocity) and when two balls collidable they
' will call this routine.

'New algorithm for OnBallBallCollision
'For stereo, xGain is a Playsound volume multiplier that provides a Constant Power pan.
'For quad, multiple PlaySound commands are launched together that are panned and faded to their maximum extents where PlaySound's PAN and FADE have the least error.
'XVol and YVol are Playsound volume multipliers that provide a Constant Power "pan" and "fade".
'Subtracting XVol or YVol from 1 yeilds an inverse response.

Sub OnBallBallCollision(ball1, ball2, velocity)
  If PFOption = 1 or PFOption = 2 Then
    PlaySound "BBcollidable", 0, (Csng(velocity) ^2 / 2000) * xGain(ball1), AudioPan(ball1), 0, Pitch(ball1), 0, 0, 0 'Left & Right stereo or Top & Bottom stereo PF Speakers.
  End If
  If PFOption = 3 Then
    PlaySound "BBcollidable", 0, (Csng(velocity) ^2 / 2000) *    XVol(ball1)  *    YVol(ball1),  -1, 0, Pitch(ball1), 0, 0, -1  'Top Left Playfield Speaker
    PlaySound "BBcollidable", 0, (Csng(velocity) ^2 / 2000) * (1-XVol(ball1)) *    YVol(ball1),   1, 0, Pitch(ball1), 0, 0, -1  'Top Right Playfield Speaker
    PlaySound "BBcollidable", 0, (Csng(velocity) ^2 / 2000) *    XVol(ball1)  * (1-YVol(ball1)), -1, 0, Pitch(ball1), 0, 0,  1  'Bottom Left Playfield Speaker
    PlaySound "BBcollidable", 0, (Csng(velocity) ^2 / 2000) * (1-XVol(ball1)) * (1-YVol(ball1)),  1, 0, Pitch(ball1), 0, 0,  1  'Bottom Right Playfield Speaker
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
    If Pan = "Lpan" Then PlaySound SoundName, 0, ReelVolAdj * 1.00, -0.12, 0, 0, 0, 1, 0  'Panned 3/4 Left at 0dB * ReelVolAdj
    If Pan = "Mpan" Then PlaySound SoundName, 0, ReelVolAdj * 0.33,  0.00, 0, 0, 0, 1, 0  'Panned Middle at -3dB * ReelVolAdj
    If Pan = "Rpan" Then PlaySound SoundName, 0, ReelVolAdj * 1.00,  0.12, 0, 0, 0, 1, 0  'Panned 3/4 Right at 0dB * ReelVolAdj
  Else
    If Pan = "Lpan" Then PlaySound SoundName, 0, ReelVolAdj * 0.33, -0.12, 0, 0, 0, 1, 0  'Panned 3/4 Left at -3dB * ReelVolAdj
    If Pan = "Mpan" Then PlaySound SoundName, 0, ReelVolAdj * 0.11,  0.00, 0, 0, 0, 1, 0  'Panned Middle at -6dB * ReelVolAdj
    If Pan = "Rpan" Then PlaySound SoundName, 0, ReelVolAdj * 0.33,  0.12, 0, 0, 0, 1, 0  'Panned 3/4 Right at -3dB * ReelVolAdj
  End If
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
dim RF1: Set RF1 = New FlipperPolarity
dim LF1: Set LF1 = New FlipperPolarity

InitPolarity

'
''*******************************************
'' Late 70's to early 80's
'
Sub InitPolarity()
        dim x, a : a = Array(LF, RF, LF1, RF1)
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
    LF1.Object = LeftFlipper1
        LF.EndPoint = EndPointLp  'you can use just a coordinate, or an object with a .x property. Using a couple of simple primitive objects
        RF.Object = RightFlipper
    RF1.Object = RightFlipper1
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
Sub TriggerLF1_Hit() : LF.Addball activeball : End Sub
Sub TriggerLF1_UnHit() :  LF.PolarityCorrect activeball : End Sub
Sub TriggerRF1_Hit() : RF.Addball activeball : End Sub
Sub TriggerRF1_UnHit() : RF.PolarityCorrect activeball : End Sub

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
  dim a : a = Array(LF, RF, LF1, RF1)
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
dim RFEndAngle, LFEndAngle, LF1EndAngle, RF1EndAngle

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
LF1EndAngle = LeftFlipper1.endAngle
RF1EndAngle = RightFlipper1.endAngle

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
RubbersD.addpoint 0, 0, 1.35        'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 1, 3.77, 1.1
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

' This subroutine updates the flipper shadows and visual primitives
Sub FlipperVisualUpdate
  dim bp
  For each BP in BP_gate2 : bp.Rotx = -Gate001.CurrentAngle*0.5 : Next
  For each BP in BP_pgate_001 : bp.Rotx = -Gate002.CurrentAngle*0.5 : Next
  For each BP in BP_gate1 : bp.Rotx = -Gate003.CurrentAngle*0.5 : Next
  For each BP in BP_Rubber___Post___Oval_001 : BP.rotz = bottomgate.currentangle : Next
    Dim BOT, b, bl
    BOT = GetBalls

  If ballsInPlay = 3 or scoreMotorLoop > 0 or scoreMotorCount > 0 Then
  Else
  End If

  For each bp in BP_flipL_001 : bp.RotZ = leftflipper.currentangle - 90 : Next
  For each bp in BP_flipL_002 : bp.RotZ = leftflipper1.currentangle - 90 : Next
  For each bp in BP_flipL_003 : bp.RotZ = leftflipper2.currentangle - 90 : Next
  For each bp in BP_flipR_001 : bp.RotZ = rightflipper.currentangle + 90 : Next
  For each bp in BP_flipR_002 : bp.RotZ = rightflipper1.currentangle + 90 : Next

  FlipperLShadow.RotZ = LeftFlipper.CurrentAngle
  FlipperRShadow.RotZ = RightFlipper.CurrentAngle
  FlipperLShadow1.RotZ = LeftFlipper1.CurrentAngle
  FlipperRShadow1.RotZ = RightFlipper1.CurrentAngle

End Sub


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


' *************************************************************************
'   JP's Reduced Display Driver Functions (based on script by Black)
' *************************************************************************


' Use FlexDMD if in FS mode
Dim UseFlexDMD
Const FlexDMDHighQuality = False
Dim FlexDMD
Dim DMDScene

Sub DMD_Init() 'default/startup values
    If UseFlexDMD Then
        Set FlexDMD = CreateObject("FlexDMD.FlexDMD")
        If Not FlexDMD is Nothing Then
            If FlexDMDHighQuality Then
                FlexDMD.TableFile = table1.Filename & ".vpx"
                FlexDMD.RenderMode = 2
                FlexDMD.Width = 256
                FlexDMD.Height = 64
                FlexDMD.Clear = True
                FlexDMD.GameName = cGameName
                FlexDMD.Run = True
                Set DMDScene = FlexDMD.NewGroup("Scene")
                DMDScene.AddActor FlexDMD.NewImage("Back", "VPX.Rockmakers DMD")
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
                DMDScene.AddActor FlexDMD.NewImage("Back", "VPX.Rockmakers DMD")
                DMDScene.GetImage("Back").SetSize FlexDMD.Width, FlexDMD.Height

                FlexDMD.LockRenderThread
                FlexDMD.Stage.AddActor DMDScene
                FlexDMD.UnlockRenderThread
            End If
        End If
    End If
End Sub

'' ===============================================================
'' The following code can be copy/pasted to have premade array for
'' movable objects:
'' - _LM suffixed arrays contains the lightmaps
'' - _BM suffixed arrays contains the bakemap
'' - _BL suffixed arrays contains both the bakemap & the lightmaps
'Dim Lsling_LM: Lsling_LM=Array(Lsling_LM_gi_lights_gi1) ' VLM.Array;LM;Lsling
'Dim Lsling_BM: Lsling_BM=Array(Lsling_BM_room) ' VLM.Array;BM;Lsling
'Dim Lsling_BL: Lsling_BL=Array(Lsling_BM_room, Lsling_LM_gi_lights_gi1) ' VLM.Array;BL;Lsling
'Dim Lsling1_LM: Lsling1_LM=Array(Lsling1_LM_gi_lights_gi1) ' VLM.Array;LM;Lsling1
'Dim Lsling1_BM: Lsling1_BM=Array(Lsling1_BM_room) ' VLM.Array;BM;Lsling1
'Dim Lsling1_BL: Lsling1_BL=Array(Lsling1_BM_room, Lsling1_LM_gi_lights_gi1) ' VLM.Array;BL;Lsling1
'Dim Lsling2_LM: Lsling2_LM=Array(Lsling2_LM_gi_lights_gi1) ' VLM.Array;LM;Lsling2
'Dim Lsling2_BM: Lsling2_BM=Array(Lsling2_BM_room) ' VLM.Array;BM;Lsling2
'Dim Lsling2_BL: Lsling2_BL=Array(Lsling2_BM_room, Lsling2_LM_gi_lights_gi1) ' VLM.Array;BL;Lsling2
'Dim Rsling_LM: Rsling_LM=Array(Rsling_LM_gi_lights_gi1) ' VLM.Array;LM;Rsling
'Dim Rsling_BM: Rsling_BM=Array(Rsling_BM_room) ' VLM.Array;BM;Rsling
'Dim Rsling_BL: Rsling_BL=Array(Rsling_BM_room, Rsling_LM_gi_lights_gi1) ' VLM.Array;BL;Rsling
'Dim Rsling1_LM: Rsling1_LM=Array(Rsling1_LM_gi_lights_gi1) ' VLM.Array;LM;Rsling1
'Dim Rsling1_BM: Rsling1_BM=Array(Rsling1_BM_room) ' VLM.Array;BM;Rsling1
'Dim Rsling1_BL: Rsling1_BL=Array(Rsling1_BM_room, Rsling1_LM_gi_lights_gi1) ' VLM.Array;BL;Rsling1
'Dim Rsling2_LM: Rsling2_LM=Array(Rsling2_LM_gi_lights_gi1) ' VLM.Array;LM;Rsling2
'Dim Rsling2_BM: Rsling2_BM=Array(Rsling2_BM_room) ' VLM.Array;BM;Rsling2
'Dim Rsling2_BL: Rsling2_BL=Array(Rsling2_BM_room, Rsling2_LM_gi_lights_gi1) ' VLM.Array;BL;Rsling2
'Dim Rubber___Post___Oval_LM: Rubber___Post___Oval_LM=Array(Rubber___Post___Oval_LM_gi_ligh) ' VLM.Array;LM;Rubber___Post___Oval
'Dim Rubber___Post___Oval_BM: Rubber___Post___Oval_BM=Array(Rubber___Post___Oval_BM_room) ' VLM.Array;BM;Rubber___Post___Oval
'Dim Rubber___Post___Oval_BL: Rubber___Post___Oval_BL=Array(Rubber___Post___Oval_BM_room, Rubber___Post___Oval_LM_gi_ligh) ' VLM.Array;BL;Rubber___Post___Oval
'Dim flipL_001_LM: flipL_001_LM=Array(flipL_001_LM_gi_lights_gi1) ' VLM.Array;LM;flipL_001
'Dim flipL_001_BM: flipL_001_BM=Array(flipL_001_BM_room) ' VLM.Array;BM;flipL_001
'Dim flipL_001_BL: flipL_001_BL=Array(flipL_001_BM_room, flipL_001_LM_gi_lights_gi1) ' VLM.Array;BL;flipL_001
'Dim flipL_002_LM: flipL_002_LM=Array(flipL_002_LM_gi_lights_gi1, flipL_002_LM_insert_lights_Lsho) ' VLM.Array;LM;flipL_002
'Dim flipL_002_BM: flipL_002_BM=Array(flipL_002_BM_room) ' VLM.Array;BM;flipL_002
'Dim flipL_002_BL: flipL_002_BL=Array(flipL_002_BM_room, flipL_002_LM_gi_lights_gi1, flipL_002_LM_insert_lights_Lsho) ' VLM.Array;BL;flipL_002
'Dim flipL_003_LM: flipL_003_LM=Array(flipL_003_LM_gi_lights_gi1) ' VLM.Array;LM;flipL_003
'Dim flipL_003_BM: flipL_003_BM=Array(flipL_003_BM_room) ' VLM.Array;BM;flipL_003
'Dim flipL_003_BL: flipL_003_BL=Array(flipL_003_BM_room, flipL_003_LM_gi_lights_gi1) ' VLM.Array;BL;flipL_003
'Dim flipR_001_LM: flipR_001_LM=Array(flipR_001_LM_gi_lights_gi1, flipR_001_LM_insert_lights_Lope, flipR_001_LM_insert_lights_Lsho) ' VLM.Array;LM;flipR_001
'Dim flipR_001_BM: flipR_001_BM=Array(flipR_001_BM_room) ' VLM.Array;BM;flipR_001
'Dim flipR_001_BL: flipR_001_BL=Array(flipR_001_BM_room, flipR_001_LM_gi_lights_gi1, flipR_001_LM_insert_lights_Lope, flipR_001_LM_insert_lights_Lsho) ' VLM.Array;BL;flipR_001
'Dim flipR_002_LM: flipR_002_LM=Array(flipR_002_LM_gi_lights_gi1, flipR_002_LM_insert_lights_Lope, flipR_002_LM_insert_lights_Lsho) ' VLM.Array;LM;flipR_002
'Dim flipR_002_BM: flipR_002_BM=Array(flipR_002_BM_room) ' VLM.Array;BM;flipR_002
'Dim flipR_002_BL: flipR_002_BL=Array(flipR_002_BM_room, flipR_002_LM_gi_lights_gi1, flipR_002_LM_insert_lights_Lope, flipR_002_LM_insert_lights_Lsho) ' VLM.Array;BL;flipR_002
'Dim gate1_LM: gate1_LM=Array(gate1_LM_gi_lights_gi1) ' VLM.Array;LM;gate1
'Dim gate1_BM: gate1_BM=Array(gate1_BM_room) ' VLM.Array;BM;gate1
'Dim gate1_BL: gate1_BL=Array(gate1_BM_room, gate1_LM_gi_lights_gi1) ' VLM.Array;BL;gate1
'Dim gate2_LM: gate2_LM=Array(gate2_LM_gi_lights_gi1) ' VLM.Array;LM;gate2
'Dim gate2_BM: gate2_BM=Array(gate2_BM_room) ' VLM.Array;BM;gate2
'Dim gate2_BL: gate2_BL=Array(gate2_BM_room, gate2_LM_gi_lights_gi1) ' VLM.Array;BL;gate2
'Dim mush1_LM: mush1_LM=Array(mush1_LM_gi_lights_gi1) ' VLM.Array;LM;mush1
'Dim mush1_BM: mush1_BM=Array(mush1_BM_room) ' VLM.Array;BM;mush1
'Dim mush1_BL: mush1_BL=Array(mush1_BM_room, mush1_LM_gi_lights_gi1) ' VLM.Array;BL;mush1
'Dim mush2_LM: mush2_LM=Array(mush2_LM_gi_lights_gi1) ' VLM.Array;LM;mush2
'Dim mush2_BM: mush2_BM=Array(mush2_BM_room) ' VLM.Array;BM;mush2
'Dim mush2_BL: mush2_BL=Array(mush2_BM_room, mush2_LM_gi_lights_gi1) ' VLM.Array;BL;mush2
'Dim mush3_LM: mush3_LM=Array(mush3_LM_gi_lights_gi1) ' VLM.Array;LM;mush3
'Dim mush3_BM: mush3_BM=Array(mush3_BM_room) ' VLM.Array;BM;mush3
'Dim mush3_BL: mush3_BL=Array(mush3_BM_room, mush3_LM_gi_lights_gi1) ' VLM.Array;BL;mush3
'Dim ring1_LM: ring1_LM=Array(ring1_LM_gi_lights_gi1) ' VLM.Array;LM;ring1
'Dim ring1_BM: ring1_BM=Array(ring1_BM_room) ' VLM.Array;BM;ring1
'Dim ring1_BL: ring1_BL=Array(ring1_BM_room, ring1_LM_gi_lights_gi1) ' VLM.Array;BL;ring1
'Dim ring2_LM: ring2_LM=Array(ring2_LM_gi_lights_gi1) ' VLM.Array;LM;ring2
'Dim ring2_BM: ring2_BM=Array(ring2_BM_room) ' VLM.Array;BM;ring2
'Dim ring2_BL: ring2_BL=Array(ring2_BM_room, ring2_LM_gi_lights_gi1) ' VLM.Array;BL;ring2
'Dim ring3_LM: ring3_LM=Array(ring3_LM_gi_lights_gi1) ' VLM.Array;LM;ring3
'Dim ring3_BM: ring3_BM=Array(ring3_BM_room) ' VLM.Array;BM;ring3
'Dim ring3_BL: ring3_BL=Array(ring3_BM_room, ring3_LM_gi_lights_gi1) ' VLM.Array;BL;ring3
'Dim skirt1_LM: skirt1_LM=Array(skirt1_LM_gi_lights_gi1) ' VLM.Array;LM;skirt1
'Dim skirt1_BM: skirt1_BM=Array(skirt1_BM_room) ' VLM.Array;BM;skirt1
'Dim skirt1_BL: skirt1_BL=Array(skirt1_BM_room, skirt1_LM_gi_lights_gi1) ' VLM.Array;BL;skirt1
'Dim skirt2_LM: skirt2_LM=Array(skirt2_LM_gi_lights_gi1) ' VLM.Array;LM;skirt2
'Dim skirt2_BM: skirt2_BM=Array(skirt2_BM_room) ' VLM.Array;BM;skirt2
'Dim skirt2_BL: skirt2_BL=Array(skirt2_BM_room, skirt2_LM_gi_lights_gi1) ' VLM.Array;BL;skirt2
'Dim skirt3_LM: skirt3_LM=Array(skirt3_LM_gi_lights_gi1) ' VLM.Array;LM;skirt3
'Dim skirt3_BM: skirt3_BM=Array(skirt3_BM_room) ' VLM.Array;BM;skirt3
'Dim skirt3_BL: skirt3_BL=Array(skirt3_BM_room, skirt3_LM_gi_lights_gi1) ' VLM.Array;BL;skirt3
'Dim sw_001_LM: sw_001_LM=Array(sw_001_LM_gi_lights_gi1) ' VLM.Array;LM;sw_001
'Dim sw_001_BM: sw_001_BM=Array(sw_001_BM_room) ' VLM.Array;BM;sw_001
'Dim sw_001_BL: sw_001_BL=Array(sw_001_BM_room, sw_001_LM_gi_lights_gi1) ' VLM.Array;BL;sw_001
'Dim sw_002_LM: sw_002_LM=Array(sw_002_LM_gi_lights_gi1, sw_002_LM_insert_lights_L200) ' VLM.Array;LM;sw_002
'Dim sw_002_BM: sw_002_BM=Array(sw_002_BM_room) ' VLM.Array;BM;sw_002
'Dim sw_002_BL: sw_002_BL=Array(sw_002_BM_room, sw_002_LM_gi_lights_gi1, sw_002_LM_insert_lights_L200) ' VLM.Array;BL;sw_002
'Dim sw_003_LM: sw_003_LM=Array(sw_003_LM_gi_lights_gi1) ' VLM.Array;LM;sw_003
'Dim sw_003_BM: sw_003_BM=Array(sw_003_BM_room) ' VLM.Array;BM;sw_003
'Dim sw_003_BL: sw_003_BL=Array(sw_003_BM_room, sw_003_LM_gi_lights_gi1) ' VLM.Array;BL;sw_003
'Dim sw_004_LM: sw_004_LM=Array(sw_004_LM_gi_lights_gi1) ' VLM.Array;LM;sw_004
'Dim sw_004_BM: sw_004_BM=Array(sw_004_BM_room) ' VLM.Array;BM;sw_004
'Dim sw_004_BL: sw_004_BL=Array(sw_004_BM_room, sw_004_LM_gi_lights_gi1) ' VLM.Array;BL;sw_004
'Dim sw_005_LM: sw_005_LM=Array(sw_005_LM_gi_lights_gi1) ' VLM.Array;LM;sw_005
'Dim sw_005_BM: sw_005_BM=Array(sw_005_BM_room) ' VLM.Array;BM;sw_005
'Dim sw_005_BL: sw_005_BL=Array(sw_005_BM_room, sw_005_LM_gi_lights_gi1) ' VLM.Array;BL;sw_005
'Dim sw_006_LM: sw_006_LM=Array(sw_006_LM_gi_lights_gi1) ' VLM.Array;LM;sw_006
'Dim sw_006_BM: sw_006_BM=Array(sw_006_BM_room) ' VLM.Array;BM;sw_006
'Dim sw_006_BL: sw_006_BL=Array(sw_006_BM_room, sw_006_LM_gi_lights_gi1) ' VLM.Array;BL;sw_006
'Dim sw_007_LM: sw_007_LM=Array(sw_007_LM_gi_lights_gi1) ' VLM.Array;LM;sw_007
'Dim sw_007_BM: sw_007_BM=Array(sw_007_BM_room) ' VLM.Array;BM;sw_007
'Dim sw_007_BL: sw_007_BL=Array(sw_007_BM_room, sw_007_LM_gi_lights_gi1) ' VLM.Array;BL;sw_007
'Dim sw_008_LM: sw_008_LM=Array(sw_008_LM_insert_lights_LKa) ' VLM.Array;LM;sw_008
'Dim sw_008_BM: sw_008_BM=Array(sw_008_BM_room) ' VLM.Array;BM;sw_008
'Dim sw_008_BL: sw_008_BL=Array(sw_008_BM_room, sw_008_LM_insert_lights_LKa) ' VLM.Array;BL;sw_008
'Dim sw_009_LM: sw_009_LM=Array(sw_009_LM_gi_lights_gi1, sw_009_LM_insert_lights_Lgate1b, sw_009_LM_insert_lights_Lopen) ' VLM.Array;LM;sw_009
'Dim sw_009_BM: sw_009_BM=Array(sw_009_BM_room) ' VLM.Array;BM;sw_009
'Dim sw_009_BL: sw_009_BL=Array(sw_009_BM_room, sw_009_LM_gi_lights_gi1, sw_009_LM_insert_lights_Lgate1b, sw_009_LM_insert_lights_Lopen) ' VLM.Array;BL;sw_009
'Dim sw_A_LM: sw_A_LM=Array(sw_A_LM_insert_lights_LA, sw_A_LM_insert_lights_LKb, sw_A_LM_insert_lights_LM) ' VLM.Array;LM;sw_A
'Dim sw_A_BM: sw_A_BM=Array(sw_A_BM_room) ' VLM.Array;BM;sw_A
'Dim sw_A_BL: sw_A_BL=Array(sw_A_BM_room, sw_A_LM_insert_lights_LA, sw_A_LM_insert_lights_LKb, sw_A_LM_insert_lights_LM) ' VLM.Array;BL;sw_A
'Dim sw_C_LM: sw_C_LM=Array(sw_C_LM_insert_lights_LC, sw_C_LM_insert_lights_LKa, sw_C_LM_insert_lights_LO) ' VLM.Array;LM;sw_C
'Dim sw_C_BM: sw_C_BM=Array(sw_C_BM_room) ' VLM.Array;BM;sw_C
'Dim sw_C_BL: sw_C_BL=Array(sw_C_BM_room, sw_C_LM_insert_lights_LC, sw_C_LM_insert_lights_LKa, sw_C_LM_insert_lights_LO) ' VLM.Array;BL;sw_C
'Dim sw_E_LM: sw_E_LM=Array(sw_E_LM_insert_lights_LE, sw_E_LM_insert_lights_LKb) ' VLM.Array;LM;sw_E
'Dim sw_E_BM: sw_E_BM=Array(sw_E_BM_room) ' VLM.Array;BM;sw_E
'Dim sw_E_BL: sw_E_BL=Array(sw_E_BM_room, sw_E_LM_insert_lights_LE, sw_E_LM_insert_lights_LKb) ' VLM.Array;BL;sw_E
'Dim sw_Ka_LM: sw_Ka_LM=Array(sw_Ka_LM_insert_lights_LC, sw_Ka_LM_insert_lights_LKa) ' VLM.Array;LM;sw_Ka
'Dim sw_Ka_BM: sw_Ka_BM=Array(sw_Ka_BM_room) ' VLM.Array;BM;sw_Ka
'Dim sw_Ka_BL: sw_Ka_BL=Array(sw_Ka_BM_room, sw_Ka_LM_insert_lights_LC, sw_Ka_LM_insert_lights_LKa) ' VLM.Array;BL;sw_Ka
'Dim sw_Kb_LM: sw_Kb_LM=Array(sw_Kb_LM_insert_lights_LA, sw_Kb_LM_insert_lights_LE, sw_Kb_LM_insert_lights_LKb) ' VLM.Array;LM;sw_Kb
'Dim sw_Kb_BM: sw_Kb_BM=Array(sw_Kb_BM_room) ' VLM.Array;BM;sw_Kb
'Dim sw_Kb_BL: sw_Kb_BL=Array(sw_Kb_BM_room, sw_Kb_LM_insert_lights_LA, sw_Kb_LM_insert_lights_LE, sw_Kb_LM_insert_lights_LKb) ' VLM.Array;BL;sw_Kb
'Dim sw_M_LM: sw_M_LM=Array(sw_M_LM_insert_lights_LA, sw_M_LM_insert_lights_LKa, sw_M_LM_insert_lights_LM) ' VLM.Array;LM;sw_M
'Dim sw_M_BM: sw_M_BM=Array(sw_M_BM_room) ' VLM.Array;BM;sw_M
'Dim sw_M_BL: sw_M_BL=Array(sw_M_BM_room, sw_M_LM_insert_lights_LA, sw_M_LM_insert_lights_LKa, sw_M_LM_insert_lights_LM) ' VLM.Array;BL;sw_M
'Dim sw_O_LM: sw_O_LM=Array(sw_O_LM_insert_lights_LC, sw_O_LM_insert_lights_LO, sw_O_LM_insert_lights_LRa) ' VLM.Array;LM;sw_O
'Dim sw_O_BM: sw_O_BM=Array(sw_O_BM_room) ' VLM.Array;BM;sw_O
'Dim sw_O_BL: sw_O_BL=Array(sw_O_BM_room, sw_O_LM_insert_lights_LC, sw_O_LM_insert_lights_LO, sw_O_LM_insert_lights_LRa) ' VLM.Array;BL;sw_O
'Dim sw_Ra_LM: sw_Ra_LM=Array(sw_Ra_LM_gi_lights_gi1, sw_Ra_LM_insert_lights_LO, sw_Ra_LM_insert_lights_LRa) ' VLM.Array;LM;sw_Ra
'Dim sw_Ra_BM: sw_Ra_BM=Array(sw_Ra_BM_room) ' VLM.Array;BM;sw_Ra
'Dim sw_Ra_BL: sw_Ra_BL=Array(sw_Ra_BM_room, sw_Ra_LM_gi_lights_gi1, sw_Ra_LM_insert_lights_LO, sw_Ra_LM_insert_lights_LRa) ' VLM.Array;BL;sw_Ra
'Dim sw_Rb_LM: sw_Rb_LM=Array(sw_Rb_LM_insert_lights_LE, sw_Rb_LM_insert_lights_LRb, sw_Rb_LM_insert_lights_LS) ' VLM.Array;LM;sw_Rb
'Dim sw_Rb_BM: sw_Rb_BM=Array(sw_Rb_BM_room) ' VLM.Array;BM;sw_Rb
'Dim sw_Rb_BL: sw_Rb_BL=Array(sw_Rb_BM_room, sw_Rb_LM_insert_lights_LE, sw_Rb_LM_insert_lights_LRb, sw_Rb_LM_insert_lights_LS) ' VLM.Array;BL;sw_Rb
'Dim sw_S_LM: sw_S_LM=Array(sw_S_LM_insert_lights_LRb, sw_S_LM_insert_lights_LS) ' VLM.Array;LM;sw_S
'Dim sw_S_BM: sw_S_BM=Array(sw_S_BM_room) ' VLM.Array;BM;sw_S
'Dim sw_S_BL: sw_S_BL=Array(sw_S_BM_room, sw_S_LM_insert_lights_LRb, sw_S_LM_insert_lights_LS) ' VLM.Array;BL;sw_S
'Dim sw_outleft_LM: sw_outleft_LM=Array(sw_outleft_LM_gi_lights_gi1, sw_outleft_LM_insert_lights_LC, sw_outleft_LM_insert_lights_LO, sw_outleft_LM_insert_lights_LRa, sw_outleft_LM_insert_lights_Lsh) ' VLM.Array;LM;sw_outleft
'Dim sw_outleft_BM: sw_outleft_BM=Array(sw_outleft_BM_room) ' VLM.Array;BM;sw_outleft
'Dim sw_outleft_BL: sw_outleft_BL=Array(sw_outleft_BM_room, sw_outleft_LM_gi_lights_gi1, sw_outleft_LM_insert_lights_LC, sw_outleft_LM_insert_lights_LO, sw_outleft_LM_insert_lights_LRa, sw_outleft_LM_insert_lights_Lsh) ' VLM.Array;BL;sw_outleft
'Dim sw_outright_LM: sw_outright_LM=Array(sw_outright_LM_gi_lights_gi1, sw_outright_LM_insert_lights_LE, sw_outright_LM_insert_lights_Lg, sw_outright_LM_insert_lights_Lo) ' VLM.Array;LM;sw_outright
'Dim sw_outright_BM: sw_outright_BM=Array(sw_outright_BM_room) ' VLM.Array;BM;sw_outright
'Dim sw_outright_BL: sw_outright_BL=Array(sw_outright_BM_room, sw_outright_LM_gi_lights_gi1, sw_outright_LM_insert_lights_LE, sw_outright_LM_insert_lights_Lg, sw_outright_LM_insert_lights_Lo) ' VLM.Array;BL;sw_outright
'
'
'
'' ===============================================================
'' The following code can be copy/pasted to adjust all the lightmaps
'' attached to a given light
'
'Sub LightHelper
' Dim l100_LM: l100_LM=Array(playfield_LM_insert_lights_L100a, german_version_LM_insert_lights_L100a, playfield_LM_insert_lights_L100b, german_version_LM_insert_lights_L100b)
' Dim l200_LM: l200_LM=Array(playfield_LM_insert_lights_L200, sw_002_LM_insert_lights_L200, german_version_LM_insert_lights_L200)
' Dim lLA_LM: lLA_LM=Array()
' Dim lLC_LM: lLC_LM=Array()
' Dim lLE_LM: lLE_LM=Array()
' Dim lLKa_LM: lLKa_LM=Array()
' Dim lLKb_LM: lLKb_LM=Array()
' Dim lLM_LM: lLM_LM=Array()
' Dim lLO_LM: lLO_LM=Array()
' Dim lLRa_LM: lLRa_LM=Array()
' Dim lLRb_LM: lLRb_LM=Array()
' Dim lLS_LM: lLS_LM=Array()
' Dim lLgate1a_LM: lLgate1a_LM=Array()
' Dim lLgate1b_LM: lLgate1b_LM=Array()
' Dim lLopen_LM: lLopen_LM=Array()
' Dim lLshootagain_LM: lLshootagain_LM=Array()
' Dim lgi1_LM: lgi1_LM=Array()
'End Sub

' ===============================================================
' The following code can serve as a base for movable position synchronization.
' You will need to adapt the part of the transform you want to synchronize
' and the source on which you want it to be synchronized.

'Sub MovableHelper
' Lsling1_BM_room.visible = x1 ' VLM.Props;BM;1;Lsling1
' Lsling1_LM_gi_lights_gi1.visible = x1 ' VLM.Props;LM;1;Lsling1
' Lsling2_BM_room.visible = x2 ' VLM.Props;BM;1;Lsling2
' Lsling2_LM_gi_lights_gi1.visible = x2 ' VLM.Props;LM;1;Lsling2
' Rsling1_BM_room.visible = x1 ' VLM.Props;BM;1;Rsling1
' Rsling1_LM_gi_lights_gi1.visible = x1 ' VLM.Props;LM;1;Rsling1
' Rsling2_BM_room.visible = x2 ' VLM.Props;BM;1;Rsling2
' Rsling2_LM_gi_lights_gi1.visible = x2 ' VLM.Props;LM;1;Rsling2
' Rubber___Post___Oval_BM_room.rotz ' VLM.Props;BM;1;Rubber - Post - Oval
' Rubber___Post___Oval_LM_gi_ligh.rotz ' VLM.Props;LM;1;Rubber - Post - Oval
' flipL_001_BM_room.rotz =  ' VLM.Props;BM;1;flipL.001
' flipL_001_LM_gi_lights_gi1.rotz =  ' VLM.Props;LM;1;flipL.001
' flipL_002_BM_room.rotz =  ' VLM.Props;BM;1;flipL.002
' flipL_002_LM_insert_lights_Lsho.rotz =  ' VLM.Props;LM;1;flipL.002
' flipL_002_LM_gi_lights_gi1.rotz =  ' VLM.Props;LM;1;flipL.002
' flipL_003_BM_room.rotz = ' VLM.Props;BM;1;flipL.003
' flipL_003_LM_gi_lights_gi1.rotz = ' VLM.Props;LM;1;flipL.003
' flipR_001_BM_room.rotz =  ' VLM.Props;BM;1;flipR.001
' flipR_001_LM_insert_lights_Lope.rotz =  ' VLM.Props;LM;1;flipR.001
' flipR_001_LM_insert_lights_Lsho.rotz =  ' VLM.Props;LM;1;flipR.001
' flipR_001_LM_gi_lights_gi1.rotz =  ' VLM.Props;LM;1;flipR.001
' flipR_002_BM_room.rotz =  ' VLM.Props;BM;1;flipR.002
' flipR_002_LM_insert_lights_Lope.rotz =  ' VLM.Props;LM;1;flipR.002
' flipR_002_LM_insert_lights_Lsho.rotz =  ' VLM.Props;LM;1;flipR.002
' flipR_002_LM_gi_lights_gi1.rotz =  ' VLM.Props;LM;1;flipR.002
' ring1_BM_room.Z = z ' VLM.Props;BM;1;ring1
' ring1_LM_gi_lights_gi1.Z = z ' VLM.Props;LM;1;ring1
' ring2_BM_room.Z = z ' VLM.Props;BM;1;ring2
' ring2_LM_gi_lights_gi1.Z = z ' VLM.Props;LM;1;ring2
' ring3_BM_room.Z = z ' VLM.Props;BM;1;ring3
' ring3_LM_gi_lights_gi1.Z = z ' VLM.Props;LM;1;ring3
' skirt1_BM_room.roty ' VLM.Props;BM;1;skirt1
' skirt1_LM_gi_lights_gi1.roty ' VLM.Props;LM;1;skirt1
' skirt2_BM_room.roty ' VLM.Props;BM;1;skirt2
' skirt2_LM_gi_lights_gi1.roty ' VLM.Props;LM;1;skirt2
' skirt3_BM_room.roty ' VLM.Props;BM;1;skirt3
' skirt3_LM_gi_lights_gi1.roty ' VLM.Props;LM;1;skirt3
' sw_outleft_BM_room.transz = z ' VLM.Props;BM;1;sw.outleft
' sw_outleft_LM_insert_lights_LC.transz = z ' VLM.Props;LM;1;sw.outleft
' sw_outleft_LM_insert_lights_LO.transz = z ' VLM.Props;LM;1;sw.outleft
' sw_outleft_LM_insert_lights_LRa.transz = z ' VLM.Props;LM;1;sw.outleft
' sw_outleft_LM_insert_lights_Lsh.transz = z ' VLM.Props;LM;1;sw.outleft
' sw_outleft_LM_gi_lights_gi1.transz = z ' VLM.Props;LM;1;sw.outleft
' sw_outright_BM_room.transz = z ' VLM.Props;BM;1;sw.outright
' sw_outright_LM_insert_lights_LE.transz = z ' VLM.Props;LM;1;sw.outright
' sw_outright_LM_insert_lights_Lg.transz = z ' VLM.Props;LM;1;sw.outright
' sw_outright_LM_insert_lights_Lo.transz = z ' VLM.Props;LM;1;sw.outright
' sw_outright_LM_gi_lights_gi1.transz = z ' VLM.Props;LM;1;sw.outright
'End Sub

' ===============================================================
' The following provides a basic synchronization mechanism were
' lightmaps are synchronized to corresponding VPX light or flasher,
' using a simple realtime timer called VLMTimer. This works great
' as a starting point but Lampz direct lightmap fading shoudl be prefered.

Sub VLMTimer_Timer
  playfield_LM_insert_lights_L100.Visible = False
  german_version_LM_insert_lights.Visible = False
  playfield_LM_insert_lights_L100.Visible = False
  german_version_LM_insert_lights.Visible = False
  playfield_LM_insert_lights_L200.Visible = False
  sw_002_LM_insert_lights_L200.Visible = False
  german_version_LM_insert_lights.Visible = False
  playfield_LM_insert_lights_LA.Visible = False
  sw_A_LM_insert_lights_LA.Visible = False
  sw_Kb_LM_insert_lights_LA.Visible = False
  sw_M_LM_insert_lights_LA.Visible = False
  german_version_LM_insert_lights.Visible = False
  playfield_LM_insert_lights_LC.Visible = False
  sw_C_LM_insert_lights_LC.Visible = False
  sw_Ka_LM_insert_lights_LC.Visible = False
  sw_O_LM_insert_lights_LC.Visible = False
  sw_outleft_LM_insert_lights_LC.Visible = False
  german_version_LM_insert_lights.Visible = False
  playfield_LM_insert_lights_LE.Visible = False
  sw_E_LM_insert_lights_LE.Visible = False
  sw_Kb_LM_insert_lights_LE.Visible = False
  sw_Rb_LM_insert_lights_LE.Visible = False
  sw_outright_LM_insert_lights_LE.Visible = False
  german_version_LM_insert_lights.Visible = False
  playfield_LM_insert_lights_LKa.Visible = False
  sw_008_LM_insert_lights_LKa.Visible = False
  sw_C_LM_insert_lights_LKa.Visible = False
  sw_Ka_LM_insert_lights_LKa.Visible = False
  sw_M_LM_insert_lights_LKa.Visible = False
  german_version_LM_insert_lights.Visible = False
  playfield_LM_insert_lights_LKb.Visible = False
  sw_A_LM_insert_lights_LKb.Visible = False
  sw_E_LM_insert_lights_LKb.Visible = False
  sw_Kb_LM_insert_lights_LKb.Visible = False
  german_version_LM_insert_lights.Visible = False
  playfield_LM_insert_lights_LM.Visible = False
  sw_A_LM_insert_lights_LM.Visible = False
  sw_M_LM_insert_lights_LM.Visible = False
  german_version_LM_insert_lights.Visible = False
  playfield_LM_insert_lights_LO.Visible = False
  sw_C_LM_insert_lights_LO.Visible = False
  sw_O_LM_insert_lights_LO.Visible = False
  sw_Ra_LM_insert_lights_LO.Visible = False
  sw_outleft_LM_insert_lights_LO.Visible = False
  german_version_LM_insert_lights.Visible = False
  playfield_LM_insert_lights_LRa.Visible = False
  sw_O_LM_insert_lights_LRa.Visible = False
  sw_Ra_LM_insert_lights_LRa.Visible = False
  sw_outleft_LM_insert_lights_LRa.Visible = False
  german_version_LM_insert_lights.Visible = False
  playfield_LM_insert_lights_LRb.Visible = False
  sw_Rb_LM_insert_lights_LRb.Visible = False
  sw_S_LM_insert_lights_LRb.Visible = False
  german_version_LM_insert_lights.Visible = False
  playfield_LM_insert_lights_LS.Visible = False
  sw_Rb_LM_insert_lights_LS.Visible = False
  sw_S_LM_insert_lights_LS.Visible = False
  german_version_LM_insert_lights.Visible = False
  playfield_LM_insert_lights_Lgat.Visible = False
  german_version_LM_insert_lights.Visible = False
  playfield_LM_insert_lights_Lgat.Visible = False
  sw_009_LM_insert_lights_Lgate1b.Visible = False
  sw_outright_LM_insert_lights_Lg.Visible = False
  german_version_LM_insert_lights.Visible = False
  playfield_LM_insert_lights_Lope.Visible = False
  flipR_001_LM_insert_lights_Lope.Visible = False
  flipR_002_LM_insert_lights_Lope.Visible = False
  sw_009_LM_insert_lights_Lopen.Visible = False
  sw_outright_LM_insert_lights_Lo.Visible = False
  german_version_LM_insert_lights.Visible = False
  playfield_LM_insert_lights_Lsho.Visible = False
  flipL_002_LM_insert_lights_Lsho.Visible = False
  flipR_001_LM_insert_lights_Lsho.Visible = False
  flipR_002_LM_insert_lights_Lsho.Visible = False
  sw_outleft_LM_insert_lights_Lsh.Visible = False
  german_version_LM_insert_lights.Visible = False
  playfield_LM_gi_lights_gi1.Visible = False
  Lsling_LM_gi_lights_gi1.Visible = False
  Lsling1_LM_gi_lights_gi1.Visible = False
  Lsling2_LM_gi_lights_gi1.Visible = False
  Rsling_LM_gi_lights_gi1.Visible = False
  Rsling1_LM_gi_lights_gi1.Visible = False
  Rsling2_LM_gi_lights_gi1.Visible = False
  Rubber___Post___Oval_LM_gi_ligh.Visible = False
  flipL_001_LM_gi_lights_gi1.Visible = False
  flipL_002_LM_gi_lights_gi1.Visible = False
  flipL_003_LM_gi_lights_gi1.Visible = False
  flipR_001_LM_gi_lights_gi1.Visible = False
  flipR_002_LM_gi_lights_gi1.Visible = False
  gate1_LM_gi_lights_gi1.Visible = False
  gate2_LM_gi_lights_gi1.Visible = False
  mush1_LM_gi_lights_gi1.Visible = False
  mush2_LM_gi_lights_gi1.Visible = False
  mush3_LM_gi_lights_gi1.Visible = False
  ring1_LM_gi_lights_gi1.Visible = False
  ring2_LM_gi_lights_gi1.Visible = False
  ring3_LM_gi_lights_gi1.Visible = False
  skirt1_LM_gi_lights_gi1.Visible = False
  skirt2_LM_gi_lights_gi1.Visible = False
  skirt3_LM_gi_lights_gi1.Visible = False
  sw_001_LM_gi_lights_gi1.Visible = False
  sw_002_LM_gi_lights_gi1.Visible = False
  sw_003_LM_gi_lights_gi1.Visible = False
  sw_004_LM_gi_lights_gi1.Visible = False
  sw_005_LM_gi_lights_gi1.Visible = False
  sw_006_LM_gi_lights_gi1.Visible = False
  sw_007_LM_gi_lights_gi1.Visible = False
  sw_009_LM_gi_lights_gi1.Visible = False
  sw_Ra_LM_gi_lights_gi1.Visible = False
  sw_outleft_LM_gi_lights_gi1.Visible = False
  sw_outright_LM_gi_lights_gi1.Visible = False
  plastics_layer1_LM_gi_lights_gi.Visible = False
  plastics_layer2_LM_gi_lights_gi.Visible = False
  german_version_LM_gi_lights_gi1.Visible = False
End Sub

Function LightFade(light, is_on, percent)
  If is_on Then
    LightFade = percent*percent*(3 - 2*percent) ' Smoothstep
  Else
    LightFade = 1 - Sqr(1 - percent*percent) '
  End If
End Function

Sub UpdateLightMapFromFlasher(flasher, lightmap, intensity_scale, sync_color)
  If flasher.Visible Then
    If sync_color Then lightmap.Color = flasher.Color
    lightmap.Opacity = intensity_scale * flasher.IntensityScale * flasher.Opacity / 1000.0
  Else
    lightmap.Opacity = 0
  End If
End Sub

Sub UpdateLightMapFromLight(light, lightmap, intensity_scale, sync_color)
  light.FadeSpeedUp = light.Intensity / 50 '100
  light.FadeSpeedDown = light.Intensity / 200
  If sync_color Then lightmap.Color = light.Colorfull
  Dim t: t = LightFade(light, light.GetInPlayStateBool(), light.GetInPlayIntensity() / (light.Intensity * light.IntensityScale))
  lightmap.Opacity = intensity_scale * light.IntensityScale * t
End Sub

' Thalamus : Exit in a clean and proper way
Sub Table1_exit
  Controller.Pause = False
  Controller.Stop
End Sub

' VLM Arrays - Start
' Arrays per baked part
Dim BP_Circle_019: BP_Circle_019=Array(BM_Circle_019, LM_gi_gi3_Circle_019)
Dim BP_Circle_024: BP_Circle_024=Array(BM_Circle_024, LM_gi_gi3_Circle_024)
Dim BP_Circle_031: BP_Circle_031=Array(BM_Circle_031, LM_gi_gi4_Circle_031, LM_gi_gi3_Circle_031)
Dim BP_DifficultyA: BP_DifficultyA=Array(BM_DifficultyA, LM_ins_Lgate1a_DifficultyA, LM_ins_Lgate1b_DifficultyA, LM_gi_gi5_DifficultyA, LM_gi_gi4_DifficultyA, LM_gi_gi1_DifficultyA)
Dim BP_DifficultyB: BP_DifficultyB=Array(BM_DifficultyB, LM_ins_Lgate1a_DifficultyB, LM_ins_Lgate1b_DifficultyB, LM_gi_gi5_DifficultyB, LM_gi_gi4_DifficultyB, LM_gi_gi1_DifficultyB)
Dim BP_Lsling: BP_Lsling=Array(BM_Lsling, LM_gi_gi4_Lsling)
Dim BP_Lsling1: BP_Lsling1=Array(BM_Lsling1, LM_gi_gi4_Lsling1)
Dim BP_Lsling2: BP_Lsling2=Array(BM_Lsling2, LM_gi_gi4_Lsling2)
Dim BP_Rsling: BP_Rsling=Array(BM_Rsling, LM_gi_gi5_Rsling)
Dim BP_Rsling1: BP_Rsling1=Array(BM_Rsling1, LM_gi_gi5_Rsling1)
Dim BP_Rsling2: BP_Rsling2=Array(BM_Rsling2, LM_gi_gi5_Rsling2)
Dim BP_Rubber___Post___Oval_001: BP_Rubber___Post___Oval_001=Array(BM_Rubber___Post___Oval_001, LM_gi_gi5_Rubber___Post___Oval_)
Dim BP_flipL_001: BP_flipL_001=Array(BM_flipL_001, LM_gi_gi4_flipL_001)
Dim BP_flipL_002: BP_flipL_002=Array(BM_flipL_002, LM_gi_gi4_flipL_002)
Dim BP_flipL_003: BP_flipL_003=Array(BM_flipL_003, LM_gi_gi2_flipL_003, LM_gi_gi1_flipL_003)
Dim BP_flipR_001: BP_flipR_001=Array(BM_flipR_001, LM_gi_gi5_flipR_001)
Dim BP_flipR_002: BP_flipR_002=Array(BM_flipR_002, LM_gi_gi5_flipR_002, LM_gi_gi4_flipR_002)
Dim BP_gate1: BP_gate1=Array(BM_gate1)
Dim BP_gate2: BP_gate2=Array(BM_gate2, LM_gi_gi1_gate2)
Dim BP_german_version: BP_german_version=Array(BM_german_version, LM_gi_gi1_german_version)
Dim BP_mush1: BP_mush1=Array(BM_mush1)
Dim BP_mush2: BP_mush2=Array(BM_mush2, LM_gi_gi1_mush2)
Dim BP_mush3: BP_mush3=Array(BM_mush3, LM_gi_gi1_mush3)
Dim BP_non_movable: BP_non_movable=Array(BM_non_movable, LM_ins_Lrar0010_non_movable, LM_ins_Lgate1a_non_movable, LM_ins_Lgate1b_non_movable, LM_gi_gi5_non_movable, LM_gi_gi4_non_movable, LM_gi_gi3_non_movable, LM_gi_gi2_non_movable, LM_gi_gi1_non_movable)
Dim BP_pgate_001: BP_pgate_001=Array(BM_pgate_001)
Dim BP_plastics_layer1: BP_plastics_layer1=Array(BM_plastics_layer1, LM_gi_gi2_plastics_layer1, LM_gi_gi1_plastics_layer1)
Dim BP_plastics_layer2: BP_plastics_layer2=Array(BM_plastics_layer2, LM_ins_Lgate1b_plastics_layer2, LM_gi_gi5_plastics_layer2, LM_gi_gi4_plastics_layer2, LM_gi_gi3_plastics_layer2, LM_gi_gi2_plastics_layer2, LM_gi_gi1_plastics_layer2)
Dim BP_playfield: BP_playfield=Array(BM_playfield, LM_ins_L100a_playfield, LM_ins_L100b_playfield, LM_ins_L200_playfield, LM_ins_Lrar006_playfield, LM_ins_Lrar003_playfield, LM_ins_Lrar008_playfield, LM_ins_Lrar004_playfield, LM_ins_Lrar007_playfield, LM_ins_Lrar005_playfield, LM_ins_Lrar002_playfield, LM_ins_Lrar001_playfield, LM_ins_Lrar009_playfield, LM_ins_Lrar0010_playfield, LM_ins_Lgate1a_playfield, LM_ins_Lgate1b_playfield, LM_ins_Lopen_playfield, LM_ins_Lshootagain_playfield, LM_gi_gi5_playfield, LM_gi_gi4_playfield, LM_gi_gi3_playfield, LM_gi_gi2_playfield, LM_gi_gi1_playfield)
Dim BP_plywood: BP_plywood=Array(BM_plywood, LM_ins_L200_plywood, LM_ins_Lrar006_plywood, LM_ins_Lrar005_plywood, LM_ins_Lrar0010_plywood, LM_ins_Lgate1b_plywood, LM_ins_Lopen_plywood, LM_gi_gi5_plywood, LM_gi_gi4_plywood, LM_gi_gi3_plywood, LM_gi_gi1_plywood)
Dim BP_rails: BP_rails=Array(BM_rails, LM_gi_gi4_rails, LM_gi_gi3_rails, LM_gi_gi1_rails)
Dim BP_skirt1: BP_skirt1=Array(BM_skirt1, LM_gi_gi1_skirt1)
Dim BP_skirt2: BP_skirt2=Array(BM_skirt2, LM_gi_gi1_skirt2)
Dim BP_skirt3: BP_skirt3=Array(BM_skirt3, LM_gi_gi1_skirt3)
Dim BP_sw_001: BP_sw_001=Array(BM_sw_001, LM_gi_gi2_sw_001, LM_gi_gi1_sw_001)
Dim BP_sw_002: BP_sw_002=Array(BM_sw_002, LM_gi_gi2_sw_002)
Dim BP_sw_003: BP_sw_003=Array(BM_sw_003, LM_gi_gi2_sw_003, LM_gi_gi1_sw_003)
Dim BP_sw_004: BP_sw_004=Array(BM_sw_004, LM_gi_gi1_sw_004)
Dim BP_sw_005: BP_sw_005=Array(BM_sw_005, LM_gi_gi1_sw_005)
Dim BP_sw_006: BP_sw_006=Array(BM_sw_006, LM_gi_gi1_sw_006)
Dim BP_sw_007: BP_sw_007=Array(BM_sw_007, LM_gi_gi1_sw_007)
Dim BP_sw_008: BP_sw_008=Array(BM_sw_008, LM_ins_Lrar006_sw_008, LM_ins_Lrar004_sw_008, LM_ins_Lrar007_sw_008, LM_ins_Lrar005_sw_008)
Dim BP_sw_009: BP_sw_009=Array(BM_sw_009, LM_ins_Lgate1b_sw_009, LM_ins_Lopen_sw_009, LM_gi_gi5_sw_009)
Dim BP_swr_001: BP_swr_001=Array(BM_swr_001, LM_ins_Lrar002_swr_001, LM_ins_Lrar001_swr_001, LM_gi_gi3_swr_001)
Dim BP_swr_002: BP_swr_002=Array(BM_swr_002, LM_ins_Lrar003_swr_002, LM_ins_Lrar002_swr_002, LM_ins_Lrar001_swr_002, LM_gi_gi3_swr_002)
Dim BP_swr_003: BP_swr_003=Array(BM_swr_003, LM_ins_Lrar003_swr_003, LM_ins_Lrar004_swr_003, LM_ins_Lrar002_swr_003)
Dim BP_swr_004: BP_swr_004=Array(BM_swr_004, LM_ins_Lrar003_swr_004, LM_ins_Lrar004_swr_004)
Dim BP_swr_005: BP_swr_005=Array(BM_swr_005, LM_ins_Lrar006_swr_005, LM_ins_Lrar004_swr_005, LM_ins_Lrar005_swr_005)
Dim BP_swr_006: BP_swr_006=Array(BM_swr_006, LM_ins_Lrar006_swr_006, LM_ins_Lrar007_swr_006, LM_ins_Lrar005_swr_006)
Dim BP_swr_007: BP_swr_007=Array(BM_swr_007, LM_ins_Lrar006_swr_007, LM_ins_Lrar008_swr_007, LM_ins_Lrar007_swr_007)
Dim BP_swr_008: BP_swr_008=Array(BM_swr_008, LM_ins_Lrar008_swr_008, LM_ins_Lrar007_swr_008, LM_ins_Lrar009_swr_008)
Dim BP_swr_009: BP_swr_009=Array(BM_swr_009, LM_ins_Lrar008_swr_009, LM_ins_Lrar009_swr_009, LM_ins_Lrar0010_swr_009)
Dim BP_swr_010: BP_swr_010=Array(BM_swr_010, LM_ins_Lrar0010_swr_010)
Dim BP_wire_001: BP_wire_001=Array(BM_wire_001, LM_gi_gi4_wire_001)
Dim BP_wire_002: BP_wire_002=Array(BM_wire_002, LM_ins_Lgate1b_wire_002, LM_gi_gi5_wire_002)
' Arrays per lighting scenario
Dim BL_gi_gi1: BL_gi_gi1=Array(LM_gi_gi1_DifficultyA, LM_gi_gi1_DifficultyB, LM_gi_gi1_flipL_003, LM_gi_gi1_gate2, LM_gi_gi1_german_version, LM_gi_gi1_mush2, LM_gi_gi1_mush3, LM_gi_gi1_non_movable, LM_gi_gi1_plastics_layer1, LM_gi_gi1_plastics_layer2, LM_gi_gi1_playfield, LM_gi_gi1_plywood, LM_gi_gi1_rails, LM_gi_gi1_skirt1, LM_gi_gi1_skirt2, LM_gi_gi1_skirt3, LM_gi_gi1_sw_001, LM_gi_gi1_sw_003, LM_gi_gi1_sw_004, LM_gi_gi1_sw_005, LM_gi_gi1_sw_006, LM_gi_gi1_sw_007)
Dim BL_gi_gi2: BL_gi_gi2=Array(LM_gi_gi2_flipL_003, LM_gi_gi2_non_movable, LM_gi_gi2_plastics_layer1, LM_gi_gi2_plastics_layer2, LM_gi_gi2_playfield, LM_gi_gi2_sw_001, LM_gi_gi2_sw_002, LM_gi_gi2_sw_003)
Dim BL_gi_gi3: BL_gi_gi3=Array(LM_gi_gi3_Circle_019, LM_gi_gi3_Circle_024, LM_gi_gi3_Circle_031, LM_gi_gi3_non_movable, LM_gi_gi3_plastics_layer2, LM_gi_gi3_playfield, LM_gi_gi3_plywood, LM_gi_gi3_rails, LM_gi_gi3_swr_001, LM_gi_gi3_swr_002)
Dim BL_gi_gi4: BL_gi_gi4=Array(LM_gi_gi4_Circle_031, LM_gi_gi4_DifficultyA, LM_gi_gi4_DifficultyB, LM_gi_gi4_Lsling, LM_gi_gi4_Lsling1, LM_gi_gi4_Lsling2, LM_gi_gi4_flipL_001, LM_gi_gi4_flipL_002, LM_gi_gi4_flipR_002, LM_gi_gi4_non_movable, LM_gi_gi4_plastics_layer2, LM_gi_gi4_playfield, LM_gi_gi4_plywood, LM_gi_gi4_rails, LM_gi_gi4_wire_001)
Dim BL_gi_gi5: BL_gi_gi5=Array(LM_gi_gi5_DifficultyA, LM_gi_gi5_DifficultyB, LM_gi_gi5_Rsling, LM_gi_gi5_Rsling1, LM_gi_gi5_Rsling2, LM_gi_gi5_Rubber___Post___Oval_, LM_gi_gi5_flipR_001, LM_gi_gi5_flipR_002, LM_gi_gi5_non_movable, LM_gi_gi5_plastics_layer2, LM_gi_gi5_playfield, LM_gi_gi5_plywood, LM_gi_gi5_sw_009, LM_gi_gi5_wire_002)
Dim BL_ins_L100a: BL_ins_L100a=Array(LM_ins_L100a_playfield)
Dim BL_ins_L100b: BL_ins_L100b=Array(LM_ins_L100b_playfield)
Dim BL_ins_L200: BL_ins_L200=Array(LM_ins_L200_playfield, LM_ins_L200_plywood)
Dim BL_ins_Lgate1a: BL_ins_Lgate1a=Array(LM_ins_Lgate1a_DifficultyA, LM_ins_Lgate1a_DifficultyB, LM_ins_Lgate1a_non_movable, LM_ins_Lgate1a_playfield)
Dim BL_ins_Lgate1b: BL_ins_Lgate1b=Array(LM_ins_Lgate1b_DifficultyA, LM_ins_Lgate1b_DifficultyB, LM_ins_Lgate1b_non_movable, LM_ins_Lgate1b_plastics_layer2, LM_ins_Lgate1b_playfield, LM_ins_Lgate1b_plywood, LM_ins_Lgate1b_sw_009, LM_ins_Lgate1b_wire_002)
Dim BL_ins_Lopen: BL_ins_Lopen=Array(LM_ins_Lopen_playfield, LM_ins_Lopen_plywood, LM_ins_Lopen_sw_009)
Dim BL_ins_Lrar001: BL_ins_Lrar001=Array(LM_ins_Lrar001_playfield, LM_ins_Lrar001_swr_001, LM_ins_Lrar001_swr_002)
Dim BL_ins_Lrar0010: BL_ins_Lrar0010=Array(LM_ins_Lrar0010_non_movable, LM_ins_Lrar0010_playfield, LM_ins_Lrar0010_plywood, LM_ins_Lrar0010_swr_009, LM_ins_Lrar0010_swr_010)
Dim BL_ins_Lrar002: BL_ins_Lrar002=Array(LM_ins_Lrar002_playfield, LM_ins_Lrar002_swr_001, LM_ins_Lrar002_swr_002, LM_ins_Lrar002_swr_003)
Dim BL_ins_Lrar003: BL_ins_Lrar003=Array(LM_ins_Lrar003_playfield, LM_ins_Lrar003_swr_002, LM_ins_Lrar003_swr_003, LM_ins_Lrar003_swr_004)
Dim BL_ins_Lrar004: BL_ins_Lrar004=Array(LM_ins_Lrar004_playfield, LM_ins_Lrar004_sw_008, LM_ins_Lrar004_swr_003, LM_ins_Lrar004_swr_004, LM_ins_Lrar004_swr_005)
Dim BL_ins_Lrar005: BL_ins_Lrar005=Array(LM_ins_Lrar005_playfield, LM_ins_Lrar005_plywood, LM_ins_Lrar005_sw_008, LM_ins_Lrar005_swr_005, LM_ins_Lrar005_swr_006)
Dim BL_ins_Lrar006: BL_ins_Lrar006=Array(LM_ins_Lrar006_playfield, LM_ins_Lrar006_plywood, LM_ins_Lrar006_sw_008, LM_ins_Lrar006_swr_005, LM_ins_Lrar006_swr_006, LM_ins_Lrar006_swr_007)
Dim BL_ins_Lrar007: BL_ins_Lrar007=Array(LM_ins_Lrar007_playfield, LM_ins_Lrar007_sw_008, LM_ins_Lrar007_swr_006, LM_ins_Lrar007_swr_007, LM_ins_Lrar007_swr_008)
Dim BL_ins_Lrar008: BL_ins_Lrar008=Array(LM_ins_Lrar008_playfield, LM_ins_Lrar008_swr_007, LM_ins_Lrar008_swr_008, LM_ins_Lrar008_swr_009)
Dim BL_ins_Lrar009: BL_ins_Lrar009=Array(LM_ins_Lrar009_playfield, LM_ins_Lrar009_swr_008, LM_ins_Lrar009_swr_009)
Dim BL_ins_Lshootagain: BL_ins_Lshootagain=Array(LM_ins_Lshootagain_playfield)
Dim BL_room: BL_room=Array(BM_Circle_019, BM_Circle_024, BM_Circle_031, BM_DifficultyA, BM_DifficultyB, BM_Lsling, BM_Lsling1, BM_Lsling2, BM_Rsling, BM_Rsling1, BM_Rsling2, BM_Rubber___Post___Oval_001, BM_flipL_001, BM_flipL_002, BM_flipL_003, BM_flipR_001, BM_flipR_002, BM_gate1, BM_gate2, BM_german_version, BM_mush1, BM_mush2, BM_mush3, BM_non_movable, BM_pgate_001, BM_plastics_layer1, BM_plastics_layer2, BM_playfield, BM_plywood, BM_rails, BM_skirt1, BM_skirt2, BM_skirt3, BM_sw_001, BM_sw_002, BM_sw_003, BM_sw_004, BM_sw_005, BM_sw_006, BM_sw_007, BM_sw_008, BM_sw_009, BM_swr_001, BM_swr_002, BM_swr_003, BM_swr_004, BM_swr_005, BM_swr_006, BM_swr_007, BM_swr_008, BM_swr_009, BM_swr_010, BM_wire_001, BM_wire_002)
' Global arrays
Dim BG_Bakemap: BG_Bakemap=Array(BM_Circle_019, BM_Circle_024, BM_Circle_031, BM_DifficultyA, BM_DifficultyB, BM_Lsling, BM_Lsling1, BM_Lsling2, BM_Rsling, BM_Rsling1, BM_Rsling2, BM_Rubber___Post___Oval_001, BM_flipL_001, BM_flipL_002, BM_flipL_003, BM_flipR_001, BM_flipR_002, BM_gate1, BM_gate2, BM_german_version, BM_mush1, BM_mush2, BM_mush3, BM_non_movable, BM_pgate_001, BM_plastics_layer1, BM_plastics_layer2, BM_playfield, BM_plywood, BM_rails, BM_skirt1, BM_skirt2, BM_skirt3, BM_sw_001, BM_sw_002, BM_sw_003, BM_sw_004, BM_sw_005, BM_sw_006, BM_sw_007, BM_sw_008, BM_sw_009, BM_swr_001, BM_swr_002, BM_swr_003, BM_swr_004, BM_swr_005, BM_swr_006, BM_swr_007, BM_swr_008, BM_swr_009, BM_swr_010, BM_wire_001, BM_wire_002)
Dim BG_Lightmap: BG_Lightmap=Array(LM_gi_gi1_DifficultyA, LM_gi_gi1_DifficultyB, LM_gi_gi1_flipL_003, LM_gi_gi1_gate2, LM_gi_gi1_german_version, LM_gi_gi1_mush2, LM_gi_gi1_mush3, LM_gi_gi1_non_movable, LM_gi_gi1_plastics_layer1, LM_gi_gi1_plastics_layer2, LM_gi_gi1_playfield, LM_gi_gi1_plywood, LM_gi_gi1_rails, LM_gi_gi1_skirt1, LM_gi_gi1_skirt2, LM_gi_gi1_skirt3, LM_gi_gi1_sw_001, LM_gi_gi1_sw_003, LM_gi_gi1_sw_004, LM_gi_gi1_sw_005, LM_gi_gi1_sw_006, LM_gi_gi1_sw_007, LM_gi_gi2_flipL_003, LM_gi_gi2_non_movable, LM_gi_gi2_plastics_layer1, LM_gi_gi2_plastics_layer2, LM_gi_gi2_playfield, LM_gi_gi2_sw_001, LM_gi_gi2_sw_002, LM_gi_gi2_sw_003, LM_gi_gi3_Circle_019, LM_gi_gi3_Circle_024, LM_gi_gi3_Circle_031, LM_gi_gi3_non_movable, LM_gi_gi3_plastics_layer2, LM_gi_gi3_playfield, LM_gi_gi3_plywood, LM_gi_gi3_rails, LM_gi_gi3_swr_001, LM_gi_gi3_swr_002, LM_gi_gi4_Circle_031, LM_gi_gi4_DifficultyA, LM_gi_gi4_DifficultyB, LM_gi_gi4_Lsling, LM_gi_gi4_Lsling1, LM_gi_gi4_Lsling2, LM_gi_gi4_flipL_001, LM_gi_gi4_flipL_002, _
  LM_gi_gi4_flipR_002, LM_gi_gi4_non_movable, LM_gi_gi4_plastics_layer2, LM_gi_gi4_playfield, LM_gi_gi4_plywood, LM_gi_gi4_rails, LM_gi_gi4_wire_001, LM_gi_gi5_DifficultyA, LM_gi_gi5_DifficultyB, LM_gi_gi5_Rsling, LM_gi_gi5_Rsling1, LM_gi_gi5_Rsling2, LM_gi_gi5_Rubber___Post___Oval_, LM_gi_gi5_flipR_001, LM_gi_gi5_flipR_002, LM_gi_gi5_non_movable, LM_gi_gi5_plastics_layer2, LM_gi_gi5_playfield, LM_gi_gi5_plywood, LM_gi_gi5_sw_009, LM_gi_gi5_wire_002, LM_ins_L100a_playfield, LM_ins_L100b_playfield, LM_ins_L200_playfield, LM_ins_L200_plywood, LM_ins_Lgate1a_DifficultyA, LM_ins_Lgate1a_DifficultyB, LM_ins_Lgate1a_non_movable, LM_ins_Lgate1a_playfield, LM_ins_Lgate1b_DifficultyA, LM_ins_Lgate1b_DifficultyB, LM_ins_Lgate1b_non_movable, LM_ins_Lgate1b_plastics_layer2, LM_ins_Lgate1b_playfield, LM_ins_Lgate1b_plywood, LM_ins_Lgate1b_sw_009, LM_ins_Lgate1b_wire_002, LM_ins_Lopen_playfield, LM_ins_Lopen_plywood, LM_ins_Lopen_sw_009, LM_ins_Lrar001_playfield, LM_ins_Lrar001_swr_001, LM_ins_Lrar001_swr_002, _
  LM_ins_Lrar0010_non_movable, LM_ins_Lrar0010_playfield, LM_ins_Lrar0010_plywood, LM_ins_Lrar0010_swr_009, LM_ins_Lrar0010_swr_010, LM_ins_Lrar002_playfield, LM_ins_Lrar002_swr_001, LM_ins_Lrar002_swr_002, LM_ins_Lrar002_swr_003, LM_ins_Lrar003_playfield, LM_ins_Lrar003_swr_002, LM_ins_Lrar003_swr_003, LM_ins_Lrar003_swr_004, LM_ins_Lrar004_playfield, LM_ins_Lrar004_sw_008, LM_ins_Lrar004_swr_003, LM_ins_Lrar004_swr_004, LM_ins_Lrar004_swr_005, LM_ins_Lrar005_playfield, LM_ins_Lrar005_plywood, LM_ins_Lrar005_sw_008, LM_ins_Lrar005_swr_005, LM_ins_Lrar005_swr_006, LM_ins_Lrar006_playfield, LM_ins_Lrar006_plywood, LM_ins_Lrar006_sw_008, LM_ins_Lrar006_swr_005, LM_ins_Lrar006_swr_006, LM_ins_Lrar006_swr_007, LM_ins_Lrar007_playfield, LM_ins_Lrar007_sw_008, LM_ins_Lrar007_swr_006, LM_ins_Lrar007_swr_007, LM_ins_Lrar007_swr_008, LM_ins_Lrar008_playfield, LM_ins_Lrar008_swr_007, LM_ins_Lrar008_swr_008, LM_ins_Lrar008_swr_009, LM_ins_Lrar009_playfield, LM_ins_Lrar009_swr_008, LM_ins_Lrar009_swr_009, _
  LM_ins_Lshootagain_playfield)
Dim BG_All: BG_All=Array(BM_Circle_019, BM_Circle_024, BM_Circle_031, BM_DifficultyA, BM_DifficultyB, BM_Lsling, BM_Lsling1, BM_Lsling2, BM_Rsling, BM_Rsling1, BM_Rsling2, BM_Rubber___Post___Oval_001, BM_flipL_001, BM_flipL_002, BM_flipL_003, BM_flipR_001, BM_flipR_002, BM_gate1, BM_gate2, BM_german_version, BM_mush1, BM_mush2, BM_mush3, BM_non_movable, BM_pgate_001, BM_plastics_layer1, BM_plastics_layer2, BM_playfield, BM_plywood, BM_rails, BM_skirt1, BM_skirt2, BM_skirt3, BM_sw_001, BM_sw_002, BM_sw_003, BM_sw_004, BM_sw_005, BM_sw_006, BM_sw_007, BM_sw_008, BM_sw_009, BM_swr_001, BM_swr_002, BM_swr_003, BM_swr_004, BM_swr_005, BM_swr_006, BM_swr_007, BM_swr_008, BM_swr_009, BM_swr_010, BM_wire_001, BM_wire_002, LM_gi_gi1_DifficultyA, LM_gi_gi1_DifficultyB, LM_gi_gi1_flipL_003, LM_gi_gi1_gate2, LM_gi_gi1_german_version, LM_gi_gi1_mush2, LM_gi_gi1_mush3, LM_gi_gi1_non_movable, LM_gi_gi1_plastics_layer1, LM_gi_gi1_plastics_layer2, LM_gi_gi1_playfield, LM_gi_gi1_plywood, LM_gi_gi1_rails, LM_gi_gi1_skirt1, _
  LM_gi_gi1_skirt2, LM_gi_gi1_skirt3, LM_gi_gi1_sw_001, LM_gi_gi1_sw_003, LM_gi_gi1_sw_004, LM_gi_gi1_sw_005, LM_gi_gi1_sw_006, LM_gi_gi1_sw_007, LM_gi_gi2_flipL_003, LM_gi_gi2_non_movable, LM_gi_gi2_plastics_layer1, LM_gi_gi2_plastics_layer2, LM_gi_gi2_playfield, LM_gi_gi2_sw_001, LM_gi_gi2_sw_002, LM_gi_gi2_sw_003, LM_gi_gi3_Circle_019, LM_gi_gi3_Circle_024, LM_gi_gi3_Circle_031, LM_gi_gi3_non_movable, LM_gi_gi3_plastics_layer2, LM_gi_gi3_playfield, LM_gi_gi3_plywood, LM_gi_gi3_rails, LM_gi_gi3_swr_001, LM_gi_gi3_swr_002, LM_gi_gi4_Circle_031, LM_gi_gi4_DifficultyA, LM_gi_gi4_DifficultyB, LM_gi_gi4_Lsling, LM_gi_gi4_Lsling1, LM_gi_gi4_Lsling2, LM_gi_gi4_flipL_001, LM_gi_gi4_flipL_002, LM_gi_gi4_flipR_002, LM_gi_gi4_non_movable, LM_gi_gi4_plastics_layer2, LM_gi_gi4_playfield, LM_gi_gi4_plywood, LM_gi_gi4_rails, LM_gi_gi4_wire_001, LM_gi_gi5_DifficultyA, LM_gi_gi5_DifficultyB, LM_gi_gi5_Rsling, LM_gi_gi5_Rsling1, LM_gi_gi5_Rsling2, LM_gi_gi5_Rubber___Post___Oval_, LM_gi_gi5_flipR_001, LM_gi_gi5_flipR_002, _
  LM_gi_gi5_non_movable, LM_gi_gi5_plastics_layer2, LM_gi_gi5_playfield, LM_gi_gi5_plywood, LM_gi_gi5_sw_009, LM_gi_gi5_wire_002, LM_ins_L100a_playfield, LM_ins_L100b_playfield, LM_ins_L200_playfield, LM_ins_L200_plywood, LM_ins_Lgate1a_DifficultyA, LM_ins_Lgate1a_DifficultyB, LM_ins_Lgate1a_non_movable, LM_ins_Lgate1a_playfield, LM_ins_Lgate1b_DifficultyA, LM_ins_Lgate1b_DifficultyB, LM_ins_Lgate1b_non_movable, LM_ins_Lgate1b_plastics_layer2, LM_ins_Lgate1b_playfield, LM_ins_Lgate1b_plywood, LM_ins_Lgate1b_sw_009, LM_ins_Lgate1b_wire_002, LM_ins_Lopen_playfield, LM_ins_Lopen_plywood, LM_ins_Lopen_sw_009, LM_ins_Lrar001_playfield, LM_ins_Lrar001_swr_001, LM_ins_Lrar001_swr_002, LM_ins_Lrar0010_non_movable, LM_ins_Lrar0010_playfield, LM_ins_Lrar0010_plywood, LM_ins_Lrar0010_swr_009, LM_ins_Lrar0010_swr_010, LM_ins_Lrar002_playfield, LM_ins_Lrar002_swr_001, LM_ins_Lrar002_swr_002, LM_ins_Lrar002_swr_003, LM_ins_Lrar003_playfield, LM_ins_Lrar003_swr_002, LM_ins_Lrar003_swr_003, LM_ins_Lrar003_swr_004, _
  LM_ins_Lrar004_playfield, LM_ins_Lrar004_sw_008, LM_ins_Lrar004_swr_003, LM_ins_Lrar004_swr_004, LM_ins_Lrar004_swr_005, LM_ins_Lrar005_playfield, LM_ins_Lrar005_plywood, LM_ins_Lrar005_sw_008, LM_ins_Lrar005_swr_005, LM_ins_Lrar005_swr_006, LM_ins_Lrar006_playfield, LM_ins_Lrar006_plywood, LM_ins_Lrar006_sw_008, LM_ins_Lrar006_swr_005, LM_ins_Lrar006_swr_006, LM_ins_Lrar006_swr_007, LM_ins_Lrar007_playfield, LM_ins_Lrar007_sw_008, LM_ins_Lrar007_swr_006, LM_ins_Lrar007_swr_007, LM_ins_Lrar007_swr_008, LM_ins_Lrar008_playfield, LM_ins_Lrar008_swr_007, LM_ins_Lrar008_swr_008, LM_ins_Lrar008_swr_009, LM_ins_Lrar009_playfield, LM_ins_Lrar009_swr_008, LM_ins_Lrar009_swr_009, LM_ins_Lshootagain_playfield)
' VLM Arrays - End


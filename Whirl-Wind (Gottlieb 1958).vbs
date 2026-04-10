'NOTES

'if gobblelight or gobblelightred are lit then layer1d.image="layer1"
'if gobblelight or gobblelightred are NOT lit (in between roto indexes) thenlayer1d.image="layer1b"

'rototarget spins on objroty

'Texture changes for lit L/R slingshots

'cig2.visible=0 when game is only 1 player or when player 1 is up in a 2 player game

'If slingleftlit=1 Then
' For each xx in slinglampsleft:xx.State = 1: Next
' For each xx in slinglampsright:xx.State = 0: Next
' For each xx in layer1acol: xx.image = "layer1": Next
' For each xx in layer1bcol: xx.image = "layer1b": Next
' For each xx in layer2acol: xx.image = "layer2": Next
' For each xx in layer2bcol: xx.image = "layer2b": Next
' For each xx in layer3acol: xx.image = "layer3": Next
' For each xx in layer3bcol: xx.image = "layer3b": Next

'If slingrightlit=1 Then
' For each xx in slinglampsleft:xx.State = 0: Next
' For each xx in slinglampsright:xx.State = 1: Next
' For each xx in layer1acol: xx.image = "layer1b": Next
' For each xx in layer1bcol: xx.image = "layer1": Next
' For each xx in layer2acol: xx.image = "layer2b": Next
' For each xx in layer2bcol: xx.image = "layer2": Next
' For each xx in layer3acol: xx.image = "layer3b": Next
' For each xx in layer3bcol: xx.image = "layer3": Next


'dead bumpers trigger with a hit to phys_deadbumps (spin), phys_deadbumps001 (Left center), phys_deadbumps002 (Right center)

'1 pt switches trigger with scoringswitches1pt and scoringswitches.


'****************************************************************
'
'          Whirl-Wind (Gottlieb 1958)
'          Script by Scottacus
'             v 1.0
'              November 2020
'
' Basic DOF config
'   90  BackGlass On Image
'   101 Left Flipper, 102 Right Flipper
'   105 Bumper1, 106 Bumper2, 107 Bumper3
'   103 Left Sling, 104 Right Sling
'   120 Kickers
'   125 Ball Release
'   130 - 138 Wire Rollovers
'   141 Chime 1-10s, 142 Chime2-100s, 143 Chime3-1000s
'   160 - 161 Rollover Buttons
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

Const cGameName = "Whirl_Wind_1958"
Const cOptions = "Whirl_Wind_1958.txt"
Const hsFileName = "Whirl Wind (Gottlieb 1958)"

'**********************************************************************************************
'Set this variable to 1 to save a PinballY High Score file to your Tables Folder
'this will let the Pinball Y front end display the high scores when searching for tables
'0 = No PinballY High Scores, 1 = Save PinballY High Scores
Const cPinballY = 2
'**********************************************************************************************

'**********************************************************************************************
'Set BallLift to 0 to enable Automatic Mechanical Ball Lift, 1 will require right magnasave to lift ball
dim BallLift
BallLift = 0
'**********************************************************************************************

'**********************************************************************************************
'Set to 0 to turn off cigs for player up
'Note neither the author of this table nor the surgeon general supports use of tobacco products
dim cigOn
cigOn = 1
'**********************************************************************************************


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
Dim deathLane


Sub Table1_init
  LoadEM
  maxPlayers = 2

  For x = 1 to maxPlayers
    Set sReel(x) = EVAL("scoreReel" & x)
  Next

  player=1
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
  If balls = "" Then balls = 5
  If chime = "" Then chime = 0
  If onBumper = "" Then onBumper = 1
  If deathLane = "" Then deathLane = 0

  replaySettings

  firstBallOut = 0
  updatePostIt
  dynamicUpdatePostIt.enabled = 1
  TiltReel.setValue(1)
  CreditReel.setvalue(credit)
  If deathLane = 1 Then
    layer1cheat.visible = 1
    layer1cheat.collidable = 1
  Else
    layer1cheat.visible = 0
    layer1cheat.collidable = 0
  End If

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

  ballShadowUpdate.enabled = True

  MatchReel.setValue(MatchNumber) 'Need to set to change if 1 point table

  tilt = False
  state = False
  gameState

  If balls = 5 Then
'   instructCard.image = "Wiggler 5 Ball Instruction"
'   coincard.image = "Balls5"
  Else
'   coincard.image = "Balls3"
'   instructCard.image = "Wiggler 3 Ball Instruction"
  End If

End Sub


'***********KeyCodes
Dim enableInitialEntry, firstBallOut
Sub Table1_KeyDown(ByVal keycode)

  If enableInitialEntry = True Then enterIntitals(keycode)

  If keycode = addCreditKey Then
    playFieldSound "coinin",0,Drain,1
    addCredit = 1
    scoreMotor5.enabled = 1
    End If

    If keycode = startGameKey Then
    If enableInitialEntry = False and operatormenu = 0 and backGlassOn = 1 Then
      If freePlay = 1 and players < 2 and firstBallOut = 0 Then startGame
      If freePlay = 0 and credit > 0 and players < 2 and firstBallOut = 0 Then
        credit = credit - 1
        playsound "Reel1", 0, reelVol, 0, 0, 0, 1, 1, 0
        creditReel.setvalue(credit)
        If B2SOn Then
          If freeplay = 0 Then controller.B2SSetCredits credit
          If freePlay = 0 and credit < 1 Then DOF 199, DOFOff
        End If
        startGame
      End If
    End If
  End If

  If keycode = PlungerKey Then
    plunger.PullBack
    playFieldSound "plungerpull", 0, plunger, 1
  End If

  If tilt = False and state = True Then
  If keycode = leftFlipperKey and contball = 0 Then
    LFPress = 1
    lf.fire
    playFieldSound SoundFX("FlipUpL",DOFFlippers), 0, leftFlipper, 1
    DOF 101,DOFOn
    playFieldSound "FlipBuzzL", -1, leftFlipper, 1
  End If

  If keycode = RightFlipperKey and contball = 0 Then
    RFPress = 1
    rf.fire
    playFieldSound SoundFX("FlipUpR",DOFFlippers), 0, RightFlipper,1
    DOF 102,DOFOn
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

    If keycode = leftFlipperKey and state = False and operatorMenu = 0 and enableInitialEntry = 0 Then
        operatorMenuTimer.Enabled = true
    End If

    If keycode = leftFlipperKey and state = False and operatorMenu = 1 Then
    options = options + 1
    If showDt = True Then If options = 2 Then options = 4 'skips non DT options
        If options = 6 Then options = 0
    optionMenu.visible = True
        playFieldSound "target", 0, SoundPointScoreMotor, 1.5
        Select Case (Options)
            Case 0:
                optionMenu.image = "FreeCoin" & freePlay
            Case 1:
                optionMenu.image = balls & "Balls"
      Case 2:
        optionMenu1.visible = 1
        optionMenu1.image = "DOF"
        optionMenu.image = "Chime" & chime
      Case 3:
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
        optionMenu.image = "SaveExit"
        optionMenu2.visible = 0
    End Select
    End If

    If keycode = RightFlipperKey and state = False and operatorMenu = 1 Then
        playFieldSound "metalhit2", 0, SoundPointScoreMotor, 0.2
      Select Case (options)
    Case 0:
            If freePlay = 0 Then
                freePlay = 1
              Else
                freePlay = 0
            End If
            optionMenu.image= "FreeCoin" & freePlay
      If freePlay = 0 Then
        If credit > 0 Then DOF 199, DOFOn
        If credit < 1 Then DOF 199, DOFOff
      Else
        DOF 199, DOFOn
      End If
        Case 1:
            If balls = 3 Then
                balls = 5
              Else
                balls = 3
            End If
'     InstructCard.image = "Wiggler " & balls & " Ball Instruction"
'     Coincard.image = "Balls" & balls
            optionMenu.image = balls & "Balls"
        Case 2:
            If chime = 0 Then
                chime= 1
        DOF 142,DOFPulse
              Else
                chime = 0
        playFieldsound SoundFX("Chime10",DOFChimes), 0, soundPoint13, 1
            End If
      optionMenu.image = "Chime" & chime
    Case 3:
      optionMenu1.visible = 1
      pfOption = pfOption + 1
      If pfOption = 4 Then pfOption = 1
      optionMenu1.image = "Sound" & pfOption

      Select Case (pfOption)
        Case 1: speaker1.visible = 1: speaker2.visible = 1: speaker3.visible = 0: speaker4.visible = 0
        Case 2: speaker5.visible = 1: speaker6.visible = 1: speaker1.visible = 0: speaker2.visible = 0
        Case 3: speaker1.visible = 1: speaker2.visible = 1: speaker3.visible = 1: speaker4.visible = 1: speaker5.visible = 0: speaker6.visible = 0
      End Select
        Case 4:
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
    If B2SOn Then controller.B2SSetTilt 1
    turnOff
  End If

  If Keycode = rightmagnasave Then
    if OKtoReleaseBall = 1 Then releaseBall
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
    If matchNumber > 10 then matchNumber = 1
    If matchNumber > 9 Then matchNumber = 0
    match
    matchNumber = matchNumber + 1

  End If

  If Keycode= 31 Then 's' key
    if rotoOn = 0 Then RotateRototarget
    cig1.visible = 1
  End If

  If Keycode = 33 Then 'f' key
    controller.B2SSetPlayerUp 91, 1
'   x=x+1
'   if x = 11 then x = 0
'   MatchReel.setValue(x)
'   if B2SOn then       controller.B2SSetMatch 34, x
  End If
'************************End Of Test Keys****************************
End Sub


dim TestChr1, TestChr2, TestChr3, TestChr

Sub Table1_KeyUp(ByVal keycode)

  If keycode = plungerKey Then
    plunger.Fire
    playFieldSound "PlungerFire", 0, plunger, 1
  End If

    If keycode = leftFlipperKey Then
        operatorMenuTimer.Enabled = False
    End If

   If tilt = False and state = True Then
    If keycode = leftFlipperKey and contball = 0 Then
      lfpress = 0
      LeftFlipper.eosTorqueAngle = EOSA
      LeftFlipper.eosTorque = EOST
      LeftFlipper.RotateToStart
      playFieldSound SoundFX("FlipDownL",DOFFlippers), 0, leftFlipper, 1
      DOF 101,DOFOff
      stopSound "FlipBuzzLA"
      stopSound "FlipBuzzLB"
      stopSound "FlipBuzzLC"
      stopSound "FlipBuzzLD"
    End If

    If keycode = RightFlipperKey and contball = 0 Then
      rfpress = 0
      RightFlipper.eosTorqueAngle = EOSA
      RightFlipper.eosTorque = EOST
      RightFlipper.rotateToStart
      playFieldSound SoundFX("FlipDownR",DOFFlippers), 0, RightFlipper, 1
      DOF 102,DOFOff
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
    '*****GI Lights On
    dim xx
    For each xx in GI:xx.State = 1: Next
    For each xx in layer1col: xx.image = "layer1": Next
    For each xx in layer2col: xx.image = "layer2": Next
    For each xx in layer3col: xx.image = "layer3": Next
    For each xx in layer4col: xx.image = "layer4": Next
'   For each xx in gatescol: xx.image = "gates": Next
'   bumpa002.image="caps2lit"
'   bumpb002.image="caps2lit"
'   bumpc002.image="caps2lit"
    playfield_off.visible = 0
    RM1.state = 1
    rotoMultiplier = 1
'   FlipInLight.state = 0
    For Each xx in gobblelight: xx.state = 1: Next

    If B2SOn Then
      controller.B2SSetCredits credit
      controller.B2SSetMatch 34, matchNumber
      controller.B2SSetTilt 33,1
      controller.B2SSetData 90, 1 'turns the backglass image on
      controller.B2SSetPlayerUp 30,0
      If credit > 0 Then DOF 199, DOFOn
      If freePlay = 1 Then DOF 199, DOFOn
    End If

    For x = 1 to maxPlayers
      If B2SOn Then controller.B2SSetScorePlayer x, score(x)
      sReel(x).setvalue(score(x))
    Next

    AlternateLights

    backGlassOn = 1
    me.enabled = False
  End If
End Sub

'***********Replay Settings
Sub replaySettings
  If balls = 5 Then
    replay(1) = 3600
    replay(2) = 4800
    replay(3) = 5400
    replay(4) = 6000
  Else
    replay(1) = 2600
    replay(2) = 3800
    replay(3) = 4400
    replay(4) = 5000
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
Dim ballInPlay, bonusScore
Sub startGame
  If state = False Then
    ballInPlay = 1
    If B2SOn Then
      controller.B2SSetCredits credit
      controller.B2SSetBallinPlay 32, ballInPlay
      controller.B2SSetCanPlay 31, 1
      controller.B2SSetPlayerup 91, 1
      controller.B2SSetGameOver 0
    End If
    If cigOn = 1 then cig1.visible = 1
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
    newGame
  Else If state = True and players < maxPlayers and ballinPlay = 1 Then
    players = players + 1
    creditReel.setvalue(credit)
    If B2SOn Then controller.B2SSetCredits credit
    If B2sOn Then controller.B2SSetCanPlay 31, players
    EVAL("CanPlay" & players).setValue(1)
    EVAL("CanPlay" & players - 1).setValue(0)
    End If
  End If
End Sub

'*********New Game
Dim endGame, roundHS
Sub newGame
  player = 1
    endGame = 0
  roundHS = 0
  gameState
  For f = 1 to 3
    EVAL("Bumper" & f).hashitevent = 1
  Next
  If B2SOn Then controller.B2SSetData 25, 0
  If B2SOn Then controller.B2SSetData 26, 0
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
    For x = 1 to 4
      EVAL("CanPlay" & x).setValue(0)
    Next
    state = False
    BIPReel.setValue(0)
    UP1.setValue(0)
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
    cig1.visible = 0
    cig2.visible = 0

    If B2SOn Then
      controller.B2SSetGameOver 35,1
      controller.B2SSetballinplay 32, 0
      controller.B2SSetData (90 + player), 0
      controller.B2SSetCanPlay 31, 0
      If credit > 0 Then DOF 199, DOFOn
      If freePlay = 1 Then DOF 199, DOFOn
    End If
    BIPReel.setValue(0)
  Else
    relBall = 1   'variable to tell score motor routine that this is a ball release run
    scoreMotor5.enabled = 1
    rotoKick = 1
    BIPReel.setValue(ballinPlay)
  End If
End Sub

'***************Drain and Release Ball
Dim drainActive
Sub drain_Hit()
  playFieldSound "Drain", 0, drain, 0.001
  repAwarded(Player) = 0
  drainActive = 1
  drain.DestroyBall 'this is used for multiball tables to avoid timing issues with balls moving in the trough
  BallsInPlay = BallsInPlay - 1
  DrainDelay.enabled = 1
End Sub


Sub DrainDelay_timer
  If BallsInPlay = 0 Then scoreBonus
  DrainDelay.enabled = 0
End Sub

Dim launched, BallsInPlay, OKtoReleaseBall
Sub releaseBall
  drain1.CreateSizedBallWithMass Ballsize/2, BallMass  'used for multiball tables (see above)
  BallsInPlay = BallsInPlay + 1
  playFieldSound SoundFX("FastKickIntoLaunchLane",DOFContactors), 0, drain, 0.5
  DOF 125, DOFPulse
  drain1.kick 270,15    'use only this and no create/destroy ball on single ball tables
  launched = 0
  OKtoReleaseBall = 0
  If B2SOn Then Controller.B2SSetBallinPlay 32, ballinPlay
  drainActive = 0
  BIPReel.setValue(ballInPlay)
End Sub

'**********Check if Scoring Bonus is True
Dim BonusFlag
Sub ScoreBonus
  bonusScore = 0
' BonusFlag = 1
' Flag100 = 0: Flag10 = 0:  Flag1000 = 0
' ScoreMotor.enabled = 1
' If BonusScore = 0 Then DoubleBonus.state = 0
' If BonusScore = 0 And ShootAgain.state = 1 Then ReleaseBall
  If BonusScore = 0 Then AdvancePlayers 'And ShootAgain.state = 0
End Sub

'**********Advance Players
Sub advancePlayers
   If players = 1 or player = players Then
    If B2SOn Then controller.B2SSetData (90 + player), 0
    player = 1
    EVAL("UP" & players).SetValue(0)
    Up1.SetValue(1)
    if cigOn = 1 then cig2.visible = 0: cig1.visible = 1
   Else
    If B2SOn Then controller.B2SSetData (90 + player), 0
    player = player + 1
    EVAL ("Up" & (player - 1)).setValue(0)
    EVAL ("Up" & player).setValue(1)
    if cigOn = 1 then Cig1.visible = 0: cig2.visible = 1
  End If
  If B2SOn Then controller.B2SSetData (90 + player), 1
  nextBall
End Sub

'**********Next Ball
Sub nextBall
    If tilt = True Then
    For f = 1 to 3  'set to number of bumpers
    EVAL("Bumper" & f).hashitevent = 1
    Next
      tilt = False
     TiltReel.setValue(0)
    If B2SOn Then
      controller.B2SSetTilt 33,0
      controller.B2SSetData 1, 1
    End If
    End If
  resetTable
  if Player = 1 then ballinPlay = ballinPlay + 1
' If player = 4 then BallinPlay = 5 'used for match testing

  If ballinPlay > balls then
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
    GameOverReel.setValue(1)
    If B2SOn then controller.B2SSetGameOver 35,1
    If B2SOn then controller.B2SSetData 80,0
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
    MatchReel.setValue(0)
    TiltReel.setValue(0)
    If B2SOn Then
      controller.B2SSetTilt 33,0
      controller.B2SSetMatch 34,0
      controller.B2SSetGameOver 35,0
      controller.B2SSetData 80,1
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
  DOF 165, DOFPulse
End Sub

'******* for ball control script
Sub BallControlOff_Hit()
    contBallInPlay = False
End Sub

Sub phys_subway_hit
  contballinplay = False
End Sub

'************Check if Ball Out of Launch Lane
Dim ballInLane
Sub ballsInPlay_hit
  If ballREnabled=1 Then
    ballREnabled = 0
    ballInLane = False
  End If
  firstBallOut = 1
End Sub

'************** Reset Table
Sub resetTable
' gate1.collidable = 1: gates004.visible = 1: gates003.visible = 0
End Sub

'************** Alternate Lights
Dim Alternate
Sub AlternateLights
  dim xx
  If Alternate = 0 then
'If slingleftlit=1 Then
    For each xx in slinglampsleft:xx.State = 1: Next
    For each xx in slinglampsright:xx.State = 0: Next
    For each xx in layer1acol: xx.image = "layer1": Next
    For each xx in layer1bcol: xx.image = "layer1b": Next
    For each xx in layer2acol: xx.image = "layer2": Next
    For each xx in layer2bcol: xx.image = "layer2b": Next
    For each xx in layer3acol: xx.image = "layer3": Next
    For each xx in layer3bcol: xx.image = "layer3b": Next
    LightRO1.state = 1
    LightRO2.state = 0
    LightRO3.state = 1
    LightRO4.state = 0
    BumpLight3001.state = 1
    BumpLight3002.state = 1
    BumpLight1001.state = 0
    BumpLight1002.state = 0
    BumpLight4001.state = 1
    BumpLight4002.state = 1
    BumpLight5001.state = 0
    BumpLight5002.state = 0
  Else
'If slingrightlit=1 Then
    For each xx in slinglampsleft:xx.State = 0: Next
    For each xx in slinglampsright:xx.State = 1: Next
    For each xx in layer1acol: xx.image = "layer1b": Next
    For each xx in layer1bcol: xx.image = "layer1": Next
    For each xx in layer2acol: xx.image = "layer2b": Next
    For each xx in layer2bcol: xx.image = "layer2": Next
    For each xx in layer3acol: xx.image = "layer3b": Next
    For each xx in layer3bcol: xx.image = "layer3": Next
    LightRO1.state = 0
    LightRO2.state = 1
    LightRO3.state = 0
    LightRO4.state = 1
    BumpLight3001.state = 0
    BumpLight3002.state = 0
    BumpLight1001.state = 1
    BumpLight1002.state = 1
    BumpLight4001.state = 0
    BumpLight4002.state = 0
    BumpLight5001.state = 1
    BumpLight5002.state = 1
  End If
  Alternate = Alternate + 1
  If Alternate = 2 then Alternate = 0
End Sub

'************** Bumpers and Skirt Animation
dim bump1step, bump2step, bump3step
Sub bumpers_Hit(Index)

  If tilt = False Then
    Select Case (Index)
      Case 0: If BumpLight1001.state = 0 Then
            addScore 1
          Else
            addScore 10
          End If
          EVAL ("skirt" & (index + 1)).enabled = 1
          PlayFieldSound SoundFX("PopBump",DOFContactors), 0, bumper1, 1
          DOF 105, DOFPulse
          DOF 110, DOFPulse
      Case 1: addScore 1
          EVAL ("skirt" & (index + 1)).enabled = 1
          playFieldSound SoundFX("PopBump",DOFContactors), 0, bumper2, 1
          DOF 106, DOFPulse
          DOF 111, DOFPulse
      Case 2: If BumpLight3001.state = 0 Then
          EVAL ("skirt" & (index + 1)).enabled = 1
            addScore 1
          Else
            addScore 10
          End If
          playFieldSound SoundFX("PopBump",DOFContactors), 0, bumper3, 1
          DOF 107, DOFPulse
          DOF 112, DOFPulse
      Case 3: addScore 10
      Case 4: addScore 10
      Case 5: addScore 5: if rotoOn = 0 Then RotateRototarget
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

Sub Skirt3_timer
  Select Case bump3Step
    Case 3: bumpskirt3.RotY=-3
    Case 4: bumpskirt3.RotY=2
    Case 5: bumpskirt3.RotY=-1
    Case 6: bumpskirt3.RotY=1
    Case 7: bumpskirt3.RotY=0: bump3Step = 0: Skirt3.enabled = 0
  End Select
  bump3Step = bump3Step + 1
End Sub

Sub bumper1On
  bumpa001.image="caps1lit"
  bumpa002.image="caps2lit"
  bumplight1.state=1
  bumpskirt1.image="skirts"
End Sub

Sub bumper1Off
  bumpa001.image="caps1"
  bumpa002.image="caps2"
  bumplight1.state=0
  bumpskirt1.image="skirts_off"
End Sub

Sub bumper2On
  bumpb001.image="caps1lit"
  bumpb002.image="caps2lit"
  bumplight2.state=1
  bumpskirt2.image="skirts"
End Sub

Sub bumper2Off
  bumpb001.image="caps1"
  bumpb002.image="caps2"
  bumplight2.state=0
  bumpskirt2.image="skirts_off"
End Sub

Sub bumper3On
  bumpc001.image="caps1lit"
  bumpc002.image="caps2lit"
  bumplight3.state=1
  bumpskirt3.image="skirts"
End Sub

Sub bumper3Off
  bumpc001.image="caps1"
  bumpc002.image="caps2"
  bumplight3.state=0
  bumpskirt3.image="skirts_off"
End Sub

'************** 10's Rubbers
Sub TensRubbers_Hit
  addScore 10
End Sub

'************** Top 200 Lights
Sub Bands1_Hit (Index)
  addScore 1
End Sub
'

'************** Wire RollOvers
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

'************** Triggers
dim rotoKick
rotoKick = 1
Sub Triggers_Hit (Index)
  wireNumber = Index
  If wireNumber < 5 Then wireAnimation.enabled = 1
  DOF (130 + Index), DOFPulse
  Select Case Index
    Case 0: If LightRO1.state = 1 Then addScore 10: Else addScore 1
    Case 1: If LightRO2.state = 1 Then addScore 10: Else addScore 1
    Case 2: If LightRO3.state = 1 Then addScore 10: Else addScore 1
    Case 3: If LightRO4.state = 1 Then addScore 10: Else addScore 1
    Case 4: addScore (RotoArray(roto.objroty/24) * 10)
    Case 5: addScore 10
    Case 6: addScore 10
    Case 7: addScore (RotoArray(roto.objroty/24) * 10)
    Case 8: addScore 1
    Case 9: addScore 1
    Case 10: addscore (100 * RotoArray(roto.objroty/24))
    Case 11: If rotoKick > 2 Then rotoKick = 1
          rotoKick = rotoKick + 1
          If rotoKick = 2 Then RotateRototarget
  End Select
End Sub

'************** RollOvers
Dim buttonNumber
Sub rollOver_Hit (Index)
  buttonNumber = (Index)
  rollOverAnimation.enabled = 1
  DOF (160 + Index), DOFPulse
' Select Case (Index)
'   Case 0: addScore (1)
'   Case 1: addScore (1)
' End Select
End Sub

Dim button
Sub rollOverAnimation_Timer
  button = button + 1
  Select Case Button
    Case 1: WiggleGateOpenButton.transz = -2
    Case 2: WiggleGateOpenButton.transz = 0
    Case 3: WiggleGateOpenButton.transz = 1
    Case 4: WiggleGateOpenButton.transz = 1
        button = 0
        rollOverAnimation.enabled = 0
  End Select
End Sub

'************ RotoTarge Rotate
'if gobblelight or gobblelightred are lit then layer1d.image="layer1"
'if gobblelight or gobblelightred are NOT lit (in between roto indexes) thenlayer1d.image="layer1b"
dim rotoRotate, rotateSteps, RotatedSteps, upperbound, lowerbound, backStep
Dim RotoArray(15)
'This is the array of targets on the roto wheel starting with the 4 that is up at the start of each new session
RotoArray(0) = 4: RotoArray(1) = 1: RotoArray(2) = 2: RotoArray(3) = 3: RotoArray(4) = 1: RotoArray(5) = 4: RotoArray(6) = 2: RotoArray(7) = 1
RotoArray(8) = 3: RotoArray(9) = 2: RotoArray(10) = 1: RotoArray(11) = 3: RotoArray(12) = 1: RotoArray(13) = 2: RotoArray(14) = 5
upperbound = 50: lowerbound = 22

Sub RotoTarget_hit
  addScore (RotoArray(roto.objroty/24) * rotoMultiplier)
  rotoBend.enabled = 1
End Sub

dim RotoAngle
Sub RotateRototarget
  rotateSteps = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
  rotatedSteps = rotateSteps
  rotoTimer.enabled = 1
  playFieldSound "rotostart", 0, layer2plate, 1
End Sub

dim rotoOn
Sub RotoTimer_Timer
  rotoOn = 1
  dim xx, yy
  RotoAngle = RotoAngle + 4 ' goes 4 degrees for each step
  if RotoAngle > 359 then RotoAngle = 0
  roto.objroty = RotoAngle
  If roto.objroty Mod 6 > 0 Then
    For each yy in gobblelight: yy.state = 0: Next
    For each yy in gobblelightred: yy.state = 0: Next
    layer1d.image="layer1b"
    layer2d.image="layer2b"
    layer2railsd.image="layer2b"
  Else
    If (roto.objroty/24) = 14 Then
      For Each yy in gobblelightred: yy.state = 1: Next
      layer1d.image="layer1"
      layer2d.image="layer2"
      layer2railsd.image="layer2"
    Else
      For Each yy in gobblelight: yy.state = 1: Next
      layer1d.image="layer1"
      layer2d.image="layer2"
      layer2railsd.image="layer2"
    End If
  End If

  For each xx in rototargets: xx.objroty = RotoAngle: Next
  rotatedSteps = rotatedSteps - 1
  If rotatedSteps = 0 Then
    backStep = rotateSteps Mod 6
    if backStep > 0 Then backUp.enabled = 1
    rotoOn = 0
    rotoTimer.enabled = 0
  End If
End Sub

Sub BackUp_Timer
  dim xx, yy
  If RotoAngle > 0 Then
    RotoAngle = RotoAngle - 4
    roto.objroty = RotoAngle
  End If

  If (roto.objroty/24) = 14 Then
    For Each yy in gobblelightred: yy.state = 1: Next
    layer1d.image="layer1"
    layer2d.image="layer2"
    layer2railsd.image="layer2"
  Else
    For Each yy in gobblelight: yy.state = 1: Next
    layer1d.image="layer1"
    layer2d.image="layer2"
    layer2railsd.image="layer2"
  End If
  layer1d.image="layer1"
  layer2d.image="layer2"
  layer2railsd.image="layer2"
  For each xx in rototargets: xx.objroty = RotoAngle: Next
  backStep = backStep - 1
  tb.text = RotoArray(roto.objroty/24)
' tb1.text = roto.objroty/24
  If backStep = 0 Then
    If roto.objroty Mod 24 = 0 Then
      BackUp.enabled = 0
      rotoOn = 0
    Else
      backStep = (roto.objroty Mod 24/4)
    End If
  End If
End Sub

dim rotoMult, rotoMultiplier
Sub rotoMultStep
  rotoMult = rotoMult + 1
  if rotoMult > 2 Then rotoMult = 0
  Select Case rotoMult
    Case 0: rotoMultiplier = 1: RM100.state = 0: RM1.state = 1
    Case 1: rotoMultiplier = 10: RM1.state = 0: RM10.state = 1
    Case 2: rotoMultiplier = 100: RM10.state = 0: RM100.state = 1
  End Select
End Sub

dim bendBack
Sub rotoBend_timer
  dim xx
  bendBack = bendBack + 1
  Select Case bendBack
    Case 1: For each xx in rototargets: xx.transy= -5:  Next
    Case 2: For each xx in rototargets: xx.transy= 2: Next
    Case 3: For each xx in rototargets: xx.transy= -1:  Next
    Case 4: For each xx in rototargets: xx.transy= 0: Next
  End Select
  If bendBack = 4 Then bendBack = 0: rotoBend.enabled = 0
End Sub

'************ Kickers
Sub kickBall(kvel, kvelz, kzlift)
  dim rangle, BOT, b, saucerScore
  BOT = GetBalls
    For b = 0 to uBound(BOT)
      If BOT(b).x < 740 and BOT(b).x > 734 and BOT(b).y > 430 and BOT(b).y < 440 and kickerUp = 1 Then
        rangle = 3.14 * (155) / 180
        BOT(b).z = BOT(b).z + kzlift
        BOT(b).velz = kvelz
        BOT(b).velx = cos(rangle)*kvel
        BOT(b).vely = sin(rangle)*kvel
      End If
      If BOT(b).x > 765 and BOT(b).x < 778 and BOT(b).y > 718 and BOT(b).y < 730 and kickerUp = 2 Then
        rangle = 3.14 * (155) / 180
        BOT(b).z = BOT(b).z + kzlift
        BOT(b).velz = kvelz
        BOT(b).velx = cos(rangle)*kvel
        BOT(b).vely = sin(rangle)*kvel
      End If
    Next
End Sub

dim kickerActive
Sub kicker_hit
    kickerActive = 1
    addScore 5
End Sub

dim kickStep1, kickStep2

''**********Sling Shot Animations
'' Rstep and Lstep  are the variables that increment the animation
''****************
Dim lStep, rStep
Sub rightSlingShot_Slingshot
  playfieldSound SoundFX("SlingShot",DOFContactors), 0, SoundPoint13, 1
  DOF 104, DOFPulse
  DOF 109, DOFPulse
    addScore(10)
    rSling.Visible = 0
    rSling2.Visible = 1
    sling1.Rotx = 10
    rStep = 0
    rightSlingShot.TimerEnabled = 1
End Sub

Sub rightSlingShot_Timer
    Select Case rStep
        Case 3:rSLing2.Visible = 0:rSLing3.Visible = 1:sling1.Rotx = 0
        Case 4:rSLing3.Visible = 0:rSLing.Visible = 1:rightSlingShot.TimerEnabled = 0
    End Select
    rStep = rStep + 1
End Sub
'
Sub leftSlingShot_Slingshot
  playfieldSound SoundFX("SlingShot",DOFContactors), 0, SoundPoint12, 1
  DOF 103, DOFPulse
  DOF 108, DOFPulse
  addScore(10)
    lSling.Visible = 0
    lSling2.Visible = 1
    sling2.Rotx = 10
    lStep = 0
    leftSlingShot.TimerEnabled = 1
End Sub

Sub leftSlingShot_Timer
    Select Case lStep
        Case 3:lSLing2.Visible = 0:lSLing3.Visible = 1:sling2.Rotx = 0
        Case 4:lSLing3.Visible = 0:lSLing.Visible = 1:leftSlingShot.TimerEnabled = 0
    End Select
    lStep = lStep + 1
End Sub

'***************Score Motor Run one full rotation
Dim scoreMotorCount, addCredit
Sub scoremotor5_Timer
  scoreMotorCount = scoreMotorCount + 1
  playFieldSound "ScoreMotorSingleFire", 0, SoundPointScoreMotor, 0. 'need to set location of score motor under the PF
  If scoremotorCount = 5 Then
    If relBall = 1 Then
      relBall = 0
      OKtoReleaseBall = 1
      If BallLift = 0 Then releaseBall
    End If
    If addCredit = 1 Then
      credit = credit + 1
      playsound "Reel1", 0, reelVol, 0, 0, 0, 1, 1, 0
      DOF 199, DOFOn
      If credit > 15 then credit = 15
      CreditReel.setValue (credit)
      If B2SOn Then
        controller.B2SSetCredits credit
        If credit > 0 Then DOF 199, DOFOn
      End If
      addCredit = 0
    End If
    scoreMotorCount = 0
    scoreMotor5.enabled = 0
  End If
End Sub

'****************Score Motor Run Timer
Dim bellRing, scoreMotorLoop, kickerUp
Sub scoreMotor_timer
  scoreMotorLoop = scoreMotorLoop + 1
  If scoreMotorLoop < 6 Then playFieldSound "ScoreMotorSingleFire", 0, SoundPointScoreMotor, 0.2

'  These Flags are passed by scores with multiple of 10, 100 or 1000
  If flag1 = 1 or flag10 = 1 or flag100 = 1 or flag1000 = 1 Then
    Select Case scoreMotorLoop
      Case 1: totalUp Point:' If point = 1000 Then advanceBonus 'this is for tables with a 1000 point bonus tree
      Case 2: If bellRing > 4 Then totalUp point
      Case 3: totalUp Point: 'If point = 1000 Then advanceBonus
      Case 4: If bellRing > 3 Then totalUp point
      Case 5: If bellRing > 2 Then totalUp point: 'If point = 1000 Then advanceBonus
      Case 6: scoreMotorLoop = 0
          rotoMultStep
          flag1 = 0
          flag10 = 0
          flag100 = 0
          flag1000 = 0
          If kickerActive = 1 Then
            kicker.kick 0,30
            playfieldSound SoundFX("SlingShot",DOFContactors), 0, SoundPoint13, 1
            DOF 120, DOFPulse
            kickerActive = 0
          End If
          scoreMotor.enabled = 0 'originally had If drainActive = 0 Then scoreMotor.Enabeled = 0 but this casued a final ball drain error in PG
    End Select
  End If
End Sub

'***************Scoring Routine
Dim flag1, flag10, flag100, flag1000, point, point10
Sub addScore(points)
  reelDone(player) = 0
  If tilt = True and kickerActive = 1 Then
    kicker.kick 0,30
    playfieldSound SoundFX("SlingShot",DOFContactors), 0, SoundPoint13, 1
    DOF 120, DOFPulse
    kickerActive = 0
  End If
  If tilt = False Then

    If points < 9 Then
'     Number Matching, decrement the match unit for each 1 point score roll over from 10 to 1
      matchNumber = matchNumber - 1
      If matchNumber < 1 Then matchNumber = 10
      bellRing = (points)
      If bellRing > 1 Then point = 1: flag1 = 1: scoreMotor.enabled = 1
      If bellRing = 1 Then
        alternateLights
        totalUp(1)
'       If chime = 0 Then
'         playFieldSound "Chime10", 0, soundPoint13, 1
'       Else
'         If B2SOn Then DOF 141,DOFPulse
'       End If
      End If
      Exit Sub
    End If

    If points > 9 and points <100 Then
'     Number Matching, decrement the match unit for each 10 point score roll up from 0 to 9
      point10 =  (points / 10)
      If matchNumber >= point10 Then
        matchNumber = matchNumber - point10
      Else
        point10 = point10 - matchNumber
        matchNumber = 10 - point10
      End If

      bellRing = (points / 10)
      If bellRing > 1 Then point = 10: flag10 = 1: scoreMotor.enabled = 1
      If bellRing = 1 Then
        totalUp(10)
        If chime = 0 Then
          playFieldSound SoundFX("Chime10",DOFChimes), 0, soundPoint13, 1
        Else
          DOF 142,DOFPulse
        End If
      End If
      Exit Sub
    End If

    If points > 99 and points < 1000 Then
      bellRing = (points / 100)
      If bellRing > 1 Then point = 100: flag100 = 1: scoreMotor.enabled = 1
      If bellRing = 1 Then
        totalUp(100)
        If chime = 0 Then
          playFieldSound SoundFX("Chime100",DOFChimes), 0, soundPoint13, 1
        Else
          DOF 143,DOFPulse
        End If
      End If
      Exit Sub
    End If

    If Points > 999 Then
      bellRing = (points / 1000)
      If bellRing > 1 Then point = 1000: flag1000 = 1: scoreMotor.enabled = 1
      If bellRing = 1 Then
        totalUp(1000)
        If chime = 0 Then
          playFieldSound SoundFX("Chime100",DOFChimes), 0, soundPoint13, 1
        Else
          DOF 143,DOFPulse
        End If
      End If
    End If

  End If
End Sub

Dim replayX,  replay(7), repAwarded(5), reelVol, reelLpan, reelRpan
reelVol = 0.3   'Volume value for reels.  0.1 a good value
reelLpan = -0.12  'Pan value for left reels.  -0.12 is 3/4 pan left
reelRpan = 0.12 'Pan value for right reels.  0.12 is 3/4 pan right

Sub totalUp(points)
  If B2SOn and showDT = False Then  'Modify "If Player" logic below if locations of reels are different
    If Player = 1 Then
    playsound "Reel1", 0, reelVol, reelLpan, 0, 0, 1, 1, 0
    Else
    playsound "Reel1", 0, reelVol, reelRpan, 0, 0, 1, 1, 0
    End If
  End If

  If showDT = True Then
    If Player = 1 Then
    playsound "Reel1", 0, reelVol, reelRpan, 0, 0, 1, 1, 0
    Else
    playsound "Reel1", 0, reelVol, reelRpan, 0, 0, 1, 1, 0
    End If
  End If

  If flag1 = 1 Then AlternateLights

  If flag10 = 1 Then
    If chime = 0 Then
      playFieldSound SoundFX("Chime10",DOFChimes), 0, soundPoint13, 1
    Else
      DOF 142,DOFPulse
    End If
  End If

  If flag100 = 1 Then
    If chime = 0 Then
      playFieldSound SoundFX("Chime100",DOFChimes), 0, soundPoint13, 1
    Else
      DOF 143,DOFPulse
    End If
  End If

  If flag1000 = 1 Then
    If chime = 0 Then
      playFieldSound SoundFX("Chime100",DOFChimes), 0, soundPoint13, 1
    Else
      DOF 143,DOFPulse
    End If
  End If

' This is for a table with a bonus tree
' If BonusFlag = 1 Then
'   If Chime = 0 Then
'     PlayFieldSound  "Chime1000", 0, SoundPoint13, 1
'   Else
'     If B2SOn Then DOF 143,DOFPulse
'   End If
' End If

  score(player) = score(player) + points
  sReel(player).addvalue(points)

'This is for score roll over lights on a table (ie 1000 Light)
  If Score(Player) > 999 Then
    If B2SOn Then controller.B2SSetData Player + 24, 1
  End If

  If B2SOn Then controller.B2SSetScorePlayer player, score(player)

  For replayX = rep(player) + 1 to 4
    If score(player) => replay(replayX) Then
      addCredit = 1
      scoreMotor5.enabled = 1
      rep(player) = rep(Player) + 1
      playsound soundFXDOF("Knocker",128,DOFPulse,DOFKnocker)
      DOF 129, DOFPulse
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

'***************Match
Sub match
  tb.text = matchNumber
' matchNumber = ((int(rnd(1)*10)) * 10) + 10
  If matchNumber = 0 Then matchNumber = 10
  MatchReel.setValue(matchNumber) 'need to change if 1 point table
  If B2SOn Then controller.B2SSetMatch 34, (matchNumber * 10)

  For i = 1 to players
    If (matchNumber * 1) = (score(i) mod 10) Then 'need to set this to match 10's or 100's
      addCredit = 1
      scoreMotor5.enabled = 1
      playsound SoundFXDOF("Knocker",128,DOFPulse,DOFKnocker)
      DOF 129, DOFPulse
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

Dim rScore(4,5), resetLoop, test, playerTest, resetFlag, reelFlag, reelStop, reelDone(4)
Sub countUp
  For playerTest = 1 to 4
    For test = 0 to 4
      rScore(playerTest,test) = Int(score(playerTest)/10^test) mod 10
    Next
  Next

  For playerTest = 1 to 4
    For x = 0 to 4
      If rscore(playerTest, x) > 0 And rscore(playerTest, x) < 9 Then score(playerTest) = score(playerTest) + 10^x
      If rScore(playerTest, x) = 9 Then score(playerTest) = score(playerTest) - (9 * 10^x)
      If rScore(playerTest, x) > 0 Then rScore(playerTest, x) = rScore(playerTest, x) + 1
      If rScore(playerTest, x) = 10 Then rScore(playerTest, x) = 0
    Next
  Next

  If score(1) = 0 and score(2) = 0 and score(3) = 0 and score(4) = 0 Then
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
  For playerTest = 1 to 2
    If B2SOn and showDT = False and reelDone(playerTest) = 0 and reelStop = 0 Then
      controller.B2SSetScorePlayer playerTest, score(playerTest)
      If reelDone(1) = 0 and playerTest = 1 Then playsound "Reel1", 0, reelVol, reelLpan, 0, 0, 1, 1, 0
      If reelDone(2) = 0 and playerTest = 2 Then playsound "Reel2", 0, reelVol, reelRpan, 0, 0, 1, 1, 0
      If score(playerTest) = 0 Then reelDone(playerTest) = 1
    End If

    If showDT = True and reelDone(playerTest) = 0 and reelStop = 0 Then
      sReel(playerTest).setvalue (score(playerTest))
      If reelDone(1) = 0 and playerTest = 1 Then playsound "Reel1", 0, reelVol, reelRpan, 0, 0, 1, 1, 0
      If reelDone(2) = 0 and playerTest = 2 Then playsound "Reel2", 0, reelVol, reelRpan, 0, 0, 1, 1, 0
      If score(playerTest) = 0 Then reelDone(playerTest) = 1
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
    If testFlag = 0 Then
      OKtoReleaseBall = 1
      If BallLift = 0 Then releaseBall
    End If
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
          If testFlag = 0 Then
            OKtoReleaseBall = 1
            If BallLift = 0 Then releaseBall
          End If
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
      If testFlag = 0 Then
        OKtoReleaseBall = 1
        If BallLift = 0 Then releaseBall
      End If
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
  For hs = 1 to 4                   'look at all 5 saved high scores
    For chY = 0 to 4                  'look at 4 player scores
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
      playsound SoundFXDOF("Knocker",128,DOFPulse,DOFKnocker)
      DOF 129, DOFPulse
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
      playSound "metalhit_thin"
    End If
    If keycode = RightFlipperKey Then
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
  playsound SoundFXDOF("Chime10",141,DOFPulse,DOFChimes)
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
    credit = cdbl (temp(26))
    freePlay = cdbl (temp(27))
    balls = cdbl (temp(28))
    matchNumber = cdbl (temp(29))
    chime = cdbl (temp(30))
    pfOption = cdbl (temp(31))
    deathLane = cdbl (temp(32))
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
    scoreFile.WriteLine credit
    scorefile.writeline freePlay
    scoreFile.WriteLine balls
    scoreFile.WriteLine matchNumber
    scoreFile.WriteLine chime
    scoreFile.WriteLine pfOption
    scoreFile.WriteLine deathLane
    scoreFile.Close
  Set scoreFile = Nothing
  Set fileObj = Nothing

'This section of code writes a file in the User Folder of VisualPinball that contains the High Score data for PinballY.
'PinballY can read this data and display the high scores on the DMD during game selection mode in PinballY.

  Set FileObj = CreateObject("Scripting.FileSystemObject")

  If cPinballY = 0 Then Exit Sub

  If Not FileObj.FolderExists(UserDirectory) Then
    Exit Sub
  End If

  FolderPath = FileObj.GetParentFolderName(UserDirectory)

  If cPinballY = 1 Then
    Set ScoreFile = FileObj.CreateTextFile(FolderPath & "/Tables/" & hsFileName & ".PinballYHighScores",True)
  Else
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
  DOF 101, DOFOff
  DOF 111, DOFOff
  RightFlipper.RotateToStart
  DOF 102, DOFOff
  DOF 112, DOFOff
' BonusScore = 0
End Sub

'*****************************************************Supporting Code Written By Others*************************************

'*****************************************
'     BALL SHADOW
'*****************************************
Dim BallShadow
BallShadow = Array (BallShadow001,BallShadow002,BallShadow003,BallShadow004,BallShadow005)

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
  Dim tmp, PI
    If PFOption=1 Then tmp = TableObj.x * 2 / table1.width-1
  If PFOption=2 Then tmp = TableObj.y * 2 / table1.height-1
  PI = 4 * ATN(1)
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
'      JP's VP10 Rolling Sounds - Modified
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

Const tnob = 5 ' total number of balls

'Change GHT, GHB and PFL values based upon the real pinball table dimensions.  Values are used by the GlassHit code.
Const GHT = 2 'Glass height in inches at top of real playfield
Const GHB = 2 'Glass height in inches at bottom of real playfield
Const PFL = 40  'Real playfield lenght in inches

ReDim rolling(tnob)
InitRolling

ReDim ArchRolling(tnob)
InitArchRolling

Dim ArchHit
Sub TopArch_Hit
  topArch.hasHitEvent = False
  ArchHit = 1
  ArchTimer.Enabled = 1
End Sub

Dim archCount
Sub ArchTimer_Timer
  archCount = archCount + 1
  If archCount = 2 Then
    archCount = 0
    toparch.hasHitEvent = True
    ArchTimer.enabled = 0
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
  Dim BOT, b, pa
  BOT = GetBalls
  pa=35000  'Playsound pitch adder for subway rolling ball sound

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

' Ball Rolling sounds
'**********************
  If PFOption = 1 or PFOption = 2 Then
    If BallVel(BOT(b) ) > 1 AND BOT(b).z < 26 and BOT(b).z > 0 Then
      rolling(b) = True
      PlaySound("BallrollingA" & b), -1, Vol(BOT(b)) * 0.2 * xGain(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, 0 'Left & Right stereo or Top & Bottom stereo PF Speakers.
    ElseIf BallVel(BOT(b) ) > 1 AND BOT(b).z =< 0 Then
      PlaySound("BallrollingA" & b), -1, Vol(BOT(b)) * 0.2 * xGain(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b))+pa, 1, 0, 0  'Left & Right stereo or Top & Bottom stereo PF Speakers.
    ElseIf rolling(b) = True Then
      StopSound("BallrollingA" & b)
      rolling(b) = False
    End If
  End If

  If PFOption = 3 Then
    If BallVel(BOT(b) ) > 1 AND BOT(b).z < 26 and BOT(b).z > 0 Then
      rolling(b) = True
      PlaySound("BallrollingA" & b), -1, Vol(BOT(b)) * 0.2 *    XVol(BOT(b))  *     YVol(BOT(b)), -1, 0, Pitch(BOT(b)), 1, 0, -1  'Top Left PF Speaker
      PlaySound("BallrollingB" & b), -1, Vol(BOT(b)) * 0.2 * (1-XVol(BOT(b))) *     YVol(BOT(b)),  1, 0, Pitch(BOT(b)), 1, 0, -1  'Top Right PF Speaker
      PlaySound("BallrollingC" & b), -1, Vol(BOT(b)) * 0.2 *    XVol(BOT(b))  * (1-YVol(BOT(b))), -1, 0, Pitch(BOT(b)), 1, 0,  1  'Bottom Left PF Speaker
      PlaySound("BallrollingD" & b), -1, Vol(BOT(b)) * 0.2 * (1-XVol(BOT(b))) * (1-YVol(BOT(b))),  1, 0, Pitch(BOT(b)), 1, 0,  1  'Bottom Right PF Speaker
    ElseIf BallVel(BOT(b) ) > 1 AND BOT(b).z =< 0 Then
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

' Arch Hit and Arch Rolling sounds
'***********************************
  If PFOption = 1 or PFOption = 2 Then
    If BallVel(BOT(b) ) > 1 AND ArchHit =1 Then
      ArchRolling(b) = True
      PlaySound("ArchHit" & b),   0, (BallVel(BOT(b))/20)^5 * 1 * xGain(BOT(b)), AudioPan(BOT(b)), 0, (BallVel(BOT(b))/40)^5, 1, 0, 0 'Left & Right stereo or Top & Bottom stereo PF Speakers.
      PlaySound("ArchRollA" & b), -1, (BallVel(BOT(b))/40)^5 * 1 * xGain(BOT(b)), AudioPan(BOT(b)), 0, (BallVel(BOT(b))/40)^5, 1, 0, 0  'Left & Right stereo or Top & Bottom stereo PF Speakers.
    Else
      If ArchRolling(b) = True Then
      StopSound("ArchRollA" & b)
      ArchRolling(b) = False
      End If
    End If
  End If

  If PFOption = 3 Then
    If BallVel(BOT(b) ) > 1 AND ArchHit =1 Then
      ArchRolling(b) = True
      PlaySound("ArchHit" & b),   0, (BallVel(BOT(b))/20)^5 * 1 *    XVol(BOT(b))  *     YVol(BOT(b)),  -1, 0, (BallVel(BOT(b))/40)^5, 1, 0, -1 'Top Left PF Speaker
      PlaySound("ArchHit" & b),   0, (BallVel(BOT(b))/20)^5 * 1 * (1-XVol(BOT(b))) *     YVol(BOT(b)),   1, 0, (BallVel(BOT(b))/40)^5, 1, 0, -1 'Top Right PF Speaker
      PlaySound("ArchHit" & b),   0, (BallVel(BOT(b))/20)^5 * 1 *    XVol(BOT(b))  *  (1-YVol(BOT(b))), -1, 0, (BallVel(BOT(b))/40)^5, 1, 0,  1 'Bottom Left PF Speaker
      PlaySound("ArchHit" & b),   0, (BallVel(BOT(b))/20)^5 * 1 * (1-XVol(BOT(b))) *  (1-YVol(BOT(b))),  1, 0, (BallVel(BOT(b))/40)^5, 1, 0,  1 'Bottom Right PF Speaker
      PlaySound("ArchRollA" & b), -1, (BallVel(BOT(b))/40)^5 * 1 *    XVol(BOT(b))  *     YVol(BOT(b)),  -1, 0, (BallVel(BOT(b))/40)^5, 1, 0, -1  'Top Left PF Speaker
      PlaySound("ArchRollB" & b), -1, (BallVel(BOT(b))/40)^5 * 1 * (1-XVol(BOT(b))) *     YVol(BOT(b)),   1, 0, (BallVel(BOT(b))/40)^5, 1, 0, -1  'Top Right PF Speaker
      PlaySound("ArchRollC" & b), -1, (BallVel(BOT(b))/40)^5 * 1 *    XVol(BOT(b))  *  (1-YVol(BOT(b))), -1, 0, (BallVel(BOT(b))/40)^5, 1, 0,  1  'Bottom Left PF Speaker
      PlaySound("ArchRollD" & b), -1, (BallVel(BOT(b))/40)^5 * 1 * (1-XVol(BOT(b))) *  (1-YVol(BOT(b))),  1, 0, (BallVel(BOT(b))/40)^5, 1, 0,  1  'Bottom Right PF Speaker
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
'Four intensities of ball bounce sound files ranging from 1 to 4 bounces.  The number of bounces increase as the ball's downward Z velocity increases.
'A BOT(b).VelZ of zero will play the sound of trapped balls (balls trapped by table objects constantly micro-bounce against them).

  If PFOption = 1 or PFOption = 2 Then
    If BOT(b).VelZ >= -3 And BOT(b).VelZ < -0.5 And BOT(b).Z < 55 And BOT(b).Z > 27 And BallinPlay => 1 Then
      PlaySound "BallDrop1" & b, 0, ABS(BOT(b).VelZ)/400 * xGain(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, 0 'Left & Right stereo or Top & Bottom stereo PF Speakers.
    ElseIf BOT(b).VelZ >= -4 And BOT(b).VelZ < -3 And BOT(b).Z < 55 And BOT(b).Z > 27 And BallinPlay => 1 Then
      PlaySound "BallDrop2" & b, 0, ABS(BOT(b).VelZ)/400 * xGain(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, 0 'Left & Right stereo or Top & Bottom stereo PF Speakers.
    ElseIf BOT(b).VelZ >= -5 And BOT(b).VelZ < -4 And BOT(b).Z < 55 And BOT(b).Z > 27 And BallinPlay => 1 Then
      PlaySound "BallDrop3" & b, 0, ABS(BOT(b).VelZ)/400 * xGain(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, 0 'Left & Right stereo or Top & Bottom stereo PF Speakers.
    ElseIf BOT(b).VelZ < -5 And BOT(b).Z < 55 And BOT(b).Z > 27 And BallinPlay => 1 Then
      PlaySound "BallDrop4" & b, 0, ABS(BOT(b).VelZ)/400 * xGain(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, 0 'Left & Right stereo or Top & Bottom stereo PF Speakers.
    End If
  End If

  If PFOption = 3 Then
    If BOT(b).VelZ >= -3 And BOT(b).VelZ < -0.5 And BOT(b).Z < 55 And BOT(b).Z > 27 And BallinPlay => 1 Then
      PlaySound "BallDrop1" & b, 0, ABS(BOT(b).VelZ)/400 *    XVol(BOT(b))  *     YVol(BOT(b)), -1, 0, Pitch(BOT(b)), 1, 0, -1  'Top Left PF Speaker
      PlaySound "BallDrop1" & b, 0, ABS(BOT(b).VelZ)/400 * (1-XVol(BOT(b))) *     YVol(BOT(b)),  1, 0, Pitch(BOT(b)), 1, 0, -1  'Top Right PF Speaker
      PlaySound "BallDrop1" & b, 0, ABS(BOT(b).VelZ)/400 *    XVol(BOT(b))  * (1-YVol(BOT(b))), -1, 0, Pitch(BOT(b)), 1, 0,  1  'Bottom Left PF Speaker
      PlaySound "BallDrop1" & b, 0, ABS(BOT(b).VelZ)/400 * (1-XVol(BOT(b))) * (1-YVol(BOT(b))),  1, 0, Pitch(BOT(b)), 1, 0,  1  'Bottom Right PF Speaker
    ElseIf BOT(b).VelZ >= -4 And BOT(b).VelZ < -3 And BOT(b).Z < 55 And BOT(b).Z > 27 And BallinPlay => 1 Then
      PlaySound "BallDrop2" & b, 0, ABS(BOT(b).VelZ)/400 *    XVol(BOT(b))  *     YVol(BOT(b)), -1, 0, Pitch(BOT(b)), 1, 0, -1  'Top Left PF Speaker
      PlaySound "BallDrop2" & b, 0, ABS(BOT(b).VelZ)/400 * (1-XVol(BOT(b))) *     YVol(BOT(b)),  1, 0, Pitch(BOT(b)), 1, 0, -1  'Top Right PF Speaker
      PlaySound "BallDrop2" & b, 0, ABS(BOT(b).VelZ)/400 *    XVol(BOT(b))  * (1-YVol(BOT(b))), -1, 0, Pitch(BOT(b)), 1, 0,  1  'Bottom Left PF Speaker
      PlaySound "BallDrop2" & b, 0, ABS(BOT(b).VelZ)/400 * (1-XVol(BOT(b))) * (1-YVol(BOT(b))),  1, 0, Pitch(BOT(b)), 1, 0,  1  'Bottom Right PF Speaker
    ElseIf BOT(b).VelZ >= -5 And BOT(b).VelZ < -4 And BOT(b).Z < 55 And BOT(b).Z > 27 And BallinPlay => 1 Then
      PlaySound "BallDrop3" & b, 0, ABS(BOT(b).VelZ)/400 *    XVol(BOT(b))  *     YVol(BOT(b)), -1, 0, Pitch(BOT(b)), 1, 0, -1  'Top Left PF Speaker
      PlaySound "BallDrop3" & b, 0, ABS(BOT(b).VelZ)/400 * (1-XVol(BOT(b))) *     YVol(BOT(b)),  1, 0, Pitch(BOT(b)), 1, 0, -1  'Top Right PF Speaker
      PlaySound "BallDrop3" & b, 0, ABS(BOT(b).VelZ)/400 *    XVol(BOT(b))  * (1-YVol(BOT(b))), -1, 0, Pitch(BOT(b)), 1, 0,  1  'Bottom Left PF Speaker
      PlaySound "BallDrop3" & b, 0, ABS(BOT(b).VelZ)/400 * (1-XVol(BOT(b))) * (1-YVol(BOT(b))),  1, 0, Pitch(BOT(b)), 1, 0,  1  'Bottom Right PF Speaker
    ElseIf BOT(b).VelZ < -5 And BOT(b).Z < 55 And BOT(b).Z > 27 And BallinPlay => 1 Then
      PlaySound "BallDrop4" & b, 0, ABS(BOT(b).VelZ)/400 *    XVol(BOT(b))  *     YVol(BOT(b)), -1, 0, Pitch(BOT(b)), 1, 0, -1  'Top Left PF Speaker
      PlaySound "BallDrop4" & b, 0, ABS(BOT(b).VelZ)/400 * (1-XVol(BOT(b))) *     YVol(BOT(b)),  1, 0, Pitch(BOT(b)), 1, 0, -1  'Top Right PF Speaker
      PlaySound "BallDrop4" & b, 0, ABS(BOT(b).VelZ)/400 *    XVol(BOT(b))  * (1-YVol(BOT(b))), -1, 0, Pitch(BOT(b)), 1, 0,  1  'Bottom Left PF Speaker
      PlaySound "BallDrop4" & b, 0, ABS(BOT(b).VelZ)/400 * (1-XVol(BOT(b))) * (1-YVol(BOT(b))),  1, 0, Pitch(BOT(b)), 1, 0,  1  'Bottom Right PF Speaker
    End If
  End If

' Glass hit sounds
'*******************
' Ball=50 units=1.0625".  Ball.z is ball center.  Balls are physically limited by Top Glass Height.  Max ball.z is 25 units below Top Glass Height.
' To ensure ball can go high enough to trigger glass hit, make Table Options/Dimensions & Slope/Top Glass Height equal to (GHT*50/1.0625) + 5

  If PFOption = 1 or PFOption = 2 Then
    If BOT(b).Z > (BOT(b).Y * ((GHT-GHB)/PFL)) + (GHB*50/1.0625) - BallSize/2 And BallinPlay => 1 Then
      PlaySound "GlassHit" & b, 0, ABS(BOT(b).VelZ)/30 * xGain(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, 0 'Left & Right stereo or Top & Bottom stereo PF Speakers.
    End If
  End If

  If PFOption = 3 Then
    If BOT(b).Z > (BOT(b).Y * ((GHT-GHB)/PFL)) + (GHB*50/1.0625) - Ballsize/2 And BallinPlay => 1 Then
      PlaySound "GlassHit" & b, 0, ABS(BOT(b).VelZ)/30 *    XVol(BOT(b))  *     YVol(BOT(b)), -1, 0, Pitch(BOT(b)), 1, 0, -1  'Top Left PF Speaker
      PlaySound "GlassHit" & b, 0, ABS(BOT(b).VelZ)/30 * (1-XVol(BOT(b))) *     YVol(BOT(b)),  1, 0, Pitch(BOT(b)), 1, 0, -1  'Top Right PF Speaker
      PlaySound "GlassHit" & b, 0, ABS(BOT(b).VelZ)/30 *    XVol(BOT(b))  * (1-YVol(BOT(b))), -1, 0, Pitch(BOT(b)), 1, 0,  1  'Bottom Left PF Speaker
      PlaySound "GlassHit" & b, 0, ABS(BOT(b).VelZ)/30 * (1-XVol(BOT(b))) * (1-YVol(BOT(b))),  1, 0, Pitch(BOT(b)), 1, 0,  1  'Bottom Right PF Speaker
    End If
  End If
  Next
End Sub

'*************Hit Sound Routines
'Eliminated the Hit Subs extra velocity criteria since the PlayFieldSoundAB command already incorporates the balls velocity.

Sub aRubberPins_Hit(idx)
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
  If BallinPlay > 0 Then  'Eliminates the thump of Trough Ball Creation balls hitting walls 9 and 14 during table initiation
  PlayFieldSoundAB "fx_rubber2", 0, 0.1
  End If
End Sub

Sub aRubberWheel_hit
  PlayFieldSoundAB "fx_rubber2", 0, 0.5
End sub

Sub aPosts_Hit(idx)
  PlayFieldSoundAB "fx_rubber2", 0, 1
End Sub


' Thalamus : This sub is used twice - this means ... this one IS NOT USED
' Not a issue though, they are the same
'Sub LeftFlipper_Collide(parm)
'  RandomSoundFlipper()
'End Sub

'Sub RightFlipper_Collide(parm)
'  RandomSoundFlipper()
'End Sub

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
  Volume = ABS(ActiveBall.vely ^3) / 50
  If ActiveBall.z > 24 Then
    If PFOption = 1 Or PFOption = 2 Then
      PlaySound "ApronHit", 0, Volume * xGain(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, 0  'Left & Right stereo or Top & Bottom stereo PF Speakers.
    End If
    If PFOption = 3 Then
      PlaySound "ApronHit", 0, Volume *    XVol(ActiveBall)  *    YVol(ActiveBall),  -1, 0, Pitch(ActiveBall), 1, 0, -1 'Top Left PF Speaker
      PlaySound "ApronHit", 0, Volume * (1-XVol(ActiveBall)) *    YVol(ActiveBall),   1, 0, Pitch(ActiveBall), 1, 0, -1 'Top Right PF Speaker
      PlaySound "ApronHit", 0, Volume *    XVol(ActiveBall)  * (1-YVol(ActiveBall)), -1, 0, Pitch(ActiveBall), 1, 0,  1 'Bottom Left PF Speaker
      PlaySound "ApronHit", 0, Volume * (1-XVol(ActiveBall)) * (1-YVol(ActiveBall)),  1, 0, Pitch(ActiveBall), 1, 0,  1 'Bottom Right PF Speaker
    End If
  End If
'Tb2.text= "Volume=" & round(Volume,2) & "     Adjust Volume devisor to keep Volume around 0.6 when dropping balls on the apron using ball control"
End Sub

Sub Saucers_Hit(idx)
  If PFOption = 1 Or PFOption = 2 Then
    PlaySound "metalhit_medium", 0, 1 * xGain(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall)-11025, 0, 0, 0  'Left & Right stereo or Top & Bottom stereo PF Speakers.
  End If
  If PFOption = 3 Then
    PlaySound "metalhit_medium", 0, 1 *    XVol(ActiveBall)  *    YVol(ActiveBall),  -1, 0, Pitch(ActiveBall)-11025, 0, 0, -1 'Top Left PF Speaker
    PlaySound "metalhit_medium", 0, 1 * (1-XVol(ActiveBall)) *    YVol(ActiveBall),   1, 0, Pitch(ActiveBall)-11025, 0, 0, -1 'Top Right PF Speaker
    PlaySound "metalhit_medium", 0, 1 *    XVol(ActiveBall)  * (1-YVol(ActiveBall)), -1, 0, Pitch(ActiveBall)-11025, 0, 0,  1 'Bottom Left PF Speaker
    PlaySound "metalhit_medium", 0, 1 * (1-XVol(ActiveBall)) * (1-YVol(ActiveBall)),  1, 0, Pitch(ActiveBall)-11025, 0, 0,  1 'Bottom Right PF Speaker
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
dim LFPress, RFPress, EOST, EOSA, EOSTnew, EOSaNew
dim FStrength, FRampUp, fElasticity, EOSRampUp, SOSRampUp
dim RFEndAngle, LFEndAngle, LF1EndAngle, RF1EndAngle, LFCount, RFCount, LiveCatch

LFEndAngle = Leftflipper.EndAngle
RFEndAngle = RightFlipper.EndAngle

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
Sub flipperTimer_Timer
' tb.text = RotoArray (roto.objroty/24)
  Pgate002.rotx = -Gate002.CurrentAngle*0.5
' tb1.text = (roto.objroty/24)
    Dim BOT, b
    BOT = GetBalls

  lFlip.rotz = leftflipper.CurrentAngle -121 'silver metal flipper obj
  lFlipR.rotz = leftflipper.CurrentAngle -121
  rFlip.rotz = RightFlipper.CurrentAngle +121
  rFlipR.rotz = RightFlipper.CurrentAngle +121

  FlipperLShadow.RotZ = LeftFlipper.CurrentAngle
  FlipperRShadow.RotZ = RightFlipper.CurrentAngle


' gates004.rotx = gate1.currentangle
' gates005.rotx = gate003.currentangle

  '--------------Flipper Tricks Section
  'What this code does is swing the flipper fast and make the flipper soft near its EOS to enable live catches.  It resets back to the base Table
  'settings once the flipper reaches the end of swing.  The code also makes the flipper starting ramp up high to simulate the stronger starting
  'coil strength and weaker at its EOS to simulate the weaker hold coil.

  If LeftFlipper.CurrentAngle = LeftFlipper.EndAngle and LFPress = 1 Then   'If the flipper is fully swung and the flipper button is pressed then...
    LeftFlipper.eosTorqueAngle = EOSaNew  'sets flipper EOS Torque Angle to .2
    LeftFlipper.eosTorque = EOStNew     'sets flipper EOS Torque to 1
    LeftFlipper.RampUp = EOSRampUp      'sets flipper ramp up to 1.5
    If LFCount = 0 Then LFCount = GameTime  'sets the variable LFCount = to the elapsed game time
    If GameTime - LFCount < LiveCatch Then  'if less than 8ms have elasped then we are in a "Live Catch" scenario
      LeftFlipper.Elasticity = 0.1    'sets flipper elasticity WAY DOWN to allow Live Catches
      If LeftFlipper.EndAngle <> LFEndAngle Then LeftFlipper.EndAngle = LFEndAngle  'Keep the flipper at its EOS and don't let it deflect
    Else
      LeftFlipper.Elasticity = fElasticity  'reset flipper elasticity to the base table setting
    End If
  Elseif LeftFlipper.CurrentAngle > LeftFlipper.startangle - 0.05  Then   'If the flipper has started its swing, make it swing fast to nearly the end...
    LeftFlipper.RampUp = SOSRampUp        'set flipper Ramp Up high
    LeftFlipper.EndAngle = LFEndAngle - 3   'swing to within 3 degrees of EOS
    LeftFlipper.Elasticity = fElasticity    'Set the elasticity to the base table elasticity
    LFCount = 0
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
    If RFCount = 0 Then RFCount = GameTime
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
End Sub

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
    x.TimeDelay = 69    '*****Important, this variable is an offset for the speed that the ball travels down the table to determine if the flippers have been fired
              'This is needed because the corrections to ball trajectory should only applied if the flippers have been fired and the ball is in the trigger zones.
              'FlipAT is set to GameTime when the ball enters the flipper trigger zones and if GameTime is less than FlipAT + this time delay then changes to velocity
              'and trajectory are applied.  If the flipper is fired before the ball enters the trigger zone then with this delay added to FlipAT the changes
              'to tragectory and velocity will not be applied.  Also if the flipper is in the final 20 degrees changes to ball values will also not be applied.
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
  LF.EndPoint = EndPointLp  'you can use just a coordinate, or an object with a .x property. Using a couple of simple primitive objects
  RF.Object = RightFlipper
  RF.EndPoint = EndPointRp
End Sub

Sub AddPt(aStr, idx, aX, aY)  'debugger wrapper for adjusting flipper script in-game
  dim a : a = Array(LF, RF)
  dim x : for each x in a
    x.addpoint aStr, idx, aX, aY
  Next
End Sub


Sub TriggerLF_Hit() : LF.Addball activeball : End Sub
Sub TriggerLF_UnHit() :  LF.PolarityCorrect activeball : End Sub
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
  End Sub

  Public Sub ProcessBalls() 'save data of balls in flipper range
    FlipAt = GameTime
    dim x : for x = 0 to uBound(balls)
      if not IsEmpty(balls(x) ) then balldata(x).Data = balls(x)
    Next
    PartialFlipCoef = ((Flipper.StartAngle - Flipper.CurrentAngle) / (Flipper.StartAngle - Flipper.EndAngle))  '% of flipper swing
    PartialFlipCoef = abs(PartialFlipCoef-1)
    if abs(Flipper.CurrentAngle - Flipper.EndAngle) < 20 Then 'last 20 degrees of swing is not dealt with
      PartialFlipCoef = 0
    End If
'   tb.text = FlipAT
  End Sub

'***********gameTime is a global variable of how long the game has progressed in ms
'***********This function lets the table know if the flipper has been fired
  Private Function FlipperOn()
'   TB.text = gameTime & ":" & (FlipAT + TimeDelay) ' ******MOVE TB into view WHEN THIS FLIPPER FUNCTIONALITY IS ADDED TO A NEW TABLE TO CHECK IF THE TIME DELAY IS LONG ENOUGH*****
    if gameTime < FlipAt + TimeDelay then FlipperOn = True
  End Function  'Timer shutoff for polaritycorrect

'***********This is turned on when a ball leaves the flipper trigger area
  Public Sub PolarityCorrect(aBall)
    if FlipperOn() then 'don't run this if the flippers are at rest
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

'   TB.text = coef

'               Applies the coef to x and y velocities
' Thalamus - patched : ' Thalamus - patched :     aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
    aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
' Thalamus - patched :     aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
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
  If B2sOn Then controller.stop
  saveHighScore
End Sub


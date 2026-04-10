'****************************************************************
'
'        Fashion Show (Gottlieb 1962)
'         Script Code by Scottacus
'              April 2021
'
' Basic DOF config
'   220 Left Flipper, 221 Right Flipper
'   203 Left sling, 204 Right sling
'   205 Bumper1,  206 Bumper2, 207 Bumper3
'   210 - 216 Top Rollovers
'   217 - 218 Side Rollovers
'   219 -220 Out Lanes
'   230 Middle Target
'   109 Launch Ball
'   124 Drain, 125 Ball Release
'   127 credit light
'   128 Knocker, 129 Knocker and Kicker Strobe
'   160 Ball In Shooter Lane
'   153 Chime 10-100s
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

Const cGameName = "FashionShow_1962"
Const cOptions = "FashionShow_1962.txt"
Const hsFileName = "Fashion Show (Gottlieb 1962)"


'**********************************************************************************************
'Set this variable to 1 to save a PinballY High Score file to your Tables Folder
'this will let the Pinball Y front end display the high scores when searching for tables
'0 = No PinballY High Scores, 1 = Save PinballY High Scores
Const cPinballY = 2
'**********************************************************************************************
  '**********************************************************************************************
  'Set BallLift to 0 to enable Automatic Mechanical Ball Lift, 1 will require start key or right magnasave to lift ball
  dim BallLiftOn
  BallLiftOn = 1
  '**********************************************************************************************

Dim Balls
Dim Replays
Dim Add1, Add10, Add100, Add1000
Dim HiSc
Dim MaxPlayers
Dim Players
Dim Player
Dim Credit
Dim Score(6)
Dim HScore(6)
Dim SReels(6)
Dim State
Dim Tilt
Dim TiltSens
Dim Target(9)
Dim BallinPlay
Dim MatchNumber
Dim BallREnabled
Dim RStep, LStep
Dim ReplayValue
Dim EndGame
Dim Bell
Dim i,j, f, ii, Object, Light, x, y
Dim AwardCheck
Dim BStop
Dim FreePlay
Dim BallInLane
Dim Ballsize,BallMass
Dim BallHomeCheck
Dim HSArray
Dim HSiArray
Dim Shift
Dim Launched
Dim Button
Dim FlipperArray
Dim GateState
Dim TargetA, TargetB
Dim BellRing
Dim HSInitial0, HSInitial1, HSInitial2
Dim RoundHS, RoundHSPlayer
Dim ScoreMil, Score100K, Score10K, ScoreK, Score100, Score10, ScoreUnit
HSArray = Array("HS_0","HS_1","HS_2","HS_3","HS_4","HS_5","HS_6","HS_7","HS_8","HS_9","HS_Space","HS_Comma","HSi_1","HSi_2","HSi_3")
HSiArray = Array("HSi_0","HSi_1","HSi_2","HSi_3","HSi_4","HSi_5","HSi_6","HSi_7","HSi_8","HSi_9","HSi_10","HSi_11","HSi_12","HSi_13","HSi_14","HSi_15","HSi_16","HSi_17","HSi_18","HSi_19","HSi_20","HSi_21","HSi_22","HSi_23","HSi_24","HSi_25","HSi_26")
Dim ShowBallShadow
BallSize = 50
BallMass = 1
Dim Chime
Dim Options
Dim PFOption
Dim BonusScore
Dim LayoutDifficulty
Dim BIP
Dim RotAngle
Dim DiskPts

Sub Table1_init
  LoadEM
  MaxPlayers=2
  Set SReels(1) = ScoreReel1
  Set SReels(2) = ScoreReel2
  Player = 1
  LoadHighScore
    If HighScore(0)="" Then HighScore(0)=1200
    If HighScore(1)="" Then HighScore(1)=1000
    If HighScore(2)="" Then HighScore(2)=800
    If HighScore(3)="" Then HighScore(3)=600
    If HighScore(4)="" Then HighScore(4)=400
  If Initial(0,1) = "" Then
    Initial(0,1) = 2: Initial(0,2) = 18: Initial(0,3) = 4
    Initial(1,1) = 1: Initial(1,2) = 1: Initial(1,3) = 1
    Initial(2,1) = 2: Initial(2,2) = 2: Initial(2,3) = 2
    Initial(3,1) = 3: Initial(3,2) = 3: Initial(3,3) = 3
    Initial(4,1) = 4: Initial(4,2) = 4: Initial(4,3) = 4
  End If
  If Credit = "" Then Credit = 0
  If FreePlay = "" Then FreePlay = 1
  If Balls = "" Then Balls = 5
  If Score(1) = "" Then Score(1) = 0
  If Score(2) = "" Then Score(2) = 0
  If Chime = "" Then Chime = 0
  If MatchNumber = "" Then MatchNumber = 4
  If ShowDT = True Then PFOption = 1
  If PFOption = "" Then PFOption = 1
  If LayoutDifficulty = "" Then LayoutDifficulty = 0
  If FreePlay = 1 Then
    If B2SOn Then DOF 127, DOFOn
  End If
  If RotAngle = "" Then RotAngle = 94
  CalcCenterScore
  RotateDisk

  CheckReplay

  contball = 0

  If LayoutDifficulty = 1 Then
    'Liberal
    render1postsl.visible = 1
    screwsl.visible = 1
    lslingl.visible = 1
    rslingl.visible = 1
    sling_l.collidable = 1

    render1postsc.visible = 0
    screwsc.visible = 0
    lsling.visible = 0
    rsling.visible = 0
    sling_c.collidable = 0


  Else
    'Conservative
    render1postsc.visible = 1
    screwsc.visible = 1
    lsling.visible = 1
    rsling.visible = 1
    sling_c.collidable = 1

    render1postsl.visible = 0
    screwsl.visible = 0
    lslingl.visible = 0
    rslingl.visible = 0
    sling_l.collidable = 0
  End If

  FirstBallOut = 0
  UpdatePostIt
  TiltReel1.SetValue(1)
  TiltReel2.SetValue(1)
  CreditReel.setvalue(Credit)

  BallinPlay = 0

  SetBackGlass.Enabled = 1

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

  SReels(1).setvalue(Score(1))
  SReels(2).setValue(Score(2))
  If B2SOn and ShowDT = False Then Controller.B2SSetScorePlayer 1, Score(1)
  If B2SOn and ShowDT = False Then Controller.B2SSetScorePlayer 2, Score(2)
  Tilt = False
  State = False
  GameState
' Initialize = 1
  If B2SOn Then
    Controller.B2SSetGameOver 35,0
  End If

  InstructCard.image = "InstructionCard" & FreePlay

End Sub


'***********KeyCodes
Dim enableInitialEntry, firstBallOut
Sub Table1_KeyDown(ByVal keycode)

  If EnableInitialEntry = True Then EnterInititals(keycode)

  If keycode = AddCreditKey Then
    PlayFieldSound "coinin",0,Drain,1
    AddCredit = 1
    ScoreMotor5.enabled = 1
    End If

    If keycode = StartGameKey Then
    If EnableInitialEntry = False and operatormenu = 0 Then
      If FreePlay = 1 and FirstBallOut = 0 and Players < 2 Then StartGame
      If FreePlay = 0 and Credit > 0 and Players < 2 and FirstBallOut = 0 Then
        Credit = Credit - 1
        Playsound "Reel"
        CreditReel.setvalue(Credit)
        If B2SOn Then
          If Freeplay = 0 Then Controller.B2SSetCredits Credit
          If FreePlay = 0 and Credit < 1 Then DOF 127, DOFOff
        End If
        StartGame
      End If
    End If
  End If

  If Keycode = StartGameKey and contball = 0 and Lifter = 0 and State = True and BIP = 0 Then
    BallLiftTimer.enabled = 1
  End If

  If keycode = PlungerKey Then
    Plunger.PullBack
    PlaySound "plungerpull",0,1,0.25,0.25
  End If

  If Tilt = False and State = True Then
  If keycode = LeftFlipperKey and contball = 0 Then
    LeftFlipper.RotateToEnd
    PlayFieldSound "FlipUpL", 0, LeftFlipper, 1
    If B2SOn Then DOF 220,DOFOn
    PlayFieldSound "FlipBuzzL", -1, LeftFlipper, 1
  End If

  If keycode = RightFlipperKey and contball = 0 Then
    RightFlipper.RotateToEnd
    PlayFieldSound "FlipUpR", 0, RightFlipper,1
    If B2SOn Then DOF 221,DOFOn
    PlayFieldSound "FlipBuzzR", -1, RightFlipper,1
  End If

  If keycode = LeftTiltKey Then
    Nudge 90, 2
    CheckTilt
  End If

  If keycode = RightTiltKey Then
    Nudge 270, 2
    CheckTilt
  End If

  If keycode = CenterTiltKey Then
    Nudge 0, 2
    CheckTilt
  End If
  End if

  If Keycode = MechanicalTilt Then
    PlayerTilt(Player) = 1
    Tilt = True
    EVAL("TiltReel" & Player).SetValue(1)
        If B2SOn Then Controller.B2SSetTilt (80+Player),1
    FreshTilt = 1
    TurnOff
  End If

    If keycode = LeftFlipperKey and State = False and OperatorMenu = 0 and EnableInitialEntry =  False Then
        OperatorMenuTimer.Enabled = true
    end if


    If keycode = LeftFlipperKey and State = False and OperatorMenu = 1 Then
    Options = Options + 1
    If ShowDt = True Then If Options = 3 Then Options = 5
        If Options =  6 Then Options = 0
    OptionMenu.visible = true
         PlayFieldSound "target", 0, Speaker5, 1.5
        Select Case (Options)
            Case 0:
                OptionMenu.image = "FreeCoin" & FreePlay
            Case 1:
        TextBox1.text = Balls
                OptionMenu.image = Balls & "Balls"
      Case 2:
        OptionMenu1.visible = 1
        OptionMenu1.image = "Layout"
        OptionMenu.image = "Difficulty" & LayoutDifficulty
      Case 3:
        OptionMenu1.image = "DOF"
        OptionMenu.image = "Chime" & Chime
      Case 4:
        OptionMenu.image = "UnderCab"
        OptionMenu1.image = "Sound" & PFOption
        OptionMenu2.visible = 1
        OptionMenu2.image = "SoundChange"
        Select Case (PFOption)
          Case 1: Speaker1.visible = 1: Speaker2.visible = 1: Speaker3.visible = 0: Speaker4.visible = 0
          Case 2: Speaker5.visible = 1: Speaker6.visible = 1: Speaker1.visible = 0: Speaker2.visible = 0
          Case 3: Speaker1.visible = 1: Speaker2.visible = 1: Speaker3.visible = 1: Speaker4.visible = 1: Speaker5.visible = 0: Speaker6.visible = 0
        End Select

      Case 5:
        For x = 1 to 6
          EVAL("Speaker" & x).visible = 0
        Next
        OptionMenu1.visible = 0
        OptionMenu.image = "SaveExit"
        OptionMenu2.visible = 0
        End Select
    End If

    If keycode = RightFlipperKey and State = False and OperatorMenu = 1 Then
      PlayFieldSound "metalhit2", 0, Speaker5, 1.5
      Select Case (Options)
    Case 0:
            If FreePlay = 0 Then
                FreePlay = 1
              Else
                FreePlay = 0
            End If
            OptionMenu.image= "FreeCoin" & FreePlay
      If FreePlay = 0 Then
        If Credit > 0 and B2SOn Then DOF 127, DOFOn
        If Credit < 1 and B2SOn Then DOF 127, DOFOff
      Else
        If B2SOn Then DOF 127, DOFOn
      End If
        Case 1:
            If Balls = 3 Then
                Balls = 5
        InstructCard.image = "InstructionCard1"
              Else
                Balls = 3
        InstructCard.image = "InstructionCard0"
            End If
            OptionMenu.image = Balls & "Balls"
    Case 2: If LayoutDifficulty = 0 Then
          LayoutDifficulty = 1 'Liberal
          render1postsl.visible = 1
          screwsl.visible = 1
          lslingl.visible = 1
          rslingl.visible = 1
          sling_l.collidable = 1
          sling_c.collidable = 0
          render1postsc.visible = 0
          screwsc.visible = 0
          lsling.visible = 0
          rsling.visible = 0



        Else
          LayoutDifficulty = 0 'Conservative
          render1postsc.visible = 1
          screwsc.visible = 1
          lsling.visible = 1
          rsling.visible = 1
          sling_l.collidable = 0
          sling_c.collidable = 1
          render1postsl.visible = 0
          screwsl.visible = 0
          lslingl.visible = 0
          rslingl.visible = 0
        End If
        OptionMenu.image = "Difficulty" & LayoutDifficulty
    Case 3:
            If Chime = 0 Then
                Chime= 1
        If B2SOn Then DOF 153,DOFPulse
              Else
                Chime = 0
        Playsound "Bell10"
            End If
      OptionMenu.image = "Chime" & Chime
    Case 4:
      OptionMenu1.visible = 1
      PFOption = PFOption + 1
      If PFOption = 4 Then PFOption = 1
      OptionMenu1.image = "Sound" & PFOption

      Select Case (PFOption)
        Case 1: Speaker1.visible = 1: Speaker2.visible = 1: Speaker3.visible = 0: Speaker4.visible = 0
        Case 2: Speaker5.visible = 1: Speaker6.visible = 1: Speaker1.visible = 0: Speaker2.visible = 0
        Case 3: Speaker1.visible = 1: Speaker2.visible = 1: Speaker3.visible = 1: Speaker4.visible = 1: Speaker5.visible = 0: Speaker6.visible = 0
      End Select
        Case 5:
            OperatorMenu = 0
            saveHighScore
      DynamicUpdatePostIt.enabled = 1
      OptionMenu.image = "FreeCoin" & FreePlay
            OptionMenu1.visible = 0
      OptionMenu.visible = 0
      OptionsMenu.visible = 0
      CheckReplay
    End Select
    End If

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

    If keycode = 203 Then Cleft = 1' Left Arrow

    If keycode = 200 Then Cup = 1' Up Arrow

    If keycode = 208 Then Cdown = 1' Down Arrow

    If keycode = 205 Then Cright = 1' Right Arrow

  If Keycode = RightMagnaSave and contball = 0 and BIP = 0 and Lifter = 0 and State = True Then
    BallLiftTimer.enabled = 1
  End If




'************************Start Of Test Keys****************************
  If keycode = 30 Then
    RotateCircleTimer.enabled = 1
  End If

  If keycode= 31 Then
    MatchNumber = MatchNumber + 1
    If MatchNumber > 10 Then MatchNumber = 1
    MatchReel.SetValue(MatchNumber)
    If B2SOn then Controller.B2SSetMatch 34, MatchNumber*10
    Match

  End If
'
'************************End Of Test Keys****************************
End Sub

dim TestChr1, TestChr2, TestChr3, TestChr, TestFlag
Sub Table1_KeyUp(ByVal keycode)

  If keycode = PlungerKey Then
    Plunger.Fire
    PlaySound "PlungerFire",0,1,0.25,0.25
  End If

    If keycode = LeftFlipperKey Then
        OperatorMenuTimer.Enabled = False
    End If

   If Tilt = False and State = True Then
    If keycode = LeftFlipperKey and contball = 0 Then
      LeftFlipper.RotateToStart
      PlayFieldSound "FlipDownL", 0, LeftFlipper, 1
      If B2SOn Then DOF 220,DOFOff
      StopSound "FlipBuzzLA"
      StopSound "FlipBuzzLB"
      StopSound "FlipBuzzLC"
      StopSound "FlipBuzzLD"
    End If

    If keycode = RightFlipperKey and contball = 0 Then
      RightFlipper.RotateToStart
      PlayFieldSound "FlipDownR", 0, RightFlipper, 1
      If B2SOn Then DOF 221,DOFOff
      StopSound "FlipBuzzRA"
      StopSound "FlipBuzzRB"
      StopSound "FlipBuzzRC"
      StopSound "FlipBuzzRD"
    End If

    If keycode = RightMagnaSave Then PlayFieldSound "BallLiftDown", 0, Plunger, .7
   End If

    If keycode = 203 Then Cleft = 0' Left Arrow

    If keycode = 200 Then Cup = 0' Up Arrow

    If keycode = 208 Then Cdown = 0' Down Arrow

    If keycode = 205 Then Cright = 0' Right Arrow

End Sub

'***********Operator Menu
Dim operatormenu

Sub OperatorMenuTimer_Timer
  Options = 0
    OperatorMenu = 1
    Displayoptions
End Sub

Sub DisplayOptions
  DynamicUpdatePostIt.enabled = 0
  UpdatePostIt
  Options = 0
    OptionsMenu.visible = True
    OptionMenu.visible = True
  OptionMenu.image = "FreeCoin" & FreePlay
End Sub

Sub CheckReplay
  If Balls = 3 Then
    Replay(1) = 600
    Replay(2) = 800
    Replay(3) = 900
    Replay(4) = 1000
  Else
    Replay(1) = 1000
    Replay(2) = 1200
    Replay(3) = 1300
    Replay(4) = 1400
  End If
End Sub


Dim backGlassOn
Sub SetBackglass_Timer
  dim xx
  PlayFieldSound "poweron",0, Plunger, 1
  centerlight.visible = 1
  centerlight1.state = 1
  playfieldmeshoff.visible = 0
  For each Light in GIlights:Light.state = 1: Next
  For each xx in render1: xx.image = "render1": Next
  For each xx in render2: xx.image = "render2": Next
  For each xx in render3: xx.image = "render3": Next
  For each xx in bumpercaps: xx.image = "bumpcap": Next
  For each xx in bumperbases: xx.image = "bumpbase": Next
  bump2.image="bumpcaplit"
  bump21.image="bumpbaselit"
  plasticwheel.image = "circlestill"
  plasticwheel1.image = "circlestillmetal"
  lockdown.image="lockdown"
  centertarget.blenddisablelighting = 1
  centertarget_switch.blenddisablelighting = 1
    MatchReel.SetValue(MatchNumber)
    If B2SOn Then
      Controller.B2SSetData 80,1
      Controller.B2SSetData 81,0
      COntroller.B2SSetData 82,0
      Controller.B2SSetData 25,0
      Controller.B2SSetData 26,0
      Controller.B2SSetCredits Credit
      Controller.B2SSetMatch 34, MatchNumber * 10
      Controller.B2SSetGameOver 35,1
      Controller.B2SSetTilt 33,1
      If Credit > 0 Then DOF 127, DOFOn
      If FreePlay = 1 Then DOF 127, DOFOn
    End If
  BackGlassOn = 1
  SetBackglass.enabled = 0
End Sub

Sub FlipperTimer_timer()
  LFlip.RotZ = LeftFlipper.CurrentAngle
  RFlip.RotZ = RightFlipper.CurrentAngle
  LFlip1.RotZ = LeftFlipper.CurrentAngle
  RFlip1.RotZ = RightFlipper.CurrentAngle
  FlipperLSh.RotZ = LeftFlipper.CurrentAngle
  FlipperRSh.RotZ = RightFlipper.CurrentAngle
End Sub

'***********Start Up Game
Sub StartGame
  If State = False Then
    If B2SOn Then
      Controller.B2SSetCredits Credit
      Controller.B2SSetData 41, 1
      Controller.B2SSetGameOver 0
      Controller.B2SSetTilt 81, 0   'tilt 1
      Controller.B2SSetTilt 82, 0   'tilt 2
      Controller.B2SSetMatch 34, 0
      Controller.B2SSetBallInPlay 32, 0
      Controller.B2SSetCanPlay 31, 1
      controller.B2SSetData 25, 0   '1000 light 1
      controller.B2SSetData 26, 0   '1000 light 2
    End If
    Bumper2Light1.state = 1
    Bump2.image = "bumpcaplit"
    Bump21.image = "bumpbaselit"
    GameOverReel.SetValue(0)
    MatchReel.SetValue(0)
    TiltReel1.SetValue(0)
    TiltReel2.SetValue(0)
    DynamicUpdatePostIt.enabled = 0
    Reel1001.SetValue(0)
    Reel1002.SetValue(0)
    Player1UpReel.setValue(1)
    CreditReel.setvalue(Credit)
    BIPReel.setValue(1)
    UpdatePostIt
    BallinPlay = 0
    Tilt = False
    FreshTilt = 0
    State = True
    Players = 1
    Score(1) = Score(1) mod 1000 'this removes any 1000 or greater digits from the score
    Score(2) = Score(2) mod 1000
    PlayersReel.setValue(1)
    BallOut = 0
    BIP = 0
    PlayerTilt(1) = 0
    PlayerTilt(2) = 0
    Player = 1
    RoundHS = 0
    RoundHSPlayer = 1
    EndGame = 0
    RoundHS = 0
    RoundHSPlayer = 1
    ResetLights
    For x = 1 to 3
      EVAL("Bumper" & x).hashitevent = 1
    Next
    ResetReel.enabled = 1
    GameStart.enabled = 1
    BallinPlay = 1
    If B2SOn Then Controller.B2SSetBallInPlay 32, 1
    BIPReel.setValue(1)
  Else If State = True and Players < MaxPlayers and FirstBallOut = 0 Then
    Players = Players + 1
    PlayersReel.setValue(2)
    CreditReel.setvalue(Credit)
      If B2SOn Then
        Controller.B2SSetCredits Credit
        Controller.B2SSetCanplay 31, Players
      End If
    End If
  End If
End Sub

'*******************Ball Drained
Sub BallDrained
  Flag100 = 0: Flag10 = 0:  Flag1 = 0
  BIP = 0
  AdvancePlayers
End Sub

'**********Advance Players
Sub AdvancePlayers
   If Players = 1 or Player = Players Then
    Player = 1
    Player1UpReel.setValue(1)
    Player2UpReel.setValue(0)
   Else
    Player = Player + 1
    Player1UpReel.setValue(0)
    Player2UpReel.setValue(1)
  End If

  If PlayerTilt(1) = 1 and Players = 1 Then
    EndGame = 1
    CheckContinue
    Exit Sub
  End If
  If PlayerTilt(1) = 1 and Players = 2 Then
    Player = 2
    Player1UpReel.setValue(0)
    Player2UpReel.setValue(1)
    If PlayerTilt(2) = 1 Then
      EndGame = 1
      CheckContinue
      Exit Sub
    End If
  End If
  If PlayerTilt(2) = 1 Then
    Player = 1
    Players = 1
    Player1UpReel.setValue(1)
    Player2UpReel.setValue(0)
    If PlayerTilt(1) = 1 Then
      EndGame = 1
      CheckContinue
      Exit Sub
    End If
  End If

  If Player = 1 Then BallinPlay = BallinPlay + 1

  If PlayerTilt(1)  = 1 and FreshTilt = 0 Then
    BallinPlay = BallinPlay + 1
  End If

  If BallinPlay > Balls and Players = 2 and Player = 1 Then
    EndGame = 1
    CheckContinue
    Exit Sub
  End If
  If BallinPlay > Balls and Players = 1 Then
    EndGame = 1
    CheckContinue
    Exit Sub
  End If
  If B2SOn Then
    Controller.B2SSetData 41, 0
    Controller.B2SSetData 42, 0
    Controller.B2SSetData (40 + player), 1
  End If
  NextBall
End Sub

'**********Next Ball
Sub NextBall
    If Tilt = True Then
    Bumper1.hashitevent = 1
    Bumper2.hashitevent = 1
    Bumper3.hashitevent = 1
      Tilt = False
    If B2SOn Then
'     Controller.B2SSetTilt 33,0
'     Controller.B2SSetData 1, 1
    End If
    End If
' If player = 4 then BallinPlay = 5 'used for match testing

  CheckContinue
End Sub

'*************Check for Continuing Game
Sub CheckContinue
  If EndGame = 1 Then
    TurnOff
    StopSound "FlipBuzzLA"
    StopSound "FlipBuzzLB"
    StopSound "FlipBuzzLC"
    StopSound "FlipBuzzLD"
    StopSound "FlipBuzzRA"
    StopSound "FlipBuzzRB"
    StopSound "FlipBuzzRC"
    StopSound "FlipBuzzRD"
    Match
    State = False
    GameState
    DynamicUpdatePostIt.enabled = 1
    SortScores
    FirstBallOut = 0
    Players = 0
    Player1UpReel.setValue(0)
    Player2UpReel.setValue(0)
    PlayersReel.setValue(0)
    ResetLights
    For x = 1 to 4
      ReplayValue = 0
      RepAwarded(x) = 0
    Next
    saveHighScore
    If B2SOn Then
      Controller.B2SSetGameOver 35,1
      If Credit > 0 Then DOF 127, DOFOn
      If FreePlay = 1 Then DOF 127, DOFOn
    End If
    Else
      ResetLights
      If ballliftOn = 0 Then BallLiftTimer.enabled = 1
    End If
  End Sub


'***************Drain and Release Ball
Dim BallOut
Sub Drain_Hit
  ArchHit = 0
  Drain.DestroyBall
  BallDrained
End Sub

Dim ScoreMotorCount, AddCredit
Sub Scoremotor5_Timer
  ScoreMotorCount = ScoreMotorCount + 1
  PlayFieldSound "ScoreMotorSingleFire", 0, SoundPoint1, 0.1
  If ScoremotorCount = 5 Then
    If AddCredit = 1 Then
      Credit = Credit + 1
      Playsound "Reel"
      If Credit > 15 then Credit = 15
      CreditReel.setvalue(Credit)
      If B2SOn Then
        Controller.B2SSetCredits Credit
        If Credit > 0 Then DOF 127, DOFOn
      End If
      AddCredit = 0
    End If
    ScoreMotorCount = 0
    ScoreMotor5.enabled = 0
  End If
End Sub

'************Game State Check
Sub GameState
  If State = False Then
'   For each light in GIlights:light.state = 0: Next
    GameOverReel.SetValue(1)
    If B2SOn Then
      Controller.B2SSetGameOver 35, 1
      Controller.B2SSetBallInPlay 32, 0
    End If
    StopSound "FlipBuzzLA"
    StopSound "FlipBuzzLB"
    StopSound "FlipBuzzLC"
    StopSound "FlipBuzzLD"
    StopSound "FlipBuzzRA"
    StopSound "FlipBuzzRB"
    StopSound "FlipBuzzRC"
    StopSound "FlipBuzzRD"
  Else
'   For each Light in GIlights:Light.state = 1: Next
    For x= 1 to 3
      EVAL("Bumper" & x).hashitevent = 1
    Next
    Tilt = False
    GameOverReel.SetValue(0)
    MatchReel.SetValue(0)
    TiltReel1.SetValue(0)
    TiltReel2.SetValue(0)
    If B2SOn then
      Controller.B2SSetTilt 81,0
      Controller.B2SSetTilt 82,0
      Controller.B2SSetMatch 34,0
      Controller.B2SSetGameOver 35,0
    End If
  End If
End Sub

'*************Ball in Launch Lane
Sub BallHome_hit
  ArchHit = 0
  BallREnabled = 1
  If B2SOn then DOF 160, DOFOn
  Set ControlBall = ActiveBall
    contballinplay = True
End Sub

Sub BallHome_unhit
  If B2SOn Then DOF 160, DOFOff
End Sub

'******* for ball control script
Sub EndControl_Hit()
    contballinplay = false
End Sub

'************Check if Ball Out of Launch Lane
' Thalamus : This sub is used twice - this means ... this one IS NOT USED
' Not a issue though, they are the same
' Sub BallsInPlay_hit
'   If BallREnabled = 1 Then
'     BallREnabled = 0
'     BallInLane = False
'   End If
'   FirstBallOut = 1
' End Sub


'***************Match Gottlieb Style
Dim MatchDisplay
Sub Match
  If MatchNumber = (Score(1) mod 10) or MatchNumber - 10 = (Score(1) mod 10) Then
    AddCredit = 1
    ScoreMotor5.enabled = 1
    Playsound SoundFXDOF("Knocker",128,DOFPulse,DOFKnocker)
    If B2SOn Then DOF 151, DOFPulse
  End If
' TextBox1.text = MatchNumber
  MatchReel.SetValue(MatchNumber)
  MatchDisplay = MatchNumber
  If MatchDisplay = 0 Then MatchDisplay = 10
  If B2SOn Then Controller.B2SSetMatch 34, MatchDisplay * 10

End Sub

'***********Rotate Flipper Shadow
Sub FlipperShadowUpdate_Timer
  FlipperLSh.RotZ = LeftFlipper.CurrentAngle
  FlipperRSh.RotZ = RightFlipper.CurrentAngle
End Sub

'************** Reset Lights For New Game
Sub ResetLights
  Bumper1Light1.state = 0
  Bumper3Light1.state = 0
  tenslingleft.state = 0
  tenslingright.state = 0
  For x = 1 to 3
    EVAL("Bump" & x).image = "bumpcap"
    EVAL("Bump" & x & "1").image = "Bumpbase"
  Next
End Sub

'************** Bumpers
Sub Bumpers_hit(Index)
  If Tilt = False Then
    Select Case (Index)
      Case 0: PlayFieldSound "PopBump", 0, Bumper1, 1
          If B2SOn Then DOF 205, DOFPulse
          If Bumper1Light1.state = 0 Then
            AddScore 1
          Else
            AddScore 10
          End If
      Case 1: PlayFieldSound "PopBump", 0, Bumper2, 1
          If B2SOn Then DOF 206, DOFPulse
          AddScore 1

      Case 2: PlayFieldSound "PopBump", 0, Bumper3, 1
          If B2SOn Then DOF 207, DOFPulse
          If Bumper3Light1.state = 0 Then
            AddScore 1
          Else
            AddScore 10
          End If
      Case 3: PlayFieldSound "PopBump", 0, Bumper4, 1
          If B2SOn Then DOF 208, DOFPulse
          If Bumper4Light1.state = 0 Then
            AddScore 1
          Else
            AddScore 10
          End If
    End Select
  End If
End Sub

'************** Targets
Dim TStep, CenterScore
Sub MiddleTarget_Hit
  centertarget.TransY=-5
  If B2SOn Then DOF 230, DOFPulse
  TStep = 0
  If Tilt = False Then
    AddScore CenterScore
  End If
  RotateCircleTimer.enabled = 1
  centertimer.enabled = 1
  centerlight.visible = 0
  centerlight1.state = 0
  centertarget.blenddisablelighting = .25
End Sub

Sub CenterTimer_timer
  Select Case TStep
    Case 3: centertarget.TransY=3
    Case 4: centertarget.TransY=-4
    Case 5: centertarget.TransY=2
    Case 6: centertarget.TransY=-3
    Case 7: centertarget.TransY=1
    Case 8: centertarget.TransY=-2
    Case 9: centertarget.TransY=0
    Case 10: centertarget.TransY=-1: CenterTimer.enabled = 0 'RotateCircleTimer.Enabled
  End Select
  TStep=TStep + 1
End Sub

'************** Slings
Sub LeftSlingShot_Slingshot
  PlayfieldSound "SlingShot", 0, SoundPoint12, 1
  If B2SOn Then DOF 203,DOFPulse
  If LayoutDifficulty = 0 Then  'this determines which sling is visible based upon layout difficulty
    lsling.Visible = 0
    lsling1.Visible = 1
  Else
    lslingl.Visible = 0
    lslingl1.Visible = 1
  End If
    leftSling.Rotx = 27
  lstep = 0
  leftSlingShot.TimerEnabled = 1
  If Tilt = False Then
    If tenslingleft.state = 0 Then
      AddScore 1
    Else
      AddScore 10
    End If
  End If
End Sub

Sub leftSlingShot_Timer
    Select Case lStep
        Case 3: leftSling.Rotx = 10
        If LayoutDifficulty = 0 Then  'this determines which sling is visible based upon layout difficulty
          lsling1.Visible = 0
          lsling2.visible = 1
        Else
          lslingl1.Visible = 0
          lslingl2.visible=1
        End If
        Case 4: leftSling.Rotx = 0
        If LayoutDifficulty = 0 Then
          lsling2.Visible = 0
          lsling3.visible = 1
        Else
          lslingl2.Visible = 0
          lslingl3.visible=1
        End If
        Case 5: If LayoutDifficulty = 0 Then
          lsling3.Visible = 0
          lsling.visible = 1
        Else
          lslingl3.Visible = 0
          lslingl.visible=1
        End If
        leftSlingShot.TimerEnabled = 0
    End Select
    lStep = lStep + 1
End Sub

Sub RightSlingShot_Slingshot
  PlayfieldSound "SlingShot", 0, SoundPoint13, 1
  If B2SOn Then DOF 204,DOFPulse
  If LayoutDifficulty = 0 Then  'this determines which sling is visible based upon layout difficulty
    rsling.Visible = 0
    rsling1.Visible = 1
  Else
    rslingl.Visible = 0
    rslingl1.Visible = 1
  End If
    RightSling.Rotx = 27
  rstep = 0
  RightSlingShot.TimerEnabled = 1
  If Tilt = False Then
    If tenslingright.state = 0 Then
      AddScore 1
    Else
      AddScore 10
    End If
  End If
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3: RightSling.Rotx = 10
        If LayoutDifficulty = 0 Then  'this determines which sling is visible based upon layout difficulty
          RSling1.Visible = 0
          rsling2.visible = 1
        Else
          rslingl1.Visible = 0
          rslingl2.visible=1
        End If
        Case 4: RightSling.Rotx = 0
        If LayoutDifficulty = 0 Then
          RSling2.Visible = 0
          rsling3.visible = 1
        Else
          rslingl2.Visible = 0
          rslingl3.visible=1
        End If
        Case 5: If LayoutDifficulty = 0 Then
          RSling3.Visible = 0
          rsling.visible = 1
        Else
          rslingl3.Visible = 0
          rslingl.visible=1
        End If
        RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

'**********Bumper skirt animations

Dim Bump1step, Bump2step, Bump3step, Bump4step, DLbumpstep, DRbumpstep
'
Sub Bumper1_hit
  bump12.RotY=-5
  Bump1Step = 0
  Bumper1.timerenabled = 1
End Sub

Sub Bumper1_timer
  Select Case Bump1Step
    Case 3: bump12.RotY=-3
    Case 4: bump12.RotY=2
    Case 5: bump12.RotY=-1
    Case 6: bump12.RotY=1
    Case 7: bump12.RotY=0:Bumper1.timerenabled = 0
  End Select
  Bump1Step=Bump1Step + 1
End Sub

Sub Bumper2_hit
  bump22.RotY=5
  Bump2Step = 0
  Bumper2.timerenabled = 1
End Sub

Sub Bumper2_timer
  Select Case Bump2Step
    Case 3: bump22.RotY=3
    Case 4: bump22.RotY=-2
    Case 5: bump22.RotY=1
    Case 6: bump22.RotY=-1
    Case 7: bump22.RotY=0:Bumper2.timerenabled = 0
  End Select
  Bump2Step=Bump2Step + 1
End Sub

Sub Bumper3_hit
  bump32.RotY=-5
  Bump3Step = 0
  Bumper3.timerenabled = 1
End Sub

Sub Bumper3_timer
  Select Case Bump3Step
    Case 3: bump32.RotY=-3
    Case 4: bump32.RotY=2
    Case 5: bump32.RotY=-1
    Case 6: bump32.RotY=1
    Case 7: bump32.RotY=0:Bumper3.timerenabled = 0
  End Select
  Bump3Step=Bump3Step + 1
End Sub

Sub dbumper0_hit
  AddScore 1
End Sub

Sub dbumper1_hit
  AddScore 1
End Sub



'*************** Triggers
Dim Lane(7)
Sub TopLanes_hit(Index)
  If Tilt = False Then
    If B2SOn Then DOF 210 + Index, DOFPulse
    Select Case (Index)
      Case 0: AddScore 20
      Case 1: AddScore 30
      Case 2: AddScore 40
      Case 3: AddScore 50
      Case 4: AddScore 40
      Case 5: AddScore 30
      Case 6: AddScore 20
    End Select
  End If
End Sub

Sub Buttons_hit(Index)
  If Tilt = False Then
    Select Case (Index)
      Case 0: Bumper1Light1.state = 1: tenslingleft.state = 1:
          Bump1.image = "bumpcaplit": Bump11.image = "bumpbaselit"
      Case 1: Bumper3Light1.state = 1: tenslingright.state = 1:
          Bump3.image = "bumpcaplit": Bump31.image = "bumpbaselit"
      Case 2: Bumper1Light1.state = 0: Bumper3Light1.state = 0: tenslingleft.state = 0: tenslingright.state = 0
          Bump1.image = "bumpcap": Bump11.image = "bumpbase": Bump3.image = "bumpcap": Bump31.image = "bumpbase"
    End Select
  End If
End Sub


Dim TriggerIndex
Sub Triggers_Hit (Index)
  If Tilt = False Then
    If B2sOn Then DOF 217 + Index, DOFPulse
    Select Case (Index)
      Case 0: AddScore CenterScore: RotateCircleTimer.enabled = 1
      Case 1: AddScore CenterScore: RotateCircleTimer.enabled = 1
      Case 2: AddScore CenterScore: RotateCircleTimer.enabled = 1
      Case 3: AddScore CenterScore: RotateCircleTimer.enabled = 1
    End Select
  End If
End Sub

Sub phys_leafs_hit
  AddScore 1
End Sub

Sub RotateDisk
  RotAngle = RotAngle - 24
  if RotAngle < 0 Then RotAngle = 358
  circle.ObjRotZ = RotAngle
  circle1.ObjRotZ = RotAngle
  CalcCenterScore
End Sub

Sub CalcCenterScore
  DiskPts = (RotAngle - 22)/24
  Select Case DiskPts
    Case 7, 8, 9: CenterScore = 20
    Case 4, 5, 6: CenterScore = 30
    Case 1,2, 3: CenterScore = 40
    Case 0, 13, 14: CenterScore = 50
    Case 12, 11: CenterScore = 100
    Case 10: CenterScore = (100 * (Int(3*Rnd)+1))
  End Select
End Sub

Sub RotateCircleTimer_timer
  RotateDisk
  centerlight.visible = 1
  centerlight1.state = 1
  centertarget.blenddisablelighting = 1
  RotateCircleTimer.enabled = 0
End Sub


Sub BallLaunchedTrigger_hit
  Launched = Launched + 1
  If CLng(Launched) Mod 2 > 0 Then
    If B2SOn Then DOF 161 ,DOFPulse
  End If
  If B2SOn Then DOF 127, DOFOff
End Sub

Dim SpecialCollect

'**************RollUnderGate
Sub Gate1_Hit
  If Tilt = False Then AddScore 5
End Sub

'**************Special
Sub Special
  AddCredit = 1
  ScoreMotor5.enabled = 1
  Playsound SoundFXDOF("Knocker",128,DOFPulse,DOFKnocker)
  If B2SOn Then DOF 151, DOFPulse
End Sub

'************Check if Ball Out of Launch Lane
Sub BallsInPlay_hit
  If BallREnabled = 1 Then
    BallREnabled = 0
    BallInLane = False
  End If
  FirstBallOut = 1
End Sub


'***************Shooter Lane Gate Animation
Sub PGateTimer_Timer
Pgate.rotx = Gate0.CurrentAngle
End Sub

'***************Scoring Routines
Dim Flag10, Flag100, Flag1, Point, Point1, Ones
Sub AddScore(points)
  If Tilt = False Then
    If Points < 10 Then
'     Number Matching, decrement the match unit for each 1 point score
      Point1 = (Points mod 10)
      For x = 1 to Point1
        MatchNumber = MatchNumber - 1
        If MatchNumber < 0 Then MatchNumber = 9
      Next

      Ones = ((Score(1) + points) Mod 10)
      OnesUnit
      TotalUp(1)
      Exit Sub
    End If

    If Points < 100 Then
      BellRing = (Points / 10)
      If BellRing > 1 Then Point = 10: Flag10 = 1: ScoreMotor.enabled = 1
      If BellRing = 1 Then
        TotalUp(10)
        If Chime = 0 Then
          PlayFieldSound "Bell10", 0, SoundPoint13, 1
        Else
          If B2SOn Then DOF 153,DOFPulse
        End If
      End If
      Exit Sub
    End If

    If Points > 99 and Points < 1000 Then
      BellRing = (Points / 100)
      If BellRing > 1 Then Point = 100: Flag100 = 1: ScoreMotor.enabled = 1
      If BellRing = 1 Then
        TotalUp(100)
        If Chime = 0 Then
          PlayFieldSound "Bell10", 0, SoundPoint13, 1
        Else
          If B2SOn Then DOF 153,DOFPulse
        End If
      End If
      Exit Sub
    End If

  End If
End Sub

Dim ReplayX, RepAwarded(5), Replay(7), ReplayText(3), TopReplay
Sub TotalUp(Points)
  PlaySound "Reel"

' If Flag1 = 1 Then
'   If Chime = 0 Then
'     PlayFieldSound  "Bell10", 0, SoundPoint13, 1
'   Else
'     If B2SOn Then DOF 153,DOFPulse
'   End If
' End If

  If Flag10 = 1 Then
    If Chime = 0 Then
      PlayFieldSound "Bell10", 0, SoundPoint13, 1
    Else
      If B2SOn Then DOF 153,DOFPulse
    End If
  End If

  If Flag100 = 1 Then
    If Chime = 0 Then
      PlayFieldSound  "Bell10", 0, SoundPoint13, 1
    Else
      If B2SOn Then DOF 153,DOFPulse
    End If
  End If

  Score(Player) = Score(Player) + Points
  SReels(Player).addvalue(Points)

  If Score(Player) > 999 Then
    If B2SOn and Player = 1 Then controller.B2SSetData 25, 1
    If B2SOn and Player = 2 Then controller.B2SSetData 26, 1
    EVAL ("Reel100" & Player).SetValue(1)
  End If

  If B2SOn Then Controller.B2SSetScorePlayer Player, Score(Player)

  If Balls = 3 Then
    TopReplay = 4
  Else
    TopReplay = 3
  End If

  For ReplayX = ReplayValue + 1 to TopReplay
    If Score(1) => Replay(ReplayX) Then
      AddCredit = 1
      ScoreMotor5.enabled = 1
      ReplayValue = ReplayValue + 1
      Playsound SoundFXDOF("Knocker",128,DOFPulse,DOFKnocker)
    End If
  Next
End Sub

'**************Score Motor Routine
Dim ScoreMotorLoop, ScoreMotorDone
Sub ScoreMotor_timer
  ScoreMotorLoop = ScoreMotorLoop + 1
  If ScoreMotorLoop < 6 Then PlayFieldSound "ScoreMotorSingleFire", 0, SoundPoint1, 0.1

'  These Flags are passed by scores with multiple of 10, 100 or 1
  If Flag10 = 1 or Flag100 = 1 or Flag1 = 1 Then
    Select Case ScoreMotorLoop
      Case 1: TotalUp Point
      Case 2: If BellRing > 1 Then TotalUp Point
      Case 3: If BellRing > 2 Then TotalUp Point
      Case 4: if BellRing > 3 Then TotalUp Point
      Case 5: If BellRing > 4 Then TotalUp Point
      Case 6: ScoreMotorLoop = 0
          Flag10 = 0
          Flag100 = 0
          Flag1 = 0
          ScoreMotor.enabled = 0
    End Select
  End If
End Sub


Sub OnesUnit

End Sub

Dim PlayerTilt(2)
'***************Tilt
Sub CheckTilt
  If Tilttimer.Enabled = True Then
   TiltSens = TiltSens + 1
   If TiltSens = 3 Then
      Tilt = True
    FreshTilt = 1
    PlayerTilt(Player) = 1
    EVAL("TiltReel" & Player).SetValue(1)
        If B2SOn Then Controller.B2SSetTilt (80+Player),1
      TurnOff
   End If
  Else
   TiltSens = 0
   Tilttimer.Enabled = True
  End If
End Sub

Sub Tilttimer_Timer()
  Tilttimer.Enabled = False
End Sub

'***********Ball Shadow Update
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
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) + 13
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) - 13
        End If
        ballShadow(b).Y = BOT(b).Y + 10
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
End Sub

'************Reset Reels

'This Sub looks at each individual digit in each players score and sets them in an array RScore.  If the value is >0 and <9
'then the players score is increased by one times the position value of that digit (ie 1 * 1000 for the 1000's digit)
'If the value of the digit is 9 then it subtracts 9 times the postion value of that digit (ie 9*100 for the 100's digit)
'so that the score is not rolled over and the next digit in line gets incremented as well (ie 9 in the 10's positon gets
'incremented so the 100's position rolls up by one as well since 90 -> 100).  Lastly the RScore array values get incremented
'by one to get ready for the next pass.

'Dim RScore(4), ResetLoop, Test, ResetFlag, ReelFlag, ReelStop
'Sub CountUp
' For Test = 0 to 2
'   RScore(Test) = Int(Score(1)/(10^Test)) mod 10
' Next
 '
  'For x = 0 to 2
  ' If Rscore(x) > 0 And Rscore(x) < 9 Then Score(1) = Score(1) + 10^x
  ' If RScore(x) = 9 Then Score(1) = Score(1) - (9 * 10^x)
  ' If RScore(x) > 0 Then RScore(x) = RScore(x) + 1
  ' If RScore(x) = 10 Then RScore(x) = 0
' Next

' If Score(1) = 0 Then
'   ReelFlag = 1
'   ReplayValue = 0
'   RepAwarded(1) = 0
' End If
'End Sub

Dim RScore(4,5), ResetLoop, Test, PlayerTest, ResetFlag, ReelFlag, ReelStop, ReelDone(4)
Sub CountUp
  For PlayerTest = 1 to 2
    For Test = 0 to 2
      RScore(PlayerTest,Test) = Int(Score(PlayerTest)/10^Test) mod 10
    Next
  Next

  For PlayerTest = 1 to 2
    For x = 0 to 2
      If Rscore(PlayerTest, x) > 0 And Rscore(PlayerTest, x) < 9 Then Score(PlayerTest) = Score(PlayerTest) + 10^x
      If RScore(PlayerTest, x) = 9 Then Score(PlayerTest) = Score(PlayerTest) - (9 * 10^x)
      If RScore(PlayerTest, x) > 0 Then RScore(PlayerTest, x) = RScore(PlayerTest, x) + 1
      If RScore(PlayerTest, x) = 10 Then RScore(PlayerTest, x) = 0
    Next
  Next
  If Score(1) = 0 and Score(2) = 0 Then
    ReelFlag = 1
    For i = 1 to MaxPlayers
      Score(i) = 0
'     Rep(i) = 0
      RepAwarded(i) = 0
    Next
  End If
End Sub


'This Sub sets each B2S reel or Desdktop reels to their new values and then plays the score motor sound each time and the
'reel sounds only if the reels are being stepped

Sub UpdateReels
  If B2SOn and ShowDT = False Then Controller.B2SSetScorePlayer 1, Score(1)
  If B2SOn and ShowDT = False Then Controller.B2SSetScorePlayer 2, Score(2)
  If ShowDT = True Then SReels(1).setvalue (Score(1))
  If ShowDT = True Then SReels(2).setvalue (Score(2))
  PlayfieldSound "ScoreMotorSingleFire", 0, rsling3, 0.2
  If ReelStop = 0 Then Playsound "reel"
  If ReelFlag = 1 Then ReelStop = 1
End Sub


'This Timer runs a loop that calls the CountUp and UpdateReels routines to step the reels up five times and Then
'check to see if they are all at zero during a two loop pause and then step them the rest of the way to zero

Dim AllZeros
Sub ResetReel_Timer
  Score(1) = (Score(1) Mod 1000)
  Score(2) = (Score(2) Mod 1000)
  ResetLoop = ResetLoop + 1
  If ResetLoop = 1 and Score(1) = 0 and Score(2) = 0 Then
    ResetLoop = 0
'   If TestFlag = 0 Then ReleaseBall
    TestFlag = 0
    AllZeros = 1
    ResetReel.enabled = 0
    Exit Sub
  End If
  Select Case ResetLoop
    Case 1: CountUp: UpdateReels
    Case 2: CountUp: UpdateReels
    Case 3: CountUp: UpdateReels
    Case 4: CountUp: UpdateReels
    Case 5: CountUp: UpdateReels
    Case 6: If ReelStop = 1 Then
          ResetLoop = 0
          ReelFlag = 0
          ReelStop = 0
  '       If TestFlag = 0 Then ReleaseBall
          TestFlag = 0
          ResetReel.enabled = 0
          Exit Sub
        End If

    Case 7:
    Case 8: CountUp: UpdateReels
    Case 9: CountUp: UpdateReels
    Case 10: CountUp: UpdateReels
    Case 11: CountUp: UpdateReels
    Case 12: CountUp: UpdateReels:
      ResetLoop = 0
      ReelFlag = 0
      ReelStop = 0
'     If TestFlag = 0 Then ReleaseBall
      TestFlag = 0
      ResetReel.enabled = 0
      Exit Sub
  End Select
End Sub


'***********Ball Lift Speed Limiter to Prevent Loss of Balls
Dim Lifter, FreshTilt
Lifter = 0
Sub BallLiftTimer_Timer
  Lifter = Lifter + 1
  If Lifter = 1 Then
    If BallinPlay < Balls + 1 Then
      PlayFieldSound "BallLift" , 0, Plunger, .7
      BallLift.CreateBall
      BallLift.kick 0,88+(Int(RND*5)),1.56
      BIP = 1
    End If
  End If
  If Lifter = 2 Then
    Lifter = 0
    FreshTilt = 0
    BIPReel.setValue(BallinPlay)
    if B2SOn Then controller.B2SSetBallInPlay 32, BallinPlay
    BallLiftTimer.enabled = 0
  End If
End Sub

'************Shut Down and De-energize for Tilt
Sub TurnOff
  For x= 1 to 3
    EVAL("Bumper" & x).hashitevent = 0
  Next
  StopSound "FlipBuzzLA"
  StopSound "FlipBuzzLB"
  StopSound "FlipBuzzLC"
  StopSound "FlipBuzzLD"
  StopSound "FlipBuzzRA"
  StopSound "FlipBuzzRB"
  StopSound "FlipBuzzRC"
  StopSound "FlipBuzzRD"
    LeftFlipper.RotateToStart
  If B2SOn Then DOF 220, DOFOff
  RightFlipper.RotateToStart
  If B2SOn Then DOF 221, DOFOff
  ResetLights
End Sub

'************************************************************************
'                         Ball Control
'************************************************************************

Dim Cup, Cdown, Cleft, Cright, contball, contballinplay, ControlBall, bcboost
Dim bcvel, bcyveloffset, bcboostmulti

bcboost = 1 'Do Not Change - default setting
bcvel = 4 'Controls the speed of the ball movement
bcyveloffset = -0.01 'Offsets the force of gravity to keep the ball from drifting vertically on the table, should be negative
bcboostmulti = 3 'Boost multiplier to ball veloctiy (toggled with the B key)

Sub BallControl_Timer()
    If Contball and ContBallInPlay then
        If Cright = 1 Then
            ControlBall.velx = bcvel*bcboost
          ElseIf Cleft = 1 Then
            ControlBall.velx = - bcvel*bcboost
          Else
            ControlBall.velx=0
        End If
        If Cup = 1 Then
            ControlBall.vely = -bcvel*bcboost
          ElseIf Cdown = 1 Then
            ControlBall.vely = bcvel*bcboost
          Else
            ControlBall.vely= bcyveloffset
        End If
    End If
End Sub

' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Function AudioPan(TableObj) ' Calculates the pan for a TableObj based on the X position on the table. "table1" is the name of the table.  New dBm algorithm for accurate pan positioning.
    Dim tmp
    If PFOption=1 Then tmp = TableObj.x * 2 / table1.width-1
  If PFOption=2 Then tmp = TableObj.y * 2 / table1.height-1
  If tmp < 0 Then
    AudioPan = -((0.8745898957*(ABS(tmp)^12.78313661)) + (0.1264569796*(ABS(tmp)^1.000771219)))
  Else
    AudioPan = (0.8745898957*(ABS(tmp)^12.78313661)) + (0.1264569796*(ABS(tmp)^1.000771219))
  End If
End Function

Function xGain(TableObj)  'New dBm algorithm used with AudioPanNew to provide a Constant Power "pan".
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

Function XVol(tableobj)  'XVol algorithm calculates a PlaySound Volume parameter multiplier for a tableobj based on its X table position.
'New dBm algorithm to provide a Constant Power "pan" & "fade" to reduce the sound positioning errors inherent in PlaySound's PAN & FADE.
Dim tmpx
  If PFOption = 3 Then
    tmpx = tableobj.x * 2 / table1.width-1
    XVol = 0.3293074856*EXP(-0.9652695455*tmpx^3 - 2.452909811*tmpx^2 - 2.597701999*tmpx) 'XVol = 1 at PF Left, XVol = 0.32931 (-3dB for PlaySound's volume parameter) at PF Center and XVol = 0 at PF Right
  End If
End Function

Function YVol(tableobj)  'YVol algorithm calculates a PlaySound Volume parameter multiplier for a tableobj based on its Y table position.
  Dim tmpy
  If PFOption = 3 Then
    tmpy = tableobj.y * 2 / table1.height-1
    YVol = 0.3293074856*EXP(-0.9652695455*tmpy^3 - 2.452909811*tmpy^2 - 2.597701999*tmpy) 'YVol = 1 at PF Top, YVol = 0.32931 (-3dB for PlaySound's volume parameter) at PF Center and YVol = 0 at PF Bottom
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

' sounds are played based on the ball speed and position

' the routine checks first for deleted balls and stops the rolling sound.

' The For loop goes through all the balls on the table and checks for the ball speed and
' if the ball is on the table (height lower than 30) then then it plays the sound
' otherwise the sound is stopped, like when the ball has stopped or is on a ramp or flying.

' The sound is played using the VOL, PAN and PITCH functions, so the volume and pitch of the sound
' will change according to the ball speed, and the PAN function will change the stereo position according
' to the position of the ball on the table.

'************************************
' What you need to add to your table
'************************************

' a timer called RollingTimer. With a fast interval, like 10
' one collision sound, in this script is called fx_collide
' as many sound files as max number of balls, with names ending with 0, 1, 2, 3, etc
' for ex. as used in this script: fx_ballrolling0, fx_ballrolling1, fx_ballrolling2, fx_ballrolling3, etc

'Added new dBm algorithm to provide a Constant Power "pan" & "fade" to reduce the sound positioning errors inherent in PlaySound's PAN & FADE.
'XVol and YVol drive multiple PlaySound commands panned and faded to their maximum extents where PlaySound's PAN and FADE have the least error.
'Subtracting XVol or YVol from 1 yeilds a inverse response.  Added new script and sounds for the ball rolling along the top arch.
'Added the looped sound files needed to enable calling the same sound twice for stereo and four times for quad under-playfield speakers.
'Examples: fx_ballrollingA0 through fx_ballrollingD4, ArchHitA0 through ArchHitD4 and ArchRollA0 through ArchRollD4.

Const tnob = 5 ' total number of balls
ReDim rolling(tnob)
InitRolling

ReDim ArchRolling(tnob)
InitArchRolling

Dim ArchHit

Sub LowerArch_Hit
  Archhit = 1
End Sub

Sub NotOnArch_Hit
  ArchHit = 0
End Sub

Dim BallDrain
Sub SubwayTrigger_Hit
  ArchHit = 1
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
  Dim BOT, b
  BOT = GetBalls

  ' stop the sound of deleted balls
  For b = UBound(BOT) + 1 to tnob
    rolling(b) = False
    StopSound("fx_ballrollingA" & b)
    StopSound("fx_ballrollingB" & b)
    StopSound("fx_ballrollingC" & b)
    StopSound("fx_ballrollingD" & b)
  Next

  ' exit the sub if no balls on the table
  If UBound(BOT) = -1 Then Exit Sub

  ' play the rolling sound for each ball
  For b = 0 to UBound(BOT)

  If PFOption = 1 or PFOption = 2 Then
    If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
      rolling(b) = True
      PlaySound("fx_ballrollingA" & b), -1, Vol(BOT(b)) * 0.2 * xGain(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, 0  'Left & Right stereo or Top & Bottom stereo PF Speakers.
    Else
      If rolling(b) = True Then
        StopSound("fx_ballrollingA" & b)
        rolling(b) = False
      End If
    End If

    If BallVel(BOT(b) ) > 1 AND ArchHit =1 Then
      ArchRolling(b) = True
      PlaySound("ArchHitA" & b),   0, (BallVel(BOT(b))/15)^5 * 1 * xGain(BOT(b)), AudioPan(BOT(b)), 0, (BallVel(BOT(b))/40)^7, 1, 0, 0  'Left & Right stereo or Top & Bottom stereo PF Speakers.
      PlaySound("ArchRollA" & b), -1, (BallVel(BOT(b))/30)^5 * 1 * xGain(BOT(b)), AudioPan(BOT(b)), 0, (BallVel(BOT(b))/40)^7, 1, 0, 0  'Left & Right stereo or Top & Bottom stereo PF Speakers.
    Else
      If ArchRolling(b) = True Then
      StopSound("ArchRollA" & b)
      ArchRolling(b) = False
      End If
    End If
  End If

  If PFOption = 3 Then
    If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
      rolling(b) = True
      PlaySound("fx_ballrollingA" & b), -1, Vol(BOT(b)) * 0.2 *    XVol(BOT(b))  *     YVol(BOT(b)), -1, 0, Pitch(BOT(b)), 1, 0, -1 'Top Left PF Speaker
      PlaySound("fx_ballrollingB" & b), -1, Vol(BOT(b)) * 0.2 * (1-XVol(BOT(b))) *     YVol(BOT(b)),  1, 0, Pitch(BOT(b)), 1, 0, -1 'Top Right PF Speaker
      PlaySound("fx_ballrollingC" & b), -1, Vol(BOT(b)) * 0.2 *    XVol(BOT(b))  * (1-YVol(BOT(b))), -1, 0, Pitch(BOT(b)), 1, 0,  1 'Bottom Left PF Speaker
      PlaySound("fx_ballrollingD" & b), -1, Vol(BOT(b)) * 0.2 * (1-XVol(BOT(b))) * (1-YVol(BOT(b))),  1, 0, Pitch(BOT(b)), 1, 0,  1 'Bottom Right PF Speaker
    Else
      If rolling(b) = True Then
        StopSound("fx_ballrollingA" & b)    'Top Left PF Speaker
        StopSound("fx_ballrollingB" & b)    'Top Right PF Speaker
        StopSound("fx_ballrollingC" & b)    'Bottom Left PF Speaker
        StopSound("fx_ballrollingD" & b)    'Bottom Right PF Speaker
        rolling(b) = False
      End If
    End If

    If BallVel(BOT(b) ) > 1 AND ArchHit =1 Then
      ArchRolling(b) = True
      PlaySound("ArchHitA" & b),   0, (BallVel(BOT(b))/20)^5 * 1 *    XVol(BOT(b))  *     YVol(BOT(b)),  -1, 0, (BallVel(BOT(b))/40)^5, 1, 0, -1  'Top Left PF Speaker
      PlaySound("ArchHitB" & b),   0, (BallVel(BOT(b))/20)^5 * 1 * (1-XVol(BOT(b))) *     YVol(BOT(b)),   1, 0, (BallVel(BOT(b))/40)^5, 1, 0, -1  'Top Right PF Speaker
      PlaySound("ArchHitC" & b),   0, (BallVel(BOT(b))/20)^5 * 1 *    XVol(BOT(b))  *  (1-YVol(BOT(b))), -1, 0, (BallVel(BOT(b))/40)^5, 1, 0,  1  'Bottom Left PF Speaker
      PlaySound("ArchHitD" & b),   0, (BallVel(BOT(b))/20)^5 * 1 * (1-XVol(BOT(b))) *  (1-YVol(BOT(b))),  1, 0, (BallVel(BOT(b))/40)^5, 1, 0,  1  'Bottom Right PF Speaker
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
  Next
End Sub

'*************Hit Sound Routines
'Eliminated the Hit Subs extra velocity criteria since the PlayFieldSoundAB command already incorporates the balls velocity.

Sub a_Pins_Hit (idx)
  PlayFieldSoundAB "pinhit_low", 0, 1
End Sub

Sub a_Targets_Hit (idx)
  PlayFieldSoundAB "target", 0, 1
End Sub

Sub a_Metals_Thin_Hit (idx)
  PlayFieldSoundAB "metalhit_thin", 0, 1
End Sub

Sub a_Metals_Medium_Hit (idx)
  PlayFieldSoundAB "metalhit_medium", 0, 1
End Sub

Sub a_Metals2_Hit (idx)
  PlayFieldSoundAB "metalhit2", 0, 1
End Sub

Sub a_Gates_Hit (idx)
  PlayFieldSound "gate4", 0,  EVAL("Gate" & idx), 1
End Sub

Sub a_Rubbers_Hit(idx)
  PlayFieldSoundAB "fx_rubber2", 0, .2
End Sub

Sub RubberWheel_hit
  PlayFieldSoundAB "fx_rubber2", 0, .7
End sub

Sub a_Posts_Hit(idx)
  PlayFieldSoundAB "fx_rubber2", 0, 1
End Sub

Sub LeftFlipper_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RightFlipper_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub LeftFlipper1_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RightFlipper1_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
  Select Case Int(Rnd*3)+1
    Case 1 : PlayFieldSoundAB "flip_hit_1", 0, 1
    Case 2 : PlayFieldSoundAB "flip_hit_2", 0, 1
    Case 3 : PlayFieldSoundAB "flip_hit_3", 0, 1
  End Select
End Sub

Sub ApronL_Hit
  PlayFieldSound "ApronHit", 0, ApronL, 1
End Sub

Sub ApronR_Hit
  PlayFieldSound "ApronHit", 0, ApronR, 1
End Sub

'**********************
' Ball Collision Sound
'**********************

'**************************************
' Explanation of the collision routine
'**************************************

' The collision is built in VP.
' You only need to add a Sub OnBallBallCollision(ball1, ball2, velocity) and when two balls collide they
' will call this routine. What you add in the sub is up to you. As an example is a simple Playsound with volume and paning
' depending of the speed of the collision.

'Added new dBm algorithm to provide a Constant Power "pan" & "fade" to reduce the sound positioning errors inherent in PlaySound's PAN & FADE.
'XVol and YVol drive multiple PlaySound commands panned and faded to their maximum extents where PlaySound's PAN and FADE have the least error.
'Subtracting XVol or YVol from 1 yeilds a inverse response.
'Added the sound files needed to enable calling the same sound twice for stereo and four times for quad under-playfield speakers.
'Examples: fx_collideA through fx_collideD.

Sub OnBallBallCollision(ball1, ball2, velocity)
  If PFOption = 1 or PFOption = 2 Then
    PlaySound "fx_collideA", 0, (Csng(velocity) ^2 / 2000) * xGain(ball1), AudioPan(ball1), 0, Pitch(ball1), 0, 0, 0  'Left & Right stereo or Top & Bottom stereo PF Speakers.
  End If
  If PFOption = 3 Then
    PlaySound "fx_collideA", 0, (Csng(velocity) ^2 / 2000) *    XVol(ball1)  *    YVol(ball1),  -1, 0, Pitch(ball1), 0, 0, -1 'Top Left Playfield Speaker
    PlaySound "fx_collideB", 0, (Csng(velocity) ^2 / 2000) * (1-XVol(ball1)) *    YVol(ball1),   1, 0, Pitch(ball1), 0, 0, -1 'Top Right Playfield Speaker
    PlaySound "fx_collideC", 0, (Csng(velocity) ^2 / 2000) *    XVol(ball1)  * (1-YVol(ball1)), -1, 0, Pitch(ball1), 0, 0,  1 'Bottom Left Playfield Speaker
    PlaySound "fx_collideD", 0, (Csng(velocity) ^2 / 2000) * (1-XVol(ball1)) * (1-YVol(ball1)),  1, 0, Pitch(ball1), 0, 0,  1 'Bottom Right Playfield Speaker
  End If
End Sub

Sub PlayFieldSound (SoundName, Looper, TableObject, VolMult)
'New dBm algorithm to provide a Constant Power "pan" & "fade" to reduce the sound positioning errors inherent in PlaySound's PAN & FADE.
'XVol and YVol drive multiple PlaySound commands panned and faded to their maximum extents where PlaySound's PAN and FADE have the least error.
'Subtracting XVol or YVol from 1 yeilds a inverse response.
'Added the looped sound files needed for to enable calling the same sound twice for stereo and four times for quad under-playfield speakers.
'Examples: BuzzLA through LD and BuzzRA through RD.

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
'New dBm algorithm to provide a Constant Power "pan" & "fade" to reduce the sound positioning errors inherent in PlaySound's PAN & FADE.
'XVol and YVol drive multiple PlaySound commands panned and faded to their maximum extents where PlaySound's PAN and FADE have the least error.
'Subtracting XVol or YVol from 1 yeilds a inverse response.

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


'***************Bubble Sort
Dim TempScore(2), TempPos(3), Position(5)
Dim BSx, BSy
'Scores are sorted high to low with Position being the player's number
Sub SortScores
  For BSx = 1 to 4
    Position(BSx) = BSx
  Next
  For BSx = 1 to 4
    For BSy = 1 to 3
      If Score(BSy) < Score(BSy+1) Then
        TempScore(1) = Score(BSy+1)
        TempPos(1) = Position(BSy+1)
        Score(BSy+1) = Score(BSy)
        Score(BSy) = TempScore(1)
        Position(BSy+1) = Position(BSy)
        Position(BSy) = TempPos(1)
      End If
    Next
  Next

  checkHighScores
' TextBox1.text = Score(1) & " " & Position(1)
' TextBox2.text = Score(2) & " " & Position(2)
' TextBox3.text = Score(3) & " " & Position(3)

End Sub

'*************Check for High Scores

Dim highScore(5), activeScore(5), hs, chX, chY, chZ, chIX, tempI(4), tempI2(4), flag, hsI, hsX
'goes through the 5 high scores one at a time and compares them to the player's scores high to
'if a player's score is higher it marks that postion with ActiveScore(x) and moves all of the other
' high scores down by one along with the high score's player initials
' also clears the new high score's initials for entry later
Sub checkHighScores
  For hs = 1 to 2                   'look at 2 player scores
    For chY = 0 to 4                  'look at all 5 saved high scores
      If score(hs) > highScore(chY) Then
        flag = flag + 1           'flag to show how many high scores needs replacing
        tempScore(1) = highScore(chY)
        highScore(chY) = score(hs)
        activeScore(hs) = chY       'ActiveScore(x) is the high score being modified with x=0 the largest and x=4 the smallest
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
'     TextBox2.text = flag
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
Dim Initial(6,5)

Sub EnterInititals(keycode)
    If KeyCode = LeftFlipperKey Then
      HSx = HSx - 1           'HSx is the inital to be displayed A-Z plus " "
      If HSx < 0 Then HSx = 26
      If HSi < 4 Then EVAL("Initial" & HSi).image = HSiArray(HSx)   'HSi is which of the three intials is being modified
      PlayFieldSound "target", 0, PTape, 1.5
    End If
    If keycode = RightFlipperKey Then
      HSx = HSx + 1
      If HSx > 26 Then HSx = 0
      If HSi < 4 Then EVAL("Initial"& HSi).image = HSiArray(HSx)
      PlayFieldSound "target", 0, PTape, 1.5
    End If
    If keycode = StartGameKey Then
      If HSi < 3 Then                 'if not on the last initial move on to the next intial
        EVAL("Initial" & HSi).image = HSiArray(HSx) 'display the initial
        Initial(ActiveScore(Flag), HSi) = HSx   'save the inital
        EVAL("InitialTimer" & HSi).enabled = 0    'turn that inital's timer off
        EVAL("Initial" & HSi).visible = 1     'make the initial not flash but be turn on
        Initial(ActiveScore(Flag),HSi + 1) = HSx  'move to the next initial and make it the same as the last initial
        EVAL("Initial" & HSi +1).image = HSiArray(HSx)  'display this intial
        EVAL("InitialTimer" & HSi + 1).enabled = 1  'make the new intial flash
        HSi = HSi + 1
        PlayfieldSound "MetalHit2", 0, PTape, 1.5             'increment HSi
      Else                    'if on the last initial then get ready yo exit the subroutine
        Initial3.visible = 1          'make the intial visible
        InitialTimer3.enabled = 0       'shut off the flashing
        Initial(ActiveScore(Flag),3) = HSx    'set last initial
        InitialEntry              'exit subroutine
        PlayFieldSound "Bell10", 0, SoundPoint13, 1
      End If
    End If
End Sub

'************Update Initials and see if more scores need to be updated
Dim eIX, initialsDone
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
    Credit = cdbl (temp(26))
    FreePlay = cdbl (temp(27))
    Balls = cdbl (temp(28))
    MatchNumber = cdbl (temp(29))
    Chime = cdbl (temp(30))
    Score(1) = cdbl (temp(31))
    Score(2) = cdbl (temp(32))
    PFOption = cdbl (temp(33))
    LayoutDifficulty = cdbl (temp(34))
    RotAngle = cdbl (temp(35))
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
    ScoreFile.WriteLine Score(1)
    ScoreFile.WriteLine Score(2)
    ScoreFile.WriteLine PFOption
    ScoreFile.WriteLine LayoutDifficulty
    ScoreFile.WriteLine RotAngle
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


'***************Static Post It Note Update
Dim  HSy

Sub UpdatePostIt
  ScoreMil = Int(HighScore(0)/1000000)
  Score100K = Int( (HighScore(0) - (ScoreMil*1000000) ) / 100000)
  Score10K = Int( (HighScore(0) - (ScoreMil*1000000) - (Score100K*100000) ) / 10000)
  ScoreK = Int( (HighScore(0) - (ScoreMil*1000000) - (Score100K*100000) - (Score10K*10000) ) / 1000)
  Score100 = Int( (HighScore(0) - (ScoreMil*1000000) - (Score100K*100000) - (Score10K*10000) - (ScoreK*1000) ) / 100)
  Score10 = Int( (HighScore(0) - (ScoreMil*1000000) - (Score100K*100000) - (Score10K*10000) - (ScoreK*1000) - (Score100*100) ) / 10)
  ScoreUnit = (HighScore(0) - (ScoreMil*1000000) - (Score100K*100000) - (Score10K*10000) - (ScoreK*1000) - (Score100*100) - (Score10*10) )

  Pscore6.image = HSArray(ScoreMil):If HighScore(0) < 1000000 Then PScore6.image = HSArray(10)
  Pscore5.image = HSArray(Score100K):If HighScore(0) < 100000 Then PScore5.image = HSArray(10)
  PScore4.image = HSArray(Score10K):If HighScore(0) < 10000 Then PScore4.image = HSArray(10)
  PScore3.image = HSArray(ScoreK):If HighScore(0) < 1000 Then PScore3.image = HSArray(10)
  PScore2.image = HSArray(Score100):If HighScore(0) < 100 Then PScore2.image = HSArray(10)
  PScore1.image = HSArray(Score10):If HighScore(0) < 10 Then PScore1.image = HSArray(10)
  PScore0.image = HSArray(ScoreUnit):If HighScore(0) < 1 Then PScore0.image = HSArray(10)
  If HighScore(0) < 1000 Then
    PComma.image = HSArray(10)
  Else
    PComma.image = HSArray(11)
  End If
  If HighScore(0) < 1000000 Then
    PComma1.image = HSArray(10)
  Else
    PComma1.image = HSArray(11)
  End If
  If HighScore(0) > 999999 Then Shift = 0 :PComma.transx = 0
  If HighScore(0) < 1000000 Then Shift = 1:PComma.transx = -10
  If HighScore(0) < 100000 Then Shift = 2:PComma.transx = -20
  If HighScore(0) < 10000 Then Shift = 3:PComma.transx = -30
  For HSy = 0 to 6
    EVAL("Pscore" & HSy).transx = (-10 * Shift)
  Next
  Initial1.image = HSiArray(Initial(0,1))
  Initial2.image = HSiArray(Initial(0,2))
  Initial3.image = HSiArray(Initial(0,3))
End Sub

'***************Show Current Score

Sub ShowScore
  ScoreMil = Int(HighScore(ActiveScore(Flag))/1000000)
  Score100K = Int( (HighScore(ActiveScore(Flag)) - (ScoreMil*1000000) ) / 100000)
  Score10K = Int( (HighScore(ActiveScore(Flag)) - (ScoreMil*1000000) - (Score100K*100000) ) / 10000)
  ScoreK = Int( (HighScore(ActiveScore(Flag)) - (ScoreMil*1000000) - (Score100K*100000) - (Score10K*10000) ) / 1000)
  Score100 = Int( (HighScore(ActiveScore(Flag)) - (ScoreMil*1000000) - (Score100K*100000) - (Score10K*10000) - (ScoreK*1000) ) / 100)
  Score10 = Int( (HighScore(ActiveScore(Flag)) - (ScoreMil*1000000) - (Score100K*100000) - (Score10K*10000) - (ScoreK*1000) - (Score100*100) ) / 10)
  ScoreUnit = (HighScore(ActiveScore(Flag)) - (ScoreMil*1000000) - (Score100K*100000) - (Score10K*10000) - (ScoreK*1000) - (Score100*100) - (Score10*10) )

  Pscore6.image = HSArray(ScoreMil):If HighScore(ActiveScore(Flag)) < 1000000 Then PScore6.image = HSArray(10)
  Pscore5.image = HSArray(Score100K):If HighScore(ActiveScore(Flag)) < 100000 Then PScore5.image = HSArray(10)
  PScore4.image = HSArray(Score10K):If HighScore(ActiveScore(Flag)) < 10000 Then PScore4.image = HSArray(10)
  PScore3.image = HSArray(ScoreK):If HighScore(ActiveScore(Flag)) < 1000 Then PScore3.image = HSArray(10)
  PScore2.image = HSArray(Score100):If HighScore(ActiveScore(Flag)) < 100 Then PScore2.image = HSArray(10)
  PScore1.image = HSArray(Score10):If HighScore(ActiveScore(Flag)) < 10 Then PScore1.image = HSArray(10)
  PScore0.image = HSArray(ScoreUnit):If HighScore(ActiveScore(Flag)) < 1 Then PScore0.image = HSArray(10)
  If HighScore(ActiveScore(Flag)) < 1000 Then
    PComma.image = HSArray(10)
  Else
    PComma.image = HSArray(11)
  End If
  If HighScore(ActiveScore(Flag)) < 1000000 Then
    PComma1.image = HSArray(10)
  Else
    PComma1.image = HSArray(11)
  End If
  If HighScore(Flag) > 999999 Then Shift = 0 :PComma.transx = 0
  If HighScore(ActiveScore(Flag)) < 1000000 Then Shift = 1:PComma.transx = -10
  If HighScore(ActiveScore(Flag)) < 100000 Then Shift = 2:PComma.transx = -20
  If HighScore(ActiveScore(Flag)) < 10000 Then Shift = 3:PComma.transx = -30
  For HSy = 0 to 6
    EVAL("Pscore" & HSy).transx = (-10 * Shift)
  Next
  Initial1.image = HSiArray(Initial(ActiveScore(Flag),1))
  Initial2.image = HSiArray(Initial(ActiveScore(Flag),2))
  Initial3.image = HSiArray(Initial(ActiveScore(Flag),3))
End Sub


'***************Dynamic Post It Note Update
Dim ScoreUpdate, DHSx

Sub DynamicUpdatePostIt_Timer
  ScoreMil = Int(HighScore(ScoreUpdate)/1000000)
  Score100K = Int( (HighScore(ScoreUpdate) - (ScoreMil*1000000) ) / 100000)
  Score10K = Int( (HighScore(ScoreUpdate) - (ScoreMil*1000000) - (Score100K*100000) ) / 10000)
  ScoreK = Int( (HighScore(ScoreUpdate) - (ScoreMil*1000000) - (Score100K*100000) - (Score10K*10000) ) / 1000)
  Score100 = Int( (HighScore(ScoreUpdate) - (ScoreMil*1000000) - (Score100K*100000) - (Score10K*10000) - (ScoreK*1000) ) / 100)
  Score10 = Int( (HighScore(ScoreUpdate) - (ScoreMil*1000000) - (Score100K*100000) - (Score10K*10000) - (ScoreK*1000) - (Score100*100) ) / 10)
  ScoreUnit = (HighScore(ScoreUpdate) - (ScoreMil*1000000) - (Score100K*100000) - (Score10K*10000) - (ScoreK*1000) - (Score100*100) - (Score10*10) )

  Pscore6.image = HSArray(ScoreMil):If HighScore(ScoreUpdate) < 1000000 Then PScore6.image = HSArray(10)
  Pscore5.image = HSArray(Score100K):If HighScore(ScoreUpdate) < 100000 Then PScore5.image = HSArray(10)
  PScore4.image = HSArray(Score10K):If HighScore(ScoreUpdate) < 10000 Then PScore4.image = HSArray(10)
  PScore3.image = HSArray(ScoreK):If HighScore(ScoreUpdate) < 1000 Then PScore3.image = HSArray(10)
  PScore2.image = HSArray(Score100):If HighScore(ScoreUpdate) < 100 Then PScore2.image = HSArray(10)
  PScore1.image = HSArray(Score10):If HighScore(ScoreUpdate) < 10 Then PScore1.image = HSArray(10)
  PScore0.image = HSArray(ScoreUnit):If HighScore(ScoreUpdate) < 1 Then PScore0.image = HSArray(10)
  If HighScore(ScoreUpdate) < 1000 Then
    PComma.image = HSArray(10)
  Else
    PComma.image = HSArray(11)
  End If
  If HighScore(ScoreUpdate) < 1000000 Then
    PComma1.image = HSArray(10)
  Else
    PComma1.image = HSArray(11)
  End If
  If HighScore(ScoreUpdate) > 999999 Then Shift = 0 :PComma.transx = 0
  If HighScore(ScoreUpdate) < 1000000 Then Shift = 1:PComma.transx = -10
  If HighScore(ScoreUpdate) < 100000 Then Shift = 2:PComma.transx = -20
  If HighScore(ScoreUpdate) < 10000 Then Shift = 3:PComma.transx = -30
  For DHSx = 0 to 6
    EVAL("Pscore" & DHSx).transx = (-10 * Shift)
  Next
  Initial1.image = HSiArray(Initial(ScoreUpdate,1))
  Initial2.image = HSiArray(Initial(ScoreUpdate,2))
  Initial3.image = HSiArray(Initial(ScoreUpdate,3))
  ScoreUpdate = ScoreUpdate + 1
  If ScoreUpdate = 5 then ScoreUpdate = 0
End Sub

Sub LightsRandom_Timer()
  Select Case Int(Rnd*2)+1
    Case 1 : DOF 200, 1
    Case 2 : DOF 200, 0
  End Select
  Select Case Int(Rnd*2)+1
    Case 1 : DOF 201, 1
    Case 2 : DOF 201, 0
  End Select
  Select Case Int(Rnd*2)+1
    Case 1 : DOF 202, 1
    Case 2 : DOF 202, 0
  End Select
End Sub

'***************Exit Table
Sub Table1_Exit()
  TurnOff
  If B2SOn Then Controller.stop
End Sub

'******************************************************
'   TRACK ALL BALL VELOCITIES
'     FOR RUBBER DAMPENER AND DROP TARGETS
'******************************************************

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

Sub RDampen_Timer()
  Cor.Update
End Sub

'******************************************************
'     FLIPPER CORRECTION FUNCTIONS
'******************************************************

Sub AddPt(aStr, idx, aX, aY)  'debugger wrapper for adjusting flipper script in-game
  dim a : a = Array(LF, RF)
  dim x : for each x in a
    x.addpoint aStr, idx, aX, aY
  Next
End Sub

Class FlipperPolarity
  Public DebugOn, Enabled
  Private FlipAt  'Timer variable (IE 'flip at 723,530ms...)
  Public TimeDelay  'delay before trigger turns off and polarity is disabled TODO set time!
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

  Public Sub Report(aChooseArray)   'debug, reports all coords in tbPL.text
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
  Private Function FlipperOn() : if gameTime < FlipAt+TimeDelay then FlipperOn = True : End If : End Function 'Timer shutoff for polaritycorrect

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
          if ballpos > 0.65 then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)        'find safety coefficient 'ycoef' data
        end if
      Next

      If BallPos = 0 Then 'no ball data meaning the ball is entering and exiting pretty close to the same position, use current values.
        BallPos = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
        if ballpos > 0.65 then  Ycoef = LinearEnvelope(aBall.Y, YcoefIn, YcoefOut)            'find safety coefficient 'ycoef' data
      End If

      'Velocity correction
      if not IsEmpty(VelocityIn(0) ) then
        Dim VelCoef
   :      VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)

        if partialflipcoef < 1 then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)

        if Enabled then aBall.Velx = aBall.Velx*VelCoef
        if Enabled then aBall.Vely = aBall.Vely*VelCoef
      End If

      'Polarity Correction (optional now)
      if not IsEmpty(PolarityIn(0) ) then
        If StartPoint > EndPoint then LR = -1 'Reverse polarity if left flipper
        dim AddX : AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR

        if Enabled then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef)
        'playsound "knocker"
      End If
    End If
    RemoveBall aBall
  End Sub
End Class

'******************************************************
'   FLIPPER POLARITY AND RUBBER DAMPENER
'     SUPPORTING FUNCTIONS
'******************************************************

' Used for flipper correction and rubber dampeners
Sub ShuffleArray(ByRef aArray, byVal offset) 'shuffle 1d array
  dim x, aCount : aCount = 0
  redim a(uBound(aArray) )
  for x = 0 to uBound(aArray) 'Shuffle objects in a temp array
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
  redim aArray(aCount-1+offset) 'Resize original array
  for x = 0 to aCount-1   'set objects back into original array
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
Function PSlope(Input, X1, Y1, X2, Y2)  'Set up line via two points, no clamping. Input X, output Y
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
  dim ii : for ii = 1 to uBound(xKeyFrame)  'find active line
    if xInput <= xKeyFrame(ii) then L = ii : exit for : end if
  Next
  if xInput > xKeyFrame(uBound(xKeyFrame) ) then L = uBound(xKeyFrame)  'catch line overrun
  Y = pSlope(xInput, xKeyFrame(L-1), yLvl(L-1), xKeyFrame(L), yLvl(L) )

  if xInput <= xKeyFrame(lBound(xKeyFrame) ) then Y = yLvl(lBound(xKeyFrame) )  'Clamp lower
  if xInput >= xKeyFrame(uBound(xKeyFrame) ) then Y = yLvl(uBound(xKeyFrame) )  'Clamp upper

  LinearEnvelope = Y
End Function

' Used for drop targets and flipper tricks
Function Distance(ax,ay,bx,by)
  Distance = SQR((ax - bx)^2 + (ay - by)^2)
End Function

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


dim RubbersD : Set RubbersD = new Dampener  'frubber
RubbersD.name = "Rubbers"
RubbersD.debugOn = False  'shows info in textbox "TBPout"
RubbersD.Print = False  'debug, reports in debugger (in vel, out cor)
'cor bounce curve (linear)
'for best results, try to match in-game velocity as closely as possible to the desired curve
'RubbersD.addpoint 0, 0, 0.935  'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 0, 0, 0.96  'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 1, 3.77, 0.96
RubbersD.addpoint 2, 5.76, 0.967  'dont take this as gospel. if you can data mine rubber elasticitiy, please help!
RubbersD.addpoint 3, 15.84, 0.874
RubbersD.addpoint 4, 56, 0.64 'there's clamping so interpolate up to 56 at least

dim SleevesD : Set SleevesD = new Dampener  'this is just rubber but cut down to 85%...
SleevesD.name = "Sleeves"
SleevesD.debugOn = False  'shows info in textbox "TBPout"
SleevesD.Print = False  'debug, reports in debugger (in vel, out cor)
SleevesD.CopyCoef RubbersD, 0.85

Class Dampener
  Public Print, debugOn 'tbpOut.text
  public name, Threshold  'Minimum threshold. Useful for Flippers, which don't have a hit threshold.
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
    RealCOR = BallSpeed(aBall) / cor.ballvel(aBall.id)
    coef = desiredcor / realcor
    if debugOn then str = name & " in vel:" & round(cor.ballvel(aBall.id),2 ) & vbnewline & "desired cor: " & round(desiredcor,4) & vbnewline & _
    "actual cor: " & round(realCOR,4) & vbnewline & "ballspeed coef: " & round(coef, 3) & vbnewline
    if Print then debug.print Round(cor.ballvel(aBall.id),2) & ", " & round(desiredcor,3)

' Thalamus - patched : ' Thalamus - patched :     aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
    aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
' Thalamus - patched :     aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
    aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
    if debugOn then TBPout.text = str
  End Sub

  Public Sub CopyCoef(aObj, aCoef) 'alternative addpoints, copy with coef
    dim x : for x = 0 to uBound(aObj.ModIn)
      addpoint x, aObj.ModIn(x), aObj.ModOut(x)*aCoef
    Next
  End Sub


  Public Sub Report()   'debug, reports all coords in tbPL.text
    if not debugOn then exit sub
    dim a1, a2 : a1 = ModIn : a2 = ModOut
    dim str, x : for x = 0 to uBound(a1) : str = str & x & ": " & round(a1(x),4) & ", " & round(a2(x),4) & vbnewline : next
    TBPout.text = str
  End Sub


End Class

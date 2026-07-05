'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~                                      ~~~
'~~~                Hunter                ~~~
'~~~            (Jennings 1935)           ~~~
'~~~                                      ~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Made by Warbler33 (VPF) / CactusDude (VPU)
'Released December 2020
'Re-released 2021
'Version 1.08
'
' v1.07 by Warbler33 (other names : CactusDude)
' Pinball size adjusted (thanks for the suggestion, Wylte) along with necessary edits to go along with this change.
' I also some minor tweaks to the artwork and added a non-tarcisio style wheel to the 1.07 zip folder.
'
' v1.08 by 7he S4ge (other names : jejepinball, jejevpuniverse)
' Add textboxs for the desktop version
' Better backglass (.directb2s file) for display images of animals (rabbit, pheasant...) and score
' Improve hits and checks : merge all in one function
' LeftMagnaSave : cheat, active hits one by one
' Animation on B2S of animals images at start of the game
'
'
' Keys :
' LeftMagnaSave : cheat, active hits one by one
'

Option Explicit

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the Controller.vbs file in order to run this table (installed with the VPX package in the scripts folder)"
On Error Goto 0

Const cGameName = "Hunter"


'~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~  Configuration  ~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~
Const BallsperGame = 10
Const FreePlay = True
Const EnableAnimationOnB2S = True ' If true we animate the animals images on B2S at start of the game
Const RollingSoundFactor = 2.5    'set sound level factor here for Ball Rolling Sound, 1=default level

Const VolumeMultiplier = 3        'adjusts table sound volume

Const ResetHighscore = False    'enable to manually reset the Highscores

Const BallSize = 20  'Ball radius

Dim i,HighScore,Credits,GameStarted,BallinPlay,BallReleaseLocked,BallsonBoard
Dim LaneClear

'--------------------------
'------  Table Init  ------
'--------------------------

Sub Hunter_Init
  'pause for display B2S on desktop on top of the playfield
  'Dim PauseTime, Start
  'PauseTime = 1    ' Set duration.
  'Start = Timer    ' Set start time.
  'Do While Timer < Start + PauseTime
  'Loop
  LaunchB2S
  FINISHGAME = False
  LoadHighScore
  If ResetHighScore Then
    SetDefaultHSTD
  end If
  UpdatePostIt

  Randomize

  if FreePlay then
    Credits = 5
    DOF 126, DOFOn
  end If

' TiltKicker.CreateBall
  TiltCount = 0
  BallinPlay = 0
  GameStarted = False
'If BallsonBoard=10 then BallReleaseLocked=True End if
  UpdateScoreReel.enabled = True
  LaneClear = True
End Sub


'~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~  Game Features  ~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~

'Sub MainTimer_Timer
' if (BallsonBoard = BallsperGame) and (not BallReleaseLocked) Then
'   If BallVel(LBall) = 0 Then
'     GameEndTimer.enabled = True
'   Else
'     if GameEndTimer.enabled Then
'       GameEndTimer.enabled = False
'     end If
'   end If
' end If
'End Sub

Sub StartGame
  if Credits > 0 Then
    if not GameStarted then
      GameEndTimer.interval = 2000
      BallReleaseLocked = False
      Playerscore = 0
      ScoreText.Text = "Score : 0"
      CheatStep = 0
      DuckSquirrelTextBox.Text = ""
      RabbitTextBox.Text = ""
      PheasantTextBox.Text = ""
      PartridgeTextBox.Text = ""
      Oldscore = 0
      Val10 = 0
      Val100 = 0
      Val1000 = 0
      Val10000 = 0
      Val100000 = 0
'     Special1 = False
'     Special2 = False
'     Special3 = False
      UpdateScoreReel.enabled = True
      GameStarted = True
      BallinPlay = 0
'     P_BallinPlay.image = CStr(0)
'     P_BallinPlay1.image = CStr(0)
      Credits = Credits - 1
      If Credits < 1 Then DOF 126, DOFOff
      if FreePlay and (Credits = 0) then
        Credits = 5
        DOF 126, DOFOn
      end If
      ' Launch animation on B2S
      If EnableAnimationOnB2S Then
        AnimationOnB2S
      End If
    end If
  end If
End Sub

Sub NextBall
  If LaneClear = False then Exit Sub end if
  if BallsonBoard <= BallsperGame Then
    BallReleaseLocked = False
    PlaySound "BallLift",0,0.5,Pan(Plunger),0.25,0,0,1,0
'   BallinPlay = BallinPlay + 1
    end If
    BallRelease.Timerenabled = True
' End If
End Sub

Sub BallRelease_Timer
  BallRelease.Timerenabled = False
  BallRelease.CreateSizedBall(20)
  BallRelease.Kick 270, 7
  DOF 122, DOFPulse
End Sub

Sub CheckFinish
  If FINISHGAME = False then Exit Sub end If
  CheckLastBall_Timer
End Sub

Sub CheckLastBall_Timer
  If BallsonBoard >= 10 then LastBall.Enabled = True
End Sub

Sub LastBall_Timer
  LastBall.Enabled = False
  EndGame
End Sub

'Sub GameEndTimer_Timer
' GameEndTimer.Enabled = False
' EndGame
'End Sub

Sub EndGame
  CheckNewHighScorePostIt(Playerscore)
  PlaySound "Chime"
  GameStarted = False
  BallinPlay = 0
End Sub

'--------------------------------
'------  Helper Functions  ------
'--------------------------------
Sub PostWall_Hit
  P_Left1.RotY = 185
  P_Left2.RotY = 185
  P_Left3.RotY = 185
  PostWall.timerenabled = True
End Sub

Sub PostWall_Timer
  PostWall.timerenabled = False
  P_Left1.RotY = 180
  P_Left2.RotY = 180
  P_Left3.RotY = 180
End Sub

Sub Drain_Hit()
' PlaySound "drain3",0,0.2,0,0.25
' DOF 121, DOFPulse
  Drain.DestroyBall
End Sub

Dim BOT2, LBall, FINISHGAME
Sub EnterPF_Hit
  BallsonBoard = BallsonBoard + 1
  If BallsonBoard = 0 then BallReleaseLocked = False end if
  If BallsonBoard = 1 then BallReleaseLocked = False end if
  If BallsonBoard = 2 then BallReleaseLocked = False end if
  If BallsonBoard = 3 then BallReleaseLocked = False end if
  If BallsonBoard = 4 then BallReleaseLocked = False end if
  If BallsonBoard = 5 then BallReleaseLocked = False end if
  If BallsonBoard = 6 then BallReleaseLocked = False end if
  If BallsonBoard = 7 then BallReleaseLocked = False end if
  If BallsonBoard = 8 then BallReleaseLocked = False end if
  If BallsonBoard = 9 then BallReleaseLocked = False end if
  If BallsonBoard = 10 then BallReleaseLocked = True end if
  If BallsonBoard >= 10 then FINISHGAME = True end If
  CheckFinish
' BallReleaseLocked = False
  'Set LBall = ActiveBall
End Sub

Dim BoardDir
Sub MoveBoardTimer_Timer

  P_Board.Transy = P_Board.Transy + BoardDir
  if P_Board.Transy < -40 Then
    KickerHoleWall.isdropped = True
  Else
    KickerHoleWall.isdropped = False
  End If

  if P_Board.Transy < -57 Then
    BoardDir = -Boarddir
    HoldBoardTimer.enabled = True
    MoveBoardTimer.enabled = False
  end If

  if P_Board.Transy >= 0 Then
    P_Board.Transy = 0
    MoveBoardTimer.enabled = False
    StartGame
  end If
End Sub

Sub HoldBoardTimer_Timer
  HoldBoardTimer.enabled = False
  MoveBoardTimer.enabled = True
  PlaySound "BallIn",0,0.75,0,0.25,0,0,1,0
end Sub

Dim Val10,Val100,Val1000,Val10000,Val100000,Score10,Score100,Score1000,Score10000,Score100000,TempScore,Oldscore
Dim PlayerScore,Special1,Special2,Special3
PlayerScore=0

'-------------------------------
'------  Keybord Handler  ------
'-------------------------------
Dim CheatStep:CheatStep = 0
Sub Hunter_KeyDown(ByVal keycode)
  If keycode = StartGameKey Then
    if BallsonBoard>9 or BallsonBoard=0 then
      If LastBall.Enabled = False then
        BallsonBoard = 0
        Ballinplay = 0
        BoardDir = -1
        MoveBoardTimer.enabled = True
        PlaySound "CoinStart",0,0.75,0,0.25,0,0,1,0
        PlayerScore=0
        ClearKickers
      end if
    end If
  End if

  If keycode = PlungerKey Then
    Plunger.PullBack
'   PlaySound "plungerpull",0,1,Pan(Plunger),0.25,0,0,1,0
  End If

  ' Cheats : active hits one by one
  If keycode = LeftMagnaSave Then
    Select Case CheatStep
      Case 0:SkillShotHit = True
      Case 1:DuckHitA = True
      Case 2:DuckHitB = True
      Case 3:SquirrelHitA = True
      Case 4:SquirrelHitB = True
      Case 5:RabbitHitA = True
      Case 6:RabbitHitB = True
      Case 7:RabbitHitC = True
      Case 8:PhesantHitA = True
      Case 9:PhesantHitB = True
      Case 10:PhesantHitC = True
      Case 11:PartridgeHitA = True
      Case 12:PartridgeHitB = True
      Case 13:PartridgeHitC = True
      Case Else : CheatStep = -1
    End Select
    HitAndWinCheck
    CheatStep = CheatStep + 1
  End If

  If keycode = RightFlipperKey or keycode = RightMagnaSave and LaneClear = True then
      if GameStarted and BallReleaseLocked = False Then
      NextBall
  '   PlaySound "ballrelease"
    end If
  End If

  If keycode = LeftTiltKey Then
    Nudge 90, 2
'   TiltIt
  End If

  If keycode = RightTiltKey Then
    Nudge 270, 2
'   TiltIt
  End If

  If keycode = CenterTiltKey OR keycode = 57 Then
    Nudge 0, 2
    StuckBallSweep
'   TiltIt
  End If

End Sub

Sub Hunter_KeyUp(ByVal keycode)
  If keycode = PlungerKey Then
    Plunger.Fire
    PlaySound "PlungerHit",0,1,Pan(Plunger),0.25,0,0,1,0
  End If

  If keycode = LeftFlipperKey Then

  End If

  If keycode = RightFlipperKey Then

  End If

End Sub

' Hide texboxs in cabinet mode
If Hunter.ShowDT = false then
    Scoretext.Visible = false
  DuckSquirrelTextBox.Visible = false
  RabbitTextBox.Visible = false
  PheasantTextBox.Visible = false
  PartridgeTextBox.Visible = false
End If

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'                  Bumper Spring Movement
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Sub SpringWall_Hit 'Not doing anything - oh well
  Spring.transX = +10000
  SpringReset.enabled = True
End Sub

Sub SpringReset_Timer()
  SpringReset.enabled = False
  Spring.transX = -10000
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'                  Lane Check
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Sub ShooterLaneLaunch_Hit
  LaneClear = False
End Sub

Sub Gate001_Hit
  LaneClear = True
End Sub


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'                  Game/Backglass Reseting
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


Sub StuckBallSweep
  If BallsonBoard = 0 then ClearKickers
  PlayerScore = 0
end Sub

Sub ClearKickers
      K100A.DestroyBall
      K100B.DestroyBall
      K100C.DestroyBall
      K100D.DestroyBall
      K200A.DestroyBall
      K200B.DestroyBall
      K200C.DestroyBall
      K200D.DestroyBall
      K200E.DestroyBall
      K200F.DestroyBall
      K300A.DestroyBall
      K300B.DestroyBall
      K300C.DestroyBall
      K300D.DestroyBall
      K400A.DestroyBall
      K400B.DestroyBall
      KPartridgeA.DestroyBall
      KPartridgeB.DestroyBall
      KPartridgeC.DestroyBall
      KPhesantA.DestroyBall
      KPhesantB.DestroyBall
      KPhesantC.DestroyBall
      KRabbitA.DestroyBall
      KRabbitB.DestroyBall
      KRabbitC.DestroyBall
      KDuckA.DestroyBall
      KDuckB.DestroyBall
      KSquirrelA.DestroyBall
      KSquirrelB.DestroyBall
      KSkillShot.DestroyBall
      Controller.B2SSetScorePlayer1 0
      Controller.B2SSetData 50,0 'Images on reel
      Controller.B2SSetData 51,0
      Controller.B2SSetData 52,0
      Controller.B2SSetData 53,0
      Controller.B2SSetData 101,0
      Controller.B2SSetData 102,0
      Controller.B2SSetData 103,0
      Controller.B2SSetData 104,0
      Controller.B2SSetData 105,0
      Controller.B2SSetData 106,0
      Controller.B2SSetData 107,0
      Controller.B2SSetData 108,0
      Controller.B2SSetData 109,0
      Controller.B2SSetData 110,0
      Controller.B2SSetData 111,0
      Controller.B2SSetData 112,0
      Controller.B2SSetData 113,0
      Controller.B2SSetData 114,0
      Controller.B2SSetData 115,0
      Controller.B2SSetData 116,0
      Controller.B2SSetData 117,0
      Controller.B2SSetData 118,0
      BallsonBoard = 0
      BallReleaseLocked = False
      ResetWins
      UpdateScoreReel.enabled = True
end Sub

Sub ResetWins
  SkillShotWin = False
  PartridgeWin = False
  PhesantWin = False
  RabbitWin = False
  SquirrelWin = False
  DuckWin = False
  DuckSquirrelTextBox.Text = ""
  RabbitTextBox.Text = ""
  PheasantTextBox.Text = ""
  PartridgeTextBox.Text = ""
'And Hits
  SkillShotHit = False
  PartridgeHitA = False
  PartridgeHitB = False
  PartridgeHitC = False
  PhesantHitA = False
  PhesantHitB = False
  PhesantHitC = False
  RabbitHitA = False
  RabbitHitB = False
  RabbitHitC = False
  SquirrelHitA = False
  SquirrelHitB = False
  DuckHitA = False
  DuckHitB = False
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'                  Backglass/Scoring
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Dim Controller

Sub LaunchB2S
  Set Controller = CreateObject("B2S.Server")
  Controller.B2SName = "Hunter"
  Controller.Run
End Sub

'Skill Shot
Dim SkillShotHit, SkillShotWin
Sub KSkillShot_Hit
  SkillShotHit = True
  Addpoints(2000)
  HitAndWinCheck
End sub

Sub KSkillShot_unHit
  SkillShotHit = False
  Addpoints(-2000)
  HitAndWinCheck
End Sub

' Check hits, enable wins and display results on B2S (lamps and animal images)
Sub HitAndWinCheck
  ' Hits lamps
  If SkillShotHit=True Then
    Controller.B2SSetData 108,1 'lamp
  Else
    Controller.B2SSetData 108,0
  End If
  If DuckHitA=True Then
    Controller.B2SSetData 101,1 'lamp
  Else
    Controller.B2SSetData 101,0
  End If
  If DuckHitB=True Then
    Controller.B2SSetData 102,1 'lamp
  Else
    Controller.B2SSetData 102,0
  End If
  If SquirrelHitA=True Then
    Controller.B2SSetData 103,1 'lamp
  Else
    Controller.B2SSetData 103,0
  End If
  If SquirrelHitB=True Then
    Controller.B2SSetData 104,1 'lamp
  Else
    Controller.B2SSetData 104,0
  End If
  If RabbitHitA=True Then
    Controller.B2SSetData 105,1 'lamp
  Else
    Controller.B2SSetData 105,0
  End If
  If RabbitHitB=True Then
    Controller.B2SSetData 106,1 'lamp
  Else
    Controller.B2SSetData 106,0
  End If
  If RabbitHitC=True Then
    Controller.B2SSetData 107,1 'lamp
  Else
    Controller.B2SSetData 107,0
  End If
  If PhesantHitA=True Then
    Controller.B2SSetData 109,1 'lamp
  Else
    Controller.B2SSetData 109,0
  End If
  If PhesantHitB=True Then
    Controller.B2SSetData 110,1 'lamp
  Else
    Controller.B2SSetData 110,0
  End If
  If PhesantHitC=True Then
    Controller.B2SSetData 111,1 'lamp
  Else
    Controller.B2SSetData 111,0
  End If
  If PartridgeHitA=True Then
    Controller.B2SSetData 112,1 'lamp
  Else
    Controller.B2SSetData 112,0
  End If
  If PartridgeHitB=True Then
    Controller.B2SSetData 113,1 'lamp
  Else
    Controller.B2SSetData 113,0
  End If
  If PartridgeHitC=True Then
    Controller.B2SSetData 114,1 'lamp
  Else
    Controller.B2SSetData 114,0
  End If

  ' Wins lamps and animal images
  If SkillShotHit=True Then
    SkillShotWin=True
    Controller.B2SSetData 108,1 'lamp
  Else
    SkillShotWin=False
    Controller.B2SSetData 108,0
  End If
  If DuckHitA=True and DuckHitB=True and SkillShotWin = True Then
    DuckWin=True
  Else
    DuckWin=False
  End If
  If SquirrelHitA=True and SquirrelHitB=True and SkillShotWin = True Then
    SquirrelWin=True
  Else
    SquirrelWin=False
  End If
  If DuckWin=True Or SquirrelWin=True Then
    Controller.B2SSetData 50,1 'animal image
    Controller.B2SSetData 115,1 'lamp
  Else
    Controller.B2SSetData 50,0 'animal image
    Controller.B2SSetData 115,0 'lamp
  End If
  If RabbitHitA=True and RabbitHitB=True and RabbitHitC=True and SkillShotWin = True Then
    RabbitWin=True
    Controller.B2SSetData 51,1 'animal image
    Controller.B2SSetData 116,1 'lamp
  Else
    RabbitWin=False
    Controller.B2SSetData 51,0 'animal image
    Controller.B2SSetData 116,0 'lamp
  End If
  If PhesantHitA=True and PhesantHitB=True and PhesantHitC=True and SkillShotWin = True Then
    PhesantWin=True
    Controller.B2SSetData 52,1 'animal image
    Controller.B2SSetData 117,1 'lamp
  Else
    PhesantWin=False
    Controller.B2SSetData 52,0 'animal image
    Controller.B2SSetData 117,0 'lamp
  End If
  If PartridgeHitA=True and PartridgeHitB=True and PartridgeHitC=True and SkillShotWin = True Then
    PartridgeWin=True
    Controller.B2SSetData 53,1 'animal image
    Controller.B2SSetData 118,1 'lamp
  Else
    PartridgeWin=False
    Controller.B2SSetData 53,0 'animal image
    Controller.B2SSetData 118,0 'lamp
  End If
End Sub

'Partridge
Dim PartridgeHitA,PartridgeHitB,PartridgeHitC,PartridgeWin
Sub KPartridgeA_Hit
  PartridgeHitA=True
  Addpoints(1500)
  HitAndWinCheck
End sub
Sub KPartridgeB_Hit
  PartridgeHitB=True
  Addpoints (1500)
  HitAndWinCheck
End sub
Sub KPartridgeC_Hit
  PartridgeHitC=True
  Addpoints(1500)
  HitAndWinCheck
End sub
Sub KPartridgeA_unHit
  Addpoints(-1500)
  PartridgeHitA = False
  HitAndWinCheck
End Sub

Sub KPartridgeB_unHit
  Addpoints(-1500)
  PartridgeHitB = False
  HitAndWinCheck
End Sub

Sub KPartridgeC_unHit
  Addpoints(-1500)
  PartridgeHitC = False
  HitAndWinCheck
End Sub

'Phesant
Dim PhesantHitA,PhesantHitB,PhesantHitC,PhesantWin
Sub KPhesantA_Hit
  PhesantHitA=True
  Addpoints(1000)
  HitAndWinCheck
End sub
Sub KPhesantB_Hit
  PhesantHitB=True
  Addpoints(1000)
  HitAndWinCheck
End sub
Sub KPhesantC_Hit
  PhesantHitC=True
  Addpoints(1000)
  HitAndWinCheck
End sub
Sub KPhesantA_unHit
  Addpoints(-1000)
  PhesantHitA = False
  HitAndWinCheck
End Sub

Sub KPhesantB_unHit
  Addpoints(-1000)
  PhesantHitB = False
  HitAndWinCheck
End Sub

Sub KPhesantC_unHit
  Addpoints(-1000)
  PhesantHitC = False
  HitAndWinCheck
End Sub

'Rabbit
Dim RabbitHitA,RabbitHitB,RabbitHitC,RabbitWin
Sub KRabbitA_Hit
  RabbitHitA=True
  Addpoints(800)
  HitAndWinCheck
End sub
Sub KRabbitB_Hit
  RabbitHitB=True
  Addpoints(800)
  HitAndWinCheck
End sub
Sub KRabbitC_Hit
  RabbitHitC=True
  Addpoints(800)
  HitAndWinCheck
End sub
Sub KRabbitA_unHit
  Addpoints(-800)
  RabbitHitA = False
  HitAndWinCheck
End Sub

Sub KRabbitB_unHit
  Addpoints(-800)
  RabbitHitB = False
  HitAndWinCheck
End Sub

Sub KRabbitC_unHit
  Addpoints(-800)
  RabbitHitC = False
  HitAndWinCheck
End Sub

'Squirrel
Dim SquirrelHitA,SquirrelHitB,SquirrelWin
Sub KSquirrelA_Hit
  SquirrelHitA=True
  Addpoints(600)
  HitAndWinCheck
End sub
Sub KSquirrelB_Hit
  SquirrelHitB=True
  Addpoints(600)
  HitAndWinCheck
End sub

Sub KSquirrelA_unHit
  Addpoints(-600)
  SquirrelHitA = False
  HitAndWinCheck
End Sub

Sub KSquirrelB_unHit
  Addpoints(-600)
  SquirrelHitB = False
  HitAndWinCheck
End Sub

'Duck
Dim DuckHitA,DuckHitB,DuckWin
Sub KDuckA_Hit
  DuckHitA=True
  Addpoints(500)
  HitAndWinCheck
End sub
Sub KDuckB_Hit
  DuckHitB=True
  Addpoints(500)
  HitAndWinCheck
End sub

Sub KDuckA_unHit
  Addpoints(-500)
  DuckHitA = False
  HitAndWinCheck
End Sub

Sub KDuckB_unHit
  Addpoints(-500)
  DuckHitB = False
  HitAndWinCheck
End Sub

'Vanilla points

'100 Points (4)
Sub K100A_Hit
  Addpoints (100)
End Sub

Sub K100B_Hit
  Addpoints (100)
End Sub

Sub K100C_Hit
  Addpoints (100)
End Sub

Sub K100D_Hit
  Addpoints (100)
End Sub

Sub K100A_unHit
  Addpoints (-100)
End Sub

Sub K100B_unHit
  Addpoints (-100)
End Sub

Sub K100C_unHit
  Addpoints (-100)
End Sub

Sub K100D_unHit
  Addpoints (-100)
End Sub

'200 Points (6)
Sub K200A_Hit
  Addpoints (200)
End Sub

Sub K200B_Hit
  Addpoints (200)
End Sub

Sub K200C_Hit
  Addpoints (200)
End Sub

Sub K200D_Hit
  Addpoints (200)
End Sub

Sub K200E_Hit
  Addpoints (200)
End Sub

Sub K200F_Hit
  Addpoints (200)
End Sub

Sub K200A_unHit
  Addpoints (-200)
End Sub

Sub K200B_unHit
  Addpoints (-200)
End Sub

Sub K200C_unHit
  Addpoints (-200)
End Sub

Sub K200D_unHit
  Addpoints (-200)
End Sub

Sub K200E_unHit
  Addpoints (-200)
End Sub

Sub K200F_unHit
  Addpoints (-200)
End Sub

'300 Points (4)
Sub K300A_Hit
  Addpoints (300)
End Sub

Sub K300B_Hit
  Addpoints (300)
End Sub

Sub K300C_Hit
  Addpoints (300)
End Sub

Sub K300D_Hit
  Addpoints (300)
End Sub

Sub K300A_unHit
  Addpoints (-300)
End Sub

Sub K300B_unHit
  Addpoints (-300)
End Sub

Sub K300C_unHit
  Addpoints (-300)
End Sub

Sub K300D_unHit
  Addpoints (-300)
End Sub

'400 Points (2)

Sub K400A_Hit
  Addpoints (400)
End Sub

Sub K400B_Hit
  Addpoints (400)
End Sub

Sub K400A_unHit
  Addpoints (-400)
End Sub

Sub K400B_unHit
  Addpoints (-400)
End Sub

'*****GI Lights On
'dim xx
'
'For each xx in GI:xx.State = 1: Next


'~~~~~~~~~~~~~~~~~~~~
'        Tilt (currently not working)
'~~~~~~~~~~~~~~~~~~~~
Dim TableTilted
Dim TiltCount
'Dim CheckTilt

'Sub TiltIt()
'   TiltCount = TiltCount + 1
'   CheckTilt
'End sub

'Sub CheckTilt
' If TiltCount > 4 then TiltKicker.Kick
'End sub

'Haven't figure out how to get the ball to kick out when you tilt it


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


'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 300) * VolumeMultiplier
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / Hunter.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
  Set LBall = ball
  BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )

  if BallVel > 0 then
    GameEndTimer.enabled = False
  end If
End Function

'Function BallVel(LBall)
'BallVel = INT(SQR((lball.VelX ^2) + (LBall.VelY ^2) ) )
'End Function

'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 11 ' total number of balls
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub RollingTimer_Timer()
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
        If BallVel(BOT(b) ) > 1 AND BOT(b).z > 30 Then
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0
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
  PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0
End Sub

'******************************************
' Explanation of the rolling sound routine
'******************************************

' sounds are played based on the ball speed and position

' the routine checks first for deleted balls and stops the rolling sound.

' The For loop goes through all the balls on the table and checks for the ball speed and
' if the ball is on the table (height lower than 30) then then it plays the sound
' otherwise the sound is stopped, like when the ball has stopped or is on a ramp or flying.

' The sound is played using the VOL, AUDIOPAN, AUDIOFADE and PITCH functions, so the volume and pitch of the sound
' will change according to the ball speed, and the AUDIOPAN & AUDIOFADE functions will change the stereo position
' according to the position of the ball on the table.


'**************************************
' Explanation of the collision routine
'**************************************

' The collision is built in VP.
' You only need to add a Sub OnBallBallCollision(ball1, ball2, velocity) and when two balls collide they
' will call this routine. What you add in the sub is up to you. As an example is a simple Playsound with volume and paning
' depending of the speed of the collision.


'Sub Gates_Hit (idx)
' PlaySound "Gate Hit", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
'End Sub

Sub Bumper_Hit '(idx)
  PlaySound "BumperHit",0,.25,0,0.25
End Sub

Sub Pegs_Hit(idx)
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "metalhit_medium", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 2 : PlaySound "Peg3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 3 : PlaySound "Peg4", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
  End Select
End Sub

Sub Guides_Hit(idx)
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "metalhit_thin", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 2 : PlaySound "metalhit2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 3 : PlaySound "Peg3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
  End Select
End Sub

'---------------------------
'----- High Score Code -----
'---------------------------
Const HighScoreFilename = "JenningsHunter.txt"

Dim HSArray,HSAHighScore, HSA1, HSA2, HSA3
Dim HSScoreM,HSScore100k, HSScore10k, HSScoreK, HSScore100, HSScore10, HSScore1, HSScorex 'Define 6 different score values for each reel to use
HSArray = Array("HS_0","HS_1","HS_2","HS_3","HS_4","HS_5","HS_6","HS_7","HS_8","HS_9","HS_Space","HS_Comma","Tape")
Const DefaultHighScore = 0

Sub LoadHighScore
  Dim FileObj
  Dim ScoreFile
  Dim TextStr
    Dim SavedDataTemp3 'HighScore
    Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if
  If Not FileObj.FileExists(UserDirectory & HighScoreFilename) then
    SetDefaultHSTD:UpdatePostIt:SaveHighScore
    Exit Sub
  End if
  Set ScoreFile=FileObj.GetFile(UserDirectory & HighScoreFilename)
  Set TextStr=ScoreFile.OpenAsTextStream(1,0)
    If (TextStr.AtEndOfStream=True) then
      SetDefaultHSTD:UpdatePostIt:SaveHighScore
      Exit Sub
    End if
    SavedDataTemp3=Textstr.ReadLine ' HighScore
    TextStr.Close
    HSAHighScore=SavedDataTemp3
    UpdatePostIt
      Set ScoreFile = Nothing
      Set FileObj = Nothing
End Sub

Sub SetDefaultHSTD  'bad data or missing file - reset and resave
  HSAHighScore = DefaultHighScore
  SaveHighScore
End Sub

Sub UpdatePostIt
  HSScorex = HSAHighScore
  TempScore = HSScorex
  HSScore1 = 0
  HSScore10 = 0
  HSScore100 = 0
  HSScoreK = 0
  HSScore10k = 0
  HSScore100k = 0
  HSScoreM = 0
  if len(TempScore) > 0 Then
    HSScore1 = cint(right(Tempscore,1))
  end If
  if len(TempScore) > 1 Then
    TempScore = Left(TempScore,len(TempScore)-1)
    HSScore10 = cint(right(Tempscore,1))
  end If
  if len(TempScore) > 1 Then
    TempScore = Left(TempScore,len(TempScore)-1)
    HSScore100 = cint(right(Tempscore,1))
  end If
  if len(TempScore) > 1 Then
    TempScore = Left(TempScore,len(TempScore)-1)
    HSScoreK = cint(right(Tempscore,1))
  end If
  if len(TempScore) > 1 Then
    TempScore = Left(TempScore,len(TempScore)-1)
    HSScore10k = cint(right(Tempscore,1))
  end If
  if len(TempScore) > 1 Then
    TempScore = Left(TempScore,len(TempScore)-1)
    HSScore100k = cint(right(Tempscore,1))
  end If
  if len(TempScore) > 1 Then
    TempScore = Left(TempScore,len(TempScore)-1)
    HSScoreM = cint(right(Tempscore,1))
  end If
  Pscore6.image = HSArray(HSScoreM):If HSScorex<1000000 Then PScore6.image = HSArray(10)
  Pscore5.image = HSArray(HSScore100K):If HSScorex<100000 Then PScore5.image = HSArray(10)
  PScore4.image = HSArray(HSScore10K):If HSScorex<10000 Then PScore4.image = HSArray(10)
  PScore3.image = HSArray(HSScoreK):If HSScorex<1000 Then PScore3.image = HSArray(10)
  PScore2.image = HSArray(HSScore100):If HSScorex<100 Then PScore2.image = HSArray(10)
  PScore1.image = HSArray(HSScore10):If HSScorex<10 Then PScore1.image = HSArray(10)
  PScore0.image = HSArray(HSScore1):If HSScorex<1 Then PScore0.image = HSArray(10)
  if HSScorex<1000 then
    PComma.image = HSArray(10)
  else
    PComma.image = HSArray(11)
  end if
  if HSScorex<1000000 then
    PComma2.image = HSArray(10)
  else
    PComma2.image = HSArray(11)
  End If
End Sub

Dim FileObj,ScoreFile
Sub SaveHighScore
  Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if
  Set ScoreFile=FileObj.CreateTextFile(UserDirectory & HighScoreFilename,True)
    ScoreFile.WriteLine HSAHighScore
    ScoreFile.Close
  Set ScoreFile=Nothing
  Set FileObj=Nothing
End Sub

Sub CheckNewHighScorePostIt(newScore)
    If CLng(newScore) > CLng(HSAHighScore) Then
'     AddSpecial
      HSAHighScore=newScore
      SaveHighScore
      UpdatePostIt
    End If
End Sub

Sub Addpoints(ScorePar)
    if GameStarted Then
      Playerscore = Playerscore + ScorePar
    end If

'   if (Playerscore >= Special1Score) and not Special1 Then
'     Special1 = True
'     AddSpecial
'   end If
'   if (Playerscore >= Special2Score) and not Special2 Then
'     Special2 = True
'     AddSpecial
'   end If
'   if (Playerscore >= Special3Score) and not Special3 Then
'     Special3 = True
'     AddSpecial
'   end If
    UpdateScoreReel.enabled = True
End Sub

'xxx

Sub UpdateScoreReel_Timer
  TempScore = Playerscore
  Score10 = 0
  Score100 = 0
  Score1000 = 0
  Score10000 = 0
  if len(TempScore) > 1 Then
    TempScore = Left(TempScore,len(TempScore)-1)
    Score10 = cstr(right(Tempscore,1))
  end If
  if len(TempScore) > 1 Then
    TempScore = Left(TempScore,len(TempScore)-1)
    Score100 = cstr(right(Tempscore,1))
  end If
  if len(TempScore) > 1 Then
    TempScore = Left(TempScore,len(TempScore)-1)
    Score1000 = cstr(right(Tempscore,1))
  end If
  if len(TempScore) > 1 Then
    TempScore = Left(TempScore,len(TempScore)-1)
    Score10000 = cstr(right(Tempscore,1))
  end If

  Val10 = ReelValue(Val10,Score10,1)
  Val100 = ReelValue(Val100,Score100,2)
  Val1000 = ReelValue(Val1000,Score1000,3)
  Val10000 = ReelValue(Val10000,Score10000,0)
  Tempscore = Val10 * 10 + Val100 * 100 + Val1000 * 1000 + Val10000 * 10000
  if Oldscore <> TempScore Then
    'playsound "solon",0,0.3,0.1,0.25
    Oldscore = TempScore
    P_Reel1.image = cstr(Val10)
    P_Reel2.image = cstr(Val100)
    P_Reel3.image = cstr(Val1000)
    P_Reel4.image = cstr(Val10000)
  Else
    UpdateScoreReel.enabled = False
  end If

  ScoreText.Text = "Score : " & Playerscore
  Controller.B2SSetScorePlayer1 Playerscore
  If DuckWin Then
    DuckSquirrelTextBox.Text = "All the ducks caught !"
  Else
    Dim ducksNumber
    If DuckHitA Then ducksNumber = ducksNumber + 1 End If
    If DuckHitB Then ducksNumber = ducksNumber + 1 End If
    If ducksNumber > 0 Then
      DuckSquirrelTextBox.Text = "Duck hit : " & ducksNumber
    End If
  End If
  If SquirrelWin Then
    DuckSquirrelTextBox.Text = "All the squirrels caught !"
  Else
    Dim squirrelsNumber
    If SquirrelHitA Then squirrelsNumber = squirrelsNumber + 1 End If
    If SquirrelHitB Then squirrelsNumber = squirrelsNumber + 1 End If
    If squirrelsNumber > 0 Then
      DuckSquirrelTextBox.Text = "Squirrel hit : " & squirrelsNumber
    End If
  End If
  If RabbitWin Then
    RabbitTextBox.Text = "All the rabbits caught !"
  Else
    Dim rabbitsNumber
    If RabbitHitA Then rabbitsNumber = rabbitsNumber + 1 End If
    If RabbitHitB Then rabbitsNumber = rabbitsNumber + 1 End If
    If RabbitHitC Then rabbitsNumber = rabbitsNumber + 1 End If
    If rabbitsNumber > 0 Then
      RabbitTextBox.Text = "Rabbit hit : " & rabbitsNumber
    End If
  End If
  If PhesantWin Then
    PheasantTextBox.Text = "All the pheasants caught !"
  Else
    Dim pheasantsNumber
    If PhesantHitA Then pheasantsNumber = pheasantsNumber + 1 End If
    If PhesantHitB Then pheasantsNumber = pheasantsNumber + 1 End If
    If PhesantHitC Then pheasantsNumber = pheasantsNumber + 1 End If
    If pheasantsNumber > 0 Then
      PheasantTextBox.Text = "Pheasant hit : " & pheasantsNumber
    End If
  End If
  If PartridgeWin Then
    PartridgeTextBox.Text = "All the partridges caught !"
  Else
    Dim partdrigesNumber
    If PartridgeHitA Then partdrigesNumber = partdrigesNumber + 1 End If
    If PartridgeHitB Then partdrigesNumber = partdrigesNumber + 1 End If
    If PartridgeHitC Then partdrigesNumber = partdrigesNumber + 1 End If
    If partdrigesNumber > 0 Then
      PartridgeTextBox.Text = "Partridge hit : " & partdrigesNumber
    End If
  End If
  If SkillShotWin Then

  End If
End Sub

Function ReelValue(ValPar,ScorPar,ChimePar)
  ReelValue = cint(ValPar)
  if ReelValue <> cint(ScorPar) Then
    ReelValue = ReelValue + 1
    if ReelValue > 9 Then
      ReelValue = 0
    end If
  end If
End Function

Sub SetScoreReel
  TempScore = Playerscore
  Score10 = 0
  Score100 = 0
  Score1000 = 0
  Score10000 = 0
  if len(TempScore) > 1 Then
    TempScore = Left(TempScore,len(TempScore)-1)
    Score10 = cstr(right(Tempscore,1))
  end If
  if len(TempScore) > 1 Then
    TempScore = Left(TempScore,len(TempScore)-1)
    Score100 = cstr(right(Tempscore,1))
  end If
  if len(TempScore) > 1 Then
    TempScore = Left(TempScore,len(TempScore)-1)
    Score1000 = cstr(right(Tempscore,1))
  end If
  if len(TempScore) > 1 Then
    TempScore = Left(TempScore,len(TempScore)-1)
    Score10000 = cstr(right(Tempscore,1))
  end If
  P_Reel1.image = cstr(Score10)
  P_Reel2.image = cstr(Score100)
  P_Reel3.image = cstr(Score1000)
  P_Reel4.image = cstr(Score10000)
End Sub

' Animation on B2S : we animate the animals images on B2S at start of the game
Dim AnimationOnB2SStep:AnimationOnB2SStep = 0
Sub AnimationOnB2S
  AnimationOnB2SStep = 0
  AnimationOnB2STimer.Interval = 400
  AnimationOnB2STimer.Enabled = True
End Sub

Sub AnimationOnB2STimer_Timer
  Select Case AnimationOnB2SStep
    Case 0:Controller.B2SSetData 50,1
    Case 1:Controller.B2SSetData 51,1
    Case 2:Controller.B2SSetData 52,1
    Case 3:Controller.B2SSetData 53,1
    Case 4:Controller.B2SSetData 53,1
    Case 5:Controller.B2SSetData 53,0
    Case 6:Controller.B2SSetData 52,0
    Case 7:Controller.B2SSetData 51,0
    Case 8:Controller.B2SSetData 50,0
    Case Else : AnimationOnB2STimer.Enabled = False
  End Select
  AnimationOnB2SStep = AnimationOnB2SStep + 1
End Sub

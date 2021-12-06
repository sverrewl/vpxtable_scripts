'### Lucky Luke Mod by kds70 & vogliadicane
'### What has been done:
'###
'### 12-09 instr.cards added with lights
'### 13-09 added flipper holding buzz sounds and generates random kicker sounds
'###       tilt dof knocker added, ending flipper buzz sounds when tilt/ game ends (line 429)
'### 14-09 line 446: flipper dof off when buttons being pressed and game ended added
'### 15-09 ballshadow / flippershadow / hopping ball sounds added
'### 17-09 todo: VolZ and other volume adjustment
'### 23-09 ballmass / size working now and modified balldrop random interval
'### 24-09 added glas hit but deactivated
'### 26-09 line 565 reel score override backglas b2s lamp 25 (player1) 26 (player2) etc integration
'###       to show when 99.999 points are overidden.
'### 29-09 droptarget back light switching added
'### 02-10 2nd layer insert lights for casting insert lights to ball
'### 08-10 2x / 3x bonus insert added / scripted. added extra ball insert and Function
'###       line 1218 and 1174 check / add extra ball when DTargets 1st time dropped
'### 11-10 inserts new (gottlieb font)
'### 13-10 extraball (one time per game and one per player works now for testing)
'### 19-10 extraball error fixed: if got an extra ball with the last 3rd ball the table thought it was the last ball; didn´t play
'###       a ball release sound again for the extra ball. fixed with a counter "DT5kHits" variable at endbonus timer line 1539
'### 13-11 added collect bonus routine & timer (line 950-978) if left / right kicker is lit
'###     DropTargets Reset / upfiring combined with dof 128 (Bumper left/right) solenoid call line 1362 and 1380
'###     ShadowMap modified by vogliadicane
'### 20-03 some music / sound effects added
'### 21-03 Pin1, Pin2, Pin3 and Pin4 ( height set to 25) and Wall78 (friction to 0.3)



'############################################################################################
'############################################################################################
'#######                                                                             ########
'#######          Fast Draw (Lucky Luke Mod)                                         ########
'#######          (Gottlieb 1975)                                                    ########
'#######                                                                             ########
'############################################################################################
'############################################################################################
' Version 1.4 mfuegemann 2016
'
' Thanks To:
' Zany for the Flipper model and texture
' JPSalas for the EM Reel images
' GTXJoe for the Highscore saving code
'
' Version 1.1:
' - GI light modulation set to 0.99 - thanks to hauntfreaks
' - Shadow layer and environment map by hauntfreaks
' - Kicker animation added
' - Coin sound added - thanks to STAT
' - Minor bugfixes
'
' Version 1.2:
' - EM Reel reset sequence fixed, no additional reset after first points
' - Primitive Apron
' - some sounds reviewed
'
' Version 1.3:
' - adjustable sound level for ball rolling sound
' - fixed standard sounds for rubbers, posts etc.
' - 2 rubber leaf switches animated
' - new inside case texture for DT mode
'
' Version 1.4:
' - B2S HighScore display fixed
' - Lane guide elasticity adjusted
' - 5K DropTarget reset animation changed
' - DOF support added by Arngrim
' - Chimes added by Arngrim
' - Option to mute Chimes added by mfuegemann ;)
' - Option to control Chime volume added


option Explicit

Dim BallSize,BallMass,LetTheBallJump

BallSize = 50 'create ball in line 386
BallMass = 1.3  'create ball in line 386

' Thalamus 2020 March : Improved directional sounds
' !! NOTE : Table not verified yet !!

' Options
' Volume devided by - lower gets higher sound

Const VolDiv = 2000    ' Lower number, louder ballrolling/collition sound
Const VolCol = 10      ' Ball collition divider ( voldiv/volcol )

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

Const cGameName = "Fast_Draw_1975"

'---------------------------
'-----  Configuration  -----
'---------------------------

Const GameMusic = 1                 'ingame western music on (1) off (0)
Const BallsperGame = 3        'the original table allows for 3 or 5 Balls
Const SetBallShadow = 1             'enable (1) disable (0)
Const ShowHighScore = True      'enable to show HighScore tape on Apron, a new HighScore will be saved anyway
Const ShowPlayfieldReel = False   'set to True to show Playfield EM-Reel and Post-It for Player Scores (instead of a B2S backglass)
Const FreePlay = False
Const RollingSoundFactor = 3.0    'set sound level factor here for Ball Rolling Sound, 1=default level
Const ChimesEnabled = True      'You have been warned!
Const ChimeVolume = 0.3       'to limit ear bleeding set between 0 and 1
Const ResetHighscore = False    'enable to manually reset the Highscores
Const Special1Score = 98000     'set the 3 replay scores (3rd for display rollover)
Const Special2Score = 123000
Const Special3Score = 150000

Dim StoreBonus,GameActive,NoofPlayers,i,HighScore,Credits,light,NoOfExtraBalls,XtraBall,ExtraBallP1,ExtraBallP2,ExtraBallP3,ExtraBallP4
Dim GotExtraBallP1,GotExtraBallP2,GotExtraBallP3,GotExtraBallP4,DT5kHits



Sub FastDraw_Init

'*** turn gi lights on when game starts ***
  For each light in GI:light.State = 0: Next

'*** set lights and arrays
'Lights(1) = Array(Extraball,Extraball2)


  LoadEM

    if SetBallShadow = 0 then
         BallShadowUpdate.enabled=False
    End if

  LoadHighScore
  HideScoreboard
  if ResetHighScore then
    SetDefaultHSTD
  end if
  if ShowHighScore then
    PTape2.image = HSArray(12)
    UpdatePostIt
  else
    PTape2.image = HSArray(10)
    PComma.image = HSArray(10)
    PComma2.image = HSArray(10)
    Pscore0.image = HSArray(10)
    PScore1.image = HSArray(10)
    PScore2.image = HSArray(10)
    PScore3.image = HSArray(10)
    PScore4.image = HSArray(10)
    PScore5.image = HSArray(10)
    PScore6.image = HSArray(10)
  end if

  Randomize

  if FreePlay then
    Credits = 5
    DOF 126, DOFOn
  end If

  if B2SOn then
    Controller.B2SSetMatch MatchValue
    Controller.B2SSetScoreRolloverPlayer1 0
    Controller.B2SSetScoreRolloverPlayer2 0
    Controller.B2SSetScoreRolloverPlayer3 0
    Controller.B2SSetScoreRolloverPlayer4 0
    Controller.B2SSetScore 6,HighScore
    Controller.B2SSetTilt 0
    for i=50 to 59
      Controller.B2SSetData i,0
    next
    if Credits=0 then
      Controller.B2SSetData 50+Credits,1
      DOF 126, DOFOff
    else
      DOF 126, DOFOn
      Controller.B2SSetData 50+Credits,1
    end if
    Controller.B2SSetGameOver 1

    for i=1 to 4
      Controller.B2SSetScorePlayer i, 0
    next
  end if

  if not FastDraw.ShowDT then
    EMReel1.visible = False
    EMReel2.visible = False
    EMReel3.visible = False
    EMReel4.visible = False
    EMReel_BiP.visible = False
    EMReel_Credits.visible = False
    CaseWall1.isdropped = True
    CaseWall2.isdropped = True
    CaseWall3.isdropped = True
    Ramp1.widthbottom = 0
    Ramp1.widthtop = 0
    Ramp15.widthbottom = 0
    Ramp15.widthtop = 0
  end If

  if not ShowPlayfieldReel or FastDraw.ShowDT Then
    ReelWall.isdropped = True
    P_Reel0.Transz = -5
    P_Reel1.Transz = -5
    P_Reel2.Transz = -5
    P_Reel3.Transz = -5
    P_Reel4.Transz = -5
    P_Reel5.Transz = -5
    P_ActivePlayer.Transz = -10
    P_Credits.Transz = -10
    P_CreditsText.Transz = -10
    P_BallinPlay.Transz = -10
    P_BallinPlayText.Transz = -10
  end If

  BallinPlay = 0
  GameActive = False
  NoofPlayers = 0
  GameStarted = False
  TargetSound = "FD_500Target"
  DTSound = "fx_droptarget"
  RolloverSound1 = "FD_Top_Rollover"
  RolloverSound2 = "FD_Top_Rollover2"
  OutlaneSound = "FD_Outlane"
  Soundlevel = 1
  for i = 1 to 4
    Val10(i) = 0
    Val100(i) = 0
    Val1000(i) = 0
    Val10000(i) = 0
    Val100000(i) = 0
    Oldscore(i) = 0
  Next
  EMReel1.ResetToZero
  EMReel2.ResetToZero
  EMReel3.ResetToZero
  EMReel4.ResetToZero
  EMReel_Credits.setvalue Credits

  P_Credits.image = cstr(Credits)
  BumperWall1.isdropped = True
  BumperWall2.isdropped = True
  BumperWall3.isdropped = True
  SwitchA_2.isdropped = True
  SwitchB_2.isdropped = True
  LeftDTBank_Cover.isdropped = True
  RightDTBank_Cover.isdropped = True

  Tilt = 0

End Sub

Dim GameStarted,NoPointsScored
Sub StartGame
  if Credits > 0 Then
    if not GameActive then
      for i = 1 to 4
        Playerscore(i) = 0
        Oldscore(i) = 0
        Val10(i) = 0
        Val100(i) = 0
        Val1000(i) = 0
        Val10000(i) = 0
        Val100000(i) = 0
        Special1(i) = False
        Special2(i) = False
        Special3(i) = False
        If B2SOn Then
          Controller.B2SSetScorePlayer i,0

          Controller.B2SSetData (24+i), 0

        end if
      Next
      EMReel1.ResetToZero
      EMReel2.ResetToZero
      EMReel3.ResetToZero
      EMReel4.ResetToZero
      if not GameStarted Then
        GameStarted = True

'### GI anschalten
      For each light in GI:light.State = 1: Next
      Playsound "target"
      DOF 20, DOFon 'b2s lamps 20 by Alex

'### GameMusic Timer on off

      EndMusic

      if GameMusic = 1 then
        m01.enabled=True
        m01_Timer
      End if

    PlaySound "BallOut Pat Woods Start2b lonesome alone",0,0.5

        Playsound "FD_GameStartwithBallrelease"

        GameStartTimer.enabled = True
        ScoreMotor
        ScoreMotorStartTimer.enabled = True
      end If
      if NoofPlayers < 4 Then
        Credits = Credits - 1
        If Credits < 1 Then DOF 126, DOFOff
        if FreePlay and (Credits = 0) then
          Credits = 5
          DOF 126, DOFOn
        end If
        if NoofPlayers > 0 Then
          Playsound "AddPlayer"
        end If
        NoofPlayers = NoofPlayers + 1
        UpdateScoreboard
        EMReel_Credits.setvalue Credits
      end If

      If B2SOn Then
        Controller.B2SSetTilt 0
        Controller.B2SSetGameOver 0
        Controller.B2SSetMatch 0
        Controller.B2SSetBallInPlay BallInPlay
        for i=50 to 59
          Controller.B2SSetData i,0
        next

        Controller.B2SSetData 50+Credits,1
        Controller.B2SSetPlayerUp 1
        Controller.B2SSetBallInPlay BallInPlay
        Controller.B2SSetCanPlay NoofPlayers
        Controller.B2SSetScore 6,HSAHighScore
      End if
      EMReel_BiP.setvalue BallinPlay
    end If
  end If
  P_Credits.image = cstr(Credits)
End Sub

'##################
'### Music Play ###
'##################

'### interval
'### 230700 milliseconds = 3 minutes 50 seconds 700 milliseconds.
'### 3 x 60 seconds x 1000 = 180000, plus 50 x 1000 = 230000, plus 700 milliseconds = 230700


Sub m01_Timer
    Dim SoundShuffle
    SoundShuffle = INT(5 * RND(1) )
    Select Case SoundShuffle
    Case 0:PlayMusic"FD_1.mp3":m01.enabled=true:m01.interval=179000    '2.59
    Case 1:PlayMusic"FD_2.mp3":m01.enabled=true:m01.interval=162000    '2.42
    Case 2:PlayMusic"FD_3.mp3":m01.enabled=true:m01.interval=186000    '3.06
    Case 3:PlayMusic"FD_4.mp3":m01.enabled=true:m01.interval=150000    '2.30
    Case 4:PlayMusic"FD_5.mp3":m01.enabled=true:m01.interval=171000    '2.51
    End Select
End Sub


Sub GameStartTimer_Timer
  GameStartTimer.enabled = False
  ScoreMotorStartTimer.enabled = False
  GotExtraBallP1 = 0
  GotExtraBallP2 = 0
  GotExtraBallP3 = 0
  GotExtraBallP4 = 0
  ExtraBallP1 = 0
  ExtraBallP2 = 0
  ExtraBallP3 = 0
  ExtraBallP4 = 0
  NoOfExtraBalls = 0
  XtraBall = 0
  DT5kHits=0
  ActivePlayer = 0
  BallinPlay = 1
  NextBall
End Sub

Sub ScoreMotorStartTimer_Timer
  ScoreMotor
End Sub

Sub NextBall

ExtraBall.State = 0: ExtraBall2nd.State = 0
ExtraBall2.State = 0: ExtraBall22nd.State = 0
ShootAgain.State = 0: ShootAgain2nd.State = 0

'*** Check if someone got an extra Ball ****
if (GotExtraBallP1 = 1) or (GotExtraBallP2 = 1) or (GotExtraBallP3 = 1) or (GotExtraBallP4 = 1) Then

    if ActivePlayer = 1 then GotExtraBallP1 = 2
    if ActivePlayer = 2 then GotExtraBallP2 = 2
    if ActivePlayer = 3 then GotExtraBallP3 = 2
    if ActivePlayer = 4 then GotExtraBallP4 = 2

  UpdateScoreboard

    P_ActivePlayer.image = "Player" & CStr(ActivePlayer)
    P_BallinPlay.image = CStr(BallinPlay)
    EMReel_BiP.setvalue BallinPlay

  SetScoreReel

    If B2SOn Then
      Controller.B2SSetBallInPlay BallInPlay
      Controller.B2SSetPlayerUp ActivePlayer
      Controller.B2SSetBallInPlay BallInPlay
      Controller.B2SSetTilt 0
    End if

    ResetPlayfield
    if BallinPlay = BallsperGame Then
      DB = True
      ExtraBonus2x.state = Lightstateon: ExtraBonus2x2nd.state = Lightstateon
    end If

    NoPointsScored = True

'   BallRelease.CreateBall
    BallRelease.CreateSizedBallWithMass Ballsize/2, BallMass
    BallRelease.Kick 90, 7
    DOF 122, DOFPulse

Else


'*** delete extra balls if not gotten wben ball wents into drain ***

  if ActivePlayer = 1 And ExtraBallP1 = 1 then GotExtraBallP1 = 2: ExtraBallP1 = 2
  if ActivePlayer = 2 And ExtraBallP2 = 1 then GotExtraBallP2 = 2: ExtraBallP2 = 2
  if ActivePlayer = 3 And ExtraBallP3 = 1 then GotExtraBallP3 = 2: ExtraBallP3 = 2
  if ActivePlayer = 4 And ExtraBallP4 = 1 then GotExtraBallP4 = 2: ExtraBallP4 = 2

'***

  if (ActivePlayer = NoofPlayers) and (BallinPlay = BallsperGame) Then

    EndGame

  Else

    UpdateScoreboard

    if ActivePlayer < NoofPlayers Then
      ActivePlayer = ActivePlayer + 1
    Else
      if ActivePlayer = NoofPlayers Then
        BallinPlay = BallinPlay + 1
        ActivePlayer = 1
      end If
    end If

    P_ActivePlayer.image = "Player" & CStr(ActivePlayer)
    P_BallinPlay.image = CStr(BallinPlay)
    EMReel_BiP.setvalue BallinPlay

    SetScoreReel

    If B2SOn Then
      Controller.B2SSetBallInPlay BallInPlay
      Controller.B2SSetPlayerUp ActivePlayer
      Controller.B2SSetBallInPlay BallInPlay
      Controller.B2SSetTilt 0
    End if

    ResetPlayfield
    if BallinPlay = BallsperGame Then
      DB = True
      ExtraBonus2x.state = Lightstateon: ExtraBonus2x2nd.state = Lightstateon
    end If
    NoPointsScored = True
'   BallRelease.CreateBall
    BallRelease.CreateSizedBallWithMass Ballsize/2, BallMass
    BallRelease.Kick 90, 7
    DOF 122, DOFPulse

  End If

End If

End Sub

Dim MatchValue

Sub EndGame

'### GI ausschalten
For each light in GI:light.State = 0: Next
For each light in GI_DTargets:light.State = 0: Next

EndMusic

PlayMusic "FD_6.mp3"

DOF 20, DOFoff 'b2s illumination lamp 20 off

'### GameMusic Timer on off
    if GameMusic = 1 then
         m01.enabled=False
    End if

  UpdateScoreboard
  for i = 1 to NoofPlayers
    CheckNewHighScorePostIt(Playerscore(i))
  next
  GameActive = False
  GameStarted = False
  MatchValue=Int(Rnd*10)*10
  If B2SOn Then
    If MatchValue = 0 Then
      Controller.B2SSetMatch 100
    Else
      Controller.B2SSetMatch MatchValue
    End If
  End if
  for i = 1 to NoofPlayers
    if MatchValue=cint(right(Playerscore(i),2)) then
      AddSpecial
    end if
  next
  BallinPlay = 0
  NoofPlayers = 0
  if B2Son Then
    Controller.B2SSetGameOver 1
    Controller.B2SSetBallInPlay BallInPlay
  end If
End Sub

Sub Drain_Hit()
  PlaySoundAtVol "drain3", Drain, 1
  PlaySound "Horse Neighing",0,1.0

DOF 121, DOFPulse
  Drain.DestroyBall
  if NoPointsScored and (Tilt = 0) Then
    Playsound "FD_DrainwithBallrelease"
    ReleaseSameBallTimer.enabled = True
  Else
    AwardBonus
  End If
   StopSound "FlipBuzzLA"
   StopSound "FlipBuzzRA"
End Sub

Sub ReleaseSameBallTimer_Timer
  ReleaseSameBallTimer.enabled = False
' BallRelease.CreateBall
  BallRelease.CreateSizedBallWithMass Ballsize/2, BallMass
  BallRelease.Kick 90, 7
  DOF 122, DOFPulse
End Sub

Sub MainTimer_Timer
  Target1_Cover.isdropped = Target1.isdropped
  Target6_Cover.isdropped = Target6.isdropped
  if not GameStarted Then
    LeftFlipper.RotateToStart
    RightFlipper.RotateToStart
        DOF 101, DOFOff   'Flipper DOF off when being holded and game ends
    DOF 102, DOFOff   'Flipper DOF off when being holded and game ends
  end If
  if FastDraw.showdt Then
    P1Light.State = (Activeplayer = 1)
    P2Light.State = (Activeplayer = 2)
    P3Light.State = (Activeplayer = 3)
    P4Light.State = (Activeplayer = 4)
  end If
End Sub


Sub TiltTimer_Timer
  if Tilt = 0 Then
    P_Tilt.visible = False
  Else
    if FastDraw.showdt Then
      P_Tilt.visible = not P_Tilt.visible
    End If
  End If
End Sub


' DT Score Reels
Dim Val10(4),Val100(4),Val1000(4),Val10000(4),Val100000(4),Score10,Score100,Score1000,Score10000,Score100000,TempScore,Oldscore(5)
Sub UpdateScoreReel_Timer
  TempScore = Playerscore(ActivePLayer)
  Score10 = 0
  Score100 = 0
  Score1000 = 0
  Score10000 = 0
  Score100000 = 0
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
  if len(TempScore) > 1 Then
    TempScore = Left(TempScore,len(TempScore)-1)
    Score100000 = cstr(right(Tempscore,1))
  end If

  Val10(ActivePLayer) = ReelValue(Val10(ActivePLayer),Score10,1)
  Val100(ActivePLayer) = ReelValue(Val100(ActivePLayer),Score100,2)
  Val1000(ActivePLayer) = ReelValue(Val1000(ActivePLayer),Score1000,3)
  Val10000(ActivePLayer) = ReelValue(Val10000(ActivePLayer),Score10000,0)
  Val100000(ActivePLayer) = ReelValue(Val100000(ActivePLayer),Score100000,0)
  Tempscore = Val10(ActivePLayer) * 10 + Val100(ActivePLayer) * 100 + Val1000(ActivePLayer) * 1000 + Val10000(ActivePLayer) * 10000 + Val100000(ActivePLayer) * 100000
  if Oldscore(ActivePLayer) <> TempScore Then
    Oldscore(ActivePLayer) = TempScore
    select Case ActivePlayer
      case 1: EMReel1.setvalue TempScore
      case 2: EMReel2.setvalue TempScore
      case 3: EMReel3.setvalue TempScore
      case 4: EMReel4.setvalue TempScore
    End Select

    If B2SOn Then
      Controller.B2SSetScorePlayer ActivePlayer, TempScore


' *** added ScoreReel Override Check and Display (B2S DOF Lamp 25 ff)
'        if (Playerscore(ActivePLayer)) >= 10000 Then
'     DOF ((Playerscore(ActivePlayer))/10000)+24, DOFon
'       if (Playerscore(ActivePlayer)) > 20000 Then
'         DOF ((Playerscore(ActivePlayer))/10000)+23, DOFoff
'       end If
'   end If

    End If

    P_Reel1.image = cstr(Val10(ActivePLayer))
    P_Reel2.image = cstr(Val100(ActivePLayer))
    P_Reel3.image = cstr(Val1000(ActivePLayer))
    P_Reel4.image = cstr(Val10000(ActivePLayer))
    P_Reel5.image = cstr(Val100000(ActivePLayer))

    If B2SOn Then Controller.B2SSetData (ActivePLayer+24), cstr(Val100000(ActivePLayer))

  Else
    UpdateScoreReel.enabled = False
  end If
End Sub

Function ReelValue(ValPar,ScorPar,ChimePar)
  ReelValue = cint(ValPar)
  if ReelValue <> cint(ScorPar) Then
    if ChimesEnabled Then
      If ChimePar = 1 Then
        PlaySound SoundFXDOF("10a",129,DOFPulse,DOFChimes),0,ChimeVolume
      end If
      If ChimePar = 2 Then
        PlaySound SoundFXDOF("100a",130,DOFPulse,DOFChimes),0,ChimeVolume
      end If
      If ChimePar = 3 Then
        PlaySound SoundFXDOF("1000a",131,DOFPulse,DOFChimes),0,ChimeVolume
      End If
    end If
    ReelValue = ReelValue + 1
    if ReelValue > 9 Then
      ReelValue = 0
    end If
  end If
End Function

Sub SetScoreReel
  TempScore = Playerscore(ActivePLayer)
  Score10 = 0
  Score100 = 0
  Score1000 = 0
  Score10000 = 0
  Score100000 = 0
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
  if len(TempScore) > 1 Then
    TempScore = Left(TempScore,len(TempScore)-1)
    Score100000 = cstr(right(Tempscore,1))
  end If
  P_Reel1.image = cstr(Score10)
  P_Reel2.image = cstr(Score100)
  P_Reel3.image = cstr(Score1000)
  P_Reel4.image = cstr(Score10000)
  P_Reel5.image = cstr(Score100000)
End Sub



' FS Post-It Playerscores
Sub UpdateScoreboard
  if not ShowPlayfieldReel or (NoofPlayers < 2) or FastDraw.ShowDT Then
    HideScoreBoard
  Else
    P_SB_Postit.image = "Postit"
    Select case NoofPlayers
      Case 2:
        P_SB_Postit.size_y = 50
        P_SB_Postit.Transy = -30
        PScore1_P.image = "Player1"
        PScore2_P.image = "Player2"
        SetScoreBoard(1)
        SetScoreBoard(2)
      Case 3:
        P_SB_Postit.size_y = 75
        P_SB_Postit.Transy = -15
        PScore1_P.image = "Player1"
        PScore2_P.image = "Player2"
        PScore3_P.image = "Player3"
        SetScoreBoard(1)
        SetScoreBoard(2)
        SetScoreBoard(3)
      Case 4:
        P_SB_Postit.size_y = 100
        P_SB_Postit.Transy = 0
        PScore1_P.image = "Player1"
        PScore2_P.image = "Player2"
        PScore3_P.image = "Player3"
        PScore4_P.image = "Player4"
        SetScoreBoard(1)
        SetScoreBoard(2)
        SetScoreBoard(3)
        SetScoreBoard(4)
    end Select
  end If
End Sub

Sub HideScoreboard
  for each obj in C_ScoreBoard
    obj.image = HSArray(10)
  Next
End Sub

Dim SBScore100k, SBScore10k, SBScoreK, SBScore100, SBScore10, SBScore1, SBTempScore,obj

Sub SetScoreBoard(PlayerPar)
  SBTempScore = PlayerScore(PlayerPar)
  SBScore1 = 0
  SBScore10 = 0
  SBScore100 = 0
  SBScoreK = 0
  SBScore10k = 0
  SBScore100k = 0
  if len(SBTempScore) > 0 Then
    SBScore1 = cint(right(SBTempscore,1))
  end If
  if len(SBTempScore) > 1 Then
    SBTempScore = Left(SBTempScore,len(SBTempScore)-1)
    SBScore10 = cint(right(SBTempscore,1))
  end If
  if len(SBTempScore) > 1 Then
    SBTempScore = Left(SBTempScore,len(SBTempScore)-1)
    SBScore100 = cint(right(SBTempscore,1))
  end If
  if len(SBTempScore) > 1 Then
    SBTempScore = Left(SBTempScore,len(SBTempScore)-1)
    SBScoreK = cint(right(SBTempscore,1))
  end If
  if len(SBTempScore) > 1 Then
    SBTempScore = Left(SBTempScore,len(SBTempScore)-1)
    SBScore10k = cint(right(SBTempscore,1))
  end If
  if len(SBTempScore) > 1 Then
    SBTempScore = Left(SBTempScore,len(SBTempScore)-1)
    SBScore100k = cint(right(SBTempscore,1))
  end If
  Select case PlayerPar
    Case 1:
      Pscore6_1.image = HSArray(SBScore100K):If PlayerScore(PlayerPar)<100000 Then Pscore6_1.image = HSArray(10)
      Pscore5_1.image = HSArray(SBScore10K):If PlayerScore(PlayerPar)<10000 Then Pscore5_1.image = HSArray(10)
      Pscore4_1.image = HSArray(SBScoreK):If PlayerScore(PlayerPar)<1000 Then Pscore4_1.image = HSArray(10)
      Pscore3_1.image = HSArray(SBScore100):If PlayerScore(PlayerPar)<100 Then Pscore3_1.image = HSArray(10)
      Pscore2_1.image = HSArray(SBScore10):If PlayerScore(PlayerPar)<10 Then Pscore2_1.image = HSArray(10)
      Pscore1_1.image = HSArray(SBScore1):If PlayerScore(PlayerPar)<1 Then Pscore1_1.image = HSArray(10)
      if PlayerScore(PlayerPar)<1000 then
        PscoreComma_1.image = HSArray(10)
      else
        PscoreComma_1.image = HSArray(11)
      end if
    Case 2:
      Pscore6_2.image = HSArray(SBScore100K):If PlayerScore(PlayerPar)<100000 Then Pscore6_2.image = HSArray(10)
      Pscore5_2.image = HSArray(SBScore10K):If PlayerScore(PlayerPar)<10000 Then Pscore5_2.image = HSArray(10)
      Pscore4_2.image = HSArray(SBScoreK):If PlayerScore(PlayerPar)<1000 Then PScore4_2.image = HSArray(10)
      Pscore3_2.image = HSArray(SBScore100):If PlayerScore(PlayerPar)<100 Then Pscore3_2.image = HSArray(10)
      Pscore2_2.image = HSArray(SBScore10):If PlayerScore(PlayerPar)<10 Then Pscore2_2.image = HSArray(10)
      Pscore1_2.image = HSArray(SBScore1):If PlayerScore(PlayerPar)<1 Then Pscore1_2.image = HSArray(10)
      if PlayerScore(PlayerPar)<1000 then
        PscoreComma_2.image = HSArray(10)
      else
        PscoreComma_2.image = HSArray(11)
      end if
    Case 3:
      Pscore6_3.image = HSArray(SBScore100K):If PlayerScore(PlayerPar)<100000 Then Pscore6_3.image = HSArray(10)
      Pscore5_3.image = HSArray(SBScore10K):If PlayerScore(PlayerPar)<10000 Then Pscore5_3.image = HSArray(10)
      Pscore4_3.image = HSArray(SBScoreK):If PlayerScore(PlayerPar)<1000 Then Pscore4_3.image = HSArray(10)
      Pscore3_3.image = HSArray(SBScore100):If PlayerScore(PlayerPar)<100 Then Pscore3_3.image = HSArray(10)
      Pscore2_3.image = HSArray(SBScore10):If PlayerScore(PlayerPar)<10 Then Pscore2_3.image = HSArray(10)
      Pscore1_3.image = HSArray(SBScore1):If PlayerScore(PlayerPar)<1 Then Pscore1_3.image = HSArray(10)
      if PlayerScore(PlayerPar)<1000 then
        PscoreComma_3.image = HSArray(10)
      else
        PscoreComma_3.image = HSArray(11)
      end if
    Case 4:
      Pscore6_4.image = HSArray(SBScore100K):If PlayerScore(PlayerPar)<100000 Then Pscore6_4.image = HSArray(10)
      Pscore5_4.image = HSArray(SBScore10K):If PlayerScore(PlayerPar)<10000 Then Pscore5_4.image = HSArray(10)
      Pscore4_4.image = HSArray(SBScoreK):If PlayerScore(PlayerPar)<1000 Then Pscore4_4.image = HSArray(10)
      Pscore3_4.image = HSArray(SBScore100):If PlayerScore(PlayerPar)<100 Then Pscore3_4.image = HSArray(10)
      Pscore2_4.image = HSArray(SBScore10):If PlayerScore(PlayerPar)<10 Then Pscore2_4.image = HSArray(10)
      Pscore1_4.image = HSArray(SBScore1):If PlayerScore(PlayerPar)<1 Then Pscore1_4.image = HSArray(10)
      if PlayerScore(PlayerPar)<1000 then
        PscoreComma_4.image = HSArray(10)
      else
        PscoreComma_4.image = HSArray(11)
      end if
  End Select
End Sub


' Keyboard Input

Sub FastDraw_KeyDown(ByVal keycode)
  If keycode = PlungerKey Then
    Plunger.PullBack
    PlaySoundAtVol "plungerpull", Plunger, 1
  End If

  if GameStarted and (Tilt = 0) Then
    If keycode = LeftFlipperKey Then
      LeftFlipper.RotateToEnd
      PlaySoundAtVol SoundFXDOF("flipperup_akiles",101,DOFOn,DOFContactors), LeftFlipper, 1
      PlayLoopSoundAtVol "FlipBuzzLA", LeftFlipper, 1
    End If

    If keycode = RightFlipperKey Then
      RightFlipper.RotateToEnd
      PlaySoundAtVol SoundFXDOF("flipperup_akiles",102,DOFOn,DOFContactors), RightFlipper, 1
     PlayLoopSoundAtVol "FlipBuzzRA", RightFlipper, 1
    End If
  end If

  If keycode = LeftTiltKey Then
    Nudge 90, 2
    checkNudge
  End If

  If keycode = RightTiltKey Then
    Nudge 270, 2
    checkNudge
  End If

  If keycode = CenterTiltKey Then
    Nudge 0, 2
    checkNudge
  End If

  if keycode = StartGameKey Then
    StartGame
  end If

  if keycode = AddCreditKey Then
    PlaysoundAtVol "Coin3", Drain, 1
    Credits = Credits + 1
    DOF 126, DOFOn
    if Credits > 9 then
      Credits = 9
    end If
    if B2Son Then
      for i=50 to 59
        Controller.B2SSetData i,0
      next

      Controller.B2SSetData 50+Credits,1
    end If
    P_Credits.image = cstr(Credits)
    EMReel_Credits.setvalue Credits
  end If

  if keycode = AddCreditKey2 Then
    PlaysoundAtVol "Coin3", Drain, 1
    Credits = Credits + 2
    DOF 126, DOFOn
    if Credits > 9 then
      Credits = 9
    end If
    if B2Son Then
      for i=50 to 59
        Controller.B2SSetData i,0
      next

      Controller.B2SSetData 50+Credits,1
    end If
    P_Credits.image = cstr(Credits)
    EMReel_Credits.setvalue Credits
  end If
End Sub

Sub FastDraw_KeyUp(ByVal keycode)
  If keycode = PlungerKey Then
    Plunger.Fire
    PlaySoundAtVol "plunger", Plunger, 1
  End If

  If keycode = LeftFlipperKey Then
    LeftFlipper.RotateToStart
    if GameStarted and (Tilt = 0) Then
      PlaySoundAtVol SoundFXDOF("flipperdown_akiles",101,DOFOff,DOFContactors), LeftFlipper, 1
      StopSound "FlipBuzzLA"
    end If
  End If

  If keycode = RightFlipperKey Then
    RightFlipper.RotateToStart
    if GameStarted and (Tilt = 0) Then
      PlaySoundAtVol SoundFXDOF("flipperdown_akiles",102,DOFOff,DOFContactors), RightFlipper, 1
      StopSound "FlipBuzzRA"
    end If
  End If
End Sub

Sub AddSpecial
  playsound SoundFXDOF("knocker",125,DOFPulse,DOFKnocker)
  Credits = Credits + 1
  DOF 126, DOFOn
  if Credits > 9 then
    Credits = 9
  end If
  if B2Son Then
    for i=50 to 59
      Controller.B2SSetData i,0
    next
    Controller.B2SSetData 50+Credits,1
  end If
  EMReel_Credits.setvalue Credits
End Sub

'*** CollectBonus L and R Kicker if lit
Sub CollectBonus

  if Tilt = 1 Then
    TotalBonus = 0
  end If
  BonusX = 1
  if DB then BonusX = 2
  if TB then BonusX = 3
  CollectBonusTimer.enabled = True
  StoreBonus = TotalBonus 'remember total to restore to playfield when collect bonus has count
end Sub

Sub CollectBonusTimer_Timer

  if TotalBonus > 0 Then
    Playsound "FD_BonusCount"
    Addpoints(1000 * BonusX)
    TotalBonus = TotalBonus - 1
    UpdateBonusLights
  Else
    CollectBonusTimer.enabled = False
    TotalBonus = StoreBonus
    UpdateBonusLights
    Playsound "FD_BonusCount"
    StoreBonus = 0
  end If

End Sub

'--- Tilt recognition ---
Dim Tilt
Sub CheckNudge
  if GameActive then
    if NudgeTimer1.enabled then
      if NudgeTimer2.enabled then
        NudgeTimer1.enabled = False
        NudgeTimer2.enabled = False
        if Tilt = 0 then
          GameTilted
        end if
      else
        NudgeTimer2.enabled = True
      end if
    else
      NudgeTimer1.enabled = True
    end if
  end if
End Sub

Sub NudgeTimer1_Timer
  NudgeTimer1.enabled = False
End Sub

Sub NudgeTimer2_Timer
  NudgeTimer2.enabled = False
End Sub

Sub GameTilted
  AdvanceBonusTimer.enabled = False
  Tilt = 1

    if B2SOn then
    Controller.B2SSetTilt 1
        StopSound "FlipBuzzLA"
        StopSound "FlipBuzzRA"
        Playsound SoundFXDOF("knocker",125,DOFPulse,DOFKnocker)
  end If

  Left500a.state = Lightstateoff: Left500a2nd.state = Lightstateoff
  Left500b.state = Lightstateoff: Left500b2nd.state = Lightstateoff
  Right500a.state = Lightstateoff: Right500a2nd.state = Lightstateoff
  Right500b.state = Lightstateoff: Right500b2nd.state = Lightstateoff
  LeftDTWL.state = Lightstateoff: LeftDTWL2nd.state = Lightstateoff
  RightDTWL.state = Lightstateoff: RightDTWL2nd.state = Lightstateoff
  b1l.state = Lightstateoff
  b2l.state = Lightstateoff
  b3l.state = Lightstateoff
  TargetSound = "target"
  DTSound = "fx_droptarget"
  RolloverSound1 = "Soloff"
  RolloverSound2 = "Soloff"
  OutlaneSound = "Soloff"
  Soundlevel = 0.3

  BumperWall1.isdropped = False
  BumperWall2.isdropped = False
  BumperWall3.isdropped = False
End Sub

'#############################################################
'#####  Fast Draw Scoring                                #####
'#############################################################
Dim BallinPlay

Dim HoleValue
Sub LKicker_Hit
  PlaySound "kicker_enter_center", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(LKicker)
  LKicker.timerenabled = False
  Select Case HoleValue
    Case 0: Playsound "FD_Hole_0",0,1,-0.07,0.05
    Case 1: Playsound "FD_Hole_1",0,1,-0.07,0.05
    Case 2: Playsound "FD_Hole_2",0,1,-0.07,0.05
    Case 3: Playsound "FD_Hole_3",0,1,-0.07,0.05
  end Select
  LKicker.timerenabled = True
  AddLeftHolePointsTimer.enabled = True
End Sub

Sub AddLeftHolePointsTimer_Timer
  AddLeftHolePointsTimer.enabled = False
  if Not LeftSpecial Then
    AddPoints(1000)
    HoleValue = 0
    if LightA Then AddPoints(1000):HoleValue = HoleValue + 1:End If
    if LightB Then AddPoints(1000):HoleValue = HoleValue + 1:End If
    if LightC Then AddPoints(1000):HoleValue = HoleValue + 1:End If
    ScoreMotor
'   if LeftSpecial Then
  Else
    LeftSpecial = False
    LeftHoleSpecial.state = Lightstateoff: LeftHoleSpecial2nd.state = Lightstateoff
    CollectBonus
    'ScoreMotor
  end If
End Sub

Sub LKicker_Timer
  LKicker.timerenabled = False
  LKicker.kick 43,23,0.85
  P_RightKicker.rotx = -85
  P_LeftKicker.rotx = -85
  Kickertimer.enabled = False
  Kickertimer.enabled = True
  DOF 117, DOFPulse
     if Tilt = 0 then
     Dim x
     x = INT(2 * RND(1) )
     Select Case x
    Case 0: Playsound "Gun1",0,1,0.07,0.05
    Case 1: Playsound "Gun2",0,1,0.07,0.05
    end Select
     end If
End Sub

Sub RKicker_Hit
  'playsound "kicker_enter_center",0,0.1,0.12,0.05
  PlaySound "kicker_enter_center", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(RKicker)
  RKicker.timerenabled = False
    Select Case HoleValue
    Case 0: Playsound "FD_Hole_0",0,1,0.07,0.05
    Case 1: Playsound "FD_Hole_1",0,1,0.07,0.05
    Case 2: Playsound "FD_Hole_2",0,1,0.07,0.05
    Case 3: Playsound "FD_Hole_3",0,1,0.07,0.05
  end Select
  RKicker.timerenabled = True
  AddRightHolePointsTimer.enabled = True
End Sub

Sub AddRightHolePointsTimer_Timer
  AddRightHolePointsTimer.enabled = False
  if Not RightSpecial Then
    AddPoints(1000)
    HoleValue = 0
    if LightA Then AddPoints(1000):HoleValue = HoleValue + 1:End If
    if LightB Then AddPoints(1000):HoleValue = HoleValue + 1:End If
    if LightC Then AddPoints(1000):HoleValue = HoleValue + 1:End If
    ScoreMotor
  Else
    RightSpecial = False
    RightHoleSpecial.state = Lightstateoff: RightHoleSpecial2nd.state = Lightstateoff
    CollectBonus
    'ScoreMotor
  end If
End Sub

Sub RKicker_Timer
  RKicker.timerenabled = False
  RKicker.kick -43,23,0.85
  P_RightKicker.rotx = -85
  P_LeftKicker.rotx = -85
  Kickertimer.enabled = False
  Kickertimer.enabled = True
  DOF 118, DOFPulse
     if Tilt = 0 then
      Dim x
      x = INT(2 * RND(1) )
      Select Case x
     Case 0:  Playsound "Gun1",0,1,0.07,0.05
     Case 1:  Playsound "Gun2",0,1,0.07,0.05
    end Select
     end If
End Sub

Sub KickerTimer_Timer
  P_RightKicker.rotx = P_RightKicker.rotx + 5
  P_LeftKicker.rotx = P_LeftKicker.rotx + 5

  if P_RightKicker.rotx >= -10 then
    P_RightKicker.rotx = -10
    P_LeftKicker.rotx = -10
    Kickertimer.enabled = False
  end if
End Sub

Sub Bumper1_Hit
  if LightA Then
    Addpoints(100)
  Else
    Addpoints(10)
  end If
  PlaySoundAtVol SoundFXDOF("FD_Bumper1",103,DOFPulse,DOFContactors), Bumper1, 1
End Sub

Sub Bumper2_Hit
  if LightC Then
    Addpoints(100)
  Else
    Addpoints(10)
  end If
  PlaySoundAtVol SoundFXDOF("FD_Bumper2",105,DOFPulse,DOFContactors), Bumper2, 1
End Sub

Sub Bumper3_Hit
  if LightB Then
    Addpoints(1000)
  Else
    Addpoints(10)
  end If
  PlaySoundAtVol SoundFXDOF("FD_Bumper3",104,DOFPulse,DOFContactors), Bumper3, 1
  if LeftSpecial Then
    RightSpecial = True
    RightHoleSpecial.state = Lightstateon: RightHoleSpecial2nd.state = Lightstateon
    LeftSpecial = False
    LeftHoleSpecial.state = Lightstateoff: LeftHoleSpecial2nd.state = Lightstateoff
  Else
    if RightSpecial Then
      RightSpecial = False
      RightHoleSpecial.state = Lightstateoff: RightHoleSpecial2nd.state = Lightstateoff
      LeftSpecial = True
      LeftHoleSpecial.state = Lightstateon: LeftHoleSpecial2nd.state = Lightstateon
    end If
  end If
End Sub

' Sub Trigger_A1_Hit:PlaySound RolloverSound1,0,Soundlevel,-0.05,0.05:Addpoints(50):Light_A:ScoreMotor:DOF 108, DOFPulse:End Sub
' Sub Trigger_A2_Hit:Playsound RolloverSound2,0,Soundlevel,-0.09,0.05:Addpoints(50):Light_A:ScoreMotor:DOF 112, DOFPulse:End Sub
' Sub Trigger_B1_Hit:Playsound RolloverSound1,0,Soundlevel,0,0.05:Addpoints(50):Light_B:ScoreMotor:DOF 109, DOFPulse:End Sub
Sub Trigger_A1_Hit:PlaySoundAtVol RolloverSound1,Trigger_A1,Soundlevel:Addpoints(50):Light_A:ScoreMotor:DOF 108, DOFPulse:End Sub
Sub Trigger_A2_Hit:PlaysoundAtVol RolloverSound2,Trigger_A2,Soundlevel:Addpoints(50):Light_A:ScoreMotor:DOF 112, DOFPulse:End Sub
Sub Trigger_B1_Hit:PlaysoundAtVol RolloverSound1,Trigger_B1,Soundlevel:Addpoints(50):Light_B:ScoreMotor:DOF 109, DOFPulse:End Sub
Sub Trigger_B2_Hit:P_LeftInlane.TransZ = -25:PlaysoundAtVol RolloverSound2,Trigger_b2,Soundlevel:Addpoints(50):Light_B:ScoreMotor:DOF 113, DOFOn:End Sub
Sub Trigger_B2_Unhit:Trigger_B2.timerenabled = True:DOF 113, DOFOff:End Sub
Sub Trigger_B2_Timer:Trigger_B2.timerenabled = False:P_LeftInlane.TransZ = 0:End Sub
Sub Trigger_B3_Hit:P_RightInlane.TransZ = -25:PlaysoundAtVol RolloverSound2,Trigger_b3,Soundlevel:Addpoints(50):Light_B:ScoreMotor:DOF 115, DOFOn:End Sub
Sub Trigger_B3_Unhit:Trigger_B3.timerenabled = True:DOF 115, DOFOff:End Sub
Sub Trigger_B3_Timer:Trigger_B3.timerenabled = False:P_RightInlane.TransZ = 0:End Sub

Sub Trigger_C1_Hit:PlaysoundAtVol RolloverSound1,Trigger_C1,Soundlevel:Addpoints(50):Light_C:ScoreMotor:DOF 110, DOFPulse:End Sub
Sub Trigger_C2_Hit:PlaysoundAtVol RolloverSound2,Trigger_C2,Soundlevel:Addpoints(50):Light_C:ScoreMotor:DOF 114, DOFPulse:End Sub

Sub Trigger_LeftOutlane_Hit:P_LeftOutlane.TransZ = -25:PlaySoundAtVol OutlaneSound,LeftOutLane,Soundlevel:AddBonus:Addpoints(500):ScoreMotor:DOF 111, DOFOn:End Sub
Sub Trigger_LeftOutlane_Unhit:Trigger_LeftOutlane.timerenabled = True:DOF 111, DOFOff:End Sub
Sub Trigger_LeftOutlane_Timer:Trigger_LeftOutlane.timerenabled = False:P_LeftOutlane.TransZ = 0:End Sub
Sub Trigger_RightOutlane_Hit:P_RightOutlane.TransZ = -25:PlaySoundAtVol OutlaneSound,Trigger_RightOutlane,Soundlevel:AddBonus:Addpoints(500):ScoreMotor:DOF 116, DOFOn:End Sub
Sub Trigger_RightOutlane_Unhit:Trigger_RightOutlane.timerenabled = True:DOF 116, DOFOff:End Sub
Sub Trigger_RightOutlane_Timer:Trigger_RightOutlane.timerenabled = False:P_RightOutlane.TransZ = 0:End Sub

Sub Target11_Hit:PlaySoundAtVol SoundFXDOF(TargetSound,106,DOFPulse,DOFContactors),Target11,Soundlevel:AddBonus:Addpoints(500):ScoreMotor:End Sub
Sub Target12_Hit:PlaySoundAtVol SoundFXDOF(TargetSound,106,DOFPulse,DOFContactors),Target12,Soundlevel:AddBonus:Addpoints(500):ScoreMotor:End Sub
Sub Target13_Hit:PlaySoundAtVol SoundFXDOF(TargetSound,107,DOFPulse,DOFContactors),Target13,Soundlevel:AddBonus:Addpoints(500):ScoreMotor:End Sub
Sub Target14_Hit:PlaySoundAtVol SoundFXDOF(TargetSound,107,DOFPulse,DOFContactors),Target14,Soundlevel:AddBonus:Addpoints(500):ScoreMotor:End Sub

'Drop Targets
Sub Target1_Hit:LightDTL1.State = 1:playsoundAtVol DTSound,Target1,Soundlevel:AddBonus:Addpoints(500):ScoreMotor:End Sub
Sub Target2_Hit:LightDTL2.State = 1:playsoundAtVol DTSound,Target2,Soundlevel:AddBonus:Addpoints(500):ScoreMotor:End Sub
Sub Target3_Hit:LightDTL3.State = 1
  playsoundAtVol DTSound,Target3,Soundlevel
  AddBonus
  if DTSpecial Then
    Addpoints(5000)
    DT5kHits = DT5kHits + 1
    Switch_ShootAgain
  Else
    Addpoints(500)
  end If
  ScoreMotor
End Sub
Sub Target4_Hit:LightDTL4.State = 1:playsoundAtVol DTSound,Target4,Soundlevel:AddBonus:Addpoints(500):ScoreMotor:End Sub
Sub Target5_Hit:LightDTL5.State = 1:playsoundAtVol DTSound,Target5,Soundlevel:AddBonus:Addpoints(500):ScoreMotor:End Sub
Sub Target6_Hit:LightDTL6.State = 1:playsoundAtVol DTSound,Target6,Soundlevel:AddBonus:Addpoints(500):ScoreMotor:End Sub
Sub Target7_Hit:LightDTL7.State = 1:playsoundAtVol DTSound,Target7,Soundlevel:AddBonus:Addpoints(500):ScoreMotor:End Sub
Sub Target8_Hit:LightDTL8.State = 1:
  playsoundAtVol DTSound,Target8,Soundlevel
  AddBonus
  if DTSpecial Then
    Addpoints(5000)
    DT5kHits = DT5kHits + 1
    Switch_ShootAgain
  Else
    Addpoints(500)
  end If
  ScoreMotor
End Sub
Sub Target9_Hit:LightDTL9.State = 1:playsoundAtVol DTSound,Target9,Soundlevel:AddBonus:Addpoints(500):ScoreMotor:End Sub
Sub Target10_Hit:LightDTL10.State = 1:playsoundAtVol DTSound,Target10,Soundlevel:AddBonus:Addpoints(500):ScoreMotor:End Sub

Dim LightA,LightB,LightC
Sub Light_A
  if not LightA and (Tilt = 0) Then
    LightA = True
    TopA.state = Lightstateoff: TopA2nd.state = Lightstateoff
    CenterA.state = Lightstateon: CenterA2nd.state = Lightstateon
    LowerA.state = Lightstateoff: LowerA2nd.state = Lightstateoff
    b1l.state = Lightstateon
    if LightA and LightB and LightC Then
      ExtraBonus2x.state = Lightstateon
            ExtraBonus2x2nd.state = Lightstateon
      DB = True
      if BallinPlay = BallsperGame Then
        TB = True
        ExtraBonus3x.state = Lightstateon
        ExtraBonus3x2nd.state = Lightstateon
      end If
    end If
  end If
End Sub

Sub Light_B
  if not LightB and (Tilt = 0) Then
    LightB = True
    TopB.state = Lightstateoff: TopB2nd.state = Lightstateoff
    CenterB.state = Lightstateon: CenterB2nd.state = Lightstateon
    LowerB1.state = Lightstateoff: LowerB2nd1.state = Lightstateoff
    LowerB2.state = Lightstateoff: LowerB2nd2.state = Lightstateoff
    b3l.state = Lightstateblinking
    if LightA and LightB and LightC Then
      ExtraBonus2x.state = Lightstateon: ExtraBonus2x2nd.state = Lightstateon
      DB = True
      if BallinPlay = BallsperGame Then
        TB = True
        ExtraBonus3x.state = Lightstateon
        ExtraBonus3x2nd.state = Lightstateon
      end If
    end If
  end If

End Sub

Sub Light_C
  if not LightC and (Tilt = 0) Then
    LightC = True
    TopC.state = Lightstateoff: TopC2nd.state = Lightstateoff
    CenterC.state = Lightstateon: CenterC2nd.state = Lightstateon
    LowerC.state = Lightstateoff: LowerC2nd.state = Lightstateoff
    b2l.state = Lightstateon
    if LightA and LightB and LightC Then
      ExtraBonus2x.state = Lightstateon: ExtraBonus2x2nd.state = Lightstateon
      DB = True
      if BallinPlay = BallsperGame Then
        TB = True
        ExtraBonus3x.state = Lightstateon
        ExtraBonus3x2nd.state = Lightstateon
      end If
    end If
  end If
End Sub

Sub CheckDTStateTimer_Timer
  if Tilt = 0 Then
    if LightA and LightB And LightC And not SpecialActivated Then
      if Target1.isdropped And Target2.isdropped And Target3.isdropped And Target4.isdropped And Target5.isdropped Then
        SpecialActivated = True
        RightSpecial = True
        RightHoleSpecial.state = Lightstateon: RightHoleSpecial2nd.state = Lightstateon
      end If
      if Target6.isdropped And Target7.isdropped And Target8.isdropped And Target9.isdropped And Target10.isdropped Then
        SpecialActivated = True
        LeftSpecial = True
        LeftHoleSpecial.state = Lightstateon: LeftHoleSpecial2nd.state = Lightstateon
      end If
    end If
    if Target1.isdropped And Target2.isdropped And Target3.isdropped And Target4.isdropped And Target5.isdropped And Target6.isdropped And Target7.isdropped And Target8.isdropped And Target9.isdropped And Target10.isdropped Then
      DTSpecial = True

'*** add / check Extra Ball ***

      if ActivePlayer = 1 then ExtraBallP1 = ExtraBallP1 + 1: end If
      if ActivePlayer = 2 then ExtraBallP2 = ExtraBallP2 + 1: end If
      if ActivePlayer = 3 then ExtraBallP3 = ExtraBallP3 + 1: end If
      if ActivePlayer = 4 then ExtraBallP4 = ExtraBallP4 + 1: end If

      if ExtraBallP1 = 1 then ExtraBall.State = 2: ExtraBall2nd.State = 2: ExtraBall2.State = 2: ExtraBall22nd.State = 2: end If
      if ExtraBallP2 = 1 then ExtraBall.State = 2: ExtraBall2nd.State = 2: ExtraBall2.State = 2: ExtraBall22nd.State = 2: end If
      if ExtraBallP3 = 1 then ExtraBall.State = 2: ExtraBall2nd.State = 2: ExtraBall2.State = 2: ExtraBall22nd.State = 2: end If
      if ExtraBallP4 = 1 then ExtraBall.State = 2: ExtraBall2nd.State = 2: ExtraBall2.State = 2: ExtraBall22nd.State = 2: end If

      'ResetLeftDT
      LeftDTBank_Cover.isdropped = False
      Target1.isdropped = False:LightDTL1.State = 0
      Target2.isdropped = False:LightDTL2.State = 0
      Target3.isdropped = False:LightDTL3.State = 0
      Target4.isdropped = False:LightDTL4.State = 0
      Target5.isdropped = False:LightDTL5.State = 0
      'ResetRightDT
      RightDTBank_Cover.isdropped = False
      Target6.isdropped = False:LightDTL6.State = 0
      Target7.isdropped = False:LightDTL7.State = 0
      Target8.isdropped = False:LightDTL8.State = 0
      Target9.isdropped = False:LightDTL9.State = 0
      Target10.isdropped = False:LightDTL10.State = 0

'*** added solenoid call when DT Targets fired up ***
  Playsound SoundFXDOF("FD_DT5KwithReset",128,DOFPulse,DOFContactors)

      Reset5KTimer.enabled = True

      LeftDTWL.state = Lightstateoff: LeftDTWL2nd.state = Lightstateoff
      Left5K.state = Lightstateon: Left5K2nd.state = Lightstateon
      RightDTWL.state = Lightstateoff: RightDTWL2nd.state = Lightstateoff
      Right5K.state = Lightstateon: Right5K2nd.state = Lightstateon

    end if
  end If
End Sub

Sub Reset5KTimer_Timer

  Reset5KTimer.enabled = False

'*** added solenoid call when DT Targets fired up ***
  Playsound SoundFXDOF("FD_DT5KwithReset",128,DOFPulse,DOFContactors)

  'DropLeftDT
  LeftDTBank_Cover.isdropped = True
  Target1.isdropped = True:LightDTL1.State = 1
  Target2.isdropped = True:LightDTL2.State = 1
  LightDTL3.State = 0
  Target4.isdropped = True:LightDTL4.State = 1
  Target5.isdropped = True:LightDTL5.State = 1

  'DropRightDT
  RightDTBank_Cover.isdropped = True
  Target6.isdropped = True:LightDTL6.State = 1
  Target7.isdropped = True:LightDTL7.State = 1
  LightDTL8.State = 0
  Target9.isdropped = True:LightDTL9.State = 1
  Target10.isdropped = True:LightDTL10.State = 1

End Sub

Dim DTSpecial,LeftSpecial,RightSpecial,SpecialActivated
'DTSpecial = 5K Drop Targets
'LeftSpecial = Left Hole Special
'RightSpecial = Right Hole Special

Sub ResetPlayfield

  DOF 128, DOFPulse  ' DOF Bumpers middle left / right for upfiring drop targets

  DT5kHits = 0

  LightA = False
  LightB = False
  LightC = False

  TopA.state = Lightstateon: TopA2nd.state = Lightstateon
  TopB.state = Lightstateon: TopB2nd.state = Lightstateon
  TopC.state = Lightstateon: TopC2nd.state = Lightstateon
  CenterA.state = Lightstateoff: CenterA2nd.state = Lightstateoff
  CenterB.state = Lightstateoff: CenterB2nd.state = Lightstateoff
  CenterC.state = Lightstateoff: CenterC2nd.state = Lightstateoff
  LowerA.state = Lightstateon: LowerA2nd.state = Lightstateon
  LowerB1.state = Lightstateon: LowerB2nd1.state = Lightstateon
  LowerB2.state = Lightstateon: LowerB2nd2.state = Lightstateon
  LowerC.state = Lightstateon: LowerC2nd.state = Lightstateon

  Left500a.state = Lightstateon: Left500a2nd.state = Lightstateon
  Left500b.state = Lightstateon: Left500b2nd.state = Lightstateon
  Right500a.state = Lightstateon: Right500a2nd.state = Lightstateon
  Right500b.state = Lightstateon: Right500b2nd.state = Lightstateon

  LeftHoleSpecial.state = Lightstateoff: LeftHoleSpecial2nd.state = Lightstateoff
  RightHoleSpecial.state = Lightstateoff: RightHoleSpecial2nd.state = Lightstateoff

  LeftDTWL.state = Lightstateon: LeftDTWL2nd.state = Lightstateon
  Left5K.state = Lightstateoff: Left5K2nd.state = Lightstateoff
  RightDTWL.state = Lightstateon: RightDTWL2nd.state = Lightstateon
  Right5K.state = Lightstateoff: Right5K2nd.state = Lightstateoff

  B1.state = Lightstateoff: B12nd.state = Lightstateoff
  B2.state = Lightstateoff: B22nd.state = Lightstateoff
  B3.state = Lightstateoff: B32nd.state = Lightstateoff
  B4.state = Lightstateoff: B42nd.state = Lightstateoff
  B5.state = Lightstateoff: B52nd.state = Lightstateoff
  B6.state = Lightstateoff: B62nd.state = Lightstateoff
  B7.state = Lightstateoff: B72nd.state = Lightstateoff
  B8.state = Lightstateoff: B82nd.state = Lightstateoff
  B9.state = Lightstateoff: B92nd.state = Lightstateoff
  B10.state = Lightstateoff: B102nd.state = Lightstateoff

' ExtraBonus.state = Lightstateoff: ExtraBonus2nd.state = Lightstateoff
  ExtraBonus2x.state = Lightstateoff: ExtraBonus2x2nd.state = Lightstateoff
  ExtraBonus3x.state = Lightstateoff: ExtraBonus3x2nd.state = Lightstateoff

  b1l.state = Lightstateoff
  b2l.state = Lightstateoff
  b3l.state = Lightstateoff

  'ResetLeftDT
  Target1.isdropped = False
  Target2.isdropped = False
  Target3.isdropped = False
  Target4.isdropped = False
  Target5.isdropped = False
  'ResetRightDT
  Target6.isdropped = False
  Target7.isdropped = False
  Target8.isdropped = False
  Target9.isdropped = False
  Target10.isdropped = False
  'Reset DropTarget Lights
  LightDTL1.State = 0
  LightDTL2.State = 0
  LightDTL3.State = 0
  LightDTL4.State = 0
  LightDTL5.State = 0
  LightDTL6.State = 0
  LightDTL7.State = 0
  LightDTL8.State = 0
  LightDTL9.State = 0
  LightDTL10.State = 0

  DB = False
  TB = False
  TotalBonus = 0
  DTSpecial = False
  LeftSpecial = False
  RightSpecial = False
  SpecialActivated = False
  Tilt = 0
  BumperWall1.isdropped = True
  BumperWall2.isdropped = True
  BumperWall3.isdropped = True

  TargetSound = "FD_500Target"
  DTSound = "FD_DropTarget"
  RolloverSound1 = "FD_Top_Rollover"
  RolloverSound2 = "FD_Top_Rollover2"
  OutlaneSound = "FD_Outlane"
  Soundlevel = 1
End Sub

Dim TotalBonus,DB,TB
'DB = Double Bonus
'TB = Triple Bonus (DB on last Ball)

Sub AddBonus
  if Tilt = 0 Then
    if not AdvanceBonusTimer.enabled Then
      if TotalBonus < 55 then
        TotalBonus = TotalBonus + 1
      end If
      UpdateBonusLights
    end If
  end If
End Sub

Dim TargetSound,DTSound,RolloverSound1,RolloverSound2,OutlaneSound,Soundlevel
Sub ScoreMotor
  if Tilt = 0 Then
    AdvanceBonusTimer.enabled = True
    Left500a.state = Lightstateoff: Left500a2nd.state = Lightstateoff
    Left500b.state = Lightstateoff: Left500b2nd.state = Lightstateoff
    Right500a.state = Lightstateoff: Right500a2nd.state = Lightstateoff
    Right500b.state = Lightstateoff: Right500b2nd.state = Lightstateoff
    LeftDTWL.state = Lightstateoff: LeftDTWL2nd.state = Lightstateoff
    RightDTWL.state = Lightstateoff: RightDTWL2nd.state = Lightstateoff
    b1l.state = Lightstateoff
    b2l.state = Lightstateoff
    b3l.state = Lightstateoff
    TargetSound = "target"
    DTSound = "TargetDrop1"
    RolloverSound1 = "Soloff"
    RolloverSound2 = "Soloff"
    OutlaneSound = "Soloff"
    Soundlevel = 0.3
  end If
End Sub

Sub AdvanceBonusTimer_Timer
  AdvanceBonusTimer.enabled = False
  Left500a.state = Lightstateon: Left500a2nd.state = Lightstateon
  Left500b.state = Lightstateon: Left500b2nd.state = Lightstateon
  Right500a.state = Lightstateon: Right500a2nd.state = Lightstateon
  Right500b.state = Lightstateon: Right500b2nd.state = Lightstateon
  if not DTSpecial Then
    LeftDTWL.state = Lightstateon: LeftDTWL2nd.state = Lightstateon
    RightDTWL.state = Lightstateon: RightDTWL2nd.state = Lightstateon
  end If
  if LightA Then
    b1l.state = Lightstateon
  end If
  if LightC Then
    b2l.state = Lightstateon
  end If
  if LightB Then
    b3l.state = Lightstateblinking
  end If
  TargetSound = "FD_500Target"
  DTSound = "fx_droptarget"
  RolloverSound1 = "FD_Top_Rollover"
  RolloverSound2 = "FD_Top_Rollover2"
  OutlaneSound = "FD_Outlane"
  Soundlevel = 1
End Sub

Dim BonusX
Sub AwardBonus
  if Tilt = 1 Then
    TotalBonus = 0
  end If
  BonusX = 1
  if DB then BonusX = 2
  if TB then BonusX = 3
  AwardBonusTimer.enabled = True
End Sub

Sub AwardBonusTimer_Timer
  if TotalBonus > 0 Then
    Playsound "FD_BonusCount"
    Addpoints(1000 * BonusX)
    TotalBonus = TotalBonus - 1
    UpdateBonusLights
  Else
    AwardBonusTimer.enabled = False
    Playsound "FD_BonusEnd"
    EndBounsTimer1.enabled = True
  end If
End Sub

Sub EndBounsTimer1_Timer
  EndBounsTimer1.enabled = False
' if (ActivePlayer < NoofPlayers) Or (BallinPlay < BallsperGame) or DT5kHits => 2 Then
  if (ActivePlayer < NoofPlayers) Or (BallinPlay < BallsperGame) or (GotExtraBallP1 = 1) or (GotExtraBallP2 = 1) or (GotExtraBallP3 = 1) or (GotExtraBallP4 = 1) Then
    DT5kHits = 0
    Playsound "FD_DrainwithBallrelease"
  end If
  EndBounsTimer2.enabled = True
End Sub

Sub EndBounsTimer2_Timer
  EndBounsTimer2.enabled = False
  NextBall
End Sub

Sub UpdateBonusLights
    B1.state = ((TotalBonus mod 10) = 1) or TotalBonus > 54
    B12nd.state = ((TotalBonus mod 10) = 1) or TotalBonus > 54
    B2.state = ((TotalBonus mod 10) = 2) or TotalBonus > 53
    B22nd.state = ((TotalBonus mod 10) = 2) or TotalBonus > 53
    B3.state = ((TotalBonus mod 10) = 3) or TotalBonus > 51
    B32nd.state = ((TotalBonus mod 10) = 3) or TotalBonus > 51
    B4.state = ((TotalBonus mod 10) = 4) or TotalBonus > 48
    B42nd.state = ((TotalBonus mod 10) = 4) or TotalBonus > 48
    B5.state = ((TotalBonus mod 10) = 5) or TotalBonus > 44
    B52nd.state = ((TotalBonus mod 10) = 5) or TotalBonus > 44
    B6.state = ((TotalBonus mod 10) = 6) or TotalBonus > 39
    B62nd.state = ((TotalBonus mod 10) = 6) or TotalBonus > 39
    B7.state = ((TotalBonus mod 10) = 7) or TotalBonus > 33
    B72nd.state = ((TotalBonus mod 10) = 7) or TotalBonus > 33
    B8.state = ((TotalBonus mod 10) = 8) or TotalBonus > 26
    B82nd.state = ((TotalBonus mod 10) = 8) or TotalBonus > 26
    B9.state = ((TotalBonus mod 10) = 9) or TotalBonus > 18
    B92nd.state = ((TotalBonus mod 10) = 9) or TotalBonus > 18
    B10.state = TotalBonus > 9
    B102nd.state = TotalBonus > 9
End Sub

Dim PlayerScore(4),ActivePlayer,Special1(4),Special2(4),Special3(4)
Sub Addpoints(ScorePar)
  if Tilt = 0 Then
    Nopointsscored = False
    if not GameActive then
      GameActive = True
    end If
    if ScorePar < 50 Then
      Playerscore(ActivePLayer) = Playerscore(ActivePLayer) + ScorePar
    Else
      if not AdvanceBonusTimer.enabled Then
        Playerscore(ActivePLayer) = Playerscore(ActivePLayer) + ScorePar
      end If
    end If

    if (Playerscore(ActivePLayer) >= Special1Score) and not Special1(ActivePlayer) Then
      Special1(ActivePlayer) = True
      AddSpecial
    end If
    if (Playerscore(ActivePLayer) >= Special2Score) and not Special2(ActivePlayer) Then
      Special2(ActivePlayer) = True
      AddSpecial
    end If
    if (Playerscore(ActivePLayer) >= Special3Score) and not Special3(ActivePlayer) Then
      Special3(ActivePlayer) = True
      AddSpecial
    end If

    UpdateScoreReel.enabled = True

  end If
End Sub

Sub RubberWall3_Hit:PlaysoundAtVol "rubber_hit_3",ActiveBall, 1:End Sub

Sub RubberSW1_Hit:Addpoints(10):Playsound "FD_10Switch",0,0.5,-0.08,0.05:End Sub
Sub RubberSW2_Hit:Addpoints(10):Playsound "FD_10Switch",0,0.5,0.08,0.05:End Sub
Sub RubberSW3_Hit:MoveSwitchA:Addpoints(10):Playsound "FD_10Switch",0,0.5,-0.08,0.05:End Sub
Sub RubberSW4_Hit:MoveSwitchB:Addpoints(10):Playsound "FD_10Switch",0,0.5,0.08,0.05:End Sub
Sub Gate1_Hit:PlaysoundAtVol "Gate5",gate1, 1:End Sub

Sub MoveSwitchA
  SwitchA_1.isdropped = True
  SwitchA_2.isdropped = False
  RubberSW3.timerenabled = False
  RubberSW3.timerenabled = True
End Sub

Sub RubberSW3_timer
  RubberSW3.timerenabled = False
  SwitchA_2.isdropped = True
  SwitchA_1.isdropped = False
End Sub

Sub MoveSwitchB
  SwitchB_1.isdropped = True
  SwitchB_2.isdropped = False
  RubberSW4.timerenabled = False
  RubberSW4.timerenabled = True
End Sub

Sub RubberSW4_timer
  RubberSW4.timerenabled = False
  SwitchB_2.isdropped = True
  SwitchB_1.isdropped = False
End Sub

Sub Flippertimer_Timer
  RFPrim.RotY = RightFlipper.currentangle
  LFPrim.RotY = LeftFlipper.currentangle
End Sub

Sub ShooterLaneLaunch_Hit
' if ActiveBall.vely < -8 then playsound "Launch",0,0.3,0.1,0.25
  if ActiveBall.vely < -8 then playsound "galopp_shooterlane",0,0.8
  DOF 124, DOFPulse
End Sub

'---------------------------
'----- High Score Code -----
'---------------------------
Const HighScoreFilename = "FastDrawLLEdition.txt"

Dim HSArray,HSAHighScore, HSA1, HSA2, HSA3
Dim HSScoreM,HSScore100k, HSScore10k, HSScoreK, HSScore100, HSScore10, HSScore1, HSScorex 'Define 6 different score values for each reel to use
HSArray = Array("HS_0","HS_1","HS_2","HS_3","HS_4","HS_5","HS_6","HS_7","HS_8","HS_9","HS_Space","HS_Comma","Tape2")
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
  end if
  if B2SOn Then
    Controller.B2SSetScore 6,HSAHighScore
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
      AddSpecial
      HSAHighScore=newScore
      SaveHighScore
      UpdatePostIt
    End If
End Sub

' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX, Rothbauerw and Herweh
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
'                     Supporting Ball & Sound Functions
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
  Vol = Csng(BallVel(ball) ^2 / VolDiv)
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
        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b))
        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If

    ' Ball Drop Sounds
    If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds, 27 standard
    PlaySound "fx_ball_drop" & b, 0, ABS(BOT(b).velz)/17, Pan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
    End If

  ' ball drop sounds from atlantis table for example
' If BOT(b).VelZ < -2 And BOT(b).Z < 55 And BOT(b).Z > 27 Then
'   PlaySound "fx_ball_drop" & Int(Rnd()*5), 0, ABS(BOT(b).VelZ)/17*5, Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b))
' End If

    ' Glass Hit Sound
' If BOT(b).VelZ < -1 And BOT(b).Z > 37 Then
'        PlaySound "AXSGlassHit" & Int(Rnd()*2)+1, 0, ABS(BOT(b).velz)/17, Pan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
' End If

    Next
End Sub


'************************************
' What you need to add to your table
'************************************

' a timer called RollingTimer. With a fast interval, like 10
' one collision sound, in this script is called fx_collide
' as many sound files as max number of balls, with names ending with 0, 1, 2, 3, etc
' for ex. as used in this script: fx_ballrolling0, fx_ballrolling1, fx_ballrolling2, fx_ballrolling3, etc


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


'**************************************
' Explanation of the collision routine
'**************************************

' The collision is built in VP.
' You only need to add a Sub OnBallBallCollision(ball1, ball2, velocity) and when two balls collide they
' will call this routine. What you add in the sub is up to you. As an example is a simple Playsound with volume and paning
' depending of the speed of the collision.


Sub Pins_Hit (idx)
  PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
  PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
  PlaySound "metalhit_thin", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
  PlaySound "metalhit_medium", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
  PlaySound "metalhit2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
  PlaySound "gate4", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinner_Spin
  PlaySoundAtVol "fx_spinner", Spinner, 1
End Sub

Sub Rubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 20 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 2 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End Sub

Sub Posts_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 2 AND finalspeed <= 16 then
    RandomSoundRubber()
  End If
End Sub

Sub RandomSoundRubber()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

Sub FastDraw_Exit()
  If B2SOn Then Controller.Stop
End Sub


' *****************************************************
' *** ball drop sounds / some special physics behaviour
' *****************************************************
' target or rubber post is hit so let the ball jump a bit
'Sub DropTargetHit()
' DropTargetSound
' TargetHit
'End Sub
'
'Sub StandupTargetHit()
' StandUpTargetSound
' TargetHit
'End Sub
'
'Sub TargetHit()
'    ActiveBall.VelZ = ActiveBall.VelZ * (0.5 + (Rnd()*LetTheBallJump + Rnd()*LetTheBallJump + 1) / 6)
'End Sub
'
'Sub RubberPostHit()
' ActiveBall.VelZ = ActiveBall.VelZ * (0.9 + (Rnd()*(LetTheBallJump-1) + Rnd()*(LetTheBallJump-1) + 1) / 6)
'End Sub
'
'Sub RubberRingHit()
' ActiveBall.VelZ = ActiveBall.VelZ * (0.8 + (Rnd()*(LetTheBallJump-1) + 1) / 6)
'End Sub


' *** ninuzzu's BALL SHADOW **************************
Dim BallShadow
BallShadow = Array (BallShadow1)

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
  If BOT(b).X < FastDraw.Width/2 Then
  BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (FastDraw.Width/2))/7)) + 13
  Else
  BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (FastDraw.Width/2))/7)) - 13
  End If
  ballShadow(b).Y = BOT(b).Y + 2
  If BOT(b).Z > 20 Then
  BallShadow(b).visible = 1
  Else
  BallShadow(b).visible = 0
  End If
  Next
End Sub

' *** ninuzzu's Flipper Shadow ************************
Sub FlipperShadowUpdate_Timer
  FlipperLSh.RotZ = LeftFlipper.CurrentAngle
  FlipperRSh.RotZ = RightFlipper.CurrentAngle
End Sub

Sub Switch_ShootAgain
  if DT5kHits = 2 then
  if ActivePlayer = 1 And ExtraBallP1 = 1 then GotExtraBallP1 = GotExtraBallP1 + 1: ExtraBallP1 = 2: ExtraBall.State = 0: ExtraBall2nd.State = 0: ExtraBall2.State = 0: ExtraBall22nd.State = 0: ShootAgain.State = 1: ShootAgain2nd.State = 1: end If
  if ActivePlayer = 2 And ExtraBallP2 = 1 then GotExtraBallP2 = GotExtraBallP2 + 1: ExtraBallP2 = 2: ExtraBall.State = 0: ExtraBall2nd.State = 0: ExtraBall2.State = 0: ExtraBall22nd.State = 0: ShootAgain.State = 1: ShootAgain2nd.State = 1: end If
  if ActivePlayer = 3 And ExtraBallP3 = 1 then GotExtraBallP3 = GotExtraBallP3 + 1: ExtraBallP3 = 2: ExtraBall.State = 0: ExtraBall2nd.State = 0: ExtraBall2.State = 0: ExtraBall22nd.State = 0: ShootAgain.State = 1: ShootAgain2nd.State = 1: end If
  if ActivePlayer = 4 And ExtraBallP4 = 1 then GotExtraBallP4 = GotExtraBallP4 + 1: ExtraBallP4 = 2: ExtraBall.State = 0: ExtraBall2nd.State = 0: ExtraBall2.State = 0: ExtraBall22nd.State = 0: ShootAgain.State = 1: ShootAgain2nd.State = 1: end If
  End If
End Sub

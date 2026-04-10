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
'### 29-03 v092
'###       script changes fix western music on / off shuts off all optional music
'### 05-2020
'###       script enhanced by scottacus: highscore feature 5players score and name entry
'
'a menu system with holding down the left flipper access to set freeplay, number of balls and music on/off
'added these variables to the save file so the player doesn't have to set them again
'fixed the ball getting stuck by the flippers, the flipper shadows are primitives and they were not set to toy and were collidible
'removed the collidibility and visibility for the nail by the left outlane
'added nfozzy's physics to the table rubbers so they are dampened
'changed the strength of the flippers and bumpers to match the period
'added nFozzy's flipper tricks to the table so now the flippers have the state of the art physics.

'### 07-2020 release v 0.95



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
BallMass = 1  'create ball in line 386

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

Const cGameName = "Fast_Draw_1975"
Const OptionsName = "LuckyLuke_1975"
Const hsFileName = "Lucky Luke (Gottlieb 1975 - FastDrawMod)"


'---------------------------
'-----  Configuration  -----
'---------------------------

'Const GameMusic = 0                 'ingame western music on (1) off (0)
'Const BallsperGame = 3       'the original table allows for 3 or 5 Balls
Const SetBallShadow = 1             'enable (1) disable (0)
Const ShowHighScore = True      'enable to show HighScore tape on Apron, a new HighScore will be saved anyway
Const ShowPlayfieldReel = False   'set to True to show Playfield EM-Reel and Post-It for Player Scores (instead of a B2S backglass)
'Const FreePlay = False
Const RollingSoundFactor = 3.0    'set sound level factor here for Ball Rolling Sound, 1=default level
Const ChimesEnabled = False     'You have been warned!
Const ChimeVolume = 0.3       'to limit ear bleeding set between 0 and 1
Const ResetHighscore = False    'enable to manually reset the Highscores
Const Special1Score = 98000     'set the 3 replay scores (3rd for display rollover)
Const Special2Score = 123000
Const Special3Score = 150000

Dim StoreBonus,GameActive,NoofPlayers,i,Credits,light,NoOfExtraBalls,XtraBall,ExtraBallP1,ExtraBallP2,ExtraBallP3,ExtraBallP4
Dim GotExtraBallP1,GotExtraBallP2,GotExtraBallP3,GotExtraBallP4,DT5kHits, tempPlayerScore

Dim hsInitial0, hsInitial1, hsInitial2, freePlay, gameMusic, ballsPerGame
Dim hsArray: hsArray = Array("HS_0","HS_1","HS_2","HS_3","HS_4","HS_5","HS_6","HS_7","HS_8","HS_9","HS_Space","HS_Comma")
Dim hsiArray: hsIArray = Array("HSi_0","HSi_1","HSi_2","HSi_3","HSi_4","HSi_5","HSi_6","HSi_7","HSi_8","HSi_9","HSi_10","HSi_11","HSi_12","HSi_13","HSi_14","HSi_15","HSi_16","HSi_17","HSi_18","HSi_19","HSi_20","HSi_21","HSi_22","HSi_23","HSi_24","HSi_25","HSi_26")



Sub FastDraw_Init

'*** turn gi lights on when game starts ***
  For each light in GI:light.State = 0: Next

'*** set lights and arrays
'Lights(1) = Array(Extraball,Extraball2)


  LoadEM

    if SetBallShadow = 0 then
         BallShadowUpdate.enabled=False
    End if

  loadHighScore

  if freePlay = "" Then freePlay = 1
  if gameMusic = "" Then gameMusic = 1
  if Credits = "" Then Credits = 0
  if ballsPerGame = "" Then ballsPerGame = 3
  If highScore(0) = "" Then highScore(0) = 1500
  If highScore(1) = "" Then highScore(1) = 1200
  If highScore(2) = "" Then highScore(2) = 1000
  If highScore(3) = "" Then highScore(3) = 800
  If HighScore(4) = "" Then highScore(4) = 600


If ballsPerGame = 3 Then
  InstructCard.image = "instcard_left_3ball"
Else
  InstructCard.image = "instcard_left_5ball"
End If

If freePlay = 0 Then
  instructCardR.image = "instcard_right"
Else
  instructCardR.image = "instcard_rightFP"
End If


  If initial(0,1) = "" Then
    initial(0,1) = 19: initial(0,2) = 5: initial(0,3) = 13
    initial(1,1) = 1: initial(1,2) = 1: initial(1,3) = 1
    initial(2,1) = 2: initial(2,2) = 2: initial(2,3) = 2
    initial(3,1) = 3: initial(3,2) = 3: initial(3,3) = 3
    initial(4,1) = 4: initial(4,2) = 4: initial(4,3) = 4
  End If

  updatePostIt
  dynamicUpdatePostIt.enabled = 1

  savehighscore

  Randomize

  if B2SOn then
    Controller.B2SSetMatch MatchValue
    Controller.B2SSetScoreRolloverPlayer1 0
    Controller.B2SSetScoreRolloverPlayer2 0
    Controller.B2SSetScoreRolloverPlayer3 0
    Controller.B2SSetScoreRolloverPlayer4 0
    Controller.B2SSetScore 6,HighScore(0)
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
  firstBallOut = 0
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

'***********Operator Menu
Dim operatormenu
Sub operatorMenuTimer_Timer
  options = 0
    operatorMenu = 1
  dynamicUpdatePostIt.enabled = 0
  updatePostIt
  options = 0
    optionsMenu.visible = True
    optionMenu.visible = True
  optionMenu.image = "FreeCoin" & freePlay
End Sub


Dim GameStarted,NoPointsScored
Sub StartGame
  if not GameActive then
    dynamicUpdatePostIt.enabled = 0
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
    updatePostIt
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

    if GameMusic = 1 then PlaySound "BallOut Pat Woods Start2b lonesome alone",0,0.5

      Playsound "FD_GameStartwithBallrelease"

      GameStartTimer.enabled = True
      ScoreMotor
      ScoreMotorStartTimer.enabled = True
    end If
    if NoofPlayers < 4 Then
      If Credits < 1 Then DOF 126, DOFOff
      if NoofPlayers > 0 Then
        Playsound "AddPlayer"
      end If
      NoofPlayers = NoofPlayers + 1
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
      Controller.B2SSetScore 6,highScore(0)
    End if
    EMReel_BiP.setvalue BallinPlay
    updatePostIt
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
    ContBallInPlay = 1

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
    ContBallInPlay = 1
  End If

End If

End Sub

Dim MatchValue

Sub EndGame

  '### GI ausschalten
  For each light in GI:light.State = 0: Next
  For each light in GI_DTargets:light.State = 0: Next

  EndMusic

  if GameMusic = 1 then PlayMusic "FD_6.mp3"

  DOF 20, DOFoff 'b2s illumination lamp 20 off

  '### GameMusic Timer on off
    if GameMusic = 1 then
         m01.enabled=False
    End if

  DynamicUpdatePostIt.enabled = 1
  Dim i
  sortScores
  checkHighScores
  firstBallOut = 0

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
  PlaySoundAt "drain3", Drain
  if GameMusic = 1 then PlaySound "Horse Neighing",0,1.0

DOF 121, DOFPulse
  Drain.DestroyBall
  ContBallInPlay = 0
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
  ContBallInPlay = 1
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
Dim Val10(4),Val100(4),Val1000(4),Val10000(4),Val100000(4),Score10,Score100,Score1000,Score10000,Score100000,Oldscore(5)
Sub UpdateScoreReel_Timer
  tempPlayerScore = Playerscore(ActivePLayer)
  Score10 = 0
  Score100 = 0
  Score1000 = 0
  Score10000 = 0
  Score100000 = 0
  if len(tempPlayerScore) > 1 Then
    tempPlayerScore = Left(tempPlayerScore,len(tempPlayerScore)-1)
    Score10 = cstr(right(tempPlayerScore,1))
  end If
  if len(tempPlayerScore) > 1 Then
    tempPlayerScore = Left(tempPlayerScore,len(tempPlayerScore)-1)
    Score100 = cstr(right(tempPlayerScore,1))
  end If
  if len(tempPlayerScore) > 1 Then
    tempPlayerScore = Left(tempPlayerScore,len(tempPlayerScore)-1)
    Score1000 = cstr(right(tempPlayerScore,1))
  end If
  if len(tempPlayerScore) > 1 Then
    tempPlayerScore = Left(tempPlayerScore,len(tempPlayerScore)-1)
    Score10000 = cstr(right(tempPlayerScore,1))
  end If
  if len(tempPlayerScore) > 1 Then
    tempPlayerScore = Left(tempPlayerScore,len(tempPlayerScore)-1)
    Score100000 = cstr(right(tempPlayerScore,1))
  end If

  Val10(ActivePLayer) = ReelValue(Val10(ActivePLayer),Score10,1)
  Val100(ActivePLayer) = ReelValue(Val100(ActivePLayer),Score100,2)
  Val1000(ActivePLayer) = ReelValue(Val1000(ActivePLayer),Score1000,3)
  Val10000(ActivePLayer) = ReelValue(Val10000(ActivePLayer),Score10000,0)
  Val100000(ActivePLayer) = ReelValue(Val100000(ActivePLayer),Score100000,0)
  tempPlayerScore = Val10(ActivePLayer) * 10 + Val100(ActivePLayer) * 100 + Val1000(ActivePLayer) * 1000 + Val10000(ActivePLayer) * 10000 + Val100000(ActivePLayer) * 100000
  if Oldscore(ActivePLayer) <> tempPlayerScore Then
    Oldscore(ActivePLayer) = tempPlayerScore
    select Case ActivePlayer
      case 1: EMReel1.setvalue tempPlayerScore
      case 2: EMReel2.setvalue tempPlayerScore
      case 3: EMReel3.setvalue tempPlayerScore
      case 4: EMReel4.setvalue tempPlayerScore
    End Select

    If B2SOn Then
      Controller.B2SSetScorePlayer ActivePlayer, tempPlayerScore


' *** added ScoreReel Override Check and Display (B2S DOF Lamp 25 ff)
'        if (Playerscore(ActivePLayer)) >= 10000 Then
'     DOF ((Playerscore(ActivePlayer))/10000)+24, DOFon
'       if (Playerscore(ActivePlayer)) > 20000 Then
'         DOF ((Playerscore(ActivePlayer))/10000)+23, DOFoff
'       end If
'   end If

    End If

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
  tempPlayerScore = Playerscore(ActivePLayer)
  Score10 = 0
  Score100 = 0
  Score1000 = 0
  Score10000 = 0
  Score100000 = 0
  if len(tempPlayerScore) > 1 Then
    tempPlayerScore = Left(tempPlayerScore,len(tempPlayerScore)-1)
    Score10 = cstr(right(tempPlayerScore,1))
  end If
  if len(tempPlayerScore) > 1 Then
    tempPlayerScore = Left(tempPlayerScore,len(tempPlayerScore)-1)
    Score100 = cstr(right(tempPlayerScore,1))
  end If
  if len(tempPlayerScore) > 1 Then
    tempPlayerScore = Left(tempPlayerScore,len(tempPlayerScore)-1)
    Score1000 = cstr(right(tempPlayerScore,1))
  end If
  if len(tempPlayerScore) > 1 Then
    tempPlayerScore = Left(tempPlayerScore,len(tempPlayerScore)-1)
    Score10000 = cstr(right(tempPlayerScore,1))
  end If
  if len(tempPlayerScore) > 1 Then
    tempPlayerScore = Left(tempPlayerScore,len(tempPlayerScore)-1)
    Score100000 = cstr(right(tempPlayerScore,1))
  end if
End Sub

'******* for ball control script
Sub endControl_Hit()
    contBallInPlay = False
End Sub

' Keyboard Input
Dim enableInitialEntry, options
Sub FastDraw_KeyDown(ByVal keycode)

  If enableInitialEntry = True Then enterIntitals(keycode)

  If keycode = PlungerKey Then
    Plunger.PullBack
    PlaySoundAt "plungerpull",ActiveBall
  End If

  if GameStarted and (Tilt = 0) and GameStarted = True and contball = 0 Then
    If keycode = LeftFlipperKey Then
      LF.fire 'LeftFlipper.RotateToEnd
      LFPress = 1
'     PlaySound SoundFXDOF("flipperup_akilesL",101,DOFOn,DOFContactors), 0, 0.1, -0.05, 0.25
      PlaySoundAtVol SoundFXDOF("flipperup_akilesL",101,DOFOn,DOFContactors), LeftFlipper, 1
            PlaySoundAt "FlipBuzzLA", ActiveBall
    End If

    If keycode = RightFlipperKey and GameStarted = True and contball = 0 Then
      RF.fire 'RightFlipper.RotateToEnd
      RFPress = 1
'     PlaySound SoundFXDOF("flipperup_akilesR",102,DOFOn,DOFContactors), 0, 0.1, 0.05, 0.25
      PlaySoundAt SoundFXDOF("flipperup_akilesR",102,DOFOn,DOFContactors), ActiveBall, 1
            PlaySoundAt "FlipBuzzRA",ActiveBall
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

  If keycode = MechanicalTilt Then
    checkNudge
  End If

  if keycode = StartGameKey Then
    If enableInitialEntry = False and operatormenu = 0 and HighScoreDelay.enabled = 0 Then
      If freePlay = 1 and noOfPlayers < 4 and firstBallOut = 0 Then startGame
      If freePlay = 0 and credits > 0 and noOfPlayers < 4 and firstBallOut = 0 Then
        credits = credits - 1
        EMReel_Credits.setvalue(credits)
        If B2SOn Then
          controller.B2SSetCredits credits
          if credits < 1 Then DOF 126, DOFOff
        End If
        startGame
      End If
    End If
  end If

    If keycode = leftFlipperKey and gameStarted = False and operatorMenu = 0 and enableInitialEntry = 0 Then
        operatorMenuTimer.Enabled = true
    End If

    If keycode = leftFlipperKey and gameStarted = False and operatorMenu = 1 Then
    options = options + 1
        If options > 3 Then options = 0
    optionMenu.visible = True
        playSound "target"
        Select Case (Options)
            Case 0:
                optionMenu.image = "FreeCoin" & freePlay
            Case 1:
                optionMenu.image = ballsperGame & "Balls"
      Case 2:
        optionMenu.image = "Music" & GameMusic
      Case 3:
        optionMenu.image = "SaveExit"
    End Select
    End If

    If keycode = rightFlipperKey and GameStarted = False and operatorMenu = 1 Then
        playSound "metalhit2"
      Select Case (options)
    Case 0:
            If freePlay = 0 Then
                freePlay = 1
        instructCardR.image = "instcard_rightFP"
              Else
                freePlay = 0
        instructCardR.image = "instcard_right"
            End If
            optionMenu.image= "FreeCoin" & freePlay
      If freePlay = 0 Then
        If credits > 0 and B2SOn Then DOF 126, DOFOn
        If credits < 1 and B2SOn Then DOF 126, DOFOff
      Else
        If B2SOn Then DOF 126, DOFOn
      End If
        Case 1:
            If ballsperGame = 3 Then
                ballsperGame = 5
        instructCard.image = "instcard_left_5ball"
'       coincard.image = "GT 5b coin"
              Else
                ballsperGame = 3
        instructCard.image = "instcard_left_3ball"
'       coincard.image = "GT 3b coin"
            End If
            optionMenu.image = ballsperGame & "Balls"
'     replaySettings
        Case 2:
            If gameMusic = 0 Then
                gameMusic= 1
        PlaySound "BallOut Pat Woods Start2b lonesome alone",0,0.5
              Else
                gameMusic = 0
        StopSound "BallOut Pat Woods Start2b lonesome alone"
            End If
      optionMenu.image = "Music" & gameMusic
        Case 3:
            operatorMenu = 0
            saveHighScore
      dynamicUpdatePostIt.enabled = 1
      optionMenu.image = "FreeCoin" & freePlay
      optionMenu.visible = 0
      optionsMenu.visible = 0
    End Select
      playSound "metalhit2"
    End If

  if keycode = AddCreditKey Then
    PlaySoundAt "Coin3", Drain, 1
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
    PlaySoundAt "Coin3", Drain, 1
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

End Sub

Sub FastDraw_KeyUp(ByVal keycode)
  If keycode = PlungerKey Then
    Plunger.Fire
    PlaySoundAt "plunger", Plunger
  End If

    If keycode = leftFlipperKey Then
        operatorMenuTimer.Enabled = False
    End If


  If keycode = LeftFlipperKey and GameStarted and contball = 0 Then
    LeftFlipper.RotateToStart
    lfpress = 0
    leftflipper.eostorqueangle = EOSA
    leftflipper.eostorque = EOST
    if GameStarted and (Tilt = 0) Then
      PlaySoundAt SoundFXDOF("flipperdown_akiles",101,DOFOff,DOFContactors), LeftFlipper
            StopSound "FlipBuzzLA"
    end If
  End If

  If keycode = RightFlipperKey and GameStarted and contball = 0 Then
    RightFlipper.RotateToStart
    rfpress = 0
    rightflipper.eostorqueangle = EOSA
    rightflipper.eostorque = EOST
    if GameStarted and (Tilt = 0) Then
      PlaySoundAt SoundFXDOF("flipperdown_akiles",102,DOFOff,DOFContactors), RightFlipper
            StopSound "FlipBuzzRA"
    end If
  End If

    If keycode = 203 then cLeft = 0' Left Arrow

    If keycode = 200 then cUp = 0' Up Arrow

    If keycode = 208 then cDown = 0' Down Arrow

    If keycode = 205 then cRight = 0' Right Arrow

    If keycode = 52 Then Zup = 0' Period
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
  PlaySoundAtBall "kicker_enter_center"
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
  PlaySoundAtBall "kicker_enter_center"
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
  PlaySoundAtVol SoundFXDOF("FD_Bumper1",103,DOFPulse,DOFContactors), ActiveBall, 1
End Sub

Sub Bumper2_Hit
  if LightC Then
    Addpoints(100)
  Else
    Addpoints(10)
  end If
  PlaySoundAtVol SoundFXDOF("FD_Bumper2",105,DOFPulse,DOFContactors), ActiveBall, 1
End Sub

Sub Bumper3_Hit
  if LightB Then
    Addpoints(1000)
  Else
    Addpoints(10)
  end If
  PlaySoundAtVol SoundFXDOF("FD_Bumper3",104,DOFPulse,DOFContactors), ActiveBall, 1
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

Sub Trigger_A1_Hit:PlaySound RolloverSound1,0,Soundlevel,-0.05,0.05:Addpoints(50):Light_A:ScoreMotor:DOF 108, DOFPulse:End Sub
Sub Trigger_A2_Hit:Playsound RolloverSound2,0,Soundlevel,-0.09,0.05:Addpoints(50):Light_A:ScoreMotor:DOF 112, DOFPulse:End Sub
Sub Trigger_B1_Hit:Playsound RolloverSound1,0,Soundlevel,0,0.05:Addpoints(50):Light_B:ScoreMotor:DOF 109, DOFPulse:End Sub
Sub Trigger_B2_Hit:P_LeftInlane.TransZ = -25:Playsound RolloverSound2,0,Soundlevel,-0.08,0.05:Addpoints(50):Light_B:ScoreMotor:DOF 113, DOFOn:End Sub
Sub Trigger_B2_Unhit:Trigger_B2.timerenabled = True:DOF 113, DOFOff:End Sub
Sub Trigger_B2_Timer:Trigger_B2.timerenabled = False:P_LeftInlane.TransZ = 0:End Sub
Sub Trigger_B3_Hit:P_RightInlane.TransZ = -25:Playsound RolloverSound2,0,Soundlevel,0.08,0.05:Addpoints(50):Light_B:ScoreMotor:DOF 115, DOFOn:End Sub
Sub Trigger_B3_Unhit:Trigger_B3.timerenabled = True:DOF 115, DOFOff:End Sub
Sub Trigger_B3_Timer:Trigger_B3.timerenabled = False:P_RightInlane.TransZ = 0:End Sub

Sub Trigger_C1_Hit:Playsound RolloverSound1,0,Soundlevel,0.05,0.05:Addpoints(50):Light_C:ScoreMotor:DOF 110, DOFPulse:End Sub
Sub Trigger_C2_Hit:Playsound RolloverSound2,0,Soundlevel,0.09,0.05:Addpoints(50):Light_C:ScoreMotor:DOF 114, DOFPulse:End Sub

Sub Trigger_LeftOutlane_Hit:P_LeftOutlane.TransZ = -25:PlaySound OutlaneSound,0,Soundlevel,-0.1,0.05:AddBonus:Addpoints(500):ScoreMotor:DOF 111, DOFOn:End Sub
Sub Trigger_LeftOutlane_Unhit:Trigger_LeftOutlane.timerenabled = True:DOF 111, DOFOff:End Sub
Sub Trigger_LeftOutlane_Timer:Trigger_LeftOutlane.timerenabled = False:P_LeftOutlane.TransZ = 0:End Sub
Sub Trigger_RightOutlane_Hit:P_RightOutlane.TransZ = -25:PlaySound OutlaneSound,0,Soundlevel,0.1,0.05:AddBonus:Addpoints(500):ScoreMotor:DOF 116, DOFOn:End Sub
Sub Trigger_RightOutlane_Unhit:Trigger_RightOutlane.timerenabled = True:DOF 116, DOFOff:End Sub
Sub Trigger_RightOutlane_Timer:Trigger_RightOutlane.timerenabled = False:P_RightOutlane.TransZ = 0:End Sub

Sub Target11_Hit:PlaySound SoundFXDOF(TargetSound,106,DOFPulse,DOFContactors),0,Soundlevel,-0.09,0.05:AddBonus:Addpoints(500):ScoreMotor:End Sub
Sub Target12_Hit:PlaySound SoundFXDOF(TargetSound,106,DOFPulse,DOFContactors),0,Soundlevel,-0.09,0.05:AddBonus:Addpoints(500):ScoreMotor:End Sub
Sub Target13_Hit:PlaySound SoundFXDOF(TargetSound,107,DOFPulse,DOFContactors),0,Soundlevel,0.09,0.05:AddBonus:Addpoints(500):ScoreMotor:End Sub
Sub Target14_Hit:PlaySound SoundFXDOF(TargetSound,107,DOFPulse,DOFContactors),0,Soundlevel,0.09,0.05:AddBonus:Addpoints(500):ScoreMotor:End Sub

'Drop Targets
Sub Target1_Hit:LightDTL1.State = 1:playsound DTSound,0,Soundlevel,-0.08,0.05:AddBonus:Addpoints(500):ScoreMotor:End Sub
Sub Target2_Hit:LightDTL2.State = 1:playsound DTSound,0,Soundlevel,-0.08,0.05:AddBonus:Addpoints(500):ScoreMotor:End Sub
Sub Target3_Hit:LightDTL3.State = 1
  playsound DTSound,0,Soundlevel,-0.08,0.05
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
Sub Target4_Hit:LightDTL4.State = 1:playsound DTSound,0,Soundlevel,-0.08,0.05:AddBonus:Addpoints(500):ScoreMotor:End Sub
Sub Target5_Hit:LightDTL5.State = 1:playsound DTSound,0,Soundlevel,-0.08,0.05:AddBonus:Addpoints(500):ScoreMotor:End Sub
Sub Target6_Hit:LightDTL6.State = 1:playsound DTSound,0,Soundlevel,0.08,0.05:AddBonus:Addpoints(500):ScoreMotor:End Sub
Sub Target7_Hit:LightDTL7.State = 1:playsound DTSound,0,Soundlevel,0.08,0.05:AddBonus:Addpoints(500):ScoreMotor:End Sub
Sub Target8_Hit:LightDTL8.State = 1:
  playsound DTSound,0,Soundlevel,0.08,0.05
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
Sub Target9_Hit:LightDTL9.State = 1:playsound DTSound,0,Soundlevel,0.08,0.05:AddBonus:Addpoints(500):ScoreMotor:End Sub
Sub Target10_Hit:LightDTL10.State = 1:playsound DTSound,0,Soundlevel,0.08,0.05:AddBonus:Addpoints(500):ScoreMotor:End Sub

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
  PlaySoundAtVol SoundFXDOF("FD_DT5KwithReset",128,DOFPulse,DOFContactors), Target8, 1

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
  PlaySoundAtVol SoundFXDOF("FD_DT5KwithReset",128,DOFPulse,DOFContactors), Target8, 1

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

Sub RubberWall3_Hit:PlaySoundAtVol "rubber_hit_3", ActiveBall, 1:End Sub

Sub RubberSW1_Hit:Addpoints(10):PlaySoundAtVol "FD_10Switch", ActiveBall, 1:End Sub
Sub RubberSW2_Hit:Addpoints(10):PlaySoundAtVol "FD_10Switch", ActiveBall, 1:End Sub
Sub RubberSW3_Hit:MoveSwitchA:Addpoints(10):PlaySoundAtVol "FD_10Switch", ActiveBall, 1:End Sub
Sub RubberSW4_Hit:MoveSwitchB:Addpoints(10):PlaySoundAtVol "FD_10Switch", ActiveBall, 1:End Sub
Sub Gate1_Hit:PlaySoundAtVol "Gate5", ActiveBall, 1:End Sub

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

dim LFPress, RFPress, EOST, EOSA, EOSTnew, EOSAnew
dim FStrength, Frampup, FElasticity, EOSRampup, SOSRampup
dim RFEndAngle, LFEndAngle, LFCount, RFCount, LiveCatch
  LFEndAngle = 70: RFEndAngle = -70

EOST = leftflipper.eostorque
EOSA = leftflipper.eostorqueangle
FStrength = LeftFlipper.strength
Frampup = LeftFlipper.rampup
FElasticity = LeftFlipper.elasticity
EOSTnew = 1.0 'FEOST
EOSAnew = 0.2
EOSRampup = 1.5
SOSRampup = 8.5
LiveCatch = 8

Sub Flippertimer_Timer
  RFPrim.RotY = RightFlipper.currentangle
  LFPrim.RotY = LeftFlipper.currentangle
  If leftflipper.currentangle = leftflipper.endangle and LFPress = 1 then
    leftflipper.eostorqueangle = EOSAnew
    leftflipper.eostorque = EOSTnew
    LeftFlipper.rampup = EOSRampup
    if LFCount = 0 Then LFCount = GameTime
    if GameTime - LFCount < LiveCatch Then
      leftflipper.Elasticity = 0.1
      If LeftFlipper.endangle <> LFEndAngle Then leftflipper.endangle = LFEndAngle
    Else
      leftflipper.Elasticity = FElasticity
    end if
  elseif leftflipper.currentangle > leftflipper.startangle - 0.05  Then
    leftflipper.rampup = SOSRampup
    leftflipper.endangle = LFEndAngle - 3
    leftflipper.Elasticity = FElasticity
    LFCount = 0
  elseif leftflipper.currentangle > leftflipper.endangle + 0.01 Then
    leftflipper.eostorque = EOST
    leftflipper.eostorqueangle = EOSA
    LeftFlipper.rampup = Frampup
    leftflipper.Elasticity = FElasticity
  end if

  If rightflipper.currentangle = rightflipper.endangle and RFPress = 1 then
    rightflipper.eostorqueangle = EOSAnew
    rightflipper.eostorque = EOSTnew
    RightFlipper.rampup = EOSRampup
    if RFCount = 0 Then RFCount = GameTime
    if GameTime - RFCount < LiveCatch Then
      rightflipper.Elasticity = 0.1
      If RightFlipper.endangle <> RFEndAngle Then rightflipper.endangle = RFEndAngle
    Else
      rightflipper.Elasticity = FElasticity
    end if
  elseif rightflipper.currentangle < rightflipper.startangle + 0.05 Then
    rightflipper.rampup = SOSRampup
    rightflipper.endangle = RFEndAngle + 3
    rightflipper.Elasticity = FElasticity
    RFCount = 0
  elseif rightflipper.currentangle < rightflipper.endangle - 0.01 Then
    rightflipper.eostorque = EOST
    rightflipper.eostorqueangle = EOSA
    RightFlipper.rampup = Frampup
    rightflipper.Elasticity = FElasticity
  end if
End Sub

Dim firstBallOut
Sub ShooterLaneLaunch_Hit
  DOF 124, DOFPulse
  Set controlBall = ActiveBall
  firstBallOut = 1
  if GameMusic = 1 then
    if ActiveBall.vely < -8 then playsound "galopp_shooterlane",0,0.8
  Else
    if ActiveBall.vely < -8 then playsound "Launch",0,0.3,0.1,0.25
  End If
End Sub

'************************************************Post It Note Section**************************************************************************
'***************Static Post It Note Update
Dim  hsY, shift, scoreMil, score100K, score10K, scoreK, scoreUnit
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
    PpComma.image = hsArray(10)
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
      If playerScore(bSy) < playerScore(bSy+1) Then
        tempScore(1) = playerScore(bSy+1)
        tempPos(1) = position(bSy+1)
        playerScore(bSy+1) = playerScore(bSy)
        playerScore(bSy) = tempScore(1)
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
      If playerScore(hs) > highScore(chY) Then
        flag = flag + 1           'flag to show how many high scores needs replacing
        tempScore(1) = highScore(chY)
        highScore(chY) = playerScore(hs)
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
      playerEntry.image = "PlayerHS" & position(Flag)
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
      playsound SoundFXDOF("Knocker",125,DOFPulse,DOFKnocker)
      enableInitialEntry = True
    End If
End Sub


'************Enter Initials Keycode Subroutine
Dim initial(6,5)
Sub enterIntitals(keycode)
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
    If keycode = startGameKey Then
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
      Else                    'if on the last initial then get ready to exit the subroutine
        enableInitialEntry = False
        highScoreDelay.enabled = 1
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
  playsound "10a"
  flag = flag - 1
' TextBox2.text = Flag
  hsI = 1
  If flag = 0 Then          'exit high score entry mode and reset variables
    NoOfPlayers = 0
    For eIX = 1 to 4
      activeScore(eIX) = 0
      position(eIX) = 0
    Next
    playerEntry.visible = 0
    scoreUpdate = 0           'go to the highest score
    updatePostIt            'display that score
  Else
    showScore
    playerEntry.image = "PlayerHS" & position(flag)
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
  hsCount = hsCount + 1
  If hsCount > 1 Then
    saveHighScore
    dynamicUpdatePostIt.enabled = 1   'turn scrolling high score back on
    hsCount = 0
    highScoreDelay.enabled = 0
  End If
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
  Dim x, y
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
  If Not fileObj.FileExists(UserDirectory & OptionsName) Then
    Exit Sub
  End If
  Set scoreFile = fileObj.GetFile(UserDirectory & OptionsName)
  Set textStr = scoreFile.OpenAsTextStream(1,0)
    If (textStr.AtEndOfStream = True) Then
      Exit Sub
    End If

    For x = 1 to 29
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
    Credits = cdbl (temp(26))
    freePlay = cdbl (temp(27))
    BallsperGame = cdbl (temp(28))
    GameMusic = cdbl (temp(29))
    Set scoreFile = Nothing
      Set fileObj = Nothing
End Sub

'************Save Scores
Sub saveHighScore
Dim hiInit(5)
Dim hiInitTemp(5)
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
  Set scoreFile = fileObj.createTextFile(userDirectory & OptionsName,True)

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
    scoreFile.writeLine Credits
    scorefile.writeline freePlay
    scoreFile.WriteLine ballsPerGame
    scoreFile.WriteLine GameMusic
    scoreFile.Close
  Set scoreFile = Nothing
  Set fileObj = Nothing

'This section of code writes a file in the User Folder of VisualPinball that contains the High Score data for PinballY.
'PinballY can read this data and display the high scores on the DMD during game selection mode in PinballY.

  Set FileObj = CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) Then
    Exit Sub
  End If
  Set ScoreFile=FileObj.CreateTextFile(UserDirectory & hsFileName & ".PinballYHighScores",True)

  For x = 0 to 4
    ScoreFile.WriteLine HighScore(x)
    ScoreFile.WriteLine HiInit(x)
  Next
  ScoreFile.Close
  Set ScoreFile=Nothing
  Set FileObj=Nothing

End Sub

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

' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 2000)
  Vol = Vol * RollingSoundFactor    'mfuegemann: adjust sound level
End Function

Function BallVelZ(ball) 'Calculates the ball speed in the -Z
    BallVelZ = INT((ball.VelZ) * -1 )
End Function

Function VolZ(ball) ' Calculates the Volume of the sound based on the ball speed in the Z
    VolZ = Csng(BallVelZ(ball) ^2 / 2000)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / FastDraw.width-1
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
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

Function AudioFade(ball)
    Dim tmp : tmp = ball.y * 2 / FastDraw.height-1
    If tmp > 0 Then
        AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
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
  PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioPan(ActiveBall)
End Sub

Sub Targets_Hit (idx)
  PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioPan(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
  PlaySound "metalhit_thin", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioPan(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
  PlaySound "metalhit_medium", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioPan(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
  PlaySound "metalhit2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioPan(ActiveBall)
End Sub

Sub Gates_Hit (idx)
  PlaySound "gate4", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioPan(ActiveBall)
End Sub

Sub Spinner_Spin
  PlaySoundAtVol "fx_spinner", Spinner, 1
End Sub

Sub Rubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 20 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioPan(ActiveBall)
  End if
  If finalspeed >= 2 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End Sub

Sub Posts_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioPan(ActiveBall)
  End if
  If finalspeed >= 2 AND finalspeed <= 16 then
    RandomSoundRubber()
  End If
End Sub

Sub RandomSoundRubber()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioPan(ActiveBall)
    Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioPan(ActiveBall)
    Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioPan(ActiveBall)
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
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioPan(ActiveBall)
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioPan(ActiveBall)
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioPan(ActiveBall)
  End Select
End Sub

Sub FastDraw_Exit()
  If B2SOn Then Controller.Stop
End Sub


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

'******************************************************
'       FLIPPER AND RUBBER CORRECTION
'******************************************************

dim LF : Set LF = New FlipperPolarity
dim RF : Set RF = New FlipperPolarity

InitPolarity

Sub InitPolarity()
  dim x, a : a = Array(LF, RF)
  for each x in a
    'safety coefficient (diminishes polarity correction only)
    'x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 0  'don't mess with these
    x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
    x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1
    x.enabled = True
    x.TimeDelay = 44
  Next


    'rf.report "Velocity"
    addpt "Velocity", 0, 0,   1
    addpt "Velocity", 1, 0.2,   1.07
    addpt "Velocity", 2, 0.41, 1.05
    addpt "Velocity", 3, 0.44, 1
    addpt "Velocity", 4, 0.65,  1.0'0.982
    addpt "Velocity", 5, 0.702, 0.968
    addpt "Velocity", 6, 0.95,  0.968
    addpt "Velocity", 7, 1.03,  0.945

    'rf.report "Polarity"
    AddPt "Polarity", 0, 0, -4.7
    AddPt "Polarity", 1, 0.16, -4.7
    AddPt "Polarity", 2, 0.33, -4.7
    AddPt "Polarity", 3, 0.37, -4.7 '4.2
    AddPt "Polarity", 4, 0.41, -4.7
    AddPt "Polarity", 5, 0.45, -4.7 '4.2
    AddPt "Polarity", 6, 0.576,-4.7
    AddPt "Polarity", 7, 0.66, -2.8'-2.1896
    AddPt "Polarity", 8, 0.743, -1.5
    AddPt "Polarity", 9, 0.81, -1.5
    AddPt "Polarity", 10, 0.88, 0

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

'Trigger Hit - .AddBall activeball
'Trigger UnHit - .PolarityCorrect activeball

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
      if not IsEmpty(balls(x) ) then
        balldata(x).Data = balls(x)
'       if DebugOn then StickL.visible = True : StickL.x = balldata(x).x
      End If
    Next
    PartialFlipCoef = ((Flipper.StartAngle - Flipper.CurrentAngle) / (Flipper.StartAngle - Flipper.EndAngle))  '% of flipper swing
    PartialFlipCoef = abs(PartialFlipCoef-1)
    if abs(Flipper.currentAngle - Flipper.EndAngle) < 20 Then 'last 20 degrees of swing is not dealt with
      PartialFlipCoef = 0
    End If
  End Sub

'***********gameTime is a global variable of how long the game has progressed in ms
'***********This function lets the table know if the flipper has been fired
  Private Function FlipperOn() : if gameTime < FlipAt+TimeDelay then FlipperOn = True : End If : End Function 'Timer shutoff for polaritycorrect

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
          'TB.TEXT = balldata(x).id & " " & BALLDATA(X).X & VBNEWLINE & FLIPPERSTART & " " & FLIPPEREND
          if ballpos > 0.65 then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)        'find safety coefficient 'ycoef' data
        end if
      Next

      'Velocity correction
      if not IsEmpty(VelocityIn(0) ) then
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

' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX, Rothbauerw, Thalamus and Herweh
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
'            Supporting Ball, Sound Functions and Math
'*********************************************************************

Function RndNum(min, max)
  RndNum = Int(Rnd() * (max-min + 1) ) + min ' Sets a random number between min and max
End Function

Const Pi = 3.1415927

Function dSin(degrees)
  dsin = sin(degrees * Pi/180)
End Function

Function dCos(degrees)
  dcos = cos(degrees * Pi/180)
End Function

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = tableobj.y * 2 / table1.height-1
' Thalamus, AudioFade - Patched
  If tmp > 0 Then
    AudioFade = Csng(tmp ^5) 'was 10
  Else
    AudioFade = Csng(-((- tmp) ^5) ) 'was 10
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = tableobj.x * 2 / table1.width-1
' Thalamus, AudioPan - Patched
  If tmp > 0 Then
    AudioPan = Csng(tmp ^5) 'was 10
  Else
    AudioPan = Csng(-((- tmp) ^5) ) 'was 10
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = ball.x * 2 / table1.width-1
' Thalamus, Pan - Patched
  If tmp > 0 Then
    Pan = Csng(tmp ^5) 'was 10
  Else
    Pan = Csng(-((- tmp) ^5) ) ' was 10
  End If
End Function

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
  Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function VolMulti(ball,Multiplier) ' Calculates the Volume of the sound based on the ball speed
  VolMulti = Csng(BallVel(ball) ^2 / 150 ) * Multiplier
End Function

Function DVolMulti(ball,Multiplier) ' Calculates the Volume of the sound based on the ball speed
  DVolMulti = Csng(BallVel(ball) ^2 / 150 ) * Multiplier
  debug.print DVolMulti
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


'The Kessel Run
'mod of 32assasin's conversion of mfugemann's Ice Cold Beer
' Graphics, sound, and flex routines by endeemillr
' added two holes to make 12 "parsecs" the goal




'############################################################################################
'############################################################################################
'#######                                                                             ########
'#######          Ice Cold Beer                                                      ########
'#######          (Taito 1983)                                                       ########
'#######                                                                             ########
'############################################################################################
'############################################################################################
' Version 1.4 FS mfuegemann 2015
'
' Thanks to :
' Zany for the Playfield and Bezel images
' Aldiode for the recording of game sounds - now complete
' gtxjoe for the Highscore PostIt code

'1.1 changes:
'- option to switch the Flipper and MagnaSave keys
'- option to use a Double Tap feature on cabs without MagnaSave keys. Double tap Your respective Flipper key for the down movement

'1.2 changes:
'- new sounds by aldiode

'1.3 changes:
'- DOF code and events added by Arngrim (using "LMEMDOFTables.txt")

'1.4 changes:
'- set option explicit, added missing variable definitions


option explicit
Randomize
'-------------------------------
'----- DIP Switch Settings -----
'-------------------------------

Const MaxBalls = 3        'Balls per Game
Const DIP_BonusCountDown = 1  'Bonus        (0)Slow: 5,4,3,2  (1)Factory: 4,3,2,2   (2)Fast: 3,2,2,1  (3)X-Fast: 2,2,1,1
Const DIP_RodPromptTime = 1   'Rod Prompt Time  (0)Easy: 5,4,2,1  (1)Factory: 4,2,1,1   (2)Hard: 2,2,1,1  (3)X-Hard: 1,1,1,1
Const RoundsToStar = 1      'Rounds to lite Star  1,2
Const ExtraBallScore = 4000   'ExtraBall  off,2000,4000,8000 (off = 999999), 2nd EB after another 10.000 Points

Const ShowHighScore = 1     'Enable Highscore Feature
Const ResetHighScore = 0    'Delete Highscore - remove again after Reset
Const FreePlay = 0        'FreePlay

Const ReverseRodKeys = 0    'switch Flipper and MagnaSave keys
Const DoubleTap = 0       'use this if You have no MagnaSave buttons, a double tap on the Flipper buttons will activate the down movement
Const DoubleTapTime = 72    'Max. interval in ms for the Double Tap to register. Please set this to the minimal interval, You can live with


Const UseFlexDMD = 1    ' 1 is on
'-------------------------
'----- Ice Cold Beer -----
'-------------------------
Dim ActiveHole,GameOn,ExtraBalls,BallInPlay,Credits,Level,Bonus,Score,Oops,NewBonus,EBgranted,RoundsAchieved,i,BonusText,ScoreText,CreditsText,HSArray,IdlePosY

Dim Gamemode
Gamemode = 1

Const LMEMTableConfig="LMEMDOFTables.txt"

Sub Table1_Init

  LoadLMEMConfig

  ' initalise the FlexDMD display
    If UseFlexDMD Then FlexDMD_Init

' if B2SOn then
    Set Controller = CreateObject("B2S.Server")
    Controller.B2SName = "Kessel Run"
    Controller.Run()
    Controller.B2SSetData 35,4
    If Err Then MsgBox "Can't Load B2S.Server."
' End If

  ApplyDIPSwitchSettings


  if ResetHighScore then
    SetDefaultHSTD
  end if

  AY = IdlePosY - 75
  BY = IdlePosY - 15

  IdleMode = 1
  Credits = 0
  if FreePlay = 1 then
    Credits = 99
    DOF 102, 1
  end if

  BallDestroyed = True
  GameBallAssigned = False
  GameOn = False
  BallInPlay = 0
  ActiveHole = 0
  ExtraBalls = 0
  Bonus = 0
  NewBonus = 0
  Score = 0
  LightStar.state = Lightstateoff
  ForceLeftUp = False
  ForceRightUp = False
  ForceRightDown = False
  ForceLeftDown = False

  IdleTimer.enabled = True

  if ShowHighScore = 1 then
    PTape.image = HSArray(12)
    UpdatePostIt
  else
    PTape.image = HSArray(10)
    PComma.image = HSArray(10)
    PScore1.image = HSArray(10)
    PScore2.image = HSArray(10)
    PScore3.image = HSArray(10)
    PScore4.image = HSArray(10)
    PScore5.image = HSArray(10)
    PScore6.image = HSArray(10)
  end if

  DoubleTapRightTimer.interval = DoubleTapTime
  DoubleTapLeftTimer.interval = DoubleTapTime
  DoubleTapRightGraceTimer.interval = DoubleTapTime - 1
  DoubleTapLeftGraceTimer.interval = DoubleTapTime - 1

  successstate = false
End Sub

Sub table1_Exit()
  If UseFlexDMD then
    If Not FlexDMD is Nothing Then
      FlexDMD.Show = False
      FlexDMD.Run = False
      FlexDMD = NULL
    End if
  End if


End Sub

Sub GameTimer_Timer
  If WrongHoleHitTimer.enabled then
    Display Digits(1),"-"
    Bonustext = "---"
    For i=1 To 3
      Display Digits(1+i),Mid(BonusText,i,1)
    Next
    Scoretext = "OOPS"
  else
    if BallInPlay > 0 then
      Display Digits(1),BallInPlay
    else
      Display Digits(1)," "
    end if
    Bonustext = ""
    For i=1 To 3 - Len(Bonus)
      Bonustext = BonusText & " "
    Next
    Bonustext = BonusText & Bonus
    For i=1 To 3
      Display Digits(1+i),Mid(BonusText,i,1)
    Next

    Scoretext = ""
    For i=1 To 4 - Len(Score)
      Scoretext = ScoreText & " "
    Next
    Scoretext = ScoreText & Score
  end if

  Creditstext = "Cr"
  For i=1 To 2 - Len(Credits)
    Creditstext = CreditsText & " "
  Next
  Creditstext = CreditsText & Credits

  if GameOn then
    GameOverLight.state = 0
    For i=1 To 4
      Display Digits(4+i),Mid(ScoreText,i,1)
    Next
  else
    GameOverLight.state = 1
    For i=1 To 4
      Display Digits(4+i),Mid(IdleText,i,1)
    Next
  end if

  if LeftUp or LeftDown or RightUp or RightDown or GracePeriodTimer.enabled then
    ForceRodTimer.enabled = False
    if GracePeriodTimer.enabled and (LeftUp or LeftDown or RightUp or RightDown) then
      GracePeriodTimer.enabled = False
      ForceRodTimer.enabled = True
      BonusTimer.enabled = True
    end if
  else
    if (not RodMoveDownTimer.enabled) and (not RodMoveUpTimer.enabled) then
      if (not BallDestroyed) and (GameBallAssigned) then
        if GameBall.z > 30 then
          ForceRodTimer.enabled = True
        end if
      end if
    end if
  end if

  if ExtraBalls > 1 then
    EB2Light.state = 1
  else
    EB2Light.state = 0
  end if

  if ExtraBalls > 0 then
    EB1Light.state = 1
  else
    EB1Light.state = 0
  end if
End Sub

Dim FArray,Farray2,ii
'FArray = Array(Flasher1,Flasher2,Flasher3,Flasher4,Flasher5,Flasher6,Flasher7,Flasher8,Flasher9,Flasher10)
'FArray2 = Array(Flasher1a,Flasher2a,Flasher3a,Flasher4a,Flasher5a,Flasher6a,Flasher7a,Flasher8a,Flasher9a,Flasher10a)
'Sub FlasherTimer_Timer
  'if (ActiveHole > 0) and (not Hole10MadeTimer.enabled) then
'   FArray(ActiveHole-1).isvisible = not FArray(ActiveHole-1).isvisible
'   For ii = 0 to 9
'     if ii <> ActiveHole-1 then
'       FArray(ii).isvisible = False
'     end if
'     FArray2(ii).isvisible = FArray(ii).isvisible
'   next
' end if
'End Sub


Dim IdleMode,IdleText
Sub IdleTimer_Timer
  IdleMode = IdleMode + 1
  if IdleMode > 17 then IdleMode = 1

  select case IdleMode
    case 1: IdleText = "    "
    case 2: IdleText = CreditsText
    case 3: IdleText = CreditsText
    case 4: IdleText = CreditsText
    case 5: IdleText = "    "
    case 6: IdleText = "HAN "
    case 7: IdleText = "SOLO"
    case 8: IdleText = "FLYS"
    case 9: IdleText = "FAL "
    case 10:IdleText = "CON "
    case 11: IdleText = "12  "
    case 12: IdleText = "PAr "
    case 13: IdleText = "SECS"
    case 14: IdleText = "    "
    case 15: IdleText = ScoreText
    case 16: IdleText = ScoreText
    case 17: IdleText = ScoreText
  end select
End Sub

Sub Drain_Hit()
  BonusTimer.enabled = False
  ForceRodTimer.enabled = False
  BallDestroyed = True
  GameBallAssigned = False
  Drain.destroyball
  DOF 114, 2
  if Oops then
    WrongHole
    if ExtraBalls > 0 then
      ExtraBalls = ExtraBalls - 1
    else
      BallInPlay = BallInPlay + 1
    If UseFlexDMD then FlexDMDUpdate
    end if

    if BallInPlay > MaxBalls then
      EndGame
      DOF 106, 2
    else
      NextBallTimer.enabled = True
      DOF 103, 2
    end if
  else
    NextBallTimer.enabled = True
    oops = true
  end if
End Sub

Sub NextBallTimer_Timer
  if (not CollectBonusTimer2.enabled) and (not CollectBonusTimer1.enabled) then
    NextBallTimer.enabled = False
    NextBall
  end if
End Sub


'--------------------------
'-----  Game Handling -----
'--------------------------
Sub StartGame
  if Credits > 0 then
    for each obj in HoleLights
      obj.state = Lightstateoff
    next
    LightStar.state = Lightstateoff
    if FreePlay = 0 then
      Credits = Credits - 1
      if Credits < 1 Then DOF 102, 0
    end if
    ExtraBalls = 0
    EBgranted = 0
    IdleTimer.enabled = False
    GameOn = True
    BallInPlay = 1
    Level = 1
    ActiveHole = 1
    Light1.state = Lightstateblinking
    Score = 0
    RoundsAchieved = 0
    NewBonus = 100
    BonusTimer.Interval = BonusCountDownL1
    ForceRodTimer.Interval = RodPromptTimeL1
    StopSound "Gameover"
    NextBall
      playsound "gameplay"
    Musictimer.Enabled = True
    If UseFlexDMD Then backtogame
    If UseFlexDMD Then FlexDMDUpdate

    If UseFlexDMD Then attractmode.enabled = False

    Oops =  true
  end if
end sub

Sub EndGame
  if UseFlexDMD then judgeplayer
  CheckNewHighScorePostIt(Score)
  BallInPlay = MaxBalls
  Bonus = 0
  IdleMode = 10
  IdleTimer.enabled = True
  GameOverTimer.enabled = True
    If UseFlexDMD Then FlexDMDUpdate
End Sub

Sub GameOverTimer_Timer
  GameOverTimer.enabled = False
  GameOn = False
  BallInPlay = 0
  stopsound "gameplay"
    Musictimer.Enabled = False
  playsound "GameOver",-1,1,0,0
End Sub

Sub NextBall
  RodMoveDownTimer.enabled = True
End Sub

Sub RodMoveDownTimer_Timer
  ForceLeftUp = False
  ForceRightUp = False
  ForceRightDown = True
  ForceLeftDown = True

  if AY > RestPosY-25 then
    ForceLeftDown = False
  end if

  if (AY > RestPosY-25) and (BY > RestPosY-1) and (not Balldestroyed) then
    RodMoveUpTimer.enabled = True
    RodMoveDownTimer.enabled = False
    ForceRightDown = False
    ForceLeftDown = False
    ForceLeftUp = True
    ForceRightUp = True
  end if
End Sub

Sub RodMoveUpTimer_Timer
  if (AY < IdlePosY) then
    ForceLeftUp = False
  end if
  if (BY < IdlePosY) then
    ForceRightUp = False
  end if
  if (AY < IdlePosY) and (BY < IdlePosY) then
    RodMoveUpTimer.enabled = False
    GracePeriodTimer.enabled = True
    ForceLeftUp = False
    ForceRightUp = False
    ForceRightDown = False
    ForceLeftDown = False
  end if
End Sub

Sub GracePeriodTimer_Timer
  GracePeriodTimer.enabled = False
  BonusTimer.enabled = True
  ForceRodTimer.enabled = True
End Sub

Sub BonusTimer_Timer
  if Bonus > 10 then
    playsound "BonusCount"
    Bonus = Bonus - 10
      If UseFlexDMD then FlexDMDUpdate
    if Bonus = 10 then
      'playsound ""   'Warning on lower Bonus Limit
    end if

      If UseFlexDMD then FlexDMDUpdate
  end if
End Sub

Dim ForceRod
Sub ForceRodTimer_Timer
  if (not BallDestroyed) and (GameBallAssigned) then
    if GameBall.z > 30 then
      ForceRod = 22
      ROMRodTimer.enabled = True
    end if
  end if
End Sub

Sub ROMRodTimer_Timer
  Playsound SoundFX("Motor1"),0,1,0,0.1
  DOF 101, 2
  ForceRod = ForceRod - RodSteps
  if By > MaxPosY then By = By - RodSteps
  if Ay > MaxPosY then Ay = Ay - RodSteps
  if ForceRod <= 0 then
    ROMRodTimer.enabled = False
  end if
End Sub

'-------------------------
'----- Hole Switches -----
'-------------------------
Sub Hole1_hit
  if ActiveHole = 1 then
    CorrectHole
  end if
End Sub

Sub Hole2_hit
  if ActiveHole = 2 then
    CorrectHole
  end if
End Sub

Sub Hole3_hit
  if ActiveHole = 3 then
    CorrectHole
  end if
End Sub

Sub Hole4_hit
  if ActiveHole = 4 then
    CorrectHole
  end if
End Sub

Sub Hole5_hit
  if ActiveHole = 5 then
    CorrectHole
  end if
End Sub

Sub Hole6_hit
  if ActiveHole = 6 then
    CorrectHole
  end if
End Sub

Sub Hole7_hit
  if ActiveHole = 7 then
    CorrectHole
  end if
End Sub

Sub Hole8_hit
  if ActiveHole = 8 then
    CorrectHole
  end if
End Sub

Sub Hole9_hit
  if ActiveHole = 9 then
    CorrectHole
  end if
End Sub

Sub Hole10_hit
  if ActiveHole = 10 then
    CorrectHole
  end if
End Sub

Sub Hole11_hit
  if ActiveHole = 11 then
    CorrectHole
  end if
End Sub

Sub Hole12_hit
  if ActiveHole = 12 then
    CorrectHole
  end if
End Sub

Sub CorrectHole
  Oops = false
  LeftUp = False
  RightUp = False
  RightDown = False
  LeftDown = False
'newsound
  stopsound "intro"
  success
  DOF 107, 2
  BonusTimer.enabled = False
  ForceRodTimer.enabled = False

  if ActiveHole = 12 then
    Hole10MadeTimer.enabled = True
    playsound "Win"
    DOF 104, 2
  end if
  CollectBonusTimer1.enabled = True
End Sub

Sub WrongHole
  LeftUp = False
  RightUp = False
  RightDown = False
  LeftDown = False
  stopsound "intro"
  stopsound successound
  wrongholesound
  BonusTimer.enabled = False
  ForceRodTimer.enabled = False
  WrongHoleHitTimer.enabled = True
End Sub

Sub WrongHoleHitTimer_Timer
  WrongHoleHitTimer.enabled = False
End Sub

Sub CollectBonusTimer1_Timer
  if not Hole10MadeTimer.enabled then
    CollectBonusTimer1.enabled = False
    CollectBonusTimer2.enabled = True
    BonusTimer.enabled = False
  end if
End Sub

Sub CollectBonusTimer2_Timer
  if Bonus > 0 then
    Score = Score + 10
    Bonus = Bonus - 10
    Playsound "CollectBonus"
    If UseFlexDMD Then FlexDMDUpdate
    if Score >= ExtraBallScore + (10000 * EBGranted) then
      Playsound "ExtraBall"
      EBGranted = 1
      ExtraBalls = ExtraBalls + 1
    end if
  else
    CollectBonusTimer2.enabled = False

    Select Case ActiveHole
      Case 1: Light1.state = 0:Light2.state = Lightstateblinking
      Case 2: Light2.state = 0:Light3.state = Lightstateblinking
      Case 3: Light3.state = 0:Light4.state = Lightstateblinking
      Case 4: Light4.state = 0:Light5.state = Lightstateblinking
      Case 5: Light5.state = 0:Light6.state = Lightstateblinking
      Case 6: Light6.state = 0:Light7.state = Lightstateblinking
      Case 7: Light7.state = 0:Light8.state = Lightstateblinking
      Case 8: Light8.state = 0:Light9.state = Lightstateblinking
      Case 9: Light9.state = 0:Light10.state = Lightstateblinking
      Case 10:Light10.state = 0:Light11.state = Lightstateblinking
      Case 11:Light11.state = 0:Light12.state = Lightstateblinking
      Case 12: Light12.state = 0:Light1.state = Lightstateblinking
    End Select

    ActiveHole = ActiveHole + 1
    if ActiveHole > 12 then
      ActiveHole = 1
      RoundsAchieved = RoundsAchieved + 1
      if RoundsAchieved >= RoundsToStar then
        LightStar.state = Lightstateon
      end if
      Level = Level + 1

      BonusTimer.Interval = BonusCountDownL4
      ForceRodTimer.Interval = RodPromptTimeL4
      if Level = 2 then
        BonusTimer.Interval = BonusCountDownL2
        ForceRodTimer.Interval = RodPromptTimeL2
      end if
      if Level = 3 then
        BonusTimer.Interval = BonusCountDownL3
        ForceRodTimer.Interval = RodPromptTimeL3
      end if
    end if
    NewBonus = ActiveHole * 100
    if NewBonus = 1000 then
      NewBonus = 900
    end if
  end if
End Sub


'Hole 10 Sequence

Dim Hole10Loop
Hole10Loop = 0
Sub Hole10MadeTimer_Timer
  Hole10Loop = Hole10Loop + 1
  for each obj in HoleLights
    obj.state = Lightstateoff
  next
  'for each obj in HoleFlashers
  ' obj.isvisible = False
' next

  Select Case Hole10Loop
    Case 1: Light1.state = Lightstateon
        Light5.state = Lightstateon
        Light6.state = Lightstateon
        Light8.state = Lightstateon
    '   Flasher1.isvisible = True
    '   Flasher1a.isvisible = True
    '   Flasher5.isvisible = True
    '   Flasher5a.isvisible = True
    '   Flasher6.isvisible = True
    '   Flasher6a.isvisible = True
    '   Flasher8.isvisible = True
    '   Flasher8a.isvisible = True
    Case 2: Light3.state = Lightstateon
        Light4.state = Lightstateon
        Light5.state = Lightstateon
        Light10.state = Lightstateon
    '   Flasher3.isvisible = True
    '   Flasher3a.isvisible = True
    '   Flasher4.isvisible = True
    '   Flasher4a.isvisible = True
    '   Flasher5.isvisible = True
    '   Flasher5a.isvisible = True
    ''    Flasher10.isvisible = True
    '   Flasher10a.isvisible = True
    Case 3: Light2.state = Lightstateon
        Light5.state = Lightstateon
        Light6.state = Lightstateon
        Light8.state = Lightstateon
        Light9.state = Lightstateon
    '   Flasher2.isvisible = True
    '   Flasher2a.isvisible = True
    '   Flasher5.isvisible = True
    '   Flasher5a.isvisible = True
    '   Flasher6.isvisible = True
    '   Flasher6a.isvisible = True
    '   Flasher8.isvisible = True
    '   Flasher8a.isvisible = True
    '   Flasher9.isvisible = True
    '   Flasher9a.isvisible = True
    Case 4: Light1.state = Lightstateon
        Light3.state = Lightstateon
        Light4.state = Lightstateon
        Light7.state = Lightstateon
        Light12.state = Lightstateon
    '   Flasher1.isvisible = True
    '   Flasher1a.isvisible = True
    '   Flasher3.isvisible = True
    '   Flasher3a.isvisible = True
    '   Flasher4.isvisible = True
    '   Flasher4a.isvisible = True
    '   Flasher7.isvisible = True
    '   Flasher7a.isvisible = True
    '   Flasher10.isvisible = True
    '   Flasher10a.isvisible = True
    Case 5: Light1.state = Lightstateon
        Light5.state = Lightstateon
        Light6.state = Lightstateon
        Light8.state = Lightstateon
    '   Flasher1.isvisible = True
    '   Flasher1a.isvisible = True
    '   Flasher5.isvisible = True
    '   Flasher5a.isvisible = True
    '   Flasher6.isvisible = True
    '   Flasher6a.isvisible = True
    '   Flasher8.isvisible = True
    ''    Flasher8a.isvisible = True
    Case 6: Light3.state = Lightstateon
        Light4.state = Lightstateon
        Light5.state = Lightstateon
        Light11.state = Lightstateon
    '   Flasher3.isvisible = True
    '   Flasher3a.isvisible = True
    '   Flasher4.isvisible = True
    '   Flasher4a.isvisible = True
    '   Flasher5.isvisible = True
    '   Flasher5a.isvisible = True
    '   Flasher10.isvisible = True
    '   Flasher10a.isvisible = True
    Case 7: Light2.state = Lightstateon
        Light5.state = Lightstateon
        Light6.state = Lightstateon
        Light7.state = Lightstateon
        Light10.state = Lightstateon
    '   Flasher2.isvisible = True
    '   Flasher2a.isvisible = True
    '   Flasher5.isvisible = True
    '   Flasher5a.isvisible = True
    '   Flasher6.isvisible = True
    '   Flasher6a.isvisible = True
    '   Flasher7.isvisible = True
    '   Flasher7a.isvisible = True
    '   Flasher10.isvisible = True
    '   Flasher10a.isvisible = True
    Case 8: Light1.state = Lightstateon
        Light3.state = Lightstateon
        Light4.state = Lightstateon
        Light8.state = Lightstateon
        Light9.state = Lightstateon
    '   Flasher1.isvisible = True
    '   Flasher1a.isvisible = True
    '   Flasher3.isvisible = True
    '   Flasher3a.isvisible = True
    '   Flasher4.isvisible = True
    '   Flasher4a.isvisible = True
    '   Flasher8.isvisible = True
    '   Flasher8a.isvisible = True
    '   Flasher9.isvisible = True
    '   Flasher9a.isvisible = True
    Case 9: Light2.state = Lightstateon
        Light4.state = Lightstateon
        Light5.state = Lightstateon
        Light9.state = Lightstateon
        Light10.state = Lightstateon
    Case 10:Light12.state = Lightstateon
        Light3.state = Lightstateon
        Light4.state = Lightstateon
        Light8.state = Lightstateon
        Light9.state = Lightstateon
    Case 11:  for each obj in HoleLights
          obj.state = Lightstateon
        next
    '   for each obj in HoleFlashers
    '     obj.isvisible = True
'       next
    Case 12:Hole10MadeTimer.enabled = False
        Hole10Loop = 0
  end select
End Sub

'Apply DIP-Settings
Dim BonusCountDownL1,BonusCountDownL2,BonusCountDownL3,BonusCountDownL4,RodPromptTimeL1,RodPromptTimeL2,RodPromptTimeL3,RodPromptTimeL4
Sub ApplyDIPSwitchSettings
  Select Case DIP_BonusCountDown
    case 0: BonusCountDownL1 = 5000
        BonusCountDownL2 = 4000
        BonusCountDownL3 = 3000
        BonusCountDownL4 = 2000
    case 2: BonusCountDownL1 = 3000
        BonusCountDownL2 = 2000
        BonusCountDownL3 = 2000
        BonusCountDownL4 = 2000
    case 3: BonusCountDownL1 = 2000
        BonusCountDownL2 = 2000
        BonusCountDownL3 = 1000
        BonusCountDownL4 = 1000
    case else
        BonusCountDownL1 = 4000
        BonusCountDownL2 = 3000
        BonusCountDownL3 = 2000
        BonusCountDownL4 = 2000
  end select

  Select Case DIP_RodPromptTime
    case 0: RodPromptTimeL1 = 5000
        RodPromptTimeL2 = 4000
        RodPromptTimeL3 = 2000
        RodPromptTimeL4 = 1000
    case 2: RodPromptTimeL1 = 2000
        RodPromptTimeL2 = 2000
        RodPromptTimeL3 = 1000
        RodPromptTimeL4 = 1000
    case 3: RodPromptTimeL1 = 1000
        RodPromptTimeL2 = 1000
        RodPromptTimeL3 = 1000
        RodPromptTimeL4 = 1000
    case else
        RodPromptTimeL1 = 4000
        RodPromptTimeL2 = 2000
        RodPromptTimeL3 = 1000
        RodPromptTimeL4 = 1000
  end select
end sub

'-------------------------
'-----  Key Handling -----
'-------------------------
Dim LeftUp,LeftDown,RightUp,RightDown,ForceLeftUp,ForceLeftDown,ForceRightUp,ForceRightDown

Sub Table1_KeyDown(ByVal keycode)
  if keycode = AddCreditKey then
    if not GameOn then
      IdleMode = 1
      playsound "coin3"
      Credits = Credits + 1
      DOF 102, 1
      if Credits > 9 then
        Credits = 9
      end if
    if UseFlexDMD then
    FlexDMDUpdate
    flexcredit
    end if
    end if

  end if
  if keycode = AddCreditKey2 then
    if not GameOn then
      IdleMode = 1
      playsound "coin3"
      Credits = Credits + 1
      DOF 102, 1
      if Credits > 9 then
        Credits = 9
      end if
    if UseFlexDMD then
    FlexDMDUpdate
    flexcredit
    end if
    end if
  end if
  if keycode = Startgamekey then
    if (not GameOn) and (BallDestroyed) then
    If UseFlexDMD then backtogame
      StartGame
    end if
  end if

  if DoubleTap = 0 then
    if ReverseRodKeys = 0 then
      If keycode = LeftFlipperKey Then
        LeftUP = True
      End If
      If keycode = LeftMagnasave Then
        LeftDown = True
      End If
      If keycode = RightFlipperKey Then
        RightUp = True
      End If
      If keycode = RightMagnasave Then
        RightDown = True
      End If
    else
      If keycode = LeftMagnasave Then
        LeftUP = True
      End If
      If keycode = LeftFlipperKey Then
        LeftDown = True
      End If
      If keycode = RightMagnasave Then
        RightUp = True
      End If
      If keycode = RightFlipperKey Then
        RightDown = True
      End If
    end if
  else
    If keycode = LeftFlipperKey Then
      if DoubleTapLeftTimer.enabled then
        LeftDown = True
        DoubleTapLeftTimer.enabled = False
      else
        DoubleTapLeftGraceTimer.enabled = False
        DoubleTapLeftGraceTimer.enabled = True
        LeftUP = True
      end if
    End If
    If keycode = RightFlipperKey Then
      if DoubleTapRightTimer.enabled then
        RightDown = True
        DoubleTapRightTimer.enabled = False
      else
        DoubleTapLeftGraceTimer.enabled = False
        DoubleTapLeftGraceTimer.enabled = True
        RightUp = True
      end if
    End If
  end if
End Sub

Sub Table1_KeyUp(ByVal keycode)
  if DoubleTap = 0 then
    if ReverseRodKeys = 0 then
      If keycode = LeftFlipperKey Then
        LeftUP = False
      End If
      If keycode = LeftMagnasave Then
        LeftDown = False
      End If
      If keycode = RightFlipperKey Then
        RightUp = False
      End If
      If keycode = RightMagnasave Then
        RightDown = False
      End If
    else
      If keycode = LeftMagnasave Then
        LeftUP = False
      End If
      If keycode = LeftFlipperKey Then
        LeftDown = False
      End If
      If keycode = RightMagnasave Then
        RightUp = False
      End If
      If keycode = RightFlipperKey Then
        RightDown = False
      End If
    end if
  else
    If keycode = LeftFlipperKey Then
      LeftUP = False
      LeftDown = False
      DoubleTapLeftTimer.enabled = False
      DoubleTapLeftTimer.enabled = True
    End If
    If keycode = RightFlipperKey Then
      RightUp = False
      RightDown = False
      DoubleTapRightTimer.enabled = False
      DoubleTapRightTimer.enabled = True
    End If

  end if
End Sub

Sub DoubleTapLeftTimer_Timer
  DoubleTapLeftTimer.enabled = False
End Sub

Sub DoubleTapLeftGraceTimer_Timer
  DoubleTapLeftGraceTimer.enabled = False
End Sub

Sub DoubleTapRightTimer_Timer
  DoubleTapRightTimer.enabled = False
End Sub

Sub DoubleTapRightGraceTimer_Timer
  DoubleTapRightGraceTimer.enabled = False
End Sub


'------------------------
'----- Rod Movement -----
'------------------------
Dim GameBall,BallDestroyed,GameBallAssigned,Ax,Ay,Bx,By,RestPosY,MaxPosY,MaxY,Vel

Ax = 0
Bx = 1200
RestPosY = BallRelease.y + 65   'Lower Y Limit
MaxPosY = 753     'Upper Y Limit
IdlePosY = 1833   'GameStart Y Position

Const RodSteps = 1.5  'Movement Speed  1.1 - 1.5

Sub RodMoveTimer_Timer
  if GameOn then
    if ForceLeftup or ForceRightup or ForceLeftdown or Forcerightdown then
      Playsound SoundFX("Motor1"),0,1,0,0.1
      DOF 101, 2
      if ForceLeftUp and (Ay > MaxPosY) then
        Ay = Ay - RodSteps
      end if
      if ForceLeftDown and (Ay < RestPosY) then
        Ay = Ay + RodSteps
      end if
      if ForceRightUp and (By > MaxPosY) then
        By = By - RodSteps
      end if
      if ForceRightDown and (By < RestPosY) then
        By = By + RodSteps
      end if
    else
      if (Leftup or Rightup or Leftdown or rightdown) and (not DoubleTapLeftGraceTimer.enabled) and (not DoubleTapRightGraceTimer.enabled) then
        if (not BallDestroyed) and (GameBallAssigned) then
          if GameBall.z > 20 then
            Playsound SoundFX("Motor1"),0,1,0,0.1
            DOF 101, 2

            if LeftUp and (Ay > MaxPosY) then
              Ay = Ay - RodSteps
            end if
            if LeftDown and (Ay < RestPosY) then
              Ay = Ay + RodSteps
            end if
            if RightUp and (By > MaxPosY) then
              By = By - RodSteps
            end if
            if RightDown and (By < RestPosY) then
              By = By + RodSteps
            end if
          end if
        end if
      end if
    end if
  end if
End Sub

Sub RodSoundTimer_Timer
  if GameOn then
    if Leftup or Rightup or Leftdown or rightdown then
      if (not BallDestroyed) and (GameBallAssigned) then
        if GameBall.z > 20 then
          if LeftUp  then
            PlaySound  "LeftRodUp",0,0.2
          end if
          if RightUp then
            Playsound "RightRodUp",0,0.2
          end if
          if LeftDown or RightDown then
            Playsound "RodDown",0,0.2
          end if
        end if
      end if
    end if
  end if
End Sub

Sub BallTimer_Timer
  'Rod animation
  P_rod.y = (Ay+By)/2

  P_LeftRod.y = ((By-Ay)/Bx)*P_LeftRod.x+Ay-25
  P_RightRod.y = ((By-Ay)/Bx)*P_RightRod.x+Ay-25

  P_rod.rotz = atn((By-Ay)/(Bx-Ax)) * 180/(4*atn(1))
  P_Rod.size_y = 75 * (SQR((Bx-Ax)^2+(By-Ay)^2) / (Bx-Ax))

  P_LeftRod.rotz = P_rod.rotz
  P_RightRod.rotz = P_rod.rotz

  'Ball above Playfield
  if (not BallDestroyed) and (GameBallAssigned) then
    if GameBall.z > 20 then
      MaxY = ((By-Ay)/Bx) * GameBall.x + Ay - 25    '30 for FS
      if GameBall.y > MaxY then
        Vel = 0.004 * abs(By-Ay)
        if Vel > 0.09 then
          Vel = 0.09
        end if
        GameBall.y = MaxY' - 1
        if By < Ay then
          GameBall.velx = GameBall.velx - GameBall.vely * Vel
        else
          GameBall.velx = GameBall.velx + GameBall.vely * Vel
        end if
        GameBall.vely = 0
      end if
    end if
  end if

  'Gate animation
  if (((By-Ay)/Bx)*BallRelease.x+Ay) > (P_Gate.y-17) then
    P_Gate.ObjRotZ = ((((By-Ay)/Bx)*BallRelease.x+Ay) - (P_Gate.y-17)) * 0.5
  else
    P_Gate.ObjRotZ = 0
  end if

  'Ball Release
  if (((By-Ay)/Bx)*BallRelease.x+Ay-25) > BallRelease.y then
    if GameOn and BallDestroyed then
'newsound
      playsound "intro"
      Bonus = Bonus + NewBonus
      If UseFlexDMD then FlexDMDUpdate
      NewBonus = 0
      Balldestroyed = False
      BallRelease.CreateSizedBall 18
      BallRelease.kick 180,1
      DOF 114, 2
    end if
  end if
End Sub

Sub Trigger1_Hit
  Set GameBall = ActiveBall
  GameBallAssigned = True
End Sub


'------------------------
'----- Display Code -----
'------------------------

Dim Digits(8),obj
Digits(1)=Array(LED1_1,LED1_2,LED1_3,LED1_4,LED1_5,LED1_6,LED1_7)
Digits(2)=Array(LED2_1,LED2_2,LED2_3,LED2_4,LED2_5,LED2_6,LED2_7)
Digits(3)=Array(LED3_1,LED3_2,LED3_3,LED3_4,LED3_5,LED3_6,LED3_7)
Digits(4)=Array(LED4_1,LED4_2,LED4_3,LED4_4,LED4_5,LED4_6,LED4_7)
Digits(5)=Array(LED5_1,LED5_2,LED5_3,LED5_4,LED5_5,LED5_6,LED5_7)
Digits(6)=Array(LED6_1,LED6_2,LED6_3,LED6_4,LED6_5,LED6_6,LED6_7)
Digits(7)=Array(LED7_1,LED7_2,LED7_3,LED7_4,LED7_5,LED7_6,LED7_7)
Digits(8)=Array(LED8_1,LED8_2,LED8_3,LED8_4,LED8_5,LED8_6,LED8_7)

Sub Display(obj,DText)
  Select case DText
    case "0": obj(0).state = 1:obj(1).state = 1:obj(2).state = 1:obj(3).state = 1:obj(4).state = 1:obj(5).state = 1:obj(6).state = 0
    case "1": obj(0).state = 0:obj(1).state = 1:obj(2).state = 1:obj(3).state = 0:obj(4).state = 0:obj(5).state = 0:obj(6).state = 0
    case "2": obj(0).state = 1:obj(1).state = 1:obj(2).state = 0:obj(3).state = 1:obj(4).state = 1:obj(5).state = 0:obj(6).state = 1
    case "3": obj(0).state = 1:obj(1).state = 1:obj(2).state = 1:obj(3).state = 1:obj(4).state = 0:obj(5).state = 0:obj(6).state = 1
    case "4": obj(0).state = 0:obj(1).state = 1:obj(2).state = 1:obj(3).state = 0:obj(4).state = 0:obj(5).state = 1:obj(6).state = 1
    case "5": obj(0).state = 1:obj(1).state = 0:obj(2).state = 1:obj(3).state = 1:obj(4).state = 0:obj(5).state = 1:obj(6).state = 1
    case "6": obj(0).state = 1:obj(1).state = 0:obj(2).state = 1:obj(3).state = 1:obj(4).state = 1:obj(5).state = 1:obj(6).state = 1
    case "7": obj(0).state = 1:obj(1).state = 1:obj(2).state = 1:obj(3).state = 0:obj(4).state = 0:obj(5).state = 0:obj(6).state = 0
    case "8": obj(0).state = 1:obj(1).state = 1:obj(2).state = 1:obj(3).state = 1:obj(4).state = 1:obj(5).state = 1:obj(6).state = 1
    case "9": obj(0).state = 1:obj(1).state = 1:obj(2).state = 1:obj(3).state = 1:obj(4).state = 0:obj(5).state = 1:obj(6).state = 1
    case "P": obj(0).state = 1:obj(1).state = 1:obj(2).state = 0:obj(3).state = 0:obj(4).state = 1:obj(5).state = 1:obj(6).state = 1
    case "L": obj(0).state = 0:obj(1).state = 0:obj(2).state = 0:obj(3).state = 1:obj(4).state = 1:obj(5).state = 1:obj(6).state = 0
    case "A": obj(0).state = 1:obj(1).state = 1:obj(2).state = 1:obj(3).state = 0:obj(4).state = 1:obj(5).state = 1:obj(6).state = 1
    case "Y": obj(0).state = 0:obj(1).state = 1:obj(2).state = 1:obj(3).state = 1:obj(4).state = 0:obj(5).state = 1:obj(6).state = 1
    case "I": obj(0).state = 0:obj(1).state = 1:obj(2).state = 1:obj(3).state = 0:obj(4).state = 0:obj(5).state = 0:obj(6).state = 0
    case "C": obj(0).state = 1:obj(1).state = 0:obj(2).state = 0:obj(3).state = 1:obj(4).state = 1:obj(5).state = 1:obj(6).state = 0
    case "O": obj(0).state = 1:obj(1).state = 1:obj(2).state = 1:obj(3).state = 1:obj(4).state = 1:obj(5).state = 1:obj(6).state = 0
    case "D": obj(0).state = 0:obj(1).state = 1:obj(2).state = 1:obj(3).state = 1:obj(4).state = 1:obj(5).state = 0:obj(6).state = 1
    case "U": obj(0).state = 0:obj(1).state = 1:obj(2).state = 1:obj(3).state = 1:obj(4).state = 1:obj(5).state = 1:obj(6).state = 0
    case "N": obj(0).state = 0:obj(1).state = 0:obj(2).state = 1:obj(3).state = 0:obj(4).state = 1:obj(5).state = 0:obj(6).state = 1
    case "H": obj(0).state = 0:obj(1).state = 1:obj(2).state = 1:obj(3).state = 0:obj(4).state = 1:obj(5).state = 1:obj(6).state = 1
    case "B": obj(0).state = 1:obj(1).state = 1:obj(2).state = 1:obj(3).state = 1:obj(4).state = 1:obj(5).state = 1:obj(6).state = 1
    case "E": obj(0).state = 1:obj(1).state = 0:obj(2).state = 0:obj(3).state = 1:obj(4).state = 1:obj(5).state = 1:obj(6).state = 1
    case "F": obj(0).state = 1:obj(1).state = 0:obj(2).state = 0:obj(3).state = 0:obj(4).state = 1:obj(5).state = 1:obj(6).state = 1
    case "R": obj(0).state = 1:obj(1).state = 1:obj(2).state = 1:obj(3).state = 0:obj(4).state = 1:obj(5).state = 1:obj(6).state = 1
    case "S": obj(0).state = 1:obj(1).state = 0:obj(2).state = 1:obj(3).state = 1:obj(4).state = 0:obj(5).state = 1:obj(6).state = 1
    case "r": obj(0).state = 0:obj(1).state = 0:obj(2).state = 0:obj(3).state = 0:obj(4).state = 1:obj(5).state = 0:obj(6).state = 1
    case " ": obj(0).state = 0:obj(1).state = 0:obj(2).state = 0:obj(3).state = 0:obj(4).state = 0:obj(5).state = 0:obj(6).state = 0
    case "-": obj(0).state = 0:obj(1).state = 0:obj(2).state = 0:obj(3).state = 0:obj(4).state = 0:obj(5).state = 0:obj(6).state = 1
  end select
End Sub


'---------------------------
'----- High Score Code -----
'---------------------------
HSArray = Array("HS_0","HS_1","HS_2","HS_3","HS_4","HS_5","HS_6","HS_7","HS_8","HS_9","HS_Space","HS_Comma","Tape")
Const HighScoreFilename = "KesselRunHighscore.txt"

Dim HSAHighScore, HSA1, HSA2, HSA3
Dim HSScore100k, HSScore10k, HSScoreK, HSScore100, HSScore10, HSScore1, HSScorex  'Define 5 different score values for each reel to use
Const DefaultHighScore = 0

LoadHighScore
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
    HSScore100K=Int (HSScorex/100000)'Calculate the value for the 100,000's digit
    HSScore10K=Int ((HSScorex-(HSScore100k*100000))/10000) 'Calculate the value for the 10,000's digit
    HSScoreK=Int((HSScorex-(HSScore100k*100000)-(HSScore10K*10000))/1000) 'Calculate the value for the 1000's digit
    HSScore100=Int((HSScorex-(HSScore100k*100000)-(HSScore10K*10000)-(HSScoreK*1000))/100) 'Calculate the value for the 100's digit
    HSScore10=Int((HSScorex-(HSScore100k*100000)-(HSScore10K*10000)-(HSScoreK*1000)-(HSScore100*100))/10) 'Calculate the value for the 10's digit
    HSScore1=Int(HSScorex-(HSScore100k*100000)-(HSScore10K*10000)-(HSScoreK*1000)-(HSScore100*100)-(HSScore10*10)) 'Calculate the value for the 1's digit

    PScore1.image = HSArray(HSScore100K):If HSScorex<100000 Then PScore1.image = HSArray(10)
    PScore2.image = HSArray(HSScore10K):If HSScorex<10000 Then PScore2.image = HSArray(10)
    PScore3.image = HSArray(HSScoreK):If HSScorex<1000 Then PScore3.image = HSArray(10)
    PScore4.image = HSArray(HSScore100):If HSScorex<100 Then PScore4.image = HSArray(10)
    PScore5.image = HSArray(HSScore10):If HSScorex<10 Then PScore5.image = HSArray(10)
    PScore6.image = HSArray(HSScore1):If HSScorex<1 Then PScore6.image = HSArray(10)
    if HSScorex<1000 then
      PComma.image = HSArray(10)
    else
      PComma.image = HSArray(11)
    end if
End Sub

Sub SaveHighScore
  Dim FileObj
  Dim ScoreFile
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

Sub CheckNewHighScorePostIt (newScore)
    If CLng(newScore) > CLng(HSAHighScore) Then
      playsound SoundFX("Knocker")
      DOF 105, 2
      HSAHighScore=newScore
      SaveHighScore
      UpdatePostIt
    End If
End Sub


'---------------------------
'--- DOF code by Arngrim ---
'---------------------------

'Dim B2SLights
Dim TextStr2,B2SOn,DOFs,Controller

sub SaveLMEMConfig
  Dim FileObj
  Dim LMConfig
  dim temp1
  dim tempb2s
  tempb2s=0
  if B2SOn=true then
    if DOFs = true then
      tempb2s=2
    else
      tempb2s=1
    end if
  else
    tempb2s=0
  end if
  Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if
  Set LMConfig=FileObj.CreateTextFile(UserDirectory & LMEMTableConfig,True)
  LMConfig.WriteLine tempb2s
  LMConfig.Close
  Set LMConfig=Nothing
  Set FileObj=Nothing

end Sub

sub LoadLMEMConfig
  Dim FileObj
  Dim LMConfig
  dim tempC
  dim tempb2s

    Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if
  If Not FileObj.FileExists(UserDirectory & LMEMTableConfig) then
    Exit Sub
  End if
  Set LMConfig=FileObj.GetFile(UserDirectory & LMEMTableConfig)
  Set TextStr2=LMConfig.OpenAsTextStream(1,0)
  If (TextStr2.AtEndOfStream=True) then
    Exit Sub
  End if
  tempC=TextStr2.ReadLine
  TextStr2.Close
  tempb2s=cdbl(tempC)
  if tempb2s=0 then
    B2SOn=false
    DOFs = false
  elseif tempb2s= 1 then
    B2SOn=true
    DOFs = false
  elseif tempb2s= 2 then
    B2SOn=true
    DOFs = true
  end if
  Set LMConfig=Nothing
  Set FileObj=Nothing
end sub

Function SoundFX (sound)
    If DOFs = true Then
        SoundFX = ""
    Else
        SoundFX = sound
    End If
End Function

Sub DOF(dofevent, dofstate)
  If B2SOn = True Then
    If dofstate = 2 Then
      Controller.B2SSetData dofevent, 1:Controller.B2SSetData dofevent, 0
    Else
      Controller.B2SSetData dofevent, dofstate
    End If
  End If
End Sub


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

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
    tmp = tableobj.y * 2 / table1.height-1
    If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = tableobj.x * 2 / table1.width-1
    If tmp > 0 Then
        AudioPan = Csng(tmp ^10)
    Else
        AudioPan = Csng(-((- tmp) ^10) )
    End If
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


'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 5 ' total number of balls
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
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
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
  PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
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


Sub Pins_Hit (idx)
  PlaySound "pinhit_low", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
  PlaySound "target", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
  PlaySound "metalhit_thin", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
  PlaySound "metalhit_medium", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
  PlaySound "metalhit2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
  PlaySound "gate4", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinner_Spin
  PlaySound "fx_spinner", 0, .25, AudioPan(Spinner), 0.25, 0, 0, 1, AudioFade(Spinner)
End Sub

Sub Rubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 20 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End Sub

Sub Posts_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    RandomSoundRubber()
  End If
End Sub

Sub RandomSoundRubber()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

'*****************STAR WARS CHANGES***********************
dim successound
dim successstate
Sub success
  starframe=0
  successstate = true
  dim soundx
  soundx = successound
  successound = INT ( 7* rnd)
  if soundx = successound Then
    successound = successound + 1
    if successound = 7 then successound = 0
  end If
  If ActiveHole<> 12 Then playsound ("h" & successound)
  'starfield.interval = 5
End Sub
dim wrongsound
Sub wrongholesound
  dim soundx
  soundx = wrongsound
  wrongsound = INT ( 15* rnd)
  if soundx = wrongsound Then
    wrongsound = wrongsound + 1
    if wrongsound = 15 then wrongsound = 0
  end If
  playsound wrongsound
End Sub


Sub Musictimer_Timer()
  stopsound "gameplay"
  playsound "gameplay"
  Musictimer.Enabled = False
End Sub


'********************************************************************************
' Flex DMD routines made possible by scutters' tutorials and scripts.
'********************************************************************************

Dim FlexDMDScene
Dim ExternalEnabled


Dim FlexDMD   ' the flex dmd display
Dim flexscore, Flexpath
Dim fso,curdir
Dim Scorearray, bonusarray, flexbonus
Dim starframe, starblur
Sub FlexDMD_Init() 'default/startup values
  starframe =0
  starblur = 0
  Scorearray=Array(0,0,0,0)
  bonusarray=Array(0,0,0)
  Set fso = CreateObject("Scripting.FileSystemObject")
  curDir = fso.GetAbsolutePathName(".")
  FlexPath = "VPX."

  Set FlexDMD = CreateObject("FlexDMD.FlexDMD")
  If Not FlexDMD is Nothing Then

    FlexDMD.GameName = "Kessel Run"
    FlexDMD.RenderMode = 2
    FlexDMD.Width = 128
    FlexDMD.Height = 32
    FlexDMD.Clear = True
    FlexDMD.Run = True
    FlexDMD.TableFile = Table1.Filename & ".vpx"
    Set FlexDMDScene = FlexDMD.NewGroup("Scene")


    With FlexDMDScene

      .Addactor FlexDMD.Newimage("starframe", "VPX.spritesheet&region=0," & (starframe * 32) & ",128,32")

      .AddActor FlexDMD.Newimage("10K",FlexPath & "8")
      .Getimage("10K").SetAlignedPosition  1,1,0

      .AddActor FlexDMD.Newimage("1000",FlexPath & "1")
      .Getimage("1000").SetAlignedPosition  26,1,0

      .AddActor FlexDMD.Newimage("100",FlexPath & "2")
      .Getimage("100").SetAlignedPosition  51,1,0

      .AddActor FlexDMD.Newimage("10",FlexPath & "0")
      .Getimage("10").SetAlignedPosition  76,1,0

      .AddActor FlexDMD.Newimage("1",FlexPath & "0")
      .Getimage("1").SetAlignedPosition  101,1,0

      .AddActor FlexDMD.NewImage("Ballbonus",FlexPath & "ball bonus")
      .GetImage("Ballbonus").SetAlignedPosition  0,0,0

      .AddActor FlexDMD.Newimage("ball",FlexPath & "s3")
      .Getimage("ball").SetAlignedPosition  39,21,0

      .AddActor FlexDMD.Newimage("bonusK",FlexPath & "s1")
      .Getimage("bonusK").SetAlignedPosition  94,21,0

      .AddActor FlexDMD.Newimage("bonus100",FlexPath & "s0")
      .Getimage("bonus100").SetAlignedPosition  102,21,0

      .AddActor FlexDMD.Newimage("bonus10",FlexPath & "s0")
      .Getimage("bonus10").SetAlignedPosition  110,21,0

      .AddActor FlexDMD.Newimage("credits",FlexPath & "s0")
      .Getimage("credits").SetAlignedPosition  81,21,0
      .GetImage("credits").Visible = False

      .AddActor FlexDMD.NewImage("Title",FlexPath & "title")
      .GetImage("Title").SetAlignedPosition  0,0,0
      .GetImage("Title").Visible = True


      .addactor FlexDMD.Newimage("creditscroll", "VPX.credits&region=0," & (creditframe) & ",128,32")
      .GetImage("creditscroll").Visible = False



    End With

    FlexDMD.LockRenderThread

    FlexDMD.Stage.AddActor FlexDMDScene

    FlexDMD.Show = True
    FlexDMD.UnlockRenderThread


  End If

End Sub




'**************
' Update FlexDMD
'**************

Dim digitoffset
Sub FlexDMDUpdate()

  if UseFlexDMD then
  If Not FlexDMD is Nothing Then FlexDMD.LockRenderThread
  If FlexDMD.Run = False Then FlexDMD.Run = True

' Build Scorearray =
Dim flexscore
flexscore= Score
flexbonus= Bonus
If flexscore> 99999 Then flexscore = flexscore - ( INT (flexscore /100000)*100000)

Scorearray(0) = INT (flexscore/10000)
Scorearray(1) = INT ((flexscore -  (Scorearray(0)*10000))/1000)
Scorearray(2) = INT ((flexscore -  (Scorearray(0)*10000) - (Scorearray(1)*1000))/100)
Scorearray(3) = INT ((flexscore -  (Scorearray(0)*10000) - (Scorearray(1)*1000) - (Scorearray(2)*100) )/10)

bonusarray(0) = INT (flexbonus/1000)
bonusarray(1) = INT ((flexbonus - (bonusarray(0)*1000))/100)
bonusarray(2) = INT ((flexbonus - (bonusarray(0)*1000)- bonusarray(1)*100)/10)

if bonusarray(0) = 0 then
  bonusarray(0) = 11
  if bonusarray(1) = 0 then
    bonusarray(1) = 11
  end if
end if
  With FlexDMD.Stage
      .GetImage("Title").Visible = False
      .GetImage("creditscroll").Visible = False
    .GetImage("Ballbonus").Visible = true
  If Scorearray(0) = 0 then
    Scorearray(0) = 11
    digitoffset = 12
    If Scorearray(1) = 0 then
    Scorearray(1) = 11
    digitoffset = 25
      If Scorearray(2) = 0 then
      Scorearray(2) = 11
      digitoffset = 37
      End If
    End If
  Else
    digitoffset= 0
  End If



    .GetImage("10K").Bitmap = FlexDMD.NewImage("10K",FlexPath & Scorearray(0)).Bitmap
    .GetImage("1000").Bitmap = FlexDMD.NewImage("1000",FlexPath & Scorearray(1)).Bitmap
    .GetImage("100").Bitmap = FlexDMD.NewImage("100",FlexPath & Scorearray(2)).Bitmap
    .GetImage("10").Bitmap = FlexDMD.NewImage("10",FlexPath & Scorearray(3)).Bitmap
    .GetImage("bonusK").Bitmap = FlexDMD.NewImage("bonusK",FlexPath & "s"& bonusarray(0)).Bitmap
    .GetImage("bonus100").Bitmap = FlexDMD.NewImage("bonus100",FlexPath & "s"& bonusarray(1)).Bitmap
    .GetImage("bonus10").Bitmap = FlexDMD.NewImage("bonus10",FlexPath & "s"& bonusarray(2)).Bitmap

    .GetImage("10K").SetAlignedPosition (1 - digitoffset),1,0
    .GetImage("1000").SetAlignedPosition (26 - digitoffset),1,0
    .GetImage("100").SetAlignedPosition (51 - digitoffset),1,0
    .GetImage("10").SetAlignedPosition (76 - digitoffset),1,0
    .GetImage("1").SetAlignedPosition (101 - digitoffset),1,0


    .GetImage("ball").Bitmap = FlexDMD.NewImage("ball",FlexPath & "s" & BallInPlay).Bitmap


  End With

  If Not FlexDMD is Nothing Then FlexDMD.UnlockRenderThread

End If

End Sub


Sub starfield_Timer()
  If UseFlexDMD then


  If Not FlexDMD is Nothing Then FlexDMD.LockRenderThread
  If FlexDMD.Run = False Then FlexDMD.Run = True
  With FlexDMD.Stage
  if successstate=true Then.GetImage("starframe").Bitmap =FlexDMD.Newimage("starframe", "VPX.spritesheetb&region=0," & (starframe * 32) & ",128,32").Bitmap
  if successstate=false Then.GetImage("starframe").Bitmap =FlexDMD.Newimage("starframe", "VPX.spritesheet&region=0," & (starframe * 32) & ",128,32").Bitmap
  End With
  If Not FlexDMD is Nothing Then FlexDMD.UnlockRenderThread
  End If

    starframe = starframe +1
    if starframe >49 Then
      starframe=0
      if successstate=true Then
        successstate= false
        'starfield.interval = 60
      End If
    End If
End Sub

Sub flexcredit
  If Not FlexDMD is Nothing Then FlexDMD.LockRenderThread
  If FlexDMD.Run = False Then FlexDMD.Run = True
  With FlexDMD.Stage
    .GetImage("Ballbonus").Bitmap =FlexDMD.Newimage("Ballbonus", "VPX.credit").Bitmap
    cleardigits
    .GetImage("credits").Visible = True
  creditframe=0
  credittimer.Enabled= False
  .GetImage("creditscroll").Visible = False
  .GetImage("Ballbonus").Visible = true
  .GetImage("credits").Bitmap = FlexDMD.NewImage("credits",FlexPath & "s" & Credits).Bitmap

  End With
  If Not FlexDMD is Nothing Then FlexDMD.UnlockRenderThread
End Sub

Sub cleardigits
  If Not FlexDMD is Nothing Then FlexDMD.LockRenderThread
  If FlexDMD.Run = False Then FlexDMD.Run = True
  With FlexDMD.Stage
    .GetImage("10K").Visible = false
    .GetImage("1000").Visible = false
    .GetImage("100").Visible = false
    .GetImage("10").Visible = false
    .GetImage("1").Visible = false
    .GetImage("ball").Visible = false
    .GetImage("bonusK").Visible = false
    .GetImage("bonus100").Visible = false
    .GetImage("bonus10").Visible = false
  End With
  If Not FlexDMD is Nothing Then FlexDMD.UnlockRenderThread
End Sub

Sub judgeplayer
  If Not FlexDMD is Nothing Then FlexDMD.LockRenderThread
  If FlexDMD.Run = False Then FlexDMD.Run = True
  With FlexDMD.Stage
    .GetImage("ball").Visible = false
    .GetImage("bonusK").Visible = false
    .GetImage("bonus100").Visible = false
    .GetImage("bonus10").Visible = false
  .GetImage("Ballbonus").Visible = True
  If Score >=2500 then
    .GetImage("Ballbonus").Bitmap =FlexDMD.Newimage("Ballbonus", "VPX.niceman").Bitmap
  Elseif score >= 1750 Then
    .GetImage("Ballbonus").Bitmap =FlexDMD.Newimage("Ballbonus", "VPX.smuggler").Bitmap
  Elseif score >= 1000 Then
    .GetImage("Ballbonus").Bitmap =FlexDMD.Newimage("Ballbonus", "VPX.scoundrel").Bitmap
  Elseif score >= 500 Then
    .GetImage("Ballbonus").Bitmap =FlexDMD.Newimage("Ballbonus", "VPX.scruffy").Bitmap
  Else.GetImage("Ballbonus").Bitmap =FlexDMD.Newimage("Ballbonus", "VPX.nerfherder").Bitmap
  End If
  If Score >= HSAHighScore then .GetImage("Ballbonus").Bitmap =FlexDMD.Newimage("Ballbonus", "VPX.best").Bitmap
  End With
  If Not FlexDMD is Nothing Then FlexDMD.UnlockRenderThread
  attractmode.Enabled = true
End Sub



Sub backtogame
  If Not FlexDMD is Nothing Then FlexDMD.LockRenderThread
  If FlexDMD.Run = False Then FlexDMD.Run = True
  With FlexDMD.Stage
    .GetImage("Ballbonus").Bitmap =FlexDMD.Newimage("Ballbonus", "VPX.ball bonus").Bitmap
    .GetImage("10K").Visible = True
    .GetImage("1000").Visible = True
    .GetImage("100").Visible = True
    .GetImage("10").Visible = True
    .GetImage("1").Visible = True
    .GetImage("ball").Visible = True
    .GetImage("bonusK").Visible = True
    .GetImage("bonus100").Visible = True
    .GetImage("bonus10").Visible = True
    .GetImage("credits").Visible = False

  .GetImage("Ballbonus").Visible = true
    credittimer.enabled = False
    creditframe = 0
    .GetImage("creditscroll").Visible = False

    FlexDMDUpdate

  End With
  If Not FlexDMD is Nothing Then FlexDMD.UnlockRenderThread
End Sub
Dim attractstatus
attractstatus = 1

Sub attractmode_Timer()

  Select Case attractstatus
  Case 1
    Highscoremode
    attractstatus = 2
  Case 2
    FlexDMDUpdate
    judgeplayer
    attractstatus = 3
  Case 3
  cleardigits
  scrollcredits
    attractstatus = 1
  attractmode.enabled = False
  End Select

End Sub
Sub Highscoremode
  backtogame
  If Not FlexDMD is Nothing Then FlexDMD.LockRenderThread
  If FlexDMD.Run = False Then FlexDMD.Run = True
  With FlexDMD.Stage
    .GetImage("Ballbonus").Bitmap =FlexDMD.Newimage("Ballbonus", "VPX.best").Bitmap

    .GetImage("ball").Visible = False
    .GetImage("bonusK").Visible = False
    .GetImage("bonus100").Visible = False
    .GetImage("bonus10").Visible = False
' Build Scorearray =
Dim flexscore
flexscore= HSAHighScore
If flexscore> 99999 Then flexscore = flexscore - ( INT (flexscore /100000)*100000)

Scorearray(0) = INT (flexscore/10000)
Scorearray(1) = INT ((flexscore -  (Scorearray(0)*10000))/1000)
Scorearray(2) = INT ((flexscore -  (Scorearray(0)*10000) - (Scorearray(1)*1000))/100)
Scorearray(3) = INT ((flexscore -  (Scorearray(0)*10000) - (Scorearray(1)*1000) - (Scorearray(2)*100) )/10)



  If Scorearray(0) = 0 then
    Scorearray(0) = 11
    digitoffset = 12
    If Scorearray(1) = 0 then
    Scorearray(1) = 11
    digitoffset = 25
      If Scorearray(2) = 0 then
      Scorearray(2) = 11
      digitoffset = 37
      End If
    End If
  Else
    digitoffset= 0
  End If



    .GetImage("10K").Bitmap = FlexDMD.NewImage("10K",FlexPath & Scorearray(0)).Bitmap
    .GetImage("1000").Bitmap = FlexDMD.NewImage("1000",FlexPath & Scorearray(1)).Bitmap
    .GetImage("100").Bitmap = FlexDMD.NewImage("100",FlexPath & Scorearray(2)).Bitmap
    .GetImage("10").Bitmap = FlexDMD.NewImage("10",FlexPath & Scorearray(3)).Bitmap

    .GetImage("10K").SetAlignedPosition (1 - digitoffset),1,0
    .GetImage("1000").SetAlignedPosition (26 - digitoffset),1,0
    .GetImage("100").SetAlignedPosition (51 - digitoffset),1,0
    .GetImage("10").SetAlignedPosition (76 - digitoffset),1,0
    .GetImage("1").SetAlignedPosition (101 - digitoffset),1,0

  End With
  If Not FlexDMD is Nothing Then FlexDMD.UnlockRenderThread

End Sub


Sub scrollcredits
  credittimer.Enabled= true
  If UseFlexDMD then
  If Not FlexDMD is Nothing Then FlexDMD.LockRenderThread
  If FlexDMD.Run = False Then FlexDMD.Run = True
  With FlexDMD.Stage

  .GetImage("Ballbonus").Visible = false
  .GetImage("creditscroll").Bitmap =FlexDMD.Newimage("creditscroll", "VPX.credits&region=0," & (creditframe) & ",128,32").Bitmap

      .GetImage("creditscroll").Visible = True
  End With
  If Not FlexDMD is Nothing Then FlexDMD.UnlockRenderThread
  End If
End Sub

dim creditframe
creditframe= 0

Sub credittimer_Timer()
  creditframe=creditframe+1
  scrollcredits
  if creditframe = 742 Then
  creditframe=0
  credittimer.Enabled= False
  if UseFlexDMD Then
  backtogame
  judgeplayer
  End If
  End If
End Sub

'*
'*        Gottlieb's Road Race (1969)
'*
'*
'*

option explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("Controller.vbs")
If Err Then MsgBox "Unable to open Controller.vbs. Ensure that it is in the scripts folder."
On Error Goto 0

Const ballsize = 50
Const ballmass = 1

Const cGameName = "RoadRace_1969"
Const HSFileName="RoadRace_69VPX.txt"

'******************************************************
'             OPTIONS
'******************************************************

'///////////////////////-----General Sound Options-----///////////////////////
'//  VolumeDial:
'//  VolumeDial is the actual global volume multiplier for the mechanical sounds.
'//  Values smaller than 1 will decrease mechanical sounds volume.
'//  Recommended values should be no greater than 1.
Const VolumeDial = 0.8

'* Set to 1 to disable manual ball control with the keyboard
Const DisableBallControl= 1

'******************************************************
'             VARIABLES
'******************************************************

Dim i
Dim obj

Dim InProgress:InProgress=false
Dim BallInPlay
Dim BallsPerGame:BallsPerGame=5
Dim Player, HighPlayer
Dim PrevScore, Score, isScoring: isScoring = False
Dim MotorRunning:MotorRunning = 0
Dim Credits:Credits=0
Dim Match
Dim TableTilted: TableTilted = False
Dim TiltCount
Dim TiltEndsGame: TiltEndsGame = 1
Dim GameOn:GameOn = False
Dim GameOver: GameOver = False

dim ScoreChecker
dim CheckAllScores
Dim TextStr

Dim HighScore: Highscore = 0
Dim HighScorePaid
Dim HighScoreReward:HighScoreReward = 1

Dim Replay1, Replay2, Replay3, Replay4
Dim Replay1Paid, Replay2Paid, Replay3Paid, Replay4Paid
Dim Replay1Table(15), Replay2Table(15), Replay3Table(15), Replay4Table(15)
Dim ReplayLevel:ReplayLevel=1
Dim ReplayTableMax

Dim OperatorMenu:OperatorMenu = 0

Dim ChimesOn: ChimesOn = 1
Dim LastChime10, LastChime100, LastChime1000

dim raceRflag, raceAflag, raceCflag, raceEflag

Dim LStep, RStep, xx

Dim TargetSequenceComplete
Dim ZeroToNineUnit

Dim TableWidth, TableHeight
TableWidth = Table1.width
TableHeight = Table1.height

Dim Desktopmode:Desktopmode = Table1.ShowDT

Dim ArrowBall:ArrowBall = Empty

'******************************************************
'             TABLE INIT
'******************************************************

Sub Table1_Init()
  LoadEM

  If Desktopmode = false then
    For each obj in DesktopCrap
      obj.visible=False
    next
  End If

  OperatorMenuBackdrop.image = "PostitBL"

  For XOpt = 1 to MaxOption
    Eval("OperatorOption"&XOpt).image = "PostitBL"
  next

  For XOpt = 1 to 256
    Eval("Option"&XOpt).image = "PostItBL"
  next

  loadhs
  if HighScore=0 then HighScore=5000

  For each obj in PlayerScores
    obj.ResetToZero
  next

  TiltReel.SetValue(0)
  ZeroToNineUnit=Int(Rnd*10)
  Match=int(Rnd*10)
  MatchReel.SetValue((Match)+1)

  BumpersOff

  For each obj in NumberLights
    obj.state=0
  next

  SetupReplayTables

  Replay1=Replay1Table(ReplayLevel)
  Replay2=Replay2Table(ReplayLevel)
  Replay3=Replay3Table(ReplayLevel)
  Replay4=Replay4Table(ReplayLevel)

  RefreshReplayCard

  InstructCard.image="IC"+FormatNumber(BallsPerGame,0)+"0"

  InitPauser5.enabled=true
  If Credits > 0 Then DOF 109, 1
End Sub

Sub Table1_exit()
  savehs
  If B2SOn Then Controller.Stop
end sub

sub InitPauser5_timer
  CreditsReel.SetValue(Credits)

  If B2SOn Then
    if Match=0 then
      Controller.B2SSetMatch 10
    else
      Controller.B2SSetMatch Match
    end if
    Controller.B2SSetTilt 0
    Controller.B2SSetCredits Credits
    Controller.B2SSetGameOver 1
    Controller.B2SSetScorePlayer 1, 0
  End If
  InitPauser5.enabled=false
end sub

'******************************************************
'             B2S LIGHTS
'******************************************************

Sub TimerR_timer
  if B2SOn then
    if raceRflag=1 then
      Controller.B2SSetData 11,0
      raceRflag=0
    else
      Controller.B2SSetData 11,1
      raceRflag=1
    end if
  end if
end sub

Sub TimerA_timer
  if B2SOn then
    if raceAflag=1 then
      Controller.B2SSetData 12,0
      raceAflag=0
    else
      Controller.B2SSetData 12,1
      raceAflag=1
    end if
  end if
end sub

Sub TimerC_timer
  if B2SOn then
    if raceCflag=1 then
      Controller.B2SSetData 13,0
      raceCflag=0
    else
      Controller.B2SSetData 13,1
      raceCflag=1
    end if
  end if
end sub

Sub TimerE_timer
  if B2SOn then
    if raceEflag=1 then
      Controller.B2SSetData 14,0
      raceEflag=0
    else
      Controller.B2SSetData 14,1
      raceEflag=1
    end if
  end if
end sub

'******************************************************
'             KEYS
'******************************************************
Const ReflipAngle = 20

Sub Table1_KeyDown(ByVal keycode)

  ' GNMOD
  if EnteringInitials then
    CollectInitials(keycode)
    exit sub
  end if

  if EnteringOptions then
    CollectOptions(keycode)
    exit sub
  end if

  If keycode = PlungerKey Then
    Plunger.PullBack
    SoundPlungerPull
    PlungerPulled = 1
  End If

  if keycode = LeftFlipperKey and InProgress = false then
    OperatorMenuTimer.Enabled = true
  end if
  ' END GNMOD

  If keycode = LeftFlipperKey and InProgress=true and TableTilted=false Then
    FlipperActivate LeftFlipper, LFPress
    LeftFlipper.RotateToEnd
    PlaySoundAtVolLoops "buzzL",LeftFlipper,0.05,-1

    If leftflipper.currentangle < leftflipper.endangle + ReflipAngle Then
      RandomSoundReflipUpLeft LeftFlipper
    Else
      SoundFlipperUpAttackLeft LeftFlipper
      RandomSoundFlipperUpLeft LeftFlipper
    End If

    DOF 101, 1
  End If

  If keycode = RightFlipperKey  and InProgress=true and TableTilted=false Then
    FlipperActivate RightFlipper, RFPress
    RightFlipper.RotateToEnd

    If rightflipper.currentangle > rightflipper.endangle - ReflipAngle Then
      RandomSoundReflipUpRight RightFlipper
    Else
      SoundFlipperUpAttackRight RightFlipper
      RandomSoundFlipperUpRight RightFlipper
    End If

    DOF 102, 1
    PlaySoundAtVolLoops "buzz",RightFlipper,0.05,-1
  End If

  If keycode = LeftTiltKey Then
    Nudge 90, 1
    SoundNudgeLeft()
    TiltIt
  End If

  If keycode = RightTiltKey Then
    Nudge 270, 1
    SoundNudgeRight()
    TiltIt
  End If

  If keycode = CenterTiltKey Then
    SoundNudgeCenter()
    Nudge 0, 1
    TiltIt
  End If

  If keycode = MechanicalTilt Then
    TiltCount=2
    TiltIt
  End If

  If keycode = AddCreditKey or keycode = 4 or keycode = 5 then
    coindelay.enabled=true
  end if

    If keycode = 46 and DisableBallControl <> 1 Then' C Key
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

  if keycode=StartGameKey and Credits>0 and InProgress=false and EnteringOptions = 0 then
    'GNMOD
    OperatorMenuTimer.Enabled = false
    'END GNMOD

    AddCredit -1
    SoundStartButton

    InProgress = True

    rst=0
    resettimer.enabled=true
    playsound "startup_norm"
  end if
End Sub

Sub Table1_KeyUp(ByVal keycode)

  ' GNMOD
  if EnteringInitials then
    exit sub
  end if

  If keycode = PlungerKey Then
    if PlungerPulled = 0 then exit sub
    If not isempty(arrowball) then
      If (ArrowBall.x > 860 and ArrowBall.y > 1650) Then
        SoundPlungerReleaseBall
      Else
        SoundPlungerReleaseNoBall
      End If
    Else
      SoundPlungerReleaseNoBall
    End If
    Plunger.Fire
  End If

  if keycode = LeftFlipperKey then
    OperatorMenuTimer.Enabled = false
  end if
  ' END GNMOD

  If keycode = LeftFlipperKey and InProgress=true and TableTilted=false Then
    FlipperDeActivate LeftFlipper, LFPress
    LeftFlipper.RotateToStart

    If LeftFlipper.currentangle < LeftFlipper.startAngle - 5 Then
      RandomSoundFlipperDownLeft LeftFlipper
    End If
    FlipperLeftHitParm = FlipperUpSoundLevel

    DOF 101, 0
    StopSound "buzzL"
  End If

  If keycode = RightFlipperKey and InProgress=true and TableTilted=false Then
    FlipperActivate RightFlipper, RFPress
    RightFlipper.RotateToStart

    If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
      RandomSoundFlipperDownRight RightFlipper
    End If
    FlipperRightHitParm = FlipperUpSoundLevel

    DOF 102, 0
    StopSound "buzz"
  End If

  If keycode = 203 then cLeft = 0' Left Arrow
  If keycode = 200 then cUp = 0' Up Arrow
  If keycode = 208 then cDown = 0' Down Arrow
  If keycode = 205 then cRight = 0' Right Arrow
  If keycode = 52 Then Zup = 0' Period
End Sub

sub coindelay_timer
  AddCredit 1
  coindelay.enabled=false
  Select Case Int(rnd*3)
    Case 0: PlaySoundAtLevelStatic ("Coin_In_1"), CoinSoundLevel, drain
    Case 1: PlaySoundAtLevelStatic ("Coin_In_2"), CoinSoundLevel, drain
    Case 2: PlaySoundAtLevelStatic ("Coin_In_3"), CoinSoundLevel, drain
  End Select
end sub

'******************************************************
'           GAME TIMERS
'     Ball Rolling and Shadows, Primitive Updates,
'******************************************************

'***********************
'     Flipper Logos
'***********************

Const tnob = 1 ' total number of balls
ReDim rolling(tnob)

Dim BallShadow
BallShadow = Array (BallShadow1)


Sub GameTimer_Timer()
  '****** Flipper shadows and primitives ******
  LFLogo.RotZ = LeftFlipper.CurrentAngle
  RFLogo.RotZ = RightFlipper.CurrentAngle
  LFLogo1.RotZ = LeftFlipper.CurrentAngle
  RFLogo1.RotZ = RightFlipper.CurrentAngle
  FlipperLSh.RotZ = LeftFlipper.currentangle
  FlipperRSh.RotZ = RightFlipper.currentangle

  '****** Animated gate ******
  PGate.Rotz = (Gate.CurrentAngle*.75) + 25

  Dim BOT
  BOT = GetBalls

  ' exit the sub if no balls on the table
    If UBound(BOT) = -1 Then
    If rolling(0) = True Then
      StopSound("BallRoll_0")
      rolling(0) = False
            BallShadow(0).visible = 0
    End If
    Exit Sub
  End If

  ' play the rolling sound for each ball
  If BallVel(Arrowball) > 1 AND ArrowBall.z < 30 Then
    rolling(0) = True
    PlaySound ("BallRoll_0"), -1, VolPlayfieldRoll(ArrowBall) * 1.1 * VolumeDial, AudioPan(ArrowBall), 0, PitchPlayfieldRoll(ArrowBall), 1, 0, AudioFade(ArrowBall)
  Else
    If rolling(0) = True Then
      StopSound("BallRoll_0")
      rolling(0) = False
    End If
  End If

    ' render the shadow for each ball
  If Arrowball.X < TableWidth/2 Then
    Ballshadow(0).X = ((Arrowball.X) - (Ballsize/16) + ((Arrowball.X - (TableWidth/2))/17))' + 13
  Else
    Ballshadow(0).X = ((Arrowball.X) + (Ballsize/16) + ((Arrowball.X - (TableWidth/2))/17))' - 13
  End If
  Ballshadow(0).Y = Arrowball.Y + 10
  If Arrowball.Z > 20 Then
    Ballshadow(0).visible = 1
  Else
    Ballshadow(0).visible = 0
  End If
End Sub

'******************************************************
'           START NEW GAME
'******************************************************
dim rst

sub resettimer_timer
  rst=rst+1

  if rst>1 and rst<12 then
    ResetReelsToZero(1)
  end if

  if rst=13 then
    'playsound "StartBall1"
  end if

  if rst=14 then
    NewGame
    resettimer.enabled=false
  end if
end sub

Sub ResetReelsToZero(reelzeroflag)
  dim d1(5)
  dim d2(5)
  dim scorestring1, scorestring2

  If reelzeroflag=1 then
    scorestring1=CStr(Score)
    scorestring1=right("00000" & scorestring1,5)
    for i=0 to 4
      d1(i)=CInt(mid(scorestring1,i+1,1))
    next
    for i=0 to 4
      if d1(i)>0 then
        d1(i)=d1(i)+1
        if d1(i)>9 then d1(i)=0
      end if
    next
    Score=(d1(0)*10000) + (d1(1)*1000) + (d1(2)*100) + (d1(3)*10) + d1(4)

    If B2SOn Then
      Controller.B2SSetScorePlayer 1, Score
    End If
    PlayerScores(0).SetValue(Score)
    PlayerScoresOn(0).SetValue(Score)
  end if
end sub

Sub NewGame()
  Player=1
  BallInPlay=1
  InProgress=true

  MatchReel.SetValue(0)

  TimerR.enabled=0
  TimerA.enabled=0
  TimerC.enabled=0
  TimerE.enabled=0

  If B2SOn Then
    Controller.B2SSetTilt 0
    Controller.B2SSetGameOver 0
    Controller.B2SSetMatch 0

    Controller.B2SSetBallInPlay BallInPlay
    Controller.B2SSetScoreRolloverPlayer1 0

    Controller.B2SSetData 11,0
    Controller.B2SSetData 12,0
    Controller.B2SSetData 13,0
    Controller.B2SSetData 14,0
  End If

  PlayerScores(0).ResetToZero
  PlayerScoresOn(0).ResetToZero

  If Desktopmode Then
    PlayerScores(0).Visible=0
    PlayerScoresOn(0).Visible=1
  End If

  Score=0

  HighScorePaid=false
  Replay1Paid=false
  Replay2Paid=false
  Replay3Paid=false
  Replay4Paid=false

  for each obj in NumberLights
    obj.state=1
  next
  for each obj in AllRolloverSpecials
    obj.state=0
  next
  SpinnerSpecialLight.state=0

  TargetSequenceComplete=0
  ResetBalls
End Sub

'******************************************************
'           START NEW BALL
'******************************************************

Sub ResetBalls()
  TableTilted=false
  TiltReel.SetValue(0)
  If B2Son then
    Controller.B2SSetTilt 0
    'Controller.B2SSetBallInPlay BallInPlay
  end if

  InstructCard.image="IC"+FormatNumber(BallsPerGame,0)+FormatNumber(BallInPlay,0)

  set arrowball = Ballrelease.CreateSizedBallWithMass (Ballsize/2, BallMass)

    Ballrelease.Kick 80,11
  PlaysoundAtVol "StartBall2-5", Drain, BallReleaseSoundLevel
  DOF 111, 2
End Sub

'******************************************************
'             NEXT BALL
'******************************************************

sub NextBallDelay_timer()
  NextBallDelay.enabled=false
  nextball
end sub

sub nextball
  BallInPlay=BallInPlay+1

  If BallInPlay>BallsPerGame then
    PlaySoundat "MotorLeer", KnockerPosition
    InProgress=false

    If B2SOn Then
      Controller.B2SSetGameOver 1
      Controller.B2SSetPlayerUp 0
      Controller.B2SSetBallInPlay 0
      Controller.B2SSetCanPlay 0
    End If

    If Desktopmode Then
      PlayerScores(0).visible=1
      PlayerScoresOn(0).visible=0
    end If

    InstructCard.image="IC"+FormatNumber(BallsPerGame,0)+"0"

    LeftFlipper.RotateToStart
    RightFlipper.RotateToStart

    BumpersOff

    Checkmatch
    CheckHighScore

    TimerR.enabled=1
    TimerA.enabled=1
    TimerC.enabled=1
    TimerE.enabled=1

    HighScoreTimer.interval=100
    HighScoreTimer.enabled=True
  Else
    ResetBalls
  End If
End sub

'******************************************************
'             DRAIN
'******************************************************

Sub Drain_Hit()
  '******* for ball control script
    contBallinPlay = False
  arrowball = empty
  Drain.DestroyBall

  RandomSoundDrain Drain
  DOF 108, 2

  NextBallDelay.enabled = true
End Sub

Sub Trigger0_hit
  Set controlBall = ActiveBall
  contBallInPlay = True
End Sub

Sub Trigger0_Unhit()
  DOF 112, 2
End Sub


'******************************************************
'           RUBBERS/SLINGS
'******************************************************

Sub RightSlingShot_Slingshot
  RandomSoundSlingshotRight
  DOF 104, 2
    RSling0.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
  AddScore 1, 0
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSling0.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  RandomSoundSlingshotLeft
  DOF 103, 2
    LSling0.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
  AddScore 1, 0
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing0.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

'***********************************
' Walls
'***********************************

Sub RubberWallSwitches_Hit(idx)
  if TableTilted=false then
    Select Case idx
      Case 0,2: AddScore 10,0
      Case 1,3,4,5,6,7: AddScore 1,0
    end Select
  end if
end Sub


'******************************************************
'             BUMPER
'******************************************************

Sub Bumper1_Hit
  If TableTilted=false then
    RandomSoundBumper bumper1
    DOF 105, 2
    If Bumper1Light.state=1 then
      AddScore 10,0
    else
      AddScore 1,0
    end if
    end if
End Sub

Sub Bumper2_Hit
  If TableTilted=false then
    RandomSoundBumper bumper2
    DOF 106, 2
    If Bumper2Light.state=1 then
      AddScore 100, 0
    else
      AddScore 10, 0
    end if
    end if
End Sub

Sub Bumper3_Hit
  If TableTilted=false then
    RandomSoundBumper bumper3
    DOF 107, 2
    If Bumper3Light.state=1 then
      AddScore 10,0
    else
      AddScore 1,0
    end if
    end if
End Sub

Sub LightYellowBumpers
  Bumper2Light.state=0
  Bumper1Light.state=1
  Bumper3Light.state=1
  Bumper2Lighta.state=0
  Bumper1Lighta.state=1
  Bumper3Lighta.state=1
  BumpersOn
End Sub

Sub LightGreenBumpers
  Bumper2Light.state=1
  Bumper1Light.state=0
  Bumper3Light.state=0
  Bumper2Lighta.state=1
  Bumper1Lighta.state=0
  Bumper3Lighta.state=0
  BumpersOn
end sub

sub BumpersOff
  Bumper1Light.visible=0
  Bumper2Light.visible=0
  Bumper3Light.visible=0
  Bumper1Lighta.visible=0
  Bumper2Lighta.visible=0
  Bumper3Lighta.visible=0
end sub

sub BumpersOn
  Bumper1Light.visible=0
  Bumper2Light.visible=0
  Bumper3Light.visible=0
  Bumper1Lighta.visible=0
  Bumper2Lighta.visible=0
  Bumper3Lighta.visible=0

  If Bumper1Light.state=1 then
    Bumper1Light.visible=1
    bumper1Lighta.visible=1
  end if

  if Bumper2Light.state=1 then
    Bumper2Light.visible=1
    Bumper2Lighta.visible=1
  end if

  If Bumper3Light.state=1 then
    Bumper3Light.visible=1
    Bumper3Lighta.visible=1
  end if
end sub

'******************************************************
'           ROLL OVERS
'******************************************************

Sub TriggerCollection_Hit(idx)
  RandomSoundRollover
  If TableTilted=false then
    DOF 200 + idx, 2
    If NumberLights(idx).state=1 then
      MultiScore 2, 100
      NumberLights(idx).state=0
      NumberLights(idx+10).state=0
      if BallsPerGame=3 then
        if idx=0 then
          NumberLights(2).state=0
          NumberLights(12).state=0
        end if
        if idx=2 then
          NumberLights(0).state=0
          NumberLights(10).state=0
        end if
        if idx=1 then
          NumberLights(3).state=0
          NumberLights(13).state=0
        end if
        if idx=3 then
          NumberLights(1).state=0
          NumberLights(11).state=0
        end if
        if idx=8 then
          NumberLights(9).state=0
          NumberLights(19).state=0
        end if
        if idx=9 then
          NumberLights(8).state=0
          NumberLights(18).state=0
        end if
      end if
    else
      MultiScore 5, 10
    end if

    if idx=0 or idx=2 or idx=5 then
      LightYellowBumpers
      ToggleSpecial
    end if
    if idx=1 or idx=3 or idx=7 then
      LightGreenBumpers
      ToggleSpecial
    end if
    if idx=8 and SpecialLight9.state=1 then
      AddSpecial
    end if
    if idx=9 and SpecialLight10.state=1 then
      AddSpecial
    end if

    CheckSpecial
  end if

end Sub


'******************************************************
'           ARROW TARGET
'******************************************************

'*****************
' Math
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

Function Distance(ax,ay,bx,by)
  Distance = SQR((ax - bx)^2 + (ay - by)^2)
End Function

Function DistancePL(px,py,ax,ay,bx,by) ' Distance between a point and a line where point is px,py
  DistancePL = ABS((by - ay)*px - (bx - ax) * py + bx*ay - by*ax)/Distance(ax,ay,bx,by)
End Function

Function Radians(Degrees)
  Radians = Degrees * PI /180
End Function

Function rAnglePP(ax,ay,bx,by)
  rAnglePP = Atn2((by - ay),(bx - ax))
End Function

Function AnglePP(ax,ay,bx,by)
  AnglePP = Atn2((by - ay),(bx - ax))*180/PI
End Function

Function getangle(angle)
  If angle > 180 Then
    getangle = angle - 360
  elseif angle < -180 Then
    getangle = angle + 360
  else
    getangle = angle
  End If
End Function

'********************
' Arrow Target Subs
'********************

dim spinx, spiny, currang, prevang, spincount, startang

spinx = spinflipcw.x
spiny = spinflipcw.y
currang = rnd*360 - 180
prevang = currang

'*****************
' Physics Calcs
'*****************

sub spincalc_timer
  dim ballang, diffang, rotang, balldist, perpdist, cwstart, cwend, ccwstart, ccwend

  prevang = currang

  if IsEmpty(arrowball) Then
    ballang = 0
    diffang = 0
    balldist = 200
    perpdist = 200
  else
    ballang = anglepp(spinx, spiny, Arrowball.x, Arrowball.y) + 90
    If ballang > 180 then ballang = ballang - 360

    diffang = getangle(currang - ballang)

    balldist =  distance(spinx, spiny, Arrowball.x, Arrowball.y)

    perpdist = DistancePL(Arrowball.x, Arrowball.y, spinx, spiny, cos(Radians(currang-90))+spinx, sin(Radians(currang-90))+spiny)
  end if

  spinflipcw.enabled = 0
  spinflipccw.enabled = 0

  If diffang > 0 then       'cw
    spinflipcw.enabled = 1

    cwstart = -360
    cwend = 360
    currang = getangle(spinflipcw.currentangle)

    if getangle(currang - prevang) < 0 then
      currang = prevang
      spinflipcw.strength = 100
      cwstart = prevang
      cwend = prevang + 10
    elseif getangle(currang - prevang) > 0 then
      if getangle(currang - prevang) > 1.2 and diffang < 90 then
        spinflipcw.strength = 400
      Else
        spinflipcw.strength = 25
      end if
      If (diffang > 90 and balldist > 80) or (diffang <= 90 and perpdist > 80) or balldist > 134  Then
        currang = prevang
        spinflipcw.strength = 100
        cwstart = prevang
        cwend = prevang + 10
      End If
    Else
      spinflipcw.strength = 0
    end If

    ccwstart = currang
    ccwend = currang - 10
    spinflipccw.strength = 100
  Else              'ccw
    spinflipccw.enabled = 1

    ccwstart = 360
    ccwend = -360
    currang = getangle(spinflipccw.currentangle)

    if getangle(currang - prevang) > 0 then
      currang = prevang
      spinflipccw.strength = 100
      ccwstart = prevang
      ccwend = prevang - 10
    elseif getangle(currang - prevang) < 0 then
      if getangle(currang - prevang) < -1.2 and diffang > -90 then
        spinflipccw.strength = 400
      Else
        spinflipccw.strength = 25
      end if
      If (diffang < -90 and balldist > 80) or (diffang >= -90 and perpdist > 80) or balldist > 134  Then
        currang = prevang
        spinflipccw.strength = 100
        ccwstart = prevang
        ccwend = prevang - 10
      End If
    Else
      spinflipccw.strength = 0
    end If

    cwstart = currang
    cwend = currang + 10
    spinflipcw.strength = 100
  End If

  spinflipcw.startangle = cwstart
  spinflipcw.endangle = cwend

  spinflipccw.endangle = ccwstart
  spinflipccw.startangle = ccwend

  ArrowPrimitive.ObjRotz = currang

  ' check if scoring shoould start
  If getangle(currang - prevang) <> 0 and spinflipcw.timerenabled = false Then
    startang = prevang
    spincount = 0
    spinflipcw.timerenabled = True
  End If
end sub

'***********************
' Spinner Switches
'***********************

'7 0 to 36
'4 36 to 72
'8 72 to 108
'3 108 to 144
'10 144 to 180
'9 -144 to -180
'2 -108 to -144
'6 -72 to -108
'1 -36 to -72
'5 0 to -36

spinflipcw.timerinterval = 50

' Scoring for Arrow
Sub spinflipcw_timer
  dim startlight, endlight

  Spincount = Spincount + 1
  If Spincount >= 4 and currang = prevang Then
    me.timerenabled = false

    startlight = getSpinnerNum(startang)
    endlight = getSpinnerNum(currang)

    If startlight <> endlight then
      Select Case Int(Rnd*2)+1
        Case 1 : MultiScore 15, 10
        Case 2 : MultiScore 5, 10
      End Select
    End If

    If Eval("SpinnerLight" & endlight).State = 1 Then
      ScoreSpinnerLight(endlight)
    End If
  End If
End Sub

Dim ScoreLight

Sub ScoreSpinnerLight(LightNum)
  ScoreLight = LightNum
  ScoreSpinner.enabled = True
End Sub

Sub ScoreSpinner_Timer
  If isScoring = False Then
    If SpinnerSpecialLight.state=1 Then
      AddSpecial
    Else
      Multiscore 2, 100
      Eval("SpinnerLight" & ScoreLight).State=0
      Eval("UpperLight" & ScoreLight).State=0
      CheckSpecial
    End If
    me.enabled = false
  End If
End Sub

Function getSpinnerNum(angle)
  If Angle >= 0 and Angle < 36 Then
    getSpinnerNum = 7
  Elseif Angle >= 36 and Angle < 72 Then
    getSpinnerNum = 4
  Elseif Angle >= 72 and Angle < 108 Then
    getSpinnerNum = 8
  Elseif Angle >= 108 and Angle < 144 Then
    getSpinnerNum = 3
  Elseif Angle >= 144 and Angle <= 180 Then
    getSpinnerNum = 10
  Elseif Angle <= 0 and Angle > -36 Then
    getSpinnerNum = 5
  Elseif Angle <= -36 and Angle > -72 Then
    getSpinnerNum = 1
  Elseif Angle <= -72 and Angle > -108 Then
    getSpinnerNum = 6
  Elseif Angle <= -108 and Angle > -144 Then
    getSpinnerNum = 2
  Elseif Angle <= -144 and Angle >= -180 Then
    getSpinnerNum = 9
  End If
End Function

Sub CheckSpecial()
  For Each Obj In NumberLights
    If Obj.State=1 Then Exit Sub
  Next
  TargetSequenceComplete=1
  SpinnerSpecialLight.State=1
  MoveSpinnerLights
  ToggleSpecial
End Sub

'**************************************

'******************************************************
'             SCORING
'******************************************************

Sub AddScore(x,y)
  If TableTilted = 0 and (isScoring = False or y = 1) Then
    PrevScore = Score
    isScoring = True

    If x = 1 Then
      PlayChime(10)
      Score = Score + 1
      AdvMatch
      If score mod 10 = 0 then AdvanceZeroToNine:PlayChime(100)
    ElseIf x = 10 Then
      PlayChime(100)
      Score = Score + 10
      If y = 0 then AdvanceZeroToNine
    ElseIf x = 100 Then
      PlayChime(1000)
      Score = Score + 100
    End If

    If MultiScoreTimer.enabled = False Then
      AddScoreTimer.enabled = True
    End If
  End If

  PlayerScores(0).SetValue(Score)
  PlayerScoresOn(0).SetValue(Score)

  If B2SOn Then
    Controller.B2SSetScorePlayer Player,Score
  End If

  If Score>Replay1 and Replay1Paid=false then
    Replay1Paid=True
    AddSpecial
  End If
  If Score>Replay2 and Replay2Paid=false then
    Replay2Paid=True
    AddSpecial
  End If
  If Score>Replay3 and Replay3Paid=false then
    Replay3Paid=True
    AddSpecial
  End If
  If Score>Replay4 and Replay4Paid=false then
    Replay4Paid=True
    AddSpecial
  End If
End Sub

Sub AddScoreTimer_Timer
  me.enabled = False
  isScoring = False
End Sub

Dim mCountDown, mScore

Sub MultiScore(cycle, Score)
  If isScoring = false Then
    mScore = Score
    mCountDown = cycle - 1
    AddScore mscore, 1
    MultiScoreTimer.enabled = true
  End If
End Sub

Sub MultiScoreTimer_Timer
  mCountDown = mCountDown - 1

  If mCountDown < 1 Then
    me.enabled = False
  End If

  AddScore mscore, 1
  If mcountdown = 10 or mcountdown = 5 Then
    me.interval = 270
  else
    me.interval = 135
  End If
End Sub

Sub PlayChime(x)
  if ChimesOn=0 then
    Select Case x
      Case 10
        If LastChime10=1 Then
          PlaySound SoundFXDOF("SpinACard_1_10_Point_Bell",141,DOFPulse,DOFChimes)
          LastChime10=0
        Else
          PlaySound SoundFXDOF("SpinACard_1_10_Point_Bell",141,DOFPulse,DOFChimes)
          LastChime10=1
        End If
      Case 100
        If LastChime100=1 Then
          PlaySound SoundFXDOF("SpinACard_100_Point_Bell",142,DOFPulse,DOFChimes)
          LastChime100=0
        Else
          PlaySound SoundFXDOF("SpinACard_100_Point_Bell",142,DOFPulse,DOFChimes)
          LastChime100=1
        End If

    End Select
  else
    Select Case x
      Case 10
        If LastChime10=1 Then
          PlaySound SoundFXDOF("SJ_Chime_10a",141,DOFPulse,DOFChimes)
          LastChime10=0
        Else
          PlaySound SoundFXDOF("SJ_Chime_10b",141,DOFPulse,DOFChimes)
          LastChime10=1
        End If
      Case 100
        If LastChime100=1 Then
          PlaySound SoundFXDOF("SJ_Chime_100a",142,DOFPulse,DOFChimes)
          LastChime100=0
        Else
          PlaySound SoundFXDOF("SJ_Chime_100b",142,DOFPulse,DOFChimes)
          LastChime100=1
        End If
      Case 1000
        If LastChime1000=1 Then
          PlaySound SoundFXDOF("SJ_Chime_1000a",143,DOFPulse,DOFChimes)
          LastChime1000=0
        Else
          PlaySound SoundFXDOF("SJ_Chime_1000b",143,DOFPulse,DOFChimes)
          LastChime1000=1
        End If
    End Select
  end if
End Sub


'******************************************************
'             CREDITS
'******************************************************

Sub AddCredit(direction)
  PlaySoundAt "click", KnockerPosition

  credits = credits + direction

  if Credits>15 then Credits=15

  If Credits < 1 Then DOF 109, 0 Else DOF 109, 1

  CreditsReel.SetValue(Credits)
  If B2SOn Then
    Controller.B2SSetCredits Credits
  End If
End Sub

Sub AddSpecial()
  KnockerSolenoid
  DOF 110, 2

  AddCredit 1
End Sub

Sub MoveSpinnerLights
  if SpinnerSpecialLight.state=1 then
    SpinnerLight1.state=0
    SpinnerLight2.state=0
    SpinnerLight3.state=0
    SpinnerLight4.state=0
    if BallsPerGame=3 then
      if ZeroToNineUnit=2 or ZeroToNineUnit=3 or ZeroToNineUnit=7 or ZeroToNineUnit=8 then
        SpinnerLight1.state=1
        SpinnerLight3.state=1
      end if
      if ZeroToNineUnit=0 or ZeroToNineUnit=1 or ZeroToNineUnit=4 or ZeroToNineUnit=5 or ZeroToNineUnit=6 or ZeroToNineUnit=9 then
        SpinnerLight2.state=1
        SpinnerLight4.state=1
      end if
    else
      if ZeroToNineUnit=2 or ZeroToNineUnit=3 then
        SpinnerLight1.state=1
      end if
      if ZeroToNineUnit=4 or ZeroToNineUnit=5 or ZeroToNineUnit=6 then
        SpinnerLight2.state=1
      end if
      if ZeroToNineUnit=7 or ZeroToNineUnit=8 then
        SpinnerLight3.state=1
      end if
      if ZeroToNineUnit=9 or ZeroToNineUnit=0 or ZeroToNineUnit=1 then
        SpinnerLight4.state=1
      end if
    end if
  end if
end sub

sub ToggleSpecial
  SpecialLight9.state=0
  SpecialLight10.state=0

  If TargetSequenceComplete=1 then
    if Bumper1Light.state=1 then
      SpecialLight10.state=1
    else
      SpecialLight9.state=1
    end if
  end if
end sub

'******************************************************
'               MATCH
'******************************************************

Sub Checkmatch
  if TableTilted=true and TiltEndsGame=1 then
    exit sub
  end if

  MatchReel.SetValue(Match+1)

  If B2SOn Then
    If Match = 0 Then
      Controller.B2SSetMatch 10
    Else
      Controller.B2SSetMatch Match
    End If
  End if
  if Match=(Score mod 10) then
    AddSpecial
  end if
End Sub

Sub AdvanceZeroToNine
  ZeroToNineUnit=ZeroToNineUnit+1
  if ZeroToNineUnit>9 then
    ZeroToNineUnit=0
  end if
  MoveSpinnerLights
end sub

Sub AdvMatch()
  Select Case (Match)
    Case 3: Match = 8
    Case 8: Match = 2
    Case 2: Match = 5
    Case 5: Match = 9
    Case 9: Match = 4
    Case 4: Match = 0
    Case 0: Match = 6
    Case 6: Match = 1
    Case 1: Match = 7
    Case 7: Match = 3
  End Select
End Sub

'******************************************************
'               TILT
'******************************************************

Sub TiltTimer_Timer()
  if TiltCount > 0 then TiltCount = TiltCount - 1
  if TiltCount = 0 then
    TiltTimer.Enabled = False
  end if
end sub

Sub TiltIt()
    TiltCount = TiltCount + 1
    if TiltCount = 3 then
      TableTilted=True
      TiltReel.SetValue(1)
      If TiltEndsGame=1 then
        BallInPlay=BallsPerGame
        If B2SOn Then
          Controller.B2SSetGameOver 1
          Controller.B2SSetPlayerUp 0
          Controller.B2SSetBallInPlay 0
          Controller.B2SSetCanPlay 0
        End If

        InstructCard.image="IC"+FormatNumber(BallsPerGame,0)+"0"
      end if

      BumpersOff

      LeftFlipper.RotateToStart
      RightFlipper.RotateToStart
      If B2Son then
        Controller.B2SSetTilt 1
      end if
    else
      TiltTimer.Interval = 500
      TiltTimer.Enabled = True
    end if

end sub

'******************************************************
'           HIGH SCORE
'******************************************************

sub CheckHighScore
  NewHighScore Score, 1
  savehs
end sub

' ============================================================================================
' GNMOD - Multiple High Score Display and Collection
' ============================================================================================
Dim EnteringInitials    ' Normally zero, set to non-zero to enter initials
EnteringInitials = 0

Dim PlungerPulled:PlungerPulled = 0

Dim SelectedChar      ' character under the "cursor" when entering initials

Dim HSTimerCount      ' Pass counter for HS timer, scores are cycled by the timer
HSTimerCount = 5      ' Timer is initially enabled, it'll wrap from 5 to 1 when it's displayed

Dim InitialString     ' the string holding the player's initials as they're entered

Dim AlphaString       ' A-Z, 0-9, space (_) and backspace (<)
Dim AlphaStringPos      ' pointer to AlphaString, move forward and backward with flipper keys
AlphaString = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_<"

Dim HSNewHigh       ' The new score to be recorded

Dim HSScore(5)        ' High Scores read in from config file
Dim HSName(5)       ' High Score Initials read in from config file

' default high scores, remove this when the scores are available from the config file
HSScore(1) = 7500
HSScore(2) = 7000
HSScore(3) = 6000
HSScore(4) = 5500
HSScore(5) = 5000

HSName(1) = "AAA"
HSName(2) = "ZZZ"
HSName(3) = "XXX"
HSName(4) = "ABC"
HSName(5) = "BBB"

Sub HighScoreTimer_Timer
  if EnteringInitials then
    if HSTimerCount = 1 then
      SetHSLine 3, InitialString & MID(AlphaString, AlphaStringPos, 1)
      HSTimerCount = 2
    else
      SetHSLine 3, InitialString
      HSTimerCount = 1
    end if
  elseif InProgress then
    SetHSLine 1, "HIGH SCORE1"
    SetHSLine 2, HSScore(1)
    SetHSLine 3, HSName(1)
    HSTimerCount = 5  ' set so the highest score will show after the game is over
    HighScoreTimer.enabled=false
  else
    ' cycle through high scores
    HighScoreTimer.interval=2000
    HSTimerCount = HSTimerCount + 1
    if HsTimerCount > 5 then
      HSTimerCount = 1
    End If
    SetHSLine 1, "HIGH SCORE"+FormatNumber(HSTimerCount,0)
    SetHSLine 2, HSScore(HSTimerCount)
    SetHSLine 3, HSName(HSTimerCount)
  end if
End Sub

Function GetHSChar(String, Index)
  dim ThisChar
  dim FileName
  ThisChar = Mid(String, Index, 1)
  FileName = "PostIt"
  if ThisChar = " " or ThisChar = "" then
    FileName = FileName & "BL"
  elseif ThisChar = "<" then
    FileName = FileName & "LT"
  elseif ThisChar = "_" then
    FileName = FileName & "SP"
  else
    FileName = FileName & ThisChar
  End If
  GetHSChar = FileName
End Function

Sub SetHsLine(LineNo, String)
  dim Letter
  dim ThisDigit
  dim ThisChar
  dim StrLen
  dim LetterLine
  dim Index
  dim StartHSArray
  dim EndHSArray
  dim LetterName
  dim xfor
  StartHSArray=array(0,1,12,22)
  EndHSArray=array(0,11,21,31)
  StrLen = len(string)
  Index = 1

  for xfor = StartHSArray(LineNo) to EndHSArray(LineNo)
    Eval("HS"&xfor).image = GetHSChar(String, Index)
    Index = Index + 1
  next

End Sub

Sub NewHighScore(NewScore, PlayNum)
  if NewScore > HSScore(5) then
    HighScoreTimer.interval = 500
    HSTimerCount = 1
    AlphaStringPos = 1    ' start with first character "A"
    EnteringInitials = 1  ' intercept the control keys while entering initials
    InitialString = ""    ' initials entered so far, initialize to empty
    SetHSLine 1, "PLAYER "+FormatNumber(PlayNum,0)
    SetHSLine 2, "ENTER NAME"
    SetHSLine 3, MID(AlphaString, AlphaStringPos, 1)
    HSNewHigh = NewScore
    For xx=1 to HighScoreReward
      AddSpecial
    next
  End if
End Sub

Sub CollectInitials(keycode)
  If keycode = LeftFlipperKey Then
    ' back up to previous character
    AlphaStringPos = AlphaStringPos - 1
    if AlphaStringPos < 1 then
      AlphaStringPos = len(AlphaString)   ' handle wrap from beginning to end
      if InitialString = "" then
        ' Skip the backspace if there are no characters to backspace over
        AlphaStringPos = AlphaStringPos - 1
      End if
    end if
    SetHSLine 3, InitialString & MID(AlphaString, AlphaStringPos, 1)
    PlaySound "DropTargetDropped"
  elseif keycode = RightFlipperKey Then
    ' advance to next character
    AlphaStringPos = AlphaStringPos + 1
    if AlphaStringPos > len(AlphaString) or (AlphaStringPos = len(AlphaString) and InitialString = "") then
      ' Skip the backspace if there are no characters to backspace over
      AlphaStringPos = 1
    end if
    SetHSLine 3, InitialString & MID(AlphaString, AlphaStringPos, 1)
    PlaySound "DropTargetDropped"
  elseif keycode = StartGameKey or keycode = PlungerKey Then
    SelectedChar = MID(AlphaString, AlphaStringPos, 1)
    if SelectedChar = "_" then
      InitialString = InitialString & " "
      PlaySound("Ding10")
    elseif SelectedChar = "<" then
      InitialString = MID(InitialString, 1, len(InitialString) - 1)
      if len(InitialString) = 0 then
        ' If there are no more characters to back over, don't leave the < displayed
        AlphaStringPos = 1
      end if
      PlaySound("Ding100")
    else
      InitialString = InitialString & SelectedChar
      PlaySound("Ding10")
    end if
    if len(InitialString) < 3 then
      SetHSLine 3, InitialString & SelectedChar
    End If
  End If
  if len(InitialString) = 3 then
    ' save the score
    for i = 5 to 1 step -1
      if i = 1 or (HSNewHigh > HSScore(i) and HSNewHigh <= HSScore(i - 1)) then
        ' Replace the score at this location
        if i < 5 then
' MsgBox("Moving " & i & " to " & (i + 1))
          HSScore(i + 1) = HSScore(i)
          HSName(i + 1) = HSName(i)
        end if
' MsgBox("Saving initials " & InitialString & " to position " & i)
        EnteringInitials = 0
        HSScore(i) = HSNewHigh
        HSName(i) = InitialString
        HSTimerCount = 5
        HighScoreTimer_Timer
        HighScoreTimer.interval = 2000
        PlaySound("Ding1000")
        exit sub
      elseif i < 5 then
        ' move the score in this slot down by 1, it's been exceeded by the new score
' MsgBox("Moving " & i & " to " & (i + 1))
        HSScore(i + 1) = HSScore(i)
        HSName(i + 1) = HSName(i)
      end if
    next
  End If

End Sub
' END GNMOD

sub savehs
  ' Based on Black's Highscore routines
  Dim FileObj
  Dim ScoreFile
  Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if
  Set ScoreFile=FileObj.CreateTextFile(UserDirectory & HSFileName,True)
    ScoreFile.WriteLine 0
    ScoreFile.WriteLine Credits
    scorefile.writeline BallsPerGame
    scorefile.writeline TiltEndsGame
    scorefile.writeline 0
    scorefile.writeline ReplayLevel
    for xx=1 to 5
      scorefile.writeline HSScore(xx)
    next
    for xx=1 to 5
      scorefile.writeline HSName(xx)
    next

    ScoreFile.Close
  Set ScoreFile=Nothing
  Set FileObj=Nothing
end sub

sub loadhs
    ' Based on Black's Highscore routines
  Dim FileObj
  Dim ScoreFile
    dim temp1,temp2,temp3,temp4,temp5,temp6,temp8,temp9,temp10,temp11,temp12,temp13,temp14,temp15,temp16,temp17

    Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if
  If Not FileObj.FileExists(UserDirectory & HSFileName) then
    Exit Sub
  End if

  Set ScoreFile=FileObj.GetFile(UserDirectory & HSFileName)
  Set TextStr=ScoreFile.OpenAsTextStream(1,0)
    If (TextStr.AtEndOfStream=True) then
      Exit Sub
    End if
    temp1=TextStr.ReadLine
    temp2=textstr.readline
    temp3=textstr.readline
    temp4=textstr.readline
    temp5=textstr.readline
    temp6=textstr.readline

    HighScore=cdbl(temp1)
    if HighScore<1 then
      temp8=textstr.readline
      temp9=textstr.readline
      temp10=textstr.readline
      temp11=textstr.readline
      temp12=textstr.readline
      temp13=textstr.readline
      temp14=textstr.readline
      temp15=textstr.readline
      temp16=textstr.readline
      temp17=textstr.readline
    end if

    TextStr.Close

      Credits=cdbl(temp2)
    BallsPerGame=cdbl(temp3)
    TiltEndsGame=cdbl(temp4)
    ReplayLevel=cdbl(temp6)

    if HighScore<1 then
      HSScore(1) = int(temp8)
      HSScore(2) = int(temp9)
      HSScore(3) = int(temp10)
      HSScore(4) = int(temp11)
      HSScore(5) = int(temp12)

      HSName(1) = temp13
      HSName(2) = temp14
      HSName(3) = temp15
      HSName(4) = temp16
      HSName(5) = temp17
    end if

    Set ScoreFile=Nothing
      Set FileObj=Nothing
end sub

Sub SetupReplayTables
  Replay1Table(1)=3400
  Replay1Table(2)=3600
  Replay1Table(3)=3800
  Replay1Table(4)=3900
  Replay1Table(5)=4100
  Replay1Table(6)=4400
  Replay1Table(7)=4600
  Replay1Table(8)=5200
  Replay1Table(9)=5300
  Replay1Table(10)=5500
  Replay1Table(11)=5800
  Replay1Table(12)=6000
  Replay1Table(13)=6300
  Replay1Table(14)=6600
  Replay1Table(15)=999000

  Replay2Table(1)=4000
  Replay2Table(2)=4200
  Replay2Table(3)=4400
  Replay2Table(4)=4500
  Replay2Table(5)=4700
  Replay2Table(6)=5000
  Replay2Table(7)=5200
  Replay2Table(8)=5800
  Replay2Table(9)=5900
  Replay2Table(10)=6100
  Replay2Table(11)=6400
  Replay2Table(12)=6600
  Replay2Table(13)=6900
  Replay2Table(14)=7200
  Replay2Table(15)=999000

  Replay3Table(1)=4600
  Replay3Table(2)=4800
  Replay3Table(3)=5000
  Replay3Table(4)=5100
  Replay3Table(5)=5300
  Replay3Table(6)=5600
  Replay3Table(7)=5800
  Replay3Table(8)=6400
  Replay3Table(9)=6500
  Replay3Table(10)=6700
  Replay3Table(11)=7000
  Replay3Table(12)=7200
  Replay3Table(13)=7500
  Replay3Table(14)=7800
  Replay3Table(15)=999000

  Replay4Table(1)=5300
  Replay4Table(2)=5400
  Replay4Table(3)=5600
  Replay4Table(4)=5700
  Replay4Table(5)=5900
  Replay4Table(6)=6200
  Replay4Table(7)=6400
  Replay4Table(8)=7000
  Replay4Table(9)=7100
  Replay4Table(10)=7300
  Replay4Table(11)=7600
  Replay4Table(12)=7800
  Replay4Table(13)=8100
  Replay4Table(14)=8400
  Replay4Table(15)=999000

  ReplayTableMax=14
end sub

Sub RefreshReplayCard
  Dim tempst1
  Dim tempst2

  tempst1=FormatNumber(BallsPerGame,0)
  tempst2=FormatNumber(ReplayLevel,0)


  ReplayCard.image = "SC" + tempst2
  Replay1=Replay1Table(ReplayLevel)
  Replay2=Replay2Table(ReplayLevel)
  Replay3=Replay3Table(ReplayLevel)
  Replay4=Replay4Table(ReplayLevel)

end sub

' ============================================================================================
' GNMOD - New Options menu
' ============================================================================================
Dim EnteringOptions
Dim CurrentOption
Dim OptionCHS
Dim MaxOption
Dim OptionHighScorePosition
Dim XOpt
Dim StartingArray
Dim EndingArray

StartingArray=Array(0,1,2,30,33,61,89,117,145,173,201,229)
EndingArray=Array(0,1,29,32,60,88,116,144,172,200,228,256)
EnteringOptions = 0
MaxOption = 9
OptionCHS = 0
OptionHighScorePosition = 0
Const OptionLinesToMark="111100011"
Const OptionLine1="" 'do not use this line
Const OptionLine2="" 'do not use this line
Const OptionLine3="" 'do not use this line
Const OptionLine4="Wedgehead Tilt Setting"
Const OptionLine5=""
Const OptionLine6=""
Const OptionLine7=""
Const OptionLine8="" 'do not use this line
Const OptionLine9="" 'do not use this line

Sub OperatorMenuTimer_Timer
  EnteringOptions = 1
  OperatorMenuTimer.enabled=false
  ShowOperatorMenu
end sub

sub ShowOperatorMenu
  OperatorMenuBackdrop.image = "OperatorMenu"

  OptionCHS = 0
  CurrentOption = 1
  DisplayAllOptions
  OperatorOption1.image = "BluePlus"
  SetHighScoreOption

End Sub

Sub DisplayAllOptions
  dim linecounter
  dim tempstring
  For linecounter = 1 to MaxOption
    tempstring=Eval("OptionLine"&linecounter)
    Select Case linecounter
      Case 1:
        tempstring=tempstring + FormatNumber(BallsPerGame,0)
        SetOptLine 1,tempstring
      Case 2:
        if Replay3Table(ReplayLevel)=999000 then
          tempstring = FormatNumber(Replay1Table(ReplayLevel),0) + "/" + FormatNumber(Replay2Table(ReplayLevel),0)
        elseif Replay4Table(ReplayLevel)=999000 then
          tempstring = FormatNumber(Replay1Table(ReplayLevel),0) + "/" + FormatNumber(Replay2Table(ReplayLevel),0) + "/" + FormatNumber(Replay3Table(ReplayLevel),0)
        else
          tempstring = FormatNumber(Replay1Table(ReplayLevel),0) + "/" + FormatNumber(Replay2Table(ReplayLevel),0) + "/" + FormatNumber(Replay3Table(ReplayLevel),0) + "/" + FormatNumber(Replay4Table(ReplayLevel),0)
        end if
      SetOptLine 2,tempstring
      Case 3:
        If OptionCHS=0 then
          tempstring = "NO"
        else
          tempstring = "YES"
        end if
        SetOptLine 3,tempstring
      Case 4:
        SetOptLine 4, tempstring
        If TiltEndsGame=0 then
          tempstring="Tilt loses ball in play"
        else
          tempstring="Tilt ends game"
        end if

        SetOptLine 5, tempstring
      Case 5:
        SetOptLine 6, tempstring

        SetOptLine 7, tempstring

      Case 6:
        SetOptLine 8, tempstring

        SetOptLine 9, tempstring

      Case 7:
        SetOptLine 10, tempstring
        SetOptLine 11, tempstring

      Case 8:

      Case 9:


    End Select

  next
end sub

sub MoveArrow
  do
    CurrentOption = CurrentOption + 1
    If CurrentOption>Len(OptionLinesToMark) then
      CurrentOption=1
    end if
  loop until Mid(OptionLinesToMark,CurrentOption,1)="1"
end sub

sub CollectOptions(ByVal keycode)
  if Keycode = LeftFlipperKey then
    PlaySound "DropTargetDropped"
    For XOpt = 1 to MaxOption
      Eval("OperatorOption"&XOpt).image = "PostitBL"
    next
    MoveArrow
    if CurrentOption<8 then
      Eval("OperatorOption"&CurrentOption).image = "BluePlus"
    elseif CurrentOption=8 then
      Eval("OperatorOption"&CurrentOption).image = "GreenCheck"
    else
      Eval("OperatorOption"&CurrentOption).image = "RedX"
    end if

  elseif Keycode = RightFlipperKey then
    PlaySound "DropTargetDropped"
    if CurrentOption = 1 then
      If BallsPerGame = 3 then
        BallsPerGame = 5
      else
        BallsPerGame = 3
      end if
      DisplayAllOptions
    elseif CurrentOption = 2 then
      ReplayLevel=ReplayLevel+1
      If ReplayLevel>ReplayTableMax then
        ReplayLevel=1
      end if
      DisplayAllOptions
    elseif CurrentOption = 3 then
      if OptionCHS = 0 then
        OptionCHS = 1

      else
        OptionCHS = 0

      end if
      DisplayAllOptions
    elseif CurrentOption = 4 then
      if TiltEndsGame=0 then
        TiltEndsGame=1
      else
        TiltEndsGame=0
      end if
      DisplayAllOptions


    elseif CurrentOption = 8 or CurrentOption = 9 then
        if OptionCHS=1 then
          HSScore(1) = 7500
          HSScore(2) = 7000
          HSScore(3) = 6000
          HSScore(4) = 5500
          HSScore(5) = 5000

          HSName(1) = "AAA"
          HSName(2) = "ZZZ"
          HSName(3) = "XXX"
          HSName(4) = "ABC"
          HSName(5) = "BBB"
        end if

        if CurrentOption = 8 then
          savehs
        else
          loadhs
        end if
        OperatorMenuBackdrop.image = "PostitBL"
        For XOpt = 1 to MaxOption
          Eval("OperatorOption"&XOpt).image = "PostitBL"
        next

        For XOpt = 1 to 256
          Eval("Option"&XOpt).image = "PostItBL"
        next
        RefreshReplayCard
        InstructCard.image="IC"+FormatNumber(BallsPerGame,0)+"0"
        EnteringOptions = 0

    end if
  end if
End Sub

Sub SetHighScoreOption

End Sub

Function GetOptChar(String, Index)
  dim ThisChar
  dim FileName
  ThisChar = Mid(String, Index, 1)
  FileName = "PostIt"
  if ThisChar = " " or ThisChar = "" then
    FileName = FileName & "BL"
  elseif ThisChar = "<" then
    FileName = FileName & "LT"
  elseif ThisChar = "_" then
    FileName = FileName & "SP"
  elseif ThisChar = "/" then
    FileName = FileName & "SL"
  elseif ThisChar = "," then
    FileName = FileName & "CM"
  else
    FileName = FileName & ThisChar
  End If
  GetOptChar = FileName
End Function

dim LineLengths(22) ' maximum number of lines
Sub SetOptLine(LineNo, String)
  Dim DispLen
    Dim StrLen
  dim xfor
  dim Letter
  dim ThisDigit
  dim ThisChar
  dim LetterLine
  dim Index
  dim LetterName
  StrLen = len(string)
  Index = 1

  StrLen = len(String)
    DispLen = StrLen
    if (DispLen < LineLengths(LineNo)) Then
        DispLen = LineLengths(LineNo)
    end If

  for xfor = StartingArray(LineNo) to StartingArray(LineNo) + DispLen
    Eval("Option"&xfor).image = GetOptChar(string, Index)
    Index = Index + 1
  next
  LineLengths(LineNo) = StrLen

End Sub

'************************************************************************
'                         Ball Control - 3 Axis
'************************************************************************

'******* for ball control script

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

'================================
'Helper Functions


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

Sub ShuffleArrays(aArray1, aArray2, offset)
  ShuffleArray aArray1, offset
  ShuffleArray aArray2, offset
End Sub


Function BallSpeed(ball) 'Calculates the ball speed
    BallSpeed = SQR(ball.VelX^2 + ball.VelY^2 + ball.VelZ^2)
End Function

Function PSlope(Input, X1, Y1, X2, Y2)  'Set up line via two points, no clamping. Input X, output Y
  dim x, y, b, m : x = input : m = (Y2 - Y1) / (X2 - X1) : b = Y2 - m*X2
  Y = M*x+b
  PSlope = Y
End Function

Function NullFunctionZ(aEnabled):End Function '1 argument null function placeholder  TODO move me or replac eme

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


Function LinearEnvelope(xInput, xKeyFrame, yLvl)
  dim y 'Y output
  dim L 'Line
  dim ii : for ii = 1 to uBound(xKeyFrame)  'find active line
    if xInput <= xKeyFrame(ii) then L = ii : exit for : end if
  Next
  if xInput > xKeyFrame(uBound(xKeyFrame) ) then L = uBound(xKeyFrame)  'catch line overrun
  Y = pSlope(xInput, xKeyFrame(L-1), yLvl(L-1), xKeyFrame(L), yLvl(L) )

  'Clamp if on the boundry lines
  'if L=1 and Y < yLvl(LBound(yLvl) ) then Y = yLvl(lBound(yLvl) )
  'if L=uBound(xKeyFrame) and Y > yLvl(uBound(yLvl) ) then Y = yLvl(uBound(yLvl) )
  'clamp 2.0
  if xInput <= xKeyFrame(lBound(xKeyFrame) ) then Y = yLvl(lBound(xKeyFrame) )  'Clamp lower
  if xInput >= xKeyFrame(uBound(xKeyFrame) ) then Y = yLvl(uBound(xKeyFrame) )  'Clamp upper

  LinearEnvelope = Y
End Function

Sub RDampen_Timer()
  Cor.Update
End Sub

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

' Thalamus - patched :     aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
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

'Tracks ball velocity for judging bounce calculations & angle
'apologies to JimmyFingers is this is what his script does. I know his tracks ball velocity too but idk how it works in particular
dim cor : set cor = New CoRTracker
cor.debugOn = False
'cor.update() - put this on a low interval timer
Class CoRTracker
  public DebugOn 'tbpIn.text
  public ballvel

  Private Sub Class_Initialize : redim ballvel(0) : End Sub
  'TODO this would be better if it didn't do the sorting every ms, but instead every time it's pulled for COR stuff
  Public Sub Update() 'tracks in-ball-velocity
    dim str, b, AllBalls, highestID : allBalls = getballs
    if uBound(allballs) < 0 then if DebugOn then str = "no balls" : TBPin.text = str : exit Sub else exit sub end if: end if
    for each b in allballs
      if b.id >= HighestID then highestID = b.id
    Next

    if uBound(ballvel) < highestID then redim ballvel(highestID)  'set bounds

    for each b in allballs
      ballvel(b.id) = BallSpeed(b)
      if DebugOn then
        dim s, bs 'debug spacer, ballspeed
        bs = round(BallSpeed(b),1)
        if bs < 10 then s = " " else s = "" end if
        str = str & b.id & ": " & s & bs & vbnewline
        'str = str & b.id & ": " & s & bs & "z:" & b.z & vbnewline
      end if
    Next
    if DebugOn then str = "ubound ballvels: " & ubound(ballvel) & vbnewline & str : if TBPin.text <> str then TBPin.text = str : end if
  End Sub
End Class


'////////////////////////////  MECHANICAL SOUNDS  ///////////////////////////
'//  This part in the script is an entire block that is dedicated to the physics sound system.
'//  Various scripts and sounds that may be pretty generic and could suit other WPC systems, but the most are tailored specifically for this table.

'///////////////////////////////  SOUNDS PARAMETERS  //////////////////////////////
Dim GlobalSoundLevel, CoinSoundLevel, PlungerReleaseSoundLevel, PlungerPullSoundLevel, NudgeLeftSoundLevel
Dim NudgeRightSoundLevel, NudgeCenterSoundLevel, StartButtonSoundLevel, RollingSoundFactor

CoinSoundLevel = 1                            'volume level; range [0, 1]
NudgeLeftSoundLevel = 1                         'volume level; range [0, 1]
NudgeRightSoundLevel = 1                        'volume level; range [0, 1]
NudgeCenterSoundLevel = 1                       'volume level; range [0, 1]
StartButtonSoundLevel = 0.1                       'volume level; range [0, 1]
PlungerReleaseSoundLevel = 0.8 '1 wjr                     'volume level; range [0, 1]
PlungerPullSoundLevel = 1                       'volume level; range [0, 1]
RollingSoundFactor = 1.1/5

'///////////////////////-----Solenoids, Kickers and Flash Relays-----///////////////////////
Dim FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel, FlipperUpAttackLeftSoundLevel, FlipperUpAttackRightSoundLevel
Dim FlipperUpSoundLevel, FlipperDownSoundLevel, FlipperLeftHitParm, FlipperRightHitParm
Dim SlingshotSoundLevel, BumperSoundFactor, KnockerSoundLevel

FlipperUpAttackMinimumSoundLevel = 0.010                      'volume level; range [0, 1]
FlipperUpAttackMaximumSoundLevel = 0.635                'volume level; range [0, 1]
FlipperUpSoundLevel = 1.0                                   'volume level; range [0, 1]
FlipperDownSoundLevel = 0.45                                  'volume level; range [0, 1]
FlipperLeftHitParm = FlipperUpSoundLevel                'sound helper; not configurable
FlipperRightHitParm = FlipperUpSoundLevel               'sound helper; not configurable
SlingshotSoundLevel = 0.95                        'volume level; range [0, 1]
BumperSoundFactor = 4.25                        'volume multiplier; must not be zero
KnockerSoundLevel = 1                           'volume level; range [0, 1]

'///////////////////////-----Ball Drops, Bumps and Collisions-----///////////////////////
Dim RubberStrongSoundFactor, RubberWeakSoundFactor, RubberFlipperSoundFactor
Dim BallBouncePlayfieldSoftFactor, BallBouncePlayfieldHardFactor, PlasticRampDropToPlayfieldSoundLevel, WireRampDropToPlayfieldSoundLevel, DelayedBallDropOnPlayfieldSoundLevel
Dim WallImpactSoundFactor, MetalImpactSoundFactor, SubwaySoundLevel, SubwayEntrySoundLevel, ScoopEntrySoundLevel
Dim SaucerLockSoundLevel, SaucerKickSoundLevel

RubberStrongSoundFactor = 0.055/5                     'volume multiplier; must not be zero
RubberWeakSoundFactor = 0.075/5                     'volume multiplier; must not be zero
RubberFlipperSoundFactor = 0.075/5                    'volume multiplier; must not be zero
BallBouncePlayfieldSoftFactor = 0.025                 'volume multiplier; must not be zero
BallBouncePlayfieldHardFactor = 0.025                 'volume multiplier; must not be zero
DelayedBallDropOnPlayfieldSoundLevel = 0.8                  'volume level; range [0, 1]
WallImpactSoundFactor = 0.075                     'volume multiplier; must not be zero
MetalImpactSoundFactor = 0.075/3
SaucerLockSoundLevel = 0.8
SaucerKickSoundLevel = 0.8

'///////////////////////-----Gates, Spinners, Rollovers and Targets-----///////////////////////

Dim GateSoundLevel, TargetSoundFactor, SpinnerSoundLevel, RolloverSoundLevel, DTSoundLevel

GateSoundLevel = 0.5/5                          'volume level; range [0, 1]
TargetSoundFactor = 0.0025 * 10                     'volume multiplier; must not be zero
DTSoundLevel = 0.25                           'volume multiplier; must not be zero
RolloverSoundLevel = 0.25                                       'volume level; range [0, 1]

'///////////////////////-----Ball Release, Guides and Drain-----///////////////////////
Dim DrainSoundLevel, BallReleaseSoundLevel, BottomArchBallGuideSoundFactor, FlipperBallGuideSoundFactor

DrainSoundLevel = 0.8                           'volume level; range [0, 1]
BallReleaseSoundLevel = 1                       'volume level; range [0, 1]
BottomArchBallGuideSoundFactor = 0.2                  'volume multiplier; must not be zero
FlipperBallGuideSoundFactor = 0.015                   'volume multiplier; must not be zero

'///////////////////////-----Loops and Lanes-----///////////////////////
Dim ArchSoundFactor
ArchSoundFactor = 0.025/5                         'volume multiplier; must not be zero


'/////////////////////////////  SOUND PLAYBACK FUNCTIONS  ////////////////////////////
'/////////////////////////////  POSITIONAL SOUND PLAYBACK METHODS  ////////////////////////////
' Positional sound playback methods will play a sound, depending on the X,Y position of the table element or depending on ActiveBall object position
' These are similar subroutines that are less complicated to use (e.g. simply use standard parameters for the PlaySound call)
' For surround setup - positional sound playback functions will fade between front and rear surround channels and pan between left and right channels
' For stereo setup - positional sound playback functions will only pan between left and right channels
' For mono setup - positional sound playback functions will not pan between left and right channels and will not fade between front and rear channels

' PlaySound full syntax - PlaySound(string, int loopcount, float volume, float pan, float randompitch, int pitch, bool useexisting, bool restart, float front_rear_fade)
' Note - These functions will not work (currently) for walls/slingshots as these do not feature a simple, single X,Y position
Sub PlaySoundAtLevelStatic(playsoundparams, aVol, tableobj)
    PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelExistingStatic(playsoundparams, aVol, tableobj)
    PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticLoop(playsoundparams, aVol, tableobj)
    PlaySound playsoundparams, -1, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticRandomPitch(playsoundparams, aVol, randomPitch, tableobj)
    PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelActiveBall(playsoundparams, aVol)
  PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ActiveBall), 0, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLevelExistingActiveBall(playsoundparams, aVol)
  PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ActiveBall), 0, 0, 1, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLeveTimerActiveBall(playsoundparams, aVol, ballvariable)
  PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ballvariable), 0, 0, 0, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelTimerExistingActiveBall(playsoundparams, aVol, ballvariable)
  PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ballvariable), 0, 0, 1, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelRoll(playsoundparams, aVol, pitch)
    PlaySound playsoundparams, -1, aVol * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

' Previous Positional Sound Subs

Sub PlaySoundAt(soundname, tableobj)
    PlaySound soundname, 1, 1 * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtVol(soundname, tableobj, aVol)
    PlaySound soundname, 1, aVol * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname)
    PlaySoundAt soundname, ActiveBall
End Sub

Sub PlaySoundAtBallVol (Soundname, aVol)
  Playsound soundname, 1,aVol * VolumeDial, AudioPan(ActiveBall), 0,0,0, 1, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtBallVolM (Soundname, aVol)
  Playsound soundname, 1,aVol * VolumeDial, AudioPan(ActiveBall), 0,0,0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtVolLoops(sound, tableobj, Vol, Loops)
  PlaySound sound, Loops, Vol * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub


' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
    tmp = tableobj.y * 2 / tableheight-1
    If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = tableobj.x * 2 / tablewidth-1
    If tmp > 0 Then
        AudioPan = Csng(tmp ^10)
    Else
        AudioPan = Csng(-((- tmp) ^10) )
    End If
End Function

Function Vol(ball) ' Calculates the volume of the sound based on the ball speed
  Vol = Csng(BallVel(ball) ^2)
End Function

Function Volz(ball) ' Calculates the volume of the sound based on the ball speed
  Volz = Csng((ball.velz) ^2)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

Function VolPlayfieldRoll(ball) ' Calculates the roll volume of the sound based on the ball speed
  VolPlayfieldRoll = RollingSoundFactor * 0.0005 * Csng(BallVel(ball) ^3)
End Function

Function PitchPlayfieldRoll(ball) ' Calculates the roll pitch of the sound based on the ball speed
    PitchPlayfieldRoll = BallVel(ball) ^2 * 15
End Function

Function RndInt(min, max)
    RndInt = Int(Rnd() * (max-min + 1) + min)' Sets a random number integer between min and max
End Function

Function RndNum(min, max)
    RndNum = Rnd() * (max-min) + min' Sets a random number between min and max
End Function

'/////////////////////////////  GENERAL SOUND SUBROUTINES  ////////////////////////////
Sub SoundStartButton()
  PlaySoundAtLevelStatic ("Start_Button"), StartButtonSoundLevel, Drain
End Sub

Sub SoundNudgeLeft()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound ("Nudge_1"), 0, NudgeLeftSoundLevel * VolumeDial, -0.1, 0.25
    Case 2 : PlaySound ("Nudge_2"), 0, NudgeLeftSoundLevel * VolumeDial, -0.1, 0.25
    Case 3 : PlaySound ("Nudge_3"), 0, NudgeLeftSoundLevel * VolumeDial, -0.1, 0.25
  End Select
End Sub

Sub SoundNudgeRight()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound ("Nudge_1"), 0, NudgeRightSoundLevel * VolumeDial, 0.1, 0.25
    Case 2 : PlaySound ("Nudge_2"), 0, NudgeRightSoundLevel * VolumeDial, 0.1, 0.25
    Case 3 : PlaySound ("Nudge_3"), 0, NudgeRightSoundLevel * VolumeDial, 0.1, 0.25
  End Select
End Sub

Sub SoundNudgeCenter()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySoundAtLevelStatic ("Nudge_1"), NudgeCenterSoundLevel * VolumeDial, Drain
    Case 2 : PlaySoundAtLevelStatic ("Nudge_2"), NudgeCenterSoundLevel * VolumeDial, Drain
    Case 3 : PlaySoundAtLevelStatic ("Nudge_3"), NudgeCenterSoundLevel * VolumeDial, Drain
  End Select
End Sub


Sub SoundPlungerPull()
  PlaySoundAtLevelStatic ("Plunger_Pull_1"), PlungerPullSoundLevel, Plunger
End Sub

Sub SoundPlungerReleaseBall()
  PlaySoundAtLevelStatic ("Plunger_Release_Ball"), PlungerReleaseSoundLevel, Plunger
End Sub

Sub SoundPlungerReleaseNoBall()
  PlaySoundAtLevelStatic ("Plunger_Release_No_Ball"), PlungerReleaseSoundLevel, Plunger
End Sub


'/////////////////////////////  KNOCKER SOLENOID  ////////////////////////////
Sub KnockerSolenoid()
  PlaySoundAtLevelStatic SoundFX("Knocker_1",DOFKnocker), KnockerSoundLevel, KnockerPosition
End Sub

'/////////////////////////////  DRAIN SOUNDS  ////////////////////////////
Sub RandomSoundDrain(drainswitch)
  Select Case Int(Rnd*11)+1
    Case 1 : PlaySoundAtLevelStatic ("Drain_1"), DrainSoundLevel, drainswitch
    Case 2 : PlaySoundAtLevelStatic ("Drain_2"), DrainSoundLevel, drainswitch
    Case 3 : PlaySoundAtLevelStatic ("Drain_3"), DrainSoundLevel, drainswitch
    Case 4 : PlaySoundAtLevelStatic ("Drain_4"), DrainSoundLevel, drainswitch
    Case 5 : PlaySoundAtLevelStatic ("Drain_5"), DrainSoundLevel, drainswitch
    Case 6 : PlaySoundAtLevelStatic ("Drain_6"), DrainSoundLevel, drainswitch
    Case 7 : PlaySoundAtLevelStatic ("Drain_7"), DrainSoundLevel, drainswitch
    Case 8 : PlaySoundAtLevelStatic ("Drain_8"), DrainSoundLevel, drainswitch
    Case 9 : PlaySoundAtLevelStatic ("Drain_9"), DrainSoundLevel, drainswitch
    Case 10 : PlaySoundAtLevelStatic ("Drain_10"), DrainSoundLevel, drainswitch
    Case 11 : PlaySoundAtLevelStatic ("Drain_11"), DrainSoundLevel, drainswitch
  End Select
End Sub


'/////////////////////////////  SLINGSHOT SOLENOID SOUNDS  ////////////////////////////
Sub RandomSoundSlingshotLeft()
  Select Case Int(Rnd*10)+1
    Case 1 : PlaySoundAtLevelStatic SoundFX("Sling_L1",DOFContactors), SlingshotSoundLevel, Sling2
    Case 2 : PlaySoundAtLevelStatic SoundFX("Sling_L2",DOFContactors), SlingshotSoundLevel, Sling2
    Case 3 : PlaySoundAtLevelStatic SoundFX("Sling_L3",DOFContactors), SlingshotSoundLevel, Sling2
    Case 4 : PlaySoundAtLevelStatic SoundFX("Sling_L4",DOFContactors), SlingshotSoundLevel, Sling2
    Case 5 : PlaySoundAtLevelStatic SoundFX("Sling_L5",DOFContactors), SlingshotSoundLevel, Sling2
    Case 6 : PlaySoundAtLevelStatic SoundFX("Sling_L6",DOFContactors), SlingshotSoundLevel, Sling2
    Case 7 : PlaySoundAtLevelStatic SoundFX("Sling_L7",DOFContactors), SlingshotSoundLevel, Sling2
    Case 8 : PlaySoundAtLevelStatic SoundFX("Sling_L8",DOFContactors), SlingshotSoundLevel, Sling2
    Case 9 : PlaySoundAtLevelStatic SoundFX("Sling_L9",DOFContactors), SlingshotSoundLevel, Sling2
    Case 10 : PlaySoundAtLevelStatic SoundFX("Sling_L10",DOFContactors), SlingshotSoundLevel, Sling2
  End Select
End Sub

Sub RandomSoundSlingshotRight()
  Select Case Int(Rnd*8)+1
    Case 1 : PlaySoundAtLevelStatic SoundFX("Sling_R1",DOFContactors), SlingshotSoundLevel, Sling1
    Case 2 : PlaySoundAtLevelStatic SoundFX("Sling_R2",DOFContactors), SlingshotSoundLevel, Sling1
    Case 3 : PlaySoundAtLevelStatic SoundFX("Sling_R3",DOFContactors), SlingshotSoundLevel, Sling1
    Case 4 : PlaySoundAtLevelStatic SoundFX("Sling_R4",DOFContactors), SlingshotSoundLevel, Sling1
    Case 5 : PlaySoundAtLevelStatic SoundFX("Sling_R5",DOFContactors), SlingshotSoundLevel, Sling1
    Case 6 : PlaySoundAtLevelStatic SoundFX("Sling_R6",DOFContactors), SlingshotSoundLevel, Sling1
    Case 7 : PlaySoundAtLevelStatic SoundFX("Sling_R7",DOFContactors), SlingshotSoundLevel, Sling1
    Case 8 : PlaySoundAtLevelStatic SoundFX("Sling_R8",DOFContactors), SlingshotSoundLevel, Sling1
  End Select
End Sub

'/////////////////////////////  BUMPER SOLENOID SOUNDS  ////////////////////////////
Sub RandomSoundBumper(bumper)
  Select Case Int(Rnd*5)+1
    Case 1 : PlaySoundAtLevelStatic SoundFX("Bumpers_Top_1",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bumper
    Case 2 : PlaySoundAtLevelStatic SoundFX("Bumpers_Top_2",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bumper
    Case 3 : PlaySoundAtLevelStatic SoundFX("Bumpers_Top_3",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bumper
    Case 4 : PlaySoundAtLevelStatic SoundFX("Bumpers_Top_4",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bumper
    Case 5 : PlaySoundAtLevelStatic SoundFX("Bumpers_Top_5",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bumper
  End Select
End Sub


'/////////////////////////////  FLIPPER BATS SOUND SUBROUTINES  ////////////////////////////
'/////////////////////////////  FLIPPER BATS SOLENOID ATTACK SOUND  ////////////////////////////
Sub SoundFlipperUpAttackLeft(flipper)
  FlipperUpAttackLeftSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
  PlaySoundAtLevelStatic ("Flipper_Attack-L01"), FlipperUpAttackLeftSoundLevel, flipper
End Sub

Sub SoundFlipperUpAttackRight(flipper)
  FlipperUpAttackRightSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
    PlaySoundAtLevelStatic ("Flipper_Attack-R01"), FlipperUpAttackLeftSoundLevel, flipper
End Sub

'/////////////////////////////  FLIPPER BATS SOLENOID CORE SOUND  ////////////////////////////
Sub RandomSoundFlipperUpLeft(flipper)
  Select Case Int(Rnd*11)+1
    Case 1 : PlaySoundAtLevelStatic SoundFX("Flipper_L01",DOFFlippers), FlipperLeftHitParm, Flipper
    Case 2 : PlaySoundAtLevelStatic SoundFX("Flipper_L02",DOFFlippers), FlipperLeftHitParm, Flipper
    Case 3 : PlaySoundAtLevelStatic SoundFX("Flipper_L07",DOFFlippers), FlipperLeftHitParm, Flipper
    Case 4 : PlaySoundAtLevelStatic SoundFX("Flipper_L08",DOFFlippers), FlipperLeftHitParm, Flipper
    Case 5 : PlaySoundAtLevelStatic SoundFX("Flipper_L09",DOFFlippers), FlipperLeftHitParm, Flipper
    Case 6 : PlaySoundAtLevelStatic SoundFX("Flipper_L10",DOFFlippers), FlipperLeftHitParm, Flipper
    Case 7 : PlaySoundAtLevelStatic SoundFX("Flipper_L12",DOFFlippers), FlipperLeftHitParm, Flipper
    Case 8 : PlaySoundAtLevelStatic SoundFX("Flipper_L14",DOFFlippers), FlipperLeftHitParm, Flipper
    Case 9 : PlaySoundAtLevelStatic SoundFX("Flipper_L18",DOFFlippers), FlipperLeftHitParm, Flipper
    Case 10 : PlaySoundAtLevelStatic SoundFX("Flipper_L20",DOFFlippers), FlipperLeftHitParm, Flipper
    Case 11 : PlaySoundAtLevelStatic SoundFX("Flipper_L26",DOFFlippers), FlipperLeftHitParm, Flipper
  End Select
End Sub

Sub RandomSoundFlipperUpRight(flipper)
  Select Case Int(Rnd*11)+1
    Case 1 : PlaySoundAtLevelStatic SoundFX("Flipper_R01",DOFFlippers), FlipperRightHitParm, Flipper
    Case 2 : PlaySoundAtLevelStatic SoundFX("Flipper_R02",DOFFlippers), FlipperRightHitParm, Flipper
    Case 3 : PlaySoundAtLevelStatic SoundFX("Flipper_R03",DOFFlippers), FlipperRightHitParm, Flipper
    Case 4 : PlaySoundAtLevelStatic SoundFX("Flipper_R04",DOFFlippers), FlipperRightHitParm, Flipper
    Case 5 : PlaySoundAtLevelStatic SoundFX("Flipper_R05",DOFFlippers), FlipperRightHitParm, Flipper
    Case 6 : PlaySoundAtLevelStatic SoundFX("Flipper_R06",DOFFlippers), FlipperRightHitParm, Flipper
    Case 7 : PlaySoundAtLevelStatic SoundFX("Flipper_R07",DOFFlippers), FlipperRightHitParm, Flipper
    Case 8 : PlaySoundAtLevelStatic SoundFX("Flipper_R08",DOFFlippers), FlipperRightHitParm, Flipper
    Case 9 : PlaySoundAtLevelStatic SoundFX("Flipper_R09",DOFFlippers), FlipperRightHitParm, Flipper
    Case 10 : PlaySoundAtLevelStatic SoundFX("Flipper_R10",DOFFlippers), FlipperRightHitParm, Flipper
    Case 11 : PlaySoundAtLevelStatic SoundFX("Flipper_R11",DOFFlippers), FlipperRightHitParm, Flipper
  End Select
End Sub

Sub RandomSoundReflipUpLeft(flipper)
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_L01",DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, Flipper
    Case 2 : PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_L02",DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, Flipper
    Case 3 : PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_L03",DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, Flipper
  End Select
End Sub

Sub RandomSoundReflipUpRight(flipper)
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_R01",DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, Flipper
    Case 2 : PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_R02",DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, Flipper
    Case 3 : PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_R03",DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, Flipper
  End Select
End Sub

Sub RandomSoundFlipperDownLeft(flipper)
  Select Case Int(Rnd*7)+1
    Case 1 : PlaySoundAtLevelStatic SoundFX("Flipper_Left_Down_1",DOFFlippers), FlipperDownSoundLevel, Flipper
    Case 2 : PlaySoundAtLevelStatic SoundFX("Flipper_Left_Down_2",DOFFlippers), FlipperDownSoundLevel, Flipper
    Case 3 : PlaySoundAtLevelStatic SoundFX("Flipper_Left_Down_3",DOFFlippers), FlipperDownSoundLevel, Flipper
    Case 4 : PlaySoundAtLevelStatic SoundFX("Flipper_Left_Down_4",DOFFlippers), FlipperDownSoundLevel, Flipper
    Case 5 : PlaySoundAtLevelStatic SoundFX("Flipper_Left_Down_5",DOFFlippers), FlipperDownSoundLevel, Flipper
    Case 6 : PlaySoundAtLevelStatic SoundFX("Flipper_Left_Down_6",DOFFlippers), FlipperDownSoundLevel, Flipper
    Case 7 : PlaySoundAtLevelStatic SoundFX("Flipper_Left_Down_7",DOFFlippers), FlipperDownSoundLevel, Flipper
  End Select
End Sub

Sub RandomSoundFlipperDownRight(flipper)
  Select Case Int(Rnd*8)+1
    Case 1 : PlaySoundAtLevelStatic SoundFX("Flipper_Right_Down_1",DOFFlippers), FlipperDownSoundLevel, Flipper
    Case 2 : PlaySoundAtLevelStatic SoundFX("Flipper_Right_Down_2",DOFFlippers), FlipperDownSoundLevel, Flipper
    Case 3 : PlaySoundAtLevelStatic SoundFX("Flipper_Right_Down_3",DOFFlippers), FlipperDownSoundLevel, Flipper
    Case 4 : PlaySoundAtLevelStatic SoundFX("Flipper_Right_Down_4",DOFFlippers), FlipperDownSoundLevel, Flipper
    Case 5 : PlaySoundAtLevelStatic SoundFX("Flipper_Right_Down_5",DOFFlippers), FlipperDownSoundLevel, Flipper
    Case 6 : PlaySoundAtLevelStatic SoundFX("Flipper_Right_Down_6",DOFFlippers), FlipperDownSoundLevel, Flipper
    Case 7 : PlaySoundAtLevelStatic SoundFX("Flipper_Right_Down_7",DOFFlippers), FlipperDownSoundLevel, Flipper
    Case 8 : PlaySoundAtLevelStatic SoundFX("Flipper_Right_Down_8",DOFFlippers), FlipperDownSoundLevel, Flipper
  End Select
End Sub

'/////////////////////////////  FLIPPER BATS BALL COLLIDE SOUND  ////////////////////////////

Sub LeftFlipperCollide(parm)
  FlipperLeftHitParm = parm/10
  If FlipperLeftHitParm > 1 Then
    FlipperLeftHitParm = 1
  End If
  FlipperLeftHitParm = FlipperUpSoundLevel * FlipperLeftHitParm
  RandomSoundRubberFlipper(parm)
End Sub

Sub RightFlipperCollide(parm)
  FlipperRightHitParm = parm/10
  If FlipperRightHitParm > 1 Then
    FlipperRightHitParm = 1
  End If
  FlipperRightHitParm = FlipperUpSoundLevel * FlipperRightHitParm
  RandomSoundRubberFlipper(parm)
End Sub

Sub RandomSoundRubberFlipper(parm)
  Select Case Int(Rnd*7)+1
    Case 1 : PlaySoundAtLevelActiveBall ("Flipper_Rubber_1"), parm  * RubberFlipperSoundFactor
    Case 2 : PlaySoundAtLevelActiveBall ("Flipper_Rubber_2"), parm  * RubberFlipperSoundFactor
    Case 3 : PlaySoundAtLevelActiveBall ("Flipper_Rubber_3"), parm  * RubberFlipperSoundFactor
    Case 4 : PlaySoundAtLevelActiveBall ("Flipper_Rubber_4"), parm  * RubberFlipperSoundFactor
    Case 5 : PlaySoundAtLevelActiveBall ("Flipper_Rubber_5"), parm  * RubberFlipperSoundFactor
    Case 6 : PlaySoundAtLevelActiveBall ("Flipper_Rubber_6"), parm  * RubberFlipperSoundFactor
    Case 7 : PlaySoundAtLevelActiveBall ("Flipper_Rubber_7"), parm  * RubberFlipperSoundFactor
  End Select
End Sub

'/////////////////////////////  ROLLOVER SOUNDS  ////////////////////////////
Sub RandomSoundRollover()
  Select Case Int(Rnd*4)+1
    Case 1 : PlaySoundAtLevelActiveBall ("Rollover_1"), RolloverSoundLevel
    Case 2 : PlaySoundAtLevelActiveBall ("Rollover_2"), RolloverSoundLevel
    Case 3 : PlaySoundAtLevelActiveBall ("Rollover_3"), RolloverSoundLevel
    Case 4 : PlaySoundAtLevelActiveBall ("Rollover_4"), RolloverSoundLevel
  End Select
End Sub

'/////////////////////////////  VARIOUS PLAYFIELD SOUND SUBROUTINES  ////////////////////////////
'/////////////////////////////  RUBBERS AND POSTS  ////////////////////////////
'/////////////////////////////  RUBBERS - EVENTS  ////////////////////////////
Sub Rubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 5 then
    RandomSoundRubberStrong 1
  End if
  If finalspeed <= 5 then
    RandomSoundRubberWeak()
  End If
End Sub

Sub RubberWheel_Hit(idx)
  'RandomSoundRubberWeakWheel
 End Sub

'/////////////////////////////  RUBBERS AND POSTS - STRONG IMPACTS  ////////////////////////////
Sub RandomSoundRubberStrong(voladj)
  Select Case Int(Rnd*10)+1
    Case 1 : PlaySoundAtLevelActiveBall ("Rubber_Strong_1"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 2 : PlaySoundAtLevelActiveBall ("Rubber_Strong_2"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 3 : PlaySoundAtLevelActiveBall ("Rubber_Strong_3"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 4 : PlaySoundAtLevelActiveBall ("Rubber_Strong_4"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 5 : PlaySoundAtLevelActiveBall ("Rubber_Strong_5"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 6 : PlaySoundAtLevelActiveBall ("Rubber_Strong_6"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 7 : PlaySoundAtLevelActiveBall ("Rubber_Strong_7"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 8 : PlaySoundAtLevelActiveBall ("Rubber_Strong_8"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 9 : PlaySoundAtLevelActiveBall ("Rubber_Strong_9"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 10 : PlaySoundAtLevelActiveBall ("Rubber_1_Hard"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6*voladj
  End Select
End Sub

'/////////////////////////////  RUBBERS AND POSTS - WEAK IMPACTS  ////////////////////////////
Sub RandomSoundRubberWeak()
  Select Case Int(Rnd*9)+1
    Case 1 : PlaySoundAtLevelActiveBall ("Rubber_1"), Vol(ActiveBall) * RubberWeakSoundFactor
    Case 2 : PlaySoundAtLevelActiveBall ("Rubber_2"), Vol(ActiveBall) * RubberWeakSoundFactor
    Case 3 : PlaySoundAtLevelActiveBall ("Rubber_3"), Vol(ActiveBall) * RubberWeakSoundFactor
    Case 4 : PlaySoundAtLevelActiveBall ("Rubber_3"), Vol(ActiveBall) * RubberWeakSoundFactor
    Case 5 : PlaySoundAtLevelActiveBall ("Rubber_5"), Vol(ActiveBall) * RubberWeakSoundFactor
    Case 6 : PlaySoundAtLevelActiveBall ("Rubber_6"), Vol(ActiveBall) * RubberWeakSoundFactor
    Case 7 : PlaySoundAtLevelActiveBall ("Rubber_7"), Vol(ActiveBall) * RubberWeakSoundFactor
    Case 8 : PlaySoundAtLevelActiveBall ("Rubber_8"), Vol(ActiveBall) * RubberWeakSoundFactor
    Case 9 : PlaySoundAtLevelActiveBall ("Rubber_9"), Vol(ActiveBall) * RubberWeakSoundFactor
  End Select
End Sub

Sub RandomSoundRubberWeakWheel()
  Select Case Int(Rnd*9)+1
    Case 1 : PlaySoundAtLevelActiveBall ("Rubber_1"), Vol(ActiveBall) * RubberWeakSoundFactor/10
    Case 2 : PlaySoundAtLevelActiveBall ("Rubber_2"), Vol(ActiveBall) * RubberWeakSoundFactor/10
    Case 3 : PlaySoundAtLevelActiveBall ("Rubber_3"), Vol(ActiveBall) * RubberWeakSoundFactor/10
    Case 4 : PlaySoundAtLevelActiveBall ("Rubber_3"), Vol(ActiveBall) * RubberWeakSoundFactor/10
    Case 5 : PlaySoundAtLevelActiveBall ("Rubber_5"), Vol(ActiveBall) * RubberWeakSoundFactor/10
    Case 6 : PlaySoundAtLevelActiveBall ("Rubber_6"), Vol(ActiveBall) * RubberWeakSoundFactor/10
    Case 7 : PlaySoundAtLevelActiveBall ("Rubber_7"), Vol(ActiveBall) * RubberWeakSoundFactor/10
    Case 8 : PlaySoundAtLevelActiveBall ("Rubber_8"), Vol(ActiveBall) * RubberWeakSoundFactor/10
    Case 9 : PlaySoundAtLevelActiveBall ("Rubber_9"), Vol(ActiveBall) * RubberWeakSoundFactor/10
  End Select
End Sub

'/////////////////////////////  WALL IMPACTS  ////////////////////////////
Sub Walls_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 5 then
    RandomSoundRubberStrong 1
  End if
  If finalspeed <= 5 then
    RandomSoundRubberWeak()
  End If
End Sub

Sub RandomSoundWall()
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    Select Case Int(Rnd*5)+1
      Case 1 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_1"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_2"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_5"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 4 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_7"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 5 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_9"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    Select Case Int(Rnd*4)+1
      Case 1 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_3"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 4 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End If
  If finalspeed < 6 Then
    Select Case Int(Rnd*3)+1
      Case 1 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End if
End Sub

'/////////////////////////////  METAL TOUCH SOUNDS  ////////////////////////////
Sub RandomSoundMetal()
  Select Case Int(Rnd*13)+1
    Case 1 : PlaySoundAtLevelActiveBall ("Metal_Touch_1"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 2 : PlaySoundAtLevelActiveBall ("Metal_Touch_2"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 3 : PlaySoundAtLevelActiveBall ("Metal_Touch_3"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 4 : PlaySoundAtLevelActiveBall ("Metal_Touch_4"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 5 : PlaySoundAtLevelActiveBall ("Metal_Touch_5"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 6 : PlaySoundAtLevelActiveBall ("Metal_Touch_6"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 7 : PlaySoundAtLevelActiveBall ("Metal_Touch_7"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 8 : PlaySoundAtLevelActiveBall ("Metal_Touch_8"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 9 : PlaySoundAtLevelActiveBall ("Metal_Touch_9"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 10 : PlaySoundAtLevelActiveBall ("Metal_Touch_10"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 11 : PlaySoundAtLevelActiveBall ("Metal_Touch_11"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 12 : PlaySoundAtLevelActiveBall ("Metal_Touch_12"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 13 : PlaySoundAtLevelActiveBall ("Metal_Touch_13"), Vol(ActiveBall) * MetalImpactSoundFactor
  End Select
End Sub

'/////////////////////////////  METAL - EVENTS  ////////////////////////////

Sub Metals_Hit (idx)
  RandomSoundMetal
End Sub

Sub DiverterFlipper_collide(idx)
  RandomSoundMetal
End Sub

'/////////////////////////////  BOTTOM ARCH BALL GUIDE  ////////////////////////////
'/////////////////////////////  BOTTOM ARCH BALL GUIDE - SOFT BOUNCES  ////////////////////////////
Sub RandomSoundBottomArchBallGuide()
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    Select Case Int(Rnd*2)+1
      Case 1 : PlaySoundAtLevelActiveBall ("Apron_Bounce_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
      Case 2 : PlaySoundAtLevelActiveBall ("Apron_Bounce_2"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
    End Select
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    Select Case Int(Rnd*2)+1
      Case 1 : PlaySoundAtLevelActiveBall ("Apron_Bounce_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
      Case 2 : PlaySoundAtLevelActiveBall ("Apron_Bounce_Soft_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
    End Select
  End If
  If finalspeed < 6 Then
    Select Case Int(Rnd*2)+1
      Case 1 : PlaySoundAtLevelActiveBall ("Apron_Bounce_Soft_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
      Case 2 : PlaySoundAtLevelActiveBall ("Apron_Medium_3"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
    End Select
  End if
End Sub

Sub Apron_Hit (idx)
  RandomSoundBottomArchBallGuide
End Sub

'/////////////////////////////  FLIPPER BALL GUIDE  ////////////////////////////
Sub RandomSoundFlipperBallGuide()
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    Select Case Int(Rnd*2)+1
      Case 1 : PlaySoundAtLevelActiveBall ("Apron_Hard_1"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
      Case 2 : PlaySoundAtLevelActiveBall ("Apron_Hard_2"),  Vol(ActiveBall) * 0.8 * FlipperBallGuideSoundFactor
    End Select
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    Select Case Int(Rnd*3)+1
      Case 1 : PlaySoundAtLevelActiveBall ("Apron_Medium_1"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
      Case 2 : PlaySoundAtLevelActiveBall ("Apron_Medium_2"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
      Case 3 : PlaySoundAtLevelActiveBall ("Apron_Medium_3"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
    End Select
  End If
  If finalspeed < 6 Then
    Select Case Int(Rnd*7)+1
      Case 1 : PlaySoundAtLevelActiveBall ("Apron_Soft_1"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
      Case 2 : PlaySoundAtLevelActiveBall ("Apron_Soft_2"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
      Case 3 : PlaySoundAtLevelActiveBall ("Apron_Soft_3"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
      Case 4 : PlaySoundAtLevelActiveBall ("Apron_Soft_4"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
      Case 5 : PlaySoundAtLevelActiveBall ("Apron_Soft_5"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
      Case 6 : PlaySoundAtLevelActiveBall ("Apron_Soft_6"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
      Case 7 : PlaySoundAtLevelActiveBall ("Apron_Soft_7"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
    End Select
  End if
End Sub

'/////////////////////////////  TARGET HIT SOUNDS  ////////////////////////////
Sub RandomSoundTargetHitStrong()
  Select Case Int(Rnd*4)+1
    Case 1 : PlaySoundAtLevelActiveBall SoundFX("Target_Hit_5",DOFTargets), Vol(ActiveBall) * 0.45 * TargetSoundFactor
    Case 2 : PlaySoundAtLevelActiveBall SoundFX("Target_Hit_6",DOFTargets), Vol(ActiveBall) * 0.45 * TargetSoundFactor
    Case 3 : PlaySoundAtLevelActiveBall SoundFX("Target_Hit_7",DOFTargets), Vol(ActiveBall) * 0.45 * TargetSoundFactor
    Case 4 : PlaySoundAtLevelActiveBall SoundFX("Target_Hit_8",DOFTargets), Vol(ActiveBall) * 0.45 * TargetSoundFactor
  End Select
End Sub

Sub RandomSoundTargetHitWeak()
  Select Case Int(Rnd*4)+1
    Case 1 : PlaySoundAtLevelActiveBall SoundFX("Target_Hit_1",DOFTargets), Vol(ActiveBall) * TargetSoundFactor
    Case 2 : PlaySoundAtLevelActiveBall SoundFX("Target_Hit_2",DOFTargets), Vol(ActiveBall) * TargetSoundFactor
    Case 3 : PlaySoundAtLevelActiveBall SoundFX("Target_Hit_3",DOFTargets), Vol(ActiveBall) * TargetSoundFactor
    Case 4 : PlaySoundAtLevelActiveBall SoundFX("Target_Hit_4",DOFTargets), Vol(ActiveBall) * TargetSoundFactor
  End Select
End Sub

Sub PlayTargetSound()
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 10 then
    RandomSoundTargetHitStrong()
    RandomSoundBallBouncePlayfieldSoft()
  Else
    RandomSoundTargetHitWeak()
  End If
End Sub

Sub Targets_Hit (idx)
  PlayTargetSound
End Sub

'/////////////////////////////  BALL BOUNCE SOUNDS  ////////////////////////////
Sub RandomSoundBallBouncePlayfieldSoft()
  Select Case Int(Rnd*9)+1
    Case 1 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_1"), volz(BSBall) * BallBouncePlayfieldSoftFactor, BSBall
    Case 2 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_2"), volz(BSBall) * BallBouncePlayfieldSoftFactor * 0.5, BSBall
    Case 3 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_3"), volz(BSBall) * BallBouncePlayfieldSoftFactor * 0.8, BSBall
    Case 4 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_4"), volz(BSBall) * BallBouncePlayfieldSoftFactor * 0.5, BSBall
    Case 5 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_5"), volz(BSBall) * BallBouncePlayfieldSoftFactor, BSBall
    Case 6 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_1"), volz(BSBall) * BallBouncePlayfieldSoftFactor * 0.2, BSBall
    Case 7 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_2"), volz(BSBall) * BallBouncePlayfieldSoftFactor * 0.2, BSBall
    Case 8 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_5"), volz(BSBall) * BallBouncePlayfieldSoftFactor * 0.2, BSBall
    Case 9 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_7"), volz(BSBall) * BallBouncePlayfieldSoftFactor * 0.3, BSBall
  End Select
End Sub

Sub RandomSoundBallBouncePlayfieldHard()
  Select Case Int(Rnd*7)+1
    Case 1 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_1"), volz(BSBall) * BallBouncePlayfieldHardFactor, BSBall
    Case 2 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_2"), volz(BSBall) * BallBouncePlayfieldHardFactor, BSBall
    Case 3 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_3"), volz(BSBall) * BallBouncePlayfieldHardFactor, BSBall
    Case 4 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_4"), volz(BSBall) * BallBouncePlayfieldHardFactor, BSBall
    Case 5 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_5"), volz(BSBall) * BallBouncePlayfieldHardFactor, BSBall
    Case 6 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_6"), volz(BSBall) * BallBouncePlayfieldHardFactor, BSBall
    Case 7 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_7"), volz(BSBall) * BallBouncePlayfieldHardFactor, BSBall
  End Select
End Sub

'/////////////////////////////  DELAYED DROP - TO PLAYFIELD - SOUND  ////////////////////////////
Sub RandomSoundDelayedBallDropOnPlayfield()
  Select Case Int(Rnd*5)+1
    Case 1 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_1_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, BSBall
    Case 2 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_2_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, BSBall
    Case 3 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_3_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, BSBall
    Case 4 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_4_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, BSBall
    Case 5 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_5_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, BSBall
  End Select
End Sub

'/////////////////////////////  BALL GATES AND BRACKET GATES SOUNDS  ////////////////////////////

Sub SoundPlayfieldGate()
  Select Case Int(Rnd*2)+1
    Case 1 : PlaySoundAtLevelStatic ("Gate_FastTrigger_1"), GateSoundLevel, Activeball
    Case 2 : PlaySoundAtLevelStatic ("Gate_FastTrigger_2"), GateSoundLevel, Activeball
  End Select
End Sub

Sub SoundHeavyGate()
  PlaySoundAtLevelStatic ("Gate_2"), GateSoundLevel, Activeball
End Sub


'/////////////////////////////  LEFT LANE ENTRANCE - SOUNDS  ////////////////////////////

Sub RandomSoundLeftArch()
  Select Case Int(Rnd*4)+1
    Case 1 : PlaySoundAtLevelActiveBall ("Arch_L1"), Vol(ActiveBall) * ArchSoundFactor
    Case 2 : PlaySoundAtLevelActiveBall ("Arch_L2"), Vol(ActiveBall) * ArchSoundFactor
    Case 3 : PlaySoundAtLevelActiveBall ("Arch_L3"), Vol(ActiveBall) * ArchSoundFactor
    Case 4 : PlaySoundAtLevelActiveBall ("Arch_L4"), Vol(ActiveBall) * ArchSoundFactor
  End Select
End Sub

Sub RandomSoundRightArch()
  Select Case Int(Rnd*4)+1
    Case 1 : PlaySoundAtLevelActiveBall ("Arch_R1"), Vol(ActiveBall) * ArchSoundFactor
    Case 2 : PlaySoundAtLevelActiveBall ("Arch_R2"), Vol(ActiveBall) * ArchSoundFactor
    Case 3 : PlaySoundAtLevelActiveBall ("Arch_R3"), Vol(ActiveBall) * ArchSoundFactor
    Case 4 : PlaySoundAtLevelActiveBall ("Arch_R4"), Vol(ActiveBall) * ArchSoundFactor
  End Select
End Sub


Sub Arch1_hit()
  If Activeball.velx > 1 Then SoundPlayfieldGate
  StopSound "Arch_L1"
  StopSound "Arch_L2"
  StopSound "Arch_L3"
  StopSound "Arch_L4"
End Sub

Sub Arch1_unhit()
  If activeball.vely < -7 Then
    RandomSoundRightArch
  End If
  'debug.print activeball.vely
End Sub

Sub Arch2_hit()
  StopSound "Arch_R1"
  StopSound "Arch_R2"
  StopSound "Arch_R3"
  StopSound "Arch_R4"
End Sub

Sub Arch2_unhit()
  If activeball.vely < -7 Then
    RandomSoundLeftArch
  End If
  'debug.print activeball.vely
End Sub

Sub Gates_Hit (idx)
  SoundHeavyGate
End Sub

'******************************************************
'                        FLIPPER TRICKS
'******************************************************

RightFlipper.timerinterval=1
Rightflipper.timerenabled=True

sub RightFlipper_timer()
        FlipperTricks LeftFlipper, LFPress, LFCount, LFEndAngle, LFState
        FlipperTricks RightFlipper, RFPress, RFCount, RFEndAngle, RFState
end sub

dim LFPress, RFPress, LFCount, RFCount
dim LFState, RFState
dim EOST, EOSA,Frampup, FElasticity,FReturn
dim RFEndAngle, LFEndAngle

EOST = leftflipper.eostorque
EOSA = leftflipper.eostorqueangle
Frampup = LeftFlipper.rampup
FElasticity = LeftFlipper.elasticity
FReturn = LeftFlipper.return
Const EOSTnew = 1.5
Const EOSAnew = 1
Const EOSRampup = 0
Dim SOSRampup
SOSRampup = 2.5
Const LiveCatch = 16
Const LiveElasticity = 0.45
Const SOSEM = 0.815
Const EOSReturn = 0.055

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
                Dim BOT, b
                BOT = GetBalls

                For b = 0 to UBound(BOT)
                        If Distance(BOT(b).x, BOT(b).y, Flipper.x, Flipper.y) < 55 Then 'check for cradle
                                If BOT(b).vely >= -0.4 Then BOT(b).vely = -0.4
                        End If
                Next
        End If
End Sub

Sub FlipperTricks (Flipper, FlipperPress, FCount, FEndAngle, FState)
        Dim Dir
        Dir = Flipper.startangle/Abs(Flipper.startangle)        '-1 for Right Flipper

        If Abs(Flipper.currentangle) > Abs(Flipper.startangle) - 0.05 Then
                If FState <> 1 Then
                        Flipper.rampup = SOSRampup
                        Flipper.endangle = FEndAngle - 3*Dir
                        Flipper.Elasticity = FElasticity * SOSEM
                        FCount = 0
                        FState = 1
                End If
        ElseIf Abs(Flipper.currentangle) <= Abs(Flipper.endangle) and FlipperPress = 1 then
                if FCount = 0 Then FCount = GameTime

                If FState <> 2 Then
                        Flipper.eostorqueangle = EOSAnew
                        Flipper.eostorque = EOSTnew
                        Flipper.rampup = EOSRampup
                        Flipper.endangle = FEndAngle
                        FState = 2
                End If
        Elseif Abs(Flipper.currentangle) > Abs(Flipper.endangle) + 0.01 and FlipperPress = 1 Then
                If FState <> 3 Then
                        Flipper.eostorque = EOST
                        Flipper.eostorqueangle = EOSA
                        Flipper.rampup = Frampup
                        Flipper.Elasticity = FElasticity
                        FState = 3
                End If

        End If
End Sub

Const LiveDistanceMin = 30  'minimum distance in vp units from flipper base live catch dampening will occur
Dim LiveDistanceMax: LiveDistanceMax = leftflipper.length  'maximum distance in vp units from flipper base live catch dampening will occur (tip protection)

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
    If Flipper.currentangle = Flipper.endangle Then RubbersD.Dampen Activeball
  End If
End Sub

Sub LeftFlipper_Collide(parm)
  'CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
  LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
  'CheckLiveCatch Activeball, RightFlipper, RFCount, parm
  RightFlipperCollide parm
End Sub

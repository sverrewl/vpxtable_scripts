'*************************************
'Jive Time (Williams 1970) - IPD No. 1298
'  Playfield - rothbauerw, bord
'  Backglass - rothbauerw
'  VPX Table - rothbauerw
'  Primitives - bord, 32assassin, kiwi, acronovum
'  Misc - borgdog, loserman, randr, acronovum, STAT
'
'  Special thanks to all those I borrowed tidbits from
'************************************

Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the Controller.vbs file in order to run this table (installed with the VPX package in the scripts folder)"
On Error Goto 0

Const ballsize = 25
Const ballmass = 1

Const cGameName = "jivetime"

'******************************************************
'             OPTIONS
'******************************************************

Const BallsPerGame = 5        '3, 5, or 10
Const WestGateOption = 1      '1 - close west gate after scoring special, 0 - keep west gate open
Const RestoreGame = 1       '1 - to remember last game on table close, 0 - to start with game over

Const VolumeDial = 2        'Change bolume of hit events
Const RollingSoundFactor = 5    'set sound level factor here for Ball Rolling Sound, 1=default level
Const ChimesOn=0          'This table uses bells, 1 turns on DOF Chimes

'******************************************************
'             VARIABLES
'******************************************************

Dim i
Dim obj
Dim InProgress
Dim BallInPlay
Dim Score
Dim HighScore
Dim Initials
Dim Credits
Dim Match
Dim Bonus
Dim SpecialLightState
Dim Replay1:Replay1 = 100000
Dim Replay2:Replay2 = 140000
Dim Replay3:Replay3 = 210000
Dim Replay4:Replay4 = 250000
Dim TableTilted
Dim TiltCount
Dim GameOn:GameOn = 0
Dim Reel1Value, Reel2Value, Reel3Value, Reel4Value, Reel5Value
Dim KickerBall

'******************************************************
'             TABLE INIT
'******************************************************

Sub Table1_Init
  LoadEM

  For each obj in GI_Flashers_DT:obj.visible = 0: Next
  For each obj in GI_Flashers:obj.visible = 0: Next
  For each obj in Targets:obj.BlendDisableLighting = 0: Next

  loadhs

  if HSA1="" then HSA1=23
  if HSA2="" then HSA2=10
  if HSA3="" then HSA3=18
  UpdatePostIt

  set KickerBall = Drain.CreateBall
  TableTilted=false
  InProgress=false
  centerpostcollide.collidable = False
  minipostcollide.collidable = False


  If Table1.ShowDT = True and Table1.ShowFSS = False then
    emp1r7.visible = true
    emp1r8.visible = true
    emp1r9.visible = true
    emp1r10.visible = true
    emp1r11.visible = true
    emp1r12.visible = true
    emcredits5.visible = true
    emcredits6.visible = true
    emcredits7.visible = true
    emcredits8.visible = true
    backglass2.visible = True
    spinner1.visible = True
    backglass3.visible = True
  End If

  If Credits > 0 Then DOF 125, 1
End Sub

Sub Table1_Exit()
  savehs
  DOF 101, 0
  DOF 102, 0
  DOF 123, 0
  DOF 125, 0
  If B2SOn Then Controller.Stop
End Sub

Dim BootCount:BootCount = 0

Sub BootTable_Timer()
  If BootCount = 0 Then
    BootCount = 1
    me.interval = 100
    PlaySoundAt "poweron", Plunger
  Else
    if Bonus = 0 Then Bonus = 1000
    EVAL("L" & Bonus).state = 1

    If match = 0 Then
      BackglassMatch.image = "JiveTimeMatch00"
    Else
      BackglassMatch.image = "JiveTimeMatch" & Match
    End If

    If RestoreGame = 0 Then BallInPlay = 0

    If BallInPlay <> 0 Then
      BackglassBall.image = "JiveTimeBall" & BallInPlay
      SpecialLight.state = SpecialLightState
      BackglassMatch.visible = False
      If B2SOn Then Controller.B2SSetMatch 0
      InProgress = True
      Drain_Hit
    Else
      BackglassMatch.visible = True
      SetB2SMatch
      If B2SOn Then Controller.B2SSetGameOver 1
    End If

    '*****GI Lights On
    For each obj in GI:obj.State = 1: Next
    For each obj in GI_Flashers:obj.visible = 1: Next
    For each obj in Targets:obj.BlendDisableLighting = 0.3: Next
    If Table1.ShowDT = True  and Table1.ShowFSS = False then
      For each obj in GI_Flashers_DT:obj.visible = 1: Next
      BackglassGI1.visible = True
      BackglassBall1.visible = True
    End If
    BumperCap1.BlendDisableLighting = 0.1

    If Credits > 0 Then Lcredit.state = 1

    BackglassGI.visible = True
    BackglassBall.visible = True

    GameOn = 1

    If B2SOn Then
      Controller.B2SSetData 80,1
      Controller.B2SSetBallInPlay BallInPlay
    End If

    me.enabled = False
  End If
End Sub

Sub BootB2S_Timer()
  If B2SOn Then
    Controller.B2SSetCredits Credits
    Controller.B2SSetScore 1,Score
    SetB2SSpinner
  End If
  me.enabled = False
End Sub

'******************************************************
'             KEYS
'******************************************************

Dim EnableBallControl
EnableBallControl = false 'Change to true to enable manual ball control (or press C in-game) via the arrow keys and B (boost movement) keys

Sub Table1_KeyDown(ByVal keycode)

  If keycode = PlungerKey Then
    Plunger.PullBack
    PlaySound "plungerpull",0,1,AudioPan(Plunger),0.25,0,0,1,AudioFade(Plunger)
  End If

  If keycode = LeftFlipperKey and InProgress=true and TableTilted=false Then
    LeftFlipper.RotateToEnd
    PlaySound SoundFX("fx_flipperup",DOFFlippers), 0, .67, AudioPan(LeftFlipper), 0.05,0,0,1,AudioFade(LeftFlipper)
    DOF 101, 1
    PlaySoundAtVolLoops "buzzL",LeftFlipper,0.05,-1
  End If

  If keycode = RightFlipperKey and InProgress=true and TableTilted=false Then
    RightFlipper.RotateToEnd
    PlaySound SoundFX("fx_flipperup",DOFFlippers), 0, .67, AudioPan(RightFlipper), 0.05,0,0,1,AudioFade(RightFlipper)
    DOF 102, 1
    PlaySoundAtVolLoops "buzz",LeftFlipper,0.05,-1
  End If

  If keycode = LeftTiltKey Then
    Nudge 90, 2
    checktilt
  End If

  If keycode = RightTiltKey Then
    Nudge 270, 2
    checktilt
  End If

  If keycode = CenterTiltKey Then
    Nudge 0, 2
    checktilt
  End If

  If keycode = MechanicalTilt Then
    gametilted
  End If

  If keycode = AddCreditKey And GameOn = 1 then
    playsound "coinin"
    AddCredit(1)
  end if

  if keycode=StartGameKey and Credits>0 and InProgress=false And GameOn = 1 And Not HSEnterMode=true then
    AddCredit(-1)
    InProgress=true
    StartNewGame.enabled = True
  end if

  If HSEnterMode Then HighScoreProcessKey(keycode)

    ' Manual Ball Control
  If keycode = 46 Then          ' C Key
    If EnableBallControl = 1 Then
      EnableBallControl = 0
    Else
      EnableBallControl = 1
    End If
  End If
    If EnableBallControl = 1 Then
    If keycode = 48 Then        ' B Key
      If BCboost = 1 Then
        BCboost = BCboostmulti
      Else
        BCboost = 1
      End If
    End If
    If keycode = 203 Then BCleft = 1  ' Left Arrow
    If keycode = 200 Then BCup = 1    ' Up Arrow
    If keycode = 208 Then BCdown = 1  ' Down Arrow
    If keycode = 205 Then BCright = 1 ' Right Arrow
  End If
End Sub

Sub Table1_KeyUp(ByVal keycode)

  If keycode = PlungerKey Then
    Plunger.Fire
    PlaySound "plunger",0,1,AudioPan(Plunger),0.25,0,0,1,AudioFade(Plunger)
  End If

  If keycode = LeftFlipperKey and InProgress=true and TableTilted=false Then
    LeftFlipper.RotateToStart
    PlaySound SoundFX("fx_flipperdown",DOFFlippers), 0, 1, AudioPan(LeftFlipper), 0.05,0,0,1,AudioFade(LeftFlipper)
    DOF 101, 0
    StopSound "buzzL"
  End If

  If keycode = RightFlipperKey and InProgress=true and TableTilted=false Then
    RightFlipper.RotateToStart
    PlaySound SoundFX("fx_flipperdown",DOFFlippers), 0, 1, AudioPan(RightFlipper), 0.05,0,0,1,AudioFade(RightFlipper)
    DOF 102, 0
    StopSound "buzz"
  End If

    'Manual Ball Control
  If EnableBallControl = 1 Then
    If keycode = 203 Then BCleft = 0  ' Left Arrow
    If keycode = 200 Then BCup = 0    ' Up Arrow
    If keycode = 208 Then BCdown = 0  ' Down Arrow
    If keycode = 205 Then BCright = 0 ' Right Arrow
  End If
End Sub



'******************************************************
'             SWITCHES
'******************************************************

Sub TopButton_Hit()
  AddScore(1000)
  playsound "sensor", 0,0.1,AudioPan(ActiveBall),0,0,0,1,AudioFade(ActiveBall)
  DOF 129, 2
End Sub

Sub UpLampPost_Hit()
  CenterPostDir = 1
  MoveCenterPost.enabled = True
  AddScore(1000)
  DOF 114, 2
End Sub

Sub UpMiniPost_Hit()
  MiniPostDir = 1
  MoveMiniPost.enabled = True
  AddScore(1000)
  DOF 117, 2
End Sub

Sub DownPostButton_Hit()
  DownPosts
  AddScore(1000)
  playsound "sensor", 0,0.1,AudioPan(ActiveBall),0,0,0,1,AudioFade(ActiveBall)
  DOF 130, 2
End Sub

Sub DownPostTarget_Hit()
  DownPosts
  AddScore(1000)
  DOF 118, 2
End Sub

Sub GreenOnTarget_Hit()
  GreenLightsOn
  AddScore(1000)
  DOF 110, 2
End Sub

Sub YellowOnTarget_Hit()
  YellowLightsOn
  AddScore(1000)
  DOF 112, 2
End Sub

Sub LeftOutLane_Hit()
  AddScore(1000)
  playsound "sensor", 0,0.1,AudioPan(ActiveBall),0,0,0,1,AudioFade(ActiveBall)
  DOF 132, 2
End Sub

Sub RightOutLane_Hit()
  AddScore(1000)
  playsound "sensor", 0,0.1,AudioPan(ActiveBall),0,0,0,1,AudioFade(ActiveBall)
  DOF 133, 2
End Sub

Sub ABTarget1_Hit()
  AdvanceBonus
  DOF 111, 2
End Sub

Sub ABTarget2_Hit()
  AdvanceBonus
  DOF 113, 2
End Sub

Sub ABTarget3_Hit()
  AdvanceBonus
  DOF 115, 2
End Sub

Sub ABTarget4_Hit()
  AdvanceBonus
  DOF 116, 2
End Sub

Sub T10Points1_Hit()
  RubberLM1.Visible = 1:RubberLM.Visible = 0:T10Points1.TimerEnabled = 1
  AddScore(10)
End Sub

Sub T10Points2_Hit()
  RubberLU1.Visible = 1:RubberLU.Visible = 0:T10Points2.TimerEnabled = 1
  AddScore(10)
End Sub

Sub T10Points3_Hit()
  AddScore(10)
End Sub

Sub T10Points4_Hit()
  RubberRUM1.Visible = 1:RubberRUM.Visible = 0:T10Points4.TimerEnabled = 1
  AddScore(10)
End Sub

Sub T10Points5_Hit()
  RubberRU1.Visible = 1:RubberRU.Visible = 0:T10Points5.TimerEnabled = 1
  AddScore(10)
End Sub

Sub ShooterLane_Unhit()
  DOF 126, 2
End Sub


'******************************************************
'             POSTS
'******************************************************

Dim MiniPostDir, CenterPostDir, MovePostSpeed
MoveMiniPost.interval = 1
MoveCenterPost.interval = 1
MovePostSpeed = 5

Sub DownPosts()
  MiniPostDir = -1
  CenterPostDir = -1
  MoveMiniPost.enabled = true
  MoveCenterPost.enabled = true
End Sub

Sub MoveMiniPost_timer()
  MiniPost.z = MiniPost.z + MiniPostDir*MovePostSpeed
  if MiniPost.z > 25 Then
    minipostcollide.collidable = True
  Else
    minipostcollide.collidable = False
  End If

  If MiniPost.z < 0.1 Then
    MiniPost.z = 0
    me.enabled = False
    DOF 127, 2
  End If
  If MiniPost.z > 49.9 Then
    MiniPost.z = 50
    me.enabled = False
    DOF 127, 2
  End If
End Sub

Sub MoveCenterPost_timer()
  CenterPost.z = CenterPost.z + CenterPostDir*MovePostSpeed/5

  if CenterPost.z > 13 Then
    centerpostcollide.collidable = True
  Else
    centerpostcollide.collidable = False
  End If


  If CenterPost.z < 6 Then
'   CenterPost.image = "3D_Centerpost"
    CenterPost.BlendDisableLighting = 0
    lp1.state=0
    lp2.state=0
  Elseif CenterPost.z < 12 Then
'   CenterPost.image = "3D_Centerpost_1"
    CenterPost.BlendDisableLighting = 0.1
    lp1.state=1
    lp2.state=0
  Elseif CenterPost.z < 18 Then
'   CenterPost.image = "3D_Centerpost_2"
    CenterPost.BlendDisableLighting = 0.2
    lp1.state=1
    lp2.state=0
  Else
'   CenterPost.image = "3D_Centerpost_2"
    CenterPost.BlendDisableLighting = 0.3
    lp1.state=1
    lp2.state=0
  End If

  If CenterPost.z < 0.1 Then
    CenterPost.z = -0
    me.enabled = False
    lp1.state=0
    lp2.state=0
    DOF 128, 2
  End If
  If CenterPost.z > 24.9 Then
    CenterPost.z = 25
    lp2.state=1
    me.enabled = False
    DOF 128, 2
  End If
End Sub

'******************************************************
'             BUMPERS
'******************************************************

Sub Bumper1_Hit()
  PlaySoundAtVol "fx_bumper1", Bumper1, 1
  DOF 107, 2
  AddScore(100)

  BumperSkirt1.roty=skirtAY(me,Activeball)
  BumperSkirt1.rotx=skirtAX(me,Activeball)
  me.timerenabled=1
End Sub

sub bumper1_timer     '****************part of bumper skirt animation
  BumperSkirt1.rotx=0
  BumperSkirt1.roty=0
  me.timerenabled=0
end sub

Sub Bumper2_Hit()
  PlaySoundAtVol "fx_bumper2", Bumper2, 1
  DOF 105, 2
  PlaySoundAtVol "fx_bumper3", Bumper3, 1
  DOF 109, 2
  Bumper3.PlayHit
  If GIG1.state = 1 Then
    AddScore(1000)
  Else
    AddScore(100)
  End If

  BumperSkirt2.roty=skirtAY(me,Activeball)
  BumperSkirt2.rotx=skirtAX(me,Activeball)
  me.timerenabled=1
End Sub

sub bumper2_timer     '****************part of bumper skirt animation
  BumperSkirt2.rotx=0
  BumperSkirt2.roty=0
  me.timerenabled=0
end sub

Sub Bumper3_Hit()
  PlaySoundAtVol "fx_bumper2", Bumper2, 1
  DOF 105, 2
  PlaySoundAtVol "fx_bumper3", Bumper3, 1
  DOF 109, 2
  Bumper2.PlayHit
  If GIG1.state = 1 Then
    AddScore(1000)
  Else
    AddScore(100)
  End If

  BumperSkirt3.roty=skirtAY(me,Activeball)
  BumperSkirt3.rotx=skirtAX(me,Activeball)
  me.timerenabled=1
End Sub

sub bumper3_timer     '****************part of bumper skirt animation
  BumperSkirt3.rotx=0
  BumperSkirt3.roty=0
  me.timerenabled=0
end sub

Sub Bumper4_Hit()
  PlaySoundAtVol "fx_bumper4", Bumper4, 1
  DOF 108, 2
  PlaySoundAtVol "fx_bumper5", Bumper5, 1
  DOF 106, 2
  Bumper5.PlayHit
  If GIY1.state = 1 Then
    AddScore(1000)
  Else
    AddScore(100)
  End If

  BumperSkirt4.roty=skirtAY(me,Activeball)
  BumperSkirt4.rotx=skirtAX(me,Activeball)
  me.timerenabled=1
End Sub

sub bumper4_timer     '****************part of bumper skirt animation
  BumperSkirt4.rotx=0
  BumperSkirt4.roty=0
  me.timerenabled=0
end sub

Sub Bumper5_Hit()
  PlaySoundAtVol "fx_bumper4", Bumper4, 1
  DOF 108, 2
  PlaySoundAtVol "fx_bumper5", Bumper5, 1
  DOF 106, 2
  Bumper4.PlayHit
  If GIY1.state = 1 Then
    AddScore(1000)
  Else
    AddScore(100)
  End If

  BumperSkirt5.roty=skirtAY(me,Activeball)
  BumperSkirt5.rotx=skirtAX(me,Activeball)
  me.timerenabled=1
End Sub

sub bumper5_timer     '****************part of bumper skirt animation
  BumperSkirt5.rotx=0
  BumperSkirt5.roty=0
  me.timerenabled=0
end sub

Sub GreenLightsOn()
  dim xx
  For each xx in GreenLights:xx.State = 1: Next
  BumperCap2.BlendDisableLighting = 0.1
  BumperCap3.BlendDisableLighting = 0.1
End Sub

Sub GreenLightsOff()
  dim xx
  For each xx in GreenLights:xx.State = 0: Next
  BumperCap2.BlendDisableLighting = 0
  BumperCap3.BlendDisableLighting = 0
End Sub

Sub YellowLightsOn()
  dim xx
  For each xx in YellowLights:xx.State = 1: Next
  BumperCap4.BlendDisableLighting = 0.1
  BumperCap5.BlendDisableLighting = 0.1
End Sub

Sub YellowLightsOff()
  dim xx
  For each xx in YellowLights:xx.State = 0: Next
  BumperCap4.BlendDisableLighting = 0
  BumperCap5.BlendDisableLighting = 0
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

'******************************************************
'             WEST GATE
'******************************************************

Sub WestGateButton2_Hit()
  playsound "sensor", 0,0.1,AudioPan(ActiveBall),0,0,0,1,AudioFade(ActiveBall)
  DOF 131, 2
  If WestGateLight.state = 1 Then
    AddBall
    If WestGateOption = 1 Then
      CloseWestGate
    End If
  Else
    AddScore(1000)
  End If
End Sub

Sub OpenWestGate()
  DiverterFlipper.rotatetoend
  DOF 122, 2
End Sub

Sub CloseWestGate()
  DiverterFlipper.rotatetostart
  DOF 122, 2
End Sub

'******************************************************
'             KICKERS
'******************************************************

TopKicker.timerinterval = 500
dim SelectKicker

Sub TopKicker_Hit
  SelectKicker = "Top"
  PlaySound "kicker_enter_center", 0, 0.1, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  If TableTilted = 0 Then
    If SpecialLight.state = 1 Then
      SpecialLight.state = 0
      AddBall
    End If
    RunSpinner
  Else
    TopKicker.timerenabled = True
  End If
End Sub

Sub TopKicker_Timer()
  TopKicker.Kick 168 + rnd(4), 13
  MiddleKicker.Kick 167 + rnd(4),13
  TWKicker1.TransY = 20
  TWKicker2.TransY = 20
  MiddleKicker.timerenabled = true
  TopKicker.timerenabled = false
  If SelectKicker = "Top" Then
    PlaySoundAtVol "popper_ball", TopKicker, 0.5
    PlaySoundAtVol "solenoid", MiddleKicker, 0.5
  Else
    PlaySoundAtVol "solenoid", TopKicker, 0.5
    PlaySoundAtVol "popper_ball", MiddleKicker, 0.5
  End If
  MiddleKicker.enabled = False
  TopKicker.enabled = False
  DOF 119, 2
  DOF 120, 2
End Sub

Sub MiddleKicker_Hit
  SelectKicker = "Middle"
  PlaySound "kicker_enter_center", 0, 0.1, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  If TableTilted = 0 Then
    RunSpinner
  Else
    TopKicker.timerenabled = True
  End If
End Sub

Sub MiddleKicker_Timer()
  TWKicker1.TransY = 0
  TWKicker2.TransY = 0
  MiddleKicker.timerenabled = False
End Sub

'******************************************************
'             SPINNER
'******************************************************

Dim SpinnerAngle, SpinnerTime, SpinnerRunTime, SpinnerStep
SpinnerTimer.interval = 10

Sub SpinnerTimer_Timer()
  SpinnerRunTime = SpinnerRunTime + (me.interval/1000)
  if SpinnerRunTime < SpinnerTime Then
    SpinnerStep = 12
  Else
    DOF 123, 0
    SpinnerStep = SpinnerStep-0.25
    If SpinnerStep < 0 Then
      SpinnerStep = 0
      me.enabled = False
    End If
  End If
  SpinnerAngle = SpinnerAngle + SpinnerStep
  If SpinnerAngle > 360 Then
    SpinnerAngle = SpinnerAngle - 360
  End If
  SetB2SSpinner
End Sub

Sub RunSpinner
  SpinnerTime = 1.25 + Int(Rnd*25)/100
  SpinnerRunTime = 0
  SpinnerTimer.enabled = true
  SpinnerScore.enabled = true
  PlaySoundAtVol "Spinner", Spinner, 0.5
  DOF 123, 1
End Sub

Sub SpinnerScore_Timer
  SpinnerAward
  me.enabled = False
End Sub

Sub SpinnerAward()
  If SpinnerAngle > 339 Or SpinnerAngle <= 15 Then
    AddScore(10000)
    TopKicker.timerenabled = true
  Elseif SpinnerAngle > 15 And SpinnerAngle <= 51 Then
    OpenWestGate
    MultiScore 5, 100
    TopKicker.timerenabled = true
  Elseif SpinnerAngle > 51 And SpinnerAngle <= 87 Then
    BonusTimer.enabled = true
  Elseif SpinnerAngle > 87 And SpinnerAngle <= 123 Then
    AddScore(1000)
    TopKicker.timerenabled = true
  Elseif SpinnerAngle > 123 And SpinnerAngle <= 159 Then
    MultiScore 3, 1000
    TopKicker.timerenabled = true
  Elseif SpinnerAngle > 159 And SpinnerAngle <= 195 Then
    DoubleBonusTimer.enabled = 1
    MultiScore 2.5, 1000
  Elseif SpinnerAngle > 195 And SpinnerAngle <= 231 Then
    AddScore(1000)
    TopKicker.timerenabled = true
  Elseif SpinnerAngle > 231 And SpinnerAngle <= 267 Then
    OpenWestGate
    MultiScore 5, 100
    TopKicker.timerenabled = true
  Elseif SpinnerAngle > 267 And SpinnerAngle <= 303 Then
    BonusTimer.enabled = true
  Elseif SpinnerAngle > 303 And SpinnerAngle <= 339 Then
    MultiScore 3, 1000
    TopKicker.timerenabled = true
  End If

End Sub

Dim PrevPos

Sub SetB2SSpinner()
  If B2SOn Then
    Dim spos
    spos = Int(((SpinnerAngle+4) mod 360)/4)

    If PrevPos <> 0 Then
      If PrevPos > 50 Then
        Controller.B2SSetData PrevPos+109, 0
      Else
        Controller.B2SSetData PrevPos+200, 0
      End If

    End If

    If spos = 0 Then spos = 90

    PrevPos = sPos

    If sPos > 50 Then
      Controller.B2SSetData sPos+109, 1
    Else
      Controller.B2SSetData sPos+200, 1
    End If
  End If
End Sub

Sub SpinnerTest_Timer()
  SpinnerAngle = SpinnerAngle + 0.1
  SetB2SSpinner
End Sub

'******************************************************
'             BONUS
'******************************************************

Dim BonusCount, AdvBonus

Sub BonusTimer_Timer()
  BonusCount = BonusCount + 1

  Select Case (BonusCount)
    Case 1:
      PlaySoundAtVol "MotorRunning", Spinner, 0.25
      PlayLoudClick
      CheckBonus
    Case 2:
      PlayLoudClick
      CheckBonus
    Case 3:
      PlayLoudClick
      CheckBonus
    Case 4:
      PlayLoudClick
      CheckBonus
    Case 5:
      PlayLoudClick
      CheckBonus

    Case 7:
      PlayQuietClick
      CheckBonus
    Case 8:
      PlayQuietClick
      CheckBonus
    Case 9:
      PlayQuietClick
      CheckBonus
    Case 10:
      PlayQuietClick
      CheckBonus
    Case 11:
      PlayQuietClick
      CheckBonus

    Case 13:
      L1000.state = 1
      TopKicker.timerenabled = true
      me.enabled = false
      BonusCount = 0
  End Select
End Sub

Sub DoubleBonusTimer_Timer()
  If L1000.state = 1 or L2000.state = 1 or L3000.state = 1 or L4000.state = 1 or L5000.state = 1 or L6000.state = 1 or L7000.state = 1 or L8000.state = 1 or L9000.state = 1 or L10000.state = 1 Then
    MultiScore 2.5, 1000
  Else
    L1000.state = 1
    TopKicker.timerenabled = true
    me.enabled = false
  End If
End Sub

Sub CheckBonus()
  If L1000.state = 1 or L2000.state = 1 or L3000.state = 1 or L4000.state = 1 or L5000.state = 1 or L6000.state = 1 or L7000.state = 1 or L8000.state = 1 or L9000.state = 1 or L10000.state = 1 Then
    DecreaseBonus
    AddScore(1000)
  End If
End Sub

Sub AdvanceBonus()
  if isScoring() = 0 Then
    AddScore 1000
    If L1000.state = 1 Then
      L2000.state = 1
      L1000.state = 0
      Bonus = 2000
      PlayLoudClick
    ElseIf L2000.state = 1 Then
      L3000.state = 1
      L2000.state = 0
      Bonus = 3000
      PlayLoudClick
    ElseIf L3000.state = 1 Then
      L4000.state = 1
      L3000.state = 0
      Bonus = 4000
      PlayLoudClick
    ElseIf L4000.state = 1 Then
      L5000.state = 1
      L4000.state = 0
      Bonus = 5000
      PlayLoudClick
    ElseIf L5000.state = 1 Then
      L6000.state = 1
      L5000.state = 0
      Bonus = 6000
      PlayLoudClick
    ElseIf L6000.state = 1 Then
      L7000.state = 1
      L6000.state = 0
      Bonus = 7000
      PlayLoudClick
    ElseIf L7000.state = 1 Then
      L8000.state = 1
      L7000.state = 0
      Bonus = 8000
      PlayLoudClick
    ElseIf L8000.state = 1 Then
      L9000.state = 1
      L8000.state = 0
      Bonus = 9000
      PlayLoudClick
    ElseIf L9000.state = 1 Then
      L10000.state = 1
      L9000.state = 0
      Bonus = 10000
      PlayLoudClick
    End If
  End If
End Sub

Sub DecreaseBonus()
  PlayLoudClick
  If L1000.state = 1 Then
    L1000.state = 0
    Bonus = 1000
  ElseIf L2000.state = 1 Then
    L1000.state = 1
    L2000.state = 0
  ElseIf L3000.state = 1 Then
    L2000.state = 1
    L3000.state = 0
  ElseIf L4000.state = 1 Then
    L3000.state = 1
    L4000.state = 0
  ElseIf L5000.state = 1 Then
    L4000.state = 1
    L5000.state = 0
  ElseIf L6000.state = 1 Then
    L5000.state = 1
    L6000.state = 0
  ElseIf L7000.state = 1 Then
    L6000.state = 1
    L7000.state = 0
  ElseIf L8000.state = 1 Then
    L7000.state = 1
    L8000.state = 0
  ElseIf L9000.state = 1 Then
    L8000.state = 1
    L9000.state = 0
  ElseIf L10000.state = 1 Then
    L10000.state = 0
    L9000.state = 1
  End If
End Sub

'******************************************************
'           SCORING/CREDITS
'******************************************************

Dim ReelClickVol:ReelClickVol=0.5
Dim PrevScore

Sub AddScore(x)
  If TableTilted = 0 Then
    PrevScore = Score
    if isScoring() = 0 Then
      If SpecialLight.state = 1 Then SpecialLight.state = 0
      If x = 10 Then
        Play10Bell
        CheckForRoll(10)
        Score = Score + 10
        AdvMatch
      ElseIf x = 100 Then
        Play100Bell
        CheckForRoll(100)
        Score = Score + 100
      ElseIf x = 1000 Then
        Play100Bell
        CheckForRoll(1000)
        Score = Score + 1000
      ElseIf x = 10000 Then
        Play100Bell
        CheckForRoll(10000)
        Score = Score + 10000
      End If
    End If
    CheckFreeGame
  End If
  'If B2SOn Then Controller.B2SSetScore 1,Score
End Sub

Sub Play10Bell()
  If ChimesON = 1 Then
    PlaySound SoundFX("10_Point_Bell_Loud",DOFChimes), 1, 1, AudioPan(Spinner), 0,0,0,1,AudioFade(Spinner)
    DOF 153, 2
  Else
    PlaySoundAt "10_Point_Bell_Loud", Spinner
  End If
End Sub

Sub Play100Bell()
  If ChimesON = 1 Then
    PlaySound SoundFX("100_Point_Bell_Loud",DOFChimes), 1, 1, AudioPan(Spinner), 0,0,0,1,AudioFade(Spinner)
    DOF 154, 2
  Else
    PlaySoundAt "100_Point_Bell_Loud", Spinner
  End If
End Sub

Sub CheckForRoll(x)
  Select Case (x)
    Case 10:
      PulseReel5.enabled = 1

      Reel5Value = (Reel5Value + 1) mod 10
      If B2SOn Then Controller.b2ssetreel 5,Reel5Value

      PlaySoundAtVol "reelclick5", emp1r5, ReelClickVol
      if emp1r5.rotx = 261 then CheckForRoll(100)
    Case 100:
      PulseReel4.enabled = 1

      Reel4Value = (Reel4Value + 1) mod 10
      If B2SOn Then Controller.b2ssetreel 4,Reel4Value

      PlaySoundAtVol "reelclick4", emp1r4, ReelClickVol
      if emp1r4.rotx = 261 then CheckForRoll(1000)
    Case 1000:
      PulseReel3.enabled = 1

      Reel3Value = (Reel3Value + 1) mod 10
      If B2SOn Then Controller.b2ssetreel 3,Reel3Value

      PlaySoundAtVol "reelclick3", emp1r3, ReelClickVol
      if emp1r3.rotx = 261 then CheckForRoll(10000)
    Case 10000:
      PulseReel2.enabled = 1

      Reel2Value = (Reel2Value + 1) mod 10
      If B2SOn Then Controller.b2ssetreel 2,Reel2Value

      PlaySoundAtVol "reelclick2", emp1r2, ReelClickVol
      if emp1r2.rotx = 261 then CheckForRoll(100000)
    Case 100000:
      PulseReel1.enabled = 1

      Reel1Value = (Reel1Value + 1) mod 10
      If B2SOn Then Controller.b2ssetreel 1,Reel1Value

      PlaySoundAtVol "reelclick1", emp1r1, ReelClickVol
  End Select
End Sub

Sub AddCredit(direction)
  If PulseCreditReel.enabled = 0 Then
    if direction > 0 and credits >= 40 Then
      'do Nothing
    Else
      if direction = 1 Then
        playsound SoundFXDOF("knocker",124,DOFPulse,DOFKnocker)
      end if
      CreditDir = direction
      credits = credits + direction
      PulseCreditReel.enabled = 1
      PlaySoundAtVol "reelclick6", emcredits1, ReelClickVol
      DOF 135, 2
    End If
  End If
  If Credits > 0 Then
    DOF 125, 1
    Lcredit.state = 1
  Else
    DOF 125, 0
    Lcredit.state = 0
  End If
  If B2SOn Then Controller.B2SSetCredits Credits
End Sub

Sub CheckFreeGame()
  If PrevScore < Replay1 And Score >= Replay1 Then FreeGame
  If PrevScore < Replay2 And Score >= Replay2 Then FreeGame
  If PrevScore < Replay3 And Score >= Replay3 Then FreeGame
  If PrevScore < Replay4 And Score >= Replay4 Then FreeGame
End Sub

Sub FreeGame()
  'playsound SoundFXDOF("knocker",124,DOFPulse,DOFKnocker)  'Added to AddCredit Sub
  AddCredit(1)
End Sub

Dim mCountDown, mScore

Sub MultiScore(cycle, Score)
  mScore = Score
  mCountDown = cycle - 1
  AddScore(mscore)
  MultiScoreTimer.enabled = true
End Sub

Sub MultiScoreTimer_Timer
  mCountDown = mCountDown - 1
  If mCountDown < 0 Then
    DecreaseBonus
  Else
    AddScore(mscore)
  End If
  If mCountDown < 0.1 Then
    me.enabled = False
  End If
End Sub

Function isScoring()
  if PulseReel2.enabled = 0 and PulseReel3.enabled = 0 and PulseReel4.enabled = 0 and PulseReel5.enabled = 0 Then
    isScoring = 0
  Else
    isScoring = 1
  End If
End Function

'******************************************************
'             EMREELS
'******************************************************

Dim ReelStep:ReelStep = 6

Sub PulseReel1_Timer
  dim done:done=0
  emp1r1.rotx = emp1r1.rotx + ReelStep
  if emp1r1.rotx > 360 then emp1r1.rotx = emp1r1.rotx - 360

  Select Case (emp1r1.rotx)
    Case 9: done = 1
    Case 45: done = 1
    Case 81: done = 1
    Case 117: done = 1
    Case 153: done = 1
    Case 189: done = 1
    Case 225: done = 1
    Case 261: done = 1
    Case 297: done = 1
    Case 333: done = 1
  End Select

  if done = 1 then
    me.enabled = 0
  End If
End Sub

Sub PulseReel2_Timer
  dim done:done=0
  emp1r2.rotx = emp1r2.rotx + ReelStep
  if emp1r2.rotx > 360 then emp1r2.rotx = emp1r2.rotx - 360

  Select Case (emp1r2.rotx)
    Case 9: done = 1
    Case 45: done = 1
    Case 81: done = 1
    Case 117: done = 1
    Case 153: done = 1
    Case 189: done = 1
    Case 225: done = 1
    Case 261: done = 1
    Case 297: done = 1
    Case 333: done = 1
  End Select

  if done = 1 then
    me.enabled = 0
  End If
End Sub

Sub PulseReel3_Timer
  dim done:done=0
  emp1r3.rotx = emp1r3.rotx + ReelStep
  if emp1r3.rotx > 360 then emp1r3.rotx = emp1r3.rotx - 360

  Select Case (emp1r3.rotx)
    Case 9: done = 1
    Case 45: done = 1
    Case 81: done = 1
    Case 117: done = 1
    Case 153: done = 1
    Case 189: done = 1
    Case 225: done = 1
    Case 261: done = 1
    Case 297: done = 1
    Case 333: done = 1
  End Select

  if done = 1 then
    me.enabled = 0
  End If
End Sub

Sub PulseReel4_Timer
  dim done:done=0
  emp1r4.rotx = emp1r4.rotx + ReelStep
  if emp1r4.rotx > 360 then emp1r4.rotx = emp1r4.rotx - 360

  Select Case (emp1r4.rotx)
    Case 9: done = 1
    Case 45: done = 1
    Case 81: done = 1
    Case 117: done = 1
    Case 153: done = 1
    Case 189: done = 1
    Case 225: done = 1
    Case 261: done = 1
    Case 297: done = 1
    Case 333: done = 1
  End Select

  if done = 1 then
    me.enabled = 0
  End If
End Sub

Sub PulseReel5_Timer
  dim done:done=0
  emp1r5.rotx = emp1r5.rotx + ReelStep
  if emp1r5.rotx > 360 then emp1r5.rotx = emp1r5.rotx - 360

  Select Case (emp1r5.rotx)
    Case 9: done = 1
    Case 45: done = 1
    Case 81: done = 1
    Case 117: done = 1
    Case 153: done = 1
    Case 189: done = 1
    Case 225: done = 1
    Case 261: done = 1
    Case 297: done = 1
    Case 333: done = 1
  End Select

  if done = 1 then
    me.enabled = 0
  End If
End Sub

Dim CreditDir

Sub PulseCreditReel_Timer
  dim done:done=0

  emcredits2.rotx = emcredits2.rotx - ReelStep*CreditDir
  if emcredits2.rotx > 360 then emcredits2.rotx = emcredits2.rotx - 360
  if emcredits2.rotx < 0 then emcredits2.rotx = emcredits2.rotx + 360

  If CreditDir = 1 and emcredits2.rotx < 333 and emcredits2.rotx >= 297 then
    emcredits1.rotx = emcredits1.rotx - ReelStep*CreditDir
  Elseif CreditDir = -1 and emcredits2.rotx <= 333 and emcredits2.rotx > 297 then
    emcredits1.rotx = emcredits1.rotx - ReelStep*CreditDir
  End If

  Select Case (emcredits2.rotx)
    Case 9: done = 1
    Case 45: done = 1
    Case 81: done = 1
    Case 117: done = 1
    Case 153: done = 1
    Case 189: done = 1
    Case 225: done = 1
    Case 261: done = 1
    Case 297: done = 1
    Case 333: done = 1
  End Select

  if done = 1 then
    me.enabled = 0
  End If
End Sub

Sub ResetReels()
  If emp1r5.rotx <> 297 then PulseReel5.enabled = true:PlaySoundAtVol "reelclick5", emp1r5, ReelClickVol
  If emp1r4.rotx <> 297 then PulseReel4.enabled = true:PlaySoundAtVol "reelclick4", emp1r4, ReelClickVol
  If emp1r3.rotx <> 297 then PulseReel3.enabled = true:PlaySoundAtVol "reelclick3", emp1r3, ReelClickVol
  If emp1r2.rotx <> 297 then PulseReel2.enabled = true:PlaySoundAtVol "reelclick2", emp1r2, ReelClickVol
  If emp1r1.rotx <> 297 then PulseReel1.enabled = true:PlaySoundAtVol "reelclick1", emp1r1, ReelClickVol
  If B2SOn Then
    If Reel5Value <> 0 Then
      Reel5Value = (Reel5Value + 1) mod 10
      If B2SOn Then Controller.b2ssetreel 5,Reel5Value
    End If
    If Reel4Value <> 0 Then
      Reel4Value = (Reel4Value + 1) mod 10
      If B2SOn Then Controller.b2ssetreel 4,Reel4Value
    End If
    If Reel3Value <> 0 Then
      Reel3Value = (Reel3Value + 1) mod 10
      If B2SOn Then Controller.b2ssetreel 3,Reel3Value
    End If
    If Reel2Value <> 0 Then
      Reel2Value = (Reel2Value + 1) mod 10
      If B2SOn Then Controller.b2ssetreel 2,Reel2Value
    End If
    If Reel1Value <> 0 Then
      Reel1Value = (Reel1Value + 1) mod 10
      If B2SOn Then Controller.b2ssetreel 1,Reel1Value
    End If
  End If
End Sub

Sub SetReels()
  dim digangles(10)
  digangles(0) = 297
  digangles(1) = 333
  digangles(2) = 9
  digangles(3) = 45
  digangles(4) = 81
  digangles(5) = 117
  digangles(6) = 153
  digangles(7) = 189
  digangles(8) = 225
  digangles(9) = 261

  If Score > 9 Then emp1r5.rotx = digangles(Int(Mid(StrReverse(score),2,1))):Reel5Value=Mid(StrReverse(score),2,1)
  If Score > 99 Then emp1r4.rotx = digangles(Int(Mid(StrReverse(score),3,1)))::Reel4Value=Mid(StrReverse(score),3,1)
  If Score > 999 Then emp1r3.rotx = digangles(Int(Mid(StrReverse(score),4,1))):Reel3Value=Mid(StrReverse(score),4,1)
  If Score > 9999 Then emp1r2.rotx = digangles(Int(Mid(StrReverse(score),5,1))):Reel2Value=Mid(StrReverse(score),5,1)
  If Score > 99999 Then emp1r1.rotx = digangles(Int(Mid(StrReverse(score),6,1))):Reel1Value=Mid(StrReverse(score),6,1)

  digangles(0) = 297
  digangles(9) = 333
  digangles(8) = 9
  digangles(7) = 45
  digangles(6) = 81
  digangles(5) = 117
  digangles(4) = 153
  digangles(3) = 189
  digangles(2) = 225
  digangles(1) = 261

  If Credits > 9 Then emcredits1.rotx = digangles(Int(Mid(StrReverse(Credits),2,1)))
  emcredits2.rotx = digangles(Int(Mid(StrReverse(Credits),1,1)))

End Sub

'******************************************************
'             DRAIN
'******************************************************

Sub Drain_Hit()
  PlaySound "drain",0,1,AudioPan(Drain),0.25,0,0,1,AudioFade(Drain)
  PlaySoundAtVol "MotorRunning", Spinner, 0.25
  DOF 121, DOFPulse
  DownPosts
  YellowLightsOff
  GreenLightsOff
  CloseWestGate
  If TableTilted = 1 Then
    BackglassTilt.visible = 0
    TableTilted = 0
    ResetTilt
  End If

  If SpecialLight.State = 0 Then
    SubtractBall
  Else
    Drain.timerenabled = 1
  End If
End Sub

Sub Drain_Timer()
  SpecialLight.state = 1
  Drain.Kick 60, 20
  DOF 136, DOFPulse
  PlaySound SoundFX("ballrelease",DOFContactors), 0,1,AudioPan(Drain),0.25,0,0,1,AudioFade(Drain)
  me.timerenabled = False
End Sub

Sub AddBall()
  BallInPlay = BallInPlay + 1
  If BallInPlay > 10 Then BallInPlay=10
  PlayQuietClick
  BackglassBall.image = "JiveTimeBall" & BallInPlay
  If B2SOn Then Controller.B2SSetBallInPlay BallInPlay
End Sub

Sub SubtractBall()
  BallInPlay = BallInPlay - 1
  if ballinplay < 0 then BallInPlay = 0
  If BallInPlay = 0 Then
    InProgress = False
    BackglassBall.image = "JiveTimeGameOver"
    If B2SOn Then Controller.B2SSetGameOver 1
    CheckMatch
    LeftFlipper.RotateToStart
    StopSound "buzzL"
    DOF 101, 0

    RightFlipper.RotateToStart
    StopSound "buzz"
    DOF 102, 0

    If Score >= HighScore Then
      HighScore = Score
      HighScoreEntryInit()
    End If
  Else
    BackglassBall.image = "JiveTimeBall" & BallInPlay
    Drain.timerenabled = 1
  End If
  PlayQuietClick
  If B2SOn Then Controller.B2SSetBallInPlay BallInPlay
End Sub


'******************************************************
'           START NEW GAME
'******************************************************

Dim NGCount

Sub StartNewGame_Timer()
  NGCount = NGCount + 1

  Select Case (NGCount)
    Case 1:
      BackglassMatch.visible = False
      BackglassTilt.visible = False
      If B2SOn Then
        Controller.B2SSetTilt 0
        Controller.B2SSetMatch 0
        'Controller.B2SSetScore 1,0
        Controller.B2SSetGameOver 0
      End If
      ResetTilt
      PlaySoundAtVol "MotorRunning", Spinner, 0.25
      PlayLoudClick
      If BallInPlay < BallsPerGame Then AddBall
      ResetReels
    Case 2:
      PlayLoudClick
      If BallInPlay < BallsPerGame Then AddBall
      ResetReels
    Case 3:
      PlayLoudClick
      If BallInPlay < BallsPerGame Then AddBall
      ResetReels
    Case 4:
      PlayLoudClick
      If BallInPlay < BallsPerGame Then AddBall
      ResetReels
    Case 5:
      PlayLoudClick
      If BallInPlay < BallsPerGame Then AddBall
      ResetReels

    Case 7:
      PlayQuietClick
      If BallInPlay < BallsPerGame Then AddBall
      ResetReels
    Case 8:
      PlayQuietClick
      If BallInPlay < BallsPerGame Then AddBall
      ResetReels
    Case 9:
      PlayQuietClick
      If BallInPlay < BallsPerGame Then AddBall
      ResetReels
    Case 10:
      PlayQuietClick
      If BallInPlay < BallsPerGame Then AddBall
      ResetReels
    Case 11:
      PlayQuietClick
      If BallInPlay < BallsPerGame Then AddBall
      ResetReels

    Case 13:
      Drain.timerenabled = true
      Score = 0
      SetReels
      If B2SOn Then
        Controller.b2ssetreel 1,0:Reel1Value = 0
        Controller.b2ssetreel 2,0:Reel2Value = 0
        Controller.b2ssetreel 3,0:Reel3Value = 0
        Controller.b2ssetreel 4,0:Reel4Value = 0
        Controller.b2ssetreel 5,0:Reel5Value = 0
      End If
      me.enabled = false
      NGCount = 0
  End Select
End Sub

'******************************************************
'               MATCH
'******************************************************

Sub CheckMatch()
  If Match = 0 Then
    BackglassMatch.image = "JiveTimeMatch00"
  Else
    BackglassMatch.image = "JiveTimeMatch" & Match
  End If

  BackglassMatch.visible = 1
  SetB2SMatch
  if Match=(Score mod 100) then
    FreeGame
  end if
End Sub

Sub AdvMatch()
  Select Case (Match)
    Case 30: Match = 80
    Case 80: Match = 20
    Case 20: Match = 50
    Case 50: Match = 90
    Case 90: Match = 40
    Case 40: Match = 0
    Case 0: Match = 60
    Case 60: Match = 10
    Case 10: Match = 70
    Case 70: Match = 30
  End Select
End Sub

Sub SetB2SMatch()
  If B2SOn Then
    If Match = 0 Then
      Controller.B2SSetMatch 100
    Else
      Controller.B2SSetMatch Match
    End If
  End If
End Sub

'******************************************************
'         SLING AND ANIMATED RUBBERS
'******************************************************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
  AddScore(10)
    PlaySound SoundFX("right_slingshot",DOFContactors), 0,1, 0.05,0.05 '0,1, AudioPan(RightSlingShot), 0.05,0,0,1,AudioFade(RightSlingShot)
  DOF 104, 2
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
  'gi1.State = 0:Gi2.State = 0
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0:'gi1.State = 1:Gi2.State = 1
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  AddScore(10)
  PlaySound SoundFX("left_slingshot",DOFContactors), 0,1, -0.05,0.05 '0,1, AudioPan(LeftSlingShot), 0.05,0,0,1,AudioFade(LeftSlingShot)
  DOF 103, 2
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
  'gi3.State = 0:Gi4.State = 0
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0:'gi3.State = 1:Gi4.State = 1
    End Select
    LStep = LStep + 1
End Sub

Sub WallLL_Hit
  RubberLL1.Visible = 1:RubberLL.Visible = 0:WallLL.TimerEnabled = 1
End Sub

Sub WallLL_Timer
  RubberLL1.Visible = 0:RubberLL.Visible = 1:me.TimerEnabled = 0
End Sub

Sub WallRL_Hit
  RubberRL1.Visible = 1:RubberRL.Visible = 0:WallRL.TimerEnabled = 1
End Sub

Sub WallRL_Timer
  RubberRL1.Visible = 0:RubberRL.Visible = 1:me.TimerEnabled = 0
End Sub

Sub WallRLM_Hit
  RubberRLM1.Visible = 1:RubberRLM.Visible = 0:WallRLM.TimerEnabled = 1
End Sub

Sub WallRLM_Timer
  RubberRLM1.Visible = 0:RubberRLM.Visible = 1:me.TimerEnabled = 0
End Sub


Sub T10Points1_Timer
  RubberLM1.Visible = 0:RubberLM.Visible = 1:me.TimerEnabled = 0
End Sub

Sub T10Points2_Timer
  RubberLU1.Visible = 0:RubberLU.Visible = 1:me.TimerEnabled = 0
End Sub

Sub T10Points4_Timer
  RubberRUM1.Visible = 0:RubberRUM.Visible = 1:me.TimerEnabled = 0
End Sub

Sub T10Points5_Timer
  RubberRU1.Visible = 0:RubberRU.Visible = 1:me.TimerEnabled = 0
End Sub

'******************************************************
'           LOAD & SAVE TABLE
'******************************************************
sub savehs
    savevalue "JiveTime_VPX", "Credits", Credits
  savevalue "JiveTime_VPX", "Match", Match
  savevalue "JiveTime_VPX", "SpinnerAngle", SpinnerAngle
  savevalue "JiveTime_VPX", "Bonus", Bonus
  savevalue "JiveTime_VPX", "HighScore", Highscore
  savevalue "JiveTime_VPX", "HSA1", HSA1
  savevalue "JiveTime_VPX", "HSA2", HSA2
  savevalue "JiveTime_VPX", "HSA3", HSA3
  savevalue "JiveTime_VPX", "BallInPlay", BallInPlay
  savevalue "JiveTime_VPX", "Score", Score
  savevalue "JiveTime_VPX", "SpecialLight", SpecialLight.state
end sub

sub loadhs
  HighScore=0
  Credits=0
  Match=0
  SpinnerAngle = 0

    dim temp
  temp = LoadValue("JiveTime_VPX", "Credits")
    If (temp <> "") then Credits = CDbl(temp)
    temp = LoadValue("JiveTime_VPX", "Match")
    If (temp <> "") then Match = CDbl(temp)
    temp = LoadValue("JiveTime_VPX", "SpinnerAngle")
    If (temp <> "") then SpinnerAngle = CDbl(temp)
    temp = LoadValue("JiveTime_VPX", "Bonus")
    If (temp <> "") then Bonus = CDbl(temp)
    temp = LoadValue("JiveTime_VPX", "HighScore")
    If (temp <> "") then HighScore = CDbl(temp)
    temp = LoadValue("JiveTime_VPX", "hsa1")
    If (temp <> "") then HSA1 = CDbl(temp)
    temp = LoadValue("JiveTime_VPX", "hsa2")
    If (temp <> "") then HSA2 = CDbl(temp)
    temp = LoadValue("JiveTime_VPX", "hsa3")
    If (temp <> "") then HSA3 = CDbl(temp)
    temp = LoadValue("JiveTime_VPX", "BallInPlay")
    If (temp <> "") then BallInPlay = CDbl(temp)
    temp = LoadValue("JiveTime_VPX", "Score")
    If (temp <> "") then Score = CDbl(temp)
    temp = LoadValue("JiveTime_VPX", "SpecialLight")
    If (temp <> "") then SpecialLightState = CDbl(temp)
  SetReels

  if HighScore=0 then HighScore=50000
end sub

'******************************************************
'               TILT
'******************************************************

Dim TiltSens

Sub CheckTilt
  If Tilttimer.Enabled = True Then
   TiltSens = TiltSens + 1
   if TiltSens = 3 Then GameTilted
  Else
   TiltSens = 1
   Tilttimer.Enabled = True
  End If
End Sub

Sub Tilttimer_Timer()
  If TiltSens > 0 Then TiltSens = TiltSens - 1
  If TiltSens = 0 Then Tilttimer.Enabled = False
End Sub

Sub GameTilted
  TableTilted = 1
  CloseWestGate
  SpecialLight.state=0
  GreenLightsOff
  YellowLightsOff
  DownPosts
  DOF 134, 2

  Bumper1.threshold = 100
  Bumper2.threshold = 100
  Bumper3.threshold = 100
  Bumper4.threshold = 100
  Bumper5.threshold = 100

  LeftSlingShot.SlingShotThreshold = 100
  RightSlingShot.SlingShotThreshold = 100


  PlayLoudClick
  BackglassTilt.visible = true
  If B2SOn Then Controller.B2SSetTilt 1


  LeftFlipper.RotateToStart
  StopSound "buzzL"
  DOF 101, 0

  RightFlipper.RotateToStart
  StopSound "buzz"
  DOF 102, 0
End Sub

Sub ResetTilt
  'DOF 134, 2
  TableTilted = 0

  Bumper1.threshold = 2
  Bumper2.threshold = 2
  Bumper3.threshold = 2
  Bumper4.threshold = 2
  Bumper5.threshold = 2

  LeftSlingShot.SlingShotThreshold = 1.5
  RightSlingShot.SlingShotThreshold = 1.5
  If B2SOn Then Controller.B2SSetTilt 0
End Sub

'******************************************************
'           REALTIME UPDATES
'******************************************************


Sub RealtimeUpdates_Timer()
  GateA.RotX = Gate1.CurrentAngle*0.6

  'West Gate
  diverter.roty = DiverterFlipper.currentangle + 90
  If DiverterFlipper.currentangle < -30 Then
    WestGateLight.state=1
  Else
    If WestGateLight.state=1 Then
      CloseWestGate
      WestGateLight.state=0
    End If
  End If

  'Spinner
  spinner.rotz = SpinnerAngle + 90
  RollingUpdate

' ********* Kicker Code
  If InRect(KickerBall.x,KickerBall.y,(TopKicker.x-35),(TopKicker.y-35),(TopKicker.x+35),(TopKicker.y-35),(TopKicker.x+35),(TopKicker.y+35),(TopKicker.x-35),(TopKicker.y+35)) Then
    If KickerBall.z < 65 Then
      TopKicker.enabled = True
    End If
  End If
' ********* End Kicker Code

' ********* Kicker Code
  If InRect(KickerBall.x,KickerBall.y,(MiddleKicker.x-35),(MiddleKicker.y-35),(MiddleKicker.x+35),(MiddleKicker.y-35),(MiddleKicker.x+35),(MiddleKicker.y+35),(MiddleKicker.x-35),(MiddleKicker.y+35)) Then
    If KickerBall.z < 65 Then
      MiddleKicker.enabled = True
    End If
  End If
' ********* End Kicker Code


  If Table1.ShowDT = True  and Table1.ShowFSS = False then
    emp1r7.rotx = emp1r1.rotx
    emp1r8.rotx = emp1r2.rotx
    emp1r9.rotx = emp1r3.rotx
    emp1r10.rotx = emp1r4.rotx
    emp1r11.rotx = emp1r5.rotx
    emcredits6.rotx = emcredits1.rotx
    emcredits5.rotx = emcredits2.rotx

    BackglassMatch1.visible = BackglassMatch.visible
    BackglassTilt1.visible = BackglassTilt.visible

    BackglassMatch1.image = BackglassMatch.image
    BackglassBall1.image = BackglassBall.image

    spinner1.rotz = SpinnerAngle + 90
  End If
End Sub

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
'   rothbauerw's Manual Ball Control
'*****************************************

Dim BCup, BCdown, BCleft, BCright
Dim ControlBallInPlay, ControlActiveBall
Dim BCvel, BCyveloffset, BCboostmulti, BCboost

BCboost = 1       'Do Not Change - default setting
BCvel = 4       'Controls the speed of the ball movement
BCyveloffset = -0.01  'Offsets the force of gravity to keep the ball from drifting vertically on the table, should be negative
BCboostmulti = 3    'Boost multiplier to ball veloctiy (toggled with the B key)

ControlBallInPlay = false

Sub StartBallControl_Hit()
  Set ControlActiveBall = ActiveBall
  ControlBallInPlay = true
End Sub

Sub StopBallControl_Hit()
  ControlBallInPlay = false
End Sub

Sub BallControlTimer_Timer()
  If EnableBallControl and ControlBallInPlay then
    If BCright = 1 Then
      ControlActiveBall.velx =  BCvel*BCboost
    ElseIf BCleft = 1 Then
      ControlActiveBall.velx = -BCvel*BCboost
    Else
      ControlActiveBall.velx = 0
    End If

    If BCup = 1 Then
      ControlActiveBall.vely = -BCvel*BCboost
    ElseIf BCdown = 1 Then
      ControlActiveBall.vely =  BCvel*BCboost
    Else
      ControlActiveBall.vely = bcyveloffset
    End If
  End If
End Sub


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

Sub RollingUpdate()
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
        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 80 Then
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b))*RollingSoundFactor, AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
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


'*****************************************
' ninuzzu's FLIPPER SHADOWS
'*****************************************

sub FlipperTimer_Timer()
  FlipperLSh.RotZ = LeftFlipper.currentangle
  FlipperRSh.RotZ = RightFlipper.currentangle

End Sub

'*****************************************
' ninuzzu's BALL SHADOW
'*****************************************
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
            BallShadow(b).X = ((BOT(b).X) + ((BOT(b).X - (Table1.Width/2))/40))
        Else
            BallShadow(b).X = ((BOT(b).X) + ((BOT(b).X - (Table1.Width/2))/40))
        End If
        ballShadow(b).Y = BOT(b).Y + 10
        If BOT(b).Z > 70 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
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
  PlaySound "pinhit_low", 0, Vol(ActiveBall)*VolumeDial, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Woods_Hit (idx)
  PlaySound "woodhit", 0, Vol(ActiveBall)*VolumeDial*2, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
  PlaySound "target", 0, Vol(ActiveBall)*VolumeDial, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
  PlaySound "metalhit_thin", 0, Vol(ActiveBall)*VolumeDial, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
  PlaySound "metalhit_medium", 0, Vol(ActiveBall)*VolumeDial, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
  PlaySound "metalhit2", 0, Vol(ActiveBall)*VolumeDial, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
  PlaySound "gate4", 0, Vol(ActiveBall)*VolumeDial, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinner_Spin
  PlaySound "fx_spinner", 0, .25*VolumeDial, AudioPan(Spinner), 0.25, 0, 0, 1, AudioFade(Spinner)
End Sub

Sub Rubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 20 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolumeDial, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End Sub

Sub Posts_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolumeDial, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    RandomSoundRubber()
  End If
End Sub

Sub RandomSoundRubber()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*VolumeDial, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall)*VolumeDial, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall)*VolumeDial, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

Sub LeftFlipper_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RightFlipper_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub DiverterFlipper_Collide(parm)
  PlaySound "metalhit2", 0, Vol(ActiveBall)*VolumeDial, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub RandomSoundFlipper()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall)*VolumeDial, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall)*VolumeDial, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall)*VolumeDial, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

'**************************************************************************
'                 Positional Sound Playback Functions by DJRobX
'**************************************************************************

Sub PlaySoundAtVol(sound, tableobj, Vol)
    PlaySound sound, 1, Vol, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'Set position at table object, vol, and loops manually.

Sub PlaySoundAtVolLoops(sound, tableobj, Vol, Loops)
    PlaySound sound, Loops, Vol, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub


Sub PlayLoudClick()
  PlaySoundAtVol "Motor_Click_Loud_Long", Spinner, 0.05
End Sub

Sub PlayQuietClick()
  PlaySoundAtVol "Motor_Click_Quiet_Long2", Spinner, 0.05
End Sub


'==========================================================================================================================================
'============================================================= START OF HIGH SCORES ROUTINES =============================================================
'==========================================================================================================================================
'
'ADD LINE TO TABLE_KEYDOWN SUB WITH THE FOLLOWING:    If HSEnterMode Then HighScoreProcessKey(keycode) AFTER THE STARTGAME ENTRY
'ADD: And Not HSEnterMode=true TO IF KEYCODE=STARTGAMEKEY
'TO SHOW THE SCORE ON POST-IT ADD LINE AT RELEVENT LOCATION THAT HAS:  UpdatePostIt
'TO INITIATE ADDING INITIALS ADD LINE AT RELEVENT LOCATION THAT HAS:  HighScoreEntryInit()
'ADD THE FOLLOWING LINES TO TABLE_INIT TO SETUP POSTIT
' if HSA1="" then HSA1=25
' if HSA2="" then HSA2=25
' if HSA3="" then HSA3=25
' UpdatePostIt
'ADD HSA1, HSA2 AND HSA3 TO SAVE AND LOAD VALUES FOR TABLE
'ADD A TIMER NAMED HighScoreFlashTimer WITH INTERVAL 100 TO TABLE
'SET HSSSCOREX BELOW TO WHATEVER VARIABLE YOU USE FOR HIGH SCORE.
'ADD OBJECTS TO PLAYFIELD (EASIEST TO JUST COPY FROM THIS TABLE)
'IMPORT POST-IT IMAGES


Dim HSA1, HSA2, HSA3
Dim HSEnterMode, hsLetterFlash, hsEnteredDigits(3), hsCurrentDigit, hsCurrentLetter
Dim HSArray
Dim HSScoreM,HSScore100k, HSScore10k, HSScoreK, HSScore100, HSScore10, HSScore1, HSScorex 'Define 6 different score values for each reel to use
HSArray = Array("Postit0","postit1","postit2","postit3","postit4","postit5","postit6","postit7","postit8","postit9","postitBL","postitCM","Tape")
Const hsFlashDelay = 4

' ***********************************************************
'  HiScore DISPLAY
' ***********************************************************

Sub UpdatePostIt
  dim tempscore
  HSScorex = HighScore
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
' if showhisc=1 and showhiscnames=1 then
'   for each objekt in hiscname:objekt.visible=1:next
    HSName1.image = ImgFromCode(HSA1, 1)
    HSName2.image = ImgFromCode(HSA2, 2)
    HSName3.image = ImgFromCode(HSA3, 3)
'   else
'   for each objekt in hiscname:objekt.visible=0:next
' end if
End Sub

Function ImgFromCode(code, digit)
  Dim Image
  if (HighScoreFlashTimer.Enabled = True and hsLetterFlash = 1 and digit = hsCurrentLetter) then
    Image = "postitBL"
  elseif (code + ASC("A") - 1) >= ASC("A") and (code + ASC("A") - 1) <= ASC("Z") then
    Image = "postit" & chr(code + ASC("A") - 1)
  elseif code = 27 Then
    Image = "PostitLT"
    elseif code = 0 Then
    image = "PostitSP"
    Else
      msgbox("Unknown display code: " & code)
  end if
  ImgFromCode = Image
End Function

Sub HighScoreEntryInit()
  HSA1=0:HSA2=0:HSA3=0
  HSEnterMode = True
  hsCurrentDigit = 0
  hsCurrentLetter = 1:HSA1=1
  HighScoreFlashTimer.Interval = 250
  HighScoreFlashTimer.Enabled = True
  hsLetterFlash = hsFlashDelay
End Sub

Sub HighScoreFlashTimer_Timer()
  hsLetterFlash = hsLetterFlash-1
  UpdatePostIt
  If hsLetterFlash=0 then 'switch back
    hsLetterFlash = hsFlashDelay
  end if
End Sub


' ***********************************************************
'  HiScore ENTER INITIALS
' ***********************************************************

Sub HighScoreProcessKey(keycode)
    If keycode = LeftFlipperKey Then
    hsLetterFlash = hsFlashDelay
    Select Case hsCurrentLetter
      Case 1:
        HSA1=HSA1-1:If HSA1=-1 Then HSA1=26 'no backspace on 1st digit
        UpdatePostIt
      Case 2:
        HSA2=HSA2-1:If HSA2=-1 Then HSA2=27
        UpdatePostIt
      Case 3:
        HSA3=HSA3-1:If HSA3=-1 Then HSA3=27
        UpdatePostIt
     End Select
    End If

  If keycode = RightFlipperKey Then
    hsLetterFlash = hsFlashDelay
    Select Case hsCurrentLetter
      Case 1:
        HSA1=HSA1+1:If HSA1>26 Then HSA1=0
        UpdatePostIt
      Case 2:
        HSA2=HSA2+1:If HSA2>27 Then HSA2=0
        UpdatePostIt
      Case 3:
        HSA3=HSA3+1:If HSA3>27 Then HSA3=0
        UpdatePostIt
     End Select
  End If

    If keycode = StartGameKey Then
    Select Case hsCurrentLetter
      Case 1:
        hsCurrentLetter=2 'ok to advance
        HSA2=HSA1 'start at same alphabet spot
'       EMReelHSName1.SetValue HSA1:EMReelHSName2.SetValue HSA2
      Case 2:
        If HSA2=27 Then 'bksp
          HSA2=0
          hsCurrentLetter=1
        Else
          hsCurrentLetter=3 'enter it
          HSA3=HSA2 'start at same alphabet spot
        End If
      Case 3:
        If HSA3=27 Then 'bksp
          HSA3=0
          hsCurrentLetter=2
        Else
          savehs 'enter it
          HighScoreFlashTimer.Enabled = False
          HSEnterMode = False
        End If
    End Select
    UpdatePostIt
    End If
End Sub


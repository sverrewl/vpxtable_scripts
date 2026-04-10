'*************************************
'Space Time (Bally 1972) - IPD No. 2262
'************************************

Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the Controller.vbs file in order to run this table (installed with the VPX package in the scripts folder)"
On Error Goto 0


Const BallSize = 25
Const BallMass = 1

Const cGameName = "spacetime_1972"


'******************************************************
'             OPTIONS
'******************************************************

Const BallsPerGame = 5        '3 or 5
Const VolumeDial = 2        'Change volume of hit events
Const RollingSoundFactor = 5    'set sound level factor here for Ball Rolling Sound, 1=default level

Dim EnableBallControl
EnableBallControl = false 'Change to true to enable manual ball control (or press C in-game) via the arrow keys and B (boost movement) keys

'******************************************************
'             VARIABLES
'******************************************************

Dim i
Dim obj
Dim InProgress
Dim BallInPlay
Dim MaxPlayers:MaxPlayers = 4
Dim Players, Player, HighPlayer
Dim PrevScore(4)
Dim Score(4)
Dim HighScore
Dim Initials
Dim Credits
Dim Match
Dim Bonus
Dim Replay1:Replay1 = 63000
Dim Replay2:Replay2 = 77000
Dim Replay3:Replay3 = 91000
Dim Replay4:Replay4 = 250000
Dim TableTilted
Dim TiltCount
Dim GameOn:GameOn = 0
Dim Reel1Value, Reel2Value, Reel3Value, Reel4Value, Reel5Value

'Hide rails
Dim DesktopMode: DesktopMode = Table1.ShowDT

If DesktopMode = True Then 'Show Desktop components
Ramp16.image="left_rail_brighter"
Ramp15.image="right_rail_brighter"
Else
Ramp16.image="bg_fs"
Ramp15.image="bg_fs"
End if

'dt text
    mtchtxt.text = "match"
    BIPtext.text = "ball in play"
    CPtext.text = "can play"
    credittxt.text = "credits"

'******************************************************
'             TABLE INIT
'******************************************************

Sub Table1_Init
  LoadEM

  loadhs

  if HSA1="" then HSA1=23
  if HSA2="" then HSA2=10
  if HSA3="" then HSA3=18
  UpdatePostIt

  Drain.CreateSizedballWithMass Ballsize,Ballmass

  TableTilted=false
  InProgress=false

  Post.collidable = false
  gatetopright.collidable=True
  gatebottomright.collidable=True

    If B2SOn or Table1.ShowDT = False then
        for each i in DT : i.visible = 0 : next
    End If

  If Credits > 0 Then DOF 125, 1
End Sub

Sub Table1_Exit()
  savehs
  DOF 101, 0
  DOF 102, 0
  DOF 123, 0
  DOF 125, 0
  DOF 150, 0
  DOF 151, 0
  DOF 152, 0
  DOF 153, 0
  DOF 154, 0
  DOF 155, 0
  DOF 156, 0
  DOF 157, 0
  DOF 158, 0
  DOF 159, 0
  DOF 160, 0

  If B2SOn Then Controller.Stop
End Sub


Dim BootCount:BootCount = 0

Sub BootTable_Timer()
  If BootCount = 0 Then
    BootCount = 1
    me.interval = 100
    PlaySoundAt "poweron", Plunger
  Else

    If match = 0 Then
      Matchtxt.text = "00"
    Else
      Matchtxt.text = Match
    End If

    gamov.text = "GAME OVER"

    SetB2SMatch
    If B2SOn Then
      Controller.B2SSetGameOver 1
      Controller.B2SSetData 80,1
    End If

        ScoreReel1.setvalue score(1)
    ScoreReel2.setvalue score(2)
    ScoreReel3.setvalue score(3)
    ScoreReel4.setvalue score(4)
    CreditReel.setvalue Credits

    '*****GI Lights On
    dim xx
    For each xx in GI:xx.State = 1: Next
    rightrollover_prim.blenddisablelighting = 0.24
    leftrollover_prim.blenddisablelighting = 0.24
    woodrails_prim.image="woodrails"
    woodrails_prim.blenddisablelighting = 0.5
    rmetalguidetop_prim.blenddisablelighting = .5
    rmetalguidetop_prim.image = "uppermetal"
    rmetalguideleft_prim.blenddisablelighting = 0.3
    rmetalguidebottom_prim.blenddisablelighting = 0.5
    rmetalguidebottom_prim.image = "lowermetalguideGION"
    outers_prim.image="outers"
    outers_prim.blenddisablelighting = 0.5
    lgate_prim.image="leftgate"
    rgate_prim.image="rightgate"
    tunnel_prim.image = "tunnel1000"
    plasticedges_prim.image="plasticedgeslit"
    plasticedgeslower_prim.image="clearguide"
    plasticedgeslower_prim.blenddisablelighting = 1
    plastics_prim.image="plastics_GION"
    top_apron_prim.image="topapron"
    top_metal_prim.image="topmetal"
    bumpcap1.image="bump1off"
    bumpcap2.image="bump1off"
    shadows.visible=1
    shadowsoff.visible=0
    GameOn = 1

    me.enabled = False
  End If
End Sub

Sub BootB2S_Timer()
  If B2SOn Then
    Controller.B2SSetCredits Credits
    Controller.B2SSetScorePlayer 1,Score(1)
    Controller.B2SSetScorePlayer 2,Score(2)
    Controller.B2SSetScorePlayer 3,Score(3)
    Controller.B2SSetScorePlayer 4,Score(4)
  End If
  me.enabled = False
End Sub

'******************************************************
'             KEYS
'******************************************************

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
    playsound "click"
    InProgress=true
    StartNewGame.enabled = True
    Players = 1
    Player = 1
    BallInPlay = 1

    canplay.text = Players
    If B2SOn Then
      Controller.B2SSetCanPlay 31,Players
      Controller.B2SSetData 40,1
    End If
   elseif keycode = StartGameKey and Credits>0 and InProgress=true and Players>0 and Players<MaxPlayers and BallInPlay<2 then
    AddCredit(-1)
    Players=Players+1
    playsound "click"

    canplay.text = Players
    If B2SOn Then
      Controller.B2SSetCanPlay 31,Players
    End If
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
'           RUBBERS/SLINGS
'******************************************************

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep, LA, LB, LC, LD, LE, RA, RB, RC, RD, RE


Sub RightSlingShot_Slingshot
  AddScore(10)
  ToggleBumpers
    PlaySound SoundFX("right_slingshot",DOFContactors), 0,1, AudioPan(Sling1), 0.05,0,0,1,AudioFade(Sling1)
  DOF 104, 2
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.Rotx = 20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.Rotx = 8
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.Rotx = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  AddScore(10)
  ToggleBumpers
    PlaySound SoundFX("left_slingshot",DOFContactors),0,1, AudioPan(Sling2), 0.05,0,0,1,AudioFade(Sling2)
  DOF 103, 2
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.Rotx = 20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.Rotx = 8
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.Rotx = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub WallLA_Hit
  AddScore(10)
  ToggleBumpers
    RubberLA.Visible = 0
    RubberLA1.Visible = 1
    LA = 0
    WallLAtimer.Enabled = 1
End Sub

Sub WallLAtimer_Timer
    Select Case LA
        Case 3:RubberLA1.Visible = 0:RubberLA2.Visible = 1
        Case 4:RubberLA2.Visible = 0:RubberLA.Visible = 1:WallLAtimer.Enabled = 0:
    End Select
    LA = LA + 1
End Sub

Sub WallLB_Hit
  AddScore(10)
  ToggleBumpers
    RubberLB.Visible = 0
    RubberLB1.Visible = 1
    LB = 0
    WallLBtimer.Enabled = 1
End Sub

Sub WallLBtimer_Timer
    Select Case LB
        Case 3:RubberLB1.Visible = 0:RubberLB2.Visible = 1
        Case 4:RubberLB2.Visible = 0:RubberLB.Visible = 1:WallLBtimer.Enabled = 0:
    End Select
    LB = LB + 1
End Sub

Sub WallLC_Hit
  AddScore(10)
  ToggleBumpers
    RubberLC.Visible = 0
    RubberLC1.Visible = 1
    LC = 0
    WallLCtimer.Enabled = 1
End Sub

Sub WallLCtimer_Timer
    Select Case LC
        Case 3:RubberLC1.Visible = 0:RubberLC2.Visible = 1
        Case 4:RubberLC2.Visible = 0:RubberLC.Visible = 1:WallLCtimer.Enabled = 0:
    End Select
    LC = LC + 1
End Sub

Sub WallLD_Hit
  AddScore(10)
  ToggleBumpers
    RubberLD.Visible = 0
    RubberLD1.Visible = 1
    LD = 0
    WallLDtimer.Enabled = 1
End Sub

Sub WallLDtimer_Timer
    Select Case LD
        Case 3:RubberLD1.Visible = 0:RubberLD2.Visible = 1
        Case 4:RubberLD2.Visible = 0:RubberLD.Visible = 1:WallLDtimer.Enabled = 0:
    End Select
    LD = LD + 1
End Sub

Sub WallLE_Hit
    RubberLE.Visible = 0
    RubberLE1.Visible = 1
    LE = 0
    WallLEtimer.Enabled = 1
End Sub

Sub WallLEtimer_Timer
    Select Case LE
        Case 3:RubberLE1.Visible = 0:RubberLE2.Visible = 1
        Case 4:RubberLE2.Visible = 0:RubberLE.Visible = 1:WallLEtimer.Enabled = 0:
    End Select
    LE = LE + 1
End Sub

Sub WallRB_Hit
    RubberRB.Visible = 0
    RubberRB1.Visible = 1
    RB = 0
    WallRBtimer.Enabled = 1
End Sub

Sub WallRBtimer_Timer
    Select Case RB
        Case 3:RubberRB1.Visible = 0:RubberRB2.Visible = 1
        Case 4:RubberRB2.Visible = 0:RubberRB.Visible = 1:WallRBtimer.Enabled = 0:
    End Select
    RB = RB + 1
End Sub

Sub WallRC_Hit
  AddScore(10)
  ToggleBumpers
    RubberRC.Visible = 0
    RubberRC1.Visible = 1
    RC = 0
    WallRCtimer.Enabled = 1
End Sub

Sub WallRCtimer_Timer
    Select Case RC
        Case 3:RubberRC1.Visible = 0:RubberRC2.Visible = 1
        Case 4:RubberRC2.Visible = 0:RubberRC.Visible = 1:WallRCtimer.Enabled = 0:
    End Select
    RC = RC + 1
End Sub

Sub WallRD_Hit
  AddScore(10)
  ToggleBumpers
    RubberRD.Visible = 0
    RubberRD1.Visible = 1
    RD = 0
    WallRDtimer.Enabled = 1
End Sub

Sub WallRDtimer_Timer
    Select Case RD
        Case 3:RubberRD1.Visible = 0:RubberRD2.Visible = 1
        Case 4:RubberRD2.Visible = 0:RubberRD.Visible = 1:WallRDtimer.Enabled = 0:
    End Select
    RD = RD + 1
End Sub

Sub WallRE_Hit
    RubberRE.Visible = 0
    RubberRE1.Visible = 1
    RE = 0
    WallREtimer.Enabled = 1
End Sub

Sub WallREtimer_Timer
    Select Case RE
        Case 3:RubberRE1.Visible = 0:RubberRE2.Visible = 1
        Case 4:RubberRE2.Visible = 0:RubberRE.Visible = 1:WallREtimer.Enabled = 0:
    End Select
    RE = RE + 1
End Sub

Sub WallCCA_Hit
  AddScore(10)
  ToggleBumpers
End Sub


Sub WallCCB_Hit
  AddScore(10)
  ToggleBumpers
End Sub

'******************************************************
'           SWITCHES/TARGETS
'******************************************************

Sub RO_1000A_Hit()
  AddScore(1000)
  playsound "sensor", 0,0.1,AudioPan(ActiveBall),0,0,0,1,AudioFade(ActiveBall)
End Sub

Sub RO_1000B_Hit()
  AddScore(1000)
  playsound "sensor", 0,0.1,AudioPan(ActiveBall),0,0,0,1,AudioFade(ActiveBall)
End Sub

Sub RO_1000C_Hit()
  AddScore(1000)
  playsound "sensor", 0,0.1,AudioPan(ActiveBall),0,0,0,1,AudioFade(ActiveBall)
End Sub

Sub Out1000L_Hit()
  AddScore(1000)
  playsound "sensor", 0,0.1,AudioPan(ActiveBall),0,0,0,1,AudioFade(ActiveBall)
End Sub

Sub Out1000R_Hit()
  AddScore(1000)
  playsound "sensor", 0,0.1,AudioPan(ActiveBall),0,0,0,1,AudioFade(ActiveBall)
End Sub


Sub RO_Collect_Hit()
  CollectTunnel
  playsound "sensor", 0,0.1,AudioPan(ActiveBall),0,0,0,1,AudioFade(ActiveBall)
End Sub

Sub RO_CollectR_Hit()
  StopTunnel
  CollectTunnel
  CloseGates
  playsound "sensor", 0,0.1,AudioPan(ActiveBall),0,0,0,1,AudioFade(ActiveBall)
End Sub

Sub RO_CollectL_Hit()
  StopTunnel
' CollectTunnel 'Disable from kicker if enabled here, beward that tunnel may be collected twice
  playsound "sensor", 0,0.1,AudioPan(ActiveBall),0,0,0,1,AudioFade(ActiveBall)
End Sub

Sub RO_Stop_Hit()
  StopTunnel
  playsound "sensor", 0,0.1,AudioPan(ActiveBall),0,0,0,1,AudioFade(ActiveBall)
End Sub

Sub T100_Hit()
  AddScore(100)
  DOF 111, 2
End Sub

Sub T100STL_Hit()
  AddScore(100)
  StartTunnel
  DOF 112, 2
End Sub

Sub T100STR_Hit()
  AddScore(100)
  StartTunnel
  DOF 113, 2
End Sub

Sub TCenter_Hit()
  UpPost
  CollectTunnel
  DOF 114, 2
End Sub

Sub TEB_Hit()
  DOF 115, 2
  If LEB.state = 1 Then
    LSA.state = 1
    LSA1.state = 1
    LSA2.state = 1
    LEB.state = 0
    LEB1.state = 0
    LEB2.state = 0
    If B2SOn Then
      Controller.B2SSetShootAgain 36, 1
    End If
  End If
End Sub

Sub DP_L_Hit()
  DownPost
  leftrollover_prim.z = -3
End Sub

Sub DP_L_UnHit()
  leftrollover_prim.z = 0
End Sub

Sub DP_R_Hit()
  DownPost
  rightrollover_prim.z = -3
End Sub

Sub DP_R_UnHit()
  rightrollover_prim.z = 0
End Sub


'******************************************************
'             POSTS/GATES
'******************************************************

Sub UpPost()
  If CenterPost_prim.TransZ=0 Then
    PlaySoundAt SoundFX("solenoid",DOFContactors), centerpost_prim
    DOF 116,2
  End If

  lp2.state = 1
  lp3.state = 1
  Post.collidable = true
  flipperleft_prim.image= "flipperleftON"
  flipperright_prim.image= "flipperrightON"
  centerpost_prim.image="3D_Centerpost_2"
  centerpost_prim.BlendDisableLighting = 0.3
  CenterPost_prim.TransZ = 25
End Sub

Sub DownPost()
  If CenterPost_prim.TransZ=25 Then
    PlaySoundAt SoundFX("solenoid",DOFContactors), centerpost_prim
    DOF 116,2
  End If
  lp2.state = 0
  lp3.state = 0
  Post.collidable = false
  flipperleft_prim.image= "flipperleftOFF"
  flipperright_prim.image= "flipperrightOFF"
  centerpost_prim.image="3D_Centerpost"
  centerpost_prim.BlendDisableLighting = 0
  CenterPost_prim.Transz = 0
End Sub

Sub OpenTopGate()
  LGOa.state = 1
  LGOa1.state = 1
  LGOa2.state = 1
  gatetopright.collidable = false
  PlaySoundAt SoundFX("click",DOFContactors), Gate1
  DOF 117,2
End Sub

Sub OpenBottomGate()
  LGOb.state = 1
  LGOb1.state = 1
  LGOb2.state = 1
  gatebottomright.collidable = false
  PlaySoundAt SoundFX("click",DOFContactors), Gate2
  DOF 118,2
End Sub


Sub CloseGates()
  LGOa.state = 0
  LGOa1.state = 0
  LGOa2.state = 0
  gatetopright.collidable = true

  LGOb.state = 0
  LGOb1.state = 0
  LGOb2.state = 0
  gatebottomright.collidable = true
End Sub

'******************************************************
'             BUMPERS
'******************************************************

Sub Bumper1_Hit()
  PlaySoundAtVol "fx_bumper1", Bumper1, 1
  DOF 107, 2

  If LBump1.state = 1 Then
    AddScore(100)
  Else
    AddScore(10)
    Bump1On
  End If
End Sub

Sub Bumper2_Hit()
  PlaySoundAtVol "fx_bumper2", Bumper2, 1
  DOF 108, 2

  If LBump2.state = 1 Then
    AddScore(100)
  Else
    AddScore(10)
    Bump2On
  End If
End Sub

Sub Bumper3_Hit()
  PlaySoundAtVol "fx_bumper3", Bumper3, 1
  DOF 109, 2

  If LBump2.state = 1 Then
    AddScore(100)
  Else
    AddScore(10)
    Bump2On
  End If
End Sub

Sub Bumper4_Hit()
  PlaySoundAtVol "fx_bumper4", Bumper4, 1
  DOF 110, 2

  If LBump1.state = 1 Then
    AddScore(100)
  Else
    AddScore(10)
    Bump1On
  End If
End Sub

Sub ToggleBumpers()
  If LBump1.state = 1 Then
    Bump2On
  Else
    Bump1On
  End If
End Sub

Sub Bump1On()
  LBump1.state=1
  LBump1a.state=1
  LBump4.state=1
  LBump4a.state=1
  LBump2.state=0
  LBump2a.state=0
  LBump3.state=0
  LBump3a.state=0
  bumpcap1.image="bump1On"
  bumpcap2.image="bump1Off"
  bumpcap1.blenddisablelighting=0.5
  bumpcap2.blenddisablelighting=0

End Sub

Sub Bump2On()
  LBump1.state=0
  LBump1a.state=0
  LBump4.state=0
  LBump4a.state=0
  LBump2.state=1
  LBump2a.state=1
  LBump3.state=1
  LBump3a.state=1
  bumpcap1.image="bump1Off"
  bumpcap2.image="bump1On"
  bumpcap1.blenddisablelighting=0
  bumpcap2.blenddisablelighting=0.5
End Sub

'******************************************************
'             KICKERS
'******************************************************

Dim KickBalls
Dim KickerBall1, KickerBall2

'*** PI returns the value for PI
Function PI()
  PI = 4*Atn(1)
End Function

Sub KickBall(kball, kangle, kvel, kvelz, kzlift)
  dim rangle
  rangle = PI * (kangle - 90) / 180

  kball.z = kball.z + kzlift
  kball.velz = kvelz
  kball.velx = cos(rangle)*kvel
  kball.vely = sin(rangle)*kvel
End Sub

'**kickers need a timer to give a pause before kickout

Dim rkickstep, lkickstep, llkickstep, ktimer

Sub kickerL_hit() 'Left Kicker at the top of the playfield
  PlaySoundat "Hole_enter", kickerL
  set KickerBall1 = activeball
  lkickstep = 0

  If TableTilted = true then
    kickerL.timerenabled=1
    else
    StartTunnel
    OpenBottomGate
    ktimer = "L"
    MultiScore 5, 100
  end if
End Sub


Sub kickerL_timer
    Select Case LKickstep
        Case 1:
      if kickerl.ballcntover > 0 then
        kickarmtop_prim.ObjRotX=-12
        KickBall KickerBall1, 167, 9, 5, 30
'       kickerL.kick 167+INT(RND*6),12+INT(RND*2)
        PlaysoundAt SoundFXDOF("popper_ball",119,DOFPulse,DOFContactors), kickerL
      Else
        kickerl.timerenabled=0
      end if
        Case 2:kickarmtop_prim.ObjRotX = -45
        Case 3:kickarmtop_prim.ObjRotX = -45
        Case 4:kickarmtop_prim.ObjRotX = -45
        Case 5:kickarmtop_prim.ObjRotX = -45
        Case 6:kickarmtop_prim.ObjRotX = -45
        Case 7:kickarmtop_prim.ObjRotX = -45
        Case 8:kickarmtop_prim.ObjRotX = -45
        Case 9:kickarmtop_prim.ObjRotX = -45
        Case 10:kickarmtop_prim.ObjRotX = -45
        Case 11:kickarmtop_prim.ObjRotX = 24
        Case 12:kickarmtop_prim.ObjRotX = 12
        Case 13:kickarmtop_prim.ObjRotX = 0:kickerL.TimerEnabled = 0
    End Select
    LKickstep = LKickstep + 1
End Sub

Sub kickerr_hit() 'Right Kicker at the top of the playfield
  PlaySoundat "Hole_enter", kickerR
  set KickerBall1 = activeball
  rkickstep = 0

  If TableTilted = true then
    kickerR.timerenabled=1
    else
    StartTunnel
    OpenTopGate
    ktimer = "R"
    MultiScore 5, 100
  end if
End Sub


Sub kickerr_timer
    Select Case rkickstep
        Case 1:
      if kickerr.ballcntover > 0 then
        kickarmtop_prim1.ObjRotX=12
        KickBall KickerBall1, -167, 9, 5, 30
'       kickerL.kick 167+INT(RND*6),12+INT(RND*2)
        PlaysoundAt SoundFXDOF("popper_ball",120,DOFPulse,DOFContactors), kickerR
      Else
        kickerr.timerenabled=0
      end if
'     kickerr.kick -167+INT(RND*6),12+INT(RND*2)
        Case 2:kickarmtop_prim1.ObjRotX = -45
        Case 3:kickarmtop_prim1.ObjRotX = -45
        Case 4:kickarmtop_prim1.ObjRotX = -45
        Case 5:kickarmtop_prim1.ObjRotX = -45
        Case 6:kickarmtop_prim1.ObjRotX = -45
        Case 7:kickarmtop_prim1.ObjRotX = -45
        Case 8:kickarmtop_prim1.ObjRotX = -45
        Case 9:kickarmtop_prim1.ObjRotX = -45
        Case 10:kickarmtop_prim1.ObjRotX = -45
        Case 11:kickarmtop_prim1.ObjRotX = 24
        Case 12:kickarmtop_prim1.ObjRotX = 12
        Case 13:kickarmtop_prim1.ObjRotX = 0:kickerr.TimerEnabled = 0
    End Select
    rkickstep = rkickstep + 1
End Sub

Sub kickerll_hit()  'Lower Left Kicker under plastic
  ktimer = "LL"
  CollectTunnel
End Sub

Sub kickerll_timer
  kickerll.kick 0,24
  PlaysoundAT SoundFXDOF("popper_ball",121,DOFPulse,DOFContactors), kickerll
  kickerll.TimerEnabled = 0
End Sub


'******************************************************
'             BONUS
'******************************************************
Dim StopTunnelTimer, TunnelValue, TunnelCount
StopTunnelTimer = 0
TunnelCount = 0
TunnelValue = 1
TunnelTimer.interval = 60

Sub CollectTunnel()
  StopTunnel
  MultiScore TunnelValue, 1000
End Sub

Sub StopTunnel()
  Stopsound "fx_motor"
  DOF 123, 0
  StopTunnelTimer = 1
End Sub

Sub StartTunnel()
  PlaySound "fx_motor", -1, 1, AudioPan(centerpost_prim), 0, 0, 1, 0, AudioFade(centerpost_prim)
  DOF 123, 1
  TunnelTimer.enabled=true
  StopTunnelTimer = 0
End Sub

Sub TunnelTimer_Timer()
  TunnelCount = TunnelCount + 1
  If TunnelCount > 14 Then TunnelCount = 0
  Select Case TunnelCount
    Case 0:
      tunnel_prim.image = "tunnel1000"
      TunnelValue = 1
      DOF 154,0
      DOF 150,1
      if StopTunnelTimer = 1 Then TunnelTimer.enabled = false
    Case 1: tunnel_prim.image = "tunnel1000a": TunnelValue = 2
    Case 2: tunnel_prim.image = "tunnel1000b"
    Case 3:
      tunnel_prim.image = "tunnel2000"
      TunnelValue = 2
      DOF 150,0
      DOF 151,1
      if StopTunnelTimer = 1 Then TunnelTimer.enabled = false
    Case 4: tunnel_prim.image = "tunnel2000a": TunnelValue = 3
    Case 5: tunnel_prim.image = "tunnel2000b"
    Case 6:
      tunnel_prim.image = "tunnel3000"
      TunnelValue = 3
      DOF 151,0
      DOF 152,1
      if StopTunnelTimer = 1 Then TunnelTimer.enabled = false
    Case 7: tunnel_prim.image = "tunnel3000a": TunnelValue = 4
    Case 8: tunnel_prim.image = "tunnel3000b"
    Case 9:
      tunnel_prim.image = "tunnel4000"
      TunnelValue = 4
      DOF 152,0
      DOF 153,1
      if StopTunnelTimer = 1 Then TunnelTimer.enabled = false
    Case 10: tunnel_prim.image = "tunnel4000a": TunnelValue = 5
    Case 11: tunnel_prim.image = "tunnel4000b"
    Case 12:
      tunnel_prim.image = "tunnel5000"
      TunnelValue = 5
      DOF 153,0
      DOF 154,1
      if StopTunnelTimer = 1 Then TunnelTimer.enabled = false
    Case 13: tunnel_prim.image = "tunnel5000a": TunnelValue = 1
    Case 14: tunnel_prim.image = "tunnel5000b"
  End Select
End Sub

'******************************************************
'           SCORING/CREDITS
'******************************************************

Sub AddScore(x)
  If TableTilted = 0 Then
    PrevScore(Player) = Score(Player)
'   if isScoring() = 0 Then
      If x = 10 Then
        Play10Bell
'       CheckForRoll(10)
        Score(Player) = Score(Player) + 10
        AdvMatch
      ElseIf x = 100 Then
        Play100Bell
'       CheckForRoll(100)
        Score(Player) = Score(Player) + 100
      ElseIf x = 1000 Then
        Play1000Bell
'       CheckForRoll(1000)
        Score(Player) = Score(Player) + 1000
      End If
'   End If
    CheckFreeGame
  End If
   EVAL("ScoreReel"&player).setvalue score(player)
  If B2SOn Then Controller.B2SSetScorePlayer Player,Score(Player)
End Sub

Sub Play10Bell()
  PlaySound SoundFX("bell10",DOFChimes), 1, 1, AudioPan(kickerr), 0,0,0,1,AudioFade(kickerr)
  DOF 141, 2
End Sub

Sub Play100Bell()
  PlaySound SoundFX("bell100",DOFChimes), 1, 1, AudioPan(kickerr), 0,0,0,1,AudioFade(kickerr)
  DOF 142, 2
End Sub

Sub Play1000Bell()
  PlaySound SoundFX("bell1000",DOFChimes), 1, 1, AudioPan(kickerr), 0,0,0,1,AudioFade(kickerr)
  DOF 143, 2
End Sub


'Sub CheckForRoll(x)
' Select Case (x)
'   Case 10:
'     PulseReel5.enabled = 1
'
'     Reel5Value = (Reel5Value + 1) mod 10
'     If B2SOn Then Controller.b2ssetreel 5,Reel5Value
'
'     PlaySoundAtVol "reelclick5", emp1r5, ReelClickVol
'     if emp1r5.rotx = 261 then CheckForRoll(100)
'   Case 100:
'     PulseReel4.enabled = 1
'
'     Reel4Value = (Reel4Value + 1) mod 10
'     If B2SOn Then Controller.b2ssetreel 4,Reel4Value
'
'     PlaySoundAtVol "reelclick4", emp1r4, ReelClickVol
'     if emp1r4.rotx = 261 then CheckForRoll(1000)
'   Case 1000:
'     PulseReel3.enabled = 1
'
'     Reel3Value = (Reel3Value + 1) mod 10
'     If B2SOn Then Controller.b2ssetreel 3,Reel3Value
'
'     PlaySoundAtVol "reelclick3", emp1r3, ReelClickVol
'     if emp1r3.rotx = 261 then CheckForRoll(10000)
'   Case 10000:
'     PulseReel2.enabled = 1
'
'     Reel2Value = (Reel2Value + 1) mod 10
'     If B2SOn Then Controller.b2ssetreel 2,Reel2Value
'
'     PlaySoundAtVol "reelclick2", emp1r2, ReelClickVol
'     if emp1r2.rotx = 261 then CheckForRoll(100000)
'   Case 100000:
'     PulseReel1.enabled = 1
'
'     Reel1Value = (Reel1Value + 1) mod 10
'     If B2SOn Then Controller.b2ssetreel 1,Reel1Value
'
'     PlaySoundAtVol "reelclick1", emp1r1, ReelClickVol
' End Select
'End Sub

Sub AddCredit(direction)
  if direction > 0 and credits >= 25 Then
    'do Nothing
  Else
    if direction = 1 Then
      'playsound SoundFXDOF("knocker",124,DOFPulse,DOFKnocker)
    end if
    credits = credits + direction
    PlaySoundAtVolLoops "reelclick", kickerr, 0.5, 0
  End If

  If Credits > 0 Then
    DOF 125, 1
  Else
    DOF 125, 0
  End If

  CreditReel.setvalue Credits
  If B2SOn Then Controller.B2SSetCredits Credits
End Sub

Sub CheckFreeGame()
  If PrevScore(Player) < Replay1 And Score(Player) >= Replay1 Then FreeGame
  If PrevScore(Player) < Replay2 And Score(Player) >= Replay2 Then FreeGame
  If PrevScore(Player) < Replay3 And Score(Player) >= Replay3 Then FreeGame
  If PrevScore(Player) < Replay4 And Score(Player) >= Replay4 Then FreeGame
End Sub

Sub FreeGame()
  playsound SoundFXDOF("knocker",124,DOFPulse,DOFKnocker)
  AddCredit(1)
End Sub

Dim mCountDown, mScore

Sub MultiScore(cycle, Score)
  mScore = Score
  mCountDown = cycle
  MultiScoreTimer.enabled = true
End Sub

Sub MultiScoreTimer_Timer
  mCountDown = mCountDown - 1
  AddScore(mscore)
  If mCountDown < 1 Then
    me.enabled = False
    If ktimer = "L" Then
      kickerL.timerenabled = true
    Elseif ktimer = "R" Then
      kickerR.timerenabled = true
    Elseif ktimer = "LL" Then
      kickerLL.timerenabled = true
    End If
    ktimer = ""
  End If
End Sub

'Function isScoring()
' if PulseReel2.enabled = 0 and PulseReel3.enabled = 0 and PulseReel4.enabled = 0 and PulseReel5.enabled = 0 Then
'   isScoring = 0
' Else
'   isScoring = 1
' End If
'End Function


'******************************************************
'             DRAIN
'******************************************************

Sub Drain_Hit()
  PlaySound "drain",0,1,AudioPan(Drain),0.25,0,0,1,AudioFade(Drain)
  PlaySoundAtVolLoops "End_Of_Game", kickerr, 0.25, 0
  DOF 126, DOFPulse

  StopTunnel
  DownPost
  CloseGates
  LEB.state = 0:LEB1.state = 0:LEB2.state = 0

  If TableTilted = 1 Then
    TableTilted = 0
    ResetTilt
  End If

  NextBall
End Sub

Sub Drain_Timer()
  StartTunnel
  Drain.Kick 60, 20
  DOF 105, DOFPulse
  PlaySound SoundFX("ballrelease",DOFContactors), 0,1,AudioPan(Drain),0.25,0,0,1,AudioFade(Drain)
  me.timerenabled = False
End Sub

Sub NextBall()
  If LSA.state = 1 Then
    LSA.state = 0
    LSA1.state = 0
    LSA2.state = 0
    If B2SOn Then Controller.B2SSetShootAgain 36,0
  Elseif Player = Players Then
    BallInPlay = BallInPlay + 1
    Player = 1
  Else
    Player = Player + 1
  End If
  if ballinplay > BallsPerGame then
    InProgress = False

    ballinplaytxt.text = ""
    gamov.text = "GAME OVER"
    If B2SOn Then
      Controller.B2SSetGameOver 1
      Controller.B2SSetBallInPlay 0
      ResetPlayerUp
            Controller.B2SSetCanPlay 31, 0
    End If

    CheckMatch

    LeftFlipper.RotateToStart
    StopSound "buzzL"
    DOF 101, 0

    RightFlipper.RotateToStart
    StopSound "buzz"
    DOF 102, 0

    If Score(2) > Score(1) Then
      HighPlayer = 2
    Else
      HighPlayer = 1
    End If

    If Score(3) > Score(HighPlayer) Then
      HighPlayer = 3
    End If

    If Score(4) > Score(HighPlayer) Then
      HighPlayer = 4
    End If


    If Score(HighPlayer) >= HighScore Then
      FreeGame
      HighScore = Score(HighPlayer)
      HighScoreEntryInit()
    End If
  Else
    ballinplaytxt.text = BallInPlay
    If B2SOn Then
      Controller.B2SSetBallInPlay BallInPlay
      SetPlayerUp Player
    End If
    Drain.timerenabled = 1
  End If
  PlayQuietClick
End Sub

Sub ResetPlayerUp()
  Controller.B2SSetData 40, 0
  Controller.B2SSetData 41, 0
  Controller.B2SSetData 42, 0
  Controller.B2SSetData 43, 0
End Sub

Sub SetPlayerUp(Player)
  ResetPlayerUp
  If Player = 1 Then
    Controller.B2SSetData 40, 1
  Elseif Player = 2 Then
    Controller.B2SSetData 41, 1
  Elseif Player = 3 Then
    Controller.B2SSetData 42, 1
  Elseif Player = 4 Then
    Controller.B2SSetData 43, 1
  End If
End Sub

'******************************************************
'           START NEW GAME
'******************************************************

Dim NGCount

Sub StartNewGame_Timer()
  NGCount = NGCount + 1

  Select Case (NGCount)
    Case 1:
      Matchtxt.text = ""
      Tilttxt.text = ""
      gamov.text = ""
      ballinplaytxt.text = ballinplay
      canplay.text = Players

      If LBump1.state = 0 and LBump2.state = 0 Then
        ToggleBumpers
      End If

      ScoreReel1.resettozero
      ScoreReel2.resettozero
      ScoreReel3.resettozero
      ScoreReel4.resettozero


      If B2SOn Then
        Controller.B2SSetTilt 0
        Controller.B2SSetMatch 0
        Controller.B2SSetScorePlayer 1,0
        Controller.B2SSetScorePlayer 2,0
        Controller.B2SSetScorePlayer 3,0
        Controller.B2SSetScorePlayer 4,0
        Controller.B2SSetGameOver 0
        Controller.B2SSetBallInPlay BallInPlay
        SetPlayerUp Player
      End If

      ResetTilt

      PlaySoundAtVol "MotorRunning", kickerr, 0.25
      PlayLoudClick
    Case 2:
      PlayLoudClick
'     ResetReels
    Case 3:
      PlayLoudClick
'     ResetReels
    Case 4:
      PlayLoudClick
'     ResetReels
    Case 5:
      PlayLoudClick
'     ResetReels
    Case 7:
      PlayQuietClick
'     ResetReels
    Case 8:
      PlayQuietClick
'     ResetReels
    Case 9:
      PlayQuietClick
'     ResetReels
    Case 10:
      PlayQuietClick
'     ResetReels
    Case 11:
      PlayQuietClick
'     ResetReels

    Case 13:
      Drain.timerenabled = true
      Score(1) = 0
      Score(2) = 0
      Score(3) = 0
      Score(4) = 0
'     SetReels

'     If B2SOn Then
'       Controller.b2ssetreel 1,0:Reel1Value = 0
'       Controller.b2ssetreel 2,0:Reel2Value = 0
'       Controller.b2ssetreel 3,0:Reel3Value = 0
'       Controller.b2ssetreel 4,0:Reel4Value = 0
'       Controller.b2ssetreel 5,0:Reel5Value = 0
'     End If
      me.enabled = false
      NGCount = 0
  End Select
End Sub

Sub PlayLoudClick()
  PlaySoundAtVol "Motor_Click_Loud_Long", kickerr, 0.05
End Sub

Sub PlayQuietClick()
  PlaySoundAtVol "Motor_Click_Quiet_Long2", kickerr, 0.05
End Sub


'******************************************************
'               MATCH
'******************************************************

Dim MatchStep

Sub CheckMatch()
  If Match = 0 Then
    Matchtxt.text = "00"
  Else
    Matchtxt.text = Match
  End If
  SetB2SMatch

  MatchStep=0
  MatchTimer.enabled = 1
End Sub

Sub MatchTimer_Timer()
  If MatchStep < Players Then
    MatchStep = MatchStep + 1
    if Match=(Score(MatchStep) mod 100) then
      FreeGame
    end if
  End If
End Sub


Sub AdvMatch()
  LEB.state = 0:LEB1.state = 0 :LEB2.state = 0
  Select Case (Match)
    Case 30: Match = 80
    Case 80: Match = 20
    Case 20: Match = 50
    Case 50: Match = 90
    Case 90: Match = 40
    Case 40: Match = 0
    Case 0: Match = 60
    Case 60: Match = 10
    Case 10: Match = 70:LEB.state = 1:LEB1.state = 1:LEB2.state = 1
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
  CloseGates
  LEB.state = 0:LEB1.state = 0:LEB2.state = 0
  DownPost

  Bumper1.threshold = 100
  Bumper2.threshold = 100
  Bumper3.threshold = 100
  Bumper4.threshold = 100

  LeftSlingShot.SlingShotThreshold = 100
  RightSlingShot.SlingShotThreshold = 100

  PlayLoudClick

  Tilttxt.text = "TILT"
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

  Bumper1.threshold = 1
  Bumper2.threshold = 1
  Bumper3.threshold = 1
  Bumper4.threshold = 1

  LeftSlingShot.SlingShotThreshold = 1.2
  RightSlingShot.SlingShotThreshold = 1.2

  Tilttxt.text = ""
  If B2SOn Then Controller.B2SSetTilt 0
End Sub

'******************************************************
'             B2S LIGHTS
'******************************************************

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

Dim LightsOn

Sub LightsOnOff_Timer()
  If LightsOn Then
    DOF 155, 0
    DOF 156, 1
    LightsOn = False
  Else
    DOF 155, 1
    DOF 156, 0
    LightsOn = True
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

Sub PlaySoundAtVolLoops(sound, tableobj, Vol, Loops)
  PlaySound sound, Loops, Vol, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtVol(sound, tableobj, Vol)
    PlaySound sound, 1, Vol, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
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


'*****************************************
'   FLIPPER SHADOWS
'*****************************************

sub FlipperTimer_Timer()
  FlipperLSh.RotZ = LeftFlipper.currentangle
  FlipperRSh.RotZ = RightFlipper.currentangle
    flipperleft_prim.rotz = leftflipper.currentangle + -120
    flipperright_prim.rotz = rightflipper.currentangle+ 120
End Sub

'*****************************************
'           BALL SHADOW
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
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/16) + ((BOT(b).X - (Table1.Width/2))/17))' + 13
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/16) + ((BOT(b).X - (Table1.Width/2))/17))' - 13
        End If
        ballShadow(b).Y = BOT(b).Y + 10
        If BOT(b).Z > 20 Then
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


'******************************************************
'           LOAD & SAVE TABLE
'******************************************************
sub savehs
    savevalue "SpaceTime_VPX", "Credits", Credits
  savevalue "SpaceTime_VPX", "Match", Match
  savevalue "SpaceTime_VPX", "HighScore", Highscore
  savevalue "SpaceTime_VPX", "HSA1", HSA1
  savevalue "SpaceTime_VPX", "HSA2", HSA2
  savevalue "SpaceTime_VPX", "HSA3", HSA3
  savevalue "SpaceTime_VPX", "BallInPlay", BallInPlay
  savevalue "SpaceTime_VPX", "Score", Score(1)
  savevalue "SpaceTime_VPX", "Score2", Score(2)
  savevalue "SpaceTime_VPX", "Score3", Score(3)
  savevalue "SpaceTime_VPX", "Score4", Score(4)
end sub

sub loadhs
  HighScore=0
  Credits=0
  Match=0

    dim temp
  temp = LoadValue("SpaceTime_VPX", "Credits")
    If (temp <> "") then Credits = CDbl(temp)
    temp = LoadValue("SpaceTime_VPX", "Match")
    If (temp <> "") then Match = CDbl(temp)
    temp = LoadValue("SpaceTime_VPX", "HighScore")
    If (temp <> "") then HighScore = CDbl(temp)
    temp = LoadValue("SpaceTime_VPX", "hsa1")
    If (temp <> "") then HSA1 = CDbl(temp)
    temp = LoadValue("SpaceTime_VPX", "hsa2")
    If (temp <> "") then HSA2 = CDbl(temp)
    temp = LoadValue("SpaceTime_VPX", "hsa3")
    If (temp <> "") then HSA3 = CDbl(temp)
    temp = LoadValue("SpaceTime_VPX", "BallInPlay")
    If (temp <> "") then BallInPlay = CDbl(temp)
    temp = LoadValue("SpaceTime_VPX", "Score")
    If (temp <> "") then Score(1) = CDbl(temp)
    temp = LoadValue("SpaceTime_VPX", "Score2")
    If (temp <> "") then Score(2) = CDbl(temp)
  temp = LoadValue("SpaceTime_VPX", "Score3")
    If (temp <> "") then Score(3) = CDbl(temp)
    temp = LoadValue("SpaceTime_VPX", "Score4")
    If (temp <> "") then Score(4) = CDbl(temp)

  if HighScore=0 then HighScore=50000
end sub



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


'******************************************************
'   FLIPPER CORRECTION SUPPORTING FUNCTIONS
'******************************************************

Sub AddPt(aStr, idx, aX, aY)  'debugger wrapper for adjusting flipper script in-game
  dim a : a = Array(LF, RF)
  dim x : for each x in a
    x.addpoint aStr, idx, aX, aY
  Next
End Sub

'Methods:
'.TimeDelay - Delay before trigger shuts off automatically. Default = 80 (ms)
'.AddPoint - "Polarity", "Velocity", "Ycoef" coordinate points. Use one of these 3 strings, keep coordinates sequential. x = %position on the flipper, y = output
'.Object - set to flipper reference. Optional.
'.StartPoint - set start point coord. Unnecessary, if .object is used.

'Called with flipper -
'ProcessBalls - catches ball data.
' - OR -
'.Fire - fires flipper.rotatetoend automatically + processballs. Requires .Object to be set to the flipper.

Class FlipperPolarity
  Public DebugOn, Enabled
  Private FlipAt  'Timer variable (IE 'flip at 723,530ms...)
  Public TimeDelay  'delay before trigger turns off and polarity is disabled TODO set time!
  private Flipper, FlipperStart, FlipperEnd, LR, PartialFlipCoef
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
  Public Property Let EndPoint(aInput) : if IsObject(aInput) then FlipperEnd = aInput.x else FlipperEnd = aInput : end if : End Property
  Public Property Get EndPoint : EndPoint = FlipperEnd : End Property

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
        if DebugOn then StickL.visible = True : StickL.x = balldata(x).x    'debug TODO
      End If
    Next
    PartialFlipCoef = ((Flipper.StartAngle - Flipper.CurrentAngle) / (Flipper.StartAngle - Flipper.EndAngle))
    PartialFlipCoef = abs(PartialFlipCoef-1)
    if abs(Flipper.currentAngle - Flipper.EndAngle) < 30 Then
      PartialFlipCoef = 0
    End If
  End Sub
  Private Function FlipperOn() : if gameTime < FlipAt+TimeDelay then FlipperOn = True : End If : End Function 'Timer shutoff for polaritycorrect

  Public Sub PolarityCorrect(aBall)
    if FlipperOn() then
      dim tmp, BallPos, x, IDX, Ycoef : Ycoef = 1
      dim teststr : teststr = "Cutoff"
      tmp = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
      if tmp < 0.1 then 'if real ball position is behind flipper, exit Sub to prevent stucks  'Disabled 1.03, I think it's the Mesh that's causing stucks, not this
        if DebugOn then TestStr = "real pos < 0.1 ( " & round(tmp,2) & ")" : tbpl.text = Teststr
        'RemoveBall aBall
        'Exit Sub
      end if

      'y safety Exit
      if aBall.VelY > -8 then 'ball going down
        if DebugOn then teststr = "y velocity: " & round(aBall.vely, 3) & "exit sub" : tbpl.text = teststr
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
        if DebugOn then set tmp = new spoofball : tmp.data = aBall : End If
        if IsEmpty(BallData(idx).id) and aBall.VelY < -12 then 'if tip hit with no collected data, do vel correction anyway
          if PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1) > 1.1 then 'adjust plz
            VelCoef = LinearEnvelope(5, VelocityIn, VelocityOut)
            if partialflipcoef < 1 then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)
            if Enabled then aBall.Velx = aBall.Velx*VelCoef'VelCoef
            if Enabled then aBall.Vely = aBall.Vely*VelCoef'VelCoef
            if DebugOn then teststr = "tip protection" & vbnewline & "velcoef: " & round(velcoef,3) & vbnewline & round(PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1),3) & vbnewline
            'debug.print teststr
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
        'playsound "fx_knocker"
        if Enabled then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef)
      End If
      'debug
      if DebugOn then
        TestStr = teststr & "%pos:" & round(BallPos,2)
        if IsEmpty(PolarityOut(0) ) then
          teststr = teststr & vbnewline & "(Polarity Disabled)" & vbnewline
        else
          teststr = teststr & "+" & round(1 *(AddX*ycoef*PartialFlipcoef),3)
          if BallPos >= PolarityOut(uBound(PolarityOut) ) then teststr = teststr & "(MAX)" & vbnewline else teststr = teststr & vbnewline end if
          if Ycoef < 1 then teststr = teststr &  "ycoef: " & ycoef & vbnewline
          if PartialFlipcoef < 1 then teststr = teststr & "PartialFlipcoef: " & round(PartialFlipcoef,4) & vbnewline
        end if

        teststr = teststr & vbnewline & "Vel: " & round(BallSpeed(tmp),2) & " -> " & round(ballspeed(aBall),2) & vbnewline
        teststr = teststr & "%" & round(ballspeed(aBall) / BallSpeed(tmp),2)
        tbpl.text = TestSTR
      end if
    Else
      'if DebugOn then tbpl.text = "td" & timedelay
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
    x.TimeDelay = 80
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

'Trigger Hit - .AddBall activeball
'Trigger UnHit - .PolarityCorrect activeball

Sub TriggerLF_Hit() : LF.Addball activeball : End Sub
Sub TriggerLF_UnHit() : LF.PolarityCorrect activeball: End Sub
Sub TriggerRF_Hit() : RF.Addball activeball : End Sub
Sub TriggerRF_UnHit() : RF.PolarityCorrect activeball : End Sub

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

Option Explicit
Randomize

Const BallSize = 50
Const BallMass = 1.7
Const TopShadow = False ' Turn on for advanced shadows


On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="fire_l3",UseSolenoids=1,UseLamps=0,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff", SCoin="coin"

LoadVPM "01560000", "S11.VBS", 3.26
Dim DesktopMode: DesktopMode = Table1.ShowDT

If DesktopMode = True Then 'Show Desktop components
Ramp16.visible=1
Ramp15.visible=1
'Primitive13.visible=1
Else
Ramp16.visible=0
Ramp15.visible=0
'Primitive13.visible=0
End if

PFTopShadow.visible = TopShadow

'*************************************************************
'Solenoid Call backs
'**********************************************************************************************************
  SolCallback(1) = "bsTrough.SolIn"
  SolCallback(2) = "bsTrough.SolOut"
  SolCallback(3) = "LRUp"
  SolCallback(4) = "LRDown"
  SolCallback(5) = "PlugM" 'Center Post
  SolCallback(6) = "bsREject.SolOut"
  SolCallback(7) = "RRUp"
  SolCallback(8) =  "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
  SolCallback(9) = "CRMotor"
  SolCallback(10) = "PFGI"
  SolCallback(13) = "SolFlamesMotor"
  SolCallback(14) = "Sol14"
  SolCallback(15) = "bsLKick.SolOut"
  SolCallback(23) = "vpmNudge.SolGameOn"
  SolCallback(29) = "RRDown"
  SolCallback(32) = "SolBell"

  SolCallback(16) = "SetLamp 116,"
  SolCallback(25) = "SetLamp 125,"
  SolCallback(26) = "SetLamp 126,"
  SolCallback(27) = "SetLamp 127,"
  SolCallback(28) = "Sol28"


  SolCallback(30) = "SetLamp 130,"
  SolCallback(31) = "SetLamp 131,"

SolCallBack(23) = "FastFlips.TiltSol"
SolCallback(sLRFlipper) = ""
SolCallback(sLLFlipper) = ""
SolCallback(sURFlipper) = ""
SolCallback(sULFlipper) = ""

Sub SolBell(enabled)
  if Enabled then PlaySound SoundFX("ding", DOFBell)
End Sub

Dim S14State:S14State = False

Sub Sol14(enabled)
  S14State = enabled
  SetLamp 150, (S14State OR FlamesMotor.Enabled)
End sub

Sub Sol28(enabled)
  SetLamp 128,enabled
End Sub

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySoundAt SoundFX("fx_Flipperup",DOFContactors),GI5:LeftFlipper.RotateToEnd
     Else
         PlaySoundAt SoundFX("fx_Flipperdown",DOFContactors),GI5:LeftFlipper.RotateToStart
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySoundAt SoundFX("fx_Flipperup",DOFContactors),GI8:RightFlipper.RotateToEnd
     Else
         PlaySoundAt SoundFX("fx_Flipperdown",DOFContactors),GI8:RightFlipper.RotateToStart
     End If
End Sub
'**********************************************************************************************************

'Solenoid Controlled toys
'**********************************************************************************************************

'Playfield GI
Sub PFGI(Enabled)
  SetLamp 151, not Enabled
  If Enabled Then
    dim xx
    For each xx in GI
      if TypeName(xx) <> "Flasher" then
        xx.State = 0
      end if
    Next
        PlaySoundAt "fx_relay_off", LightCenterBuilding
    plasticedges_prim.blenddisablelighting = 0
    rightrail_prim.blenddisablelighting = 0.05
    leftrail_prim.blenddisablelighting = 0.05
    backguide3_prim.blenddisablelighting = 0
    outers.blenddisablelighting = 0.1
    woodrail_prim.blenddisablelighting = 0
  Else
    For each xx in GI
      if TypeName(xx) <> "Flasher" then
        xx.State = 1
      end if
    Next
        PlaySoundAt "fx_relay", LightCenterBuilding
    plasticedges_prim.blenddisablelighting = 0.65
    rightrail_prim.blenddisablelighting = 0.15
    leftrail_prim.blenddisablelighting = 0.15
    backguide3_prim.blenddisablelighting = 0.15
    outers.blenddisablelighting = 0.25
    woodrail_prim.blenddisablelighting = 0.05
  End If
End Sub

Sub PlugM(enabled) 'Center Post
  if enabled then
    If PlugA.IsDropped then
      PlaySoundAt "solenoid", Plug
      Controller.Switch(26)=0
      Plug.TransZ = 0
      PlugA.IsDropped = 0
    else
      PlaySoundAt "fx_rubber2", Plug
      Controller.Switch(26)=1
      Plug.TransZ = 46
      PlugA.IsDropped = 1
    end if
  end if
End Sub

Sub PlugA_Hit()
  RandomBump 1, 3000
End Sub

Const RampMin=-10

'******Raising ramps
Dim LRDir, LRPos, RRDir, RRPos
LRDir = -1:LRPos = RampMin
RRDir = -1:RRPos = RampMin

'*****Ramp Init
LRampA.Collidable = 1:LRampB.Collidable = 0
RRampA.Collidable = 1:RRampB.Collidable = 0
LadderA.Collidable = 1




'Left Ramp
Sub LRUp(enabled)
  if enabled and LRPos = RampMin then
    Controller.Switch(41)=0
    PlaySoundAt "fx_leftmotor", LRamp
    LRDir = 1
    LRampTimer.Enabled = 1
    LRampA.Collidable = 0
    LRampB.Collidable = 1
  end if
End Sub

Sub LRDown(enabled)
  if enabled and LRPos = 0 then
    Controller.Switch(41)=1
    PlaySoundAt"fx_leftmotor", LRamp
    LRDir = -1
    LRampTimer.Enabled = 1
    LRampA.Collidable = 1
    LRampB.Collidable = 0
  end if
End Sub

Sub LRampTimer_Timer()
  LRPos = LRPos + LRDir
  If LRPos = 0 OR LRPos = RampMin then
    StopSound "fx_leftmotor"
    me.enabled = false
  end if
  LRamp.ObjRotx = LRPos
  LRampend.ObjRotx = LRPos
  leftliftmetal.ObjRotx = LRPos
End Sub



'RightRamp
Sub RRUp(enabled)
  if enabled and RRPos = RampMin then
    Controller.Switch(42)=0
    RRDir = 1
    RRampTimer.Enabled = 1
    PlaySoundAt "fx_rightmotor", RRamp
    RRampA.Collidable = 0
    RRampB.Collidable = 1
  end if
End Sub

Sub RRDown(enabled)
  if enabled and RRPos = 0 then
    Controller.Switch(42)=1
    RRampTimer.Enabled = 1
    RRDir = -1
    PlaySoundAt "fx_rightmotor", RRamp
    RRampA.Collidable = 1
    RRampB.Collidable = 0
  end if
End Sub

Sub RRampTimer_Timer()
  RRPos = RRPos + RRDir
  If RRPos = 0 or RRPos = RampMin then
    StopSound "fx_rightmotor"
    me.enabled = False
  end if
  RRamp.ObjRotx = RRPos
  RRampend.ObjRotx = RRPos
  rightliftmetal.ObjRotx = RRPos
End Sub

'Ladder Ramp
Dim StartLadderMotion:StartLadderMotion=0
Dim CRDir, CRPos: CRPos = 0: CRDir = -1

Function IsBallNearCenterRamp
  Dim Bot, b, cx, cy
  BOT = GetBalls
    For b = 0 to UBound(BOT)
    cx = BOT(b).x
    cy = BOT(b).y
    if cx > 370 AND cx < 490 AND cy > 195 AND cy < 395 Then
      IsBallNearCenterRamp = True
      exit Function
    end if
  Next
  IsBallNearCenterRamp = False
End Function


Sub CRMotor(enabled)
  if enabled then
    StartLadderMotion = Timer
    CRampTimer.Enabled = 1
  Else
    CRampTimer.Enabled = 0
  end if
End Sub

Sub CRampTimer_Timer()

  If StartLadderMotion > 0 AND Timer > StartLadderMotion Then
    ' Don't start motion yet if a ball is near ramp
    if IsBallNearCenterRamp then Exit Sub
    ' Make a half-propped ladder collideable
    ' This should prevent balls from rolling onto it from the sides unnaturally
    LadderMid.Collidable = 1
    if CRPos = 0 then
      CRDir = -.2':Ladder.ObjRotx = -24
    elseif CRpos = -30 then
      CRDir = 2:LadderA.Collidable = 0:LBWall.IsDropped = 1':Ladder.ObjRotx = 0
    end if
    PlaySoundAtVol "fx_centermotor", l9, 1
    Controller.Switch(63)=0
    Controller.Switch(59)=0
    StartLadderMotion = 0
  end if
  if StartLadderMotion = 0 Then
    CRPos = CRPos + CRDir
    If CRPos > 0 then CRPos = 0
    If CRPos < -30 then CRPos = -30
    Ladder.ObjRotx = CRPos
    Ladderrivets.ObjRotx = CRPos
    If CRPos = -30 then
      Controller.Switch(63)=1
      Controller.Switch(59)=0
      LadderMid.Collidable = 0
      LadderA.Collidable = 1:LBWall.IsDropped = 0
      StopSound "fx_centermotor"
      StartLadderMotion = Timer+1
    end If
    If CRPos = 0 then
      Controller.Switch(63)=0
      Controller.Switch(59)=1
      LadderMid.Collidable = 0
      StopSound "fx_centermotor"
      StartLadderMotion = Timer+1
    end If
  end if
End Sub

'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************

Dim bsTrough, bsLKick, bsREject

Sub Table1_Init
  vpmInit Me
  On Error Resume Next
    With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine = "Fire"&chr(13)&"You Suck"
    .HandleMechanics=0
    .HandleKeyboard=0
    .ShowDMDOnly=1
    .ShowFrame=0
    .ShowTitle=0
        .hidden = 0'1!!!!!
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With
     On Error Goto 0

  PinMAMETimer.Interval=PinMAMEInterval
  PinMAMETimer.Enabled=1

    vpmNudge.TiltSwitch=1
    vpmNudge.Sensitivity=1
    vpmNudge.TiltObj=Array(LeftSlingshot,RightSlingshot,LeftSlingshot1,RightSlingshot1)

      Set bsTrough = New cvpmBallStack
          bsTrough.InitSw 10, 56, 55, 54, 0, 0, 0, 0
          bsTrough.InitKick BallRelease, 90, 4
          bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
          bsTrough.Balls = 3

      Set bsREject = New cvpmBallStack
          bsREject.InitSaucer sw53, 53, 180, 5
          bsREject.KickZ = 0.5
          bsREject.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)

      Set bsLKick = New cvpmBallStack
          bsLKick.InitSaucer sw50, 50, 0, 45
          bsLKick.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)

  End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)
  If KeyDownHandler(keycode) Then Exit Sub
  If keycode = PlungerKey Then Plunger.Pullback:playsoundat "plungerpull", Plunger
  'if KeyCode = LeftTiltKey Then Nudge 90, 4
  'if KeyCode = RightTiltKey Then Nudge 270, 4
  'if KeyCode = CenterTiltKey Then Nudge 0, 4
   If KeyCode = LeftFlipperKey then FastFlips.FlipL True : Controller.Switch(23)=1 ' FastFlips.FlipUL True
     If KeyCode = RightFlipperKey then FastFlips.FlipR True : Controller.Switch(24)=1 'FastFlips.FlipUR True

'************************   Start Ball Control 1/3
  if keycode = 46 then        ' C Key
    If contball = 1 Then
      contball = 0
    Else
      contball = 1
    End If
  End If
  if keycode = 48 then        'B Key
    If bcboost = 1 Then
      bcboost = bcboostmulti
    Else
      bcboost = 1
    End If
  End If
  if keycode = 203 then bcleft = 1    ' Left Arrow
  if keycode = 200 then bcup = 1      ' Up Arrow
  if keycode = 208 then bcdown = 1    ' Down Arrow
  if keycode = 205 then bcright = 1   ' Right Arrow
'************************   End Ball Control 1/3

End Sub

Sub Table1_KeyUp(ByVal KeyCode)
  If KeyUpHandler(keycode) Then Exit Sub
  If keycode = PlungerKey Then Plunger.Fire:PlaySoundat "plunger", Plunger
     If KeyCode = LeftFlipperKey then FastFlips.FlipL False : Controller.Switch(23)=0'FastFlips.FlipUL False
     If KeyCode = RightFlipperKey then FastFlips.FlipR False : Controller.Switch(24)=0 'FastFlips.FlipUR False

'************************   Start Ball Control 2/3
  if keycode = 203 then bcleft = 0    ' Left Arrow
  if keycode = 200 then bcup = 0      ' Up Arrow
  if keycode = 208 then bcdown = 0    ' Down Arrow
  if keycode = 205 then bcright = 0   ' Right Arrow
'************************   End Ball Control 2/33

End Sub

'************************   Start Ball Control 3/3
Sub StartControl_Hit()
  Set ControlBall = ActiveBall
  contballinplay = true
End Sub

Sub StopControl_Hit()
  contballinplay = false
End Sub

Dim bcup, bcdown, bcleft, bcright, contball, contballinplay, ControlBall, bcboost
Dim bcvel, bcyveloffset, bcboostmulti

bcboost = 1   'Do Not Change - default setting
bcvel = 4   'Controls the speed of the ball movement
bcyveloffset = -0.01  'Offsets the force of gravity to keep the ball from drifting vertically on the table, should be negative
bcboostmulti = 3  'Boost multiplier to ball veloctiy (toggled with the B key)

Sub BallControl_Timer()
  If Contball and ContBallInPlay then
    If bcright = 1 Then
      ControlBall.velx = bcvel*bcboost
    ElseIf bcleft = 1 Then
      ControlBall.velx = - bcvel*bcboost
    Else
      ControlBall.velx=0
    End If

    If bcup = 1 Then
      ControlBall.vely = -bcvel*bcboost
    ElseIf bcdown = 1 Then
      ControlBall.vely = bcvel*bcboost
    Else
      ControlBall.vely= bcyveloffset
    End If
  End If
End Sub
'************************   End Ball Control 3/3


Sub Rubber100Wall_hit:vpmTimer.PulseSw 100:Rubber100.visible = 0::Rubber100a.visible = 1:Rubber100Wall.timerenabled = 1:End Sub
Sub Rubber100Wall_timer:Rubber100.visible = 1::Rubber100a.visible = 0: Rubber100Wall.timerenabled= 0:End Sub



'**********************************************************************************************************

 ' Drain hole and kickers
Sub Drain_Hit:bsTrough.addball me : playsoundat "drain", Drain : End Sub
Sub sw50_Hit:bsLKick.AddBall 0 : playsoundat "popper_ball", sw50: End Sub
Sub sw53_Hit:bsREject.AddBall 0 : playsoundat "popper_ball", sw53: End Sub

 'Rollovers
  Sub sw25_Hit:Controller.Switch(25) = 1:PlaySoundAt "rollover", ActiveBall:End Sub
  Sub sw25_UnHit:Controller.Switch(25) = 0:End Sub
  Sub sw33_Hit:Controller.Switch(33) = 1:PlaySoundAt "rollover", ActiveBall:End Sub
  Sub sw33_UnHit:Controller.Switch(33) = 0:End Sub
  Sub sw34_Hit:Controller.Switch(34) = 1:PlaySoundAt "rollover", ActiveBall:End Sub
  Sub sw34_UnHit:Controller.Switch(34) = 0:End Sub
  Sub sw35_Hit:Controller.Switch(35) = 1:PlaySoundAt "rollover", ActiveBall:End Sub
  Sub sw35_UnHit:Controller.Switch(35) = 0:End Sub
  Sub sw36_Hit:Controller.Switch(36) = 1:PlaySoundAt "rollover", ActiveBall:End Sub
  Sub sw36_UnHit:Controller.Switch(36) = 0:End Sub

  Sub sw39_Hit:Controller.Switch(39) = 1:PlaySoundAt "rollover", ActiveBall:End Sub
  Sub sw39_UnHit:Controller.Switch(39) = 0:End Sub
  Sub sw40_Hit:Controller.Switch(40) = 1:PlaySoundAt "rollover", ActiveBall:End Sub
  Sub sw40_UnHit:Controller.Switch(40) = 0:End Sub
  Sub sw43_Hit:Controller.Switch(43) = 1:PlaySoundAt "rollover", ActiveBall:End Sub
  Sub sw43_UnHit:Controller.Switch(43) = 0:End Sub
  Sub sw44_Hit:Controller.Switch(44) = 1:PlaySoundAt "rollover", ActiveBall:End Sub
  Sub sw44_UnHit:Controller.Switch(44) = 0:End Sub
  Sub sw45_Hit:Controller.Switch(45) = 1:PlaySoundAt "rollover", ActiveBall:End Sub
  Sub sw45_UnHit:Controller.Switch(45) = 0:End Sub
  Sub sw46_Hit:Controller.Switch(46) = 1:PlaySoundAt "rollover", ActiveBall:End Sub
  Sub sw46_UnHit:Controller.Switch(46) = 0:End Sub
  Sub sw47_Hit:Controller.Switch(47) = 1:PlaySoundAt "rollover", ActiveBall:End Sub
  Sub sw47_UnHit:Controller.Switch(47) = 0:End Sub

'Gate Trigges
Sub sw27_hit:vpmTimer.pulseSw 27 : End Sub
Sub sw28_hit:vpmTimer.pulseSw 28 : End Sub

'Ramp Triggers
  Sub sw37_Hit:vpmTimer.AddTimer 5,"vpmTimer.PulseSw 37'":PlaySoundAt "rollover", ActiveBall:End Sub
  Sub sw38_Hit:vpmTimer.AddTimer 5,"vpmTimer.PulseSw 38'":PlaySoundAt "rollover", ActiveBall:End Sub

'  Sub sw37_Hit:Controller.Switch(37) = 1:PlaySoundAt "rollover", ActiveBall:Debug.Print "37 Hit":End Sub
'  Sub sw37_UnHit:Controller.Switch(37) = 0:Debug.Print "37 UnHit":End Sub
'  Sub sw38_Hit:Controller.Switch(38) = 1:PlaySoundAt "rollover", ActiveBall:Debug.Print "38 Hit":End Sub
'  Sub sw38_UnHit:Controller.Switch(38) = 0:Debug.Print "38 UnHit":End Sub

'Ramp BackWall StandUp Targets
  Sub sw31_Hit:vpmTimer.PulseSw 31:PlaySoundAt "Target", ActiveBall:sw31a.transY = -10:Me.TimerEnabled = 1:End Sub
  Sub sw31_Timer:sw31a.transY = 0:Me.TimerEnabled = 0:End Sub
  Sub sw32_Hit:vpmTimer.PulseSw 32:PlaySoundAt "Target", ActiveBall:sw32a.transY = -10:Me.TimerEnabled = 1:End Sub
  Sub sw32_Timer:sw32a.transY = 0:Me.TimerEnabled = 0:End Sub


'Upper Right Corner
Sub sw48_Hit:vpmTimer.PulseSw 48:PlaySoundAt "fx_rubber2", ActiveBall:sw48a.transY = -10:sw48b.transX = -10:sw48c.transY = -10:Me.TimerEnabled = 1:End Sub
Sub sw48_Timer:sw48a.transY = 0:sw48b.transX = 0:sw48c.transY = 0:Me.TimerEnabled = 0:End Sub

'Captive Ball Triggers
Sub sw49_Hit:Controller.Switch(49) = 1:PlaySoundAt "rollover", ActiveBall:End Sub
Sub sw49_UnHit:Controller.Switch(49) = 0:End Sub
Sub sw52_Hit:Controller.Switch(52) = 1:PlaySoundAt "rollover", ActiveBall:End Sub
Sub sw52_UnHit:Controller.Switch(52) = 0:End Sub

'Center Ramp
Sub sw60_Hit:vpmTimer.PulseSw 60:ActiveBall.VelY = (ActiveBall.VelY * .5):PlaySoundAtBallVol "MetalRampCollide",1:End Sub
'Sub sw60_UnHit:Controller.Switch(60) = 0:Debug.Print "60 UnHit":End Sub

'******Stand Up Targets
Sub sw11_Hit:vpmTimer.PulseSw 11:End Sub
Sub sw12_Hit:vpmTimer.PulseSw 12:End Sub
Sub sw13_Hit:vpmTimer.PulseSw 13:End Sub
Sub sw17_Hit:vpmTimer.PulseSw 17:End Sub
Sub sw18_Hit:vpmTimer.PulseSw 18:End Sub
Sub sw19_Hit:vpmTimer.PulseSw 19:End Sub
Sub sw14_Hit:vpmTimer.PulseSw 14:End Sub
Sub sw15_Hit:vpmTimer.PulseSw 15:End Sub
Sub sw16_Hit:vpmTimer.PulseSw 16:End Sub
Sub sw20_Hit:vpmTimer.PulseSw 20:End Sub
Sub sw21_Hit:vpmTimer.PulseSw 21:End Sub
Sub sw22_Hit:vpmTimer.PulseSw 22:End Sub

Sub TRBallDrop1_Hit:PlaySoundAt "Balldrop", TRBallDrop1:End Sub
Sub TRBallDrop2_Hit:PlaySoundAt "Balldrop", TRBallDrop2:End Sub

'***************************************************
  ' Flames and animation
'***************************************************'

Dim FlamePos:FlamePos = 1


'*****Flames Motor/Animation
    Sub SolFlamesMotor(enabled)
      FlamesMotor.enabled = enabled
      SetLamp 150, (S14State OR FlamesMotor.Enabled)
    End Sub

    Sub FlamesMotor_Timer()
      FlamePos = FlamePos- 1
      If FlamePos = 0 then FlamePos = 200
      Flames.ImageA = "fanm" & FlamePos
      Flames.ImageB = "fanm" & FlamePos
    End Sub


'***************************************************
'       JP's VP10 Fading Lamps & Flashers
'       Based on PD's Fading Light System
' SetLamp 0 is Off
' SetLamp 1 is On
' fading for non opacity objects is 4 steps
'***************************************************

Dim LampState(200), FadingLevel(200)
Dim FlashSpeedUp(200), FlashSpeedDown(200), FlashMin(200), FlashMax(200), FlashLevel(200)

InitLamps()             ' turn off the lights and flashers and reset them to the default parameters
Set MotorCallback = GetRef("GameTimer")

Sub GameTimer
    Dim chgLamp, num, chg, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp) Then
        For ii = 0 To UBound(chgLamp)
            LampState(chgLamp(ii, 0) ) = chgLamp(ii, 1)       'keep the real state in an array
            FadingLevel(chgLamp(ii, 0) ) = chgLamp(ii, 1) + 4 'actual fading step
        Next
    End If
    UpdateLamps
  UpdateScores
  UpdateFlippers
  UpdateBallShadow
  RollingUpdate
End Sub


 Sub UpdateLamps
  NfadeL 4, l4
  NfadeL 5, l5
  NfadeL 6, l6
  NfadeL 7, l7
  NfadeL 8, l8
  NfadeL 9, l9
  NfadeLm 10, l10a
  NfadeL 10, l10
  NfadeL 11, l11
  NfadeL 12, l12
  NfadeL 13, l13
  NfadeL 14, l14
  NfadeL 15, l15
  NfadeL 16, l16
  NfadeL 17, l17
  NfadeL 18, l18
  NfadeL 19, l19
  NfadeL 20, l20
  NfadeL 21, l21
  NfadeL 22, l22
  NfadeL 23, l23
  NfadeL 24, l24
  NfadeL 25, l25
  NfadeLm 26, l26
  NFadeObj 26, Plug, "PostON", "Post"
  NfadeL 27, l27
  NfadeL 28, l28
  NfadeL 29, l29
  NfadeL 30, l30
  NfadeL 31, l31
  NfadeL 32, l32
  NfadeLm 33, l33
  FadeObj 33, leftrail_prim, "leftrailLIT", "leftrailb", "leftraila", "leftrail"
  NfadeLm 34, l34
  FadeObj 34, rightrail_prim, "rightrailLIT", "rightrailb", "rightraila", "rightrail"
  SetLamp 135, LampState(35) OR LampState(125)
  Flashm 135, Building3
  'Flash 135, Building3Smoke
  NfadeL 37, l37
  NfadeL 38, l38
  NfadeL 39, l39
  NfadeL 40, l40
  NfadeL 41, l41
  NfadeL 42, l42
  NfadeL 44, l44
  NfadeL 45, l45
  NfadeL 46, l46
  NfadeL 47, l47
  NfadeL 48, l48
  SetLamp 136, LampState(49) OR LampState(126)
  Flashm 136, Building2
  'Flash 136, Building2Smoke
  NfadeL 61, l61
  NfadeL 62, l62
  NFadeLm 53, CenterBuildingGlow1
  NFadeLm 53, CenterBuildingGlow2
  NFadeLm 53, CenterBuildingGlow3
  NFadeLm 53, LightCenterBuildingb
  NFadeL 53, LightCenterBuildinga
  'Flash 53, centerBuilding
  Flashm 55, BuildingBack
  Flash 55, Building6Glow
  SetLamp 138, LampState(57) OR LampState(130)
  Flash 138, Building5
  SetLamp 137, LampState(59) OR LampState(131)
  Flash 137, Building8
  NfadeL 63, l63
  NfadeL 64, l64

'Solenoid Controlled Flashers
  NfadeLm 116, W1
  NfadeLm 116, W2
  NfadeLm 116, W3
  NfadeLm 116, W4
  NfadeLm 116, W5
  NfadeLm 116, W6
  NfadeLm 116, W7
  NfadeLm 116, W8
  NfadeLm 116, W9
  NfadeLm 116, W10
  NfadeLm 116, W11
  NfadeL 116, W12
  NFadeLm 125, LightBuilding3
  NFadeLm 125, LightBuilding3a
  'NFadeLm 125, f125a
  'NFadeL 125, f125b
  'FadeObjm 125, Building3, "b3On", "b3A", "b3B", "b3"
  NFadeLm 126, Lightbuilding2
  NFadeLm 126, Lightbuilding2a
  'NFadeLm 126, f126a
  'NFadeL 126, f126b

  'FadeObjm 126, Building2, "b2On", "b2A", "b2B", "b2"

  NFadeL 127, LightCenterBuilding

'FadeObj 127, centerBuilding, "centerOn", "centerA", "centerB", "center"
  Flash 150, Flames
  NfadeL 128, F128

  NFadeLm 130, LightBuilding8
  NFadeL 130, LightBuilding8a
  NFadeLm 131, LightBuilding5
  NFadeL 131, LightBuilding5a

  'Flasher GI

  Flashm 151, gi13
  Flashm 151, gi14
  Flashm 151, gi15
  Flash 151, gi16
  ' Building 4 seems to come on with GI
  Flash 151, Building4


  'Buildings 6 and 7 always seems to be on per videos, they are supposed
  'to illuminate the skills shot.
    'Buiding6 has additional lamps tied to 55, we'll leave the glow for that.
  SetLamp 152, 1   '??
  Flashm 152, BuildingSkill
  NFadeLm 152, SkillGlow1
  NFadeLm 152, SkillGlow2
  Flash 152, Building7



 End Sub

' div lamp subs

Sub InitLamps()
    Dim x
    For x = 0 to 200
        LampState(x) = 0        ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4      ' used to track the fading state
        FlashSpeedUp(x) = 0.2   ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.08 ' slower speed when turning off the flasher
        FlashMax(x) = 1         ' the maximum value when on, usually 1
        FlashMin(x) = 0         ' the minimum value when off, usually 0
        FlashLevel(x) = 0       ' the intensity of the flashers, usually from 0 to 1
    Next
  FlashSpeedUp(150)=0.2
  FlashSpeedDown(150)=0.05
End Sub

Sub AllLampsOff
    Dim x
    For x = 0 to 200
        SetLamp x, 0
    Next
End Sub

Sub SetLamp(nr, value)
    If value <> LampState(nr) Then
        LampState(nr) = abs(value)
        FadingLevel(nr) = abs(value) + 4
    End If
End Sub


Sub FadeFlames(nr)
    Select Case FadingLevel(nr)
        Case 2:FlamesFade = 0:FadingLevel(nr) = 0 'Off
        Case 3:FlamesFade = 1:FadingLevel(nr) = 2 'fading...
        Case 4:FlamesFade = 2:FadingLevel(nr) = 3 'fading...
        Case 5:FlamesFade = 3:FadingLevel(nr) = 1 'ON
    End Select
End Sub
' Lights: used for VP10 standard lights, the fading is handled by VP itself



Sub NFadeL(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.state = 0:FadingLevel(nr) = 0
        Case 5:object.state = 1:FadingLevel(nr) = 1
    End Select
End Sub

Sub NFadeLm(nr, object) ' used for multiple lights
    Select Case FadingLevel(nr)
        Case 4:object.state = 0
        Case 5:object.state = 1
    End Select
End Sub

'Lights, Ramps & Primitives used as 4 step fading lights
'a,b,c,d are the images used from on to off

Sub FadeObj(nr, object, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 4:object.image = b:FadingLevel(nr) = 6                   'fading to off...
        Case 5:object.image = a:FadingLevel(nr) = 1                   'ON
        Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1             'wait
        Case 9:object.image = c:FadingLevel(nr) = FadingLevel(nr) + 1 'fading...
        Case 10, 11, 12:FadingLevel(nr) = FadingLevel(nr) + 1         'wait
        Case 13:object.image = d:FadingLevel(nr) = 0                  'Off
    End Select
End Sub

Sub FadeObjm(nr, object, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 4:object.image = b
        Case 5:object.image = a
        Case 9:object.image = c
        Case 13:object.image = d
    End Select
End Sub

Sub NFadeObj(nr, object, a, b)
    Select Case FadingLevel(nr)
        Case 4:object.image = b:FadingLevel(nr) = 0 'off
        Case 5:object.image = a:FadingLevel(nr) = 1 'on
    End Select
End Sub

Sub NFadeObjm(nr, object, a, b)
    Select Case FadingLevel(nr)
        Case 4:object.image = b
        Case 5:object.image = a
    End Select
End Sub

' Flasher / primitive objects

Sub Flash(nr, object)
    Select Case FadingLevel(nr)
        Case 4 'off
            FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown(nr)
            If FlashLevel(nr) < FlashMin(nr) Then
                FlashLevel(nr) = FlashMin(nr)
                FadingLevel(nr) = 0 'completely off
            End if
        Case 5 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
            If FlashLevel(nr) > FlashMax(nr) Then
                FlashLevel(nr) = FlashMax(nr)
                FadingLevel(nr) = 1 'completely on
            End if
    End Select
  Flashm nr, Object
End Sub

Sub Flashm(nr, object) 'multiple flashers, it just sets the flashlevel
    if TypeName(object) = "Primitive" Then
    object.BlendDisableLighting = FlashLevel(nr)
  else
    Object.IntensityScale = FlashLevel(nr)
  end if
End Sub

 'Reels
Sub FadeReel(nr, reel)
    Select Case FadingLevel(nr)
        Case 2:FadingLevel(nr) = 0
        Case 3:FadingLevel(nr) = 2
        Case 4:reel.Visible = 0:FadingLevel(nr) = 3
        Case 5:reel.Visible = 1:FadingLevel(nr) = 1
    End Select
End Sub

 'Inverted Reels
Sub FadeIReel(nr, reel)
    Select Case FadingLevel(nr)
        Case 2:FadingLevel(nr) = 0
        Case 3:FadingLevel(nr) = 2
        Case 4:reel.Visible = 1:FadingLevel(nr) = 3
        Case 5:reel.Visible = 0:FadingLevel(nr) = 1
    End Select
End Sub


'**********************************************************************************************************
'Digital Display
'**********************************************************************************************************
 Dim Digits(28)
 Digits(0)=Array(a00, a05, a0c, a0d, a08, a01, a06, a0f, a02, a03, a04, a07, a0b, a0a, a09, a0e)
 Digits(1)=Array(a10, a15, a1c, a1d, a18, a11, a16, a1f, a12, a13, a14, a17, a1b, a1a, a19, a1e)
 Digits(2)=Array(a20, a25, a2c, a2d, a28, a21, a26, a2f, a22, a23, a24, a27, a2b, a2a, a29, a2e)
 Digits(3)=Array(a30, a35, a3c, a3d, a38, a31, a36, a3f, a32, a33, a34, a37, a3b, a3a, a39, a3e)
 Digits(4)=Array(a40, a45, a4c, a4d, a48, a41, a46, a4f, a42, a43, a44, a47, a4b, a4a, a49, a4e)
 Digits(5)=Array(a50, a55, a5c, a5d, a58, a51, a56, a5f, a52, a53, a54, a57, a5b, a5a, a59, a5e)
 Digits(6)=Array(a60, a65, a6c, a6d, a68, a61, a66, a6f, a62, a63, a64, a67, a6b, a6a, a69, a6e)

 Digits(7)=Array(a70, a75, a7c, a7d, a78, a71, a76, a7f, a72, a73, a74, a77, a7b, a7a, a79, a7e)
 Digits(8)=Array(a80, a85, a8c, a8d, a88, a81, a86, a8f, a82, a83, a84, a87, a8b, a8a, a89, a8e)
 Digits(9)=Array(a90, a95, a9c, a9d, a98, a91, a96, a9f, a92, a93, a94, a97, a9b, a9a, a99, a9e)
 Digits(10)=Array(aa0, aa5, aac, aad, aa8, aa1, aa6, aaf, aa2, aa3, aa4, aa7, aab, aaa, aa9, aae)
 Digits(11)=Array(ab0, ab5, abc, abd, ab8, ab1, ab6, abf, ab2, ab3, ab4, ab7, abb, aba, ab9, abe)
 Digits(12)=Array(ac0, ac5, acc, acd, ac8, ac1, ac6, acf, ac2, ac3, ac4, ac7, acb, aca, ac9, ace)
 Digits(13)=Array(ad0, ad5, adc, add, ad8, ad1, ad6, adf, ad2, ad3, ad4, ad7, adb, ada, ad9, ade)

 ' 3rd Player
Digits(14) = Array(LED150,LED151,LED152,LED153,LED154,LED155,LED156)
Digits(15) = Array(LED160,LED161,LED162,LED163,LED164,LED165,LED166)
Digits(16) = Array(LED170,LED171,LED172,LED173,LED174,LED175,LED176)
Digits(17) = Array(LED180,LED181,LED182,LED183,LED184,LED185,LED186)
Digits(18) = Array(LED190,LED191,LED192,LED193,LED194,LED195,LED196)
Digits(19) = Array(LED200,LED201,LED202,LED203,LED204,LED205,LED206)
Digits(20) = Array(LED210,LED211,LED212,LED213,LED214,LED215,LED216)

' 4th Player
Digits(21) = Array(LED220,LED221,LED222,LED223,LED224,LED225,LED226)
Digits(22) = Array(LED230,LED231,LED232,LED233,LED234,LED235,LED236)
Digits(23) = Array(LED240,LED241,LED242,LED243,LED244,LED245,LED246)
Digits(24) = Array(LED250,LED251,LED252,LED253,LED254,LED255,LED256)
Digits(25) = Array(LED260,LED261,LED262,LED263,LED264,LED265,LED266)
Digits(26) = Array(LED270,LED271,LED272,LED273,LED274,LED275,LED276)
Digits(27) = Array(LED280,LED281,LED282,LED283,LED284,LED285,LED286)


 Sub UpdateScores
    Dim ChgLED, ii, jj, num, chg, stat, obj, b, x
    ChgLED=Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
    If DesktopMode = True Then
       For ii=0 To UBound(chgLED)
          num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)
      if (num < 28) then
              For Each obj In Digits(num)
                   If chg And 1 Then obj.State=stat And 1
                   chg=chg\2 : stat=stat\2
                  Next
      Else
             end if
        Next
     end if
    End If
 End Sub

'**********************************************************************************************************
'**********************************************************************************************************
' Start of VPX functions
'**********************************************************************************************************
'**********************************************************************************************************

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep, RStep1, Lstep1

Sub RightSlingShot_Slingshot
  vpmTimer.PulseSw 58
    PlaySoundAt SoundFX("right_slingshot",DOFContactors), GI4
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0:
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  vpmTimer.PulseSw 57
    PlaySoundAt SoundFX("left_slingshot",DOFContactors),GI2
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot1_Slingshot
  vpmTimer.PulseSw 62
    PlaySoundAt SoundFX("right_slingshot",DOFContactors), GI4
    R1Sling.Visible = 0
    R1Sling1.Visible = 1
    sling3.TransZ = -20
    RStep1 = 0
    RightSlingShot1.TimerEnabled = 1
End Sub

Sub RightSlingShot1_Timer
    Select Case RStep1
        Case 3:R1SLing1.Visible = 0:R1SLing2.Visible = 1:sling3.TransZ = -10
        Case 4:R1SLing2.Visible = 0:R1SLing.Visible = 1:sling3.TransZ = 0:RightSlingShot1.TimerEnabled = 0:
    End Select
    RStep1 = RStep1 + 1
End Sub

Sub LeftSlingShot1_Slingshot
  vpmTimer.PulseSw 61
    PlaySoundAt SoundFX("left_slingshot",DOFContactors),GI2
    L1Sling.Visible = 0
    L1Sling1.Visible = 1
    sling4.TransZ = -20
    LStep1 = 0
    LeftSlingShot1.TimerEnabled = 1
End Sub

Sub LeftSlingShot1_Timer
    Select Case LStep1
        Case 3:L1SLing1.Visible = 0:L1SLing2.Visible = 1:sling4.TransZ = -10
        Case 4:L1SLing2.Visible = 0:L1SLing.Visible = 1:sling4.TransZ = 0:LeftSlingShot1.TimerEnabled = 0
    End Select
    LStep1 = LStep1 + 1
End Sub
'*********************************************************************
'                 Positional Sound Playback Functions
'*********************************************************************

' Similar subroutines that are less complicated to use (e.g. simply use standard parameters for the PlaySound call)
Sub PlaySoundAt(soundname, tableobj)
    PlaySound soundname, 1, 1, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtVol(soundname, tableobj, vol)
    PlaySound soundname, 1, vol, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
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
    Vol = Csng(BallVel(ball) ^2 / 400)
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

Const tnob = 3 ' total number of balls
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
        If BallVel(BOT(b) ) > 1 Then
            rolling(b) = True
      if BOT(b).z < 30 then
        If Table1.VersionMinor > 3 OR Table1.VersionMajor > 10 Then
          PlaySound "fx_ballrolling" & b, -1, Vol(BOT(b) ) * .5 , AudioPan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b))
        Else
          PlaySound "fx_ballrolling" & b, -1, Vol(BOT(b) ) * .5 , AudioPan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0
        End If
      Else ' ball in air, probably on plastic.
        If Table1.VersionMinor > 3 OR Table1.VersionMajor > 10 Then
          PlaySound "fx_ballrolling" & b, -1, Vol(BOT(b) ) * .2  , AudioPan(BOT(b) ), 0, Pitch(BOT(b) ) + 40000, 1, 0, AudioFade(BOT(b))
        Else
          PlaySound "fx_ballrolling" & b, -1, Vol(BOT(b) ) * .2 , AudioPan(BOT(b) ), 0, Pitch(BOT(b) ) + 40000, 1, 0
        End If
      End If
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
  PlaySound "fx_collide", 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub


'*****************************************
' ninuzzu's FLIPPER SHADOWS
'*****************************************

sub UpdateFlippers()
  FlipperLSh.RotZ = LeftFlipper.currentangle
  FlipperRSh.RotZ = RightFlipper.currentangle
  batleft.objrotz = LeftFlipper.CurrentAngle + 1
  batright.objrotz = RightFlipper.CurrentAngle - 1
End Sub

'*****************************************
' ninuzzu's BALL SHADOW
'*****************************************
Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5)
ReDim BallHangTime(tnob)

Sub UpdateBallShadow()
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
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/20)) + 6
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/20)) - 6
        End If
        ballShadow(b).Y = BOT(b).Y + 12
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    ' Ball can "rest" on the horeshoe ramp.   Ball sticking hack - if it stops for more than 1 second give it a nudge to the right.
    if BOT(b).Y < 130 and BallVel(BOT(b)) < .01 Then
      if Timer > BallHangTime(b) + 1 then
        if Bot(B).X > 435 then
          BOT(b).VelX = .6
        Else
          Bot(b).VelX = -.6
        end if
      end if
    Else
      BallHangTime(b)=Timer
    end if
    Next
End Sub


'**********************************************************************************************************
' Random Ramp Bumps by RustyCardores & DJRobX - Best used to compliment Rusty & Rob's Raised Ramp RollingBall Script
' Switches added to ramps in key bend locations and called from these collections(idx).
'**********************************************************************************************************

Dim NextHit:NextHit = 0

Sub PlasticRampbumps_Hit(idx)
  RandomBump 1, Pitch(ActiveBall)
End Sub

Sub MetalRampbumps_Hit(idx)
  RandomBump 1, 6000
End Sub

Sub RandomBump(voladj, freq)
  dim BumpSnd:BumpSnd= "fx_rampbump" & CStr(Int(Rnd*7)+1)
  If Table1.VersionMinor > 3 OR Table1.VersionMajor > 10 Then
    PlaySound BumpSnd, 0, Vol(ActiveBall)*voladj, AudioPan(ActiveBall), 0, freq, 0, 1, AudioFade(ActiveBall)
  Else
    PlaySound BumpSnd, 0, Vol(ActiveBall)*voladj, AudioPan(ActiveBall), 0, freq, 0, 1
  End If
End Sub

'Set all as per ball position & speed, but Vol Multiplier may be used eg; PlaySoundAtBallVol "sound",3

Sub PlaySoundAtBallVol(sound, VolMult)
  If Table1.VersionMinor > 3 OR Table1.VersionMajor > 10 Then
    PlaySound sound, 0, Vol(ActiveBall) * VolMult, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
  Else
    PlaySound sound, 0, Vol(ActiveBall) * VolMult, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1
  End If
End Sub
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
  if BallVel(ActiveBall) > .3 and Timer > NextHit then
    RandomBump 2, 20000
    NextHit = Timer + .03 + (Rnd * .2)
  end if
End Sub

Sub Metals_Medium_Hit (idx)
  if BallVel(ActiveBall) > .3 and Timer > NextHit then
    PlaySound "metalhit_medium", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    NextHit = Timer + .03 + (Rnd * .2)
  end if
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

Sub RubbersBands_Hit(idx)
  If BallVel(ActiveBall) > 20 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall)*2, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  else
    RandomSoundRubber()
  End If
End Sub

Sub RubbersPosts_Hit(idx)
  If BallVel(ActiveBall) > 16 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall)*2, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  else
    RandomSoundRubber()
  End If
End Sub

Sub RubbersRingsSleeves_Hit(idx)
  If BallVel(ActiveBall) > 16 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall)*2, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  else
    RandomSoundRubber()
  End If
End Sub

Sub RandomSoundRubber()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*2, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall)*2, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall)*2, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

Sub LeftFlipper_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RightFlipper_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub centerBuilding_hit()
  RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall)*2, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall)*2, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall)*2, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub


'**********************************************************************************************************
'cFastFlips by nFozzy
'**********************************************************************************************************
dim FastFlips
Set FastFlips = new cFastFlips
with FastFlips
  .CallBackL = "SolLflipper"  'Point these to flipper subs
  .CallBackR = "SolRflipper"  '...
' .CallBackUL = "SolULflipper"'...(upper flippers, if needed)
' .CallBackUR = "SolURflipper"'...
  .TiltObjects = True 'Optional, if True calls vpmnudge.solgameon automatically. IF YOU GET A LINE 1 ERROR, DISABLE THIS! (or setup vpmNudge.TiltObj!)
' .InitDelay "FastFlips", 100     'Optional, if > 0 adds some compensation for solenoid jitter (occasional problem on Bram Stoker's Dracula)
' .DebugOn = False    'Debug, always-on flippers. Call FastFlips.DebugOn True or False in debugger to enable/disable.
end with

Class cFastFlips
  Public TiltObjects, DebugOn
  Private SubL, SubUL, SubR, SubUR, FlippersEnabled, Delay, LagCompensation, Name

  Private Sub Class_Initialize()
    Delay = 0 : FlippersEnabled = False : DebugOn = False : LagCompensation = False
  End Sub

  'set callbacks
  Public Property Let CallBackL(aInput)  : Set SubL  = GetRef(aInput) : End Property
  Public Property Let CallBackUL(aInput) : Set SubUL = GetRef(aInput) : End Property
  Public Property Let CallBackR(aInput)  : Set SubR  = GetRef(aInput) : End Property
  Public Property Let CallBackUR(aInput) : Set SubUR = GetRef(aInput) : End Property
  Public Sub InitDelay(aName, aDelay) : Name = aName : delay = aDelay : End Sub 'Create Delay

  'call callbacks
  Public Sub FlipL(aEnabled)
    if not FlippersEnabled and not DebugOn then Exit Sub
    subL aEnabled
  End Sub

  Public Sub FlipR(aEnabled)
    if not FlippersEnabled and not DebugOn then Exit Sub
    subR aEnabled
  End Sub

  Public Sub FlipUL(aEnabled)
    if not FlippersEnabled and not DebugOn then Exit Sub
    subUL aEnabled
  End Sub

  Public Sub FlipUR(aEnabled)
    if not FlippersEnabled and not DebugOn then Exit Sub
    subUR aEnabled
  End Sub

  Public Sub TiltSol(aEnabled)  'Handle solenoid / Delay (if delayinit)
    if delay > 0 and not aEnabled then  'handle delay
      vpmtimer.addtimer Delay, Name & ".FireDelay" & "'"
      LagCompensation = True
    else
      if Delay > 0 then LagCompensation = False
      EnableFlippers(aEnabled)
    end if
  End Sub

  Sub FireDelay() : if LagCompensation then EnableFlippers False End If : End Sub

  Private Sub EnableFlippers(aEnabled)
    FlippersEnabled = aEnabled
    if TiltObjects then vpmnudge.solgameon aEnabled
    If Not aEnabled then
      subL False
      subR False
      if not IsEmpty(subUL) then subUL False
      if not IsEmpty(subUR) then subUR False
    End If
  End Sub

End Class
'**********************************************************************************************************
'**********************************************************************************************************
' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


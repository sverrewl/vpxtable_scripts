Option Explicit
Randomize

' !!!!! Before release: check all FIXME and solve them !!!!


'///////////////////////-----General Sound Options-----///////////////////////
'//  VolumeDial:
'//  VolumeDial is the actual global volume multiplier for the mechanical sounds.
'//  Values smaller than 1 will decrease mechanical sounds volume.
'//  Recommended values should be no greater than 1.
Const VolumeDial = 1

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

Const cSingleLFlip = 0
Const cSingleRFlip = 0

Const cGameName="wrlok_l3",UseSolenoids=2,UseLamps=0,UseGI=0,SSolenoidOn="Solenoid",SSolenoidOff="SolOff"
Const SCoin=""

LoadVPM "01560000", "S7.VBS", 3.26

Dim DesktopMode: DesktopMode = Table1.ShowDT

' FIXME VB
If DesktopMode = True Then 'Show Desktop components
displaytimer.enabled=1
'lockdown.visible=1
Else
displaytimer.enabled=0
'lockdown.visible=0
End if

' using table width and height in script slows down the performance
dim tablewidth: tablewidth = table1.width
dim tableheight: tableheight = table1.height

Dim BallShadows: Ballshadows=1      '******************set to 1 to turn on Ball shadows
Dim FlipperShadows: FlipperShadows=1  '***********set to 1 to turn on Flipper shadows

Dim EnableBallControl
EnableBallControl = false 'Change to true to enable manual ball control (or press C in-game) via the arrow keys and B (boost movement) keys

'Solenoid Call backs
'**********************************************************************************************************
SolCallback(1)  = "Sol1DropUp"
SolCallback(2)  = "Sol2DropUp"
SolCallback(3)  = "Sol3DropUp"
SolCallback(4)  = "SolRelease"
SolCallback(5)  = "dt1drop"
SolCallback(6)  = "dt2drop"
SolCallback(7)  = "dt3drop"
SolCallback(8)  = "SolGi"
SolCallback(15) = "SolKnocker"
SolCallback(23) = "vpmNudge.SolGameOn"

Sub SolGi(Enabled)
  Debug.Print "SolGi " & Enabled
  If Enabled Then
    Lampz.State(101) = 1
  Else
    Lampz.State(101) = 0
  End If
End Sub

'******************************************************
'         FLIPPERS
'******************************************************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Const ReflipAngle = 20

Sub SolLFlipper(Enabled)
  If Enabled Then
    LF.Fire

    If leftflipper.currentangle < leftflipper.endangle + ReflipAngle Then
      RandomSoundReflipUpLeft LeftFlipper
    Else
      SoundFlipperUpAttackLeft LeftFlipper
      RandomSoundFlipperUpLeft LeftFlipper
    End If
  Else
    LeftFlipper.RotateToStart
    If LeftFlipper.currentangle < LeftFlipper.startAngle - 5 Then
      RandomSoundFlipperDownLeft LeftFlipper
    End If
    FlipperLeftHitParm = FlipperUpSoundLevel
    End If
End Sub

Sub SolRFlipper(Enabled)
  If Enabled Then
    RF.Fire
'   RightFlipper001.rotatetoend
    If rightflipper.currentangle > rightflipper.endangle - ReflipAngle Then
      RandomSoundReflipUpRight RightFlipper
    Else
      SoundFlipperUpAttackRight RightFlipper
      RandomSoundFlipperUpRight RightFlipper
    End If
  Else
    RightFlipper.RotateToStart
'   rightflipper001.rotatetostart
    If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
      RandomSoundFlipperDownRight RightFlipper
    End If
    FlipperRightHitParm = FlipperUpSoundLevel
  End If
End Sub

'******************************************************
'         KNOCKER
'******************************************************
Sub SolKnocker(Enabled)
  If enabled Then
    KnockerSolenoid 'Add knocker position object
  End If
End Sub

'******************************************************
'       SOLENOID CONTROLLED TOYS
'******************************************************

Sub dt1drop(enabled)
  If enabled Then
    DTRaise 33
    PlaySoundAt SoundFX(DTResetSound,DOFContactors), DT3_1_BM_Dark_Room
  End If
 End Sub

Sub dt2drop(enabled)
  If enabled Then
    DTRaise 34
    PlaySoundAt SoundFX(DTResetSound,DOFContactors), DT3_2_BM_Dark_Room
  End If
End Sub

Sub dt3drop(enabled)
  If enabled Then
    DTRaise 35
    PlaySoundAt SoundFX(DTResetSound,DOFContactors), DT3_3_BM_Dark_Room
  End If
End Sub

Sub Sol1DropUp(enabled)
  if enabled then
    PlaySoundAt SoundFX(DTResetSound,DOFContactors), DT1_2_BM_Dark_Room
    DTRaise 27
    DTRaise 28
    DTRaise 29
  end if
End Sub

Sub Sol2DropUp(enabled)
  if enabled then
    PlaySoundAt SoundFX(DTResetSound,DOFContactors), DT2_2_BM_Dark_Room
    DTRaise 30
    DTRaise 31
    DTRaise 32
  end if
End Sub

Sub Sol3DropUp(enabled)
  if enabled then
    PlaySoundAt SoundFX(DTResetSound,DOFContactors), DT3_2_BM_Dark_Room
    DTRaise 33
    DTRaise 34
    DTRaise 35
  end if
End Sub



'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************

Dim bsTrough, bslsaucer

Sub Table1_Init
  vpmInit Me
' On Error Resume Next
  With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine = "Meteor (Stern 1978)... or is it?"
    .HandleMechanics=0
    .HandleKeyboard=0
    .ShowDMDOnly=1
    .ShowFrame=0
    .ShowTitle=0
        .hidden = 1
    If Err Then MsgBox Err.Description
  End With
  On Error Goto 0
    Controller.SolMask(0)=0
      vpmTimer.AddTimer 2000,"Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the timer to renable all the solenoids after 2 seconds
    Controller.Run
  If Err Then MsgBox Err.Description
  On Error Goto 0

  PinMAMETimer.Interval=PinMAMEInterval
  PinMAMETimer.Enabled=1
  vpmNudge.TiltSwitch = 1
  vpmNudge.Sensitivity = 3
  vpmNudge.TiltObj = Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingshot)

  ' Initialize dark/lit room level
  Dim x: x = LoadValue(cGameName, "LIGHT")
    If x <> "" Then LightLevel = CInt(x) Else LightLevel = 50
  UpdateLightLevel

  ' Hide all lights except from ball reflections
  For Each x in GetElements
    If typename(x) = "Light" Then x.Visible = False
  Next

  ' FIXME Light up the GI (it stays off otherwise)
  SolGi True
End Sub

  Drain.Createball
  Controller.Switch(9) = 1

Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)
  If keycode = LeftFlipperKey Then FlipperActivate LeftFlipper, LFPress
  If keycode = RightFlipperKey Then FlipperActivate RightFlipper, RFPress: Controller.Switch(36) = 1

  If keycode = LeftTiltKey Then Nudge 90, 5:SoundNudgeLeft()
  If keycode = RightTiltKey Then Nudge 270, 5:SoundNudgeRight()
  If keycode = CenterTiltKey Then Nudge 0, 3:SoundNudgeCenter()

  If keycode = LeftMagnaSave Then
    LightLevel = LightLevel - 10
    If LightLevel < 0 Then LightLevel = 100
    SaveValue cGameName, "LIGHT", LightLevel
    UpdateLightLevel
  End If
  If keycode = RightMagnaSave Then
    LightLevel = LightLevel + 10
    If LightLevel > 100 Then LightLevel = 0
    SaveValue cGameName, "LIGHT", LightLevel
    UpdateLightLevel
  End If

  If keycode = keyInsertCoin1 or keycode = keyInsertCoin2 or keycode = keyInsertCoin3 or keycode = keyInsertCoin4 Then
    Select Case Int(rnd*3)
      Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
      Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
      Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25

    End Select
  End If

  If KeyCode = PlungerKey Then Plunger.Pullback:SoundPlungerPull()

  If KeyDownHandler(keycode) Then Exit Sub

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

Sub Table1_KeyUp(ByVal KeyCode)

  If keycode = LeftFlipperKey Then FlipperDeActivate LeftFlipper, LFPress
  If keycode = RightFlipperKey Then FlipperDeActivate RightFlipper, RFPress: Controller.Switch(36) = 0

  If KeyCode = PlungerKey Then
    Plunger.Fire
    If plungelane=1 Then
      SoundPlungerReleaseBall()     'Plunger release sound when there is a ball in shooter lane
    Else
      SoundPlungerReleaseNoBall()     'Plunger release sound when there is no ball in shooter lane
    End If
  End If

  If KeyUpHandler(keycode) Then Exit Sub

    'Manual Ball Control
  If EnableBallControl = 1 Then
    If keycode = 203 Then BCleft = 0  ' Left Arrow
    If keycode = 200 Then BCup = 0    ' Up Arrow
    If keycode = 208 Then BCdown = 0  ' Down Arrow
    If keycode = 205 Then BCright = 0 ' Right Arrow
  End If

End Sub

dim plungelane

Sub ballatplunger_hit
  plungelane=1
  plungewire.transz=-6
End Sub

Sub ballatplunger_unhit
  plungelane=0
  plungewire.transz=0
End Sub

'******************************************************
'     FLIPPER TRICKS
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
Const EOSTnew = 1
Const EOSAnew = 1
Const EOSRampup = 0
Dim SOSRampup
SOSRampup = 2.5
Const LiveCatch = 16
Const LiveElasticity = 0.45
Const SOSEM = 0.815
Const EOSReturn = 0.045

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
  Dir = Flipper.startangle/Abs(Flipper.startangle)  '-1 for Right Flipper

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
Const LiveDistanceMax = 114  'maximum distance in vp units from flipper base live catch dampening will occur (tip protection)

Sub CheckLiveCatch(ball, Flipper, FCount, parm) 'Experimental new live catch
  Dim Dir
  Dir = Flipper.startangle/Abs(Flipper.startangle)    '-1 for Right Flipper
  Dim LiveCatchBounce                             'If live catch is not perfect, it won't freeze ball totally
  Dim CatchTime : CatchTime = GameTime - FCount

  if CatchTime <= LiveCatch and parm > 6 and ABS(Flipper.x - ball.x) > LiveDistanceMin and ABS(Flipper.x - ball.x) < LiveDistanceMax Then
    if CatchTime <= LiveCatch*0.5 Then            'Perfect catch only when catch time happens in the beginning of the window
      LiveCatchBounce = 0
    else
      LiveCatchBounce = Abs((LiveCatch/2) - CatchTime)  'Partial catch when catch happens a bit late
    end If

    If LiveCatchBounce = 0 and ball.velx * Dir > 0 Then ball.velx = 0
    ball.vely = LiveCatchBounce * (32 / LiveCatch) ' Multiplier for inaccuracy bounce
    ball.angmomx= 0
    ball.angmomy= 0
    ball.angmomz= 0
  End If
End Sub

Sub LeftFlipper_Collide(parm)
  CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
  LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
  CheckLiveCatch Activeball, RightFlipper, RFCount, parm
  RightFlipperCollide parm
End Sub

'**********************************************************************************************************
'**********************************************************************************************************

 'switches
'**********************************************************************************************************

Sub Drain_Hit()
  Controller.Switch(9) = 1
  RandomSoundDrain drain
End Sub

Sub Drain_UnHit()  'Drain
  Controller.Switch(9) = 0
End Sub

Sub SolRelease(enabled)
  If enabled Then
    If Drain.BallCntOver = 0 Then
      PlaySoundAt SoundFX("solon",DOFContactors), Drain
    Else
      PlaySoundAt SoundFX("gamestart",DOFContactors), Drain
    End If
    Drain.kick 66, 20
  End If
End Sub

'Drop Targets
Sub Sw27_Hit:DTHit 27:End Sub
Sub Sw28_Hit:DTHit 28:End Sub
Sub Sw29_Hit:DTHit 29:End Sub
Sub Sw30_Hit:DTHit 30:End Sub
Sub Sw31_Hit:DTHit 31:End Sub
Sub Sw32_Hit:DTHit 32:End Sub
Sub Sw33_Hit:DTHit 33:End Sub
Sub Sw34_Hit:DTHit 34:End Sub
Sub Sw35_Hit:DTHit 35:End Sub

'Stand Up Targets
 Sub Tsw19_Hit: PrimStandupTgtHit 19, Tsw19, HT_BM_Dark_Room: vpmTimer.PulseSw 19 : End Sub
 Sub Tsw19_Timer: PrimStandupTgtMove 19, Tsw19, HT_BM_Dark_Room: End Sub

'Bumpers
Sub Bumper1_Hit
  vpmTimer.PulseSw 13
  RandomSoundBumperMiddle BR1_BM_Dark_Room
  BR1_BM_Dark_Room.rotx = skirtAX(me,Activeball)
  BR1_BM_Dark_Room.roty = skirtAY(me,Activeball)
  UpdateBR1
  me.timerinterval = 150
  me.timerenabled=1
End Sub

sub bumper1_timer
  BR1_BM_Dark_Room.rotx = 0
  BR1_BM_Dark_Room.roty = 0
  UpdateBR1
  me.timerenabled=0
end sub

Sub Bumper2_Hit
  vpmTimer.PulseSw 14
  RandomSoundBumperMiddle BR2_BM_Dark_Room
  BR2_BM_Dark_Room.roty=skirtAY(me,Activeball)
  BR2_BM_Dark_Room.rotx=skirtAX(me,Activeball)
  UpdateBR2
  me.timerinterval = 150
  me.timerenabled=1
End Sub

sub bumper2_timer
  BR2_BM_Dark_Room.rotx = 0
  BR2_BM_Dark_Room.roty = 0
  UpdateBR2
  me.timerenabled=0
end sub

Sub Bumper3_Hit
  vpmTimer.PulseSw 15
  RandomSoundBumperMiddle BR3_BM_Dark_Room
  BR3_BM_Dark_Room.roty=skirtAY(me,Activeball)
  BR3_BM_Dark_Room.rotx=skirtAX(me,Activeball)
  UpdateBR3
  me.timerinterval = 150
  me.timerenabled=1
End Sub

sub bumper3_timer
  BR3_BM_Dark_Room.rotx = 0
  BR3_BM_Dark_Room.roty = 0
  UpdateBR3
  me.timerenabled=0
end sub

'******************************************************
'     SKIRT ANIMATION FUNCTIONS
'******************************************************
' NOTE: set bumper object timer to around 150-175 in order to be able
' to actually see the animation, adjust to your liking

Const PI = 3.1415926
Const SkirtTilt=3   'angle of skirt tilting in degrees

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

 'Wire Triggers
 Sub sw23_Hit  : Controller.Switch(23) = 1 : wire1_BM_Dark_Room.transz=-6: UpdateWires: End Sub
 Sub sw23_UnHit: Controller.Switch(23) = 0 : wire1_BM_Dark_Room.transz=0: UpdateWires: End Sub
 Sub sw24_Hit  : Controller.Switch(24) = 1 : wire2_BM_Dark_Room.transz=-6: UpdateWires: End Sub
 Sub sw24_UnHit: Controller.Switch(24) = 0 : wire2_BM_Dark_Room.transz=0: UpdateWires: End Sub
 Sub sw25_Hit  : Controller.Switch(25) = 1 : wire3_BM_Dark_Room.transz=-6: UpdateWires: End Sub
 Sub sw25_UnHit: Controller.Switch(25) = 0 : wire3_BM_Dark_Room.transz=0: UpdateWires: End Sub
 Sub sw26_Hit  : Controller.Switch(26) = 1 : wire4_BM_Dark_Room.transz=-6: UpdateWires: End Sub
 Sub sw26_UnHit: Controller.Switch(26) = 0 : wire4_BM_Dark_Room.transz=0: UpdateWires: End Sub
 Sub sw16_Hit  : Controller.Switch(16) = 1 : wire5_BM_Dark_Room.transz=-6: UpdateWires: End Sub
 Sub sw16_UnHit: Controller.Switch(16) = 0 : wire5_BM_Dark_Room.transz=0: UpdateWires: End Sub
 Sub sw17_Hit  : Controller.Switch(17) = 1 : wire6_BM_Dark_Room.transz=-6: UpdateWires: End Sub
 Sub sw17_UnHit: Controller.Switch(17) = 0 : wire6_BM_Dark_Room.transz=0: UpdateWires: End Sub
 Sub sw18_Hit  : Controller.Switch(18) = 1 : wire7_BM_Dark_Room.transz=-6: UpdateWires: End Sub
 Sub sw18_UnHit: Controller.Switch(18) = 0 : wire7_BM_Dark_Room.transz=0: UpdateWires: End Sub

'rollunder gates
 Sub sw22_Hit  : Controller.Switch(22) = 1 : End Sub
 Sub sw22_UnHit: Controller.Switch(22) = 0 : End Sub
 Sub sw37_Hit  : Controller.Switch(37) = 1 : End Sub
 Sub sw37_UnHit: Controller.Switch(37) = 0 : End Sub
 Sub sw38_Hit  : Controller.Switch(38) = 1 : End Sub
 Sub sw38_UnHit: Controller.Switch(38) = 0 : End Sub

'Spinner
Sub sw10_Spin : vpmTimer.PulseSw (10) :PlaySoundAtLevelStatic ("fx_spinner"), gatesoundlevel, SpinL_BM_Dark_Room: End Sub
Sub sw11_Spin : vpmTimer.PulseSw (11) :PlaySoundAtLevelStatic ("fx_spinner"), gatesoundlevel, SpinRT_BM_Dark_Room: End Sub
Sub sw12_Spin : vpmTimer.PulseSw (12) :PlaySoundAtLevelStatic ("fx_spinner"), gatesoundlevel, SpinRB_BM_Dark_Room: End Sub

'***********Rotate Spinner
Dim Angle

Sub SpinnerTimer_Timer
  ' FIXME VB
  Exit Sub
  Spinmesh.Rotx = sw10.CurrentAngle
  Angle = (sin (sw10.CurrentAngle-180))
  SpinnerRod.TransX = sin( (sw10.CurrentAngle+180) * (2*PI/360)) * 12
  SpinnerRod.TransZ = sin( (sw10.CurrentAngle- 90) * (2*PI/360)) * 3.5
  Dim SpinnerRadius: SpinnerRadius=7
  Spinmesh1.Rotx = sw11.CurrentAngle
  Angle = (sin (sw11.CurrentAngle-180))
  SpinnerRod1.TransX = sin( (sw11.CurrentAngle+180) * (2*PI/360)) * 12
  SpinnerRod1.TransZ = sin( (sw11.CurrentAngle- 90) * (2*PI/360)) * 3.5
  Spinmesh2.Rotx = sw12.CurrentAngle
  Angle = (sin (sw12.CurrentAngle-180))
  SpinnerRod2.TransX = sin( (sw12.CurrentAngle+180) * (2*PI/360)) * 12
  SpinnerRod2.TransZ = sin( (sw12.CurrentAngle- 90) * (2*PI/360)) * 3.5
End Sub


'**********************************************************************************************************
'Digital Display
'**********************************************************************************************************
Dim Digits(32)

' 1st Player
Digits(0) = Array(LED10,LED11,LED12,LED13,LED14,LED15,LED16)
Digits(1) = Array(LED20,LED21,LED22,LED23,LED24,LED25,LED26)
Digits(2) = Array(LED30,LED31,LED32,LED33,LED34,LED35,LED36)
Digits(3) = Array(LED40,LED41,LED42,LED43,LED44,LED45,LED46)
Digits(4) = Array(LED50,LED51,LED52,LED53,LED54,LED55,LED56)
Digits(5) = Array(LED60,LED61,LED62,LED63,LED64,LED65,LED66)
Digits(6) = Array(LED70,LED71,LED72,LED73,LED74,LED75,LED76)

' 2nd Player
Digits(7) = Array(LED80,LED81,LED82,LED83,LED84,LED85,LED86)
Digits(8) = Array(LED90,LED91,LED92,LED93,LED94,LED95,LED96)
Digits(9) = Array(LED100,LED101,LED102,LED103,LED104,LED105,LED106)
Digits(10) = Array(LED110,LED111,LED112,LED113,LED114,LED115,LED116)
Digits(11) = Array(LED120,LED121,LED122,LED123,LED124,LED125,LED126)
Digits(12) = Array(LED130,LED131,LED132,LED133,LED134,LED135,LED136)
Digits(13) = Array(LED140,LED141,LED142,LED143,LED144,LED145,LED146)

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

' Credits
Digits(28) = Array(LED4,LED2,LED6,LED7,LED5,LED1,LED3)
Digits(29) = Array(LED18,LED9,LED27,LED28,LED19,LED8,LED17)
' Balls
Digits(30) = Array(LED39,LED37,LED48,LED49,LED47,LED29,LED38)
Digits(31) = Array(LED67,LED58,LED69,LED77,LED68,LED57,LED59)

Sub DisplayTimer_Timer
  Dim ChgLED,ii,num,chg,stat,obj
  ChgLed = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
  If Not IsEmpty(ChgLED) Then
    If not B2SOn then
      For ii = 0 To UBound(chgLED)
        num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
        if (num < 32) then
          For Each obj In Digits(num)
            If chg And 1 Then obj.State = stat And 1
            chg = chg\2 : stat = stat\2
          Next
        else
        end if
      next
    end if
  end if
End Sub


'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep, LAstep, LBstep, LB2step, LCstep, LC2step, RCstep, RAStep, RA2Step, RBStep, RB2Step

Sub RightSlingShot_Slingshot
  vpmTimer.PulseSw 21
  RandomSoundSlingshotRight slingR_BM_Dark_Room
  RSling1.Visible = 1
  slingR_BM_Dark_Room.TransZ = -20
  RStep = 0
  RightSlingShot.TimerEnabled = 1
  UpdateSlings
End Sub

Sub RightSlingShot_Timer
  Select Case RStep
    Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:slingR_BM_Dark_Room.TransZ = -10
    Case 4:RSLing2.Visible = 0:slingR_BM_Dark_Room.TransZ = 0:RightSlingShot.TimerEnabled = 0:
  End Select
  RStep = RStep + 1
  UpdateSlings
End Sub

Sub LeftSlingShot_Slingshot
  vpmTimer.PulseSw 20
  RandomSoundSlingshotLeft slingL_BM_Dark_Room
  LSling1.Visible = 1
  slingL_BM_Dark_Room.TransZ = -20
  LStep = 0
  LeftSlingShot.TimerEnabled = 1
  UpdateSlings
End Sub

Sub LeftSlingShot_Timer
  Select Case LStep
    Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:slingL_BM_Dark_Room.TransZ = -10
    Case 4:LSLing2.Visible = 0:slingL_BM_Dark_Room.TransZ = 0:LeftSlingShot.TimerEnabled = 0:
  End Select
  LStep = LStep + 1
  UpdateSlings
End Sub

Const WallPrefix    = "T" 'Change this based on your naming convention
Const PrimitivePrefix   = "PrimT"'Change this based on your naming convention
Const PrimitiveBumperPrefix = "BumperRing" 'Change this based on your naming convention
Dim primCnt(120), primDir(120), primBmprDir(120)

'****************************************************************************
'***** Primitive Standup Target Animation
'****************************************************************************
'USAGE:   Sub sw1_Hit:  PrimStandupTgtHit  1, Sw1, PrimSw1: End Sub
'USAGE:   Sub Sw1_Timer:  PrimStandupTgtMove 1, Sw1, PrimSw1: End Sub

Const StandupTgtMovementDir = "TransY"
Const StandupTgtMovementMax = 6

Sub PrimStandupTgtHit (swnum, wallName, primName)
  PlayTargetSound
  vpmTimer.PulseSw swnum
  primCnt(swnum) = 0                  'Reset count
  wallName.TimerInterval = 10 'Set timer interval
  wallName.TimerEnabled = 1   'Enable timer
End Sub

Sub PrimStandupTgtMove (swnum, wallName, primName)
  Select Case StandupTgtMovementDir
    Case "TransX":
      Select Case primCnt(swnum)
        Case 0:   primName.TransX = -StandupTgtMovementMax
        Case 1:   primName.TransX = -StandupTgtMovementMax *2/3
        Case 2:   primName.TransX = -StandupTgtMovementMax /3
        Case 3:   primName.TransX = 0
        Case else:  wallName.TimerEnabled = 0
      End Select
    Case "TransY":
      Select Case primCnt(swnum)
        Case 0:   primName.TransY = -StandupTgtMovementMax
        Case 1:   primName.TransY = -StandupTgtMovementMax * 2/3
        Case 2:   primName.TransY = -StandupTgtMovementMax /3
        Case 3:   primName.TransY = 0
        Case else:  wallName.TimerEnabled = 0
      End Select
    Case "TransZ":
      Select Case primCnt(swnum)
        Case 0:   primName.TransZ = -StandupTgtMovementMax
        Case 1:   primName.TransZ = -StandupTgtMovementMax * 2/3
        Case 2:   primName.TransZ = -StandupTgtMovementMax /3
        Case 3:   primName.TransZ = 0
        Case else:  wallName.TimerEnabled = 0
      End Select
  End Select
  primCnt(swnum) = primCnt(swnum) + 1
  UpdateHitTarget
End Sub

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

Dim BIPL

Sub StartBallControl_UnHit()
  If Activeball.vely > 0 Then BIPL = 1 Else BIPL = 0
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

if flippershadows=1 then
  batleftshadow.visible=1
  batrightshadow.visible=1
else
  batleftshadow.visible=0
  batrightshadow.visible=0
end if

'*****************************************
'     FLIPPER SHADOWS
'*****************************************

sub FlipperTimer_Timer()
  batleftshadow.RotZ = LeftFlipper.currentangle
  batrightshadow.RotZ = RightFlipper.currentangle
  UpdateFlipper
  UpdateSpinners
  UpdateGates
End Sub

if ballshadows=1 then
  BallShadowUpdate.enabled=1
else
  BallShadowUpdate.enabled=0
end if

'*****************************************
'     BALL SHADOW
'*****************************************
Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5,BallShadow6,BallShadow7,BallShadow8,BallShadow9,BallShadow10)

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
    BallShadow(b).X = BOT(b).X - (TableWidth/2 - BOT(b).X)/20
    ballShadow(b).Y = BOT(b).Y + 10
    If BOT(b).Z > 20 Then
      BallShadow(b).visible = 1
    Else
      BallShadow(b).visible = 0
    End If
  Next
End Sub

'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 10 ' total number of balls
ReDim rolling(tnob)
InitRolling

Dim DropCount
ReDim DropCount(tnob)

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
    StopSound("BallRoll_" & b)
  Next

  ' exit the sub if no balls on the table
  If UBound(BOT) = -1 Then Exit Sub

  ' play the rolling sound for each ball

  For b = 0 to UBound(BOT)
    If BallVel(BOT(b)) > 1 AND BOT(b).z < 30 Then
      rolling(b) = True
      PlaySound ("BallRoll_" & b), -1, VolPlayfieldRoll(BOT(b)) * 1.1 * VolumeDial, AudioPan(BOT(b)), 0, PitchPlayfieldRoll(BOT(b)), 1, 0, AudioFade(BOT(b))

    Else
      If rolling(b) = True Then
        StopSound("BallRoll_" & b)
        rolling(b) = False
      End If
    End If

    '***Ball Drop Sounds***
    If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
      If DropCount(b) >= 5 Then
        DropCount(b) = 0
        If BOT(b).velz > -7 Then
          RandomSoundBallBouncePlayfieldSoft BOT(b)
        Else
          RandomSoundBallBouncePlayfieldHard BOT(b)
        End If
      End If
    End If
    If DropCount(b) < 5 Then
      DropCount(b) = DropCount(b) + 1
    End If
  Next
End Sub

'******************************************************
'   FLIPPER CORRECTION INITIALIZATION
'******************************************************

dim LF : Set LF = New FlipperPolarity
dim RF : Set RF = New FlipperPolarity

InitPolarity

Sub InitPolarity()
  dim x, a : a = Array(LF, RF)
  for each x in a
    x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
    x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1
    x.enabled = True
    x.TimeDelay = 80
  Next

  AddPt "Polarity", 0, 0, 0
  AddPt "Polarity", 1, 0.05, -2.7
  AddPt "Polarity", 2, 0.33, -2.7
  AddPt "Polarity", 3, 0.37, -2.7
  AddPt "Polarity", 4, 0.41, -2.7
  AddPt "Polarity", 5, 0.45, -2.7
  AddPt "Polarity", 6, 0.576,-2.7
  AddPt "Polarity", 7, 0.66, -1.8
  AddPt "Polarity", 8, 0.743, -0.5
  AddPt "Polarity", 9, 0.81, -0.5
  AddPt "Polarity", 10, 0.88, 0

  addpt "Velocity", 0, 0,   1
  addpt "Velocity", 1, 0.16, 1.06
  addpt "Velocity", 2, 0.41,  1.05
  addpt "Velocity", 3, 0.53,  1'0.982
  addpt "Velocity", 4, 0.702, 0.968
  addpt "Velocity", 5, 0.95,  0.968
  addpt "Velocity", 6, 1.03,  0.945

  LF.Object = LeftFlipper
  LF.EndPoint = EndPointLp
  RF.Object = RightFlipper
  RF.EndPoint = EndPointRp
End Sub

Sub TriggerLF_Hit() : LF.Addball activeball : End Sub
Sub TriggerLF_UnHit() : LF.PolarityCorrect activeball : End Sub
Sub TriggerRF_Hit() : RF.Addball activeball : End Sub
Sub TriggerRF_UnHit() : RF.PolarityCorrect activeball : End Sub

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
'   RUBBER POST AND SLEEVE DAMPENERS
'******************************************************
'These are data mined bounce curves,
'dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.

Sub dPosts_Hit(idx)
  RubbersD.dampen Activeball
End Sub

Sub dSleeves_Hit(idx)
  SleevesD.Dampen Activeball
End Sub

dim RubbersD : Set RubbersD = new Dampener  'frubber
RubbersD.name = "Rubbers"

'cor bounce curve (linear)
'for best results, try to match in-game velocity as closely as possible to the desired curve
'RubbersD.addpoint 0, 0, 0.935  'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 0, 0, 1.11  'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 1, 3.77, 0.96
RubbersD.addpoint 2, 5.76, 0.967  'dont take this as gospel. if you can data mine rubber elasticitiy, please help!
RubbersD.addpoint 3, 15.84, 0.874
RubbersD.addpoint 4, 56, 0.64 'there's clamping so interpolate up to 56 at least

dim SleevesD : Set SleevesD = new Dampener  'this is just rubber but cut down to 85%...
SleevesD.name = "Sleeves"
SleevesD.CopyCoef RubbersD, 0.85

Class Dampener
  public name, Threshold  'Minimum threshold. Useful for Flippers, which don't have a hit threshold.
  Public ModIn, ModOut
  Private Sub Class_Initialize : redim ModIn(0) : redim Modout(0): End Sub

  Public Sub AddPoint(aIdx, aX, aY)
    ShuffleArrays ModIn, ModOut, 1 : ModIn(aIDX) = aX : ModOut(aIDX) = aY : ShuffleArrays ModIn, ModOut, 0
  End Sub

  public sub Dampen(aBall)
    if threshold then if BallSpeed(aBall) < threshold then exit sub end if end if
    dim RealCOR, DesiredCOR, str, coef
    DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
    RealCOR = BallSpeed(aBall) / cor.ballvel(aBall.id)
    coef = desiredcor / realcor

' Thalamus - patched :     aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
    aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
  End Sub

  Public Sub CopyCoef(aObj, aCoef) 'alternative addpoints, copy with coef
    dim x : for x = 0 to uBound(aObj.ModIn)
      addpoint x, aObj.ModIn(x), aObj.ModOut(x)*aCoef
    Next
  End Sub

End Class

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
'   DROP TARGETS INITIALIZATION
'******************************************************

'Define a variable for each drop target
Dim DT27, DT28, DT29, DT30, DT31, DT32, DT33, DT34, DT35

'Set array with drop target objects
'
'DropTargetvar = Array(primary, secondary, prim, switch, animate)
'   primary:      primary target wall to determine drop
' secondary:      wall used to simulate the ball striking a bent or offset target after the initial Hit
' prim:       primitive target used for visuals and animation
'             IMPORTANT!!!
'             rotz must be used for orientation
'             rotx to bend the target back
'             transz to move it up and down
'             the pivot point should be in the center of the target on the x, y and at or below the playfield (0) on z
' switch:       ROM switch number
' animate:      Arrary slot for handling the animation instrucitons, set to 0
'
' Values for annimate: 1 - bend target (hit to primary), 2 - drop target (hit to secondary), 3 - brick target (high velocity hit to secondary), -1 - raise target

' 1 Bank
DT27 = Array(sw27, sw27y, DT1_1_BM_Dark_Room, 27, 0)
DT28 = Array(sw28, sw28y, DT1_2_BM_Dark_Room, 28, 0)
DT29 = Array(sw29, sw29y, DT1_3_BM_Dark_Room, 29, 0)

' 2 Bank
DT30 = Array(sw30, sw30y, DT2_1_BM_Dark_Room, 30, 0)
DT31 = Array(sw31, sw31y, DT2_2_BM_Dark_Room, 31, 0)
DT32 = Array(sw32, sw32y, DT2_3_BM_Dark_Room, 32, 0)

' 3 Bank
DT33 = Array(sw33, sw33y, DT3_1_BM_Dark_Room, 33, 0)
DT34 = Array(sw34, sw34y, DT3_2_BM_Dark_Room, 34, 0)
DT35 = Array(sw35, sw35y, DT3_3_BM_Dark_Room, 35, 0)


'Add all the Drop Target Arrays to Drop Target Animation Array
' DTAnimationArray = Array(DT1, DT2, ....)
Dim DTArray
DTArray = Array(DT27, DT28, DT29, DT30, DT31, DT32, DT33, DT34, DT35)

'Configure the behavior of Drop Targets.
Const DTDropSpeed = 110       'in milliseconds
Const DTDropUpSpeed = 40      'in milliseconds
Const DTDropUnits = 50        'VP units primitive drops
Const DTDropUpUnits = 5       'VP units primitive raises above the up position on drops up
Const DTMaxBend = 8         'max degrees primitive rotates when hit
Const DTDropDelay = 20      'time in milliseconds before target drops (due to friction/impact of the ball)
Const DTRaiseDelay = 40     'time in milliseconds before target drops back to normal up position after the solendoid fires to raise the target
Const DTBrickVel = 30       'velocity at which the target will brick, set to '0' to disable brick

Const DTEnableBrick = 0     'Set to 0 to disable bricking, 1 to enable bricking
Const DTHitSound = "Target_Hit_1" 'Drop Target Hit sound
Const DTDropSound = "DTDrop"    'Drop Target Drop sound
Const DTResetSound = "DTReset"  'Drop Target reset sound

Const DTMass = 0.2        'Mass of the Drop Target (between 0 and 1), higher values provide more resistance

'******************************************************
'       DROP TARGETS FUNCTIONS
'******************************************************

Sub DTHit(switch)
  Dim i
  i = DTArrayID(switch)

  PlayTargetSound
  DTArray(i)(4) =  DTCheckBrick(Activeball,DTArray(i)(2))
  DTArray(i)(4) =  DTCheckBrick(Activeball,DTArray(i)(2))
  If DTArray(i)(4) = 1 or DTArray(i)(4) = 3 Then
    DTBallPhysics Activeball, DTArray(i)(2).rotz, DTMass
  End If

  DoDTAnim
End Sub

Sub DTRaise(switch)
  Dim i
  i = DTArrayID(switch)

  DTArray(i)(4) = -1
  DoDTAnim
End Sub

Sub DTDrop(switch)
  Dim i
  i = DTArrayID(switch)

  DTArray(i)(4) = 1
  DoDTAnim
End Sub

Function DTArrayID(switch)
  Dim i
  For i = 0 to uBound(DTArray)
    If DTArray(i)(3) = switch Then DTArrayID = i:Exit Function
  Next
End Function


sub DTBallPhysics(aBall, angle, mass)
  dim rangle,bangle,calc1, calc2, calc3
  rangle = (angle - 90) * 3.1416 / 180
  bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))

  calc1 = cor.BallVel(aball.id) * cos(bangle - rangle) * (aball.mass - mass) / (aball.mass + mass)
  calc2 = cor.BallVel(aball.id) * sin(bangle - rangle) * cos(rangle + 4*Atn(1)/2)
  calc3 = cor.BallVel(aball.id) * sin(bangle - rangle) * sin(rangle + 4*Atn(1)/2)

  aBall.velx = calc1 * cos(rangle) + calc2
  aBall.vely = calc1 * sin(rangle) + calc3
End Sub

'Add Timer name DTAnim to editor to handle drop target animations
DTAnim.interval = 10
DTAnim.enabled = True

Sub DTAnim_Timer()
  DoDTAnim
End Sub

'Check if target is hit on it's face or sides and whether a 'brick' occurred
Function DTCheckBrick(aBall, dtprim)
  dim bangle, bangleafter, rangle, rangle2, Xintersect, Yintersect, cdist, perpvel, perpvelafter
  rangle = (dtprim.rotz - 90) * 3.1416 / 180
  rangle2 = dtprim.rotz * 3.1416 / 180
  bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))
  bangleafter = Atn2(aBall.vely,aball.velx)

  Xintersect = (aBall.y - dtprim.y - tan(bangle) * aball.x + tan(rangle2) * dtprim.x) / (tan(rangle2) - tan(bangle))
  Yintersect = tan(rangle2) * Xintersect + (dtprim.y - tan(rangle2) * dtprim.x)

  cdist = Distance(dtprim.x, dtprim.y, Xintersect, Yintersect)

  perpvel = cor.BallVel(aball.id) * cos(bangle-rangle)
  perpvelafter = BallSpeed(aBall) * cos(bangleafter - rangle)

  If perpvel <= 0 or perpvelafter >= 0 Then
    DTCheckBrick = 0
  ElseIf DTEnableBrick = 1 and  perpvel > DTBrickVel and DTBrickVel <> 0 and cdist < 8 Then
    DTCheckBrick = 3
  Else
    DTCheckBrick = 1
  End If
End Function


Sub DoDTAnim()
  Dim i
  For i=0 to Ubound(DTArray)
    DTArray(i)(4) = DTAnimate(DTArray(i)(0),DTArray(i)(1),DTArray(i)(2),DTArray(i)(3),DTArray(i)(4))
  Next
  'table specific code
' DTShadow psw25, drop4shadow, "drop4"
' DTShadow psw26, drop3shadow, "drop3"
' DTShadow psw27, drop2shadow, "drop2"
' DTShadow psw28, drop1shadow, "drop1"
  UpdateTargets
End Sub

Function DTAnimate(primary, secondary, prim, switch,  animate)
  dim transz
  Dim animtime, rangle

  rangle = prim.rotz * 3.1416 / 180

  DTAnimate = animate

  if animate = 0  Then
    primary.uservalue = 0
    DTAnimate = 0
    Exit Function
  Elseif primary.uservalue = 0 then
    primary.uservalue = gametime
  end if

  animtime = gametime - primary.uservalue

  If animate = 1 and animtime < DTDropDelay Then
    primary.collidable = 0
    secondary.collidable = 1
    prim.rotx = DTMaxBend * cos(rangle)
    prim.roty = DTMaxBend * sin(rangle)
    DTAnimate = 1
    Exit Function
  elseif animate = 1 and animtime > DTDropDelay Then
    primary.collidable = 0
    secondary.collidable = 1
    prim.rotx = DTMaxBend * cos(rangle)
    prim.roty = DTMaxBend * sin(rangle)
    animate = 2
    PlaySoundAt SoundFX(DTDropSound,DOFDropTargets),prim
  End If

  if animate = 2 Then
    transz = (animtime - DTDropDelay)/DTDropSpeed *  DTDropUnits * -1
    if prim.transz > -DTDropUnits  Then
      prim.transz = transz
    end if

    prim.rotx = DTMaxBend * cos(rangle)/2
    prim.roty = DTMaxBend * sin(rangle)/2

    if prim.transz <= -DTDropUnits Then
      prim.transz = -DTDropUnits
      secondary.collidable = 0
      controller.Switch(Switch) = 1
      primary.uservalue = 0
      DTAnimate = 0
      Exit Function
    Else
      DTAnimate = 2
      Exit Function
    end If
  End If

  If animate = 3 and animtime < DTDropDelay Then
    primary.collidable = 0
    secondary.collidable = 1
    prim.rotx = DTMaxBend * cos(rangle)
    prim.roty = DTMaxBend * sin(rangle)
  elseif animate = 3 and animtime > DTDropDelay Then
    primary.collidable = 1
    secondary.collidable = 0
    prim.rotx = 0
    prim.roty = 0
    primary.uservalue = 0
    DTAnimate = 0
    Exit Function
  End If

  if animate = -1 Then
    transz = (1 - (animtime)/DTDropUpSpeed) *  DTDropUnits * -1

    If prim.transz = -DTDropUnits Then
      Dim BOT, b
      BOT = GetBalls

      For b = 0 to UBound(BOT)
        If InRect(BOT(b).x,BOT(b).y,prim.x-25,prim.y-10,prim.x+25, prim.y-10,prim.x+25,prim.y+25,prim.x -25,prim.y+25) Then
          BOT(b).velz = 20
        End If
      Next
    End If

    if prim.transz < 0 Then
      prim.transz = transz
    elseif transz > 0 then
      prim.transz = transz
    end if

    if prim.transz > DTDropUpUnits then
      DTAnimate = -2
      prim.rotx = 0
      prim.roty = 0
      primary.uservalue = gametime
    end if
    primary.collidable = 0
    secondary.collidable = 1
    controller.Switch(Switch) = 0

  End If

  if animate = -2 and animtime > DTRaiseDelay Then
    prim.transz = (animtime - DTRaiseDelay)/DTDropSpeed *  DTDropUnits * -1 + DTDropUpUnits
    if prim.transz < 0 then
      prim.transz = 0
      primary.uservalue = 0
      DTAnimate = 0

      primary.collidable = 1
      secondary.collidable = 0
    end If
  End If
End Function

'' Table specific function
'Sub  DTShadow(prim, shadowprim, imagename)
' If prim.transz < -DTDropUnits/2 Then
'   shadowprim.image ="blank"
' Else
'   shadowprim.image = imagename
' End If
'End Sub

'******************************************************
'   FLIPPER POLARITY, DAMPENER, AND DROP TARGET
'       SUPPORTING FUNCTIONS
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

' Used for drop targets
Function Atn2(dy, dx)
  dim pi
  pi = 4*Atn(1)

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

' Used for drop targets
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

' Used for drop targets
Function Distance(ax,ay,bx,by)
  Distance = SQR((ax - bx)^2 + (ay - by)^2)
End Function

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
Dim RubberStrongSoundFactor, RubberWeakSoundFactor, RubberFlipperSoundFactor,BallWithBallCollisionSoundFactor
Dim BallBouncePlayfieldSoftFactor, BallBouncePlayfieldHardFactor, PlasticRampDropToPlayfieldSoundLevel, WireRampDropToPlayfieldSoundLevel, DelayedBallDropOnPlayfieldSoundLevel
Dim WallImpactSoundFactor, MetalImpactSoundFactor, SubwaySoundLevel, SubwayEntrySoundLevel, ScoopEntrySoundLevel
Dim SaucerLockSoundLevel, SaucerKickSoundLevel

BallWithBallCollisionSoundFactor = 3.2                  'volume multiplier; must not be zero
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
  PlaySound ("Start_Button"), 0, StartButtonSoundLevel, 0, 0.25
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
    Case 1 : PlaySound ("Nudge_1"), 0, NudgeCenterSoundLevel * VolumeDial, 0, 0.25
    Case 2 : PlaySound ("Nudge_2"), 0, NudgeCenterSoundLevel * VolumeDial, 0, 0.25
    Case 3 : PlaySound ("Nudge_3"), 0, NudgeCenterSoundLevel * VolumeDial, 0, 0.25
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

'/////////////////////////////  TROUGH BALL RELEASE SOLENOID SOUNDS  ////////////////////////////
Sub RandomSoundBallRelease(drainswitch)
  Select Case Int(Rnd*7)+1
    Case 1 : PlaySoundAtLevelStatic SoundFX("BallRelease1",DOFContactors), BallReleaseSoundLevel, drainswitch
    Case 2 : PlaySoundAtLevelStatic SoundFX("BallRelease2",DOFContactors), BallReleaseSoundLevel, drainswitch
    Case 3 : PlaySoundAtLevelStatic SoundFX("BallRelease3",DOFContactors), BallReleaseSoundLevel, drainswitch
    Case 4 : PlaySoundAtLevelStatic SoundFX("BallRelease4",DOFContactors), BallReleaseSoundLevel, drainswitch
    Case 5 : PlaySoundAtLevelStatic SoundFX("BallRelease5",DOFContactors), BallReleaseSoundLevel, drainswitch
    Case 6 : PlaySoundAtLevelStatic SoundFX("BallRelease6",DOFContactors), BallReleaseSoundLevel, drainswitch
    Case 7 : PlaySoundAtLevelStatic SoundFX("BallRelease7",DOFContactors), BallReleaseSoundLevel, drainswitch
  End Select
End Sub



'/////////////////////////////  SLINGSHOT SOLENOID SOUNDS  ////////////////////////////
Sub RandomSoundSlingshotLeft(sling)
  Select Case Int(Rnd*10)+1
    Case 1 : PlaySoundAtLevelStatic SoundFX("Sling_L1",DOFContactors), SlingshotSoundLevel, Sling
    Case 2 : PlaySoundAtLevelStatic SoundFX("Sling_L2",DOFContactors), SlingshotSoundLevel, Sling
    Case 3 : PlaySoundAtLevelStatic SoundFX("Sling_L3",DOFContactors), SlingshotSoundLevel, Sling
    Case 4 : PlaySoundAtLevelStatic SoundFX("Sling_L4",DOFContactors), SlingshotSoundLevel, Sling
    Case 5 : PlaySoundAtLevelStatic SoundFX("Sling_L5",DOFContactors), SlingshotSoundLevel, Sling
    Case 6 : PlaySoundAtLevelStatic SoundFX("Sling_L6",DOFContactors), SlingshotSoundLevel, Sling
    Case 7 : PlaySoundAtLevelStatic SoundFX("Sling_L7",DOFContactors), SlingshotSoundLevel, Sling
    Case 8 : PlaySoundAtLevelStatic SoundFX("Sling_L8",DOFContactors), SlingshotSoundLevel, Sling
    Case 9 : PlaySoundAtLevelStatic SoundFX("Sling_L9",DOFContactors), SlingshotSoundLevel, Sling
    Case 10 : PlaySoundAtLevelStatic SoundFX("Sling_L10",DOFContactors), SlingshotSoundLevel, Sling
  End Select
End Sub

Sub RandomSoundSlingshotRight(sling)
  Select Case Int(Rnd*8)+1
    Case 1 : PlaySoundAtLevelStatic SoundFX("Sling_R1",DOFContactors), SlingshotSoundLevel, Sling
    Case 2 : PlaySoundAtLevelStatic SoundFX("Sling_R2",DOFContactors), SlingshotSoundLevel, Sling
    Case 3 : PlaySoundAtLevelStatic SoundFX("Sling_R3",DOFContactors), SlingshotSoundLevel, Sling
    Case 4 : PlaySoundAtLevelStatic SoundFX("Sling_R4",DOFContactors), SlingshotSoundLevel, Sling
    Case 5 : PlaySoundAtLevelStatic SoundFX("Sling_R5",DOFContactors), SlingshotSoundLevel, Sling
    Case 6 : PlaySoundAtLevelStatic SoundFX("Sling_R6",DOFContactors), SlingshotSoundLevel, Sling
    Case 7 : PlaySoundAtLevelStatic SoundFX("Sling_R7",DOFContactors), SlingshotSoundLevel, Sling
    Case 8 : PlaySoundAtLevelStatic SoundFX("Sling_R8",DOFContactors), SlingshotSoundLevel, Sling
  End Select
End Sub

'/////////////////////////////  BUMPER SOLENOID SOUNDS  ////////////////////////////
Sub RandomSoundBumperTop(Bump)
  Select Case Int(Rnd*5)+1
    Case 1 : PlaySoundAtLevelStatic SoundFX("Bumpers_Top_1",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
    Case 2 : PlaySoundAtLevelStatic SoundFX("Bumpers_Top_2",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
    Case 3 : PlaySoundAtLevelStatic SoundFX("Bumpers_Top_3",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
    Case 4 : PlaySoundAtLevelStatic SoundFX("Bumpers_Top_4",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
    Case 5 : PlaySoundAtLevelStatic SoundFX("Bumpers_Top_5",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
  End Select
End Sub

Sub RandomSoundBumperMiddle(bump)
  Select Case Int(Rnd*5)+1
    Case 1 : PlaySoundAtLevelStatic SoundFX("Bumpers_Middle_1",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
    Case 2 : PlaySoundAtLevelStatic SoundFX("Bumpers_Middle_2",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
    Case 3 : PlaySoundAtLevelStatic SoundFX("Bumpers_Middle_3",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
    Case 4 : PlaySoundAtLevelStatic SoundFX("Bumpers_Middle_4",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
    Case 5 : PlaySoundAtLevelStatic SoundFX("Bumpers_Middle_5",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
  End Select
End Sub

Sub RandomSoundBumperBottom(Bump)
  Select Case Int(Rnd*5)+1
    Case 1 : PlaySoundAtLevelStatic SoundFX("Bumpers_Bottom_1",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
    Case 2 : PlaySoundAtLevelStatic SoundFX("Bumpers_Bottom_2",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
    Case 3 : PlaySoundAtLevelStatic SoundFX("Bumpers_Bottom_3",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
    Case 4 : PlaySoundAtLevelStatic SoundFX("Bumpers_Bottom_4",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
    Case 5 : PlaySoundAtLevelStatic SoundFX("Bumpers_Bottom_5",DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
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

Sub Rollovers_Hit(idx)
  RandomSoundRollover
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

Sub ShooterDiverter_collide(idx)
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

'/////////////////////////////  BOTTOM ARCH BALL GUIDE - HARD HITS  ////////////////////////////
Sub RandomSoundBottomArchBallGuideHardHit()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySoundAtLevelActiveBall ("Apron_Hard_Hit_1"), BottomArchBallGuideSoundFactor * 0.25
    Case 2 : PlaySoundAtLevelActiveBall ("Apron_Hard_Hit_2"), BottomArchBallGuideSoundFactor * 0.25
    Case 3 : PlaySoundAtLevelActiveBall ("Apron_Hard_Hit_3"), BottomArchBallGuideSoundFactor * 0.25
  End Select
End Sub


Sub Apron_Hit (idx)
  If Abs(cor.ballvelx(activeball.id) < 4) and cor.ballvely(activeball.id) > 7 then
    RandomSoundBottomArchBallGuideHardHit()
  Else
    RandomSoundBottomArchBallGuide
  End If
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
    RandomSoundBallBouncePlayfieldSoft Activeball
  Else
    RandomSoundTargetHitWeak()
  End If
End Sub

Sub Targets_Hit (idx)
  PlayTargetSound
End Sub

'/////////////////////////////  BALL BOUNCE SOUNDS  ////////////////////////////
Sub RandomSoundBallBouncePlayfieldSoft(aBall)
  Select Case Int(Rnd*9)+1
    Case 1 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_1"), volz(aBall) * BallBouncePlayfieldSoftFactor, aBall
    Case 2 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_2"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.5, aBall
    Case 3 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_3"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.8, aBall
    Case 4 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_4"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.5, aBall
    Case 5 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_5"), volz(aBall) * BallBouncePlayfieldSoftFactor, aBall
    Case 6 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_1"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
    Case 7 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_2"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
    Case 8 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_5"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
    Case 9 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_7"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.3, aBall
  End Select
End Sub

Sub RandomSoundBallBouncePlayfieldHard(aBall)
  Select Case Int(Rnd*7)+1
    Case 1 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_1"), volz(aBall) * BallBouncePlayfieldHardFactor, aBall
    Case 2 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_2"), volz(aBall) * BallBouncePlayfieldHardFactor, aBall
    Case 3 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_3"), volz(aBall) * BallBouncePlayfieldHardFactor, aBall
    Case 4 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_4"), volz(aBall) * BallBouncePlayfieldHardFactor, aBall
    Case 5 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_5"), volz(aBall) * BallBouncePlayfieldHardFactor, aBall
    Case 6 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_6"), volz(aBall) * BallBouncePlayfieldHardFactor, aBall
    Case 7 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_7"), volz(aBall) * BallBouncePlayfieldHardFactor, aBall
  End Select
End Sub

'/////////////////////////////  DELAYED DROP - TO PLAYFIELD - SOUND  ////////////////////////////
Sub RandomSoundDelayedBallDropOnPlayfield(aBall)
  Select Case Int(Rnd*5)+1
    Case 1 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_1_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 2 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_2_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 3 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_3_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 4 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_4_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 5 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_5_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
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

Sub Gates_hit(idx)
  SoundHeavyGate
End Sub

Sub GatesWire_hit(idx)
  SoundPlayfieldGate
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
  If activeball.velx < -8 Then
    RandomSoundRightArch
  End If
End Sub

Sub Arch2_hit()
  If Activeball.velx < 1 Then SoundPlayfieldGate
  StopSound "Arch_R1"
  StopSound "Arch_R2"
  StopSound "Arch_R3"
  StopSound "Arch_R4"
End Sub

Sub Arch2_unhit()
  If activeball.velx > 10 Then
    RandomSoundLeftArch
  End If
End Sub

'/////////////////////////////  SAUCERS (KICKER HOLES)  ////////////////////////////

Sub SoundSaucerLock()
  Select Case Int(Rnd*2)+1
    Case 1: PlaySoundAtLevelStatic ("Saucer_Enter_1"), SaucerLockSoundLevel, Activeball
    Case 2: PlaySoundAtLevelStatic ("Saucer_Enter_2"), SaucerLockSoundLevel, Activeball
  End Select
End Sub

Sub SoundSaucerKick(scenario, saucer)
  Select Case scenario
    Case 0: PlaySoundAtLevelStatic SoundFX("Saucer_Empty", DOFContactors), SaucerKickSoundLevel, saucer
    Case 1: PlaySoundAtLevelStatic SoundFX("Saucer_Kick", DOFContactors), SaucerKickSoundLevel, saucer
  End Select
End Sub

'/////////////////////////////  BALL COLLISION SOUND  ////////////////////////////
Sub OnBallBallCollision(ball1, ball2, velocity)
  Dim snd
  Select Case Int(Rnd*7)+1
    Case 1 : snd = "Ball_Collide_1"
    Case 2 : snd = "Ball_Collide_2"
    Case 3 : snd = "Ball_Collide_3"
    Case 4 : snd = "Ball_Collide_4"
    Case 5 : snd = "Ball_Collide_5"
    Case 6 : snd = "Ball_Collide_6"
    Case 7 : snd = "Ball_Collide_7"
  End Select

  PlaySound (snd), 0, Csng(velocity) ^2 / 200 * BallWithBallCollisionSoundFactor * VolumeDial, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub





'******************************************************
'****  LAMPZ by nFozzy
'******************************************************
'
' Lampz is a utility designed to manage and fade the lights and light-related objects on a table that is being driven by a ROM.
' To set up Lampz, one must populate the Lampz.MassAssign array with VPX Light objects, where the index of the MassAssign array
' corrisponds to the ROM index of the associated light. More that one Light object can be associated with a single MassAssign index (not shown in this example)
' Optionally, callbacks can be assigned for each index using the Lampz.Callback array. This is very useful for allowing 3D Insert primitives
' to be controlled by the ROM. Note, the aLvl parameter (i.e. the fading level that ranges between 0 and 1) is appended to the callback call.

Sub SetModLamp(id, val)
  Lampz.state(id) = val
End Sub

Dim NullFader : set NullFader = new NullFadingObject
Dim Lampz : Set Lampz = New LampFader
InitLampsNF               ' Setup lamp assignments
LampTimer.Interval = 16   ' Using fixed value so the fading speed is same for every fps
LampTimer.Enabled = 1

Sub LampTimer_Timer()
  dim x, chglamp
  chglamp = Controller.ChangedLamps
  If Not IsEmpty(chglamp) Then
    For x = 0 To UBound(chglamp)      'nmbr = chglamp(x, 0), state = chglamp(x, 1)
      Lampz.state(chglamp(x, 0)) = chglamp(x, 1)
    next
  End If
  Lampz.Update1 'update (fading logic only)
End Sub

dim FrameTime, InitFrameTime : InitFrameTime = 0
LampTimer2.Interval = -1
LampTimer2.Enabled = True
Sub LampTimer2_Timer()
  FrameTime = gametime - InitFrameTime : InitFrameTime = gametime 'Count frametime. Unused atm?
  Lampz.Update 'updates on frametime (Object updates only)
End Sub

Function FlashLevelToIndex(Input, MaxSize)
  FlashLevelToIndex = cInt(MaxSize * Input)
End Function

'***Material Swap***
'Fade material for green, red, yellow colored Bulb prims
Sub FadeMaterialColoredBulb(pri, group, ByVal aLvl) 'cp's script
  if Lampz.UseFunction then aLvl = Lampz.FilterOut(aLvl)  'Callbacks don't get this filter automatically
  Select case FlashLevelToIndex(aLvl, 3)
    Case 0:pri.Material = group(0) 'Off
    Case 1:pri.Material = group(1) 'Fading...
    Case 2:pri.Material = group(2) 'Fading...
    Case 3:pri.Material = group(3) 'Full
  End Select
  'if tb.text <> pri.image then tb.text = pri.image : 'debug.print pri.image end If 'debug
  pri.blenddisablelighting = aLvl * 1 'Intensity Adjustment
End Sub

'Fade material for red, yellow colored bulb Filiment prims
Sub FadeMaterialColoredFiliment(pri, group, ByVal aLvl) 'cp's script
  if Lampz.UseFunction then aLvl = Lampz.FilterOut(aLvl)  'Callbacks don't get this filter automatically
  Select case FlashLevelToIndex(aLvl, 3)
    Case 0:pri.Material = group(0) 'Off
    Case 1:pri.Material = group(1) 'Fading...
    Case 2:pri.Material = group(2) 'Fading...
    Case 3:pri.Material = group(3) 'Full
  End Select
  'if tb.text <> pri.image then tb.text = pri.image : 'debug.print pri.image end If 'debug
  pri.blenddisablelighting = aLvl * 50  'Intensity Adjustment
End Sub

Sub DisableLighting(pri, DLintensity, ByVal aLvl) 'cp's script  DLintensity = disabled lighting intesity
  if Lampz.UseFunction then aLvl = Lampz.FilterOut(aLvl)  'Callbacks don't get this filter automatically
  pri.blenddisablelighting = aLvl * DLintensity
End Sub

Sub UpdateLightMap(lightmap, intensity, ByVal aLvl)
  if Lampz.UseFunction then aLvl = Lampz.FilterOut(aLvl)  'Callbacks don't get this filter automatically
  lightmap.Opacity = aLvl * intensity
End Sub

Sub UpdateGiMap(lightmap, intensity, ByVal aLvl)
  if Lampz.UseFunction then aLvl = Lampz.FilterOut(aLvl)  'Callbacks don't get this filter automatically
  lightmap.Opacity = 0.25 * aLvl * intensity
End Sub

Sub InitLampsNF()
  Lampz.Filter = "LampFilter" 'Puts all lamp intensityscale output (no callbacks) through this function before updating

  'Adjust fading speeds (max level / full MS fading time)
  Dim x
  for x = 0 to 150 : Lampz.FadeSpeedUp(x) = 1/5 : Lampz.FadeSpeedDown(x) = 1/20 : next

  Lampz.Modulate(100) = 1/100 : Lampz.FadeSpeedUp(100) = 100/5 : Lampz.FadeSpeedDown(100) = 100/5

  ' Lampz.MassAssign(100) = L100 ' VLM.Lampz;Lit Room
  Lampz.Callback(100) = "UpdateLightMap Playfield_LM_Lit_Room, 100.00, " ' VLM.Lampz;Lit Room
  Lampz.Callback(100) = "UpdateLightMap Over_LM_Lit_Room, 100.00, " ' VLM.Lampz;Lit Room
  Lampz.Callback(100) = "UpdateLightMap BR1_LM_Lit_Room, 100.00, " ' VLM.Lampz;Lit Room
  Lampz.Callback(100) = "UpdateLightMap BR2_LM_Lit_Room, 100.00, " ' VLM.Lampz;Lit Room
  Lampz.Callback(100) = "UpdateLightMap BR3_LM_Lit_Room, 100.00, " ' VLM.Lampz;Lit Room
  Lampz.Callback(100) = "UpdateLightMap DT1_1_LM_Lit_Room, 100.00, " ' VLM.Lampz;Lit Room
  Lampz.Callback(100) = "UpdateLightMap DT1_2_LM_Lit_Room, 100.00, " ' VLM.Lampz;Lit Room
  Lampz.Callback(100) = "UpdateLightMap DT1_3_LM_Lit_Room, 100.00, " ' VLM.Lampz;Lit Room
  Lampz.Callback(100) = "UpdateLightMap DT2_1_LM_Lit_Room, 100.00, " ' VLM.Lampz;Lit Room
  Lampz.Callback(100) = "UpdateLightMap DT2_2_LM_Lit_Room, 100.00, " ' VLM.Lampz;Lit Room
  Lampz.Callback(100) = "UpdateLightMap DT2_3_LM_Lit_Room, 100.00, " ' VLM.Lampz;Lit Room
  Lampz.Callback(100) = "UpdateLightMap DT3_1_LM_Lit_Room, 100.00, " ' VLM.Lampz;Lit Room
  Lampz.Callback(100) = "UpdateLightMap DT3_2_LM_Lit_Room, 100.00, " ' VLM.Lampz;Lit Room
  Lampz.Callback(100) = "UpdateLightMap DT3_3_LM_Lit_Room, 100.00, " ' VLM.Lampz;Lit Room
  Lampz.Callback(100) = "UpdateLightMap FlipperL_LM_Lit_Room, 100.00, " ' VLM.Lampz;Lit Room
  Lampz.Callback(100) = "UpdateLightMap FlipperR_LM_Lit_Room, 100.00, " ' VLM.Lampz;Lit Room
  Lampz.Callback(100) = "UpdateLightMap Gate_1_LM_Lit_Room, 100.00, " ' VLM.Lampz;Lit Room
  Lampz.Callback(100) = "UpdateLightMap Gate_2_LM_Lit_Room, 100.00, " ' VLM.Lampz;Lit Room
  Lampz.Callback(100) = "UpdateLightMap Gate_3_LM_Lit_Room, 100.00, " ' VLM.Lampz;Lit Room
  Lampz.Callback(100) = "UpdateLightMap Gate_4_LM_Lit_Room, 100.00, " ' VLM.Lampz;Lit Room
  Lampz.Callback(100) = "UpdateLightMap Gate_5_LM_Lit_Room, 100.00, " ' VLM.Lampz;Lit Room
  Lampz.Callback(100) = "UpdateLightMap Gate_6_LM_Lit_Room, 100.00, " ' VLM.Lampz;Lit Room
  Lampz.Callback(100) = "UpdateLightMap HT_LM_Lit_Room, 100.00, " ' VLM.Lampz;Lit Room
  Lampz.Callback(100) = "UpdateLightMap SpinL_LM_Lit_Room, 100.00, " ' VLM.Lampz;Lit Room
  Lampz.Callback(100) = "UpdateLightMap SpinRB_LM_Lit_Room, 100.00, " ' VLM.Lampz;Lit Room
  Lampz.Callback(100) = "UpdateLightMap SpinRT_LM_Lit_Room, 100.00, " ' VLM.Lampz;Lit Room
  Lampz.Callback(100) = "UpdateLightMap slingL_LM_Lit_Room, 100.00, " ' VLM.Lampz;Lit Room
  Lampz.Callback(100) = "UpdateLightMap slingR_LM_Lit_Room, 100.00, " ' VLM.Lampz;Lit Room
  Lampz.Callback(100) = "UpdateLightMap wire1_LM_Lit_Room, 100.00, " ' VLM.Lampz;Lit Room
  Lampz.Callback(100) = "UpdateLightMap wire2_LM_Lit_Room, 100.00, " ' VLM.Lampz;Lit Room
  Lampz.Callback(100) = "UpdateLightMap wire3_LM_Lit_Room, 100.00, " ' VLM.Lampz;Lit Room
  Lampz.Callback(100) = "UpdateLightMap wire4_LM_Lit_Room, 100.00, " ' VLM.Lampz;Lit Room
  Lampz.Callback(100) = "UpdateLightMap wire6_LM_Lit_Room, 100.00, " ' VLM.Lampz;Lit Room
  Lampz.Callback(100) = "UpdateLightMap wire7_LM_Lit_Room, 100.00, " ' VLM.Lampz;Lit Room

  For Each x in Gi : Lampz.MassAssign(101) = x : Next
  ' Lampz.MassAssign(101) = L101 ' VLM.Lampz;GI.Lights
  Lampz.Callback(101) = "UpdateLightMap Playfield_LM_GI_Lights, 100.00, " ' VLM.Lampz;GI.Lights
  Lampz.Callback(101) = "UpdateLightMap Over_LM_GI_Lights, 100.00, " ' VLM.Lampz;GI.Lights
  Lampz.Callback(101) = "UpdateLightMap BR1_LM_GI_Lights, 100.00, " ' VLM.Lampz;GI.Lights
  Lampz.Callback(101) = "UpdateLightMap BR2_LM_GI_Lights, 100.00, " ' VLM.Lampz;GI.Lights
  Lampz.Callback(101) = "UpdateLightMap BR3_LM_GI_Lights, 100.00, " ' VLM.Lampz;GI.Lights
  Lampz.Callback(101) = "UpdateLightMap DT1_1_LM_GI_Lights, 100.00, " ' VLM.Lampz;GI.Lights
  Lampz.Callback(101) = "UpdateLightMap DT1_2_LM_GI_Lights, 100.00, " ' VLM.Lampz;GI.Lights
  Lampz.Callback(101) = "UpdateLightMap DT1_3_LM_GI_Lights, 100.00, " ' VLM.Lampz;GI.Lights
  Lampz.Callback(101) = "UpdateLightMap DT2_1_LM_GI_Lights, 100.00, " ' VLM.Lampz;GI.Lights
  Lampz.Callback(101) = "UpdateLightMap DT2_2_LM_GI_Lights, 100.00, " ' VLM.Lampz;GI.Lights
  Lampz.Callback(101) = "UpdateLightMap DT2_3_LM_GI_Lights, 100.00, " ' VLM.Lampz;GI.Lights
  Lampz.Callback(101) = "UpdateLightMap DT3_1_LM_GI_Lights, 100.00, " ' VLM.Lampz;GI.Lights
  Lampz.Callback(101) = "UpdateLightMap DT3_2_LM_GI_Lights, 100.00, " ' VLM.Lampz;GI.Lights
  Lampz.Callback(101) = "UpdateLightMap DT3_3_LM_GI_Lights, 100.00, " ' VLM.Lampz;GI.Lights
  Lampz.Callback(101) = "UpdateLightMap FlipperL_LM_GI_Lights, 100.00, " ' VLM.Lampz;GI.Lights
  Lampz.Callback(101) = "UpdateLightMap FlipperR_LM_GI_Lights, 100.00, " ' VLM.Lampz;GI.Lights
  Lampz.Callback(101) = "UpdateLightMap Gate_1_LM_GI_Lights, 100.00, " ' VLM.Lampz;GI.Lights
  Lampz.Callback(101) = "UpdateLightMap Gate_2_LM_GI_Lights, 100.00, " ' VLM.Lampz;GI.Lights
  Lampz.Callback(101) = "UpdateLightMap Gate_3_LM_GI_Lights, 100.00, " ' VLM.Lampz;GI.Lights
  Lampz.Callback(101) = "UpdateLightMap Gate_4_LM_GI_Lights, 100.00, " ' VLM.Lampz;GI.Lights
  Lampz.Callback(101) = "UpdateLightMap Gate_5_LM_GI_Lights, 100.00, " ' VLM.Lampz;GI.Lights
  Lampz.Callback(101) = "UpdateLightMap Gate_6_LM_GI_Lights, 100.00, " ' VLM.Lampz;GI.Lights
  Lampz.Callback(101) = "UpdateLightMap HT_LM_GI_Lights, 100.00, " ' VLM.Lampz;GI.Lights
  Lampz.Callback(101) = "UpdateLightMap SpinL_LM_GI_Lights, 100.00, " ' VLM.Lampz;GI.Lights
  Lampz.Callback(101) = "UpdateLightMap SpinRB_LM_GI_Lights, 100.00, " ' VLM.Lampz;GI.Lights
  Lampz.Callback(101) = "UpdateLightMap SpinRT_LM_GI_Lights, 100.00, " ' VLM.Lampz;GI.Lights
  Lampz.Callback(101) = "UpdateLightMap slingL_LM_GI_Lights, 100.00, " ' VLM.Lampz;GI.Lights
  Lampz.Callback(101) = "UpdateLightMap slingR_LM_GI_Lights, 100.00, " ' VLM.Lampz;GI.Lights
  Lampz.Callback(101) = "UpdateLightMap wire1_LM_GI_Lights, 100.00, " ' VLM.Lampz;GI.Lights
  Lampz.Callback(101) = "UpdateLightMap wire2_LM_GI_Lights, 100.00, " ' VLM.Lampz;GI.Lights
  Lampz.Callback(101) = "UpdateLightMap wire3_LM_GI_Lights, 100.00, " ' VLM.Lampz;GI.Lights
  Lampz.Callback(101) = "UpdateLightMap wire4_LM_GI_Lights, 100.00, " ' VLM.Lampz;GI.Lights
  Lampz.Callback(101) = "UpdateLightMap wire5_LM_GI_Lights, 100.00, " ' VLM.Lampz;GI.Lights
  Lampz.Callback(101) = "UpdateLightMap wire6_LM_GI_Lights, 100.00, " ' VLM.Lampz;GI.Lights
  Lampz.Callback(101) = "UpdateLightMap wire7_LM_GI_Lights, 100.00, " ' VLM.Lampz;GI.Lights

  Lampz.MassAssign(10) = L10 ' VLM.Lampz;Inserts-L10
  Lampz.Callback(10) = "UpdateLightMap Playfield_LM_Inserts_L10, 100.00, " ' VLM.Lampz;Inserts-L10
  Lampz.Callback(10) = "UpdateLightMap Gate_4_LM_Inserts_L10, 100.00, " ' VLM.Lampz;Inserts-L10
  Lampz.Callback(10) = "UpdateLightMap SpinRT_LM_Inserts_L10, 100.00, " ' VLM.Lampz;Inserts-L10
  Lampz.MassAssign(11) = L11 ' VLM.Lampz;Inserts-L11
  Lampz.Callback(11) = "UpdateLightMap Playfield_LM_Inserts_L11, 100.00, " ' VLM.Lampz;Inserts-L11
  Lampz.Callback(11) = "UpdateLightMap Over_LM_Inserts_L11, 100.00, " ' VLM.Lampz;Inserts-L11
  Lampz.Callback(11) = "UpdateLightMap SpinRB_LM_Inserts_L11, 100.00, " ' VLM.Lampz;Inserts-L11
  Lampz.MassAssign(12) = L12 ' VLM.Lampz;Inserts-L12
  Lampz.Callback(12) = "UpdateLightMap Playfield_LM_Inserts_L12, 100.00, " ' VLM.Lampz;Inserts-L12
  Lampz.Callback(12) = "UpdateLightMap Gate_1_LM_Inserts_L12, 100.00, " ' VLM.Lampz;Inserts-L12
  Lampz.MassAssign(13) = L13 ' VLM.Lampz;Inserts-L13
  Lampz.Callback(13) = "UpdateLightMap Playfield_LM_Inserts_L13, 100.00, " ' VLM.Lampz;Inserts-L13
  Lampz.Callback(13) = "UpdateLightMap BR1_LM_Inserts_L13, 100.00, " ' VLM.Lampz;Inserts-L13
  Lampz.Callback(13) = "UpdateLightMap BR3_LM_Inserts_L13, 100.00, " ' VLM.Lampz;Inserts-L13
  Lampz.MassAssign(14) = L14 ' VLM.Lampz;Inserts-L14
  Lampz.Callback(14) = "UpdateLightMap Playfield_LM_Inserts_L14, 100.00, " ' VLM.Lampz;Inserts-L14
  Lampz.Callback(14) = "UpdateLightMap BR3_LM_Inserts_L14, 100.00, " ' VLM.Lampz;Inserts-L14
  Lampz.MassAssign(15) = L15 ' VLM.Lampz;Inserts-L15
  Lampz.Callback(15) = "UpdateLightMap Playfield_LM_Inserts_L15, 100.00, " ' VLM.Lampz;Inserts-L15
  Lampz.MassAssign(16) = L16 ' VLM.Lampz;Inserts-L16
  Lampz.Callback(16) = "UpdateLightMap Playfield_LM_Inserts_L16, 100.00, " ' VLM.Lampz;Inserts-L16
  Lampz.MassAssign(17) = L17 ' VLM.Lampz;Inserts-L17
  Lampz.Callback(17) = "UpdateLightMap Playfield_LM_Inserts_L17, 100.00, " ' VLM.Lampz;Inserts-L17
  Lampz.MassAssign(18) = L18 ' VLM.Lampz;Inserts-L18
  Lampz.Callback(18) = "UpdateLightMap Playfield_LM_Inserts_L18, 100.00, " ' VLM.Lampz;Inserts-L18
  Lampz.Callback(18) = "UpdateLightMap Gate_5_LM_Inserts_L18, 100.00, " ' VLM.Lampz;Inserts-L18
  Lampz.MassAssign(19) = L19 ' VLM.Lampz;Inserts-L19
  Lampz.Callback(19) = "UpdateLightMap Playfield_LM_Inserts_L19, 100.00, " ' VLM.Lampz;Inserts-L19
  Lampz.Callback(19) = "UpdateLightMap BR1_LM_Inserts_L19, 100.00, " ' VLM.Lampz;Inserts-L19
  Lampz.Callback(19) = "UpdateLightMap HT_LM_Inserts_L19, 100.00, " ' VLM.Lampz;Inserts-L19
  Lampz.MassAssign(20) = L20 ' VLM.Lampz;Inserts-L20
  Lampz.Callback(20) = "UpdateLightMap Playfield_LM_Inserts_L20, 100.00, " ' VLM.Lampz;Inserts-L20
  Lampz.Callback(20) = "UpdateLightMap BR3_LM_Inserts_L20, 100.00, " ' VLM.Lampz;Inserts-L20
  Lampz.Callback(20) = "UpdateLightMap DT1_1_LM_Inserts_L20, 100.00, " ' VLM.Lampz;Inserts-L20
  Lampz.Callback(20) = "UpdateLightMap DT1_2_LM_Inserts_L20, 100.00, " ' VLM.Lampz;Inserts-L20
  Lampz.Callback(20) = "UpdateLightMap DT1_3_LM_Inserts_L20, 100.00, " ' VLM.Lampz;Inserts-L20
  Lampz.MassAssign(21) = L21 ' VLM.Lampz;Inserts-L21
  Lampz.Callback(21) = "UpdateLightMap Playfield_LM_Inserts_L21, 100.00, " ' VLM.Lampz;Inserts-L21
  Lampz.Callback(21) = "UpdateLightMap DT2_1_LM_Inserts_L21, 100.00, " ' VLM.Lampz;Inserts-L21
  Lampz.Callback(21) = "UpdateLightMap DT2_2_LM_Inserts_L21, 100.00, " ' VLM.Lampz;Inserts-L21
  Lampz.Callback(21) = "UpdateLightMap DT2_3_LM_Inserts_L21, 100.00, " ' VLM.Lampz;Inserts-L21
  Lampz.MassAssign(22) = L22 ' VLM.Lampz;Inserts-L22
  Lampz.Callback(22) = "UpdateLightMap Playfield_LM_Inserts_L22, 100.00, " ' VLM.Lampz;Inserts-L22
  Lampz.Callback(22) = "UpdateLightMap DT3_1_LM_Inserts_L22, 100.00, " ' VLM.Lampz;Inserts-L22
  Lampz.Callback(22) = "UpdateLightMap DT3_2_LM_Inserts_L22, 100.00, " ' VLM.Lampz;Inserts-L22
  Lampz.Callback(22) = "UpdateLightMap DT3_3_LM_Inserts_L22, 100.00, " ' VLM.Lampz;Inserts-L22
  Lampz.MassAssign(23) = L23 ' VLM.Lampz;Inserts-L23
  Lampz.Callback(23) = "UpdateLightMap Playfield_LM_Inserts_L23, 100.00, " ' VLM.Lampz;Inserts-L23
  Lampz.Callback(23) = "UpdateLightMap BR1_LM_Inserts_L23, 100.00, " ' VLM.Lampz;Inserts-L23
  Lampz.Callback(23) = "UpdateLightMap DT1_1_LM_Inserts_L23, 100.00, " ' VLM.Lampz;Inserts-L23
  Lampz.Callback(23) = "UpdateLightMap DT1_2_LM_Inserts_L23, 100.00, " ' VLM.Lampz;Inserts-L23
  Lampz.Callback(23) = "UpdateLightMap DT1_3_LM_Inserts_L23, 100.00, " ' VLM.Lampz;Inserts-L23
  Lampz.MassAssign(24) = L24 ' VLM.Lampz;Inserts-L24
  Lampz.Callback(24) = "UpdateLightMap Playfield_LM_Inserts_L24, 100.00, " ' VLM.Lampz;Inserts-L24
  Lampz.Callback(24) = "UpdateLightMap DT2_1_LM_Inserts_L24, 100.00, " ' VLM.Lampz;Inserts-L24
  Lampz.Callback(24) = "UpdateLightMap DT2_2_LM_Inserts_L24, 100.00, " ' VLM.Lampz;Inserts-L24
  Lampz.Callback(24) = "UpdateLightMap DT2_3_LM_Inserts_L24, 100.00, " ' VLM.Lampz;Inserts-L24
  Lampz.Callback(24) = "UpdateLightMap SpinRB_LM_Inserts_L24, 100.00, " ' VLM.Lampz;Inserts-L24
  Lampz.MassAssign(25) = L25 ' VLM.Lampz;Inserts-L25
  Lampz.Callback(25) = "UpdateLightMap Playfield_LM_Inserts_L25, 100.00, " ' VLM.Lampz;Inserts-L25
  Lampz.Callback(25) = "UpdateLightMap BR2_LM_Inserts_L25, 100.00, " ' VLM.Lampz;Inserts-L25
  Lampz.Callback(25) = "UpdateLightMap DT3_1_LM_Inserts_L25, 100.00, " ' VLM.Lampz;Inserts-L25
  Lampz.Callback(25) = "UpdateLightMap DT3_2_LM_Inserts_L25, 100.00, " ' VLM.Lampz;Inserts-L25
  Lampz.Callback(25) = "UpdateLightMap DT3_3_LM_Inserts_L25, 100.00, " ' VLM.Lampz;Inserts-L25
  Lampz.MassAssign(26) = L26 ' VLM.Lampz;Inserts-L26
  Lampz.Callback(26) = "UpdateLightMap Playfield_LM_Inserts_L26, 100.00, " ' VLM.Lampz;Inserts-L26
  Lampz.MassAssign(27) = L27 ' VLM.Lampz;Inserts-L27
  Lampz.Callback(27) = "UpdateLightMap Playfield_LM_Inserts_L27, 100.00, " ' VLM.Lampz;Inserts-L27
  Lampz.MassAssign(28) = L28 ' VLM.Lampz;Inserts-L28
  Lampz.Callback(28) = "UpdateLightMap Playfield_LM_Inserts_L28, 100.00, " ' VLM.Lampz;Inserts-L28
  Lampz.Callback(28) = "UpdateLightMap BR1_LM_Inserts_L28, 100.00, " ' VLM.Lampz;Inserts-L28
  Lampz.Callback(28) = "UpdateLightMap BR2_LM_Inserts_L28, 100.00, " ' VLM.Lampz;Inserts-L28
  Lampz.MassAssign(29) = L29 ' VLM.Lampz;Inserts-L29
  Lampz.Callback(29) = "UpdateLightMap Playfield_LM_Inserts_L29, 100.00, " ' VLM.Lampz;Inserts-L29
  Lampz.Callback(29) = "UpdateLightMap BR2_LM_Inserts_L29, 100.00, " ' VLM.Lampz;Inserts-L29
  Lampz.Callback(29) = "UpdateLightMap BR3_LM_Inserts_L29, 100.00, " ' VLM.Lampz;Inserts-L29
  Lampz.MassAssign(30) = L30 ' VLM.Lampz;Inserts-L30
  Lampz.Callback(30) = "UpdateLightMap Playfield_LM_Inserts_L30, 100.00, " ' VLM.Lampz;Inserts-L30
  Lampz.MassAssign(31) = L31 ' VLM.Lampz;Inserts-L31
  Lampz.Callback(31) = "UpdateLightMap Playfield_LM_Inserts_L31, 100.00, " ' VLM.Lampz;Inserts-L31
  Lampz.MassAssign(32) = L32 ' VLM.Lampz;Inserts-L32
  Lampz.Callback(32) = "UpdateLightMap Playfield_LM_Inserts_L32, 100.00, " ' VLM.Lampz;Inserts-L32
  Lampz.MassAssign(33) = L33 ' VLM.Lampz;Inserts-L33
  Lampz.Callback(33) = "UpdateLightMap Playfield_LM_Inserts_L33, 100.00, " ' VLM.Lampz;Inserts-L33
  Lampz.Callback(33) = "UpdateLightMap BR2_LM_Inserts_L33, 100.00, " ' VLM.Lampz;Inserts-L33
  Lampz.MassAssign(34) = L34 ' VLM.Lampz;Inserts-L34
  Lampz.Callback(34) = "UpdateLightMap Playfield_LM_Inserts_L34, 100.00, " ' VLM.Lampz;Inserts-L34
  Lampz.MassAssign(35) = L35 ' VLM.Lampz;Inserts-L35
  Lampz.Callback(35) = "UpdateLightMap Playfield_LM_Inserts_L35, 100.00, " ' VLM.Lampz;Inserts-L35
  Lampz.MassAssign(36) = L36 ' VLM.Lampz;Inserts-L36
  Lampz.Callback(36) = "UpdateLightMap Playfield_LM_Inserts_L36, 100.00, " ' VLM.Lampz;Inserts-L36
  Lampz.MassAssign(37) = L37 ' VLM.Lampz;Inserts-L37
  Lampz.Callback(37) = "UpdateLightMap Playfield_LM_Inserts_L37, 100.00, " ' VLM.Lampz;Inserts-L37
  Lampz.MassAssign(38) = L38 ' VLM.Lampz;Inserts-L38
  Lampz.Callback(38) = "UpdateLightMap Playfield_LM_Inserts_L38, 100.00, " ' VLM.Lampz;Inserts-L38
  Lampz.MassAssign(39) = L39 ' VLM.Lampz;Inserts-L39
  Lampz.Callback(39) = "UpdateLightMap Playfield_LM_Inserts_L39, 100.00, " ' VLM.Lampz;Inserts-L39
  Lampz.Callback(39) = "UpdateLightMap BR1_LM_Inserts_L39, 100.00, " ' VLM.Lampz;Inserts-L39
  Lampz.Callback(39) = "UpdateLightMap HT_LM_Inserts_L39, 100.00, " ' VLM.Lampz;Inserts-L39
  Lampz.MassAssign(40) = L40 ' VLM.Lampz;Inserts-L40
  Lampz.Callback(40) = "UpdateLightMap Playfield_LM_Inserts_L40, 100.00, " ' VLM.Lampz;Inserts-L40
  Lampz.MassAssign(41) = L41 ' VLM.Lampz;Inserts-L41
  Lampz.Callback(41) = "UpdateLightMap Playfield_LM_Inserts_L41, 100.00, " ' VLM.Lampz;Inserts-L41
  Lampz.MassAssign(42) = L42 ' VLM.Lampz;Inserts-L42
  Lampz.Callback(42) = "UpdateLightMap Playfield_LM_Inserts_L42, 100.00, " ' VLM.Lampz;Inserts-L42
  Lampz.MassAssign(43) = L43 ' VLM.Lampz;Inserts-L43
  Lampz.Callback(43) = "UpdateLightMap Playfield_LM_Inserts_L43, 100.00, " ' VLM.Lampz;Inserts-L43
  Lampz.MassAssign(44) = L44 ' VLM.Lampz;Inserts-L44
  Lampz.Callback(44) = "UpdateLightMap Playfield_LM_Inserts_L44, 100.00, " ' VLM.Lampz;Inserts-L44
  Lampz.MassAssign(45) = L45 ' VLM.Lampz;Inserts-L45
  Lampz.Callback(45) = "UpdateLightMap Playfield_LM_Inserts_L45, 100.00, " ' VLM.Lampz;Inserts-L45
  Lampz.MassAssign(46) = L46 ' VLM.Lampz;Inserts-L46
  Lampz.Callback(46) = "UpdateLightMap Playfield_LM_Inserts_L46, 100.00, " ' VLM.Lampz;Inserts-L46
  Lampz.MassAssign(47) = L47 ' VLM.Lampz;Inserts-L47
  Lampz.Callback(47) = "UpdateLightMap Playfield_LM_Inserts_L47, 100.00, " ' VLM.Lampz;Inserts-L47
  Lampz.MassAssign(49) = L49 ' VLM.Lampz;Inserts-L49
  Lampz.Callback(49) = "UpdateLightMap Playfield_LM_Inserts_L49, 100.00, " ' VLM.Lampz;Inserts-L49
  Lampz.MassAssign(50) = L50 ' VLM.Lampz;Inserts-L50
  Lampz.Callback(50) = "UpdateLightMap Playfield_LM_Inserts_L50, 100.00, " ' VLM.Lampz;Inserts-L50
  Lampz.MassAssign(51) = L51 ' VLM.Lampz;Inserts-L51
  Lampz.Callback(51) = "UpdateLightMap Playfield_LM_Inserts_L51, 100.00, " ' VLM.Lampz;Inserts-L51
  Lampz.MassAssign(52) = L52 ' VLM.Lampz;Inserts-L52
  Lampz.Callback(52) = "UpdateLightMap Playfield_LM_Inserts_L52, 100.00, " ' VLM.Lampz;Inserts-L52
  Lampz.MassAssign(53) = L53 ' VLM.Lampz;Inserts-L53
  Lampz.Callback(53) = "UpdateLightMap Playfield_LM_Inserts_L53, 100.00, " ' VLM.Lampz;Inserts-L53
  Lampz.MassAssign(54) = L54 ' VLM.Lampz;Inserts-L54
  Lampz.Callback(54) = "UpdateLightMap Playfield_LM_Inserts_L54, 100.00, " ' VLM.Lampz;Inserts-L54
  Lampz.MassAssign(55) = L55 ' VLM.Lampz;Inserts-L55
  Lampz.Callback(55) = "UpdateLightMap Playfield_LM_Inserts_L55, 100.00, " ' VLM.Lampz;Inserts-L55
  Lampz.MassAssign(57) = L57 ' VLM.Lampz;Inserts-L57
  Lampz.Callback(57) = "UpdateLightMap Playfield_LM_Inserts_L57, 100.00, " ' VLM.Lampz;Inserts-L57
  Lampz.Callback(57) = "UpdateLightMap slingL_LM_Inserts_L57, 100.00, " ' VLM.Lampz;Inserts-L57
  Lampz.MassAssign(58) = L58 ' VLM.Lampz;Inserts-L58
  Lampz.Callback(58) = "UpdateLightMap Playfield_LM_Inserts_L58, 100.00, " ' VLM.Lampz;Inserts-L58
  Lampz.Callback(58) = "UpdateLightMap BR3_LM_Inserts_L58, 100.00, " ' VLM.Lampz;Inserts-L58
  Lampz.MassAssign(59) = L59 ' VLM.Lampz;Inserts-L59
  Lampz.Callback(59) = "UpdateLightMap Playfield_LM_Inserts_L59, 100.00, " ' VLM.Lampz;Inserts-L59
  Lampz.Callback(59) = "UpdateLightMap BR3_LM_Inserts_L59, 100.00, " ' VLM.Lampz;Inserts-L59
  Lampz.MassAssign(60) = L60 ' VLM.Lampz;Inserts-L60
  Lampz.Callback(60) = "UpdateLightMap Playfield_LM_Inserts_L60, 100.00, " ' VLM.Lampz;Inserts-L60
  Lampz.Callback(60) = "UpdateLightMap BR2_LM_Inserts_L60, 100.00, " ' VLM.Lampz;Inserts-L60
  Lampz.MassAssign(61) = L61 ' VLM.Lampz;Inserts-L61
  Lampz.Callback(61) = "UpdateLightMap Playfield_LM_Inserts_L61, 100.00, " ' VLM.Lampz;Inserts-L61
  Lampz.MassAssign(62) = L62 ' VLM.Lampz;Inserts-L62
  Lampz.Callback(62) = "UpdateLightMap Playfield_LM_Inserts_L62, 100.00, " ' VLM.Lampz;Inserts-L62
  Lampz.Callback(62) = "UpdateLightMap BR2_LM_Inserts_L62, 100.00, " ' VLM.Lampz;Inserts-L62
  Lampz.MassAssign(63) = L63 ' VLM.Lampz;Inserts-L63
  Lampz.Callback(63) = "UpdateLightMap Playfield_LM_Inserts_L63, 100.00, " ' VLM.Lampz;Inserts-L63
  Lampz.MassAssign(8) = L8 ' VLM.Lampz;Inserts-L8
  Lampz.Callback(8) = "UpdateLightMap Playfield_LM_Inserts_L8, 100.00, " ' VLM.Lampz;Inserts-L8
  Lampz.Callback(8) = "UpdateLightMap BR2_LM_Inserts_L8, 100.00, " ' VLM.Lampz;Inserts-L8
  Lampz.MassAssign(9) = L9 ' VLM.Lampz;Inserts-L9
  Lampz.Callback(9) = "UpdateLightMap Playfield_LM_Inserts_L9, 100.00, " ' VLM.Lampz;Inserts-L9
  Lampz.Callback(9) = "UpdateLightMap BR1_LM_Inserts_L9, 100.00, " ' VLM.Lampz;Inserts-L9
  Lampz.Callback(9) = "UpdateLightMap BR3_LM_Inserts_L9, 100.00, " ' VLM.Lampz;Inserts-L9
  Lampz.Callback(9) = "UpdateLightMap SpinL_LM_Inserts_L9, 100.00, " ' VLM.Lampz;Inserts-L9

  'Turn off all lamps on startup
  Lampz.Init  'This just turns state of any lamps to 1

  'Immediate update to turn on GI, turn off lamps
  Lampz.Update
End Sub


'====================
'Class jungle nf
'====================

'No-op object instead of adding more conditionals to the main loop
'It also prevents errors if empty lamp numbers are called, and it's only one object
'should be g2g?

Class NullFadingObject : Public Property Let IntensityScale(input) : : End Property : End Class

'version 0.11 - Mass Assign, Changed modulate style
'version 0.12 - Update2 (single -1 timer update) update method for core.vbs
'Version 0.12a - Filter can now be accessed via 'FilterOut'
'Version 0.12b - Changed MassAssign from a sub to an indexed property (new syntax: lampfader.MassAssign(15) = Light1 )
'Version 0.13 - No longer requires setlocale. Callback() can be assigned multiple times per index
' Note: if using multiple 'LampFader' objects, set the 'name' variable to avoid conflicts with callbacks

Class LampFader
  Public FadeSpeedDown(150), FadeSpeedUp(150)
  Private Lock(150), Loaded(150), OnOff(150)
  Public UseFunction
  Private cFilter
  Public UseCallback(150), cCallback(150)
  Public Lvl(150), Obj(150)
  Private Mult(150)
  Public FrameTime
  Private InitFrame
  Public Name

  Sub Class_Initialize()
    InitFrame = 0
    dim x : for x = 0 to uBound(OnOff)  'Set up fade speeds
      FadeSpeedDown(x) = 1/100  'fade speed down
      FadeSpeedUp(x) = 1/80   'Fade speed up
      UseFunction = False
      lvl(x) = 0
      OnOff(x) = 0
      Lock(x) = True : Loaded(x) = False
      Mult(x) = 1
    Next
    Name = "LampFaderNF" 'NEEDS TO BE CHANGED IF THERE'S MULTIPLE OF THESE OBJECTS, OTHERWISE CALLBACKS WILL INTERFERE WITH EACH OTHER!!
    for x = 0 to uBound(OnOff)    'clear out empty obj
      if IsEmpty(obj(x) ) then Set Obj(x) = NullFader' : Loaded(x) = True
    Next
  End Sub

  Public Property Get Locked(idx) : Locked = Lock(idx) : End Property   ''debug.print Lampz.Locked(100) 'debug
  Public Property Get state(idx) : state = OnOff(idx) : end Property
  Public Property Let Filter(String) : Set cFilter = GetRef(String) : UseFunction = True : End Property
  Public Function FilterOut(aInput) : if UseFunction Then FilterOut = cFilter(aInput) Else FilterOut = aInput End If : End Function
  'Public Property Let Callback(idx, String) : cCallback(idx) = String : UseCallBack(idx) = True : End Property
  Public Property Let Callback(idx, String)
    UseCallBack(idx) = True
    'cCallback(idx) = String 'old execute method
    'New method: build wrapper subs using ExecuteGlobal, then call them
    cCallback(idx) = cCallback(idx) & "___" & String  'multiple strings dilineated by 3x _

    dim tmp : tmp = Split(cCallback(idx), "___")

    dim str, x : for x = 0 to uBound(tmp) 'build proc contents
      'If Not tmp(x)="" then str = str & "  " & tmp(x) & " aLVL" & "  '" & x & vbnewline  'more verbose
      If Not tmp(x)="" then str = str & tmp(x) & " aLVL:"
    Next
    'msgbox "Sub " & name & idx & "(aLvl):" & str & "End Sub"
    dim out : out = "Sub " & name & idx & "(aLvl):" & str & "End Sub"
    ExecuteGlobal Out

  End Property

  Public Property Let state(ByVal idx, input) 'Major update path
    if TypeName(input) <> "Double" and typename(input) <> "Integer"  and typename(input) <> "Long" then
      If input Then
        input = 1
      Else
        input = 0
      End If
    End If
    if Input <> OnOff(idx) then  'discard redundant updates
      OnOff(idx) = input
      Lock(idx) = False
      Loaded(idx) = False
    End If
  End Property

  'Mass assign, Builds arrays where necessary
  'Sub MassAssign(aIdx, aInput)
  Public Property Let MassAssign(aIdx, aInput)
    If typename(obj(aIdx)) = "NullFadingObject" Then 'if empty, use Set
      if IsArray(aInput) then
        obj(aIdx) = aInput
      Else
        Set obj(aIdx) = aInput
      end if
    Else
      Obj(aIdx) = AppendArray(obj(aIdx), aInput)
    end if
  end Property

  Sub SetLamp(aIdx, aOn) : state(aIdx) = aOn : End Sub  'Solenoid Handler

  Public Sub TurnOnStates() 'If obj contains any light objects, set their states to 1 (Fading is our job!)
    dim debugstr
    dim idx : for idx = 0 to uBound(obj)
      if IsArray(obj(idx)) then
        'debugstr = debugstr & "array found at " & idx & "..."
        dim x, tmp : tmp = obj(idx) 'set tmp to array in order to access it
        for x = 0 to uBound(tmp)
          if typename(tmp(x)) = "Light" then DisableState tmp(x)' : debugstr = debugstr & tmp(x).name & " state'd" & vbnewline
          tmp(x).intensityscale = 0.001 ' this can prevent init stuttering
        Next
      Else
        if typename(obj(idx)) = "Light" then DisableState obj(idx)' : debugstr = debugstr & obj(idx).name & " state'd (not array)" & vbnewline
        obj(idx).intensityscale = 0.001 ' this can prevent init stuttering
      end if
    Next
    ''debug.print debugstr
  End Sub
  Private Sub DisableState(ByRef aObj) : aObj.FadeSpeedUp = 1000 : aObj.State = 1 : End Sub 'turn state to 1

  Public Sub Init() 'Just runs TurnOnStates right now
    TurnOnStates
  End Sub

  Public Property Let Modulate(aIdx, aCoef) : Mult(aIdx) = aCoef : Lock(aIdx) = False : Loaded(aIdx) = False: End Property
  Public Property Get Modulate(aIdx) : Modulate = Mult(aIdx) : End Property

  Public Sub Update1()   'Handle all boolean numeric fading. If done fading, Lock(x) = True. Update on a '1' interval Timer!
    dim x : for x = 0 to uBound(OnOff)
      if not Lock(x) then 'and not Loaded(x) then
        if OnOff(x) > 0 then 'Fade Up
          Lvl(x) = Lvl(x) + FadeSpeedUp(x)
          if Lvl(x) >= OnOff(x) then Lvl(x) = OnOff(x) : Lock(x) = True
        else 'fade down
          Lvl(x) = Lvl(x) - FadeSpeedDown(x)
          if Lvl(x) <= 0 then Lvl(x) = 0 : Lock(x) = True
        end if
      end if
    Next
  End Sub

  Public Sub Update2()   'Both updates on -1 timer (Lowest latency, but less accurate fading at 60fps vsync)
    FrameTime = gametime - InitFrame : InitFrame = GameTime 'Calculate frametime
    dim x : for x = 0 to uBound(OnOff)
      if not Lock(x) then 'and not Loaded(x) then
        if OnOff(x) > 0 then 'Fade Up
          Lvl(x) = Lvl(x) + FadeSpeedUp(x) * FrameTime
          if Lvl(x) >= OnOff(x) then Lvl(x) = OnOff(x) : Lock(x) = True
        else 'fade down
          Lvl(x) = Lvl(x) - FadeSpeedDown(x) * FrameTime
          if Lvl(x) <= 0 then Lvl(x) = 0 : Lock(x) = True
        end if
      end if
    Next
    Update
  End Sub

  Public Sub Update() 'Handle object updates. Update on a -1 Timer! If done fading, loaded(x) = True
    dim x,xx, aLvl : for x = 0 to uBound(OnOff)
      if not Loaded(x) then
        aLvl = Lvl(x)*Mult(x)
        if IsArray(obj(x) ) Then  'if array
          If UseFunction then
            for each xx in obj(x) : xx.IntensityScale = cFilter(aLvl) : Next
          Else
            for each xx in obj(x) : xx.IntensityScale = aLvl : Next
          End If
        else            'if single lamp or flasher
          If UseFunction then
            obj(x).Intensityscale = cFilter(aLvl)
          Else
            obj(x).Intensityscale = aLvl
          End If
        end if
        'if TypeName(lvl(x)) <> "Double" and typename(lvl(x)) <> "Integer" and typename(lvl(x)) <> "Long" then msgbox "uhh " & 2 & " = " & lvl(x)
        'If UseCallBack(x) then execute cCallback(x) & " " & (Lvl(x)) 'Callback
        If UseCallBack(x) then Proc name & x,aLvl 'Proc
        If Lock(x) Then
          if Lvl(x) = OnOff(x) or Lvl(x) = 0 then Loaded(x) = True  'finished fading
        end if
      end if
    Next
  End Sub
End Class


'Lamp Filter
Function LampFilter(aLvl)
  LampFilter = aLvl^1.6 'exponential curve?
End Function


'Helper functions
Sub Proc(string, Callback)  'proc using a string and one argument
  'On Error Resume Next
  dim p : Set P = GetRef(String)
  P Callback
  If err.number = 13 then  msgbox "Proc error! No such procedure: " & vbnewline & string
  if err.number = 424 then msgbox "Proc error! No such Object"
End Sub

Function AppendArray(ByVal aArray, aInput)  'append one value, object, or Array onto the end of a 1 dimensional array
  if IsArray(aInput) then 'Input is an array...
    dim tmp : tmp = aArray
    If not IsArray(aArray) Then 'if not array, create an array
      tmp = aInput
    Else            'Append existing array with aInput array
      Redim Preserve tmp(uBound(aArray) + uBound(aInput)+1) 'If existing array, increase bounds by uBound of incoming array
      dim x : for x = 0 to uBound(aInput)
        if isObject(aInput(x)) then
          Set tmp(x+uBound(aArray)+1 ) = aInput(x)
        Else
          tmp(x+uBound(aArray)+1 ) = aInput(x)
        End If
      Next
      AppendArray = tmp  'return new array
    End If
  Else 'Input is NOT an array...
    If not IsArray(aArray) Then 'if not array, create an array
      aArray = Array(aArray, aInput)
    Else
      Redim Preserve aArray(uBound(aArray)+1) 'If array, increase bounds by 1
      if isObject(aInput) then
        Set aArray(uBound(aArray)) = aInput
      Else
        aArray(uBound(aArray)) = aInput
      End If
    End If
    AppendArray = aArray 'return new array
  End If
End Function

'******************************************************
'****  END LAMPZ
'******************************************************




'////////////////////////////  LIGHTMAP SUPPORT  ///////////////////////////

Dim LightLevel : LightLevel = 50

Sub UpdateLightLevel
  Lampz.state(100) = LightLevel
End Sub

Sub UpdateBR1
  Dim r : r = BR1_BM_Dark_Room.rotx
  BR1_LM_Lit_Room.rotx = r ' VLM.Props;LM;1;BR1
  BR1_LM_GI_Lights.rotx = r ' VLM.Props;LM;1;BR1
  BR1_LM_Inserts_L13.rotx = r ' VLM.Props;LM;1;BR1
  BR1_LM_Inserts_L19.rotx = r ' VLM.Props;LM;1;BR1
  BR1_LM_Inserts_L23.rotx = r ' VLM.Props;LM;1;BR1
  BR1_LM_Inserts_L28.rotx = r ' VLM.Props;LM;1;BR1
  BR1_LM_Inserts_L39.rotx = r ' VLM.Props;LM;1;BR1
  BR1_LM_Inserts_L9.rotx = r ' VLM.Props;LM;1;BR1
  r = BR1_BM_Dark_Room.roty
  BR1_LM_Lit_Room.roty = r ' VLM.Props;LM;2;BR1
  BR1_LM_GI_Lights.roty = r ' VLM.Props;LM;2;BR1
  BR1_LM_Inserts_L13.roty = r ' VLM.Props;LM;2;BR1
  BR1_LM_Inserts_L19.roty = r ' VLM.Props;LM;2;BR1
  BR1_LM_Inserts_L23.roty = r ' VLM.Props;LM;2;BR1
  BR1_LM_Inserts_L28.roty = r ' VLM.Props;LM;2;BR1
  BR1_LM_Inserts_L39.roty = r ' VLM.Props;LM;2;BR1
  BR1_LM_Inserts_L9.roty = r ' VLM.Props;LM;2;BR1
End Sub

Sub UpdateBR2
  Dim r : r = BR2_BM_Dark_Room.rotx
  BR2_LM_Lit_Room.rotx = r ' VLM.Props;LM;1;BR2
  BR2_LM_GI_Lights.rotx = r ' VLM.Props;LM;1;BR2
  BR2_LM_Inserts_L25.rotx = r ' VLM.Props;LM;1;BR2
  BR2_LM_Inserts_L28.rotx = r ' VLM.Props;LM;1;BR2
  BR2_LM_Inserts_L29.rotx = r ' VLM.Props;LM;1;BR2
  BR2_LM_Inserts_L33.rotx = r ' VLM.Props;LM;1;BR2
  BR2_LM_Inserts_L60.rotx = r ' VLM.Props;LM;1;BR2
  BR2_LM_Inserts_L62.rotx = r ' VLM.Props;LM;1;BR2
  BR2_LM_Inserts_L8.rotx = r ' VLM.Props;LM;1;BR2
  r = BR2_BM_Dark_Room.roty
  BR2_LM_Lit_Room.roty = r ' VLM.Props;LM;2;BR2
  BR2_LM_GI_Lights.roty = r ' VLM.Props;LM;2;BR2
  BR2_LM_Inserts_L25.roty = r ' VLM.Props;LM;2;BR2
  BR2_LM_Inserts_L28.roty = r ' VLM.Props;LM;2;BR2
  BR2_LM_Inserts_L29.roty = r ' VLM.Props;LM;2;BR2
  BR2_LM_Inserts_L33.roty = r ' VLM.Props;LM;2;BR2
  BR2_LM_Inserts_L60.roty = r ' VLM.Props;LM;2;BR2
  BR2_LM_Inserts_L62.roty = r ' VLM.Props;LM;2;BR2
  BR2_LM_Inserts_L8.roty = r ' VLM.Props;LM;2;BR2
End Sub

Sub UpdateBR3
  Dim r : r = BR3_BM_Dark_Room.rotx
  BR3_LM_Lit_Room.rotx = r ' VLM.Props;LM;1;BR3
  BR3_LM_GI_Lights.rotx = r ' VLM.Props;LM;1;BR3
  BR3_LM_Inserts_L13.rotx = r ' VLM.Props;LM;1;BR3
  BR3_LM_Inserts_L14.rotx = r ' VLM.Props;LM;1;BR3
  BR3_LM_Inserts_L20.rotx = r ' VLM.Props;LM;1;BR3
  BR3_LM_Inserts_L29.rotx = r ' VLM.Props;LM;1;BR3
  BR3_LM_Inserts_L58.rotx = r ' VLM.Props;LM;1;BR3
  BR3_LM_Inserts_L59.rotx = r ' VLM.Props;LM;1;BR3
  BR3_LM_Inserts_L9.rotx = r ' VLM.Props;LM;1;BR3
  r = BR3_BM_Dark_Room.roty
  BR3_LM_Lit_Room.roty = r ' VLM.Props;LM;2;BR3
  BR3_LM_GI_Lights.roty = r ' VLM.Props;LM;2;BR3
  BR3_LM_Inserts_L13.roty = r ' VLM.Props;LM;2;BR3
  BR3_LM_Inserts_L14.roty = r ' VLM.Props;LM;2;BR3
  BR3_LM_Inserts_L20.roty = r ' VLM.Props;LM;2;BR3
  BR3_LM_Inserts_L29.roty = r ' VLM.Props;LM;2;BR3
  BR3_LM_Inserts_L58.roty = r ' VLM.Props;LM;2;BR3
  BR3_LM_Inserts_L59.roty = r ' VLM.Props;LM;2;BR3
  BR3_LM_Inserts_L9.roty = r ' VLM.Props;LM;2;BR3
End Sub

Sub UpdateTargets
  Dim z
  z = DT1_1_BM_Dark_Room.transz
  DT1_1_LM_Lit_Room.transz = z ' VLM.Props;LM;1;DT1.1
  DT1_1_LM_GI_Lights.transz = z ' VLM.Props;LM;1;DT1.1
  DT1_1_LM_Inserts_L20.transz = z ' VLM.Props;LM;1;DT1.1
  DT1_1_LM_Inserts_L23.transz = z ' VLM.Props;LM;1;DT1.1
  z = DT1_2_BM_Dark_Room.transz
  DT1_2_LM_Lit_Room.transz = z ' VLM.Props;LM;1;DT1.2
  DT1_2_LM_GI_Lights.transz = z ' VLM.Props;LM;1;DT1.2
  DT1_2_LM_Inserts_L20.transz = z ' VLM.Props;LM;1;DT1.2
  DT1_2_LM_Inserts_L23.transz = z ' VLM.Props;LM;1;DT1.2
  z = DT1_3_BM_Dark_Room.transz
  DT1_3_LM_Lit_Room.transz = z ' VLM.Props;LM;1;DT1.3
  DT1_3_LM_GI_Lights.transz = z ' VLM.Props;LM;1;DT1.3
  DT1_3_LM_Inserts_L20.transz = z ' VLM.Props;LM;1;DT1.3
  DT1_3_LM_Inserts_L23.transz = z ' VLM.Props;LM;1;DT1.3
  z = DT2_1_BM_Dark_Room.transz
  DT2_1_LM_Lit_Room.transz = z ' VLM.Props;LM;1;DT2.1
  DT2_1_LM_GI_Lights.transz = z ' VLM.Props;LM;1;DT2.1
  DT2_1_LM_Inserts_L21.transz = z ' VLM.Props;LM;1;DT2.1
  DT2_1_LM_Inserts_L24.transz = z ' VLM.Props;LM;1;DT2.1
  z = DT2_2_BM_Dark_Room.transz
  DT2_2_LM_Lit_Room.transz = z ' VLM.Props;LM;1;DT2.2
  DT2_2_LM_GI_Lights.transz = z ' VLM.Props;LM;1;DT2.2
  DT2_2_LM_Inserts_L21.transz = z ' VLM.Props;LM;1;DT2.2
  DT2_2_LM_Inserts_L24.transz = z ' VLM.Props;LM;1;DT2.2
  z = DT2_3_BM_Dark_Room.transz
  DT2_3_LM_Lit_Room.transz = z ' VLM.Props;LM;1;DT2.3
  DT2_3_LM_GI_Lights.transz = z ' VLM.Props;LM;1;DT2.3
  DT2_3_LM_Inserts_L21.transz = z ' VLM.Props;LM;1;DT2.3
  DT2_3_LM_Inserts_L24.transz = z ' VLM.Props;LM;1;DT2.3
  z = DT3_1_BM_Dark_Room.transz
  DT3_1_LM_Lit_Room.transz = z ' VLM.Props;LM;1;DT3.1
  DT3_1_LM_GI_Lights.transz = z ' VLM.Props;LM;1;DT3.1
  DT3_1_LM_Inserts_L22.transz = z ' VLM.Props;LM;1;DT3.1
  DT3_1_LM_Inserts_L25.transz = z ' VLM.Props;LM;1;DT3.1
  z = DT3_2_BM_Dark_Room.transz
  DT3_2_LM_Lit_Room.transz = z ' VLM.Props;LM;1;DT3.2
  DT3_2_LM_GI_Lights.transz = z ' VLM.Props;LM;1;DT3.2
  DT3_2_LM_Inserts_L22.transz = z ' VLM.Props;LM;1;DT3.2
  DT3_2_LM_Inserts_L25.transz = z ' VLM.Props;LM;1;DT3.2
  z = DT3_3_BM_Dark_Room.transz
  DT3_3_LM_Lit_Room.transz = z ' VLM.Props;LM;1;DT3.3
  DT3_3_LM_GI_Lights.transz = z ' VLM.Props;LM;1;DT3.3
  DT3_3_LM_Inserts_L22.transz = z ' VLM.Props;LM;1;DT3.3
  DT3_3_LM_Inserts_L25.transz = z ' VLM.Props;LM;1;DT3.3
End Sub

Sub UpdateFlipper
  Dim r

  r = LeftFlipper.currentangle
  FlipperL_BM_Dark_Room.ObjRotZ = r
  FlipperL_LM_Lit_Room.ObjRotZ = r ' VLM.Props;LM;1;FlipperL
  FlipperL_LM_GI_Lights.ObjRotZ = r ' VLM.Props;LM;1;FlipperL

  r = RightFlipper.currentangle
  FlipperR_BM_Dark_Room.ObjRotZ = r
  FlipperR_LM_Lit_Room.ObjRotZ = r ' VLM.Props;LM;1;FlipperR
  FlipperR_LM_GI_Lights.ObjRotZ = r ' VLM.Props;LM;1;FlipperR
End Sub

Sub UpdateGates
  Dim r

  r = Gate001.CurrentAngle
  Gate_1_BM_Dark_Room.roty = r
  Gate_1_LM_Lit_Room.rotx = r ' VLM.Props;LM;1;Gate.1
  Gate_1_LM_GI_Lights.rotx = r ' VLM.Props;LM;1;Gate.1
  Gate_1_LM_Inserts_L12.rotx = r ' VLM.Props;LM;1;Gate.1

  r = Gate004.CurrentAngle
  Gate_2_BM_Dark_Room.roty = r
  Gate_2_LM_Lit_Room.roty = r ' VLM.Props;LM;1;Gate.2
  Gate_2_LM_GI_Lights.roty = r ' VLM.Props;LM;1;Gate.2

  r = Gate005.CurrentAngle
  Gate_3_BM_Dark_Room.roty = r
  Gate_3_LM_Lit_Room.roty = r ' VLM.Props;LM;1;Gate.3
  Gate_3_LM_GI_Lights.roty = r ' VLM.Props;LM;1;Gate.3

  r = Gate003.CurrentAngle
  Gate_4_BM_Dark_Room.roty = r
  Gate_4_LM_Lit_Room.roty = r ' VLM.Props;LM;1;Gate.4
  Gate_4_LM_GI_Lights.roty = r ' VLM.Props;LM;1;Gate.4
  Gate_4_LM_Inserts_L10.roty = r ' VLM.Props;LM;1;Gate.4

  r = -Gate006.CurrentAngle
  Gate_5_BM_Dark_Room.roty = r
  Gate_5_LM_Lit_Room.roty = r ' VLM.Props;LM;1;Gate.5
  Gate_5_LM_GI_Lights.roty = r ' VLM.Props;LM;1;Gate.5
  Gate_5_LM_Inserts_L18.roty = r ' VLM.Props;LM;1;Gate.5

  r = Gate002.CurrentAngle
  Gate_6_BM_Dark_Room.roty = r
  Gate_6_LM_Lit_Room.roty = r ' VLM.Props;LM;1;Gate.6
  Gate_6_LM_GI_Lights.roty = r ' VLM.Props;LM;1;Gate.6
End Sub

Sub UpdateHitTarget
  Dim y : y = HT_BM_Dark_Room.TransY
  HT_LM_Lit_Room.transy = y ' VLM.Props;LM;1;HT
  HT_LM_GI_Lights.transy = y ' VLM.Props;LM;1;HT
  HT_LM_Inserts_L19.transy = y ' VLM.Props;LM;1;HT
  HT_LM_Inserts_L39.transy = y ' VLM.Props;LM;1;HT
End Sub

Sub UpdateSpinners
  Dim r

  r = sw10.CurrentAngle
  SpinL_BM_Dark_Room.rotx = r
  SpinL_LM_Lit_Room.rotx = r ' VLM.Props;LM;1;SpinL
  SpinL_LM_GI_Lights.rotx = r ' VLM.Props;LM;1;SpinL
  SpinL_LM_Inserts_L9.rotx = r ' VLM.Props;LM;1;SpinL

  r = sw12.CurrentAngle
  SpinRB_BM_Dark_Room.rotx = r
  SpinRB_LM_Lit_Room.rotx = r ' VLM.Props;LM;1;SpinRB
  SpinRB_LM_GI_Lights.rotx = r ' VLM.Props;LM;1;SpinRB
  SpinRB_LM_Inserts_L11.rotx = r ' VLM.Props;LM;1;SpinRB
  SpinRB_LM_Inserts_L24.rotx = r ' VLM.Props;LM;1;SpinRB

  r = sw11.CurrentAngle
  SpinRT_BM_Dark_Room.rotx = r
  SpinRT_LM_Lit_Room.rotx = r ' VLM.Props;LM;1;SpinRT
  SpinRT_LM_GI_Lights.rotx = r ' VLM.Props;LM;1;SpinRT
  SpinRT_LM_Inserts_L10.rotx = r ' VLM.Props;LM;1;SpinRT
End Sub

Sub UpdateSlings
  Dim z

  z = slingL_BM_Dark_Room.TransZ
  slingL_LM_Lit_Room.TransZ = z ' VLM.Props;LM;1;slingL
  slingL_LM_GI_Lights.TransZ = z ' VLM.Props;LM;1;slingL
  slingL_LM_Inserts_L57.TransZ = z ' VLM.Props;LM;1;slingL

  z = slingR_BM_Dark_Room.TransZ
  slingR_LM_Lit_Room.TransZ = z ' VLM.Props;LM;1;slingR
  slingR_LM_GI_Lights.TransZ = z ' VLM.Props;LM;1;slingR
End Sub

Sub UpdateWires
  Dim z
  z = wire1_BM_Dark_Room.transz
  wire1_LM_Lit_Room.transz = z ' VLM.Props;LM;1;wire1
  wire1_LM_GI_Lights.transz = z ' VLM.Props;LM;1;wire1
  z = wire2_BM_Dark_Room.transz
  wire2_LM_Lit_Room.transz = z ' VLM.Props;LM;1;wire2
  wire2_LM_GI_Lights.transz = z ' VLM.Props;LM;1;wire2
  z = wire3_BM_Dark_Room.transz
  wire3_LM_Lit_Room.transz = z ' VLM.Props;LM;1;wire3
  wire3_LM_GI_Lights.transz = z ' VLM.Props;LM;1;wire3
  z = wire4_BM_Dark_Room.transz
  wire4_LM_Lit_Room.transz = z ' VLM.Props;LM;1;wire4
  wire4_LM_GI_Lights.transz = z ' VLM.Props;LM;1;wire4
  z = wire5_BM_Dark_Room.transz
  wire5_LM_GI_Lights.transz = z ' VLM.Props;LM;1;wire5
  z = wire6_BM_Dark_Room.transz
  wire6_LM_Lit_Room.transz = z ' VLM.Props;LM;1;wire6
  wire6_LM_GI_Lights.transz = z ' VLM.Props;LM;1;wire6
  z = wire7_BM_Dark_Room.transz
  wire7_LM_Lit_Room.transz = z ' VLM.Props;LM;1;wire7
  wire7_LM_GI_Lights.transz = z ' VLM.Props;LM;1;wire7
End Sub

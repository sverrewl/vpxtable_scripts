'############################################################################################
'############################################################################################
'#######                                                                             ########
'#######                          Nemesis  (Peyper 1986)                             ########
'#######                                                                             ########
'############################################################################################
'############################################################################################

' VPX version by Batch 2020

' VP9 version by MFuegemann (2014)

option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

LoadVPM "01560000","PEYPER.VBS",3.2

Dim DesktopMode: DesktopMode = Nemesis.ShowDT

If DesktopMode = True Then 'Show Desktop components
  lrail.visible=1
  rrail.visible=1
  SideCab.visible=1
Else
  lrail.visible=0
  rrail.visible=0
  SideCab.visible=0
End if

'*************************************
'        GLOBAL CONFIGURATIONS
'*************************************

Const cgamename = "nemesis"   'name of ROM file to be used
Const FreePlay = 1        'set to 1 for a faked FreePlay / otherwise, set to 0

Const UseSolenoids=2,UseLamps=1,SSolenoidOn="SolOn",SSolenoidOff="SolOff",SFlipperOn="FlipperUp",SFlipperOff="FlipperDown",SCoin="coin3"

'***********************************
'        SOLENOID ASSIGNMENT
'***********************************

'Sol1 LSling
SolCallback(6)="DropTargetBank1.SolDropUp"
'Sol3 RSling
'Sol4 Bumper
'Sol5 Bumper
SolCallback(2)="DropTargetBank2.SolDropUp"
SolCallback(7)="SolKnocker"
SolCallback(8)="bsTrough.SolOut"
'SolCallback(30)="SolGameOver"

Ballsize = 50
BallMass = 1

'***************************
'        TABLE INIT
'***************************

Dim bsTrough,obj,DropTargetBank1,DropTargetBank2,cLeftCaptive,cRightCaptive,Flipperactive

Sub Nemesis_Init
  vpminit me

  dim xx
  For each xx in GIBulbs:xx.State = 1:Next
  For each xx in GIPlastics:xx.State = 1:Next

  Controller.GameName=cGameName
  Controller.SplashInfoLine="Nemesis" & vbNewLine & "converted from VP9 to VPX by Batch"
  Controller.HandleKeyboard=False
  Controller.ShowTitle=0
  Controller.ShowFrame=0
  Controller.ShowDMDOnly=1
  Controller.Hidden = 0     'enable to hide DMD if You use a B2S backglass

  '    'DMD position for 3 Monitor Setup
  '    'Controller.Games(cGameName).Settings.Value("dmd_pos_x")=3850    'set this to 0 if You cannot find the DMD
  '    'Controller.Games(cGameName).Settings.Value("dmd_pos_y")=300   'set this to 0 if You cannot find the DMD
  '    'Controller.Games(cGameName).Settings.Value("dmd_width")=505
  '    'Controller.Games(cGameName).Settings.Value("dmd_height")=155
  '    'Controller.Games(cGameName).Settings.Value("rol")=0
  '
  ' 'Controller.Games(cGameName).Settings.Value("ddraw") = 0             'set to 0 if You have problems with DMD showing or table stutter

  Controller.HandleMechanics=0
  Controller.Run
  If Err Then MsgBox Err.Description
    On Error Goto 0

  PinMAMETimer.Interval=PinMAMEInterval
  PinMAMETimer.Enabled = true

  vpmNudge.TiltSwitch=-5
  vpmNudge.Sensitivity=5
  vpmNudge.TiltObj = Array(LeftFlipper,RightFlipper)

  vpmMapLights AllLights

  Set bsTrough=New cvpmBallStack
  bsTrough.InitSw 0,0.1,0,0,0,0,0,0     '0.1 = Switch 0
  bsTrough.InitKick BallRelease,90,5
  bsTrough.InitExitSnd SoundFX("BallRel",DOFContactors),SoundFX("Solenoid",DOFContactors)
  bsTrough.Balls=1

  set DropTargetBank1 = new cvpmDropTarget
  DropTargetBank1.InitDrop Array(DT1,DT2), Array(18,19)
  DropTargetBank1.InitSnd SoundFX("fx_DropTarget",DOFDropTargets),SoundFX("TargetBankreset1",DOFContactors)
  DropTargetBank1.CreateEvents "DropTargetBank1"

  set DropTargetBank2 = new cvpmDropTarget
  DropTargetBank2.InitDrop Array(DT3,DT4,DT5), Array(22,23,24)
  DropTargetBank2.InitSnd SoundFX("fx_DropTarget",DOFDropTargets),SoundFX("TargetBankreset1",DOFContactors)
  DropTargetBank2.CreateEvents "DropTargetBank2"

  Set cLeftCaptive=New cvpmCaptiveBall
  cLeftCaptive.InitCaptive LeftCaptiveTrigger,LeftCaptiveWall,LeftCaptiveKicker,-22
  cLeftCaptive.Start
  cLeftCaptive.ForceTrans = 0.5
  cLeftCaptive.MinForce = 3.5'
  cLeftCaptive.CreateEvents "cLeftCaptive"

  Set cRightCaptive=New cvpmCaptiveBall
  cRightCaptive.InitCaptive RightCaptiveTrigger,RightCaptiveWall,RightCaptiveKicker,-22
  cRightCaptive.Start
  cRightCaptive.ForceTrans = 0.5
  cRightCaptive.MinForce = 3.5
  cRightCaptive.CreateEvents "cRightCaptive"

End Sub

'******************************
'        TROUGH HANDLER
'******************************

'Sub Drain_Hit()
' SoundTimer.enabled = False
' bsTrough.AddBall Me
' playsound "Drain5"
'End Sub

Sub Drain_Hit()
  SoundTimer.enabled = False
  vpmTimer.addTimer 2200, "bsTrough.addball Drain '"
  playsoundAtVol "Drain5", Drain, 1
End Sub

'*******************************
'       KEYBOARD HANDLER
'*******************************

Sub Nemesis_KeyDown(ByVal keycode)
  if keycode = startgamekey then
    if FreePlay then
      vpmTimer.PulseSw -3
    end if
    Controller.Switch(-4) = 1
  end if

  If keycode = PlungerKey Then
    Plunger.PullBack
  End If

  If keycode = LeftFlipperKey Then
    if Flipperactive then
      LF.fire'LeftFlipper.RotatetoEnd
      controller.switch(103) = 1
      PlaySoundAtVol SoundFX("FlipperUp",DOFFlippers), LeftFlipper, 1
    end if
  End If

  If keycode = RightFlipperKey Then
    if Flipperactive then
      RF.fire'RightFlipper.RotatetoEnd
      controller.switch(101) = 1
      PlaySoundAtVol SoundFX("FlipperUp",DOFFlippers), RightFlipper, 1
    end if

  If keycode = LeftFlipperKey Then FlipperActivate LeftFlipper, LFPress
    If keycode = RightFlipperKey Then FlipperActivate RightFlipper, RFPress
    End If

  if keycode = LeftMagnaSave then
    vpmNudge.DoNudge 45, 2
  end if
  if keycode = RightMagnaSave then
    vpmNudge.DoNudge -45, 2
  end if

  If vpmKeyDown(KeyCode) Then Exit Sub
End Sub

Sub Nemesis_KeyUp(ByVal keycode)
  if keycode = startgamekey then
    Controller.Switch(-4) = 0
  end if

  If keycode = PlungerKey Then
    Plunger.Fire
    PlaySoundAtVol "fx_plunger", Plunger, 1
  End If

  If keycode = LeftFlipperKey Then
    controller.switch(103) = 0
    LeftFlipper.RotatetoStart
    if Flipperactive then
      PlaySoundAtVol SoundFX("FlipperDown",DOFFlippers), LeftFlipper, 1
    end if
  End If

  If keycode = RightFlipperKey Then
    controller.switch(101) = 0
    RightFlipper.RotatetoStart
    if Flipperactive then
      PlaySoundAtVol SoundFX("FlipperDown",DOFFlippers), RightFlipper, 1
    end if

  If keycode = LeftFlipperKey Then FlipperDeActivate LeftFlipper, LFPress
    If keycode = RightFlipperKey Then FlipperDeActivate RightFlipper, RFPress
    End If

  If vpmKeyUp(KeyCode) Then Exit Sub
End Sub

'*********************************************************************
'   FLIPPER CORRECTION INITIALIZATION (NFOZZY / ROTHBAUERW)
'*********************************************************************

dim LF : Set LF = New FlipperPolarity
dim RF : Set RF = New FlipperPolarity

InitPolarity

Sub InitPolarity()
  dim x, a : a = Array(LF, RF)
  for each x in a
    x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
    x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1
    x.enabled = True
    x.TimeDelay = 60
  Next

  AddPt "Polarity", 0, 0, 0
  AddPt "Polarity", 1, 0.05, -5
  AddPt "Polarity", 2, 0.4, -5
  AddPt "Polarity", 3, 0.6, -4.5
  AddPt "Polarity", 4, 0.65, -4.0
  AddPt "Polarity", 5, 0.7, -3.5
  AddPt "Polarity", 6, 0.75, -3.0
  AddPt "Polarity", 7, 0.8, -2.5
  AddPt "Polarity", 8, 0.85, -2.0
  AddPt "Polarity", 9, 0.9,-1.5
  AddPt "Polarity", 10, 0.95, -1.0
  AddPt "Polarity", 11, 1, -0.5
  AddPt "Polarity", 12, 1.1, 0
  AddPt "Polarity", 13, 1.3, 0

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
'     FLIPPER CORRECTION FUNCTIONS (NFOZZY / ROTHBAUERW)
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
'     FLIPPER POLARITY AND RUBBER DAMPENER
'    SUPPORTING FUNCTIONS (NFOZZY / ROTHBAUERW)
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

'******************************************************
'     FLIPPER TRICKS (NFOZZY / ROTHBAUERW)
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
  RandomSoundRubberFlipper(parm)
  CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
  'LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
  RandomSoundRubberFlipper(parm)
  CheckLiveCatch Activeball, RightFlipper, RFCount, parm
  'RightFlipperCollide parm
End Sub

Sub RandomSoundRubberFlipper(parm)
  Select Case Int(Rnd*7)+1
    Case 1 : PlaySoundAtBallVol "TOM_Rubber_Flipper_Normal_1", parm
    Case 2 : PlaySoundAtBallVol "TOM_Rubber_Flipper_Normal_2", parm
    Case 3 : PlaySoundAtBallVol "TOM_Rubber_Flipper_Normal_3", parm
    Case 4 : PlaySoundAtBallVol "TOM_Rubber_Flipper_Normal_4", parm
    Case 5 : PlaySoundAtBallVol "TOM_Rubber_Flipper_Normal_5", parm
    Case 6 : PlaySoundAtBallVol "TOM_Rubber_Flipper_Normal_6", parm
    Case 7 : PlaySoundAtBallVol "TOM_Rubber_Flipper_Normal_7", parm
  End Select
End Sub

'******************************
'        SWITCH HANDLER
'******************************
' Cabinet switches
'Const swStartButton    = -4
'Const swCoin1          = -3
'Const swCoin2          = -1
'Const swCoin3          = -2
'Const swTilt           = -5
'Const swSlamDoorHit    = -6
'Const swLRFlip         = 101
'Const swLLFlip         = 103

'**********************************
'      SLINGSHOT ANIMATIONS
'**********************************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
  vpmtimer.pulsesw 5
  PlaySoundAtVol SoundFX("LSling",DOFContactors), Sling1, 1
  RSling.Visible = 0
  RSling1.Visible = 1
  SLING1.TransZ = -20
  RStep = 0
  RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
  Select Case RStep
    Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:SLING1.TransZ = -10
    Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:SLING1.TransZ = 0:RightSlingShot.TimerEnabled = 0:
  End Select
  RStep = RStep + 1
End Sub


Sub LeftSlingShot_Slingshot
  vpmtimer.pulsesw 3
  PlaySoundAtVol SoundFX("RSling",DOFContactors), Sling2, 1
  LSling.Visible = 0
  LSling1.Visible = 1
  SLING2.TransZ = -20
  LStep = 0
  LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
  Select Case LStep
    Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:SLING2.TransZ = -10
    Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:SLING2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
  End Select
  LStep = LStep + 1
End Sub

'****************************
'         ROLLOVERS
'****************************

'Top Left Rollover
sub sw28_hit:Controller.Switch(28)=1:PlaySoundAtVol "fx_Rollover", ActiveBall, 1:End Sub
sub sw28_unhit:Controller.Switch(28)=0:End Sub

'500/500 Rollovers
sub sw10_hit:Controller.Switch(8)=1: PlaySoundAtVol "fx_Rollover", ActiveBall, 1:End Sub
sub sw10_unhit:Controller.Switch(8)=0:End Sub
sub sw11_hit:Controller.Switch(9)=1: PlaySoundAtVol "fx_Rollover", ActiveBall, 1:End Sub
sub sw11_unhit:Controller.Switch(9)=0 :End Sub
sub sw12_hit:Controller.Switch(12)=1: PlaySoundAtVol "fx_Rollover", ActiveBall, 1:End Sub
sub sw12_unhit:Controller.Switch(12)=0:End Sub
sub sw13_hit:Controller.Switch(13)=1 :PlaySoundAtVol "fx_Rollover", ActiveBall, 1:End Sub
sub sw13_unhit:Controller.Switch(13)=0 :End Sub
sub sw14_hit:Controller.Switch(14)=1 :PlaySoundAtVol "fx_Rollover", ActiveBall, 1:End Sub
sub sw14_unhit:Controller.Switch(14)=0 :End Sub
sub sw15_hit:Controller.Switch(15)=1 :PlaySoundAtVol "fx_Rollover", ActiveBall, 1:End Sub
sub sw15_unhit:Controller.Switch(15)=0 :End Sub
sub sw16_hit:Controller.Switch(16)=1 :PlaySoundAtVol "fx_Rollover", ActiveBall, 1:End Sub
sub sw16_unhit:Controller.Switch(16)=0 :End Sub
sub sw17_hit:Controller.Switch(17)=1 :PlaySoundAtVol "fx_Rollover", ActiveBall, 1:End Sub
sub sw17_unhit:Controller.Switch(17)=0 :End Sub

'Outlanes
sub sw34_hit:Controller.Switch(34)=1 :PlaySoundAtVol "fx_Rollover", ActiveBall, 1:End Sub
sub sw34_unhit:Controller.Switch(34)=0 :End Sub
sub sw35_hit:Controller.Switch(35)=1 :PlaySoundAtVol "fx_Rollover", ActiveBall, 1:End Sub
sub sw35_unhit:Controller.Switch(35)=0 :End Sub

'Captive Balls and Target behind DT
sub sw27_hit:Controller.Switch(27)=1
  PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
sub sw27_unhit:Controller.Switch(27)=0:End Sub

'*********************************
'          ROUND TARGETS
'*********************************

Sub Target1_Hit:vpmTimer.PulseSw 25:PlaySoundAtVol SoundFX("fx_RoundTarget",DOFTargets), ActiveBall, 1:End Sub  'captive 1
Sub Target4_Hit:vpmTimer.PulseSw 26:PlaySoundAtVol SoundFX("fx_RoundTarget",DOFTargets), ActiveBall, 1:End Sub   'captive 2

Sub Target2_Hit:vpmTimer.PulseSw 31:PlaySoundAtVol SoundFX("fx_RoundTarget",DOFTargets), ActiveBall, 1:End Sub   'Top 1Mio
Sub Target3_Hit:vpmTimer.PulseSw 29:PlaySoundAtVol SoundFX("fx_RoundTarget",DOFTargets), ActiveBall, 1:End Sub   'Center

'****************************
'          RUBBERS
'****************************

Sub Rubber1_Hit:vpmTimer.PulseSw 36:PlaySoundAtVol "fx_Rubber", ActiveBall, 1:End Sub
Sub Rubber2_Hit:vpmTimer.PulseSw 36:PlaySoundAtVol "fx_Rubber", ActiveBall, 1:End Sub
Sub Rubber3_Hit:vpmTimer.PulseSw 36:PlaySoundAtVol "fx_Rubber", ActiveBall, 1:End Sub
Sub Rubber4_Hit:vpmTimer.PulseSw 36:PlaySoundAtVol "fx_Rubber", ActiveBall, 1:End Sub

Sub Rubber11_Hit :PlaySoundAtVol "fx_Rubber", ActiveBall, 1:End Sub
Sub Rubber12_Hit :PlaySoundAtVol "fx_Rubber", ActiveBall, 1:End Sub
Sub Rubber13_Hit :PlaySoundAtVol "fx_Rubber", ActiveBall, 1:End Sub
Sub Rubber14_Hit :PlaySoundAtVol "fx_Rubber", ActiveBall, 1:End Sub
Sub Rubber15_Hit :PlaySoundAtVol "fx_Rubber", ActiveBall, 1:End Sub
Sub Rubber16_Hit :PlaySoundAtVol "fx_Rubber", ActiveBall, 1:End Sub
Sub Rubber17_Hit :PlaySoundAtVol "fx_Rubber", ActiveBall, 1:End Sub
Sub Rubber18_Hit :PlaySoundAtVol "fx_Rubber", ActiveBall, 1:End Sub
Sub Rubber19_Hit :PlaySoundAtVol "fx_Rubber", ActiveBall, 1:End Sub
Sub Rubber20_Hit :PlaySoundAtVol "fx_Rubber", ActiveBall, 1:End Sub
Sub Rubber21_Hit :PlaySoundAtVol "fx_Rubber", ActiveBall, 1:End Sub
Sub Rubber22_Hit :PlaySoundAtVol "fx_Rubber", ActiveBall, 1:End Sub
Sub Rubber23_Hit :PlaySoundAtVol "fx_Rubber", ActiveBall, 1:End Sub
Sub Rubber24_Hit :PlaySoundAtVol "fx_Rubber", ActiveBall, 1:End Sub
Sub Rubber25_Hit :PlaySoundAtVol "fx_Rubber", ActiveBall, 1:End Sub

'****************************************************************************
'         PHYSICS DAMPENERS - RUBBER FUNCTIONS (NFOZZY / ROTHBAUERW)
'****************************************************************************

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

'**************************************************************************************************
'   TRACK ALL BALL VELOCITIES FOR RUBBER DAMPENER AND DROP TARGETS (NFOZZY / ROTHBAUERW)
'**************************************************************************************************

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

'****************************
'        DROP TARGETS
'****************************

Sub DT1_Dropped:DropTargetBank1.Hit 1:End Sub
Sub DT2_Dropped:DropTargetBank1.Hit 2:End Sub
Sub DT3_Dropped:DropTargetBank2.Hit 1:End Sub
Sub DT4_Dropped:DropTargetBank2.Hit 2:End Sub
Sub DT5_Dropped:DropTargetBank2.Hit 3:End Sub

'*****************************
'          WIRE RAMPS
'*****************************

Sub WireRamp9_Hit :PlaySoundAtVol "WireRamp", WireRamp9, 1:End Sub

'*****************************
'         RUBBER POSTS
'*****************************

Sub RubberPost1_Hit :PlaySoundAtVol "fx_rubber_post", ActiveBall, 1:End Sub
Sub RubberPost2_Hit :PlaySoundAtVol "fx_rubber_post", ActiveBall, 1:End Sub
Sub RubberPost3_Hit :PlaySoundAtVol "fx_rubber_post", ActiveBall, 1:End Sub
Sub RubberPost4_Hit :PlaySoundAtVol "fx_rubber_post", ActiveBall, 1:End Sub
Sub RubberPost5_Hit :PlaySoundAtVol "fx_rubber_post", ActiveBall, 1:End Sub
Sub RubberPost6_Hit :PlaySoundAtVol "fx_rubber_post", ActiveBall, 1:End Sub
Sub RubberPost7_Hit :PlaySoundAtVol "fx_rubber_post", ActiveBall, 1:End Sub
Sub RubberPost8_Hit :PlaySoundAtVol "fx_rubber_post", ActiveBall, 1:End Sub

Sub RubberPost13_Hit :PlaySoundAtVol "fx_rubber_post", ActiveBall, 1:End Sub
Sub RubberPost14_Hit :PlaySoundAtVol "fx_rubber_post", ActiveBall, 1:End Sub
Sub RubberPost15_Hit :PlaySoundAtVol "fx_rubber_post", ActiveBall, 1:End Sub
Sub RubberPost16_Hit :PlaySoundAtVol "fx_rubber_post", ActiveBall, 1:End Sub
Sub RubberPost17_Hit :PlaySoundAtVol "fx_rubber_post", ActiveBall, 1:End Sub
Sub RubberPost18_Hit :PlaySoundAtVol "fx_rubber_post", ActiveBall, 1:End Sub
Sub RubberPost19_Hit :PlaySoundAtVol "fx_rubber_post", ActiveBall, 1:End Sub
Sub RubberPost20_Hit :PlaySoundAtVol "fx_rubber_post", ActiveBall, 1:End Sub
Sub RubberPost21_Hit :PlaySoundAtVol "fx_rubber_post", ActiveBall, 1:End Sub
Sub RubberPost22_Hit :PlaySoundAtVol "fx_rubber_post", ActiveBall, 1:End Sub
Sub RubberPost23_Hit :PlaySoundAtVol "fx_rubber_post", ActiveBall, 1:End Sub
Sub RubberPost24_Hit :PlaySoundAtVol "fx_rubber_post", ActiveBall, 1:End Sub
Sub RubberPost25_Hit :PlaySoundAtVol "fx_rubber_post", ActiveBall, 1:End Sub
Sub RubberPost26_Hit :PlaySoundAtVol "fx_rubber_post", ActiveBall, 1:End Sub
Sub RubberPost27_Hit :PlaySoundAtVol "fx_rubber_post", ActiveBall, 1:End Sub

Sub RubberPost29_Hit :PlaySoundAtVol "fx_rubber_post", ActiveBall, 1:End Sub
Sub RubberPost30_Hit :PlaySoundAtVol "fx_rubber_post", ActiveBall, 1:End Sub
Sub RubberPost31_Hit :PlaySoundAtVol "fx_rubber_post", ActiveBall, 1:End Sub
Sub RubberPost32_Hit :PlaySoundAtVol "fx_rubber_post", ActiveBall, 1:End Sub
Sub RubberPost33_Hit :PlaySoundAtVol "fx_rubber_post", ActiveBall, 1:End Sub
Sub RubberPost34_Hit :PlaySoundAtVol "fx_rubber_post", ActiveBall, 1:End Sub
Sub RubberPost35_Hit :PlaySoundAtVol "fx_rubber_post", ActiveBall, 1:End Sub
Sub RubberPost36_Hit :PlaySoundAtVol "fx_rubber_post", ActiveBall, 1:End Sub
Sub RubberPost37_Hit :PlaySoundAtVol "fx_rubber_post", ActiveBall, 1:End Sub
Sub RubberPost38_Hit :PlaySoundAtVol "fx_rubber_post", ActiveBall, 1:End Sub
Sub RubberPost39_Hit :PlaySoundAtVol "fx_rubber_post", ActiveBall, 1:End Sub
Sub RubberPost40_Hit :PlaySoundAtVol "fx_rubber_post", ActiveBall, 1:End Sub
Sub RubberPost41_Hit :PlaySoundAtVol "fx_rubber_post", ActiveBall, 1:End Sub
Sub RubberPost42_Hit :PlaySoundAtVol "fx_rubber_post", ActiveBall, 1:End Sub
Sub RubberPost43_Hit :PlaySoundAtVol "fx_rubber_post", ActiveBall, 1:End Sub
Sub RubberPost44_Hit :PlaySoundAtVol "fx_rubber_post", ActiveBall, 1:End Sub
Sub RubberPost45_Hit :PlaySoundAtVol "fx_rubber_post", ActiveBall, 1:End Sub
Sub RubberPost46_Hit :PlaySoundAtVol "fx_rubber_post", ActiveBall, 1:End Sub
Sub RubberPost47_Hit :PlaySoundAtVol "fx_rubber_post", ActiveBall, 1:End Sub
Sub RubberPost48_Hit :PlaySoundAtVol "fx_rubber_post", ActiveBall, 1:End Sub
Sub RubberPost49_Hit :PlaySoundAtVol "fx_rubber_post", ActiveBall, 1:End Sub
Sub RubberPost50_Hit :PlaySoundAtVol "fx_rubber_post", ActiveBall, 1:End Sub
Sub RubberPost51_Hit :PlaySoundAtVol "fx_rubber_post", ActiveBall, 1:End Sub
Sub RubberPost52_Hit :PlaySoundAtVol "fx_rubber_post", ActiveBall, 1:End Sub
Sub RubberPost53_Hit :PlaySoundAtVol "fx_rubber_post", ActiveBall, 1:End Sub

'*****************************
'        RUBBER SLEEVES
'*****************************

Sub RubberSleeve2_Hit :PlaySoundAtVol "fx_metalwire", ActiveBall, 1:End Sub
Sub RubberSleeve4_Hit :PlaySoundAtVol "fx_metalwire", ActiveBall, 1:End Sub
Sub RubberSleeve6_Hit :PlaySoundAtVol "fx_metalwire", ActiveBall, 1:End Sub
Sub RubberSleeve8_Hit :PlaySoundAtVol "fx_metalwire", ActiveBall, 1:End Sub
Sub RubberSleeve9_Hit :PlaySoundAtVol "fx_metalwire", ActiveBall, 1:End Sub
Sub RubberSleeve10_Hit :PlaySoundAtVol "fx_metalwire", ActiveBall, 1:End Sub



'***********************************
'          HELPER FUNCTIONS
'***********************************

Sub LeftCaptiveHelper_Hit
  if Activeball.vely < 0 then
    RightCaptiveKicker.kick -22, abs(ActiveBall.vely) * 0.3
  end if
End Sub

Sub RightCaptiveHelper_Hit
  if Activeball.vely < 0 then
    LeftCaptiveKicker.kick -22, abs(ActiveBall.vely) * 0.3
  end if
End Sub

'****************************
'           GATES
'****************************

Sub Gate_Hit:PlaySoundAtVol "Gate5", Gate, 1:End Sub
Sub Gate1_Hit:PlaySoundAtVol "Gate5", Gate1, 1:End Sub

'****************************
'          BUMPERS
'****************************

Dim Bumper1_Dir,Bumper2_Dir

Sub Bumper1_Hit : vpmTimer.PulseSw(1) : playsoundAtVol SoundFX("bumper2",DOFContactors), ActiveBall, 1: End Sub
Sub Bumper2_Hit : vpmTimer.PulseSw(2) : playsoundAtVol SoundFX("bumper2",DOFContactors), ActiveBall, 1: End Sub

'BUMPER RINGS
Sub Bumper1_Timer
  P_B1Ring.TransZ = P_B1Ring.TransZ + Bumper1_Dir * 2.5
  if P_B1Ring.TransZ <= -30 then Bumper1_Dir = 1
  if P_B1Ring.TransZ >= 0 then
    Bumper1.Timerenabled = False
    Bumper1_Dir = 0
  end if
End Sub
Sub Bumper2_Timer
  P_B2Ring.TransZ = P_B2Ring.TransZ + Bumper2_Dir * 2.5
  if P_B2Ring.TransZ <= -30 then Bumper2_Dir = 1
  if P_B2Ring.TransZ >= 0 then
    Bumper2.Timerenabled = False
    Bumper2_Dir = 0
  end if
End Sub

'********************************************
'         ADDITIONAL LAMPS CALLBACK
'********************************************

Dim LeftCaptiveActive,RightCaptiveActive

Sub CBTimer_Timer
  if Controller.lamp(33) or Controller.lamp(40) then
    LeftCaptiveActive = True
    RightCaptiveActive = False
  end if
  if Controller.lamp(34) or Controller.lamp(39) then
    LeftCaptiveActive = False
    RightCaptiveActive = True
  end if
  if Controller.lamp(48) then
    Lamp48.state = 1
  else
    Lamp48.state = 0
  end if
  Flipperactive = Controller.lamp(54)
End sub

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
        If BOT(b).X < Nemesis.Width/2 Then
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Nemesis.Width/2))/21)) + 6
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Nemesis.Width/2))/21)) - 6
        End If
        ballShadow(b).Y = BOT(b).Y + 4
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
End Sub

'***********************************
'       DIP SWITCH SETTINGS
'***********************************

'B        A
'87654321 87654321

Sub editDips
  Dim vpmDips:Set vpmDips=New cvpmDips
  With vpmDips
    .AddForm 700,280,"Nemesis - DIP switches"

    .AddFrame 0,0,110,"Bonus Feature",&H00000100,Array("Adaptive",0,"Fixed",&H00000100)'dip B-1
    .AddFrame 0,45,110,"Match Feature",&H00000002,Array("Credit",0,"Double Score",&H00000002)'dip A-2
    .AddChk 0,95,150,Array("Attract Sound",&H00000004)  'dip A-3

'   .AddChk 0,115,150,Array("Enable Features below",&H00001000) 'dip B-5
'   .AddChk 15,130,150,Array("Show game-on time",&H00002000)  'dip B-6
'   .AddChk 15,145,150,Array("Show inserted coins",&H00004000)  'dip B-7
'   .AddChk 15,160,150,Array("Show played games",&H00008000)  'dip B-8
'
'   .AddChk 0,180,150,Array("Erase memory A",&H00000400)    'dip B-2
'   .AddChk 0,195,150,Array("Erase memory B",&H00000200)    'dip B-3
    .AddFrame 160,0,110,"Coins per game",&H00000018,Array("1-3",0,"1-5",&H00000018,"1-6",&H00000008,"2-6",&H00000010)'dip A-4&A-5
    .AddFrame 160,75,110,"Score threshold",&H000000C0,Array("1,400,000 points",&H000000C0,"1,600,000 points",&H00000080,"1,500,000 points",&H00000040,"1,700,000 points",0)'dip A-7&A-8
    .AddFrame 160,150,110,"Balls per game",&H00000020,Array("5 balls",0,"3 balls",&H00000020)'dip A-6
    .AddLabel 0,220,280,20,"After hitting OK, press F3 to reset game with new settings."
    .ViewDips
  End With
End Sub
Set vpmShowDips=GetRef("editDips")

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
'                     Supporting Ball & Sound Functions
'*********************************************************************

Function RndNum(min, max)
  RndNum = Int(Rnd() * (max-min + 1) ) + min ' Sets a random number between min and max
End Function

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "Nemesis" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = tableobj.y * 2 / Nemesis.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "Nemesis" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = tableobj.x * 2 / Nemesis.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "Nemesis" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = ball.x * 2 / Nemesis.width-1
  If tmp > 0 Then
    Pan = Csng(tmp ^10)
  Else
    Pan = Csng(-((- tmp) ^10) )
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


'**********************************************
'        JP SALAS' VP10 ROLLING SOUNDS
'**********************************************

Const tnob = 3 ' total number of balls
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
    If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
      PlaySoundAtBOTBallZ "fx_ball_drop" & b, BOT(b)
      'debug.print BOT(b).velz
    End If
    Next
End Sub

'****************************
'    BALL COLISION SOUND
'****************************

Sub OnBallBallCollision(ball1, ball2, velocity)
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / (VolDiv/VolCol), Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

Sub SolKnocker(Enabled)
    If Enabled Then PlaySoundAt SoundFX("knocker",DOFKnocker), Plunger
End Sub

'**********************************
'      DESTRUK'S DISPLAY CODE
'**********************************

Dim Digits(53)
Digits(0)=Array(Light1,Light2,Light3,Light4,Light5,Light6,Light7,Light8)'ok
Digits(1)=Array(Light15,Light9,Light10,Light11,Light12,Light13,Light14)'ok
Digits(2)=Array(Light23,Light16,Light18,Light19,Light20,Light21,Light22)'ok
Digits(3)=Array(Light30,Light24,Light25,Light26,Light27,Light28,Light29,Light31)'ok
Digits(4)=Array(Light38,Light32,Light33,Light34,Light35,Light36,Light37)'ok
Digits(5)=Array(Light45,Light39,Light40,Light41,Light42,Light43,Light44)'ok
Digits(6)=Array(Light53,Light46,Light47,Light49,Light50,Light51,Light52)'ok

Digits(7)=Array(Light54,Light55,Light57,Light59,Light60,Light61,Light62,Light63)'ok
Digits(8)=Array(Light70,Light64,Light65,Light66,Light67,Light68,Light69)'ok
Digits(9)=Array(Light77,Light71,Light72,Light73,Light74,Light75,Light76)'ok
Digits(10)=Array(Light84,Light78,Light79,Light80,Light81,Light82,Light83,Light85)'ok
Digits(11)=Array(Light92,Light86,Light87,Light88,Light89,Light90,Light91)'ok
Digits(12)=Array(Light99,Light93,Light94,Light95,Light96,Light97,Light98)'ok
Digits(13)=Array(Light106,Light100,Light101,Light102,Light103,Light104,Light105)'ok

Digits(14)=Array(Light107,Light108,Light109,Light110,Light111,Light112,Light113,Light114)'ok
Digits(15)=Array(Light121,Light115,Light116,Light117,Light118,Light119,Light120)'ok
Digits(16)=Array(Light128,Light122,Light123,Light124,Light125,Light126,Light127)'ok
Digits(17)=Array(Light135,Light129,Light130,Light131,Light132,Light133,Light134,Light136)'ok
Digits(18)=Array(Light143,Light137,Light138,Light139,Light140,Light141,Light142)'ok
Digits(19)=Array(Light150,Light144,Light145,Light146,Light147,Light148,Light149)'ok
Digits(20)=Array(Light157,Light151,Light152,Light153,Light154,Light155,Light156)'ok

Digits(21)=Array(Light158,Light159,Light160,Light161,Light162,Light163,Light164,Light165)'ok
Digits(22)=Array(Light172,Light166,Light167,Light168,Light169,Light170,Light171)'ok
Digits(23)=Array(Light179,Light173,Light174,Light175,Light176,Light177,Light178)'ok
Digits(24)=Array(Light186,Light180,Light181,Light182,Light183,Light184,Light185,Light187)'ok
Digits(25)=Array(Light194,Light188,Light189,Light190,Light191,Light192,Light193)'ok
Digits(26)=Array(Light201,Light195,Light196,Light197,Light198,Light199,Light200)'ok
Digits(27)=Array(Light208,Light202,Light203,Light204,Light205,Light206,Light207)'ok

Digits(28)=Array(Light209,Light210,Light211,Light212,Light213,Light214,Light215)'ok
Digits(29)=Array(Light216,Light217,Light218,Light219,Light220,Light221,Light222)'ok

Digits(30)=Array(Light223,Light224,Light225,Light226,Light227,Light228,Light229)'ok
Digits(31)=Array(Light230,Light231,Light232,Light233,Light234,Light235,Light236)'ok

Digits(32)=Array(Light237,Light238,Light239,Light240,Light241,Light242,Light243)'ok

Digits(33)=Array(Light17,Light48,Light56,Light58,Light244,Light245,Light246,Light247)'ok
Digits(34)=Array(Light254,Light248,Light249,Light250,Light251,Light252,Light253)'ok
Digits(35)=Array(Light261,Light255,Light256,Light257,Light258,Light259,Light260)'ok
Digits(36)=Array(Light268,Light262,Light263,Light264,Light265,Light266,Light267,Light269)'ok
Digits(37)=Array(Light276,Light270,Light271,Light272,Light273,Light274,Light275)'ok
Digits(38)=Array(Light283,Light277,Light278,Light279,Light280,Light281,Light282)'ok
Digits(39)=Array(Light290,Light284,Light285,Light286,Light287,Light288,Light289)'ok

Digits(40)=Array(Light291,Light292,Light293,Light294,Light295,Light296,Light297,Light298)'ok
Digits(41)=Array(Light305,Light299,Light300,Light301,Light302,Light303,Light304)'ok
Digits(42)=Array(Light312,Light306,Light307,Light308,Light309,Light310,Light311)'ok
Digits(43)=Array(Light319,Light313,Light314,Light315,Light316,Light317,Light318,Light320)'ok
Digits(44)=Array(Light327,Light321,Light322,Light323,Light324,Light325,Light326)'ok
Digits(45)=Array(Light334,Light328,Light329,Light330,Light331,Light332,Light333)'ok
Digits(46)=Array(Light341,Light335,Light336,Light337,Light338,Light339,Light340)'ok

Digits(47)=Array(Light342,Light343,Light344,Light345,Light346,Light347,Light348,Light349)'ok
Digits(48)=Array(Light356,Light350,Light351,Light352,Light353,Light354,Light355)'ok
Digits(49)=Array(Light363,Light357,Light358,Light359,Light360,Light361,Light362)'ok
Digits(50)=Array(Light370,Light364,Light365,Light366,Light367,Light368,Light369,Light371)'ok
Digits(51)=Array(Light378,Light372,Light373,Light374,Light375,Light376,Light377)'ok
Digits(52)=Array(Light385,Light379,Light380,Light381,Light382,Light383,Light384)'ok
Digits(53)=Array(Light392,Light386,Light387,Light388,Light389,Light390,Light391)'ok


Sub DisplayTimer_Timer
  Dim ChgLED,ii,num,chg,stat,obj
  ChgLED = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
  If Not IsEmpty(ChgLED) Then
    For ii = 0 To UBound(chgLED)
      num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
        For Each obj In Digits(num)
          If chg And 1 Then obj.State = stat And 1
          chg = chg\2 : stat = stat\2
        Next
    Next
  End If
End Sub

Sub Nemesis_Exit()
  If B2SOn Then
    Controller.Pause = False
    Controller.Stop
  End If
End Sub



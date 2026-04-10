Option Explicit
Randomize
setlocale(1033)


On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

'Pinball Stern 1977 V1.0
' Thalamus 2019 July : Improved directional sounds
' !! NOTE : Table not verified yet !!

' Options
' Volume devided by - lower gets higher sound

Const VolDiv = 3000    ' Lower number, louder ballrolling/collition sound
Const VolCol = 10      ' Ball collition divider ( voldiv/volcol )



'******Rom Choice*******

'***Scoring and lights are correct but no 3rd and 4th player***
Const cGameName="pinball",UseSolenoids=2,UseLamps=0,UseGI=0,SCoin="coin"

'***adds 3rd and 4th player but lights are different***
'Const cGameName="stingray",UseSolenoids=2,UseLamps=0,UseGI=0,SCoin="coin"

'***Free play rom***
'Const cGameName="pinbalfp",UseSolenoids=2,UseLamps=0,UseGI=0,SCoin="coin"



LoadVPM "01000100", "Bally.VBS", 1.2
Dim DesktopMode: DesktopMode = Table1.ShowDT


If DesktopMode = True Then 'Show Desktop components
l45.x=874.1237
l15.x=58.12378
L31.x=870.1237
l47.x=64.12378
l63.x=872.1237
Ramp16.visible=1
Ramp15.visible=1
Primitive13.visible=1
Else
l45.x=1083.458
l15.x=1083.458
L31.x=1083.458
l47.x=1083.458
l63.x=1083.458
Ramp16.visible=0
Ramp15.visible=0
Primitive13.visible=0
End if

'*************************************************************
'Solenoid Call backs
'**********************************************************************************************************
SolCallback(1)  = "vpmSolSound SoundFX(""10"",DOFChimes),"
SolCallback(2)  = "vpmSolSound SoundFX(""100"",DOFChimes),"
SolCallback(3)  = "vpmSolSound SoundFX(""1000"",DOFChimes),"
SolCallback(4)  = "vpmSolSound SoundFX(""10000"",DOFChimes),"
SolCallback(5)  = "SolDT" 'Drop Targets
SolCallback(6)  = "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(7)  = "bsTrough.SolOut"
SolCallback(9)  = "bsSaucer.SolOut" 'Top Kicker
SolCallback(8)  = "bsSaucer1.SolOut" 'Side Kicker
SolCallback(19) = "vpmNudge.SolGameOn"

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

'CHECKED
'*****************************************
'FLIPPER
'*****************************************

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFx("fx_Flipperup",DOFContactors),LeftFlipper, 1:LF.fire'LeftFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFx("fx_Flipperdown",DOFContactors),LeftFlipper,.1:LeftFlipper.RotateToStart
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFx("fx_Flipperup",DOFContactors),RightFlipper, 1:RF.fire'RightFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFx("fx_Flipperdown",DOFContactors),RightFlipper,.1:RightFlipper.RotateToStart
     End If
End Sub


'CHECKED
'******************************************************
'     STEPS 2-4 (FLIPPER POLARITY SETUP
'******************************************************

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
    x.TimeDelay = 44
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
Sub TriggerLF_UnHit() : LF.PolarityCorrect activeball : End Sub
Sub TriggerRF_Hit() : RF.Addball activeball : End Sub
Sub TriggerRF_UnHit() : RF.PolarityCorrect activeball : End Sub

'CHECKED
'**********************************************************************************************************
RightFlipper.timerinterval=1
rightflipper.timerenabled=True

sub RightFlipper_timer()

  If leftflipper.currentangle = leftflipper.endangle and LFPress = 1 then
    leftflipper.eostorqueangle = EOSAnew
    leftflipper.eostorque = EOSTnew
    LeftFlipper.rampup = EOSRampup
    if LFCount < LiveCatch Then
      LFCount = LFCount + 1
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
    if RFCount < LiveCatch Then
      RFCount = RFCount + 1
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

end sub

dim LFPress, RFPress, EOST, EOSA, EOSTnew, EOSAnew
dim FStrength, Frampup, FElasticity, EOSRampup, SOSRampup
dim RFEndAngle, LFEndAngle, LFCount, RFCount, LiveCatch

EOST = leftflipper.eostorque
EOSA = leftflipper.eostorqueangle
FStrength = LeftFlipper.strength
Frampup = LeftFlipper.rampup
FElasticity = LeftFlipper.elasticity
EOSTnew = 1.0 'FEOST
EOSAnew = 0.2
EOSRampup = 1.5
SOSRampup = 8.5
LiveCatch = 12

LFEndAngle = Leftflipper.endangle
RFEndAngle = RightFlipper.endangle

'CHECKED
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

'*****************************************
'FLIPPER SHADOWS
'*****************************************

sub FlipperTimer_Timer()
  FlipperLSh.RotZ = LeftFlipper.currentangle
  FlipperRSh.RotZ = RightFlipper.currentangle

End Sub



'Primitive Droptarget Reset
Sub SolDT(Enabled)
  If Enabled Then
    dtBank.DropSol_On 'Drop Target Wall reset
    PrimDropTgtUp dtBank, 1, 25, 0, 1
    PrimDropTgtUp dtBank, 2, 26, 0, 0
    PrimDropTgtUp dtBank, 3, 27, 0, 0
    PrimDropTgtUp dtBank, 4, 28, 0, 0
    PrimDropTgtUp dtBank, 5, 29, 0, 0
    debug.print "dtDrop"
  End If
End Sub

'**********************************************************************************************************

'Solenoid Controlled toys
'**********************************************************************************************************

'*****GI Lights On
dim xx
For each xx in GI:xx.State = 1: Next
'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************
Dim bsTrough, dtBank, bsSaucer, bsSaucer1, kickstep1

Sub Table1_Init
  vpmInit Me
  On Error Resume Next
    With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine = "Pinball (Stern 1977)"
    .HandleMechanics=0
    .HandleKeyboard=0
    .ShowDMDOnly=1
    .ShowFrame=0
    .ShowTitle=0
        .hidden = 1
    '.Games(cGameName).Settings.Value("sound")=1
    '.PuPHide = 1
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With
     On Error Goto 0

  PinMAMETimer.Interval = PinMAMEInterval
  PinMAMETimer.Enabled = True

  vpmNudge.TiltSwitch = 7
  vpmNudge.Sensitivity = 1
  vpmNudge.TiltObj = Array(Bumper1, Bumper2, LeftSlingshot, RightSlingshot)


  Set bsTrough = New cvpmBallStack ' Trough handler
    bsTrough.InitSw 0,8,0,0,0,0,0,0
    bsTrough.InitKick Ballrelease, 90, 5
    bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
    bsTrough.Balls = 1

  'Drop Targets
  Set dtBank = new cvpmDropTarget
    dtBank.InitDrop Array(Sw25, Sw26, Sw27, Sw28, Sw29), Array(25,26,27,28,29)
    dtBank.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

  Set bsSaucer=New cvpmBallStack
        bsSaucer.InitSaucer SW12,12,177,18
        bsSaucer.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
        bsSaucer.KickAngleVar=10
        bsSaucer.KickForceVar = 3

  Set bsSaucer1=New cvpmBallStack
        bsSaucer1.InitSaucer SW13,13,135,18
        bsSaucer1.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
        bsSaucer1.KickAngleVar=10
        bsSaucer1..KickForceVar = 3

End Sub


'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)
  If KeyDownHandler(keycode) Then Exit Sub
  If keycode = PlungerKey Then Plunger.Pullback:playsoundAtVol"plungerpull", Plunger, 1
  If keycode = LeftFlipperKey Then LFPress = 1
  If keycode = RightFlipperKey Then rfpress = 1
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
  If KeyUpHandler(keycode) Then Exit Sub
  If keycode = PlungerKey Then Plunger.Fire:PlaySoundAtVol"plunger", Plunger, 1
  If keycode = LeftFlipperKey Then
    lfpress = 0
    leftflipper.eostorqueangle = EOSA
    leftflipper.eostorque = EOST
  End If
  If keycode = RightFlipperKey Then
    rfpress = 0
    rightflipper.eostorqueangle = EOSA
    rightflipper.eostorque = EOST
  End If
End Sub

'**********************************************************************************************************

 ' Drain hole and kickers
Sub Drain_Hit:bsTrough.addball me : playsoundAtVol"drain" , Drain, 1: End Sub
Sub sw12_Hit:bsSaucer.AddBall 0 : playsoundAtVol "popper_ball", ActiveBall, 1: End Sub
Sub sw13_Hit:bsSaucer1.AddBall 0 : playsoundAtVol "popper_ball", ActiveBall, 1: End Sub

'Drop Targets
Sub Sw25_Hit: PrimDropTgtDown dtBank, 1, 25: End Sub
Sub Sw25_Timer: PrimDropTgtMove 25: End Sub
Sub Sw26_Hit: PrimDropTgtDown dtBank, 2, 26: End Sub
Sub Sw26_Timer: PrimDropTgtMove 26: End Sub
Sub Sw27_Hit: PrimDropTgtDown dtBank, 3, 27: End Sub
Sub Sw27_Timer: PrimDropTgtMove 27: End Sub
Sub Sw28_Hit: PrimDropTgtDown dtBank, 4, 28: End Sub
Sub Sw28_Timer: PrimDropTgtMove 28: End Sub
Sub Sw29_Hit: PrimDropTgtDown dtBank, 5, 29: End Sub
Sub Sw29_Timer: PrimDropTgtMove 29: End Sub

'Wire Triggers
Sub SW32_Hit:Controller.Switch(32)=1 : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub SW32_unHit:Controller.Switch(32)=0:End Sub
Sub SW24_Hit:Controller.Switch(24)=1 : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub SW24_unHit:Controller.Switch(24)=0:End Sub
Sub SW23_Hit:Controller.Switch(23)=1 : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub SW23_unHit:Controller.Switch(23)=0:End Sub
Sub SW31_Hit:Controller.Switch(31)=1 : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub SW31_unHit:Controller.Switch(31)=0:End Sub


'Bumpers
Sub Bumper1_Hit : vpmTimer.PulseSw(40) : playsoundAtVol SoundFX("fx_bumper1",DOFContactors), ActiveBall, 1: End Sub
Sub Bumper2_Hit : vpmTimer.PulseSw(39) : playsoundAtVol SoundFX("fx_bumper1",DOFContactors), ActiveBall, 1: End Sub
Sub Bumper3_Hit : vpmTimer.PulseSw(38) : playsoundAtVol SoundFX("fx_bumper1",DOFContactors), ActiveBall, 1: End Sub

'Spinners
Sub sw30_Spin:vpmTimer.PulseSw 30 : playsoundAtVol"fx_spinner" , sw30, 1: End Sub

Dim Step1, Step2', Step3, Step4, Step5

' 'Scoring Rubber
Sub sw21a_hit:vpmTimer.pulseSw 21 : playsoundAtVol"rubber_hit_2" , ActiveBall, 1: End Sub
Sub sw21b_hit:vpmTimer.pulseSw 21 : playsoundAtVol"rubber_hit_2" , ActiveBall, 1: End Sub
Sub sw21c_Hit:vpmTimer.PulseSw 21 : playsoundAtVol"rubber_hit_2" , ActiveBall, 1: Step2=1: sr2R.visible=0: sr2R1.visible=1: me.timerenabled=1: End Sub
Sub sw21d_hit:vpmTimer.pulseSw 21 : playsoundAtVol"rubber_hit_2" , ActiveBall, 1: End Sub
Sub sw22a_Hit:vpmTimer.PulseSw 22 : playsoundAtVol"rubber_hit_2" , ActiveBall, 1: Step1=1: sr1R.visible=0: sr1R1.visible=1: me.timerenabled=1: End Sub
Sub sw22b_hit:vpmTimer.pulseSw 22 : playsoundAtVol"rubber_hit_2" , ActiveBall, 1: End Sub


'Scoring rubbers animations

sub sw22a_timer
  select case Step1
    Case 1: sr1R1.visible=0: sr1r.visible=1
    Case 2: sr1R.visible=0: sr1R2.visible=1
    Case 3: sr1R2.visible=0: sr1R.visible=1: me.timerenabled=0
  end Select
  Step1=Step1+1
end sub

sub sw21c_timer
  select case Step2
    Case 1: sr2R1.visible=0: sr2r.visible=1
    Case 2: sr2R.visible=0: sr2R2.visible=1
    Case 3: sr2R2.visible=0: sr2R.visible=1: me.timerenabled=0
  end Select
  Step2=Step2+1
end sub


'Scoring Slingshots
Sub sw22a_Slingshot:vpmTimer.PulseSw 22:End Sub
Sub sw22b_Slingshot:vpmTimer.PulseSw 22:End Sub

Sub sw21a_Slingshot:vpmTimer.PulseSw 21:End Sub
Sub sw21b_Slingshot:vpmTimer.PulseSw 21:End Sub
Sub sw21c_Slingshot:vpmTimer.PulseSw 21:End Sub
Sub sw21d_Slingshot:vpmTimer.PulseSw 21:End Sub

'**********************************************************************************************************
'Rollover Targets
'**********************************************************************************************************

Sub sw14a_Hit
  sw14ap.z = -1.5
  Controller.Switch(14)=1 : playsoundAtVol"rollover", ActiveBall, 1
End Sub

Sub sw14a_UnHit
  sw14ap.z = .5
  Controller.Switch(14)=0
End Sub

Sub sw14b_Hit
  sw14bp.z = -1.5
  Controller.Switch(14)=1 : playsoundAtVol"rollover", ActiveBall, 1
End Sub

Sub sw14b_UnHit
  sw14bp.z = .5
  Controller.Switch(14)=0
End Sub

Sub sw14c_Hit
  sw14cp.z = -1.5
  Controller.Switch(14)=1 : playsoundAtVol"rollover", ActiveBall, 1
End Sub

Sub sw14c_UnHit
  sw14cp.z = .5
  Controller.Switch(14)=0
End Sub

Sub sw14d_Hit
  sw14dp.z = -1.5
  Controller.Switch(14)=1 : playsoundAtVol"rollover", ActiveBall, 1
End Sub

Sub sw14d_UnHit
  sw14dp.z = .5
  Controller.Switch(14)=0
End Sub

'****************************************************************************
'*****
'***** Primitive Animation subroutines          (gtxjoe v1.05)
'*****
'****************************************************************************
Const WallPrefix    = "Sw" 'Change this based on your naming convention
Const PrimitivePrefix   = "PrimSw"'Change this based on your naming convention
Const PrimitiveBumperPrefix = "BR" 'Change this based on your naming convention
Dim primCnt(100), primDir(100), primBmprDir(6)


'************************************************************************
'***** Primitive Drop Target Animation
'************************************************************************
'USAGE:  Sub Sw13_Hit: PrimDropTgtDown RBank, 3, 13: End Sub
'USAGE:  Sub Sw13_Timer: PrimDropTgtMove 13: End Sub
'USAGE:  Sub solRBankReset (enabled): If enabled Then PrimDropTgtUp RBank, 1, 13, 0, 1: PrimDropTgtUp RBank, 2, 12, 0, 0: PrimDropTgtUp RBank, 3, 11, 0, 0: End If: End Sub

Const DropTgtMovementDir = "transz"
Const DropTgtMovementMax = 47

Sub PrimDropTgtDown (targetbankname, targetbanknum, swnum)
  PrimDropTgtAnimate swnum, 0
  targetbankname.Hit targetbanknum
End Sub
Sub PrimDropTgtUp  (targetbankname, targetbanknum, swnum, resetvpmtarget, resetvpmbank)
  PrimDropTgtAnimate swnum, 1

  If resetvpmtarget = 1 Then targetbankname.UnHit targetbanknum
  If resetvpmbank = 1 Then targetbankname.DropSol_On
End Sub

Sub PrimDropTgtMove (swNum) 'Customize direction as needed
  If primDir(swNum) = 1 Then 'Up
    Select Case primCnt(swNum)
      Case 0: Execute PrimitivePrefix & swnum & "." & DropTgtMovementDir & "=" & -DropTgtMovementMax * .75
      Case 1: Execute PrimitivePrefix & swnum & "." & DropTgtMovementDir & "=" & -DropTgtMovementMax * .25
      Case 2,3,4: Execute PrimitivePrefix & swnum & "." & DropTgtMovementDir & "=" & 10
      Case 5: Execute PrimitivePrefix & swnum & "." & DropTgtMovementDir & "=" & 0
      Case else: Execute wallPrefix & swnum & ".TimerEnabled = 0"
    End Select
  Else 'Down
    Select Case primCnt(swNum)
      Case 0: Execute PrimitivePrefix & swnum & "." & DropTgtMovementDir & "=" & -DropTgtMovementMax * .25
      Case 1: Execute PrimitivePrefix & swnum & "." & DropTgtMovementDir & "=" & -DropTgtMovementMax * .5
      Case 2: Execute PrimitivePrefix & swnum & "." & DropTgtMovementDir & "=" & -DropTgtMovementMax * .75
      Case 3: Execute PrimitivePrefix & swnum & "." & DropTgtMovementDir & "=" & -DropTgtMovementMax
      Case else: Execute wallPrefix & swnum & ".TimerEnabled = 0"

    End Select
  End If
  primCnt(swnum) = primCnt(swnum) + 1
End Sub

Sub PrimDropTgtAnimate  (swnum, dir)
  primCnt(swnum) = 0
  primDir(swnum) = dir
  Execute wallPrefix & swnum & ".TimerInterval = 10"
  Execute wallPrefix & swnum & ".TimerEnabled = 1"
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
LampTimer.Interval = 5 'lamp fading speed
LampTimer.Enabled = 1

' Lamp & Flasher Timers

Sub LampTimer_Timer()
    Dim chgLamp, num, chg, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp) Then
        For ii = 0 To UBound(chgLamp)
            LampState(chgLamp(ii, 0) ) = chgLamp(ii, 1)       'keep the real state in an array
            FadingLevel(chgLamp(ii, 0) ) = chgLamp(ii, 1) + 4 'actual fading step
        Next
    End If
    UpdateLamps
End Sub


Sub UpdateLamps()
NfadeL 1, L1
NfadeL 2, L2
NfadeL 3, L3
'NfadeL 4, L4  'not used
NfadeL 5, L5
'NfadeL 6, L6  'not used
'NfadeL 7, L7 not used
'NfadeL 8, L8 'not used
NfadeL 9, L9
NfadeL 10, L10
NfadeL 11, L11
NfadeL 12, L12
'NfadeL 13, L13  'Backglass Ball in play
'NFadeL 14, L14  'Backglass 1 player
'NfadeL 15, L15  ''EM Reel player 1-2 Backglass player 1 circle UNVERIFIED
'NfadeL 16, L16  'not used
NfadeL 17, L17
NfadeL 18, L18
NfadeL 19, L19
NfadeL 20, L20
NfadeL 21, L21
'NfadeL 22, L22 'not used
'NfadeL 23, L23 'not used
'NfadeL 24, L24 'not used
NfadeL 25, L25
'NfadeL 26, L26 'not used
'NfadeL 27, L27  'Backglass Game Over
NfadeL 28, L28
'NfadeL 29, L29 'Backglass Highscore to date
'NfadeL 30, L30 'Backglass 2 player
'NfadeL 31, L31 ''EM Reel player 2-2 Backglass player 2 circle UNVERIFIED
'NfadeL 32, L32 'not used
NfadeL 33, L33
NfadeL 34, L34
'NfadeL 35, L35   'not used
NfadeL 36, L36
NfadeL 37, L37
NfadeL 38, L38
'NfadeL 39, L39 'not used
'NfadeL 40, L40 'not used
NfadeL 41, L41
'NfadeL 42, L42 'not used
'NfadeL 43, L43 'not used
NfadeL 44, L44
'NfadeL 45, L45  ''EM Reel Gameover2
'NfadeL 46, L46   'not used
'NfadeL 47, L47  ''EM Reel player 3-2 Backglass player 3 circle UNVERIFIED
'NfadeL 48, L48 'not used
NfadeL 49, L49
NfadeL 50, L50
'NfadeL 51, L51  'not used
NfadeL 52, L52
NfadeL 53, L53
NfadeL 54, L54
'NfadeL 55, L55 'not used
'NfadeL 56, L56 'not used
NfadeL 57, L57
NfadeL 58, L58
NfadeLm 59, L59a
NfadeLm 59, L59b
NfadeL 59, L59c
NfadeL 60, L60
NfadeL 60, L60
'NfadeL 61, L61  'Backglass Tilt
'NfadeL 62, L62 'not used
'NfadeL 63, L63  ''EM Reel player 4-2 Backglass player 4 circle UNVERIFIED
'NfadeL 64, L64 'not used
'NfadeL 65, L65 'not used


End Sub

' div lamp subs

Sub InitLamps()
    Dim x
    For x = 0 to 200
        LampState(x) = 0        ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4      ' used to track the fading state
        FlashSpeedUp(x) = 0.4   ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.2 ' slower speed when turning off the flasher
        FlashMax(x) = 1         ' the maximum value when on, usually 1
        FlashMin(x) = 0         ' the minimum value when off, usually 0
        FlashLevel(x) = 0       ' the intensity of the flashers, usually from 0 to 1
    Next
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

' Flasher objects

Sub Flash(nr, object)
    Select Case FadingLevel(nr)
        Case 4 'off
            FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown(nr)
            If FlashLevel(nr) < FlashMin(nr) Then
                FlashLevel(nr) = FlashMin(nr)
                FadingLevel(nr) = 0 'completely off
            End if
            Object.IntensityScale = FlashLevel(nr)
        Case 5 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
            If FlashLevel(nr) > FlashMax(nr) Then
                FlashLevel(nr) = FlashMax(nr)
                FadingLevel(nr) = 1 'completely on
            End if
            Object.IntensityScale = FlashLevel(nr)
    End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it just sets the flashlevel
    Object.IntensityScale = FlashLevel(nr)
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
' 1st Player
Digits(0) = Array(LED10,LED11,LED12,LED13,LED14,LED15,LED16)
Digits(1) = Array(LED20,LED21,LED22,LED23,LED24,LED25,LED26)
Digits(2) = Array(LED30,LED31,LED32,LED33,LED34,LED35,LED36)
Digits(3) = Array(LED40,LED41,LED42,LED43,LED44,LED45,LED46)
Digits(4) = Array(LED50,LED51,LED52,LED53,LED54,LED55,LED56)
Digits(5) = Array(LED60,LED61,LED62,LED63,LED64,LED65,LED66)

' 2nd Player
Digits(6) = Array(LED80,LED81,LED82,LED83,LED84,LED85,LED86)
Digits(7) = Array(LED90,LED91,LED92,LED93,LED94,LED95,LED96)
Digits(8) = Array(LED100,LED101,LED102,LED103,LED104,LED105,LED106)
Digits(9) = Array(LED110,LED111,LED112,LED113,LED114,LED115,LED116)
Digits(10) = Array(LED120,LED121,LED122,LED123,LED124,LED125,LED126)
Digits(11) = Array(LED130,LED131,LED132,LED133,LED134,LED135,LED136)

' 3rd Player
Digits(12) = Array(LED150,LED151,LED152,LED153,LED154,LED155,LED156)
Digits(13) = Array(LED160,LED161,LED162,LED163,LED164,LED165,LED166)
Digits(14) = Array(LED170,LED171,LED172,LED173,LED174,LED175,LED176)
Digits(15) = Array(LED180,LED181,LED182,LED183,LED184,LED185,LED186)
Digits(16) = Array(LED190,LED191,LED192,LED193,LED194,LED195,LED196)
Digits(17) = Array(LED200,LED201,LED202,LED203,LED204,LED205,LED206)

' 4th Player
Digits(18) = Array(LED220,LED221,LED222,LED223,LED224,LED225,LED226)
Digits(19) = Array(LED230,LED231,LED232,LED233,LED234,LED235,LED236)
Digits(20) = Array(LED240,LED241,LED242,LED243,LED244,LED245,LED246)
Digits(21) = Array(LED250,LED251,LED252,LED253,LED254,LED255,LED256)
Digits(22) = Array(LED260,LED261,LED262,LED263,LED264,LED265,LED266)
Digits(23) = Array(LED270,LED271,LED272,LED273,LED274,LED275,LED276)

' Credits
Digits(24) = Array(LED4,LED2,LED6,LED7,LED5,LED1,LED3)
Digits(25) = Array(LED18,LED9,LED27,LED28,LED19,LED8,LED17)
' Balls
Digits(26) = Array(LED39,LED37,LED48,LED49,LED47,LED29,LED38)
Digits(27) = Array(LED67,LED58,LED69,LED77,LED68,LED57,LED59)

Sub DisplayTimer_Timer
  Dim ChgLED,ii,num,chg,stat,obj
  ChgLed = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
If Not IsEmpty(ChgLED) Then
    If DesktopMode = True Then
    For ii = 0 To UBound(chgLED)
      num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
      if (num < 28) then
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


'**********************************************************************************************************
'**********************************************************************************************************


'Stern Electronic Pinball
'added by Inkochnito
Sub editDips
Dim vpmDips : Set vpmDips = New cvpmDips
With vpmDips
.AddForm 700,400,"Electronic Pinball - DIP switches"
.AddFrame 2,0,190,"Maximum credits",&H00070000,Array("5 credits",0,"10 credits",&H00010000,"15 credits",&H00020000,"20 credits",&H00030000,"25 credits",&H00040000,"30 credits",&H00050000,"35 credits",&H00060000,"40 credits",&H00070000)'dip 17&18&19
.AddFrame 2,130,190,"Outlane special award",&HC0000000,Array("100,000 points",0,"free ball",&H40000000,"free game",&H80000000,"free ball and free game",&HC0000000)'dip 31&32
.AddChk 2,210,180,Array("Match feature",&H00100000)'dip 21
.AddChk 2,230,115,Array("Credits display",&H00080000)'dip 20
.AddFrame 205,0,190,"Balls per game",&H00000040,Array("3 balls",0,"5 balls",&H00000040)'dip 7
.AddFrame 205,46,190,"Pro or novice adjust",32768,Array("all lites on",0,"lites alternating",32768)'dip 16
.AddFrame 205,92,190,"High score award",&H00000020,Array("extra ball",0,"replay",&H00000020)'dip 6
.AddFrame 205,139,190,"Drop targets 2nd time down",&H00800000,Array("special lite on only",0,"special lite on and 1 replay awarded",&H00800000)'dip 24
.AddFrame 205,186,190,"Melody option",&H00000080,Array("2 tones only",0,"full melody",&H00000080)'dip 8
.AddFrame 205,233,190,"High game to date award",&H00004000,Array("points",0,"replay",&H00004000)'dip 15
.AddLabel 50,290,300,20,"After hitting OK, press F3 to reset game with new settings."
.ViewDips
End With
End Sub
Set vpmShowDips = GetRef("editDips")



' *********************************************************************
' *********************************************************************

          'Start of VPX call back Functions

' *********************************************************************
' *********************************************************************

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
  vpmTimer.PulseSw 36
    PlaySoundAtVol SoundFX("right_slingshot",DOFContactors), sling1, 1
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
  vpmTimer.PulseSw 37
    PlaySoundAtVol SoundFX("left_slingshot",DOFContactors), sling2, 1
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0:
    End Select
    LStep = LStep + 1
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
    Next
End Sub

' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX and Rothbauerw
' PlaySoundAtVol sound, 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
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

'Set position as bumperX and Vol manually.

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
      If BallVel(BOT(b) ) > 1 Then
        rolling(b) = True
        if BOT(b).z < 30 Then ' Ball on playfield
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), AudioPan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        Else ' Ball on raised ramp
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*.5, AudioPan(BOT(b) ), 0, Pitch(BOT(b) )+50000, 1, 0, AudioFade(BOT(b) )
        End If
      Else
        If rolling(b) = True Then
          StopSound("fx_ballrolling" & b)
          rolling(b) = False
        End If
      End If
    If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
      PlaySoundAtBOTBallZ "fx_ball_drop" & b, BOT(b)
    End If
    Next
End Sub

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / (VolDiv/VolCol), AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
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

' Sub Spinner_Spin
'   PlaySound "fx_spinner",0,.25,0,0.25
' End Sub

Sub Rubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 20 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End Sub

Sub Posts_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
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


' Thalamus - 2021-04-30 : added proper exit

' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


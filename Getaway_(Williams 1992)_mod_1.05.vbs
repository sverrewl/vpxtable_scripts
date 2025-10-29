'*******************************************************************************
'   _____ ______ _______  __          __ __     __ _    _  _____ _____ _____
'  / ____|  ____|__   __|/\ \        / /\\ \   / /| |  | |/ ____|_   _|_   _|
' | |  __| |__     | |  /  \ \  /\  / /  \\ \_/ (_) |__| | (___   | |   | |
' | | |_ |  __|    | | / /\ \ \/  \/ / /\ \\   /  |  __  |\___ \  | |   | |
' | |__| | |____   | |/ ____ \  /\  / ____ \| |  _| |  | |____) |_| |_ _| |_
'  \_____|______|  |_/_/    \_\/  \/_/    \_\_| (_)_|  |_|_____/|_____|_____|
'
'*******************************************************************************
'
'The Getaway: High Speed II - Williams 1992
'
'Enjoy and remember, winners don't drive drunk.
'
'Big thanks for original authors (32assassin, flupper1, ganjafarmer)!

'MOD 1.0 chokeee and many good guys

Option Explicit
Randomize

'----- Shadow Options -----
Const DynamicBallShadowsOn = 1    '0 = no dynamic ball shadow ("triangles" near slings and such), 1 = enable dynamic ball shadow
Const AmbientBallShadowOn = 1   '0 = Static shadow under ball ("flasher" image, like JP's)
                  '1 = Moving ball shadow ("primitive" object, like ninuzzu's) - This is the only one that behaves like a true shadow!
                  '2 = flasher image shadow, but it moves like ninuzzu's

'///////////////////////-----General Sound Options-----///////////////////////
'// VolumeDial:
'// VolumeDial is the actual global volume multiplier for the mechanical sounds.
'// Values smaller than 1 will decrease mechanical sounds volume.
'// Recommended values should be no greater than 1.

Const BallRollVolume = 0.9      'Level of ball rolling volume. Value between 0 and 1
Const VolumeDial = 1.5
Const TargetBouncerEnabled = 1    '0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.7   'Level of bounces. Recommmended value of 0.7

'********VR Room Init - ********************
Dim Stuff
Dim VRRoom
Dim Thing
Dim Scratches
Dim VRLogoOn
RevolvingLamp False  ' added for VR Beacon init
'*******************************************


'** VR OPTIONS **************
'****************************

VRRoom = 0        '0 = Desktop/FS    1 = VR Room
Scratches = 1   '0 = Scratches OFF 1 = Scratches ON
VRLogoOn = False  'False = VR logo OFF   True = VR Logo ON

'** VR OPTIONS END **********
'****************************


'Set VR Room
if VRRoom = 1 Then
  for each Thing in VRCab : Thing.visible = 1 : next
  for each Thing in VRMinimal : Thing.visible = 1 : next

  SetBackglass

  BeaconTimer.enabled = True  ' 15 seconds at first table load...
  BeaconAttractTimer.enabled = True  ' 15 seconds at first table load...
  Playsound "Relay_GI_Off"
  BeaconSoundTimer.enabled = True

  if Scratches = 1 then GlassImpurities.visible = true

  'Move the sideblades somewhat because I lowered tha cab
  Sideblades.Rotx = -0.65
  Sideblades.size_Z = 93

  'bumper lights intensity was turning black in VR - this fixes...
  F121f.intensity = 25
  F123e.intensity = 25
  F118.intensity = 25

  'swap out main plastics with a better material for VR..
  Wall7.visible = false
  Wall7VR.visible = True

  if VRLogoOn = True then VRLogo.visible = true

  ' Backglass flashers
  VRBGFL21_1.visible = 1
  VRBGFL21_2.visible = 1
  VRBGFL21_3.visible = 1
  VRBGFL21_4.visible = 1
  VRBGFL18_1.visible = 1
  VRBGFL18_2.visible = 1
  VRBGFL18_3.visible = 1
  VRBGFL18_4.visible = 1
  VRBGFL23_1.visible = 1
  VRBGFL23_2.visible = 1
  VRBGFL23_3.visible = 1
  VRBGFL23_4.visible = 1
  VRBGFL19_1.visible = 1
  VRBGFL19_2.visible = 1
  VRBGFL19_3.visible = 1
  VRBGFL19_4.visible = 1
  VRBGFL24_1.visible = 1
  VRBGFL24_2.visible = 1
  VRBGFL24_3.visible = 1
  VRBGFL24_4.visible = 1
end If

' End VR Room init..
'************************************************************************************************************
'************************************************************************************************************

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const BallSize = 50
Const BallMass = 1
Dim tablewidth: tablewidth = Table1.width
Dim tableheight: tableheight = Table1.height

Const cGameName="gw_l5", UseVPMModSol = 2, UseSolenoids=2, UseLamps=1, SSolenoidOn="SolOn", SSolenoidOff="SolOff", SCoin="coin", UseVPMDMD = True

LoadVPM "01560000", "WPC.VBS", 3.46

Dim DesktopMode: DesktopMode = Table1.ShowDT
if Table1.ShowDT = true then Pincab_metals.visible = true

Dim ramptrig:ramptrig=0

Dim Wbl: Wbl=1 ' wobble the supercharger on each loop
If Wbl=0 Then Trigger10.enabled=0

'*************************************************************
'Solenoid Call backs
'**********************************************************************************************************
SolCallback(2)  = "SolRampUp"
SolCallback(3)  = "SolRampDown"
SolCallback(4)  = "LockPost"
SolCallback(7)  = "SolKnocker"
SolCallback(8)  = "SolKickback"              'Kickback
SolCallback(9)  = "bsEjectHole.SolOut"
SolCallback(10) = "SuperchargerDiverter"
SolCallback(11) = "bsTrough.SolOut"
SolCallback(12) = "SolPlunger"        'AutoPlunger
SolCallback(16) = "bsTrough.SolIn"
SolModCallback(17) = "FlashSolMod117"     'Right Middle PF
SolModCallback(18) = "FlashSolMod118"     'Motor toy lights 'BB supercharger
SolModCallback(19) = "FlashSolMod119"     'Right Slingshot  'BB left Copter
SolModCallback(20) = "FlashSolMod120"    'PF
SolModCallback(21) = "FlashSolMod121"     'left Middle PF   'BB Left Ramp
SolModCallback(22) = "FlashSolMod122"     'left Middle PF
SolModCallback(23) = "FlashSolMod123"     'Right Middle PF  'BB flipper
SolModCallback(24) = "FlashSolMod124"             'BB Right Copter
'SolCallBack(31) = "FastFlips.TiltSol"
SolCallback(27) = "RevolvingLamp"       'For VR Beacon Topper
SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"
'SolCallback(sURFlipper) = "SolURFlipper"


'******************************************
'       FLIPPERS
'******************************************

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
    rightflipper1.RotateToEnd

    If rightflipper.currentangle > rightflipper.endangle - ReflipAngle Then
      RandomSoundReflipUpRight RightFlipper
    Else
      SoundFlipperUpAttackRight RightFlipper
      RandomSoundFlipperUpRight RightFlipper
    End If
  Else
    RightFlipper.RotateToStart:RightFlipper1.RotateToStart
    If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
      RandomSoundFlipperDownRight RightFlipper
    End If
    FlipperRightHitParm = FlipperUpSoundLevel
  End If
End Sub

' Flipper collide subs
Sub LeftFlipper_Collide(parm)
  CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
  LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
  CheckLiveCatch Activeball, RightFlipper, RFCount, parm
  RightFlipperCollide parm
End Sub


Sub LeftFlipper_Animate() : batleft.objrotz = LeftFlipper.CurrentAngle + 1 : batleftshadow.objrotz = batleft.objrotz : End Sub
Sub RightFlipper_Animate() : batright.objrotz = RightFlipper.CurrentAngle - 1 : batrightshadow.objrotz  = batright.objrotz : End Sub
Sub RightFlipper1_Animate() : batright1.objrotz = RightFlipper1.CurrentAngle - 1 : batrightshadow1.objrotz  = batright1.objrotz : End Sub

'******************************************************
'  SLINGSHOT CORRECTION FUNCTIONS
'******************************************************
' To add these slingshot corrections:
'   - On the table, add the endpoint primitives that define the two ends of the Slingshot
' - Initialize the SlingshotCorrection objects in InitSlingCorrection
'   - Call the .VelocityCorrect methods from the respective _Slingshot event sub


dim LS : Set LS = New SlingshotCorrection
dim RS : Set RS = New SlingshotCorrection

InitSlingCorrection

Sub InitSlingCorrection

  LS.Object = LeftSlingshot
  LS.EndPoint1 = EndPoint1LS
  LS.EndPoint2 = EndPoint2LS

  RS.Object = RightSlingshot
  RS.EndPoint1 = EndPoint1RS
  RS.EndPoint2 = EndPoint2RS

  'Slingshot angle corrections (pt, BallPos in %, Angle in deg)
  ' These values are best guesses. Retune them if needed based on specific table research.
  AddSlingsPt 0, 0.00,  -4
  AddSlingsPt 1, 0.45,  -7
  AddSlingsPt 2, 0.48,  0
  AddSlingsPt 3, 0.52,  0
  AddSlingsPt 4, 0.55,  7
  AddSlingsPt 5, 1.00,  4

End Sub


Sub AddSlingsPt(idx, aX, aY)        'debugger wrapper for adjusting flipper script in-game
  dim a : a = Array(LS, RS)
  dim x : for each x in a
    x.addpoint idx, aX, aY
  Next
End Sub

'' The following sub are needed, however they may exist somewhere else in the script. Uncomment below if needed
'Dim PI: PI = 4*Atn(1)
'Function dSin(degrees)
' dsin = sin(degrees * Pi/180)
'End Function
'Function dCos(degrees)
' dcos = cos(degrees * Pi/180)
'End Function
'
Function RotPoint(x,y,angle)
    dim rx, ry
    rx = x*dCos(angle) - y*dSin(angle)
    ry = x*dSin(angle) + y*dCos(angle)
    RotPoint = Array(rx,ry)
End Function

Class SlingshotCorrection
  Public DebugOn, Enabled
  private Slingshot, SlingX1, SlingX2, SlingY1, SlingY2

  Public ModIn, ModOut
  Private Sub Class_Initialize : redim ModIn(0) : redim Modout(0): Enabled = True : End Sub

  Public Property let Object(aInput) : Set Slingshot = aInput : End Property
  Public Property Let EndPoint1(aInput) : SlingX1 = aInput.x: SlingY1 = aInput.y: End Property
  Public Property Let EndPoint2(aInput) : SlingX2 = aInput.x: SlingY2 = aInput.y: End Property

  Public Sub AddPoint(aIdx, aX, aY)
    ShuffleArrays ModIn, ModOut, 1 : ModIn(aIDX) = aX : ModOut(aIDX) = aY : ShuffleArrays ModIn, ModOut, 0
    If gametime > 100 then Report
  End Sub

  Public Sub Report()         'debug, reports all coords in tbPL.text
    If not debugOn then exit sub
    dim a1, a2 : a1 = ModIn : a2 = ModOut
    dim str, x : for x = 0 to uBound(a1) : str = str & x & ": " & round(a1(x),4) & ", " & round(a2(x),4) & vbnewline : next
    TBPout.text = str
  End Sub


  Public Sub VelocityCorrect(aBall)
    dim BallPos, XL, XR, YL, YR

    'Assign right and left end points
    If SlingX1 < SlingX2 Then
      XL = SlingX1 : YL = SlingY1 : XR = SlingX2 : YR = SlingY2
    Else
      XL = SlingX2 : YL = SlingY2 : XR = SlingX1 : YR = SlingY1
    End If

    'Find BallPos = % on Slingshot
    If Not IsEmpty(aBall.id) Then
      If ABS(XR-XL) > ABS(YR-YL) Then
        BallPos = PSlope(aBall.x, XL, 0, XR, 1)
      Else
        BallPos = PSlope(aBall.y, YL, 0, YR, 1)
      End If
      If BallPos < 0 Then BallPos = 0
      If BallPos > 1 Then BallPos = 1
    End If

    'Velocity angle correction
    If not IsEmpty(ModIn(0) ) then
      Dim Angle, RotVxVy
      Angle = LinearEnvelope(BallPos, ModIn, ModOut)
      'debug.print " BallPos=" & BallPos &" Angle=" & Angle
      'debug.print " BEFORE: aBall.Velx=" & aBall.Velx &" aBall.Vely" & aBall.Vely
      RotVxVy = RotPoint(aBall.Velx,aBall.Vely,Angle)
      If Enabled then aBall.Velx = RotVxVy(0)
      If Enabled then aBall.Vely = RotVxVy(1)
      'debug.print " AFTER: aBall.Velx=" & aBall.Velx &" aBall.Vely" & aBall.Vely
      'debug.print " "
    End If
  End Sub

End Class



'******************************************************
'****  PHYSICS DAMPENERS
'******************************************************
'
' These are data mined bounce curves,
' dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
' Requires tracking ballspeed to calculate COR


'******************************************************
' VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************

sub TargetBouncer(aBall,defvalue)
    dim zMultiplier, vel, vratio
    if TargetBouncerEnabled = 1 and aball.z < 30 then
        'debug.print "velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz
        vel = BallSpeed(aBall)
        if aBall.velx = 0 then vratio = 1 else vratio = aBall.vely/aBall.velx
        Select Case Int(Rnd * 6) + 1
            Case 1: zMultiplier = 0.2*defvalue
      Case 2: zMultiplier = 0.25*defvalue
            Case 3: zMultiplier = 0.3*defvalue
      Case 4: zMultiplier = 0.4*defvalue
            Case 5: zMultiplier = 0.45*defvalue
            Case 6: zMultiplier = 0.5*defvalue
        End Select
        aBall.velz = abs(vel * zMultiplier * TargetBouncerFactor)
        aBall.velx = sgn(aBall.velx) * sqr(abs((vel^2 - aBall.velz^2)/(1+vratio^2)))
        aBall.vely = aBall.velx * vratio
        'debug.print "---> velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz
        'debug.print "conservation check: " & BallSpeed(aBall)/vel
  end if
end sub

' Add targets or posts to the TargetBounce collection if you want to activate the targetbouncer code from them
Sub TargetBounce_Hit
  TargetBouncer activeball, 1
End Sub


Sub dPosts_Hit(idx)
  RubbersD.dampen Activeball
  TargetBouncer Activeball, 1
End Sub

Sub dSleeves_Hit(idx)
  SleevesD.Dampen Activeball
  TargetBouncer Activeball, 0.7
End Sub


dim RubbersD : Set RubbersD = new Dampener        'frubber
RubbersD.name = "Rubbers"
RubbersD.debugOn = False        'shows info in textbox "TBPout"
RubbersD.Print = False        'debug, reports in debugger (in vel, out cor)
'cor bounce curve (linear)
'for best results, try to match in-game velocity as closely as possible to the desired curve
'RubbersD.addpoint 0, 0, 0.935        'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 0, 0, 1.1        'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 1, 3.77, 0.97
RubbersD.addpoint 2, 5.76, 0.967        'dont take this as gospel. if you can data mine rubber elasticitiy, please help!
RubbersD.addpoint 3, 15.84, 0.874
RubbersD.addpoint 4, 56, 0.64        'there's clamping so interpolate up to 56 at least

dim SleevesD : Set SleevesD = new Dampener        'this is just rubber but cut down to 85%...
SleevesD.name = "Sleeves"
SleevesD.debugOn = False        'shows info in textbox "TBPout"
SleevesD.Print = False        'debug, reports in debugger (in vel, out cor)
SleevesD.CopyCoef RubbersD, 0.85

'######################### Add new FlippersD Profile
'#########################    Adjust these values to increase or lessen the elasticity

dim FlippersD : Set FlippersD = new Dampener
FlippersD.name = "Flippers"
FlippersD.debugOn = False
FlippersD.Print = False
FlippersD.addpoint 0, 0, 1.1
FlippersD.addpoint 1, 3.77, 0.99
FlippersD.addpoint 2, 6, 0.99

Class Dampener
  Public Print, debugOn 'tbpOut.text
  public name, Threshold         'Minimum threshold. Useful for Flippers, which don't have a hit threshold.
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
    RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id)+0.0001)
    coef = desiredcor / realcor
    if debugOn then str = name & " in vel:" & round(cor.ballvel(aBall.id),2 ) & vbnewline & "desired cor: " & round(desiredcor,4) & vbnewline & _
    "actual cor: " & round(realCOR,4) & vbnewline & "ballspeed coef: " & round(coef, 3) & vbnewline
    if Print then debug.print Round(cor.ballvel(aBall.id),2) & ", " & round(desiredcor,3)

    aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
    if debugOn then TBPout.text = str
  End Sub

  public sub Dampenf(aBall, parm) 'Rubberizer is handle here
    dim RealCOR, DesiredCOR, str, coef
    DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
    RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id)+0.0001)
    coef = desiredcor / realcor
    If abs(aball.velx) < 2 and aball.vely < 0 and aball.vely > -3.75 then
      aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
    End If
  End Sub

  Public Sub CopyCoef(aObj, aCoef) 'alternative addpoints, copy with coef
    dim x : for x = 0 to uBound(aObj.ModIn)
      addpoint x, aObj.ModIn(x), aObj.ModOut(x)*aCoef
    Next
  End Sub


  Public Sub Report()         'debug, reports all coords in tbPL.text
    if not debugOn then exit sub
    dim a1, a2 : a1 = ModIn : a2 = ModOut
    dim str, x : for x = 0 to uBound(a1) : str = str & x & ": " & round(a1(x),4) & ", " & round(a2(x),4) & vbnewline : next
    TBPout.text = str
  End Sub

End Class


'******************************************************
'   FLIPPER CORRECTION INITIALIZATION
'******************************************************

dim LF : Set LF = New FlipperPolarity
dim RF : Set RF = New FlipperPolarity

InitPolarity

Sub InitPolarity()
  dim x, a : a = Array(LF, RF)
  for each x in a
    x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1        'disabled
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

  addpt "Velocity", 0, 0,         1
  addpt "Velocity", 1, 0.16, 1.06
  addpt "Velocity", 2, 0.41,         1.05
  addpt "Velocity", 3, 0.53,         1'0.982
  addpt "Velocity", 4, 0.702, 0.968
  addpt "Velocity", 5, 0.95,  0.968
  addpt "Velocity", 6, 1.03,         0.945

  LF.Object = LeftFlipper
  LF.EndPoint = EndPointLp
  RF.Object = RightFlipper
  RF.EndPoint = EndPointRp
End Sub

' Flipper trigger hit subs
Sub TriggerLF_Hit() : LF.Addball activeball : End Sub
Sub TriggerLF_UnHit() : LF.PolarityCorrect activeball : End Sub
Sub TriggerRF_Hit() : RF.Addball activeball : End Sub
Sub TriggerRF_UnHit() : RF.PolarityCorrect activeball : End Sub


'******************************************************
'  FLIPPER CORRECTION FUNCTIONS
'******************************************************

Sub AddPt(aStr, idx, aX, aY)        'debugger wrapper for adjusting flipper script in-game
  dim a : a = Array(LF, RF)
  dim x : for each x in a
    x.addpoint aStr, idx, aX, aY
  Next
End Sub

Class FlipperPolarity
  Public DebugOn, Enabled
  Private FlipAt        'Timer variable (IE 'flip at 723,530ms...)
  Public TimeDelay        'delay before trigger turns off and polarity is disabled TODO set time!
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

  Public Sub Report(aChooseArray)         'debug, reports all coords in tbPL.text
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
  Private Function FlipperOn() : if gameTime < FlipAt+TimeDelay then FlipperOn = True : End If : End Function        'Timer shutoff for polaritycorrect

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
          if ballpos > 0.65 then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)                                'find safety coefficient 'ycoef' data
        end if
      Next

      If BallPos = 0 Then 'no ball data meaning the ball is entering and exiting pretty close to the same position, use current values.
        BallPos = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
        if ballpos > 0.65 then  Ycoef = LinearEnvelope(aBall.Y, YcoefIn, YcoefOut)                                                'find safety coefficient 'ycoef' data
      End If

      'Velocity correction
      if not IsEmpty(VelocityIn(0) ) then
        Dim VelCoef
        VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)

        if partialflipcoef < 1 then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)

        if Enabled then aBall.Velx = aBall.Velx*VelCoef
        if Enabled then aBall.Vely = aBall.Vely*VelCoef
      End If

      'Polarity Correction (optional now)
      if not IsEmpty(PolarityIn(0) ) then
        If StartPoint > EndPoint then LR = -1        'Reverse polarity if left flipper
        dim AddX : AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR

        if Enabled then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef)
      End If
    End If
    RemoveBall aBall
  End Sub
End Class


'******************************************************
'  FLIPPER POLARITY AND RUBBER DAMPENER SUPPORTING FUNCTIONS
'******************************************************

' Used for flipper correction and rubber dampeners
Sub ShuffleArray(ByRef aArray, byVal offset) 'shuffle 1d array
  dim x, aCount : aCount = 0
  redim a(uBound(aArray) )
  for x = 0 to uBound(aArray)        'Shuffle objects in a temp array
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
  redim aArray(aCount-1+offset)        'Resize original array
  for x = 0 to aCount-1                'set objects back into original array
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
Function PSlope(Input, X1, Y1, X2, Y2)        'Set up line via two points, no clamping. Input X, output Y
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
  dim ii : for ii = 1 to uBound(xKeyFrame)        'find active line
    if xInput <= xKeyFrame(ii) then L = ii : exit for : end if
  Next
  if xInput > xKeyFrame(uBound(xKeyFrame) ) then L = uBound(xKeyFrame)        'catch line overrun
  Y = pSlope(xInput, xKeyFrame(L-1), yLvl(L-1), xKeyFrame(L), yLvl(L) )

  if xInput <= xKeyFrame(lBound(xKeyFrame) ) then Y = yLvl(lBound(xKeyFrame) )         'Clamp lower
  if xInput >= xKeyFrame(uBound(xKeyFrame) ) then Y = yLvl(uBound(xKeyFrame) )        'Clamp upper

  LinearEnvelope = Y
End Function



'******************************************************
'  FLIPPER TRICKS
'******************************************************

RightFlipper.timerinterval=1
Rightflipper.timerenabled=True

sub RightFlipper_timer()
  FlipperTricks LeftFlipper, LFPress, LFCount, LFEndAngle, LFState
  FlipperTricks RightFlipper, RFPress, RFCount, RFEndAngle, RFState
  FlipperNudge RightFlipper, RFEndAngle, RFEOSNudge, LeftFlipper, LFEndAngle
  FlipperNudge LeftFlipper, LFEndAngle, LFEOSNudge,  RightFlipper, RFEndAngle
end sub

Dim LFEOSNudge, RFEOSNudge

Sub FlipperNudge(Flipper1, Endangle1, EOSNudge1, Flipper2, EndAngle2)
  Dim b, BOT
  BOT = GetBalls

  If Flipper1.currentangle = Endangle1 and EOSNudge1 <> 1 Then
    EOSNudge1 = 1
    'debug.print Flipper1.currentangle &" = "& Endangle1 &"--"& Flipper2.currentangle &" = "& EndAngle2
    If Flipper2.currentangle = EndAngle2 Then
      For b = 0 to Ubound(BOT)
        If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper1) Then
          'Debug.Print "ball in flip1. exit"
          exit Sub
        end If
      Next
      For b = 0 to Ubound(BOT)
        If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper2) Then
          BOT(b).velx = BOT(b).velx / 1.3
          BOT(b).vely = BOT(b).vely - 0.5
        end If
      Next
    End If
  Else
    If Abs(Flipper1.currentangle) > Abs(EndAngle1) + 30 then EOSNudge1 = 0
  End If
End Sub

'*****************
' Maths
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


'*************************************************
'  Check ball distance from Flipper for Rem
'*************************************************

Function Distance(ax,ay,bx,by)
  Distance = SQR((ax - bx)^2 + (ay - by)^2)
End Function

Function DistancePL(px,py,ax,ay,bx,by) ' Distance between a point and a line where point is px,py
  DistancePL = ABS((by - ay)*px - (bx - ax) * py + bx*ay - by*ax)/Distance(ax,ay,bx,by)
End Function

Function Radians(Degrees)
  Radians = Degrees * PI /180
End Function

Function AnglePP(ax,ay,bx,by)
  AnglePP = Atn2((by - ay),(bx - ax))*180/PI
End Function

Function DistanceFromFlipper(ballx, bally, Flipper)
  DistanceFromFlipper = DistancePL(ballx, bally, Flipper.x, Flipper.y, Cos(Radians(Flipper.currentangle+90))+Flipper.x, Sin(Radians(Flipper.currentangle+90))+Flipper.y)
End Function

Function FlipperTrigger(ballx, bally, Flipper)
  Dim DiffAngle
  DiffAngle  = ABS(Flipper.currentangle - AnglePP(Flipper.x, Flipper.y, ballx, bally) - 90)
  If DiffAngle > 180 Then DiffAngle = DiffAngle - 360

  If DistanceFromFlipper(ballx,bally,Flipper) < 48 and DiffAngle <= 90 and Distance(ballx,bally,Flipper.x,Flipper.y) < Flipper.Length Then
    FlipperTrigger = True
  Else
    FlipperTrigger = False
  End If
End Function


'*************************************************
'  End - Check ball distance from Flipper for Rem
'*************************************************

dim LFPress, RFPress, LFCount, RFCount
dim LFState, RFState
dim EOST, EOSA,Frampup, FElasticity,FReturn
dim RFEndAngle, LFEndAngle

Const FlipperCoilRampupMode = 0     '0 = fast, 1 = medium, 2 = slow (tap passes should work)

LFState = 1
RFState = 1
EOST = leftflipper.eostorque
EOSA = leftflipper.eostorqueangle
Frampup = LeftFlipper.rampup
FElasticity = LeftFlipper.elasticity
FReturn = LeftFlipper.return
'Const EOSTnew = 1 'EM's to late 80's
Const EOSTnew = 0.8 '90's and later
Const EOSAnew = 1
Const EOSRampup = 0
Dim SOSRampup
Select Case FlipperCoilRampupMode
  Case 0:
    SOSRampup = 2.5
  Case 1:
    SOSRampup = 6
  Case 2:
    SOSRampup = 8.5
End Select

Const LiveCatch = 16
Const LiveElasticity = 0.45
Const SOSEM = 0.815
'Const EOSReturn = 0.055  'EM's
'Const EOSReturn = 0.045  'late 70's to mid 80's
Const EOSReturn = 0.018  'mid 80's to early 90's
'Const EOSReturn = 0.025  'mid 90's and later

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
    Dim b, BOT
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
Const LiveDistanceMax = 114  'maximum distance in vp units from flipper base live catch dampening will occur (tip protection)

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
        If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 1 Then FlippersD.Dampenf Activeball, parm
  End If
End Sub


'******************************************************
'****  END FLIPPER CORRECTIONS
'******************************************************



'Solenoid Controlled toys
'**********************************************************************************************************

Sub SolKnocker(Enabled)
        If enabled Then
                KnockerSolenoid 'Add knocker position object
        End If
End Sub


Sub SolRampUp(Enabled)
    If Enabled Then
        Controller.Switch(54) = False
        RampaMovil.Collidable = False
        playsound SoundFX("fx_diverterUp",DOFContactors)
    div_ramp.ObjRotY=15
    ramptrig=0
    End If
End Sub

Sub SolRampDown(Enabled)
    If Enabled Then
        Controller.Switch(54) = true
        RampaMovil.Collidable = True
        playsound SoundFX("fx_diverterUp",DOFContactors)
    div_ramp.ObjRotY=0
    ramptrig=1
    End If
End sub

Sub SuperchargerDiverter(enabled)
  If Enabled Then
    Wall29.collidable=True
    sc_div.ObjRotZ=0
    PlaySound SoundFX("fx_DiverterUp",DOFContactors)
  Else
    Wall29.collidable=False
    sc_div.ObjRotZ=-43
    PlaySound SoundFX("sc_diverter",DOFContactors)
  End If
End Sub

'plunger
Sub SolPlunger(Enabled)
  If Enabled Then
    plungerIM.AutoFire
  End If
End Sub

'kickback
Sub SolKickback(Enabled)
    If enabled Then
       Plunger1.Fire
       PlaySound SoundFX("Popper",DOFContactors)
    Else
       Plunger1.PullBack
    End If
End Sub

'kicker ball bounce

'hit to enable wall lock
Sub e_kick_Hit()
  ek.enabled=1
  PlaySound "metalhit_medium"
  PlaySound "fx_rr6"
End Sub

'lock ball between walls (80ms)
Sub ek_Timer
  l_kick.collidable=1
  kk.enabled=1
  ek.enabled=0
End Sub

'destroy walls and enable real kicker (400ms)
Sub kk_timer
  sw77.enabled=1
  l_kick.collidable=0
  kk.enabled=0
End Sub

'disable real kicker after release
Sub k_off_Hit()
  sw77.enabled=0
End Sub


'VR Beacon...

Sub RevolvingLamp(Enabled)
  If VRroom > 0 then
    If Enabled Then
      BeaconTimer.enabled = True
      Playsound "Relay_GI_Off"

      BeaconSoundTimer.interval = 1
      BeaconSoundTimer.enabled = True
      'BeaconSoundTimer.interval = 2050
      'We need to turn ON all the reflections and main flasher light here..

      BeaconFR.visible = true

      BeaconREF1.visible = true
      BeaconREF1b.visible = true
      BeaconREF1.visible = true
      BeaconREF1b.visible = true
      BeaconREF1.visible = true
      BeaconREF1b.visible = true
      BeaconREF2.visible = true
      BeaconREF2b.visible = true
      BeaconREF2.visible = true
      BeaconREF2b.visible = true
      BeaconREF2.visible = true
      BeaconREF2b.visible = true

    Else
      BeaconTimer.enabled = False
      BeaconSoundTimer.enabled = False
      BeaconSoundTimer.interval = 1


      Stopsound "Test_Beaconcut3"
      Playsound "Relay_Flash_Off"
      'We need to turn OFF all the reflections and main flasher light here..

      BeaconFR.visible = false

      BeaconREF1.visible = false
      BeaconREF1b.visible = false
      BeaconREF1.visible = false
      BeaconREF1b.visible = false
      BeaconREF1.visible = false
      BeaconREF1b.visible = false
      BeaconREF2.visible = false
      BeaconREF2b.visible = false
      BeaconREF2.visible = false
      BeaconREF2b.visible = false
      BeaconREF2.visible = false
      BeaconREF2b.visible = false

    End If
  End if
End Sub


'Visible Lock, original PD

  dim postdown : postdown = false

  Sub LockPost(enabled)
  visibleLock.Solexit enabled
    If enabled then
      If postdown = false Then : PosteArriba.IsDropped=1 : postlock.z=-10 : Playsound "sc_diverter" : vpmTimer.AddTimer 400, "RaisePost" : End If
      postdown = true
    End If
  End Sub

 Sub RaisePost(aSw) : If postdown = true Then : PosteArriba.IsDropped = 0 : postlock.z=45 : PlaySound "sc_diverter" : End If : postdown = false : End Sub

'**************************
'           GI
'**************************
dim gilvl:gilvl = 1   'General Illumination light state tracked for Dynamic Ball Shadows
dim gilvl1,gilvl2,gilvl3,gilvl4,gilvl5 : gilvl1 = 0 : gilvl2 = 0 : gilvl3 = 0 : gilvl4 = 0 : gilvl5 = 0
Set GiCallback2 = GetRef("UpdateGI")
Sub UpdateLights(col, v)
  Dim xx : For each xx in col
    if TypeName(xx) = "Light" then
      xx.State = v
    Else
      xx.Opacity = v * 100.0
    End If
  Next
End Sub
Sub UpdateGI(no, v)
  'Debug.Print no & " = " & v
  Select Case no
    Case 0 : gilvl1 = v : UpdateLights GIString1, v ' Playfield String #1
    Case 1 : gilvl2 = v : UpdateLights GIString2, v ' Playfield String #2
    Case 2 : gilvl3 = v : UpdateLights GIString3, v ' Backglass String #3
    Case 3 : gilvl4 = v : UpdateLights GIString4, v ' Backglass String #4
    Case 4 : gilvl5 = v : UpdateLights GIString5, v ' Backglass String #5
  End Select
  BGBright.Opacity = 50.0 + 50.0 * (gilvl3 + gilvl4 + gilvl5) / 3.0
  Dim vavg : vavg = 0.5 * (gilvl1 + gilvl2)
  GIOverall.State = vavg

  wireRamps.blenddisableLighting=0.1 + vavg * 0.3
  metalwalls.blenddisableLighting=0.1 + vavg * 0.3
  '   For each xx in Plastics:xx.BlendDisableLighting = .41 * v: Next
  ''    batleft.BlendDisableLighting = .11 * v:batright.BlendDisableLighting = .11 * v:batright1.BlendDisableLighting = .11 * v
  Primitive35.BlendDisableLighting = .04 * vavg: Primitive36.BlendDisableLighting = .04 * vavg 'sling cover
  Primitive12.BlendDisableLighting = .06 * vavg
  Primitive56.BlendDisableLighting = .03 * vavg
  Primitive57.BlendDisableLighting = .03 * vavg
  Primitive46.BlendDisableLighting = .2 * vavg
  '   Primitive11.BlendDisableLighting = .11 * v:Primitive12.BlendDisableLighting = .11 * v
  '   l66.BlendDisableLighting = .11 * v:l67.BlendDisableLighting = .11 * v
  Wall26.BlendDisableLighting = .08 * vavg 'ramp decal
  '   Primitive45.BlendDisableLighting = .21 * v : Primitive46.BlendDisableLighting = .31 * v' drive 65
  donut.BlendDisableLighting = .04 * vavg ' donut heaven
  Primitive34.BlendDisableLighting = .04 * vavg
  '   dome1.BlendDisableLighting = .15 * v:dome2.BlendDisableLighting = .15 * v:dome3.BlendDisableLighting = .15 * v ' traffic lights
  superramp_p.BlendDisableLighting = vavg * 2 ' supercharger and ramp
  '   supercharger_p.BlendDisableLighting = 1.0 - 0.9 * v:
  '   refL.opacity=7 * v:refR.opacity=7 * v
  '   PinCab_Backglass.blenddisablelighting = 0.2 + v * 7.8
  ' Domes lighten..
  redlight.blenddisablelighting = 0.3 * vavg
  orangelight.blenddisablelighting = 0.3 * vavg
  greenlight.blenddisablelighting = 0.3 * vavg

  wall_refl.visible=1
  RLS.visible=1
  RLS1.visible=1
  RRS.visible=1
  RRS1.visible=1
  ' refL.visible=1
  refR.visible=1
  Fbumper1.visible=1
  Fbumper2.visible=1
  Fbumper3.visible=1
  f_r1.visible=1
  f_r2.visible=1
  f_r3.visible=1
  f_r4.visible=1
  GIon.visible=1
  GIoff.visible=1

  wall_refl.Opacity = 100 * vavg
  RLS.Opacity = 100 * vavg
  RLS1.Opacity = 100 * vavg
  RRS.Opacity = 100 * vavg
  RRS1.Opacity = 100 * vavg
  ' refL.Opacity = 100 * vavg
  refR.Opacity = 100 * vavg
  Fbumper1.Opacity = 100 * vavg
  Fbumper2.Opacity = 100 * vavg
  Fbumper3.Opacity = 100 * vavg
  f_r1.Opacity = 100 * vavg
  f_r2.Opacity = 100 * vavg
  f_r3.Opacity = 100 * vavg
  f_r4.Opacity = 100 * vavg
  GIon.Opacity = 100 * vavg
  GIoff.Opacity = 100 * vavg


  If vavg > 0.5 And giLvl < 0.5 Then
    Sideblades.image="sidewallsON"
    ' DOF 101, DOFOn
    If False Then
    wall_refl.visible=1
    RLS.visible=1
    RLS1.visible=1
    RRS.visible=1
    RRS1.visible=1
    '   refL.visible=1
    refR.visible=1
    Fbumper1.visible=1
    Fbumper2.visible=1
    Fbumper3.visible=1
    f_r1.visible=1
    f_r2.visible=1
    f_r3.visible=1
    f_r4.visible=1
    GIon.visible=1
    GIoff.visible=0
    End If
  ElseIf vavg < 0.5 And giLvl > 0.5 Then
    Sideblades.image="sidewallsOFF"
    ' DOF 101, DOFOff
    If False Then
    wall_refl.visible=0
    RLS.visible=0
    RLS1.visible=0
    RRS.visible=0
    RRS1.visible=0
'   refL.visible=0
    refR.visible=0
    Fbumper1.visible=0
    Fbumper2.visible=0
    Fbumper3.visible=0
    f_r1.visible=0
    f_r2.visible=0
    f_r3.visible=0
    f_r4.visible=0
    GIoff.visible=1
    GIon.visible=0
    End If
  End If
  gilvl = vavg
End Sub

Sub Siren(enabled)
    If enabled Then

  Else

  End If
End Sub


'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************
Dim bsTrough, bsEjectHole, visibleLock
Dim bsSwitch74, bsSwitch75, bsSwitch76

Sub Table1_Init
  vpmInit Me
  vpmMapLights Lamps
  On Error Resume Next
    With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine = "Gateway High Speed II Williams" & chr(13) & "You Suck"
    .HandleMechanics=0
    .HandleKeyboard=0
    .ShowDMDOnly=1
    .ShowFrame=0
    .ShowTitle=0
        .hidden = 0
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With
     On Error Goto 0

    vpmNudge.TiltSwitch = 14
    vpmNudge.Sensitivity = 3
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingshot)

  Set bsTrough=New cvpmBallStack
    bsTrough.InitSw 55,58,57,56,0,0,0,0
    bsTrough.InitKick BallRelease,85,7
'   bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
    bsTrough.Balls=3

  Set bsEjectHole=New cvpmBallStack
    bsEjectHole.InitSaucer sw77,77,100,10
'     bsEjectHole.InitExitSnd SoundFX("fx_saucer_exit",DOFContactors), SoundFX("Solenoid",DOFContactors)

  Set visibleLock = New cvpmVLock
    visibleLock.InitVLock Array(sw76,sw75, sw74),Array(k76, k75, k74), Array(76,75,74)
    visibleLock.InitSnd SoundFX("sc_diverter",DOFContactors), SoundFX("Solenoid",DOFContactors)
    visibleLock.ExitDir = 180
    visibleLock.ExitForce = 0
    visibleLock.createevents "visibleLock"

    Controller.Switch(22) = True  ' Coin Door Closed
    Controller.Switch(24) = True  ' Always Closed Switch

    Set bsSwitch74 = New cvpmBallStack : bsSwitch74.Initsw 0,0,0,0,0,0,0,0
    Set bsSwitch75 = New cvpmBallStack : bsSwitch75.Initsw 0,0,0,0,0,0,0,0
    Set bsSwitch76 = New cvpmBallStack : bsSwitch76.Initsw 0,0,0,0,0,0,0,0

  Plunger1.PullBack

  FlashSolMod117 0
  FlashSolMod118 0
  FlashSolMod119 0
  FlashSolMod120 0
  FlashSolMod121 0
  FlashSolMod122 0
  FlashSolMod123 0
  FlashSolMod124 0

End Sub


'**********
'Timer Code
'**********
Sub FrameTimer_Timer()
  RollingUpdate         'update rolling sounds
' DoDTAnim            'handle drop target animations
' DoSTAnim            'handle stand up target animations
  if Vrroom > 0 then VRStartButton.blenddisableLighting = L68.IntensityScale ' for front start button VR
  If DynamicBallShadowsOn Or AmbientBallShadowOn Then DynamicBSUpdate 'update ball shadows
End Sub

Sub CorTimer_Timer() : Cor.Update : End Sub


'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Dim BIPL : BIPL=0

Sub Table1_KeyDown(ByVal KeyCode)
  If keycode = keyInsertCoin1 or keycode = keyInsertCoin2 or keycode = keyInsertCoin3 or keycode = keyInsertCoin4 Then
    Select Case Int(rnd*3)
                        Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
                        Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
                        Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
    End Select
  End If
  if keycode=StartGameKey then soundStartButton()
  If keycode = PlungerKey or keycode = LockBarKey Then Controller.Switch(34) = True
  if KeyCode = LeftTiltKey Then Nudge 90, 5:SoundNudgeLeft()
  if KeyCode = RightTiltKey Then Nudge 270, 5:SoundNudgeRight()
  if KeyCode = CenterTiltKey Then Nudge 0, 5:SoundNudgeCenter()
  If keycode = LeftMagnaSave Then:Controller.Switch(33) = 1:End If
  'for vr..
  If keycode = PlungerKey and VRroom > 0 Then VRshifter.objrotX = VRShifter.ObjrotX -20
  If keycode = LeftMagnaSave and VRroom > 0 Then VRshifter.objrotX = VRShifter.ObjrotX +20: End If
  If keycode = LeftFlipperKey and VRroom > 0 then VRFlipperButtonLeft.x = VRFlipperButtonLeft.x + 5
  If keycode = RightFlipperKey and VRroom > 0 then VRFlipperButtonRight.x = VRFlipperButtonRight.x - 5
  if keycode=StartGameKey and VRroom >0 then VRStartButton.y = VRStartButton.y - 5: VRStartButton2.y = VRStartButton2.y -5

  If keycode = RightFlipperKey Then FlipperActivate RightFlipper, RFPress End If
  If keycode = LeftFlipperKey Then FlipperActivate LeftFlipper, LFPress End If

  If KeyDownHandler(keycode) Then Exit Sub
  '(Do not use Exit Sub, this script does not handle switch handling at all!)
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
  If keycode = PlungerKey or keycode = LockBarKey Then Controller.Switch(34) = False
  If keycode = LeftMagnaSave Then:Controller.Switch(33) = False:End If
  'for vr..
  If keycode = PlungerKey and VRroom > 0 Then VRshifter.objrotX = VRShifter.ObjrotX +20
  If keycode = LeftMagnaSave and VRroom > 0 Then VRshifter.objrotX = VRShifter.ObjrotX -20: End If
  If keycode = RightFlipperKey and VRroom > 0 then VRFlipperButtonRight.x = VRFlipperButtonRight.x + 5
  If keycode = LeftFlipperKey and VRroom > 0 then VRFlipperButtonLeft.x = VRFlipperButtonLeft.x - 5
  if keycode=StartGameKey and VRroom >0 then VRStartButton.y = VRStartButton.y + 5: VRStartButton2.y = VRStartButton2.y +5

  If keycode = RightFlipperKey Then FlipperDeActivate RightFlipper, RFPress End If
  If keycode = LeftFlipperKey Then FlipperDeActivate LeftFlipper, LFPress End If

  'If KeyCode = PlungerKey Then
  '                Plunger.Fire
  '                If BIPL = 1 Then
  '                        SoundPlungerReleaseBall()                        'Plunger release sound when there is a ball in shooter lane
  '                Else
  '                        SoundPlungerReleaseNoBall()                        'Plunger release sound when there is no ball in shooter lane
  '                End If
  '        End If
  '
  If KeyUpHandler(keycode) Then Exit Sub
End Sub


Dim plungerIM
Set plungerIM = New cvpmImpulseP
With plungerIM
  .InitImpulseP ShooterLane, 55, 0.1
  .Random 0.2
  .Switch 78
  .InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
  .CreateEvents "plungerIM"
End With



'Sub Table1_KeyDown(ByVal KeyCode)
' If KeyDownHandler(keycode) Then Exit Sub
' If keycode = PlungerKey Then Controller.Switch(34) = True: shifter.objrotX = Shifter.ObjrotX -20
' if KeyCode = LeftTiltKey Then Nudge 90, 4
' if KeyCode = RightTiltKey Then Nudge 270, 4
' if KeyCode = CenterTiltKey Then Nudge 0, 12
' If keycode = LeftMagnaSave Then:Controller.Switch(33) = 1: shifter.objrotX = Shifter.ObjrotX +20: End If
' If KeyCode = LeftFlipperKey then FlipperActivate LeftFlipper, LFPress
' If KeyCode = RightFlipperKey then FlipperActivate RightFlipper, RFPress
''(Do not use Exit Sub, this script does not handle switch handling at all!)



' If keycode = LeftTiltKey Then Nudge 90, 5:SoundNudgeLeft()
'        If keycode = RightTiltKey Then Nudge 270, 5:SoundNudgeRight()
'        If keycode = CenterTiltKey Then Nudge 0, 3:SoundNudgeCenter()
'
'        If keycode = keyInsertCoin1 or keycode = keyInsertCoin2 or keycode = keyInsertCoin3 or keycode = keyInsertCoin4 Then
'                Select Case Int(rnd*3)
'                        Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
'                        Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
'                        Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
'
'                End Select
'        End If
'
'        If KeyCode = PlungerKey Then Plunger.Pullback:SoundPlungerPull()
'
'        if keycode=StartGameKey then soundStartButton()

'End Sub
'
'Sub Table1_KeyUp(ByVal KeyCode)
' If KeyUpHandler(keycode) Then Exit Sub
' If keycode = PlungerKey Then Controller.Switch(34) = False : shifter.objrotX = Shifter.ObjrotX +20
' If keycode = LeftMagnaSave Then:Controller.Switch(33) = False : shifter.objrotX = Shifter.ObjrotX -20: End If
' If KeyCode = LeftFlipperKey then FlipperDeactivate LeftFlipper, LFPress
' If KeyCode = RightFlipperKey then FlipperDeactivate RightFlipper, RFPress
'
' If KeyCode = PlungerKey Then
'                Plunger.Fire
'                If BIPL = 1 Then
'                        SoundPlungerReleaseBall()                        'Plunger release sound when there is a ball in shooter lane
'                Else
'                        SoundPlungerReleaseNoBall()                        'Plunger release sound when there is no ball in shooter lane
'                End If
'        End If
'
'
'End Sub





'Dim plungerIM
' Set plungerIM = New cvpmImpulseP
' With plungerIM
'   .InitImpulseP ShooterLane, 46, 0.1
'   .Random 0.2
'   .Switch 78
'   .InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
'   .CreateEvents "plungerIM"
' End With

'**********************************************************************************************************
 ' Drain hole and kickers
Sub Drain_Hit:bsTrough.addball me : RandomSoundDrain Drain : End Sub
Sub sw77_Hit(): bsEjectHole.Addball me : SoundSaucerLock : End Sub
Sub sw77_UnHit(): bsEjectHole.Addball me : SoundSaucerKick 1, sw77 : End Sub
Sub BallRelease_UnHit(): RandomSoundBallRelease ballrelease : End Sub

'Bumpers
Sub Bumper1_Hit : vpmTimer.PulseSw(61) : RandomSoundBumperTop Bumper1: End Sub
Sub Bumper2_Hit : vpmTimer.PulseSw(62) : RandomSoundBumperMiddle bumper2: End Sub
Sub Bumper3_Hit : vpmTimer.PulseSw(63) : RandomSoundBumperBottom Bumper3: End Sub

'Wire Triggers
Sub SW15_Hit:Controller.Switch(15)=1 :  End Sub
Sub SW15_unHit:Controller.Switch(15)=0:End Sub
Sub SW16_Hit:Controller.Switch(16)=1 :  End Sub
Sub SW16_unHit:Controller.Switch(16)=0:End Sub
Sub SW17_Hit:Controller.Switch(17)=1 :  End Sub
Sub SW17_unHit:Controller.Switch(17)=0:End Sub
Sub SW18_Hit:Controller.Switch(18)=1 :  End Sub
Sub SW18_unHit:Controller.Switch(18)=0:End Sub
Sub SW25_Hit:Controller.Switch(25)=1 :  End Sub
Sub SW25_unHit:Controller.Switch(25)=0:End Sub
Sub SW26_Hit:Controller.Switch(26)=1 :  End Sub
Sub SW26_unHit:Controller.Switch(26)=0:End Sub
Sub SW27_Hit:Controller.Switch(27)=1 :  End Sub
Sub SW27_unHit:Controller.Switch(27)=0:End Sub
Sub SW28_Hit:Controller.Switch(28)=1 :  End Sub
Sub SW28_unHit:Controller.Switch(28)=0:End Sub
Sub SW71_Hit:Controller.Switch(71)=1 :  End Sub
Sub SW71_unHit:Controller.Switch(71)=0:End Sub
Sub SW72_Hit:Controller.Switch(72)=1 :  End Sub
Sub SW72_unHit:Controller.Switch(72)=0:End Sub
Sub SW73_Hit:Controller.Switch(73)=1 :  End Sub
Sub SW73_unHit:Controller.Switch(73)=0:End Sub

Sub SW78_Hit : Controller.Switch(78)=1 :  BIPL=1 : playsound"rollover" : End Sub
Sub SW78_unHit : Controller.Switch(78)=0:BIPL=0 :End Sub


Sub lock74_Hit:stopsound "fx_metalrolling" : playsound"WireRamp_Hit" : End Sub

'Stand Up Targets

Sub sw36_Hit
  TargetBouncer Activeball, 1
  vpmTimer.PulseSw 36
End Sub

Sub sw37_Hit
  TargetBouncer Activeball, 1
  vpmTimer.PulseSw 37
End Sub

Sub sw38_Hit
  TargetBouncer Activeball, 1
  vpmTimer.PulseSw 38
End Sub

Sub sw41_Hit
  TargetBouncer Activeball, 1
  vpmTimer.PulseSw 41
End Sub

Sub sw42_Hit
  TargetBouncer Activeball, 1
  vpmTimer.PulseSw 42
End Sub

Sub sw43_Hit
  TargetBouncer Activeball, 1
  vpmTimer.PulseSw 43
End Sub

Sub sw51_Hit
  TargetBouncer Activeball, 1
  vpmTimer.PulseSw 51
End Sub

Sub sw52_Hit
  TargetBouncer Activeball, 1
  vpmTimer.PulseSw 52
End Sub

Sub sw53_Hit
  TargetBouncer Activeball, 1
  vpmTimer.PulseSw 53
End Sub

Sub sw44_Hit
  TargetBouncer Activeball, 1
  vpmTimer.PulseSw 44
End Sub

Sub sw45_Hit
  TargetBouncer Activeball, 1
  vpmTimer.PulseSw 45
End Sub

Sub sw46_Hit
  TargetBouncer Activeball, 1
  vpmTimer.PulseSw 46
End Sub

Sub sw86_Hit
  TargetBouncer Activeball, 1
  vpmTimer.PulseSw 86
End Sub

Sub sw87_Hit
  TargetBouncer Activeball, 1
  vpmTimer.PulseSw 87
End Sub

Sub sw88_Hit
  TargetBouncer Activeball, 1
  vpmTimer.PulseSw 88
End Sub

 'Stand Up Targets
' Sub sw36_Hit: vpmTimer.pulseSw 36:End Sub
' Sub sw37_Hit: vpmTimer.pulseSw 37:End Sub
' Sub sw38_Hit: vpmTimer.pulseSw 38:End Sub
' Sub sw41_Hit: vpmTimer.pulseSw 41:End Sub
' Sub sw42_Hit: vpmTimer.pulseSw 42:End Sub
' Sub sw43_Hit: vpmTimer.pulseSw 43:End Sub
' Sub sw51_Hit: vpmTimer.pulseSw 51:End Sub
' Sub sw52_Hit: vpmTimer.pulseSw 52:End Sub
' Sub sw53_Hit: vpmTimer.pulseSw 53:End Sub
' Sub sw44_Hit: vpmTimer.pulseSw 44:End Sub
' Sub sw45_Hit: vpmTimer.pulseSw 45:End Sub
' Sub sw46_Hit: vpmTimer.pulseSw 46:End Sub
' Sub sw86_Hit: vpmTimer.pulseSw 86:End Sub
' Sub sw87_Hit: vpmTimer.pulseSw 87:End Sub
' Sub sw88_Hit: vpmTimer.pulseSw 88:End Sub

 'Ramp Triggers
Sub SW65_Hit:Controller.Switch(65)=1 : playsound"fx_metalrolling" : : playsound"WireRamp_Hit" :End Sub
Sub SW65_unHit:Controller.Switch(65)=0:End Sub
Sub SW67_Hit:Controller.Switch(67)=1 : playsound"rollover" : End Sub
Sub SW67_unHit:Controller.Switch(67)=0:End Sub
Sub SW84_Hit:Controller.Switch(84)=1 : playsound"rollover" : End Sub
Sub SW84_unHit:Controller.Switch(84)=0:End Sub

 'supercharger

Sub SW81_Hit
  Controller.Switch(81)=true
  'activeball.velY=2
  activeball.velX=activeball.velX+11
End Sub

Sub SW81_unHit
  Controller.Switch(81)=false
End Sub

Sub SW82_Hit
  Controller.Switch(82)=true
  'activeball.velY=3
  activeball.velX=activeball.velX+15
End Sub

Sub SW82_unHit
  Controller.Switch(82)=false
End Sub

Sub SW83_Hit
  Controller.Switch(83)=true
  'activeball.velY=5
  activeball.velX=activeball.velX+19
End Sub

Sub SW83_unHit
  Controller.Switch(83)=false
End Sub

Sub SW85_Hit
  vpmTimer.pulseSw 85
  'Controller.Switch(85)=true
  playsound"sc_loop2"
End Sub

Sub SW85_unHit
  'Controller.Switch(85)=false
End Sub

Sub Trigger11_Hit()
  Wall30.collidable=1
End Sub

Sub Wall30_hit()
  Wall30.collidable=0
  drop.enabled=0
  drop.enabled=1
End Sub

Sub Trigger12_Hit()
  Wall31.collidable=1
End Sub

Sub Wall31_hit()
  Wall31.collidable=0
  drop.enabled=0
  drop.enabled=1
End Sub

Sub Trigger9_Hit()
  Primitive29.transz=1
  wobble1.enabled=1
End Sub

'Sub Trigger10_Hit()
' supercharger_p.transx=-2
' primitive41.transx=-2
' dome118.transx=-2
' dome121.transx=-2
' dome123.transx=-2
' bolt12.transx=2
' bolt13.transx=2
' wobble1.enabled=1
'End Sub
'
'Sub wobble1_Timer
' Primitive29.transz=0
' supercharger_p.transx=0
' primitive41.transx=0
' dome118.transx=0
' dome121.transx=0
' dome123.transx=0
' bolt12.transx=0
' bolt13.transx=0
' wobble1.enabled=0
'End Sub

Sub Trigger1_Hit():activeball.velY=activeball.velY * 0.8:playSound "fx_metalrolling":End Sub

Sub Trigger2_Hit():stopSound "fx_shortmetal" :playSound "WireRamp_Hit" : End Sub

Sub Trigger3_Hit():stopSound "fx_metalrolling" :playSound "WireRamp_Hit" : End Sub

Sub Trigger4_Hit():playSound "fx_rr7":End Sub

Sub Trigger5_Hit():playSound "fx_rrenter":End Sub

Sub Trigger_slide_Hit():playSound "sc_loop2 1":End Sub

Sub Trigger6_Hit()
  If ramptrig=1 Then
    playSound "fx_rrenter"
  End If
End Sub

Sub Trigger7_Hit():playSound "fx_shortmetal":End Sub

Sub Trigger8_Hit():playSound "fx_rr7":End Sub

Sub Gate1_Hit():playSound "gate":End Sub

Sub Gate5_Hit():playSound "gate":End Sub

'Sub BallReleaseGate_Hit():playSound "gate":End Sub

Sub drop_Timer
  playSound "fx_ballrampdrop"
  drop.enabled=0
End Sub




'***************************************************************
'****  VPW DYNAMIC BALL SHADOWS by Iakki, Apophis, and Wylte
'***************************************************************

' *** These are usually defined elsewhere (ballrolling), but activate here if necessary
'Const tnob = 10 ' total number of balls
Const lob = 0 'locked balls on start; might need some fiddling depending on how your locked balls are done
'Dim tablewidth: tablewidth = Table1.width
'Dim tableheight: tableheight = Table1.height



' *** This segment goes within the RollingUpdate sub, so that if Ambient...=0 and Dynamic...=0 the entire DynamicBSUpdate sub can be skipped for max performance
' *** Change gBOT to BOT if using existing getballs code
' *** Includes lines commonly found there, for reference:
' ' stop the sound of deleted balls
' For b = UBound(gBOT) + 1 to tnob
'   If AmbientBallShadowOn = 0 Then BallShadowA(b).visible = 0
'   ...rolling(b) = False
'   ...StopSound("BallRoll_" & b)
' Next
'
' ...rolling and drop sounds...

'   If DropCount(b) < 5 Then
'     DropCount(b) = DropCount(b) + 1
'   End If
'
'   ' "Static" Ball Shadows
'   If AmbientBallShadowOn = 0 Then
'     If gBOT(b).Z > 30 Then
'       BallShadowA(b).height=gBOT(b).z - BallSize/4    'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping the ramp
'     Else
'       BallShadowA(b).height=gBOT(b).z - BallSize/2 + 5
'     End If
'     BallShadowA(b).Y = gBOT(b).Y + Ballsize/5 + offsetY
'     BallShadowA(b).X = gBOT(b).X + offsetX
'     BallShadowA(b).visible = 1
'   End If

' *** Required Functions, enable these if they are not already present elswhere in your table
Function max(a,b)
  if a > b then
    max = a
  Else
    max = b
  end if
end Function

'Function Distance(ax,ay,bx,by)
' Distance = SQR((ax - bx)^2 + (ay - by)^2)
'End Function

'Dim PI: PI = 4*Atn(1)

'Function Atn2(dy, dx)
' If dx > 0 Then
'   Atn2 = Atn(dy / dx)
' ElseIf dx < 0 Then
'   If dy = 0 Then
'     Atn2 = pi
'   Else
'     Atn2 = Sgn(dy) * (pi - Atn(Abs(dy / dx)))
'   end if
' ElseIf dx = 0 Then
'   if dy = 0 Then
'     Atn2 = 0
'   else
'     Atn2 = Sgn(dy) * pi / 2
'   end if
' End If
'End Function

'Function AnglePP(ax,ay,bx,by)
' AnglePP = Atn2((by - ay),(bx - ax))*180/PI
'End Function

'****** End Part B:  Code and Functions ******


'****** Part C:  The Magic ******

' *** These define the appearance of shadows in your table  ***

'Ambient (Room light source)
Const AmbientBSFactor     = 0.9 '0 to 1, higher is darker
Const AmbientMovement   = 2   '1 to 4, higher means more movement as the ball moves left and right
Const offsetX       = 0   'Offset x position under ball (These are if you want to change where the "room" light is for calculating the shadow position,)
Const offsetY       = 0   'Offset y position under ball  (for example 5,5 if the light is in the back left corner)
'Dynamic (Table light sources)
Const DynamicBSFactor     = 0.95  '0 to 1, higher is darker
Const Wideness        = 20  'Sets how wide the dynamic ball shadows can get (20 +5 thinness is technically most accurate for lights at z ~25 hitting a 50 unit ball)
Const Thinness        = 5   'Sets minimum as ball moves away from source

' ***                           ***

' *** Trim or extend these to *match* the number of balls/primitives/flashers on the table!
dim objrtx1(5), objrtx2(5)
dim objBallShadow(5)
Dim OnPF(5)
Dim BallShadowA
BallShadowA = Array (BallShadowA0,BallShadowA1,BallShadowA2,BallShadowA3,BallShadowA4)
Dim DSSources(30), numberofsources', DSGISide(30) 'Adapted for TZ with GI left / GI right

Dim ClearSurface:ClearSurface = True    'Variable for hiding flasher shadow on wire and clear plastic ramps
                  'Intention is to set this either globally or in a similar manner to RampRolling sounds

'Initialization
DynamicBSInit

sub DynamicBSInit()
  Dim iii, source

  for iii = 0 to tnob - 1               'Prepares the shadow objects before play begins
    Set objrtx1(iii) = Eval("RtxBallShadow" & iii)
    objrtx1(iii).material = "RtxBallShadow" & iii
    objrtx1(iii).z = 1 + iii/1000 + 0.01      'Separate z for layering without clipping
    objrtx1(iii).visible = 0

    Set objrtx2(iii) = Eval("RtxBall2Shadow" & iii)
    objrtx2(iii).material = "RtxBallShadow2_" & iii
    objrtx2(iii).z = 1 + iii/1000 + 0.02
    objrtx2(iii).visible = 0

    Set objBallShadow(iii) = Eval("BallShadow" & iii)
    objBallShadow(iii).material = "BallShadow" & iii
    UpdateMaterial objBallShadow(iii).material,1,0,0,0,0,0,AmbientBSFactor,RGB(0,0,0),0,0,False,True,0,0,0,0
    objBallShadow(iii).Z = 1 + iii/1000 + 0.04
    objBallShadow(iii).visible = 0

    BallShadowA(iii).Opacity = 100*AmbientBSFactor
    BallShadowA(iii).visible = 0
  Next

  iii = 0

  For Each Source in DynamicSources
    DSSources(iii) = Array(Source.x, Source.y)
'   If Instr(Source.name , "Left") > 0 Then DSGISide(iii) = 0 Else DSGISide(iii) = 1  'Adapted for TZ with GI left / GI right
    iii = iii + 1
  Next
  numberofsources = iii
end sub


Sub BallOnPlayfieldNow(yeh, num)    'Only update certain things once, save some cycles
  If yeh Then
    OnPF(num) = True
'   debug.print "Back on PF"
    UpdateMaterial objBallShadow(num).material,1,0,0,0,0,0,AmbientBSFactor,RGB(0,0,0),0,0,False,True,0,0,0,0
    objBallShadow(num).size_x = 5
    objBallShadow(num).size_y = 4.5
    objBallShadow(num).visible = 1
    BallShadowA(num).visible = 0
  Else
    OnPF(num) = False
'   debug.print "Leaving PF"
    If Not ClearSurface Then
      BallShadowA(num).visible = 1
      objBallShadow(num).visible = 0
    Else
      objBallShadow(num).visible = 1
    End If
  End If
End Sub

Sub DynamicBSUpdate
  Dim falloff: falloff = 150 'Max distance to light sources, can be changed dynamically if you have a reason
  Dim ShadowOpacity1, ShadowOpacity2
  Dim s, LSd, iii
  Dim dist1, dist2, src1, src2
  Dim BOT: BOT=getballs 'Uncomment if you're deleting balls - Don't do it! #SaveTheBalls

  'Hide shadow of deleted balls
  For s = UBound(BOT) + 1 to tnob - 1
    objrtx1(s).visible = 0
    objrtx2(s).visible = 0
    objBallShadow(s).visible = 0
    BallShadowA(s).visible = 0
  Next

  If UBound(BOT) < lob Then Exit Sub    'No balls in play, exit

'The Magic happens now
  For s = lob to UBound(BOT)

' *** Normal "ambient light" ball shadow
  'Layered from top to bottom. If you had an upper pf at for example 80 units and ramps even above that, your segments would be z>110; z<=110 And z>100; z<=100 And z>30; z<=30 And z>20; Else invisible

    If AmbientBallShadowOn = 1 Then     'Primitive shadow on playfield, flasher shadow in ramps
      If BOT(s).Z > 30 Then             'The flasher follows the ball up ramps while the primitive is on the pf
        If OnPF(s) Then BallOnPlayfieldNow False, s   'One-time update

        If Not ClearSurface Then              'Don't show this shadow on plastic or wire ramps (table-wide variable, for now)
          BallShadowA(s).X = BOT(s).X + offsetX
          BallShadowA(s).Y = BOT(s).Y + BallSize/5
          BallShadowA(s).height=BOT(s).z - BallSize/4   'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping the ramp
        Else
          If BOT(s).X < tablewidth/2 Then
            objBallShadow(s).X = ((BOT(s).X) - (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) + offsetX + 5
          Else
            objBallShadow(s).X = ((BOT(s).X) + (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) + offsetX - 5
          End If
          objBallShadow(s).Y = BOT(s).Y + BallSize/10 + offsetY
          objBallShadow(s).size_x = 5 * ((BOT(s).Z+BallSize)/80)      'Shadow gets larger and more diffuse as it moves up
          objBallShadow(s).size_y = 4.5 * ((BOT(s).Z+BallSize)/80)
          UpdateMaterial objBallShadow(s).material,1,0,0,0,0,0,AmbientBSFactor*(30/(BOT(s).Z)),RGB(0,0,0),0,0,False,True,0,0,0,0
        End If

      Elseif BOT(s).Z <= 30 And BOT(s).Z > 20 Then  'On pf, primitive only
        If Not OnPF(s) Then BallOnPlayfieldNow True, s

        If BOT(s).X < tablewidth/2 Then
          objBallShadow(s).X = ((BOT(s).X) - (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) + offsetX + 5
        Else
          objBallShadow(s).X = ((BOT(s).X) + (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) + offsetX - 5
        End If
        objBallShadow(s).Y = BOT(s).Y + offsetY
'       objBallShadow(s).Z = BOT(s).Z + s/1000 + 0.04   'Uncomment (and adjust If/Elseif height logic) if you want the primitive shadow on an upper/split pf

      Else                        'Under pf, no shadows
        objBallShadow(s).visible = 0
        BallShadowA(s).visible = 0
      end if

    Elseif AmbientBallShadowOn = 2 Then   'Flasher shadow everywhere
      If BOT(s).Z > 30 Then             'In a ramp
        If Not ClearSurface Then              'Don't show this shadow on plastic or wire ramps (table-wide variable, for now)
          BallShadowA(s).X = BOT(s).X + offsetX
          BallShadowA(s).Y = BOT(s).Y + BallSize/5
          BallShadowA(s).height=BOT(s).z - BallSize/4   'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping the ramp
        Else
          BallShadowA(s).X = BOT(s).X + offsetX
          BallShadowA(s).Y = BOT(s).Y + offsetY
        End If
      Elseif BOT(s).Z <= 30 And BOT(s).Z > 20 Then  'On pf
        BallShadowA(s).visible = 1
        If BOT(s).X < tablewidth/2 Then
          BallShadowA(s).X = ((BOT(s).X) - (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) + offsetX + 5
        Else
          BallShadowA(s).X = ((BOT(s).X) + (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) + offsetX - 5
        End If
        BallShadowA(s).Y = BOT(s).Y + offsetY
        BallShadowA(s).height=0.1
      Else                      'Under pf
        BallShadowA(s).visible = 0
      End If
    End If

' *** Dynamic shadows
    If DynamicBallShadowsOn Then
      If BOT(s).Z < 30 And BOT(s).X < 850 Then  'Parameters for where the shadows can show, here they are not visible above the table (no upper pf) or in the plunger lane
        dist1 = falloff:
        dist2 = falloff
        For iii = 0 to numberofsources - 1 ' Search the 2 nearest influencing lights
          LSd = Distance(BOT(s).x, BOT(s).y, DSSources(iii)(0), DSSources(iii)(1)) 'Calculating the Linear distance to the Source
          If LSd < falloff And gilvl > 0.5 Then
'         If LSd < dist2 And ((DSGISide(iii) = 0 And gilvl>0.5) Or (DSGISide(iii) = 1 And gilvl>0.5)) Then  'Adapted for TZ with GI left / GI right
            dist2 = dist1
            dist1 = LSd
            src2 = src1
            src1 = iii
          End If
        Next
        ShadowOpacity1 = 0
        If dist1 < falloff Then
          objrtx1(s).visible = 1 : objrtx1(s).X = BOT(s).X : objrtx1(s).Y = BOT(s).Y
          'objrtx1(s).Z = BOT(s).Z - 25 + s/1000 + 0.01 'Uncomment if you want to add shadows to an upper/lower pf
          objrtx1(s).rotz = AnglePP(DSSources(src1)(0), DSSources(src1)(1), BOT(s).X, BOT(s).Y) + 90
          ShadowOpacity1 = 1 - dist1 / falloff
          objrtx1(s).size_y = Wideness * ShadowOpacity1 + Thinness
          UpdateMaterial objrtx1(s).material,1,0,0,0,0,0,ShadowOpacity1*DynamicBSFactor^3,RGB(0,0,0),0,0,False,True,0,0,0,0
        Else
          objrtx1(s).visible = 0
        End If
        ShadowOpacity2 = 0
        If dist2 < falloff Then
          objrtx2(s).visible = 1 : objrtx2(s).X = BOT(s).X : objrtx2(s).Y = BOT(s).Y + offsetY
          'objrtx2(s).Z = BOT(s).Z - 25 + s/1000 + 0.02 'Uncomment if you want to add shadows to an upper/lower pf
          objrtx2(s).rotz = AnglePP(DSSources(src2)(0), DSSources(src2)(1), BOT(s).X, BOT(s).Y) + 90
          ShadowOpacity2 = 1 - dist2 / falloff
          objrtx2(s).size_y = Wideness * ShadowOpacity2 + Thinness
          UpdateMaterial objrtx2(s).material,1,0,0,0,0,0,ShadowOpacity2*DynamicBSFactor^3,RGB(0,0,0),0,0,False,True,0,0,0,0
        Else
          objrtx2(s).visible = 0
        End If
        If AmbientBallShadowOn = 1 Then
          'Fades the ambient shadow (primitive only) when it's close to a light
          UpdateMaterial objBallShadow(s).material,1,0,0,0,0,0,AmbientBSFactor*(1 - max(ShadowOpacity1, ShadowOpacity2)),RGB(0,0,0),0,0,False,True,0,0,0,0
        Else
          BallShadowA(s).Opacity = 100 * AmbientBSFactor * (1 - max(ShadowOpacity1, ShadowOpacity2))
        End If
      Else 'Hide dynamic shadows everywhere else, just in case
        objrtx2(s).visible = 0 : objrtx1(s).visible = 0
      End If
    End If
  Next
End Sub
'****************************************************************
'****  END VPW DYNAMIC BALL SHADOWS by Iakki, Apophis, and Wylte
'****************************************************************



'******************************************************
'*****   FLUPPER DOMES
'******************************************************


Sub FlashSolMod117(flvl)
  ObjLevel(4) = flvl
  UpdateFlasher 4
  'F117.IntensityScale = flvl
  F117a.IntensityScale = flvl
  F117b.IntensityScale = flvl
  F117c.IntensityScale = flvl
  F117d.IntensityScale = flvl
  F117e.IntensityScale = flvl
End Sub

Sub FlashSolMod118(flvl)
  ObjLevel(2) = flvl
  UpdateFlasher 2
  F118.IntensityScale = flvl
  F118a.IntensityScale = flvl
  F118b.IntensityScale = flvl
End Sub

Sub FlashSolMod119(flvl)
  ObjLevel(8) = flvl
  UpdateFlasher 8
  F119.IntensityScale = flvl
  F119a.IntensityScale = flvl
  F119b.IntensityScale = flvl
End Sub

Sub FlashSolMod120(flvl)
  F120.State = flvl
  F120a.State = flvl
End Sub

Sub FlashSolMod121(flvl)
  ObjLevel(1) = flvl
  UpdateFlasher 1
  ObjLevel(6) = flvl
  UpdateFlasher 6
' F121.IntensityScale = flvl
  F121a.IntensityScale = flvl
  F121b.IntensityScale = flvl
' F121c.IntensityScale = flvl
  F121d.IntensityScale = flvl
  F121e.IntensityScale = flvl
  F121f.IntensityScale = flvl
  F121g.IntensityScale = flvl
  F121h.IntensityScale = flvl
  F121i.IntensityScale = flvl
  F121j.IntensityScale = flvl
' F121k.IntensityScale = flvl
End Sub

Sub FlashSolMod122(flvl)
  ObjLevel(5) = flvl
  UpdateFlasher 5
  F122.IntensityScale = flvl
' F122a.IntensityScale = flvl
  f122b.IntensityScale = flvl
  F122c.IntensityScale = flvl
  F122d.IntensityScale = flvl
End Sub

Sub FlashSolMod123(flvl)
  ObjLevel(3) = flvl
  UpdateFlasher 3
  ObjLevel(7) = flvl
  UpdateFlasher 7
' F123.IntensityScale = flvl
  F123a.IntensityScale = flvl
' F123b.IntensityScale = flvl
  F123c.IntensityScale = flvl
  F123d.IntensityScale = flvl
  F123e.IntensityScale = flvl
  F123f.IntensityScale = flvl
  F123g.IntensityScale = flvl
  F123i.IntensityScale = flvl
  F123j.IntensityScale = flvl
' F123k.IntensityScale = flvl
End Sub

Sub FlashSolMod124(flvl)
  ObjLevel(9) = flvl
  UpdateFlasher 9
  F124.IntensityScale = flvl
  F124a.IntensityScale = flvl
  F124b.IntensityScale = flvl
End Sub

Dim TestFlashers, TableRef, FlasherLightIntensity, FlasherFlareIntensity, FlasherBloomIntensity, FlasherOffBrightness

                ' *********************************************************************
TestFlashers = 0        ' *** set this to 1 to check position of flasher object       ***
Set TableRef = Table1     ' *** change this, if your table has another name             ***
FlasherLightIntensity = 0.3   ' *** lower this, if the VPX lights are too bright (i.e. 0.1)   ***
FlasherFlareIntensity = 0.1   ' *** lower this, if the flares are too bright (i.e. 0.1)     ***
FlasherBloomIntensity = 0.3   ' *** lower this, if the blooms are too bright (i.e. 0.1)     ***
FlasherOffBrightness = 0.3    ' *** brightness of the flasher dome when switched off (range 0-2)  ***
                ' *********************************************************************

Dim ObjLevel(20), objbase(20), objlit(20), objflasher(20), objbloom(20), objlight(20)

InitFlasher 1, "red"
InitFlasher 2, "red"
InitFlasher 3, "red"

InitFlasher 4, "white"
InitFlasher 5, "white"
InitFlasher 6, "white"
InitFlasher 7, "white"
InitFlasher 8, "white"
InitFlasher 9, "white"

'rotate the flasher with the command below (first argument = flasher nr, second argument = angle in degrees)
'RotateFlasher 1,17 : RotateFlasher 2,0 : RotateFlasher 3,90 : RotateFlasher 4,90

Sub InitFlasher(nr, col)
  ' store all objects in an array for use in FlashFlasher subroutine
  Set objbase(nr) = Eval("Flasherbase" & nr) : Set objlit(nr) = Eval("Flasherlit" & nr)
  Set objflasher(nr) = Eval("Flasherflash" & nr) : Set objlight(nr) = Eval("Flasherlight" & nr)
  Set objbloom(nr) = Eval("Flasherbloom" & nr)
  ' If the flasher is parallel to the playfield, rotate the VPX flasher object for POV and place it at the correct height
  If objbase(nr).RotY = 0 Then
    objbase(nr).ObjRotZ =  atn( (tablewidth/2 - objbase(nr).x) / (objbase(nr).y - tableheight*1.1)) * 180 / 3.14159
    objflasher(nr).RotZ = objbase(nr).ObjRotZ : objflasher(nr).height = objbase(nr).z + 60
  End If
  ' set all effects to invisible and move the lit primitive at the same position and rotation as the base primitive
  objlight(nr).IntensityScale = 0 : objlit(nr).visible = 0 : objlit(nr).material = "Flashermaterial" & nr
  objlit(nr).RotX = objbase(nr).RotX : objlit(nr).RotY = objbase(nr).RotY : objlit(nr).RotZ = objbase(nr).RotZ
  objlit(nr).ObjRotX = objbase(nr).ObjRotX : objlit(nr).ObjRotY = objbase(nr).ObjRotY : objlit(nr).ObjRotZ = objbase(nr).ObjRotZ
  objlit(nr).x = objbase(nr).x : objlit(nr).y = objbase(nr).y : objlit(nr).z = objbase(nr).z
  objbase(nr).BlendDisableLighting = FlasherOffBrightness
  ' set the texture and color of all objects
  select case objbase(nr).image
    Case "dome2basewhite" : objbase(nr).image = "dome2base" & col : objlit(nr).image = "dome2lit" & col :
    Case "ronddomebasewhite" : objbase(nr).image = "ronddomebase" & col : objlit(nr).image = "ronddomelit" & col
    Case "domeearbasewhite" : objbase(nr).image = "domeearbase" & col : objlit(nr).image = "domeearlit" & col
  end select
  If TestFlashers = 0 Then objflasher(nr).imageA = "domeflashwhite" : objflasher(nr).visible = 0 : End If
  select case col
    Case "red" :    objlight(nr).color = RGB(255,32,4) : objflasher(nr).color = RGB(255,32,4) : objbloom(nr).color = RGB(255,32,4)
    Case "white" :  objlight(nr).color = RGB(255,240,150) : objflasher(nr).color = RGB(100,86,59) : objbloom(nr).color = RGB(255,240,150)
  end select
  objlight(nr).colorfull = objlight(nr).color
  If TableRef.ShowDT and ObjFlasher(nr).RotX = -45 Then
    objflasher(nr).height = objflasher(nr).height - 20 * ObjFlasher(nr).y / tableheight
    ObjFlasher(nr).y = ObjFlasher(nr).y + 10
  End If
  objflasher(nr).visible = 1
  objbloom(nr).visible = 1
  objlit(nr).visible = 1
  UpdateFlasher nr
End Sub

Sub RotateFlasher(nr, angle) : angle = ((angle + 360 - objbase(nr).ObjRotZ) mod 180)/30 : objbase(nr).showframe(angle) : objlit(nr).showframe(angle) : End Sub

Sub UpdateFlasher(nr)
  If VRRoom > 0 Then
    Select Case nr
      case 1:   'solenoir 21
        VRBGFL21_1.opacity = 75 * ObjLevel(nr)
        VRBGFL21_2.opacity = 75 * ObjLevel(nr)
        VRBGFL21_3.opacity = 75 * ObjLevel(nr)
        VRBGFL21_4.opacity = 110 * ObjLevel(nr)^(3/2.5)
      case 2:   'solenoid 18
        VRBGFL18_1.opacity = 75 * ObjLevel(nr)
        VRBGFL18_2.opacity = 75 * ObjLevel(nr)
        VRBGFL18_3.opacity = 75 * ObjLevel(nr)
        VRBGFL18_4.opacity = 110 * ObjLevel(nr)^(3/2.5)
      case 3:   'solenoid 23
        VRBGFL23_1.opacity = 75 * ObjLevel(nr)
        VRBGFL23_2.opacity = 75 * ObjLevel(nr)
        VRBGFL23_3.opacity = 75 * ObjLevel(nr)
        VRBGFL23_4.opacity = 110 * ObjLevel(nr)^(3/2.5)
      case 8:
        VRBGFL19_1.opacity = 100 * ObjLevel(nr)
        VRBGFL19_2.opacity = 100 * ObjLevel(nr)
        VRBGFL19_3.opacity = 100 * ObjLevel(nr)
        VRBGFL19_4.opacity = 105 * ObjLevel(nr)^(3/2.5)
      case 9:
        VRBGFL24_1.opacity = 100 * ObjLevel(nr)
        VRBGFL24_2.opacity = 100 * ObjLevel(nr)
        VRBGFL24_3.opacity = 100 * ObjLevel(nr)
        VRBGFL24_4.opacity = 105 * ObjLevel(nr)^(3/2.5)
    End Select
  End If
  objflasher(nr).opacity = 1000 *  FlasherFlareIntensity * ObjLevel(nr)
  objbloom(nr).opacity = 100 *  FlasherBloomIntensity * ObjLevel(nr)
  objlight(nr).IntensityScale = 0.5 * FlasherLightIntensity * ObjLevel(nr)^(3/2.5)
  objbase(nr).BlendDisableLighting =  FlasherOffBrightness + 10 * ObjLevel(nr)^(3/2.5)
  objlit(nr).BlendDisableLighting = 100 * ObjLevel(nr)^(2/2.5)
  UpdateMaterial "Flashermaterial" & nr,0,0,0,0,0,0,ObjLevel(nr),RGB(255,255,255),0,0,False,True,0,0,0,0
End Sub

'******************************************************
'******  END FLUPPER DOMES
'******************************************************




'**********************************************************************************************************
'**********************************************************************************************************
' Start of VPX functions
'**********************************************************************************************************
'**********************************************************************************************************

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
  vpmTimer.pulseSw 32
  RS.VelocityCorrect(ActiveBall)
    RandomSoundSlingshotRight SLING1
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  vpmTimer.pulseSw 31
  LS.VelocityCorrect(ActiveBall)
    RandomSoundSlingshotLeft SLING2
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



'******************************************************
'                BALL ROLLING AND DROP SOUNDS
'******************************************************

Const tnob = 3 ' total number of balls
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

Sub RollingUpdate()
  Dim b, BOT
  BOT = GetBalls

  ' stop the sound of deleted balls
  For b = UBound(BOT) + 1 to tnob - 1
    ' Comment the next line if you are not implementing Dyanmic Ball Shadows
    If AmbientBallShadowOn = 0 Then BallShadowA(b).visible = 0
    rolling(b) = False
    StopSound("BallRoll_" & b)
  Next

  ' exit the sub if no balls on the table
  If UBound(BOT) = -1 Then Exit Sub

  ' play the rolling sound for each ball

  For b = 0 to UBound(BOT)
    If BallVel(BOT(b)) > 1 AND BOT(b).z < 30 Then
      rolling(b) = True
      PlaySound ("BallRoll_" & b), -1, VolPlayfieldRoll(BOT(b)) * BallRollVolume * VolumeDial, AudioPan(BOT(b)), 0, PitchPlayfieldRoll(BOT(b)), 1, 0, AudioFade(BOT(b))

    Else
      If rolling(b) = True Then
        StopSound("BallRoll_" & b)
        rolling(b) = False
      End If
    End If

    ' Ball Drop Sounds
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

    ' "Static" Ball Shadows
    ' Comment the next If block, if you are not implementing the Dyanmic Ball Shadows
    If AmbientBallShadowOn = 0 Then
      If BOT(b).Z > 30 Then
        BallShadowA(b).height=BOT(b).z - BallSize/4   'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping the ramp
      Else
        BallShadowA(b).height=0.1
      End If
      BallShadowA(b).Y = BOT(b).Y + offsetY
      BallShadowA(b).X = BOT(b).X + offsetX
      BallShadowA(b).visible = 1
    End If
  Next
End Sub

'Sub Targets_Hit (idx)
' PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
'End Sub
'
'Sub Spinner_Spin
' PlaySound "fx_spinner",0,.25,0,0.25
'End Sub


Sub bounce_timer()
    PlaySound "fx_bounce"
    bounce.enabled=0
End Sub





'////////////////////////////  MECHANICAL SOUNDS  ///////////////////////////
'//  This part in the script is an entire block that is dedicated to the physics sound system.
'//  Various scripts and sounds that may be pretty generic and could suit other WPC systems, but the most are tailored specifically for this table.


'///////////////////////////////  SOUNDS PARAMETERS  //////////////////////////////
Dim GlobalSoundLevel, CoinSoundLevel, PlungerReleaseSoundLevel, PlungerPullSoundLevel, NudgeLeftSoundLevel
Dim NudgeRightSoundLevel, NudgeCenterSoundLevel, StartButtonSoundLevel, RollingSoundFactor

CoinSoundLevel = 1                                                                                                                'volume level; range [0, 1]
NudgeLeftSoundLevel = 1                                                                                                        'volume level; range [0, 1]
NudgeRightSoundLevel = 1                                                                                                'volume level; range [0, 1]
NudgeCenterSoundLevel = 1                                                                                                'volume level; range [0, 1]
StartButtonSoundLevel = 0.1                                                                                                'volume level; range [0, 1]
PlungerReleaseSoundLevel = 0.8 '1 wjr                                                                                        'volume level; range [0, 1]
PlungerPullSoundLevel = 1                                                                                                'volume level; range [0, 1]
RollingSoundFactor = 1.1/5

'///////////////////////-----Solenoids, Kickers and Flash Relays-----///////////////////////
Dim FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel, FlipperUpAttackLeftSoundLevel, FlipperUpAttackRightSoundLevel
Dim FlipperUpSoundLevel, FlipperDownSoundLevel, FlipperLeftHitParm, FlipperRightHitParm
Dim SlingshotSoundLevel, BumperSoundFactor, KnockerSoundLevel

FlipperUpAttackMinimumSoundLevel = 0.050                                                           'volume level; range [0, 1]
FlipperUpAttackMaximumSoundLevel = 0.635                                                                'volume level; range [0, 1]
FlipperUpSoundLevel = 1.0                                                                        'volume level; range [0, 1]
FlipperDownSoundLevel = 0.45                                                                      'volume level; range [0, 1]
FlipperLeftHitParm = FlipperUpSoundLevel                                                                'sound helper; not configurable
FlipperRightHitParm = FlipperUpSoundLevel                                                                'sound helper; not configurable
SlingshotSoundLevel = 1                                                                                                'volume level; range [0, 1]
BumperSoundFactor = 7.25                                                                                                'volume multiplier; must not be zero
KnockerSoundLevel = 1                                                                                                         'volume level; range [0, 1]

'///////////////////////-----Ball Drops, Bumps and Collisions-----///////////////////////
Dim RubberStrongSoundFactor, RubberWeakSoundFactor, RubberFlipperSoundFactor,BallWithBallCollisionSoundFactor
Dim BallBouncePlayfieldSoftFactor, BallBouncePlayfieldHardFactor, PlasticRampDropToPlayfieldSoundLevel, WireRampDropToPlayfieldSoundLevel, DelayedBallDropOnPlayfieldSoundLevel
Dim WallImpactSoundFactor, MetalImpactSoundFactor, SubwaySoundLevel, SubwayEntrySoundLevel, ScoopEntrySoundLevel
Dim SaucerLockSoundLevel, SaucerKickSoundLevel

BallWithBallCollisionSoundFactor = 3.2                                                                        'volume multiplier; must not be zero
RubberStrongSoundFactor = 0.055/5                                                                                        'volume multiplier; must not be zero
RubberWeakSoundFactor = 0.075/5                                                                                        'volume multiplier; must not be zero
RubberFlipperSoundFactor = 0.2/5                                                                                'volume multiplier; must not be zero
BallBouncePlayfieldSoftFactor = 0.025                                                                        'volume multiplier; must not be zero
BallBouncePlayfieldHardFactor = 0.025                                                                        'volume multiplier; must not be zero
DelayedBallDropOnPlayfieldSoundLevel = 0.8                                                                        'volume level; range [0, 1]
WallImpactSoundFactor = 0.075                                                                                        'volume multiplier; must not be zero
MetalImpactSoundFactor = 0.075/3
SaucerLockSoundLevel = 0.8
SaucerKickSoundLevel = 0.8

'///////////////////////-----Gates, Spinners, Rollovers and Targets-----///////////////////////

Dim GateSoundLevel, TargetSoundFactor, SpinnerSoundLevel, RolloverSoundLevel, DTSoundLevel

GateSoundLevel = 0.5/5                                                                                                        'volume level; range [0, 1]
TargetSoundFactor = 0.0025 * 10                                                                                        'volume multiplier; must not be zero
DTSoundLevel = 0.25                                                                                                                'volume multiplier; must not be zero
RolloverSoundLevel = 0.25                                                                      'volume level; range [0, 1]
SpinnerSoundLevel = 0.5                                                                      'volume level; range [0, 1]

'///////////////////////-----Ball Release, Guides and Drain-----///////////////////////
Dim DrainSoundLevel, BallReleaseSoundLevel, BottomArchBallGuideSoundFactor, FlipperBallGuideSoundFactor

DrainSoundLevel = 0.8                                                                                                                'volume level; range [0, 1]
BallReleaseSoundLevel = 1                                                                                                'volume level; range [0, 1]
BottomArchBallGuideSoundFactor = 0.2                                                                        'volume multiplier; must not be zero
FlipperBallGuideSoundFactor = 0.015                                                                                'volume multiplier; must not be zero

'///////////////////////-----Loops and Lanes-----///////////////////////
Dim ArchSoundFactor
ArchSoundFactor = 0.025/5                                                                                                        'volume multiplier; must not be zero


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


'******************************************************
'  Fleep  Supporting Ball & Sound Functions
'******************************************************

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
    tmp = tableobj.y * 2 / tableheight-1

        if tmp > 7000 Then
                tmp = 7000
        elseif tmp < -7000 Then
                tmp = -7000
        end if

    If tmp > 0 Then
                AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = tableobj.x * 2 / tablewidth-1

        if tmp > 7000 Then
                tmp = 7000
        elseif tmp < -7000 Then
                tmp = -7000
        end if

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
        VolPlayfieldRoll = RollingSoundFactor * 0.001 * Csng(BallVel(ball) ^3)
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
        PlaySound ("Nudge_" & Int(Rnd*2)+1), 0, NudgeLeftSoundLevel * VolumeDial, -0.1, 0.25
End Sub

Sub SoundNudgeRight()
        PlaySound ("Nudge_" & Int(Rnd*2)+1), 0, NudgeRightSoundLevel * VolumeDial, 0.1, 0.25
End Sub

Sub SoundNudgeCenter()
        PlaySound ("Nudge_" & Int(Rnd*2)+1), 0, NudgeCenterSoundLevel * VolumeDial, 0, 0.25
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
        PlaySoundAtLevelStatic ("Drain_" & Int(Rnd*11)+1), DrainSoundLevel, drainswitch
End Sub

'/////////////////////////////  TROUGH BALL RELEASE SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundBallRelease(drainswitch)
        PlaySoundAtLevelStatic SoundFX("BallRelease" & Int(Rnd*7)+1,DOFContactors), BallReleaseSoundLevel, drainswitch
End Sub

'/////////////////////////////  SLINGSHOT SOLENOID SOUNDS  ////////////////////////////
Sub RandomSoundSlingshotLeft(sling)
        PlaySoundAtLevelStatic SoundFX("Sling_L" & Int(Rnd*27)+1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

Sub RandomSoundSlingshotRight(sling)
        PlaySoundAtLevelStatic SoundFX("Sling_R" & Int(Rnd*26)+1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

'/////////////////////////////  BUMPER SOLENOID SOUNDS  ////////////////////////////
Sub RandomSoundBumperTop(Bump)
        PlaySoundAtLevelStatic SoundFX("Bumpers_Top_" & Int(Rnd*23)+1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperMiddle(Bump)
        PlaySoundAtLevelStatic SoundFX("Bumpers_Middle_" & Int(Rnd*26)+1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperBottom(Bump)
        PlaySoundAtLevelStatic SoundFX("Bumpers_Bottom_" & Int(Rnd*29)+1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

'/////////////////////////////  SPINNER SOUNDS  ////////////////////////////
Sub SoundSpinner(spinnerswitch)
        PlaySoundAtLevelStatic ("Spinner"), SpinnerSoundLevel, spinnerswitch
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
        PlaySoundAtLevelStatic SoundFX("Flipper_L0" & Int(Rnd*9)+1,DOFFlippers), FlipperLeftHitParm, Flipper
End Sub

Sub RandomSoundFlipperUpRight(flipper)
        PlaySoundAtLevelStatic SoundFX("Flipper_R0" & Int(Rnd*9)+1,DOFFlippers), FlipperRightHitParm, Flipper
End Sub

Sub RandomSoundReflipUpLeft(flipper)
        PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_L0" & Int(Rnd*3)+1,DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundReflipUpRight(flipper)
        PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_R0" & Int(Rnd*3)+1,DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperDownLeft(flipper)
        PlaySoundAtLevelStatic SoundFX("Flipper_Left_Down_" & Int(Rnd*7)+1,DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperDownRight(flipper)
        PlaySoundAtLevelStatic SoundFX("Flipper_Right_Down_" & Int(Rnd*8)+1,DOFFlippers), FlipperDownSoundLevel, Flipper
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
        PlaySoundAtLevelActiveBall ("Flipper_Rubber_" & Int(Rnd*7)+1), parm  * RubberFlipperSoundFactor
End Sub

'/////////////////////////////  ROLLOVER SOUNDS  ////////////////////////////
Sub RandomSoundRollover()
        PlaySoundAtLevelActiveBall ("Rollover_" & Int(Rnd*4)+1), RolloverSoundLevel
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
        PlaySoundAtLevelActiveBall ("Rubber_" & Int(Rnd*9)+1), Vol(ActiveBall) * RubberWeakSoundFactor
End Sub

'/////////////////////////////  WALL IMPACTS  ////////////////////////////
Sub Walls_Hit(idx)
        RandomSoundWall()
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
        PlaySoundAtLevelActiveBall ("Metal_Touch_" & Int(Rnd*13)+1), Vol(ActiveBall) * MetalImpactSoundFactor
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
                PlaySoundAtLevelActiveBall ("Apron_Bounce_"& Int(Rnd*2)+1), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
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
        PlaySoundAtLevelActiveBall ("Apron_Hard_Hit_" & Int(Rnd*3)+1), BottomArchBallGuideSoundFactor * 0.25
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
                PlaySoundAtLevelActiveBall ("Apron_Medium_" & Int(Rnd*3)+1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
        End If
        If finalspeed < 6 Then
                PlaySoundAtLevelActiveBall ("Apron_Soft_" & Int(Rnd*7)+1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
        End If
End Sub

'/////////////////////////////  TARGET HIT SOUNDS  ////////////////////////////
Sub RandomSoundTargetHitStrong()
        PlaySoundAtLevelActiveBall SoundFX("Target_Hit_" & Int(Rnd*4)+5,DOFTargets), Vol(ActiveBall) * 0.45 * TargetSoundFactor
End Sub

Sub RandomSoundTargetHitWeak()
        PlaySoundAtLevelActiveBall SoundFX("Target_Hit_" & Int(Rnd*4)+1,DOFTargets), Vol(ActiveBall) * TargetSoundFactor
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
        PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_" & Int(Rnd*7)+1), volz(aBall) * BallBouncePlayfieldHardFactor, aBall
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
        PlaySoundAtLevelStatic ("Gate_FastTrigger_" & Int(Rnd*2)+1), GateSoundLevel, Activeball
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
        PlaySoundAtLevelActiveBall ("Arch_L" & Int(Rnd*4)+1), Vol(ActiveBall) * ArchSoundFactor
End Sub

Sub RandomSoundRightArch()
        PlaySoundAtLevelActiveBall ("Arch_R" & Int(Rnd*4)+1), Vol(ActiveBall) * ArchSoundFactor
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
        PlaySoundAtLevelStatic ("Saucer_Enter_" & Int(Rnd*2)+1), SaucerLockSoundLevel, Activeball
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


'///////////////////////////  DROP TARGET HIT SOUNDS  ///////////////////////////

Sub RandomSoundDropTargetReset(obj)
        PlaySoundAtLevelStatic SoundFX("Drop_Target_Reset_" & Int(Rnd*6)+1,DOFContactors), 1, obj
End Sub

Sub SoundDropTargetDrop(obj)
        PlaySoundAtLevelStatic ("Drop_Target_Down_" & Int(Rnd*6)+1), 200, obj
End Sub

'/////////////////////////////  GI AND FLASHER RELAYS  ////////////////////////////

Const RelayFlashSoundLevel = 0.315                                                                        'volume level; range [0, 1];
Const RelayGISoundLevel = 1.05                                                                        'volume level; range [0, 1];

Sub Sound_GI_Relay(toggle, obj)
        Select Case toggle
                Case 1
                        PlaySoundAtLevelStatic ("Relay_GI_On"), 0.025*RelayGISoundLevel, obj
                Case 0
                        PlaySoundAtLevelStatic ("Relay_GI_Off"), 0.025*RelayGISoundLevel, obj
        End Select
End Sub

Sub Sound_Flash_Relay(toggle, obj)
        Select Case toggle
                Case 1
                        PlaySoundAtLevelStatic ("Relay_Flash_On"), 0.025*RelayFlashSoundLevel, obj
                Case 0
                        PlaySoundAtLevelStatic ("Relay_Flash_Off"), 0.025*RelayFlashSoundLevel, obj
        End Select
End Sub

'/////////////////////////////////////////////////////////////////
'                                        End Mechanical Sounds
'/////////////////////////////////////////////////////////////////



'******************************************************
'                TRACK ALL BALL VELOCITIES
'                 FOR RUBBER DAMPENER AND DROP TARGETS
'******************************************************

dim cor : set cor = New CoRTracker

Class CoRTracker
  public ballvel, ballvelx, ballvely

  Private Sub Class_Initialize : redim ballvel(0) : redim ballvelx(0): redim ballvely(0) : End Sub

  Public Sub Update()        'tracks in-ball-velocity
    dim str, b, AllBalls, highestID : allBalls = getballs
    for each b in allballs
      if b.id >= HighestID then highestID = b.id
    Next
    if uBound(ballvel) < highestID then redim ballvel(highestID)        'set bounds
    if uBound(ballvelx) < highestID then redim ballvelx(highestID)        'set bounds
    if uBound(ballvely) < highestID then redim ballvely(highestID)        'set bounds
    for each b in allballs
      ballvel(b.id) = BallSpeed(b)
      ballvelx(b.id) = b.velx
      ballvely(b.id) = b.vely
    Next
  End Sub
End Class




'*******VR Code (Rawd)*********

BeaconREF1.opacity = 0
BeaconREF2.opacity = 150

Dim BeaconRefOn
BeaconRefOn = true
Dim BeaconPos:BeaconPos = 0

Sub BeaconTimer_Timer ' DJRobX code....
  BeaconPos = BeaconPos + 3

  if BeaconPos = 180 then BeaconRefOn = true  ' turns on beacon reflection timing..

  if BeaconPos = 360 then BeaconPos = 0
  BeaconRedInt.RotY = BeaconPos
    BeaconRed.BlendDisableLighting=.5 * abs(sin((BeaconPos+90) * 6.28 / 360))
  BeaconFR.RotY = BeaconPos

  if BeaconRefOn = true then
    BeaconREF1b.opacity = 2

    if BeaconREF1.opacity < 150 then BeaconREF1.opacity = BeaconREF1.opacity +3
    if BeaconREF1.opacity > 10 then BeaconREF1b.opacity = 5
    if BeaconREF1.opacity > 50 then BeaconREF1b.opacity = 10
    if BeaconREF1.opacity > 75 then BeaconREF1b.opacity = 7
    if BeaconREF1.opacity > 100 then BeaconREF1b.opacity = 15
    if BeaconREF1.opacity > 125 then BeaconREF1b.opacity = 17
    if BeaconREF1.opacity > 145 then BeaconREF1b.opacity = 20

    if BeaconREF1.x < 1000 then
      BeaconREF1.x = BeaconREF1.x+30
      BeaconREF1b.x = BeaconREF1b.x+30
      BeaconREF1.y = BeaconREF1.y-36
      BeaconREF1b.y = BeaconREF1b.y-36
    end if

    if BeaconREF1.x => 0 then
      BeaconREF2.x = BeaconREF2.x+30
      BeaconREF2b.x = BeaconREF2b.x+30
      BeaconREF2.y = BeaconREF2.y+36
      BeaconREF2b.y = BeaconREF2b.y+36
      if BeaconREF2.opacity > 0 then BeaconREF2.opacity = BeaconREF2.opacity -3
      if BeaconREF2.opacity < 150 then BeaconREF2b.opacity = 15
      if BeaconREF2.opacity < 125 then BeaconREF2b.opacity = 13
      if BeaconREF2.opacity < 100 then BeaconREF2b.opacity = 10
      if BeaconREF2.opacity < 75 then BeaconREF2b.opacity = 7
      if BeaconREF2.opacity < 50 then BeaconREF2b.opacity = 5
      if BeaconREF2.opacity < 10 then BeaconREF2b.opacity = 0
    end If

    ' The numbers below were found through trial and error and do not reflect their position in the editor.  Annoying.
    if BeaconREF2.x => 2200 Then
      BeaconREF1.x = -1210
      BeaconREF1b.x = -1210
      BeaconREF1.y = 250
      BeaconREF1b.y = 250
      BeaconREF1.opacity = 0
      BeaconREF1b.opacity = 0
      BeaconREF2.x = 0
      BeaconREF2b.x = 0
      BeaconREF2.y = -2400
      BeaconREF2b.y = -2400
      BeaconREF2.opacity = 200
      BeaconREF2b.opacity = 20
      BeaconRefOn = false
    end If
  end if
End Sub

Sub BeaconSoundTimer_timer()
  BeaconSoundTimer.interval = 2080
  Playsound "Test_Beaconcut3"
End Sub


Sub BeaconAttractTimer_Timer
  BeaconFR.visible = false
  BeaconTimer.enabled = False

  'We need to turn OFF all the reflections and main flasher light here..
  BeaconFR.visible = false
  BeaconREF1.visible = false
  BeaconREF1b.visible = false
  BeaconREF1.visible = false
  BeaconREF1b.visible = false
  BeaconREF1.visible = false
  BeaconREF1b.visible = false
  BeaconREF2.visible = false
  BeaconREF2b.visible = false
  BeaconREF2.visible = false
  BeaconREF2b.visible = false
  BeaconREF2.visible = false
  BeaconREF2b.visible = false
  BeaconSoundTimer.enabled = False
  BeaconSoundTimer.interval = 1
  Stopsound "Test_BeaconCut3"
  Playsound "Relay_Flash_Off"
  BeaconAttractTimer.enabled = false ' don't run again.
End Sub

'******************************************************
'*******  Set Up Backglass  *******
'******************************************************


Sub SetBackglass()
  Dim obj


  For Each obj In VRBackglass
    obj.x = obj.x - 30
    obj.height = - obj.y + 235
    obj.y = -90 'adjusts the distance from the backglass towards the user
    obj.rotx=-86.5
  Next
End Sub

' Thalamus : Exit in a clean and proper way
Sub Table1_exit
  Controller.Pause = False
  Controller.Stop
End Sub


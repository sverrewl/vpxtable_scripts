'############################################################################################
'#######                                                                             ########
'#######          Tom & Jerry                                                        ########
'#######          (Original Mod of Hollywood Heat by watacaractr)                    ########
'#######                                                                             ########
'############################################################################################

' Version 1.0 - watacaractr - 2018
' - Original version which includes all base graphics, primitives, meshes, logic, scripting, etc.  A beautiful rendition of a ROM mod.  Thanks watacaractr!

' Version 1.2 - Gedankekojote97 - 2022
' - Added nFozzy physics, Fleep sounds and LUT selector.

' Version 1.31 - TastyWasps / psiomicron / Rawd / Sixtoe - 2023
' - Hybrid version to complete this VPX package.
' - Integrated psiomicron / rawd / sixtoe awesome VR room modeled after Jerry's mouse hole house.
' - Added a minimal VR Room with cabinet artwork for VR.
' - Desktop background upgraded to include Tom & Jerry logo.

Option Explicit
Randomize

Const cGameName = "tomjerry"

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

LoadVPM "01210000", "sys80.VBS", 3.1

'**********************************************************
'********       OPTIONS     *******************************
'**********************************************************

Dim BallShadows: Ballshadows=1      '******************set to 1 to turn on Ball shadows
Dim FlipperShadows: FlipperShadows=1  '***********set to 1 to turn on Flipper shadows
Dim ROMSounds: ROMSounds=0        '**********set to 0 for no rom sounds, 1 to play rom sounds.. mostly used for testing


Const BallBright = 0.5        '0 - Normal, 1 - Bright
Const VolumeDial = 0.8
Const BallRollVolume = 0.5      'Level of ball rolling volume. Value between 0 and 1
Const RampRollVolume = 0.5      'Level of ramp rolling volume. Value between 0 and 1

'************************************************
Const UseSolenoids = 2
Const UseLamps = True
Const UseSync = False
Const UseGI = False

' Standard Sounds
Const SSolenoidOn = "fx_solenoid"
Const SSolenoidOff = "fx_solenoidoff"
Const SCoin = ""
Const tnob = 5

Const AmbientBallShadowOn = 1
Const DynamicBallShadowsOn = 1
Const RubberizerEnabled = 1
Const TargetBouncerEnabled = 1    '0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.7   'Level of bounces. Recommmended value of 0.7

'----- VR Room Choice -----
Const VRRoomChoice = 1    ' 1 = Jerry's Mouse House, 2 = Minimal Room

Dim VR_Room, VR_Obj

If RenderingMode = 2 Then
  VR_Room = VRRoomChoice
Else
  VR_Room = 0
End If

If VR_Room = 1 Then
  For Each VR_Obj in TJRoom : VR_Obj.Visible = 1 : Next
  For Each VR_Obj in MinimalRoom : VR_Obj.Visible = 0 : Next
  SideRailLeft.Visible = 0
  SideRailRight.Visible = 0
  Primitive52.Visible = 0
ElseIf VR_Room = 2 Then
  For Each VR_Obj in TJRoom : VR_Obj.Visible = 0 : Next
  For Each VR_Obj in MinimalRoom : VR_Obj.Visible = 1 : Next
  SideRailLeft.Visible = 0
  SideRailRight.Visible = 0
  Primitive52.Visible = 0
Else
  For Each VR_Obj in TJRoom : VR_Obj.Visible = 0 : Next
  For Each VR_Obj in MinimalRoom : VR_Obj.Visible = 0 : Next
End If

Dim tablewidth: tablewidth = TomandJerry.width
Dim tableheight: tableheight = TomandJerry.height

Dim bsTrough, bslLock, bsrLock, dtRight, dtLeft, objekt, xx
Dim bMusicOn, startgamesound, multiballs

Sub TomandJerry_Init
  LoadLUT
     With Controller
         .GameName = cGameName
         If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
         .SplashInfoLine = "Tom & Jerry 1.2"&chr(13)&"Williams 2018"
         .HandleKeyboard = 0
         .ShowTitle = 0
         .ShowDMDOnly = 1
         .ShowFrame = 0
         .HandleMechanics = False
     .Games(cGameName).Settings.Value("sound") = ROMSounds
     .Hidden = 0
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With

  If B2SOn then
    for each objekt in backdropstuff : objekt.visible = 0 : next
  End If

  Intensity 'sets GI brightness depending on day/night slider settings

  multiballs=1
  ALlightsTimer.uservalue=1

  CaptiveKick.createball
  CaptiveKick.kick 0,0

    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

'************************Trough

  sw76.CreateBall
  sw20.CreateBall
  sw10.CreateBall
  Controller.Switch(76) = 1
  Controller.Switch(20) = 1
  Controller.Switch(10) = 1

'Kickers

    Set bslLock=New cvpmBallStack
    with bslLock
        .InitSaucer sw46,46,-170,10
    end with

    Set bsrLock=New cvpmBallStack
    with bsrLock
        .InitSaucer sw56,56,120,11
    end with

' Nudging
  vpmNudge.TiltSwitch = 57
    vpmNudge.Sensitivity = 1
    vpmNudge.TiltObj = Array(leftslingshot, rightslingshot, TopSlingShot, URightSlingShot, bumper1)

    Set dtLeft=New cvpmDropTarget
        dtLeft.InitDrop Array(sw40,sw50,sw60),Array(40,50,60)
        dtLeft.InitSnd SoundFX("",DOFDropTargets),SoundFX("Drop_Target_Reset_1",DOFDropTargets)

    Set dtRight=New cvpmDropTarget
        dtRight.InitDrop Array(sw41,sw51,sw61),Array(41,51,61)
        dtRight.InitSnd SoundFX("",DOFDropTargets),SoundFX("Drop_Target_Reset_2",DOFDropTargets)

  if ballshadows=1 then
        BallShadowUpdate.enabled=1
      else
        BallShadowUpdate.enabled=0
    end if

    if flippershadows=1 then
        FlipperLSh.visible=1
        FlipperRSh.visible=1
        MiniFlipperLSh.visible=1
        MiniFlipperRSh.visible=1
      else
        FlipperLSh.visible=0
        FlipperRSh.visible=0
        MiniFlipperLSh.visible=0
        MiniFlipperRSh.visible=0
    end if

  vpmMapLights CPULights

  startgamesound=1

  PlaySound "T&Jambient", -1, .05     'adjust the volume using the second parameter.

End Sub

'************************************************
' Solenoids
'************************************************
SolCallback(1) =    "SolLLock"
SolCallback(2) =    "SolRLock"
SolCallback(4) =    "dtDrop2"
SolCallback(5) =    "SolLeftTargetReset"
SolCallback(6) =    "SolRightTargetReset"
SolCallback(7) =     "dtDrop3"
SolCallback(8) =    "solknocker"
SolCallback(9) =    "solballrelease"

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

'//////////////////////////////////////////////////////////////////////
'// TIMERS
'//////////////////////////////////////////////////////////////////////


' The game timer interval is 10 ms
Sub GameTimer_Timer()
  Cor.Update            'update ball tracking
  RollingUpdate         'update rolling sounds
End Sub


' The frame timer interval is -1, so executes at the display frame rate
Sub FrameTimer_Timer()
  FlipperVisualUpdate       'update flipper shadows and primitives
End Sub

'***************
' Random Music
'***************

Dim song, numsongs, currentsong
numsongs=5    'Set to number of songs, name songs T&JbackgroundMusic0, T&JbackgroundMusic1, T&JbackgroundMusic2.. etc

Sub PlaySong 'random
    If bMusicOn Then
    currentsong = INT((numsongs-1) * RND(1) )
        song = "bgout_T&JbackgroundMusic" & currentsong & ".mp3"
        PlayMusic song
    End If
End Sub

Sub TomandJerry_MusicDone
    If bMusicOn Then
    if multiballs>1 then
      PlayMusic "bgout_T&Jmultiballmadness.mp3"     'if still in multiball repeat multiball music
      else
      PlaySong              'if not in multiball, play random song.
    end if
    End If
End Sub

'//////////////////////////////////////////////////////////////////////
'// Ball
'//////////////////////////////////////////////////////////////////////

If BallBright Then
  TomandJerry.BallImage = "ball_HDR_brighter"
Else
  TomandJerry.BallImage = "ball_HDR"
End If

'//////////////////////////////////////////////////////////////////////
'// FLIPPERS
'//////////////////////////////////////////////////////////////////////

 sub dtDrop2(enabled):dtLeft.hit 2:dtRight.hit 2:l7.state=0:end sub
 sub dtDrop3(enabled):dtLeft.hit 3:dtRight.hit 3:l8.state=0:end sub

Const ReflipAngle = 20

' Flipper Solenoid Callbacks (these subs mimics how you would handle flippers in ROM based tables)
Sub SolLFlipper(Enabled)
  If Enabled Then
    LF.Fire  'leftflipper.rotatetoend
        LeftFlipper1.RotateToEnd
    If leftflipper.currentangle < leftflipper.endangle + ReflipAngle Then
      RandomSoundReflipUpLeft LeftFlipper
    Else
      SoundFlipperUpAttackLeft LeftFlipper
      RandomSoundFlipperUpLeft LeftFlipper
    End If
  Else
    LeftFlipper.RotateToStart
        LeftFlipper1.RotateToStart
    If LeftFlipper.currentangle < LeftFlipper.startAngle - 5 Then
      RandomSoundFlipperDownLeft LeftFlipper
    End If
    FlipperLeftHitParm = FlipperUpSoundLevel
  End If
End Sub

Sub SolRFlipper(Enabled)
  If Enabled Then
    RF.Fire 'rightflipper.rotatetoend
    RightFlipper1.RotateToEnd
    If rightflipper.currentangle > rightflipper.endangle - ReflipAngle Then
      RandomSoundReflipUpRight RightFlipper
    Else
      SoundFlipperUpAttackRight RightFlipper
      RandomSoundFlipperUpRight RightFlipper
    End If
  Else
    RightFlipper.RotateToStart
    RightFlipper1.RotateToStart
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

' This subroutine updates the flipper shadows and visual primitives
Sub FlipperVisualUpdate
  FlipperLSh.RotZ = LeftFlipper.CurrentAngle
  FlipperRSh.RotZ = RightFlipper.CurrentAngle
End Sub

dim LF : Set LF = New FlipperPolarity
dim RF : Set RF = New FlipperPolarity

InitPolarity

'
''*******************************************
'' Late 70's to early 80's
'
'Sub InitPolarity()
'        dim x, a : a = Array(LF, RF)
'        for each x in a
'                x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1        'disabled
'                x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1
'                x.enabled = True
'                x.TimeDelay = 80
'        Next
'
'        AddPt "Polarity", 0, 0, 0
'        AddPt "Polarity", 1, 0.05, -2.7
'        AddPt "Polarity", 2, 0.33, -2.7
'        AddPt "Polarity", 3, 0.37, -2.7
'        AddPt "Polarity", 4, 0.41, -2.7
'        AddPt "Polarity", 5, 0.45, -2.7
'        AddPt "Polarity", 6, 0.576,-2.7
'        AddPt "Polarity", 7, 0.66, -1.8
'        AddPt "Polarity", 8, 0.743, -0.5
'        AddPt "Polarity", 9, 0.81, -0.5
'        AddPt "Polarity", 10, 0.88, 0
'
'        addpt "Velocity", 0, 0,         1
'        addpt "Velocity", 1, 0.16, 1.06
'        addpt "Velocity", 2, 0.41,         1.05
'        addpt "Velocity", 3, 0.53,         1'0.982
'        addpt "Velocity", 4, 0.702, 0.968
'        addpt "Velocity", 5, 0.95,  0.968
'        addpt "Velocity", 6, 1.03,         0.945
'
'        LF.Object = LeftFlipper
'        LF.EndPoint = EndPointLp
'        RF.Object = RightFlipper
'        RF.EndPoint = EndPointRp
'End Sub
'
'
'
''*******************************************
'' Mid 80's
'
'Sub InitPolarity()
'        dim x, a : a = Array(LF, RF)
'        for each x in a
'                x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1        'disabled
'                x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1
'                x.enabled = True
'                x.TimeDelay = 80
'        Next
'
'        AddPt "Polarity", 0, 0, 0
'        AddPt "Polarity", 1, 0.05, -3.7
'        AddPt "Polarity", 2, 0.33, -3.7
'        AddPt "Polarity", 3, 0.37, -3.7
'        AddPt "Polarity", 4, 0.41, -3.7
'        AddPt "Polarity", 5, 0.45, -3.7
'        AddPt "Polarity", 6, 0.576,-3.7
'        AddPt "Polarity", 7, 0.66, -2.3
'        AddPt "Polarity", 8, 0.743, -1.5
'        AddPt "Polarity", 9, 0.81, -1
'        AddPt "Polarity", 10, 0.88, 0
'
'        addpt "Velocity", 0, 0,         1
'        addpt "Velocity", 1, 0.16, 1.06
'        addpt "Velocity", 2, 0.41,         1.05
'        addpt "Velocity", 3, 0.53,         1'0.982
'        addpt "Velocity", 4, 0.702, 0.968
'        addpt "Velocity", 5, 0.95,  0.968
'        addpt "Velocity", 6, 1.03,         0.945
'
'        LF.Object = LeftFlipper
'        LF.EndPoint = EndPointLp
'        RF.Object = RightFlipper
'        RF.EndPoint = EndPointRp
'End Sub
'
'


'*******************************************
'  Late 80's early 90's

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



'
''*******************************************
'' Early 90's and after
'
'Sub InitPolarity()
'        dim x, a : a = Array(LF, RF)
'        for each x in a
'                x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1        'disabled
'                x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1
'                x.enabled = True
'                x.TimeDelay = 60
'        Next
'
'        AddPt "Polarity", 0, 0, 0
'        AddPt "Polarity", 1, 0.05, -5.5
'        AddPt "Polarity", 2, 0.4, -5.5
'        AddPt "Polarity", 3, 0.6, -5.0
'        AddPt "Polarity", 4, 0.65, -4.5
'        AddPt "Polarity", 5, 0.7, -4.0
'        AddPt "Polarity", 6, 0.75, -3.5
'        AddPt "Polarity", 7, 0.8, -3.0
'        AddPt "Polarity", 8, 0.85, -2.5
'        AddPt "Polarity", 9, 0.9,-2.0
'        AddPt "Polarity", 10, 0.95, -1.5
'        AddPt "Polarity", 11, 1, -1.0
'        AddPt "Polarity", 12, 1.05, -0.5
'        AddPt "Polarity", 13, 1.1, 0
'        AddPt "Polarity", 14, 1.3, 0
'
'        addpt "Velocity", 0, 0,         1
'        addpt "Velocity", 1, 0.16, 1.06
'        addpt "Velocity", 2, 0.41,         1.05
'        addpt "Velocity", 3, 0.53,         1'0.982
'        addpt "Velocity", 4, 0.702, 0.968
'        addpt "Velocity", 5, 0.95,  0.968
'        addpt "Velocity", 6, 1.03,         0.945
'
'        LF.Object = LeftFlipper
'        LF.EndPoint = EndPointLp
'        RF.Object = RightFlipper
'        RF.EndPoint = EndPointRp
'End Sub


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
Const EOSTnew = 1 'EM's to late 80's
'Const EOSTnew = 0.8 '90's and later
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
Const EOSReturn = 0.035  'mid 80's to early 90's
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


'//////////////////////////////////////////////////////////////////////
'// End Flipper
'//////////////////////////////////////////////////////////////////////

Dim GILevel, DayNight

Sub Intensity
  If DayNight <= 20 Then
      GILevel = .5
  ElseIf DayNight <= 40 Then
      GILevel = .4125
  ElseIf DayNight <= 60 Then
      GILevel = .325
  ElseIf DayNight <= 80 Then
      GILevel = .2375
  Elseif DayNight <= 100  Then
      GILevel = .15
  End If

  For each xx in GI: xx.IntensityScale = xx.IntensityScale * (GILevel): Next
  For each xx in DTLights: xx.IntensityScale = xx.IntensityScale * (GILevel): Next

End Sub

Sub ALlightsTimer_timer

  if me.uservalue>1 then
      ALlights(me.uservalue-2).state=0
      ALlightsA(me.uservalue-2).state=0
    elseif me.uservalue=1 then
      ALlights(9).state=0
      ALlightsA(9).state=0
    else
      ALlights(8).state=0
      ALlightsA(8).state=0
  end if

  ALlights(me.uservalue).state=1
  ALlightsA(me.uservalue).state=1

  me.uservalue = me.uservalue+1
  if me.uservalue>9 then me.uservalue=0
End sub

Sub FlipperTimer_Timer

'testbox1.text =

' BumperLightA1.state=BumperLight1.state

  LFlip.RotY = LeftFlipper.CurrentAngle
  RFlip.RotY = RightFlipper.CurrentAngle
  LFlip1.RotY = LeftFlipper1.CurrentAngle-90
  RFlip1.RotY = RightFlipper1.CurrentAngle-90
  Pgate.Rotz = Gate.CurrentAngle*0.7

  if sw40.isdropped then
    Lsw40.state=GI3.state
'       l1.state=0
      else
    Lsw40.state=0
'     l1.state=1
    end if

  if sw50.isdropped then
    Lsw50.state=GI3.state
'       l13.state=0
       else
    Lsw50.state=0
'       l13.state=1
    end if

  if sw60.isdropped then
    Lsw60.state=GI3.state
'       l31.state=0
       else
    Lsw60.state=0
'       l31.state=1
    end if

if sw41.isdropped then
    Lsw1.state=GI17.state
'        l6.state=0
      else
    Lsw1.state=0
'   l6.state=1
  end if

  if sw51.isdropped then
    Lsw2.state=GI17.state
'        l7.state=0
    else
    Lsw2.state=0
'   l7.state=1
  end if

  if sw61.isdropped then
    Lsw3.state=GI17.state
'        l8.state=0
    else
    Lsw3.state=0
'   l8.state=1
  end if

    if FlipperShadows=1 then
    FlipperLSh.RotZ = LeftFlipper.currentangle
    FlipperRSh.RotZ = RightFlipper.currentangle
      MiniFlipperLSh.RotZ = LeftFlipper1.currentangle
        MiniFlipperRSh.RotZ = RightFlipper1.currentangle
    end if
End Sub


' Ball locks / kickers

Sub sw46_unhit()
PlaysoundAt "VUKOut",sw46
End Sub

Sub sw56_unhit()
PlaysoundAt "VUKOut",sw56
End Sub

Sub sw46_Hit:
  F13b1.Duration 2, 1000, 0
  F13b2.Duration 2, 1000, 0
  PlaySound "VUKEnter"
  PlaySound "Spike bark"
  bslLock.AddBall 0:
End Sub

Sub sw56_Hit:
  PlaySound "VUKEnter"
  bsrLock.AddBall 0:
End Sub

Sub SolLLock(enabled)
  If enabled and LeftKickTimer.enabled=0 Then
    bslLock.ExitSol_On
    PkickarmL.RotZ = 15
    LeftKickTimer.uservalue = 0
    LeftKickTimer.Enabled = 1
  End If
End Sub

Sub LeftKickTimer_timer
  select case me.uservalue
    case 5:
    PkickarmL.rotz=0
    me.enabled=0
  end Select
  me.uservalue = me.uservalue+1
End Sub

Sub SolRLock(enabled)
  If enabled and RightKickTimer.enabled=0 Then
   if multiballs<2 then
    multiballs=3

    if l0.state=1 and RightKickTimer.enabled=0 then           'L0 is lit when game enabled
      bMusicOn=false          'turn off autoreplay
      endmusic            'stop regular music
      playmusic "bgout_T&Jmultiballmadness.mp3" 'playmultiball music
      bMusicOn=true         'turn back oon auto play so when multiball music ends music starts again
    end if
   end if

    bsrLock.ExitSol_On
    PkickarmR.RotZ = 15
    RightKickTimer.uservalue = 0
    RightKickTimer.Enabled = 1
  End If
End Sub

Sub RightKickTimer_timer
  select case me.uservalue
    case 5:
    PkickarmR.rotz=0
    me.enabled=0
  end Select
  me.uservalue = me.uservalue+1
End Sub


'******************************************************
'     TROUGH BASED ON NFOZZY'S via bord's pink panther
'******************************************************

Sub sw10_Hit():Controller.Switch(10) = 1:UpdateTrough:End Sub
Sub sw10_UnHit():Controller.Switch(10) = 0:UpdateTrough:End Sub
Sub sw20_Hit():Controller.Switch(20) = 1:UpdateTrough:End Sub
Sub sw20_UnHit():Controller.Switch(20) = 0:UpdateTrough:End Sub
Sub sw76_Hit():Controller.Switch(76) = 1:UpdateTrough:End Sub
Sub sw76_UnHit():Controller.Switch(76) = 0:UpdateTrough:End Sub

Sub UpdateTrough()
  UpdateTroughTimer.Interval = 500
  UpdateTroughTimer.Enabled = 1
End Sub

Sub UpdateTroughTimer_Timer()
  If sw76.BallCntOver = 0 Then sw20.kick 60, 1
  UpdateTroughTimer1.Interval = 500
  UpdateTroughTimer1.Enabled = 1
  Me.Enabled = 0
End Sub

Sub UpdateTroughTimer1_Timer()
  If sw20.BallCntOver = 0 Then sw10.kick 60, 1
  Me.Enabled = 0
End Sub

'******************************************************
'       DRAIN & RELEASE
'******************************************************

Sub sw66_Hit()
  if l0.state=1 then RandomSoundDrain sw66 : Playsound "Drain"
  if multiballs>0 then multiballs=multiballs-1
  if multiballs<2 then              'down to one ball after multiball
    endmusic                  'stop multiball music
    PlayMusic "bgout_T&JbackgroundMusic" & currentsong & ".mp3"                 'play regular musuic
  end if
  startgamesound = 1
    if sw20.BallCntOver > 0 or (sw76.ballcntover>0 and (sw46.ballcntover >0 or sw56.ballcntover>0)) or (sw46.ballcntover>0 and sw56.ballcntover>0) then
    bMusicOn = False  'turn off autoreplay music
    endMusic
  end if
  UpdateTrough
  Controller.Switch(66) = 1
End Sub

Sub sw66_UnHit()
  Controller.Switch(66) = 0
End Sub

Sub solballrelease(enabled)
  If enabled Then
    sw66.kick 60,40
  End If
End Sub

set Lights(12) = L12

Set LampCallback = GetRef("UpdateMultipleLamps")
Sub UpdateMultipleLamps
    if controller.lamp(12)=true and sw76.BallCntOver > 0 then
    sw76.kick 60, 7

    if bMusicOn = False then
      bMusicOn = TRUE 'TRUE
      playsong        'play random song
    end if
RandomSoundBallRelease sw76
    UpdateTrough
    end if

    if controller.lamp(2)=true then dtleft.hit 1:l1.state=0:dtRight.hit 1:l6.state=0
End Sub

 sub startgametimer_timer     '****make random sound for starting game
  Dim v
  if l0.state=1 then
    v = INT(2 * RND(1) )
    Select Case v
      Case 0:PlaySound"introducing"
      Case 1:PlaySound"introducing"
        End Select
  end if
  me.enabled=0
end sub

'//////////////////////////////////////////////////////////////////////
'// LUT
'//////////////////////////////////////////////////////////////////////

'//////////////---- LUT (Colour Look Up Table) ----//////////////
'0 = Fleep Natural Dark 1
'1 = Fleep Natural Dark 2
'2 = Fleep Warm Dark
'3 = Fleep Warm Bright
'4 = Fleep Warm Vivid Soft
'5 = Fleep Warm Vivid Hard
'6 = Skitso Natural and Balanced
'7 = Skitso Natural High Contrast
'8 = 3rdaxis Referenced THX Standard
'9 = CalleV Punchy Brightness and Contrast
'10 = HauntFreaks Desaturated
'11 = Tomate Washed Out
'12 = VPW Original 1 to 1
'13 = Bassgeige
'14 = Blacklight
'15 = B&W Comic Book

Dim LUTset, DisableLUTSelector, LutToggleSound, bLutActive
LutToggleSound = True
LoadLUT
'LUTset = 0     ' Override saved LUT for debug
SetLUT
DisableLUTSelector = 0  ' Disables the ability to change LUT option with magna saves in game when set to 1

Sub SetLUT  'AXS
  TomandJerry.ColorGradeImage = "LUT" & LUTset
end sub

Sub LUTBox_Timer
  LUTBox.TimerEnabled = 0
  LUTBox.Visible = 0
End Sub

Sub ShowLUT
  LUTBox.visible = 1
  Select Case LUTSet
    Case 0: LUTBox.text = "Fleep Natural Dark 1"
    Case 1: LUTBox.text = "Fleep Natural Dark 2"
    Case 2: LUTBox.text = "Fleep Warm Dark"
    Case 3: LUTBox.text = "Fleep Warm Bright"
    Case 4: LUTBox.text = "Fleep Warm Vivid Soft"
    Case 5: LUTBox.text = "Fleep Warm Vivid Hard"
    Case 6: LUTBox.text = "Skitso Natural and Balanced"
    Case 7: LUTBox.text = "Skitso Natural High Contrast"
    Case 8: LUTBox.text = "3rdaxis Referenced THX Standard"
    Case 9: LUTBox.text = "CalleV Punchy Brightness and Contrast"
    Case 10: LUTBox.text = "HauntFreaks Desaturated"
      Case 11: LUTBox.text = "Tomate washed out"
        Case 12: LUTBox.text = "VPW original 1on1"
        Case 13: LUTBox.text = "bassgeige"
        Case 14: LUTBox.text = "blacklight"
        Case 15: LUTBox.text = "B&W Comic Book"
    Case 16: LUTBox.text = "Skitso New ColorLut"
  End Select
  LUTBox.TimerEnabled = 1
End Sub

Sub SaveLUT
  Dim FileObj
  Dim ScoreFile

  Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if

  if LUTset = "" then LUTset = 0 'failsafe

  Set ScoreFile=FileObj.CreateTextFile(UserDirectory & "TomandJerryLUT.txt",True)
  ScoreFile.WriteLine LUTset
  Set ScoreFile=Nothing
  Set FileObj=Nothing
End Sub
Sub LoadLUT
    bLutActive = False
  Dim FileObj, ScoreFile, TextStr
  dim rLine

  Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    LUTset=0
    Exit Sub
  End if
  If Not FileObj.FileExists(UserDirectory & "TomandJerryLUT.txt") then
    LUTset=0
    Exit Sub
  End if
  Set ScoreFile=FileObj.GetFile(UserDirectory & "TomandJerryLUT.txt")
  Set TextStr=ScoreFile.OpenAsTextStream(1,0)
    If (TextStr.AtEndOfStream=True) then
      Exit Sub
    End if
    rLine = TextStr.ReadLine
    If rLine = "" then
      LUTset=0
      Exit Sub
    End if
    LUTset = int (rLine)
    Set ScoreFile = Nothing
      Set FileObj = Nothing
End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub TomandJerry_KeyDown(ByVal KeyCode)

  'LUT controls
  If keycode = LeftMagnaSave Then bLutActive = True

  If keycode = RightMagnaSave Then
    If bLutActive Then
      If DisableLUTSelector = 0 then
        LUTSet = LUTSet  - 1
        If LutSet < 0 Then LUTSet = 16
          SetLUT
          ShowLUT
        End If
    End If
  End If

  If keycode = keyInsertCoin1 or keycode = keyInsertCoin2 or keycode = keyInsertCoin3 or keycode = keyInsertCoin4 Then
    Select Case Int(rnd*3)
      Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
      Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
      Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
    End Select
  End If

    If keycode = RightFlipperKey Then FlipperActivate RightFlipper, RFPress End If
    If keycode = LeftFlipperKey Then FlipperActivate LeftFlipper, LFPress End If

    If keycode = PlungerKey Then
    Plunger.Pullback
    SoundPlungerPull()
    TimerPlunger.Enabled = True
    TimerPlunger2.Enabled = False
  End If

  If KeyCode = startgamekey and startgametimer.enabled=0 and startgamesound=1 then
    startgamesound=0
    startgametimer.enabled=1
  end if

    If KeyDownHandler(keycode) Then Exit Sub

End Sub

Sub TomandJerry_KeyUp(ByVal KeyCode)

  'LUT controls
  If keycode = LeftMagnaSave Then bLutActive = False

    If keycode = RightFlipperKey Then FlipperDeActivate RightFlipper, RFPress End If
    If keycode = LeftFlipperKey Then FlipperDeActivate LeftFlipper, LFPress End If

    If keycode = PlungerKey Then
    Plunger.Fire
    SoundPlungerReleaseBall()
    TimerPlunger.Enabled = False
    TimerPlunger2.Enabled = True
    PinCab_Shooter.Y = -73.66
  End If

    If KeyUpHandler(keycode) Then Exit Sub

End Sub

'Drop Targets
 Sub Sw40_Dropped:dtleft.Hit 1 : l1.state=1:StopSound"punch1":PlaySound"punch1": PlaySoundAt "Drop_Target_Down_1",DropTargetSound001:End Sub
 Sub Sw50_Dropped:dtleft.Hit 2 : l13.state=1:StopSound"punch2":PlaySound"punch2": PlaySoundAt "Drop_Target_Down_2",DropTargetSound001:End Sub
 Sub Sw60_Dropped:dtleft.Hit 3 : l31.state=1:StopSound"punch3":PlaySound"punch3": PlaySoundAt "Drop_Target_Down_3",DropTargetSound001:End Sub

 Sub Sw41_Dropped:dtRight.Hit 1: l6.state=0:StopSound"boink1":PlaySound"boink1": PlaySoundAt "Drop_Target_Down_4",DropTargetSound:End Sub
 Sub Sw51_Dropped:dtRight.Hit 2: l7.state=0:StopSound"boink2":PlaySound"boink2": PlaySoundAt "Drop_Target_Down_5",DropTargetSound:End Sub
 Sub Sw61_Dropped:dtRight.Hit 3: l8.state=0:StopSound"boink3":PlaySound"boink3": PlaySoundAt "Drop_Target_Down_6",DropTargetSound:End Sub


 Sub SolRightTargetReset(enabled)
    dim xx
    if enabled then
        dtRight.SolDropUp enabled
    l6.state=2
    l7.state=2
    l8.state=2
    end if
End Sub

Sub SolLeftTargetReset(enabled)
    dim xx
    if enabled then
        dtLeft.SolDropUp enabled
        l1.state=0
    l13.state=0
    l31.state=0
    end if
End Sub

'Bumpers

Sub bumper1_Hit
vpmTimer.PulseSw 30
RandomSoundBumperMiddle Bumper1
PlaySoundAt SoundFX("Tomboing",DOFContactors), Bumper1:
DOF 206, DOFPulse
BumperLight1.duration 0, 200, 1
BumperLightA1.duration 0, 200, 1
F13a1.duration 0, 200, 1
F13a2.duration 0, 200, 1
Light2.duration 1, 200, 0
GI15.duration 0, 200, 1
Light4.duration 1, 200, 0
GI21.duration 0, 200, 1
Flasher5a.duration 1, 200, 0
Flasher5a1.duration 1, 200, 0
Flasher5aLIGHTsec.duration 1, 200, 0
Flasher5a1LIGHTsec.duration 1, 200, 0
End Sub

Sub Tomscream_Hit()
if Activeball.vely > 0 then
Playsound "Tomscream1"
end if
end sub

Sub Tomscream2_Hit()
if Activeball.vely > 0 then
Playsound "Tomscream2"
end if
end sub

Sub bumper2_Hit
vpmTimer.PulseSw (60):
RandomSoundBumperTop Bumper2
PlaySoundAt SoundFX("fx_bumper2",DOFContactors), Bumper2:
DOF 207, DOFPulse
BumperLight2.duration 2, 200, 1
End Sub

Sub bumper3_Hit
vpmTimer.PulseSw (60):
RandomSoundBumperBottom Bumper3
PlaySoundAt SoundFX("fx_bumper3",DOFContactors), Bumper3:
DOF 208, DOFPulse
BumperLight3.duration 2, 200, 1
End Sub

'Wire Triggers

Sub SW38_Hit
al16.Duration 1, 300, 2
End Sub

Sub SW38a_Hit
al16.Duration 0,0,0
end sub

Sub SW39_Hit
al17.Duration 1, 300, 2
End Sub

Sub SW39a_Hit
al17.Duration 0,0,0
end sub

Sub SW42_Hit:Controller.Switch(42)=1:
    if Activeball.vely > 0 then
    Dim w
      w = INT(3 * RND(1) )
      Select Case w
      Case 0:PlaySound"zing3"
      Case 1:PlaySound""
      Case 2:PlaySound""
            End Select
End if
End Sub
Sub SW42_unHit:Controller.Switch(42)=0:End Sub

Sub SW52_Hit:Controller.Switch(52)=1:
    if Activeball.vely > 0 then
    Dim w
      w = INT(3 * RND(1) )
      Select Case w
      Case 0:PlaySound"zing2"
      Case 1:PlaySound""
      Case 2:PlaySound""
            End Select
End if
End Sub
Sub SW52_unHit:Controller.Switch(52)=0:End Sub

Sub SW62_Hit:Controller.Switch(62)=1:
    if Activeball.vely > 0 then
    Dim w
      w = INT(3 * RND(1) )
      Select Case w
      Case 0:PlaySound"zing1"
      Case 1:PlaySound""
      Case 2:PlaySound""
            End Select
End if
End Sub
Sub SW62_unHit:Controller.Switch(62)=0:End Sub

Sub SW73_Hit:Controller.Switch(73)=1:
End Sub
Sub SW73_unHit:Controller.Switch(73)=0:End Sub
Sub SW45_Hit:Controller.Switch(45)=1:
End Sub
Sub SW45_unHit:Controller.Switch(45)=0:End Sub
Sub SW55_Hit:Controller.Switch(55)=1:
  DOF 302, DOFOn
    al12.Duration 2, 600, 0
    al15.Duration 5, 1000, 0
End Sub
Sub SW55_unHit:Controller.Switch(55)=0:DOF 302, DOFOff:End Sub
Sub SW55a_Hit:Controller.Switch(55)=1:
  DOF 301, DOFOn
    al13.Duration 2, 600, 0
End Sub
Sub SW55a_unHit:Controller.Switch(55)=0:DOF 301, DOFOff:End Sub
Sub SW65_Hit:Controller.Switch(65)=1:
    al14.Duration 2, 600, 0
End Sub
Sub SW65_unHit:Controller.Switch(65)=0:End Sub
Sub SW75_Hit:Controller.Switch(75)=1:
    al11.Duration 2, 600, 0
End Sub
Sub SW75_unHit:Controller.Switch(75)=0:End Sub
Sub SW72_Hit:Controller.Switch(72)=1:End Sub
Sub SW72_unHit:Controller.Switch(72)=0:End Sub
Sub SW70_Hit:Controller.Switch(70)=1:
End Sub
Sub SW70_unHit:Controller.Switch(70)=0:End Sub
Sub SW74_Hit:Controller.Switch(74)=1:End Sub
Sub SW74_unHit:Controller.Switch(74)=0:End Sub
Sub SW31_Hit:Controller.Switch(31)=1:End Sub
Sub SW31_unHit:Controller.Switch(31)=0:End Sub

Sub SW80_hit
  If Jerry.roty=0 Then
    me.uservalue = 0
    me.timerinterval = 10
    me.timerenabled = 1
  End If
end sub

Sub SW80_timer
  jerry.roty = - me.uservalue * 3
  if jerry.roty < -120 Then
    jerry.roty = -120
    me.timerenabled = 0
  end if
  jerrySHADOW.roty = jerry.roty
  me.uservalue=me.uservalue+1
end sub

Sub SW80c_Hit
  Playsound "Bell"
End Sub

Sub SW81_hit
  If Jerry.roty=-120 Then
    me.uservalue=40
    me.timerinterval = 10
    me.timerenabled= 1
  End If
end sub

Sub SW81_timer
  jerry.roty = - me.uservalue * 3
  if jerry.roty > 0 Then
    jerry.roty = 0
    me.timerenabled = 0
  end if
  jerrySHADOW.roty = jerry.roty
  me.uservalue=me.uservalue - 1
end sub


Sub SW82_hit
  if TomHead.roty=0 and not sw90.timerenabled and not sw93.timerenabled then
    me.uservalue = 0
    me.timerinterval = 10
    me.timerenabled = 1
  End If
end sub

Sub SW82_timer
  TomHead.roty = me.uservalue * 0.75
  if TomHead.roty > 30 Then
    TomHead.roty = 30
    me.timerenabled = 0
  end if
  TomEYES.roty = TomHead.roty
  TomEyeLids.roty = TomHead.roty
  me.uservalue=me.uservalue+1
end sub

Sub SW83_hit
  if TomHead.roty=30 then
    me.uservalue = 40
    me.timerinterval = 10
    me.timerenabled = 1
  End If
end sub

Sub SW83_timer
  TomHead.roty = me.uservalue * 0.75
  if TomHead.roty < 0 Then
    TomHead.roty = 0
    me.timerenabled = 0
  end if
  TomEYES.roty = TomHead.roty
  TomEyeLids.roty = TomHead.roty
  me.uservalue=me.uservalue-1
end sub

Sub SW86_hit
  if TomEyeLids.rotx>6.5 then
    TomEyeLids.rotx = 0
    me.timerinterval = 30
    me.timerenabled = 1
  End If
end sub

Sub SW86_timer
  If TomEyeLids.rotx < 2 Then
    TomEyeLids.rotx = TomEyeLids.rotx + 0.1
  Elseif TomEyeLids.rotx < 6 Then
    TomEyeLids.rotx = TomEyeLids.rotx + 0.2
  Else
    TomEyeLids.rotx = TomEyeLids.rotx + 0.15
    if TomEyeLids.rotx > 7.5 Then
      TomEyeLids.rotx = 7.5
      me.timerenabled = 0
    End If
  End If
  'debug.print TomEyeLids.rotx
end sub

Sub SW88_Hit
TomHeadLight.Duration 1,1,1
end sub

Sub SW89_Hit
TomHeadLight.Duration 0,0,0
end sub

Sub SW90_hit
  if TomHead.roty=0  and not sw93.timerenabled and not sw82.timerenabled then
    me.uservalue = 1
    me.timerinterval = 10
    me.timerenabled = 1
  End If
end sub

Sub SW90_timer
  TomHead.roty = TomHead.roty + (me.uservalue * 0.4)
  if TomHead.roty >= 12 Then
    me.uservalue = -1
  end if
  if TomHead.roty < 0 Then
    TomHead.roty = 0
    me.timerenabled = 0
  End If
  TomEyes.roty = TomHead.roty
  TomEyeLids.roty = TomHead.roty
end sub

Sub SW93_hit
  if TomHead.roty=0 and not sw90.timerenabled and not sw82.timerenabled then
    me.uservalue = -1
    me.timerinterval = 10
    me.timerenabled = 1
  End If
end sub

Sub SW93_timer
  TomHead.roty = TomHead.roty + (me.uservalue * 0.6)
  if TomHead.roty <= -18 Then
    me.uservalue = 1
  end if
  if TomHead.roty > 0 Then
    TomHead.roty = 0
    me.timerenabled = 0
  End If
  TomEyes.roty = TomHead.roty
  TomEyeLids.roty = TomHead.roty
end sub

Sub SW95_hit
  Silhouette.transx=0
  me.uservalue = 0
  me.timerinterval = 10
  me.timerenabled = 1
end sub

Sub SW95_timer
  Silhouette.transx = - me.uservalue * 1
  If Silhouette.transx < -170 Then
    me.timerenabled = 0
    Silhouette.transx = 0
  End If
  me.uservalue=me.uservalue+1
end sub

Sub SW95a_Hit
  WindowA.Duration 1,1950,0
  WindowB.Duration 1,1950,0
  WindowC.Duration 1,1950,0
  WindowD.Duration 1,1950,0
end sub

Sub Tomtrough_hit
  if Activeball.vely < 0 then
    me.uservalue=1
    'me.timerenabled= 1
    Dim w
    w = INT(3 * RND(1) )
    Select Case w
      Case 0:PlaySound"Tom Scream Cannon"
      Case 1:PlaySound""
            Case 2:PlaySound""
    End Select
  End if
End Sub

Sub SW96_hit()
  if Activeball.vely < 0 then
    If Trashcan.rotz=0 then
      me.uservalue = 0
      me.timerenabled = 1
      me.timerinterval = 10
      Playsound "garbage can"
    end if
  end if
end sub

Sub SW96_timer
  if me.uservalue > 10 and me.uservalue < 21 then
    Trashcan.rotz = Trashcan.rotz - 1.2
  elseif me.uservalue > 20 and me.uservalue < 31 then
    Trashcan.rotz = Trashcan.rotz + 1.2
  elseif me.uservalue > 30 and me.uservalue < 41 then
    Trashcan.rotz = Trashcan.rotz - 1.2
  elseif me.uservalue > 40 and me.uservalue < 51 then
    Trashcan.rotz = Trashcan.rotz + 1.2
  elseif me.uservalue > 50 and me.uservalue < 61 then
    Trashcan.rotz = Trashcan.rotz - 1.2
  elseif me.uservalue > 60 and me.uservalue < 71 then
    Trashcan.rotz = Trashcan.rotz + 1.2
  elseif me.uservalue > 70 and me.uservalue < 81 then
    Trashcan.rotz = Trashcan.rotz - 1.2
  elseif me.uservalue > 80 and me.uservalue < 91 then
    Trashcan.rotz = Trashcan.rotz + 1.2
  elseif me.uservalue > 90 and me.uservalue < 101 then
    Trashcan.rotz = Trashcan.rotz - 1.2
  elseif me.uservalue > 99 then
    me.timerenabled = 0
    Trashcan.rotz = 0
  End If
  TrashcanSHADOW.rotz = Trashcan.rotz
  TrashcanLid.rotz = Trashcan.rotz
  'debug.print Trashcan.rotz
  me.uservalue=me.uservalue+1
end sub

Sub SW96a_Hit()
    if Activeball.vely < 0 then
        WindowA.Duration 1,3000,0
        WindowB.Duration 1,3000,0
        WindowC.Duration 1,3000,0
        WindowD.Duration 1,3000,0
    end if
end sub

Sub SW96c_hit()
  if Activeball.vely < 0 then
    me.uservalue=1
    'me.timerenabled= 1
    Dim w
    w = INT(3 * RND(1) )
    Select Case w
      Case 0: PlaySound "Thomas get this mouse"
      Case 1: PlaySound "Thomas"
      Case 2: Playsound ""
    End Select
  End if
End Sub


Sub SW98_hit()
  if Activeball.vely > 0 then
    If JerryHAMMER.roty = 0 then
      me.uservalue = 0
      me.timerenabled = 1
      me.timerinterval = 10
      Playsound "Tom inlane left"
    end if
  end if
end sub

Sub SW98_timer
  if me.uservalue < 11 Then
    'do nothing
  elseif me.uservalue < 48 then
    JerryHammer.roty = JerryHammer.roty + 3
  elseif me.uservalue < 68 then
    If me.uservalue = 58 then StarLightsLeft.Duration 1, 1000, 0
  else
    JerryHammer.roty = JerryHammer.roty - 3
    if JerryHammer.roty < 0 Then
      JerryHammer.roty = 0
      me.timerenabled = 0
    end If
  end if
  JerryHAMMERshadow.roty = JerryHammer.roty
  me.uservalue=me.uservalue+1
end sub

Sub SW99_hit()
  if Activeball.vely > 0 then
    If MusclesKnife.roty=0 then
      me.uservalue = 0
      me.timerenabled = 1
      me.timerinterval = 10
    end if
  end if
end sub

Sub SW99_timer
  if me.uservalue < 11 Then
    if me.uservalue = 2 then Playsound "Tom inlane right"
  elseif me.uservalue < 52 then
    MusclesKnife.roty = MusclesKnife.roty - 3.5
  elseif me.uservalue < 72 then
    If me.uservalue = 62 then StarLightsRight.Duration 1, 1000, 0
  else
    MusclesKnife.roty = MusclesKnife.roty + 3.5
    if MusclesKnife.roty > 0 Then
      MusclesKnife.roty = 0
      me.timerenabled = 0
    end If
  end if
  MuscleswithknifeSHADOW.roty = MusclesKnife.roty
  me.uservalue=me.uservalue+1
end sub


Sub SW100_Hit
GI11.Duration 0, 0, 0
end sub

Sub SW101_Hit:Controller.Switch(55)=1:
    PlaySound "ramp"
  if sw46.ballcntover=1 and sw56.ballcntover=1 then
  end if
end sub


Sub SW102_hit
  If Spike.roty=0 and not sw106.timerenabled Then
    me.uservalue=0
    me.timerinterval = 10
    me.timerenabled= 1
  End If
end sub

Sub SW102_timer
  If me.uservalue < 20 Then
    Spike.roty = Spike.roty + 1.5
  elseif me.uservalue < 90 Then
    'do Nothing
  Else
    Spike.roty = Spike.roty - 1.5
    if Spike.roty < 0 Then
      Spike.roty = 0
      me.timerenabled = 0
    end if
  end if
  SpikeSHADOW.roty = Spike.roty
  me.uservalue=me.uservalue + 1
end sub

Sub SW102b_hit
GI17.Duration 2, 1000, 1
GI2.Duration 2, 1000, 1
GI11.Duration 2, 1000, 1
LkickR1.Duration 2, 1000, 1
LkickR.Duration 2, 1000, 1
Lbulb4.Duration 2, 1000, 1
lsw1.Duration 2, 1000, 1
lsw2.Duration 2, 1000, 1
lsw3.Duration 2, 1000, 1
GI35.Duration 2, 1000, 1
GI37.Duration 2, 1000, 1
GI18.Duration 2, 1000, 1
Lbulb11.Duration 2, 1000, 1
Lbulb14.Duration 2, 1000, 1
end sub

Sub SW102c_hit
  if Activeball.vely > 0 then
    me.uservalue=1
    'me.timerenabled= 1
    Dim w
    w = INT(2 * RND(1) )
    Select Case w
      Case 0:PlaySound"That's my boy"
      Case 1:PlaySound"Listen pussycat"
    End Select
  End if
End Sub

Sub SW103_Hit
  GI9.Duration 2, 900, 1
    PlaySound ""
end sub

Sub leftoutlane_Hit()
  if Activeball.vely > 0 then
    me.uservalue=1
    me.timerenabled= 1
  end if
end sub

Sub leftoutlane_timer
select Case me.uservalue
Case 1: Playsound "Bomb fuse"
Case 2: Playsound ""
Case 3: Playsound ""
Case 4: Playsound ""
Case 5: Playsound ""
Case 6: Playsound ""
Case 7: Playsound ""
Case 8: Playsound ""
Case 9: Playsound ""
Case 10: Stopsound "Bomb fuse"
Case 11: Playsound "Bomb left"
Case 12: GI32.Duration 2, 1050, 0: me.timerenabled=0
end select
me.uservalue=me.uservalue+1
end sub

Sub rightoutlane_Hit()
  if Activeball.vely > 0 then
    me.uservalue=1
    me.timerenabled= 1
  end if
end sub

Sub rightoutlane_timer
select Case me.uservalue
Case 1: Playsound "Bomb fuse"
Case 2: Playsound ""
Case 3: Playsound ""
Case 4: Playsound ""
Case 5: Playsound ""
Case 6: Playsound ""
Case 7: Playsound ""
Case 8: Playsound ""
Case 9: Playsound ""
Case 10: Stopsound "Bomb fuse"
Case 11: Playsound "Bomb right"
Case 12: GI32.Duration 2, 1000, 0: me.timerenabled=0
end select
me.uservalue=me.uservalue+1
end sub

Sub SW104_Hit
Llsling2.Duration 2, 500, 1
end sub


Sub SW105_hit
  If Tuffy.transx = 0 then
    F13a.Duration 0, 2900, 0
    al6a.Duration 0, 2900, 0
    F13.Duration 0, 2900, 0
    me.uservalue=1
    me.timerenabled= 1
    me.timerinterval = 10
  End If
end sub

Sub SW105_timer
  if me.uservalue < 10 then
    Tuffy.transx = Tuffy.transx - 1
  elseif me.uservalue < 40 then
    Tuffy.transx = Tuffy.transx - 1.5
  elseif me.uservalue < 50 then
    Tuffy.transx = Tuffy.transx - 1.3
  elseif me.uservalue < 420 then
    if me.uservalue = 60 then PlaySound "Touche Pussycat"
    if me.uservalue = 70 then TuffyLight.Duration 1, 3300, 0
    if me.uservalue = 80 then TomHeadLight.Duration 1, 2000, 0
  elseif me.uservalue < 430 then
    Tuffy.transx = Tuffy.transx + 1.3
  elseif me.uservalue < 460 then
    Tuffy.transx = Tuffy.transx + 1.5
  else
    Tuffy.transx = Tuffy.transx + 1
    if Tuffy.transx > 0 Then
      Tuffy.transx = 0
      me.timerenabled = 0
    End If
  End If
  TuffyShadow.transx = Tuffy.transx
  me.uservalue=me.uservalue+1
end sub

Sub SW106_hit
  If Spike.roty=0 and not sw102.timerenabled Then
    me.uservalue=0
    me.timerinterval = 10
    me.timerenabled= 1
  End If
end sub

Sub SW106_timer
  If me.uservalue < 20 Then
    if me.uservalue = 0 then GI37.Duration 2, 1500, 0
    if me.uservalue = 10 then Lbulb11.Duration 2, 1500, 0
  elseif me.uservalue < 60 Then
    Spike.roty = Spike.roty + 2
  elseif me.uservalue < 140 Then
    'DO NOTHING
  Else
    Spike.roty = Spike.roty - 2
    if Spike.roty < 0 Then
      Spike.roty = 0
      me.timerenabled = 0
    end if
  end if
  SpikeSHADOW.roty = Spike.roty
  me.uservalue=me.uservalue + 1
end sub


Sub SW107_hit()
  if Activeball.vely > 0 then
    if Tyke.rotz=0 then
      me.uservalue = -1
      me.timerinterval = 10
      me.timerenabled = 1
    end if
  end if
end sub

Sub SW107_timer
  Tyke.rotz = Tyke.rotz + (me.uservalue * 0.9)
  if Tyke.rotz < - 26.5 Then
    me.uservalue = 1
  end If
  if tyke.rotz > 0 Then
    tyke.rotz = 0
    me.timerenabled = 0
  end if
  TykeSHADOW.rotz = Tyke.rotz
end sub

Sub SWTomsplat_hit
if Activeball.vely > 0 then
me.uservalue=1
'me.timerenabled= 1
Dim w
      w = INT(3 * RND(1) )
      Select Case w
      Case 0:PlaySound"Tom splat"
      Case 1:PlaySound""
      Case 2:PlaySound""
End Select
End if
End Sub

Sub SWTykebark_hit
if Activeball.vely > 0 then
me.uservalue=1
'me.timerenabled= 1
Dim w
      w = INT(3 * RND(1) )
      Select Case w
      Case 0:PlaySound"Tyke bark"
      Case 1:PlaySound""
      Case 2:PlaySound""
End Select
End if
End Sub

Sub SW108_Hit
    if Activeball.vely < 0 then
        Flasher5a.Duration 1, 700, 0
        Flasher5aLIGHTsec.duration 1, 700, 0
        Playsound "Ramp"
        me.uservalue=1
    me.timerenabled= 1
  end if
end sub

Sub SW109_Hit
Flasher5a1.Duration 1, 700, 0
Flasher5a1LIGHTsec.Duration 1, 700, 0
end sub

Sub JerryFuseOn_Hit
  if Activeball.vely > 0 then
    Playsound "Cannon fuse"
        Fuse.Duration 1, 1, 2
        FuseReflect1.Duration 1, 1, 2
        FuseReflect2.Duration 1, 1, 2
    me.uservalue=1
    me.timerenabled= 1
  end if
end sub

Sub JerryFuseOff_Hit
  if Activeball.vely < 0 then
    Stopsound "Cannon fuse"
        Playsound "Cannon fire"
        Fuse.Duration 0, 0, 0
        FuseReflect1.Duration 0, 0, 0
        FuseReflect2.Duration 0, 0, 0
    me.uservalue=1
    me.timerenabled= 1
  end if
end sub

Sub Cannon_Hit()
  if Activeball.vely < 0 then
    me.uservalue=1
    me.timerenabled= 1
  end if
end sub

Sub Cannon_timer
select Case me.uservalue
Case 1: CannonFire.Duration 2, 500, 0
Case 2: CannonFireReflect.Duration 2, 500, 0
Case 3: TomTroughLight.Duration 1, 750, 0: me.timerenabled=0
end select
me.uservalue=me.uservalue+1
end sub

'Targets
Sub sw43_Hit:vpmTimer.PulseSw (43):
    PlaySound "target1"
End Sub
Sub sw53_Hit:vpmTimer.PulseSw (53):
    PlaySound "target2"
End Sub
Sub sw63_Hit:vpmTimer.PulseSw (63):
    PlaySound "target3"
End Sub
Sub sw44_Hit:vpmTimer.PulseSw (44):
    PlaySound "target1"
End Sub
Sub sw54_Hit:vpmTimer.PulseSw (54):
    PlaySound "target2"
End Sub
Sub sw64_Hit:vpmTimer.PulseSw (64):
    PlaySound "Jerry gulp"
End Sub
Sub sw73a_Hit:vpmTimer.PulseSw (73):
    PlaySound "Jerry gulp":
  DOF 301, DOFPulse
End Sub



Sub SolKnocker(Enabled)
    If Enabled Then PlaySound SoundFX("Knocker",DOFKnocker)
End Sub

'**********Rubber Animations

sub RslingA_hit
  SlingA.visible=0
  SlingA1.visible=1
  me.uservalue=1
  Me.timerenabled=1
end sub

sub RslingA_timer                 'default 50 timer
  select case me.uservalue
    Case 1: SlingA1.visible=0: SlingA.visible=1
    case 2: SlingA.visible=0: SlingA2.visible=1
    Case 3: SlingA2.visible=0: SlingA.visible=1: Me.timerenabled=0
  end Select
  me.uservalue=me.uservalue+1
end sub


'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep, Tstep, URstep


Sub RightSlingShot_Slingshot
    RandomSoundSlingshotRight slingR
    DOF 202, DOFPulse
    vpmtimer.PulseSw(32)
    RSling.Visible = 0
    RSling1.Visible = 1
    slingR.objroty = -15
    RStep = 0
    RightSlingShot.TimerEnabled = 1
    GI6.Duration 2, 300, 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:slingR.objroty = -7
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:slingR.objroty = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
    RandomSoundSlingshotLeft slingL
    DOF 201, DOFPulse
    vpmtimer.pulsesw(32)
    LSling.Visible = 0
    LSling1.Visible = 1
    slingL.objroty = 15
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
    GI7.Duration 2, 300, 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:slingL.objroty = 7
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:slingL.objroty=0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub TopSlingShot_Slingshot
    RandomSoundSlingshotLeft slingT
    DOF 203, DOFPulse
    vpmtimer.pulsesw(32)
    Tsling.Visible = 0
    Tsling1.Visible = 1
    slingT.objroty = 15
    TStep = 0
    TopSlingShot.TimerEnabled = 1
End Sub

Sub TopSlingShot_Timer
    Select Case TStep
        Case 3:Tsling1.Visible = 0:Tsling2.Visible = 1:slingT.objroty = 7
        Case 4:Tsling2.Visible = 0:Tsling.Visible = 1:slingT.objroty=0:TopSlingShot.TimerEnabled = 0
    End Select
    TStep = TStep + 1
End Sub

Sub URightSlingShot_Slingshot
    RandomSoundSlingshotRight slingUR
    DOF 204, DOFPulse
    vpmtimer.PulseSw(32)
    URSling.Visible = 0
    URSling1.Visible = 1
    slingUR.objroty = -15
    URStep = 0
    URightSlingShot.TimerEnabled = 1
    Playsound "Wham"
    GI5.Duration 2, 300, 1
End Sub

Sub URightSlingShot_Timer
    Select Case URStep
        Case 3:URSLing1.Visible = 0:URSLing2.Visible = 1:slingUR.objroty = -7
        Case 4:URSLing2.Visible = 0:URSLing.Visible = 1:slingUR.objroty = 0:URightSlingShot.TimerEnabled = 0
    End Select
    URStep = URStep + 1
End Sub

'*********************************
'GI uselamps workaround by nfozzy
'*********************************

dim GIlamps : set GIlamps = New GIcatcherobject
Class GIcatcherObject   'object that disguises itself as a light. (UseLamps workaround for System80 GI circuit)
    Public Property Let State(input)
        dim x
        if input = 1 then 'If GI switch is engaged, turn off GI.
            for each x in gi : x.state = 0 : next
        elseif input = 0 then
            for each x in gi : x.state = 1 : next
        end if
        'tb.text = "gitcatcher.state = " & input    'debug
    End Property
End Class


set Lights(1) = GIlamps 'GI circuit


'*****************************************
'           BALL SHADOW by ninnuzu
'*****************************************
Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5)

Sub BallShadowUpdate_timer()
    Dim BOT, b
  Dim maxXoffset
  maxXoffset=6
    BOT = GetBalls

  ' render the shadow for each ball
    For b = 0 to UBound(BOT)
    BallShadow(b).X = BOT(b).X-maxXoffset*(1-(Bot(1).X)/(TomandJerry.Width/2))
    BallShadow(b).Y = BOT(b).Y + 12
    If BOT(b).Z > 0 and BOT(b).Z < 30 Then
      BallShadow(b).visible = 1
    Else
      BallShadow(b).visible = 0
    End If
  Next
End Sub

'Gottlieb Hollywood Heat
'added by Inkochnito
Sub editDips
  Dim vpmDips : Set vpmDips = New cvpmDips
  With vpmDips

    .AddForm 700,400,"Tom & Jerry - DIP Switches"
    .AddFrame 2,4,190,"Maximum Credits",49152,Array("8 Credits",0,"10 Credits",32768,"15 Credits",&H00004000,"20 Credits",49152)'dip 15&16
    .AddFrame 2,80,190,"Coin Chute 1 and 2 Control",&H00002000,Array("Separate",0,"Same",&H00002000)'dip 14
    .AddFrame 2,126,190,"Playfield Special",&H00200000,Array("Replay",0,"Extra Ball",&H00200000)'dip 22
    .AddFrame 2,172,190,"High Games to Date Control",&H00000020,Array("No Effect",0,"Reset High Games 2-5 on Power Off",&H00000020)'dip6
    .AddFrame 2,218,190,"Added a Letter to CARTOON when",&H40000000,Array("Top Rollovers are Completed",0,"Top Rollovers are Lit",&H40000000)'dip 31
    .AddFrame 2,264,190,"Drop Target Lamps",&H80000000,Array("Reset Ball to Ball",0,"Memorize Ball to Ball",&H80000000)'dip 32
    .AddFrame 205,4,190,"High Score to Date Awards",&H00C00000,Array("Not Displayed and No Award",0,"Displayed and No Award",&H00800000,"Displayed and 2 Replays",&H00400000,"Displayed and 3 Replays",&H00C00000)'dip 23&24
    '.AddFrame 205,80,190,"Balls per game",&H01000000,Array("5 balls",0,"3 balls",&H01000000)'dip 25
    .AddFrame 205,126,190,"Replay Limit",&H04000000,Array("No Limit",0,"One per Ball",&H04000000)'dip 27
    .AddFrame 205,172,190,"Novelty",&H08000000,Array("Normal",0,"Extra Ball and Replay Score 500000",&H08000000)'dip 28
    .AddFrame 205,218,190,"Game Mode",&H10000000,Array("Replay",0,"Extra Ball",&H10000000)'dip 29
    .AddFrame 205,264,190,"3rd Coin Chute Credits Control",&H20000000,Array("No Effect",0,"Add 9",&H20000000)'dip 30
    .AddChk 205,316,180,Array("Match Feature",&H02000000)'dip 26
    '.AddChk 2,316,190,Array("Attract Sound",&H00000040)'dip 7
    .AddLabel 50,340,300,20,"After Hitting OK, Press F3 to Reset Game with New Settings."
    .ViewDips
  End With

End Sub
Set vpmShowDips = GetRef("editDips")

'Finding an individual dip state based on scapino's Strikes and spares dip code - from unclewillys pinball pool


 Sub DipsTimer_Timer()
  Dim TheDips(32)
  Dim BPG
    Dim DipsNumber

  DipsNumber = Controller.Dip(3)

  TheDips(32) = Int(DipsNumber/128)
  If TheDips(32) = 1 then DipsNumber = DipsNumber - 128 end if
  TheDips(31) = Int(DipsNumber/64)
  If TheDips(31) = 1 then DipsNumber = DipsNumber - 64 end if
  TheDips(30) = Int(DipsNumber/32)
  If TheDips(30) = 1 then DipsNumber = DipsNumber - 32 end if
  TheDips(29) = Int(DipsNumber/16)
  If TheDips(29) = 1 then DipsNumber = DipsNumber - 16 end if
  TheDips(28) = Int(DipsNumber/8)
  If TheDips(28) = 1 then DipsNumber = DipsNumber - 8 end if
  TheDips(27) = Int(DipsNumber/4)
  If TheDips(27) = 1 then DipsNumber = DipsNumber - 4 end if
  TheDips(26) = Int(DipsNumber/2)
  If TheDips(26) = 1 then DipsNumber = DipsNumber - 2 end if
  TheDips(25) = Int(DipsNumber)

  BPG = TheDips(25)
  If BPG = 1 then
'   instcard.image="InstCard3Balls"
    Else
'   instcard.image="InstCard5Balls"
  End if
  DipsTimer.enabled=0
 End Sub

Sub TomandJerry_Paused:Controller.Pause = 1:End Sub
Sub TomandJerry_unPaused:Controller.Pause = 0:End Sub

Sub TomandJerry_Exit
SaveLUT
  Controller.Games(cGameName).Settings.Value("sound")=1
  Controller.Stop
End Sub

'//////////////////////////////////////////////////////////////////////
'// TargetBounce
'//////////////////////////////////////////////////////////////////////

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

'//////////////////////////////////////////////////////////////////////
'// PHYSICS DAMPENERS
'//////////////////////////////////////////////////////////////////////

' These are data mined bounce curves,
' dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
' Requires tracking ballspeed to calculate COR



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

' Thalamus - patched :     aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
    aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
    if debugOn then TBPout.text = str
  End Sub

  public sub Dampenf(aBall, parm) 'Rubberizer is handle here
    dim RealCOR, DesiredCOR, str, coef
    DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
    RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id)+0.0001)
    coef = desiredcor / realcor
    If abs(aball.velx) < 2 and aball.vely < 0 and aball.vely > -3.75 then
' Thalamus - patched :       aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
    aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
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

'//////////////////////////////////////////////////////////////////////
'// TRACK ALL BALL VELOCITIES FOR RUBBER DAMPENER AND DROP TARGETS
'//////////////////////////////////////////////////////////////////////

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

'//////////////////////////////////////////////////////////////////////
'// RAMP ROLLING SFX
'//////////////////////////////////////////////////////////////////////

'Ball tracking ramp SFX 1.0
'   Reqirements:
'          * Import A Sound File for each ball on the table for plastic ramps.  Call It RampLoop<Ball_Number> ex: RampLoop1, RampLoop2, ...
'          * Import a Sound File for each ball on the table for wire ramps. Call it WireLoop<Ball_Number> ex: WireLoop1, WireLoop2, ...
'          * Create a Timer called RampRoll, that is enabled, with a interval of 100
'          * Set RampBAlls and RampType variable to Total Number of Balls
' Usage:
'          * Setup hit events and call WireRampOn True or WireRampOn False (True = Plastic ramp, False = Wire Ramp)
'          * To stop tracking ball
'                 * call WireRampOff
'                 * Otherwise, the ball will auto remove if it's below 30 vp units
'

dim RampMinLoops : RampMinLoops = 4

' RampBalls
'      Setup:        Set the array length of x in RampBalls(x,2) Total Number of Balls on table + 1:  if tnob = 5, then RammBalls(6,2)
'      Description:
dim RampBalls(6,2)
'x,0 = ball x,1 = ID, 2 = Protection against ending early (minimum amount of updates)
'0,0 is boolean on/off, 0,1 unused for now
RampBalls(0,0) = False

' RampType
'     Setup: Set this array to the number Total number of balls that can be tracked at one time + 1.  5 ball multiball then set value to 6
'     Description: Array type indexed on BallId and a values used to deterimine what type of ramp the ball is on: False = Wire Ramp, True = Plastic Ramp
dim RampType(6)

Sub WireRampOn(input)  : Waddball ActiveBall, input : RampRollUpdate: End Sub
Sub WireRampOff() : WRemoveBall ActiveBall.ID : End Sub


' WaddBall (Active Ball, Boolean)
'     Description: This subroutine is called from WireRampOn to Add Balls to the RampBalls Array
Sub Waddball(input, RampInput)  'Add ball
  ' This will loop through the RampBalls array checking each element of the array x, position 1
  ' To see if the the ball was already added to the array.
  ' If the ball is found then exit the subroutine
  dim x : for x = 1 to uBound(RampBalls)  'Check, don't add balls twice
    if RampBalls(x, 1) = input.id then
      if Not IsEmpty(RampBalls(x,1) ) then Exit Sub 'Frustating issue with BallId 0. Empty variable = 0
    End If
  Next

  ' This will itterate through the RampBalls Array.
  ' The first time it comes to a element in the array where the Ball Id (Slot 1) is empty.  It will add the current ball to the array
  ' The RampBalls assigns the ActiveBall to element x,0 and ball id of ActiveBall to 0,1
  ' The RampType(BallId) is set to RampInput
  ' RampBalls in 0,0 is set to True, this will enable the timer and the timer is also turned on
  For x = 1 to uBound(RampBalls)
    if IsEmpty(RampBalls(x, 1)) then
      Set RampBalls(x, 0) = input
      RampBalls(x, 1) = input.ID
      RampType(x) = RampInput
      RampBalls(x, 2) = 0
      'exit For
      RampBalls(0,0) = True
      RampRoll.Enabled = 1   'Turn on timer
      'RampRoll.Interval = RampRoll.Interval 'reset timer
      exit Sub
    End If
    if x = uBound(RampBalls) then   'debug
      Debug.print "WireRampOn error, ball queue is full: " & vbnewline & _
      RampBalls(0, 0) & vbnewline & _
      Typename(RampBalls(1, 0)) & " ID:" & RampBalls(1, 1) & "type:" & RampType(1) & vbnewline & _
      Typename(RampBalls(2, 0)) & " ID:" & RampBalls(2, 1) & "type:" & RampType(2) & vbnewline & _
      Typename(RampBalls(3, 0)) & " ID:" & RampBalls(3, 1) & "type:" & RampType(3) & vbnewline & _
      Typename(RampBalls(4, 0)) & " ID:" & RampBalls(4, 1) & "type:" & RampType(4) & vbnewline & _
      Typename(RampBalls(5, 0)) & " ID:" & RampBalls(5, 1) & "type:" & RampType(5) & vbnewline & _
      " "
    End If
  next
End Sub

' WRemoveBall (BallId)
'    Description: This subroutine is called from the RampRollUpdate subroutine
'                 and is used to remove and stop the ball rolling sounds
Sub WRemoveBall(ID)   'Remove ball
  'Debug.Print "In WRemoveBall() + Remove ball from loop array"
  dim ballcount : ballcount = 0
  dim x : for x = 1 to Ubound(RampBalls)
    if ID = RampBalls(x, 1) then 'remove ball
      Set RampBalls(x, 0) = Nothing
      RampBalls(x, 1) = Empty
      RampType(x) = Empty
      StopSound("RampLoop" & x)
      StopSound("wireloop" & x)
    end If
    'if RampBalls(x,1) = Not IsEmpty(Rampballs(x,1) then ballcount = ballcount + 1
    if not IsEmpty(Rampballs(x,1)) then ballcount = ballcount + 1
  next
  if BallCount = 0 then RampBalls(0,0) = False  'if no balls in queue, disable timer update
End Sub

Sub RampRoll_Timer():RampRollUpdate:End Sub

Sub RampRollUpdate()    'Timer update
  dim x : for x = 1 to uBound(RampBalls)
    if Not IsEmpty(RampBalls(x,1) ) then
      if BallVel(RampBalls(x,0) ) > 1 then ' if ball is moving, play rolling sound
        If RampType(x) then
          PlaySound("RampLoop" & x), -1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * VolumeDial, AudioPan(RampBalls(x,0)), 0, BallPitchV(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
          StopSound("wireloop" & x)
        Else
          StopSound("RampLoop" & x)
          PlaySound("wireloop" & x), -1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * VolumeDial, AudioPan(RampBalls(x,0)), 0, BallPitch(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
        End If
        RampBalls(x, 2) = RampBalls(x, 2) + 1
      Else
        StopSound("RampLoop" & x)
        StopSound("wireloop" & x)
      end if
      if RampBalls(x,0).Z < 30 and RampBalls(x, 2) > RampMinLoops then  'if ball is on the PF, remove  it
        StopSound("RampLoop" & x)
        StopSound("wireloop" & x)
        Wremoveball RampBalls(x,1)
      End If
    Else
      StopSound("RampLoop" & x)
      StopSound("wireloop" & x)
    end if
  next
  if not RampBalls(0,0) then RampRoll.enabled = 0

End Sub

' This can be used to debug the Ramp Roll time.  You need to enable the tbWR timer on the TextBox
Sub tbWR_Timer()  'debug textbox
  me.text = "on? " & RampBalls(0, 0) & " timer: " & RampRoll.Enabled & vbnewline & _
  "1 " & Typename(RampBalls(1, 0)) & " ID:" & RampBalls(1, 1) & " type:" & RampType(1) & " Loops:" & RampBalls(1, 2) & vbnewline & _
  "2 " & Typename(RampBalls(2, 0)) & " ID:" & RampBalls(2, 1) & " type:" & RampType(2) & " Loops:" & RampBalls(2, 2) & vbnewline & _
  "3 " & Typename(RampBalls(3, 0)) & " ID:" & RampBalls(3, 1) & " type:" & RampType(3) & " Loops:" & RampBalls(3, 2) & vbnewline & _
  "4 " & Typename(RampBalls(4, 0)) & " ID:" & RampBalls(4, 1) & " type:" & RampType(4) & " Loops:" & RampBalls(4, 2) & vbnewline & _
  "5 " & Typename(RampBalls(5, 0)) & " ID:" & RampBalls(5, 1) & " type:" & RampType(5) & " Loops:" & RampBalls(5, 2) & vbnewline & _
  "6 " & Typename(RampBalls(6, 0)) & " ID:" & RampBalls(6, 1) & " type:" & RampType(6) & " Loops:" & RampBalls(6, 2) & vbnewline & _
  " "
End Sub


Function BallPitch(ball) ' Calculates the pitch of the sound based on the ball speed
    BallPitch = pSlope(BallVel(ball), 1, -1000, 60, 10000)
End Function

Function BallPitchV(ball) ' Calculates the pitch of the sound based on the ball speed Variation
  BallPitchV = pSlope(BallVel(ball), 1, -4000, 60, 7000)
End Function

'//////////////////////////////////////////////////////////////////////
'// RAMP TRIGGERS
'//////////////////////////////////////////////////////////////////////

Sub ramptrigger01_hit()
  WireRampOn True 'Play Plastic Ramp Sound
End Sub

Sub ramptrigger02_hit()
  WireRampOff ' Turn off the Plastic Ramp Sound
End Sub

Sub ramptrigger02_unhit()
  WireRampOn False ' On Wire Ramp Pay Wire Ramp Sound
End Sub

Sub ramptrigger03_hit()
  WireRampOff ' Exiting Wire Ramp Stop Playing Sound
End Sub

Sub ramptrigger03_unhit()
  PlaySoundAt "WireRamp_Stop", ramptrigger03
End Sub

'//////////////////////////////////////////////////////////////////////
'// Ball Rolling
'//////////////////////////////////////////////////////////////////////

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
  Dim BOT, b
  BOT = GetBalls

  ' stop the sound of deleted balls
  For b = UBound(BOT) + 1 to tnob
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
        BallShadowA(b).height=BOT(b).z - BallSize/4   'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping
      Else
        BallShadowA(b).height=BOT(b).z - BallSize/2 + 5
      End If
      BallShadowA(b).Y = BOT(b).Y + Ballsize/5 + fovY
      BallShadowA(b).X = BOT(b).X
      BallShadowA(b).visible = 1
    End If
  Next
End Sub

'//////////////////////////////////////////////////////////////////////
'// Mechanic Sounds
'//////////////////////////////////////////////////////////////////////

' This part in the script is an entire block that is dedicated to the physics sound system.
' Various scripts and sounds that may be pretty generic and could suit other WPC systems, but the most are tailored specifically for the TOM table

' Many of the sounds in this package can be added by creating collections and adding the appropriate objects to those collections.
' Create the following new collections:
'   Metals (all metal objects, metal walls, metal posts, metal wire guides)
'   Apron (the apron walls and plunger wall)
'   Walls (all wood or plastic walls)
'   Rollovers (wire rollover triggers, star triggers, or button triggers)
'   Targets (standup or drop targets, these are hit sounds only ... you will want to add separate dropping sounds for drop targets)
'   Gates (plate gates)
'   GatesWire (wire gates)
'   Rubbers (all rubbers including posts, sleeves, pegs, and bands)
' When creating the collections, make sure "Fire events for this collection" is checked.
' You'll also need to make sure "Has Hit Event" is checked for each object placed in these collections (not necessary for gates and triggers).
' Once the collections and objects are added, the save, close, and restart VPX.
'
' Tutorial vides by Apophis
' Part 1:   https://youtu.be/PbE2kNiam3g
' Part 2:   https://youtu.be/B5cm1Y8wQsk
' Part 3:   https://youtu.be/eLhWyuYOyGg


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
SpinnerSoundLevel = 0.5                                       'volume level; range [0, 1]

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
  PlaySoundAtLevelStatic SoundFX("Sling_L" & Int(Rnd*10)+1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

Sub RandomSoundSlingshotRight(sling)
  PlaySoundAtLevelStatic SoundFX("Sling_R" & Int(Rnd*8)+1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

'/////////////////////////////  BUMPER SOLENOID SOUNDS  ////////////////////////////
Sub RandomSoundBumperTop(Bump)
  PlaySoundAtLevelStatic SoundFX("Bumpers_Top_" & Int(Rnd*5)+1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperMiddle(Bump)
  PlaySoundAtLevelStatic SoundFX("Bumpers_Middle_" & Int(Rnd*5)+1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperBottom(Bump)
  PlaySoundAtLevelStatic SoundFX("Bumpers_Bottom_" & Int(Rnd*5)+1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
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


'/////////////////////////////////////////////////////////////////
'         End Mechanical Sounds
'/////////////////////////////////////////////////////////////////


' VR PLUNGER ANIMATION
Sub TimerPlunger_Timer
  If PinCab_Shooter.Y < 26.34 then
    PinCab_Shooter.Y = PinCab_Shooter.Y + 5
  End If
End Sub

Sub TimerPlunger2_Timer
  PinCab_Shooter.Y = -73.66 + (5* Plunger.Position) -20
End Sub

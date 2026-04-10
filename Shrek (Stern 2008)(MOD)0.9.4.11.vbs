'******************************************************
'* SHREK VPX 1.0***** *********************************
'******************************************************
'* STERN PINBALL 2008 *********************************
'******************************************************
'* SAM SYSTEM *****************************************
'******************************************************
'* ROM VERSION 1.41 (shr_141.bin)                     *
'******************************************************
'* CREATED BY NINUZZU AND TOM TOWER********************
'******************************************************
'* Many thanks to Groni, without him this wouldn't be *
'* possible.The table is based on Family Guy by Groni.*
'* Many thanks to comicalman, I used his FP table as  *
'* a resource for playfield images and toys.          *
'* Thanks to arngrim for the DOF script               *
'* Also a huge thanks to Freneticamnesic, who fixed   *
'* the bug with the flashers and to DJRobX, who fixed *
'* the center post script.                            *
'* Shrek, Fiona and Donkey models made by GLXB        *
'* What I did:                                        *
'* Readapted the textures to fit Groni's models       *
'* New Plastics from scatch                           *
'* Some textures redone from scratch                  *
'* Made a new model for the mirror                    *
'* Reworked some 3d models                            *
'* Added light system to Donkey mini-pinball          *
'* Other little stuff (lights, model positioning)     *
'******************************************************


'0.9.4 - iaakki - new baked sideblades
'0.9.4.2 - Paladin - cortracker , flipper scripting , fixed upper left flipper , removed rubber walls from dposts
'0.9.4.3 - Paladin - added Apophis beautified Script with flipper solenoid edit , added physics material several ramps
'0.9.4.4 - Paladin - added various missing physics materials
'0.9.4.5 - Gedankekojote97 - Added Dynamic Ballshadows. Added Flupper Domes (Lifted 30 VPUnits). Added new GI Lights. Added missing Fading script for the sidewalls. Added Slingshot Corrections
'0.9.4.6 - apophis - Fixed FlipperActivate and FlipperDeactivate calls in KeyDown and KeyUp subs. Updated FlipperNudge sub to latest. Reduced SSR scale to 0.2. Increased slingshot strength a little. Decreased PF friction a little.
'0.9.4.7 - Paladin - Reimported the physics material and added Fast Flips to the table_init
'0.9.4.8 - apophis - Adjust strength of flippers, slingshots, and scoop kicker. Reduced nudge strength. Increased tilt sensitivity. Added targetbouncer. Added Flacidizer. Added blocker walls. Locked all objects.
'0.9.4.9 - fluffhead35 - Added Fleep Sound.
'0.9.4.10 - fluffhead35 - fixed flipper sounds


Option Explicit
Randomize

'******************************************************
'  User Options
'******************************************************

'----- Mute Rom for PUP -----
' Changing this value may lead to an initial crash and reload recommend saving after change.
' This will turn sound off in VPinMAME for shr_141 , keep this in mind if using vpmalias

Dim MuteRom
MuteRom = False 'MuteRom = True for ROM Sound OFF  / MuteRom = False for ROM Sound ON

'******************************************************
'* VPM INIT *******************************************
'******************************************************

'----- General Sound Options -----
Const VolumeDial = 0.8        'Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Const BallRollVolume = 0.5      'Level of ball rolling volume. Value between 0 and 1
Const RampRollVolume = 0.5      'Level of ramp rolling volume. Value between 0 and 1

Dim Ballsize,BallMass
BallSize = 50
BallMass = 1

Const tnob = 9
Const lob = 4

Const AmbientBallShadowOn = 1
Const DynamicBallShadowsOn = 1

Dim tablewidth: tablewidth = Table1.width
Dim tableheight: tableheight = Table1.height

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Dim DesktopMode:DesktopMode = Table1.ShowDT
Dim UseVPMDMD:UseVPMDMD = DesktopMode

LoadVPM "01560000","sam.vbs",3.43

'********************
'Standard definitions
'********************

Const UseSolenoids = 1
Const UseLamps = 0
Const UseSync = 0
Const HandleMech = 0
Const SSolenoidOff = ""
Const SCoin = ""

'******************************************************
'* Sound Options **************************************
'******************************************************
If MuteRom = True Then
  Controller.Games("shr_141").Settings.Value("sound")=0
Else
  Controller.Games("shr_141").Settings.Value("sound")=1
End If
'******************************************************
'* ROM VERSION ****************************************
'******************************************************

Const cGameName = "shr_141"

'******************************************************
'* VARIABLES ******************************************
'******************************************************

Dim bsTrough, PlungerIM, BIP, SwampScoop, MerlinVUK, BABYBank, PinocchioTarget, CPUPos, CPDPos, MiniPF, MiniRight, MiniLeft
Dim  CastleGatePos, CastleGatePos2, CGUp, DonkeyPos, DonkeyPos2, DonkeyPos3, DonkeyDir, DonkeyPinball



'******************************************************
'* KEYS ***********************************************
'******************************************************

Sub Table1_KeyDown(ByVal keycode)
  If keycode = PlungerKey Then Plunger.PullBack : SoundPlungerPull()
  if keycode=StartGameKey then soundStartButton()

  If Keycode = RightFlipperKey then
    FlipperActivate RightFlipper, RFPress
    Controller.Switch(90)=1
    Controller.Switch(82)=1
    If MiniPF.Balls=0 Then
      'PlaySound SoundFX("Stern_MiniFlipperUp2",DOFContactors)
      MiniPF_RightFlipper.RotateToEnd
      If MiniPF_RightFlipper.currentangle < MiniPF_RightFlipper.endangle + ReflipAngle Then
        RandomSoundReflipUpRight MiniPF_RightFlipper
      Else
        SoundFlipperUpAttackRight MiniPF_RightFlipper
        RandomSoundFlipperUpRight MiniPF_RightFlipper
      End If
      DOF 101, 1
      MiniRight=1
    End If
  End If

  If Keycode = LeftFlipperKey then
    FlipperActivate LeftFlipper, LFPress
    Controller.Switch(84)=1
    If MiniPF.Balls=0 Then
      'PlaySound SoundFX("Stern_MiniFlipperUp1",DOFContactors)
      MiniPF_LeftFlipper.RotateToEnd
      If MiniPF_LeftFlipper.currentangle < MiniPF_LeftFlipper.endangle + ReflipAngle Then
        RandomSoundReflipUpLeft MiniPF_LeftFlipper
      Else
        SoundFlipperUpAttackLeft MiniPF_LeftFlipper
        RandomSoundFlipperUpLeft MiniPF_LeftFlipper
      End If
      DOF 102, 1
      MiniLeft=1
    End If
  End If

  If keycode = LeftTiltKey Then Nudge 90, 1:SoundNudgeLeft():End If
  If keycode = RightTiltKey Then Nudge 270, 1:SoundNudgeRight():End If
  If keycode = CenterTiltKey Then Nudge 0, 1:SoundNudgeCenter():End If

  If keycode = keyInsertCoin1 or keycode = keyInsertCoin2 or keycode = keyInsertCoin3 or keycode = keyInsertCoin4 Then
    Select Case Int(rnd*3)
      Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
      Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
      Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25

    End Select
  End If
  If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal keycode)

  If keycode = PlungerKey Then
    Plunger.Fire
    If BIP=1 then
      SoundPlungerReleaseBall()     'Plunger release sound when there is a ball in shooter lane
    elseIf BIP=0 then
      SoundPlungerReleaseNoBall()     'Plunger release sound when there is no ball in shooter lane
    End If
  End If

  If Keycode = RightFlipperKey then
    FlipperDeActivate RightFlipper, RFPress
    Controller.Switch(90)=0
    Controller.Switch(82)=0

    If MiniRight=1 Then
      'PlaySound SoundFX("Stern_MiniFlipperDown2",DOFContactors)
      MiniPF_RightFlipper.RotateToStart
      If MiniPF_RightFlipper.currentangle < MiniPF_RightFlipper.startAngle - 5 Then
        RandomSoundFlipperDownRight MiniPF_RightFlipper
      End If
      FlipperRightHitParm = FlipperUpSoundLevel
      DOF 101, 0
      MiniRight=0
    End If
  End If

  If Keycode = LeftFlipperKey then
    FlipperDeActivate LeftFlipper, LFPress
    Controller.Switch(84)=0
    If MiniLeft=1 Then
      'PlaySound SoundFX("Stern_MiniFlipperDown1",DOFContactors)
      MiniPF_LeftFlipper.RotateToStart
      If MiniPF_LeftFlipper.currentangle < MiniPF_LeftFlipper.startAngle - 5 Then
        RandomSoundFlipperDownLeft MiniPF_LeftFlipper
      End If
      FlipperLeftHitParm = FlipperUpSoundLevel
      DOF 102, 0
      MiniLeft=0
    End If
  End If

  If vpmKeyUp(keycode) Then Exit Sub
End Sub

'******************************************************
'******************************************************
'******************************************************
'* TABLE INIT *****************************************
'******************************************************
'******************************************************
'******************************************************

Sub Table1_Init

  '* ROM AND DMD ****************************************

  With Controller
    .GameName = cGameName
    .SplashInfoLine = "SHREK - STERN 2008"
    .HandleMechanics = 0
    .HandleKeyboard = 0
    .ShowDMDOnly = 1
    .ShowFrame = 0
    .ShowTitle = 0
    .Hidden = 0
    On Error Resume Next
    .Run
    If Err Then MsgBox Err.Description
    On Error Goto 0
  End With

  '* PINMAME TIMER **************************************

  PinMAMETimer.Interval = PinMAMEInterval
  PinMAMETimer.Enabled = 1

  '* STARTUP CALLS **************************************
  CPC.isdropped=1
  kicker1.CreateSizedBallWithMass BallSize/2,Ballmass
  kicker1.kick 180, 1
  kicker1.enabled = 0
  kicker2.CreateSizedBallWithMass BallSize/2,Ballmass
  kicker2.kick 180, 1
  kicker2.enabled = 0
  kicker3.CreateSizedBallWithMass BallSize/2,Ballmass
  kicker3.kick 180, 1
  kicker3.enabled = 0
  kicker4.CreateSizedBallWithMass BallSize/2,Ballmass
  kicker4.kick 180, 1
  kicker4.enabled = 0

  '* NUGDE **********************************************

  vpmNudge.TiltSwitch=-7
  vpmNudge.Sensitivity=7
  vpmNudge.TiltObj=Array(bumper1,bumper2,bumper3,LeftSlingshot,RightSlingshot)

  '* TROUGH *********************************************

  Set bsTrough = New cvpmBallStack
  With bsTrough
    .InitSw 0,21,20,19,18,0,0,0
    .InitKick BallRelease, 90, 8
    .InitExitSnd SoundFX("Stern_Release",DOFContactors), ""
    .Balls = 4
  End With

  '* MINI TROUGH ****************************************

  Set MiniPF = New cvpmBallStack
  With MiniPF
    .InitSaucer sw55,55, 90, 35
    .InitExitSnd "", ""
    .InitAddSnd ""
  End With
  sw55.createsizedball(14.6875)     'minipinball 5/8"
  MiniPF.AddBall 0


  '* SWAMP SCOOP *******************************************

  Set SwampScoop = New cvpmBallStack
  With SwampScoop
    .InitSw 0,13,0,0,0,0,0,0
    .InitKick sw13, 184, 35
    .InitExitSnd SoundFX("Stern_Scoopexit",DOFContactors), ""
  End With

  '* MERLIN VUK *****************************************

  Set MerlinVUK = New cvpmBallStack
  With MerlinVUK
    .InitSaucer sw64,64, 183, 4
    .InitExitSnd "",""
    .InitAddSnd ""
  End With

  '* BABY TARGETS 4-BANK DROPTARGETS ********************

  Set BABYBank = New cvpmDropTarget
  With BABYBank
    .InitDrop Array(sw47,sw44,sw46,sw45), Array(44,47,45,46)
    .Initsnd "", ""
  End With

  '* PINOCCHIO DROPTARGET *******************************

  Set PinocchioTarget = New cvpmDropTarget
  With PinocchioTarget
    .InitDrop Array(Array(sw9, sw9w)), Array(9)
    .Initsnd "", ""
  End With

  '* IMPULSE PLUNGER ************************************

  Const IMPowerSetting = 33
  Const IMTime = 0.33
  Set PlungerIM = New cvpmImpulseP
  With plungerIM
    .InitImpulseP BIPREG, IMPowerSetting, IMTime
    .InitExitSnd "Stern_Autoplung", ""
    .CreateEvents "PlungerIM"
  End With

  '* Fast Flips *****************************************
    On Error Resume Next
    InitVpmFFlipsSAM
    If Err Then MsgBox "You need the latest sam.vbs in order to run this table, available with vp10.5 rev3434"
    On Error Goto 0

  '******************************************************
  '******************************************************
  '******************************************************
  '* TABLE INIT END *************************************
  '******************************************************
  '******************************************************
  '******************************************************

End Sub

Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub
Sub table1_exit():Controller.Stop:End sub

'******************************************************
'* SOLENOID CALLS *************************************
'******************************************************

SolCallback(1)  = "solTrough"
SolCallback(2)  = "AutoLaunch"
SolCallback(3)  = "BankReset"
SolCallback(4)  = "CPDSol"                   'Ball Saver Down
SolCallback(5)  = "MerlinEject"
SolCallback(6)  = "PinocchioReset"
'SolCallback(7)  = "LeftSlingHit"            'Left Slingshot
'SolCallback(8)  = "RightSlingHit"           'Right Slingshot
'SolCallback(9)  = "Bump1Sol"                'Bottom Bumper
'SolCallback(10) = "Bump2Sol"                'Right Bumper
'SolCallback(11) = "Bump3Sol"                'Top Bumper
SolCallback(12) = "CPUSol"                   'Ball Saver Up
SolCallback(13) = "Scoopexit"                'Swamp Eject Scoop
SolCallback(15) = "LFlipper"
SolCallback(16) = "RFlipper"
SolCallback(17) = "Miniflipper_Left"
SolCallback(18) = "Miniflipper_Right"
SolCallback(19) = "CastleGuardSol"
SolCallBack(20) = "DonkeyMove"        'Donkey Motor Drive
SolCallBack(21) = "MiniPFSol"
SolCallback(23) = "Flash1"      'Flash Lower Left
SolCallback(25) = "Flash3"      'Flash Backpanel Left
SolCallback(26) = "Flash4"      'Flash Backpanel Center
SolCallback(27) = "Flash5"      'Flash Backpanel Right
SolCallback(28) = "Flash6"      'Flash Magic Mirror
SolCallback(29) = "setlamp 199,"      'Flash Fiona
SolCallback(30) = "setlamp 190,"      'Flash RIght Orbit (Spinner)
SolCallback(31) = "setlamp 191,"      'Flash Pop Bumpers
SolCallback(32) = "Flash2"      'Flash Lower Right


'****************************************************************************
'PHYSICS DAMPENERS
'****************************************************************************

'These are data mined bounce curves,
'dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
'Requires tracking ballspeed to calculate COR


Sub dPosts_Hit(idx)
  TargetBouncer activeball, 1
  RubbersD.dampen Activeball
End Sub

Sub dSleeves_Hit(idx)
  TargetBouncer activeball, 0.7
  SleevesD.Dampen Activeball
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
'// FLIPPERS
'//////////////////////////////////////////////////////////////////////


' Flipper collide subs
Sub LeftFlipper_Collide(parm)
  CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
  LeftFlipperCollide parm     'fleep stuff
End Sub


Sub RightFlipper_Collide(parm)
  CheckLiveCatch Activeball, RightFlipper, RFCount, parm
  RightFlipperCollide parm      'fleep stuff
End Sub

' This subroutine updates the flipper shadows and visual primitives
Sub FlipperVisualUpdate
  FlipperLSh.RotZ = LeftFlipper.CurrentAngle
  FlipperRSh.RotZ = RightFlipper.CurrentAngle
End Sub

dim LF : Set LF = New FlipperPolarity
dim RF : Set RF = New FlipperPolarity

InitPolarity

''*******************************************
'' Early 90's and after
'
Sub InitPolarity()
  dim x, a : a = Array(LF, RF)
  for each x in a
    x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1        'disabled
    x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1
    x.enabled = True
    x.TimeDelay = 60
  Next

  AddPt "Polarity", 0, 0, 0
  AddPt "Polarity", 1, 0.05, -5.5
  AddPt "Polarity", 2, 0.4, -5.5
  AddPt "Polarity", 3, 0.6, -5.0
  AddPt "Polarity", 4, 0.65, -4.5
  AddPt "Polarity", 5, 0.7, -4.0
  AddPt "Polarity", 6, 0.75, -3.5
  AddPt "Polarity", 7, 0.8, -3.0
  AddPt "Polarity", 8, 0.85, -2.5
  AddPt "Polarity", 9, 0.9,-2.0
  AddPt "Polarity", 10, 0.95, -1.5
  AddPt "Polarity", 11, 1, -1.0
  AddPt "Polarity", 12, 1.05, -0.5
  AddPt "Polarity", 13, 1.1, 0
  AddPt "Polarity", 14, 1.3, 0

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
'                        FLIPPER TRICKS
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
Const PI = 3.1415927

Function dSin(degrees)
  dsin = sin(degrees * Pi/180)
End Function

Function dCos(degrees)
  dcos = cos(degrees * Pi/180)
End Function

'*************************************************
' Check ball distance from Flipper for Rem
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

' Used for drop targets and stand up targets
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

'*************************************************
' End - Check ball distance from Flipper for Rem
'*************************************************

dim LFPress, RFPress, LFCount, RFCount
dim LFState, RFState
dim EOST, EOSA,Frampup, FElasticity,FReturn
dim RFEndAngle, LFEndAngle

LFState = 1
RFState = 1
EOST = leftflipper.eostorque
EOSA = leftflipper.eostorqueangle
Frampup = LeftFlipper.rampup
FElasticity = LeftFlipper.elasticity
FReturn = LeftFlipper.return
Const EOSTnew = 0.8
Const EOSAnew = 1
Const EOSRampup = 0
Dim SOSRampup
SOSRampup = 2.5
Const LiveCatch = 16
Const LiveElasticity = 0.45
Const SOSEM = 0.815
Const EOSReturn = 0.025
Const SOSReturn = 0.027 'adjust this value until you get the right amount of movement when the flipper is hit when down

LFEndAngle = Leftflipper.endangle
RFEndAngle = RightFlipper.endangle

Sub FlipperActivate(Flipper, FlipperPress)
  FlipperPress = 1
  Flipper.Elasticity = FElasticity
  Flipper.return = FReturn

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
      Flipper.Return = SOSReturn 'set new return strength when flipper is down
      FCount = 0
      FState = 1
    End If
  ElseIf Abs(Flipper.currentangle) <= Abs(Flipper.endangle) and FlipperPress = 1 then
    if FCount = 0 Then FCount = GameTime

    If FState <> 2 Then
      Flipper.return = FReturn 'insure original return strength is used when flipping - could get reset above after keypress before flipper moves
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

Sub RDampen_Timer()
  Cor.Update
End Sub



'******************************************************
' VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************

Const TargetBouncerEnabled = 1    '0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 1 'Level of bounces. Recommmended value of 0.7

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



'******************************************************
'* CASTLE GUARD GATE **********************************
'******************************************************

Sub CastleGuardsw_Hit()
  Controller.Switch(35)=0
  CastleGuardUp.enabled=1
  CGUp=1
  CastleGuardsw.isdropped=1
End Sub

Sub CastleGuardSol (enabled)
  If enabled then
    Controller.Switch(35)=1
  End if
End Sub

Sub CastleGuardUp_Timer()
  Select Case CastleGatePos
    Case 0: CastleGuardP.RotX=0:CastleGatePos=1
    Case 1: CastleGuardP.RotX=0:
    Case 2: CastleGuardP.RotX=-15
    Case 3: CastleGuardP.RotX=-30
    Case 4: CastleGuardP.RotX=-45
    Case 5: CastleGuardP.RotX=-60
    Case 6: CastleGuardP.RotX=-75
    Case 7: CastleGuardP.RotX=-90:Me.Enabled=0:CastleGatePos=0
  End Select
  If CastleGatePos>0 then CastleGatePos=CastleGatePos+1
End Sub

Sub CastleGuardDown_Timer()
  Select Case CastleGatePos2
    Case 0: CastleGuardP.RotX=-90:CastleGatePos2=1
    Case 1: CastleGuardP.RotX=-90:
    Case 2: CastleGuardP.RotX=-75
    Case 3: CastleGuardP.RotX=-60
    Case 4: CastleGuardP.RotX=-45
    Case 5: CastleGuardP.RotX=-30
    Case 6: CastleGuardP.RotX=-15
    Case 7: CastleGuardP.RotX=0:CastleGuardsw.isdropped=0:CGUp=0
    Case 8: CastleGuardP.RotX=5
    Case 9: CastleGuardP.RotX=10
    Case 10: CastleGuardP.RotX=15
    Case 11: CastleGuardP.RotX=20
    Case 12: CastleGuardP.RotX=15
    Case 13: CastleGuardP.RotX=10
    Case 14: CastleGuardP.RotX=5
    Case 15: CastleGuardP.RotX=0
    Case 16: CastleGuardP.RotX=-5
    Case 17: CastleGuardP.RotX=-10
    Case 18: CastleGuardP.RotX=-15
    Case 19: CastleGuardP.RotX=-10
    Case 20: CastleGuardP.RotX=-5
    Case 21: CastleGuardP.RotX=0
    Case 22: CastleGuardP.RotX=5
    Case 23: CastleGuardP.RotX=10
    Case 24: CastleGuardP.RotX=15
    Case 25: CastleGuardP.RotX=10
    Case 26: CastleGuardP.RotX=5
    Case 27: CastleGuardP.RotX=0
    Case 28: CastleGuardP.RotX=-5
    Case 29: CastleGuardP.RotX=-10
    Case 30: CastleGuardP.RotX=-15
    Case 31: CastleGuardP.RotX=-10
    Case 32: CastleGuardP.RotX=-5
    Case 33: CastleGuardP.RotX=0
    Case 34: CastleGuardP.RotX=5
    Case 35: CastleGuardP.RotX=10
    Case 36: CastleGuardP.RotX=10
    Case 37: CastleGuardP.RotX=5
    Case 38: CastleGuardP.RotX=0
    Case 39: CastleGuardP.RotX=-5
    Case 40: CastleGuardP.RotX=-10
    Case 41: CastleGuardP.RotX=-10
    Case 42: CastleGuardP.RotX=-5
    Case 43: CastleGuardP.RotX=0
    Case 44: CastleGuardP.RotX=5
    Case 45: CastleGuardP.RotX=10
    Case 46: CastleGuardP.RotX=5
    Case 47: CastleGuardP.RotX=0
    Case 48: CastleGuardP.RotX=-5
    Case 49: CastleGuardP.RotX=-10
    Case 50: CastleGuardP.RotX=-5
    Case 51: CastleGuardP.RotX=0
    Case 52: CastleGuardP.RotX=5
    Case 53: CastleGuardP.RotX=0
    Case 54: CastleGuardP.RotX=-5
    Case 55: CastleGuardP.RotX=0
    Case 56: CastleGuardP.RotX=5
    Case 57: CastleGuardP.RotX=0:Me.Enabled=0:CastleGatePos2=0
  End Select
  If CastleGatePos2>0 then CastleGatePos2=CastleGatePos2+1
End Sub

'******************************************************
'* MINI PLAYFIELD *************************************
'******************************************************

Sub Miniflipper_Right(Enabled)
  If Enabled Then
    MiniPF_RightFlipper.RotateToEnd

    If MiniPF_RightFlipper.currentangle < MiniPF_RightFlipper.endangle + ReflipAngle Then
      RandomSoundReflipUpLeft MiniPF_RightFlipper
    Else
      SoundFlipperUpAttackLeft MiniPF_RightFlipper
      RandomSoundFlipperUpLeft MiniPF_RightFlipper
    End If
  Else
    MiniPF_RightFlipper.RotateToStart
    If MiniPF_RightFlipper.currentangle < MiniPF_RightFlipper.startAngle - 5 Then
      RandomSoundFlipperDownLeft MiniPF_RightFlipper
    End If
    FlipperRightHitParm = FlipperUpSoundLevel
    End If
End Sub

Sub Miniflipper_Left(Enabled)
  If Enabled Then
    DonkeySWTimer.Enabled=1
    MiniPF_LeftFlipper.RotateToEnd

    If MiniPF_LeftFlipper.currentangle < MiniPF_LeftFlipper.endangle + ReflipAngle Then
      RandomSoundReflipUpLeft MiniPF_LeftFlipper
    Else
      SoundFlipperUpAttackLeft MiniPF_LeftFlipper
      RandomSoundFlipperUpLeft MiniPF_LeftFlipper
    End If
  Else
    MiniPF_LeftFlipper.RotateToStart
    If MiniPF_LeftFlipper.currentangle < MiniPF_LeftFlipper.startAngle - 5 Then
      RandomSoundFlipperDownLeft MiniPF_LeftFlipper
    End If
    FlipperLeftHitParm = FlipperUpSoundLevel
    End If
End Sub

Sub MiniPFSol(enabled)
  If enabled then
    MiniPF.ExitSol_On
  end if
End Sub

Sub sw55_Hit():MiniPF.AddBall 0:End Sub

Sub DonkeySWTimer_Timer()
  DonkeySW.enabled=1
End Sub

Sub DonkeySW_hit()
  DonkeyPinball=2
End Sub

Sub DonkeyTurn_Hit()
  If DonkeyPinball=0 then
    DonkeyTimer2.enabled=1
    DonkeyPinball=3
    DonkeySWTimer.enabled=1
  End If
End Sub

Sub DonkeyReset_Hit
  If DonkeyPinball >0 then
    DonkeyPinball = 0
  End If
End Sub

Sub DonkeyDelay_Timer()
  DonkeyTimer.Enabled = 1
  DonkeyDelay.Enabled = 0
End Sub

'******************************************************
'* DONKEY MOVEMENT ************************************
'******************************************************

Sub DonkeyMove (enabled)
  If enabled AND DonkeyPinball=0 then
    DonkeyDelay.Enabled = 1
  Else
    If enabled AND DonkeyPinball=2 then
      DonkeyTimer3.Enabled = 1
    End If
  End If
End Sub

Sub DonkeyTimer_Timer()
  Select Case DonkeyPos
    Case 0: DonkeyP.RotY=0:DonkeyPos3=0:DonkeyPos=1
    Case 1: DonkeyP.RotY=+2
    Case 2: DonkeyP.RotY=+4
    Case 3: DonkeyP.RotY=+6
    Case 4: DonkeyP.RotY=+8
    Case 5: DonkeyP.RotY=+10
    Case 6: DonkeyP.RotY=+12
    Case 7: DonkeyP.RotY=+10
    Case 8: DonkeyP.RotY=+8
    Case 9: DonkeyP.RotY=+6
    Case 10: DonkeyP.RotY=+4
    Case 11: DonkeyP.RotY=+2
    Case 12: DonkeyP.RotY=0:DonkeyPos=1:DonkeyTimer.Enabled=0:DonkeyPinball=0
  End Select
  If DonkeyPos>0 then DonkeyPos=DonkeyPos+1
  Sh5.RotZ = DonkeyP.RotY
End Sub

Sub DonkeyTimer2_Timer()
  Select Case DonkeyPos2
    Case 0: DonkeyP.RotY=0:DonkeyPos2=1
    Case 1: DonkeyP.RotY=0:
    Case 2: DonkeyP.RotY=15
    Case 3: DonkeyP.RotY=30
    Case 4: DonkeyP.RotY=45
    Case 5: DonkeyP.RotY=60
    Case 6: DonkeyP.RotY=75
    Case 7: DonkeyP.RotY=90
    Case 8: DonkeyP.RotY=105
    Case 9: DonkeyP.RotY=120:DonkeyTimer2.Enabled=0:DonkeyPos2=0:DonkeyPinball=3
  End Select
  If DonkeyPos2>0 then DonkeyPos2=DonkeyPos2+1
  Sh5.RotZ = DonkeyP.RotY
End Sub

Sub DonkeyTimer3_Timer()
  Select Case DonkeyPos3
    Case 0: DonkeyP.RotY=120:DonkeyPos3=1
    Case 1: DonkeyP.RotY=105
    Case 2: DonkeyP.RotY=90
    Case 3: DonkeyP.RotY=70
    Case 4: DonkeyP.RotY=60
    Case 5: DonkeyP.RotY=45
    Case 6: DonkeyP.RotY=30
    Case 7: DonkeyP.RotY=15
    Case 8: DonkeyP.RotY=0
    Case 9: DonkeyP.RotY=0:DonkeyTimer3.Enabled=0:DonkeyPos3=0:DonkeySWTimer.enabled=0:DonkeyPinball=0:DonkeySW.enabled=0
  End Select
  If DonkeyPos3>0 then DonkeyPos3=DonkeyPos3+1
  Sh5.RotZ = DonkeyP.RotY
End Sub

'******************************************************
'* TROUGH *********************************************
'******************************************************

Sub solTrough(enabled)
  If enabled then
    bsTrough.ExitSol_On
    vpmTimer.PulseSw 22
  end if
End Sub

'******************************************************
'* AUTOLAUNCH *****************************************
'******************************************************

Sub AutoLaunch(Enabled)
  If Enabled Then
    PlungerIM.AutoFire
  End If
End Sub

'******************************************************
'* SWAMP SCOOP ****************************************
'******************************************************

Dim bBall, bZpos

Sub Scoopexit(enabled)
  If enabled then
    SwampScoop.ExitSol_On
  end if
End Sub

Sub sw13_Hit
  Set bBall = ActiveBall
  'PlaySound "Stern_Scoopenter"
  SoundSaucerKick 1, sw13
  bZpos = 50
  Me.TimerInterval = 2
  Me.TimerEnabled = 1
End Sub

Sub sw13_Timer
  bBall.Z = bZpos
  bZpos = bZpos-2
  If bZpos <40 Then
    Me.TimerEnabled = 0
    Me.DestroyBall
    SwampScoop.AddBall Me
  End If
End Sub

'******************************************************
'* MERLIN VUK *****************************************
'******************************************************

Sub MerlinEject(enabled)
  If enabled then
    MerlinVUK.ExitSol_On
  End if
End Sub

Sub sw64_Hit:SoundSaucerKick 1, sw64:MerlinVUK.AddBall 0:End Sub

'******************************************************
'* GINGY SPINNER **************************************
'******************************************************

Sub sw39_Spin:SoundSpinner sw39 : vpmTimer.PulseSw 39:End Sub

'******************************************************
'* DRAIN **********************************************
'******************************************************

Sub Drain_Hit()
  RandomSoundDrain Drain
  bsTrough.AddBall Me
End Sub

'******************************************************
'* POP BUMPERS ****************************************
'******************************************************

Dim dirRing1 : dirRing1 = -1
Dim dirRing2 : dirRing2 = -1
Dim dirRing3 : dirRing3 = -1

Sub Bumper1_Hit : vpmTimer.PulseSw 32 : RandomSoundBumperTop bumper1 : Me.TimerEnabled = 1 : End Sub
Sub Bumper2_Hit : vpmTimer.PulseSw 31 : RandomSoundBumperMiddle bumper2 : Me.TimerEnabled = 1 : End Sub
Sub Bumper3_Hit : vpmTimer.PulseSw 30 : RandomSoundBumperBottom bumper3 : Me.TimerEnabled = 1 : End Sub

Sub Bumper1_timer()
  BR1.Z = BR1.Z + (5 * dirRing1)
  If BR1.Z <= -35 Then dirRing1 = 1
  If BR1.Z >= 0 Then
    dirRing1 = -1
    BR1.Z = 0
    Me.TimerEnabled = 0
  End If
End Sub

Sub Bumper2_timer()
  BR2.Z = BR2.Z + (5 * dirRing2)
  If BR2.Z <= -35 Then dirRing2 = 1
  If BR2.Z >= 0 Then
    dirRing2 = -1
    BR2.Z = 0
    Me.TimerEnabled = 0
  End If
End Sub

Sub Bumper3_timer()
  BR3.Z = BR3.Z + (5 * dirRing3)
  If BR3.Z <= -35 Then dirRing3 = 1
  If BR3.Z >= 0 Then
    dirRing3 = -1
    BR3.Z = 0
    Me.TimerEnabled = 0
  End If
End Sub

'******************************************************
'* TARGETS ********************************************
'******************************************************

Sub sw3_Hit:vpmTimer.PulseSw 3:If CGUp= 1 then CastleGuardDown.enabled=1:End If:End Sub

Sub sw4_Hit:vpmTimer.PulseSw 4:End Sub

Sub sw5_Hit:vpmTimer.PulseSw 5:End Sub

Sub sw8_Hit:vpmTimer.PulseSw 8:End Sub

Sub sw10_Hit:vpmTimer.PulseSw 10:End Sub

Sub sw41_Hit:vpmTimer.PulseSw 41:End Sub

Sub sw42_Hit:vpmTimer.PulseSw 42:End Sub

Sub sw43_Hit:vpmTimer.PulseSw 43:End Sub


'******************************************************
'* MINIPF TARGETS **************************************
'******************************************************

Sub sw50_Hit:vpmTimer.PulseSw 50:End Sub
Sub sw51_Hit:vpmTimer.PulseSw 51:End Sub

'******************************************************
'* MINI PF SWITCHES ***********************************
'******************************************************

Sub sw52_Hit:Controller.Switch(52) = 1:End Sub
Sub sw52_Unhit:Controller.Switch(52) = 0:End Sub

Sub sw53_Hit:Controller.Switch(53) = 1:End Sub
Sub sw53_Unhit:Controller.Switch(53) = 0:End Sub

Sub sw54_Hit:Controller.Switch(54) = 1:End Sub
Sub sw54_Unhit:Controller.Switch(54) = 0:End Sub

'******************************************************
'* ROLLOVER SWITCHES **********************************
'******************************************************

Sub sw6_Hit:Controller.Switch(6) = 1:sw6p.Z=-2:End Sub
Sub sw6_Unhit:Controller.Switch(6) = 0:sw6p.Z=0:End Sub

Sub sw7_Hit:Controller.Switch(7) = 1:sw7p.Z=-2:End Sub
Sub sw7_Unhit:Controller.Switch(7) = 0:sw7p.Z=0:End Sub

Sub sw23_Hit:Controller.Switch(23) = 1:sw23p.Z=-2:BIP=1:End Sub
Sub sw23_Unhit:Controller.Switch(23) = 0:sw23p.Z=0:BIP=0:End Sub

Sub sw24_Hit:Controller.Switch(24) = 1:sw24p.Z=-2:End Sub
Sub sw24_Unhit:Controller.Switch(24) = 0:sw24p.Z=0:End Sub

Sub sw25_Hit:Controller.Switch(25) = 1:sw25p.Z=-2:End Sub
Sub sw25_Unhit:Controller.Switch(25) = 0:sw25p.Z=0:End Sub

Sub sw28_Hit:Controller.Switch(28) = 1:sw28p.Z=-2:End Sub
Sub sw28_Unhit:Controller.Switch(28) = 0:sw28p.Z=0:End Sub

Sub sw29_Hit:Controller.Switch(29) = 1:sw29p.Z=-2:End Sub
Sub sw29_Unhit:Controller.Switch(29) = 0:sw29p.Z=0:End Sub

Sub sw40_Hit:Controller.Switch(40) = 1:sw40p.Z=-2:End Sub
Sub sw40_Unhit:Controller.Switch(40) = 0:sw40p.Z=0:End Sub

Sub sw48_Hit:Controller.Switch(48) = 1:End Sub
Sub sw48_Unhit:Controller.Switch(48) = 0:End Sub

Sub sw57_Hit:Controller.Switch(57) = 1:End Sub
Sub sw57_Unhit:Controller.Switch(57) = 0:End Sub

'******************************************************
'* SLINGSHOTS *****************************************
'******************************************************
Dim LStep, RStep
Sub LeftSlingShot_Slingshot
  LS.VelocityCorrect(ActiveBall)
  RandomSoundSlingshotLeft sling1
  'Playsound SoundFX("Stern_LeftSlingshot",DOFContactors),0,1,-0.05,0.05
  vpmTimer.PulseSw 26
  LSling.Visible = 0
  LSling1.Visible = 1
  sling1.TransZ = -20
  LStep = 0
  LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
  Select Case LStep
    Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling1.TransZ = -10
    Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling1.TransZ = 0:LeftSlingShot.TimerEnabled = 0
  End Select
  LStep = LStep + 1
End Sub


Sub RightSlingShot_Slingshot
  RS.VelocityCorrect(ActiveBall)
  RandomSoundSlingshotRight sling2
  'Playsound SoundFX("Stern_RightSlingshot",DOFContactors),0,1,0.05,0.05
  vpmTimer.PulseSw 27
  RSling.Visible = 0
  RSling1.Visible = 1
  sling2.TransZ = -20
  RStep = 0
  RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
  Select Case RStep
    Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling2.TransZ = -10
    Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling2.TransZ = 0:RightSlingShot.TimerEnabled = 0
  End Select
  RStep = RStep + 1
End Sub

'******************************************************
'* CENTER POST ****************************************
'******************************************************
'* DJRobX fixed the ball saver emulation error. Game  *
'* seems to fire both up and down to put the saver    *
'* down. Create a "latch" to hold it down'            *
'******************************************************

Dim SaverPos, SaverTime
SaverPos = 0
SaverTime = Now

Sub CPDSol (enabled)
  If enabled then
    SaverTime = DateAdd("s",3, Now)
    CPUSol (true)
  End If
End Sub

Sub CPUSol (enabled)
  If enabled then
    If SaverTime > Now Then
      If SaverPos = 1 Then
        ' ** Move Saver Down **'
        Controller.Switch(2) = 1
        Controller.Switch(1) = 0
        CenterpostDown.Enabled=1
        SaverPos=0
        CPC.isdropped=1
      End If
    Else
      If SaverPos = 0 Then
        ' ** Move Saver Up ** '
        Controller.Switch(2) = 0
        Controller.Switch(1) = 1
        CenterpostUp.Enabled=1
        SaverPos=1
        CPC.isdropped=0
      End If
    End If
  End If
End Sub

Sub CenterpostUp_Timer()
  Select Case CPUPos
    Case 0: CenterPost.Z=-15
    Case 1: CenterPost.Z=-05:Playsound SoundFX("Stern_Centerpost_Up",DOFContactors)
    Case 2: CenterPost.Z=0
    Case 3: CenterPost.Z=0:CPUPos=0:CenterpostUp.Enabled=0
  End Select
  If CPUPos=>0 then CPUPos=CPUPos+1
End Sub

Sub CenterpostDown_Timer()
  Select Case CPDPos
    Case 0: CenterPost.Z=-05
    Case 1: CenterPost.Z=-15:Playsound SoundFX("Stern_Centerpost_Down",DOFContactors)
    Case 2: CenterPost.Z=-26
    Case 3: CenterPost.Z=-26:CPDPos=0:CenterpostDown.Enabled=0
  End Select
  If CPDPos=>0 then CPDPos=CPDPos+1
End Sub

'******************************************************
'* MAGIC MIRROR TARGET ********************************
'******************************************************
Dim MirrorPos:MirrorPos=1
Dim MirrorDir : MirrorDir=0
Sub sw49_Hit
  vpmTimer.PulseSw 49
  Me.TimerEnabled = 1
  PlaySound SoundFX("Stern_Beercanhit",DOFContactors)
  TargetBouncer activeball, 1.4
End Sub

Sub sw49_Timer()
  MirrorP.Rotx = MirrorP.Rotx - 2*MirrorPos
  If MirrorP.Rotx < - 10 AND MirrorPos=1 Then MirrorPos=-1
  If MirrorP.Rotx > 6 AND MirrorPos=-1 Then MirrorDir=1 : MirrorPos=1
  If MirrorP.Rotx <0 AND MirrorDir=1 Then MirrorP.Rotx =0 : MirrorDir=0 : Me.TimerEnabled = 0
End Sub

'******************************************************
'* PINOCCHIO TARGET ***********************************
'******************************************************

Dim sw9Dir

Sub Pinocchioreset (enabled)
  If enabled then
    RandomSoundDropTargetReset sw9p
    sw9Dir = 1
    sw9.TimerEnabled = 1
    sw9w.isdropped=0
    sw9.isdropped=0
    PinocchioTarget.DropSol_On
  End If
End Sub

Sub sw9_Hit
  SoundDropTargetDrop sw9p
  PinocchioTarget.Hit 1
  sw9Dir=-1
  Me.TimerEnabled = 1
  TargetBouncer activeball, 1.4
End Sub

Sub sw9_Timer()
  Select Case sw9Dir
    Case -1
      sw9P.z = sw9P.z -5
      If sw9p.Z <  -65 Then sw9p.Z = -65 : Me.TimerEnabled = 0
    Case 1
      sw9P.z = sw9P.z + 5
      If sw9p.Z >  0 Then sw9p.Z = 0 : Me.TimerEnabled = 0
  End Select
End Sub

'******************************************************
'* BABY TARGETBANK ************************************
'******************************************************
Dim sw44Dir, sw45Dir, sw46Dir, sw47Dir

Sub Bankreset (enabled)
  If enabled then
    RandomSoundDropTargetReset sw47p
    sw44Dir = 1
    sw44.TimerEnabled = 1
    sw45Dir = 1
    sw45.TimerEnabled = 1
    sw46Dir = 1
    sw46.TimerEnabled = 1
    sw47Dir = 1
    sw47.TimerEnabled = 1
    BABYBank.DropSol_On
  End If
End Sub

Sub sw44_Hit:SoundDropTargetDrop sw44p: BABYBank.Hit 2:sw44Dir = -1:Me.TimerEnabled = 1:End Sub
Sub sw45_Hit:SoundDropTargetDrop sw45p: BABYBank.Hit 4:sw45Dir = -1:Me.TimerEnabled = 1:End Sub
Sub sw46_Hit:SoundDropTargetDrop sw46p: BABYBank.Hit 3:sw46Dir = -1:Me.TimerEnabled = 1:End Sub
Sub sw47_Hit:SoundDropTargetDrop sw47p: BABYBank.Hit 1:sw47Dir = -1:Me.TimerEnabled = 1:End Sub

'* Y TARGET *******************************************

Sub sw44_Timer()
  Select Case sw44Dir
    Case -1
      sw44P.z = sw44P.z -5
      If sw44p.Z <  -60 Then sw44p.Z = -60 : Me.TimerEnabled = 0
    Case 1
      sw44P.z = sw44P.z + 5
      If sw44p.Z >  0 Then sw44p.Z = 0 : Me.TimerEnabled = 0
  End Select
End Sub

'* B TARGET *******************************************

Sub sw45_Timer()
  Select Case sw45Dir
    Case -1
      sw45P.z = sw45P.z -5
      If sw45p.Z <  -60 Then sw45p.Z = -60 : Me.TimerEnabled = 0
    Case 1
      sw45P.z = sw45P.z + 5
      If sw45p.Z >  0 Then sw45p.Z = 0 : Me.TimerEnabled = 0
  End Select
End Sub

'* A TARGET *******************************************
Sub sw46_Timer()
  Select Case sw46Dir
    Case -1
      sw46P.z = sw46P.z -5
      If sw46p.Z <  -60 Then sw46p.Z = -60 : Me.TimerEnabled = 0
    Case 1
      sw46P.z = sw46P.z + 5
      If sw46p.Z >  0 Then sw46p.Z = 0 : Me.TimerEnabled = 0
  End Select
End Sub

'* B TARGET *******************************************
Sub sw47_Timer()
  Select Case sw47Dir
    Case -1
      sw47P.z = sw47P.z -5
      If sw47p.Z <  -60 Then sw47p.Z = -60 : Me.TimerEnabled = 0
    Case 1
      sw47P.z = sw47P.z + 5
      If sw47p.Z >  0 Then sw47p.Z = 0 : Me.TimerEnabled = 0
  End Select
End Sub

'******************************************************
'* FLIPPER CALLS **************************************
'******************************************************

SolCallback(sLRFlipper) = "RFlipper"
SolCallback(sLLFlipper) = "LFlipper"

Const ReflipAngle = 20

Sub LFlipper(Enabled)
  If Enabled Then
    LF.Fire  'leftflipper.rotatetoend
    LeftFlipperSmall.RotateToEnd
    If leftflipper.currentangle < leftflipper.endangle + ReflipAngle Then
      RandomSoundReflipUpLeft LeftFlipper
    Else
      SoundFlipperUpAttackLeft LeftFlipper
      RandomSoundFlipperUpLeft LeftFlipper
    End If
  Else
    LeftFlipper.RotateToStart
    LeftFlipperSmall.RotateToStart
    If LeftFlipper.currentangle < LeftFlipper.startAngle - 5 Then
      RandomSoundFlipperDownLeft LeftFlipper
    End If
    FlipperLeftHitParm = FlipperUpSoundLevel
    End If
End Sub

Sub RFlipper(Enabled)
  If Enabled Then
    RF.Fire 'rightflipper.rotatetoend

    If rightflipper.currentangle > rightflipper.endangle - ReflipAngle Then
      RandomSoundReflipUpRight RightFlipper
    Else
      SoundFlipperUpAttackRight RightFlipper
      RandomSoundFlipperUpRight RightFlipper
    End If
  Else
    RightFlipper.RotateToStart
    If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
      RandomSoundFlipperDownRight RightFlipper
    End If
    FlipperRightHitParm = FlipperUpSoundLevel
  End If
End Sub


'******************************************************
'* LEFT GATE ******************************************
'******************************************************

Sub LeftGate_hit:vpmTimer.PulseSw 33:LeftGateFl.RotateToEnd:End Sub
Sub LeftGate_Unhit:LeftGateFl.RotateToStart:End Sub

'******************************************************
'* REAL TIME UPDATES *********************************
'******************************************************

Set MotorCallback = GetRef("GameTimer")

Sub GameTimer
  RightFlipperP.Rotz = RightFlipper.CurrentAngle
  LeftFlipperP.Rotz = LeftFlipper.CurrentAngle
  LeftFlipperSmallP.RotY = LeftFlipperSmall.CurrentAngle
  MiniPF_LeftFlipperP.Rotz = MiniPF_LeftFlipper.CurrentAngle
  MiniPF_RightFlipperP.Rotz = MiniPF_RightFlipper.CurrentAngle
  FlipperLSh.RotZ = LeftFlipper.currentangle
  FlipperRSh.RotZ = RightFlipper.currentangle
  Gate1p.Roty = Gate1.CurrentAngle
  LeftGatep.Rotx = LeftGate.CurrentAngle
  RightSpinnerP.Rotx = sw39.CurrentAngle
  RightSpinnerP1.Rotx = sw39.CurrentAngle
  LeftGateSwitch.Rotx = LeftGateFl.CurrentAngle
  RollingSoundUpdate
End Sub

'//////////////////////////////////////////////////////////////////////
'// TIMERS
'//////////////////////////////////////////////////////////////////////


' The game timer interval is 10 ms
Sub GameTimer_Timer()
  Cor.Update            'update ball tracking
End Sub


' The frame timer interval is -1, so executes at the display frame rate
Sub FrameTimer_Timer()
  FlipperVisualUpdate       'update flipper shadows and primitives
  If DynamicBallShadowsOn Or AmbientBallShadowOn Then DynamicBSUpdate 'update ball shadows
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
LampTimer.Interval = 10 'lamp fading speed
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

Sub InitLamps()
  Dim x
  For x = 0 to 200
    LampState(x) = 0         ' current light state, independent of the fading level. 0 is off and 1 is on
    FadingLevel(x) = 4       ' used to track the fading state
    FlashSpeedUp(x) = 0.5    ' faster speed when turning on the flasher
    FlashSpeedDown(x) = 0.35 ' slower speed when turning off the flasher
    FlashMax(x) = 1          ' the maximum value when on, usually 1
    FlashMin(x) = 0          ' the minimum value when off, usually 0
    FlashLevel(x) = 0        ' the intensity of the flashers, usually from 0 to 1
  Next
End Sub

Sub UpdateLamps
  NFadeL 3, l3
  NFadeL 4, l4
  NFadeL 5, l5
  NFadeL 6, l6
  NFadeL 7, l7
  NFadeL 8, l8
  NFadeL 9, l9
  NFadeL 10, l10
  NFadeL 11, l11
  NFadeL 12, l12
  NFadeL 13, l13
  NFadeL 14, l14
  NFadeL 15, l15
  NFadeL 16, l16
  NFadeL 17, l17
  NFadeL 18, l18
  NFadeL 19, l19
  NFadeL 20, l20
  NFadeL 21, l21
  NFadeL 22, l22
  NFadeL 23, l23
  NFadeL 24, l24
  NFadeL 25, l25
  NFadeL 26, l26
  NFadeL 27, l27
  NFadeL 28, l28
  NFadeL 29, l29
  NFadeL 30, l30
  NFadeL 31, l31
  NFadeL 32, l32
  NFadeL 33, l33
  NFadeL 34, l34
  NFadeL 35, l35
  NFadeL 36, l36
  NFadeL 37, l37
  NFadeL 38, l38
  NFadeL 39, l39
  NFadeL 40, l40
  NFadeL 41, l41
  NFadeL 42, l42
  NFadeL 43, l43
  NFadeL 44, l44
  NFadeL 45, l45
  NFadeL 46, l46
  NFadeL 47, l47
  NFadeL 48, l48
  NFadeL 49, l49
  NFadeL 50, l50
  NFadeL 51, l51
  NFadeL 52, l52
  NFadeL 53, l53
  NFadeL 54, l54
  NFadeL 55, l55
  NFadeL 56, l56
  NFadeL 57, l57
  NFadeL 58, l58

  NFadeLm 61, l61
  NFadeL 61, l61a
  NFadeL 62, l62
  NFadeL 63, l63
  NFadeL 64, l64
  NFadeL 65, l65
  NFadeL 66, l66
  NFadeL 67, l67
  NFadeL 68, l68
  NFadeLm 69, cpl
  'FadeObj 69, CenterPost, "3D_Centerpost_3", "3D_Centerpost_2", "3D_Centerpost_1", "3D_Centerpost"
  NFadeL 70, l70
  Flashm 70, l70a
  Flash 70, l70a1



  NFadeL 199, F29
  NFadeL 190, F30
  NFadeL 191, F31

  'Mini Playfield LED Inserts
  NFadeL 125, LED1       'F
  NFadeL 124, LED2       'I
  NFadeL 123, LED3       'O
  NFadeL 122, LED4       'N
  NFadeL 121, LED5       'A

  NFadeL 98, LED6        'M
  NFadeL 99, LED7        'A
  NFadeL 100, LED8       'N

  NFadeL 114, LED9       'S
  NFadeL 89, LED10       'H
  NFadeL 90, LED11       'R
  NFadeL 91, LED12       'E
  NFadeL 92, LED13       'K

  NFadeL 108, LED14      'P
  NFadeL 107, LED15      'U
  NFadeL 106, LED16      'S
  NFadeL 105, LED17      'S

  NFadeL 84, LED18       'C
  NFadeL 83, LED19       'H
  NFadeL 82, LED20       'A
  NFadeL 81, LED21       'R
  NFadeL 97, LED22       'M

  'Flashers
  Flashm 76, Strip1
  NFadeLm 76, Strip1a
  Flashm 76, Strip2
  NFadeLm 76, Strip2a
  Flashm 76, Strip3
  NFadeLm 76, strip3a
  Flashm 76, Strip4
  NFadeLm 76, strip4a
  Flashm 76, Strip5
  NFadeLm 76, strip5a
  Flashm 76, Strip6
  NFadeLm 76, strip6a
  Flashm 76, Strip7
  NFadeLm 76, strip7a
  Flashm 76, Strip8
  NFadeLm 76, strip8a
  Flashm 76, Strip9
  NFadeLm 76, strip9a
  Flash 76, Strip10
  NFadeL 76, strip10a

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
    Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1            'wait
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



'******************************************************
'* GI  *******************************************
'******************************************************
'******************************************************

'******************************************************
'* GI  *******************************************
'******************************************************
'******************************************************

Set GICallback = GetRef("GIUpdate")

dim giTargetLvl, gilvl
Sub GIUpdate(no, Enabled)
Dim x
    For each x in GI
        x.State = ABS(Enabled)
    Next
  if Enabled Then
    giTargetLvl = 1
  Else
    giTargetLvl = 0
  end if

  GI_fade.enabled = true

End Sub

sub GI_fade_timer
' debug.print gilvl
  if gilvl < gitargetlvl then
    gilvl = gilvl + 0.3
    if gilvl > 1 then gilvl = 1
    if gilvl > gitargetlvl then gilvl = gitargetlvl
  elseif gilvl > gitargetlvl then
    gilvl = gilvl * 0.9 - 0.01
    if gilvl < 0 then gilvl = 0
    if gilvl < gitargetlvl then gilvl = gitargetlvl
  Else
    GI_fade.enabled = false
  end if

  PinCab_BladesON.opacity = 100 * gilvl

end sub

'*****************************************
' Ball Shadow
'*****************************************

Dim BallShadow
BallShadow = Array (BallShadow1, BallShadow2, BallShadow3)

Sub BallShadowUpdate()
  Dim BOT, b, shadowZ
  BOT = GetBalls

  ' render the shadow for each ball
  For b = 0 to UBound(BOT)
    If BOT(b).X < Table1.Width/2 Then
      BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) + 10
    Else
      BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) - 10
    End If

    If BOT(b).Z < 10 and BOT(b).Z>90 then BallShadow(b).Z = 1 else BallShadow(b).Z=BOT(b).z-20

    BallShadow(b).Y = BOT(b).Y + 40
    If BOT(b).Z > 20 Then
      BallShadow(b).visible = 1
    Else
      BallShadow(b).visible = 0
    End If
  Next
End Sub


'*****************************************
'    JP's VP10 Collision & Rolling Sounds
'*****************************************

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

Sub RollingSoundUpdate()
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


'******************************************************
'****  END BALL ROLLING AND DROP SOUNDS
'******************************************************


' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************


Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
  Dim tmp
  tmp = ball.x * 2 / table1.width-1
  If tmp> 0 Then
    Pan = Csng(tmp ^10)
  Else
    Pan = Csng(-((- tmp) ^10) )
  End If
End Function


Sub DropTargets_Hit (idx)

  SoundDropTargetDrop idx
  'Playsound SoundFX("Stern_Droptargethit",DOFContactors), 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
End Sub


'Sub Posts_Hit(idx)
' dim finalspeed
' finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
' If finalspeed > 16 then
'   PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
' End if
' If finalspeed >= 6 AND finalspeed <= 16 then
'   RandomSoundRubber()
' End If
'End Sub

'Sub RandomSoundRubber()
' Select Case Int(Rnd*3)+1
'   Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
'   Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
'   Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
' End Select
'End Sub


sub PlasticRampStart_hit: WireRampOn True : End Sub
Sub WirerollSND_Hit:WireRampOff : WireRampOn False : End Sub ' Playsound"Stern_Wireroll":End Sub
Sub BallDropSound_Hit: WireRampOff : RandomSoundWireRampStop BallDropSound :  End Sub 'Stopsound"Stern_Wireroll":Playsound"Stern_Balldrop":End Sub


'******************************************************
'*****   FLUPPER DOMES
'******************************************************
' Based on FlupperDoms2.2

' What you need in your table to use these flashers:
' Open this table and your table both in VPX
' Export all the materials domebasemat, Flashermaterial0 - 20 and import them in your table
' Export all textures (images) starting with the name "dome" and "ronddome" and import them into your table with the same names
' Export all textures (images) starting with the name "flasherbloom" and import them into your table with the same names
' Copy a set of 4 objects flasherbase, flasherlit, flasherlight and flasherflash from layer 7 to your table
' If you duplicate the four objects for a new flasher dome, be sure that they all end with the same number (in the 0-20 range)
' Copy the flasherbloom flashers from layer 10 to your table. you will need to make one per flasher dome that you plan to make
' Select the correct flasherbloom texture for each flasherbloom flasher, per flasher dome
' Copy the script below

' Place your flasher base primitive where you want the flasher located on your Table
' Then run InitFlasher in the script with the number of your flasher objects and the color of the flasher.  This will align the flasher object, light object, and
' flasher lit primitive.  It will also assign the appropriate flasher bloom images to the flasher bloom object.
'
' Example: InitFlasher 1, "green"
'
' Color Options: "blue", "green", "red", "purple", "yellow", "white", and "orange"

' You can use the RotateFlasher call to align the Rotz/ObjRotz of the flasher primitives with "handles".  Don't set those values in the editor,
' call the RotateFlasher sub instead (this call will likely crash VP if it's call for the flasher primitives without "handles")
'
' Example: RotateFlasher 1, 180     'where 1 is the flasher number and 180 is the angle of Z rotation

' For flashing the flasher use in the script: "ObjLevel(1) = 1 : FlasherFlash1_Timer"
' This should also work for flashers with variable flash levels from the rom, just use ObjLevel(1) = xx from the rom (in the range 0-1)
'
' Notes (please read!!):
' - Setting TestFlashers = 1 (below in the ScriptsDirectory) will allow you to see how the flasher objects are aligned (need the targetflasher image imported to your table)
' - The rotation of the primitives with "handles" is done with a script command, not on the primitive itself (see RotateFlasher below)
' - Color of the objects are set in the script, not on the primitive itself
' - Screws are optional to copy and position manually
' - If your table is not named "Table1" then change the name below in the script
' - Every flasher uses its own material (Flashermaterialxx), do not use it for anything else
' - Lighting > Bloom Strength affects how the flashers look, do not set it too high
' - Change RotY and RotX of flasherbase only when having a flasher something other then parallel to the playfield
' - Leave RotX of the flasherflash object to -45; this makes sure that the flash effect is visible in FS and DT
' - If you want to resize a flasher, be sure to resize flasherbase, flasherlit and flasherflash with the same percentage
' - If you think that the flasher effects are too bright, change flasherlightintensity and/or flasherflareintensity below

' Some more notes for users of the v1 flashers and/or JP's fading lights routines:
' - Delete all textures/primitives/script/materials in your table from the v1 flashers and scripts before you start; they don't mix well with v2
' - Remove flupperflash(m) routines if you have them; they do not work with this new script
' - Do not try to mix this v2 script with the JP fading light routine (that is making it too complicated), just use the example script below

' example script for rom based tables (non modulated):

' SolCallback(25)="FlashRed"
'
' Sub FlashRed(flstate)
' If Flstate Then
'   Objlevel(1) = 1 : FlasherFlash1_Timer
' End If
' End Sub

' example script for rom based tables (modulated):

' SolModCallback(25)="FlashRed"
'
' Sub FlashRed(level)
' Objlevel(1) = level/255 : FlasherFlash1_Timer
' End Sub


 Sub Flash1(Enabled)
  If Enabled Then
    ObjTargetLevel(1) = 1
  Else
    ObjTargetLevel(1) = 0
  End If
  FlasherFlash1_Timer
' Sound_Flash_Relay enabled, Flasherbase1
 End Sub

  Sub Flash2(Enabled)
  If Enabled Then
    ObjTargetLevel(2) = 1
  Else
    ObjTargetLevel(2) = 0
  End If
  FlasherFlash2_Timer
' Sound_Flash_Relay enabled, Flasherbase2
 End Sub

 Sub Flash3(Enabled)
  If Enabled Then
    ObjTargetLevel(3) = 1
  Else
    ObjTargetLevel(3) = 0
  End If
  FlasherFlash3_Timer
' Sound_Flash_Relay enabled, Flasherbase3
 End Sub

  Sub Flash4(Enabled)
  If Enabled Then
    ObjTargetLevel(4) = 1
  Else
    ObjTargetLevel(4) = 0
  End If
  FlasherFlash4_Timer
' Sound_Flash_Relay enabled, Flasherbase4
 End Sub

 Sub Flash5(Enabled)
  If Enabled Then
    ObjTargetLevel(5) = 1
  Else
    ObjTargetLevel(5) = 0
  End If
  FlasherFlash5_Timer
' Sound_Flash_Relay enabled, Flasherbase5
 End Sub

 Sub Flash6(Enabled)
  If Enabled Then
    ObjTargetLevel(6) = 1
  Else
    ObjTargetLevel(6) = 0
  End If
  FlasherFlash6_Timer
' Sound_Flash_Relay enabled, Flasherbase6
 End Sub



Dim TestFlashers, TableRef, FlasherLightIntensity, FlasherFlareIntensity, FlasherBloomIntensity, FlasherOffBrightness

                ' *********************************************************************
TestFlashers = 0        ' *** set this to 1 to check position of flasher object       ***
Set TableRef = Table1       ' *** change this, if your table has another name             ***
FlasherLightIntensity = 0.1   ' *** lower this, if the VPX lights are too bright (i.e. 0.1)   ***
FlasherFlareIntensity = 0.3   ' *** lower this, if the flares are too bright (i.e. 0.1)     ***
FlasherBloomIntensity = 0.2   ' *** lower this, if the blooms are too bright (i.e. 0.1)     ***
FlasherOffBrightness = 0.5    ' *** brightness of the flasher dome when switched off (range 0-2)  ***
                ' *********************************************************************

Dim ObjLevel(20), objbase(20), objlit(20), objflasher(20), objbloom(20), objlight(20), ObjTargetLevel(20)
'Dim tablewidth, tableheight : tablewidth = TableRef.width : tableheight = TableRef.height
'initialise the flasher color, you can only choose from "green", "red", "purple", "blue", "white" and "yellow"

InitFlasher 1, "yellow"
InitFlasher 2, "yellow"
InitFlasher 3, "green"
InitFlasher 4, "yellow"
InitFlasher 5, "white"
InitFlasher 6, "white"

' rotate the flasher with the command below (first argument = flasher nr, second argument = angle in degrees)
'RotateFlasher 1,17 : RotateFlasher 2,0 : RotateFlasher 3,90 : RotateFlasher 4,90
RotateFlasher 1,180
RotateFlasher 2,180
RotateFlasher 3,180
RotateFlasher 4,180
RotateFlasher 5,180


Sub InitFlasher(nr, col)
  ' store all objects in an array for use in FlashFlasher subroutine
  Set objbase(nr) = Eval("Flasherbase" & nr) : Set objlit(nr) = Eval("Flasherlit" & nr)
  Set objflasher(nr) = Eval("Flasherflash" & nr) : Set objlight(nr) = Eval("Flasherlight" & nr)
  Set objbloom(nr) = Eval("Flasherbloom" & nr)
  ' If the flasher is parallel to the playfield, rotate the VPX flasher object for POV and place it at the correct height
  If objbase(nr).RotY = 0 Then
    objbase(nr).ObjRotZ =  atn( (tablewidth/2 - objbase(nr).x) / (objbase(nr).y - tableheight*1.1)) * 180 / 3.14159
    objflasher(nr).RotZ = objbase(nr).ObjRotZ : objflasher(nr).height = objbase(nr).z + 40
  End If
  ' set all effects to invisible and move the lit primitive at the same position and rotation as the base primitive
  objlight(nr).IntensityScale = 0 : objlit(nr).visible = 0 : objlit(nr).material = "Flashermaterial" & nr
  objlit(nr).RotX = objbase(nr).RotX : objlit(nr).RotY = objbase(nr).RotY : objlit(nr).RotZ = objbase(nr).RotZ
  objlit(nr).ObjRotX = objbase(nr).ObjRotX : objlit(nr).ObjRotY = objbase(nr).ObjRotY : objlit(nr).ObjRotZ = objbase(nr).ObjRotZ
  objlit(nr).x = objbase(nr).x : objlit(nr).y = objbase(nr).y : objlit(nr).z = objbase(nr).z
  objbase(nr).BlendDisableLighting = FlasherOffBrightness

  'rothbauerw
  'Adjust the position of the flasher object to align with the flasher base.
  'Comment out these lines if you want to manually adjust the flasher object
  If objbase(nr).roty > 135 then
    objflasher(nr).y = objbase(nr).y + 50
    objflasher(nr).height = objbase(nr).z + 20
  Else
    objflasher(nr).y = objbase(nr).y + 20
    objflasher(nr).height = objbase(nr).z + 50
  End If
  objflasher(nr).x = objbase(nr).x

  'rothbauerw
  'Adjust the position of the light object to align with the flasher base.
  'Comment out these lines if you want to manually adjust the flasher object
  objlight(nr).x = objbase(nr).x
  objlight(nr).y = objbase(nr).y
  objlight(nr).bulbhaloheight = objbase(nr).z -10

  'rothbauerw
  'Assign the appropriate bloom image basked on the location of the flasher base
  'Comment out these lines if you want to manually assign the bloom images
  dim xthird, ythird
  xthird = tablewidth/3
  ythird = tableheight/3

  If objbase(nr).x >= xthird and objbase(nr).x <= xthird*2 then
    objbloom(nr).imageA = "flasherbloomCenter"
    objbloom(nr).imageB = "flasherbloomCenter"
  elseif objbase(nr).x < xthird and objbase(nr).y < ythird then
    objbloom(nr).imageA = "flasherbloomUpperLeft"
    objbloom(nr).imageB = "flasherbloomUpperLeft"
  elseif  objbase(nr).x > xthird*2 and objbase(nr).y < ythird then
    objbloom(nr).imageA = "flasherbloomUpperRight"
    objbloom(nr).imageB = "flasherbloomUpperRight"
  elseif objbase(nr).x < xthird and objbase(nr).y < ythird*2 then
    objbloom(nr).imageA = "flasherbloomCenterLeft"
    objbloom(nr).imageB = "flasherbloomCenterLeft"
  elseif  objbase(nr).x > xthird*2 and objbase(nr).y < ythird*2 then
    objbloom(nr).imageA = "flasherbloomCenterRight"
    objbloom(nr).imageB = "flasherbloomCenterRight"
  elseif objbase(nr).x < xthird and objbase(nr).y < ythird*3 then
    objbloom(nr).imageA = "flasherbloomLowerLeft"
    objbloom(nr).imageB = "flasherbloomLowerLeft"
  elseif  objbase(nr).x > xthird*2 and objbase(nr).y < ythird*3 then
    objbloom(nr).imageA = "flasherbloomLowerRight"
    objbloom(nr).imageB = "flasherbloomLowerRight"
  end if

  ' set the texture and color of all objects
  select case objbase(nr).image
    Case "dome2basewhite" : objbase(nr).image = "dome2base" & col : objlit(nr).image = "dome2lit" & col :
    Case "ronddomebasewhite" : objbase(nr).image = "ronddomebase" & col : objlit(nr).image = "ronddomelit" & col
    Case "domeearbasewhite" : objbase(nr).image = "domeearbase" & col : objlit(nr).image = "domeearlit" & col
  end select
  If TestFlashers = 0 Then objflasher(nr).imageA = "domeflashwhite" : objflasher(nr).visible = 0 : End If
  select case col
    Case "blue" :   objlight(nr).color = RGB(4,120,255) : objflasher(nr).color = RGB(200,255,255) : objbloom(nr).color = RGB(4,120,255) : objlight(nr).intensity = 5000
    Case "green" :  objlight(nr).color = RGB(12,255,4) : objflasher(nr).color = RGB(12,255,4) : objbloom(nr).color = RGB(12,255,4)
    Case "red" :    objlight(nr).color = RGB(255,32,4) : objflasher(nr).color = RGB(255,32,4) : objbloom(nr).color = RGB(255,32,4)
    Case "purple" : objlight(nr).color = RGB(230,49,255) : objflasher(nr).color = RGB(255,64,255) : objbloom(nr).color = RGB(230,49,255)
    Case "yellow" : objlight(nr).color = RGB(200,173,25) : objflasher(nr).color = RGB(255,200,50) : objbloom(nr).color = RGB(200,173,25)
    Case "white" :  objlight(nr).color = RGB(255,240,150) : objflasher(nr).color = RGB(100,86,59) : objbloom(nr).color = RGB(255,240,150)
    Case "orange" :  objlight(nr).color = RGB(255,70,0) : objflasher(nr).color = RGB(255,70,0) : objbloom(nr).color = RGB(255,70,0)
  end select
  objlight(nr).colorfull = objlight(nr).color
  If TableRef.ShowDT and ObjFlasher(nr).RotX = -45 Then
    objflasher(nr).height = objflasher(nr).height - 20 * ObjFlasher(nr).y / tableheight
    ObjFlasher(nr).y = ObjFlasher(nr).y + 10
  End If
End Sub

Sub RotateFlasher(nr, angle) : angle = ((angle + 360 - objbase(nr).ObjRotZ) mod 180)/30 : objbase(nr).showframe(angle) : objlit(nr).showframe(angle) : End Sub

Sub FlashFlasher(nr)
  If not objflasher(nr).TimerEnabled Then objflasher(nr).TimerEnabled = True : objflasher(nr).visible = 1 : objbloom(nr).visible = 1 : objlit(nr).visible = 1 : End If
  objflasher(nr).opacity = 1000 *  FlasherFlareIntensity * ObjLevel(nr)^2.5
  objbloom(nr).opacity = 100 *  FlasherBloomIntensity * ObjLevel(nr)^2.5
  objlight(nr).IntensityScale = 0.5 * FlasherLightIntensity * ObjLevel(nr)^3
  objbase(nr).BlendDisableLighting =  FlasherOffBrightness + 10 * ObjLevel(nr)^3
  objlit(nr).BlendDisableLighting = 10 * ObjLevel(nr)^2
  UpdateMaterial "Flashermaterial" & nr,0,0,0,0,0,0,ObjLevel(nr),RGB(255,255,255),0,0,False,True,0,0,0,0
  if round(ObjTargetLevel(nr),1) > round(ObjLevel(nr),1) Then
    ObjLevel(nr) = ObjLevel(nr) + 0.3
    if ObjLevel(nr) > 1 then ObjLevel(nr) = 1
  Elseif round(ObjTargetLevel(nr),1) < round(ObjLevel(nr),1) Then
    ObjLevel(nr) = ObjLevel(nr) * 0.85 - 0.01
    if ObjLevel(nr) < 0 then ObjLevel(nr) = 0
  Else
    ObjLevel(nr) = round(ObjTargetLevel(nr),1)
    objflasher(nr).TimerEnabled = False
  end if
  'ObjLevel(nr) = ObjLevel(nr) * 0.9 - 0.01
  If ObjLevel(nr) < 0 Then objflasher(nr).TimerEnabled = False : objflasher(nr).visible = 0 : objbloom(nr).visible = 0 : objlit(nr).visible = 0 : End If
End Sub

Sub FlasherFlash1_Timer() : FlashFlasher(1) : End Sub
Sub FlasherFlash2_Timer() : FlashFlasher(2) : End Sub
Sub FlasherFlash3_Timer() : FlashFlasher(3) : End Sub
Sub FlasherFlash4_Timer() : FlashFlasher(4) : End Sub
Sub FlasherFlash5_Timer() : FlashFlasher(5) : End Sub
Sub FlasherFlash6_Timer() : FlashFlasher(6) : End Sub


'******************************************************
'******  END FLUPPER DOMES
'******************************************************

'//////////////////////////////////////////////////////////////////////
'// Dynamic Ball Shadows
'//////////////////////////////////////////////////////////////////////

Const fovY          = 0   'Offset y position under ball to account for layback or inclination (more pronounced need further back)
Const DynamicBSFactor     = 0.95  '0 to 1, higher is darker
Const AmbientBSFactor     = 0.7 '0 to 1, higher is darker
Const AmbientMovement   = 2   '1 to 4, higher means more movement as the ball moves left and right
Const Wideness        = 20  'Sets how wide the dynamic ball shadows can get (20 +5 thinness should be most realistic for a 50 unit ball)
Const Thinness        = 5   'Sets minimum as ball moves away from source


' *** This segment goes within the RollingUpdate sub, so that if Ambient...=0 and Dynamic...=0 the entire DynamicBSUpdate sub can be skipped for max performance
' ' stop the sound of deleted balls
' For b = UBound(BOT) + 1 to tnob
'   If AmbientBallShadowOn = 0 Then BallShadowA(b).visible = 0
'   ...rolling(b) = False
'   ...StopSound("BallRoll_" & b)
' Next
'
'...rolling and drop sounds...

'   If DropCount(b) < 5 Then
'     DropCount(b) = DropCount(b) + 1
'   End If
'
'   ' "Static" Ball Shadows
'   If AmbientBallShadowOn = 0 Then
'     If BOT(b).Z > 30 Then
'       BallShadowA(b).height=BOT(b).z - BallSize/4   'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping
'     Else
'       BallShadowA(b).height=BOT(b).z - BallSize/2 + 5
'     End If
'     BallShadowA(b).Y = BOT(b).Y + Ballsize/5 + fovY
'     BallShadowA(b).X = BOT(b).X
'     BallShadowA(b).visible = 1
'   End If

' *** Required Functions, enable these if they are not already present elswhere in your table
Function DistanceFast(x, y)
  dim ratio, ax, ay
  ax = abs(x)         'Get absolute value of each vector
  ay = abs(y)
  ratio = 1 / max(ax, ay)   'Create a ratio
  ratio = ratio * (1.29289 - (ax + ay) * ratio * 0.29289)
  if ratio > 0 then     'Quickly determine if it's worth using
    DistanceFast = 1/ratio
  Else
    DistanceFast = 0
  End if
end Function

Function max(a,b)
  if a > b then
    max = a
  Else
    max = b
  end if
end Function

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
'
'Function AnglePP(ax,ay,bx,by)
' AnglePP = Atn2((by - ay),(bx - ax))*180/PI
'End Function

'****** End Part B:  Code and Functions ******


'****** Part C:  The Magic ******
Dim sourcenames, currentShadowCount, DSSources(44), numberofsources, numberofsources_hold
sourcenames = Array ("","","","","","","","","","","","")
currentShadowCount = Array (0,0,0,0,0,0,0,0,0,0,0,0)

' *** Trim or extend these to match the number of balls/primitives/flashers on the table!
dim objrtx1(12), objrtx2(12)
dim objBallShadow(12)
Dim BallShadowA
BallShadowA = Array (BallShadowA0,BallShadowA1,BallShadowA2,BallShadowA3,BallShadowA4,BallShadowA5,BallShadowA6,BallShadowA7,BallShadowA8,BallShadowA9,BallShadowA10,BallShadowA11)

DynamicBSInit

sub DynamicBSInit()
  Dim iii, source

  for iii = 0 to tnob                 'Prepares the shadow objects before play begins
    Set objrtx1(iii) = Eval("RtxBallShadow" & iii)
    objrtx1(iii).material = "RtxBallShadow" & iii
    objrtx1(iii).z = iii/1000 + 0.01
    objrtx1(iii).visible = 0

    Set objrtx2(iii) = Eval("RtxBall2Shadow" & iii)
    objrtx2(iii).material = "RtxBallShadow2_" & iii
    objrtx2(iii).z = (iii)/1000 + 0.02
    objrtx2(iii).visible = 0

    currentShadowCount(iii) = 0

    Set objBallShadow(iii) = Eval("BallShadow" & iii)
    objBallShadow(iii).material = "BallShadow" & iii
    UpdateMaterial objBallShadow(iii).material,1,0,0,0,0,0,AmbientBSFactor,RGB(0,0,0),0,0,False,True,0,0,0,0
    objBallShadow(iii).Z = iii/1000 + 0.04
    objBallShadow(iii).visible = 0

    BallShadowA(iii).Opacity = 100*AmbientBSFactor
    BallShadowA(iii).visible = 0
  Next

  iii = 0

  For Each Source in DynamicSources
    DSSources(iii) = Array(Source.x, Source.y)
    iii = iii + 1
  Next
  numberofsources = iii
  numberofsources_hold = iii
end sub


Sub DynamicBSUpdate
  Dim falloff:  falloff = 150     'Max distance to light sources, can be changed if you have a reason
  Dim ShadowOpacity, ShadowOpacity2
  Dim s, Source, LSd, currentMat, AnotherSource, BOT, iii
  BOT = GetBalls

  'Hide shadow of deleted balls
  For s = UBound(BOT) + 1 to tnob
    objrtx1(s).visible = 0
    objrtx2(s).visible = 0
    objBallShadow(s).visible = 0
    BallShadowA(s).visible = 0
  Next

  If UBound(BOT) < lob Then Exit Sub    'No balls in play, exit

'The Magic happens now
  For s = lob to UBound(BOT)

' *** Normal "ambient light" ball shadow
  'Layered from top to bottom. If you had an upper pf at for example 80 and ramps even above that, your segments would be z>110; z<=110 And z>100; z<=100 And z>30; z<=30 And z>20; Else invisible

    If AmbientBallShadowOn = 1 Then     'Primitive shadow on playfield, flasher shadow in ramps
      If BOT(s).Z > 30 Then             'The flasher follows the ball up ramps while the primitive is on the pf
        If BOT(s).X < tablewidth/2 Then
          objBallShadow(s).X = ((BOT(s).X) - (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) + 5
        Else
          objBallShadow(s).X = ((BOT(s).X) + (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) - 5
        End If
        objBallShadow(s).Y = BOT(s).Y + BallSize/10 + fovY
        objBallShadow(s).visible = 1

        BallShadowA(s).X = BOT(s).X
        BallShadowA(s).Y = BOT(s).Y + BallSize/5 + fovY
        BallShadowA(s).height=BOT(s).z - BallSize/4   'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping
        BallShadowA(s).visible = 1
      Elseif BOT(s).Z <= 30 And BOT(s).Z > 20 Then  'On pf, primitive only
        objBallShadow(s).visible = 1
        If BOT(s).X < tablewidth/2 Then
          objBallShadow(s).X = ((BOT(s).X) - (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) + 5
        Else
          objBallShadow(s).X = ((BOT(s).X) + (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) - 5
        End If
        objBallShadow(s).Y = BOT(s).Y + fovY
        BallShadowA(s).visible = 0
      Else                      'Under pf, no shadows
        objBallShadow(s).visible = 0
        BallShadowA(s).visible = 0
      end if

    Elseif AmbientBallShadowOn = 2 Then   'Flasher shadow everywhere
      If BOT(s).Z > 30 Then             'In a ramp
        BallShadowA(s).X = BOT(s).X
        BallShadowA(s).Y = BOT(s).Y + BallSize/5 + fovY
        BallShadowA(s).height=BOT(s).z - BallSize/4   'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping
        BallShadowA(s).visible = 1
      Elseif BOT(s).Z <= 30 And BOT(s).Z > 20 Then  'On pf
        BallShadowA(s).visible = 1
        If BOT(s).X < tablewidth/2 Then
          BallShadowA(s).X = ((BOT(s).X) - (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) + 5
        Else
          BallShadowA(s).X = ((BOT(s).X) + (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) - 5
        End If
        BallShadowA(s).Y = BOT(s).Y + Ballsize/10 + fovY
        BallShadowA(s).height=BOT(s).z - BallSize/2 + 5
      Else                      'Under pf
        BallShadowA(s).visible = 0
      End If
    End If

' *** Dynamic shadows
    If DynamicBallShadowsOn Then
      If BOT(s).Z < 30 Then 'And BOT(s).Y < (TableHeight - 200) Then 'Or BOT(s).Z > 105 Then    'Defining when and where (on the table) you can have dynamic shadows
        For iii = 0 to numberofsources - 1
          LSd=DistanceFast((BOT(s).x-DSSources(iii)(0)),(BOT(s).y-DSSources(iii)(1))) 'Calculating the Linear distance to the Source
          If LSd < falloff And gilvl > 0 Then               'If the ball is within the falloff range of a light and light is on (we will set numberofsources to 0 when GI is off)
            currentShadowCount(s) = currentShadowCount(s) + 1   'Within range of 1 or 2
            if currentShadowCount(s) = 1 Then           '1 dynamic shadow source
              sourcenames(s) = iii
              currentMat = objrtx1(s).material
              objrtx2(s).visible = 0 : objrtx1(s).visible = 1 : objrtx1(s).X = BOT(s).X : objrtx1(s).Y = BOT(s).Y + fovY
  '           objrtx1(s).Z = BOT(s).Z - 25 + s/1000 + 0.01            'Uncomment if you want to add shadows to an upper/lower pf
              objrtx1(s).rotz = AnglePP(DSSources(iii)(0), DSSources(iii)(1), BOT(s).X, BOT(s).Y) + 90
              ShadowOpacity = (falloff-LSd)/falloff                 'Sets opacity/darkness of shadow by distance to light
              objrtx1(s).size_y = Wideness*ShadowOpacity+Thinness           'Scales shape of shadow with distance/opacity
              UpdateMaterial currentMat,1,0,0,0,0,0,ShadowOpacity*DynamicBSFactor^2,RGB(0,0,0),0,0,False,True,0,0,0,0
              If AmbientBallShadowOn = 1 Then
                currentMat = objBallShadow(s).material                  'Brightens the ambient primitive when it's close to a light
                UpdateMaterial currentMat,1,0,0,0,0,0,AmbientBSFactor*(1-ShadowOpacity),RGB(0,0,0),0,0,False,True,0,0,0,0
              Else
                BallShadowA(s).Opacity = 100*AmbientBSFactor*(1-ShadowOpacity)
              End If

            Elseif currentShadowCount(s) = 2 Then
                                  'Same logic as 1 shadow, but twice
              currentMat = objrtx1(s).material
              AnotherSource = sourcenames(s)
              objrtx1(s).visible = 1 : objrtx1(s).X = BOT(s).X : objrtx1(s).Y = BOT(s).Y + fovY
  '           objrtx1(s).Z = BOT(s).Z - 25 + s/1000 + 0.01              'Uncomment if you want to add shadows to an upper/lower pf
              objrtx1(s).rotz = AnglePP(DSSources(AnotherSource)(0),DSSources(AnotherSource)(1), BOT(s).X, BOT(s).Y) + 90
              ShadowOpacity = (falloff-DistanceFast((BOT(s).x-DSSources(AnotherSource)(0)),(BOT(s).y-DSSources(AnotherSource)(1))))/falloff
              objrtx1(s).size_y = Wideness*ShadowOpacity+Thinness
              UpdateMaterial currentMat,1,0,0,0,0,0,ShadowOpacity*DynamicBSFactor^3,RGB(0,0,0),0,0,False,True,0,0,0,0

              currentMat = objrtx2(s).material
              objrtx2(s).visible = 1 : objrtx2(s).X = BOT(s).X : objrtx2(s).Y = BOT(s).Y + fovY
  '           objrtx2(s).Z = BOT(s).Z - 25 + s/1000 + 0.02              'Uncomment if you want to add shadows to an upper/lower pf
              objrtx2(s).rotz = AnglePP(DSSources(iii)(0), DSSources(iii)(1), BOT(s).X, BOT(s).Y) + 90
              ShadowOpacity2 = (falloff-LSd)/falloff
              objrtx2(s).size_y = Wideness*ShadowOpacity2+Thinness
              UpdateMaterial currentMat,1,0,0,0,0,0,ShadowOpacity2*DynamicBSFactor^3,RGB(0,0,0),0,0,False,True,0,0,0,0
              If AmbientBallShadowOn = 1 Then
                currentMat = objBallShadow(s).material                  'Brightens the ambient primitive when it's close to a light
                UpdateMaterial currentMat,1,0,0,0,0,0,AmbientBSFactor*(1-max(ShadowOpacity,ShadowOpacity2)),RGB(0,0,0),0,0,False,True,0,0,0,0
              Else
                BallShadowA(s).Opacity = 100*AmbientBSFactor*(1-max(ShadowOpacity,ShadowOpacity2))
              End If
            end if
          Else
            currentShadowCount(s) = 0
            BallShadowA(s).Opacity = 100*AmbientBSFactor
          End If
        Next
      Else                  'Hide dynamic shadows everywhere else
        objrtx2(s).visible = 0 : objrtx1(s).visible = 0
      End If
    End If
  Next
End Sub

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


'///////////////////////////  DROP TARGET HIT SOUNDS  ///////////////////////////

Sub RandomSoundDropTargetReset(obj)
  PlaySoundAtLevelStatic SoundFX("Drop_Target_Reset_" & Int(Rnd*6)+1,DOFContactors), 1, obj
End Sub

Sub SoundDropTargetDrop(obj)
  PlaySoundAtLevelStatic ("Drop_Target_Down_" & Int(Rnd*6)+1), 200, obj
End Sub

'/////////////////////////////  GI AND FLASHER RELAYS  ////////////////////////////

Const RelayFlashSoundLevel = 0.315                  'volume level; range [0, 1];
Const RelayGISoundLevel = 1.05                  'volume level; range [0, 1];

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

Sub RandomSoundWireRampStop(obj)
  Select Case Int(rnd*6)
    Case 0: PlaySoundAtVol "wireramp_stop", obj, 0.2*volumedial
    Case 1: PlaySoundAtVol "wireramp_stop2", obj, 0.2*volumedial
    Case 2: PlaySoundAtVol "wireramp_stop3", obj, 0.2*volumedial
    Case 3: PlaySoundAtVol "wireramp_stop", obj, 0.1*volumedial
    Case 4: PlaySoundAtVol "wireramp_stop2", obj, 0.1*volumedial
    Case 5: PlaySoundAtVol "wireramp_stop3", obj, 0.1*volumedial
  End Select
End Sub

Sub RandomSoundRampStop(obj)
  Select Case Int(rnd*3)
    Case 0: PlaySoundAtLevelActiveBall ("Rubber_Strong_1"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
    Case 1: PlaySoundAtLevelActiveBall ("Rubber_Strong_2"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
    Case 2: PlaySoundAtLevelActiveBall ("Rubber_1_Hard"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
  End Select
End Sub


'/////////////////////////////////////////////////////////////////
'         End Mechanical Sounds
'/////////////////////////////////////////////////////////////////

'******************************************************
'****  FLEEP MECHANICAL SOUNDS
'******************************************************

'******************************************************
'**** RAMP ROLLING SFX
'******************************************************

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



'******************************************************
'**** END RAMP ROLLING SFX
'******************************************************


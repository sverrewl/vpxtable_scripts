Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01530000","SEGA.vbs",3.1
'********************************************
'**     Game Specific Code Starts Here     **
'********************************************

Const cGameName="startrp",UseSolenoids=1,UseLamps=1,UseGI=0,UseSync=1
Const SSolenoidOn="SolOn",SSolenoidOff="SolOff",SFlipperOn="FlipperUp",SFlipperOff="FlipperDown",SCoin="Coin3"

'*********************************
' FLASHER and GI INTENSITY
Dim bigflashers
Dim BrightGI
Dim Flasherspeed

bigflashers = 1 ' 1= Big flasher images and flashes, 0 = only basic flashers
BrightGI = 0 ' 0 = default GI brightness, 1 = for brighter GI (if you want to run the table darker
Flasherspeed = 1 ' 1= faster flasher speed up / down, 0 = Slower flashers

'*********************************



'**************************************
'**     Bind Events To Solenoids     **
'**************************************

SolCallBack(1)="bsTrough.SolOut"
SolCallBack(2)= "Auto_Plunger"
SolCallBack(3)="bsBottom.SolOut"
SolCallBack(4)="bsTop.SolOut"
SolCallback(5)="LMag.MagnetOn="
SolCallback(6)="RMag.MagnetOn="
SolCallBack(7)="SolBrainBug"
SolCallBack(8)="vpmSolSound SoundFX(""Knocker"",DOFKnocker)," 'dispenser - Set adjustment to OFF to get knocker sound
'SolCallBack(9)="vpmSolSound SoundFX(""Jet3"",DOFContactors),"
'SolCallBack(10)="vpmSolSound SoundFX(""Jet3"",DOFContactors),"
'SolCallBack(11)="vpmSolSound SoundFX(""Jet3"",DOFContactors),"
SolCallBack(12)="vpmSolSound SoundFX(""Sling"",DOFContactors),"
SolCallBack(13)="vpmSolSound SoundFX(""Sling"",DOFContactors),"
'SolCallBack(14)="vpmSolFlipper Flipper1,Nothing,"
SolCallback(14) = "SolMiniFlipper"
'SolCallback(17)="SolMotorOn"
SolCallBack(23)="SolBrainFlash"
'SolCallBack(24)='optional coil
SolCallBack(25)="Sol25"     'red x 4
SolCallBack(26)="Sol26"     'yellow x 4
SolCallBack(27)="Sol27"     'green x 4
SolCallBack(28)="Sol28"     'blue x 4
SolCallBack(29)="SolWarrior"        'flash multiball x 4
SolCallBack(30)="vpmFlasher Array(F30A,F30B,F30C),"     'left ramp x 4
SolCallBack(31)="vpmFlasher Array(F31A,F31B,F31C),"     'right ramp x 4
SolCallBack(32)="vpmFlasher Array(F32A,F32B),"        'pops x 2
'SolCallBack(sLLFlipper)="vpmSolFlipper LeftFlipper,Nothing,"
'SolCallBack(sLRFlipper)="vpmSolFlipper RightFlipper,Nothing,"

SolCallBack(25)="SetFlash 81, "     'red x 4
SolCallBack(26)="SetFlash 82, "   'yellow x 4
SolCallBack(28)="SetFlash 83, "     'blue x 4
SolCallBack(27)="SetFlash 84, "     'blue x 4
SolCallBack(29)="SolWarrior"        'flash multiball x 4
SolCallBack(30)="SetFlash 86, "     'left ramp x 4
SolCallBack(31)="SetFlash 87, "     'right ramp x 4
'SolCallBack(32)="SetFlasher 88, "        'pops x 2


SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
  If Enabled Then
    PlaySoundAtVol SoundFX("FlipperUp",DOFFlippers), LeftFlipper, 1
    LeftFlipper.RotateToEnd
   Else
    PlaySoundAtVol SoundFX("FlipperDown",DOFFlippers), LeftFlipper, 1
    LeftFlipper.RotateToStart
   End If
End Sub

Sub SolRFlipper(Enabled)
   If Enabled Then
    PlaySoundAtVol SoundFX("FlipperUp",DOFFlippers), RightFlipper, 1
    RightFlipper.RotateToEnd
   Else
    PlaySoundAtVol SoundFX("FlipperDown",DOFFlippers), RightFlipper, 1
    RightFlipper.RotateToStart
   End If
End Sub

Sub SolMiniFlipper(Enabled)
  If Enabled Then
    PlaySoundAtVol SoundFX("FlipperUp",DOFFlippers), Flipper1, 1
    Flipper1.RotateToEnd
  Else
    PlaySoundAtVol SoundFX("FlipperDown",DOFFlippers), Flipper1, 1
    Flipper1.RotateToStart
  End If
End Sub

'**********************************************************************************************************

'Solenoid Controlled toys
'**********************************************************************************************************


Sub Auto_Plunger(Enabled) 'plunger
  If Enabled Then
    PlungerIM.AutoFire
  End If
End Sub


Dim BugLit
BugLit=False

Sub SolWarrior(Enabled)
    ' TODO - lighhts on bug here!!
  BugLit = Enabled

  ' warrior bug flasher
  if Buglit = true Then
    setflash 199, true
  Else
    setflash 199, false
  End If



End Sub

Dim BrainState
BrainState=False

Sub SolBrainFlash(Enabled)
  If Enabled Then
    BrainState=True
    BrainFrame=BrainFrame+1
  Else
    BrainState=False
    BrainFrame=BrainFrame-1
  End If
  EMReel1.SetValue BrainFrame
End Sub



Sub Sol25(Enabled)
  If Enabled Then
FlasherLight1.State=LightStateOn
FlasherLight2.State=LightStateOn
FlasherLight3.State=LightStateOn
FlasherLight4.State=LightStateOn
    F1A.IsDropped=0
    F1B.IsDropped=0
    F1C.IsDropped=0
    F1D.IsDropped=0
  Else
FlasherLight1.State=LightStateOff
FlasherLight2.State=LightStateOff
FlasherLight3.State=LightStateOff
FlasherLight4.State=LightStateOff
    F1A.IsDropped=1
    F1B.IsDropped=1
    F1C.IsDropped=1
    F1D.IsDropped=1
  End If
End Sub

Sub Sol26(Enabled)
  If Enabled Then
FlasherLight5.State=LightStateOn
FlasherLight6.State=LightStateOn
FlasherLight7.State=LightStateOn
FlasherLight8.State=LightStateOn
    F2A.IsDropped=0
    F2B.IsDropped=0
    F2C.IsDropped=0
    F2D.IsDropped=0
  Else
FlasherLight5.State=LightStateOff
FlasherLight6.State=LightStateOff
FlasherLight7.State=LightStateOff
FlasherLight8.State=LightStateOff
    F2A.IsDropped=1
    F2B.IsDropped=1
    F2C.IsDropped=1
    F2D.IsDropped=1
  End If

End Sub

Sub Sol27(Enabled)
  If Enabled Then
FlasherLight9.State=LightStateOn
FlasherLight10.State=LightStateOn
FlasherLight11.State=LightStateOn
FlasherLight12.State=LightStateOn
    F3A.IsDropped=0
    F3B.IsDropped=0
    F3C.IsDropped=0
    F3D.IsDropped=0
  Else
FlasherLight9.State=LightStateOff
FlasherLight10.State=LightStateOff
FlasherLight11.State=LightStateOff
FlasherLight12.State=LightStateOff
    F3A.IsDropped=1
    F3B.IsDropped=1
    F3C.IsDropped=1
    F3D.IsDropped=1
  End If
End Sub

Sub Sol28(Enabled)
  If Enabled Then
FlasherLight13.State=LightStateOn
FlasherLight14.State=LightStateOn
FlasherLight15.State=LightStateOn
FlasherLight16.State=LightStateOn
    F4A.IsDropped=0
    F4B.IsDropped=0
    F4C.IsDropped=0
    F4D.IsDropped=0
  Else
FlasherLight13.State=LightStateOff
FlasherLight14.State=LightStateOff
FlasherLight15.State=LightStateOff
FlasherLight16.State=LightStateOff
    F4A.IsDropped=1
    F4B.IsDropped=1
    F4C.IsDropped=1
    F4D.IsDropped=1
  End If
End Sub

 Dim BrainE,BrainFrame
 BrainE=False:BrainFrame=0

Sub SolBrainBug(Enabled)
If Enabled Then
  EMReel1.TimerEnabled=0
  EMReel1.TimerEnabled=1
  Brain.IsDropped=0
  BrainBottom.IsDropped=0
  BrainE=True
  EMReel1_Timer
    pWorm.TransZ = +80
Else
  EMReel1.TimerEnabled=0
  EMReel1.TimerEnabled=1
  BrainE=False
  EMReel1_Timer
    pWorm.TransZ = -80
End If
End Sub

Sub EMReel1_Timer
  If BrainE=True Then
    If BrainFrame<4 Then
      BrainFrame=BrainFrame+2
    Else
      EMReel1.TimerEnabled=0
      Controller.Switch(37)=1
      Controller.Switch(39)=1
    End If
  Else
    If BrainFrame>1 Then
      BrainFrame=BrainFrame-2
    Else
      EMReel1.TimerEnabled=0
      Controller.Switch(37)=0
      Controller.Switch(39)=0
      Brain.IsDropped=1
      BrainBottom.IsDropped=1
    End If
  End If
  EMReel1.SetValue BrainFrame
 End Sub

Sub Table1_KeyDown(ByVal KeyCode)
  If KeyCode=RightMagnaSave Then Controller.Switch(88)=1
  If KeyCode=LeftFlipperKey Then Controller.Switch(63)=1
  If KeyCode=RightFlipperKey Then Controller.Switch(64)=1
  If KeyDownHandler(KeyCode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
  If KeyCode=RightMagnaSave Then Controller.Switch(88)=0
  If KeyCode=LeftFlipperKey Then Controller.Switch(63)=0
  If KeyCode=RightFlipperKey Then Controller.Switch(64)=0
  If KeyUpHandler(KeyCode) Then Exit Sub
End Sub

ExtraKeyHelp=KeyUpperRight&vbTab&"Right Mini Flipper"

Dim bsTrough,bsBottom,bsTop,LMag,RMag,BugWalls,BugTargets,mBug
BugTargets=Array(BugT6,BugT5,BugT4,BugT3,BugT2,BugT1,BugT6,BugT5,BugT4,BugT3,BugT2,BugT1)

 Sub SetDisplayToElement(Element)
  If Controller.Version<="01500000" Then
    ' forget it, version is to old
    Exit Sub
  End If
    Dim playerRect
  playerRect=Controller.GetClientRect(GetPlayerHwnd)
  Dim playerWidth, playerHeight
  playerWidth=playerRect(2)-playerRect(0)
  playerHeight=playerRect(3)-playerRect(1)
  Dim Game
  Set Game=Controller.Game
  Dim x,y
    x=Element.x*playerWidth/1000.0-1
  y=Element.y*playerHeight/750.0-1
  Game.Settings.SetDisplayPosition x,y,GetPlayerHwnd
  Set Game=nothing
End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)
  If KeyCode=RightMagnaSave Then Controller.Switch(88)=1
  If KeyCode=LeftFlipperKey Then Controller.Switch(63)=1
  If KeyCode=RightFlipperKey Then Controller.Switch(64)=1
  If keycode = PlungerKey Then Plunger.Pullback:playsoundAtVol "plungerpull", Plunger, 1
  If KeyDownHandler(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
  If KeyCode=RightMagnaSave Then Controller.Switch(88)=0
  If KeyCode=LeftFlipperKey Then Controller.Switch(63)=0
  If KeyCode=RightFlipperKey Then Controller.Switch(64)=0
  If keycode = PlungerKey Then Plunger.Fire:PlaySoundAtVol "plunger", Plunger, 1
  If KeyUpHandler(keycode) Then Exit Sub
End Sub

  ' Impulse Plunger
  dim plungerIM

  Const IMPowerSetting = 40 ' Plunger Power
  Const IMTime = 0.6        ' Time in seconds for Full Plunge
  Set plungerIM = New cvpmImpulseP
  With plungerIM
    .InitImpulseP swPlunger, IMPowerSetting, IMTime
    .Random 0.3
    .Switch 16
    .InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
    .CreateEvents "plungerIM"
  End With

'**********************************************************************************************************

 ' Drain hole and kickers
Sub Drain_Hit:bsTrough.addball me : playsoundAtVol "drain", Drain, 1: End Sub
Sub sw45_Hit:bsBottom.AddBall 0 : playsoundAtVol "popper_ball" , sw45, 1: End Sub

'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************

Sub Table1_Init
    if table1.VersionMinor < 6 AND table1.VersionMajor = 10 then MsgBox "This table requires VPX 10.6, you have " & table1.VersionMajor & "." & table.VersionMinor
  if VPinMAMEDriverVer < 3.57 then MsgBox "This table requires core.vbs 3.57 or higher, which is included with VPX 10.6.  You have " & VPinMAMEDriverVer & ". Be sure scripts folder is up to date, and that there are no old .vbs files in your table folder."


  LeftGuide.IsDropped=1 'disables collision detection
  RightGuide.IsDropped=1 'disables collision detection
FlasherLight1.State=LightStateOff
FlasherLight2.State=LightStateOff
FlasherLight3.State=LightStateOff
FlasherLight4.State=LightStateOff
FlasherLight5.State=LightStateOff
FlasherLight6.State=LightStateOff
FlasherLight7.State=LightStateOff
FlasherLight8.State=LightStateOff
FlasherLight9.State=LightStateOff
FlasherLight10.State=LightStateOff
FlasherLight11.State=LightStateOff
FlasherLight12.State=LightStateOff
FlasherLight13.State=LightStateOff
FlasherLight14.State=LightStateOff
FlasherLight15.State=LightStateOff
FlasherLight16.State=LightStateOff
  D6L.IsDropped=1
  F1A.IsDropped=1
  F2A.IsDropped=1
  F3A.IsDropped=1
  F4A.IsDropped=1
  F1B.IsDropped=1
  F2B.IsDropped=1
  F3B.IsDropped=1
  F4B.IsDropped=1
  F1C.IsDropped=1
  F2C.IsDropped=1
  F3C.IsDropped=1
  F4C.IsDropped=1
  F1D.IsDropped=1
  F2D.IsDropped=1
  F3D.IsDropped=1
  F4D.IsDropped=1
  Brain.IsDropped=1
  BrainBottom.IsDropped=1
  BugT1.IsDropped=1
  BugT2.IsDropped=1
  BugT3.IsDropped=1
  BugT4.IsDropped=1
  BugT5.IsDropped=1
  WL1.IsDropped=1
  WL2.IsDropped=1
  WL3.IsDropped=1
  WL4.IsDropped=1
  WL5.IsDropped=1
  WL6.IsDropped=1
  WL56.IsDropped=1
  WL64.IsDropped=1
  Controller.GameName=cGameName
  Controller.SplashInfoLine="Starship Troopers"
  Controller.HandleKeyboard=0
  Controller.ShowTitle=0
  Controller.ShowDMDOnly=1
  Controller.ShowFrame=0
    Controller.HandleMechanics=False
  On Error Resume Next
  'SetDisplayToElement Textbox1
  Controller.Run GetPlayerHwnd
    If Err Then MsgBox Err.Description
  On Error Goto 0

  if Controller.Version < "03030000" then MsgBox "This table requires vpinmame 3.3 rev 4895 or later. Yours reoprts " & Controller.Version & ".  Please get the latest SAMBuild from VPUniverse"

  PinMAMETimer.Interval=PinMAMEInterval:PinMAMETimer.Enabled=1
  vpmNudge.TiltSwitch=56:vpmNudge.Sensitivity=5
  vpmNudge.TiltObj=Array(LeftSlingshot,RightSlingshot,Bumper1,Bumper2,Bumper3)

  Set bsTrough=New cvpmBallStack
  bsTrough.InitSw 0,15,14,13,12,0,0,0
  bsTrough.InitKick BallRelease,90,8
  bsTrough.InitExitSnd SoundFX("fx_ballRel",DOFContactors),SoundFX("SolOn",DOFContactors)
  bsTrough.Balls=4

  Set bsBottom=New cvpmBallStack
  bsBottom.InitSaucer VUK,45,120,10
  bsBottom.InitExitSnd SoundFX("Balleject",DOFContactors), SoundFX("SolOn",DOFContactors)

  Set bsTop=New cvpmBallStack
  bsTop.InitSw 0,46,0,0,0,0,0,0
  bsTop.InitKick VUKTop,180,10
  bsTop.InitExitSnd SoundFX("fx_ScoopExit",DOFContactors), SoundFX("SolOn",DOFContactors)

  Set LMag=New cvpmMagnet
  LMag.InitMagnet LMagnet,25
  LMag.GrabCenter=True

  Set RMag=New cvpmMagnet
  RMag.InitMagnet RMagnet,25
  LMag.GrabCenter=True

' Controller.Switch(33)=1 'Set Bug Mech to Home position
  vpmMapLights AllLights

     Set mBug=New cvpmMyMech
  ' In VPX 10.7+ should use "vpmMechFourStepSol" which happens to be the same as setting both StepSol and TwoDirSol
  ' mBug.MType=vpmMechFourStepSol+vpmMechStopEnd+vpmMechLinear+vpmMechFast
  mBug.MType=vpmMechStepSol+vpmMechTwoDirSol+vpmMechStopEnd+vpmMechLinear+vpmMechFast
  mBug.Sol1=17
  mBug.Length=600'150
  mBug.Steps=160
  mBug.AddSw 34,0,1
  mBug.AddSw 33,159,160
    mBug.Callback=GetRef("UpdateBug")
  mBug.Start

'TT TESTING
'*********************************
GiOff()

' if user doesn't want bright / big flashers --> disable them
If bigflashers = 0 Then
  RF001.visible = False
  RF002.visible = False
  RF003.visible = False
  RF004.visible = False

  bf001.visible = False
  bf002.visible = False
  bf003.visible = False
  bf004.visible = False

  YF001.visible = False
  YF002.visible = False
  YF003.visible = False
  YF004.visible = False

  GF001.visible = False
  GF002.visible = False
  GF003.visible = False
  GF004.visible = False

  Warriorflash1.visible = False
  Warriorflash2.visible = False
  Warriorflash3.visible = False

End If
'**********************************

End Sub


Sub UpdateBug(NewPos, aSpeed, LastPos)
  ' Debug.Print NewPos
  dim Position, LastPosition
  Position = NewPos * 11 / 160
  LastPosition = LastPos * 11 / 160

  BugTargets(LastPosition).IsDropped=1
  BugTargets(Position).IsDropped=0

  BugWarSledPrim.TransY = NewPos

'/  Select Case Position
'   Case 0:Controller.Switch(33)=1
'   Case 1:Controller.Switch(33)=0
'   Case 4:Controller.Switch(34)=0
'   Case 5:Controller.Switch(34)=1
'   Case 6:Controller.Switch(33)=1
'   Case 7:Controller.Switch(33)=0
'   Case 10:Controller.Switch(34)=0
'   Case 11:Controller.Switch(34)=1
' End Select
End Sub



Class cvpmMyMech
  Public Sol1, Sol2, MType, Length, Steps, Acc, Ret
  Private mMechNo, mNextSw, mSw(), mLastPos, mLastSpeed, mCallback

  Private Sub Class_Initialize
    ReDim mSw(10)
    gNextMechNo = gNextMechNo + 1 : mMechNo = gNextMechNo : mNextSw = 0 : mLastPos = 0 : mLastSpeed = 0
    MType = 0 : Length = 0 : Steps = 0 : Acc = 0 : Ret = 0 : vpmTimer.addResetObj Me
  End Sub

  Public Sub AddSw(aSwNo, aStart, aEnd)
    mSw(mNextSw) = Array(aSwNo, aStart, aEnd, 0)
    mNextSw = mNextSw + 1
  End Sub

  Public Sub AddPulseSwNew(aSwNo, aInterval, aStart, aEnd)
    If Controller.Version >= "01200000" Then
      mSw(mNextSw) = Array(aSwNo, aStart, aEnd, aInterval)
    Else
      mSw(mNextSw) = Array(aSwNo, -aInterval, aEnd - aStart + 1, 0)
    End If
    mNextSw = mNextSw + 1
  End Sub

  Public Sub Start
    Dim sw, ii
    With Controller
      .Mech(1) = Sol1 : .Mech(2) = Sol2 : .Mech(3) = Length
      .Mech(4) = Steps : .Mech(5) = MType : .Mech(6) = Acc : .Mech(7) = Ret
      ii = 10
      For Each sw In mSw
        If IsArray(sw) Then
          .Mech(ii) = sw(0) : .Mech(ii+1) = sw(1)
          .Mech(ii+2) = sw(2) : .Mech(ii+3) = sw(3)
          ii = ii + 10
        End If
      Next
      .Mech(0) = mMechNo
    End With
    If IsObject(mCallback) Then mCallBack 0, 0, 0 : mLastPos = 0 : vpmTimer.EnableUpdate Me, True, True  ' <------- All for this.
  End Sub

  Public Property Get Position : Position = Controller.GetMech(mMechNo) : End Property
  Public Property Get Speed    : Speed = Controller.GetMech(-mMechNo)   : End Property
  Public Property Let Callback(aCallBack) : Set mCallback = aCallBack : End Property

  Public Sub Update
    Dim currPos, speed
    currPos = Controller.GetMech(mMechNo)
    speed = Controller.GetMech(-mMechNo)
    If currPos < 0 Or (mLastPos = currPos And mLastSpeed = speed) Then Exit Sub
    mCallBack currPos, speed, mLastPos : mLastPos = currPos : mLastSpeed = speed
  End Sub

  Public Sub Reset : Start : End Sub

End Class

Dim Old44,Old45,Old23,Old24,New44,New45,New23,New24
Old44=0:Old45=0:Old23=0:Old24=0:New44=0:New45=0:New23=0:New24=0
Dim Old1,Old2,Old3,Old4,Old5,Old6,Old34,Old56,Old64
Dim New1,New2,New3,New4,New5,New6,New34,New56,New64
Old1=0:Old2=0:Old3=0:Old4=0:Old5=0:Old6=0:Old34=0:Old56=0:Old64=0
New1=0:New2=0:New3=0:New4=0:New5=0:New6=0:New34=0:New56=0:New64=0

Set LampCallback=GetRef("UpdateMultipleLamps")
Sub UpdateMultipleLamps
  New44=Controller.Lamp(44)
  New45=Controller.Lamp(45)
  New23=Controller.Lamp(23)
  New24=Controller.Lamp(24)
  If (New44<>Old44) Or (New45<>Old45) Or (New23<>Old23) Or (New24<>Old24) Then
    EMReel1.SetValue BrainFrame
    Old44=New44
    Old45=New45
    Old23=New23
    Old24=New24
  End If
  New1=Controller.Lamp(1)
  If New1<>Old1 Then
    If New1 Then
      WL1.IsDropped=0
    Else
      WL1.IsDropped=1
    End If
  Old1=New1
  End If
  New2=Controller.Lamp(2)
  If New2<>Old2 Then
    If New2 Then
      WL2.IsDropped=0
    Else
      WL2.IsDropped=1
    End If
  Old2=New2
  End If
  New3=Controller.Lamp(3)
  If New3<>Old3 Then
    If New3 Then
      WL3.IsDropped=0
    Else
      WL3.IsDropped=1
    End If
  Old3=New3
  End If
  New4=Controller.Lamp(4)
  If New4<>Old4 Then
    If New4 Then
      WL4.IsDropped=0
    Else
      WL4.IsDropped=1
    End If
  Old4=New4
  End If
  New5=Controller.Lamp(5)
  If New5<>Old5 Then
    If New5 Then
      WL5.IsDropped=0
    Else
      WL5.IsDropped=1
    End If
  Old5=New5
  End If
  New6=Controller.Lamp(6)
  If New6<>Old6 Then
    If New6 Then
      WL6.IsDropped=0
    Else
      WL6.IsDropped=1
    End If
  Old6=New6
  End If
  New34=Controller.Lamp(34)
  If New34<>Old34 Then
    If New34 Then
      D6L.IsDropped=0
    Else
      D6L.IsDropped=1
    End If
  Old34=New34
  End If
  New56=Controller.Lamp(56)
  If New56<>Old56 Then
    If New56 Then
      WL56.IsDropped=0
    Else
      WL56.IsDropped=1
    End If
  Old56=New56
  End If
  New64=Controller.Lamp(64)
  If New64<>Old64 Then
    If New64 Then
      WL64.IsDropped=0
    Else
      WL64.IsDropped=1
    End If
  Old64=New64
  End If
End Sub

Sub LMagnet_Hit:LMag.AddBall ActiveBall:End Sub
Sub LMagnet_UnHit:LMag.RemoveBall ActiveBall:End Sub
Sub RMagnet_Hit:RMag.AddBall ActiveBall:End Sub
Sub RMagnet_UnHit:RMag.RemoveBall ActiveBall:End Sub

Sub SS_Hit:SS.DestroyBall:vpmTimer.PulseSwitch 9,300,"ATV":End Sub    '9
Sub BK_Hit:BK.DestroyBall:vpmTimer.PulseSwitch 10,240,"ATV":End Sub   '10
Sub Drain_Hit:bsTrough.AddBall Me:End Sub               '12/13/14/15
Sub S16_Hit:Controller.Switch(16)=1:End Sub               '16
Sub S16_unHit:Controller.Switch(16)=0:End Sub
Sub S25_Hit:Controller.Switch(25)=1:End Sub               '25
Sub S25_unHit:Controller.Switch(25)=0:End Sub
Sub S26_Hit:Controller.Switch(26)=1:End Sub               '26
Sub S26_unHit:Controller.Switch(26)=0:End Sub
Sub T28_Hit:vpmTimer.PulseSw 28:End Sub                 '28
Sub T29_Hit:vpmTimer.PulseSw 29:End Sub                 '29
Sub T30_Hit:vpmTimer.PulseSw 30:End Sub                 '30
Sub T31_Hit:vpmTimer.PulseSw 31:End Sub                 '31
Sub BugT1_Hit:vpmTimer.PulseSw 35:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub                  '35
Sub BugT2_Hit:vpmTimer.PulseSw 35:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub
Sub BugT3_Hit:vpmTimer.PulseSw 35:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub
Sub BugT4_Hit:vpmTimer.PulseSw 35:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub
Sub BugT5_Hit:vpmTimer.PulseSw 35:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub
Sub BugT6_Hit:vpmTimer.PulseSw 35:End Sub
Sub Brain_Hit:vpmTimer.PulseSw 38:End Sub               '38
Sub PE_Hit:PE.DestroyBall:vpmTimer.PulseSwitch 40,100,"ATV":End Sub   '40
Sub ATV(swNo):bsTop.AddBall 0:End Sub
Sub S41_Hit:Controller.Switch(41)=1:End Sub               '41
Sub S41_unHit:Controller.Switch(41)=0:End Sub
Sub S42_Hit:Controller.Switch(42)=1:End Sub               '42
Sub S42_unHit:Controller.Switch(42)=0:End Sub
Sub S43_Hit:Controller.Switch(43)=1:End Sub               '43
Sub S43_unHit:Controller.Switch(43)=0:End Sub
Sub VUK_Hit:bsBottom.AddBall 0:End Sub                  '45
Sub VUKTop_Hit:bsTop.AddBall Me:End Sub                 '46
Sub S47_Hit:Controller.Switch(47)=1:End Sub               '47
Sub S47_unHit:Controller.Switch(47)=0:End Sub
Sub S48_Hit:Controller.Switch(48)=1:End Sub               '48
Sub S48_unHit:Controller.Switch(48)=0:End Sub
Sub Bumper1_Hit:vpmTimer.PulseSw 49:PlaySoundAtVol SoundFX("Jet1",DOFContactors), ActiveBall, 1:End Sub               '49
Sub Bumper2_Hit:vpmTimer.PulseSw 50:PlaysoundAtVol SoundFX("Jet2",DOFContactors), ActiveBall, 1:End Sub               '50
Sub Bumper3_Hit:vpmTimer.PulseSw 51:PlaysoundAtVol SoundFX("Jet1",DOFContactors), ActiveBall, 1:End Sub               '51
Sub S52_Hit:Controller.Switch(52)=1:End Sub               '52
Sub S52_unHit:Controller.Switch(52)=0:End Sub
Sub S53_Hit:Controller.Switch(53)=1:End Sub               '53
Sub S53_unHit:Controller.Switch(53)=0:End Sub
Sub S57_Hit:Controller.Switch(57)=1:End Sub               '57
Sub S57_unHit:Controller.Switch(57)=0:End Sub
Sub S58_Hit:Controller.Switch(58)=1:End Sub               '58
Sub S58_unHit:Controller.Switch(58)=0:End Sub
Sub LeftSlingshot_Slingshot:vpmTimer.PulseSw 59:End Sub         '59
Sub S60_Hit:Controller.Switch(60)=1:End Sub               '60
Sub S60_unHit:Controller.Switch(60)=0:End Sub
Sub S61_Hit:Controller.Switch(61)=1:End Sub               '61
Sub S61_unHit:Controller.Switch(61)=0:End Sub
Sub RightSlingshot_Slingshot:vpmTimer.PulseSw 62:End Sub        '62

 Sub Trigger1_Hit:ActiveBall.VelZ=0:End Sub
 Sub Trigger2_Hit:ActiveBall.VelZ=0:End Sub

'Missing Motor Sound


'*****************************
'**     Switch Handling     **
'*****************************

Sub Drain_Hit:bsTrough.AddBall Me:PlaysoundAtVol "drain", drain, 1:End Sub
Sub sw16_Hit : : playsoundAtVol"rollover" , ActiveBall, 1: End Sub 'Coded to impulse plunger
'Sub sw16_unHit : Controller.Switch (16)=0:End Sub
Sub S53_Hit:Controller.switch(18) = True : playsoundAtVol"RampMidway" , ActiveBall, 1: End Sub
Sub S53_Unhit:Controller.switch(18) = False:end sub
Sub S26_Hit:Controller.switch(19) = True : playsoundAtVol"fx_metalrolling" , ActiveBall, 1: End Sub
Sub S26_Unhit:Controller.switch(19) = False:end sub
Sub S99_Hit:Controller.switch(20) = True : playsoundAtVol"RampUp" , ActiveBall, 1: End Sub
Sub S99_Unhit:Controller.switch(20) = False:end sub
Sub S98_Hit:Controller.switch(22) = True : playsoundAtVol"RampUp" , ActiveBall, 1: End Sub
Sub S98_Unhit:Controller.switch(22) = False:end sub
Sub S57_Hit:Controller.switch(57) = True : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub S57_Unhit:Controller.switch(57) = False:end sub
Sub S58_Hit:Controller.switch(58) = True : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub S58_Unhit:Controller.switch(58) = False:end sub
Sub S60_Hit:Controller.switch(60) = True : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub S60_Unhit:Controller.switch(60) = False:end sub
Sub S61_Hit:Controller.switch(61) = True : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub S60_Unhit:Controller.switch(61) = False:end sub
Sub S47_Hit:Controller.switch(37) = True : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub S47_Unhit:Controller.switch(37) = False:end sub
Sub S48_Hit:Controller.switch(39) = True : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub S48_Unhit:Controller.switch(39) = False:end sub
Sub S41_Hit:Controller.switch(30) = True : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub S41_Unhit:Controller.switch(30) = False:end sub
Sub S42_Hit:Controller.switch(31) = True : playsoundAtVol"RampMidway" , ActiveBall, 1: End Sub
Sub S42_Unhit:Controller.switch(31) = False:end sub
Sub S43_Hit:Controller.switch(32) = True : playsoundAtVol"RampEnd" , ActiveBall, 1: End Sub
Sub S43_Unhit:Controller.switch(32) = False:end sub
Sub Trigger9_Hit:Controller.switch(9) = True : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub Trigger9_Unhit:Controller.switch(9) = False:end sub


 'Stand Up Targets
 'Stand Up Targets
Sub sw17_Hit:vpmTimer.PulseSw 17:End Sub
Sub sw18_Hit:vpmTimer.PulseSw 18:End Sub
Sub sw19_Hit:vpmTimer.PulseSw 19:End Sub
Sub sw20_Hit:vpmTimer.PulseSw 20:End Sub
Sub sw21_Hit:vpmTimer.PulseSw 21:End Sub
Sub sw22_Hit:vpmTimer.PulseSw 22:End Sub
Sub sw23_Hit:vpmTimer.PulseSw 23:End Sub
Sub sw24_Hit:vpmTimer.PulseSw 24:End Sub
Sub sw27_Hit:vpmTimer.PulseSw 27:End Sub
Sub sw28_Hit:vpmTimer.PulseSw 28:End Sub


'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
  vpmTimer.PulseSw 62
    PlaySoundAtVol SoundFX("right_slingshot",DOFContactors), sling1, 1
    Rubber001.Visible = 0
    sling1.rotx = 20
    RStep = 0
    Rubber001.TimerEnabled = 1

End Sub


Sub LeftSlingShot_Slingshot
  vpmTimer.PulseSw 59
    PlaySoundAtVol SoundFX("left_slingshot",DOFContactors), sling2, 1
    Rubber002.Visible = 0
    LStep = 0
    sling2.rotx = 20
    Rubber002.TimerEnabled = 1

End Sub

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


'********************************************************************
'      JP's VP10 Rolling Sounds (+rothbauerw's Dropping Sounds)
'********************************************************************

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

    For b = 0 to UBound(BOT)
        ' play the rolling sound for each ball
        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If

        ' play ball drop sounds
        If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
            PlaySound "fx_ball_drop" & b, 0, ABS(BOT(b).velz)/17, AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
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

' Wire ramp sounds

Sub RampSound1_Hit: PlaySoundAtVol"fx_metalrolling", ActiveBall, 1: End Sub
Sub RampSound2_Hit: PlaySoundAtVol"fx_metalrolling", ActiveBall, 1: End Sub
Sub RampSound3_Hit: PlaySoundAtVol"fx_metalrolling", ActiveBall, 1: End Sub



'******************
' RealTime Updates
'******************
Set MotorCallback = GetRef("GameTimer")

Sub GameTimer
  UpdateFlipperLogo
End Sub

Sub UpdateFlipperLogo
  LogoSx.RotZ = LeftFlipper.CurrentAngle - 90
  LogoDx1.RotZ = Flipper1.CurrentAngle - 90
  LogoDx.RotZ = RightFlipper.CurrentAngle + 90
End Sub



'*******************'
' TT
'*******************


Sub GiOn()
Dim light
  For Each light in GILights
    light.State=LightStateOn
  Next

  If BrightGI = 1 Then
    GIBlue1.state = LightStateOn
    GIBlue2.state = LightStateOn
    GI038b.state = LightStateOn
    GI039b.state = LightStateOn
    GI038br.state = LightStateOn
    GI039br.state = LightStateOn
    GI040br.state = LightStateOn
    GI041br.state = LightStateOn
    GI038c.state = LightStateOn
    GI039c.state = LightStateOn
    GI038d.state = LightStateOn
    GI039d.state = LightStateOn
  End If

End Sub

Sub GiOff()
Dim light
  For Each light in GILights
    light.State=LightStateOff
  Next

  'If BrightGI = 1 Then
    GIBlue1.state = LightStateOff
    GIBlue2.state = LightStateOff
    GI038b.state = LightStateOff
    GI039b.state = LightStateOff
    GI038br.state = LightStateOff
    GI039br.state = LightStateOff
    GI040br.state = LightStateOff
    GI041br.state = LightStateOff
    GI038c.state = LightStateOff
    GI039c.state = LightStateOff
    GI038d.state = LightStateOff
    GI039d.state = LightStateOff
  'End If

End Sub


'*********************
'  Flasher fading sub
'    vpm version 1 - JP Salas
'*********************

Dim FlashState(200), FlashLevel(200)
Dim FlashSpeedUp, FlashSpeedDown

FlashInit()
FlasherTimer.Interval = 10
FlasherTimer.Enabled = 1

Sub FlashInit
    Dim i
    For i = 0 to 200
        FlashState(i) = 0
        FlashLevel(i) = 0
    Next

if Flasherspeed = 1 Then
  FlashSpeedUp = 100    ' fast speed when turning on the flasher
  FlashSpeedDown = 40  ' slow speed when turning off the flasher, gives a smooth fading
End If

if Flasherspeed = 0 Then
  FlashSpeedUp = 70    ' fast speed when turning on the flasher
  FlashSpeedDown = 10  ' slow speed when turning off the flasher, gives a smooth fading
End If

' you could also change the default images for each flasher or leave it as in the editor
' for example
' Flasher1.Image = "fr"
AllFlashOff()

FlashLight1.state = 0

End Sub

Sub AllFlashOff
    Dim i
    For i = 0 to 200
        FlashState(i) = 0
    Next
End Sub


Sub FlasherTimer_Timer()


' GI TESTING - TT
'**********************

' Added simple script to disable GI when no balls in play
' Can be deleted if real GI call from ROM is found

    Dim BOT, b
    BOT = GetBalls


    ' exit the sub if no balls on the table
    If UBound(BOT) = -1 Then
    DOF 101, DOFOff
    GiOff()
  End If

  If UBound(BOT) > -1 and GI001.state = LightStateOff Then
    DOF 101, DOFOn
    GiOn()
  End If

'**************************

    Flashlight 83, Flashlight1
    Flashlight 83, Flashlight2
    Flashlight 83, Flashlight3
    Flashlight 83, Flashlight4



        Flashm 83, bf001
    Flashm 83, bf002
        Flashm 83, bf004
    Flash 83, bf003


      Flashlight 82, Flashlight5
    Flashlight 82, Flashlight8
    Flashlight 82, Flashlight9
    Flashlight 82, Flashlight10



        Flashm 82, Yf001
    Flashm 82, Yf002
        Flashm 82, Yf004
    Flash 82, Yf003



      Flashlight 81, Flashlight7
    Flashlight 81, Flashlight14
    Flashlight 81, Flashlight15
    Flashlight 81, Flashlight16


        Flashm 81, Rf001
    Flashm 81, Rf002
        Flashm 81, Rf004
    Flash 81, Rf003


    Flashlight 84, Flashlight6
    Flashlight 84, Flashlight11
    Flashlight 84, Flashlight12
    Flashlight 84, Flashlight13



        Flashm 84, Gf001
    Flashm 84, Gf002
        Flashm 84, Gf004
    Flash 84, Gf003



'   flashm 99, flasher1
'   flashm 99, flasher3
'   flashm 99, flasher4
'   flashm 99, flasher5
'   flashm 99, flasher6
'   flashm 99, flasher7
'   flashm 99, flasher8
'   flashm 99, flasher9
'   flashm 99, flasher10
'   flashm 99, flasher11
'   flashm 99, flasher12
'   flashm 99, flasher13
'   flashm 99, flasher14
'   flashm 99, flasher15
'   flashm 99, flasher16
'   flashm 99, flasher17
'   flash 99, flasher2

'ramp flashers
   flashm 86, LRFnew1
   flashm 86, LRFnew2
   flashm 86, LRFnew3
   flashm 86, LRFnew4
   flash 86, LRFnew5

   flashm 87, LRFnew6
   flashm 87, LRFnew7
   flashm 87, LRFnew8
   flashm 87, LRFnew9
   flash 87, LRFnew10


' flash 114, RLSHalo
' flash 113, LLSHalo
' flash 34, yellowhalo5
'

' flash 112, f112
' flash 110, f110
' flashm 139, f139b
' flashm 96, f96b
' flash 139, f139
' flash 96, f96
' flash 109, f109
' flash 111, f111

'Warriorflash
    Flashlight 199, WarriorFlashlight1
    Flash 199, Warriorflash1




End Sub

Sub Flashlight(nr, Object)

  dim lstate

  if FlashState(nr) = -1 then lstate = 0 end If
  if FlashState(nr) = -2 then lstate = 1 end If

  object.state = lstate

End Sub

Sub SetFlash(nr, stat)
    FlashState(nr) = ABS(stat)
End Sub

Sub Flash(nr, object)
    Select Case FlashState(nr)
        Case 0 'off
            FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown
            If FlashLevel(nr) < 0 Then
                FlashLevel(nr) = 0
                FlashState(nr) = -1 'completely off
            End if
            Object.opacity = FlashLevel(nr)
        Case 1 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp
            If FlashLevel(nr) > 255 Then
                FlashLevel(nr) = 255
                FlashState(nr) = -2 'completely on
            End if
            Object.opacity = FlashLevel(nr)
    End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it doesn't change the flashstate
    Select Case FlashState(nr)
        Case 0         'off
            Object.opacity = FlashLevel(nr)
        Case 1         ' on
            Object.opacity = FlashLevel(nr)
    End Select
End Sub


'********************
' END TT modifications
'********************

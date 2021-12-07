'************************************************************
'************************************************************
'
'  Pinball Magic (CAPCOM 1995) - IPD No. 3596
'
'   Credits:
'
'     VPX by unclewilly, rothbauerw
'     Contributions from:
'        DJRobX (switch debugging and stage functionality, this never worked before, IT DOES NOW!!!)
'        randr (hand and misc primitivies)
'
'
'************************************************************
'************************************************************

' Thalamus 2018-07-24
' Tables has its own "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Table supports useSolenoid=2 but script updates is needed - what are they ?
' No special SSF tweaks yet.

Option Explicit
Randomize

'******************************************************
'             OPTIONS
'******************************************************

Const EnableBallControl = False   'set to false to disable the C key from taking manual control
Const FlippersOnOff = 1       '0 for always on flippers, 1 to simulate "GameOn" solenoid (DOF 110 for left flipper, DOF 111 for right flipper)
Const BypassMagnet = 1        '0 for default behavior, 1 to bypass magnet activation for trunk to prevent rom from enabling coil protection (DOF 112 for shaker)
Const PlungerOption = 1       '0 for fullscreen plunger, 1 for desktop plunger

Const VolumeDial = 5        'Change bolume of hit events
Const RollingSoundFactor = 1    'set sound level factor here for Ball Rolling Sound, 1=default level


Dim ContrastSetting, GlowAmountDay, InsertBrightnessDay, GlowAmountNight, InsertBrightnessNight
' *** Contrast level, possible values are 0 - 7, can be done in game with magnasave keys **
' *** 0: bright, good for desktop view, daytime settings in insert lighting below *********
' *** 1: same as 0, but with nighttime settings in insert lighting below ******************
' *** 2: darker, better contrast, daytime settings in insert lighting below ***************
' *** 3: same as 2, but with nighttime settings in insert lighting below ******************
' *** etc for 4-7; default is 3 ***********************************************************
ContrastSetting = 3

' *** Insert Lighting settings ************************************************************
' *** The settings below together with ContrastSetting determines how the lighting looks **
' *** for all values: 1.0 = default, useful range 0.1 - 5 *********************************
GlowAmountDay = 0.05
InsertBrightnessDay = 0.8
GlowAmountNight = 0.5
InsertBrightnessNight = 0.6

'******************************************************
'           STANDARD DEFINITIONS
'******************************************************


Const cGameName = "pmv112"
Const BallSize = 50
Const BallMass = 1.7

Const UseSolenoids = 1
Const UseLamps = 0
Const UseSync = 1
Const HandleMech = 0

'Standard Sounds
Const SSolenoidOn = "Solenoid"
Const SSolenoidOff = ""
Const SCoin = "coin"

'******************************************************
'           TABLE INIT
'******************************************************

if Version < 10400 then msgbox "This table requires Visual Pinball 10.4 beta or newer!" & vbnewline & "Your version: " & Version/1000

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01560000", "capcom.VBS", 3.26

Dim DesktopMode, NightDay:DesktopMode = Table1.ShowDT:NightDay = Table1.NightDay
Dim VarHidden, UseVPMDMD, xx,  ClearBall, eyeBall1, eyeBall2, eyeBall3

Sub Table1_Init
  vpmInit Me
  With Controller
    .GameName =  cGameName
    If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
    .SplashInfoLine = "Pinball Magic" & vbNewLine & "by unclewilly/rothbauerw VPX"
    .HandleKeyboard = 0
    .ShowTitle = 0
    .ShowDMDOnly = 1
    .ShowFrame = 0
    .HandleMechanics = 0
    .Hidden = VarHidden
    On Error Resume Next
    .Run GetPlayerHWnd
    If Err Then MsgBox Err.Description
    On Error Goto 0
  End With

  'Nudging
    vpmNudge.TiltSwitch=10
    vpmNudge.Sensitivity=1
    vpmNudge.TiltObj=Array(Bumper1b,LeftSlingshot,RightSlingshot)

  WandDiv.IsDropped=1:WandDiv1.IsDropped=0
  StageDiverter3.isDropped=1
  Kickback.PullBack

    '************  Trough **************************
  Set eyeBall1 = Slot1.CreateSizedballWithMass (Ballsize/2,Ballmass)
  Set eyeBall2 = Slot2.CreateSizedballWithMass (Ballsize/2,Ballmass)
  Set eyeBall3 = Slot3.CreateSizedballWithMass (Ballsize/2,Ballmass)
  'Slot4.CreateSizedballWithMass Ballsize/2,Ballmass
  Controller.Switch(74) = 1
  Controller.Switch(75) = 1
  Controller.Switch(76) = 1

    '************  Captive Ball **************************
  Kicker3.CreateSizedballWithMass Ballsize/2,Ballmass
  kicker3.kick 0, 0
  set ClearBall = Kicker4.CreateSizedballWithMass(30,Ballmass*3/4)
  ClearBall.visible = False
  kicker4.kick 180, 10
  kicker4.enabled = false


  '**Main Timer init
  PinMAMETimer.Interval = PinMAMEInterval
  PinMAMETimer.Enabled = 1

  '** Stage
  Controller.Switch(59)= 1
  Controller.Switch(58)= 0
  CreateLBall

  '** Wand
  Controller.Switch(67)=1
  Controller.Switch(68)=0
  CriticDiv.IsDropped = 1
  TrunkDiv.IsDropped = 0
  ramp19.collidable = false

  If DesktopMode = False Then
    Ramp15.visible = False
    Ramp16.visible = False
    Ramp17.visible = False
  End If

  If PlungerOption = 1 Then
    Plunger.visible = False
    Plunger1.visible = True
    Plunger.MechPlunger = False
    Plunger1.MechPlunger = True
    Plunger.Pullback
  Else
    Plunger.visible = True
    Plunger1.visible = False
    Plunger.MechPlunger = True
    Plunger1.MechPlunger = False
    Plunger1.Pullback
  End If

  ColorGrade

End Sub

Sub Table1_Paused:Controller.Pause = 1:End Sub
Sub Table1_unPaused:Controller.Pause = 0:End Sub
Sub Table1_Exit
  DOF 101, DOFOff
  DOF 102, DOFOff
  DOF 103, DOFOff
  DOF 104, DOFOff
  DOF 105, DOFOff
  DOF 110, DOFOff
  DOF 111, DOFOff
  DOF 112, DOFOff
  Controller.Stop
End Sub

'******************************************************
'             KEYS
'******************************************************

Sub Table1_KeyDown(ByVal keycode)
  If keycode = LeftFlipperKey Then SolLFlipper(True)
  If keycode = RightFlipperKey Then SolRFlipper(True)
  If keycode = plungerkey then
    If PlungerOption = 1 then
      plunger1.PullBack
    Else
      plunger.PullBack
    End If
    PlaySound "plungerpull",0,1,AudioPan(Plunger),0.25,0,0,1,AudioFade(Plunger)
  End If
  If keycode = LeftTiltKey Then Nudge 90, 4
  If keycode = RightTiltKey Then Nudge 270, 4
  If keycode = CenterTiltKey Then Nudge 0, 5
  If KeyCode = KeyFront then Controller.Switch(11)=1

  If keycode = RightMagnaSave Then
    ContrastSetting = ContrastSetting + 1
    If ContrastSetting > 7 Then ContrastSetting = 7 End If
    ColorGrade
    Primitive6.collidable=0
  End If
  If keycode = LeftMagnaSave Then
    ContrastSetting = ContrastSetting - 1
    If ContrastSetting < 0 Then ContrastSetting = 0 End If
    ColorGrade
  End If

    ' Manual Ball Control
  If keycode = 46 Then          ' C Key
    If contball = 1 Then
      contball = 0
    Else
      contball = 1
    End If
  End If
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

  If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal keycode)
  If keycode = LeftFlipperKey Then SolLFlipper(False)
  If keycode = RightFlipperKey Then SolRFlipper(False)
  If keycode = plungerkey then
    If PlungerOption = 1 then
      plunger1.fire
    Else
      plunger.fire
    End If
    PlaySound "plunger",0,1,AudioPan(Plunger),0.25,0,0,1,AudioFade(Plunger)
  End If
  If KeyCode = KeyFront then Controller.Switch(11)=0

    'Manual Ball Control
  If keycode = 203 Then BCleft = 0  ' Left Arrow
  If keycode = 200 Then BCup = 0    ' Up Arrow
  If keycode = 208 Then BCdown = 0  ' Down Arrow
  If keycode = 205 Then BCright = 0 ' Right Arrow

' Thalamus - sometime ball gets stuck so lets let free by using the right magna save.

  If keycode = RightMagnaSave Then
    Primitive6.collidable=1
  End If

  If vpmKeyUp(keycode) Then Exit Sub
End Sub

'******************************************************
'             SOLENOIDS
'******************************************************

SolCallback(1)="SolTrunk"
SolCallback(2)="SolGenie"
SolCallback(5)="SolKB"
SolCallback(6)="SolTroughOut"
SolCallback(7)="SolTroughIn"
SolCallBack(8)="vpmSolSound SoundFX(""Knocker"",DOFKnocker),"

SolCallback(11)="SolStageDoors"
SolCallback(12)="SolStageKicker"
SolCallback(13)="SolStageDiverter"
SolCallback(14)="SolWandDiverter"
SolCallback(17)="SolWand"
SolCallback(18)="SolElevator"
SolCallback(19)="DropReset"
SolCallback(20)="SolMag"

SolCallback(26)="SetLamp 176,"
SolCallback(27)="SetLamp 177,"
SolCallback(28)="SetLamp 178,"
SolCallback(29)="SetLamp 179,"
SolCallback(30)="SetLamp 180,"
SolCallback(31)="SetLamp 181,"
SolCallback(32)="SetLamp 182,"

'**************
' Solenoid Subs
'**************

'***** Flippers *************************

'SolCallback(sLRFlipper) = "SolRFlipper"
'SolCallback(sLLFlipper) = "SolLFlipper"

Dim LeftKeyDown, RightKeyDown, FlippersEnabled

Function AllInTrough(drainhit)
  dim intrough
  intrough = 0
  If FlippersOnOff Then
    If FlippersEnabled Then
      If InRect(eyeball1.x, eyeball1.y, 378, 2096, 823, 1800, 856, 1868, 536, 2096) then intrough = intrough + 1
      If InRect(eyeball2.x, eyeball2.y, 378, 2096, 823, 1800, 856, 1868, 536, 2096) then intrough = intrough + 1
      If InRect(eyeball3.x, eyeball3.y, 378, 2096, 823, 1800, 856, 1868, 536, 2096) then intrough = intrough + 1
      If intrough = 3 then
        AllInTrough = True
      Elseif Intrough = 2 and controller.Switch(27) and drainhit Then
        AllInTrough = True
      Elseif InTrough = 1 and controller.Switch(27) and controller.Switch(26) and drainhit Then
        AllInTrough = True
      Else
        AllInTrough = False
      End If
    Else
      AllinTrough = True
    End If
  Else
    AllinTrough = False
  End If
End Function

Sub SolLFlipper(Enabled)
    If Enabled Then
    If NOT AllInTrough(0) Then
      PlaySound SoundFX("fx_flipperup",DOFFlippers), 0, .67, AudioPan(LeftFlipper), 0.05,0,0,1,AudioFade(LeftFlipper)
      LeftFlipper.RotateToEnd
      DOF 110, DOFOn
    Else
      DisableFlippers
    End If
    LeftKeyDown = True
    Else
    If NOT AllInTrough(0) Then
      PlaySound SoundFX("fx_flipperdown",DOFFlippers), 0, 1, AudioPan(LeftFlipper), 0.05,0,0,1,AudioFade(LeftFlipper)
      LeftFlipper.RotateToStart
      DOF 110, DOFOff
    End if
    LeftKeyDown = False
    End If
End Sub
'
Sub SolRFlipper(Enabled)
    If Enabled Then
    If NOT AllInTrough(0) Then
      PlaySound SoundFX("fx_flipperup",DOFFlippers), 0, .67, AudioPan(RightFlipper), 0.05,0,0,1,AudioFade(RightFlipper)
      RightFlipper.RotateToEnd
      DOF 111, DOFOn
    Else
      DisableFlippers
    End If
    RightKeyDown = True
    Else
    If NOT AllInTrough(0) Then
      PlaySound SoundFX("fx_flipperdown",DOFFlippers), 0, 1, AudioPan(RightFlipper), 0.05,0,0,1,AudioFade(RightFlipper)
      RightFlipper.RotateToStart
      DOF 111, DOFOff
    End If
    RightKeyDown = False
    End If
End Sub

Sub DisableFlippers()
  FlippersEnabled = False

  Bumper1b.threshold = 100
  LeftSlingShot.slingshotthreshold = 100
  RightSlingShot.slingshotthreshold = 100
  RightSlingShot1.slingshotthreshold = 100

  If LeftFlipper.currentAngle < LeftFlipper.StartAngle Then
    PlaySound SoundFX("fx_flipperdown",DOFFlippers), 0, 1, AudioPan(LeftFlipper), 0.05,0,0,1,AudioFade(LeftFlipper)
    LeftFlipper.RotateToStart
    DOF 110, DOFOff
  End If
  If RightFlipper.currentAngle > RightFlipper.StartAngle Then
    PlaySound SoundFX("fx_flipperdown",DOFFlippers), 0, 1, AudioPan(RightFlipper), 0.05,0,0,1,AudioFade(RightFlipper)
    RightFlipper.RotateToStart
    DOF 111, DOFOff
  End If
End Sub

Sub ActivateFlippers()
  FlippersEnabled = True

  Bumper1b.threshold = 3
  LeftSlingShot.slingshotthreshold = 3
  RightSlingShot.slingshotthreshold = 3
  RightSlingShot1.slingshotthreshold = 3

  If RightKeyDown and RightFlipper.currentAngle < RightFlipper.StartAngle + 1 Then
    SolRFlipper(True)
  End If
  If LeftKeyDown and LeftFlipper.currentAngle > LeftFlipper.StartAngle - 1 Then
    SolLFlipper(True)
  End If
End Sub


'***** Kick Back *************************
Sub SolKB(enabled)
  If enabled Then
    Kickback.Fire
    PlaySoundAt SoundFX("AutoPlunger",DOFContactors), kickback
  Else
    Kickback.Pullback
  end If
End Sub

'***** Trunk Lock ************************
TrunkWall1.IsDropped = 1:TrunkWall2.IsDropped = 1
Sub SolTrunk(enabled)
  If enabled Then
    TrunkDoorP.TransY = -100
    TrunkDoor.IsDropped = 1
    TrunkDoor.TimerEnabled = 1
    PlaySoundAt SoundFX("solenoid",DOFContactors), TrunkDoorP
  Else

  End If
End Sub

Sub TrunkDoor_Timer
  TrunkDoorP.TransY = 0
  TrunkDoor.IsDropped = 0
  TrunkDoor.TimerEnabled = 0
End Sub

Sub Sw25_hit()
  DOF 112, DOFOff
  Controller.Switch(25) = 1
End Sub

Sub Sw25_Unhit()
  Controller.Switch(25) = 0
End Sub

Sub Sw26_hit()
  TrunkWall2.IsDropped = 0
  Controller.Switch(26) = 1
End Sub

Sub Sw26_Unhit()
  TrunkWall2.IsDropped = 1
  Controller.Switch(26) = 0
End Sub

Sub Sw27_hit()
  TrunkWall1.IsDropped = 0
  Controller.Switch(27) = 1
End Sub

Sub Sw27_Unhit()
  TrunkWall1.IsDropped = 1
  Controller.Switch(27) = 0
  If controller.Switch(28) = true then
    ActivateFlippers
  End If
End Sub

'*********** Wand ************************
Dim WandPos,WandDir, Wandcount
WandPos = 42:WandDir = -1
HandP.ObjRotZ = WandPos
WandP.ObjRotZ = WandPos
WandR.ObjRotZ = WandPos

Sub SolWand(enabled)
If Enabled Then
  DOF 101, DOFOn
  WandT.Enabled = 1
Else
  WandT.enabled = 0
end if
End Sub

Sub WandT_Timer
  if wandcount = 0 then
    PlaysoundAt SoundFX("Motor",DOFGear), HandP
  end If
  wandcount = wandcount + 1
  If wandcount = 10 then wandcount = 0
  WandPos = WandPos + (WandDir * 0.1)
  If WandPos > 42 then WandPos = 42:WandDir = -1
  If WandPos < -22 then WandPos = -22:WandDir = 1
  If WandPos >38 then
    Controller.Switch(67)=1
    CriticDiv.IsDropped = 1
    TrunkDiv.IsDropped = 0
  Else
    DOF 101, DOFOff
    Controller.Switch(67)=0
  end If
  If WandPos < -18 then
    Controller.Switch(68)=1
    CriticDiv.IsDropped = 0
    TrunkDiv.IsDropped = 1
  Else
    DOF 101, DOFOff
    Controller.Switch(68)=0
  end If
  If WandPos > -18 and wandPos < 38 Then
    CriticDiv.IsDropped = 0
    TrunkDiv.IsDropped = 0
  end If
  HandP.ObjRotZ = WandPos
  WandP.ObjRotZ = WandPos
  WandR.ObjRotZ = WandPos
End Sub

Sub SolMag(enabled)
  If enabled Then
    Ramp19.Collidable = True
  Else
    Ramp19.Collidable = False
  End If
End Sub

'*******End Wand

Sub SolStageDiverter(enabled)
  If enabled then
    PlaySoundAt SoundFX("solenoid",DOFContactors), Stagein
    StageDiverter.IsDropped = 1
    StageDiverter3.IsDropped = 0
  Else
    PlaySoundAt SoundFX("solenoid",DOFContactors), Stagein
    StageDiverter.IsDropped = 0
    StageDiverter3.IsDropped = 1
  end if
End Sub

Sub SolWandDiverter(enabled)
  If enabled then
    PlaySoundAt SoundFX("solenoid",DOFContactors), sw51a
    WandDiv.IsDropped=0:WandDiv1.IsDropped=1
  Else
    PlaySoundAt SoundFX("solenoid",DOFContactors), sw51a
    WandDiv.IsDropped=1:WandDiv1.IsDropped=0
  End If
End Sub

Sub WandDiv_Timer()
  WandDiv.IsDropped=1:WandDiv1.IsDropped=0
  WandDiv.TimerEnabled = 0
End Sub

Sub Stagein_Hit()
  PlaySoundAt "kicker_enter_center", stagein
  controller.Switch(60) = 1
End Sub

Sub Stagein_UnHit()
  controller.Switch(60) = 0
End Sub

Dim MyBall, ELActive
ElActive = 0

Sub Elevator_Hit()
  PlaySoundAt "kicker_enter_center", stagein
  Set MyBall=ActiveBall
  Controller.Switch(57)=1
  ElActive = 1
End Sub

Sub Elevator_unHit()
  Controller.Switch(57)=false
End Sub

Dim LBall, EDown, EPos
EPos = 300
Edown = 1

Sub CreateLBall()
  Set LBall=LevBall.CreateBall
  LBall.Z = EPos
End Sub

Sub solElevator(enabled)
  If Enabled Then
    DOF 102, DOFOn
    EMotor.Enabled = 1
  Else
    DOF 102, DOFOff
    EMotor.Enabled = 0
    If Epos > 250 then
      If ElActive = 1 then
        MyBall.Z=258
        MyBall.X=785
        Elactive = 0
        elevator.kick 90, 5
        controller.Switch(57)=false
        ActivateFlippers
      End If
    End If
  end If
end Sub

Sub Emotor_Timer()
  PlaysoundAt SoundFX("Motor",DOFGear), Stagein
  Dim ElDelt
  If Epos > 289.9 then
    ElDelt = 0.5
  Else
    ElDelt = 2.5
  End If

  If EDown = 0 Then
    Epos = Epos + ElDelt
    LBall.Z = EPos
    'LBall.visible = true
    el.z = EPos/2
    If EPos > 299 then
      EDown = 1
      'emotor.enabled = false
    End If
  else
    Epos = Epos - ElDelt
    LBall.Z = EPos
    'LBall.visible = False
    el.z = EPos/2
    If Epos < 116 Then
      EDown = 0
      'emotor.enabled = false
    End If
  end if
  if Epos < 125 then
    Controller.Switch(59) = 0
    Controller.Switch(58) = 1
  elseif EPos > 298 then
    Controller.Switch(59) = 1
    Controller.Switch(58) = 0
  else
    Controller.Switch(59) = 1
    Controller.Switch(58) = 1
  end if
End Sub

Sub SolStageKicker(enabled)
  If Enabled then
    SKick
  End If
End Sub

Sub SKick()
  stagein.Kickz 160,25,80,0
  PlaySoundAt "popper_ball", stagein
  DOF 105, DOFPulse
End Sub

Dim DoorDir, CloseTime
DoorDir = 2
CloseTime = Now

Sub SolStageDoors(enabled)
  If enabled then
    DoorDir = 2
    DoorT.timerenabled = 1
    ' Power is on keep it open.
    CloseTime = DateAdd("s", 1000, Now)
    DOF 103,DOFOn
  Else
    ' Power is off, start to close if we remain in this state for more than 1s.
    CloseTime = DateAdd("s",1, Now)
  end if
End Sub

Dim DoorPos, DrOpn
DoorPos = 0:DrOpn = 0

Sub DoorT_Timer()
  ' Power has dissipated, close.
  if  Now > CloseTime then
    DoorDir = -2
    DOF 103,DOFOn
  end if

  DoorPos = DoorPos + DoorDir

  If DoorPos >= 40 then
    DoorPos = 40
    DOF 103,DOFOff
  ElseIf DoorPos < 0 then
    DoorPos = 0
    DoorT.timerenabled = 0
    DOF 103,DOFOff
  Else
    PlaysoundAt SoundFX("Motor",DOFGear), Stagein
  End If

  DoorL.TransX = -DoorPos
  DoorR.TransX = DoorPos
End Sub

'***********************************************
 'Kickers, drains, poppers

'******************************************************
'     TROUGH BASED ON NFOZZY'S
'******************************************************

Sub Slot3_Hit():Controller.Switch(74) = 1:UpdateTrough:End Sub
Sub Slot3_UnHit():Controller.Switch(74) = 0:UpdateTrough:End Sub
Sub Slot2_Hit():Controller.Switch(75) = 1:End Sub
Sub Slot2_UnHit():Controller.Switch(75) = 0:UpdateTrough:End Sub
Sub Slot1_Hit():Controller.Switch(76) = 1:End Sub
Sub Slot1_UnHit():Controller.Switch(76) = 0:UpdateTrough:End Sub

Sub UpdateTrough()
  UpdateTroughTimer.Interval = 300
  UpdateTroughTimer.Enabled = 1
End Sub

Sub UpdateTroughTimer_Timer()
  If Slot1.BallCntOver = 0 Then Slot2.kick 60, 9
  If Slot2.BallCntOver = 0 Then Slot3.kick 60, 9
  Me.Enabled = 0
End Sub

'******************************************************
'       DRAIN & RELEASE
'******************************************************

Sub Drain_Hit()
  PlaySoundAt "drain", drain
  UpdateTrough
  Controller.Switch(73) = 1
  If AllInTrough(True) Then DisableFlippers
End Sub

Sub Drain_UnHit()
  Controller.Switch(73) = 0
End Sub

Sub SolTroughIn(enabled)
  If enabled Then
    Drain.kick 90,20
    PlaySoundAt SoundFX(SSolenoidOn,DOFContactors), drain
  End If
End Sub

Sub SolTroughOut(enabled)
  If enabled Then
    PlaySoundAt SoundFX("ballrelease",DOFContactors), Slot1
    Slot1.kick 90, 9
    UpdateTrough
  End If
End Sub


Sub SolGenie(enabled)   'GenieTrough
  sw28.kick 321, 46
  PlaySoundAt SoundFX("AutoPlunger",DOFContactors), sw28
End Sub


'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep, RStep1

Sub RightSlingShot_Slingshot
  vpmTimer.PulseSw 22
    PlaySound SoundFX("right_slingshot",DOFContactors),0,1, AudioPan(SLING1), 0.05,0,0,1,AudioFade(SLING1)
    RS.Visible = 0
    RS1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1

End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RS1.Visible = 0:RS2.Visible = 1:sling1.TransZ = -10
        Case 4:RS2.Visible = 0:RS.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub RightSlingShot1_Slingshot
  vpmTimer.PulseSw 46
    PlaySound SoundFX("right_slingshot",DOFContactors),0,1, AudioPan(SLING3), 0.05,0,0,1,AudioFade(SLING3)
    RtS.Visible = 0
    RtS1.Visible = 1
    sling3.TransZ = -20
    RStep1 = 0
    RightSlingShot1.TimerEnabled = 1

End Sub

Sub RightSlingShot1_Timer
    Select Case RStep1
        Case 3:RtS1.Visible = 0:RtS2.Visible = 1:sling3.TransZ = -10
        Case 4:RtS2.Visible = 0:RtS.Visible = 1:sling3.TransZ = 0:RightSlingShot1.TimerEnabled = 0
    End Select
    RStep1 = RStep1 + 1
End Sub

Sub LeftSlingShot_Slingshot
  vpmTimer.PulseSw 19
    PlaySound SoundFX("left_slingshot",DOFContactors),0,1, AudioPan(SLING2), 0.05,0,0,1,AudioFade(SLING2)
    LS.Visible = 0
    LS1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1

End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LS1.Visible = 0:LS2.Visible = 1:sling2.TransZ = -10
        Case 4:LS2.Visible = 0:LS.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

'AutomaticUpdates
Sub UpdatesTimer_Timer()
  RollingUpdate
  PrimBall1.x = ClearBall.x
  PrimBall1.y = ClearBall.y
  PrimBall1.z = ClearBall.z
  FlipperLSh.RotZ = LeftFlipper.currentangle
  FlipperRSh.RotZ = RightFlipper.currentangle
  BallShadowUpdate
  BallControlTimer
End Sub

'************Bumper1

Sub Bumper1b_Hit()
  vpmTimer.PulseSw 69
    PlaySoundAt SoundFX("fx_bumper1",DOFContactors), Bumper1b
End Sub

'**********Drop Targets
Sub sw41_Hit()
  Controller.Switch(41) = 1
  PlaySoundAt SoundFX("droptarget",DOFDropTargets), sw41
End Sub

Sub sw42_Hit()
  Controller.Switch(42) = 1
  PlaySoundAt SoundFX("droptarget",DOFDropTargets), sw42
End Sub

Sub sw43_Hit()
  Controller.Switch(43) = 1
  PlaySoundAt SoundFX("droptarget",DOFDropTargets), sw43
End Sub

Sub sw44_Hit()
  Controller.Switch(44) = 1
  PlaySoundAt SoundFX("droptarget",DOFDropTargets), sw44
End Sub

Sub sw45_Hit()
  Controller.Switch(45) = 1
  PlaySoundAt SoundFX("droptarget",DOFDropTargets), sw45
End Sub

Sub DropReset(enabled)
  If enabled then
    Controller.Switch(41) = 0
    Controller.Switch(42) = 0
    Controller.Switch(43) = 0
    Controller.Switch(44) = 0
    Controller.Switch(45) = 0
    sw41.IsDropped = 0:sw42.IsDropped = 0:sw43.IsDropped = 0
    sw44.IsDropped = 0:sw45.IsDropped = 0
    PlaySoundAt SoundFx("drop_reset",DOFContactors), sw43
  end if
End Sub
'****Gates
Sub sw49_Hit()
  PlaySoundAt "gate", sw49
  vpmTimer.PulseSw 49
End Sub
'********Standup Targets
Sub sw34_Hit()
  vpmTimer.PulseSw 34
    PlaySoundAt SoundFX("target",DOFTargets),sw34
End Sub

Sub sw31_Hit()
  vpmTimer.PulseSw 31
    PlaySoundAt SoundFX("target",DOFTargets),sw31
End Sub

Sub sw65_Hit()
  vpmTimer.PulseSw 65
    PlaySoundAt SoundFX("target",DOFTargets),sw65
End Sub

Sub sw70_Hit()
  vpmTimer.PulseSw 70
    PlaySoundAt SoundFX("target",DOFTargets),sw70
End Sub

Sub sw71_Hit()
  vpmTimer.PulseSw 71
    PlaySoundAt SoundFX("target",DOFTargets),sw71
End Sub

Sub sw72_Hit()
  vpmTimer.PulseSw 72
    PlaySoundAt SoundFX("target",DOFTargets),sw72
End Sub

'********Rollovers
Sub Sw17_hit()
  PlaySoundAt "sensor",sw17
  sensorStall Activeball
  Controller.Switch(17) = 1
End Sub

Sub Sw17_unhit()
  Controller.Switch(17) = 0
End Sub


Sub Sw18_hit()
  PlaySoundAt "sensor",sw18
  sensorStall Activeball
  vpmTimer.PulseSw 18
End Sub

Sub Sw23_hit()
  PlaySoundAt "sensor",sw23
  sensorStall Activeball
  vpmTimer.PulseSw 23
End Sub

Sub Sw24_hit()
  PlaySoundAt "sensor",sw24
  sensorStall Activeball
  vpmTimer.PulseSw 24
End Sub

Sub Sw33_hit()
  PlaySoundAt "sensor",sw33
  sensorStall Activeball
  controller.switch(33)=1
  ActivateFlippers
End Sub

Sub Sw33_unhit()
  controller.switch(33)=0
End Sub

Sub Sw35_hit()
  vpmTimer.PulseSw 35
End Sub


Sub Sw36_hit()
  PlaySoundAt "sensor",sw36
  sensorStall Activeball
  vpmTimer.PulseSw 36
End Sub

Sub Sw37_hit()
  PlaySoundAt "sensor",sw37
  sensorStall Activeball
  vpmTimer.PulseSw 37
End Sub

Sub Sw38_hit()
  PlaySoundAt "sensor",sw38
  sensorStall Activeball
  vpmTimer.PulseSw 38
End Sub

Sub Sw40_hit()
  PlaySoundAt "sensor",sw40
  sensorStall Activeball
  vpmTimer.PulseSw 40
End Sub

Sub sw51_Hit()
    vpmTimer.PulseSw 51
End Sub

Sub SpinnerTrig_Unhit()
  sensorStall Activeball
End Sub

Sub SensorStall(ball)
  Dim speedFactor
  speedFactor = 0.875
  If ballvel(ball) >= 33 Then speedfactor = 0.825
  ball.velx = ball.velx * speedFactor
  ball.vely = ball.vely * speedFactor
End Sub

Dim CriticsBall

Sub sw51a_Hit()
  Set CriticsBall = ActiveBall
  if ActiveBall.vely > 15 Then ActiveBall.vely = 15
  sw51a.timerenabled = true
  If trunkDiv.isdropped = 0 and BypassMagnet = 1 Then
    DOF 112, DOFOn
  Else
    vpmTimer.PulseSw 51
  End If
End Sub

Sub sw51a_Timer()
  If CriticsBall.z < 170 Then
    sw51a.timerenabled = false
    playsoundat "balldrop", CriticsBall
  Elseif CriticsBall.x < 360 Then
    sw51a.timerenabled = false
  End If
End Sub

Dim WireRampBall

Sub sw52_Hit()
  Set WireRampBall = ActiveBall
  sw52.timerenabled = true
  Controller.Switch(52) = 1
End Sub

Sub sw52_UnHit()
  Controller.Switch(52) = 0
End Sub

Sub sw52_Timer()
  PlaySound("WireRamp"), -1, Vol(WireRampBall)*VolumeDial*10, AudioPan(WireRampBall), 0, Pitch(WireRampBall), 1, 0, AudioFade(WireRampBall)
  If WireRampBall.z < 170 Then
    me.timerenabled = False
    StopSound "WireRamp"
  end if
End Sub


Sub Hathole_hit: PlaySoundat "fx_subway3", Hathole: Hathole.Kick -90, 10 : End Sub


'*******Genie Trough


Sub Sw30_hit()
  Controller.Switch(30) = 1
End Sub

Sub Sw30_Unhit()
  Controller.Switch(30) = 0
End Sub

Sub Sw29_hit()
  Controller.Switch(29) = 1
End Sub

Sub Sw29_Unhit()
  Controller.Switch(29) = 0
End Sub

Sub Sw28_hit()
  Controller.Switch(28) = 1
  PlaySoundat "kicker_enter_center", sw28
End Sub

Sub Sw28_Unhit()
  Controller.Switch(28) = 0
End Sub
'*****Spinner
Sub Spinner_Spin()
  vpmTimer.PulseSw 66
    PlaySoundAt "fx_spinner", spinner
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

Dim LampCount

Sub LampTimer_Timer()
    Dim chgLamp, num, chg, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp) Then
        For ii = 0 To UBound(chgLamp)
      If ChgLamp(ii,1) = 1 Then
        LampCount = LampCount + 1
      Else
        LampCount = LampCount - 1
      End If
      LampState(chgLamp(ii, 0) ) = chgLamp(ii, 1)       'keep the real state in an array
            FadingLevel(chgLamp(ii, 0) ) = chgLamp(ii, 1) + 4 'actual fading step
        Next
    End If
    UpdateLamps
End Sub

Dim TiltCount

Sub UpdateLamps

  If FlippersEnabled and LampCount = 5 Then
    TiltCount = TiltCount + 1
    If TiltCount > 9 Then
      DisableFlippers
      TiltCount = 0
    End If
  Else
    TiltCount = 0
  End If

    NFadeL 5, l5
    NFadeL 6, l6
    NFadeL 7, l7
    NFadeL 8, l8
    NFadeL 9, l9
    NFadeLm 10, l10a
    NFadeLm 10, l10b
    NFadeLm 10, l10c
    NFadeLm 10, l10d
    NFadeL 10, l10
    NFadeLm 11, l11a
    NFadeLm 11, l11b
    NFadeLm 11, l11c
    NFadeLm 11, l11d
    NFadeLm 11, l11e
    NFadeL 11, l11
    NFadeL 12, l12
    NFadeL 13, l13
    NFadeL 14, l14
    NFadeL 15, l15
    NFadeL 16, l16

    NFadeObjm 17, l17, "LEOn", "LE" ', "LEA", "LEB"
  Flash   17, l24F

    NFadeObjm 18, l18, "TAOn", "TA" ', "TAA", "TAB
  Flash   18, l23F

    NFadeLm 19, l19e
    NFadeLm 19, l19d
    NFadeLm 19, l19c
    NFadeLm 19, l19b
    NFadeLm 19, l19a
    NFadeL 19, l19
  NFadeObjm 20, L20P, "LRampPlOn", "LRampPl"
  NFadeLm 20, L20
  NFadeLm 20, L20b
  NFadeLm 20, l20c
  Flash   20, L20f

    NFadeObjm 21, l21, "LshOn", "Lsh" ', "LshA", "LshB"
  Flash   21, l22F

    NFadeObjm 22, l22, "LsOn", "Ls" ', "LsA", "LsB"
  Flash   22, l21F

    NFadeObjm 23, l23, "LLOn", "LL" ', "LLA", "LLB"
  Flash   23, l18F

    NFadeObjm 24, l24, "cbL1On", "cbL1" ', "cbL1A", "cbL1B"
  Flashm  24, l17F1
  Flashm  24, l17F2
  Flash   24, l17F
    Flashm 25, l25a
    Flash 25, l25
    Flashm 26, l26a
    Flash 26, l26
    Flashm 27, l27a
    Flash 27, l27

    Flashm 30, l30a
    Flash 30, l30
    Flashm 31, l31a
    Flash 31, l31
    Flashm 32, l32a
    Flash 32, l32
    NFadeLm 33, l33
    Flash 33, l33F
    NFadeLm 34, l34
    Flash 34, l34F
    NFadeLm 35, l35
    Flash 35, l35F
    NFadeL 36, l36
    NFadeL 37, l37
'    NFadeL 38, l38 -> GI 8AB?
    NFadeLm 39, l39a
    NFadeLm 39, l39
    Flash 39, l39F
    NFadeL 40, l40
    NFadeLm 41, l41
    Flash 41, l41F
    NFadeLm 42, l42
    Flash 42, l42F
    NFadeLm 43, l43
    Flash 43, l43F
    NFadeL 44, l44
    NFadeL 45, l45
'    NFadeL 46, l46 -> GI 9AB
    NFadeLm 47, l47
    NFadeLm 47, l47a
    Flash 47, l47F
    NFadeLm 48, l48a
    NFadeLm 48, l48b
    NFadeLm 48, l48
    Flash 48, l48F
    NFadeLm 49, l49
    Flash 49, l49F
    NFadeLm 50, l50
    Flash 50, l50F
    NFadeLm 51, l51
    Flash 51, l51F
    NFadeL 52, l52
    NFadeLm 53, l53
    Flash 53, l53F
    NFadeL 54, l54
    NFadeLm 55, l55b
    NFadeLm 55, l55A
  NFadeLm 55, l182a
  NFadeLm 55, l182b
  NFadeLm 55, l182c
  NFadeLm 55, l182d
    NFadeL 55, l55
    NFadeLm 56, l56
    NFadeLm 56, l56a
    NFadeLm 56, l56b
    NFadeLm 56, l56c
    NFadeL 56, l56d
    NFadeLm 57, l57
    NFadeLm 57, l57a
    NFadeLm 57, l57b
    Flash 57, l57F
    NFadeLm 58, l58
    Flash 58, l58F
    NFadeLm 59, l59
    NFadeLm 59, l59a
    NFadeLm 59, l59b
    NFadeLm 59, l59c
    Flash 59, l59F
    NFadeL 60, l60
    NFadeLm 61, l61
    Flash 61, l61F
    NFadeLm 62, l62
    NFadeLm 62, l62a
    Flash 62, l62F
    NFadeL 63, l63
    NFadeLm 64, l64
    NFadeLm 64, l64a
    NFadeLm 64, l64b
    NFadeLm 64, l64c
    NFadeLm 64, l64e
    NFadeL 64, l64d
'    NFadeL 65, l65
'    NFadeL 66, l66
'    NFadeL 67, l67
'    NFadeL 68, l68
'    NFadeL 71, l71
'    NFadeL 72, l72
'    NFadeL 73, l73
'    NFadeL 74, l74
'    NFadeL 75, l75
    NFadeL 76, l76
    NFadeL 77, l77
    NFadeL 78, l78
    NFadeL 79, l79
    NFadeLm 80, l80
    NFadeLm 81, l81
    Flash 81, l81F
    NFadeL 82, l82
'    NFadeL 83, l83
'    NFadeL 84, l84
    NFadeL 85, l85
    NFadeL 86, l86
  BlendObjm 87, l87p
  'NFadeObjm 87, L87P, "bulbcover1_redOn", "bulbcover1_red"
    NFadeL 87, l87l
  BlendObjm 88, l88p
  'NFadeObjm 88, L88P, "bulbcover1_redOn", "bulbcover1_red"
    NFadeL 88, l88L
    NFadeL 89, l89
    NFadeL 90, l90
    NFadeL 91, l91
    NFadeL 94, l94
    NFadeL 95, l95
    NFadeL 96, l96
  BlendObjm 97, l97p
    NFadeObjm 97, L97P, "bulbcover1_greenOn", "bulbcover1_green"
  Flash   97, L97f
  BlendObjm 98, l98p
    NFadeObjm 98, L98P, "bulbcover1_whiteOn", "bulbcover1_white"
  Flash   98, L98f
  BlendObjm 99, l99p
    NFadeObjm 99, L99P, "bulbcover1_greenOn", "bulbcover1_green"
  Flash   99, L99f
    Flash 100, l100F
    NFadeL 103, l103
    NFadeL 104, l104
    NFadeL 109, l109
    NFadeL 112, l112
    NFadeL 121, l121
  BlendObjm 113, l113p
    NFadeObjm 113, L113P, "bulbcover1_greenOn", "bulbcover1_green"
  Flash   113, L113f
  BlendObjm 115, l114p
    NFadeObjm 115, L114P, "TopFlasherWhiteA", "TopFlasherWhite"
  Flash   115, L114
  BlendObjm 117, l115p
  NFadeObjm 117, L115P, "bulbcover1_blueOn", "bulbcover1_blue"
  Flash   117, L115f
  BlendObjm 116, l116p
  NFadeObjm 116, L116P, "bulbcover1_yellowOn", "bulbcover1_yellow"
  Flash   116, L116f
  BlendObjm 114, l117p
  NFadeObjm 114, L117P, "bulbcover1_redOn", "bulbcover1_red"
  Flash   114, L117f
  FadeObj 118, HandP, "hanwithringOn", "hanwithringA", "hanwithringB", "hanwithring"
  NFadeLm 176, L176
  Flash   176, L176f
  NFadeLm 177, L13
  Flash   177, L177f
  Blend2Objm 178, l178p
  NFadeMatm 178, L178p, "flasheron", "flasheroff"
  NFadeObjm 178, L178P, "dome_green_on", "dome_green_off"
  NFadeLm 178, L178L
  Flashm  178, L178f2
  Blend2Objm 178, l179p
  NFadeMatm 178, L179p, "flasheron", "flasheroff"
  NFadeObjm 178, L179P, "dome_red_on", "dome_red_off"
  NFadeLm 178, L179L
  Flash   178, L179f
  NFadeLm 179, L179s2
  NFadeL 179, L179s1

  NFadeL 180, l180a

    NFadeLm 181, l181
    NFadeLm 181, F181a
  Blend2Objm 181, F181
    FadeObj 181, F181, "FlasherTestOn", "FlasherTestA", "FlasherTestB", "FlasherTest"
  NFadeLm 182, L182
  Flash   182, L182fa


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

Sub BlendObjm(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.blenddisablelighting = 0.3
        Case 5:object.blenddisablelighting = 1
        Case 9:object.blenddisablelighting = 0.3
        Case 13:object.blenddisablelighting = 0
    End Select
End Sub

Sub Blend2Objm(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.blenddisablelighting = 0.1
        Case 5:object.blenddisablelighting = 0.3
        Case 9:object.blenddisablelighting = 0.1
        Case 13:object.blenddisablelighting = 0
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

Sub NFadeMatm(nr, object, a, b)
    Select Case FadingLevel(nr)
        Case 4:object.material = b
        Case 5:object.material = a
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

' Desktop Objects: Reels & texts (you may also use lights on the desktop)

' Reels

Sub FadeR(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.SetValue 1:FadingLevel(nr) = 6                   'fading to off...
        Case 5:object.SetValue 0:FadingLevel(nr) = 1                   'ON
        Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1             'wait
        Case 9:object.SetValue 2:FadingLevel(nr) = FadingLevel(nr) + 1 'fading...
        Case 10, 11, 12:FadingLevel(nr) = FadingLevel(nr) + 1          'wait
        Case 13:object.SetValue 3:FadingLevel(nr) = 0                  'Off
    End Select
End Sub

Sub FadeRm(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.SetValue 1
        Case 5:object.SetValue 0
        Case 9:object.SetValue 2
        Case 3:object.SetValue 3
    End Select
End Sub

'Texts

Sub NFadeT(nr, object, message)
    Select Case FadingLevel(nr)
        Case 4:object.Text = "":FadingLevel(nr) = 0
        Case 5:object.Text = message:FadingLevel(nr) = 1
    End Select
End Sub

Sub NFadeTm(nr, object, b)
    Select Case FadingLevel(nr)
        Case 4:object.Text = ""
        Case 5:object.Text = message
    End Select
End Sub

Sub ChangeGlow(day)
  Dim Light
  If day Then
    For Each Light in GlowLights : Light.IntensityScale = GlowAmountDay: Light.FadeSpeedUp = Light.Intensity * GlowAmountDay / 2: Light.FadeSpeedDown = Light.FadeSpeedUp / 5: Next 'fadespeeddown = fadespeedup / 25
    For Each Light in InsertLights : Light.IntensityScale = InsertBrightnessDay : Light.FadeSpeedUp = Light.Intensity * InsertBrightnessDay / 2 : Light.FadeSpeedDown = Light.FadeSpeedUp / 5 : Next 'fadespeeddown = fadespeedup / 25
  Else
    For Each Light in GlowLights : Light.IntensityScale = GlowAmountNight: Light.FadeSpeedUp = Light.Intensity * GlowAmountNight / 2: Light.FadeSpeedDown = Light.FadeSpeedUp / 5: Next 'fadespeeddown = fadespeedup / 25
    For Each Light in InsertLights : Light.IntensityScale = InsertBrightnessNight : Light.FadeSpeedUp = Light.Intensity * InsertBrightnessNight / 2 : Light.FadeSpeedDown = Light.FadeSpeedUp / 5 : Next 'fadespeeddown = fadespeedup / 25
  End If
End Sub

Sub ColorGrade()
  Dim lutlevel, ContrastLut
  Lutlevel = 0
  If ContrastSetting=0 or ContrastSetting=2 or ContrastSetting=4 or ContrastSetting=6 Then
    ChangeGlow(True)
    ContrastLut = ContrastSetting / 2
  Else
    ChangeGlow(False)
    ContrastLut = ContrastSetting / 2 - 0.5
  End If
  table1.ColorGradeImage = "LUT" & ContrastLut & "_" & lutlevel
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

Sub PlaySoundAtExisting(soundname, tableobj)
    PlaySound soundname, 1, 1, AudioPan(tableobj), 0,0,1, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname)
    PlaySoundAt soundname, ActiveBall
End Sub

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / table1.width-1
    If tmp> 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

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
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) + (ball.velz^2) ) )
End Function

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 1000, Pan(ball1), 0, Pitch(ball1), 0, 0
End Sub

'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 9 ' total number of balls
ReDim rolling(tnob)
InitRolling

Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5,BallShadow6,BallShadow7,BallShadow8,BallShadow9)

Sub BallShadowUpdate()
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
    BallShadow(b).X = BOT(b).X
    ballShadow(b).Y = BOT(b).Y + 10
        If BOT(b).Z > 120 and BOT(b).Z < 130 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next

  If CheckInPlay(eyeBall1) Then
    CheckXLocation(eyeBall1)
  ElseIf CheckInPlay(eyeBall2) Then
    CheckXLocation(eyeBall2)
  ElseIf CheckInPlay(eyeBall3) Then
    CheckXLocation(eyeBall3)
  Else
    'CheckXLocation(eyeBall1)
  End If

End Sub


Function CheckInPlay(xball)
  If InRect(xball.x, xball.y, 0, 0, 952, 0, 952, 2096, 0, 2096) and Not (InRect(xball.x, xball.y, 25, 1100, 75, 1100, 75, 1840, 25, 1840) and xball.z > 177) and not InRect(xball.x, xball.y, 378, 2096, 823, 1800, 856, 1868, 536, 2096) Then
    CheckInPlay = True
  Else
    CheckInPlay = False
  End If
End Function


Sub CheckXLocation(xball)
  eyesl.transx = (xball.x - table1.width/2)/100 - (xball.y - table1.height/2)/400
  eyesl.transz = -(xball.x - table1.width/2)/100 - (xball.y - table1.height/2)/400
  eyesr.transx = (xball.x - table1.width/2)/100 + (xball.y - table1.height/2)/400
  eyesr.transz = (xball.x - table1.width/2)/100 - (xball.y - table1.height/2)/400
End Sub


Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Dim SkillVel

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
        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 130 Then
      StopSound("plasticroll" & b)
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b))*RollingSoundFactor, AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
        ElseIf BallVel(BOT(b) ) > 1 AND BOT(b).z > 130 Then
      StopSound("fx_ballrolling" & b)
            rolling(b) = True
            PlaySound("plasticroll" & b), -1, Vol(BOT(b))*RollingSoundFactor*10, AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
        Else
            If rolling(b) = True Then
        StopSound("plasticroll" & b)
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If
    If BOT(b).z > 220 Then
      If InRect(BOT(b).x, BOT(b).y, 825,896,869,958,814,995,772,935) and BOT(b).z > 285 Then
        SkillVel = BallVel(BOT(b))
      ElseIf InRect(BOT(b).x, BOT(b).y, 684,980,720,966,770,1035,688,1060) Then
        BOT(b).velx = - 2 * SkillVel
      Elseif InRect(BOT(b).x, BOT(b).y, 360,0,860,0,860,50,360,50) and BOT(b).velx < 2 Then
        BOT(b).velx = BOT(b).velx*1.1
      Elseif InRect(BOT(b).x, BOT(b).y, 900,200,950,200,950,840,900,840) and BOT(b).vely < 7.5 Then
        BOT(b).vely = 7.5
      Elseif InRect(BOT(b).x, BOT(b).y, 900,840,950,840,950,1480,900,1480) and BOT(b).vely < 7.5 and BOT(b).z > 320 Then
        BOT(b).vely = 7.5
      End If
    End If
    Next
End Sub

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

Sub RandomSoundFlipper()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall)*VolumeDial, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall)*VolumeDial, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall)*VolumeDial, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

Sub Balldrop1_Hit()
    PlaySoundAt "Balldrop", Balldrop1
End Sub

Sub Balldrop2_Hit()
    PlaySoundAt "Balldrop", Balldrop2
End Sub

Sub Scoop_Hit()
    PlaySoundAt "Scoop_Enter", scoop
End Sub

'*** Spiral Ramp Sounds ***
Dim SpiralBall

Sub SpiralRamp_hit()
  if Activeball.vely < 0 Then
    Set SpiralBall = ActiveBall
    SpiralRamp.timerenabled = true
  Else
    SpiralRamp.timerenabled = false
    StopSound "WireRamp"
  End If
End Sub

Sub SpiralRamp_Timer()
  PlaySound("WireRamp"), -1, Vol(SpiralBall)*VolumeDial*10, AudioPan(SpiralBall), 0, Pitch(SpiralBall), 1, 0, AudioFade(SpiralBall)

  'debug.print SpiralBall.x & chr(9) & spiralball.y & chr(9) & SpiralBall.z & chr(9) & SpiralBall.velx  & chr(9) & SpiralBall.vely & chr(9) & SpiralBall.velz
  If Not InRect(SpiralBall.x,SpiralBall.y,870,845,970,845,970,1060,870,1060) and SpiralBall.z < 170 Then
    me.timerenabled = False
    StopSound "WireRamp"
    playsoundat "balldrop", SpiralBall
  end if
End Sub


 '*****************************************************************
 'Functions
 '*****************************************************************

'*** PI returns the value for PI

Function PI()

  PI = 4*Atn(1)

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
'   rothbauerw's Manual Ball Control
'*****************************************

Dim BCup, BCdown, BCleft, BCright
Dim contBall, ControlBallInPlay, ControlActiveBall
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

Sub BallControlTimer()
  If contBall and EnableBallControl and ControlBallInPlay then
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

' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


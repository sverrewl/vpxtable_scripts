'Monte Carlo
'By Gottlieb/Premier
'1987

Option Explicit
Dim VRPosterR
Dim VRPosterL
Dim VRLogo
Dim Glass
Dim GlassScratch
Dim BackglassSelect
Dim BladeSelect

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Dim TableWidth  : TableWidth  = Table1.Width
Dim TableHeight : TableHeight = Table1.Height

Const Angle = 20
Const ReflipAngle = 20

Const BallSize = 50
Const BallMass = 1

'Dim UseVPMDMD
'UseVPMDMD = 1

'********************* VR AND FSS OPTIONS ******************************************
'***********************************************************************************
'VR Logo - Set VRLogo = 0 to turn off VR Logo.
VRLogo = 1

'VR Poster-Right - Set VRPosterR = 0 to turn off VR Poster-Right.
VRPosterR = 1

'VR Poster-Left - Set VRPosterL = 0 to turn off VR Poster-Left.
VRPosterL = 1

'VR and FSS table Glass - Set Glass = 0 to turn off VR playfield Glass.
Glass = 1

'VR and FSS Playfield glass Scratches - Set to 0 if you want to turn them off.
GlassScratch = 1

'VR and FSS select backglass - Set 0 for NOS, set 1 for alternative.
BackglassSelect = 0

'For all views blade select - Set 0 for black marble, set 1 for casino chips.
BladeSelect = 0

'***********************************************************************************
'******** Copy from this green line to next green line and insert it before the Table1_init *******

' VRRoom set based on RenderingMode

Dim VRRoom, DesktopMode, Obj4
Dim VRRoomChoice : VRRoomChoice = 0 '0 - Minimal Room (only applies when using VR headset)
DesktopMode = Table1.ShowDT
Dim VRTest
VRTest = 0
If RenderingMode = 2 or vrtest =1 Then VRRoom = VRRoomChoice + 1 Else VRRoom = 0

If VRRoom=1 Then
  for each Obj4 in ColRoomMinimal : Obj4.visible = 1 : next
  for each Obj4 in ColBackdrop : Obj4.visible = 0 : next
End If
If VRRoom=0 Then
  for each Obj4 in ColRoomMinimal : Obj4.visible = 0 : next
  for each Obj4 in ColBackdrop : Obj4.visible = 1 : next
End If

'******************************************************************************************************

if VRPosterR = 1 then VR_Poster1R.visible = true:VR_Poster2R.visible = true
if VRPosterR = 0 then VR_Poster1R.visible = false:VR_Poster2R.visible = false

if VRPosterL = 1 then VR_Poster3L.visible = true:VR_Poster4L.visible = true
if VRPosterL = 0 then VR_Poster3L.visible = false:VR_Poster4L.visible = false

if VRLogo = 1 then VR_Logo.visible = true
if VRLogo = 0 then VR_Logo.visible = false

If BackglassSelect = 1 then PinCab_Backglass.Image = "backglassalt" Else PinCab_Backglass.Image = "mcbackglass"
If BladeSelect = 1 then PinCab_Blades.Image = "PinCab_Blades_Chips" Else PinCab_Blades.Image = "PinCab_Blades_Marble"

'**************************************************************

'******************************************************************************************************
'If VR or FSS is selected then switch between desktop segments to VR/FSS segments
'******************************************************************************************************
Dim segs
If VRRoom <> 0 or Table1.ShowFSS = -1 then
  UpdateLedsF.Enabled = True
  For Each segs in VRFSSsegments
    'segs.Visible = 1
    segs.Y = -110
  Next
  For Each segs in dtdisplay
    segs.Visible = 0
  Next
Else
  DisplayTimer.Enabled = True
  For Each segs in dtdisplay
    segs.Visible = 1
  Next
End If

Dim VarRol,VarHidden, loopvar
If Table1.ShowDT = true and Table1.ShowFSS = 0 then
  VarRol=0
  VarHidden=1
  If VRRoom = 0 then
    For Each loopvar in dtdisplay
      loopvar.Visible = True
    Next
  End If
  If VRRoom = 0 then VR_Poster1R.Visible=0:VR_Poster3L.Visible=0:VR_Logo.Visible=0
  DisplayTimer.Enabled=True
Else
  VarRol=1
  VarHidden=0
  For Each loopvar in dtdisplay
    loopvar.Visible = False
  Next
  DisplayTimer.Enabled=False
End If

If B2SOn = true or Table1.ShowFSS = -1 Then
  VarHidden=1
End If


Const cGameName="mntecrlo",UseSolenoids=1,UseLamps=1,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff",SFlipperOn="",SFlipperOff=""
Const SCoin="",cCredits=""

LoadVPM "01210000","sys80.vbs",3.10

'Sub LoadVPM(VPMver, VBSfile, VBSver)
' On Error Resume Next
'   If ScriptEngineMajorVersion < 5 Then MsgBox "VB Script Engine 5.0 or higher required"
'   ExecuteGlobal GetTextFile(VBSFile)
'   If Err Then MsgBox "Unable to open " & VBSfile & ". Ensure that it is in the same folder as this table. " & vbNewLine & Err.Description : Err.Clear
'   Set Controller = CreateObject("VPinMAME.Controller")
'   If Err Then MsgBox "Can't Load VPinMAME." & vbNewLine & Err.Description
'   If VPMver>"" Then If Controller.Version < VPMver Or Err Then MsgBox "VPinMAME ver " & VPMver & " required." : Err.Clear
'   If VPinMAMEDriverVer < VBSver Or Err Then MsgBox VBSFile & " ver " & VBSver & " or higher required."
' On Error Goto 0
'End Sub

Set LampCallback=GetRef("UpdateMultipleLamps")

Sub UpdateMultipleLamps
  LI4=Light4.State 'BALL GATE
    If LI4<>OL4 Then
      IF LI4=1 Then
        Gate.RotateToEnd
      Else
        Gate.RotateToStart
      End If
    End If
  OL4=LI4
  LI12=Light12.State 'BALL RELEASE
  If LI12<>OL Then
    If LI12=0 Then
      If TroughBalls>0 Then
        BallRelease.CreateBall
        BallRelease.Kick 180,3
        RandomSoundBallRelease Ballrelease
        TroughBalls=TroughBalls-1
      End If
      If TroughBalls=0 Then Controller.Switch(42)=0
    End If
  End If
  OL=LI12
  LIT17=Light17a.State 'Spinning Wheel
  If LIT17<>OL17 Then
    If LIT17=1 Then
      'spinnertimer.Enabled = True
    'If Not PreEmpt.Enabled And Not TwoTimer.Enabled And Not BallCreated Then
      Short
      'PreEmpt.Enabled=1
    'End If
    End If
  End If
  OL17=LIT17
  LI13=Light13.State 'STARGATE RAMP
  If LI13<>OL13 Then
    Ramp12.HeightBottom = 160
    Ramp12.Collidable = False
  primitive012.visible=1
  primitive013.visible=0
  metalbar2.visible=0
  metalbar.visible=1
  rod1.visible=1
  rod1b.visible=1
  rod2.visible=0
  rod2b.visible=0
    'Kicker6.Enabled = False
    'Kicker7.Enabled = False
    If LI13=0 Then
    Ramp12.HeightBottom = 100
    Ramp12.Collidable = True
  primitive012.visible=0
  primitive013.visible=1
  metalbar2.visible=1
  metalbar.visible=0
  rod1.visible=0
  rod1b.visible=0
  rod2.visible=1
  rod2b.visible=1
    'Kicker6.Enabled = True
    'Kicker7.Enabled = True
  End If
  End If
  OL13=LI13
End Sub

'Sub PreEmpt_Timer
  'TTable.MotorOn=False
  'TWOTimer.Enabled=1
  'Me.Enabled=0
'End Sub

'Sub TwoTimer_Timer
'Me.Enabled=0
'End Sub

SolCallback(5)="dt123.SolDropUp"
SolCallback(1)="dt45.SolDropUp"
SolCallback(6)="dt678.SolDropUp"
SolCallback(2)="dt910.SolDropUp"
'solcallback(3)="K3.SolOut"
'solcallback(4)="K2.SolOut"
'Solcallback(7)="K1.SolOut"
SolCallback(3)="SolLock3Out"
SolCallback(4)="SolLock2Out"
SolCallback(7)="SolLock1Out"
SolCallback(8)="vpmSolSound""knocker"","
SolCallback(9)="SolOutHole"
SolCallback(sLLFlipper) = "SolLFlipper"
SolCallback(sLRFlipper) = "SolRFlipper"


'***************************************************************************************
' GAME OVER BLOCKER WALLS
'
' Monte Carlo ejects balls from the upper lock kickers both during multiball and during
' game-over cleanup. Since the same kickers are used for both situations, a custom
' detection system is used to determine whether the blockers should be raised or lowered.
'
' The collection "gameoverblockers" contains invisible walls that protect the bumpers
' and slingshots from stray balls rolling around the playfield after game over.
'
' HOW IT WORKS:
'
' 1. Gate1 is used to detect active gameplay. Whenever a ball passes through Gate1,
'    LastBallEvent is set to "release". This indicates that a ball has recently entered
'    normal play and that any upcoming lock eject is likely multiball-related.
'
' 2. Whenever a ball drains, LastBallEvent is set to "drain".
'
' 3. When a lock kicker eject occurs, CheckLockKickForGameOver examines LastBallEvent.
'
' 4. If the last event was "drain", the eject is assumed to be game-over cleanup.
'    The blocker walls are raised for 5 seconds using GameOverBlockerTimer, preventing
'    dead balls from reaching the bumpers and slingshots.
'
' 5. If the last event was "release", the eject is assumed to be the start of a
'    3-ball multiball. The blockers are lowered and MBGuardBalls is set to 2.
'
' 6. MBGuardBalls represents the two additional balls created by the 3-ball multiball.
'    Each drain decrements this value.
'
' 7. While MBGuardBalls is greater than zero, all lock kicker ejects are treated as
'    multiball activity and the blockers remain lowered. This prevents lock catch/
'    re-eject behavior during multiball from incorrectly raising the walls.
'
' 8. After two drains, MBGuardBalls reaches zero and future lock kicker ejects are
'    once again evaluated using LastBallEvent.
'
' TIMERS:
'
' GameOverBlockerTimer
'    Keeps the blocker walls raised for 5 seconds after a suspected game-over eject.
'
' COLLECTIONS:
'
' gameoverblockers
'    Collection of invisible walls that are raised/lowered using IsDropped.
'
'***************************************************************************************


Dim LastBallEvent
LastBallEvent = "release"
Dim MBGuardBalls
MBGuardBalls = 0

Sub Gate1_hit
LastBallEvent = "release"
End Sub

Sub SolLock1Out(Enabled)
  LeftKickArm002.RotZ = LeftKickArm002.RotZ + 16
  Arm1Timer.enabled=1
  K1.SolOut Enabled
  If Enabled Then CheckLockKickForGameOver
  'playsound "knocker_1"
End Sub

Sub Arm1Timer_timer
  LeftKickArm002.RotZ = 0
  Arm1Timer.enabled=0
End sub

Sub SolLock2Out(Enabled)
  LeftKickArm.RotZ = LeftKickArm.RotZ + 16
  Arm2Timer.enabled=1
  K2.SolOut Enabled
  If Enabled Then CheckLockKickForGameOver
End Sub

Sub Arm2Timer_timer
  LeftKickArm.RotZ = 0
  Arm2Timer.enabled=0
End sub

Sub SolLock3Out(Enabled)
  LeftKickArm001.RotZ = LeftKickArm001.RotZ + 16
  Arm3Timer.enabled=1
  K3.SolOut Enabled
  If Enabled Then CheckLockKickForGameOver
End Sub

Sub Arm3Timer_timer
  LeftKickArm001.RotZ = 0
  Arm3Timer.enabled=0
End sub

Sub DropGameOverBlockers()
  Dim obj
  For Each obj In gameoverblockers
    obj.IsDropped = True
  Next
End Sub

Sub RaiseGameOverBlockers()
  Dim obj
  For Each obj In gameoverblockers
    obj.IsDropped = False
  Next
End Sub

Sub CheckLockKickForGameOver()
    If MBGuardBalls > 0 Then Exit Sub

    If LastBallEvent = "drain" Then
        RaiseGameOverBlockers
        GameOverBlockerTimer.Interval = 5000
        GameOverBlockerTimer.Enabled = False
        GameOverBlockerTimer.Enabled = True
    Else
        DropGameOverBlockers
        MBGuardBalls = 2
    End If
End Sub

Sub GameOverBlockerTimer_Timer()
  DropGameOverBlockers
  GameOverBlockerTimer.Enabled = False
End Sub

'*****'depending if kicker kicks are multiball or game over, blocker walls will be raised to prevent bumper/slingshot hits

Sub MoveHammer(hammer, rotZ)
  If rotZ = 0 Then
    hammer.RotZ = 0
  Else
    LeftKickArm002.RotZ = LeftKickArm002.RotZ + 16
  End If
End Sub


Sub SolOutHole(Enabled)
If Enabled Then
  If DrainBalls=True Then
  TroughBalls=TroughBalls+1
  Controller.Switch(65)=0
  Controller.Switch(42)=1
  Drain.DestroyBall
  DrainBalls=False
  End If
End If
End Sub

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

Dim bsTrough,dt123,dt45,dt678,dt910,K1,K2,K3,TroughBalls,DrainBalls,LI4,OL4,LI13,OL13,LI12,OL,LIT17,OL17,Place,MyBall,X
TroughBalls=2:DrainBalls=True:OL=0:OL13=0:OL4=0:Place=0:OL17=0
Dim BallLocate,TTable,Frame,BallCreated
Dim kickmag,kickmag1
Frame=0:Place=0:BallCreated=False
'BallLocate=Array(B0,B1,B2,B3,B4,B5,B6,B7,B8,B9,B10,B11)

'*****************************************************************************************************
' VR PLUNGER ANIMATION
'
' Code needed to animate the plunger. If you pull the plunger it will move in VR.
' IMPORTANT: there are two numeric values in the code that define the postion of the plunger and the
' range in which it can move. The first numeric value is the actual y position of the plunger primitive
' and the second is the actual y position + 135 to determine the range in which it can move.
'
' You need to to select the Primary_plunger primitive you copied from the
' template and copy the value of the Y position
' (e.g. 2114) into the code. The value that determines the range of the plunger is always the y
' position + 135 (e.g. 2249).
'
'*****************************************************************************************************

Sub TimerPlunger_Timer
  If VR_Primary_plunger.Y < 2213 then
      VR_Primary_plunger.Y = VR_Primary_plunger.Y + 3
  End If
End Sub

Sub TimerPlunger2_Timer
 'debug.print plunger.position
  VR_Primary_plunger.Y = 2113 + (5* Plunger.Position) -20
End Sub

'*****************************************************************************************************

Sub Table1_Init
Controller.Games(cGameName).Settings.Value("dmd_red")=0
Controller.Games(cGameName).Settings.Value("dmd_green")=223
Controller.Games(cGameName).Settings.Value("dmd_blue")=223
'For X=0 To 11:BallLocate(X).IsDropped=1:RWall(X).IsDropped=1:Next
'Barrier.IsDropped=1
'Barrier2.IsDropped=1
'RWall(Frame).IsDropped=0
'BallLocate(Place).IsDropped=0
'RWheel Frame,Place,1
   With Controller
      .GameName=cGameName
      .SplashInfoLine="MONTE CARLO BY GOTTLIEB/PREMIER 1987"
    .HandleMechanics=0
    .HandleKeyboard=0
    .ShowDMDOnly=1
    .ShowFrame=0
    .ShowTitle=0
    .Hidden=VarHidden
    .Games(cGameName).Settings.Value("rol")=VarRol
    .Run
    If Err Then MsgBox Err.Description
  End With
  On Error Goto 0
   PinmameTimer.Interval=PinmameInterval
   PinmameTimer.Enabled=1

   vpmNudge.TiltSwitch=57
   vpmNudge.Sensitivity=5
   vpmNudge.TiltObj=Array(Bumper1a,Bumper2a,Bumper3a,LeftSlingshot,RightSlingshot)

  Set TTable=New cvpmTurnTable
  TTable.InitTurnTable TT,40
  TTable.SpinUp=40
  TTable.SpinDown=20

   Set dt123=New cvpmDropTarget
   dt123.InitDrop Array(Target1,Target2,Target3),Nothing
   dt123.InitSnd "Drop_Target_Down_1","Drop_Target_Reset_1"
   dt123.CreateEvents "dt123"

   Set dt45=New cvpmDropTarget
   dt45.InitDrop Array(Target4,Target5),Nothing
   dt45.initsnd "Drop_Target_Down_2","Drop_Target_Reset_2"
   dt45.CreateEvents "dt45"

   Set dt678=New cvpmDropTarget
   dt678.InitDrop Array(Target6,Target7,Target8),Nothing
   dt678.InitSnd "Drop_Target_Down_3","Drop_Target_Reset_3"
   dt678.CreateEvents "dt678"

   Set dt910=New cvpmDropTarget
   dt910.InitDrop Array(Target9,Target10),Nothing
   dt910.InitSnd "Drop_Target_Down_4","Drop_Target_Reset_4"
   dt910.CreateEvents "dt910"

   Set K1=New cvpmBallstack
   K1.InitSaucer Kicker1,45,184,1
   K1.KickAngleVar=5
   K1.InitExitSnd "saucer_kick","solenoid"

   Set K2=New cvpmBallstack
   K2.InitSaucer Kicker2,35,100,1
   K2.KickAngleVar=5
   K2.InitExitSnd "saucer_kick","solenoid"

   Set K3=new cvpmBallstack
   K3.InitSaucer Kicker3,25,270,1
   K3.KickAngleVar=5
   K3.InitExitSnd "saucer_kick","solenoid"

  Set kickmag=new cvpmMagnet
  kickmag.InitMagnet Trigger1, 14
  kickmag.GrabCenter = True

  Set kickmag1=new cvpmMagnet
  kickmag1.InitMagnet Trigger2, 14
  kickmag1.GrabCenter = True

Controller.Switch(42)=1
Controller.Switch(65)=1

vpmMapLights AllLights
vpmCreateEvents AllSwitches
' primitive013.visible=0
DropGameOverBlockers 'used for raising and lowering walls that prevent bumper/sling hits
Set MyBall = SW10.CreateBall
SW10.Kick 0,1
End Sub

'drop target shadows
Sub DTShadowTimer_Timer()
  Shadow_DT1.Visible = Target1.IsDropped = False
  Shadow_DT2.Visible = Target2.IsDropped = False
  Shadow_DT3.Visible = Target3.IsDropped = False
  Shadow_DT4.Visible = Target4.IsDropped = False
  Shadow_DT5.Visible = Target5.IsDropped = False
  Shadow_DT6.Visible = Target6.IsDropped = False
  Shadow_DT7.Visible = Target7.IsDropped = False
  Shadow_DT8.Visible = Target8.IsDropped = False
End Sub
'end drop target shadows

Sub Table1_exit()

  Controller.Stop

End Sub

Sub Table1_KeyDown(ByVal KeyCode)

'************************************************************

  If keycode = LeftFlipperKey Then FlipperActivate LeftFlipper, LFPress
  If keycode = RightFlipperKey Then FlipperActivate RightFlipper, RFPress

  If keycode = AddCreditKey Or keycode = AddCreditKey2 Then
    Select Case Int(Rnd * 3)
      Case 0: PlaySound "Coin_In_1", 0, CoinSoundLevel, 0, 0.25
      Case 1: PlaySound "Coin_In_2", 0, CoinSoundLevel, 0, 0.25
      Case 2: PlaySound "Coin_In_3", 0, CoinSoundLevel, 0, 0.25
    End Select
  End If

  If KeyCode = LeftFlipperKey Then
    VR_CabFlipperLeft.X = VR_CabFlipperLeft.X +10
  End If
  If KeyCode = RightFlipperKey Then
'primitive012.visible=1
'primitive013.visible=0
'metalbar2.visible=0
'metalbar.visible=1
'rod1.visible=1
'rod1b.visible=1
'rod2.visible=0
'rod2b.visible=0
    ' Animate VR Right flipper
    VR_CabFlipperRight.X = VR_CabFlipperRight.X - 10
  End If
  If Keycode = StartGameKey Then
'DropGameOverBlockers
    ' Animate VR Startbutton
    VR_Cab_StartButton.y = VR_Cab_StartButton.y -5
  End If
  If keycode = PlungerKey Then
'RaiseGameOverBlockers
    Plunger.Pullback
    TimerPlunger.Enabled = True
    TimerPlunger2.Enabled = False
  End If

'************************************************************
  If KeyCode=RightFlipperKey Then Controller.Switch(75)=1
  If KeyCode=PlungerKey Then Plunger.Pullback
  If vpmKeyDown(KeyCode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
'************************************************************
  If keycode = LeftFlipperKey Then FlipperDeActivate LeftFlipper, LFPress
  If keycode = RightFlipperKey Then FlipperDeActivate RightFlipper, RFPress

  If KeyCode = LeftFlipperKey Then
    ' Animate VR Left flipper
    VR_CabFlipperLeft.X = VR_CabFlipperLeft.X -10
  End If
  If KeyCode = RightFlipperKey Then
    ' Animate VR Right flipper
    VR_CabFlipperRight.X = VR_CabFlipperRight.X +10
  End If
  If keycode = PlungerKey Then
    Plunger.Fire
    TimerPlunger.Enabled = False
    TimerPlunger2.Enabled = True
    VR_Primary_plunger.Y = 1098
  End If
  If Keycode = StartGameKey Then
    ' Animate VR Startbutton
    VR_Cab_StartButton.y = VR_Cab_StartButton.y +5
  End If
'************************************************************
  If KeyCode=RightFlipperKey Then Controller.Switch(75)=0
  If KeyCode=PlungerKey Then Plunger.Fire:PlaySound "plunger"
  If vpmKeyUp(KeyCode) Then Exit Sub
End Sub

Sub Drain_Hit
  RandomSoundDrain Drain
  LastBallEvent = "drain"
    If MBGuardBalls > 0 Then MBGuardBalls = MBGuardBalls - 1 'helps with blocker that block bumpers/slings at game over
  Controller.Switch(65)=1
  DrainBalls=True
End Sub
Sub Bumper1a_Hit:vpmTimer.PulseSw 20:RandomSoundBumperTop Bumper1a:End Sub
Sub Bumper2a_Hit:vpmTimer.PulseSw 30:RandomSoundBumperMiddle Bumper2a:End Sub
Sub Bumper3a_Hit:vpmTimer.PulseSw 40:RandomSoundBumperBottom Bumper3a:End Sub
Sub Kicker1_Hit:soundsaucerlock:K1.AddBall 0:kickmag.MagnetOn = False:End Sub
Sub Kicker2_Hit:soundsaucerlock:K2.AddBall 0:End Sub
Sub Kicker3_Hit:soundsaucerlock:K3.AddBall 0:End Sub
Sub LeftSlingshot_Slingshot()
  LS.VelocityCorrect ActiveBall
  RandomSoundSlingshotLeft Primitive39
  vpmTimer.PulseSw 53
  PlaySound SoundFX("",DOFContactors), 0,1, -0.05,0.05 '0,1, AudioPan(LeftSlingShot), 0.05,0,0,1,AudioFade(LeftSlingShot)
End Sub
Sub RightSlingshot_Slingshot
  RS.VelocityCorrect ActiveBall
  RandomSoundSlingshotRight EndPoint1RS
  vpmTimer.PulseSw 53
  PlaySound SoundFX("",DOFContactors), 0,1, 0.05,0.05 '0,1, AudioPan(RightSlingShot), 0.05,0,0,1,AudioFade(RightSlingShot)
End Sub
Dim MY
Sub Kicker4_Hit:MY=-INT(ActiveBall.VelY)/3:Me.DestroyBall:Kicker4.TimerEnabled=True:End Sub
Sub Kicker4_Timer:Kicker5.CreateBall:Kicker5.Kick 180,MY:Kicker4.TimerEnabled=False:End Sub
'Sub Kicker6_Hit:Me.DestroyBall:Kicker7.CreateBall:Kicker7.Kick 0,20:End Sub

'Sub Roulette_Timer
'If TTable.MotorOn=True Then
' RWall(Frame).IsDropped=1
' Frame=Frame+1
' If Frame>11 Then Frame=0
' RWall(Frame).IsDropped=0
'Else
' RWall(Frame).IsDropped=1
' Frame=Frame+1
' If Frame>11 Then Frame=0
' RWall(Frame).IsDropped=0
'If TTable.Speed<1 Then
' Kicker9.Enabled=1
' Kicker10.Enabled=1
' Kicker11.Enabled=1
' Kicker12.Enabled=1
' Kicker13.Enabled=1
' Kicker14.Enabled=1
' Kicker15.Enabled=1
' Kicker16.Enabled=1
' Kicker17.Enabled=1
' Kicker18.Enabled=1
' Kicker19.Enabled=1
' Kicker20.Enabled=1
' Roulette.Enabled=0
'End If
'End If
'End Sub

'Dim RWall
'RWall=Array(W1,W2,W3,W4,W5,W6,W7,W8,W9,W10,W11,W12)

Sub Short
'Barrier.IsDropped=0
'Barrier2.IsDropped=0
'TTable.MotorOn=True
'BallLocate(Place).IsDropped=1
'BallCreated=True
'PlaySound"popper"
'Select Case Place
'Case 0:Set MyBall=Kicker9.CreateBall:MyBall.Image="Small":Kicker9.Kick 90,5
'Case 1:Set MyBall=Kicker10.CreateBall:MyBall.Image="Small":Kicker10.Kick 120,5
'Case 2:Set MyBall=Kicker11.CreateBall:MyBall.Image="Small":Kicker11.Kick 150,5
'Case 3:Set MyBall=Kicker12.CreateBall:MyBall.Image="Small":Kicker12.Kick 180,5
'Case 4:Set MyBall=Kicker13.CreateBall:MyBall.Image="Small":Kicker13.Kick 210,5
'Case 5:Set MyBall=Kicker14.CreateBall:MyBall.Image="Small":Kicker14.Kick 240,5
'Case 6:Set MyBall=Kicker15.CreateBall:MyBall.Image="Small":Kicker15.Kick 270,5
'Case 7:Set MyBall=Kicker16.CreateBall:MyBall.Image="Small":Kicker16.Kick 300,6
'Case 8:Set MyBall=Kicker17.CreateBall:MyBall.Image="Small":Kicker17.Kick 330,7
'Case 9:Set MyBall=Kicker18.CreateBall:MyBall.Image="Small":Kicker18.Kick 0,8
'Case 10:Set MyBall=Kicker19.CreateBall:MyBall.Image="Small":Kicker19.Kick 30,7
'Case 11:Set MyBall=Kicker20.CreateBall:MyBall.Image="Small":Kicker20.Kick 60,6
'End Select
RWheel Frame,Place,0
'TTable.AddBall MyBall
'TTable.AffectBall MyBall
'Roulette_Timer
'Roulette.Enabled=1
End Sub

'Sub Kicker9_Hit:Place=0:BallLocate(Place).IsDropped=0:TTable.RemoveBall MyBall:Me.DestroyBall:Barrier.IsDropped=1:Barrier2.IsDropped=1:ResetKickers:RWheel Frame,Place,1:BallCreated=False:End Sub
'Sub Kicker10_Hit:Place=1:BallLocate(Place).IsDropped=0:TTable.RemoveBall MyBall:Me.DestroyBall:Barrier.IsDropped=1:Barrier2.IsDropped=1:ResetKickers:RWheel Frame,Place,1:BallCreated=False:End Sub
'Sub Kicker11_Hit:Place=2:BallLocate(Place).IsDropped=0:TTable.RemoveBall MyBall:Me.DestroyBall:Barrier.IsDropped=1:Barrier2.IsDropped=1:ResetKickers:RWheel Frame,Place,1:BallCreated=False:End Sub
'Sub Kicker12_Hit:Place=3:BallLocate(Place).IsDropped=0:TTable.RemoveBall MyBall:Me.DestroyBall:Barrier.IsDropped=1:Barrier2.IsDropped=1:ResetKickers:RWheel Frame,Place,1:BallCreated=False:End Sub
'Sub Kicker13_Hit:Place=4:BallLocate(Place).IsDropped=0:TTable.RemoveBall MyBall:Me.DestroyBall:Barrier.IsDropped=1:Barrier2.IsDropped=1:ResetKickers:RWheel Frame,Place,1:BallCreated=False:End Sub
'Sub Kicker14_Hit:Place=5:BallLocate(Place).IsDropped=0:TTable.RemoveBall MyBall:Me.DestroyBall:Barrier.IsDropped=1:Barrier2.IsDropped=1:ResetKickers:RWheel Frame,Place,1:BallCreated=False:End Sub
'Sub Kicker15_Hit:Place=6:BallLocate(Place).IsDropped=0:TTable.RemoveBall MyBall:Me.DestroyBall:Barrier.IsDropped=1:Barrier2.IsDropped=1:ResetKickers:RWheel Frame,Place,1:BallCreated=False:End Sub
'Sub Kicker16_Hit:Place=7:BallLocate(Place).IsDropped=0:TTable.RemoveBall MyBall:Me.DestroyBall:Barrier.IsDropped=1:Barrier2.IsDropped=1:ResetKickers:RWheel Frame,Place,1:BallCreated=False:End Sub
'Sub Kicker17_Hit:Place=8:BallLocate(Place).IsDropped=0:TTable.RemoveBall MyBall:Me.DestroyBall:Barrier.IsDropped=1:Barrier2.IsDropped=1:ResetKickers:RWheel Frame,Place,1:BallCreated=False:End Sub
'Sub Kicker18_Hit:Place=9:BallLocate(Place).IsDropped=0:TTable.RemoveBall MyBall:Me.DestroyBall:Barrier.IsDropped=1:Barrier2.IsDropped=1:ResetKickers:RWheel Frame,Place,1:BallCreated=False:End Sub
'Sub Kicker19_Hit:Place=10:BallLocate(Place).IsDropped=0:TTable.RemoveBall MyBall:Me.DestroyBall:Barrier.IsDropped=1:Barrier2.IsDropped=1:ResetKickers:RWheel Frame,Place,1:BallCreated=False:End Sub
'Sub Kicker20_Hit:Place=11:BallLocate(Place).IsDropped=0:TTable.RemoveBall MyBall:Me.DestroyBall:Barrier.IsDropped=1:Barrier2.IsDropped=1:ResetKickers:RWheel Frame,Place,1:BallCreated=False:End Sub

'Sub ResetKickers
' Kicker9.Enabled=0
' Kicker10.Enabled=0
' Kicker11.Enabled=0
' Kicker12.Enabled=0
' Kicker13.Enabled=0
' Kicker14.Enabled=0
' Kicker15.Enabled=0
' Kicker16.Enabled=0
' Kicker17.Enabled=0
' Kicker18.Enabled=0
' Kicker19.Enabled=0
' Kicker20.Enabled=0
'End Sub

'Gottlieb Monte Carlo
'added by Inkochnito
Sub editDips
  Dim vpmDips : Set vpmDips = New cvpmDips
  With vpmDips
    .AddForm  700,400,"Monte Carlo - DIP switches"
    .AddFrame 2,4,190,"Maximum credits",49152,Array("8 credits",0,"10 credits",32768,"15 credits",&H00004000,"20 credits",49152)'dip 15&16
    .AddFrame 2,80,190,"Coin chute 1 and 2 control",&H00002000,Array("seperate",0,"same",&H00002000)'dip 14
    .AddFrame 2,126,190,"Playfield special",&H00200000,Array("replay",0,"extra ball",&H00200000)'dip 22
    .AddFrame 2,172,190,"High games to date control",&H00000020,Array("no effect",0,"reset high games 2-5 on power off",&H00000020)'dip 6
    .AddFrame 2,218,190,"10.000.000 points light",&H00000080,Array("lights random",0,"lights after 40 switch hits",&H00000080)'dip 8
    .AddFrame 2,264,190,"Match number control",&H40000000,Array("match with display",0,"match with roulette",&H40000000)'dip 31
    .AddFrame 2,310,190,"Ball lock control",&H80000000,Array("one lock at the time",0,"all locks on",&H80000000)'dip 32
    .AddFrame 205,4,190,"High game to date awards",&H00C00000,Array("not displayed and no award",0,"displayed and no award",&H00800000,"displayed and 2 replays",&H00400000,"displayed and 3 replays",&H00C00000)'dip 23&24
    .AddFrame 205,80,190,"Balls per game",&H01000000,Array("5 balls",0,"3 balls",&H01000000)'dip 25
    .AddFrame 205,126,190,"Replay limit",&H04000000,Array("no limit",0,"one per game",&H04000000)'dip 27
    .AddFrame 205,172,190,"Novelty",&H08000000,Array("normal",0,"extra ball and replay scores 500K",&H08000000)'dip 28
    .AddFrame 205,218,190,"Game mode",&H10000000,Array("replay",0,"extra ball",&H10000000)'dip 29
    .AddFrame 205,264,190,"3rd coin chute credits control",&H20000000,Array("no effect",0,"add 9",&H20000000)'dip 30
    .AddChk 205,316,180,Array("Match feature",&H02000000)'dip 26
    .AddChk 205,331,190,Array("Attract sound",&H00000040)'dip 7
    .AddLabel 50,360,300,20,"After hitting OK, press F3 to reset game with new settings."
    .ViewDips
  End With
End Sub
Set vpmShowDips = GetRef("editDips")

Sub Trigger1_Hit()
  kickmag.MagnetOn = True
End Sub

Sub Trigger2_Hit()
  kickmag1.MagnetOn = True
End Sub

Sub flashertimer_Timer()
  If Light15.State=1 then
    Flasher2.Visible=True:LF2.state=1:Flasher3.Visible=True:Flasher1.Visible=True:LF1.state=1:Flasher4.Visible=True::bg3.Visible=False:bg4.Visible=False:bg5.Visible=False
  Else
    Flasher2.Visible=False:LF2.state=0:Flasher3.Visible=False:Flasher1.Visible=False:LF1.state=0:Flasher4.Visible=False:bg3.Visible=True:bg4.Visible=True:bg5.Visible=True
  End If
  If Light14.State=1 then
    Flasher5.Visible=True:Flasher6.Visible=True:LF6.state=1:Flasher7.Visible=True:Flasher8.Visible=True:LF8.state=1:bg1.Visible=False:bg2.Visible=False:bg6.Visible=False
  Else
    Flasher5.Visible=False:Flasher6.Visible=False:LF6.state=0:Flasher7.Visible=False:Flasher8.Visible=False:LF8.state=0:bg1.Visible=True:bg2.Visible=True:bg6.Visible=True
  End If

  If Light9.State=1 then Flasher9.Visible=True else Flasher9.Visible=False
  If Light24.State=1 then Flasher10.Visible=True else Flasher10.Visible=False
  If Light8.State=1 then Flasher11.Visible=True else Flasher11.Visible=False
  If Light7.State=1 then Flasher12.Visible=True else Flasher12.Visible=False
  If Light6.State=1 then Flasher13.Visible=True else Flasher13.Visible=False
  If Light5.State=1 then Flasher14.Visible=True else Flasher14.Visible=False
End Sub


'EVERYTHING FOR THE ROULETTE WHEEL IS BELOW
'-------------------------------------------------------------------------------------------
Dim LoopObj, spinctr, RouletteState, LastRouletteLight, RouletteLockout
Dim kickernum:kickernum = 0

Const ROULETTE_IDLE = 0
Const ROULETTE_SPINNING = 1
Const ROULETTE_INITIALKICK = 2
Const ROULETTE_FINAL = 3

spinctr = 0
RouletteState = ROULETTE_IDLE
LastRouletteLight = 0
RouletteLockout = False

Sub EnableRealRouletteKickers(aEnabled)
  SW00.Enabled = aEnabled
  SW01.Enabled = aEnabled
  SW02.Enabled = aEnabled
  SW03.Enabled = aEnabled
  SW04.Enabled = aEnabled
  SW05.Enabled = aEnabled
  SW10.Enabled = aEnabled
  SW11.Enabled = aEnabled
  SW12.Enabled = aEnabled
  SW13.Enabled = aEnabled
  SW14.Enabled = aEnabled
  SW15.Enabled = aEnabled
End Sub

Sub ClearRouletteSwitches()
  Controller.Switch(0)=0
  Controller.Switch(1)=0
  Controller.Switch(2)=0
  Controller.Switch(3)=0
  Controller.Switch(4)=0
  Controller.Switch(5)=0
  Controller.Switch(10)=0
  Controller.Switch(11)=0
  Controller.Switch(12)=0
  Controller.Switch(13)=0
  Controller.Switch(14)=0
  Controller.Switch(15)=0
End Sub

Sub ReleaseRouletteBall()
  SW00.Kick 90, 1
  SW01.Kick 120, 1
  SW02.Kick 150, 1
  SW03.Kick 180, 1
  SW04.Kick 210, 1
  SW05.Kick 240, 1
  SW10.Kick 270, 1
  SW11.Kick 300, 1
  SW12.Kick 330, 1
  SW13.Kick 0, 1
  SW14.Kick 30, 1
  SW15.Kick 60, 1
End Sub

Sub StartRoulette()
  If RouletteLockout Then Exit Sub
  If RouletteState <> ROULETTE_IDLE Then Exit Sub

  RouletteLockout = True
  RouletteLockoutTimer.Interval = 10000
  RouletteLockoutTimer.Enabled = False
  RouletteLockoutTimer.Enabled = True

  ClearRouletteSwitches
  RouletteFinalTimer.Enabled = False
  EnableRealRouletteKickers False

  ReleaseRouletteBall
  PlaySoundAt "roulette", spindisk
  DisableRouletteKickersHard

  RouletteState = ROULETTE_SPINNING

  TTable.MotorOn = True
  TTable.AddBall MyBall
  TTable.AffectBall MyBall
End Sub

Sub RouletteLockoutTimer_Timer()
  RouletteLockoutTimer.Enabled = False
  RouletteLockout = False
End Sub

Sub FireAllRouletteKickers()
  SW00.Kick 90  + Rnd * 20 - 10, 3 + Rnd * 5
  SW01.Kick 120 + Rnd * 20 - 10, 3 + Rnd * 5
  SW02.Kick 150 + Rnd * 20 - 10, 3 + Rnd * 5
  SW03.Kick 180 + Rnd * 20 - 10, 3 + Rnd * 5
  SW04.Kick 210 + Rnd * 20 - 10, 3 + Rnd * 5
  SW05.Kick 240 + Rnd * 20 - 10, 3 + Rnd * 5
  SW10.Kick 270 + Rnd * 20 - 10, 3 + Rnd * 5
  SW11.Kick 300 + Rnd * 20 - 10, 3 + Rnd * 5
  SW12.Kick 330 + Rnd * 20 - 10, 3 + Rnd * 5
  SW13.Kick 0   + Rnd * 20 - 10, 3 + Rnd * 5
  SW14.Kick 30  + Rnd * 20 - 10, 3 + Rnd * 5
  SW15.Kick 60  + Rnd * 20 - 10, 3 + Rnd * 5
End Sub

Sub StartRouletteInitialKick()
  If RouletteState <> ROULETTE_SPINNING Then Exit Sub

  RouletteState = ROULETTE_INITIALKICK

  TTable.RemoveBall MyBall
  FireAllRouletteKickers
  DisableRouletteKickersHard

  RouletteFinalTimer.Interval = 4000 + Int(Rnd * 2500) '4000-6499 ms
  RouletteFinalTimer.Enabled = False
  RouletteFinalTimer.Enabled = True
End Sub

Sub RouletteFinalTimer_Timer()
  RouletteFinalTimer.Enabled = False

  If RouletteState = ROULETTE_INITIALKICK Then
    RouletteState = ROULETTE_FINAL
    EnableRealRouletteKickers True
  End If
End Sub

Sub FinishRoulette()
  RouletteState = ROULETTE_IDLE
  RouletteFinalTimer.Enabled = False
  EnableRealRouletteKickers False
End Sub

Sub spinnertimer_Timer()
  If Light17a.State = 1 And LastRouletteLight = 0 And RouletteLockout = False Then
    toppertimer.Enabled = True
    w009.TimerEnabled = True
    StartRoulette
  End If

  LastRouletteLight = Light17a.State

  If RouletteState = ROULETTE_SPINNING Then
    Spindisk.RotZ = Spindisk.RotZ + 30
    Spindiskb.RotZ = Spindisk.RotZ
    cone.RotY = cone.RotY + 30
    sphere.RotY = sphere.RotY + 30

    If Spindisk.RotZ >= 360 Then
      Spindisk.RotZ = 0
      Spindiskb.RotZ = 0
      cone.RotY = -95
      sphere.RotY = 0
    End If

    spinctr = spinctr + 1
    If spinctr > 11 Then spinctr = 0

    If TTable.Speed >= 30 Then TTable.MotorOn = False

    If TTable.Speed <= 1 And TTable.MotorOn = False Then
      StartRouletteInitialKick
    End If
  End If
End Sub

Sub RouletteKickerHit(kickerIndex)
  If RouletteState <> ROULETTE_FINAL Then Exit Sub

  RWheel spinctr, kickerIndex, 1
  kickernum = kickerIndex
  FinishRoulette
End Sub

Sub SW00_Hit:Stopsound "roulette":PlaySoundAt "r" & (Int(Rnd * 4) + 3), spindisk:RouletteKickerHit 0:End Sub
Sub SW01_Hit:Stopsound "roulette":PlaySoundAt "r" & (Int(Rnd * 4) + 3), spindisk:RouletteKickerHit 1:End Sub
Sub SW02_Hit:Stopsound "roulette":PlaySoundAt "r" & (Int(Rnd * 4) + 3), spindisk:RouletteKickerHit 2:End Sub
Sub SW03_Hit:Stopsound "roulette":PlaySoundAt "r" & (Int(Rnd * 4) + 3), spindisk:RouletteKickerHit 3:End Sub
Sub SW04_Hit:Stopsound "roulette":PlaySoundAt "r" & (Int(Rnd * 4) + 3), spindisk:RouletteKickerHit 4:End Sub
Sub SW05_Hit:Stopsound "roulette":PlaySoundAt "r" & (Int(Rnd * 4) + 3), spindisk:RouletteKickerHit 5:End Sub
Sub SW10_Hit:Stopsound "roulette":PlaySoundAt "r" & (Int(Rnd * 4) + 3), spindisk:RouletteKickerHit 6:End Sub
Sub SW11_Hit:Stopsound "roulette":PlaySoundAt "r" & (Int(Rnd * 4) + 3), spindisk:RouletteKickerHit 7:End Sub
Sub SW12_Hit:Stopsound "roulette":PlaySoundAt "r" & (Int(Rnd * 4) + 3), spindisk:RouletteKickerHit 8:End Sub
Sub SW13_Hit:Stopsound "roulette":PlaySoundAt "r" & (Int(Rnd * 4) + 3), spindisk:RouletteKickerHit 9:End Sub
Sub SW14_Hit:Stopsound "roulette":PlaySoundAt "r" & (Int(Rnd * 4) + 3), spindisk:RouletteKickerHit 10:End Sub
Sub SW15_Hit:Stopsound "roulette":PlaySoundAt "r" & (Int(Rnd * 4) + 3), spindisk:RouletteKickerHit 11:End Sub

Sub DisableRouletteKickersHard()
  SW00.Enabled = False
  SW01.Enabled = False
  SW02.Enabled = False
  SW03.Enabled = False
  SW04.Enabled = False
  SW05.Enabled = False
  SW10.Enabled = False
  SW11.Enabled = False
  SW12.Enabled = False
  SW13.Enabled = False
  SW14.Enabled = False
  SW15.Enabled = False
End Sub

Dim colorandnum
Sub RWheel(GR,GS,GT)
TextBox006.Text = GT
Select Case GR
  Case 0:Select Case GS
      Case 0:Controller.Switch(0)=GT:colorandnum ="Green 0"
      Case 1:Controller.Switch(1)=GT:colorandnum ="Red 1"
      Case 2:Controller.Switch(2)=GT:colorandnum ="Red 2"
      Case 3:Controller.Switch(3)=GT:colorandnum ="Black 3"
      Case 4:Controller.Switch(4)=GT:colorandnum ="Black 4"
      Case 5:Controller.Switch(5)=GT:colorandnum ="Red 5"
      Case 6:Controller.Switch(10)=GT:colorandnum ="Green 00"
      Case 7:Controller.Switch(11)=GT:colorandnum ="Black 6"
      Case 8:Controller.Switch(12)=GT:colorandnum ="Black 7"
      Case 9:Controller.Switch(13)=GT:colorandnum ="Red 8"
      Case 10:Controller.Switch(14)=GT:colorandnum ="Red 9"
      Case 11:Controller.Switch(15)=GT:colorandnum ="Black 10"
    End Select
  Case 1:Select Case GS
      Case 0:Controller.Switch(15)=GT:colorandnum ="Black 10"
      Case 1:Controller.Switch(0)=GT:colorandnum ="Green 0"
      Case 2:Controller.Switch(1)=GT:colorandnum ="Red 1"
      Case 3:Controller.Switch(2)=GT:colorandnum ="Red 2"
      Case 4:Controller.Switch(3)=GT:colorandnum ="Black 3"
      Case 5:Controller.Switch(4)=GT:colorandnum ="Black 4"
      Case 6:Controller.Switch(5)=GT:colorandnum ="Red 5"
      Case 7:Controller.Switch(10)=GT:colorandnum ="Green 00"
      Case 8:Controller.Switch(11)=GT:colorandnum ="Black 6"
      Case 9:Controller.Switch(12)=GT:colorandnum ="Black 7"
      Case 10:Controller.Switch(13)=GT:colorandnum ="Red 8"
      Case 11:Controller.Switch(14)=GT:colorandnum ="Red 9"
    End Select
  Case 2:Select Case GS
      Case 0:Controller.Switch(14)=GT:colorandnum ="Red 9"
      Case 1:Controller.Switch(15)=GT:colorandnum ="Black 10"
      Case 2:Controller.Switch(0)=GT:colorandnum ="Green 0"
      Case 3:Controller.Switch(1)=GT:colorandnum ="Red 1"
      Case 4:Controller.Switch(2)=GT:colorandnum ="Red 2"
      Case 5:Controller.Switch(3)=GT:colorandnum ="Black 3"
      Case 6:Controller.Switch(4)=GT:colorandnum ="Black 4"
      Case 7:Controller.Switch(5)=GT:colorandnum ="Red 5"
      Case 8:Controller.Switch(10)=GT:colorandnum ="Green 00"
      Case 9:Controller.Switch(11)=GT:colorandnum ="Black 6"
      Case 10:Controller.Switch(12)=GT:colorandnum ="Black 7"
      Case 11:Controller.Switch(13)=GT:colorandnum ="Red 8"
    End Select
  Case 3:Select Case GS
      Case 0:Controller.Switch(13)=GT:colorandnum ="Red 8"
      Case 1:Controller.Switch(14)=GT:colorandnum ="Red 9"
      Case 2:Controller.Switch(15)=GT:colorandnum ="Black 10"
      Case 3:Controller.Switch(0)=GT:colorandnum ="Green 0"
      Case 4:Controller.Switch(1)=GT:colorandnum ="Red 1"
      Case 5:Controller.Switch(2)=GT:colorandnum ="Red 2"
      Case 6:Controller.Switch(3)=GT:colorandnum ="Black 3"
      Case 7:Controller.Switch(4)=GT:colorandnum ="Black 4"
      Case 8:Controller.Switch(5)=GT:colorandnum ="Red 5"
      Case 9:Controller.Switch(10)=GT:colorandnum ="Green 00"
      Case 10:Controller.Switch(11)=GT:colorandnum ="Black 6"
      Case 11:Controller.Switch(12)=GT:colorandnum ="Black 7"
    End Select
  Case 4:Select Case GS
      Case 0:Controller.Switch(12)=GT:colorandnum ="Black 7"
      Case 1:Controller.Switch(13)=GT:colorandnum ="Red 8"
      Case 2:Controller.Switch(14)=GT:colorandnum ="Red 9"
      Case 3:Controller.Switch(15)=GT:colorandnum ="Black 10"
      Case 4:Controller.Switch(0)=GT:colorandnum ="Green 0"
      Case 5:Controller.Switch(1)=GT:colorandnum ="Red 1"
      Case 6:Controller.Switch(2)=GT:colorandnum ="Red 2"
      Case 7:Controller.Switch(3)=GT:colorandnum ="Black 3"
      Case 8:Controller.Switch(4)=GT:colorandnum ="Black 4"
      Case 9:Controller.Switch(5)=GT:colorandnum ="Red 5"
      Case 10:Controller.Switch(10)=GT:colorandnum ="Green 00"
      Case 11:Controller.Switch(11)=GT:colorandnum ="Black 6"
    End Select
  Case 5:Select Case GS
      Case 0:Controller.Switch(11)=GT:colorandnum ="Black 6"
      Case 1:Controller.Switch(12)=GT:colorandnum ="Black 7"
      Case 2:Controller.Switch(13)=GT:colorandnum ="Red 8"
      Case 3:Controller.Switch(14)=GT:colorandnum ="Red 9"
      Case 4:Controller.Switch(15)=GT:colorandnum ="Black 10"
      Case 5:Controller.Switch(0)=GT:colorandnum ="Green 0"
      Case 6:Controller.Switch(1)=GT:colorandnum ="Red 1"
      Case 7:Controller.Switch(2)=GT:colorandnum ="Red 2"
      Case 8:Controller.Switch(3)=GT:colorandnum ="Black 3"
      Case 9:Controller.Switch(4)=GT:colorandnum ="Black 4"
      Case 10:Controller.Switch(5)=GT:colorandnum ="Red 5"
      Case 11:Controller.Switch(10)=GT:colorandnum ="Green 00"
    End Select
  Case 6:Select Case GS
      Case 0:Controller.Switch(10)=GT:colorandnum ="Green 00"
      Case 1:Controller.Switch(11)=GT:colorandnum ="Black 6"
      Case 2:Controller.Switch(12)=GT:colorandnum ="Black 7"
      Case 3:Controller.Switch(13)=GT:colorandnum ="Red 8"
      Case 4:Controller.Switch(14)=GT:colorandnum ="Red 9"
      Case 5:Controller.Switch(15)=GT:colorandnum ="Black 10"
      Case 6:Controller.Switch(0)=GT:colorandnum ="Green 0"
      Case 7:Controller.Switch(1)=GT:colorandnum ="Red 1"
      Case 8:Controller.Switch(2)=GT:colorandnum ="Red 2"
      Case 9:Controller.Switch(3)=GT:colorandnum ="Black 3"
      Case 10:Controller.Switch(4)=GT:colorandnum ="Black 4"
      Case 11:Controller.Switch(5)=GT:colorandnum ="Red 5"
    End Select
  Case 7:Select Case GS
      Case 0:Controller.Switch(5)=GT:colorandnum ="Red 5"
      Case 1:Controller.Switch(10)=GT:colorandnum ="Green 00"
      Case 2:Controller.Switch(11)=GT:colorandnum ="Black 6"
      Case 3:Controller.Switch(12)=GT:colorandnum ="Black 7"
      Case 4:Controller.Switch(13)=GT:colorandnum ="Red 8"
      Case 5:Controller.Switch(14)=GT:colorandnum ="Red 9"
      Case 6:Controller.Switch(15)=GT:colorandnum ="Black 10"
      Case 7:Controller.Switch(0)=GT:colorandnum ="Green 0"
      Case 8:Controller.Switch(1)=GT:colorandnum ="Red 1"
      Case 9:Controller.Switch(2)=GT:colorandnum ="Red 2"
      Case 10:Controller.Switch(3)=GT:colorandnum ="Black 3"
      Case 11:Controller.Switch(4)=GT:colorandnum ="Black 4"
    End Select
  Case 8:Select Case GS
      Case 0:Controller.Switch(4)=GT:colorandnum ="Black 4"
      Case 1:Controller.Switch(5)=GT:colorandnum ="Red 5"
      Case 2:Controller.Switch(10)=GT:colorandnum ="Green 00"
      Case 3:Controller.Switch(11)=GT:colorandnum ="Black 6"
      Case 4:Controller.Switch(12)=GT:colorandnum ="Black 7"
      Case 5:Controller.Switch(13)=GT:colorandnum ="Red 8"
      Case 6:Controller.Switch(14)=GT:colorandnum ="Red 9"
      Case 7:Controller.Switch(15)=GT:colorandnum ="Black 10"
      Case 8:Controller.Switch(0)=GT:colorandnum ="Green 0"
      Case 9:Controller.Switch(1)=GT:colorandnum ="Red 1"
      Case 10:Controller.Switch(2)=GT:colorandnum ="Red 2"
      Case 11:Controller.Switch(3)=GT:colorandnum ="Black 3"
    End Select
  Case 9:Select Case GS
      Case 0:Controller.Switch(3)=GT:colorandnum ="Black 3"
      Case 1:Controller.Switch(4)=GT:colorandnum ="Black 4"
      Case 2:Controller.Switch(5)=GT:colorandnum ="Red 5"
      Case 3:Controller.Switch(10)=GT:colorandnum ="Green 00"
      Case 4:Controller.Switch(11)=GT:colorandnum ="Black 6"
      Case 5:Controller.Switch(12)=GT:colorandnum ="Black 7"
      Case 6:Controller.Switch(13)=GT:colorandnum ="Red 8"
      Case 7:Controller.Switch(14)=GT:colorandnum ="Red 9"
      Case 8:Controller.Switch(15)=GT:colorandnum ="Black 10"
      Case 9:Controller.Switch(0)=GT:colorandnum ="Green 0"
      Case 10:Controller.Switch(1)=GT:colorandnum ="Red 1"
      Case 11:Controller.Switch(2)=GT:colorandnum ="Red 2"
    End Select
  Case 10:Select Case GS
      Case 0:Controller.Switch(2)=GT:colorandnum ="Red 2"
      Case 1:Controller.Switch(3)=GT:colorandnum ="Black 3"
      Case 2:Controller.Switch(4)=GT:colorandnum ="Black 4"
      Case 3:Controller.Switch(5)=GT:colorandnum ="Red 5"
      Case 4:Controller.Switch(10)=GT:colorandnum ="Green 00"
      Case 5:Controller.Switch(11)=GT:colorandnum ="Black 6"
      Case 6:Controller.Switch(12)=GT:colorandnum ="Black 7"
      Case 7:Controller.Switch(13)=GT:colorandnum ="Red 8"
      Case 8:Controller.Switch(14)=GT:colorandnum ="Red 9"
      Case 9:Controller.Switch(15)=GT:colorandnum ="Black 10"
      Case 10:Controller.Switch(0)=GT:colorandnum ="Green 0"
      Case 11:Controller.Switch(1)=GT:colorandnum ="Red 1"
    End Select
  Case 11:Select Case GS
      Case 0:Controller.Switch(1)=GT:colorandnum ="Red 1"
      Case 1:Controller.Switch(2)=GT:colorandnum ="Red 2"
      Case 2:Controller.Switch(3)=GT:colorandnum ="Black 3"
      Case 3:Controller.Switch(4)=GT:colorandnum ="Black 4"
      Case 4:Controller.Switch(5)=GT:colorandnum ="Red 5"
      Case 5:Controller.Switch(10)=GT:colorandnum ="Green 00"
      Case 6:Controller.Switch(11)=GT:colorandnum ="Black 6"
      Case 7:Controller.Switch(12)=GT:colorandnum ="Black 7"
      Case 8:Controller.Switch(13)=GT:colorandnum ="Red 8"
      Case 9:Controller.Switch(14)=GT :colorandnum ="Red 9"
      Case 10:Controller.Switch(15)=GT:colorandnum ="Black 10"
      Case 11:Controller.Switch(0)=GT:colorandnum ="Green 0"
    End Select
End Select
End Sub

'LED taken from Victory Table (Gottlieb1987) by Sinbad
'https://vpinball.com/VPBdownloads/victory-gottlieb-1987-2-0-1/

Dim Digits(40)
Digits(0) = Array(a00, a05, a0c, a0d, a08, a01, a06, a0f, a02, a03, a04, a07, a0b, a0a, a09, a0e)
Digits(1) = Array(a10, a15, a1c, a1d, a18, a11, a16, a1f, a12, a13, a14, a17, a1b, a1a, a19, a1e)
Digits(2) = Array(a20, a25, a2c, a2d, a28, a21, a26, a2f, a22, a23, a24, a27, a2b, a2a, a29, a2e)
Digits(3) = Array(a30, a35, a3c, a3d, a38, a31, a36, a3f, a32, a33, a34, a37, a3b, a3a, a39, a3e)
Digits(4) = Array(a40, a45, a4c, a4d, a48, a41, a46, a4f, a42, a43, a44, a47, a4b, a4a, a49, a4e)
Digits(5) = Array(a50, a55, a5c, a5d, a58, a51, a56, a5f, a52, a53, a54, a57, a5b, a5a, a59, a5e)
Digits(6) = Array(a60, a65, a6c, a6d, a68, a61, a66, a6f, a62, a63, a64, a67, a6b, a6a, a69, a6e)
Digits(7) = Array(a70, a75, a7c, a7d, a78, a71, a76, a7f, a72, a73, a74, a77, a7b, a7a, a79, a7e)
Digits(8) = Array(a80, a85, a8c, a8d, a88, a81, a86, a8f, a82, a83, a84, a87, a8b, a8a, a89, a8e)
Digits(9) = Array(a90, a95, a9c, a9d, a98, a91, a96, a9f, a92, a93, a94, a97, a9b, a9a, a99, a9e)
Digits(10) = Array(aa0, aa5, aac, aad, aa8, aa1, aa6, aaf, aa2, aa3, aa4, aa7, aab, aaa, aa9, aae)
Digits(11) = Array(ab0, ab5, abc, abd, ab8, ab1, ab6, abf, ab2, ab3, ab4, ab7, abb, aba, ab9, abe)
Digits(12) = Array(ac0, ac5, acc, acd, ac8, ac1, ac6, acf, ac2, ac3, ac4, ac7, acb, aca, ac9, ace)
Digits(13) = Array(ad0, ad5, adc, add, ad8, ad1, ad6, adf, ad2, ad3, ad4, ad7, adb, ada, ad9, ade)
Digits(14) = Array(ae0, ae5, aec, aed, ae8, ae1, ae6, aef, ae2, ae3, ae4, ae7, aeb, aea, ae9, aee)
Digits(15) = Array(af0, af5, afc, afd, af8, af1, af6, aff, af2, af3, af4, af7, afb, afa, af9, afe)
Digits(16) = Array(b00, b05, b0c, b0d, b08, b01, b06, b0f, b02, b03, b04, b07, b0b, b0a, b09, b0e)
Digits(17) = Array(b10, b15, b1c, b1d, b18, b11, b16, b1f, b12, b13, b14, b17, b1b, b1a, b19, b1e)
Digits(18) = Array(b20, b25, b2c, b2d, b28, b21, b26, b2f, b22, b23, b24, b27, b2b, b2a, b29, b2e)
Digits(19) = Array(b30, b35, b3c, b3d, b38, b31, b36, b3f, b32, b33, b34, b37, b3b, b3a, b39, b3e)
Digits(20) = Array(b40, b45, b4c, b4d, b48, b41, b46, b4f, b42, b43, b44, b47, b4b, b4a, b49, b4e)
Digits(21) = Array(b50, b55, b5c, b5d, b58, b51, b56, b5f, b52, b53, b54, b57, b5b, b5a, b59, b5e)
Digits(22) = Array(b60, b65, b6c, b6d, b68, b61, b66, b6f, b62, b63, b64, b67, b6b, b6a, b69, b6e)
Digits(23) = Array(b70, b75, b7c, b7d, b78, b71, b76, b7f, b72, b73, b74, b77, b7b, b7a, b79, b7e)
Digits(24) = Array(b80, b85, b8c, b8d, b88, b81, b86, b8f, b82, b83, b84, b87, b8b, b8a, b89, b8e)
Digits(25) = Array(b90, b95, b9c, b9d, b98, b91, b96, b9f, b92, b93, b94, b97, b9b, b9a, b99, b9e)
Digits(26) = Array(ba0, ba5, bac, bad, ba8, ba1, ba6, baf, ba2, ba3, ba4, ba7, bab, baa, ba9, bae)
Digits(27) = Array(bb0, bb5, bbc, bbd, bb8, bb1, bb6, bbf, bb2, bb3, bb4, bb7, bbb, bba, bb9, bbe)
Digits(28) = Array(bc0, bc5, bcc, bcd, bc8, bc1, bc6, bcf, bc2, bc3, bc4, bc7, bcb, bca, bc9, bce)
Digits(29) = Array(bd0, bd5, bdc, bdd, bd8, bd1, bd6, bdf, bd2, bd3, bd4, bd7, bdb, bda, bd9, bde)
Digits(30) = Array(be0, be5, bec, bed, be8, be1, be6, bef, be2, be3, be4, be7, beb, bea, be9, bee)
Digits(31) = Array(bf0, bf5, bfc, bfd, bf8, bf1, bf6, bff, bf2, bf3, bf4, bf7, bfb, bfa, bf9, bfe)
Digits(32) = Array(c00, c05, c0c, c0d, c08, c01, c06, c0f, c02, c03, c04, c07, c0b, c0a, c09, c0e)
Digits(33) = Array(c10, c15, c1c, c1d, c18, c11, c16, c1f, c12, c13, c14, c17, c1b, c1a, c19, c1e)
Digits(34) = Array(c20, c25, c2c, c2d, c28, c21, c26, c2f, c22, c23, c24, c27, c2b, c2a, c29, c2e)
Digits(35) = Array(c30, c35, c3c, c3d, c38, c31, c36, c3f, c32, c33, c34, c37, c3b, c3a, c39, c3e)
Digits(36) = Array(c40, c45, c4c, c4d, c48, c41, c46, c4f, c42, c43, c44, c47, c4b, c4a, c49, c4e)
Digits(37) = Array(c50, c55, c5c, c5d, c58, c51, c56, c5f, c52, c53, c54, c57, c5b, c5a, c59, c5e)
Digits(38) = Array(c60, c65, c6c, c6d, c68, c61, c66, c6f, c62, c63, c64, c67, c6b, c6a, c69, c6e)
Digits(39) = Array(c70, c75, c7c, c7d, c78, c71, c76, c7f, c72, c73, c74, c77, c7b, c7a, c79, c7e)

Sub DisplayTimer_Timer
    Dim ChgLED, ii, jj, num, chg, stat, obj, b, x
    ChgLED=Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
       For ii=0 To UBound(chgLED)
          num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)
      if (num < 40) then
              For Each obj In Digits(num)
                   If chg And 1 Then obj.State=stat And 1
                   chg=chg\2 : stat=stat\2
                  Next
      Else
             end if
        Next
    End If
 End Sub

'VR and FSS Segment Display - Thanks to Rajo Joey
'*********************************************************************

Dim DigitsF(40)

DigitsF(0) = Array(fa00, fa05, fa0c, fa0d, fa08, fa01, fa06, fa0f, fa02, fa03, fa04, fa07, fa0b, fa0a, fa09, fa0e)
DigitsF(1) = Array(fa10, fa15, fa1c, fa1d, fa18, fa11, fa16, fa1f, fa12, fa13, fa14, fa17, fa1b, fa1a, fa19, fa1e)
DigitsF(2) = Array(fa20, fa25, fa2c, fa2d, fa28, fa21, fa26, fa2f, fa22, fa23, fa24, fa27, fa2b, fa2a, fa29, fa2e)
DigitsF(3) = Array(fa30, fa35, fa3c, fa3d, fa38, fa31, fa36, fa3f, fa32, fa33, fa34, fa37, fa3b, fa3a, fa39, fa3e)
DigitsF(4) = Array(fa40, fa45, fa4c, fa4d, fa48, fa41, fa46, fa4f, fa42, fa43, fa44, fa47, fa4b, fa4a, fa49, fa4e)
DigitsF(5) = Array(fa50, fa55, fa5c, fa5d, fa58, fa51, fa56, fa5f, fa52, fa53, fa54, fa57, fa5b, fa5a, fa59, fa5e)
DigitsF(6) = Array(fa60, fa65, fa6c, fa6d, fa68, fa61, fa66, fa6f, fa62, fa63, fa64, fa67, fa6b, fa6a, fa69, fa6e)
DigitsF(7) = Array(fa70, fa75, fa7c, fa7d, fa78, fa71, fa76, fa7f, fa72, fa73, fa74, fa77, fa7b, fa7a, fa79, fa7e)
DigitsF(8) = Array(fa80, fa85, fa8c, fa8d, fa88, fa81, fa86, fa8f, fa82, fa83, fa84, fa87, fa8b, fa8a, fa89, fa8e)
DigitsF(9) = Array(fa90, fa95, fa9c, fa9d, fa98, fa91, fa96, fa9f, fa92, fa93, fa94, fa97, fa9b, fa9a, fa99, fa9e)
DigitsF(10) = Array(faa0, faa5, faac, faad, faa8, faa1, faa6, faaf, faa2, faa3, faa4, faa7, faab, faaa, faa9, faae)
DigitsF(11) = Array(fab0, fab5, fabc, fabd, fab8, fab1, fab6, fabf, fab2, fab3, fab4, fab7, fabb, faba, fab9, fabe)
DigitsF(12) = Array(fac0, fac5, facc, facd, fac8, fac1, fac6, facf, fac2, fac3, fac4, fac7, facb, faca, fac9, face)
DigitsF(13) = Array(fad0, fad5, fadc, fadd, fad8, fad1, fad6, fadf, fad2, fad3, fad4, fad7, fadb, fada, fad9, fade)
DigitsF(14) = Array(fae0, fae5, faec, faed, fae8, fae1, fae6, faef, fae2, fae3, fae4, fae7, faeb, faea, fae9, faee)
DigitsF(15) = Array(faf0, faf5, fafc, fafd, faf8, faf1, faf6, faff, faf2, faf3, faf4, faf7, fafb, fafa, faf9, fafe)

DigitsF(16) = Array(fb00, fb05, fb0c, fb0d, fb08, fb01, fb06, fb0f, fb02, fb03, fb04, fb07, fb0b, fb0a, fb09, fb0e)
DigitsF(17) = Array(fb10, fb15, fb1c, fb1d, fb18, fb11, fb16, fb1f, fb12, fb13, fb14, fb17, fb1b, fb1a, fb19, fb1e)
DigitsF(18) = Array(fb20, fb25, fb2c, fb2d, fb28, fb21, fb26, fb2f, fb22, fb23, fb24, fb27, fb2b, fb2a, fb29, fb2e)
DigitsF(19) = Array(fb30, fb35, fb3c, fb3d, fb38, fb31, fb36, fb3f, fb32, fb33, fb34, fb37, fb3b, fb3a, fb39, fb3e)
DigitsF(20) = Array(fb40, fb45, fb4c, fb4d, fb48, fb41, fb46, fb4f, fb42, fb43, fb44, fb47, fb4b, fb4a, fb49, fb4e)
DigitsF(21) = Array(fb50, fb55, fb5c, fb5d, fb58, fb51, fb56, fb5f, fb52, fb53, fb54, fb57, fb5b, fb5a, fb59, fb5e)
DigitsF(22) = Array(fb60, fb65, fb6c, fb6d, fb68, fb61, fb66, fb6f, fb62, fb63, fb64, fb67, fb6b, fb6a, fb69, fb6e)
DigitsF(23) = Array(fb70, fb75, fb7c, fb7d, fb78, fb71, fb76, fb7f, fb72, fb73, fb74, fb77, fb7b, fb7a, fb79, fb7e)
DigitsF(24) = Array(fb80, fb85, fb8c, fb8d, fb88, fb81, fb86, fb8f, fb82, fb83, fb84, fb87, fb8b, fb8a, fb89, fb8e)
DigitsF(25) = Array(fb90, fb95, fb9c, fb9d, fb98, fb91, fb96, fb9f, fb92, fb93, fb94, fb97, fb9b, fb9a, fb99, fb9e)
DigitsF(26) = Array(fba0, fba5, fbac, fbad, fba8, fba1, fba6, fbaf, fba2, fba3, fba4, fba7, fbab, fbaa, fba9, fbae)
DigitsF(27) = Array(fbb0, fbb5, fbbc, fbbd, fbb8, fbb1, fbb6, fbbf, fbb2, fbb3, fbb4, fbb7, fbbb, fbba, fbb9, fbbe)
DigitsF(28) = Array(fbc0, fbc5, fbcc, fbcd, fbc8, fbc1, fbc6, fbcf, fbc2, fbc3, fbc4, fbc7, fbcb, fbca, fbc9, fbce)
DigitsF(29) = Array(fbd0, fbd5, fbdc, fbdd, fbd8, fbd1, fbd6, fbdf, fbd2, fbd3, fbd4, fbd7, fbdb, fbda, fbd9, fbde)
DigitsF(30) = Array(fbe0, fbe5, fbec, fbed, fbe8, fbe1, fbe6, fbef, fbe2, fbe3, fbe4, fbe7, fbeb, fbea, fbe9, fbee)
DigitsF(31) = Array(fbf0, fbf5, fbfc, fbfd, fbf8, fbf1, fbf6, fbff, fbf2, fbf3, fbf4, fbf7, fbfb, fbfa, fbf9, fbfe)

DigitsF(32) = Array(fc00, fc05, fc0c, fc0d, fc08, fc01, fc06, fc0f, fc02, fc03, fc04, fc07, fc0b, fc0a, fc09, fc0e)
DigitsF(33) = Array(fc10, fc15, fc1c, fc1d, fc18, fc11, fc16, fc1f, fc12, fc13, fc14, fc17, fc1b, fc1a, fc19, fc1e)
DigitsF(34) = Array(fc20, fc25, fc2c, fc2d, fc28, fc21, fc26, fc2f, fc22, fc23, fc24, fc27, fc2b, fc2a, fc29, fc2e)
DigitsF(35) = Array(fc30, fc35, fc3c, fc3d, fc38, fc31, fc36, fc3f, fc32, fc33, fc34, fc37, fc3b, fc3a, fc39, fc3e)
DigitsF(36) = Array(fc40, fc45, fc4c, fc4d, fc48, fc41, fc46, fc4f, fc42, fc43, fc44, fc47, fc4b, fc4a, fc49, fc4e)
DigitsF(37) = Array(fc50, fc55, fc5c, fc5d, fc58, fc51, fc56, fc5f, fc52, fc53, fc54, fc57, fc5b, fc5a, fc59, fc5e)
DigitsF(38) = Array(fc60, fc65, fc6c, fc6d, fc68, fc61, fc66, fc6f, fc62, fc63, fc64, fc67, fc6b, fc6a, fc69, fc6e)
DigitsF(39) = Array(fc70, fc75, fc7c, fc7d, fc78, fc71, fc76, fc7f, fc72, fc73, fc74, fc77, fc7b, fc7a, fc79, fc7e)

Sub UpdateLedsF_Timer
    Dim ChgLED, ii, num, chg, stat, obj
    ChgLED = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
        For ii = 0 To UBound(chgLED)
            num = chgLED(ii, 0):chg = chgLED(ii, 1):stat = chgLED(ii, 2)
            For Each obj In DigitsF(num)
                If chg And 1 Then obj.Visible = stat And 1
                chg = chg \ 2:stat = stat \ 2
            Next
        Next
    End If
End Sub

'*************************************************************************************************************************

Sub TextBox002_Timer()
  TextBox002.Text = spindisk.RotZ / 30
  TextBox004.Text = kickernum
  TextBox005.Text = colorandnum
End Sub

' ***************** CODE BELOW IS FOR THE VR CLOCK.  This was originally taken from Rascal VP9 clock table *******************************
'*****************************************************************************************************************************************
Dim CurrentMinute ' for VR clock

Sub ClockTimer_Timer()

    'ClockHands Below *********************************************************
  Pminutes.RotAndTra2 = (Minute(Now())+(Second(Now())/100))*6
  Phours.RotAndTra2 = Hour(Now())*30+(Minute(Now())/2)
      Pseconds.RotAndTra2 = (Second(Now()))*6
  CurrentMinute=Minute(Now())
End Sub



'**************************************************************
Dim TL,TL1,TL2,Stopctr:TL=0:TL1=34:TL2=68:Stopctr=0
Sub toppertimer_Timer()
  If TL > 0 then topperlights(TL-1).Visible = 0
  If TL1 > 0 then topperlights(TL1-1).Visible = 0
  If TL2 > 0 then topperlights(TL2-1).Visible = 0
  If TL=101 then TL=0
  If TL1=101 then TL1=0
  If TL2=101 then TL2=0
  topperlights(TL).Visible = 1:topperlights(TL1).Visible = 1:topperlights(TL2).Visible = 1
  TL=TL+1:TL1=TL1+1:TL2=TL2+1:Stopctr=Stopctr+1
  If Stopctr=300 Then
    Dim offloop
    toppertimer.Enabled = False:w009.TimerEnabled = False:Stopctr=0
    For Each offloop in topperlights
      offloop.Visible = 0
    Next
    For Each offloop in topperwinner
      offloop.Visible = 0
    Next
  End If
End Sub

Dim TW
Sub w009_Timer()
  For Each TW in topperwinner
    TW.Visible = not TW.Visible
  Next
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


'********************************************************************
'      JP's VP10 Rolling Sounds (+rothbauerw's Dropping Sounds)
'********************************************************************

Const tnob = 5

'**********************
' Ball Collision Sound
'**********************


'*****************************************
' ninuzzu's FLIPPER SHADOWS v2
'*****************************************

Sub LeftFlipperShadow_Timer()
    FlipperLSh.RotZ = LeftFlipper.CurrentAngle
End Sub

Sub RightFlipperShadow_Timer()
    FlipperRSh.RotZ = RightFlipper.CurrentAngle
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
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/21)) + 6
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/21)) - 6
        End If
        ballShadow(b).Y = BOT(b).Y + 4
        If BOT(b).Z > 120 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
End Sub


'**************************************************
'        Flipper Collision Subs
'NOTE: COpy and overwrite collision sound from original collision subs over
'RandomSoundFlipper()' below
'**************************************************'

Sub LeftFlipper_Collide(parm)
  CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
  LF.ReProcessBalls ActiveBall
  LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
  CheckLiveCatch Activeball, RightFlipper, RFCount, parm
  RF.ReProcessBalls ActiveBall
  RightFlipperCollide parm
End Sub



'******************************************************
'   ZDMP:  RUBBER  DAMPENERS
'******************************************************
' These are data mined bounce curves,
' dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
' Requires tracking ballspeed to calculate COR

' Put all the Post and Pin objects in dPosts collection. Make sure dPosts fires hit events.
Sub dPosts_Hit(idx)
  RubbersD.dampen ActiveBall
  TargetBouncer ActiveBall, 1
End Sub

' This collection contains the bottom sling posts. They are not in the dPosts collection so that the TargetBouncer is not applied to them, but they should still have dampening applied
' If you experience airballs with posts or targets, consider adding them to this collection
Sub NoTargetBouncer_Hit
    RubbersD.dampen ActiveBall
End Sub

' Put all the Sleeve objects in dSleeves collection. Make sure dSleeves fires hit events.
Sub dSleeves_Hit(idx)
  SleevesD.Dampen ActiveBall
  TargetBouncer ActiveBall, 0.7
End Sub

Dim RubbersD        'frubber
Set RubbersD = New Dampener
RubbersD.name = "Rubbers"
RubbersD.debugOn = False  'shows info in textbox "TBPout"
RubbersD.Print = False    'debug, reports In debugger (In vel, out cor); cor bounce curve (linear)

'for best results, try to match in-game velocity as closely as possible to the desired curve
'   RubbersD.addpoint 0, 0, 0.935   'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 0, 0, 1.1    'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 1, 3.77, 0.97
RubbersD.addpoint 2, 5.76, 0.967  'dont take this as gospel. if you can data mine rubber elasticitiy, please help!
RubbersD.addpoint 3, 15.84, 0.874
RubbersD.addpoint 4, 56, 0.64    'there's clamping so interpolate up to 56 at least

Dim SleevesD  'this is just rubber but cut down to 85%...
Set SleevesD = New Dampener
SleevesD.name = "Sleeves"
SleevesD.debugOn = False  'shows info in textbox "TBPout"
SleevesD.Print = False    'debug, reports In debugger (In vel, out cor)
SleevesD.CopyCoef RubbersD, 0.85

'######################### Add new FlippersD Profile
'######################### Adjust these values to increase or lessen the elasticity

Dim FlippersD
Set FlippersD = New Dampener
FlippersD.name = "Flippers"
FlippersD.debugOn = False
FlippersD.Print = False
FlippersD.addpoint 0, 0, 1.1
FlippersD.addpoint 1, 3.77, 0.99
FlippersD.addpoint 2, 6, 0.99

Class Dampener
  Public Print, debugOn   'tbpOut.text
  Public name, Threshold  'Minimum threshold. Useful for Flippers, which don't have a hit threshold.
  Public ModIn, ModOut
  Private Sub Class_Initialize
    ReDim ModIn(0)
    ReDim Modout(0)
  End Sub

  Public Sub AddPoint(aIdx, aX, aY)
    ShuffleArrays ModIn, ModOut, 1
    ModIn(aIDX) = aX
    ModOut(aIDX) = aY
    ShuffleArrays ModIn, ModOut, 0
    If GameTime > 100 Then Report
  End Sub

  Public Sub Dampen(aBall)
    If threshold Then
      If BallSpeed(aBall) < threshold Then Exit Sub
    End If
    Dim RealCOR, DesiredCOR, str, coef
    DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
    RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id) + 0.0001)
    coef = desiredcor / realcor
    If debugOn Then str = name & " In vel:" & Round(cor.ballvel(aBall.id),2 ) & vbNewLine & "desired cor: " & Round(desiredcor,4) & vbNewLine & _
    "actual cor: " & Round(realCOR,4) & vbNewLine & "ballspeed coef: " & Round(coef, 3) & vbNewLine
    If Print Then Debug.print Round(cor.ballvel(aBall.id),2) & ", " & Round(desiredcor,3)

    aBall.velx = aBall.velx * coef
    aBall.vely = aBall.vely * coef
    aBall.velz = aBall.velz * coef
    If debugOn Then TBPout.text = str
  End Sub

  Public Sub Dampenf(aBall, parm) 'Rubberizer is handle here
    Dim RealCOR, DesiredCOR, str, coef
    DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
    RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id) + 0.0001)
    coef = desiredcor / realcor
    If Abs(aball.velx) < 2 And aball.vely < 0 And aball.vely >  - 3.75 Then
      aBall.velx = aBall.velx * coef
      aBall.vely = aBall.vely * coef
    End If
  End Sub

  Public Sub CopyCoef(aObj, aCoef) 'alternative addpoints, copy with coef
    Dim x
    For x = 0 To UBound(aObj.ModIn)
      addpoint x, aObj.ModIn(x), aObj.ModOut(x) * aCoef
    Next
  End Sub

  Public Sub Report() 'debug, reports all coords in tbPL.text
    If Not debugOn Then Exit Sub
    Dim a1, a2
    a1 = ModIn
    a2 = ModOut
    Dim str, x
    For x = 0 To UBound(a1)
      str = str & x & ": " & Round(a1(x),4) & ", " & Round(a2(x),4) & vbNewLine
    Next
    TBPout.text = str
  End Sub
End Class


'******************************************************
'   ZBOU: VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************

Const TargetBouncerEnabled = 1    '0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.9   'Level of bounces. Recommmended value of 0.7-1.0

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


'******************************************************
'   ZFLE:  FLEEP MECHANICAL SOUNDS
'******************************************************

' This part in the script is an entire block that is dedicated to the physics sound system.
' Various scripts and sounds that may be pretty generic and could suit other WPC systems, but the most are tailored specifically for the TOM table

' Many of the sounds in this package can be added by creating collections and adding the appropriate objects to those collections.
' Create the following new collections:
'  Metals (all metal objects, metal walls, metal posts, metal wire guides)
'  Apron (the apron walls and plunger wall)
'  Walls (all wood or plastic walls)
'  Rollovers (wire rollover triggers, star triggers, or button triggers)
'  Targets (standup or drop targets, these are hit sounds only ... you will want to add separate dropping sounds for drop targets)
'  Gates (plate gates)
'  GatesWire (wire gates)
'  Rubbers (all rubbers including posts, sleeves, pegs, and bands)
' When creating the collections, make sure "Fire events for this collection" is checked.
' You'll also need to make sure "Has Hit Event" is checked for each object placed in these collections (not necessary for gates and triggers).
' Once the collections and objects are added, the save, close, and restart VPX.
'
' Many places in the script need to be modified to include the correct sound effect subroutine calls. The tutorial videos linked below demonstrate
' how to make these updates. But in summary the following needs to be updated:
' - Nudging, plunger, coin-in, start button sounds will be added to the keydown and keyup subs.
' - Flipper sounds in the flipper solenoid subs. Flipper collision sounds in the flipper collide subs.
' - Bumpers, slingshots, drain, ball release, knocker, spinner, and saucers in their respective subs
' - Ball rolling sounds sub
'
' Tutorial videos by Apophis
' Audio : Adding Fleep Part 1         https://youtu.be/rG35JVHxtx4?si=zdN9W4cZWEyXbOz_
' Audio : Adding Fleep Part 2         https://youtu.be/dk110pWMxGo?si=2iGMImXXZ0SFKVCh
' Audio : Adding Fleep Part 3         https://youtu.be/ESXWGJZY_EI?si=6D20E2nUM-xAw7xy


'///////////////////////////////  SOUNDS PARAMETERS  //////////////////////////////
Dim GlobalSoundLevel, CoinSoundLevel, PlungerReleaseSoundLevel, PlungerPullSoundLevel, NudgeLeftSoundLevel
Dim NudgeRightSoundLevel, NudgeCenterSoundLevel, StartButtonSoundLevel, RollingSoundFactor

CoinSoundLevel = 1            'volume level; range [0, 1]
NudgeLeftSoundLevel = 1        'volume level; range [0, 1]
NudgeRightSoundLevel = 1        'volume level; range [0, 1]
NudgeCenterSoundLevel = 1        'volume level; range [0, 1]
StartButtonSoundLevel = 0.1      'volume level; range [0, 1]
PlungerReleaseSoundLevel = 0.8 '1 wjr   'volume level; range [0, 1]
PlungerPullSoundLevel = 1        'volume level; range [0, 1]
RollingSoundFactor = 1.1 / 5

'///////////////////////-----Solenoids, Kickers and Flash Relays-----///////////////////////
Dim FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel, FlipperUpAttackLeftSoundLevel, FlipperUpAttackRightSoundLevel
Dim FlipperUpSoundLevel, FlipperDownSoundLevel, FlipperLeftHitParm, FlipperRightHitParm
Dim SlingshotSoundLevel, BumperSoundFactor, KnockerSoundLevel

FlipperUpAttackMinimumSoundLevel = 0.010    'volume level; range [0, 1]
FlipperUpAttackMaximumSoundLevel = 0.635    'volume level; range [0, 1]
FlipperUpSoundLevel = 1.0            'volume level; range [0, 1]
FlipperDownSoundLevel = 0.45          'volume level; range [0, 1]
FlipperLeftHitParm = FlipperUpSoundLevel    'sound helper; not configurable
FlipperRightHitParm = FlipperUpSoundLevel    'sound helper; not configurable
SlingshotSoundLevel = 0.95            'volume level; range [0, 1]
BumperSoundFactor = 4.25            'volume multiplier; must not be zero
KnockerSoundLevel = 1              'volume level; range [0, 1]

'///////////////////////-----Ball Drops, Bumps and Collisions-----///////////////////////
Dim RubberStrongSoundFactor, RubberWeakSoundFactor, RubberFlipperSoundFactor,BallWithBallCollisionSoundFactor
Dim BallBouncePlayfieldSoftFactor, BallBouncePlayfieldHardFactor, PlasticRampDropToPlayfieldSoundLevel, WireRampDropToPlayfieldSoundLevel, DelayedBallDropOnPlayfieldSoundLevel
Dim WallImpactSoundFactor, MetalImpactSoundFactor, SubwaySoundLevel, SubwayEntrySoundLevel, ScoopEntrySoundLevel
Dim SaucerLockSoundLevel, SaucerKickSoundLevel

BallWithBallCollisionSoundFactor = 3.2      'volume multiplier; must not be zero
RubberStrongSoundFactor = 0.055 / 5      'volume multiplier; must not be zero
RubberWeakSoundFactor = 0.075 / 5        'volume multiplier; must not be zero
RubberFlipperSoundFactor = 0.075 / 5      'volume multiplier; must not be zero
BallBouncePlayfieldSoftFactor = 0.025      'volume multiplier; must not be zero
BallBouncePlayfieldHardFactor = 0.025      'volume multiplier; must not be zero
DelayedBallDropOnPlayfieldSoundLevel = 0.8    'volume level; range [0, 1]
WallImpactSoundFactor = 0.075          'volume multiplier; must not be zero
MetalImpactSoundFactor = 0.075 / 3
SaucerLockSoundLevel = 0.8
SaucerKickSoundLevel = 0.8

'///////////////////////-----Gates, Spinners, Rollovers and Targets-----///////////////////////

Dim GateSoundLevel, TargetSoundFactor, SpinnerSoundLevel, RolloverSoundLevel, DTSoundLevel

GateSoundLevel = 0.5 / 5      'volume level; range [0, 1]
TargetSoundFactor = 0.0025 * 10  'volume multiplier; must not be zero
DTSoundLevel = 0.25        'volume multiplier; must not be zero
RolloverSoundLevel = 0.25      'volume level; range [0, 1]
SpinnerSoundLevel = 0.5      'volume level; range [0, 1]

'///////////////////////-----Ball Release, Guides and Drain-----///////////////////////
Dim DrainSoundLevel, BallReleaseSoundLevel, BottomArchBallGuideSoundFactor, FlipperBallGuideSoundFactor

DrainSoundLevel = 0.8          'volume level; range [0, 1]
BallReleaseSoundLevel = 1        'volume level; range [0, 1]
BottomArchBallGuideSoundFactor = 0.2  'volume multiplier; must not be zero
FlipperBallGuideSoundFactor = 0.015  'volume multiplier; must not be zero

'///////////////////////-----Loops and Lanes-----///////////////////////
Dim ArchSoundFactor
ArchSoundFactor = 0.025 / 5      'volume multiplier; must not be zero

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
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelExistingStatic(playsoundparams, aVol, tableobj)
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticLoop(playsoundparams, aVol, tableobj)
  PlaySound playsoundparams, - 1, min(aVol,1) * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticRandomPitch(playsoundparams, aVol, randomPitch, tableobj)
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelActiveBall(playsoundparams, aVol)
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(ActiveBall), 0, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLevelExistingActiveBall(playsoundparams, aVol)
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(ActiveBall), 0, 0, 1, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLeveTimerActiveBall(playsoundparams, aVol, ballvariable)
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(ballvariable), 0, 0, 0, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelTimerExistingActiveBall(playsoundparams, aVol, ballvariable)
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(ballvariable), 0, 0, 1, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelRoll(playsoundparams, aVol, pitch)
  PlaySound playsoundparams, - 1, min(aVol,1) * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

' Previous Positional Sound Subs

Sub PlaySoundAt(soundname, tableobj)
  PlaySound soundname, 1, 1 * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtVol(soundname, tableobj, aVol)
  PlaySound soundname, 1, min(aVol,1) * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname)
  PlaySoundAt soundname, ActiveBall
End Sub

Sub PlaySoundAtBallVol (Soundname, aVol)
  PlaySound soundname, 1,min(aVol,1) * VolumeDial, AudioPan(ActiveBall), 0,0,0, 1, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtBallVolM (Soundname, aVol)
  PlaySound soundname, 1,min(aVol,1) * VolumeDial, AudioPan(ActiveBall), 0,0,0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtVolLoops(sound, tableobj, Vol, Loops)
  PlaySound sound, Loops, Vol * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'******************************************************
'  Fleep  Supporting Ball & Sound Functions
'******************************************************

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / tableheight - 1

  If tmp > 7000 Then
    tmp = 7000
  ElseIf tmp <  - 7000 Then
    tmp =  - 7000
  End If

  If tmp > 0 Then
    AudioFade = CSng(tmp ^ 10)
  Else
    AudioFade = CSng( - (( - tmp) ^ 10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / tablewidth - 1

  If tmp > 7000 Then
    tmp = 7000
  ElseIf tmp <  - 7000 Then
    tmp =  - 7000
  End If

  If tmp > 0 Then
    AudioPan = CSng(tmp ^ 10)
  Else
    AudioPan = CSng( - (( - tmp) ^ 10) )
  End If
End Function

Function Vol(ball) ' Calculates the volume of the sound based on the ball speed
  Vol = CSng(BallVel(ball) ^ 2)
End Function

Function Volz(ball) ' Calculates the volume of the sound based on the ball speed
  Volz = CSng((ball.velz) ^ 2)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
  Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
  BallVel = Int(Sqr((ball.VelX ^ 2) + (ball.VelY ^ 2) ) )
End Function

Function VolPlayfieldRoll(ball) ' Calculates the roll volume of the sound based on the ball speed
  VolPlayfieldRoll = RollingSoundFactor * 0.0005 * CSng(BallVel(ball) ^ 3)
End Function

Function PitchPlayfieldRoll(ball) ' Calculates the roll pitch of the sound based on the ball speed
  PitchPlayfieldRoll = BallVel(ball) ^ 2 * 15
End Function

Function RndInt(min, max) ' Sets a random number integer between min and max
  RndInt = Int(Rnd() * (max - min + 1) + min)
End Function

Function RndNum(min, max) ' Sets a random number between min and max
  RndNum = Rnd() * (max - min) + min
End Function

'/////////////////////////////  GENERAL SOUND SUBROUTINES  ////////////////////////////

Sub SoundStartButton()
  PlaySound ("Start_Button"), 0, StartButtonSoundLevel, 0, 0.25
End Sub

Sub SoundNudgeLeft()
  PlaySound ("Nudge_" & Int(Rnd * 2) + 1), 0, NudgeLeftSoundLevel * VolumeDial, - 0.1, 0.25
End Sub

Sub SoundNudgeRight()
  PlaySound ("Nudge_" & Int(Rnd * 2) + 1), 0, NudgeRightSoundLevel * VolumeDial, 0.1, 0.25
End Sub

Sub SoundNudgeCenter()
  PlaySound ("Nudge_" & Int(Rnd * 2) + 1), 0, NudgeCenterSoundLevel * VolumeDial, 0, 0.25
End Sub

Sub SoundPlungerPull()
  PlaySoundAtLevelStatic ("Plunger_Pull_1"), PlungerPullSoundLevel, Plunger
End Sub

Sub SoundPlungerReleaseBall()
  PlaySoundAtLevelStatic ("Plunger_Release_Ball"), PlungerReleaseSoundLevel, Plunger
End Sub

Sub SoundPlungerReleaseBall2()
  PlaySoundAtLevelStatic ("Plunger_Release_Ball"), PlungerReleaseSoundLevel, kickback
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
  PlaySoundAtLevelStatic ("Drain_" & Int(Rnd * 11) + 1), DrainSoundLevel, drainswitch
End Sub

'/////////////////////////////  TROUGH BALL RELEASE SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundBallRelease(drainswitch)
  PlaySoundAtLevelStatic SoundFX("BallRelease" & Int(Rnd * 7) + 1,DOFContactors), BallReleaseSoundLevel, drainswitch
End Sub

'/////////////////////////////  SLINGSHOT SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundSlingshotLeft(sling)
  PlaySoundAtLevelStatic SoundFX("Sling_L" & Int(Rnd * 10) + 1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

Sub RandomSoundSlingshotRight(sling)
  PlaySoundAtLevelStatic SoundFX("Sling_R" & Int(Rnd * 8) + 1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

'/////////////////////////////  BUMPER SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundBumperTop(Bump)
  PlaySoundAtLevelStatic SoundFX("Bumpers_Top_" & Int(Rnd * 5) + 1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperMiddle(Bump)
  PlaySoundAtLevelStatic SoundFX("Bumpers_Middle_" & Int(Rnd * 5) + 1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperBottom(Bump)
  PlaySoundAtLevelStatic SoundFX("Bumpers_Bottom_" & Int(Rnd * 5) + 1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperLeft(Bump)
  PlaySoundAtLevelStatic SoundFX("Bumpers_Left_" & Int(Rnd * 5) + 1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

'/////////////////////////////  SPINNER SOUNDS  ////////////////////////////

Sub SoundSpinner(spinnerswitch)
  PlaySoundAtLevelStatic ("Spinner"), SpinnerSoundLevel, spinnerswitch
End Sub

'/////////////////////////////  FLIPPER BATS SOUND SUBROUTINES  ////////////////////////////
'/////////////////////////////  FLIPPER BATS SOLENOID ATTACK SOUND  ////////////////////////////

Sub SoundFlipperUpAttackLeft(flipper)
  FlipperUpAttackLeftSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
  PlaySoundAtLevelStatic SoundFX("Flipper_Attack-L01",DOFFlippers), FlipperUpAttackLeftSoundLevel, flipper
End Sub

Sub SoundFlipperUpAttackRight(flipper)
  FlipperUpAttackRightSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
  PlaySoundAtLevelStatic SoundFX("Flipper_Attack-R01",DOFFlippers), FlipperUpAttackLeftSoundLevel, flipper
End Sub

'/////////////////////////////  FLIPPER BATS SOLENOID CORE SOUND  ////////////////////////////

Sub RandomSoundFlipperUpLeft(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_L0" & Int(Rnd * 9) + 1,DOFFlippers), FlipperLeftHitParm, Flipper
End Sub

Sub RandomSoundFlipperUpRight(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_R0" & Int(Rnd * 9) + 1,DOFFlippers), FlipperRightHitParm, Flipper
End Sub

Sub RandomSoundReflipUpLeft(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_L0" & Int(Rnd * 3) + 1,DOFFlippers), (RndNum(0.8, 1)) * FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundReflipUpRight(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_R0" & Int(Rnd * 3) + 1,DOFFlippers), (RndNum(0.8, 1)) * FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperDownLeft(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_Left_Down_" & Int(Rnd * 7) + 1,DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperDownRight(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_Right_Down_" & Int(Rnd * 8) + 1,DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

'/////////////////////////////  FLIPPER BATS BALL COLLIDE SOUND  ////////////////////////////

Sub LeftFlipperCollide(parm)
  FlipperLeftHitParm = parm / 10
  If FlipperLeftHitParm > 1 Then
    FlipperLeftHitParm = 1
  End If
  FlipperLeftHitParm = FlipperUpSoundLevel * FlipperLeftHitParm
  RandomSoundRubberFlipper(parm)
End Sub

Sub RightFlipperCollide(parm)
  FlipperRightHitParm = parm / 10
  If FlipperRightHitParm > 1 Then
    FlipperRightHitParm = 1
  End If
  FlipperRightHitParm = FlipperUpSoundLevel * FlipperRightHitParm
  RandomSoundRubberFlipper(parm)
End Sub

Sub RandomSoundRubberFlipper(parm)
  PlaySoundAtLevelActiveBall ("Flipper_Rubber_" & Int(Rnd * 7) + 1), parm * RubberFlipperSoundFactor
End Sub

'/////////////////////////////  ROLLOVER SOUNDS  ////////////////////////////

Sub RandomSoundRollover()
  PlaySoundAtLevelActiveBall ("Rollover_" & Int(Rnd * 4) + 1), RolloverSoundLevel
End Sub

Sub Rollovers_Hit(idx)
  RandomSoundRollover
End Sub

'/////////////////////////////  VARIOUS PLAYFIELD SOUND SUBROUTINES  ////////////////////////////
'/////////////////////////////  RUBBERS AND POSTS  ////////////////////////////
'/////////////////////////////  RUBBERS - EVENTS  ////////////////////////////

Sub Rubbers_Hit(idx)
  Dim finalspeed
  finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
  If finalspeed > 5 Then
    RandomSoundRubberStrong 1
  End If
  If finalspeed <= 5 Then
    RandomSoundRubberWeak()
  End If
End Sub

Sub Posts_Hit(idx)
  Dim finalspeed
  finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
  If finalspeed > 5 Then
    RandomSoundRubberStrong 1
  End If
  If finalspeed <= 5 Then
    RandomSoundRubberWeak()
  End If
End Sub

'/////////////////////////////  RUBBERS AND POSTS - STRONG IMPACTS  ////////////////////////////

Sub RandomSoundRubberStrong(voladj)
  Select Case Int(Rnd * 10) + 1
    Case 1
      PlaySoundAtLevelActiveBall ("Rubber_Strong_1"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 2
      PlaySoundAtLevelActiveBall ("Rubber_Strong_2"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 3
      PlaySoundAtLevelActiveBall ("Rubber_Strong_3"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 4
      PlaySoundAtLevelActiveBall ("Rubber_Strong_4"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 5
      PlaySoundAtLevelActiveBall ("Rubber_Strong_5"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 6
      PlaySoundAtLevelActiveBall ("Rubber_Strong_6"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 7
      PlaySoundAtLevelActiveBall ("Rubber_Strong_7"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 8
      PlaySoundAtLevelActiveBall ("Rubber_Strong_8"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 9
      PlaySoundAtLevelActiveBall ("Rubber_Strong_9"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 10
      PlaySoundAtLevelActiveBall ("Rubber_1_Hard"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6 * voladj
  End Select
End Sub

'/////////////////////////////  RUBBERS AND POSTS - WEAK IMPACTS  ////////////////////////////

Sub RandomSoundRubberWeak()
  PlaySoundAtLevelActiveBall ("Rubber_" & Int(Rnd * 9) + 1), Vol(ActiveBall) * RubberWeakSoundFactor
End Sub

'/////////////////////////////  WALL IMPACTS  ////////////////////////////

Sub Walls_Hit(idx)
  RandomSoundWall()
End Sub

Sub RandomSoundWall()
  Dim finalspeed
  finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
  If finalspeed > 16 Then
    Select Case Int(Rnd * 5) + 1
      Case 1
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_1"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_2"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_5"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 4
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_7"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 5
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_9"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End If
  If finalspeed >= 6 And finalspeed <= 16 Then
    Select Case Int(Rnd * 4) + 1
      Case 1
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_3"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 4
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End If
  If finalspeed < 6 Then
    Select Case Int(Rnd * 3) + 1
      Case 1
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End If
End Sub

'/////////////////////////////  METAL TOUCH SOUNDS  ////////////////////////////

Sub RandomSoundMetal()
  PlaySoundAtLevelActiveBall ("Metal_Touch_" & Int(Rnd * 13) + 1), Vol(ActiveBall) * MetalImpactSoundFactor
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
  Dim finalspeed
  finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
  If finalspeed > 16 Then
    PlaySoundAtLevelActiveBall ("Apron_Bounce_" & Int(Rnd * 2) + 1), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
  End If
  If finalspeed >= 6 And finalspeed <= 16 Then
    Select Case Int(Rnd * 2) + 1
      Case 1
        PlaySoundAtLevelActiveBall ("Apron_Bounce_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
      Case 2
        PlaySoundAtLevelActiveBall ("Apron_Bounce_Soft_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
    End Select
  End If
  If finalspeed < 6 Then
    Select Case Int(Rnd * 2) + 1
      Case 1
        PlaySoundAtLevelActiveBall ("Apron_Bounce_Soft_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
      Case 2
        PlaySoundAtLevelActiveBall ("Apron_Medium_3"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
    End Select
  End If
End Sub

'/////////////////////////////  BOTTOM ARCH BALL GUIDE - HARD HITS  ////////////////////////////

Sub RandomSoundBottomArchBallGuideHardHit()
  PlaySoundAtLevelActiveBall ("Apron_Hard_Hit_" & Int(Rnd * 3) + 1), BottomArchBallGuideSoundFactor * 0.25
End Sub

Sub Apron_Hit (idx)
  If Abs(cor.ballvelx(ActiveBall.id) < 4) And cor.ballvely(ActiveBall.id) > 7 Then
    RandomSoundBottomArchBallGuideHardHit()
  Else
    RandomSoundBottomArchBallGuide
  End If
End Sub

'/////////////////////////////  FLIPPER BALL GUIDE  ////////////////////////////

Sub RandomSoundFlipperBallGuide()
  Dim finalspeed
  finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
  If finalspeed > 16 Then
    Select Case Int(Rnd * 2) + 1
      Case 1
        PlaySoundAtLevelActiveBall ("Apron_Hard_1"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
      Case 2
        PlaySoundAtLevelActiveBall ("Apron_Hard_2"),  Vol(ActiveBall) * 0.8 * FlipperBallGuideSoundFactor
    End Select
  End If
  If finalspeed >= 6 And finalspeed <= 16 Then
    PlaySoundAtLevelActiveBall ("Apron_Medium_" & Int(Rnd * 3) + 1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
  End If
  If finalspeed < 6 Then
    PlaySoundAtLevelActiveBall ("Apron_Soft_" & Int(Rnd * 7) + 1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
  End If
End Sub

'/////////////////////////////  TARGET HIT SOUNDS  ////////////////////////////

Sub RandomSoundTargetHitStrong()
  PlaySoundAtLevelActiveBall SoundFX("Target_Hit_" & Int(Rnd * 4) + 5,DOFTargets), Vol(ActiveBall) * 0.45 * TargetSoundFactor
End Sub

Sub RandomSoundTargetHitWeak()
  PlaySoundAtLevelActiveBall SoundFX("Target_Hit_" & Int(Rnd * 4) + 1,DOFTargets), Vol(ActiveBall) * TargetSoundFactor
End Sub

Sub PlayTargetSound()
  Dim finalspeed
  finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
  If finalspeed > 10 Then
    RandomSoundTargetHitStrong()
    RandomSoundBallBouncePlayfieldSoft ActiveBall
  Else
    RandomSoundTargetHitWeak()
  End If
End Sub

Sub Targets_Hit (idx)
  PlayTargetSound
End Sub

'/////////////////////////////  BALL BOUNCE SOUNDS  ////////////////////////////

Sub RandomSoundBallBouncePlayfieldSoft(aBall)
  Select Case Int(Rnd * 9) + 1
    Case 1
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_1"), volz(aBall) * BallBouncePlayfieldSoftFactor, aBall
    Case 2
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_2"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.5, aBall
    Case 3
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_3"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.8, aBall
    Case 4
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_4"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.5, aBall
    Case 5
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_5"), volz(aBall) * BallBouncePlayfieldSoftFactor, aBall
    Case 6
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_1"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
    Case 7
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_2"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
    Case 8
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_5"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
    Case 9
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_7"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.3, aBall
  End Select
End Sub

Sub RandomSoundBallBouncePlayfieldHard(aBall)
  PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_" & Int(Rnd * 7) + 1), volz(aBall) * BallBouncePlayfieldHardFactor, aBall
End Sub

'/////////////////////////////  DELAYED DROP - TO PLAYFIELD - SOUND  ////////////////////////////

Sub RandomSoundDelayedBallDropOnPlayfield(aBall)
  Select Case Int(Rnd * 5) + 1
    Case 1
      PlaySoundAtLevelStatic ("Ball_Drop_Playfield_1_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 2
      PlaySoundAtLevelStatic ("Ball_Drop_Playfield_2_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 3
      PlaySoundAtLevelStatic ("Ball_Drop_Playfield_3_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 4
      PlaySoundAtLevelStatic ("Ball_Drop_Playfield_4_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 5
      PlaySoundAtLevelStatic ("Ball_Drop_Playfield_5_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
  End Select
End Sub

'/////////////////////////////  BALL GATES AND BRACKET GATES SOUNDS  ////////////////////////////

Sub SoundPlayfieldGate()
  PlaySoundAtLevelStatic ("Gate_FastTrigger_" & Int(Rnd * 2) + 1), GateSoundLevel, ActiveBall
End Sub

Sub SoundHeavyGate()
  PlaySoundAtLevelStatic ("Gate_2"), GateSoundLevel, ActiveBall
End Sub

Sub Gates_hit(idx)
  SoundHeavyGate
End Sub

Sub GatesWire_hit(idx)
  SoundPlayfieldGate
End Sub

'/////////////////////////////  LEFT LANE ENTRANCE - SOUNDS  ////////////////////////////

Sub RandomSoundLeftArch()
  PlaySoundAtLevelActiveBall ("Arch_L" & Int(Rnd * 4) + 1), Vol(ActiveBall) * ArchSoundFactor
End Sub

Sub RandomSoundRightArch()
  PlaySoundAtLevelActiveBall ("Arch_R" & Int(Rnd * 4) + 1), Vol(ActiveBall) * ArchSoundFactor
End Sub

Sub Arch1_hit()
  If ActiveBall.velx > 1 Then SoundPlayfieldGate
  StopSound "Arch_L1"
  StopSound "Arch_L2"
  StopSound "Arch_L3"
  StopSound "Arch_L4"
End Sub

Sub Arch1_unhit()
  If ActiveBall.velx <  - 8 Then
    RandomSoundRightArch
  End If
End Sub

Sub Arch2_hit()
  If ActiveBall.velx < 1 Then SoundPlayfieldGate
  StopSound "Arch_R1"
  StopSound "Arch_R2"
  StopSound "Arch_R3"
  StopSound "Arch_R4"
End Sub

Sub Arch2_unhit()
  If ActiveBall.velx > 10 Then
    RandomSoundLeftArch
  End If
End Sub

'/////////////////////////////  SAUCERS (KICKER HOLES)  ////////////////////////////

Sub SoundSaucerLock()
  PlaySoundAtLevelStatic ("Saucer_Enter_" & Int(Rnd * 2) + 1), SaucerLockSoundLevel, ActiveBall
End Sub

Sub SoundSaucerKick(scenario, saucer)
  Select Case scenario
    Case 0
      PlaySoundAtLevelStatic SoundFX("Saucer_Empty", DOFContactors), SaucerKickSoundLevel, saucer
    Case 1
      PlaySoundAtLevelStatic SoundFX("Saucer_Kick", DOFContactors), SaucerKickSoundLevel, saucer
  End Select
End Sub

'/////////////////////////////  BALL COLLISION SOUND  ////////////////////////////

Sub OnBallBallCollision(ball1, ball2, velocity)

  FlipperCradleCollision ball1, ball2, velocity

  Dim snd
  Select Case Int(Rnd * 7) + 1
    Case 1
      snd = "Ball_Collide_1"
    Case 2
      snd = "Ball_Collide_2"
    Case 3
      snd = "Ball_Collide_3"
    Case 4
      snd = "Ball_Collide_4"
    Case 5
      snd = "Ball_Collide_5"
    Case 6
      snd = "Ball_Collide_6"
    Case 7
      snd = "Ball_Collide_7"
  End Select

  PlaySound (snd), 0, CSng(velocity) ^ 2 / 200 * BallWithBallCollisionSoundFactor * VolumeDial, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

'///////////////////////////  DROP TARGET HIT SOUNDS  ///////////////////////////

Sub RandomSoundDropTargetReset(obj)
  PlaySoundAtLevelStatic SoundFX("Drop_Target_Reset_" & Int(Rnd * 6) + 1,DOFContactors), 1, obj
End Sub

Sub SoundDropTargetDrop(obj)
  PlaySoundAtLevelStatic ("Drop_Target_Down_" & Int(Rnd * 6) + 1), 200, obj
End Sub

'/////////////////////////////  GI AND FLASHER   ////////////////////////////

Const RelayFlashSoundLevel = 0.315  'volume level; range [0, 1];
Const RelayGISoundLevel = 1.05    'volume level; range [0, 1];

Sub Sound_GI_Relay(toggle, obj)
  Select Case toggle
    Case 1
      PlaySoundAtLevelStatic ("Relay_GI_On"), 0.025 * RelayGISoundLevel, obj
    Case 0
      PlaySoundAtLevelStatic ("Relay_GI_Off"), 0.025 * RelayGISoundLevel, obj
  End Select
End Sub

Sub Sound_Flash_Relay(toggle, obj)
  Select Case toggle
    Case 1
      PlaySoundAtLevelStatic ("Relay_Flash_On"), 0.025 * RelayFlashSoundLevel, obj
    Case 0
      PlaySoundAtLevelStatic ("Relay_Flash_Off"), 0.025 * RelayFlashSoundLevel, obj
  End Select
End Sub

'/////////////////////////////////////////////////////////////////
'         End Mechanical Sounds
'/////////////////////////////////////////////////////////////////


'******************************************************
'****  END FLEEP MECHANICAL SOUNDS
'******************************************************

'**********************************
'   ZMAT: General Math Functions
'**********************************
' These get used throughout the script.

Dim PI
PI = 4 * Atn(1)

Function dSin(degrees)
  dsin = Sin(degrees * Pi / 180)
End Function

Function dCos(degrees)
  dcos = Cos(degrees * Pi / 180)
End Function

Function Atn2(dy, dx)
  If dx > 0 Then
    Atn2 = Atn(dy / dx)
  ElseIf dx < 0 Then
    If dy = 0 Then
      Atn2 = pi
    Else
      Atn2 = Sgn(dy) * (pi - Atn(Abs(dy / dx)))
    End If
  ElseIf dx = 0 Then
    If dy = 0 Then
      Atn2 = 0
    Else
      Atn2 = Sgn(dy) * pi / 2
    End If
  End If
End Function

Function ArcCos(x)
  If x = 1 Then
    ArcCos = 0/180*PI
  ElseIf x = -1 Then
    ArcCos = 180/180*PI
  Else
    ArcCos = Atn(-x/Sqr(-x * x + 1)) + 2 * Atn(1)
  End If
End Function

Function max(a,b)
  If a > b Then
    max = a
  Else
    max = b
  End If
End Function

Function min(a,b)
  If a > b Then
    min = b
  Else
    min = a
  End If
End Function

' Used for drop targets
Function InRect(px,py,ax,ay,bx,by,cx,cy,dx,dy) 'Determines if a Points (px,py) is inside a 4 point polygon A-D in Clockwise/CCW order
  Dim AB, BC, CD, DA
  AB = (bx * py) - (by * px) - (ax * py) + (ay * px) + (ax * by) - (ay * bx)
  BC = (cx * py) - (cy * px) - (bx * py) + (by * px) + (bx * cy) - (by * cx)
  CD = (dx * py) - (dy * px) - (cx * py) + (cy * px) + (cx * dy) - (cy * dx)
  DA = (ax * py) - (ay * px) - (dx * py) + (dy * px) + (dx * ay) - (dy * ax)

  If (AB <= 0 And BC <= 0 And CD <= 0 And DA <= 0) Or (AB >= 0 And BC >= 0 And CD >= 0 And DA >= 0) Then
    InRect = True
  Else
    InRect = False
  End If
End Function

Function InRotRect(ballx,bally,px,py,angle,ax,ay,bx,by,cx,cy,dx,dy)
  Dim rax,ray,rbx,rby,rcx,rcy,rdx,rdy
  Dim rotxy
  rotxy = RotPoint(ax,ay,angle)
  rax = rotxy(0) + px
  ray = rotxy(1) + py
  rotxy = RotPoint(bx,by,angle)
  rbx = rotxy(0) + px
  rby = rotxy(1) + py
  rotxy = RotPoint(cx,cy,angle)
  rcx = rotxy(0) + px
  rcy = rotxy(1) + py
  rotxy = RotPoint(dx,dy,angle)
  rdx = rotxy(0) + px
  rdy = rotxy(1) + py

  InRotRect = InRect(ballx,bally,rax,ray,rbx,rby,rcx,rcy,rdx,rdy)
End Function

Function RotPoint(x,y,angle)
  Dim rx, ry
  rx = x * dCos(angle) - y * dSin(angle)
  ry = x * dSin(angle) + y * dCos(angle)
  RotPoint = Array(rx,ry)
End Function


'*******************************************
'  ZOPT: User Options
'*******************************************

Dim VolumeDial : VolumeDial = 0.8             ' Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Dim BallRollVolume : BallRollVolume = 0.5     ' Level of ball rolling volume. Value between 0 and 1
Dim RampRollVolume : RampRollVolume = 0.5     ' Level of ramp rolling volume. Value between 0 and 1

' Called when options are tweaked by the player.
' - 0: game has started, good time to load options and adjust accordingly
' - 1: an option has changed
' - 2: options have been reseted
' - 3: player closed the tweak UI, good time to update staticly prerendered parts
' Table1.Option arguments are:
' - option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional arry of literal strings
Dim dspTriggered : dspTriggered = False
Dim LUToption
Sub Table1_OptionEvent(ByVal eventId)
    If eventId = 1 And Not dspTriggered Then dspTriggered = True : DisableStaticPreRendering = True : End If
  LUToption = Table1.Option("LUT Color Scheme", 0, 38, 1, 34, 0, Array("Dark 0","Dark 1", "Dark 2", "Dark 3", "Dark 4", "Dark 5", "Dark 6", "Dark 7", "Dark 8", "Dark 9", "Dark 10", "LUT Warm 0", "LUT Warm 1", "LUT Warm 2", "LUT Warm 3", "LUT Warm 4", "LUT Warm 5", "LUT Warm 6","LUT Warm 7","LUT Warm 8","LUT Warm 9","LUT Warm 10","Fleep Natural Dark 1","Fleep Natural Dark 2","Fleep Warm Dark","Fleep Warm Bright","Fleep Warm Vivid Soft","Fleep Warm Vivid Hard","Skitso Natural and Balanced","Skitso Natural High Contrast","3rdaxis Referenced THX Standard","CalleV Punchy Brightness and Contrast","HauntFreaks Desaturated","Tomate Washed Out","VPW Original 1 to 1","Bassgeige","Blacklight","B&W Comic Book","Skitso"))
    Table1.ColorGradeImage = "LUT" & LUToption
    ' Sound volumes
    VolumeDial = Table1.Option("Mech Volume", 0, 1, 0.01, 0.8, 1)
    BallRollVolume = Table1.Option("Ball Roll Volume", 0, 1, 0.01, 0.9, 1)
    RampRollVolume = Table1.Option("Ramp Roll Volume", 0, 1, 0.01, 0.9, 1)


    If eventId = 3 And dspTriggered Then dspTriggered = False : DisableStaticPreRendering = False : End If
End Sub


'******************************************************
' ZBRL: BALL ROLLING AND DROP SOUNDS
'******************************************************
'
' Be sure to call RollingUpdate in a timer with a 10ms interval see the GameTimer_Timer() sub

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
  Dim b', BOT
  gBOT = GetBalls


  ' stop the sound of deleted balls
  For b = UBound(gBOT) + 1 to tnob - 1
    rolling(b) = False
    StopSound("BallRoll_" & b)
  Next

  ' exit the sub if no balls on the table
  If UBound(gBOT) = -1 Then Exit Sub

  ' play the rolling sound for each ball

  For b = 0 to UBound(gBOT)
    If BallVel(gBOT(b)) > 1 AND gBOT(b).z < 130 AND gBOT(b).z > 80 Then
      rolling(b) = True
      PlaySound ("BallRoll_" & b), -1, VolPlayfieldRoll(gBOT(b)) * BallRollVolume * VolumeDial, AudioPan(gBOT(b)), 0, PitchPlayfieldRoll(gBOT(b)), 1, 0, AudioFade(gBOT(b))

    Else
      If rolling(b) = True Then
        StopSound("BallRoll_" & b)
        rolling(b) = False
      End If
    End If

    ' Ball Drop Sounds
    If gBOT(b).VelZ < -1 and gBOT(b).z < 155 and gBOT(b).z > 127 Then 'height adjust for ball drop sounds
      If DropCount(b) >= 5 Then
        DropCount(b) = 0
        If gBOT(b).velz > -7 Then
          RandomSoundBallBouncePlayfieldSoft gBOT(b)
        Else
          RandomSoundBallBouncePlayfieldHard gBOT(b)
        End If
      End If
    End If
    If DropCount(b) < 5 Then
      DropCount(b) = DropCount(b) + 1
    End If
  Next
End Sub


'******************************************************
'****  END BALL ROLLING AND DROP SOUNDS
'******************************************************

'******************************************************
' ZSSC: SLINGSHOT CORRECTION FUNCTIONS by apophis
'******************************************************
' To add these slingshot corrections:
'  - On the table, add the endpoint primitives that define the two ends of the Slingshot
'  - Initialize the SlingshotCorrection objects in InitSlingCorrection
'  - Call the .VelocityCorrect methods from the respective _Slingshot event sub

Dim LS: Set LS = New SlingshotCorrection
Dim RS: Set RS = New SlingshotCorrection
'Dim TS: Set TS = New SlingshotCorrection

InitSlingCorrection

Sub InitSlingCorrection
  LS.Object = LeftSlingshot
  LS.EndPoint1 = EndPoint1LS
  LS.EndPoint2 = EndPoint2LS

  RS.Object = RightSlingshot
  RS.EndPoint1 = EndPoint1RS
  RS.EndPoint2 = EndPoint2RS

  'TS.Object = TopSlingshot
  'TS.EndPoint1 = EndPoint1TS
  'TS.EndPoint2 = EndPoint2TS

  'Slingshot angle corrections (pt, BallPos in %, Angle in deg)
  ' These values are best guesses. Retune them if needed based on specific table research.
  AddSlingsPt 0, 0.00, - 4
  AddSlingsPt 1, 0.45, - 7
  AddSlingsPt 2, 0.48,  0
  AddSlingsPt 3, 0.52,  0
  AddSlingsPt 4, 0.55,  7
  AddSlingsPt 5, 1.00,  4
End Sub

Sub AddSlingsPt(idx, aX, aY)    'debugger wrapper for adjusting flipper script In-game
  Dim a
  a = Array(LS, RS)
  Dim x
  For Each x In a
    x.addpoint idx, aX, aY
  Next
End Sub

' The following sub are needed, however they exist in the ZMAT maths section of the script. Uncomment below if needed
'Dim PI: PI = 4*Atn(1)
'Function dSin(degrees)
' dsin = sin(degrees * Pi/180)
'End Function
'Function dCos(degrees)
' dcos = cos(degrees * Pi/180)
'End Function
'
'Function RotPoint(x,y,angle)
' dim rx, ry
' rx = x*dCos(angle) - y*dSin(angle)
' ry = x*dSin(angle) + y*dCos(angle)
' RotPoint = Array(rx,ry)
'End Function

Class SlingshotCorrection
  Public DebugOn, Enabled
  Private Slingshot, SlingX1, SlingX2, SlingY1, SlingY2

  Public ModIn, ModOut

  Private Sub Class_Initialize
    ReDim ModIn(0)
    ReDim Modout(0)
    Enabled = True
  End Sub

  Public Property Let Object(aInput)
    Set Slingshot = aInput
  End Property

  Public Property Let EndPoint1(aInput)
    SlingX1 = aInput.x
    SlingY1 = aInput.y
  End Property

  Public Property Let EndPoint2(aInput)
    SlingX2 = aInput.x
    SlingY2 = aInput.y
  End Property

  Public Sub AddPoint(aIdx, aX, aY)
    ShuffleArrays ModIn, ModOut, 1
    ModIn(aIDX) = aX
    ModOut(aIDX) = aY
    ShuffleArrays ModIn, ModOut, 0
    If GameTime > 100 Then Report
  End Sub

  Public Sub Report() 'debug, reports all coords in tbPL.text
    If Not debugOn Then Exit Sub
    Dim a1, a2
    a1 = ModIn
    a2 = ModOut
    Dim str, x
    For x = 0 To UBound(a1)
      str = str & x & ": " & Round(a1(x),4) & ", " & Round(a2(x),4) & vbNewLine
    Next
    TBPout.text = str
  End Sub


  Public Sub VelocityCorrect(aBall)
    Dim BallPos, XL, XR, YL, YR

    'Assign right and left end points
    If SlingX1 < SlingX2 Then
      XL = SlingX1
      YL = SlingY1
      XR = SlingX2
      YR = SlingY2
    Else
      XL = SlingX2
      YL = SlingY2
      XR = SlingX1
      YR = SlingY1
    End If

    'Find BallPos = % on Slingshot
    If Not IsEmpty(aBall.id) Then
      If Abs(XR - XL) > Abs(YR - YL) Then
        BallPos = PSlope(aBall.x, XL, 0, XR, 1)
      Else
        BallPos = PSlope(aBall.y, YL, 0, YR, 1)
      End If
      If BallPos < 0 Then BallPos = 0
      If BallPos > 1 Then BallPos = 1
    End If

    'Velocity angle correction
    If Not IsEmpty(ModIn(0) ) Then
      Dim Angle, RotVxVy
      Angle = LinearEnvelope(BallPos, ModIn, ModOut)
      '   debug.print " BallPos=" & BallPos &" Angle=" & Angle
      '   debug.print " BEFORE: aBall.Velx=" & aBall.Velx &" aBall.Vely" & aBall.Vely
      RotVxVy = RotPoint(aBall.Velx,aBall.Vely,Angle)
      If Enabled Then aBall.Velx = RotVxVy(0)
      If Enabled Then aBall.Vely = RotVxVy(1)
      '   debug.print " AFTER: aBall.Velx=" & aBall.Velx &" aBall.Vely" & aBall.Vely
      '   debug.print " "
    End If
  End Sub
End Class

Dim gBot
Sub GameTimer_Timer()
  'gBOT = GetBalls
  Cor.Update
  RollingUpdate
End Sub

'The CorTimer interval should be 10. It's sole purpose is to update the Cor (physics) calculations
'CorTimer.Interval = 10
'Sub CorTimer_Timer(): Cor.Update: End Sub


'--- Add this near the top of your script ---
Function IIf(condition, truePart, falsePart)
    If condition Then
        IIf = truePart
    Else
        IIf = falsePart
    End If
End Function

'******************************************************
'   ZRRL: RAMP ROLLING SFX
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
dim RampBalls(5,2)
'x,0 = ball x,1 = ID, 2 = Protection against ending early (minimum amount of updates)
'0,0 is boolean on/off, 0,1 unused for now
RampBalls(0,0) = False

' RampType
'     Setup: Set this array to the number Total number of balls that can be tracked at one time + 1.  5 ball multiball then set value to 6
'     Description: Array type indexed on BallId and a values used to deterimine what type of ramp the ball is on: False = Wire Ramp, True = Plastic Ramp
dim RampType(5)

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
      if RampBalls(x,0).Z < 130 and RampBalls(x, 2) > RampMinLoops then 'if ball is on the PF, remove  it
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


'Ramp triggers


'Ramp triggers

Sub RampStart1_hit
  If ActiveBall.Z > 126 Then
        WireRampOn True           ' true means plastic ramp here
  End If
End Sub

Sub RampStart2_hit
        WireRampOn True           ' true means plastic ramp here
End Sub

Sub RampStart3_hit
        WireRampOn True           ' true means plastic ramp here
End Sub

Sub RampStart4_hit
        WireRampOn True           ' true means plastic ramp here
End Sub

Sub RampStart5_hit
        WireRampOn True           ' true means plastic ramp here
End Sub

Sub WireStart1_hit
        WireRampOn False           ' true means plastic ramp here
End Sub

Sub WireStart2_hit
        WireRampOn False           ' true means plastic ramp here
End Sub

Sub WireStart3_hit
        WireRampOn False           ' true means plastic ramp here
End Sub

Sub WireStart4_hit
        WireRampOn False           ' true means plastic ramp here
End Sub

Sub WireStart5_hit
        WireRampOn False           ' true means plastic ramp here
End Sub

Sub Switch2Wire1_hit
    WireRampOff
    ' optional "clunk" sound:
    RandomSoundRampStop Me
        WireRampOn False          ' wire
End Sub

Sub Switch2Wire2_hit
    WireRampOff
    ' optional "clunk" sound:
    RandomSoundRampStop Me
        WireRampOn False          ' wire
End Sub

Sub Switch2Wire3_hit
    WireRampOff
    ' optional "clunk" sound:
    RandomSoundRampStop Me
        WireRampOn False          ' wire
End Sub

Sub Switch2Flat4_hit
    WireRampOff
    ' optional "clunk" sound:
    RandomSoundRampStop Me
        WireRampOn True          ' wire
End Sub

Sub Switch2Flat5_hit
    WireRampOff
    ' optional "clunk" sound:
    RandomSoundRampStop Me
        WireRampOn True          ' wire
End Sub

Sub trLowerRampStart_Hit()
    If ActiveBall.VelY < 0 Then   ' ball moving away from player
        WireRampOn True           ' true means plastic ramp here
    Else
        WireRampOff               ' stop sound
    End If
End Sub

Sub trUpperRampWireEnd_Hit
    WireRampOff
    ' optional "clunk" sound:
    RandomSoundRampStop Me
End Sub

Sub trLowerrampend_Hit 'end of ramp
    WireRampOff
    ' optional "clunk" sound:
    RandomSoundRampStop Me
End Sub

Sub trUpperRampWireStart_Hit
    WireRampOff
    ' optional "clunk" sound:
    RandomSoundRampStop Me
        WireRampOn False          ' wire
End Sub


Sub RampTrigger1_Hit
  If activeball.vely < 0 Then
    WireRampOn True
  Else
    WireRampOff
  End If
End Sub

Sub RampTrigger2_Hit
  If activeball.vely < 0 Then
    WireRampOn True
  Else
    WireRampOff
  End If
End Sub

Sub RampTrigger3_Hit
    if abs(activeball.AngMomZ) > 70 then activeball.AngMomZ = 50
    activeball.AngMomZ = -abs(activeball.AngMomZ) * 3
    WireRampOff
End Sub

Sub RampTrigger4_Hit
    if abs(activeball.AngMomZ) > 70 then activeball.AngMomZ = 50
    activeball.AngMomZ = abs(activeball.AngMomZ) * 3
    WireRampOff
End Sub

Sub RampTrigger5_Hit
  If activeball.vely < 0 Then
    WireRampOn True
  Else
    WireRampOff
  End If
End Sub

Sub RampTrigger6_Hit
  WireRampOff
End Sub

Sub RandomSoundRampStop(obj)
  Select Case Int(rnd*3)
    Case 0: PlaySoundAtVol "wireramp_stop1", obj, 0.02*volumedial:PlaySoundAtLevelActiveBall ("Rubber_Strong_1"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
    Case 1: PlaySoundAtVol "wireramp_stop2", obj, 0.02*volumedial:PlaySoundAtLevelActiveBall ("Rubber_Strong_2"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
    Case 2: PlaySoundAtVol "wireramp_stop3", obj, 0.02*volumedial:PlaySoundAtLevelActiveBall ("Rubber_1_Hard"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
  End Select
End Sub




'******************************************************
'**** END RAMP ROLLING SFX
'******************************************************

'******************************************************
' ZNFF:  FLIPPER CORRECTIONS by nFozzy
'******************************************************
'
' There are several steps for taking advantage of nFozzy's flipper solution.  At a high level we'll need the following:
' 1. flippers with specific physics settings
' 2. custom triggers for each flipper (TriggerLF, TriggerRF)
' 3. and, special scripting
'
' TriggerLF and RF should now be 27 vp units from the flippers. In addition, 3 degrees should be added to the end angle
' when creating these triggers.
'
' RF.ReProcessBalls Activeball and LF.ReProcessBalls Activeball must be added the flipper_collide subs.
'
' A common mistake is incorrect flipper length.  A 3-inch flipper with rubbers will be about 3.125 inches long.
' This translates to about 147 vp units.  Therefore, the flipper start radius + the flipper length + the flipper end
' radius should  equal approximately 147 vp units. Another common mistake is is that sometimes the right flipper
' angle was set with a large postive value (like 238 or something). It should be using negative value (like -122).
'
' The following settings are a solid starting point for various eras of pinballs.
' |                    | EM's           | late 70's to mid 80's | mid 80's to early 90's | mid 90's and later |
' | ------------------ | -------------- | --------------------- | ---------------------- | ------------------ |
' | Mass               | 1              | 1                     | 1                      | 1                  |
' | Strength           | 500-1000 (750) | 1400-1600 (1500)      | 2000-2600              | 3200-3300 (3250)   |
' | Elasticity         | 0.88           | 0.88                  | 0.88                   | 0.88               |
' | Elasticity Falloff | 0.15           | 0.15                  | 0.15                   | 0.15               |
' | Fricition          | 0.8-0.9        | 0.9                   | 0.9                    | 0.9                |
' | Return Strength    | 0.11           | 0.09                  | 0.07                   | 0.055              |
' | Coil Ramp Up       | 2.5            | 2.5                   | 2.5                    | 2.5                |
' | Scatter Angle      | 0              | 0                     | 0                      | 0                  |
' | EOS Torque         | 0.4            | 0.4                   | 0.375                  | 0.375              |
' | EOS Torque Angle   | 4              | 4                     | 6                      | 6                  |
'

'******************************************************
' Flippers Polarity (Select appropriate sub based on era)
'******************************************************

Dim LF : Set LF = New FlipperPolarity
Dim RF : Set RF = New FlipperPolarity
'Dim ULF : Set ULF = New FlipperPolarity

InitPolarity

'
''*******************************************
'' Late 70's to early 80's
'
'Sub InitPolarity()
'   dim x, a : a = Array(LF, RF)
' for each x in a
'   x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
'   x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
'   x.enabled = True
'   x.TimeDelay = 80
'   x.DebugOn=False ' prints some info in debugger
'
'
'        x.AddPt "Polarity", 0, 0, 0
'        x.AddPt "Polarity", 1, 0.05, - 2.7
'        x.AddPt "Polarity", 2, 0.16, - 2.7
'        x.AddPt "Polarity", 3, 0.22, - 0
'        x.AddPt "Polarity", 4, 0.25, - 0
'        x.AddPt "Polarity", 5, 0.3, - 1
'        x.AddPt "Polarity", 6, 0.4, - 2
'        x.AddPt "Polarity", 7, 0.5, - 2.7
'        x.AddPt "Polarity", 8, 0.65, - 1.8
'        x.AddPt "Polarity", 9, 0.75, - 0.5
'        x.AddPt "Polarity", 10, 0.81, - 0.5
'        x.AddPt "Polarity", 11, 0.88, 0
'        x.AddPt "Polarity", 12, 1.3, 0
'
'   x.AddPt "Velocity", 0, 0, 0.85
'   x.AddPt "Velocity", 1, 0.15, 0.85
'   x.AddPt "Velocity", 2, 0.2, 0.9
'   x.AddPt "Velocity", 3, 0.23, 0.95
'   x.AddPt "Velocity", 4, 0.41, 0.95
'   x.AddPt "Velocity", 5, 0.53, 0.95 '0.982
'   x.AddPt "Velocity", 6, 0.62, 1.0
'   x.AddPt "Velocity", 7, 0.702, 0.968
'   x.AddPt "Velocity", 8, 0.95,  0.968
'   x.AddPt "Velocity", 9, 1.03,  0.945
'   x.AddPt "Velocity", 10, 1.5,  0.945
'
' Next
'
' ' SetObjects arguments: 1: name of object 2: flipper object: 3: Trigger object around flipper
'    LF.SetObjects "LF", LeftFlipper, TriggerLF
'    RF.SetObjects "RF", RightFlipper, TriggerRF
'End Sub
'
'
'
''*******************************************
'' Mid 80's
'
Sub InitPolarity()
   dim x, a : a = Array(LF, RF)
  for each x in a
    x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
    x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
    x.enabled = True
    x.TimeDelay = 80
    x.DebugOn=False ' prints some info in debugger

    x.AddPt "Polarity", 0, 0, 0
    x.AddPt "Polarity", 1, 0.05, - 3.7
    x.AddPt "Polarity", 2, 0.16, - 3.7
    x.AddPt "Polarity", 3, 0.22, - 0
    x.AddPt "Polarity", 4, 0.25, - 0
    x.AddPt "Polarity", 5, 0.3, - 2
    x.AddPt "Polarity", 6, 0.4, - 3
    x.AddPt "Polarity", 7, 0.5, - 3.7
    x.AddPt "Polarity", 8, 0.65, - 2.3
    x.AddPt "Polarity", 9, 0.75, - 1.5
    x.AddPt "Polarity", 10, 0.81, - 1
    x.AddPt "Polarity", 11, 0.88, 0
    x.AddPt "Polarity", 12, 1.3, 0

    x.AddPt "Velocity", 0, 0, 0.85
    x.AddPt "Velocity", 1, 0.15, 0.85
    x.AddPt "Velocity", 2, 0.2, 0.9
    x.AddPt "Velocity", 3, 0.23, 0.95
    x.AddPt "Velocity", 4, 0.41, 0.95
    x.AddPt "Velocity", 5, 0.53, 0.95 '0.982
    x.AddPt "Velocity", 6, 0.62, 1.0
    x.AddPt "Velocity", 7, 0.702, 0.968
    x.AddPt "Velocity", 8, 0.95,  0.968
    x.AddPt "Velocity", 9, 1.03,  0.945
    x.AddPt "Velocity", 10, 1.5,  0.945

  Next

  ' SetObjects arguments: 1: name of object 2: flipper object: 3: Trigger object around flipper
    LF.SetObjects "LF", LeftFlipper, TriggerLF
    RF.SetObjects "RF", RightFlipper, TriggerRF
End Sub
'
''*******************************************
''  Late 80's early 90's
'
'Sub InitPolarity()
' dim x, a : a = Array(LF, RF)
' for each x in a
'   x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
'   x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
'   x.enabled = True
'   x.TimeDelay = 60
'   x.DebugOn=False ' prints some info in debugger
'
'   x.AddPt "Polarity", 0, 0, 0
'   x.AddPt "Polarity", 1, 0.05, - 5
'   x.AddPt "Polarity", 2, 0.16, - 5
'   x.AddPt "Polarity", 3, 0.22, - 0
'   x.AddPt "Polarity", 4, 0.25, - 0
'   x.AddPt "Polarity", 5, 0.3, - 2
'   x.AddPt "Polarity", 6, 0.4, - 3
'   x.AddPt "Polarity", 7, 0.5, - 4.0
'   x.AddPt "Polarity", 8, 0.7, - 3.5
'   x.AddPt "Polarity", 9, 0.75, - 3.0
'   x.AddPt "Polarity", 10, 0.8, - 2.5
'   x.AddPt "Polarity", 11, 0.85, - 2.0
'   x.AddPt "Polarity", 12, 0.9, - 1.5
'   x.AddPt "Polarity", 13, 0.95, - 1.0
'   x.AddPt "Polarity", 14, 1, - 0.5
'   x.AddPt "Polarity", 15, 1.1, 0
'   x.AddPt "Polarity", 16, 1.3, 0
'
'   x.AddPt "Velocity", 0, 0, 0.85
'   x.AddPt "Velocity", 1, 0.15, 0.85
'   x.AddPt "Velocity", 2, 0.2, 0.9
'   x.AddPt "Velocity", 3, 0.23, 0.95
'   x.AddPt "Velocity", 4, 0.41, 0.95
'   x.AddPt "Velocity", 5, 0.53, 0.95 '0.982
'   x.AddPt "Velocity", 6, 0.62, 1.0
'   x.AddPt "Velocity", 7, 0.702, 0.968
'   x.AddPt "Velocity", 8, 0.95,  0.968
'   x.AddPt "Velocity", 9, 1.03,  0.945
'   x.AddPt "Velocity", 10, 1.5,  0.945

' Next
'
' ' SetObjects arguments: 1: name of object 2: flipper object: 3: Trigger object around flipper
' LF.SetObjects "LF", LeftFlipper, TriggerLF
' RF.SetObjects "RF", RightFlipper, TriggerRF
'End Sub

'*******************************************
' Early 90's and after


'******************************************************
'  FLIPPER CORRECTION FUNCTIONS
'******************************************************

' modified 2023 by nFozzy
' Removed need for 'endpoint' objects
' Added 'createvents' type thing for TriggerLF / TriggerRF triggers.
' Removed AddPt function which complicated setup imo
' made DebugOn do something (prints some stuff in debugger)
'   Otherwise it should function exactly the same as before\
' modified 2024 by rothbauerw
' Added Reprocessballs for flipper collisions (LF.Reprocessballs Activeball and RF.Reprocessballs Activeball must be added to the flipper collide subs
' Improved handling to remove correction for backhand shots when the flipper is raised

Class FlipperPolarity
  Public DebugOn, Enabled
  Private FlipAt    'Timer variable (IE 'flip at 723,530ms...)
  Public TimeDelay    'delay before trigger turns off and polarity is disabled
  Private Flipper, FlipperStart, FlipperEnd, FlipperEndY, LR, PartialFlipCoef, FlipStartAngle
  Private Balls(20), balldata(20)
  Private Name

  Dim PolarityIn, PolarityOut
  Dim VelocityIn, VelocityOut
  Dim YcoefIn, YcoefOut
  Public Sub Class_Initialize
    ReDim PolarityIn(0)
    ReDim PolarityOut(0)
    ReDim VelocityIn(0)
    ReDim VelocityOut(0)
    ReDim YcoefIn(0)
    ReDim YcoefOut(0)
    Enabled = True
    TimeDelay = 50
    LR = 1
    Dim x
    For x = 0 To UBound(balls)
      balls(x) = Empty
      Set Balldata(x) = new SpoofBall
    Next
  End Sub

  Public Sub SetObjects(aName, aFlipper, aTrigger)

    If TypeName(aName) <> "String" Then MsgBox "FlipperPolarity: .SetObjects error: first argument must be a String (And name of Object). Found:" & TypeName(aName) End If
    If TypeName(aFlipper) <> "Flipper" Then MsgBox "FlipperPolarity: .SetObjects error: Second argument must be a flipper. Found:" & TypeName(aFlipper) End If
    If TypeName(aTrigger) <> "Trigger" Then MsgBox "FlipperPolarity: .SetObjects error: third argument must be a trigger. Found:" & TypeName(aTrigger) End If
    If aFlipper.EndAngle > aFlipper.StartAngle Then LR = -1 Else LR = 1 End If
    Name = aName
    Set Flipper = aFlipper
    FlipperStart = aFlipper.x
    FlipperEnd = Flipper.Length * Sin((Flipper.StartAngle / 57.295779513082320876798154814105)) + Flipper.X ' big floats for degree to rad conversion
    FlipperEndY = Flipper.Length * Cos(Flipper.StartAngle / 57.295779513082320876798154814105)*-1 + Flipper.Y

    Dim str
    str = "Sub " & aTrigger.name & "_Hit() : " & aName & ".AddBall ActiveBall : End Sub'"
    ExecuteGlobal(str)
    str = "Sub " & aTrigger.name & "_UnHit() : " & aName & ".PolarityCorrect ActiveBall : End Sub'"
    ExecuteGlobal(str)

  End Sub

  ' Legacy: just no op
  Public Property Let EndPoint(aInput)

  End Property

  Public Sub AddPt(aChooseArray, aIDX, aX, aY) 'Index #, X position, (in) y Position (out)
    Select Case aChooseArray
      Case "Polarity"
        ShuffleArrays PolarityIn, PolarityOut, 1
        PolarityIn(aIDX) = aX
        PolarityOut(aIDX) = aY
        ShuffleArrays PolarityIn, PolarityOut, 0
      Case "Velocity"
        ShuffleArrays VelocityIn, VelocityOut, 1
        VelocityIn(aIDX) = aX
        VelocityOut(aIDX) = aY
        ShuffleArrays VelocityIn, VelocityOut, 0
      Case "Ycoef"
        ShuffleArrays YcoefIn, YcoefOut, 1
        YcoefIn(aIDX) = aX
        YcoefOut(aIDX) = aY
        ShuffleArrays YcoefIn, YcoefOut, 0
    End Select
  End Sub

  Public Sub AddBall(aBall)
    Dim x
    For x = 0 To UBound(balls)
      If IsEmpty(balls(x)) Then
        Set balls(x) = aBall
        Exit Sub
      End If
    Next
  End Sub

  Private Sub RemoveBall(aBall)
    Dim x
    For x = 0 To UBound(balls)
      If TypeName(balls(x) ) = "IBall" Then
        If aBall.ID = Balls(x).ID Then
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
    Dim x
    For x = 0 To UBound(balls)
      If Not IsEmpty(balls(x)) Then
        pos = pSlope(Balls(x).x, FlipperStart, 0, FlipperEnd, 1)
      End If
    Next
  End Property

  Public Sub ProcessBalls() 'save data of balls in flipper range
    FlipAt = GameTime
    Dim x
    For x = 0 To UBound(balls)
      If Not IsEmpty(balls(x)) Then
        balldata(x).Data = balls(x)
      End If
    Next
    FlipStartAngle = Flipper.currentangle
    PartialFlipCoef = ((Flipper.StartAngle - Flipper.CurrentAngle) / (Flipper.StartAngle - Flipper.EndAngle))
    PartialFlipCoef = abs(PartialFlipCoef-1)
  End Sub

  Public Sub ReProcessBalls(aBall) 'save data of balls in flipper range
    If FlipperOn() Then
      Dim x
      For x = 0 To UBound(balls)
        If Not IsEmpty(balls(x)) Then
          if balls(x).ID = aBall.ID Then
            If isempty(balldata(x).ID) Then
              balldata(x).Data = balls(x)
            End If
          End If
        End If
      Next
    End If
  End Sub

  'Timer shutoff for polaritycorrect
  Private Function FlipperOn()
    If GameTime < FlipAt+TimeDelay Then
      FlipperOn = True
    End If
  End Function

  Public Sub PolarityCorrect(aBall)
    If FlipperOn() Then
      Dim tmp, BallPos, x, IDX, Ycoef, BalltoFlip, BalltoBase, NoCorrection, checkHit
      Ycoef = 1

      'y safety Exit
      If aBall.VelY > -8 Then 'ball going down
        RemoveBall aBall
        Exit Sub
      End If

      'Find balldata. BallPos = % on Flipper
      For x = 0 To UBound(Balls)
        If aBall.id = BallData(x).id And Not IsEmpty(BallData(x).id) Then
          idx = x
          BallPos = PSlope(BallData(x).x, FlipperStart, 0, FlipperEnd, 1)
          BalltoFlip = DistanceFromFlipperAngle(BallData(x).x, BallData(x).y, Flipper, FlipStartAngle)
          If ballpos > 0.65 Then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)                'find safety coefficient 'ycoef' data
        End If
      Next

      If BallPos = 0 Then 'no ball data meaning the ball is entering and exiting pretty close to the same position, use current values.
        BallPos = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
        If ballpos > 0.65 Then  Ycoef = LinearEnvelope(aBall.Y, YcoefIn, YcoefOut)                        'find safety coefficient 'ycoef' data
        NoCorrection = 1
      Else
        checkHit = 50 + (20 * BallPos)

        If BalltoFlip > checkHit or (PartialFlipCoef < 0.5 and BallPos > 0.22) Then
          NoCorrection = 1
        Else
          NoCorrection = 0
        End If
      End If

      'Velocity correction
      If Not IsEmpty(VelocityIn(0) ) Then
        Dim VelCoef
        VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)

        'If partialflipcoef < 1 Then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)

        If Enabled Then aBall.Velx = aBall.Velx*VelCoef
        If Enabled Then aBall.Vely = aBall.Vely*VelCoef
      End If

      'Polarity Correction (optional now)
      If Not IsEmpty(PolarityIn(0) ) Then
        Dim AddX
        AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR

        If Enabled and NoCorrection = 0 Then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef*VelCoef)
      End If
      If DebugOn Then debug.print "PolarityCorrect" & " " & Name & " @ " & GameTime & " " & Round(BallPos*100) & "%" & " AddX:" & Round(AddX,2) & " Vel%:" & Round(VelCoef*100)
    End If
    RemoveBall aBall
  End Sub
End Class

'******************************************************
'  FLIPPER POLARITY AND RUBBER DAMPENER SUPPORTING FUNCTIONS
'******************************************************

' Used for flipper correction and rubber dampeners
Sub ShuffleArray(ByRef aArray, byVal offset) 'shuffle 1d array
  Dim x, aCount
  aCount = 0
  ReDim a(UBound(aArray) )
  For x = 0 To UBound(aArray)   'Shuffle objects in a temp array
    If Not IsEmpty(aArray(x) ) Then
      If IsObject(aArray(x)) Then
        Set a(aCount) = aArray(x)
      Else
        a(aCount) = aArray(x)
      End If
      aCount = aCount + 1
    End If
  Next
  If offset < 0 Then offset = 0
  ReDim aArray(aCount-1+offset)   'Resize original array
  For x = 0 To aCount-1       'set objects back into original array
    If IsObject(a(x)) Then
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
  BallSpeed = Sqr(ball.VelX^2 + ball.VelY^2 + ball.VelZ^2)
End Function

' Used for flipper correction and rubber dampeners
Function PSlope(Input, X1, Y1, X2, Y2)    'Set up line via two points, no clamping. Input X, output Y
  Dim x, y, b, m
  x = input
  m = (Y2 - Y1) / (X2 - X1)
  b = Y2 - m*X2
  Y = M*x+b
  PSlope = Y
End Function

' Used for flipper correction
Class spoofball
  Public X, Y, Z, VelX, VelY, VelZ, ID, Mass, Radius
  Public Property Let Data(aBall)
    With aBall
      x = .x
      y = .y
      z = .z
      velx = .velx
      vely = .vely
      velz = .velz
      id = .ID
      mass = .mass
      radius = .radius
    End With
  End Property
  Public Sub Reset()
    x = Empty
    y = Empty
    z = Empty
    velx = Empty
    vely = Empty
    velz = Empty
    id = Empty
    mass = Empty
    radius = Empty
  End Sub
End Class

' Used for flipper correction and rubber dampeners
Function LinearEnvelope(xInput, xKeyFrame, yLvl)
  Dim y 'Y output
  Dim L 'Line
  'find active line
  Dim ii
  For ii = 1 To UBound(xKeyFrame)
    If xInput <= xKeyFrame(ii) Then
      L = ii
      Exit For
    End If
  Next
  If xInput > xKeyFrame(UBound(xKeyFrame) ) Then L = UBound(xKeyFrame)    'catch line overrun
  Y = pSlope(xInput, xKeyFrame(L-1), yLvl(L-1), xKeyFrame(L), yLvl(L) )

  If xInput <= xKeyFrame(LBound(xKeyFrame) ) Then Y = yLvl(LBound(xKeyFrame) )     'Clamp lower
  If xInput >= xKeyFrame(UBound(xKeyFrame) ) Then Y = yLvl(UBound(xKeyFrame) )    'Clamp upper

  LinearEnvelope = Y
End Function

'******************************************************
'  FLIPPER TRICKS
'******************************************************
' To add the flipper tricks you must
'  - Include a call to FlipperCradleCollision from within OnBallBallCollision subroutine
'  - Include a call the CheckLiveCatch from the LeftFlipper_Collide and RightFlipper_Collide subroutines
'  - Include FlipperActivate and FlipperDeactivate in the Flipper solenoid subs

RightFlipper.timerinterval = 1
Rightflipper.timerenabled = True

Sub RightFlipper_timer()
  FlipperTricks LeftFlipper, LFPress, LFCount, LFEndAngle, LFState
  FlipperTricks RightFlipper, RFPress, RFCount, RFEndAngle, RFState
  FlipperNudge RightFlipper, RFEndAngle, RFEOSNudge, LeftFlipper, LFEndAngle
  FlipperNudge LeftFlipper, LFEndAngle, LFEOSNudge,  RightFlipper, RFEndAngle
End Sub

Dim LFEOSNudge, RFEOSNudge

Sub FlipperNudge_old(Flipper1, Endangle1, EOSNudge1, Flipper2, EndAngle2)
  Dim b
     Dim BOT
     BOT = GetBalls

  If Flipper1.currentangle = Endangle1 And EOSNudge1 <> 1 Then
    EOSNudge1 = 1
    '   debug.print Flipper1.currentangle &" = "& Endangle1 &"--"& Flipper2.currentangle &" = "& EndAngle2
    If Flipper2.currentangle = EndAngle2 Then
      For b = 0 To UBound(BOT)
        If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper1) Then
          'Debug.Print "ball in flip1. exit"
          Exit Sub
        End If
      Next
      For b = 0 To UBound(BOT)
        If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper2) Then
          BOT(b).velx = BOT(b).velx / 1.3
          BOT(b).vely = BOT(b).vely - 0.5
        End If
      Next
    End If
  Else
    If Abs(Flipper1.currentangle) > Abs(EndAngle1) + 30 Then EOSNudge1 = 0
  End If
End Sub


Sub FlipperNudge(Flipper1, Endangle1, EOSNudge1, Flipper2, EndAngle2)
  Dim b
     Dim BOT
     BOT = GetBalls

  If Flipper1.currentangle = Endangle1 And EOSNudge1 <> 1 Then
    EOSNudge1 = 1
    '   debug.print Flipper1.currentangle &" = "& Endangle1 &"--"& Flipper2.currentangle &" = "& EndAngle2
    If Flipper2.currentangle = EndAngle2 Then
      For b = 0 To UBound(BOT)
        If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper1) Then
          'Debug.Print "ball in flip1. exit"
          Exit Sub
        End If
      Next
      For b = 0 To UBound(BOT)
        If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper2) Then
          BOT(b).velx = BOT(b).velx / 1.3
          BOT(b).vely = BOT(b).vely - 0.5
        End If
      Next
    End If
  Else
    If Abs(Flipper1.currentangle) > Abs(EndAngle1) + 30 Then EOSNudge1 = 0
  End If
End Sub


Dim FCCDamping: FCCDamping = 0.4

Sub FlipperCradleCollision(ball1, ball2, velocity)
  if velocity < 0.7 then exit sub   'filter out gentle collisions
    Dim DoDamping, coef
    DoDamping = false
    'Check left flipper
    If LeftFlipper.currentangle = LFEndAngle Then
    If FlipperTrigger(ball1.x, ball1.y, LeftFlipper) OR FlipperTrigger(ball2.x, ball2.y, LeftFlipper) Then DoDamping = true
    End If
    'Check right flipper
    If RightFlipper.currentangle = RFEndAngle Then
    If FlipperTrigger(ball1.x, ball1.y, RightFlipper) OR FlipperTrigger(ball2.x, ball2.y, RightFlipper) Then DoDamping = true
    End If
    If DoDamping Then
    coef = FCCDamping
        ball1.velx = ball1.velx * coef: ball1.vely = ball1.vely * coef: ball1.velz = ball1.velz * coef
        ball2.velx = ball2.velx * coef: ball2.vely = ball2.vely * coef: ball2.velz = ball2.velz * coef
    End If
End Sub



'*************************************************
'  Check ball distance from Flipper for Rem
'*************************************************

Function Distance(ax,ay,bx,by)
  Distance = Sqr((ax - bx) ^ 2 + (ay - by) ^ 2)
End Function

Function DistancePL(px,py,ax,ay,bx,by) 'Distance between a point and a line where point Is px,py
  DistancePL = Abs((by - ay) * px - (bx - ax) * py + bx * ay - by * ax) / Distance(ax,ay,bx,by)
End Function

Function Radians(Degrees)
  Radians = Degrees * PI / 180
End Function

Function AnglePP(ax,ay,bx,by)
  AnglePP = Atn2((by - ay),(bx - ax)) * 180 / PI
End Function

Function DistanceFromFlipper(ballx, bally, Flipper)
  DistanceFromFlipper = DistancePL(ballx, bally, Flipper.x, Flipper.y, Cos(Radians(Flipper.currentangle + 90)) + Flipper.x, Sin(Radians(Flipper.currentangle + 90)) + Flipper.y)
End Function

Function DistanceFromFlipperAngle(ballx, bally, Flipper, Angle)
  DistanceFromFlipperAngle = DistancePL(ballx, bally, Flipper.x, Flipper.y, Cos(Radians(Angle + 90)) + Flipper.x, Sin(Radians(angle + 90)) + Flipper.y)
End Function

Function FlipperTrigger(ballx, bally, Flipper)
  Dim DiffAngle
  DiffAngle = Abs(Flipper.currentangle - AnglePP(Flipper.x, Flipper.y, ballx, bally) - 90)
  If DiffAngle > 180 Then DiffAngle = DiffAngle - 360

  If DistanceFromFlipper(ballx,bally,Flipper) < 48 And DiffAngle <= 90 And Distance(ballx,bally,Flipper.x,Flipper.y) < Flipper.Length Then
    FlipperTrigger = True
  Else
    FlipperTrigger = False
  End If
End Function



'*************************************************
'  End - Check ball distance from Flipper for Rem
'*************************************************

Dim LFPress, RFPress, ULFPress, LFCount, RFCount
Dim LFState, RFState
Dim EOST, EOSA,Frampup, FElasticity,FReturn
Dim RFEndAngle, LFEndAngle

Const FlipperCoilRampupMode = 0 '0 = fast, 1 = medium, 2 = slow (tap passes should work)

LFState = 1
RFState = 1
EOST = leftflipper.eostorque
EOSA = leftflipper.eostorqueangle
Frampup = LeftFlipper.rampup
FElasticity = LeftFlipper.elasticity
FReturn = LeftFlipper.return
Const EOSTnew = 1.5 'EM's to late 80's - new recommendation by rothbauerw (previously 1)
'Const EOSTnew = 1.2 '90's and later - new recommendation by rothbauerw (previously 0.8)
Const EOSAnew = 1
Const EOSRampup = 0
Dim SOSRampup
Select Case FlipperCoilRampupMode
  Case 0
    SOSRampup = 2.5
  Case 1
    SOSRampup = 6
  Case 2
    SOSRampup = 8.5
End Select

Const LiveCatch = 16
Const LiveElasticity = 0.45
Const SOSEM = 0.815
'   Const EOSReturn = 0.055  'EM's
'   Const EOSReturn = 0.045  'late 70's to mid 80's
   Const EOSReturn = 0.035  'mid 80's to early 90's
'   Const EOSReturn = 0.025  'mid 90's and later

LFEndAngle = Leftflipper.endangle
RFEndAngle = RightFlipper.endangle

Sub FlipperActivate(Flipper, FlipperPress)
  FlipperPress = 1
  Flipper.Elasticity = FElasticity

  Flipper.eostorque = EOST
  Flipper.eostorqueangle = EOSA
End Sub

Sub FlipperDeactivate(Flipper, FlipperPress)
    Dim BOT, b

    FlipperPress = 0
    Flipper.eostorqueangle = EOSA
    Flipper.eostorque = EOST * EOSReturn / FReturn

    If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 0.1 Then
        BOT = GetBalls   ' fresh list of valid balls on the table

        If IsArray(BOT) Then
            For b = 0 To UBound(BOT)
                If IsObject(BOT(b)) Then
                    ' check for cradle near this flipper
                    If Distance(BOT(b).x, BOT(b).y, Flipper.x, Flipper.y) < 55 Then
                        ' clamp downward speed a bit
                        If BOT(b).VelY >= -0.4 Then BOT(b).VelY = -0.4
                    End If
                End If
            Next
        End If
    End If
End Sub

Sub FlipperTricks (Flipper, FlipperPress, FCount, FEndAngle, FState)
  Dim Dir
  Dir = Flipper.startangle / Abs(Flipper.startangle) '-1 for Right Flipper

  If Abs(Flipper.currentangle) > Abs(Flipper.startangle) - 0.05 Then
    If FState <> 1 Then
      Flipper.rampup = SOSRampup
      Flipper.endangle = FEndAngle - 3 * Dir
      Flipper.Elasticity = FElasticity * SOSEM
      FCount = 0
      FState = 1
    End If
  ElseIf Abs(Flipper.currentangle) <= Abs(Flipper.endangle) And FlipperPress = 1 Then
    If FCount = 0 Then FCount = GameTime

    If FState <> 2 Then
      Flipper.eostorqueangle = EOSAnew
      Flipper.eostorque = EOSTnew
      Flipper.rampup = EOSRampup
      Flipper.endangle = FEndAngle
      FState = 2
    End If
  ElseIf Abs(Flipper.currentangle) > Abs(Flipper.endangle) + 0.01 And FlipperPress = 1 Then
    If FState <> 3 Then
      Flipper.eostorque = EOST
      Flipper.eostorqueangle = EOSA
      Flipper.rampup = Frampup
      Flipper.Elasticity = FElasticity
      FState = 3
    End If
  End If
End Sub

Const LiveDistanceMin = 5  'minimum distance In vp units from flipper base live catch dampening will occur
Const LiveDistanceMax = 114 'maximum distance in vp units from flipper base live catch dampening will occur (tip protection)
Const BaseDampen = 0.55

Sub CheckLiveCatch(ball, Flipper, FCount, parm) 'Experimental new live catch
    Dim Dir, LiveDist
    Dir = Flipper.startangle / Abs(Flipper.startangle)    '-1 for Right Flipper
    Dim LiveCatchBounce   'If live catch is not perfect, it won't freeze ball totally
    Dim CatchTime
    CatchTime = GameTime - FCount
    LiveDist = Abs(Flipper.x - ball.x)

    If CatchTime <= LiveCatch And parm > 3 And LiveDist > LiveDistanceMin And LiveDist < LiveDistanceMax Then
        If CatchTime <= LiveCatch * 0.5 Then   'Perfect catch only when catch time happens in the beginning of the window
            LiveCatchBounce = 0
        Else
            LiveCatchBounce = Abs((LiveCatch / 2) - CatchTime)  'Partial catch when catch happens a bit late
        End If

        If LiveCatchBounce = 0 And ball.velx * Dir > 0 And LiveDist > 30 Then ball.velx = 0

        If ball.velx * Dir > 0 And LiveDist < 30 Then
            ball.velx = BaseDampen * ball.velx
            ball.vely = BaseDampen * ball.vely
            ball.angmomx = BaseDampen * ball.angmomx
            ball.angmomy = BaseDampen * ball.angmomy
            ball.angmomz = BaseDampen * ball.angmomz
        Elseif LiveDist > 30 Then
            ball.vely = LiveCatchBounce * (32 / LiveCatch) ' Multiplier for inaccuracy bounce
            ball.angmomx = 0
            ball.angmomy = 0
            ball.angmomz = 0
        End If
    Else
        If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 1 Then FlippersD.Dampenf ActiveBall, parm
    End If
End Sub

'******************************************************
'****  END FLIPPER CORRECTIONS
'******************************************************




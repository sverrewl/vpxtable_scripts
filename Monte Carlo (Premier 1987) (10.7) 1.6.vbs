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
If RenderingMode = 2 Then VRRoom = VRRoomChoice + 1 Else VRRoom = 0

If VRRoom=1 Then
  for each Obj4 in ColRoomMinimal : Obj4.visible = 1 : next
  for each Obj4 in ColBackdrop : Obj4.visible = 0 : next
End If
If VRRoom=0 Then
  for each Obj4 in ColRoomMinimal : Obj4.visible = 0 : next
  for each Obj4 in ColBackdrop : Obj4.visible = 1 : next
End If

'******************************************************************************************************

if GlassScratch = 1 then GlassImpurities.visible = true:GlassImpurities1.visible = true
if GlassScratch = 0 then GlassImpurities.visible = false:GlassImpurities1.visible = false

if VRPosterR = 1 then VR_Poster1R.visible = true:VR_Poster2R.visible = true
if VRPosterR = 0 then VR_Poster1R.visible = false:VR_Poster2R.visible = false

if VRPosterL = 1 then VR_Poster3L.visible = true:VR_Poster4L.visible = true
if VRPosterL = 0 then VR_Poster3L.visible = false:VR_Poster4L.visible = false

if Glass = 1 then WindowGlass.visible = true
if Glass = 0 then WindowGlass.visible = false

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


Const cGameName="mntecrlo",UseSolenoids=1,UseLamps=1,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff",SFlipperOn="fx_Flipperup",SFlipperOff="fx_Flipperdown"
Const SCoin="coin3",cCredits=""

LoadVPM "01210000","sys80.vbs",3.10

'Sub LoadVPM(VPMver, VBSfile, VBSver)
' On Error Resume Next
'   If ScriptEngineMajorVersion < 5 Then MsgBox "VB Script Engine 5.0 or higher required"
'    ExecuteGlobal GetTextFile(VBSFile)
'   If Err Then MsgBox "Unable to open " & VBSfile & ". Ensure that it is in the same folder as this table. " & vbNewLine & Err.Description : Err.Clear
'    Set Controller = CreateObject("VPinMAME.Controller")
'   If Err Then MsgBox "Can't Load VPinMAME." & vbNewLine & Err.Description
'    If VPMver>"" Then If Controller.Version < VPMver Or Err Then MsgBox "VPinMAME ver " & VPMver & " required." : Err.Clear
'    If VPinMAMEDriverVer < VBSver Or Err Then MsgBox VBSFile & " ver " & VBSver & " or higher required."
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
        PlaySoundAtVol "ballrel", BallRelease, 1
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
   'Kicker6.Enabled = False
    'Kicker7.Enabled = False
    If LI13=0 Then
    Ramp12.HeightBottom = 100
    Ramp12.Collidable = True
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
solcallback(3)="K3.SolOut"
solcallback(4)="K2.SolOut"
Solcallback(7)="K1.SolOut"
SolCallback(8)="vpmSolSound""knocker"","
SolCallback(9)="SolOutHole"
SolCallback(sLLFlipper)="vpmSolFlipper LeftFlipper,nothing,"
SolCallback(sLRFlipper)="vpmSolFlipper RightFlipper,nothing,"

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

' Thalamus : Was missing 'vpminit me'
vpminit me

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
   dt123.InitSnd "flapclos","flapopen"
   dt123.CreateEvents "dt123"

   Set dt45=New cvpmDropTarget
   dt45.InitDrop Array(Target4,Target5),Nothing
   dt45.initsnd "flapclos","flapopen"
   dt45.CreateEvents "dt45"

   Set dt678=New cvpmDropTarget
   dt678.InitDrop Array(Target6,Target7,Target8),Nothing
   dt678.InitSnd "flapclos","flapopen"
   dt678.CreateEvents "dt678"

   Set dt910=New cvpmDropTarget
   dt910.InitDrop Array(Target9,Target10),Nothing
   dt910.InitSnd "flapclos","flapopen"
   dt910.CreateEvents "dt910"

    ' Thalamus - more randomness to kickers pls
   Set K1=New cvpmBallstack
   K1.KickForceVar = 3
   K1.KickAngleVar = 3
   K1.InitSaucer Kicker1,45,165,1
   K1.KickAngleVar=5
   K1.InitExitSnd "popper","solenoid"

   Set K2=New cvpmBallstack
   K2.KickForceVar = 3
   K2.KickAngleVar = 3
   K2.InitSaucer Kicker2,35,100,1
   K2.KickAngleVar=5
   K2.InitExitSnd "popper","solenoid"

   Set K3=new cvpmBallstack
   K3.KickForceVar = 3
   K3.KickAngleVar = 3
   K3.InitSaucer Kicker3,25,270,1
   K3.KickAngleVar=5
   K3.InitExitSnd "popper","solenoid"

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
End Sub

Sub Table1_exit()

  SaveLUT 'add this line in the sub Table1_exit

End Sub

Sub Table1_KeyDown(ByVal KeyCode)
'************************************************************
  If KeyCode = LeftFlipperKey Then
    ' Animate VR Left flipper
    VR_CabFlipperLeft.X = VR_CabFlipperLeft.X +10
  End If
  If KeyCode = RightFlipperKey Then
    ' Animate VR Right flipper
    VR_CabFlipperRight.X = VR_CabFlipperRight.X - 10
  End If
  If Keycode = StartGameKey Then
    ' Animate VR Startbutton
    VR_Cab_StartButton.y = VR_Cab_StartButton.y -5
  End If
  If keycode = PlungerKey Then
    Plunger.Pullback
    TimerPlunger.Enabled = True
    TimerPlunger2.Enabled = False
  End If
' LUT-Changer
  If Keycode = LeftMagnaSave Then
        LUTSet = LUTSet  + 1
    if LutSet > 15 then LUTSet = 0
        lutsetsounddir = 1
      If LutToggleSound then
        If lutsetsounddir = 1 And LutSet <> 15 Then
          Playsound "click", 0, 1, 0.1, 0.25
        End If
      If lutsetsounddir = -1 And LutSet <> 15 Then
          Playsound "click", 0, 1, 0.1, 0.25
        End If
      If LutSet = 15 Then
          Playsound "gun", 0, 1, 0.1, 0.25
        End If
        LutSlctr.enabled = true
      SetLUT
      ShowLUT
    End If
  End If
'************************************************************
  If KeyCode=RightFlipperKey Then Controller.Switch(75)=1
 If KeyCode=PlungerKey Then Plunger.Pullback
 If vpmKeyDown(KeyCode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
'************************************************************
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
 If KeyCode=PlungerKey Then Plunger.Fire:PlaySoundAtVol "plunger", Plunger ,1
 If vpmKeyUp(KeyCode) Then Exit Sub
End Sub

Sub Drain_Hit:Controller.Switch(65)=1:DrainBalls=True:End Sub
Sub Bumper1a_Hit:vpmTimer.PulseSw 20:PlaySoundAtVol "jet1", ActiveBall, 1:End Sub
Sub Bumper2a_Hit:vpmTimer.PulseSw 30:PlaySoundAtVol "jet1", ActiveBall, 1:End Sub
Sub Bumper3a_Hit:vpmTimer.PulseSw 40:PlaySoundAtVol "jet1", ActiveBall, 1:End Sub
Sub Kicker1_Hit:K1.AddBall 0:kickmag.MagnetOn = False:End Sub
Sub Kicker2_Hit:K2.AddBall 0:End Sub
Sub Kicker3_Hit:K3.AddBall 0:End Sub
Sub LeftSlingshot_Slingshot()
  vpmTimer.PulseSw 53
  PlaySoundAtVol SoundFX("left_slingshot",DOFContactors), ActiveBall, 1
End Sub
Sub RightSlingshot_Slingshot
  vpmTimer.PulseSw 53
  PlaySoundAtVol SoundFX("right_slingshot",DOFContactors), ActiveBall, 1
End Sub
Dim MY
Sub Kicker4_Hit:MY=-INT(ActiveBall.VelY)/3:Me.DestroyBall:Kicker4.TimerEnabled=True:End Sub
Sub Kicker4_Timer:Kicker5.CreateBall:Kicker5.Kick 180,MY:Kicker4.TimerEnabled=False:End Sub
'Sub Kicker6_Hit:Me.DestroyBall:Kicker7.CreateBall:Kicker7.Kick 0,20:End Sub

'Sub Roulette_Timer
'If TTable.MotorOn=True Then
'  RWall(Frame).IsDropped=1
' Frame=Frame+1
'  If Frame>11 Then Frame=0
' RWall(Frame).IsDropped=0
'Else
'  RWall(Frame).IsDropped=1
' Frame=Frame+1
'  If Frame>11 Then Frame=0
' RWall(Frame).IsDropped=0
'If TTable.Speed<1 Then
'  Kicker9.Enabled=1
'  Kicker10.Enabled=1
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
'  Kicker9.Enabled=0
'  Kicker10.Enabled=0
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
    Flasher2.Visible=True:Flasher3.Visible=True:Flasher1.Visible=True:Flasher4.Visible=True::bg3.Visible=False:bg4.Visible=False:bg5.Visible=False
  Else
    Flasher2.Visible=False:Flasher3.Visible=False:Flasher1.Visible=False:Flasher4.Visible=False:bg3.Visible=True:bg4.Visible=True:bg5.Visible=True
  End If
  If Light14.State=1 then
    Flasher5.Visible=True:Flasher6.Visible=True:Flasher7.Visible=True:Flasher8.Visible=True:bg1.Visible=False:bg2.Visible=False:bg6.Visible=False
  Else
    Flasher5.Visible=False:Flasher6.Visible=False:Flasher7.Visible=False:Flasher8.Visible=False:bg1.Visible=True:bg2.Visible=True:bg6.Visible=True
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
Set MyBall=SW10.CreateBall
SW10.Kick 0,1

Dim LoopObj, spinctr
spinctr = 0

'MAINTIMER THAT SPINS THE BALL AND THE DISK, ENABLED WHEN LIGHT17A IS ON
Sub spinnertimer_Timer()
  If Light17a.State=1 Then
    toppertimer.Enabled = True:w009.TimerEnabled = True
    Spindisk.RotZ=Spindisk.RotZ+30:cone.RotY=cone.RotY+30:sphere.RotY=sphere.RotY+30
    If Spindisk.RotZ = 360 then Spindisk.RotZ = 0:cone.RotY = -95:sphere.RotY = 0
    If TTable.MotorOn=False then
      For Each LoopObj in spinnerkickers
        LoopObj.Kick 0, 1
        LoopObj.Enabled=false
      Next
      TTable.MotorOn=True
      TTable.AddBall MyBall
      TTable.AffectBall MyBall
    End If
    spinctr=spinctr+1
    If spinctr>11 then spinctr=0
  End If
  If TTable.Speed>=20 then
    TTable.MotorOn=False
  End If
  If TTable.Speed =< 1 and TTable.MotorOn=false Then
    For Each LoopObj in spinnerkickers
      LoopObj.Enabled=True
    Next
    TTable.RemoveBall MyBall
  End If
End Sub

'ALL 12 OF THE KICKERS IN THE SPINNING ROULETTE WHEEL.
Dim kickernum:kickernum = 0
Sub SW00_Hit:RWheel spinctr,0,1:kickernum=0:End Sub
'Sub SW00_Unhit:RWheel spinctr,0,0:End Sub
Sub SW01_Hit:RWheel spinctr,1,1:kickernum=1:End Sub
'Sub SW01_Unhit:RWheel spinctr,1,0:End Sub
Sub SW02_Hit:RWheel spinctr,2,1:kickernum=2:End Sub
'Sub SW02_Unhit:RWheel spinctr,2,0:End Sub
Sub SW03_Hit:RWheel spinctr,3,1:kickernum=3:End Sub
'Sub SW03_Unhit:RWheel spinctr,3,0:End Sub
Sub SW04_Hit:RWheel spinctr,4,1:kickernum=4:End Sub
'Sub SW04_Unhit:RWheel spinctr,4,0:End Sub
Sub SW05_Hit:RWheel spinctr,5,1:kickernum=5:End Sub
'Sub SW05_Unhit:RWheel spinctr,5,0:End Sub
Sub SW10_Hit:RWheel spinctr,6,1:kickernum=6:End Sub
'Sub SW10_Unhit:RWheel spinctr,6,0:End Sub
Sub SW11_Hit:RWheel spinctr,7,1:kickernum=7:End Sub
'Sub SW11_Unhit:RWheel spinctr,7,0:End Sub
Sub SW12_Hit:RWheel spinctr,8,1:kickernum=8:End Sub
'Sub SW12_Unhit:RWheel spinctr,8,0:End Sub
Sub SW13_Hit:RWheel spinctr,9,1:kickernum=9:End Sub
'Sub SW13_Unhit:RWheel spinctr,9,0:End Sub
Sub SW14_Hit:RWheel spinctr,10,1:kickernum=10:End Sub
'Sub SW14_Unhit:RWheel spinctr,10,0:End Sub
Sub SW15_Hit:RWheel spinctr,11,1:kickernum=11:End Sub
'Sub SW15_Unhit:RWheel spinctr,11,0:End Sub

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
'LUT (Colour Look Up Table)

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

Dim LUTset, LutToggleSound
LutToggleSound = True
LoadLUT
SetLUT

'LUT selector timer

dim lutsetsounddir
sub LutSlctr_timer
  If lutsetsounddir = 1 And LutSet <> 15 Then
    Playsound "click", 0, 1, 0.1, 0.25
  End If
  If lutsetsounddir = -1 And LutSet <> 15 Then
    Playsound "click", 0, 1, 0.1, 0.25
  End If
  If LutSet = 15 Then
    Playsound "gun", 0, 1, 0.1, 0.25
  End If
  LutSlctr.enabled = False
end sub

'LUT Subs

Sub SetLUT
  Table1.ColorGradeImage = "LUT" & LUTset
end sub

Sub LUTBox_Timer
  LUTBox.TimerEnabled = 0
  LUTBox.Visible = 0
  LUTBack.visible = 0
  VRLutdesc.visible = 0
End Sub

Sub ShowLUT
  LUTBox.visible = 1
  LUTBack.visible = 1
  VRLutdesc.visible = 1

  Select Case LUTSet
        Case 0: LUTBox.text = "Fleep Natural Dark 1": VRLUTdesc.imageA = "LUTcase0"
    Case 1: LUTBox.text = "Fleep Natural Dark 2": VRLUTdesc.imageA = "LUTcase1"
    Case 2: LUTBox.text = "Fleep Warm Dark": VRLUTdesc.imageA = "LUTcase2"
    Case 3: LUTBox.text = "Fleep Warm Bright": VRLUTdesc.imageA = "LUTcase3"
    Case 4: LUTBox.text = "Fleep Warm Vivid Soft": VRLUTdesc.imageA = "LUTcase4"
    Case 5: LUTBox.text = "Fleep Warm Vivid Hard": VRLUTdesc.imageA = "LUTcase5"
    Case 6: LUTBox.text = "Skitso Natural and Balanced": VRLUTdesc.imageA = "LUTcase6"
    Case 7: LUTBox.text = "Skitso Natural High Contrast": VRLUTdesc.imageA = "LUTcase7"
    Case 8: LUTBox.text = "3rdaxis Referenced THX Standard": VRLUTdesc.imageA = "LUTcase8"
    Case 9: LUTBox.text = "CalleV Punchy Brightness and Contrast": VRLUTdesc.imageA = "LUTcase9"
    Case 10: LUTBox.text = "HauntFreaks Desaturated" : VRLUTdesc.imageA = "LUTcase10"
      Case 11: LUTBox.text = "Tomate washed out": VRLUTdesc.imageA = "LUTcase11"
    Case 12: LUTBox.text = "VPW original 1on1": VRLUTdesc.imageA = "LUTcase12"
    Case 13: LUTBox.text = "bassgeige": VRLUTdesc.imageA = "LUTcase13"
    Case 14: LUTBox.text = "blacklight": VRLUTdesc.imageA = "LUTcase14"
    Case 15: LUTBox.text = "B&W Comic Book": VRLUTdesc.imageA = "LUTcase15"
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

  if LUTset = "" then LUTset = 12 'failsafe

  Set ScoreFile=FileObj.CreateTextFile(UserDirectory & "TableLUT.txt",True) 'Rename the tableLUT
  ScoreFile.WriteLine LUTset
  Set ScoreFile=Nothing
  Set FileObj=Nothing

End Sub

Sub LoadLUT

  Dim FileObj, ScoreFile, TextStr
  dim rLine

  Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    LUTset=12
    Exit Sub
  End if
  If Not FileObj.FileExists(UserDirectory & "TableLUT.txt") then  'Rename the tableLUT
    LUTset=12
    Exit Sub
  End if
  Set ScoreFile=FileObj.GetFile(UserDirectory & "TableLUT.txt")  'Rename the tableLUT
  Set TextStr=ScoreFile.OpenAsTextStream(1,0)
    If (TextStr.AtEndOfStream=True) then
      Exit Sub
    End if
    rLine = TextStr.ReadLine
    If rLine = "" then
      LUTset=12
      Exit Sub
    End if
    LUTset = int (rLine)
    Set ScoreFile = Nothing
      Set FileObj = Nothing
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
'            Supporting Ball, Sound Functions and Math
'*********************************************************************

Function RndNum(min, max)
  RndNum = Int(Rnd() * (max-min + 1) ) + min ' Sets a random number between min and max
End Function

Const Pi = 3.1415927

Function dSin(degrees)
  dsin = sin(degrees * Pi/180)
End Function

Function dCos(degrees)
  dcos = cos(degrees * Pi/180)
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
        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 130 AND BOT(b).z > 100 Then
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b))+5, AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If

        ' play ball drop sounds
        If BOT(b).VelZ < -1 and BOT(b).z < 155 and BOT(b).z > 127 Then 'height adjust for ball drop sounds
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
' ninuzzu's FLIPPER SHADOWS v2
'*****************************************

Sub LeftFlipper_Timer()
    FlipperLSh.RotZ = LeftFlipper.CurrentAngle
End Sub

Sub RightFlipper_Timer()
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

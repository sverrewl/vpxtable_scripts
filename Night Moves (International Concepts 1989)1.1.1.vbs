'*******************************************************************************************************
'
'                  Night Moves International Concepts 1989 v1.1.0
'               http://www.ipdb.org/machine.cgi?id=3507
'
'                     Created by Kiwi
'
'*******************************************************************************************************

Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

'********************************************** OPTIONS ************************************************

'****************************************** Volume Settings **********************************

Const RolVol = 1  'Ball Rolling
Const ColVol = 1  'Ball Collision
Const RubVol = 1  'Rubbers Collision
Const NudVol = 1  'Nudge
Const BreVol = 1  'Ball Release
Const PlpVol = 1  'Plunger pull
Const PlfVol = 1  'Plunger fire
Const SliVol = 1  'Slingshots
Const BumVol = 1  'Bumpers
Const SwiVol = 1  'Rollovers
Const TarVol = 1  'Targets
Const GatVol = 1  'Gates
Const SpiVol = 1  'Spinner
Const KicVol = 0.5  'Kicker catch
Const KieVol = 1  'Kicker eject
Const KidVol = 1  'Kicker Drain
Const FluVol = 1  'Flippers up
Const FldVol = 1  'Flippers down
Const KnoVol = 1  'Knocker

'************************ ROM

Const cGameName = "nmoves"

'************************ Ball: 50 unit is standard ball size ** Mass=(50^3)/125000 ,(BallSize^3)/125000

Const BallSize = 50

'Const BallMass = 1

'************************ Ball Shadow: 0 hidden , 1 visible

Const BallSHW = 1

'************************ Cab Options , 0 invisible, 1 visible

BlackFrame.Visible = 1
MetalRail.Visible = 1

'************************ Color Grading LUT: 1=Active, any other value=disabled

Const LutEnabled = 1

'************************ Slingshot mode: 0 Standard with walls , 1 Flippers

Const SlingM = 0

'************************ Slingshot hit threshold, with flippers (parm)

Const ThSling = 3

'************************ FlexDMD: 0 invisible, 1 visible

Const FlexDMDShow = 0

'************************ 0 Playfield Digits, 1 Playfield DMD (FlexDMD Flasher)

Const PlayfieldDMD = 0

'******************************************** OPTIONS END **********************************************

LoadVPM "01120100", "SYS80.vbs", 3.37

Dim bsTrough, bsLK, bsRK, x, PinPlay, UseFlexDMD

Const UseSolenoids = 2
Const UseLamps = 0
Const UseGI = 0
Const UseSync = 0
Const HandleMech = 0

' Standard Sounds
'Const SSolenoidOn = ""
'Const SSolenoidOff = ""
'Const SFlipperOn = ""
'Const SFlipperOff = ""
Const SCoin = "coin3"

'************
' Table init.
'************

Sub Table1_Init

' Thalamus : Was missing 'vpminit me'
  vpminit me

  With Controller
    .GameName = cGameName
    .SplashInfoLine = "Night Moves (International Concepts 1989)" & vbNewLine & "VPX table by Kiwi 1.1.0"
    .HandleMechanics = 0
    .HandleKeyboard = 0
    .ShowDMDOnly = 1
    .ShowFrame = 0
    .ShowTitle = 0
    .Hidden = 1
'   .Games(cGameName).Settings.Value("dmd_pos_x")=0
'   .Games(cGameName).Settings.Value("dmd_pos_y")=0
'   .Games(cGameName).Settings.Value("dmd_width")=400
'   .Games(cGameName).Settings.Value("dmd_height")=92
'   .Games(cGameName).Settings.Value("rol") = 0
'   .Games(cGameName).Settings.Value("sound") = 1
'   .Games(cGameName).Settings.Value("ddraw") = 1
    .Games(cGameName).Settings.Value("dmd_red")=50
    .Games(cGameName).Settings.Value("dmd_green")=100
    .Games(cGameName).Settings.Value("dmd_blue")=220
    If Err Then MsgBox Err.Description
  End With
  On Error Goto 0
  Controller.SolMask(0) = 0
  vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the timer to renable all the solenoids after 2 seconds
  Controller.Run

  ' Nudging
  vpmNudge.TiltSwitch = 57
  vpmNudge.Sensitivity = 3
  vpmNudge.TiltObj = Array(Bumper1, Bumper2, LeftSlingShot, RightSlingShot)

  ' Trough
' Set bsTrough = New cvpmBallStack
' With bsTrough
'   .InitSw 56, 0, 46, 0, 0, 0, 0, 0
'   .InitKick BallRelease, 63, 6
'   .InitEntrySnd "Solenoid", "Solenoid"
'   .InitExitSnd "popper", "Solenoid"
'   .Balls = 2
' End With

Set bsTrough = New cvpmTrough
  With bsTrough
     .Size = 1
     .EntrySw = 56
     .InitSwitches Array(46)
     .InitExit BallRelease, 53, 12
'     .InitEntrySounds "Solenoid", "Solenoid", "Solenoid"
'     .InitExitSounds "popper", "popper"
'     .InitExitSounds SoundFX("popper",DOFContactors), SoundFX("popper",DOFContactors)
     .Balls = 1
'     .CreateEvents "bsTrough", Outhole
'     .isDebug = True
End With

  ' Top Left Kicker
  Set bsLK = New cvpmBallStack
  With bsLK
    .InitSaucer sw26, 26, 108, 10
'   .InitExitSnd "popper", "Solenoid2"
    .KickForceVar = 0.5
    .KickAngleVar = 0.5
  End With

  ' Top Right Kicker
  Set bsRK = New cvpmBallStack
  With bsRK
    .InitSaucer sw36, 36, 200, 8
'   .InitExitSnd "popper", "Solenoid2"
    .KickForceVar = 0.5
    .KickAngleVar = 0.5
  End With

' Main Timer init
  PinMAMETimer.Interval = PinMAMEInterval
  PinMAMETimer.Enabled = 1

' Init drain
     DrainStart.CreateBall:DrainStart.Kick 90, 1
     vpmTimer.AddTimer 2000, "RollingTimer.Enabled=1'"

' Init Lut
  TextLUT.Visible = 0
  TextLUT.Text = Table1.ColorGradeImage
  LoadLut

' FlexDMD
 If FlexDMDShow + PlayfieldDMD > 0 Then UseFlexDMD = 1
 If UseFlexDMD Then FlexDMD_Init
  DotMatrix.TimerEnabled = PlayfieldDMD * UseFlexDMD

End Sub

Sub Table1_Exit()
  Controller.Stop
 If UseFlexDMD Then
    If Not FlexDMD is Nothing Then
      FlexDMD.Show = False
      FlexDMD.Run = False
      FlexDMD = NULL
    End If
  End If
End Sub
Sub Table1_Paused:Controller.Pause = 1:End Sub
Sub Table1_unPaused:Controller.Pause = 0:End Sub

'**********
' Keys
'**********

Sub Table1_KeyDown(ByVal KeyCode)
  If KeyCode = LeftFlipperKey Then Controller.Switch(52) = 1:PulsanteSx.TransX = 8
  If KeyCode = RightFlipperKey Then Controller.Switch(53) = 1:PulsanteDx.TransX = 8
  If keycode = LeftMagnaSave Then bLutActive = True:TextLUT.Visible = 1
  If keycode = LeftMagnaSave Then Controller.Switch(6) = 1:AvanzaSx.TransX = 8
  If keycode = RightMagnaSave Then Controller.Switch(16) = 1:AvanzaDx.TransX = 8
  If keycode = RightMagnaSave Then
  If bLutActive And LutEnabled = 1 Then NextLUT:End If
  End If
  If keycode = PlungerKey Then Plunger.Pullback:PlaySoundAtVol "fx_plungerpull", Plunger, PlpVol
  If keycode = LeftTiltKey Then Nudge 90, 4:PlaySound SoundFX("fx_nudge",0), 0, NudVol, -0.1, 0.25
  If keycode = RightTiltKey Then Nudge 270, 4:PlaySound SoundFX("fx_nudge",0), 0, NudVol, 0.1, 0.25
  If keycode = CenterTiltKey Then Nudge 0, 5:PlaySound SoundFX("fx_nudge",0), 0, NudVol, 0, 0.25
  If keycode = StartGameKey Then PulsanteStart.TransX = 8
  If vpmKeyDown(keycode) Then Exit Sub

    'debug key
    If KeyCode = "3" Then
        SetLamp 131, 1
        SetLamp 133, 1
        SetLamp 134, 1
        SetLamp 137, 1
    End If
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
  If KeyCode = LeftFlipperKey Then Controller.Switch(52) = 0:PulsanteSx.TransX = 0
  If KeyCode = RightFlipperKey Then Controller.Switch(53) = 0:PulsanteDx.TransX = 0
  If keycode = LeftMagnaSave Then Controller.Switch(6) = 0:bLutActive = False:TextLUT.Visible = 0:AvanzaSx.TransX = 0
  If keycode = RightMagnaSave Then Controller.Switch(16) = 0:AvanzaDx.TransX = 0
  If keycode = PlungerKey Then Plunger.Fire:PlaySoundAtVol "plunger2", Plunger, PlfVol * (Plunger.Position/25)
  If keycode = StartGameKey Then PulsanteStart.TransX = 0
  If vpmKeyUp(keycode) Then Exit Sub

    'debug key
    If KeyCode = "3" Then
        SetLamp 131, 0
        SetLamp 133, 0
        SetLamp 134, 0
        SetLamp 137, 0
    End If
End Sub

'*********
'   LUT
'*********

Dim bLutActive, LUTImage

Sub LoadLUT
  bLutActive = False
    x = LoadValue(cGameName, "LUTImage")
    If(x <> "") Then LUTImage = x Else LUTImage = 0
  UpdateLUT
End Sub

Sub SaveLUT
    SaveValue cGameName, "LUTImage", LUTImage
End Sub

Sub NextLUT:LUTImage = (LUTImage +1) MOD 9:UpdateLUT:SaveLUT:End Sub

Sub UpdateLUT
Select Case LutImage
Case 0:Table1.ColorGradeImage = "LUT0"
Case 1:Table1.ColorGradeImage = "LUT1"
Case 2:Table1.ColorGradeImage = "LUT2"
Case 3:Table1.ColorGradeImage = "LUT3"
Case 4:Table1.ColorGradeImage = "LUT4"
Case 5:Table1.ColorGradeImage = "LUT5"
Case 6:Table1.ColorGradeImage = "LUT6"
Case 7:Table1.ColorGradeImage = "LUT7"
Case 8:Table1.ColorGradeImage = "LUT8"
End Select
  TextLUT.Text = Table1.ColorGradeImage
End Sub

'*********
' Switches
'*********

' Slings
Dim LStep, RStep

Sub LeftSlingshot_Slingshot:vpmTimer.PulseSw 42:LeftSling.Visible=1:SxEmKickerT1.TransX=-25:LStep=0:Me.TimerEnabled=1:PlaySoundAtVol SoundFX("Slingshot",DOFContactors), SxEmKickerT1, SliVol:End Sub
Sub LeftSlingshot_Timer
  Select Case LStep
    Case 0:LeftSling.Visible = 1
    Case 1: 'pause
    Case 2:LeftSling.Visible = 0 :LeftSling1.Visible = 1:SxEmKickerT1.TransX=-21
    Case 3:LeftSling1.Visible = 0:LeftSling2.Visible = 1:SxEmKickerT1.TransX=-17:SlingFSx1.RotateToStart:SlingFSx2.RotateToStart
    Case 4:LeftSling2.Visible = 0:Me.TimerEnabled = 0:SxEmKickerT1.TransX=0
  End Select
  LStep = LStep + 1
End Sub

Sub SlingFSx1_Collide(parm)
 If SxEmKickerT1.TransX=0 And PinPlay=1 And parm > ThSling Then
  vpmTimer.PulseSw 42:PlaySoundAtVol SoundFX("Slingshot",DOFContactors), SxEmKickerT1, SliVol:LStep=0
  LeftSling.Visible=1:SxEmKickerT1.TransX=-25:LeftSlingshot.TimerEnabled=1:SlingFSx1.RotateToEnd:SlingFSx2.RotateToEnd
End If
End Sub
Sub SlingFSx2_Collide(parm)
 If SxEmKickerT1.TransX=0 And PinPlay=1 And parm > ThSling Then
  vpmTimer.PulseSw 42:PlaySoundAtVol SoundFX("Slingshot",DOFContactors), SxEmKickerT1, SliVol:LStep=0
  LeftSling.Visible=1:SxEmKickerT1.TransX=-25:LeftSlingshot.TimerEnabled=1:SlingFSx1.RotateToEnd:SlingFSx2.RotateToEnd
End If
End Sub

Sub RightSlingshot_Slingshot:vpmTimer.PulseSw 43:RightSling.Visible=1:DxEmKickerT1.TransX=-25:RStep=0:Me.TimerEnabled=1:PlaySoundAtVol SoundFX("Slingshot",DOFContactors), DxEmKickerT1, SliVol:End Sub
Sub RightSlingshot_Timer
  Select Case RStep
    Case 0:RightSling.Visible = 1
    Case 1: 'pause
    Case 2:RightSling.Visible = 0 :RightSling1.Visible = 1:DxEmKickerT1.TransX=-21
    Case 3:RightSling1.Visible = 0:RightSling2.Visible = 1:DxEmKickerT1.TransX=-17:SlingFDx1.RotateToStart:SlingFDx2.RotateToStart
    Case 4:RightSling2.Visible = 0:Me.TimerEnabled = 0:DxEmKickerT1.TransX=0
  End Select
  RStep = RStep + 1
End Sub

Sub SlingFDx1_Collide(parm)
 If DxEmKickerT1.TransX=0 And PinPlay=1 And parm > ThSling Then
  vpmTimer.PulseSw 43:PlaySoundAtVol SoundFX("Slingshot",DOFContactors), DxEmKickerT1, SliVol:RStep=0
  RightSling.Visible=1:DxEmKickerT1.TransX=-25:RightSlingshot.TimerEnabled=1:SlingFDx1.RotateToEnd:SlingFDx2.RotateToEnd
End If
End Sub
Sub SlingFDx2_Collide(parm)
 If DxEmKickerT1.TransX=0 And PinPlay=1 And parm > ThSling Then
  vpmTimer.PulseSw 43:PlaySoundAtVol SoundFX("Slingshot",DOFContactors), DxEmKickerT1, SliVol:RStep=0
  RightSling.Visible=1:DxEmKickerT1.TransX=-25:RightSlingshot.TimerEnabled=1:SlingFDx1.RotateToEnd:SlingFDx2.RotateToEnd
End If
End Sub

' Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 50:PlaySoundAtVol SoundFX("jet1",DOFContactors), Bumper1, BumVol:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 51:PlaySoundAtVol SoundFX("jet1",DOFContactors), Bumper2, BumVol:End Sub

' Spinner
Sub sw25_Spin:vpmTimer.PulseSw 25:PlaySoundAtVol "spinner", sw25, SpiVol:End Sub

' Eject holes
Sub DrainStart_Hit:Me.TimerEnabled=1:PlaysoundAtVol "drain1a", DrainStart, KidVol:End Sub
Sub DrainStart_Timer:Me.TimerEnabled=0:DrainStart.Kick 90, 1:End Sub
Sub Drain_Hit:bsTrough.AddBall Me:End Sub
Sub sw26_Hit:PlaysoundAtVol "fx_kicker_catch", sw26, KicVol:bsLK.AddBall 0:End Sub
Sub sw36_Hit:PlaysoundAtVol "fx_kicker_catch", sw36, KicVol:bsRK.AddBall 0:End Sub

' Ball release sound and DOF call
Sub BallRelTrig_Hit:PlaySoundAtVol SoundFX("popper",DOFContactors), Drain, BreVol:End Sub

' Rollovers
Sub sw2_Hit:  Controller.Switch(2) = 1:PlaySoundAtVol "sensor", sw2, SwiVol:End Sub
Sub sw2_UnHit:Controller.Switch(2) = 0:End Sub
Sub sw3_Hit:  Controller.Switch(3) = 1:PlaySoundAtVol "sensor", sw3, SwiVol:End Sub
Sub sw3_UnHit:Controller.Switch(3) = 0:End Sub
Sub sw5_Hit:  Controller.Switch(5) = 1:PlaySoundAtVol "sensor", sw5, SwiVol:End Sub
Sub sw5_UnHit:Controller.Switch(5) = 0:End Sub
Sub sw12_Hit:  Controller.Switch(12) = 1:PlaySoundAtVol "sensor", sw12, SwiVol:End Sub
Sub sw12_UnHit:Controller.Switch(12) = 0:End Sub
Sub sw13_Hit:  Controller.Switch(13) = 1:PlaySoundAtVol "sensor", sw13, SwiVol:End Sub
Sub sw13_UnHit:Controller.Switch(13) = 0:End Sub
Sub sw15_Hit:  Controller.Switch(15) = 1:PlaySoundAtVol "sensor", sw15, SwiVol:End Sub
Sub sw15_UnHit:Controller.Switch(15) = 0:End Sub
Sub sw22_Hit:  Controller.Switch(22) = 1:PlaySoundAtVol "sensor", sw22, SwiVol:End Sub
Sub sw22_UnHit:Controller.Switch(22) = 0:End Sub
Sub sw23_Hit:  Controller.Switch(23) = 1:PlaySoundAtVol "sensor", sw23, SwiVol:End Sub
Sub sw23_UnHit:Controller.Switch(23) = 0:End Sub

Sub sw32_Hit:  Controller.Switch(32) = 1:PlaySoundAtVol "sensor", sw32, SwiVol:End Sub
Sub sw32_UnHit:Controller.Switch(32) = 0:End Sub
Sub sw33_Hit:  Controller.Switch(33) = 1:PlaySoundAtVol "sensor", sw33, SwiVol:End Sub
Sub sw33_UnHit:Controller.Switch(33) = 0:End Sub
Sub sw40_Hit:  Controller.Switch(40) = 1:PlaySoundAtVol "sensor", sw40, SwiVol:End Sub
Sub sw40_UnHit:Controller.Switch(40) = 0:End Sub
Sub sw41_Hit:  Controller.Switch(41) = 1:PlaySoundAtVol "sensor", sw41, SwiVol:End Sub
Sub sw41_UnHit:Controller.Switch(41) = 0:End Sub

' Targets
Sub sw0_Hit:vpmTimer.PulseSw 0:PlaySoundAtBallVol SoundFX("target",DOFTargets), TarVol:End Sub
Sub sw1_Hit:vpmTimer.PulseSw 1:PlaySoundAtBallVol SoundFX("target",DOFTargets), TarVol:End Sub
Sub sw10_Hit:vpmTimer.PulseSw 10:PlaySoundAtBallVol SoundFX("target",DOFTargets), TarVol:End Sub
Sub sw11_Hit:vpmTimer.PulseSw 11:PlaySoundAtBallVol SoundFX("target",DOFTargets), TarVol:End Sub
Sub sw20_Hit:vpmTimer.PulseSw 20:PlaySoundAtBallVol SoundFX("target",DOFTargets), TarVol:End Sub
Sub sw21_Hit:vpmTimer.PulseSw 21:PlaySoundAtBallVol SoundFX("target",DOFTargets), TarVol:End Sub
Sub sw30_Hit:vpmTimer.PulseSw 30:PlaySoundAtBallVol SoundFX("target",DOFTargets), TarVol:End Sub
Sub sw31_Hit:vpmTimer.PulseSw 31:PlaySoundAtBallVol SoundFX("target",DOFTargets), TarVol:End Sub

' Gates
Sub Gate1_Hit:PlaySoundAtBallVol "Gate5", GatVol:End Sub
Sub Gate2_Hit:PlaySoundAtBallVol "Gate5", GatVol:End Sub
Sub Gate3_Hit:PlaySoundAtBallVol "Gate5", GatVol:End Sub

'*********
'Solenoids
'*********

SolCallback(1) = "setlamp 131,"
SolCallback(2) = "bsLKBallRelease"
SolCallback(3) = "setlamp 133,"
SolCallback(4) = "setlamp 134,"
SolCallback(5) = "bsRKBallRelease"
SolCallback(6) = "bsTrough.SolOut"
SolCallback(7) = "setlamp 137,"
SolCallback(8) = "KnockerSol"
SolCallback(9) = "bsTrough.SolIn"
SolCallback(10) = "SolRun"

'Sol1 Top Left Domes
'Sol2 Top Left Hole
'Sol3 Bottom Left Dome
'Sol4 Bottom Right Dome
'Sol5 Top Right Hole
'Sol6 Ball Release
'Sol7 Top Right Side Domes
'Sol8 Knocker
'Sol9 Outhole

Sub bsLKBallRelease(Enabled)
 If Enabled Then
  bsLK.ExitSol_On
  EjectTimerSx.Enabled=1
  EjectArmSx.RotX = 20
  PlaySoundAtVol SoundFX("popper",DOFContactors), sw26, KieVol
End If
End Sub

Sub EjectTimerSx_Timer
  EjectArmSx.RotX = EjectArmSx.RotX -2
 If EjectArmSx.RotX = 0 Then:EjectTimerSx.Enabled = 0
End Sub

Sub bsRKBallRelease(Enabled)
 If Enabled Then
  bsRK.ExitSol_On
  EjectTimerDx.Enabled=1
  EjectArmDx.RotX = 20
  PlaySoundAtVol SoundFX("popper",DOFContactors), sw36, KieVol
End If
End Sub

Sub EjectTimerDx_Timer
  EjectArmDx.RotX = EjectArmDx.RotX -2
 If EjectArmDx.RotX = 0 Then:EjectTimerDx.Enabled = 0
End Sub

Sub KnockerSol(Enabled)
 If Enabled Then
  PlaySoundAtVol SoundFX("Knocker",DOFKnocker), gi2a, KnoVol
End If
End Sub

Sub SolRun(Enabled)
  vpmNudge.SolGameOn Enabled
 If Enabled Then
  PinPlay=1
Else
  PinPlay=0
  SlingFSx1.RotateToStart:SlingFSx2.RotateToStart
  SlingFSx1.RotateToStart:SlingFSx2.RotateToStart
End If
End Sub

'**************
' Flipper Subs
'**************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
' Controller.Switch(52) = ABS(enabled)
  If Enabled Then
    PlaySoundAtVol SoundFX("flipperup1",DOFFlippers), LeftFlipper, FluVol:LeftFlipper.RotateToEnd
  Else
        PlaySoundAtVol SoundFX("flipperdown1",DOFFlippers), LeftFlipper, FldVol:LeftFlipper.RotateToStart
  End If
End Sub

Sub SolRFlipper(Enabled)
' Controller.Switch(53) = ABS(enabled)
  If Enabled Then
    PlaySoundAtVol SoundFX("flipperup1",DOFFlippers), RightFlipper, FluVol:RightFlipper.RotateToEnd
  Else
        PlaySoundAtVol SoundFX("flipperdown1",DOFFlippers), RightFlipper, FldVol:RightFlipper.RotateToStart
  End If
End Sub

Sub LeftFlipper_Collide(parm)
  PlaySound "rubber_flipper", 0, parm / 10, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper_Collide(parm)
  PlaySound "rubber_flipper", 0, parm / 10, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

'******************
' RealTime Updates
'******************
Set MotorCallback = GetRef("GameTimer")

Sub GameTimer
  UpdateFlipperLogos
End Sub

Dim PI
PI = Round(4 * Atn(1), 6) '3.1415926535897932384626433832795

Sub UpdateFlipperLogos
  FlipperSx.RotZ = LeftFlipper.CurrentAngle
  FlipperDx.RotZ = RightFlipper.CurrentAngle
  pSpinnerRod.TransX = sin( (sw25.CurrentAngle+180) * (PI/180)) * 8
  pSpinnerRod.TransZ = cos( (sw25.CurrentAngle+180) * (PI/180)) * 8
  pSpinnerRod.RotY = sin( (sw25.CurrentAngle-180) * (PI/180)) * 6
  WireGate1.RotX = Gate1.CurrentAngle * 1.4
  WireGate2.RotX = Gate2.CurrentAngle
  WireGate3.RotX = Gate3.CurrentAngle + 26
  PlangerRod.TransY = (Plunger.Position * Plunger.Stroke/25) - (Plunger.Stroke/(1/Plunger.ParkPosition))
  Pomello.TransY = (Plunger.Position * Plunger.Stroke/25) - (Plunger.Stroke/(1/Plunger.ParkPosition))
  Molla.TransY = (Plunger.Position * Plunger.Stroke/25) - (Plunger.Stroke/(1/Plunger.ParkPosition))
 If Plunger.Position < (Plunger.ParkPosition * 25) Then
  Molla.Size_Y = Plunger.Position * (4/Plunger.ParkPosition) + ((Plunger.ParkPosition-Plunger.Position/25) * (42/(Plunger.Stroke*Plunger.ParkPosition)*100))
Else
  Molla.Size_Y = 100
End If
End Sub

'**********************************
'       JP's VP10 Fading Lamps & Flashers v2
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
    UpdateLeds
End Sub

Sub UpdateLamps

' NFadeLm 0, gio1
' NFadeLm 0, gio2
' NFadeLm 0, gio3
' NFadeL 0, gio4

  NFadeLim 1, gi1
  NFadeLim 1, gi1a
  NFadeLim 1, gi2
  NFadeLim 1, gi2a
  NFadeLim 1, gi3
  NFadeLim 1, gi4
  NFadeLim 1, gi5
  NFadeLim 1, gi5a
  NFadeLim 1, gi6
  NFadeLim 1, gi6a
  NFadeLim 1, gi7
  NFadeLim 1, gi7a
  NFadeLim 1, gi8
  NFadeLim 1, gi8a
  NFadeLim 1, gi9
  NFadeLim 1, gi9a
  NFadeLim 1, gi10
  NFadeLim 1, gi10a
  NFadeLim 1, gi11
  NFadeLim 1, gi11a
  NFadeLim 1, gi12
  NFadeLim 1, gi12a
  NFadeLim 1, gi13
  NFadeLim 1, gi14
  NFadeLim 1, gi14a
  NFadeLim 1, gi15
  NFadeLim 1, gi16
  NFadeLim 1, gi17
  NFadeLim 1, gi18
  NFadeLim 1, gi19
  NFadeLim 1, gi20
' NFadeLim 1, gi21
  NFadeLim 1, gi22

  FlashDLIm 1, 1, CDSlotS1
  FlashDLIm 1, 1, CDSlotC1
  FlashDLI 1, 1, CDSlotD1

  NFadeL 2, l2
  NFadeL 3, l3

  NFadeL 5, l5
  NFadeL 6, l6
  NFadeL 7, l7
  NFadeL 8, l8
  NFadeL 9, l9
  NFadeL 10, l10
  NFadeL 11, l11

  NFadeLm 12, l12
  NFadeLm 12, l12a
  Flash 12, fl12

  NFadeLm 13, l13
  NFadeLm 13, l13a
  Flash 13, fl13

  NFadeLm 14, l14
  NFadeLm 14, l14a
  Flash 14, fl14

  NFadeLm 15, l15
  NFadeLm 15, l15a
  Flash 15, fl15

  NFadeLm 16, l16
  NFadeLm 16, l16a
  Flash 16, fl16

  NFadeLm 17, l17
  NFadeLm 17, l17a
  Flash 17, fl17

  NFadeLm 18, l18
  Flash 18, fl18

  NFadeLm 19, l19
  Flash 19, fl19

  NFadeLm 20, l20
  Flash 20, fl20

  NFadeLm 21, l21
  Flash 21, fl21

  NFadeLm 22, l22
  Flash 22, fl22

  NFadeLm 23, l23
  Flash 23, fl23

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

  NFadeLm 131, f131
  NFadeLm 131, f131a
  Flashm 131, f131b
  Flash 131, f131c

  NFadeLm 133, f133
  NFadeLm 133, f133a
  NFadeLm 133, f133b
  Flash 133, f133c

  NFadeLm 134, f134
  NFadeLm 134, f134a
  NFadeLm 134, f134b
  Flash 134, f134c

  NFadeLm 137, f137
  NFadeLm 137, f137a
  Flashm 137, f137b
  Flash 137, f137c

End Sub

' div lamp subs

Sub InitLamps()
    Dim x
    For x = 0 to 200
        LampState(x) = 0        ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4      ' used to track the fading state
        FlashSpeedUp(x) = 0.2   ' faster speed when turning on the flasher
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

Sub SetModLamp(nr, level)
  FlashLevel(nr) = level /150 'lights & flashers
End Sub

' Lights: old method, using 4 images

Sub FadeL(nr, light, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 4:light.image = b:FadingLevel(nr) = 6                   'fading to off...
        Case 5:light.image = a:light.State = 1:FadingLevel(nr) = 1   'ON
        Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1           'wait
        Case 9:light.image = c:FadingLevel(nr) = FadingLevel(nr) + 1 'fading...
        Case 10, 11, 12:FadingLevel(nr) = FadingLevel(nr) + 1        'wait
        Case 13:light.image = d:Light.State = 0:FadingLevel(nr) = 0  'Off
    End Select
End Sub

Sub FadeLm(nr, light, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 4:light.image = b
        Case 5:light.image = a
        Case 9:light.image = c
        Case 13:light.image = d
    End Select
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

Sub LightMod(nr, object) ' modulated lights used as flashers
    Object.IntensityScale = FlashLevel(nr)
  Object.State = 1
End Sub

' Inverted

Sub NFadeLi(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.state = 1:FadingLevel(nr) = 1
        Case 5:object.state = 0:FadingLevel(nr) = 0
    End Select
End Sub

Sub NFadeLim(nr, object) ' used for multiple lights
    Select Case FadingLevel(nr)
        Case 4:object.state = 1
        Case 5:object.state = 0
    End Select
End Sub

'Ramps & Primitives used as 4 step fading lights
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

Sub FlashMod(nr, object) 'sets the flashlevel from the SolModCallback
    Object.IntensityScale = FlashLevel(nr)
End Sub

Sub FastFlash(nr, object)
    Select Case FadingLevel(nr)
    Case 4:object.Visible = 0:FadingLevel(nr) = 0 'off
    Case 5:object.Visible = 1:FadingLevel(nr) = 1 'on
    Object.IntensityScale = FlashMax(nr)
    End Select
End Sub

Sub FastFlashm(nr, object)
    Select Case FadingLevel(nr)
    Case 4:object.Visible = 0 'off
    Case 5:object.Visible = 1 'on
    Object.IntensityScale = FlashMax(nr)
    End Select
End Sub

' Objects DisableLighting Inverted

Sub FlashDLI(nr, Limite, object)
    Select Case FadingLevel(nr)
        Case 4 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
            If FlashLevel(nr) > FlashMax(nr) Then
                FlashLevel(nr) = FlashMax(nr)
                FadingLevel(nr) = 1 'completely on
            End if
            Object.BlendDisableLighting = Limite*FlashLevel(nr)
        Case 5 'off
            FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown(nr)
            If FlashLevel(nr) < FlashMin(nr) Then
                FlashLevel(nr) = FlashMin(nr)
                FadingLevel(nr) = 0 'completely off
            End if
            Object.BlendDisableLighting = Limite*FlashLevel(nr)
    End Select
End Sub

Sub FlashDLIm(nr, Limite, object) 'multiple objects, it just sets the flashlevel
    Object.BlendDisableLighting = Limite*FlashLevel(nr)
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

' *********************************************************************
'           Wall, rubber and metal hit sounds
' *********************************************************************

Sub Rubbers_Hit(idx):PlaySoundAtBallVol "rubber1", RubVol:End Sub

' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / Table1.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10))
    End If
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = (SQR((ball.VelX ^2) + (ball.VelY ^2)))
End Function

Function AudioFade(ball) 'only on VPX 10.4 and newer
    Dim tmp
    tmp = ball.y * 2 / Table1.height-1
    If tmp > 0 Then
        AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10))
    End If
End Function

'*****************************************
'PlaySound(string, int loopcount, float volume, float pan, float randompitch, int pitch, bool useexisting, bool restart, float front_rear_fade)

Sub PlaySoundAt(soundname, tableobj) 'play sound at X and Y position of an object, mostly bumpers, flippers and other fast objects
    PlaySound soundname, 0, 1, Pan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtVol(soundname, tableobj, Vol) 'play sound at X and Y position of an object, mostly bumpers, flippers and other fast objects
    PlaySound soundname, 0, Vol, Pan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname) ' play a sound at the ball position, like rubbers, targets
    PlaySound soundname, 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtBallVol(soundname, VolMult) ' play a sound at the ball position, like rubbers, targets
    PlaySound soundname, 0, Vol(ActiveBall) * VolMult, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 19 ' total number of balls
Const lob = 0  'number of locked balls
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub RollingTimer_Timer()
    Dim BOT, b, ballpitch, ballvol
    BOT = GetBalls

    ' stop the sound of deleted balls
    For b = UBound(BOT) + 1 to tnob
        rolling(b) = False
    aBallShadow(b).Visible = 0
        StopSound("fx_ballrolling" & b)
    Next

    ' exit the sub if no balls on the table
    If UBound(BOT) = lob - 1 Then Exit Sub 'there no extra balls on this table

    ' play the rolling sound for each ball
    For b = lob to UBound(BOT)

    aBallShadow(b).X = BOT(b).X
    aBallShadow(b).Y = BOT(b).Y
    aBallShadow(b).Height = BOT(b).Z - (BallSize / 2)+1

        If BallVel(BOT(b)) > 1 Then
            If BOT(b).z < 30 Then
                ballpitch = Pitch(BOT(b))
                ballvol = Vol(BOT(b))
            Else
                ballpitch = Pitch(BOT(b)) + 20000 'increase the pitch on a ramp
                ballvol = Vol(BOT(b)) * 4
            End If
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, ballvol*RolVol, Pan(BOT(b)), 0, ballpitch, 1, 0, AudioFade(BOT(b))
        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If

    ' Ball Shadow
    If BallSHW = 1 Then
      aBallShadow(b).Visible = 1
    Else
      aBallShadow(b).Visible = 0
    End If

    Next
End Sub

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
    PlaySound("fx_collide"), 0, Csng(ColVol*((velocity) ^2 / 200)), Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

'*****************
' Leds Display
'*****************

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

Sub UpdateLeds
  Dim ChgLED, ii, jj, num, chg, stat, obj
    ChgLED=Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
      If UseFlexDMD Then LockRender
'     If Not FlexDMD is Nothing Then FlexDMD.LockRenderThread
      For ii=0 To UBound(chgLED)
        num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)
        If UseFlexDMD Then UpdateFlexChar num, stat

        If DotMatrix.TimerEnabled = 0 Then
          For Each obj In Digits(num)
            If chg And 1 Then obj.State=stat And 1
            chg=chg\2 : stat=stat\2
          Next
          End If

      Next
      If UseFlexDMD Then UnlockRender
'     If Not FlexDMD is Nothing Then FlexDMD.UnlockRenderThread
'   End If
'   End If

  Else

  End If
End Sub

Sub LockRender
 If Not FlexDMD is Nothing Then FlexDMD.LockRenderThread
End Sub

Sub UnlockRender
 If Not FlexDMD is Nothing Then FlexDMD.UnlockRenderThread
End Sub

'Night Moves
'DIP switches by Inkochnito
'Add coins chute by Mike da Spike

Sub editDips
  Dim vpmDips : Set vpmDips = New cvpmDips
  With vpmDips
        .AddForm 700,400,"Night Moves - DIP switches"

        .AddFrame 2,4,190,"Left Coin Chute (Coins/Credit)",&H0000001F,Array("4/1",&H0000000D,"2/1",&H0000000A,"1/1",&H00000000,"1/2",&H00000010) 'Dip 1-5
        .AddFrame 2,80,190,"Right Coin Chute (Coins/Credit)",&H00001F00,Array("4/1",&H00000D00,"2/1",&H00000A00,"1/1",&H00000000,"1/2",&H00001000) 'Dip 9-13
        .AddFrame 2,160,190,"Center Coin Chute (Coins/Credit)",&H001F0000,Array("4/1",&H000D0000,"2/1",&H000A0000,"1/1",&H00000000,"1/2",&H00010000) 'Dip 17-21
        .AddFrame 2,240,190,"3rd coin chute credits control",&H20000000,Array("no effect",0,"add 9",&H20000000)'dip 30

        .AddFrame 207,4,190,"Maximum credits",49152,Array("8 credits",0,"10 credits",32768,"15 credits",&H00004000,"20 credits",49152)'dip 15&16
        .AddFrame 207,80,190,"Coin chute left and right control",&H00002000,Array("seperate",0,"same",&H00002000)'dip 14
        .AddFrame 207,126,190,"Playfield special",&H00200000,Array("replay",0,"extra ball",&H00200000)'dip 22
        .AddFrame 207,172,190,"High games to date control",&H00000020,Array("no effect",0,"reset high games 2-5 on power off",&H00000020)'dip 6
        .AddFrame 207,218,190,"Attract mode sound",&H00000040,Array("off",0,"on",&H00000040)'dip 7
        .AddFrame 207,264,190,"Auto Percentage control",&H00000080,Array("disabled",0,"enabled",&H00000080)'dip 8

        .AddFrame 412,4,190,"High game to date awards",&H00C00000,Array("not displayed and no award",0,"displayed and no award",&H00800000,"displayed and 2 replays",&H00400000,"displayed and 3 replays",&H00C00000)'dip 23&24
        .AddFrame 412,80,190,"Balls per game",&H01000000,Array("5 balls",0,"3 balls",&H01000000)'dip 25
        .AddFrame 412,126,190,"Replay limit",&H04000000,Array("no limit",0,"one per game",&H04000000)'dip 27
        .AddFrame 412,172,190,"Novelty",&H08000000,Array("normal",0,"extra ball and replay scores 500K",&H08000000)'dip 28
        .AddFrame 412,218,190,"Game mode",&H10000000,Array("replay",0,"extra ball",&H10000000)'dip 29
        .AddFrame 412,264,190,"Extra Ball frequency 3 ball/5 ball",&HC0000000,Array("6/10",0,"8/13",&H80000000,"9/15",&H40000000,"10/17",&HC0000000)'dip 32

        .AddChk 2,360,120,Array("Match feature",&H02000000)'dip 26
        .AddLabel 50,380,300,20,"After hitting OK, press F3 to reset game with new settings."
    End With
    Dim extra
    extra = Controller.Dip(4) + Controller.Dip(5)*256
    extra = vpmDips.ViewDipsExtra(extra)
    Controller.Dip(4) = extra And 255
    Controller.Dip(5) = (extra And 65280)\256 And 255
End Sub

Set vpmShowDips=GetRef("editDips")

'*****************
' FlexDMD
'*****************

'flexdmd image constants
Const DMD_A = "VPX.DMD_A"
Const DMD_B = "VPX.DMD_B"
Const DMD_C = "VPX.DMD_C"
Const DMD_D = "VPX.DMD_D"
Const DMD_E = "VPX.DMD_E"
Const DMD_F = "VPX.DMD_F"
Const DMD_G = "VPX.DMD_G"
Const DMD_H = "VPX.DMD_H"
Const DMD_I = "VPX.DMD_I"
Const DMD_J = "VPX.DMD_J"
Const DMD_K = "VPX.DMD_K"
Const DMD_L = "VPX.DMD_L"
Const DMD_M = "VPX.DMD_M"
Const DMD_N = "VPX.DMD_N"
Const DMD_O = "VPX.DMD_O"
Const DMD_P = "VPX.DMD_P"
Const DMD_Q = "VPX.DMD_Q"
Const DMD_R = "VPX.DMD_R"
Const DMD_S = "VPX.DMD_S"
Const DMD_T = "VPX.DMD_T"
Const DMD_U = "VPX.DMD_U"
Const DMD_V = "VPX.DMD_V"
Const DMD_W = "VPX.DMD_W"
Const DMD_X = "VPX.DMD_X"
Const DMD_Y = "VPX.DMD_Y"
Const DMD_Z = "VPX.DMD_Z"

Const DMD_1 = "VPX.DMD_1"
Const DMD_2 = "VPX.DMD_2"
Const DMD_3 = "VPX.DMD_3"
Const DMD_4 = "VPX.DMD_4"
Const DMD_6 = "VPX.DMD_6"
Const DMD_7 = "VPX.DMD_7"
Const DMD_8 = "VPX.DMD_8"
Const DMD_9 = "VPX.DMD_9"

Const DMD_1dot = "VPX.DMD_1dot"
Const DMD_2dot = "VPX.DMD_2dot"
Const DMD_3dot = "VPX.DMD_3dot"
Const DMD_4dot = "VPX.DMD_4dot"
Const DMD_6dot = "VPX.DMD_6dot"
Const DMD_7dot = "VPX.DMD_7dot"
Const DMD_8dot = "VPX.DMD_8dot"
Const DMD_9dot = "VPX.DMD_9dot"

Const DMD_Odot = "VPX.DMD_Odot"
Const DMD_Sdot = "VPX.DMD_Sdot"

Const DMD_Space = "VPX.DMD_Space"
Const DMD_SpaceDot = "VPX.DMD_SpaceDot"

Const DMD_Ampersand = "VPX.DMD_Ampersand"
Const DMD_Asterick = "VPX.DMD_Asterick"
Const DMD_BSlash = "VPX.DMD_BSlash"
Const DMD_CloseBracket = "VPX.DMD_CloseBracket"
Const DMD_Colon = "VPX.DMD_Colon"
Const DMD_Dollar = "VPX.DMD_Dollar"
Const DMD_Equals = "VPX.DMD_Equals"
Const DMD_Exclamation = "VPX.DMD_Exclamation"
Const DMD_FSlash = "VPX.DMD_FSlash"
Const DMD_GreaterThan = "VPX.DMD_GreaterThan"
Const DMD_Hash = "VPX.DMD_Hash"
Const DMD_LessThan = "VPX.DMD_LessThan"
Const DMD_Minus = "VPX.DMD_Minus"
Const DMD_OpenBracket = "VPX.DMD_OpenBracket"
Const DMD_Percent = "VPX.DMD_Percent"
Const DMD_Plus = "VPX.DMD_Plus"
Const DMD_Question = "VPX.DMD_Question"
Const DMD_Quote = "VPX.DMD_Quote"
Const DMD_SemiColon = "VPX.DMD_SemiColon"
Const DMD_SingleQuote = "VPX.DMD_SingleQuote"

Dim FlexDMD
Dim FlexDMDDict
Dim FlexDMDScene

Sub FlexDMD_Init() 'default/startup values

  ' flex dmd variables
  Dim FlexDMDFont
  Dim FlexPath

  ' populate the lookup dictionary for mapping display characters
  FlexDictionary_Init

  'setup flex dmd
  Set FlexDMD = CreateObject("FlexDMD.FlexDMD")
  If Not FlexDMD is Nothing Then


    FlexDMD.GameName = cGameName
    FlexDMD.TableFile = Table1.Filename & ".vpx"
    FlexDMD.RenderMode = 0
    FlexDMD.Color = RGB(50, 188, 255)
    DotMatrix.Color = RGB(0, 128, 255)
    FlexDMD.Width = 128
    FlexDMD.Height = 32
    FlexDMD.Clear = True
    FlexDMD.Run = True

    FlexDMD.LockRenderThread

    Set FlexDMDScene = FlexDMD.NewGroup("Scene")


    With FlexDMDScene
      'populate blank display
'     .AddActor FlexDMD.NewImage("Back", "VPX.DMD_Background")
      '40 segment display holders
      .AddActor FlexDMD.NewImage("Seg0", DMD_Space)
      .GetImage("Seg0").SetAlignedPosition 4,0,0
      .AddActor FlexDMD.NewImage("Seg1", DMD_Space)
      .GetImage("Seg1").SetAlignedPosition 10,0,0
      .AddActor FlexDMD.NewImage("Seg2", DMD_Space)
      .GetImage("Seg2").SetAlignedPosition 16,0,0
      .AddActor FlexDMD.NewImage("Seg3", DMD_Space)
      .GetImage("Seg3").SetAlignedPosition 22,0,0
      .AddActor FlexDMD.NewImage("Seg4", DMD_Space)
      .GetImage("Seg4").SetAlignedPosition 28,0,0
      .AddActor FlexDMD.NewImage("Seg5", DMD_Space)
      .GetImage("Seg5").SetAlignedPosition 34,0,0
      .AddActor FlexDMD.NewImage("Seg6", DMD_Space)
      .GetImage("Seg6").SetAlignedPosition 40,0,0
      .AddActor FlexDMD.NewImage("Seg7", DMD_Space)
      .GetImage("Seg7").SetAlignedPosition 46,0,0
      .AddActor FlexDMD.NewImage("Seg8", DMD_Space)
      .GetImage("Seg8").SetAlignedPosition 52,0,0
      .AddActor FlexDMD.NewImage("Seg9", DMD_Space)
      .GetImage("Seg9").SetAlignedPosition 58,0,0
      .AddActor FlexDMD.NewImage("Seg10", DMD_Space)
      .GetImage("Seg10").SetAlignedPosition 64,0,0
      .AddActor FlexDMD.NewImage("Seg11", DMD_Space)
      .GetImage("Seg11").SetAlignedPosition 70,0,0
      .AddActor FlexDMD.NewImage("Seg12", DMD_Space)
      .GetImage("Seg12").SetAlignedPosition 76,0,0
      .AddActor FlexDMD.NewImage("Seg13", DMD_Space)
      .GetImage("Seg13").SetAlignedPosition 82,0,0
      .AddActor FlexDMD.NewImage("Seg14", DMD_Space)
      .GetImage("Seg14").SetAlignedPosition 88,0,0
      .AddActor FlexDMD.NewImage("Seg15", DMD_Space)
      .GetImage("Seg15").SetAlignedPosition 94,0,0
      .AddActor FlexDMD.NewImage("Seg16", DMD_Space)
      .GetImage("Seg16").SetAlignedPosition 100,0,0
      .AddActor FlexDMD.NewImage("Seg17", DMD_Space)
      .GetImage("Seg17").SetAlignedPosition 106,0,0
      .AddActor FlexDMD.NewImage("Seg18", DMD_Space)
      .GetImage("Seg18").SetAlignedPosition 112,0,0
      .AddActor FlexDMD.NewImage("Seg19", DMD_Space)
      .GetImage("Seg19").SetAlignedPosition 118,0,0

      .AddActor FlexDMD.NewImage("Seg20", DMD_Space)
      .GetImage("Seg20").SetAlignedPosition 4,16,0
      .AddActor FlexDMD.NewImage("Seg21", DMD_Space)
      .GetImage("Seg21").SetAlignedPosition 10,16,0
      .AddActor FlexDMD.NewImage("Seg22", DMD_Space)
      .GetImage("Seg22").SetAlignedPosition 16,16,0
      .AddActor FlexDMD.NewImage("Seg23", DMD_Space)
      .GetImage("Seg23").SetAlignedPosition 22,16,0
      .AddActor FlexDMD.NewImage("Seg24", DMD_Space)
      .GetImage("Seg24").SetAlignedPosition 28,16,0
      .AddActor FlexDMD.NewImage("Seg25", DMD_Space)
      .GetImage("Seg25").SetAlignedPosition 34,16,0
      .AddActor FlexDMD.NewImage("Seg26", DMD_Space)
      .GetImage("Seg26").SetAlignedPosition 40,16,0
      .AddActor FlexDMD.NewImage("Seg27", DMD_Space)
      .GetImage("Seg27").SetAlignedPosition 46,16,0
      .AddActor FlexDMD.NewImage("Seg28", DMD_Space)
      .GetImage("Seg28").SetAlignedPosition 52,16,0
      .AddActor FlexDMD.NewImage("Seg29", DMD_Space)
      .GetImage("Seg29").SetAlignedPosition 58,16,0
      .AddActor FlexDMD.NewImage("Seg30", DMD_Space)
      .GetImage("Seg30").SetAlignedPosition 64,16,0
      .AddActor FlexDMD.NewImage("Seg31", DMD_Space)
      .GetImage("Seg31").SetAlignedPosition 70,16,0
      .AddActor FlexDMD.NewImage("Seg32", DMD_Space)
      .GetImage("Seg32").SetAlignedPosition 76,16,0
      .AddActor FlexDMD.NewImage("Seg33", DMD_Space)
      .GetImage("Seg33").SetAlignedPosition 82,16,0
      .AddActor FlexDMD.NewImage("Seg34", DMD_Space)
      .GetImage("Seg34").SetAlignedPosition 88,16,0
      .AddActor FlexDMD.NewImage("Seg35", DMD_Space)
      .GetImage("Seg35").SetAlignedPosition 94,16,0
      .AddActor FlexDMD.NewImage("Seg36", DMD_Space)
      .GetImage("Seg36").SetAlignedPosition 100,16,0
      .AddActor FlexDMD.NewImage("Seg37", DMD_Space)
      .GetImage("Seg37").SetAlignedPosition 106,16,0
      .AddActor FlexDMD.NewImage("Seg38", DMD_Space)
      .GetImage("Seg38").SetAlignedPosition 112,16,0
      .AddActor FlexDMD.NewImage("Seg39", DMD_Space)
      .GetImage("Seg39").SetAlignedPosition 118,16,0

    End With

    FlexDMD.Stage.AddActor FlexDMDScene

    FlexDMD.Show = FlexDMDShow
    FlexDMD.UnlockRenderThread

  End If

End Sub

'*****************
' Flasher DMD
'*****************

Sub DotMatrix_Timer
  Dim DMDp
  If FlexDMD.RenderMode = 0 Then
    DMDp = FlexDMD.DmdColoredPixels
    If Not IsEmpty(DMDp) Then
      DMDWidth = FlexDMD.Width
      DMDHeight = FlexDMD.Height
      DMDColoredPixels = DMDp
    End If
  Else
    DMDp = FlexDMD.DmdPixels
    If Not IsEmpty(DMDp) Then
      DMDWidth = FlexDMD.Width
      DMDHeight = FlexDMD.Height
      DMDPixels = DMDp
    End If
  End If
End Sub

'*****************
' FlexDictionary
'*****************

Sub FlexDictionary_Init

  Set FlexDMDDict = CreateObject("Scripting.Dictionary")

  FlexDMDDict.Add 0, DMD_Space
  FlexDMDDict.Add 63, DMD_O
  FlexDMDDict.Add 8704, DMD_1
  FlexDMDDict.Add 2139, DMD_2
  FlexDMDDict.Add 2127, DMD_3
  FlexDMDDict.Add 2150, DMD_4
  FlexDMDDict.Add 2157, DMD_S
  FlexDMDDict.Add 2173, DMD_6
  FlexDMDDict.Add 7, DMD_7
  FlexDMDDict.Add 2175,DMD_8
  FlexDMDDict.Add 2159,DMD_9

  FlexDMDDict.Add 191,DMD_Odot
  FlexDMDDict.Add 8832, DMD_1dot
  FlexDMDDict.Add 2267, DMD_2dot
  FlexDMDDict.Add 2255, DMD_3dot
  FlexDMDDict.Add 2278, DMD_4dot
  FlexDMDDict.Add 2285, DMD_Sdot
  FlexDMDDict.Add 2301, DMD_6dot
  FlexDMDDict.Add 135, DMD_7dot
  FlexDMDDict.Add 2303, DMD_8dot
  FlexDMDDict.Add 2287, DMD_9dot

  FlexDMDDict.Add 2167, DMD_A
  FlexDMDDict.Add 10767, DMD_B
  FlexDMDDict.Add 57, DMD_C
  FlexDMDDict.Add 8719, DMD_D
  FlexDMDDict.Add 121, DMD_E
  FlexDMDDict.Add 113, DMD_F
  FlexDMDDict.Add 2109, DMD_G
  FlexDMDDict.Add 2166, DMD_H
  FlexDMDDict.Add 8713, DMD_I
  FlexDMDDict.Add 30, DMD_J
  FlexDMDDict.Add 5232, DMD_K
  FlexDMDDict.Add 56, DMD_L
  FlexDMDDict.Add 1334, DMD_M
  FlexDMDDict.Add 4406, DMD_N
  ' "O" = 0
  FlexDMDDict.Add 2163, DMD_P
  FlexDMDDict.Add 4159, DMD_Q
  FlexDMDDict.Add 6259, DMD_R
  ' "S" = 5
  FlexDMDDict.Add 8705, DMD_T
  FlexDMDDict.Add 62, DMD_U
  FlexDMDDict.Add 17456, DMD_V
  FlexDMDDict.Add 20534, DMD_W
  FlexDMDDict.Add 21760, DMD_X
  FlexDMDDict.Add 9472, DMD_Y
  FlexDMDDict.Add 17417, DMD_Z

  FlexDMDDict.Add &h400,DMD_SingleQuote
  FlexDMDDict.Add 16640, DMD_CloseBracket
  FlexDMDDict.Add 5120, DMD_OpenBracket
  FlexDMDDict.Add 2120, DMD_Equals
  FlexDMDDict.Add 10275, DMD_Question
  FlexDMDDict.Add 2112, DMD_Minus
  FlexDMDDict.Add 10861, DMD_Dollar
  FlexDMDDict.Add 6144, DMD_GreaterThan
  FlexDMDDict.Add 65535, DMD_Hash
  FlexDMDDict.Add 32576, DMD_Asterick
  FlexDMDDict.Add 10816, DMD_Plus

End Sub

'*****************
' FlexCharacter
'*****************

Sub UpdateFlexChar(id, value)

  If id < 40 Then
    If FlexDMDDict.Exists (value) Then
      FlexDMDScene.GetImage("Seg" & id).Bitmap = FlexDMD.NewImage("", FlexDMDDict.Item (value)).Bitmap
    Else
      FlexDMDScene.GetImage("Seg" & id).Bitmap = FlexDMD.NewImage("", DMD_Space).Bitmap
    End If
  End If
End Sub

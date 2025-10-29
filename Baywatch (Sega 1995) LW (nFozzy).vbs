'*******************************************************************************************************
'
'                      Baywatch Sega 1995 VPX v1.0.0
'               http://www.ipdb.org/machine.cgi?id=2848
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
Const MroVol = 1  'Wire Ramps Rolling
Const ProVol = 1  'Plastic Ramps Rolling
Const ColVol = 1  'Ball Collision
Const RubVol = 1  'Rubbers Collision
Const MhiVol = 1  'Metals Hit
Const DroVol = 1  'Ball drop (4)
Const NudVol = 1  'Nudge
Const TuiVol = 0.5  'TroughLock on
Const TuoVol = 0.5  'TroughLock off
Const TufVol = 0.3  'TroughUp fire
Const ApfVol = 1  'AutoPlungers fire (3)
Const SliVol = 1  'Slingshots (2)
Const BumVol = 1  'Bumpers (3)
Const SwiVol = 1  'Rollovers (17)
Const TarVol = 1  'Targets (12)
Const DtaVol = 1  'Droptargets (3)
Const DtrVol = 1  'Droptarget reset
Const GaoVol = 0.4  'Control Gates & Trapdoor open (3)
Const GacVol = 0.4  'Control Gates & Trapdoor close (3)
Const GatVol = 1  'Gates (2)
Const SpiVol = 1  'Spinner
Const VukVol = 1  'VUK catch
Const VufVol = 1  'VUK & Shark Super Scoop & Lower Super Vuk fire (3)
Const SssVol = 1  'Shark Super Scoop catch
Const KidVol = 1  'Kicker Drain
Const FluVol = 0.6  'Flippers up
Const FldVol = 0.6  'Flippers down
Const KnoVol = 1  'Knocker
Const GirVol = 0.5  'General Illumination Relay click

'************************ ROM

Const cGameName = "bay_401"

'************************ Ball

Const BallSize = 50

Const BallMass = 1.02   'Mass=(50^3)/125000 ,(BallSize^3)/125000

'************************ Ball Shadow : 0 hidden , 1 visible

Const BallSHW = 1

'************ CabRails and rail lights Hidden/Visible in FS mode : 0 hidden , 1 visible

Const RailsLights = 1

Const bladeArt  = 2     '1=On (Art1), 2=On (Art2), 3=On (Art3), 4=On (BlackWood), 0=Sideblades Off.

Const HouseArt  = 1     '0= Original, 1 Update House

'************************ Flashers Intensity

Const Lumen = 10

'************************ VR Room ' 0 = 360 room,  1 = Normal room

Const VrEnv = 0
const VRTopper = 0        ' 0 = hide topper, 1 = topper vissible

'************ DMD '0 for VPinMAME DMD visible in DT mode, 1 for Text DMD visible in DT mode

Const VPMorTextDMD = 1

Dim VR_Room: Dim UseVPMDMD : Dim VarHidden
If Renderingmode = 2 then VR_ROOM = True Else VR_ROOM = False
If Table1.ShowDT = True and VR_ROOM = False Then
    UseVPMDMD = VPMorTextDMD
    VarHidden = VPMorTextDMD
Else
    UseVPMDMD = 0
    VarHidden = 0   'Put 1 if you whant DMD Hidden in FS mode
    ScoreText.Visible = 0
End If


' If B2SOn = True Then VarHidden = 1

'************ Slingshot mode : 0 Walls , 1 Flippers

Const SlingM = 0

'************ Slingshot hit threshold, with flippers (parm)

Const ThSling = 3

'************ Ombre fori PF

Dim x
For each x in aOmbre
  x.IntensityScale = x.IntensityScale * 0.95
  x.Height = -24
Next

'******************************************** FSS Init

 If Table1.ShowFSS = True Then
  ScoreText.Visible = False
End If

 If Table1.ShowFSS = False Then
  FlasherDMD.Visible = False
End If

f1r.Y=151:f1ra.Y=151:f2r.Y=151:f2ra.Y=151:f3r.Y=151:f3ra.Y=151:f4r.Y=151:f5r.Y=151:f5ra.Y=151:f6r.Y=151:f6ra.Y=151:f7r.Y=151:f8r.Y=151:f8ra.Y=151
fbb1.Y=151:fbb2.Y=151:fbb3.Y=151:fbb4.Y=151:fbb5.Y=151:fbb6.Y=151:fbb7.Y=151:fbb8.Y=151:fbb9.Y=151:fbb10.Y=151
fbb11.Y=151:fbb12.Y=151:fbb13.Y=151:fbb14.Y=151:fbb15.Y=151:fbb16.Y=151:fbb17.Y=151:fbb18.Y=151:fbb19.Y=151:fbb20.Y=151
fbb21.Y=151:fbb22.Y=151:fbb23.Y=151:fbb24.Y=151:fbb25.Y=151:fbb26.Y=151:fbb27.Y=151:fbb28.Y=151:fbb29.Y=151:fbb30.Y=151
fbb31.Y=151:fbb32.Y=151:fbb33.Y=151

'******************************************** OPTIONS END **********************************************

LoadVPM "01560000", "de2.VBS", 3.26

'Set Controller = CreateObject("b2s.server")

Dim bsTrough, bsTroughUp, bsUpVUK, bsSSScoop, bsLowSVuk, cdtBank, PinPlay

Const UseSolenoids = 2
Const UseLamps = 0
Const UseGI = 0
Const UseSync = 0
'Const HandleMech = 0

' Standard Sounds
Const SSolenoidOn = "Solenoid"
Const SSolenoidOff = ""
Const SFlipperOn = ""
Const SFlipperOff = ""
Const SCoin = "coin3"

'************
' Table init.
'************

Sub Table1_Init
  vpmInit me
  With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
    .SplashInfoLine = "Baywatch (Sega 1995)" & vbNewLine & "VPX table by Kiwi 1.0.0"
    .HandleMechanics = 0
    .HandleKeyboard = 0
    .ShowDMDOnly = 1
    .ShowFrame = 0
    .ShowTitle = 0
    .Hidden = VarHidden
'   .DoubleSize = 1
'   .Games(cGameName).Settings.Value("dmd_pos_x")=0
'   .Games(cGameName).Settings.Value("dmd_pos_y")=0
'   .Games(cGameName).Settings.Value("dmd_width")=768
'   .Games(cGameName).Settings.Value("dmd_height")=256
'   .Games(cGameName).Settings.Value("rol") = 0
'   .Games(cGameName).Settings.Value("ddraw") = 0
'   .Games(cGameName).Settings.Value("sound") = 1
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With

  ' Nudging
  vpmNudge.TiltSwitch = swTilt
  vpmNudge.Sensitivity = 2
  vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingShot, RightSlingShot)

  ' Trough
  Set bsTrough = New cvpmBallStack
  With bsTrough
    .InitSw 0, 14, 13, 12, 11, 10, 0, 0
'   .InitKick BallRelease, 80, 27
'   .InitEntrySnd "Solenoid", "Solenoid"
'   .InitExitSnd SoundFX("solenoid",DOFContactors), SoundFX("popper",DOFContactors)
    .Balls = 5
  End With

  Set bsTroughUp = New cvpmBallStack
  With bsTroughUp
    .initSaucer BallRelease, 15 , 0, 24
    .KickZ = 1.56
'   .InitExitSnd SoundFX("popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
  End With

  ' Upper VUK
  Set bsUpVUK = New cvpmBallStack
  With bsUpVUK
    .InitSaucer sw44, 44, 0, 36
    .KickZ = 1.56
'   .InitExitSnd SoundFX("popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
  End With

  ' Shark Super Scoop
    Set bsSSScoop = New cvpmBallStack
    With bsSSScoop
    .InitSaucer sw56, 56, 0, 45
    .KickForceVar = 4
'   .KickAngleVar = 1
    .KickZ = 1.56
'   .InitExitSnd SoundFX("popper",DOFContactors), SoundFX("vuk_enter",DOFContactors)
    End With

    ' Lower Super Vuk (Subway)
    Set bsLowSVuk = New cvpmBallStack
    With bsLowSVuk
    .InitSaucer sw46, 46, 0, 30
    .KickZ = 1.56
'   .InitExitSnd SoundFX("popper",DOFContactors), SoundFX("vuk_enter",DOFContactors)
    End With

  ' Drop targets
  set cdtBank = new cvpmdroptarget
  With cdtBank
    .initdrop array(sw22, sw23, sw24), array(22, 23, 24)
'   .initsnd SoundFX("DROPTARG",DOFContactors), SoundFX("DTResetB",DOFContactors)
  End With

  ' Main Timer init
  PinMAMETimer.Interval = PinMAMEInterval
  PinMAMETimer.Enabled = 1

  Plunger1.Pullback
  LaserKick.Pullback
  UpBLaunch.Pullback
  RampLG1.Collidable=1
  RampLG2.Collidable=0
  RampLG3.Collidable=0
  Botola.Open=0
  RampLG6.Collidable=0

  GIBulb13gt22.Visible=0
  GIBulb13gt23.Visible=0
  GIBulb13gt24.Visible=0

  SlingFSx1.Enabled = SlingM
  SlingFSx2.Enabled = SlingM
  SlingFDx1.Enabled = SlingM
  SlingFDx2.Enabled = SlingM

  RightSlingshot.IsDropped = SlingM
  LeftSlingshot.IsDropped = SlingM

  '************ Rails and rail lights

   If Table1.ShowDT = False And RailsLights = 0 Then
    CabRailSx.Visible = 0
    CabRailDx.Visible = 0
  End If

  'VR stuff
  Dim vrObj
  If VR_Room Then
    backbox.Visible = False
    'backGlass.Visible = False
    FlasherDMD.Visible = False
    Lockdownbar.Visible = False
    CabRailSx.Visible = False
    CabRailDx.Visible = False
    Serratura.Visible = False
    Wall36.Visible = False
    Wall36.SideVisible = False
    Wall37.Visible = False
    Wall37.SideVisible = False
    Backglass.BlendDisableLighting = 0.15 'original was a bit too dark
    'BBDown ??

    Select case VRenv
      Case 1 'Normal VR room
        For Each vrObj in colVRCabinet
          vrObj.BlendDisableLighting = 0.2
        Next
        For Each vrObj in colVRBeach
          vrObj.Visible = False
        Next
        For Each vrObj in colVRRoom
          vrObj.Visible = True
        Next
        ClockTimer.enabled = True
        BeerTimer.enabled = True
      Case Else 'for beach sphere and typos by user
        For Each vrObj in colVRCabinet
          vrObj.BlendDisableLighting = 0.9
        Next
        For Each vrObj in colVRRoom
          vrObj.Visible = False
        Next
        For Each vrObj in colVRBeach
          vrObj.Visible = true
        Next
    End Select
    If VRTopper = 0 then Pincab_BackBoxTopper.visible = False else Pincab_BackBoxTopper.visible = True
  Else
    For Each vrObj in colVRStuff
      vrObj.Visible = False
    Next
  End If
End Sub

' Choose Side Blades
  if bladeArt = 1 then
    PinCab_Blades.Image = "Sidewalls BW"
    PinCab_Blades.visible = 1
    elseif bladeArt = 2 then
    PinCab_Blades.Image = "Sidewalls BW2"
    PinCab_Blades.visible = 1
    elseif bladeArt = 3 then
    PinCab_Blades.Image = "Sidewalls BW4"
    PinCab_Blades.visible = 1
    elseif bladeArt = 4 then
    PinCab_Blades.Image = "Sidewalls Black"
    PinCab_Blades.visible = 1
  elseif bladeArt = 0 then
    PinCab_Blades.visible = 0
  End if


' Choose House Image
  if HouseArt = 0 then
    LifeGuard.Image = "LifeGuardMap1"
        LifeGuardPasserella.Image = "LifeGuardMap1"
    LifeGuard.visible = 1
  elseif bladeArt = 1 then
    LifeGuard.Image = "LifeGuardMap3"
        LifeGuardPasserella.Image = "LifeGuardMap3"
    LifeGuard.visible = 1
  End if



Sub Table1_Exit():Controller.Stop:End Sub
Sub Table1_Paused:Controller.Pause = 1:End Sub
Sub Table1_unPaused:Controller.Pause = 0:End Sub

Sub Table1_KeyDown(ByVal keycode)

  If KeyCode = LeftFlipperKey Then Controller.Switch(63) = 1 : Pincab_FlipperButtonLeft.X = Pincab_FlipperButtonLeft.X +6
  If KeyCode = RightFlipperKey Then Controller.Switch(64) = 1 : Pincab_FlipperButtonRight.X = Pincab_FlipperButtonRight.X -6
  If keycode = PlungerKey Then Controller.Switch(50) = 1 : Pincab_LaunchButton.z = Pincab_LaunchButton.z - 3 : Pincab_LaunchButton.y = Pincab_LaunchButton.y -5
  If KeyCode = KeyFront Then Controller.Switch(8) = 1
  If keycode = LockBarKey Then Controller.Switch(50) = 1
' If KeyCode = LeftFlipperKey And PinPlay=1 Then LeftFlipper.RotateToEnd:LeftFlipper1.RotateToEnd:Controller.Switch(63) = 1
' If KeyCode = RightFlipperKey And PinPlay=1 Then RightFlipper.RotateToEnd:RightFlipper1.RotateToEnd:Controller.Switch(64) = 1
  If keycode = LeftTiltKey Then Nudge 90, 4:PlaySound SoundFX("fx_nudge",0), 0, NudVol, -0.1, 0.25
  If keycode = RightTiltKey Then Nudge 270, 4:PlaySound SoundFX("fx_nudge",0), 0, NudVol, 0.1, 0.25
  If keycode = CenterTiltKey Then Nudge 0, 5:PlaySound SoundFX("fx_nudge",0), 0, NudVol, 0, 0.25
  If keycode = StartGameKey then Pincab_StartButton.y = Pincab_StartButton.y - 4 : Pincab_StartButtoninside.y = Pincab_StartButtoninside.y - 4
  If KeyDownHandler(KeyCode) Then Exit Sub
' If vpmKeyDown(keycode) Then Exit Sub



    'debug key (M)
  If keycode = "50" Then
        setlamp 191, 1
        setlamp 192, 1
        setlamp 193, 1
        setlamp 194, 1
        setlamp 195, 1
        setlamp 196, 1
        setlamp 197, 1
        setlamp 198, 1
'  Stop1.IsDropped=ABS(Stop1.IsDropped+1)
End If


End Sub

Sub Table1_KeyUp(ByVal keycode)

  If KeyCode = LeftFlipperKey Then Controller.Switch(63) = 0 : Pincab_FlipperButtonLeft.X = Pincab_FlipperButtonLeft.X -6
  If KeyCode = RightFlipperKey Then Controller.Switch(64) = 0 : Pincab_FlipperButtonRight.X = Pincab_FlipperButtonRight.X +6
  If keycode = PlungerKey Then Controller.Switch(50) = 0 : Pincab_LaunchButton.z = Pincab_LaunchButton.z + 3 : Pincab_LaunchButton.y = Pincab_LaunchButton.y +5
  If keycode = LockBarKey Then Controller.Switch(50) = 0
  If KeyCode = KeyFront Then Controller.Switch(8) = 0
' If KeyCode = LeftFlipperKey And PinPlay=1 Then LeftFlipper.RotateToStart:LeftFlipper1.RotateToStart:Controller.Switch(63) = 0
' If KeyCode = RightFlipperKey And PinPlay=1 Then RightFlipper.RotateToStart:RightFlipper1.RotateToStart:Controller.Switch(64) = 0
  If keycode = StartGameKey then Pincab_StartButton.y = Pincab_StartButton.y + 4 :  Pincab_StartButtoninside.y = Pincab_StartButtoninside.y +4
  If KeyUpHandler(KeyCode) Then Exit Sub
' If vpmKeyUp(keycode) Then Exit Sub
    'debug key (M)
  If keycode = "50" Then
        setlamp 191, 0
        setlamp 192, 0
        setlamp 193, 0
        setlamp 194, 0
        setlamp 195, 0
        setlamp 196, 0
        setlamp 197, 0
        setlamp 198, 0
End If

End Sub
'*********
' Switches
'*********

' Slings

Dim LStep, RStep

Sub LeftSlingshot_Slingshot:LS.VelocityCorrect(Activeball):vpmTimer.PulseSw 61:PlaySoundAtVol SoundFX("Slingshot",DOFContactors), SxEmKickerT1, SliVol:LeftSling.Visible=1:SxEmKickerT1.TransX=-25:LStep=0:Me.TimerEnabled=1:End Sub
Sub RightSlingshot_Slingshot:RS.VelocityCorrect(Activeball):vpmTimer.PulseSw 62:PlaySoundAtVol SoundFX("Slingshot",DOFContactors), DxEmKickerT1, SliVol:RightSling.Visible=1:DxEmKickerT1.TransX=-25:RStep=0:Me.TimerEnabled=1:End Sub

Sub SlingFSx1_Collide(parm)
 If SxEmKickerT1.TransX=0 And PinPlay=1 And parm > ThSling Then
  vpmTimer.PulseSw 61:PlaySoundAtVol SoundFX("Slingshot",DOFContactors), SxEmKickerT1, SliVol:LStep=0
  LeftSling.Visible=1:SxEmKickerT1.TransX=-25:LeftSlingshot.TimerEnabled=1:SlingFSx1.RotateToEnd:SlingFSx2.RotateToEnd
End If
End Sub
Sub SlingFSx2_Collide(parm)
 If SxEmKickerT1.TransX=0 And PinPlay=1 And parm > ThSling Then
  vpmTimer.PulseSw 61:PlaySoundAtVol SoundFX("Slingshot",DOFContactors), SxEmKickerT1, SliVol:LStep=0
  LeftSling.Visible=1:SxEmKickerT1.TransX=-25:LeftSlingshot.TimerEnabled=1:SlingFSx1.RotateToEnd:SlingFSx2.RotateToEnd
End If
End Sub

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

Sub SlingFDx1_Collide(parm)
 If DxEmKickerT1.TransX=0 And PinPlay=1 And parm > ThSling Then
  vpmTimer.PulseSw 62:PlaySoundAtVol SoundFX("Slingshot",DOFContactors), DxEmKickerT1, SliVol:RStep=0
  RightSling.Visible=1:DxEmKickerT1.TransX=-25:RightSlingshot.TimerEnabled=1:SlingFDx1.RotateToEnd:SlingFDx2.RotateToEnd
End If
End Sub
Sub SlingFDx2_Collide(parm)
 If DxEmKickerT1.TransX=0 And PinPlay=1 And parm > ThSling Then
  vpmTimer.PulseSw 62:PlaySoundAtVol SoundFX("Slingshot",DOFContactors), DxEmKickerT1, SliVol:RStep=0
  RightSling.Visible=1:DxEmKickerT1.TransX=-25:RightSlingshot.TimerEnabled=1:SlingFDx1.RotateToEnd:SlingFDx2.RotateToEnd
End If
End Sub

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

' Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 41:PlaySoundAtVol SoundFX("jet1",DOFContactors), Bumper1, BumVol:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 42:PlaySoundAtVol SoundFX("jet1",DOFContactors), Bumper2, BumVol:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 43:PlaySoundAtVol SoundFX("jet1",DOFContactors), Bumper3, BumVol:End Sub

' Eject holes
Sub Drain_Hit:bsTrough.AddBall Me:PlaysoundAtVol "drain1a", Drain, KidVol:End Sub
Sub sw44_Hit:PlaysoundAtVol "fx_kicker_enter1", sw44, VuKVol:bsUpVUK.AddBall 0:End Sub
Sub sw56_Hit:PlaysoundAtVol "vuk_enter", sw56, SssVol:bsSSScoop.AddBall 0:End Sub
Sub sw46_Hit:PlaysoundAtVol "fx_kicker_enter1", sw46, SssVol:bsLowSVuk.AddBall 0:End Sub

' Spinner
Sub sw52_Spin:vpmTimer.PulseSw 52:PlaySoundAtVol "spinner", sw52, SpiVol:End Sub

' Rollovers
Sub sw16_Hit:  Controller.Switch(16) = 1:PlaySoundAtVol "sensor", sw16, SwiVol:End Sub
Sub sw16_UnHit:Controller.Switch(16) = 0:End Sub
Sub sw26_Hit:  Controller.Switch(26) = 1:PlaySoundAtVol "sensor", sw26, SwiVol:End Sub
Sub sw26_UnHit:Controller.Switch(26) = 0:End Sub
Sub sw28_Hit:  Controller.Switch(28) = 1:PlaySoundAtVol "sensor", sw28, SwiVol:End Sub
Sub sw28_UnHit:Controller.Switch(28) = 0:End Sub
Sub sw30_Hit:  Controller.Switch(30) = 1:PlaySoundAtVol "sensor", sw30, SwiVol:End Sub
Sub sw30_UnHit:Controller.Switch(30) = 0:End Sub
Sub sw33_Hit:  Controller.Switch(33) = 1:PlaySoundAtVol "sensor", sw33, SwiVol:End Sub
Sub sw33_UnHit:Controller.Switch(33) = 0:End Sub

Sub sw38_Hit:  Controller.Switch(38) = 1:PlaySoundAtVol "sensor", sw38, SwiVol:End Sub
Sub sw38_UnHit:Controller.Switch(38) = 0:End Sub
Sub sw39_Hit:  Controller.Switch(39) = 1:PlaySoundAtVol "sensor", sw39, SwiVol:End Sub
Sub sw39_UnHit:Controller.Switch(39) = 0:End Sub
Sub sw40_Hit:  Controller.Switch(40) = 1:PlaySoundAtVol "sensor", sw40, SwiVol:End Sub
Sub sw40_UnHit:Controller.Switch(40) = 0:End Sub
Sub sw45_Hit:  Controller.Switch(45) = 1:PlaySoundAtVol "sensor", sw45, SwiVol:End Sub
Sub sw45_UnHit:Controller.Switch(45) = 0:End Sub

Sub sw47_Hit:  Controller.Switch(47) = 1:PlaySoundAtVol "sensor", sw47, SwiVol:End Sub
Sub sw47_UnHit:Controller.Switch(47) = 0:End Sub
Sub sw49_Hit:  Controller.Switch(49) = 1:PlaySoundAtVol "sensor", sw49, SwiVol:End Sub
Sub sw49_UnHit:Controller.Switch(49) = 0:End Sub

Sub sw54_Hit:  Controller.Switch(54) = 1:PlaySoundAtVol "sensor", sw54, SwiVol:End Sub
Sub sw54_UnHit:Controller.Switch(54) = 0:End Sub
Sub sw55_Hit:  Controller.Switch(55) = 1:PlaySoundAtVol "sensor", sw55, SwiVol:End Sub
Sub sw55_UnHit:Controller.Switch(55) = 0:End Sub
Sub sw57_Hit:  Controller.Switch(57) = 1:PlaySoundAtVol "sensor", sw57, SwiVol:End Sub
Sub sw57_UnHit:Controller.Switch(57) = 0:End Sub
Sub sw58_Hit:  Controller.Switch(58) = 1:PlaySoundAtVol "sensor", sw58, SwiVol:End Sub
Sub sw58_UnHit:Controller.Switch(58) = 0:End Sub
Sub sw59_Hit:  Controller.Switch(59) = 1:PlaySoundAtVol "sensor", sw59, SwiVol:End Sub
Sub sw59_UnHit:Controller.Switch(59) = 0:End Sub
Sub sw60_Hit:  Controller.Switch(60) = 1:PlaySoundAtVol "sensor", sw60, SwiVol:End Sub
Sub sw60_UnHit:Controller.Switch(60) = 0:End Sub

' Targets
Sub sw9_Hit:vpmTimer.PulseSw 9:Psw9.TransX=-5:Me.TimerEnabled=1:PlaySoundAtBallVol SoundFX("target",DOFTargets), TarVol:End Sub
Sub sw9_Timer:Psw9.TransX=0:Me.TimerEnabled=0:End Sub
Sub sw53_Hit:vpmTimer.PulseSw 53:Psw53.TransX=-5:Me.TimerEnabled=1:PlaySoundAtBallVol SoundFX("target",DOFTargets), TarVol:End Sub
Sub sw53_Timer:Psw53.TransX=0:Me.TimerEnabled=0:End Sub

Sub sw17_Hit:vpmTimer.PulseSw 17:Psw17.TransX=-5:Me.TimerEnabled=1:PlaySoundAtBallVol SoundFX("target",DOFTargets), TarVol:End Sub
Sub sw17_Timer:Psw17.TransX=0:Me.TimerEnabled=0:End Sub
Sub sw18_Hit:vpmTimer.PulseSw 18:Psw18.TransX=-5:Me.TimerEnabled=1:PlaySoundAtBallVol SoundFX("target",DOFTargets), TarVol:End Sub
Sub sw18_Timer:Psw18.TransX=0:Me.TimerEnabled=0:End Sub
Sub sw19_Hit:vpmTimer.PulseSw 19:Psw19.TransX=-5:Me.TimerEnabled=1:PlaySoundAtBallVol SoundFX("target",DOFTargets), TarVol:End Sub
Sub sw19_Timer:Psw19.TransX=0:Me.TimerEnabled=0:End Sub
Sub sw20_Hit:vpmTimer.PulseSw 20:Psw20.TransX=-5:Me.TimerEnabled=1:PlaySoundAtBallVol SoundFX("target",DOFTargets), TarVol:End Sub
Sub sw20_Timer:Psw20.TransX=0:Me.TimerEnabled=0:End Sub
Sub sw21_Hit:vpmTimer.PulseSw 21:Psw21.TransX=-5:Me.TimerEnabled=1:PlaySoundAtBallVol SoundFX("target",DOFTargets), TarVol:End Sub
Sub sw21_Timer:Psw21.TransX=0:Me.TimerEnabled=0:End Sub
Sub sw31_Hit:vpmTimer.PulseSw 31:Psw31.TransX=-5:Me.TimerEnabled=1:PlaySoundAtBallVol SoundFX("target",DOFTargets), TarVol:End Sub
Sub sw31_Timer:Psw31.TransX=0:Me.TimerEnabled=0:End Sub
Sub sw32_Hit:vpmTimer.PulseSw 32:Psw32.TransX=-5:Me.TimerEnabled=1:PlaySoundAtBallVol SoundFX("target",DOFTargets), TarVol:End Sub
Sub sw32_Timer:Psw32.TransX=0:Me.TimerEnabled=0:End Sub
Sub sw36_Hit:vpmTimer.PulseSw 36:Psw36.TransX=-5:Me.TimerEnabled=1:PlaySoundAtBallVol SoundFX("target",DOFTargets), TarVol:End Sub
Sub sw36_Timer:Psw36.TransX=0:Me.TimerEnabled=0:End Sub
Sub sw37_Hit:vpmTimer.PulseSw 37:Psw37.TransX=-5:Me.TimerEnabled=1:PlaySoundAtBallVol SoundFX("target",DOFTargets), TarVol:End Sub
Sub sw37_Timer:Psw37.TransX=0:Me.TimerEnabled=0:End Sub
Sub sw48_Hit:vpmTimer.PulseSw 48:Psw48.TransX=-5:Me.TimerEnabled=1:PlaySoundAtBallVol SoundFX("target",DOFTargets), TarVol:End Sub
Sub sw48_Timer:Psw48.TransX=0:Me.TimerEnabled=0:End Sub

' Droptargets
Sub sw22_Hit():PlaySoundAtVol SoundFX("DROPTARG",DOFDropTargets), sw22, DtaVol:End Sub
Sub sw22_Dropped:cdtbank.Hit 1:GIBulb13gt22.Visible=1:End Sub
Sub sw23_Hit():PlaySoundAtVol SoundFX("DROPTARG",DOFDropTargets), sw23, DtaVol:End Sub
Sub sw23_Dropped:cdtbank.Hit 2:GIBulb13gt23.Visible=1:End Sub
Sub sw24_Hit():PlaySoundAtVol SoundFX("DROPTARG",DOFDropTargets), sw24, DtaVol:End Sub
Sub sw24_Dropped:cdtbank.Hit 3:GIBulb13gt24.Visible=1:End Sub

' Gates
Sub LowerGate_Hit():PlaySoundAtBallVol "Gate51", GatVol:End Sub 'Necessari ?
Sub UpperGate_Hit():PlaySoundAtBallVol "Gate51", GatVol:End Sub 'Necessari ?
Sub Gate25_Hit():vpmTimer.PulseSw 25:PlaySoundAtBallVol "Gate51", GatVol:End Sub
Sub Gate3_Hit():PlaySoundAtBallVol "Gate51", GatVol:End Sub

' Fx Sounds
Sub swfx1_Hit:PlaySoundAtVol "fx_InMetalrolling", swfx1, MhiVol*0.2:End Sub
Sub swfx2_Hit:PlaySoundAtBallVol "fx_InMetalrolling", MhiVol:End Sub
Sub swfx3_Hit:PlaySoundAtBallVol "fx_InMetalrolling", MhiVol:End Sub
Sub swfx4_Hit:PlaySoundAtBallVol "fx_InMetalrolling", MhiVol:End Sub

Sub TriggerOSxR_Hit:PlaySoundAtBallVol "fx_ballrampendhit", MhiVol*3:End Sub
Sub EndWY_Hit:PlaySoundAtBallVol "fx_metalhit", MhiVol*3:End Sub

'Sub TriggerOLR_Hit:Me.TimerEnabled = 1:Me.TimerInterval = 150:End Sub
'Sub TriggerOLR_Timer:Me.TimerEnabled = 0:PlaySoundAtVol "fx_ballhit", TriggerOLR, DroVol:End Sub

Sub TriggerOLT_Hit:PlaySoundAtBallVol "fx_ballrampendhit", MhiVol*2:End Sub
Sub TriggerOVUK_Hit:PlaySoundAtBallVol "fx_ballrampendhit", MhiVol*2:End Sub

'Sub TriggerOSxR_Hit:Me.TimerEnabled = 1:Me.TimerInterval = 150:End Sub
'Sub TriggerOSxR_Timer:Me.TimerEnabled = 0:PlaySoundAtVol "fx_ballhit", TriggerOSxR, DroVol:End Sub

Sub THelper1_Hit:ActiveBall.VelZ=0:End Sub

'*********
'Solenoids
'*********

SolCallback(1) = "TroughLock"
SolCallback(2) = "bsTroughUpSolOut"
SolCallback(3) = "AutoBallLaunch"
SolCallback(4) = "UpBallLaunch"
SolCallback(5) = "bsUpVUKSolOut"
SolCallback(6) = "bsLowSVukSolOut"
SolCallback(7) = "bsSSScoopSolOut"
SolCallback(8) = "KnockerSound"
SolCallback(9) = "UpperG"
'SolCallback(10) = "LRABRelay"
SolCallback(11) = "GIRelay"
SolCallback(12) = "LockTrap"
SolCallback(13) = "AutoFlipper"
SolCallback(14) = "LowerG"
SolCallback(16) = "dtcbank"
SolCallback(22) = "LaserkickBack"
SolCallback(23) = "SolRun"

SolCallback(25) = "setlamp 191,"
SolCallback(26) = "setlamp 192,"
SolCallback(27) = "setlamp 193,"
SolCallback(28) = "setlamp 194,"
SolCallback(29) = "setlamp 195,"
SolCallback(30) = "setlamp 196,"
SolCallback(31) = "setlamp 197,"
SolCallback(32) = "setlamp 198,"

Sub TroughLock(Enabled)
 If Enabled Then
  bsTrough.ExitSol_On
  vpmCreateBall BallRelease
  bsTroughUp.AddBall 0
  PlaySoundAtVol SoundFX("metalhit",DOFContactors), BallRelease, TuiVol
Else
  PlaySoundAtVol SoundFX("fx_solenoidoff",DOFContactors), BallRelease, TuoVol
End If
End Sub

Sub bsTroughUpSolOut(Enabled)
 If Enabled Then
  bsTroughUp.ExitSol_On
  PlaySoundAtVol SoundFX("TVUKOut",DOFContactors), BallRelease, TufVol
End If
End Sub

Sub AutoBallLaunch(Enabled)
 If Enabled Then
  Plunger1.Fire
  PlaySoundAtVol SoundFX("plunger",DOFContactors), Plunger1, ApfVol
  Else
  Plunger1.Pullback
End If
End Sub

Sub UpBallLaunch(Enabled)
 If Enabled Then
  UpBLaunch.Fire
  PlaySoundAtVol SoundFX("plunger",DOFContactors), UpBLaunch, ApfVol
  Else
  UpBLaunch.Pullback
End If
End Sub

Sub bsUpVUKSolOut(Enabled)
 If Enabled Then
  bsUpVUK.ExitSol_On
  VUKArm.Z = -20
  VUKArmTimer.Enabled = 1
  PlaySoundAtVol SoundFX("popper",DOFContactors), sw44, VufVol
End If
End Sub

Sub VUKArmTimer_Timer()
  VUKArm.Z = -40
  VUKArmTimer.Enabled = 0
End Sub

Sub bsLowSVukSolOut(Enabled)
 If Enabled Then
  bsLowSVuk.ExitSol_On
  PlaySoundAtVol SoundFX("popper",DOFContactors), sw46, VufVol
End If
End Sub

Sub bsSSScoopSolOut(Enabled)
 If Enabled Then
  bsSSScoop.ExitSol_On
  PlaySoundAtVol SoundFX("popper",DOFContactors), sw56, VufVol
End If
End Sub

Sub KnockerSound(Enabled)
 If Enabled Then
  PlaySoundAtVol SoundFX("Knocker",DOFKnocker), l27, KnoVol
End If
End Sub

Sub UpperG(Enabled)
 If Enabled Then
  UpperGate.Open=1
  PlaySoundAtVol SoundFX("fx_solenoidon",DOFContactors), UpperGate, GaoVol
Else
  UpperGate.Open=0
  PlaySoundAtVol SoundFX("fx_solenoidoff",DOFContactors), UpperGate, GacVol
End If
End Sub

Sub GIRelay(Enabled)
If Enabled Then
   Table1.ColorGradeImage = "ColorGrade_4"
else
   Table1.ColorGradeImage = "ColorGradeLUT256x16_extraConSat"
end if
  Dim GIoffon
  GIoffon = ABS(ABS(Enabled) -1)
  SetLamp 200, GIoffon
  PlaySoundAtVol "sensor", ViteEsa13, GirVol
End Sub

Sub LockTrap(Enabled)
 If Enabled Then
  RampLG1.Collidable=0
  RampLG2.Collidable=1
  RampLG3.Collidable=1
  RampLG6.Collidable=1
  Botola.Open=1
  PlaySoundAtVol SoundFX("fx_solenoidon",DOFContactors), Trapdoor, GaoVol
Else
  RampLG1.Collidable=1
  RampLG2.Collidable=0
  RampLG3.Collidable=0
  RampLG6.Collidable=0
  Botola.Open=0
  PlaySoundAtVol SoundFX("fx_solenoidoff",DOFContactors), Trapdoor, GacVol
End If
End Sub

Sub AutoFlipper(Enabled)
 If Enabled Then
  LeftFlipper1.RotateToEnd:PlaySoundAtVolPitch SoundFX("flipperup1",DOFFlippers), LeftFlipper1, FluVol, 10
  RightFlipper1.RotateToEnd:PlaySoundAtVolPitch SoundFX("flipperup1",DOFFlippers), RightFlipper1, FluVol*0.9, 50
Else
  LeftFlipper1.RotateToStart:PlaySoundAtVolPitch SoundFX("flipperdown1",DOFFlippers), LeftFlipper1, FldVol, 10
  RightFlipper1.RotateToStart:PlaySoundAtVolPitch SoundFX("flipperdown1",DOFFlippers), RightFlipper1, FldVol*0.9, 50
End If
End Sub

Sub LowerG(Enabled)
 If Enabled Then
  LowerGate.Open=1
  PlaySoundAtVol SoundFX("fx_solenoidon",DOFContactors), LowerGate, GaoVol
Else
  LowerGate.Open=0
  PlaySoundAtVol SoundFX("fx_solenoidoff",DOFContactors), LowerGate, GacVol
End If
End Sub

Sub dtcbank(Enabled)
 If Enabled Then
  cdtbank.DropSol_On
  GIBulb13gt22.Visible=0
  GIBulb13gt23.Visible=0
  GIBulb13gt24.Visible=0
End If
  PlaySoundAtVol SoundFX("DTResetB",DOFContactors), sw23, DtrVol
End Sub

Sub LaserKickBack(Enabled)
 If Enabled Then
  LaserKick.Fire
  LaserKick.FireSpeed=130+(20*RND)
  PlaySoundAtVol SoundFX("plunger",DOFContactors), LaserKick, ApfVol
  Else
  LaserKick.Pullback
End If
End Sub

Sub SolRun(Enabled)
  vpmNudge.SolGameOn Enabled
 If Enabled Then
  PinPlay=1
' LeftSlingShot.Disabled=0
' RightSlingShot.Disabled=0
Else
  PinPlay=0
  SlingFSx1.RotateToStart:SlingFSx2.RotateToStart
  SlingFSx1.RotateToStart:SlingFSx2.RotateToStart
' LeftFlipper.RotateToStart
' RightFlipper.RotateToStart
' LeftFlipper1.RotateToStart
' RightFlipper1.RotateToStart
' LeftSlingShot.Disabled=1
' RightSlingShot.Disabled=1
End If
End Sub

'**************
' Flipper Subs
'**************

'SolCallback(sULFlipper)
'SolCallback(sURFlipper)
SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
  If Enabled Then
     FlipperActivate LeftFlipper, LFPress
    FlipperActivate LeftFlipper1, LFPress1
    LF.Fire
    PlaySoundAtVol SoundFX("Flipper_L01",DOFFlippers), LeftFlipper, FluVol:LeftFlipper1.RotateToEnd
    'PlaySoundAtVolPitch SoundFX("flipperup1",DOFFlippers), LeftFlipper1, FluVol, 10
  Else
         FlipperDeActivate LeftFlipper, LFPress
        FlipperDeActivate LeftFlipper1, LFPress1
    If LeftFlipper.CurrentAngle < LeftFlipper.StartAngle - 5 Then
      PlaySoundAtVol SoundFX("flipperdown1",DOFFlippers), LeftFlipper, FldVol
      'PlaySoundAtVolPitch SoundFX("flipperdown1",DOFFlippers), LeftFlipper1, FldVol, 10
    End If
        LeftFlipper.RotateToStart:LeftFlipper1.RotateToStart
  End If
End Sub

Sub SolRFlipper(Enabled)
  If Enabled Then
     FlipperActivate RightFlipper, RFPress
      FlipperActivate RightFlipper1, RFPress1
      RF.Fire
    PlaySoundAtVol SoundFX("Flipper_R01",DOFFlippers), RightFlipper, FluVol:RightFlipper1.RotateToEnd
    'PlaySoundAtVolPitch SoundFX("flipperup1",DOFFlippers), RightFlipper1, FluVol*0.9, 50
  Else
        FlipperDeActivate RightFlipper, RFPress
        FlipperDeActivate RightFlipper1, RFPress1
    If RightFlipper.CurrentAngle > RightFlipper.StartAngle + 5 Then
      PlaySoundAtVol SoundFX("flipperdown1",DOFFlippers), RightFlipper, FldVol
      'PlaySoundAtVolPitch SoundFX("flipperdown1",DOFFlippers), RightFlipper1, FldVol*0.9, 50
    End If
    RightFlipper.RotateToStart:RightFlipper1.RotateToStart
  End If
End Sub

Sub LeftFlipper_Collide(parm)
    CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
  LF.ReProcessBalls ActiveBall
  PlaySound "Flipper_Rubber", 0, parm / 60, pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper_Collide(parm)
    CheckLiveCatch Activeball, RightFlipper, RFCount, parm
  RF.ReProcessBalls ActiveBall
  PlaySound "Flipper_Rubber", 0, parm / 60, pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub LeftFlipper1_Collide(parm)
  PlaySound "Flipper_Rubber", 0, parm / 60, pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper1_Collide(parm)
  PlaySound "Flipper_Rubber", 0, parm / 60, pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
End Sub

'******************
' RealTime Updates
'******************
Set MotorCallback = GetRef("GameTimer")

Sub GameTimer
  UpdateFlipperLogo
End Sub

Sub UpdateFlipperLogo
  LogoSx.RotZ = LeftFlipper.CurrentAngle - 90
  LogoSx1.RotZ = LeftFlipper1.CurrentAngle - 90
  LogoDx.RotZ = RightFlipper.CurrentAngle + 90
  Pinna.RotZ = RightFlipper1.CurrentAngle
  Trapdoor.RotX = Botola.CurrentAngle
  LowerGatePrim.RotX = LowerGate.CurrentAngle / 2
  UpperGatePrim.RotX = UpperGate.CurrentAngle / 2
  UGateArm.RotX = UpperGate.CurrentAngle / 10 - 8
  WireGLG.RotX = Gate25.CurrentAngle
  WireGTS.RotX = Gate3.CurrentAngle
lfs.RotZ = LeftFlipper.CurrentAngle
  rfs.RotZ = RightFlipper.CurrentAngle
End Sub

'**********************************
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

Sub UpdateLamps
' FadeL 0, l0, "On", "F66", "F33", "Off"
  NFadeL 1, l1
  NFadeL 2, l2
  NFadeL 3, l3
  NFadeL 4, l4
  NFadeL 5, l5
  FlashVR 7, Pincab_LaunchButton
  NFadeL 9, l9
  NFadeL 10, l10
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
  NFadeLm 25, lL25
  NFadeL 25, lR25
  NFadeL 26, l26
  NFadeL 27, l27
  NFadeL 28, l28
  NFadeL 29, l29
  NFadeL 30, l30
  NFadeL 31, l31
  NFadeL 32, l32
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
  NFadeL 59, l59
  NFadeL 60, l60
  NFadeL 61, l61
  NFadeL 62, l62
  NFadeL 63, l63
  FlashVR 64, Pincab_StartButton

  NFadeLm 200, lgi1
  NFadeLm 200, lgi2
  NFadeLm 200, GIBulb1
  NFadeLm 200, GIBulb1g
  NFadeLm 200, GIBulb2
  NFadeLm 200, GIBulb2g
  NFadeLm 200, GIBulb3
  NFadeLm 200, GIBulb3g
  NFadeLm 200, GIBulb4
  NFadeLm 200, GIBulb4g
  NFadeLm 200, GIBulb5
  NFadeLm 200, GIBulb5g
  NFadeLm 200, GIBulb6
  NFadeLm 200, GIBulb6g
  NFadeLm 200, GIBulb7
  NFadeLm 200, GIBulb7g
  NFadeLm 200, GIBulb8
  NFadeLm 200, GIBulb8g
  NFadeLm 200, GIBulb9
  NFadeLm 200, GIBulb9g
  NFadeLm 200, GIBulb10
  NFadeLm 200, GIBulb10g
  NFadeLm 200, GIBulb11
  NFadeLm 200, GIBulb11g
  NFadeLm 200, GIBulb12
  NFadeLm 200, GIBulb12g
  NFadeLm 200, GIBulb13
  NFadeLm 200, GIBulb13g
  NFadeLm 200, GIBulb13gt22
  NFadeLm 200, GIBulb13gt23
  NFadeLm 200, GIBulb13gt24
  NFadeLm 200, GIBulb14
  NFadeLm 200, GIBulb15
  NFadeLm 200, GIBulb15g
  NFadeLm 200, GIBulb16
  NFadeLm 200, GIBulb16g
  NFadeLm 200, GIBulb17
  NFadeLm 200, GIBulb17g
  NFadeLm 200, GIBulb18
  NFadeLm 200, GIBulb18g
  NFadeLm 200, GIBulb19
  NFadeLm 200, GIBulb19g
  NFadeLm 200, GIBulb20
  NFadeLm 200, GIBulb20g
  NFadeLm 200, GIBulb21
  NFadeLm 200, GIBulb21g
  NFadeLm 200, GIBulb22
  NFadeLm 200, GIBulb22g
  NFadeLm 200, GIBulb23
  NFadeLm 200, GIBulb23g
  NFadeLm 200, GIBulb24
  NFadeLm 200, GIBulb24g
  NFadeLm 200, GIBulb25
  NFadeLm 200, GIBulb25g
  NFadeLm 200, GIBulb26
  NFadeLm 200, GIBulb26g
  NFadeLm 200, GIBulb27
  NFadeLm 200, GIBulb27g
  NFadeLm 200, GIBulb28
  NFadeLm 200, GIBulb28g
  NFadeLm 200, GIBulb29
  NFadeLm 200, GIBulb29g
  NFadeLm 200, GIBulb30
  NFadeLm 200, GIBulb30g
  NFadeLm 200, GIBulb31
  NFadeLm 200, GIBulb31g
  NFadeLm 200, GIBulb32
  NFadeLm 200, GIBulb32a
  NFadeLm 200, GIBulb33
  NFadeLm 200, GIBulb33a
  NFadeLm 200, GIBumb2
  NFadeLm 200, GIBumb3

  Flashm 200, fgit1
  Flashm 200, fgit2
  Flashm 200, fgit3
  Flashm 200, fgit4
  Flashm 200, fgit5
  Flashm 200, fgit6
  Flashm 200, fgit8
  Flashm 200, fgit17
  Flashm 200, fgit26

  Flashm 200, fbb1
  Flashm 200, fbb2
  Flashm 200, fbb3
  Flashm 200, fbb4
  Flashm 200, fbb5
  Flashm 200, fbb6
  Flashm 200, fbb7
  Flashm 200, fbb8
  Flashm 200, fbb9
  Flashm 200, fbb10
  Flashm 200, fbb11
  Flashm 200, fbb12
  Flashm 200, fbb13
  Flashm 200, fbb14
  Flashm 200, fbb15
  Flashm 200, fbb16
  Flashm 200, fbb17
  Flashm 200, fbb18
  Flashm 200, fbb19
  Flashm 200, fbb20
  Flashm 200, fbb21
  Flashm 200, fbb22
  Flashm 200, fbb23
  Flashm 200, fbb24
  Flashm 200, fbb25
  Flashm 200, fbb26
  Flashm 200, fbb27
  Flashm 200, fbb28
  Flashm 200, fbb29
  Flashm 200, fbb30
  Flashm 200, fbb31
  Flashm 200, fbb32
  Flash 200, fbb33

  NFadeLm 191, l191
  NFadeLm 191, l191a
    NFadeLm 191, l191b
  NFadeLm 191, l191c
  Flashm 191, f1r
  Flashm 191, f1ra
'    FlashDLm 191, 3, DomeBW191
' FlashDL 191, 3, DomeBW191a


  NFadeLm 192, l192
  NFadeLm 192, l192a
  NFadeLm 192, l192b
  Flashm 192, f2r
  Flash 192, f2ra

  NFadeLm 193, l193
  NFadeLm 193, l193a
    NFadeLm 193, l193b
  NFadeLm 193, l193c
  Flashm 193, f3r
  Flashm 193, f3ra
' FlashDL 193, 2, DomeBW193

  NFadeLm 194, l194
NFadeLm 194, l194b
NFadeLm 194, l194c
NFadeLm 194, l194d
NFadeLm 194, l194e
  Flashm 194, f4r
' FlashDLm 194, 2, DomeBW194
' FlashDL 194, 2, DomeBW194a

  Flashm 195, f195a
  Flashm 195, f195b
  Flashm 195, f5r
  Flash 195, f5ra

NFadeLm 196, l196b
NFadeLm 196, l196c
  Flashm 196, f196
  Flashm 196, f6r
  Flashm 196, f6ra
' FlashDL 196, 2, DomeBW196

  NFadeLm 197, l197a
  NFadeLm 197, l197b
  NFadeLm 197, l197c
  Flash 197, f7r

  NFadeLm 198, l198
NFadeLm 198, l198b
NFadeLm 198, l198c
  Flashm 198, f8r
  Flashm 198, f8ra
' FlashDL 198, 2, DomeBW198

  FlashDLm 6, 1, BulboF6
  Flash 6, f6
  FlashDLm 11, 1, BulboF11
  Flash 11, f11
  FlashDLm 12, 1, BulboF12
  Flash 12, f12
  FlashDLm 13, 1, BulboF13
  Flash 13, f13
  FlashDLm 33, 1, BulboF33
  Flash 33, f33
  FlashDLm 34, 1, BulboF34
  Flash 34, f34
  FlashDLm 35, 1, BulboF35
  Flash 35, f35

End Sub

' div lamp subs

Sub InitLamps()
    Dim x
    For x = 0 to 200
        LampState(x) = 0        ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4      ' used to track the fading state
        FlashSpeedUp(x) = 0.8   ' faster speed when turning on the flasher
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

' Ramps & Primitives used as 4 step fading lights
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
            Object.IntensityScale = Lumen*FlashLevel(nr)
        Case 5 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
            If FlashLevel(nr) > FlashMax(nr) Then
                FlashLevel(nr) = FlashMax(nr)
                FadingLevel(nr) = 1 'completely on
            End if
            Object.IntensityScale = Lumen*FlashLevel(nr)
    End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it just sets the flashlevel
    Object.IntensityScale = Lumen*FlashLevel(nr)
End Sub

'Flash a primitive
Sub FlashVR(nr, object)
    Select Case FadingLevel(nr)
        Case 4 'off
            FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown(nr)
            If FlashLevel(nr) < FlashMin(nr) Then
                FlashLevel(nr) = FlashMin(nr)
                FadingLevel(nr) = 0 'completely off
            End if
            Object.BlendDisableLighting = FlashLevel(nr)
        Case 5 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
            If FlashLevel(nr) > FlashMax(nr) Then
                FlashLevel(nr) = FlashMax(nr)
                FadingLevel(nr) = 1 'completely on
            End if
            Object.BlendDisableLighting = FlashLevel(nr)
    End Select
End Sub

Sub FlashVRm(nr, object) 'multiple flashers, it just sets the flashlevel
    Object.BlendDisableLighting = FlashLevel(nr)
End Sub

' Objects DisableLighting

Sub FlashDL(nr, Limite, object)
    Select Case FadingLevel(nr)
        Case 4 'off
            FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown(nr)
            If FlashLevel(nr) < FlashMin(nr) Then
                FlashLevel(nr) = FlashMin(nr)
                FadingLevel(nr) = 0 'completely off
            End if
            Object.BlendDisableLighting = Limite*FlashLevel(nr)
        Case 5 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
            If FlashLevel(nr) > FlashMax(nr) Then
                FlashLevel(nr) = FlashMax(nr)
                FadingLevel(nr) = 1 'completely on
            End if
            Object.BlendDisableLighting = Limite*FlashLevel(nr)
    End Select
End Sub

Sub FlashDLm(nr, Limite, object) 'multiple flashers, it just sets the flashlevel
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

Sub SubWayOut_Hit()
  SubWayOut.TimerEnabled=1
End Sub

Sub SubWayOut_Timer()
  SubWayOut.Kick 0,15,1.56
  SubWayOut.TimerEnabled=0
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

'*** Determines if a Points (px,py) is inside a 4 point polygon A-D in Clockwise/CCW order


'*****************************************

Sub PlaySoundAt(soundname, tableobj) 'play sound at X and Y position of an object, mostly bumpers, flippers and other fast objects
    PlaySound soundname, 0, 1, Pan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtVol(soundname, tableobj, Vol) 'play sound at X and Y position of an object, mostly bumpers, flippers and other fast objects
    PlaySound soundname, 0, Vol, Pan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtVolPitch(soundname, tableobj, Vol, Pitch) 'play sound at X and Y position of an object, mostly bumpers, flippers and other fast objects
    PlaySound soundname, 0, Vol, Pan(tableobj), 0, Pitch, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname) ' play a sound at the ball position, like rubbers, targets
    PlaySound soundname, 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtBallVol(soundname, VolMult) ' play a sound at the ball position, like rubbers, targets
    PlaySound soundname, 0, Vol(ActiveBall) * VolMult, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

'*******************************************
'   JP's VP10 Rolling Sounds + Ballshadow
' uses a collection of shadows, aBallShadow
'*******************************************

Const tnob = 20 ' total number of balls
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
    Dim BOT, b
    BOT = GetBalls
' TextBox001.Text = (UBound(BOT))+1

    ' stop the sound of deleted balls
    For b = UBound(BOT) + 1 to tnob
    rolling(b) = False
    aBallShadow(b).Visible = 0
    StopSound("fx_ballrolling" & b)
    StopSound("fx_Rolling_Plastic" & b)
    StopSound("fx_Rolling_Metal" & b)
    Next

    ' exit the sub if no balls on the table
    If UBound(BOT) = lob - 1 Then Exit Sub

    ' play the rolling sound for each ball and draw the shadow
    For b = lob to UBound(BOT)

    aBallShadow(b).X = BOT(b).X
    aBallShadow(b).Y = BOT(b).Y
    aBallShadow(b).Height = BOT(b).Z - 20  '(BallSize / 2)

        If BallVel(BOT(b)) > 1 Then
            rolling(b) = True

'Playfield
      If BOT(b).z < 30 Then
          StopSound("fx_Rolling_Metal" & b):StopSound("fx_Rolling_Plastic" & b)
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*RolVol, Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
      Else
'Wire Ramps
'55-70 Uscita U 80 95
        If InRect(BOT(b).x, BOT(b).y, 16,266,62,266,135,1509,16,1509) And BOT(b).z < 100 And BOT(b).z > 79 Then
          StopSound("fx_ballrolling" & b):StopSound("fx_Rolling_Plastic" & b)
          PlaySound("fx_Rolling_Metal" & b), -1, Vol(BOT(b) )*3*MroVol, Pan(BOT(b) ), 0, Pitch(BOT(b) )*0.1, 1, 0, AudioFade(BOT(b) )
'90-96 Uscita Y 115 121
      ElseIf InRect(BOT(b).x, BOT(b).y, 704,159,1144,159,78,883,78,822) And BOT(b).z < 122 And BOT(b).z > 114 Then
          StopSound("fx_ballrolling" & b):StopSound("fx_Rolling_Plastic" & b)
          PlaySound("fx_Rolling_Metal" & b), -1, Vol(BOT(b) )*3*MroVol, Pan(BOT(b) ), 0, Pitch(BOT(b) )*0.1, 1, 0, AudioFade(BOT(b) )
'62-150 Lancio
      ElseIf InRect(BOT(b).x, BOT(b).y, 810,1350,952,1350,1160,1930,883,1930) And BOT(b).z > 85 Then
          StopSound("fx_ballrolling" & b):StopSound("fx_Rolling_Plastic" & b)
          PlaySound("fx_Rolling_Metal" & b), -1, Vol(BOT(b) )*3*MroVol, Pan(BOT(b) ), 0, Pitch(BOT(b) )*0.1, 1, 0, AudioFade(BOT(b) )
      ElseIf InRect(BOT(b).x, BOT(b).y, 0,450,70,450,952,1350,810,1350) And BOT(b).z > 134 Then
          StopSound("fx_ballrolling" & b):StopSound("fx_Rolling_Plastic" & b)
          PlaySound("fx_Rolling_Metal" & b), -1, Vol(BOT(b) )*3*MroVol, Pan(BOT(b) ), 0, Pitch(BOT(b) )*0.1, 1, 0, AudioFade(BOT(b) )
      ElseIf InRect(BOT(b).x, BOT(b).y, 5,125,70,125,70,450,5,450) And BOT(b).z > 149 Then
          StopSound("fx_ballrolling" & b):StopSound("fx_Rolling_Plastic" & b)
          PlaySound("fx_Rolling_Metal" & b), -1, Vol(BOT(b) )*3*MroVol, Pan(BOT(b) ), 0, Pitch(BOT(b) )*0.1, 1, 0, AudioFade(BOT(b) )
      ElseIf InRect(BOT(b).x, BOT(b).y, 5,7,534,7,534,125,5,125) And BOT(b).z < 155 And BOT(b).z > 95 Then
          StopSound("fx_ballrolling" & b):StopSound("fx_Rolling_Plastic" & b)
          PlaySound("fx_Rolling_Metal" & b), -1, Vol(BOT(b) )*3*MroVol, Pan(BOT(b) ), 0, Pitch(BOT(b) )*0.1, 1, 0, AudioFade(BOT(b) )
'215-110 VUK 135
      ElseIf InRect(BOT(b).x, BOT(b).y, 360,226,470,226,945,955,600,955) And BOT(b).z > 135 Then
          StopSound("fx_ballrolling" & b):StopSound("fx_Rolling_Plastic" & b)
          PlaySound("fx_Rolling_Metal" & b), -1, Vol(BOT(b) )*3*MroVol, Pan(BOT(b) ), 0, Pitch(BOT(b) )*0.1, 1, 0, AudioFade(BOT(b) )
'171-110 Uscita Lifeguard
      ElseIf InRect(BOT(b).x, BOT(b).y, 865,307,936,307,936,728,892,728) And BOT(b).z > 135 Then
          StopSound("fx_ballrolling" & b):StopSound("fx_Rolling_Plastic" & b)
          PlaySound("fx_Rolling_Metal" & b), -1, Vol(BOT(b) )*3*MroVol, Pan(BOT(b) ), 0, Pitch(BOT(b) )*0.1, 1, 0, AudioFade(BOT(b) )
      ElseIf InRect(BOT(b).x, BOT(b).y, 362,0,1100,307,865,307,342,64) And BOT(b).z > 170 Then
          StopSound("fx_ballrolling" & b):StopSound("fx_Rolling_Plastic" & b)
          PlaySound("fx_Rolling_Metal" & b), -1, Vol(BOT(b) )*3*MroVol, Pan(BOT(b) ), 0, Pitch(BOT(b) )*0.1, 1, 0, AudioFade(BOT(b) )

'Plastic Ramps
      Else
          StopSound("fx_Rolling_Metal" & b):StopSound("fx_ballrolling" & b)
          PlaySound("fx_Rolling_Plastic" & b), -1, Vol(BOT(b) )*3*ProVol, Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
      End If
      End If

        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
        StopSound("fx_Rolling_Plastic" & b)
        StopSound("fx_Rolling_Metal" & b)
                rolling(b) = False
            End If
        End If

    ' play ball drop sounds
    If BOT(b).VelZ < -8 and BOT(b).VelY < 11 And BOT(b).z < 50 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
      PlaySound ("fx_ballhit" & b), 0, (ABS(BOT(b).velz)/17)*DroVol, Pan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
    End If

    ' Ball Shadow
    If BallSHW = 1 Then
      aBallShadow(b).Visible = 1
    Else
      aBallShadow(b).Visible = 0
    End If

    Next
End Sub

'0-180 Lifeguard

'0-65 Lifeguard U
'0-103 Y Sx
'0-115 Y Dx

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
     FlipperCradleCollision ball1, ball2, velocity
    PlaySound("fx_collide"), 0, Csng(ColVol*((velocity) ^2 / 200)), Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

'*********************************
' Diverse Collection Hit Sounds
'*********************************

Sub aMetals_Hit(idx):PlaySoundAtBall "fx_MetalHit":End Sub
'Sub aMetalWires_Hit(idx):PlaySoundAtBall "fx_MetalWire":End Sub
Sub aRubber_Bands_Hit(idx):PlaySoundAtBall "fx_rubber_band":End Sub
'Sub aRubber_LongBands_Hit(idx):PlaySoundAtBall "fx_rubber_longband":End Sub
'Sub aRubber_Posts_Hit(idx):PlaySoundAtBall "fx_rubber_post":End Sub
'Sub aRubber_Pins_Hit(idx):PlaySoundAtBall "fx_rubber_pin":End Sub
'Sub aRubber_Pegs_Hit(idx):PlaySoundAtBall "fx_rubber_peg":End Sub
Sub aPlastics_Hit(idx):PlaySoundAtBall "fx_PlasticHit":End Sub
Sub aGates_Hit(idx):PlaySoundAtBall "fx_Gate":End Sub
Sub aWoods_Hit(idx):PlaySoundAtBall "fx_Woodhit":End Sub



'***************************************************************************
'VR Beer Bubble Code - Rawd
'***************************************************************************
Sub BeerTimer_Timer()

  Randomize(21)
  BeerBubble1.z = BeerBubble1.z + Rnd(1)*0.5
  if BeerBubble1.z > -771 then BeerBubble1.z = -955
  BeerBubble2.z = BeerBubble2.z + Rnd(1)*1
  if BeerBubble2.z > -768 then BeerBubble2.z = -955
  BeerBubble3.z = BeerBubble3.z + Rnd(1)*1
  if BeerBubble3.z > -768 then BeerBubble3.z = -955
  BeerBubble4.z = BeerBubble4.z + Rnd(1)*0.75
  if BeerBubble4.z > -774 then BeerBubble4.z = -955
  BeerBubble5.z = BeerBubble5.z + Rnd(1)*1
  if BeerBubble5.z > -771 then BeerBubble5.z = -955
  BeerBubble6.z = BeerBubble6.z + Rnd(1)*1
  if BeerBubble6.z > -774 then BeerBubble6.z = -955
  BeerBubble7.z = BeerBubble7.z + Rnd(1)*0.8
  if BeerBubble7.z > -768 then BeerBubble7.z = -955
  BeerBubble8.z = BeerBubble8.z + Rnd(1)*1
  if BeerBubble8.z > -771 then BeerBubble8.z = -955
End Sub


'***************************************************************************
'VR Clock code below - THANKS RASCAL
'***************************************************************************


' VR Clock code below....
Sub ClockTimer_Timer()
  VRClockMinutes.RotAndTra2 = (Minute(Now())+(Second(Now())/100))*6
  VRClockhours.RotAndTra2 = Hour(Now())*30+(Minute(Now())/2)
    VRClockseconds.RotAndTra2 = (Second(Now()))*6
End Sub

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
'Sub InitPolarity()
'   dim x, a : a = Array(LF, RF)
' for each x in a
'   x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
'   x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
'   x.enabled = True
'   x.TimeDelay = 80
'   x.DebugOn=False ' prints some info in debugger
'
'   x.AddPt "Polarity", 0, 0, 0
'   x.AddPt "Polarity", 1, 0.05, - 3.7
'   x.AddPt "Polarity", 2, 0.16, - 3.7
'   x.AddPt "Polarity", 3, 0.22, - 0
'   x.AddPt "Polarity", 4, 0.25, - 0
'   x.AddPt "Polarity", 5, 0.3, - 2
'   x.AddPt "Polarity", 6, 0.4, - 3
'   x.AddPt "Polarity", 7, 0.5, - 3.7
'   x.AddPt "Polarity", 8, 0.65, - 2.3
'   x.AddPt "Polarity", 9, 0.75, - 1.5
'   x.AddPt "Polarity", 10, 0.81, - 1
'   x.AddPt "Polarity", 11, 0.88, 0
'   x.AddPt "Polarity", 12, 1.3, 0
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

Sub InitPolarity()
  Dim x, a
  a = Array(LF, RF)
  For Each x In a
    x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
    x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
    x.enabled = True
    x.TimeDelay = 60
    x.DebugOn=False ' prints some info in debugger

    x.AddPt "Polarity", 0, 0, 0
    x.AddPt "Polarity", 1, 0.05, - 5.5
    x.AddPt "Polarity", 2, 0.16, - 5.5
    x.AddPt "Polarity", 3, 0.20, - 0.75
    x.AddPt "Polarity", 4, 0.25, - 1.25
    x.AddPt "Polarity", 5, 0.3, - 1.75
    x.AddPt "Polarity", 6, 0.4, - 3.5
    x.AddPt "Polarity", 7, 0.5, - 5.25
    x.AddPt "Polarity", 8, 0.7, - 4.0
    x.AddPt "Polarity", 9, 0.75, - 3.5
    x.AddPt "Polarity", 10, 0.8, - 3.0
    x.AddPt "Polarity", 11, 0.85, - 2.5
    x.AddPt "Polarity", 12, 0.9, - 2.0
    x.AddPt "Polarity", 13, 0.95, - 1.5
    x.AddPt "Polarity", 14, 1, - 1.0
    x.AddPt "Polarity", 15, 1.05, -0.5
    x.AddPt "Polarity", 16, 1.1, 0
    x.AddPt "Polarity", 17, 1.3, 0

    x.AddPt "Velocity", 0, 0, 0.85
    x.AddPt "Velocity", 1, 0.23, 0.85
    x.AddPt "Velocity", 2, 0.27, 1
    x.AddPt "Velocity", 3, 0.3, 1
    x.AddPt "Velocity", 4, 0.35, 1
    x.AddPt "Velocity", 5, 0.6, 1 '0.982
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
    FlipperTricks LeftFlipper1, LFPress1, LFCount1, LFEndAngle1, LFState1
  FlipperTricks RightFlipper1, RFPress1, RFCount1, RFEndAngle1, RFState1
  FlipperNudge RightFlipper, RFEndAngle, RFEOSNudge, LeftFlipper, LFEndAngle
  FlipperNudge LeftFlipper, LFEndAngle, LFEOSNudge,  RightFlipper, RFEndAngle
End Sub

Dim LFEOSNudge, RFEOSNudge

Sub FlipperNudge(Flipper1, Endangle1, EOSNudge1, Flipper2, EndAngle2)
  Dim b
     Dim gBOT
     gBOT = GetBalls

  If Flipper1.currentangle = Endangle1 And EOSNudge1 <> 1 Then
    EOSNudge1 = 1
    '   debug.print Flipper1.currentangle &" = "& Endangle1 &"--"& Flipper2.currentangle &" = "& EndAngle2
    If Flipper2.currentangle = EndAngle2 Then
      For b = 0 To UBound(gBOT)
        If FlipperTrigger(gBOT(b).x, gBOT(b).y, Flipper1) Then
          'Debug.Print "ball in flip1. exit"
          Exit Sub
        End If
      Next
      For b = 0 To UBound(gBOT)
        If FlipperTrigger(gBOT(b).x, gBOT(b).y, Flipper2) Then
          gBOT(b).velx = gBOT(b).velx / 1.3
          gBOT(b).vely = gBOT(b).vely - 0.5
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

Dim LFPress, RFPress, LFCount, RFCount
Dim LFState, RFState
Dim EOST, EOSA,Frampup, FElasticity,FReturn
Dim RFEndAngle, LFEndAngle
Dim LFPress1, LFCount1, LFEndAngle1, LFState1
Dim RFPress1, RFCount1, RFEndAngle1, RFState1

Const FlipperCoilRampupMode = 0 '0 = fast, 1 = medium, 2 = slow (tap passes should work)

LFState = 1
LFState1 = 1
RFState = 1
RFState1 = 1
EOST = leftflipper.eostorque
EOSA = leftflipper.eostorqueangle
Frampup = LeftFlipper.rampup
FElasticity = LeftFlipper.elasticity
FReturn = LeftFlipper.return
'Const EOSTnew = 1.5 'EM's to late 80's - new recommendation by rothbauerw (previously 1)
Const EOSTnew = 1.2 '90's and later - new recommendation by rothbauerw (previously 0.8)
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
'   Const EOSReturn = 0.035  'mid 80's to early 90's
Const EOSReturn = 0.025  'mid 90's and later

LFEndAngle = Leftflipper.endangle
LFEndAngle1 = Leftflipper1.endangle
RFEndAngle = RightFlipper.endangle
RFEndAngle1 = RightFlipper1.endangle

Sub FlipperActivate(Flipper, FlipperPress)
  FlipperPress = 1
  Flipper.Elasticity = FElasticity

  Flipper.eostorque = EOST
  Flipper.eostorqueangle = EOSA
End Sub

Sub FlipperDeactivate(Flipper, FlipperPress)
  FlipperPress = 0
  Flipper.eostorqueangle = EOSA
  Flipper.eostorque = EOST * EOSReturn / FReturn

  If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 0.1 Then
    Dim b, gBOT
        gBOT = GetBalls

    For b = 0 To UBound(gBOT)
      If Distance(gBOT(b).x, gBOT(b).y, Flipper.x, Flipper.y) < 55 Then 'check for cradle
        If gBOT(b).vely >= - 0.4 Then gBOT(b).vely =  - 0.4
      End If
    Next
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





'******************************************************
'   ZDMP:  RUBBER  DAMPENERS
'******************************************************
' These are data mined bounce curves,
' dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
' Requires tracking ballspeed to calculate COR

Sub dPosts_Hit(idx)
  RubbersD.dampen ActiveBall
  TargetBouncer ActiveBall, 1
End Sub

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
      aBall.velz = aBall.velz * coef
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
'  TRACK ALL BALL VELOCITIES
'  FOR RUBBER DAMPENER AND DROP TARGETS
'******************************************************

Dim cor
Set cor = New CoRTracker

Class CoRTracker
  Public ballvel, ballvelx, ballvely

  Private Sub Class_Initialize
    ReDim ballvel(0)
    ReDim ballvelx(0)
    ReDim ballvely(0)
  End Sub

  Public Sub Update() 'tracks in-ball-velocity
    Dim str, b, AllBalls, highestID
    allBalls = GetBalls

    For Each b In allballs
      If b.id >= HighestID Then highestID = b.id
    Next

    If UBound(ballvel) < highestID Then ReDim ballvel(highestID)  'set bounds
    If UBound(ballvelx) < highestID Then ReDim ballvelx(highestID)  'set bounds
    If UBound(ballvely) < highestID Then ReDim ballvely(highestID)  'set bounds

    For Each b In allballs
      ballvel(b.id) = BallSpeed(b)
      ballvelx(b.id) = b.velx
      ballvely(b.id) = b.vely
    Next
  End Sub
End Class

' Note, cor.update must be called in a 10 ms timer. The example table uses the GameTimer for this purpose, but sometimes a dedicated timer call RDampen is used.

Sub RDampen_Timer
  Cor.Update
End Sub

'******************************************************
'****  END PHYSICS DAMPENERS
'******************************************************



'******************************************************
'   ZBOU: VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************

Const TargetBouncerEnabled = 1    '0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.9  'Level of bounces. Recommmended value of 0.7-1

Sub TargetBouncer(aBall,defvalue)
  Dim zMultiplier, vel, vratio
  If TargetBouncerEnabled = 1 And aball.z < 30 Then
    '   debug.print "velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz
    vel = BallSpeed(aBall)
    If aBall.velx = 0 Then vratio = 1 Else vratio = aBall.vely / aBall.velx
    Select Case Int(Rnd * 6) + 1
      Case 1
        zMultiplier = 0.2 * defvalue
      Case 2
        zMultiplier = 0.25 * defvalue
      Case 3
        zMultiplier = 0.3 * defvalue
      Case 4
        zMultiplier = 0.4 * defvalue
      Case 5
        zMultiplier = 0.45 * defvalue
      Case 6
        zMultiplier = 0.5 * defvalue
    End Select
    aBall.velz = Abs(vel * zMultiplier * TargetBouncerFactor)
    aBall.velx = Sgn(aBall.velx) * Sqr(Abs((vel ^ 2 - aBall.velz ^ 2) / (1 + vratio ^ 2)))
    aBall.vely = aBall.velx * vratio
    '   debug.print "---> velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz
    '   debug.print "conservation check: " & BallSpeed(aBall)/vel
  End If
End Sub

'Add targets or posts to the TargetBounce collection if you want to activate the targetbouncer code from them
Sub TargetBounce_Hit(idx)
  TargetBouncer ActiveBall, 1
End Sub



'******************************************************
' ZSSC: SLINGSHOT CORRECTION FUNCTIONS by apophis
'******************************************************
' To add these slingshot corrections:
'  - On the table, add the endpoint primitives that define the two ends of the Slingshot
'  - Initialize the SlingshotCorrection objects in InitSlingCorrection
'  - Call the .VelocityCorrect methods from the respective _Slingshot event sub

Dim LS
Set LS = New SlingshotCorrection
Dim RS
Set RS = New SlingshotCorrection

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

'' The following sub are needed, however they may exist somewhere else in the script. Uncomment below if needed
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


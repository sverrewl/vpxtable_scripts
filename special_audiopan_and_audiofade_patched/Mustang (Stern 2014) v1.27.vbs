'##################################################
'#######                                   ########
'#######          Mustang                  ########
'#######          Stern 2014               ########
'#######                                   ########
'##################################################

'***************************************************
'* Original VPX - Dozer, Dark, Stefenaustria
'* Original VPX VR Room - Senseless
'*
'* VPX 10.7x, 10.8x - TastyWasps - Oct. 2023
'*
'* nFozzy Physics, Fleep Sounds, VR Room, Flupper Domes
'* Lighting Tweaks, Dynamic Ball Shadows, Multi-Ball Cradle Physics
'* Script Organization, Ramp Refractions, Gameplay Balance, PWM Flashers
'* VR Cabinet Assets - Senseless
'* Playfield Upscales - MovieGuru
'* Arcade VR Room Assets - Senseless
'* Arcade VR Room Setup - DGrimmReaper
'* Testing - VPW Team
'****************************************************
Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

'********************
' Table Options
'********************

'----- Target Bouncer Levels -----
Const TargetBouncerEnabled = 1    ' 0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.7   ' Level of bounces. Recommmended value of 0.7

'----- General Sound Options -----
Const VolumeDial = 0.8        ' Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Const BallRollVolume = 0.5      ' Level of ball rolling volume. Value between 0 and 1
Const RampRollVolume = 0.5      ' Level of ramp rolling volume. Value between 0 and 1

'----- Ball Shadow Options -----
Const DynamicBallShadowsOn = 1    ' 0 = no dynamic ball shadow ("triangles" near slings and such), 1 = enable dynamic ball shadow
Const AmbientBallShadowOn = 1   ' 0 = Static shadow under ball ("flasher" image, like JP's)
'                 ' 1 = Moving ball shadow ("primitive" object, like ninuzzu's) - This is the only one that shows up on the pf when in ramps and fades when close to lights!
'                 ' 2 = flasher image shadow, but it moves like ninuzzu's

'----- Table Options -----
Const Inteceptor = 0    ' Replace yellow Showroom Car with Mad Max Inteceptor
Const Flasher_Halos = 0   ' Render Halos around flashers.
Const Tacho_Mod = 1     ' Render spinner activated tacho on right ramp.
Const Erratic_Scoop = 1   ' Make the Mustang scoop / saucer behave more organically.
Const Car_Color = 9     ' Turntable car color (Select from list below)
              ' 1 Light Blue, 2 Dark Blue, 3 White, 4 Green, 5 Orange. 6 Randr, 7 Red, 8 Purple, 9 Yellow

'----- VR Table Options -----
Const DMDReflection = 0   ' (0 = not visible, 1 = visible)  - Visible DMD reflections on glassplate
Const GlassScratches = 0  ' (0 = none, 1 = normal.  2 = more. 3 = less) - Visible scratches on glassplate
Const SidewallChoice = 3  ' (0 = default, 1 = custom 1, 2= custom 2, 3 = Official sideblades) - Different sidewalls
Const CoinOnGlass = 0   ' (0 = no coin on glass, 1 = coin on glass)
Const VRRoomChoice = 1    ' (1 = Minimal Room. 2 = Arcade Room (Senseless))

'----- VPM DMD Option -----
Const ShowVPMDMD = True   ' False = Do not show native VPinMame DMD in desktop/vr,  True = Show native VPinMame DMD in desktop/vr

Const BallSize = 50
Const BallMass = 1

Const UseVPMModSol = 1

'*************************************************************
' Desktop Setup
'*************************************************************
Dim VarHidden, UseVPMDMD

If Table.ShowDT = True Then
  L104.bulbhaloheight = 260
  L104a.bulbhaloheight = 260
  L103.bulbhaloheight = 186
  L103a.bulbhaloheight = 186
  L101.bulbhaloheight = 258
  L101a.bulbhaloheight = 258
  L102.bulbhaloheight = 280
  L102a.bulbhaloheight = 280
  L107.bulbhaloheight = 250
  L107a.bulbhaloheight = 250
  L106.bulbhaloheight = 270
  L106a.bulbhaloheight = 270
  L105.bulbhaloheight = 250
  L105a.bulbhaloheight = 250
  xm.bulbhaloheight = 250
  xu.bulbhaloheight = 250
  xs.bulbhaloheight = 250
  xt.bulbhaloheight = 250
  xa.bulbhaloheight = 250
  xn.bulbhaloheight = 250
  xg.bulbhaloheight = 250
  Ramp16.visible = 1
  Ramp15.visible = 1
  If ShowVPMDMD = True Then
    UseVPMDMD = True
  End If
  VarHidden = 1
Else
  UseVPMDMD = False
  VarHidden = 0
  Ramp16.visible = 0
  Ramp15.visible = 0
End If

LoadVPM "01560000", "sam.VBS", 3.10

Sub LoadVPM(VPMver, VBSfile, VBSver)
  On Error Resume Next
  If ScriptEngineMajorVersion < 5 Then MsgBox "VB Script Engine 5.0 or higher required"
  ExecuteGlobal GetTextFile(VBSfile)
  If Err Then MsgBox "Unable to open " & VBSfile & ". Ensure that it is in the same folder as this table. " & vbNewLine & Err.Description
  If Table.ShowDT = True Then
    Set Controller = CreateObject("VPinMAME.Controller")
    B2SOn = 0
  Else
    Set Controller = CreateObject("B2S.server")
    'Set Controller = CreateObject("VPinMAME.Controller")
    B2SOn = 1
  End If
  If Err Then MsgBox "Can't Load VPinMAME." & vbNewLine & Err.Description
  If VPMver > "" Then If Controller.Version < VPMver Or Err Then MsgBox "VPinMAME ver " & VPMver & " required."
  If VPMver > "" Then If Controller.Version < VPMver Or Err Then MsgBox "VPinMAME ver " & VPMver & " required."
  If VPinMAMEDriverVer < VBSver Or Err Then MsgBox VBSFile & " ver " & VBSver & " or higher required."
  On Error Goto 0
End Sub

'********************
' Standard definitions
'********************
Const cGameName = "mt_145hc"

Const UseSolenoids = 1
Const UseLamps = 0
Const UseSync = 0
Const HandleMech = 0

' Standard Sounds
Const SSolenoidOn = "Solenoid"
Const SSolenoidOff = ""
Const SCoin = ""

Dim tablewidth:tablewidth = Table.width
Dim tableheight:tableheight = Table.height
Dim gilvl: gilvl = 1

Const tnob = 6
Const lob = 0

'*************************************************************
' VR Room Auto-Detect
'*************************************************************
Dim VR_Obj

Dim DesktopMode: DesktopMode = Table.ShowDT

If RenderingMode = 2 and VRRoomChoice = 1 Then
  Ramp16.visible = 0
  Ramp15.visible = 0
  Prim_Cabside.Visible = 0
  Light4.Visible = 0
  Redline.Visible = 0
  Primary_Front_Shades.Visible = 1
  Primary_Front_Shades.SideVisible = 1
  For Each VR_Obj in VRCabinet : VR_Obj.Visible = 1 : Next
  For Each VR_Obj in VRMinimalRoom : VR_Obj.Visible = 1 : Next
  For Each VR_Obj in VRArcadeRoom : VR_Obj.Visible = 0 : Next
  If ShowVPMDMD = True Then
    UseVPMDMD = True
  End If
Else
  If RenderingMode = 2 and VRRoomChoice = 2 Then
    Ramp16.visible = 0
    Ramp15.visible = 0
    Prim_Cabside.Visible = 0
    Light4.Visible = 0
    Redline.Visible = 0
    Primary_Front_Shades.Visible = 1
    Primary_Front_Shades.SideVisible = 1
    For Each VR_Obj in VRCabinet : VR_Obj.Visible = 1 : Next
    For Each VR_Obj in VRMinimalRoom : VR_Obj.Visible = 0 : Next
    For Each VR_Obj in VRArcadeRoom : VR_Obj.Visible = 1 : Next
    If ShowVPMDMD = True Then
      UseVPMDMD = True
    End If
  Else
    For Each VR_Obj in VRCabinet : VR_Obj.Visible = 0 : Next
    For Each VR_Obj in VRMinimalRoom : VR_Obj.Visible = 0 : Next
    For Each VR_Obj in VRArcadeRoom : VR_Obj.Visible = 0 : Next
  End If
End If

'************
' Table init
'************
Dim xx, x
Dim Bump1, Bump2, Bump3, Bump4, Mech3bank,bsTrough,bsRHole,DTBank5,turntable, cbRight, XTurn, mspinmagnet
Dim PlungerIM, B2SOn

Sub Table_Init
  With Controller
    vpmInit Me
    .GameName = cGameName
   If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
   .SplashInfoLine = "Mustang (Stern 2014)"
    .HandleKeyboard = 0
   .ShowTitle = 0
    .ShowDMDOnly = 1
    .ShowFrame = 0
    .HandleMechanics = 1
    .Hidden = VarHidden
   On Error Resume Next
    .Run GetPlayerHWnd
    If Err Then MsgBox Err.Description
  End With

  On Error Goto 0

  Set bsTrough = New cvpmBallStack
  bsTrough.InitSw 0, 23, 22, 21, 20, 19, 18, 0
  bsTrough.InitKick BallRelease, 90, 8
  bsTrough.InitExitSnd "BallRelease1", "Solenoid"
  bsTrough.Balls = 6

  Set bsRHole = New cvpmBallStack
  With bsRHole
    .InitSaucer sw43, 43, 185, 20
    .KickZ = 0.4
    .InitExitSnd "fx_solenoid", "Solenoid"
    .KickForceVar = 2
  End With

  Set cbRight = New cvpmCaptiveBall
  With cbRight
    .InitCaptive RCaptTrigger, RCaptWall, Array(RCaptKicker1, RCaptKicker1a), 10
    .RestSwitch = 9
    .NailedBalls = 1
    .ForceTrans = 1
    .MinForce = 7
    .Start
  End With

  RCaptKicker1.CreateBall

  Set mSpinMagnet = New cvpmMagnet
  With mSpinMagnet
    .InitMagnet SpinMagnet, 20
    '.Solenoid = 35 'own solenoid sub
    .GrabCenter = 0
    .Size = 100
    .CreateEvents "mSpinMagnet"
    End With

  '*** Nudging
  vpmNudge.TiltSwitch=-7
  vpmNudge.Sensitivity=1
  vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3,Bumper4,LeftSlingshot,RightSlingshot)

  '*** Main Timer
  PinMAMETimer.Interval = PinMAMEInterval
 PinMAMETimer.Enabled = 1

 '*** Slings
    For Each xx in RhammerA:xx.IsDropped=1:Next
    For Each xx in RhammerB:xx.IsDropped=1:Next
    For Each xx in RhammerC:xx.IsDropped=1:Next
    For Each xx in LhammerA:xx.IsDropped=1:Next
    For Each xx in LhammerB:xx.IsDropped=1:Next
    For Each xx in LhammerC:xx.IsDropped=1:Next

 '*** Drop Targets
  Set DTBank5 = New cvpmDropTarget
  With DTBank5
    .InitDrop Array(sw34,sw35,sw36,sw37,sw38), Array(34,35,36,37,38)
'   .InitSnd "Drop_Target_Down_1", "Drop_Target_Reset_1"
  End With

  '*** Stand Ups
  sw1a.IsDropped=1
  sw2a.IsDropped=1
  sw3a.IsDropped=1
  sw4a.IsDropped=1
  sw5a.IsDropped=1
  sw41a.IsDropped=1
  sw42a.IsDropped=1
  sw54a.IsDropped=1
  sw55a.IsDropped=1

  Plunger1.Pullback
  'mSpinMagnet.MagnetOn = True

  If erratic_scoop = 1 Then
    sw43.enabled = 0
  Else
    sw43.enabled = 1
  End If

  ' Fast Flips coding for Stern SAM
  On Error Resume Next
  InitVpmFFlipsSAM
  If Err Then MsgBox "You need the latest sam.vbs in order to run this table, available with vp10.5 rev3434"
  On Error Goto 0

End Sub

Sub Table_Paused:Controller.Pause = 1:End Sub
Sub Table_unPaused:Controller.Pause = 0:End Sub
Sub Table_Exit:Controller.Pause = 0:Controller.Stop:End Sub

 '*** Keys
Sub Table_KeyDown(ByVal keycode)

  If KeyCode = MechanicalTilt Then
    vpmTimer.PulseSw vpmNudge.TiltSwitch
    Exit Sub
  End If

  If Keycode = StartGameKey Then
    Controller.Switch(16) = 1
    Primary_StartButton.y = Primary_StartButton.y - 5
    Primary_StartButton2.y = Primary_StartButton2.y - 5
  End If

  If Keycode = 3 Then Controller.Switch(15) = 1

  If Keycode = RightMagnasave or keycode = LockBarKey Then
    Controller.Switch(71) = 1
    PinCab_FireButton_1.Z = PinCab_FireButton_1.Z -6
    PinCab_FireButton_2.Z = PinCab_FireButton_2.Z -6
    PinCab_FireButton_inner.Z = PinCab_FireButton_inner.Z -6
  End If


  If keycode = LeftFlipperKey Then
    FlipperActivate LeftFlipper, LFPress
    Primary_FlipperButtonLeft.X = Primary_FlipperButtonLeft.X + 10
  End If

  If keycode = RightFlipperKey Then
    FlipperActivate RightFlipper, RFPress
        Primary_FlipperButtonRight.X = Primary_FlipperButtonRight.X - 10
  End If

  If KeyCode = PlungerKey Then
    Plunger.Pullback
    SoundPlungerPull()
    TimerPlunger.Enabled = True
    TimerPlunger2.Enabled = False
  End If

  If keycode = LeftTiltKey Then Nudge 90, 1:SoundNudgeLeft()
  If keycode = RightTiltKey Then Nudge 270, 1:SoundNudgeRight()
  If keycode = CenterTiltKey Then Nudge 0, 2:SoundNudgeCenter()

  If keycode = keyInsertCoin1 or keycode = keyInsertCoin2 or keycode = keyInsertCoin3 or keycode = keyInsertCoin4 Then
    Select Case Int(rnd*3)
      Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
      Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
      Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
    End Select
  End If

  If Keycode = AddCreditKey Then
    Primary_Coin_on_glass.visible = 0 ' Hide the coin on the glassplate - VR feature
  End If

  If keycode = StartGameKey then soundStartButton()

    If vpmKeyDown(keycode) Then Exit Sub

End Sub

Sub Table_KeyUp(ByVal keycode)

  If Keycode = RightMagnasave or keycode = LockBarKey Then
    Controller.Switch(71) = 0
    PinCab_FireButton_1.Z = PinCab_FireButton_1.Z + 6
    PinCab_FireButton_2.Z = PinCab_FireButton_2.Z + 6
    PinCab_FireButton_inner.Z = PinCab_FireButton_inner.Z + 6
  End If

  If Keycode = 3 Then Controller.Switch(15) = 0

  If Keycode = StartGameKey Then
    Controller.Switch(16) = 0
    Primary_StartButton.y = Primary_StartButton.y + 5
    Primary_StartButton2.y = Primary_StartButton2.y + 5
  End If

  If keycode = LeftFlipperKey Then
    FlipperDeActivate LeftFlipper, LFPress
    Primary_FlipperButtonLeft.X = Primary_FlipperButtonLeft.X - 10
  End If

  If keycode = RightFlipperKey Then
    FlipperDeActivate RightFlipper, RFPress
    Primary_FlipperButtonRight.X = Primary_FlipperButtonRight.X + 10
  End If

  If KeyCode = PlungerKey Then
    Plunger.Fire
    SoundPlungerReleaseBall()
    TimerPlunger.Enabled = False
    TimerPlunger2.Enabled = True
    Primary_plunger.Y = -170
  End If

  If vpmKeyUp(keycode) Then Exit Sub

End Sub

'*** Solenoids
SolCallback(1) = "solTrough"
SolCallback(2) = "solAutofire"
SolCallback(3) = "lower_ramp"
SolCallback(4) = "lower_hold"
SolCallback(5) = "upper_ramp"
SolCallback(6) = "upper_hold"
SolCallback(7) = "Reset_Drops"
'SolCallback(8) = shaker
'SolCallback(9) = l pop
'SolCallback(10) = r pop
'SolCallback(11) = B pop
'SolCallback(12) = T pop
'SolCallback(13) = l sling
'SolCallback(14) = r sling
SolModCallback(17) = "Flash17"
SolModCallback(18) = "Flash18"
SolModCallback(19) = "FlashMod19"
SolModCallback(20) = "FlashMod20"
SolModCallback(21) = "Flash21"
SolModCallback(23) = "Flash23"
SolModCallback(25) = "Flash25"
SolModCallback(26) = "Flash26"
SolModCallback(27) = "Flash27"
SolCallback(28) = "Flash28"
SolModCallback(29) = "Flash29"
SolCallback(30) = "Flash30"

SolCallback(15) = "SolLFlipper"
SolCallback(16) = "SolRFlipper"
SolCallback(22) = "Ttable"
SolCallback(51) = "Left_Orbit_Gate"
SolCallback(52) = "Right_Orbit_Gate"
SolCallback(56) = "LHDTD"
SolCallback(55) = "LHDTU"
SolCallback(54) = "RHDTD"
SolCallback(53) = "RHDTU"
SolCallback(59) = "Kicker_Out"
SolCallback(60) = "Ramp_Div"

Sub Flash17(Enabled)
  F17.Intensity = Enabled / 10
  F17a.Intensity = Enabled / 10
End Sub

Sub Flash18(Enabled)
  F18.Intensity = Enabled / 10
  F18a.Intensity = Enabled / 10
End Sub

' Flupper Dome #2 (Sling)
Sub FlashMod19(level)
  ObjTargetLevel(2) = level/255
  FlasherFlash2_Timer
End Sub

' Flupper Dome #1 (Sling)
Sub FlashMod20(level)
  ObjTargetLevel(1) = level/255
  FlasherFlash1_Timer
End Sub

Sub Flash21(Enabled)
  F21Flash2.Opacity = Enabled * 1.5
  F21F2.Opacity = Enabled * 1.5
  Fl21.Opacity = Enabled * 1.5
  F21.Intensity = Enabled / 30
  F21a.Intensity = Enabled / 10
  F21b.Intensity = Enabled / 10
  F21c.Intensity = Enabled / 50
  F21d.Intensity = Enabled / 10
End Sub

Sub Flash23(Enabled)
  F23Flash1.Opacity = Enabled * 1.5
  F23Flash2.Opacity = Enabled * 1.5
  F23.Intensity = Enabled / 30
  F23a.Intensity = Enabled / 10
  F23b.Intensity = Enabled / 10
  F23c.Intensity = Enabled / 50
  F23d.Intensity = Enabled / 10
End Sub

Sub Flash25(Enabled)
  F25.Intensity = Enabled / 10
  F25a.Intensity = Enabled / 10
End Sub

Sub Flash26(Enabled)
  F26.Intensity = Enabled / 10
  F26a.Intensity = Enabled / 10
End Sub

Sub Flash27(Enabled)
  F27.Intensity = Enabled / 10
  F27a.Intensity = Enabled / 10
End Sub

Sub Flash28(Enabled)
  If enabled Then
    F28.State = 1
    F28a.State = 1
    F28b.State = 1
    F28c.State = 1
  Else
    F28.State = 0
    F28a.State = 0
    F28b.State = 0
    F28c.State = 0
  End If
End Sub

Sub Flash29(Enabled)
  F29.Intensity = Enabled / 20
  F29a.Intensity = Enabled / 20
  F29b.Intensity = Enabled / 20
End Sub

Sub Flash30(Enabled)
  If Enabled Then
    F30.State = 1
    F30a.State = 1
  Else
    F30.State = 0
    F30a.State = 0
  End If
End Sub

Dim shopstep

Sub shoph_Timer()
  Select Case shopstep
    Case 1:
      mSpinMagnet.MagnetOn = False
      sw43.enabled = 1
      shopstep = 2
    Case 2:
      me.enabled = 0
  End Select
End Sub

Sub Reset_Drops(Enabled)
  If enabled Then
    DTBank5.DropSol_On
    RandomSoundDropTargetReset sw36
  End If
End Sub

Sub Wall261_Hit()
  If erratic_scoop = 1 Then
    mSpinMagnet.MagnetOn = True
    shopstep = 1
    shoph.enabled = 1
  End If
End Sub

Sub LHDTD(Enabled)
  If enabled Then
    LHDT.isdropped = 1
    LHDTW.isdropped = 1
    LHDTW1.isdropped = 1
    PlaySound "dtresetl"
    LHDBU.enabled = 1
    Controller.Switch(57) = 1
  End If
End Sub

Sub LHDBU_Timer()
  If LHDT.isdropped = 0 Then
    LHDT.isdropped = 1
  End If
  me.enabled = 0
End Sub

Sub LHDTU(Enabled)
  If enabled Then
    LHDT.isdropped = 0
    LHDTW.isdropped = 0
    LHDTW1.isdropped = 0
    PlaySound "dtresetl"
    Controller.Switch(57) = 0
  End If
End Sub

Sub RHDTD(Enabled)
  If enabled Then
    XRHDT.isdropped = 1
    XRHDTW.isdropped = 1
    XRHDTW1.isdropped = 1
    PlaySound "dtresetr"
    RHDBU.enabled = 1
    Controller.Switch(56) = 1
  End If
End Sub

Sub RHDBU_Timer()
  If XRHDT.isdropped = 0 Then
    XRHDT.isdropped = 1
  End If
  me.enabled = 0
End Sub

Sub RHDTU(Enabled)
  If enabled Then
    XRHDT.isdropped = 0
    XRHDTW.isdropped = 0
    XRHDTW1.isdropped = 0
    PlaySound "dtresetr"
    Controller.Switch(56) = 0
  End If
End Sub

Sub Ramp_Div(Enabled)
  If enabled Then
    LeftRampDiv.Enabled = 1
    PlaySound "diverter"
  Else
    LeftRampDiv.Enabled = 0
    PlaySound "diverter"
  End If
End Sub

Sub Left_Orbit_Gate(Enabled)
  If enabled Then
    Lorbit.open = 1
    PlaySound "dtl"
  Else
    Lorbit.open = 0
  End If
End Sub

Sub Right_Orbit_Gate(Enabled)
  If enabled Then
    Rorbit.open = 1
    PlaySound "dtr"
  Else
    Rorbit.open = 0
  End If
End Sub

Sub Kicker_Out(Enabled)
  If enabled Then
    BsRHole.ExitSol_On
    If Erratic_Scoop = 1 Then
      sw43.enabled = 0
    End If
  End If
End Sub

Sub upper_ramp(Enabled)
  If Enabled Then
    PlaySound "dtl"
  End If
 End Sub

Sub upper_hold(Enabled)
    If enabled Then
    trpos = 2
        top_ramp.enabled = 1
        topramp5.collidable = 1
  Else
    trpos = 1
    top_ramp.enabled = 1
        topramp5.collidable = 0
    PlaySound "dtl"
  End If
End Sub

Sub lower_ramp(Enabled)
  If Enabled Then
    PlaySound "dtl"
  End If
End Sub

Sub lower_hold(Enabled)
  If Enabled Then
    brpos = 2
    bottom_ramp.enabled = 1
    BotRamp2.collidable = 1
  Else
    brpos = 1
    bottom_ramp.enabled = 1
    BotRamp2.collidable = 0
    PlaySound "dtl"
  End If
End Sub

Dim brpos
brpos = 0

Sub Bottom_Ramp_Timer()
  Select Case brpos
    Case 1:
      If BotRamp.heightbottom => 60 Then
        BotRamp.heightbottom = 60
        me.enabled = 0
      End If
      BotRamp.heightbottom = BotRamp.heightbottom + 1
      BotRamp_Cover.heightbottom = BotRamp_Cover.heightbottom  + 1.5
      BotRamp_Cover.heighttop = BotRamp_Cover.heighttop  + 1
    Case 2:
      If BotRamp.heightbottom <= 0 Then
        BotRamp.heightbottom = 0
        me.enabled = 0
      End If
      BotRamp.heightbottom = BotRamp.heightbottom - 1
      BotRamp_Cover.heightbottom = BotRamp_Cover.heightbottom  - 1.5
      BotRamp_Cover.heighttop = BotRamp_Cover.heighttop  - 1
  End Select
End Sub

Dim trpos
trpos = 0

Sub Top_Ramp_Timer()
  Select Case trpos
    Case 1:
      If topRamp.heightbottom => 120 Then
        topRamp.heightbottom = 120
        topRamp2.collidable = 1
        Controller.Switch(50) = 1
        me.enabled = 0
      End If
      topRamp.heightbottom = topRamp.heightbottom + 1
      topRamp_Cover.heightbottom = topRamp_Cover.heightbottom  + 1
      topRamp_Cover.heighttop = topRamp_Cover.heighttop  + 0.5
    Case 2:
      If topRamp.heightbottom <= 60 Then
        topRamp.heightbottom = 60
        topRamp2.collidable = 0
        Controller.Switch(50) = 0
        me.enabled = 0
      End If
      topRamp.heightbottom = topRamp.heightbottom - 1
      topRamp_Cover.heightbottom = topRamp_Cover.heightbottom  - 1
      topRamp_Cover.heighttop = topRamp_Cover.heighttop  - 0.5
  End Select
End Sub

Sub solTrough(Enabled)
  If Enabled Then
   bsTrough.ExitSol_On
   vpmTimer.PulseSw 22
 End If
 End Sub

Sub solAutofire(Enabled)
  If Enabled Then
    'PlungerIM.AutoFire
        Plunger1.Fire
  Else
    Plunger1.Pullback
  End If
End Sub

 ' Captive Ball Right
Sub RCaptTrigger_Hit:cbRight.TrigHit ActiveBall:End Sub
Sub RCaptTrigger_UnHit:cbRight.TrigHit 0:End Sub
Sub RCaptWall_Hit:PlaySound "ball_collide_1":cbRight.BallHit ActiveBall:End Sub
Sub RCaptKicker1a_Hit:cbRight.BallReturn Me:End Sub

'******************************************
' Use FlipperTimers to call div subs
'******************************************

Dim LFTCount:LFTCount=1

Sub LeftFlipperTimer_Timer()
  If LFTCount < 6 Then
    LFTCount = LFTCount + 1
    LeftFlipper.Strength = StartLeftFlipperStrength*(LFTCount/6)
  Else
    Me.Enabled=0
  End If
End Sub

Dim RFTCount:RFTCount=1

Sub RightFlipperTimer_Timer()
  If RFTCount < 6 Then
    RFTCount = RFTCount + 1
    RightFlipper.Strength = StartRightFlipperStrength*(RFTCount/6)
  Else
    Me.Enabled=0
  End If
End Sub

Const ReflipAngle = 20

Sub SolLFlipper(Enabled)
  If Enabled Then
    LF.fire

    If leftflipper.currentangle < leftflipper.endangle + ReflipAngle Then
      RandomSoundReflipUpLeft LeftFlipper
    Else
      SoundFlipperUpAttackLeft LeftFlipper
      RandomSoundFlipperUpLeft LeftFlipper
    End If
     Else
     LFTCount=1

    LeftFlipper.RotateToStart
    If LeftFlipper.currentangle < LeftFlipper.startAngle - 5 Then
      RandomSoundFlipperDownLeft LeftFlipper
    End If
    FlipperLeftHitParm = FlipperUpSoundLevel
  End If
End Sub

Sub SolRFlipper(Enabled)
  If Enabled Then
    RF.fire

    If rightflipper.currentangle < rightflipper.endangle + ReflipAngle Then
      RandomSoundReflipUpRight RightFlipper
    Else
      SoundFlipperUpAttackRight RightFlipper
      RandomSoundFlipperUpRight RightFlipper
    End If
     Else
     RFTCount=1

    RightFlipper.RotateToStart
    If RightFlipper.currentangle < RightFlipper.startAngle - 5 Then
      RandomSoundFlipperDownRight RightFlipper
    End If
    FlipperRightHitParm = FlipperUpSoundLevel
  End If
End Sub

' Drains and Kickers
Sub Drain_Hit()
  RandomSoundDrain Drain
  bsTrough.AddBall Me
End Sub

Sub orbitpost(Enabled)
  If Enabled Then
    UpPost.Isdropped=false
    UpPost2.Isdropped=false
  Else
    UpPost.Isdropped=true
    UpPost2.Isdropped=true
  End If
End Sub

' Switches
Sub sw1_Hit:sw1.IsDropped = 1:sw1a.IsDropped = 0:Me.TimerEnabled = 1:vpmTimer.PulseSw 1:End Sub
Sub sw1_Timer:sw1.IsDropped = 0:sw1a.IsDropped = 1:Me.TimerEnabled = 0:End Sub
Sub sw2_Hit:sw2.IsDropped = 1:sw2a.IsDropped = 0:Me.TimerEnabled = 1:vpmTimer.PulseSw 2:End Sub
Sub sw2_Timer:sw2.IsDropped = 0:sw2a.IsDropped = 1:Me.TimerEnabled = 0:End Sub
Sub sw3_Hit:sw3.IsDropped = 1:sw3a.IsDropped = 0:Me.TimerEnabled = 1:vpmTimer.PulseSw 3:End Sub
Sub sw3_Timer:sw3.IsDropped = 0:sw3a.IsDropped = 1:Me.TimerEnabled = 0:End Sub
Sub sw4_Hit:sw4.IsDropped = 1:sw4a.IsDropped = 0:Me.TimerEnabled = 1:vpmTimer.PulseSw 4:End Sub
Sub sw4_Timer:sw4.IsDropped = 0:sw4a.IsDropped = 1:Me.TimerEnabled = 0:End Sub
Sub sw5_Hit:sw5.IsDropped = 1:sw5a.IsDropped = 0:Me.TimerEnabled = 1:vpmTimer.PulseSw 5:End Sub
Sub sw5_Timer:sw5.IsDropped = 0:sw5a.IsDropped = 1:Me.TimerEnabled = 0:End Sub
Sub sw8_Hit:Controller.Switch(8) = 1:End Sub
Sub sw8_UnHit:Controller.Switch(8) = 0:End Sub
'sw9 captive ball front
Sub sw10_Hit:Controller.Switch(10)=1:End Sub
Sub sw10_unHit:Controller.Switch(10)=0:End Sub
Sub sw11_Hit:la11.IsDropped = 1:Controller.Switch(11) = 1:End Sub
Sub sw11_UnHit:la11.IsDropped = 0:Controller.Switch(11) = 0:End Sub
Sub sw12_Hit:la12.IsDropped = 1:Controller.Switch(12) = 1:End Sub
Sub sw12_UnHit:la12.IsDropped = 0:Controller.Switch(12) = 0:End Sub
Sub sw13_Hit:la13.IsDropped = 1:Controller.Switch(13) = 1::End Sub
Sub sw13_UnHit:la13.IsDropped = 0:Controller.Switch(13) = 0:End Sub
Sub sw14_Hit:la14.IsDropped = 1:Controller.Switch(14) = 1::End Sub
Sub sw14_UnHit:la14.IsDropped = 0:Controller.Switch(14) = 0:End Sub
Sub sw24_Hit:la24.IsDropped = 1:Controller.Switch(24) = 1::End Sub
Sub sw24_UnHit:la24.IsDropped = 0:Controller.Switch(24) = 0:End Sub
Sub sw25_Hit:la25.IsDropped = 1:Controller.Switch(25) = 1:End Sub
Sub sw25_UnHit:la25.IsDropped = 0:Controller.Switch(25) = 0:End Sub
Sub sw28_Hit:la28.IsDropped = 1:Controller.Switch(28) = 1:End Sub
Sub sw28_UnHit:la28.IsDropped = 0:Controller.Switch(28) = 0:End Sub
Sub sw29_Hit:la29.IsDropped = 1:Controller.Switch(29) = 1:End Sub
Sub sw29_UnHit:la29.IsDropped = 0:Controller.Switch(29) = 0:End Sub
Sub sw34_Hit:DTBank5.Hit 1:SoundDropTargetDrop sw34:End Sub'Light54.State = 1:End Sub
Sub sw35_Hit:DTBank5.Hit 2:SoundDropTargetDrop sw35:End Sub'Light53.State = 1:End Sub
Sub sw36_Hit:DTBank5.Hit 3:SoundDropTargetDrop sw36:End Sub'Light51.State = 1:End Sub
Sub sw37_Hit:DTBank5.Hit 4:SoundDropTargetDrop sw37:End Sub'Light50.State = 1:End Sub
Sub sw38_Hit:DTBank5.Hit 5:SoundDropTargetDrop sw38:End Sub'Light52.State = 1:End Sub
Sub sw39_Hit:Controller.Switch(39) = 1:PlaySound "subway2":End Sub'Gi_Off:Gi_BO.enabled = 1:End Sub
Sub sw39_UnHit:Controller.Switch(39) = 0:End Sub
Sub sw40_Hit:Controller.Switch(40) = 1:PlaySound "subway2":End Sub'Gi_Off:Gi_BO.enabled = 1:End Sub
Sub sw40_UnHit:Controller.Switch(40) = 0:End Sub
Sub sw41_Hit:sw41.IsDropped = 1:sw41a.IsDropped = 0:Me.TimerEnabled = 1:vpmTimer.PulseSw 41:End Sub
Sub sw41_Timer:sw41.IsDropped = 0:sw41a.IsDropped = 1:Me.TimerEnabled = 0:End Sub
Sub sw42_Hit:sw42.IsDropped = 1:sw42a.IsDropped = 0:Me.TimerEnabled = 1:vpmTimer.PulseSw 42:End Sub
Sub sw42_Timer:sw42.IsDropped = 0:sw42a.IsDropped = 1:Me.TimerEnabled = 0:End Sub
Sub sw44_Hit:Controller.Switch(44)=1:End Sub
Sub sw44_unHit:Controller.Switch(44)=0:End Sub
Sub sw45_Hit:Controller.Switch(45)=1:End Sub
Sub sw45_unHit:Controller.Switch(45)=0:End Sub
Sub sw46_Hit:Controller.Switch(46) = 1:bowl_sw.rotatetoend:End Sub
Sub sw46_UnHit:Controller.Switch(46) = 0:bowl_sw.rotatetostart:End Sub
Sub sw47_Hit:Controller.Switch(47) = 1:End Sub
Sub sw47_UnHit:Controller.Switch(47) = 0:End Sub'GI_On:End Sub
Sub sw48_Spin:vpmTimer.PulseSw 48:playsound"spinner": tach.enabled = 1:End Sub
Sub sw54_Hit:sw54.IsDropped = 1:sw54a.IsDropped = 0:Me.TimerEnabled = 1:vpmTimer.PulseSw 54:End Sub
Sub sw54_Timer:sw54.IsDropped = 0:sw54a.IsDropped = 1:Me.TimerEnabled = 0:End Sub
Sub sw55_Hit:sw55.IsDropped = 1:sw55a.IsDropped = 0:Me.TimerEnabled = 1:vpmTimer.PulseSw 55:End Sub
Sub sw55_Timer:sw55.IsDropped = 0:sw55a.IsDropped = 1:Me.TimerEnabled = 0:End Sub
Sub lhdtw_Hit:vpmTimer.PulseSw 57:Playsound "target":End Sub
Sub xrhdtw_Hit:vpmTimer.PulseSw 56:Playsound "target":End Sub

' Scoop
 Dim aBall, aZpos
 Dim bBall, bZpos

' Right Saucer
Sub sw43_Hit
  SoundSaucerLock
  Me.TimerInterval = 2
  Me.TimerEnabled = 1
End Sub

Sub sw43_Timer
  Me.TimerEnabled = 0
  bsRHole.AddBall Me
End Sub

' Slingshots
Dim LStep, RStep

Sub LeftSlingShot_Slingshot
  LS.VelocityCorrect(ActiveBall)
  RandomSoundSlingshotLeft zCol_Rubber_Corner_004
  For each xx in LHammerA:xx.IsDropped = 0:Next
   vpmTimer.PulseSw 26
  LStep = 0
  Me.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
  Select Case LStep
   Case 0: 'pause
    Case 1: 'pause
    Case 2
      For each xx in LHammerA:xx.IsDropped = 1:Next
     For Each xx in LHammerB:xx.IsDropped = 0:Next
   Case 3
      For each xx in LHammerB:xx.IsDropped = 1:Next
     For each xx in LHammerC:xx.IsDropped = 0:Next
   Case 4
      For each xx in LHammerC:xx.IsDropped = 1:Next
     Me.TimerEnabled = 0
 End Select
  LStep = LStep + 1
End Sub

 Sub RightSlingShot_Slingshot
  RS.VelocityCorrect(ActiveBall)
  RandomSoundSlingshotRight zCol_Rubber_Corner_008
  For each xx in RHammerA:xx.IsDropped = 0:Next
   vpmTimer.PulseSw 27
  RStep = 0
  Me.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
 Select Case RStep
   Case 0: 'pause
    Case 1: 'pause
    Case 2:
      For each xx in RHammerA:xx.IsDropped = 1:Next
     For each xx in RHammerB:xx.IsDropped = 0:Next
   Case 3
      For each xx in RHammerB:xx.IsDropped = 1:Next
     For each xx in RHammerC:xx.IsDropped = 0:Next
   Case 4
      For each xx in RHammerC:xx.IsDropped = 1:Next
     Me.TimerEnabled = 0
 End Select
  RStep = RStep + 1
End Sub

' Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 30:RandomSoundBumperMiddle Bumper1:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 32:RandomSoundBumperBottom Bumper2:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 33:RandomSoundBumperTop Bumper3:End Sub
Sub Bumper4_Hit:vpmTimer.PulseSw 31:RandomSoundBumperMiddle Bumper1:End Sub


Sub LeftFlipper_Collide(parm)
  CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
    LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
  CheckLiveCatch Activeball, RightFlipper, RFCount, parm
  RightFlipperCollide parm
End Sub

Dim LampState(400), FlashLevel(200)

AllLampsOff()
LampTimer.Interval = 35
LampTimer.Enabled = 1

Sub LampTimer_Timer()
    Dim chgLamp, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp) Then
        For ii = 0 To UBound(chgLamp)
      LampState(chgLamp(ii, 0) ) = chgLamp(ii, 1)
        Next
    End If
    UpdateLamps
End Sub

Sub AllLampsOff()
    On Error Resume Next

  Dim x
  For x = 0 to 360
    LampState(x) = 0
  Next

  UpdateLamps:UpdateLamps:Updatelamps
End Sub

Sub SetLamp(nr, value)
  If value <> LampState(nr) Then
    LampState(nr) = abs(value)
  End If
End Sub

Sub SetModLamp(nr, level)
    FlashLevel(nr) = level / 150 'lights & flashers
End Sub

Sub SetModLampPWM(nr, level)
    FlashLevel(nr) = level / 150 'lights & flashers
End Sub

Dim xxtt

Sub UpdateLamps
  On Error Resume Next

  For each xxtt in GI_ALL:xxtt.Intensity = LampState(98) / 20:next

  If LampState(98) < 10 Then
    Table.ColorGradeImage = "ColorGrade_off"
  Else
    Table.ColorGradeImage = "ColorGrade_on"
  End If

  If sw34.isdropped AND LampState(98) > 10 Then
    Light54.State = 1
  Else
    Light54.State = 0
  End If

  if sw35.isdropped AND LampState(98) > 10 Then
    Light53.State = 1
  else
    Light53.State = 0
  End If

  if sw36.isdropped AND LampState(98) > 10 Then
    Light51.State = 1
  else
    Light51.State = 0
  End If

  if sw37.isdropped AND LampState(98) > 10 Then
    Light50.State = 1
  else
    Light50.State = 0
  End If

  if sw38.isdropped AND LampState(98) > 10 Then
    Light52.State = 1
  else
    Light52.State = 0
  End If

  l3.State = Lampstate(3)
  l4.State = Lampstate(4)
  l5.State = Lampstate(5)
  l6.State = Lampstate(6)
  l7.State = Lampstate(7)
  l8.State = Lampstate(8)

  xg.State = Lampstate(9)
  xn.State = Lampstate(10)
  xa.State = Lampstate(11)
  xt.State = Lampstate(12)
  xs.State = Lampstate(13)
  xu.State = Lampstate(14)
  xm.State = Lampstate(15)

  L33.State = Lampstate(33)
  L34.State = Lampstate(34)
  L35.State = Lampstate(35)

  L16.State = Lampstate(16)
  L17.State = Lampstate(17)
  L18.State = Lampstate(18)
  L19.State = Lampstate(19)

  L20.State = Lampstate(20)
  L21.State = Lampstate(21)
  L22.State = Lampstate(22)
  L23.State = Lampstate(23)

  L20a.State = Lampstate(20)
  L21a.State = Lampstate(21)
  L22a.State = Lampstate(22)
  L23a.State = Lampstate(23)

  L24.State = Lampstate(24)
  L26.State = Lampstate(26)
  L27.State = Lampstate(27)
  L28.State = Lampstate(28)
  L29.State = Lampstate(29)
  L30.State = Lampstate(30)
  L31.State = Lampstate(31)
  L32.State = Lampstate(32)
  L25.State = Lampstate(25)

  L36.State = Lampstate(36)
  L37.State = Lampstate(37)
  L38.State = Lampstate(38)

  L39.State = Lampstate(39)
  L40.State = Lampstate(40)
  L41.State = Lampstate(41)

  L44.State = Lampstate(44)
  L45.State = Lampstate(45)
  L46.State = Lampstate(46)

  L47.State = Lampstate(47)
  L48.State = Lampstate(48)
  L42.State = Lampstate(42)
  L43.State = Lampstate(43)
  L49.State = Lampstate(49)
  L50.State = Lampstate(50)
  L51.State = Lampstate(51)

  L52.State = Lampstate(52)
  L53.State = Lampstate(53)
  L54.State = Lampstate(54)

  L55.State = Lampstate(55)
  L56.State = Lampstate(56)
  L57.State = Lampstate(57)

  L58.State = Lampstate(58)
  L59.State = Lampstate(59)
  L60.State = Lampstate(60)
  L61.State = Lampstate(61)
  L62.State = Lampstate(62)
  L63.State = Lampstate(63)
  L64.State = Lampstate(64)

  L65.State = Lampstate(65)
  L66.State = Lampstate(66)
  L67.State = Lampstate(67)
  L68.State = Lampstate(68)
  L69.State = Lampstate(69)
  L70.State = Lampstate(70)

  L73.State = Lampstate(73)

  L77.State = Lampstate(77)
  L78.State = Lampstate(78)
  L79.State = Lampstate(79)
  L79a.State = Lampstate(79)
  L80.State = Lampstate(80)

  L108.State = Lampstate(112)

  'L39.State = Lampstate(39)

  L81.State = Lampstate(81)
  L82.State = Lampstate(82)
  L83.State = Lampstate(83)
  L84.State = Lampstate(84)
  L85.State = Lampstate(85)

  L86.State = Lampstate(86)
  L87.State = Lampstate(87)
  L88.State = Lampstate(88)
  L89.State = Lampstate(89)
  L90.State = Lampstate(90)

  L91.State = Lampstate(91)
  L92.State = Lampstate(92)
  L93.State = Lampstate(93)
  L94.State = Lampstate(94)
  L95.State = Lampstate(95)

  L96.State = Lampstate(96)
  L97.State = Lampstate(97)

  L99.State = Lampstate(102)
  L99a.State = Lampstate(102)
  L99b.State = Lampstate(102)

  L98.State = Lampstate(103)
  L98a.State = Lampstate(103)
  L98b.State = Lampstate(103)

  L100.State = Lampstate(104)
  L100a.State = Lampstate(104)
  L100b.State = Lampstate(104)

  L103.State = Lampstate(107)
  L103a.State = Lampstate(107)
  L103b.State = Lampstate(107)
  L104.State = Lampstate(108)
  L104a.State = Lampstate(108)
  L104b.State = Lampstate(108)

  L101.State = Lampstate(105)
  L101a.State = Lampstate(105)
  L101b.State = Lampstate(105)
  L102.State = Lampstate(106)
  L102a.State = Lampstate(106)
  L102b.State = Lampstate(106)
  L105.State = Lampstate(109)
  L105a.State = Lampstate(109)
  L105b.State = Lampstate(109)

  L106.State = Lampstate(110)
  L106a.State = Lampstate(110)
  L106b.State = Lampstate(110)

  L107.State = Lampstate(111)
  L107a.State = Lampstate(111)
  L107b.State = Lampstate(111)

  rgb1.Color = RGB(0,0,0)
  rgb1.ColorFull = RGB(Lampstate(125),Lampstate(123),Lampstate(124))

  rgb2.Color = RGB(0,0,0)
  rgb2.ColorFull = RGB(Lampstate(128),Lampstate(126),Lampstate(127))

  rgb3.Color = RGB(0,0,0)
  rgb3.ColorFull = RGB(Lampstate(122),Lampstate(120),Lampstate(121))

  rgb8.Color = RGB(0,0,0)
  rgb8.ColorFull = RGB(Lampstate(119),Lampstate(117),Lampstate(118))

  rgb4.Color = RGB(0,0,0)
  rgb4.ColorFull = RGB(Lampstate(141),Lampstate(139),Lampstate(140))

  rgb5.Color = RGB(0,0,0)
  rgb5.ColorFull = RGB(Lampstate(135),Lampstate(133),Lampstate(134))

  rgb6.Color = RGB(0,0,0)
  rgb6.ColorFull = RGB(Lampstate(132),Lampstate(130),Lampstate(131))

  rgb7.Color = RGB(0,0,0)
  rgb7.ColorFull = RGB(Lampstate(138),Lampstate(136),Lampstate(137))

  Pincab_FireButtonLight.Color = RGB(0,0,0)
  Pincab_FireButtonLight.ColorFull = RGB(Lampstate(144),Lampstate(142),Lampstate(143))

  N8.State = Lampstate(98)
  N9.State = Lampstate(99)
  OH.State = Lampstate(100)
  OH1.State = Lampstate(101)

End Sub

'*************************
' Plunger kicker animation
'*************************
Sub Trigger1_hit
  PlaySound "DROP_LEFT"
 End Sub

 Sub Trigger2_hit
  PlaySound "DROP_RIGHT"
 End Sub

 Sub RHD_hit
  PlaySound "DROP_RIGHT"
 End Sub

Sub Ttable(Enabled)
  If enabled Then
    car_wheel.enabled = 1
  else
    car_wheel.enabled = 0
  End If
End Sub

Sub car_wheel_Timer()

  If Disc2.ObjRotz => 360 Then
    Disc2.ObjRotz = 0
  End If

  If Disc2.ObjRotz > 3 AND Disc2.ObjRotz < 53 AND Controller.Switch (52) = False Then
    Controller.Switch (52) = 1
  else
    if Disc2.ObjRotz => 57 AND Disc2.ObjRotz <=60 Then
      Controller.Switch (52) = 0
    End If
  End If

  If Disc2.ObjRotz > 60 AND Disc2.ObjRotz < 108 AND Controller.Switch (52) = False Then
    Controller.Switch (52) = 1
  else
    if Disc2.ObjRotz => 112 AND Disc2.ObjRotz <=115 Then
      Controller.Switch (52) = 0
    End If
  End If

  If Disc2.ObjRotz > 115 AND Disc2.ObjRotz < 158 AND Controller.Switch (52) = False Then
    Controller.Switch (52) = 1
  else
    if Disc2.ObjRotz => 162 AND Disc2.ObjRotz <=165 Then
      Controller.Switch (52) = 0
    End If
  End If

  If Disc2.ObjRotz > 165 AND Disc2.ObjRotz < 208 AND Controller.Switch (52) = False Then
    Controller.Switch (52) = 1
  else
    if Disc2.ObjRotz => 212 AND Disc2.ObjRotz <=215 Then
      Controller.Switch (52) = 0
    End If
  End If

  If Disc2.ObjRotz > 215 AND Disc2.ObjRotz < 258 AND Controller.Switch (52) = False Then
    Controller.Switch (52) = 1
  else
    if Disc2.ObjRotz => 264 AND Disc2.ObjRotz <=267 Then
      Controller.Switch (52) = 0
    End If
  End If

  If Disc2.ObjRotz > 267 AND Disc2.ObjRotz < 308 AND Controller.Switch (52) = False Then
    Controller.Switch (52) = 1
  else
    if Disc2.ObjRotz => 312 AND Disc2.ObjRotz <=315 Then
      Controller.Switch (52) = 0
    End If
  End If

  If Disc2.ObjRotz > 315 AND Disc2.ObjRotz < 353 AND Controller.Switch (52) = False Then
    Controller.Switch (52) = 1
  else
    if Disc2.ObjRotz => 358 Then
      Controller.Switch (52) = 0
    End If
  End If

  If Disc2.ObjRotz > 352 Then
    Controller.Switch (53) = 1
  else
    if Disc2.ObjRotz => 3 AND Disc2.ObjRotz <= 5 Then
      Controller.Switch (53) = 0
    End If
  End If

  If Controller.Switch(52) = True Then
    Light61.State = 1
  else
    Light61.State = 0
  End If

  If Controller.Switch(53) = True Then
    Light62.State = 1
  else
    Light62.State = 0
  End If

  Primitive3.ObjRotZ = Primitive3.ObjRotZ + 0.1
  Primitive53.ObjRotZ = Primitive3.ObjRotZ + 0.1
  Disc2.ObjRotZ = Disc2.ObjRotZ + 0.1
  Disc1.ObjRotZ = Disc1.ObjRotZ - 0.2
End Sub

Dim tachpos

Tachpos = 1

Sub tach_timer()'255
  Select Case tachpos
    Case 1:
      If Disc4.Rotz => 200 Then
        redline.state = 2
      End If
      If Disc4.Rotz => 255 Then
        tachpos = 2
        Disc4.Rotz = 255
      End If
      me.enabled = 0
      Disc4.rotz = Disc4.rotz + 5
    Case 2:
      If Disc4.Rotz <= 200 Then
        redline.state = 0
      End If
      If Disc4.Rotz <= 0 Then
        tachpos = 1
        Disc4.Rotz = 0
        me.enabled = 0
      End If
      Disc4.rotz = Disc4.rotz - 5
  End Select
End Sub

'Sub DivKick1_Hit()
  'me.DestroyBall
  'DivKick2.CreateBall
  'DivKick2.Kick 110,2
  'PlaySound "metalrolling2"
'End Sub

Sub LeftRampDiv_Hit
  ActiveBall.X = 143
  ActiveBall.Y = 813
  ActiveBall.Z = 137
  activeball.velx = 0
  activeball.vely = 0
  activeball.velz = 0
End Sub

'*** Timers
Sub GameTimer_Timer
  RollingUpdate
  RightFlipperP.Rotz = RightFlipper.CurrentAngle
  LeftFlipperP.Rotz = LeftFlipper.CurrentAngle
End Sub

Sub FrameTimer_Timer
  If DynamicBallShadowsOn Or AmbientBallShadowOn Then DynamicBSUpdate ' Update ball shadows
End Sub

'*** Triggers
Sub LRampT_Hit
  WireRampOn False
End Sub

Sub BotLeftRampEnd_Hit
  WireRampOff
End Sub

Sub TopLeftPlasticRamp_Hit
  WireRampOn True
End Sub

Sub TopLeftWireRamp_Hit
  WireRampOn False
End Sub

Sub BottomLeftWireRamp_Hit
  WireRampOff
End Sub

Sub TopLeftPlasticRampContinue_Hit
  WireRampOn True
End Sub

'*** Table Color/Model Options
Select Case Car_Color
  Case 1
    Primitive3.image = "blue1"
  Case 2
    Primitive3.image = "blue2"
  Case 3
    Primitive3.image = "diamond_white"
  Case 4
    Primitive3.image = "green"
  Case 5
    Primitive3.image = "orange"
  Case 6
    Primitive3.image = "randrs_car"
  Case 7
    Primitive3.image = "red"
  Case 8
    Primitive3.image = "purple"
  Case 9
    Primitive3.image = "yellow2"
End Select

If Inteceptor = 1 Then
  Primitive3.visible = 0
  Primitive53.visible = 1
Else
  Primitive3.visible = 1
  Primitive53.visible = 0
End If

Dim dtff

If Flasher_Halos = 1 Then
  For each dtff in Flashers:dtff.visible = 1:next
Else
  For each dtff in Flashers:dtff.visible = 0:next
End If

If Tacho_Mod = 1 Then
  Disc4.visible = 1
  Disc2.visible = 1
  Disc3.visible = 1
  Tacho_Bod.visible = 1
Else
  Disc4.visible = 0
  Disc2.visible = 0
  Disc3.visible = 0
  Tacho_Bod.visible = 0
End If

If DMDReflection = 1 Then
  Primary_DMD_Reflection.visible=1
Else
  Primary_DMD_Reflection.visible=0
End If

Select Case SidewallChoice
  Case 0
    Primary_Sidewalls.image = "Cab_sidewalls"
  Case 1
    Primary_Sidewalls.image = "Cab_sidewalls_custom_1"
  Case 2
    Primary_Sidewalls.image = "Cab_sidewalls_custom_2"
  Case 3
    Primary_Sidewalls.image = "Cab_sidewalls_custom_3"
End Select

Select Case GlassScratches
  Case 0
    Primary_Glass_scratches.visible=0:Primary_windowglass.visible=0
  Case 1
    Primary_Glass_scratches.imageA= "Cab_glass_scratches":Primary_windowglass.visible=1
  Case 2
    Primary_Glass_scratches.imageA= "Cab_glass_scratches_custom_1":Primary_windowglass.visible=1
  Case 3
    Primary_Glass_scratches.imageA= "Cab_glass_scratches_custom_2":Primary_windowglass.visible=1
End Select

If CoinOnGlass = 0 Then
  Primary_Coin_on_glass.visible = 0
End If

'******************************************************
' ZPHY:  GENERAL ADVICE ON PHYSICS
'******************************************************
'
' It's advised that flipper corrections, dampeners, and general physics settings should all be updated per these
' examples as all of these improvements work together to provide a realistic physics simulation.
'
' Tutorial videos provided by Bord
' Adding nFozzy roth physics : pt1 rubber dampeners         https://youtu.be/AXX3aen06FM
' Adding nFozzy roth physics : pt2 flipper physics          https://youtu.be/VSBFuK2RCPE
' Adding nFozzy roth physics : pt3 other elements           https://youtu.be/JN8HEJapCvs
'
' Note: BallMass must be set to 1. BallSize should be set to 50 (in other words the ball radius is 25)
'
' Recommended Table Physics Settings
' | Gravity Constant             | 0.97      |
' | Playfield Friction           | 0.15-0.25 |
' | Playfield Elasticity         | 0.25      |
' | Playfield Elasticity Falloff | 0         |
' | Playfield Scatter            | 0         |
' | Default Element Scatter      | 2         |
'
' Bumpers
' | Force         | 9.5-10.5 |
' | Hit Threshold | 1.6-2    |
' | Scatter Angle | 2        |
'
' Slingshots
' | Hit Threshold      | 2    |
' | Slingshot Force    | 4-5  |
' | Slingshot Theshold | 2-3  |
' | Elasticity         | 0.85 |
' | Friction           | 0.8  |
' | Scatter Angle      | 1    |


'******************************************************
' Flippers Polarity (Select appropriate sub based on era)
'******************************************************

Dim LF
Set LF = New FlipperPolarity
Dim RF
Set RF = New FlipperPolarity

InitPolarity

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
    x.AddPt "Polarity", 1, 0.05, -5.5
    x.AddPt "Polarity", 2, 0.4, -5.5
    x.AddPt "Polarity", 3, 0.6, -5.0
    x.AddPt "Polarity", 4, 0.65, -4.5
    x.AddPt "Polarity", 5, 0.7, -4.0
    x.AddPt "Polarity", 6, 0.75, -3.5
    x.AddPt "Polarity", 7, 0.8, -3.0
    x.AddPt "Polarity", 8, 0.85, -2.5
    x.AddPt "Polarity", 9, 0.9,-2.0
    x.AddPt "Polarity", 10, 0.95, -1.5
    x.AddPt "Polarity", 11, 1, -1.0
    x.AddPt "Polarity", 12, 1.05, -0.5
    x.AddPt "Polarity", 13, 1.1, 0
    x.AddPt "Polarity", 14, 1.3, 0

    x.AddPt "Velocity", 0, 0,    1
    x.AddPt "Velocity", 1, 0.160, 1.06
    x.AddPt "Velocity", 2, 0.410, 1.05
    x.AddPt "Velocity", 3, 0.530, 1'0.982
    x.AddPt "Velocity", 4, 0.702, 0.968
    x.AddPt "Velocity", 5, 0.95,  0.968
    x.AddPt "Velocity", 6, 1.03,  0.945
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
'   Otherwise it should function exactly the same as before

Class FlipperPolarity
  Public DebugOn, Enabled
  Private FlipAt    'Timer variable (IE 'flip at 723,530ms...)
  Public TimeDelay    'delay before trigger turns off and polarity is disabled
  Private Flipper, FlipperStart, FlipperEnd, FlipperEndY, LR, PartialFlipCoef
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
      If Not IsEmpty(balls(x) ) Then
        pos = pSlope(Balls(x).x, FlipperStart, 0, FlipperEnd, 1)
      End If
    Next
  End Property

  Public Sub ProcessBalls() 'save data of balls in flipper range
    FlipAt = GameTime
    Dim x
    For x = 0 To UBound(balls)
      If Not IsEmpty(balls(x) ) Then
        balldata(x).Data = balls(x)
      End If
    Next
    PartialFlipCoef = ((Flipper.StartAngle - Flipper.CurrentAngle) / (Flipper.StartAngle - Flipper.EndAngle))
    PartialFlipCoef = abs(PartialFlipCoef-1)
  End Sub
  'Timer shutoff for polaritycorrect
  Private Function FlipperOn()
    If GameTime < FlipAt+TimeDelay Then
      FlipperOn = True
    End If
  End Function

  Public Sub PolarityCorrect(aBall)
    If FlipperOn() Then
      Dim tmp, BallPos, x, IDX, Ycoef
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
          If ballpos > 0.65 Then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)                'find safety coefficient 'ycoef' data
        End If
      Next

      If BallPos = 0 Then 'no ball data meaning the ball is entering and exiting pretty close to the same position, use current values.
        BallPos = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
        If ballpos > 0.65 Then  Ycoef = LinearEnvelope(aBall.Y, YcoefIn, YcoefOut)                        'find safety coefficient 'ycoef' data
      End If

      'Velocity correction
      If Not IsEmpty(VelocityIn(0) ) Then
        Dim VelCoef
        VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)

        If partialflipcoef < 1 Then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)

        If Enabled Then aBall.Velx = aBall.Velx*VelCoef
        If Enabled Then aBall.Vely = aBall.Vely*VelCoef
      End If

      'Polarity Correction (optional now)
      If Not IsEmpty(PolarityIn(0) ) Then
        Dim AddX
        AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR

        If Enabled Then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef)
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

RightFlipper.timerinterval = 1
Rightflipper.timerenabled = True

Sub RightFlipper_timer()
  FlipperTricks LeftFlipper, LFPress, LFCount, LFEndAngle, LFState
  FlipperTricks RightFlipper, RFPress, RFCount, RFEndAngle, RFState
  FlipperNudge RightFlipper, RFEndAngle, RFEOSNudge, LeftFlipper, LFEndAngle
  FlipperNudge LeftFlipper, LFEndAngle, LFEOSNudge,  RightFlipper, RFEndAngle
End Sub

Dim LFEOSNudge, RFEOSNudge

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

'*****************
' Maths
'*****************

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

Const FlipperCoilRampupMode = 0 '0 = fast, 1 = medium, 2 = slow (tap passes should work)

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
  FlipperPress = 0
  Flipper.eostorqueangle = EOSA
  Flipper.eostorque = EOST * EOSReturn / FReturn

  If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 0.1 Then
    Dim b, BOT
    BOT = GetBalls

    For b = 0 To UBound(BOT)
      If Distance(BOT(b).x, BOT(b).y, Flipper.x, Flipper.y) < 55 Then 'check for cradle
        If BOT(b).vely >= - 0.4 Then BOT(b).vely =  - 0.4
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

Const LiveDistanceMin = 30  'minimum distance In vp units from flipper base live catch dampening will occur
Const LiveDistanceMax = 114 'maximum distance in vp units from flipper base live catch dampening will occur (tip protection)

Sub CheckLiveCatch(ball, Flipper, FCount, parm) 'Experimental new live catch
  Dim Dir
  Dir = Flipper.startangle / Abs(Flipper.startangle)  '-1 for Right Flipper
  Dim LiveCatchBounce                                                           'If live catch is not perfect, it won't freeze ball totally
  Dim CatchTime
  CatchTime = GameTime - FCount

  If CatchTime <= LiveCatch And parm > 6 And Abs(Flipper.x - ball.x) > LiveDistanceMin And Abs(Flipper.x - ball.x) < LiveDistanceMax Then
    If CatchTime <= LiveCatch * 0.5 Then                        'Perfect catch only when catch time happens in the beginning of the window
      LiveCatchBounce = 0
    Else
      LiveCatchBounce = Abs((LiveCatch / 2) - CatchTime)    'Partial catch when catch happens a bit late
    End If

    If LiveCatchBounce = 0 And ball.velx * Dir > 0 Then ball.velx = 0
    ball.vely = LiveCatchBounce * (32 / LiveCatch) ' Multiplier for inaccuracy bounce
    ball.angmomx = 0
    ball.angmomy = 0
    ball.angmomz = 0
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
'
Sub RDampen_Timer
  Cor.Update
End Sub

'******************************************************
'****  END PHYSICS DAMPENERS
'******************************************************



'******************************************************
'   ZBOU: VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************

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

Sub AddSlingsPt(idx, aX, aY)    'debugger wrapper for adjusting flipper script in-game
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
Function RotPoint(x,y,angle)
  dim rx, ry
  rx = x*dCos(angle) - y*dSin(angle)
  ry = x*dSin(angle) + y*dCos(angle)
  RotPoint = Array(rx,ry)
End Function

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
    If gametime > 100 Then Report
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

'******************************************************
'****  END SLINGSHOT CORRECTIONS
'******************************************************

'******************************************************
'   ZFLE:  FLEEP MECHANICAL SOUNDS
'******************************************************


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
RubberFlipperSoundFactor = 0.375 / 5      'volume multiplier; must not be zero
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
  PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelExistingStatic(playsoundparams, aVol, tableobj)
  PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticLoop(playsoundparams, aVol, tableobj)
  PlaySound playsoundparams, - 1, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
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
  PlaySound playsoundparams, - 1, aVol * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
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
  PlaySound soundname, 1,aVol * VolumeDial, AudioPan(ActiveBall), 0,0,0, 1, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtBallVolM (Soundname, aVol)
  PlaySound soundname, 1,aVol * VolumeDial, AudioPan(ActiveBall), 0,0,0, 0, AudioFade(ActiveBall)
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

' Thalamus, AudioFade patched
	If tmp > 0 Then
		AudioFade = CSng(tmp ^ 5) 'was 10
	Else
		AudioFade = CSng( - (( - tmp) ^ 5) ) 'was 10
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

' Thalamus, AudioPan patched
	If tmp > 0 Then
		AudioPan = CSng(tmp ^ 5) 'was 10
	Else
		AudioPan = CSng( - (( - tmp) ^ 5) ) 'was 10
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

  FlipperCradleCollision ball1, ball2, velocity

End Sub


'///////////////////////////  DROP TARGET HIT SOUNDS  ///////////////////////////

Sub RandomSoundDropTargetReset(obj)
  PlaySoundAtLevelStatic SoundFX("Drop_Target_Reset_" & Int(Rnd * 6) + 1,DOFContactors), 1, obj
End Sub

Sub SoundDropTargetDrop(obj)
  PlaySoundAtLevelStatic ("Drop_Target_Down_" & Int(Rnd * 6) + 1), 200, obj
End Sub

'/////////////////////////////  GI AND FLASHER RELAYS  ////////////////////////////

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

'******************************************************
'****  END FLEEP MECHANICAL SOUNDS
'******************************************************


'******************************************************
' ZBRL:  BALL ROLLING AND DROP SOUNDS
'******************************************************

' Be sure to call RollingUpdate in a timer with a 10ms interval see the GameTimer_Timer() sub

ReDim rolling(tnob)
InitRolling

Dim DropCount
ReDim DropCount(tnob)

Sub InitRolling
  Dim i
  For i = 0 To tnob
    rolling(i) = False
  Next
End Sub

Sub RollingUpdate()
  Dim b
  Dim gBOT
  gBOT = GetBalls

  ' stop the sound of deleted balls
  For b = UBound(gBOT) + 1 To tnob - 1
    ' Comment the next line if you are not implementing Dyanmic Ball Shadows
    If AmbientBallShadowOn = 0 Then BallShadowA(b).visible = 0
    rolling(b) = False
    StopSound("BallRoll_" & b)
  Next

  ' exit the sub if no balls on the table
  If UBound(gBOT) =  - 1 Then Exit Sub

  ' play the rolling sound for each ball
  For b = 0 To UBound(gBOT)
    If BallVel(gBOT(b)) > 1 And gBOT(b).z < 30 Then
      rolling(b) = True
      PlaySound ("BallRoll_" & b), - 1, VolPlayfieldRoll(gBOT(b)) * BallRollVolume * VolumeDial, AudioPan(gBOT(b)), 0, PitchPlayfieldRoll(gBOT(b)), 1, 0, AudioFade(gBOT(b))
    Else
      If rolling(b) = True Then
        StopSound("BallRoll_" & b)
        rolling(b) = False
      End If
    End If

    ' Ball Drop Sounds
    If gBOT(b).VelZ <  - 1 And gBOT(b).z < 55 And gBOT(b).z > 27 Then 'height adjust for ball drop sounds
      If DropCount(b) >= 5 Then
        DropCount(b) = 0
        If gBOT(b).velz >  - 7 Then
          RandomSoundBallBouncePlayfieldSoft gBOT(b)
        Else
          RandomSoundBallBouncePlayfieldHard gBOT(b)
        End If
      End If
    End If

    If DropCount(b) < 5 Then
      DropCount(b) = DropCount(b) + 1
    End If

    ' "Static" Ball Shadows
    ' Comment the next If block, if you are not implementing the Dynamic Ball Shadows
    If AmbientBallShadowOn = 0 Then
      If gBOT(b).Z > 30 Then
        BallShadowA(b).height = gBOT(b).z - BallSize / 4    'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping the ramp
      Else
        BallShadowA(b).height = 0.1
      End If
      BallShadowA(b).Y = gBOT(b).Y + offsetY
      BallShadowA(b).X = gBOT(b).X + offsetX
      BallShadowA(b).visible = 1
    End If
  Next
End Sub

'******************************************************
'****  END BALL ROLLING AND DROP SOUNDS
'******************************************************

'******************************************************
'   ZRRL: RAMP ROLLING SFX
'******************************************************

'Ball tracking ramp SFX 1.0
'   Reqirements:
'     * Import A Sound File for each ball on the table for plastic ramps.  Call It RampLoop<Ball_Number> ex: RampLoop1, RampLoop2, ...
'     * Import a Sound File for each ball on the table for wire ramps. Call it WireLoop<Ball_Number> ex: WireLoop1, WireLoop2, ...
'     * Create a Timer called RampRoll, that is enabled, with a interval of 100
'     * Set RampBAlls and RampType variable to Total Number of Balls
' Usage:
'     * Setup hit events and call WireRampOn True or WireRampOn False (True = Plastic ramp, False = Wire Ramp)
'     * To stop tracking ball
'        * call WireRampOff
'        * Otherwise, the ball will auto remove if it's below 30 vp units
'

Dim RampMinLoops
RampMinLoops = 4

' RampBalls
' Setup:  Set the array length of x in RampBalls(x,2) Total Number of Balls on table + 1:  if tnob = 5, then RampBalls(6,2)
Dim RampBalls(7,2)
'x,0 = ball x,1 = ID, 2 = Protection against ending early (minimum amount of updates)

'0,0 is boolean on/off, 0,1 unused for now
RampBalls(0,0) = False

' RampType
' Setup: Set this array to the number Total number of balls that can be tracked at one time + 1.  5 ball multiball then set value to 6
' Description: Array type indexed on BallId and a values used to deterimine what type of ramp the ball is on: False = Wire Ramp, True = Plastic Ramp
Dim RampType(7)

Sub WireRampOn(input)
  Waddball ActiveBall, input
  RampRollUpdate
End Sub

Sub WireRampOff()
  WRemoveBall ActiveBall.ID
End Sub

' WaddBall (Active Ball, Boolean)
Sub Waddball(input, RampInput) 'This subroutine is called from WireRampOn to Add Balls to the RampBalls Array
  ' This will loop through the RampBalls array checking each element of the array x, position 1
  ' To see if the the ball was already added to the array.
  ' If the ball is found then exit the subroutine
  Dim x
  For x = 1 To UBound(RampBalls)  'Check, don't add balls twice
    If RampBalls(x, 1) = input.id Then
      If Not IsEmpty(RampBalls(x,1) ) Then Exit Sub 'Frustating issue with BallId 0. Empty variable = 0
    End If
  Next

  ' This will itterate through the RampBalls Array.
  ' The first time it comes to a element in the array where the Ball Id (Slot 1) is empty.  It will add the current ball to the array
  ' The RampBalls assigns the ActiveBall to element x,0 and ball id of ActiveBall to 0,1
  ' The RampType(BallId) is set to RampInput
  ' RampBalls in 0,0 is set to True, this will enable the timer and the timer is also turned on
  For x = 1 To UBound(RampBalls)
    If IsEmpty(RampBalls(x, 1)) Then
      Set RampBalls(x, 0) = input
      RampBalls(x, 1) = input.ID
      RampType(x) = RampInput
      RampBalls(x, 2) = 0
      'exit For
      RampBalls(0,0) = True
      RampRoll.Enabled = 1   'Turn on timer
      'RampRoll.Interval = RampRoll.Interval 'reset timer
      Exit Sub
    End If
    If x = UBound(RampBalls) Then  'debug
      Debug.print "WireRampOn error, ball queue Is full: " & vbNewLine & _
      RampBalls(0, 0) & vbNewLine & _
      TypeName(RampBalls(1, 0)) & " ID:" & RampBalls(1, 1) & "type:" & RampType(1) & vbNewLine & _
      TypeName(RampBalls(2, 0)) & " ID:" & RampBalls(2, 1) & "type:" & RampType(2) & vbNewLine & _
      TypeName(RampBalls(3, 0)) & " ID:" & RampBalls(3, 1) & "type:" & RampType(3) & vbNewLine & _
      TypeName(RampBalls(4, 0)) & " ID:" & RampBalls(4, 1) & "type:" & RampType(4) & vbNewLine & _
      TypeName(RampBalls(5, 0)) & " ID:" & RampBalls(5, 1) & "type:" & RampType(5) & vbNewLine & _
      " "
    End If
  Next
End Sub

' WRemoveBall (BallId)
Sub WRemoveBall(ID) 'This subroutine is called from the RampRollUpdate subroutine and is used to remove and stop the ball rolling sounds
  '   Debug.Print "In WRemoveBall() + Remove ball from loop array"
  Dim ballcount
  ballcount = 0
  Dim x
  For x = 1 To UBound(RampBalls)
    If ID = RampBalls(x, 1) Then 'remove ball
      Set RampBalls(x, 0) = Nothing
      RampBalls(x, 1) = Empty
      RampType(x) = Empty
      StopSound("RampLoop" & x)
      StopSound("wireloop" & x)
    End If
    'if RampBalls(x,1) = Not IsEmpty(Rampballs(x,1) then ballcount = ballcount + 1
    If Not IsEmpty(Rampballs(x,1)) Then ballcount = ballcount + 1
  Next
  If BallCount = 0 Then RampBalls(0,0) = False  'if no balls in queue, disable timer update
End Sub

Sub RampRoll_Timer()
  RampRollUpdate
End Sub

Sub RampRollUpdate()  'Timer update
  Dim x
  For x = 1 To UBound(RampBalls)
    If Not IsEmpty(RampBalls(x,1) ) Then
      If BallVel(RampBalls(x,0) ) > 1 Then ' if ball is moving, play rolling sound
        If RampType(x) Then
          PlaySound("RampLoop" & x), - 1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * VolumeDial, AudioPan(RampBalls(x,0)), 0, BallPitchV(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
          StopSound("wireloop" & x)
        Else
          StopSound("RampLoop" & x)
          PlaySound("wireloop" & x), - 1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * VolumeDial, AudioPan(RampBalls(x,0)), 0, BallPitch(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
        End If
        RampBalls(x, 2) = RampBalls(x, 2) + 1
      Else
        StopSound("RampLoop" & x)
        StopSound("wireloop" & x)
      End If
      If RampBalls(x,0).Z < 30 And RampBalls(x, 2) > RampMinLoops Then  'if ball is on the PF, remove  it
        StopSound("RampLoop" & x)
        StopSound("wireloop" & x)
        Wremoveball RampBalls(x,1)
      End If
    Else
      StopSound("RampLoop" & x)
      StopSound("wireloop" & x)
    End If
  Next
  If Not RampBalls(0,0) Then RampRoll.enabled = 0
End Sub

' This can be used to debug the Ramp Roll time.  You need to enable the tbWR timer on the TextBox
Sub tbWR_Timer()  'debug textbox
  Me.text = "on? " & RampBalls(0, 0) & " timer: " & RampRoll.Enabled & vbNewLine & _
  "1 " & TypeName(RampBalls(1, 0)) & " ID:" & RampBalls(1, 1) & " type:" & RampType(1) & " Loops:" & RampBalls(1, 2) & vbNewLine & _
  "2 " & TypeName(RampBalls(2, 0)) & " ID:" & RampBalls(2, 1) & " type:" & RampType(2) & " Loops:" & RampBalls(2, 2) & vbNewLine & _
  "3 " & TypeName(RampBalls(3, 0)) & " ID:" & RampBalls(3, 1) & " type:" & RampType(3) & " Loops:" & RampBalls(3, 2) & vbNewLine & _
  "4 " & TypeName(RampBalls(4, 0)) & " ID:" & RampBalls(4, 1) & " type:" & RampType(4) & " Loops:" & RampBalls(4, 2) & vbNewLine & _
  "5 " & TypeName(RampBalls(5, 0)) & " ID:" & RampBalls(5, 1) & " type:" & RampType(5) & " Loops:" & RampBalls(5, 2) & vbNewLine & _
  "6 " & TypeName(RampBalls(6, 0)) & " ID:" & RampBalls(6, 1) & " type:" & RampType(6) & " Loops:" & RampBalls(6, 2) & vbNewLine & _
  " "
End Sub

Function BallPitch(ball) ' Calculates the pitch of the sound based on the ball speed
  BallPitch = pSlope(BallVel(ball), 1, - 1000, 60, 10000)
End Function

Function BallPitchV(ball) ' Calculates the pitch of the sound based on the ball speed Variation
  BallPitchV = pSlope(BallVel(ball), 1, - 4000, 60, 7000)
End Function

'******************************************************
'**** END RAMP ROLLING SFX
'******************************************************


'******************************************************
'   ZFLD:  FLUPPER DOMES
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
' Example: RotateFlasher 1, 180    'where 1 is the flasher number and 180 is the angle of Z rotation

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
'   ObjTargetLevel(1) = 1
' Else
'   ObjTargetLevel(1) = 0
' End If
'   FlasherFlash1_Timer
' End Sub

' example script for rom based tables (modulated):

' SolModCallback(25)="FlashRed"
'
' Sub FlashRed(level)
' ObjTargetLevel(1) = level/255 : FlasherFlash1_Timer
' End Sub

Sub Flash1(Enabled)
  If Enabled Then
    ObjTargetLevel(1) = 1
  Else
    ObjTargetLevel(1) = 0
  End If
  FlasherFlash1_Timer
  Sound_Flash_Relay enabled, Flasherbase1
End Sub

Sub Flash2(Enabled)
  If Enabled Then
    ObjTargetLevel(2) = 1
  Else
    ObjTargetLevel(2) = 0
  End If
  FlasherFlash2_Timer
  Sound_Flash_Relay enabled, Flasherbase2
End Sub

Sub Flash3(Enabled)
  If Enabled Then
    ObjTargetLevel(3) = 1
  Else
    ObjTargetLevel(3) = 0
  End If
  FlasherFlash3_Timer
  Sound_Flash_Relay enabled, Flasherbase3
End Sub

Sub Flash4(Enabled)
  If Enabled Then
    ObjTargetLevel(4) = 1
  Else
    ObjTargetLevel(4) = 0
  End If
  FlasherFlash4_Timer
  Sound_Flash_Relay enabled, Flasherbase1
End Sub

Dim TestFlashers, TableRef, FlasherLightIntensity, FlasherFlareIntensity, FlasherBloomIntensity, FlasherOffBrightness

' *********************************************************************
TestFlashers = 0        ' *** set this to 1 to check position of flasher object      ***
Set TableRef = Table       ' *** change this, if your table has another name           ***
FlasherLightIntensity = 0.1  ' *** lower this, if the VPX lights are too bright (i.e. 0.1)     ***
FlasherFlareIntensity = 0.3  ' *** lower this, if the flares are too bright (i.e. 0.1)       ***
FlasherBloomIntensity = 0.2  ' *** lower this, if the blooms are too bright (i.e. 0.1)       ***
FlasherOffBrightness = 0.5    ' *** brightness of the flasher dome when switched off (range 0-2)  ***
' *********************************************************************

Dim ObjLevel(20), objbase(20), objlit(20), objflasher(20), objbloom(20), objlight(20), ObjTargetLevel(20)
'Dim tablewidth, tableheight : tablewidth = TableRef.width : tableheight = TableRef.height

'initialise the flasher color, you can only choose from "green", "red", "purple", "blue", "white" and "yellow"
InitFlasher 1, "yellow"
InitFlasher 2, "yellow"

' rotate the flasher with the command below (first argument = flasher nr, second argument = angle in degrees)
'   RotateFlasher 1,17
'   RotateFlasher 2,0
'   RotateFlasher 3,90
'   RotateFlasher 4,90

Sub InitFlasher(nr, col)
  ' store all objects in an array for use in FlashFlasher subroutine
  Set objbase(nr) = Eval("Flasherbase" & nr)
  Set objlit(nr) = Eval("Flasherlit" & nr)
  Set objflasher(nr) = Eval("Flasherflash" & nr)
  Set objlight(nr) = Eval("Flasherlight" & nr)
  Set objbloom(nr) = Eval("Flasherbloom" & nr)

  ' If the flasher is parallel to the playfield, rotate the VPX flasher object for POV and place it at the correct height
  If objbase(nr).RotY = 0 Then
    objbase(nr).ObjRotZ = Atn( (tablewidth / 2 - objbase(nr).x) / (objbase(nr).y - tableheight * 1.1)) * 180 / 3.14159
    objflasher(nr).RotZ = objbase(nr).ObjRotZ
    objflasher(nr).height = objbase(nr).z + 40
  End If

  ' set all effects to invisible and move the lit primitive at the same position and rotation as the base primitive
  objlight(nr).IntensityScale = 0
  objlit(nr).visible = 0
  objlit(nr).material = "Flashermaterial" & nr
  objlit(nr).RotX = objbase(nr).RotX
  objlit(nr).RotY = objbase(nr).RotY
  objlit(nr).RotZ = objbase(nr).RotZ
  objlit(nr).ObjRotX = objbase(nr).ObjRotX
  objlit(nr).ObjRotY = objbase(nr).ObjRotY
  objlit(nr).ObjRotZ = objbase(nr).ObjRotZ
  objlit(nr).x = objbase(nr).x
  objlit(nr).y = objbase(nr).y
  objlit(nr).z = objbase(nr).z
  objbase(nr).BlendDisableLighting = FlasherOffBrightness

  'rothbauerw
  'Adjust the position of the flasher object to align with the flasher base.
  'Comment out these lines if you want to manually adjust the flasher object
  If objbase(nr).roty > 135 Then
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
  objlight(nr).bulbhaloheight = objbase(nr).z - 10

  'rothbauerw
  'Assign the appropriate bloom image basked on the location of the flasher base
  'Comment out these lines if you want to manually assign the bloom images
  Dim xthird, ythird
  xthird = tablewidth / 3
  ythird = tableheight / 3
  If objbase(nr).x >= xthird And objbase(nr).x <= xthird * 2 Then
    objbloom(nr).imageA = "flasherbloomCenter"
    objbloom(nr).imageB = "flasherbloomCenter"
  ElseIf objbase(nr).x < xthird And objbase(nr).y < ythird Then
    objbloom(nr).imageA = "flasherbloomUpperLeft"
    objbloom(nr).imageB = "flasherbloomUpperLeft"
  ElseIf  objbase(nr).x > xthird * 2 And objbase(nr).y < ythird Then
    objbloom(nr).imageA = "flasherbloomUpperRight"
    objbloom(nr).imageB = "flasherbloomUpperRight"
  ElseIf objbase(nr).x < xthird And objbase(nr).y < ythird * 2 Then
    objbloom(nr).imageA = "flasherbloomCenterLeft"
    objbloom(nr).imageB = "flasherbloomCenterLeft"
  ElseIf  objbase(nr).x > xthird * 2 And objbase(nr).y < ythird * 2 Then
    objbloom(nr).imageA = "flasherbloomCenterRight"
    objbloom(nr).imageB = "flasherbloomCenterRight"
  ElseIf objbase(nr).x < xthird And objbase(nr).y < ythird * 3 Then
    objbloom(nr).imageA = "flasherbloomLowerLeft"
    objbloom(nr).imageB = "flasherbloomLowerLeft"
  ElseIf  objbase(nr).x > xthird * 2 And objbase(nr).y < ythird * 3 Then
    objbloom(nr).imageA = "flasherbloomLowerRight"
    objbloom(nr).imageB = "flasherbloomLowerRight"
  End If

  ' set the texture and color of all objects
  Select Case objbase(nr).image
    Case "dome2basewhite"
      objbase(nr).image = "dome2base" & col
      objlit(nr).image = "dome2lit" & col

    Case "ronddomebasewhite"
      objbase(nr).image = "ronddomebase" & col
      objlit(nr).image = "ronddomelit" & col

    Case "domeearbasewhite"
      objbase(nr).image = "domeearbase" & col
      objlit(nr).image = "domeearlit" & col
  End Select
  If TestFlashers = 0 Then
    objflasher(nr).imageA = "domeflashwhite"
    objflasher(nr).visible = 0
  End If
  Select Case col
    Case "blue"
      objlight(nr).color = RGB(4,120,255)
      objflasher(nr).color = RGB(200,255,255)
      objbloom(nr).color = RGB(4,120,255)
      objlight(nr).intensity = 5000

    Case "green"
      objlight(nr).color = RGB(12,255,4)
      objflasher(nr).color = RGB(12,255,4)
      objbloom(nr).color = RGB(12,255,4)

    Case "red"
      objlight(nr).color = RGB(255,32,4)
      objflasher(nr).color = RGB(255,32,4)
      objbloom(nr).color = RGB(255,32,4)

    Case "purple"
      objlight(nr).color = RGB(230,49,255)
      objflasher(nr).color = RGB(255,64,255)
      objbloom(nr).color = RGB(230,49,255)

    Case "yellow"
      objlight(nr).color = RGB(200,173,25)
      objflasher(nr).color = RGB(255,200,50)
      objbloom(nr).color = RGB(200,173,25)

    Case "white"
      objlight(nr).color = RGB(255,240,150)
      objflasher(nr).color = RGB(100,86,59)
      objbloom(nr).color = RGB(255,240,150)

    Case "orange"
      objlight(nr).color = RGB(255,70,0)
      objflasher(nr).color = RGB(255,70,0)
      objbloom(nr).color = RGB(255,70,0)
  End Select
  objlight(nr).colorfull = objlight(nr).color
  If TableRef.ShowDT And ObjFlasher(nr).RotX =  - 45 Then
    objflasher(nr).height = objflasher(nr).height - 20 * ObjFlasher(nr).y / tableheight
    ObjFlasher(nr).y = ObjFlasher(nr).y + 10
  End If
End Sub

Sub RotateFlasher(nr, angle)
  angle = ((angle + 360 - objbase(nr).ObjRotZ) Mod 180) / 30
  objbase(nr).showframe(angle)
  objlit(nr).showframe(angle)
End Sub

Sub FlashFlasher(nr)
  If Not objflasher(nr).TimerEnabled Then
    objflasher(nr).TimerEnabled = True
    objflasher(nr).visible = 1
    objbloom(nr).visible = 1
    objlit(nr).visible = 1
  End If
  objflasher(nr).opacity = 1000 * FlasherFlareIntensity * ObjLevel(nr) ^ 2.5
  objbloom(nr).opacity = 100 * FlasherBloomIntensity * ObjLevel(nr) ^ 2.5
  objlight(nr).IntensityScale = 0.5 * FlasherLightIntensity * ObjLevel(nr) ^ 3
  objbase(nr).BlendDisableLighting = FlasherOffBrightness + 10 * ObjLevel(nr) ^ 3
  objlit(nr).BlendDisableLighting = 10 * ObjLevel(nr) ^ 2
  UpdateMaterial "Flashermaterial" & nr,0,0,0,0,0,0,ObjLevel(nr),RGB(255,255,255),0,0,False,True,0,0,0,0
  If Round(ObjTargetLevel(nr),1) > Round(ObjLevel(nr),1) Then
    ObjLevel(nr) = ObjLevel(nr) + 0.3
    If ObjLevel(nr) > 1 Then ObjLevel(nr) = 1
  ElseIf Round(ObjTargetLevel(nr),1) < Round(ObjLevel(nr),1) Then
    ObjLevel(nr) = ObjLevel(nr) * 0.85 - 0.01
    If ObjLevel(nr) < 0 Then ObjLevel(nr) = 0
  Else
    ObjLevel(nr) = Round(ObjTargetLevel(nr),1)
    objflasher(nr).TimerEnabled = False
  End If
  '   ObjLevel(nr) = ObjLevel(nr) * 0.9 - 0.01
  If ObjLevel(nr) < 0 Then
    objflasher(nr).TimerEnabled = False
    objflasher(nr).visible = 0
    objbloom(nr).visible = 0
    objlit(nr).visible = 0
  End If
End Sub

Sub UpdateCaps(nr, aValue) ' nf UseVPMModSol dynamic fading version of 'FlashFlasher' (PWM Update). You can see most of the ramp up has been commented out in favor of Pinmame's curves
  if aValue > 0 then
    objflasher(nr).visible = 1 : objlit(nr).visible = 1 ': objbloom(nr).visible = 1
    Select Case nr
      Case 1:             '1,2,3 are same event, so doing this only for 1
        Flasherbloom1.ImageA="flasherbloomLowerLeft"
        Flasherbloom1.ImageB="flasherbloomLowerLeft"
        Flasherbloom1.rotx = -186.1
        Flasherbloom1.roty = 180
        Flasherbloom1.color = RGB(230,49,255)
        Flasherbloom1.amount = 750
        Flasherbloom1.visible = 1
      Case 2:
        Flasherbloom2.ImageA="flasherbloomLowerRight"
        Flasherbloom2.ImageB="flasherbloomLowerRight"
        FlasherBloom2.rotx = -6.1
        Flasherbloom2.roty = 180
        Flasherbloom2.color = RGB(255,32,4)
        Flasherbloom2.amount = 750
        Flasherbloom2.visible = 1
      Case 3:
        Flasherbloom3.ImageA="flasherbloomUpperLeft"
        Flasherbloom3.ImageB="flasherbloomUpperLeft"
        Flasherbloom3.rotx = -6.1
        Flasherbloom3.roty = 0
        Flasherbloom3.color = RGB(230,49,255)
        Flasherbloom3.amount = 750
        Flasherbloom3.visible = 1
      Case 4:
        Flasherbloom4.ImageA="flasherbloomUpperRight"
        Flasherbloom4.ImageB="flasherbloomUpperRight"
        Flasherbloom4.rotx = -186.1
        Flasherbloom4.roty = 180
        Flasherbloom4.color = RGB(255,32,4)
        Flasherbloom4.amount = 500
        Flasherbloom4.visible = 1
    End Select
  Else
    objflasher(nr).visible = 0 : objlit(nr).visible = 0 ': objbloom(nr).visible = 0
  End If

  objflasher(nr).opacity = 1000 *  FlasherFlareIntensity * aValue' * ObjLevel(nr)^2.5

  'objbloom(nr).opacity = 100 *  FlasherBloomIntensity * aValue' * ObjLevel(nr)^2.5
  Select Case nr
    Case 1:
      Flasherbloom1.opacity = 100 *  FlasherBloomIntensity * aValue '* MaxLevel^2.5
    Case 2:
      Flasherbloom2.opacity = 100 *  FlasherBloomIntensity * aValue '* MaxLevel^2.5
    Case 3:
      Flasherbloom3.opacity = 100 *  FlasherBloomIntensity * aValue '* MaxLevel^2.5
    Case 4:
      Flasherbloom4.opacity = 100 *  FlasherBloomIntensity * aValue '* MaxLevel^2.5
  End Select
  objlight(nr).IntensityScale = 0.5 * FlasherLightIntensity * aValue' * ObjLevel(nr)^3
  objbase(nr).BlendDisableLighting =  FlasherOffBrightness + 10 * aValue '* ObjLevel(nr)^3
  objlit(nr).BlendDisableLighting = 10 * aValue '* ObjLevel(nr)^2
  UpdateMaterial "Flashermaterial" & nr, 0,0,0,0,0,0,Round(aValue,1),RGB(255,255,255),0,0,False,True,0,0,0,0

End Sub


Sub FlasherFlash1_Timer()
  FlashFlasher(1)
End Sub
Sub FlasherFlash2_Timer()
  FlashFlasher(2)
End Sub
Sub FlasherFlash3_Timer()
  FlashFlasher(3)
End Sub
Sub FlasherFlash4_Timer()
  FlashFlasher(4)
End Sub

'******************************************************
'******  END FLUPPER DOMES
'******************************************************

'***************************************************************
' ZSHA: VPW DYNAMIC BALL SHADOWS by Iakki, Apophis, and Wylte
'***************************************************************

'****** INSTRUCTIONS please read ******

'****** Part A:  Table Elements ******
'
' Import the "bsrtx8" and "ballshadow" images
' Import the shadow materials file (3 sets included) (you can also export the 3 sets from this table to create the same file)
' Copy in the BallShadowA flasher set and the sets of primitives named BallShadow#, RtxBallShadow#, and RtxBall2Shadow#
' * Count from 0 up, with at least as many objects each as there can be balls, including locked balls.  You'll get an "eval" warning if tnob is higher
' * Warning:  If merging with another system (JP's ballrolling), you may need to check tnob math and add an extra BallShadowA# flasher (out of range error)
' Ensure you have a timer with a -1 interval that is always running
' Set plastic ramps DB to *less* than the ambient shadows (-11000) if you want to see the pf shadow through the ramp
' Place triggers at the start of each ramp *type* (solid, clear, wire) and one at the end if it doesn't return to the base pf
' * These can share duties as triggers for RampRolling sounds

' Create a collection called DynamicSources that includes all light sources you want to cast ball shadows
' It's recommended that you be selective in which lights go in this collection, as there are limitations:
' 1. The shadows can "pass through" solid objects and other light sources, so be mindful of where the lights would actually able to cast shadows
' 2. If there are more than two equidistant sources, the shadows can suddenly switch on and off, so places like top and bottom lanes need attention
' 3. At this time the shadows get the light on/off from tracking gilvl, so if you have lights you want shadows for that are on at different times you will need to either:
' a) remove this restriction (shadows think lights are always On)
' b) come up with a custom solution (see TZ example in script)
' After confirming the shadows work in general, use ball control to move around and look for any weird behavior

'****** End Part A:  Table Elements ******


'****** Part B:  Code and Functions ******

' *** Timer sub
' The "DynamicBSUpdate" sub should be called by a timer with an interval of -1 (framerate)
' Example timer sub:

'Sub FrameTimer_Timer()
' If DynamicBallShadowsOn Or AmbientBallShadowOn Then DynamicBSUpdate 'update ball shadows
'End Sub

' *** These are usually defined elsewhere (ballrolling), but activate here if necessary
'Const tnob = 10 ' total number of balls
'Const lob = 0  'locked balls on start; might need some fiddling depending on how your locked balls are done
'Dim tablewidth: tablewidth = Table1.width
'Dim tableheight: tableheight = Table1.height

' *** User Options - Uncomment here or move to top for easy access by players
'----- Shadow Options -----
'Const DynamicBallShadowsOn = 1   '0 = no dynamic ball shadow ("triangles" near slings and such), 1 = enable dynamic ball shadow
'Const AmbientBallShadowOn = 1    '0 = Static shadow under ball ("flasher" image, like JP's)
'                 '1 = Moving ball shadow ("primitive" object, like ninuzzu's) - This is the only one that shows up on the pf when in ramps and fades when close to lights!
'                 '2 = flasher image shadow, but it moves like ninuzzu's

' *** The following segment goes within the RollingUpdate sub, so that if Ambient...=0 and Dynamic...=0 the entire DynamicBSUpdate sub can be skipped for max performance
' ** Change gBOT to BOT if using existing getballs code
' ** Double commented lines commonly found there included for reference:

''  ' stop the sound of deleted balls
''  For b = UBound(gBOT) + 1 to tnob
'   If AmbientBallShadowOn = 0 Then BallShadowA(b).visible = 0
''    ...rolling(b) = False
''    ...StopSound("BallRoll_" & b)
''  Next
''
'' ...rolling and drop sounds...
''
''    If DropCount(b) < 5 Then
''      DropCount(b) = DropCount(b) + 1
''    End If
''
'   ' "Static" Ball Shadows
'   If AmbientBallShadowOn = 0 Then
'     BallShadowA(b).visible = 1
'     BallShadowA(b).X = gBOT(b).X + offsetX
'     If gBOT(b).Z > 30 Then
'       BallShadowA(b).height=gBOT(b).z - BallSize/4 + b/1000 'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping the ramp
'       BallShadowA(b).Y = gBOT(b).Y + offsetY + BallSize/10
'     Else
'       BallShadowA(b).height=gBOT(b).z - BallSize/2 + 1.04 + b/1000
'       BallShadowA(b).Y = gBOT(b).Y + offsetY
'     End If
'   End If

' *** Place this inside the table init, just after trough balls are added to gBOT
'
' Add balls to shadow dictionary
' For Each xx in gBOT
'   bsDict.Add xx.ID, bsNone
' Next

' *** Example RampShadow trigger subs:

'Sub ClearRampStart_hit()
' bsRampOnClear     'Shadow on ramp and pf below
'End Sub

'Sub SolidRampStart_hit()
' bsRampOn        'Shadow on ramp only
'End Sub

'Sub WireRampStart_hit()
' bsRampOnWire      'Shadow only on pf
'End Sub

'Sub RampEnd_hit()
' bsRampOff ActiveBall.ID 'Back to default shadow behavior
'End Sub


' *** Required Functions, enable these if they are not already present elswhere in your table
Function max(a,b)
  If a > b Then
    max = a
  Else
    max = b
  End If
End Function

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
Const AmbientBSFactor = 0.9  '0 To 1, higher is darker
Const AmbientMovement = 1    '1+ higher means more movement as the ball moves left and right
Const offsetX = 0        'Offset x position under ball (These are if you want to change where the "room" light is for calculating the shadow position,)
Const offsetY = 5        'Offset y position under ball (^^for example 5,5 if the light is in the back left corner)

'Dynamic (Table light sources)
Const DynamicBSFactor = 0.90  '0 To 1, higher is darker
Const Wideness = 20      'Sets how wide the dynamic ball shadows can get (20 +5 thinness is technically most accurate for lights at z ~25 hitting a 50 unit ball)
Const Thinness = 5        'Sets minimum as ball moves away from source

' *** Trim or extend these to match the number of balls/primitives/flashers on the table!  (will throw errors if there aren't enough objects)
Dim objrtx1(6), objrtx2(6)
Dim objBallShadow(6)
Dim OnPF(6)
Dim BallShadowA
BallShadowA = Array (BallShadowA0,BallShadowA1,BallShadowA2,BallShadowA3,BallShadowA4,BallShadowA5)
Dim DSSources(30), numberofsources', DSGISide(30) 'Adapted for TZ with GI left / GI right

' *** The Shadow Dictionary
Dim bsDict
Set bsDict = New cvpmDictionary
Const bsNone = "None"
Const bsWire = "Wire"
Const bsRamp = "Ramp"
Const bsRampClear = "Clear"

'Initialization
DynamicBSInit

Sub DynamicBSInit()
  Dim iii, source

  'Prepare the shadow objects before play begins
  For iii = 0 To tnob - 1
    Set objrtx1(iii) = Eval("RtxBallShadow" & iii)
    objrtx1(iii).material = "RtxBallShadow" & iii
    objrtx1(iii).z = 1 + iii / 1000 + 0.01  'Separate z for layering without clipping
    objrtx1(iii).visible = 0

    Set objrtx2(iii) = Eval("RtxBall2Shadow" & iii)
    objrtx2(iii).material = "RtxBallShadow2_" & iii
    objrtx2(iii).z = 1 + iii / 1000 + 0.02
    objrtx2(iii).visible = 0

    Set objBallShadow(iii) = Eval("BallShadow" & iii)
    objBallShadow(iii).material = "BallShadow" & iii
    UpdateMaterial objBallShadow(iii).material,1,0,0,0,0,0,AmbientBSFactor,RGB(0,0,0),0,0,False,True,0,0,0,0
    objBallShadow(iii).Z = 1 + iii / 1000 + 0.04
    objBallShadow(iii).visible = 0

    BallShadowA(iii).Opacity = 100 * AmbientBSFactor
    BallShadowA(iii).visible = 0
  Next

  iii = 0

  For Each Source In DynamicSources
    DSSources(iii) = Array(Source.x, Source.y)
    '   If Instr(Source.name , "Left") > 0 Then DSGISide(iii) = 0 Else DSGISide(iii) = 1  'Adapted for TZ with GI left / GI right
    iii = iii + 1
  Next
  numberofsources = iii
End Sub

Sub BallOnPlayfieldNow(onPlayfield, ballNum)  'Whether a ball is currently on the playfield. Only update certain things once, save some cycles
  Dim gBOT: gBOT = getballs

  If onPlayfield Then
    OnPF(ballNum) = True
    bsRampOff gBOT(ballNum).ID
    '   debug.print "Back on PF"
    UpdateMaterial objBallShadow(ballNum).material,1,0,0,0,0,0,AmbientBSFactor,RGB(0,0,0),0,0,False,True,0,0,0,0
    objBallShadow(ballNum).size_x = 5
    objBallShadow(ballNum).size_y = 4.5
    objBallShadow(ballNum).visible = 1
    BallShadowA(ballNum).visible = 0
    BallShadowA(ballNum).Opacity = 100 * AmbientBSFactor
  Else
    OnPF(ballNum) = False
    '   debug.print "Leaving PF"
  End If
End Sub

Sub DynamicBSUpdate
  Dim falloff 'Max distance to light sources, can be changed dynamically if you have a reason
  falloff = 150
  Dim ShadowOpacity1, ShadowOpacity2
  Dim s, LSd, iii
  Dim dist1, dist2, src1, src2
  Dim bsRampType
  Dim gBOT: gBOT = getballs

  'Hide shadow of deleted balls
  For s = UBound(gBOT) + 1 To tnob - 1
    objrtx1(s).visible = 0
    objrtx2(s).visible = 0
    objBallShadow(s).visible = 0
    BallShadowA(s).visible = 0
  Next

  If UBound(gBOT) < lob Then Exit Sub 'No balls in play, exit

  'The Magic happens now
  For s = lob To UBound(gBOT)
    ' *** Normal "ambient light" ball shadow
    'Layered from top to bottom. If you had an upper pf at for example 80 units and ramps even above that, your Elseif segments would be z>110; z<=110 And z>100; z<=100 And z>30; z<=30 And z>20; Else (under 20)

    'Primitive shadow on playfield, flasher shadow in ramps
    If AmbientBallShadowOn = 1 Then
      '** Above the playfield
      If gBOT(s).Z > 30 Then
        If OnPF(s) Then BallOnPlayfieldNow False, s   'One-time update
        bsRampType = getBsRampType(gBOT(s).id)
        '   debug.print bsRampType

        If Not bsRampType = bsRamp Then 'Primitive visible on PF
          objBallShadow(s).visible = 1
          objBallShadow(s).X = gBOT(s).X + (gBOT(s).X - (tablewidth / 2)) / (Ballsize / AmbientMovement) + offsetX
          objBallShadow(s).Y = gBOT(s).Y + offsetY
          objBallShadow(s).size_x = 5 * ((gBOT(s).Z + BallSize) / 80) 'Shadow gets larger and more diffuse as it moves up
          objBallShadow(s).size_y = 4.5 * ((gBOT(s).Z + BallSize) / 80)
          UpdateMaterial objBallShadow(s).material,1,0,0,0,0,0,AmbientBSFactor * (30 / (gBOT(s).Z)),RGB(0,0,0),0,0,False,True,0,0,0,0
        Else 'Opaque, no primitive below
          objBallShadow(s).visible = 0
        End If

        If bsRampType = bsRampClear Or bsRampType = bsRamp Then 'Flasher visible on opaque ramp
          BallShadowA(s).visible = 1
          BallShadowA(s).X = gBOT(s).X + offsetX
          BallShadowA(s).Y = gBOT(s).Y + offsetY + BallSize / 10
          BallShadowA(s).height = gBOT(s).z - BallSize / 4 + s / 1000 'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping the ramp
          If bsRampType = bsRampClear Then BallShadowA(s).Opacity = 50 * AmbientBSFactor
        ElseIf bsRampType = bsWire Or bsRampType = bsNone Then 'Turn it off on wires or falling out of a ramp
          BallShadowA(s).visible = 0
        End If

        '** On pf, primitive only
      ElseIf gBOT(s).Z <= 30 And gBOT(s).Z > 20 Then
        If Not OnPF(s) Then BallOnPlayfieldNow True, s
        objBallShadow(s).X = gBOT(s).X + (gBOT(s).X - (tablewidth / 2)) / (Ballsize / AmbientMovement) + offsetX
        objBallShadow(s).Y = gBOT(s).Y + offsetY
        '   objBallShadow(s).Z = gBOT(s).Z + s/1000 + 0.04    'Uncomment (and adjust If/Elseif height logic) if you want the primitive shadow on an upper/split pf

        '** Under pf, flasher shadow only
      Else
        If OnPF(s) Then BallOnPlayfieldNow False, s
        objBallShadow(s).visible = 0
        BallShadowA(s).visible = 1
        BallShadowA(s).X = gBOT(s).X + offsetX
        BallShadowA(s).Y = gBOT(s).Y + offsetY
        BallShadowA(s).height = gBOT(s).z - BallSize / 4 + s / 1000
      End If

      'Flasher shadow everywhere
    ElseIf AmbientBallShadowOn = 2 Then
      If gBOT(s).Z > 30 Then 'In a ramp
        BallShadowA(s).X = gBOT(s).X + offsetX
        BallShadowA(s).Y = gBOT(s).Y + offsetY + BallSize / 10
        BallShadowA(s).height = gBOT(s).z - BallSize / 4 + s / 1000 'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping the ramp
      ElseIf gBOT(s).Z <= 30 And gBOT(s).Z > 20 Then 'On pf
        BallShadowA(s).visible = 1
        BallShadowA(s).X = gBOT(s).X + (gBOT(s).X - (tablewidth / 2)) / (Ballsize / AmbientMovement) + offsetX
        BallShadowA(s).Y = gBOT(s).Y + offsetY
        BallShadowA(s).height = 1.04 + s / 1000
      Else 'Under pf
        BallShadowA(s).X = gBOT(s).X + offsetX
        BallShadowA(s).Y = gBOT(s).Y + offsetY
        BallShadowA(s).height = gBOT(s).z - BallSize / 4 + s / 1000
      End If
    End If

    ' *** Dynamic shadows
    If DynamicBallShadowsOn Then
      If gBOT(s).Z < 30 And gBOT(s).X < 850 Then 'Parameters for where the shadows can show, here they are not visible above the table (no upper pf) or in the plunger lane
        dist1 = falloff
        dist2 = falloff
        For iii = 0 To numberofsources - 1 'Search the 2 nearest influencing lights
          LSd = Distance(gBOT(s).x, gBOT(s).y, DSSources(iii)(0), DSSources(iii)(1)) 'Calculating the Linear distance to the Source
          If LSd < falloff And gilvl > 0 Then
            '   If LSd < dist2 And ((DSGISide(iii) = 0 And Lampz.State(100)>0) Or (DSGISide(iii) = 1 And Lampz.State(104)>0)) Then  'Adapted for TZ with GI left / GI right
            dist2 = dist1
            dist1 = LSd
            src2 = src1
            src1 = iii
          End If
        Next
        ShadowOpacity1 = 0
        If dist1 < falloff Then
          objrtx1(s).visible = 1
          objrtx1(s).X = gBOT(s).X
          objrtx1(s).Y = gBOT(s).Y
          '   objrtx1(s).Z = gBOT(s).Z - 25 + s/1000 + 0.01 'Uncomment if you want to add shadows to an upper/lower pf
          objrtx1(s).rotz = AnglePP(DSSources(src1)(0), DSSources(src1)(1), gBOT(s).X, gBOT(s).Y) + 90
          ShadowOpacity1 = 1 - dist1 / falloff
          objrtx1(s).size_y = Wideness * ShadowOpacity1 + Thinness
          UpdateMaterial objrtx1(s).material,1,0,0,0,0,0,ShadowOpacity1 * DynamicBSFactor ^ 3,RGB(0,0,0),0,0,False,True,0,0,0,0
        Else
          objrtx1(s).visible = 0
        End If
        ShadowOpacity2 = 0
        If dist2 < falloff Then
          objrtx2(s).visible = 1
          objrtx2(s).X = gBOT(s).X
          objrtx2(s).Y = gBOT(s).Y + offsetY
          '   objrtx2(s).Z = gBOT(s).Z - 25 + s/1000 + 0.02 'Uncomment if you want to add shadows to an upper/lower pf
          objrtx2(s).rotz = AnglePP(DSSources(src2)(0), DSSources(src2)(1), gBOT(s).X, gBOT(s).Y) + 90
          ShadowOpacity2 = 1 - dist2 / falloff
          objrtx2(s).size_y = Wideness * ShadowOpacity2 + Thinness
          UpdateMaterial objrtx2(s).material,1,0,0,0,0,0,ShadowOpacity2 * DynamicBSFactor ^ 3,RGB(0,0,0),0,0,False,True,0,0,0,0
        Else
          objrtx2(s).visible = 0
        End If
        If AmbientBallShadowOn = 1 Then
          'Fades the ambient shadow (primitive only) when it's close to a light
          UpdateMaterial objBallShadow(s).material,1,0,0,0,0,0,AmbientBSFactor * (1 - max(ShadowOpacity1, ShadowOpacity2)),RGB(0,0,0),0,0,False,True,0,0,0,0
        Else
          BallShadowA(s).Opacity = 100 * AmbientBSFactor * (1 - max(ShadowOpacity1, ShadowOpacity2))
        End If
      Else 'Hide dynamic shadows everywhere else, just in case
        objrtx2(s).visible = 0
        objrtx1(s).visible = 0
      End If
    End If
  Next
End Sub

' *** Ramp type definitions

Sub bsRampOnWire()
  If bsDict.Exists(ActiveBall.ID) Then
    bsDict.Item(ActiveBall.ID) = bsWire
  Else
    bsDict.Add ActiveBall.ID, bsWire
  End If
End Sub

Sub bsRampOn()
  If bsDict.Exists(ActiveBall.ID) Then
    bsDict.Item(ActiveBall.ID) = bsRamp
  Else
    bsDict.Add ActiveBall.ID, bsRamp
  End If
End Sub

Sub bsRampOnClear()
  If bsDict.Exists(ActiveBall.ID) Then
    bsDict.Item(ActiveBall.ID) = bsRampClear
  Else
    bsDict.Add ActiveBall.ID, bsRampClear
  End If
End Sub

Sub bsRampOff(idx)
  If bsDict.Exists(idx) Then
    bsDict.Item(idx) = bsNone
  End If
End Sub

Function getBsRampType(id)
  Dim retValue
  If bsDict.Exists(id) Then
    retValue = bsDict.Item(id)
  Else
    retValue = bsNone
  End If
  getBsRampType = retValue
End Function

'****************************************************************
'****  END VPW DYNAMIC BALL SHADOWS by Iakki, Apophis, and Wylte
'****************************************************************

' VR Plunger Animation
Sub TimerPlunger_Timer
  If Primary_plunger.Y < -30 then
      Primary_plunger.Y = Primary_plunger.Y + 5
  End If
End Sub

Sub TimerPlunger2_Timer
  Primary_plunger.Y = -170 + (5 * Plunger.Position) - 20
End Sub

'**********************
' Pool Light timer ....
'**********************

Sub PoolLightTimer2_timer()
  If VRRoomChoice = 2 Then
    PoolTable.Image = "PoolTableLIT"   'change pool table image to the Lit one...
    PoolLight1.visible = true
    PoolShadow.visible = true
    Randomize (21)
    PoolLightTimer2.Interval = 20 * ((7 * Rnd) + 1)  ' 20 *  a random number between 1.05 and 7.95
    PoolLightTimer2.enabled = False
    PoolLightTimer1.enabled = True
  Else
    PoolTable.Image = "PoolTableLIT"   'change pool table image to the Lit one...
    PoolLight1.visible = False
    PoolShadow.visible = False
    PoolLightTimer2.Interval = 20 * ((7 * Rnd) + 1)  ' 20 *  a random number between 1.05 and 7.95
    PoolLightTimer2.enabled = False
    PoolLightTimer1.enabled = False
  End If
End Sub

Sub PoolLightTimer1_timer()
  If VRRoomChoice = 2 Then
    PoolTable.Image = "PoolTableOff"
    PoolLight1.visible = false
    PoolShadow.visible = false
    Randomize (21)
    PoolLightTimer1.Interval = 600 * ((7 * Rnd) + 1)  ' 600 *  a random number between 1.05 and 7.95
    PoolLightTimer1.enabled = false
    PoolLightTimer2.enabled = true
  Else
    PoolTable.Image = "PoolTableOff"
    PoolLight1.visible = false
    PoolShadow.visible = false
    Randomize (21)
    PoolLightTimer1.Interval = 600 * ((7 * Rnd) + 1)  ' 600 *  a random number between 1.05 and 7.95
    PoolLightTimer1.enabled = false
    PoolLightTimer2.enabled = false
  End If
End Sub

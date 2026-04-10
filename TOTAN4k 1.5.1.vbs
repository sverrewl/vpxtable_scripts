'************************************************************
'************************************************************
'
'  Tales of the Arabian Nights / IPD No. 3824 / May, 1996 / 4 Players
'
'   Credits:
'
'   VPX - JPSalas 2015; Flupper, rothbauerw 2019
'   High res playfield and plastics - Clark Kent
'   Beta Testing and Table Reference - randr
'   DOF - arngrim
'   Fixes/physics nFozzy - Iaaki
'   Special thanks:
'     Knorr, for sounds from his soundpack
'       nFozzy, for physics techniques
'     iaakki - updated to latest Roth/NF physics code
'
'************************************************************
'************************************************************

Option Explicit
Randomize
Dim luts, lutpos

'******************************************************
'             OPTIONS
'******************************************************

Const SkillShotWall = 1       '0 - no wall
                  '1 - adds a wall to prevent ball from flying onto table
                  'Ball can fly onto table from skillshot on real machine, set to 0 for more realistic play
Const ForceSideRailsFS = False    'True - Show siderails in fullscreen mode; False - do not show
Const BallShadowOn = 1        '1 - on; 0 - off
Const BallReflections = 1     '1 - on; 0 - off
Const VolumeDial = 10       'Change volume of hit events
Const RollingSoundFactor = 1    'Change volume of rolling sounds
Const FlasherTimerInterval = 17   'synchronize flasher frequency with screen refresh rate (ms)
                  ' 17 for 60 fps, 20 for 100 fps, 13 for 75 fps, 17 for 120fps, 14 for 144 fps
luts = array("LUTtotan2022", "LUTtotan2022alt1", "LUT1on1", "LUTconsat", "LUTtotan1",  "LUTtotan2", "LUTtotan3", "LUTtotan4", "LUTtotan5","LUTtotan6","LUTrobertmstotan0_darker", "LUTblacklight", "LUTVogliadicane70", "LUTVogliadicane80", "LUTmandolin", "LUTbassgeige1", "LUTbassgeige2", "LUTbassgeigemeddark", "LUTbassgeigemeddarkwhite", "LUTbassgeigeultrdark", "LUTbassgeigeultrdarkwhite", "LUtfleep", "LUTmlager8", "LUTmlager8night" )
'lutpos = 0 : not necessary anymore, lutpos is saved from previous game session
Const EnableMagnasaveLUT = 1    ' 1 - on; 0 - off; if on then the right and left magnasave button let's you rotate all LUT's
Const Show3dScrews = 1        ' 1 - on; 0 - off; display screws as 3d primitives, not just as texture; 1 is not necessary for fullscreen
Const RedWhiteFlipperBat = 0    ' 1 - on; 0 - off; if on, flipperbats are traditional red/white, off = purple bats
Const VRRoom = 1          ' 0 = big room, 1 = small room

'******************************************************
'           STANDARD DEFINITIONS
'******************************************************

Const UseVPMModSol = 2
Const UseSolenoids = 2
Const UseLamps = 1
Const SCoin = "fx_Coin"
Dim BallMass, BallSize: Ballsize = 50 : Ballmass = 1.0

'******************************************************
'           TABLE INIT
'******************************************************

' using table width and height in script slows down the performance
dim tablewidth:  tablewidth = Table1.width
dim tableheight: tableheight = Table1.height

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Dim DesktopMode: DesktopMode = Table1.ShowDT
Dim FSSMode:     FSSMode = Table1.ShowFSS
Dim UseVPMDMD:   UseVPMDMD = DesktopMode

'******** VR room and view dependant settings *******

' VRRoom set based on RenderingMode

Dim obj, backboxflash

If RenderingMode = 2 Then                   ' VR
  Light63.TransmissionScale = 10
  for each Obj in ColDesktop : Obj.visible = 0 : next     ' hide DT lockbar and rails
  If VRRoom = 0 Then
    for each Obj in ColRoom : Obj.visible = 1 : next    ' big VR room
  Else
    for each Obj in ColRoomSmall : Obj.visible = 1 : next   ' small VR room
  End If
  backboxflash = True
Else
  If DesktopMode and not FSSmode Then             ' desktopmode
    for each Obj in ColRoom : Obj.visible = 0 : next    ' hide big VR room (should not be necessary, but to make sure)
    for each Obj in ColRoomSmall : Obj.visible = 0 : next ' hide small VR room (should not be necessary, but to make sure)
    for each Obj in ColDesktop : Obj.visible = 1 : next   ' display DT lockbar and rails
    Light003.y = Light003.y + 100
    FlasherFlash117.y = FlasherFlash117.y + 100
    Light63.FalloffPower = 8
    LampPr.material = "MetalLampGoldOffDT"
    backboxflash = False
  Else
    If FssMode Then                     ' FSS mode
      scoretext.x = -500
      primitive001.visible = True             ' show backbox
      primitive002.visible = True
      primitive003.visible = True
      DMD.visible = True
      backboxflash = True
      LeftRail.visible = True : RightRail.visible = True
    Else                          ' cabinet mode
      for each Obj in ColRoom : Obj.visible = 0 : next  ' hide big VR room (should not be necessary, but to make sure)
      for each Obj in ColDesktop : Obj.visible = 0 : next ' hide small VR room (should not be necessary, but to make sure)
      If ForceSideRailsFS Then LeftRail.visible = True : RightRail.visible = True
      backboxflash = False
    End If
  End If
End If

' ***********************

LoadVPM "03060000","WPC.vbs",3.56

Const cGameName = "totan_14"

Dim bsL, vlLock, cbLeft, cbRight, mVanishMagnet, mLockMagnet, mRampMagnet
Dim x, SpinnerBall1, SpinnerBall2, TOTBall1, TOTBall2, TOTBall3, TOTBall4, CapL1, CapL2, CapR1, CapR2

Set GICallback2 = GetRef("UpdateGI")
Const VPM_MODOUT_BULB_89_20V_DC_WPC   = 301 ' Incandescent #89/906 Bulb connected to 12V, commonly used for flashers

Sub Table1_Init
  vpmInit Me
  vpmMapLights Insertlights
  With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
    .SplashInfoLine = "Tales of the Arabian Nights - Williams 1996"
    .Games(cGameName).Settings.Value("rol") = 0 'set it to 1 to rotate the DMD to the left
    .HandleKeyboard = 0
    .ShowTitle = 0
    .ShowDMDOnly = 1
    .ShowFrame = 0
    .HandleMechanics = 0
    .Hidden = DesktopMode
    '.SetDisplayPosition 0, 0, GetPlayerHWnd 'uncomment this line if you don't see the DMD
    On Error Resume Next
    .Run GetPlayerHWnd
    If Err Then MsgBox Err.Description
    On Error Goto 0
    .Switch(22) = 1 'close coin door
    .Switch(24) = 1 'and keep it closed
  End With

  ' Nudging
  vpmNudge.TiltSwitch = 14
  vpmNudge.Sensitivity = 4
  vpmNudge.TiltObj = Array(bumper1, bumper2, bumper3, LeftSlingshot, RightSlingshot)

  TextBox001.visible = 0
  On Error Resume Next
    dim testnormalmaps : testnormalmaps = LampPr.objectSpaceNormalMap
    If Err Then
      textbox001.text = "Please use VPX 10.6 revision 3751 or higher"
      textbox001.visible = 1
    End IF

  LampPr.blenddisablelighting = 4
  sw23p.blenddisablelighting = 3
  flasherflash123.intensityscale = 0
  FlasherFlash122.intensityscale = 0
  FlasherFlash122b.intensityscale = 0
  FlasherFlash122c.intensityscale = 0
  Flasher5.intensityscale = 0
  Light87.intensityscale = 0
  Light88.intensityscale = 0
  FlasherFlash117.intensityscale = 0
  FlasherFlash119.intensityscale = 0
  FlasherFlash120.intensityscale = 0
  Light89.intensityscale = 0
  Light90.intensityscale = 0
  Light91.intensityscale = 0
  Light92.intensityscale = 0
  Light93.intensityscale = 0
  Light95.intensityscale = 0
  Light96.intensityscale = 0
  Light97.intensityscale = 0
  FlasherFlash124.intensityscale = 0
  FlasherFlash125.intensityscale = 0
  Flasherflash125b.intensityscale = 0
  FlasherFlash127.intensityscale = 0
  Light94.intensityscale = 0
  Light100.intensityscale = 0
  Light99.intensityscale = 0
  genieflash.material = "lit10"
  Light98.intensityscale = 0
  Light101.intensityscale = 0
  Light102.intensityscale = 0
  Light103.intensityscale = 0
  Light104.intensityscale = 0
  Light7.intensityscale = 0
  Light105.intensityscale = 0
  Flasherflash116.intensityscale = 0
  Flasherflash118.intensityscale = 0
  Light106.intensityscale = 0
  Flasherflash128.intensityscale = 0
  Light003.intensityscale = 0
  Light107.intensityscale = 0
  Flasher2.intensityscale = 0
  Flasher3.intensityscale = 0
  Flasher4.intensityscale = 0
  Flasher5.intensityscale = 0
  Light77.intensityscale = 0
  sw16l1.intensityscale = 0
  sw27l1.intensityscale = 0

  ' adjust timerinterval to user settting
  FlasherFlash116.TimerInterval = FlasherTimerInterval
  FlasherFlash117.TimerInterval = FlasherTimerInterval
  FlasherFlash118.TimerInterval = FlasherTimerInterval
  FlasherFlash119.TimerInterval = FlasherTimerInterval
  FlasherFlash120.TimerInterval = FlasherTimerInterval
  FlasherFlash122.TimerInterval = FlasherTimerInterval
  FlasherFlash123.TimerInterval = FlasherTimerInterval
  FlasherFlash124.TimerInterval = FlasherTimerInterval
  FlasherFlash125.TimerInterval = FlasherTimerInterval
  FlasherFlash126.TimerInterval = FlasherTimerInterval
  FlasherFlash127.TimerInterval = FlasherTimerInterval
  FlasherFlash128.TimerInterval = FlasherTimerInterval


  '************  Trough **************************
  Set TOTBall4 = sw35.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set TOTBall3 = sw34.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set TOTBall2 = sw33.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set TOTBall1 = sw32.CreateSizedballWithMass(Ballsize/2,Ballmass)

  Controller.Switch(35) = 1
  Controller.Switch(34) = 1
  Controller.Switch(33) = 1
  Controller.Switch(32) = 1

  Set mVanishMagnet = New cvpmMagnet
  With mVanishMagnet
    .InitMagnet VanishMagnet, 30
    .GrabCenter = 1
    .CreateEvents "mVanishMagnet"
  End With

  Set mLockMagnet = New cvpmMagnet
  With mLockMagnet
    .InitMagnet LockMagnet, 100
    .Solenoid = 6
    .GrabCenter = 1
    .CreateEvents "mLockMagnet"
  End With

  Set mRampMagnet = New cvpmMagnet
  With mRampMagnet
    .InitMagnet RMagnet, 50
    .Solenoid = 8
    .CreateEvents "mRampMagnet"
  End With

  ' Main Timer init
  PinMAMETimer.Interval = PinMAMEInterval
  PinMAMETimer.Enabled = 1

  ' Init other dropwalls - animations
  SolSpikerLeft 0:SolSpikerRight 0
  HideVanish.isDropped = 0

  UpdateGI 0, 1:UpdateGI 1, 1:UpdateGI 2, 1

  'Spinner
  Set SpinnerBall1 = SpinnerKick.CreateSizedballWithMass(35/2,Ballmass*1.375)
  SpinnerBall1.visible = False
  Spinnerkick.kick 0,0,0
  Spinnerkick.enabled = False

  Set SpinnerBall2 = SpinnerKick2.CreateSizedballWithMass(35/2,Ballmass*1.375)
  SpinnerBall2.visible = False
  Spinnerkick2.kick 0,0,0
  Spinnerkick2.enabled = False

  sw25wall.collidable = false
  LoopPostDiverter.isdropped = True

  Set CapL1 = CapKicker1.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set CapL2 = CapKicker1a.CreateSizedballWithMass(Ballsize/2,Ballmass)

  CapKicker1.kick 0,0,0
  CapKicker1a.kick 0,0,0


  Set CapR1 = CapKicker2.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set CapR2 = CapKicker2a.CreateSizedballWithMass(Ballsize/2,Ballmass)

  CapKicker2.kick 0,0,0
  CapKicker2a.kick 0,0,0

  CapR1.FrontDecal = "NoScratches"
  CapR2.FrontDecal = "NoScratches"
  CapL1.FrontDecal = "NoScratches"
  CapL2.FrontDecal = "NoScratches"

  Wall001.collidable = False
  Wall002.collidable = False

  If SkillShotWall = 1 Then
    WallSSBox.collidable = True
  Else
    WallSSBox.collidable = False
  End If


  If RedWhiteFlipperBat = 1 Then
    leftflip.image = "flipperbatother"
    rightflip.image = "flipperbatother"
  End If

  dim x
  x = LoadValue(cGameName, "LUTPOS") : If x <> "" Then lutpos = Cint(x) Else Lutpos = 0

  table1.ColorGradeImage = luts(lutpos)
End Sub

Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub


'******************************************************
'             KEYS
'******************************************************

Sub table1_KeyDown(ByVal Keycode)
  dim tekst
  If keycode = LeftTiltKey Then Nudge 90, 2:PlaySound SoundFX("fx_nudge",0), 0, 1, -0.1, 0.25
  If keycode = RightTiltKey Then Nudge 270, 2:PlaySound SoundFX("fx_nudge",0), 0, 1, 0.1, 0.25
  If keycode = CenterTiltKey Then Nudge 0, 2:PlaySound SoundFX("fx_nudge",0), 0, 1, 0, 0.25
  If keycode = PlungerKey Then PlaySoundAt "fx_PlungerPull", Plunger:Plunger.Pullback
  If keycode = RightMagnaSave and EnableMagnasaveLUT = 1 then
    lutpos = lutpos + 1 : If lutpos > ubound(luts) Then lutpos = 0 : end if
        call myChangeLut
    playsound "LutChange"
  End if

  If keycode = LeftMagnaSave and EnableMagnasaveLUT = 1 then
    lutpos = lutpos - 1 : If lutpos < 0 Then lutpos = ubound(luts) : end if
        call myChangeLut
    playsound "LutChange2"
    end if

  If keycode = LeftFlipperKey Then FlipperActivate LeftFlipper, LFPress
  If keycode = RightFlipperKey Then FlipperActivate RightFlipper, RFPress

  If keycode = LeftFlipperKey Then
      VR_CabFlipperLeft.X = VR_CabFlipperLeft.X + 10
  End If

  If keycode = RightFlipperKey Then
      VR_CabFlipperRight.X = VR_CabFlipperRight.X - 10
  End If

  If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub table1_KeyUp(ByVal Keycode)
  If keycode = PlungerKey Then PlaySoundAt "fx_plunger",Plunger:Plunger.Fire

  If keycode = LeftFlipperKey Then
    FlipperDeActivate LeftFlipper, LFPress
  End If
  If keycode = RightFlipperKey Then
    FlipperDeActivate RightFlipper, RFPress
  End If

  If keycode = LeftFlipperKey Then
      VR_CabFlipperLeft.X = VR_CabFlipperLeft.X - 10
  End If

  If keycode = RightFlipperKey Then
      VR_CabFlipperRight.X = VR_CabFlipperRight.X + 10
  End If

  If vpmKeyUp(keycode) Then Exit Sub
End Sub

'******************************************************
'             SOLENOIDS
'******************************************************

SolCallback(1) = "SolSpikerLeft"
SolCallback(2) = "SolSpikerRight"
SolCallback(3) = "SolVanishDrop"
SolCallback(4) = "SolLockRelease"
SolCallback(5) = "SolBazaarKick"
SolCallback(6) = "SolLockMagnet"
SolCallback(7) = "vpmSolSound SoundFX(""fx_Knocker"",DOFKnocker),"
SolCallback(8) = "SolRampMagnet"
SolCallback(9) = "SolRelease"
SolCallback(15) = "SolLeftKicker"
SolCallback(21) = "RampDiverter"
SolCallback(34) = "SolPlayFDiv"
SolCallback(35) = "SolVanishMagnet"
SolCallback(36) = "SolLoopDiv"

'******************************************************
'         INLANE CAGES
'******************************************************



Sub sw36_Hit:Controller.Switch(36) = 1:End Sub
Sub sw36_Unhit:Controller.Switch(36) = 0:End Sub

Sub sw37_Hit:Controller.Switch(37) = 1:End Sub
Sub sw37_Unhit:Controller.Switch(37) = 0:End Sub

Sub SolSpikerLeft(Enabled)
  If Enabled Then
    sw36.timerenabled = true
  Else
    sw36.timerenabled = false
    vpmSolWall Array(s11, s12, s13, s14, s15, s16, s17, s18), False, NOT Enabled
    leftcage.z = -69
    StopSound "fx_cage-hold"
    PlaySoundAt SoundFX("fx_cage-off",DOFContactors), sw36
  End If
End Sub

Sub SolSpikerRight(Enabled)
  If Enabled Then
    sw37.timerenabled = true
  Else
    sw37.timerenabled = false
    vpmSolWall Array(s21, s22, s23, s24, s25, s26, s27, s28), False, NOT Enabled
    rightcage.z = -69
    StopSound "fx_cage-hold"
    PlaySoundAt SoundFX("fx_cage-off",DOFContactors), sw37
  End If
End Sub

sw36.timerinterval = 10
sw37.timerinterval = 10

Sub sw36_timer()
  Dim BOT, b, RaiseCage
  BOT = GetBalls
  RaiseCage = True

  For b = 0 to UBound(BOT)
    If InCircle(BOT(b).x, BOT(b).y, sw36a.x, sw36a.y, 67) and Not InCircle(BOT(b).x, BOT(b).y, sw36a.x, sw36a.y, 40) Then
      RaiseCage = False
    End If
  Next

  If RaiseCage Then
    leftcage.z = 0
    vpmSolWall Array(s11, s12, s13, s14, s15, s16, s17, s18), False, NOT True
    PlaySoundAt SoundFX("fx_cage-on",DOFContactors),sw36
    PlaySoundAtLoop SoundFX("fx_cage-hold",DOFShaker), sw36
    me.timerenabled = false
  End If
End Sub

Sub sw37_timer()
  Dim BOT, b, RaiseCage
  BOT = GetBalls
  RaiseCage = True

  For b = 0 to UBound(BOT)
    If InCircle(BOT(b).x, BOT(b).y, sw37a.x, sw37a.y, 67) and Not InCircle(BOT(b).x, BOT(b).y, sw37a.x, sw37a.y, 40) Then
      RaiseCage = False
    End If
  Next

  If RaiseCage Then
    rightcage.z = 0
    vpmSolWall Array(s21, s22, s23, s24, s25, s26, s27, s28), False, NOT True
    PlaySoundAt SoundFX("fx_cage-on",DOFContactors), sw37
    PlaySoundAtLoop SoundFX("fx_cage-hold",DOFShaker), sw37
    me.timerenabled = false
  End If
End Sub


Function InCircle(x, y, cirx, ciry, radius)
  If sqr((x-cirx)^2+(y-ciry)^2) < radius Then
    InCircle = True
  Else
    InCircle = False
  End If
End Function

'******************************************************
'         VANISH HOLE
'******************************************************

Dim VanishHoleX, VanishHoleY, VanishDir
VanishHoleX = 385
VanishHoleY = 491

' When vanish hole opens move ball into kicker
Sub SolVanishDrop(enabled)
  HideVanish.Isdropped = enabled
  If enabled Then
    PlaySoundAt SoundFX("fx_Solenoid",DOFContactors), VanishMagnet
    VanishDir = -1
    VanishHole.timerenabled = true
  Else
    VanishDir = 1
    VanishHole.timerenabled = true
  End If
End Sub

VanishHole.timerinterval = 10

Sub VanishHole_Timer()
  vanishholemagnet.transy = vanishholemagnet.transy + VanishDir * 10
  If vanishholemagnet.transy < -120 Then
    vanishholemagnet.transy = -120
    me.timerenabled = false
  Elseif vanishholemagnet.transy > 0 Then
    vanishholemagnet.transy = 0
    me.timerenabled = false
  End If
End Sub

Sub VanishHole_Hit
  vpmTimer.PulseSw 12:
  PlaySoundAtBall "fx_balldrop"
End Sub

' Magnet power is pulsed so wait before turning power off
Sub SolVanishMagnet(enabled)
  HideVanish.TimerEnabled = Not enabled
  If enabled Then
    mVanishMagnet.MagnetOn = True
    PlaySoundAtLoop SoundFx("fx_magnet",DOFShaker), VanishMagnet
  Else
    StopSound "fx_magnet"
  End If
End Sub

' Magnet is turned off
' Sends balls in random direction
Sub HideVanish_Timer
  Dim dir, speed, ball
  For Each ball In mVanishMagnet.Balls
    With ball
      If(.X - VanishHoleX) ^2 + (.Y - VanishHoleY) ^2 < 15 * 15 Then
        dir = Rnd * 6.28:speed = 15 + Rnd * 5
        .VelX = speed * Sin(dir): .VelY = speed * Cos(dir)
      End If
    End With
  Next
  Me.TimerEnabled = False:mVanishMagnet.MagnetOn = False
End Sub

'******************************************************
'           LOCK
'******************************************************

Sub LockMagnet_hit()
  If mLockMagnet.MagnetOn Then
    activeball.vely = activeball.vely/10
    activeball.velx = activeball.velx/10
  End If
End Sub

Sub LockMagnet1_unhit()
  If mLockMagnet.MagnetOn Then
    activeball.vely = 0
    activeball.velx = 0
  End If
End Sub

Sub SolLockMagnet(enabled)
  If enabled Then
    PlaySoundAtLoop SoundFx("fx_magnet",DOFShaker), VanishMagnet
  Else
    StopSound "fx_magnet"
  End If
End Sub

Sub sw67_hit:Controller.switch(67)=1:PlaySoundAtBall "fx_sensor":End Sub
Sub sw67_unhit:Controller.switch(67)=0:End Sub

Sub sw68_hit:Controller.switch(68)=1:PlaySoundAtBall "fx_sensor":End Sub
Sub sw68_unhit:Controller.switch(68)=0:End Sub

Sub Lock1_hit:PlaySoundAtBall"fx_kicker-enter":Controller.switch(66)=1:End Sub

Sub SolLockRelease(enabled)
  controller.Switch(66) = 0
  If lock1.BallCntOver = 0 Then
    PlaySoundAt SoundFX("fx_Solenoid",DOFContactors), lock1
  Else
    PlaySoundAt SoundFX("fx_kicker2",DOFContactors), lock1
  End If
  Lock1.kick 160,10
  if enabled then
    wall001.collidable = True
    Wall002.collidable = True
  Else
    wall001.collidable = False
    Wall002.collidable = False
  End If

End Sub

'******************************************************
'         BAZAAR SCOOP
'******************************************************

Sub sw25_Hit()
  sw25wall.collidable = true
  PlaySoundAtVol "fx_hole-enter", 0.3, sw25
  Controller.Switch(25) = 1
End Sub

Sub SolBazaarKick(Enabled)
  If Enabled Then
    sw25.kickz 192 + Rnd*1,17.5, 0,126
    sw25wall.collidable = false
    PlaySoundAtVol SoundFX("fx_Popper",DOFContactors),0.3, sw25
    Controller.Switch(25) = 0
  End If
End Sub

'******************************************************
'         RAMP MAGNET
'******************************************************

Sub SolRampMagnet(Enabled)
  Dim BOT, b

  If enabled Then
    PlaySoundAtLoop SoundFx("fx_magnet",DOFShaker), VanishMagnet
  Else
    StopSound "fx_magnet"
    BOT = GetBalls

    For b = 0 to UBound(BOT)
      If(BOT(b).X - RMagnet.x)^2 + (BOT(b).Y - RMagnet.Y)^2 < 15 * 15 Then
        BOT(b).VelX = 5
        BOT(b).VelY = 2
      End If
    Next
  end if
End Sub

'******************************************************
'           TROUGH
'******************************************************

Sub sw34_Hit():Controller.Switch(34) = 1:UpdateTrough:End Sub
Sub sw34_UnHit():Controller.Switch(34) = 0:UpdateTrough:End Sub
Sub sw33_Hit():Controller.Switch(33) = 1:UpdateTrough:End Sub
Sub sw33_UnHit():Controller.Switch(33) = 0:UpdateTrough:End Sub
Sub sw32_Hit():Controller.Switch(32) = 1:UpdateTrough:End Sub
Sub sw32_UnHit():Controller.Switch(32) = 0:UpdateTrough:End Sub

Sub UpdateTrough()
  UpdateTroughTimer.Interval = 300
  UpdateTroughTimer.Enabled = 1
End Sub

Sub UpdateTroughTimer_Timer()
  If sw32.BallCntOver = 0 Then sw33.kick 60, 9
  If sw33.BallCntOver = 0 Then sw34.kick 60, 9
  If sw34.BallCntOver = 0 Then sw35.kick 60, 20
  Me.Enabled = 0
End Sub

'******************************************************
'         DRAIN & RELEASE
'******************************************************

Sub sw35_Hit() 'Drain
  UpdateTrough
  Controller.Switch(35) = 1
  PlaySoundAtVol "fx_drain", 0.3, sw35
End Sub

Sub sw35_UnHit()  'Drain
  Controller.Switch(35) = 0
End Sub

Sub SolRelease(enabled)
  If enabled Then
    If sw32.BallCntOver = 0 Then
      PlaySoundAt SoundFX("fx_Solenoid",DOFContactors), sw32
    Else
      PlaySoundAtVol SoundFX("fx_ballrel",DOFContactors), 0.5, sw32
      vpmTimer.PulseSw 31
    End If
    sw32.kick 60, 9
  End If
End Sub


'******************************************************
'         LEFT KICKER
'******************************************************

Sub sw38_hit
  PlaySoundAtBall "fx_kicker-enter"
  Controller.switch(38)=1
End Sub

Sub SolLeftKicker(enabled)
  controller.Switch(38) = 0
  If sw38.BallCntOver = 0 Then
    PlaySoundAt SoundFX("fx_Solenoid",DOFContactors), sw38
  Else
    PlaySoundAt SoundFX("fx_kicker2",DOFContactors), sw38
  End If
  sw38.kick 168,18
End Sub


'******************************************************
'         RAMP DIVERTER
'******************************************************

Dim RampDivPos, RampDivDir
RampDivPos = 3
RampDivDir = 4 'this is the direction and the speed
RampDiv.Isdropped = 0

Sub RampDiverter(Enabled)
  PlaySoundAt SoundFX("fx_diverter",DOFContactors), RRHit7
  If Enabled Then
    RampDivDir = -4
    RampDiv.TimerEnabled = 1
  Else
    RampDivDir = 4
    RampDiv.TimerEnabled = 1
  End If
End Sub

Sub RampDiv_Timer
  RampDivP.ObjRotZ = RampDivPos
  RampDivP1.ObjRotZ = RampDivPos
  RampDivPos = RampDivPos + RampDivDir
  If RampDivPos < -21 Then
    RampDivPos = -21
    RampDiv.IsDropped = 1
    Me.TimerEnabled = 0
    Exit Sub
  End If
  If RampDivPos > 1 Then
    RampDivP.ObjRotz = 1
    RampDivP1.ObjRotz = 1
    RampDiv.IsDropped = 0
    Me.TimerEnabled = 0
    Exit Sub
  End If
End Sub

'******************************************************
'         PLAYFIELD DIVERTER
'******************************************************

Sub SolPlayFDiv(Enabled)
  vpmSolDiverter PlayFDiv, False, Enabled
  PlaySoundAt SoundFX("fx_diverter",DOFContactors), PlayFDiv
End Sub


'******************************************************
'         LOOP DIVERTER
'******************************************************

Sub SolLoopDiv(Enabled)
  vpmSolWall LoopPostDiverter,False,Not Enabled
  If Enabled Then
    PlaySoundAt SoundFX("fx_solenoid",DOFContactors), Bumper2
  Else
    PlaySoundAt SoundFX("fx_soloff",DOFContactors), Bumper2
  End If
End Sub


'******************************************************
'         FLIPPERS
'******************************************************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Const cSingleLFlip = False
Const cSingleRFlip = False

Sub SolLFlipper(Enabled)
  If Enabled Then
    PlaySoundAt SoundFX("fx_flipperupL",DOFFlippers), LeftFlipper
    LF.fire  'LeftFlipper.RotateToEnd
  Else
  if leftflipper.currentangle < leftflipper.startangle - 5 then
    PlaySoundAt SoundFX("fx_flipperdownL",DOFFlippers), LeftFlipper
  end if
    LeftFlipper.RotateToStart
  End If
End Sub

Sub SolRFlipper(Enabled)
   If Enabled Then
    PlaySoundAt SoundFX("fx_flipperupR",DOFFlippers), RightFlipper
    RF.fire  'RightFlipper.RotateToEnd
  Else
  if RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
    PlaySoundAt SoundFX("fx_flipperdownR",DOFFlippers), RightFlipper
  end if
    RightFlipper.RotateToStart
  End If
End Sub

Sub LeftFlipper_Collide(parm)
  CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
' LeftFlipperCollide parm
  PlaySoundAtBallVol "fx_rubber_flipper",  Vol(ActiveBall)*VolumeDial
End Sub

Sub RightFlipper_Collide(parm)
  CheckLiveCatch Activeball, RightFlipper, RFCount, parm
' RightFlipperCollide parm
  PlaySoundAtBallVol "fx_rubber_flipper",  Vol(ActiveBall)*VolumeDial
End Sub

'If OldFlippers = 0 then
' rightflipper.timerenabled=True
' RightFlipper.timerinterval=1
'end if
'
'sub RightFlipper_timer()
'
' If leftflipper.currentangle = leftflipper.endangle and LFPress = 1 then
'   leftflipper.eostorqueangle = EOSAnew
'   leftflipper.eostorque = EOSTnew
'   LeftFlipper.rampup = EOSRampup
'   leftflipper.Elasticity = 0.1
'   if leftflipper.endangle < LFEndAngle then
'     leftflipper.endangle = leftflipper.endangle + 0.2
'   Else
'     leftflipper.Elasticity = 0.83
'   end if
' elseif leftflipper.currentangle > leftflipper.startangle - 0.05  Then
'   leftflipper.rampup = SOSRampup
'   leftflipper.endangle = LFEndAngle - 3
'   leftflipper.Elasticity = 0.83
' elseif leftflipper.currentangle > leftflipper.endangle + 0.01 Then
'   leftflipper.eostorque = EOST
'   leftflipper.eostorqueangle = EOSA
'   LeftFlipper.rampup = Frampup
'   leftflipper.Elasticity = 0.83
' end if
'
' If rightflipper.currentangle = rightflipper.endangle and RFPress = 1 then
'   rightflipper.eostorqueangle = EOSAnew
'   rightflipper.eostorque = EOSTnew
'   RightFlipper.rampup = EOSRampup
'   rightflipper.Elasticity = 0.1
'   if rightflipper.endangle > RFEndAngle then
'     rightflipper.endangle = rightflipper.endangle - 0.2
'   Else
'     rightflipper.Elasticity = 0.83
'   end if
'
' elseif rightflipper.currentangle < rightflipper.startangle + 0.05 Then
'   rightflipper.rampup = SOSRampup
'   rightflipper.endangle = RFEndAngle + 3
'   rightflipper.Elasticity = 0.83
' elseif rightflipper.currentangle < rightflipper.endangle - 0.01 Then
'   rightflipper.eostorque = EOST
'   rightflipper.eostorqueangle = EOSA
'   RightFlipper.rampup = Frampup
'   rightflipper.Elasticity = 0.83
' end if

' If leftflipper.currentangle <> leftflipper.startangle and leftflipper.currentangle <> leftflipper.endangle and leftflipper.currentangle < leftflipper.endangle + 0.08 and leftflipper.currentangle > leftflipper.endangle + 0.01 Then
'   debug.print leftflipper.currentangle & " " & leftflipper.eostorque & " " & LeftFlipper.eostorqueangle
'   PlaySoundAtVol "fx_knocker", 0.1, LeftFlipper
' End If
'end sub

'dim LFPress, RFPress, EOST, EOSA, EOSTnew, EOSAnew
'dim FStrength, Frampup, EOSRampup, SOSRampup
'dim RFEndAngle, LFEndAngle
'
'EOST = leftflipper.eostorque
'EOSA = leftflipper.eostorqueangle
'If oldphysics = 1 Then
' LeftFlipper.strength = Leftflipper.strength - 100
'End If
'FStrength = LeftFlipper.strength
'Frampup = LeftFlipper.rampup
'EOSTnew = 1.0 'FEOST
'EOSAnew = 0.2
'EOSRampup = 1.5
'SOSRampup = 8.5
'
'LFEndAngle = Leftflipper.endangle
'RFEndAngle = RightFlipper.endangle

'******************************************************
'         SLINGSHOTS
'******************************************************

Sub LeftSlingShot_Slingshot
  PlaySoundAt SoundFx("fx_slingshot",DOFContactors), LSling
  leftsling.PlayAnim 0,0.2
  vpmTimer.PulseSw 51
End Sub

Sub RightSlingShot_Slingshot
  PlaySoundAt SoundFx("fx_slingshot",DOFContactors), RSling
  rightsling.PlayAnim 0,0.2
  vpmTimer.PulseSw 52
End Sub

'******************************************************
'         BUMPERS
'******************************************************

Sub Bumper1_Hit:vpmTimer.PulseSw 53:PlaySoundAt SoundFX("fx_bumper",DOFContactors), Bumper1:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 55:PlaySoundAt SoundFX("fx_bumper",DOFContactors), Bumper2:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 54:PlaySoundAt SoundFX("fx_bumper",DOFContactors), Bumper3:End Sub

'******************************************************
'         TRIGGERS & TARGETS
'******************************************************

' Shooter Lane
Sub sw18_Hit:Controller.Switch(18) = 1:PlaySoundAtBall "fx_sensor":End Sub
Sub sw18_Unhit:Controller.Switch(18) = 0:End Sub

' Left Inlane and Outlane
Sub sw16_Hit:vpmTimer.PulseSw 16:PlaySoundAtBall "fx_sensor":sw16p.transz=3:End Sub
Sub sw16_Unhit:sw16p.transz=0:End Sub
Sub sw26_Hit:vpmTimer.PulseSw 26:PlaySoundAtBall "fx_sensor":End Sub

' Right Inlane and Outlane
Sub sw17_Hit:vpmTimer.PulseSw 17:PlaySoundAtBall "fx_sensor":End Sub
Sub sw27_Hit:vpmTimer.PulseSw 27:PlaySoundAtBall "fx_sensor":sw27p.transz=3:End Sub
Sub sw27_Unhit:sw27p.transz=0:End Sub

' Harem Sneak
Sub HaremSneak_Hit:vpmTimer.PulseSw 11:PlaySound "fx_haremdrop", 0, 1, pan(ActiveBall):End Sub

' Ramp Entrance
Sub Gate3_Hit:
  vpmTimer.PulseSw 15
  If ActiveBall.VelY < 0 Then 'on the way up
    PlaySoundAtBall "fx_rrenter"
  Else
    PlaySoundAtBall "fx_gate"
  End If
End Sub

' Wire Ramp
Sub sw28_Hit:vpmTimer.PulseSw 28:PlaySoundAtBall "fx_sensor":PlaySoundAtBallVol "fx_metalrolling", 0.5:End Sub

' Left Ramp
Sub sw41_Hit:vpmTimer.PulseSw 41:PlaySoundAtBall "fx_sensor":End Sub

' Left Loop
Sub sw43_Hit:Controller.Switch(43) = 1:PlaySoundAtBall "fx_sensor":End Sub
Sub sw43_Unhit:Controller.Switch(43) = 0:End Sub

' Left Inner Loop
Sub sw44_Hit:vpmTimer.PulseSw 44:PlaySoundAtBall "fx_sensor":End Sub

' Right Inner Loop
Sub sw45_Hit:vpmTimer.PulseSw 45:PlaySoundAtBall "fx_sensor":End Sub

' Right Ramp Made
Sub sw47_Hit:poolSwitch.roty=-5:vpmTimer.PulseSw 47:PlaySoundAtBall "fx_sensor":End Sub
Sub sw47_unHit:poolSwitch.roty=0:End Sub

Dim TargetDelay,TargetRotz, SkillRotxy
TargetDelay = 50
TargetRotz = -7
SkillRotxy = -1.25


' Skillshot

Sub sw63_Hit(ball)
  if sw63p.objrotx = 0 then
    vpmTimer.PulseSw 63
    PlaySoundAt "fx_SkillShot-Falling", ball
    sw63p.objrotx = skillrotxy
    sw63p.objroty = skillrotxy
    sw63p1.objrotx = skillrotxy
    sw63p1.objroty = skillrotxy
  end if
End Sub

Sub sw63_unHit()
  If sw63p.objrotx = skillrotxy then
    sw63p.objrotx=skillrotxy/2:sw63p.objroty=skillrotxy/2
    sw63p1.objrotx=skillrotxy/2:sw63p1.objroty=skillrotxy/2
    vpmTimer.AddTimer TargetDelay,"sw63p.objrotx=0:sw63p.objroty=0:sw63p1.objrotx=0:sw63p1.objroty=0'"
  end if
End Sub

Sub sw64_Hit(ball)
  if sw64p.objrotx = 0 then
    vpmTimer.PulseSw 64
    PlaySoundAt "fx_SkillShot-Falling", ball
    sw64p.objrotx = skillrotxy
    sw64p.objroty = skillrotxy
    sw64p1.objrotx = skillrotxy
    sw64p1.objroty = skillrotxy
  end if
End Sub

Sub sw64_unHit()
  If sw64p.objrotx = skillrotxy then
    sw64p.objrotx=skillrotxy/2:sw64p.objroty=skillrotxy/2
    sw64p1.objrotx=skillrotxy/2:sw64p1.objroty=skillrotxy/2
    vpmTimer.AddTimer TargetDelay,"sw64p.objrotx=0:sw64p.objroty=0:sw64p1.objrotx=0:sw64p1.objroty=0'"
  end if
End Sub


Sub sw65_Hit(ball)
  if sw65p.objrotx = 0 then
    vpmTimer.PulseSw 65
    PlaySoundAt "fx_SkillShot-Falling", ball
    sw65p.objrotx = skillrotxy
    sw65p.objroty = skillrotxy
    sw65p1.objrotx = skillrotxy
    sw65p1.objroty = skillrotxy
  end if
End Sub

Sub sw65_unHit()
  If sw65p.objrotx = skillrotxy then
    sw65p.objrotx=skillrotxy/2:sw65p.objroty=skillrotxy/2
    sw65p1.objrotx=skillrotxy/2:sw65p1.objroty=skillrotxy/2
    vpmTimer.AddTimer TargetDelay,"sw65p.objrotx=0:sw65p.objroty=0:sw65p1.objrotx=0:sw65p1.objroty=0'"
  end if
End Sub


' Genie Standup Target
Sub sw23_Hit
  TargetBouncer Activeball, 1
  vpmTimer.PulseSw 23
  PlaySoundAtBallVol SoundFX("fx_target",DOFContactors), Vol(ActiveBall)*VolumeDial
  sw23p.rotz=TargetRotz
  sw23p1.rotz=TargetRotz
  vpmTimer.AddTimer TargetDelay,"sw23p.rotz=Targetrotz/2:sw23p1.rotz=Targetrotz/2'"
  vpmTimer.AddTimer TargetDelay+TargetDelay,"sw23p.rotz=0:sw23p1.rotz=0'"
End Sub

' Mini Standup Targets
Sub sw46_Hit
  TargetBouncer Activeball, 1.2
  vpmTimer.PulseSw 46
  PlaySoundAtBallVol SoundFX("fx_target",DOFContactors), Vol(ActiveBall)*VolumeDial
  DOF 101, DOFPulse
  sw46p.rotz=TargetRotz
  sw46p1.rotz=TargetRotz
  vpmTimer.AddTimer TargetDelay,"sw46p.rotz=Targetrotz/2:sw46p1.rotz=Targetrotz/2'"
  vpmTimer.AddTimer TargetDelay+TargetDelay,"sw46p.rotz=0:sw46p1.rotz=0'"
End Sub

Sub sw46b_Hit
  TargetBouncer Activeball, 1.2
  vpmTimer.PulseSw 46
  PlaySoundAtBallVol SoundFX("fx_target",DOFContactors), Vol(ActiveBall)*VolumeDial
  DOF 102, DOFPulse
  sw46bp.rotz=TargetRotz
  sw46bp1.rotz=TargetRotz
  vpmTimer.AddTimer TargetDelay,"sw46bp.rotz=Targetrotz/2:sw46bp1.rotz=Targetrotz/2'"
  vpmTimer.AddTimer TargetDelay+TargetDelay,"sw46bp.rotz=0:sw46bp1.rotz=0'"
End Sub

' Right Captive Target
Sub sw48_Hit
  vpmTimer.PulseSw 48
  PlaySoundAtBallVol SoundFX("fx_target",DOFContactors), Vol(ActiveBall)*VolumeDial
  sw48p.rotz=TargetRotz
  sw48p1.rotz=TargetRotz
  vpmTimer.AddTimer TargetDelay,"sw48p.rotz=Targetrotz/2:sw48p1.rotz=Targetrotz/2'"
  vpmTimer.AddTimer TargetDelay+TargetDelay,"sw48p.rotz=0:sw48p1.rotz=0'"
End Sub

' Left Captive Target
Sub sw58_Hit
  vpmTimer.PulseSw 58
  PlaySoundAtBallVol SoundFX("fx_target",DOFContactors), Vol(ActiveBall)*VolumeDial
  sw58p.rotz=TargetRotz
  sw58p1.rotz=TargetRotz
  vpmTimer.AddTimer TargetDelay,"sw58p.rotz=Targetrotz/2:sw58p1.rotz=Targetrotz/2'"
  vpmTimer.AddTimer TargetDelay+TargetDelay,"sw58p.rotz=0:sw58p1.rotz=0'"
End Sub

' Left Target Bank
Sub sw61a_Hit
  TargetBouncer Activeball, 1
  vpmTimer.PulseSw 61
  PlaySoundAtBallVol SoundFX("fx_target",DOFContactors), Vol(ActiveBall)*VolumeDial
  sw61ap.rotz=TargetRotz
  vpmTimer.AddTimer TargetDelay,"sw61ap.rotz=Targetrotz/2'"
  vpmTimer.AddTimer TargetDelay+TargetDelay,"sw61ap.rotz=0'"
End Sub

Sub sw61b_Hit
  TargetBouncer Activeball, 1
  vpmTimer.PulseSw 61
  PlaySoundAtBallVol SoundFX("fx_target",DOFContactors), Vol(ActiveBall)*VolumeDial
  sw61bp.rotz=TargetRotz
  vpmTimer.AddTimer TargetDelay,"sw61bp.rotz=Targetrotz/2'"
  vpmTimer.AddTimer TargetDelay+TargetDelay,"sw61bp.rotz=0'"
End Sub

Sub sw61c_Hit
  TargetBouncer Activeball, 1
  vpmTimer.PulseSw 61
  PlaySoundAtBallVol SoundFX("fx_target",DOFContactors), Vol(ActiveBall)*VolumeDial
  sw61cp.rotz=TargetRotz
  vpmTimer.AddTimer TargetDelay,"sw61cp.rotz=Targetrotz/2'"
  vpmTimer.AddTimer TargetDelay+TargetDelay,"sw61cp.rotz=0'"
End Sub

' Right Target Bank
Sub sw62a_Hit
  TargetBouncer Activeball, 1
  vpmTimer.PulseSw 62
  PlaySoundAtBallVol SoundFX("fx_target",DOFContactors), Vol(ActiveBall)*VolumeDial
  sw62ap.rotz=TargetRotz
  vpmTimer.AddTimer TargetDelay,"sw62ap.rotz=Targetrotz/2'"
  vpmTimer.AddTimer TargetDelay+TargetDelay,"sw62ap.rotz=0'"
End Sub

Sub sw62b_Hit
  TargetBouncer Activeball, 1
  vpmTimer.PulseSw 62
  PlaySoundAtBallVol SoundFX("fx_target",DOFContactors), Vol(ActiveBall)*VolumeDial
  sw62bp.rotz=TargetRotz
  vpmTimer.AddTimer TargetDelay,"sw62bp.rotz=Targetrotz/2'"
  vpmTimer.AddTimer TargetDelay+TargetDelay,"sw62bp.rotz=0'"
End Sub

Sub sw62c_Hit
  TargetBouncer Activeball, 1
  vpmTimer.PulseSw 62
  PlaySoundAtBallVol SoundFX("fx_target",DOFContactors), Vol(ActiveBall)*VolumeDial
  sw62cp.rotz=TargetRotz
  vpmTimer.AddTimer TargetDelay,"sw62cp.rotz=Targetrotz/2'"
  vpmTimer.AddTimer TargetDelay+TargetDelay,"sw62cp.rotz=0'"
End Sub

'******************************************************
'           GENIE
'******************************************************

Const GenieMass = 0.8
Dim GenieVel: GenieVel = 0
Dim GenieAngle: GenieAngle = 0
Dim GenieDecay: GenieDecay = 0.8

Dim BounceAcc: BounceAcc = 0.5
Dim BounceSpeed: BounceSpeed = 0.75


Sub UpdateGenie()

  GenieAngle = GenieAngle + bheight(GenieVel/2.2, BounceAcc/2.2, BounceSpeed)
  GenieVel = (GenieVel - BounceAcc*(BounceSpeed)^2/2)

  If GenieAngle > 10 Then
    GenieAngle = 10
    GenieVel = 0
  End If

  If GenieAngle < 0 Then
    GenieAngle = 0
    GenieVel = -GenieVel * GenieDecay
  End If

  If GenieAngle > 4.5 Then
    Controller.Switch(42) = 1
  Else
    Controller.Switch(42) = 0
  End If

  GenieP.ObjRotX = -GenieAngle
  GenieP1.ObjRotX = -GenieAngle
  Genieflash.ObjRotX = -GenieAngle
End Sub

Function bvel(acc, height)
  bvel = sqr(2*acc*height)*0.75
End Function

Function bheight(vel, acc, time)
  bheight = vel*time  - (acc * time^2)/2
End Function

Function ECVel(Velocity1, Mass1, Velocity2, Mass2)
  ECVel = (Mass1 - Mass2)/(Mass1 + Mass2) * Velocity1  + 2 * Mass2/(Mass1 + Mass2)*Velocity2
End Function

Sub GenieTrig_hit()
  GenieCol(Activeball)
End Sub

Sub GenieTrig1_hit()
  GenieCol(Activeball)
End Sub

Sub GenieCol(ball)
  If ball.vely < 322.1 - GenieAngle/2.2 Then
    If GenieAngle < 4 Then
      PlaySoundAtBallVol "fx_geniehit",  Csng(ball.vely^2 / 2000)
    End If

    Ball.vely = ECVel(Ball.vely, BallMass, -GenieVel, GenieMass)
    GenieVel = ECVel(GenieVel, GenieMass, -ball.vely, BallMass)
  End If
End Sub

'******************************************************
'             LAMP
'******************************************************

Dim discPosition, discSpinSpeed, discLastPos, SpinCounter, maxvel
dim spinAngle, degAngle, spinAngle2, degAngle2, startAngle, postSpeedFactor
dim discX, discY
startAngle = -15
discX = 510
discY = 840
PostSpeedFactor = 130'90

Const cDiscSpeedMult = 32 '35           ' Affects speed transfer to object (deg/sec)
Const cDiscFriction = 0.55  '1.0          ' Friction coefficient (deg/sec/sec)
Const cDiscMinSpeed = 0.05            ' Object stops at this speed (deg/sec)
'Const cDiscMinSpeed = 5            ' use this value if you want to enable DOF for lamp below in script
Const cDiscRadius = 60

'Wobble
Const discSpringConst = -70 '-100
Const discSpringAngle = 30
Const discSpringRange = 25
'End Wobble

Sub SpinnerBallTimer_Timer()
  Dim oldDiscSpeed, discFriction
  oldDiscSpeed = discSpinSpeed

  discPosition = discPosition + discSpinSpeed * Me.Interval / 1000

  if ABS(discSpinSpeed) < 200 Then
    discFriction = 6 ' was 6
  else
    discFriction = cDiscFriction
  end if
  discSpinSpeed = discSpinSpeed * (1 - discFriction * Me.Interval / 1000)

  Do While discPosition < 0 : discPosition = discPosition + 360 : Loop
  Do While discPosition > 360 : discPosition = discPosition - 360 : Loop

  'Wobble

  Dim UpperRange, LowerRange
  UpperRange = discSpringAngle + discSpringRange
  LowerRange = discSpringAngle - discSpringRange

  If abs(discSpinSpeed) < 400 Then
    If discPosition > LowerRange and discPosition < discSpringAngle Then
      discSpinSpeed = newDiscSpinSpeed(discSpinSpeed ,discPosition - LowerRange, Me.Interval / 1000)
    ElseIf discPosition > discSpringAngle and discPosition < UpperRange  Then
      discSpinSpeed = newDiscSpinSpeed(discSpinSpeed ,discPosition - UpperRange, Me.Interval / 1000)
    ElseIf discPosition > LowerRange+180 and discPosition < discSpringAngle+180 Then
      discSpinSpeed = newDiscSpinSpeed(discSpinSpeed ,discPosition - LowerRange - 180, Me.Interval / 1000)
    ElseIf discPosition > discSpringAngle+180 and discPosition < UpperRange+180  Then
      discSpinSpeed = newDiscSpinSpeed(discSpinSpeed ,discPosition - UpperRange - 180, Me.Interval / 1000)
    End If
  End If
  'End Wobble

  If Abs(discSpinSpeed) < cDiscMinSpeed Then
    discSpinSpeed = 0
    'DOF 103,DOFOff
  Else
    'DOF 103,DOFOn
  End If

  If discSpinSpeed < 0 and discPosition < 210 and discPosition > 30 and discLastPos <> 180 Then
    PlaySoundAt SoundFX("fx_lamp",DOFGear), LampPr1
    vpmTimer.PulseSw 56
    discLastPos = 180
  ElseIf discSpinSpeed < 0 and (discPosition >= 210  or discPosition < 30) and discLastPos <> 360 Then
    PlaySoundAt SoundFX("fx_lamp",DOFGear), LampPr1
    vpmTimer.PulseSw 56
    discLastPos = 360
  ElseIf discSpinSpeed > 0 and discPosition < 210  and discPosition > 30 and discLastPos <> 180 Then
    PlaySoundAt SoundFX("fx_lamp",DOFGear), LampPr1
    vpmTimer.PulseSw 57
    discLastPos = 180
  ElseIf discSpinSpeed > 0 and (discPosition >= 210  or discPosition < 30) and discLastPos <> 360 Then
    PlaySoundAt SoundFX("fx_lamp",DOFGear), LampPr1
    vpmTimer.PulseSw 57
    discLastPos = 360
  End If

  degAngle = -180 + startAngle + discPosition
  degAngle2 = degAngle + 180

  spinAngle = PI * (degAngle) / 180
  spinAngle2 = PI * (degAngle2) / 180

  SpinnerBall1.x = discX + (cDiscRadius * Cos(spinAngle))
  SpinnerBall1.y = discY + (cDiscRadius * Sin(spinAngle))
  SpinnerBall1.z = 25

  If ABS(discSpinSpeed*sin(spinAngle)/postSpeedFactor) < 0.05 Then
    SpinnerBall1.velx = 0.05
  Else
    SpinnerBall1.velx = - discSpinSpeed*sin(spinAngle)/postSpeedFactor
  End If

  If Abs(discSpinSpeed*cos(spinAngle)/postSpeedFactor) < 0.05 Then
    SpinnerBall1.vely = 0.05
  Else
    SpinnerBall1.vely = discSpinSpeed*cos(spinAngle)/postSpeedFactor    '0.05
  End If

  SpinnerBall1.velz = 0

  SpinnerBall2.x = discX + (cDiscRadius * Cos(spinAngle2))
  SpinnerBall2.y = discY + (cDiscRadius * Sin(spinAngle2))
  SpinnerBall2.z = 25


  If ABS(discSpinSpeed*sin(spinAngle2)/postSpeedFactor) < 0.05 Then
    SpinnerBall2.velx = 0.05
  Else
    SpinnerBall2.velx = - discSpinSpeed*sin(spinAngle2)/postSpeedFactor
  End If

  If Abs(discSpinSpeed*cos(spinAngle2)/postSpeedFactor) < 0.05 Then
    SpinnerBall2.vely = 0.05
  Else
    SpinnerBall2.vely = discSpinSpeed*cos(spinAngle2)/postSpeedFactor   '0.05
  End If

  SpinnerBall2.velz = 0

  LampPr.objrotz = discPosition + 75
  LampPr1.objrotz = discPosition + 75
  LampPr3.objrotz = discPosition + 75
  LampPr001.objrotz = discPosition + 75
  Flasher2.RotZ = discPosition + 75
  Flasher3.RotZ = discPosition + 75

End Sub

Function newDiscSpinSpeed(spinspeed, springangle, springtime)
  newDiscSpinSpeed = spinspeed + discSpringConst * springangle * springtime
End Function

'********************************************
' Ball Collision, spinner collision and Sound
'********************************************

Sub OnBallBallCollision(ball1, ball2, velocity)
  dim collAngle,bvelx,bvely,hitball, whichBall
  If ball1.radius < 23 or ball2.radius < 23 then

    If ball1.radius < 23 Then
      collAngle = GetCollisionAngle(ball1.x,ball1.y,ball2.x,ball2.y)
      set hitball = ball2
      If ball1.x = SpinnerBall1.x and ball1.y = SpinnerBall1.y Then
        whichball = 1
      Else
        whichball = 2
      End If
    else
      collAngle = GetCollisionAngle(ball2.x,ball2.y,ball1.x,ball1.y)
      set hitball = ball1
      If ball2.x = SpinnerBall1.x and ball2.y = SpinnerBall1.y Then
        whichball = 1
      Else
        whichball = 2
      End If
    End If

    dim discAngle

    If whichBall = 1 Then
      discAngle = NormAngle(spinAngle)
    Else
      discAngle = NormAngle(spinAngle2)
    End If

'   discSpinSpeed = discspinspeed + ecvel(0,1.5,sin(collAngle - discAngle)*velocity,BallMass * ABS(sin(collAngle - discAngle))) * cDiscSpeedMult

'   PlaySound "fx_lamphit", 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 1, 0, AudioFade(ball1)

    dim sineOfAngle, sineOfAngleSqr
    sineOfAngle = sin(collAngle - discAngle)

    discSpinSpeed = discspinspeed + ecvel(0,1.5,sineOfAngle*velocity,BallMass) * cDiscSpeedMult

    PlaySound "fx_lamphit", 0, Csng(velocity) ^2 / 2000 / 3, AudioPan(ball1), 0, Pitch(ball1), 1, 0, AudioFade(ball1)


  Else
    If ball1.z > 10 and ball2.z > 10 Then
      PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
    End If
  End If
End Sub

Function GetCollisionAngle(ax, ay, bx, by)
  Dim ang
  Dim collisionV:Set collisionV = new jVector
  collisionV.SetXY ax - bx, ay - by
  GetCollisionAngle = collisionV.ang
End Function

Function NormAngle(angle)
  NormAngle = angle
''  Dim pi:pi = 3.14159265358979323846
  Do While NormAngle>2 * pi
    NormAngle = NormAngle - 2 * pi
  Loop
  Do While NormAngle <0
    NormAngle = NormAngle + 2 * pi
  Loop
End Function

Class jVector
     Private m_mag, m_ang, pi

     Sub Class_Initialize
         m_mag = CDbl(0)
         m_ang = CDbl(0)
         pi = CDbl(3.14159265358979323846)
     End Sub

     Public Function add(anothervector)
         Dim tx, ty, theta
         If TypeName(anothervector) = "jVector" then
             Set add = new jVector
             add.SetXY x + anothervector.x, y + anothervector.y
         End If
     End Function

     Public Function multiply(scalar)
         Set multiply = new jVector
         multiply.SetXY x * scalar, y * scalar
     End Function

     Sub ShiftAxes(theta)
         ang = ang - theta
     end Sub

     Sub SetXY(tx, ty)

         if tx = 0 And ty = 0 Then
             ang = 0
          elseif tx = 0 And ty <0 then
             ang = - pi / 180 ' -90 degrees
          elseif tx = 0 And ty>0 then
             ang = pi / 180   ' 90 degrees
         else
             ang = atn(ty / tx)
             if tx <0 then ang = ang + pi ' Add 180 deg if in quadrant 2 or 3
         End if

         mag = sqr(tx ^2 + ty ^2)
     End Sub

     Property Let mag(nmag)
         m_mag = nmag
     End Property

     Property Get mag
         mag = m_mag
     End Property

     Property Let ang(nang)
         m_ang = nang
         Do While m_ang>2 * pi
             m_ang = m_ang - 2 * pi
         Loop
         Do While m_ang <0
             m_ang = m_ang + 2 * pi
         Loop
     End Property

     Property Get ang
         Do While m_ang>2 * pi
             m_ang = m_ang - 2 * pi
         Loop
         Do While m_ang <0
             m_ang = m_ang + 2 * pi
         Loop
         ang = m_ang
     End Property

     Property Get x
         x = m_mag * cos(ang)
     End Property

     Property Get y
         y = m_mag * sin(ang)
     End Property

     Property Get dump
         dump = "vector "
         Select Case CInt(ang + pi / 8)
             case 0, 8:dump = dump & "->"
             case 1:dump = dump & "/'"
             case 2:dump = dump & "/\"
             case 3:dump = dump & "'\"
             case 4:dump = dump & "<-"
             case 5:dump = dump & ":/"
             case 6:dump = dump & "\/"
             case 7:dump = dump & "\:"
         End Select

         dump = dump & " mag:" & CLng(mag * 10) / 10 & ", ang:" & CLng(ang * 180 / pi) & ", x:" & CLng(x * 10) / 10 & ", y:" & CLng(y * 10) / 10
     End Property
End Class

'******************************************************
'           FUNCTIONS
'******************************************************

'*** PI returns the value for PI
'Function PI()
' PI = 4*Atn(1)
'End Function

'*** Determines if a Points (px,py) is inside a 4 point polygon A-D in Clockwise/CCW order
'Function InRect(px,py,ax,ay,bx,by,cx,cy,dx,dy)
' Dim AB, BC, CD, DA
' AB = (bx*py) - (by*px) - (ax*py) + (ay*px) + (ax*by) - (ay*bx)
' BC = (cx*py) - (cy*px) - (bx*py) + (by*px) + (bx*cy) - (by*cx)
' CD = (dx*py) - (dy*px) - (cx*py) + (cy*px) + (cx*dy) - (cy*dx)
' DA = (ax*py) - (ay*px) - (dx*py) + (dy*px) + (dx*ay) - (dy*ax)
'
' If (AB <= 0 AND BC <=0 AND CD <= 0 AND DA <= 0) Or (AB >= 0 AND BC >=0 AND CD >= 0 AND DA >= 0) Then
'   InRect = True
' Else
'   InRect = False
' End If
'End Function

'Function Distance(ax,ay,bx,by)
' Distance = SQR((ax - bx)^2 + (ay - by)^2)
'End Function

'Function DistancePL(px,py,ax,ay,bx,by) ' Distance between a point and a line where point is px,py
' DistancePL = ABS((by - ay)*px - (bx - ax) * py + bx*ay - by*ax)/Distance(ax,ay,bx,by)
'End Function

'Function AnglePP(ax,ay,bx,by)
' AnglePP = Atn2((by - ay),(bx - ax))*180/PI
'End Function

'Function Atn2(dy, dx)
' If dx > 0 Then
'   Atn2 = Atn(dy / dx)
' ElseIf dx < 0 Then
'   Atn2 = Sgn(dy) * (PI - Atn(Abs(dy / dx)))
' ElseIf dy = 0 Then
'   Atn2 = 0
' Else
'   Atn2 = Sgn(dy) * Pi / 2
' End If
'End Function


'******************************************************
'       GI / LAMPS / FLASHERS
'******************************************************

' *****************************************
' *** GI                              *****
' *****************************************
Dim globalGI
Sub UpdateGI(no, step)
  If no <> 2 Then  Exit Sub

  globalGI = CInt(4 + (1 + step * 6) * 6)
  If globalGI < 4 + 2 * 6 Then DOF 200, DOFOff
  If globalGI > 4 + 7 * 6 then DOF 200, DOFOn

  For Each obj In FadingGIlights : obj.state = step : next

  If globalGI < 30 and leftsling.image = "slingtexgion" Then
    leftsling.image = "slingtexgioff" : rightsling.image = "slingtexgioff"
    leftcage.image = "tex5gioff" : rightcage.image = "tex5gioff"
  End If
  If globalGI > 30 and leftsling.image = "slingtexgioff" Then
    leftsling.image = "slingtexgion" : rightsling.image = "slingtexgion"
    leftcage.image = "tex5" : rightcage.image = "tex5"
  End If

  Dim BallColor,BallFade
  BallFade = 111 + (globalGI - 10) * 4
  BallColor = RGB(BallFade,BallFade,BallFade)

  TOTBall1.color = BallColor
  TOTBall2.color = BallColor
  TOTBall3.color = BallColor
  TOTBall4.color = BallColor

  CapR1.color = BallColor
  CapR2.color = BallColor
  CapL1.color = BallColor
  CapL2.color = BallColor

  Dim pfrscale: pfrscale = (globalGI - 10) / 4

  sw16l1.intensityscale = (globalGI-10)/36
  sw27l1.intensityscale = (globalGI-10)/36
  Light77.intensityscale = (globalGI-10)/36

  dim calc : calc = 56 - globalGI

  If Show3dScrews = 1 Then
    texscrewsgion.material = "lit" & globalGI
  End If

  tex1gioff.material = "lithalf" & calc

  tex3gion.material = "lit" & globalGI
  tex3gion1.material = "lit" & globalGI
  tex2gion.material = "lit" & globalGI
  tex2bgion.material = "lit" & globalGI
  rampgion.material = "lithalf" & globalGI
  sidewallsgion.material = "lit" & globalGI
  tex5gion.material = "lit" & globalGI
  GenieP1.material = "lit" & globalGI
  kickerleftgion.material = "lithalf" & globalGI
  kickerrightgion.material = "lithalf" & globalGI
  LampPr3.material = "glossy" & globalGI

  apronmesh.material = "color" & globalGI
  tex5gioff1.blenddisablelighting =  ((globalGI-10)/36)
  tex5gioff1.material = "color" & globalGI
  'Gate2b.material = "glossy10" '& globalGI

  RampDivP1.material = "lit" & globalGI
  PlayFDivP1.material = "lit" & globalGI
  sw63p1.material = "lit" & globalGI
  sw64p1.material = "lit" & globalGI
  sw65p1.material = "lit" & globalGI
  sw58p1.material = "lit" & globalGI
  sw46p1.material = "lit" & globalGI
  sw23p1.material = "lit" & globalGI
  sw46bp1.material = "lit" & globalGI
  sw48p1.material = "lit" & globalGI
  Primitive46.blenddisablelighting = (globalGI-10)/36

  Bulb1.blenddisablelighting = 1 + 400 * ((globalGI-10)/36)^2
  Bulb2.blenddisablelighting = 1 + 400 * ((globalGI-10)/36)^2

  tex3gion.blenddisablelighting = 1 + (globalGI-10)/36
  tex2gion.blenddisablelighting = 1 + (globalGI-10)/36
  tex2bgion.blenddisablelighting = 1 + (globalGI-10)/36
  tex1gioff.blenddisablelighting = (calc-10)/72 '0.2 + (calc-10)/72
  rampgion.blenddisablelighting = 1 + (globalGI-10)/36
  sidewallsgion.blenddisablelighting = 1 + (globalGI-10)/36
  tex5gion.blenddisablelighting = 0.5 + (globalGI-10)/36
  If sw16p.blenddisablelighting  < 0.99 Then sw16p.blenddisablelighting = (globalGI-10)/36 : End If
  If sw27p.blenddisablelighting  < 0.99 Then sw27p.blenddisablelighting = (globalGI-10)/36 : End If

  calc = round((globalGI - 10) / 2) + 10
  LampPr1.material = "color" & calc

  If RedWhiteFlipperBat = 0 Then
    leftflip.material = "color" & globalGI
    leftflip.blenddisablelighting = (globalGI-10)/36
    rightflip.material = "color" & globalGI
    rightflip.blenddisablelighting = (globalGI-10)/36
  Else
    leftflip.material = "color" & calc
    leftflip.blenddisablelighting =  0 + 1.5 * (globalGI-10)/36
    rightflip.material = "color" & calc
    rightflip.blenddisablelighting = 0 + 1.5 * (globalGI-10)/36
  End If

  flasher2.IntensityScale =  1 * (globalGI-10)/36
  flasher3.IntensityScale =  1 * (globalGI-10)/36
  Flasher4.IntensityScale =  0.1 * (globalGI-10)/36

  calc = 0.2 + 20 * (globalGI-10)/36

  sw61ap.blenddisablelighting = calc
  sw61bp.blenddisablelighting = calc
  sw61cp.blenddisablelighting = calc
  sw62ap.blenddisablelighting = calc
  sw62bp.blenddisablelighting = calc
  sw62cp.blenddisablelighting = calc

  If globalGI = 46 and tex1gioff.visible Then
    tex1gioff.visible = false
    tex2gioff.visible = false
    tex2bgioff.visible = false
    tex5gioff.visible = false
    GenieP.visible = false
    LampPr.visible = false
    kickerleftgioff.visible = false
    kickrightgioff.visible = false
    sidewallsgioff.visible = false
    RampDivP.visible = false
    PlayFDivP.visible = false
    sw63p.visible = false
    sw64p.visible = false
    sw65p.visible = false
    sw58p.visible = false
    sw46p.visible = false
    sw23p.visible = false
    sw46bp.visible = false
    sw48p.visible = false
    texscrewsgioff.visible = false
  End If

  If globalGI > 10 and globalGI < 46 then
    tex1gioff.visible = True
    tex2gioff.visible = True
    tex2bgioff.visible = True
    tex5gioff.visible = True
    GenieP.visible = True
    LampPr.visible = True
    kickerleftgioff.visible = True
    kickrightgioff.visible = True
    sidewallsgioff.visible = True
    RampDivP.visible = True
    PlayFDivP.visible = True
    sw63p.visible = True
    sw64p.visible = True
    sw65p.visible = True
    sw58p.visible = True
    sw46p.visible = True
    sw23p.visible = True
    sw46bp.visible = True
    sw48p.visible = True
    tex2gion.visible = True
    tex2bgion.visible = True
    tex3gion.visible = True
    tex3gion1.visible = True
    tex5gion.visible = True
    rampgion.visible = True
    GenieP1.visible = True
    LampPr3.visible = True
    kickerleftgion.visible = True
    kickerrightgion.visible = True
    sidewallsgion.visible = True
    RampDivP1.visible = True
    PlayFDivP1.visible = True
    sw63p1.visible = True
    sw64p1.visible = True
    sw65p1.visible = True
    sw58p1.visible = True
    sw46p1.visible = True
    sw23p1.visible = True
    sw46bp1.visible = True
    sw48p1.visible = True
    If show3dscrews = 1 Then
      texscrewsgioff.visible = True
      texscrewsgion.visible = True
    End If
  End If

  If globalGI = 10 and tex2gion.visible Then
    tex2gion.visible = false
    tex2bgion.visible = false
    tex3gion.visible = false
    tex3gion1.visible = False
    tex5gion.visible = false
    rampgion.visible = false
    GenieP1.visible = False
    LampPr3.visible = False
    kickerleftgion.visible = false
    kickerrightgion.visible = false
    sidewallsgion.visible = False
    RampDivP1.visible = False
    PlayFDivP1.visible = False
    sw63p1.visible = false
    sw64p1.visible = false
    sw65p1.visible = false
    sw58p1.visible = false
    sw46p1.visible = false
    sw23p1.visible = false
    sw46bp1.visible = false
    sw48p1.visible = false
    texscrewsgion.visible = False
  End If
End Sub

' *****************************************
' *** insert lights                   *****
' *****************************************

Sub GraphicsTimer_Timer()
  leftflip.ObjRotZ = LeftFlipper.CurrentAngle - 90
  rightflip.ObjRotZ = RightFlipper.CurrentAngle -270
  batleftshadow.rotz = LeftFlipper.CurrentAngle
  batrightshadow.rotz  = RightFlipper.CurrentAngle

  If PlayFDiv.currentangle < PlayFDivStart + 0.5 Then
    PlayFDivp.objrotz = 0.5
    PlayFDivP1.objrotz = 0.5
  Elseif PlayFDiv.currentangle < PlayFDivStart + 12.75 Then
    PlayFDivP.objrotz = PlayFDiv.currentangle - PlayFDivStart
    PlayFDivP1.objrotz = PlayFDiv.currentangle - PlayFDivStart
  Else
    PlayFDivP.objrotz = 12.75
    PlayFDivP1.objrotz = 12.75
  End If

  if PlayFDiv.currentangle < 181 Then
    PlayFDiv.enabled = False
    diverterclosed.collidable = True
    diverteropen.collidable = False
  Elseif PlayFDiv.currentangle > 193.2 Then
    PlayFDiv.enabled=True
    diverterclosed.collidable = False
    diverteropen.collidable = True
  Else
    PlayFDiv.enabled=True
    diverterclosed.collidable = False
    diverteropen.collidable = False
  End If
End Sub

' *****************************************
' *** Flashers                        *****
' *****************************************

SolModCallback(16) = "SetModLamp 116,"
SolModCallback(17) = "SetModLamp 117,"
SolModCallback(18) = "SetModLamp 118,"
SolModCallback(19) = "SetModLamp 119,"
SolModCallback(20) = "SetModLamp 120,"
SolModCallback(22) = "SetModLamp 122,"
SolModCallback(23) = "SetModLamp 123,"
SolModCallback(24) = "SetModLamp 124,"
SolModCallback(25) = "SetModLamp 125,"
SolModCallback(26) = "SetModLamp 126,"
SolModCallback(27) = "SetModLamp 127,"
SolModCallback(28) = "SetModLamp 128,"

Dim flashFactor : flashFactor = 0.25 ' To allow easy testing and adjustments

Sub SetModLamp(nr, level)
  'Debug.Print nr & " => " & level
  dim isOn : if level > 0.5 Then isOn = True Else isOn = False End If
  level = level * flashFactor
  dim flashx3 : flashx3 = level 'level^3
  dim flashcolor : flashcolor = round (level * 36) + 10
  Select Case nr
    case 116:
      flasherflash116.intensityscale = 5
      flasherflash116.state = 5 * flashx3
      tex1flashlefteject.visible = True 'isOn
      tex1flashlefteject.blenddisablelighting = 20 * flashx3
      tex1flashlefteject.material = "lit" & flashcolor
      tex3flashlefteject.visible = True 'isOn
      tex3flashlefteject.blenddisablelighting = 1 + 40 * flashx3
      tex3flashlefteject.material = "lit" & flashcolor
      rampflashlefteject.visible = True 'isOn
      rampflashlefteject.blenddisablelighting = 40 * flashx3
      rampflashlefteject.material = "lit" & flashcolor
      If backboxflash Then Flasher002.opacity = 4000 * Flashx3 : End If
    case 117:
      tex1flashinlane.visible = True 'isOn
      ' tex1flashinlane.blenddisablelighting = flashx3
      tex1flashinlane.material = "lit" & flashcolor
      tex2flashinlane.visible = True 'isOn
      tex2flashinlane.blenddisablelighting = 1 + flashx3
      tex2flashinlane.material = "lit" & flashcolor
      tex3flashinlane.visible = True 'isOn
      tex3flashinlane.blenddisablelighting = 1 + 3 * flashx3
      tex3flashinlane.material = "lit" & flashcolor
      tex3flashinlane1.visible = True 'isOn
      tex3flashinlane1.blenddisablelighting = 1 + 3 * flashx3
      tex3flashinlane1.material = "lit" & flashcolor
      sidewallsflash.visible = True 'isOn
      sidewallsflash.blenddisablelighting = 1 + flashx3
      sidewallsflash.material = "lit" & flashcolor
      sw16p.blenddisablelighting = 20 * level
      sw27p.blenddisablelighting = 20 * level
      rampflash.visible = True 'isOn
      rampflash.blenddisablelighting = flashx3
      rampflash.material = "lit" & flashcolor
      Light87.intensityscale = 1
      Light87.state = flashx3
      Light88.intensityscale = 1
      Light88.state = flashx3
      Light003.intensityscale = 0.5
      Light003.state = 0.5 * level
      FlasherFlash117.intensityscale = 0.5
      FlasherFlash117.state = 0.5 * level
      sidewallsgioff.blenddisablelighting = 1 + 3 * flashx3
      sidewallsgion.blenddisablelighting = 1 + 3 * flashx3
    case 118:
      flasherflash118.intensityscale = 20 * flashx3
      Light106.intensityscale = 5 * flashx3
    case 119:
      flasherflash119.intensityscale = 5
      flasherflash119.state = flashx3
      Light91.intensityscale = 5
      Light91.state = flashx3
      Light90.intensityscale = 5
      Light90.state = flashx3
      Light89.intensityscale = 1
      Light89.state = level
    case 120:
      tex1flashbazaar.visible = True 'isOn
      tex1flashbazaar.blenddisablelighting = 40 * flashx3
      tex1flashbazaar.material = "lit" & flashcolor
      tex2flashbazaar.visible = True 'isOn
      tex2flashbazaar.blenddisablelighting = 10 * flashx3
      tex2flashbazaar.material = "lit" & flashcolor
      rampflashbazaar.visible = True 'isOn
      rampflashbazaar.blenddisablelighting = 10 * flashx3
      rampflashbazaar.material = "lit" & flashcolor
      tex3flashbazaar.visible = True 'isOn
      tex3flashbazaar.blenddisablelighting = 10 * flashx3
      tex3flashbazaar.material = "lit" & flashcolor
      flasherflash120.state = level
      flasherflash120.intensityscale = 5
    case 122:
      flasherflash122.intensityscale = 20
      flasherflash122.state = flashx3
      FlasherFlash122b.intensityscale = 10
      FlasherFlash122b.state = level
      FlasherFlash122c.intensityscale = 1
      FlasherFlash122c.state = flashx3
      Light92.intensityscale = 5
      Light92.state = flashx3
      Light93.intensityscale = 5
      Light93.state = flashx3
    Case 123: ' Flasher under lamp
      Flasherflash123.intensityscale = 10
      Flasherflash123.state = flashx3
      LampPr001.visible = True 'isOn
      LampPr001.blenddisablelighting = 10 * level 'flashx3
      Light86.state = level
      Light84.state = 1.0 - level
      Light85.state = 1.0 - level
      Flasher5.IntensityScale =  0.5 * level
      if 0.1 * level > 0.1 * (globalGI-10)/36 Then
        Flasher4.IntensityScale =  0.1 * level
      Else
        Flasher4.IntensityScale =  0.1 * (globalGI-10)/36
      End If
    Case 124:
      flasherflash124.intensityscale = 5
      flasherflash124.state = flashx3
      Light95.intensityscale = 5
      Light95.state = flashx3
      Light96.intensityscale = 5
      Light96.state = flashx3
      Light97.intensityscale = 1
      Light97.state = level
    Case 125:
      sidewallsflashstarttale.visible = True 'isOn
      sidewallsflashstarttale.blenddisablelighting = 8 * flashx3
      sidewallsflashstarttale.material = "lit" & flashcolor
      tex1flashstarttale.visible = True 'isOn
      tex1flashstarttale.blenddisablelighting = 1 * flashx3
      tex1flashstarttale.material = "lit" & flashcolor
      rampflashstarttale.visible = True 'isOn
      rampflashstarttale.blenddisablelighting = 1 * flashx3
      rampflashstarttale.material = "lit" & flashcolor
      tex3flashstarttale.visible = True 'isOn
      tex3flashstarttale.blenddisablelighting = 2 * flashx3
      tex3flashstarttale.material = "lit" & flashcolor
      flasherflash125.intensityscale = 5
      flasherflash125.state = level
      Light105.intensityscale = 1
      Light105.state = level
      flasherflash125b.intensityscale = 3
      flasherflash125b.state = flashx3
      Light107.intensityscale = 30
      Light107.state = flashx3
      If backboxflash Then Flasher001.opacity = 4000 * Flashx3 : End If
    Case 126:
      tex1flashjet.visible = True 'isOn
      tex1flashjet.blenddisablelighting = 10 * flashx3
      tex1flashjet.material = "lit" & flashcolor
      sidewallsflashjet.visible = True 'isOn
      sidewallsflashjet.blenddisablelighting = 10 * flashx3
      sidewallsflashjet.material = "lit" & flashcolor
      rampflashjet.visible = True 'isOn
      rampflashjet.blenddisablelighting = 40 * flashx3
      rampflashjet.material = "lit" & flashcolor
      tex3flashjet.visible = True 'isOn
      tex3flashjet.blenddisablelighting = 100 * flashx3
      tex3flashjet.material = "lit" & flashcolor
      flasherflash126.state = level
      flasherflash126.intensityscale = 5
      Light98.intensityscale = 5
      Light98.State = flashx3
      Light102.intensityscale = 5
      Light102.State = flashx3
      Light101.intensityscale = 1
      Light101.State = flashx3
      Light103.intensityscale = 5
      Light103.State = flashx3
      Light104.intensityscale = 5
      Light104.State = flashx3
      Light7.intensityscale = 5
      Light7.State = flashx3
      If backboxflash Then Flasher004.opacity = 4000 * Flashx3 : End If
    Case 127:
      genieflash.visible = True 'isOn
      genieflash.blenddisablelighting = 1 + 20 * level^0.25
      genieflash.material = "lit" & flashcolor
      flasherflash127.intensityscale = 5 * flashx3 : Light94.intensityscale = 5 * flashx3 : Light100.intensityscale = 5 * flashx3: Light99.intensityscale = level
      If backboxflash Then Flasher003.opacity = 4000 * Flashx3 : End If
    case 128:
      dim flashcolor2 : flashcolor2 = round (level * 36 / 2) + 10
      flasherflash128.intensityscale = 20 * flashx3
      tex1flashramp.visible = True 'isOn
      tex1flashramp.blenddisablelighting = 10 * flashx3
      tex1flashramp.material = "lit" & flashcolor2
      rampflashramp.visible = True 'isOn
      rampflashramp.blenddisablelighting = 5 * flashx3
      rampflashramp.material = "lit" & flashcolor2
  End select
End Sub

'******************************************************
'           SOUNDS
'******************************************************

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

Sub PlaySoundAtLoop(soundname, tableobj)
    PlaySound soundname, -1, 1, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtVol(soundname, aVol, tableobj)
    PlaySound soundname, 1, aVol, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub


Sub PlaySoundAtExisting(soundname, tableobj)
    PlaySound soundname, 1, 1, AudioPan(tableobj), 0,0,1, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname)
    PlaySoundAt soundname, ActiveBall
End Sub

Sub PlaySoundAtBallVol (Soundname, aVol)
  Playsound soundname, 1,aVol, AudioPan(ActiveBall), 0,0,0, 1, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtBallVolM (Soundname, aVol)
  Playsound soundname, 1,aVol, AudioPan(ActiveBall), 0,0,0, 0, AudioFade(ActiveBall)
End Sub


'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
    tmp = tableobj.y * 2 / tableheight-1
    If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = tableobj.x * 2 / tablewidth-1
    If tmp > 0 Then
        AudioPan = Csng(tmp ^10)
    Else
        AudioPan = Csng(-((- tmp) ^10) )
    End If
End Function

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallSpeed(ball) ^2 / 2000)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallSpeed(ball) * 20
End Function
'
'Function BallSpeed(ball) 'Calculates the ball speed
'    BallSpeed = SQR(ball.VelX^2 + ball.VelY^2 + ball.VelZ^2)
'End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / tablewidth-1
    If tmp> 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

'*********************************************************************
'                       Collection Sounds
'*********************************************************************

Sub aApron_Hit(idx):PlaySoundAtBallVolM "fx_apron", Vol(ActiveBall)*VolumeDial:End Sub
Sub aRubbers_Hit(idx):PlaySoundAtBallVol "fx_rubber", Vol(ActiveBall)*VolumeDial:End Sub
Sub aPostRubbers_Hit(idx):PlaySoundAtBallVol "fx_postrubber", Vol(ActiveBall)*VolumeDial:End Sub
Sub aMetals_Hit(idx):PlaySoundAtBallVolM "fx_MetalHit", Vol(ActiveBall)*VolumeDial/100:End Sub
Sub aPlastics_Hit(idx):PlaySoundAtBallVol "fx_PlasticHit", Vol(ActiveBall)*VolumeDial*5:End Sub
Sub aGates_Hit(idx):PlaySoundAtBallVol "fx_Gate", Vol(ActiveBall)*VolumeDial:End Sub
Sub aWoods_Hit(idx):PlaySoundAtBallVolM "fx_Woodhit", Vol(ActiveBall)*VolumeDial/5:End Sub
Sub aSkillShot_Hit(idx):PlaySoundAtBallVol "fx_MetalHit", Vol(ActiveBall)*VolumeDial/100:End Sub

'*********************************************************************
'                       Ramp Sounds
'*********************************************************************

Sub REnd1_Hit()
    PlaySoundAtBall "fx_ExitRampToPlayfield"
End Sub


Sub REnd2_Hit()
    StopSound "fx_metalrolling"
    PlaySoundAtBall "fx_ExitRampToPlayfield"
End Sub


' left ramp sounds
Sub LRHit0_Hit:PlaySoundAtBall "fx_lr1":End Sub
Sub LRHit1_Hit:PlaySoundAtBall "fx_lr2":End Sub
Sub LRHit2_Hit:PlaySoundAtBall "fx_lr3":End Sub
Sub LRHit3_Hit:PlaySoundAtBall "fx_lr4":End Sub
Sub LRHit4_Hit:PlaySoundAtBall "fx_lr5":End Sub
Sub LRHit5_Hit:PlaySoundAtBall "fx_lr6":End Sub
Sub LRHit6_Hit:PlaySoundAtBall "fx_lr7":End Sub

'right ramp sounds
Sub RRHit0_Hit:PlaySoundAtBall "fx_rr1":End Sub
Sub RRHit1_Hit:PlaySoundAtBall "fx_rr2":End Sub
Sub RRHit2_Hit:PlaySoundAtBall "fx_rr3":End Sub
Sub RRHit3_Hit:PlaySoundAtBall "fx_rr4":End Sub
Sub RRHit4_Hit:PlaySoundAtBall "fx_rr5":End Sub
Sub RRHit5_Hit:PlaySoundAtBall "fx_rr6":End Sub
Sub RRHit6_Hit:PlaySoundAtBall "fx_lr1":End Sub
Sub RRHit7_Hit:PlaySoundAtBall "fx_lr1":End Sub
Sub RRHit8_Hit:PlaySoundAtBall "fx_rr7":End Sub

Sub CenterRampHelper_unHit()
  Activeball.vely = Activeball.velY + 10
End Sub


'*********************************************************************
'           Game Timer, Ball Rolling, Ball Shadows, Ball Drop
'*********************************************************************

Const tnob = 10 ' total number of balls
ReDim rolling(tnob)
InitRolling

Dim BallShadow, BallRefl
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5,BallShadow6,BallShadow7,BallShadow8,BallShadow9,BallShadow10,BallShadow11)
BallRefl = Array (BallRefl1,BallRefl2,BallRefl3,BallRefl4,BallRefl5,BallRefl6,BallRefl7,BallRefl8,BallRefl9,BallRefl10,BallRefl11)


Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Const PlayFDivStart = 180.7

Sub GameTimer_timer()

  Dim BOT, b
  BOT = GetBalls

  ' play the rolling sound for each ball
  For b = 0 to UBound(BOT)
    If BallSpeed(BOT(b) ) > 1 AND BOT(b).z < 27 and BOT(b).radius > 23  Then
      rolling(b) = True
      PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b))*RollingSoundFactor, AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
    Else
      If rolling(b) = True Then
        StopSound("fx_ballrolling" & b)
        rolling(b) = False
      End If
    End If

    '***Ball Reflections***
    dim shift
    If BallReflections = 1 Then
      shift = (BOT(b).X - tablewidth/2)/20 * 1 / Bot(b).Y
      BallRefl(b).z = BOT(b).z - 24
      BallRefl(b).Size_z = 40 - BOT(b).y/130
      BallRefl(b).X = BOT(b).X - shift
      BallRefl(b).Y = BOT(b).Y + 5 - abs(shift)
      BallRefl(b).RotY = 180 + (BOT(b).X - tablewidth/2)/15
    End If

    '***Ball Shadows***
    If BallShadowOn = 1 Then
      BallShadow(b).X = BOT(b).X
      ballShadow(b).Y = BOT(b).Y + 10
    End If

    If BOT(b).Z > 24 and BOT(b).Z < 35 and BOT(b).radius > 23  and not inrect(BOT(b).x,BOT(b).y,183,925,227,925,227,969,183,969) Then
      BallShadow(b).visible = 1
      BallRefl(b).visible = 1
    Else
      BallShadow(b).visible = 0
      BallRefl(b).visible = 0
    End If

    '***Ball Drop Sounds***
    If BOT(b).radius > 23 and BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
      PlaySound "fx_ball_drop" & b, 0, (ABS(BOT(b).velz)/17)^2, AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
    End If

    '*** Skillshot Animation and Sounds***
    if BOT(b).z > 100 and BOT(b).x > 817 and BOT(b).y > 614 and BOT(b).y < 1056 then
      If BOT(b).velz < -1 and BOT(b).z < 215 and BOT(b).z > 177 Then
        PlaySound "fx_SkillShot-HitMetal", 0, 0.75*VolumeDial, AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
      End If

      If BOT(b).z < 176 and BOT(b).z > 165 and BOT(b).x > 856 and BOT(b).x < 906 then
        If BOT(b).y > 702 and BOT(b).y < 755 Then
          sw63_Hit BOT(b)
        Elseif BOT(b).y > 843 and BOT(b).y < 907 Then
          sw64_Hit BOT(b)
        Elseif BOT(b).y > 991 and BOT(b).y < 1054 Then
          sw65_Hit BOT(b)
        Else
          sw63_unHit:sw64_unHit:sw65_unHit
        End If
      ElseIf BOT(b).z < 165 and BOT(b).z > 130 Then
        sw63_unHit:sw64_unHit:sw65_unHit
      End If
    end if

  Next

  UpdateGenie
  cor.update

End Sub

'****************************************************************
' FLIPPER CORRECTION INITIALIZATION
'****************************************************************

Const LiveCatch = 16

dim LF : Set LF = New FlipperPolarity
dim RF : Set RF = New FlipperPolarity

InitPolarity

Sub InitPolarity()
  dim x, a : a = Array(LF, RF)
  for each x in a
    x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
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

  addpt "Velocity", 0, 0,   1
  addpt "Velocity", 1, 0.16, 1.06
  addpt "Velocity", 2, 0.41,  1.05
  addpt "Velocity", 3, 0.53,  1'0.982
  addpt "Velocity", 4, 0.702, 0.968
  addpt "Velocity", 5, 0.95,  0.968
  addpt "Velocity", 6, 1.03,  0.945

  LF.Object = LeftFlipper
  LF.EndPoint = EndPointLp
  RF.Object = RightFlipper
  RF.EndPoint = EndPointRp
End Sub

Sub TriggerLF_Hit() : LF.Addball activeball : End Sub
Sub TriggerLF_UnHit() : LF.PolarityCorrect activeball : End Sub
Sub TriggerRF_Hit() : RF.Addball activeball : End Sub
Sub TriggerRF_UnHit() : RF.PolarityCorrect activeball : End Sub

'****************************************************************
' FLIPPER CORRECTION FUNCTIONS
'****************************************************************

Sub AddPt(aStr, idx, aX, aY)  'debugger wrapper for adjusting flipper script in-game
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

'****************************************************************
' FLIPPER POLARITY AND RUBBER DAMPENER SUPPORTING FUNCTIONS
'****************************************************************

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

Function max(a,b)
  if a > b then
    max = a
  Else
    max = b
  end if
end Function

'****************************************************************
' Check ball distance from Flipper for Rem
'****************************************************************

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

'****************************************************************
' End - Check ball distance from Flipper for Rem
'****************************************************************

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

'Const LiveCatch = 16
Const LiveElasticity = 0.45
Const SOSEM = 0.815
Const EOSReturn = 0.025  'mid 90's and later

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

'   debug.print LiveCatchBounce
    If LiveCatchBounce = 0 and ball.velx * Dir > 0 Then ball.velx = 0
    ball.vely = LiveCatchBounce * (32 / LiveCatch) ' Multiplier for inaccuracy bounce
    ball.angmomx= 0
    ball.angmomy= 0
    ball.angmomz= 0
    Else
        If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 1 Then FlippersD.Dampenf Activeball, parm
  End If
End Sub

'****************************************************************
' PHYSICS DAMPENERS
'****************************************************************
'These are data mined bounce curves,
'dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
'Requires tracking ballspeed to calculate COR


Sub dPosts_Hit(idx)
  RubbersD.dampen Activeball
  TargetBouncer Activeball, 1
End Sub

'sling bottom corners removed from targetbouncer
Sub zCol_Rubber_Corner_2_Hit(idx)
  RubbersD.dampen Activeball
End Sub

Sub zCol_Rubber_Corner_3_Hit(idx)
  RubbersD.dampen Activeball
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

'****************************************************************
' TRACK ALL BALL VELOCITIES FOR RUBBER DAMPENER AND DROP TARGETS
'****************************************************************

dim cor : set cor = New CoRTracker

Class CoRTracker
  public ballvel, ballvelx, ballvely

  Private Sub Class_Initialize : redim ballvel(0) : redim ballvelx(0): redim ballvely(0) : End Sub

  Public Sub Update() 'tracks in-ball-velocity
    dim str, b, AllBalls, highestID : allBalls = GetBalls

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

Function InRotRect(ballx,bally,px,py,angle,ax,ay,bx,by,cx,cy,dx,dy)
    Dim rax,ray,rbx,rby,rcx,rcy,rdx,rdy
    Dim rotxy
    rotxy = RotPoint(ax,ay,angle)
    rax = rotxy(0)+px : ray = rotxy(1)+py
    rotxy = RotPoint(bx,by,angle)
    rbx = rotxy(0)+px : rby = rotxy(1)+py
    rotxy = RotPoint(cx,cy,angle)
    rcx = rotxy(0)+px : rcy = rotxy(1)+py
    rotxy = RotPoint(dx,dy,angle)
    rdx = rotxy(0)+px : rdy = rotxy(1)+py

    InRotRect = InRect(ballx,bally,rax,ray,rbx,rby,rcx,rcy,rdx,rdy)
End Function

'*********************************************************************
' VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'*********************************************************************

Const TargetBouncerEnabled = 1    '0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.8   'Level of bounces. Recommmended value of 0.7

'sub TargetBouncer(aBall,defvalue)
'    dim zMultiplier, vel, vratio
'    if TargetBouncerEnabled = 1 and aball.z < 30 then
'        debug.print "velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz
'        vel = BallSpeed(aBall)
'        if aBall.velx = 0 then vratio = 1 else vratio = aBall.vely/aBall.velx
'        Select Case Int(Rnd * 6) + 1
'            Case 1: zMultiplier = 0.2*defvalue
'     Case 2: zMultiplier = 0.25*defvalue
'            Case 3: zMultiplier = 0.3*defvalue
'     Case 4: zMultiplier = 0.4*defvalue
'            Case 5: zMultiplier = 0.45*defvalue
'            Case 6: zMultiplier = 0.5*defvalue
'        End Select
'        aBall.velz = abs(vel * zMultiplier * TargetBouncerFactor)
'        aBall.velx = sgn(aBall.velx) * sqr(abs((vel^2 - aBall.velz^2)/(1+vratio^2)))
'        aBall.vely = aBall.velx * vratio
'        debug.print "---> velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz
'        'debug.print "conservation check: " & BallSpeed(aBall)/vel
' end if
'end sub

''iaakki - TargetBouncer for standup targets
sub TargetBouncer(aBall,defvalue)
  Dim zMultiplier, vel
  'vel = BallSpeed(aBall)
  if TargetBouncerEnabled <> 0 and aball.z < 30 and aBall.vely > 0 then
    'debug.print "vely: " & activeball.vely
    Select Case Int(Rnd * 4) + 1
      Case 1: zMultiplier = defvalue+1.1
      Case 2: zMultiplier = defvalue+1.05
      Case 3: zMultiplier = defvalue+0.7
      Case 4: zMultiplier = defvalue+0.3
    End Select
    aBall.velz = aBall.velz * zMultiplier * TargetBouncerFactor
    'debug.print "----> velz: " & activeball.velz
    'debug.print "conservation check: " & BallSpeed(aBall)/vel
  end if
end sub

' *************************  script for LUT text display and rolling text *********************************
' script related to the script below is in keydown magnasave keys script and luts/lutpos script at the top of the script
' other objects used: layer 7 all Text0?? objects ("A") and the textures "32" to "96"                                                                       "

dim textindex : textindex = 1
dim charobj(55), glyph(201)
InitDisplayText

Sub myChangeLut
  table1.ColorGradeImage = luts(lutpos)
  DisplayText lutpos, luts(lutpos)
  SaveValue cGameName, "LUTPOS", lutpos
  vpmTimer.AddTimer 2000, "If lutpos = " & lutpos & " then for anr = 10 to 54 : charobj(anr).visible = 0 : next'"
End Sub

Sub InitDisplayText
  Dim anr
  For anr = 10 to 54 : set charobj(anr) = eval("text0" & anr) : charobj(anr).visible = 0 : Next
  For anr = 32 to 96 : glyph(anr) = anr : next
  For anr = 0 to 31 : glyph(anr) = 32 : next
  for anr = 97 to 122 : glyph(anr)  = anr - 32 : next
  for anr = 123 to 200 : glyph(anr) = 32 : next
End Sub

Sub DisplayText(nr, luttext)
  dim tekst, anr
  for anr = 10 to 54 : charobj(anr).imageA = 32 : charobj(anr).visible = 1 : next
  tekst = "lutpos:" & nr
  For anr = 1 to len(tekst) : charobj(43 + anr).imageA = glyph(asc(mid(tekst, anr, 1))) : Next
  For anr = 1 to len(luttext) : charobj(9 + anr).imageA = glyph(asc(mid(luttext, anr, 1))) : Next
End Sub


' Thalamus : Exit in a clean and proper way
Sub Table1_exit
  Controller.Pause = False
  Controller.Stop
End Sub


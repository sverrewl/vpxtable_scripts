'*********************************************************************
'1995 Williams DIRTY HARRY
'*********************************************************************
'*********************************************************************
'Pinball Machine designed by Barry Oursler
'*********************************************************************
'*********************************************************************
'recreated for Visual Pinball by Knorr
'*********************************************************************
'*********************************************************************
'I would like to give my sincere thanks to
'Mfuegemann, Freneticamnesic, Toxie and the VPdevs, Clark Kent and
'Gigalula for always being so friendly, helpful and motivating while
'building this table.
'*********************************************************************

Option Explicit
Randomize

' Thalamus 2018-07-20
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' Thalamus 2018-09-09 : Improved directional sounds

' !! NOTE : Table not verified yet !!

Const VolDiv = 2000    ' Lower number, louder ballrolling/collition sound
Const VolCol = 10      ' Ball collition divider ( voldiv/volcol )

' The rest of the values are multipliers
'
'  .5 = lower volume
' 1.5 = higher volume

Const VolBump   = 2    ' Bumpers volume.
Const VolBall   = 1    ' Ball volume.
Const VolGates  = 1    ' Gates volume.
Const VolMetal  = 1    ' Metals volume.
Const VolRH     = 1    ' Rubber hits volume.
Const VolPo     = 1    ' Rubber posts volume.
Const VolPi     = 1    ' Rubber pins volume.
Const VolTarg   = 1    ' Targets volume.
Const VolKick   = 1    ' Kicker volume.
Const VolSpin   = 2    ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.

Dim PinBlades

'******************************************************************************************
'* TABLE OPTIONS **************************************************************************
'******************************************************************************************

'***********  Activate Pinblades     ****************************************************

PinBlades = 1     '0 = no PinBlades, 1 = with PinBlades

'V2.0 Shiny Mod by Flupper
'Retextured warehouses with Blender generated ambient occlusion map, including the plastics below the warehouses
'Gun replaced by shiny version
'Replaced bullethole plastics with new primitives and textures
'Added Blender generated ambient occlusion map on the playfield
'Replaced environment map
'Added pinblades with option to switch on/off

'V2.0
'new Ramps (thanks to flupper!)
'added more Gi
'script clean up by JP (thanks for all the help!)
'reworked enviroment


'V1.9
'added primitive to HQ

'V1.8
'New Plastic/WireRamps
'reworked primitives
'Bug Fixes

'V1.7
'Added controller.vbs
'Bug Fixes

'V1.6
'First Release For VP10.0.0
'CrimeWave Gunfix (only 1 Ball can be in the Gun)
'reduced gun model
'reduced warehouse model (again)
'reduced png for smaller file size
'BallSize is now 52

'V1.5
'Fixed Multiball
'Changed Primitves Rampentry Metals

'V1.4
'Added Global Light for Flasher
'Added Plunger Animation
'Cleaned up RightRamp Primitive
'Added Dropwall to Warehouse so only one Ball can be in
'new Slingshot Plastics Primitives and added Walls for Collision

'V1.3
'Updated Ball Rolling/Collision Script
'small changes with primitives

'V1.2
'Fixed CrimeWave Multiball
'Added Global Lightning (thanks for helping Fren)
'Added missing Lights for BumperCap and BankRobber
'added images for ON/OFF effects
'Changed Sound for Plunger

'V1.1
'Reduced meshes in the warehouse model (almost the half)
'improved lightning for environment (thanks to Fren)
'reduced lightning for the inserts
'reduced size of the playfield.jpg and the warehouse
'minor changes with flashers

'V1.0
'First Release For VP10 Beta


On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01560000", "WPC.VBS", 3.36

'********************
'Standard definitions
'********************

Const cGameName = "dh_lx2"
Const UseSolenoids = 2
Const UseLamps = 1
Const SSolenoidOn = "SolOn"
Const SSolenoidOff = "SolOff"
Const SFlipperOn = "FlipperUp"
Const SFlipperOff = "FlipperDown"
Const SCoin = "Coin5"

Set GiCallback2 = GetRef("UpdateGI")
BSize = 25.5

' Standard Sounds
' Const SSolenoidOn = "Solenoid"
' Const SSolenoidOff = ""
' Const SFlipperOn = "FlipperUp"
' Const SFlipperOff = "FlipperDown"
' Const SCoin = "Coin"

Dim bsTrough, BallInGun, bsSafeHouse, LeftPopper, WareHousePopper, GunPopper, RightMagnet

Set LampCallback = GetRef("UpdateMultipleLamps")
Set MotorCallback = GetRef("RealTimeUpdates")

'************
' Table init.
'************

Sub Table1_Init
    vpmInit Me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
        .Games(cGameName).Settings.Value("rol") = 0
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = 0
        '.SetDisplayPosition 0, 0, GetPlayerHWnd 'uncomment this line if you don't see the DMD
        On Error Resume Next
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
        On Error Goto 0
        .Switch(22) = 1 'close coin door
        .Switch(24) = 1 'and keep it close
        PinMAMETimer.Interval = PinMAMEInterval
        PinMAMETimer.Enabled = true
        vpmNudge.TiltSwitch = 14
        vpmNudge.Sensitivity = 2
    End With

    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0, 32, 33, 34, 35, 0, 0, 0
        .InitKick BallRelease, 90, 10
        .InitExitSnd SoundFX("BallRelease", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .Balls = 4
        .IsTrough = 1
    End With

    Set bsSafeHouse = New cvpmBallStack
    bsSafeHouse.InitSaucer sw73, 73, 167, 22
    bsSafeHouse.InitExitSnd SoundFX("SafeHouseKick", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    bsSafeHouse.KickForceVar = 2
    bsSafeHouse.KickAngleVar = 0.8

    Set LeftPopper = New cvpmBallStack
    With LeftPopper
        .InitSw 0, 47, 0, 0, 0, 0, 0, 0
        .InitKick sw47, 180, 15
        .InitExitSnd SoundFX("HeadquarterKick", DOFContactors), "fx_Solenoid"
    End With

    Set WarehousePopper = New cvpmBallStack
    With WarehousePopper
        .InitSw 0, 46, 0, 0, 0, 0, 0, 0
        .InitKick sw46, 2, 10
        .KickZ = 1
        .InitExitSnd SoundFX("WareHouseKick", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .KickBalls = 1
    End With

    Set GunPopper = New cvpmBallStack
    With GunPopper
        .InitSw 0, 45, 0, 0, 0, 0, 0, 0
        .InitKick sw45, 105, 7
        .InitExitSnd SoundFX("GunPopper", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    End With

    Set RightMagnet = New cvpmMagnet
    With RightMagnet
        .InitMagnet RMagnet, 100
        .Solenoid = 35
        .GrabCenter = 1
        .CreateEvents "RightMagnet"
    End With

    DiverterOn.isDropped = 1
    DiverterOn2.isDropped = 1
    DiverterOff.isDropped = 0
    Warehousedw.isDropped = 1
    If table1.ShowDT = False then
        Ramp16.WidthTop = 0
        Ramp16.WidthBottom = 0
        Ramp15.WidthTop = 0
        Ramp15.WidthBottom = 0
    If PinBlades = 0 Then
      Korpus.visible = 1
      Korpus.Size_Y = 1.7
      primitive2.visible = 0
      primitive3.visible = 0
    Else
      Korpus.visible = 0
    End If
  Else
    primitive2.visible = 0
    primitive3.visible = 0
    Korpus.visible = 1
    If PinBlades = 1 Then
      Korpus.image = "korpuspinblade"
    End If

    End if
End Sub

'******
'Trough
'******

Sub SolRelease(Enabled)
    If Enabled Then
        If bsTrough.Balls = 4 Then vpmTimer.PulseSw 31
        If bsTrough.Balls > 0 Then bsTrough.ExitSol_On
    End If
End Sub

Sub Drain_Hit
    PlaySoundAtVol "Balltruhe", Drain, 1
    bsTrough.AddBall Me
End Sub

'*********
'Safehouse
'*********

Sub sw73_Hit
    PlaySoundAtVol "SafeHouseHit", sw73, 1
    bsSafeHouse.AddBall Me
End Sub

'**********
'LeftPopper
'**********

Dim aBall

Sub HQHole_Hit
    PlaySoundAtVol "HeadquarterHit", HQHole, 1
    Set aBall = ActiveBall:Me.TimerEnabled = 1

    LeftPopper.AddBall 1
End Sub

Sub HQHole_Timer
    Do While aBall.Z > 0
        aBall.Z = aBall.Z -5
        Exit Sub
    Loop
    Me.DestroyBall
    Me.TimerEnabled = 0
End Sub

'*********
'Warehouse
'*********

Sub WarehouseEntry_Hit 'Warehouse
  PlaySoundAtVol "WareHouseHit", WarehouseEntry, 1
  WarehousePopper.AddBall Me
  Warehousedw.isDropped = 0
End Sub

Sub Warehousedwtrigger_Hit
    Warehousedw.isDropped = 1
End Sub

'*********
'GunPopper
'*********

Sub TrapDoorKicker_Hit
  PlaySoundAtVol "HeadquarterHit", TrapDoorKicker, 1
  GunPopper.AddBall Me
End Sub

'************
'TrapDoorRamp
'************

Sub TrapDoorLow(Enabled)
    PlaySoundAtVol "TrapDoorHigh", TrapDoorP, 1
    If Enabled then
        TrapDoorP.RotX = TrapDoorP.RotX + 25
        TrapDoorKicker.Enabled = True
    Else
     TrapDoorP.RotX = TrapDoorP.RotX - 25
     PlaySoundAtVol "TrapDoorLow", TrapDoorP, 1
       TrapDoorKicker.Enabled = False
    End If
End Sub

'*********
'MagnumGun
'*********

Sub SolGunLaunch(Enabled)
     If Enabled AND BallInGun then
     vpmCreateBall GunKick
         GunKick.kick GPos, 50
         PlaySoundAtVol "GunShot", BallP, VolKick
         controller.switch(3) = 0
         BallInGun = 0
         BallP.Visible = False
  Else
    Controller.switch(44) = 0
'   sw44.Enabled = True
    vpmTimer.AddTimer 200, "sw44.Enabled = True'"
     End If
 End Sub



Sub SolGunMotor(Enabled)
    If Enabled Then
       PlaySoundAtVol SoundFX("GunMotor",DOFGear), gunscoop, 1
       GDir = -1
       UpdateGun.Enabled=1
       Controller.switch(77) = 1
     Else
       UpdateGun.Enabled=0
       Controller.switch(77) = 0
     StopSound "GunMotor"
  End If

End Sub

Dim GPos, GDir
GPos = -50
GDir = -50
Sub updategun_Timer()
    '    StopSound"": PlaySound ""
    GPos = GPos + GDir
    If GPos <= -98 Then GDir = 1
    If GPos >= -4 Then Controller.switch(76) = 1 Else Controller.switch(76) = 0
    If GPos >= -2 Then GDir = -1
    MagnumGun.RotY = GPos
End Sub

Sub sw44_hit()
    sw44.Enabled = False
    PlaySoundAtVol "BallFallInGun", ActiveBall, 1
    StopSound "WireRamp"
    RightWireStart2.Enabled = True
    Controller.switch(44) = 1
    me.DestroyBall
    BallInGun = 1
    BallP.Visible = True
End Sub

Sub GunLoadHelper_Hit()
    ActiveBall.VelX = -2
End Sub

'******
'Magnet
'******

Sub SolMagnetOn(Enabled)
    If Enabled then
        RightMagnet.MagnetOn = True
    Else
        RightMagnet.MagnetOn = False
    End if
End Sub

'*************
'RightLoopGate
'*************

Sub RightLoopGate(Enabled)
  If Enabled then
    PlaysoundAtVol "gate", sw42, 1
    GateR.open = True
  Else
  GateR.open = False
  End if
End sub

'******
'Plunger
'******

Dim AP

Sub AutoPlunge(Enabled)
    if enabled then
        AP = True
        AutoKicker.Kick 0, 48
        PlaySoundAtVol SoundFX("Plunger", DOFContactors), AutoKicker, 1
    End if
End Sub

Sub PlungerPTimer()
    if AP = True and PlungerP.TransZ < 45 then PlungerP.TransZ = PlungerP.TransZ + 10
    if AP = False and PlungerP.TransZ > 0 then PlungerP.TransZ = PlungerP.TransZ -10
    if PlungerP.TransZ >= 45 then AP = False
End Sub

'*********
'Solenoids
'*********

SolCallback(1) = "SolRelease"
SolCallback(2) = "AutoPlunge"
SolCallback(3) = "SolGunLaunch"
SolCallback(4) = "WarehousePopper.SolOut"
SolCallback(5) = "GunPopper.SolOut"
SolCallback(7) = "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
' SolCallback(8) = "TrapDoorHigh"
SolCallback(14) = "LeftPopper.SolOut"
SolCallBack(15) = "vpmSolDiverter diverterR,""DiverterRight"","
SolCallback(16) = "TrapDoorLow"
SolCallBack(20) = "SolGunMotor"
SolCallback(26) = "bsSafeHouse.SolOut"
SolCallback(27) = "SolDiverterHold"
SolCallBack(28) = "RightLoopGate"
SolCallBack(35) = "SolMagnetOn"

'*********
'Flasher
'*********

SolCallback(17) = "Multi117"
SolCallback(18) = "Multi118"
SolCallback(19) = "Multi119"
SolCallback(21) = "Multi121"
SolCallback(22) = "Multi122"
SolCallback(23) = "Multi123"
SolCallback(24) = "Multi124"

'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
    If KeyCode = MechanicalTilt Then
    If keycode = LockBarKey Then Controller.Switch(11) = 1
    If keycode = RightMagnaSave Then Controller.Switch(11) = 1
        vpmTimer.PulseSw vpmNudge.TiltSwitch
        Exit Sub
    End if

    If keycode = PlungerKey Then Controller.Switch(11) = 1
    If keycode = keyFront Then Controller.Switch(23) = 1
    If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If vpmKeyUp(keycode) Then Exit Sub
    If keycode = PlungerKey Then Controller.Switch(11) = 0
    If keycode = keyFront Then Controller.Switch(23) = 0
    If keycode = PlungerKey Then Controller.Switch(11) = 0
    If keycode = LockBarKey Then Controller.Switch(11) = 0
End Sub

'**************
' Flipper Subs
'**************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFX("FlipperUpLeft", DOFContactors), LeftFlipper, VolFlip:LeftFlipper.RotateToEnd
    Else
        PlaySoundAtVol SoundFX("FlipperDown", DOFContactors), LeftFlipper, VolFlip:LeftFlipper.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFX("FlipperUpRightBoth", DOFContactors), RightFlipper, VoLFlip:RightFlipper.RotateToEnd:RightFlipper1.RotateToEnd
    Else
        PlaySoundAtVol SoundFX("FlipperDown", DOFContactors),RightFlipper, VolFlip:RightFlipper.RotateToStart:RightFlipper1.RotateToStart
    End If
End Sub

'*********
' Switches
'*********

Sub sw15_Hit:Controller.Switch(15) = 1:sw15wire.RotX = 15:PlaySoundAtVol "metalhit_thin", ActiveBall, 1:End Sub 'shooterlane'
Sub sw15_UnHit:Controller.Switch(15) = 0:sw15wire.RotX = 0:End Sub
Sub sw16_Hit:Controller.Switch(16) = 1:sw16wire.RotX = 15:PlaySoundAtVol "metalhit_thin", ActiveBall, 1:End Sub 'right outlane'
Sub sw16_UnHit:Controller.Switch(16) = 0:sw16wire.RotX = 0:End Sub
Sub sw17_Hit:Controller.Switch(17) = 1:sw17wire.RotX = 15:PlaySoundAtVol "metalhit_thin", ActiveBall, 1:End Sub 'right inlane'
Sub sw17_UnHit:Controller.Switch(17) = 0:sw17wire.RotX = 0:End Sub
Sub sw26_Hit:Controller.Switch(26) = 1:sw26wire.RotX = 15:PlaySoundAtVol "metalhit_thin", ActiveBall, 1:End Sub 'left outlane'
Sub sw26_UnHit:Controller.Switch(26) = 0:sw26wire.RotX = 0:End Sub
Sub sw25_Hit:Controller.Switch(25) = 1:sw25wire.RotX = 15:PlaySoundAtVol "metalhit_thin", ActiveBall, 1:End Sub 'left inlane'
Sub sw25_UnHit:Controller.Switch(25) = 0:sw25wire.RotX = 0:End Sub
Sub sw71_Hit:Controller.Switch(71) = 1:sw71wire.RotX = 15:PlaySoundAtVol "metalhit_thin", ActiveBall, 1:End Sub 'left loop'
Sub sw71_UnHit:Controller.Switch(71) = 0:sw71wire.RotX = 0:End Sub
Sub sw66_Hit:Controller.Switch(66) = 1:sw66wire.RotX = 15:PlaySoundAtVol "metalhit_thin", ActiveBall, 1:End Sub 'left rollover (bumper)'
Sub sw66_UnHit:Controller.Switch(66) = 0:sw66wire.RotX = 0:End Sub
Sub sw67_Hit:Controller.Switch(67) = 1:sw67wire.RotX = 15:PlaySoundAtVol "metalhit_thin", ActiveBall, 1:End Sub 'middle rollover (bumper)'
Sub sw67_UnHit:Controller.Switch(67) = 0:sw67wire.RotX = 0:End Sub
Sub sw68_Hit:Controller.Switch(68) = 1:sw68wire.RotX = 15:PlaySoundAtVol "metalhit_thin", ActiveBall, 1:End Sub 'right rollover (bumper)'
Sub sw68_UnHit:Controller.Switch(68) = 0:sw68wire.RotX = 0:End Sub
Sub sw42_Hit:Controller.Switch(42) = 1:End Sub                                              'right loop'
Sub sw42_UnHit:Controller.Switch(42) = 0:End Sub

Sub BallDrop1_Hit():StopSound "WireRamp":End Sub
Sub BallDrop2_Hit():PlaySoundAtVol "BallDrop", ActiveBall, 1:End Sub
Sub BallDrop3_Hit():PlaySoundAtVol "BallDrop", ActiveBall, 1:End Sub
Sub ZHelper_Hit():ActiveBall.VelZ = ActiveBall.VelZ -5:End Sub

'*********
' Ramps
'*********

Sub sw41_Hit:vpmTimer.pulseSw 41:End Sub
Sub sw43_Hit:vpmTimer.pulseSw 43:End Sub
Sub sw51_Hit:vpmTimer.pulseSw 51:End Sub
Sub sw38_Hit:vpmTimer.pulseSw 38:End Sub

'RampSounds

Dim SoundBall

' Thalamus - TODO
Sub MiddleWireStart_Hit
    Set SoundBall = Activeball             'Ball-assignment
    playsound "WireRamp", -1, 0.6, 0, 0.35 '-1 = looping, Panning = -1 bis 1, je nach Position
End Sub

Sub LeftWireStart_Hit
    Set SoundBall = Activeball             'Ball-assignment
    playsound "WireRamp", -1, 0.6, 0, 0.35 '-1 = looping, Panning = -1 bis 1, je nach Position
End Sub

Sub RightWireStart1_Hit
    RightWireStart2.Enabled = False
    Set SoundBall = Activeball             'Ball-assignment
    playsound "WireRamp", -1, 0.6, 0, 0.35 '-1 = looping, Panning = -1 bis 1, je nach Position
End Sub

Sub RightWireStart2_Hit
    Set SoundBall = Activeball             'Ball-assignment
    playsound "WireRamp", -1, 0.6, 0, 0.35 '-1 = looping, Panning = -1 bis 1, je nach Position
End Sub

'***************
' StandupTargets
'***************

Sub Standup27_Hit:vpmTimer.pulseSw 27:Standup27p.RotY = Standup27p.RotY -3:PlaysoundAtVol SoundFX("target", DOFContactors), ActiveBall, 1:Me.TimerEnabled = 1:End Sub
Sub Standup27_Timer:Standup27p.RotY = Standup27p.RotY + 3:Me.TimerEnabled = 0:End Sub
Sub Standup28_Hit:vpmTimer.pulseSw 28:Standup28p.RotY = Standup28p.RotY -3:PlaysoundAtVol SoundFX("target", DOFContactors), ActiveBall, 1:Me.TimerEnabled = 1:End Sub
Sub Standup28_Timer:Standup28p.RotY = Standup28p.RotY + 3:Me.TimerEnabled = 0:End Sub
Sub Standup58_Hit:vpmTimer.pulseSw 58:Standup58p.RotY = Standup58p.RotY -3:PlaysoundAtVol SoundFX("target", DOFContactors), ActiveBall, 1:Me.TimerEnabled = 1:End Sub
Sub Standup58_Timer:Standup58p.RotY = Standup58p.RotY + 3:Me.TimerEnabled = 0:End Sub
Sub Standup57_Hit:vpmTimer.pulseSw 57:Standup57p.RotY = Standup57p.RotY -3:PlaysoundAtVol SoundFX("target", DOFContactors), ActiveBall, 1:Me.TimerEnabled = 1:End Sub
Sub Standup57_Timer:Standup57p.RotY = Standup57p.RotY + 3:Me.TimerEnabled = 0:End Sub
Sub Standup56_Hit:vpmTimer.pulseSw 56:Standup56p.RotY = Standup56p.RotY -3:PlaysoundAtVol SoundFX("target", DOFContactors), ActiveBall, 1:Me.TimerEnabled = 1:End Sub
Sub Standup56_Timer:Standup56p.RotY = Standup56p.RotY + 3:Me.TimerEnabled = 0:End Sub
Sub Standup54_Hit:vpmTimer.pulseSw 54:Standup54p.RotX = Standup54p.RotX + 3:PlaysoundAtVol SoundFX("target", DOFContactors), ActiveBall, 1:Me.TimerEnabled = 1:End Sub
Sub Standup54_Timer:Standup54p.RotX = Standup54p.RotX -3:Me.TimerEnabled = 0:End Sub
Sub Standup55_Hit:vpmTimer.pulseSw 55:Standup55p.RotY = Standup55p.RotY -3:PlaysoundAtVol SoundFX("target", DOFContactors), ActiveBall, 1:Me.TimerEnabled = 1:End Sub
Sub Standup55_Timer:Standup55p.RotY = Standup55p.RotY + 3:Me.TimerEnabled = 0:End Sub
Sub Standup18_Hit:vpmTimer.pulseSw 18:Standup18p.RotY = Standup18p.RotY -3:PlaysoundAtVol SoundFX("target", DOFContactors), ActiveBall, 1:Me.TimerEnabled = 1:End Sub
Sub Standup18_Timer:Standup18p.RotY = Standup18p.RotY + 3:Me.TimerEnabled = 0:End Sub

'*********
' Bumper
'*********

Sub Bumper63_hit:vpmTimer.pulseSw 63:PlaysoundAtVol SoundFX("BumperLeft", DOFContactors), ActiveBall, 1:Me.TimerEnabled = 1:End Sub
Sub Bumper63_Timer:Me.Timerenabled = 0:End Sub

Sub Bumper64_hit:vpmTimer.pulseSw 64:PlaysoundAtVol SoundFX("BumperMiddle", DOFContactors), ActiveBall, 1:Me.TimerEnabled = 1:End Sub
Sub Bumper64_Timer:Me.Timerenabled = 0:End Sub

Sub Bumper65_hit:vpmTimer.pulseSw 65:PlaysoundAtVol SoundFX("BumperRight", DOFContactors), ActiveBall, 1:Me.TimerEnabled = 1:End Sub
Sub Bumper65_Timer:Me.Timerenabled = 0:End Sub

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
    PlaySoundAtVol SoundFX ("SlingshotLeft",DOFContactors), sling1, 1
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
    vpmTimer.PulseSw 62
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
    PlaySoundAtVol SoundFX("SlingshotRight",DOFContactors), sling2, 1
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
    vpmTimer.PulseSw 61
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

'********
'Diverter
'********

Sub SolDiverterHold(Enabled)
    DiverterOFF.IsDropped = Enabled:DiverterOn.IsDropped = Not Enabled
    DiverterOn2.IsDropped = Not Enabled
    If Enabled then Playsound "DiverterLeft":End if
End Sub

'*********
'Update GI
'*********

Dim xx
Dim gistep
gistep = 1 / 8

Sub UpdateGI(no, step)
    If step = 0 OR step = 7 then exit sub
    Select Case no

        'Bottom String
        Case 4
            For each xx in GIString1:xx.IntensityScale = gistep * step:next
            if step = 1 then Table1.ColorGradeImage = "-70"
            if step = 2 then Table1.ColorGradeImage = "-60"
            if step = 3 then Table1.ColorGradeImage = "-50"
            if step = 4 then Table1.ColorGradeImage = "-40"
            if step = 5 then Table1.ColorGradeImage = "-30"
            if step = 6 then Table1.ColorGradeImage = "-20"
            if step = 7 then Table1.ColorGradeImage = "-10"
            if step = 8 then Table1.ColorGradeImage = "ColorGradeLUT256x16_ConSat"

        'Left String
        Case 1
            For each xx in GIString2:xx.IntensityScale = gistep * step:next

        'Right String
        Case 0
            For each xx in GIString3:xx.IntensityScale = gistep * step:next
    End Select
End Sub

'*************
'  VP Lights
'*************

InitLamps

Sub InitLamps()
    Set Lights(11) = l11
    Set Lights(12) = l12
    Set Lights(13) = l13
    Set Lights(14) = l14
    Set Lights(15) = l15
    Set Lights(16) = l16
    Set Lights(17) = l17
    Set Lights(18) = l18
    Set Lights(21) = l21
    Set Lights(22) = l22
    Set Lights(23) = l23
    Set Lights(24) = l24
    Set Lights(25) = l25
    Set Lights(26) = l26
    Set Lights(27) = l27
    Set Lights(28) = l28
    Set Lights(31) = l31
    Set Lights(32) = l32
    Set Lights(33) = l33
    Set Lights(34) = l34
    Set Lights(35) = l35
    Set Lights(36) = l36
    Set Lights(37) = l37
    Set Lights(38) = l38
    Set Lights(41) = l41
    Set Lights(42) = l42
    Set Lights(43) = l43
    Set Lights(44) = l44
    Set Lights(45) = l45
    Set Lights(46) = l46
    Set Lights(47) = l47
    Set Lights(48) = l48
    Set Lights(51) = l51
    Set Lights(52) = l52
    Set Lights(53) = l53
    Set Lights(54) = l54
    Set Lights(55) = l55
    Set Lights(56) = l56
    Set Lights(57) = l57
    Set Lights(58) = l58
    Set Lights(61) = l61
    Set Lights(62) = l62
    Set Lights(63) = l63
    Set Lights(64) = l64
    Set Lights(65) = l65
    Set Lights(66) = l66
    Set Lights(67) = l67
    Set Lights(68) = l68
    Set Lights(77) = l77
    Set Lights(78) = l78
    Set Lights(81) = l81
    Set Lights(82) = l82
    Set Lights(83) = l83
    Set Lights(84) = l84
    Set Lights(85) = l85
    Set Lights(86) = l86
End Sub

Sub UpdateMultipleLamps
    If l48.state = 1 then bulbyellow.image = "bulbcover1_yellowOn":l84a.state = 1:else bulbyellow.image = "bulbcover1_yellow":l84a.state = 0
    If l47.state = 1 then bulbred.image = "bulbcover1_redOn":else bulbred.image = "bulbcover1_red"
    If l85.state = 1 then domesmall.image = "domesmallredOn":else domesmall.image = "domesmallred"
    If l84.state = 1 then domesmall1.image = "domesmallredOn":else domesmall1.image = "domesmallred"
    If l38.state = 1 then Bankrobber.image = "bankrobbermesh1On":else Bankrobber.image = "bankrobbermesh1"
End Sub

Sub Multi117(Enabled)
    If Enabled Then
        l117a.State = 1
        l117b.state = 1
        l117c.state = 1
        l117d.state = 1
        l117e.state = 1
        Dome1.Image = "dome3_orange_On"
        domesmall2.Image = "domesmallredOn"
    Else
        l117a.State = 0
        l117b.state = 0
        l117c.state = 0
        l117d.state = 0
        l117e.state = 0
        Dome1.Image = "dome3_orange"
        domesmall2.Image = "domesmallred"
    End If
End Sub

Sub Multi118(Enabled)
    If Enabled Then
        l118a.State = 1
        l118b.state = 1
        l118c.state = 1
        l118d.state = 1
        l118e.state = 1
        l118f.state = 1
    Else
        l118a.State = 0
        l118b.state = 0
        l118c.state = 0
        l118d.state = 0
        l118e.state = 0
        l118f.state = 0
    End If
End Sub

Sub Multi119(Enabled)
    If Enabled Then
        l119a.State = 1
        l119b.state = 1
        l119c.state = 1
        l119d.state = 1
        l119e.state = 1
        l119f.state = 1
    Else
        l119a.State = 0
        l119b.state = 0
        l119c.state = 0
        l119d.state = 0
        l119e.state = 0
        l119f.state = 0
    End If
End Sub

Sub Multi122(Enabled)
    If Enabled Then
        l122a.State = 1
        l122b.state = 1
    Else
        l122a.State = 0
        l122b.state = 0
    End If
End Sub

Sub Multi123(Enabled)
    If Enabled Then
        l123a.State = 1
        l123b.state = 1
        l123c.state = 1
        l123d.state = 1
        l123ab.state = 1
        l123ab1.state = 1
        Dome5.Image = "dome3_blue_On"
        Dome3.Image = "dome3_clear_On"
        GIWhite1.State = 1
    Else
        l123a.State = 0
        l123b.state = 0
        l123c.state = 0
        l123d.state = 0
        l123ab.state = 0
        l123ab1.state = 0
        Dome5.Image = "dome3_blue"
        Dome3.Image = "dome3_clear"
        GIWhite1.State = 0
    End If
End Sub

Sub Multi124(Enabled)
    If Enabled Then
        l124a.State = 1
        l124b.state = 1
        l124ab.state = 1
        l124ab1.state = 1
        l124c.state = 1
        l124d.state = 1
        Dome2.Image = "dome3_clear_On"
        Dome4.Image = "dome3_blue_On"
        GIWhite1.State = 1
    Else
        l124a.State = 0
        l124b.state = 0
        l124c.state = 0
        l124d.state = 0
        l124ab.State = 0
        l124ab1.state = 0
        Dome2.Image = "dome3_clear"
        Dome4.Image = "dome3_blue"
        GIWhite1.State = 0
    End If
End Sub

Sub Multi121(Enabled)
    If Enabled Then
        l121a.State = 1
        l121b.state = 1
        l121c.state = 1
    Else
        l121a.State = 0
        l121b.state = 0
        l121c.state = 0
    End If
End Sub

' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX
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

'Set position as bumperX and Vol manually.

Sub PlaySoundAtBumperVol(sound, tableobj, Vol)
  PlaySound sound, 1, Vol, AudioPan(tableobj), 0,0,1, 1, AudioFade(tableobj)
End Sub

'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

Function RndNum(min, max)
    RndNum = Int(Rnd() * (max-min + 1) ) + min ' Sets a random number between min and max
End Function

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

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / table1.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
  Vol = Csng(BallVel(ball) ^2 / VolDiv)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
  Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
  BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

Function Vol2(ball1, ball2) ' Calculates the Volume of the sound based on the speed of two balls
    Vol2 = (Vol(ball1) + Vol(ball2) ) / 2
End Function

'*****************************************
'    JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 9 ' total number of balls
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = FALSE
    Next
End Sub

Sub RollingSoundUpdate()
    Dim BOT, b
    BOT = GetBalls

    ' stop the sound of deleted balls
    For b = UBound(BOT) + 1 to tnob
        rolling(b) = FALSE
        StopSound("fx_ballrolling" & b)
    Next

    ' exit the sub if no balls on the table
    If UBound(BOT) = -1 Then Exit Sub

    ' play the rolling sound for each ball

    For b = 0 to UBound(BOT)
      If BallVel(BOT(b) ) > 1 Then
        rolling(b) = True
        if BOT(b).z < 30 Then ' Ball on playfield
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), AudioPan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        Else ' Ball on raised ramp
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*.5, AudioPan(BOT(b) ), 0, Pitch(BOT(b) )+50000, 1, 0, AudioFade(BOT(b) )
        End If
      Else
        If rolling(b) = True Then
          StopSound("fx_ballrolling" & b)
          rolling(b) = False
        End If
      End If
    Next
End Sub

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / (VolDiv/VolCol), AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

'****************************
'     Realtime Updates
' called by the MotorCallBack
'****************************

Sub RealTimeUpdates
    'flippers
    LeftFlipperP.RotY = LeftFlipper.CurrentAngle
    RightFlipperP1.RotY = RightFlipper1.CurrentAngle
    RightFlipperP.RotY = RightFlipper.CurrentAngle
    ' rolling sound
    RollingSoundUpdate
    ' Plunger update
    PlungerPTimer
    ' ramp gate
    SpinnerP.RotX = Spinner1.currentangle + 95
    ' diverter
    DiverterP.RotY = diverterR.CurrentAngle
End Sub

'******************************
' Diverse Collection Hit Sounds
'******************************

Sub Pins_Hit(idx)
    PlaySound "pinhit_low", 0, Vol(ActiveBall)*VolPi, Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
End Sub

Sub Targets_Hit(idx)
    PlaySound SoundFX("target", DOFContactors)*VolTarg, 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
End Sub

Sub Metals_Thin_Hit(idx)
    PlaySound "metalhit_thin", 0, Vol(ActiveBall)*VolMetal, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub Metals_Medium_Hit(idx)
    PlaySound "metalhit_medium", 0, Vol(ActiveBall)*VolMetal, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub Metals2_Hit(idx)
    PlaySound "metalhit2", 0, Vol(ActiveBall)*VolMetal, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub Gates_Hit(idx)
    PlaySound "gate4", 0, Vol(ActiveBall)*VolGates, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub Spinner_Spin
    PlaySoundAtVol "fx_spinner", Spinner, VolSpin
End Sub

Sub Rubbers_Hit(idx)
    dim finalspeed
    finalspeed = SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
    If finalspeed > 20 then
        PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    End if
    If finalspeed >= 6 AND finalspeed <= 20 then
        RandomSoundRubber()
    End If
End Sub

Sub Posts_Hit(idx)
    dim finalspeed
    finalspeed = SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
    If finalspeed > 16 then
        PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolPo, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    End if
    If finalspeed >= 6 AND finalspeed <= 16 then
        RandomSoundRubber()
    End If
End Sub

Sub RandomSoundRubber()
    Select Case Int(Rnd * 3) + 1
        Case 1:PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
        Case 2:PlaySound "rubber_hit_2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
        Case 3:PlaySound "rubber_hit_3", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    End Select
End Sub

Sub LeftFlipper_Collide(parm)
    RandomSoundFlipper()
End Sub

Sub RightFlipper_Collide(parm)
    RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
    Select Case Int(Rnd * 3) + 1
        Case 1:PlaySound "flip_hit_1", 0, 1, Pan(ActiveBall)*VolRH, 0, Pitch(ActiveBall), 1, 0
        Case 2:PlaySound "flip_hit_2", 0, 1, Pan(ActiveBall)*VolRH, 0, Pitch(ActiveBall), 1, 0
        Case 3:PlaySound "flip_hit_3", 0, 1, Pan(ActiveBall)*VolRH, 0, Pitch(ActiveBall), 1, 0
    End Select
End Sub


' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  On Error Resume Next
  Controller.Pause = False
  Controller.Stop
End Sub


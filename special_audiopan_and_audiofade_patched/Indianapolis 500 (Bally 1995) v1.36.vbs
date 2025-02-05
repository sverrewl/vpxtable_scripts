' Indianapolis 500 - Bally 1995
' JPSalas - VP9 v1.0
' Dorsola - Turbo Handler Code
' Dozer - VP10 Conversion / Enhancements - August 2017
' TastyWasps - VPW Enhancements & VR Room - April 2023

Option Explicit
Randomize
Const UseVPMModSol = 1

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

'*******************************************
'  Table Options
'*******************************************
Const ToyMod = 1

' Render a rotating Indy Car sign up the back left of the Playfield.
Const Rotating_Track = 1

' Change the color of the upper Indy Car Toy.
' 1 = Black / 2 = Blue / 3 = Red
Const Rotating_Car_Color = 1

' Rotate the track under the upper Indy Car Toy.
' 1 = Original White / 2 = Asphalt with line / 3 = Asphalt with broken line
Const Track_Type = 3

' With this on, the VUK hole to feed the turbo will catch balls less precisely. The real
' game suffers from random ball fails when shot into this VUK. Turn off for precise kicker grabs.
Const Sloppy_Turbo_Vuk = 0

' Render a relection of the turbo on the left sidewall.
Const Turbo_Reflection = 0

' Flipper colors
' 1 = Blue / 2 = Black
Const Flipper_Bat_Color = 2

'Change the color of the lamps in the grid light targets on the playfield.
' 1=Red / 2=Green / 3=Blue / 4 = Yellow / 5=Purple / 6=Cyan / 7=Orange
Const GridLamp = 1

'Change the color of the lamp underneath the turbo impellor on the turbo mech.
' 1=Red / 2=Green / 3=Blue / 4 = Yellow / 5=Purple / 6=Cyan / 7=Orange
Const TurboLamp = 1

' Sound Volume Options (0-1)
Const VolumeDial = 0.8       ' Mechanical sounds.
Const BallRollVolume = 0.5   ' Level of ball rolling volume.
Const RampRollVolume = 0.5   ' Level of ramp rolling volume.

' Target Bouncer Options
Const TargetBouncerEnabled = 1    ' 0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.7   ' Level of bounces. Recommmended value of 0.7

' Ball Shadow Options
Const DynamicBallShadowsOn = 1    '0 = no dynamic ball shadow ("triangles" near slings and such), 1 = enable dynamic ball shadow
Const AmbientBallShadowOn = 1   '0 = Static shadow under ball ("flasher" image, like JP's)
'                 '1 = Moving ball shadow ("primitive" object, like ninuzzu's) - This is the only one that shows up on the pf when in ramps and fades when close to lights!
'                 '2 = flasher image shadow, but it moves like ninuzzu's

' Dozer's "Northern Lights" option - High lighting reflections above playfield for interesting effects. 1 = On 0 = Off
Dim Dozer_Northern_Lights
Dozer_Northern_Lights = 1

Const Dozer_Cab = 0
' Leave this OFF!

'----- VR Room Auto-Detect -----
Dim VRRoom, VR_Obj

If RenderingMode = 2 Then

  Dozer_Northern_Lights = 0  ' Set to 1 for interesting non standard lighting reflections above playfield in VR.

  For Each VR_Obj in VRCabinet : VR_Obj.Visible = 1 : Next

  Ramp15.Visible = 0
  Ramp13.Visible = 0
  Wall13.Visible = 0
  Wall13.Sidevisible = false
  Wall71.Visible = 0
  Wall71.Sidevisible = false
  Wall72.Visible = 0
  Wall72.Sidevisible = false
  Backdrop.Visible = 0
  Backdrop.Sidevisible = false

  For Each VR_Obj in VRMinimalRoom : VR_Obj.Visible = 1 : Next

Else
  For Each VR_Obj in VRCabinet : VR_Obj.Visible = 0 : Next
  For Each VR_Obj in VRMinimalRoom : VR_Obj.Visible = 0 : Next
End If

'*******************************************
'  Constants and Global Variables
'*******************************************
Const BallSize = 50
Const BallMass = 1
Const tnob = 5

Dim varHidden, UseVPMDMD, TableWidth, TableHeight, gilvl

TableWidth = table1.width
TableHeight = table1.height

gilvl = 1 ' General Illumination light state tracked for Dynamic Ball Shadows

If table1.ShowDT = True then
  UseVPMDMD = true
  VarHidden = 1
Else
  UseVPMDMD = false
  VarHidden = 0
End If

'*******************************************
'  Customization Code from User Options
'*******************************************

If Dozer_Northern_Lights = 0 Then
  ref1.Visible = 0
  ref2.Visible = 0
  ref3.Visible = 0
  ref4.Visible = 0
  ref5.Visible = 0
  ref6.Visible = 0
  ref7.Visible = 0
' f26_Side_Flash1.Visible = 0
End If

If toymod = 1 Then
  modtoy.enabled = 1
  track1.visible = 1
  indy_shaft1.visible = 1
  modref.visible = 1
  TopCar_Base2.visible = 1
  Prim_Spot3.visible = 1
  lamp_post.visible = 1
End If

If Rotating_Car_Color = 1 Then
  Indy_Top.image = "vi0_main_black"
ElseIf Rotating_Car_Color = 2 Then
  Indy_Top.image = "vi0_main_blue"
Else
  Indy_Top.image = "vi0_main"
End If

If Flipper_Bat_Color = 1 Then
  batleft.image = "flipper_white_blue"
  batright.image = "flipper_white_blue"
  batright1.image = "flipper_white_blue"
Else
  batleft.image = "flipper_white_black"
  batright.image = "flipper_white_black"
  batright1.image = "flipper_white_black"
End If

If Track_Type = 1 Then
  track.image = "track_white"
End If

If Track_Type = 2 Then
  track.image = "track_line"
End If

If Track_Type = 3 Then
  track.image = "track_broken"
End If

If Turbo_Reflection = 1 Then
  TREF.visible = 1
  Turbo_Side_Flash2.opacity = 20
Else
  TREF.visible = 0
  Turbo_Side_Flash2.opacity = 0
End If

If Sloppy_Turbo_VUK = 1 Then
  SW62.enabled = 0
Else
  SW62.enabled = 1
End If

Dim Bulb

If gridlamp = 1 Then
  For Each bulb in tlights
    bulb.Color=RGB(255,0,0)
  Next
End If

If gridlamp = 2 Then
  For Each bulb in tlights
    bulb.Color=RGB(0,255,0)
  Next
End If

If gridlamp = 3 Then
  For Each bulb in tlights
    bulb.Color=RGB(0,0,255)
  Next
End If

If gridlamp = 4 Then
  For Each bulb in tlights
    bulb.Color=RGB(255,255,0)
  Next
End If

If gridlamp = 5 Then
  For Each bulb in tlights
    bulb.Color=RGB(255,0,255)
  Next
End If

If gridlamp = 6 Then
  For Each bulb in tlights
    bulb.Color=RGB(0,255,255)
  Next
End If

If gridlamp = 7 Then
  For Each bulb in tlights
    bulb.Color=RGB(255,102,0)
  Next
End If

If turbolamp = 1 Then
  For Each bulb in tblights
    bulb.Color=RGB(255,0,0)
  Next
End If

If turbolamp = 2 Then
  For Each bulb in tblights
    bulb.Color=RGB(0,255,0)
  Next
End If

If turbolamp = 3 Then
  For Each bulb in tblights
    bulb.Color=RGB(0,0,255)
  Next
End If

If turbolamp = 4 Then
  For Each bulb in tblights
    bulb.Color=RGB(255,255,0)
  Next
End If

If turbolamp = 5 Then
  For Each bulb in tblights
    bulb.Color=RGB(255,0,255)
  Next
End If

If turbolamp = 6 Then
  For Each bulb in tblights
    bulb.Color=RGB(0,255,255)
  Next
End If

If turbolamp = 7 Then
  For Each bulb in tblights
    bulb.Color=RGB(255,102,0)
  Next
End If

LoadVPM "02800000", "WPC.VBS", 3.55

' Init Table
Const cGameName = "i500_11r"
Const UseSolenoids = 2
Const cSingleLFlip = 0
'Const cSingleRFlip = 0 ' Commented out for staged flipping.
Const UseLamps = 1
Const UseGI = 0
Const UseSync = 0
Const FlasherTest = 0

Dim x, bsTrough, bsLE, bsUE, bsTBP, bsBBP, bump1, bump2, bump3, BallInShooterLane, TopCarPos, MSpinMagnet, MSpinMagnet1

' Standard Sounds
Const SSolenoidOn = "Solenoid"    'Solenoid activates
Const SSolenoidOff = ""           'Solenoid deactivates
Const SFlipperOn = "FlipperUp"    'Flipper activated
Const SFlipperOff = "FlipperDown" 'Flipper deactivated
Const SCoin = "Coin"              'Coin inserted


' Table init.
Sub Table1_Init
  Dim i
  With Controller
    vpmInit Me
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
    .SplashInfoLine = "Indianapolis 500 - (Bally 1995)"
    .HandleKeyboard = 0
    .ShowTitle = 0
    .ShowDMDOnly = 1
    .ShowFrame = 0
    .HandleMechanics = 1
    .Hidden = VarHidden
    .Games(cGameName).Settings.Value("sound") = 1
    '.Games(cGameName).Settings.Value("rol") = 0 'Do Not Rotate
    On Error Resume Next
    .Run GetPlayerHWnd
    If Err Then MsgBox Err.Description
    On Error Goto 0
  End With

  Controller.Switch(22) = 1 'door closed
  Controller.Switch(24) = 0 'door always closed

  ' Main Timer init
  PinMAMETimer.Interval = PinMAMEInterval
  PinMAMETimer.Enabled = 1
  'StopShake 'StartShake

  ' Nudging
  vpmNudge.TiltSwitch = 14
  vpmNudge.Sensitivity = 5
  vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingshot)

  Set bsTrough = New cvpmTrough
  With bsTrough
    .Size = 4
    .InitSwitches Array(42, 43, 44, 45)
    .InitExit BallRelease, 60, 6
    .Balls = 4
  End With

  Set mSpinMagnet = New cvpmMagnet
  With mSpinMagnet
    .InitMagnet SpinMagnet, 70
    '.Solenoid = 35 'own solenoid sub
    .GrabCenter = 0
    .Size = 100
    .CreateEvents "mSpinMagnet"
  End With

  ' Lower Eject
  Set bsLE = New cvpmBallStack
  bsLE.InitSaucer sw65, 65, 328+(RND*5), 23+(RND*2)

  ' Upper Eject
  Set bsUE = New cvpmBallStack
  bsUE.InitSaucer sw64, 64, 50, 19
  bsUE.KickAngleVar = 4
  bsUE.KickForceVar = 4

  ' Top Ball Popper
  Set bsTBP = New cvpmBallStack
  bsTBP.InitSaucer sw61, 61, 0, 0
  bsTBP.InitKick sw61a, 140, 9

  ' Bottom Ball Popper
  Set bsBBP = New cvpmBallStack
  bsBBP.InitSaucer sw62, 62, 0, 0
  bsBBP.InitKick sw62a, 330, 16

  DiverterON.IsDropped = 1
  BallInShooterLane = 0
  TopCarPos = 1
  Plunger.Pullback
  Init_Turbo
End Sub

Sub Table1_Paused:Controller.Pause = True:End Sub
Sub Table1_unPaused:Controller.Pause = False:End Sub

' Keys
Sub Table1_KeyDown(ByVal Keycode)

  If keycode = LeftFlipperKey Then
    FlipperActivate LeftFlipper, LFPress
    PinCab_LeftFlipperButton.X = PinCab_LeftFlipperButton.X + 10
  End If

  If keycode = RightFlipperKey Then
    FlipperActivate RightFlipper, RFPress
    PinCab_RightFlipperButton.X = PinCab_RightFlipperButton.X - 10
  End If

  If keycode = KeyUpperRight Then
    SolURFlipper 1
  End If

  If keycode = LeftTiltKey Then Nudge 90, 1:SoundNudgeLeft()
  If keycode = RightTiltKey Then Nudge 270, 1:SoundNudgeRight()
  If keycode = CenterTiltKey Then Nudge 0, 2:SoundNudgeCenter()

  If keycode = StartGameKey Then Controller.Switch(13) = 1: soundStartButton()
  If keycode = keyFront Then Controller.Switch(23) = 1

  If keycode = keyInsertCoin1 or keycode = keyInsertCoin2 or keycode = keyInsertCoin3 or keycode = keyInsertCoin4 Then
    Select Case Int(rnd*3)
      Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
      Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
      Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
    End Select
  End If

  If keycode = PlungerKey Then
    Controller.Switch(11) = 1
    If BallInShooterLane = 1 Then
      SoundPlungerReleaseBall()
    Else
      SoundPlungerReleaseNoBall()
    End If
    LaunchButton.Y = 755.8526
    Flasher1.Y = 2442.821
  End If

  If vpmKeyDown(keycode) Then Exit Sub

End Sub

Sub Table1_KeyUp(ByVal Keycode)

  If keycode = StartGameKey Then Controller.Switch(13) = 0
  If keycode = keyFront Then Controller.Switch(23) = 0

  If keycode = PlungerKey Then
    Controller.Switch(11) = 0
    LaunchButton.Y = 759.8526
    Flasher1.Y = 2446.821
  End If

  If keycode = LeftFlipperKey Then
    PinCab_LeftFlipperButton.X = PinCab_LeftFlipperButton.X - 10
  End If

  If keycode = RightFlipperKey Then
    PinCab_RightFlipperButton.X = PinCab_RightFlipperButton.X + 10
  End If

  If keycode = KeyUpperRight Then
    SolURFlipper 0
  End If

  If vpmKeyUp(keycode) Then Exit Sub

End Sub

' Solenoids
SolCallback(1) = "SolPlungBall"
SolCallback(2) = "SolTopPopper"
SolCallback(3) = "UpMidSaucer"
SolCallback(4) = "slot_kicker"
SolCallback(5) = "SolBottomPopper"
SolCallback(7) = "Knocker"
SolCallback(13) = "SolBallRelease"
SolModCallback(14) = "Sol14"
SolModCallback(15) = "Sol15"
SolModCallback(16) = "Sol16"
SolCallback(18) = "SolTopCar"
SolModCallback(19) = "Sol19"
SolModCallback(20) = "Sol20"
SolModCallback(21) = "Sol21"
SolModCallback(22) = "Sol22"
SolModCallback(23) = "Sol23"
SolModCallback(24) = "Sol24"
SolModCallback(25) = "Sol25"
SolModCallback(26) = "Sol26"
SolModCallback(27) = "Sol27"
SolModCallback(28) = "Sol28"
SolCallback(36) = "SolDiverterHold"

' Solenoid Subs
Sub Knocker(Enabled)
  If enabled Then
    KnockerSolenoid
  End If
End Sub

Sub Sol19(Enabled)
  F19a.Intensity = Enabled / 10
  If enabled > 1 Then
    F19.State = 1
  Else
    F19.State = 0
  End If
End Sub

Sub Sol20(Enabled)
  F20a.Intensity = Enabled / 10
  If enabled > 1 Then
    F20.State = 1
  Else
    F20.State = 0
  End If
End Sub

Sub Sol21(Enabled)
  F21a.Intensity = Enabled / 10
  If enabled > 1 Then
    F21.State = 1
    F21r.State = 1
  Else
    F21.State = 0
    F21r.State = 0
  End If
End Sub

Sub Sol22(Enabled)
  F22a.Intensity = Enabled / 10
  If enabled > 1 Then
    F22.State = 1
  Else
    F22.State = 0
  End If
End Sub

Sub Sol26(Enabled)
  F26_Lamp1.Intensity = Enabled / 10
  F26_Lamp2.Intensity = Enabled / 20
  F26_Side_Flash1.opacity = Enabled / 2
End Sub

Sub Sol16(Enabled)
  F16_Lamp1.Intensity = Enabled / 10
  F16_Lamp2.Intensity = Enabled / 20
  F16_Side_Flash1.opacity = Enabled / 2
End Sub

Sub Sol14(Enabled) ' Upper Red Dome Flasher
  ObjTargetLevel(4) = Enabled/255
  FlasherFlash4_Timer
  F14_Lamp1.Intensity = Enabled / 10
  F14_Lamp2.Intensity = Enabled / 10
  F14_Lamp3.Intensity = Enabled / 10

  If toymod = 1 Then
    t1.Intensity = Enabled / 10
    Turbo_Side_Flash3.opacity = Enabled / 2
    Turbo_Side_Flash1.opacity = Enabled / 2
  End If

  If Enabled > 1 AND toymod = 1 Then
    Track1.image = "disc2Db_asp2log"
    Track3.visible = 1
    Indy_Shaft3.visible = 1
  Else
    Track1.image = "disc2Db_asp2log_dark"
    Track3.visible = 0
    Indy_Shaft3.visible = 0
  End If

  F14_Side_Flash1.opacity = Enabled / 2
  If enabled > 1 Then
    F14_Side_Flash2.visible = 1
  Else
    F14_Side_Flash2.visible = 0
  End If
End Sub

Sub Sol15(Enabled) ' Upper Blue Dome Flasher
  ObjTargetLevel(3) = Enabled/255
  FlasherFlash3_Timer
  F15_Lamp1.Intensity = Enabled / 10
  F15_Lamp2.Intensity = Enabled / 10
  F15_Lamp3.Intensity = Enabled / 10
  F15_Side_Flash1.opacity = Enabled / 2
  F15_Side_Flash2.opacity = Enabled / 2

  If toymod = 1 Then
    t1.Intensity = Enabled / 10
    Turbo_Side_Flash3.opacity = Enabled / 2
    Turbo_Side_Flash1.opacity = Enabled / 2
  End If

  If enabled > 1 Then
    F15_Side_Flash3.visible = 1
  Else
    F15_Side_Flash3.visible = 0
  End If

  If Enabled > 1 AND toymod = 1 Then
    Track3.visible = 1
    Indy_Shaft3.visible = 1
    Track1.image = "disc2Db_asp2log"
  Else
    Track3.visible = 0
    Indy_Shaft3.visible = 0
    Track1.image = "disc2Db_asp2log_dark"
  End If
End Sub

Sub Sol23(Enabled)
  Flash23_Lamp1.Intensity = Enabled / 10
  Flash23_Lamp2.Intensity = Enabled / 10
  Flash23_Lamp3.Intensity = Enabled / 10
  Flash23_Lamp4.Intensity = Enabled / 10
End Sub

Sub Sol24(Enabled)
  Flash24_Lamp1.Intensity = Enabled / 10
  Flash24_Lamp2.Intensity = Enabled / 10
  Flash24_Lamp3.Intensity = Enabled / 10
  Flash24_Lamp4.Intensity = Enabled / 10
End Sub

Sub Sol25(Enabled)
  Flash25_Lamp1.Intensity = Enabled / 10
  Flash25_Lamp2.Intensity = Enabled / 10
  Flash25_Lamp3.Intensity = Enabled / 10
  Flash25_Lamp4.Intensity = Enabled / 10
End Sub

Sub Sol27(Enabled) ' Lower Left Red Flasher
  ObjTargetLevel(1) = Enabled/255
  FlasherFlash1_Timer
  F27_Lamp1.Intensity = Enabled / 10
  F27_Lamp2.Intensity = Enabled / 10
  F27_Lamp3.Intensity = Enabled / 10
  F27_Lamp4.Intensity = Enabled / 10
  F27_Lamp5.Intensity = Enabled / 20
  F27_Lamp6.Intensity = Enabled / 10
  F27_Lamp7.Intensity = Enabled / 10
  F27_Lamp8.Intensity = Enabled / 10
  F27_Lamp9.Intensity = Enabled / 10
  F27_Lamp10.Intensity = Enabled / 10
  F27_Lamp11.Intensity = Enabled / 20
  F27_LowBoy.Intensity = Enabled / 20
  F27_Flash.opacity = Enabled / 2
  F27_Side_Flash.opacity = Enabled / 2
  F27_Side_Flash1.opacity = Enabled / 2
End Sub

Sub Sol28(Enabled) ' Lower Right Yellow Flasher
  ObjTargetLevel(2) = Enabled/255
  FlasherFlash2_Timer
  F28_LowBoy.Intensity = Enabled / 20
  F28_Lamp1.Intensity = Enabled / 10
  F28_Lamp2.Intensity = Enabled / 10
  F28_Lamp3.Intensity = Enabled / 10
  F28_Lamp4.Intensity = Enabled / 10
  F28_Lamp5.Intensity = Enabled / 20
  F28_Flash1.opacity = Enabled / 2
  F28_Side_Flash.opacity = Enabled / 2
  If enabled > 1 Then
    F17s4.visible = 1
  Else
    F17s4.visible = 0
  End If
End Sub

Sub UpMidSaucer(Enabled)
  If enabled Then
    bsUE.ExitSol_on
    SoundSaucerKick 1, sw64
  End If
End Sub

Sub SolPlungBall(Enabled)
  If Enabled Then
    Plunger.Fire
    PlaySoundat SoundFX("popper",DOFContactors), sw25
    PlaySoundat "rail_low_slower", sw25
  End If
End Sub

Sub SolBallRelease(Enabled)
  If Enabled Then
    bsTrough.ExitSol_On
    RandomSoundBallRelease BallRelease
    If bsTrough.Balls Then
      vpmTimer.PulseSw 41
    End If
  End If
End Sub

Sub SolDiverterHold(Enabled)
  If Enabled Then
    PlaySoundat SoundFX("Solenoid",DOFContactors), diverterhold
    ibpos = 1
    Indy_Lower.enabled = 1
    DiverterOff.IsDropped = 1
    'carr.state=ABS(carr.state-1)
    DiverterOn.IsDropped = 0
    If Dozer_Cab = 1 Then
      DOF 166,2
    End If
  Else
    PlaySoundat SoundFX("Solenoid",DOFContactors), diverterhold
    ibpos = 2
    Indy_Lower.enabled = 1
    DiverterOff.IsDropped = 0
    'carr.state=ABS(carr.state-1)
    DiverterOn.IsDropped = 1
    DiverterHold.Kick 180, 3
    If Dozer_Cab = 1 Then
      DOF 166,2
    End If
  End If
End Sub

'Lower Racecar Diverter Code.
Dim ibpos

ibpos = 1

Sub Indy_Lower_Timer()
  Select Case ibpos
    Case 1:
      If Indy_Bottom.objrotz <= -20 Then
        Indy_Bottom.objrotz = -20
        TL1.State = 2
        TL2.State = 2
        me.enabled = 0
      End If
      Indy_Bottom.objrotz = Indy_Bottom.objrotz - 1
    Case 2:
      If Indy_Bottom.objrotz => 0 Then
        Indy_Bottom.objrotz = 0
        TL1.State = 0
        TL2.State = 0
        me.enabled = 0
      End If
    Indy_Bottom.objrotz = Indy_Bottom.objrotz + 1
  End Select
End Sub

' Upper Right Slot Kicker Out
Sub slot_kicker(enabled)
  If enabled Then
    Ukick = 1
    upper_kicker.enabled = 1
    bsle.exitSol_On
    SoundSaucerKick 1, sw65
  End If
End Sub

Dim ukick

Sub Upper_Kicker_Timer()
  Select Case ukick
    Case 1:
      If Hammer.RotZ => 45 Then
        ukick = 2
      End If
      Hammer.RotZ = Hammer.RotZ + 1
    Case 2:
      If Hammer.RotZ <= 0 Then
        Hammer.RotZ = 0
        me.enabled = 0
      End If
      Hammer.RotZ = Hammer.RotZ - 1
  End Select
End Sub

'***************************
' RealTime Updates / Timers
'***************************

Set MotorCallback = GetRef("GameTimer")

Sub GameTimer
  batleft.objrotz = LeftFlipper.CurrentAngle + 1
  batright.objrotz = RightFlipper.CurrentAngle + 1
  batright1.objrotz = RightFlipper2.CurrentAngle + 1
  batleftshadow.objrotz = batleft.objrotz
  batrightshadow.objrotz  = batright.objrotz
  batrightshadow1.objrotz  = RightFlipper2.CurrentAngle + 1
End Sub

Sub modtoy_timer()
  Track1.roty = Track1.roty + 1
  'Track2.roty = Track1.roty + 1
  Track3.roty = Track1.roty + 1
  modref.roty = modref.roty + 1
  'modref1.roty = modref1.roty + 1
  Indy_shaft1.objrotz = Indy_shaft1.objrotz - 1
  TopCar_Base2.objrotz = TopCar_Base2.objrotz - 1
End Sub

Sub FrameTimer_Timer()
  If DynamicBallShadowsOn Or AmbientBallShadowOn Then DynamicBSUpdate ' Update ball shadows
End Sub

'**************
' Flipper Subs
'**************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"
SolCallback(sURFlipper) = "SolURFlipper" ' Uncomment for Staged Flippers

'******************************************
' Use FlipperTimers to call div subs
'******************************************
Const ReflipAngle = 20

Sub SolLFlipper(Enabled) ' Left flipper solenoid callback
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

Sub SolRFlipper(Enabled) ' Right flipper solenoid callback
  If Enabled Then
    RF.Fire
    'RightFlipper2.RotateToEnd ' Comment this line for Staged Flippers

    If rightflipper.currentangle > rightflipper.endangle - ReflipAngle Then
      RandomSoundReflipUpRight RightFlipper
    Else
      SoundFlipperUpAttackRight RightFlipper
      RandomSoundFlipperUpRight RightFlipper
    End If
  Else
    RightFlipper.RotateToStart
    'RightFlipper2.RotateToStart ' Comment this line for Staged Flippers
    If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
      RandomSoundFlipperDownRight RightFlipper
    End If
    FlipperRightHitParm = FlipperUpSoundLevel
  End If
End Sub

' Uncomment for Staged Flippers
Sub SolURFlipper(Enabled)
     If Enabled Then
         PlaySoundatvol SoundFX("right_flipper_up", DOFFlippers),rightflipper2, 0.5
         RightFlipper2.RotateToEnd
     Else
         PlaySoundatvol SoundFX("right_flipper_down", DOFFlippers),rightflipper2, 0.1
     RightFlipper2.RotateToStart
     End If
End Sub

' Lanes
Sub sw15_Hit:Controller.Switch(15) = 1:End Sub
Sub sw15_UnHit:Controller.Switch(15) = 0:End Sub

Sub sw16_Hit:Controller.Switch(16) = 1:End Sub
Sub sw16_UnHit:Controller.Switch(16) = 0:End Sub

Sub sw17_Hit:Controller.Switch(17) = 1:End Sub
Sub sw17_UnHit:Controller.Switch(17) = 0:End Sub

Sub sw18_Hit:Controller.Switch(18) = 1:End Sub
Sub sw18_UnHit:Controller.Switch(18) = 0:End Sub

Sub sw37_Hit:Controller.Switch(37) = 1:End Sub
Sub sw37_UnHit:Controller.Switch(37) = 0:End Sub

Sub sw38_Hit:Controller.Switch(38) = 1:End Sub
Sub sw38_UnHit:Controller.Switch(38) = 0:End Sub

Sub sw51_Hit:Controller.Switch(51) = 1:End Sub
Sub sw51_UnHit:Controller.Switch(51) = 0:End Sub

Sub sw52_Hit:Controller.Switch(52) = 1:End Sub
Sub sw52_UnHit:Controller.Switch(52) = 0:End Sub

Sub sw53_Hit:Controller.Switch(53) = 1:End Sub
Sub sw53_UnHit:Controller.Switch(53) = 0:End Sub

' Other switches
Sub sw25_Hit:Controller.Switch(25) = 1:BallInShooterLane = 1:PlaySoundatball "sensor":End Sub
Sub sw25_Unhit:Controller.Switch(25) = 0:BallInShooterLane = 0:End Sub
Sub sw35_Hit:Controller.Switch(35) = 1:PlaySoundat "gateback_low",sw35:End Sub
Sub sw35_Unhit:Controller.Switch(35) = 0:End Sub
Sub sw36_Hit:Controller.Switch(36) = 1:PlaySoundat "gateback_low",sw36:End Sub
Sub sw36_Unhit:Controller.Switch(36) = 0:End Sub
Sub sw75_Hit:Controller.Switch(75) = 1:PlaySoundat "gateback_low",sw75:End Sub
Sub sw75_Unhit:Controller.Switch(75) = 0:End Sub
Sub sw76_Hit:Controller.Switch(76) = 1:PlaySoundat "gateback_low",sw76:End Sub
Sub sw76_Unhit:Controller.Switch(76) = 0:End Sub


' Targets
Sub sw28_Hit:vpmTimer.PulseSw 28:End Sub
Sub sw31_Hit:vpmTimer.PulseSw 31:End Sub
Sub sw32_Hit:vpmTimer.PulseSw 32:End Sub
Sub sw34_Hit:vpmTimer.PulseSw 34:End Sub
Sub sw46_Hit:vpmTimer.PulseSw 46:End Sub
Sub sw47_Hit:vpmTimer.PulseSw 47:End Sub
Sub sw48_Hit:vpmTimer.PulseSw 48:End Sub
Sub sw55_Hit:vpmTimer.PulseSw 55:End Sub
Sub sw56_Hit:vpmTimer.PulseSw 56:End Sub
Sub sw57_Hit:vpmTimer.PulseSw 57:End Sub
Sub sw58_Hit:vpmTimer.PulseSw 58:End Sub

' Vuks
Dim popperTBall, popperTZpos
Dim popperBBall, popperBZpos

Sub sw61_Hit:PlaySoundat "kicker_enter",sw61:sw61a.Enabled=1:bsTBP.AddBall Me:End Sub
Sub sw62_Hit:PlaySoundatvol "warehousehit",sw62, 0.2:sw62a.Enabled=1:bsBBP.AddBall Me:End Sub

Sub SolTopPopper(Enabled)
  If Enabled Then
    If bsTBP.Balls Then
      sw61.Enabled = 0
      Set popperTBall = sw61a.Createball
      popperTBall.Z = 0
      popperTZpos = 0
      SoundSaucerKick 1, sw61a
      sw61a.TimerInterval = 2
      sw61a.TimerEnabled = 1
    End If
  End If
End Sub

Sub sw61a_Timer
  popperTBall.Z = popperTZpos
  popperTZpos = popperTZpos + 10
  If popperTZpos> 170 Then
    sw61a.TimerEnabled = 0
    sw61a.DestroyBall
    bsTBP.ExitSol_On
  End If
End Sub

Sub SolBottomPopper(Enabled)
  If Enabled Then
    If bsBBP.Balls Then
      sw62.Enabled = 0
      Set popperBBall = sw62a.Createball
      popperBBall.Z = 0
      popperBZpos = 0
      SoundSaucerKick 1, sw62a
      sw62a.TimerInterval = 2
      sw62a.TimerEnabled = 1
    End If
  End If
End Sub

Sub sw62a_Timer
  popperBBall.Z = popperBZpos
  popperBZpos = popperBZpos + 10
  If popperBZpos> 110 Then
    sw62a.TimerEnabled = 0
    sw62a.DestroyBall
    bsBBP.ExitSol_On
    If Sloppy_Turbo_VUK = 1 Then
      sw62.Enabled = 0 '' TEST
    Else
      sw62.Enabled = 1
    End If
  End If
End Sub

Dim tballcount
tballcount = 0

Sub sw62b_Hit()
  sw62b.Destroyball
  tballcount = tballcount + 1
  StopSound "rail"
  PlaySoundat "fx_metalclank",Turbo_Bottom
  vpmTimer.AddTimer 250, "AddBallToTurbo"
  If tballcount = 1 Then
    ball1.visible = 1
    TREF.image = "disc2db_1"
  End If
  If tballcount = 2 Then
    ball2.visible = 1
    TREF.image = "disc2db_2"
  End If
  If tballcount = 3 Then
    ball3.visible = 1
    TREF.image = "disc2db_3"
  End If
  If tballcount = 4 Then
    ball4.visible = 1
    TREF.image = "disc2db_4"
  End If
End Sub

' Saucers
Sub sw64_Hit
  SoundSaucerLock
  bsUE.AddBall Me
End Sub

Sub sw65_Hit
  SoundSaucerLock
  bsLE.AddBall Me
End Sub

Sub Drain1_Hit
  RandomSoundDrain Drain1
  bsTrough.AddBall Me
End Sub

' Bumpers
Sub Bumper1_Hit
  vpmTimer.PulseSw 72
  RandomSoundBumperTop Bumper1
End Sub

Sub Bumper2_Hit
  vpmTimer.PulseSw 73
  RandomSoundBumperMiddle Bumper2
End Sub

Sub Bumper3_Hit
  vpmTimer.PulseSw 74
  RandomSoundBumperBottom Bumper3
End Sub

' Ramp Rolling Sounds
Sub LeftRampStart_hit
  WireRampOn True
  bsRampOnClear
End Sub

Sub RightRampStart_hit
  WireRampOn True
  bsRampOnClear
End Sub

Sub TOUT_Hit ' The wire ramp trigger from turbo spinner.
  WireRampOn False
  bsRampOnWire
End Sub

Sub TriggerSpeedwayKicker_Hit
  WireRampOn True
End Sub

'************************************************************************
'         Slingshots Animation
'************************************************************************

Dim LStep, RStep

Sub LeftSlingShot_Slingshot
  LS.VelocityCorrect(ActiveBall)
    RandomSoundSlingshotLeft Sling1
    vpmTimer.PulseSw 26
  LSling.Visible = 0
  LSling1.Visible = 1
  sling1.TransZ = -20
  LStep = 0
  LeftSlingShot.TimerEnabled = 1
    LeftSlingShot.TimerInterval = 10
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
    RandomSoundSlingshotRight Sling2
  vpmTimer.PulseSw 27
  RSling.Visible = 0
  RSling1.Visible = 1
  sling2.TransZ = -20
  RStep = 0
  RightSlingShot.TimerEnabled = 1
    RightSlingShot.TimerInterval = 10
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling2.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

 Sub sw54_Slingshot()
  vpmTimer.PulseSw 54
  playsoundat SoundFX("Sling_R1",DOFContactors),BR3
 End Sub

'***********
' Update GI
'***********
Set GiCallback2 = GetRef("UpdateGI")

Dim xx
Dim gistep
gistep = 1/8

Sub UpdateGI(no, step)

  If step < 5 Then
    F27_LowBoy.State = 1
    F28_LowBoy.State = 1
  Else
    F27_LowBoy.State = 0
    F28_LowBoy.State = 0
  End If

  If step => 0 Then
    For each xx in GITL:xx.state = 1:next
    For each xx in GITR:xx.state = 1:next
    For each xx in GIB:xx.state = 1:next
  End If

  For each xx in GIF:xx.opacity = step * 30:next

  If step = 0 then
  Table1.ColorGradeImage = "LUT0"
  else
  if step = 1 Then
  Table1.ColorGradeImage = "LUT0_1"
  else
  if step = 2 Then
  Table1.ColorGradeImage = "LUT0_2"
  else
  if step = 3 Then
  Table1.ColorGradeImage = "LUT0_3"
  else
  if step = 4 Then
  Table1.ColorGradeImage = "LUT0_4"
  else
  if step = 5 Then
  Table1.ColorGradeImage = "LUT0_5"
  else
  if step = 6 Then
  Table1.ColorGradeImage = "LUT0_6"
  else
  if step = 7 Then
  Table1.ColorGradeImage = "LUT0_7"
  else
  if step = 8 Then
  Table1.ColorGradeImage = "LUT0_8"
  End If
  End If
  End If
  End If
  End If
  End If
  End If
  End If
  End If

  If step = 8 Then
    DOF 302, DOFOn
  Else
    DOF 302, DOFOff
  End If

  If step => 1 Then
    PinCab_Backglass.blenddisablelighting = 5
  Else
    PinCab_Backglass.blenddisablelighting = 0.2
  End If

  Select Case no

    Case 0 'top left
      For each xx in GITL:xx.IntensityScale = gistep * step:next
      ref5.opacity = step * 2
      ref7.opacity = step * 2

      If Dozer_Cab = 1 Then
        If step = 1 then Controller.B2SSetData 201,0:Controller.B2SSetData 202,0:Controller.B2SSetData 203,0:Controller.B2SSetData 204,0:Controller.B2SSetData 205,0:Controller.B2SSetData 206,0
        If step = 2 then Controller.B2SSetData 201,1:Controller.B2SSetData 202,0:Controller.B2SSetData 203,0:Controller.B2SSetData 204,0:Controller.B2SSetData 205,0:Controller.B2SSetData 206,0
        If step = 3 then Controller.B2SSetData 202,1:Controller.B2SSetData 201,0:Controller.B2SSetData 203,0:Controller.B2SSetData 204,0:Controller.B2SSetData 205,0:Controller.B2SSetData 206,0
        If step = 4 then Controller.B2SSetData 203,1:Controller.B2SSetData 202,0:Controller.B2SSetData 201,0:Controller.B2SSetData 204,0:Controller.B2SSetData 205,0:Controller.B2SSetData 206,0
        If step = 5 then Controller.B2SSetData 204,1:Controller.B2SSetData 202,0:Controller.B2SSetData 203,0:Controller.B2SSetData 201,0:Controller.B2SSetData 205,0:Controller.B2SSetData 206,0
        If step = 6 then Controller.B2SSetData 205,1:Controller.B2SSetData 202,0:Controller.B2SSetData 203,0:Controller.B2SSetData 204,0:Controller.B2SSetData 201,0:Controller.B2SSetData 206,0
        If step = 8 then Controller.B2SSetData 206,1:Controller.B2SSetData 202,0:Controller.B2SSetData 203,0:Controller.B2SSetData 204,0:Controller.B2SSetData 205,0:Controller.B2SSetData 201,0
      End If

    Case 1 'top right
      For each xx in GITR:xx.IntensityScale = gistep * step:next
      ref6.opacity = step * 2

      If Dozer_Cab = 1 Then
        If step = 1 then Controller.B2SSetData 101,0:Controller.B2SSetData 102,0:Controller.B2SSetData 103,0:Controller.B2SSetData 104,0:Controller.B2SSetData 105,0:Controller.B2SSetData 106,0
        If step = 2 then Controller.B2SSetData 101,1:Controller.B2SSetData 102,0:Controller.B2SSetData 103,0:Controller.B2SSetData 104,0:Controller.B2SSetData 105,0:Controller.B2SSetData 106,0
        If step = 3 then Controller.B2SSetData 102,1:Controller.B2SSetData 101,0:Controller.B2SSetData 103,0:Controller.B2SSetData 104,0:Controller.B2SSetData 105,0:Controller.B2SSetData 106,0
        If step = 4 then Controller.B2SSetData 103,1:Controller.B2SSetData 102,0:Controller.B2SSetData 101,0:Controller.B2SSetData 104,0:Controller.B2SSetData 105,0:Controller.B2SSetData 106,0
        If step = 5 then Controller.B2SSetData 104,1:Controller.B2SSetData 102,0:Controller.B2SSetData 103,0:Controller.B2SSetData 101,0:Controller.B2SSetData 105,0:Controller.B2SSetData 106,0
        If step = 6 then Controller.B2SSetData 105,1:Controller.B2SSetData 102,0:Controller.B2SSetData 103,0:Controller.B2SSetData 104,0:Controller.B2SSetData 101,0:Controller.B2SSetData 106,0
        If step = 8 then Controller.B2SSetData 106,1:Controller.B2SSetData 102,0:Controller.B2SSetData 103,0:Controller.B2SSetData 104,0:Controller.B2SSetData 105,0:Controller.B2SSetData 101,0
      End If

    Case 2 'bottom
      For each xx in GIB:xx.IntensityScale = gistep * step:next
      ref1.opacity = step * 2
      ref2.opacity = step * 2
      ref3.opacity = step * 2
      ref4.opacity = step * 2

      If Dozer_Cab = 1 Then
        If step = 1 then Controller.B2SSetData 111,0:Controller.B2SSetData 112,0:Controller.B2SSetData 113,0:Controller.B2SSetData 114,0:Controller.B2SSetData 115,0:Controller.B2SSetData 116,0
        If step = 2 then Controller.B2SSetData 111,1:Controller.B2SSetData 112,0:Controller.B2SSetData 113,0:Controller.B2SSetData 114,0:Controller.B2SSetData 115,0:Controller.B2SSetData 116,0
        If step = 3 then Controller.B2SSetData 112,1:Controller.B2SSetData 111,0:Controller.B2SSetData 113,0:Controller.B2SSetData 114,0:Controller.B2SSetData 115,0:Controller.B2SSetData 116,0
        If step = 4 then Controller.B2SSetData 113,1:Controller.B2SSetData 112,0:Controller.B2SSetData 111,0:Controller.B2SSetData 114,0:Controller.B2SSetData 115,0:Controller.B2SSetData 116,0
        If step = 5 then Controller.B2SSetData 114,1:Controller.B2SSetData 112,0:Controller.B2SSetData 113,0:Controller.B2SSetData 111,0:Controller.B2SSetData 115,0:Controller.B2SSetData 116,0
        If step = 6 then Controller.B2SSetData 115,1:Controller.B2SSetData 112,0:Controller.B2SSetData 113,0:Controller.B2SSetData 114,0:Controller.B2SSetData 111,0:Controller.B2SSetData 116,0
        If step = 8 then Controller.B2SSetData 116,1:Controller.B2SSetData 112,0:Controller.B2SSetData 113,0:Controller.B2SSetData 114,0:Controller.B2SSetData 115,0:Controller.B2SSetData 111,0
      End If

   End Select
End Sub


'**********
 ' TopCar
 '**********

 Sub SolTopCar(Enabled)
     If Enabled Then
         PlaySoundat SoundFX("motor",DOFGear),Indy_Top
         Indy_Upper.Enabled = 1
         Indy_up_Flash.Enabled = 1
     Else
         Indy_Upper.Enabled = 0
         Indy_up_Flash.Enabled = 0
     XBallShadow1.visible = 0
     IU1.State = 0
     IU2.State = 0
     IU3.State = 0
     IU4.State = 0
     IU5.State = 0
     IUFLASH1.visible = 0
         IUFLASH2.visible = 0
     Car_Shadow.visible = 1
         F17s4.visible = 0
     End If
 End Sub

Sub Indy_Upper_Timer()
  indy_top.objrotz = indy_top.objrotz - 1
  indy_shaft.objrotz = indy_shaft.objrotz - 1
  topcar_base.objrotz = topcar_base.objrotz - 1
  topcar_base1.objrotz = topcar_base1.objrotz - 1
  If Rotating_Track = 1 Then
    track.objrotz = track.objrotz + 1
  End If
  car_shadow.objrotz = car_shadow.objrotz - 1
  XBallShadow1.objrotz = XBallShadow1.objrotz - 1
End Sub

Dim iupflash

iupflash = 1

Sub Indy_Up_Flash_Timer
  Select Case iupflash
    Case 1:
      XBallShadow1.visible = 0
      IU1.State = 0
      IU2.State = 0
      IU3.State = 0
      IU4.State = 0
      IU5.State = 0
      IUFLASH1.visible = 0
      IUFLASH2.visible = 0
      Car_Shadow.visible = 0
      F17s4.visible = 0
      iupflash = 2
    Case 2
      XBallShadow1.visible = 1
      IU1.State = 1
      IU2.State = 1
      IU3.State = 1
      IU4.State = 1
      IU5.State = 1
      IUFLASH1.visible = 1
      IUFLASH2.visible = 1
      F17s4.visible = 1
      Car_Shadow.visible = 1
      iupflash = 1
  End Select
End Sub


'*********************************************************************************************
' Turbo Handler Code by Dorsola - Code Additions for Primitive Objects Turbo (Dozer).
' Theory of operation:
' - TurboPopper    shoots a ball into the turbo mechanism
' - TurboIndex     gets triggered every quarter turn of the turbo mech
' - TurboBallSense detects the quadrant slats every quarter turn and a locked ball
'
' * If storing balls, turbo turns at low speed and and keeps balls in turbo
' * If releasing balls, turbo spins up until centrIfugal force ejects each ball onto the ramp.
'
' Known Issues:
' * Turbo lock behavior is not always consistent, but is believed to be accurately simulated
'**********************************************************************************************

Dim BallInTurboIndex(4), TurboSpeed, TurboMotorState, TurboRotatePosition, TurboCounter, TurboTargetCounter

Sub CloseBallSense:Controller.Switch(63) = 1:End Sub

Sub OpenBallSense:Controller.Switch(63) = 0:End Sub

Sub CloseTurboIndex:Controller.Switch(66) = 1:End Sub
Sub OpenTurboIndex:Controller.Switch(66) = 0:End Sub

Sub ShowTurboBalls
  ShowTurboPos TurboPos1A, BallInTurboIndex(1) And TurboRotatePosition = 0
  ShowTurboPos TurboPos1B, BallInTurboIndex(1) And TurboRotatePosition = 1
  ShowTurboPos TurboPos1C, BallInTurboIndex(1) And TurboRotatePosition = 2
  ShowTurboPos TurboPos1D, BallInTurboIndex(1) And TurboRotatePosition = 3
  ShowTurboPos TurboPos2A, BallInTurboIndex(2) And TurboRotatePosition = 0
  ShowTurboPos TurboPos2B, BallInTurboIndex(2) And TurboRotatePosition = 1
  ShowTurboPos TurboPos2C, BallInTurboIndex(2) And TurboRotatePosition = 2
  ShowTurboPos TurboPos2D, BallInTurboIndex(2) And TurboRotatePosition = 3
  ShowTurboPos TurboPos3A, BallInTurboIndex(3) And TurboRotatePosition = 0
  ShowTurboPos TurboPos3B, BallInTurboIndex(3) And TurboRotatePosition = 1
  ShowTurboPos TurboPos3C, BallInTurboIndex(3) And TurboRotatePosition = 2
  ShowTurboPos TurboPos3D, BallInTurboIndex(3) And TurboRotatePosition = 3
  ShowTurboPos TurboPos4A, BallInTurboIndex(4) And TurboRotatePosition = 0
  ShowTurboPos TurboPos4B, BallInTurboIndex(4) And TurboRotatePosition = 1
  ShowTurboPos TurboPos4C, BallInTurboIndex(4) And TurboRotatePosition = 2
  ShowTurboPos TurboPos4D, BallInTurboIndex(4) And TurboRotatePosition = 3
End Sub

Sub ShowTurboWalls
  Select Case TurboRotatePosition
    Case 0:TurboWallA.isdropped = 0:TurboWallB.isdropped = 1:TurboWallC.isdropped = 1:TurboWallD.isdropped = 1
    Case 1:TurboWallA.isdropped = 1:TurboWallB.isdropped = 0:TurboWallC.isdropped = 1:TurboWallD.isdropped = 1
    Case 2:TurboWallA.isdropped = 1:TurboWallB.isdropped = 1:TurboWallC.isdropped = 0:TurboWallD.isdropped = 1
    Case 3:TurboWallA.isdropped = 1:TurboWallB.isdropped = 1:TurboWallC.isdropped = 1:TurboWallD.isdropped = 0
  End Select
End Sub

'Dozer Code Mod
Sub ShowTurboPos(kicker, show)
  If show Then
    If kicker.Enabled = 0 Then
      'kicker.CreateBall
      kicker.Enabled = 1
    End If
  Else
         If kicker.Enabled Then
      'kicker.DestroyBall
      kicker.Enabled = 0
    End If
  End If
End Sub

'Dozer Code Mod
Sub Init_Turbo()
  Dim count
  for count = 1 to 4:BallInTurboIndex(count) = 0:next
  TurboSpeed = 0
  TurboCounter = 0
  TurboRotatePosition = 0
  TurboMotorState = Controller.GetMech(0)
  OpenTurboIndex
  OpenBallSense
  ShowTurboBalls
  ShowTurboWalls
End Sub

Sub AdvanceTurboBalls()
  Dim overflow
  overflow = BallInTurboIndex(4)
  Dim count
  For count = 3 to 1 step - 1
    BallInTurboIndex(count + 1) = BallInTurboIndex(count)
  Next

  BallInTurboIndex(1) = overflow
End Sub

Sub RotateTurbo()
  TurboRotatePosition = (TurboRotatePosition + 1) mod 4
  Select Case TurboRotatePosition
    Case 0
      AdvanceTurboBalls
      OpenTurboIndex
      If BallInTurboIndex(2) Then
        CloseBallSense
      Else
        OpenBallSense
      End If

      If TurboSpeed> 0 and TurboMotorState = 0 Then TurboSpeed = 0

    Case 1
      CloseTurboIndex
      CloseBallSense

    Case 2
      CloseTurboIndex
      CloseBallSense
      If TurboSpeed = 2 And BallInTurboIndex(2) Then
        TurboPopperEject.Enabled = 1
        TurboPopperEject.CreateBall
        TurboPopperEject.Kick 80, 26
        tballcount = tballcount - 1
        PlaySoundat "metalhit",TurboPopperEject
        ball1.visible = 0
        ball2.visible = 0
        ball3.visible = 0
        ball4.visible = 0
        TREF.image = "disc2db"
        BallInTurboIndex(2) = 0
      End If
    Case 3
      CloseTurboIndex
      CloseBallSense
  End Select

  ShowTurboBalls
  ShowTurboWalls
End Sub

Sub AddBallToTurbo(no)
  Dim index
  index = 1

  If BallInTurboIndex(1) Then
    While BallInTurboIndex(index) = 1 and index <= 4
      index = index + 1
    WEnd
  End If

  If index> 4 Then
    sw62b.CreateBall
    sw62b_Hit
  Else
    BallInTurboIndex(index) = 1
    If index = 2 Then CloseBallSense
  End If

  ShowTurboBalls
  ShowTurboWalls
End Sub

Sub TurboMasterTimer_Timer()
  Dim Temp
  Temp = Controller.GetMech(0)
  If Temp <> TurboMotorState Then
    TurboMotorState = Temp
    If Temp> 0 Then TurboSpeed = Temp
      'Dozer Code
      If Temp = 1 Then
        Turbo_Slow.enabled = 1
        PlaySoundatvol SoundFX("TurboMotor_Low",DOFGear),Turbo_Bottom, 0.05
        'PlaySound "TurboMotor_Low", 1, 0.2, AudioPan(Turbo_Bottom), 0,0,0, 1, AudioFade(Turbo_Bottom)
        DOF 303, DOFOn
        Turbo_Fast.enabled = 0:StopSound "TurboMotor":DOF 301, DOFOff
        Turbo_Flash_fast.enabled = 0
        Turbo_Flash_slow.enabled = 1
      else
      If Temp = 2 Then
        Turbo_Slow.enabled = 0:StopSound "TurboMotor_Low":DOF 303, DOFOff
        Turbo_Fast.enabled = 1
        PlaySoundatvol SoundFX("TurboMotor",DOFGear),Turbo_Bottom, 0.05
        DOF 301, DOFOn
        Turbo_Flash_fast.enabled = 1
        Turbo_Flash_slow.enabled = 0
      else
        Turbo_Fast.enabled = 0
        Turbo_Slow.enabled = 0
        trepos=1:Turbo_Repos.enabled = 1
      End If
    End If
    'Dozer Code
  End If

  Select Case TurboSpeed
    Case 0
      TurboCounter = 0
    Case 1
      TurboCounter = TurboCounter + 1
      If TurboCounter> 2 Then
        RotateTurbo
        TurboCounter = 0
      End If
    Case 2
      RotateTurbo
      TurboCounter = 0
  End Select
End Sub


' Code to handle Primitive Turbo and Ball Handling (Dozer).
Dim tfpos,trepos
tfpos = 60

Sub Turbo_Fast_timer()
  If tfpos = 360 Then
    tfpos = 0
  End If
  impellor.objrotz = impellor.objrotz + 1
  tref.rotz = tref.rotz - 1
  modref.rotz = modref.rotz - 1
  turbo_bottom.objrotz = turbo_bottom.objrotz + 1
  ball1.objrotz = ball1.objrotz + 1
  ball2.objrotz = ball2.objrotz + 1
  ball3.objrotz = ball3.objrotz + 1
  ball4.objrotz = ball4.objrotz + 1
  tfpos = tfpos + 1
End Sub

Dim tflashf
tflashf = 1

Sub Turbo_Flash_Fast_Timer
  Select Case tflashf
    Case 1:
      Fturbo.visible = 0:Turbo_Side_Flash2.visible = 0:tflashf = 2
    Case 2
      Fturbo.visible = 1:Turbo_Side_Flash2.visible = 1:tflashf = 1
  End Select
End Sub

Dim tflashs
tflashs = 1
Sub Turbo_Flash_Slow_Timer
  Select Case tflashs
    Case 1:
      Fturbo.visible = 0:Turbo_Side_Flash2.visible = 0:tflashs = 2
    Case 2
      Fturbo.visible = 1:Turbo_Side_Flash2.visible = 1:tflashs = 1
  End Select
End Sub

Sub Turbo_Slow_timer()
  If tfpos = 360 Then
    tfpos = 0
  End If
  impellor.objrotz = impellor.objrotz + 1
  tref.rotz = tref.rotz - 1
  modref.rotz = modref.rotz - 1
  turbo_bottom.objrotz = turbo_bottom.objrotz + 1
  ball1.objrotz = ball1.objrotz + 1
  ball2.objrotz = ball2.objrotz + 1
  ball3.objrotz = ball3.objrotz + 1
  ball4.objrotz = ball4.objrotz + 1
  tfpos = tfpos + 1
End Sub

Sub Turbo_Repos_Timer()
  Select Case trepos
    Case 1:
      If NOT tfpos = 60 Then
        trepos = 2
      End If
    Case 2:
      If tfpos => 60 Then
        impellor.objrotz = 60
        tfpos = 60
        Fturbo.visible = 1
        turbo_flash_slow.enabled = 0
        turbo_flash_fast.enabled = 0
        StopSound "TurboMotor_Low":DOF 303, DOFOff
        StopSound "TurboMotor":DOF 301, DOFOff
        vpmTimer.AddTimer 200, "RestopSound"
        me.enabled = 0
      End If
      If tfpos = 360 Then
        tfpos = 0
      End If
      impellor.objrotz = impellor.objrotz + 1
      tref.rotz = tref.rotz - 1
      modref.rotz = modref.rotz - 1
      turbo_bottom.objrotz = turbo_bottom.objrotz + 1
      ball1.objrotz = ball1.objrotz + 1
      ball2.objrotz = ball2.objrotz + 1
      ball3.objrotz = ball3.objrotz + 1
      ball4.objrotz = ball4.objrotz + 1
      tfpos = tfpos + 1
  End Select
End Sub

Sub RestopSound(no)
  StopSound "TurboMotor_Low":DOF 303, DOFOff
  StopSound "TurboMotor":DOF 301, DOFOff
End Sub

'Redundant but kept until release.
Sub turbo_watch_timer()
  If tballcount = 1 Then
    ball1.visible = 1
  End If
  If tballcount = 2 Then
    ball2.visible = 1
  End If
  If tballcount = 3 Then
    ball3.visible = 1
  End If
  If tballcount = 4 Then
    ball4.visible = 1
  End If
End Sub

'Insert Lights
Set Lights(11)=L11
Set Lights(12)=L12
Set Lights(13)=L13
Set Lights(15)=L15
Set Lights(18)=L18
'Set Lights(14)=L14
'Set Lights(16)=L16
'Set Lights(17)=L17
Lights(14)=Array(L14,L14a)
Lights(16)=Array(L16,L16a)
Lights(17)=Array(L17,L17a)
'Set Lights(21)=L21
'Set Lights(22)=L22
'Set Lights(23)=L23
Lights(21)=Array(L21,L21a)
Lights(22)=Array(L22,L22a)
Lights(23)=Array(L23,L23a)
Set Lights(24)=L24
Set Lights(25)=L25
Set Lights(26)=L26
Set Lights(27)=L27
Set Lights(28)=L28
Set Lights(31)=L31
Set Lights(32)=L32
Set Lights(33)=L33
Set Lights(34)=L34
Set Lights(35)=L35
Set Lights(36)=L36
Set Lights(37)=L37
Set Lights(38)=L38
Set Lights(41)=L41
Set Lights(42)=L42
Set Lights(43)=L43
Set Lights(44)=L44
Set Lights(45)=L45
Set Lights(46)=L46
Set Lights(47)=L47
Set Lights(48)=L48
Set Lights(57)=L57
Set Lights(58)=L58
Set Lights(61)=L61
Set Lights(62)=L62
Set Lights(63)=L63
Set Lights(51)=L51
Set Lights(52)=L52
Set Lights(53)=L53
Set Lights(54)=L54
Set Lights(55)=L55
Set Lights(56)=L56
Set Lights(64)=L64
Set Lights(65)=L65
Set Lights(66)=L66
Set Lights(75)=L75x
Set Lights(76)=L76x
Set Lights(77)=L77x
Set Lights(78)=L78x
Lights(71)=Array(L71x,L71)
Lights(72)=Array(L72x,L72)
Lights(73)=Array(L73x,L73)
Lights(74)=Array(L74x,L74)
Lights(75)=Array(L75x,L75x1)
Lights(76)=Array(L76x,L76x1)
Lights(77)=Array(L77x,L77x1)
Lights(78)=Array(L78x,L78x1)
Lights(81)=Array(L81x,L81x1)
Lights(82)=Array(L82x,L82x1)
Lights(83)=Array(L83x,L83x1)
Lights(84)=Array(L84x,L84x1)


'----------------------------
' Misc Hitty Stuff
'----------------------------

Sub TurboErrorCatcher_Hit()
  Me.DestroyBall
  TurboPopperEject.Enabled = 1
  TurboPopperEject.CreateBall
  TurboPopperEject.Kick 80, 26
  tballcount = tballcount - 1
End Sub

Sub TurboPopperEject_Unhit() : TurboPopperEject.Enabled = 0 : End Sub
Sub sw61a_UnHit() : PlaySoundat "balldrop",sw61a : sw61a.Enabled=0 : End Sub
Sub sw62a_UnHit() : sw62a.Enabled=0 : End Sub
Sub sw62b_UnHit() : End Sub

Sub RHD_Hit()
  SoundSaucerLock
End Sub

Sub RHD2_Hit()
  PlaySoundatvol "balldrop",RHD2, 0.5
End Sub

Sub LHD_Hit()
  PlaySoundatvol "balldrop",sw16, 0.5
  'StopSound "subway2"
End Sub

Sub MetalStop_Hit()
  PlaySoundat "fx_metalclank",Primitive61
  'StopSound "rail_low_slower"
End Sub

Sub Wall720_Hit()
  PlaySoundat "metalhit",sw65
End Sub

Sub Wall14_Hit()
  PlaySoundat "metalhit",whramp1
End Sub

Sub Wall205_Hit()
  PlaySoundat "metalhit",sw47
End Sub

Sub Wall27_Hit()
  PlaySoundat "metalhit",disc2
End Sub

Sub DiverterHold_Hit()
  PlaySoundat "plastichit",Diverterhold
End Sub

Sub gate2_hit()
  PlaySoundat "gateback_low",gate2
End Sub

Sub gate3_hit()
  PlaySoundat "gateback_low",gate3
End Sub

Sub Tback_Hit()
  If Sloppy_Turbo_VUK = 1 Then
    PlaySoundat "metalhit",whramp1
    mSpinMagnet.MagnetOn = True
    tppos1 = 1:tpopenable.enabled = 1
  End If
End Sub

Dim tppos1

Sub tpopenable_timer()
  Select Case tppos1
    Case 1:
      mSpinMagnet.MagnetOn = False
      sw62.enabled = 1
      tppos1 = 2
    Case 2:
      me.enabled = 0
  End Select
End Sub

'*****************************************
'      Other Table Sounds
'*****************************************
Sub LeftFlipper_Collide(parm)
  CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
  LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
  CheckLiveCatch Activeball, RightFlipper, RFCount, parm
  RightFlipperCollide parm
End Sub

Sub RightFlipper2_Collide(parm)
  CheckLiveCatch Activeball, RightFlipper2, RFCount, parm
  RightFlipperCollide parm
End Sub

Sub Dampen(dt,df,r)           'dt is threshold speed, df is dampen factor 0 to 1 (higher more dampening), r is randomness
  Dim dfRandomness
  r=cint(r)
  dfRandomness=INT(RND*(2*r+1))
  df=df+(r-dfRandomness)*.01
  If ABS(activeball.velx) > dt Then activeball.velx=activeball.velx*(1-df*(ABS(activeball.velx)/100))
  If ABS(activeball.vely) > dt Then activeball.vely=activeball.vely*(1-df*(ABS(activeball.vely)/100))
End Sub

Sub Table1_Exit
  Controller.Stop
End Sub

'******************************************************
' ZNFF:  FLIPPER CORRECTIONS by nFozzy
'******************************************************

'******************************************************
' Flippers Polarity (Select appropriate sub based on era)
'******************************************************

Dim LF, RF
Set LF = New FlipperPolarity
Set RF = New FlipperPolarity

InitPolarity

'*******************************************
' Early 90's and after

Sub InitPolarity()
  dim x, a : a = Array(LF, RF)
  for each x in a
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

'' Flipper trigger hit subs
'Sub TriggerLF_Hit()
' LF.Addball activeball
'End Sub
'Sub TriggerLF_UnHit()
' LF.PolarityCorrect activeball
'End Sub
'Sub TriggerRF_Hit()
' RF.Addball activeball
'End Sub
'Sub TriggerRF_UnHit()
' RF.PolarityCorrect activeball
'End Sub

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
  private Flipper, FlipperStart, FlipperEnd, FlipperEndY, LR, PartialFlipCoef
  Private Balls(20), balldata(20)
  private Name

  dim PolarityIn, PolarityOut
  dim VelocityIn, VelocityOut
  dim YcoefIn, YcoefOut
  Public Sub Class_Initialize
    redim PolarityIn(0) : redim PolarityOut(0) : redim VelocityIn(0) : redim VelocityOut(0) : redim YcoefIn(0) : redim YcoefOut(0)
    Enabled = True : TimeDelay = 50 : LR = 1:  dim x : for x = 0 to uBound(balls) : balls(x) = Empty : set Balldata(x) = new SpoofBall : next
  End Sub

  Public Sub SetObjects(aName, aFlipper, aTrigger)

    if typename(aName) <> "String" then msgbox "FlipperPolarity: .SetObjects error: first argument must be a string (and name of Object). Found:" & typename(aName) end if
    if typename(aFlipper) <> "Flipper" then msgbox "FlipperPolarity: .SetObjects error: second argument must be a flipper. Found:" & typename(aFlipper) end if
    if typename(aTrigger) <> "Trigger" then msgbox "FlipperPolarity: .SetObjects error: third argument must be a trigger. Found:" & typename(aTrigger) end if
    if aFlipper.EndAngle > aFlipper.StartAngle then LR = -1 Else LR = 1 End If
    Name = aName
    Set Flipper = aFlipper : FlipperStart = aFlipper.x
    FlipperEnd = Flipper.Length * sin((Flipper.StartAngle / 57.295779513082320876798154814105)) + Flipper.X ' big floats for degree to rad conversion
    FlipperEndY = Flipper.Length * cos(Flipper.StartAngle / 57.295779513082320876798154814105)*-1 + Flipper.Y

    dim str : str = "sub " & aTrigger.name & "_Hit() : " & aName & ".AddBall ActiveBall : End Sub'"
    ExecuteGlobal(str)
    str = "sub " & aTrigger.name & "_UnHit() : " & aName & ".PolarityCorrect ActiveBall : End Sub'"
    ExecuteGlobal(str)

  End Sub

  Public Property Let EndPoint(aInput) :  : End Property ' Legacy: just no op

  Public Sub AddPt(aChooseArray, aIDX, aX, aY) 'Index #, X position, (in) y Position (out)
    Select Case aChooseArray
      case "Polarity" : ShuffleArrays PolarityIn, PolarityOut, 1 : PolarityIn(aIDX) = aX : PolarityOut(aIDX) = aY : ShuffleArrays PolarityIn, PolarityOut, 0
      Case "Velocity" : ShuffleArrays VelocityIn, VelocityOut, 1 :VelocityIn(aIDX) = aX : VelocityOut(aIDX) = aY : ShuffleArrays VelocityIn, VelocityOut, 0
      Case "Ycoef" : ShuffleArrays YcoefIn, YcoefOut, 1 :YcoefIn(aIDX) = aX : YcoefOut(aIDX) = aY : ShuffleArrays YcoefIn, YcoefOut, 0
    End Select
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
  Private Function FlipperOn() : if gameTime < FlipAt+TimeDelay then FlipperOn = True : End If : End Function   'Timer shutoff for polaritycorrect

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
          if ballpos > 0.65 then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)                'find safety coefficient 'ycoef' data
        end if
      Next

      If BallPos = 0 Then 'no ball data meaning the ball is entering and exiting pretty close to the same position, use current values.
        BallPos = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
        if ballpos > 0.65 then  Ycoef = LinearEnvelope(aBall.Y, YcoefIn, YcoefOut)                        'find safety coefficient 'ycoef' data
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
        dim AddX : AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR

        if Enabled then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef)
      End If
      if DebugOn then debug.print "PolarityCorrect" & " " & Name & " @ " & gametime & " " & Round(BallPos*100) & "%" & " AddX:" & Round(AddX,2) & " Vel%:" & Round(VelCoef*100)
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
  for x = 0 to uBound(aArray)   'Shuffle objects in a temp array
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
  redim aArray(aCount-1+offset)   'Resize original array
  for x = 0 to aCount-1       'set objects back into original array
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
Function PSlope(Input, X1, Y1, X2, Y2)    'Set up line via two points, no clamping. Input X, output Y
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
  dim ii : for ii = 1 to uBound(xKeyFrame)    'find active line
    if xInput <= xKeyFrame(ii) then L = ii : exit for : end if
  Next
  if xInput > xKeyFrame(uBound(xKeyFrame) ) then L = uBound(xKeyFrame)    'catch line overrun
  Y = pSlope(xInput, xKeyFrame(L-1), yLvl(L-1), xKeyFrame(L), yLvl(L) )

  if xInput <= xKeyFrame(lBound(xKeyFrame) ) then Y = yLvl(lBound(xKeyFrame) )     'Clamp lower
  if xInput >= xKeyFrame(uBound(xKeyFrame) ) then Y = yLvl(uBound(xKeyFrame) )    'Clamp upper

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

Function DistancePL(px,py,ax,ay,bx,by) 'Distance between a point and a line where point is px,py
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
    Dim b', BOT
    '   BOT = GetBalls

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

Const LiveDistanceMin = 30  'minimum distance in vp units from flipper base live catch dampening will occur
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
    If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 1 Then FlippersD.Dampenf Activeball, parm
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
  RubbersD.dampen Activeball
  TargetBouncer Activeball, 1
End Sub

Sub dSleeves_Hit(idx)
  SleevesD.Dampen Activeball
  TargetBouncer Activeball, 0.7
End Sub

Dim RubbersD        'frubber
Set RubbersD = New Dampener
RubbersD.name = "Rubbers"
RubbersD.debugOn = False  'shows info in textbox "TBPout"
RubbersD.Print = False    'debug, reports in debugger (in vel, out cor); cor bounce curve (linear)

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
SleevesD.Print = False    'debug, reports in debugger (in vel, out cor)
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
    If gametime > 100 Then Report
  End Sub

  Public Sub Dampen(aBall)
    If threshold Then
      If BallSpeed(aBall) < threshold Then Exit Sub
    End If
    Dim RealCOR, DesiredCOR, str, coef
    DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
    RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id) + 0.0001)
    coef = desiredcor / realcor
    If debugOn Then str = name & " in vel:" & Round(cor.ballvel(aBall.id),2 ) & vbNewLine & "desired cor: " & Round(desiredcor,4) & vbNewLine & _
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
    allBalls = getballs

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
'   ZBOU: VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************
Sub TargetBouncer(aBall,defvalue)
  Dim zMultiplier, vel, vratio
  If TargetBouncerEnabled = 1 And aball.z < 30 Then
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
  End If
End Sub

'Add targets or posts to the TargetBounce collection if you want to activate the targetbouncer code from them
Sub TargetBounce_Hit(idx)
  TargetBouncer activeball, 1
End Sub

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
' Tutorial vides by Apophis
' Audio : Adding Fleep Part 1       https://youtu.be/rG35JVHxtx4
' Audio : Adding Fleep Part 2       https://youtu.be/dk110pWMxGo
' Audio : Adding Fleep Part 3       https://youtu.be/ESXWGJZY_EI


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
  finalspeed = Sqr(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
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
  finalspeed = Sqr(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
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
  finalspeed = Sqr(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
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
  If Abs(cor.ballvelx(activeball.id) < 4) And cor.ballvely(activeball.id) > 7 Then
    RandomSoundBottomArchBallGuideHardHit()
  Else
    RandomSoundBottomArchBallGuide
  End If
End Sub

'/////////////////////////////  FLIPPER BALL GUIDE  ////////////////////////////

Sub RandomSoundFlipperBallGuide()
  Dim finalspeed
  finalspeed = Sqr(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
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
  finalspeed = Sqr(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 10 Then
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
  PlaySoundAtLevelStatic ("Gate_FastTrigger_" & Int(Rnd * 2) + 1), GateSoundLevel, Activeball
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
  PlaySoundAtLevelActiveBall ("Arch_L" & Int(Rnd * 4) + 1), Vol(ActiveBall) * ArchSoundFactor
End Sub

Sub RandomSoundRightArch()
  PlaySoundAtLevelActiveBall ("Arch_R" & Int(Rnd * 4) + 1), Vol(ActiveBall) * ArchSoundFactor
End Sub

Sub Arch1_hit()
  If Activeball.velx > 1 Then SoundPlayfieldGate
  StopSound "Arch_L1"
  StopSound "Arch_L2"
  StopSound "Arch_L3"
  StopSound "Arch_L4"
End Sub

Sub Arch1_unhit()
  If activeball.velx <  - 8 Then
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
  PlaySoundAtLevelStatic ("Saucer_Enter_" & Int(Rnd * 2) + 1), SaucerLockSoundLevel, Activeball
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

'/////////////////////////////////////////////////////////////////
'         End Mechanical Sounds
'/////////////////////////////////////////////////////////////////

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
  Dim BOT
  BOT = GetBalls

  ' stop the sound of deleted balls
  For b = UBound(BOT) + 1 To tnob - 1
    ' Comment the next line if you are not implementing Dyanmic Ball Shadows
    If AmbientBallShadowOn = 0 Then BallShadowA(b).visible = 0
    rolling(b) = False
    StopSound("BallRoll_" & b)
  Next

  ' exit the sub if no balls on the table
  If UBound(BOT) =  - 1 Then Exit Sub

  ' play the rolling sound for each ball
  For b = 0 To UBound(BOT)
    If BallVel(BOT(b)) > 1 And BOT(b).z < 30 Then
      rolling(b) = True
      PlaySound ("BallRoll_" & b), - 1, VolPlayfieldRoll(BOT(b)) * BallRollVolume * VolumeDial, AudioPan(BOT(b)), 0, PitchPlayfieldRoll(BOT(b)), 1, 0, AudioFade(BOT(b))
    Else
      If rolling(b) = True Then
        StopSound("BallRoll_" & b)
        rolling(b) = False
      End If
    End If

    ' Ball Drop Sounds
    If BOT(b).VelZ <  - 1 And BOT(b).z < 55 And BOT(b).z > 27 Then 'height adjust for ball drop sounds
      If DropCount(b) >= 5 Then
        DropCount(b) = 0
        If BOT(b).velz >  - 7 Then
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
    ' Comment the next If block, if you are not implementing the Dynamic Ball Shadows
    If AmbientBallShadowOn = 0 Then
      If BOT(b).Z > 30 Then
        BallShadowA(b).height = BOT(b).z - BallSize / 4   'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping the ramp
      Else
        BallShadowA(b).height = 0.1
      End If
      BallShadowA(b).Y = BOT(b).Y + offsetY
      BallShadowA(b).X = BOT(b).X + offsetX
      BallShadowA(b).visible = 1
    End If
  Next
End Sub

'******************************************************
'****  END BALL ROLLING AND DROP SOUNDS
'******************************************************

Sub GameTimer_Timer()
  RollingUpdate   ' Update rolling sounds
End Sub

'******************************************************
'   ZRRL: RAMP ROLLING SFX
'******************************************************

'Ball tracking ramp SFX 1.0
'   Reqirements:
'     * Import a Sound File for each ball on the table for plastic ramps.  Call It RampLoop<Ball_Number> ex: RampLoop1, RampLoop2, ...
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
Dim RampBalls(6,2)
'x,0 = ball x,1 = ID, 2 = Protection against ending early (minimum amount of updates)

'0,0 is boolean on/off, 0,1 unused for now
RampBalls(0,0) = False

' RampType
' Setup: Set this array to the number Total number of balls that can be tracked at one time + 1.  5 ball multiball then set value to 6
' Description: Array type indexed on BallId and a values used to deterimine what type of ramp the ball is on: False = Wire Ramp, True = Plastic Ramp
Dim RampType(6)

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
      Debug.print "WireRampOn error, ball queue is full: " & vbNewLine & _
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

Dim TestFlashers, TableRef, FlasherLightIntensity, FlasherFlareIntensity, FlasherBloomIntensity, FlasherOffBrightness

' *********************************************************************
TestFlashers = 0        ' *** set this to 1 to check position of flasher object      ***
Set TableRef = Table1      ' *** change this, if your table has another name           ***
FlasherLightIntensity = 0.1  ' *** lower this, if the VPX lights are too bright (i.e. 0.1)     ***
FlasherFlareIntensity = 0.3  ' *** lower this, if the flares are too bright (i.e. 0.1)       ***
FlasherBloomIntensity = 0.2  ' *** lower this, if the blooms are too bright (i.e. 0.1)       ***
FlasherOffBrightness = 0.5    ' *** brightness of the flasher dome when switched off (range 0-2)  ***
' *********************************************************************

Dim ObjLevel(20), objbase(20), objlit(20), objflasher(20), objbloom(20), objlight(20), ObjTargetLevel(20)
'Dim tablewidth, tableheight : tablewidth = TableRef.width : tableheight = TableRef.height

'initialise the flasher color, you can only choose from "green", "red", "purple", "blue", "white" and "yellow"
InitFlasher 1, "red"
InitFlasher 2, "yellow"
InitFlasher 3, "blue"
InitFlasher 4, "red"

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
Const lob = 0 ' Locked balls on start

'****** Part A:  Table Elements ******

'****** Part B:  Code and Functions ******

' *** Required Functions, enable these if they are not already present elswhere in your table
Function max(a,b)
  If a > b Then
    max = a
  Else
    max = b
  End If
End Function

'****** Part C:  The Magic ******

' *** These define the appearance of shadows in your table  ***

'Ambient (Room light source)
Const AmbientBSFactor = 0.9  '0 to 1, higher is darker
Const AmbientMovement = 1    '1+ higher means more movement as the ball moves left and right
Const offsetX = 0        'Offset x position under ball (These are if you want to change where the "room" light is for calculating the shadow position,)
Const offsetY = 5        'Offset y position under ball (^^for example 5,5 if the light is in the back left corner)

'Dynamic (Table light sources)
Const DynamicBSFactor = 0.90  '0 to 1, higher is darker
Const Wideness = 20      'Sets how wide the dynamic ball shadows can get (20 +5 thinness is technically most accurate for lights at z ~25 hitting a 50 unit ball)
Const Thinness = 5        'Sets minimum as ball moves away from source

' *** Trim or extend these to match the number of balls/primitives/flashers on the table!  (will throw errors if there aren't enough objects)
Dim objrtx1(5), objrtx2(5)
Dim objBallShadow(5)
Dim OnPF(5)
Dim BallShadowA
BallShadowA = Array (BallShadowA0,BallShadowA1,BallShadowA2,BallShadowA3,BallShadowA4)
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
  If onPlayfield Then
    Dim gBOT: gBOT=getballs
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
  Dim gBOT: gBOT=getballs 'Uncomment if you're destroying balls - Not recommended! #SaveTheBalls

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


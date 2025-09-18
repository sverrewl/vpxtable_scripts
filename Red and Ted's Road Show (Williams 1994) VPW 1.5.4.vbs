'*************   VPW Presents   *************
'Red & Ted's Road Show (Williams 1994)
'https://www.ipdb.org/machine.cgi?id=1972
'********************************************
'
'         --- Red & Teds ---
'  ____                 _   ____  _
' |  _ \ ___   __ _  __| | / ___|| |__   _____      __
' | |_) / _ \ / _` |/ _` | \___ \| '_ \ / _ \ \ /\ / /
' |  _ < (_) | (_| | (_| |  ___) | | | | (_) \ V  V /
' |_| \_\___/ \__,_|\__,_| |____/|_| |_|\___/ \_/\_/
'
'
'***************
' VPW Road Crew
'***************
'Sixtoe - Project Lead
'Skitso - Lighting
'Clark Kent - Physics Tweaking, real table comparison
'Additional Script Work - apophis, fluffhead, Wylte
'Testing - Rik, Pinstratsdan, VPW Team
'
'Original Table Thanks;
'Knorr and Clark Kent

Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

'*******************************************
' User Options
'*******************************************

'----- Shadow Options -----
Const DynamicBallShadowsOn = 1    '0 = no dynamic ball shadow ("triangles" near slings and such), 1 = enable dynamic ball shadow
Const AmbientBallShadowOn = 1   '0 = Static shadow under ball ("flasher" image, like JP's)
                  '1 = Moving ball shadow ("primitive" object, like ninuzzu's) - This is the only one that shows up on the pf when in ramps and fades when close to lights!
                  '2 = flasher image shadow, but it moves like ninuzzu's

'----- General Sound Options -----
'Const VolumeDial = 0.8       'Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
'Const BallRollVolume = 0.8     'Level of ball rolling volume. Value between 0 and 1
'Const RampRollVolume = 0.9     'Level of ramp rolling volume. Value between 0 and 1

Const AltWireSounds = 0             'Change to 1 to Enable Alternate Wire Ramp Sounds



Dim ShakerSoundVolume: ShakerSoundVolume = 1  ' Shaker Volume (Ranges between 0 and 1)
Dim ShakerIntensity : ShakerIntensity = 3     ' 0 = Off, 1 = Low, 2 = Normal, 3 = High
Dim VRRoomChoice : VRRoomChoice = 1       ' 1 - Minimal Room, 2 - Ultra Minimal
Dim RightOutpostMod : RightOutpostMod = 1   ' 0 = Easy, 1 = Normal, 2 = Hard
Dim LeftOutpostMod : LeftOutpostMod = 1     ' 0 = Easy, 1 = Normal, 2 = Hard
Dim CustomDecals : CustomDecals = 1       ' 0 - Original, 1 - Custom Target and Plunger Gate Stickers.
Dim FlipperDecals : FlipperDecals = 1     ' 0 = original, 1 = new
Dim VolumeDial : VolumeDial = 0.8             ' Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Dim BallRollVolume : BallRollVolume = 0.8     ' Level of ball rolling volume. Value between 0 and 1
Dim RampRollVolume : RampRollVolume = 0.9     ' Level of ramp rolling volume. Value between 0 and 1

' Called when options are tweaked by the player.
' - 0: game has started, good time to load options and adjust accordingly
' - 1: an option has changed
' - 2: options have been reseted
' - 3: player closed the tweak UI, good time to update staticly prerendered parts
' Table1.Option arguments are:
' - option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional arry of literal strings
Dim dspTriggered : dspTriggered = False
Sub Table1_OptionEvent(ByVal eventId)
    If eventId = 1 And Not dspTriggered Then dspTriggered = True : DisableStaticPreRendering = True : End If

    Dim x, v

  ' Right Outlane  difficulty
    RightOutpostMod = Table1.Option("Right Outpost Difficulty", 0, 2, 1, 1, 0, Array("Easy", "Normal", "Hard"))
    If RightOutpostMod = 0 Then     'Easy
    zCol_Rubber_Post013.Collidable = 1
    zCol_Rubber_Post005.Collidable = 0
    zCol_Rubber_Post012.Collidable = 0
    RubberAPRightUpper.Visible = 0
    RubberAPRightLower.Visible = 0
    RubberAPRightUpperEasy.Visible = 1
    RubberAPRightLowerEasy.Visible = 1
    RubberAPRightUpperHard.Visible = 0
    RubberAPRightLowerHard.Visible = 0
    AdjustingPostRightOutlane.Visible = 0
    AdjustingPostRightOutlaneEasy.Visible = 1
    AdjustingPostRightOutlaneHard.Visible = 0
  ElseIf RightOutpostMod = 1 Then 'Normal
    zCol_Rubber_Post005.Collidable = 1
    zCol_Rubber_Post013.Collidable = 0
    zCol_Rubber_Post012.Collidable = 0
    RubberAPRightUpper.Visible = 1
    RubberAPRightLower.Visible = 1
    RubberAPRightUpperEasy.Visible = 0
    RubberAPRightLowerEasy.Visible = 0
    RubberAPRightUpperHard.Visible = 0
    RubberAPRightLowerHard.Visible = 0
    AdjustingPostRightOutlane.Visible = 1
    AdjustingPostRightOutlaneEasy.Visible = 0
    AdjustingPostRightOutlaneHard.Visible = 0
  Else              'Hard
    zCol_Rubber_Post012.Collidable = 1
    zCol_Rubber_Post005.Collidable = 0
    zCol_Rubber_Post013.Collidable = 0
    RubberAPRightUpper.Visible = 0
    RubberAPRightLower.Visible = 0
    RubberAPRightUpperEasy.Visible = 0
    RubberAPRightLowerEasy.Visible = 0
    RubberAPRightUpperHard.Visible = 1
    RubberAPRightLowerHard.Visible = 1
    AdjustingPostRightOutlane.Visible = 0
    AdjustingPostRightOutlaneEasy.Visible = 0
    AdjustingPostRightOutlaneHard.Visible = 1
  End If

  ' Left Outlane  difficulty
    LeftOutpostMod = Table1.Option("Left Outpost Difficulty", 0, 2, 1, 1, 0, Array("Easy", "Normal", "Hard"))
    If LeftOutpostMod = 0 Then    'Easy
    zCol_Rubber_Post010.Collidable = 0
    zCol_Rubber_Post014.Collidable = 1
    zCol_Rubber_Post015.Collidable = 0
    RubberBand005.Collidable = 1
    RubberBand006.Collidable = 0
    RubberBand001.Collidable = 0
    RubberAPLeftUpper.Visible = 0
    RubberAPLeftLower.Visible = 0
    RubberAPLeftUpperEasy.Visible = 1
    RubberAPLeftLowerEasy.Visible = 1
    RubberAPLeftUpperHard.Visible = 0
    RubberAPLeftLowerHard.Visible = 0
    AdjustingPostLeftOutlane.Visible = 0
    AdjustingPostLeftOutlaneEasy.Visible = 1
    AdjustingPostLeftOutlaneHard.Visible = 0
  ElseIf LeftOutpostMod = 1 Then  'Normal
    zCol_Rubber_Post010.Collidable = 1
    zCol_Rubber_Post014.Collidable = 0
    zCol_Rubber_Post015.Collidable = 0
    RubberBand005.Collidable = 0
    RubberBand006.Collidable = 0
    RubberBand001.Collidable = 1
    RubberAPLeftUpper.Visible = 1
    RubberAPLeftLower.Visible = 1
    RubberAPLeftUpperEasy.Visible = 0
    RubberAPLeftLowerEasy.Visible = 0
    RubberAPLeftUpperHard.Visible = 0
    RubberAPLeftLowerHard.Visible = 0
    AdjustingPostLeftOutlane.Visible = 1
    AdjustingPostLeftOutlaneEasy.Visible = 0
    AdjustingPostLeftOutlaneHard.Visible = 0
  Else              'Hard
    zCol_Rubber_Post010.Collidable = 0
    zCol_Rubber_Post014.Collidable = 0
    zCol_Rubber_Post015.Collidable = 1
    RubberBand005.Collidable = 0
    RubberBand006.Collidable = 1
    RubberBand001.Collidable = 0
    RubberAPLeftUpper.Visible = 0
    RubberAPLeftLower.Visible = 0
    RubberAPLeftUpperEasy.Visible = 0
    RubberAPLeftLowerEasy.Visible = 0
    RubberAPLeftUpperHard.Visible = 1
    RubberAPLeftLowerHard.Visible = 1
    AdjustingPostLeftOutlane.Visible = 0
    AdjustingPostLeftOutlaneEasy.Visible = 0
    AdjustingPostLeftOutlaneHard.Visible = 1
  End If

  'Custom Decals
  CustomDecals = Table1.Option("Custom Decals", 0, 1, 1, 1, 0, Array("Not Visible", "Visible"))
  If CustomDecals = 1 Then
    T28a.image = "TargetBZ3"    'Blast Zone
    T28b.image = "TargetBZ2"
    T28c.image = "TargetBZ1"

    T34a.image = "HitTargetRad2"  'Radio Targets
    T34b.image = "HitTargetRad3"
    T34c.image = "HitTargetRad2"

    T35.image = "HitTargetRad1"   'Right Ramp Target
    T36.image = "HitTargetRad1"

    T81.image = "HitTargetConeWhite"    'Cones
    T82.image = "HitTargetConeOrange"
    T83.image = "HitTargetConeYellow"
    T84.image = "HitTargetConeOrange"
    PlungerGateDecal.visible = 1
  Else
    T28a.image = "HitTargetY"   'Blast Zone
    T28b.image = "HitTargetY"
    T28c.image = "HitTargetY"

    T34a.image = "HitTarget"    'Radio Targets
    T34b.image = "HitTarget"
    T34c.image = "HitTarget"

    T35.image = "HitTarget"     'Right Ramp Target
    T36.image = "HitTarget"

    T81.image = "HitTargetW"      'Cones
    T82p.image = "HitTarget"
    T83.image = "HitTargetY"
    T84.image = "HitTarget"
    PlungerGateDecal.visible = 0
  End If

  'Flipper Decals
  FlipperDecals = Table1.Option("Flipper Decals", 0, 1, 1, 1, 0, Array("Not Visible", "Visible"))
  If FlipperDecals = 1 Then
    FlipperLPDecals.visible = 1
    FlipperRPDecals.visible = 1
  Else
    FlipperLPDecals.visible = 0
    FlipperRPDecals.visible = 0
  End If

  'Desktop DMD
  v = Table1.Option("Desktop DMD", 0, 1, 1, 1, 0, Array("Not Visible", "Visible"))
  If v = 1 Then
    Scoretext.Visible = 1
  Else
    Scoretext.Visible = 0
  End If

  ' Side rails (always for VR, and if not in cabinet mode)
  v = Table1.Option("Cabinet Rails", 0, 1, 1, 1, 0, Array("Off", "On"))
  If v = 1 Then
    PinCab_Rails.Visible = 1
  Else
    PinCab_Rails.Visible = 0
  End If


  ' Shaker SSF Intensity
    ShakerIntensity = Table1.Option("Shaker SSF Intensity", 0, 3, 1, 2, 0, Array("Off", "Low", "Normal", "High"))

  ' Flasher Brightness
    v = Table1.Option("Flasher Brightness", 0, 2, 1, 1, 0, Array("Low", "Mid", "High"))
  Select Case v
    Case 0: FlasherLightIntensity = 0.3: FlasherFlareIntensity = 0.3: FlasherBloomIntensity = 0
    Case 1: FlasherLightIntensity = 0.3: FlasherFlareIntensity = 0.3: FlasherBloomIntensity = 0.2
    Case 2: FlasherLightIntensity = 0.4: FlasherFlareIntensity = 0.5: FlasherBloomIntensity = 0.5
  End Select


  ' Playfield Reflections
    v = Table1.Option("Playfield Reflections", 0, 1, 1, 1, 0, Array("Off", "On"))
  If v = 1 Then
    playfield_mesh_vis.ReflectionProbe = "Playfield Reflections"
  Else
    playfield_mesh_vis.ReflectionProbe = ""
  End If

  'VR Room
  VRRoomChoice = Table1.Option("VR Room", 1, 2, 1, 1, 0, Array("Minimal", "Ultra Minimal"))
  SetupVRRoom

    ' Sound volumes
  ShakerSoundVolume = Table1.Option("Shaker SSF Volume", 0, 1, 0.01, 0.8, 1)
    VolumeDial = Table1.Option("Mech Volume", 0, 1, 0.01, 0.8, 1)
    RampRollVolume = Table1.Option("Ramp Volume", 0, 1, 0.01, 0.5, 1)
    BallRollVolume = Table1.Option("Ball Roll Volume", 0, 1, 0.01, 0.5, 1)

    If eventId = 3 And dspTriggered Then dspTriggered = False : DisableStaticPreRendering = False : End If
End Sub


'*******************************************
' Constants and Global Variables
'*******************************************

Const BallSize = 50         'Ball size must be 50
Const BallMass = 1          'Ball mass must be 1
Const tnob = 4            'Total number of balls
Const lob = 0           'Locked balls

Dim tablewidth: tablewidth = Table1.width
Dim tableheight: tableheight = Table1.height
Const UseVPMModSol = True

Dim DesktopMode:DesktopMode = Table1.ShowDT
Dim UseVPMColoredDMD, CabinetMode, VRRoom
If RenderingMode = 2 Then UseVPMColoredDMD = True Else UseVPMColoredDMD = DesktopMode

'Sub ImplicitDMD_Init
'   Me.x = 5
'   Me.y = 5
'   Me.width = 160 * 2
'   Me.height = 55 * 2
'   Me.visible = true
'   Me.intensityScale = 1.5
'   Me.FontColor = RGB(255, 88, 32)
'End Sub

LoadVPM "01560000", "wpc.VBS", 3.36

'**************************
' Standard definitions
'**************************

Const cGameName = "rs_l6"
Const UseSolenoids = 2
Const UseLamps = 0
Const UseSync = 0
Const HandleMech = 0
Const SSolenoidOn = ""
Const SSolenoidOff = ""
Const SCoin = ""


'****************
' Table init
'****************

Dim bsTrough, StartCity, bsLockKickout, bsLock, bsRed, BullDozerMech, TedJawMech, RedJawMech
Dim PPL, aBall, aBall77a, aBall77b, aBall77c, aBall77d

Dim RSBall1, RSBall2, RSBall3, RSBall4, gBOT

Sub table1_Init
    vpmInit Me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
    .SplashInfoLine = "Red and 's Road Show (Williams 1994)" & vbNewLine & "VPW"
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        On Error Resume Next
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
    End With

  '************  Trough **************
  Set RSBall4 = sw42.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set RSBall3 = sw43.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set RSBall2 = sw44.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set RSBall1 = sw45.CreateSizedballWithMass(Ballsize/2,Ballmass)
  gBOT = Array(RSBall1,RSBall2,RSBall3,RSBall4)

  Controller.Switch(42) = 1
  Controller.Switch(43) = 1
  Controller.Switch(44) = 1
  Controller.Switch(45) = 1

  Controller.Switch(22) = 1 'close coin door
  Controller.Switch(24) = 1 'and keep it close

  PinMAMETimer.Interval = PinMAMEInterval
  PinMAMETimer.Enabled = true

  vpmNudge.TiltSwitch = 14
  vpmNudge.Sensitivity = 5
  vpmNudge.TiltObj=Array(Leftjetbumper,rightjetbumper,topjetbumper,Slingshotleft,slingshotright)

' Set StartCity = New cvpmBallStack
' With StartCity
'   .InitSaucer sw78, 78, 0, 60
'   .KickZ=0.8
'        .InitExitSnd SoundFX("StartCity_SolOut",DOFContactors), SoundFX("solenoid",DOFContactors)
'    End With

' Thalamus - add some randomness
  'Red Kickout
    Set bsRed = new cvpmBallStack
    With bsRed
    .InitSaucer sw25, 25, 195, 20
    .InitExitSnd SoundFX("SafeHouseKick",DOFContactors), SoundFX("solenoid",DOFContactors)
    .KickForceVar = 3
    .KickAngleVar = 3
    End With

    Set BulldozerMech = New cvpmMech
    With BulldozerMech
        .MType = vpmMechOneSol + vpmMechReverse + vpmMechNonLinear + vpmMechFast
        .Sol1 = 23
        .Length = DozerSpeed
        .Steps = 11
        .AddSw 12, 0, 0
        .AddSw 15, 10, 11
        .Callback = GetRef("UpdateDozer")
        .Start
    End With

  Set TedJawMech = New cvpmMech
  With TedJawMech
    .MType = vpmMechOneDirSol + vpmMechStopEnd + vpmMechNonLinear + vpmMechFast
    .Sol1 = 20
    .Sol2 = 19
    .length = jawspeed
    .steps = 18
    .callback = getRef("UpdateJawTed")
    .start
  End With

  Set RedJawMech = New cvpmMech
  With RedJawMech
    .MType = vpmMechOneDirSol + vpmMechStopEnd + vpmMechNonLinear + vpmMechFast
    .Sol1 = 17
    .Sol2 = 18
    .length = jawspeed
    .steps = 18
    .callback = getRef("UpdateJawRed")
    .start
  End With

  diverterleftramp1.isDropped = 1

    If table1.ShowDT = False then
        sideblades.Size_Y = 3
    sideblades.Z = 60
    GiFlasher3.Height = 200
    GiFlasher4.Height = 200
    GiFlasher.Height = 220
    GiFlasher1.Height = 220
    GiFlasher2.Height = 220
    GiFlasher5.Height = 220
    GiFlasher6.Height = 220
    GiFlasher7.Height = 220
    GiFlasher8.Height = 220
    Else
    GiFlasher3.Height = 100
    GiFlasher4.Height = 100
    GiFlasher.Height = 120
    GiFlasher1.Height = 120
    GiFlasher2.Height = 120
    GiFlasher5.Height = 120
    GiFlasher6.Height = 120
    GiFlasher7.Height = 120
    GiFlasher8.Height = 120
    End if

  sol3lockup.isDropped = 0
  lockl.state = 1

  vpmTimer.AddTimer 200, "InitAllFlashers'"

End Sub

Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub
Sub table1_Exit:Controller.Stop:End Sub


Sub InitAllFlashers
  InitFlasher 1, "white"
  InitFlasher 2, "white"
  InitFlasher 3, "yellow"
  InitFlasher 4, "yellow"
  InitFlasher 5, "red"
  InitFlasher 6, "red"
  InitFlasher 7, "yellow"
  InitFlasher 8, "white"
  InitFlasher 9, "white"
  InitFlasher 10, "orange"
  InitFlasher 11, "white"

  ' rotate the flasher with the command below (first argument = flasher nr, second argument = angle in degrees)
  RotateFlasher 9,0 : RotateFlasher 10,-17' : RotateFlasher 3,90 : RotateFlasher 4,90

  FlInitBumper 1, "yellow"
  FlInitBumper 2, "red"
  FlInitBumper 3, "white"
End Sub



'****************************
' Animations
'****************************

Sub LeftFlipper_Animate
  dim a: a = LeftFlipper.CurrentAngle
  FlipperLP.ObjRotZ = a
  FlipperLPDecals.ObjRotZ = a
  batleftshadow.ObjRotZ = a
End Sub

Sub RightFlipper_Animate
  dim a: a = RightFlipper.CurrentAngle
  FlipperRP.ObjRotZ = a
  FlipperRPDecals.ObjRotZ = a
  batrightshadow.ObjRotZ = a
End Sub

Sub LeftFlipper1_Animate
  dim a: a = LeftFlipper1.CurrentAngle
  FlipperL2P.ObjRotZ = a
  batleft2shadow.ObjRotZ = a
End Sub

Sub LeftFlipper2_Animate
  dim a: a = LeftFlipper2.CurrentAngle
  FlipperL3P.ObjRotZ = a
  batleft3shadow.ObjRotZ = a
End Sub

Sub Diverter1_animate
  Diverter1P.RotY = Diverter1.CurrentAngle
End Sub

Sub Diverter2_animate
  Diverter2P.RotY = Diverter2.CurrentAngle
End Sub

Sub sw51spinner_animate
  SpinnerPrimitive.RotX = sw51spinner.CurrentAngle +90
End Sub

Sub GatePlungerLane_animate
  GatePlungerLaneP.RotX = GatePlungerLane.CurrentAngle +90
End Sub




'****************************
' Timers
'****************************

Sub Frametimer_Timer  'The frame timer interval is -1 (once per frame).
  RollingUpdate   'update rolling sounds
  If DynamicBallShadowsOn Or AmbientBallShadowOn Then DynamicBSUpdate 'update ball shadows
  If VRRoom <> 0 Then
    Pincab_ShooterR.Y = 1162 + (5* PlungerR.Position) -20
    Pincab_ShooterRtop.Y = 1162 + (5* PlungerR.Position) -20
    Pincab_ShooterL.Y = 1162 + (5* PlungerL.Position) -20
    Pincab_ShooterLtop.Y = 1162 + (5* PlungerL.Position) -20
  End If


  ' Flashers
  l158.state = Flasherlight10.state
  l158.IntensityScale = Flasherlight10.IntensityScale
  f158.state = Flasherlight10.state
  f158.IntensityScale = Flasherlight10.IntensityScale
  p158.visible = Flasherlit10.visible
  p158.blenddisablelighting = 2 + Flasherlight10.state


  sideflasher1.visible = Flasherbloom2.visible
  sideflasher1.opacity = Flasherbloom2.opacity

  sideflasher2.visible = Flasherbloom2.visible
  sideflasher2.opacity = Flasherbloom2.opacity

  sideflasher3.visible = Flasherbloom3.visible
  sideflasher3.opacity = Flasherbloom3.opacity

  sideflasher4.visible = Flasherbloom3.visible
  sideflasher4.opacity = Flasherbloom3.opacity

  sideflasher5.visible = Flasherbloom6.visible
  sideflasher5.opacity = Flasherbloom6.opacity

  sideflasher6.visible = Flasherbloom6.visible
  sideflasher6.opacity = Flasherbloom6.opacity

  sideflasher7.visible = Flasherbloom7.visible
  sideflasher7.opacity = Flasherbloom7.opacity

  sideflasher8.visible = Flasherbloom8.visible
  sideflasher8.opacity = Flasherbloom8.opacity

  sideflasher9.visible = Flasherbloom9.visible
  sideflasher9.opacity = Flasherbloom9.opacity

  sideflasher10.visible = Flasherflash10.visible
  sideflasher10.opacity = Flasherflash10.opacity

  sideflasher11.visible = Flasherbloom11.visible
  sideflasher11.opacity = Flasherbloom11.opacity

  f11a.visible = Flasherlight11.visible
  f11a.IntensityScale = Flasherlight11.IntensityScale
  f11b.visible = Flasherlight11.visible
  f11b.IntensityScale = Flasherlight11.IntensityScale

End Sub

Sub GameTimer_Timer() 'The game timer interval is 10 ms
  Cor.Update      'update ball tracking (this sometimes goes in the RDampen_Timer sub)
End Sub




'**********
' Keys
'**********

Sub Table1_KeyDown(ByVal keycode)
  If keycode = LeftFlipperKey Then
    PinCab_ButtonL.X = PinCab_ButtonL.X + 10
  End If
  If keycode = RightFlipperKey Then
    PinCab_ButtonR.X = PinCab_ButtonR.X - 10
  End If

  If PPL Then
    If keycode = PlungerKey Then PlungerL.Pullback : SoundPlungerPull
  Else
    If keycode = PlungerKey Then PlungerR.Pullback : SoundPlungerPull
  End If
  If keycode = LeftMagnaSave Then PlungerL.Pullback : SoundPlungerPull
  If keycode = LeftTiltKey Then Nudge 90, 1 : SoundNudgeLeft
  If keycode = RightTiltKey Then Nudge 270, 1 : SoundNudgeRight
  If keycode = CenterTiltKey Then Nudge 0, 1 : SoundNudgeCenter
  If keycode = StartGameKey Then
    SoundStartButton
    StartButton.y = StartButton.y - 4
    StartButtonInner.y = StartButtonInner.y - 4
  End If
  If Keycode = KeyFront Then Controller.Switch(23) = 1
  If keycode = AddCreditKey or keycode = AddCreditKey2 Then
    Select Case Int(rnd*3)
      Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
      Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
      Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
    End Select
  End If
  If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal keycode)
  If keycode = LeftFlipperKey Then
    PinCab_ButtonL.X = PinCab_ButtonL.X - 10
  End If
  If keycode = RightFlipperKey Then
    PinCab_ButtonR.X = PinCab_ButtonR.X + 10
  End If

  If keycode = StartGameKey Then
    StartButton.y = StartButton.y + 4
    StartButtonInner.y = StartButtonInner.y + 4
  End If

  If PPL Then
    If keycode = PlungerKey Then PlungerL.Fire : SoundPlungerReleaseBall
  Else
    If keycode = PlungerKey Then PlungerR.Fire : SoundPlungerReleaseBall
  End If
  If Keycode = KeyFront Then Controller.Switch(23) = 0
  If keycode = LeftMagnaSave Then PlungerL.Fire : SoundPlungerPull
  If vpmKeyUp(keycode) Then Exit Sub
End Sub

'*********
'Solenoids
'*********

SolCallback(1) = "SolRelease"
SolCallback(2) = "LowerLeftDiverter"
SolCallback(3) = "LockupPin"
SolCallback(4) = "UpperLeftDiverter"
SolCallback(5) = "UpperRightDiverter"
SolCallback(6) = "StartCitySolOut"
SolCallback(7) = "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallBack(8) = "SolLockKickOut"
SolCallBack(9) = "TedEyesLeft"
SolCallBack(10) = "TedLidsDown"
SolCallBack(11) = "TedLidsUp"
SolCallBack(12) = "TedEyesRight"
SolCallBack(13) = "RedLidsDown"
SolCallBack(14) = "RedEyesLeft"
SolCallBack(15) = "RedLidsUp"
SolCallBack(16) = "RedEyesRight"
SolCallBack(17) = "RedMotorOn"
'SolCallBack(18) = "RedMotorDir"
'SolCallBack(19) = "TedMotorDir"
SolCallBack(20) = "TedMotorOn"
SolCallBack(23) = "BullDozerMotor"
SolCallBack(24) = "bsRed.SolOut"
SolCallBack(28) = "ShakerMotorSol"

SolModCallBack(51) = "FlashSolMod51" ' Little Flipper Flasher (Modulated)
SolModCallBack(52) = "FlashSolMod52" ' Left Ramp Flasher (Modulated)
SolModCallBack(53) = "FlashSolMod53" ' Back White Flashers (Modulated)
SolModCallBack(54) = "FlashSolMod54" ' Back Yellow Flashers (Modulated)
SolModCallBack(55) = "FlashSolMod55" ' Back Red Flashers (Modulated)
SolModCallBack(56) = "FlashSolMod56" ' Blasting Zone Flashers (Modulated)
SolModCallBack(57) = "FlashSolMod57" ' Right Ramp Flasher (Modulated)
SolModCallBack(58) = "FlashSolMod58" ' Jets Flasher (Modulated)

SolCallback(sLRFlipper) = "SolRFlipper" 'Right Flipper
SolCallback(sLLFlipper) = "SolLFlipper" 'Left Flipper
SolCallback(sULFlipper) = "SolUFlipper" 'Upper Left Flippers

'TedsMouth
Sub sw11_Hit()
  vpmTimer.PulseSw 11
End Sub

Sub TedTrig_Hit()
  PlaySoundAt "Ted_Hit2", TedTrig
End Sub

Sub sw48_Hit()  'Ted Mouth Plastic
  vpmTimer.PulseSw 48
  PlaySoundAtLevelStatic ("DiverterLeft_Close"), BlastZoneDrop * volumedial, Activeball
End Sub

'Dozer
Sub sw47_Hit()
  vpmTimer.PulseSw 47
End Sub

Sub dozerwall_Hit()
  shakedozer
  PlaySoundAt "flip_hit_3", dozerwallp
End Sub

Sub sw37_Hit()  'Red Mouth Plastic
  PlaySoundAtLevelStatic ("DiverterLeft_Close"), BlastZoneDrop * volumedial, Activeball
End Sub

Sub sw37a_Hit()
  vpmTimer.PulseSw 37
End Sub


'RedsMouth
Sub sw25_Hit()
  PlaySoundAt "SafeHouseHit", sw25
  bsRed.AddBall Me
End Sub

'diverter
Sub UpperRightDiverter(Enabled)
  If Enabled then
    diverterr1.visible = 1
    diverterr2.visible = 0
    diverterrightramp.isDropped = 1
    PlaySound "DiverterRight_Open"
  Else
    diverterr1.visible = 0
    diverterr2.visible = 1
    diverterrightramp.isDropped = 0
    PlaySound "DiverterLeft_Close"
  End if
End Sub

Sub UpperLeftDiverter(Enabled)
  If Enabled then
    diverterl1.visible = 1
    diverterl2.visible = 0
    diverterleftramp1.isDropped = 0
    diverterleftramp2.isDropped = 1
    PlaySound "DiverterRight_Open"
  Else
    diverterl1.visible = 0
    diverterl2.visible = 1
    diverterleftramp1.isDropped = 1
    diverterleftramp2.isDropped = 0
    PlaySound "DiverterLeft_Close"
  End if
End Sub

Sub LowerLeftDiverter(enabled)
  If Enabled Then
    diverter1.rotatetoend
    diverter2.rotatetoend
    PlaySound "DiverterRight_Open"
  Else
    diverter1.rotatetostart
    diverter2.rotatetostart
    PlaySound" DiverterLeft_Close"
  End If
End Sub

Sub StartCitySolOut(Enabled)
  If Enabled Then
    PlaySound SoundFX("StartCity_SolOut",DOFContactors)
    sw78.KickZ 1, 36+rnd*5, 80, 5
  End If
End Sub

'*********
' Bumper
'*********
Sub Bumper3_hit:vpmTimer.pulseSw 63:RandomSoundBumperTop Bumper3:Me.TimerEnabled = 1:End Sub
Sub Bumper3_Timer:Me.Timerenabled = 0:End Sub

Sub Bumper2_hit:vpmTimer.pulseSw 65:RandomSoundBumperMiddle Bumper2:Me.TimerEnabled = 1:End Sub
Sub Bumper2_Timer:Me.Timerenabled = 0:End Sub

Sub Bumper1_hit:vpmTimer.pulseSw 64:RandomSoundBumperBottom Bumper1:Me.TimerEnabled = 1:End Sub
Sub Bumper1_Timer:Me.Timerenabled = 0:End Sub


'*********
' Switches
'*********
Sub sw16_Hit:Controller.Switch(16) = 1:sw16wire.RotX = 15:End Sub
Sub sw16_UnHit:Controller.Switch(16) = 0:sw16wire.RotX = 0:End Sub

Sub sw17_Hit:Controller.Switch(17) = 1:sw17wire.RotX = 15:End Sub
Sub sw17_UnHit:Controller.Switch(17) = 0:sw17wire.RotX = 0:End Sub

Sub sw18_Hit:Controller.Switch(18) = 1:sw18wire.RotX = 15:End Sub
Sub sw18_UnHit:Controller.Switch(18) = 0:sw18wire.RotX = 0:End Sub

Sub sw26_Hit:Controller.Switch(26) = 1:sw26wire.RotX = 15:End Sub
Sub sw26_UnHit:Controller.Switch(26) = 0:sw26wire.RotX = 0:End Sub

Sub sw27_Hit:Controller.Switch(27) = 1:sw27wire.RotX = 15:End Sub
Sub sw27_UnHit:Controller.Switch(27) = 0:sw27wire.RotX = 0:End Sub

Sub sw31_Hit:Controller.Switch(31) = 1:sw31wire.RotX = 15:End Sub
Sub sw31_UnHit:Controller.Switch(31) = 0:sw31wire.RotX = 0:End Sub

Sub sw32_Hit:Controller.Switch(32) = 1:sw32wire.RotX = 15:End Sub
Sub sw32_UnHit:Controller.Switch(32) = 0:sw32wire.RotX = 0:End Sub

Sub sw33_Hit:Controller.Switch(33) = 1:sw33wire.RotX = 15:End Sub
Sub sw33_UnHit:Controller.Switch(33) = 0:sw33wire.RotX = 0:End Sub

Sub sw38_Hit:Controller.Switch(38) = 1:sw38wire.RotX = 15:End Sub
Sub sw38_UnHit:Controller.Switch(38) = 0:sw38wire.RotX = 0:End Sub

Sub sw46_Hit:Controller.Switch(46) = 1:sw46wire.RotX = 15:End Sub
Sub sw46_UnHit:Controller.Switch(46) = 0:sw46wire.RotX = 0:End Sub

Sub sw55_Hit:Controller.Switch(55) = 1:sw55wire.RotX = 15:End Sub
Sub sw55_UnHit:Controller.Switch(55) = 0:sw55wire.RotX = 0:End Sub

Sub sw58_Hit:Controller.Switch(58) = 1:PPL = True:PlungerL.MechPlunger = 1:PlungerR.MechPlunger = 0:sw58wire.RotX = 15:StopSound "WireRamp":End Sub
Sub sw58_UnHit:Controller.Switch(58) = 0:PPL = False:PlungerL.MechPlunger = 0:PlungerR.MechPlunger = 1:sw58wire.RotX = 0:End Sub

Sub sw85_Hit:Controller.Switch(85) = 1:sw85wire.RotX = 15:End Sub
Sub sw85_UnHit:Controller.Switch(85) = 0:sw85wire.RotX = 0:End Sub

Sub sw86_Hit:Controller.Switch(86) = 1:sw86wire.RotX = 15:End Sub
Sub sw86_UnHit:Controller.Switch(86) = 0:sw86wire.RotX = 0:End Sub

'Sub sw73_Hit:Controller.Switch(73) = 1:PlaySound "metalhit_thin":vpmTimer.AddTimer 100, "BallDropSound'":StopSound "wireramp_right":End Sub
Sub sw73_Hit:Controller.Switch(73) = 1:WireRampOff:End Sub
Sub sw73_UnHit:Controller.Switch(73) = 0:RandomSoundWireRampStop sw73 : End Sub

'Sub sw74_Hit:Controller.Switch(74) = 1:PlaySound "metalhit_thin":vpmTimer.AddTimer 100, "WireRampHitS'":PlaySound "wireramp_right":End Sub
Sub sw74_Hit:Controller.Switch(74) = 1:WireRampOff : WireRampOn False :End Sub
Sub sw74_UnHit:Controller.Switch(74) = 0:End Sub

'Sub sw75_Hit:Controller.Switch(75) = 1:PlaySound "metalhit_thin":vpmTimer.AddTimer 100, "WireRampHitS'":End Sub
Sub sw75_Hit:Controller.Switch(75) = 1: WireRampOff :End Sub
Sub sw75_UnHit:Controller.Switch(75) = 0:RandomSoundWireRampStop sw75:End Sub

Sub sw76_Hit:Controller.Switch(76) = 1:PlaySoundAt "metalhit_thin", sw76:End Sub
Sub sw76_UnHit:Controller.Switch(76) = 0:End Sub

'BlastZone
Sub sw77_Hit:Controller.Switch(77) = 1:End Sub
Sub sw77_UnHit:Controller.Switch(77) = 0:End Sub

Sub sw77TrigL_Hit:RandomSoundBlastZone:End Sub
Sub sw77TrigR_Hit:PlaySoundAtLevelStatic ("BlastZone_HitR"), BlastZoneDrop * volumedial, Activeball:End Sub

Sub startcitytrig_Hit:PlaySoundAtLevelStatic ("StartCity_Hit2"), SaucerLockSoundLevel, Activeball:End Sub

Sub sw51spinner_Spin():vpmTimer.PulseSw 51:SoundSpinner sw51spinner:End Sub

Sub WireRampHitS()
  PlaySound "wireramp_stop", 0, 0.3, -0.5
End Sub

'Targets
Sub T28a_Hit:vpmTimer.PulseSw 28:End Sub
Sub T28b_Hit:vpmTimer.PulseSw 28:End Sub
Sub T28c_Hit:vpmTimer.PulseSw 28:End Sub

Sub T34a_Hit:vpmTimer.PulseSw 34:End Sub
Sub T34b_Hit:vpmTimer.PulseSw 34:End Sub
Sub T34c_Hit:vpmTimer.PulseSw 34:End Sub

Sub T35_Hit:vpmTimer.PulseSw 35:End Sub
Sub T36_Hit:vpmTimer.PulseSw 36:End Sub

Sub T81_Hit:vpmTimer.PulseSw 81:End Sub
Sub T82_Hit:vpmTimer.PulseSw 82: T82P.RotX = T82P.RotX +5::Me.TimerEnabled = 1:End Sub
Sub T82_Timer:T82P.RotX = T82P.RotX -5:Me.TimerEnabled = 0:End Sub
Sub T83_Hit:vpmTimer.PulseSw 83:End Sub
Sub T84_Hit:vpmTimer.PulseSw 84:End Sub

'Ramps
Sub sw56_Hit:Controller.Switch(56) = 1:WireSwitchLeftRamp.RotX = 110:PlaySound "metalhit_thin":End Sub
Sub sw56_UnHit:Controller.Switch(56) = 0:WireSwitchLeftRamp.RotX = 90:End Sub
Sub sw57_Hit:Controller.Switch(57) = 1:End Sub
Sub sw57_UnHit:Controller.Switch(57) = 0:End Sub

Sub sw71_Hit:Controller.Switch(71) = 1:End Sub
Sub sw71_UnHit:Controller.Switch(71) = 0:End Sub
Sub sw72_Hit:Controller.Switch(72) = 1:WireSwitchRightRamp.RotX = 110:PlaySound "metalhit_thin":End Sub
Sub sw72_UnHit:Controller.Switch(72) = 0:WireSwitchRightRamp.RotX = 90:End Sub

Sub RRampSStart_hit: WireRampOn True :End Sub
Sub RRampSStop1_hit: WireRampOff : End Sub
Sub RWireSStart_hit: WireRampOff : WireRampOn False : End Sub
Sub RWireSStop_hit: WireRampOff: RandomSoundWireRampStop RWireSStop: End Sub

Sub LRampSStart_hit: WireRampOn True: End Sub
Sub LWireSStart1_hit: WireRampOff : WireRampOn False : End Sub

Sub FRWireSStart1_hit: WireRampOn False: End Sub
Sub LWireSStop_hit: WireRampOff : RandomSoundWireRampStop LWireSStop: End Sub
sub LWireSStart2_hit: WireRampOff : WireRampOn False : End Sub


'KickerAnimation
'Sub sw78_Hit:StartCity.AddBall 1:SoundSaucerLock:End Sub
'Sub sw78_UnHit:End Sub
Sub sw78_Hit:Controller.Switch(78) = 1:SoundSaucerLock:End Sub
Sub sw78_UnHit:Controller.Switch(78) = 0:End Sub

'******************************************************
'           TROUGH
'******************************************************

Sub sw42_Hit():Controller.Switch(42) = 1:UpdateTrough:End Sub
Sub sw42_UnHit():Controller.Switch(42) = 0:UpdateTrough:End Sub
Sub sw43_Hit():Controller.Switch(43) = 1:UpdateTrough:End Sub
Sub sw43_UnHit():Controller.Switch(43) = 0:UpdateTrough:End Sub
Sub sw44_Hit():Controller.Switch(44) = 1:UpdateTrough:End Sub
Sub sw44_UnHit():Controller.Switch(44) = 0:UpdateTrough:End Sub
Sub sw45_Hit():Controller.Switch(45) = 1:UpdateTrough:End Sub
Sub sw45_UnHit():Controller.Switch(45) = 0:UpdateTrough:End Sub

Sub UpdateTrough()
  UpdateTroughTimer.Interval = 300
  UpdateTroughTimer.Enabled = 1
End Sub

Sub UpdateTroughTimer_Timer()
  If sw42.BallCntOver = 0 Then sw43.kick 60, 9
  If sw43.BallCntOver = 0 Then sw44.kick 60, 9
  If sw44.BallCntOver = 0 Then sw45.kick 60, 9
  Me.Enabled = 0
End Sub

'******************************************************
'         DRAIN & RELEASE
'******************************************************

Sub Drain_Hit()
  RandomSoundDrain drain
  UpdateTrough
  vpmtimer.AddTimer 500, "Drain.kick 60, 20'"
End Sub

Sub SolRelease(enabled)
  If enabled Then
        vpmTimer.PulseSw 41
    sw42.kick 60, 9
    RandomSoundBallRelease sw42
  End If
End Sub

'******************************************************
' LOCKUP
'******************************************************
Dim KickerBall54

Sub sw52_Hit   : Controller.Switch(52) = 1 : End Sub
Sub sw52_UnHit : Controller.Switch(52) = 0 : End Sub
Sub sw53_Hit   : Controller.Switch(53) = 1 : End Sub
Sub sw53_UnHit : Controller.Switch(53) = 0 : End Sub
Sub sw54_Hit   : Controller.Switch(54) = 1 : set KickerBall54 = activeball : End Sub
Sub sw54_UnHit : Controller.Switch(54) = 0 : End Sub

Sub sw100_Hit
  vpmtimer.AddTimer 200, "PlaySound ""BallDropLoud"", 0, 1, AudioPan(sw100), 0, 0, 0, 0, AudioFade(sw100)'"
End Sub

Sub sw52trig_hit : PlaySoundAt "Lock_Hit2", sw52 : End Sub

Sub LockUpPin(enabled)
  If enabled then
    sol3lockup.IsDropped = 1
        PlaySound "SafeHouseKick"
  Else
    sol3lockup.IsDropped = 0
    PlaySound "LockUpPin"
    End If
End Sub

Sub SolLockKickOut(Enabled)
  If Enabled then
    If Controller.Switch(54) <> 0 Then
      KickBall KickerBall54, 180, 0, 70, 0
      SoundSaucerKick 1,sw54
      Controller.Switch(54) = 0
    End If
  End If
End Sub


Sub KickBall(kball, kangle, kvel, kvelz, kzlift)
  dim rangle
  rangle = PI * (kangle - 90) / 180
  kball.z = kball.z + kzlift
  kball.velz = kvelz
  kball.velx = cos(rangle)*kvel
  kball.vely = sin(rangle)*kvel
End Sub


'***************
' ToysAnimation
'***************

'Bulldozer

Sub UpdateDozer(aNewPos,aSpeed,aLastPos)
  DozerRotate.Enabled = 1
End Sub

Dim DozerSpeed, DozerStep

'fast Mech
DozerSpeed = 192
DozerStep = 0.4166666666667*DozerSpeed/4
Sub DozerRotate_Timer()
  If bulldozermech.position +90 > TedBar.RotX then
  TedBar.RotX = TedBar.RotX +(11/DozerStep)
  End if
  If bulldozermech.position +90 < TedBar.RotX then
  TedBar.RotX = TedBar.RotX -(11/DozerStep)
  End if
  TedBarPl.RotX = TedBar.RotX
  dozerscrews.RotX = TedBar.RotX
  If TedBar.RotX > 95 then dozerwall.isDropped = 1: Else dozerwall.isDropped = 0
End Sub

'Shake Dozer
Dim DozerPos

Sub ShakeDozer
    DozerPos = 3 '8
    DozerTimer.Enabled = 1
End Sub

Sub DozerTimer_Timer
    TedBar.RotAndTra4 = DozerPos
    TedBarPl.RotAndTra4 = DozerPos
    If DozerPos = 0 Then DozerTimer.Enabled = False:Exit Sub
    If DozerPos < 0 Then
        DozerPos = ABS(DozerPos) - 1
    Else
        DozerPos = - DozerPos + 1
    End If
End Sub

'Sound
Sub BullDozerMotor(Enabled)
  If enabled then
    PlaySound SoundFX("BulldozerMotor_Up",DofGear)
  Else
    StopSound SoundFX("BulldozerMotor_Up",DofGear)
  End If
End Sub


'TED
Sub UpdateJawTed(aNewPos,aSpeed,aLastPos)
' TedJaw.RotX = 73 +TedJawMech.position
  TedJawTimer.Enabled = 1
  DOF 201, DOFPulse
' Debug.Print "aNewPos & aSpeed & aLastPos"
End Sub


Dim JawSpeed, JawStep
JawSpeed = 90
JawStep = 0.4166666666667*jawspeed/4


Sub TedJawTimer_Timer()
  If TedJawMech.position +73 > TedJaw.RotX then
  TedJaw.RotX = TedJaw.RotX +(18/JawStep)
  End if
  If TedJawMech.position +73 < TedJaw.RotX then
  TedJaw.RotX = TedJaw.RotX -(18/JawStep)
  End if
  If TedJaw.RotX < 75 then sw48.isDropped = 1:Else sw48.isDropped = 0
' If TedJawMech.position +73 = TedJaw1.RotX then PlaySound "Knocker"
End Sub

Sub TedLidsDown(Enabled)
  If Enabled then
    TedEyelid.RotX = 18
    PlaySound "TedLids_Down"
  End if
End Sub


Sub TedLidsUp(Enabled)
  If Enabled then
     If TedEyelid.RotX = 88 then TedEyelid.RotX = 111
    PlaySound "TedLids_Up"
    Else
    TedEyelid.RotX = 88
  End if
End Sub

Sub TedEyesRight(Enabled)
  If Enabled then
    TedEyeRight.RotY = -25
    TedEyeLeft.RotY = -25
    PlaySound "TedEyes_Right"
  Else
    TedEyeRight.RotY = 0
    TedEyeLeft.RotY = 0
  End if
End Sub

Sub TedEyesLeft(Enabled)
  If Enabled then
    TedEyeRight.RotY = 25
    TedEyeLeft.RotY = 25
    PlaySound "TedEyes_Left"
  Else
    TedEyeRight.RotY = 0
    TedEyeLeft.RotY = 0
  End if
End Sub

'Sound
Sub TedMotorOn(enabled)
  If enabled then
    PlaySound SoundFX("BulldozerMotor_Down3",DofGear)
  Else
    StopSound SoundFX("BulldozerMotor_Down3",DofGear)
  End if
End Sub


'RED

Sub UpdateJawRed(aNewPos,aSpeed,aLastPos)
' If RedJaw.RotX > 105 then sw37.isDropped = 1: Else sw37.isDropped = 0
  RedJawTimer.Enabled = 1
  DOF 201, DOFPulse
End Sub

Sub RedJawTimer_Timer()
  If RedJawMech.position +90 > RedJaw.RotX then
  RedJaw.RotX = RedJaw.RotX +(18/JawStep)
  End if
  If RedJawMech.position +90 < RedJaw.RotX then
  RedJaw.RotX = RedJaw.RotX -(18/JawStep)
  End if
  If RedJaw.RotX > 105 then sw37.isDropped = 1:Else sw37.isDropped = 0
End Sub


Sub RedLidsDown(Enabled)
  If Enabled then
    RedEyelid.RotX = 18
    PlaySound "RedLids_Down"
  End if
End Sub

Sub RedLidsUp(Enabled)
  If Enabled then
    If RedEyelid.RotX = 85 then RedEyelid.RotX = 108
    PlaySound "RedLids_Up"
    Else
    RedEyelid.RotX = 85
  End if
End Sub

Sub RedEyesRight(Enabled)
  If Enabled then
    RedEyeRight.RotY = -17
    RedEyeLeft.RotY = -17
    PlaySound "RedEyes_Right"
  Else
    RedEyeRight.RotY = 0
    RedEyeLeft.RotY = 0
  End if
End Sub

Sub RedEyesLeft(Enabled)
  If Enabled then
    RedEyeRight.RotY = 17
    RedEyeLeft.RotY = 17
    PlaySound "RedEyes_Left"
  Else
    RedEyeRight.RotY = 0
    RedEyeLeft.RotY = 0
  End if
End Sub

'Sound
Sub RedMotorOn(enabled)
  If enabled then
    PlaySound SoundFX("BulldozerMotor_Down2",DofGear)
  Else
    StopSound SoundFX("BulldozerMotor_Down2",DofGear)
  End if
End Sub

'**********************************************
' Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'**********************************************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
  vpmTimer.PulseSw(62)
  RandomSoundSlingshotRight Sling1
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
  RStep = 0
  RightSlingShot.TimerEnabled = 1
    RightSlingShot.TimerInterval = 10
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  vpmTimer.PulseSw(61)
  RandomSoundSlingshotLeft Sling2
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
  LStep = 0
  LeftSlingShot.TimerEnabled = 1
  LeftSlingShot.TimerInterval  = 10
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

'******************************************
' Flipppers
'******************************************

Const ReflipAngle = 20

Sub SolLFlipper(Enabled)
  If Enabled Then
    FlipperActivate LeftFlipper, LFPress
    LF.Fire  'leftflipper.rotatetoend
      If leftflipper.currentangle < leftflipper.endangle + ReflipAngle Then
      RandomSoundReflipUpLeft LeftFlipper
    Else
      SoundFlipperUpAttackLeft LeftFlipper
      RandomSoundFlipperUpLeft LeftFlipper
    End If
  Else
    FlipperDeActivate LeftFlipper, LFPress
    LeftFlipper.RotateToStart
      If LeftFlipper.currentangle < LeftFlipper.startAngle - 5 Then
      RandomSoundFlipperDownLeft LeftFlipper
    End If
    FlipperLeftHitParm = FlipperUpSoundLevel
  End If
End Sub

Sub SolUFlipper(Enabled)
  If Enabled Then
    FlipperActivate LeftFlipper1, LFPress1
    FlipperActivate LeftFlipper2, LFPress2
    LeftFlipper1.RotateToEnd
    LeftFlipper2.RotateToEnd
    If leftflipper1.currentangle < leftflipper1.endangle + ReflipAngle Then
      RandomSoundReflipUpLeft LeftFlipper
    Else
      SoundFlipperUpAttackLeft LeftFlipper1
      RandomSoundFlipperUpLeft LeftFlipper1
    End If
  Else
    FlipperDeActivate LeftFlipper, LFPress
    FlipperDeActivate LeftFlipper1, LFPress1
    FlipperDeActivate LeftFlipper2, LFPress2
    LeftFlipper.RotateToStart
    LeftFlipper1.RotateToStart
    LeftFlipper2.RotateToStart
    If LeftFlipper.currentangle < LeftFlipper.startAngle - 5 Then
      RandomSoundFlipperDownLeft LeftFlipper
    End If
    FlipperLeftHitParm = FlipperUpSoundLevel
  End If
End Sub


Sub SolRFlipper(Enabled)
  If Enabled Then
    FlipperActivate RightFlipper, RFPress
    RF.Fire 'rightflipper.rotatetoend
    If rightflipper.currentangle > rightflipper.endangle - ReflipAngle Then
      RandomSoundReflipUpRight RightFlipper
    Else
      SoundFlipperUpAttackRight RightFlipper
      RandomSoundFlipperUpRight RightFlipper
    End If
  Else
    FlipperDeActivate RightFlipper, RFPress
    RightFlipper.RotateToStart
    If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
      RandomSoundFlipperDownRight RightFlipper
    End If
    FlipperRightHitParm = FlipperUpSoundLevel
  End If
End Sub


' Flipper collide subs
Sub LeftFlipper_Collide(parm)
  CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
  LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
  CheckLiveCatch Activeball, RightFlipper, RFCount, parm
  RightFlipperCollide parm
End Sub



'******************************************************
'****  LAMPZ by nFozzy
'******************************************************
'
' Lampz is a utility designed to manage and fade the lights and light-related objects on a table that is being driven by a ROM.
' To set up Lampz, one must populate the Lampz.MassAssign array with VPX Light objects, where the index of the MassAssign array
' corrisponds to the ROM index of the associated light. More that one Light object can be associated with a single MassAssign index (not shown in this example)
' Optionally, callbacks can be assigned for each index using the Lampz.Callback array. This is very useful for allowing 3D Insert primitives
' to be controlled by the ROM. Note, the aLvl parameter (i.e. the fading level that ranges between 0 and 1) is appended to the callback call.
'
Dim NullFader : set NullFader = new NullFadingObject
Dim Lampz : Set Lampz = New LampFader
Dim ModLampz : Set ModLampz = New DynamicLamps
LampTimer.Interval = 16
LampTimer.Enabled = 1

Sub LampTimer_Timer()
  dim x, chglamp
  chgLamp = Controller.ChangedLamps
  If Not IsEmpty(chglamp) Then
    For x = 0 To UBound(chglamp)      'nmbr = chglamp(x, 0), state = chglamp(x, 1)
      Lampz.state(chglamp(x, 0)) = chglamp(x, 1)
    next
  End If
  Lampz.Update1 'update (fading logic only)
  ModLampz.Update1
End Sub

dim FrameTime, InitFrameTime : InitFrameTime = 0
LampTimer2.Interval = -1
LampTimer2.Enabled = True
Sub LampTimer2_Timer()
  FrameTime = gametime - InitFrameTime : InitFrameTime = gametime 'Count frametime. Unused atm?
  Lampz.Update 'updates on frametime (Object updates only)
  ModLampz.Update
End Sub

Sub DisableLighting(pri, DLintensity, ByVal aLvl) 'cp's script  DLintensity = disabled lighting intesity
  if Lampz.UseFunction then aLvl = Lampz.FilterOut(aLvl)  'Callbacks don't get this filter automatically
  pri.blenddisablelighting = aLvl * DLintensity
End Sub

Sub InitLampsNF()

  'Filtering (comment out to disable)
  Lampz.Filter = "LampFilter" 'Puts all lamp intensityscale output (no callbacks) through this function before updating
  ModLampz.Filter = "LampFilter"

  'Adjust fading speeds (1 / full MS fading time)
  dim x : for x = 0 to 140 : Lampz.FadeSpeedUp(x) = 1/40 : Lampz.FadeSpeedDown(x) = 1/150 : next
  for x = 0 to 5 : ModLampz.FadeSpeedUp(x) = 1/20 : ModLampz.FadeSpeedDown(x) = 1/60 : Next
  for x = 6 to 28 : ModLampz.FadeSpeedUp(x) = 1/20 : ModLampz.FadeSpeedDown(x) = 1/60 : Next


  'Lamp Assignments
  'MassAssign is an optional way to do assignments. It'll create arrays automatically / append objects to existing arrays

  Lampz.MassAssign(11) = l11
  Lampz.MassAssign(11) = l11h
  Lampz.Callback(11) = "DisableLighting p11, 50,"
  Lampz.MassAssign(12) = l12
  Lampz.MassAssign(12) = l12h
  Lampz.Callback(12) = "DisableLighting p12, 50,"
  Lampz.MassAssign(13) = l13
  Lampz.MassAssign(13) = l13h
  Lampz.Callback(13) = "DisableLighting p13, 50,"
  Lampz.MassAssign(14) = l14
  Lampz.MassAssign(14) = l14h
  Lampz.Callback(14) = "DisableLighting p14, 50,"
  Lampz.MassAssign(15) = l15
  Lampz.MassAssign(15) = l15h
  Lampz.Callback(15) = "DisableLighting p15, 50,"
  Lampz.MassAssign(16) = l16
  Lampz.MassAssign(16) = l16h
  Lampz.Callback(16) = "DisableLighting p16, 50,"
  Lampz.MassAssign(17) = l17
  Lampz.MassAssign(17) = l17h
  Lampz.Callback(17) = "DisableLighting p17, 50,"
  Lampz.MassAssign(18) = l18
  Lampz.MassAssign(18) = l18h
  Lampz.Callback(18) = "DisableLighting p18, 50,"

  Lampz.MassAssign(21) = l21
  Lampz.MassAssign(21) = l21h
  Lampz.Callback(21) = "DisableLighting p21, 50,"
  Lampz.MassAssign(22) = l22
  Lampz.MassAssign(22) = l22h
  Lampz.Callback(22) = "DisableLighting p22, 50,"
  Lampz.MassAssign(23) = l23
  Lampz.MassAssign(23) = l23h
  Lampz.Callback(23) = "DisableLighting p23, 50,"
  Lampz.MassAssign(24) = l24
  Lampz.MassAssign(24) = l24h
  Lampz.Callback(24) = "DisableLighting p24, 50,"
  Lampz.MassAssign(25) = l25
  Lampz.MassAssign(25) = l25h
  Lampz.Callback(25) = "DisableLighting p25, 50,"
  Lampz.MassAssign(26) = l26
  Lampz.MassAssign(26) = l26h
  Lampz.Callback(26) = "DisableLighting p26, 50,"
  Lampz.MassAssign(27) = l27
  Lampz.MassAssign(27) = l27h
  Lampz.Callback(27) = "DisableLighting p27, 50,"
  Lampz.MassAssign(28) = l28
  Lampz.MassAssign(28) = l28h
  Lampz.Callback(28) = "DisableLighting p28, 50,"

  Lampz.MassAssign(31) = l31
  Lampz.MassAssign(31) = l31h
  Lampz.Callback(31) = "DisableLighting p31, 50,"
  Lampz.MassAssign(32) = l32
  Lampz.MassAssign(32) = l32h
  Lampz.Callback(32) = "DisableLighting p32, 50,"
  Lampz.MassAssign(33) = l33
  Lampz.MassAssign(33) = l33h
  Lampz.Callback(33) = "DisableLighting p33, 20,"
  Lampz.MassAssign(34) = l34
  Lampz.MassAssign(34) = l34h
  Lampz.Callback(34) = "DisableLighting p34, 1,"
  Lampz.MassAssign(35) = l35
  Lampz.MassAssign(35) = l35h
  Lampz.Callback(35) = "DisableLighting p35, 20,"
  Lampz.MassAssign(36) = l36
  Lampz.MassAssign(36) = l36h
  Lampz.Callback(36) = "DisableLighting p36, 30,"
  Lampz.MassAssign(37) = l37
  Lampz.MassAssign(37) = l37h
  Lampz.Callback(37) = "DisableLighting p37, 50,"
  Lampz.MassAssign(38) = l38
  Lampz.MassAssign(38) = l38h
  Lampz.Callback(38) = "DisableLighting p38, 1,"

  Lampz.MassAssign(41) = l41
  Lampz.MassAssign(41) = l41h
  Lampz.Callback(41) = "DisableLighting p41, 50,"
  Lampz.MassAssign(42) = l42
  Lampz.MassAssign(42) = l42h
  Lampz.Callback(42) = "DisableLighting p42, 50,"

  Lampz.MassAssign(43) = l43            'Centre Playfield Insert
  Lampz.MassAssign(43) = l43h
  Lampz.Callback(43) = "DisableLighting p43, 50,"
  Lampz.MassAssign(43) = f43            'Radio Toy Lamp
  Lampz.Callback(43) = "DisableLighting bc43, 5,"

  Lampz.MassAssign(44) = l44
  Lampz.MassAssign(44) = l44h
  Lampz.Callback(44) = "DisableLighting p44, 50,"
  Lampz.MassAssign(45) = l45
  Lampz.MassAssign(45) = l45h
  Lampz.Callback(45) = "DisableLighting p45, 30,"
  Lampz.MassAssign(46) = l46
  Lampz.MassAssign(46) = l46h
  Lampz.Callback(46) = "DisableLighting p46, 30,"
  Lampz.MassAssign(47) = l47
  Lampz.MassAssign(47) = l47h
  Lampz.Callback(47) = "DisableLighting p47, 30,"
  Lampz.MassAssign(48) = l48
  Lampz.MassAssign(48) = l48h
  Lampz.Callback(48) = "DisableLighting p48, 30,"

  Lampz.MassAssign(51) = l51
  Lampz.MassAssign(51) = l51h
  Lampz.Callback(51) = "DisableLighting p51, 0.1,"
  Lampz.MassAssign(52) = l52
  Lampz.MassAssign(52) = l52h
  Lampz.Callback(52) = "DisableLighting p52, 50,"
  Lampz.MassAssign(53) = l53
  Lampz.MassAssign(53) = l53h
  Lampz.Callback(53) = "DisableLighting p53, 50,"
  Lampz.MassAssign(54) = l54
  Lampz.MassAssign(54) = l54h
  Lampz.Callback(54) = "DisableLighting p54, 50,"
  Lampz.MassAssign(55) = l55
  Lampz.MassAssign(55) = l55h
  Lampz.Callback(55) = "DisableLighting p55, 50,"
  Lampz.MassAssign(56) = l56
  Lampz.MassAssign(56) = l56h
  Lampz.Callback(56) = "DisableLighting p56, 50,"
  Lampz.MassAssign(57) = l57
  Lampz.MassAssign(57) = l57h
  Lampz.Callback(57) = "DisableLighting p57, 10,"
  Lampz.MassAssign(58) = l58
  Lampz.MassAssign(58) = l58h
  Lampz.Callback(58) = "DisableLighting p58, 50,"

  Lampz.MassAssign(61) = l61
  Lampz.MassAssign(61) = l61h
  Lampz.Callback(61) = "DisableLighting p61, 0.1,"
  Lampz.MassAssign(62) = l62
  Lampz.MassAssign(62) = l62h
  Lampz.Callback(62) = "DisableLighting p62, 0.1,"
  Lampz.MassAssign(63) = l63
  Lampz.MassAssign(63) = l63h
  Lampz.Callback(63) = "DisableLighting p63, 0.1,"
  Lampz.MassAssign(64) = l64
  Lampz.MassAssign(64) = l64h
  Lampz.Callback(64) = "DisableLighting p64, 2,"
  Lampz.MassAssign(65) = l65
  Lampz.MassAssign(65) = l65h
  Lampz.Callback(65) = "DisableLighting p65, 50,"
  Lampz.MassAssign(66) = l66
  Lampz.MassAssign(66) = l66h
  Lampz.Callback(66) = "DisableLighting p66, 50,"
  Lampz.MassAssign(67) = l67
  Lampz.MassAssign(67) = l67h
  Lampz.Callback(67) = "DisableLighting p67, 50,"
  Lampz.MassAssign(68) = l68
  Lampz.MassAssign(68) = l68h
  Lampz.Callback(68) = "DisableLighting p68, 50,"

  Lampz.MassAssign(71) = l71
  Lampz.MassAssign(71) = l71h
  Lampz.Callback(71) = "DisableLighting p71, 50,"
  Lampz.MassAssign(72) = l72
  Lampz.MassAssign(72) = l72h
  Lampz.Callback(72) = "DisableLighting p72, 50,"
  Lampz.MassAssign(73) = l73
  Lampz.MassAssign(73) = l73h
  Lampz.Callback(73) = "DisableLighting p73, 50,"
  Lampz.MassAssign(74) = l74
  Lampz.MassAssign(74) = l74h
  Lampz.Callback(74) = "DisableLighting p74, 50,"
  Lampz.MassAssign(75) = l75
  Lampz.MassAssign(75) = l75h
  Lampz.Callback(75) = "DisableLighting p75, 50,"
  Lampz.MassAssign(76) = l76
  Lampz.MassAssign(76) = l76h
  Lampz.Callback(76) = "DisableLighting p76, 8,"
  Lampz.MassAssign(77) = f77            'Bulbs Above Playfield
  Lampz.Callback(77) = "DisableLighting bc77, 5,"
  Lampz.MassAssign(78) = l78
  Lampz.MassAssign(78) = l78h
  Lampz.Callback(78) = "DisableLighting p78, 50,"

  Lampz.MassAssign(81) = f81            'Bulbs Above Playfield
  Lampz.Callback(81) = "DisableLighting bc81, 5,"
  Lampz.MassAssign(82) = f82            'Bulbs Above Playfield
  Lampz.Callback(82) = "DisableLighting bc82, 5,"
  Lampz.MassAssign(83) = f83            'Bulbs Above Playfield
  Lampz.Callback(83) = "DisableLighting bc83, 2,"

  Lampz.MassAssign(85) = f85            'Bulbs Above Playfield
  Lampz.Callback(85) = "DisableLighting bc85, 5,"
  Lampz.Callback(85) = "DisableLighting BobsBunker, 0.015,"

' Lampz.MassAssign(84) = f84l           'Bulbs Left Ramp
' Lampz.MassAssign(84) = f84r
' Lampz.Callback(84) = "DisableLighting bc84a, 5,"
' Lampz.Callback(84) = "DisableLighting bc84b, 5,"
'
' Lampz.MassAssign(86) = f86l           'Bulbs Right Ramp
' Lampz.MassAssign(86) = f86r
' Lampz.Callback(86) = "DisableLighting bc86a, 5,"
' Lampz.Callback(86) = "DisableLighting bc86b, 5,"

  Lampz.Callback(84) = "WigWag1"
  Lampz.MassAssign(94) = f84l             'using spare lamp idxs 94-97 for blinking bulbs
  Lampz.Callback(94) = "DisableLighting bc84a, 5,"
  Lampz.MassAssign(95) = f84r
  Lampz.Callback(95) = "DisableLighting bc84b, 5,"

  Lampz.Callback(86) = "WigWag2"
  Lampz.MassAssign(96) = f86l
  Lampz.Callback(96) = "DisableLighting bc86a, 5,"
  Lampz.MassAssign(97) = f86r
  Lampz.Callback(97) = "DisableLighting bc86b, 5,"

    'VR Room Start Button & Extra Ball Button
  Lampz.MassAssign(88) = l88            'Start button
  Lampz.Callback(88) = "DisableLighting StartButton, 8,"
  Lampz.Callback(88) = "DisableLighting StartButtonInner, 8,"

  Lampz.MassAssign(87) = l87          'Extra Ball Button
  Lampz.Callback(87) = "DisableLighting ExtraballButton, 8,"
  Lampz.Callback(87) = "DisableLighting ExtraBallButtonInner, 8,"



'********************************************************

'****************************************************************
'           GI assignments
'****************************************************************

  ModLampz.Callback(0) = "GIUpdates"
  ModLampz.Callback(1) = "GIUpdates"
  ModLampz.Callback(2) = "GIUpdates"
  ModLampz.Callback(3) = "GIUpdates"
  ModLampz.Callback(4) = "GIupdates"

  ModLampz.MassAssign(0)= ColToArray(PlayfieldInsert1)  ' Playfield / Insert 1
  ModLampz.MassAssign(1)= ColToArray(PlayfieldInsert2)  ' Playfield / Insert 2
  ModLampz.MassAssign(2)= ColToArray(PlayfieldInsert3)  ' Playfield / Insert 3
  ModLampz.MassAssign(3)= ColToArray(RightPlayfield)    ' Right Playfield
  ModLampz.MassAssign(4)= ColToArray(LeftPlayfield)   ' Left Playfield

  'Turn off all lamps on startup
  lampz.Init  'This just turns state of any lamps to 1
  ModLampz.Init

  'Immediate update to turn on GI, turn off lamps
  lampz.update
  ModLampz.Update

End Sub

Function ColtoArray(aDict)  'converts a collection to an indexed array. Indexes will come out random probably.
  redim a(999)
  dim count : count = 0
  dim x  : for each x in aDict : set a(Count) = x : count = count + 1 : Next
  redim preserve a(count-1) : ColtoArray = a
End Function

Sub SetLamp(aNr, aOn)
  Lampz.state(aNr) = abs(aOn)
End Sub

Sub SetModLamp(aNr, aInput)
  ModLampz.state(aNr) = abs(aInput)/255
End Sub

'**************************************
'********* NFOZZY's WIGWAG CODE *******
'**************************************

f84l.timerinterval = 500 'adjust blink interval
f86l.timerinterval = f84l.timerinterval

sub WigWag1(f)
  if f > 0 then
    f84l.timerenabled=true : lampz.state(94)=true
  else
    f84l.timerenabled=False : Lampz.state(94)=false : Lampz.state(95)=false
  end If
End Sub

sub f84l_timer()
  Lampz.state(94) = not lampz.state(94) ' note: 'not' operator will only work if the state value is boolean
  Lampz.state(95) = not lampz.state(95)
end Sub


sub WigWag2(f)
  if f > 0 then
    f86l.timerenabled=true : Lampz.state(96)=true
  else
    f86l.timerenabled=False : Lampz.state(96)=false : Lampz.state(97)=false
  end If
End Sub

sub f86l_timer()
  Lampz.state(96) = not lampz.state(96)
  Lampz.state(97) = not lampz.state(97)
end Sub


'****************************************************************
'           GI assignments
'****************************************************************

dim ballbrightness
const ballbrightMax = 255
const ballbrightMin = 115

Set GICallback2 = GetRef("SetGI")

dim gilvl:gilvl = 1

Sub SetGI(aNr, aValue)
  'debug.print "GI nro: " & aNr & " and step: " & aValue
  ModLampz.SetGI aNr, aValue 'Redundant. Could reassign GI indexes here
End Sub

Sub GIupdates(ByVal aLvl) 'GI update odds and ends go here
  if Lampz.UseFunction then aLvl = LampFilter(aLvl) 'Callbacks don't get this filter automatically

  dim gi0lvl,gi1lvl,gi2lvl,gi3lvl,gi4lvl, ii
  if Lampz.UseFunction then   'Callbacks don't get this filter automatically
    gi0lvl = LampFilter(ModLampz.Lvl(0))
    gi1lvl = LampFilter(ModLampz.Lvl(1))
    gi2lvl = LampFilter(ModLampz.Lvl(2))
    gi3lvl = LampFilter(ModLampz.Lvl(3))
    gi4lvl = LampFilter(ModLampz.Lvl(4))
  Else
    gi0lvl = ModLampz.Lvl(0)
    gi1lvl = ModLampz.Lvl(1)
    gi2lvl = ModLampz.Lvl(2)
    gi3lvl = ModLampz.Lvl(3)
    gi4lvl = ModLampz.Lvl(4)
  end if

  FlBumperFadeTarget(1) = gi1lvl
  FlBumperFadeTarget(2) = gi0lvl
  FlBumperFadeTarget(3) = gi2lvl

  FlipperLP.blenddisablelighting = 0.5 * gi4lvl + 0.2
  FlipperL2P.blenddisablelighting = 0.2 * gi4lvl + 0.2
  FlipperRP.blenddisablelighting = 0.5 * gi3lvl + 0.2
  FlipperL3P.blenddisablelighting =0.2 * gi4lvl + 0.2
  TedBar.blenddisablelighting =0.075 * gi4lvl
' Pincab_Backglass.blenddisablelighting = 4 * gi3lvl

end sub

'**********************************************************************
'     START FLUPPER BUMPERS
'**********************************************************************

' prepare some global vars to dim/brighten objects when using day-night slider
Dim DayNightAdjust , DNA30, DNA45, DNA90
If NightDay < 10 Then
  DNA30 = 0 : DNA45 = (NightDay-10)/20 : DNA90 = 0 : DayNightAdjust = 0.4
Else
  DNA30 = (NightDay-10)/30 : DNA45 = (NightDay-10)/45 : DNA90 = (NightDay-10)/90 : DayNightAdjust = NightDay/25
End If

Dim FlBumperFadeActual(6), FlBumperFadeTarget(6), FlBumperColor(6), FlBumperTop(6), FlBumperSmallLight(6), Flbumperbiglight(6)
Dim FlBumperDisk(6), FlBumperBase(6), FlBumperBulb(6), FlBumperscrews(6), FlBumperActive(6), FlBumperHighlight(6)
Dim cnt : For cnt = 1 to 6 : FlBumperActive(cnt) = False : Next

' colors available are red, white, blue, orange, yellow, green, purple and blacklight

' FlInitBumper 1, "yellow"
' FlInitBumper 2, "red"
' FlInitBumper 3, "white"

' ### uncomment the statement below to change the color for all bumpers ###
' Dim ind : For ind = 1 to 5 : FlInitBumper ind, "green" : next

Sub FlInitBumper(nr, col)
  FlBumperActive(nr) = True
  ' store all objects in an array for use in FlFadeBumper subroutine
  FlBumperFadeActual(nr) = 1 : FlBumperFadeTarget(nr) = 1.1: FlBumperColor(nr) = col
  Set FlBumperTop(nr) = Eval("bumpertop" & nr) : FlBumperTop(nr).material = "bumpertopmat" & nr
  Set FlBumperSmallLight(nr) = Eval("bumpersmalllight" & nr) : Set Flbumperbiglight(nr) = Eval("bumperbiglight" & nr)
  Set FlBumperDisk(nr) = Eval("bumperdisk" & nr) : Set FlBumperBase(nr) = Eval("bumperbase" & nr)
  Set FlBumperBulb(nr) = Eval("bumperbulb" & nr) : FlBumperBulb(nr).material = "bumperbulbmat" & nr
  Set FlBumperscrews(nr) = Eval("bumperscrews" & nr): FlBumperscrews(nr).material = "bumperscrew" & col
  Set FlBumperHighlight(nr) = Eval("bumperhighlight" & nr)
  ' set the color for the two VPX lights
  select case col
    Case "red"
      FlBumperSmallLight(nr).color = RGB(255,4,0) : FlBumperSmallLight(nr).colorfull = RGB(255,24,0)
      FlBumperBigLight(nr).color = RGB(255,32,0) : FlBumperBigLight(nr).colorfull = RGB(255,32,0)
      FlBumperHighlight(nr).color = RGB(64,255,0)
      FlBumperSmallLight(nr).BulbModulateVsAdd = 0.98
      FlBumperSmallLight(nr).TransmissionScale = 0
    Case "white"
      FlBumperBigLight(nr).color = RGB(255,230,190) : FlBumperBigLight(nr).colorfull = RGB(255,230,190)
      FlBumperHighlight(nr).color = RGB(255,180,100) :
      FlBumperSmallLight(nr).TransmissionScale = 0
      FlBumperSmallLight(nr).BulbModulateVsAdd = 0.99
    Case "yellow"
      FlBumperSmallLight(nr).color = RGB(255,230,4) : FlBumperSmallLight(nr).colorfull = RGB(255,230,4)
      FlBumperBigLight(nr).color = RGB(255,240,50) : FlBumperBigLight(nr).colorfull = RGB(255,240,50)
      FlBumperHighlight(nr).color = RGB(255,255,220)
      FlBumperSmallLight(nr).BulbModulateVsAdd = 1
      FlBumperSmallLight(nr).TransmissionScale = 0
  end select
End Sub

Sub FlFadeBumper(nr, Z)
  FlBumperBase(nr).BlendDisableLighting = 0.5 * DayNightAdjust
' UpdateMaterial(string, float wrapLighting, float roughness, float glossyImageLerp, float thickness, float edge, float edgeAlpha, float opacity,
'               OLE_COLOR base, OLE_COLOR glossy, OLE_COLOR clearcoat, VARIANT_BOOL isMetal, VARIANT_BOOL opacityActive,
'               float elasticity, float elasticityFalloff, float friction, float scatterAngle) - updates all parameters of a material
  FlBumperDisk(nr).BlendDisableLighting = (0.5 - Z * 0.3 )* DayNightAdjust

  select case FlBumperColor(nr)

    Case "red"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1-Z, 1-Z, 1-Z, 0.9999, RGB(255, 16 - 11*Z + 16 * sin(Z*3.14),0), RGB(255,255,255), RGB(32,32,32), false, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 17 + 100 * Z / (1 + DNA30^2)
      FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 18 * Z / (1 + DNA90)
      FlBumperBulb(nr).BlendDisableLighting = 20 * DayNightAdjust + 9000 * (0.03 * Z +0.97 * Z^10)
      Flbumperbiglight(nr).intensity = 8 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 2000 * (Z^3) / (1 + DNA90)
      MaterialColor "bumpertopmat" & nr, RGB(255,20 + Z*4,8-Z*8)

    Case "white"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1-Z, 1-Z, 1-Z, 0.9999, RGB(255,230 - 100 * Z, 200 - 150 * Z), RGB(255,255,255), RGB(32,32,32), false, true, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 20 + 180 * Z / (1 + DNA30)
      FlBumperTop(nr).BlendDisableLighting = 5 * DayNightAdjust + 30 * Z
      FlBumperBulb(nr).BlendDisableLighting = 18 * DayNightAdjust + 3000 * (0.03 * Z +0.97 * Z^10)
      Flbumperbiglight(nr).intensity = 11 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 1000 * (Z^3) / (1 + DNA90)
      FlBumperSmallLight(nr).color = RGB(255,255 - 20*Z,255-65*Z) : FlBumperSmallLight(nr).colorfull = RGB(255,255 - 20*Z,255-65*Z)
      MaterialColor "bumpertopmat" & nr, RGB(255,235 - z*36,220 - Z*90)

    Case "yellow"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1-Z, 1-Z, 1-Z, 0.9999, RGB(255, 180 + 40*z, 48* Z), RGB(255,255,255), RGB(32,32,32), false, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 17 + 200 * Z / (1 + DNA30^2)
      FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 40 * Z / (1 + DNA90)
      FlBumperBulb(nr).BlendDisableLighting = 12 * DayNightAdjust + 2000 * (0.03 * Z +0.97 * Z^10)
      Flbumperbiglight(nr).intensity = 9 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 1000 * (Z^3) / (1 + DNA90)
      MaterialColor "bumpertopmat" & nr, RGB(255,200, 24 - 24 * z)

  end select
End Sub

Sub BumperTimer_Timer
  dim nr
  For nr = 1 to 6
    If FlBumperFadeActual(nr) < FlBumperFadeTarget(nr) and FlBumperActive(nr)  Then
      FlBumperFadeActual(nr) = FlBumperFadeActual(nr) + (FlBumperFadeTarget(nr) - FlBumperFadeActual(nr)) * 0.8
      If FlBumperFadeActual(nr) > 0.99 Then FlBumperFadeActual(nr) = 1 : End If
      FlFadeBumper nr, FlBumperFadeActual(nr)
    End If
    If FlBumperFadeActual(nr) > FlBumperFadeTarget(nr) and FlBumperActive(nr)  Then
      FlBumperFadeActual(nr) = FlBumperFadeActual(nr) + (FlBumperFadeTarget(nr) - FlBumperFadeActual(nr)) * 0.4 / (FlBumperFadeActual(nr) + 0.1)
      If FlBumperFadeActual(nr) < 0.01 Then FlBumperFadeActual(nr) = 0 : End If
      FlFadeBumper nr, FlBumperFadeActual(nr)
    End If
  next
End Sub
'**********************************************************************
'     END FLUPPER BUMPERS
'**********************************************************************


'****************************************
'****************************************
'         FLASHERS (MODULATED)
'****************************************
'****************************************

Sub FlashSolMod53(flLvl)  'Back White Flashers
  ObjTargetLevel(1) = flLvl / 255 : FlasherFlash1_Timer
  ObjTargetLevel(2) = flLvl / 255 : FlasherFlash2_Timer
End Sub

Sub FlashSolMod54(flLvl)  'Back Yellow Flashers
  ObjTargetLevel(3) = flLvl / 255 : FlasherFlash3_Timer
  ObjTargetLevel(4) = flLvl / 255 : FlasherFlash4_Timer
End Sub

Sub FlashSolMod55(flLvl)  'Back Red Flashers
  ObjTargetLevel(5) = flLvl / 255 : FlasherFlash5_Timer
  ObjTargetLevel(6) = flLvl / 255 : FlasherFlash6_Timer
End Sub

Sub FlashSolMod52(flLvl)  'Left Ramp Flasher
  ObjTargetLevel(7) = flLvl / 255 : FlasherFlash7_Timer
End Sub

Sub FlashSolMod57(flLvl)  'Right Ramp Flasher
  ObjTargetLevel(8) = flLvl / 255 : FlasherFlash8_Timer
End Sub

Sub FlashSolMod51(flLvl)  'Little Flipper Flasher
  ObjTargetLevel(9) = flLvl / 255 : FlasherFlash9_Timer
End Sub

Sub FlashSolMod58(flLvl)
  ObjTargetLevel(10) = flLvl / 255 : FlasherFlash10_Timer
End Sub

Sub FlashSolMod56(flLvl)
  ObjTargetLevel(11) = flLvl / 255 : FlasherFlash11_Timer
End Sub


Dim TestFlashers, TableRef, FlasherLightIntensity, FlasherFlareIntensity, FlasherBloomIntensity, FlasherOffBrightness

                ' *********************************************************************
TestFlashers = 0        ' *** set this to 1 to check position of flasher object       ***
Set TableRef = Table1     ' *** change this, if your table has another name             ***
FlasherLightIntensity = 0.2   ' *** lower this, if the VPX lights are too bright (i.e. 0.1)   *** default 0.2
FlasherFlareIntensity = 0.1   ' *** lower this, if the flares are too bright (i.e. 0.1)     *** default 0.3
FlasherBloomIntensity = 0.8   ' *** lower this, if the blooms are too bright (i.e. 0.1)     *** default 0.2
FlasherOffBrightness = 0.3    ' *** brightness of the flasher dome when switched off (range 0-2)  ***
                ' ********************************************************************

Dim ObjLevel(20), objbase(20), objlit(20), objflasher(20), objbloom(20), objlight(20), ObjTargetLevel(20)
'Dim tablewidth, tableheight : tablewidth = TableRef.width : tableheight = TableRef.height
'initialise the flasher color, you can only choose from "green", "red", "purple", "blue", "white" and "yellow"


Sub InitFlasher(nr, col)
  ' store all objects in an array for use in FlashFlasher subroutine
  Set objbase(nr) = Eval("Flasherbase" & nr) : Set objlit(nr) = Eval("Flasherlit" & nr)
  Set objflasher(nr) = Eval("Flasherflash" & nr) : Set objlight(nr) = Eval("Flasherlight" & nr)
  Set objbloom(nr) = Eval("Flasherbloom" & nr)
  ' If the flasher is parallel to the playfield, rotate the VPX flasher object for POV and place it at the correct height
  If objbase(nr).RotY = 0 Then
    objbase(nr).ObjRotZ =  atn( (tablewidth/2 - objbase(nr).x) / (objbase(nr).y - tableheight*1.1)) * 180 / 3.14159
    objflasher(nr).RotZ = objbase(nr).ObjRotZ : objflasher(nr).height = objbase(nr).z + 60
  End If
  ' set all effects to invisible and move the lit primitive at the same position and rotation as the base primitive
  objlight(nr).IntensityScale = 0 : objlit(nr).visible = 0 : objlit(nr).material = "Flashermaterial" & nr
  objlit(nr).RotX = objbase(nr).RotX : objlit(nr).RotY = objbase(nr).RotY : objlit(nr).RotZ = objbase(nr).RotZ
  objlit(nr).ObjRotX = objbase(nr).ObjRotX : objlit(nr).ObjRotY = objbase(nr).ObjRotY : objlit(nr).ObjRotZ = objbase(nr).ObjRotZ
  objlit(nr).x = objbase(nr).x : objlit(nr).y = objbase(nr).y : objlit(nr).z = objbase(nr).z
  objbase(nr).BlendDisableLighting = FlasherOffBrightness
  ' set the texture and color of all objects
  select case objbase(nr).image
    Case "dome2basewhite" : objbase(nr).image = "dome2base" & col : objlit(nr).image = "dome2lit" & col :
    Case "ronddomebasewhite" : objbase(nr).image = "ronddomebase" & col : objlit(nr).image = "ronddomelit" & col
    Case "domeearbasewhite" : objbase(nr).image = "domeearbase" & col : objlit(nr).image = "domeearlit" & col
  end select
  If TestFlashers = 0 Then objflasher(nr).imageA = "domeflashwhite" : objflasher(nr).visible = 0 : End If
  select case col
    Case "blue" :   objlight(nr).color = RGB(4,120,255) : objflasher(nr).color = RGB(200,255,255) : objbloom(nr).color = RGB(4,120,255) : objlight(nr).intensity = 5000
    Case "green" :  objlight(nr).color = RGB(12,255,4) : objflasher(nr).color = RGB(12,255,4) : objbloom(nr).color = RGB(12,255,4)
    Case "red" :    objlight(nr).color = RGB(255,32,4) : objflasher(nr).color = RGB(255,32,4) : objbloom(nr).color = RGB(255,32,4)
    Case "purple" : objlight(nr).color = RGB(230,49,255) : objflasher(nr).color = RGB(255,64,255) : objbloom(nr).color = RGB(230,49,255)
    Case "yellow" : objlight(nr).color = RGB(200,173,25) : objflasher(nr).color = RGB(255,200,50) : objbloom(nr).color = RGB(200,173,25)
    Case "white" :  objlight(nr).color = RGB(255,240,150) : objflasher(nr).color = RGB(100,86,59) : objbloom(nr).color = RGB(255,240,150)
    Case "orange" :  objlight(nr).color = RGB(255,70,0) : objflasher(nr).color = RGB(255,70,0) : objbloom(nr).color = RGB(255,70,0)
  end select
  objlight(nr).colorfull = objlight(nr).color
  If TableRef.ShowDT and ObjFlasher(nr).RotX = -45 Then
    objflasher(nr).height = objflasher(nr).height - 20 * ObjFlasher(nr).y / tableheight
    ObjFlasher(nr).y = ObjFlasher(nr).y + 10
  End If
End Sub

Sub RotateFlasher(nr, angle) : angle = ((angle + 360 - objbase(nr).ObjRotZ) mod 180)/30 : objbase(nr).showframe(angle) : objlit(nr).showframe(angle) : End Sub

Sub FlashFlasher(nr)
  If not objflasher(nr).TimerEnabled Then
    objflasher(nr).TimerEnabled = True
    objflasher(nr).visible = 1
    objlit(nr).visible = 1
    If FlasherBloomIntensity > 0 Then
      Select Case nr
        Case 1,2: objbloom(2).visible = 1
        Case 3,4: objbloom(3).visible = 1
        Case 5,6: objbloom(6).visible = 1
        Case 7: objbloom(7).visible = 1
        Case 8: objbloom(8).visible = 1
        Case 9: objbloom(9).visible = 1
        Case 10,11: objbloom(11).visible = 1
      End Select
    End If
  End If


  if round(ObjTargetLevel(nr),1) > round(ObjLevel(nr),1) Then
    ObjLevel(nr) = ObjLevel(nr) + 0.3
    if ObjLevel(nr) > 1 then ObjLevel(nr) = 1
  Elseif round(ObjTargetLevel(nr),1) < round(ObjLevel(nr),1) Then
    ObjLevel(nr) = ObjLevel(nr) * 0.85 - 0.01
    if ObjLevel(nr) < 0.01 then ObjLevel(nr) = 0
  Else
    ObjLevel(nr) = round(ObjTargetLevel(nr),1)
'   debug.print "stop timer"
    objflasher(nr).TimerEnabled = False
  end if

  objflasher(nr).opacity = 1000 *  FlasherFlareIntensity * ObjLevel(nr)^2.5
  objbloom(nr).opacity = 100 *  FlasherBloomIntensity * ObjLevel(nr)^2.5
  objlight(nr).IntensityScale = 0.5 * FlasherLightIntensity * ObjLevel(nr)^3
  objbase(nr).BlendDisableLighting =  FlasherOffBrightness + 10 * ObjLevel(nr)^3
  objlit(nr).BlendDisableLighting = 10 * ObjLevel(nr)^2
  UpdateMaterial "Flashermaterial" & nr,0,0,0,0,0,0,ObjLevel(nr),RGB(255,255,255),0,0,False,True,0,0,0,0

  If ObjLevel(nr) < 0.01 Then objflasher(nr).TimerEnabled = False : objflasher(nr).visible = 0 : objbloom(nr).visible = 0 : objlit(nr).visible = 0 : End If
End Sub

Sub FlasherFlash1_Timer() : FlashFlasher(1) : End Sub
Sub FlasherFlash2_Timer() : FlashFlasher(2) : End Sub
Sub FlasherFlash3_Timer() : FlashFlasher(3) : End Sub
Sub FlasherFlash4_Timer() : FlashFlasher(4) : End Sub
Sub FlasherFlash5_Timer() : FlashFlasher(5) : End Sub
Sub FlasherFlash6_Timer() : FlashFlasher(6) : End Sub
Sub FlasherFlash7_Timer() : FlashFlasher(7) : End Sub
Sub FlasherFlash8_Timer() : FlashFlasher(8) : End Sub
Sub FlasherFlash9_Timer() : FlashFlasher(9) : End Sub
Sub FlasherFlash10_Timer() : FlashFlasher(10) : End Sub
Sub FlasherFlash11_Timer() : FlashFlasher(11) : End Sub




'**********************************************************************
'     LAMPZ
'**********************************************************************

InitLampsNF              ' Setup lamp assignments

'**********************************************************************
'Class jungle nf
'**********************************************************************

'No-op object instead of adding more conditionals to the main loop
'It also prevents errors if empty lamp numbers are called, and it's only one object
'should be g2g?

Class NullFadingObject : Public Property Let IntensityScale(input) : : End Property : End Class

'version 0.11 - Mass Assign, Changed modulate style
'version 0.12 - Update2 (single -1 timer update) update method for core.vbs
'Version 0.12a - Filter can now be accessed via 'FilterOut'
'Version 0.12b - Changed MassAssign from a sub to an indexed property (new syntax: lampfader.MassAssign(15) = Light1 )
'Version 0.13 - No longer requires setlocale. Callback() can be assigned multiple times per index
'Note: if using multiple 'LampFader' objects, set the 'name' variable to avoid conflicts with callbacks

Class LampFader
  Public FadeSpeedDown(140), FadeSpeedUp(140)
  Private Lock(140), Loaded(140), OnOff(140)
  Public UseFunction
  Private cFilter
  Public UseCallback(140), cCallback(140)
  Public Lvl(140), Obj(140)
  Private Mult(140)
  Public FrameTime
  Private InitFrame
  Public Name

  Sub Class_Initialize()
    InitFrame = 0
    dim x : for x = 0 to uBound(OnOff)  'Set up fade speeds
      FadeSpeedDown(x) = 1/100  'fade speed down
      FadeSpeedUp(x) = 1/80   'Fade speed up
      UseFunction = False
      lvl(x) = 0
      OnOff(x) = False
      Lock(x) = True : Loaded(x) = False
      Mult(x) = 1
    Next
    Name = "LampFaderNF" 'NEEDS TO BE CHANGED IF THERE'S MULTIPLE OF THESE OBJECTS, OTHERWISE CALLBACKS WILL INTERFERE WITH EACH OTHER!!
    for x = 0 to uBound(OnOff)    'clear out empty obj
      if IsEmpty(obj(x) ) then Set Obj(x) = NullFader' : Loaded(x) = True
    Next
  End Sub

  Public Property Get Locked(idx) : Locked = Lock(idx) : End Property   ''debug.print Lampz.Locked(100) 'debug
  Public Property Get state(idx) : state = OnOff(idx) : end Property
  Public Property Let Filter(String) : Set cFilter = GetRef(String) : UseFunction = True : End Property
  Public Function FilterOut(aInput) : if UseFunction Then FilterOut = cFilter(aInput) Else FilterOut = aInput End If : End Function
  'Public Property Let Callback(idx, String) : cCallback(idx) = String : UseCallBack(idx) = True : End Property
  Public Property Let Callback(idx, String)
    UseCallBack(idx) = True
    'cCallback(idx) = String 'old execute method
    'New method: build wrapper subs using ExecuteGlobal, then call them
    cCallback(idx) = cCallback(idx) & "___" & String  'multiple strings dilineated by 3x _

    dim tmp : tmp = Split(cCallback(idx), "___")

    dim str, x : for x = 0 to uBound(tmp) 'build proc contents
      'If Not tmp(x)="" then str = str & "  " & tmp(x) & " aLVL" & "  '" & x & vbnewline  'more verbose
      If Not tmp(x)="" then str = str & tmp(x) & " aLVL:"
    Next
    'msgbox "Sub " & name & idx & "(aLvl):" & str & "End Sub"
    dim out : out = "Sub " & name & idx & "(aLvl):" & str & "End Sub"
    'if idx = 132 then msgbox out 'debug
    ExecuteGlobal Out

  End Property

  Public Property Let state(ByVal idx, input) 'Major update path
    if Input <> OnOff(idx) then  'discard redundant updates
      OnOff(idx) = input
      Lock(idx) = False
      Loaded(idx) = False
    End If
  End Property

  'Mass assign, Builds arrays where necessary
  'Sub MassAssign(aIdx, aInput)
  Public Property Let MassAssign(aIdx, aInput)
    If typename(obj(aIdx)) = "NullFadingObject" Then 'if empty, use Set
      if IsArray(aInput) then
        obj(aIdx) = aInput
      Else
        Set obj(aIdx) = aInput
      end if
    Else
      Obj(aIdx) = AppendArray(obj(aIdx), aInput)
    end if
  end Property

  Sub SetLamp(aIdx, aOn) : state(aIdx) = aOn : End Sub  'Solenoid Handler

  Public Sub TurnOnStates() 'If obj contains any light objects, set their states to 1 (Fading is our job!)
    dim debugstr
    dim idx : for idx = 0 to uBound(obj)
      if IsArray(obj(idx)) then
        'debugstr = debugstr & "array found at " & idx & "..."
        dim x, tmp : tmp = obj(idx) 'set tmp to array in order to access it
        for x = 0 to uBound(tmp)
          if typename(tmp(x)) = "Light" then DisableState tmp(x)' : debugstr = debugstr & tmp(x).name & " state'd" & vbnewline
          tmp(x).intensityscale = 0.001 ' this can prevent init stuttering
        Next
      Else
        if typename(obj(idx)) = "Light" then DisableState obj(idx)' : debugstr = debugstr & obj(idx).name & " state'd (not array)" & vbnewline
        obj(idx).intensityscale = 0.001 ' this can prevent init stuttering
      end if
    Next
    ''debug.print debugstr
  End Sub
  Private Sub DisableState(ByRef aObj) : aObj.FadeSpeedUp = 1000 : aObj.State = 1 : End Sub 'turn state to 1

  Public Sub Init() 'Just runs TurnOnStates right now
    TurnOnStates
  End Sub

  Public Property Let Modulate(aIdx, aCoef) : Mult(aIdx) = aCoef : Lock(aIdx) = False : Loaded(aIdx) = False: End Property
  Public Property Get Modulate(aIdx) : Modulate = Mult(aIdx) : End Property

  Public Sub Update1()   'Handle all boolean numeric fading. If done fading, Lock(x) = True. Update on a '1' interval Timer!
'   dim x : for x = 0 to uBound(OnOff)
'     if not Lock(x) then 'and not Loaded(x) then
'       if OnOff(x) then 'Fade Up
'         Lvl(x) = Lvl(x) + FadeSpeedUp(x)
'         if Lvl(x) >= 1 then Lvl(x) = 1 : Lock(x) = True
'       elseif Not OnOff(x) then 'fade down
'         Lvl(x) = Lvl(x) - FadeSpeedDown(x)
'         if Lvl(x) <= 0 then Lvl(x) = 0 : Lock(x) = True
'       end if
'     end if
'   Next
    FrameTime = gametime - InitFrame : InitFrame = GameTime 'Calculate frametime
    dim x : for x = 0 to uBound(OnOff)
      if not Lock(x) then 'and not Loaded(x) then
        if OnOff(x) then 'Fade Up
          Lvl(x) = Lvl(x) + FadeSpeedUp(x) * FrameTime
          if Lvl(x) >= 1 then Lvl(x) = 1 : Lock(x) = True
        elseif Not OnOff(x) then 'fade down
          Lvl(x) = Lvl(x) - FadeSpeedDown(x) * FrameTime
          if Lvl(x) <= 0 then Lvl(x) = 0 : Lock(x) = True
        end if
      end if
    Next
  End Sub

  Public Sub Update2()   'Both updates on -1 timer (Lowest latency, but less accurate fading at 60fps vsync)
    FrameTime = gametime - InitFrame : InitFrame = GameTime 'Calculate frametime
    dim x : for x = 0 to uBound(OnOff)
      if not Lock(x) then 'and not Loaded(x) then
        if OnOff(x) then 'Fade Up
          Lvl(x) = Lvl(x) + FadeSpeedUp(x) * FrameTime
          if Lvl(x) >= 1 then Lvl(x) = 1 : Lock(x) = True
        elseif Not OnOff(x) then 'fade down
          Lvl(x) = Lvl(x) - FadeSpeedDown(x) * FrameTime
          if Lvl(x) <= 0 then Lvl(x) = 0 : Lock(x) = True
        end if
      end if
    Next
    Update
  End Sub

  Public Sub Update() 'Handle object updates. Update on a -1 Timer! If done fading, loaded(x) = True
    dim x,xx : for x = 0 to uBound(OnOff)
      if not Loaded(x) then
        if IsArray(obj(x) ) Then  'if array
          If UseFunction then
            for each xx in obj(x) : xx.IntensityScale = cFilter(Lvl(x)*Mult(x)) : Next
          Else
            for each xx in obj(x) : xx.IntensityScale = Lvl(x)*Mult(x) : Next
          End If
        else            'if single lamp or flasher
          If UseFunction then
            obj(x).Intensityscale = cFilter(Lvl(x)*Mult(x))
          Else
            obj(x).Intensityscale = Lvl(x)
          End If
        end if
        if TypeName(lvl(x)) <> "Double" and typename(lvl(x)) <> "Integer" then msgbox "uhh " & 2 & " = " & lvl(x)
        'If UseCallBack(x) then execute cCallback(x) & " " & (Lvl(x)) 'Callback
        If UseCallBack(x) then Proc name & x,Lvl(x)*mult(x) 'Proc
        If Lock(x) Then
          if Lvl(x) = 1 or Lvl(x) = 0 then Loaded(x) = True 'finished fading
        end if
      end if
    Next
  End Sub
End Class


Class DynamicLamps 'Lamps that fade up and down. GI and Flasher handling
  Public Loaded(50), FadeSpeedDown(50), FadeSpeedUp(50)
  Private Lock(50), SolModValue(50)
  Private UseCallback(50), cCallback(50)
  Public Lvl(50)
  Public Obj(50)
  Private UseFunction, cFilter
  private Mult(50)
  Public Name

  Public FrameTime
  Private InitFrame

  Private Sub Class_Initialize()
    InitFrame = 0
    dim x : for x = 0 to uBound(Obj)
      FadeSpeedup(x) = 0.01
      FadeSpeedDown(x) = 0.01
      lvl(x) = 0.0001 : SolModValue(x) = 0
      Lock(x) = True : Loaded(x) = False
      mult(x) = 1
      Name = "DynamicFaderNF" 'NEEDS TO BE CHANGED IF THERE'S MULTIPLE OBJECTS, OTHERWISE CALLBACKS WILL INTERFERE WITH EACH OTHER!!
      if IsEmpty(obj(x) ) then Set Obj(x) = NullFader' : Loaded(x) = True
    next
  End Sub

  Public Property Get Locked(idx) : Locked = Lock(idx) : End Property
  'Public Property Let Callback(idx, String) : cCallback(idx) = String : UseCallBack(idx) = True : End Property
  Public Property Let Filter(String) : Set cFilter = GetRef(String) : UseFunction = True : End Property
  Public Function FilterOut(aInput) : if UseFunction Then FilterOut = cFilter(aInput) Else FilterOut = aInput End If : End Function

  Public Property Let Callback(idx, String)
    UseCallBack(idx) = True
    'cCallback(idx) = String 'old execute method
    'New method: build wrapper subs using ExecuteGlobal, then call them
    cCallback(idx) = cCallback(idx) & "___" & String  'multiple strings dilineated by 3x _

    dim tmp : tmp = Split(cCallback(idx), "___")

    dim str, x : for x = 0 to uBound(tmp) 'build proc contents
      'debugstr = debugstr & x & "=" & tmp(x) & vbnewline
      'If Not tmp(x)="" then str = str & "  " & tmp(x) & " aLVL" & "  '" & x & vbnewline  'more verbose
      If Not tmp(x)="" then str = str & tmp(x) & " aLVL:"
    Next

    dim out : out = "Sub " & name & idx & "(aLvl):" & str & "End Sub"
    'if idx = 132 then msgbox out 'debug
    ExecuteGlobal Out
  End Property


  Public Property Let State(idx,Value)
    'If Value = SolModValue(idx) Then Exit Property ' Discard redundant updates
    If Value <> SolModValue(idx) Then ' Discard redundant updates
      SolModValue(idx) = Value
      Lock(idx) = False : Loaded(idx) = False
    End If
  End Property

  Public Property Get state(idx) : state = SolModValue(idx) : end Property

  'Mass assign, Builds arrays where necessary
  'Sub MassAssign(aIdx, aInput)
  Public Property Let MassAssign(aIdx, aInput)
    If typename(obj(aIdx)) = "NullFadingObject" Then 'if empty, use Set
      if IsArray(aInput) then
        obj(aIdx) = aInput
      Else
        Set obj(aIdx) = aInput
      end if
    Else
      Obj(aIdx) = AppendArray(obj(aIdx), aInput)
    end if
  end Property

  'solcallback (solmodcallback) handler
  Sub SetLamp(aIdx, aInput) : state(aIdx) = aInput : End Sub  '0->1 Input
  Sub SetModLamp(aIdx, aInput) : state(aIdx) = aInput/255 : End Sub '0->255 Input
  Sub SetGI(aIdx, ByVal aInput) : if aInput = 8 then aInput = 7 end if : state(aIdx) = aInput/7 : End Sub '0->8 WPC GI input

  Public Sub TurnOnStates() 'If obj contains any light objects, set their states to 1 (Fading is our job!)
    dim debugstr
    dim idx : for idx = 0 to uBound(obj)
      if IsArray(obj(idx)) then
        'debugstr = debugstr & "array found at " & idx & "..."
        dim x, tmp : tmp = obj(idx) 'set tmp to array in order to access it
        for x = 0 to uBound(tmp)
          if typename(tmp(x)) = "Light" then DisableState tmp(x) ': debugstr = debugstr & tmp(x).name & " state'd" & vbnewline

        Next
      Else
        if typename(obj(idx)) = "Light" then DisableState obj(idx) ': debugstr = debugstr & obj(idx).name & " state'd (not array)" & vbnewline

      end if
    Next
    'debug.print debugstr
  End Sub

  Private Sub DisableState(ByRef aObj) : aObj.FadeSpeedUp = 1000 : aObj.State = 1 : End Sub 'turn state to 1

  Public Sub Init() 'just call turnonstates for now
    TurnOnStates
  End Sub

  Public Property Let Modulate(aIdx, aCoef) : Mult(aIdx) = aCoef : Lock(aIdx) = False : Loaded(aIdx) = False: End Property
  Public Property Get Modulate(aIdx) : Modulate = Mult(aIdx) : End Property

  Public Sub Update1()   'Handle all numeric fading. If done fading, Lock(x) = True
    'dim stringer
'   dim x : for x = 0 to uBound(Lvl)
'     'stringer = "Locked @ " & SolModValue(x)
'     if not Lock(x) then 'and not Loaded(x) then
'       If lvl(x) < SolModValue(x) then '+
'         'stringer = "Fading Up " & lvl(x) & " + " & FadeSpeedUp(x)
'         Lvl(x) = Lvl(x) + FadeSpeedUp(x)
'         if Lvl(x) >= SolModValue(x) then Lvl(x) = SolModValue(x) : Lock(x) = True
'       ElseIf Lvl(x) > SolModValue(x) Then '-
'         Lvl(x) = Lvl(x) - FadeSpeedDown(x)
'         'stringer = "Fading Down " & lvl(x) & " - " & FadeSpeedDown(x)
'         if Lvl(x) <= SolModValue(x) then Lvl(x) = SolModValue(x) : Lock(x) = True
'       End If
'     end if
'   Next
    FrameTime = gametime - InitFrame : InitFrame = GameTime 'Calculate frametime
    dim x : for x = 0 to uBound(Lvl)
      if not Lock(x) then 'and not Loaded(x) then
        If lvl(x) < SolModValue(x) then '+
          Lvl(x) = Lvl(x) + FadeSpeedUp(x) * FrameTime
          if Lvl(x) >= SolModValue(x) then Lvl(x) = SolModValue(x) : Lock(x) = True
        ElseIf Lvl(x) > SolModValue(x) Then '-
          Lvl(x) = Lvl(x) - FadeSpeedDown(x) * FrameTime
          if Lvl(x) <= SolModValue(x) then Lvl(x) = SolModValue(x) : Lock(x) = True
        End If
      end if
    Next
    'tbF.text = stringer
  End Sub

  Public Sub Update2()   'Both updates on -1 timer (Lowest latency, but less accurate fading at 60fps vsync)
    FrameTime = gametime - InitFrame : InitFrame = GameTime 'Calculate frametime
    dim x : for x = 0 to uBound(Lvl)
      if not Lock(x) then 'and not Loaded(x) then
        If lvl(x) < SolModValue(x) then '+
          Lvl(x) = Lvl(x) + FadeSpeedUp(x) * FrameTime
          if Lvl(x) >= SolModValue(x) then Lvl(x) = SolModValue(x) : Lock(x) = True
        ElseIf Lvl(x) > SolModValue(x) Then '-
          Lvl(x) = Lvl(x) - FadeSpeedDown(x) * FrameTime
          if Lvl(x) <= SolModValue(x) then Lvl(x) = SolModValue(x) : Lock(x) = True
        End If
      end if
    Next
    Update
  End Sub

  Public Sub Update() 'Handle object updates. Update on a -1 Timer! If done fading, loaded(x) = True
    dim x,xx
    for x = 0 to uBound(Lvl)
      if not Loaded(x) then
        if IsArray(obj(x) ) Then  'if array
          If UseFunction then
            for each xx in obj(x) : xx.IntensityScale = cFilter(abs(Lvl(x))*mult(x)) : Next
          Else
            for each xx in obj(x) : xx.IntensityScale = Lvl(x)*mult(x) : Next
          End If
        else            'if single lamp or flasher
          If UseFunction then
            obj(x).Intensityscale = cFilter(abs(Lvl(x))*mult(x))
          Else
            obj(x).Intensityscale = Lvl(x)*mult(x)
          End If
        end if
        'If UseCallBack(x) then execute cCallback(x) & " " & (Lvl(x)*mult(x)) 'Callback
        If UseCallBack(x) then Proc name & x,Lvl(x)*mult(x) 'Proc
        If Lock(x) Then
          Loaded(x) = True
        end if
      end if
    Next
  End Sub
End Class


'Lamp Filter
Function LampFilter(aLvl)
  LampFilter = aLvl^1.6 'exponential curve?
End Function


'Helper functions
Sub Proc(string, Callback)  'proc using a string and one argument
  'On Error Resume Next
  dim p : Set P = GetRef(String)
  P Callback
  If err.number = 13 then  msgbox "Proc error! No such procedure: " & vbnewline & string
  if err.number = 424 then msgbox "Proc error! No such Object"
End Sub

Function AppendArray(ByVal aArray, aInput)  'append one value, object, or Array onto the end of a 1 dimensional array
  if IsArray(aInput) then 'Input is an array...
    dim tmp : tmp = aArray
    If not IsArray(aArray) Then 'if not array, create an array
      tmp = aInput
    Else            'Append existing array with aInput array
      Redim Preserve tmp(uBound(aArray) + uBound(aInput)+1) 'If existing array, increase bounds by uBound of incoming array
      dim x : for x = 0 to uBound(aInput)
        if isObject(aInput(x)) then
          Set tmp(x+uBound(aArray)+1 ) = aInput(x)
        Else
          tmp(x+uBound(aArray)+1 ) = aInput(x)
        End If
      Next
      AppendArray = tmp  'return new array
    End If
  Else 'Input is NOT an array...
    If not IsArray(aArray) Then 'if not array, create an array
      aArray = Array(aArray, aInput)
    Else
      Redim Preserve aArray(uBound(aArray)+1) 'If array, increase bounds by 1
      if isObject(aInput) then
        Set aArray(uBound(aArray)) = aInput
      Else
        aArray(uBound(aArray)) = aInput
      End If
    End If
    AppendArray = aArray 'return new array
  End If
End Function

'******************************************************
'****  END LAMPZ
'******************************************************



'******************************************************
' VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************

Const TargetBouncerEnabled = 1    '0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.9   'Level of bounces. Recommmended value of 0.7

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

' Add targets or posts to the TargetBounce collection if you want to activate the targetbouncer code from them
Sub TargetBounce_Hit(idx)
  TargetBouncer activeball, 1
End Sub

'******************************************************
'****  GNEREAL ADVICE ON PHYSICS
'******************************************************
'
' It's advised that flipper corrections, dampeners, and general physics settings should all be updated per these
' examples as all of these improvements work together to provide a realistic physics simulation.
'
' Tutorial videos provided by Bord
' Flippers:   https://www.youtube.com/watch?v=FWvM9_CdVHw
' Dampeners:  https://www.youtube.com/watch?v=tqsxx48C6Pg
' Physics:    https://www.youtube.com/watch?v=UcRMG-2svvE
'
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
'****  FLIPPER CORRECTIONS by nFozzy
'******************************************************
'
' There are several steps for taking advantage of nFozzy's flipper solution.  At a high level we?ll need the following:
' 1. flippers with specific physics settings
' 2. custom triggers for each flipper (TriggerLF, TriggerRF)
' 3. an object or point to tell the script where the tip of the flipper is at rest (EndPointLp, EndPointRp)
' 4. and, special scripting
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
' | EOS Torque         | 0.3            | 0.3                   | 0.275                  | 0.275              |
' | EOS Torque Angle   | 4              | 4                     | 6                      | 6                  |


'******************************************************
' Flippers Polarity (Select appropriate sub based on era)
'******************************************************

dim LF : Set LF = New FlipperPolarity
dim RF : Set RF = New FlipperPolarity

InitPolarity

''*******************************************
'' Early 90's and after

Sub InitPolarity()
        dim x, a : a = Array(LF, RF)
        for each x in a
                x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1        'disabled
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

        addpt "Velocity", 0, 0,         1
        addpt "Velocity", 1, 0.16, 1.06
        addpt "Velocity", 2, 0.41,         1.05
        addpt "Velocity", 3, 0.53,         1'0.982
        addpt "Velocity", 4, 0.702, 0.968
        addpt "Velocity", 5, 0.95,  0.968
        addpt "Velocity", 6, 1.03,         0.945

        LF.Object = LeftFlipper
        LF.EndPoint = EndPointLp
        RF.Object = RightFlipper
        RF.EndPoint = EndPointRp
End Sub


' Flipper trigger hit subs
Sub TriggerLF_Hit() : LF.Addball activeball : End Sub
Sub TriggerLF_UnHit() : LF.PolarityCorrect activeball : End Sub
Sub TriggerRF_Hit() : RF.Addball activeball : End Sub
Sub TriggerRF_UnHit() : RF.PolarityCorrect activeball : End Sub

'******************************************************
'  FLIPPER CORRECTION FUNCTIONS
'******************************************************

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


'******************************************************
'  SLINGSHOT CORRECTION FUNCTIONS
'******************************************************
' To add these slingshot corrections:
'   - On the table, add the endpoint primitives that define the two ends of the Slingshot
' - Initialize the SlingshotCorrection objects in InitSlingCorrection
'   - Call the .VelocityCorrect methods from the respective _Slingshot event sub

dim LS : Set LS = New SlingshotCorrection
dim RS : Set RS = New SlingshotCorrection

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
  AddSlingsPt 0, 0.00,  -4
  AddSlingsPt 1, 0.45,  -7
  AddSlingsPt 2, 0.48,  0
  AddSlingsPt 3, 0.52,  0
  AddSlingsPt 4, 0.55,  7
  AddSlingsPt 5, 1.00,  4

End Sub


Sub AddSlingsPt(idx, aX, aY)        'debugger wrapper for adjusting flipper script in-game
  dim a : a = Array(LS, RS)
  dim x : for each x in a
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
'    dim rx, ry
'    rx = x*dCos(angle) - y*dSin(angle)
'    ry = x*dSin(angle) + y*dCos(angle)
'    RotPoint = Array(rx,ry)
'End Function

Class SlingshotCorrection
  Public DebugOn, Enabled
  private Slingshot, SlingX1, SlingX2, SlingY1, SlingY2

  Public ModIn, ModOut
  Private Sub Class_Initialize : redim ModIn(0) : redim Modout(0): Enabled = True : End Sub

  Public Property let Object(aInput) : Set Slingshot = aInput : End Property
  Public Property Let EndPoint1(aInput) : SlingX1 = aInput.x: SlingY1 = aInput.y: End Property
  Public Property Let EndPoint2(aInput) : SlingX2 = aInput.x: SlingY2 = aInput.y: End Property

  Public Sub AddPoint(aIdx, aX, aY)
    ShuffleArrays ModIn, ModOut, 1 : ModIn(aIDX) = aX : ModOut(aIDX) = aY : ShuffleArrays ModIn, ModOut, 0
    If gametime > 100 then Report
  End Sub

  Public Sub Report()         'debug, reports all coords in tbPL.text
    If not debugOn then exit sub
    dim a1, a2 : a1 = ModIn : a2 = ModOut
    dim str, x : for x = 0 to uBound(a1) : str = str & x & ": " & round(a1(x),4) & ", " & round(a2(x),4) & vbnewline : next
    TBPout.text = str
  End Sub


  Public Sub VelocityCorrect(aBall)
    dim BallPos, XL, XR, YL, YR

    'Assign right and left end points
    If SlingX1 < SlingX2 Then
      XL = SlingX1 : YL = SlingY1 : XR = SlingX2 : YR = SlingY2
    Else
      XL = SlingX2 : YL = SlingY2 : XR = SlingX1 : YR = SlingY1
    End If

    'Find BallPos = % on Slingshot
    If Not IsEmpty(aBall.id) Then
      If ABS(XR-XL) > ABS(YR-YL) Then
        BallPos = PSlope(aBall.x, XL, 0, XR, 1)
      Else
        BallPos = PSlope(aBall.y, YL, 0, YR, 1)
      End If
      If BallPos < 0 Then BallPos = 0
      If BallPos > 1 Then BallPos = 1
    End If

    'Velocity angle correction
    If not IsEmpty(ModIn(0) ) then
      Dim Angle, RotVxVy
      Angle = LinearEnvelope(BallPos, ModIn, ModOut)
      'debug.print " BallPos=" & BallPos &" Angle=" & Angle
      'debug.print " BEFORE: aBall.Velx=" & aBall.Velx &" aBall.Vely" & aBall.Vely
      RotVxVy = RotPoint(aBall.Velx,aBall.Vely,Angle)
      If Enabled then aBall.Velx = RotVxVy(0)
      If Enabled then aBall.Vely = RotVxVy(1)
      'debug.print " AFTER: aBall.Velx=" & aBall.Velx &" aBall.Vely" & aBall.Vely
      'debug.print " "
    End If
  End Sub

End Class



'******************************************************
'  FLIPPER POLARITY. RUBBER DAMPENER, AND SLINGSHOT CORRECTION SUPPORTING FUNCTIONS
'******************************************************


Sub AddPt(aStr, idx, aX, aY)        'debugger wrapper for adjusting flipper script in-game
  dim a : a = Array(LF, RF)
  dim x : for each x in a
    x.addpoint aStr, idx, aX, aY
  Next
End Sub


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

sub RightFlipper_Timer()
  FlipperTricks LeftFlipper, LFPress, LFCount, LFEndAngle, LFState
  FlipperTricks RightFlipper, RFPress, RFCount, RFEndAngle, RFState
  FlipperNudge RightFlipper, RFEndAngle, RFEOSNudge, LeftFlipper, LFEndAngle
  FlipperNudge LeftFlipper, LFEndAngle, LFEOSNudge,  RightFlipper, RFEndAngle
end sub

Dim LFEOSNudge, RFEOSNudge

Sub FlipperNudge(Flipper1, Endangle1, EOSNudge1, Flipper2, EndAngle2)
  Dim b

  If Flipper1.currentangle = Endangle1 and EOSNudge1 <> 1 Then
    EOSNudge1 = 1
    'debug.print Flipper1.currentangle &" = "& Endangle1 &"--"& Flipper2.currentangle &" = "& EndAngle2
    If Flipper2.currentangle = EndAngle2 Then
      For b = 0 to Ubound(gBOT)
        If FlipperTrigger(gBOT(b).x, gBOT(b).y, Flipper1) Then
          'Debug.Print "ball in flip1. exit"
          exit Sub
        end If
      Next
      For b = 0 to Ubound(gBOT)
        If FlipperTrigger(gBOT(b).x, gBOT(b).y, Flipper2) Then
          gBOT(b).velx = gBOT(b).velx / 1.3
          gBOT(b).vely = gBOT(b).vely - 0.5
        end If
      Next
    End If
  Else
    If Abs(Flipper1.currentangle) > Abs(EndAngle1) + 30 then EOSNudge1 = 0
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


'*************************************************
'  Check ball distance from Flipper for Rem
'*************************************************

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


'*************************************************
'  End - Check ball distance from Flipper for Rem
'*************************************************

dim LFPress, LFPress1, LFPress2, RFPress, LFCount, RFCount
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
'Const EOSTnew = 1 'EM's to late 80's
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

Const LiveCatch = 16
Const LiveElasticity = 0.45
Const SOSEM = 0.815
'Const EOSReturn = 0.055  'EM's
'Const EOSReturn = 0.045  'late 70's to mid 80's
'Const EOSReturn = 0.035  'mid 80's to early 90's
Const EOSReturn = 0.025  'mid 90's and later

Const SOSReturn = 0.025

LFEndAngle = Leftflipper.endangle
RFEndAngle = RightFlipper.endangle

Sub FlipperActivate(Flipper, FlipperPress)
  FlipperPress = 1
  Flipper.Elasticity = FElasticity
  Flipper.return = FReturn

  Flipper.eostorque = EOST
  Flipper.eostorqueangle = EOSA
End Sub

Sub FlipperDeactivate(Flipper, FlipperPress)
  FlipperPress = 0
  Flipper.eostorqueangle = EOSA
  Flipper.eostorque = EOST*EOSReturn/FReturn

  If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 0.1 Then
    Dim b

    For b = 0 to UBound(gBOT)
      If Distance(gBOT(b).x, gBOT(b).y, Flipper.x, Flipper.y) < 55 Then 'check for cradle
        If gBOT(b).vely >= -0.4 Then gBOT(b).vely = -0.4
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
      Flipper.Return = SOSReturn 'set new return strength when flipper is down
      FCount = 0
      FState = 1
    End If
  ElseIf Abs(Flipper.currentangle) <= Abs(Flipper.endangle) and FlipperPress = 1 then
    if FCount = 0 Then FCount = GameTime

    If FState <> 2 Then
      Flipper.return = FReturn 'insure original return strength is used when flipping - could get reset above after keypress before flipper moves
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

    If LiveCatchBounce = 0 and ball.velx * Dir > 0 Then ball.velx = 0
    ball.vely = LiveCatchBounce * (32 / LiveCatch) ' Multiplier for inaccuracy bounce
    ball.angmomx= 0
    ball.angmomy= 0
    ball.angmomz= 0
    Else
        If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 1 Then FlippersD.Dampenf Activeball, parm
  End If
End Sub


'******************************************************
'****  END FLIPPER CORRECTIONS
'******************************************************






'******************************************************
'****  PHYSICS DAMPENERS
'******************************************************
'
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



'******************************************************
'  TRACK ALL BALL VELOCITIES
'  FOR RUBBER DAMPENER AND DROP TARGETS
'******************************************************

dim cor : set cor = New CoRTracker

Class CoRTracker
  public ballvel, ballvelx, ballvely

  Private Sub Class_Initialize : redim ballvel(0) : redim ballvelx(0): redim ballvely(0) : End Sub

  Public Sub Update() 'tracks in-ball-velocity
    dim str, b, highestID
    dim BOT : BOT = getballs

    for each b in BOT
      if b.id >= HighestID then highestID = b.id
    Next

    if uBound(ballvel) < highestID then redim ballvel(highestID)  'set bounds
    if uBound(ballvelx) < highestID then redim ballvelx(highestID)  'set bounds
    if uBound(ballvely) < highestID then redim ballvely(highestID)  'set bounds

    for each b in BOT
      ballvel(b.id) = BallSpeed(b)
      ballvelx(b.id) = b.velx
      ballvely(b.id) = b.vely
    Next
  End Sub
End Class

' Note, cor.update must be called in a 10 ms timer. The example table uses the GameTimer for this purpose, but sometimes a dedicated timer call RDampen is used.
'
'Sub RDampen_Timer
' Cor.Update
'End Sub



'******************************************************
'****  END PHYSICS DAMPENERS
'******************************************************




'******************************************************
'****  BALL ROLLING AND DROP SOUNDS
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
  Dim b

  ' stop the sound of deleted balls
  For b = UBound(gBOT) + 1 to tnob
    ' Comment the next line if you are not implementing Dyanmic Ball Shadows
    If AmbientBallShadowOn = 0 Then BallShadowA(b).visible = 0
    rolling(b) = False
    StopSound("BallRoll_" & b)
  Next

  ' exit the sub if no balls on the table
  If UBound(gBOT) = -1 Then Exit Sub

  ' play the rolling sound for each ball

  For b = 0 to UBound(gBOT)
    If BallVel(gBOT(b)) > 1 AND gBOT(b).z < 30 Then
      rolling(b) = True
      PlaySound ("BallRoll_" & b), -1, VolPlayfieldRoll(gBOT(b)) * BallRollVolume , AudioPan(gBOT(b)), 0, PitchPlayfieldRoll(gBOT(b)), 1, 0, AudioFade(gBOT(b))

    Else
      If rolling(b) = True Then
        StopSound("BallRoll_" & b)
        rolling(b) = False
      End If
    End If

    ' Ball Drop Sounds
    If gBOT(b).VelZ < -1 and gBOT(b).z < 55 and gBOT(b).z > 27 Then 'height adjust for ball drop sounds
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

    ' "Static" Ball Shadows
    ' Comment the next If block, if you are not implementing the Dyanmic Ball Shadows
    If AmbientBallShadowOn = 0 Then
      If gBOT(b).Z > 30 Then
        BallShadowA(b).height=gBOT(b).z - BallSize/4    'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping
      Else
        BallShadowA(b).height=gBOT(b).z - BallSize/2 + 5
      End If
      BallShadowA(b).Y = gBOT(b).Y + Ballsize/5 + fovY
      BallShadowA(b).X = gBOT(b).X
      BallShadowA(b).visible = 1
    End If
  Next
End Sub


'******************************************************
'****  END BALL ROLLING AND DROP SOUNDS
'******************************************************


'******************************************************
'****  FLEEP MECHANICAL SOUNDS
'******************************************************

' This part in the script is an entire block that is dedicated to the physics sound system.
' Various scripts and sounds that may be pretty generic and could suit other WPC systems, but the most are tailored specifically for the TOM table

' Many of the sounds in this package can be added by creating collections and adding the appropriate objects to those collections.
' Create the following new collections:
'   Metals (all metal objects, metal walls, metal posts, metal wire guides)
'   Apron (the apron walls and plunger wall)
'   Walls (all wood or plastic walls)
'   Rollovers (wire rollover triggers, star triggers, or button triggers)
'   Targets (standup or drop targets, these are hit sounds only ... you will want to add separate dropping sounds for drop targets)
'   Gates (plate gates)
'   GatesWire (wire gates)
'   Rubbers (all rubbers including posts, sleeves, pegs, and bands)
' When creating the collections, make sure "Fire events for this collection" is checked.
' You'll also need to make sure "Has Hit Event" is checked for each object placed in these collections (not necessary for gates and triggers).
' Once the collections and objects are added, the save, close, and restart VPX.
'
' Tutorial vides by Apophis
' Part 1:   https://youtu.be/PbE2kNiam3g
' Part 2:   https://youtu.be/B5cm1Y8wQsk
' Part 3:   https://youtu.be/eLhWyuYOyGg

Dim BIPL : BIPL = False       'Ball in plunger lane

'/////////////////////////////////  CARTRIDGES  /////////////////////////////////
'
'//  /Fleep:
'//  A cheap version of cartridges use. Added just for the shaker sounds
'//  Mechanical sounds naming convention: <CARTRIDGE>_<Soundset_Name>
'//
'//  Cartridge name is composed using the following convention:
'//  <TABLE MANUFACTURER ABBREVIATION>_<TABLE NAME ABBREVIATION>_<SOUNDSET REVISION NUMBER>
'//
Const Cartridge_Table_Specifics     = "SY_TNA_REV02" 'Spooky Total Nuclear Annihilation Cartridge REV02

'//  Cartridges wouldn't have been possible without audio recordings provided by:
'//  Spooky Total Nuclear Annihilation - WildDogArcade, Ed and Gary



'///////////////////////////////  SOUNDS PARAMETERS  //////////////////////////////
Dim GlobalSoundLevel, CoinSoundLevel, PlungerReleaseSoundLevel, PlungerPullSoundLevel, NudgeLeftSoundLevel
Dim NudgeRightSoundLevel, NudgeCenterSoundLevel, StartButtonSoundLevel, RollingSoundFactor
Dim ShakerSoundLevel

CoinSoundLevel = 1                            'volume level; range [0, 1]
NudgeLeftSoundLevel = 1                         'volume level; range [0, 1]
NudgeRightSoundLevel = 1                        'volume level; range [0, 1]
NudgeCenterSoundLevel = 1                       'volume level; range [0, 1]
StartButtonSoundLevel = 0.1                       'volume level; range [0, 1]
PlungerReleaseSoundLevel = 0.8 '1 wjr                     'volume level; range [0, 1]
PlungerPullSoundLevel = 1                       'volume level; range [0, 1]
RollingSoundFactor = 1.1/5
'ShakerSoundLevel = 1                         'volume level; range [0, 1]
ShakerSoundLevel = ShakerSoundVolume                  'volume level gets const value

'////////////////////////////////  SOUND HELPERS  ///////////////////////////////
Dim SoundOn : SoundOn = 1
Dim SoundOff : SoundOff = 0


'///////////////////////-----Solenoids, Kickers and Flash Relays-----///////////////////////
Dim FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel, FlipperUpAttackLeftSoundLevel, FlipperUpAttackRightSoundLevel
Dim FlipperUpSoundLevel, FlipperDownSoundLevel, FlipperLeftHitParm, FlipperRightHitParm
Dim SlingshotSoundLevel, BumperSoundFactor, KnockerSoundLevel

FlipperUpAttackMinimumSoundLevel = 0.010                      'volume level; range [0, 1]
FlipperUpAttackMaximumSoundLevel = 0.635                'volume level; range [0, 1]
FlipperUpSoundLevel = 1.0                                   'volume level; range [0, 1]
FlipperDownSoundLevel = 0.45                                  'volume level; range [0, 1]
FlipperLeftHitParm = FlipperUpSoundLevel                'sound helper; not configurable
FlipperRightHitParm = FlipperUpSoundLevel               'sound helper; not configurable
SlingshotSoundLevel = 0.95                        'volume level; range [0, 1]
BumperSoundFactor = 4.25                        'volume multiplier; must not be zero
KnockerSoundLevel = 1                           'volume level; range [0, 1]

'///////////////////////-----Ball Drops, Bumps and Collisions-----///////////////////////
Dim RubberStrongSoundFactor, RubberWeakSoundFactor, RubberFlipperSoundFactor,BallWithBallCollisionSoundFactor
Dim BallBouncePlayfieldSoftFactor, BallBouncePlayfieldHardFactor, PlasticRampDropToPlayfieldSoundLevel, WireRampDropToPlayfieldSoundLevel, DelayedBallDropOnPlayfieldSoundLevel
Dim WallImpactSoundFactor, MetalImpactSoundFactor, SubwaySoundLevel, SubwayEntrySoundLevel, ScoopEntrySoundLevel
Dim SaucerLockSoundLevel, SaucerKickSoundLevel, BlastZoneDrop

BallWithBallCollisionSoundFactor = 3.2                  'volume multiplier; must not be zero
RubberStrongSoundFactor = 0.055/5                     'volume multiplier; must not be zero
RubberWeakSoundFactor = 0.075/5                     'volume multiplier; must not be zero
RubberFlipperSoundFactor = 0.475/5                    'volume multiplier; must not be zero
BallBouncePlayfieldSoftFactor = 0.025                 'volume multiplier; must not be zero
BallBouncePlayfieldHardFactor = 0.025                 'volume multiplier; must not be zero
DelayedBallDropOnPlayfieldSoundLevel = 0.8                  'volume level; range [0, 1]
WallImpactSoundFactor = 0.075                     'volume multiplier; must not be zero
MetalImpactSoundFactor = 0.075/3
SaucerLockSoundLevel = 0.8
SaucerKickSoundLevel = 0.8
BlastZoneDrop = 0.3002




'///////////////////////-----Gates, Spinners, Rollovers and Targets-----///////////////////////

Dim GateSoundLevel, TargetSoundFactor, SpinnerSoundLevel, RolloverSoundLevel, DTSoundLevel

GateSoundLevel = 0.5/5                          'volume level; range [0, 1]
TargetSoundFactor = 0.0025 * 10                     'volume multiplier; must not be zero
DTSoundLevel = 0.25                           'volume multiplier; must not be zero
RolloverSoundLevel = 0.25                                       'volume level; range [0, 1]
SpinnerSoundLevel = 0.1                                       'volume level; range [0, 1]

'///////////////////////-----Ball Release, Guides and Drain-----///////////////////////
Dim DrainSoundLevel, BallReleaseSoundLevel, BottomArchBallGuideSoundFactor, FlipperBallGuideSoundFactor

DrainSoundLevel = 0.8                           'volume level; range [0, 1]
BallReleaseSoundLevel = 1                       'volume level; range [0, 1]
BottomArchBallGuideSoundFactor = 0.2                  'volume multiplier; must not be zero
FlipperBallGuideSoundFactor = 0.015                   'volume multiplier; must not be zero

'///////////////////////-----Loops and Lanes-----///////////////////////
Dim ArchSoundFactor
ArchSoundFactor = 0.025/5                         'volume multiplier; must not be zero


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
  PlaySound playsoundparams, -1, min(aVol,1) * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
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
  PlaySound playsoundparams, -1, min(aVol,1) * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
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
  Playsound soundname, 1,min(aVol,1) * VolumeDial, AudioPan(ActiveBall), 0,0,0, 1, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtBallVolM (Soundname, aVol)
  Playsound soundname, 1,min(aVol,1) * VolumeDial, AudioPan(ActiveBall), 0,0,0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtVolLoops(sound, tableobj, Vol, Loops)
  PlaySound sound, Loops, Vol * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub


'******************************************************
'  Fleep  Supporting Ball & Sound Functions
'******************************************************

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
    tmp = tableobj.y * 2 / tableheight-1

  if tmp > 7000 Then
    tmp = 7000
  elseif tmp < -7000 Then
    tmp = -7000
  end if

    If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
  'Debug.Print "" & tableobj.name
    Dim tmp
    tmp = tableobj.x * 2 / tablewidth-1

  if tmp > 7000 Then
    tmp = 7000
  elseif tmp < -7000 Then
    tmp = -7000
  end if

    If tmp > 0 Then
        AudioPan = Csng(tmp ^10)
    Else
        AudioPan = Csng(-((- tmp) ^10) )
    End If
End Function

Function Vol(ball) ' Calculates the volume of the sound based on the ball speed
  Vol = Csng(BallVel(ball) ^2)
End Function

Function Volz(ball) ' Calculates the volume of the sound based on the ball speed
  Volz = Csng((ball.velz) ^2)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
  Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
  BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

Function VolPlayfieldRoll(ball) ' Calculates the roll volume of the sound based on the ball speed
  VolPlayfieldRoll = RollingSoundFactor * 0.0005 * Csng(BallVel(ball) ^3)
End Function

Function PitchPlayfieldRoll(ball) ' Calculates the roll pitch of the sound based on the ball speed
  PitchPlayfieldRoll = BallVel(ball) ^2 * 15
End Function

Function RndInt(min, max)
  RndInt = Int(Rnd() * (max-min + 1) + min)' Sets a random number integer between min and max
End Function

Function RndNum(min, max)
  RndNum = Rnd() * (max-min) + min' Sets a random number between min and max
End Function

'/////////////////////////////  GENERAL SOUND SUBROUTINES  ////////////////////////////
Sub SoundStartButton()
  PlaySound ("Start_Button"), 0, StartButtonSoundLevel, 0, 0.25
End Sub

Sub SoundNudgeLeft()
  PlaySound ("Nudge_" & Int(Rnd*2)+1), 0, NudgeLeftSoundLevel * VolumeDial, -0.1, 0.25
End Sub

Sub SoundNudgeRight()
  PlaySound ("Nudge_" & Int(Rnd*2)+1), 0, NudgeRightSoundLevel * VolumeDial, 0.1, 0.25
End Sub

Sub SoundNudgeCenter()
  PlaySound ("Nudge_" & Int(Rnd*2)+1), 0, NudgeCenterSoundLevel * VolumeDial, 0, 0.25
End Sub


Sub SoundPlungerPull()
  PlaySoundAtLevelStatic ("Plunger_Pull_1"), PlungerPullSoundLevel, PlungerR
End Sub

Sub SoundPlungerReleaseBall()
  PlaySoundAtLevelStatic ("Plunger_Release_Ball"), PlungerReleaseSoundLevel, PlungerR
End Sub

Sub SoundPlungerReleaseNoBall()
  PlaySoundAtLevelStatic ("Plunger_Release_No_Ball"), PlungerReleaseSoundLevel, PlungerR
End Sub


'//////////////////////////////////  SHAKER  ////////////////////////////////////
Sub ShakerMotorSol(Enabled)
  If Enabled then
    SoundShakerDOF SoundOn, DOFOn
    'Debug.print "ShakerMotorSol enabled"
  Else
    SoundShakerDOF SoundOff, DOFOff
    'Debug.print "ShakerMotorSol disabled"
  End if
End Sub


'DOF numbers for shaker intensity
'126 - Shaker Intensity 1
'127 - Shaker Intensity 2
'128 - Shaker Intensity 3

Sub SoundShakerDOF(toggle, DOFstate)
  Select Case toggle
    Case SoundOn
      Select Case ShakerIntensity
        Case 0
        Case 1
          PlaySound SoundFXDOF(Cartridge_Table_Specifics & "_Shaker_Level_01_RampUp_and_Loop", 126, DOFstate, DOFShaker), -1, ShakerSoundLevel * VolumeDial, 0, 0, 0, 1, 1, 0
        Case 2
          PlaySound SoundFXDOF(Cartridge_Table_Specifics & "_Shaker_Level_02_RampUp_and_Loop", 127, DOFstate, DOFShaker), -1, ShakerSoundLevel * VolumeDial, 0, 0, 0, 1, 1, 0
        Case 3
          PlaySound SoundFXDOF(Cartridge_Table_Specifics & "_Shaker_Level_03_RampUp_and_Loop", 128, DOFstate, DOFShaker), -1, ShakerSoundLevel * VolumeDial, 0, 0, 0, 1, 1, 0
      End Select
    Case SoundOff
      Select Case ShakerIntensity
        Case 0
        Case 1
          PlaySound SoundFXDOF(Cartridge_Table_Specifics & "_Shaker_Level_01_RampDown_Only", 126, DOFstate, DOFShaker), 0, ShakerSoundLevel * VolumeDial
          StopSound Cartridge_Table_Specifics & "_Shaker_Level_01_RampUp_and_Loop"
        Case 2
          PlaySound SoundFXDOF(Cartridge_Table_Specifics & "_Shaker_Level_02_RampDown_Only", 127, DOFstate, DOFShaker), 0, ShakerSoundLevel * VolumeDial
          StopSound Cartridge_Table_Specifics & "_Shaker_Level_02_RampUp_and_Loop"
        Case 3
          PlaySound SoundFXDOF(Cartridge_Table_Specifics & "_Shaker_Level_03_RampDown_Only", 128, DOFstate, DOFShaker), 0, ShakerSoundLevel * VolumeDial
          StopSound Cartridge_Table_Specifics & "_Shaker_Level_03_RampUp_and_Loop"
      End Select
  End Select
End Sub

'/////////////////////////////  KNOCKER SOLENOID  ////////////////////////////
Sub KnockerSolenoid()
  PlaySoundAtLevelStatic SoundFX("Knocker_1",DOFKnocker), KnockerSoundLevel, KnockerPosition
End Sub

'/////////////////////////////  DRAIN SOUNDS  ////////////////////////////
Sub RandomSoundDrain(drainswitch)
  PlaySoundAtLevelStatic ("Drain_" & Int(Rnd*11)+1), DrainSoundLevel, drainswitch
End Sub

'/////////////////////////////  TROUGH BALL RELEASE SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundBallRelease(drainswitch)
  PlaySoundAtLevelStatic SoundFX("BallRelease" & Int(Rnd*7)+1,DOFContactors), BallReleaseSoundLevel, drainswitch
End Sub

'/////////////////////////////  SLINGSHOT SOLENOID SOUNDS  ////////////////////////////
Sub RandomSoundSlingshotLeft(sling)
  PlaySoundAtLevelStatic SoundFX("Sling_L" & Int(Rnd*10)+1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

Sub RandomSoundSlingshotRight(sling)
  PlaySoundAtLevelStatic SoundFX("Sling_R" & Int(Rnd*8)+1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

'/////////////////////////////  BUMPER SOLENOID SOUNDS  ////////////////////////////
Sub RandomSoundBumperTop(Bump)
' Debug.Print "Bump"
  PlaySoundAtLevelStatic SoundFX("Bumpers_Top_" & Int(Rnd*5)+1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperMiddle(Bump)
  PlaySoundAtLevelStatic SoundFX("Bumpers_Middle_" & Int(Rnd*5)+1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperBottom(Bump)
  PlaySoundAtLevelStatic SoundFX("Bumpers_Bottom_" & Int(Rnd*5)+1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

'/////////////////////////////  SPINNER SOUNDS  ////////////////////////////
Sub SoundSpinner(spinnerswitch)
  PlaySoundAtLevelStatic ("Spinner"), SpinnerSoundLevel, spinnerswitch
End Sub


'/////////////////////////////  FLIPPER BATS SOUND SUBROUTINES  ////////////////////////////
'/////////////////////////////  FLIPPER BATS SOLENOID ATTACK SOUND  ////////////////////////////
Sub SoundFlipperUpAttackLeft(flipper)
  FlipperUpAttackLeftSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
  PlaySoundAtLevelStatic ("Flipper_Attack-L01"), FlipperUpAttackLeftSoundLevel, flipper
End Sub

Sub SoundFlipperUpAttackRight(flipper)
  FlipperUpAttackRightSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
  PlaySoundAtLevelStatic ("Flipper_Attack-R01"), FlipperUpAttackLeftSoundLevel, flipper
End Sub

'/////////////////////////////  FLIPPER BATS SOLENOID CORE SOUND  ////////////////////////////
Sub RandomSoundFlipperUpLeft(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_L0" & Int(Rnd*9)+1,DOFFlippers), FlipperLeftHitParm, Flipper
End Sub

Sub RandomSoundFlipperUpRight(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_R0" & Int(Rnd*9)+1,DOFFlippers), FlipperRightHitParm, Flipper
End Sub

Sub RandomSoundReflipUpLeft(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_L0" & Int(Rnd*3)+1,DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundReflipUpRight(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_R0" & Int(Rnd*3)+1,DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperDownLeft(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_Left_Down_" & Int(Rnd*7)+1,DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperDownRight(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_Right_Down_" & Int(Rnd*8)+1,DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

'/////////////////////////////  FLIPPER BATS BALL COLLIDE SOUND  ////////////////////////////

Sub LeftFlipperCollide(parm)
  FlipperLeftHitParm = parm/10
  If FlipperLeftHitParm > 1 Then
    FlipperLeftHitParm = 1
  End If
  FlipperLeftHitParm = FlipperUpSoundLevel * FlipperLeftHitParm
  RandomSoundRubberFlipper(parm)
End Sub

Sub RightFlipperCollide(parm)
  FlipperRightHitParm = parm/10
  If FlipperRightHitParm > 1 Then
    FlipperRightHitParm = 1
  End If
  FlipperRightHitParm = FlipperUpSoundLevel * FlipperRightHitParm
  RandomSoundRubberFlipper(parm)
End Sub

Sub RandomSoundRubberFlipper(parm)
  PlaySoundAtLevelActiveBall ("Flipper_Rubber_" & Int(Rnd*7)+1), parm  * RubberFlipperSoundFactor
End Sub

'/////////////////////////////  ROLLOVER SOUNDS  ////////////////////////////
Sub RandomSoundRollover()
  PlaySoundAtLevelActiveBall ("Rollover_" & Int(Rnd*4)+1), RolloverSoundLevel
End Sub

Sub Rollovers_Hit(idx)
  RandomSoundRollover
End Sub

'/////////////////////////////  VARIOUS PLAYFIELD SOUND SUBROUTINES  ////////////////////////////
'/////////////////////////////  RUBBERS AND POSTS  ////////////////////////////
'/////////////////////////////  RUBBERS - EVENTS  ////////////////////////////
Sub Rubbers_Hit(idx)
  dim finalspeed
  finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 5 then
    RandomSoundRubberStrong 1
  End if
  If finalspeed <= 5 then
    RandomSoundRubberWeak()
  End If
End Sub

'/////////////////////////////  RUBBERS AND POSTS - STRONG IMPACTS  ////////////////////////////
Sub RandomSoundRubberStrong(voladj)
  Select Case Int(Rnd*10)+1
    Case 1 : PlaySoundAtLevelActiveBall ("Rubber_Strong_1"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 2 : PlaySoundAtLevelActiveBall ("Rubber_Strong_2"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 3 : PlaySoundAtLevelActiveBall ("Rubber_Strong_3"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 4 : PlaySoundAtLevelActiveBall ("Rubber_Strong_4"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 5 : PlaySoundAtLevelActiveBall ("Rubber_Strong_5"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 6 : PlaySoundAtLevelActiveBall ("Rubber_Strong_6"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 7 : PlaySoundAtLevelActiveBall ("Rubber_Strong_7"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 8 : PlaySoundAtLevelActiveBall ("Rubber_Strong_8"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 9 : PlaySoundAtLevelActiveBall ("Rubber_Strong_9"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 10 : PlaySoundAtLevelActiveBall ("Rubber_1_Hard"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6*voladj
  End Select
End Sub

'/////////////////////////////  RUBBERS AND POSTS - WEAK IMPACTS  ////////////////////////////
Sub RandomSoundRubberWeak()
  PlaySoundAtLevelActiveBall ("Rubber_" & Int(Rnd*9)+1), Vol(ActiveBall) * RubberWeakSoundFactor
End Sub

'/////////////////////////////  WALL IMPACTS  ////////////////////////////
Sub Walls_Hit(idx)
  RandomSoundWall()
End Sub

Sub RandomSoundWall()
  dim finalspeed
  finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    Select Case Int(Rnd*5)+1
      Case 1 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_1"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_2"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_5"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 4 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_7"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 5 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_9"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    Select Case Int(Rnd*4)+1
      Case 1 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_3"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 4 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End If
  If finalspeed < 6 Then
    Select Case Int(Rnd*3)+1
      Case 1 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End if
End Sub

'/////////////////////////////  METAL TOUCH SOUNDS  ////////////////////////////
Sub RandomSoundMetal()
  PlaySoundAtLevelActiveBall ("Metal_Touch_" & Int(Rnd*13)+1), Vol(ActiveBall) * MetalImpactSoundFactor
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
  dim finalspeed
  finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    PlaySoundAtLevelActiveBall ("Apron_Bounce_"& Int(Rnd*2)+1), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    Select Case Int(Rnd*2)+1
      Case 1 : PlaySoundAtLevelActiveBall ("Apron_Bounce_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
      Case 2 : PlaySoundAtLevelActiveBall ("Apron_Bounce_Soft_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
    End Select
  End If
  If finalspeed < 6 Then
    Select Case Int(Rnd*2)+1
      Case 1 : PlaySoundAtLevelActiveBall ("Apron_Bounce_Soft_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
      Case 2 : PlaySoundAtLevelActiveBall ("Apron_Medium_3"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
    End Select
  End if
End Sub

'/////////////////////////////  BOTTOM ARCH BALL GUIDE - HARD HITS  ////////////////////////////
Sub RandomSoundBottomArchBallGuideHardHit()
  PlaySoundAtLevelActiveBall ("Apron_Hard_Hit_" & Int(Rnd*3)+1), BottomArchBallGuideSoundFactor * 0.25
End Sub

Sub Apron_Hit (idx)
  If Abs(cor.ballvelx(activeball.id) < 4) and cor.ballvely(activeball.id) > 7 then
    RandomSoundBottomArchBallGuideHardHit()
  Else
    RandomSoundBottomArchBallGuide
  End If
End Sub

'/////////////////////////////  FLIPPER BALL GUIDE  ////////////////////////////
Sub RandomSoundFlipperBallGuide()
  dim finalspeed
  finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    Select Case Int(Rnd*2)+1
      Case 1 : PlaySoundAtLevelActiveBall ("Apron_Hard_1"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
      Case 2 : PlaySoundAtLevelActiveBall ("Apron_Hard_2"),  Vol(ActiveBall) * 0.8 * FlipperBallGuideSoundFactor
    End Select
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    PlaySoundAtLevelActiveBall ("Apron_Medium_" & Int(Rnd*3)+1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
  End If
  If finalspeed < 6 Then
    PlaySoundAtLevelActiveBall ("Apron_Soft_" & Int(Rnd*7)+1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
  End If
End Sub

'/////////////////////////////  TARGET HIT SOUNDS  ////////////////////////////
Sub RandomSoundTargetHitStrong()
  PlaySoundAtLevelActiveBall SoundFX("Target_Hit_" & Int(Rnd*4)+5,DOFTargets), Vol(ActiveBall) * 0.45 * TargetSoundFactor
End Sub

Sub RandomSoundTargetHitWeak()
  PlaySoundAtLevelActiveBall SoundFX("Target_Hit_" & Int(Rnd*4)+1,DOFTargets), Vol(ActiveBall) * TargetSoundFactor
End Sub

Sub PlayTargetSound()
  dim finalspeed
  finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 10 then
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
  Select Case Int(Rnd*9)+1
    Case 1 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_1"), volz(aBall) * BallBouncePlayfieldSoftFactor, aBall
    Case 2 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_2"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.5, aBall
    Case 3 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_3"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.8, aBall
    Case 4 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_4"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.5, aBall
    Case 5 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_5"), volz(aBall) * BallBouncePlayfieldSoftFactor, aBall
    Case 6 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_1"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
    Case 7 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_2"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
    Case 8 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_5"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
    Case 9 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_7"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.3, aBall
  End Select
End Sub

Sub RandomSoundBallBouncePlayfieldHard(aBall)
  PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_" & Int(Rnd*7)+1), volz(aBall) * BallBouncePlayfieldHardFactor, aBall
End Sub

'/////////////////////////////  DELAYED DROP - TO PLAYFIELD - SOUND  ////////////////////////////
Sub RandomSoundDelayedBallDropOnPlayfield(aBall)
  Select Case Int(Rnd*5)+1
    Case 1 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_1_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 2 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_2_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 3 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_3_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 4 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_4_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 5 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_5_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
  End Select
End Sub

'/////////////////////////////  BALL GATES AND BRACKET GATES SOUNDS  ////////////////////////////

Sub SoundPlayfieldGate()
  PlaySoundAtLevelStatic ("Gate_FastTrigger_" & Int(Rnd*2)+1), GateSoundLevel, Activeball
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
  PlaySoundAtLevelActiveBall ("Arch_L" & Int(Rnd*4)+1), Vol(ActiveBall) * ArchSoundFactor
End Sub

Sub RandomSoundRightArch()
  PlaySoundAtLevelActiveBall ("Arch_R" & Int(Rnd*4)+1), Vol(ActiveBall) * ArchSoundFactor
End Sub


Sub Arch1_hit()
  If Activeball.velx > 1 Then SoundPlayfieldGate
  StopSound "Arch_L1"
  StopSound "Arch_L2"
  StopSound "Arch_L3"
  StopSound "Arch_L4"
End Sub

Sub Arch1_unhit()
  If activeball.velx < -8 Then
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
  PlaySoundAtLevelStatic ("Saucer_Enter_" & Int(Rnd*2)+1), SaucerLockSoundLevel, Activeball
End Sub

Sub SoundSaucerKick(scenario, saucer)
  Select Case scenario
    Case 0: PlaySoundAtLevelStatic SoundFX("Saucer_Empty", DOFContactors), SaucerKickSoundLevel, saucer
    Case 1: PlaySoundAtLevelStatic SoundFX("Saucer_Kick", DOFContactors), SaucerKickSoundLevel, saucer
  End Select
End Sub

'/////////////////////////////  BALL COLLISION SOUND  ////////////////////////////
Sub OnBallBallCollision(ball1, ball2, velocity)
  Dim snd
  Select Case Int(Rnd*7)+1
    Case 1 : snd = "Ball_Collide_1"
    Case 2 : snd = "Ball_Collide_2"
    Case 3 : snd = "Ball_Collide_3"
    Case 4 : snd = "Ball_Collide_4"
    Case 5 : snd = "Ball_Collide_5"
    Case 6 : snd = "Ball_Collide_6"
    Case 7 : snd = "Ball_Collide_7"
  End Select

  PlaySound (snd), 0, Csng(velocity) ^2 / 200 * BallWithBallCollisionSoundFactor * VolumeDial, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)

  FlipperCradleCollision ball1, ball2, velocity
End Sub


'///////////////////////////  DROP TARGET HIT SOUNDS  ///////////////////////////

Sub RandomSoundDropTargetReset(obj)
  PlaySoundAtLevelStatic SoundFX("Drop_Target_Reset_" & Int(Rnd*6)+1,DOFContactors), 1, obj
End Sub

Sub SoundDropTargetDrop(obj)
  PlaySoundAtLevelStatic ("Drop_Target_Down_" & Int(Rnd*6)+1), 200, obj
End Sub

'/////////////////////////////  GI AND FLASHER RELAYS  ////////////////////////////

Const RelayFlashSoundLevel = 0.315                  'volume level; range [0, 1];
Const RelayGISoundLevel = 1.05                  'volume level; range [0, 1];

Sub Sound_GI_Relay(toggle, obj)
  Select Case toggle
    Case 1
      PlaySoundAtLevelStatic ("Relay_GI_On"), 0.025*RelayGISoundLevel, obj
    Case 0
      PlaySoundAtLevelStatic ("Relay_GI_Off"), 0.025*RelayGISoundLevel, obj
  End Select
End Sub

Sub Sound_Flash_Relay(toggle, obj)
  Select Case toggle
    Case 1
      PlaySoundAtLevelStatic ("Relay_Flash_On"), 0.025*RelayFlashSoundLevel, obj
    Case 0
      PlaySoundAtLevelStatic ("Relay_Flash_Off"), 0.025*RelayFlashSoundLevel, obj
  End Select
End Sub

Sub RandomSoundWireRampStop(obj)
  Select Case Int(rnd*6)
    Case 0: PlaySoundAtVol "wireramp_stop", obj, 0.2*volumedial
    Case 1: PlaySoundAtVol "wireramp_stop2", obj, 0.2*volumedial
    Case 2: PlaySoundAtVol "wireramp_stop3", obj, 0.2*volumedial
    Case 3: PlaySoundAtVol "wireramp_stop", obj, 0.1*volumedial
    Case 4: PlaySoundAtVol "wireramp_stop2", obj, 0.1*volumedial
    Case 5: PlaySoundAtVol "wireramp_stop3", obj, 0.1*volumedial
  End Select
End Sub

Sub RandomSoundRampStop(obj)
  Select Case Int(rnd*3)
    Case 0: PlaySoundAtLevelActiveBall ("Rubber_Strong_1"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
    Case 1: PlaySoundAtLevelActiveBall ("Rubber_Strong_2"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
    Case 2: PlaySoundAtLevelActiveBall ("Rubber_1_Hard"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
  End Select
End Sub

Sub RandomSoundBlastZone()
    Select Case Int(Rnd * 3) + 1
    Case 1:PlaySoundAtLevelStatic ("BlastZone_Hit1"), BlastZoneDrop * volumedial, Activeball
        Case 2:PlaySoundAtLevelStatic ("BlastZone_Hit2"), BlastZoneDrop * volumedial, Activeball
        Case 3:PlaySoundAtLevelStatic ("BlastZone_Hit3"), BlastZoneDrop * volumedial, Activeball
    End Select
End Sub

'/////////////////////////////////////////////////////////////////
'         End Mechanical Sounds
'/////////////////////////////////////////////////////////////////


'******************************************************
'**** RAMP ROLLING SFX
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
dim RampBalls(6,2)
'x,0 = ball x,1 = ID, 2 = Protection against ending early (minimum amount of updates)
'0,0 is boolean on/off, 0,1 unused for now
RampBalls(0,0) = False

' RampType
'     Setup: Set this array to the number Total number of balls that can be tracked at one time + 1.  5 ball multiball then set value to 6
'     Description: Array type indexed on BallId and a values used to deterimine what type of ramp the ball is on: False = Wire Ramp, True = Plastic Ramp
dim RampType(6)

Sub WireRampOn(input) : Waddball ActiveBall, input : RampRollUpdate: End Sub
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
      If AltWireSounds Then
        StopSound("wireloop_rt" & x)
            Else
        StopSound("wireloop" & x)
            End If
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
          PlaySound("RampLoop" & x), -1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume, AudioPan(RampBalls(x,0)), 0, BallPitchV(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
          If AltWireSounds Then
            StopSound("wireloop_rt" & x)
          Else
            StopSound("wireloop" & x)
          End If
        Else
          StopSound("RampLoop" & x)
          If AltWireSounds Then
            PlaySound("wireloop_rt" & x), -1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume, AudioPan(RampBalls(x,0)), 0, BallPitch(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
          Else
            PlaySound("wireloop" & x), -1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume, AudioPan(RampBalls(x,0)), 0, BallPitch(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
          End If
        End If
        RampBalls(x, 2) = RampBalls(x, 2) + 1
      Else
        StopSound("RampLoop" & x)
        If AltWireSounds Then
          StopSound("wireloop_rt" & x)
        Else
          StopSound("wireloop" & x)
        End If
      end if
      if RampBalls(x,0).Z < 30 and RampBalls(x, 2) > RampMinLoops then  'if ball is on the PF, remove  it
        StopSound("RampLoop" & x)
        If AltWireSounds Then
          StopSound("wireloop_rt" & x)
        Else
          StopSound("wireloop" & x)
        End If
        Wremoveball RampBalls(x,1)
      End If
    Else
      StopSound("RampLoop" & x)
      If AltWireSounds Then
        StopSound("wireloop_rt" & x)
            Else
        StopSound("wireloop" & x)
            End If
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



'******************************************************
'**** END RAMP ROLLING SFX
'******************************************************


'***************************************************************
'****  VPW DYNAMIC BALL SHADOWS by Iakki, Apophis, and Wylte
'***************************************************************

'****** INSTRUCTIONS please read ******

'****** Part A:  Table Elements ******
'
' Import the "bsrtx7" and "ballshadow" images
' Import the shadow materials file (3 sets included) (you can also export the 3 sets from this table to create the same file)
' Copy in the BallShadowA flasher set and the sets of primitives named BallShadow#, RtxBallShadow#, and RtxBall2Shadow#
' * Count from 0 up, with at least as many objects each as there can be balls, including locked balls
' Ensure you have a timer with a -1 interval that is always running

' Create a collection called DynamicSources that includes all light sources you want to cast ball shadows
' It's recommended that you be selective in which lights go in this collection, as there are limitations:
' 1. The shadows can "pass through" solid objects and other light sources, so be mindful of where the lights would actually able to cast shadows
' 2. If there are more than two equidistant sources, the shadows can suddenly switch on and off, so places like top and bottom lanes need attention
' 3. At this time the shadows get the light on/off from tracking gilvl, so if you have lights you want shadows for that are on at different times you will need to:
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

' *** This segment goes within the RollingUpdate sub, so that if Ambient...=0 and Dynamic...=0 the entire DynamicBSUpdate sub can be skipped for max performance
' *** Includes lines commonly found there, for reference:
' ' stop the sound of deleted balls
' For b = UBound(gBOT) + 1 to tnob
'   If AmbientBallShadowOn = 0 Then BallShadowA(b).visible = 0
'   ...rolling(b) = False
'   ...StopSound("BallRoll_" & b)
' Next
'
' ...rolling and drop sounds...

'   If DropCount(b) < 5 Then
'     DropCount(b) = DropCount(b) + 1
'   End If
'
'   ' "Static" Ball Shadows
'   If AmbientBallShadowOn = 0 Then
'     If gBOT(b).Z > 30 Then
'       BallShadowA(b).height=gBOT(b).z - BallSize/4    'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping
'     Else
'       BallShadowA(b).height=gBOT(b).z - BallSize/2 + 5
'     End If
'     BallShadowA(b).Y = gBOT(b).Y + Ballsize/5 + fovY
'     BallShadowA(b).X = gBOT(b).X
'     BallShadowA(b).visible = 1
'   End If

' *** Required Functions, enable these if they are not already present elswhere in your table
'Function max(a,b)
' if a > b then
'   max = a
' Else
'   max = b
' end if
'end Function

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

Dim DSSources(30), numberofsources', DSGISide(30) 'Adapted for TZ with GI left / GI right

' *** Trim or extend these to *match* the number of balls/primitives/flashers on the table!
dim objrtx1(5), objrtx2(5)
dim objBallShadow(5)
Dim BallShadowA
BallShadowA = Array (BallShadowA0,BallShadowA1,BallShadowA2,BallShadowA3,BallShadowA4)

DynamicBSInit

sub DynamicBSInit()
  Dim iii, source

  for iii = 0 to tnob - 1               'Prepares the shadow objects before play begins
    Set objrtx1(iii) = Eval("RtxBallShadow" & iii)
    objrtx1(iii).material = "RtxBallShadow" & iii
    objrtx1(iii).z = 1 + iii/1000 + 0.21      'Separate z for layering without clipping
    objrtx1(iii).visible = 0

    Set objrtx2(iii) = Eval("RtxBall2Shadow" & iii)
    objrtx2(iii).material = "RtxBallShadow2_" & iii
    objrtx2(iii).z = 1 + iii/1000 + 0.22
    objrtx2(iii).visible = 0

    Set objBallShadow(iii) = Eval("BallShadow" & iii)
    objBallShadow(iii).material = "BallShadow" & iii
    UpdateMaterial objBallShadow(iii).material,1,0,0,0,0,0,AmbientBSFactor,RGB(0,0,0),0,0,False,True,0,0,0,0
    objBallShadow(iii).Z = 1 + iii/1000 + 0.24
    objBallShadow(iii).visible = 0

    BallShadowA(iii).Opacity = 100*AmbientBSFactor
    BallShadowA(iii).visible = 0
  Next

  iii = 0

  For Each Source in DynamicSources
    DSSources(iii) = Array(Source.x, Source.y)
'   If Instr(Source.name , "Left") > 0 Then DSGISide(iii) = 0 Else DSGISide(iii) = 1  'Adapted for TZ with GI left / GI right
    iii = iii + 1
  Next
  numberofsources = iii
end sub

' *** These define the appearance of shadows in your table
Const DynamicBSFactor     = 0.95  '0 to 1, higher is darker
Const AmbientBSFactor     = 0.9 '0 to 1, higher is darker
Const AmbientMovement   = 2   '1 to 4, higher means more movement as the ball moves left and right
Const Wideness        = 20  'Sets how wide the dynamic ball shadows can get (20 +5 thinness is technically most accurate for lights at 25+ hitting a 50 unit ball)
Const Thinness        = 5   'Sets minimum as ball moves away from source
Const fovY          = 0   'Offset y position under ball to account for layback or inclination (personal preference)

Sub DynamicBSUpdate
  Dim falloff: falloff = 150 'Max distance to light sources, can be changed if you have a reason
  Dim ShadowOpacity1, ShadowOpacity2
  Dim s, LSd, iii
  Dim dist1, dist2, src1, src2
' Dim BOT: BOT=getballs

  'Hide shadow of deleted balls
  For s = UBound(gBOT) + 1 to tnob - 1
    objrtx1(s).visible = 0
    objrtx2(s).visible = 0
    objBallShadow(s).visible = 0
    BallShadowA(s).visible = 0
  Next

  If UBound(gBOT) < lob Then Exit Sub   'No balls in play, exit

'The Magic happens now
  For s = lob to UBound(gBOT)

' *** Normal "ambient light" ball shadow
  'Layered from top to bottom. If you had an upper pf at for example 80 units and ramps even above that, your segments would be z>110; z<=110 And z>100; z<=100 And z>30; z<=30 And z>20; Else invisible

    If AmbientBallShadowOn = 1 Then     'Primitive shadow on playfield, flasher shadow in ramps
      If gBOT(s).Z > 30 Then              'The flasher follows the ball up ramps while the primitive is on the pf
        If gBOT(s).X < tablewidth/2 Then
          objBallShadow(s).X = ((gBOT(s).X) - (Ballsize/10) + ((gBOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) + 5
        Else
          objBallShadow(s).X = ((gBOT(s).X) + (Ballsize/10) + ((gBOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) - 5
        End If
        objBallShadow(s).Y = gBOT(s).Y + BallSize/10 + fovY
        objBallShadow(s).visible = 1

        BallShadowA(s).X = gBOT(s).X
        BallShadowA(s).Y = gBOT(s).Y + BallSize/5 + fovY
        BallShadowA(s).height=gBOT(s).z - BallSize/4    'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping
        BallShadowA(s).visible = 1
      Elseif gBOT(s).Z <= 30 And gBOT(s).Z > 20 Then  'On pf, primitive only
        objBallShadow(s).visible = 1
        If gBOT(s).X < tablewidth/2 Then
          objBallShadow(s).X = ((gBOT(s).X) - (Ballsize/10) + ((gBOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) + 5
        Else
          objBallShadow(s).X = ((gBOT(s).X) + (Ballsize/10) + ((gBOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) - 5
        End If
        objBallShadow(s).Y = gBOT(s).Y + fovY
        BallShadowA(s).visible = 0
      Else                      'Under pf, no shadows
        objBallShadow(s).visible = 0
        BallShadowA(s).visible = 0
      end if

    Elseif AmbientBallShadowOn = 2 Then   'Flasher shadow everywhere
      If gBOT(s).Z > 30 Then              'In a ramp
        BallShadowA(s).X = gBOT(s).X
        BallShadowA(s).Y = gBOT(s).Y + BallSize/5 + fovY
        BallShadowA(s).height=gBOT(s).z - BallSize/4    'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping
        BallShadowA(s).visible = 1
      Elseif gBOT(s).Z <= 30 And gBOT(s).Z > 20 Then  'On pf
        BallShadowA(s).visible = 1
        If gBOT(s).X < tablewidth/2 Then
          BallShadowA(s).X = ((gBOT(s).X) - (Ballsize/10) + ((gBOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) + 5
        Else
          BallShadowA(s).X = ((gBOT(s).X) + (Ballsize/10) + ((gBOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) - 5
        End If
        BallShadowA(s).Y = gBOT(s).Y + Ballsize/10 + fovY
        BallShadowA(s).height=gBOT(s).z - BallSize/2 + 5
      Else                      'Under pf
        BallShadowA(s).visible = 0
      End If
    End If

' *** Dynamic shadows
    If DynamicBallShadowsOn Then
      If gBOT(s).Z < 30 Then
        dist1 = falloff
        dist2 = falloff
        For iii = 0 to numberofsources - 1 ' Search the 2 nearest influencing lights
          LSd = Distance(gBOT(s).x, gBOT(s).y, DSSources(iii)(0), DSSources(iii)(1)) 'Calculating the Linear distance to the Source
          If LSd < dist1 And gilvl > 0 Then
'         If LSd < dist2 And ((DSGISide(iii) = 0 And Lampz.State(100)>0) Or (DSGISide(iii) = 1 And Lampz.State(104)>0)) Then  'Adapted for TZ with GI left / GI right
            dist2 = dist1
            dist1 = LSd
            src2 = src1
            src1 = iii
          End If
        Next
        ShadowOpacity1 = 0
        If dist1 < falloff Then
          objrtx1(s).visible = 1 : objrtx1(s).X = gBOT(s).X : objrtx1(s).Y = gBOT(s).Y + fovY
          'objrtx1(s).Z = BOT(s).Z - 25 + s/1000 + 0.01 'Uncomment if you want to add shadows to an upper/lower pf
          objrtx1(s).rotz = AnglePP(DSSources(src1)(0), DSSources(src1)(1), gBOT(s).X, gBOT(s).Y) + 90
          ShadowOpacity1 = 1 - dist1 / falloff
          objrtx1(s).size_y = Wideness * ShadowOpacity1 + Thinness
          UpdateMaterial objrtx1(s).material,1,0,0,0,0,0,ShadowOpacity1*DynamicBSFactor^3,RGB(0,0,0),0,0,False,True,0,0,0,0
        Else
          objrtx1(s).visible = 0
        End If
        ShadowOpacity2 = 0
        If dist2 < falloff Then
          objrtx2(s).visible = 1 : objrtx2(s).X = gBOT(s).X : objrtx2(s).Y = gBOT(s).Y + fovY
          'objrtx2(s).Z = BOT(s).Z - 25 + s/1000 + 0.01 'Uncomment if you want to add shadows to an upper/lower pf
          objrtx2(s).rotz = AnglePP(DSSources(src2)(0), DSSources(src2)(1), gBOT(s).X, gBOT(s).Y) + 90
          ShadowOpacity2 = 1 - dist2 / falloff
          objrtx2(s).size_y = Wideness * ShadowOpacity2 + Thinness
          UpdateMaterial objrtx2(s).material,1,0,0,0,0,0,ShadowOpacity2*DynamicBSFactor^3,RGB(0,0,0),0,0,False,True,0,0,0,0
        Else
          objrtx2(s).visible = 0
        End If
        If AmbientBallShadowOn = 1 Then
          'Fades the ambient shadow (primitive only) when it's close to a light
          UpdateMaterial objBallShadow(s).material,1,0,0,0,0,0,AmbientBSFactor*(1 - max(ShadowOpacity1, ShadowOpacity2)),RGB(0,0,0),0,0,False,True,0,0,0,0
        Else
          BallShadowA(s).Opacity = 100 * AmbientBSFactor * (1 - max(ShadowOpacity1, ShadowOpacity2))
        End If
      Else 'Hide dynamic shadows everywhere else
        objrtx2(s).visible = 0 : objrtx1(s).visible = 0
      End If
    End If
  Next
End Sub
'****************************************************************
'****  END VPW DYNAMIC BALL SHADOWS by Iakki, Apophis, and Wylte
'****************************************************************




'****************************************************************
'   VR Mode
'****************************************************************
'0 - VR Room Off, 1 - Minimal Room, 2 - Ultra Minimal

DIM VRThings
Sub SetupVRRoom()
  If RenderingMode = 2 Then VRRoom = VRRoomChoice Else VRRoom = 0
  If VRRoom <> 0 Then
    ScoreText.visible = 0
    DMD.visible = 1
    PinCab_Rails.visible = 1
    Table1.PlayfieldReflectionStrength = 40
    If VRRoom = 1 Then
      for each VRThings in VR_Cab:VRThings.visible = 1:Next
      for each VRThings in VR_Min:VRThings.visible = 1:Next
      Pincab_ShooterRtop.visible = 1
      Pincab_ShooterLtop.visible = 1
    End If
    If VRRoom = 2 Then
      for each VRThings in VR_Cab:VRThings.visible = 0:Next
      for each VRThings in VR_Min:VRThings.visible = 0:Next
      PinCab_Backbox.visible = 1
      PinCab_Backglass.visible = 1
    End If
  Else
    for each VRThings in VR_Cab:VRThings.visible = 0:Next
    for each VRThings in VR_Min:VRThings.visible = 0:Next
  End if
End Sub




' Hack to return Narnia ball back in play
Sub Narnia_Timer
    Dim b
    For b = 0 to UBound(gBOT)
        if gBOT(b).z < -200 Then
'           msgbox "Ball " &b& " in Narnia X: " & gBOT(b).x &" Y: "&gBOT(b).y & " Z: "&gBOT(b).z
'           debug.print "Move narnia ball ("& gBOT(b).x &" Y: "&gBOT(b).y & " Z: "&gBOT(b).z&") to upper left vuk"
'           gBOT(b).x = 568
'           gBOT(b).y = 2016
'           gBOT(b).z = 26
            gBOT(b).velx = 0 : gBOT(b).vely = 0 : gBOT(b).velz = 0
            gBOT(b).x = 150 : gBOT(b).y = 387  : gBOT(b).z = 0
        end if
    next
end sub



'*************
'VPW Changelog
'*************
'v0.001 - Sixtoe - Script butchery, physical trough and blast zone (save the balls!)
'v0.002 - Sixtoe - Continued script butchery, nfozzy physics, fleep sound calls (but no sounds)
'v0.003 - Sixtoe - Implemented physical ted lock, well, I say that, but it doesn't work. Locks if you hold a ball above sw52.
'v0.004 - apophis - Got the multiball lock mech working
'v0.005 - Sixtoe - Lighting! Flupper Flashers, Lampz, GI, Flupper Bumpers, Light Shaping, Blooms, Cab Lockdown & Rails, loads of other tweaks.
'v0.006 - Sixtoe - Fixed bumpers (thanks Oqqsan), Added VR cabinet and minimal room, fixed some VR issues, fluppers flippers added, flipper shadows added.
'v0.007 - Skitso - Full GI remake. Color / brightness corrected textures, replaced PF texture to one without ramp shadows, improved over pf bulbs, fixed red bulbs over ramp entrances, improved spots, rotated bumper caps to correct position, new ball, new lut, tons of DL tweaking.
'v0.008 - Sixtoe - Primitive Inserts added and hooked up (need tuning), new playfield cutout & insert text layer, city kicker tuned down, hooked up blast zone playfield flasher.
'v0.009 - Skitso - Tuned primitive inserts. Reverted map edge inserts back to original white (instead of red/blue/white, should maybe add this as a mod?) Reverted banana insert/flasher as a skitso style. Further improvements to GI, added a bit more oooph to flashers
'v0.010 - Skitso - Improved flashers
'v0.011 - fluffhead35 - Wiring up Fleep Sounds, Ball Roll Sounds, and Wire Ramp Sounds.  Removed Duplicate function lampfilter
'v0.012 - Sixtoe - Removed duplicate bumpers, changed some physics materials, adjusted flipper physics surrounds, adjusted right vuk kicker, turned up the bumpers, increased the table slope to 6.5, added delay script to drain, changed a few post settings,
'v0.013 - Sixtoe - Messed around with the subways again, changed head materials, underplayfield wall installed, rectangle spotlight hoods installed, moved ball lock exit wall blocker, adjusted spinner sound, set physics to mid 90's onwards, flipper and physics tweaks as per nfozzy and clark kents suggestions,
'v0.013a - ClarkKent - Tweaked physics
'v0.014 - Sixtoe - Reinforced the entire table to try and prevent wall gaps, hooked up sound for scoops, added dynamic shadows, tidied up and removed unused resources
'v0.015 - Sixtoe - Re-added missing inlane walls.
'v0.016 - Sixtoe - Fixed broken playfield material
'v0.017 - N/A
'v0.018 - Sixtoe - Redid playfield mesh (twice), spent ages trying to tweak city start hole, replaced metal walls, added wigwag flashing lights for ramps (thanks nfozzy & iaakki), added flipper bounces (thanks apophis), changed wall entrance profile on left uturn (thanks clarkkent), fixed dynamic shadow depth bias (thanks apophis), fixed right wire ramp exit, changed tedtrig, changed left plunger strength, adjusted outlane rubbers
'v0.019 - Sixtoe - Tweaked lower small flipper area (thanks ClarkKent), adjusted blastzone sounds (thanks wytle), added hit sounds to Red and Ted's mouth plastics, fixed some other sounds, changed depth bias for upper flipper shadows, blastzone wall cutout added, updated VR cabinet images (thanks ClarkKent).
'v0.020 - fluffhead35 - Removed VolumeDial from the BallRoll and RampRoll sound math to make the sound louder.  Uncommented shadow code from RollingUpdate.  Added optoin to use Alternate WireRamp sounds from original table.
'v0.021 - Sixtoe - Fixed lower stubby flipper ball trap, fixed apron sound
'     Clark Kent - tweaked flippers, adjusted rubbers, adjusted position of the left slingshot rubber bands, moved pegs, tweaked bumpers, reduced elasticity of gates,
'v0.022/23 - Clark Kent
'v0.024 - apophis - Updated flasher code to support modulated flasher levels
'v0.025/26 - Clark Kent
'v0.027 - apophis - Updated sw78 kick strength. Fixed Lampz Update1 to work with -1 timer. Modified fading speeds per as necessary. Implemented some performance improvements when flashers stack.
'v0.027a - apophis - Fixed Bumper and flasher initialization conflict. Fixed some insert bloom falloff sizes.
'v0.032 - Sixtoe - Fixed flashers, fixed bumpers, fixed bumper central flasher, added aftermarket target decals, add delay to dropsound,
'v0.033 - Clark Kent - Wall002 adjusted
'v0.034 - Wylte - Adjusted clear post material, added arch sounds to BlockerWall1, enabled hit sounds for metal walls, set FS pov
'v0.034a - apophis - Fixed bumpers.
'v1.02 - Fleep - Added support for shaker motor sounds: added a cheap version of cartridges use just for the shaker sounds and added new shaker configuration selection
'v1.05 - Fleep - Added replaced fleep mechanical sounds cartridge used for the shaker motor sounds with new REV02
'v1.06 - Fleep - Fixed fleep REV02 sounds not being played back
'v1.07 - Wylte - Fix BlockerWall1 sound/collection, eliminate getballs, wrapped FrameTimer plunger animation to fix VR plunger y overflow when VR=0
'v1.1 - Release
'v1.2.0 - Sixtoe - Adjusted lock mech geometry, changed where narnia code drops ball.
'v1.2.1 - apophis - Added blockerwall near Ted's right cheek to prevent stuck ball. Added wall in subway lock to prevent sticky kicker.
'v1.2.3 - Sixtoe - Rebuilt subway lock, again.
'v1.2.4 - apophis - Updated the shaker code.
'v1.2.5 - ClarkKent -  removed primitive ball guides inside the heads from MetalWallP primitive, added new scoops with bent corners and springs inside the heads, improved the Löcher texture for the mounting plates of the heads with correct wood type.
'v1.2.6 - apophis - Converted sw54 from a kicker to a switch-based VUK. Made some subway objects invisible.
'v1.2.7 - ClarkKent - added plunger lane groove, corrected orientation of the mounting plate of Red a little bit including repositioning Red‘s head and all corresponding primitives, adjusted desktop POV to maximize view size
'v1.3.0 - ClarkKent - small randomization of the strength of the Start City Event scoop kicker, new collidable Start City Event scoop to fix rebounces of the ball, slight reposition of the bumper caps, fixed position of some screws and bolts, corrected size, shape and position of the blue plastics in the lower back right in front of the backboard
'v1.3.1 - Retro27 - VR Room reworks, New Metal works, tweets to cabinet, New plungers, added Start button, extra ball button. Adjusted PlungerR to 90 strenght
'v1.3.2 - ClarkKent - Plastics improved, higher resolution, more details, color corrected. Flasher dome color corrected.
'v1.3.3 - ClarkKent - Primitive ball guides adjusted, some plastics adjusted
'v1.3.4 - ClarkKent - Custom Target Decals color correction, Ramp Decals higher resolution, Apron Decals higher resolution
'v1.3.5 - Primetime5k - Added staged flipper scripting
'v1.3.6 - apophis - Added the flipper coillide subs back in. Now using GetBalls in Cor.Update. Fixed flasher code so they completely turn off. Automated VRRoom and CabinetMode
'v1.3.7 - ClarkKent - left ramp exit improved, apron improved at the plungers
'v1.3.8 - apophis - VolumeDial fix. Added FlipperCradleCollision.
'v1.4 - Release
'v1.4.1 - ClarkKent - replaced small flipper primitive and corrected colors, new accurate spinner and bracket primitive, adjusted BlastZoneDrop sound volume, adjusted bumpers, adjusted scoop, adjusted textures screws and bolts
'v1.4.2 - ClarkKent - added cables to the spinner bracket, added rules card to the table options
'v1.4.3 - ClarkKent - added switchable flipper bat decals and plunger gate decal
'v1.4.4 - ClarkKent - adjusted eyelids
'v1.4.5 - ClarkKent - adjusted some primitive material settings
'v1.4.6 - apophis - made left plunger only active when ball is in left plunger lane, otherwise right plunger is active.
'v1.4.7 - ClarkKent - adjusted fully open eyelids angle, made head primitives collidable
'v1.4.8 - ClarkKent - adjusted Desktop and Cabinet view, replaced left and right adjusting posts with correct metal posts primitives
'v1.4.9 - ClarkKent - adding table options, rules card brighter in table options
'v1.4.10 - apophis - Merged in Nestorgian's option updates. Moved animations from frametimer to _animate subs.
'v1.4.11 - apophis - Reduced reflection of PF on ball.  Added options for playfield reflections and flasher brightness. Changed some timer intervals from 1 to 20 (with associated changes so they still work correctly).
'v1.4.12 - apophis - Fixed VR room option.
'v1.5 Release
'v1.5.1 - apophis - fixed lighting on plastics and ramp textures.
'v1.5.2 - DGrimmReaper - Animated VR buttons
'v1.5.3 - ClarkKent - flipper strength to 3100 to be even closer to the real machine, fixed startcitytrig playing sound ony when hit from above, changed UseVPMColoredDMD
'v1.5.4 - Sixtoe - Small adjustments

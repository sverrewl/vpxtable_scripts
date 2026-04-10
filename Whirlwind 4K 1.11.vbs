' __      __  __  __  ______   ____    __       __      __  ______   __  __  ____
'/\ \  __/\ \/\ \/\ \/\__  _\ /\  _`\ /\ \     /\ \  __/\ \/\__  _\ /\ \/\ \/\  _`\ 
'\ \ \/\ \ \ \ \ \_\ \/_/\ \/ \ \ \L\ \ \ \    \ \ \/\ \ \ \/_/\ \/ \ \ `\\ \ \ \/\ \ 
' \ \ \ \ \ \ \ \  _  \ \ \ \  \ \ ,  /\ \ \  __\ \ \ \ \ \ \ \ \ \  \ \ , ` \ \ \ \ \ 
'  \ \ \_/ \_\ \ \ \ \ \ \_\ \__\ \ \\ \\ \ \L\ \\ \ \_/ \_\ \ \_\ \__\ \ \`\ \ \ \_\ \
'   \ `\___x___/\ \_\ \_\/\_____\\ \_\ \_\ \____/ \ `\___x___/ /\_____\\ \_\ \_\ \____/
'    '\/__//__/  \/_/\/_/\/_____/ \/_/\/ /\/___/   '\/__//__/  \/_____/ \/_/\/_/\/___/
'
'
'****************************************************************************************************************************************
'****************************************************************************************************************************************
'
'  Whirlwind / IPD No. 2765 / January, 1990 / 4 Players
'
' credits:
' Base version VP9 of Herweh used and some bits and pieces from Walamab's Whirlwind for VPX
' Flupper: 3d models, 3d inserts, new flipper bats, photoshop work, rendering, lighting, all scripting related to visuals (except VR)
' Rothbauerw: all the physics coding and objects
' Fleep: sound engineering, sound scripting
' Clark Kent: scanned playfield
' Blackmoor: owns actual machine, sound recordings, scanned plastics, reference images, reference videos for the lighting, testing
' Cosmic80 for cleaning plastics textures for previous vp9 version
' Wrd1972: reference photos for ramps placement, additional sound recordings
' Leojreimroc: scoring displays and backglass lighting in VR
' Rawd & 3rdaxis: VR room, cabinet, backbox, topper and VR script added
' VPW: testing the prerelease beta
'

Option explicit
Randomize
Dim luts, lutpos, FlasherIntensity, FlasherIntensityDome
Dim VRObject, CellarOpt

'****************************************************************************************************************************************
'****************************************************************************************************************************************


'************** VR Options **************************************************************************************************************
Const VRRoom = 0        '0 = Off, 1 = VR Room On
Const VRFlashingBackglass = 1 '0 = Static Backglass, 1 = Flashing Backglass
Const Glass = 1         '0 = Off 1 = On - VR Glass scratches
'****************************************************************************************************************************************
If VRRoom > 0 Then LoadVRRoom() Else SetNoVRRoom() : End If

'************** Cellar Eject Difficulty Option ******************************************************************************************
CellarOpt = 0           '0 = Normal, 1 = Difficult

'************** Flipper Rubberizer Option ***********************************************************************************************
Const RubberizerEnabled = 1     '0 = normal flip rubber, 1 = more lively rubber for flips, 2 = a different rubberizer

' ****** siderails & lockbar visible in FS mode, default = false  ***********************************************************************
Const SideRails = False

' ****** disable ballshadow by setting this to 0, default = 1 ***************************************************************************
Const BallShadowOn = 1

' ****** overall flasher intensity, default = 11, range = 1 to 11 ***********************************************************************
FlasherIntensityDome = 11   ' for flasher effects on the domes
FlasherIntensity = 11       ' for flasher effects on the playfield, walls, etc

' ****** LUT overall contrast & brightness setting **************************************************************************************
luts = array("ColorGradeLUT256x16_ConSat", "ColorGradeLUT256x16_ConSatgam", "LUTww01", "LUT01", "LUT1_1", "ColorGradeLUT256x16_1to1")
lutpos = 0            '  set the nr of the LUT you want to use (0 = first in the list above, 1 = second, etc); 0 is the default
Const EnableMagnasave = 1   ' 1 - on; 0 - off; if on then the magnasave button let's you rotate all LUT's

' ****** rolling text enable / disable **************************************************************************************************
Const EnableRollingText = True

'***********  Use staged flippers (dual leaf switches)  *******************************
'RightMagnaSave could be used by changing the following line (but will interfere with the LUT changer, so disable that):
'keyStagedFlipperR = KeyUpperRight
'to
'keyStagedFlipperR = RightMagnaSave
Const StagedFlipperMod = 0    '0 = not staged, 1 - staged (dual leaf switches)

'////////////////////////////////////////////////////////////////////////////////
'////          Mechanical Sounds Options, by Fleep                           ////
'////////////////////////////////////////////////////////////////////////////////
'
'//////////////////////////  MECHANICAL SOUNDS OPTIONS  /////////////////////////
'//  This section allows to set various general sound options for the mechanical sounds.
'//  For the entire sound system scripts see mechanical sounds block down below in the project.
'
'////////////////////////////  GENERAL SOUND OPTIONS  ///////////////////////////
'
'//  PositionalSoundPlaybackConfiguration:
'//  Specifies the sound playback configuration. Options:
'//  1 = Mono
'//  2 = Stereo (Only L+R channels)
'//  3 = Surround (Full surround sound support for SSF - Front L + Front R + Rear L + Rear R channels)
Const PositionalSoundPlaybackConfiguration = 3
'
'
'//  RelaysPosition:
'//  1 = Relays positioned with power board (Provides sound spread from the left and right back channels)
'//  2 = Relays positioned with GI strings (Provides sound spread from left/right/front/rear surround channels)
Const RelaysPosition = 2
'
'
'//  SpinWheelsMotorConfiguration:
'//  1 = Spin Wheels Motor Sound disabled, DOF disabled
'//  2 = Spin Wheels Motor Sound disabled, DOF enabled
'//  3 = Spin Wheels Motor Sound enabled, DOF disabled
'//  4 = Spin Wheels Motor Sound enabled, DOF enabled
Const SpinWheelsMotorConfiguration = 4
'
'
'//  BackwallFanBlowerConfiguration:
'//  1 = FanBlower Sound disabled, DOF disabled
'//  2 = FanBlower Sound disabled, DOF enabled
'//  3 = FanBlower Sound enabled, DOF disabled
'//  4 = FanBlower Sound enabled, DOF enabled
Const BackwallFanBlowerConfiguration = 1
'
'
'//  VolumeDial:
'//  VolumeDial is the actual global volume multiplier for the mechanical sounds.
'//  Values smaller than 1 will decrease mechanical sounds volume.
'//  Recommended values should be no greater than 1.
'//  Default value is 0.8 and is optimal.
Const VolumeDial = 0.8
'
'
'//  PlayfieldRollVolumeDial:
'//  PlayfieldRollVolumeDial is a constant volume multiplier for the playfield rolling sound.
'//  Default value should be 1 which will guarantee a proper carefully calculated dynamic volume changes profile.
'//  Any values different than the default will impact the volume level and the dynamic voume changes profile.
Const PlayfieldRollVolumeDial = 1
'
'
'//  RampsRollVolumeDial:
'//  RampsRollVolumeDial is a constant volume multiplier for all ramps rolling sounds.
'//  This includes all types of plastic and metal ramps.
'//  Default value should be 1 which will guarantee a proper carefully calculated dynamic volume changes profile.
'//  Any values different than the default will impact the volume level and the dynamic voume changes profile.
'//  For a specific ramp volume change please refer to the corresponding volume variables in the script sound paramter section.
Const RampsRollVolumeDial = 1
'
'
'////////////////////////////////////////////////////////////////////////////////
'////          End of Fleep Mechanical Sounds Options                        ////
'////////////////////////////////////////////////////////////////////////////////

' Changes in version 1.1:
'
'- Fleep: Reworked flipper coil logic to have more dynamic loudness and added new dampened flipper up sounds that engaged if there are nearby ball(s)
'- Fleep: Added PlayfieldRollVolumeDial and RamopsRollVolumeDial to allow playfield and ramps rolling sound volume in global sound options
'- Fleep: Left and Right orbit lanes ball roll sound volume readjusted
'- Fleep: Left and Right Cellars enter sound loudness is now dynamic and depend on ball velocity
'- Fleep: Readjusted ramp lifter/down coils loudness
'- Fleep: 1 Bank and 3 Bank Drop Target coil release samples reprocessed
'- Fleep: Added additional sounds for Inner Lane Ball Guide Hits and readjust loudness
'- Fleep: Plastic Ramps rolling and enter loudness calibration
'- Fleep: AC Solenoid loudness readjusted
'- Fleep: Other minor sound script changes
'- Fleep: Reprocessed and repositioned locking kicker sounds
'- Fleep: Reprocessed plunger pull sound to correlate with the actual plunger speed
'- Fleep: Plunger pull sound now stops when plugner released
'- leojreimroc: new VR backglass, thanks to new image provided by Sheltemke
'- Rothbauerw: Adjusted plunger release speed to 86, it does make it around the loop occasionally.
'- Rothbauerw: Adjusted Wall001 to feed the loop more smoothly on plunge
'- Rothbauerw: Check to see if all three balls are in trough before running credits.  Resets if all three balls are not in trough.
'- Rothbauerw: Reset credits when SolRelease is fired or ball is drained
'- Rothbauerw: Staged flippers added with a script option
'- Rothbauerw: added ramp ceiling on upper ramp to prevent ball loss
'- Rawd fixed a misalignment of the VR walls
'- Flupper: doubled droptarget drop and release speed, in sync with the sounds
'- Flupper: Fixed missing screws
'- Flupper: Fixed depth bias issues in the bottom right flasher dome
'- Flupper: Fixed colored dots on the bumper tops in VR
'- Flupper: Screen space reflect set default to off for better performance (in table options)
'- Flupper: new LUT text display, which is also visible in VR
'- Flupper: easter egg added, rolling text starts after 30 seconds without a key press
'- Flupper: Fixed 2 mixed up backglass lights (500k and quick multiball) in VR
'- Flupper: Better suggested VR settings in the second settings screenshot
'- Flupper: added "making off" PDF, which is essentially the same as the WIP thread on VPForums

'******************************************************
'           STANDARD DEFINITIONS
'******************************************************

If FlasherIntensity < 1 Then FlasherIntensity = 1
If FlasherIntensity > 11 Then FlasherIntensity = 11
FlasherIntensity = FlasherIntensity / 11

If FlasherIntensityDome < 1 Then FlasherIntensityDome = 1
If FlasherIntensityDome > 11 Then FlasherIntensityDome = 11
FlasherIntensityDome = 0.5 * FlasherIntensityDome / 11

If EnableRollingText Then
  Text011.TimerEnabled = True
Else
  Text011.TimerEnabled = False
End If

Dim BallMass ,BallSize
Ballsize = 50
Ballmass = 1.0

Const UseSolenoids=2
Const UseLamps=0
Const UseGI=0


' using table width and height in script slows down the performance
Dim tablewidth: tablewidth = Table1.width
Dim tableheight: tableheight = Table1.height
Dim DesktopMode: DesktopMode = Table1.ShowDT
Dim GIStateLower, GIStateUpper, GInrLower, GInrUpper
GIStateLower = False : GiStateUpper = False

'******************************************************
'           TABLE INIT
'******************************************************

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01560000", "S11.VBS", 3.26

Const cGameName="whirl_l3"

Dim bsTrough, bsSaucer, trCellar
Dim ttMiddleSpinner, ttLeftSpinner, ttRightSpinner
Dim dtT, dtL
Dim WWBall1, WWBall2, WWBall3

Sub Table1_Init
  vpmInit Me
  On Error Resume Next
    With Controller
      .GameName = cGameName
       If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
      .SplashInfoLine = "Whirlwind Williams 1990"
      .HandleMechanics=0
      .HandleKeyboard=0
      .ShowDMDOnly=1
      .ShowFrame=0
      .ShowTitle=0
      .hidden = desktopmode


       On Error Resume Next
       .Run GetPlayerHWnd
       If Err Then MsgBox Err.Description
       On Error Goto 0
    End With
  On Error Goto 0

  InitLights InsertOn, InsertLights

  Dim obj, multiply : multiply = 1.5
  For Each obj In InsertOff : obj.blenddisablelighting = 4: Next
  For Each obj in InsertLights : obj.IntensityScale = 0.666 * multiply: obj.FadeSpeedUp = 4* obj.FadeSpeedUp  * multiply : obj.FadeSpeedDown = 4*obj.FadeSpeedDown  * multiply : Next

  SwitchGI

  PinMAMETimer.Interval = PinMAMEInterval
  PinMAMETimer.Enabled  = True

  vpmNudge.TiltSwitch   = 1
  vpmNudge.Sensitivity  = 3
  vpmNudge.TiltObj    = Array(Bumper1,Bumper2,Bumper3,Bumper4,Bumper5,Bumper6,LeftSlingshot,RightSlingshot)

  '************  Trough **************************
  Set WWBall3 = sw11.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set WWBall2 = sw12.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set WWBall1 = sw13.CreateSizedballWithMass(Ballsize/2,Ballmass)

  Controller.Switch(11) = 1
  Controller.Switch(12) = 1
  Controller.Switch(13) = 1

  'Setup Top Kicker
  Set bsSaucer = New cvpmBallStack
    bsSaucer.InitSaucer sw43, 43, 110, 15
    'bsSaucer.InitExitSnd  SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
    bsSaucer.KickForceVar = 3
    bsSaucer.KickAngleVar = 2

  ' Setup Turntables
  Set ttLeftSpinner = New cvpmTurntable
    ttLeftSpinner.InitTurntable LeftSpinTrigger, 50
    ttLeftSpinner.SpinDown = 0
    ttLeftSpinner.SpinUp = 0
    ttLeftSpinner.CreateEvents "ttLeftSpinner"


  Set ttMiddleSpinner = New cvpmTurntable
    ttMiddleSpinner.InitTurntable MidSpinTrigger, 50
    ttMiddleSpinner.SpinDown = 0
    ttMiddleSpinner.SpinUp = 0
    ttMiddleSpinner.CreateEvents "ttMiddleSpinner"

  Set ttRightSpinner = New cvpmTurntable
    ttRightSpinner.InitTurntable RightSpinTrigger, 50
    ttRightSpinner.SpinDown = 0
    ttRightSpinner.SpinUp = 0
    ttRightSpinner.CreateEvents "ttRightSpinner"

  'Drop Targets
  Set dtT=New cvpmDropTarget
    dtT.InitDrop sw26,26
    'dtT.InitSnd SoundFX("",DOFDropTargets),SoundFX("W_1Bank_DropTarget_Reset_1",DOFContactors)

' Set dtL=New cvpmDropTarget
'   dtL.InitDrop Array(sw27,sw28,sw29),Array(27,28,29)
    'dtL.InitSnd SoundFX("",DOFContactors),SoundFX("W_3Bank_DropTarget_Reset_1",DOFContactors)

  Kickback.Pullback

  SpinnerStep = 10

  Controller.Switch(42) = 1
  SolRightRampEntryDown(1)

  Diverter_closed.collidable = False


  If SideRails or DesktopMode Then
    leftrail.visible = True
    RightRail.visible = True
    Lockbar.visible = True
  End If
  If VRroom > 0 then              'VRADDED
        leftrail.visible = false  'VRADDED
    RightRail.visible = false 'VRADDED
    Lockbar.visible = false   'VRADDED
  End if
  If DesktopMode Then
    table1.BloomStrength = 0.4
  Else
    If WindowWidth < 2000 Then table1.BloomStrength = 0.2
    EMReel001.visible = false : EMReel002.visible = false : EMReel003.visible = false: EMReel004.visible = false
    EMReel005.visible = false : EMReel006.visible = false : EMReel007.visible = false: EMReel008.visible = false
  End If

  table1.ColorGradeImage = luts(lutpos)

  If StagedFlipperMod = 1 Then
    keyStagedFlipperR = KeyUpperRight ' change to RightMagnaSave if you want to use MagnaSave instead
  End If

'VRADDED ******************************************************
  If VRRoom > 0 Then
    SetBackglass
  End If
'END VRADDED **************************************************

End Sub

Sub Table1_Paused:Controller.Pause = 1:End Sub
Sub Table1_unPaused:Controller.Pause = 0:End Sub
Sub Table1_exit():Controller.Pause = False:Controller.Stop:End Sub

' *******************************
' ***** Solenoid Call backs *****
' *******************************
'
' Whirlwind features some unusual callback numbering, probably due to the manual using C and A encoding with an aux board
' Conversion rules (when referring to the manual):
' xxA = xx (xx = 01 - 08)
' xxC : 01C = 25 *** 02C = 26 *** 03C = 27 *** 04C = 28 *** 05C = 29 *** 06C = 30 *** 07C = 31 *** 08C = 32
' 09 - 22 = 09 - 22 (no change)
' above 23 : 023 = 37 *** 024 = 38 *** 025 = 39 *** 026 = 40 *** 027 = 41


SolCallback(1) = "SolOutHole"       'Solenoid 01A - Outhole Kicker
SolCallback(2) = "SolRelease"       'Solenoid 02A - Shooter Lane Feeder
SolCallback(3) = "SolRightRampEntryLifter"  'Solenoid 03A - Right Ramp Entry Lifter
SolCallback(4) = "SolLeftLockingKickback" 'Solenoid 04A - Left Locking Kickback
SolCallback(5) = "SolTopKickout"      'Solenoid 05A - Top Right Eject
SolCallback(6) = "SolKnocker"         'Solenoid 06A - Knocker (Backbox)
SolCallback(7) = "SolLeftDTUp"        'Solenoid 07A - 3-Bank Drop Target Reset
SolCallback(8) = "SolTopDTUp"         'Solenoid 08A - 1-Bank Drop Target Reset
SolCallback(11) = "SolUpperGI"        'Solenoid 11 - Upper Playfield G.I. Relay
SolCallback(12) = "SolACselectRelay"    'Solenoid 12 - Solenoid A/C Select Relay
SolCallback(13) = "SolRampDiverter"     'Solenoid 13 - Ramp Diverter
SolCallback(14) = "SolCellar"       'Solenoid 14 - Cellar Kickback
SolCallback(16) = "SolLowerGI"        'Solenoid 16 - Lower Playfield & Insert/Backbox GI Relays
SolCallback(22) = "SolRightRampEntryDown" 'Solenoid 22 - Right Ramp Entry Down
SolCallback(41) = "SolSpinWheelsMotor"    'Solenoid 27 - Wheels Spinner Motor

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"
SolCallback(sURFlipper) = "SolURFlipper"

' Setup Flasher Callbacks
SolCallback(25) =  "FlasherSet 25,"     'Solenoid 01C - Bottom Right Flasher
SolCallback(26) =  "FlasherSet 26,"     'Solenoid 02C - Spinner Flasher
SolCallback(27) =  "FlasherSet 27,"     'Solenoid 03C - Ramp Top / Upper Jets Flashers
SolCallback(28) =  "Setlamp 128,"     'Solenoid 04C - Ramp Upper Middle & Million Flashers
SolCallback(29) =  "FlasherSet 29,"     'Solenoid 05C - Ramp Lower Middle / Lower Jets Flashers
SolCallback(30) =  "Setlamp 130,"     'Solenoic 06C - Ramp Bottom (Low Playfield) Flasher
SolCallback(31) =  "FlasherSet 31,"     'Solenoid 07C - 3-Bank Drop Target & Middle Target Flashers
SolCallback(32) =  "FlasherSet 32,"     'Solenoid 08C - Million + & Compass Flashers

'backwall Lightning
SolCallback(37) = "Setlamp 137,"    'Solenoid 23 - Left Lightning (BP) Flashers
SolCallback(39) = "Setlamp 139,"    'Solenoid 25 - Middle Thunder (BP) Flasher
SolCallback(40) = "Setlamp 140,"    'Solenoid 26 - Right Thunder (BP) Flashers

'backwall VR_fan
SolCallback(38)       = "SolBackwallFan"  'Solenoid 24 - Blower Motor (Backbox)


Sub FlasherSet(nr, value)
  Select Case nr
    Case 25
      If value  then
        '***
        Call Sound_Flasher_Relay(SoundOn, F25Position)
        Flasher1c_enabled = true
        ObjLevel(9) = 1 : FlasherFlash9_Timer
      Else
        Call Sound_Flasher_Relay(SoundOff, F25Position)
        Flasher1c_enabled = false
      End If
    Case 26
      Light058.state = value
      Light059.state = value
      Light004.state = value
      If value  then
        Call Sound_Flasher_Relay(SoundOn, F26Position)
      Else
        Call Sound_Flasher_Relay(SoundOff, F26Position)
      End If
    Case 27
      Light028.state = value
      Light027.state = value
      If value  then
        '***
        Call Sound_Flasher_Relay(SoundOn, F27Position)
        Flasher3c_enabled = true
        ObjLevel(4) = 1 : FlasherFlash4_Timer
      Else
        Call Sound_Flasher_Relay(SoundOff, F27Position)
        Flasher3c_enabled = false
      End If
      Case 29
      Light017.state = value
      Light035.state = value
      If value  then
        '***
        Call Sound_Flasher_Relay(SoundOn, F29Position)
        Flasher5c_enabled = true
        ObjLevel(2) = 1 : FlasherFlash2_Timer
      Else
        Call Sound_Flasher_Relay(SoundOff, F29Position)
        Flasher5c_enabled = false
      End If
    Case 31
      Light034.state = value
      Light003.state = value

      If value  then
        Call Sound_Flasher_Relay(SoundOn, F31Position)
        Primitive115.BlendDisableLighting = 200000
        Primitive056.BlendDisableLighting = 2000
      Else
        Call Sound_Flasher_Relay(SoundOff, F31Position)
        Primitive115.BlendDisableLighting = 0
        Primitive056.BlendDisableLighting = 0
      End If
    Case 32
      Light012.state = value
      Light033.state = value
      If value  then
        Call Sound_Flasher_Relay(SoundOn, F32Position)
      Else
        Call Sound_Flasher_Relay(SoundOff, F32Position)
      End If
    End Select
End Sub

Dim Flasher6c_enabled, Flasher5c_enabled, Flasher4c_enabled, Flasher3c_enabled, Flasher1c_enabled, Flasher23_enabled, Flasher25_enabled, Flasher26_enabled
Flasher6c_enabled = False : Flasher5c_enabled = False : Flasher4c_enabled = False : Flasher3c_enabled = False
Flasher1c_enabled = False : Flasher23_enabled = False : Flasher25_enabled = False : Flasher26_enabled = False

Sub SetLamp(nr, value)
  Select Case nr
    Case 128
      If value  then
        '***
        Call Sound_Flasher_Relay(SoundOn, F128Position)
        ObjLevel(3) = 1 : FlasherFlash3_Timer
        Flasher4c_enabled = true
      Else
        Call Sound_Flasher_Relay(SoundOff, F128Position)
        Flasher4c_enabled = false
      End If
    Case 130
      If value  then
        '***
        Call Sound_Flasher_Relay(SoundOn, F130Position)
        Flasher6c_enabled = true
        ObjLevel(1) = 1 : FlasherFlash1_Timer
      Else
        Call Sound_Flasher_Relay(SoundOff, F130Position)
        Flasher6c_enabled = false
      End If
    Case 137
      If value  then
        Call Sound_Flasher_Relay(SoundOn, GIUpperPosition)
        Flasher23_enabled = True
      Else
        Call Sound_Flasher_Relay(SoundOff, GIUpperPosition)
        Flasher23_enabled = False
      End If
    Case 139
      If value  then
        Call Sound_Flasher_Relay(SoundOn, GIUpperPosition)
        Flasher25_enabled = True
      Else
        Call Sound_Flasher_Relay(SoundOff, GIUpperPosition)
        Flasher25_enabled = False
      End If
    Case 140
      If value  then
        Call Sound_Flasher_Relay(SoundOn, GIUpperPosition)
        Flasher26_enabled = True
              Else
        Call Sound_Flasher_Relay(SoundOff, GIUpperPosition)
        Flasher26_enabled = False
      End If
    End Select
End Sub

'******************************************************
'           TROUGH
'******************************************************

' switch order in manual is incorrect, the script below adds a second ball
' when pressing start when ball 2 is in the shooter lane
'Sub sw13_Hit():Controller.Switch(13) = 1:UpdateTrough:End Sub
'Sub sw13_UnHit():Controller.Switch(13) = 0:UpdateTrough:End Sub
'Sub sw12_Hit():Controller.Switch(12) = 1:UpdateTrough:End Sub
'Sub sw12_UnHit():Controller.Switch(12) = 0:UpdateTrough:End Sub
'Sub sw11_Hit():Controller.Switch(11) = 1:UpdateTrough:End Sub
'Sub sw11_UnHit():Controller.Switch(11) = 0:UpdateTrough:End Sub

Sub sw13_Hit():Controller.Switch(11) = 1:UpdateTrough:End Sub
Sub sw13_UnHit():Controller.Switch(11) = 0:UpdateTrough:End Sub
Sub sw12_Hit():Controller.Switch(12) = 1:UpdateTrough:End Sub
Sub sw12_UnHit():Controller.Switch(12) = 0:UpdateTrough:End Sub
Sub sw11_Hit():Controller.Switch(13) = 1:UpdateTrough:End Sub
Sub sw11_UnHit():Controller.Switch(13) = 0:UpdateTrough:End Sub


Sub UpdateTrough()
  UpdateTroughTimer.Interval = 300
  UpdateTroughTimer.Enabled = 1
End Sub

Sub UpdateTroughTimer_Timer()
  If sw13.BallCntOver = 0 Then sw12.kick 60, 9
  If sw12.BallCntOver = 0 Then sw11.kick 60, 9
  Me.Enabled = 0
End Sub

'******************************************************
'         DRAIN & RELEASE
'******************************************************

Sub sw10_Hit() 'Drain
  RandomSoundOutholeHit()
  Controller.Switch(10) = 1
End Sub

Sub sw10_UnHit()  'Drain
    Controller.Switch(10) = 0
    ResetRollingTextTimer
End Sub

Sub SolOutHole(enabled)
  If enabled Then
    UpdateTrough
    sw10.kick 60, 20
    RandomSoundOutholeKicker()
  End If
End Sub

Sub SolRelease(enabled)
  If enabled Then
    sw13.kick 60, 9
    RandomSoundShooterFeeder()
    If EnableRollingText Then
      ResetRollingTextTimer
    End If
  End If
End Sub

'RandomSoundBottomArchBallGuideSoftHit - Soft Bounces
Sub Wall015_Hit() : RandomSoundBottomArchBallGuideSoftHit() : End Sub
Sub Wall018_Hit() : RandomSoundBottomArchBallGuideSoftHit() : End Sub

'RandomSoundBottomArchBallGuideHardHit - Hard Hit
Sub Wall016_Hit() : RandomSoundBottomArchBallGuideHardHit() : End Sub
Sub Wall017_Hit() : RandomSoundBottomArchBallGuideHardHit() : End Sub

'RandomSoundFlipperBallGuide
Sub Wall11_Hit() : RandomSoundFlipperBallGuide() : End Sub
Sub Wall21_Hit() : RandomSoundFlipperBallGuide() : End Sub

'RandomSoundWireformAntiRebountRail
'Sub rail001_Hit() : RandomSoundWireformAntiRebountRail() : End Sub
'Sub rail002_Hit() : RandomSoundWireformAntiRebountRail() : End Sub


'**********************************************************************************************************
'Solenoid Controlled toys
'**********************************************************************************************************

'Lock ball
Sub SolLeftLockingKickback(enabled)
  if enabled Then
    RandomSoundLockingKickerSolenoid()
    Kickback.Fire
'   Kicker3.kick 0,30
'   Kicker3.enabled = False
'   Kicker2.kick 0,30
'   Kicker2.enabled = False
'   Kicker1.kick 0,30
  Else
    Kickback.Pullback
  end if
End Sub


'**************** Cellar Kickout ****************
dim cellarkick_vel,cellarkick_vel_variance
dim cellarkick_angle, cellarkick_angle_variance

cellarkick_vel = 38     'velocity
cellarkick_vel_variance = 0.5 'velocity variance
cellarkick_angle = Loopy_Left.objrotz - 5 'adjust objrotz on loop mesh to adjust kickout direction
cellarkick_angle_variance = 0.0 'Angle variance


Sub sw19_hit()
  vpmtimer.PulseSw 19
End Sub

Sub Sw20_Hit()
  Controller.Switch(20) = 1
  scoopKickerOverflow.Enabled = True
End Sub

Sub Sw20_UnHit()
  Controller.Switch(20) = 0
  scoopKickerOverflow.Enabled = False
End Sub

sub SolCellar(enabled)
  if enabled Then
    scoopKickerOverflow.Kick KickoutVariance(cellarkick_angle,cellarkick_angle_variance), KickoutVariance(cellarkick_vel, cellarkick_vel_variance)
    sw20.Kick KickoutVariance(cellarkick_angle,cellarkick_angle_variance), KickoutVariance(cellarkick_vel, cellarkick_vel_variance)
    if CellarOpt = 1 then
      cellarwall.collidable = 1
      vpmTimer.AddTimer 200,"cellarwall.collidable = 0'"
    Else
            cellarwall.collidable = 0
    end if
    SoundCellerKickout()
    SoundCellarKickoutBallDrop()
  end if
end sub

Function KickoutVariance(aNumber, aVariance)  'strength, variance
  KickoutVariance = aNumber + ((Rnd*2)-1)*aVariance
End Function


'Upper PF GI
Sub SolUpperGI(enabled)
  'debug.print "SolUpperGI " & enabled
  If Enabled Then
    Sound_UpperGI_Relay(SoundOn)
  Else
    Sound_UpperGI_Relay(SoundOff)
  End If
  GIStateUpper = Enabled
  If Enabled Then GInrUpper = 0 Else GInrUpper = 1 : End If
  SwitchGI
End Sub


'Lower PF GI
Sub SolLowerGI(enabled)
  'debug.print "SolLowerGI " & enabled
  If Enabled Then
    Sound_LowerGI_Relay(SoundOn)
    If VRRoom > 0 Then                      'VRAdded***
      dim xx
      For each xx in VRBackglassGI:xx.visible = 0: Next
      For each xx in VRBGGILetters:xx.visible = 1: Next
        If VRFlashingBackglass = 1 Then
          BGFlashersFast.enabled = False
          BGFlashersMid.enabled = False
          BGFlashersSlow.enabled = False
        End If
    End If                            'VRAdded***
'   dim xxxx
'   For each xxxx in LowerGI:xxxx.State = 0: Next
      'TextBox002.text = "Loff"
  Else
    Sound_LowerGI_Relay(SoundOff)
    If VRRoom > 0 Then                      'VRAdded**
      For each xx in VRBackglassGI:xx.visible = 1: Next
      For each xx in VRBGGILetters:xx.visible = 0: Next
        If VRFlashingBackglass = 1 Then
          BGFlashersFast.enabled = True
          BGFlashersMid.enabled = True
          BGFlashersSlow.enabled = True
        End If
    End If                            'VRAdeed**
'   For each xxxx in LowerGI:xxxx.State = 1: Next\ 
    'TextBox002.text = "Lon"
  End If
  GIStateLower = Enabled
  If Enabled Then GInrLower = 0 Else GInrLower = 1 : End If
  SwitchGI
End Sub

Sub SwitchGI()
  Dim nr, obj
  If GIStateUpper Then
    ' *** Upper GI is OFF ***
    coll028.image = "TexGIoff2" : coll029.image = "TexGIoff2"
    bumperbulb001.image = "grey" : bumperbulb001.material = "bumperbulbGI" : bumperbulb001.blenddisablelighting = 0 : Light001.state = 0
    for each nr in InsertLightsUpperGI : nr.intensityscale = 1 : Next
    sw26.blenddisablelighting = 0.5 : Flasherbase003.Blenddisablelighting = 1
    If GIStateLower Then
      ' *** Lower GI is OFF ***
      playfield_render.image = "texGIoff0" : Primitive112.image = "texGIoff0" : Primitive113.image = "texGIoff0" : Primitive114.image = "texGIoff0" : Primitive114.blenddisablelighting = 0.2
      playfield_render.blenddisablelighting = 0.3 '0.4 0.1
      playfield_render.visible = 1
      coll026.image = "texGIoff3" : coll027.image = "texGIoff3"
      For each nr in tex1objects: nr.image = "texGIoff1" : nr.material = "collection1dark" : Next
      MiddleWheel.blenddisablelighting = 0 : RightWheel.blenddisablelighting = 0 : LeftWheel.blenddisablelighting = 0
      Flasherbase1.blenddisablelighting = 0.1 : Flasherbase2.blenddisablelighting = 0.1 : Flasherbase3.blenddisablelighting = 0.1 : Flasherbase4.blenddisablelighting = 0.1
      FlasherBase9.blenddisablelighting = 0
      for each nr in InsertLightsLowerGI : nr.intensityscale = 1: Next
      batlightleft.state = 0 : batlightright.state = 0 : URightBat.blenddisablelighting = 0.1
      sw29p.blenddisablelighting = 0.5 : sw28p.blenddisablelighting = 0.5 : sw27p.blenddisablelighting = 0.5
      Flasherbase001.blenddisablelighting = 1.2 : Flasherbase002.blenddisablelighting = 0.8
      For Each obj In BallReflectionLights : obj.state = false: Next
      leftsling.image = "slingtexgioff" : rightsling.image = "slingtexgioff"
    Else
      ' *** Lower GI is ON ***
      playfield_render.image = "texGIonlower0" : Primitive112.image = "texGIonlower0" : Primitive113.image = "texGIonlower0" : Primitive114.image = "texGIonlower0" : Primitive114.blenddisablelighting = 1
      playfield_render.blenddisablelighting = 0.6
      playfield_render.visible = 1
      coll026.image = "texGIonlower3" : coll027.image = "texGIonlower3"
      For each nr in tex1objects: nr.image = "texGIonlower1" : nr.material = "collection1" : Next
      MiddleWheel.blenddisablelighting = 0.5 : RightWheel.blenddisablelighting = 0.5 : LeftWheel.blenddisablelighting = 0.5
      Flasherbase1.blenddisablelighting = 1 : Flasherbase2.blenddisablelighting = 0.4 : Flasherbase3.blenddisablelighting = 0.1 : Flasherbase4.blenddisablelighting = 0.1
      FlasherBase9.blenddisablelighting = 1
      for each nr in InsertLightsLowerGI : nr.intensityscale = 0.4 : Next
      batlightleft.state = 1 : batlightright.state = 1 : URightBat.blenddisablelighting = 0.4
      sw29p.blenddisablelighting = 1 : sw28p.blenddisablelighting = 1 : sw27p.blenddisablelighting = 1
      Flasherbase001.blenddisablelighting = 10 : Flasherbase002.blenddisablelighting = 1.4
      For Each obj In BallReflectionLights : obj.state = True: Next
      leftsling.image = "slingtexgion" : rightsling.image = "slingtexgion"
    End If
  Else
    ' *** Upper GI is ON ***
    coll028.image = "TexGIon2" : coll028.image = "TexGIon2"
    bumperbulb001.image = "bumperbulbon" : bumperbulb001.material = "bumperbulb" : bumperbulb001.blenddisablelighting = 0 : Light001.state = 1
    for each nr in InsertLightsUpperGI : nr.intensityscale = 0.4 : Next
    sw26.blenddisablelighting = 1 : Flasherbase003.Blenddisablelighting = 4
    If GIStateLower Then
      ' *** Lower GI is OFF ***
      playfield_render.image = "texGIonupper0" : Primitive112.image = "texGIonupper0" : Primitive113.image = "texGIonupper0" : Primitive114.image = "texGIonupper0" : Primitive114.blenddisablelighting = 0.4
      playfield_render.blenddisablelighting = 0.7 : playfield_render.visible = 1
      coll026.image = "texGIonupper3" : coll027.image = "texGIonupper3"
      For each nr in tex1objects: nr.image = "texGIonupper1" : nr.material = "collection1" : Next
      MiddleWheel.blenddisablelighting = 0 : RightWheel.blenddisablelighting = 0 : LeftWheel.blenddisablelighting = 0
      Flasherbase1.blenddisablelighting = 0.1 : Flasherbase2.blenddisablelighting = 0.1 : Flasherbase3.blenddisablelighting = 1 : Flasherbase4.blenddisablelighting = 1
      FlasherBase9.blenddisablelighting = 0.1
      for each nr in InsertLightsLowerGI : nr.intensityscale = 0.8 : Next
      batlightleft.state = 0 : batlightright.state = 0 : URightBat.blenddisablelighting = 0.1
      sw29p.blenddisablelighting = 0.5 : sw28p.blenddisablelighting = 0.5 : sw27p.blenddisablelighting = 0.5
      Flasherbase001.blenddisablelighting = 1.2 : Flasherbase002.blenddisablelighting = 0.8
      For Each obj In BallReflectionLights : obj.state = false: Next
      leftsling.image = "slingtexgioff" : rightsling.image = "slingtexgioff"
    Else
      ' *** Lower GI is ON ***
      playfield_render.visible = 0 : Primitive112.image = "texGIon0" : Primitive113.image = "texGIon0" : Primitive114.image = "texGIon0" : Primitive114.blenddisablelighting = 1
      coll026.image = "texGIon3" : coll027.image = "texGIon3"
      For each nr in tex1objects: nr.image = "texGIon1" : nr.material = "collection1" : Next
      MiddleWheel.blenddisablelighting = 0.5 : RightWheel.blenddisablelighting = 0.5 : LeftWheel.blenddisablelighting = 0.5
      Flasherbase1.blenddisablelighting = 1 : Flasherbase2.blenddisablelighting = 0.8 : Flasherbase3.blenddisablelighting = 1 : Flasherbase4.blenddisablelighting = 1
      FlasherBase9.blenddisablelighting = 1
      for each nr in InsertLightsLowerGI : nr.intensityscale = 0.4 : Next
      batlightleft.state = 1 : batlightright.state = 1 : URightBat.blenddisablelighting = 0.4
      sw29p.blenddisablelighting = 1 : sw28p.blenddisablelighting = 1 : sw27p.blenddisablelighting = 1
      Flasherbase001.blenddisablelighting = 10 : Flasherbase002.blenddisablelighting = 1.4
      For Each obj In BallReflectionLights : obj.state = True: Next
      leftsling.image = "slingtexgion" : rightsling.image = "slingtexgion"
    End If
  End If
  If VRRoom > 0 Then
    FlBumperRedFadeVR 3, BumperCurrent(3)
    FlBumperRedFadeVR 4, BumperCurrent(4)
    FlBumperRedFadeVR 5, BumperCurrent(5)
    FlBumperYellowFadeVR 0, BumperCurrent(0)
    FlBumperYellowFadeVR 1, BumperCurrent(1)
    FlBumperYellowFadeVR 2, BumperCurrent(2)
  Else
    FlBumperRedFade 3, BumperCurrent(3)
    FlBumperRedFade 4, BumperCurrent(4)
    FlBumperRedFade 5, BumperCurrent(5)
    FlBumperYellowFade 0, BumperCurrent(0)
    FlBumperYellowFade 1, BumperCurrent(1)
    FlBumperYellowFade 2, BumperCurrent(2)
  End If
End Sub


'Solenoid A/C Select Relay
'In its de-energized state, the Relay connects the 'circuit A power' to 16 "controlled" and "switched" solenoids (identified in the table with no suffix letter or the letter A, after the solenoid number).
'Individual solenoid operation then depends on the game program enabling the ground path for solenoid actuation via the driver transistor associated with each solenoid circuit.
'For example, the game program can actuate the Outhole Kicker solenoid (sol. 01A), via the driver transistor Q33.
'
'"When the game program determines that the Solenoid A/C Select Relay (sol. 12) must be energized, the relay connects 'circuit C power' to eight group C solenoids (01C through 08C).
'Now, driver transistor Q33 can actuate the Bottom Right Flasher circuit (sol. 01C) which has two lamp circuits, one to the Insert Board and one to the playfield.
'Using this "multiplexing" technique, the same driver transistor can control actuation of two separate solenoid circuits."
Sub SolACselectRelay(enabled)
  If Enabled Then
    Sound_Solenoid_AC(CircuitC)
  Else
    Sound_Solenoid_AC(CircuitA)
  End If
End Sub

Dim DivStep : DivStep = 0
Dim DivOpen : DivOpen = False

Dim SolRampDiverterFlag
Sub SolRampDiverter(enabled)
  if enabled Then
    SolRampDiverterFlag = 1
    RandomSoundRampDiverterDivert()
    RandomSoundRampDiverterHold(SoundOn)
    diverter_closed.collidable = True 'Enable Divert ball to lock
    diverter.collidable = False       'Disable Divert ball back to playfield
    DivOpen = True
    DiverterTimer.enabled = True
  Else
    SolRampDiverterFlag = 0
    RandomSoundRampDiverterBack()
    RandomSoundRampDiverterHold(SoundOff)
    diverter_closed.collidable = False 'Disable divert ball to lock
    diverter.collidable = True         'Enable Divert ball back to playfield
    DivOpen = False
    DiverterTimer.enabled = True
  end if
End Sub

Sub DiverterTimer_Timer()
  DivStep = Divstep + 1
  If DivOpen Then
    coll024.RotZ = StUp(Divstep) * 19/15 + 1
  Else
    coll024.RotZ = StDn(Divstep) * 19/15 + 1
  End If
  If Divstep = 8 Then DiverterTimer.enabled = False : DivStep = 0
End Sub

Dim RampStep : RampStep = 0
Dim RampDirUp
Dim StUp(8), StDn(8)
StUp(0) = -15 : StUp(1) = -14.6 : StUp(2) = -14 : StUp(3) = -13 : StUp(4) = -10 : StUp(5) = -4 : StUp(6) = -2 : StUp(7) = -1 : StUp(8) = 0
StDn(0) = 0 : StDn(1) = -0.4 : StDn(2) = -0.8 : StDn(3) = -1.2 : StDn(4) = -2 : StDn(5) = -10 : StDn(6) = -14 : StDn(7) = -14.5 : StDn(8) = -15


sub SolRightRampEntryLifter(enabled)
  if enabled Then
    RandomSoundRightRampEntryLifter()
    RightRampDown.collidable = 0
    RightRampUp.collidable = 1
    Controller.Switch(42) = 0
    RampDirUp = True
    RampTimer.enabled = True
  Else
  end if
end Sub

Sub SolRightRampEntryDown(enabled)
  if enabled Then
    RandomSoundRightRampEntryDown()
    RightRampDown.Collidable = 1
    RightRampup.Collidable = 0
    Controller.Switch(42) = 1
    RampDirUp = False
    RampTimer.enabled = True
  Else
  end if
End Sub

Sub RampTimer_Timer()
  RampStep = Rampstep + 1
  If RampDirUp Then
    coll022.RotX = StUp(Rampstep)
    'Primitive094.x = 24.7572 - Rampstep * (24.75 - 6.38) /8
    'Primitive094.y = -25.54732 - Rampstep * (33.71 - 25.55) / 8
    'Primitive094.RotZ = 2 - 2 * RampStep / 8
    UpdateMaterial "Flapshadow",0,0,0,0,0,0,(0.9 - 0.6 * RampStep/8),RGB(255,255,255),0,0,False,True,0,0,0,0
  Else
    coll022.RotX = StDn(Rampstep)
    'Primitive094.x = 6.38 + Rampstep * (24.75 - 6.38) /8
    'Primitive094.y = -33.71 + Rampstep * (33.71 - 25.55) / 8
    'Primitive094.RotZ = 2 * RampStep / 8
    UpdateMaterial "Flapshadow",0,0,0,0,0,0,(0.3 + 0.6 * RampStep/8),RGB(255,255,255),0,0,False,True,0,0,0,0
  End If
  If RampStep = 8 Then RampTimer.enabled = False : RampStep = 0
End Sub

Sub TDBallLift(ball, pos)
  If InRect(ball.x, ball.y, 665, 728, 750, 512, 806, 534, 725, 753) Then
    ball.z = 110 * Pos - (110 * Pos * (1 - Distance(ball.x,ball.y,787,505)/250)) + 25
    'ball.velz = 10
  End If
End Sub

'Function Distance(ax,ay,bx,by)
' Distance = SQR((ax - bx)^2 + (ay - by)^2)
'End Function



'******************************************************
'         SPINNING DISCS
'******************************************************

Sub SolSpinWheelsMotor(enabled)
  If enabled Then
    ttMiddleSpinner.MotorOn = True
    ttLeftSpinner.MotorOn = True
    ttRightSpinner.MotorOn = True
    SoundSpinWheelsMotorsTransients(RampUp)
    SoundSpinWheelsMotorsLoop(SoundOn)
  ' ttMiddleSpinner.Speed = 25
  ' ttLeftSpinner.Speed = -25
  ' ttRightSpinner.Speed = -25
    SpinnerMotorOff = False
    SpinnerTimer.Interval = 10
    SpinnerTimer.enabled = True
  Else
    SpinnerMotorOff = True

    SoundSpinWheelsMotorsLoop(SoundOff)
    'If spinning disc at top speed (MaxTTSpeed) then
    If TTSpeed = MaxTTSpeed Then
      SoundSpinWheelsMotorsTransients(RampDownSlow)
    Else
      SoundSpinWheelsMotorsTransients(RampDownFast)
    End If
  End If
End Sub

'Spinning Discs Animation TimerSoundSpinWheelsMotorsLoop
Dim SpinnerMotorOff, SpinnerStep, ss
Dim MaxSpinnerStep, SpinnerRampUp, SpinnerRampDown
MaxSpinnerStep =20
SpinnerRampup = 2
SpinnerRampDown = 0.35

Dim MaxTTSpeed, TTRampUp, TTRampDown, TTSpeed
MaxTTSpeed = 25
TTRampUp = (SpinnerRampUp/MaxSpinnerStep)*MaxTTSpeed
TTRampDown = (SpinnerRampDown/MaxSpinnerStep)*MaxTTSpeed

Sub SpinnerTimer_Timer()

  If Not(SpinnerMotorOff) Then
    If SpinnerStep < MaxSpinnerStep Then
      SpinnerStep = SpinnerStep + SpinnerRampUp
      TTSpeed = TTSpeed + TTRampUp
    Else
      SpinnerStep = MaxSpinnerStep
      TTSpeed = MaxTTSpeed
    End If
  Else
    if SpinnerStep < 0 Then
      SpinnerStep = 0
      TTSpeed = 0
      SpinnerTimer.enabled = False
      ttleftSpinner.speed = 0
      ttLeftSpinner.MotorOn = False
      ttMiddleSpinner.Speed = 0
      ttMiddleSpinner.MotorOn = False
      ttRightSpinner.speed = 0
      ttRightSpinner.MotorOn = False
    Else
      'slow the rate of spin by decreasing rotation step
      SpinnerStep = SpinnerStep - SpinnerRampDown
      TTSpeed = TTSpeed - TTRampDown
    End If
  End If

  ttMiddleSpinner.Speed = TTSpeed
  ttLeftSpinner.Speed = -TTSpeed
  ttRightSpinner.Speed = -TTSpeed
  LeftWheel.RotY = ss
  MiddleWheel.RotZ = ss
  RightWheel.RotY = ss
  ss = ss + SpinnerStep
  If ss > 360 Then ss = ss - 360
End Sub

'******************************************************
'           FAN
'******************************************************

Sub SolBackwallFan(enabled)
  If enabled Then
    SoundBackwallFanBlowerTransients(RampUp)
    SoundBackwallFanBlowerLoop(SoundOn)
    'Debug.Print "SolBackwallFan = Enabled"
    If VRroom > 0 then BlowerTimer.enabled = true 'VRADDED

  Else
    SoundBackwallFanBlowerTransients(RampDown)
    SoundBackwallFanBlowerLoop(SoundOff)
    'Debug.Print "SolBackwallFan = Disabled"
        If VRroom > 0 then BlowerRampDownTimer.enabled = true  'VRADDED
  end If
End Sub

'VRADDED ***************************************************************
Dim VRFanSpeed
VRFanSpeed = 15

sub BlowerTimer_Timer()
VR_Fan.RotY = VR_Fan.RotY + VRFanSpeed
end sub

sub BlowerRampDownTimer_Timer()
BlowerTimer.interval = BlowerTimer.interval + 5
VRFanSpeed = VRFanSpeed - 1

if BlowerTimer.interval > 100 then
BlowerTimer.enabled = false
BlowerTimer.interval = 11
VRFanSpeed = 15
BlowerRampDownTimer.enabled = false
exit sub
end If
End Sub
' End VRADDED ************************************************************


'**********************************************************************************************************
'Key handling
'**********************************************************************************************************




Sub Table1_KeyDown(ByVal KeyCode)
' If EnableRollingText Then
'   ResetRollingTextTimer
' End If
  'If KeyDownHandler(keycode) Then Exit Sub
  If keycode = StartGameKey Then
     soundStartButton()
    If VRroom >0 then VR_StartButton.y = VR_StartButton.y -5  'VRADDED StartButton
  End if
  If keycode = PlungerKey Then Plunger.Pullback : SoundPlungerPull()

  If keycode = PlungerKey and VRRoom > 0 Then  'VRADDED
    TimerVRPlunger.Enabled = True     'VRADDED
    TimerVRPlunger2.Enabled = False     'VRADDED
  End If                    'VRADDED

  If keycode = LeftTiltKey Then
    Nudge 90, 4
    SoundNudgeLeft()
    End if
  If keycode = RightTiltKey Then
    Nudge 270, 4
    SoundNudgeRight()
    End if
  If keycode = CenterTiltKey Then
    Nudge 0, 5
    SoundNudgeCenter()
    End if
  If keycode = LeftFlipperKey Then
    FlipperActivate LeftFlipper, LFPress
        if VRroom > 0 then VR_FB_Left.X = VR_FB_Left.X +10 'VRADDED
        End if

  If keycode = RightFlipperKey Then
    FlipperActivate RightFlipper, RFPress
    If StagedFlipperMod <> 1 Then FlipperActivate RightFlipper1, RFPress1
    if VRroom > 0 then VR_FB_Right.X = VR_FB_Right.X - 10 'VRADDED
  End if

  If StagedFlipperMod = 1 Then
    If keycode = keyStagedFlipperR Then FlipperActivate RightFlipper1, RFPress1
  End If

  If keycode = RightMagnaSave and EnableMagnasave = 1 then
    ResetRollingTextTimer
    lutpos = lutpos + 1 : If lutpos > ubound(luts) Then lutpos = 0 : end if
        call myChangeLut
    playsound "LutChange"
  End if

  If keycode = LeftMagnaSave and EnableMagnasave = 1 then
    ResetRollingTextTimer
    lutpos = lutpos - 1 : If lutpos < 0 Then lutpos = ubound(luts) : end if
        call myChangeLut
    playsound "LutChange2"
    end if


  If keycode = keyInsertCoin1 or keycode = keyInsertCoin2 or keycode = keyInsertCoin3  Then
    Select Case Int(rnd*3)
      Case 0: PlaySoundAtLevelStatic ("Coin_In_1"), CoinSoundLevel, sw10
      Case 1: PlaySoundAtLevelStatic ("Coin_In_2"), CoinSoundLevel, sw10
      Case 2: PlaySoundAtLevelStatic ("Coin_In_3"), CoinSoundLevel, sw10
    End Select
  End If

  If vpmKeyDown(keycode) Then Exit Sub
End Sub



Sub Table1_KeyUp(ByVal KeyCode)
  'If KeyUpHandler(keycode) Then Exit Sub
  If keycode = PlungerKey Then
    Plunger.Fire
    If controller.switch(59) = True Then  'If true then ball in shooter lane, else no ball is shooter lane
      SoundPlungerPullStop()
      SoundPlungerReleaseBall()     'Plunger release sound when there is a ball in shooter lane
    Else
      SoundPlungerPullStop()
      SoundPlungerReleaseNoBall()     'Plunger release sound when there is no ball in shooter lane
    End If
  End If

  If keycode = LeftFlipperKey Then
    FlipperDeActivate LeftFlipper, LFPress
    if VRRoom >0 then VR_FB_Left.X = VR_FB_Left.X -10 'VRADDED
  End If
  If keycode = RightFlipperKey Then
    FlipperDeActivate RightFlipper, RFPress
    If StagedFlipperMod <> 1 Then FlipperDeActivate RightFlipper1, RFPress1
    if VRRoom >0 then VR_FB_Right.X = VR_FB_Right.X +10 'VRADDED
  End If

  If StagedFlipperMod = 1 Then
    If keycode = keyStagedFlipperR Then FlipperDeActivate RightFlipper1, RFPress1
  End If

'VRADDED *******************************************************************************************************************************
  If keycode = PlungerKey and VRRoom > 0 Then
    TimerVRPlunger.Enabled = False
    TimerVRPlunger2.Enabled = True
    TAFPlungerNew.Y = 1097.224
    TAFPlungerNew1.Y = 1097.224
  End If

  If Keycode = StartGameKey Then
    If VRroom >0 then VR_StartButton.y = VR_StartButton.y +5
  End If

  If vpmKeyUp(keycode) Then Exit Sub
End Sub

Sub TimerVRPlunger_Timer
  If TAFPlungerNew.Y < 1200 then
       TAFPlungerNew.Y = TAFPlungerNew.Y + 5
  End If
  If TAFPlungerNew1.Y < 1200 then
       TAFPlungerNew1.Y = TAFPlungerNew1.Y + 5
  End If
End Sub

Sub TimerVRPlunger2_Timer
  'debug.print plunger.position
  TAFPlungerNew.Y = 1106.345 + (5* Plunger.Position) -20
  TAFPlungerNew1.Y = 1106.345 + (5* Plunger.Position) -20
End Sub
'END VRADDED ******************************************************************************************************************************


Const ReflipAngle = 20
Const QuickFlipAngle = 20

Sub SolLFlipper(Enabled)
   If Enabled Then
    LF.Fire   'LeftFlipper.RotateToEnd
    If LeftFlipper.currentangle < LeftFlipper.endangle + ReflipAngle Then
      'Play partial flip sound and stop any flip down sound
      'Debug.print "Flip Reflip"
      StopAnyFlipperLowerLeftDown()
      RandomSoundLowerLeftReflip()
    Else
      'Debug.print "LeftFlipper.currentangle = " &LeftFlipper.currentangle
      'Debug.print "LeftFlipper.startangle = " &LeftFlipper.startangle
      'Debug.print "LeftFlipper.endangle = " &LeftFlipper.endangle
      'LeftFlipper.RotateToEnd
      'Play full flip sound
      'Debug.print "Flip Up"
      If BallNearLF = 0 Then
        SoundFlipperUpAttackLeft()
        RandomSoundFlipperLowerLeftUpFullStroke()
      End If
      If BallNearLF = 1 Then
        Select Case Int(Rnd*2)+1
          Case 1 : RandomSoundFlipperLowerLeftUpDampenedStroke()
          Case 2 : RandomSoundFlipperLowerLeftUpFullStroke()
        End Select
      End If
    End If
   Else
    LeftFlipper.RotateToStart
    If LeftFlipper.currentangle < LeftFlipper.startAngle - 5 Then
      'Play flip down sound
      'Debug.print "Flip Down"
      RandomSoundLowerLeftDown()
    End If
    If LeftFlipper.currentangle < LeftFlipper.startAngle + QuickFlipAngle and LeftFlipper.currentangle <> LeftFlipper.endangle Then
      'Play quick flip sound and stop any flip up sound
      'Debug.print "Flip Quick"
      'StopAnyFlipperLowerLeftUp()
      'RandomSoundLowerLeftQuickFlipUp()
    Else
      FlipperLeftLowerHitParm = FlipperUpSoundLevel
    End If
   End If
 End Sub

Sub SolRFlipper(Enabled)
  If Enabled Then
    RF.Fire 'RightFlipper.RotateToEnd
    If RightFlipper.currentangle > RightFlipper.endangle - ReflipAngle Then
      'Play partial flip sound and stop any flip down sound
      StopAnyFlipperLowerRightDown()
      RandomSoundLowerRightReflip()
    Else
      'Play full flip sound
      'RightFlipper.RotateToEnd
      If BallNearRF = 0 Then
        SoundFlipperUpAttackRight()
        RandomSoundFlipperLowerRightUpFullStroke()
      End If

      If BallNearRF = 1 Then
        Select Case Int(Rnd*2)+1
          Case 1 : RandomSoundFlipperLowerRightUpDampenedStroke()
          Case 2 : RandomSoundFlipperLowerRightUpFullStroke()
        End Select
      End If
    End If
  Else
    RightFlipper.RotateToStart
    If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
      'Play flip down sound
      RandomSoundLowerRightDown()
    End If
    If RightFlipper.currentangle < RightFlipper.startAngle + QuickFlipAngle and RightFlipper.currentangle <> RightFlipper.endangle Then
      'Play quick flip sound and stop any flip up sound
      'StopAnyFlipperLowerRightUp()
      'RandomSoundLowerRightQuickFlipUp()
    Else
      FlipperRightLowerHitParm = FlipperUpSoundLevel
    End If
  End If
End Sub

Sub SolURFlipper(Enabled)
  If Enabled Then
    RightFlipper1.RotateToEnd
    If RightFlipper1.currentangle > RightFlipper1.endangle - ReflipAngle Then
      'Play partial flip sound and stop any flip down sound
      StopAnyFlipperUpperRightDown()
      RandomSoundUpperRightReflip()
    Else
      'Play full flip sound
      'RightFlipper1.RotateToEnd
      RandomSoundFlipperUpperRightUpFullStroke()
    End If
  Else
    RightFlipper1.RotateToStart
    If RightFlipper1.currentangle > RightFlipper1.startAngle + 5 Then
      'Play flip down sound
      RandomSoundUpperRightDown()
    End If
    If RightFlipper1.currentangle < RightFlipper1.startAngle + QuickFlipAngle and RightFlipper1.currentangle <> RightFlipper1.endangle Then
      'Play quick flip sound and stop any flip up sound
      'StopAnyFlipperUpperRightUp()
      'RandomSoundUpperRightQuickFlipUp()
    Else
      FlipperRightUpperHitParm = FlipperUpSoundLevel
    End If
  End If
End Sub


'******************************************************
'                        FLIPPER TRICKS
'******************************************************

RightFlipper.timerinterval=1
Rightflipper.timerenabled=True

sub RightFlipper_timer()
        FlipperTricks LeftFlipper, LFPress, LFCount, LFEndAngle, LFState
        FlipperTricks RightFlipper, RFPress, RFCount, RFEndAngle, RFState
  FlipperTricks RightFlipper1, RFPress1, RFCount1, RFEndAngle1, RFState1
        FlipperNudge RightFlipper, RFEndAngle, RFEOSNudge, LeftFlipper, LFEndAngle
        FlipperNudge LeftFlipper, LFEndAngle, LFEOSNudge,  RightFlipper, RFEndAngle
end sub

Dim LFEOSNudge, RFEOSNudge

Sub FlipperNudge(Flipper1, Endangle1, EOSNudge1, Flipper2, EndAngle2)
        Dim BOT, b

        If Flipper1.currentangle = Endangle1 and EOSNudge1 <> 1 Then
                EOSNudge1 = 1
                'debug.print Flipper1.currentangle &" = "& Endangle1 &"--"& Flipper2.currentangle &" = "& EndAngle2
                If Flipper2.currentangle = EndAngle2 Then
                        BOT = GetBalls
                        For b = 0 to Ubound(BOT)
                                If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper1) Then
                                        'Debug.Print "ball in flip1. exit"
                                         exit Sub
                                end If
                        Next
                        For b = 0 to Ubound(BOT)
                                If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper2) Then
                                        BOT(b).velx = BOT(b).velx / 1.7
                                        BOT(b).vely = BOT(b).vely - 1
                                end If
                        Next
                End If
        Else
                If Flipper1.currentangle <> EndAngle1 then
                        EOSNudge1 = 0
                end if
        End If
End Sub

'*****************
' Maths
'*****************
Const PI = 3.1415927

Function dSin(degrees)
        dsin = sin(degrees * Pi/180)
End Function

Function dCos(degrees)
        dcos = cos(degrees * Pi/180)
End Function

Function dAtn(degrees)
  dAtn = atn(degrees * Pi/180)
End Function

Function RndInt(min, max)
    RndInt = Int(Rnd() * (max-min + 1) + min)' Sets a random number integer between min and max
End Function

Function RndNum(min, max)
    RndNum = Rnd() * (max-min) + min' Sets a random number between min and max
End Function

'*************************************************
' Check ball distance from Flipper for Rem
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

' Used for drop targets and stand up targets
Function Atn2(dy, dx)
        dim pi
        pi = 4*Atn(1)

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

'*************************************************
' End - Check ball distance from Flipper for Rem
'*************************************************

dim LFPress, RFPress, RFPress1, LFCount, RFCount, RFCount1
dim LFState, RFState, RFState1
dim EOST, EOSA,Frampup, FElasticity, FReturn
dim RFEndAngle, RFEndAngle1, LFEndAngle

EOST = leftflipper.eostorque
EOSA = leftflipper.eostorqueangle
Frampup = LeftFlipper.rampup
FElasticity = LeftFlipper.elasticity
FReturn = LeftFlipper.return
Const EOSTnew = 0.8
Const EOSAnew = 1
Const EOSRampup = 0
Dim SOSRampup
SOSRampup = 2.5
Const LiveCatch = 16
Const LiveElasticity = 0.45
Const SOSEM = 0.815
Const EOSReturn = 0.035

LFEndAngle = Leftflipper.endangle
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
        Flipper.eostorque = EOST*EOSReturn/FReturn


        If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 0.1 Then
                Dim BOT, b
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

                If LiveCatchBounce = 0 and ball.velx * Dir > 0 Then ball.velx = 0
                ball.vely = LiveCatchBounce * (32 / LiveCatch) ' Multiplier for inaccuracy bounce
                ball.angmomx= 0
                ball.angmomy= 0
                ball.angmomz= 0
        End If
End Sub

'*******************************************
'  VPW Rubberizer by Iaakki
'*******************************************

' iaakki Rubberizer
sub Rubberizer(parm)
  if parm < 10 And parm > 2 And Abs(activeball.angmomz) < 10 then
    'debug.print "parm: " & parm & " momz: " & activeball.angmomz &" vely: "& activeball.vely
    activeball.angmomz = activeball.angmomz * 1.2
    activeball.vely = activeball.vely * 1.2
    'debug.print ">> newmomz: " & activeball.angmomz&" newvely: "& activeball.vely
  Elseif parm <= 2 and parm > 0.2 and activeball.vely < 0 Then
    'debug.print "* parm: " & parm & " momz: " & activeball.angmomz &" vely: "& activeball.vely
    activeball.angmomz = activeball.angmomz * -1.1
    activeball.vely = activeball.vely * 1.4
    'debug.print "**** >> newmomz: " & activeball.angmomz&" newvely: "& activeball.vely
  end if
end sub


' apophis Rubberizer
sub Rubberizer2(parm)
  if parm < 10 And parm > 2 And Abs(activeball.angmomz) < 10 then
    'debug.print "parm: " & parm & " momz: " & activeball.angmomz &" vely: "& activeball.vely
    activeball.angmomz = -activeball.angmomz * 2
    activeball.vely = activeball.vely * 1.2
    'debug.print ">> newmomz: " & activeball.angmomz&" newvely: "& activeball.vely
  Elseif parm <= 2 and parm > 0.2 and activeball.vely < 0 Then
    'debug.print "* parm: " & parm & " momz: " & activeball.angmomz &" vely: "& activeball.vely
    activeball.angmomz = -activeball.angmomz * 0.5
    activeball.vely = activeball.vely * (1.2 + rnd(1)/3 )
    'debug.print "**** >> newmomz: " & activeball.angmomz&" newvely: "& activeball.vely
  end if
end sub

'**********************************************************************************************************

 ' Drain and Kickers
'Sub Drain_Hit:bsTrough.addball me : playsound"drain" : End Sub

' Saucer
Sub sw43_Hit()
  bsSaucer.addball 0
  RandomSoundEjectHoleEnter()
End Sub

Sub SolTopKickout(Enabled)
  If Enabled then
    bsSaucer.SolOut 1
    RandomSoundEjectHoleSolenoid()
    RandomSoundEjectBallBump()
  End If
End Sub


' Drop Targets
Sub SolLeftDTUp(Enabled)
  If enabled Then
    'Debug.Print "SolLeftDTUp"
    RandomSoundDropTargetLeft()
    'dtL.SolDropUp 1
    'dtL.DropSol_On
    DTRaise 27
    DTRaise 28
    DTRaise 29

  End If
End Sub


Sub SolTopDTUp(Enabled)
  If enabled Then
    'Debug.Print "SolTopDTUp"
    RandomSoundDropTargetTop()
    'dtT.SolDropUp 1
    dtT.DropSol_On
  End If
End Sub


'Sub kCellarRight_Hit() : vpmTimer.PulseSw  19 : End Sub
'Sub kCellarLeft_Hit() : Controller.Switch(20) = 1 : End Sub
'Sub kCellarLeft_Unhit() : Controller.Switch(20) = 0 : End Sub

' Cellar Drops
Sub CellarRightTrigger_Hit()
  If activeball.VelZ < 0 Then
    'debug.print "CellarRightTrigger_Hit"
    SoundCellarRightEnter()
  End If
End Sub

Sub CellarLeftTrigger_Hit()
  If activeball.VelZ < 0 Then
    'debug.print "CellarLeftTrigger_Hit()"
    SoundCellarLeftEnter()
  End If
End Sub


' Knocker
Sub SolKnocker(enabled)
  If enabled Then
    KnockerSolenoid()
  End If
End Sub




'Wire triggers - Outlane triggers
Sub sw15_Hit() : Controller.Switch(15) = 1 : RandomSoundRollover() : RandomSoundOutlaneRollover() : End Sub
Sub sw15_Unhit() : Controller.Switch(15) = 0 : End Sub
Sub sw16_Hit() : Controller.Switch(16) = 1 : RandomSoundRollover() : RandomSoundOutlaneRollover() : End Sub
Sub sw16_Unhit() : Controller.Switch(16) = 0 : End Sub
Sub sw17_Hit() : Controller.Switch(17) = 1 : RandomSoundRollover() : RandomSoundOutlaneRollover() : End Sub
Sub sw17_Unhit() : Controller.Switch(17) = 0 : End Sub
Sub sw18_Hit() : Controller.Switch(18) = 1 : RandomSoundRollover() : RandomSoundOutlaneRollover() : End Sub
Sub sw18_Unhit() : Controller.Switch(18) = 0 : End Sub

'Outlane - Walls & Primitives - Right
Sub Wire_Rail_005_Hit() : RandomSoundOutlaneWalls() : End Sub
Sub Wire_Rail_009_Hit() : RandomSoundOutlaneWalls() : End Sub
Sub zCol_Band001_Hit() : RandomSoundOutlaneWalls() : End Sub

'Outlane - Walls & Primitives - Right
Sub Wall25_Hit() : RandomSoundOutlaneWalls() : End Sub
Sub Wire_Rail_003_Hit() : RandomSoundOutlaneWalls() : End Sub
Sub zCol_Band003_Hit() : RandomSoundOutlaneWalls() : End Sub




'Wire triggers - Other
Sub sw36_Hit() : Controller.Switch(36) = 1 : RandomSoundRollover() : End Sub
Sub sw36_Unhit() : Controller.Switch(36) = 0 : End Sub
Sub sw37_Hit() : Controller.Switch(37) = 1 : RandomSoundRollover() : End Sub
Sub sw37_Unhit() : Controller.Switch(37) = 0 : End Sub
Sub sw38_Hit() : Controller.Switch(38) = 1 : RandomSoundRollover() : End Sub
Sub sw38_Unhit() : Controller.Switch(38) = 0 : End Sub
Sub sw39_Hit() : Controller.Switch(39) = 1 : RandomSoundRollover() : End Sub
Sub sw39_Unhit() : Controller.Switch(39) = 0 : End Sub
Sub sw59_Hit() : Controller.Switch(59) = 1 : RandomSoundRollover() : End Sub
Sub sw59_Unhit() : Controller.Switch(59) = 0 : End Sub

Dim TargetDelay,TargetRotX
TargetDelay = 50
TargetRotX = 7

'Stand Up Targets
Sub sw21_Hit()
  vpmTimer.PulseSw 21
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 10 then
    RandomSoundTargetHitStrong()
    RandomSoundBallBouncePlayfieldSoft()
  End if
  If finalspeed <= 10 then
    RandomSoundTargetHitWeak()
  End If
  coll019.rotX=TargetRotX
  vpmTimer.AddTimer TargetDelay,"coll019.rotX=TargetrotX/2'"
  vpmTimer.AddTimer TargetDelay+TargetDelay,"coll019.rotX=0'"
End Sub

Sub sw25_Hit()
  vpmTimer.PulseSw 25
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 10 then
    RandomSoundTargetHitStrong()
    RandomSoundBallBouncePlayfieldSoft()
  End if
  If finalspeed <= 10 then
    RandomSoundTargetHitWeak()
  End If
  coll020.rotX=TargetRotX
  vpmTimer.AddTimer TargetDelay,"coll020.rotX=TargetrotX/2'"
  vpmTimer.AddTimer TargetDelay+TargetDelay,"coll020.rotX=0'"
End Sub

Sub sw30_Hit()
  vpmTimer.PulseSw 30
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 10 then
    RandomSoundTargetHitStrong()
    RandomSoundBallBouncePlayfieldSoft()
  End if
  If finalspeed <= 10 then
    RandomSoundTargetHitWeak()
  End If
  coll021.rotX=TargetRotX
  vpmTimer.AddTimer TargetDelay,"coll021.rotX=TargetrotX/2'"
  vpmTimer.AddTimer TargetDelay+TargetDelay,"coll021.rotX=0'"
End Sub

Sub sw47_Hit()
  vpmTimer.PulseSw 47
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 10 then
    RandomSoundTargetHitStrong()
    RandomSoundBallBouncePlayfieldSoft()
  End if
  If finalspeed <= 10 then
    RandomSoundTargetHitWeak()
  End If
  coll018.rotX=TargetRotX
  vpmTimer.AddTimer TargetDelay,"coll018.rotX=TargetrotX/2'"
  vpmTimer.AddTimer TargetDelay+TargetDelay,"coll018.rotX=0'"
End Sub

Sub sw48_Hit()
  vpmTimer.PulseSw 48
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 10 then
    RandomSoundTargetHitStrong()
    RandomSoundBallBouncePlayfieldSoft()
  End if
  If finalspeed <= 10 then
    RandomSoundTargetHitWeak()
  End If
  coll017.rotX=TargetRotX
  vpmTimer.AddTimer TargetDelay,"coll017.rotX=TargetrotX/2'"
  vpmTimer.AddTimer TargetDelay+TargetDelay,"coll017.rotX=0'"
End Sub


'Drop Targets
Sub Sw26_Dropped:dtT.Hit 1 :End Sub
'Sub Sw27_Dropped:dtL.Hit 1 : sw27.image = "target3dropped" : End Sub
'Sub Sw28_Dropped:dtL.Hit 2 : sw28.image = "target2dropped" : End Sub
'Sub Sw29_Dropped:dtL.Hit 3 : sw29.image = "target1dropped" : End Sub
'Sub Sw29_Raised: sw29.image = "target1" : End Sub
'Sub Sw28_Raised: sw28.image = "target2" : End Sub
'Sub Sw27_Raised: sw27.image = "target3" : End Sub

Sub sw26_Hit() '1-Bank Drop Target
  'Debug.Print "sw26_Hit()"
  SoundDropTargetsw26()
End Sub

'Sub sw27_Hit() '3-Bank Drop Target
' 'Debug.Print "sw27_Hit()"
' SoundDropTargetsw27()
'End Sub
'
'Sub sw28_Hit() '3-Bank Drop Target
' 'Debug.Print "sw28_Hit()"
' SoundDropTargetsw28()
'End Sub
'
'Sub sw29_Hit() '3-Bank Drop Target
' 'Debug.Print "sw29_Hit()"
' SoundDropTargetsw29()
'End Sub

Sub sw27_Hit() '3-Bank Drop Target
  'Debug.Print "sw27_Hit()"
  DTHit 27
  SoundDropTargetsw27()
End Sub

Sub sw28_Hit() '3-Bank Drop Target
  'Debug.Print "sw28_Hit()"
  DTHit 28
  SoundDropTargetsw28()
End Sub

Sub sw29_Hit() '3-Bank Drop Target
  'Debug.Print "sw29_Hit()"
  DTHit 29
  SoundDropTargetsw29()
End Sub

'******************************************************
'****  DROP TARGETS by Rothbauerw
'******************************************************
' This solution improves the physics for drop targets to create more realistic behavior. It allows the ball
' to move through the target enabling the ability to score more than one target with a well placed shot.
' It also handles full drop target animation, including deflection on hit and a slight lift when the drop
' targets raise, switch handling, bricking, and popping the ball up if it's over the drop target when it raises.
'
'Add a Timer named DTAnim to editor to handle drop & standup target animations, or run them off an always-on 10ms timer (GameTimer)
'DTAnim.interval = 10
'DTAnim.enabled = True

'Sub DTAnim_Timer
' DoDTAnim
' DoSTAnim
'End Sub

' For each drop target, we'll use two wall objects for physics calculations and one primitive for visuals and
' animation. We will not use target objects.  Place your drop target primitive the same as you would a VP drop target.
' The primitive should have it's pivot point centered on the x and y axis and at or just below the playfield
' level on the z axis. Orientation needs to be set using Rotz and bending deflection using Rotx. You'll find a hooded
' target mesh in this table's example. It uses the same texture map as the VP drop targets.

'******************************************************
'  DROP TARGETS INITIALIZATION
'******************************************************

'Define a variable for each drop target
Dim DT27, DT28, DT29

'Set array with drop target objects
'
'DropTargetvar = Array(primary, secondary, prim, swtich, animate)
'   primary:  primary target wall to determine drop
'   secondary:  wall used to simulate the ball striking a bent or offset target after the initial Hit
'   prim:    primitive target used for visuals and animation
'          IMPORTANT!!!
'          rotz must be used for orientation
'          rotx to bend the target back
'          transz to move it up and down
'          the pivot point should be in the center of the target on the x, y and at or below the playfield (0) on z
'   switch:  ROM switch number
'   animate:  Array slot for handling the animation instrucitons, set to 0
'          Values for animate: 1 - bend target (hit to primary), 2 - drop target (hit to secondary), 3 - brick target (high velocity hit to secondary), -1 - raise target
'   isDropped:  Boolean which determines whether a drop target is dropped. Set to false if they are initially raised, true if initially dropped.

DT27 = Array(sw27, sw27a, sw27p, 27, 0, False)
DT28 = Array(sw28, sw28a, sw28p, 28, 0, False)
DT29 = Array(sw29, sw29a, sw29p, 29, 0, False)

Dim DTArray
DTArray = Array(DT27, DT28, DT29)

Const UsingROM = True

'Configure the behavior of Drop Targets.
Const DTDropSpeed = 110 'in milliseconds
Const DTDropUpSpeed = 40 'in milliseconds
Const DTDropUnits = 46 'VP units primitive drops so top of at or below the playfield
Const DTDropUpUnits = 10 'VP units primitive raises above the up position on drops up
Const DTMaxBend = 8 'max degrees primitive rotates when hit
Const DTDropDelay = 20 'time in milliseconds before target drops (due to friction/impact of the ball)
Const DTRaiseDelay = 40 'time in milliseconds before target drops back to normal up position after the solenoid fires to raise the target
Const DTBrickVel = 30 'velocity at which the target will brick, set to '0' to disable brick
Const DTEnableBrick = 0 'Set to 0 to disable bricking, 1 to enable bricking
Const DTMass = 0.2 'Mass of the Drop Target (between 0 and 1), higher values provide more resistance

'******************************************************
'  DROP TARGETS FUNCTIONS
'******************************************************

Sub DTHit(switch)
  Dim i
  i = DTArrayID(switch)

' PlayTargetSound
  DTArray(i)(4) = DTCheckBrick(Activeball,DTArray(i)(2))
  If DTArray(i)(4) = 1 Or DTArray(i)(4) = 3 Or DTArray(i)(4) = 4 Then
    DTBallPhysics Activeball, DTArray(i)(2).rotz, DTMass
  End If
  DoDTAnim
End Sub

Sub DTRaise(switch)
  Dim i
  i = DTArrayID(switch)

  DTArray(i)(4) =  - 1
  DoDTAnim
End Sub

Sub DTDrop(switch)
  Dim i
  i = DTArrayID(switch)

  DTArray(i)(4) = 1
  DoDTAnim
End Sub

Function DTArrayID(switch)
  Dim i
  For i = 0 To UBound(DTArray)
    If DTArray(i)(3) = switch Then
      DTArrayID = i
      Exit Function
    End If
  Next
End Function

Sub DTBallPhysics(aBall, angle, mass)
  Dim rangle,bangle,calc1, calc2, calc3
  rangle = (angle - 90) * 3.1416 / 180
  bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))

  calc1 = cor.BallVel(aball.id) * Cos(bangle - rangle) * (aball.mass - mass) / (aball.mass + mass)
  calc2 = cor.BallVel(aball.id) * Sin(bangle - rangle) * Cos(rangle + 4 * Atn(1) / 2)
  calc3 = cor.BallVel(aball.id) * Sin(bangle - rangle) * Sin(rangle + 4 * Atn(1) / 2)

  aBall.velx = calc1 * Cos(rangle) + calc2
  aBall.vely = calc1 * Sin(rangle) + calc3
End Sub

'Check if target is hit on it's face or sides and whether a 'brick' occurred
Function DTCheckBrick(aBall, dtprim)
  Dim bangle, bangleafter, rangle, rangle2, Xintersect, Yintersect, cdist, perpvel, perpvelafter, paravel, paravelafter
  rangle = (dtprim.rotz - 90) * 3.1416 / 180
  rangle2 = dtprim.rotz * 3.1416 / 180
  bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))
  bangleafter = Atn2(aBall.vely,aball.velx)

  Xintersect = (aBall.y - dtprim.y - Tan(bangle) * aball.x + Tan(rangle2) * dtprim.x) / (Tan(rangle2) - Tan(bangle))
  Yintersect = Tan(rangle2) * Xintersect + (dtprim.y - Tan(rangle2) * dtprim.x)

  cdist = Distance(dtprim.x, dtprim.y, Xintersect, Yintersect)

  perpvel = cor.BallVel(aball.id) * Cos(bangle - rangle)
  paravel = cor.BallVel(aball.id) * Sin(bangle - rangle)

  perpvelafter = BallSpeed(aBall) * Cos(bangleafter - rangle)
  paravelafter = BallSpeed(aBall) * Sin(bangleafter - rangle)

  If perpvel > 0 And  perpvelafter <= 0 Then
    If DTEnableBrick = 1 And  perpvel > DTBrickVel And DTBrickVel <> 0 And cdist < 8 Then
      DTCheckBrick = 3
    Else
      DTCheckBrick = 1
    End If
  ElseIf perpvel > 0 And ((paravel > 0 And paravelafter > 0) Or (paravel < 0 And paravelafter < 0)) Then
    DTCheckBrick = 4
  Else
    DTCheckBrick = 0
  End If
End Function

Sub DoDTAnim()
  Dim i
  For i = 0 To UBound(DTArray)
    DTArray(i)(4) = DTAnimate(DTArray(i)(0),DTArray(i)(1),DTArray(i)(2),DTArray(i)(3),DTArray(i)(4))
  Next
End Sub

Function DTAnimate(primary, secondary, prim, switch, animate)
  Dim transz, switchid
  Dim animtime, rangle

  switchid = switch

  Dim ind
  ind = DTArrayID(switchid)

  rangle = prim.rotz * PI / 180

  DTAnimate = animate

  If animate = 0 Then
    primary.uservalue = 0
    DTAnimate = 0
    Exit Function
  ElseIf primary.uservalue = 0 Then
    primary.uservalue = gametime
  End If

  animtime = gametime - primary.uservalue

  If (animate = 1 Or animate = 4) And animtime < DTDropDelay Then
    primary.collidable = 0
    If animate = 1 Then secondary.collidable = 1 Else secondary.collidable = 0
    prim.rotx = DTMaxBend * Cos(rangle)
    prim.roty = DTMaxBend * Sin(rangle)
    DTAnimate = animate
    Exit Function
  ElseIf (animate = 1 Or animate = 4) And animtime > DTDropDelay Then
    primary.collidable = 0
    If animate = 1 Then secondary.collidable = 1 Else secondary.collidable = 0
    prim.rotx = DTMaxBend * Cos(rangle)
    prim.roty = DTMaxBend * Sin(rangle)
    animate = 2
'   SoundDropTargetDrop prim
  End If


  If animate = 2 Then
    transz = (animtime - DTDropDelay) / DTDropSpeed * DTDropUnits *  - 1
    If prim.transz >  - DTDropUnits  Then
      prim.transz = transz
    End If

    prim.rotx = DTMaxBend * Cos(rangle) / 2
    prim.roty = DTMaxBend * Sin(rangle) / 2

    If prim.transz <= - DTDropUnits Then
      prim.transz =  - DTDropUnits
      secondary.collidable = 0
      DTArray(ind)(5) = True 'Mark target as dropped
      if switchid = 27 then sw27p.image = "target1dropped"
      if switchid = 28 then sw28p.image = "target2dropped"
      if switchid = 29 then sw29p.image = "target3dropped"
      If UsingROM Then
        controller.Switch(Switchid) = 1
      Else
        DTAction switchid
      End If
      primary.uservalue = 0
      DTAnimate = 0
      Exit Function
    Else
      DTAnimate = 2
      Exit Function
    End If
  End If

  If animate = 3 And animtime < DTDropDelay Then
    primary.collidable = 0
    secondary.collidable = 1
    prim.rotx = DTMaxBend * Cos(rangle)
    prim.roty = DTMaxBend * Sin(rangle)
  ElseIf animate = 3 And animtime > DTDropDelay Then
    primary.collidable = 1
    secondary.collidable = 0
    prim.rotx = 0
    prim.roty = 0
    primary.uservalue = 0
    DTAnimate = 0
    Exit Function
  End If

  If animate =  - 1 Then
    transz = (1 - (animtime) / DTDropUpSpeed) * DTDropUnits *  - 1

    If prim.transz =  - DTDropUnits Then
      Dim b
      Dim BOT
      BOT = GetBalls

      For b = 0 To UBound(BOT)
        If InRotRect(BOT(b).x,BOT(b).y,prim.x, prim.y, prim.rotz, - 25, - 10,25, - 10,25,25, - 25,25) And BOT(b).z < prim.z + DTDropUnits + 25 Then
          BOT(b).velz = 20
        End If
      Next
    End If

    If prim.transz < 0 Then
      prim.transz = transz
    ElseIf transz > 0 Then
      prim.transz = transz
    End If

    If prim.transz > DTDropUpUnits Then
      DTAnimate =  - 2
      prim.transz = DTDropUpUnits
      prim.rotx = 0
      prim.roty = 0
      primary.uservalue = gametime
    End If
    primary.collidable = 0
    secondary.collidable = 1
    DTArray(ind)(5) = False 'Mark target as not dropped
    if switchid = 27 then sw27p.image = "target1"
    if switchid = 28 then sw28p.image = "target2"
    if switchid = 29 then sw29p.image = "target3"
    If UsingROM Then controller.Switch(Switchid) = 0
  End If

  If animate =  - 2 And animtime > DTRaiseDelay Then
    prim.transz = (animtime - DTRaiseDelay) / DTDropSpeed * DTDropUnits *  - 1 + DTDropUpUnits
    If prim.transz < 0 Then
      prim.transz = 0
      primary.uservalue = 0
      DTAnimate = 0

      primary.collidable = 1
      secondary.collidable = 0
    End If
  End If
End Function

Sub DTAction(switchid)
  Select Case switchid
    Case 1
    Addscore 1000
    ShadowDT(0).visible = False

    Case 2
    Addscore 1000
    ShadowDT(1).visible = False

    Case 3
    Addscore 1000
    ShadowDT(2).visible = False
  End Select
End Sub

'******************************************************
'  DROP TARGET
'  SUPPORTING FUNCTIONS
'******************************************************

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
'****  END DROP TARGETS
'******************************************************




'Gate Triggers
Sub sw33_Hit() : vpmTimer.PulseSw 33 : End Sub
Sub sw40_Hit() : vpmtimer.PulseSw 40 : End Sub


Sub sw33g_Hit() : SoundRampBrktGate1(SoundOn) : End Sub 'A-12506 - Ball Gate Wire
Sub sw40g_Hit() : SoundRampBrktGate2(SoundOn) : End Sub 'A-13376 - Ball Gate Wire
Sub Gate1_Hit() : SoundBallGate1() : End Sub 'A-8244-L - Ball Gate Flap Assembly (Part #17)
Sub Gate3_Hit() : SoundBallGate3() : End Sub 'A-8244-R - Ball Gate Flap Assembly (Part #32)
Sub BallReleaseGate_Hit() : SoundBallReleaseGate() : End Sub




'Spinner
Sub sw41_Spin() : vpmTimer.PulseSw 41 : SoundSpinner() : End Sub

'Ramp triggers
Sub sw34_Hit() : vpmTimer.PulseSw 34 : RandomSoundRollover() : End Sub
Sub sw35_Hit() : vpmTimer.PulseSw 35 : RandomSoundRollover() : End Sub
Sub sw44_Hit() : vpmTimer.PulseSw 44 : RandomSoundRollover() : End Sub
Sub sw45_Hit() : vpmTimer.PulseSw 45 : RandomSoundRollover() : End Sub

'Bumpers
Sub Bumper1_Hit() : vpmTimer.PulseSw 49 : RandomSoundBumperTopLeft() : End Sub     'Left Top Jet Bumper
Sub Bumper2_Hit() : vpmtimer.PulseSw 50 : RandomSoundBumperTopRight() : End Sub    'Right Top Jet Bumper
Sub Bumper3_Hit() : vpmTimer.PulseSw 51 : RandomSoundBumperTopLower() : End Sub    'Lower Top Jet Bumper
Sub Bumper4_Hit() : vpmTimer.PulseSw 52 : RandomSoundBumperBottomLeft() : End Sub  'Left Bottom Jet Bumper
Sub Bumper5_Hit() : vpmTimer.PulseSw 53 : RandomSoundBumperBottomRight() : End Sub 'Right Bottom Jet Bumper
Sub Bumper6_Hit() : vpmTimer.PulseSw 54 : RandomSoundBumperBottomTop() : End Sub   'Top Bottom Jet Bumper










'Scoring Rubbers
Sub sw60_Hit()
  vpmTimer.PulseSw 60
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 10 then
    RandomSoundRubberStrong()
  End if
  If finalspeed <= 10 then
    RandomSoundRubberWeak()
  End If
End Sub

Sub sw61_Hit()
  vpmTimer.PulseSw 61
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 10 then
    RandomSoundRubberStrong()
  End if
  If finalspeed <= 10 then
    RandomSoundRubberWeak()
  End If
End Sub

'Lock Ball
'Sub Kicker1_Hit() : Kicker2.enabled = True : End Sub
'Sub Kicker2_Hit() : Kicker3.enabled = True : End Sub

Sub sw22_Hit() :
  Controller.Switch(22) = 1
  RandomSoundRollover()
  'debug.print "hit lock 1"
  End Sub

Sub sw22_Unhit()
  Controller.Switch(22) = 0
  'debug.print "unhit lock 1"
  End Sub

Sub sw23_Hit()
  Controller.Switch(23) = 1
  RandomSoundRollover()
  'debug.print "hit lock 2"
  End Sub

Sub sw23_Unhit()
  Controller.Switch(23) = 0
  'debug.print "unhit lock 2"
  End Sub

Sub sw24_Hit()
  Controller.Switch(24) = 1
  RandomSoundRollover()
  'debug.print"hit Lock 3"
End Sub

Sub sw24_Unhit()
  Controller.Switch(24) = 0
  'debug.print "unhit lock 3"
  End Sub

Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5,BallShadow6,BallShadow7,BallShadow8,BallShadow9,BallShadow10,BallShadow11)

Sub GameTimer_timer()
  RollingSoundUpdate
  cor.Update
  DoDTAnim
End Sub

'Generic Ramp Sounds
'Sub Trigger1_Hit() : playsound"Ball Drop" : End Sub
'Sub Trigger2_Hit() : playsound"Ball Drop" : End Sub
'Sub Trigger3_Hit() : playsound"Ball Drop" : End Sub
'Sub Trigger4_Hit() : playsound"Wire Ramp" : End Sub

'**********************************************************************************************************
'Digital Display
'**********************************************************************************************************
 Dim Digits(32)
 Digits(0)=Array(a00, a05, a0c, a0d, a08, a01, a06, a0f, a02, a03, a04, a07, a0b, a0a, a09, a0e)
 Digits(1)=Array(a10, a15, a1c, a1d, a18, a11, a16, a1f, a12, a13, a14, a17, a1b, a1a, a19, a1e)
 Digits(2)=Array(a20, a25, a2c, a2d, a28, a21, a26, a2f, a22, a23, a24, a27, a2b, a2a, a29, a2e)
 Digits(3)=Array(a30, a35, a3c, a3d, a38, a31, a36, a3f, a32, a33, a34, a37, a3b, a3a, a39, a3e)
 Digits(4)=Array(a40, a45, a4c, a4d, a48, a41, a46, a4f, a42, a43, a44, a47, a4b, a4a, a49, a4e)
 Digits(5)=Array(a50, a55, a5c, a5d, a58, a51, a56, a5f, a52, a53, a54, a57, a5b, a5a, a59, a5e)
 Digits(6)=Array(a60, a65, a6c, a6d, a68, a61, a66, a6f, a62, a63, a64, a67, a6b, a6a, a69, a6e)
 Digits(7)=Array(a70, a75, a7c, a7d, a78, a71, a76, a7f, a72, a73, a74, a77, a7b, a7a, a79, a7e)
 Digits(8)=Array(a80, a85, a8c, a8d, a88, a81, a86, a8f, a82, a83, a84, a87, a8b, a8a, a89, a8e)
 Digits(9)=Array(a90, a95, a9c, a9d, a98, a91, a96, a9f, a92, a93, a94, a97, a9b, a9a, a99, a9e)
 Digits(10)=Array(aa0, aa5, aac, aad, aa8, aa1, aa6, aaf, aa2, aa3, aa4, aa7, aab, aaa, aa9, aae)
 Digits(11)=Array(ab0, ab5, abc, abd, ab8, ab1, ab6, abf, ab2, ab3, ab4, ab7, abb, aba, ab9, abe)
 Digits(12)=Array(ac0, ac5, acc, acd, ac8, ac1, ac6, acf, ac2, ac3, ac4, ac7, acb, aca, ac9, ace)
 Digits(13)=Array(ad0, ad5, adc, add, ad8, ad1, ad6, adf, ad2, ad3, ad4, ad7, adb, ada, ad9, ade)
 Digits(14)=Array(ae0, ae5, aec, aed, ae8, ae1, ae6, aef, ae2, ae3, ae4, ae7, aeb, aea, ae9, aee)
 Digits(15)=Array(af0, af5, afc, afd, af8, af1, af6, aff, af2, af3, af4, af7, afb, afa, af9, afe)

 Digits(16)=Array(b00, b05, b0c, b0d, b08, b01, b06, b0f, b02, b03, b04, b07, b0b, b0a, b09, b0e)
 Digits(17)=Array(b10, b15, b1c, b1d, b18, b11, b16, b1f, b12, b13, b14, b17, b1b, b1a, b19, b1e)
 Digits(18)=Array(b20, b25, b2c, b2d, b28, b21, b26, b2f, b22, b23, b24, b27, b2b, b2a, b29, b2e)
 Digits(19)=Array(b30, b35, b3c, b3d, b38, b31, b36, b3f, b32, b33, b34, b37, b3b, b3a, b39, b3e)
 Digits(20)=Array(b40, b45, b4c, b4d, b48, b41, b46, b4f, b42, b43, b44, b47, b4b, b4a, b49, b4e)
 Digits(21)=Array(b50, b55, b5c, b5d, b58, b51, b56, b5f, b52, b53, b54, b57, b5b, b5a, b59, b5e)
 Digits(22)=Array(b60, b65, b6c, b6d, b68, b61, b66, b6f, b62, b63, b64, b67, b6b, b6a, b69, b6e)
 Digits(23)=Array(b70, b75, b7c, b7d, b78, b71, b76, b7f, b72, b73, b74, b77, b7b, b7a, b79, b7e)
 Digits(24)=Array(b80, b85, b8c, b8d, b88, b81, b86, b8f, b82, b83, b84, b87, b8b, b8a, b89, b8e)
 Digits(25)=Array(b90, b95, b9c, b9d, b98, b91, b96, b9f, b92, b93, b94, b97, b9b, b9a, b99, b9e)
 Digits(26)=Array(ba0, ba5, bac, bad, ba8, ba1, ba6, baf, ba2, ba3, ba4, ba7, bab, baa, ba9, bae)
 Digits(27)=Array(bb0, bb5, bbc, bbd, bb8, bb1, bb6, bbf, bb2, bb3, bb4, bb7, bbb, bba, bb9, bbe)
 Digits(28)=Array(bc0, bc5, bcc, bcd, bc8, bc1, bc6, bcf, bc2, bc3, bc4, bc7, bcb, bca, bc9, bce)
 Digits(29)=Array(bd0, bd5, bdc, bdd, bd8, bd1, bd6, bdf, bd2, bd3, bd4, bd7, bdb, bda, bd9, bde)
 Digits(30)=Array(be0, be5, bec, bed, be8, be1, be6, bef, be2, be3, be4, be7, beb, bea, be9, bee)
 Digits(31)=Array(bf0, bf5, bfc, bfd, bf8, bf1, bf6, bff, bf2, bf3, bf4, bf7, bfb, bfa, bf9, bfe)

 Sub DisplayTimer_Timer
  If VRRoom = 0 Then
    Dim ChgLED, ii, jj, num, chg, stat, obj, b, x
    ChgLED=Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
      If DesktopMode = True Then
        For ii=0 To UBound(chgLED)
          num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)
          if (num < 32) then
            For Each obj In Digits(num)
               If chg And 1 Then obj.intensity = 30 * (stat and 1) + 0.4 'obj.State=stat And 1
               chg=chg\2 : stat=stat\2
              Next
          Else
          end if
        Next
      end if
    End IF
    End If
 End Sub



'Hide Reels if DT false
  If DesktopMode = True Then
  dim xxxxxx
    For each xxxxxx in BG:xxxxxx.visible = 1: xxxxxx.state = 1 : xxxxxx.intensity = 0.4: Next
  Else
    For each xxxxxx in BG:xxxxxx.visible = 0: Next
  End If

'**********************************************************************************************************
'**********************************************************************************************************

'****************************************************************************************************
'****************************************************************************************************
'                 Start of VPX Call Backs
'****************************************************************************************************
'****************************************************************************************************

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
  RandomSoundSlingshotRight()
  vpmTimer.PulseSw 56
    RStep = 0
    RightSlingShot.TimerEnabled = 1
  rightsling.PlayAnim 0,0.2
End Sub


Sub RightSlingShot_Timer
    Select Case RStep
     Case 4: RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  RandomSoundSlingshotLeft()
  vpmTimer.PulseSw 55
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
  leftsling.PlayAnim 0,0.2
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
    Case 4: LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub




'************************************
' What you need to add to your table
'************************************

' a timer called RollingTimer. With a fast interval, like 10
' one collision sound, in this script is called fx_collide
' as many sound files as max number of balls, with names ending with 0, 1, 2, 3, etc
' for ex. as used in this script: fx_ballrolling0, fx_ballrolling1, fx_ballrolling2, fx_ballrolling3, etc


'******************************************
' Explanation of the rolling sound routine
'******************************************

' sounds are played based on the ball speed and position

' the routine checks first for deleted balls and stops the rolling sound.

' The For loop goes through all the balls on the table and checks for the ball speed and
' if the ball is on the table (height lower than 30) then then it plays the sound
' otherwise the sound is stopped, like when the ball has stopped or is on a ramp or flying.

' The sound is played using the VOL, PAN and PITCH functions, so the volume and pitch of the sound
' will change according to the ball speed, and the PAN function will change the stereo position according
' to the position of the ball on the table.


'**************************************
' Explanation of the collision routine
'**************************************

' The collision is built in VP.
' You only need to add a Sub OnBallBallCollision(ball1, ball2, velocity) and when two balls collide they
' will call this routine. What you add in the sub is up to you. As an example is a simple Playsound with volume and paning
' depending of the speed of the collision.


Sub Pins_Hit (idx)
  'PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
End Sub

Sub Metals_Thin_Hit (idx)
  RandomSoundWall()
End Sub

Sub Metals_Medium_Hit (idx)
  RandomSoundWall()
End Sub

Sub Metals2_Hit (idx)
  RandomSoundMetal()
End Sub

Sub Gates_Hit (idx)
  'PlaySound "gate4", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub


'******************************************************
'       FLIPPER AND RUBBER CORRECTION
'******************************************************

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

        addpt "Velocity", 0, 0,         1
        addpt "Velocity", 1, 0.16, 1.06
        addpt "Velocity", 2, 0.41,         1.05
        addpt "Velocity", 3, 0.53,         1'0.982
        addpt "Velocity", 4, 0.702, 0.968
        addpt "Velocity", 5, 0.95,  0.968
        addpt "Velocity", 6, 1.03,         0.945

        AddPt "Polarity", 0, 0, 0
        AddPt "Polarity", 1, 0.05, -5
        AddPt "Polarity", 2, 0.4, -5
        AddPt "Polarity", 3, 0.6, -4.5
        AddPt "Polarity", 4, 0.65, -4.0
        AddPt "Polarity", 5, 0.7, -3.5
        AddPt "Polarity", 6, 0.75, -3.0
        AddPt "Polarity", 7, 0.8, -2.5
        AddPt "Polarity", 8, 0.85, -2.0
        AddPt "Polarity", 9, 0.9,-1.5
        AddPt "Polarity", 10, 0.95, -1.0
        AddPt "Polarity", 11, 1, -0.5
        AddPt "Polarity", 12, 1.1, 0
        AddPt "Polarity", 13, 1.3, 0

  LF.Object = LeftFlipper
  LF.EndPoint = EndPointLp  'you can use just a coordinate, or an object with a .x property. Using a couple of simple primitive objects
  RF.Object = RightFlipper
  RF.EndPoint = EndPointRp
End Sub

Sub AddPt(aStr, idx, aX, aY)  'debugger wrapper for adjusting flipper script in-game
  dim a : a = Array(LF, RF)
  dim x : for each x in a
    x.addpoint aStr, idx, aX, aY
  Next
End Sub

Sub TriggerLF_Hit() : LF.Addball activeball : End Sub
Sub TriggerLF_UnHit() : LF.PolarityCorrect activeball : End Sub
Sub TriggerRF_Hit() : RF.Addball activeball : End Sub
Sub TriggerRF_UnHit() : RF.PolarityCorrect activeball : End Sub

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
         :                         VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)

                                if partialflipcoef < 1 then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)

                                if Enabled then aBall.Velx = aBall.Velx*VelCoef
                                if Enabled then aBall.Vely = aBall.Vely*VelCoef
                        End If

                        'Polarity Correction (optional now)
                        if not IsEmpty(PolarityIn(0) ) then
                                If StartPoint > EndPoint then LR = -1        'Reverse polarity if left flipper
                                dim AddX : AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR

                                if Enabled then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef)
                                'playsound "fx_knocker"
                        End If
                End If
                RemoveBall aBall
        End Sub
End Class

'================================
'Helper Functions


Sub ShuffleArray(ByRef aArray, byVal offset) 'shuffle 1d array
  dim x, aCount : aCount = 0
  redim a(uBound(aArray) )
  for x = 0 to uBound(aArray) 'Shuffle objects in a temp array
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
  redim aArray(aCount-1+offset) 'Resize original array
  for x = 0 to aCount-1   'set objects back into original array
    if IsObject(a(x)) then
      Set aArray(x) = a(x)
    Else
      aArray(x) = a(x)
    End If
  Next
End Sub

Sub ShuffleArrays(aArray1, aArray2, offset)
  ShuffleArray aArray1, offset
  ShuffleArray aArray2, offset
End Sub


Function BallSpeed(ball) 'Calculates the ball speed
    BallSpeed = SQR(ball.VelX^2 + ball.VelY^2 + ball.VelZ^2)
End Function

Function PSlope(Input, X1, Y1, X2, Y2)  'Set up line via two points, no clamping. Input X, output Y
  dim x, y, b, m : x = input : m = (Y2 - Y1) / (X2 - X1) : b = Y2 - m*X2
  Y = M*x+b
  PSlope = Y
End Function

Function NullFunctionZ(aEnabled):End Function '1 argument null function placeholder  TODO move me or replac eme

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


'****************************************************************************
'PHYSICS DAMPENERS

'These are data mined bounce curves,
'dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
'Requires tracking ballspeed to calculate COR


Sub dPosts_Hit(idx)
  RubbersD.dampen Activeball
   dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 10 then
    RandomSoundRubberStrong()
  End if
  If finalspeed <= 10 then
    RandomSoundRubberWeak()
  End If
End Sub

Sub dSleeves_Hit(idx)
  SleevesD.Dampen Activeball
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 10 then
    RandomSoundRubberStrong()
  End if
  If finalspeed <= 10 then
    RandomSoundRubberWeak()
  End If
End Sub

dim RubbersD : Set RubbersD = new Dampener  'frubber
RubbersD.name = "Rubbers"
RubbersD.debugOn = False  'shows info in textbox "TBPout"
RubbersD.Print = False  'debug, reports in debugger (in vel, out cor)
'cor bounce curve (linear)
'for best results, try to match in-game velocity as closely as possible to the desired curve
'RubbersD.addpoint 0, 0, 0.935  'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 0, 0, 0.935 '0.96 'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 1, 3.77, 0.935 '0.96
RubbersD.addpoint 2, 5.76, 0.942 '0.967 'dont take this as gospel. if you can data mine rubber elasticitiy, please help!
RubbersD.addpoint 3, 15.84, 0.874
RubbersD.addpoint 4, 56, 0.64 'there's clamping so interpolate up to 56 at least

dim SleevesD : Set SleevesD = new Dampener  'this is just rubber but cut down to 85%...
SleevesD.name = "Sleeves"
SleevesD.debugOn = False  'shows info in textbox "TBPout"
SleevesD.Print = False  'debug, reports in debugger (in vel, out cor)
SleevesD.CopyCoef RubbersD, 0.85

Class Dampener
  Public Print, debugOn 'tbpOut.text
  public name, Threshold  'Minimum threshold. Useful for Flippers, which don't have a hit threshold.
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
    RealCOR = BallSpeed(aBall) / cor.ballvel(aBall.id)
    coef = desiredcor / realcor
    if debugOn then str = name & " in vel:" & round(cor.ballvel(aBall.id),2 ) & vbnewline & "desired cor: " & round(desiredcor,4) & vbnewline & _
    "actual cor: " & round(realCOR,4) & vbnewline & "ballspeed coef: " & round(coef, 3) & vbnewline
    if Print then debug.print Round(cor.ballvel(aBall.id),2) & ", " & round(desiredcor,3)

' Thalamus - patched :     aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
    aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
    'playsound "fx_knocker"
    if debugOn then TBPout.text = str
  End Sub

  Public Sub CopyCoef(aObj, aCoef) 'alternative addpoints, copy with coef
    dim x : for x = 0 to uBound(aObj.ModIn)
      addpoint x, aObj.ModIn(x), aObj.ModOut(x)*aCoef
    Next
  End Sub


  Public Sub Report()   'debug, reports all coords in tbPL.text
    if not debugOn then exit sub
    dim a1, a2 : a1 = ModIn : a2 = ModOut
    dim str, x : for x = 0 to uBound(a1) : str = str & x & ": " & round(a1(x),4) & ", " & round(a2(x),4) & vbnewline : next
    TBPout.text = str
  End Sub


End Class

'Tracks ball velocity for judging bounce calculations & angle
'apologies to JimmyFingers is this is what his script does. I know his tracks ball velocity too but idk how it works in particular
'dim cor : set cor = New CoRTracker
'cor.debugOn = False
''cor.update() - put this on a low interval timer
'Class CoRTracker
' public DebugOn 'tbpIn.text
' public ballvel
'
' Private Sub Class_Initialize : redim ballvel(0) : End Sub
' 'TODO this would be better if it didn't do the sorting every ms, but instead every time it's pulled for COR stuff
' Public Sub Update() 'tracks in-ball-velocity
'   dim str, b, AllBalls, highestID : allBalls = getballs
'   'if uBound(allballs) < 0 then if DebugOn then str = "no balls" : TBPin.text = str : exit Sub else exit sub end if: end if
'   for each b in allballs
'     if b.id >= HighestID then highestID = b.id
'   Next
'
'   if uBound(ballvel) < highestID then redim ballvel(highestID)  'set bounds
'
'   for each b in allballs
'     ballvel(b.id) = BallSpeed(b)
''      if DebugOn then
''        dim s, bs 'debug spacer, ballspeed
''        bs = round(BallSpeed(b),1)
''        if bs < 10 then s = " " else s = "" end if
''        str = str & b.id & ": " & s & bs & vbnewline
''        'str = str & b.id & ": " & s & bs & "z:" & b.z & vbnewline
''      end if
'   Next
'   'if DebugOn then str = "ubound ballvels: " & ubound(ballvel) & vbnewline & str : if TBPin.text <> str then TBPin.text = str : end if
' End Sub
'End Class

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


Function LinearEnvelope(xInput, xKeyFrame, yLvl)
  dim y 'Y output
  dim L 'Line
  dim ii : for ii = 1 to uBound(xKeyFrame)  'find active line
    if xInput <= xKeyFrame(ii) then L = ii : exit for : end if
  Next
  if xInput > xKeyFrame(uBound(xKeyFrame) ) then L = uBound(xKeyFrame)  'catch line overrun
  Y = pSlope(xInput, xKeyFrame(L-1), yLvl(L-1), xKeyFrame(L), yLvl(L) )

  'Clamp if on the boundry lines
  'if L=1 and Y < yLvl(LBound(yLvl) ) then Y = yLvl(lBound(yLvl) )
  'if L=uBound(xKeyFrame) and Y > yLvl(uBound(yLvl) ) then Y = yLvl(uBound(yLvl) )
  'clamp 2.0
  if xInput <= xKeyFrame(lBound(xKeyFrame) ) then Y = yLvl(lBound(xKeyFrame) )  'Clamp lower
  if xInput >= xKeyFrame(uBound(xKeyFrame) ) then Y = yLvl(uBound(xKeyFrame) )  'Clamp upper

  LinearEnvelope = Y
End Function


'////////////////////////////////////////////////////////////////////////////////
'////          Mechanical Sounds, by Fleep                                   ////
'////                     Last Updated: January, 2022                        ////
'////////////////////////////////////////////////////////////////////////////////
'
'/////////////////////////////////  CARTRIDGES  /////////////////////////////////
'
'//  Specify which mechanical sound cartridge to use for each group of elements.
'//  Mechanical sounds naming convention: <CARTRIDGE>_<Soundset_Name>
'//
'//  Cartridge name is composed using the following convention:
'//  <TABLE MANUFACTURER ABBREVIATION>_<TABLE NAME ABBREVIATION>_<SOUNDSET REVISION NUMBER>
'//
'//  General Mechanical Sounds Cartridges:
Const Cartridge_Bumpers         = "WS_WHD_REV01" 'Williams Whirlwind Cartridge REV01
Const Cartridge_Slingshots        = "WS_WHD_REV01"
Const Cartridge_Flippers        = "WS_WHD_REV01"
Const Cartridge_Kickers         = "WS_WHD_REV01"
Const Cartridge_Diverters       = "WS_DNR_REV01" 'Williams Diner Cartridge REV01
Const Cartridge_Knocker         = "WS_WHD_REV02" 'Williams Whirlwind Cartridge REV02
Const Cartridge_Relays          = "WS_WHD_REV01"
Const Cartridge_Trough          = "WS_WHD_REV01"
Const Cartridge_Rollovers       = "WS_WHD_REV01"
Const Cartridge_Targets         = "WS_WHD_REV01"
Const Cartridge_Gates         = "WS_WHD_REV01"
Const Cartridge_Spinner         = "SY_TNA_REV01" 'Spooky Total Nuclear Annihilation Cartridge REV01
Const Cartridge_Rubber_Hits       = "WS_WHD_REV01"
Const Cartridge_Metal_Hits        = "WS_WHD_REV01"
Const Cartridge_Plastic_Hits      = "WS_WHD_REV01"
Const Cartridge_Wood_Hits       = "WS_WHD_REV01"
Const Cartridge_Cabinet_Sounds      = "WS_WHD_REV01"
Const Cartridge_Drain         = "WS_WHD_REV01"
Const Cartridge_Apron         = "WS_WHD_REV01"
Const Cartridge_Ball_Roll       = "BY_TOM_REV01" 'Bally Theatre of Magic Cartridge REV01
Const Cartridge_BallBallCollision   = "BY_WDT_REV01" 'Bally WHO Dunnit Cartridge REV01
Const Cartridge_Ball_Drop_Bump      = "WS_WHD_REV01"
Const Cartridge_Plastic_Ramps     = "WS_WHD_REV01"
Const Cartridge_Metal_Ramps       = "WS_WHD_REV01"
Const Cartridge_Ball_Guides       = "WS_WHD_REV01"
Const Cartridge_Table_Specifics     = "WS_WHD_REV01"


'////////////////////////////  SOUND SOURCE CREDITS  ////////////////////////////
'//  Special thanks go to the following contributors who have provided audio
'//  footage recordings:
'//
'//  Williams Whirlwind - Blackmoor, wrd1972
'//  Williams Diner - Nick Rusis
'//  Spooky Total Nuclear Annihilation - WildDogArcade, Ed and Gary
'//  Bally Theatre of Magic - CalleV, nickbuol
'//  Bally WHO Dunnit - Amazaley1


'///////////////////////////////  USER PARAMETERS  //////////////////////////////
'
'//  Sounds Parameter with suffix "SoundLevel" can have any value in range [0..1]
'//  Sounds Parameter with suffix "SoundMultiplier" can have any value


'///////////////////////////  SOLENOIDS (COILS) CONFIG  /////////////////////////

'//  FLIPPER COILS:
'//  Flippers in this table: Lower Left Flipper, Lower Right Flipper, Upper Right Fliiper
Dim FlipperUpSoundLevel, FlipperDownSoundLevel, FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel
Dim FlipperLeftLowerHitParm, FlipperRightUpperHitParm, FlipperRightLowerHitParm

'//  Flipper Up Attacks initialize during playsound subs
Dim FlipperUpAttackLeftSoundLevel, FlipperUpAttackRightSoundLevel


FlipperUpSoundLevel = 1
FlipperDownSoundLevel = 0.65
FlipperUpAttackMinimumSoundLevel = 0.010
FlipperUpAttackMaximumSoundLevel = 0.435


'//  Flipper Hit Param initialize with FlipperUpSoundLevel
'//  and dynamically modified calculated by ball flipper collision
FlipperLeftLowerHitParm = FlipperUpSoundLevel
FlipperRightLowerHitParm = FlipperUpSoundLevel
FlipperRightUpperHitParm = FlipperUpSoundLevel


'//  CONTROLLED / SWITCHED COILS:
'//  Solenoid 01A = Outhole Kicker
'//  Solenoid 02A = Shooter Feeder
'//  Solenoid 03A = Right Ramp Lifter
'//  Solenoid 04A = Left Locking Kickback
'//  Solenoid 05A = Top Eject
'//  Solenoid 06A = Knocker
'//  Solenoid 07A,08A = 3-Bank Drop Target Reset, 1-Bank Drop Target Reset
'//  Solenoid 13 = Diverter
'//  Solenoid 14 = Under Playfield Kickbig
'//  Solenoid 09,10,15,17,19,21 = 3 Lower Bumpers, 3 Upper Bumpers
'//  Solenoid 18,20 = Left Kicker (Slingshot), Right Kicker (Slingshot)
'//  Solenoid 22 = Right Ramp Down

Dim Solenoid_OutholeKicker_SoundLevel, Solenoid_ShooterFeeder_SoundLevel
Dim Solenoid_RightRampLifter_SoundLevel, Solenoid_LeftLockingKickback_SoundLevel
Dim Solenoid_TopEject_SoundLevel, Solenoid_Knocker_SoundLevel, Solenoid_DropTargetReset_SoundLevel
Dim Solenoid_Diverter_Enabled_SoundLevel, Solenoid_Diverter_Hold_SoundLevel, Solenoid_Diverter_Disabled_SoundLevel
Dim Solenoid_UnderPlayfieldKickbig_SoundLevel, Solenoid_Bumper_SoundMultiplier
Dim Solenoid_Slingshot_SoundLevel, Solenoid_RightRampDown_SoundLevel

Solenoid_OutholeKicker_SoundLevel = 1
Solenoid_ShooterFeeder_SoundLevel = 1
Solenoid_RightRampLifter_SoundLevel = 0.3
Solenoid_RightRampDown_SoundLevel = 0.3
Solenoid_LeftLockingKickback_SoundLevel = 1
Solenoid_TopEject_SoundLevel = 1
Solenoid_Knocker_SoundLevel = 1
Solenoid_DropTargetReset_SoundLevel = 1
Solenoid_Diverter_Enabled_SoundLevel = 1
Solenoid_Diverter_Hold_SoundLevel = 0.7
Solenoid_Diverter_Disabled_SoundLevel = 0.4
Solenoid_UnderPlayfieldKickbig_SoundLevel = 1
Solenoid_Bumper_SoundMultiplier = 0.004 '8
Solenoid_Slingshot_SoundLevel = 1


'//  RELAYS:
'//  Solenoid 16 = Lower Playfield Relay GI Relay (P/N 5580-12145-00) / Backbox GI Relay (P/N 5580-09555-01)
'//  Solenoid 11 = Upper Playfield Relay GI Relay (P/N 5580-12145-00)
'//  Solenoid 12 = Solenoid A/C Select Relay (5580-09555-01)
'//  Fake Solenoid = Flahser Relay

Dim RelayLowerGISoundLevel, RelayUpperGISoundLevel, RelaySolenoidACSelectSoundLevel, RelayFlasherSoundLevel
RelayLowerGISoundLevel = 0.45
RelayUpperGISoundLevel = 0.45
RelaySolenoidACSelectSoundLevel = 0.3
RelayFlasherSoundLevel = 0.015


'//  EXTRA SOLENOIDS:
'//  Solenoid 24 = Blower Motor (Ontop Backbox)
'//  Solenoid 27 = Spin Wheels Motor
Dim Solenoid_BlowerMotor_SoundLevel, Solenoid_SpinWheelsMotor_SoundLevel
Solenoid_BlowerMotor_SoundLevel = 0.2
Solenoid_SpinWheelsMotor_SoundLevel = 0.2


'////////////////////////////  SWITCHES SOUND CONFIG  ///////////////////////////
Dim Switch_Gate_SoundLevel, SpinnerSoundLevel, RolloverSoundLevel, OutLaneRolloverSoundLevel, TargetSoundFactor

Switch_Gate_SoundLevel = 1
SpinnerSoundLevel = 0.2
RolloverSoundLevel = 0.55
OutLaneRolloverSoundLevel = 0.8
TargetSoundFactor = 0.8


'////////////////////  BALL HITS, BUMPS, DROPS SOUND CONFIG  ////////////////////
Dim BallWithBallCollisionSoundFactor, BallBouncePlayfieldSoftFactor, BallBouncePlayfieldHardFactor
Dim WallImpactSoundFactor, MetalImpactSoundFactor, WireformAntiRebountRailSoundFactor
Dim RubberStrongSoundFactor, RubberWeakSoundFactor, RubberFlipperSoundFactor
Dim BottomArchBallGuideSoundFactor, FlipperBallGuideSoundFactor, OutlaneWallsSoundFactor
Dim EjectBallBumpSoundLevel, EjectHoleEnterSoundLevel
Dim RightRampMetalWireDropToPlayfieldSoundLevel, LeftPlasticRampDropToLockSoundLevel, LeftPlasticRampDropToPlayfieldSoundLevel
Dim CellarLeftEnterSoundLevel, CellarRightEnterSoundLevel, CellerKickouBallDroptSoundLevel


BallWithBallCollisionSoundFactor = 3.2
BallBouncePlayfieldSoftFactor = 0.015
BallBouncePlayfieldHardFactor = 0.0075
WallImpactSoundFactor = 0.075
MetalImpactSoundFactor = 0.075
RubberStrongSoundFactor = 0.045
RubberWeakSoundFactor = 0.055
RubberFlipperSoundFactor = 0.65
BottomArchBallGuideSoundFactor = 0.2
FlipperBallGuideSoundFactor = 0.015
WireformAntiRebountRailSoundFactor = 0.04
OutlaneWallsSoundFactor = 1
EjectBallBumpSoundLevel = 1
RightRampMetalWireDropToPlayfieldSoundLevel = 1
LeftPlasticRampDropToLockSoundLevel = 1
LeftPlasticRampDropToPlayfieldSoundLevel = 1
EjectHoleEnterSoundLevel = 0.75
CellerKickouBallDroptSoundLevel = 1
CellarLeftEnterSoundLevel = 0.85
CellarRightEnterSoundLevel = 0.85


'///////////////////////  OTHER PLAYFIELD ELEMENTS CONFIG  //////////////////////
Dim RollingSoundFactor, RollingOnDiscSoundFactor, BallReleaseShooterLaneSoundLevel
Dim LeftPlasticRampEnteranceSoundLevel, RightPlasticRampEnteranceSoundLevel
Dim LeftPlasticRampRollSoundFactor, RightPlasticRampRollSoundFactor
Dim LeftMetalWireRampRollSoundFactor, RightPlasticRampHitsSoundLevel, LeftPlasticRampHitsSoundLevel
Dim SpinningDiscRolloverSoundFactor, SpinningDiscRolloverBumpSoundLevel
Dim LaneSoundFactor, LaneEnterSoundFactor, InnerLaneSoundFactor
Dim LaneLoudImpactMinimumSoundLevel, LaneLoudImpactMaximumSoundLevel

RollingSoundFactor = 50
RollingOnDiscSoundFactor = 1.5
BallReleaseShooterLaneSoundLevel = 1
LeftPlasticRampEnteranceSoundLevel = 0.1
RightPlasticRampEnteranceSoundLevel = 0.1
LeftPlasticRampRollSoundFactor = 0.2
RightPlasticRampRollSoundFactor = 0.2
LeftMetalWireRampRollSoundFactor = 1
RightPlasticRampHitsSoundLevel = 1
LeftPlasticRampHitsSoundLevel = 1
SpinningDiscRolloverSoundFactor = 0.05
SpinningDiscRolloverBumpSoundLevel = 0.3
LaneEnterSoundFactor = 0.9
InnerLaneSoundFactor = 0.0005
LaneSoundFactor = 0.0004
LaneLoudImpactMinimumSoundLevel = 0
LaneLoudImpactMaximumSoundLevel = 0.4


'///////////////////////////  CABINET SOUND PARAMETERS  /////////////////////////
Dim NudgeLeftSoundLevel, NudgeRightSoundLevel, NudgeCenterSoundLevel, StartButtonSoundLevel
Dim PlungerReleaseSoundLevel, PlungerPullSoundLevel, CoinSoundLevel

NudgeLeftSoundLevel = 1
NudgeRightSoundLevel = 1
NudgeCenterSoundLevel = 1
StartButtonSoundLevel = 0.1
PlungerReleaseSoundLevel = 1
PlungerPullSoundLevel = 1
CoinSoundLevel = 1


'///////////////////////////  MISC SOUND PARAMETERS  ////////////////////////////
Dim LutToggleSoundLevel :
LutToggleSoundLevel = 0.5


'////////////////////////////////  SOUND HELPERS  ///////////////////////////////
Dim SoundOn : SoundOn = 1
Dim SoundOff : SoundOff = 0
Dim Up : Up = 0
Dim Down : Down = 1
Dim RampUp : RampUp = 1
Dim RampDown : RampDown = 0
Dim RampDownSlow : RampDownSlow = 1
Dim RampDownFast : RampDownFast = 2
Dim CircuitA : CircuitA = 0
Dim CircuitC : CircuitC = 1

'//  Helper for Main (Lower) flippers dampened stroke
Dim BallNearLF : BallNearLF = 0
Dim BallNearRF : BallNearRF = 0

Sub TriggerBallNearLF_Hit()
  'Debug.Print "BallNearLF = 1"
  BallNearLF = 1
End Sub

Sub TriggerBallNearLF_UnHit()
  'Debug.Print "BallNearLF = 0"
  BallNearLF = 0
End Sub

Sub TriggerBallNearRF_Hit()
  'Debug.Print "BallNearRF = 1"
  BallNearRF = 1
End Sub

Sub TriggerBallNearRF_UnHit()
  'Debug.Print "BallNearLF = 0"
  BallNearRF = 0
End Sub


'///////////////////////  SOUND PLAYBACK SUBS / FUNCTIONS  //////////////////////
'//////////////////////  POSITIONAL SOUND PLAYBACK METHODS  /////////////////////
'//  Positional sound playback methods will play a sound, depending on the X,Y position of the table element or depending on ActiveBall object position
'//  These are similar subroutines that are less complicated to use (e.g. simply use standard parameters for the PlaySound call)
'//  For surround setup - positional sound playback functions will fade between front and rear surround channels and pan between left and right channels
'//  For stereo setup - positional sound playback functions will only pan between left and right channels
'//  For mono setup - positional sound playback functions will not pan between left and right channels and will not fade between front and rear channels
'//
'//  PlaySound full syntax - PlaySound(string, int loopcount, float volume, float pan, float randompitch, int pitch, bool useexisting, bool restart, float front_rear_fade)
'//  Note - These functions will not work (currently) for walls/slingshots as these do not feature a simple, single X,Y position

Sub PlaySoundAtLevelStatic(playsoundparams, aVol, tableobj)
    PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelExistingStatic(playsoundparams, aVol, tableobj)
    PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticLoop(playsoundparams, aVol, tableobj)
    PlaySound playsoundparams, -1, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelExistingStaticLoop(playsoundparams, aVol, tableobj)
    PlaySound playsoundparams, -1, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
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
    PlaySound playsoundparams, -1, aVol * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub


'//////////////////////  SUPPORTING BALL & SOUND FUNCTIONS  /////////////////////
'//  Fades between front and back of the table
'//  (for surround systems or 2x2 speakers, etc), depending on the Y position
'//  on the table.
Function AudioFade(tableobj)
  Dim tmp
  Select Case PositionalSoundPlaybackConfiguration
    Case 1
      AudioFade = 0
    Case 2
      AudioFade = 0
    Case 3
      tmp = tableobj.y * 2 / tableheight-1
      If tmp > 0 Then
        AudioFade = Csng(tmp ^10)
      Else
        AudioFade = Csng(-((- tmp) ^10) )
      End If
  End Select
End Function

'//  Calculates the pan for a tableobj based on the X position on the table.
Function AudioPan(tableobj)
  Dim tmp
  Select Case PositionalSoundPlaybackConfiguration
    Case 1
      AudioPan = 0
    Case 2
      tmp = tableobj.x * 2 / tablewidth-1
      If tmp > 0 Then
        AudioPan = Csng(tmp ^10)
      Else
        AudioPan = Csng(-((- tmp) ^10) )
      End If
    Case 3
      tmp = tableobj.x * 2 / tablewidth-1

      If tmp > 0 Then
        AudioPan = Csng(tmp ^10)
      Else
        AudioPan = Csng(-((- tmp) ^10) )
      End If
  End Select
End Function

'//  Calculates the volume of the sound based on the ball speed
Function Vol(ball)
  Vol = Csng(BallVel(ball) ^2)
End Function

'//  Calculates the pitch of the sound based on the ball speed
Function Pitch(ball)
    Pitch = BallVel(ball) * 20
End Function

'//  Calculates the ball speed
Function BallVel(ball)
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

'//  Calculates the roll volume of the sound based on the ball speed
Dim TempBallVel
Function VolPlayfieldRoll(ball) ' Calculates the roll volume of the sound based on the ball speed
  TempBallVel = Csng((INT(SQR((ball.VelX^2)+(ball.VelY^2))))/RollingSoundFactor)
  If TempBallVel = 1 Then TempBallVel = 0.999
  If TempBallVel = 0 Then TempBallVel = 0.001
  'debug.print TempBallVel
  TempBallVel = Csng(1/(1+(0.275*(((0.75*TempBallVel)/(1-TempBallVel))^(-2)))))
  VolPlayfieldRoll = TempBallVel
End Function

'//  Calculates the roll volume of the sound based on the ball speed
Function VolSpinningDiscRoll(ball)
  VolSpinningDiscRoll = RollingOnDiscSoundFactor * 0.1 * Csng(BallVel(ball) ^3)
End Function

'//  Calculates the roll volume of the sound based on the ball speed
Dim TempBallVelPlastic
Function VolPlasticMetalRampRoll(ball)
  'VolPlasticMetalRampRoll = RollingOnDiscSoundFactor * 0.0005 * Csng(BallVel(ball) ^3)
  TempBallVelPlastic = Csng((INT(SQR((ball.VelX^2)+(ball.VelY^2))))/RollingSoundFactor)
  If TempBallVelPlastic = 1 Then TempBallVelPlastic = 0.999
  If TempBallVelPlastic = 0 Then TempBallVelPlastic = 0.001
  'debug.print TempBallVel
  TempBallVelPlastic = Csng(1/(1+(0.275*(((0.75*TempBallVelPlastic)/(1-TempBallVelPlastic))^(-2)))))
  VolPlasticMetalRampRoll = TempBallVelPlastic



End Function

'//  Calculates the roll pitch of the sound based on the ball speed
Dim TempPitchBallVel
Function PitchPlayfieldRoll(ball) ' Calculates the roll pitch of the sound based on the ball speed
  'PitchPlayfieldRoll = BallVel(ball) ^2 * 15
  'PitchPlayfieldRoll = Csng(BallVel(ball))/50 * 10000
  'PitchPlayfieldRoll = (1-((Csng(BallVel(ball))/50)^0.2)) * 20000

  'PitchPlayfieldRoll = (2*((Csng(BallVel(ball)))^0.7))/(2+(Csng(BallVel(ball)))) * 16000
  TempPitchBallVel = Csng((INT(SQR((ball.VelX^2)+(ball.VelY^2))))/50)
  If TempPitchBallVel = 1 Then TempPitchBallVel = 0.999
  If TempPitchBallVel = 0 Then TempPitchBallVel = 0.001
  TempPitchBallVel = Csng(1/(1+(0.275*(((0.75*TempPitchBallVel)/(1-TempPitchBallVel))^(-2))))) * 10000
  PitchPlayfieldRoll = TempPitchBallVel
End Function

'//  Calculates the pitch of the sound based on the ball speed.
'//  Used for plastic ramps roll sound
Function PitchPlasticRamp(ball)
    PitchPlasticRamp = BallVel(ball) * 20
End Function

'//  Determines if a Points (px,py) is inside a 4 point polygon A-D
'//  in Clockwise/CCW order
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

'//  Determines if a point (px,py) in inside a circle with a center of
'//  (cx,cy) coordinates and circleradius
Function InCircle(px,py,cx,cy,circleradius)
  Dim distance
  distance = SQR(((px-cx)^2) + ((py-cy)^2))

  If (distance < circleradius) Then
    InCircle = True
  Else
    InCircle = False
  End If
End Function

'///////////////////////////  PLAY SOUNDS SUBROUTINES  //////////////////////////
'//
'//  These Subroutines implement all mechanical playsounds including timers
'//
'//////////////////////////  GENERAL SOUND SUBROUTINES  /////////////////////////
Sub SoundStartButton()
  PlaySoundAtLevelStatic (Cartridge_Cabinet_Sounds & "_Start_Button"), StartButtonSoundLevel, StartButtonPosition
End Sub

Sub SoundPlungerPull()
  PlaySoundAtLevelStatic (Cartridge_Cabinet_Sounds & "_Plunger_Pull_Slow"), PlungerPullSoundLevel, Plunger
End Sub

Sub SoundPlungerPullStop()
  StopSound Cartridge_Cabinet_Sounds & "_Plunger_Pull_Slow"
End Sub

Sub SoundPlungerReleaseBall()
  PlaySoundAtLevelStatic (Cartridge_Cabinet_Sounds & "_Plunger_Release_Ball_" & Int(Rnd*3)+1), PlungerReleaseSoundLevel, Plunger
End Sub

Sub SoundPlungerReleaseNoBall()
  PlaySoundAtLevelStatic (Cartridge_Cabinet_Sounds & "_Plunger_Release_Empty"), PlungerReleaseSoundLevel, Plunger
End Sub

Sub SoundNudgeLeft()
  PlaySound ("Nudge_" & Int(Rnd*3)+1), 0, NudgeLeftSoundLevel * VolumeDial, -0.1, 0.25
End Sub

Sub SoundNudgeRight()
  PlaySound ("Nudge_" & Int(Rnd*3)+1), 0, NudgeRightSoundLevel * VolumeDial, 0.1, 0.25
End Sub

Sub SoundNudgeCenter()
  PlaySoundAtLevelStatic ("Nudge_" & Int(Rnd*3)+1), NudgeCenterSoundLevel * VolumeDial, sw10
End Sub



'/////////////////////////////  PLASTIC RAMP TIMERS  ////////////////////////////
'Sub PlasticRampTimer1_Timer()
' ' 4 point polygon that contains a complete right plastic ramp
' If ballvariablePlasticRampTimer1 > 1 and InRect(ballvariablePlasticRampTimer1.x, ballvariablePlasticRampTimer1.y, 232,1556,3,1556,3,3,961,168) and Not InRect(ballvariablePlasticRampTimer1.x, ballvariablePlasticRampTimer1.y, 720,377,801,389,767,668,658,649) And Not InRect(ballvariablePlasticRampTimer1.x, ballvariablePlasticRampTimer1.y, 3,3,950,3,950,1512,3,350) and Not InRect(ballvariablePlasticRampTimer1.x, ballvariablePlasticRampTimer1.y, 350,572,292,659,66,472,123,293) And Not InRect(ballvariablePlasticRampTimer1.x, ballvariablePlasticRampTimer1.y, 950,1280,850,1280,850,2095,950,2095) Then
'   PlaySound (Cartridge_Plastic_Ramps & "_Ramp_Right_Plastic_BallRoll_" & b), 0, (VolPlasticMetalRampRoll(BOT(b)))/10 * VolumeDial, AudioPan(BOT(b)), 0, 0, 1, 0, AudioFade(BOT(b))
' Else
'   Me.Enabled = 0
' End If
'End Sub


'//////  JP'S VP10 ROLLING SOUNDS & FLEEP RAMPS / SPINNING DISC ROLLOVER  ///////
Const tnob = 11 ' total number of balls
ReDim rolling(tnob)
ReDim ramprolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
  ramprolling(i) = False
    Next
End Sub

Dim SpinningDiscRollover_Left_Up, SpinningDiscRollover_Left_Down
Dim SpinningDiscRollover_Middle_Up, SpinningDiscRollover_Middle_Down
Dim SpinningDiscRollover_Right_Up, SpinningDiscRollover_Right_Down

SpinningDiscRollover_Left_Up = 0
SpinningDiscRollover_Left_Down = 0
SpinningDiscRollover_Middle_Up = 0
SpinningDiscRollover_Middle_Down = 0
SpinningDiscRollover_Right_Up = 0
SpinningDiscRollover_Right_Down = 0

Sub RollingSoundUpdate()
  Dim BOT, b
  BOT = GetBalls
  For b = 0 to UBound(BOT)
      If BOT(b).z < 27 and BOT(b).z > 23 Then

        ' play the rolling sound for each ball
        If BallVel(BOT(b) ) > 0.01 Then
          rolling(b) = True
          PlaySound (Cartridge_Ball_Roll & "_Ball_Roll_" & b), -1, VolPlayfieldRoll(BOT(b)) * PlayfieldRollVolumeDial * VolumeDial * 1.25, AudioPan(BOT(b)), 0, PitchPlayfieldRoll(BOT(b)), 1, 0, AudioFade(BOT(b))

          ' play rolling sound on left disc as well
          If InCircle(BOT(b).x, BOT(b).y, 353, 1226, 60) Then
            If BOT(b).VelY < 0 Then
              'Ball rolls upwards
              SpinningDiscRollover_Left_Up = 1
              PlaySound Cartridge_Table_Specifics & "_Spinning_Disc_Left_Ball_Rollover_RollUp", 1, VolSpinningDiscRoll(BOT(b)) * SpinningDiscRolloverSoundFactor * VolumeDial, AudioPan(BOT(b)), 0, 0, 1, 0, AudioFade(BOT(b))
            Else
              'Ball rolls downwards
              SpinningDiscRollover_Left_Down = 1
              PlaySound Cartridge_Table_Specifics & "_Spinning_Disc_Left_Ball_Rollover_RollDown", 1, VolSpinningDiscRoll(BOT(b)) * SpinningDiscRolloverSoundFactor * VolumeDial, AudioPan(BOT(b)), 0, 0, 1, 0, AudioFade(BOT(b))
            End If
          Else
            If SpinningDiscRollover_Left_Up Then
              SoundSpinningDiscRolloverBumpLeft()
              SpinningDiscRollover_Left_Up = 0
              StopSound Cartridge_Table_Specifics & "_Spinning_Disc_Left_Ball_Rollover_RollUp"
            End If
            If SpinningDiscRollover_Left_Down Then
              SoundSpinningDiscRolloverBumpLeft()
              SpinningDiscRollover_Left_Down = 0
              StopSound Cartridge_Table_Specifics & "_Spinning_Disc_Left_Ball_Rollover_RollDown"
            End If
          End If

          ' play rolling sound on middle disc as well
          If InCircle(BOT(b).x, BOT(b).y, 533, 1184, 95) Then
            If BOT(b).VelY < 0 Then
              'Ball rolls upwards
              SpinningDiscRollover_Middle_Up = 1
              PlaySound Cartridge_Table_Specifics & "_Spinning_Disc_Middle_Ball_Rollover_RollUp", 1, VolSpinningDiscRoll(BOT(b)) * SpinningDiscRolloverSoundFactor * VolumeDial, AudioPan(BOT(b)), 0, 0, 1, 0, AudioFade(BOT(b))
            Else
              'Ball rolls downwards
              SpinningDiscRollover_Middle_Down = 1
              PlaySound Cartridge_Table_Specifics & "_Spinning_Disc_Middle_Ball_Rollover_RollDown", 1, VolSpinningDiscRoll(BOT(b)) * SpinningDiscRolloverSoundFactor * VolumeDial, AudioPan(BOT(b)), 0, 0, 1, 0, AudioFade(BOT(b))
            End If
          Else
            If SpinningDiscRollover_Middle_Up Then
              SoundSpinningDiscRolloverBumpMiddle()
              SpinningDiscRollover_Middle_Up = 0
              StopSound Cartridge_Table_Specifics & "_Spinning_Disc_Middle_Ball_Rollover_RollUp"
            End If
            If SpinningDiscRollover_Middle_Down Then
              SoundSpinningDiscRolloverBumpMiddle()
              SpinningDiscRollover_Middle_Down = 0
              StopSound Cartridge_Table_Specifics & "_Spinning_Disc_Middle_Ball_Rollover_RollDown"
            End If
          End If

          ' play rolling sound on right disc as well
          If InCircle(BOT(b).x, BOT(b).y, 715, 1254, 70) Then
            If BOT(b).VelY < 0 Then
              'Ball rolls upwards
              SpinningDiscRollover_Right_Up = 1
              PlaySound Cartridge_Table_Specifics & "_Spinning_Disc_Right_Ball_Rollover_RollUp", 1, VolSpinningDiscRoll(BOT(b)) * SpinningDiscRolloverSoundFactor * VolumeDial, AudioPan(BOT(b)), 0, 0, 1, 0, AudioFade(BOT(b))
            Else
              'Ball rolls downwards
              SpinningDiscRollover_Right_Down = 1
              PlaySound Cartridge_Table_Specifics & "_Spinning_Disc_Right_Ball_Rollover_RollDown", 1, VolSpinningDiscRoll(BOT(b)) * SpinningDiscRolloverSoundFactor * VolumeDial, AudioPan(BOT(b)), 0, 0, 1, 0, AudioFade(BOT(b))
            End If
          Else
            If SpinningDiscRollover_Right_Up Then
              SoundSpinningDiscRolloverBumpRight()
              SpinningDiscRollover_Right_Up = 0
              StopSound Cartridge_Table_Specifics & "_Spinning_Disc_Right_Ball_Rollover_RollUp"
            End If
            If SpinningDiscRollover_Right_Down Then
              SoundSpinningDiscRolloverBumpRight()
              SpinningDiscRollover_Right_Down = 0
              StopSound Cartridge_Table_Specifics & "_Spinning_Disc_Right_Ball_Rollover_RollDown"
            End If
          End If
        End If

      Else
        If BOT(b).z > 46 Then

          ' Play the right plastic ramp rolling sound for each ball
          If BallVel(BOT(b) ) > 1 and InRect(BOT(b).x, BOT(b).y, 232,1483,3,1483,3,3,961,168) and Not InRect(BOT(b).x, BOT(b).y, 720,377,801,389,767,668,658,649) And Not InRect(BOT(b).x, BOT(b).y, 3,3,950,3,950,1512,3,350) and Not InRect(BOT(b).x, BOT(b).y, 350,572,292,659,66,472,123,293) And Not InRect(BOT(b).x, BOT(b).y, 950,1280,850,1280,850,2095,950,2095) Then
            ramprolling(b) = True
            PlaySound (Cartridge_Plastic_Ramps & "_Ramp_Right_Plastic_BallRoll_" & b), 0, (VolPlasticMetalRampRoll(BOT(b))) * RightPlasticRampRollSoundFactor * RampsRollVolumeDial * VolumeDial * 1.25, AudioPan(BOT(b)), 0, 0, 1, 0, AudioFade(BOT(b))
            'Debug.Print Cartridge_Plastic_Ramps & "_Ramp_Right_Plastic_BallRoll_" & b

          End If

          ' Play the left ramp - plastic part sound for each ball
          If BallVel(BOT(b) ) > 1 and InRect(BOT(b).x, BOT(b).y, 3,3,950,3,950,1512,3,350) and Not InRect(BOT(b).x, BOT(b).y, 950,352,806,349,715,1232,950,1396) and Not InRect(BOT(b).x, BOT(b).y, 950,1280,850,1280,850,2095,950,2095) and Not InRect(BOT(b).x, BOT(b).y, 350,572,292,659,66,472,123,293) Then
            ramprolling(b) = True
            PlaySound (Cartridge_Plastic_Ramps & "_Ramp_Left_Plastic_BallRoll_" & b), 0, (VolPlasticMetalRampRoll(BOT(b))) * LeftPlasticRampRollSoundFactor * RampsRollVolumeDial * VolumeDial * 1.25, AudioPan(BOT(b)), 0, 0, 1, 0, AudioFade(BOT(b))
            'Debug.Print Cartridge_Plastic_Ramps & "_Ramp_Left_Plastic_BallRoll_" & b
          End If

          ' Play the left ramp - wire part sound for each ball
          If BallVel(BOT(b) ) > 1 and InRect(BOT(b).x, BOT(b).y, 950,352,806,349,715,1232,950,1396) and Not InRect(BOT(b).x, BOT(b).y, 950,1280,850,1280,850,2095,950,2095) Then
            ramprolling(b) = True
            StopSound (Cartridge_Metal_Ramps & "_Ramp_Left_Plastic_BallRoll_" & b)
            PlaySound (Cartridge_Metal_Ramps & "_Ramp_Left_Metal_Wire_BallRoll_" & b), 0, (VolPlasticMetalRampRoll(BOT(b))) * LeftMetalWireRampRollSoundFactor * RampsRollVolumeDial * VolumeDial * 1.25, AudioPan(BOT(b)), 0, 0, 1, 0, AudioFade(BOT(b))
            'Debug.Print Cartridge_Metal_Ramps & "_Ramp_Left_Metal_Wire_BallRoll_" & b
          End If
        Else
          If ramprolling(b) = True Then
            ramprolling(b) = False
            StopSound (Cartridge_Plastic_Ramps & "_Ramp_Left_Plastic_BallRoll_" & b)
            StopSound (Cartridge_Metal_Ramps & "_Ramp_Left_Metal_Wire_BallRoll_" & b)
            StopSound (Cartridge_Plastic_Ramps & "_Ramp_Right_Plastic_BallRoll_" & b)
          End If
        End If
      End If

      If rolling(b) = True and (BallVel(BOT(b) ) <= 1 or BOT(B).z < 23 or BOT(b).z > 27) Then
        StopSound (Cartridge_Ball_Roll & "_Ball_Roll_" & b)
        rolling(b) = False
      End If


      If ramprolling(b) = True and BOT(b).z <= 46 Then
        ramprolling(b) = False
        StopSound (Cartridge_Plastic_Ramps & "_Ramp_Left_Plastic_BallRoll_" & b)
        StopSound (Cartridge_Metal_Ramps & "_Ramp_Left_Metal_Wire_BallRoll_" & b)
        StopSound (Cartridge_Plastic_Ramps & "_Ramp_Right_Plastic_BallRoll_" & b)
      End If

      '***Ball Shadows***
      If BallShadowOn = 1 Then
        BallShadow(b).X = BOT(b).X
        ballShadow(b).Y = BOT(b).Y + 10
      End If

      If BOT(b).Z > 24 and BOT(b).Z < 35 and BOT(b).radius > 23  and not inrect(BOT(b).x,BOT(b).y,183,925,227,925,227,969,183,969) Then
        BallShadow(b).visible = 1
      Else
        BallShadow(b).visible = 0
      End If
  Next
End Sub


'///////////////////////  JP'S VP10 BALL COLLISION SOUND  ///////////////////////
Sub OnBallBallCollision(ball1, ball2, velocity)
  PlaySound (Cartridge_BallBallCollision & "_BallBall_Collide_" & Int(Rnd*7)+1), 0, Csng(velocity) ^2 / 200 * BallWithBallCollisionSoundFactor * VolumeDial, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub



'///////////////////////////////  CELLAR SOUNDS  ///////////////////////////////
Sub SoundCellarLeftEnter()
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  PlaySoundAtLevelStatic (Cartridge_Kickers & "_Scoop_Left_Cellar_Enter_" & Int(Rnd*4)+1), CellarLeftEnterSoundLevel * finalspeed/40, CellarLeftTrigger
End Sub

Sub SoundCellarRightEnter()
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  PlaySoundAtLevelStatic (Cartridge_Kickers & "_Scoop_Right_Cellar_Enter_" & Int(Rnd*5)+1), CellarRightEnterSoundLevel * finalspeed/40, CellarRightTrigger
End Sub


Sub SoundCellerKickout()
  PlaySoundAtLevelStatic SoundFX(Cartridge_Kickers & "_Scoop_Kickout_Cellar_" & Int(Rnd*3)+1,DOFContactors), Solenoid_UnderPlayfieldKickbig_SoundLevel, ScoopKickerOverflow
End Sub

Sub SoundCellarKickoutBallDrop()
  PlaySoundAtLevelStatic (Cartridge_Kickers & "_Scoop_BallDrop_After_Kickout_" & Int(Rnd*2)+1), CellerKickouBallDroptSoundLevel, ScoopKickerOverflow
End Sub



'///////////////////////////  GENERAL ROLLOVER SOUNDS  //////////////////////////
Sub RandomSoundRollover()
  PlaySoundAtLevelActiveBall (Cartridge_Rollovers & "_Rollover_All_" & Int(Rnd*4)+1), RolloverSoundLevel
End Sub



'///////////////////////////  OUTLANE ROLLOVER SOUNDS  //////////////////////////
Sub RandomSoundOutlaneRollover()
  PlaySoundAtLevelActiveBall (Cartridge_Rollovers & "_Rollover_Outlane_" & Int(Rnd*4)+1), OutLaneRolloverSoundLevel
End Sub



'////////////////////  BALL GATES AND BRACKET GATES SOUNDS  /////////////////////
Sub SoundRampBrktGate1(toggle)
  If toggle = SoundOn Then
    PlaySoundAtLevelStatic (Cartridge_Gates & "_Bracket_Gate_1"), Switch_Gate_SoundLevel * 0.005 , sw33g
  End If
  If toggle = SoundOff Then
    Stopsound Cartridge_Gates & "_Bracket_Gate_1"
  End If
End Sub

Sub SoundRampBrktGate2(toggle)
  If toggle = SoundOn Then
    PlaySoundAtLevelStatic (Cartridge_Gates & "_Bracket_Gate_2"), Switch_Gate_SoundLevel * 0.005, sw40g
  End If
  If toggle = SoundOff Then
    Stopsound Cartridge_Gates & "_Bracket_Gate_2"
  End If
End Sub

Sub SoundBallGate1()
  PlaySoundAtLevelStatic (Cartridge_Gates & "_Oneway_Ball_Gate_" & Int(Rnd*3)+1), Switch_Gate_SoundLevel, Gate1
End Sub

Sub SoundBallGate3()
  PlaySoundAtLevelStatic (Cartridge_Gates & "_Oneway_Ball_Gate_" & Int(Rnd*3)+1), Switch_Gate_SoundLevel, Gate3
End Sub

Sub SoundBallReleaseGate()
  PlaySoundAtLevelStatic (Cartridge_Gates & "_Oneway_Ball_Gate_" & Int(Rnd*3)+1), Switch_Gate_SoundLevel * 0.5, BallReleaseGate
End Sub



'///////////////////////////  DROP TARGET HIT SOUNDS  ///////////////////////////
Sub SoundDropTargetsw26()
  PlaySoundAtLevelStatic (Cartridge_Targets & "_Drop_Target_1Bank_Release_Down_" & Int(Rnd*2)+1), Vol(ActiveBall) * TargetSoundFactor, sw26
End Sub

Sub SoundDropTargetsw27()
  PlaySoundAtLevelStatic (Cartridge_Targets & "_Drop_Target_3Bank_Release_Down_" & Int(Rnd*6)+1), Vol(ActiveBall) * TargetSoundFactor, sw27p
End Sub

Sub SoundDropTargetsw28()
  PlaySoundAtLevelStatic (Cartridge_Targets & "_Drop_Target_3Bank_Release_Down_" & Int(Rnd*6)+1), Vol(ActiveBall) * TargetSoundFactor, sw28p
End Sub

Sub SoundDropTargetsw29()
  PlaySoundAtLevelStatic (Cartridge_Targets & "_Drop_Target_3Bank_Release_Down_" & Int(Rnd*6)+1), Vol(ActiveBall) * TargetSoundFactor, sw29p
End Sub



'/////////////////////  DROP TARGET RESET SOLENOID SOUNDS  //////////////////////
Sub RandomSoundDropTargetTop()
  PlaySoundAtLevelStatic SoundFX(Cartridge_Targets & "_Drop_Target_1Bank_Reset_Up_" & Int(Rnd*8)+1,DOFContactors), Solenoid_DropTargetReset_SoundLevel, sw26
End Sub

Sub RandomSoundDropTargetLeft()
  PlaySoundAtLevelStatic SoundFX(Cartridge_Targets & "_Drop_Target_3Bank_Reset_Up_" & Int(Rnd*6)+1,DOFContactors), Solenoid_DropTargetReset_SoundLevel, sw28p
End Sub



'///////////////////////////////////  SPINNER  //////////////////////////////////
Sub SoundSpinner()
  PlaySoundAtLevelStatic (Cartridge_Spinner & "_Spinner_Spin_Loop"), SpinnerSoundLevel, sw41
End Sub


'//////////////////////////  STADNING TARGET HIT SOUNDS  ////////////////////////
Sub RandomSoundTargetHitStrong()
  Select Case Int(Rnd*4)+1
    Case 1 : PlaySoundAtLevelActiveBall SoundFX(Cartridge_Targets & "_Target_Hit_5",DOFTargets), Vol(ActiveBall) * TargetSoundFactor
    Case 2 : PlaySoundAtLevelActiveBall SoundFX(Cartridge_Targets & "_Target_Hit_6",DOFTargets), Vol(ActiveBall) * TargetSoundFactor
    Case 3 : PlaySoundAtLevelActiveBall SoundFX(Cartridge_Targets & "_Target_Hit_7",DOFTargets), Vol(ActiveBall) * TargetSoundFactor
    Case 4 : PlaySoundAtLevelActiveBall SoundFX(Cartridge_Targets & "_Target_Hit_8",DOFTargets), Vol(ActiveBall) * TargetSoundFactor
  End Select
End Sub

Sub RandomSoundTargetHitWeak()
  Select Case Int(Rnd*4)+1
    Case 1 : PlaySoundAtLevelActiveBall SoundFX(Cartridge_Targets & "_Target_Hit_1",DOFTargets), Vol(ActiveBall) * TargetSoundFactor
    Case 2 : PlaySoundAtLevelActiveBall SoundFX(Cartridge_Targets & "_Target_Hit_2",DOFTargets), Vol(ActiveBall) * TargetSoundFactor
    Case 3 : PlaySoundAtLevelActiveBall SoundFX(Cartridge_Targets & "_Target_Hit_3",DOFTargets), Vol(ActiveBall) * TargetSoundFactor
    Case 4 : PlaySoundAtLevelActiveBall SoundFX(Cartridge_Targets & "_Target_Hit_4",DOFTargets), Vol(ActiveBall) * TargetSoundFactor
  End Select
End Sub



'/////////////////////////////  BALL BOUNCE SOUNDS  /////////////////////////////
Sub RandomSoundBallBouncePlayfieldSoft()
  Select Case Int(Rnd*7)+1
    Case 1 : PlaySoundAtLevelActiveBall (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_2"), Vol(ActiveBall) * BallBouncePlayfieldSoftFactor
    Case 2 : PlaySoundAtLevelActiveBall (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_12"), Vol(ActiveBall) * BallBouncePlayfieldSoftFactor
    Case 3 : PlaySoundAtLevelActiveBall (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_14"), Vol(ActiveBall) * BallBouncePlayfieldSoftFactor
    Case 4 : PlaySoundAtLevelActiveBall (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_18"), Vol(ActiveBall) * BallBouncePlayfieldSoftFactor
    Case 5 : PlaySoundAtLevelActiveBall (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_19"), Vol(ActiveBall) * BallBouncePlayfieldSoftFactor
    Case 6 : PlaySoundAtLevelActiveBall (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_20"), Vol(ActiveBall) * BallBouncePlayfieldSoftFactor
    Case 7 : PlaySoundAtLevelActiveBall (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_21"), Vol(ActiveBall) * BallBouncePlayfieldSoftFactor
  End Select
End Sub

Sub RandomSoundBallBouncePlayfieldHard()
  Select Case Int(Rnd*12)+1
    Case 1 : PlaySoundAtLevelActiveBall (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_1"), Vol(ActiveBall) * BallBouncePlayfieldHardFactor
    Case 2 : PlaySoundAtLevelActiveBall (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_3"), Vol(ActiveBall) * BallBouncePlayfieldHardFactor
    Case 3 : PlaySoundAtLevelActiveBall (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_7"), Vol(ActiveBall) * BallBouncePlayfieldHardFactor
    Case 4 : PlaySoundAtLevelActiveBall (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_8"), Vol(ActiveBall) * BallBouncePlayfieldHardFactor
    Case 5 : PlaySoundAtLevelActiveBall (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_9"), Vol(ActiveBall) * BallBouncePlayfieldHardFactor
    Case 6 : PlaySoundAtLevelActiveBall (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_11"), Vol(ActiveBall) * BallBouncePlayfieldHardFactor
    Case 7 : PlaySoundAtLevelActiveBall (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_13"), Vol(ActiveBall) * BallBouncePlayfieldHardFactor
    Case 8 : PlaySoundAtLevelActiveBall (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_15"), Vol(ActiveBall) * BallBouncePlayfieldHardFactor
    Case 9 : PlaySoundAtLevelActiveBall (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_16"), Vol(ActiveBall) * BallBouncePlayfieldHardFactor
    Case 10 : PlaySoundAtLevelActiveBall (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_17"), Vol(ActiveBall) * BallBouncePlayfieldHardFactor
    Case 11 : PlaySoundAtLevelActiveBall (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_22"), Vol(ActiveBall) * BallBouncePlayfieldHardFactor
    Case 12 : PlaySoundAtLevelActiveBall (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_23"), Vol(ActiveBall) * BallBouncePlayfieldHardFactor
  End Select
End Sub



'/////////////////////  RAMPS BALL DROP TO PLAYFIELD SOUNDS  ////////////////////
'/////////  METAL WIRE - RIGHT RAMP - EXIT HOLE TO PLAYFIELD - SOUNDS  //////////
Sub RandomSoundLeftRampDropToPlayfield()
  PlaySoundAtLevelStatic (Cartridge_Metal_Ramps & "_Ramp_Right_Metal_Wire_Drop_to_Playfield_" & Int(Rnd*2)+1), RightRampMetalWireDropToPlayfieldSoundLevel, RHelper3
End Sub


'//////////////  PLASTIC - LEFT RAMP - EXIT HOLE TO LOCK - SOUND  ///////////////
Sub RandomSoundRightRampLeftExitDropToLock()
  PlaySoundAtLevelStatic (Cartridge_Plastic_Ramps & "_Ramp_Right_Plastic_Drop_to_Lock_" & Int(Rnd*4)+1), LeftPlasticRampDropToLockSoundLevel, RHelper2
End Sub


'////////////  PLASTIC - LEFT RAMP - EXIT HOLE TO PLAYFIELD - SOUND  ////////////
Sub RandomSoundRightRampRightExitDropToPlayfield()
  PlaySoundAtLevelStatic (Cartridge_Plastic_Ramps & "_Ramp_Right_Plastic_Drop_to_Playfield_" & Int(Rnd*2)+1), LeftPlasticRampDropToPlayfieldSoundLevel, RHelper1
End Sub


'//////////////////  METAL WIRE RIGHT RAMP - 2ND PART RAMP SUB  /////////////////
'Sub SoundRightPlasticRampPart2(toggle, ballvariablePlasticRampTimer1)
' Set ballvariablePlasticRampTimer1 = ActiveBall
' Select Case toggle
'   Case SoundOn
'     PlasticRampTimer1.Interval = 10
'     PlasticRampTimer1.Enabled = 1
'   Case SoundOff
'     PlasticRampTimer1.Enabled = 0
'     StopSound Cartridge_Plastic_Ramps & "_Ramp_Right_Plastic_BallRoll_" & b
' End Select
'End Sub


'////////////////////////////  RAMP ENTRANCE EVENTS  ////////////////////////////
'/////////////////////////  RIGHT RAMP ENTRANCE SOUNDS  /////////////////////////
Sub RRAMPUP_Hit()
  ' Play the right plastic ramp entrance lifter/down sound for each ball
  If BallVel(ActiveBall) > 1 and ActiveBall.VelY < 0 Then
    'Ball rolls upwards
    PlaySoundAtLevelTimerExistingActiveBall (Cartridge_Plastic_Ramps & "_Ramp_Right_Plastic_Ramp_Enter_" & Int(Rnd*4)+1), RightPlasticRampEnteranceSoundLevel, ActiveBall
  End If
End Sub

Sub RRAMPDOWN_Hit()
  ' Play the right plastic ramp entrance lifter/down sound for each ball
  If BallVel(ActiveBall) > 1 and ActiveBall.VelY > 0 Then
    'Ball rolls downwards
    PlaySoundAtLevelTimerExistingActiveBall (Cartridge_Plastic_Ramps & "_Ramp_Right_Plastic_Ramp_RollBack_" & Int(Rnd*2)+1), RightPlasticRampEnteranceSoundLevel, ActiveBall
  End If
End Sub


'//////////////////////////  LEFT RAMP ENTRANCE SOUNDS  /////////////////////////
Sub LRAMPUP_Hit()
  ' Play the left ramp plastic ramp entrance for each ball
  If BallVel(ActiveBall) > 1 and ActiveBall.VelY < 0 Then
    'Ball rolls upwards
    PlaySoundAtLevelTimerExistingActiveBall (Cartridge_Plastic_Ramps & "_Ramp_Left_Plastic_Enter_" & Int(Rnd*4)+1), LeftPlasticRampEnteranceSoundLevel, ActiveBall
  End If
End Sub

Sub LRAMPDOWN_Hit()
  ' Play the left ramp plastic ramp entrance for each ball
  If BallVel(ActiveBall) > 1 and ActiveBall.VelY > 0 Then
    'Ball rolls downwards
    PlaySoundAtLevelTimerExistingActiveBall (Cartridge_Plastic_Ramps & "_Ramp_Left_Plastic_RollBack_" & Int(Rnd*2)+1), LeftPlasticRampEnteranceSoundLevel, ActiveBall
  End If
End Sub


'////////////////////////////////////  DRAIN  ///////////////////////////////////
'///////////////////////////////  OUTHOLE SOUNDS  ///////////////////////////////
Sub RandomSoundOutholeHit()
  PlaySoundAtLevelStatic (Cartridge_Trough & "_Outhole_Drain_Hit_" & Int(Rnd*4)+1), Solenoid_OutholeKicker_SoundLevel, sw10
End Sub

Sub RandomSoundOutholeKicker()
  PlaySoundAtLevelStatic SoundFX(Cartridge_Trough & "_Outhole_Kicker_" & Int(Rnd*4)+1,DOFContactors), Solenoid_OutholeKicker_SoundLevel, sw10
End Sub

'/////////////////////  BALL SHOOTER FEEDER SOLENOID SOUNDS  ////////////////////
Sub RandomSoundShooterFeeder()
  PlaySoundAtLevelStatic SoundFX(Cartridge_Trough & "_Shooter_Feeder_" & Int(Rnd*6)+1,DOFContactors), Solenoid_ShooterFeeder_SoundLevel, sw13
End Sub


'///////  SHOOTER LANE - BALL RELEASE ROLL IN SHOOTER LANE SOUND - SOUND  ///////
Sub SoundBallReleaseShooterLane(toggle)
  Select Case toggle
    Case SoundOn
      PlaySoundAtLevelActiveBall (Cartridge_Table_Specifics & "_Ball_Launch_from_Shooter_Lane"), BallReleaseShooterLaneSoundLevel
    Case SoundOff
      StopSound Cartridge_Table_Specifics & "_Ball_Launch_from_Shooter_Lane"
  End Select
End Sub


'//////////////////////////////  KNOCKER SOLENOID  //////////////////////////////
Sub KnockerSolenoid()
  'PlaySoundAtLevelStatic SoundFX(Cartridge_Knocker & "_Knocker_Coil",DOFKnocker), Solenoid_Knocker_SoundLevel, KnockerPosition
  PlaySound SoundFX(Cartridge_Knocker & "_Knocker_Coil",DOFKnocker), 0, Solenoid_Knocker_SoundLevel
End Sub


'/////////////////////////////  EJECT HOLD SOLENOID  ////////////////////////////
Sub RandomSoundEjectHoleSolenoid()
  PlaySoundAtLevelStatic (Cartridge_Kickers & "_Eject_Kickout_" & Int(Rnd*8)+1), Solenoid_TopEject_SoundLevel, sw43
End Sub


'///////////////////////////////  EJECT BALL BUMP  //////////////////////////////
Sub RandomSoundEjectBallBump()
  PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_After_Eject_" & Int(Rnd*3)+1), EjectBallBumpSoundLevel, EjectBallPosition
End Sub


'///////////////////////////  EJECT HOLD BALL ENTER  ////////////////////////////
Sub RandomSoundEjectHoleEnter()
  PlaySoundAtLevelStatic (Cartridge_Kickers & "_Eject_Enter_" & Int(Rnd*2)+1), EjectHoleEnterSoundLevel, sw43
End Sub


'///////////////////////////  LOCKING KICKER SOLENOID  //////////////////////////
Sub RandomSoundLockingKickerSolenoid()
  PlaySoundAtLevelStatic (Cartridge_Kickers & "_Locking_Kickback_" & Int(Rnd*4)+1), Solenoid_LeftLockingKickback_SoundLevel, LockingPosition
End Sub


'//////////////////////////  SLINGSHOT SOLENOID SOUNDS  /////////////////////////
Sub RandomSoundSlingshotLeft()
  PlaySoundAtLevelStatic SoundFX(Cartridge_Slingshots & "_Slingshot_Left_" & Int(Rnd*7)+1,DOFContactors), Solenoid_Slingshot_SoundLevel, LeftSlingshotPosition
End Sub

Sub RandomSoundSlingshotRight()
  PlaySoundAtLevelStatic SoundFX(Cartridge_Slingshots & "_Slingshot_Right_" & Int(Rnd*7)+1,DOFContactors), Solenoid_Slingshot_SoundLevel, RightSlingshotPosition
End Sub


'///////////////////////////  BUMPER SOLENOID SOUNDS  ///////////////////////////
'////////////////////////////////  BUMPERS - TOP  ///////////////////////////////
Sub RandomSoundBumperTopLeft()
  'Debug.Print Vol(ActiveBall) * Solenoid_Bumper_SoundMultiplier
  PlaySoundAtLevelStatic SoundFX(Cartridge_Bumpers & "_Jet_Bumper_Top_Left_" & Int(Rnd*5)+1,DOFContactors), Vol(ActiveBall) * Solenoid_Bumper_SoundMultiplier, Bumper1
End Sub

Sub RandomSoundBumperTopRight()
  'Debug.Print Vol(ActiveBall) * Solenoid_Bumper_SoundMultiplier
  PlaySoundAtLevelStatic SoundFX(Cartridge_Bumpers & "_Jet_Bumper_Top_Right_" & Int(Rnd*5)+1,DOFContactors), Vol(ActiveBall) * Solenoid_Bumper_SoundMultiplier, Bumper2
End Sub

Sub RandomSoundBumperTopLower()
  'Debug.Print Vol(ActiveBall) * Solenoid_Bumper_SoundMultiplier
  PlaySoundAtLevelStatic SoundFX(Cartridge_Bumpers & "_Jet_Bumper_Top_Low_" & Int(Rnd*7)+1,DOFContactors), Vol(ActiveBall) * Solenoid_Bumper_SoundMultiplier, Bumper3
End Sub


'///////////////////////////////  BUMPERS - BOTTOM  /////////////////////////////
Sub RandomSoundBumperBottomLeft()
  'Debug.Print Vol(ActiveBall) * Solenoid_Bumper_SoundMultiplier
  PlaySoundAtLevelStatic SoundFX(Cartridge_Bumpers & "_Jet_Bumper_Bottom_Left_" & Int(Rnd*6)+1,DOFContactors), Vol(ActiveBall) * Solenoid_Bumper_SoundMultiplier, Bumper4
End Sub

Sub RandomSoundBumperBottomRight()
  'Debug.Print Vol(ActiveBall) * Solenoid_Bumper_SoundMultiplier
  PlaySoundAtLevelStatic SoundFX(Cartridge_Bumpers & "_Jet_Bumper_Bottom_Right_" & Int(Rnd*5)+1,DOFContactors), Vol(ActiveBall) * Solenoid_Bumper_SoundMultiplier, Bumper5
End Sub

Sub RandomSoundBumperBottomTop()
  'Debug.Print Vol(ActiveBall) * Solenoid_Bumper_SoundMultiplier
  PlaySoundAtLevelStatic SoundFX(Cartridge_Bumpers & "_Jet_Bumper_Bottom_Top_" & Int(Rnd*5)+1,DOFContactors), Vol(ActiveBall) * Solenoid_Bumper_SoundMultiplier, Bumper6
End Sub


'///////////////////////  FLIPPER BATS SOUND SUBROUTINES  ///////////////////////
'/////////////////////  FLIPPER BATS SOLENOID ATTACK SOUND  /////////////////////
Sub SoundFlipperUpAttackLeft()
  FlipperUpAttackLeftSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
  PlaySoundAtLevelStatic (Cartridge_Flippers & "_Flipper_Attack_L"), FlipperUpAttackLeftSoundLevel, LeftFlipper
End Sub

Sub SoundFlipperUpAttackRight()
  FlipperUpAttackRightSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
    PlaySoundAtLevelStatic (Cartridge_Flippers & "_Flipper_Attack_R"), FlipperUpAttackLeftSoundLevel, RightFlipper
End Sub


'///////////////////////  FLIPPER BATS SOLENOID CORE SOUND  /////////////////////
Sub RandomSoundFlipperLowerLeftUpFullStroke()
  PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Lower_Left_Up_Full_Stroke_" & Int(Rnd*9)+1,DOFFlippers), FlipperLeftLowerHitParm, LeftFlipper
End Sub

Sub RandomSoundFlipperLowerRightUpFullStroke()
  PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Lower_Right_Up_Full_Stroke_" & Int(Rnd*8)+1,DOFFlippers), FlipperRightLowerHitParm, RightFlipper
End Sub

Sub RandomSoundFlipperLowerLeftUpDampenedStroke()
  PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Lower_Left_Up_Dampened_Stroke_" & Int(Rnd*7)+1,DOFFlippers), FlipperLeftLowerHitParm * 1.2, LeftFlipper
End Sub

Sub RandomSoundFlipperLowerRightUpDampenedStroke()
  PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Lower_Right_Up_Dampened_Stroke_" & Int(Rnd*7)+1,DOFFlippers), FlipperRightLowerHitParm * 1.2, RightFlipper
End Sub

Sub RandomSoundFlipperUpperRightUpFullStroke()
  PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Upper_Right_Up_Full_Stroke_" & Int(Rnd*8)+1,DOFFlippers), FlipperRightUpperHitParm, RightFlipper1
End Sub

Sub RandomSoundLowerLeftReflip()
  PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Lower_Left_Reflip_" & Int(Rnd*3)+1,DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, LeftFlipper
End Sub

Sub RandomSoundLowerRightReflip()
  PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Lower_Right_Reflip_" & Int(Rnd*2)+1,DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, RightFlipper
End Sub

Sub RandomSoundUpperRightReflip()
  PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Upper_Right_Reflip_" & Int(Rnd*3)+1,DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, RightFlipper1
End Sub

Sub RandomSoundLowerLeftQuickFlipUp()
  PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Lower_Left_QuickFlip_Up_" & Int(Rnd*3)+1,DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, LeftFlipper
End Sub

Sub RandomSoundLowerRightQuickFlipUp()
  PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Lower_Right_QuickFlip_Up_" & Int(Rnd*3)+1,DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, RightFlipper
End Sub

Sub RandomSoundUpperRightQuickFlipUp()
  PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Upper_Right_Reflip_" & Int(Rnd*3)+1,DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, RightFlipper1
End Sub

Sub RandomSoundLowerLeftDown()
  PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Lower_Left_Down_" & Int(Rnd*6)+1,DOFFlippers), FlipperDownSoundLevel, LeftFlipper
End Sub

Sub RandomSoundLowerRightDown()
  PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Lower_Right_Down_" & Int(Rnd*6)+1,DOFFlippers), FlipperDownSoundLevel, RightFlipper
End Sub

Sub RandomSoundUpperRightDown()
  PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Upper_Right_Down_" & Int(Rnd*4)+1,DOFFlippers), FlipperDownSoundLevel, RightFlipper1
End Sub

Sub StopAnyFlipperLowerLeftDown()
  StopSound Cartridge_Flippers & "_Flipper_Lower_Left_Down_1"
  StopSound Cartridge_Flippers & "_Flipper_Lower_Left_Down_2"
  StopSound Cartridge_Flippers & "_Flipper_Lower_Left_Down_3"
  StopSound Cartridge_Flippers & "_Flipper_Lower_Left_Down_4"
  StopSound Cartridge_Flippers & "_Flipper_Lower_Left_Down_5"
  StopSound Cartridge_Flippers & "_Flipper_Lower_Left_Down_6"
End Sub

Sub StopAnyFlipperLowerRightDown()
  StopSound Cartridge_Flippers & "_Flipper_Lower_Right_Down_1"
  StopSound Cartridge_Flippers & "_Flipper_Lower_Right_Down_2"
  StopSound Cartridge_Flippers & "_Flipper_Lower_Right_Down_3"
  StopSound Cartridge_Flippers & "_Flipper_Lower_Right_Down_4"
  StopSound Cartridge_Flippers & "_Flipper_Lower_Right_Down_5"
  StopSound Cartridge_Flippers & "_Flipper_Lower_Right_Down_6"
End Sub

Sub StopAnyFlipperUpperRightDown()
  StopSound Cartridge_Flippers & "_Flipper_Upper_Right_Down_1"
  StopSound Cartridge_Flippers & "_Flipper_Upper_Right_Down_2"
  StopSound Cartridge_Flippers & "_Flipper_Upper_Right_Down_3"
  StopSound Cartridge_Flippers & "_Flipper_Upper_Right_Down_4"
End Sub

Sub StopAnyFlipperLowerLeftUp()
  StopSound Cartridge_Flippers & "_Flipper_Lower_Left_Up_Full_Stroke_1"
  StopSound Cartridge_Flippers & "_Flipper_Lower_Left_Up_Full_Stroke_2"
  StopSound Cartridge_Flippers & "_Flipper_Lower_Left_Up_Full_Stroke_3"
  StopSound Cartridge_Flippers & "_Flipper_Lower_Left_Up_Full_Stroke_4"
  StopSound Cartridge_Flippers & "_Flipper_Lower_Left_Up_Full_Stroke_5"
  StopSound Cartridge_Flippers & "_Flipper_Lower_Left_Up_Full_Stroke_6"
  StopSound Cartridge_Flippers & "_Flipper_Lower_Left_Up_Full_Stroke_7"
  StopSound Cartridge_Flippers & "_Flipper_Lower_Left_Up_Full_Stroke_8"
  StopSound Cartridge_Flippers & "_Flipper_Lower_Left_Up_Full_Stroke_9"
End Sub

Sub StopAnyFlipperLowerRightUp()
  StopSound Cartridge_Flippers & "_Flipper_Lower_Right_Up_Full_Stroke_1"
  StopSound Cartridge_Flippers & "_Flipper_Lower_Right_Up_Full_Stroke_2"
  StopSound Cartridge_Flippers & "_Flipper_Lower_Right_Up_Full_Stroke_3"
  StopSound Cartridge_Flippers & "_Flipper_Lower_Right_Up_Full_Stroke_4"
  StopSound Cartridge_Flippers & "_Flipper_Lower_Right_Up_Full_Stroke_5"
  StopSound Cartridge_Flippers & "_Flipper_Lower_Right_Up_Full_Stroke_6"
  StopSound Cartridge_Flippers & "_Flipper_Lower_Right_Up_Full_Stroke_7"
  StopSound Cartridge_Flippers & "_Flipper_Lower_Right_Up_Full_Stroke_8"
End Sub

Sub StopAnyFlipperUpperRightUp()
  StopSound Cartridge_Flippers & "_Flipper_Upper_Right_Up_Full_Stroke_1"
  StopSound Cartridge_Flippers & "_Flipper_Upper_Right_Up_Full_Stroke_2"
  StopSound Cartridge_Flippers & "_Flipper_Upper_Right_Up_Full_Stroke_3"
  StopSound Cartridge_Flippers & "_Flipper_Upper_Right_Up_Full_Stroke_4"
  StopSound Cartridge_Flippers & "_Flipper_Upper_Right_Up_Full_Stroke_5"
  StopSound Cartridge_Flippers & "_Flipper_Upper_Right_Up_Full_Stroke_6"
  StopSound Cartridge_Flippers & "_Flipper_Upper_Right_Up_Full_Stroke_7"
  StopSound Cartridge_Flippers & "_Flipper_Upper_Right_Up_Full_Stroke_8"
End Sub


'///////////////////////  FLIPPER BATS BALL COLLIDE SOUND  //////////////////////
dim angdamp, veldamp
angdamp = 0.2
veldamp = 0.8

Sub LeftFlipper_Collide(parm)
  CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
  if RubberizerEnabled = 1 then Rubberizer(parm)
  if RubberizerEnabled = 2 then Rubberizer2(parm)

' If parm > 4  and leftflipper.currentangle < 90 Then
'   activeball.angmomx=activeball.angmomx*angdamp
'   activeball.angmomy=activeball.angmomy*angdamp
'   activeball.angmomz=activeball.angmomz*angdamp
'   If  activeball.velx > 0 Then activeball.velx = activeball.velx * veldamp
' End If


  If parm => 22 Then
    ' Strong hit safe values boundary
    ' Flipper stroke dampened
    FlipperLeftLowerHitParm = FlipperUpSoundLevel * 0.1
  Else
    If parm =< 1 Then
      ' Weak hit safe values boundary
      ' Flipper stroke full
      FlipperLeftLowerHitParm = FlipperUpSoundLevel
    Else
      ' Fully modulated hit
      FlipperLeftLowerHitParm = FlipperUpSoundLevel * (1-(parm/25))
    End If
  End If

  RandomSoundRubberFlipper(parm)
End Sub

Sub RightFlipper_Collide(parm)
  CheckLiveCatch Activeball, RightFlipper, RFCount, parm
  if RubberizerEnabled = 1 then Rubberizer(parm)
  if RubberizerEnabled = 2 then Rubberizer2(parm)

' If parm > 4  and rightflipper.currentangle > -90 Then
'   activeball.angmomx=activeball.angmomx*angdamp
'   activeball.angmomy=activeball.angmomy*angdamp
'   activeball.angmomz=activeball.angmomz*angdamp
'   If  activeball.velx < 0 Then activeball.velx = activeball.velx * veldamp
' End If

  If parm => 22 Then
    ' Strong hit safe values boundary
    ' Flipper stroke dampened
    FlipperRightLowerHitParm = FlipperUpSoundLevel * 0.1
  Else
    If parm =< 1 Then
      ' Weak hit safe values boundary
      ' Flipper stroke full
      FlipperRightLowerHitParm = FlipperUpSoundLevel
    Else
      ' Fully modulated hit
      FlipperRightLowerHitParm = FlipperUpSoundLevel * (1-(parm/25))
    End If
  End If

  RandomSoundRubberFlipper(parm)
End Sub

Sub RightFlipper1_Collide(parm)
' If parm > 4  and rightflipper1.currentangle > -90 Then
'   activeball.angmomx=activeball.angmomx*angdamp
'   activeball.angmomy=activeball.angmomy*angdamp
'   activeball.angmomz=activeball.angmomz*angdamp
'   If  activeball.velx < 0 Then activeball.velx = activeball.velx * veldamp
' End If

  If parm => 22 Then
    ' Strong hit safe values boundary
    ' Flipper stroke dampened
    FlipperRightUpperHitParm = FlipperUpSoundLevel * 0.1
  Else
    If parm =< 1 Then
      ' Weak hit safe values boundary
      ' Flipper stroke full
      FlipperRightUpperHitParm = FlipperUpSoundLevel
    Else
      ' Fully modulated hit
      FlipperRightUpperHitParm = FlipperUpSoundLevel * (1-(parm/25))
    End If
  End If

  RandomSoundRubberFlipper(parm)
End Sub


Sub RandomSoundRubberFlipper(parm)
  PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Flipper_Hit_" & Int(Rnd*7)+1), parm / 25 * RubberFlipperSoundFactor
End Sub


'////////////////////////////  RAMPS SOUND EVENTS  //////////////////////////////
'//////////////////////  PLASTIC RIGHT RAMP SOUND EVENTS  ///////////////////////
Sub RRHelper1_Hit()
  If ActiveBall.VelX < 0 And ActiveBall.VelY < 0 Then PlaySoundAtLevelStatic (Cartridge_Plastic_Ramps & "_Ramp_Right_Plastic_Hit_1"), RightPlasticRampHitsSoundLevel, RRHelper1 : 'Debug.print Cartridge_Plastic_Ramps & "_Ramp_Right_Plastic_Hit_1"
End Sub

Sub RRHelper2_Hit()
  If ActiveBall.VelX < 0 Then PlaySoundAtLevelStatic (Cartridge_Plastic_Ramps & "_Ramp_Right_Plastic_Hit_2"), RightPlasticRampHitsSoundLevel, RRHelper2 : 'Debug.print Cartridge_Plastic_Ramps & "_Ramp_Right_Plastic_Hit_2"
End Sub

Sub RRHelper3_Hit()
  If ActiveBall.VelY > 0 Then PlaySoundAtLevelStatic (Cartridge_Plastic_Ramps & "_Ramp_Right_Plastic_Hit_3"), RightPlasticRampHitsSoundLevel, RRHelper3 : 'Debug.print Cartridge_Plastic_Ramps & "_Ramp_Right_Plastic_Hit_3"
End Sub

Sub RRHelper4_Hit()
  If ActiveBall.VelY > 0 Then PlaySoundAtLevelStatic (Cartridge_Plastic_Ramps & "_Ramp_Right_Plastic_Hit_4"), RightPlasticRampHitsSoundLevel, RRHelper4 : 'Debug.print Cartridge_Plastic_Ramps & "_Ramp_Right_Plastic_Hit_4"
End Sub

Sub RRHelper5_Hit()
  If ActiveBall.VelY > 0 Then PlaySoundAtLevelStatic (Cartridge_Plastic_Ramps & "_Ramp_Right_Plastic_Hit_5"), RightPlasticRampHitsSoundLevel, RRHelper4 : 'Debug.print Cartridge_Plastic_Ramps & "_Ramp_Right_Plastic_Hit_5"
End Sub

'///////////////////////  PLASTIC LEFT RAMP SOUND EVENTS  ///////////////////////
Sub LRHelper1_Hit()
  If ActiveBall.VelY < 0 Then PlaySoundAtLevelStatic (Cartridge_Plastic_Ramps & "_Ramp_Left_Plastic_Hit_1"), LeftPlasticRampHitsSoundLevel, LRHelper1 : 'Debug.print Cartridge_Plastic_Ramps & "_Ramp_Left_Plastic_Hit_1"
End Sub

Sub LRHelper2_Hit()
  If ActiveBall.VelY < 0 And ActiveBall.VelX > 0 Then PlaySoundAtLevelStatic (Cartridge_Plastic_Ramps & "_Ramp_Left_Plastic_Hit_2"), LeftPlasticRampHitsSoundLevel, LRHelper2 : 'Debug.print Cartridge_Plastic_Ramps & "_Ramp_Left_Plastic_Hit_2"
End Sub

Sub LRHelper3_Hit()
  If ActiveBall.VelX > 0 Then PlaySoundAtLevelStatic (Cartridge_Plastic_Ramps & "_Ramp_Left_Plastic_Hit_3"), LeftPlasticRampHitsSoundLevel, LRHelper3 : 'Debug.print Cartridge_Plastic_Ramps & "_Ramp_Left_Plastic_Hit_3"
End Sub

Sub LRHelper4_Hit()
  If ActiveBall.VelY > 0 Then PlaySoundAtLevelStatic (Cartridge_Plastic_Ramps & "_Ramp_Left_Plastic_Hit_4"), LeftPlasticRampHitsSoundLevel, LRHelper4 : 'Debug.print Cartridge_Plastic_Ramps & "_Ramp_Left_Plastic_Hit_4"
End Sub


'///////////////////////  RIGHT RAMP ENTRY LIFTER SOLENOID  /////////////////////
Sub RandomSoundRightRampEntryLifter()
  PlaySoundAtLevelStatic (Cartridge_Diverters & "_Ramp_Up_" & Int(Rnd*4)+1), Solenoid_RightRampLifter_SoundLevel, RampEntryLifterPosition
End Sub


'///////////////////////  RIGHT RAMP ENTRY DOWN SOLENOID  ///////////////////////
Sub RandomSoundRightRampEntryDown()
  PlaySoundAtLevelStatic (Cartridge_Diverters & "_Ramp_Down_" & Int(Rnd*6)+1), Solenoid_RightRampDown_SoundLevel, RampEntryDownPosition
End Sub


'///////////////////////  RAMP DIVERTER SOLENOID - DIVERT  //////////////////////
Sub RandomSoundRampDiverterDivert()
  PlaySoundAtLevelStatic (Cartridge_Diverters & "_Diverter_Divert_" & Int(Rnd*4)+1), Solenoid_Diverter_Enabled_SoundLevel, DiverterPosition
End Sub

'////////////////////////  RAMP DIVERTER SOLENOID - BACK  ///////////////////////
Sub RandomSoundRampDiverterBack()
  PlaySoundAtLevelStatic (Cartridge_Diverters & "_Diverter_Back_" & Int(Rnd*4)+1), Solenoid_Diverter_Disabled_SoundLevel, DiverterPosition
End Sub


'////////////////////  RAMP DIVERTER SOLENOID - MAGNET SOUND  ///////////////////
Sub RandomSoundRampDiverterHold(toggle)
  Select Case toggle
    Case SoundOn
      PlaySoundAtLevelStaticLoop SoundFX(Cartridge_Diverters & "_Diverter_Hold_Loop",DOFShaker), Solenoid_Diverter_Hold_SoundLevel, DiverterPosition
    Case SoundOff
      StopSound Cartridge_Diverters & "_Diverter_Hold_Loop"
  End Select
End Sub


'////////////////////////////  SPINNING DISCS SOUNDS  ///////////////////////////
'///////////////////////  SPINNING DISCS MOTOR TRANSIENTS  //////////////////////
Sub SoundSpinWheelsMotorsTransients(toggle)
  Select Case toggle
    Case RampUp
      Select Case SpinWheelsMotorConfiguration
        Case 1:
        Case 2:
          PlaySoundAtLevelStatic SoundFX(Cartridge_Table_Specifics & "_Spin_Wheels_Motors_Start_" & Int(Rnd*13)+1,DOFShaker), 0, MIdSpinTrigger
        Case 3:
          PlaySoundAtLevelStatic (Cartridge_Table_Specifics & "_Spin_Wheels_Motors_Start_" & Int(Rnd*13)+1), Solenoid_SpinWheelsMotor_SoundLevel, MIdSpinTrigger
        Case 4:
          PlaySoundAtLevelStatic SoundFX(Cartridge_Table_Specifics & "_Spin_Wheels_Motors_Start_" & Int(Rnd*13)+1,DOFShaker), Solenoid_SpinWheelsMotor_SoundLevel, MIdSpinTrigger
      End Select
    Case RampDownSlow
      Select Case SpinWheelsMotorConfiguration
        Case 1:
        Case 2:
          PlaySoundAtLevelStatic SoundFX(Cartridge_Table_Specifics & "_Spin_Wheels_Motors_Stop_Slow_" & Int(Rnd*3)+1,DOFShaker), 0, MIdSpinTrigger
        Case 3:
          PlaySoundAtLevelStatic (Cartridge_Table_Specifics & "_Spin_Wheels_Motors_Stop_Slow_" & Int(Rnd*3)+1), Solenoid_SpinWheelsMotor_SoundLevel, MIdSpinTrigger
        Case 4:
          PlaySoundAtLevelStatic SoundFX(Cartridge_Table_Specifics & "_Spin_Wheels_Motors_Stop_Slow_" & Int(Rnd*3)+1,DOFShaker), Solenoid_SpinWheelsMotor_SoundLevel, MIdSpinTrigger
      End Select
    Case RampDownFast
      Select Case SpinWheelsMotorConfiguration
        Case 1:
        Case 2:
          PlaySoundAtLevelStatic SoundFX(Cartridge_Table_Specifics & "_Spin_Wheels_Motors_Stop_Fast_" & Int(Rnd*3)+1,DOFShaker), 0, MIdSpinTrigger
        Case 3:
          PlaySoundAtLevelStatic (Cartridge_Table_Specifics & "_Spin_Wheels_Motors_Stop_Fast_" & Int(Rnd*3)+1), Solenoid_SpinWheelsMotor_SoundLevel, MIdSpinTrigger
        Case 4:
          PlaySoundAtLevelStatic SoundFX(Cartridge_Table_Specifics & "_Spin_Wheels_Motors_Stop_Fast_" & Int(Rnd*3)+1,DOFShaker), Solenoid_SpinWheelsMotor_SoundLevel, MIdSpinTrigger
      End Select
  End Select
End Sub


'//////////////////////////  SPINNING DISCS MOTOR LOOP  /////////////////////////
Sub SoundSpinWheelsMotorsLoop(toggle)
  Select Case toggle
    Case SoundOn
      Select Case SpinWheelsMotorConfiguration
        Case 1:
        Case 2: PlaySoundAtLevelExistingStaticLoop SoundFX(Cartridge_Table_Specifics & "_Spin_Wheels_Motors_Loop_Regular",DOFShaker), 0, MIdSpinTrigger
        Case 3: PlaySoundAtLevelExistingStaticLoop (Cartridge_Table_Specifics & "_Spin_Wheels_Motors_Loop_Regular"), Solenoid_SpinWheelsMotor_SoundLevel, MIdSpinTrigger
        Case 4: PlaySoundAtLevelExistingStaticLoop SoundFX(Cartridge_Table_Specifics & "_Spin_Wheels_Motors_Loop_Regular",DOFShaker), Solenoid_SpinWheelsMotor_SoundLevel, MIdSpinTrigger
      End Select
    Case SoundOff
      Select Case SpinWheelsMotorConfiguration
        Case 1:
        Case 2: StopSound (Cartridge_Table_Specifics & "_Spin_Wheels_Motors_Loop_Regular")
        Case 3: StopSound (Cartridge_Table_Specifics & "_Spin_Wheels_Motors_Loop_Regular")
        Case 4: StopSound (Cartridge_Table_Specifics & "_Spin_Wheels_Motors_Loop_Regular")
      End Select
  End Select
End Sub


'/////////////////////  SPINNING DISCS ROLLOVER BUMP SOUNDS  ////////////////////
Sub SoundSpinningDiscRolloverBumpLeft()
  Select Case Int(Rnd*5)+1
    Case 1 : PlaySoundAtLevelStatic (Cartridge_Table_Specifics & "_Spinning_Disc_Ball_Rollover_Bump_1"), 1.00 * SpinningDiscRolloverBumpSoundLevel, LeftSpinTrigger
    Case 2 : PlaySoundAtLevelStatic (Cartridge_Table_Specifics & "_Spinning_Disc_Ball_Rollover_Bump_2"), 0.90 * SpinningDiscRolloverBumpSoundLevel, LeftSpinTrigger
    Case 3 : PlaySoundAtLevelStatic (Cartridge_Table_Specifics & "_Spinning_Disc_Ball_Rollover_Bump_3"), 0.75 * SpinningDiscRolloverBumpSoundLevel, LeftSpinTrigger
    Case 4 : PlaySoundAtLevelStatic (Cartridge_Table_Specifics & "_Spinning_Disc_Ball_Rollover_Bump_4"), 0.75 * SpinningDiscRolloverBumpSoundLevel, LeftSpinTrigger
    Case 5 : PlaySoundAtLevelStatic (Cartridge_Table_Specifics & "_Spinning_Disc_Ball_Rollover_Bump_5"), 0.95 * SpinningDiscRolloverBumpSoundLevel, LeftSpinTrigger
  End Select
End Sub

Sub SoundSpinningDiscRolloverBumpMiddle()
  Select Case Int(Rnd*5)+1
    Case 1 : PlaySoundAtLevelStatic (Cartridge_Table_Specifics & "_Spinning_Disc_Ball_Rollover_Bump_1"), 1.00 * SpinningDiscRolloverBumpSoundLevel, MIdSpinTrigger
    Case 2 : PlaySoundAtLevelStatic (Cartridge_Table_Specifics & "_Spinning_Disc_Ball_Rollover_Bump_2"), 0.90 * SpinningDiscRolloverBumpSoundLevel, MIdSpinTrigger
    Case 3 : PlaySoundAtLevelStatic (Cartridge_Table_Specifics & "_Spinning_Disc_Ball_Rollover_Bump_3"), 0.75 * SpinningDiscRolloverBumpSoundLevel, MIdSpinTrigger
    Case 4 : PlaySoundAtLevelStatic (Cartridge_Table_Specifics & "_Spinning_Disc_Ball_Rollover_Bump_4"), 0.75 * SpinningDiscRolloverBumpSoundLevel, MIdSpinTrigger
    Case 5 : PlaySoundAtLevelStatic (Cartridge_Table_Specifics & "_Spinning_Disc_Ball_Rollover_Bump_5"), 0.95 * SpinningDiscRolloverBumpSoundLevel, MIdSpinTrigger
  End Select
End Sub

Sub SoundSpinningDiscRolloverBumpRight()
  Select Case Int(Rnd*5)+1
    Case 1 : PlaySoundAtLevelStatic (Cartridge_Table_Specifics & "_Spinning_Disc_Ball_Rollover_Bump_1"), 1.00 * SpinningDiscRolloverBumpSoundLevel, RightSpinTrigger
    Case 2 : PlaySoundAtLevelStatic (Cartridge_Table_Specifics & "_Spinning_Disc_Ball_Rollover_Bump_2"), 0.90 * SpinningDiscRolloverBumpSoundLevel, RightSpinTrigger
    Case 3 : PlaySoundAtLevelStatic (Cartridge_Table_Specifics & "_Spinning_Disc_Ball_Rollover_Bump_3"), 0.75 * SpinningDiscRolloverBumpSoundLevel, RightSpinTrigger
    Case 4 : PlaySoundAtLevelStatic (Cartridge_Table_Specifics & "_Spinning_Disc_Ball_Rollover_Bump_4"), 0.75 * SpinningDiscRolloverBumpSoundLevel, RightSpinTrigger
    Case 5 : PlaySoundAtLevelStatic (Cartridge_Table_Specifics & "_Spinning_Disc_Ball_Rollover_Bump_5"), 0.95 * SpinningDiscRolloverBumpSoundLevel, RightSpinTrigger
  End Select
End Sub


'//////////////////////////  BACKWALL FAN BLOWER SOUNDS  ////////////////////////
'////////////////////////  BACKWALL FAN BLOWER TRANSIENTS  //////////////////////
Sub SoundBackwallFanBlowerTransients(toggle)
  Select Case toggle
    Case RampUp:
      Select Case BackwallFanBlowerConfiguration
        Case 1:
        Case 2: PlaySoundAtLevelStatic SoundFX(Cartridge_Table_Specifics & "_Fan_Blower_Motor_Start",DOFShaker), 0, BackwallFanPosition
        Case 3: PlaySoundAtLevelStatic (Cartridge_Table_Specifics & "_Fan_Blower_Motor_Start"), Solenoid_BlowerMotor_SoundLevel, BackwallFanPosition
        Case 4: PlaySoundAtLevelStatic SoundFX(Cartridge_Table_Specifics & "_Fan_Blower_Motor_Start",DOFShaker), Solenoid_BlowerMotor_SoundLevel, BackwallFanPosition
      End Select
    Case RampDown:
      Select Case BackwallFanBlowerConfiguration
        Case 1:
        Case 2: PlaySoundAtLevelStatic SoundFX(Cartridge_Table_Specifics & "_Fan_Blower_Motor_Stop",DOFShaker), 0, BackwallFanPosition
        Case 3: PlaySoundAtLevelStatic (Cartridge_Table_Specifics & "_Fan_Blower_Motor_Stop"), Solenoid_BlowerMotor_SoundLevel, BackwallFanPosition
        Case 4: PlaySoundAtLevelStatic SoundFX(Cartridge_Table_Specifics & "_Fan_Blower_Motor_Stop",DOFShaker), Solenoid_BlowerMotor_SoundLevel, BackwallFanPosition
      End Select
  End Select
End Sub


'//////////////////////////  BACKWALL FAN BLOWER LOOP  //////////////////////////
Sub SoundBackwallFanBlowerLoop(toggle)
  Select Case toggle
    Case SoundOn:
      Select Case BackwallFanBlowerConfiguration
        Case 1:
        Case 2: PlaySoundAtLevelStatic SoundFX(Cartridge_Table_Specifics & "_Fan_Blower_Motor_Loop_Delayed",DOFShaker), 0, BackwallFanPosition
        Case 3: PlaySoundAtLevelStatic (Cartridge_Table_Specifics & "_Fan_Blower_Motor_Loop_Delayed"), Solenoid_BlowerMotor_SoundLevel, BackwallFanPosition
        Case 4: PlaySoundAtLevelStatic SoundFX(Cartridge_Table_Specifics & "_Fan_Blower_Motor_Loop_Delayed",DOFShaker), Solenoid_BlowerMotor_SoundLevel, BackwallFanPosition
      End Select
    Case SoundOff:
      Select Case BackwallFanBlowerConfiguration
        Case 1:
        Case 2: Stopsound Cartridge_Table_Specifics & "_Fan_Blower_Motor_Loop_Delayed"
        Case 3: Stopsound Cartridge_Table_Specifics & "_Fan_Blower_Motor_Loop_Delayed"
        Case 4: Stopsound Cartridge_Table_Specifics & "_Fan_Blower_Motor_Loop_Delayed"
      End Select
  End Select
End Sub


'//////////////////////////  SOLENOID A/C SELECT RELAY  /////////////////////////
Sub Sound_Solenoid_AC(toggle)
  Select Case toggle
    Case CircuitA
      If RelaysPosition = 1 Then PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_AC_Select_Relay_Side_A"), RelaySolenoidACSelectSoundLevel, GIUpperPosition
      If RelaysPosition = 2 Then PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_AC_Select_Relay_Side_A"), RelaySolenoidACSelectSoundLevel, ACSelectPosition
    Case CircuitC
      If RelaysPosition = 1 Then PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_AC_Select_Relay_Side_C"), RelaySolenoidACSelectSoundLevel, GIUpperPosition
      If RelaysPosition = 2 Then PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_AC_Select_Relay_Side_C"), RelaySolenoidACSelectSoundLevel, ACSelectPosition
  End Select
End Sub


'//////////////////////////  GENERAL ILLUMINATION RELAYS  ///////////////////////
Sub Sound_LowerGI_Relay(toggle)
  Select Case toggle
    Case SoundOn
      If RelaysPosition = 1 Then PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_Lower_Playfield_and_Backbox_GI_Relay_On"), RelayLowerGISoundLevel, GIUpperPosition
      If RelaysPosition = 2 Then PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_Lower_Playfield_and_Backbox_GI_Relay_On"), RelayLowerGISoundLevel, GIUpperPosition
    Case SoundOff
      If RelaysPosition = 1 Then PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_Lower_Playfield_and_Backbox_GI_Relay_Off"), RelayLowerGISoundLevel, GIUpperPosition
      If RelaysPosition = 2 Then PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_Lower_Playfield_and_Backbox_GI_Relay_Off"), RelayLowerGISoundLevel, GIUpperPosition
  End Select
End Sub

Sub Sound_UpperGI_Relay(toggle)
  Select Case toggle
    Case SoundOn
      If RelaysPosition = 1 Then PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_Upper_Playfield_GI_Relay_On"), RelayUpperGISoundLevel, GIUpperPosition
      If RelaysPosition = 2 Then PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_Upper_Playfield_GI_Relay_On"), RelayUpperGISoundLevel, GILowerPosition
    Case SoundOff
      If RelaysPosition = 1 Then PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_Upper_Playfield_GI_Relay_Off"), RelayUpperGISoundLevel, GIUpperPosition
      If RelaysPosition = 2 Then PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_Upper_Playfield_GI_Relay_Off"), RelayUpperGISoundLevel, GILowerPosition
  End Select
End Sub


'///////////////////////////////  FLASHERS RELAY  ///////////////////////////////
Sub Sound_Flasher_Relay(toggle, tableobj)
  Select Case toggle
    Case SoundOn
      If RelaysPosition = 1 Then PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_Flashers_Relay_On"), RelayFlasherSoundLevel, GIUpperPosition
      If RelaysPosition = 2 Then PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_Flashers_Relay_On"), RelayFlasherSoundLevel, tableobj
    Case SoundOff
      If RelaysPosition = 1 Then PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_Flashers_Relay_Off"), RelayFlasherSoundLevel, GIUpperPosition
      If RelaysPosition = 2 Then PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_Flashers_Relay_Off"), RelayFlasherSoundLevel, tableobj
  End Select
End Sub



'////////////////////  VARIOUS PLAYFIELD SOUND SUBROUTINES  /////////////////////
'/////////////////////////////  RUBBERS AND POSTS  //////////////////////////////
'/////////////////////////////  RUBBERS - EVENTS  ///////////////////////////////
Sub dRubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 10 then
    RandomSoundRubberStrong()
  End if
  If finalspeed <= 10 then
    RandomSoundRubberWeak()
  End If
End Sub


'/////////////////////  RUBBERS AND POSTS - STRONG IMPACTS  /////////////////////
Sub RandomSoundRubberStrong()
  Select Case Int(Rnd*10)+1
    Case 1 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_1"), Vol(ActiveBall) * RubberStrongSoundFactor
    Case 2 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_2"), Vol(ActiveBall) * RubberStrongSoundFactor
    Case 3 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_3"), Vol(ActiveBall) * RubberStrongSoundFactor
    Case 4 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_4"), Vol(ActiveBall) * RubberStrongSoundFactor
    Case 5 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_5"), Vol(ActiveBall) * RubberStrongSoundFactor
    Case 6 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_6"), Vol(ActiveBall) * RubberStrongSoundFactor
    Case 7 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_7"), Vol(ActiveBall) * RubberStrongSoundFactor
    Case 8 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_8"), Vol(ActiveBall) * RubberStrongSoundFactor
    Case 9 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_9"), Vol(ActiveBall) * RubberStrongSoundFactor
    Case 10 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_10"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
  End Select
End Sub


'///////////////////////  RUBBERS AND POSTS - WEAK IMPACTS  /////////////////////
Sub RandomSoundRubberWeak()
  PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_" & Int(Rnd*8)+1), Vol(ActiveBall) * RubberWeakSoundFactor
End Sub


'///////////////////////////////  WALL IMPACTS  /////////////////////////////////
Sub RandomSoundWall()
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    Select Case Int(Rnd*5)+1
      Case 1 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_1"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_2"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_5"), Vol(ActiveBall) * WallImpactSoundFactor * 0.05
      Case 4 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_7"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 5 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_9"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    Select Case Int(Rnd*4)+1
      Case 1 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_3"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 4 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End If
  If finalspeed < 6 Then
    Select Case Int(Rnd*3)+1
      Case 1 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End if
End Sub


'/////////////////////////////  WALL IMPACTS EVENTS  ////////////////////////////
Sub Wall27_Hit()
  RandomSoundMetal()
End Sub

Sub Wall13_Hit()
  RandomSoundMetal()
End Sub

Sub Wall30_Hit()
  RandomSoundMetal()
End Sub

Sub Wall18_Hit()
  RandomSoundMetal()
End Sub

Sub Wall14_Hit()
  RandomSoundMetal()
End Sub

Sub Wall19_Hit()
  RandomSoundMetal()
End Sub

Sub Wall16_Hit()
  RandomSoundMetal()
End Sub

Sub Wall17_Hit()
  RandomSoundMetal()
End Sub

Sub Wall46_Hit()
  PlaySoundAtLevelStatic (Cartridge_Metal_Hits & "_InnerLane_Left_Wall_3"), Vol(ActiveBall) * MetalImpactSoundFactor, LeftInnerLanePosition
End Sub


'////////////////////////////  INNER LEFT LANE WALLS  ///////////////////////////
Sub Wall28_Hit()
  PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_InnerLane_Left_Wall_" & Int(Rnd*3)+1), Vol(ActiveBall) * MetalImpactSoundFactor
End Sub

Sub Wall29_Hit()
  PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_InnerLane_Left_Wall_" & Int(Rnd*3)+1), Vol(ActiveBall) * MetalImpactSoundFactor
End Sub

Sub Wall30_Hit()
  PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_InnerLane_Left_Wall_" & Int(Rnd*3)+1), Vol(ActiveBall) * MetalImpactSoundFactor
End Sub


'/////////////////////////////  METAL TOUCH SOUNDS  /////////////////////////////
Sub RandomSoundMetal()
  Select Case Int(Rnd*20)+1
    Case 1 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_1"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 2 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_2"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 3 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_3"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 4 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_4"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 5 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_5"), Vol(ActiveBall) * 0.02 * MetalImpactSoundFactor
    Case 6 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_6"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 7 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_7"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 8 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_8"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 9 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_9"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 10 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_10"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 11 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_11"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 12 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_12"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 13 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_13"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 14 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_14"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 15 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_15"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 16 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_16"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 17 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_17"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 18 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_18"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 19 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_19"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 20 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_20"), Vol(ActiveBall) * MetalImpactSoundFactor
  End Select
End Sub


'///////////////////////////////  OUTLANES WALLS  ///////////////////////////////
Sub RandomSoundOutlaneWalls()
  PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Outlane_Wall_" & Int(Rnd*9)+1), OutlaneWallsSoundFactor
End Sub


'///////////////////////////  BOTTOM ARCH BALL GUIDE  ///////////////////////////
'///////////////////  BOTTOM ARCH BALL GUIDE - SOFT BOUNCES  ////////////////////
Sub RandomSoundBottomArchBallGuideSoftHit()
  PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Arch_Ball_Guide_Hit_Soft_" & Int(Rnd*4)+1), BottomArchBallGuideSoundFactor
End Sub


'//////////////////////  BOTTOM ARCH BALL GUIDE - HARD HITS  ////////////////////
Sub RandomSoundBottomArchBallGuideHardHit()
  PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Arch_Ball_Guide_Hit_Hard_" & Int(Rnd*3)+1), BottomArchBallGuideSoundFactor * 3
End Sub


'//////////////////////////////  FLIPPER BALL GUIDE  ////////////////////////////
Sub RandomSoundFlipperBallGuide()
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    Select Case Int(Rnd*2)+1
      Case 1 : PlaySoundAtLevelActiveBall (Cartridge_Apron & "_Apron_Hit_Hard_1"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
      Case 2 : PlaySoundAtLevelActiveBall (Cartridge_Apron & "_Apron_Hit_Hard_2"),  Vol(ActiveBall) * 0.8 * FlipperBallGuideSoundFactor
    End Select
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    PlaySoundAtLevelActiveBall (Cartridge_Apron & "_Apron_Hit_Medium_" & Int(Rnd*3)+1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
  End If
  If finalspeed < 6 Then
    PlaySoundAtLevelActiveBall (Cartridge_Apron & "_Apron_Hit_Soft_" & Int(Rnd*7)+1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
  End if
End Sub



'/////////////////////////  WIREFORM ANTI-REBOUNT RAILS  ////////////////////////
Sub RandomSoundWireformAntiRebountRail()
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed >= 10 then
    Select Case Int(Rnd*5)+1
      Case 1 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Wireform_Anti_Rebound_Rail_3"),  Vol(ActiveBall) * WireformAntiRebountRailSoundFactor
      Case 2 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Wireform_Anti_Rebound_Rail_4"),  Vol(ActiveBall) * WireformAntiRebountRailSoundFactor
      Case 3 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Wireform_Anti_Rebound_Rail_5"),  Vol(ActiveBall) * WireformAntiRebountRailSoundFactor
      Case 4 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Wireform_Anti_Rebound_Rail_6"),  Vol(ActiveBall) * WireformAntiRebountRailSoundFactor
      Case 5 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Wireform_Anti_Rebound_Rail_7"),  Vol(ActiveBall) * WireformAntiRebountRailSoundFactor
    End Select
  End if
  If finalspeed < 10 Then
    Select Case Int(Rnd*3)+1
      Case 1 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Wireform_Anti_Rebound_Rail_1"),  Vol(ActiveBall) * WireformAntiRebountRailSoundFactor
      Case 2 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Wireform_Anti_Rebound_Rail_2"),  Vol(ActiveBall) * WireformAntiRebountRailSoundFactor
      Case 3 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Wireform_Anti_Rebound_Rail_8"),  Vol(ActiveBall) * WireformAntiRebountRailSoundFactor
    End Select
  End if
End Sub


'////////////////////////////  LANES AND INNER LOOPS  ///////////////////////////
'////////////////////  INNER LOOPS - LEFT ENTRANCE - EVENTS  ////////////////////
Sub LeftInnerLaneTriggerUp_Hit()
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 10 then
    If ActiveBall.VelY < 0 Then RandomSoundInnerLaneEnter()
  End If
End Sub

Sub LeftInnerLaneTriggerDown_Hit()
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 7 then
    If ActiveBall.VelY > 0 Then RandomSoundInnerLaneEnter()
  End If
End Sub


'/////////////////////  INNER LOOPS - LEFT ENTRANCE - SOUNDS  ///////////////////
Sub RandomSoundInnerLaneEnter()
  PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_InnerLane_Ball_Guide_Hit_" & Int(Rnd*20)+1), Vol(ActiveBall) * InnerLaneSoundFactor
End Sub


'//////////////////////////////  LEFT LANE ENTRANCE  ////////////////////////////
'/////////////////////////  LEFT LANE ENTRANCE - EVENTS  ////////////////////////
Sub LeftLaneTrigger_Hit()
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 10 then
    If ActiveBall.VelY < 0 Then PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Lane_Left_Ball_Enter_Hit"), Vol(ActiveBall) * LaneEnterSoundFactor : RandomSoundLaneLeftEnter()
  End If
End Sub


'/////////////////////////  LEFT LANE ENTRANCE - SOUNDS  ////////////////////////
Sub RandomSoundLaneLeftEnter()
  PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Lane_Left_Ball_Roll_" & Int(Rnd*2)+1), Vol(ActiveBall) * LaneSoundFactor
End Sub


'/////////////////////////////  RIGHT LANE ENTRANCE  ////////////////////////////
'////////////////////////  RIGHT LANE ENTRANCE - EVENTS  ////////////////////////
Sub RightLaneTrigger_Hit()
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 10 then
    If ActiveBall.VelY < 0 Then PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Lane_Right_Ball_Enter_Hit"), Vol(ActiveBall) * LaneEnterSoundFactor : RandomSoundLaneRightEnter()
  End If
End Sub


'/////////////////  RIGHT LANE ENTRANCE (RIGHT ORBIT) - SOUNDS  /////////////////
Sub RandomSoundLaneRightEnter()
  PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Lane_Right_Ball_Roll_" & Int(Rnd*3)+1), Vol(ActiveBall) * LaneSoundFactor
End Sub


'/////////////////////////  RAMP HELPERS BALL VARIABLES  ////////////////////////
'Dim ballvariablePlasticRampTimer1


'/////////  PLASTIC LEFT RAMP - RIGHT EXIT HOLE - TO PLAYFIELD - EVENT  /////////
Sub RHelper1_Hit()
  RandomSoundRightRampRightExitDropToPlayfield()
' Call SoundRightPlasticRampPart2(SoundOff, ballvariablePlasticRampTimer1)
End Sub

'/////////  PLASTIC LEFT RAMP - LEFT EXIT HOLE - TO PLAYFIELD - EVENT  /////////
Sub RHelper2_Hit()
  RandomSoundRightRampLeftExitDropToLock()
' Call SoundRightPlasticRampPart2(SoundOff, ballvariablePlasticRampTimer1)
End Sub

'/////////  METAL WIRE RIGHT RAMP - EXIT HOLE - TO PLAYFIELD - EVENT  //////////
Sub RHelper3_Hit()
  RandomSoundLeftRampDropToPlayfield()
End Sub


'//////////////////  PLASTIC LEFT RAMP - PART 2 SOUND - EVENT  //////////////////
'Sub RHelper4_Hit()
' Call SoundRightPlasticRampPart2(SoundOn, ballvariablePlasticRampTimer1)
'End Sub

'////////////////////////////////////////////////////////////////////////////////
'////          End of Fleep Mechanical Sounds Section                        ////
'////////////////////////////////////////////////////////////////////////////////

'END  of Fleep Audio

' *****************************************
' *** insert lights                   *****
' *****************************************

Dim PFLights(100,5), PFLightsCount(100), PFInsertOnPrim(100), PFInsertOnPrimMult(100), PFInsertMax
Dim PFInsertState(100), PFInsertStateNew(100), PFInsertChange(100), FlBumperScrews(6)
Dim FlBumperTop(6), FlBumperBase(6), FlBumperLightBig(6), FlBumperLightMedium(6), FlBumperLightsmall(6), FlBumperBulb(6)
PFInsertMax= 0


Sub InitLights(aColl, aColl2)
  Dim obj, idx, mult
  For idx = 0 to 100 : PfInsertChange(idx) = False : Next
  For Each obj In aColl
    idx = obj.BlendDisableLightingFromBelow : mult = obj.BlendDisableLighting
    obj.BlendDisableLightingFromBelow = 1
    Set PFInsertOnPrim(idx) = obj : PfInsertOnPrimMult(idx) = mult
    If idx > PfInsertMax then PfInsertMax = idx
    PfInsertState(idx) = 0
  Next
  For Each obj In aColl2
    idx = obj.TimerInterval
    Set PFLights(idx, PFLightsCount(idx)) = obj
    PFLightsCount(idx) = PFLightsCount(idx) + 1
  Next
  For idx = 0 to 5
    Set FlBumperTop(idx) = Eval("bumpertop" & (idx + 33))
    Set FlBumperLightBig(idx) = Eval("bumperlightbig" & (idx + 33))
    Set FlBumperLightMedium(idx) = Eval("bumperlightmedium" & (idx + 33))
    Set FlBumperLightSmall(idx) = Eval("bumperlightsmall" & (idx + 33))
    Set FlBumperBulb(idx) = Eval("bumperbulb" & (idx + 33))
    Set FlBumperBase(idx) = Eval("bumperbase" & (idx + 33))
    Set FlBumperScrews(idx) = Eval("bumperscrews" & (idx + 33))
  Next
End Sub

dim BumperCurrent(5), BumperNew(5)
BumperCurrent(0) = 0 : BumperCurrent(1) = 0 : BumperCurrent(2) = 0 : BumperCurrent(3) = 0 : BumperCurrent(4) = 0 : BumperCurrent(5) = 0
dim CellarLeftCurrent, CellarRightCurrent, CellarLeftNew, CellarRightNew
CellarLeftCurrent = 0 : CellarRightCurrent = 0

Sub LampTimer_Timer()
  Dim chgLamp, ii, nr
  chgLamp = Controller.ChangedLamps
  If Not IsEmpty(chgLamp) Then
    For ii = 0 To UBound(chgLamp)
      For nr = 1 to PFLightsCount(chglamp(ii,0)) : PFLights(chglamp(ii,0),nr - 1).state = chglamp(ii,1) : Next
      Select Case chglamp(ii,0)
        case 2 : EMReel001.SetValue(chglamp(ii,1))
        case 3 : EMReel002.SetValue(chglamp(ii,1))
        case 4 : EMReel003.SetValue(chglamp(ii,1))
        case 5 : EMReel004.SetValue(chglamp(ii,1))
        case 6 : EMReel005.SetValue(chglamp(ii,1))
        case 7 : EMReel006.SetValue(chglamp(ii,1))
        case 8 : EMReel007.SetValue(chglamp(ii,1))
      End Select
      PfInsertStateNew(chglamp(ii,0)) = chglamp(ii,1)
      If Not IsEmpty(PFInsertOnPrim(chglamp(ii,0))) Then
        PfInsertChange(chglamp(ii,0)) = True
      Else
        If chglamp(ii,0) > 32 and chglamp(ii,0) < 39 Then
          If Not IsEmpty(FlBumperTop(chglamp(ii,0) - 33)) Then PfInsertChange(chglamp(ii,0)) = True
        Else
          If chglamp(ii,0) = 39 or chglamp(ii,0) = 40 Then PfInsertChange(chglamp(ii,0)) = True : end if
        End If
      End If
    Next
  End If
  For ii = 0 To PfInsertMax
    If PfInsertChange(ii) Then
      Select case ii
        case 33, 34, 35
          BumperNew(ii - 33) = PfInsertStateNew(ii) * 3
        case 36, 37, 38
          BumperNew(ii - 33) = PfInsertStateNew(ii) * 3
        case 39
          CellarLeftNew = PfInsertStateNew(ii) * 7
        case 40
          CellarRightNew = PfInsertStateNew(ii) * 7
        Case else
          If PFInsertState(ii) < PFInsertStateNew(ii) Then
            PFInsertState(ii) = PFInsertState(ii)  + 2 * ((PFInsertStateNew(ii) - PFInsertState(ii)) / 3)^2
          Else
            PFInsertState(ii) = PFInsertState(ii)  + (PFInsertStateNew(ii) - PFInsertState(ii)) / 1.5
          End If
          If abs(PFInsertState(ii) - PFInsertStateNew(ii)) < 0.01 Then PFInsertState(ii) = PFInsertStatenew(ii) : PfInsertChange(ii) = False
          PFInsertOnPrim(ii).BlendDisableLighting =  10000 * PFInsertOnPrimMult(ii) * PFInsertState(ii)
      End Select
    End If
  Next

End Sub

Sub FlBumperRedFade(nr, fade)
  dim faderange : faderange = fade / 3
  FlBumperTop(nr).image = "bumpertexredfade" & fade : FlBumperTop(nr).BlendDisableLighting = 0.4 + 0.6 * faderange
  FlBumperLightBig(nr).intensity = faderange * 5
  FlBumperLightMedium(nr).opacity = 5000 + faderange * 1000 : FlBumperLightMedium(nr).ModulateVsAdd = 1 - 0.05 * faderange
  FlBumperBulb(nr).image = "bumperbulbredfade" & fade : FlBumperBulb(nr).BlendDisableLighting = 0.2 + 31.5 * faderange '1.5 instead of 0.2
  FlBumperBase(nr).image = "bumperbaseredfade" & fade: FlBumperBase(nr).BlendDisableLighting = 0.6 + 0.5 * GInrUpper + 0.4 * faderange
  FlBumperBase(nr).material = "bumperbasefade" & fade
  FlBumperLightSmall(nr).intensity = 80 - 30 * faderange : FlBumperLightSmall(nr).color = RGB(255, 32 * faderange,0)
  FlBumperLightSmall(nr).colorfull = RGB (255,16 * faderange,0)
  FlBumperScrews(nr).material = "bumperred" & fade : FlBumperScrews(nr).BlendDisableLighting = 1 + faderange * 5
  Select case nr
    case 3 : coll032.material = "bumperrampfade" & fade
    case 4 : coll034.material = "bumperrampfade" & fade
  End Select
End Sub

Sub FlBumperRedFadeVR(nr, fade)
  dim faderange : faderange = fade / 3
  FlBumperTop(nr).image = "bumpertexredfadeVR" & fade : FlBumperTop(nr).BlendDisableLighting = 0.4 + 1.6 * faderange
  FlBumperLightBig(nr).intensity = faderange * 5
  FlBumperLightMedium(nr).opacity = 3000 + faderange * 6000 : FlBumperLightMedium(nr).ModulateVsAdd = 1 - 0.1 * faderange
  FlBumperBulb(nr).image = "bumperbulbredfade" & fade : FlBumperBulb(nr).BlendDisableLighting = 0.2 + 31.5 * faderange
  FlBumperBase(nr).image = "bumperbaseredfade" & fade: FlBumperBase(nr).BlendDisableLighting = 0.6 + 0.5 * GInrUpper + 0.4 * faderange
  FlBumperBase(nr).material = "bumperbasefade" & fade
  FlBumperLightSmall(nr).intensity = 200 - 125 * faderange : FlBumperLightSmall(nr).color = RGB(255, 32 * faderange,0)
  FlBumperLightSmall(nr).colorfull = RGB (255,16 * faderange,0)
  FlBumperScrews(nr).material = "bumperred" & fade : FlBumperScrews(nr).BlendDisableLighting = 1 + faderange * 5
  Select case nr
    case 3 : coll032.material = "bumperrampfade" & fade
    case 4 : coll034.material = "bumperrampfade" & fade
  End Select
End Sub

Sub FlBumperYellowFade(nr, fade)
  dim faderange : faderange = fade / 3
  FlBumperTop(nr).image = "bumpertexyellowfade" & fade : FlBumperTop(nr).BlendDisableLighting = 0.4 + 0.2 * GInrLower + 1.4 * faderange
  FlBumperLightBig(nr).intensity = faderange * 5
  FlBumperLightMedium(nr).opacity = 10000 + faderange * 2000 : FlBumperLightMedium(nr).ModulateVsAdd = 1 - 0.05 * faderange
  FlBumperBulb(nr).image = "bumperbulbfade" & fade : FlBumperBulb(nr).BlendDisableLighting = 0.2 + 0.5 * GInrLower + 14.3 * faderange
  FlBumperBase(nr).image = "bumperbaseyellowfade" & fade: FlBumperBase(nr).BlendDisableLighting = 0.5 + 0.5 * GInrLower + 0.5 * faderange
  FlBumperLightSmall(nr).intensity = 40 + 60 * faderange : FlBumperLightSmall(nr).color = RGB(255,190 + 15 * faderange,32 * faderange)
  FlBumperLightSmall(nr).colorfull = RGB (255,190,32 * faderange)
  FlBumperScrews(nr).material = "bumperyellow" & fade : FlBumperScrews(nr).BlendDisableLighting = 1 + faderange * 5
  Select case nr
    case 0 : coll031.material = "bumperrampfade" & fade
    case 1 : coll030.material = "bumperrampfade" & fade
  End Select
End Sub

Sub FlBumperYellowFadeVR(nr, fade)
  dim faderange : faderange = fade / 3
  FlBumperTop(nr).image = "bumpertexyellowfadeVR" & fade : FlBumperTop(nr).BlendDisableLighting = 0.4 + 0.2 * GInrLower + 1.4 * faderange
  FlBumperLightBig(nr).intensity = faderange * 5
  FlBumperLightMedium(nr).opacity = 10000 + faderange * 2000 : FlBumperLightMedium(nr).ModulateVsAdd = 1 - 0.05 * faderange
  FlBumperBulb(nr).image = "bumperbulbfade" & fade : FlBumperBulb(nr).BlendDisableLighting = 0.2 + 0.5 * GInrLower + 20 * faderange
  FlBumperBase(nr).image = "bumperbaseyellowfade" & fade: FlBumperBase(nr).BlendDisableLighting = 0.5 + 0.5 * GInrLower + 0.5 * faderange
  FlBumperLightSmall(nr).intensity = 40 + 90 * faderange : FlBumperLightSmall(nr).color = RGB(255,190 + 15 * faderange,32 * faderange)
  FlBumperLightSmall(nr).colorfull = RGB (255,190,32 * faderange)
  FlBumperScrews(nr).material = "bumperyellow" & fade : FlBumperScrews(nr).BlendDisableLighting = 1 + faderange * 5
  Select case nr
    case 0 : coll031.material = "bumperrampfade" & fade
    case 1 : coll030.material = "bumperrampfade" & fade
  End Select
End Sub


Dim looper : looper = -1


Dim DFlasherIntensity : DFlasherIntensity = FlasherIntensity^2
Dim Mult100, Mult120, Mult15, Mult8, Mult5, Mult10000, Mult10, Mult200, Mult2000, Mult140, Mult220, Mult100000, Mult20
Mult100 = 100 * DFlasherIntensity : Mult120 = 120 * DFlasherIntensity : Mult15 = 15 * DFlasherIntensity : Mult8 = 8 * DFlasherIntensity
Mult10000 = 10000 * DFlasherIntensity : Mult10 = 10 * DFlasherIntensity : Mult200 = 200 * DFlasherIntensity : Mult2000 = 2000 * DFlasherIntensity
Mult140 = 140 * DFlasherIntensity : Mult220 = 220 * DFlasherIntensity : Mult100000 = 100000 * DFlasherIntensity : Mult20 = 20 * DFlasherIntensity



Sub SyncRefreshTimer_Timer()

  If sw41.currentangle < 180 Then
    Light004.intensity = 500 * ((180 - sw41.CurrentAngle mod 180)/180)^8
  Else
    Light004.intensity = 500 * ((sw41.CurrentAngle mod 181)/180)^8
  End If

  LeftBat.ObjRotZ = LeftFlipper.CurrentAngle : batleftshadow.ObjRotZ = LeftFlipper.CurrentAngle
  RightBat.ObjRotZ = RightFlipper.CurrentAngle : batrightshadow.ObjRotZ = RightFlipper.CurrentAngle
  URightBat.ObjRotZ = RightFlipper1.CurrentAngle : baturightshadow.ObjRotZ = RightFlipper1.CurrentAngle

  dim ii
  For ii = 0 to 2
    If BumperCurrent(ii) < BumperNew(ii) Then
      BumperCurrent(ii) = BumperCurrent(ii) + 1
      If VRRoom > 0 Then
        FlBumperYellowFadeVR ii, BumperCurrent(ii)
      Else
        FlBumperYellowFade ii, BumperCurrent(ii)
      End If
    End If
    If BumperCurrent(ii) > BumperNew(ii) Then
      BumperCurrent(ii) = BumperCurrent(ii) - 1
      If VRRoom > 0 Then
        FlBumperYellowFadeVR ii, BumperCurrent(ii)
      Else
        FlBumperYellowFade ii, BumperCurrent(ii)
      End If
    End If
  Next
  For ii = 3 to 5
    If BumperCurrent(ii) < BumperNew(ii) Then
      BumperCurrent(ii) = BumperCurrent(ii) + 1
      If VRRoom > 0 Then
        FlBumperredFadeVR ii, BumperCurrent(ii)
      Else
        FlBumperredFade ii, BumperCurrent(ii)
      End If
    End If
    If BumperCurrent(ii) > BumperNew(ii) Then
      BumperCurrent(ii) = BumperCurrent(ii) - 1
      If VRRoom > 0 Then
        FlBumperredFadeVR ii, BumperCurrent(ii)
      Else
        FlBumperredFade ii, BumperCurrent(ii)
      End If
    End If
  Next

  If CellarLeftCurrent < CellarLeftNew Then
    CellarLeftCurrent = CellarLeftCurrent + 1
    BulbLeft.image = "simplelightwhite" & CellarLeftCurrent
    BulbLeft.blenddisablelighting = CellarLeftCurrent / 3
    LightCellarLeft.intensity = (CellarLeftCurrent / 7) ^ 3 * 30
  End If
  If CellarLeftCurrent > CellarLeftNew Then
    CellarLeftCurrent = CellarLeftCurrent - 1
    BulbLeft.image = "simplelightwhite" & CellarLeftCurrent
    BulbLeft.blenddisablelighting = CellarLeftCurrent / 3
    LightCellarLeft.intensity = (CellarLeftCurrent / 7) ^ 3 * 30
  End If

  If CellarRightCurrent < CellarRightNew Then
    CellarRightCurrent = CellarRightCurrent + 1
    BulbRight.image = "simplelightwhite" & CellarRightCurrent
    BulbRight.blenddisablelighting = CellarRightCurrent / 3
    LightCellarRight.intensity = (CellarRightCurrent / 7) ^ 3 * 30
  End If
  If CellarRightCurrent > CellarRightNew Then
    CellarRightCurrent = CellarRightCurrent - 1
    BulbRight.image = "simplelightwhite" & CellarRightCurrent
    BulbRight.blenddisablelighting = CellarRightCurrent / 4
    LightCellarRight.intensity = (CellarRightCurrent / 7) ^ 3 * 30
  End If

  Dim obj
  looper = looper + 1
  If looper > 3 Then looper = 0 : End If
  Select Case looper
    case 0
      If Flasher6c_enabled Then
        For Each obj In Flasher1coll : obj.image = "texflasher6c" : obj.blenddisablelighting = Mult120 : obj.visible = true : Next
        playfield_renderflasher1.image = "texflasher6cpf" : playfield_renderflasher1.blenddisablelighting = Mult200 : playfield_renderflasher1.visible = true
        blackwoodflasher1.image = "texflasher6cpf" : blackwoodflasher1.blenddisablelighting = Mult100 : blackwoodflasher1.visible = true
        aproncardsflasher1.image = "texflasher6cpf" : aproncardsflasher1.blenddisablelighting = Mult100 : aproncardsflasher1.visible = true
        sidewallsflasher1.image = "texflasher6cpf" : sidewallsflasher1.blenddisablelighting = Mult15 : sidewallsflasher1.visible = true
        Light002.state = 1
      Else
        DisableFlasher1 : Light002.state = 0
      End If
      If Flasher1c_enabled Then
        For Each obj In Flasher2coll : obj.image = "texflasher1c" : obj.blenddisablelighting = Mult8 : obj.visible = true : Next
        playfield_renderflasher2.image = "texflasher1cpf" : playfield_renderflasher2.blenddisablelighting = Mult100 : playfield_renderflasher2.visible = true
        blackwoodflasher2.image = "texflasher1cpf" : blackwoodflasher2.blenddisablelighting = Mult100 : blackwoodflasher2.visible = true
        aproncardsflasher2.image = "texflasher1cpf" : aproncardsflasher2.blenddisablelighting = Mult100 : aproncardsflasher2.visible = true
        sidewallsflasher2.image = "texflasher1cpf" : sidewallsflasher2.blenddisablelighting = Mult5 : sidewallsflasher2.visible = true
      Else
        DisableFlasher2
      End If
    case 1
      If Flasher5c_enabled Then
        For Each obj In Flasher1coll : obj.image = "texflasher5c" : obj.blenddisablelighting = Mult120 : obj.visible = true : Next
        playfield_renderflasher1.image = "texflasher5cpf" : playfield_renderflasher1.blenddisablelighting = Mult200 : playfield_renderflasher1.visible = true
        blackwoodflasher1.image = "texflasher5cpf" : blackwoodflasher1.blenddisablelighting = Mult100 : blackwoodflasher1.visible = true
        aproncardsflasher1.image = "texflasher5cpf" : aproncardsflasher1.blenddisablelighting = Mult100 : aproncardsflasher1.visible = true
        sidewallsflasher1.image = "texflasher5cpf" : sidewallsflasher1.blenddisablelighting = Mult5 : sidewallsflasher1.visible = true
      Else
        DisableFlasher1
      End If
      If Flasher23_enabled Then
        For Each obj In Flasher2coll : obj.image = "texflasher23" : obj.blenddisablelighting = Mult8 : obj.visible = true : Next
        playfield_renderflasher2.image = "texflasher23pf" : playfield_renderflasher2.blenddisablelighting = Mult2000 : playfield_renderflasher2.visible = true
        blackwoodflasher2.image = "texflasher23pf" : blackwoodflasher2.blenddisablelighting = Mult100 : blackwoodflasher2.visible = true
        aproncardsflasher2.image = "texflasher23pf" : aproncardsflasher2.blenddisablelighting = Mult100 : aproncardsflasher2.visible = true
        sidewallsflasher2.image = "texflasher23pf" : sidewallsflasher2.blenddisablelighting = Mult10 : sidewallsflasher2.visible = true
        Flasherflash002.visible = true
      Else
        'DisableFlasher2
        Flasherflash002.visible = false
      End If
    case 2
      If Flasher4c_enabled Then
        For Each obj In Flasher1coll : obj.image = "texflasher4c" : obj.blenddisablelighting = Mult120 : obj.visible = true : Next
        playfield_renderflasher1.image = "texflasher4cpf" : playfield_renderflasher1.blenddisablelighting = Mult140 : playfield_renderflasher1.visible = true
        blackwoodflasher1.image = "texflasher4cpf" : blackwoodflasher1.blenddisablelighting = Mult100 : blackwoodflasher1.visible = true
        aproncardsflasher1.image = "texflasher4cpf" : aproncardsflasher1.blenddisablelighting = Mult100 : aproncardsflasher1.visible = true
        sidewallsflasher1.image = "texflasher4cpf" : sidewallsflasher1.blenddisablelighting = Mult10 : sidewallsflasher1.visible = true
      Else
        DisableFlasher1
      End If
          If Flasher25_enabled Then
        For Each obj In Flasher2coll : obj.image = "texflasher25" : obj.blenddisablelighting = Mult15 : obj.visible = true : Next
        playfield_renderflasher2.image = "texflasher25pf" : playfield_renderflasher2.blenddisablelighting = Mult10000 : playfield_renderflasher2.visible = true
        blackwoodflasher2.image = "texflasher25pf" : blackwoodflasher2.blenddisablelighting = Mult100 : blackwoodflasher2.visible = true
        aproncardsflasher2.image = "texflasher25pf" : aproncardsflasher2.blenddisablelighting = Mult100 : aproncardsflasher2.visible = true
        sidewallsflasher2.image = "texflasher25pf" : sidewallsflasher2.blenddisablelighting = Mult15 : sidewallsflasher2.visible = true
        Flasherflash001.visible = true
      Else
        DisableFlasher2
        Flasherflash001.visible = false
      End If
    case 3
      If Flasher3c_enabled Then
        For Each obj In Flasher1coll : obj.image = "texflasher3c" : obj.blenddisablelighting = Mult220 : obj.visible = true : Next
        playfield_renderflasher1.image = "texflasher3cpf" : playfield_renderflasher1.blenddisablelighting = Mult200 : playfield_renderflasher1.visible = true
        blackwoodflasher1.image = "texflasher3cpf" : blackwoodflasher1.blenddisablelighting = Mult100 : blackwoodflasher1.visible = true
        aproncardsflasher1.image = "texflasher3cpf" : aproncardsflasher1.blenddisablelighting = Mult100 : aproncardsflasher1.visible = true
        sidewallsflasher1.image = "texflasher3cpf" : sidewallsflasher1.blenddisablelighting = Mult5 : sidewallsflasher1.visible = true
      Else
        DisableFlasher1
      End If
      If Flasher26_enabled Then
        For Each obj In Flasher2coll : obj.image = "texflasher26" : obj.blenddisablelighting = Mult15 : obj.visible = true : Next
        playfield_renderflasher2.image = "texflasher26pf" : playfield_renderflasher2.blenddisablelighting = Mult100000 : playfield_renderflasher2.visible = true
        blackwoodflasher2.image = "texflasher26pf" : blackwoodflasher2.blenddisablelighting = Mult100 : blackwoodflasher2.visible = true
        aproncardsflasher2.image = "texflasher26pf" : aproncardsflasher2.blenddisablelighting = Mult100 : aproncardsflasher2.visible = true
        sidewallsflasher2.image = "texflasher26pf" : sidewallsflasher2.blenddisablelighting = Mult15 : sidewallsflasher2.visible = true
        Flasherflash003.visible = true
      Else
        DisableFlasher2
        Flasherflash003.visible = false
      End If

  End Select
End Sub

Sub DisableFlasher1()
  Dim obj
  If playfield_renderflasher1.visible Then
    For Each obj In Flasher1coll : obj.visible = false : Next
    playfield_renderflasher1.visible = false : blackwoodflasher1.visible = false : aproncardsflasher1.visible = false : sidewallsflasher1.visible = false
  End If
End Sub

Sub DisableFlasher2()
  Dim obj
  If playfield_renderflasher2.visible Then
    For Each obj In Flasher2coll : obj.visible = false : Next
    playfield_renderflasher2.visible = false : blackwoodflasher2.visible = false : aproncardsflasher2.visible = false : sidewallsflasher2.visible = false
  End If
End Sub


Dim TestFlashers, TableRef, FlasherLightIntensity, FlasherFlareIntensity, FlasherOffBrightness

                          ' *********************************************************************
TestFlashers = 0                  ' *** set this to 1 to check position of flasher object       ***
Set TableRef = Table1                 ' *** change this, if your table has another name             ***
FlasherLightIntensity = 1 * FlasherIntensityDome  ' *** lower this, if the VPX lights are too bright (i.e. 0.1)   ***
FlasherFlareIntensity = 3 * FlasherIntensityDome  ' *** lower this, if the flares are too bright (i.e. 0.1)     ***
FlasherOffBrightness = 1              ' *** brightness of the flasher dome when switched off (range 0-2)  ***
                          ' *********************************************************************

Dim ObjLitIntensity, ObjFadeIntensity
ObjLitIntensity = 300 * FlasherIntensityDome^0.75
ObjFadeIntensity = 100 * FlasherIntensityDome^0.75

Dim Dmult20, Dmult10, Dmult40
Dmult20 = 20 * FlasherIntensityDome^2 : Dmult10 = 10 * FlasherIntensityDome^2 : Dmult40 = 40 * FlasherIntensityDome

Dim ObjLevel(20), objbase(20), objlit(20), objflasher(20), objlight(20), objfade(20)
Dim tablewidthFl, tableheightFl : tablewidthFl = TableRef.width : tableheightFl = TableRef.height
'initialise the flasher color, you can only choose from "green", "red", "purple", "blue", "white" and "yellow"
InitFlasher 1, "red" : InitFlasher 2, "red" : InitFlasher 3, "red" : InitFlasher 4, "red"
InitFlasher 9, "red"
RotateFlasher 9,2

Sub InitFlasher(nr, col)
  ' store all objects in an array for use in FlashFlasher subroutine
  Set objbase(nr) = Eval("Flasherbase" & nr) : Set objlit(nr) = Eval("Flasherlit" & nr)
  Set objflasher(nr) = Eval("Flasherflash" & nr) : Set objlight(nr) = Eval("Flasherlight" & nr)
  Set objfade(nr) = Eval("Flasherfade" & nr)
  ' If the flasher is parallel to the playfield, rotate the VPX flasher object for POV and place it at the correct height
  If objbase(nr).RotY = 0 Then
    objbase(nr).ObjRotZ =  atn( (tablewidthFl/2 - objbase(nr).x) / (objbase(nr).y - tableheightFl*1.1)) * 180 / 3.14159
    objflasher(nr).RotZ = objbase(nr).ObjRotZ : objflasher(nr).height = objbase(nr).z + 60
  End If
  ' set all effects to invisible and move the lit primitive at the same position and rotation as the base primitive
  objlight(nr).IntensityScale = 0 : objlit(nr).visible = 0 : objfade(nr).visible = 0 : objlit(nr).material = "Flashermaterial" & nr
  objlit(nr).RotX = objbase(nr).RotX : objlit(nr).RotY = objbase(nr).RotY : objlit(nr).RotZ = objbase(nr).RotZ
  objlit(nr).ObjRotX = objbase(nr).ObjRotX : objlit(nr).ObjRotY = objbase(nr).ObjRotY : objlit(nr).ObjRotZ = objbase(nr).ObjRotZ
  objlit(nr).x = objbase(nr).x : objlit(nr).y = objbase(nr).y : objlit(nr).z = objbase(nr).z
  objfade(nr).RotX = objbase(nr).RotX : objfade(nr).RotY = objbase(nr).RotY : objfade(nr).RotZ = objbase(nr).RotZ
  objfade(nr).ObjRotX = objbase(nr).ObjRotX : objfade(nr).ObjRotY = objbase(nr).ObjRotY : objfade(nr).ObjRotZ = objbase(nr).ObjRotZ
  objfade(nr).x = objbase(nr).x : objfade(nr).y = objbase(nr).y : objfade(nr).z = objbase(nr).z
  objbase(nr).BlendDisableLighting = FlasherOffBrightness
  ' set the texture and color of all objects
  select case objbase(nr).image
    Case "dome2basewhite" : objbase(nr).image = "dome2base" & col : objlit(nr).image = "dome2lit" & col :
    Case "ronddomebasewhite" : objbase(nr).image = "ronddomebase" & col : objlit(nr).image = "ronddomelit" & col
    Case "domeearbasewhite" : objbase(nr).image = "domeearbase" & col : objlit(nr).image = "domeearlit" & col
  end select
  If TestFlashers = 0 Then objflasher(nr).imageA = "domeflashwhite" : objflasher(nr).visible = 0 : End If
  select case col
    Case "blue" :   objlight(nr).color = RGB(4,120,255) : objflasher(nr).color = RGB(200,255,255) : objlight(nr).intensity = 5000
    Case "green" :  objlight(nr).color = RGB(12,255,4) : objflasher(nr).color = RGB(12,255,4)
    Case "red" :    objlight(nr).color = RGB(255,32,4) : objflasher(nr).color = RGB(255,32,4)
    Case "purple" : objlight(nr).color = RGB(230,49,255) : objflasher(nr).color = RGB(255,64,255)
    Case "yellow" : objlight(nr).color = RGB(200,173,25) : objflasher(nr).color = RGB(255,200,50)
    Case "white" :  objlight(nr).color = RGB(255,240,150) : objflasher(nr).color = RGB(100,86,59)
  end select
  objlight(nr).colorfull = objlight(nr).color
  If TableRef.ShowDT and ObjFlasher(nr).RotX = -45 Then
    objflasher(nr).height = objflasher(nr).height - 20 * ObjFlasher(nr).y / tableheight
    ObjFlasher(nr).y = ObjFlasher(nr).y + 10
  End If
  ObjLevel(nr) = 1
  FlashFlasher(nr)
End Sub

Sub RotateFlasher(nr, angle) : angle = ((angle + 360 - objbase(nr).ObjRotZ) mod 180)/30 : objbase(nr).showframe(angle) : objlit(nr).showframe(angle) : objfade(nr).showframe(angle) : End Sub

Sub FlashFlasher(nr)
  Dim tmpvar
  If not objflasher(nr).TimerEnabled Then
    objflasher(nr).TimerEnabled = True : objflasher(nr).visible = 1 : objlit(nr).visible = 1 : objfade(nr).visible = 1
    Select case nr
      case 1 : rampflash6c.visible = true
      case 2 : rampflash5c.visible = true
      case 3 : rampflash4c.visible = true
      case 4 : rampflash3c.visible = true
      case 9 : coll005flasher001.visible = true
    End Select
  End If
  objflasher(nr).opacity = 1000 *  FlasherFlareIntensity * ObjLevel(nr)^2.5
  objlight(nr).IntensityScale = 0.5 * FlasherLightIntensity * ObjLevel(nr)^3
  Select case nr
    case 1 : rampflash6c.blenddisablelighting = Dmult10 * ObjLevel(nr)^3 : If GIStateLower Then tmpvar = 0.1 else tmpvar = 1 : end if
    case 2 : rampflash5c.blenddisablelighting = Dmult20 * ObjLevel(nr)^2 : If GIStateLower Then tmpvar = 0.1 else tmpvar = 1 : end if
    case 3 : rampflash4c.blenddisablelighting = Dmult10 * ObjLevel(nr)^3 : If GIStateUpper Then tmpvar = 0 else tmpvar = 1 : end if
    case 4 : rampflash3c.blenddisablelighting = Dmult20 * ObjLevel(nr)^3 : If GIStateUpper Then tmpvar = 0 else tmpvar = 1 : end if
    case 9 : If GIStateLower Then tmpvar = 0.1 else tmpvar = 1 : end if : coll005flasher001.blenddisablelighting = Dmult20 * ObjLevel(nr)^3
  End Select
  objbase(nr).BlendDisableLighting =  tmpvar + Dmult40 * ObjLevel(nr)^3
  objlit(nr).BlendDisableLighting = ObjLitIntensity * ObjLevel(nr)^2
  objfade(nr).BlendDisableLighting = ObjFadeIntensity * ObjLevel(nr)^0.2
  UpdateMaterial "Flashermaterial" & nr,0,0,0,0,0,0,ObjLevel(nr),RGB(255,255,255),0,0,False,True,0,0,0,0
  ObjLevel(nr) = ObjLevel(nr) * 0.9 - 0.005
  If ObjLevel(nr) < 0 Then
    objflasher(nr).TimerEnabled = False : objflasher(nr).visible = 0 : objlit(nr).visible = 0 : objfade(nr).visible = 0
    Select Case nr
      case 1 : rampflash6c.visible = false
      case 2 : rampflash5c.visible = false
      case 3 : rampflash4c.visible = false
      case 4 : rampflash3c.visible = false
      case 9 : coll005flasher001.visible = false
    End Select
  End If
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



'VRADDED *********************************************************************************************************************

'************* VR Clock ****************
Dim CurrentMinute

Sub ClockTimer_Timer()
  Pminutes.RotAndTra2 = (Minute(Now())+(Second(Now())/100))*6
  Phours.RotAndTra2 = Hour(Now())*30+(Minute(Now())/2)
  Pseconds.RotAndTra2 = (Second(Now()))*6
  CurrentMinute=Minute(Now())
End Sub

'***************************************


'******************************************************
'*******  Set Up Backglass  *******
'******************************************************

  Dim xoff,yoff,zoff,xrot,zscale, xcen,ycen

Sub SetBackglass()
  Dim obj
  xoff = -20
  yoff = 61
  zoff = 699
  xrot = -90
  center_digits()

  For Each obj In BackglassLow
    obj.x = obj.x
    obj.height = - obj.y + 325
    obj.y = 90 'adjusts the distance from the backglass towards the user
    obj.rotx=-86
  Next


  For Each obj In BackglassMid
    obj.x = obj.x
    obj.height = - obj.y + 325
    obj.y = 15 'adjusts the distance from the backglass towards the user
    obj.rotx=-86
  Next


  For Each obj In BackglassHigh
    obj.x = obj.x
    obj.height = - obj.y + 325
    obj.y = 0 'adjusts the distance from the backglass towards the user
    obj.rotx=-86
  Next
End Sub

Set LampCallback = GetRef("UpdateMultipleLamps")

Sub UpdateMultipleLamps()
  IF VRRoom <> 0 Then
  If Controller.Lamp(2) = 0 Then: BGFL2_UpperJets.visible=0: else: BGFL2_UpperJets.visible=1
  If Controller.Lamp(3) = 0 Then: BGFL6_500k.visible=0: else: BGFL6_500k.visible=1
  If Controller.Lamp(4) = 0 Then: BGFL4_XBall.visible=0: else: BGFL4_XBall.visible=1
  If Controller.Lamp(5) = 0 Then: BGFL5_3Bank100k.visible=0: else: BGFL5_3Bank100k.visible=1
  If Controller.Lamp(6) = 0 Then: BGFL3_MultiBall.visible=0: else: BGFL3_MultiBall.visible=1   '????? Manual says 250k light, backglass says multiball
  If Controller.Lamp(7) = 0 Then: BGFL7_Million.visible=0: else: BGFL7_Million.visible=1
  If Controller.Lamp(8) = 0 Then: BGFL8_LowerJets.visible=0: else: BGFL8_LowerJets.visible=1
  End If
End Sub



' ***************************************************************************
'Alphanumeric Setup Code
' ****************************************************************************

Sub center_digits()
  Dim ix, xx, yy, yfact, xfact, xobj

  zscale = 0.0000002

  xcen = 0  '(130 /2) - (92 / 2)
  ycen = (780 /2 ) + (203/2)'(780 /2 ) + (203 /2)

  for ix = 0 to 31
    For Each xobj In DigitsVR(ix)

      xx = xobj.x

      xobj.x = (xoff - xcen) + xx
      yy = xobj.y ' get the yoffset before it is changed
      xobj.y = yoff

      If (yy < 0.) then
        yy = yy * -1
      end if

      xobj.height = (zoff - ycen) + yy - (yy * (zscale))
      xobj.rotx = xrot

    Next
  Next

'moving the bottom digits forward just a bit...
For Each xobj In DigitsVR2
xobj.y = xobj.y + 8
next

end sub


Dim DigitsVR(32)
DigitsVR(0)=Array(ax00, ax05, ax0c, ax0d, ax08, ax01, ax06, ax0f, ax02, ax03, ax04, ax07, ax0b, ax0a, ax09, ax0e)
DigitsVR(1)=Array(ax10, ax15, ax1c, ax1d, ax18, ax11, ax16, ax1f, ax12, ax13, ax14, ax17, ax1b, ax1a, ax19, ax1e)
DigitsVR(2)=Array(ax20, ax25, ax2c, ax2d, ax28, ax21, ax26, ax2f, ax22, ax23, ax24, ax27, ax2b, ax2a, ax29, ax2e)
DigitsVR(3)=Array(ax30, ax35, ax3c, ax3d, ax38, ax31, ax36, ax3f, ax32, ax33, ax34, ax37, ax3b, ax3a, ax39, ax3e)
DigitsVR(4)=Array(ax40, ax45, ax4c, ax4d, ax48, ax41, ax46, ax4f, ax42, ax43, ax44, ax47, ax4b, ax4a, ax49, ax4e)
DigitsVR(5)=Array(ax50, ax55, ax5c, ax5d, ax58, ax51, ax56, ax5f, ax52, ax53, ax54, ax57, ax5b, ax5a, ax59, ax5e)
DigitsVR(6)=Array(ax60, ax65, ax6c, ax6d, ax68, ax61, ax66, ax6f, ax62, ax63, ax64, ax67, ax6b, ax6a, ax69, ax6e)
DigitsVR(7)=Array(ax70, ax75, ax7c, ax7d, ax78, ax71, ax76, ax7f, ax72, ax73, ax74, ax77, ax7b, ax7a, ax79, ax7e)
DigitsVR(8)=Array(ax80, ax85, ax8c, ax8d, ax88, ax81, ax86, ax8f, ax82, ax83, ax84, ax87, ax8b, ax8a, ax89, ax8e)
DigitsVR(9)=Array(ax90, ax95, ax9c, ax9d, ax98, ax91, ax96, ax9f, ax92, ax93, ax94, ax97, ax9b, ax9a, ax99, ax9e)
DigitsVR(10)=Array(axa0, axa5, axac, axad, axa8, axa1, axa6, axaf, axa2, axa3, axa4, axa7, axab, axaa, axa9, axae)
DigitsVR(11)=Array(axb0, axb5, axbc, axbd, axb8, axb1, axb6, axbf, axb2, axb3, axb4, axb7, axbb, axba, axb9, axbe)
DigitsVR(12)=Array(axc0, axc5, axcc, axcd, axc8, axc1, axc6, axcf, axc2, axc3, axc4, axc7, axcb, axca, axc9, axce)
DigitsVR(13)=Array(axd0, axd5, axdc, axdd, axd8, axd1, axd6, axdf, axd2, axd3, axd4, axd7, axdb, axda, axd9, axde)
DigitsVR(14)=Array(axe0, axe5, axec, axed, axe8, axe1, axe6, axef, axe2, axe3, axe4, axe7, axeb, axea, axe9, axee)
DigitsVR(15)=Array(axf0, axf5, axfc, axfd, axf8, axf1, axf6, axff, axf2, axf3, axf4, axf7, axfb, axfa, axf9, axfe)

DigitsVR(16)=Array(a2ax00, a2ax05, a2ax0c, a2ax0d, a2ax08, a2ax01, a2ax06, a2ax0f, a2ax02, a2ax03, a2ax04, a2ax07, a2ax0b, a2ax0a, a2ax09, a2ax0e)
DigitsVR(17)=Array(a2ax10, a2ax15, a2ax1c, a2ax1d, a2ax18, a2ax11, a2ax16, a2ax1f, a2ax12, a2ax13, a2ax14, a2ax17, a2ax1b, a2ax1a, a2ax19, a2ax1e)
DigitsVR(18)=Array(a2ax20, a2ax25, a2ax2c, a2ax2d, a2ax28, a2ax21, a2ax26, a2ax2f, a2ax22, a2ax23, a2ax24, a2ax27, a2ax2b, a2ax2a, a2ax29, a2ax2e)
DigitsVR(19)=Array(a2ax30, a2ax35, a2ax3c, a2ax3d, a2ax38, a2ax31, a2ax36, a2ax3f, a2ax32, a2ax33, a2ax34, a2ax37, a2ax3b, a2ax3a, a2ax39, a2ax3e)
DigitsVR(20)=Array(a2ax40, a2ax45, a2ax4c, a2ax4d, a2ax48, a2ax41, a2ax46, a2ax4f, a2ax42, a2ax43, a2ax44, a2ax47, a2ax4b, a2ax4a, a2ax49, a2ax4e)
DigitsVR(21)=Array(a2ax50, a2ax55, a2ax5c, a2ax5d, a2ax58, a2ax51, a2ax56, a2ax5f, a2ax52, a2ax53, a2ax54, a2ax57, a2ax5b, a2ax5a, a2ax59, a2ax5e)
DigitsVR(22)=Array(a2ax60, a2ax65, a2ax6c, a2ax6d, a2ax68, a2ax61, a2ax66, a2ax6f, a2ax62, a2ax63, a2ax64, a2ax67, a2ax6b, a2ax6a, a2ax69, a2ax6e)
DigitsVR(23)=Array(a2ax70, a2ax75, a2ax7c, a2ax7d, a2ax78, a2ax71, a2ax76, a2ax7f, a2ax72, a2ax73, a2ax74, a2ax77, a2ax7b, a2ax7a, a2ax79, a2ax7e)
DigitsVR(24)=Array(a2ax80, a2ax85, a2ax8c, a2ax8d, a2ax88, a2ax81, a2ax86, a2ax8f, a2ax82, a2ax83, a2ax84, a2ax87, a2ax8b, a2ax8a, a2ax89, a2ax8e)
DigitsVR(25)=Array(a2ax90, a2ax95, a2ax9c, a2ax9d, a2ax98, a2ax91, a2ax96, a2ax9f, a2ax92, a2ax93, a2ax94, a2ax97, a2ax9b, a2ax9a, a2ax99, a2ax9e)
DigitsVR(26)=Array(a2axa0, a2axa5, a2axac, a2axad, a2axa8, a2axa1, a2axa6, a2axaf, a2axa2, a2axa3, a2axa4, a2axa7, a2axab, a2axaa, a2axa9, a2axae)
DigitsVR(27)=Array(a2axb0, a2axb5, a2axbc, a2axbd, a2axb8, a2axb1, a2axb6, a2axbf, a2axb2, a2axb3, a2axb4, a2axb7, a2axbb, a2axba, a2axb9, a2axbe)
DigitsVR(28)=Array(a2axc0, a2axc5, a2axcc, a2axcd, a2axc8, a2axc1, a2axc6, a2axcf, a2axc2, a2axc3, a2axc4, a2axc7, a2axcb, a2axca, a2axc9, a2axce)
DigitsVR(29)=Array(a2axd0, a2axd5, a2axdc, a2axdd, a2axd8, a2axd1, a2axd6, a2axdf, a2axd2, a2axd3, a2axd4, a2axd7, a2axdb, a2axda, a2axd9, a2axde)
DigitsVR(30)=Array(a2axe0, a2axe5, a2axec, a2axed, a2axe8, a2axe1, a2axe6, a2axef, a2axe2, a2axe3, a2axe4, a2axe7, a2axeb, a2axea, a2axe9, a2axee)
DigitsVR(31)=Array(a2axf0, a2axf5, a2axfc, a2axfd, a2axf8, a2axf1, a2axf6, a2axff, a2axf2, a2axf3, a2axf4, a2axf7, a2axfb, a2axfa, a2axf9, a2axfe)


Sub DisplayTimerVR_Timer
    Dim ChgLED, ii, jj, num, chg, stat, obj, b, x
    ChgLED=Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED)Then

       For ii=0 To UBound(chgLED)
        num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)
        'if (num < 32) then
        if (num < 32) then
          For Each obj In DigitsVR(num)
             If chg And 1 Then obj.visible=stat And 1
             chg=chg\2 : stat=stat\2
            Next
        Else
        end if
      Next
    end if
 End Sub


Dim Flasherseq1
Dim Flasherseq2
Dim Flasherseq3
Dim flbulb

Sub BGFlashersFast_Timer
  Select Case Flasherseq1
    Case 1:For each flbulb in VRBGFlasher1:flbulb.visible = 1: Next
    Case 2:For each flbulb in VRBGFlasher1:flbulb.visible = 1: Next
    Case 3:For each flbulb in VRBGFlasher1:flbulb.visible = 0: Next
    Case 4:For each flbulb in VRBGFlasher1:flbulb.visible = 0: Next
    Case 5:For each flbulb in VRBGFlasher1:flbulb.visible = 0: Next
    Case 6:For each flbulb in VRBGFlasher1:flbulb.visible = 0: Next
    Case 7:For each flbulb in VRBGFlasher3:flbulb.visible = 1: Next
    Case 8:For each flbulb in VRBGFlasher3:flbulb.visible = 0: Next
    Case 9:For each flbulb in VRBGFlasher3:flbulb.visible = 0: Next
  End Select
  Flasherseq1 = Flasherseq1 + 1
  If Flasherseq1 > 9 Then
  Flasherseq1 = 1
  End if
End Sub

Sub BGFlashersMid_Timer
  Select Case Flasherseq3
    Case 1:For each flbulb in VRBGFlasher2:flbulb.visible = 1: Next
    Case 2:For each flbulb in VRBGFlasher2:flbulb.visible = 0: Next
    Case 3:For each flbulb in VRBGFlasher2:flbulb.visible = 0: Next
    Case 4:For each flbulb in VRBGFlasher2:flbulb.visible = 0: Next
    Case 5:For each flbulb in VRBGFlasher2:flbulb.visible = 0: Next
    Case 6:For each flbulb in VRBGFlasher2:flbulb.visible = 0: Next
    Case 7:For each flbulb in VRBGFlasher5:flbulb.visible = 1: Next
    Case 8:For each flbulb in VRBGFlasher5:flbulb.visible = 0: Next
    Case 9:For each flbulb in VRBGFlasher5:flbulb.visible = 0: Next
    Case 10:For each flbulb in VRBGFlasher5:flbulb.visible = 0: Next
    Case 11:For each flbulb in VRBGFlasher5:flbulb.visible = 0: Next
  End Select
  Flasherseq3 = Flasherseq3 + 1
  If Flasherseq3 > 11 Then
  Flasherseq3 = 1
  End if
End Sub

Sub BGFlashersSlow_Timer
  Select Case Flasherseq2
    Case 1:For each flbulb in VRBGFlasher4:flbulb.visible = 1: Next
    Case 2:For each flbulb in VRBGFlasher4:flbulb.visible = 0: Next
    Case 3:For each flbulb in VRBGFlasher4:flbulb.visible = 0: Next
    Case 4:For each flbulb in VRBGFlasher4:flbulb.visible = 0: Next
    Case 5:For each flbulb in VRBGFlasher4:flbulb.visible = 0: Next
    Case 6:For each flbulb in VRBGFlasher4:flbulb.visible = 0: Next
    Case 7:For each flbulb in VRBGFlasher4:flbulb.visible = 0: Next
    Case 8:For each flbulb in VRBGFlasher4:flbulb.visible = 0: Next
  End Select
  Flasherseq2 = Flasherseq2 + 1
  If Flasherseq2 > 8 Then
  Flasherseq2 = 1
  End if
End Sub


' loads VR Room if set..
Sub LoadVRRoom()
DisplayTimerVR.enabled = true  'turns on display timer for VR segmented scoring
ClockTimer.Enabled=1  'enables VR clock if VR room is running.
CloudTimer.Enabled=1
timervrplunger2.Enabled=1

If Glass = 1 Then
  GlassImpurities.visible = 1
  Else
  GlassImpurities.visible = 0
End If

For Each VRObject in VRBackglassStuff: VRObject.visible = 1: Next

  If VRFlashingBackglass = 1 Then
    BGFlashersFast.enabled = True
    BGFlashersMid.enabled = True
    BGFlashersSlow.enabled = True
  Else
    BGFlashersFast.enabled = False
    BGFlashersMid.enabled = False
    BGFlashersSlow.enabled = False
    For each VRobject in VRBGFlasher3:vrobject.amount = 50: Next
    BGBulb32.amount = 50
    BGBulb67.amount = 50
  End If
end sub


Sub SetNoVRRoom()
  For Each VRObject in VRBackglassStuff: VRObject.visible = 0: Next
  For Each VRObject in VRStuff: VRObject.visible = 0: Next
    For Each VRObject in VRCab: VRObject.visible = 0: Next
  BGFlashersFast.enabled = False
  BGFlashersMid.enabled = False
  BGFlashersSlow.enabled = False
End Sub

Sub CloudTimer_Timer()
  VR_SkyTube.Rotz = VR_SkyTube.Rotz +.22
End Sub



' END VRADDED ***********************************************************************************************************************************************************


' *************************  script for LUT text display and rolling text *********************************
' script related to the script below is in keydown magnasave keys script and luts/lutpos script at the top of the script
' other objects used: layer 7 all Text0?? objects ("A") and the textures "32" to "96"

dim rollingtext
rollingtext = "                                      ******************************* This is free software, like not paid for in any way *******************************"
rollingtext = rollingtext & "                   Whirlwind credits:                 base version VP9 of Herweh used and some bits and pieces from Walamab's Whirlwind for VPX"
rollingtext = rollingtext & "       Flupper: 3d models, 3d inserts, new flipper bats, photoshop work, rendering, lighting, all scripting related to visuals except for VR"
rollingtext = rollingtext & "          Rothbauerw: all the physics coding and objects"
rollingtext = rollingtext & "        Fleep: sound engineering, sound scripting               Clark Kent: scanned playfield              "
rollingtext = rollingtext & "Blackmoor: owns actual machine, sound recordings, scanned plastics, reference images, reference videos for the lighting, testing"
rollingtext = rollingtext & "         Cosmic80 for cleaning plastics textures for previous vp9 version             Wrd1972: reference photos for ramps placement, additional sound recordings"
rollingtext = rollingtext & "           Leojreimroc: scoring displays and backglass lighting in VR            Rawd & 3rdaxis: VR room, cabinet, backbox, topper and VR script added"
rollingtext = rollingtext & "   VPW: testing the prerelease beta          "
rollingtext = rollingtext & "                                                                         "

dim textindex : textindex = 1
dim charobj(55), glyph(201)
InitDisplayText

'If VRRoom = 1 Then text010.TimerInterval = 100 : End if

Sub text010_Timer()
  dim tekst
  tekst = Mid(rollingtext, textindex, 34)
  DisplayText -1, tekst
  textindex = textindex + 1
  If textindex > len(rollingtext) - 36 then textindex = 1 : end if
End Sub

Sub text011_Timer()
  Dim anr
  If Controller.switch(11) and Controller.Switch(12) and Controller.Switch(13) Then
    text010.TimerEnabled = True
    For anr = 1 to 34 : charobj(9 + anr).opacity = 1500 : next
  Else
    ResetRollingTextTimer
  End If
End Sub

Sub myChangeLut
  table1.ColorGradeImage = luts(lutpos)
  DisplayText lutpos, luts(lutpos)
  vpmTimer.AddTimer 2000, "If lutpos = " & lutpos & " then for anr = 10 to 54 : charobj(anr).visible = 0 : next'"
End Sub

Sub ResetRollingTextTimer()
  Dim anr
  text010.TimerEnabled = False : text011.TimerEnabled = False : text011.TimerEnabled = True : textindex = 1
  For anr = 10 to 43 : charobj(anr).opacity = 4000 : charobj(anr).y = 1800 : charobj(anr).height = 150 : charobj(anr).visible = 0 : next
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
  If nr > -1 then
    tekst = "lutpos:" & nr
    For anr = 1 to len(tekst) : charobj(43 + anr).imageA = glyph(asc(mid(tekst, anr, 1))) : Next
  End If
  For anr = 1 to len(luttext)
    charobj(9 + anr).imageA = glyph(asc(mid(luttext, anr, 1)))
    If nr = -1 Then
      charobj(9 + anr).y = 1500 + sin(((textindex * 4 + anr)/20)*3.14) * 100
      charobj(9 + anr).height = 150 + cos(((textindex * 4 + anr)/20)*3.14) * 100
    End If
  Next
End Sub










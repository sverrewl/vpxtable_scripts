'
'  _____                _   _
' |  ___| __ ___  _ __ | |_(_) ___ _ __
' | |_ | '__/ _ \| '_ \| __| |/ _ \ '__|
' |  _|| | | (_) | | | | |_| |  __/ |
' |_|  |_|  \___/|_| |_|\__|_|\___|_|
'
'
'               Bally Midway Manufacturing Company (1980)
'       IPD No. 959 / November, 1980 / 4 Player
'               https://www.ipdb.org/machine.cgi?id=959
'
'
'********************************************************************************************************************************
'******** Work done by UnclePaulie on Hybrid version 0.01 - 2.0 *********
'     Implemented latest VPW standards, physics, GI VLM, 3D inserts, AllLamps routines, shadows, backglass for desktop,
'   roth drop and standup targets, standalone compatibility, sling corrections, optional DIP adjustments, updated flipper physics,
'   VR and fully hybrid, playfield_mesh , Flupper style bumpers and lighting. Used backglass images from Hauntfreaks.
'   Redbone updated Bord's scanned playfield, as well other graphics updates: plastics, flippers, targets, apron, and spinner images.
'   Iaakki assisted with inlane physics, upper left drop target bounce physics, and some insert color adjustments.
'   Full details at the bottom of the script.
'********************************************************************************************************************************
'
'
' PLEASE NOTE:
' - All game options can be adjusted by pressing F12 and then left flipper button during the game. Press start button to save your options.
'

Option Explicit
Randomize
SetLocale 1033      'Forces VBS to use english to stop crashes.

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0


'***********************************************************************
'* TABLE OPTIONS *******************************************************
'***********************************************************************

' ALL OPTIONS ARE IN THE F12 MENU, EXCEPT FOR THE SetDIPSwitches and ForceVR BELOW

const SetDIPSwitches= 0   'If you want to set the dips differently, set to 1, and then hit F6 to launch.
const ForceVR = 0     'Force VR mode on desktop, 0 is off (default): 1 is to force VR on.


'User Defined acrylic edge colors (This is if user defined option is selected for plastic edge color... Choose your own RGB)

  Dim UserAcrylColor

  UserAcrylColor = RGB(255,255,240) 'default is RGB(255,255,240)


'******************************************************

' The default ROM menu has the replay threshold values to off.  The scorecard shows replay levels of 440,000 and 680,000.  (M-1508-95-F and M-1508-95-B)
' To set these follow the instructions below.  If you downloaded UnclePaulie's version off VPU, a .nvram file is included
' which has those settings preset to those levels.
' ALSO:  Self Test positions at 16, 17 (Major Mode - Replays at score levels), 18 (sound), and 19 (HSTD - 3 credits) set to 03
'   Sound Options:  18 (set to 00 - chimes / no background.  01 - noises / no background.  02: scoreing chime effect.  03: Noise effect WITH background)
' IF you want to change those values, here's the procedure:

' Here is the procedure to set the selftest postions on this Bally 35 MPU table for the preset replay levels:

' 1. Press 7 to enter test mode. Continue to repeatedly press 7 until you see 01 appear in the ball in play/match display
' 2. Press and HOLD the 1 key until it reaches the value you want to set it to. The value will increase by 1000 while you are holding in the 1 key.
'   If you happen to go past the value you want to set it to, while holding in the 1 key press and release the 7 key.
'   The value will then decrement by 1000 while you are holding in the 1 key.
'   When you reach the value you want, release the 1 key.
' 3.  Press 7 to scroll the next settings and change values in the similar manner.
' 4.  Press F3 to save your changes and reset the table.

' Here are the recommended values to change to:

  ' 01: 440,000 - First threshold to award a replay
  ' 02: 680,000 - First threshold to award a replay
  ' 16: set to 02 for extra ball to be awarded on game specials
  ' 17: set to 03 for scoring thresholds award of replay
  ' 18: set to 01 for noise effect without background
  ' 19: set to 03 for 3 credits awarded for HSTD or over 10,000,000

'***********************************************************************

'
'*********************************************************************************************************************************
' === TABLE OF CONTENTS  ===
'
' You can quickly jump to a section by searching the four letter tag (ZXXX)
'
' ZVAR: Constants and Global Variable
' ZTIM: Timers
' ZINI: Table Initialization and Exiting
'   ZOPT: User Options
'   ZMAT: General Math Functions
' ZANI: Misc Animations
'   ZBBR: Ball Brightness
'   ZRBR: Room Brightness
' ZKEY: Key Press Handling
' ZSOL: Solenoids & Flashers
' ZDRN: Drain, Trough, and Ball Release
' ZFLP: Flippers
' ZSLG: Slingshot Animations
' ZSSC: Slingshot Corrections
' ZSWI: Switches
'   ZBRL: Ball Rolling and Drop Sounds
'   ZRRL: Ramp Rolling Sound Effects
'   ZFLE: Fleep Mechanical Sounds
' ZNFF: Flipper Corrections
'   ZDMP: Rubber Dampeners
'   ZRDT: Drop Targets
' ZRST: Stand-Up Targets
'   ZBOU: VPW TargetBouncer for targets and posts
'   ZGIU: GI updates
'   Z3DI: 3D Inserts
'   ZFLD: Flupper Domes
'   ZFLB: Flupper Bumpers
' ZSHA: Ambient Ball Shadows
'   ZVRR: VR Room / VR Cabinet
'
'*********************************************************************************************************************************

  Dim VR_Room, cab_mode, DesktopMode

  If ForceVR = 1 Then
    DesktopMode = False
    VR_Room = 1
    cab_mode = 0
  Else
    DesktopMode = Table1.ShowDT
    If RenderingMode = 2 Then VR_Room=1 Else VR_Room=0
    if Not DesktopMode and VR_Room=0 Then cab_mode=1 Else cab_mode=0
  End If

' ****************************************************

'******************************************************
'  ZVAR: Constants and Global Variables
'******************************************************
Const BallSize = 50
Const BallMass = 1
Const tnob = 1 ' total number of balls
Const lob = 0   'number of locked balls

Dim tablewidth: tablewidth = Table1.width
Dim tableheight: tableheight = Table1.height
Dim BIPL : BIPL = False   'Ball in plunger lane
Dim DT21up, DT22up, DT23up, DT33up, DT34up, DT35up
Dim flbumpvalue1: flbumpvalue1 = 0
Dim flbumpvalue2: flbumpvalue2 = 0
Dim flbumpvalue3: flbumpvalue3 = 0
Dim gilvl
Dim PlasProt
Const rcolor = 201
Const gcolor = 204
Const bcolor = 202
Dim gate4soundrdy: gate4soundrdy = 0  'used for gate4 metal hit sounds after ball goes through gate.
Dim l40on: l40on = 0          'used for additional bdl on spinner prim when insert lit.

Dim FBall1, gBOT

'  Standard definitions
Const cGameName = "frontier"    'PinMAME ROM name
Const UseSolenoids = 2      '1 = Normal Flippers, 2 = Fastflips
Const UseLamps = 1        '0 = Custom lamp handling, 1 = Built-in VPX handling (using light number in light timer)
Const UseSync = 0
Const HandleMech = 0

' Insert brightness scale
const insertfactor    = 1   ' adjusts the level of the brightness of the inserts when on

LoadVPM "01000100", "Bally.VBS", 3.02  'Frontier

'******************************************************
' VLM Arrays
'******************************************************

' VLM  Arrays - Start
' Arrays per baked part
Dim BP_Playfield: BP_Playfield=Array(LM_All_Lights_gi_001_Playfield, LM_All_Lights_gi_002_Playfield, LM_All_Lights_gi_003_Playfield, LM_All_Lights_gi_004_Playfield, LM_All_Lights_gi_005_Playfield, LM_All_Lights_gi_006_Playfield, LM_All_Lights_gi_007_Playfield, LM_All_Lights_gi_008_Playfield, LM_All_Lights_gi_009_Playfield, LM_All_Lights_gi_010_Playfield, LM_All_Lights_gi_011_Playfield, LM_All_Lights_gi_012_Playfield, LM_All_Lights_gi_013_Playfield, LM_All_Lights_gi_014_Playfield, LM_All_Lights_gi_015_Playfield, LM_All_Lights_gi_016_Playfield, LM_All_Lights_gi_017_Playfield, LM_All_Lights_gi_018_Playfield, LM_All_Lights_gi_019_Playfield, LM_All_Lights_gi_020_Playfield, LM_All_Lights_gi_021_Playfield, LM_All_Lights_gi_022_Playfield, LM_All_Lights_gi_023_Playfield)
Dim BP_Parts: BP_Parts=Array(BM_Parts, LM_All_Lights_gi_002_Parts, LM_All_Lights_gi_003_Parts, LM_All_Lights_gi_004_Parts, LM_All_Lights_gi_005_Parts, LM_All_Lights_gi_007_Parts, LM_All_Lights_gi_009_Parts, LM_All_Lights_gi_010_Parts, LM_All_Lights_gi_011_Parts, LM_All_Lights_gi_016_Parts, LM_All_Lights_gi_017_Parts, LM_All_Lights_gi_019_Parts, LM_All_Lights_gi_020_Parts, LM_All_Lights_gi_021_Parts, LM_All_Lights_l12_Parts, LM_All_Lights_l28_Parts, LM_All_Lights_l44_Parts)
' Arrays per lighting scenario
Dim BL_All_Lights_gi_001: BL_All_Lights_gi_001=Array(LM_All_Lights_gi_001_Playfield)
Dim BL_All_Lights_gi_002: BL_All_Lights_gi_002=Array(LM_All_Lights_gi_002_Playfield,LM_All_Lights_gi_002_Parts)
Dim BL_All_Lights_gi_003: BL_All_Lights_gi_003=Array(LM_All_Lights_gi_003_Playfield,LM_All_Lights_gi_003_Parts)
Dim BL_All_Lights_gi_004: BL_All_Lights_gi_004=Array(LM_All_Lights_gi_004_Playfield,LM_All_Lights_gi_004_Parts)
Dim BL_All_Lights_gi_005: BL_All_Lights_gi_005=Array(LM_All_Lights_gi_005_Playfield,LM_All_Lights_gi_005_Parts)
Dim BL_All_Lights_gi_006: BL_All_Lights_gi_006=Array(LM_All_Lights_gi_006_Playfield)
Dim BL_All_Lights_gi_007: BL_All_Lights_gi_007=Array(LM_All_Lights_gi_007_Playfield,LM_All_Lights_gi_007_Parts)
Dim BL_All_Lights_gi_008: BL_All_Lights_gi_008=Array(LM_All_Lights_gi_008_Playfield)
Dim BL_All_Lights_gi_009: BL_All_Lights_gi_009=Array(LM_All_Lights_gi_009_Playfield,LM_All_Lights_gi_009_Parts)
Dim BL_All_Lights_gi_010: BL_All_Lights_gi_010=Array(LM_All_Lights_gi_010_Playfield,LM_All_Lights_gi_010_Parts)
Dim BL_All_Lights_gi_011: BL_All_Lights_gi_011=Array(LM_All_Lights_gi_011_Playfield,LM_All_Lights_gi_011_Parts)
Dim BL_All_Lights_gi_012: BL_All_Lights_gi_012=Array(LM_All_Lights_gi_012_Playfield)
Dim BL_All_Lights_gi_013: BL_All_Lights_gi_013=Array(LM_All_Lights_gi_013_Playfield)
Dim BL_All_Lights_gi_014: BL_All_Lights_gi_014=Array(LM_All_Lights_gi_014_Playfield)
Dim BL_All_Lights_gi_015: BL_All_Lights_gi_015=Array(LM_All_Lights_gi_015_Playfield)
Dim BL_All_Lights_gi_016: BL_All_Lights_gi_016=Array(LM_All_Lights_gi_016_Playfield,LM_All_Lights_gi_016_Parts)
Dim BL_All_Lights_gi_017: BL_All_Lights_gi_017=Array(LM_All_Lights_gi_017_Playfield,LM_All_Lights_gi_017_Parts)
Dim BL_All_Lights_gi_018: BL_All_Lights_gi_018=Array(LM_All_Lights_gi_018_Playfield)
Dim BL_All_Lights_gi_019: BL_All_Lights_gi_019=Array(LM_All_Lights_gi_019_Playfield,LM_All_Lights_gi_019_Parts)
Dim BL_All_Lights_gi_020: BL_All_Lights_gi_020=Array(LM_All_Lights_gi_020_Playfield,LM_All_Lights_gi_020_Parts)
Dim BL_All_Lights_gi_021: BL_All_Lights_gi_021=Array(LM_All_Lights_gi_021_Playfield,LM_All_Lights_gi_021_Parts)
Dim BL_All_Lights_gi_022: BL_All_Lights_gi_022=Array(LM_All_Lights_gi_022_Playfield)
Dim BL_All_Lights_gi_023: BL_All_Lights_gi_023=Array(LM_All_Lights_gi_023_Playfield)
Dim BL_All_Lights_l12: BL_All_Lights_l12=Array(LM_All_Lights_l12_Parts)
Dim BL_All_Lights_l28: BL_All_Lights_l28=Array(LM_All_Lights_l28_Parts)
Dim BL_All_Lights_l44: BL_All_Lights_l44=Array(LM_All_Lights_l44_Parts)
Dim BL_World: BL_World=Array(BM_Parts)
' Global arrays
Dim BG_Bakemap: BG_Bakemap=Array(BM_Parts)
Dim BG_Lightmap: BG_Lightmap=Array(LM_All_Lights_gi_001_Playfield, LM_All_Lights_gi_002_Playfield, LM_All_Lights_gi_003_Playfield, LM_All_Lights_gi_004_Playfield, LM_All_Lights_gi_005_Playfield, LM_All_Lights_gi_006_Playfield, LM_All_Lights_gi_007_Playfield, LM_All_Lights_gi_008_Playfield, LM_All_Lights_gi_009_Playfield, LM_All_Lights_gi_010_Playfield, LM_All_Lights_gi_011_Playfield, LM_All_Lights_gi_012_Playfield, LM_All_Lights_gi_013_Playfield, LM_All_Lights_gi_014_Playfield, LM_All_Lights_gi_015_Playfield, LM_All_Lights_gi_016_Playfield, LM_All_Lights_gi_017_Playfield, LM_All_Lights_gi_018_Playfield, LM_All_Lights_gi_019_Playfield, LM_All_Lights_gi_020_Playfield, LM_All_Lights_gi_021_Playfield, LM_All_Lights_gi_022_Playfield, LM_All_Lights_gi_023_Playfield, LM_All_Lights_gi_002_Parts, LM_All_Lights_gi_003_Parts, LM_All_Lights_gi_004_Parts, LM_All_Lights_gi_005_Parts, LM_All_Lights_gi_007_Parts, LM_All_Lights_gi_009_Parts, LM_All_Lights_gi_010_Parts, LM_All_Lights_gi_011_Parts, LM_All_Lights_gi_016_Parts, LM_All_Lights_gi_017_Parts, LM_All_Lights_gi_019_Parts, LM_All_Lights_gi_020_Parts, LM_All_Lights_gi_021_Parts, LM_All_Lights_l12_Parts, LM_All_Lights_l28_Parts, LM_All_Lights_l44_Parts)
Dim BG_All: BG_All=Array(LM_All_Lights_gi_001_Playfield, LM_All_Lights_gi_002_Playfield, LM_All_Lights_gi_003_Playfield, LM_All_Lights_gi_004_Playfield, LM_All_Lights_gi_005_Playfield, LM_All_Lights_gi_006_Playfield, LM_All_Lights_gi_007_Playfield, LM_All_Lights_gi_008_Playfield, LM_All_Lights_gi_009_Playfield, LM_All_Lights_gi_010_Playfield, LM_All_Lights_gi_011_Playfield, LM_All_Lights_gi_012_Playfield, LM_All_Lights_gi_013_Playfield, LM_All_Lights_gi_014_Playfield, LM_All_Lights_gi_015_Playfield, LM_All_Lights_gi_016_Playfield, LM_All_Lights_gi_017_Playfield, LM_All_Lights_gi_018_Playfield, LM_All_Lights_gi_019_Playfield, LM_All_Lights_gi_020_Playfield, LM_All_Lights_gi_021_Playfield, LM_All_Lights_gi_022_Playfield, LM_All_Lights_gi_023_Playfield, BM_Parts, LM_All_Lights_gi_002_Parts, LM_All_Lights_gi_003_Parts, LM_All_Lights_gi_004_Parts, LM_All_Lights_gi_005_Parts, LM_All_Lights_gi_007_Parts, LM_All_Lights_gi_009_Parts, LM_All_Lights_gi_010_Parts, LM_All_Lights_gi_011_Parts, LM_All_Lights_gi_016_Parts, LM_All_Lights_gi_017_Parts, LM_All_Lights_gi_019_Parts, LM_All_Lights_gi_020_Parts, LM_All_Lights_gi_021_Parts, LM_All_Lights_l12_Parts, LM_All_Lights_l28_Parts, LM_All_Lights_l44_Parts)
' VLM  Arrays - End

'******************************************************
'  ZTIM: Timers
'******************************************************

'The FrameTimer interval should be -1, so executes at the display frame rate
'The frame timer should be used to update anything visual, like some animations, shadows, etc.
'However, a lot of animations will be handled in their respective _animate subroutines.

Dim FrameTime, InitFrameTime
InitFrameTime = 0

FrameTimer.Interval = -1
Sub FrameTimer_Timer()
  FrameTime = gametime - InitFrameTime 'Calculate FrameTime as some animuations could use this
  InitFrameTime = gametime  'Count frametime
  'Add animation stuff here
  BSUpdate
  UpdateBallBrightness
  RollingUpdate       'Update rolling sounds
  DoDTAnim          'Drop target animations
  DoSTAnim          'Standup target animations
  BSUpdate          'Update ambient ball shadows
  AnimateBumperSkirts     'Base skirts will animate when a ball touches them

    If VR_Room = 0 and cab_mode = 0 Then
        DisplayTimer
    End If

  If VR_Room=1 Then
    VRDisplayTimer
  End If

End Sub

'The CorTimer interval should be 10. It's sole purpose is to update the Cor (physics) calculations

CorTimer.Interval = 10
Sub CorTimer_Timer(): Cor.Update: End Sub

'******************************************************
'  ZINI: Table Initialization and Exiting
'******************************************************

Sub Table1_Init
  vpminit me
  With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
    .SplashInfoLine = "Frontier (Bally 1980)"&chr(13)&"by UnclePaulie"
    .HandleKeyboard = 0
    .ShowTitle = 0
    .ShowDMDOnly = 1
    .ShowFrame = 0
    .HandleMechanics = 0
    .Games(cGameName).Settings.Value("sound") = 1 ' Set sound (0=OFF, 1=ON)
  End With
  On Error Resume Next
  Controller.Run
  If Err Then MsgBox Err.Description
  On Error Goto 0

  'Map all lamps to the corresponding ROM output using the value of TimerInterval of each light object
  vpmMapLights AllLamps     'Make a collection called "AllLamps" and put all the light objects in it.

  'Nudging
  vpmNudge.TiltSwitch = 7
  vpmNudge.Sensitivity = 3
  vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingshot)

  'Trough - Creates a ball in the kicker switch and gives that ball used an individual name.
  Set FBall1 = BallRelease.CreateSizedballWithMass(Ballsize/2,Ballmass)

  'Forces the trough switches to "on" at table boot so the game logic knows there are balls in the trough.
  Controller.Switch(8) = 1

  '***Setting up a ball array (collection), must contain all the balls you create on the table.
  gBOT = Array(FBall1)

  ' GI initially off
  PFGI 0
  gilvl = 0

  ' GI comes on after 2 seconds
  vpmTimer.AddTimer 2000, "PFGI 1'"

  if VR_Room = 1 Then
    setup_backglass()
  End If

' If user wants to set dip switches themselves it will force them to set it via F6.
  If SetDIPSwitches = 0 Then
    SetDefaultDips
  End If

' Make drop target shadows visible
  Dim xx
  for each xx in ShadowDT
    xx.visible=False
  Next

' Drop Target Variable state for DT Shadows
  DT21up=1
  DT22up=1
  DT23up=1
  DT33up=1
  DT34up=1
  DT35up=1

' Initialize Bumpers
    FlFadeBumper 1, 0
    FlFadeBumper 2, 0
    FlFadeBumper 3, 0
    flbumpvalue1 = 0
    flbumpvalue2 = 0
    flbumpvalue3 = 0

End Sub

' Thalamus - should be table1
Sub Table1_Paused:Controller.Pause = 1:End Sub
Sub Table1_unPaused:Controller.Pause = 0:End Sub
' Thalamus, Table is named Table1 - not Table
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


'*******************************************
'  ZOPT: User Options
'*******************************************

' Table1.Option arguments are:
' - option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional arry of literal strings

' Called when options are tweaked by the player.
' - 0: game has started, good time to load options and adjust accordingly
' - 1: an option has changed
' - 2: options have been reseted
' - 3: player closed the tweak UI, good time to update staticly prerendered parts




const lightlevelinitial = 0.15          ' Initial Level of room lighting (0 to 1), where 0 is dark and 100 is brightest
Dim LightLevel : LightLevel = lightlevelinitial ' Level of room lighting
Dim ColorLUT : ColorLUT = 1           ' Color desaturation LUTs: 1 to 11, where 1 is normal and 11 is black'n'white
Dim VolumeDial : VolumeDial = 0.8             ' Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Dim BallRollVolume : BallRollVolume = 0.5     ' Level of ball rolling volume. Value between 0 and 1
Dim RampRollVolume : RampRollVolume = 0.5     ' Level of ramp rolling volume. Value between 0 and 1
Dim ROMVolume                 ' Level of ROM volume. Value between 0 and 1
Dim WallClock, poster, poster2, CustomWalls
Dim ChooseBall: ChooseBall = 4
              '0 = dark, 1 = not as dark, 2 = bright, 3 = brightest, 4 = warm, 5 = MrBallDark, 6 = DarthVito Ball1
              '7 = DarthVito HDR1, 8 = DarthVito HDR Bright, 9 = DarthVito HDR2 (Default), 10 = DarthVito ballHDR3,
              '11 = DarthVito ballHDR4, 12 = DarthVito ballHDRdark, 13 = Borg Ball, 14 = SteelBall2


' Bumper brightness scale
dim bumperscale:    bumperscale =     0.22  ' adjusts the level of blenddisablelighting and intensity on bumpers when lights are not on
dim bumperflashlevel:   bumperflashlevel =  0.4   ' adjusts the level of bumper lights when on and / or flashing


Dim FlipperColor, ScrewCapColor, cabsideblades, PlasticColor, TargetColor, undercab

Dim dspTriggered : dspTriggered = False
Sub Table1_OptionEvent(ByVal eventId)
    If eventId = 1 And Not dspTriggered Then dspTriggered = True : DisableStaticPreRendering = True : End If

  ROMVolume = Table1.Option("ROM Volume (-32 off; 0 max) - Requires Restart; altsound set to 0", -32, 0, 1, 0, 0)

    ' Ball Choice
    ChooseBall = Table1.Option("Ball Choice", 0, 14, 1, 4, 0, Array("0: Hauntfreaks dark", "1: Hauntfreaks not as dark", _
      "2: Hauntfreaks bright", "3: Hauntfreaks brightest", "4: Hauntfreaks warm (default)", "5: MrBallDark", "6: DarthVito Ball1", _
      "7: DarthVito HDR1", "8: DarthVito HDR Bright", "9: DarthVito HDR2", "10: DarthVito ballHDR3", _
      "11: DarthVito ballHDR4", "12: DarthVito ballHDRdark", "13: Borg Ball", "14: SteelBall2"))
  ChangeBall(ChooseBall)

  ' Bumper Options
  bumperscale = Table1.Option("Bumper brightness when not lit (default 22%)", 0, 1, 0.01, 0.22, 1)
  bumperflashlevel = Table1.Option("Bumper brightness when lit or flashing (default 40%)", 0, 1, 0.01, 0.40, 1)
  FlFadeBumper 1, flbumpvalue1
  FlFadeBumper 2, flbumpvalue2
  FlFadeBumper 3, flbumpvalue3

  ' General Options
  FlipperColor = Table1.Option("Flipper Color", 1, 3, 1, 1, 0, Array("Red (default)", "Blue", "Yellow"))
  PlasticColor = Table1.Option("Plastic Edge Color", 0, 6, 1, 0, 0, Array("Off (default)", "Orange/Red", "Yellow", "Blue", "User Defined in Script", "Rainbow", "Random"))
  TargetColor = Table1.Option("Stand Up Target Color", 0, 1, 1, 0, 0, Array("Red Worn with streaks(default)", "Red"))
  ScrewCapColor = Table1.Option("Screw Cap Color", 0, 5, 1, 0, 0, Array("White (default)", "Black", "Bronze", "Silver", "Blue/Green", "Gold"))
  cabsideblades = Table1.Option("Cabinet Sideblades - full cabinet only", 0, 1, 1, 0, 0, Array("Extra Tall - default in cab mode", "Lowered Blades"))

  ' VR Room Environment

  WallClock = Table1.Option("VR Wall Clock On/Off", 0, 1, 1, 1, 0, Array("Off", "On"))
    CustomWalls = Table1.Option("VR Room Environment", 0, 4, 1, 2, 0, _
    Array("UP Original", "Sixtoe Arcade", "DarthVito Home", "DarthVito Plaster", "DarthVito Blue"))
  Poster = Table1.Option("VR Poster 1 On/Off", 0, 1, 1, 1, 0, Array("Off", "On"))
  Poster2 = Table1.Option("VR Poster 2 On/Off", 0, 1, 1, 1, 0, Array("Off", "On"))
  undercab = Table1.Option("VR Undercab lighting On/Off", 0, 1, 1, 1, 0, Array("Off", "On"))

  SetupOptions

    ' Sound volumes
  VolumeDial = Table1.Option("Mech Volume", 0, 1, 0.01, 0.8, 1)
  BallRollVolume = Table1.Option("Ball Roll Volume", 0, 1, 0.01, 0.5, 1)
  RampRollVolume = Table1.Option("Ramp Roll Volume", 0, 1, 0.01, 0.5, 1)

  ' Color Saturation
    ColorLUT = Table1.Option("Color Saturation", 1, 11, 1, 1, 0, _
    Array("Normal", "Desaturated 10%", "Desaturated 20%", "Desaturated 30%", "Desaturated 40%", "Desaturated 50%", _
        "Desaturated 60%", "Desaturated 70%", "Desaturated 80%", "Desaturated 90%", "Black 'n White"))
  if ColorLUT = 1 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1" ' the image was blank for default, i.e.:  just ""  - this fix if for VR Headtracking
  if ColorLUT = 2 Then Table1.ColorGradeImage = "colorgradelut256x16-10"
  if ColorLUT = 3 Then Table1.ColorGradeImage = "colorgradelut256x16-20"
  if ColorLUT = 4 Then Table1.ColorGradeImage = "colorgradelut256x16-30"
  if ColorLUT = 5 Then Table1.ColorGradeImage = "colorgradelut256x16-40"
  if ColorLUT = 6 Then Table1.ColorGradeImage = "colorgradelut256x16-50"
  if ColorLUT = 7 Then Table1.ColorGradeImage = "colorgradelut256x16-60"
  if ColorLUT = 8 Then Table1.ColorGradeImage = "colorgradelut256x16-70"
  if ColorLUT = 9 Then Table1.ColorGradeImage = "colorgradelut256x16-80"
  if ColorLUT = 10 Then Table1.ColorGradeImage = "colorgradelut256x16-90"
  if ColorLUT = 11 Then Table1.ColorGradeImage = "colorgradelut256x16-100"

  ' Room brightness
' LightLevel = Table1.Option("Table Brightness (Ambient Light Level)", 0, 1, 0.01, .5, 1)
  LightLevel = NightDay/100
  SetRoomBrightness   'Uncomment this line for lightmapped tables.

    If eventId = 3 And dspTriggered Then dspTriggered = False : DisableStaticPreRendering = False : End If

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
'  ZANI: Misc Animations
'******************************************************

Sub LeftFlipper_Animate
  dim a: a = LeftFlipper.CurrentAngle
  FlipperLSh.RotZ = a
  Lflipmesh1.RotZ =a
End Sub

Sub RightFlipper_Animate
  dim a: a = RightFlipper.CurrentAngle
  FlipperRSh.RotZ = a
  Rflipmesh1.RotZ =a
End Sub

Sub Gate_001_Animate
  dim a: a = (Gate_001.CurrentAngle * -0.5) + 6.5
  GateFlap_001.Rotx = a
End Sub

Sub Gate_002_Animate
  dim a: a = (Gate_002.CurrentAngle * -0.5) + 6.5
  GateFlap_002.Rotx = a
End Sub

Sub Gate_003_Animate
  dim a: a = (Gate_003.CurrentAngle - 25)
  Gate_003p.Rotx = max(a,20)
  Gate_003Rod.z = ((Gate_003p.Rotx * 0.089) + 63.22)
End Sub

Sub Gate_004_Animate
  dim a: a = (Gate_004.CurrentAngle)
  Gate_004p.Rotx = max(a,0)
End Sub


''''' Bumper Animations

Sub Bumper1_Animate
  Dim a
  a = Bumper1.CurrentRingOffset
  BumperRing1.transz = a
End Sub

Sub Bumper2_Animate
  Dim a
  a = Bumper2.CurrentRingOffset
  BumperRing2.transz = a
End Sub

Sub Bumper3_Animate
  Dim a
  a = Bumper3.CurrentRingOffset
  BumperRing3.transz = a
End Sub

Dim Bumpers : Bumpers = Array(Bumper1, Bumper2, Bumper3)

Sub AnimateBumperSkirts
  dim r, g, s, x, y, b, tz
  ' Animate Bumper switch (experimental)
  For r = 0 To 2
    g = 10000.0
    For s = 0 to UBound(gBOT)
      x = Bumpers(r).x - gBOT(s).x
      y = Bumpers(r).y - gBOT(s).y
      b = x * x + y * y
      If b < g Then g = b
    Next
    tz = 0
    If g < 80 * 80 Then
      tz = -4.5 '-3 (UP changed the distance to be 6 VPU... so bumper skirts are now z=-1.5 and than transz is total of -1.5 + -4.5 = -6)
    End If
    If r = 0 Then bumperdisk1.transZ = tz
    If r = 1 Then bumperdisk2.transZ = tz
    If r = 2 Then bumperdisk3.transZ = tz
  Next
End Sub


'******************************************************
'   ZBBR: BALL BRIGHTNESS
'******************************************************

Const BallBrightness =  1       'Ball brightness - Value between 0 and 1 (0=Dark ... 1=Bright)

' Constants for plunger lane ball darkening.
' You can make a temporary wall in the plunger lane area and use the co-ordinates from the corner control points.
Const PLOffset = 0.65     'Minimum ball brightness scale in plunger lane
Const PLLeft = 860        'X position of punger lane left
Const PLRight = 952       'X position of punger lane right
Const PLTop = 1060        'Y position of punger lane top
Const PLBottom = 1900       'Y position of punger lane bottom
Dim PLGain: PLGain = (1-PLOffset)/(PLTop-PLBottom)

Sub UpdateBallBrightness
  Dim s, b_base, b_r, b_g, b_b, d_w
  b_base = 120 * BallBrightness + 135*gilvl ' orig was 120 and 70

  For s = 0 To UBound(gBOT)
    ' Handle z direction
    d_w = b_base*(1 - (gBOT(s).z-25)/500)
    If d_w < 30 Then d_w = 30
    ' Handle plunger lane
    If InRect(gBOT(s).x,gBOT(s).y,PLLeft,PLBottom,PLLeft,PLTop,PLRight,PLTop,PLRight,PLBottom) Then
      d_w = d_w*(PLOffset+PLGain*(gBOT(s).y-PLBottom))
    End If
    ' Assign color
    b_r = Int(d_w)
    b_g = Int(d_w)
    b_b = Int(d_w)
    If b_r > 255 Then b_r = 255
    If b_g > 255 Then b_g = 255
    If b_b > 255 Then b_b = 255
    gBOT(s).color = b_r + (b_g * 256) + (b_b * 256 * 256)
    'debug.print "--- ball.color level="&b_r
  Next
End Sub


'****************************
'   ZRBR: Room Brightness
'****************************

Sub SetRoomBrightness
Dim girubbercolor_r, girubbercolor_g, girubbercolor_b, x

' VR Room

  For each x in VRRoomLL: x.blenddisablelighting = 1 * (LightLevel/lightlevelinitial): Next
  For each x in VRCabLL: x.blenddisablelighting = 0.5 * (LightLevel/lightlevelinitial): Next
  For each x in LLVRButtons: x.blenddisablelighting = 0.1125 * (LightLevel/lightlevelinitial): Next
  VRposter.blenddisablelighting = 0.5 * (LightLevel/lightlevelinitial)
  VRposter2.blenddisablelighting = 0.5 * (LightLevel/lightlevelinitial)
  VR_Clockface.opacity = 100 * (LightLevel/lightlevelinitial)
  BGBright.opacity = 25 + 75 * (LightLevel/lightlevelinitial)
  BGDark.opacity = 50 * LightLevel/lightlevelinitial
  VRNewFloor.opacity = 100 * (LightLevel/lightlevelinitial)
  For each x in TopHardware: x.blenddisablelighting = 0.05 * (LightLevel/lightlevelinitial): Next
  UndercabLight.intensity = 15 * gilvl

' GI and GI Bulb Prims
  For each x in GI: x.State = gilvl: Next
  For each x in GIBulbPrims: x.blenddisablelighting = (1.25 * (gilvl)): Next

' rubbers
  girubbercolor_r = (gilvl * 53 * LightLevel/lightlevelinitial) + rcolor
  girubbercolor_g = (gilvl * 50 * LightLevel/lightlevelinitial) + gcolor
  girubbercolor_b = (gilvl * 42 * LightLevel/lightlevelinitial) + bcolor
  MaterialColor "Rubber White",RGB(girubbercolor_r,girubbercolor_g,girubbercolor_b)

' flippers
  Lflipmesh1.blenddisablelighting = (0.065 * (gilvl)) + (0.035 * LightLevel/lightlevelinitial)
  Rflipmesh1.blenddisablelighting = (0.065 * (gilvl)) + (0.035 * LightLevel/lightlevelinitial)

' targets
  For each x in GITargets: x.blenddisablelighting = (0.1 * gilvl) + (0.025 * LightLevel/lightlevelinitial): Next
  sw30p.blenddisablelighting = (0.02 * gilvl) + (0.005 * LightLevel/lightlevelinitial)
  For each x in GIdtTargets: x.blenddisablelighting = (0.075 * gilvl) + (0.025 * LightLevel/lightlevelinitial): Next

' prims
  PlungerCover.blenddisablelighting = 0.025 * gilvl * LightLevel/lightlevelinitial
  LaneGuard.blenddisablelighting = 0.025 * gilvl * LightLevel/lightlevelinitial
  For each x in GIScrewCaps: x.blenddisablelighting = ((0.025 * gilvl) + (0.025 * LightLevel/lightlevelinitial)): Next
  For each x in apronwallgi: x.blenddisablelighting = 0.05 * gilvl * LightLevel/lightlevelinitial: Next
  BM_Parts.blenddisablelighting = LightLevel/lightlevelinitial
  OuterPrimBlack_DT.blenddisablelighting = (0.05 * gilvl) + (0.05 * LightLevel/lightlevelinitial)
  For each x in GICutoutPrims: x.blenddisablelighting = (0.15 * (gilvl)) + (0.1 * LightLevel/lightlevelinitial): Next
  cutout_prim_bulb_large.blenddisablelighting = (0.5 * (gilvl)) + (0.1 * LightLevel/lightlevelinitial)
  For each x in GILaneGuides: x.blenddisablelighting = (0.125 * (gilvl)) + (0.125 * LightLevel/lightlevelinitial) : Next
  For each x in GIMetal: x.blenddisablelighting = (0.015 * gilvl) + (0.01 * LightLevel/lightlevelinitial): Next
  For each x in GIMetalWall: x.blenddisablelighting = (0.0575 * gilvl) + (0.0025 * LightLevel/lightlevelinitial): Next
  For each x in GIMetalRails: x.blenddisablelighting = (0.005 * gilvl) + (0.005 * LightLevel/lightlevelinitial): Next
  For each x in GIMetalPosts: x.blenddisablelighting = (0.005 * gilvl) + (0.005 * LightLevel/lightlevelinitial): Next
  For each x in GILeafswitch: x.blenddisablelighting = ((0.2 * gilvl) + (0.005 * LightLevel/lightlevelinitial)): Next
  For each x in GIPegs: x.blenddisablelighting = (0.15 * gilvl) + (0.005 * LightLevel/lightlevelinitial) : Next
  PegPlasticRail_001.blenddisablelighting = (0.05 * gilvl) + (0.005 * LightLevel/lightlevelinitial)
  PegPlasticRail_002.blenddisablelighting = (0.05 * gilvl) + (0.005 * LightLevel/lightlevelinitial)
  Sling1.blenddisablelighting = (0.005 * gilvl) + (0.005 * LightLevel/lightlevelinitial)
  Sling2.blenddisablelighting = (0.005 * gilvl) + (0.005 * LightLevel/lightlevelinitial)
  MaterialColor "Plastic with an image cards",RGB(girubbercolor_r,girubbercolor_g,girubbercolor_b)

'saucer
  SaucerBally.blenddisablelighting = (0.125 * gilvl) + (0.025 * LightLevel/lightlevelinitial)
  kickarmtop_prim.blenddisablelighting = (0.125 * gilvl) + (0.025 * LightLevel/lightlevelinitial)

'spinner
  SpinnerPrim.blenddisablelighting = ((0.25 * gilvl) + (0.0625 * l40on)) + (((0.0625 * l40on) + (0.05) * LightLevel/lightlevelinitial))

' Pegs by Bumpers controlled in bumper animation sub
  l26_animate


' *** UPDATE as needed after baking and the GI collections above for whatever is baked.

' BM_Parts.blenddisablelighting = LightLevel/lightlevelinitial
' BM_SW_Parts.blenddisablelighting = LightLevel/lightlevelinitial
' BM_SW_Parts1.blenddisablelighting = 0.1*LightLevel/lightlevelinitial
' BM_Posts.opacity = ((35 * gilvl) + (40 * LightLevel/lightlevelinitial))


End Sub


'******************************************************
' ZKEY: Key Press Handling
'******************************************************

'*********************
' VR Plunger code
'*********************

Sub TimerVRPlunger_Timer
  If VR_Shooter.Y < 86 then
      VR_Shooter.Y = VR_Shooter.Y + 5
  End If
End Sub

Sub TimerVRPlunger1_Timer
  VR_Shooter.Y = -4 + (5* Plunger.Position) -20
End Sub


'******************************************************

Sub Table1_KeyDown(ByVal keycode) '***What to do when a button is pressed***

  If keycode = LeftTiltKey Then Nudge 90, 0.5 : SoundNudgeLeft
  If keycode = RightTiltKey Then Nudge 270, 0.5 : SoundNudgeRight
  If keycode = CenterTiltKey Then Nudge 0, 0.5 : SoundNudgeCenter

  'If keycode = keyInsertCoin1 or keycode = keyInsertCoin2 or keycode = keyInsertCoin3 or keycode = keyInsertCoin4 Then 'Use this for ROM based games
  If keycode = AddCreditKey or keycode = AddCreditKey2 Then
    Select Case Int(rnd*3)
      Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
      Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
      Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
    End Select
  End If

  If KeyCode = PlungerKey Then
    Plunger.Pullback:SoundPlungerPull()
    TimerVRPlunger.Enabled = True
    TimerVRPlunger1.Enabled = False
    VR_Shooter.Y = -4
  End If

  If keycode = LeftFlipperKey Then
    VRFlipperButtonLeft.X = 2112 + 6
  End If
  If keycode = RightFlipperKey Then
    VRFlipperButtonRight.X = 2094 - 6
  End If

  If keycode = StartGameKey Then
    StartButton.y = 1916 - 5
    SoundStartButton
  End If

  If KeyDownHandler(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal keycode)   '***What to do when a button is released***
  If keycode = PlungerKey Then
    dim pspeed: pspeed = RndInt (1,3)
    If pspeed = 1 then Plunger.firespeed = 102
    If pspeed = 2 then Plunger.firespeed = 107
    If pspeed = 3 then Plunger.firespeed = 109
    Plunger.Fire
    If BIPL = 1 Then
      SoundPlungerReleaseBall()     'Plunger release sound when there is a ball in shooter lane
    Else
      SoundPlungerReleaseNoBall()     'Plunger release sound when there is no ball in shooter lane
    End If
    TimerVRPlunger.Enabled = False
    TimerVRPlunger1.Enabled = True
    VR_Shooter.Y = -4
  End If

  If keycode = LeftFlipperKey Then
    VRFlipperButtonLeft.X = 2112
  End If

  If keycode = RightFlipperKey Then
    VRFlipperButtonRight.X = 2094
  End If

  If keycode = StartGameKey Then
    StartButton.y = 1916
  End If

  If KeyUpHandler(keycode) Then Exit Sub
End Sub


'******************************************************
' ZSOL: Solenoids & Flashers
'******************************************************

SolCallback(6)  = "SolKnocker"        ' Knocker solenoid
SolCallback(7)  = "SolBallRelease"      ' Trough to Plunger Lane
SolCallback(8)  = "SolInlineDTReset"    ' Inline Targets
SolCallback(9)  = "SolDTReset"        ' 3 Drop Target Reset
SolCallback(10) = "SolSaucer"       ' Saucer
SolCallback(17) = "SolGate"         ' Top Right Gate


' *** UPDATE

'SolCallback(19) =  "vpmNudge.SolGameOn"        'sEnable


SolCallback(sLRFlipper) = "SolRFlipper" 'Right Flipper
SolCallback(sLLFlipper) = "SolLFlipper" 'Left Flipper


'******************************************************
' ZDRN: Drain, Trough, and Ball Release
'******************************************************

'********************* TROUGH *************************

Sub BallRelease_Hit   : Controller.Switch(8) = 1 : End Sub
Sub BallRelease_UnHit : RandomSoundBallRelease BallRelease : Controller.Switch(8) = 0 : End Sub


Sub UpdateTrough
  UpdateTroughTimer.Interval = 300
  UpdateTroughTimer.Enabled = 1
End Sub

Sub UpdateTroughTimer_Timer
  If BallRelease.BallCntOver = 0 Then Drain.kick 60, 16
  Me.Enabled = 0
End Sub


'*****************  DRAIN & RELEASE  ******************

Dim RNDKickValue1, RNDKickAngle1  'Random Values for saucer kick and angles

Sub SolBallRelease(enabled)
  If enabled Then
    BallRelease.kick 57, 10
    RandomSoundBallRelease BallRelease
  End If
End Sub

Sub Drain_Hit()
  RandomSoundDrain Drain
  gate4soundrdy = 0 ' used for gate4 metal sound when hit after first hit.
  UpdateTrough
End Sub


'**********************  SAUCER  ***********************
dim RKick

Sub sw12_Hit
  SoundSaucerLock
  Controller.Switch(12) = 1
  kickarmtop_prim.RotX = 6
End Sub

Sub SolSaucer(Enabled)
  If Enabled Then
    If sw12.ballcntover > 0 then
      RNDKickAngle1 = -14
      RNDKickValue1 = RndInt(23,25)     ' Generate random value variance of +/- 1 from 24
      sw12.kick RNDKickAngle1, RNDKickValue1
      sw12.timerenabled = 1
      RKick = 0
    End If
  End If
End Sub


Sub sw12_Timer()
    Select Case RKick
        Case 1:kickarmtop_prim.RotX = -20
        Case 2:kickarmtop_prim.RotX = -20
        Case 3:kickarmtop_prim.RotX = -20
        Case 4:kickarmtop_prim.RotX = -20
        Case 5:kickarmtop_prim.RotX = -20
        Case 6:kickarmtop_prim.RotX = -20
        Case 7:kickarmtop_prim.RotX = -20
        Case 8:kickarmtop_prim.RotX = -20
        Case 9:kickarmtop_prim.RotX = -20
        Case 10:kickarmtop_prim.RotX = -20
        Case 11:kickarmtop_prim.RotX = -10
        Case 12:kickarmtop_prim.RotX = -5
        Case 13:kickarmtop_prim.RotX = 0:sw12.timerenabled = 0
    End Select
    RKick = RKick + 1
End Sub

Sub sw12_Unhit
  SoundSaucerKick 1, sw12
  controller.switch(12) = 0
End Sub


'*******************************************
'  Block Reflections over saucer
'*******************************************

sub BallReflectionMask_Hit() : table1.BallPlayfieldReflectionScale = 0 : end Sub
sub BallReflectionMask_UnHit() : table1.BallPlayfieldReflectionScale = 1 : end Sub


'******************************************************
' ZFLP: FLIPPERS
'******************************************************
Const ReflipAngle = 20

Sub SolLFlipper(Enabled)
  If Enabled Then
    FlipperActivate LeftFlipper, LFPress
    LF.Fire
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

Sub SolRFlipper(Enabled)
  If Enabled Then
    FlipperActivate RightFlipper, RFPress
    RF.Fire
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


'Flipper collide subs
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



'************************************************************
' ZSLG: Slingshot Animations
'************************************************************

Dim RStep, LStep

Sub RightSlingShot_Slingshot
  RS.VelocityCorrect(ActiveBall)
  vpmTimer.PulseSw 36
  RandomSoundSlingshotRight SLING1
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.rotx = 12
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.rotx = 10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.rotx = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub


Sub LeftSlingShot_Slingshot
  LS.VelocityCorrect(ActiveBall)
  vpmTimer.PulseSw 37
  RandomSoundSlingshotLeft SLING2
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.rotx = 12
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.rotx = 10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.rotx = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub


'******************************************************
' ZSSC: SLINGSHOT CORRECTION FUNCTIONS by apophis
'******************************************************
' To add these slingshot corrections:
'  - On the table, add the endpoint primitives that define the two ends of the Slingshot
'  - Initialize the SlingshotCorrection objects in InitSlingCorrection
'  - Call the .VelocityCorrect methods from the respective _Slingshot event sub

Dim LS: Set LS = New SlingshotCorrection
Dim RS: Set RS = New SlingshotCorrection

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
  a = Array(RS)
  Dim x
  For Each x In a
    x.addpoint idx, aX, aY
  Next
End Sub


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


'************************************************************
' ZSWI: SWITCHES
'************************************************************


'*******************************************
'Spinners
'*******************************************

dim spindamping: spindamping = 0.9925
dim ntimelastspin, ntimecurrentspin, ntimedifspin

Sub sw27trig_hit()
    'debug.print abs(gBOT(0).vely)
  ntimelastspin = gametime

  If abs(gBOT(0).vely) < 10 Then
    spindamping = 0.9925
  Else
    spindamping = .995
    spindampingtimer.enabled = 1
  End If
End Sub

Sub spindampingtimer_timer()
  If spindamping > .994 Then
    spindamping = spindamping - 0.00015
  Else
    spindamping = .96
    spindampingtimer.enabled = 0
  End If
End Sub

Sub sw27_Spin()
  ntimecurrentspin = gametime
  ntimedifspin = ntimecurrentspin - ntimelastspin
'   debug.print (ntimedifspin)

  sw27.damping = spindamping

  If ntimedifspin < 750 Then
    vpmtimer.PulseSw 27
    SoundSpinner sw27
  End If

  ntimelastspin = ntimecurrentspin
End Sub


'***********Rotate Spinner
Dim SpinnerRadius: SpinnerRadius= 7

Sub sw27_Animate
  Dim a
  a = sw27.CurrentAngle
  SpinnerPrim.Rotx = -a
  SpinnerRod.TransY = sin((a) * (PI/180)-0.25) * -SpinnerRadius
  SpinnerRod.TransZ = (cos((a + 180) * (PI/180))+1) * SpinnerRadius

End Sub


'************************* Bumpers **************************
Sub Bumper1_Hit(): RandomSoundBumperTop Bumper1: vpmTimer.PulseSw 40: End Sub
Sub Bumper2_Hit(): RandomSoundBumperTop Bumper2: vpmTimer.PulseSw 39: End Sub
Sub Bumper3_Hit(): RandomSoundBumperBottom Bumper3: vpmTimer.PulseSw 38: End Sub


'********************* Leaf Sensors **********************

Sub Phys_Sensor_001_hit: vpmTimer.PulseSw 1: End Sub
Sub Phys_Sensor_002_hit: vpmTimer.PulseSw 1: End Sub
Sub Phys_Sensor_003_hit: vpmTimer.PulseSw 1: End Sub
Sub Phys_Sensor_004_hit: vpmTimer.PulseSw 1: End Sub

'************************ Rollovers *************************

Sub swBIPL_Hit(): BIPL = 1: RampLShooterLane.collidable = 1 : End Sub 'Plunger Lane Rollover Switch
Sub swBIPL_UnHit: BIPL = 0: End Sub

Sub sw2_Hit(): Controller.Switch(2) = 1: End Sub
Sub sw2_UnHit: Controller.Switch(2) = 0: End Sub
Sub sw3_Hit(): Controller.Switch(3) = 1: End Sub
Sub sw3_UnHit: Controller.Switch(3) = 0: End Sub
Sub sw4_Hit(): Controller.Switch(4) = 1: End Sub
Sub sw4_UnHit: Controller.Switch(4) = 0: End Sub
Sub sw5_Hit(): Controller.Switch(5) = 1: End Sub
Sub sw5_UnHit: Controller.Switch(5) = 0: End Sub
Sub sw13_Hit(): Controller.Switch(13) = 1: End Sub
Sub sw13_UnHit: Controller.Switch(13) = 0: End Sub
Sub sw14_Hit(): Controller.Switch(14) = 1: End Sub
Sub sw14_UnHit: Controller.Switch(14) = 0: End Sub
Sub sw15_Hit(): Controller.Switch(15) = 1: End Sub
Sub sw15_UnHit: Controller.Switch(15) = 0: End Sub


'*******************************************
' Rollover Animations
'*******************************************

Sub sw2_Animate:  psw2.transz = sw2.CurrentAnimOffset: End Sub
Sub sw3_Animate:  psw3.transz = sw3.CurrentAnimOffset: End Sub
Sub sw4_Animate:  psw4.transz = sw4.CurrentAnimOffset: End Sub
Sub sw5_Animate:  psw5.transz = sw5.CurrentAnimOffset: End Sub
Sub sw13_Animate: psw13.transz = sw13.CurrentAnimOffset: End Sub
Sub sw14_Animate: psw14.transz = sw14.CurrentAnimOffset: End Sub
Sub sw15_Animate: psw15.transz = sw15.CurrentAnimOffset: End Sub


'*******************************************
'Star Triggers
'*******************************************

Sub sw25_Hit:vpmTimer.PulseSw 25:me.timerenabled = 0:AnimateStar star25, sw25, 1:End Sub
Sub sw25_UnHit:me.timerinterval = 7:me.timerenabled = 1:End Sub
Sub sw25_timer:AnimateStar star25, sw25, 0:End Sub

Sub AnimateStar(prim, sw, action) ' Action = 1 - to drop, 0 to raise
  If action = 1 Then
    prim.transz = -4
  Else
    prim.transz = prim.transz + 0.5
    if prim.transz = -2 and Rnd() < 0.05 Then
      sw.timerenabled = 0
    Elseif prim.transz >= 0 Then
      prim.transz = 0
      sw.timerenabled = 0
    End If
  End If
End Sub


'*******************************************
'  Knocker
'*******************************************

Sub SolKnocker(Enabled)
  If enabled Then
    KnockerSolenoid
  End If
End Sub


'*******************************************
'  Gate Solenoid to open gate under hood
'*******************************************

Sub SolGate(Enabled)
  Gate_003.Open = Enabled
End Sub


'*******************************************
' Gate Metal Hit sounds when gates are closed
' Note: Gate_004_Trig_sound_rdy turns off when drain is hit or at initial Start
'*******************************************

Sub Gate_003_Trig_sound_hit() ' doesn't need a sound trigger to enable, as it starts in closed position
  If Gate_003.open = false then
    RandomSoundMetal
  End If
End Sub

Sub Gate_004_Trig_sound_rdy_hit()
  gate4soundrdy = 1
End Sub

Sub Gate_004_Trig_sound_hit()
  If Gate_004.open = false then
    if gate4soundrdy =1 Then
      RandomSoundMetal
    End If
  End If
End Sub


'********************* Standup Targets **********************

Sub sw24_hit: STHit 24: End Sub
Sub sw29_hit: STHit 29: End Sub
Sub sw30_hit: STHit 30: End Sub
Sub sw31_hit: STHit 31: End Sub

Sub sw24o_Hit: TargetBouncer ActiveBall, 1: End Sub
Sub sw29o_Hit: TargetBouncer ActiveBall, 1: End Sub
Sub sw30o_Hit: TargetBouncer ActiveBall, 0.5: End Sub
Sub sw31o_Hit: TargetBouncer ActiveBall, 1: End Sub


'********************** Drop Targets ************************
Sub sw21_Hit: DTMass = DTMassWeak: DTHit 21: TargetBouncerIaakki Activeball, 1: DT21up = 0: End Sub
Sub sw22_Hit: DTMass = DTMassWeak: DTHit 22: TargetBouncerIaakki Activeball, 1: DT22up = 0: End Sub
Sub sw23_Hit: DTMass = DTMassWeak: DTHit 23: TargetBouncerIaakki Activeball, 1: DT23up = 0: End Sub
Sub sw33_Hit: DTMass = DTMassStrong: DTHit 33: TargetBouncer Activeball, 0.5: DT33up = 0: End Sub
Sub sw34_Hit: DTMass = DTMassStrong: DTHit 34: TargetBouncer Activeball, 0.5: DT34up = 0: End Sub
Sub sw35_Hit: DTMass = DTMassStrong: DTHit 35: TargetBouncer Activeball, 0.5: DT35up = 0: End Sub


'********************************************
' Drop Target Solenoid Controls
'********************************************


Sub SolDTReset (enabled)
  If enabled then
    RandomSoundDropTargetReset sw34p
    DTRaise 33
    DTRaise 34
    DTRaise 35
    DT33up=1
    DT34up=1
    DT35up=1
    dtsh33.visible=True
    dtsh34.visible=True
    dtsh35.visible=True
  End if
End Sub

Sub SolInlineDTReset (enabled)
  If enabled then
    RandomSoundDropTargetReset sw22p
    DTRaise 21
    DTRaise 22
    DTRaise 23
    DT21up=1
    DT22up=1
    DT23up=1
    dtsh21.visible=True
    dtsh22.visible=True
    dtsh23.visible=True
  End if
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
  Dim b'

  ' stop the sound of deleted balls
  For b = UBound(gBOT) + 1 to tnob - 1
    rolling(b) = False
    StopSound("BallRoll_" & b)
  Next

  ' exit the sub if no balls on the table
  If UBound(gBOT) = -1 Then Exit Sub

  ' play the rolling sound for each ball

  For b = 0 to UBound(gBOT)
    If BallVel(gBOT(b)) > 1 AND gBOT(b).z < 30 Then
      rolling(b) = True
      PlaySound ("BallRoll_" & b), -1, VolPlayfieldRoll(gBOT(b)) * BallRollVolume * VolumeDial, AudioPan(gBOT(b)), 0, PitchPlayfieldRoll(gBOT(b)), 1, 0, AudioFade(gBOT(b))

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
  Next
End Sub


'******************************************************
'****  END BALL ROLLING AND DROP SOUNDS
'******************************************************


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

'/////////////////////////////  GI AND FLASHER RELAYS  ////////////////////////////

Const RelayFlashSoundLevel = 0.315  'volume level; range [0, 1];
Const RelayGISoundLevel = 1   'volume level; range [0, 1];

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
' ZNFF:  FLIPPER CORRECTIONS by nFozzy
'******************************************************
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

Sub InitPolarity()
   dim x, a : a = Array(LF, RF)
  for each x in a
    x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
    x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
    x.enabled = True
    x.TimeDelay = 80
    x.DebugOn=False ' prints some info in debugger


        x.AddPt "Polarity", 0, 0, 0
        x.AddPt "Polarity", 1, 0.05, - 2.7
        x.AddPt "Polarity", 2, 0.16, - 2.7
        x.AddPt "Polarity", 3, 0.22, - 0
        x.AddPt "Polarity", 4, 0.25, - 0
        x.AddPt "Polarity", 5, 0.3, - 1
        x.AddPt "Polarity", 6, 0.4, - 2
        x.AddPt "Polarity", 7, 0.5, - 2.7
        x.AddPt "Polarity", 8, 0.65, - 1.8
        x.AddPt "Polarity", 9, 0.75, - 0.5
        x.AddPt "Polarity", 10, 0.81, - 0.5
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

Sub FlipperNudge(Flipper1, Endangle1, EOSNudge1, Flipper2, EndAngle2)
  Dim b

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

Const LiveCatch = 12
Const LiveElasticity = 0.45
Const SOSEM = 0.815
'   Const EOSReturn = 0.055  'EM's
   Const EOSReturn = 0.045  'late 70's to mid 80's
' Const EOSReturn = 0.035  'mid 80's to early 90's
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
    Dim b

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
'   aBall.velz = aBall.velz * coef
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



'******************************************************
'****  END PHYSICS DAMPENERS
'******************************************************


'******************************************************
'   ZRDT:  DROP TARGETS by Rothbauerw
'******************************************************


'******************************************************
'  DROP TARGETS INITIALIZATION
'******************************************************

Class DropTarget
  Private m_primary, m_secondary, m_prim, m_sw, m_animate, m_isDropped

  Public Property Get Primary(): Set Primary = m_primary: End Property
  Public Property Let Primary(input): Set m_primary = input: End Property

  Public Property Get Secondary(): Set Secondary = m_secondary: End Property
  Public Property Let Secondary(input): Set m_secondary = input: End Property

  Public Property Get Prim(): Set Prim = m_prim: End Property
  Public Property Let Prim(input): Set m_prim = input: End Property

  Public Property Get Sw(): Sw = m_sw: End Property
  Public Property Let Sw(input): m_sw = input: End Property

  Public Property Get Animate(): Animate = m_animate: End Property
  Public Property Let Animate(input): m_animate = input: End Property

  Public Property Get IsDropped(): IsDropped = m_isDropped: End Property
  Public Property Let IsDropped(input): m_isDropped = input: End Property

  Public default Function init(primary, secondary, prim, sw, animate, isDropped)
    Set m_primary = primary
    Set m_secondary = secondary
    Set m_prim = prim
    m_sw = sw
    m_animate = animate
    m_isDropped = isDropped

    Set Init = Me
  End Function
End Class


'Define a variable for each drop target
Dim DT21, DT22, DT23, DT33, DT34, DT35

'Set array with drop target objects
'
'DropTargetvar = Array(primary, secondary, prim, swtich, animate)
'   primary:      primary target wall to determine drop
' secondary:      wall used to simulate the ball striking a bent or offset target after the initial Hit
' prim:       primitive target used for visuals and animation
'             IMPORTANT!!!
'             rotz must be used for orientation
'             rotx to bend the target back
'             transz to move it up and down
'             the pivot point should be in the center of the target on the x, y and at or below the playfield (0) on z
' switch:       ROM switch number
' animate:      Array slot for handling the animation instrucitons, set to 0
'           Values for animate: 1 - bend target (hit to primary), 2 - drop target (hit to secondary), 3 - brick target (high velocity hit to secondary), -1 - raise target
'   isDropped:      Boolean which determines whether a drop target is dropped. Set to false if they are initially raised, true if initially dropped.

Set DT21 = (new DropTarget)(sw21, sw21a, sw21p, 21, 0, false)
Set DT22 = (new DropTarget)(sw22, sw22a, sw22p, 22, 0, false)
Set DT23 = (new DropTarget)(sw23, sw23a, sw23p, 23, 0, false)
Set DT33 = (new DropTarget)(sw33, sw33a, sw33p, 33, 0, false)
Set DT34 = (new DropTarget)(sw34, sw34a, sw34p, 34, 0, false)
Set DT35 = (new DropTarget)(sw35, sw35a, sw35p, 35, 0, false)

Dim DTArray
DTArray = Array(DT21, DT22, DT23, DT33, DT34, DT35)

'Configure the behavior of Drop Targets.
Const DTDropSpeed = 90 'in milliseconds
Const DTDropUpSpeed = 40 'in milliseconds
Const DTDropUnits = 44 'VP units primitive drops so top of at or below the playfield
Const DTDropUpUnits = 10 'VP units primitive raises above the up position on drops up
Const DTMaxBend = 2 '8 'max degrees primitive rotates when hit
Const DTDropDelay = 20 'time in milliseconds before target drops (due to friction/impact of the ball)
Const DTRaiseDelay = 40 'time in milliseconds before target drops back to normal up position after the solenoid fires to raise the target
Const DTBrickVel = 30 'velocity at which the target will brick, set to '0' to disable brick
Const DTEnableBrick = 0 '0 'Set to 0 to disable bricking, 1 to enable bricking

Const DTHitSound = "" 'Drop Target Hit sound
Const DTDropSound = "DropTarget_Down" 'Drop Target Drop sound
Const DTResetSound = "DropTarget_Up" 'Drop Target reset sound

Dim DTMass: DTMass = 0.2 'initial value
Const DTMassWeak = 0.02 'Mass of the Drop Target (between 0 and 1), higher values provide more resistance
Const DTMassStrong = 0.2 'Mass of the Drop Target (between 0 and 1), higher values provide more resistance


'******************************************************
'  DROP TARGETS FUNCTIONS
'******************************************************

Dim DTShadow(6)

DTShadowInit 1
DTShadowInit 2
DTShadowInit 3
DTShadowInit 4
DTShadowInit 5
DTShadowInit 6

' Initializes the drop targets for shadow logic below
Sub DTShadowInit(dtnbr)

  if dtnbr = 1 Then
    Set DTShadow(dtnbr) = Eval("dtsh" & 21)
  elseif dtnbr = 2 Then
    Set DTShadow(dtnbr) = Eval("dtsh" & 22)
  elseif dtnbr = 3 Then
    Set DTShadow(dtnbr) = Eval("dtsh" & 23)
  elseif dtnbr = 4 Then
    Set DTShadow(dtnbr) = Eval("dtsh" & 33)
  elseif dtnbr = 5 Then
    Set DTShadow(dtnbr) = Eval("dtsh" & 34)
  elseif dtnbr = 6 Then
    Set DTShadow(dtnbr) = Eval("dtsh" & 35)
  End If
End Sub

Sub DTHit(switch)
  Dim i, swmod

  i = DTArrayID(switch)
  If switch = 21 Then
    swmod = 1
  Elseif switch = 22 then
    swmod = 2
  Elseif switch = 23 then
    swmod = 3
  Elseif switch = 33 then
    swmod = 4
  Elseif switch = 34 then
    swmod = 5
  Elseif switch = 35 then
    swmod = 6
  End If
  i = DTArrayID(switch)

  PlayTargetSound
  DTArray(i).animate = DTCheckBrick(ActiveBall,DTArray(i).prim)
  If DTArray(i).animate = 1 Or DTArray(i).animate = 3 Or DTArray(i).animate = 4 Then
    DTBallPhysics ActiveBall, DTArray(i).prim.rotz, DTMass
' Controls Drop Shadow for a direct hit only
    if swmod=1 or swmod=2 or swmod=3 or swmod=4 or swmod=5 or swmod=6 then
      DTShadow(swmod).visible = 0
    End If
  End If
  DoDTAnim
End Sub

Sub DTRaise(switch)
  Dim i
  i = DTArrayID(switch)

  DTArray(i).animate =  - 1
  DoDTAnim
End Sub

Sub DTDrop(switch)
  Dim i
  i = DTArrayID(switch)

  DTArray(i).animate = 1
  DoDTAnim
End Sub

Function DTArrayID(switch)
  Dim i
  For i = 0 To UBound(DTArray)
    If DTArray(i).sw = switch Then
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
    DTArray(i).animate = DTAnimate(DTArray(i).primary,DTArray(i).secondary,DTArray(i).prim,DTArray(i).sw,DTArray(i).animate)
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
    primary.uservalue = GameTime
  End If

  animtime = GameTime - primary.uservalue

  If (animate = 1 Or animate = 4) And animtime < DTDropDelay Then
    primary.collidable = 0
    If animate = 1 Then secondary.collidable = 1 Else secondary.collidable = 0
    prim.rotx = DTMaxBend * Cos(rangle)
    prim.roty = DTMaxBend * Sin(rangle)
    DTAnimate = animate
    Exit Function
  ElseIf (animate = 1 Or animate = 4) And animtime > DTDropDelay Then
    primary.collidable = 0
    If animate = 1 Then secondary.collidable = 1 Else secondary.collidable = 1 'If animate = 1 Then secondary.collidable = 1 Else secondary.collidable = 0 'updated by rothbauerw to account for edge case
    prim.rotx = DTMaxBend * Cos(rangle)
    prim.roty = DTMaxBend * Sin(rangle)
    animate = 2
    SoundDropTargetDrop prim
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
      DTArray(ind).isDropped = True 'Mark target as dropped
      controller.Switch(Switchid mod 100) = 1
      prim.reflectionenabled = False  'Turn drop target reflection on when above playfield
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
      For b = 0 To UBound(gBOT)
        If InRotRect(gBOT(b).x,gBOT(b).y,prim.x, prim.y, prim.rotz, - 25, - 10,25, - 10,25,25, - 25,25) And gBOT(b).z < prim.z + DTDropUnits + 25 Then
          gBOT(b).velz = 20
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
      primary.uservalue = GameTime
    End If
    primary.collidable = 0
    secondary.collidable = 1
    DTArray(ind).isDropped = False 'Mark target as not dropped
    controller.Switch(Switchid mod 100) = 0
    prim.reflectionenabled = True 'Turn drop target reflection off when below playfield
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

Function DTDropped(switchid)
  Dim ind
  ind = DTArrayID(switchid)

  DTDropped = DTArray(ind).isDropped
End Function


'******************************************************
'****  END DROP TARGETS
'******************************************************


'******************************************************
'****  END DROP TARGETS
'******************************************************


'******************************************************
' ZRST: STAND-UP TARGET INITIALIZATION
'******************************************************

Class StandupTarget
  Private m_primary, m_prim, m_sw, m_animate

  Public Property Get Primary(): Set Primary = m_primary: End Property
  Public Property Let Primary(input): Set m_primary = input: End Property

  Public Property Get Prim(): Set Prim = m_prim: End Property
  Public Property Let Prim(input): Set m_prim = input: End Property

  Public Property Get Sw(): Sw = m_sw: End Property
  Public Property Let Sw(input): m_sw = input: End Property

  Public Property Get Animate(): Animate = m_animate: End Property
  Public Property Let Animate(input): m_animate = input: End Property

  Public default Function init(primary, prim, sw, animate)
    Set m_primary = primary
    Set m_prim = prim
    m_sw = sw
    m_animate = animate

    Set Init = Me
  End Function
End Class


'Define a variable for each stand-up target
' NOTE: code modified to work with Swinging Target in the center

Dim ST24,ST29,ST30,ST31


'StandupTargetvar = Array(primary, prim, swtich)
'   primary:      vp target to determine target hit
' prim:       primitive target used for visuals and animation
'             IMPORTANT!!!
'             transy must be used to offset the target animation
' switch:       ROM switch number
' animate:      Arrary slot for handling the animation instrucitons, set to 0
'
'You will also need to add a secondary hit object for each stand up (name sw11o, sw12o, and sw13o on the example Table1)
'these are inclined primitives to simulate hitting a bent target and should provide so z velocity on high speed impacts

Set ST24 = (new StandupTarget)(sw24, sw24p, 24, 0)
Set ST29 = (new StandupTarget)(sw29, sw29p, 29, 0)
Set ST30 = (new StandupTarget)(sw30, sw30p, 30, 0)
Set ST31 = (new StandupTarget)(sw31, sw31p, 31, 0)


'Add all the Stand-up Target Arrays to Stand-up Target Animation Array
'   STAnimationArray = Array(ST1, ST2, ....)
Dim STArray
STArray = Array(ST24,ST29,ST30,ST31)

'Configure the behavior of Stand-up Targets
Const STAnimStep =  1.5     'vpunits per animation step (control return to Start)
Const STMaxOffset = 9       'max vp units target moves when hit

Const STMass = 0.2        'Mass of the Stand-up Target (between 0 and 1), higher values provide more resistance

'******************************************************
'       STAND-UP TARGETS FUNCTIONS
'******************************************************

Sub STHit(switch)
  Dim i
  i = STArrayID(switch)

  PlayTargetSound
  STArray(i).animate = STCheckHit(ActiveBall,STArray(i).primary)

  If STArray(i).animate <> 0 Then
    DTBallPhysics ActiveBall, STArray(i).primary.orientation, STMass
  End If
  DoSTAnim
End Sub

Function STArrayID(switch)
  Dim i
  For i = 0 To UBound(STArray)
    If STArray(i).sw = switch Then
      STArrayID = i
      Exit Function
    End If
  Next
End Function

Function STCheckHit(aBall, target) 'Check if target is hit on it's face
  Dim bangle, bangleafter, rangle, rangle2, perpvel, perpvelafter, paravel, paravelafter
  rangle = (target.orientation - 90) * 3.1416 / 180
  bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))
  bangleafter = Atn2(aBall.vely,aball.velx)

  perpvel = cor.BallVel(aball.id) * Cos(bangle - rangle)
  paravel = cor.BallVel(aball.id) * Sin(bangle - rangle)

  perpvelafter = BallSpeed(aBall) * Cos(bangleafter - rangle)
  paravelafter = BallSpeed(aBall) * Sin(bangleafter - rangle)

  If perpvel > 0 And  perpvelafter <= 0 Then
    STCheckHit = 1
  ElseIf perpvel > 0 And ((paravel > 0 And paravelafter > 0) Or (paravel < 0 And paravelafter < 0)) Then
    STCheckHit = 1
  Else
    STCheckHit = 0
  End If
End Function

Sub DoSTAnim()
  Dim i
  For i = 0 To UBound(STArray)
    STArray(i).animate = STAnimate(STArray(i).primary,STArray(i).prim,STArray(i).sw,STArray(i).animate)
  Next
End Sub

Function STAnimate(primary, prim, switch,  animate)
  Dim animtime

  STAnimate = animate

  If animate = 0  Then
    primary.uservalue = 0
    STAnimate = 0
    Exit Function
  ElseIf primary.uservalue = 0 Then
    primary.uservalue = GameTime
  End If

  animtime = GameTime - primary.uservalue

  If animate = 1 Then
    primary.collidable = 0
    prim.transy =  - STMaxOffset
    vpmTimer.PulseSw switch mod 100
    STAnimate = 2
    Exit Function
  ElseIf animate = 2 Then
    prim.transy = prim.transy + STAnimStep
    If prim.transy >= 0 Then
      prim.transy = 0
      primary.collidable = 1
      STAnimate = 0
      Exit Function
    Else
      STAnimate = 2
    End If
  End If
End Function


'******************************************************
'***  END STAND-UP TARGETS
'******************************************************



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


sub TargetBouncerIaakki(aBall,defvalue)
  Dim zMultiplier, vel
  vel = BallSpeed(aBall)
' debug.print "bounce"
  if TargetBouncerEnabled = 1 and aball.z < 30 then
    'debug.print "velz: " & activeball.velz
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


'****************************************************************
'****  END VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'****************************************************************


'******************************************************
'   Z3DI:   3D INSERTS
'******************************************************

dim insert_white_on: insert_white_on = insertfactor * 30
const insert_white_off = 0.5
dim insert_red_on: insert_red_on = insertfactor * 22
const insert_red_off = 4
dim insert_redlight_on: insert_redlight_on = insertfactor * 20
const insert_redlight_off = 4
dim insert_orange_on: insert_orange_on = insertfactor * 16
const insert_orange_off = 1
dim insert_yellow_on: insert_yellow_on = insertfactor * 18
const insert_yellow_off = 0.5
dim insert_green_on: insert_green_on = insertfactor * 24
const insert_green_off = 0.25


const a_insert_white_on = .15
const a_insert_white_off = .35
const a_insert_red_on = 0.1
const a_insert_red_off = 0.15
const a_insert_redlight_on = 0.1
const a_insert_redlight_off = 0.15
const a_insert_orange_on = 0.1
const a_insert_orange_off = 0.105
const a_insert_yellow_on = 1
const a_insert_yellow_off = 1
const a_insert_green_on = 1
const a_insert_green_off = 1


const bulb_on = 10
const bulb_off = 1

const insert_credit_on = 0.1
const insert_credit_off = 0.05



'************************
'**** White Inserts

Sub l7_animate
  p7.BlendDisableLighting = insert_white_off + insert_white_on * (l7.GetInPlayIntensity / l7.Intensity)
  p7off.BlendDisableLighting = a_insert_white_off - a_insert_white_on * (l7.GetInPlayIntensity / l7.Intensity)
  bulb7.BlendDisableLighting = bulb_off + bulb_on * (l7.GetInPlayIntensity / l7.Intensity)
End Sub

Sub l8_animate
  p8.BlendDisableLighting = insert_white_off + insert_white_on * (l8.GetInPlayIntensity / l8.Intensity)
  p8off.BlendDisableLighting = a_insert_white_off - a_insert_white_on * (l8.GetInPlayIntensity / l8.Intensity)
  bulb8.BlendDisableLighting = bulb_off + bulb_on * (l8.GetInPlayIntensity / l8.Intensity)
End Sub

Sub l23_animate
  p23.BlendDisableLighting = insert_white_off + insert_white_on * (l23.GetInPlayIntensity / l23.Intensity)
  p23off.BlendDisableLighting = a_insert_white_off - a_insert_white_on * (l23.GetInPlayIntensity / l23.Intensity)
  bulb23.BlendDisableLighting = bulb_off + bulb_on * (l23.GetInPlayIntensity / l23.Intensity)
End Sub

Sub l24_animate
  p24.BlendDisableLighting = insert_white_off + insert_white_on * (l24.GetInPlayIntensity / l24.Intensity)
  p24off.BlendDisableLighting = a_insert_white_off - a_insert_white_on * (l24.GetInPlayIntensity / l24.Intensity)
  bulb24.BlendDisableLighting = bulb_off + bulb_on * (l24.GetInPlayIntensity / l24.Intensity)
End Sub

Sub l39_animate
  p39.BlendDisableLighting = insert_white_off + insert_white_on * (l39.GetInPlayIntensity / l39.Intensity)
  p39off.BlendDisableLighting = a_insert_white_off - a_insert_white_on * (l39.GetInPlayIntensity / l39.Intensity)
  bulb39.BlendDisableLighting = bulb_off + bulb_on * (l39.GetInPlayIntensity / l39.Intensity)
End Sub

Sub l40_animate
  p40.BlendDisableLighting = insert_white_off + insert_white_on * (l40.GetInPlayIntensity / l40.Intensity)
  p40off.BlendDisableLighting = a_insert_white_off - a_insert_white_on * (l40.GetInPlayIntensity / l40.Intensity)
  bulb40.BlendDisableLighting = bulb_off + bulb_on * (l40.GetInPlayIntensity / l40.Intensity)
  l40on = l40.GetInPlayIntensity / l40.Intensity
  SetRoomBrightness
End Sub

Sub l55_animate
  p55.BlendDisableLighting = insert_white_off + insert_white_on * (l55.GetInPlayIntensity / l55.Intensity)
  p55off.BlendDisableLighting = a_insert_white_off - a_insert_white_on * (l55.GetInPlayIntensity / l55.Intensity)
  bulb55.BlendDisableLighting = bulb_off + bulb_on * (l55.GetInPlayIntensity / l55.Intensity)
End Sub

Sub l56_animate
  p56.BlendDisableLighting = insert_white_off + insert_white_on * (l56.GetInPlayIntensity / l56.Intensity)
  p56off.BlendDisableLighting = a_insert_white_off - a_insert_white_on * (l56.GetInPlayIntensity / l56.Intensity)
  bulb56.BlendDisableLighting = bulb_off + bulb_on * (l56.GetInPlayIntensity / l56.Intensity)
End Sub


'************************
'**** White Inserts (small)

Sub l12_animate
  p12.BlendDisableLighting = insert_white_off + insert_white_on * (l12.GetInPlayIntensity / l12.Intensity)
  p12off.BlendDisableLighting = a_insert_white_off - a_insert_white_on * (l12.GetInPlayIntensity / l12.Intensity)
  bulb12.BlendDisableLighting = bulb_off + bulb_on * (l12.GetInPlayIntensity / l12.Intensity)
End Sub

Sub l28_animate
  p28.BlendDisableLighting = insert_white_off + insert_white_on * (l28.GetInPlayIntensity / l28.Intensity)
  p28off.BlendDisableLighting = a_insert_white_off - a_insert_white_on * (l28.GetInPlayIntensity / l28.Intensity)
  bulb28.BlendDisableLighting = bulb_off + bulb_on * (l28.GetInPlayIntensity / l28.Intensity)
End Sub

Sub l46_animate
  p46.BlendDisableLighting = insert_white_off + insert_white_on * (l46.GetInPlayIntensity / l46.Intensity)
  p46off.BlendDisableLighting = a_insert_white_off - a_insert_white_on * (l46.GetInPlayIntensity / l46.Intensity)
  bulb46.BlendDisableLighting = bulb_off + bulb_on * (l46.GetInPlayIntensity / l46.Intensity)
End Sub

Sub l52_animate
  p52.BlendDisableLighting = insert_white_off + insert_white_on * (l52.GetInPlayIntensity / l52.Intensity)
  p52off.BlendDisableLighting = a_insert_white_off - a_insert_white_on * (l52.GetInPlayIntensity / l52.Intensity)
  bulb52.BlendDisableLighting = bulb_off + bulb_on * (l52.GetInPlayIntensity / l52.Intensity)
End Sub

Sub l54_animate
  p54.BlendDisableLighting = insert_white_off + insert_white_on * (l54.GetInPlayIntensity / l54.Intensity)
  p54off.BlendDisableLighting = a_insert_white_off - a_insert_white_on * (l54.GetInPlayIntensity / l54.Intensity)
  bulb54.BlendDisableLighting = bulb_off + bulb_on * (l54.GetInPlayIntensity / l54.Intensity)
End Sub

Sub l63_animate
  p63.BlendDisableLighting = insert_white_off + insert_white_on * (l63.GetInPlayIntensity / l63.Intensity)
  p63off.BlendDisableLighting = a_insert_white_off - a_insert_white_on * (l63.GetInPlayIntensity / l63.Intensity)
  bulb63.BlendDisableLighting = bulb_off + bulb_on * (l63.GetInPlayIntensity / l63.Intensity)
End Sub


'************************
'**** White Inserts (animals)

Sub l1_animate
  p1.BlendDisableLighting = insert_white_off + insert_white_on * (l1.GetInPlayIntensity / l1.Intensity)
  p1off.BlendDisableLighting = a_insert_white_off - a_insert_white_on * (l1.GetInPlayIntensity / l1.Intensity)
  bulb1.BlendDisableLighting = bulb_off + bulb_on * (l1.GetInPlayIntensity / l1.Intensity)
End Sub

Sub l2_animate
  p2.BlendDisableLighting = insert_white_off + insert_white_on * (l2.GetInPlayIntensity / l2.Intensity)
  p2off.BlendDisableLighting = a_insert_white_off - a_insert_white_on * (l2.GetInPlayIntensity / l2.Intensity)
  bulb2.BlendDisableLighting = bulb_off + bulb_on * (l2.GetInPlayIntensity / l2.Intensity)
End Sub

Sub l3_animate
  p3.BlendDisableLighting = insert_white_off + insert_white_on * (l3.GetInPlayIntensity / l3.Intensity)
  p3off.BlendDisableLighting = a_insert_white_off - a_insert_white_on * (l3.GetInPlayIntensity / l3.Intensity)
  bulb3.BlendDisableLighting = bulb_off + bulb_on * (l3.GetInPlayIntensity / l3.Intensity)
End Sub

Sub l4_animate
  p4.BlendDisableLighting = insert_white_off + insert_white_on * (l4.GetInPlayIntensity / l4.Intensity)
  p4off.BlendDisableLighting = a_insert_white_off - a_insert_white_on * (l4.GetInPlayIntensity / l4.Intensity)
  bulb4.BlendDisableLighting = bulb_off + bulb_on * (l4.GetInPlayIntensity / l4.Intensity)
End Sub

Sub l5_animate
  p5.BlendDisableLighting = insert_white_off + insert_white_on * (l5.GetInPlayIntensity / l5.Intensity)
  p5off.BlendDisableLighting = a_insert_white_off - a_insert_white_on * (l5.GetInPlayIntensity / l5.Intensity)
  bulb5.BlendDisableLighting = bulb_off + bulb_on * (l5.GetInPlayIntensity / l5.Intensity)
End Sub

Sub l17_animate
  p17.BlendDisableLighting = insert_white_off + insert_white_on * (l17.GetInPlayIntensity / l17.Intensity)
  p17off.BlendDisableLighting = a_insert_white_off - a_insert_white_on * (l17.GetInPlayIntensity / l17.Intensity)
  bulb17.BlendDisableLighting = bulb_off + bulb_on * (l17.GetInPlayIntensity / l17.Intensity)
End Sub

Sub l18_animate
  p18.BlendDisableLighting = insert_white_off + insert_white_on * (l18.GetInPlayIntensity / l18.Intensity)
  p18off.BlendDisableLighting = a_insert_white_off - a_insert_white_on * (l18.GetInPlayIntensity / l18.Intensity)
  bulb18.BlendDisableLighting = bulb_off + bulb_on * (l18.GetInPlayIntensity / l18.Intensity)
End Sub

Sub l19_animate
  p19.BlendDisableLighting = insert_white_off + insert_white_on * (l19.GetInPlayIntensity / l19.Intensity)
  p19off.BlendDisableLighting = a_insert_white_off - a_insert_white_on * (l19.GetInPlayIntensity / l19.Intensity)
  bulb19.BlendDisableLighting = bulb_off + bulb_on * (l19.GetInPlayIntensity / l19.Intensity)
End Sub

Sub l20_animate
  p20.BlendDisableLighting = insert_white_off + insert_white_on * (l20.GetInPlayIntensity / l20.Intensity)
  p20off.BlendDisableLighting = a_insert_white_off - a_insert_white_on * (l20.GetInPlayIntensity / l20.Intensity)
  bulb20.BlendDisableLighting = bulb_off + bulb_on * (l20.GetInPlayIntensity / l20.Intensity)
End Sub

Sub l21_animate
  p21.BlendDisableLighting = insert_white_off + insert_white_on * (l21.GetInPlayIntensity / l21.Intensity)
  p21off.BlendDisableLighting = a_insert_white_off - a_insert_white_on * (l21.GetInPlayIntensity / l21.Intensity)
  bulb21.BlendDisableLighting = bulb_off + bulb_on * (l21.GetInPlayIntensity / l21.Intensity)
End Sub

Sub l33_animate
  p33.BlendDisableLighting = insert_white_off + insert_white_on * (l33.GetInPlayIntensity / l33.Intensity)
  p33off.BlendDisableLighting = a_insert_white_off - a_insert_white_on * (l33.GetInPlayIntensity / l33.Intensity)
  bulb33.BlendDisableLighting = bulb_off + bulb_on * (l33.GetInPlayIntensity / l33.Intensity)
End Sub

Sub l34_animate
  p34.BlendDisableLighting = insert_white_off + insert_white_on * (l34.GetInPlayIntensity / l34.Intensity)
  p34off.BlendDisableLighting = a_insert_white_off - a_insert_white_on * (l34.GetInPlayIntensity / l34.Intensity)
  bulb34.BlendDisableLighting = bulb_off + bulb_on * (l34.GetInPlayIntensity / l34.Intensity)
End Sub

Sub l35_animate
  p35.BlendDisableLighting = insert_white_off + insert_white_on * (l35.GetInPlayIntensity / l35.Intensity)
  p35off.BlendDisableLighting = a_insert_white_off - a_insert_white_on * (l35.GetInPlayIntensity / l35.Intensity)
  bulb35.BlendDisableLighting = bulb_off + bulb_on * (l35.GetInPlayIntensity / l35.Intensity)
End Sub

Sub l36_animate
  p36.BlendDisableLighting = insert_white_off + insert_white_on * (l36.GetInPlayIntensity / l36.Intensity)
  p36off.BlendDisableLighting = a_insert_white_off - a_insert_white_on * (l36.GetInPlayIntensity / l36.Intensity)
  bulb36.BlendDisableLighting = bulb_off + bulb_on * (l36.GetInPlayIntensity / l36.Intensity)
End Sub

Sub l37_animate
  p37.BlendDisableLighting = insert_white_off + insert_white_on * (l37.GetInPlayIntensity / l37.Intensity)
  p37off.BlendDisableLighting = a_insert_white_off - a_insert_white_on * (l37.GetInPlayIntensity / l37.Intensity)
  bulb37.BlendDisableLighting = bulb_off + bulb_on * (l37.GetInPlayIntensity / l37.Intensity)
End Sub


'************************
'**** Red Inserts (lighter / normal red ones)

Sub l53_animate
  p53.BlendDisableLighting = insert_redlight_off + insert_redlight_on * (l53.GetInPlayIntensity / l53.Intensity)
  p53off.BlendDisableLighting = a_insert_redlight_off - a_insert_redlight_on * (l53.GetInPlayIntensity / l53.Intensity)
  bulb53.BlendDisableLighting = bulb_off + bulb_on * (l53.GetInPlayIntensity / l53.Intensity)
End Sub


'************************
'**** Red Inserts (smaller ones)

Sub l44_animate
  p44.BlendDisableLighting = insert_redlight_off + insert_redlight_on * (l44.GetInPlayIntensity / l44.Intensity)
  p44off.BlendDisableLighting = a_insert_redlight_off - a_insert_redlight_on * (l44.GetInPlayIntensity / l44.Intensity)
  bulb44.BlendDisableLighting = bulb_off + bulb_on * (l44.GetInPlayIntensity / l44.Intensity)
End Sub

Sub l49_animate
  p49.BlendDisableLighting = insert_redlight_off + insert_redlight_on * (l49.GetInPlayIntensity / l49.Intensity)
  p49off.BlendDisableLighting = a_insert_redlight_off - a_insert_redlight_on * (l49.GetInPlayIntensity / l49.Intensity)
  bulb49.BlendDisableLighting = bulb_off + bulb_on * (l49.GetInPlayIntensity / l49.Intensity)
End Sub

Sub l50_animate
  p50.BlendDisableLighting = insert_redlight_off + insert_redlight_on * (l50.GetInPlayIntensity / l50.Intensity)
  p50off.BlendDisableLighting = a_insert_redlight_off - a_insert_redlight_on * (l50.GetInPlayIntensity / l50.Intensity)
  bulb50.BlendDisableLighting = bulb_off + bulb_on * (l50.GetInPlayIntensity / l50.Intensity)
End Sub

Sub l60_animate
  p60.BlendDisableLighting = insert_redlight_off + insert_redlight_on * (l60.GetInPlayIntensity / l60.Intensity)
  p60off.BlendDisableLighting = a_insert_redlight_off - a_insert_redlight_on * (l60.GetInPlayIntensity / l60.Intensity)
  bulb60.BlendDisableLighting = bulb_off + bulb_on * (l60.GetInPlayIntensity / l60.Intensity)
End Sub

Sub l62_animate
  p62.BlendDisableLighting = insert_redlight_off + insert_redlight_on * (l62.GetInPlayIntensity / l62.Intensity)
  p62off.BlendDisableLighting = a_insert_redlight_off - a_insert_redlight_on * (l62.GetInPlayIntensity / l62.Intensity)
  bulb62.BlendDisableLighting = bulb_off + bulb_on * (l62.GetInPlayIntensity / l62.Intensity)
End Sub

'************************
'**** Orange Inserts

Sub l11_animate
  p11.BlendDisableLighting = insert_orange_off + insert_orange_on * (l11.GetInPlayIntensity / l11.Intensity)
  p11off.BlendDisableLighting = a_insert_orange_off - a_insert_orange_on * (l11.GetInPlayIntensity / l11.Intensity)
  bulb11.BlendDisableLighting = bulb_off + bulb_on * (l11.GetInPlayIntensity / l11.Intensity)
End Sub

Sub l51_animate
  p51.BlendDisableLighting = insert_orange_off + insert_orange_on * (l51.GetInPlayIntensity / l51.Intensity)
  p51off.BlendDisableLighting = a_insert_orange_off - a_insert_orange_on * (l51.GetInPlayIntensity / l51.Intensity)
  bulb51.BlendDisableLighting = bulb_off + bulb_on * (l51.GetInPlayIntensity / l51.Intensity)
End Sub


'************************
'**** Green Inserts

Sub l6_animate
  p6.BlendDisableLighting = insert_green_off + insert_green_on * (l6.GetInPlayIntensity / l6.Intensity)
  p6off.BlendDisableLighting = a_insert_green_off - a_insert_green_on * (l6.GetInPlayIntensity / l6.Intensity)
  bulb6.BlendDisableLighting = bulb_off + bulb_on * (l6.GetInPlayIntensity / l6.Intensity)
End Sub

Sub l22_animate
  p22.BlendDisableLighting = insert_green_off + insert_green_on * (l22.GetInPlayIntensity / l22.Intensity)
  p22off.BlendDisableLighting = a_insert_green_off - a_insert_green_on * (l22.GetInPlayIntensity / l22.Intensity)
  bulb22.BlendDisableLighting = bulb_off + bulb_on * (l22.GetInPlayIntensity / l22.Intensity)
End Sub

Sub l30_animate
  p30.BlendDisableLighting = insert_green_off + insert_green_on * (l30.GetInPlayIntensity / l30.Intensity)
  p30off.BlendDisableLighting = a_insert_green_off - a_insert_green_on * (l30.GetInPlayIntensity / l30.Intensity)
  bulb30.BlendDisableLighting = bulb_off + bulb_on * (l30.GetInPlayIntensity / l30.Intensity)
End Sub

Sub l38_animate
  p38.BlendDisableLighting = insert_green_off + insert_green_on * (l38.GetInPlayIntensity / l38.Intensity)
  p38off.BlendDisableLighting = a_insert_green_off - a_insert_green_on * (l38.GetInPlayIntensity / l38.Intensity)
  bulb38.BlendDisableLighting = bulb_off + bulb_on * (l38.GetInPlayIntensity / l38.Intensity)
End Sub

Sub l47_animate
  p47.BlendDisableLighting = insert_green_off + insert_green_on * (l47.GetInPlayIntensity / l47.Intensity)
  p47off.BlendDisableLighting = a_insert_green_off - a_insert_green_on * (l47.GetInPlayIntensity / l47.Intensity)
  bulb47.BlendDisableLighting = bulb_off + bulb_on * (l47.GetInPlayIntensity / l47.Intensity)
End Sub

Sub l58_animate
  p58.BlendDisableLighting = insert_green_off + insert_green_on * (l58.GetInPlayIntensity / l58.Intensity)
  p58off.BlendDisableLighting = a_insert_green_off - a_insert_green_on * (l58.GetInPlayIntensity / l58.Intensity)
  bulb58.BlendDisableLighting = bulb_off + bulb_on * (l58.GetInPlayIntensity / l58.Intensity)
End Sub


'************************
'**** Yellow Inserts (smaller ones)

Sub l14_animate
  p14.BlendDisableLighting = insert_yellow_off + insert_yellow_on * (l14.GetInPlayIntensity / l14.Intensity)
  p14off.BlendDisableLighting = a_insert_yellow_off - a_insert_yellow_on * (l14.GetInPlayIntensity / l14.Intensity)
  bulb14.BlendDisableLighting = bulb_off + bulb_on * (l14.GetInPlayIntensity / l14.Intensity)
End Sub

Sub l31_animate
  p31.BlendDisableLighting = insert_yellow_off + insert_yellow_on * (l31.GetInPlayIntensity / l31.Intensity)
  p31off.BlendDisableLighting = a_insert_yellow_off - a_insert_yellow_on * (l31.GetInPlayIntensity / l31.Intensity)
  bulb31.BlendDisableLighting = bulb_off + bulb_on * (l31.GetInPlayIntensity / l31.Intensity)
End Sub

Sub l57_animate
  p57.BlendDisableLighting = insert_yellow_off + insert_yellow_on * (l57.GetInPlayIntensity / l57.Intensity)
  p57off.BlendDisableLighting = a_insert_yellow_off - a_insert_yellow_on * (l57.GetInPlayIntensity / l57.Intensity)
  bulb57.BlendDisableLighting = bulb_off + bulb_on * (l57.GetInPlayIntensity / l57.Intensity)
End Sub


'************************
'**** Orange Arrow Inserts

Sub l15_animate
  p15.BlendDisableLighting = insert_orange_off + insert_orange_on * (l15.GetInPlayIntensity / l15.Intensity)
  p15off.BlendDisableLighting = a_insert_orange_off - a_insert_orange_on * (l15.GetInPlayIntensity / l15.Intensity)
  bulb15.BlendDisableLighting = bulb_off + bulb_on * (l15.GetInPlayIntensity / l15.Intensity)
End Sub


'************************
'**** Yellow Arrow Inserts

Sub l9_animate
  p9.BlendDisableLighting = insert_yellow_off + insert_yellow_on * (l9.GetInPlayIntensity / l9.Intensity)
  p9off.BlendDisableLighting = a_insert_yellow_off - a_insert_yellow_on * (l9.GetInPlayIntensity / l9.Intensity)
  bulb9.BlendDisableLighting = bulb_off + bulb_on * (l9.GetInPlayIntensity / l9.Intensity)
End Sub

Sub l25_animate
  p25.BlendDisableLighting = insert_yellow_off + insert_yellow_on * (l25.GetInPlayIntensity / l25.Intensity)
  p25off.BlendDisableLighting = a_insert_yellow_off - a_insert_yellow_on * (l25.GetInPlayIntensity / l25.Intensity)
  bulb25.BlendDisableLighting = bulb_off + bulb_on * (l25.GetInPlayIntensity / l25.Intensity)
End Sub

Sub l41_animate
  p41.BlendDisableLighting = insert_yellow_off + insert_yellow_on * (l41.GetInPlayIntensity / l41.Intensity)
  p41off.BlendDisableLighting = a_insert_yellow_off - a_insert_yellow_on * (l41.GetInPlayIntensity / l41.Intensity)
  bulb41.BlendDisableLighting = bulb_off + bulb_on * (l41.GetInPlayIntensity / l41.Intensity)
End Sub


'************************
'**** Star Trigger Lamps

Sub lstar_animate   'sw25
   star25.BlendDisableLighting = .02 * insertfactor * (lstar.GetInPlayIntensity / lstar.Intensity) + 0.005
   pstar25.BlendDisableLighting = .98 * insertfactor * (lstar.GetInPlayIntensity / lstar.Intensity) + 0.02
End Sub


'************************
'**** Credit Insert

Sub l59_animate: p59.BlendDisableLighting = insert_credit_off + insert_credit_on * (l59.GetInPlayIntensity / l59.Intensity): End Sub


'************************
'**** Bumper Lamps

Sub l10_animate
  flbumpvalue1 = bumperflashlevel * (l10.GetInPlayIntensity / l10.Intensity)
  FlFadeBumper 1, flbumpvalue1
End Sub

Sub l42_animate
  flbumpvalue2 = bumperflashlevel * (l42.GetInPlayIntensity / l42.Intensity)
  FlFadeBumper 2, flbumpvalue2
End Sub

Sub l26_animate
  flbumpvalue3 = bumperflashlevel * (l26.GetInPlayIntensity / l26.Intensity)
  FlFadeBumper 3, flbumpvalue3
  Hole_001.blenddisablelighting = ((0.1 * (l26.GetInPlayIntensity / l26.Intensity)) + (0.1 * LightLevel/lightlevelinitial))
  PegPlasticT_043.blenddisablelighting = (((0.075 * (l26.GetInPlayIntensity / l26.Intensity)) + (0.075 * gilvl)) + (0.005 * LightLevel/lightlevelinitial))
  PegPlasticT_044.blenddisablelighting = (((0.075 * (l26.GetInPlayIntensity / l26.Intensity)) + (0.075 * gilvl)) + (0.005 * LightLevel/lightlevelinitial))
End Sub



'******************************************************
'*****   END 3D INSERTS
'******************************************************



'******************************************************
'****  ZGIU:  GI Control
'******************************************************

'**** SetRelayGI is called from SolCallback.
sub PFGI (Enabled)

Dim x
' debug.print "SetRelayGI value: " & Enabled

  if Enabled = 1 Then

    If DesktopMode = True or VR_Room = 1 Then:  Sound_GI_Relay 1, KnockerPosition

    If DesktopMode = True then
      L_DT_1.intensity = 7
      L_DT_2.intensity = 6
      L_DT_2a.intensity = 6
      L_DT_Player_001.intensity = 2
      L_DT_Player_002.intensity = 2
      L_DT_Player_003.intensity = 2
      L_DT_Player_004.intensity = 2
    Else
      for each x in dt_bg_lights: x.intensity = 0.1: Next
    End If

    if VR_Room = 1 Then
      BGDark.Visible = 0
      BGBright.Visible = 1
    Else
      BGDark.Visible = 0
      BGBright.Visible = 0
    End If

  ' Set Drop Target Shadows to state when GI is back on.
    if DT21up=1 Then dtsh21.visible = 1
    if DT22up=1 Then dtsh22.visible = 1
    if DT23up=1 Then dtsh23.visible = 1
    if DT33up=1 Then dtsh33.visible = 1
    if DT34up=1 Then dtsh34.visible = 1
    if DT35up=1 Then dtsh35.visible = 1

  gilvl = 1

    Else

    If DesktopMode = True or VR_Room = 1 Then:  Sound_GI_Relay 0, KnockerPosition

      for each x in dt_bg_lights: x.intensity = 0.1: Next

    if VR_Room = 1 Then
      BGDark.Visible = 1
      BGBright.Visible = 0
    Else
      BGDark.Visible = 0
      BGBright.Visible = 0
    End If

  ' Set Drop Target Shadows to off when GI is off.
    for each x in ShadowDT
      x.visible=False
    Next

  gilvl = 0

    End If

  SetRoomBrightness

End Sub



'******************************************************
'****  END GI Control
'******************************************************



'******************************************************
'   ZFLB:  FLUPPER BUMPERS
'******************************************************

' prepare some global vars to dim/brighten objects when using day-night slider
Dim DayNightAdjust , DNA30, DNA45, DNA90
If NightDay < 10 Then
  DNA30 = 0
  DNA45 = (NightDay - 10) / 20
  DNA90 = 0
  DayNightAdjust = 0.4
Else
  DNA30 = (NightDay - 10) / 30
  DNA45 = (NightDay - 10) / 45
  DNA90 = (NightDay - 10) / 90
  DayNightAdjust = NightDay / 25
End If

Dim FlBumperFadeActual(6), FlBumperFadeTarget(6), FlBumperColor(6), FlBumperTop(6), FlBumperSmallLight(6), Flbumperbiglight(6)
Dim FlBumperDisk(6), FlBumperBase(6), FlBumperBulb(6), FlBumperscrews(6), FlBumperActive(6), FlBumperHighlight(6)
Dim cnt
For cnt = 1 To 6
  FlBumperActive(cnt) = False
Next

' colors available are red, white, blue, orange, yellow, green, purple and blacklight
FlInitBumper 1, "frontier"
FlInitBumper 2, "frontier"
FlInitBumper 3, "frontier"

Sub FlInitBumper(nr, col)
  FlBumperActive(nr) = True

  ' store all objects in an array for use in FlFadeBumper subroutine
  FlBumperFadeActual(nr) = 1
  FlBumperFadeTarget(nr) = 1.1
  FlBumperColor(nr) = col
  Set FlBumperTop(nr) = Eval("bumpertop" & nr)
  FlBumperTop(nr).material = "bumpertopmat" & nr
  Set FlBumperSmallLight(nr) = Eval("bumpersmalllight" & nr)
  Set Flbumperbiglight(nr) = Eval("bumperbiglight" & nr)
  Set FlBumperDisk(nr) = Eval("bumperdisk" & nr)
  Set FlBumperBase(nr) = Eval("bumperbase" & nr)
  Set FlBumperBulb(nr) = Eval("bumperbulb" & nr)
  FlBumperBulb(nr).material = "bumperbulbmat" & nr
  Set FlBumperscrews(nr) = Eval("bumperscrews" & nr)
  FlBumperscrews(nr).material = "bumperscrew" & col
  Set FlBumperHighlight(nr) = Eval("bumperhighlight" & nr)

  ' set the color for the two VPX lights
  Select Case col

    Case "frontier"
    FlBumperHighlight(nr).color = RGB(239,192,112)
    FlBumperSmallLight(nr).TransmissionScale = 0.1
    FlBumperSmallLight(nr).BulbModulateVsAdd = 0.99
    FlBumperSmallLight(nr).color = RGB(228,112,37)
    FlBumperSmallLight(nr).colorfull = RGB(239,192,112)
    FlBumperBigLight(nr).color = RGB(228,112,37)
    FlBumperBigLight(nr).colorfull = RGB(239,192,112)
    FlBumperDisk(nr).BlendDisableLighting = 0.5

  End Select
End Sub

Sub FlFadeBumper(nr, Z)

  '   UpdateMaterial(string, float wrapLighting, float roughness, float glossyImageLerp, float thickness, float edge, float edgeAlpha, float opacity,
  '        OLE_COLOR base, OLE_COLOR glossy, OLE_COLOR clearcoat, VARIANT_BOOL isMetal, VARIANT_BOOL opacityActive,
  '        float elasticity, float elasticityFalloff, float friction, float scatterAngle) - updates all parameters of a material
  FlBumperDisk(nr).BlendDisableLighting = (0.75 + Z * 3 )
  FlBumperBase(nr).BlendDisableLighting = (0.75 + Z * 4 )

  Select Case FlBumperColor(nr)

    Case "frontier"
    MaterialColor "bumpertopmat" & nr, RGB(255,236 - z*6,155 - Z*15)
    MaterialColor "bumperbulbmat" & nr, RGB(239 -z*11,192 - z*80,112 - Z*85)
    FlBumperTop(nr).BlendDisableLighting = bumperscale * (0.05+ 0.45 * Z)
    FlBumperBulb(nr).BlendDisableLighting = bumperscale * (10 + 2000 * z)
    Flbumperbiglight(nr).intensity = bumperscale * (10 * Z )
    FlBumperHighlight(nr).opacity = bumperscale * (1000 * (Z^3))
    FlBumperSmallLight(nr).color = RGB(239 -z*11,192 - z*80,112 - Z*85)
    FlBumperSmallLight(nr).colorfull = RGB(239 -z*11,192 - z*80,112 - Z*85)
    FlBumperSmallLight(nr).intensity = bumperscale * (1000 * Z)

  End Select
End Sub


'******************************************************
'******  END FLUPPER BUMPERS
'******************************************************


'***************************************************************
' ZSHA: Ambient ball shadows
'***************************************************************

' For dynamic ball shadows, Check the "Raytraced ball shadows" box for the specific light.
' Also make sure the light's z position is around 25 (mid ball)

'Ambient (Room light source)
Const AmbientBSFactor = 0.9    '0 To 1, higher is darker
Const AmbientMovement = 1    '1+ higher means more movement as the ball moves left and right
Const offsetX = 0        'Offset x position under ball (These are if you want to change where the "room" light is for calculating the shadow position,)
Const offsetY = 0        'Offset y position under ball (^^for example 5,5 if the light is in the back left corner)

' *** Trim or extend these to match the number of balls/primitives/flashers on the table!  (will throw errors if there aren't enough objects)
Dim objBallShadow(1)

'Initialization
BSInit

Sub BSInit()
  Dim iii
  'Prepare the shadow objects before play begins
  For iii = 0 To tnob - 1
    Set objBallShadow(iii) = Eval("BallShadow" & iii)
    objBallShadow(iii).material = "BallShadow" & iii
    UpdateMaterial objBallShadow(iii).material,1,0,0,0,0,0,AmbientBSFactor,RGB(0,0,0),0,0,False,True,0,0,0,0
    objBallShadow(iii).Z = 3 + iii / 1000
    objBallShadow(iii).visible = 0
  Next
End Sub


Sub BSUpdate
  Dim s: For s = lob To UBound(gBOT)
    ' *** Normal "ambient light" ball shadow

    'Primitive shadow on playfield, flasher shadow in ramps
    '** If on main and upper pf
    If gBOT(s).Z > 20 And gBOT(s).Z < 30 Then
      objBallShadow(s).visible = 1
      objBallShadow(s).X = gBOT(s).X + (gBOT(s).X - (tablewidth / 2)) / (Ballsize / AmbientMovement) + offsetX
      objBallShadow(s).Y = gBOT(s).Y + offsetY
      'objBallShadow(s).Z = gBOT(s).Z + s/1000 + 1.04 - 25

    '** No shadow if ball is off the main playfield (this may need to be adjusted per table)
    Else
      objBallShadow(s).visible = 0
    End If
  Next
End Sub


'*******************************************
'Digital Display
'*******************************************

Dim Digits(32)
' 1st Player
Digits(0) = Array(LED10,LED11,LED12,LED13,LED14,LED15,LED16,LEDc17)
Digits(1) = Array(LED20,LED21,LED22,LED23,LED24,LED25,LED26)
Digits(2) = Array(LED30,LED31,LED32,LED33,LED34,LED35,LED36)
Digits(3) = Array(LED40,LED41,LED42,LED43,LED44,LED45,LED46,LEDc47)
Digits(4) = Array(LED50,LED51,LED52,LED53,LED54,LED55,LED56)
Digits(5) = Array(LED60,LED61,LED62,LED63,LED64,LED65,LED66)
Digits(6) = Array(LED70,LED71,LED72,LED73,LED74,LED75,LED76)

' 2nd Player
Digits(7) = Array(LED80,LED81,LED82,LED83,LED84,LED85,LED86,LEDc87)
Digits(8) = Array(LED90,LED91,LED92,LED93,LED94,LED95,LED96)
Digits(9) = Array(LED100,LED101,LED102,LED103,LED104,LED105,LED106)
Digits(10) = Array(LED110,LED111,LED112,LED113,LED114,LED115,LED116,LEDc117)
Digits(11) = Array(LED120,LED121,LED122,LED123,LED124,LED125,LED126)
Digits(12) = Array(LED130,LED131,LED132,LED133,LED134,LED135,LED136)
Digits(13) = Array(LED140,LED141,LED142,LED143,LED144,LED145,LED146)

' 3rd Player
Digits(14) = Array(LED150,LED151,LED152,LED153,LED154,LED155,LED156,LEDc157)
Digits(15) = Array(LED160,LED161,LED162,LED163,LED164,LED165,LED166)
Digits(16) = Array(LED170,LED171,LED172,LED173,LED174,LED175,LED176)
Digits(17) = Array(LED180,LED181,LED182,LED183,LED184,LED185,LED186,LEDc187)
Digits(18) = Array(LED190,LED191,LED192,LED193,LED194,LED195,LED196)
Digits(19) = Array(LED200,LED201,LED202,LED203,LED204,LED205,LED206)
Digits(20) = Array(LED210,LED211,LED212,LED213,LED214,LED215,LED216)

' 4th Player
Digits(21) = Array(LED220,LED221,LED222,LED223,LED224,LED225,LED226,LEDc227)
Digits(22) = Array(LED230,LED231,LED232,LED233,LED234,LED235,LED236)
Digits(23) = Array(LED240,LED241,LED242,LED243,LED244,LED245,LED246)
Digits(24) = Array(LED250,LED251,LED252,LED253,LED254,LED255,LED256,LEDc257)
Digits(25) = Array(LED260,LED261,LED262,LED263,LED264,LED265,LED266)
Digits(26) = Array(LED270,LED271,LED272,LED273,LED274,LED275,LED276)
Digits(27) = Array(LED280,LED281,LED282,LED283,LED284,LED285,LED286)

' Credits
Digits(28) = Array(LED4,LED2,LED6,LED7,LED5,LED1,LED3)
Digits(29) = Array(LED18,LED9,LED27,LED28,LED19,LED8,LED17)

' Balls
Digits(30) = Array(LED39,LED37,LED48,LED49,LED47,LED29,LED38)
Digits(31) = Array(LED67,LED58,LED69,LED77,LED68,LED57,LED59)

Sub DisplayTimer
  Dim ChgLED,ii,num,chg,stat,obj
  ChgLed = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
If Not IsEmpty(ChgLED) Then
    If DesktopMode = True Then
    For ii = 0 To UBound(chgLED)
      num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
      if (num < 32) then
        For Each obj In Digits(num)
          If chg And 1 Then obj.State = stat And 1
          chg = chg\2 : stat = stat\2
        Next
      else
      end if
    next
    end if
end if
End Sub


'******************************************************
' VR Backglass Digital Display
'******************************************************

Dim xoff,yoff1, yoff2, yoff3, yoff4, yoff5, zoff,xrot,zscale, xcen,ycen

Sub Setup_Backglass()

  xoff = -20
  yoff1 = 35 ' this is where you adjust the forward/backward position for player 1 score
  yoff2 = 35 ' this is where you adjust the forward/backward position for player 2 score
  yoff3 = 35 ' this is where you adjust the forward/backward position for player 3 score
  yoff4 = 35 ' this is where you adjust the forward/backward position for player 4 score
  yoff5 = 35 ' this is where you adjust the forward/backward position for credits and ball in play
  zoff = 638
  xrot = -90

  CenterVRDigits()

End Sub

Sub CenterVRDigits()
  Dim ix, xx, yy, yfact, xfact, xobj

  zscale = 0.0000001

  xcen = (130 /2) - (92 / 2)
  ycen = (780 /2 ) + (203 /2)

  for ix = 0 to 6
    For Each xobj In VRDigits(ix)

      xx = xobj.x

      xobj.x = (xoff - xcen) + xx
      yy = xobj.y ' get the yoffset before it is changed
      xobj.y = yoff1

      If (yy < 0.) then
        yy = yy * -1
      end if

      xobj.height = (zoff - ycen) + yy - (yy * (zscale))
      xobj.rotx = xrot
    Next
  Next

  for ix = 7 to 13
    For Each xobj In VRDigits(ix)

      xx = xobj.x

      xobj.x = (xoff - xcen) + xx
      yy = xobj.y ' get the yoffset before it is changed
      xobj.y = yoff2

      If (yy < 0.) then
        yy = yy * -1
      end if

      xobj.height = (zoff - ycen) + yy - (yy * (zscale))
      xobj.rotx = xrot
    Next
  Next

  for ix = 14 to 20
    For Each xobj In VRDigits(ix)

      xx = xobj.x

      xobj.x = (xoff - xcen) + xx
      yy = xobj.y ' get the yoffset before it is changed
      xobj.y = yoff3

      If (yy < 0.) then
        yy = yy * -1
      end if

      xobj.height = (zoff - ycen) + yy - (yy * (zscale))
      xobj.rotx = xrot
    Next
  Next

  for ix = 21 to 27
    For Each xobj In VRDigits(ix)

      xx = xobj.x

      xobj.x = (xoff - xcen) + xx
      yy = xobj.y ' get the yoffset before it is changed
      xobj.y = yoff4

      If (yy < 0.) then
        yy = yy * -1
      end if

      xobj.height = (zoff - ycen) + yy - (yy * (zscale))
      xobj.rotx = xrot
    Next
  Next

  for ix = 28 to 31
    For Each xobj In VRDigits(ix)

      xx = xobj.x

      xobj.x = (xoff - xcen) + xx
      yy = xobj.y ' get the yoffset before it is changed
      xobj.y = yoff5

      If (yy < 0.) then
        yy = yy * -1
      end if

      xobj.height = (zoff - ycen) + yy - (yy * (zscale))
      xobj.rotx = xrot
    Next
  Next

End Sub

Dim VRDigits(32)
VRDigits(0) = Array(LED1x0,LED1x1,LED1x2,LED1x3,LED1x4,LED1x5,LED1x6,led1x7)
VRDigits(1) = Array(LED2x0,LED2x1,LED2x2,LED2x3,LED2x4,LED2x5,LED2x6)
VRDigits(2) = Array(LED3x0,LED3x1,LED3x2,LED3x3,LED3x4,LED3x5,LED3x6)
VRDigits(3) = Array(LED4x0,LED4x1,LED4x2,LED4x3,LED4x4,LED4x5,LED4x6,led4x7)
VRDigits(4) = Array(LED5x0,LED5x1,LED5x2,LED5x3,LED5x4,LED5x5,LED5x6)
VRDigits(5) = Array(LED6x0,LED6x1,LED6x2,LED6x3,LED6x4,LED6x5,LED6x6)
VRDigits(6) = Array(LED7x0,LED7x1,LED7x2,LED7x3,LED7x4,LED7x5,LED7x6)

VRDigits(7) = Array(LED8x0,LED8x1,LED8x2,LED8x3,LED8x4,LED8x5,LED8x6,led8x7)
VRDigits(8) = Array(LED9x0,LED9x1,LED9x2,LED9x3,LED9x4,LED9x5,LED9x6)
VRDigits(9) = Array(LED10x0,LED10x1,LED10x2,LED10x3,LED10x4,LED10x5,LED10x6)
VRDigits(10) = Array(LED11x0,LED11x1,LED11x2,LED11x3,LED11x4,LED11x5,LED11x6,led11x7)
VRDigits(11) = Array(LED12x0,LED12x1,LED12x2,LED12x3,LED12x4,LED12x5,LED12x6)
VRDigits(12) = Array(LED13x0,LED13x1,LED13x2,LED13x3,LED13x4,LED13x5,LED13x6)
VRDigits(13) = Array(LED14x0,LED14x1,LED14x2,LED14x3,LED14x4,LED14x5,LED14x6)

VRDigits(14) = Array(LED1x000,LED1x001,LED1x002,LED1x003,LED1x004,LED1x005,LED1x006,LED1x007)
VRDigits(15) = Array(LED1x100,LED1x101,LED1x102,LED1x103,LED1x104,LED1x105,LED1x106)
VRDigits(16) = Array(LED1x200,LED1x201,LED1x202,LED1x203,LED1x204,LED1x205,LED1x206)
VRDigits(17) = Array(LED1x300,LED1x301,LED1x302,LED1x303,LED1x304,LED1x305,LED1x306,LED1x307)
VRDigits(18) = Array(LED1x400,LED1x401,LED1x402,LED1x403,LED1x404,LED1x405,LED1x406)
VRDigits(19) = Array(LED1x500,LED1x501,LED1x502,LED1x503,LED1x504,LED1x505,LED1x506)
VRDigits(20) = Array(LED1x600,LED1x601,LED1x602,LED1x603,LED1x604,LED1x605,LED1x606)

VRDigits(21) = Array(LED2x000,LED2x001,LED2x002,LED2x003,LED2x004,LED2x005,LED2x006,led2x007)
VRDigits(22) = Array(LED2x100,LED2x101,LED2x102,LED2x103,LED2x104,LED2x105,LED2x106)
VRDigits(23) = Array(LED2x200,LED2x201,LED2x202,LED2x203,LED2x204,LED2x205,LED2x206)
VRDigits(24) = Array(LED2x300,LED2x301,LED2x302,LED2x303,LED2x304,LED2x305,LED2x306,led2x307)
VRDigits(25) = Array(LED2x400,LED2x401,LED2x402,LED2x403,LED2x404,LED2x405,LED2x406)
VRDigits(26) = Array(LED2x500,LED2x501,LED2x502,LED2x503,LED2x504,LED2x505,LED2x506)
VRDigits(27) = Array(LED2x600,LED2x601,LED2x602,LED2x603,LED2x604,LED2x605,LED2x606)

VRDigits(28) = Array(LEDax300,LEDax301,LEDax302,LEDax303,LEDax304,LEDax305,LEDax306)
VRDigits(29) = Array(LEDbx400,LEDbx401,LEDbx402,LEDbx403,LEDbx404,LEDbx405,LEDbx406)
VRDigits(30) = Array(LEDcx500,LEDcx501,LEDcx502,LEDcx503,LEDcx504,LEDcx505,LEDcx506)
VRDigits(31) = Array(LEDdx600,LEDdx601,LEDdx602,LEDdx603,LEDdx604,LEDdx605,LEDdx606)

dim DisplayColor

DisplayColor =  RGB(255,40,1)

Sub VRDisplayTimer
    Dim ChgLED, ii, jj, num, chg, stat, obj, b, x
    ChgLED=Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
       For ii=0 To UBound(chgLED)
          num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)
      if (num < 34) then
              For Each obj In VRDigits(num)
'                   If chg And 1 Then obj.visible=stat And 1
           If chg And 1 Then FadeDisplay obj, stat And 1
                   chg=chg\2 : stat=stat\2
              Next
      Else
             end if
        Next
    End If
 End Sub

Sub FadeDisplay(object, onoff)
  If OnOff = 1 Then
    object.color = DisplayColor
    Object.Opacity = 12
  Else
    Object.Color = RGB(1,1,1)
    Object.Opacity = 8
  End If
End Sub

Sub InitVRDigits()
  dim tmp, x, obj
  for x = 0 to uBound(VRDigits)
    if IsArray(VRDigits(x) ) then
      For each obj in VRDigits(x)
        obj.height = obj.height + 18
        FadeDisplay obj, 0
      next
    end If
  Next
End Sub

InitVRDigits


'******************************************************
' Ball brightness code
'******************************************************



'****************************************************************
' Ball option setup
'****************************************************************

Dim CustomBallImage(14), CustomBallLogoMode(14), CustomBallDecal(14), CustomBallRS(14), CustomBulbIntensity(14)

' dark
CustomBallImage(0) =    "ball-dark"
CustomBallLogoMode(0) =   False
CustomBallDecal(0) =    "JPBall-Scratches"
CustomBulbIntensity(0) =  1.0
CustomBallRS(0) =       0.4

' not as dark
CustomBallImage(1) =    "ball_HDR"
CustomBallLogoMode(1) =   False
CustomBallDecal(1) =    "Scratches"
CustomBulbIntensity(1) =  1.0
CustomBallRS(1) =       0.4

' bright
CustomBallImage(2) =    "ball-light-hf"
CustomBallLogoMode(2) =   False
CustomBallDecal(2) =    "scratchedmorelight"
CustomBulbIntensity(2) =  1.0
CustomBallRS(2) =       0.4

' brightest
CustomBallImage(3) =    "ball-lighter-hf"
CustomBallLogoMode(3) =   False
CustomBallDecal(3) =    "scratchedmorelight"
CustomBulbIntensity(3) =  1.0
CustomBallRS(3) =       0.4

' warm
CustomBallImage(4) =    "ball-hf-warm"
CustomBallLogoMode(4) =   False
CustomBallDecal(4) =    "scratchedmorelight"
CustomBulbIntensity(4) =  1.0
CustomBallRS(4) =       0.4

' MrBallDark
CustomBallImage(5) =    "ball_MRBallDark"
CustomBallLogoMode(5) =   False
CustomBallDecal(5) =    "ball_scratches"
CustomBulbIntensity(5) =  1.0
CustomBallRS(5) =       0.4

' DarthVito Ball
CustomBallImage(6) =    "ball"
CustomBallLogoMode(6) =   False
CustomBallDecal(6) =    "ball_scratches"
CustomBulbIntensity(6) =  1.0
CustomBallRS(6) =       0.4

' DarthVito HDR1
CustomBallImage(7) =    "ball_HDR"
CustomBallLogoMode(7) =   False
CustomBallDecal(7) =    "ball_scratches"
CustomBulbIntensity(7) =  1.0
CustomBallRS(7) =       0.4

' DarthVito HDR Bright
CustomBallImage(8) =    "ball_HDR_bright"
CustomBallLogoMode(8) =   False
CustomBallDecal(8) =    "ball_scratches"
CustomBulbIntensity(8) =  1.0
CustomBallRS(8) =       0.4

' DarthVito HDR2
CustomBallImage(9) =    "ball_HDR3b"
CustomBallLogoMode(9) =   False
CustomBallDecal(9) =    "ball_scratches"
CustomBulbIntensity(9) =  1.0
CustomBallRS(9) =       0.4

' DarthVito ballHDR3
CustomBallImage(10) =     "ball_HDR3b2"
CustomBallLogoMode(10) =  False
CustomBallDecal(10) =     "ball_scratches"
CustomBulbIntensity(10) =   1.0
CustomBallRS(10) =      0.4

' DarthVito ballHDR4
CustomBallImage(11) =     "ball_HDR3b3"
CustomBallLogoMode(11) =  False
CustomBallDecal(11) =     "ball_scratches"
CustomBulbIntensity(11) =   1.0
CustomBallRS(11) =      0.4

' DarthVito ballHDRdark
CustomBallImage(12) =     "ball_HDR3dark"
CustomBallLogoMode(12) =  False
CustomBallDecal(12) =     "ball_scratches"
CustomBulbIntensity(12) =   1.0
CustomBallRS(12) =      0.4

' Borg Ball
CustomBallImage(13) =     "BorgBall"
CustomBallLogoMode(13) =  False
CustomBallDecal(13) =     "ball_scratches"
CustomBulbIntensity(13) =   1.0
CustomBallRS(13) =      0.4

' SteelBall2
CustomBallImage(14) =     "SteelBall2"
CustomBallLogoMode(14) =  False
CustomBallDecal(14) =     "ball_scratches"
CustomBulbIntensity(14) =   1.0
CustomBallRS(14) =      0.4


Sub ChangeBall(ballnr)
  Dim ii
  table1.BallDecalMode = CustomBallLogoMode(ballnr)
  table1.BallFrontDecal = CustomBallDecal(ballnr)
  table1.DefaultBulbIntensityScale = CustomBulbIntensity(ballnr)
  table1.BallImage = CustomBallImage(ballnr)
  For ii = 0 to UBound(gBOT)
    gBOT(ii).Image = CustomBallImage(ballnr)
    gBOT(ii).FrontDecal = CustomBallDecal(ballnr)
    gBOT(ii).PlayfieldReflectionScale = CustomBallRS(ballnr)
  Next
End Sub

'******************************************************
' Options
'******************************************************

Sub SetupOptions

Dim C

'  ROM Volume Settings
    With Controller
        .Games(cGameName).Settings.Value("volume") = ROMVolume
  End With

' Plastic Edge Color

  If PlasticColor = 6 Then
    PlasProt = RndInt(0,5)
  Else
    PlasProt = PlasticColor
  End if

  SetAcrylicColor PlasProt

' Flipper Color

  if FlipperColor = 1 Then
    Lflipmesh1.image = "flipper_red-left"
    Rflipmesh1.image = "flipper_red-right"
  Elseif FlipperColor = 2 Then
    Lflipmesh1.image = "flipper_blue-left"
    Rflipmesh1.image = "flipper_blue-right"
  Elseif FlipperColor = 3 Then
    Lflipmesh1.image = "flipper_yellow-left"
    Rflipmesh1.image = "flipper_yellow-right"
  End If

' Screw Cap Color
  If ScrewCapColor  = 1  Then
    For each C in GIScrewCaps
      C.material = "Metal Lamp Black"
    Next
  Elseif ScrewCapColor  = 2 Then
    For each C in GIScrewCaps
      C.material = "Metal Lamp Bronze"
    Next
  Elseif ScrewCapColor  = 3 Then
    For each C in GIScrewCaps
      C.material = "Metal Lamp Silver"
    Next
  Elseif ScrewCapColor  = 4 Then
    For each  C in GIScrewCaps
      C.material = "Metal Lamp BlueGreen"
    Next
  Elseif ScrewCapColor  = 5 Then
    For each C in GIScrewCaps
      C.material = "Metal Lamp Gold"
    Next
  Else
    For each C in GIScrewCaps
      C.material = "Metal Lamp White"
    Next
  End If

' Standup Target Color
  If TargetColor  = 1  Then
    For each C in GITargets
      C.image = "target_red"
    Next
    sw30p.image = "target1color" 'middle
  Else
    For each C in GITargets
      C.image = "target_red-streak"
    Next
    sw30p.image = "target1color-streak" 'middle
  End If

SetupVRRoom

End Sub


'******************************************************
' DIP Switches
'******************************************************

'Bally Frontier

Sub editDips
  if SetDIPSwitches = 1 Then
    Dim vpmDips:Set vpmDips = New cvpmDips
    With vpmDips
      .AddForm 315, 370, "Frontier DIP Switch Settings"
      .AddFrame 2, 10, 140, "Balls per game", &HC0000000, Array("2", &HC0000000, "3", 0, "4", &H80000000, "5", &H40000000)
      .AddFrame 160, 10, 140, "Max Credits", &H03000000, Array("10", 0, "15", &H01000000, "25", &H02000000, "40", &H03000000)
      .AddFrame 2, 90, 140, "A && C Lanes are", &H00000020, Array("Separate", 0, "Linked", &H00000020)
      .AddFrame 160, 90, 140, "Frontier Bonus Max", &H00400000, Array("60,000", 0, "110,000", &H00400000)
      .AddFrame 2, 140, 140, "Outlane Special lights", &H00000080, Array("Alternate", 0, "Both light", &H00000080)
      .AddFrame 160, 140, 140, "Outlane Special lights when", &H00002000, Array("3X lights", 0, "4X lights", &H00002000)
      .AddFrame 2, 190, 140, "Feeder Lane Special lights", &H20000000, Array("Alternate", 0, "Both light", &H20000000)
      .AddChk 160, 200, 148, Array("Credit display", &H04000000)
      .AddChk 160, 220, 148, Array("Match", &H08000000)
      .AddChk 7, 250, 300, Array("A-B-C Extra Ball light on at start of game", &H00000040)
      .AddChk 7, 270, 300, Array("A-B-C lane memory", &H00004000)
      'Have to use 32768 because vbscript treats &H00008000 as -32768
      .AddChk 7, 290, 300, Array("Outhole collects Frontier Bonus", 32768)
      .AddChk 7, 310, 300, Array("Gate memory", &H00800000)
      .AddLabel 7, 340, 300, 20, "After hitting OK, press F3 to reset game with new settings."
      .ViewDips
    End With
  End If
End Sub
Set vpmShowDips = GetRef("editDips")


'**********************************************************************************************************
'**********************************************************************************************************

'Recommended settings by Bally, but for conservative/harder play.

'  *** This is the default dips settings ***

'1-5  00000   1 credit/coins chute left
'6    0   A and C rollover lane lites: Only puts that lane out
'7    0   A-B-C extra ball: Lite comes on after making A-B-C lanes
'8    1   Special Outlane: Both lites comes on
'9-13   00000   1 credit/coins chute right
'14   1   Special Outlane lites and side target:  Comes on when drop target 4X is on.
'15   0   A-B-C Lite Recall:  Any A-B-C lite out will come back on for next ball.
'16   1   Frontier Bonus score adj: Going in outhole Frontier bonus will be collected.
'17-20  0000  Credit/coins chutes 2 same as chute 1
'21-22  00    n/a
'23   0   Frontier bonus will stop at 60,000
'24   1   Gate Recall: going into outhole with gate up will keep gate up for next ball.
'25-26  11    Max credits of 40
'27   1     Credit display lite is on
'28   1     Match feature on
'29   0   n/a
'30   1   Flipper feed lane lite adj: Both lites come on.
'31-32  00    ball per game = 3



Sub SetDefaultDips
  If SetDIPSwitches = 0 Then
    Controller.Dip(0) = 128   '8-1 = 10000000
    Controller.Dip(1) = 160   '16-9 = 10100000
    Controller.Dip(2) = 128   '24-17 = 10000000
    Controller.Dip(3) = 47    '32-25 = 00101111
    Controller.Dip(4) = 0
    Controller.Dip(5) = 0
  End If
End Sub


'******************************************************
' ZVRR
'******************************************************

Sub SetupVRRoom

Dim VRThings

if VR_Room = 0 and cab_mode = 0 Then
  for each VRThings in VRStuff:VRThings.visible = 0:Next
  for each VRThings in VRClock:VRThings.visible = 0:Next
  for each VRThings in VRBackglass:VRThings.visible = 0:Next
  for each VRThings in TopHardware:VRThings.visible = 1:Next
  OuterPrimBlack_DT.z = -318
  OuterPrimBlack_DT.size_y = 165
  FrontWall.sidevisible = 1
  for each VRThings in dt_bg_lights: VRThings.State = 1:Next

Elseif VR_Room = 0 and cab_mode = 1 Then
  for each VRThings in VRStuff:VRThings.visible = 0:Next
  for each VRThings in VRClock:VRThings.visible = 0:Next
  for each VRThings in VRBackglass:VRThings.visible = 0:Next
  for each VRThings in DTBackglass:VRThings.visible = 0: Next
  for each VRThings in TopHardware:VRThings.visible = 0:Next
    if cabsideblades = 0 Then
      OuterPrimBlack_DT.z = -180
      OuterPrimBlack_DT.size_y = 220
    Else
      OuterPrimBlack_DT.z = -318
      OuterPrimBlack_DT.size_y = 165
    End If
  FrontWall.sidevisible = 0
  for each VRThings in dt_bg_lights: VRThings.State = 0:Next
  for each VRThings in dt_bg_lights: VRThings.Visible = 0:Next

Else
  for each VRThings in VRStuff:VRThings.visible = 1:Next
  for each VRThings in VRClock:VRThings.visible = WallClock:Next
  for each VRThings in DTBackglass:VRThings.visible = 0: Next
  for each VRThings in TopHardware:VRThings.visible = 1:Next
  OuterPrimBlack_DT.z = -318
  OuterPrimBlack_DT.size_y = 165
  FrontWall.sidevisible = 1
  for each VRThings in dt_bg_lights: VRThings.State = 0:Next
  for each VRThings in dt_bg_lights: VRThings.Visible = 0:Next

'Custom Walls, Floor, and Roof
  if CustomWalls = 1 Then
    VR_Wall_Left.image = "VR_Wall_Left"
    VR_Wall_Right.image = "VR_Wall_Right"
    VR_Floor.image = "VR_Floor"
    VR_Roof.image = "VR_Roof"
  Elseif CustomWalls = 2 Then
    VR_Wall_Left.image = "VR_Wall_Left_DV2"
    VR_Wall_Right.image = "VR_Wall_Right_DV2"
    VR_Floor.image = "VR_Room Floor"
    VR_Roof.image = "VR_Room Roof 2"
  Elseif CustomWalls = 3 Then
    VR_Wall_Left.image = "VR_Wall_Left_DV1"
    VR_Wall_Right.image = "VR_Wall_Right_DV1"
    VR_Floor.image = "VR_Room Floor"
    VR_Roof.image = "VR_Room Roof 2"
  Elseif CustomWalls = 4 Then
    VR_Wall_Left.image = "VR_Wall_Left_DV3"
    VR_Wall_Right.image = "VR_Wall_Right_DV3"
    VR_Floor.image = "VR_Room Floor"
    VR_Roof.image = "VR_Room Roof 2"
  Else
    VR_Wall_Left.image = "wallpaper left"
    VR_Wall_Right.image = "wallpaper_right"
    VR_Floor.image = "FloorCompleteMap2"
    VR_Roof.image = "Light Gray"
  end if

  If poster = 1 Then
    VRposter.visible = 1
  Else
    VRposter.visible = 0
  End If

  If poster2 = 1 Then
    VRposter2.visible = 1
  Else
    VRposter2.visible = 0
  End If

  if undercab = 1 Then
    UndercabLight.state = 1
  Else
    UndercabLight.state = 0
  End If

End If

End Sub


'******************************************************
' VR CLOCK
'******************************************************

Dim CurrentMinute ' for VR clock

Sub ClockTimer_Timer()

    'ClockHands Below *********************************************************
  Pminutes.RotAndTra2 = (Minute(Now())+(Second(Now())/100))*6
  Phours.RotAndTra2 = Hour(Now())*30+(Minute(Now())/2)
  Pseconds.RotAndTra2 = (Second(Now()))*6
  CurrentMinute=Minute(Now())

End Sub


'******************************************************
' LAMP CALLBACK for the bckglass flasher lamps
'******************************************************

if VR_Room = 0 AND cab_mode = 0 Then
  Set LampCallback = GetRef("UpdateDTLamps")
End If

if VR_Room = 1 Then
  Set LampCallback = GetRef("UpdateVRLamps")
End If


Sub UpdateDTLamps()
  CreditsREEL.setValue(1)
  If Controller.Lamp(43) = 0 Then: ShootAgainReel.setValue(0):  Else: ShootAgainReel.setValue(1)    'Shoot Again
  If Controller.Lamp(13) = 0 Then: BIPReel.setValue(0):     Else: BIPReel.setValue(1)         'Ball in Play
  If Controller.Lamp(61) = 0 Then: TiltReel.setValue(0):      Else: TiltReel.setValue(1)        'Tilt
  If Controller.Lamp(45) = 0 Then: GameOverReel.setValue(0):    Else: GameOverReel.setValue(1)      'Game Over
  If Controller.Lamp(27) = 0 Then: MatchReel.setValue(0):     Else: MatchReel.setValue(1)       'Match
  If Controller.Lamp(29) = 0 Then: HighScoreReel.setValue(0):   Else: HighScoreReel.setValue(1)     'High Score

End Sub


Sub UpdateVRLamps()
  If Controller.Lamp(43) = 0 Then: VR_ShootAgain.visible=0:     Else: VR_ShootAgain.visible=1       'Shoot again
  If Controller.Lamp(13) = 0 Then: VR_BIP.visible=0:        Else: VR_BIP.visible=1          'Ball in  play
  If Controller.Lamp(61) = 0 Then: VR_Tilt.visible=0:       Else: VR_Tilt.visible=1         'Tilt
  If Controller.Lamp(45) = 0 Then: VR_GameOver.visible=0:     Else: VR_GameOver.visible=1       'Game Over
  If Controller.Lamp(27) = 0 Then: VR_Match.visible=0:      Else: VR_Match.visible=1        'Match
  If Controller.Lamp(29) = 0 Then: VR_HighScore.visible=0:    Else: VR_HighScore.visible=1      'High Score

End Sub


'*******************************
'Acrylic plastics
'*******************************


Sub SetMaterialColor(name, new_color)
    Dim wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle
    GetMaterial name, wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle
    base = new_color
    UpdateMaterial name, wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle
End Sub

Sub SetMaterialGlossyColor(name, new_color)
    Dim wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle
    GetMaterial name, wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle
    glossy = new_color
    UpdateMaterial name, wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle
End Sub

Sub SetMaterialCKColor(name, new_color)
    Dim wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle
    GetMaterial name, wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle
    clearcoat = new_color
    UpdateMaterial name, wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle
End Sub


DIM ProtCol

sub SetAcrylicColor(ColorOption)

  for each Protcol in Acrylics:Protcol.SideVisible = 1:Next
  for each Protcol in AcrylicsLEDMod:Protcol.SideVisible = 0:Next

  if PlasProt = 0 Then
    for each Protcol in AcrylicsLEDMod:Protcol.visible = 0:Next
    for each Protcol in Acrylics:Protcol.visible = 1:Next
  Else
    for each Protcol in AcrylicsLEDMod:Protcol.visible = 1:Next
    for each Protcol in Acrylics:Protcol.visible = 0:Next
  end If

  if Not ColorOption = 5 then
    RGBAcrylTimer.enabled = false
  End If

  if Not ColorOption = 0 then
    for each Protcol in Acrylics:Protcol.blenddisablelighting = 1:Next
    for each Protcol in Acrylics:Protcol.Topmaterial = "AcrylicBLTops":Next
    for each Protcol in Acrylics:Protcol.Sidematerial = "AcrylicBLSides":Next
    for each Protcol in Acrylics:Protcol.image = "Colour_Trans_Grey":Next
    for each Protcol in Acrylics:Protcol.Sideimage = "Colour_Cream_Trans":Next
  End If

  If ColorOption = 0 Then    'clear
    for each Protcol in Acrylics:Protcol.blenddisablelighting = 0:Next
    for each Protcol in Acrylics:Protcol.Topmaterial = "acrylic top":Next
    for each Protcol in Acrylics:Protcol.Sidematerial = "acrylic side":Next
    for each Protcol in Acrylics:Protcol.image = "":Next
    for each Protcol in Acrylics:Protcol.Sideimage = "":Next
    End If

  If ColorOption = 1 Then    'orangered
        SetMaterialColor "AcrylicBLSides",GIlvlRGB(RGB(255,8,0))
        SetMaterialColor "AcrylicBLTops",GIlvlRGB(RGB(0,0,0))
        SetMaterialGlossyColor "AcrylicBLSides",GIlvlRGB(RGB(255,8,0))
        SetMaterialGlossyColor "AcrylicBLTops",GIlvlRGB(RGB(0,0,0))
        SetMaterialCKColor "AcrylicBLSides",GIlvlRGB(RGB(255,8,0))
        SetMaterialCKColor "AcrylicBLTops",GIlvlRGB(RGB(0,0,0))
    End If

  If ColorOption = 2 Then    'yellow
        SetMaterialColor "AcrylicBLSides",GIlvlRGB(RGB(255,128,0))
        SetMaterialColor "AcrylicBLTops",GIlvlRGB(RGB(0,0,0))
        SetMaterialGlossyColor "AcrylicBLSides",GIlvlRGB(RGB(255,128,0))
        SetMaterialGlossyColor "AcrylicBLTops",GIlvlRGB(RGB(0,0,0))
        SetMaterialCKColor "AcrylicBLSides",GIlvlRGB(RGB(255,128,0))
        SetMaterialCKColor "AcrylicBLTops",GIlvlRGB(RGB(0,0,0))
    End If

  If ColorOption = 3 Then    'blue
        SetMaterialColor "AcrylicBLSides",GIlvlRGB(RGB(15,65,100))
        SetMaterialColor "AcrylicBLTops",GIlvlRGB(RGB(0,0,0))
        SetMaterialGlossyColor "AcrylicBLSides",GIlvlRGB(RGB(15,65,100))
        SetMaterialGlossyColor "AcrylicBLTops",GIlvlRGB(RGB(0,0,0))
        SetMaterialCKColor "AcrylicBLSides",GIlvlRGB(RGB(15,65,100))
        SetMaterialCKColor "AcrylicBLTops",GIlvlRGB(RGB(0,0,0))
    End If

  If ColorOption = 4 Then 'user defined
    dim currentColor : currentColor = GIlvlRGB(UserAcrylColor)
    SetMaterialColor "AcrylicBLSides",GIlvlRGB(UserAcrylColor)
    SetMaterialColor "AcrylicBLTops",GIlvlRGB(RGB(0,0,0))
        SetMaterialGlossyColor "AcrylicBLSides",GIlvlRGB(UserAcrylColor)
        SetMaterialGlossyColor "AcrylicBLTops",GIlvlRGB(RGB(0,0,0))
    SetMaterialCKColor "AcrylicBLSides",GIlvlRGB(UserAcrylColor)
    SetMaterialCKColor "AcrylicBLTops",GIlvlRGB(RGB(0,0,0))
  end if

  If ColorOption = 5 Then 'rainbow
    RGBAcrylTimer.enabled = true
  end if

end sub


Function GIlvlRGB(color)
  Dim red, green, blue, dimFactor, dimEquation

  dimFactor = 5
' dimEquation = ((1+((dimFactor-1)*gilvl))/dimFactor) 'decided to leave LED eges always on if option chosen... rather than turn on and off with each GI.
  dimEquation = ((1+(dimFactor-1))/dimFactor)

  red = color And &HFF
  green = (color \ &H100) And &HFF
  blue = (color \ &H10000) And &HFF

  GIlvlRGB = RGB(red*dimEquation, green*dimEquation, blue*dimEquation)
End Function

dim RGBAcrylHueAngle : RGBAcrylHueAngle = 0
sub RGBAcrylTimer_timer
  RGBAcrylHueAngle = RGBAcrylHueAngle + 1
  if RGBAcrylHueAngle > 359 then RGBAcrylHueAngle = 0

  RGBAcrylHue RGBAcrylHueAngle

  SetMaterialColor "AcrylicBLSides",RGBAcrylColor
  SetMaterialColor "AcrylicBLTops",GIlvlRGB(RGB(0,0,0))
  SetMaterialGlossyColor "AcrylicBLSides",RGBAcrylColor
  SetMaterialGlossyColor "AcrylicBLTops",GIlvlRGB(RGB(0,0,0))
  SetMaterialCKColor "AcrylicBLSides",RGBAcrylColor
  SetMaterialCKColor "AcrylicBLTops",GIlvlRGB(RGB(0,0,0))
end sub

dim RGBAcrylColor: RGBAcrylColor = RGB(240,20,20)

Sub RGBAcrylHue(HueDegrees)
    dim ColorR, ColorG, ColorB
    dim outR, outG, outB

    'UW's from Hue angle
    dim U : U = cos(HueDegrees * PI/180)
    dim W : W = sin(HueDegrees * PI/180)
    'Initial color
    ColorR = 240
    ColorG = 20
    ColorB = 20
    'RGB values from Hue
    outR = Cint((.299+.701*U+.168*W)*ColorR + (.587-.587*U+.330*W)*ColorG + (.114-.114*U-.497*W) * ColorB)
    outG = Cint((.299-.299*U-.328*W)*ColorR + (.587+.413*U+.035*W)*ColorG + (.114-.114*U+.292*W) * ColorB)
    outB = Cint((.299-.300*U+1.25*W)*ColorR + (.587-.588*U-1.05*W)*ColorG + (.114+.886*U-.203*W) * ColorB)

    'limits as equations above may not be that precise
    if outR > 255 then outR = 255
    if outG > 255 then outG = 255
    if outB > 255 then outB = 255

    if outR < 0 then outR = 0
    if outG < 0 then outG = 0
    if outB < 0 then outB = 0

    RGBAcrylColor = GIlvlRGB(RGB(outR, outG, outB))
end sub


'*******************************
'end plastics
'*******************************


'******** Work done by UnclePaulie on Hybrid version v.01 - 2.0 *********

'0.01 Started with Supersonic as same MPU, style, physics, etc.  Updated to latest VPW standards, and updates I did on Centaur and Alien Star.
'   Utilized playfield from Bord that he had scanned in at 150 dpi.  Redbone cleaned up.
'   Created base script to utilize latest VPW standards
'   Automated the desktop, VR, and cab mode
'   New apron prim and adjusted apron and physics.
'   New rail guide prims, heights, and matched exactly with pinballshenanigan's restoration
'   Added flippers from centair (new prim UP updated) - will need new colors
'   Added slings
'   Added undercab lighting option to VR and tied to gilvl.
'   Placed upper rail and side wall.
'   Added the star trigger... tied to l48 lamp, and sw25
'   Added VR Room environment, and VR cabinet and backglass.
'   Added desktop lockdown bar.  Added backrail prim, and outerwall prims.  Adjusted appropriately.
'   Using same lockdownbar and side rails for desktop and VR.  New siderails with metal side plates
'   Updated the ROM settings to be defaulted to the manual recommendations
'   Created a .nvram file to account for high score levels, and other settings
'   Set bumpers, and created flupper style.
'   Updated flipper postion, flipper triggers, flipper shadows, and physics
'   Updated flipper animations, nudging parameters, flipper physics and tricks
'   Added Bally style long gate prims and animated them.
'   Implemented full F12 menu system and tied everything to light room brightness
'   Adjusted room and cab brightness levels to be controlled via F12 menu for day/night
'   Added slings, sling corrections, and sling physics
'   Added new ball images and options and ball brightness code.
'   Added bumper brightness lit and not lit options to F12 menu
'   Added screw cap color, target color, peg coler, and flipper color options to F12 menu
'   Added ball options into the F12 menu based off the lotr's VPW table code
'   Updated the GI to come on after 2 seconds and stay on.
'   Set plunger.fire strength with a bit of randomness.
'   Set the bloom strength and screen space reflections scale to 0.2
'   Added a black wall under playfield to ensure cutouts have a black bottom.
'   Updated the GI to be controlled by not only GI, but also F12 lightlevel changes.
'   Wrote code and added a ball reflection mask for the saucer to eliminate playfield reflections.
'   Updated ball brightness area in plunger lane.
'   Corrected the audiopan and audiofade functions to raise to power of 5 instead of 10.  It wasn't panning correctly.
'   Updated the sounds for the slings SSF by unbalancing them.
'   Added top wall and image, antisnubber rails, lower rails, gate prims and gate
'   Added nvram file for replay thresholds, and instructions on how to change it.
'   Updated rollovers, slings, and animated them.
'   Added new Gate2 hood, gate, and gate rod.  Adjusted all gates and spinner.
'   Added all the cutout prims for GI bulbs, leafs, etc.
'   Added all physics elements for posts, rails, etc.
'   Added all rubbers, and adjusted the leaf sensors.
'   Added lower plastic guide rails in upper right orbit.
'   Added playfield mesh (with .76 edge width, and 3/16" drop.
'   Turned the playfield_mesh_vis to static to make sure you don't see insert holes.
'   Fine tuned the saucer kick and angle, with some variance.
'   Update pf mesh and mesh visible, made sure to have reflections on it.
'   Added Roth drop targets.  Updated target image.  Added code for DT shadows to ensure shadows go off on a solid hit only
'   Updated bumper strength and scatter angle slightly.
'   Adjusted physics sizes of guide rails to 5.
'   Updated standup targets to latest VPW code, and standalone capable, and updated target physics materials
'   Added bumper ring primitives and animated them.  Also animated the bumper skirts based off code from Beat the Clock by VPW
'   Updated saucer kick prim arm in the sacuer and properaly animated it to recess and timed kick ball out.
'   Tied the bumper lights to correct all lamps
'   Adjusted flipper physics to match pinballshenanigans and PAPA the best I could.  PS has them pointing right at the center insert.PAPA is at bottom. Went little below center.
'   Added a little targetbounce to the drop targets and lowered the center stand up target.
'   Added code for DT shadows to ensure shadows go off on a solid hit only
'   Tied the star lights to GI.
'   Created a new bumpercap image.  Changed transimission scale, modulation, intensity and warmth of bumper lights, as well as the bdl of the prim.
'   Tied bumper lights to all lamps and added logic to control the primitives
'   Cut holes in the playfield, and added all cutout prims
'   Added all the white and metal screws and metal post pegs,and rotated screws.
'   Added spinner and animated the the spinner and rod.
'   Updated the spinner spins to match online videos spinner duration.  Used code and logic similar to what I did on dragonfist.
'   Used slingshots from centaur and image.
'   Added an under physics wall beneath the flippers.
'   Added new desktop backglass and enabled the backglass reels, SA, BIP, Tilt, etc for Desktop Mode
'   Adjusted the slingshot hit and threshold values.  Too sensitive before.
'   Cut holes in the playfield for inserts.  Created borders around inserts, and placed Cooper Black Regular Font.  Created animal images for insert overlays.
'   Solid insert images, I put scratches on them like Centaur, but wen 15% brighter
'   Added 3D inserts with latest alllamps animation
'   Added insert wall blocks to eliminate seeing light bulbs through playfield
'   Changed the glow rings to have intensity of 12.
'   Ensured all rails guides were correct height.
'   Added bumper screws
'   Created plastic image based off of Bord's scanned assets
'   Created plastic walls, acrylics, and placed screws
'   Added LED plastics
'   Added GI tops to plastics.
'   Added GI tops for the upper lane guides
'   Updated the plastic post primitives.  Old one was blocky, and not very round.
'   Added basic GI
'   Performed a VLM bake of the GI Lighting. Had to remove the acrylics in the bake. They caused too much reflections in dark lanes and other areas on playfield.
'   Moved the lightmaps to 0.1 above the playfield.  Moved the light insert glows to 0.2 (above the GI)
'   Updated ALL the GI script to control everything.
'   Redbone updated the flipper, bumpercap, target, wall, and spinner images.
'   VR Backglass based on most recent Hauntfreaks backglass image
'   Added a backglass black background prim behind backglass flashers (flashers were too transparent at low opacity)
'   Added VR cabinet and VR Backglass images, text, and digits, and animated the backglass text
'   Updated ruleset in table info, and created a gamehelp image for the F12 rules
'   Raised the flipper prims to z height of 1.  Realistically, should be slightly above playfield.
'   Raised scaled the cabinet sideblades for cabmode option slightly.  Also removed the frontwall for cabmode.
'   Updated drop target shadow images.
'   Adjusted the dtshadow image opacity and height.
'   Added triggers and logic to enable metal hit sounds to closed gates (3 and 4).
'   Attempted several bakes of the spinner, but the insert in front is too far away and not strong enough for good bake.  However, added a little bdl on it when L40 is lit instead.
'   Redbone updated the apron and adjusted the wall plywood grain show through.
'   Adjusted color of apron walls to match apron and tied all to GI.
'   Updated the plunger cover image based off the apron colors.
'   Added code to DTAniate to turn prim.reflectionenabled true or false depending where it is.
'0.02 Baked the inner walls and have GI and insert reflections on them.
'   Removed unused images and materials.
'   Redbone added middle standup targets worn streaks image
'   Redbone added instruction cards (M-1508-95-F and M-1508-95-B).
'   Converted images to webp
'   Redbone upscaled the plastics image and swapped out the sling plastics with new images.
'   Redbone added noise to the apron for some texture.  Also added some exposure to the spinner image.
'   Changed the distance of bumper skirt animation to be total of 6 VPU... so bumper skirts are now z=-1.5 and than transz is -4.5... so total of -1.5 + -4.5 = -6)
'0.03 Changed the metal red peg material to clearcoat layer of 0,0,0
'   Adjusted the white flipper material to be more warm (antique white rgb), and adjusted the GI levels.
'   Increased the hit height of the saucer to 12, as on a very slow roll, it could just sit there and not roll in.
'   Adjusted the material on the credit light, the on intensity, and the height.
'   Adjusted the wire rail guide ends and inlane wires to have different elasticity and other physics on it.  Should have some spring to it.
'   Broke the inlane rail guide physics into two sections for different elasticity.
'   Widened the right inlane guide physics by one vpu to keep same size and distance from slings as on left.
'0.04 Adjusted the right inlane top portion to have more scatter angle.
'   Increased the green and yellow inserts on bdl.
'   Adjusted the materials and bdl of white inserts to be whiter off, and more warm in the on state.
'   Added more color to the plywood cutout material to light brown (196 164 132)
'   Increased the elasticity of the top lane posts.  It makes nudge play more important.
'   Increased elasticity of the right inlane area (more than left side to compensate for larger Distance)
'0.05 Made the green and yellow inserts more true to color.  Also lowered the white off insert brightness.
'0.06 iaakki: adjusted various physics regarding the drop targets; including physical angle of dt, mass, targetbouncer, elasticity, and friction.
'0.07 Uncle Paulie changed all the drop target physics and targetbouncer back to what it was, but ONLY left the left drop targets to what iaakki did for bounce and physics
'   Changed the angle of the dt a touch.
'   Changed dtmass of only the 3 left dt's from 0.2 to 0.02.
'   Redbone updated the apron graphics for higher rez, and texture.
'   Aligned the instruction cards.


'2.0.0  Released

' Thanks to Redbone for updating the playfield, plastics, targets, apron, and spinner images. Thanks to Bord for the playfield scan. Thanks to Hauntfreaks for providing backglass images
' Thanks to the iaakki for helping with the left drop target physics and the inlane guide rail physics to get the springiness like a real table.
' And thanks to the VPW team for testing and feedback.

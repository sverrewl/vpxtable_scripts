'
' _____  _____ _       ___  ______   ______ ___________ _____
'/  ___||  _  | |     / _ \ | ___ \  | ___ \_   _|  _  \  ___|
'\ `--. | | | | |    / /_\ \| |_/ /  | |_/ / | | | | | | |__
' `--. \| | | | |    |  _  ||    /   |    /  | | | | | |  __|
'/\__/ /\ \_/ / |____| | | || |\ \   | |\ \ _| |_| |/ /| |___
'\____/  \___/\_____/\_| |_/\_| \_|  \_| \_|\___/|___/ \____/
'
'
'     D. Gottlieb & Company (1977-1983), MPU Gottlieb System 1
'     Solar Ride / IPD No. 2239 / February, 1979 / 4 Players
'     https://www.ipdb.org/machine.cgi?id=2239
'     Production: 8,800 units
'
'
'********************************************************************************************************************************
'******** Work done by UnclePaulie on Hybrid version 0.01 - 2.0 *********
'     Implemented latest VPW standards, physics, GI VLM, 3D inserts, AllLamps routines, shadows, backglass for desktop,
'   roth targets, standalone compatibility, sling corrections, optional DIP adjustments, updated flipper physics,
'   VR and fully hybrid, playfield mesh, Flupper style bumpers and lighting. Used VR backglass images from Hauntfreaks.
'   Redbone updated the playfield graphics (from a cwgoss28 scan), as well other graphics: plastics, flippers, targets, and apron images.
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

const SetDIPSwitches= 0   ' If you want to set the dips differently, set to 1, and then hit F6 to launch.
const ForceVR = 0     ' Force VR mode on desktop, 0 is off (default): 1 is to force VR on.

' Insert brightness scale
const insertfactor  = 1   ' Adjusts the level of the brightness of the inserts when on

'User Defined acrylic edge colors (This is if user defined option is selected for plastic edge color... Choose your own RGB)

  Dim UserAcrylColor

  UserAcrylColor = RGB(255,255,240) 'default is RGB(255,255,240)


'******************************************************
' The default ROM menu has the replay threshold values to off.  The scorecard shows replay levels of 170,000, 350,000 and 430,000.  To set these
' follow the instructions below.  If you downloaded UnclePaulie's version off VPU, a .nvram file is included which has those settings
' preset to those levels to match the instruction card (left: B-18559-2;   Right: A-19135).

' Here is the procedure to set the selftest postions on this Gottlieb system 80 table for the preset replay levels:

' 1. Press 7 to enter adjustment mode
' 2. Continue to press 7 until you see 7 in the credit display.  This is the first replay level.
' 3. Press 1 to set the first replay level.  The score will increment by 10,000 points with each press of the 1 key. (Note: If you need to reset the score, release the 1 key then press the 9 key to set the score to zero then press 1 again to set the score to what you want.
' 4. When you reach the score you wish to set the first replay level to, press 7 to accept the change and advance to the 2nd replay level.
' 5. Repeat steps 3 and 4 to set the second replay level and advance to the 3rd replay level.
' 6. Repeat steps 3 and 4 to set the third replay level and advance to the high score to date.
' 7. Repeat steps 3 and 4 to set the high score to date and advance to the 1st and 3rd display check.
' 8. Press F3 to reset the table.  (I recommend to also restart the table)

' Here are the recommended values to change to:

  ' 7: 170,000 - First threshold to award a replay
  ' 8: 350,000 - Second threshold to award a replay
  ' 9: 430,000 - Third threshold to award a replay

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
' ZRVT: Vari Targets
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
Dim BIPL :  BIPL = False    'Ball in plunger lane
Dim GOn:  GOn = 0     ' Game is in play
Dim flbumpvalue1: flbumpvalue1 = 0
Dim flbumpvalue2: flbumpvalue2 = 0
Dim flbumpvalue3: flbumpvalue3 = 0
Dim gilvl
Dim PlasProt
Const rcolor = 211
Const gcolor = 200
Const bcolor = 169

Dim SRBall1, gBOT

'  Standard definitions
Const cGameName = "solaride"  'PinMAME ROM name
Const UseSolenoids = 2      '1 = Normal Flippers, 2 = Fastflips
Const UseLamps = 1        '0 = Custom lamp handling, 1 = Built-in VPX handling (using light number in light timer)
Const UseSync = 0
Const HandleMech = 0


Dim DT30up,DT31up,DT32up,DT33up,DT34up

LoadVPM"01150000","GTS1.VBS",3.22

'******************************************************
' VLM Arrays
'******************************************************

' VLM  Arrays - Start

' Arrays per baked part
Dim BP_Playfield: BP_Playfield=Array(LM_All_Lights_gi_001_Playfield, LM_All_Lights_gi_002_Playfield, LM_All_Lights_gi_003_Playfield, LM_All_Lights_gi_004_Playfield, LM_All_Lights_gi_005_Playfield, LM_All_Lights_gi_006_Playfield, LM_All_Lights_gi_007_Playfield, LM_All_Lights_gi_008_Playfield, LM_All_Lights_gi_009_Playfield, LM_All_Lights_gi_010_Playfield, LM_All_Lights_gi_011_Playfield, LM_All_Lights_gi_012_Playfield, LM_All_Lights_gi_013_Playfield, LM_All_Lights_gi_015_Playfield, LM_All_Lights_gi_016_Playfield, LM_All_Lights_gi_017_Playfield, LM_All_Lights_gi_018_Playfield)
Dim BP_Wood_Walls: BP_Wood_Walls=Array(BM_Wood_Walls, LM_All_Lights_gi_001_Wood_Walls, LM_All_Lights_gi_002_Wood_Walls, LM_All_Lights_gi_003_Wood_Walls, LM_All_Lights_gi_004_Wood_Walls, LM_All_Lights_gi_005_Wood_Walls, LM_All_Lights_gi_006_Wood_Walls, LM_All_Lights_gi_007_Wood_Walls, LM_All_Lights_gi_008_Wood_Walls, LM_All_Lights_gi_009_Wood_Walls, LM_All_Lights_gi_010_Wood_Walls, LM_All_Lights_gi_011_Wood_Walls, LM_All_Lights_gi_012_Wood_Walls, LM_All_Lights_gi_013_Wood_Walls, LM_All_Lights_gi_015_Wood_Walls, LM_All_Lights_gi_016_Wood_Walls, LM_All_Lights_gi_017_Wood_Walls, LM_All_Lights_gi_018_Wood_Walls)
Dim BP_MetalWalls: BP_MetalWalls=Array(BM_MetalWalls, LM_All_Lights_gi_001_MetalWalls, LM_All_Lights_gi_002_MetalWalls, LM_All_Lights_gi_003_MetalWalls, LM_All_Lights_gi_004_MetalWalls, LM_All_Lights_gi_005_MetalWalls, LM_All_Lights_gi_007_MetalWalls, LM_All_Lights_gi_008_MetalWalls, LM_All_Lights_gi_013_MetalWalls, LM_All_Lights_gi_015_MetalWalls, LM_All_Lights_gi_017_MetalWalls, LM_All_Lights_gi_018_MetalWalls, LM_All_Lights_l12_MetalWalls, LM_All_Lights_l14_MetalWalls)

' Arrays per lighting scenario
Dim BL_All_Lights_gi_001: BL_All_Lights_gi_001=Array(LM_All_Lights_gi_001_Playfield,LM_All_Lights_gi_001_Wood_Walls,LM_All_Lights_gi_001_MetalWalls)
Dim BL_All_Lights_gi_002: BL_All_Lights_gi_002=Array(LM_All_Lights_gi_002_Playfield,LM_All_Lights_gi_002_Wood_Walls,LM_All_Lights_gi_002_MetalWalls)
Dim BL_All_Lights_gi_003: BL_All_Lights_gi_003=Array(LM_All_Lights_gi_003_Playfield,LM_All_Lights_gi_003_Wood_Walls,LM_All_Lights_gi_003_MetalWalls)
Dim BL_All_Lights_gi_004: BL_All_Lights_gi_004=Array(LM_All_Lights_gi_004_Playfield,LM_All_Lights_gi_004_Wood_Walls,LM_All_Lights_gi_004_MetalWalls)
Dim BL_All_Lights_gi_005: BL_All_Lights_gi_005=Array(LM_All_Lights_gi_005_Playfield,LM_All_Lights_gi_005_Wood_Walls,LM_All_Lights_gi_005_MetalWalls)
Dim BL_All_Lights_gi_006: BL_All_Lights_gi_006=Array(LM_All_Lights_gi_006_Playfield,LM_All_Lights_gi_006_Wood_Walls)
Dim BL_All_Lights_gi_007: BL_All_Lights_gi_007=Array(LM_All_Lights_gi_007_Playfield,LM_All_Lights_gi_007_Wood_Walls,LM_All_Lights_gi_007_MetalWalls)
Dim BL_All_Lights_gi_008: BL_All_Lights_gi_008=Array(LM_All_Lights_gi_008_Playfield,LM_All_Lights_gi_008_Wood_Walls,LM_All_Lights_gi_008_MetalWalls)
Dim BL_All_Lights_gi_009: BL_All_Lights_gi_009=Array(LM_All_Lights_gi_009_Playfield,LM_All_Lights_gi_009_Wood_Walls)
Dim BL_All_Lights_gi_010: BL_All_Lights_gi_010=Array(LM_All_Lights_gi_010_Playfield,LM_All_Lights_gi_010_Wood_Walls)
Dim BL_All_Lights_gi_011: BL_All_Lights_gi_011=Array(LM_All_Lights_gi_011_Playfield,LM_All_Lights_gi_011_Wood_Walls)
Dim BL_All_Lights_gi_012: BL_All_Lights_gi_012=Array(LM_All_Lights_gi_012_Playfield,LM_All_Lights_gi_012_Wood_Walls)
Dim BL_All_Lights_gi_013: BL_All_Lights_gi_013=Array(LM_All_Lights_gi_013_Playfield,LM_All_Lights_gi_013_Wood_Walls,LM_All_Lights_gi_013_MetalWalls)
Dim BL_All_Lights_gi_015: BL_All_Lights_gi_015=Array(LM_All_Lights_gi_015_Playfield,LM_All_Lights_gi_015_Wood_Walls,LM_All_Lights_gi_015_MetalWalls)
Dim BL_All_Lights_gi_016: BL_All_Lights_gi_016=Array(LM_All_Lights_gi_016_Playfield,LM_All_Lights_gi_016_Wood_Walls)
Dim BL_All_Lights_gi_017: BL_All_Lights_gi_017=Array(LM_All_Lights_gi_017_Playfield,LM_All_Lights_gi_017_Wood_Walls,LM_All_Lights_gi_017_MetalWalls)
Dim BL_All_Lights_gi_018: BL_All_Lights_gi_018=Array(LM_All_Lights_gi_018_Playfield,LM_All_Lights_gi_018_Wood_Walls,LM_All_Lights_gi_018_MetalWalls)
Dim BL_World: BL_World=Array(BM_Wood_Walls,BM_MetalWalls)

' Global arrays
Dim BG_Bakemap: BG_Bakemap=Array(BM_Wood_Walls,BM_MetalWalls)
Dim BG_Lightmap: BG_Lightmap=Array(LM_All_Lights_gi_001_Playfield, LM_All_Lights_gi_002_Playfield, LM_All_Lights_gi_003_Playfield, LM_All_Lights_gi_004_Playfield, LM_All_Lights_gi_005_Playfield, LM_All_Lights_gi_006_Playfield, LM_All_Lights_gi_007_Playfield, LM_All_Lights_gi_008_Playfield, LM_All_Lights_gi_009_Playfield, LM_All_Lights_gi_010_Playfield, LM_All_Lights_gi_011_Playfield, LM_All_Lights_gi_012_Playfield, LM_All_Lights_gi_013_Playfield, LM_All_Lights_gi_015_Playfield, LM_All_Lights_gi_016_Playfield, LM_All_Lights_gi_017_Playfield, LM_All_Lights_gi_018_Playfield,_
LM_All_Lights_gi_001_Wood_Walls, LM_All_Lights_gi_002_Wood_Walls, LM_All_Lights_gi_003_Wood_Walls, LM_All_Lights_gi_004_Wood_Walls, LM_All_Lights_gi_005_Wood_Walls, LM_All_Lights_gi_006_Wood_Walls, LM_All_Lights_gi_007_Wood_Walls, LM_All_Lights_gi_008_Wood_Walls, LM_All_Lights_gi_009_Wood_Walls, LM_All_Lights_gi_010_Wood_Walls, LM_All_Lights_gi_011_Wood_Walls, LM_All_Lights_gi_012_Wood_Walls, LM_All_Lights_gi_013_Wood_Walls, LM_All_Lights_gi_015_Wood_Walls, LM_All_Lights_gi_016_Wood_Walls, LM_All_Lights_gi_017_Wood_Walls, LM_All_Lights_gi_018_Wood_Walls,_
LM_All_Lights_gi_001_MetalWalls, LM_All_Lights_gi_002_MetalWalls, LM_All_Lights_gi_003_MetalWalls, LM_All_Lights_gi_004_MetalWalls, LM_All_Lights_gi_005_MetalWalls, LM_All_Lights_gi_007_MetalWalls, LM_All_Lights_gi_008_MetalWalls, LM_All_Lights_gi_013_MetalWalls, LM_All_Lights_gi_015_MetalWalls, LM_All_Lights_gi_017_MetalWalls, LM_All_Lights_gi_018_MetalWalls, LM_All_Lights_l12_MetalWalls, LM_All_Lights_l14_MetalWalls)

Dim BG_All: BG_All=Array(LM_All_Lights_gi_001_Playfield, LM_All_Lights_gi_002_Playfield, LM_All_Lights_gi_003_Playfield, LM_All_Lights_gi_004_Playfield, LM_All_Lights_gi_005_Playfield, LM_All_Lights_gi_006_Playfield, LM_All_Lights_gi_007_Playfield, LM_All_Lights_gi_008_Playfield, LM_All_Lights_gi_009_Playfield, LM_All_Lights_gi_010_Playfield, LM_All_Lights_gi_011_Playfield, LM_All_Lights_gi_012_Playfield, LM_All_Lights_gi_013_Playfield, LM_All_Lights_gi_015_Playfield, LM_All_Lights_gi_016_Playfield, LM_All_Lights_gi_017_Playfield, LM_All_Lights_gi_018_Playfield,_
BM_Wood_Walls, LM_All_Lights_gi_001_Wood_Walls, LM_All_Lights_gi_002_Wood_Walls, LM_All_Lights_gi_003_Wood_Walls, LM_All_Lights_gi_004_Wood_Walls, LM_All_Lights_gi_005_Wood_Walls, LM_All_Lights_gi_006_Wood_Walls, LM_All_Lights_gi_007_Wood_Walls, LM_All_Lights_gi_008_Wood_Walls, LM_All_Lights_gi_009_Wood_Walls, LM_All_Lights_gi_010_Wood_Walls, LM_All_Lights_gi_011_Wood_Walls, LM_All_Lights_gi_012_Wood_Walls, LM_All_Lights_gi_013_Wood_Walls, LM_All_Lights_gi_015_Wood_Walls, LM_All_Lights_gi_016_Wood_Walls, LM_All_Lights_gi_017_Wood_Walls, LM_All_Lights_gi_018_Wood_Walls,_
BM_MetalWalls, LM_All_Lights_gi_001_MetalWalls, LM_All_Lights_gi_002_MetalWalls, LM_All_Lights_gi_003_MetalWalls, LM_All_Lights_gi_004_MetalWalls, LM_All_Lights_gi_005_MetalWalls, LM_All_Lights_gi_007_MetalWalls, LM_All_Lights_gi_008_MetalWalls, LM_All_Lights_gi_013_MetalWalls, LM_All_Lights_gi_015_MetalWalls, LM_All_Lights_gi_017_MetalWalls, LM_All_Lights_gi_018_MetalWalls, LM_All_Lights_l12_MetalWalls, LM_All_Lights_l14_MetalWalls)
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
  UpdateBallBrightness    'Control ball brightness in plunger lane
  RollingUpdate       'Update rolling sounds
  DoDTAnim          'Drop target animations
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
    .SplashInfoLine = "Solar Ride (Gottlieb 1979)"&chr(13)&"by UnclePaulie"
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
  vpmNudge.TiltSwitch = 4
  vpmNudge.Sensitivity = 3
  vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, RightSlingshot, LeftFlipper, RightFlipper, LeftFlipper1)

  'Trough - Creates a ball in the kicker switch and gives that ball used an individual name.
  Set SRBall1 = BallRelease.CreateSizedballWithMass(Ballsize/2,Ballmass)

  'Forces the trough switches to "on" at table boot so the game logic knows there are balls in the trough.
  Controller.Switch(66) = 1

  'Saucer initializations
  Controller.Switch(41) = 0
  Controller.Switch(42) = 0

  '***Setting up a ball array (collection), must contain all the balls you create on the table.
  gBOT = Array(SRBall1)

  ' GI initially off
  PFGI 0
  gilvl = 0
  BGGI 0

  if VR_Room = 1 Then
    setup_backglass()
  End If

' If user wants to set dip switches themselves it will force them to set it via F6.
  If SetDIPSwitches = 0 Then
    SetDefaultDips
  End If

' Initialize Bumpers
  FlFadeBumper 1, 0
  FlFadeBumper 2, 0
  FlFadeBumper 3, 0
  flbumpvalue1 = 0
  flbumpvalue2 = 0
  flbumpvalue3 = 0

' Make drop target shadows visible
  Dim xx
  for each xx in ShadowDT
    xx.visible=0
  Next

' Drop Target Variable state for DT Shadows
  DT30up=1
  DT31up=1
  DT32up=1
  DT33up=1
  DT34up=1

End Sub

Sub Table1_Paused:Controller.Pause = 1:End Sub
Sub Table1_unPaused:Controller.Pause = 0:End Sub
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
Dim ROMVolume                 ' Level of ROM volume. Value between 0 and 1
Dim WallClock, poster, poster2, CustomWalls
Dim ChooseBall: ChooseBall = 4
              '0 = dark, 1 = not as dark, 2 = bright, 3 = brightest, 4 = warm, 5 = MrBallDark, 6 = DarthVito Ball1
              '7 = DarthVito HDR1, 8 = DarthVito HDR Bright, 9 = DarthVito HDR2 (Default), 10 = DarthVito ballHDR3,
              '11 = DarthVito ballHDR4, 12 = DarthVito ballHDRdark, 13 = Borg Ball, 14 = SteelBall2


' Bumper brightness scale
dim bumperscale:    bumperscale =     0.22  ' adjusts the level of blenddisablelighting and intensity on bumpers when lights are not on
dim bumperflashlevel:   bumperflashlevel =  0.4   ' adjusts the level of bumper lights when on and / or flashing

Dim FlipperColor, ScrewCapColor, cabsideblades, PlasticColor, undercab, PegColor, ChimesorTones, Bumperflasher, SlingFlasher, BumperdiskColor, LaneguideColor, InnerCabColor

Dim dspTriggered : dspTriggered = False
Sub Table1_OptionEvent(ByVal eventId)
    If eventId = 1 And Not dspTriggered Then dspTriggered = True : DisableStaticPreRendering = True : End If

  ' ROM and Solenoid Sounds
  ROMVolume = Table1.Option("ROM Volume (-32 off; 0 max) - Requires Restart; altsound set to 0", -32, 0, 1, 0, 0)
  ChimesorTones = Table1.Option("Scoring Sounds: chimes or tones", 0, 2, 1, 0, 0, Array("None (default)", "Chimes", "Tones"))

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
  animatebumperlights

  ' General Options
  FlipperColor = Table1.Option("Flipper Color", 0, 3, 1, 0, 0, Array("Red (default)", "Yellow", "Black", "Orange"))
  PegColor = Table1.Option("Peg Color", 0, 4, 1, 4, 0, Array("Acrylic", "Red", "Yellow", "Blue", "White (default)"))
  PlasticColor = Table1.Option("Plastic Edge Color", 0, 6, 1, 0, 0, Array("Off (default)", "Orange/Red", "Yellow", "Blue", "User Defined in Script", "Rainbow", "Random"))
  LaneguideColor = Table1.Option("Lane Guide Color", 0, 4, 1, 0, 0, Array("White/Red (default)", "All White", "All Red", "All Yellow", "Yellow/Red"))
  BumperdiskColor = Table1.Option("Bumper Disk Color", 0, 1, 1, 0, 0, Array("Red (default)", "White"))
  BumperFlasher = Table1.Option("Bumper Flashers Option", 0, 1, 1, 0, 0, Array("Original No Flash (default)", "Flasher On When Bumper Hit"))
  SlingFlasher = Table1.Option("Sling Flashers Option", 0, 1, 1, 0, 0, Array("Original No Flash (default)", "Flasher On When Sling Hit"))
  ScrewCapColor = Table1.Option("Screw Cap Color", 0, 5, 1, 3, 0, Array("White", "Black", "Bronze", "Silver (default)", "Blue/Green", "Gold"))
  InnerCabColor = Table1.Option("Inside Cab Side Wall Color", 0, 1, 1, 0, 0, Array("Yellow (default)", "Black"))
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
  lflipb.RotZ = a
  lflipr.RotZ = a
End Sub

Sub RightFlipper_Animate
  dim a: a = RightFlipper.CurrentAngle
  FlipperRSh.RotZ = a
  rflipb.RotZ = a
  rflipr.RotZ = a
End Sub

Sub LeftFlipper1_Animate
  dim a: a = LeftFlipper1.CurrentAngle
  FlipperLSh1.RotZ = a
  lflipb1.RotZ = a
  lflipr1.RotZ = a
End Sub

Sub Gate_001_Animate
  dim a: a = (Gate_001.CurrentAngle * 0.5) + 25
  GateFlap_001.Rotz = a
End Sub

Sub Gate_002_Animate
  dim a: a = -Gate_002.CurrentAngle
    GateU2Wire_002.Rotx = a
  pScoreSpring.TransX = sin( (Gate_002.CurrentAngle+180) * (2*PI/360)) * 5
  pScoreSpring.TransY = sin( (Gate_002.CurrentAngle- 90) * (2*PI/360)) * 5
End Sub

' Gate Metal Hit sounds when gates are closed

Sub Gate_001_Trig_sound_hit() ' doesn't need a sound trigger to enable, as it starts in closed position
  If Gate_001.open = false then
    RandomSoundMetal
  End If
End Sub

' Bumper Animations

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
Const PLLeft = 850        'X position of plunger lane left
Const PLRight = 952       'X position of plunger lane right
Const PLTop = 480         'Y position of plunger lane top
Const PLBottom = 1900       'Y position of plunger lane bottom
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
Dim x

' VR Room

  For each x in VRRoomLL: x.blenddisablelighting = 1 * (LightLevel/lightlevelinitial): Next
  For each x in VRCabLL: x.blenddisablelighting = 0.5 * (LightLevel/lightlevelinitial): Next
  For each x in LLVRButtons: x.blenddisablelighting = 0.1125 * (LightLevel/lightlevelinitial): Next
  VRposter.blenddisablelighting = 0.5 * (LightLevel/lightlevelinitial)
  VRposter2.blenddisablelighting = 0.5 * (LightLevel/lightlevelinitial)
  VR_Clockface.opacity = 100 * (LightLevel/lightlevelinitial)
  BGBright.opacity = 25 + 75 * (LightLevel/lightlevelinitial)
  BGDark.opacity = 50 * LightLevel/lightlevelinitial
  For each x in TopHardware: x.blenddisablelighting = 0.05 * (LightLevel/lightlevelinitial): Next
  UndercabLight.intensity = 15 * gilvl

  if gilvl > 0 AND undercab = 1 Then
    VRNewFloor.opacity = 100 * gilvl
  Elseif gilvl > 0 AND undercab = 0 Then
    VRNewFloor.opacity = 100 * (LightLevel/lightlevelinitial)
  Else
    VRNewFloor.opacity = 100 * (LightLevel/lightlevelinitial)
  End If

' GI and GI Bulb Prims
  For each x in GI: x.State = gilvl: Next
  For each x in GIBulbPrims: x.blenddisablelighting = (4 * (gilvl)): Next

' rubbers

  UpdateMaterial "Rubber White",0,0,0,0,0,0,1,RGB(((gilvl * 44) + rcolor),((gilvl * 40) + gcolor),((gilvl * 71) + bcolor)),0,0,False,True,0,0,0,0

' flippers
  Lflipb.blenddisablelighting = (0.0375 * (gilvl)) + (0.1125 * LightLevel/lightlevelinitial)
  Lflipb1.blenddisablelighting = (0.0375 * (gilvl)) + (0.1125 * LightLevel/lightlevelinitial)
  Rflipb.blenddisablelighting = (0.0375 * (gilvl)) + (0.1125 * LightLevel/lightlevelinitial)
  Lflipr.blenddisablelighting = (0.0375 * (gilvl)) + (0.1125 * LightLevel/lightlevelinitial)
  Lflipr1.blenddisablelighting = (0.0375 * (gilvl)) + (0.1125 * LightLevel/lightlevelinitial)
  Rflipr.blenddisablelighting = (0.0375 * (gilvl)) + (0.1125 * LightLevel/lightlevelinitial)

' targets
  For each x in GIdtTargets: x.blenddisablelighting = (0.055 * gilvl) + (0.02 * LightLevel/lightlevelinitial): Next

' prims (Star Triggers)
  For each x in GIStars: x.blenddisablelighting = 0.03 * gilvl * insertfactor + (LightLevel/lightlevelinitial * 0.005): Next
  For each x in GIpStars: x.blenddisablelighting = 1.98 * gilvl * insertfactor + (LightLevel/lightlevelinitial * 0.02): Next

' prims
  ApronPrim.blenddisablelighting = ((0.0495 * gilvl) + (0.005 * LightLevel/lightlevelinitial))
  Instruction_Card.blenddisablelighting = (0.0495 * gilvl) + (0.005 * LightLevel/lightlevelinitial)
  Score_Card.blenddisablelighting = (0.0495 * gilvl) + (0.005 * LightLevel/lightlevelinitial)
  PlungerCoverPrim.blenddisablelighting = ((0.0495 * gilvl) + (0.005 * LightLevel/lightlevelinitial))
  LaneGuard.blenddisablelighting = 0.025 * gilvl * LightLevel/lightlevelinitial
  RubberWheel.blenddisablelighting = (0.05 * gilvl) + (0.001 * LightLevel/lightlevelinitial)
  For each x in GIScrewCaps: x.blenddisablelighting = ((0.025 * gilvl) + (0.025 * LightLevel/lightlevelinitial)): Next
  OuterPrimBlack_DT.blenddisablelighting = (0.0625 * gilvl) + (0.0375 * LightLevel/lightlevelinitial)
  For each x in GICutoutPrims: x.blenddisablelighting = (0.25 * (gilvl)) + (0.1 * LightLevel/lightlevelinitial): Next
  For each x in GICutoutHoles: x.blenddisablelighting = (0.2 * (gilvl)) + (0.2 * LightLevel/lightlevelinitial): Next
  For each x in GILaneGuides: x.blenddisablelighting = (0.125 * (gilvl)) + (0.05 * LightLevel/lightlevelinitial) : Next
  For each x in GIMetal: x.blenddisablelighting = (0.015 * gilvl) + (0.01 * LightLevel/lightlevelinitial): Next
  For each x in GIMetalRails: x.blenddisablelighting = (0.008 * gilvl) + (0.007 * LightLevel/lightlevelinitial): Next
  For each x in GIMetalPosts: x.blenddisablelighting = (0.005 * gilvl) + (0.005 * LightLevel/lightlevelinitial): Next
  For each x in GILeafswitch: x.blenddisablelighting = ((0.2 * gilvl) + (0.005 * LightLevel/lightlevelinitial)): Next
  Sling1.blenddisablelighting = (0.2 * gilvl) + (0.005 * LightLevel/lightlevelinitial)
  ArchPrim.blenddisablelighting = 0.05 * gilvl * LightLevel/lightlevelinitial
  ArchRail.blenddisablelighting = (0.15 * gilvl) + (0.01 * LightLevel/lightlevelinitial)

  If PegColor = 0 Then
    For each x in GIPegs: x.blenddisablelighting = (0.06 * gilvl) + (0.015 * LightLevel/lightlevelinitial) : Next
  Elseif PegColor = 4 Then
    For each x in GIPegs: x.blenddisablelighting = (0.05 * gilvl) + (0.0325 * LightLevel/lightlevelinitial) : Next
  Else
    For each x in GIPegs: x.blenddisablelighting = (0.05 * gilvl) + (0.075 * LightLevel/lightlevelinitial) : Next
  End If

' Bumpers - top two center are controlled in bumper animation sub
  bumperGI_animate

' Saucers
  animatesaucerlights

' GI lightmaps
  For each x in LMGI: x.opacity = (75 * gilvl): Next

' Baked Walls
  For each x in BP_Wood_Walls: x.opacity = (85 * gilvl): Next
  For each x in BP_MetalWalls: x.opacity = (65 * gilvl): Next
  For each x in BP_Wood_Walls: x.blenddisablelighting = (LightLevel/lightlevelinitial) : Next
  For each x in BP_MetalWalls: x.blenddisablelighting = (LightLevel/lightlevelinitial) : Next

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
    StartButton.y = 2798.14 - 5
    SoundStartButton
  End If

  If KeyDownHandler(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal keycode)   '***What to do when a button is released***
  If keycode = PlungerKey Then
    Plunger.firespeed = RndInt (98,102)
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
    StartButton.y = 2798.14
  End If

  If KeyUpHandler(keycode) Then Exit Sub
End Sub


'******************************************************
' ZSOL: Solenoids & Flashers
'******************************************************

SolCallback(1)  = "SolBallRelease"    ' Trough to Plunger Lane
SolCallback(2)  = "SolKnocker"      ' Knocker solenoid
SolCallback(3)  = "SolSound3"     ' 10's   Sounds (Chimes/Tones/Alt/None)
SolCallback(4)  = "SolSound4"     ' 100's  Sounds (Chimes/Tones/Alt/None)
SolCallback(5)  = "SolSound5"     ' 1000's Sounds (Chimes/Tones/Alt/None)
SolCallback(6) =  "SolSaucerRight"    ' Right Saucer kick
SolCallback(7)  = "SolSaucerLeft"     ' Left Saucer kick
SolCallback(8)  = "DTBankReset"     ' Droptargets reset

SolCallback(17) = "SolGameOn"     ' Game is in play.  Used on this table for enabling chimes/tones

SolCallback(sLRFlipper) = "SolRFlipper" 'Right Flipper
SolCallback(sLLFlipper) = "SolLFlipper" 'Left Flipper


' Solenoid Sounds

Sub SolGameOn (Enabled)
  If Enabled Then
    GOn = 1
  Else
    GOn = 0
  End If
End Sub

Sub SolSound3 (Enabled)
  if Enabled = True AND GOn = 1 Then
    if ChimesorTones = 1 Then
      PlaySoundAtLevelStatic SoundFX("SJ_Chime_10a",DOFChimes), 1, ChimesPosition
    Elseif ChimesorTones = 2 Then
      PlaySoundAtLevelStatic SoundFX("10tone",DOFChimes), 1, ChimesPosition
    Else
      ' no sounds
    End If
  End If
End Sub

Sub SolSound4 (Enabled)
  if Enabled = True AND GOn = 1 Then
    if ChimesorTones = 1 Then
      PlaySoundAtLevelStatic SoundFX("SJ_Chime_100a",DOFChimes), 1, ChimesPosition
    Elseif ChimesorTones = 2 Then
      PlaySoundAtLevelStatic SoundFX("100tone",DOFChimes), 1, ChimesPosition
    Else
      ' no sounds
    End If
  End If
End Sub

Sub SolSound5 (Enabled)
  if Enabled = True AND GOn = 1 Then
    if ChimesorTones = 1 Then
      PlaySoundAtLevelStatic SoundFX("SJ_Chime_1000a",DOFChimes), 1, ChimesPosition
    Elseif ChimesorTones = 2 Then
      PlaySoundAtLevelStatic SoundFX("1000tone",DOFChimes), 1, ChimesPosition
    Else
      ' no sounds
    End If
  End If
End Sub

'******************************************************
' ZDRN: Drain, Trough, and Ball Release
'******************************************************

'********************* TROUGH *************************

Sub BallRelease_Hit   : Controller.Switch(66) = 1 : End Sub
Sub BallRelease_UnHit : RandomSoundBallRelease BallRelease : Controller.Switch(66) = 0 : End Sub


Sub UpdateTrough
  UpdateTroughTimer.Interval = 300
  UpdateTroughTimer.Enabled = 1
End Sub

Sub UpdateTroughTimer_Timer
  If BallRelease.BallCntOver = 0 Then Drain.kick 60, 16
  Me.Enabled = 0
End Sub


'*****************  DRAIN & RELEASE  ******************

Sub SolBallRelease(enabled)
  If enabled Then
    BallRelease.kick 57, 10
    RandomSoundBallRelease BallRelease
  End If
End Sub

Sub Drain_Hit()
  RandomSoundDrain Drain
  UpdateTrough
End Sub


'******************************************************
' ZFLP: FLIPPERS
'******************************************************
Const ReflipAngle = 20

Sub SolLFlipper(Enabled)
  If Enabled Then
    FlipperActivate LeftFlipper, LFPress
    FlipperActivate LeftFlipper1, LFPress1
    LF.Fire
    LF1.Fire
    If leftflipper.currentangle < leftflipper.endangle + ReflipAngle Then
      RandomSoundReflipUpLeft LeftFlipper
    Else
      SoundFlipperUpAttackLeft LeftFlipper
      RandomSoundFlipperUpLeft LeftFlipper
    End If
  Else
    FlipperDeActivate LeftFlipper, LFPress
    FlipperDeActivate LeftFlipper1, LFPress1
    LeftFlipper.RotateToStart
    LeftFlipper1.RotateToStart
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

Sub LeftFlipper1_Collide(parm)
  CheckLiveCatch Activeball, LeftFlipper1, LFCount1, parm
  LF1.ReProcessBalls ActiveBall
  LeftFlipperCollide parm
End Sub


'************************************************************
' ZSLG: Slingshot Animations
'************************************************************

Dim RStep

Sub RightSlingShot_Slingshot
  RS.VelocityCorrect(ActiveBall)
  vpmTimer.PulseSw 13
  DOF 203, DOFPulse
  RandomSoundSlingshotRight SLING1
    RSling.Visible = 0
    RSling1.Visible = 1
    Sling1.objroty = -10
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.objroty = -6: if SlingFlasher = 1 Then slingrightflash.state = 1
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.objroty = 0:RightSlingShot.TimerEnabled = 0: vpmTimer.Addtimer 100, "slingrightflash.state = 0'"
    End Select
    RStep = RStep + 1
End Sub


'******************************************************
' ZSSC: SLINGSHOT CORRECTION FUNCTIONS by apophis
'******************************************************
' To add these slingshot corrections:
'  - On the table, add the endpoint primitives that define the two ends of the Slingshot
'  - Initialize the SlingshotCorrection objects in InitSlingCorrection
'  - Call the .VelocityCorrect methods from the respective _Slingshot event sub

Dim RS: Set RS = New SlingshotCorrection

InitSlingCorrection

Sub InitSlingCorrection

  RS.Object = RightSlingshot
  RS.EndPoint1 = EndPoint1RS
  RS.EndPoint2 = EndPoint2RS

  'Slingshot angle corrections (pt, BallPos in %, Angle in deg)
  ' These values are best guesses. Retune them if needed based on specific table research.
  AddSlingsPt 0, 0.00,    -3
  AddSlingsPt 1, 0.30,    -5
  AddSlingsPt 2, 0.40,  -30
  AddSlingsPt 3, 0.60,  30
  AddSlingsPt 4, 0.70,  5
  AddSlingsPt 5, 1.00,  3

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


'************************* Bumpers **************************
Sub Bumper1_Hit()
  RandomSoundBumperTop Bumper1
  vpmTimer.PulseSw 43
  DOF 201, 2
  if BumperFlasher = 1 Then
    bumper1flash.state = 1
    vpmTimer.AddTimer 100, "bumper1flash.state = 0'"
  end If
End Sub

Sub Bumper2_Hit()
  RandomSoundBumperTop Bumper2
  vpmTimer.PulseSw 43
  DOF 202, 2
  if BumperFlasher = 1 Then
    bumper2flash.state = 1
    vpmTimer.AddTimer 100, "bumper2flash.state = 0'"
  end If
End Sub

Sub Bumper3_Hit()
  RandomSoundBumperMiddle Bumper3
  vpmTimer.PulseSw 40
  if BumperFlasher = 1 Then
    bumper3flash.state = 1
    vpmTimer.AddTimer 100, "bumper3flash.state = 0'"
  end If
End Sub


'********************* Leaf Sensors **********************

Sub Phys_Sensor_001_hit: vpmTimer.PulseSw 13: RightUpRubber.enabled = true: End Sub
Sub Phys_Sensor_002_hit: vpmTimer.PulseSw 13: RightMidRubber.enabled = true: End Sub
Sub Phys_Sensor_003_hit: vpmTimer.PulseSw 13: LeftMidRubber.enabled = true: End Sub
Sub Phys_Sensor_004_hit
  vpmTimer.PulseSw 13
  if activeball.velx > 4 Then
    LeftRubber.enabled = true
  End If
End Sub

Sub Phys_Sensor_005_hit
  vpmTimer.PulseSw 13
    'debug.print "ball vel x: " & activeball.velx
  if activeball.velx > 10 Then
    LeftMidRubber2.enabled = true
  elseif activeball.velx > 4 Then
    LeftMidRubber3.enabled = true
  End If
End Sub

Sub Phys_Sensor_006_hit: vpmTimer.PulseSw 13: LeftUpRubber.enabled = true: End Sub
Sub Phys_Sensor_007_hit: TopRubber.enabled = true: End Sub
Sub Phys_Sensor_008_hit: RightRubber.enabled = true: End Sub
Sub Phys_Sensor_009_hit: LaneRubber.enabled = true: End Sub

' Leaf Sensor Rubber movement

dim LeftRubberCounter : LeftRubberCounter = 0
sub LeftRubber_Timer

  Select Case LeftRubberCounter
    Case 1:
      ' rubber in
      Rubber_Left.visible = 0
      Rubber_Lefta.visible = 1
      Rubber_Leftb.visible = 0
      LeafSwitch_004.objroty = -8
      if SlingFlasher = 1 Then slingleftflash.state = 1
    Case 3:
      ' rubber out
      Rubber_Left.visible = 0
      Rubber_Lefta.visible = 0
      Rubber_Leftb.visible = 1
      LeafSwitch_004.objroty = 5
      vpmTimer.Addtimer 100, "slingleftflash.state = 0'"
    Case Else:
      ' default rubber position
      Rubber_Left.visible = 1
      Rubber_Lefta.visible = 0
      Rubber_Leftb.visible = 0
      LeafSwitch_004.objroty = 0
  End Select

  if LeftRubberCounter >= 4 Then
    LeftRubberCounter = 0
    LeftRubber.enabled = false
  Else
    LeftRubberCounter = LeftRubberCounter + 1
  end If
end Sub

dim RightUpRubberCounter : RightUpRubberCounter = 0
sub RightUpRubber_Timer

  Select Case RightUpRubberCounter
    Case 1:
      ' rubber in
      Rubber_UpRight.visible = 0
      Rubber_UpRighta.visible = 1
      Rubber_UpRightb.visible = 0
      LeafSwitch_001.objroty = 8
    Case 3:
      ' rubber out
      Rubber_UpRight.visible = 0
      Rubber_UpRighta.visible = 0
      Rubber_UpRightb.visible = 1
      LeafSwitch_001.objroty = -5
    Case Else:
      ' default rubber position
      Rubber_UpRight.visible = 1
      Rubber_UpRighta.visible = 0
      Rubber_UpRightb.visible = 0
      LeafSwitch_001.objroty = 0
  End Select

  if RightUpRubberCounter >= 4 Then
    RightUpRubberCounter = 0
    RightUpRubber.enabled = false
  Else
    RightUpRubberCounter = RightUpRubberCounter + 1
  end If
end Sub


dim LeftUpRubberCounter : LeftUpRubberCounter = 0
sub LeftUpRubber_Timer

  Select Case LeftUpRubberCounter
    Case 1:
      ' rubber in
      Rubber_UpLeft.visible = 0
      Rubber_UpLefta.visible = 1
      Rubber_UpLeftb.visible = 0
      LeafSwitch_006.objroty = -5
    Case 3:
      ' rubber out
      Rubber_UpLeft.visible = 0
      Rubber_UpLefta.visible = 0
      Rubber_UpLeftb.visible = 1
      LeafSwitch_006.objroty = 3
    Case Else:
      ' default rubber position
      Rubber_UpLeft.visible = 1
      Rubber_UpLefta.visible = 0
      Rubber_UpLeftb.visible = 0
      LeafSwitch_006.objroty = 0
  End Select

  if LeftUpRubberCounter >= 4 Then
    LeftUpRubberCounter = 0
    LeftUpRubber.enabled = false
  Else
    LeftUpRubberCounter = LeftUpRubberCounter + 1
  end If
end Sub

dim LeftMidRubberCounter : LeftMidRubberCounter = 0
sub LeftMidRubber_Timer

  Select Case LeftMidRubberCounter
    Case 1:
      ' rubber in
      Rubber_MidLeft.visible = 0
      Rubber_MidLefta.visible = 1
      Rubber_MidLeftb.visible = 0
      Rubber_MidLeftc.visible = 0
      Rubber_MidLeftd.visible = 0
      Rubber_MidLefte.visible = 0
      Rubber_MidLeftf.visible = 0
      LeafSwitch_003.objroty = -4
    Case 3:
      ' rubber out
      Rubber_MidLeft.visible = 0
      Rubber_MidLefta.visible = 0
      Rubber_MidLeftb.visible = 1
      Rubber_MidLeftc.visible = 0
      Rubber_MidLeftd.visible = 0
      Rubber_MidLefte.visible = 0
      Rubber_MidLeftf.visible = 0
      LeafSwitch_003.objroty = 2
    Case Else:
      ' default rubber position
      Rubber_MidLeft.visible = 1
      Rubber_MidLefta.visible = 0
      Rubber_MidLeftb.visible = 0
      Rubber_MidLeftc.visible = 0
      Rubber_MidLeftd.visible = 0
      Rubber_MidLefte.visible = 0
      Rubber_MidLeftf.visible = 0
      LeafSwitch_003.objroty = 0
  End Select

  if LeftMidRubberCounter >= 4 Then
    LeftMidRubberCounter = 0
    LeftMidRubber.enabled = false
  Else
    LeftMidRubberCounter = LeftMidRubberCounter + 1
  end If
end Sub


dim LeftMidRubber2Counter : LeftMidRubber2Counter = 0
sub LeftMidRubber2_Timer

  Select Case LeftMidRubber2Counter
    Case 1:
      ' rubber in
      Rubber_MidLeft.visible = 0
      Rubber_MidLefta.visible = 0
      Rubber_MidLeftb.visible = 0
      Rubber_MidLeftc.visible = 1
      Rubber_MidLeftd.visible = 0
      Rubber_MidLefte.visible = 0
      Rubber_MidLeftf.visible = 0
      LeafSwitch_005.objroty = -8
    Case 3:
      ' rubber out
      Rubber_MidLeft.visible = 0
      Rubber_MidLefta.visible = 0
      Rubber_MidLeftb.visible = 0
      Rubber_MidLeftc.visible = 0
      Rubber_MidLeftd.visible = 1
      Rubber_MidLefte.visible = 0
      Rubber_MidLeftf.visible = 0
      LeafSwitch_005.objroty = 5
    Case Else:
      ' default rubber position
      Rubber_MidLeft.visible = 1
      Rubber_MidLefta.visible = 0
      Rubber_MidLeftb.visible = 0
      Rubber_MidLeftc.visible = 0
      Rubber_MidLeftd.visible = 0
      Rubber_MidLefte.visible = 0
      Rubber_MidLeftf.visible = 0
      LeafSwitch_005.objroty = 0
  End Select

  if LeftMidRubber2Counter >= 4 Then
    LeftMidRubber2Counter = 0
    LeftMidRubber2.enabled = false
  Else
    LeftMidRubber2Counter = LeftMidRubber2Counter + 1
  end If
end Sub


dim LeftMidRubber3Counter : LeftMidRubber3Counter = 0
sub LeftMidRubber3_Timer

  Select Case LeftMidRubber3Counter
    Case 1:
      ' rubber in
      Rubber_MidLeft.visible = 0
      Rubber_MidLefta.visible = 0
      Rubber_MidLeftb.visible = 0
      Rubber_MidLeftc.visible = 0
      Rubber_MidLeftd.visible = 0
      Rubber_MidLefte.visible = 1
      Rubber_MidLeftf.visible = 0
      LeafSwitch_005.objroty = -5
    Case 3:
      ' rubber out
      Rubber_MidLeft.visible = 0
      Rubber_MidLefta.visible = 0
      Rubber_MidLeftb.visible = 0
      Rubber_MidLeftc.visible = 0
      Rubber_MidLeftd.visible = 0
      Rubber_MidLefte.visible = 0
      Rubber_MidLeftf.visible = 1
      LeafSwitch_005.objroty = 3
    Case Else:
      ' default rubber position
      Rubber_MidLeft.visible = 1
      Rubber_MidLefta.visible = 0
      Rubber_MidLeftb.visible = 0
      Rubber_MidLeftc.visible = 0
      Rubber_MidLeftd.visible = 0
      Rubber_MidLefte.visible = 0
      Rubber_MidLeftf.visible = 0
      LeafSwitch_005.objroty = 0
  End Select

  if LeftMidRubber3Counter >= 4 Then
    LeftMidRubber3Counter = 0
    LeftMidRubber3.enabled = false
  Else
    LeftMidRubber3Counter = LeftMidRubber3Counter + 1
  end If
end Sub



dim RightMidRubberCounter : RightMidRubberCounter = 0
sub RightMidRubber_Timer

  Select Case RightMidRubberCounter
    Case 1:
      ' rubber in
      Rubber_MidRight.visible = 0
      Rubber_MidRighta.visible = 1
      Rubber_MidRightb.visible = 0
      LeafSwitch_002.objroty = -6
    Case 3:
      ' rubber out
      Rubber_MidRight.visible = 0
      Rubber_MidRighta.visible = 0
      Rubber_MidRightb.visible = 1
      LeafSwitch_002.objroty = 3
    Case Else:
      ' default rubber position
      Rubber_MidRight.visible = 1
      Rubber_MidRighta.visible = 0
      Rubber_MidRightb.visible = 0
      LeafSwitch_002.objroty = 0
  End Select

  if RightMidRubberCounter >= 4 Then
    RightMidRubberCounter = 0
    RightMidRubber.enabled = false
  Else
    RightMidRubberCounter = RightMidRubberCounter + 1
  end If
end Sub


dim TopRubberCounter : TopRubberCounter = 0
sub TopRubber_Timer

  Select Case TopRubberCounter
    Case 1:
      ' rubber in
      Rubber_Top.visible = 0
      Rubber_Topa.visible = 1
      Rubber_Topb.visible = 0
    Case 3:
      ' rubber out
      Rubber_Top.visible = 0
      Rubber_Topa.visible = 0
      Rubber_Topb.visible = 1
    Case Else:
      ' default rubber position
      Rubber_Top.visible = 1
      Rubber_Topa.visible = 0
      Rubber_Topb.visible = 0
  End Select

  if TopRubberCounter >= 4 Then
    TopRubberCounter = 0
    TopRubber.enabled = false
  Else
    TopRubberCounter = TopRubberCounter + 1
  end If
end Sub


dim RightRubberCounter : RightRubberCounter = 0
sub RightRubber_Timer

  Select Case RightRubberCounter
    Case 1:
      ' rubber in
      Rubber_Right.visible = 0
      Rubber_Righta.visible = 1
      Rubber_Rightb.visible = 0
    Case 3:
      ' rubber out
      Rubber_Right.visible = 0
      Rubber_Righta.visible = 0
      Rubber_Rightb.visible = 1
    Case Else:
      ' default rubber position
      Rubber_Right.visible = 1
      Rubber_Righta.visible = 0
      Rubber_Rightb.visible = 0
  End Select

  if RightRubberCounter >= 4 Then
    RightRubberCounter = 0
    RightRubber.enabled = false
  Else
    RightRubberCounter = RightRubberCounter + 1
  end If
end Sub


dim LaneRubberCounter : LaneRubberCounter = 0
sub LaneRubber_Timer

  Select Case LaneRubberCounter
    Case 1:
      ' rubber in
      Rubber_Lane.visible = 0
      Rubber_Lanea.visible = 1
      Rubber_Laneb.visible = 0
    Case 3:
      ' rubber out
      Rubber_Lane.visible = 0
      Rubber_Lanea.visible = 0
      Rubber_Laneb.visible = 1
    Case Else:
      ' default rubber position
      Rubber_Lane.visible = 1
      Rubber_Lanea.visible = 0
      Rubber_Laneb.visible = 0
  End Select

  if LaneRubberCounter >= 4 Then
    LaneRubberCounter = 0
    LaneRubber.enabled = false
  Else
    LaneRubberCounter = LaneRubberCounter + 1
  end If
end Sub


'************************ Rollovers *************************

Sub swBIPL_Hit(): BIPL = 1: End Sub 'Plunger Lane Rollover Switch
Sub swBIPL_UnHit: BIPL = 0: End Sub

Sub sw10_Hit(): Controller.Switch(10) = 1: End Sub
Sub sw10_UnHit: Controller.Switch(10) = 0: End Sub
Sub sw10a_Hit(): Controller.Switch(10) = 1: End Sub
Sub sw10a_UnHit: Controller.Switch(10) = 0: End Sub
Sub sw11_Hit(): Controller.Switch(11) = 1: End Sub
Sub sw11_UnHit: Controller.Switch(11) = 0: End Sub
Sub sw20_Hit(): Controller.Switch(20) = 1: End Sub
Sub sw20_UnHit: Controller.Switch(20) = 0: End Sub
Sub sw21_Hit(): Controller.Switch(21) = 1: End Sub
Sub sw21_UnHit: Controller.Switch(21) = 0: End Sub
Sub sw22_Hit(): Controller.Switch(22) = 1: End Sub
Sub sw22_UnHit: Controller.Switch(22) = 0: End Sub
Sub sw23_Hit(): Controller.Switch(23) = 1: End Sub
Sub sw23_UnHit: Controller.Switch(23) = 0: End Sub


'*******************************************
' Rollover Animations
'*******************************************

Sub sw10_Animate: psw10.transz = sw10.CurrentAnimOffset: End Sub
Sub sw10a_Animate:  psw10a.transz = sw10a.CurrentAnimOffset: End Sub
Sub sw11_Animate: psw11.transz = sw11.CurrentAnimOffset: End Sub
Sub sw20_Animate: psw20.transz = sw20.CurrentAnimOffset: End Sub
Sub sw21_Animate: psw21.transz = sw21.CurrentAnimOffset: End Sub
Sub sw22_Animate: psw22.transz = sw22.CurrentAnimOffset: End Sub
Sub sw23_Animate: psw23.transz = sw23.CurrentAnimOffset: End Sub


'*******************************************
'Star Rollovers
'*******************************************

Sub sw12_Hit:vpmTimer.PulseSw 12:me.timerenabled = 0:AnimateStar star12, sw12, 1:End Sub
Sub sw12_UnHit:me.timerinterval = 7:me.timerenabled = 1:End Sub
Sub sw12_timer:AnimateStar star12, sw12, 0:End Sub

Sub sw122_Hit:vpmTimer.PulseSw 12:me.timerenabled = 0:AnimateStar star122, sw122, 1:End Sub
Sub sw122_UnHit:me.timerinterval = 7:me.timerenabled = 1:End Sub
Sub sw122_timer:AnimateStar star122, sw122, 0:End Sub

Sub sw123_Hit:vpmTimer.PulseSw 12:me.timerenabled = 0:AnimateStar star123, sw123, 1:End Sub
Sub sw123_UnHit:me.timerinterval = 7:me.timerenabled = 1:End Sub
Sub sw123_timer:AnimateStar star123, sw123, 0:End Sub

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
'  Saucer
'*******************************************

'Random Values for saucer kick and angles
Dim RNDKickForce41
Dim RNDKickAngle41
Dim RNDKickForce42
Dim RNDKickAngle42

Sub SolSaucerRight(Enabled)
  If Enabled Then
  If sw41.BallCntOver = 1 Then
    RNDKickForce41 = RndInt(14,15)
    RNDKickAngle41 = RndInt(235,237)
    sw41.kick RNDKickAngle41,RNDKickForce41
    sw41.uservalue=1
    sw41.timerenabled=1
    PkickarmSW41.rotz=15
  End If
  End if
End Sub

Sub sw41_timer
  select case sw41.uservalue
    case 2:
      PkickarmSW41.rotz=0
      me.timerenabled=0
  end Select
  sw41.uservalue=sw41.uservalue+1
End Sub

Sub sw41_Hit
  SoundSaucerLock
  Controller.Switch(41) = 1
End Sub

Sub sw41_Unhit
  SoundSaucerKick 1, sw41
  controller.switch(41) = 0
End Sub


Sub SolSaucerLeft(Enabled)
  If Enabled Then
  If sw42.BallCntOver = 1 Then
    RNDKickForce42 = RndInt(16,17)
    RNDKickAngle42 = RndInt(179,181)
    sw42.kick RNDKickAngle42,RNDKickForce42
    sw42.uservalue=1
    sw42.timerenabled=1
    PkickarmSW42.rotz=15
  End If
  End if
End Sub

Sub sw42_timer
  select case sw42.uservalue
    case 2:
      PkickarmSW42.rotz=0
      me.timerenabled=0
  end Select
  sw42.uservalue=sw42.uservalue+1
End Sub

Sub sw42_Hit
  SoundSaucerLock
  Controller.Switch(42) = 1
End Sub

Sub sw42_Unhit
  SoundSaucerKick 1, sw42
  controller.switch(42) = 0
End Sub


'*******************************************
'  Block Reflections over saucer
'*******************************************

sub BallReflectionMask41_Hit() : table1.BallPlayfieldReflectionScale = 0 : end Sub
sub BallReflectionMask41_UnHit() : table1.BallPlayfieldReflectionScale = 1 : end Sub
sub BallReflectionMask42_Hit() : table1.BallPlayfieldReflectionScale = 0 : end Sub
sub BallReflectionMask42_UnHit() : table1.BallPlayfieldReflectionScale = 1 : end Sub


'*******************************************
'  Gates
'*******************************************

Sub Gate_002_Trigger_Hit:vpmTimer.PulseSw(52):End Sub

'*******************************************
'  Knocker
'*******************************************

Sub SolKnocker(Enabled)
  If enabled Then
    KnockerSolenoid
  End If
End Sub

'*******************************************
'  Drop and Stand Up Targets
'*******************************************

'********************** Drop Targets ************************

Sub sw30_Hit: DTHit 30: TargetBouncer Activeball, 1.0: DT30up = 0: End Sub
Sub sw31_Hit: DTHit 31: TargetBouncer Activeball, 1.0: DT31up = 0: End Sub
Sub sw32_Hit: DTHit 32: TargetBouncer Activeball, 1.0: DT32up = 0: End Sub
Sub sw33_Hit: DTHit 33: TargetBouncer Activeball, 1.0: DT33up = 0: End Sub
Sub sw34_Hit: DTHit 34: TargetBouncer Activeball, 1.0: DT34up = 0: End Sub


'********************************************
' Drop Target Solenoid Controls
'********************************************

const dtdelaytime = 0

Sub DTBankReset (enabled)
  If enabled then
  vpmTimer.AddTimer dtdelaytime, "RandomSoundDropTargetReset sw32p'"
    vpmTimer.AddTimer dtdelaytime, "DTRaise 30'"
    vpmTimer.AddTimer dtdelaytime, "DTRaise 31'"
    vpmTimer.AddTimer dtdelaytime, "DTRaise 32'"
    vpmTimer.AddTimer dtdelaytime, "DTRaise 33'"
    vpmTimer.AddTimer dtdelaytime, "DTRaise 34'"
    vpmTimer.AddTimer dtdelaytime, "DT30up=1'"
    vpmTimer.AddTimer dtdelaytime, "DT31up=1'"
    vpmTimer.AddTimer dtdelaytime, "DT32up=1'"
    vpmTimer.AddTimer dtdelaytime, "DT33up=1'"
    vpmTimer.AddTimer dtdelaytime, "DT34up=1'"
    vpmTimer.AddTimer dtdelaytime, "dtsh30.visible=True'"
    vpmTimer.AddTimer dtdelaytime, "dtsh31.visible=True'"
    vpmTimer.AddTimer dtdelaytime, "dtsh32.visible=True'"
    vpmTimer.AddTimer dtdelaytime, "dtsh33.visible=True'"
    vpmTimer.AddTimer dtdelaytime, "dtsh34.visible=True'"
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
Dim LF1 : Set LF1 = New FlipperPolarity


InitPolarity

'
''*******************************************
'' Late 70's to early 80's

Sub InitPolarity()
   dim x, a : a = Array(LF, RF, LF1)
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
    LF1.SetObjects "LF1", LeftFlipper1, TriggerLF1
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
  FlipperNudge RightFlipper, RFEndAngle, RFEOSNudge, LeftFlipper, LFEndAngle
  FlipperNudge LeftFlipper, LFEndAngle, LFEOSNudge,  RightFlipper, RFEndAngle
  FlipperNudge LeftFlipper1, LFEndAngle, LFEOSNudge,  RightFlipper, RFEndAngle
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
    'Check left1 flipper
    If LeftFlipper1.currentangle = LFEndAngle1 Then
    If FlipperTrigger(ball1.x, ball1.y, LeftFlipper1) OR FlipperTrigger(ball2.x, ball2.y, LeftFlipper1) Then DoDamping = true
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

Dim LFPress, RFPress, LFCount, RFCount, LFPress1, LFCount1
Dim LFState, RFState, LFState1
Dim EOST, EOSA,Frampup, FElasticity,FReturn
Dim RFEndAngle, LFEndAngle, LFEndAngle1

Const FlipperCoilRampupMode = 0 '0 = fast, 1 = medium, 2 = slow (tap passes should work)

LFState = 1
RFState = 1
LFState1 = 1
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
LFEndAngle1 = Leftflipper1.endangle

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
Dim DT30, DT31, DT32, DT33, DT34

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

Set DT30 = (new DropTarget)(sw30, sw30a, sw30p, 30, 0, false)
Set DT31 = (new DropTarget)(sw31, sw31a, sw31p, 31, 0, false)
Set DT32 = (new DropTarget)(sw32, sw32a, sw32p, 32, 0, false)
Set DT33 = (new DropTarget)(sw33, sw33a, sw33p, 33, 0, false)
Set DT34 = (new DropTarget)(sw34, sw34a, sw34p, 34, 0, false)


Dim DTArray
DTArray = Array(DT30, DT31, DT32, DT33, DT34)

'Configure the behavior of Drop Targets.
Const DTDropSpeed = 90 'in milliseconds
Const DTDropUpSpeed = 40 'in milliseconds
Const DTDropUnits = 49 'VP units primitive drops so top of at or below the playfield
Const DTDropUpUnits = 10 'VP units primitive raises above the up position on drops up
Const DTMaxBend = 2 '8 'max degrees primitive rotates when hit
Const DTDropDelay = 20 'time in milliseconds before target drops (due to friction/impact of the ball)
Const DTRaiseDelay = 40 'time in milliseconds before target drops back to normal up position after the solenoid fires to raise the target
Const DTBrickVel = 30 'velocity at which the target will brick, set to '0' to disable brick
Const DTEnableBrick = 0 '0 'Set to 0 to disable bricking, 1 to enable bricking

Const DTHitSound = "" 'Drop Target Hit sound
Const DTDropSound = "DropTarget_Down" 'Drop Target Drop sound
Const DTResetSound = "DropTarget_Up" 'Drop Target reset sound

Const DTMass = 0.1 'Mass of the Drop Target (between 0 and 1), higher values provide more resistance

'******************************************************
'  DROP TARGETS FUNCTIONS
'******************************************************

Dim DTShadow(5)

DTShadowInit 1
DTShadowInit 2
DTShadowInit 3
DTShadowInit 4
DTShadowInit 5

' Initializes the drop targets for shadow logic below
Sub DTShadowInit(dtnbr)

  if dtnbr = 1 Then
    Set DTShadow(dtnbr) = Eval("dtsh" & 30)
  elseif dtnbr = 2 Then
    Set DTShadow(dtnbr) = Eval("dtsh" & 31)
  elseif dtnbr = 3 Then
    Set DTShadow(dtnbr) = Eval("dtsh" & 32)
  elseif dtnbr = 4 Then
    Set DTShadow(dtnbr) = Eval("dtsh" & 33)
  elseif dtnbr = 5 Then
    Set DTShadow(dtnbr) = Eval("dtsh" & 34)
  End If
End Sub

Sub DTHit(switch)
  Dim i, swmod

  If switch = 30 Then
    swmod = 1
  Elseif switch = 31 then
    swmod = 2
  Elseif switch = 32 then
    swmod = 3
  Elseif switch = 33 then
    swmod = 4
  Elseif switch = 34 then
    swmod = 5
  End If

  i = DTArrayID(switch)

  PlayTargetSound
  DTArray(i).animate = DTCheckBrick(ActiveBall,DTArray(i).prim)
  If DTArray(i).animate = 1 Or DTArray(i).animate = 3 Or DTArray(i).animate = 4 Then
    DTBallPhysics ActiveBall, DTArray(i).prim.rotz, DTMass
' Controls Drop Shadow for a direct hit only
    if swmod=1 or swmod=2 or swmod=3 or swmod=4 or swmod=5 then
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


'****************************************************************
'****  END VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'****************************************************************


'******************************************************
'   Z3DI:   3D INSERTS
'******************************************************

dim insert_white_on: insert_white_on = insertfactor * 30
dim insert_white_arrow_on: insert_white_arrow_on = insertfactor * 20
const insert_white_off = 1.5
dim insert_red_on: insert_red_on = insertfactor * 16
const insert_red_off = 2
dim insert_wine_on: insert_wine_on = insertfactor * 24
const insert_wine_off = 2
dim insert_orange_on: insert_orange_on = insertfactor * 5
const insert_orange_off = 1
dim insert_yellow_on: insert_yellow_on = insertfactor * 12
const insert_yellow_off = 1.5
dim insert_green_on: insert_green_on = insertfactor * 30
const insert_green_off = 0.5

const a_insert_white_on = 0.75
const a_insert_white_off = 0.95
const a_insert_red_on = 1.5
const a_insert_red_off = 2
const a_insert_wine_on = 1.5
const a_insert_wine_off = 2
const a_insert_orange_on = 0.5
const a_insert_orange_off = 0.105
const a_insert_yellow_on = 1
const a_insert_yellow_off = 1
const a_insert_green_on = 1.0
const a_insert_green_off = 1.25

const bulb_on = 10
const bulb_off = 1

const insert_credit_on = 0.15
const insert_credit_off = 0.1


'************************
'**** White Inserts

Sub l12_animate
  p12.BlendDisableLighting = insert_white_off + insert_white_on * (l12.GetInPlayIntensity / l12.Intensity)
  p12off.BlendDisableLighting = a_insert_white_off - a_insert_white_on * (l12.GetInPlayIntensity / l12.Intensity)
  bulb12.BlendDisableLighting = bulb_off + bulb_on * (l12.GetInPlayIntensity / l12.Intensity)
End Sub

Sub l14_animate
  p14.BlendDisableLighting = insert_white_off + insert_white_on * (l14.GetInPlayIntensity / l14.Intensity)
  p14off.BlendDisableLighting = a_insert_white_off - a_insert_white_on * (l14.GetInPlayIntensity / l14.Intensity)
  bulb14.BlendDisableLighting = bulb_off + bulb_on * (l14.GetInPlayIntensity / l14.Intensity)
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

Sub l25_animate
  p25.BlendDisableLighting = insert_white_off + insert_white_on * (l25.GetInPlayIntensity / l25.Intensity)
  p25off.BlendDisableLighting = a_insert_white_off - a_insert_white_on * (l25.GetInPlayIntensity / l25.Intensity)
  bulb25.BlendDisableLighting = bulb_off + bulb_on * (l25.GetInPlayIntensity / l25.Intensity)
End Sub

Sub l26_animate
  p26.BlendDisableLighting = insert_white_off + insert_white_on * (l26.GetInPlayIntensity / l26.Intensity)
  p26off.BlendDisableLighting = a_insert_white_off - a_insert_white_on * (l26.GetInPlayIntensity / l26.Intensity)
  bulb26.BlendDisableLighting = bulb_off + bulb_on * (l26.GetInPlayIntensity / l26.Intensity)
End Sub

Sub l27_animate
  p27.BlendDisableLighting = insert_white_off + insert_white_on * (l27.GetInPlayIntensity / l27.Intensity)
  p27off.BlendDisableLighting = a_insert_white_off - a_insert_white_on * (l27.GetInPlayIntensity / l27.Intensity)
  bulb27.BlendDisableLighting = bulb_off + bulb_on * (l27.GetInPlayIntensity / l27.Intensity)
End Sub

Sub l28_animate
  p28.BlendDisableLighting = insert_white_off + insert_white_on * (l28.GetInPlayIntensity / l28.Intensity)
  p28off.BlendDisableLighting = a_insert_white_off - a_insert_white_on * (l28.GetInPlayIntensity / l28.Intensity)
  bulb28.BlendDisableLighting = bulb_off + bulb_on * (l28.GetInPlayIntensity / l28.Intensity)
End Sub

Sub l29_animate
  p29.BlendDisableLighting = insert_white_off + insert_white_on * (l29.GetInPlayIntensity / l29.Intensity)
  p29off.BlendDisableLighting = a_insert_white_off - a_insert_white_on * (l29.GetInPlayIntensity / l29.Intensity)
  bulb29.BlendDisableLighting = bulb_off + bulb_on * (l29.GetInPlayIntensity / l29.Intensity)
End Sub

Sub l30_animate
  p30.BlendDisableLighting = insert_white_off + insert_white_on * (l30.GetInPlayIntensity / l30.Intensity)
  p30off.BlendDisableLighting = a_insert_white_off - a_insert_white_on * (l30.GetInPlayIntensity / l30.Intensity)
  bulb30.BlendDisableLighting = bulb_off + bulb_on * (l30.GetInPlayIntensity / l30.Intensity)
End Sub

Sub l31_animate
  p31.BlendDisableLighting = insert_white_off + insert_white_on * (l31.GetInPlayIntensity / l31.Intensity)
  p31off.BlendDisableLighting = a_insert_white_off - a_insert_white_on * (l31.GetInPlayIntensity / l31.Intensity)
  bulb31.BlendDisableLighting = bulb_off + bulb_on * (l31.GetInPlayIntensity / l31.Intensity)
End Sub

Sub l32_animate
  p32.BlendDisableLighting = insert_white_off + insert_white_on * (l32.GetInPlayIntensity / l32.Intensity)
  p32off.BlendDisableLighting = a_insert_white_off - a_insert_white_on * (l32.GetInPlayIntensity / l32.Intensity)
  bulb32.BlendDisableLighting = bulb_off + bulb_on * (l32.GetInPlayIntensity / l32.Intensity)
End Sub

Sub l33_animate
  p33.BlendDisableLighting = insert_white_off + insert_white_on * (l33.GetInPlayIntensity / l33.Intensity)
  p33off.BlendDisableLighting = a_insert_white_off - a_insert_white_on * (l33.GetInPlayIntensity / l33.Intensity)
  bulb33.BlendDisableLighting = bulb_off + bulb_on * (l33.GetInPlayIntensity / l33.Intensity)
End Sub


'************************
'**** Yellow Inserts

Sub l13_animate
  p13.BlendDisableLighting = insert_yellow_off + insert_yellow_on * (l13.GetInPlayIntensity / l13.Intensity)
  p13off.BlendDisableLighting = a_insert_yellow_off - a_insert_yellow_on * (l13.GetInPlayIntensity / l13.Intensity)
  bulb13.BlendDisableLighting = bulb_off + bulb_on * (l13.GetInPlayIntensity / l13.Intensity)
End Sub

Sub l15_animate
  p15.BlendDisableLighting = insert_yellow_off + insert_yellow_on * (l15.GetInPlayIntensity / l15.Intensity)
  p15off.BlendDisableLighting = a_insert_yellow_off - a_insert_yellow_on * (l15.GetInPlayIntensity / l15.Intensity)
  bulb15.BlendDisableLighting = bulb_off + bulb_on * (l15.GetInPlayIntensity / l15.Intensity)
End Sub

Sub l16_animate
  p16.BlendDisableLighting = insert_yellow_off + insert_yellow_on * (l16.GetInPlayIntensity / l16.Intensity)
  p16off.BlendDisableLighting = a_insert_yellow_off - a_insert_yellow_on * (l16.GetInPlayIntensity / l16.Intensity)
  bulb16.BlendDisableLighting = bulb_off + bulb_on * (l16.GetInPlayIntensity / l16.Intensity)
End Sub

Sub l18_animate
  p18.BlendDisableLighting = insert_yellow_off + insert_yellow_on * (l18.GetInPlayIntensity / l18.Intensity)
  p18off.BlendDisableLighting = a_insert_yellow_off - a_insert_yellow_on * (l18.GetInPlayIntensity / l18.Intensity)
  bulb18.BlendDisableLighting = bulb_off + bulb_on * (l18.GetInPlayIntensity / l18.Intensity)
End Sub


'************************
'**** Green Inserts

Sub l5_animate
  p5.BlendDisableLighting = insert_green_off + insert_green_on * (l5.GetInPlayIntensity / l5.Intensity)
  p5off.BlendDisableLighting = a_insert_green_off - a_insert_green_on * (l5.GetInPlayIntensity / l5.Intensity)
  bulb5.BlendDisableLighting = bulb_off + bulb_on * (l5.GetInPlayIntensity / l5.Intensity)
End Sub

Sub l6_animate
  p6.BlendDisableLighting = insert_green_off + insert_green_on * (l6.GetInPlayIntensity / l6.Intensity)
  p6off.BlendDisableLighting = a_insert_green_off - a_insert_green_on * (l6.GetInPlayIntensity / l6.Intensity)
  bulb6.BlendDisableLighting = bulb_off + bulb_on * (l6.GetInPlayIntensity / l6.Intensity)
End Sub

Sub l7_animate
  p7.BlendDisableLighting = insert_green_off + insert_green_on * (l7.GetInPlayIntensity / l7.Intensity)
  p7off.BlendDisableLighting = a_insert_green_off - a_insert_green_on * (l7.GetInPlayIntensity / l7.Intensity)
  bulb7.BlendDisableLighting = bulb_off + bulb_on * (l7.GetInPlayIntensity / l7.Intensity)
End Sub

Sub l8_animate
  p8.BlendDisableLighting = insert_green_off + insert_green_on * (l8.GetInPlayIntensity / l8.Intensity)
  p8off.BlendDisableLighting = a_insert_green_off - a_insert_green_on * (l8.GetInPlayIntensity / l8.Intensity)
  bulb8.BlendDisableLighting = bulb_off + bulb_on * (l8.GetInPlayIntensity / l8.Intensity)
End Sub

Sub l10_animate
  p10.BlendDisableLighting = insert_green_off + insert_green_on * (l10.GetInPlayIntensity / l10.Intensity)
  p10off.BlendDisableLighting = a_insert_green_off - a_insert_green_on * (l10.GetInPlayIntensity / l10.Intensity)
  bulb10.BlendDisableLighting = bulb_off + bulb_on * (l10.GetInPlayIntensity / l10.Intensity)
End Sub


'************************
'**** Red Inserts

Sub l9_animate
  p9.BlendDisableLighting = insert_red_off + insert_red_on * (l9.GetInPlayIntensity / l9.Intensity)
  p9off.BlendDisableLighting = a_insert_red_off - a_insert_red_on * (l9.GetInPlayIntensity / l9.Intensity)
  bulb9.BlendDisableLighting = bulb_off + bulb_on * (l9.GetInPlayIntensity / l9.Intensity)
End Sub


'************************
'**** Wine / Pinkish / Red Inserts

Sub l4_animate
  p4.BlendDisableLighting = insert_wine_off + insert_wine_on * (l4.GetInPlayIntensity / l4.Intensity)
  p4off.BlendDisableLighting = a_insert_wine_off - a_insert_wine_on * (l4.GetInPlayIntensity / l4.Intensity)
  bulb4.BlendDisableLighting = bulb_off + bulb_on * (l4.GetInPlayIntensity / l4.Intensity)
End Sub

Sub l11_animate   'want brighter when on vs other wine insert
  p11.BlendDisableLighting = insert_red_off + insert_red_on * 2 * (l11.GetInPlayIntensity / l11.Intensity)
  p11off.BlendDisableLighting = a_insert_red_off - a_insert_red_on * (l11.GetInPlayIntensity / l11.Intensity)
  bulb11.BlendDisableLighting = bulb_off + bulb_on * (l11.GetInPlayIntensity / l11.Intensity)
End Sub


'************************
'**** Bumper Lamps

Sub bumperGI_animate
  flbumpvalue1 = bumperflashlevel * (bumperGI.GetInPlayIntensity / bumperGI.Intensity)
  flbumpvalue2 = bumperflashlevel * (bumperGI.GetInPlayIntensity / bumperGI.Intensity)
  flbumpvalue3 = bumperflashlevel * (bumperGI.GetInPlayIntensity / bumperGI.Intensity)
  FlFadeBumper 1, flbumpvalue1
  FlFadeBumper 2, flbumpvalue2
  FlFadeBumper 3, flbumpvalue3
End Sub

Sub animatebumperlights 'used when light level in room is adjusted
  bumperGI_animate
End Sub

'************************
'**** Saucer Lamps

Sub LightSw41_animate
  Pkickerhole1.blenddisablelighting = (0.075 * (LightSw41.GetInPlayIntensity / LightSw41.Intensity)) + (0.005 * LightLevel/lightlevelinitial)
  PkickarmSW41.blenddisablelighting = (0.025 * (LightSw41.GetInPlayIntensity / LightSw41.Intensity)) + (0.005 * LightLevel/lightlevelinitial)
End Sub

Sub LightSw42_animate
  Pkickerhole2.blenddisablelighting = (0.075 * (LightSw42.GetInPlayIntensity / LightSw41.Intensity)) + (0.005 * LightLevel/lightlevelinitial)
  PkickarmSW42.blenddisablelighting = (0.025 * (LightSw42.GetInPlayIntensity / LightSw41.Intensity)) + (0.005 * LightLevel/lightlevelinitial)
End Sub

Sub animatesaucerlights
  LightSw41_animate
  LightSw41_animate
End Sub

'******************************************************
'*****   END 3D INSERTS
'******************************************************



'******************************************************
'****  ZGIU:  GI Control
'******************************************************

sub PFGI (Enabled)

Dim x

  if Enabled = 1 Then

    If gilvl = 0 Then: Sound_GI_Relay 1, Bumper3

  ' Set Drop Target Shadows to state when GI is back on.
    if DT30up=1 Then dtsh30.visible = 1
    if DT31up=1 Then dtsh31.visible = 1
    if DT32up=1 Then dtsh32.visible = 1
    if DT33up=1 Then dtsh33.visible = 1
    if DT34up=1 Then dtsh34.visible = 1

    gilvl = 1   'Storing the latest GI fading state into global variable, so one can use it elsewhere too.
    BGGI 1      'Turn backglass GI lights on for desktop or VR.
    Else

    if gilvl = 1 Then: Sound_GI_Relay 0, bumper3

    ' Set Drop Target Shadows to off when GI is goes off.
      for each x in ShadowDT
        x.visible=0
      Next

  gilvl = 0   'Storing the latest GI fading state into global variable, so one can use it elsewhere too.

    End If

  SetRoomBrightness

End Sub


sub BGGI (Enabled)

  if Enabled = 1 Then

    If DesktopMode = True then
      L_DT_1.intensity = 7
      L_DT_1a.intensity = 5
      L_DT_2.intensity = 6
      L_DT_2a.intensity = 7
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

  Else
    Dim x

    for each x in dt_bg_lights: x.intensity = 0.1: Next

    if VR_Room = 1 Then
      BGDark.Visible = 1
      BGBright.Visible = 0
    Else
      BGDark.Visible = 0
      BGBright.Visible = 0
    End If
  End If
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

Dim FlBumperFadeActual(6), FlBumperFadeTarget(6), FlBumperColor(6), FlBumperTop(6), FlBumperCapBase(6), FlBumperSmallLight(6), Flbumperbiglight(6)
Dim FlBumperDisk(6), FlBumperBase(6), FlBumperBulb(6), FlBumperscrews(6), FlBumperActive(6), FlBumperHighlight(6), FLBumperlight(6), FlBumperRing(6)
Dim cnt
For cnt = 1 To 6
  FlBumperActive(cnt) = False
Next

' Note these are called now in SetupOptions, as bumper disk colors are an option switchable in the F12 menu
'FlInitBumper 1, "SolarRide"
'FlInitBumper 2, "SolarRide"
'FlInitBumper 3, "SolarRide"

Sub FlInitBumper(nr, col)
  FlBumperActive(nr) = True

  ' store all objects in an array for use in FlFadeBumper subroutine
  FlBumperFadeActual(nr) = 1
  FlBumperFadeTarget(nr) = 1.1
  FlBumperColor(nr) = col
  Set FlBumperTop(nr) = Eval("bumpertop" & nr)
  Set FlBumperCapBase(nr) = Eval("bumpercapbase" & nr)
  Set FlBumperBase(nr) = Eval("bumperbase" & nr)
  FlBumperTop(nr).material = "bumpertopmat" & nr
  FlBumperCapBase(nr).material = "bumpertopmata" & nr
  Set FlBumperSmallLight(nr) = Eval("bumpersmalllight" & nr)
  Set Flbumperbiglight(nr) = Eval("bumperbiglight" & nr)
  Set FlBumperDisk(nr) = Eval("bumperdisk" & nr)
  Set FlBumperBase(nr) = Eval("bumperbase" & nr)
  Set FlBumperBulb(nr) = Eval("bumperbulb" & nr)
  FlBumperBulb(nr).material = "bumperbulbmat" & nr
  Set FlBumperscrews(nr) = Eval("bumperscrews" & nr)
  FlBumperscrews(nr).material = "bumperscrew" & col
  Set FlBumperHighlight(nr) = Eval("bumperhighlight" & nr)
  Set FlBumperLight(nr) = Eval("bumperlight" & nr)
  Set FlBumperRing(nr) = Eval("BumperRing" & nr)

  ' set the color for the two VPX lights
  Select Case col

    Case "SRred"
      FlBumperSmallLight(nr).color = RGB(228,112,37)
      FlBumperSmallLight(nr).colorfull = RGB(239,192,112)
      FlBumperBigLight(nr).color = RGB(228,112,37)
      FlBumperBigLight(nr).colorfull = RGB(239,192,112)
      FlBumperHighlight(nr).color = RGB(239,192,112)
      FlBumperSmallLight(nr).TransmissionScale = 0.1
      FlBumperSmallLight(nr).BulbModulateVsAdd = 0.99
      FlBumperTop(nr).BlendDisableLighting = 0.005
      MaterialColor "bumpertopmat" & nr, RGB(239,192,112)
      FlBumperCapBase(nr).BlendDisableLighting = 0.005
      MaterialColor "bumpertopmata" & nr, RGB(239,192,112)
      MaterialColor "bumperdisk" & nr, RGB(255,0,0)
      FlBumperDisk(nr).BlendDisableLighting = 0.15
      FlBumperRing(nr).BlendDisableLighting = 0
      FLBumperLight(nr).intensity = 0.1

    Case "SRwhite"
      FlBumperSmallLight(nr).color = RGB(228,112,37)
      FlBumperSmallLight(nr).colorfull = RGB(239,192,112)
      FlBumperBigLight(nr).color = RGB(228,112,37)
      FlBumperBigLight(nr).colorfull = RGB(239,192,112)
      FlBumperHighlight(nr).color = RGB(239,192,112)
      FlBumperSmallLight(nr).TransmissionScale = 0.1
      FlBumperSmallLight(nr).BulbModulateVsAdd = 0.99
      FlBumperTop(nr).BlendDisableLighting = 0.005
      MaterialColor "bumpertopmat" & nr, RGB(239,192,112)
      FlBumperCapBase(nr).BlendDisableLighting = 0.005
      MaterialColor "bumpertopmata" & nr, RGB(239,192,112)
      MaterialColor "bumperdisk" & nr, RGB(103,102,85)
      FlBumperDisk(nr).BlendDisableLighting = 0.15
      FlBumperRing(nr).BlendDisableLighting = 0
      FLBumperLight(nr).intensity = 0.1

  End Select
End Sub

Sub FlFadeBumper(nr, Z)

  '   UpdateMaterial(string, float wrapLighting, float roughness, float glossyImageLerp, float thickness, float edge, float edgeAlpha, float opacity,
  '        OLE_COLOR base, OLE_COLOR glossy, OLE_COLOR clearcoat, VARIANT_BOOL isMetal, VARIANT_BOOL opacityActive,
  '        float elasticity, float elasticityFalloff, float friction, float scatterAngle) - updates all parameters of a material



  Select Case FlBumperColor(nr)

    Case "SRred"

    MaterialColor "bumpertopmat" & nr, RGB(239,198 - z*6,127 - Z*15)
    MaterialColor "bumpertopmata" & nr, RGB(239,198 - z*6,127 - Z*15)
    MaterialColor "bumperbulbmat" & nr, RGB(239 -z*11,192 - z*80,112 - Z*85)
    MaterialColor "bumperbase" & nr, RGB(239,228 - z * 36,202 - Z * 90)
    FlBumperTop(nr).BlendDisableLighting = bumperscale * ((0.3 * LightLevel/lightlevelinitial) + 0.45 * Z)
    FlBumperCapBase(nr).BlendDisableLighting = bumperscale * ((0.15 * LightLevel/lightlevelinitial) + 0.45 * Z)
    FlBumperRing(nr).BlendDisableLighting = bumperscale * ((0.025 * LightLevel/lightlevelinitial) + 0.1 * Z)
    FlBumperBulb(nr).BlendDisableLighting = bumperscale * ((10 * LightLevel/lightlevelinitial) + 2000 * Z)
    Flbumperbiglight(nr).intensity = bumperscale * (10 * Z )
    FlBumperHighlight(nr).opacity = bumperscale * (1000 * (Z^3))
    FlBumperSmallLight(nr).color = RGB(239 -z*11,192 - z*80,112 - Z*85)
    FlBumperSmallLight(nr).colorfull = RGB(239 -z*11,192 - z*80,112 - Z*85)
    FlBumperSmallLight(nr).intensity = bumperscale * (1000 * Z)
    FlBumperLight(nr).color = RGB(239 -z*11,192 - z*80,112 - Z*85)
    FlBumperLight(nr).colorfull = RGB(239 -z*11,192 - z*80,112 - Z*85)
    FLBumperLight(nr).intensity =  ((50 * Z)*bumperflashlevel)
    FlBumperDisk(nr).BlendDisableLighting = (bumperscale * 0.68 * LightLevel/lightlevelinitial) + (Z * 2 * bumperflashlevel)
    FlBumperBase(nr).BlendDisableLighting = (bumperscale * 2.25 * LightLevel/lightlevelinitial) + (2.5 * Z * bumperflashlevel)

    Case "SRwhite"

    MaterialColor "bumpertopmat" & nr, RGB(239,198 - z*6,127 - Z*15)
    MaterialColor "bumpertopmata" & nr, RGB(239,198 - z*6,127 - Z*15)
    MaterialColor "bumperbulbmat" & nr, RGB(239 -z*11,192 - z*80,112 - Z*85)
    MaterialColor "bumperbase" & nr, RGB(239,228 - z * 36,202 - Z * 90)
    FlBumperTop(nr).BlendDisableLighting = bumperscale * ((0.3 * LightLevel/lightlevelinitial) + 0.45 * Z)
    FlBumperCapBase(nr).BlendDisableLighting = bumperscale * ((0.15 * LightLevel/lightlevelinitial) + 0.45 * Z)
    FlBumperRing(nr).BlendDisableLighting = bumperscale * ((0.025 * LightLevel/lightlevelinitial) + 0.1 * Z)
    FlBumperBulb(nr).BlendDisableLighting = bumperscale * ((10 * LightLevel/lightlevelinitial) + 2000 * Z)
    Flbumperbiglight(nr).intensity = bumperscale * (10 * Z )
    FlBumperHighlight(nr).opacity = bumperscale * (1000 * (Z^3))
    FlBumperSmallLight(nr).color = RGB(239 -z*11,192 - z*80,112 - Z*85)
    FlBumperSmallLight(nr).colorfull = RGB(239 -z*11,192 - z*80,112 - Z*85)
    FlBumperSmallLight(nr).intensity = bumperscale * (1000 * Z)
    FlBumperLight(nr).color = RGB(239 -z*11,192 - z*80,112 - Z*85)
    FlBumperLight(nr).colorfull = RGB(239 -z*11,192 - z*80,112 - Z*85)
    FLBumperLight(nr).intensity =  ((50 * Z)*bumperflashlevel)
    FlBumperDisk(nr).BlendDisableLighting = (bumperscale * 0.68 * LightLevel/lightlevelinitial) + (Z * 2 * bumperflashlevel)
    FlBumperBase(nr).BlendDisableLighting = (bumperscale * 2.25 * LightLevel/lightlevelinitial) + (2.5 * Z * bumperflashlevel)

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

Dim Digits(28)

Digits(0)=Array(a001,a002,a003,a004,a005,a006,a007,LXM,a008)
Digits(1)=Array(a1,a2,a3,a4,a5,a6,a7,LXM,a8)
Digits(2)=Array(a9,a10,a11,a12,a13,a14,a15,LXM,a16)
Digits(3)=Array(a17,a18,a19,a20,a21,a22,a23,LXM,a24)
Digits(4)=Array(a25,a26,a27,a28,a29,a30,a31,LXM,a32)
Digits(5)=Array(a33,a34,a35,a36,a37,a38,a39,LXM,a40)
Digits(6)=Array(a009,a010,a011,a012,a013,a014,a015,LXM,a016)
Digits(7)=Array(a49,a50,a51,a52,a53,a54,a55,LXM,a56)
Digits(8)=Array(a57,a58,a59,a60,a61,a62,a63,LXM,a64)
Digits(9)=Array(a65,a66,a67,a68,a69,a70,a71,LXM,a72)
Digits(10)=Array(a73,a74,a75,a76,a77,a78,a79,LXM,a80)
Digits(11)=Array(a81,a82,a83,a84,a85,a86,a87,LXM,a88)
Digits(12)=Array(a017,a018,a019,a020,a021,a022,a023,LXM,a024)
Digits(13)=Array(a97,a98,a99,a100,a101,a102,a103,LXM,a104)
Digits(14)=Array(a105,a106,a107,a108,a109,a110,a111,LXM,a112)
Digits(15)=Array(a113,a114,a115,a116,a117,a118,a119,LXM,a120)
Digits(16)=Array(a121,a122,a123,a124,a125,a126,a127,LXM,a128)
Digits(17)=Array(a129,a130,a131,a132,a133,a134,a135,LXM,a136)
Digits(18)=Array(a025,a026,a027,a028,a029,a030,a031,LXM,a032)
Digits(19)=Array(a145,a146,a147,a148,a149,a150,a151,LXM,a152)
Digits(20)=Array(a153,a154,a155,a156,a157,a158,a159,LXM,a160)
Digits(21)=Array(a161,a162,a163,a164,a165,a166,a167,LXM,a168)
Digits(22)=Array(a169,a170,a171,a172,a173,a174,a175,LXM,a176)
Digits(23)=Array(a177,a178,a179,a180,a181,a182,a183,LXM,a184)

'Ball in Play and Credit displays

Digits(24)=Array(e00,e01,e02,e03,e04,e05,e06,LXM)
Digits(25)=Array(e10,e11,e12,e13,e14,e15,e16,LXM)
Digits(26)=Array(f00,f01,f02,f03,f04,f05,f06,LXM)
Digits(27)=Array(f10,f11,f12,f13,f14,f15,f16,LXM)


Sub DisplayTimer
  Dim ChgLED,ii,num,chg,stat,obj
  ChgLed = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
  If Not IsEmpty(ChgLED) Then
    For ii = 0 To UBound(chgLED)
      num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
      if (num < 28) then
        For Each obj In Digits(num)
          If cab_mode = 1 OR VR_Room =1 Then
            obj.intensity=0
          Else
            obj.intensity= 1.5
          End If
          If chg And 1 Then obj.State = stat And 1
          chg = chg\2 : stat = stat\2
        Next
            Else
                For Each obj In Digits(num)
                    If chg And 1 Then obj.State = stat And 1
                    chg = chg \ 2:stat = stat \ 2
                Next
      end If
    next
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

  for ix = 0 to 5
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

  for ix = 6 to 11
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

  for ix = 12 to 17
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

  for ix = 18 to 23
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

  for ix = 24 to 27
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


'*****************************************************************************************************
' VR LED Displays
'*****************************************************************************************************

dim DisplayColor
DisplayColor =  RGB(1,48,135)


Dim VRDigits(28)

VRDigits(0) = Array(LED1x0,LED1x1,LED1x2,LED1x3,LED1x4,LED1x5,LED1x6, n1, LED1x8)
VRDigits(1) = Array(LED2x0,LED2x1,LED2x2,LED2x3,LED2x4,LED2x5,LED2x6, n1, LED2x8)
VRDigits(2) = Array(LED3x0,LED3x1,LED3x2,LED3x3,LED3x4,LED3x5,LED3x6, n1, LED3x8)
VRDigits(3) = Array(LED4x0,LED4x1,LED4x2,LED4x3,LED4x4,LED4x5,LED4x6, n1, LED4x8)
VRDigits(4) = Array(LED5x0,LED5x1,LED5x2,LED5x3,LED5x4,LED5x5,LED5x6, n1, LED5x8)
VRDigits(5) = Array(LED6x0,LED6x1,LED6x2,LED6x3,LED6x4,LED6x5,LED6x6, n1, LED6x8)

VRDigits(6) = Array(LED8x0,LED8x1,LED8x2,LED8x3,LED8x4,LED8x5,LED8x6, n1, LED8x8)
VRDigits(7) = Array(LED9x0,LED9x1,LED9x2,LED9x3,LED9x4,LED9x5,LED9x6, n1, LED9x8)
VRDigits(8) = Array(LED10x0,LED10x1,LED10x2,LED10x3,LED10x4,LED10x5,LED10x6, n1, LED10x8)
VRDigits(9) = Array(LED11x0,LED11x1,LED11x2,LED11x3,LED11x4,LED11x5,LED11x6, n1, LED11x8)
VRDigits(10) = Array(LED12x0,LED12x1,LED12x2,LED12x3,LED12x4,LED12x5,LED12x6, n1, LED12x8)
VRDigits(11) = Array(LED13x0,LED13x1,LED13x2,LED13x3,LED13x4,LED13x5,LED13x6, n1, LED13x8)

VRDigits(12) = Array(LED1x000,LED1x001,LED1x002,LED1x003,LED1x004,LED1x005,LED1x006, n1, LED1x008)
VRDigits(13) = Array(LED1x100,LED1x101,LED1x102,LED1x103,LED1x104,LED1x105,LED1x106, n1, LED1x108)
VRDigits(14) = Array(LED1x200,LED1x201,LED1x202,LED1x203,LED1x204,LED1x205,LED1x206, n1, LED1x208)
VRDigits(15) = Array(LED1x300,LED1x301,LED1x302,LED1x303,LED1x304,LED1x305,LED1x306, n1, LED1x308)
VRDigits(16) = Array(LED1x400,LED1x401,LED1x402,LED1x403,LED1x404,LED1x405,LED1x406, n1, LED1x408)
VRDigits(17) = Array(LED1x500,LED1x501,LED1x502,LED1x503,LED1x504,LED1x505,LED1x506, n1, LED1x508)

VRDigits(18) = Array(LED2x000,LED2x001,LED2x002,LED2x003,LED2x004,LED2x005,LED2x006, n1, LED2x008)
VRDigits(19) = Array(LED2x100,LED2x101,LED2x102,LED2x103,LED2x104,LED2x105,LED2x106, n1, LED2x108)
VRDigits(20) = Array(LED2x200,LED2x201,LED2x202,LED2x203,LED2x204,LED2x205,LED2x206, n1, LED2x208)
VRDigits(21) = Array(LED2x300,LED2x301,LED2x302,LED2x303,LED2x304,LED2x305,LED2x306, n1, LED2x308)
VRDigits(22) = Array(LED2x400,LED2x401,LED2x402,LED2x403,LED2x404,LED2x405,LED2x406, n1, LED2x408)
VRDigits(23) = Array(LED2x500,LED2x501,LED2x502,LED2x503,LED2x504,LED2x505,LED2x506, n1, LED2x508)

'Ball in Play and Credit displays

VRDigits(24) = Array(LEDcx500,LEDcx501,LEDcx502,LEDcx503,LEDcx504,LEDcx505,LEDcx506, n1)
VRDigits(25) = Array(LEDdx600,LEDdx601,LEDdx602,LEDdx603,LEDdx604,LEDdx605,LEDdx606, n1)
VRDigits(26) = Array(LEDax300,LEDax301,LEDax302,LEDax303,LEDax304,LEDax305,LEDax306, n1)
VRDigits(27) = Array(LEDbx400,LEDbx401,LEDbx402,LEDbx403,LEDbx404,LEDbx405,LEDbx406, n1)



Sub VRDisplayTimer
  Dim ii, jj, obj, b, x
  Dim ChgLED,num, chg, stat
  ChgLED=Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED) Then
      For ii=0 To UBound(chgLED)
        num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)
        if num < 28 Then
          For Each obj In VRDigits(num)
   '                  If chg And 1 Then obj.visible=stat And 1    'if you use the object color for off; turn the display object visible to not visible on the playfield, and uncomment this line out.
            If chg And 1 Then FadeDisplay obj, stat And 1
            chg=chg\2 : stat=stat\2
          Next
        Else
          For Each obj In VRDigits(num)
            If chg And 1 Then obj.State = stat And 1
            chg = chg \ 2:stat = stat \ 2
          Next
        End If
      Next
    End If
End Sub


Sub FadeDisplay(object, onoff)
  If OnOff = 1 Then
    object.color = DisplayColor
    Object.Opacity = 5
  Else
    Object.Color = RGB(1,1,1)
    Object.Opacity = 1
  End If
End Sub

Sub InitDigits()
  dim tmp, x, xobj
  for x = 0 to uBound(VRDigits)
    if IsArray(VRDigits(x) ) then
      For each xobj in VRDigits(x)
        xobj.height = xobj.height + 0
        FadeDisplay xobj, 0
      next
    end If
  Next
End Sub



If VR_Room=1 Then
  InitDigits
End If

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

  If FlipperColor = 1 Then
    rflipr.material = "Rubber Yellow1"
    lflipr.material = "Rubber Yellow2"
    lflipr1.material = "Rubber Yellow3"
  Elseif FlipperColor = 2 Then
    rflipr.material = "Rubber Black1"
    lflipr.material = "Rubber Black2"
    lflipr1.material = "Rubber Black3"
  Elseif FlipperColor = 3 Then
    rflipr.material = "Rubber Orange1"
    lflipr.material = "Rubber Orange2"
    lflipr1.material = "Rubber Orange3"
  Else
    rflipr.material = "Red Rubber1"
    lflipr.material = "Red Rubber2"
    lflipr1.material = "Red Rubber3"
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

' Peg Color
  If PegColor  = 1  Then
    For each C in GIPegs
      C.material = "Metal Red Pegs"
    Next
  ElseIf PegColor  = 2  Then
    For each C in GIPegs
      C.material = "Metal Yellow Pegs"
    Next
  ElseIf PegColor  = 3  Then
    For each C in GIPegs
      C.material = "Metal Blue Pegs"
    Next
  ElseIf PegColor  = 4  Then
    For each C in GIPegs
      C.material = "White Pegs"
    Next
  Else
    For each C in GIPegs
      C.material = "Acrylic Posts"
    Next
  End If

' Bumper Disk Color
  If BumperdiskColor  = 1  Then
    bumperdisk1.image = "Flbumperdisk-left white"
    bumperdisk2.image = "Flbumperdisk-right white"
    bumperdisk3.image = "Flbumperdisk-center white"
      FlInitBumper 1, "SRwhite"
      FlInitBumper 2, "SRwhite"
      FlInitBumper 3, "SRwhite"
      animatebumperlights
  Else
    bumperdisk1.image = "Flbumperdisk-left red"
    bumperdisk2.image = "Flbumperdisk-right red"
    bumperdisk3.image = "Flbumperdisk-center red"
    FlInitBumper 1, "SRred"
    FlInitBumper 2, "SRred"
    FlInitBumper 3, "SRred"
    animatebumperlights
  End If

' Lane Guide Color
  If LaneguideColor = 1  Then
    For each C in GIlaneguides
      C.material = "Plastic Lane Guides White"
    Next
  Elseif LaneguideColor = 2 Then
    For each C in GIlaneguides
      C.material = "Plastic Lane Guides Red"
    Next
  Elseif LaneguideColor = 3 Then
    For each C in GIlaneguides
      C.material = "Plastic Lane Guides Yellow"
    Next
  Elseif LaneguideColor = 4 Then
    LaneGuidePlastic_001.material = "Plastic Lane Guides Red"
    LaneGuidePlastic_002.material = "Plastic Lane Guides Red"
    LaneGuidePlastic_003.material = "Plastic Lane Guides Red"
    LaneGuidePlastic_004.material = "Plastic Lane Guides Yellow"
    LaneGuidePlastic_005.material = "Plastic Lane Guides Yellow"
  Else
    LaneGuidePlastic_001.material = "Plastic Lane Guides Red"
    LaneGuidePlastic_002.material = "Plastic Lane Guides Red"
    LaneGuidePlastic_003.material = "Plastic Lane Guides Red"
    LaneGuidePlastic_004.material = "Plastic Lane Guides White"
    LaneGuidePlastic_005.material = "Plastic Lane Guides White"
  End If

' Interior Cab Side Color
  If InnerCabColor = 1 Then
    OuterPrimBlack_DT.image = "OuterPrim_DT_UV-Black"
  Else
    OuterPrimBlack_DT.image = "OuterPrim_DT_UV yellow"
  End If

SetupVRRoom

End Sub


'******************************************************
' DIP Switches
'******************************************************

'Gottlieb System 1 games with 3 tones only (or chimes).
'These are: Charlie's Angels, Cleopatra, Close Encounters, Count Down, Dragon, Joker Poker, Pinball Pool, Sinbad, Solar Ride.
'For the other System 1 games use the file called "Gottlieb System 1 with multi-mode sound.txt"
'editdips originally done by Inkochnito

Sub editDips
  if SetDIPSwitches = 1 Then
    Dim vpmDips:Set vpmDips = New cvpmDips
    With vpmDips
        .AddForm 700, 400, "System 1 (3 tones) - DIP switches"
        .AddFrame 205, 0, 190, "Maximum credits", &H00030000, Array("5 credits", 0, "8 credits", &H00020000, "10 credits", &H00010000, "15 credits", &H00030000) 'dip 17&18
        .AddFrame 0, 0, 190, "Coin chute control", &H00040000, Array("seperate", 0, "same", &H00040000)                                                          'dip 19
        .AddFrame 0, 46, 190, "Game mode", &H00000400, Array("extra ball", 0, "replay", &H00000400)                                                              'dip 11
        .AddFrame 0, 92, 190, "High game to date awards", &H00200000, Array("no award", 0, "3 replays", &H00200000)                                              'dip 22
        .AddFrame 0, 138, 190, "Balls per game", &H00000100, Array("5 balls", 0, "3 balls", &H00000100)                                                          'dip 9
        .AddFrame 0, 184, 190, "Tilt effect", &H00000800, Array("game over", 0, "ball in play only", &H00000800)                                                 'dip 12
        .AddChk 205, 80, 190, Array("Match feature", &H00000200)                                                                                                 'dip 10
        .AddChk 205, 95, 190, Array("Credits displayed", &H00001000)                                                                                             'dip 13
        .AddChk 205, 110, 190, Array("Play credit button tune", &H00002000)                                                                                      'dip 14
        .AddChk 205, 125, 190, Array("Play tones when scoring", &H00080000)                                                                                      'dip 20
        .AddChk 205, 140, 190, Array("Play coin switch tune", &H00400000)                                                                                        'dip 23
        .AddChk 205, 155, 190, Array("High game to date displayed", &H00100000)                                                                                  'dip 21
        .AddLabel 50, 240, 300, 20, "After hitting OK, press F3 to reset game with new settings."
        .ViewDips
    End With
  End If
End Sub
Set vpmShowDips = GetRef("editDips")


'**********************************************************************************************************
'**********************************************************************************************************

'Recommended 3 ball settings by Gottlieb; but purposefully choosing conservative play where applicable

'  *** This is the default dips settings ***

'1-4  0000  1 credit/ 1st coin; 2 credits for second coin chute #1
'5-8  0000  Credit/coins chutes 2 same as chute 1
'9    1     Balls per game on = 3; off = 5
'10   1     Match feature on
'11   1   Game Mode: Replay
'12   1   Tilt Effect: Ball in play only
'13   1     Credit display on
'14   0   Credit button tune off
'15   0   n/a
'16   0   n/a
'17-18  11    Max credits of 15
'19   1   Coin chute control:  1 and 2 the same.
'20   1   Tones when scoring
'21   1     High Score to Date displayed
'22   1     High Score to Date awards 3 credits
'23   0   Coin switch tune off
'24   0   n/a

Sub SetDefaultDips
  If SetDIPSwitches = 0 Then
    Controller.Dip(0) = 0   '8-1  = 00000000
    Controller.Dip(1) = 31  '16-9   = 00011111
    Controller.Dip(2) = 63  '24-17  = 00111111
    Controller.Dip(3) = 0
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
  VRCoinVisualBlock.visible = 0
  VRCoinVisualBlock.sidevisible = 0
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
  VRCoinVisualBlock.visible = 0
  VRCoinVisualBlock.sidevisible = 0
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
  VRCoinVisualBlock.visible = 1
  VRCoinVisualBlock.sidevisible = 1
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

if VR_Room = 0 AND cab_mode = 1 Then
  Set LampCallback = GetRef("UpdateCabLamps")
End If

Sub UpdateDTLamps()
  CreditsREEL.setValue(1)
  If Controller.Lamp(4) = 0 Then: ShootAgainReel.setValue(0): Else: ShootAgainReel.setValue(1)    'Shoot Again
  If Controller.Lamp(3) = 0 Then: HighScoreReel.setValue(0):  Else: HighScoreReel.setValue(1)     'High Score

  If Controller.Lamp(2) = 0 Then                                    'Tilt
    TiltReel.setValue(0)
    pfgi 1
  Else
    TiltReel.setValue(1)
    pfgi 0
  End If

  If Controller.Lamp(1) = 0 Then                                    'Game Over and BIP
    GameOverReel.setValue(1)
    BIPReel.setValue(0)
    MatchReel.setValue(1)
  Else
    GameOverReel.setValue(0)
    BIPReel.setValue(1)
    MatchReel.setValue(0)
  End If

End Sub


Sub UpdateVRLamps()
  If Controller.Lamp(4) = 0 Then: VR_ShootAgain.visible=0:  Else: VR_ShootAgain.visible=1       'Shoot again
  If Controller.Lamp(3) = 0 Then: VR_HighScore.visible=0:   Else: VR_HighScore.visible=1      'High Score

  If Controller.Lamp(2) = 0 Then                                    'Tilt
    VR_Tilt.visible=0
    pfgi 1
  Else
    VR_Tilt.visible=1
    pfgi 0
  End If


  If Controller.Lamp(1) = 0 Then                                    'Game Over and BIP
    VR_GameOver.visible=1
    VR_BIP.visible=0
    VR_Match.visible=1
  Else
    VR_GameOver.visible=0
    VR_BIP.visible=1
    VR_Match.visible=0
  End If

End Sub


Sub UpdateCabLamps()
  If Controller.Lamp(2) = 0 Then                                    'Tilt
    pfgi 1
  Else
    pfgi 0
  End If
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

'******** Work done by UnclePaulie on Hybrid version 0.01 - 2.0 *********
'v0.01
' INITIAL SETUP
'   Solar Ride playfield scan, clean up, and hole cutouts provided by CainArg and then scaled by Redbone. UP fixed the upper right star rollover hole and missing GI hole.
'   Started with UP's Buck Rogers table as baseline for script, physics, etc. to utilize latest VPW standards
'   Created base script to utilize latest VPW standards
'   Update the table1.init with verbiage for table/author and other settings.
'   Ensure table1.init stop/pause code correct.
'   Corrected the audiopan and audiofade functions to raise to power of 5 instead of 10.  It wasn't panning correctly.
'   Updated the sounds for the slings SSF by unbalancing them in the sound manager.
'   Automated the desktop, VR, and cab mode
'   Set ball reflectivity to 0.4
'   Set the bloom strength and screen space reflections scale to 0.2
'   Adjusted the playfield material to be white base, and 0.5 glossy layer.
'   Updated ball choices and ball brightness code.
'   Removed lamptimer, ramproll timer and variable
'   Adjusted ball brightness in plunger lane.
'   Updated table info.
'   Added to table1.init values for default targets, GI, switches, shadows, etc.
'   Updated ruleset in table info, and created a gamehelp image for the F12 rules
' ROM / DIP settings
'   Created a DIPs menu and default DIPs at game launch.
'   Set the ROM dip switch settings to match the instruction card (left: B-18559-2;   Right: A-19135)
'   Created a .nvram file to set the high score replay levels of 170,000, 350,000 and 430,000; and initial high score of zero.  Also 3 replays for HSTD.
'   Added sounds based on solenoids and chime sounds if wanted.
' STAR TARGETS
'   Added star trigger, prims, and animation.  Tied to GI.
'   Created RED star on and off primitives
'   Added a light under the star targets for reflections, and tied to GI. Adjusted bdl and other materials as well.  Added spearate insertoff1, 2, and 3 for animation.
'   Ensured the star insertoff material to be 0.99 transparency
'   Changed the shinyness of the star's insertoff materials to be 0.65
' DROP TARGETS
'   Rothbauerw drop targets implemented
'   Added code for a slight delay in the drop target reset. However changed to delay of zero.
'   Set dtmass at 0.1 per latest VPW example table.
'   Adjusted the elasticity of the drop targets to 0.625 per example table, and added them to "targets" collection.
'   Created all new primitives for the drop targets.  Created new UV Map and images.  Had to ensure the front of prim was facing forward for roth dt hit animations.
'   Had to increase the amount to drop VP units the drop target drops with the new prim.
' DROP TARGET SHADOWS
'   Added drop target shadows and adjusted the opacity
'   Added functionality to ensure ONLY direct hits triggers a drop target hit and for shadows
' ROLLOVERS
'   Added rollovers, and animated them.
' WALLS
'   Added wood walls, used a natural wood image
'   BUMPERS
'   Set bumpers, and created flupper style.
'   Adjusted the bumper physics, increased strength
'   Added bumper ring primitives and animated them.  Also animated the bumper skirts based off code from Beat the Clock by VPW
'   Changed the distance of bumper skirt animation to be total of 6 VPU... so bumper skirts are now z=-1.5 and than transz is -4.5... so total of -1.5 + -4.5 = -6)
'   Changed bumper image transimission scale, modulation, intensity and warmth of bumper lights, as well as the bdl of the prim.
'   Updated the bumperbiglight color
'   Updated bumper subroutines, colors, brightness, and added options for on and off bumper brightness
'   Added another bumper light for to see through cap edges
'   Softened the color on the bumper big lights (on the playfield) so not as deep color.
'   Tied bumper lights to GI and added logic to control the primitives
'   UP usedbumper image from on borgdog image, rescaled and adjusted to fit correctly
'   Adjusted each bumper cap prim left and up by 2 VPU's.
' SOUNDS
'   Updated to latest Fleep sounds per example table.  Some sounds went mono.
'   Added tones and chimes for solenoid sounds.  Created option in F12 menu.
' OUTERCAB WALLS
'   Added outer prim cab walls and adjusted to fit table.  Sized to 22 vpu wide. (about 15/32" wide... typical plywood width)
'   Turned reflectivity off on the OuterPrimBlack_DT sidewalls.
'   Adjusted OuterPrimBlack_DT to be used for both desktop and cab, just adjusted when cabmode tall sideblades
'   Added yellow outer cab walls, and made black as an option.
' DESKTOP
'   Added new desktop backglass background and animated the desktop backglass lights.
'   Added new desktop backglass and enabled the backglass reels, SA, BIP, Tilt, etc for Desktop Mode
'   Turned off desktop backglass lamp reflectivity.
'   Added desktop lockdown bar.  Added backrail prim, and outerwall prims.  Adjusted appropriately.
'   Using same lockdownbar and side rails for desktop and VR.  New siderails with metal side plates used from Alien Star
'   Put code in to control match and BIP in desktop game on and off modes.
' FLIPPERS
'   Added flippers from Buck Rogers, flipper physics, accurate flipper prims and images.
'   Updated flipper postion, flipper triggers, flipper shadows, and physics
'   Updated flipper animations, nudging parameters, flipper physics and tricks
'   Updated the LiveCatch for flippers to 12
'   Raised the flipper prims to z height of 1.  Realistically, should be slightly above playfield.
' APRON
'   Added apron, plunger cover, and sidewood images from Buck Rogers.
'   Added physical nubs at bottom of apron on both sides, as well as pegs in the corners.
' INSTRUCTION CARDS
'   Created prims for the instruction cards, new UV maps, and placed redbone's image in there.
' PLUNGER
'   Adjusted plunger groove lane
'   Added plunger and adjusted firespeed and other plunger settings
'   Set plunger.fire strength with a bit of randomness.  Also set the plunger momentum transfer to 1.1
' GUIDE RAILS
'   Primitive and physics for wire inlanes.
'   Primitive and physics for solid inlanes
'   Adjusted the wire rail guide ends and inlane wires to have different elasticity and other physics on it.  Should have some spring to it.
'   Broke the inlane rail guide physics into two sections for different elasticity.
'   Changed the metal prims of the guide rails to be active, not staic, so they can be controlled with lightlevel and gilvl changes.
'   Added top wall and image, antisnubber rails, lower rails
'   Lowered the left guide rail near the sling rubber from 10 to 7.5 vp units, as it wwas hitting the rubber.
' PEGS and PEG POSTS
'   Placed all peg posts, metal pegs, and white screw caps.  Also put longer peg inside posts where they go through plastics.
'   Added all rubber physics, and "more bouncy" physics to the top rollover lanes.
'   Added little pegs near the apron corners.  And smaller nubs near bottom
'   Created white default pegs, and have options for red, yellow, blue, and acrylic.
'   Modified the white peg material, as well as split options for bdl on each color.
' RUBBERS
'   Placed rubbers around posts and pegs, as well as the physics elements
'   Adjusted the rubbers to a height of 31.5 to align with gotlieb pegs.
'   Animated leaf and open rubbers
'   Added more bounce to the top arch rubbers.
'   Changed the rubber material equation to change with GI, and updatematerial.
' LEAF SWITCHES
'   Added rubber and leaf switch animation.
'   Added 6 10-point switches with sensors and tied logic to them.
'   On the long rubber, I tied the hit to x velocity and changed the distance of the rubber move based on speed.
' HOLE CUTOUTS
'   Added cutouts, bulb prims, triggers and trigger prims
' SCREWS
'   Added all the white, bumper, and metal screws and rotated them.
'   Ensured screws were not collidable and only toys.  No need for that.
' SLINGS
'   Added slings, sling corrections, and sling physics.  Assigned the correct sling switch id's.
'   Used gottlieb sling prim and had to change to rotobjy value vs rotx for the animation.
'   Updated slingshot correction curves based on latest VPW example table.
' TOP PLASTIC LANE GUIDES
'   Added gottlieb top lane guides. Made the material white, and adjusted the bdl on and off.
'   Added GI tops for the upper lane guides
'   Updated the primitives in blender to get a more 3D look to it.
' WALLS
'   Added metal walls, guide rails, and all associated physics
'   Adjusted walls as close as possible to emulate shots.
'   Meticulously placed the right and left orbit wall to match exact shots on videos.
'   Created a new metal wall texture to be a smaller texture used for baking with a slight pinball mark on it
'   Added physics walls in plastic areas.
' ARCH
'   Added arch prim from Borgdog's Gottlieb table.  Created white image, and used a ramp and image from HF for certificate.
' PLASTICS
'   UP created plastic image based off of borgdog's prior table images. Upscaled, and placed per vpx blueprint.
'   Created plastic walls, acrylics, and placed and rotated screws
'   Added option for LED plastics edges
'   Added GI tops to plastics.
'   Redbone used the plastics image and reupscaled and postioned to match my template.
' RUBBER BUMPER WHEEL
'   Added a rubber bumper wheel based off prims from hauntfreaks and borgdog.  I modified in blender and combined to get the look I wanted.
'   Added rubber bouncy physics to the bumber wheel.
' GATES
'   Added right plate gate, and right wire gate
'   Added gate trigger sounds and gate trigger ready logic to the plate gate.
'   Added Gottlieb style long gate prims and animated them.
'   Added light metal gate hit sounds at the top right and left gates when closed
'   Created a completely new gatewire primitive and ensured it rotated on local axis.
'   Created new gate wire bracket to match real table.
'   Added spring mechanism and animated it as the wire gate gets hit.
'   Added a gatewire trigger to pulse switch 52
'   F12 OPTIONS
'   Implemented full F12 menu system and tied everything to light room brightness
'   Adjusted room and cab brightness levels to be controlled via F12 menu for day/night
'   Added options for colors on flippers, screw caps, lane guides, bumper disks, and drop targets.
'   Added new ball images and options and ball brightness code.
'   Added an option for a flashers on slings and bumpers.  Each selectable separately in the F12 options.
'   Added 100ms to the sling flasher option
' SAUCERS
'   Created playfield_mesh for saucer holes and other holes to below playfield
'   Added gottlieb saucer with pivot arm.
'   Added a black wall under playfield to ensure cutouts have a black bottom, and cutout hole for saucer.
'   Had to modify the playfield to get a nicer wood grain around saucers.  Existing pf wasn't consistent.
'   Tuned the saucers to match online videos as close as I could
'   Turned reflectivity off in the saucer prim elements
'   Put a ball reflection mask off in the saucer.
'   Removed visibility of psaucerfloors: only using for physics on gottlieb sys1.
'   Added saucer lights to help with baking, as well as tie the bdl of the saucers to GI and the animated light subs.
' PHYSICS
'   Adjusted all physics settings tuned to videos: plunger, slings, rubber sling areas, saucer, lane guide rails, bumpers, EVERY shot based off videos.
'   Adjusted the slingshot hit and threshold values.
'   Watched hours of videos to ensure accuracy of physics, ball rolloffs, flippers, etc.
'   Changed the physics by the flippers to a wall, as once inawhile a ball will stick there (like what I saw in the capture ball awhile ago)
'   Using peg placement based off playfield scan I had.
' VR
'   Added VR Room environment, and VR cabinet and backglass.
'   Added the VR Room options to the F12 menu
'   Adjusted the room brightness and the cab levels based off the day/night mode... controlled in the F12 menu.
'   Animated the VR plunger, flipper buttons, start button.
'   Added VR undercab lighting and made it an option in F12 menu
'   Changed material on VR cab to remove shininess and light reflections
'   Added VR cab and backbox images, VR posters.  All darthvito environments, and mine, and sixtoes.
'   Added VR backglass and text images (from Hauntfreaks).  Adjusted location and lighting of VR LEDs.
'   Animated the backglass text. (In GIMP I used a supernova lighting filter 50 radius, color match of bg area, overlay for text on VR BG)
'   Fixed the undercab LED lighting GI script in VR when day/night mode adjusts (with option on or offset) as well as the new floor.  (was first done in Star Gazer)
'     Enabled the backglass SA, Tilt, GameOver, Match, BIP, HSTD for VR Mode.
'   Added a small coin block wall to hide seeing inside of cab through coin return opening
'   Added a plunger cover wall to hide VR plunger in gap near lockbar
' GI
'   Learned that the GI is ALSO controlled by the tilt solenoid under the playfield.
'    A switch contact in the relay is normally closed to leave the playfield GI on. However, when the game is in tilt mode, the T relay is engaged,
'    and playfield GI and controlled lamps will be turned off. There is also a switch contact to disable the playfield solenoids.
'   GI comes on when tilt solenoid is active.(not after 2 seconds like other tables)
'   Added logic to tilt controller to turn pfgi on and off.
'   Added logic to the pfgi subroutine to ensure sound of relays is accurate on gilvl change only.
'   Had to add a updatelamps routine for cab mode to control GI and tilt.
'   Moved backglass GI to come on with other GI for desktop and VR modes.
'   Updated the GI to be controlled by not only GI, but also F12 lightlevel changes.
'   Set the bdl on the GI bulb prims to 4.
'   Added screw caps, and all peg posts to GI.
'   Added main GI (hidden for baking)
'   Adjusted the GItops for the lane guides intensity up a little, as well as the falloff power.
'   Added updated lane guide prims to have a more 3D effect on the pointy part.
'   Changed the GItop on the lane guides to transmission of .2 and intensity of 25.  Also changed the material to have .985 transparency.
'   Updated the VRRoomLL collection to include the start button.
' INSERTS
'   Added 3D inserts with latest alllamps animation.
'   Updated the glows around the insert lights.
'   Added an insert wall block under the inserts for the bulb showing through other insert prims and playfield.
'   Added an insert factor variable to increase the brightness of the inserts.
'   Warmed the white inserts up some.
'   Adjusted colors of all inserts to match playfield, pictures and videos online.  Had to create a wine/red/pink type color for extra ball.
'   Created a different solid off image.
'   Changed all light fading to 120 fade up, and 160 fade down.
' GI BAKING
'   Did a VLM bake of the GI
'   Moved the lightmaps to 0.1 above the playfield.  Moved the light insert glows to 0.2 (above the GI)
'   Decreased the opacity of the GI playfield lighting to 75%.  It was washing out some of the white on playfield before.
'   Turned on raytracing on bumperbiglights and bumpersmalllights.
'   IMAGE and Option UPDATES
'   Redbone updated several images for the instruction cards, apron, and plunger cover.
'   UP updated the apron image to remove the shadow effect that Redbone had.  It didn't line up with the Apron prim quite right.
' DTShadows
'   Created new drop target shadow images
'   Readjusted drop target shadows after GI bake.
'   Set the dtshadows to come on with GI.
'   Baked Walls
'   Baked the metal walls and have GI and insert reflections on them.
'   Baked the inner wood walls and have GI reflections on them.
'   Updated ALL the GI script to control baked walls.
'   Added a VLM.bakesolid material to the BM_walls primitive to avoid any incorrect hue.
'   AFTERBAKE ADJUSTMENTS
'   Updated the GI and setroombrightness routines
'   Removed unused images and materials.
'   Converted images to webp
'   Added a glass top physics to ensure that if the ball gets hit by a resetting target, it'll stay on the table.
'   Added DOF to the bumpers
'   Added DOF to the left sling
'2.0.0  Released
'   Added wall15 to metal collection (little wall in plunger lane that ball rests on.


' Thanks to CainArg for providing the playfield scan, cleanup, and hole cuts.
' Thanks to Redbone for updating the playfield scale, as well as updating the plastics, apron, and instruction cards.

'Notes on physics:
' https://www.youtube.com/watch?v=G7dpJJ-JPWU  at 9:28
'   Tommy says those curved metal backs behind saucers made it VERY difficult to get ball into saucer.  On purpose to be hard.


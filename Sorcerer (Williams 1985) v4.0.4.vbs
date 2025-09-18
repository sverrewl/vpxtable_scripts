'
'
'            __
'  -_-/    ,-||-,   -__ /\     ,- _~.   ,- _~, -__ /\     ,- _~, -__ /\
' (_ /    ('|||  )    || \,   (' /|    (' /| /   || \,   (' /| /   || \,
'(_ --_  (( |||--))  /|| /   ((  ||   ((  ||/=  /|| /   ((  ||/=  /|| /
'  --_ ) (( |||--))  \||/-   ((  ||   ((  ||    \||/-   ((  ||    \||/-
' _/  ))  ( / |  )    ||  \   ( / |    ( / |     ||  \   ( / |     ||  \
'(_-_-     -____-   _---_-|,   -____-   -____- _---_-|,   -____- _---_-|,
'
'
'       Williams's Sorcerer / IPD No. 2242 / March, 1985 / 4 Players
'     https://www.ipdb.org/machine.cgi?id=2242
'     VPX version 4.0 by UnclePaulie
'
'
'********************************************************************************************************************************
'******** Work done by UnclePaulie on Hybrid version 1.0 - 4.0 *********
' original by JPSalas (2018)
' V1.0 - V4.0 by UnclePaulie
'   v1.0 Added VR room
'   v2.0 Updated to include hybrid, fleep sounds, nfozzy physics, roth targets (2021)
'   v3.0 GI updates, physics, slings, gBOT trough logic, 3D inserts, Lampz (2022)
'   v4.0 Overhauled table completely to updated VPW physics, desktop backglass, VR room options, alllamps, options, sounds, more realistic play
'       Added latest VPW standards, physics, GI VLM, 3D inserts, playfield mesh, allLamps routines, shadows, updated backglass for desktop,
'     roth drop targets, standalone compatibility, sling corrections, updated flipper physics, VR and fully hybrid, and staged flippers.
'     Playfield, plastics, and other images provided by Redbone.
'   Full details at the bottom of the script.
'********************************************************************************************************************************

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

' ALL OPTIONS ARE IN THE F12 MENU, EXCEPT FOR THE BALL COLOR


const ChangeBall  = 1   '0 = Ball changes are done via the drop down table menu
              '1 = ball selection per the below BallLightness option (default)
const BallLightness = 8   '0 = dark, 1 = not as dark, 2 = bright, 3 = brightest, 4 = MrBallDark, 5 = DarthVito Ball1
              '6 = DarthVito HDR1, 7 = DarthVito HDR Bright, 8 = DarthVito HDR2 (Default), 9 = DarthVito ballHDR3,  10 = DarthVito ballHDR4
              '11 = DarthVito ballHDRdark, 12 = Borg Ball, 13 = SteelBall2

' Insert brightness scale
const insertfactor    = 1   ' adjusts the level of the brightness of the inserts when on

' Force VR mode on desktop
Const ForceVR = 0     '0 is off (default): 1 is to force VR on.

'***********************************************************************
'* Operator Menu Settings & Instructions via the ROM and displays ******
'*  full description of adjustments at bottom of script. ***************
'***********************************************************************

'Press Advance (8) until 04 is shown in the Credit Display (or move Auto/Manual switch to Auto first)
'Auto-Up / Manual-Down switch is controled by 7(Press ENDto see a graphical respresentation of the toggle switch position)
'9 will exit/reset


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
Const tnob = 2 ' total number of balls
Const lob = 0   'number of locked balls


Dim tablewidth: tablewidth = Table1.width
Dim tableheight: tableheight = Table1.height
Dim BIPL : BIPL = 0       'Ball in plunger lane
Dim DT34up,DT35up,DT36up
Dim SBall1, SBall2, gBOT

'Const cGameName = "sorcr_l1"
Const cGameName = "sorcr_l2"

Const UseSolenoids = 2

Const UseLamps = 1
Const UseGI = 0
Const UseSync = 0
Const HandleMech = 0

LoadVPM "01550000", "S11.VBS", 3.26

'******************************************************
' VLM Arrays
'******************************************************

' VLM  Arrays - Start
' Arrays per baked part
'Dim BP_Playfield: BP_Playfield=Array(BM_Playfield, LM_All_Lights_F6_Playfield, LM_All_Lights_F7_Playfield, LM_All_Lights_gi_001_Playfield, LM_All_Lights_gi_002_Playfield, LM_All_Lights_gi_003_Playfield, LM_All_Lights_gi_004_Playfield, LM_All_Lights_gi_005_Playfield, LM_All_Lights_gi_006_Playfield, LM_All_Lights_gi_007_Playfield, LM_All_Lights_gi_008_Playfield, LM_All_Lights_gi_009_Playfield, LM_All_Lights_gi_010_Playfield, LM_All_Lights_gi_011_Playfield, LM_All_Lights_gi_012_Playfield, LM_All_Lights_gi_013_Playfield, LM_All_Lights_gi_014_Playfield, LM_All_Lights_gi_015_Playfield, LM_All_Lights_gi_016_Playfield, LM_All_Lights_gi_017_Playfield, LM_All_Lights_gi_018_Playfield, LM_All_Lights_gi_019_Playfield, LM_All_Lights_gi_lane_001_Playf, LM_All_Lights_gi_lane_002_Playf, LM_All_Lights_gi_lane_003_Playf, LM_All_Lights_gi_lane_004_Playf, LM_All_Lights_gi_lane_005_Playf)
' Arrays per lighting scenario
Dim BL_All_Lights_F6: BL_All_Lights_F6=Array(LM_All_Lights_F6_Playfield)
Dim BL_All_Lights_F7: BL_All_Lights_F7=Array(LM_All_Lights_F7_Playfield)
Dim BL_All_Lights_gi_001: BL_All_Lights_gi_001=Array(LM_All_Lights_gi_001_Playfield)
Dim BL_All_Lights_gi_002: BL_All_Lights_gi_002=Array(LM_All_Lights_gi_002_Playfield)
Dim BL_All_Lights_gi_003: BL_All_Lights_gi_003=Array(LM_All_Lights_gi_003_Playfield)
Dim BL_All_Lights_gi_004: BL_All_Lights_gi_004=Array(LM_All_Lights_gi_004_Playfield)
Dim BL_All_Lights_gi_005: BL_All_Lights_gi_005=Array(LM_All_Lights_gi_005_Playfield)
Dim BL_All_Lights_gi_006: BL_All_Lights_gi_006=Array(LM_All_Lights_gi_006_Playfield)
Dim BL_All_Lights_gi_007: BL_All_Lights_gi_007=Array(LM_All_Lights_gi_007_Playfield)
Dim BL_All_Lights_gi_008: BL_All_Lights_gi_008=Array(LM_All_Lights_gi_008_Playfield)
Dim BL_All_Lights_gi_009: BL_All_Lights_gi_009=Array(LM_All_Lights_gi_009_Playfield)
Dim BL_All_Lights_gi_010: BL_All_Lights_gi_010=Array(LM_All_Lights_gi_010_Playfield)
Dim BL_All_Lights_gi_011: BL_All_Lights_gi_011=Array(LM_All_Lights_gi_011_Playfield)
Dim BL_All_Lights_gi_012: BL_All_Lights_gi_012=Array(LM_All_Lights_gi_012_Playfield)
Dim BL_All_Lights_gi_013: BL_All_Lights_gi_013=Array(LM_All_Lights_gi_013_Playfield)
Dim BL_All_Lights_gi_014: BL_All_Lights_gi_014=Array(LM_All_Lights_gi_014_Playfield)
Dim BL_All_Lights_gi_015: BL_All_Lights_gi_015=Array(LM_All_Lights_gi_015_Playfield)
Dim BL_All_Lights_gi_016: BL_All_Lights_gi_016=Array(LM_All_Lights_gi_016_Playfield)
Dim BL_All_Lights_gi_017: BL_All_Lights_gi_017=Array(LM_All_Lights_gi_017_Playfield)
Dim BL_All_Lights_gi_018: BL_All_Lights_gi_018=Array(LM_All_Lights_gi_018_Playfield)
Dim BL_All_Lights_gi_019: BL_All_Lights_gi_019=Array(LM_All_Lights_gi_019_Playfield)
Dim BL_All_Lights_gi_lane_001: BL_All_Lights_gi_lane_001=Array(LM_All_Lights_gi_lane_001_Playf)
Dim BL_All_Lights_gi_lane_002: BL_All_Lights_gi_lane_002=Array(LM_All_Lights_gi_lane_002_Playf)
Dim BL_All_Lights_gi_lane_003: BL_All_Lights_gi_lane_003=Array(LM_All_Lights_gi_lane_003_Playf)
Dim BL_All_Lights_gi_lane_004: BL_All_Lights_gi_lane_004=Array(LM_All_Lights_gi_lane_004_Playf)
Dim BL_All_Lights_gi_lane_005: BL_All_Lights_gi_lane_005=Array(LM_All_Lights_gi_lane_005_Playf)
'Dim BL_World: BL_World=Array(BM_Playfield)
' Global arrays
'Dim BG_Bakemap: BG_Bakemap=Array(BM_Playfield)
Dim BG_Lightmap: BG_Lightmap=Array(LM_All_Lights_F6_Playfield, LM_All_Lights_F7_Playfield, LM_All_Lights_gi_001_Playfield, LM_All_Lights_gi_002_Playfield, LM_All_Lights_gi_003_Playfield, LM_All_Lights_gi_004_Playfield, LM_All_Lights_gi_005_Playfield, LM_All_Lights_gi_006_Playfield, LM_All_Lights_gi_007_Playfield, LM_All_Lights_gi_008_Playfield, LM_All_Lights_gi_009_Playfield, LM_All_Lights_gi_010_Playfield, LM_All_Lights_gi_011_Playfield, LM_All_Lights_gi_012_Playfield, LM_All_Lights_gi_013_Playfield, LM_All_Lights_gi_014_Playfield, LM_All_Lights_gi_015_Playfield, LM_All_Lights_gi_016_Playfield, LM_All_Lights_gi_017_Playfield, LM_All_Lights_gi_018_Playfield, LM_All_Lights_gi_019_Playfield, LM_All_Lights_gi_lane_001_Playf, LM_All_Lights_gi_lane_002_Playf, LM_All_Lights_gi_lane_003_Playf, LM_All_Lights_gi_lane_004_Playf, LM_All_Lights_gi_lane_005_Playf)
Dim BG_All: BG_All=Array(LM_All_Lights_F6_Playfield, LM_All_Lights_F7_Playfield, LM_All_Lights_gi_001_Playfield, LM_All_Lights_gi_002_Playfield, LM_All_Lights_gi_003_Playfield, LM_All_Lights_gi_004_Playfield, LM_All_Lights_gi_005_Playfield, LM_All_Lights_gi_006_Playfield, LM_All_Lights_gi_007_Playfield, LM_All_Lights_gi_008_Playfield, LM_All_Lights_gi_009_Playfield, LM_All_Lights_gi_010_Playfield, LM_All_Lights_gi_011_Playfield, LM_All_Lights_gi_012_Playfield, LM_All_Lights_gi_013_Playfield, LM_All_Lights_gi_014_Playfield, LM_All_Lights_gi_015_Playfield, LM_All_Lights_gi_016_Playfield, LM_All_Lights_gi_017_Playfield, LM_All_Lights_gi_018_Playfield, LM_All_Lights_gi_019_Playfield, LM_All_Lights_gi_lane_001_Playf, LM_All_Lights_gi_lane_002_Playf, LM_All_Lights_gi_lane_003_Playf, LM_All_Lights_gi_lane_004_Playf, LM_All_Lights_gi_lane_005_Playf)

'Dim BG_All: BG_All=Array(BM_Playfield, LM_All_Lights_F6_Playfield, LM_All_Lights_F7_Playfield, LM_All_Lights_gi_001_Playfield, LM_All_Lights_gi_002_Playfield, LM_All_Lights_gi_003_Playfield, LM_All_Lights_gi_004_Playfield, LM_All_Lights_gi_005_Playfield, LM_All_Lights_gi_006_Playfield, LM_All_Lights_gi_007_Playfield, LM_All_Lights_gi_008_Playfield, LM_All_Lights_gi_009_Playfield, LM_All_Lights_gi_010_Playfield, LM_All_Lights_gi_011_Playfield, LM_All_Lights_gi_012_Playfield, LM_All_Lights_gi_013_Playfield, LM_All_Lights_gi_014_Playfield, LM_All_Lights_gi_015_Playfield, LM_All_Lights_gi_016_Playfield, LM_All_Lights_gi_017_Playfield, LM_All_Lights_gi_018_Playfield, LM_All_Lights_gi_019_Playfield, LM_All_Lights_gi_lane_001_Playf, LM_All_Lights_gi_lane_002_Playf, LM_All_Lights_gi_lane_003_Playf, LM_All_Lights_gi_lane_004_Playf, LM_All_Lights_gi_lane_005_Playf)
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
  SpinnerTimer
  UpdateBallBrightness
  RollingUpdate       'Update rolling sounds
  DoSTAnim          'Standup target animations
  DoDTAnim          'Drop target animations
  BSUpdate          'Update ambient ball shadows

   If VR_Room = 0 and cab_mode = 0 Then
        DisplayTimer
    End If

  If VR_Room=1 Then
    VRDisplayTimer
    VRBGtext
  End If

End Sub


'The CorTimer interval should be 10. It's sole purpose is to update the Cor (physics) calculations

CorTimer.Interval = 10
Sub CorTimer_Timer(): Cor.Update: End Sub


'******************************************************
'  ZINI: Table Initialization and Exiting
'******************************************************

Sub Table1_Init
    vpmInit me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Sorcerer - Williams 1985" & vbNewLine & "by UnclePaulie"
    .Games(cGameName).Settings.Value("sound") = 1 ' Set sound (0=OFF, 1=ON)
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


  'Map all lamps to the corresponding ROM output using the value of TimerInterval of each light object
  vpmMapLights AllLamps     'Make a collection called "AllLamps" and put all the light objects in it.

  On Error Goto 0

    ' Nudging
    vpmNudge.TiltSwitch = 1
    vpmNudge.Sensitivity = 3
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingshot)

  'Ball initializations need for physical trough
  Set SBall1 = BallRelease.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set SBall2 = Slot1.CreateSizedballWithMass(Ballsize/2,Ballmass)
  gBOT = Array(SBall1, SBall2)

  'Forces the trough switches to "on" at table boot so the game logic knows there are balls in the trough.
  Controller.Switch(28) = 0
  Controller.Switch(29) = 1
  Controller.Switch(30) = 1

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1


' GI initially off
  SolGi 1

' GI comes on after 2 seconds
  vpmTimer.AddTimer 2000, "SolGi 0'"

  if VR_Room = 1 Then
    setup_backglass()
  End If

  SetupOptions

' Make drop target shadows visible

Dim xx
  for each xx in ShadowDT
    xx.visible=True
  Next

' Drop Target Variable state for DT Shadows
  DT34up=1
  DT35up=1
  DT36up=1


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


Dim LightLevel : LightLevel = 0.15        ' Level of room lighting (0 to 1), where 0 is dark and 100 is brightest
Dim ColorLUT : ColorLUT = 1           ' Color desaturation LUTs: 1 to 11, where 1 is normal and 11 is black'n'white
Dim VolumeDial : VolumeDial = 0.8             ' Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Dim BallRollVolume : BallRollVolume = 0.5     ' Level of ball rolling volume. Value between 0 and 1
Dim RampRollVolume : RampRollVolume = 0.5     ' Level of ramp rolling volume. Value between 0 and 1
Dim ROMVolume                 ' Level of ROM volume. Value between 0 and 1
Dim WallClock, poster, poster2, CustomWalls
Dim FlipperColor, ScrewCapColor, SpinnerColor, CardOption, cabsideblades

Dim dspTriggered : dspTriggered = False
Sub Table1_OptionEvent(ByVal eventId)
    If eventId = 1 And Not dspTriggered Then dspTriggered = True : DisableStaticPreRendering = True : End If

  ROMVolume = Table1.Option("ROM Volume (-32 off; 0 max) - Requires Restart; altsound set to 0", -32, 0, 1, 0, 0)

  ' Table Options
  FlipperColor = Table1.Option("Flipper Color", 0, 1, 1, 0, 0, Array("Red", "Black"))
  SpinnerColor = Table1.Option("Spinner Style", 0, 1, 1, 0, 0, Array("Normal", "Glitter"))
  ScrewCapColor = Table1.Option("Screw Cap Color", 0, 5, 1, 0, 0, Array("White", "Black", "Bronze", "Silver", "Blue/Green", "Gold"))
  CardOption = Table1.Option("Instruction Cards Color", 0, 1, 1, 0, 0, Array("Black", "White"))
  cabsideblades = Table1.Option("Cabinet Sideblades - full cabinet only", 0, 1, 1, 0, 0, Array("Extra Tall - default in cab mode", "Lowered Blades"))

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
  SetRoomBrightness LightLevel   'Uncomment this line for lightmapped tables.
    If eventId = 3 And dspTriggered Then dspTriggered = False : DisableStaticPreRendering = False : End If

  ' VR Room Environment
  WallClock = Table1.Option("VR Wall Clock On/Off", 0, 1, 1, 1, 0, Array("Off", "On"))
    CustomWalls = Table1.Option("VR Room Environment", 0, 4, 1, 2, 0, _
    Array("UP Original", "Sixtoe Arcade", "DarthVito Home", "DarthVito Plaster", "DarthVito Blue"))
  Poster = Table1.Option("VR Poster 1 On/Off", 0, 1, 1, 1, 0, Array("Off", "On"))
  Poster2 = Table1.Option("VR Poster 2 On/Off", 0, 1, 1, 1, 0, Array("Off", "On"))

  SetupRoom

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
  lflip.rotz = a
End Sub

Sub LeftFlipper1_Animate
  dim a: a = LeftFlipper1.CurrentAngle
  FlipperLSh1.RotZ = a
  lflip2.rotz = a
End Sub

Sub RightFlipper_Animate
  dim a: a = RightFlipper.CurrentAngle
  FlipperRSh.RotZ = a
  rflip.rotz = a
End Sub

Sub Gate3_Animate
  dim a: a = (Gate3.CurrentAngle * -0.5) + 6.5
  GateFlap3.Rotx = a
End Sub

Sub Gate4_Animate
  dim a: a = (Gate4.CurrentAngle * -0.5) + 6.5
  GateFlap4.Rotx = a
End Sub

Sub Gate5_Animate
  dim a: a = max(Gate5.CurrentAngle,0)
  Gate5p.Rotx = a
End Sub


'******************************************************
'   ZBBR: BALL BRIGHTNESS
'******************************************************

Const BallBrightness =  1       'Ball brightness - Value between 0 and 1 (0=Dark ... 1=Bright)

' Constants for plunger lane ball darkening.
' You can make a temporary wall in the plunger lane area and use the co-ordinates from the corner control points.
Const PLOffset = 0.65     'Minimum ball brightness scale in plunger lane
Const PLLeft = 853.5      'X position of punger lane left
Const PLRight = 952       'X position of punger lane right
Const PLTop = 950       'Y position of punger lane top
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

Sub SetRoomBrightness (level)
Dim llelement, girubbercolor, x, gibumpercolor1, gibumpercolor2, gibumpercolor3, gibumpercolor4
  For each llelement in VRRoomLL
    llelement.blenddisablelighting = level*6.7
  Next
  For each llelement in VRCabLL
    llelement.blenddisablelighting = level*2.3
  Next
  VRposter.blenddisablelighting = level * 3.35
  VRposter2.blenddisablelighting = level * 3.35
  VR_Clockface.opacity = level*670
  For each llelement in TopHardware
    llelement.blenddisablelighting = level * 0.3
  Next

' GI and GI Bulb Prims
  For each x in GI: x.State = gilvl: Next
  For Each x in GIBulbPrims: x.blenddisablelighting = (2*(gilvl)): Next

' rubbers
  girubbercolor = (128 * (gilvl * level/0.15)) + 128
  MaterialColor "Rubber White",RGB(girubbercolor,girubbercolor,girubbercolor)

' Bumpers
  BumperTop_001.BlendDisableLighting = ((0.1 * gilvl) + 0.15)
  BumperTop_002.BlendDisableLighting = ((0.1 * gilvl) + 0.15)
  BumperTop_003.BlendDisableLighting = ((0.1 * gilvl) + 0.15)
  bumperhighlight1.opacity = (1500*(gilvl))
  bumperhighlight2.opacity = (1500*(gilvl))
  bumperhighlight3.opacity = (1500*(gilvl))

  bumperbase1.BlendDisableLighting = 0.015 * gilvl
  bumperbase2.BlendDisableLighting = 0.015 * gilvl
  bumperbase3.BlendDisableLighting = 0.015 * gilvl
  bumperdisk1.BlendDisableLighting = 0.3 + (0.1 * gilvl)
  bumperdisk2.BlendDisableLighting = 0.3 + (0.1 * gilvl)
  bumperdisk3.BlendDisableLighting = 0.3 + (0.1 * gilvl)

' flippers
  For Each x in GIFlip: x.blenddisablelighting = ((0.25 * gilvl) + 0.35) * (level/0.15): Next

' Sleeves
  For each x in GISleeves: x.blenddisablelighting = ((0.015 * gilvl) + 0.01) * (level/0.15): Next

' targets
  For each x in GITargets: x.blenddisablelighting = ((0.25 * gilvl)  + 0.15) * (level/0.15): Next
  For each x in GIdtTargets: x.blenddisablelighting = ((0.15 * gilvl)  + 0.15) * (level/0.15): Next

' prims (Red Pegs)
  For each x in GIPegs: x.blenddisablelighting = ((0.15 * gilvl)  + 0.005) * (level/0.15): Next

' prims (Cutouts)
  For each x in GICutoutPrims: x.blenddisablelighting = ((0.25 * gilvl)  + 0.15) * (level/0.15): Next
  For each x in GICutoutWalls: x.blenddisablelighting = ((0.1 * gilvl)  + 0.2) * (level/0.15): Next

' prims (apron, brackets, spinner)

    ApronWilliams.blenddisablelighting = ((0.15 * gilvl) + 0.35) * (level/0.15)
    CardLeft.blenddisablelighting = ((0.15 * gilvl) + 0.35) * (level/0.15)
    CardRight.blenddisablelighting = ((0.15 * gilvl) + 0.35) * (level/0.15)
    apronplunger.blenddisablelighting = ((0.105 * gilvl) + 0.35) * (level/0.15)
    SpinnerPrim.blenddisablelighting = ((0.05 * gilvl) + 0.05) * (level/0.15)
    SpinnerPrim2.blenddisablelighting = ((0.05 * gilvl) + 0.05) * (level/0.15)

' prims (White Screws Main)
  For each x in GIScrewCaps: x.blenddisablelighting = ((0.05 * gilvl) + 0.005) * (level/0.15): Next

' metals and laneguard
  For each x in GIMetal: x.blenddisablelighting = ((0.0075 * gilvl) + 0.0075) * (level/0.15): Next
  For each x in GIMetalWalls: x.blenddisablelighting = ((0.05 * gilvl) + 0.0075) * (level/0.15): Next

' leaf switch sensors
  For each x in GILeafSwitch: x.blenddisablelighting = ((0.1 * gilvl) + 0.1) * (level/0.15): Next
  Sling1.blenddisablelighting = (0.005 * (gilvl * level/0.15))

' ramps
if gilvl = 0 Then
  lramp1.image="ramp1off"
  lramp2.image="ramp2off"
  lramp4.image="rampwiresoff"
  rramp.image="ramplidoffpng"
  rramp2.image="lockplateoff"
  outerb.blenddisablelighting = .05 * (level/0.15)
  lramp1.blenddisablelighting = 0.6 * (level/0.15)
  lramp2.blenddisablelighting = 0.6 * (level/0.15)
  lramp3.blenddisablelighting = 0.6 * (level/0.15)
  lramp4.blenddisablelighting = 0.6 * (level/0.15)
  rramp.blenddisablelighting = 0.2 * (level/0.15)
  rramp1.blenddisablelighting = 0.2 * (level/0.15)
  rramp2.blenddisablelighting = 0.5 * (level/0.15)
Else
  lramp1.image="ramp1on"
  lramp2.image="ramp2on"
  lramp4.image="rampwireson"
  rramp.image="ramplidonpng"
  rramp2.image="lockplateon"
  outerb.blenddisablelighting = 0.6 * (level/0.15)
  lramp1.blenddisablelighting = 0.6 * (level/0.15)
  lramp2.blenddisablelighting = 0.6 * (level/0.15)
  lramp3.blenddisablelighting = 0.6 * (level/0.15)
  lramp4.blenddisablelighting = 0.6 * (level/0.15)
  rramp.blenddisablelighting = 0.2 * (level/0.15)
  rramp1.blenddisablelighting = 0.2 * (level/0.15)
  rramp2.blenddisablelighting = 0.5 * (level/0.15)
End If

End Sub


'******************************************************
' ZKEY: Key Press Handling
'******************************************************

'*********************
' VR Plunger code
'*********************

Sub TimerVRPlunger_Timer
  If PinCab_Shooter.Y < -260 then
       PinCab_Shooter.Y = PinCab_Shooter.Y + 5
  End If
End Sub

Sub TimerVRPlunger1_Timer
  PinCab_Shooter.Y = -351 + (5* Plunger.Position) -20
End Sub


'*********************
' Keys
'*********************

Sub Table1_KeyDown(ByVal Keycode)
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
    PinCab_Shooter.Y = -351
  End If

  If keycode = LeftFlipperKey Then
    VRFlipperButtonLeft.X = 2112.278 + 8
  End If

  If keycode = RightFlipperKey Then
    Controller.Switch(44) = 1
    VRFlipperButtonRight.X = 2099.209 - 8
  End If

  If keycode = StartGameKey Then
    StartButton.y = 811.9485 - 5
    SoundStartButton
  End If

  If KeyDownHandler(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal Keycode)

  If keycode = PlungerKey Then
    Plunger.Fire
    If BIPL = 1 Then
      SoundPlungerReleaseBall()     'Plunger release sound when there is a ball in shooter lane
    Else
      SoundPlungerReleaseNoBall()     'Plunger release sound when there is no ball in shooter lane
    End If
    TimerVRPlunger.Enabled = False
        TimerVRPlunger1.Enabled = True
    PinCab_Shooter.Y = -351
  End If

  If keycode = LeftFlipperKey Then
    VRFlipperButtonLeft.X = 2112.278
  End If
  If keycode = RightFlipperKey Then
    Controller.Switch(44) = 0
    VRFlipperButtonRight.X = 2099.209
  End If

  If keycode = StartGameKey Then
    StartButton.y = 811.9485
  End If

  If KeyUpHandler(keycode) Then Exit Sub
End Sub


'******************************************************
' ZSOL: Solenoids & Flashers
'******************************************************

SolCallback(1) = "SolOuthole"     ' Outhole
SolCallback(2) = "SolBallRelease"   ' Ball Release
SolCallback(3) = "SolMBEject"     ' Multi-Ball Eject
SolCallback(4) = "SolDtBank"    ' Three Bank Reset
SolCallback(6) = "Flash6"         ' Drop Target Flasher
SolCallback(7) = "Flash7"         ' Center Target Bank Flashers
SolCallback(8) = "sol8subroutine" ' Back Panel Eye Flashers, and Backbox Flasher
SolCallback(11) = "SolGi"           ' General Illumination
SolCallback(14) = "SolKnocker"    ' Knocker
SolCallback(15) = "SolBell"     ' Sorcerer Bell
'SolCallback(16)
'SolCallback(17)     = ""       ' Left Sling
'SolCallback(18)     = ""       ' Right Sling
'SolCallback(19)    = ""      ' Left Jet Bumper
'SolCallback(20)    = ""        ' Bottom Jet Bumper
'SolCallback(21)    = ""        ' Right Jet Bumper
'SolCallback(23) = "vpmNudge.SolGameOn"

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"
SolCallback(sULFlipper) = "SolULFlipper"

SolCallback(5) = "Sol5"     ' Demon Backdrop Flasher
SolCallback(19)  = "Sol19"  ' Backbox Flasher

Sub Sol5(Enabled)
  If Enabled Then
    f5flash.state = 1
    L_DT_1.intensity = 10
  Else
    f5flash.state = 0
    if giLvl = 1 then
      L_DT_1.intensity = 5
    Else
      L_DT_1.intensity = 0
    End If
  End If
End Sub

sub sol8subroutine(Enabled)
  if enabled then
    l_f8.state = 1
  Else
    l_f8.state = 0
  end If
end Sub


Sub Sol19(Enabled)
  If Enabled Then
    f19flash.state = 1
  Else
    f19flash.state = 0
  End If
End Sub


sub Flash6(Enabled)
  If Enabled Then
    f6.state = 1
    Sound_Flash_Relay 1, bumper2
  Else
    f6.state = 0
    Sound_Flash_Relay 0, bumper2
  End If
End Sub

sub Flash7(Enabled)
  If Enabled Then
    f7.state = 1
    Sound_Flash_Relay 1, bumper3
  Else
    f7.state = 0
    Sound_Flash_Relay 0, bumper3
  End If
End Sub


'******************************************************
' ZDRN: Drain, Trough, and Ball Release
'******************************************************

Sub BallRelease_Hit   : Controller.Switch(30) = 1 : UpdateTrough : End Sub
Sub BallRelease_UnHit   : Controller.Switch(30) = 0 : UpdateTrough : End Sub
Sub Slot1_Hit       : Controller.Switch(29) = 1 : UpdateTrough : End Sub
Sub Slot1_UnHit     : Controller.Switch(29) = 0 : UpdateTrough : End Sub

Sub UpdateTrough
  UpdateTroughTimer.Interval = 300
  UpdateTroughTimer.Enabled = 1
End Sub

Sub UpdateTroughTimer_Timer
  If BallRelease.BallCntOver = 0 Then Slot1.kick 60, 10
  Me.Enabled = 0
End Sub


'*****************  DRAIN & RELEASE  ******************

Sub SolOuthole(Enabled)
  If Enabled Then
    Drain.kick 60, 16
    UpdateTrough
  End If
End Sub

Sub SolBallRelease(enabled)
  If enabled Then
    BallRelease.kick 60, 12
    RandomSoundBallRelease BallRelease
    UpdateTrough
  End If
End Sub

Sub Drain_Hit()
  Controller.Switch(28) = 1
  RandomSoundDrain Drain
End Sub

Sub Drain_UnHit()
  Controller.Switch(28) = 0
End Sub


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

Sub SolULFlipper(Enabled)
  If Enabled Then
    FlipperActivate LeftFlipper1, ULFPress
    LeftFlipper1.RotateToEnd
    If leftflipper1.currentangle < leftflipper1.endangle + ReflipAngle Then
      RandomSoundReflipUpLeft LeftFlipper1
    Else
      SoundFlipperUpAttackLeft LeftFlipper1
      RandomSoundFlipperUpLeft LeftFlipper1
    End If
  Else
    FlipperDeactivate LeftFlipper1, ULFPress
    LeftFlipper1.RotateToStart
    If LeftFlipper1.currentangle < LeftFlipper1.startAngle - 5 Then
      RandomSoundFlipperDownLeft LeftFlipper1
    End If
    FlipperLeft1HitParm = FlipperUpSoundLevel
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
  LeftFlipperCollide parm
End Sub


'************************************************************
' ZSLG: Slingshot Animations
'************************************************************

Dim RStep, LStep, UStep

Sub RightSlingShot_Slingshot
  RS.VelocityCorrect(ActiveBall)
  vpmTimer.PulseSw 33
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
  vpmTimer.PulseSw 32
  RandomSoundSlingshotRight SLING2
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
' Leaf Sensors
'*******************************************

Sub Phys_RB_Sensor_sw39_Hit: vpmTimer.PulseSw 39: End Sub
Sub Phys_RB_Sensor_sw40_Hit: vpmTimer.PulseSw 40: End Sub
Sub Phys_RB_Sensor_sw42_Hit: vpmTimer.PulseSw 42: End Sub
Sub Phys_RB_Sensor_sw43_Hit: vpmTimer.PulseSw 43: End Sub


'*******************************************
' Bumpers
'*******************************************

Sub Bumper1_Hit
  vpmTimer.PulseSw 21
  RandomSoundBumperTop Bumper1
End Sub

Sub Bumper2_Hit
  vpmTimer.PulseSw 23
  RandomSoundBumperMiddle Bumper2
End Sub

Sub Bumper3_Hit
  vpmTimer.PulseSw 22
  RandomSoundBumperBottom Bumper3
End Sub


'*******************************************
' Kickers, Saucers
'*******************************************

Sub sw37_Hit
  SoundSaucerLock
  controller.Switch(37)=1
End Sub

Sub SolMBEject (Enabled)
  if enabled Then
  SoundSaucerKick 1, sw37
  saucerwall_phys.IsDropped = True
  sw37.kick 180, Int(Rnd * 10) + 20
  controller.switch(37) = 0
  vpmTimer.AddTimer 1000, "saucerwall_phys.IsDropped = False'"
  end if
End sub

'*******************************************
' Rollovers
'*******************************************

Sub sw17_Hit:Controller.Switch(17) = 1:End Sub
Sub sw17_UnHit:Controller.Switch(17) = 0:End Sub
Sub sw18_Hit:Controller.Switch(18) = 1:End Sub
Sub sw18_UnHit:Controller.Switch(18) = 0:End Sub
Sub sw19_Hit:Controller.Switch(19) = 1:End Sub
Sub sw19_UnHit:Controller.Switch(19) = 0:End Sub
Sub sw20_Hit:Controller.Switch(20) = 1:End Sub
Sub sw20_UnHit:Controller.Switch(20) = 0:End Sub
Sub sw24_Hit:Controller.Switch(24) = 1:End Sub
Sub sw24_UnHit:Controller.Switch(24) = 0:End Sub
Sub sw25_Hit:Controller.Switch(25) = 1:End Sub
Sub sw25_UnHit:Controller.Switch(25) = 0:End Sub
Sub sw26_Hit:Controller.Switch(26) = 1:End Sub
Sub sw26_UnHit:Controller.Switch(26) = 0:End Sub
Sub sw27_Hit:Controller.Switch(27) = 1:End Sub
Sub sw27_UnHit:Controller.Switch(27) = 0:End Sub
Sub sw38_Hit:Controller.Switch(38) = 1:End Sub
Sub sw38_UnHit:Controller.Switch(38) = 0:End Sub

Sub sw31_Hit
  Controller.Switch(31) = 1
  BIPL = 1
  ShooterLaneGuideRail.collidable = 0
End Sub

Sub sw31_UnHit:Controller.Switch(31) = 0: BIPL = 0: End Sub

' Control the guide rail that's supposed to move when ball goes through shooterlane, and under right plastic.
Sub TriggerShooterGate_Hit: ShooterLaneGuideRail.collidable = 1: End Sub
Sub Triggershootergatesound_Hit: RandomSoundWall: End Sub

'*******************************************
' Rollover Animations
'*******************************************
Sub sw17_Animate: psw17.transz = sw17.CurrentAnimOffset: End Sub
Sub sw18_Animate: psw18.transz = sw18.CurrentAnimOffset: End Sub
Sub sw19_Animate: psw19.transz = sw19.CurrentAnimOffset: End Sub
Sub sw20_Animate: psw20.transz = sw20.CurrentAnimOffset: End Sub
Sub sw24_Animate: psw24.transz = sw24.CurrentAnimOffset: End Sub
Sub sw25_Animate: psw25.transz = sw25.CurrentAnimOffset: End Sub
Sub sw26_Animate: psw26.transz = sw26.CurrentAnimOffset: End Sub
Sub sw27_Animate: psw27.transz = sw27.CurrentAnimOffset: End Sub
Sub sw31_Animate: psw31.transz = sw31.CurrentAnimOffset: End Sub
Sub sw38_Animate: psw38.transz = sw38.CurrentAnimOffset: End Sub

'*******************************************
'  Drop and Stand Up Targets
'*******************************************

' Drop Targets

Sub Sw34_Hit: DTHit 34: TargetBouncer Activeball, 1.5: End Sub
Sub Sw35_Hit: DTHit 35: TargetBouncer Activeball, 1.5: End Sub
Sub Sw36_Hit: DTHit 36: TargetBouncer Activeball, 1.5: End Sub


' **** Reset Drop Targets by Solenoid callout

Sub SolDtBank(enabled)
  if enabled then
    RandomSoundDropTargetReset sw35p
    DTRaise 34
    DTRaise 35
    DTRaise 36
    DT34up = 1
    DT35up = 1
    DT36up = 1
  dim xx
    for each xx in ShadowDT
      xx.visible=True
     Next
  end if
End Sub


' Stand Up Targets

Sub sw10_Hit: STHit 10: End Sub
Sub sw11_Hit: STHit 11: End Sub
Sub sw12_Hit: STHit 12: End Sub
Sub sw13_Hit: STHit 13: End Sub
Sub sw14_Hit: STHit 14: End Sub
Sub sw15_Hit: STHit 15: End Sub


'*******************************************
'Spinners
'*******************************************

Sub sw9_Spin
  vpmTimer.PulseSw 9
  SoundSpinner sw9
End Sub

Sub sw16_Spin
  vpmTimer.PulseSw 16
  SoundSpinner sw16
End Sub

'***********Rotate Spinner
Dim SpinnerRadius: SpinnerRadius= 7

Sub SpinnerTimer

  SpinnerPrim2.Rotx = sw9.CurrentAngle
  SpinnerRod2.TransZ = (cos((sw9.CurrentAngle + 180) * (PI/180))+1) * SpinnerRadius
  SpinnerRod2.TransY = sin((sw9.CurrentAngle) * (PI/180)-0.25) * -SpinnerRadius

  SpinnerPrim.Rotx = -sw16.CurrentAngle
  SpinnerRod.TransY = sin((sw16.CurrentAngle) * (PI/180)-0.25) * -SpinnerRadius
  SpinnerRod.TransZ = (cos((sw16.CurrentAngle + 180) * (PI/180))+1) * SpinnerRadius

End Sub


'*******************************************
'  Knocker
'*******************************************

Sub SolKnocker(Enabled)
  If enabled Then
    KnockerSolenoid
  End If
End Sub

Sub SolBell(Enabled)
  If enabled Then
    KnockerBell
  End If
End Sub


'*******************************************
'  Ramp Triggers
'*******************************************

Sub Ramp11Enter_hit()
  WireRampOn True 'Play Plastic Ramp Sound
End Sub

Sub Ramp12Enter_hit()
  WireRampOff ' Turn off the Plastic Ramp Sound
End Sub

Sub Ramp12Enter_unhit()
  WireRampOn False ' On Wire Ramp Play Wire Ramp Sound
End Sub

Sub Ramp12Exit_hit()
  PlaySoundAt "WireRamp_Stop", Ramp12Exit
End Sub



'******************************
' Setup Backglass
'******************************

Dim xoff,yoff,zoff,xrot,zscale, xcen,ycen, ix, xx, yy, xobj

Sub setup_backglass()

  xoff = -20
  yoff = 78
  zoff = 699
  xrot = -90
  zscale = 0.0000001

  xcen = 0  '(130 /2) - (92 / 2)
  ycen = (780 /2 ) + (203 /2)

  for ix = 0 to 31
    For Each xobj In VRDigits(ix)

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


' this is for lining up the backglass flashers on top of a backglass image

  Dim obj

  For Each obj In VRBackglassFlash
    obj.x = obj.x
    obj.height = - obj.y
    obj.y = 78 'adjusts the distance from the backglass towards the user
  Next


End Sub



'*******************************************
'Digital Display
'*******************************************

Dim Digits(34)
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
    Dim ChgLED, ii, jj, num, chg, stat, obj, b, x
    ChgLED=Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
       For ii=0 To UBound(chgLED)
          num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)
        For Each obj In Digits(num)
                   If chg And 1 Then obj.State=stat And 1
                   chg=chg\2 : stat=stat\2
                  Next
    Next
    End If
 End Sub


'******************************************************
' VR Backglass Digital Display
'******************************************************

Dim VRDigits(32)
VRDigits(0) = Array(LED1x0,LED1x1,LED1x2,LED1x3,LED1x4,LED1x5,LED1x6)
VRDigits(1) = Array(LED2x0,LED2x1,LED2x2,LED2x3,LED2x4,LED2x5,LED2x6)
VRDigits(2) = Array(LED3x0,LED3x1,LED3x2,LED3x3,LED3x4,LED3x5,LED3x6)
VRDigits(3) = Array(LED4x0,LED4x1,LED4x2,LED4x3,LED4x4,LED4x5,LED4x6)
VRDigits(4) = Array(LED5x0,LED5x1,LED5x2,LED5x3,LED5x4,LED5x5,LED5x6)
VRDigits(5) = Array(LED6x0,LED6x1,LED6x2,LED6x3,LED6x4,LED6x5,LED6x6)
VRDigits(6) = Array(LED7x0,LED7x1,LED7x2,LED7x3,LED7x4,LED7x5,LED7x6)

VRDigits(7) = Array(LED8x0,LED8x1,LED8x2,LED8x3,LED8x4,LED8x5,LED8x6)
VRDigits(8) = Array(LED9x0,LED9x1,LED9x2,LED9x3,LED9x4,LED9x5,LED9x6)
VRDigits(9) = Array(LED10x0,LED10x1,LED10x2,LED10x3,LED10x4,LED10x5,LED10x6)
VRDigits(10) = Array(LED11x0,LED11x1,LED11x2,LED11x3,LED11x4,LED11x5,LED11x6)
VRDigits(11) = Array(LED12x0,LED12x1,LED12x2,LED12x3,LED12x4,LED12x5,LED12x6)
VRDigits(12) = Array(LED13x0,LED13x1,LED13x2,LED13x3,LED13x4,LED13x5,LED13x6)
VRDigits(13) = Array(LED14x0,LED14x1,LED14x2,LED14x3,LED14x4,LED14x5,LED14x6)

VRDigits(14) = Array(LED1x000,LED1x001,LED1x002,LED1x003,LED1x004,LED1x005,LED1x006)
VRDigits(15) = Array(LED1x100,LED1x101,LED1x102,LED1x103,LED1x104,LED1x105,LED1x106)
VRDigits(16) = Array(LED1x200,LED1x201,LED1x202,LED1x203,LED1x204,LED1x205,LED1x206)
VRDigits(17) = Array(LED1x300,LED1x301,LED1x302,LED1x303,LED1x304,LED1x305,LED1x306)
VRDigits(18) = Array(LED1x400,LED1x401,LED1x402,LED1x403,LED1x404,LED1x405,LED1x406)
VRDigits(19) = Array(LED1x500,LED1x501,LED1x502,LED1x503,LED1x504,LED1x505,LED1x506)
VRDigits(20) = Array(LED1x600,LED1x601,LED1x602,LED1x603,LED1x604,LED1x605,LED1x606)

VRDigits(21) = Array(LED2x000,LED2x001,LED2x002,LED2x003,LED2x004,LED2x005,LED2x006)
VRDigits(22) = Array(LED2x100,LED2x101,LED2x102,LED2x103,LED2x104,LED2x105,LED2x106)
VRDigits(23) = Array(LED2x200,LED2x201,LED2x202,LED2x203,LED2x204,LED2x205,LED2x206)
VRDigits(24) = Array(LED2x300,LED2x301,LED2x302,LED2x303,LED2x304,LED2x305,LED2x306)
VRDigits(25) = Array(LED2x400,LED2x401,LED2x402,LED2x403,LED2x404,LED2x405,LED2x406)
VRDigits(26) = Array(LED2x500,LED2x501,LED2x502,LED2x503,LED2x504,LED2x505,LED2x506)
VRDigits(27) = Array(LED2x600,LED2x601,LED2x602,LED2x603,LED2x604,LED2x605,LED2x606)

VRDigits(28) = Array(LEDax300,LEDax301,LEDax302,LEDax303,LEDax304,LEDax305,LEDax306)
VRDigits(29) = Array(LEDbx400,LEDbx401,LEDbx402,LEDbx403,LEDbx404,LEDbx405,LEDbx406)
VRDigits(30) = Array(LEDcx500,LEDcx501,LEDcx502,LEDcx503,LEDcx504,LEDcx505,LEDcx506)
VRDigits(31) = Array(LEDdx600,LEDdx601,LEDdx602,LEDdx603,LEDdx604,LEDdx605,LEDdx606)



dim DisplayColor, DisplayColorG
DisplayColor =  RGB(255,40,1)

Sub VRDisplayTimer
  Dim ii, jj, obj, b, x
  Dim ChgLED,num, chg, stat
  ChgLED=Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED) Then
      For ii=0 To UBound(chgLED)
        num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)
        For Each obj In VRDigits(num)
 '                  If chg And 1 Then obj.visible=stat And 1    'if you use the object color for off; turn the display object visible to not visible on the playfield, and uncomment this line out.
          If chg And 1 Then FadeDisplay obj, stat And 1
          chg=chg\2 : stat=stat\2
        Next
      Next
    End If
End Sub

Sub FadeDisplay(object, onoff)
  If OnOff = 1 Then
    object.Color = DisplayColor
    Object.Opacity = 16
  Else
    Object.Color = RGB(5,10,10)
    Object.Opacity = 1
  End If
End Sub

Sub InitDigits()
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


If VR_Room = 1 Then
  InitDigits
End If



' ******************************************************************************************
'      LAMP CALLBACK for the 6 backglass flasher lamps (not the solenoid conrolled ones)
' ******************************************************************************************

if VR_Room = 0 and cab_mode = 0 Then
  Set LampCallback = GetRef("UpdateDTLamps")
End If

if VR_Room = 1 Then
  Set LampCallback = GetRef("UpdateVRLamps")
End If


Sub UpdateDTLamps()
  CreditsREEL.setValue(1)
  If Controller.Lamp(1) = 0 Then: GameOverReel.setValue(0):   Else: GameOverReel.setValue(1) 'Game Over
  If Controller.Lamp(2) = 0 Then: MatchReel.setValue(0):      Else: MatchReel.setValue(1) 'Match
  If Controller.Lamp(3) = 0 Then: TiltReel.setValue(0):     Else: TiltReel.setValue(1) 'Tilt
  If Controller.Lamp(4) = 0 Then: HighScoreReel.setValue(0):    Else: HighScoreReel.setValue(1) 'High Score
  If Controller.Lamp(5) = 0 Then: ShootAgainReel.setValue(0):   Else: ShootAgainReel.setValue(1) 'Shoot Again
  If Controller.Lamp(6) = 0 Then: BIPReel.setValue(0):      Else: BIPReel.setValue(1) 'Ball in Play
End Sub

Sub UpdateVRLamps()
  If Controller.Lamp(1) = 0 Then: f1bg.visible = 0:       Else: f1bg.visible = 1  'Game Over
  If Controller.Lamp(2) = 0 Then: f2bg.visible = 0:       Else: f2bg.visible = 1  'Match
  If Controller.Lamp(3) = 0 Then: f3bg.visible = 0:       Else: f3bg.visible = 1  'Tilt
  If Controller.Lamp(4) = 0 Then: f4bg.visible = 0:       Else: f4bg.visible = 1  'High Score
  If Controller.Lamp(5) = 0 Then: f5bg.visible = 0:       Else: f5bg.visible = 1  'Shoot Again
  If Controller.Lamp(6) = 0 Then: f6bg.visible = 0:       Else: f6bg.visible = 1  'Ball in Play
End Sub


' controls the VR Backglass text for game over, BIP, etc.  Attempt to time the change when GI changes.
dim vrbgopacity

Sub VRBGtext
  vrbgopacity = (1200 - (1175*VRgilvl))
    f1bg.opacity = vrbgopacity
    f2bg.opacity = vrbgopacity
    f3bg.opacity = vrbgopacity
    f4bg.opacity = vrbgopacity
    f5bg.opacity = vrbgopacity
    f6bg.opacity = vrbgopacity
End Sub


'******************************************************
' ZNFF:  FLIPPER CORRECTIONS by nFozzy
'******************************************************
'

'******************************************************
' Flippers Polarity (Select appropriate sub based on era)
'******************************************************

Dim LF : Set LF = New FlipperPolarity
Dim RF : Set RF = New FlipperPolarity
Dim ULF : Set ULF = New FlipperPolarity

InitPolarity

'
''*******************************************
'' Mid 80's
'
Sub InitPolarity()
   dim x, a : a = Array(LF, RF)
  for each x in a
    x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
    x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
    x.enabled = True
    x.TimeDelay = 80
    x.DebugOn=False ' prints some info in debugger

    x.AddPt "Polarity", 0, 0, 0
    x.AddPt "Polarity", 1, 0.05, - 3.7
    x.AddPt "Polarity", 2, 0.16, - 3.7
    x.AddPt "Polarity", 3, 0.22, - 0
    x.AddPt "Polarity", 4, 0.25, - 0
    x.AddPt "Polarity", 5, 0.3, - 2
    x.AddPt "Polarity", 6, 0.4, - 3
    x.AddPt "Polarity", 7, 0.5, - 3.7
    x.AddPt "Polarity", 8, 0.65, - 2.3
    x.AddPt "Polarity", 9, 0.75, - 1.5
    x.AddPt "Polarity", 10, 0.81, - 1
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
    ULF.SetObjects "ULF", LeftFlipper1, TriggerULF

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

Dim LFPress, RFPress, ULFPress, LFCount, RFCount
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

Const LiveCatch = 16
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

Dim DT34, DT35, DT36


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
' animate:      Arrary slot for handling the animation instrucitons, set to 0
'
' Values for annimate: 1 - bend target (hit to primary), 2 - drop target (hit to secondary), 3 - brick target (high velocity hit to secondary), -1 - raise target

Set DT34 = (new DropTarget)(sw34, sw34a, sw34p, 34, 0, false)
Set DT35 = (new DropTarget)(sw35, sw35a, sw35p, 35, 0, false)
Set DT36 = (new DropTarget)(sw36, sw36a, sw36p, 36, 0, false)

Dim DTArray
DTArray = Array(DT34, DT35, DT36)


'Configure the behavior of Drop Targets.
Const DTDropSpeed = 90 'in milliseconds
Const DTDropUpSpeed = 40 'in milliseconds
Const DTDropUnits = 44 'VP units primitive drops so top of at or below the playfield
Const DTDropUpUnits = 10 'VP units primitive raises above the up position on drops up
Const DTMaxBend = 2 'max degrees primitive rotates when hit
Const DTDropDelay = 20 'time in milliseconds before target drops (due to friction/impact of the ball)
Const DTRaiseDelay = 40 'time in milliseconds before target drops back to normal up position after the solenoid fires to raise the target
Const DTBrickVel = 30 'velocity at which the target will brick, set to '0' to disable brick
Const DTEnableBrick = 0 '0 'Set to 0 to disable bricking, 1 to enable bricking

Const DTHitSound = "" 'Drop Target Hit sound
Const DTDropSound = "DropTarget_Down" 'Drop Target Drop sound
Const DTResetSound = "DropTarget_Up" 'Drop Target reset sound

Const DTMass = 0.2 'Mass of the Drop Target (between 0 and 1), higher values provide more resistance



'******************************************************
'  DROP TARGETS FUNCTIONS
'******************************************************

Dim DTShadow(3)

DTShadowInit 1
DTShadowInit 2
DTShadowInit 3


' Initializes the drop targets for shadow logic below
Sub DTShadowInit(dtnbr)

  if dtnbr = 1 Then
    Set DTShadow(dtnbr) = Eval("dtsh" & 34)
  elseif dtnbr = 2 Then
    Set DTShadow(dtnbr) = Eval("dtsh" & 35)
  elseif dtnbr = 3 Then
    Set DTShadow(dtnbr) = Eval("dtsh" & 36)
  End If
End Sub


Sub DTHit(switch)
  Dim i, swmod

  If switch = 34 Then
    swmod = 1
  Elseif switch = 35 then
    swmod = 2
  Elseif switch = 36 then
    swmod = 3
  End If

  i = DTArrayID(switch)

  PlayTargetSound
  DTArray(i).animate = DTCheckBrick(ActiveBall,DTArray(i).prim)
  If DTArray(i).animate = 1 Or DTArray(i).animate = 3 Or DTArray(i).animate = 4 Then
    DTBallPhysics ActiveBall, DTArray(i).prim.rotz, DTMass

' Controls Drop Shadow for a direct hit only
      if swmod = 1 Then
        DT34up = 0
        DTShadow(1).visible = 0
      Elseif swmod = 2 Then
        DT35up = 0
        DTShadow(2).visible = 0
      Elseif swmod = 3 Then
        DT36up = 0
        DTShadow(3).visible = 0
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
      'Dim gBOT
      'gBOT = GetBalls

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
'  DROP TARGET
'  SUPPORTING FUNCTIONS
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

Dim ST10, ST11, ST12, ST13, ST14, ST15

'Set array with stand-up target objects
'
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

Set ST10 = (new StandupTarget)(sw10, psw10,10, 0)
Set ST11 = (new StandupTarget)(sw11, psw11,11, 0)
Set ST12 = (new StandupTarget)(sw12, psw12,12, 0)
Set ST13 = (new StandupTarget)(sw13, psw13,13, 0)
Set ST14 = (new StandupTarget)(sw14, psw14,14, 0)
Set ST15 = (new StandupTarget)(sw15, psw15,15, 0)

Dim STArray
STArray = Array(ST10, ST11, ST12, ST13, ST14, ST15)

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

'****************************************************************
'****  END VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'****************************************************************



'******************************************************
'   Z3DI:   3D INSERTS
'******************************************************

dim insert_white_on: insert_white_on = insertfactor * 250
dim insert_red_on: insert_red_on = insertfactor * 200
dim insert_orange_on: insert_orange_on = insertfactor * 200
dim insert_green_on: insert_green_on = insertfactor * 180

dim insert_red_tri_on: insert_red_tri_on = insertfactor * 60
dim insert_orange_tri_on: insert_orange_on = insertfactor * 200
dim insert_green_tri_on: insert_green_tri_on = insertfactor * 60

dim inprim    ' Setting the inserts off lighting.  default of 1 is a little too low.
  for each inprim in InsertPrimsOff
    inprim.blenddisablelighting = 1.5
  Next


'************************
'**** Red Inserts

Sub li8_animate: p8.BlendDisableLighting = insert_red_on * (li8.GetInPlayIntensity / li8.Intensity): End Sub
Sub li9_animate: p9.BlendDisableLighting = insert_red_tri_on * (li9.GetInPlayIntensity / li9.Intensity): End Sub
Sub li10_animate: p10.BlendDisableLighting = insert_red_tri_on * (li10.GetInPlayIntensity / li10.Intensity): End Sub
Sub li11_animate: p11.BlendDisableLighting = insert_red_tri_on * (li11.GetInPlayIntensity / li11.Intensity): End Sub
Sub li12_animate: p12.BlendDisableLighting = insert_red_tri_on * (li12.GetInPlayIntensity / li12.Intensity): End Sub
Sub li13_animate: p13.BlendDisableLighting = insert_red_tri_on * (li13.GetInPlayIntensity / li13.Intensity): End Sub
Sub li14_animate: p14.BlendDisableLighting = insert_red_tri_on * (li14.GetInPlayIntensity / li14.Intensity): End Sub
Sub li15_animate: p15.BlendDisableLighting = insert_red_tri_on * (li15.GetInPlayIntensity / li15.Intensity): End Sub
Sub li16_animate: p16.BlendDisableLighting = insert_red_tri_on * (li16.GetInPlayIntensity / li16.Intensity): End Sub
Sub li22_animate: p22.BlendDisableLighting = insert_red_tri_on * (li22.GetInPlayIntensity / li22.Intensity): End Sub
Sub li23_animate: p23.BlendDisableLighting = insert_red_on * (li23.GetInPlayIntensity / li23.Intensity): End Sub
Sub li24_animate: p24.BlendDisableLighting = insert_red_on * (li24.GetInPlayIntensity / li24.Intensity): End Sub
Sub li25_animate: p25.BlendDisableLighting = insert_red_on * (li25.GetInPlayIntensity / li25.Intensity): End Sub
Sub li28_animate: p28.BlendDisableLighting = insert_red_on * (li28.GetInPlayIntensity / li28.Intensity): End Sub
Sub li29_animate: p29.BlendDisableLighting = insert_red_on * (li29.GetInPlayIntensity / li29.Intensity): End Sub
Sub li30_animate: p30.BlendDisableLighting = insert_red_on * (li30.GetInPlayIntensity / li30.Intensity): End Sub
Sub li42_animate: p42.BlendDisableLighting = insert_red_on * (li42.GetInPlayIntensity / li42.Intensity): End Sub
Sub li43_animate: p43.BlendDisableLighting = insert_red_on * (li43.GetInPlayIntensity / li43.Intensity): End Sub
Sub li44_animate: p44.BlendDisableLighting = insert_red_on * (li44.GetInPlayIntensity / li44.Intensity): End Sub
Sub li45_animate: p45.BlendDisableLighting = insert_red_on * (li45.GetInPlayIntensity / li45.Intensity): End Sub
Sub li46_animate: p46.BlendDisableLighting = insert_red_on * (li46.GetInPlayIntensity / li46.Intensity): End Sub
Sub li55_animate: p55.BlendDisableLighting = insert_red_on * (li55.GetInPlayIntensity / li55.Intensity): End Sub

'************************
'**** White Inserts

Sub li26_animate: p26.BlendDisableLighting = insert_white_on * (li26.GetInPlayIntensity / li26.Intensity): End Sub
Sub li27_animate: p27.BlendDisableLighting = insert_white_on * (li27.GetInPlayIntensity / li27.Intensity): End Sub
Sub li33_animate: p33.BlendDisableLighting = insert_white_on * (li33.GetInPlayIntensity / li33.Intensity): End Sub
Sub li34_animate: p34.BlendDisableLighting = insert_white_on * (li34.GetInPlayIntensity / li34.Intensity): End Sub
Sub li35_animate: p35.BlendDisableLighting = insert_white_on * (li35.GetInPlayIntensity / li35.Intensity): End Sub
Sub li36_animate: p36.BlendDisableLighting = insert_white_on * (li36.GetInPlayIntensity / li36.Intensity): End Sub
Sub li37_animate: p37.BlendDisableLighting = insert_white_on * (li37.GetInPlayIntensity / li37.Intensity): End Sub
Sub li38_animate: p38.BlendDisableLighting = insert_white_on * (li38.GetInPlayIntensity / li38.Intensity): End Sub
Sub li39_animate: p39.BlendDisableLighting = insert_white_on * (li39.GetInPlayIntensity / li39.Intensity): End Sub
Sub li40_animate: p40.BlendDisableLighting = insert_white_on * (li40.GetInPlayIntensity / li40.Intensity): End Sub
Sub li41_animate: p41.BlendDisableLighting = insert_white_on * (li41.GetInPlayIntensity / li41.Intensity): End Sub
Sub li53_animate: p53.BlendDisableLighting = insert_white_on * (li53.GetInPlayIntensity / li53.Intensity): End Sub

'************************
'**** Orange Inserts

Sub li7_animate: p7.BlendDisableLighting = insert_orange_on * (li7.GetInPlayIntensity / li7.Intensity): End Sub
Sub li21_animate: p21.BlendDisableLighting = insert_orange_tri_on * (li21.GetInPlayIntensity / li21.Intensity): End Sub
Sub li31_animate: p31.BlendDisableLighting = insert_orange_on * (li31.GetInPlayIntensity / li31.Intensity): End Sub
Sub li32_animate: p32.BlendDisableLighting = insert_orange_on * (li32.GetInPlayIntensity / li32.Intensity): End Sub

'************************
'**** Green Inserts

Sub li17_animate: p17.BlendDisableLighting = insert_green_on * (li17.GetInPlayIntensity / li17.Intensity): End Sub
Sub li18_animate: p18.BlendDisableLighting = insert_green_on * (li18.GetInPlayIntensity / li18.Intensity): End Sub
Sub li19_animate: p19.BlendDisableLighting = insert_green_on * (li19.GetInPlayIntensity / li19.Intensity): End Sub
Sub li20_animate: p20.BlendDisableLighting = insert_green_on * (li20.GetInPlayIntensity / li20.Intensity): End Sub
Sub li47_animate: p47.BlendDisableLighting = insert_green_tri_on * (li47.GetInPlayIntensity / li47.Intensity): End Sub
Sub li48_animate: p48.BlendDisableLighting = insert_green_tri_on * (li48.GetInPlayIntensity / li48.Intensity): End Sub
Sub li49_animate: p49.BlendDisableLighting = insert_green_on * (li49.GetInPlayIntensity / li49.Intensity): End Sub
Sub li50_animate: p50.BlendDisableLighting = insert_green_on * (li50.GetInPlayIntensity / li50.Intensity): End Sub
Sub li51_animate: p51.BlendDisableLighting = insert_green_on * (li51.GetInPlayIntensity / li51.Intensity): End Sub
Sub li52_animate: p52.BlendDisableLighting = insert_green_on * (li52.GetInPlayIntensity / li52.Intensity): End Sub
Sub li54_animate: p54.BlendDisableLighting = insert_green_on * (li54.GetInPlayIntensity / li54.Intensity): End Sub


'************************
'**** Backwall and VR Backglass Eyes Flasher

Sub l_f8_animate
  for each bgbulb in f8flash
    bgbulb.opacity = 100 * (l_f8.GetInPlayIntensity / l_f8.Intensity)
  Next

  if VR_Room = 1 Then
    for each bgbulb in f8VRflash
      bgbulb.visible = 1
      FlBGL28.opacity = 350 * (l_f8.GetInPlayIntensity / l_f8.Intensity)
      FlBGL29.opacity = 350 * (l_f8.GetInPlayIntensity / l_f8.Intensity)
      FlBGL30.opacity = 500 * (l_f8.GetInPlayIntensity / l_f8.Intensity)
      FlBGL31.opacity = 500 * (l_f8.GetInPlayIntensity / l_f8.Intensity)
    Next
  End If
End Sub


'************************
'**** VR Backglass Flash

If VR_Room = 1 Then

dim bgbulb

  for each bgbulb in BGLampCntl: bgbulb.visible = 1: Next

  Sub l56_animate: FlBGL49.opacity = 350 * (l56.GetInPlayIntensity / l56.Intensity): FlBGL50.opacity = 350 * (l56.GetInPlayIntensity / l56.Intensity): End Sub
  Sub l57_animate: FlBGL47.opacity = 350 * (l57.GetInPlayIntensity / l57.Intensity): FlBGL48.opacity = 350 * (l57.GetInPlayIntensity / l57.Intensity): End Sub
  Sub l58_animate: FlBGL45.opacity = 350 * (l58.GetInPlayIntensity / l58.Intensity): FlBGL46.opacity = 350 * (l58.GetInPlayIntensity / l58.Intensity): End Sub
  Sub l59_animate: FlBGL43.opacity = 350 * (l59.GetInPlayIntensity / l59.Intensity): FlBGL44.opacity = 350 * (l59.GetInPlayIntensity / l59.Intensity): End Sub
  Sub l60_animate: FlBGL41.opacity = 350 * (l60.GetInPlayIntensity / l60.Intensity): FlBGL42.opacity = 350 * (l60.GetInPlayIntensity / l60.Intensity): End Sub
  Sub l61_animate: FlBGL33.opacity = 350 * (l61.GetInPlayIntensity / l61.Intensity): FlBGL34.opacity = 350 * (l61.GetInPlayIntensity / l61.Intensity): End Sub
  Sub l62_animate: FlBGL35.opacity = 350 * (l62.GetInPlayIntensity / l62.Intensity): FlBGL36.opacity = 350 * (l62.GetInPlayIntensity / l62.Intensity): End Sub
  Sub l63_animate: FlBGL37.opacity = 350 * (l63.GetInPlayIntensity / l63.Intensity): FlBGL38.opacity = 350 * (l63.GetInPlayIntensity / l63.Intensity): End Sub
  Sub l64_animate: FlBGL39.opacity = 350 * (l64.GetInPlayIntensity / l64.Intensity): FlBGL40.opacity = 350 * (l64.GetInPlayIntensity / l64.Intensity): End Sub

  Sub l_gi_animate
    for each bgbulb in GIBackglass
      if VR_Room = 0 Then
        bgbulb.visible = 0
      Else
        bgbulb.visible = 1
        bgbulb.opacity = 350 * (l_gi.GetInPlayIntensity / l_gi.Intensity)
      End If
    Next
  End Sub

  Sub f5flash_animate:  FlBGL27.opacity = 350 * (f5flash.GetInPlayIntensity / f5flash.Intensity): End Sub
  Sub f19flash_animate:   FlBGL32.opacity = 350 * (f19flash.GetInPlayIntensity / f19flash.Intensity): End Sub

Else
  for each bgbulb in GIBackglass: bgbulb.visible = 0: Next
  for each bgbulb in BGLampCntl:  bgbulb.visible = 0: Next

End If



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
Dim RampBalls(3,2)
'x,0 = ball x,1 = ID, 2 = Protection against ending early (minimum amount of updates)

'0,0 is boolean on/off, 0,1 unused for now
RampBalls(0,0) = False

' RampType
' Setup: Set this array to the number Total number of balls that can be tracked at one time + 1.  5 ball multiball then set value to 6
' Description: Array type indexed on BallId and a values used to deterimine what type of ramp the ball is on: False = Wire Ramp, True = Plastic Ramp
Dim RampType(3)

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

Sub RandomSoundRampStop(obj)
  Select Case Int(rnd*3)
    Case 0: PlaySoundAtVol "wireramp_stop1", obj, 0.2*VolumeDial:PlaySoundAtLevelActiveBall ("Rubber_Strong_1"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
    Case 1: PlaySoundAtVol "wireramp_stop2", obj, 0.2*VolumeDial:PlaySoundAtLevelActiveBall ("Rubber_Strong_2"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
    Case 2: PlaySoundAtVol "wireramp_stop3", obj, 0.2*VolumeDial:PlaySoundAtLevelActiveBall ("Rubber_1_Hard"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
  End Select
End Sub

'******************************************************
'**** END RAMP ROLLING SFX
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
Dim FlipperUpSoundLevel, FlipperDownSoundLevel, FlipperLeftHitParm, FlipperRightHitParm, FlipperLeft1HitParm
Dim SlingshotSoundLevel, BumperSoundFactor, KnockerSoundLevel

FlipperUpAttackMinimumSoundLevel = 0.010    'volume level; range [0, 1]
FlipperUpAttackMaximumSoundLevel = 0.635    'volume level; range [0, 1]
FlipperUpSoundLevel = 1.0            'volume level; range [0, 1]
FlipperDownSoundLevel = 0.45          'volume level; range [0, 1]
FlipperLeftHitParm = FlipperUpSoundLevel    'sound helper; not configurable
FlipperRightHitParm = FlipperUpSoundLevel    'sound helper; not configurable
FlipperLeft1HitParm = FlipperUpSoundLevel   'sound helper; not configurable
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

' Thalamus, AudioFade - Patched
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

' Thalamus, AudioPan - Patched
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

Sub KnockerBell()
  PlaySoundAtLevelStatic SoundFX("Bell",DOFKnocker), KnockerSoundLevel, KnockerPosition
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

Sub LeftFlipper1Collide(parm)
  FlipperLeft1HitParm = parm / 10
  If FlipperLeft1HitParm > 1 Then
    FlipperLeft1HitParm = 1
  End If
  FlipperLeft1HitParm = FlipperUpSoundLevel * FlipperLeft1HitParm
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
Dim objBallShadow(2)

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



'******************************************************
' Ball brightness code
'******************************************************

If ChangeBall = 1 Then
  if BallLightness = 0 Then
    table1.BallImage="ball-dark"
    table1.BallFrontDecal="JPBall-Scratches"
  elseif BallLightness = 1 Then
    table1.BallImage="ball_HDR"
    table1.BallFrontDecal="Scratches"
  elseif BallLightness = 2 Then
    table1.BallImage="ball-light-hf"
    table1.BallFrontDecal="scratchedmorelight"
  elseif BallLightness = 3 Then
    table1.BallImage="ball-lighter-hf"
    table1.BallFrontDecal="scratchedmorelight"
  elseif BallLightness = 4 Then
    table1.BallImage="ball_MRBallDark"
    table1.BallFrontDecal="ball_scratches"
  elseif BallLightness = 5 Then
    table1.BallImage="ball"
    table1.BallFrontDecal="ball_scratches"
  elseif BallLightness = 6 Then
    table1.BallImage="ball_HDR"
    table1.BallFrontDecal="ball_scratches"
  elseif BallLightness = 7 Then
    table1.BallImage="ball_HDR_bright"
    table1.BallFrontDecal="ball_scratches"
  elseif BallLightness = 8 Then
    table1.BallImage="ball_HDR3b"
    table1.BallFrontDecal="ball_scratches"
  elseif BallLightness = 9 Then
    table1.BallImage="ball_HDR3b2"
    table1.BallFrontDecal="ball_scratches"
  elseif BallLightness = 10 Then
    table1.BallImage="ball_HDR3b3"
    table1.BallFrontDecal="ball_scratches"
  elseif BallLightness = 11 Then
    table1.BallImage="ball_HDR3dark"
    table1.BallFrontDecal="ball_scratches"
  elseif BallLightness = 12 Then
    table1.BallImage="BorgBall"
    table1.BallFrontDecal="ball_scratches"
  elseif BallLightness = 13 Then
    table1.BallImage="SteelBall2"
    table1.BallFrontDecal="ball_scratches"
  else
    table1.BallImage="ball"
    table1.BallFrontDecal="ball_scratches"
  End If
End If




'******************************************************
' Options
'******************************************************

Sub SetupOptions

Dim C

'  ROM Volume Settings
    With Controller
        .Games(cGameName).Settings.Value("volume") = ROMVolume
  End With

' Flipper Color
  If FlipperColor = 1 Then
    For Each C in GIFlip
      C.image = "williamsbatwhiteblack"
    Next
  Else
    For Each C in GIFlip
      C.image = "williamsbatwhitered"
    Next
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

' Spinner Color
  If SpinnerColor = 0 Then
    SpinnerPrim.image = "Spinner-Reversed"
    SpinnerPrim2.image = "Spinner"
  Else
    SpinnerPrim.image = "Spinner-Reversed-glitter"
    SpinnerPrim2.image = "Spinner-glitter"
  End If

' Instruction  Color
  If CardOption = 0 Then
    cardleft.image = "cardleft-black"
    cardright.image = "cardright-black"
  Else
    cardleft.image = "cardleft-white"
    cardright.image = "cardright-white"
  End If


End Sub

'******************************************************
' ZVRR
'******************************************************

Sub SetupRoom

Dim VRThings

if VR_Room = 0 and cab_mode = 0 Then
  for each VRThings in VRStuff:VRThings.visible = 0:Next
  for each VRThings in VRClock:VRThings.visible = 0:Next
  for each VRThings in VRBackglass:VRThings.visible = 0:Next
  for each VRThings in TopHardware:VRThings.visible = 1:Next
  OuterPrimBlack_DT.z = -251
  L_DT_1.State = 1
  L_DT_2.State = 1
Elseif VR_Room = 0 and cab_mode = 1 Then
  for each VRThings in VRStuff:VRThings.visible = 0:Next
  for each VRThings in VRClock:VRThings.visible = 0:Next
  for each VRThings in VRBackglass:VRThings.visible = 0:Next
  for each VRThings in DTBackglass:VRThings.visible = 0: Next
  for each VRThings in DTBGDisplay:VRThings.visible = 0: Next
  for each VRThings in TopHardware:VRThings.visible = 0:Next
    if cabsideblades = 0 Then
      OuterPrimBlack_DT.z = -151
    Else
      OuterPrimBlack_DT.z = -251
    End If
  L_DT_1.Visible = 0
  L_DT_1.State = 0
  L_DT_2.Visible = 0
  L_DT_2.State = 0
Else
  for each VRThings in VRStuff:VRThings.visible = 1:Next
  for each VRThings in VRClock:VRThings.visible = WallClock:Next
  for each VRThings in DTBackglass:VRThings.visible = 0: Next
  for each VRThings in DTBGDisplay:VRThings.visible = 0: Next
  for each VRThings in TopHardware:VRThings.visible = 1:Next
  OuterPrimBlack_DT.z = -251
  L_DT_1.Visible = 0
  L_DT_1.State = 0
  L_DT_2.Visible = 0
  L_DT_2.State = 0

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
'****  ZGIU:  GI Control
'******************************************************

'**** Global variable to hold current GI light intensity. Not Mandatory, but may help in other lighting subs and timers
'**** This value is updated always when GI state is being updated

dim gilvl, VRgilvl

Sub SolGi(enabled)
Dim x

' debug.print "SolGi value: " & Enabled

    If Enabled Then
    gilvl = 0
    vpmTimer.AddTimer 50, "VRgilvl = 0'"

    Sound_GI_Relay 0, Bumper3

    ' Set Drop Target Shadows to off when GI is goes off.
      for each x in ShadowDT
        x.visible=0
      Next

    L_DT_1.intensity = 0
    L_DT_2.intensity = 0

    Else
    gilvl = 1
    VRgilvl = 1

    Sound_GI_Relay 1, Bumper3 'Note: Bumper3 is just used for sound positioning. Can be anywhere that makes sense.

    ' Set Drop Target Shadows to state when GI is back on.
      if DT34up=1 Then dtsh34.visible = 1
      if DT35up=1 Then dtsh35.visible = 1
      if DT36up=1 Then dtsh36.visible = 1

    If DesktopMode = True then
      L_DT_1.intensity = 5
      L_DT_2.intensity = 2.5
    Else
      L_DT_1.intensity = 0
      L_DT_2.intensity = 0
    End If
End If

  SetRoomBrightness LightLevel

End Sub



'******************************************************
'****  END GI Control
'******************************************************




' VR Room Mods for Version 1.0 by UnclePaulie

'******** Revisions done on VR version 1.0

' - Added all the VR Cab primatives (Used from Defender VR and Comet VR, as similar Williams Table); adjusted as appropriate
' - Matched cabinet colors for flipper buttons (normally black, but also over a metal side rail, so made them red)
' - Score alphanumeric displays and code Used from Fathom VR Table and rearranged
' - Backglass Image from Hauntfreak's B2S file... Note:  ONLY the image.  The rest of display done in code.
' - Added plunger primary and code for movement
' - Added two different VR room environments
' - Added GI to backglass
' - Ramp lrail, rrail, 15 and 16 removed
' - Changed the material of Wall8 to black MetalGuide
' - Lowered the lane triggers suface height by -20, by tying to *switch surface
' - Changed the ball image and added jp_scratches for decal to enhance the ball rolling realism
' - Changed the flipper color to white (the real table is white... at least the ones I saw on google.)
' - Sixtoe recommended changing the GI intenisity of GI1, 3, and 17 to 10 (was 40).  Looks much better in VR.

'******** Revisions done on VR version 1.1
' - Moved the flipper buttons in.
' - The top flipper was still yellow.  Changed to white
' - Modified the backglass to have lighted interactivity (added primative and several flashers)
' - Added fastflips by solenoids = 2

'******** Revisions done on VR version 1.2
' - Added Material _noextrashadingvr to the backglass
' - Changed trigger material and the legs, lockdownbar, plunger, and rails to Metal4
' - Cleaned up the backglass flasher code and removed preexisting backglass Lights
' - Completely animated the backglass lamps
' - Rawd added ball shadows

'******** Revisions done by UnclePaulie on Hybrid version 0.1 - 2.0 *********

' v0.1: Enabled desktop, VR, and cabinet modes
'     Added lighter ball images provided by hauntfreaks and g5k ball scratches, changeable in script
' v0.2  Fixed the sol8 backglass and eyes working.  Needed to create a metal black dynamic material for side rails.
'     Need to ensure that Metal Black, Plastic Black, and Plastic materials are not active for the eyes back flasher to work.
'   Created a dynamic playfield material to be active
'   Added a sol8subroutine to handle both the dragon eyes on the playfield and 4 backglass flashers.
' v0.3  Added LUT options with save option
'     Updated the playfield for slings, triggers holes, and target holes, and added plywood hole images.
'   Modified the sling primatives and sling walls to drop below playfield.  Looks better in VR.
'   Added an optional VR Clock, animated the plunger for VR keyboard use, and animated the flippers
'   Had to change materials on each flipper button to animate properly
' v0.4: Added Fleep Sounds, nFozzy / Roth physics and flippers.  Updated the flipper physics per recommendations.
'     Updated ramprolling sounds by fluffhead35 as well as ballroll and ramproll amplication sounds
'   Moved flippers slightly to match closer to real table.  Also adjusted a couple posts on lower playfield to match real table.
'     Physics (mid 80's);  table physics, gravity constant.  Also sling and bumper thresholds changed to match table physics
'   Added VPW Dynamic Ball Shadows to the solution done Iakki, Apophis, and Wylte
'   Added ramp triggers for ramp rolling sounds
' v0.5: Added new drop targets to the solution done by Rothbauerw, as well as drop target shadows
'   Slightly adjusted the ramp entrance, and the sleeve location.  It wasn't aligned just right with playfield.
'   Added a couple walls behind stand up target.  Ball got behind once.
' v0.6  Added a LUT flasher for VR, so you can see what LUT is being chosen when in VR Mode.
'   Commented out drop target shadows... couldn't see anyway, as playfield dard, and it made lights too dim by targets
'   Enabled show mesh on bulb, and broughht halo heighht down.  Looks better in VR.  More work could be done later to them.
' v0.7  Corrected the demon light and bonus holdover light being swapped error.  Li22 and Li23 were swapped.  (thanks to Lumi for finding)
'   Adjusted the plunger physics settings to increase the pull speed, length, and adjust the strength.(Thanks PinStratsDan)
'   Adjusted some of the flipper physics settings (EOST, and EOSReturn, friction, and strength) (Thanks PinStratsDan)
'   Very minor adjustment to the outlane posts and drain post location.
'   Made the plastics collidable to prevent a bounced ball from going in.
' v0.8  Increased the visibility of the dynamic ball shadows.
'   Rawd added a glass and scratches.
'   Added code to make the glass and scratches optional.
' v0.9  Glass and scratches only work in VR now.  Not used for cab or desktop
'   Turned rails off in cab mode
'   Changed EOST to 0.4, to help with flipper tricks (small flicks and cradle separation)
' v0.10 Changed EOST to 0.3, and return strength to .09
' v0.11 Added ramp primative from Bord.  Updated to Bord's playfield, but added holes back in.
' v0.12 Added flasher blooms, new primitives for gi on and off from Bord.  Updated the GI lights to Bord's GI update.
'   Imported Bord's materials and images
'   Changed playfield scatter to 0.1 and elements scatter to 12.
'   Reduced the wire ramp friction slightly, as ball moved very slow.
' v0.13 Changed environment_shiny3blur4 environmental image.
'   Corrected z fighting on walls with new primitives
'   Bord recommended a few mods.  Turned Gate 2 visibility off.  Removed flipper adjustment pins
' v0.14 Modified cab slightly in VR to accomodate for new flashers
' v0.15 Changed playfield material to active, for see through holes in VR.
' v2.0  Release Version
' v2.1  Added Wall16 side image of "shadow" to stop seeing a white strip under backwall demon eyes.
'   Adjusted primative29 location slightly up.  Also, put EOS Torque on flippers to .275, strength to 1800, friction to .9, and return strength to .07.
'   Adjusted the friction on ramps down slightly to .06.  Also the elasticity to 0.2.
'   Adjusted the saucer kicker sw37 to strength 25, and kickforce variance to 5 to kick out towards upper flipper a bit more.
'   Adjusted size of post048, so it won't interfere with ramp, based on feedback from Sixtoe.
'   Adjusted MetalGuide3 out a bit... was causing a weird bounce back off the ramp coming down.  Thanks to Wylte.
'   Adjusted EOSReturn to 0.035, based on suggestion from Rothbauerw.
' v2.2  Fixed flippers still being activated after a game ends.
'   Automated desktop and cabinet mode.  Still need to manually select if in VR mode.
'   Ensured the ROM sounds were activated (others mod tables turned PINMAME sounds off, and didn't turn back on when exited)
' v2.3  Redid backdrop image for desktop mode (not as busy).
'   Redid the text boxes on backdrop to Reels, as some resolutions don't size text boxes correctly.
'   Moved the desktop backglass dragon flasher slightly.
' v2.4  Updated to latest Fleep Sounds.
'   Added Apophis sling corrections, and updated the physics.
'   Changed code for ALL VPW physics, dampener, shadows, rubberizer, bouncer, rolling sounds, etc.
'   Fixed Ramp11 rolling sound
'   Changed nudge sensitivity to 6
'     Added updated knocker position and code
' v2.5  Completely changed the trough logic to handle balls on table
'   Changed the apron wall structure to create a trough
'     No longer destroy balls.  Use a global gBOT throughout all of script.
'   Removed all the getballs calls and am only useing gBOT.  Will help eliminate stutter on slower CPU machines
'   Added groove to plunger lane, and small ball release ramp to get up.
'   Simplified .hidden to only DesktopMode
'   Reduced the number of dynamice sources to 6.  Was 12, didn't need that many.
'   Added updated GI relay sounds
'   Updated drop target hit and reset sounds and subs.
'   Updated the use of the TargetBounce collection for usage on targets.
' v2.6  Drop target shadows working
'   Added updated standup targets with VPW physics, original code created by Rothbauerw and VPW team
'   Added GI brightness to all targets
'   Changed each stand up target to it's own material, needed for hit animation.
'   Added addtional balls with varying brightness
'   Added basic ball GI brightness level
' v2.7  Added updated playfield and insert text for 3D inserts
'   Updated to Lampz lighting routine by VPW.
'   Adjusted all the insert light intensities, fade, falloff, heights, colors, etc.
'   Adjusted the flasher fading up and down to match real table.
' v2.8  Added insert light blooms
' v2.9  Corrected issue with GI states and Lampz.
'   Slight flipper tweak to EOSReturn, and Return Strength
'     Added option for cabinet side blades
'   Added VR Topper, VR Posters/flyers, and VR Cord/outlet
'   Corrected reflectivity issue with top left plate gate.  Weird reflections...
'   Reduced the width of the triggers a little.
'   Cleaned up the setup backglass code for VR
' v2.10 Added missing rubber physics wall behind drop targets
'   Added GI Brightness increase to inserts when GI is off.
' v2.11 Added new Williams flipper primitives and had added GI levels to them
' v2.12 Added GI brightness to the white screws, bumpers, rubbers, and updated the targets GI to new dl method
' v2.13 Added GI levels to VR Backglass
'   Added varying Lampz aLvl brightness fade levels for backglass lamps, and playfield flashers
' v2.14 Updated POV for cab use
'   Increased the BIP, game over, etc. lights in VR backglass slightly.
'   Adjusted the white insert lights (1-8) intensity down slightly. Triangle lights up slightly. Top prims up slightly.
'   Reduced the amount of insert intensity update when GI on.
'   Adjusted the Apron, Plunger Cover, Plunger, and Plunger Guide for GI brightness levels.
' v2.15 Adjusted the bloom shapes, the green color, the intensities, and the falloff radius, per Apophis recommendations.
' v2.16 Increased GI intensity from 8 to 12; and increased default bulb intensity scale to 2 (per Apophis and nFozzy recommendations)
' v3.0  Released Version

'3.1.1  Updated to latest Fleep Sounds
'   Removed LUTs
'   Cleaned up layers
'   Corrected dimensions on table.  Requires everything to be moved.
'   Updated physics materials, playfield overall physics, playfield lighting
'   Added all new ball images and options.  Removed old dynamic ball shadows code, images, and material.
'   Updated targets to latest VPW code, and standalone capable, and updated target physics materials, also ensure dtshadows only work on direct hit
'   Automated the desktop, VR, and cab mode
'   Updated script to latest VPW script
'   Added options menu via F12
'   Updated flipper animations and ballbrightness subs
'   Updated nudging parameters.
'   Updated flipper physics and tricks
'   Implemented some recommendations from apophis on flippers and target physics
'   Added a new desktop backglass and animations
'   Updated playfield, plastics, bumper, target, spinner images from redbone.
'   Moved all elements to correct positions
'   Changed y scale of all triangle prims to 100 (was 380). And the z height to 0 (was -3.5) Could see inside prim element.
'   Changed all insert glow light intensity to 15.
'     Added new backrail prim and updated the backwall image
'   Added new rollover prims, adjusted the cutouts, and trigger sizes.
'     Redid plunger, plunger groove, pluger metal guide stop (needed a new Prim)
'   Updated VR Room environments, removed topper, added grill, improved LED on VR backglass
'   Updated peg primitives, physics, material, image, transparency, and GI control
'   Updated slings, sling corrections and lower peg no bounce
'   Removed old baked walls and shadows and GI
'   Added new walls and new lane tops
'   Put new sized rubbers in and updated the physics collections on the rubber physics and all that.
'     Added leaf sensor prims, and all new guard rails and physics.
'   Added new long gates, and animated them.  Also wire gate.
'     Added guide rail under right plastic in shooter lane, as well as associated logic and sounds.
'   Added spinner prim and animations
'   Added new lanereturn guides for base metal and top acrylic plastic.  Adjusted materials as well.
'   Added new apron and plunger cover primitives, and added associated physics.
'   Added all plywood cutouts and GI bulb prims.
'   Adjusted the ramp width at bottom and side walls to match prim better. Lowered friction on wire ramp. Follows physics of tables watched online.
'   Changed ramp walls to real walls.
'   Added all the screws, screw caps, and inside peg posts
'   Updated, rotated, and aligned all the targets.
'   Added new GI lights
'3.1.2  Added 3D inserts, had to add orange inserts (missed before).
'   Have GI come on after 2 seconds.
'   Removed Lampz.  Implemented AllLamps fading routines built into VPinmame
'   Adjusted colors on the inserts look correct.
'   Added playfield mesh for saucer
'   Updated GI routines
'   Modified the VR Backglass GI for text.  Added delay in gilvl variable for VR.
'   Added new flipper primitives and images.
'   Set the inserts off lighting to 1.5.  default of 1 was a little too low.
'   Added a new outerprim wall with image from redbone
'   All metal walls have new images from redbone
'   Redid the dimensions of the VR Cab body, backbox, rails, lockdownbar, etc.  Now fits the inner cab walls.
'   Using same lockdownbar and side rails for desktop and VR.  New siderails with metal side plates
'   Adjusted desktop pov
'   Updated to my standard materials for playfield, targets.
'   Adjusted room and cab brightness levels to be controlled via F12 menu for day/night
'   Removed visible glass and impurities.  Added only glass physics.
'   Added table ruleset and info
'   Small update to target shadows.
'   Redbone updated the plastics and text overlay images.  Changed overlay text material to black base.
'   Added new acrylics to plastics.  Reshaped and sized plastics.
'   Updated the GI to be controlled by not only GI, but also F12 lightlevel changes.
'   Added GI tops to plastics.
'   New bumper primitives, images, materials, and tied to both GI and room brightness.
'3.1.3  Baked the GI and flashers
'   Added options for flippers, spinners, instruction cards, and screwcaps
'   Added bumper colors to gi and roombrightness sub.
'3.1.4  Fixed lamp assigment for L47
'   Moved upper left sleeve and near by rubber and post slightly for better physics ball roll.
'   Update to plastics and playfield and bulb material
'   Added GI to metal walls.
'   Defaulted ROM sounds to 0.  If you use alternate sounds, need at zero.
'   Updated cab pov for legacy
'   Fixed apron side wall edge hanging over.
'3.1.5  Redbone updated plastic colors in some locations.
'   Minor physics adjustment to ramp11 to ensure slow roll down accurate.
'   Update to bumper gi light transmit setting and bumper tops translucency for day/night slider affect.
'   Added different bumper base and skirt to stay on at day/night slider adjustments.
'   Changed to a graphic for F12 instructions / rules
'   Primetime5k added staged flipper support; flipper sub clean-up. UP updated sections to match example table
'   Updated ramp11 top width as you could see ball slightly through ramp side.
'   Moved top part of left ramp prim slightly to cover too much light leaks.
'4.0.0  Converted images to webp
'   Corrected cab side blade option
'   Removed unused images and materials.
'   Moved plunger cover slightly.
'   Put translucency of bumpers back to 0.
'   Released
'4.0.1  Minor graphics update... upper right plastic had a piece protruding through the edge of the ramp.
'   Rewrote the instructions to modify the ROM.
'4.0.2  Updated right card images for 3 ball.
'4.0.3  Added missing solenoid routine for missing knocker1 (wasn't in manual), added solenoid routine for bell.

' Thanks to Redbone for the graphics updates, and apophis for physics testing and feedback! Also Primetime5k for staged flipper support.

'4.0.4  Corrected sling switch pulses to 32 and 33.  (error came in on a cut and paste)



'------------------------------------------------------------
'Operator Menu Settings & Instructions in the ROM
'------------------------------------------------------------

'Start table as usual
'Press (8) to get into ROM menu
'Press (7) to start correct advancement
'Press (8) several times until 04 is shown in the Credit Display.  'Auto-Up / Manual-Down switch is controlled by (7)
'Press (8) repeatedly until you get to the adjustment number desired.
'   For example, background sound is defaulted to on.  It is adjustment number 35.  You'll see a 1 in the upper left display.  Press (1) to change that value.   To turn off sound, set the value to 0.
' Note:  Press (7) to reverse the advancement up or down.
'(9) will exit/reset

'Note: Auto-Up / Manual-Down (7)  also controls the direction the menus will advance


'Bookkeeping Displays (04 is in Credit Display; function numbers below appear in the Ball-In-Play Display; sub-values appear in player score displays)
'--------------------
'00 Game Identification (2535 1)
'01 Coins, Left Chute (closest to coin-door hinge)
'02 Coins, Center Chute
'03 Coins, Right Chute
'04 Total Paid-Credits
'05 Special Credits
'06 Replay-Score Credits
'07 Match Credits
'08 Total Credits
'09 Total Extra Balls
'10 Ball Time in Minutes
'11 Total Balls Played
'12 High Scores
'13 Backup High-Scores

'42 Times MULTI-BALL play was achieved
'43 Number of 3Xs in MULTI-BALL play
'44 Number of 5Xs in MULTI-BALL play
'45 Number of Ramps over 5X
'46 Number of Bonus Holdovers
'47 Number of Playfield Specials (from spotting S-O-R-C-E-R-E-R)
'48 Times A-B-C-D was completed

'Adjustments (factory defaults in brackets [])
'-----------
'14 Replay-Level 1 [1,000,000]
'15 Replay-Level 2 (or 2nd highest score) [00 = off]
'16 Replay-Level 3 (or 3rd highest score) [00 = off]
'17 Replay-Level 4 (or 4th highest socre) [00 = off]
'18 Maximum Credits [30]
'19 Standard and Custom Pricing-Control [01/02]
'20 Left   Coin-Slot Multiplier [01]
'21 Center Coin-Slot Multiplier [04/10]
'22 Right  Coin-Slot Multiplier [01/03]
'23 Coin Units Required For Credit [01]
'24 Units Required For Bonus Credit [00]
'25 Minimum Coin-Units [00]

'26 Match [00]
'   00: Standard Match (awards 10% replays)
'   01: Off

'27 Special [00]
'     00: Awards Credit
'     01: Awards Extra Ball
'   02: Awards Points

'28 Replay [00]
'    00: Awards Credit
'     01: Awards Extra Ball
'     02: No Award

'29 Maximum Plumb-Bob Tilts (including warnings) [03]
'30 Number of Balls [03]

'31 Game Adjustment #1 - Extra Ball Lamp in Memory [00]
'     00:  No-Conservative
'     01: Yes-Liberal

'32 Game Adjustment #2 - Reset Interval For Drop Targets [05]
'     00-09
'     00: Fast
'     09: Slow

'33 Game Adjustment #3 - Bell [01]
'     00: Off
'     01: On

'34 Game Adjustment #4 - Blink Interval for Drop Target Lights [05]
'     00-09
'     00: Fast
'     09: Slow

'35 Game Adjustment #5 - Background Sound [01]
'     00: Off
'     01: On

'36 Game Adjustment #7 - Sound for Attract Mode [01]
'     00: Off
'     01: On

'37 Game Adjustment #8 - Bonus Holdover in Memory [00]
'     00: Off
'     01: On

'38 Game Adjustment #9 - S-O-R-C-E-R-E-R Target Memory [00]
'     00:  No-Conservative
'     01: Yes-Liberal

'40 Maximum high-score credits [03]
'   00: Displays high scores without credit payouts

'50 Special Function
'   15: Auto-Cycle Mode
'   35: Zero bookkeeping totals
'   45: Restore factory settings & zero bookkeeping totals

'-------------------------------
'End of Operator Adjustment List
'-------------------------------

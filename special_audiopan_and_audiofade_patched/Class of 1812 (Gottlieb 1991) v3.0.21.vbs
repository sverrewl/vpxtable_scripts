' Class of 1812 / IPD No. 528 / August, 1991 / 4 Players
'
'  ,- _~. ,,                                 /\        /|   /\\   /|   /\
' (' /|   ||   _                            ||        /||  || || /||  (  )
'((  ||   ||  < \,  _-_,  _-_,        /'\\ =||=        ||   \ /   ||    //
'((  ||   ||  /-|| ||_.  ||_.        || ||  ||         ||   /\\   ||   //
' ( / |   || (( ||  ~ ||  ~ ||       || ||  ||         ||  // \\  ||  /(
'  -____- \\  \/\\ ,-_-  ,-_-        \\,/   \\,       ,/-' || || ,/-' {___
'                                                           \\/
'

' Version 3.0 by UnclePaulie 2023
' Table includes Hybrid VR/desktop/cabinet modes, enhanced playfield by Redbone, VPW physics, Fleep sounds, Lampz, 3D inserts, new GI,
'   sling corrections, targets, saucers, playfield mesh, etc.
' Tomate created a VLM and UnclePaulie updated the script.
' Updated Version 2 by UnclePaulie, which was originally done by Tom Tower & Ninuzzu
' Complete implementation details and credits found at the bottom of the script.

' Version 3.0.15 by UnclePaulie and Tomate 2024
' Tomate did VLM baking, and UnclePaule incorporated the associated script and updates to pwm and VPW standards.


' **************************************************************
' ******  Gotlieb Add 3 Credits Issue or other ROM issues ******
' **************************************************************

' - If your ROM is adding 3 credits... you need to go into service menu and change center coin slot to 1 credit
'   - Press 7 when starting the game to enter ROM menu
'   - Immediately Press 8 to enter game adjustments
'   - Press 7 to advance to the adjustment you want to modify.  The center slot credit menu is adjustment number 11.
'   - Press 8 or 9 to increase or decrease the setting.  (In this case, hit 8 twice, to go from 3 to 1)
'   - Press F3 to save

' ***   If the ROM has all the settings blank (i.e. high score, max credits, add credits, etc.)
'     This is due to a blank .nv file.  All the settings are defaulted to 0

' - Follow the instructions above, but you need to modify all the settings.  1 coin = 1 credit.  Max 3 credits,  replay credit, etc.


'*********************************************************************************************************************************
' === TABLE OF CONTENTS  ===
'
' You can quickly jump to a section by searching the four letter tag (ZXXX)
'
' ZVAR: Constants and Global Variable
' ZTIM: Timers
' ZINI: Table Initialization and Exiting
' ZOPT: User Options
'   ZMAT: General Math Functions
' ZANI: Misc Animations
'   ZBBR: Ball Brightness
' ZKEY: Key Press Handling
' ZSOL: Solenoids & Flashers
' ZDRN: Drain, Trough, and Ball Release
' ZFLP: Flippers
' ZSLG: Slingshot Animations
' ZSSC: Slingshot Corrections
' ZSWI: Switches
' ZVUK: VUKs and Kickers
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


Option Explicit
Randomize
SetLocale 1033

'******************************************************
' Minimum Requirements
'******************************************************

' VPX: version 10.8.0 RC3 or higher
' VPM: version 3.6.0 or higher
' B2S: version 2.1.1 or higher (if using b2s)


'*******************************************
'  ZOPT: User Options
'*******************************************

' ****************************************************
' Desktop, Cab, and VR OPTIONS
' ****************************************************

' Desktop, Cab, and VR Room are automatically selected.  However if in VR Room mode, you can change the environment with the magna save buttons.

' ****************************************************
' ALL OPTIONS ARE IN THE F12 MENU, EXCEPT FOR THE BALL COLOR BELOW

const ChangeBall  = 1   '0 = Ball changes are done via the drop down table menu
              '1 = ball selection per the below BallLightness option (default)
const BallLightness = 7   '0 = dark, 1 = not as dark, 2 = bright, 3 = brightest, 4 = MrBallDark, 5 = DarthVito Ball1
              '6 = DarthVito HDR1, 7 = DarthVito HDR Bright (Default), 8 = DarthVito HDR2, 9 = DarthVito ballHDR3,  10 = DarthVito ballHDR4
              '11 = DarthVito ballHDRdark, 12 = Borg Ball, 13 = SteelBall2


' ****************************************************
' *** If playing in desktop mode:
' ****************************************************

const DTBGAnimation = 1 '0 = static desktop mode backglass, 1 = desktop mode lamps interactive

'******************************************************
' VR Room Options
'******************************************************

const WallClock   = 1   '1 Shows the clock in the VR minimal rooms only
const topper    = 1   '0 = Off 1= On - Topper visible in VR Room only
const poster    = 1   '1 Shows the flyer posters in the VR room only
const poster2     = 1   '1 Shows the flyer posters in the VR room only
const CustomWalls   = 2   '0 = UP's Original Minimal Walls, floor, and roof
              '1 = Sixtoe's arcade style
              '2 = DarthVito's Updated home walls with lights
              '3 = DarthVito's plaster home walls
              '4 = DarthVito's blue home walls
Const FrontKeys   = 0   '1 = Enables Magna Buttons for front Left / Right VR Buttons

'******************************************************
' F12 Menu Options
'******************************************************

Dim LightLevel : LightLevel = 0.25        ' Level of room lighting (0 to 1), where 0 is dark and 100 is brightest
Dim ColorLUT : ColorLUT = 1           ' Color desaturation LUTs: 1 to 11, where 1 is normal and 11 is black'n'white
Dim VolumeDial : VolumeDial = 0.8             ' Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Dim BallRollVolume : BallRollVolume = 0.5     ' Level of ball rolling volume. Value between 0 and 1
Dim RampRollVolume : RampRollVolume = 0.5   ' Level of ramp rolling volume. Value between 0 and 1

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

    ' Sound volumes
    VolumeDial = Table1.Option("Mech Volume", 0, 1, 0.01, 0.8, 1)
    BallRollVolume = Table1.Option("Ball Roll Volume", 0, 1, 0.01, 0.5, 1)
  RampRollVolume = Table1.Option("Ramp Roll Volume", 0, 1, 0.01, 0.5, 1)

  ' Room brightness
' LightLevel = Table1.Option("Table Brightness (Ambient Light Level)", 0, 1, 0.01, .5, 1)
  LightLevel = NightDay/100

  If eventId = 3 And dspTriggered Then dspTriggered = False : DisableStaticPreRendering = False : End If
End Sub


'******************************************************
'  ZVAR: Constants and Global Variables
'******************************************************

Dim VR_Room, cab_mode, DesktopMode: DesktopMode = Table1.ShowDT
If RenderingMode = 2 Then VR_Room=1 Else VR_Room=0      'VRRoom set based on RenderingMode starting in version 10.72
if Not DesktopMode and VR_Room=0 Then cab_mode=1 Else cab_mode=0

Const UseSolenoids  = 2
Const cSingleLFlip  = 0
Const cSingleRFlip  = 0
Const UseLamps = 1        '0 = Custom lamp handling, 1 = Built-in VPX handling (using light number in light timer)
Const UseGI     = 0
Const cGameName   = "clas1812"
Const UsingROM    = True    'The UsingROM flag is to indicate code that requires ROM usage. Mostly for instructional purposes only.
Const BallSize    = 50
Const BallMass    = 1
Const UseSync   = 0
Const tnob = 2          ' total number of balls : 2 (trough)
Const lob = 0         ' Locked balls at table start

Const UseVPMModSol = 2      'Set to 2 for PWM flashers, inserts, and GI. Requires VPinMame 3.6

Dim tablewidth: tablewidth = Table1.width
Dim tableheight: tableheight = Table1.height
Dim i, cBall1, cball2, gBOT
Dim BIPL : BIPL = False       ' Ball in plunger lane
Dim gilvl: gilvl = 1


Dim DT16up, DT17up, DT26up, DT27up, DT36up, DT37up, DT46up, DT47up


On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "02800000", "Class1812.VBS", 3.50


'******************************************************
' VLM Arrays
'******************************************************

' VLM  Arrays - Start
' Arrays per baked part
Dim BP_BottomTeeth: BP_BottomTeeth=Array(BM_BottomTeeth, LM_FL_Flasherflash2_BottomTeeth, LM_FL_Flasherflash4_BottomTeeth, LM_FL_Flasherflash5_BottomTeeth, LM_IN_L36_BottomTeeth, LM_IN_L67_BottomTeeth, LM_IN_Li2_BottomTeeth, LM_GI_BottomTeeth)
Dim BP_Layer0: BP_Layer0=Array(BM_Layer0, LM_FL_Flasherflash1_Layer0, LM_FL_Flasherflash2_Layer0, LM_FL_Flasherflash3_Layer0, LM_FL_Flasherflash4_Layer0, LM_FL_Flasherflash5_Layer0, LM_IN_L46_Layer0, LM_IN_L47_Layer0, LM_IN_L67_Layer0, LM_IN_Li10_Layer0, LM_IN_Li11_Layer0, LM_IN_Li12_Layer0, LM_IN_Li13_Layer0, LM_IN_Li14_Layer0, LM_IN_Li2_Layer0, LM_IN_bumperbiglight1_Layer0, LM_IN_bumperbiglight2_Layer0, LM_FL_f25_Layer0, LM_GIS_gi_001_Layer0, LM_GIS_gi_002_Layer0, LM_GIS_gi_003_Layer0, LM_GIS_gi_004_Layer0, LM_GIS_gi_005_Layer0, LM_GIS_gi_006_Layer0, LM_GIS_gi_007_Layer0, LM_GIS_gi_008_Layer0, LM_GI_Layer0)
Dim BP_Layer1: BP_Layer1=Array(BM_Layer1, LM_FL_Flasherflash1_Layer1, LM_FL_Flasherflash2_Layer1, LM_FL_Flasherflash3_Layer1, LM_FL_Flasherflash4_Layer1, LM_FL_Flasherflash5_Layer1, LM_IN_L36_Layer1, LM_IN_L37_Layer1, LM_IN_L46_Layer1, LM_IN_L47_Layer1, LM_IN_L67_Layer1, LM_IN_Li10_Layer1, LM_IN_Li11_Layer1, LM_IN_Li12_Layer1, LM_IN_Li13_Layer1, LM_IN_Li14_Layer1, LM_IN_Li2_Layer1, LM_IN_bumperbiglight1_Layer1, LM_IN_bumperbiglight2_Layer1, LM_IN_bumperbiglight3_Layer1, LM_FL_f25_Layer1, LM_GI_Layer1)
Dim BP_Layer2: BP_Layer2=Array(BM_Layer2, LM_FL_Flasherflash1_Layer2, LM_FL_Flasherflash2_Layer2, LM_FL_Flasherflash3_Layer2, LM_FL_Flasherflash4_Layer2, LM_FL_Flasherflash5_Layer2, LM_IN_L67_Layer2, LM_IN_Li12_Layer2, LM_IN_Li2_Layer2, LM_IN_bumperbiglight1_Layer2, LM_IN_bumperbiglight2_Layer2, LM_FL_f25_Layer2, LM_GI_Layer2)
Dim BP_Layer3: BP_Layer3=Array(BM_Layer3, LM_FL_Flasherflash1_Layer3, LM_FL_Flasherflash2_Layer3, LM_FL_Flasherflash4_Layer3, LM_FL_Flasherflash5_Layer3, LM_IN_L36_Layer3, LM_IN_L67_Layer3, LM_IN_Li10_Layer3, LM_IN_Li11_Layer3, LM_IN_Li12_Layer3, LM_IN_Li13_Layer3, LM_IN_Li14_Layer3, LM_IN_Li2_Layer3, LM_GI_Layer3)
Dim BP_Parts: BP_Parts=Array(BM_Parts, LM_FL_Flasherflash1_Parts, LM_FL_Flasherflash2_Parts, LM_FL_Flasherflash3_Parts, LM_FL_Flasherflash4_Parts, LM_FL_Flasherflash5_Parts, LM_IN_L36_Parts, LM_IN_L37_Parts, LM_IN_L46_Parts, LM_IN_L47_Parts, LM_IN_L67_Parts, LM_IN_Li10_Parts, LM_IN_Li11_Parts, LM_IN_Li12_Parts, LM_IN_Li13_Parts, LM_IN_Li14_Parts, LM_IN_Li2_Parts, LM_IN_bumperbiglight1_Parts, LM_IN_bumperbiglight2_Parts, LM_IN_bumperbiglight3_Parts, LM_FL_f25_Parts, LM_GIS_gi_001_Parts, LM_GIS_gi_002_Parts, LM_GIS_gi_003_Parts, LM_GIS_gi_004_Parts, LM_GIS_gi_005_Parts, LM_GIS_gi_006_Parts, LM_GIS_gi_007_Parts, LM_GIS_gi_008_Parts, LM_GI_Parts, LM_IN_l3_Parts, LM_IN_l4_Parts)
Dim BP_Playfield: BP_Playfield=Array(BM_Playfield, LM_FL_Flasherflash1_Playfield, LM_FL_Flasherflash2_Playfield, LM_FL_Flasherflash3_Playfield, LM_FL_Flasherflash4_Playfield, LM_FL_Flasherflash5_Playfield, LM_IN_L67_Playfield, LM_IN_Li11_Playfield, LM_IN_Li12_Playfield, LM_IN_Li13_Playfield, LM_IN_Li2_Playfield, LM_IN_bumperbiglight1_Playfield, LM_IN_bumperbiglight2_Playfield, LM_IN_bumperbiglight3_Playfield, LM_GIS_gi_001_Playfield, LM_GIS_gi_002_Playfield, LM_GIS_gi_003_Playfield, LM_GIS_gi_004_Playfield, LM_GIS_gi_005_Playfield, LM_GIS_gi_006_Playfield, LM_GIS_gi_007_Playfield, LM_GIS_gi_008_Playfield, LM_GI_Playfield, LM_IN_l3_Playfield)
Dim BP_Prim_LeftFlipper: BP_Prim_LeftFlipper=Array(BM_Prim_LeftFlipper, LM_GIS_gi_001_Prim_LeftFlipper, LM_GIS_gi_002_Prim_LeftFlipper, LM_GIS_gi_003_Prim_LeftFlipper, LM_GIS_gi_004_Prim_LeftFlipper, LM_GIS_gi_005_Prim_LeftFlipper, LM_GIS_gi_006_Prim_LeftFlipper, LM_GIS_gi_007_Prim_LeftFlipper, LM_GIS_gi_008_Prim_LeftFlipper, LM_GI_Prim_LeftFlipper)
Dim BP_Prim_RightFlipper: BP_Prim_RightFlipper=Array(BM_Prim_RightFlipper, LM_GIS_gi_001_Prim_RightFlipper, LM_GIS_gi_002_Prim_RightFlipper, LM_GIS_gi_003_Prim_RightFlipper, LM_GIS_gi_004_Prim_RightFlipper, LM_GIS_gi_005_Prim_RightFlipper, LM_GIS_gi_006_Prim_RightFlipper, LM_GIS_gi_007_Prim_RightFlipper, LM_GIS_gi_008_Prim_RightFlipper, LM_GI_Prim_RightFlipper)
Dim BP_heart: BP_heart=Array(BM_heart, LM_FL_Flasherflash1_heart, LM_FL_Flasherflash2_heart, LM_FL_Flasherflash3_heart, LM_FL_Flasherflash4_heart, LM_FL_Flasherflash5_heart, LM_IN_L67_heart, LM_IN_Li2_heart, LM_FL_f25_heart, LM_GI_heart)
' Arrays per lighting scenario
Dim BL_FL_Flasherflash1: BL_FL_Flasherflash1=Array(LM_FL_Flasherflash1_Layer0, LM_FL_Flasherflash1_Layer1, LM_FL_Flasherflash1_Layer2, LM_FL_Flasherflash1_Layer3, LM_FL_Flasherflash1_Parts, LM_FL_Flasherflash1_Playfield, LM_FL_Flasherflash1_heart)
Dim BL_FL_Flasherflash2: BL_FL_Flasherflash2=Array(LM_FL_Flasherflash2_BottomTeeth, LM_FL_Flasherflash2_Layer0, LM_FL_Flasherflash2_Layer1, LM_FL_Flasherflash2_Layer2, LM_FL_Flasherflash2_Layer3, LM_FL_Flasherflash2_Parts, LM_FL_Flasherflash2_Playfield, LM_FL_Flasherflash2_heart)
Dim BL_FL_Flasherflash3: BL_FL_Flasherflash3=Array(LM_FL_Flasherflash3_Layer0, LM_FL_Flasherflash3_Layer1, LM_FL_Flasherflash3_Layer2, LM_FL_Flasherflash3_Parts, LM_FL_Flasherflash3_Playfield, LM_FL_Flasherflash3_heart)
Dim BL_FL_Flasherflash4: BL_FL_Flasherflash4=Array(LM_FL_Flasherflash4_BottomTeeth, LM_FL_Flasherflash4_Layer0, LM_FL_Flasherflash4_Layer1, LM_FL_Flasherflash4_Layer2, LM_FL_Flasherflash4_Layer3, LM_FL_Flasherflash4_Parts, LM_FL_Flasherflash4_Playfield, LM_FL_Flasherflash4_heart)
Dim BL_FL_Flasherflash5: BL_FL_Flasherflash5=Array(LM_FL_Flasherflash5_BottomTeeth, LM_FL_Flasherflash5_Layer0, LM_FL_Flasherflash5_Layer1, LM_FL_Flasherflash5_Layer2, LM_FL_Flasherflash5_Layer3, LM_FL_Flasherflash5_Parts, LM_FL_Flasherflash5_Playfield, LM_FL_Flasherflash5_heart)
Dim BL_FL_f25: BL_FL_f25=Array(LM_FL_f25_Layer0, LM_FL_f25_Layer1, LM_FL_f25_Layer2, LM_FL_f25_Parts, LM_FL_f25_heart)
Dim BL_GI: BL_GI=Array(LM_GI_BottomTeeth, LM_GI_Layer0, LM_GI_Layer1, LM_GI_Layer2, LM_GI_Layer3, LM_GI_Parts, LM_GI_Playfield, LM_GI_Prim_LeftFlipper, LM_GI_Prim_RightFlipper, LM_GI_heart)
Dim BL_GIS_gi_001: BL_GIS_gi_001=Array(LM_GIS_gi_001_Layer0, LM_GIS_gi_001_Parts, LM_GIS_gi_001_Playfield, LM_GIS_gi_001_Prim_LeftFlipper, LM_GIS_gi_001_Prim_RightFlipper)
Dim BL_GIS_gi_002: BL_GIS_gi_002=Array(LM_GIS_gi_002_Layer0, LM_GIS_gi_002_Parts, LM_GIS_gi_002_Playfield, LM_GIS_gi_002_Prim_LeftFlipper, LM_GIS_gi_002_Prim_RightFlipper)
Dim BL_GIS_gi_003: BL_GIS_gi_003=Array(LM_GIS_gi_003_Layer0, LM_GIS_gi_003_Parts, LM_GIS_gi_003_Playfield, LM_GIS_gi_003_Prim_LeftFlipper, LM_GIS_gi_003_Prim_RightFlipper)
Dim BL_GIS_gi_004: BL_GIS_gi_004=Array(LM_GIS_gi_004_Layer0, LM_GIS_gi_004_Parts, LM_GIS_gi_004_Playfield, LM_GIS_gi_004_Prim_LeftFlipper, LM_GIS_gi_004_Prim_RightFlipper)
Dim BL_GIS_gi_005: BL_GIS_gi_005=Array(LM_GIS_gi_005_Layer0, LM_GIS_gi_005_Parts, LM_GIS_gi_005_Playfield, LM_GIS_gi_005_Prim_LeftFlipper, LM_GIS_gi_005_Prim_RightFlipper)
Dim BL_GIS_gi_006: BL_GIS_gi_006=Array(LM_GIS_gi_006_Layer0, LM_GIS_gi_006_Parts, LM_GIS_gi_006_Playfield, LM_GIS_gi_006_Prim_LeftFlipper, LM_GIS_gi_006_Prim_RightFlipper)
Dim BL_GIS_gi_007: BL_GIS_gi_007=Array(LM_GIS_gi_007_Layer0, LM_GIS_gi_007_Parts, LM_GIS_gi_007_Playfield, LM_GIS_gi_007_Prim_LeftFlipper, LM_GIS_gi_007_Prim_RightFlipper)
Dim BL_GIS_gi_008: BL_GIS_gi_008=Array(LM_GIS_gi_008_Layer0, LM_GIS_gi_008_Parts, LM_GIS_gi_008_Playfield, LM_GIS_gi_008_Prim_LeftFlipper, LM_GIS_gi_008_Prim_RightFlipper)
Dim BL_IN_L36: BL_IN_L36=Array(LM_IN_L36_BottomTeeth, LM_IN_L36_Layer1, LM_IN_L36_Layer3, LM_IN_L36_Parts)
Dim BL_IN_L37: BL_IN_L37=Array(LM_IN_L37_Layer1, LM_IN_L37_Parts)
Dim BL_IN_L46: BL_IN_L46=Array(LM_IN_L46_Layer0, LM_IN_L46_Layer1, LM_IN_L46_Parts)
Dim BL_IN_L47: BL_IN_L47=Array(LM_IN_L47_Layer0, LM_IN_L47_Layer1, LM_IN_L47_Parts)
Dim BL_IN_L67: BL_IN_L67=Array(LM_IN_L67_BottomTeeth, LM_IN_L67_Layer0, LM_IN_L67_Layer1, LM_IN_L67_Layer2, LM_IN_L67_Layer3, LM_IN_L67_Parts, LM_IN_L67_Playfield, LM_IN_L67_heart)
Dim BL_IN_Li10: BL_IN_Li10=Array(LM_IN_Li10_Layer0, LM_IN_Li10_Layer1, LM_IN_Li10_Layer3, LM_IN_Li10_Parts)
Dim BL_IN_Li11: BL_IN_Li11=Array(LM_IN_Li11_Layer0, LM_IN_Li11_Layer1, LM_IN_Li11_Layer3, LM_IN_Li11_Parts, LM_IN_Li11_Playfield)
Dim BL_IN_Li12: BL_IN_Li12=Array(LM_IN_Li12_Layer0, LM_IN_Li12_Layer1, LM_IN_Li12_Layer2, LM_IN_Li12_Layer3, LM_IN_Li12_Parts, LM_IN_Li12_Playfield)
Dim BL_IN_Li13: BL_IN_Li13=Array(LM_IN_Li13_Layer0, LM_IN_Li13_Layer1, LM_IN_Li13_Layer3, LM_IN_Li13_Parts, LM_IN_Li13_Playfield)
Dim BL_IN_Li14: BL_IN_Li14=Array(LM_IN_Li14_Layer0, LM_IN_Li14_Layer1, LM_IN_Li14_Layer3, LM_IN_Li14_Parts)
Dim BL_IN_Li2: BL_IN_Li2=Array(LM_IN_Li2_BottomTeeth, LM_IN_Li2_Layer0, LM_IN_Li2_Layer1, LM_IN_Li2_Layer2, LM_IN_Li2_Layer3, LM_IN_Li2_Parts, LM_IN_Li2_Playfield, LM_IN_Li2_heart)
Dim BL_IN_bumperbiglight1: BL_IN_bumperbiglight1=Array(LM_IN_bumperbiglight1_Layer0, LM_IN_bumperbiglight1_Layer1, LM_IN_bumperbiglight1_Layer2, LM_IN_bumperbiglight1_Parts, LM_IN_bumperbiglight1_Playfield)
Dim BL_IN_bumperbiglight2: BL_IN_bumperbiglight2=Array(LM_IN_bumperbiglight2_Layer0, LM_IN_bumperbiglight2_Layer1, LM_IN_bumperbiglight2_Layer2, LM_IN_bumperbiglight2_Parts, LM_IN_bumperbiglight2_Playfield)
Dim BL_IN_bumperbiglight3: BL_IN_bumperbiglight3=Array(LM_IN_bumperbiglight3_Layer1, LM_IN_bumperbiglight3_Parts, LM_IN_bumperbiglight3_Playfield)
Dim BL_IN_l3: BL_IN_l3=Array(LM_IN_l3_Parts, LM_IN_l3_Playfield)
Dim BL_IN_l4: BL_IN_l4=Array(LM_IN_l4_Parts)
Dim BL_World: BL_World=Array(BM_BottomTeeth, BM_Layer0, BM_Layer1, BM_Layer2, BM_Layer3, BM_Parts, BM_Playfield, BM_Prim_LeftFlipper, BM_Prim_RightFlipper, BM_heart)
' Global arrays
Dim BG_Bakemap: BG_Bakemap=Array(BM_BottomTeeth, BM_Layer0, BM_Layer1, BM_Layer2, BM_Layer3, BM_Parts, BM_Playfield, BM_Prim_LeftFlipper, BM_Prim_RightFlipper, BM_heart)
Dim BG_Lightmap: BG_Lightmap=Array(LM_FL_Flasherflash1_Layer0, LM_FL_Flasherflash1_Layer1, LM_FL_Flasherflash1_Layer2, LM_FL_Flasherflash1_Layer3, LM_FL_Flasherflash1_Parts, LM_FL_Flasherflash1_Playfield, LM_FL_Flasherflash1_heart, LM_FL_Flasherflash2_BottomTeeth, LM_FL_Flasherflash2_Layer0, LM_FL_Flasherflash2_Layer1, LM_FL_Flasherflash2_Layer2, LM_FL_Flasherflash2_Layer3, LM_FL_Flasherflash2_Parts, LM_FL_Flasherflash2_Playfield, LM_FL_Flasherflash2_heart, LM_FL_Flasherflash3_Layer0, LM_FL_Flasherflash3_Layer1, LM_FL_Flasherflash3_Layer2, LM_FL_Flasherflash3_Parts, LM_FL_Flasherflash3_Playfield, LM_FL_Flasherflash3_heart, LM_FL_Flasherflash4_BottomTeeth, LM_FL_Flasherflash4_Layer0, LM_FL_Flasherflash4_Layer1, LM_FL_Flasherflash4_Layer2, LM_FL_Flasherflash4_Layer3, LM_FL_Flasherflash4_Parts, LM_FL_Flasherflash4_Playfield, LM_FL_Flasherflash4_heart, LM_FL_Flasherflash5_BottomTeeth, LM_FL_Flasherflash5_Layer0, LM_FL_Flasherflash5_Layer1, LM_FL_Flasherflash5_Layer2, LM_FL_Flasherflash5_Layer3, _
  LM_FL_Flasherflash5_Parts, LM_FL_Flasherflash5_Playfield, LM_FL_Flasherflash5_heart, LM_FL_f25_Layer0, LM_FL_f25_Layer1, LM_FL_f25_Layer2, LM_FL_f25_Parts, LM_FL_f25_heart, LM_GI_BottomTeeth, LM_GI_Layer0, LM_GI_Layer1, LM_GI_Layer2, LM_GI_Layer3, LM_GI_Parts, LM_GI_Playfield, LM_GI_Prim_LeftFlipper, LM_GI_Prim_RightFlipper, LM_GI_heart, LM_GIS_gi_001_Layer0, LM_GIS_gi_001_Parts, LM_GIS_gi_001_Playfield, LM_GIS_gi_001_Prim_LeftFlipper, LM_GIS_gi_001_Prim_RightFlipper, LM_GIS_gi_002_Layer0, LM_GIS_gi_002_Parts, LM_GIS_gi_002_Playfield, LM_GIS_gi_002_Prim_LeftFlipper, LM_GIS_gi_002_Prim_RightFlipper, LM_GIS_gi_003_Layer0, LM_GIS_gi_003_Parts, LM_GIS_gi_003_Playfield, LM_GIS_gi_003_Prim_LeftFlipper, LM_GIS_gi_003_Prim_RightFlipper, LM_GIS_gi_004_Layer0, LM_GIS_gi_004_Parts, LM_GIS_gi_004_Playfield, LM_GIS_gi_004_Prim_LeftFlipper, LM_GIS_gi_004_Prim_RightFlipper, LM_GIS_gi_005_Layer0, LM_GIS_gi_005_Parts, LM_GIS_gi_005_Playfield, LM_GIS_gi_005_Prim_LeftFlipper, LM_GIS_gi_005_Prim_RightFlipper, _
  LM_GIS_gi_006_Layer0, LM_GIS_gi_006_Parts, LM_GIS_gi_006_Playfield, LM_GIS_gi_006_Prim_LeftFlipper, LM_GIS_gi_006_Prim_RightFlipper, LM_GIS_gi_007_Layer0, LM_GIS_gi_007_Parts, LM_GIS_gi_007_Playfield, LM_GIS_gi_007_Prim_LeftFlipper, LM_GIS_gi_007_Prim_RightFlipper, LM_GIS_gi_008_Layer0, LM_GIS_gi_008_Parts, LM_GIS_gi_008_Playfield, LM_GIS_gi_008_Prim_LeftFlipper, LM_GIS_gi_008_Prim_RightFlipper, LM_IN_L36_BottomTeeth, LM_IN_L36_Layer1, LM_IN_L36_Layer3, LM_IN_L36_Parts, LM_IN_L37_Layer1, LM_IN_L37_Parts, LM_IN_L46_Layer0, LM_IN_L46_Layer1, LM_IN_L46_Parts, LM_IN_L47_Layer0, LM_IN_L47_Layer1, LM_IN_L47_Parts, LM_IN_L67_BottomTeeth, LM_IN_L67_Layer0, LM_IN_L67_Layer1, LM_IN_L67_Layer2, LM_IN_L67_Layer3, LM_IN_L67_Parts, LM_IN_L67_Playfield, LM_IN_L67_heart, LM_IN_Li10_Layer0, LM_IN_Li10_Layer1, LM_IN_Li10_Layer3, LM_IN_Li10_Parts, LM_IN_Li11_Layer0, LM_IN_Li11_Layer1, LM_IN_Li11_Layer3, LM_IN_Li11_Parts, LM_IN_Li11_Playfield, LM_IN_Li12_Layer0, LM_IN_Li12_Layer1, LM_IN_Li12_Layer2, LM_IN_Li12_Layer3, _
  LM_IN_Li12_Parts, LM_IN_Li12_Playfield, LM_IN_Li13_Layer0, LM_IN_Li13_Layer1, LM_IN_Li13_Layer3, LM_IN_Li13_Parts, LM_IN_Li13_Playfield, LM_IN_Li14_Layer0, LM_IN_Li14_Layer1, LM_IN_Li14_Layer3, LM_IN_Li14_Parts, LM_IN_Li2_BottomTeeth, LM_IN_Li2_Layer0, LM_IN_Li2_Layer1, LM_IN_Li2_Layer2, LM_IN_Li2_Layer3, LM_IN_Li2_Parts, LM_IN_Li2_Playfield, LM_IN_Li2_heart, LM_IN_bumperbiglight1_Layer0, LM_IN_bumperbiglight1_Layer1, LM_IN_bumperbiglight1_Layer2, LM_IN_bumperbiglight1_Parts, LM_IN_bumperbiglight1_Playfield, LM_IN_bumperbiglight2_Layer0, LM_IN_bumperbiglight2_Layer1, LM_IN_bumperbiglight2_Layer2, LM_IN_bumperbiglight2_Parts, LM_IN_bumperbiglight2_Playfield, LM_IN_bumperbiglight3_Layer1, LM_IN_bumperbiglight3_Parts, LM_IN_bumperbiglight3_Playfield, LM_IN_l3_Parts, LM_IN_l3_Playfield, LM_IN_l4_Parts)
Dim BG_All: BG_All=Array(BM_BottomTeeth, BM_Layer0, BM_Layer1, BM_Layer2, BM_Layer3, BM_Parts, BM_Playfield, BM_Prim_LeftFlipper, BM_Prim_RightFlipper, BM_heart, LM_FL_Flasherflash1_Layer0, LM_FL_Flasherflash1_Layer1, LM_FL_Flasherflash1_Layer2, LM_FL_Flasherflash1_Layer3, LM_FL_Flasherflash1_Parts, LM_FL_Flasherflash1_Playfield, LM_FL_Flasherflash1_heart, LM_FL_Flasherflash2_BottomTeeth, LM_FL_Flasherflash2_Layer0, LM_FL_Flasherflash2_Layer1, LM_FL_Flasherflash2_Layer2, LM_FL_Flasherflash2_Layer3, LM_FL_Flasherflash2_Parts, LM_FL_Flasherflash2_Playfield, LM_FL_Flasherflash2_heart, LM_FL_Flasherflash3_Layer0, LM_FL_Flasherflash3_Layer1, LM_FL_Flasherflash3_Layer2, LM_FL_Flasherflash3_Parts, LM_FL_Flasherflash3_Playfield, LM_FL_Flasherflash3_heart, LM_FL_Flasherflash4_BottomTeeth, LM_FL_Flasherflash4_Layer0, LM_FL_Flasherflash4_Layer1, LM_FL_Flasherflash4_Layer2, LM_FL_Flasherflash4_Layer3, LM_FL_Flasherflash4_Parts, LM_FL_Flasherflash4_Playfield, LM_FL_Flasherflash4_heart, LM_FL_Flasherflash5_BottomTeeth, _
  LM_FL_Flasherflash5_Layer0, LM_FL_Flasherflash5_Layer1, LM_FL_Flasherflash5_Layer2, LM_FL_Flasherflash5_Layer3, LM_FL_Flasherflash5_Parts, LM_FL_Flasherflash5_Playfield, LM_FL_Flasherflash5_heart, LM_FL_f25_Layer0, LM_FL_f25_Layer1, LM_FL_f25_Layer2, LM_FL_f25_Parts, LM_FL_f25_heart, LM_GI_BottomTeeth, LM_GI_Layer0, LM_GI_Layer1, LM_GI_Layer2, LM_GI_Layer3, LM_GI_Parts, LM_GI_Playfield, LM_GI_Prim_LeftFlipper, LM_GI_Prim_RightFlipper, LM_GI_heart, LM_GIS_gi_001_Layer0, LM_GIS_gi_001_Parts, LM_GIS_gi_001_Playfield, LM_GIS_gi_001_Prim_LeftFlipper, LM_GIS_gi_001_Prim_RightFlipper, LM_GIS_gi_002_Layer0, LM_GIS_gi_002_Parts, LM_GIS_gi_002_Playfield, LM_GIS_gi_002_Prim_LeftFlipper, LM_GIS_gi_002_Prim_RightFlipper, LM_GIS_gi_003_Layer0, LM_GIS_gi_003_Parts, LM_GIS_gi_003_Playfield, LM_GIS_gi_003_Prim_LeftFlipper, LM_GIS_gi_003_Prim_RightFlipper, LM_GIS_gi_004_Layer0, LM_GIS_gi_004_Parts, LM_GIS_gi_004_Playfield, LM_GIS_gi_004_Prim_LeftFlipper, LM_GIS_gi_004_Prim_RightFlipper, LM_GIS_gi_005_Layer0, _
  LM_GIS_gi_005_Parts, LM_GIS_gi_005_Playfield, LM_GIS_gi_005_Prim_LeftFlipper, LM_GIS_gi_005_Prim_RightFlipper, LM_GIS_gi_006_Layer0, LM_GIS_gi_006_Parts, LM_GIS_gi_006_Playfield, LM_GIS_gi_006_Prim_LeftFlipper, LM_GIS_gi_006_Prim_RightFlipper, LM_GIS_gi_007_Layer0, LM_GIS_gi_007_Parts, LM_GIS_gi_007_Playfield, LM_GIS_gi_007_Prim_LeftFlipper, LM_GIS_gi_007_Prim_RightFlipper, LM_GIS_gi_008_Layer0, LM_GIS_gi_008_Parts, LM_GIS_gi_008_Playfield, LM_GIS_gi_008_Prim_LeftFlipper, LM_GIS_gi_008_Prim_RightFlipper, LM_IN_L36_BottomTeeth, LM_IN_L36_Layer1, LM_IN_L36_Layer3, LM_IN_L36_Parts, LM_IN_L37_Layer1, LM_IN_L37_Parts, LM_IN_L46_Layer0, LM_IN_L46_Layer1, LM_IN_L46_Parts, LM_IN_L47_Layer0, LM_IN_L47_Layer1, LM_IN_L47_Parts, LM_IN_L67_BottomTeeth, LM_IN_L67_Layer0, LM_IN_L67_Layer1, LM_IN_L67_Layer2, LM_IN_L67_Layer3, LM_IN_L67_Parts, LM_IN_L67_Playfield, LM_IN_L67_heart, LM_IN_Li10_Layer0, LM_IN_Li10_Layer1, LM_IN_Li10_Layer3, LM_IN_Li10_Parts, LM_IN_Li11_Layer0, LM_IN_Li11_Layer1, LM_IN_Li11_Layer3, _
  LM_IN_Li11_Parts, LM_IN_Li11_Playfield, LM_IN_Li12_Layer0, LM_IN_Li12_Layer1, LM_IN_Li12_Layer2, LM_IN_Li12_Layer3, LM_IN_Li12_Parts, LM_IN_Li12_Playfield, LM_IN_Li13_Layer0, LM_IN_Li13_Layer1, LM_IN_Li13_Layer3, LM_IN_Li13_Parts, LM_IN_Li13_Playfield, LM_IN_Li14_Layer0, LM_IN_Li14_Layer1, LM_IN_Li14_Layer3, LM_IN_Li14_Parts, LM_IN_Li2_BottomTeeth, LM_IN_Li2_Layer0, LM_IN_Li2_Layer1, LM_IN_Li2_Layer2, LM_IN_Li2_Layer3, LM_IN_Li2_Parts, LM_IN_Li2_Playfield, LM_IN_Li2_heart, LM_IN_bumperbiglight1_Layer0, LM_IN_bumperbiglight1_Layer1, LM_IN_bumperbiglight1_Layer2, LM_IN_bumperbiglight1_Parts, LM_IN_bumperbiglight1_Playfield, LM_IN_bumperbiglight2_Layer0, LM_IN_bumperbiglight2_Layer1, LM_IN_bumperbiglight2_Layer2, LM_IN_bumperbiglight2_Parts, LM_IN_bumperbiglight2_Playfield, LM_IN_bumperbiglight3_Layer1, LM_IN_bumperbiglight3_Parts, LM_IN_bumperbiglight3_Playfield, LM_IN_l3_Parts, LM_IN_l3_Playfield, LM_IN_l4_Parts)
' VLM  Arrays - End


'******************************************************
'  ZTIM: Timers
'******************************************************

Dim FrameTime, InitFrameTime
InitFrameTime = 0

FrameTimer.Interval = -1
Sub FrameTimer_Timer()
  FrameTime = gametime - InitFrameTime 'Calculate FrameTime as some animuations could use this
  InitFrameTime = gametime  'Count frametime

  'Add animation stuff here
  RollingUpdate         'update rolling sounds
  DoDTAnim            'handle drop target animations
  DoSTAnim            'handle stand up target animations
  UpdateBallBrightness      'GI for the ball
  BSUpdate            'update ambient ball shadows

    If VR_Room = 0 and cab_mode = 0 Then: DisplayTimer: End If
  If VR_Room=1 Then: VRDisplayTimer: End If

End Sub

CorTimer.Interval = 10
Sub CorTimer_Timer()
  Cor.Update            'update ball tracking
End Sub


'******************************************************
'  ZINI: Table Initialization and Exiting
'******************************************************

Dim xx

Sub Table1_Init
  vpmInit Me
    With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
    .SplashInfoLine = "Class of 1812 (Gottlieb 1991)"&chr(13)&"by UnclePaulie"
    .HandleKeyboard = 0
    .ShowTitle = 0
    .ShowDMDOnly = 1
    .ShowFrame = 0
    .HandleMechanics = 0
    .dip(0) = &H00  'Set DIP to USA
    .Hidden = 1
    .Games(cGameName).Settings.Value("sound") = 1 ' Set sound (0=OFF, 1=ON)
    If Err Then MsgBox Err.Description
    End With
    On Error Goto 0
      Controller.Run
    If Err Then MsgBox Err.Description
    On Error Goto 0

  'Map all lamps to the corresponding ROM output using the value of TimerInterval of each light object
  vpmMapLights AllLamps     'Make a collection called "AllLamps" and put all the light objects in it.

    '** Nudging
  vpmNudge.TiltSwitch = 151
  vpmNudge.Sensitivity = 7
  vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingshot, LeftFlipper, RightFlipper)

  DiverterOn.IsDropped=1

'Ball initializations need for physical trough and ball shadows
  Set cBall1 = Slot1.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set cBall2 = Slot2.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Controller.Switch(44) = 1

'Ball initializations
  gBOT = Array(cBall1, cball2)

' Kicker initialization
  Controller.Switch(52) = 0
  Controller.Switch(53) = 0

    SetBackglass_Low
    SetBackglass_Mid
    SetBackglass_High

' GI on at Start
  For each xx in GI:xx.State = 1: Next
  Sound_GI_Relay 1, Relay_GI
' GI Bulb Prims
  For Each xx in BulbPrims: xx.blenddisablelighting = 2: Next

' Make drop target shadows visible
  for each xx in ShadowDTLeft
    xx.visible=True
  Next

  for each xx in ShadowDTRight
    xx.visible=True
  Next

' Drop Target Variable state

  DT16up=1
  DT17up=1
  DT26up=1
  DT27up=1
  DT36up=1
  DT37up=1
  DT46up=1
  DT47up=1

End Sub


Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub

Sub Table1_Paused: Controller.Pause = True: End Sub
Sub Table1_unPaused: Controller.Pause = False: End Sub


'**********************************
'   ZMAT: General Math Functions
'**********************************

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

''''' Flipper Animations

Sub LeftFlipper_Animate
  dim a: a = LeftFlipper.CurrentAngle
  FlipperLSh.RotZ = a
  'Add any left flipper related animations here
  Dim BP
  For each BP in BP_Prim_LeftFlipper:BP.rotz = a : Next
End Sub

Sub RightFlipper_Animate
  dim a: a = RightFlipper.CurrentAngle
  FlipperRSh.RotZ = a
  'Add any right flipper related animations here
  Dim BP
  For each BP in BP_Prim_RightFlipper:BP.rotz = a : Next
End Sub


''''' Gate Animations

Sub LGate_Animate
  LeftGate.RotY = - LGate.currentangle
End Sub

Sub RGate_Animate
  RightGate.RotY = - RGate.currentangle
End Sub


''''' Spinner Animations

Sub sw15_Animate
  sw15s.RotX = - sw15.currentangle
End Sub


'******************************************************
'   ZBBR: BALL BRIGHTNESS
'******************************************************

Const BallBrightness =  1       'Ball brightness - Value between 0 and 1 (0=Dark ... 1=Bright)

' Constants for plunger lane ball darkening.
' You can make a temporary wall in the plunger lane area and use the co-ordinates from the corner control points.
Const PLOffset = 0.25     'Minimum ball brightness scale in plunger lane
Const PLLeft = 865        'X position of punger lane left
Const PLRight = 934       'X position of punger lane right
Const PLTop = 375         'Y position of punger lane top
Const PLBottom = 1800       'Y position of punger lane bottom
Dim PLGain: PLGain = (1-PLOffset)/(PLTop-PLBottom)

Sub UpdateBallBrightness
  Dim s, b_base, b_r, b_g, b_b, d_w
' b_base = 70 * BallBrightness + 185*gilvl
  b_base = 120 * BallBrightness + 135*gilvl

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


'******************************************************
' ZKEY: Key Press Handling
'******************************************************


Sub Table1_KeyDown(ByVal keycode)

'**** added to fix the enter initials problem, so you can use the flippers
  If KeyCode=LeftFlipperKey Then Controller.Switch(4)=1
  If KeyCode=RightFlipperKey Then Controller.Switch(5)=1
'****

  If KeyCode = PlungerKey Then
    Plunger.PullBack
    Plunger.Pullback:SoundPlungerPull()
    TimerVRPlunger.Enabled = True
    TimerVRPlunger1.Enabled = False
    VR_Primary_plunger.Y = 2110.428
  End If

  If keycode = LeftTiltKey Then Nudge 90, 0.5 ': SoundNudgeLeft
  If keycode = RightTiltKey Then Nudge 270, 0.5 ': SoundNudgeRight
  If keycode = CenterTiltKey Then Nudge 0, 0.5 ': SoundNudgeCenter
  If keycode = StartGameKey then  PincabStartButton.y = PincabStartButton.y - 4 : l80.y = l80.y - 4

    If Keycode = LeftFlipperKey Then
    VR_FlipperButtonLeft.X = 2101.424 + 6
  End if
    If Keycode = RightFlipperKey Then
    VR_FlipperButtonRight.X = 2102.786 - 6
  End if

  If Frontkeys = 1 and keycode = RightMagnaSave Then
    Controller.Switch(5)=1
    VR_TopFiveRight.Y = VR_TopFiveRight.Y - 6: l80.y = l80.y - 6
  End If

  If Frontkeys = 1 and keycode = LeftMagnaSave Then
    Controller.Switch(4)=1
    VR_TopFiveLeft.Y = VR_TopFiveLeft.Y - 6: l80.y = l80.y - 6
  End If


  'If keycode = keyInsertCoin1 or keycode = keyInsertCoin2 or keycode = keyInsertCoin3 or keycode = keyInsertCoin4 Then 'Use this for ROM based games
  If keycode = AddCreditKey or keycode = AddCreditKey2 Then
    Select Case Int(rnd*3)
      Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
      Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
      Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
    End Select
  End If

  If keycode = StartGameKey Then
    SoundStartButton
  End If


    If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal keycode)
'**** added to fix the enter initials problem, so you can use the flippers
  If KeyCode=LeftFlipperKey Then Controller.Switch(4)=0
  If KeyCode=RightFlipperKey Then Controller.Switch(5)=0
'****
  If keycode = StartGameKey then PincabStartButton.y = PincabStartButton.y + 4 : l80.y = l80.y + 4

  If keycode = PlungerKey Then
    Plunger.Fire
    If BIPL = 1 Then
      SoundPlungerReleaseBall()     'Plunger release sound when there is a ball in shooter lane
    Else
      SoundPlungerReleaseNoBall()     'Plunger release sound when there is no ball in shooter lane
    End If
    TimerVRPlunger.Enabled = False
        TimerVRPlunger1.Enabled = True
    VR_Primary_plunger.Y = 2110.428
  End If

  If Keycode = LeftFlipperKey Then
    VR_FlipperButtonLeft.x = 2101.424
  End if
  If keycode = RightFlipperKey Then
    VR_FlipperButtonRight.X = 2102.786
  End If

  If Frontkeys = 1 and keycode = RightMagnaSave Then
    Controller.Switch(5)=0
    VR_TopFiveRight.Y = VR_TopFiveRight.Y + 6: l80.y = l80.y + 6
  End If

  If Frontkeys = 1 and keycode = LeftMagnaSave Then
    Controller.Switch(4)=0
    VR_TopFiveLeft.Y = VR_TopFiveLeft.Y + 6: l80.y = l80.y + 6
  End If


  if vpmKeyUp(keycode) Then Exit Sub
End Sub


'******************************************************
' VR Plunger code
'******************************************************

Sub TimerVRPlunger_Timer
  If VR_Primary_plunger.Y < 2220 then
       VR_Primary_plunger.Y = VR_Primary_plunger.Y + 5
  End If
End Sub

Sub TimerVRPlunger1_Timer
  VR_Primary_plunger.Y = 2110.428 + (5* Plunger.Position) -20
End Sub

'******************************************************
' ZSOL: Solenoids & Flashers
'******************************************************

SolCallback(6)  = "SolVUK"          '6 - Bottom Scoop
SolCallback(7)  = "SolSaucerUp"       '7 - Top Upkicker
SolCallback(8)  = "SolDtBankLeft"     '8 - Left Bank Reset
SolCallback(9)  = "SolDtBankRight"      '9 - Right Bank Reset
SolCallback(10) = "SolDiv"          '10 - Diverter
SolCallback(11) = "SolHeart"        '11 - Heart
SolCallback(12) = "SolTeeth"        '12 - Teeth
SolModCallback(22) = "DomeFlashWhite"           '22 - Flasher:Ramps #1 (2) White dome and yellow dome bottom
SolModCallback(23) = "DomeFlashRed"           '23 - Flasher:Ramps #2 (2) Red dome and yellow dome top
SolModCallback(24) = "DomeFlashBlue"          '24 - Flasher:Ramps #3 (2) Blue dome and under ramp

SolModCallback(25) = "SolHeartFlash"          '25 - Flasher:Heart (2)
SolModCallback(26) = "BGGI"         '26 - LightBox Relay
SolCallback(28) = "ReleaseBall"       '28 - Ball Release
SolCallback(29) = "SolOuthole"        '29 - Outhole.
SolCallback(30) = "SolKnocker"        '30 - Knocker
SolModCallback(31) = "PFGI"         '31 - Playfield GI / Tilt Relay

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"


'***********************************************
' *** Backglass Solenoid Controlled Flashers ***
'***********************************************

' THESE ARE OFF BY 1.  12--> 13; 13-->14, ETC.  (manual starts at zero)
SolModCallback(13) = "SolFLamp12"         '12 - Flasher:Lightning (3) **
SolModCallback(14) = "SolFLamp13"         '13 - Flasher:Willie Wolf (2) **
SolModCallback(15) = "SolFLamp14"         '14 - Flasher:Belly LaGhost (2) **
SolModCallback(16) = "SolFLamp15"         '15 - Flasher:Vamp do Tramp (2) **
SolModCallback(17) = "SolFLamp16"         '16 - Flasher:Mumsy N. Law (2) **
SolModCallback(18) = "SolFLamp17"         '17 - Flasher:Zom McCoffin (2) **
SolModCallback(19) = "SolFLamp18"         '18 - Flasher:Zom McCoffin (2) **
SolModCallback(20) = "SolFLamp19"         '19 - Flasher:Grover Cleaever (2) **
SolModCallback(21) = "SolFLamp20"         '20 - Flasher:Peter PieceMeal (2) **


'***********************************************
'******  DOME Solenoid Calls
'***********************************************

 Sub DomeFlashWhite(level)
  Flasherflash1.state = level
  Flasherflash5.state = level
  if level= 1 then
    Sound_Flash_Relay 1, sw45
    Sound_Flash_Relay 1, sw55
  elseif level = 0 Then
    Sound_Flash_Relay 0, sw45
    Sound_Flash_Relay 0, sw55
  end if
 End Sub

 Sub DomeFlashBlue(level)
  Flasherflash2.state = level
  if level= 1 then
    Sound_Flash_Relay 1, sw45
  elseif level = 0 Then
    Sound_Flash_Relay 0, sw45
  end if

 End Sub

 Sub DomeFlashRed(level)
  Flasherflash3.state = level
  Flasherflash4.state = level
  if level= 1 then
    Sound_Flash_Relay 1, sw45
    Sound_Flash_Relay 1, sw55
  elseif level = 0 Then
    Sound_Flash_Relay 0, sw45
    Sound_Flash_Relay 0, sw55
  end if
 End Sub


'******************************************************
' ZDRN: Drain, Trough, and Ball Release
'******************************************************

'********************* TROUGH *************************

Sub Slot2_Hit():Controller.Switch(44) = 1:UpdateTrough:End Sub
Sub Slot2_UnHit():Controller.Switch(44) = 0:UpdateTrough:End Sub
Sub Slot1_UnHit():UpdateTrough:End Sub

Sub UpdateTrough()
  Slot1.TimerInterval = 300
  Slot1.TimerEnabled = 1
End Sub

Sub Slot1_Timer()
  Me.TimerEnabled = 0
  If Slot1.BallCntOver = 0 Then Slot2.kick 60, 8
End Sub

'Drain & Release
Sub Drain_Hit()
  RandomSoundDrain Drain
  UpdateTrough
  Controller.Switch(54) = 1
End Sub

Sub Drain_UnHit()
  Controller.Switch(54) = 0
End Sub

Sub SolOuthole(enabled)
  If enabled Then
    Drain.kick 60,16
  SoundSaucerKick 1, Drain
  End If
End Sub

Sub ReleaseBall(enabled)
  If enabled Then
    RandomSoundBallRelease Slot1
    Slot1.kick 60, 10
    UpdateTrough
  End If
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

Dim LStep,RStep

Sub LeftSlingshot_Slingshot
  LS.VelocityCorrect(ActiveBall)
  RandomSoundSlingshotLeft SLING1
  vpmTimer.PulseSw 13
  LSling.Visible = 0
  LSling1.Visible = 1
  sling1.TransZ = -20
  LStep = 0
  Me.TimerInterval = 20
  Me.TimerEnabled = 1
End Sub

Sub LeftSlingshot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling1.TransZ = 0:Me.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingshot_Slingshot
  RS.VelocityCorrect(ActiveBall)
  RandomSoundSlingshotRight SLING2
  vpmTimer.PulseSw 14
  RSling.Visible = 0
  RSling1.Visible = 1
  sling2.TransZ = -20
  RStep = 0
  Me.TimerInterval = 20
  Me.TimerEnabled = 1
End Sub

Sub RightSlingshot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling2.TransZ = 0:Me.TimerEnabled = 0
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
  a = Array(LS, RS)
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
' Bumpers
'*******************************************

Sub Bumper1_Hit 'Left Top
  RandomSoundBumperTop Bumper1
  vpmTimer.PulseSw 10
End Sub

Sub Bumper2_Hit 'Right Top
  RandomSoundBumperTop Bumper2
  vpmTimer.PulseSw 11
End Sub

Sub Bumper3_Hit 'Bottom
  RandomSoundBumperBottom Bumper3
  vpmTimer.PulseSw 12
End Sub


'*******************************************
' Spinner
'*******************************************

'Spinner
Sub sw15_Spin
  vpmtimer.pulsesw 15
  SoundSpinner sw15
End Sub

'*******************************************
'Drop Targets
'*******************************************

Sub sw16_Hit
  DTHit 16
  DT16up=0
End Sub

Sub sw26_Hit
  DTHit 26
  DT26up=0
End Sub

Sub sw36_Hit
  DTHit 36
  DT36up=0
End Sub

Sub sw46_Hit
  DTHit 46
  DT46up=0
End Sub

Sub sw17_Hit
  DTHit 17
  DT17up=0
End Sub

Sub sw27_Hit
  DTHit 27
  DT27up=0
End Sub

Sub sw37_Hit
  DTHit 37
  DT37up=0
End Sub

Sub sw47_Hit
  DTHit 47
  DT47up=0
End Sub

' **** Reset Drop Targets by Solenoid callout

Sub SolDtBankLeft(enabled)
  dim xx
  if enabled then
    RandomSoundDropTargetReset sw16p
    DTRaise 16
    DTRaise 26
    DTRaise 36
    DTRaise 46
    DT16up=1
    DT26up=1
    DT36up=1
    DT46up=1
    if GIlvl=1 Then: for each xx in ShadowDTLeft: xx.visible=True: Next: End If
  end if
End Sub

Sub SolDtBankRight(enabled)
  dim xx
  if enabled then
    RandomSoundDropTargetReset sw17p
    DTRaise 17
    DTRaise 27
    DTRaise 37
    DTRaise 47
    DT17up=1
    DT27up=1
    DT37up=1
    DT47up=1
    if GIlvl=1 Then: for each xx in ShadowDTRight: xx.visible=True: Next: End If
  end if
End Sub


'*******************************************
'Standup Targets
'*******************************************

Sub sw20_Hit
  STHit 20
End Sub

Sub sw20o_Hit
  TargetBouncer ActiveBall, 1
End Sub

Sub sw21_Hit
  STHit 21
End Sub

Sub sw21o_Hit
  TargetBouncer ActiveBall, 1
End Sub

Sub sw22_Hit
  STHit 22
End Sub

Sub sw22o_Hit
  TargetBouncer ActiveBall, 1
End Sub

Sub sw30_Hit
  STHit 30
End Sub

Sub sw30o_Hit
  TargetBouncer ActiveBall, 1
End Sub

Sub sw31_Hit
  STHit 31
End Sub

Sub sw31o_Hit
  TargetBouncer ActiveBall, 1
End Sub

Sub sw32_Hit
  STHit 32
End Sub

Sub sw32o_Hit
  TargetBouncer ActiveBall, 1
End Sub

Sub sw40_Hit
  STHit 40
End Sub

Sub sw40o_Hit
  TargetBouncer ActiveBall, 1
End Sub

Sub sw41_Hit
  STHit 41
End Sub

Sub sw41o_Hit
  TargetBouncer ActiveBall, 1
End Sub

Sub sw42_Hit
  STHit 42
End Sub

Sub sw42o_Hit
  TargetBouncer ActiveBall, 1
End Sub


'*******************************************
'Rollovers
'*******************************************

Sub Sw23_Hit:Controller.Switch(23)=1: End Sub
Sub Sw23_UnHit:Controller.Switch(23)=0: End Sub
Sub Sw24_Hit:Controller.Switch(24)=1: End Sub
Sub Sw24_UnHit:Controller.Switch(24)=0: End Sub
Sub Sw25_Hit:Controller.Switch(25)=1: End Sub
Sub Sw25_UnHit:Controller.Switch(25)=0: End Sub

Sub Sw33_Hit:Controller.Switch(33)=1: End Sub
Sub Sw33_UnHit:Controller.Switch(33)=0: End Sub
Sub Sw34_Hit:Controller.Switch(34)=1: End Sub
Sub Sw34_UnHit:Controller.Switch(34)=0: End Sub
Sub Sw35_Hit:Controller.Switch(35)=1: End Sub
Sub Sw35_UnHit:Controller.Switch(35)=0: End Sub
Sub Sw43_Hit
  Controller.Switch(43)=1
  BIPL=True
End Sub
Sub Sw43_UnHit
  Controller.Switch(43)=0
  BIPL=False
End Sub
Sub Sw45_Hit:Controller.Switch(45)=1: End Sub
Sub Sw45_UnHit:Controller.Switch(45)=0: End Sub
Sub Sw55_Hit:Controller.Switch(55)=1: End Sub
Sub Sw55_UnHit:Controller.Switch(55)=0: End Sub
'Optos
Sub Sw50_Hit():vpmTimer.PulseSw 50: End Sub
Sub Sw51_Hit():vpmTimer.PulseSw 51: End Sub


'*******************************************
'  Ramp Triggers
'*******************************************
Sub rrenter_hit()
  WireRampOn True 'Play Plastic Ramp Sound
End Sub

Sub rrenter1_hit()
  WireRampOff ' Turn off the Plastic Ramp Sound
End Sub

Sub rrrampend_hit()
  If DiverterDir = -1 Then
    WireRampOff ' Turn off the Plastic Ramp Sound
  End If
End Sub

Sub RRail_hit()
  WireRampOff ' Turn off the Plastic Ramp Sound
End Sub

Sub RRail_unhit()
  WireRampOn False ' On Wire Ramp Play Wire Ramp Sound
End Sub

Sub RRExit_hit()
  WireRampOff ' Exiting Wire Ramp Stop Playing Sound
End Sub

Sub LRenter_hit()
  WireRampOn True 'Play Plastic Ramp Sound
End Sub

Sub LRenter1_hit()
  WireRampOff ' Turn off the Plastic Ramp Sound
End Sub

Sub LRail_hit()
  WireRampOff ' Turn off the Plastic Ramp Sound
End Sub

Sub LRail_unhit()
  WireRampOn False ' On Wire Ramp Play Wire Ramp Sound
End Sub

Sub LRExit_hit()
  WireRampOff ' Exiting Wire Ramp Stop Playing Sound
End Sub


'*******************************************
'  Lamp Solenoids
'*******************************************

Sub SolFLamp12(level)
  if VR_Room = 1 Then LFLamp12.state = level
End Sub

Sub SolFLamp13(level)
  if DesktopMode = True AND DTBGAnimation = 1 Then: L_DT_Willie.intensity = 6*level: End If
  if VR_Room = 1 Then LFLamp13.state = level
End Sub

Sub SolFLamp14(level)
  if DesktopMode = True AND DTBGAnimation = 1 Then: L_DT_Belly.intensity = 6*level: End If
  if VR_Room = 1 Then LFLamp14.state = level
End Sub

Sub SolFLamp15(level)
  if DesktopMode = True AND DTBGAnimation = 1 Then: L_DT_Vamp.intensity = 6*level: End If
  if VR_Room = 1 Then LFLamp15.state = level
End Sub

Sub SolFLamp16(level)
  if DesktopMode = True AND DTBGAnimation = 1 Then: L_DT_Mumsy.intensity = 6*level: End If
  if VR_Room = 1 Then LFLamp16.state = level
End Sub

Sub SolFLamp17(level)
  if VR_Room = 1 Then LFLamp17.state = level
End Sub

Sub SolFLamp18(level)
  if VR_Room = 1 Then LFLamp18.state = level
End Sub

Sub SolFLamp19(level)
  if VR_Room = 1 Then LFLamp19.state = level
End Sub

Sub SolFLamp20(level)
  if DesktopMode = True AND DTBGAnimation = 1 Then: L_DT_Peter.intensity = 6*level: End If
  if VR_Room = 1 Then LFLamp20.state = level
End Sub


'******************************************************
' Solenoids Routines
'******************************************************

'*********** Knocker
Sub SolKnocker(enabled)
  If Enabled Then
    KnockerSolenoid
  End If
End Sub


'*********** Scoop
Dim RNDKickValue1, RNDKickAngle1  'Random Values for saucer kick and angles

Sub sw52_Hit
  SoundSaucerLock
  Controller.Switch(52) = 1
End Sub

Sub SolVUK(Enabled)
  If Enabled Then
    sw52.timerinterval=500
    sw52.timerEnabled=1
    sw52p.transZ=30
  End If
End Sub

sub sw52_timer
  me.timerEnabled=0
  sw52p.transZ=10
  RNDKickAngle1 = RndInt(199,201)   ' Generate random value variance of +/- 1 from 200
  RNDKickValue1 = RndInt(24,26)   ' Generate random value variance of +/- 1 from 25
  sw52.kick RNDKickAngle1, RNDKickValue1,1
end Sub

Sub sw52_Unhit
  SoundSaucerKickCenter 1, sw52
  controller.switch(52) = 0
End Sub


'*********** Top Kicker
Sub sw53k_Hit
  SoundSaucerLock
  Controller.Switch(53) = 1
End Sub

Sub SolSaucerUp(Enabled)
  If Enabled Then
    sw53k.kick -30,25, 0
  End If
End Sub

Sub sw53k_Unhit
  SoundSaucerKick 1, sw53k
  controller.switch(53) = 0
End Sub


'*********** Diverter
Dim DiverterDir

Sub SolDiv(Enabled)
  If Enabled = -1 Then
    DiverterOff.IsDropped=1
    DiverterOn.IsDropped=0
    DiverterDir = -1
    Diverter.Interval = 5:Diverter.Enabled = 1
    PlaySoundAt SoundFX("DiverterOn",DOFContactors),DiverterP
  Else
    DiverterOff.IsDropped=0
    DiverterOn.IsDropped=1
    DiverterDir = +1
    Diverter.Interval = 5:Diverter.Enabled = 1
    PlaySoundAt SoundFX("DiverterOff",DOFContactors),DiverterP
    End If
End Sub

Sub Diverter_Timer()
DiverterP.RotZ=DiverterP.RotZ+DiverterDir
If DiverterP.RotZ>5 AND DiverterDir=1 Then Me.Enabled=0:DiverterP.RotZ=5
If DiverterP.RotZ<-35 AND DiverterDir=-1 Then Me.Enabled=0:DiverterP.RotZ=-35
End Sub


'********* Teeth
Sub SolTeeth(enabled)
  Dim BP
  If enabled = -1 Then
    For each BP in BP_BottomTeeth:BP.RotX=0: Next
  Else
    For each BP in BP_BottomTeeth:BP.RotX=-25: Next
  End If
End Sub


'********* Heart
dim heartbeat: heartbeat=0


Sub SolHeart(enabled)
  If enabled = -1 Then
  heartbeat = 1
       heartMove.enabled=1:PlaySoundAt SoundFX("SolenoidOn",DOFContactors),bm_heart
    Else
  heartbeat = 0
    PlaySoundAt SoundFX("SolenoidOff",DOFContactors),bm_heart
  End If
End Sub

dim pump, heartD
sub heartMove_timer()
  Dim bp
    pump=pump+10
    heartD=dsin(pump)
    if pump>180 then pump=0: for each BP in BP_Heart: BP.size_y=1: Next: Me.enabled=0
  for each BP in BP_Heart: BP.size_y=(heartD/8)+1.2: Next
end sub


Sub SolHeartFlash(level)
  F25.state = level
End Sub


'********************************************
'              Display Output
'********************************************

dim DisplayColor, DisplayColorG
DisplayColor =  RGB(0,255,30)

Sub VRDisplayTimer
  Dim ii, jj, obj, b, x
  Dim ChgLED,num, chg, stat
  ChgLED=Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED) Then
      For ii=0 To UBound(chgLED)
        num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)
        For Each obj In VRDigits(num)
          If chg And 1 Then FadeDisplay obj, stat And 1
          chg=chg\2 : stat=stat\2
        Next
      Next
    End If
End Sub

Sub FadeDisplay(object, onoffstat)
  If OnOffstat = 1 Then
    object.color = DisplayColor
  Else
    Object.Color = RGB(1,1,1)
  End If
End Sub

Dim VRDigits(40)

VRDigits(0)=Array(D1,D2,D3,D4,D5,D6,D7,D8,D9,D10,D11,D12,D13,D14,D15)
VRDigits(1)=Array(D16,D17,D18,D19,D20,D21,D22,D23,D24,D25,D26,D27,D28,D29,D30)
VRDigits(2)=Array(D31,D32,D33,D34,D35,D36,D37,D38,D39,D40,D41,D42,D43,D44,D45)
VRDigits(3)=Array(D46,D47,D48,D49,D50,D51,D52,D53,D54,D55,D56,D57,D58,D59,D60)
VRDigits(4)=Array(D61,D62,D63,D64,D65,D66,D67,D68,D69,D70,D71,D72,D73,D74,D75)
VRDigits(5)=Array(D76,D77,D78,D79,D80,D81,D82,D83,D84,D85,D86,D87,D88,D89,D90)
VRDigits(6)=Array(D91,D92,D93,D94,D95,D96,D97,D98,D99,D100,D101,D102,D103,D104,D105)
VRDigits(7)=Array(D106,D107,D108,D109,D110,D111,D112,D113,D114,D115,D116,D117,D118,D119,D120)
VRDigits(8)=Array(D121,D122,D123,D124,D125,D126,D127,D128,D129,D130,D131,D132,D133,D134,D135)
VRDigits(9)=Array(D136,D137,D138,D139,D140,D141,D142,D143,D144,D145,D146,D147,D148,D149,D150)
VRDigits(10)=Array(D151,D152,D153,D154,D155,D156,D157,D158,D159,D160,D161,D162,D163,D164,D165)
VRDigits(11)=Array(D166,D167,D168,D169,D170,D171,D172,D173,D174,D175,D176,D177,D178,D179,D180)
VRDigits(12)=Array(D181,D182,D183,D184,D185,D186,D187,D188,D189,D190,D191,D192,D193,D194,D195)
VRDigits(13)=Array(D196,D197,D198,D199,D200,D201,D202,D203,D204,D205,D206,D207,D208,D209,D210)
VRDigits(14)=Array(D211,D212,D213,D214,D215,D216,D217,D218,D219,D220,D221,D222,D223,D224,D225)
VRDigits(15)=Array(D226,D227,D228,D229,D230,D231,D232,D233,D234,D235,D236,D237,D238,D239,D240)

VRDigits(16)=Array(D241,D242,D243,D244,D245,D246,D247,D248,D249,D250,D251,D252,D253,D254,D255)
VRDigits(17)=Array(D256,D257,D258,D259,D260,D261,D262,D263,D264,D265,D266,D267,D268,D269,D270)
VRDigits(18)=Array(D271,D272,D273,D274,D275,D276,D277,D278,D279,D280,D281,D282,D283,D284,D285)
VRDigits(19)=Array(D286,D287,D288,D289,D290,D291,D292,D293,D294,D295,D296,D297,D298,D299,D300)
VRDigits(20)=Array(D301,D302,D303,D304,D305,D306,D307,D308,D309,D310,D311,D312,D313,D314,D315)
VRDigits(21)=Array(D316,D317,D318,D319,D320,D321,D322,D323,D324,D325,D326,D327,D328,D329,D330)
VRDigits(22)=Array(D331,D332,D333,D334,D335,D336,D337,D338,D339,D340,D341,D342,D343,D344,D345)
VRDigits(23)=Array(D346,D347,D348,D349,D350,D351,D352,D353,D354,D355,D356,D357,D358,D359,D360)
VRDigits(24)=Array(D361,D362,D363,D364,D365,D366,D367,D368,D369,D370,D371,D372,D373,D374,D375)
VRDigits(25)=Array(D376,D377,D378,D379,D380,D381,D382,D383,D384,D385,D386,D387,D388,D389,D390)
VRDigits(26)=Array(D391,D392,D393,D394,D395,D396,D397,D398,D399,D400,D401,D402,D403,D404,D405)
VRDigits(27)=Array(D406,D407,D408,D409,D410,D411,D412,D413,D414,D415,D416,D417,D418,D419,D420)
VRDigits(28)=Array(D421,D422,D423,D424,D425,D426,D427,D428,D429,D430,D431,D432,D433,D434,D435)
VRDigits(29)=Array(D436,D437,D438,D439,D440,D441,D442,D443,D444,D445,D446,D447,D448,D449,D450)
VRDigits(30)=Array(D451,D452,D453,D454,D455,D456,D457,D458,D459,D460,D461,D462,D463,D464,D465)
VRDigits(31)=Array(D466,D467,D468,D469,D470,D471,D472,D473,D474,D475,D476,D477,D478,D479,D480)

VRDigits(32)=Array(D481,D482,D483,D484,D485,D486,D487,D488,D489,D490,D491,D492,D493,D494,D495)
VRDigits(33)=Array(D496,D497,D498,D499,D500,D501,D502,D503,D504,D505,D506,D507,D508,D509,D510)
VRDigits(34)=Array(D511,D512,D513,D514,D515,D516,D517,D518,D519,D520,D521,D522,D523,D524,D525)
VRDigits(35)=Array(D526,D527,D528,D529,D530,D531,D532,D533,D534,D535,D536,D537,D538,D539,D540)
VRDigits(36)=Array(D541,D542,D543,D544,D545,D546,D547,D548,D549,D550,D551,D552,D553,D554,D555)
VRDigits(37)=Array(D556,D557,D558,D559,D560,D561,D562,D563,D564,D565,D566,D567,D568,D569,D570)
VRDigits(38)=Array(D571,D572,D573,D574,D575,D576,D577,D578,D579,D580,D581,D582,D583,D584,D585)
VRDigits(39)=Array(D586,D587,D588,D589,D590,D591,D592,D593,D594,D595,D596,D597,D598,D599,D600)

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

'**********************************************************************************************************
'Digital Display
'**********************************************************************************************************
Dim Digits(40)
Digits(0) = Array(a00, a05, a0c, a0d, a08, a01, a06, a0f, a02, a03, a04, a07, a0b, a0a, a09, a0e)
Digits(1) = Array(a10, a15, a1c, a1d, a18, a11, a16, a1f, a12, a13, a14, a17, a1b, a1a, a19, a1e)
Digits(2) = Array(a20, a25, a2c, a2d, a28, a21, a26, a2f, a22, a23, a24, a27, a2b, a2a, a29, a2e)
Digits(3) = Array(a30, a35, a3c, a3d, a38, a31, a36, a3f, a32, a33, a34, a37, a3b, a3a, a39, a3e)
Digits(4) = Array(a40, a45, a4c, a4d, a48, a41, a46, a4f, a42, a43, a44, a47, a4b, a4a, a49, a4e)
Digits(5) = Array(a50, a55, a5c, a5d, a58, a51, a56, a5f, a52, a53, a54, a57, a5b, a5a, a59, a5e)
Digits(6) = Array(a60, a65, a6c, a6d, a68, a61, a66, a6f, a62, a63, a64, a67, a6b, a6a, a69, a6e)
Digits(7) = Array(a70, a75, a7c, a7d, a78, a71, a76, a7f, a72, a73, a74, a77, a7b, a7a, a79, a7e)
Digits(8) = Array(a80, a85, a8c, a8d, a88, a81, a86, a8f, a82, a83, a84, a87, a8b, a8a, a89, a8e)
Digits(9) = Array(a90, a95, a9c, a9d, a98, a91, a96, a9f, a92, a93, a94, a97, a9b, a9a, a99, a9e)
Digits(10) = Array(aa0, aa5, aac, aad, aa8, aa1, aa6, aaf, aa2, aa3, aa4, aa7, aab, aaa, aa9, aae)
Digits(11) = Array(ab0, ab5, abc, abd, ab8, ab1, ab6, abf, ab2, ab3, ab4, ab7, abb, aba, ab9, abe)
Digits(12) = Array(ac0, ac5, acc, acd, ac8, ac1, ac6, acf, ac2, ac3, ac4, ac7, acb, aca, ac9, ace)
Digits(13) = Array(ad0, ad5, adc, add, ad8, ad1, ad6, adf, ad2, ad3, ad4, ad7, adb, ada, ad9, ade)
Digits(14) = Array(ae0, ae5, aec, aed, ae8, ae1, ae6, aef, ae2, ae3, ae4, ae7, aeb, aea, ae9, aee)
Digits(15) = Array(af0, af5, afc, afd, af8, af1, af6, aff, af2, af3, af4, af7, afb, afa, af9, afe)

Digits(16) = Array(b00, b05, b0c, b0d, b08, b01, b06, b0f, b02, b03, b04, b07, b0b, b0a, b09, b0e)
Digits(17) = Array(b10, b15, b1c, b1d, b18, b11, b16, b1f, b12, b13, b14, b17, b1b, b1a, b19, b1e)
Digits(18) = Array(b20, b25, b2c, b2d, b28, b21, b26, b2f, b22, b23, b24, b27, b2b, b2a, b29, b2e)
Digits(19) = Array(b30, b35, b3c, b3d, b38, b31, b36, b3f, b32, b33, b34, b37, b3b, b3a, b39, b3e)
Digits(20) = Array(b40, b45, b4c, b4d, b48, b41, b46, b4f, b42, b43, b44, b47, b4b, b4a, b49, b4e)
Digits(21) = Array(b50, b55, b5c, b5d, b58, b51, b56, b5f, b52, b53, b54, b57, b5b, b5a, b59, b5e)
Digits(22) = Array(b60, b65, b6c, b6d, b68, b61, b66, b6f, b62, b63, b64, b67, b6b, b6a, b69, b6e)
Digits(23) = Array(b70, b75, b7c, b7d, b78, b71, b76, b7f, b72, b73, b74, b77, b7b, b7a, b79, b7e)
Digits(24) = Array(b80, b85, b8c, b8d, b88, b81, b86, b8f, b82, b83, b84, b87, b8b, b8a, b89, b8e)
Digits(25) = Array(b90, b95, b9c, b9d, b98, b91, b96, b9f, b92, b93, b94, b97, b9b, b9a, b99, b9e)
Digits(26) = Array(ba0, ba5, bac, bad, ba8, ba1, ba6, baf, ba2, ba3, ba4, ba7, bab, baa, ba9, bae)
Digits(27) = Array(bb0, bb5, bbc, bbd, bb8, bb1, bb6, bbf, bb2, bb3, bb4, bb7, bbb, bba, bb9, bbe)
Digits(28) = Array(bc0, bc5, bcc, bcd, bc8, bc1, bc6, bcf, bc2, bc3, bc4, bc7, bcb, bca, bc9, bce)
Digits(29) = Array(bd0, bd5, bdc, bdd, bd8, bd1, bd6, bdf, bd2, bd3, bd4, bd7, bdb, bda, bd9, bde)
Digits(30) = Array(be0, be5, bec, bed, be8, be1, be6, bef, be2, be3, be4, be7, beb, bea, be9, bee)
Digits(31) = Array(bf0, bf5, bfc, bfd, bf8, bf1, bf6, bff, bf2, bf3, bf4, bf7, bfb, bfa, bf9, bfe)

Digits(32) = Array(c00, c05, c0c, c0d, c08, c01, c06, c0f, c02, c03, c04, c07, c0b, c0a, c09, c0e)
Digits(33) = Array(c10, c15, c1c, c1d, c18, c11, c16, c1f, c12, c13, c14, c17, c1b, c1a, c19, c1e)
Digits(34) = Array(c20, c25, c2c, c2d, c28, c21, c26, c2f, c22, c23, c24, c27, c2b, c2a, c29, c2e)
Digits(35) = Array(c30, c35, c3c, c3d, c38, c31, c36, c3f, c32, c33, c34, c37, c3b, c3a, c39, c3e)
Digits(36) = Array(c40, c45, c4c, c4d, c48, c41, c46, c4f, c42, c43, c44, c47, c4b, c4a, c49, c4e)
Digits(37) = Array(c50, c55, c5c, c5d, c58, c51, c56, c5f, c52, c53, c54, c57, c5b, c5a, c59, c5e)
Digits(38) = Array(c60, c65, c6c, c6d, c68, c61, c66, c6f, c62, c63, c64, c67, c6b, c6a, c69, c6e)
Digits(39) = Array(c70, c75, c7c, c7d, c78, c71, c76, c7f, c72, c73, c74, c77, c7b, c7a, c79, c7e)

 Sub DisplayTimer
    Dim ChgLED, ii, num, chg, stat, obj
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
' ZBRL: BALL ROLLING AND DROP SOUNDS
'******************************************************

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
dim RampBalls(3,2)
'x,0 = ball x,1 = ID, 2 = Protection against ending early (minimum amount of updates)
'0,0 is boolean on/off, 0,1 unused for now
RampBalls(0,0) = False

' RampType
'     Setup: Set this array to the number Total number of balls that can be tracked at one time + 1.  5 ball multiball then set value to 6
'     Description: Array type indexed on BallId and a values used to deterimine what type of ramp the ball is on: False = Wire Ramp, True = Plastic Ramp
dim RampType(5)

Sub WireRampOn(input)  : Waddball ActiveBall, input : RampRollUpdate: End Sub
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
      StopSound("wireloop" & x)
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
          PlaySound("RampLoop" & x), -1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * VolumeDial, AudioPan(RampBalls(x,0)), 0, BallPitchV(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
          StopSound("wireloop" & x)
        Else
          StopSound("RampLoop" & x)
          PlaySound("wireloop" & x), -1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * VolumeDial, AudioPan(RampBalls(x,0)), 0, BallPitch(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
        End If
        RampBalls(x, 2) = RampBalls(x, 2) + 1
      Else
        StopSound("RampLoop" & x)
        StopSound("wireloop" & x)
      end if
      if RampBalls(x,0).Z < 30 and RampBalls(x, 2) > RampMinLoops then  'if ball is on the PF, remove  it
        StopSound("RampLoop" & x)
        StopSound("wireloop" & x)
        Wremoveball RampBalls(x,1)
      End If
    Else
      StopSound("RampLoop" & x)
      StopSound("wireloop" & x)
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

Sub RandomSoundRampStop(obj)
  Select Case Int(rnd*3)
    Case 0: PlaySoundAtVol "wireramp_stop1", obj, 0.02*volumedial:PlaySoundAtLevelActiveBall ("Rubber_Strong_1"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
    Case 1: PlaySoundAtVol "wireramp_stop2", obj, 0.02*volumedial:PlaySoundAtLevelActiveBall ("Rubber_Strong_2"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
    Case 2: PlaySoundAtVol "wireramp_stop3", obj, 0.02*volumedial:PlaySoundAtLevelActiveBall ("Rubber_1_Hard"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
  End Select
End Sub


'******************************************************
'**** END RAMP ROLLING SFX
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

' Thalamus, AudioFade patched
	If tmp > 0 Then
		AudioFade = CSng(tmp ^ 5) 'was 10
	Else
		AudioFade = CSng( - (( - tmp) ^ 5) ) 'was 10
	End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
  'debug.print tableobj.name
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

Sub SoundSaucerKickCenter(scenario, saucer)
  Select Case scenario
    Case 0: PlaySoundAtLevelStatic SoundFX("Saucer_Empty", DOFContactors), SaucerKickSoundLevel, saucer
    Case 1: PlaySoundAtLevelStatic SoundFX("CenterEject", DOFContactors), SaucerKickSoundLevel, saucer
  End Select
End Sub

'/////////////////////////////  BALL COLLISION SOUND  ////////////////////////////

Sub OnBallBallCollision(ball1, ball2, velocity)
  dim collAngle,bvelx,bvely,hitball

  If ball1.radius < 23 or ball2.radius < 23 then

    If ball1.radius < 23 Then
      collAngle = GetCollisionAngle(ball1.x,ball1.y,ball2.x,ball2.y)
      bvelx = ball2.velx
      bvely = ball2.vely
      set hitball = ball2
    else
      collAngle = GetCollisionAngle(ball2.x,ball2.y,ball1.x,ball1.y)
      bvelx = ball1.velx
      bvely = ball1.vely
      set hitball = ball1
    End If

    dim discAngle
    discAngle = NormAngle(spinAngle)

    Dim mball, mdisc, rdisc, idisc

    discSpinSpeed = discSpinSpeed + sqr(bVelX ^2 + bVelY ^2) * sin(collAngle - discAngle) * cDiscSpeedMult

' Rubber hit random sounds
    Dim finalspeed
    finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
    If finalspeed > 5 Then
      RandomSoundRubberStrong 1
    End If
    If finalspeed <= 5 Then
      RandomSoundRubberWeak()
    End If

  Else

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

  End If

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


'/////////////////////////////  Diverters  ////////////////////////////
Dim Solenoid_Diverter_Enabled_SoundLevel, Solenoid_Diverter_Disabled_SoundLevel

Solenoid_Diverter_Enabled_SoundLevel = 1
Solenoid_Diverter_Disabled_SoundLevel = 0.4

'///////////////////////  RAMP DIVERTER SOLENOID - DIVERT  //////////////////////
Sub SoundRampDiverterDivert(obj)
  PlaySoundAtLevelStatic ("Diverter_Divert"), Solenoid_Diverter_Enabled_SoundLevel, obj
End Sub

'///////////////////////  RAMP DIVERTER SOLENOID - BACK  ///////////////////////
Sub SoundRampDiverterBack(obj)
  PlaySoundAtLevelStatic ("Diverter_Back"), Solenoid_Diverter_Disabled_SoundLevel, obj
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


'******************************************************
' Flippers Polarity (Select appropriate sub based on era)
'******************************************************

Dim LF : Set LF = New FlipperPolarity
Dim RF : Set RF = New FlipperPolarity

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
Dim DT16, DT17, DT26, DT27, DT36, DT37, DT46, DT47

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


Set DT16 = (new DropTarget)(sw16, sw16a, sw16p, 16, 0, false)
Set DT17 = (new DropTarget)(sw17, sw17a, sw17p, 17, 0, false)
Set DT26 = (new DropTarget)(sw26, sw26a, sw26p, 26, 0, false)
Set DT27 = (new DropTarget)(sw27, sw27a, sw27p, 27, 0, false)
Set DT36 = (new DropTarget)(sw36, sw36a, sw36p, 36, 0, false)
Set DT37 = (new DropTarget)(sw37, sw37a, sw37p, 37, 0, false)
Set DT46 = (new DropTarget)(sw46, sw46a, sw46p, 46, 0, false)
Set DT47 = (new DropTarget)(sw47, sw47a, sw47p, 47, 0, false)


Dim DTArray
DTArray = Array(DT16, DT17, DT26, DT27, DT36, DT37, DT46, DT47)

'Configure the behavior of Drop Targets.
Const DTDropSpeed = 110 'in milliseconds
Const DTDropUpSpeed = 40 'in milliseconds
Const DTDropUnits = 48 'VP units primitive drops so top of at or below the playfield
Const DTDropUpUnits = 10 'VP units primitive raises above the up position on drops up
Const DTMaxBend = 0 'max degrees primitive rotates when hit
Const DTDropDelay = 20 'time in milliseconds before target drops (due to friction/impact of the ball)
Const DTRaiseDelay = 40 'time in milliseconds before target drops back to normal up position after the solenoid fires to raise the target
Const DTBrickVel = 30 'velocity at which the target will brick, set to '0' to disable brick

Const DTEnableBrick = 0 'Set to 0 to disable bricking, 1 to enable bricking
Const DTHitSound = "" 'Drop Target Hit sound
Const DTDropSound = "DropTarget_Down" 'Drop Target Drop sound
Const DTResetSound = "DropTarget_Up" 'Drop Target reset sound

Const DTMass = 0.2 'Mass of the Drop Target (between 0 and 1), higher values provide more resistance


'******************************************************
'  DROP TARGETS FUNCTIONS
'******************************************************

' Initial Drop Target Shadows - Avoids a light DT hit and shadows go off when not strong enough hit to drop the target.

Dim DTShadow(8)

DTShadowInit 1
DTShadowInit 2
DTShadowInit 3
DTShadowInit 4
DTShadowInit 5
DTShadowInit 6
DTShadowInit 7
DTShadowInit 8

' Initializes the drop targets for shadow logic below
Sub DTShadowInit(dtnbr)

  if dtnbr = 1 Then
    Set DTShadow(dtnbr) = Eval("dtsh" & 16)
  elseif dtnbr = 2 Then
    Set DTShadow(dtnbr) = Eval("dtsh" & 17)
  elseif dtnbr = 3 Then
    Set DTShadow(dtnbr) = Eval("dtsh" & 26)
  elseif dtnbr = 4 Then
    Set DTShadow(dtnbr) = Eval("dtsh" & 27)
  elseif dtnbr = 5 Then
    Set DTShadow(dtnbr) = Eval("dtsh" & 36)
  elseif dtnbr = 6 Then
    Set DTShadow(dtnbr) = Eval("dtsh" & 37)
  elseif dtnbr = 7 Then
    Set DTShadow(dtnbr) = Eval("dtsh" & 46)
  elseif dtnbr = 8 Then
    Set DTShadow(dtnbr) = Eval("dtsh" & 47)
  End If
End Sub


Sub DTHit(switch)
  Dim i, swmod

  i = DTArrayID(switch)
  If switch = 16 Then
    swmod = 1
  Elseif switch = 17 then
    swmod = 2
  Elseif switch = 26 then
    swmod = 3
  Elseif switch = 27 then
    swmod = 4
  Elseif switch = 36 then
    swmod = 5
  Elseif switch = 37 then
    swmod = 6
  Elseif switch = 46 then
    swmod = 7
  Elseif switch = 47 then
    swmod = 8
  End If

  PlayTargetSound
  DTArray(i).animate =  DTCheckBrick(Activeball,DTArray(i).prim)
  If DTArray(i).animate = 1 or DTArray(i).animate = 3 or DTArray(i).animate = 4 Then
    DTBallPhysics Activeball, DTArray(i).prim.rotz, DTMass

' Controls Drop Shadow for a direct hit only
    if swmod=1 or swmod=2 or swmod=3 or swmod=4 or swmod=5 or swmod=6 or swmod=7 or swmod=8 then
      DTShadow(swmod).visible = 0
    End If

  End If
  DoDTAnim
End Sub

Sub DTRaise(switch)
  Dim i
  i = DTArrayID(switch)

  DTArray(i).animate = -1
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
  For i = 0 to uBound(DTArray)
    If DTArray(i).sw = switch Then DTArrayID = i:Exit Function
  Next
End Function


sub DTBallPhysics(aBall, angle, mass)
  dim rangle,bangle,calc1, calc2, calc3
  rangle = (angle - 90) * 3.1416 / 180
  bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))

  calc1 = cor.BallVel(aball.id) * cos(bangle - rangle) * (aball.mass - mass) / (aball.mass + mass)
  calc2 = cor.BallVel(aball.id) * sin(bangle - rangle) * cos(rangle + 4*Atn(1)/2)
  calc3 = cor.BallVel(aball.id) * sin(bangle - rangle) * sin(rangle + 4*Atn(1)/2)

  aBall.velx = calc1 * cos(rangle) + calc2
  aBall.vely = calc1 * sin(rangle) + calc3
End Sub


'Check if target is hit on it's face or sides and whether a 'brick' occurred
Function DTCheckBrick(aBall, dtprim)
  dim bangle, bangleafter, rangle, rangle2, Xintersect, Yintersect, cdist, perpvel, perpvelafter, paravel, paravelafter
  rangle = (dtprim.rotz - 90) * 3.1416 / 180
  rangle2 = dtprim.rotz * 3.1416 / 180
  bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))
  bangleafter = Atn2(aBall.vely,aball.velx)

  Xintersect = (aBall.y - dtprim.y - tan(bangle) * aball.x + tan(rangle2) * dtprim.x) / (tan(rangle2) - tan(bangle))
  Yintersect = tan(rangle2) * Xintersect + (dtprim.y - tan(rangle2) * dtprim.x)

  cdist = Distance(dtprim.x, dtprim.y, Xintersect, Yintersect)

  perpvel = cor.BallVel(aball.id) * cos(bangle-rangle)
  paravel = cor.BallVel(aball.id) * sin(bangle-rangle)

  perpvelafter = BallSpeed(aBall) * cos(bangleafter - rangle)
  paravelafter = BallSpeed(aBall) * sin(bangleafter - rangle)

  If perpvel > 0 and  perpvelafter <= 0 Then
    If DTEnableBrick = 1 and  perpvel > DTBrickVel and DTBrickVel <> 0 and cdist < 8 Then
      DTCheckBrick = 3
    Else
      DTCheckBrick = 1
    End If
  ElseIf perpvel > 0 and ((paravel > 0 and paravelafter > 0) or (paravel < 0 and paravelafter < 0)) Then
    DTCheckBrick = 4
  Else
    DTCheckBrick = 0
  End If
End Function


Sub DoDTAnim()
  Dim i
  For i=0 to Ubound(DTArray)
    DTArray(i).animate = DTAnimate(DTArray(i).primary,DTArray(i).secondary,DTArray(i).prim,DTArray(i).sw,DTArray(i).animate)
  Next
End Sub

Function DTAnimate(primary, secondary, prim, switch, animate)
  dim transz, switchid
  Dim animtime, rangle

  switchid = switch

  Dim ind
  ind = DTArrayID(switchid)

  rangle = prim.rotz * PI / 180

  DTAnimate = animate

  if animate = 0  Then
    primary.uservalue = 0
    DTAnimate = 0
    Exit Function
  Elseif primary.uservalue = 0 then
    primary.uservalue = gametime
  end if

  animtime = gametime - primary.uservalue

  If (animate = 1 or animate = 4) and animtime < DTDropDelay Then
    primary.collidable = 0
  If animate = 1 then secondary.collidable = 1 else secondary.collidable= 0
    prim.rotx = DTMaxBend * cos(rangle)
    prim.roty = DTMaxBend * sin(rangle)
    DTAnimate = animate
    Exit Function
    elseif (animate = 1 or animate = 4) and animtime > DTDropDelay Then
    primary.collidable = 0
    If animate = 1 then secondary.collidable = 1 else secondary.collidable= 0
    prim.rotx = DTMaxBend * cos(rangle)
    prim.roty = DTMaxBend * sin(rangle)
    animate = 2
    SoundDropTargetDrop prim
  End If

  if animate = 2 Then

    transz = (animtime - DTDropDelay)/DTDropSpeed *  DTDropUnits * -1
    if prim.transz > -DTDropUnits  Then
      prim.transz = transz
    end if

    prim.rotx = DTMaxBend * cos(rangle)/2
    prim.roty = DTMaxBend * sin(rangle)/2

    if prim.transz <= -DTDropUnits Then
      prim.transz = -DTDropUnits
      secondary.collidable = 0
      DTArray(ind).isDropped = true 'Mark target as dropped
      if UsingROM then controller.Switch(Switchid) = 1
      primary.uservalue = 0
      DTAnimate = 0
      Exit Function
    Else
      DTAnimate = 2
      Exit Function
    end If
  End If

  If animate = 3 and animtime < DTDropDelay Then
    primary.collidable = 0
    secondary.collidable = 1
    prim.rotx = DTMaxBend * cos(rangle)
    prim.roty = DTMaxBend * sin(rangle)
  elseif animate = 3 and animtime > DTDropDelay Then
    primary.collidable = 1
    secondary.collidable = 0
    prim.rotx = 0
    prim.roty = 0
    primary.uservalue = 0
    DTAnimate = 0
    Exit Function
  End If

  if animate = -1 Then
    transz = (1 - (animtime)/DTDropUpSpeed) *  DTDropUnits * -1

    If prim.transz = -DTDropUnits Then
      Dim b

      For b = 0 to UBound(gBOT)
        If InRotRect(gBOT(b).x,gBOT(b).y,prim.x, prim.y, prim.rotz, -25,-10,25,-10,25,25,-25,25) and gBOT(b).z < prim.z+DTDropUnits+25 Then
          gBOT(b).velz = 20
        End If
      Next
    End If

    if prim.transz < 0 Then
      prim.transz = transz
    elseif transz > 0 then
      prim.transz = transz
    end if

    if prim.transz > DTDropUpUnits then
      DTAnimate = -2
      prim.transz = DTDropUpUnits
      prim.rotx = 0
      prim.roty = 0
      primary.uservalue = gametime
    end if
    primary.collidable = 0
    secondary.collidable = 1
    DTArray(ind).isDropped = false 'Mark target as not dropped
    if UsingROM then controller.Switch(Switchid) = 0
  End If

  if animate = -2 and animtime > DTRaiseDelay Then
    prim.transz = (animtime - DTRaiseDelay)/DTDropSpeed *  DTDropUnits * -1 + DTDropUpUnits
    if prim.transz < 0 then
      prim.transz = 0
      primary.uservalue = 0
      DTAnimate = 0

      primary.collidable = 1
      secondary.collidable = 0
    end If
  End If
End Function


'******************************************************
'****  END DROP TARGETS
'******************************************************


'******************************************************
'   STAND-UP TARGET INITIALIZATION
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

Dim ST20, ST21, ST22, ST30, ST31, ST32, ST40, ST41, ST42

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

Set ST20 = (new StandupTarget)(sw20, psw20,20, 0)
Set ST21 = (new StandupTarget)(sw21, psw21,21, 0)
Set ST22 = (new StandupTarget)(sw22, psw22,22, 0)
Set ST30 = (new StandupTarget)(sw30, psw30,30, 0)
Set ST31 = (new StandupTarget)(sw31, psw31,31, 0)
Set ST32 = (new StandupTarget)(sw32, psw32,32, 0)
Set ST40 = (new StandupTarget)(sw40, psw40,40, 0)
Set ST41 = (new StandupTarget)(sw41, psw41,41, 0)
Set ST42 = (new StandupTarget)(sw42, psw42,42, 0)

Dim STArray
STArray = Array(ST20, ST21, ST22, ST30, ST31, ST32, ST40, ST41, ST42)


'Configure the behavior of Stand-up Targets
Const STAnimStep =  1.5         'vpunits per animation step (control return to Start)
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
    If UsingROM Then
      vpmTimer.PulseSw switch
    Else
      STAction switch
    End If
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


Sub STAction(Switch)
  Select Case Switch
    Case 11
      Addscore 1000
      Flash1 True 'Demo of the flasher
      vpmTimer.AddTimer 150,"Flash1 False'"   'Disable the flash after short time, just like a ROM would do

    Case 12
      Addscore 1000
      Flash2 True 'Demo of the flasher
      vpmTimer.AddTimer 150,"Flash2 False'"   'Disable the flash after short time, just like a ROM would do

    Case 13
      Addscore 1000
      Flash3 True 'Demo of the flasher
      vpmTimer.AddTimer 150,"Flash3 False'"   'Disable the flash after short time, just like a ROM would do
  End Select
End Sub


'******************************************************
'   END STAND-UP TARGETS
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

Sub l3_animate: p3.BlendDisableLighting = 100 * (l3.GetInPlayIntensity / l3.Intensity): End Sub
Sub l4_animate: p4.BlendDisableLighting = 100 * (l4.GetInPlayIntensity / l4.Intensity): End Sub
Sub l5_animate: p5.BlendDisableLighting = 100 * (l5.GetInPlayIntensity / l5.Intensity): End Sub
Sub l6_animate: p6.BlendDisableLighting = 200 * (l6.GetInPlayIntensity / l6.Intensity): End Sub
Sub l7_animate: p7.BlendDisableLighting = 100 * (l7.GetInPlayIntensity / l7.Intensity): End Sub
Sub l15_animate: p15.BlendDisableLighting = 200 * (l15.GetInPlayIntensity / l15.Intensity): End Sub
Sub l16_animate: p16.BlendDisableLighting = 200 * (l16.GetInPlayIntensity / l16.Intensity): End Sub
Sub l17_animate: p17.BlendDisableLighting = 200 * (l17.GetInPlayIntensity / l17.Intensity): End Sub
Sub l20_animate: p20.BlendDisableLighting = 200 * (l20.GetInPlayIntensity / l20.Intensity): End Sub
Sub l21_animate: p21.BlendDisableLighting = 200 * (l21.GetInPlayIntensity / l21.Intensity): End Sub
Sub l22_animate: p22.BlendDisableLighting = 200 * (l22.GetInPlayIntensity / l22.Intensity): End Sub
Sub l23_animate: p23.BlendDisableLighting = 200 * (l23.GetInPlayIntensity / l23.Intensity): End Sub
Sub l25_animate: p25.BlendDisableLighting = 200 * (l25.GetInPlayIntensity / l25.Intensity): End Sub
Sub l26_animate: p26.BlendDisableLighting = 200 * (l26.GetInPlayIntensity / l26.Intensity): End Sub
Sub l27_animate: p27.BlendDisableLighting = 200 * (l27.GetInPlayIntensity / l27.Intensity): End Sub
Sub l30_animate: p30.BlendDisableLighting = 200 * (l30.GetInPlayIntensity / l30.Intensity): End Sub
Sub l31_animate: p31.BlendDisableLighting = 200 * (l31.GetInPlayIntensity / l31.Intensity): End Sub
Sub l32_animate: p32.BlendDisableLighting = 200 * (l32.GetInPlayIntensity / l32.Intensity): End Sub
Sub l33_animate: p33.BlendDisableLighting = 200 * (l33.GetInPlayIntensity / l33.Intensity): End Sub
Sub l34_animate: p34.BlendDisableLighting = 200 * (l34.GetInPlayIntensity / l34.Intensity): End Sub
Sub l35_animate: p35.BlendDisableLighting = 200 * (l35.GetInPlayIntensity / l35.Intensity): End Sub
Sub l40_animate: p40.BlendDisableLighting = 200 * (l40.GetInPlayIntensity / l40.Intensity): End Sub
Sub l41_animate: p41.BlendDisableLighting = 200 * (l41.GetInPlayIntensity / l41.Intensity): End Sub
Sub l42_animate: p42.BlendDisableLighting = 200 * (l42.GetInPlayIntensity / l42.Intensity): End Sub
Sub l43_animate: p43.BlendDisableLighting = 200 * (l43.GetInPlayIntensity / l43.Intensity): End Sub
Sub l44_animate: p44.BlendDisableLighting = 200 * (l44.GetInPlayIntensity / l44.Intensity): End Sub
Sub l45_animate: p45.BlendDisableLighting = 200 * (l45.GetInPlayIntensity / l45.Intensity): End Sub
Sub l50_animate: p50.BlendDisableLighting = 200 * (l50.GetInPlayIntensity / l50.Intensity): End Sub
Sub l51_animate: p51.BlendDisableLighting = 200 * (l51.GetInPlayIntensity / l51.Intensity): End Sub
Sub l52_animate: p52.BlendDisableLighting = 200 * (l52.GetInPlayIntensity / l52.Intensity): End Sub
Sub l53_animate: p53.BlendDisableLighting = 200 * (l53.GetInPlayIntensity / l53.Intensity): End Sub
Sub l54_animate: p54.BlendDisableLighting = 200 * (l54.GetInPlayIntensity / l54.Intensity): End Sub
Sub l55_animate: p55.BlendDisableLighting = 200 * (l55.GetInPlayIntensity / l55.Intensity): End Sub
Sub l56_animate: p56.BlendDisableLighting = 200 * (l56.GetInPlayIntensity / l56.Intensity): End Sub
Sub l57_animate: p57.BlendDisableLighting = 100 * (l57.GetInPlayIntensity / l57.Intensity): End Sub
Sub l60_animate: p60.BlendDisableLighting = 200 * (l60.GetInPlayIntensity / l60.Intensity): End Sub
Sub l61_animate: p61.BlendDisableLighting = 200 * (l61.GetInPlayIntensity / l61.Intensity): End Sub
Sub l62_animate: p62.BlendDisableLighting = 200 * (l62.GetInPlayIntensity / l62.Intensity): End Sub
Sub l63_animate: p63.BlendDisableLighting = 200 * (l63.GetInPlayIntensity / l63.Intensity): End Sub
Sub l64_animate: p64.BlendDisableLighting = 200 * (l64.GetInPlayIntensity / l64.Intensity): End Sub
Sub l65_animate: p65.BlendDisableLighting = 200 * (l65.GetInPlayIntensity / l65.Intensity): End Sub
Sub l66_animate: p66.BlendDisableLighting = 200 * (l66.GetInPlayIntensity / l66.Intensity): End Sub
Sub l70_animate: p70.BlendDisableLighting = 200 * (l70.GetInPlayIntensity / l70.Intensity): End Sub
Sub l71_animate: p71.BlendDisableLighting = 200 * (l71.GetInPlayIntensity / l71.Intensity): End Sub
Sub l72_animate: p72.BlendDisableLighting = 200 * (l72.GetInPlayIntensity / l72.Intensity): End Sub
Sub l73_animate: p73.BlendDisableLighting = 200 * (l73.GetInPlayIntensity / l73.Intensity): End Sub
Sub l74_animate: p74.BlendDisableLighting = 200 * (l74.GetInPlayIntensity / l74.Intensity): End Sub
Sub l75_animate: p75.BlendDisableLighting = 200 * (l75.GetInPlayIntensity / l75.Intensity): End Sub
Sub l76_animate: p76.BlendDisableLighting = 200 * (l76.GetInPlayIntensity / l76.Intensity): End Sub
Sub l77_animate: p77.BlendDisableLighting = 200 * (l77.GetInPlayIntensity / l77.Intensity): End Sub
Sub l85_animate: p85.BlendDisableLighting = 200 * (l85.GetInPlayIntensity / l85.Intensity): End Sub
Sub l86_animate: p86.BlendDisableLighting = 200 * (l86.GetInPlayIntensity / l86.Intensity): End Sub
Sub l90_animate: p90.BlendDisableLighting = 200 * (l90.GetInPlayIntensity / l90.Intensity): End Sub
Sub l91_animate: p91.BlendDisableLighting = 200 * (l91.GetInPlayIntensity / l91.Intensity): End Sub
Sub l92_animate: p92.BlendDisableLighting = 200 * (l92.GetInPlayIntensity / l92.Intensity): End Sub
Sub l93_animate: p93.BlendDisableLighting = 200 * (l93.GetInPlayIntensity / l93.Intensity): End Sub
Sub l94_animate: p94.BlendDisableLighting = 200 * (l94.GetInPlayIntensity / l94.Intensity): End Sub
Sub l95_animate: p95.BlendDisableLighting = 200 * (l95.GetInPlayIntensity / l95.Intensity): End Sub
Sub l96_animate: p96.BlendDisableLighting = 200 * (l96.GetInPlayIntensity / l96.Intensity): End Sub
Sub l97_animate: p97.BlendDisableLighting = 200 * (l97.GetInPlayIntensity / l97.Intensity): End Sub
Sub l100_animate: p100.BlendDisableLighting = 200 * (l100.GetInPlayIntensity / l100.Intensity): End Sub
Sub l101_animate: p101.BlendDisableLighting = 200 * (l101.GetInPlayIntensity / l101.Intensity): End Sub
Sub l102_animate: p102.BlendDisableLighting = 200 * (l102.GetInPlayIntensity / l102.Intensity): End Sub
Sub l103_animate: p103.BlendDisableLighting = 200 * (l103.GetInPlayIntensity / l103.Intensity): End Sub
Sub l104_animate: p104.BlendDisableLighting = 200 * (l104.GetInPlayIntensity / l104.Intensity): End Sub
Sub l105_animate: p105.BlendDisableLighting = 200 * (l105.GetInPlayIntensity / l105.Intensity): End Sub
Sub l106_animate: p106.BlendDisableLighting = 200 * (l106.GetInPlayIntensity / l106.Intensity): End Sub
Sub l107_animate: p107.BlendDisableLighting = 200 * (l107.GetInPlayIntensity / l107.Intensity): End Sub
Sub l110_animate: p110.BlendDisableLighting = 200 * (l110.GetInPlayIntensity / l110.Intensity): End Sub
Sub l111_animate: p111.BlendDisableLighting = 200 * (l111.GetInPlayIntensity / l111.Intensity): End Sub
Sub l112_animate: p112.BlendDisableLighting = 200 * (l112.GetInPlayIntensity / l112.Intensity): End Sub
Sub l113_animate: p113.BlendDisableLighting = 200 * (l113.GetInPlayIntensity / l113.Intensity): End Sub
Sub l114_animate: p114.BlendDisableLighting = 200 * (l114.GetInPlayIntensity / l114.Intensity): End Sub
Sub l115_animate: p115.BlendDisableLighting = 200 * (l115.GetInPlayIntensity / l115.Intensity): End Sub
Sub l116_animate: p116.BlendDisableLighting = 200 * (l116.GetInPlayIntensity / l116.Intensity): End Sub
Sub l117_animate: p117.BlendDisableLighting = 200 * (l117.GetInPlayIntensity / l117.Intensity): End Sub

' Purple Lower Light (want some bdl when off)
Sub l0_animate: p0.BlendDisableLighting = 225 * (l0.GetInPlayIntensity / l0.Intensity): p0off.BlendDisableLighting = 1.75: End Sub

' White Middle Lights (want them to have some bdl when off)
Sub l80_animate: p80.BlendDisableLighting = 200 * (l80.GetInPlayIntensity / l80.Intensity): p80off.BlendDisableLighting = 7: End Sub
Sub l81_animate: p81.BlendDisableLighting = 200 * (l81.GetInPlayIntensity / l81.Intensity): p81off.BlendDisableLighting = 7: End Sub
Sub l82_animate: p82.BlendDisableLighting = 200 * (l82.GetInPlayIntensity / l82.Intensity): p82off.BlendDisableLighting = 7: End Sub
Sub l83_animate: p83.BlendDisableLighting = 200 * (l83.GetInPlayIntensity / l83.Intensity): p83off.BlendDisableLighting = 7: End Sub
Sub l84_animate: p84.BlendDisableLighting = 200 * (l84.GetInPlayIntensity / l84.Intensity): p84off.BlendDisableLighting = 7: End Sub

' Bat Lights
Sub l10_animate: Dim BL: for each BL in BL_IN_Li10: BL.Opacity = 10 * (l10.GetInPlayIntensity / l10.Intensity): Next: End Sub
Sub l11_animate: Dim BL: for each BL in BL_IN_Li11: BL.Opacity = 10 * (l11.GetInPlayIntensity / l11.Intensity): Next: End Sub
Sub l12_animate: Dim BL: for each BL in BL_IN_Li12: BL.Opacity = 10 * (l12.GetInPlayIntensity / l12.Intensity): Next: End Sub
Sub l13_animate: Dim BL: for each BL in BL_IN_Li13: BL.Opacity = 10 * (l13.GetInPlayIntensity / l13.Intensity): Next: End Sub
Sub l14_animate: Dim BL: for each BL in BL_IN_Li14: BL.Opacity = 10 * (l14.GetInPlayIntensity / l14.Intensity): Next: End Sub

' Ramp Lights
Sub l36_animate: Dim BL: for each BL in BL_IN_L36: BL.Opacity = 30 * (l36.GetInPlayIntensity / l36.Intensity): Next: End Sub
Sub l37_animate: Dim BL: for each BL in BL_IN_L37: BL.Opacity = 30 * (l37.GetInPlayIntensity / l37.Intensity): Next: End Sub
Sub l46_animate: Dim BL: for each BL in BL_IN_L46: BL.Opacity = 30 * (l46.GetInPlayIntensity / l46.Intensity): Next: End Sub
Sub l47_animate: Dim BL: for each BL in BL_IN_L47: BL.Opacity = 30 * (l47.GetInPlayIntensity / l47.Intensity): Next: End Sub
Sub l67_animate: Dim BL: for each BL in BL_IN_L67: BL.Opacity = 125 * (l67.GetInPlayIntensity / l67.Intensity): Next: End Sub

' Teeth Flasher Lights
Sub l2_animate
  Dim BL: for each BL in BL_IN_Li2: BL.Opacity = 175 * (l2.GetInPlayIntensity / l2.Intensity): Next
  Dim BP: for each BP in BP_BottomTeeth: BP.blenddisablelighting = 2.5 * (l2.GetInPlayIntensity / l2.Intensity) + 2.5: Next
End Sub


'******************************************************
'*****   END 3D INSERTS
'******************************************************


'******************************************************
'****  ZGIU:  GI Control
'******************************************************

Sub PFGI(level)
dim xgi, x, girubbercolor

debug.print level

  For each xgi in GI:xgi.State = (1-level): Next

    If level = 1 Then

    Sound_GI_Relay 0, Relay_GI
    gilvl = 0

    ' Set Drop Target Shadows to off when GI is goes off.
      for each xx in ShadowDTLeft
        xx.visible=False
      Next

      for each xx in ShadowDTRight
        xx.visible=False
      Next
    Elseif level =0 Then
    Sound_GI_Relay 1, Relay_GI
  Else

    ' Set Drop Target Shadows to state when GI is back on.
      if DT16up=1 Then dtsh16.visible = 1
      if DT17up=1 Then dtsh17.visible = 1
      if DT26up=1 Then dtsh26.visible = 1
      if DT27up=1 Then dtsh27.visible = 1
      if DT36up=1 Then dtsh36.visible = 1
      if DT37up=1 Then dtsh37.visible = 1
      if DT46up=1 Then dtsh46.visible = 1
      if DT47up=1 Then dtsh47.visible = 1
    gilvl=1
    End If

' Standup Targets
  For each x in GITargets: x.blenddisablelighting = 0.15 * (1-Level) + .05: Next
  For each x in GITargetsTop: x.blenddisablelighting = 0.2 * (1-Level) + .005: Next

' Drop Targets
  For each x in GIDropTargetsRight: x.blenddisablelighting = 0.125 * (1-Level) + .025: Next
  For each x in GIDropTargetsLeft: x.blenddisablelighting = 0.25 * (1-Level) + .05: Next

  'spinner
  sw15s.blenddisablelighting = 0.06 * (1-Level) + .005

' GI Bulb Prims
  For Each x in BulbPrims: x.blenddisablelighting = 2*(1-level): Next

' rubbers
  girubbercolor = 200*(1-Level) + 55
  MaterialColor "Rubber White",RGB(girubbercolor,girubbercolor,girubbercolor)

' Metals
  For each x in GIMetals: x.blenddisablelighting = 0.025 * (1-Level) + 0.01: Next

End Sub


Sub BGGI(level)

  PinCab_Backglass.blenddisablelighting = 0.2 + 0.6*(1-Level)

  if DesktopMode = True AND DTBGAnimation = 1 Then

      L_DT_1812.intensity = 6*(1-Level)
      L_DT_Text.intensity = 3*(1-Level)

    If level = 0 Then
      Sound_GI_Relay 1, Relay_BGGI
    Elseif level = 1 Then
      Sound_GI_Relay 0, Relay_BGGI
    End If
  End If
End Sub


'******************************************************
'****  END GI Control
'******************************************************



'***************************************************************
' ZSHA: Ambient ball shadows
'***************************************************************

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
'     objBallShadow(s).visible = 0
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


'**********************************************
'*******  Set Up Backglass Flashers *******
'**********************************************
' this is for lining up the backglass flashers on top of a backglass image

Sub SetBackglass_Low()
  Dim obj

  For Each obj In Backglass_Low
    obj.x = obj.x
    obj.height = - obj.y + 350
    obj.y = -70 'adjusts the distance from the backglass towards the user
  Next
End Sub


Sub SetBackglass_Mid()
  Dim obj

  For Each obj In Backglass_Mid
    obj.x = obj.x
    obj.height = - obj.y + 320
    obj.y = -90 'adjusts the distance from the backglass towards the user
  Next
End Sub


Sub SetBackglass_High()
  Dim obj

  For Each obj In Backglass_High
    obj.x = obj.x
    obj.height = - obj.y + 290
    obj.y = -100 'adjusts the distance from the backglass towards the user
  Next
End Sub


'***************************************************************
'   ZVRR: VR Room / VR Cabinet
'***************************************************************

Dim VRThings

' Desktop Mode
if VR_Room = 0 and cab_mode = 0 Then
  for each VRThings in VRStuff:VRThings.visible = 0:Next
  for each VRThings in VRClock:VRThings.visible = 0:Next
  for each VRThings in VRBackglass:VRThings.visible = 0:Next
  For each VRThings in DT_BG_Lights: VRThings.Visible = 1: Next
  For each VRThings in DT_Lamps: VRThings.Visible = 1: Next
  For each VRThings in DT_Lamps: VRThings.state = 1: Next
  SideRails.visible = 1

' Cabinet Mode
Elseif VR_Room = 0 and cab_mode = 1 Then
  for each VRThings in VRStuff:VRThings.visible = 0:Next
  for each VRThings in VRClock:VRThings.visible = 0:Next
  for each VRThings in VRBackglass:VRThings.visible = 0:Next
  For each VRThings in DT_BG_Lights: VRThings.Visible = 0: Next
  For each VRThings in DT_Lamps: VRThings.Visible = 0: Next
  For each VRThings in DT_Lamps: VRThings.state = 0: Next
  SideRails.visible = 0

' VR Mode
Else
  for each VRThings in VRStuff:VRThings.visible = 1:Next
  for each VRThings in VRClock:VRThings.visible = WallClock:Next
  for each VRThings in VRBackglass:VRThings.visible = 1:Next
  For each VRThings in DT_BG_Lights:VRThings.Visible = 0:Next
  For each VRThings in DT_Lamps: VRThings.Visible = 0: Next
  For each VRThings in DT_Lamps: VRThings.state = 0: Next
  SideRails.visible = 1

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

  If topper = 1 Then
    Primary_topper.visible = 1
  Else
    Primary_topper.visible = 0
  End If

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

'*******************************************
' VR Clock
'*******************************************

Dim CurrentMinute ' for VR clock

Sub ClockTimer_Timer()

    'ClockHands Below *********************************************************
  Pminutes.RotAndTra2 = (Minute(Now())+(Second(Now())/100))*6
  Phours.RotAndTra2 = Hour(Now())*30+(Minute(Now())/2)
  Pseconds.RotAndTra2 = (Second(Now()))*6
  CurrentMinute=Minute(Now())

End Sub

' ****************************************************


'******************************************************************
'******** Revisions done on VR version 1.0 by UnclePaulie *********
'******************************************************************

' - Ball shadow subroutine off, turned back on.
' - Added all the VR Cab primatives
' - Backglass Image from Blacksad's B2S file... Note:  ONLY used the dark image.
' - Added the alphanumeric displays from Mousin' Around, and then increased to 40 digits.
' - Added apron walls to block seeing through
' - Moved the apron forward, and adjusted the height so you can see ball come out
' - Matched colors for flipper buttons to cabinet - made red (real ones black, but hard to see)
' - Added two different VR room environments
' - Rails on left and right removed
' - The back flashers are too bright (F22r1,1x,2,2x,3,4,F23r,F24r1,r2,F25x) Changed all to 100
' - The (primative 33) was really bright.  Set blend disable underneath to 1.  Also set regular blend to .3.
' - The light13 was high, turned intensity down to 7. (was 35)
' - Changed the intensity of the lights (L36,37,46,47) to 50.  Was 150 and 200.
' - Changed the intensity of the lights (L67,67a) to 5.  Was 10.
' - Changed tilt sensitivity to 1
' - Added GI to backglass
' - Changed the ball image and added jp_scratches for decal to enhance the ball rolling realism
' - Modified the shapes of lights 1,2,3,4,5,6,15,17.  These are near the slings and flippers.  The original author
'   was trying to have the light blocked by the posts, but it looked really bad in VR.
' - Fixed issue of entering initials when getting a high score.  It required service buttons, however I tied to the left/right flippers now

'******** Revisions done on VR version 1.0 by Sixtoe *********

' - Sixtoe made some lighting modifications to flashers.
'   - Got rid of large lights (F22a,23,24a) and replaced with Flashers 152,153,154 to enable a bloom effect
'   - Cleaned up layer 7 lights to not overlap so much.  The left and top yellow lightgs.  Looks much better in VR.
'   - Moved the right red sidewall flasher as it was in the middle of the table.

'******** Revisions done on VR version 1.01 by UnclePaulie *********
' NOTE:  THIS IS NO LONGER VALID.  REAL FIX IS IN 1.3.1
' - There is an issue with the tilt, that if you tilt, and a ball is in play, you might get two balls when it kicks the next ball out
'   The solution was to change the trough slot2_hit to slot2_unhit.  This was resident in Tom Tower and Ninuzzu's original version.

'******** Revisions done on VR version 1.1 *********
' - Update to the cabinet, backbox, grill to match the real table.  Also, a new plunger, and added buttons in the front to match.

'******** Revisions done on VR version 1.2 *********
' NOTE:  THIS IS NO LONGER VALID.  REAL FIX IS IN 1.3.1
' ***  Updated two errors in original table.  Sub Slot2 was originally Slot2_Hit().  That caused too many balls to be on table in tilt mode
' ***  once that was changed, there was an issue that caused during the heartbeat multiball, the 2nd ball didn't come into outlane
' ***  the issue was that the Slot1_Hit() updated the trough in the subroutine.  Needed to remove UpdateTrough

'******** Revisions done on VR version 1.3 *********
' Animated the flashers on the backglass

'******** Revisions done on VR version 1.3.1 *********
' Finally fixed the trough issue, and the tilt issue, and the multiball issue.  The issue was that slot1 and slot2 subroutines were reversed.

'******** Revisions done by UnclePaulie on Hybrid version 0.1 - 2.0 *********
' v0.1: Enabled desktop, VR, and cabinet modes
'   Added an optional VR Clock, animated the plunger for VR keyboard use, and animated the flippers
'     Updated the playfield for slings and triggers holes, and added plywood holes.
' v0.2: Added Fleep Sounds
'   Added nFozzy / Roth physics and flippers.  Updated the flipper physics per recommendations.
'     Physics (early 90's);  table physics, gravity constant.  Also sling and bumper thresholds changed to match table physics
'   Corrected the Left and Right Gate rotation to RotY direction.  Was incorrect in original version.
'   Aligned SW55.  Also put correct color for primative8.
' v0.3: Added ramp triggers for ramp rolling sounds
'   Changed wall15 height to 70 (was 50).  Ball could jump over.
' v0.4: Changed the drop targets to the solution done by Rothbauerw
'   Modified the sling primatives and sling walls to drop below playfield.  Looks better in VR.
'   Added VPW Dynamic Ball Shadows to the solution done Iakki, Apophis, and Wylte
'   Added a couple wallblocks, as there were a couple times the ball bounced above the bumpers into the open area.
'     Also made the plastic primative on right side collidable.  Not sure this is the problem, as the bumper force and scatter angle
'     were increased during my physics updates.  I adjusted them down a little, while keeping the wall blocks.
'   Changed the color of the ramp entrances to black (that's what it is on the real table.)
'   Fixed the teeth movement in VR (solenoid12 conflict)
'     Continued to fine tune physics on flippers.
' v0.5  Added Flupper Domes.
'   Removed all the old flashers for the domes, and the flasherblooms Sixtoe added in old VR version.
'   Adjusted dome lighting and bloom intensity
' v0.6  Changed halo height of lighting, as it seemed to "float" above the playfield.  Also clicked, "show mesh" so you can see them lit.
'   Added LUT options with save option
'   Turned off reflection of some primatieves (23, 33, 34, and sw15s)
'   Added script for cab_mode
'   Turned down intensity of reflection on ball
'   Adjusted cab POV
'   Changed plywood cutout wood images
'   Adjusted some of the GI intensities at the top, and the bumpers.
' v0.7  bord - added playfield mesh, removed fallthrough "centerhole" kicker, resized sw52 to give the ball a chance to settle before the grab
' v0.8  Updated ramprolling sounds by fluffhead35 and adding ballroll and ramproll amplication sounds
' v0.9  Added lighter ball images provided by hauntfreaks and g5k ball scratches, changeable in script
' v2.0  Release version.
'   Added taller ramp walls and ramp ceilings to prevent balls jumping onto the two ramps (corner conditions)
'   Added invisible glass to keep ball in play (corner condition)
' v2.1  Updated ballrolling logic per recent update by fluffhead35 and apophis
'   Fixed the SW15 spinner not registering.  There was an ' in front of the vpmtimerpulse command.  Thanks to Lumi for noticing the problem.
' v2.2  Hid the rails in cabinet mode
' v2.3  Sixtoe provided a modified primative 45 for the saucer.
' v2.4  Added redbone images, playfield, apron, bumper cap, ramps, ramp entries, optional blood hand
'   Added movieguru upscaled apron cards
'     Updated the playfield material
'   Changed layer names
'     Added blood hand option
'   Fixed issue with flippers staying on after game over, and during tilt.
'     Updated the ball images and options
'   Removed display and lamp timers and added to frametimer
'     Automated the vr/cab/desktop
'   Removed/replaced solenoid on and off calls to just the sounds.
'   Updated the saucer kick routines, and added another sub for center sacuer kick sound
'   Added gBOT array
'   Updated to latest Fleep sounds and supporting functions
'   Updated to latest dynamic shadows, shadow images, flipper corrections, targetbouncer, collidesubs, rampshadows and sounds
'   Updated ramp sounds and shadows and GI relays
'     Added sling corrections
'   Updated the trough slot kick power and location, and added plunger grooves
'   Updated drop and stand up target arrays to standalone
'     Added VR posters, topper, outlet and cord, and standardized the VR/cab/desktop code.
' v2.5  Changed to Lampz and 3D prims
'   Updated the sol routines and light/fading for teeth, heart, and ramp bulbs.
'   Fixed the solenoid calls for playfield and backglass GI.  They were tied together on the old and original tables.
'   Added relay for BGGI.
'   Fixed an issue in the VR Backglass lamps.  All solenoids are 1 number off (manual started at 0).
'   Added flasher relay sounds
'     Added GI lighting to prims, balls, metals, targets, etc.
'     Added desktop backglass lamps and interactivity
'   Added option for desktop backglass lamps
' v2.6  Added flupper bumpers, modified
'   Adjusted the bumper flash level and bumper scale for bumper brightness on and off.
'   Moved location of Door Prize sign, as you could see the ball through it.  Also adjusted with GI levels
'   Adjusted the flasher relay sounds down slightly
'   Adjusted the nudge strength and the tilt sensitivity
'   Changed how magna saves control the LUT changes.  Also, added text for VR.
'     Removed the bumper screws from Primitive46 in blender, and added new ones.
' v2.7  New GI bulb prims, GI lights
' v2.8  Added GI bake and shadows
'   Adjusted the color of the GI lights
'   Added the toys to the GI routines.
'     Added DT Shadows, and modded the Drop Target routines to only turn shadows on when the DT actually drops (avoids a light Hit)
'   Updated DT solenoid routines to accomodate DT Shadows for game start, and when GI is on.
'     The upper triangle inserts had lines in them.  Mistake in VPW example table inserts, needed to be YSize=90, and Z = 0
'   Made the plastics collidable so not to accidentally get ball stuck under.
'   Added GI routine to all the toys, as well as the plastics
'   Updated the heartbeat blenddisablelighting to be controlled via GI routine, or the heartbeat timer
' v2.9  Redbone updated the apron and right plastics image.
'   Updated the playfield, GI, and shadow images for a slight imperfection in 2x-5x holes
'   Changed all images to webp format
'   Updated the cab pov
' v3.0  Small update to height of bumpers
'   Fixed DT shadow images as you could see shadow over dt holes and on dt.
'   Slight adjustment to sling location
'   Needed differentr material on bolts
'   Released
'v3.0.1 The right ramp under the hand plastics had a corner condition where there was a very slight collision with the top of the ball, causing difficult or slow up ramp.
'   Changed Primitive37 to be not collidable and added a simple invisible physics wall in it's place with a cutout near the ramp entrance.
'   Also changed Primitive27 to not collidable and added a physics wall.
'v3.0.2 Had to change desktopmode = 1 to = True in script to get desktop backglass lamps to operate correctly.

' Tomate
'v 3.0.3  Tomate added VLM bakes and moved layers around
'v 3.0.4  Tomate improved VLM bakes and moved layers around

' UnclePaulie
'v3.0.5 Removed FlasherGI, AO shadow image, and GI Tops.
'   Updated script to new VPW format standards
'   Removed Lampz and lamp call of 140 and 141 for old GI control
'   Added options for various balls.  Changed the updateball brightness code.
'   Made a collection called "AllLamps" and put all the light objects in it.   and changed all lamp timers to value
'   Changed script to alllamps and insert animations
'   Changed GI lights to have raytraced, show reflections on balls, states off, fader: Incandescent, fader 20, 40
'   Removed old LUT images and selection.  Changed to F12 menu and new VPW LUT color adjustments.
'   Updated to newer VR room images
'   Changed flipper shadow to visible and new depth bias.
'   Turned sling prims and sling sensors visible back on.  Moved to visible stuff layer.
'   Removed bsshadow code on ramps.
'   Removed flupper domes, and replaced with a simple light to control the new vpm light arrays.
'   Had to rename the dome lights to flasherflash to match the VLM lightarrays that tomate used.
'   Updated the heart movement to include the full BP_heart array.
'   Deleted unused prims.
'   Updated to latest flipper triggers and code.
'   Adjusted the inserts brightness
'   Removed flupper bumpers.  Only kept bumpersmall light, but named biglight (Tomate used that in light arrays).
'   Tied bumper lights to GI.
'   Removed lots of unused images and materials.
'     Added the new notargetbounce physics to the bottom sling posts
'   Modified the blue inserts slightly lower intensity.
'   Updated the smaller bulb light animations, as well as teeth
'   Animated the bumper rings by simply making them visible on the bumpers.
'   Got desktop backglass animation to work with new light routines.
'   Updated other visible parts GI effects.
'   Turned off raytracing on gi_009.  It was causing ball shadows all over playfield.
'   Added a VR pincab side behind new walls.  Also added a backwallglassholder and backwallscrews to VR.
'   Made a few modifications to the bottom teeth lighting.
' Tomate
'v3.0.8 Updated to the 4K batch and correted textures and lighting on the teeth and heart
' UnclePaulie
'v3.0.9 Updated DisableStaticPreRendering functionality to be consistent with VPX 10.8.1 API
'   Tied the heart flasher to VLM lights.  Deleted f25x flasher.
'   Adjusted the sounds of the ramp ball rolling sounces to 0.5 (was 0.1).  Updated ramp roll script.
'   Adjusted the GI on/off levels of the stand up targets and the spinner
' Tomate
'v3.0.10 Updated to a new bake with the VR sideblades to be the default.
'    Broke up the metal coin door, legs, metal parts prims separately.
'    Created a lockbar and siderails for both desktop and VR modes.
'v3.0.11 New bake.  Removed naming of prim with the "+"'s in it.  Fixed Wall20 height.
'    Tomate provided an updated mesh for the siderails used in desktop and VR to close a gap at top.
' UnclePaulie
'v3.0.12 Adjusted the script to account for the new sideblades and rails.  Deleted old ones.
'    Adjusted VR floor angle and height slightly to aling with Tomate's legs.
'    Deleted Wall20.
'    Increased the heart flash, the teethflash, and the back ramp light higher.
'    Changed flasher relay off sound to elseif=0
'    Changed insert light fader to LED none, letting pwm level handle it.
'    Apophis provided some guidance on pwm flashers.
'    Added F25 hidden lamp, changed script to control those, updated F25 lightmaps to point to them.
'    Use SolModCallback for flashers and set fader to NONE.
'    Lowered the light intensity of the l80-84.
'    Changed solenoid flashers logic to handle pwm and accurate level calls
'    VR outlet and cords were missing materials
'    Removed pincabgrill reflections
'    Change VR elements to be static rendering.
'    Modded pincabblades to just a back wall for VR.
'    Changed default ball to 7, and adjusted GI off brightness level.
'v3.0.13 Updated the heartflash routine and changed the SolFLamp backglass flashers for VR logic and faders
'    Changed LM_IN_L67_BottomTeeth nestmap image to VLM.Nestmap4 to get to work in VR.
' Tomate
'v3.0.14 New 4k batch added
'    PF set as "hide parts behind" and "reflections enabled"
'        turn visible all the prims needed
'        set bumpers rings visible
'        set "hide parts behind" for Layer1
' UnclePaulie
'v3.0.15 Re-enabled some functions lost in the new 4K bake.  Heartflash on F25 increased back to 175.
'    Added blend disable lighting to the bulb prims
'v3.0.16 Light 47 (blue lock light on right ramp wasn't working.  Needed to change timer to 47.  Also lightblooms 20, 22, 97, 100 were incorrect.
'v3.0.17 Borgdog found a couple light errors.  L97, and L46 timers were not right.
'v3.0.18 Forget to put physics walls in under the plastics (old version I had plastics collidable)
'    Made sw52p visible
'    Changed material on siderails.
'    Tomate updated the back plastic text to be more readable.
' Retro27
'v3.0.19 Reworked the artwork and overhauled the VR Cabinet.
'    Added option for magna saves for front buttons.  Redid timer plunger operation.
' UnclePaulie
'v3.0.20 One of the plastics physics walls (Plastics_Physics_Wall) protruded slightly in the ramp. Modded slightly to remove.
'    Added a default LUT.  Caused some users headtracking issues.
'    Corrected VR Plunger timing in VR with new prims Retro added.
'    Turned reflections off on VR prims added by Retro, and added missed VR prim to VRStuff collection.
'    Several new Retro prims were on incorrect layers.
'v3.0.21 Minor tweak to Prim37PhysicsWall, from 153 to 157, and physicswallblockback to 157.  Also lowered Ramp1 to height of 44.
'    Also changed material on VR_CoinmReturn2... was using rubberwhite, which could cause issues.

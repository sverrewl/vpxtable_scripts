'Eight Ball Deluxe 1.0.1
'based on a script by 32assassin
'Version 1.0 by Bord
'Version 2.0 by Bord and UnclePaulie

' All Revisions notes at the end of the script.

Option Explicit
Randomize


'***********************************************************************
'* TABLE OPTIONS *******************************************************
'***********************************************************************

' ALL OPTIONS ARE IN THE F12 MENU, EXCEPT FOR THE BALL COLOR and SetDIPSwitches BELOW

const ChangeBall  = 1   '0 = Ball changes are done via the drop down table menu
              '1 = ball selection per the below BallLightness option (default)
const BallLightness = 5   '0 = dark, 1 = not as dark, 2 = bright, 3 = brightest, 4 = MrBallDark, 5 = DarthVito Ball1 (Default)
              '6 = DarthVito HDR1, 7 = DarthVito HDR Bright, 8 = DarthVito HDR2, 9 = DarthVito ballHDR3,  10 = DarthVito ballHDR4
              '11 = DarthVito ballHDRdark, 12 = Borg Ball, 13 = SteelBall2

const SetDIPSwitches= 0   'If you want to set the dips differently, set to 1, and then hit F6 to launch.


'***********************************************************************
' VR USERS... THE F12 MENU ISN'T AVAILABLE AND WILL NEED TO SELECT OPTIONS HERE


' User Options for In-Game Menu F12 then use Flipper and Magna Keys
'   NOTE: The detailed definitions of F12 Table Options are at the bottom of the script.

  'variables
  Dim VolumeDial, BallRollVolume, ROMVolume
  Dim WallClock, poster, poster2, CustomWalls

  Dim VR_Room, cab_mode, DesktopMode: DesktopMode = Table1.ShowDT
  If RenderingMode = 2 Then VR_Room=1 Else VR_Room=0
  if Not DesktopMode and VR_Room=0 Then cab_mode=1 Else cab_mode=0

' NOTE:  Options for desktop and cab are done via F12 menu.  VR users have to manually select here in the script.

  IF VR_Room = 1 Then

WallClock   = 1 '1 Shows the clock in the VR minimal rooms only
poster      = 1 '1 Shows the flyer posters in the VR room only
poster2     = 1 '1 Shows the flyer posters in the VR room only
CustomWalls   = 2 '0=UP's Original Minimal Walls, floor, and roof
              '1=Sixtoe's arcade style
              '2=DarthVito's Updated home walls with lights
              '3=DarthVito's plaster home walls
              '4=DarthVito's blue home walls
ROMVolume   = 0 'ROM Volume (-32 is off; 0 is max db) - Requires Table Restart

VolumeDial    = 0.8 'Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
BallRollVolume  = 0.5 'Level of ball rolling volume. Value between 0 and 1

  End If


'***********************************************************************


' Table1.Option arguments are:
' - option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional arry of literal strings

' Called when options are tweaked by the player.
' - 0: game has started, good time to load options and adjust accordingly
' - 1: an option has changed
' - 2: options have been reseted
' - 3: player closed the tweak UI, good time to update staticly prerendered parts


Dim LightLevel      ' Level of room lighting (0 to 1), where 0 is dark and 100 is brightest
Dim ColorLUT      ' Color desaturation LUTs: 1 to 11, where 1 is normal and 11 is black'n'white


Sub Table1_OptionEvent(ByVal eventId)

  if VR_Room = 0 Then ' only allow F12 menu for desktop and cab.... not currently working in VR

    If eventId = 1 Then DisableStaticPreRendering = True

  ROMVolume = Table1.Option("ROM Volume (-32 is off; 0 is max db) - Requires Table Restart", -32, 0, 1, 0, 0)

  SetupOptions

    ' Sound volumes
  VolumeDial = Table1.Option("Mech Volume", 0, 1, 0.01, 0.8, 1)
  BallRollVolume = Table1.Option("Ball Roll Volume", 0, 1, 0.01, 0.5, 1)

  ' Color Saturation
    ColorLUT = Table1.Option("Color Saturation", 1, 11, 1, 1, 0, _
    Array("Normal", "Desaturated 10%", "Desaturated 20%", "Desaturated 30%", "Desaturated 40%", "Desaturated 50%", _
        "Desaturated 60%", "Desaturated 70%", "Desaturated 80%", "Desaturated 90%", "Black 'n White"))
  if ColorLUT = 1 Then Table1.ColorGradeImage = ""
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
  LightLevel = Table1.Option("Table Brightness (Ambient Light Level)", 0, 1, 0.01, .65, 1)
  SetRoomBrightness LightLevel   'Uncomment this line for lightmapped tables.

    If eventId = 3 Then DisableStaticPreRendering = False

  End If

End Sub


On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const UseSolenoids=2,UseLamps=1,UseGI=0
Const cGameName="eballdlx"


'******************************************************
' Constants and Global Variables
'******************************************************

Const BallSize = 50
Const BallMass = 1
Const tnob = 1 ' total number of balls
Const lob = 0   'number of locked balls


Dim tablewidth: tablewidth = Table1.width
Dim tableheight: tableheight = Table1.height
Dim BIPL : BIPL = False       'Ball in plunger lane
Dim BP

LoadVPM "01560000", "Bally.vbs", 3.26


'******************************************************
' VLM Arrays
'******************************************************

' VLM  Arrays - Start
' Arrays per baked part
Dim BP_BumperRing1: BP_BumperRing1=Array(BM_BumperRing1, LM_GI_GI_21_BumperRing1, LM_GI_GI_18_BumperRing1, LM_Ins1_L117_BumperRing1)
Dim BP_BumperRing2: BP_BumperRing2=Array(BM_BumperRing2, LM_GI_GI_21_BumperRing2, LM_GI_GI_20_BumperRing2)
Dim BP_BumperRing3: BP_BumperRing3=Array(BM_BumperRing3, LM_GI_GI_4_BumperRing3, LM_GI_GI_20_BumperRing3, LM_Ins1_L101_BumperRing3)
Dim BP_BumperSkirt1: BP_BumperSkirt1=Array(BM_BumperSkirt1, LM_Ins1_L70_BumperSkirt1, LM_GI_GI_21_BumperSkirt1, LM_GI_GI_17_BumperSkirt1, LM_GI_GI_25_BumperSkirt1, LM_GI_GI_24_BumperSkirt1, LM_GI_GI_23_BumperSkirt1, LM_GI_GI_22_BumperSkirt1, LM_GI_GI_18_BumperSkirt1, LM_GI_GI_20_BumperSkirt1, LM_Ins1_L119_BumperSkirt1, LM_Ins1_L118_BumperSkirt1, LM_Ins1_L117_BumperSkirt1, LM_Ins1_L101_BumperSkirt1)
Dim BP_BumperSkirt2: BP_BumperSkirt2=Array(BM_BumperSkirt2, LM_GI_gi1_BumperSkirt2, LM_Ins1_L86_BumperSkirt2, LM_GI_GI_21_BumperSkirt2, LM_GI_GI_19_BumperSkirt2, LM_GI_GI_17_BumperSkirt2, LM_GI_GI_25_BumperSkirt2, LM_GI_GI_24_BumperSkirt2, LM_GI_GI_23_BumperSkirt2, LM_GI_GI_22_BumperSkirt2, LM_GI_GI_18_BumperSkirt2, LM_GI_GI_20_BumperSkirt2, LM_Ins1_L119a_BumperSkirt2, LM_Ins1_L117a_BumperSkirt2, LM_Ins1_L101_BumperSkirt2, LM_Ins1_L101a_BumperSkirt2)
Dim BP_BumperSkirt3: BP_BumperSkirt3=Array(BM_BumperSkirt3, LM_Ins1_L7_BumperSkirt3, LM_Ins1_L102_BumperSkirt3, LM_GI_GI_4_BumperSkirt3, LM_GI_GI_21_BumperSkirt3, LM_GI_GI_20_BumperSkirt3, LM_Ins1_L117_BumperSkirt3, LM_Ins1_L101_BumperSkirt3, LM_Ins1_L101a_BumperSkirt3)
Dim BP_Flip1: BP_Flip1=Array(BM_Flip1, LM_Ins1_L103_Flip1, LM_Ins1_L1_Flip1, LM_Ins1_L33_Flip1, LM_Ins1_L28_Flip1, LM_Ins1_L60_Flip1, LM_GI_gi1_Flip1, LM_GI_GI_9_Flip1, LM_GI_GI_14_Flip1, LM_GI_GI_10_Flip1, LM_GI_GI_15_Flip1, LM_GI_GI_13_Flip1, LM_GI_GI_4_Flip1, LM_GI_GI_8_Flip1, LM_GI_GI_11_Flip1, LM_Ins2_L44_Flip1, LM_Ins2_L87_Flip1, LM_Ins2_L11_Flip1)
Dim BP_Flip2: BP_Flip2=Array(BM_Flip2, LM_Ins1_L103_Flip2, LM_Ins1_L1_Flip2, LM_Ins1_L17_Flip2, LM_Ins1_L28_Flip2, LM_Ins1_L60_Flip2, LM_GI_gi1_Flip2, LM_GI_GI_9_Flip2, LM_GI_GI_14_Flip2, LM_GI_GI_10_Flip2, LM_GI_GI_13_Flip2, LM_GI_GI_8_Flip2, LM_GI_GI_11_Flip2, LM_Ins2_L12_Flip2, LM_Ins2_L71_Flip2, LM_Ins2_L11_Flip2)
Dim BP_Flip3: BP_Flip3=Array(BM_Flip3, LM_Ins1_L103_Flip3, LM_Ins1_L17_Flip3, LM_Ins1_L49_Flip3, LM_Ins1_L2_Flip3, LM_Ins1_L34_Flip3, LM_Ins1_L35_Flip3, LM_Ins1_L50_Flip3, LM_Ins1_L20_Flip3, LM_Ins1_L4_Flip3, LM_Ins1_L51_Flip3, LM_Ins1_L58_Flip3, LM_Ins1_L65_Flip3, LM_Ins1_L81_Flip3, LM_Ins1_L97_Flip3, LM_GI_gi1_Flip3, LM_Ins1_L114_Flip3, LM_Ins1_L63_Flip3, LM_GI_GI_9_Flip3, LM_GI_GI_14_Flip3, LM_GI_GI_13_Flip3, LM_GI_GI_4_Flip3, LM_GI_GI_8_Flip3, LM_GI_GI_11_Flip3, LM_Ins1_L101_Flip3, LM_Ins2_L41_Flip3)
Dim BP_Gateplate2: BP_Gateplate2=Array(BM_Gateplate2, LM_GI_GI_22_Gateplate2, LM_Ins1_L118_Gateplate2)
Dim BP_Gateplate5: BP_Gateplate5=Array(BM_Gateplate5, LM_GI_gi1_Gateplate5, LM_Ins1_L119_Gateplate5)
Dim BP_Lane_Guide___2a: BP_Lane_Guide___2a=Array(BM_Lane_Guide___2a, LM_GI_GI_24_Lane_Guide___2a, LM_GI_GI_23_Lane_Guide___2a, LM_GI_GI_22_Lane_Guide___2a)
Dim BP_Layer_2: BP_Layer_2=Array(BM_Layer_2, LM_Ins1_L26_Layer_2, LM_GI_gi1_Layer_2, LM_Ins1_L7_Layer_2, LM_Ins1_L115_Layer_2, LM_Ins1_L102_Layer_2, LM_Ins1_L70_Layer_2, LM_Ins1_L86_Layer_2, LM_GI_GI_9_Layer_2, LM_GI_GI_14_Layer_2, LM_GI_GI_10_Layer_2, LM_GI_GI_15_Layer_2, LM_GI_GI_13_Layer_2, LM_GI_GI_21_Layer_2, LM_GI_GI_19_Layer_2, LM_GI_GI_17_Layer_2, LM_GI_GI_25_Layer_2, LM_GI_GI_24_Layer_2, LM_GI_GI_23_Layer_2, LM_GI_GI_22_Layer_2, LM_GI_GI_1_Layer_2, LM_GI_GI_6_Layer_2, LM_GI_GI_18_Layer_2, LM_GI_GI_20_Layer_2, LM_Ins1_L119a_Layer_2, LM_Ins1_L119_Layer_2, LM_Ins1_L118a_Layer_2, LM_Ins1_L118_Layer_2, LM_Ins1_L117a_Layer_2, LM_Ins1_L117_Layer_2, LM_GI_GI_11_Layer_2, LM_Ins1_L101_Layer_2, LM_Ins1_L101a_Layer_2, LM_Ins2_L68_Layer_2)
Dim BP_Parts: BP_Parts=Array(BM_Parts, LM_Ins1_L17_Parts, LM_Ins1_L60_Parts, LM_Ins1_L10_Parts, LM_Ins1_L26_Parts, LM_Ins1_L42_Parts, LM_Ins1_L58_Parts, LM_Ins1_L81_Parts, LM_Ins1_L97_Parts, LM_Ins1_L82_Parts, LM_GI_gi1_Parts, LM_Ins1_L114_Parts, LM_Ins1_L5_Parts, LM_Ins1_L54_Parts, LM_Ins1_L115_Parts, LM_Ins1_L102_Parts, LM_Ins1_L70_Parts, LM_Ins1_L86_Parts, LM_Ins1_L59_Parts, LM_GI_GI_9_Parts, LM_GI_GI_14_Parts, LM_GI_GI_10_Parts, LM_GI_GI_15_Parts, LM_GI_GI_13_Parts, LM_GI_GI_4_Parts, LM_GI_GI_8_Parts, LM_GI_GI_21_Parts, LM_GI_GI_19_Parts, LM_GI_GI_17_Parts, LM_GI_GI_25_Parts, LM_GI_GI_24_Parts, LM_GI_GI_23_Parts, LM_GI_GI_22_Parts, LM_GI_GI_1_Parts, LM_GI_GI_6_Parts, LM_GI_GI_18_Parts, LM_GI_GI_20_Parts, LM_Ins1_L119a_Parts, LM_Ins1_L119_Parts, LM_Ins1_L118a_Parts, LM_Ins1_L118_Parts, LM_Ins1_L117a_Parts, LM_Ins1_L117_Parts, LM_GI_GI_11_Parts, LM_Ins1_L101_Parts, LM_Ins1_L101a_Parts, LM_Ins2_L25_Parts, LM_Ins2_L9_Parts, LM_Ins2_L66_Parts, LM_Ins2_L113_Parts)
Dim BP_Playfield: BP_Playfield=Array(BM_Playfield, LM_Ins1_L103_Playfield, LM_Ins1_L1_Playfield, LM_Ins1_L17_Playfield, LM_Ins1_L33_Playfield, LM_Ins1_L49_Playfield, LM_Ins1_L2_Playfield, LM_Ins1_L18_Playfield, LM_Ins1_L34_Playfield, LM_Ins1_L35_Playfield, LM_Ins1_L50_Playfield, LM_Ins1_L3_Playfield, LM_Ins1_L19_Playfield, LM_Ins1_L36_Playfield, LM_Ins1_L20_Playfield, LM_Ins1_L4_Playfield, LM_Ins1_L51_Playfield, LM_Ins1_L28_Playfield, LM_Ins1_L60_Playfield, LM_Ins1_L10_Playfield, LM_Ins1_L26_Playfield, LM_Ins1_L42_Playfield, LM_Ins1_L58_Playfield, LM_Ins1_L65_Playfield, LM_Ins1_L81_Playfield, LM_Ins1_L97_Playfield, LM_Ins1_L82_Playfield, LM_GI_gi1_Playfield, LM_Ins1_L114_Playfield, LM_Ins1_L40_Playfield, LM_Ins1_L38_Playfield, LM_Ins1_L24_Playfield, LM_Ins1_L22_Playfield, LM_Ins1_L8_Playfield, LM_Ins1_L6_Playfield, LM_Ins1_L55_Playfield, LM_Ins1_L53_Playfield, LM_Ins1_L39_Playfield, LM_Ins1_L37_Playfield, LM_Ins1_L23_Playfield, LM_Ins1_L21_Playfield, LM_Ins1_L7_Playfield, LM_Ins1_L5_Playfield, _
  LM_Ins1_L63_Playfield, LM_Ins1_L54_Playfield, LM_Ins1_L115_Playfield, LM_Ins1_L102_Playfield, LM_Ins1_L70_Playfield, LM_Ins1_L86_Playfield, LM_Ins1_L59_Playfield, LM_GI_GI_9_Playfield, LM_GI_GI_14_Playfield, LM_GI_GI_10_Playfield, LM_GI_GI_15_Playfield, LM_GI_GI_13_Playfield, LM_GI_GI_4_Playfield, LM_GI_GI_8_Playfield, LM_GI_GI_21_Playfield, LM_GI_GI_19_Playfield, LM_GI_GI_17_Playfield, LM_GI_GI_25_Playfield, LM_GI_GI_24_Playfield, LM_GI_GI_23_Playfield, LM_GI_GI_22_Playfield, LM_GI_GI_1_Playfield, LM_GI_GI_6_Playfield, LM_GI_GI_18_Playfield, LM_GI_GI_20_Playfield, LM_Ins1_L119a_Playfield, LM_Ins1_L119_Playfield, LM_Ins1_L118a_Playfield, LM_Ins1_L118_Playfield, LM_Ins1_L117a_Playfield, LM_Ins1_L117_Playfield, LM_GI_GI_11_Playfield, LM_Ins1_L101_Playfield, LM_Ins1_L101a_Playfield, LM_Ins2_L12_Playfield, LM_Ins2_L41_Playfield, LM_Ins2_L57_Playfield, LM_Ins2_L67_Playfield, LM_Ins2_L83_Playfield, LM_Ins2_L99_Playfield, LM_Ins2_L25_Playfield, LM_Ins2_L9_Playfield, LM_Ins2_L44_Playfield, LM_Ins2_L47_Playfield, _
  LM_Ins2_L113_Playfield, LM_Ins2_L68_Playfield, LM_Ins2_L84_Playfield, LM_Ins2_L100_Playfield, LM_Ins2_L116_Playfield, LM_Ins2_L69_Playfield, LM_Ins2_L85_Playfield, LM_Ins2_L71_Playfield, LM_Ins2_L87_Playfield)
Dim BP_ROstar: BP_ROstar=Array(BM_ROstar, LM_GI_gi1_ROstar, LM_GI_GI_22_ROstar, LM_GI_GI_18_ROstar, LM_GI_GI_20_ROstar, LM_Ins1_L119_ROstar, LM_Ins1_L118_ROstar, LM_Ins1_L117_ROstar)
Dim BP_bulbs: BP_bulbs=Array(BM_bulbs, LM_Ins1_L97_bulbs, LM_Ins1_L82_bulbs, LM_GI_gi1_bulbs, LM_Ins1_L115_bulbs, LM_GI_GI_9_bulbs, LM_GI_GI_14_bulbs, LM_GI_GI_10_bulbs, LM_GI_GI_15_bulbs, LM_GI_GI_13_bulbs, LM_GI_GI_4_bulbs, LM_GI_GI_8_bulbs, LM_GI_GI_21_bulbs, LM_GI_GI_19_bulbs, LM_GI_GI_17_bulbs, LM_GI_GI_25_bulbs, LM_GI_GI_24_bulbs, LM_GI_GI_23_bulbs, LM_GI_GI_22_bulbs, LM_GI_GI_1_bulbs, LM_GI_GI_6_bulbs, LM_GI_GI_18_bulbs, LM_GI_GI_20_bulbs, LM_Ins1_L119a_bulbs, LM_Ins1_L119_bulbs, LM_Ins1_L118a_bulbs, LM_Ins1_L118_bulbs, LM_Ins1_L117a_bulbs, LM_Ins1_L117_bulbs, LM_GI_GI_11_bulbs, LM_Ins1_L101_bulbs, LM_Ins1_L101a_bulbs, LM_Ins2_L12_bulbs, LM_Ins2_L41_bulbs, LM_Ins2_L57_bulbs, LM_Ins2_L67_bulbs, LM_Ins2_L83_bulbs, LM_Ins2_L99_bulbs, LM_Ins2_L25_bulbs, LM_Ins2_L9_bulbs, LM_Ins2_L44_bulbs, LM_Ins2_L66_bulbs, LM_Ins2_L98_bulbs, LM_Ins2_L47_bulbs, LM_Ins2_L113_bulbs, LM_Ins2_L68_bulbs, LM_Ins2_L84_bulbs, LM_Ins2_L100_bulbs, LM_Ins2_L116_bulbs, LM_Ins2_L69_bulbs, LM_Ins2_L85_bulbs, LM_Ins2_L71_bulbs, _
  LM_Ins2_L87_bulbs, LM_Ins2_L11_bulbs)
Dim BP_bumpcapA_3_002: BP_bumpcapA_3_002=Array(BM_bumpcapA_3_002, LM_Ins1_L102_bumpcapA_3_002, LM_Ins1_L70_bumpcapA_3_002, LM_Ins1_L86_bumpcapA_3_002, LM_GI_GI_21_bumpcapA_3_002)
Dim BP_gatewire3: BP_gatewire3=Array(BM_gatewire3, LM_Ins1_L119a_gatewire3)
Dim BP_gatewire4: BP_gatewire4=Array(BM_gatewire4, LM_GI_gi1_gatewire4, LM_Ins1_L118_gatewire4)
Dim BP_ins_mesh: BP_ins_mesh=Array(BM_ins_mesh, LM_Ins1_L103_ins_mesh, LM_Ins1_L50_ins_mesh, LM_Ins1_L3_ins_mesh, LM_Ins1_L19_ins_mesh, LM_Ins1_L36_ins_mesh, LM_Ins1_L20_ins_mesh, LM_Ins1_L4_ins_mesh, LM_Ins1_L51_ins_mesh, LM_Ins1_L28_ins_mesh, LM_Ins1_L60_ins_mesh, LM_Ins1_L10_ins_mesh, LM_Ins1_L26_ins_mesh, LM_Ins1_L42_ins_mesh, LM_Ins1_L58_ins_mesh, LM_Ins1_L65_ins_mesh, LM_Ins1_L81_ins_mesh, LM_Ins1_L97_ins_mesh, LM_Ins1_L114_ins_mesh, LM_Ins1_L40_ins_mesh, LM_Ins1_L38_ins_mesh, LM_Ins1_L24_ins_mesh, LM_Ins1_L22_ins_mesh, LM_Ins1_L8_ins_mesh, LM_Ins1_L6_ins_mesh, LM_Ins1_L55_ins_mesh, LM_Ins1_L53_ins_mesh, LM_Ins1_L39_ins_mesh, LM_Ins1_L37_ins_mesh, LM_Ins1_L23_ins_mesh, LM_Ins1_L21_ins_mesh, LM_Ins1_L7_ins_mesh, LM_Ins1_L5_ins_mesh, LM_Ins1_L63_ins_mesh, LM_GI_GI_13_ins_mesh, LM_GI_GI_8_ins_mesh, LM_GI_GI_21_ins_mesh, LM_GI_GI_11_ins_mesh, LM_Ins1_L101a_ins_mesh, LM_Ins2_L12_ins_mesh, LM_Ins2_L41_ins_mesh, LM_Ins2_L57_ins_mesh, LM_Ins2_L67_ins_mesh, LM_Ins2_L83_ins_mesh, LM_Ins2_L99_ins_mesh, _
  LM_Ins2_L25_ins_mesh, LM_Ins2_L9_ins_mesh, LM_Ins2_L44_ins_mesh, LM_Ins2_L66_ins_mesh, LM_Ins2_L98_ins_mesh, LM_Ins2_L47_ins_mesh, LM_Ins2_L113_ins_mesh, LM_Ins2_L68_ins_mesh, LM_Ins2_L84_ins_mesh, LM_Ins2_L100_ins_mesh, LM_Ins2_L116_ins_mesh, LM_Ins2_L69_ins_mesh, LM_Ins2_L85_ins_mesh, LM_Ins2_L71_ins_mesh, LM_Ins2_L87_ins_mesh, LM_Ins2_L11_ins_mesh)
Dim BP_lsling: BP_lsling=Array(BM_lsling, LM_GI_gi1_lsling, LM_GI_GI_9_lsling, LM_GI_GI_14_lsling, LM_GI_GI_10_lsling, LM_GI_GI_15_lsling, LM_GI_GI_13_lsling, LM_GI_GI_8_lsling, LM_GI_GI_11_lsling)
Dim BP_lsling1: BP_lsling1=Array(BM_lsling1, LM_GI_gi1_lsling1, LM_GI_GI_9_lsling1, LM_GI_GI_14_lsling1, LM_GI_GI_10_lsling1, LM_GI_GI_15_lsling1, LM_GI_GI_13_lsling1, LM_GI_GI_8_lsling1, LM_GI_GI_11_lsling1)
Dim BP_lsling2: BP_lsling2=Array(BM_lsling2, LM_GI_gi1_lsling2, LM_GI_GI_9_lsling2, LM_GI_GI_14_lsling2, LM_GI_GI_10_lsling2, LM_GI_GI_15_lsling2, LM_GI_GI_13_lsling2, LM_GI_GI_8_lsling2, LM_GI_GI_11_lsling2)
Dim BP_plastics: BP_plastics=Array(BM_plastics, LM_Ins1_L10_plastics, LM_Ins1_L26_plastics, LM_Ins1_L82_plastics, LM_GI_gi1_plastics, LM_Ins1_L54_plastics, LM_Ins1_L115_plastics, LM_Ins1_L70_plastics, LM_GI_GI_9_plastics, LM_GI_GI_14_plastics, LM_GI_GI_10_plastics, LM_GI_GI_15_plastics, LM_GI_GI_13_plastics, LM_GI_GI_4_plastics, LM_GI_GI_8_plastics, LM_GI_GI_21_plastics, LM_GI_GI_19_plastics, LM_GI_GI_17_plastics, LM_GI_GI_25_plastics, LM_GI_GI_24_plastics, LM_GI_GI_22_plastics, LM_GI_GI_1_plastics, LM_GI_GI_6_plastics, LM_GI_GI_18_plastics, LM_GI_GI_20_plastics, LM_Ins1_L119a_plastics, LM_Ins1_L119_plastics, LM_Ins1_L118a_plastics, LM_Ins1_L118_plastics, LM_Ins1_L117a_plastics, LM_Ins1_L117_plastics, LM_GI_GI_11_plastics, LM_Ins1_L101_plastics, LM_Ins1_L101a_plastics, LM_Ins2_L41_plastics, LM_Ins2_L57_plastics, LM_Ins2_L66_plastics)
Dim BP_psw1: BP_psw1=Array(BM_psw1, LM_Ins1_L65_psw1, LM_Ins1_L81_psw1, LM_Ins1_L97_psw1, LM_Ins1_L82_psw1, LM_GI_gi1_psw1, LM_Ins1_L114_psw1, LM_GI_GI_9_psw1, LM_GI_GI_15_psw1, LM_GI_GI_13_psw1, LM_GI_GI_4_psw1, LM_GI_GI_8_psw1, LM_GI_GI_23_psw1, LM_GI_GI_18_psw1, LM_GI_GI_20_psw1, LM_Ins1_L119_psw1, LM_Ins1_L118_psw1, LM_Ins1_L117_psw1, LM_Ins1_L101_psw1, LM_Ins2_L66_psw1, LM_Ins2_L113_psw1)
Dim BP_psw17: BP_psw17=Array(BM_psw17, LM_GI_gi1_psw17, LM_Ins1_L38_psw17, LM_Ins1_L24_psw17, LM_Ins1_L22_psw17, LM_Ins1_L8_psw17, LM_Ins1_L6_psw17, LM_Ins1_L55_psw17, LM_Ins1_L53_psw17, LM_Ins1_L39_psw17, LM_Ins1_L37_psw17, LM_Ins1_L23_psw17, LM_Ins1_L21_psw17, LM_Ins1_L7_psw17, LM_Ins1_L5_psw17, LM_Ins1_L115_psw17, LM_GI_GI_9_psw17, LM_GI_GI_4_psw17, LM_GI_GI_8_psw17, LM_GI_GI_21_psw17, LM_GI_GI_25_psw17, LM_GI_GI_24_psw17, LM_GI_GI_1_psw17, LM_GI_GI_20_psw17, LM_Ins1_L118a_psw17, LM_Ins1_L118_psw17, LM_Ins1_L117a_psw17, LM_Ins1_L117_psw17, LM_Ins1_L101a_psw17, LM_Ins2_L67_psw17, LM_Ins2_L83_psw17, LM_Ins2_L99_psw17)
Dim BP_psw18: BP_psw18=Array(BM_psw18, LM_GI_gi1_psw18, LM_Ins1_L40_psw18, LM_Ins1_L38_psw18, LM_Ins1_L24_psw18, LM_Ins1_L22_psw18, LM_Ins1_L8_psw18, LM_Ins1_L6_psw18, LM_Ins1_L55_psw18, LM_Ins1_L53_psw18, LM_Ins1_L39_psw18, LM_Ins1_L37_psw18, LM_Ins1_L23_psw18, LM_Ins1_L21_psw18, LM_Ins1_L7_psw18, LM_Ins1_L5_psw18, LM_Ins1_L63_psw18, LM_GI_GI_9_psw18, LM_GI_GI_14_psw18, LM_GI_GI_15_psw18, LM_GI_GI_4_psw18, LM_GI_GI_8_psw18, LM_GI_GI_21_psw18, LM_GI_GI_25_psw18, LM_GI_GI_20_psw18, LM_Ins1_L118a_psw18, LM_Ins1_L117a_psw18, LM_Ins1_L101a_psw18, LM_Ins2_L67_psw18, LM_Ins2_L83_psw18)
Dim BP_psw19: BP_psw19=Array(BM_psw19, LM_Ins1_L36_psw19, LM_GI_gi1_psw19, LM_Ins1_L40_psw19, LM_Ins1_L38_psw19, LM_Ins1_L24_psw19, LM_Ins1_L22_psw19, LM_Ins1_L8_psw19, LM_Ins1_L6_psw19, LM_Ins1_L55_psw19, LM_Ins1_L53_psw19, LM_Ins1_L39_psw19, LM_Ins1_L37_psw19, LM_Ins1_L23_psw19, LM_Ins1_L21_psw19, LM_Ins1_L7_psw19, LM_Ins1_L5_psw19, LM_GI_GI_9_psw19, LM_GI_GI_14_psw19, LM_GI_GI_4_psw19, LM_GI_GI_8_psw19, LM_GI_GI_21_psw19, LM_GI_GI_25_psw19, LM_GI_GI_24_psw19, LM_GI_GI_23_psw19, LM_GI_GI_22_psw19, LM_GI_GI_20_psw19, LM_GI_GI_11_psw19, LM_Ins2_L67_psw19, LM_Ins2_L83_psw19)
Dim BP_psw2: BP_psw2=Array(BM_psw2, LM_Ins1_L81_psw2, LM_Ins1_L97_psw2, LM_Ins1_L82_psw2, LM_GI_gi1_psw2, LM_Ins1_L114_psw2, LM_Ins1_L63_psw2, LM_GI_GI_9_psw2, LM_GI_GI_15_psw2, LM_GI_GI_13_psw2, LM_GI_GI_8_psw2, LM_GI_GI_21_psw2, LM_GI_GI_24_psw2, LM_GI_GI_23_psw2, LM_GI_GI_18_psw2, LM_GI_GI_20_psw2, LM_Ins1_L119_psw2, LM_Ins1_L118_psw2, LM_Ins1_L117_psw2, LM_Ins1_L101_psw2)
Dim BP_psw20: BP_psw20=Array(BM_psw20, LM_Ins1_L36_psw20, LM_Ins1_L42_psw20, LM_GI_gi1_psw20, LM_Ins1_L40_psw20, LM_Ins1_L38_psw20, LM_Ins1_L24_psw20, LM_Ins1_L22_psw20, LM_Ins1_L8_psw20, LM_Ins1_L6_psw20, LM_Ins1_L55_psw20, LM_Ins1_L53_psw20, LM_Ins1_L39_psw20, LM_Ins1_L37_psw20, LM_Ins1_L23_psw20, LM_Ins1_L21_psw20, LM_Ins1_L7_psw20, LM_Ins1_L5_psw20, LM_GI_GI_9_psw20, LM_GI_GI_14_psw20, LM_GI_GI_15_psw20, LM_GI_GI_13_psw20, LM_GI_GI_4_psw20, LM_GI_GI_8_psw20, LM_GI_GI_21_psw20, LM_GI_GI_25_psw20, LM_GI_GI_24_psw20, LM_GI_GI_22_psw20, LM_GI_GI_11_psw20, LM_Ins1_L101_psw20)
Dim BP_psw21: BP_psw21=Array(BM_psw21, LM_Ins1_L3_psw21, LM_Ins1_L19_psw21, LM_Ins1_L36_psw21, LM_Ins1_L20_psw21, LM_GI_gi1_psw21, LM_Ins1_L40_psw21, LM_Ins1_L38_psw21, LM_Ins1_L24_psw21, LM_Ins1_L22_psw21, LM_Ins1_L8_psw21, LM_Ins1_L6_psw21, LM_Ins1_L55_psw21, LM_Ins1_L53_psw21, LM_Ins1_L39_psw21, LM_Ins1_L37_psw21, LM_Ins1_L23_psw21, LM_Ins1_L21_psw21, LM_Ins1_L7_psw21, LM_Ins1_L5_psw21, LM_Ins1_L63_psw21, LM_GI_GI_9_psw21, LM_GI_GI_14_psw21, LM_GI_GI_10_psw21, LM_GI_GI_13_psw21, LM_GI_GI_4_psw21, LM_GI_GI_8_psw21, LM_GI_GI_21_psw21, LM_GI_GI_25_psw21, LM_GI_GI_24_psw21, LM_GI_GI_22_psw21, LM_GI_GI_11_psw21, LM_Ins1_L101_psw21)
Dim BP_psw22: BP_psw22=Array(BM_psw22, LM_Ins1_L18_psw22, LM_Ins1_L3_psw22, LM_Ins1_L19_psw22, LM_Ins1_L36_psw22, LM_Ins1_L20_psw22, LM_GI_gi1_psw22, LM_Ins1_L40_psw22, LM_Ins1_L38_psw22, LM_Ins1_L24_psw22, LM_Ins1_L22_psw22, LM_Ins1_L8_psw22, LM_Ins1_L6_psw22, LM_Ins1_L55_psw22, LM_Ins1_L53_psw22, LM_Ins1_L39_psw22, LM_Ins1_L37_psw22, LM_Ins1_L23_psw22, LM_Ins1_L21_psw22, LM_Ins1_L7_psw22, LM_Ins1_L5_psw22, LM_Ins1_L63_psw22, LM_GI_GI_14_psw22, LM_GI_GI_10_psw22, LM_GI_GI_13_psw22, LM_GI_GI_4_psw22, LM_GI_GI_8_psw22, LM_GI_GI_21_psw22, LM_GI_GI_24_psw22, LM_GI_GI_22_psw22, LM_GI_GI_20_psw22, LM_GI_GI_11_psw22, LM_Ins1_L101_psw22)
Dim BP_psw23: BP_psw23=Array(BM_psw23, LM_Ins1_L18_psw23, LM_Ins1_L3_psw23, LM_Ins1_L19_psw23, LM_Ins1_L36_psw23, LM_Ins1_L20_psw23, LM_Ins1_L42_psw23, LM_GI_gi1_psw23, LM_Ins1_L40_psw23, LM_Ins1_L38_psw23, LM_Ins1_L24_psw23, LM_Ins1_L22_psw23, LM_Ins1_L8_psw23, LM_Ins1_L6_psw23, LM_Ins1_L55_psw23, LM_Ins1_L53_psw23, LM_Ins1_L39_psw23, LM_Ins1_L37_psw23, LM_Ins1_L23_psw23, LM_Ins1_L21_psw23, LM_Ins1_L5_psw23, LM_Ins1_L63_psw23, LM_GI_GI_10_psw23, LM_GI_GI_15_psw23, LM_GI_GI_13_psw23, LM_GI_GI_4_psw23, LM_GI_GI_8_psw23, LM_GI_GI_21_psw23, LM_GI_GI_24_psw23, LM_GI_GI_22_psw23, LM_Ins1_L117_psw23, LM_GI_GI_11_psw23, LM_Ins1_L101_psw23)
Dim BP_psw3: BP_psw3=Array(BM_psw3, LM_Ins1_L82_psw3, LM_GI_gi1_psw3, LM_Ins1_L114_psw3, LM_Ins1_L63_psw3, LM_GI_GI_15_psw3, LM_GI_GI_24_psw3, LM_GI_GI_23_psw3, LM_GI_GI_18_psw3, LM_GI_GI_20_psw3, LM_Ins1_L119_psw3, LM_Ins1_L118a_psw3, LM_Ins1_L118_psw3, LM_Ins1_L117_psw3, LM_Ins1_L101_psw3)
Dim BP_psw33: BP_psw33=Array(BM_psw33, LM_GI_gi1_psw33, LM_Ins1_L54_psw33, LM_Ins1_L115_psw33, LM_GI_GI_21_psw33, LM_GI_GI_19_psw33, LM_GI_GI_17_psw33, LM_GI_GI_25_psw33, LM_GI_GI_1_psw33, LM_GI_GI_6_psw33, LM_Ins1_L119a_psw33, LM_Ins1_L118a_psw33, LM_Ins1_L117a_psw33, LM_Ins1_L117_psw33, LM_Ins1_L101a_psw33)
Dim BP_psw4: BP_psw4=Array(BM_psw4, LM_Ins1_L82_psw4, LM_GI_gi1_psw4, LM_GI_GI_24_psw4, LM_GI_GI_23_psw4, LM_GI_GI_22_psw4, LM_GI_GI_18_psw4, LM_GI_GI_20_psw4, LM_Ins1_L119_psw4, LM_Ins1_L118_psw4, LM_Ins1_L117_psw4, LM_Ins1_L101_psw4)
Dim BP_rails: BP_rails=Array(BM_rails, LM_GI_gi1_rails, LM_GI_GI_9_rails, LM_GI_GI_15_rails, LM_GI_GI_13_rails, LM_GI_GI_4_rails, LM_GI_GI_8_rails, LM_GI_GI_21_rails, LM_GI_GI_17_rails, LM_GI_GI_25_rails, LM_GI_GI_20_rails, LM_Ins1_L119_rails, LM_Ins1_L118a_rails, LM_Ins1_L118_rails, LM_GI_GI_11_rails, LM_Ins1_L101a_rails)
Dim BP_rsling: BP_rsling=Array(BM_rsling, LM_GI_GI_9_rsling, LM_GI_GI_14_rsling, LM_GI_GI_10_rsling, LM_GI_GI_15_rsling, LM_GI_GI_13_rsling, LM_GI_GI_4_rsling, LM_GI_GI_11_rsling)
Dim BP_rsling1: BP_rsling1=Array(BM_rsling1, LM_GI_GI_9_rsling1, LM_GI_GI_14_rsling1, LM_GI_GI_10_rsling1, LM_GI_GI_15_rsling1, LM_GI_GI_13_rsling1, LM_GI_GI_4_rsling1, LM_GI_GI_11_rsling1)
Dim BP_rsling2: BP_rsling2=Array(BM_rsling2, LM_GI_GI_9_rsling2, LM_GI_GI_14_rsling2, LM_GI_GI_10_rsling2, LM_GI_GI_15_rsling2, LM_GI_GI_13_rsling2, LM_GI_GI_4_rsling2, LM_GI_GI_11_rsling2)
Dim BP_slingarmL: BP_slingarmL=Array(BM_slingarmL, LM_GI_GI_9_slingarmL, LM_GI_GI_11_slingarmL)
Dim BP_slingarmR: BP_slingarmR=Array(BM_slingarmR, LM_GI_GI_15_slingarmR, LM_GI_GI_13_slingarmR)
Dim BP_sw12: BP_sw12=Array(BM_sw12, LM_GI_GI_24_sw12, LM_GI_GI_23_sw12, LM_GI_GI_22_sw12)
Dim BP_sw13: BP_sw13=Array(BM_sw13, LM_GI_GI_24_sw13, LM_GI_GI_23_sw13, LM_GI_GI_22_sw13)
Dim BP_sw14: BP_sw14=Array(BM_sw14, LM_Ins1_L34_sw14, LM_GI_GI_14_sw14, LM_GI_GI_8_sw14, LM_GI_GI_11_sw14)
Dim BP_sw15: BP_sw15=Array(BM_sw15, LM_GI_GI_13_sw15)
Dim BP_sw31: BP_sw31=Array(BM_sw31, LM_Ins1_L19_sw31, LM_GI_GI_9_sw31, LM_GI_GI_13_sw31)
Dim BP_sw32: BP_sw32=Array(BM_sw32, LM_GI_GI_15_sw32, LM_GI_GI_11_sw32)
Dim BP_tsw25: BP_tsw25=Array(BM_tsw25, LM_GI_gi1_tsw25, LM_Ins1_L21_tsw25, LM_Ins1_L5_tsw25, LM_GI_GI_21_tsw25, LM_GI_GI_20_tsw25, LM_Ins1_L118a_tsw25)
Dim BP_tsw26: BP_tsw26=Array(BM_tsw26, LM_GI_gi1_tsw26, LM_GI_GI_4_tsw26, LM_GI_GI_21_tsw26)
Dim BP_tsw27: BP_tsw27=Array(BM_tsw27, LM_GI_gi1_tsw27, LM_Ins1_L37_tsw27, LM_GI_GI_4_tsw27, LM_GI_GI_21_tsw27, LM_Ins1_L101a_tsw27)
Dim BP_tsw28: BP_tsw28=Array(BM_tsw28, LM_GI_gi1_tsw28, LM_Ins1_L53_tsw28, LM_GI_GI_4_tsw28, LM_GI_GI_21_tsw28)
Dim BP_tsw29: BP_tsw29=Array(BM_tsw29, LM_GI_gi1_tsw29, LM_GI_GI_4_tsw29, LM_GI_GI_21_tsw29)
Dim BP_tsw30: BP_tsw30=Array(BM_tsw30, LM_GI_gi1_tsw30, LM_Ins1_L22_tsw30, LM_GI_GI_13_tsw30, LM_GI_GI_4_tsw30)
Dim BP_tsw5: BP_tsw5=Array(BM_tsw5, LM_GI_gi1_tsw5, LM_GI_GI_23_tsw5, LM_GI_GI_22_tsw5, LM_GI_GI_18_tsw5, LM_Ins1_L119_tsw5, LM_Ins1_L118_tsw5, LM_Ins1_L117_tsw5, LM_Ins1_L101_tsw5)
' Arrays per lighting scenario
Dim BL_GI_GI_1: BL_GI_GI_1=Array(LM_GI_GI_1_Layer_2, LM_GI_GI_1_Parts, LM_GI_GI_1_Playfield, LM_GI_GI_1_bulbs, LM_GI_GI_1_plastics, LM_GI_GI_1_psw17, LM_GI_GI_1_psw33)
Dim BL_GI_GI_10: BL_GI_GI_10=Array(LM_GI_GI_10_Flip1, LM_GI_GI_10_Flip2, LM_GI_GI_10_Layer_2, LM_GI_GI_10_Parts, LM_GI_GI_10_Playfield, LM_GI_GI_10_bulbs, LM_GI_GI_10_lsling, LM_GI_GI_10_lsling1, LM_GI_GI_10_lsling2, LM_GI_GI_10_plastics, LM_GI_GI_10_psw21, LM_GI_GI_10_psw22, LM_GI_GI_10_psw23, LM_GI_GI_10_rsling, LM_GI_GI_10_rsling1, LM_GI_GI_10_rsling2)
Dim BL_GI_GI_11: BL_GI_GI_11=Array(LM_GI_GI_11_Flip1, LM_GI_GI_11_Flip2, LM_GI_GI_11_Flip3, LM_GI_GI_11_Layer_2, LM_GI_GI_11_Parts, LM_GI_GI_11_Playfield, LM_GI_GI_11_bulbs, LM_GI_GI_11_ins_mesh, LM_GI_GI_11_lsling, LM_GI_GI_11_lsling1, LM_GI_GI_11_lsling2, LM_GI_GI_11_plastics, LM_GI_GI_11_psw19, LM_GI_GI_11_psw20, LM_GI_GI_11_psw21, LM_GI_GI_11_psw22, LM_GI_GI_11_psw23, LM_GI_GI_11_rails, LM_GI_GI_11_rsling, LM_GI_GI_11_rsling1, LM_GI_GI_11_rsling2, LM_GI_GI_11_slingarmL, LM_GI_GI_11_sw14, LM_GI_GI_11_sw32)
Dim BL_GI_GI_13: BL_GI_GI_13=Array(LM_GI_GI_13_Flip1, LM_GI_GI_13_Flip2, LM_GI_GI_13_Flip3, LM_GI_GI_13_Layer_2, LM_GI_GI_13_Parts, LM_GI_GI_13_Playfield, LM_GI_GI_13_bulbs, LM_GI_GI_13_ins_mesh, LM_GI_GI_13_lsling, LM_GI_GI_13_lsling1, LM_GI_GI_13_lsling2, LM_GI_GI_13_plastics, LM_GI_GI_13_psw1, LM_GI_GI_13_psw2, LM_GI_GI_13_psw20, LM_GI_GI_13_psw21, LM_GI_GI_13_psw22, LM_GI_GI_13_psw23, LM_GI_GI_13_rails, LM_GI_GI_13_rsling, LM_GI_GI_13_rsling1, LM_GI_GI_13_rsling2, LM_GI_GI_13_slingarmR, LM_GI_GI_13_sw15, LM_GI_GI_13_sw31, LM_GI_GI_13_tsw30)
Dim BL_GI_GI_14: BL_GI_GI_14=Array(LM_GI_GI_14_Flip1, LM_GI_GI_14_Flip2, LM_GI_GI_14_Flip3, LM_GI_GI_14_Layer_2, LM_GI_GI_14_Parts, LM_GI_GI_14_Playfield, LM_GI_GI_14_bulbs, LM_GI_GI_14_lsling, LM_GI_GI_14_lsling1, LM_GI_GI_14_lsling2, LM_GI_GI_14_plastics, LM_GI_GI_14_psw18, LM_GI_GI_14_psw19, LM_GI_GI_14_psw20, LM_GI_GI_14_psw21, LM_GI_GI_14_psw22, LM_GI_GI_14_rsling, LM_GI_GI_14_rsling1, LM_GI_GI_14_rsling2, LM_GI_GI_14_sw14)
Dim BL_GI_GI_15: BL_GI_GI_15=Array(LM_GI_GI_15_Flip1, LM_GI_GI_15_Layer_2, LM_GI_GI_15_Parts, LM_GI_GI_15_Playfield, LM_GI_GI_15_bulbs, LM_GI_GI_15_lsling, LM_GI_GI_15_lsling1, LM_GI_GI_15_lsling2, LM_GI_GI_15_plastics, LM_GI_GI_15_psw1, LM_GI_GI_15_psw18, LM_GI_GI_15_psw2, LM_GI_GI_15_psw20, LM_GI_GI_15_psw23, LM_GI_GI_15_psw3, LM_GI_GI_15_rails, LM_GI_GI_15_rsling, LM_GI_GI_15_rsling1, LM_GI_GI_15_rsling2, LM_GI_GI_15_slingarmR, LM_GI_GI_15_sw32)
Dim BL_GI_GI_17: BL_GI_GI_17=Array(LM_GI_GI_17_BumperSkirt1, LM_GI_GI_17_BumperSkirt2, LM_GI_GI_17_Layer_2, LM_GI_GI_17_Parts, LM_GI_GI_17_Playfield, LM_GI_GI_17_bulbs, LM_GI_GI_17_plastics, LM_GI_GI_17_psw33, LM_GI_GI_17_rails)
Dim BL_GI_GI_18: BL_GI_GI_18=Array(LM_GI_GI_18_BumperRing1, LM_GI_GI_18_BumperSkirt1, LM_GI_GI_18_BumperSkirt2, LM_GI_GI_18_Layer_2, LM_GI_GI_18_Parts, LM_GI_GI_18_Playfield, LM_GI_GI_18_ROstar, LM_GI_GI_18_bulbs, LM_GI_GI_18_plastics, LM_GI_GI_18_psw1, LM_GI_GI_18_psw2, LM_GI_GI_18_psw3, LM_GI_GI_18_psw4, LM_GI_GI_18_tsw5)
Dim BL_GI_GI_19: BL_GI_GI_19=Array(LM_GI_GI_19_BumperSkirt2, LM_GI_GI_19_Layer_2, LM_GI_GI_19_Parts, LM_GI_GI_19_Playfield, LM_GI_GI_19_bulbs, LM_GI_GI_19_plastics, LM_GI_GI_19_psw33)
Dim BL_GI_GI_20: BL_GI_GI_20=Array(LM_GI_GI_20_BumperRing2, LM_GI_GI_20_BumperRing3, LM_GI_GI_20_BumperSkirt1, LM_GI_GI_20_BumperSkirt2, LM_GI_GI_20_BumperSkirt3, LM_GI_GI_20_Layer_2, LM_GI_GI_20_Parts, LM_GI_GI_20_Playfield, LM_GI_GI_20_ROstar, LM_GI_GI_20_bulbs, LM_GI_GI_20_plastics, LM_GI_GI_20_psw1, LM_GI_GI_20_psw17, LM_GI_GI_20_psw18, LM_GI_GI_20_psw19, LM_GI_GI_20_psw2, LM_GI_GI_20_psw22, LM_GI_GI_20_psw3, LM_GI_GI_20_psw4, LM_GI_GI_20_rails, LM_GI_GI_20_tsw25)
Dim BL_GI_GI_21: BL_GI_GI_21=Array(LM_GI_GI_21_BumperRing1, LM_GI_GI_21_BumperRing2, LM_GI_GI_21_BumperSkirt1, LM_GI_GI_21_BumperSkirt2, LM_GI_GI_21_BumperSkirt3, LM_GI_GI_21_Layer_2, LM_GI_GI_21_Parts, LM_GI_GI_21_Playfield, LM_GI_GI_21_bulbs, LM_GI_GI_21_bumpcapA_3_002, LM_GI_GI_21_ins_mesh, LM_GI_GI_21_plastics, LM_GI_GI_21_psw17, LM_GI_GI_21_psw18, LM_GI_GI_21_psw19, LM_GI_GI_21_psw2, LM_GI_GI_21_psw20, LM_GI_GI_21_psw21, LM_GI_GI_21_psw22, LM_GI_GI_21_psw23, LM_GI_GI_21_psw33, LM_GI_GI_21_rails, LM_GI_GI_21_tsw25, LM_GI_GI_21_tsw26, LM_GI_GI_21_tsw27, LM_GI_GI_21_tsw28, LM_GI_GI_21_tsw29)
Dim BL_GI_GI_22: BL_GI_GI_22=Array(LM_GI_GI_22_BumperSkirt1, LM_GI_GI_22_BumperSkirt2, LM_GI_GI_22_Gateplate2, LM_GI_GI_22_Lane_Guide___2a, LM_GI_GI_22_Layer_2, LM_GI_GI_22_Parts, LM_GI_GI_22_Playfield, LM_GI_GI_22_ROstar, LM_GI_GI_22_bulbs, LM_GI_GI_22_plastics, LM_GI_GI_22_psw19, LM_GI_GI_22_psw20, LM_GI_GI_22_psw21, LM_GI_GI_22_psw22, LM_GI_GI_22_psw23, LM_GI_GI_22_psw4, LM_GI_GI_22_sw12, LM_GI_GI_22_sw13, LM_GI_GI_22_tsw5)
Dim BL_GI_GI_23: BL_GI_GI_23=Array(LM_GI_GI_23_BumperSkirt1, LM_GI_GI_23_BumperSkirt2, LM_GI_GI_23_Lane_Guide___2a, LM_GI_GI_23_Layer_2, LM_GI_GI_23_Parts, LM_GI_GI_23_Playfield, LM_GI_GI_23_bulbs, LM_GI_GI_23_psw1, LM_GI_GI_23_psw19, LM_GI_GI_23_psw2, LM_GI_GI_23_psw3, LM_GI_GI_23_psw4, LM_GI_GI_23_sw12, LM_GI_GI_23_sw13, LM_GI_GI_23_tsw5)
Dim BL_GI_GI_24: BL_GI_GI_24=Array(LM_GI_GI_24_BumperSkirt1, LM_GI_GI_24_BumperSkirt2, LM_GI_GI_24_Lane_Guide___2a, LM_GI_GI_24_Layer_2, LM_GI_GI_24_Parts, LM_GI_GI_24_Playfield, LM_GI_GI_24_bulbs, LM_GI_GI_24_plastics, LM_GI_GI_24_psw17, LM_GI_GI_24_psw19, LM_GI_GI_24_psw2, LM_GI_GI_24_psw20, LM_GI_GI_24_psw21, LM_GI_GI_24_psw22, LM_GI_GI_24_psw23, LM_GI_GI_24_psw3, LM_GI_GI_24_psw4, LM_GI_GI_24_sw12, LM_GI_GI_24_sw13)
Dim BL_GI_GI_25: BL_GI_GI_25=Array(LM_GI_GI_25_BumperSkirt1, LM_GI_GI_25_BumperSkirt2, LM_GI_GI_25_Layer_2, LM_GI_GI_25_Parts, LM_GI_GI_25_Playfield, LM_GI_GI_25_bulbs, LM_GI_GI_25_plastics, LM_GI_GI_25_psw17, LM_GI_GI_25_psw18, LM_GI_GI_25_psw19, LM_GI_GI_25_psw20, LM_GI_GI_25_psw21, LM_GI_GI_25_psw33, LM_GI_GI_25_rails)
Dim BL_GI_GI_4: BL_GI_GI_4=Array(LM_GI_GI_4_BumperRing3, LM_GI_GI_4_BumperSkirt3, LM_GI_GI_4_Flip1, LM_GI_GI_4_Flip3, LM_GI_GI_4_Parts, LM_GI_GI_4_Playfield, LM_GI_GI_4_bulbs, LM_GI_GI_4_plastics, LM_GI_GI_4_psw1, LM_GI_GI_4_psw17, LM_GI_GI_4_psw18, LM_GI_GI_4_psw19, LM_GI_GI_4_psw20, LM_GI_GI_4_psw21, LM_GI_GI_4_psw22, LM_GI_GI_4_psw23, LM_GI_GI_4_rails, LM_GI_GI_4_rsling, LM_GI_GI_4_rsling1, LM_GI_GI_4_rsling2, LM_GI_GI_4_tsw26, LM_GI_GI_4_tsw27, LM_GI_GI_4_tsw28, LM_GI_GI_4_tsw29, LM_GI_GI_4_tsw30)
Dim BL_GI_GI_6: BL_GI_GI_6=Array(LM_GI_GI_6_Layer_2, LM_GI_GI_6_Parts, LM_GI_GI_6_Playfield, LM_GI_GI_6_bulbs, LM_GI_GI_6_plastics, LM_GI_GI_6_psw33)
Dim BL_GI_GI_8: BL_GI_GI_8=Array(LM_GI_GI_8_Flip1, LM_GI_GI_8_Flip2, LM_GI_GI_8_Flip3, LM_GI_GI_8_Parts, LM_GI_GI_8_Playfield, LM_GI_GI_8_bulbs, LM_GI_GI_8_ins_mesh, LM_GI_GI_8_lsling, LM_GI_GI_8_lsling1, LM_GI_GI_8_lsling2, LM_GI_GI_8_plastics, LM_GI_GI_8_psw1, LM_GI_GI_8_psw17, LM_GI_GI_8_psw18, LM_GI_GI_8_psw19, LM_GI_GI_8_psw2, LM_GI_GI_8_psw20, LM_GI_GI_8_psw21, LM_GI_GI_8_psw22, LM_GI_GI_8_psw23, LM_GI_GI_8_rails, LM_GI_GI_8_sw14)
Dim BL_GI_GI_9: BL_GI_GI_9=Array(LM_GI_GI_9_Flip1, LM_GI_GI_9_Flip2, LM_GI_GI_9_Flip3, LM_GI_GI_9_Layer_2, LM_GI_GI_9_Parts, LM_GI_GI_9_Playfield, LM_GI_GI_9_bulbs, LM_GI_GI_9_lsling, LM_GI_GI_9_lsling1, LM_GI_GI_9_lsling2, LM_GI_GI_9_plastics, LM_GI_GI_9_psw1, LM_GI_GI_9_psw17, LM_GI_GI_9_psw18, LM_GI_GI_9_psw19, LM_GI_GI_9_psw2, LM_GI_GI_9_psw20, LM_GI_GI_9_psw21, LM_GI_GI_9_rails, LM_GI_GI_9_rsling, LM_GI_GI_9_rsling1, LM_GI_GI_9_rsling2, LM_GI_GI_9_slingarmL, LM_GI_GI_9_sw31)
Dim BL_GI_gi1: BL_GI_gi1=Array(LM_GI_gi1_BumperSkirt2, LM_GI_gi1_Flip1, LM_GI_gi1_Flip2, LM_GI_gi1_Flip3, LM_GI_gi1_Gateplate5, LM_GI_gi1_Layer_2, LM_GI_gi1_Parts, LM_GI_gi1_Playfield, LM_GI_gi1_ROstar, LM_GI_gi1_bulbs, LM_GI_gi1_gatewire4, LM_GI_gi1_lsling, LM_GI_gi1_lsling1, LM_GI_gi1_lsling2, LM_GI_gi1_plastics, LM_GI_gi1_psw1, LM_GI_gi1_psw17, LM_GI_gi1_psw18, LM_GI_gi1_psw19, LM_GI_gi1_psw2, LM_GI_gi1_psw20, LM_GI_gi1_psw21, LM_GI_gi1_psw22, LM_GI_gi1_psw23, LM_GI_gi1_psw3, LM_GI_gi1_psw33, LM_GI_gi1_psw4, LM_GI_gi1_rails, LM_GI_gi1_tsw25, LM_GI_gi1_tsw26, LM_GI_gi1_tsw27, LM_GI_gi1_tsw28, LM_GI_gi1_tsw29, LM_GI_gi1_tsw30, LM_GI_gi1_tsw5)
Dim BL_Ins1_L1: BL_Ins1_L1=Array(LM_Ins1_L1_Flip1, LM_Ins1_L1_Flip2, LM_Ins1_L1_Playfield)
Dim BL_Ins1_L10: BL_Ins1_L10=Array(LM_Ins1_L10_Parts, LM_Ins1_L10_Playfield, LM_Ins1_L10_ins_mesh, LM_Ins1_L10_plastics)
Dim BL_Ins1_L101: BL_Ins1_L101=Array(LM_Ins1_L101_BumperRing3, LM_Ins1_L101_BumperSkirt1, LM_Ins1_L101_BumperSkirt2, LM_Ins1_L101_BumperSkirt3, LM_Ins1_L101_Flip3, LM_Ins1_L101_Layer_2, LM_Ins1_L101_Parts, LM_Ins1_L101_Playfield, LM_Ins1_L101_bulbs, LM_Ins1_L101_plastics, LM_Ins1_L101_psw1, LM_Ins1_L101_psw2, LM_Ins1_L101_psw20, LM_Ins1_L101_psw21, LM_Ins1_L101_psw22, LM_Ins1_L101_psw23, LM_Ins1_L101_psw3, LM_Ins1_L101_psw4, LM_Ins1_L101_tsw5)
Dim BL_Ins1_L101a: BL_Ins1_L101a=Array(LM_Ins1_L101a_BumperSkirt2, LM_Ins1_L101a_BumperSkirt3, LM_Ins1_L101a_Layer_2, LM_Ins1_L101a_Parts, LM_Ins1_L101a_Playfield, LM_Ins1_L101a_bulbs, LM_Ins1_L101a_ins_mesh, LM_Ins1_L101a_plastics, LM_Ins1_L101a_psw17, LM_Ins1_L101a_psw18, LM_Ins1_L101a_psw33, LM_Ins1_L101a_rails, LM_Ins1_L101a_tsw27)
Dim BL_Ins1_L102: BL_Ins1_L102=Array(LM_Ins1_L102_BumperSkirt3, LM_Ins1_L102_Layer_2, LM_Ins1_L102_Parts, LM_Ins1_L102_Playfield, LM_Ins1_L102_bumpcapA_3_002)
Dim BL_Ins1_L103: BL_Ins1_L103=Array(LM_Ins1_L103_Flip1, LM_Ins1_L103_Flip2, LM_Ins1_L103_Flip3, LM_Ins1_L103_Playfield, LM_Ins1_L103_ins_mesh)
Dim BL_Ins1_L114: BL_Ins1_L114=Array(LM_Ins1_L114_Flip3, LM_Ins1_L114_Parts, LM_Ins1_L114_Playfield, LM_Ins1_L114_ins_mesh, LM_Ins1_L114_psw1, LM_Ins1_L114_psw2, LM_Ins1_L114_psw3)
Dim BL_Ins1_L115: BL_Ins1_L115=Array(LM_Ins1_L115_Layer_2, LM_Ins1_L115_Parts, LM_Ins1_L115_Playfield, LM_Ins1_L115_bulbs, LM_Ins1_L115_plastics, LM_Ins1_L115_psw17, LM_Ins1_L115_psw33)
Dim BL_Ins1_L117: BL_Ins1_L117=Array(LM_Ins1_L117_BumperRing1, LM_Ins1_L117_BumperSkirt1, LM_Ins1_L117_BumperSkirt3, LM_Ins1_L117_Layer_2, LM_Ins1_L117_Parts, LM_Ins1_L117_Playfield, LM_Ins1_L117_ROstar, LM_Ins1_L117_bulbs, LM_Ins1_L117_plastics, LM_Ins1_L117_psw1, LM_Ins1_L117_psw17, LM_Ins1_L117_psw2, LM_Ins1_L117_psw23, LM_Ins1_L117_psw3, LM_Ins1_L117_psw33, LM_Ins1_L117_psw4, LM_Ins1_L117_tsw5)
Dim BL_Ins1_L117a: BL_Ins1_L117a=Array(LM_Ins1_L117a_BumperSkirt2, LM_Ins1_L117a_Layer_2, LM_Ins1_L117a_Parts, LM_Ins1_L117a_Playfield, LM_Ins1_L117a_bulbs, LM_Ins1_L117a_plastics, LM_Ins1_L117a_psw17, LM_Ins1_L117a_psw18, LM_Ins1_L117a_psw33)
Dim BL_Ins1_L118: BL_Ins1_L118=Array(LM_Ins1_L118_BumperSkirt1, LM_Ins1_L118_Gateplate2, LM_Ins1_L118_Layer_2, LM_Ins1_L118_Parts, LM_Ins1_L118_Playfield, LM_Ins1_L118_ROstar, LM_Ins1_L118_bulbs, LM_Ins1_L118_gatewire4, LM_Ins1_L118_plastics, LM_Ins1_L118_psw1, LM_Ins1_L118_psw17, LM_Ins1_L118_psw2, LM_Ins1_L118_psw3, LM_Ins1_L118_psw4, LM_Ins1_L118_rails, LM_Ins1_L118_tsw5)
Dim BL_Ins1_L118a: BL_Ins1_L118a=Array(LM_Ins1_L118a_Layer_2, LM_Ins1_L118a_Parts, LM_Ins1_L118a_Playfield, LM_Ins1_L118a_bulbs, LM_Ins1_L118a_plastics, LM_Ins1_L118a_psw17, LM_Ins1_L118a_psw18, LM_Ins1_L118a_psw3, LM_Ins1_L118a_psw33, LM_Ins1_L118a_rails, LM_Ins1_L118a_tsw25)
Dim BL_Ins1_L119: BL_Ins1_L119=Array(LM_Ins1_L119_BumperSkirt1, LM_Ins1_L119_Gateplate5, LM_Ins1_L119_Layer_2, LM_Ins1_L119_Parts, LM_Ins1_L119_Playfield, LM_Ins1_L119_ROstar, LM_Ins1_L119_bulbs, LM_Ins1_L119_plastics, LM_Ins1_L119_psw1, LM_Ins1_L119_psw2, LM_Ins1_L119_psw3, LM_Ins1_L119_psw4, LM_Ins1_L119_rails, LM_Ins1_L119_tsw5)
Dim BL_Ins1_L119a: BL_Ins1_L119a=Array(LM_Ins1_L119a_BumperSkirt2, LM_Ins1_L119a_Layer_2, LM_Ins1_L119a_Parts, LM_Ins1_L119a_Playfield, LM_Ins1_L119a_bulbs, LM_Ins1_L119a_gatewire3, LM_Ins1_L119a_plastics, LM_Ins1_L119a_psw33)
Dim BL_Ins1_L17: BL_Ins1_L17=Array(LM_Ins1_L17_Flip2, LM_Ins1_L17_Flip3, LM_Ins1_L17_Parts, LM_Ins1_L17_Playfield)
Dim BL_Ins1_L18: BL_Ins1_L18=Array(LM_Ins1_L18_Playfield, LM_Ins1_L18_psw22, LM_Ins1_L18_psw23)
Dim BL_Ins1_L19: BL_Ins1_L19=Array(LM_Ins1_L19_Playfield, LM_Ins1_L19_ins_mesh, LM_Ins1_L19_psw21, LM_Ins1_L19_psw22, LM_Ins1_L19_psw23, LM_Ins1_L19_sw31)
Dim BL_Ins1_L2: BL_Ins1_L2=Array(LM_Ins1_L2_Flip3, LM_Ins1_L2_Playfield)
Dim BL_Ins1_L20: BL_Ins1_L20=Array(LM_Ins1_L20_Flip3, LM_Ins1_L20_Playfield, LM_Ins1_L20_ins_mesh, LM_Ins1_L20_psw21, LM_Ins1_L20_psw22, LM_Ins1_L20_psw23)
Dim BL_Ins1_L21: BL_Ins1_L21=Array(LM_Ins1_L21_Playfield, LM_Ins1_L21_ins_mesh, LM_Ins1_L21_psw17, LM_Ins1_L21_psw18, LM_Ins1_L21_psw19, LM_Ins1_L21_psw20, LM_Ins1_L21_psw21, LM_Ins1_L21_psw22, LM_Ins1_L21_psw23, LM_Ins1_L21_tsw25)
Dim BL_Ins1_L22: BL_Ins1_L22=Array(LM_Ins1_L22_Playfield, LM_Ins1_L22_ins_mesh, LM_Ins1_L22_psw17, LM_Ins1_L22_psw18, LM_Ins1_L22_psw19, LM_Ins1_L22_psw20, LM_Ins1_L22_psw21, LM_Ins1_L22_psw22, LM_Ins1_L22_psw23, LM_Ins1_L22_tsw30)
Dim BL_Ins1_L23: BL_Ins1_L23=Array(LM_Ins1_L23_Playfield, LM_Ins1_L23_ins_mesh, LM_Ins1_L23_psw17, LM_Ins1_L23_psw18, LM_Ins1_L23_psw19, LM_Ins1_L23_psw20, LM_Ins1_L23_psw21, LM_Ins1_L23_psw22, LM_Ins1_L23_psw23)
Dim BL_Ins1_L24: BL_Ins1_L24=Array(LM_Ins1_L24_Playfield, LM_Ins1_L24_ins_mesh, LM_Ins1_L24_psw17, LM_Ins1_L24_psw18, LM_Ins1_L24_psw19, LM_Ins1_L24_psw20, LM_Ins1_L24_psw21, LM_Ins1_L24_psw22, LM_Ins1_L24_psw23)
Dim BL_Ins1_L26: BL_Ins1_L26=Array(LM_Ins1_L26_Layer_2, LM_Ins1_L26_Parts, LM_Ins1_L26_Playfield, LM_Ins1_L26_ins_mesh, LM_Ins1_L26_plastics)
Dim BL_Ins1_L28: BL_Ins1_L28=Array(LM_Ins1_L28_Flip1, LM_Ins1_L28_Flip2, LM_Ins1_L28_Playfield, LM_Ins1_L28_ins_mesh)
Dim BL_Ins1_L3: BL_Ins1_L3=Array(LM_Ins1_L3_Playfield, LM_Ins1_L3_ins_mesh, LM_Ins1_L3_psw21, LM_Ins1_L3_psw22, LM_Ins1_L3_psw23)
Dim BL_Ins1_L33: BL_Ins1_L33=Array(LM_Ins1_L33_Flip1, LM_Ins1_L33_Playfield)
Dim BL_Ins1_L34: BL_Ins1_L34=Array(LM_Ins1_L34_Flip3, LM_Ins1_L34_Playfield, LM_Ins1_L34_sw14)
Dim BL_Ins1_L35: BL_Ins1_L35=Array(LM_Ins1_L35_Flip3, LM_Ins1_L35_Playfield)
Dim BL_Ins1_L36: BL_Ins1_L36=Array(LM_Ins1_L36_Playfield, LM_Ins1_L36_ins_mesh, LM_Ins1_L36_psw19, LM_Ins1_L36_psw20, LM_Ins1_L36_psw21, LM_Ins1_L36_psw22, LM_Ins1_L36_psw23)
Dim BL_Ins1_L37: BL_Ins1_L37=Array(LM_Ins1_L37_Playfield, LM_Ins1_L37_ins_mesh, LM_Ins1_L37_psw17, LM_Ins1_L37_psw18, LM_Ins1_L37_psw19, LM_Ins1_L37_psw20, LM_Ins1_L37_psw21, LM_Ins1_L37_psw22, LM_Ins1_L37_psw23, LM_Ins1_L37_tsw27)
Dim BL_Ins1_L38: BL_Ins1_L38=Array(LM_Ins1_L38_Playfield, LM_Ins1_L38_ins_mesh, LM_Ins1_L38_psw17, LM_Ins1_L38_psw18, LM_Ins1_L38_psw19, LM_Ins1_L38_psw20, LM_Ins1_L38_psw21, LM_Ins1_L38_psw22, LM_Ins1_L38_psw23)
Dim BL_Ins1_L39: BL_Ins1_L39=Array(LM_Ins1_L39_Playfield, LM_Ins1_L39_ins_mesh, LM_Ins1_L39_psw17, LM_Ins1_L39_psw18, LM_Ins1_L39_psw19, LM_Ins1_L39_psw20, LM_Ins1_L39_psw21, LM_Ins1_L39_psw22, LM_Ins1_L39_psw23)
Dim BL_Ins1_L4: BL_Ins1_L4=Array(LM_Ins1_L4_Flip3, LM_Ins1_L4_Playfield, LM_Ins1_L4_ins_mesh)
Dim BL_Ins1_L40: BL_Ins1_L40=Array(LM_Ins1_L40_Playfield, LM_Ins1_L40_ins_mesh, LM_Ins1_L40_psw18, LM_Ins1_L40_psw19, LM_Ins1_L40_psw20, LM_Ins1_L40_psw21, LM_Ins1_L40_psw22, LM_Ins1_L40_psw23)
Dim BL_Ins1_L42: BL_Ins1_L42=Array(LM_Ins1_L42_Parts, LM_Ins1_L42_Playfield, LM_Ins1_L42_ins_mesh, LM_Ins1_L42_psw20, LM_Ins1_L42_psw23)
Dim BL_Ins1_L49: BL_Ins1_L49=Array(LM_Ins1_L49_Flip3, LM_Ins1_L49_Playfield)
Dim BL_Ins1_L5: BL_Ins1_L5=Array(LM_Ins1_L5_Parts, LM_Ins1_L5_Playfield, LM_Ins1_L5_ins_mesh, LM_Ins1_L5_psw17, LM_Ins1_L5_psw18, LM_Ins1_L5_psw19, LM_Ins1_L5_psw20, LM_Ins1_L5_psw21, LM_Ins1_L5_psw22, LM_Ins1_L5_psw23, LM_Ins1_L5_tsw25)
Dim BL_Ins1_L50: BL_Ins1_L50=Array(LM_Ins1_L50_Flip3, LM_Ins1_L50_Playfield, LM_Ins1_L50_ins_mesh)
Dim BL_Ins1_L51: BL_Ins1_L51=Array(LM_Ins1_L51_Flip3, LM_Ins1_L51_Playfield, LM_Ins1_L51_ins_mesh)
Dim BL_Ins1_L53: BL_Ins1_L53=Array(LM_Ins1_L53_Playfield, LM_Ins1_L53_ins_mesh, LM_Ins1_L53_psw17, LM_Ins1_L53_psw18, LM_Ins1_L53_psw19, LM_Ins1_L53_psw20, LM_Ins1_L53_psw21, LM_Ins1_L53_psw22, LM_Ins1_L53_psw23, LM_Ins1_L53_tsw28)
Dim BL_Ins1_L54: BL_Ins1_L54=Array(LM_Ins1_L54_Parts, LM_Ins1_L54_Playfield, LM_Ins1_L54_plastics, LM_Ins1_L54_psw33)
Dim BL_Ins1_L55: BL_Ins1_L55=Array(LM_Ins1_L55_Playfield, LM_Ins1_L55_ins_mesh, LM_Ins1_L55_psw17, LM_Ins1_L55_psw18, LM_Ins1_L55_psw19, LM_Ins1_L55_psw20, LM_Ins1_L55_psw21, LM_Ins1_L55_psw22, LM_Ins1_L55_psw23)
Dim BL_Ins1_L58: BL_Ins1_L58=Array(LM_Ins1_L58_Flip3, LM_Ins1_L58_Parts, LM_Ins1_L58_Playfield, LM_Ins1_L58_ins_mesh)
Dim BL_Ins1_L59: BL_Ins1_L59=Array(LM_Ins1_L59_Parts, LM_Ins1_L59_Playfield)
Dim BL_Ins1_L6: BL_Ins1_L6=Array(LM_Ins1_L6_Playfield, LM_Ins1_L6_ins_mesh, LM_Ins1_L6_psw17, LM_Ins1_L6_psw18, LM_Ins1_L6_psw19, LM_Ins1_L6_psw20, LM_Ins1_L6_psw21, LM_Ins1_L6_psw22, LM_Ins1_L6_psw23)
Dim BL_Ins1_L60: BL_Ins1_L60=Array(LM_Ins1_L60_Flip1, LM_Ins1_L60_Flip2, LM_Ins1_L60_Parts, LM_Ins1_L60_Playfield, LM_Ins1_L60_ins_mesh)
Dim BL_Ins1_L63: BL_Ins1_L63=Array(LM_Ins1_L63_Flip3, LM_Ins1_L63_Playfield, LM_Ins1_L63_ins_mesh, LM_Ins1_L63_psw18, LM_Ins1_L63_psw2, LM_Ins1_L63_psw21, LM_Ins1_L63_psw22, LM_Ins1_L63_psw23, LM_Ins1_L63_psw3)
Dim BL_Ins1_L65: BL_Ins1_L65=Array(LM_Ins1_L65_Flip3, LM_Ins1_L65_Playfield, LM_Ins1_L65_ins_mesh, LM_Ins1_L65_psw1)
Dim BL_Ins1_L7: BL_Ins1_L7=Array(LM_Ins1_L7_BumperSkirt3, LM_Ins1_L7_Layer_2, LM_Ins1_L7_Playfield, LM_Ins1_L7_ins_mesh, LM_Ins1_L7_psw17, LM_Ins1_L7_psw18, LM_Ins1_L7_psw19, LM_Ins1_L7_psw20, LM_Ins1_L7_psw21, LM_Ins1_L7_psw22)
Dim BL_Ins1_L70: BL_Ins1_L70=Array(LM_Ins1_L70_BumperSkirt1, LM_Ins1_L70_Layer_2, LM_Ins1_L70_Parts, LM_Ins1_L70_Playfield, LM_Ins1_L70_bumpcapA_3_002, LM_Ins1_L70_plastics)
Dim BL_Ins1_L8: BL_Ins1_L8=Array(LM_Ins1_L8_Playfield, LM_Ins1_L8_ins_mesh, LM_Ins1_L8_psw17, LM_Ins1_L8_psw18, LM_Ins1_L8_psw19, LM_Ins1_L8_psw20, LM_Ins1_L8_psw21, LM_Ins1_L8_psw22, LM_Ins1_L8_psw23)
Dim BL_Ins1_L81: BL_Ins1_L81=Array(LM_Ins1_L81_Flip3, LM_Ins1_L81_Parts, LM_Ins1_L81_Playfield, LM_Ins1_L81_ins_mesh, LM_Ins1_L81_psw1, LM_Ins1_L81_psw2)
Dim BL_Ins1_L82: BL_Ins1_L82=Array(LM_Ins1_L82_Parts, LM_Ins1_L82_Playfield, LM_Ins1_L82_bulbs, LM_Ins1_L82_plastics, LM_Ins1_L82_psw1, LM_Ins1_L82_psw2, LM_Ins1_L82_psw3, LM_Ins1_L82_psw4)
Dim BL_Ins1_L86: BL_Ins1_L86=Array(LM_Ins1_L86_BumperSkirt2, LM_Ins1_L86_Layer_2, LM_Ins1_L86_Parts, LM_Ins1_L86_Playfield, LM_Ins1_L86_bumpcapA_3_002)
Dim BL_Ins1_L97: BL_Ins1_L97=Array(LM_Ins1_L97_Flip3, LM_Ins1_L97_Parts, LM_Ins1_L97_Playfield, LM_Ins1_L97_bulbs, LM_Ins1_L97_ins_mesh, LM_Ins1_L97_psw1, LM_Ins1_L97_psw2)
Dim BL_Ins2_L100: BL_Ins2_L100=Array(LM_Ins2_L100_Playfield, LM_Ins2_L100_bulbs, LM_Ins2_L100_ins_mesh)
Dim BL_Ins2_L11: BL_Ins2_L11=Array(LM_Ins2_L11_Flip1, LM_Ins2_L11_Flip2, LM_Ins2_L11_bulbs, LM_Ins2_L11_ins_mesh)
Dim BL_Ins2_L113: BL_Ins2_L113=Array(LM_Ins2_L113_Parts, LM_Ins2_L113_Playfield, LM_Ins2_L113_bulbs, LM_Ins2_L113_ins_mesh, LM_Ins2_L113_psw1)
Dim BL_Ins2_L116: BL_Ins2_L116=Array(LM_Ins2_L116_Playfield, LM_Ins2_L116_bulbs, LM_Ins2_L116_ins_mesh)
Dim BL_Ins2_L12: BL_Ins2_L12=Array(LM_Ins2_L12_Flip2, LM_Ins2_L12_Playfield, LM_Ins2_L12_bulbs, LM_Ins2_L12_ins_mesh)
Dim BL_Ins2_L25: BL_Ins2_L25=Array(LM_Ins2_L25_Parts, LM_Ins2_L25_Playfield, LM_Ins2_L25_bulbs, LM_Ins2_L25_ins_mesh)
Dim BL_Ins2_L41: BL_Ins2_L41=Array(LM_Ins2_L41_Flip3, LM_Ins2_L41_Playfield, LM_Ins2_L41_bulbs, LM_Ins2_L41_ins_mesh, LM_Ins2_L41_plastics)
Dim BL_Ins2_L44: BL_Ins2_L44=Array(LM_Ins2_L44_Flip1, LM_Ins2_L44_Playfield, LM_Ins2_L44_bulbs, LM_Ins2_L44_ins_mesh)
Dim BL_Ins2_L47: BL_Ins2_L47=Array(LM_Ins2_L47_Playfield, LM_Ins2_L47_bulbs, LM_Ins2_L47_ins_mesh)
Dim BL_Ins2_L57: BL_Ins2_L57=Array(LM_Ins2_L57_Playfield, LM_Ins2_L57_bulbs, LM_Ins2_L57_ins_mesh, LM_Ins2_L57_plastics)
Dim BL_Ins2_L66: BL_Ins2_L66=Array(LM_Ins2_L66_Parts, LM_Ins2_L66_bulbs, LM_Ins2_L66_ins_mesh, LM_Ins2_L66_plastics, LM_Ins2_L66_psw1)
Dim BL_Ins2_L67: BL_Ins2_L67=Array(LM_Ins2_L67_Playfield, LM_Ins2_L67_bulbs, LM_Ins2_L67_ins_mesh, LM_Ins2_L67_psw17, LM_Ins2_L67_psw18, LM_Ins2_L67_psw19)
Dim BL_Ins2_L68: BL_Ins2_L68=Array(LM_Ins2_L68_Layer_2, LM_Ins2_L68_Playfield, LM_Ins2_L68_bulbs, LM_Ins2_L68_ins_mesh)
Dim BL_Ins2_L69: BL_Ins2_L69=Array(LM_Ins2_L69_Playfield, LM_Ins2_L69_bulbs, LM_Ins2_L69_ins_mesh)
Dim BL_Ins2_L71: BL_Ins2_L71=Array(LM_Ins2_L71_Flip2, LM_Ins2_L71_Playfield, LM_Ins2_L71_bulbs, LM_Ins2_L71_ins_mesh)
Dim BL_Ins2_L83: BL_Ins2_L83=Array(LM_Ins2_L83_Playfield, LM_Ins2_L83_bulbs, LM_Ins2_L83_ins_mesh, LM_Ins2_L83_psw17, LM_Ins2_L83_psw18, LM_Ins2_L83_psw19)
Dim BL_Ins2_L84: BL_Ins2_L84=Array(LM_Ins2_L84_Playfield, LM_Ins2_L84_bulbs, LM_Ins2_L84_ins_mesh)
Dim BL_Ins2_L85: BL_Ins2_L85=Array(LM_Ins2_L85_Playfield, LM_Ins2_L85_bulbs, LM_Ins2_L85_ins_mesh)
Dim BL_Ins2_L87: BL_Ins2_L87=Array(LM_Ins2_L87_Flip1, LM_Ins2_L87_Playfield, LM_Ins2_L87_bulbs, LM_Ins2_L87_ins_mesh)
Dim BL_Ins2_L9: BL_Ins2_L9=Array(LM_Ins2_L9_Parts, LM_Ins2_L9_Playfield, LM_Ins2_L9_bulbs, LM_Ins2_L9_ins_mesh)
Dim BL_Ins2_L98: BL_Ins2_L98=Array(LM_Ins2_L98_bulbs, LM_Ins2_L98_ins_mesh)
Dim BL_Ins2_L99: BL_Ins2_L99=Array(LM_Ins2_L99_Playfield, LM_Ins2_L99_bulbs, LM_Ins2_L99_ins_mesh, LM_Ins2_L99_psw17)
Dim BL_World: BL_World=Array(BM_BumperRing1, BM_BumperRing2, BM_BumperRing3, BM_BumperSkirt1, BM_BumperSkirt2, BM_BumperSkirt3, BM_Flip1, BM_Flip2, BM_Flip3, BM_Gateplate2, BM_Gateplate5, BM_Lane_Guide___2a, BM_Layer_2, BM_Parts, BM_Playfield, BM_ROstar, BM_bulbs, BM_bumpcapA_3_002, BM_gatewire3, BM_gatewire4, BM_ins_mesh, BM_lsling, BM_lsling1, BM_lsling2, BM_plastics, BM_psw1, BM_psw17, BM_psw18, BM_psw19, BM_psw2, BM_psw20, BM_psw21, BM_psw22, BM_psw23, BM_psw3, BM_psw33, BM_psw4, BM_rails, BM_rsling, BM_rsling1, BM_rsling2, BM_slingarmL, BM_slingarmR, BM_sw12, BM_sw13, BM_sw14, BM_sw15, BM_sw31, BM_sw32, BM_tsw25, BM_tsw26, BM_tsw27, BM_tsw28, BM_tsw29, BM_tsw30, BM_tsw5)
' Global arrays
Dim BG_Bakemap: BG_Bakemap=Array(BM_BumperRing1, BM_BumperRing2, BM_BumperRing3, BM_BumperSkirt1, BM_BumperSkirt2, BM_BumperSkirt3, BM_Flip1, BM_Flip2, BM_Flip3, BM_Gateplate2, BM_Gateplate5, BM_Lane_Guide___2a, BM_Layer_2, BM_Parts, BM_Playfield, BM_ROstar, BM_bulbs, BM_bumpcapA_3_002, BM_gatewire3, BM_gatewire4, BM_ins_mesh, BM_lsling, BM_lsling1, BM_lsling2, BM_plastics, BM_psw1, BM_psw17, BM_psw18, BM_psw19, BM_psw2, BM_psw20, BM_psw21, BM_psw22, BM_psw23, BM_psw3, BM_psw33, BM_psw4, BM_rails, BM_rsling, BM_rsling1, BM_rsling2, BM_slingarmL, BM_slingarmR, BM_sw12, BM_sw13, BM_sw14, BM_sw15, BM_sw31, BM_sw32, BM_tsw25, BM_tsw26, BM_tsw27, BM_tsw28, BM_tsw29, BM_tsw30, BM_tsw5)
Dim BG_Lightmap: BG_Lightmap=Array(LM_GI_GI_1_Layer_2, LM_GI_GI_1_Parts, LM_GI_GI_1_Playfield, LM_GI_GI_1_bulbs, LM_GI_GI_1_plastics, LM_GI_GI_1_psw17, LM_GI_GI_1_psw33, LM_GI_GI_10_Flip1, LM_GI_GI_10_Flip2, LM_GI_GI_10_Layer_2, LM_GI_GI_10_Parts, LM_GI_GI_10_Playfield, LM_GI_GI_10_bulbs, LM_GI_GI_10_lsling, LM_GI_GI_10_lsling1, LM_GI_GI_10_lsling2, LM_GI_GI_10_plastics, LM_GI_GI_10_psw21, LM_GI_GI_10_psw22, LM_GI_GI_10_psw23, LM_GI_GI_10_rsling, LM_GI_GI_10_rsling1, LM_GI_GI_10_rsling2, LM_GI_GI_11_Flip1, LM_GI_GI_11_Flip2, LM_GI_GI_11_Flip3, LM_GI_GI_11_Layer_2, LM_GI_GI_11_Parts, LM_GI_GI_11_Playfield, LM_GI_GI_11_bulbs, LM_GI_GI_11_ins_mesh, LM_GI_GI_11_lsling, LM_GI_GI_11_lsling1, LM_GI_GI_11_lsling2, LM_GI_GI_11_plastics, LM_GI_GI_11_psw19, LM_GI_GI_11_psw20, LM_GI_GI_11_psw21, LM_GI_GI_11_psw22, LM_GI_GI_11_psw23, LM_GI_GI_11_rails, LM_GI_GI_11_rsling, LM_GI_GI_11_rsling1, LM_GI_GI_11_rsling2, LM_GI_GI_11_slingarmL, LM_GI_GI_11_sw14, LM_GI_GI_11_sw32, LM_GI_GI_13_Flip1, LM_GI_GI_13_Flip2, _
  LM_GI_GI_13_Flip3, LM_GI_GI_13_Layer_2, LM_GI_GI_13_Parts, LM_GI_GI_13_Playfield, LM_GI_GI_13_bulbs, LM_GI_GI_13_ins_mesh, LM_GI_GI_13_lsling, LM_GI_GI_13_lsling1, LM_GI_GI_13_lsling2, LM_GI_GI_13_plastics, LM_GI_GI_13_psw1, LM_GI_GI_13_psw2, LM_GI_GI_13_psw20, LM_GI_GI_13_psw21, LM_GI_GI_13_psw22, LM_GI_GI_13_psw23, LM_GI_GI_13_rails, LM_GI_GI_13_rsling, LM_GI_GI_13_rsling1, LM_GI_GI_13_rsling2, LM_GI_GI_13_slingarmR, LM_GI_GI_13_sw15, LM_GI_GI_13_sw31, LM_GI_GI_13_tsw30, LM_GI_GI_14_Flip1, LM_GI_GI_14_Flip2, LM_GI_GI_14_Flip3, LM_GI_GI_14_Layer_2, LM_GI_GI_14_Parts, LM_GI_GI_14_Playfield, LM_GI_GI_14_bulbs, LM_GI_GI_14_lsling, LM_GI_GI_14_lsling1, LM_GI_GI_14_lsling2, LM_GI_GI_14_plastics, LM_GI_GI_14_psw18, LM_GI_GI_14_psw19, LM_GI_GI_14_psw20, LM_GI_GI_14_psw21, LM_GI_GI_14_psw22, LM_GI_GI_14_rsling, LM_GI_GI_14_rsling1, LM_GI_GI_14_rsling2, LM_GI_GI_14_sw14, LM_GI_GI_15_Flip1, LM_GI_GI_15_Layer_2, LM_GI_GI_15_Parts, LM_GI_GI_15_Playfield, LM_GI_GI_15_bulbs, LM_GI_GI_15_lsling, LM_GI_GI_15_lsling1, _
  LM_GI_GI_15_lsling2, LM_GI_GI_15_plastics, LM_GI_GI_15_psw1, LM_GI_GI_15_psw18, LM_GI_GI_15_psw2, LM_GI_GI_15_psw20, LM_GI_GI_15_psw23, LM_GI_GI_15_psw3, LM_GI_GI_15_rails, LM_GI_GI_15_rsling, LM_GI_GI_15_rsling1, LM_GI_GI_15_rsling2, LM_GI_GI_15_slingarmR, LM_GI_GI_15_sw32, LM_GI_GI_17_BumperSkirt1, LM_GI_GI_17_BumperSkirt2, LM_GI_GI_17_Layer_2, LM_GI_GI_17_Parts, LM_GI_GI_17_Playfield, LM_GI_GI_17_bulbs, LM_GI_GI_17_plastics, LM_GI_GI_17_psw33, LM_GI_GI_17_rails, LM_GI_GI_18_BumperRing1, LM_GI_GI_18_BumperSkirt1, LM_GI_GI_18_BumperSkirt2, LM_GI_GI_18_Layer_2, LM_GI_GI_18_Parts, LM_GI_GI_18_Playfield, LM_GI_GI_18_ROstar, LM_GI_GI_18_bulbs, LM_GI_GI_18_plastics, LM_GI_GI_18_psw1, LM_GI_GI_18_psw2, LM_GI_GI_18_psw3, LM_GI_GI_18_psw4, LM_GI_GI_18_tsw5, LM_GI_GI_19_BumperSkirt2, LM_GI_GI_19_Layer_2, LM_GI_GI_19_Parts, LM_GI_GI_19_Playfield, LM_GI_GI_19_bulbs, LM_GI_GI_19_plastics, LM_GI_GI_19_psw33, LM_GI_GI_20_BumperRing2, LM_GI_GI_20_BumperRing3, LM_GI_GI_20_BumperSkirt1, LM_GI_GI_20_BumperSkirt2, _
  LM_GI_GI_20_BumperSkirt3, LM_GI_GI_20_Layer_2, LM_GI_GI_20_Parts, LM_GI_GI_20_Playfield, LM_GI_GI_20_ROstar, LM_GI_GI_20_bulbs, LM_GI_GI_20_plastics, LM_GI_GI_20_psw1, LM_GI_GI_20_psw17, LM_GI_GI_20_psw18, LM_GI_GI_20_psw19, LM_GI_GI_20_psw2, LM_GI_GI_20_psw22, LM_GI_GI_20_psw3, LM_GI_GI_20_psw4, LM_GI_GI_20_rails, LM_GI_GI_20_tsw25, LM_GI_GI_21_BumperRing1, LM_GI_GI_21_BumperRing2, LM_GI_GI_21_BumperSkirt1, LM_GI_GI_21_BumperSkirt2, LM_GI_GI_21_BumperSkirt3, LM_GI_GI_21_Layer_2, LM_GI_GI_21_Parts, LM_GI_GI_21_Playfield, LM_GI_GI_21_bulbs, LM_GI_GI_21_bumpcapA_3_002, LM_GI_GI_21_ins_mesh, LM_GI_GI_21_plastics, LM_GI_GI_21_psw17, LM_GI_GI_21_psw18, LM_GI_GI_21_psw19, LM_GI_GI_21_psw2, LM_GI_GI_21_psw20, LM_GI_GI_21_psw21, LM_GI_GI_21_psw22, LM_GI_GI_21_psw23, LM_GI_GI_21_psw33, LM_GI_GI_21_rails, LM_GI_GI_21_tsw25, LM_GI_GI_21_tsw26, LM_GI_GI_21_tsw27, LM_GI_GI_21_tsw28, LM_GI_GI_21_tsw29, LM_GI_GI_22_BumperSkirt1, LM_GI_GI_22_BumperSkirt2, LM_GI_GI_22_Gateplate2, LM_GI_GI_22_Lane_Guide___2a, _
  LM_GI_GI_22_Layer_2, LM_GI_GI_22_Parts, LM_GI_GI_22_Playfield, LM_GI_GI_22_ROstar, LM_GI_GI_22_bulbs, LM_GI_GI_22_plastics, LM_GI_GI_22_psw19, LM_GI_GI_22_psw20, LM_GI_GI_22_psw21, LM_GI_GI_22_psw22, LM_GI_GI_22_psw23, LM_GI_GI_22_psw4, LM_GI_GI_22_sw12, LM_GI_GI_22_sw13, LM_GI_GI_22_tsw5, LM_GI_GI_23_BumperSkirt1, LM_GI_GI_23_BumperSkirt2, LM_GI_GI_23_Lane_Guide___2a, LM_GI_GI_23_Layer_2, LM_GI_GI_23_Parts, LM_GI_GI_23_Playfield, LM_GI_GI_23_bulbs, LM_GI_GI_23_psw1, LM_GI_GI_23_psw19, LM_GI_GI_23_psw2, LM_GI_GI_23_psw3, LM_GI_GI_23_psw4, LM_GI_GI_23_sw12, LM_GI_GI_23_sw13, LM_GI_GI_23_tsw5, LM_GI_GI_24_BumperSkirt1, LM_GI_GI_24_BumperSkirt2, LM_GI_GI_24_Lane_Guide___2a, LM_GI_GI_24_Layer_2, LM_GI_GI_24_Parts, LM_GI_GI_24_Playfield, LM_GI_GI_24_bulbs, LM_GI_GI_24_plastics, LM_GI_GI_24_psw17, LM_GI_GI_24_psw19, LM_GI_GI_24_psw2, LM_GI_GI_24_psw20, LM_GI_GI_24_psw21, LM_GI_GI_24_psw22, LM_GI_GI_24_psw23, LM_GI_GI_24_psw3, LM_GI_GI_24_psw4, LM_GI_GI_24_sw12, LM_GI_GI_24_sw13, LM_GI_GI_25_BumperSkirt1, _
  LM_GI_GI_25_BumperSkirt2, LM_GI_GI_25_Layer_2, LM_GI_GI_25_Parts, LM_GI_GI_25_Playfield, LM_GI_GI_25_bulbs, LM_GI_GI_25_plastics, LM_GI_GI_25_psw17, LM_GI_GI_25_psw18, LM_GI_GI_25_psw19, LM_GI_GI_25_psw20, LM_GI_GI_25_psw21, LM_GI_GI_25_psw33, LM_GI_GI_25_rails, LM_GI_GI_4_BumperRing3, LM_GI_GI_4_BumperSkirt3, LM_GI_GI_4_Flip1, LM_GI_GI_4_Flip3, LM_GI_GI_4_Parts, LM_GI_GI_4_Playfield, LM_GI_GI_4_bulbs, LM_GI_GI_4_plastics, LM_GI_GI_4_psw1, LM_GI_GI_4_psw17, LM_GI_GI_4_psw18, LM_GI_GI_4_psw19, LM_GI_GI_4_psw20, LM_GI_GI_4_psw21, LM_GI_GI_4_psw22, LM_GI_GI_4_psw23, LM_GI_GI_4_rails, LM_GI_GI_4_rsling, LM_GI_GI_4_rsling1, LM_GI_GI_4_rsling2, LM_GI_GI_4_tsw26, LM_GI_GI_4_tsw27, LM_GI_GI_4_tsw28, LM_GI_GI_4_tsw29, LM_GI_GI_4_tsw30, LM_GI_GI_6_Layer_2, LM_GI_GI_6_Parts, LM_GI_GI_6_Playfield, LM_GI_GI_6_bulbs, LM_GI_GI_6_plastics, LM_GI_GI_6_psw33, LM_GI_GI_8_Flip1, LM_GI_GI_8_Flip2, LM_GI_GI_8_Flip3, LM_GI_GI_8_Parts, LM_GI_GI_8_Playfield, LM_GI_GI_8_bulbs, LM_GI_GI_8_ins_mesh, LM_GI_GI_8_lsling, _
  LM_GI_GI_8_lsling1, LM_GI_GI_8_lsling2, LM_GI_GI_8_plastics, LM_GI_GI_8_psw1, LM_GI_GI_8_psw17, LM_GI_GI_8_psw18, LM_GI_GI_8_psw19, LM_GI_GI_8_psw2, LM_GI_GI_8_psw20, LM_GI_GI_8_psw21, LM_GI_GI_8_psw22, LM_GI_GI_8_psw23, LM_GI_GI_8_rails, LM_GI_GI_8_sw14, LM_GI_GI_9_Flip1, LM_GI_GI_9_Flip2, LM_GI_GI_9_Flip3, LM_GI_GI_9_Layer_2, LM_GI_GI_9_Parts, LM_GI_GI_9_Playfield, LM_GI_GI_9_bulbs, LM_GI_GI_9_lsling, LM_GI_GI_9_lsling1, LM_GI_GI_9_lsling2, LM_GI_GI_9_plastics, LM_GI_GI_9_psw1, LM_GI_GI_9_psw17, LM_GI_GI_9_psw18, LM_GI_GI_9_psw19, LM_GI_GI_9_psw2, LM_GI_GI_9_psw20, LM_GI_GI_9_psw21, LM_GI_GI_9_rails, LM_GI_GI_9_rsling, LM_GI_GI_9_rsling1, LM_GI_GI_9_rsling2, LM_GI_GI_9_slingarmL, LM_GI_GI_9_sw31, LM_GI_gi1_BumperSkirt2, LM_GI_gi1_Flip1, LM_GI_gi1_Flip2, LM_GI_gi1_Flip3, LM_GI_gi1_Gateplate5, LM_GI_gi1_Layer_2, LM_GI_gi1_Parts, LM_GI_gi1_Playfield, LM_GI_gi1_ROstar, LM_GI_gi1_bulbs, LM_GI_gi1_gatewire4, LM_GI_gi1_lsling, LM_GI_gi1_lsling1, LM_GI_gi1_lsling2, LM_GI_gi1_plastics, LM_GI_gi1_psw1, _
  LM_GI_gi1_psw17, LM_GI_gi1_psw18, LM_GI_gi1_psw19, LM_GI_gi1_psw2, LM_GI_gi1_psw20, LM_GI_gi1_psw21, LM_GI_gi1_psw22, LM_GI_gi1_psw23, LM_GI_gi1_psw3, LM_GI_gi1_psw33, LM_GI_gi1_psw4, LM_GI_gi1_rails, LM_GI_gi1_tsw25, LM_GI_gi1_tsw26, LM_GI_gi1_tsw27, LM_GI_gi1_tsw28, LM_GI_gi1_tsw29, LM_GI_gi1_tsw30, LM_GI_gi1_tsw5, LM_Ins1_L1_Flip1, LM_Ins1_L1_Flip2, LM_Ins1_L1_Playfield, LM_Ins1_L10_Parts, LM_Ins1_L10_Playfield, LM_Ins1_L10_ins_mesh, LM_Ins1_L10_plastics, LM_Ins1_L101_BumperRing3, LM_Ins1_L101_BumperSkirt1, LM_Ins1_L101_BumperSkirt2, LM_Ins1_L101_BumperSkirt3, LM_Ins1_L101_Flip3, LM_Ins1_L101_Layer_2, LM_Ins1_L101_Parts, LM_Ins1_L101_Playfield, LM_Ins1_L101_bulbs, LM_Ins1_L101_plastics, LM_Ins1_L101_psw1, LM_Ins1_L101_psw2, LM_Ins1_L101_psw20, LM_Ins1_L101_psw21, LM_Ins1_L101_psw22, LM_Ins1_L101_psw23, LM_Ins1_L101_psw3, LM_Ins1_L101_psw4, LM_Ins1_L101_tsw5, LM_Ins1_L101a_BumperSkirt2, LM_Ins1_L101a_BumperSkirt3, LM_Ins1_L101a_Layer_2, LM_Ins1_L101a_Parts, LM_Ins1_L101a_Playfield, LM_Ins1_L101a_bulbs, _
  LM_Ins1_L101a_ins_mesh, LM_Ins1_L101a_plastics, LM_Ins1_L101a_psw17, LM_Ins1_L101a_psw18, LM_Ins1_L101a_psw33, LM_Ins1_L101a_rails, LM_Ins1_L101a_tsw27, LM_Ins1_L102_BumperSkirt3, LM_Ins1_L102_Layer_2, LM_Ins1_L102_Parts, LM_Ins1_L102_Playfield, LM_Ins1_L102_bumpcapA_3_002, LM_Ins1_L103_Flip1, LM_Ins1_L103_Flip2, LM_Ins1_L103_Flip3, LM_Ins1_L103_Playfield, LM_Ins1_L103_ins_mesh, LM_Ins1_L114_Flip3, LM_Ins1_L114_Parts, LM_Ins1_L114_Playfield, LM_Ins1_L114_ins_mesh, LM_Ins1_L114_psw1, LM_Ins1_L114_psw2, LM_Ins1_L114_psw3, LM_Ins1_L115_Layer_2, LM_Ins1_L115_Parts, LM_Ins1_L115_Playfield, LM_Ins1_L115_bulbs, LM_Ins1_L115_plastics, LM_Ins1_L115_psw17, LM_Ins1_L115_psw33, LM_Ins1_L117_BumperRing1, LM_Ins1_L117_BumperSkirt1, LM_Ins1_L117_BumperSkirt3, LM_Ins1_L117_Layer_2, LM_Ins1_L117_Parts, LM_Ins1_L117_Playfield, LM_Ins1_L117_ROstar, LM_Ins1_L117_bulbs, LM_Ins1_L117_plastics, LM_Ins1_L117_psw1, LM_Ins1_L117_psw17, LM_Ins1_L117_psw2, LM_Ins1_L117_psw23, LM_Ins1_L117_psw3, LM_Ins1_L117_psw33, LM_Ins1_L117_psw4, _
  LM_Ins1_L117_tsw5, LM_Ins1_L117a_BumperSkirt2, LM_Ins1_L117a_Layer_2, LM_Ins1_L117a_Parts, LM_Ins1_L117a_Playfield, LM_Ins1_L117a_bulbs, LM_Ins1_L117a_plastics, LM_Ins1_L117a_psw17, LM_Ins1_L117a_psw18, LM_Ins1_L117a_psw33, LM_Ins1_L118_BumperSkirt1, LM_Ins1_L118_Gateplate2, LM_Ins1_L118_Layer_2, LM_Ins1_L118_Parts, LM_Ins1_L118_Playfield, LM_Ins1_L118_ROstar, LM_Ins1_L118_bulbs, LM_Ins1_L118_gatewire4, LM_Ins1_L118_plastics, LM_Ins1_L118_psw1, LM_Ins1_L118_psw17, LM_Ins1_L118_psw2, LM_Ins1_L118_psw3, LM_Ins1_L118_psw4, LM_Ins1_L118_rails, LM_Ins1_L118_tsw5, LM_Ins1_L118a_Layer_2, LM_Ins1_L118a_Parts, LM_Ins1_L118a_Playfield, LM_Ins1_L118a_bulbs, LM_Ins1_L118a_plastics, LM_Ins1_L118a_psw17, LM_Ins1_L118a_psw18, LM_Ins1_L118a_psw3, LM_Ins1_L118a_psw33, LM_Ins1_L118a_rails, LM_Ins1_L118a_tsw25, LM_Ins1_L119_BumperSkirt1, LM_Ins1_L119_Gateplate5, LM_Ins1_L119_Layer_2, LM_Ins1_L119_Parts, LM_Ins1_L119_Playfield, LM_Ins1_L119_ROstar, LM_Ins1_L119_bulbs, LM_Ins1_L119_plastics, LM_Ins1_L119_psw1, LM_Ins1_L119_psw2, _
  LM_Ins1_L119_psw3, LM_Ins1_L119_psw4, LM_Ins1_L119_rails, LM_Ins1_L119_tsw5, LM_Ins1_L119a_BumperSkirt2, LM_Ins1_L119a_Layer_2, LM_Ins1_L119a_Parts, LM_Ins1_L119a_Playfield, LM_Ins1_L119a_bulbs, LM_Ins1_L119a_gatewire3, LM_Ins1_L119a_plastics, LM_Ins1_L119a_psw33, LM_Ins1_L17_Flip2, LM_Ins1_L17_Flip3, LM_Ins1_L17_Parts, LM_Ins1_L17_Playfield, LM_Ins1_L18_Playfield, LM_Ins1_L18_psw22, LM_Ins1_L18_psw23, LM_Ins1_L19_Playfield, LM_Ins1_L19_ins_mesh, LM_Ins1_L19_psw21, LM_Ins1_L19_psw22, LM_Ins1_L19_psw23, LM_Ins1_L19_sw31, LM_Ins1_L2_Flip3, LM_Ins1_L2_Playfield, LM_Ins1_L20_Flip3, LM_Ins1_L20_Playfield, LM_Ins1_L20_ins_mesh, LM_Ins1_L20_psw21, LM_Ins1_L20_psw22, LM_Ins1_L20_psw23, LM_Ins1_L21_Playfield, LM_Ins1_L21_ins_mesh, LM_Ins1_L21_psw17, LM_Ins1_L21_psw18, LM_Ins1_L21_psw19, LM_Ins1_L21_psw20, LM_Ins1_L21_psw21, LM_Ins1_L21_psw22, LM_Ins1_L21_psw23, LM_Ins1_L21_tsw25, LM_Ins1_L22_Playfield, LM_Ins1_L22_ins_mesh, LM_Ins1_L22_psw17, LM_Ins1_L22_psw18, LM_Ins1_L22_psw19, LM_Ins1_L22_psw20, LM_Ins1_L22_psw21, _
  LM_Ins1_L22_psw22, LM_Ins1_L22_psw23, LM_Ins1_L22_tsw30, LM_Ins1_L23_Playfield, LM_Ins1_L23_ins_mesh, LM_Ins1_L23_psw17, LM_Ins1_L23_psw18, LM_Ins1_L23_psw19, LM_Ins1_L23_psw20, LM_Ins1_L23_psw21, LM_Ins1_L23_psw22, LM_Ins1_L23_psw23, LM_Ins1_L24_Playfield, LM_Ins1_L24_ins_mesh, LM_Ins1_L24_psw17, LM_Ins1_L24_psw18, LM_Ins1_L24_psw19, LM_Ins1_L24_psw20, LM_Ins1_L24_psw21, LM_Ins1_L24_psw22, LM_Ins1_L24_psw23, LM_Ins1_L26_Layer_2, LM_Ins1_L26_Parts, LM_Ins1_L26_Playfield, LM_Ins1_L26_ins_mesh, LM_Ins1_L26_plastics, LM_Ins1_L28_Flip1, LM_Ins1_L28_Flip2, LM_Ins1_L28_Playfield, LM_Ins1_L28_ins_mesh, LM_Ins1_L3_Playfield, LM_Ins1_L3_ins_mesh, LM_Ins1_L3_psw21, LM_Ins1_L3_psw22, LM_Ins1_L3_psw23, LM_Ins1_L33_Flip1, LM_Ins1_L33_Playfield, LM_Ins1_L34_Flip3, LM_Ins1_L34_Playfield, LM_Ins1_L34_sw14, LM_Ins1_L35_Flip3, LM_Ins1_L35_Playfield, LM_Ins1_L36_Playfield, LM_Ins1_L36_ins_mesh, LM_Ins1_L36_psw19, LM_Ins1_L36_psw20, LM_Ins1_L36_psw21, LM_Ins1_L36_psw22, LM_Ins1_L36_psw23, LM_Ins1_L37_Playfield, _
  LM_Ins1_L37_ins_mesh, LM_Ins1_L37_psw17, LM_Ins1_L37_psw18, LM_Ins1_L37_psw19, LM_Ins1_L37_psw20, LM_Ins1_L37_psw21, LM_Ins1_L37_psw22, LM_Ins1_L37_psw23, LM_Ins1_L37_tsw27, LM_Ins1_L38_Playfield, LM_Ins1_L38_ins_mesh, LM_Ins1_L38_psw17, LM_Ins1_L38_psw18, LM_Ins1_L38_psw19, LM_Ins1_L38_psw20, LM_Ins1_L38_psw21, LM_Ins1_L38_psw22, LM_Ins1_L38_psw23, LM_Ins1_L39_Playfield, LM_Ins1_L39_ins_mesh, LM_Ins1_L39_psw17, LM_Ins1_L39_psw18, LM_Ins1_L39_psw19, LM_Ins1_L39_psw20, LM_Ins1_L39_psw21, LM_Ins1_L39_psw22, LM_Ins1_L39_psw23, LM_Ins1_L4_Flip3, LM_Ins1_L4_Playfield, LM_Ins1_L4_ins_mesh, LM_Ins1_L40_Playfield, LM_Ins1_L40_ins_mesh, LM_Ins1_L40_psw18, LM_Ins1_L40_psw19, LM_Ins1_L40_psw20, LM_Ins1_L40_psw21, LM_Ins1_L40_psw22, LM_Ins1_L40_psw23, LM_Ins1_L42_Parts, LM_Ins1_L42_Playfield, LM_Ins1_L42_ins_mesh, LM_Ins1_L42_psw20, LM_Ins1_L42_psw23, LM_Ins1_L49_Flip3, LM_Ins1_L49_Playfield, LM_Ins1_L5_Parts, LM_Ins1_L5_Playfield, LM_Ins1_L5_ins_mesh, LM_Ins1_L5_psw17, LM_Ins1_L5_psw18, LM_Ins1_L5_psw19, _
  LM_Ins1_L5_psw20, LM_Ins1_L5_psw21, LM_Ins1_L5_psw22, LM_Ins1_L5_psw23, LM_Ins1_L5_tsw25, LM_Ins1_L50_Flip3, LM_Ins1_L50_Playfield, LM_Ins1_L50_ins_mesh, LM_Ins1_L51_Flip3, LM_Ins1_L51_Playfield, LM_Ins1_L51_ins_mesh, LM_Ins1_L53_Playfield, LM_Ins1_L53_ins_mesh, LM_Ins1_L53_psw17, LM_Ins1_L53_psw18, LM_Ins1_L53_psw19, LM_Ins1_L53_psw20, LM_Ins1_L53_psw21, LM_Ins1_L53_psw22, LM_Ins1_L53_psw23, LM_Ins1_L53_tsw28, LM_Ins1_L54_Parts, LM_Ins1_L54_Playfield, LM_Ins1_L54_plastics, LM_Ins1_L54_psw33, LM_Ins1_L55_Playfield, LM_Ins1_L55_ins_mesh, LM_Ins1_L55_psw17, LM_Ins1_L55_psw18, LM_Ins1_L55_psw19, LM_Ins1_L55_psw20, LM_Ins1_L55_psw21, LM_Ins1_L55_psw22, LM_Ins1_L55_psw23, LM_Ins1_L58_Flip3, LM_Ins1_L58_Parts, LM_Ins1_L58_Playfield, LM_Ins1_L58_ins_mesh, LM_Ins1_L59_Parts, LM_Ins1_L59_Playfield, LM_Ins1_L6_Playfield, LM_Ins1_L6_ins_mesh, LM_Ins1_L6_psw17, LM_Ins1_L6_psw18, LM_Ins1_L6_psw19, LM_Ins1_L6_psw20, LM_Ins1_L6_psw21, LM_Ins1_L6_psw22, LM_Ins1_L6_psw23, LM_Ins1_L60_Flip1, LM_Ins1_L60_Flip2, _
  LM_Ins1_L60_Parts, LM_Ins1_L60_Playfield, LM_Ins1_L60_ins_mesh, LM_Ins1_L63_Flip3, LM_Ins1_L63_Playfield, LM_Ins1_L63_ins_mesh, LM_Ins1_L63_psw18, LM_Ins1_L63_psw2, LM_Ins1_L63_psw21, LM_Ins1_L63_psw22, LM_Ins1_L63_psw23, LM_Ins1_L63_psw3, LM_Ins1_L65_Flip3, LM_Ins1_L65_Playfield, LM_Ins1_L65_ins_mesh, LM_Ins1_L65_psw1, LM_Ins1_L7_BumperSkirt3, LM_Ins1_L7_Layer_2, LM_Ins1_L7_Playfield, LM_Ins1_L7_ins_mesh, LM_Ins1_L7_psw17, LM_Ins1_L7_psw18, LM_Ins1_L7_psw19, LM_Ins1_L7_psw20, LM_Ins1_L7_psw21, LM_Ins1_L7_psw22, LM_Ins1_L70_BumperSkirt1, LM_Ins1_L70_Layer_2, LM_Ins1_L70_Parts, LM_Ins1_L70_Playfield, LM_Ins1_L70_bumpcapA_3_002, LM_Ins1_L70_plastics, LM_Ins1_L8_Playfield, LM_Ins1_L8_ins_mesh, LM_Ins1_L8_psw17, LM_Ins1_L8_psw18, LM_Ins1_L8_psw19, LM_Ins1_L8_psw20, LM_Ins1_L8_psw21, LM_Ins1_L8_psw22, LM_Ins1_L8_psw23, LM_Ins1_L81_Flip3, LM_Ins1_L81_Parts, LM_Ins1_L81_Playfield, LM_Ins1_L81_ins_mesh, LM_Ins1_L81_psw1, LM_Ins1_L81_psw2, LM_Ins1_L82_Parts, LM_Ins1_L82_Playfield, LM_Ins1_L82_bulbs, _
  LM_Ins1_L82_plastics, LM_Ins1_L82_psw1, LM_Ins1_L82_psw2, LM_Ins1_L82_psw3, LM_Ins1_L82_psw4, LM_Ins1_L86_BumperSkirt2, LM_Ins1_L86_Layer_2, LM_Ins1_L86_Parts, LM_Ins1_L86_Playfield, LM_Ins1_L86_bumpcapA_3_002, LM_Ins1_L97_Flip3, LM_Ins1_L97_Parts, LM_Ins1_L97_Playfield, LM_Ins1_L97_bulbs, LM_Ins1_L97_ins_mesh, LM_Ins1_L97_psw1, LM_Ins1_L97_psw2, LM_Ins2_L100_Playfield, LM_Ins2_L100_bulbs, LM_Ins2_L100_ins_mesh, LM_Ins2_L11_Flip1, LM_Ins2_L11_Flip2, LM_Ins2_L11_bulbs, LM_Ins2_L11_ins_mesh, LM_Ins2_L113_Parts, LM_Ins2_L113_Playfield, LM_Ins2_L113_bulbs, LM_Ins2_L113_ins_mesh, LM_Ins2_L113_psw1, LM_Ins2_L116_Playfield, LM_Ins2_L116_bulbs, LM_Ins2_L116_ins_mesh, LM_Ins2_L12_Flip2, LM_Ins2_L12_Playfield, LM_Ins2_L12_bulbs, LM_Ins2_L12_ins_mesh, LM_Ins2_L25_Parts, LM_Ins2_L25_Playfield, LM_Ins2_L25_bulbs, LM_Ins2_L25_ins_mesh, LM_Ins2_L41_Flip3, LM_Ins2_L41_Playfield, LM_Ins2_L41_bulbs, LM_Ins2_L41_ins_mesh, LM_Ins2_L41_plastics, LM_Ins2_L44_Flip1, LM_Ins2_L44_Playfield, LM_Ins2_L44_bulbs, LM_Ins2_L44_ins_mesh, _
  LM_Ins2_L47_Playfield, LM_Ins2_L47_bulbs, LM_Ins2_L47_ins_mesh, LM_Ins2_L57_Playfield, LM_Ins2_L57_bulbs, LM_Ins2_L57_ins_mesh, LM_Ins2_L57_plastics, LM_Ins2_L66_Parts, LM_Ins2_L66_bulbs, LM_Ins2_L66_ins_mesh, LM_Ins2_L66_plastics, LM_Ins2_L66_psw1, LM_Ins2_L67_Playfield, LM_Ins2_L67_bulbs, LM_Ins2_L67_ins_mesh, LM_Ins2_L67_psw17, LM_Ins2_L67_psw18, LM_Ins2_L67_psw19, LM_Ins2_L68_Layer_2, LM_Ins2_L68_Playfield, LM_Ins2_L68_bulbs, LM_Ins2_L68_ins_mesh, LM_Ins2_L69_Playfield, LM_Ins2_L69_bulbs, LM_Ins2_L69_ins_mesh, LM_Ins2_L71_Flip2, LM_Ins2_L71_Playfield, LM_Ins2_L71_bulbs, LM_Ins2_L71_ins_mesh, LM_Ins2_L83_Playfield, LM_Ins2_L83_bulbs, LM_Ins2_L83_ins_mesh, LM_Ins2_L83_psw17, LM_Ins2_L83_psw18, LM_Ins2_L83_psw19, LM_Ins2_L84_Playfield, LM_Ins2_L84_bulbs, LM_Ins2_L84_ins_mesh, LM_Ins2_L85_Playfield, LM_Ins2_L85_bulbs, LM_Ins2_L85_ins_mesh, LM_Ins2_L87_Flip1, LM_Ins2_L87_Playfield, LM_Ins2_L87_bulbs, LM_Ins2_L87_ins_mesh, LM_Ins2_L9_Parts, LM_Ins2_L9_Playfield, LM_Ins2_L9_bulbs, LM_Ins2_L9_ins_mesh, _
  LM_Ins2_L98_bulbs, LM_Ins2_L98_ins_mesh, LM_Ins2_L99_Playfield, LM_Ins2_L99_bulbs, LM_Ins2_L99_ins_mesh, LM_Ins2_L99_psw17)
Dim BG_All: BG_All=Array(BM_BumperRing1, BM_BumperRing2, BM_BumperRing3, BM_BumperSkirt1, BM_BumperSkirt2, BM_BumperSkirt3, BM_Flip1, BM_Flip2, BM_Flip3, BM_Gateplate2, BM_Gateplate5, BM_Lane_Guide___2a, BM_Layer_2, BM_Parts, BM_Playfield, BM_ROstar, BM_bulbs, BM_bumpcapA_3_002, BM_gatewire3, BM_gatewire4, BM_ins_mesh, BM_lsling, BM_lsling1, BM_lsling2, BM_plastics, BM_psw1, BM_psw17, BM_psw18, BM_psw19, BM_psw2, BM_psw20, BM_psw21, BM_psw22, BM_psw23, BM_psw3, BM_psw33, BM_psw4, BM_rails, BM_rsling, BM_rsling1, BM_rsling2, BM_slingarmL, BM_slingarmR, BM_sw12, BM_sw13, BM_sw14, BM_sw15, BM_sw31, BM_sw32, BM_tsw25, BM_tsw26, BM_tsw27, BM_tsw28, BM_tsw29, BM_tsw30, BM_tsw5, LM_GI_GI_1_Layer_2, LM_GI_GI_1_Parts, LM_GI_GI_1_Playfield, LM_GI_GI_1_bulbs, LM_GI_GI_1_plastics, LM_GI_GI_1_psw17, LM_GI_GI_1_psw33, LM_GI_GI_10_Flip1, LM_GI_GI_10_Flip2, LM_GI_GI_10_Layer_2, LM_GI_GI_10_Parts, LM_GI_GI_10_Playfield, LM_GI_GI_10_bulbs, LM_GI_GI_10_lsling, LM_GI_GI_10_lsling1, LM_GI_GI_10_lsling2, LM_GI_GI_10_plastics, _
  LM_GI_GI_10_psw21, LM_GI_GI_10_psw22, LM_GI_GI_10_psw23, LM_GI_GI_10_rsling, LM_GI_GI_10_rsling1, LM_GI_GI_10_rsling2, LM_GI_GI_11_Flip1, LM_GI_GI_11_Flip2, LM_GI_GI_11_Flip3, LM_GI_GI_11_Layer_2, LM_GI_GI_11_Parts, LM_GI_GI_11_Playfield, LM_GI_GI_11_bulbs, LM_GI_GI_11_ins_mesh, LM_GI_GI_11_lsling, LM_GI_GI_11_lsling1, LM_GI_GI_11_lsling2, LM_GI_GI_11_plastics, LM_GI_GI_11_psw19, LM_GI_GI_11_psw20, LM_GI_GI_11_psw21, LM_GI_GI_11_psw22, LM_GI_GI_11_psw23, LM_GI_GI_11_rails, LM_GI_GI_11_rsling, LM_GI_GI_11_rsling1, LM_GI_GI_11_rsling2, LM_GI_GI_11_slingarmL, LM_GI_GI_11_sw14, LM_GI_GI_11_sw32, LM_GI_GI_13_Flip1, LM_GI_GI_13_Flip2, LM_GI_GI_13_Flip3, LM_GI_GI_13_Layer_2, LM_GI_GI_13_Parts, LM_GI_GI_13_Playfield, LM_GI_GI_13_bulbs, LM_GI_GI_13_ins_mesh, LM_GI_GI_13_lsling, LM_GI_GI_13_lsling1, LM_GI_GI_13_lsling2, LM_GI_GI_13_plastics, LM_GI_GI_13_psw1, LM_GI_GI_13_psw2, LM_GI_GI_13_psw20, LM_GI_GI_13_psw21, LM_GI_GI_13_psw22, LM_GI_GI_13_psw23, LM_GI_GI_13_rails, LM_GI_GI_13_rsling, LM_GI_GI_13_rsling1, _
  LM_GI_GI_13_rsling2, LM_GI_GI_13_slingarmR, LM_GI_GI_13_sw15, LM_GI_GI_13_sw31, LM_GI_GI_13_tsw30, LM_GI_GI_14_Flip1, LM_GI_GI_14_Flip2, LM_GI_GI_14_Flip3, LM_GI_GI_14_Layer_2, LM_GI_GI_14_Parts, LM_GI_GI_14_Playfield, LM_GI_GI_14_bulbs, LM_GI_GI_14_lsling, LM_GI_GI_14_lsling1, LM_GI_GI_14_lsling2, LM_GI_GI_14_plastics, LM_GI_GI_14_psw18, LM_GI_GI_14_psw19, LM_GI_GI_14_psw20, LM_GI_GI_14_psw21, LM_GI_GI_14_psw22, LM_GI_GI_14_rsling, LM_GI_GI_14_rsling1, LM_GI_GI_14_rsling2, LM_GI_GI_14_sw14, LM_GI_GI_15_Flip1, LM_GI_GI_15_Layer_2, LM_GI_GI_15_Parts, LM_GI_GI_15_Playfield, LM_GI_GI_15_bulbs, LM_GI_GI_15_lsling, LM_GI_GI_15_lsling1, LM_GI_GI_15_lsling2, LM_GI_GI_15_plastics, LM_GI_GI_15_psw1, LM_GI_GI_15_psw18, LM_GI_GI_15_psw2, LM_GI_GI_15_psw20, LM_GI_GI_15_psw23, LM_GI_GI_15_psw3, LM_GI_GI_15_rails, LM_GI_GI_15_rsling, LM_GI_GI_15_rsling1, LM_GI_GI_15_rsling2, LM_GI_GI_15_slingarmR, LM_GI_GI_15_sw32, LM_GI_GI_17_BumperSkirt1, LM_GI_GI_17_BumperSkirt2, LM_GI_GI_17_Layer_2, LM_GI_GI_17_Parts, _
  LM_GI_GI_17_Playfield, LM_GI_GI_17_bulbs, LM_GI_GI_17_plastics, LM_GI_GI_17_psw33, LM_GI_GI_17_rails, LM_GI_GI_18_BumperRing1, LM_GI_GI_18_BumperSkirt1, LM_GI_GI_18_BumperSkirt2, LM_GI_GI_18_Layer_2, LM_GI_GI_18_Parts, LM_GI_GI_18_Playfield, LM_GI_GI_18_ROstar, LM_GI_GI_18_bulbs, LM_GI_GI_18_plastics, LM_GI_GI_18_psw1, LM_GI_GI_18_psw2, LM_GI_GI_18_psw3, LM_GI_GI_18_psw4, LM_GI_GI_18_tsw5, LM_GI_GI_19_BumperSkirt2, LM_GI_GI_19_Layer_2, LM_GI_GI_19_Parts, LM_GI_GI_19_Playfield, LM_GI_GI_19_bulbs, LM_GI_GI_19_plastics, LM_GI_GI_19_psw33, LM_GI_GI_20_BumperRing2, LM_GI_GI_20_BumperRing3, LM_GI_GI_20_BumperSkirt1, LM_GI_GI_20_BumperSkirt2, LM_GI_GI_20_BumperSkirt3, LM_GI_GI_20_Layer_2, LM_GI_GI_20_Parts, LM_GI_GI_20_Playfield, LM_GI_GI_20_ROstar, LM_GI_GI_20_bulbs, LM_GI_GI_20_plastics, LM_GI_GI_20_psw1, LM_GI_GI_20_psw17, LM_GI_GI_20_psw18, LM_GI_GI_20_psw19, LM_GI_GI_20_psw2, LM_GI_GI_20_psw22, LM_GI_GI_20_psw3, LM_GI_GI_20_psw4, LM_GI_GI_20_rails, LM_GI_GI_20_tsw25, LM_GI_GI_21_BumperRing1, _
  LM_GI_GI_21_BumperRing2, LM_GI_GI_21_BumperSkirt1, LM_GI_GI_21_BumperSkirt2, LM_GI_GI_21_BumperSkirt3, LM_GI_GI_21_Layer_2, LM_GI_GI_21_Parts, LM_GI_GI_21_Playfield, LM_GI_GI_21_bulbs, LM_GI_GI_21_bumpcapA_3_002, LM_GI_GI_21_ins_mesh, LM_GI_GI_21_plastics, LM_GI_GI_21_psw17, LM_GI_GI_21_psw18, LM_GI_GI_21_psw19, LM_GI_GI_21_psw2, LM_GI_GI_21_psw20, LM_GI_GI_21_psw21, LM_GI_GI_21_psw22, LM_GI_GI_21_psw23, LM_GI_GI_21_psw33, LM_GI_GI_21_rails, LM_GI_GI_21_tsw25, LM_GI_GI_21_tsw26, LM_GI_GI_21_tsw27, LM_GI_GI_21_tsw28, LM_GI_GI_21_tsw29, LM_GI_GI_22_BumperSkirt1, LM_GI_GI_22_BumperSkirt2, LM_GI_GI_22_Gateplate2, LM_GI_GI_22_Lane_Guide___2a, LM_GI_GI_22_Layer_2, LM_GI_GI_22_Parts, LM_GI_GI_22_Playfield, LM_GI_GI_22_ROstar, LM_GI_GI_22_bulbs, LM_GI_GI_22_plastics, LM_GI_GI_22_psw19, LM_GI_GI_22_psw20, LM_GI_GI_22_psw21, LM_GI_GI_22_psw22, LM_GI_GI_22_psw23, LM_GI_GI_22_psw4, LM_GI_GI_22_sw12, LM_GI_GI_22_sw13, LM_GI_GI_22_tsw5, LM_GI_GI_23_BumperSkirt1, LM_GI_GI_23_BumperSkirt2, LM_GI_GI_23_Lane_Guide___2a, _
  LM_GI_GI_23_Layer_2, LM_GI_GI_23_Parts, LM_GI_GI_23_Playfield, LM_GI_GI_23_bulbs, LM_GI_GI_23_psw1, LM_GI_GI_23_psw19, LM_GI_GI_23_psw2, LM_GI_GI_23_psw3, LM_GI_GI_23_psw4, LM_GI_GI_23_sw12, LM_GI_GI_23_sw13, LM_GI_GI_23_tsw5, LM_GI_GI_24_BumperSkirt1, LM_GI_GI_24_BumperSkirt2, LM_GI_GI_24_Lane_Guide___2a, LM_GI_GI_24_Layer_2, LM_GI_GI_24_Parts, LM_GI_GI_24_Playfield, LM_GI_GI_24_bulbs, LM_GI_GI_24_plastics, LM_GI_GI_24_psw17, LM_GI_GI_24_psw19, LM_GI_GI_24_psw2, LM_GI_GI_24_psw20, LM_GI_GI_24_psw21, LM_GI_GI_24_psw22, LM_GI_GI_24_psw23, LM_GI_GI_24_psw3, LM_GI_GI_24_psw4, LM_GI_GI_24_sw12, LM_GI_GI_24_sw13, LM_GI_GI_25_BumperSkirt1, LM_GI_GI_25_BumperSkirt2, LM_GI_GI_25_Layer_2, LM_GI_GI_25_Parts, LM_GI_GI_25_Playfield, LM_GI_GI_25_bulbs, LM_GI_GI_25_plastics, LM_GI_GI_25_psw17, LM_GI_GI_25_psw18, LM_GI_GI_25_psw19, LM_GI_GI_25_psw20, LM_GI_GI_25_psw21, LM_GI_GI_25_psw33, LM_GI_GI_25_rails, LM_GI_GI_4_BumperRing3, LM_GI_GI_4_BumperSkirt3, LM_GI_GI_4_Flip1, LM_GI_GI_4_Flip3, LM_GI_GI_4_Parts, _
  LM_GI_GI_4_Playfield, LM_GI_GI_4_bulbs, LM_GI_GI_4_plastics, LM_GI_GI_4_psw1, LM_GI_GI_4_psw17, LM_GI_GI_4_psw18, LM_GI_GI_4_psw19, LM_GI_GI_4_psw20, LM_GI_GI_4_psw21, LM_GI_GI_4_psw22, LM_GI_GI_4_psw23, LM_GI_GI_4_rails, LM_GI_GI_4_rsling, LM_GI_GI_4_rsling1, LM_GI_GI_4_rsling2, LM_GI_GI_4_tsw26, LM_GI_GI_4_tsw27, LM_GI_GI_4_tsw28, LM_GI_GI_4_tsw29, LM_GI_GI_4_tsw30, LM_GI_GI_6_Layer_2, LM_GI_GI_6_Parts, LM_GI_GI_6_Playfield, LM_GI_GI_6_bulbs, LM_GI_GI_6_plastics, LM_GI_GI_6_psw33, LM_GI_GI_8_Flip1, LM_GI_GI_8_Flip2, LM_GI_GI_8_Flip3, LM_GI_GI_8_Parts, LM_GI_GI_8_Playfield, LM_GI_GI_8_bulbs, LM_GI_GI_8_ins_mesh, LM_GI_GI_8_lsling, LM_GI_GI_8_lsling1, LM_GI_GI_8_lsling2, LM_GI_GI_8_plastics, LM_GI_GI_8_psw1, LM_GI_GI_8_psw17, LM_GI_GI_8_psw18, LM_GI_GI_8_psw19, LM_GI_GI_8_psw2, LM_GI_GI_8_psw20, LM_GI_GI_8_psw21, LM_GI_GI_8_psw22, LM_GI_GI_8_psw23, LM_GI_GI_8_rails, LM_GI_GI_8_sw14, LM_GI_GI_9_Flip1, LM_GI_GI_9_Flip2, LM_GI_GI_9_Flip3, LM_GI_GI_9_Layer_2, LM_GI_GI_9_Parts, LM_GI_GI_9_Playfield, _
  LM_GI_GI_9_bulbs, LM_GI_GI_9_lsling, LM_GI_GI_9_lsling1, LM_GI_GI_9_lsling2, LM_GI_GI_9_plastics, LM_GI_GI_9_psw1, LM_GI_GI_9_psw17, LM_GI_GI_9_psw18, LM_GI_GI_9_psw19, LM_GI_GI_9_psw2, LM_GI_GI_9_psw20, LM_GI_GI_9_psw21, LM_GI_GI_9_rails, LM_GI_GI_9_rsling, LM_GI_GI_9_rsling1, LM_GI_GI_9_rsling2, LM_GI_GI_9_slingarmL, LM_GI_GI_9_sw31, LM_GI_gi1_BumperSkirt2, LM_GI_gi1_Flip1, LM_GI_gi1_Flip2, LM_GI_gi1_Flip3, LM_GI_gi1_Gateplate5, LM_GI_gi1_Layer_2, LM_GI_gi1_Parts, LM_GI_gi1_Playfield, LM_GI_gi1_ROstar, LM_GI_gi1_bulbs, LM_GI_gi1_gatewire4, LM_GI_gi1_lsling, LM_GI_gi1_lsling1, LM_GI_gi1_lsling2, LM_GI_gi1_plastics, LM_GI_gi1_psw1, LM_GI_gi1_psw17, LM_GI_gi1_psw18, LM_GI_gi1_psw19, LM_GI_gi1_psw2, LM_GI_gi1_psw20, LM_GI_gi1_psw21, LM_GI_gi1_psw22, LM_GI_gi1_psw23, LM_GI_gi1_psw3, LM_GI_gi1_psw33, LM_GI_gi1_psw4, LM_GI_gi1_rails, LM_GI_gi1_tsw25, LM_GI_gi1_tsw26, LM_GI_gi1_tsw27, LM_GI_gi1_tsw28, LM_GI_gi1_tsw29, LM_GI_gi1_tsw30, LM_GI_gi1_tsw5, LM_Ins1_L1_Flip1, LM_Ins1_L1_Flip2, LM_Ins1_L1_Playfield, _
  LM_Ins1_L10_Parts, LM_Ins1_L10_Playfield, LM_Ins1_L10_ins_mesh, LM_Ins1_L10_plastics, LM_Ins1_L101_BumperRing3, LM_Ins1_L101_BumperSkirt1, LM_Ins1_L101_BumperSkirt2, LM_Ins1_L101_BumperSkirt3, LM_Ins1_L101_Flip3, LM_Ins1_L101_Layer_2, LM_Ins1_L101_Parts, LM_Ins1_L101_Playfield, LM_Ins1_L101_bulbs, LM_Ins1_L101_plastics, LM_Ins1_L101_psw1, LM_Ins1_L101_psw2, LM_Ins1_L101_psw20, LM_Ins1_L101_psw21, LM_Ins1_L101_psw22, LM_Ins1_L101_psw23, LM_Ins1_L101_psw3, LM_Ins1_L101_psw4, LM_Ins1_L101_tsw5, LM_Ins1_L101a_BumperSkirt2, LM_Ins1_L101a_BumperSkirt3, LM_Ins1_L101a_Layer_2, LM_Ins1_L101a_Parts, LM_Ins1_L101a_Playfield, LM_Ins1_L101a_bulbs, LM_Ins1_L101a_ins_mesh, LM_Ins1_L101a_plastics, LM_Ins1_L101a_psw17, LM_Ins1_L101a_psw18, LM_Ins1_L101a_psw33, LM_Ins1_L101a_rails, LM_Ins1_L101a_tsw27, LM_Ins1_L102_BumperSkirt3, LM_Ins1_L102_Layer_2, LM_Ins1_L102_Parts, LM_Ins1_L102_Playfield, LM_Ins1_L102_bumpcapA_3_002, LM_Ins1_L103_Flip1, LM_Ins1_L103_Flip2, LM_Ins1_L103_Flip3, LM_Ins1_L103_Playfield, LM_Ins1_L103_ins_mesh, _
  LM_Ins1_L114_Flip3, LM_Ins1_L114_Parts, LM_Ins1_L114_Playfield, LM_Ins1_L114_ins_mesh, LM_Ins1_L114_psw1, LM_Ins1_L114_psw2, LM_Ins1_L114_psw3, LM_Ins1_L115_Layer_2, LM_Ins1_L115_Parts, LM_Ins1_L115_Playfield, LM_Ins1_L115_bulbs, LM_Ins1_L115_plastics, LM_Ins1_L115_psw17, LM_Ins1_L115_psw33, LM_Ins1_L117_BumperRing1, LM_Ins1_L117_BumperSkirt1, LM_Ins1_L117_BumperSkirt3, LM_Ins1_L117_Layer_2, LM_Ins1_L117_Parts, LM_Ins1_L117_Playfield, LM_Ins1_L117_ROstar, LM_Ins1_L117_bulbs, LM_Ins1_L117_plastics, LM_Ins1_L117_psw1, LM_Ins1_L117_psw17, LM_Ins1_L117_psw2, LM_Ins1_L117_psw23, LM_Ins1_L117_psw3, LM_Ins1_L117_psw33, LM_Ins1_L117_psw4, LM_Ins1_L117_tsw5, LM_Ins1_L117a_BumperSkirt2, LM_Ins1_L117a_Layer_2, LM_Ins1_L117a_Parts, LM_Ins1_L117a_Playfield, LM_Ins1_L117a_bulbs, LM_Ins1_L117a_plastics, LM_Ins1_L117a_psw17, LM_Ins1_L117a_psw18, LM_Ins1_L117a_psw33, LM_Ins1_L118_BumperSkirt1, LM_Ins1_L118_Gateplate2, LM_Ins1_L118_Layer_2, LM_Ins1_L118_Parts, LM_Ins1_L118_Playfield, LM_Ins1_L118_ROstar, LM_Ins1_L118_bulbs, _
  LM_Ins1_L118_gatewire4, LM_Ins1_L118_plastics, LM_Ins1_L118_psw1, LM_Ins1_L118_psw17, LM_Ins1_L118_psw2, LM_Ins1_L118_psw3, LM_Ins1_L118_psw4, LM_Ins1_L118_rails, LM_Ins1_L118_tsw5, LM_Ins1_L118a_Layer_2, LM_Ins1_L118a_Parts, LM_Ins1_L118a_Playfield, LM_Ins1_L118a_bulbs, LM_Ins1_L118a_plastics, LM_Ins1_L118a_psw17, LM_Ins1_L118a_psw18, LM_Ins1_L118a_psw3, LM_Ins1_L118a_psw33, LM_Ins1_L118a_rails, LM_Ins1_L118a_tsw25, LM_Ins1_L119_BumperSkirt1, LM_Ins1_L119_Gateplate5, LM_Ins1_L119_Layer_2, LM_Ins1_L119_Parts, LM_Ins1_L119_Playfield, LM_Ins1_L119_ROstar, LM_Ins1_L119_bulbs, LM_Ins1_L119_plastics, LM_Ins1_L119_psw1, LM_Ins1_L119_psw2, LM_Ins1_L119_psw3, LM_Ins1_L119_psw4, LM_Ins1_L119_rails, LM_Ins1_L119_tsw5, LM_Ins1_L119a_BumperSkirt2, LM_Ins1_L119a_Layer_2, LM_Ins1_L119a_Parts, LM_Ins1_L119a_Playfield, LM_Ins1_L119a_bulbs, LM_Ins1_L119a_gatewire3, LM_Ins1_L119a_plastics, LM_Ins1_L119a_psw33, LM_Ins1_L17_Flip2, LM_Ins1_L17_Flip3, LM_Ins1_L17_Parts, LM_Ins1_L17_Playfield, LM_Ins1_L18_Playfield, _
  LM_Ins1_L18_psw22, LM_Ins1_L18_psw23, LM_Ins1_L19_Playfield, LM_Ins1_L19_ins_mesh, LM_Ins1_L19_psw21, LM_Ins1_L19_psw22, LM_Ins1_L19_psw23, LM_Ins1_L19_sw31, LM_Ins1_L2_Flip3, LM_Ins1_L2_Playfield, LM_Ins1_L20_Flip3, LM_Ins1_L20_Playfield, LM_Ins1_L20_ins_mesh, LM_Ins1_L20_psw21, LM_Ins1_L20_psw22, LM_Ins1_L20_psw23, LM_Ins1_L21_Playfield, LM_Ins1_L21_ins_mesh, LM_Ins1_L21_psw17, LM_Ins1_L21_psw18, LM_Ins1_L21_psw19, LM_Ins1_L21_psw20, LM_Ins1_L21_psw21, LM_Ins1_L21_psw22, LM_Ins1_L21_psw23, LM_Ins1_L21_tsw25, LM_Ins1_L22_Playfield, LM_Ins1_L22_ins_mesh, LM_Ins1_L22_psw17, LM_Ins1_L22_psw18, LM_Ins1_L22_psw19, LM_Ins1_L22_psw20, LM_Ins1_L22_psw21, LM_Ins1_L22_psw22, LM_Ins1_L22_psw23, LM_Ins1_L22_tsw30, LM_Ins1_L23_Playfield, LM_Ins1_L23_ins_mesh, LM_Ins1_L23_psw17, LM_Ins1_L23_psw18, LM_Ins1_L23_psw19, LM_Ins1_L23_psw20, LM_Ins1_L23_psw21, LM_Ins1_L23_psw22, LM_Ins1_L23_psw23, LM_Ins1_L24_Playfield, LM_Ins1_L24_ins_mesh, LM_Ins1_L24_psw17, LM_Ins1_L24_psw18, LM_Ins1_L24_psw19, LM_Ins1_L24_psw20, _
  LM_Ins1_L24_psw21, LM_Ins1_L24_psw22, LM_Ins1_L24_psw23, LM_Ins1_L26_Layer_2, LM_Ins1_L26_Parts, LM_Ins1_L26_Playfield, LM_Ins1_L26_ins_mesh, LM_Ins1_L26_plastics, LM_Ins1_L28_Flip1, LM_Ins1_L28_Flip2, LM_Ins1_L28_Playfield, LM_Ins1_L28_ins_mesh, LM_Ins1_L3_Playfield, LM_Ins1_L3_ins_mesh, LM_Ins1_L3_psw21, LM_Ins1_L3_psw22, LM_Ins1_L3_psw23, LM_Ins1_L33_Flip1, LM_Ins1_L33_Playfield, LM_Ins1_L34_Flip3, LM_Ins1_L34_Playfield, LM_Ins1_L34_sw14, LM_Ins1_L35_Flip3, LM_Ins1_L35_Playfield, LM_Ins1_L36_Playfield, LM_Ins1_L36_ins_mesh, LM_Ins1_L36_psw19, LM_Ins1_L36_psw20, LM_Ins1_L36_psw21, LM_Ins1_L36_psw22, LM_Ins1_L36_psw23, LM_Ins1_L37_Playfield, LM_Ins1_L37_ins_mesh, LM_Ins1_L37_psw17, LM_Ins1_L37_psw18, LM_Ins1_L37_psw19, LM_Ins1_L37_psw20, LM_Ins1_L37_psw21, LM_Ins1_L37_psw22, LM_Ins1_L37_psw23, LM_Ins1_L37_tsw27, LM_Ins1_L38_Playfield, LM_Ins1_L38_ins_mesh, LM_Ins1_L38_psw17, LM_Ins1_L38_psw18, LM_Ins1_L38_psw19, LM_Ins1_L38_psw20, LM_Ins1_L38_psw21, LM_Ins1_L38_psw22, LM_Ins1_L38_psw23, _
  LM_Ins1_L39_Playfield, LM_Ins1_L39_ins_mesh, LM_Ins1_L39_psw17, LM_Ins1_L39_psw18, LM_Ins1_L39_psw19, LM_Ins1_L39_psw20, LM_Ins1_L39_psw21, LM_Ins1_L39_psw22, LM_Ins1_L39_psw23, LM_Ins1_L4_Flip3, LM_Ins1_L4_Playfield, LM_Ins1_L4_ins_mesh, LM_Ins1_L40_Playfield, LM_Ins1_L40_ins_mesh, LM_Ins1_L40_psw18, LM_Ins1_L40_psw19, LM_Ins1_L40_psw20, LM_Ins1_L40_psw21, LM_Ins1_L40_psw22, LM_Ins1_L40_psw23, LM_Ins1_L42_Parts, LM_Ins1_L42_Playfield, LM_Ins1_L42_ins_mesh, LM_Ins1_L42_psw20, LM_Ins1_L42_psw23, LM_Ins1_L49_Flip3, LM_Ins1_L49_Playfield, LM_Ins1_L5_Parts, LM_Ins1_L5_Playfield, LM_Ins1_L5_ins_mesh, LM_Ins1_L5_psw17, LM_Ins1_L5_psw18, LM_Ins1_L5_psw19, LM_Ins1_L5_psw20, LM_Ins1_L5_psw21, LM_Ins1_L5_psw22, LM_Ins1_L5_psw23, LM_Ins1_L5_tsw25, LM_Ins1_L50_Flip3, LM_Ins1_L50_Playfield, LM_Ins1_L50_ins_mesh, LM_Ins1_L51_Flip3, LM_Ins1_L51_Playfield, LM_Ins1_L51_ins_mesh, LM_Ins1_L53_Playfield, LM_Ins1_L53_ins_mesh, LM_Ins1_L53_psw17, LM_Ins1_L53_psw18, LM_Ins1_L53_psw19, LM_Ins1_L53_psw20, LM_Ins1_L53_psw21, _
  LM_Ins1_L53_psw22, LM_Ins1_L53_psw23, LM_Ins1_L53_tsw28, LM_Ins1_L54_Parts, LM_Ins1_L54_Playfield, LM_Ins1_L54_plastics, LM_Ins1_L54_psw33, LM_Ins1_L55_Playfield, LM_Ins1_L55_ins_mesh, LM_Ins1_L55_psw17, LM_Ins1_L55_psw18, LM_Ins1_L55_psw19, LM_Ins1_L55_psw20, LM_Ins1_L55_psw21, LM_Ins1_L55_psw22, LM_Ins1_L55_psw23, LM_Ins1_L58_Flip3, LM_Ins1_L58_Parts, LM_Ins1_L58_Playfield, LM_Ins1_L58_ins_mesh, LM_Ins1_L59_Parts, LM_Ins1_L59_Playfield, LM_Ins1_L6_Playfield, LM_Ins1_L6_ins_mesh, LM_Ins1_L6_psw17, LM_Ins1_L6_psw18, LM_Ins1_L6_psw19, LM_Ins1_L6_psw20, LM_Ins1_L6_psw21, LM_Ins1_L6_psw22, LM_Ins1_L6_psw23, LM_Ins1_L60_Flip1, LM_Ins1_L60_Flip2, LM_Ins1_L60_Parts, LM_Ins1_L60_Playfield, LM_Ins1_L60_ins_mesh, LM_Ins1_L63_Flip3, LM_Ins1_L63_Playfield, LM_Ins1_L63_ins_mesh, LM_Ins1_L63_psw18, LM_Ins1_L63_psw2, LM_Ins1_L63_psw21, LM_Ins1_L63_psw22, LM_Ins1_L63_psw23, LM_Ins1_L63_psw3, LM_Ins1_L65_Flip3, LM_Ins1_L65_Playfield, LM_Ins1_L65_ins_mesh, LM_Ins1_L65_psw1, LM_Ins1_L7_BumperSkirt3, LM_Ins1_L7_Layer_2, _
  LM_Ins1_L7_Playfield, LM_Ins1_L7_ins_mesh, LM_Ins1_L7_psw17, LM_Ins1_L7_psw18, LM_Ins1_L7_psw19, LM_Ins1_L7_psw20, LM_Ins1_L7_psw21, LM_Ins1_L7_psw22, LM_Ins1_L70_BumperSkirt1, LM_Ins1_L70_Layer_2, LM_Ins1_L70_Parts, LM_Ins1_L70_Playfield, LM_Ins1_L70_bumpcapA_3_002, LM_Ins1_L70_plastics, LM_Ins1_L8_Playfield, LM_Ins1_L8_ins_mesh, LM_Ins1_L8_psw17, LM_Ins1_L8_psw18, LM_Ins1_L8_psw19, LM_Ins1_L8_psw20, LM_Ins1_L8_psw21, LM_Ins1_L8_psw22, LM_Ins1_L8_psw23, LM_Ins1_L81_Flip3, LM_Ins1_L81_Parts, LM_Ins1_L81_Playfield, LM_Ins1_L81_ins_mesh, LM_Ins1_L81_psw1, LM_Ins1_L81_psw2, LM_Ins1_L82_Parts, LM_Ins1_L82_Playfield, LM_Ins1_L82_bulbs, LM_Ins1_L82_plastics, LM_Ins1_L82_psw1, LM_Ins1_L82_psw2, LM_Ins1_L82_psw3, LM_Ins1_L82_psw4, LM_Ins1_L86_BumperSkirt2, LM_Ins1_L86_Layer_2, LM_Ins1_L86_Parts, LM_Ins1_L86_Playfield, LM_Ins1_L86_bumpcapA_3_002, LM_Ins1_L97_Flip3, LM_Ins1_L97_Parts, LM_Ins1_L97_Playfield, LM_Ins1_L97_bulbs, LM_Ins1_L97_ins_mesh, LM_Ins1_L97_psw1, LM_Ins1_L97_psw2, LM_Ins2_L100_Playfield, _
  LM_Ins2_L100_bulbs, LM_Ins2_L100_ins_mesh, LM_Ins2_L11_Flip1, LM_Ins2_L11_Flip2, LM_Ins2_L11_bulbs, LM_Ins2_L11_ins_mesh, LM_Ins2_L113_Parts, LM_Ins2_L113_Playfield, LM_Ins2_L113_bulbs, LM_Ins2_L113_ins_mesh, LM_Ins2_L113_psw1, LM_Ins2_L116_Playfield, LM_Ins2_L116_bulbs, LM_Ins2_L116_ins_mesh, LM_Ins2_L12_Flip2, LM_Ins2_L12_Playfield, LM_Ins2_L12_bulbs, LM_Ins2_L12_ins_mesh, LM_Ins2_L25_Parts, LM_Ins2_L25_Playfield, LM_Ins2_L25_bulbs, LM_Ins2_L25_ins_mesh, LM_Ins2_L41_Flip3, LM_Ins2_L41_Playfield, LM_Ins2_L41_bulbs, LM_Ins2_L41_ins_mesh, LM_Ins2_L41_plastics, LM_Ins2_L44_Flip1, LM_Ins2_L44_Playfield, LM_Ins2_L44_bulbs, LM_Ins2_L44_ins_mesh, LM_Ins2_L47_Playfield, LM_Ins2_L47_bulbs, LM_Ins2_L47_ins_mesh, LM_Ins2_L57_Playfield, LM_Ins2_L57_bulbs, LM_Ins2_L57_ins_mesh, LM_Ins2_L57_plastics, LM_Ins2_L66_Parts, LM_Ins2_L66_bulbs, LM_Ins2_L66_ins_mesh, LM_Ins2_L66_plastics, LM_Ins2_L66_psw1, LM_Ins2_L67_Playfield, LM_Ins2_L67_bulbs, LM_Ins2_L67_ins_mesh, LM_Ins2_L67_psw17, LM_Ins2_L67_psw18, LM_Ins2_L67_psw19, _
  LM_Ins2_L68_Layer_2, LM_Ins2_L68_Playfield, LM_Ins2_L68_bulbs, LM_Ins2_L68_ins_mesh, LM_Ins2_L69_Playfield, LM_Ins2_L69_bulbs, LM_Ins2_L69_ins_mesh, LM_Ins2_L71_Flip2, LM_Ins2_L71_Playfield, LM_Ins2_L71_bulbs, LM_Ins2_L71_ins_mesh, LM_Ins2_L83_Playfield, LM_Ins2_L83_bulbs, LM_Ins2_L83_ins_mesh, LM_Ins2_L83_psw17, LM_Ins2_L83_psw18, LM_Ins2_L83_psw19, LM_Ins2_L84_Playfield, LM_Ins2_L84_bulbs, LM_Ins2_L84_ins_mesh, LM_Ins2_L85_Playfield, LM_Ins2_L85_bulbs, LM_Ins2_L85_ins_mesh, LM_Ins2_L87_Flip1, LM_Ins2_L87_Playfield, LM_Ins2_L87_bulbs, LM_Ins2_L87_ins_mesh, LM_Ins2_L9_Parts, LM_Ins2_L9_Playfield, LM_Ins2_L9_bulbs, LM_Ins2_L9_ins_mesh, LM_Ins2_L98_bulbs, LM_Ins2_L98_ins_mesh, LM_Ins2_L99_Playfield, LM_Ins2_L99_bulbs, LM_Ins2_L99_ins_mesh, LM_Ins2_L99_psw17)
' VLM  Arrays - End



'******************************************************
' Timers
'******************************************************

Dim FrameTime, InitFrameTime
InitFrameTime = 0

' The frame timer interval is -1, so executes at the display frame rate
Sub FrameTimer_Timer()

  FrameTime = gametime - InitFrameTime 'Calculate FrameTime as some animations could use this
  InitFrameTime = gametime  'Count frametime

  'Add animation stuff here
  RollingUpdate         'update rolling sounds
  DoDTAnim            'handle drop target animations
  UpdateDropTargets
  DoSTAnim            'handle stand up target animations
  UpdateStandupTargets
  AnimateBumperSkirts

  BSUpdate 'update ambient ball shadows

    If VR_Room = 0 and cab_mode = 0 Then
        DisplayTimer
    End If

  If VR_Room=1 Then
    VRDisplayTimer
  End If

  LatchedLamp52 'added to control dt1 timing : Latching Controler.Lamp(52)
  LatchedSolBallRel 'added to control dt1 timing : Latching SolBallRelease
  LatchedSolReset 'added to control dt2 timing : Latching SolReset

End Sub

'The CorTimer interval should be 10. It's sole purpose is to update the Cor (physics) calculations
CorTimer.Interval = 10
Sub CorTimer_Timer(): Cor.Update: End Sub


'******************************************************
' Solenoid Call backs
'******************************************************

SolCallback(6) = "SolKnocker"
SolCallback(7) = "SolDT8ballReset"
SolCallback(8) = "SolSaucer"
SolCallback(9) = "SolBallRelease"
SolCallback(10) = "SolReset"
SolCallback(11) = "DTlower3"
SolCallback(12) = "DTlower4"
SolCallback(13) = "DTlower5"
SolCallback(14) = "DTlower6"
SolCallback(15) = "DTlower7"
SolCallback(19) = "Gameinplay"

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"


'******************************************************
' General Math Functions
'******************************************************
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
' Initiate Table
'******************************************************

Dim x

Dim EBall1, gBOT

 Sub Table1_Init
  vpmInit Me
  On Error Resume Next
    With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine = "Eight Ball Deluxe (Bally 1980)"&chr(13)&"by Bord and UnclePaulie"
    .HandleMechanics=0
    .HandleKeyboard=0
    .ShowDMDOnly=1
    .ShowFrame=0
    .ShowTitle=0
        .hidden = 1
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With
     On Error Goto 0

     PinMAMETimer.Interval = PinMAMEInterval
     PinMAMETimer.Enabled = 1

     vpmNudge.TiltSwitch = 7
     vpmNudge.Sensitivity = 3
     vpmNudge.TiltObj = Array(Bumper1,Bumper2,Bumper3, LeftSlingshot, RightSlingshot)

  'Ball initializations need for physical trough
  Set EBall1 = BallRelease.CreateSizedballWithMass(Ballsize/2,Ballmass)
  gBOT = Array(EBall1)

  Controller.Switch(8) = 0

  if VR_Room = 1 Then
    setup_backglass()
  End If

  VR_Primary_Backglass.blenddisablelighting = 1

  vpmMapLights insertlamps

' set slings to visible at startup

  Dim BP
  For each BP in BP_lsling: BP.visible = 1: Next
  For each BP in BP_lsling1: BP.visible = 0: Next
  For each BP in BP_lsling2: BP.visible = 0: Next
  For each BP in BP_rsling: BP.visible = 1: Next
  For each BP in BP_rsling1: BP.visible = 0: Next
  For each BP in BP_rsling2: BP.visible = 0: Next

' set flipper3 prims to new location.  (slight adjustment over hole and avoids ball roll bumpout... physical wall moved slightly too)
  For each BP in BP_Flip3:BP.x = 91.5: Next
  For each BP in BP_Flip3:BP.y = 1064.5: Next
  LeftFlipper1.x = 91.5
  LeftFlipper1.y = 1064.5

' If user wants to set dip switches themselves it will force them to set it via F6.
  If SetDIPSwitches = 0 Then
    SetDefaultDips
  End If

If vr_room = 1 Then ' F12 menu only working for desktop and cab.  VR mode requires the setup be done here.
  SetupOptions ' enable all options above
End if

End Sub

Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub
Sub table1_exit
  Controller.Pause = False
  Controller.stop
End Sub


'******************************************************
' Misc Animations
'******************************************************

''''' Flippers

Sub LeftFlipper_Animate
  dim a: a = LeftFlipper.CurrentAngle
  FlipperLSh.RotZ = a
  'Add any left flipper related animations here
  Dim BP
  For each BP in BP_Flip2:BP.rotz = a -174 : Next
End Sub

Sub RightFlipper_Animate
  dim a: a = RightFlipper.CurrentAngle
  FlipperRSh.RotZ = a
  'Add any right flipper related animations here
  Dim BP
  For each BP in BP_Flip1:BP.rotz = a + 186 : Next
End Sub

Sub LeftFlipper1_Animate
    dim a: a = LeftFlipper1.CurrentAngle
    FlipperLSh1.RotZ = a
  'Add any right flipper related animations here
  Dim BP
  For each BP in BP_Flip3:BP.rotz = a - 174: Next
End Sub



''''' Drop Targets

Sub UpdateDropTargets
  dim BP, tz, rx, ry

  tz = BM_psw1.transz
  rx = BM_psw1.rotx
  ry = BM_psw1.roty
  For each BP in BP_psw1: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_psw2.transz
  rx = BM_psw2.rotx
  ry = BM_psw2.roty
  For each BP in BP_psw2: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_psw3.transz
  rx = BM_psw3.rotx
  ry = BM_psw3.roty
  For each BP in BP_psw3: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_psw4.transz
  rx = BM_psw4.rotx
  ry = BM_psw4.roty
  For each BP in BP_psw4: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_psw17.transz
  rx = BM_psw17.rotx
  ry = BM_psw17.roty
  For each BP in BP_psw17: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_psw18.transz
  rx = BM_psw18.rotx
  ry = BM_psw18.roty
  For each BP in BP_psw18: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_psw19.transz
  rx = BM_psw19.rotx
  ry = BM_psw19.roty
  For each BP in BP_psw19: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_psw20.transz
  rx = BM_psw20.rotx
  ry = BM_psw20.roty
  For each BP in BP_psw20: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_psw21.transz
  rx = BM_psw21.rotx
  ry = BM_psw21.roty
  For each BP in BP_psw21: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_psw22.transz
  rx = BM_psw22.rotx
  ry = BM_psw22.roty
  For each BP in BP_psw22: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_psw23.transz
  rx = BM_psw23.rotx
  ry = BM_psw23.roty
  For each BP in BP_psw23: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_psw33.transz
  rx = BM_psw33.rotx
  ry = BM_psw33.roty
  For each BP in BP_psw33: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

End Sub


''''' Standup Targets

Sub UpdateStandupTargets
  dim BP, ty

  ty = BM_tsw5.transy
  For each BP in BP_tsw5 : BP.transy = ty: Next

  ty = BM_tsw25.transy
  For each BP in BP_tsw25 : BP.transy = ty: Next

  ty = BM_tsw26.transy
  For each BP in BP_tsw26 : BP.transy = ty: Next

  ty = BM_tsw27.transy
  For each BP in BP_tsw27 : BP.transy = ty: Next

  ty = BM_tsw28.transy
  For each BP in BP_tsw28 : BP.transy = ty: Next

  ty = BM_tsw29.transy
  For each BP in BP_tsw29 : BP.transy = ty: Next

  ty = BM_tsw30.transy
  For each BP in BP_tsw30 : BP.transy = ty: Next

End Sub



''''' Bumper Animations

Sub Bumper1_Animate
  Dim z, BP
  z = Bumper1.CurrentRingOffset
  For Each BP in BP_BumperRing1 : BP.transz = z: Next
End Sub

Sub Bumper2_Animate
  Dim z, BP
  z = Bumper2.CurrentRingOffset
  For Each BP in BP_BumperRing2 : BP.transz = z: Next
End Sub

Sub Bumper3_Animate
  Dim z, BP
  z = Bumper3.CurrentRingOffset
  For Each BP in BP_BumperRing3 : BP.transz = z: Next
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
      tz = -3
    End If
    If r = 0 Then For Each x in BP_BumperSkirt1: x.transZ = tz: Next
    If r = 1 Then For Each x in BP_BumperSkirt2: x.transZ = tz: Next
    If r = 2 Then For Each x in BP_BumperSkirt3: x.transZ = tz: Next
  Next
End Sub


''''' Gate Animations

Sub Gate2_Animate
  Dim a : a = (Gate2.CurrentAngle * -0.5) + 15
  Dim BP
    For Each BP in BP_Gateplate2
      BP.rotx = a
      BP.z = 71.5
    Next
End Sub

Sub Gate3_Animate
  Dim a : a = Gate3.CurrentAngle
  Dim BP
    For Each BP in BP_gatewire3
      BP.rotx = a
      BP.z = 60
    Next
End Sub

Sub Gate4_Animate
  Dim a : a = Gate4.CurrentAngle
    For Each BP in BP_gatewire4
      BP.rotx = a
      BP.z = 60
    Next
End Sub

Sub Gate5_Animate
  Dim a : a = (Gate5.CurrentAngle * -0.5) + 15
  Dim BP
    For Each BP in BP_Gateplate5
      BP.rotx = a
      BP.z = 71.5
    Next
End Sub


''''' Star Trigger Animations

Sub sw35_Animate
  Dim z : z = sw35.CurrentAnimOffset
  Dim BP : For Each BP in BP_ROstar : BP.transz = z:Next
End Sub

''''' Rollover Trigger Animations

Sub sw12_Animate
  Dim z : z = sw12.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw12 : BP.transz = z:Next
End Sub

Sub sw13_Animate
  Dim z : z = sw13.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw13 : BP.transz = z:Next
End Sub

Sub sw14_Animate
  Dim z : z = sw14.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw14 : BP.transz = z:Next
End Sub

Sub sw15_Animate
  Dim z : z = sw15.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw15 : BP.transz = z:Next
End Sub

Sub sw31_Animate
  Dim z : z = sw31.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw31 : BP.transz = z:Next
End Sub

Sub sw32_Animate
  Dim z : z = sw32.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw32 : BP.transz = z:Next
End Sub


'******************************************************
' Drain, Trough, and Ball Release Code
'******************************************************

Sub Drain_Hit()
  Controller.Switch(8) = 1
  RandomSoundDrain Drain
  UpdateTrough
End Sub

Sub Drain_UnHit()
  Controller.Switch(8) = 0
End Sub

Sub KickOutOfOuthole()
  BallRelease.kick 60, 14
End Sub

Sub UpdateTrough
  UpdateTroughTimer.Interval = 300
  UpdateTroughTimer.Enabled = 1
End Sub

Sub UpdateTroughTimer_Timer
  If Drain.BallCntOver = 1 Then Drain.kick 60, 6
  Me.Enabled = 0
End Sub

Sub BIPL_Trigger_Hit
  BIPL=True
End Sub

Sub BIPL_Trigger_UnHit
  BIPL=False
End Sub

'******************************************************
' Saucer
'******************************************************

Sub KickBall(kball, kangle, kvel, kvelz, kzlift)
  dim rangle
  rangle = PI * (kangle - 90) / 180
  kball.z = kball.z + kzlift
  kball.velz = kvelz
  kball.velx = cos(rangle)*kvel
  kball.vely = sin(rangle)*kvel
End Sub

Dim KickerBall

dim kickstep1, kickstep2

Sub sw34_Hit
  set KickerBall = activeball
  Controller.Switch(34) = 1
  SoundSaucerLock
End Sub

Sub sw34_unHit
  Controller.Switch(34) = 0
End Sub

Sub sw34_timer
    Select Case kickstep2
    Case 9:sw34.TimerEnabled = 0: if sw34.ballcntover > 0 then SolSaucer -1
    End Select
   kickstep2 = kickstep2 + 1
End Sub


'******************************************************
' Solenoid Controlled toys
'******************************************************

dim BallRelLatched: BallRelLatched =0
Dim SolResetLatched: SolResetLatched = 0
dim Lamp52On: Lamp52On = 0
dim DT17Dropped: DT17Dropped = 0
dim DT18Dropped: DT18Dropped = 0

Sub SolBallRelease(Enabled)
  if Enabled Then
    BallRelLatched = 1
  End If
End Sub

sub LatchedLamp52()
  if controller.lamp(52) Then
    Lamp52On = 1
  end If
end Sub

Sub LatchedSolBallRel()
  If Lamp52On  = 1 Then
    If BallRelLatched = 1 Then
      KickOutOfOuthole
    End If
  Else
    If BallRelLatched = 1 Then
      If DT17Dropped = 0 Then
        DTDrop 17
        DT17Dropped = 1
      End If
    End If
  End If
End Sub

Sub SolSaucer(Enable)
  If Enable then
    If Lamp52On = 1 Then
      kickstep2 = 0
      If controller.switch(34) then
        KickBall KickerBall, 330, 18, 6, 26
        SoundSaucerKick 1, sw34
      Else
        SoundSaucerKick 0, sw34
      End If
      sw34.timerenabled = 1
    Else
      dt7reset
    End If
  End If
End Sub

Sub BallRelSoundOn_Hit
  RandomSoundBallRelease BallRelease
End Sub

Sub ResetLatches_Hit
  BallRelLatched = 0
  SolResetLatched = 0
  Lamp52On = 0
  DT17Dropped = 0
  DT18Dropped = 0
End Sub

Sub DTLower3(Enabled)
  If Enabled Then
    DTDrop 19
  End If
End Sub

Sub DTLower4(Enabled)
  If Enabled Then
    DTDrop 20
  End If
End Sub

Sub DTLower5(Enabled)
  If Enabled Then
    DTDrop 21
  End If
End Sub

Sub DTLower6(Enabled)
  If Enabled Then
    DTDrop 22
  End If
End Sub

Sub DTLower7(Enabled)
  If Enabled Then
    DTDrop 23
  End If
End Sub


'******************************************************
' GI Lights On
'******************************************************
dim xx
For each xx in GI:xx.State = 1: Next


'******************************************************
' VR Plunger code
'******************************************************

Sub TimerVRPlunger_Timer
  If VR_Primary_plunger.Y < 2193 then
      VR_Primary_plunger.Y = VR_Primary_plunger.Y + 5
  End If
End Sub

Sub TimerVRPlunger1_Timer
  VR_Primary_plunger.Y = 2103 + (5* Plunger.Position) -20
End Sub


'******************************************************
' Keys
'******************************************************

Sub Table1_KeyDown(ByVal KeyCode)

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
    VR_Primary_plunger.Y = 2103
  End If

  If keycode = LeftFlipperKey Then
    VRFlipperButtonLeft.X = 2115.435 + 8
  End If
  If keycode = RightFlipperKey Then
    VRFlipperButtonRight.X = 2098.233 - 8
  End If

  If keycode = StartGameKey Then
    VR_StartButton.y = 1918 - 5
    SoundStartButton
  End If

  If KeyDownHandler(keycode) Then Exit Sub

End Sub

Sub Table1_KeyUp(ByVal KeyCode)

  If keycode = PlungerKey Then
    Plunger.Fire
    If BIPL = 1 Then
      SoundPlungerReleaseBall()     'Plunger release sound when there is a ball in shooter lane
    Else
      SoundPlungerReleaseNoBall()     'Plunger release sound when there is no ball in shooter lane
    End If
    TimerVRPlunger.Enabled = False
        TimerVRPlunger1.Enabled = True
    VR_Primary_plunger.Y = 2103
  End If

  If keycode = LeftFlipperKey Then
    VRFlipperButtonLeft.X = 2115.435
  End If

  If keycode = RightFlipperKey Then
    VRFlipperButtonRight.X = 2098.233
  End If

  If keycode = StartGameKey Then
    VR_StartButton.y = 1918
  End If

  If KeyUpHandler(keycode) Then Exit Sub

End Sub




'****************************
'   ZRBR: Room Brightness
'****************************

'This code only applies to lightmapped tables. It is here for reference.
'NOTE: Objects bightness will be affected by the Day/Night slider only if their blenddisablelighting property is less than 1.
'      Lightmapped table primitives have their blenddisablelighting equal to 1, therefore we need this SetRoomBrightness sub
'      to handle updating their effective ambient brighness.

' Update these arrays if you want to change more materials with room light level
Dim RoomBrightnessMtlArray: RoomBrightnessMtlArray = Array("VLM.Bake.Active","VLM.Bake.Solid")

Sub SetRoomBrightness(lvl)
  If lvl > 1 Then lvl = 1
  If lvl < 0 Then lvl = 0

  ' Lighting level
  Dim v: v=(lvl * 245 + 10)/255

  Dim i: For i = 0 to UBound(RoomBrightnessMtlArray)
    ModulateMaterialBaseColor RoomBrightnessMtlArray(i), i, v
  Next
End Sub

Dim SavedMtlColorArray
SaveMtlColors
Sub SaveMtlColors
  ReDim SavedMtlColorArray(UBound(RoomBrightnessMtlArray))
  Dim i: For i = 0 to UBound(RoomBrightnessMtlArray)
    SaveMaterialBaseColor RoomBrightnessMtlArray(i), i
  Next
End Sub

Sub SaveMaterialBaseColor(name, idx)
  Dim wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle
  GetMaterial name, wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle
  SavedMtlColorArray(idx) = round(base,0)
End Sub


Sub ModulateMaterialBaseColor(name, idx, val)
  Dim wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle
  Dim red, green, blue, saved_base, new_base

  'First get the existing material properties
  GetMaterial name, wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle

  'Get saved color
  saved_base = SavedMtlColorArray(idx)

  'Next extract the r,g,b values from the base color
  red = saved_base And &HFF
  green = (saved_base \ &H100) And &HFF
  blue = (saved_base \ &H10000) And &HFF
  'msgbox red & " " & green & " " & blue

  'Create new color scaled down by 'val', and update the material
  new_base = RGB(red*val, green*val, blue*val)
  UpdateMaterial name, wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, new_base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle
End Sub




'******************************************************
' Flippers
'******************************************************

Const ReflipAngle = 20

' Flipper Solenoid Callbacks (these subs mimics how you would handle flippers in ROM based tables)
Sub SolLFlipper(Enabled)
  If Enabled Then
    FlipperActivate LeftFlipper, LFPress
    LF.Fire  'leftflipper.rotatetoend
    LeftFlipper1.rotatetoend
    If leftflipper.currentangle < leftflipper.endangle + ReflipAngle Then
      RandomSoundReflipUpLeft LeftFlipper
    Else
      SoundFlipperUpAttackLeft LeftFlipper
      RandomSoundFlipperUpLeft LeftFlipper
    End If
  Else
    FlipperDeActivate LeftFlipper, LFPress
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


'******************************************************
' Drop Targets
'******************************************************

Sub sw1_Hit
  DTHit 1
End Sub

Sub sw2_Hit
  DTHit 2
End Sub

Sub sw3_Hit
  DTHit 3
End Sub

Sub sw4_Hit
  DTHit 4
End Sub

Sub sw17_Hit
  DTHit 17
End Sub

Sub sw18_Hit
  DTHit 18
End Sub

Sub sw19_Hit
  DTHit 19
End Sub

Sub sw20_Hit
  DTHit 20
End Sub

Sub sw21_Hit
  DTHit 21
End Sub

Sub sw22_Hit
  DTHit 22
End Sub

Sub sw23_Hit
  DTHit 23
End Sub

Sub sw33_Hit
  DTHit 33
End Sub


' **** Reset Drop Targets by Solenoid callout

Sub SolDT8ballReset(enabled)
  dim xx
  if enabled then
    RandomSoundDropTargetReset BM_psw33
    DTRaise 33
  end if
End Sub



Sub SolReset(Enabled)
  If Enabled Then
    SolResetLatched = 1
    dt4reset
  End If
End Sub


Sub LatchedSolReset()
  If Lamp52On = 0 then
    If SolResetLatched = 1 Then
      If DT18Dropped = 0 Then
        DTDrop 18
        DT18Dropped = 1
        SolResetLatched = 0
      End If
    End If
  End If
End Sub



Sub dt7reset()
  RandomSoundDropTargetReset BM_psw20
  DTRaise 17
  DTRaise 18
  DTRaise 19
  DTRaise 20
  DTRaise 21
  DTRaise 22
  DTRaise 23
End Sub

Sub dt4reset()
  RandomSoundDropTargetReset BM_psw1
  DTRaise 1
  DTRaise 2
  DTRaise 3
  DTRaise 4
End Sub


'******************************************************
' Rollovers
'******************************************************

 Sub sw12_Hit:Controller.Switch(12) = 1 : Lamp52On = 0: End Sub  'Remove LatchedLamp52
 Sub sw12_UnHit:Controller.Switch(12) = 0:End Sub
 Sub sw13_Hit:Controller.Switch(13) = 1 : Lamp52On = 0: End Sub  'Remove LatchedLamp52
 Sub sw13_UnHit:Controller.Switch(13) = 0:End Sub
 Sub sw14_Hit:Controller.Switch(14) = 1 : End Sub
 Sub sw14_UnHit:Controller.Switch(14) = 0:End Sub
 Sub sw15_Hit:Controller.Switch(15) = 1 : End Sub
 Sub sw15_UnHit:Controller.Switch(15) = 0:End Sub
 Sub sw31_Hit:Controller.Switch(31) = 1 : End Sub
 Sub sw31_UnHit:Controller.Switch(31) = 0:End Sub
 Sub sw32_Hit:Controller.Switch(32) = 1 : End Sub
 Sub sw32_UnHit:Controller.Switch(32) = 0:End Sub


'******************************************************
' Bumpers
'******************************************************

Sub Bumper1_Hit : vpmTimer.PulseSw(38) : RandomSoundBumperTop Bumper1: End Sub
Sub Bumper2_Hit : vpmTimer.PulseSw(39) : RandomSoundBumperTop Bumper2: End Sub
Sub Bumper3_Hit : vpmTimer.PulseSw(40) : RandomSoundBumperBottom Bumper3: End Sub


'******************************************************
' Standup Targets
'******************************************************

Sub sw5_Hit
  STHit 5
End Sub

Sub sw5o_Hit
  TargetBouncer Activeball, 1
End Sub

Sub sw25_Hit
  STHit 25
End Sub

Sub sw25o_Hit
  TargetBouncer Activeball, 1
End Sub

Sub sw26_Hit
  STHit 26
End Sub

Sub sw26o_Hit
  TargetBouncer Activeball, 1
End Sub

Sub sw27_Hit
  STHit 27
End Sub

Sub sw27o_Hit
  TargetBouncer Activeball, 1
End Sub

Sub sw28_Hit
  STHit 28
End Sub

Sub sw28o_Hit
  TargetBouncer Activeball, 1
End Sub

Sub sw29_Hit
  STHit 29
End Sub

Sub sw29o_Hit
  TargetBouncer Activeball, 1
End Sub

Sub sw30_Hit
  STHit 30
End Sub

Sub sw30o_Hit
  TargetBouncer Activeball, 1
End Sub


'******************************************************
' Star Trigger
'******************************************************

 Sub sw35_Hit:Controller.Switch(35) = 1 : End Sub
 Sub sw35_UnHit:Controller.Switch(35) = 0:End Sub


'******************************************************
' Knocker
'******************************************************

Sub SolKnocker(Enabled)
  If enabled Then
    KnockerSolenoid
  End If
End Sub


'******************************************************
' Upper leaf sensor switches
'******************************************************

Sub sw24001_hit
  vpmTimer.PulseSw 24
End Sub


'******************************************************
' Digital Display
'******************************************************

Dim Digits(32)
' 1st Player
Digits(0) = Array(LED10,LED11,LED12,LED13,LED14,LED15,LED16)
Digits(1) = Array(LED20,LED21,LED22,LED23,LED24,LED25,LED26)
Digits(2) = Array(LED30,LED31,LED32,LED33,LED34,LED35,LED36)
Digits(3) = Array(LED40,LED41,LED42,LED43,LED44,LED45,LED46)
Digits(4) = Array(LED50,LED51,LED52,LED53,LED54,LED55,LED56)
Digits(5) = Array(LED60,LED61,LED62,LED63,LED64,LED65,LED66)
Digits(6) = Array(LED70,LED71,LED72,LED73,LED74,LED75,LED76)

' 2nd Player
Digits(7) = Array(LED80,LED81,LED82,LED83,LED84,LED85,LED86)
Digits(8) = Array(LED90,LED91,LED92,LED93,LED94,LED95,LED96)
Digits(9) = Array(LED100,LED101,LED102,LED103,LED104,LED105,LED106)
Digits(10) = Array(LED110,LED111,LED112,LED113,LED114,LED115,LED116)
Digits(11) = Array(LED120,LED121,LED122,LED123,LED124,LED125,LED126)
Digits(12) = Array(LED130,LED131,LED132,LED133,LED134,LED135,LED136)
Digits(13) = Array(LED140,LED141,LED142,LED143,LED144,LED145,LED146)

' 3rd Player
Digits(14) = Array(LED150,LED151,LED152,LED153,LED154,LED155,LED156)
Digits(15) = Array(LED160,LED161,LED162,LED163,LED164,LED165,LED166)
Digits(16) = Array(LED170,LED171,LED172,LED173,LED174,LED175,LED176)
Digits(17) = Array(LED180,LED181,LED182,LED183,LED184,LED185,LED186)
Digits(18) = Array(LED190,LED191,LED192,LED193,LED194,LED195,LED196)
Digits(19) = Array(LED200,LED201,LED202,LED203,LED204,LED205,LED206)
Digits(20) = Array(LED210,LED211,LED212,LED213,LED214,LED215,LED216)

' 4th Player
Digits(21) = Array(LED220,LED221,LED222,LED223,LED224,LED225,LED226)
Digits(22) = Array(LED230,LED231,LED232,LED233,LED234,LED235,LED236)
Digits(23) = Array(LED240,LED241,LED242,LED243,LED244,LED245,LED246)
Digits(24) = Array(LED250,LED251,LED252,LED253,LED254,LED255,LED256)
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

Dim VRDigits(32)
' 1st Player
VRDigits(0) = Array(ax00,ax05,ax0c,ax0d,ax08,ax01,ax06)
VRDigits(1) = Array(ax10,ax15,ax1c,ax1d,ax18,ax11,ax16)
VRDigits(2) = Array(ax20,ax25,ax2c,ax2d,ax28,ax21,ax26)
VRDigits(3) = Array(ax30,ax35,ax3c,ax3d,ax38,ax31,ax36)
VRDigits(4) = Array(ax40,ax45,ax4c,ax4d,ax48,ax41,ax46)
VRDigits(5) = Array(ax50,ax55,ax5c,ax5d,ax58,ax51,ax56)
VRDigits(6) = Array(ax60,ax65,ax6c,ax6d,ax68,ax61,ax66)
' 2nd Player
VRDigits(7) = Array(ax70,ax75,ax7c,ax7d,ax78,ax71,ax76)
VRDigits(8) = Array(ax80,ax85,ax8c,ax8d,ax88,ax81,ax86)
VRDigits(9) = Array(ax90,ax95,ax9c,ax9d,ax98,ax91,ax96)
VRDigits(10) = Array(ax100,ax105,ax10c,ax10d,ax108,ax101,ax106)
VRDigits(11) = Array(ax110,ax115,ax11c,ax11d,ax118,ax111,ax116)
VRDigits(12) = Array(ax120,ax125,ax12c,ax12d,ax128,ax121,ax126)
VRDigits(13) = Array(ax130,ax135,ax13c,ax13d,ax138,ax131,ax136)

' 3rd Player
VRDigits(14) = Array(ax140,ax145,ax14c,ax14d,ax148,ax141,ax146)
VRDigits(15) = Array(ax150,ax155,ax15c,ax15d,ax158,ax151,ax156)
VRDigits(16) = Array(ax160,ax165,ax16c,ax16d,ax168,ax161,ax166)
VRDigits(17) = Array(ax170,ax175,ax17c,ax17d,ax178,ax171,ax176)
VRDigits(18) = Array(ax180,ax185,ax18c,ax18d,ax188,ax181,ax186)
VRDigits(19) = Array(ax190,ax195,ax19c,ax19d,ax198,ax191,ax196)
VRDigits(20) = Array(ax200,ax205,ax20c,ax20d,ax208,ax201,ax206)

' 4th Player
VRDigits(21) = Array(ax210,ax215,ax21c,ax21d,ax218,ax211,ax216)
VRDigits(22) = Array(ax220,ax225,ax22c,ax22d,ax228,ax221,ax226)
VRDigits(23) = Array(ax230,ax235,ax23c,ax23d,ax238,ax231,ax236)
VRDigits(24) = Array(ax240,ax245,ax24c,ax24d,ax248,ax241,ax246)
VRDigits(25) = Array(ax250,ax255,ax25c,ax25d,ax258,ax251,ax256)
VRDigits(26) = Array(ax260,ax265,ax26c,ax26d,ax268,ax261,ax266)
VRDigits(27) = Array(ax270,ax275,ax27c,ax27d,ax278,ax271,ax276)

' Credits
VRDigits(28) = Array(ax280,ax285,ax28c,ax28d,ax288,ax281,ax286)
VRDigits(29) = Array(ax290,ax295,ax29c,ax29d,ax298,ax291,ax296)
' Balls
VRDigits(30) = Array(ax300,ax305,ax30c,ax30d,ax308,ax301,ax306)
VRDigits(31) = Array(ax310,ax315,ax31c,ax31d,ax318,ax311,ax316)


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
    Object.Opacity = 12
  Else
    Object.Color = RGB(5,10,10)
    Object.Opacity = 4
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



'******************************************************
' Setup Backglass for VR
'******************************************************

Dim xoff,yoff,zoff,xrot,zscale, xcen,ycen

Sub setup_backglass()

  Dim ix, xx, yy, yfact, xfact, xobj
  xoff =475
  yoff =142
  zoff =990
  xrot = -90

  zscale = -.4

  xcen =(1017 /2) - (80 / 2)
  ycen = (1786 /2 ) + (325 /2)

  yfact =0 'y fudge factor (ycen was wrong so fix)
  xfact = 1


  for ix =0 to 31
  For Each xobj In VRDigits(ix)

  xx =xobj.x

  xobj.x = (xoff -xcen) + xx +xfact
  yy = xobj.y ' get the yoffset before it is changed
  xobj.y =yoff
    xobj.y =xobj.y - 38

    If(yy < 0.) then
    yy = yy * -1
    end if

  xobj.height =( zoff - ycen) + yy - (yy * (zscale)) + yfact

  xobj.rotx = xrot

  Next
  Next
end sub


'******************************************************
' DIP Switches
'******************************************************

Sub editDips
  if SetDIPSwitches = 1 Then
    Dim vpmDips : Set vpmDips = New cvpmDips
    With vpmDips
      .AddForm 315,450,"Eight Ball Deluxe - DIP switches"
      .AddFrame 2,10,140,"Balls per game",&HC0000000,Array("2",&HC0000000,"3",0,"4",&H80000000,"5",&H40000000)'dip 31&32
      .AddFrame 2,90,140,"C & D lanes are",&H00000040,Array("separate",0,"linked",&H00000040)'dip 7
      .AddFrame 2,140,140,"Saucer collects bonus",32768,Array("without multiplier",0,"with multiplier",32768)'dip 16
      .AddFrame 2,190,140,"Replays awarded per game",&H10000000,Array("1 per player per game",0,"unlimited replays",&H10000000)'dip 29
      .AddFrame 2,240,140,"Backbox Deluxe advances",&H00200000,Array("when special is lit",0,"every time",&H00200000)'dip 22
      .AddFrame 2,290,298,"Left lane score sequence",&H00002000,Array("no lite, 10K, 30K, 50K, X-ball, 70K, special, 70K stays on",&H00002000,"1st ball - no special lite, next ball - no extra ball lite",0)'dip 14
      .AddChk 7,345,190,Array("Drop targets not reset after Deluxe",&H00400000)'dip 23
      .AddChk 7,365,160,Array("Playfield Deluxe memory",&H00100000)'dip 21
      .AddChk 7,385,160,Array("Voice on during attact mode",&H20000000)'dip 30
      .AddFrame 160,10,140,"Maximum credits",&H03000000,Array("10",0,"15",&H01000000,"25",&H02000000,"40",&H03000000)'dip 25&26
      .AddFrame 160,90,140,"Making A-B-C-D spots",&H00000080,Array("1 target",0,"2 targets",&H00000080)'dip 8
      .AddFrame 160,140,140,"Target && Deluxe specials",&H00000020,Array("scores only once per ball",0,"alternates with 50k",&H00000020)'dip 6
      .AddFrame 160,190,140,"Eight Ball scores special on",&H00004000,Array("third time",0,"second and third time",&H00004000)'dip 15
      .AddFrame 160,240,140,"Deluxe special scored",&H00800000,Array("after 50K",0,"before 50K",&H00800000)'dip 24
      .AddChk 210,345,90,Array("Credit display",&H04000000)'dip 27
      .AddChk 210,365,90,Array("Match",&H08000000)'dip 28
      .AddLabel 7,410,308,20,"Set selftest position 16,17,18 and 19 to 03 for the best gameplay."
      .AddLabel 7,430,300,20,"After hitting OK, press F3 to reset game with new settings."
      .ViewDips
    End With
  End If
End Sub
Set vpmShowDips = GetRef("editDips")


'  *** This is the default dips settings ***

' 1-5   00000   1 credit per coin
' 6   1   50k and special alternate (0 only 1 special per ball)
' 7   0   C and D are separate (1 they are linked)
' 8   0   ABCD awards 1 target down (1 is 2 targets)

' 9-13  00000   1 credit per coin
' 14    0   conservate light sequence (1 is liberal)
' 15    0   8 ball special conservative (3rd time gets replay, 1=liberal and 2 and 3rd times get replay)
' 16    0   saucer gets rack and scores values only - no multiplier.  liberal(1) gets that and gets multiplier

' 17-20   0000  1 credit per coin and same as coin chute
' 21    1   playfield deluxe memory.  1 keeps lights on after 8 ball for next ball
' 22    1   Backbox deluxe advances 1 step each time deluxe is done. (0 special also has to be lit)
' 23    0   8ball and deluxe resets 7 drop targets (1 does not reset)
' 24    1   Deluxe special. 1 - before 50k.  0 after 50

' 25-26 00    10 max credits
' 28    1   match
' 27    1   credit displayed
' 29    0   0 = 1 replay per player per game; 1 = all replays earned will be collected
' 30    1   1=voice; 0=no voice
' 31-32 00    3 balls per game



Sub SetDefaultDips
  If SetDIPSwitches = 0 Then
    Controller.Dip(0) = 32    '8-1 = 00100000
    Controller.Dip(1) = 0   '16-9 = 00000000
    Controller.Dip(2) = 176   '24-17 = 10110000
    Controller.Dip(3) = 44    '32-25 = 00101100
    Controller.Dip(4) = 0
    Controller.Dip(5) = 0
  End If
End Sub


'******************************************************
' Sling Shot Animations
'******************************************************

Dim RStep, Lstep

Sub RightSlingShot_Slingshot
  Dim BP
  RS.VelocityCorrect(ActiveBall)
  vpmTimer.PulseSw 36
  RandomSoundSlingshotRight BM_slingarmR
  For each BP in BP_rsling: BP.visible = 0: Next
  For each BP in BP_rsling1: BP.visible = 1: Next
  For each BP in BP_slingarmR: BP.TransZ = -20: Next
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
  Dim BP
    Select Case RStep
        Case 3
      For each BP in BP_rsling1: BP.visible = 0: Next
      For each BP in BP_rsling2: BP.visible = 1: Next
      For each BP in BP_slingarmR: BP.TransZ = -10: Next
    Case 4
      For each BP in BP_rsling2: BP.visible = 0: Next
      For each BP in BP_rsling: BP.visible = 1: Next
      For each BP in BP_slingarmR: BP.TransZ = 0: Next
      RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  Dim BP
  LS.VelocityCorrect(ActiveBall)
  vpmTimer.PulseSw 37
  RandomSoundSlingshotLeft BM_slingarmL
  For each BP in BP_lsling: BP.visible = 0: Next
  For each BP in BP_lsling1: BP.visible = 1: Next
  For each BP in BP_slingarmL: BP.TransZ = -20: Next
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
  Dim BP
    Select Case LStep
    Case 3
      For each BP in BP_lsling1: BP.visible = 0: Next
      For each BP in BP_lsling2: BP.visible = 1: Next
      For each BP in BP_slingarmL: BP.TransZ = -10: Next
    Case 4
      For each BP in BP_lsling2: BP.visible = 0: Next
      For each BP in BP_lsling: BP.visible = 1: Next
      For each BP in BP_slingarmL: BP.TransZ = 0: Next
      LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub




'******************************************************
' Ambient ball shadows
'******************************************************

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
    If gBOT(s).Z > 20 Then
      objBallShadow(s).visible = 1
      objBallShadow(s).X = gBOT(s).X + (gBOT(s).X - (tablewidth / 2)) / (Ballsize / AmbientMovement) + offsetX
      objBallShadow(s).Y = gBOT(s).Y + offsetY
      'objBallShadow(s).Z = gBOT(s).Z + s/1000 + 1.04 - 25

    '** Under pf, no shadow
    Else
      objBallShadow(s).visible = 0
    End If
  Next
End Sub

'******************************************************
' VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
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

'******************************************************
' END VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************



'******************************************************
' Flippers Polarity (Select appropriate sub based on era)
'******************************************************

dim LF : Set LF = New FlipperPolarity
dim RF : Set RF = New FlipperPolarity

InitPolarity


''*******************************************
'' Late 70's to early 80's
'
Sub InitPolarity()
        dim x, a : a = Array(LF, RF)
  For Each x In a
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
' FLIPPER CORRECTION FUNCTIONS
'******************************************************

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
' SLINGSHOT CORRECTION FUNCTIONS
'******************************************************

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
' FLIPPER POLARITY AND RUBBER DAMPENER SUPPORTING FUNCTIONS
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
' FLIPPER TRICKS
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
  '   Dim BOT
  '   BOT = GetBalls

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



'******************************************************
' Check ball distance from Flipper for Rem
'******************************************************

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

'******************************************************
' End - Check ball distance from Flipper for Rem
'******************************************************

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
'Const EOSReturn = 0.035  'mid 80's to early 90's
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
' END FLIPPER CORRECTIONS
'******************************************************



'******************************************************
' RUBBER  DAMPENERS
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
' TRACK ALL BALL VELOCITIES, FOR RUBBER DAMPENER AND DROP TARGETS
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
' END PHYSICS DAMPENERS
'******************************************************



'******************************************************
' DROP TARGETS by Rothbauerw
'******************************************************

' DROP TARGETS INITIALIZATION

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

Dim DT1, DT2, DT3, DT4, DT17, DT18, DT19, DT20, DT21, DT22, DT23, DT33

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
'         Use the function DTDropped(switchid) to check a target's drop status.

Set DT1 = (new DropTarget)(sw1, sw1a, BM_psw1, 1, 0, false)
Set DT2 = (new DropTarget)(sw2, sw2a, BM_psw2, 2, 0, false)
Set DT3 = (new DropTarget)(sw3, sw3a, Bm_psw3, 3, 0, false)
Set DT4 = (new DropTarget)(sw4, sw4a, BM_psw4, 4, 0, false)
Set DT17 = (new DropTarget)(sw17, sw17a, BM_psw17, 17, 0, false)
Set DT18 = (new DropTarget)(sw18, sw18a, BM_psw18, 18, 0, false)
Set DT19 = (new DropTarget)(sw19, sw19a, BM_psw19, 19, 0, false)
Set DT20 = (new DropTarget)(sw20, sw20a, BM_psw20, 20, 0, false)
Set DT21 = (new DropTarget)(sw21, sw21a, BM_psw21, 21, 0, false)
Set DT22 = (new DropTarget)(sw22, sw22a, BM_psw22, 22, 0, false)
Set DT23 = (new DropTarget)(sw23, sw23a, BM_psw23, 23, 0, false)
Set DT33 = (new DropTarget)(sw33, sw33a, BM_psw33, 33, 0, false)


Dim DTArray
DTArray = Array(DT1, DT2, DT3, DT4, DT17, DT18, DT19, DT20, DT21, DT22, DT23, DT33)

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
Const DTHitSound = "" 'Drop Target Hit sound
Const DTDropSound = "DropTarget_Down" 'Drop Target Drop sound
Const DTResetSound = "DropTarget_Up" 'Drop Target reset sound

Const DTMass = 0.2 'Mass of the Drop Target (between 0 and 1), higher values provide more resistance


'******************************************************
' DROP TARGETS FUNCTIONS
'******************************************************

Sub DTHit(switch)
  Dim i
  i = DTArrayID(switch)

  PlayTargetSound
  DTArray(i).animate = DTCheckBrick(ActiveBall,DTArray(i).prim)
  If DTArray(i).animate = 1 Or DTArray(i).animate = 3 Or DTArray(i).animate = 4 Then
    DTBallPhysics ActiveBall, DTArray(i).prim.rotz, DTMass
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
' END DROP TARGETS
'******************************************************




'******************************************************
' STAND-UP TARGETS
'******************************************************

' STAND-UP TARGET INITIALIZATION

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
Dim ST5, ST25, ST26, ST27, ST28, ST29, ST30

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
'You will also need to add a secondary hit object for each stand up
'these are inclined primitives to simulate hitting a bent target and should provide so z velocity on high speed impacts

Set ST5 = (new StandupTarget)(sw5, BM_tsw5,5, 0)
Set ST25 = (new StandupTarget)(sw25, BM_tsw25,25, 0)
Set ST26 = (new StandupTarget)(sw26, BM_tsw26,26, 0)
Set ST27 = (new StandupTarget)(sw27, BM_tsw27,27, 0)
Set ST28 = (new StandupTarget)(sw28, BM_tsw28,28, 0)
Set ST29 = (new StandupTarget)(sw29, BM_tsw29,29, 0)
Set ST30 = (new StandupTarget)(sw30, BM_tsw30,30, 0)

'Add all the Stand-up Target Arrays to Stand-up Target Animation Array
' STAnimationArray = Array(ST1, ST2, ....)

Dim STArray
STArray = Array(ST5, ST25, ST26, ST27, ST28, ST29, ST30)

'Configure the behavior of Stand-up Targets
Const STAnimStep =  1.5     'vpunits per animation step (control return to Start)
Const STMaxOffset = 9       'max vp units target moves when hit

Const STMass = 0.2        'Mass of the Stand-up Target (between 0 and 1), higher values provide more resistance

'******************************************************
' STAND-UP TARGETS FUNCTIONS
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
' END STAND-UP TARGETS
'******************************************************



'******************************************************
' BALL ROLLING AND DROP SOUNDS
'******************************************************


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

  ' stop the sound of deleted balls
  For b = UBound(gBOT) + 1 To tnob - 1
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

  Next
End Sub

'******************************************************
' END BALL ROLLING AND DROP SOUNDS
'******************************************************




'******************************************************
' FLEEP MECHANICAL SOUNDS
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
' Fleep Supporting Ball & Sound Functions
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
' END FLEEP MECHANICAL SOUNDS
'******************************************************




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

'  ROM Volume Settings
    With Controller
        .Games(cGameName).Settings.Value("volume") = ROMVolume
  End With

End Sub



Dim VRThings

if VR_Room = 0 and cab_mode = 0 Then
  for each VRThings in VRStuff:VRThings.visible = 0:Next
  for each VRThings in VRClock:VRThings.visible = 0:Next
  for each VRThings in VRBackglass:VRThings.visible = 0:Next
  for each VRThings in BP_rails:VRThings.visible = 1:Next
  VR_Front_Wall_Block.visible = 0
  VR_Front_Wall_Block.sidevisible = 0
  EndTroughWall.visible = 0
  EndTroughWall.sidevisible = 0
  EndTroughWall2.visible = 0
  EndTroughWall2.sidevisible = 0
Elseif VR_Room = 0 and cab_mode = 1 Then
  for each VRThings in VRStuff:VRThings.visible = 0:Next
  for each VRThings in VRClock:VRThings.visible = 0:Next
  for each VRThings in VRBackglass:VRThings.visible = 0:Next
  for each VRThings in DTBackglass:VRThings.visible = 0: Next
  for each VRThings in BP_rails:VRThings.visible = 0:Next
  VR_Front_Wall_Block.visible = 0
  VR_Front_Wall_Block.sidevisible = 0
  EndTroughWall.visible = 0
  EndTroughWall.sidevisible = 0
  EndTroughWall2.visible = 0
  EndTroughWall2.sidevisible = 0
Else
  for each VRThings in VRStuff:VRThings.visible = 1:Next
  for each VRThings in VRClock:VRThings.visible = WallClock:Next
  for each VRThings in DTBackglass:VRThings.visible = 0: Next
  for each VRThings in BP_rails:VRThings.visible = 0:Next
  VR_Front_Wall_Block.visible = 1
  VR_Front_Wall_Block.sidevisible = 1
  EndTroughWall.visible = 1
  EndTroughWall.sidevisible = 1
  EndTroughWall2.visible = 1
  EndTroughWall2.sidevisible = 1

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

DIM GIP: GIP = 0

Sub Gameinplay (Enabled)
  If Enabled Then
    GIP = 1
  Else
    GIP = 0
  End If
End Sub

if VR_Room = 1 Then
  Set LampCallback = GetRef("UpdateMultipleLamps")
End If

Sub UpdateMultipleLamps()

  If Controller.Lamp(29) = 0 Then: FlBGL29.visible=0: else: FlBGL29.visible=1 ' High Score To Date
  If Controller.Lamp(13) = 0 Then: FlBGL13.visible=0: else: FlBGL13.visible=1 ' Ball In Play
  If Controller.Lamp(27) = 0 Then: FlBGL27.visible=0: else: FlBGL27.visible=1 ' Match
  If Controller.Lamp(45) = 0 Then: FlBGL45.visible=0: else: FlBGL45.visible=1 ' Game Over
  If Controller.Lamp(61) = 0 Then: FlBGL61.visible=0: else: FlBGL61.visible=1 ' Tilt
  If Controller.Lamp(14) = 0 Then: FlBGL14.visible=0: else: FlBGL14.visible=1 ' D
  If Controller.Lamp(30) = 0 Then: FlBGL30.visible=0: else: FlBGL30.visible=1 ' E1
  If Controller.Lamp(46) = 0 Then: FlBGL46.visible=0: else: FlBGL46.visible=1 ' L
  If Controller.Lamp(62) = 0 Then: FlBGL62.visible=0: else: FlBGL62.visible=1 ' U
  If Controller.Lamp(15) = 0 Then: FlBGL15.visible=0: else: FlBGL15.visible=1 ' X
  If Controller.Lamp(31) = 0 Then: FlBGL31.visible=0: else: FlBGL31.visible=1 ' E2

  If GIP = 1 Then
    If Controller.Lamp(43) = 0 Then: FlBGL43.visible=0: else: FlBGL43.visible=1 ' Shoot Again
  End If

End Sub

if VR_Room = 0 and cab_mode = 0 Then
  Set LampCallback = GetRef("UpdateDTLamps")
End If

Sub UpdateDTLamps()

  If Controller.Lamp(29) = 0 Then: HighScoreReel.setValue(0):   Else: HighScoreReel.setValue(1) 'High Score
  If Controller.Lamp(13) = 0 Then: BIPReel.setValue(0):     Else: BIPReel.setValue(1) 'Ball in Play
  If Controller.Lamp(27) = 0 Then: MatchReel.setValue(0):     Else: MatchReel.setValue(1) 'Match
  If Controller.Lamp(45) = 0 Then: GameOverReel.setValue(0):    Else: GameOverReel.setValue(1) 'Game Over
  If Controller.Lamp(61) = 0 Then: TiltReel.setValue(0):      Else: TiltReel.setValue(1) 'Tilt
  If Controller.Lamp(14) = 0 Then: DTLightD.state =  0: Else: DTLightD.state =  1  'D
  If Controller.Lamp(30) = 0 Then: DTLightE.state =  0: Else: DTLightE.state =  1  'E
  If Controller.Lamp(46) = 0 Then: DTLightL.state =  0: Else: DTLightL.state =  1  'L
  If Controller.Lamp(62) = 0 Then: DTLightU.state =  0: Else: DTLightU.state =  1  'U
  If Controller.Lamp(15) = 0 Then: DTLightX.state =  0: Else: DTLightX.state =  1  'X
  If Controller.Lamp(31) = 0 Then: DTLightE2.state = 0: Else: DTLightE2.state = 1  'E2

  If GIP = 1 Then
    If Controller.Lamp(43) = 0 Then: ShootAgainReel.setValue(0):  Else: ShootAgainReel.setValue(1) 'Shoot Again
  End If

End Sub



'******** Revisions done by Bord and UnclePaulie on Hybrid version 1.01 - 2.0 *********

'UnclePaulie edits
' v1.01 Used main VR cabinet assets and graphics from a version Sixtoe did awhile back
'   Enabled hybrid version for desktop, VR, and cabinet modes.  Automated cab and desktop mode.
'   Added various timers for full implementation
'   Modded the VR cab prims positions.  Also added posters, clock, and other minor VR elements. Animated flippers and start button.
'   Animated the VR backglass and used the image from BlackSad's B2S
'   Removed manual ball control.
'   Adjusted cab pov
' v1.02 Completely created a new desktop background
'     Added lighter ball images provided by hauntfreaks and g5k ball scratches, changeable in script
' v1.03 Added Fleep sound files and VPW physics primitives, and adjusted physics materials to zcol materials.
' v1.04 Adjusted the flippers slightly... they were off, and too long.
' v1.05 Added all the VPW script for physics, sounds, etc.
' v1.06 Created a new trough system for single ball games with only a outhole solenoid and not a BallRelease solenoid
'   Changed the apron wall structure to create a trough
'     No longer destroy balls.  Use a global gBOT throughout all of script.
'   Created new saucer catch and kick routine for the gBOT
'   Added VPW Dynamic Ball Shadows to the solution done by Iakki, Apophis, and Wylte
'     Added nFozzy / Roth physics and flippers.  Updated the flipper physics per recommendations.
'     Updated ball rolling sounds
'   Added Apophis sling corrections, and updated the physics.
' v1.07 Added new drop targets to the solution done by Rothbauerw - code needed to be redone, as solenoids don't match manual
'   Lined up the drop targets correctly.  They weren't straight before.
' v1.08 Investigation into solenoid operation, rewrote ball release code.  Added an Lamp52On variable and sub to be controlled by lamp 52 controller.  Went away before ballrel solenoid.  Only affected 2-4 players.
'   Redid fastflips.. don't need all the extra code, just solenoids = 2
'   Cleaned up timers that weren't used anymore.
' v1.09 Added updated standup targets with VPW physics, original code created by Rothbauerw and VPW team
'   Updated the use of the TargetBounce collection for usage on standup targets.
'   Added groove to plunger lane, and small ball release ramp to get up.
'     Added updated knocker position and code
'   Slight mod to desktop background picture
'     Updated the images on the drop Targets, and lowered the drop units so that they are level with playfield.
'   Moved the 7 bank of targets to line up.  Could see in VR.
'     Adjusted the sling prims, sensors, and triggers to align correctly.
' v1.10 Cleaned code up.
'   Added DTBackglass collection
'   Added blacksidewalls for cabinet mode as an option
' v1.11 Updated to Lampz lighting routine by VPW. (removed the lamp2 timer, and put in the frame timer)
'   Increased the sizes of all the targets to more closely match a real table.  Aligned the target walls to match real target.
' v1.12 Adjusted the table physics, gravity constant, flipper physics, etc.  Also sling and bumper thresholds changed to match table physics
'   Corrected the playfield cutouts (some trigger holes weren't fully transparent, could really see in VR.
' v1.13 Was still a timing issue on some PCs (tested on older and faster machines) with Controller.Lamp(52) and SolBallRelease happening too quick.
'   - Had to ensure they were solidly latched, and then reset once ball comes out of trough.  Had to control sound of Drop Target as well.
'   Made the Plasticsrim_prim collidable, as once a ball jumped over a wall.
'   Adjusted the environment image, blooms, and other settings.
'   Adjusted the VR backglass primitive
'   Slight mod to trough with a timer. Also moved the ballrelease sound to a trigger.
'   Adjusted the sling and bumper physics
' v1.14 Fixed an error in the SW29 z scale.
'   Slight adjustement to the environment lighting
'bord edits
' v1.15 swapped in Bally flippers
' v1.16 nearly all physics replaced with meshes based on new model/scale
' v1.17 physics fixes to apron/Drain
'   plunger position changed
'   gate physics Update
'   added mesh playfield, removed VP kicker, added saucer physics elements, replaced with non-visually binary saucer scripting (should now be able to miss the collect shot per PAPA video)
' v1.18 UnclePaulie automated VR room, cab, and desktop.
'   Updated to latest VPW script for dynamic shadows.
' v1.19 bord lightmap pass 1 2k render
'UnclePaulie edits
' v1.20 Updated the script to get up to latest VPW standards
'   Removed calls to various old lights when drop targets would go down
'   Updated the sling routines for the new baked prims.
'   Moved the sling correction physical prims
'   Updated code for DT and ST's, and resized the DT and ST physics
'   Converted code to standalone capability
'   Updated to latest flipper code and triggers by rothbauerw.  Triggers are 27 VPU, and 3 degrees added.
'   Updated the flipper physics EOS torqe, end angle, and strength to match VPW recommendations.
'   Updated the bumper force, threshold and scatter.  Matches PAPA play, and meets VPW recommendations.
'   Removed catchme wall.
'   Removed dynamic shadows, Lampz, and dynamic sources collection.
'   Found an error in the early 2k bake for drop targets. They were placed backwards.
'     For testing, I rotated them all 180 rotz and objrotz, then moved them into place.  And dropped too far.
'   Animated Drop and standup targets, flippers, bumper rings, bumper skirts, gates, star trigger, and lane rollover triggers.
'   Had to adjust the physics of the gates slightly and align with relation to ball.
'   Added a small invisible gate wall near the upper left gate.
'   Adjusted the plunger location... stuck out too far.  Also, adjusted the pull speed, etc.
'   Adjusted the power of the ballrelease kicker
'   Walls are now a prim... needed to add to the walls collection for sounds.
'   Aligned trigger rollover physics
'   Added rubber physics to rubbers sound collection.
'   Tied lower metal guide rail to metal sounds collection.
'   Removed duplicate wall005 and ramp006
'   Lower plastic walls weren't tied to sounds, so tied to walls collection
'   Added physics walls to lower guide rails and all plastics areas.  Avoids balls getting stuck inside.
'   Removed old code that vibrated old rubber meshes now that they are gone.
'   Fixed SW24 sensor hit, and removed old meshes and code.
'   Removed old kickarm animation
'   Added ruleset
'   Flipper shadows, changed to "hide Parts behind".
'   Updated fleep sounds, audio, and ensure code is same.
'bord edits
' v1.21 new lightmap with fixed lamp assignments for semi-transparent inserts
'   standup and drop target orientation corrected
'UnclePaulie edits
' v1.22
'   Updated the sling rubbers to be not visible at startup in the script.
'   Added new ball images and options
'   Added F12 menu for various options
'   Set default DIPS, and simplified the code.  Also corrected the F6 menu for two switches.
'   Updated VR element placements to align with updated table.
'   Removed VR topper. Changed the VR clock image. BM_lockdown invisible for VR mode, adjusted off LEDs.  Added a VR front wall to block seeing behind.
'   Removed unused images
'   Removed old kicker primitive and images
'   Put flipper activate and deactivate in flipper subs.
'   Removed duplicate guard rails behind the flippers and in the outlanes.  Also the right flipper rail was duplicated.
'   Replaced the anit snubber rails with physics prims.
'   Fixed the physics under the apron, and slightly adjusted the trough physics.
'   Added an end trough wall for VR.
'   Corrected Shoot Again light for VR and Desktop (it should only come on when game is in play.)
'bord edits
' v1.23 new 4k render of the game
'UnclePaulie edits
' v1.24 Adjusted the gatewire VLMprims z height to 60 and the gateplate VLMprims to 71.5 in the script.
'   Also changed the gate animation formula for rotation angle (went a bit too far before)
'   Converted several images to webp format
'   Deleted desktop side rails, as Bord put new ones into the bake.
'   Slightly increased VR cab and associated VR prims to accomodate for VLM sidewall reflections.
'   Fixed the sling animations to account for VLM BP arrays
' RC2 Adjusted the upper flipper location and the BP_Flip3 arrays prims.  Also the flipper end angle.
'   Changed the depth bias on flipper shadows to -1000
'   Added room brightness options to the F12 menu, adjusted day/night to 35, and default room brightness to 65
'   Adjusted the tilt sensitivity to 3 (down from 6).  Also, the nudge strength from 1 to 0.5.
'   Added an additional ball and scratch images
'   Corrected all lightmaps to have Unshaded Additive Blend checked.
' RC3 Updated cabmode to remove BP_rails
'   Added an apron physic wall.
'   VR Elements had reflections enabled... turned them off, as you could see in VR.
' RC4 Redid the playfield mesh and made visible.  Also had to do BM_Playfield reflections with a -1 render reflections.
'   Adjusted playfield friction to .25.  Bumpers down to 11.  And slings to force of 3.75 and threshold of 2.  Also adjusted the rubberposts elasticity up to .9125
'   Adjusted the ball reflection scale down to 0.3
' RC4a  Removed all the VLM Dynamic reflections
'   Turned reflection off for two VR end trough walls.
'   Reduced overall reflection level to 10. Playfield slope to 5.25.
'   Adjusted gate5 by 6 degrees to get better bounce
' RC4b  Changed the friction on Gate5 to 0.4.  Helps the bounce off to hit the opposite side of the top.
'   DarthVito provided several new ball images in the ball options.  Chose one of them as new default.

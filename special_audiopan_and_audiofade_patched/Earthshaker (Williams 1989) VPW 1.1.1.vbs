'
'     ___  _  __ ___     __       _       ___ __
'     )_  /_) )_) ) )_) (_ ` )_) /_) )_/  )_  )_) |
'    (__ / / / \ ( ( ( .__) ( ( / / /  ) (__ / \  o
'
'          Earthshaker (Williams 1989) VPW
'
'
'
'  VPW Seismologists
'  ~~~~~~~~~~~~~~~~~
'  Friscopinball - Project Lead, asset scans
'  Tomate - 3D Modeling and Rendering
'  Sixtoe - Building Physical Table, Lots of VR Stuff
'  Apophis - Blender toolkit scripting, physics updates, animations, options
'  UnclePaulie - VR backglass and alphanumeric DMD
'  TastyWasps - Additional physics, Fleep updates, Deluxe VR Room
'  Bord, Benji & Oqqsan - merged WIP versions, physics updates
'  Primetime5k - Staged flipper functionality, testing
'  Niwak - VPX Blender Toolkit, VPX Part Library
'  VPW Testers - PinStratsDan, Cliffy, Bietekwiet, Lumi, Darth Vito, Studlygoorite, jarr3, somatic, DGrimmReaper, lminimart
'  Fliptronic - Playfield assets, research
'  Based on works by 32Assassin
'
'
'
'
'************************************************************************************************
' === TABLE OF CONTENTS  ===
'
' You can quickly jump to a section by searching the four letter tag (ZXXX)
'
' ZCON: Constants and Global Variables
' ZVLM: VLM Arrays
' ZLOA: Load Stuff
' ZTIM: Main Timers
' ZINI: Table Initialization and Exiting
'   ZOPT: User Options
'   ZBRI: Room Brightness
'   ZKEY: Key Press Handling
' ZSOL: Solenoids
'   ZFLA: PWM Flasher Stuff
' ZFLP: Flippers
' ZTOY: Solenoid Controlled toys
' ZGII: General Illumination
' ZDRN: Drain, Trough, and Ball Release
' ZSWI: Switches
' ZSTA: Stand-up Targets
' ZDTA: Drop Targets
' ZSLG: Slingshots
' ZVUK: VUK's
' ZCND: California Nevada Diverter
' ZEIA: Earthquake Institute Animation
'   ZABS: Ambient ball shadows
' ZLED: LEDs
'   ZMAT: General Math Functions
' ZNFF: Flipper Corrections
'   ZDMP: Rubber Dampeners
' ZSSC: Slingshot Corrections
'   ZBOU: VPW TargetBouncer
' ZSFX: Mechanical Sound effects
' ZBRL: Ball Rolling and Drop Sounds
'   ZRRL: Ramp Rolling Sound effects
'   ZANI: Misc Animations
' ZVRR: VR Room Stuff
' ZCHL: Change Log
'
'************************************************************************************************

Option Explicit
Randomize
SetLocale 1033


'*******************************************
' ZCON: Constants and Global Variables
'*******************************************

Dim esball1, esball2, esball3, gBOT
Dim bsLSaucer
Dim VRRoom
Dim DesktopMode: DesktopMode = Table1.ShowDT

Dim tablewidth: tablewidth = Table1.width
Dim tableheight: tableheight = Table1.height

Const BallSize = 50     ' Ball size must be 50
Const BallMass = 1      ' Ball mass must be 1
Const tnob = 3        ' Total number of balls
Const lob = 0       ' Total number of locked balls


'*******************************************
' ZVLM: VLM Arrays
'*******************************************


' VLM  Arrays - Start
' Arrays per baked part
Dim BP_BC1: BP_BC1=Array(BM_BC1, LM_GI_BC1, LM_IN_L17_BC1, LM_IN_L24_BC1, LM_IN_L41_BC1, LM_IN_L42_BC1, LM_IN_L43_BC1, LM_IN_L7_BC1, LM_IN_L8_BC1, LM_FL_f1_BC1, LM_FL_f10_BC1, LM_FL_f2_BC1, LM_FL_f3_BC1)
Dim BP_BC2: BP_BC2=Array(BM_BC2, LM_GIS_GI_13_BC2, LM_GI_BC2, LM_IN_L17_BC2, LM_IN_L20_BC2, LM_IN_L21_BC2, LM_IN_L23_BC2, LM_IN_L24_BC2, LM_IN_L25_BC2, LM_IN_L41_BC2, LM_IN_L42_BC2, LM_IN_L43_BC2, LM_IN_L48_BC2, LM_IN_L7_BC2, LM_IN_L8_BC2, LM_FL_f1_BC2, LM_FL_f10_BC2, LM_FL_f2_BC2, LM_FL_f3_BC2, LM_FL_f6_BC2, LM_FL_f9_BC2)
Dim BP_BC3: BP_BC3=Array(BM_BC3, LM_GI_BC3, LM_IN_L42_BC3, LM_IN_L8_BC3, LM_FL_f10_BC3, LM_FL_f114_BC3, LM_FL_f3_BC3, LM_FL_f6_BC3)
Dim BP_BR1: BP_BR1=Array(BM_BR1, LM_GI_BR1, LM_IN_L24_BR1, LM_IN_L41_BR1, LM_IN_L42_BR1, LM_IN_L43_BR1, LM_IN_L7_BR1, LM_IN_L8_BR1, LM_FL_f10_BR1, LM_FL_f2_BR1, LM_FL_f3_BR1)
Dim BP_BR2: BP_BR2=Array(BM_BR2, LM_GIS_GI_13_BR2, LM_GI_BR2, LM_IN_L42_BR2, LM_IN_L43_BR2, LM_IN_L8_BR2, LM_FL_f2_BR2, LM_FL_f3_BR2, LM_FL_f6_BR2)
Dim BP_BR3: BP_BR3=Array(BM_BR3, LM_GI_BR3, LM_IN_L42_BR3, LM_IN_L8_BR3, LM_FL_f10_BR3, LM_FL_f114_BR3, LM_FL_f2_BR3, LM_FL_f6_BR3)
Dim BP_BS1: BP_BS1=Array(BM_BS1, LM_GI_BS1, LM_IN_L21_BS1, LM_IN_L23_BS1, LM_IN_L24_BS1, LM_IN_L41_BS1, LM_IN_L42_BS1, LM_IN_L43_BS1, LM_IN_L7_BS1, LM_IN_L8_BS1, LM_FL_f1_BS1, LM_FL_f10_BS1, LM_FL_f2_BS1, LM_FL_f3_BS1)
Dim BP_BS2: BP_BS2=Array(BM_BS2, LM_GI_BS2, LM_IN_L17_BS2, LM_IN_L18_BS2, LM_IN_L19_BS2, LM_IN_L20_BS2, LM_IN_L21_BS2, LM_IN_L22_BS2, LM_IN_L23_BS2, LM_IN_L41_BS2, LM_IN_L43_BS2, LM_IN_L7_BS2, LM_IN_L8_BS2, LM_FL_f10_BS2, LM_FL_f114_BS2, LM_FL_f2_BS2, LM_FL_f3_BS2, LM_FL_f7_BS2)
Dim BP_BS3: BP_BS3=Array(BM_BS3, LM_GI_BS3, LM_IN_L18_BS3, LM_IN_L19_BS3, LM_IN_L21_BS3, LM_IN_L22_BS3, LM_IN_L25_BS3, LM_IN_L41_BS3, LM_IN_L42_BS3, LM_IN_L43_BS3, LM_IN_L8_BS3, LM_FL_f10_BS3, LM_FL_f114_BS3, LM_FL_f2_BS3, LM_FL_f3_BS3)
Dim BP_EIprim: BP_EIprim=Array(BM_EIprim, LM_GI_EIprim, LM_IN_L17_EIprim, LM_IN_L18_EIprim, LM_IN_L19_EIprim, LM_IN_L20_EIprim, LM_IN_L21_EIprim, LM_IN_L22_EIprim, LM_IN_L23_EIprim, LM_IN_L24_EIprim, LM_IN_L25_EIprim, LM_IN_L43_EIprim, LM_IN_L49a_EIprim, LM_IN_L7_EIprim, LM_FL_f1_EIprim, LM_FL_f10_EIprim, LM_FL_f114_EIprim, LM_FL_f114a_EIprim, LM_FL_f2_EIprim, LM_FL_f3_EIprim, LM_FL_f4_EIprim, LM_FL_f5_EIprim, LM_FL_f6_EIprim, LM_FL_f7_EIprim, LM_FL_f9_EIprim)
Dim BP_Gate2: BP_Gate2=Array(BM_Gate2, LM_FL_f4_Gate2)
Dim BP_LEMK: BP_LEMK=Array(BM_LEMK, LM_GIS_GI_1_LEMK, LM_GIS_GI_17_LEMK, LM_GIS_GI_20_LEMK, LM_GI_LEMK, LM_IN_L9_LEMK, LM_FL_f1_LEMK)
Dim BP_LFlipper: BP_LFlipper=Array(BM_LFlipper, LM_GIS_GI_1_LFlipper, LM_GIS_GI_16_LFlipper, LM_GIS_GI_17_LFlipper, LM_GIS_GI_2_LFlipper, LM_GIS_GI_20_LFlipper, LM_GIS_GI_29_LFlipper, LM_IN_L32_LFlipper)
Dim BP_LRubber1: BP_LRubber1=Array(BM_LRubber1, LM_GIS_GI_10_LRubber1, LM_GIS_GI_15_LRubber1, LM_GIS_GI_17_LRubber1, LM_GIS_GI_20_LRubber1, LM_GI_LRubber1, LM_IN_L16_LRubber1, LM_IN_L38_LRubber1, LM_FL_f1_LRubber1, LM_FL_f114_LRubber1, LM_FL_f114a_LRubber1)
Dim BP_LRubber2: BP_LRubber2=Array(BM_LRubber2, LM_GIS_GI_10_LRubber2, LM_GIS_GI_15_LRubber2, LM_GIS_GI_17_LRubber2, LM_GIS_GI_20_LRubber2, LM_GI_LRubber2, LM_IN_L16_LRubber2, LM_IN_L38_LRubber2, LM_FL_f1_LRubber2, LM_FL_f114_LRubber2, LM_FL_f114a_LRubber2)
Dim BP_LRubber3: BP_LRubber3=Array(BM_LRubber3, LM_GIS_GI_10_LRubber3, LM_GIS_GI_15_LRubber3, LM_GIS_GI_17_LRubber3, LM_GIS_GI_20_LRubber3, LM_GI_LRubber3, LM_IN_L16_LRubber3, LM_IN_L38_LRubber3, LM_FL_f1_LRubber3, LM_FL_f114_LRubber3, LM_FL_f114a_LRubber3)
Dim BP_LRubber4: BP_LRubber4=Array(BM_LRubber4, LM_GIS_GI_10_LRubber4, LM_GIS_GI_15_LRubber4, LM_GIS_GI_17_LRubber4, LM_GIS_GI_20_LRubber4, LM_GI_LRubber4, LM_IN_L16_LRubber4, LM_IN_L38_LRubber4, LM_FL_f1_LRubber4, LM_FL_f114_LRubber4, LM_FL_f114a_LRubber4)
Dim BP_LSling1: BP_LSling1=Array(BM_LSling1, LM_GIS_GI_17_LSling1, LM_GIS_GI_20_LSling1, LM_GI_LSling1, LM_IN_L9_LSling1, LM_FL_f6_LSling1)
Dim BP_LSling2: BP_LSling2=Array(BM_LSling2, LM_GIS_GI_17_LSling2, LM_GIS_GI_20_LSling2, LM_GI_LSling2)
Dim BP_LUFlipper: BP_LUFlipper=Array(BM_LUFlipper, LM_GIS_GI_10_LUFlipper, LM_GIS_GI_15_LUFlipper, LM_GI_LUFlipper, LM_FL_f1_LUFlipper, LM_FL_f114_LUFlipper, LM_FL_f114a_LUFlipper, LM_FL_f2_LUFlipper)
Dim BP_Layer1: BP_Layer1=Array(BM_Layer1, LM_GIS_GI_1_Layer1, LM_GIS_GI_10_Layer1, LM_GIS_GI_15_Layer1, LM_GIS_GI_17_Layer1, LM_GIS_GI_2_Layer1, LM_GIS_GI_20_Layer1, LM_GIS_GI_29_Layer1, LM_GI_Layer1, LM_IN_L15_Layer1, LM_IN_L16_Layer1, LM_IN_L17_Layer1, LM_IN_L18_Layer1, LM_IN_L19_Layer1, LM_IN_L20_Layer1, LM_IN_L21_Layer1, LM_IN_L22_Layer1, LM_IN_L23_Layer1, LM_IN_L24_Layer1, LM_IN_L25_Layer1, LM_IN_L34_Layer1, LM_IN_L35_Layer1, LM_IN_L36_Layer1, LM_IN_L37_Layer1, LM_IN_L38_Layer1, LM_IN_L39_Layer1, LM_IN_L41_Layer1, LM_IN_L42_Layer1, LM_IN_L43_Layer1, LM_IN_L49a_Layer1, LM_IN_L54_Layer1, LM_IN_L55_Layer1, LM_IN_L57a_Layer1, LM_IN_L7_Layer1, LM_IN_L8_Layer1, LM_FL_f1_Layer1, LM_FL_f10_Layer1, LM_FL_f114_Layer1, LM_FL_f114a_Layer1, LM_FL_f2_Layer1, LM_FL_f3_Layer1, LM_FL_f4_Layer1, LM_FL_f5_Layer1, LM_FL_f6_Layer1, LM_FL_f7_Layer1, LM_FL_f8_Layer1, LM_FL_f9_Layer1)
Dim BP_Layer2: BP_Layer2=Array(BM_Layer2, LM_GI_Layer2, LM_IN_L17_Layer2, LM_IN_L18_Layer2, LM_IN_L20_Layer2, LM_IN_L21_Layer2, LM_IN_L23_Layer2, LM_IN_L41_Layer2, LM_IN_L7_Layer2, LM_IN_L8_Layer2, LM_FL_f10_Layer2, LM_FL_f2_Layer2, LM_FL_f3_Layer2, LM_FL_f4_Layer2, LM_FL_f5_Layer2, LM_FL_f7_Layer2, LM_FL_f8_Layer2, LM_FL_f9_Layer2)
Dim BP_Layer3: BP_Layer3=Array(BM_Layer3, LM_GI_Layer3, LM_FL_f10_Layer3, LM_FL_f4_Layer3, LM_FL_f5_Layer3, LM_FL_f6_Layer3, LM_FL_f7_Layer3, LM_FL_f8_Layer3, LM_FL_f9_Layer3)
Dim BP_Layer4: BP_Layer4=Array(BM_Layer4, LM_GI_Layer4, LM_FL_f10_Layer4)
Dim BP_Layer5: BP_Layer5=Array(BM_Layer5, LM_FL_f4_Layer5, LM_FL_f9_Layer5)
Dim BP_Layer6: BP_Layer6=Array(BM_Layer6, LM_GIS_GI_13_Layer6, LM_GI_Layer6, LM_IN_L28_Layer6, LM_IN_L30_Layer6, LM_IN_L31_Layer6, LM_IN_L57a_Layer6, LM_FL_f114_Layer6, LM_FL_f114a_Layer6, LM_FL_f2_Layer6, LM_FL_f6_Layer6, LM_FL_f7_Layer6)
Dim BP_Layer7: BP_Layer7=Array(BM_Layer7, LM_GI_Layer7, LM_FL_f6_Layer7)
Dim BP_Lockdownbar: BP_Lockdownbar=Array(BM_Lockdownbar, LM_GI_Lockdownbar, LM_IN_L49a_Lockdownbar)
Dim BP_Parts: BP_Parts=Array(BM_Parts, LM_GIS_GI_1_Parts, LM_GIS_GI_10_Parts, LM_GIS_GI_13_Parts, LM_GIS_GI_15_Parts, LM_GIS_GI_16_Parts, LM_GIS_GI_17_Parts, LM_GIS_GI_2_Parts, LM_GIS_GI_20_Parts, LM_GIS_GI_29_Parts, LM_GI_Parts, LM_IN_L13_Parts, LM_IN_L15_Parts, LM_IN_L16_Parts, LM_IN_L17_Parts, LM_IN_L18_Parts, LM_IN_L19_Parts, LM_IN_L20_Parts, LM_IN_L21_Parts, LM_IN_L22_Parts, LM_IN_L23_Parts, LM_IN_L24_Parts, LM_IN_L25_Parts, LM_IN_L27_Parts, LM_IN_L28_Parts, LM_IN_L2_Parts, LM_IN_L30_Parts, LM_IN_L31_Parts, LM_IN_L32_Parts, LM_IN_L34_Parts, LM_IN_L36_Parts, LM_IN_L38_Parts, LM_IN_L39_Parts, LM_IN_L3_Parts, LM_IN_L41_Parts, LM_IN_L42_Parts, LM_IN_L43_Parts, LM_IN_L46_Parts, LM_IN_L47_Parts, LM_IN_L48_Parts, LM_IN_L49a_Parts, LM_IN_L4_Parts, LM_IN_L50_Parts, LM_IN_L52_Parts, LM_IN_L53_Parts, LM_IN_L54_Parts, LM_IN_L55_Parts, LM_IN_L56_Parts, LM_IN_L57a_Parts, LM_IN_L5_Parts, LM_IN_L6_Parts, LM_IN_L7_Parts, LM_IN_L8_Parts, LM_IN_L9_Parts, LM_FL_f1_Parts, LM_FL_f10_Parts, LM_FL_f114_Parts, LM_FL_f114a_Parts, _
  LM_FL_f2_Parts, LM_FL_f3_Parts, LM_FL_f4_Parts, LM_FL_f5_Parts, LM_FL_f6_Parts, LM_FL_f7_Parts, LM_FL_f8_Parts, LM_FL_f9_Parts)
Dim BP_Playfield: BP_Playfield=Array(BM_Playfield, LM_GIS_GI_1_Playfield, LM_GIS_GI_10_Playfield, LM_GIS_GI_13_Playfield, LM_GIS_GI_15_Playfield, LM_GIS_GI_16_Playfield, LM_GIS_GI_17_Playfield, LM_GIS_GI_2_Playfield, LM_GIS_GI_20_Playfield, LM_GIS_GI_29_Playfield, LM_GI_Playfield, LM_IN_L10_Playfield, LM_IN_L11_Playfield, LM_IN_L12_Playfield, LM_IN_L13_Playfield, LM_IN_L14_Playfield, LM_IN_L15_Playfield, LM_IN_L16_Playfield, LM_IN_L17_Playfield, LM_IN_L18_Playfield, LM_IN_L19_Playfield, LM_IN_L1_Playfield, LM_IN_L20_Playfield, LM_IN_L21_Playfield, LM_IN_L22_Playfield, LM_IN_L23_Playfield, LM_IN_L24_Playfield, LM_IN_L25_Playfield, LM_IN_L28_Playfield, LM_IN_L29_Playfield, LM_IN_L2_Playfield, LM_IN_L30_Playfield, LM_IN_L31_Playfield, LM_IN_L32_Playfield, LM_IN_L33_Playfield, LM_IN_L34_Playfield, LM_IN_L35_Playfield, LM_IN_L36_Playfield, LM_IN_L37_Playfield, LM_IN_L38_Playfield, LM_IN_L39_Playfield, LM_IN_L3_Playfield, LM_IN_L40_Playfield, LM_IN_L41_Playfield, LM_IN_L42_Playfield, LM_IN_L43_Playfield, _
  LM_IN_L44_Playfield, LM_IN_L45_Playfield, LM_IN_L46_Playfield, LM_IN_L47_Playfield, LM_IN_L48_Playfield, LM_IN_L49a_Playfield, LM_IN_L4_Playfield, LM_IN_L51_Playfield, LM_IN_L52_Playfield, LM_IN_L53_Playfield, LM_IN_L54_Playfield, LM_IN_L55_Playfield, LM_IN_L56_Playfield, LM_IN_L57a_Playfield, LM_IN_L5_Playfield, LM_IN_L6_Playfield, LM_IN_L7_Playfield, LM_IN_L8_Playfield, LM_IN_L9_Playfield, LM_FL_f1_Playfield, LM_FL_f10_Playfield, LM_FL_f114_Playfield, LM_FL_f114a_Playfield, LM_FL_f2_Playfield, LM_FL_f3_Playfield, LM_FL_f4_Playfield, LM_FL_f5_Playfield, LM_FL_f6_Playfield, LM_FL_f7_Playfield, LM_FL_f8_Playfield, LM_FL_f9_Playfield)
Dim BP_REMK: BP_REMK=Array(BM_REMK, LM_GIS_GI_13_REMK, LM_GIS_GI_16_REMK, LM_GIS_GI_2_REMK, LM_GIS_GI_29_REMK, LM_GI_REMK, LM_IN_L13_REMK)
Dim BP_RFlipper: BP_RFlipper=Array(BM_RFlipper, LM_GIS_GI_1_RFlipper, LM_GIS_GI_16_RFlipper, LM_GIS_GI_17_RFlipper, LM_GIS_GI_2_RFlipper, LM_GIS_GI_20_RFlipper, LM_GI_RFlipper, LM_IN_L32_RFlipper)
Dim BP_RRubber1: BP_RRubber1=Array(BM_RRubber1, LM_GIS_GI_13_RRubber1, LM_GIS_GI_16_RRubber1, LM_GIS_GI_29_RRubber1, LM_GI_RRubber1, LM_IN_L28_RRubber1, LM_IN_L30_RRubber1, LM_IN_L31_RRubber1, LM_FL_f114_RRubber1, LM_FL_f6_RRubber1)
Dim BP_RRubber2: BP_RRubber2=Array(BM_RRubber2, LM_GIS_GI_13_RRubber2, LM_GIS_GI_16_RRubber2, LM_GIS_GI_29_RRubber2, LM_GI_RRubber2, LM_IN_L28_RRubber2, LM_IN_L30_RRubber2, LM_IN_L31_RRubber2, LM_FL_f114_RRubber2, LM_FL_f6_RRubber2)
Dim BP_RRubber3: BP_RRubber3=Array(BM_RRubber3, LM_GIS_GI_13_RRubber3, LM_GIS_GI_16_RRubber3, LM_GIS_GI_29_RRubber3, LM_GI_RRubber3, LM_IN_L28_RRubber3, LM_IN_L30_RRubber3, LM_IN_L31_RRubber3, LM_FL_f114_RRubber3, LM_FL_f114a_RRubber3, LM_FL_f6_RRubber3)
Dim BP_RRubber4: BP_RRubber4=Array(BM_RRubber4, LM_GIS_GI_13_RRubber4, LM_GIS_GI_16_RRubber4, LM_GIS_GI_29_RRubber4, LM_GI_RRubber4, LM_IN_L28_RRubber4, LM_IN_L30_RRubber4, LM_IN_L31_RRubber4, LM_FL_f114_RRubber4, LM_FL_f6_RRubber4)
Dim BP_RSling1: BP_RSling1=Array(BM_RSling1, LM_GIS_GI_16_RSling1, LM_GIS_GI_2_RSling1, LM_GIS_GI_29_RSling1, LM_FL_f1_RSling1)
Dim BP_RSling2: BP_RSling2=Array(BM_RSling2, LM_GIS_GI_16_RSling2, LM_FL_f1_RSling2)
Dim BP_cabRails: BP_cabRails=Array(BM_cabRails, LM_GIS_GI_10_cabRails, LM_GIS_GI_13_cabRails, LM_GIS_GI_16_cabRails, LM_GIS_GI_17_cabRails, LM_GIS_GI_20_cabRails, LM_GI_cabRails, LM_IN_L17_cabRails, LM_IN_L18_cabRails, LM_IN_L19_cabRails, LM_IN_L20_cabRails, LM_IN_L21_cabRails, LM_IN_L22_cabRails, LM_IN_L24_cabRails, LM_IN_L25_cabRails, LM_IN_L49a_cabRails, LM_IN_L57a_cabRails, LM_FL_f1_cabRails, LM_FL_f10_cabRails, LM_FL_f114_cabRails, LM_FL_f114a_cabRails, LM_FL_f2_cabRails, LM_FL_f3_cabRails, LM_FL_f4_cabRails, LM_FL_f5_cabRails, LM_FL_f6_cabRails, LM_FL_f7_cabRails, LM_FL_f9_cabRails)
Dim BP_diverter: BP_diverter=Array(BM_diverter, LM_GI_diverter, LM_FL_f4_diverter, LM_FL_f7_diverter, LM_FL_f9_diverter)
Dim BP_gate1: BP_gate1=Array(BM_gate1, LM_GI_gate1, LM_FL_f114_gate1, LM_FL_f114a_gate1, LM_FL_f2_gate1, LM_FL_f7_gate1)
Dim BP_lpost: BP_lpost=Array(BM_lpost, LM_GIS_GI_10_lpost, LM_GIS_GI_15_lpost, LM_GIS_GI_20_lpost, LM_GI_lpost, LM_IN_L16_lpost, LM_IN_L37_lpost, LM_FL_f1_lpost, LM_FL_f114a_lpost, LM_FL_f6_lpost)
Dim BP_rpost: BP_rpost=Array(BM_rpost, LM_GIS_GI_13_rpost, LM_GIS_GI_20_rpost, LM_GI_rpost, LM_IN_L29_rpost, LM_IN_L30_rpost, LM_IN_L31_rpost, LM_FL_f1_rpost, LM_FL_f114_rpost, LM_FL_f6_rpost)
Dim BP_sw14: BP_sw14=Array(BM_sw14, LM_GIS_GI_16_sw14, LM_GIS_GI_17_sw14)
Dim BP_sw15: BP_sw15=Array(BM_sw15)
Dim BP_sw16: BP_sw16=Array(BM_sw16)
Dim BP_sw17: BP_sw17=Array(BM_sw17, LM_GIS_GI_17_sw17, LM_GIS_GI_2_sw17, LM_GIS_GI_20_sw17)
Dim BP_sw18: BP_sw18=Array(BM_sw18, LM_GIS_GI_16_sw18, LM_GIS_GI_17_sw18, LM_GIS_GI_20_sw18)
Dim BP_sw19: BP_sw19=Array(BM_sw19, LM_GIS_GI_10_sw19, LM_GIS_GI_20_sw19, LM_GI_sw19, LM_FL_f1_sw19, LM_FL_f114_sw19, LM_FL_f114a_sw19, LM_FL_f6_sw19)
Dim BP_sw21: BP_sw21=Array(BM_sw21, LM_GI_sw21, LM_FL_f114_sw21, LM_FL_f114a_sw21, LM_FL_f6_sw21)
Dim BP_sw22: BP_sw22=Array(BM_sw22, LM_GI_sw22, LM_FL_f114_sw22, LM_FL_f114a_sw22, LM_FL_f6_sw22)
Dim BP_sw23: BP_sw23=Array(BM_sw23, LM_GI_sw23, LM_IN_L57a_sw23, LM_FL_f3_sw23, LM_FL_f6_sw23, LM_FL_f7_sw23)
Dim BP_sw24: BP_sw24=Array(BM_sw24, LM_GIS_GI_13_sw24, LM_GIS_GI_16_sw24, LM_GI_sw24, LM_IN_L28_sw24, LM_IN_L6_sw24, LM_FL_f1_sw24, LM_FL_f114a_sw24, LM_FL_f6_sw24)
Dim BP_sw27: BP_sw27=Array(BM_sw27, LM_GI_sw27, LM_IN_L17_sw27, LM_IN_L20_sw27, LM_IN_L23_sw27, LM_IN_L24_sw27, LM_IN_L41_sw27, LM_IN_L43_sw27, LM_IN_L8_sw27, LM_FL_f2_sw27, LM_FL_f3_sw27)
Dim BP_sw28: BP_sw28=Array(BM_sw28, LM_GI_sw28, LM_IN_L17_sw28, LM_IN_L18_sw28, LM_IN_L24_sw28, LM_IN_L43_sw28, LM_IN_L8_sw28, LM_FL_f114_sw28, LM_FL_f3_sw28)
Dim BP_sw29: BP_sw29=Array(BM_sw29, LM_GI_sw29, LM_IN_L43_sw29, LM_FL_f114_sw29, LM_FL_f2_sw29, LM_FL_f3_sw29)
Dim BP_sw30: BP_sw30=Array(BM_sw30, LM_GI_sw30, LM_IN_L48_sw30, LM_IN_L53_sw30, LM_IN_L54_sw30, LM_IN_L56_sw30, LM_FL_f1_sw30, LM_FL_f114_sw30, LM_FL_f114a_sw30, LM_FL_f2_sw30, LM_FL_f4_sw30, LM_FL_f6_sw30, LM_FL_f7_sw30, LM_FL_f8_sw30, LM_FL_f9_sw30)
Dim BP_sw31: BP_sw31=Array(BM_sw31, LM_GI_sw31, LM_FL_f4_sw31, LM_FL_f7_sw31, LM_FL_f9_sw31)
Dim BP_sw32: BP_sw32=Array(BM_sw32, LM_FL_f4_sw32)
Dim BP_sw33: BP_sw33=Array(BM_sw33, LM_GI_sw33, LM_FL_f10_sw33, LM_FL_f3_sw33)
Dim BP_sw34: BP_sw34=Array(BM_sw34, LM_FL_f10_sw34, LM_FL_f3_sw34)
Dim BP_sw35: BP_sw35=Array(BM_sw35, LM_GI_sw35, LM_FL_f10_sw35, LM_FL_f3_sw35)
Dim BP_sw36: BP_sw36=Array(BM_sw36)
Dim BP_sw41: BP_sw41=Array(BM_sw41, LM_GI_sw41, LM_IN_L20_sw41, LM_IN_L23_sw41, LM_IN_L41_sw41, LM_IN_L43_sw41, LM_IN_L7_sw41, LM_IN_L8_sw41, LM_FL_f10_sw41, LM_FL_f2_sw41, LM_FL_f3_sw41, LM_FL_f4_sw41, LM_FL_f5_sw41)
Dim BP_sw41a: BP_sw41a=Array(BM_sw41a, LM_FL_f3_sw41a)
Dim BP_sw43: BP_sw43=Array(BM_sw43, LM_GI_sw43, LM_FL_f7_sw43)
Dim BP_sw44: BP_sw44=Array(BM_sw44, LM_GI_sw44, LM_IN_L49a_sw44, LM_FL_f4_sw44, LM_FL_f7_sw44, LM_FL_f9_sw44)
Dim BP_sw45: BP_sw45=Array(BM_sw45, LM_GI_sw45, LM_FL_f1_sw45, LM_FL_f10_sw45, LM_FL_f2_sw45)
Dim BP_sw46: BP_sw46=Array(BM_sw46, LM_GIS_GI_10_sw46, LM_GIS_GI_15_sw46, LM_GI_sw46, LM_FL_f1_sw46, LM_FL_f2_sw46, LM_FL_f3_sw46)
Dim BP_sw50: BP_sw50=Array(BM_sw50)
Dim BP_underPF: BP_underPF=Array(BM_underPF, LM_GIS_GI_1_underPF, LM_GIS_GI_10_underPF, LM_GIS_GI_13_underPF, LM_GIS_GI_15_underPF, LM_GIS_GI_16_underPF, LM_GIS_GI_17_underPF, LM_GIS_GI_20_underPF, LM_GI_underPF, LM_IN_L10_underPF, LM_IN_L11_underPF, LM_IN_L12_underPF, LM_IN_L13_underPF, LM_IN_L14_underPF, LM_IN_L15_underPF, LM_IN_L16_underPF, LM_IN_L1_underPF, LM_IN_L21_underPF, LM_IN_L23_underPF, LM_IN_L26_underPF, LM_IN_L27_underPF, LM_IN_L28_underPF, LM_IN_L29_underPF, LM_IN_L2_underPF, LM_IN_L30_underPF, LM_IN_L31_underPF, LM_IN_L32_underPF, LM_IN_L33_underPF, LM_IN_L34_underPF, LM_IN_L35_underPF, LM_IN_L36_underPF, LM_IN_L37_underPF, LM_IN_L38_underPF, LM_IN_L39_underPF, LM_IN_L3_underPF, LM_IN_L40_underPF, LM_IN_L41_underPF, LM_IN_L42_underPF, LM_IN_L43_underPF, LM_IN_L44_underPF, LM_IN_L45_underPF, LM_IN_L46_underPF, LM_IN_L47_underPF, LM_IN_L48_underPF, LM_IN_L49a_underPF, LM_IN_L4_underPF, LM_IN_L50_underPF, LM_IN_L51_underPF, LM_IN_L52_underPF, LM_IN_L53_underPF, LM_IN_L54_underPF, _
  LM_IN_L55_underPF, LM_IN_L56_underPF, LM_IN_L57a_underPF, LM_IN_L5_underPF, LM_IN_L6_underPF, LM_IN_L7_underPF, LM_IN_L8_underPF, LM_IN_L9_underPF, LM_FL_f1_underPF, LM_FL_f114_underPF, LM_FL_f114a_underPF, LM_FL_f2_underPF, LM_FL_f3_underPF, LM_FL_f4_underPF, LM_FL_f6_underPF, LM_FL_f7_underPF)
' Arrays per lighting scenario
Dim BL_FL_f1: BL_FL_f1=Array(LM_FL_f1_BC1, LM_FL_f1_BC2, LM_FL_f1_BS1, LM_FL_f1_EIprim, LM_FL_f1_LEMK, LM_FL_f1_LRubber1, LM_FL_f1_LRubber2, LM_FL_f1_LRubber3, LM_FL_f1_LRubber4, LM_FL_f1_LUFlipper, LM_FL_f1_Layer1, LM_FL_f1_Parts, LM_FL_f1_Playfield, LM_FL_f1_RSling1, LM_FL_f1_RSling2, LM_FL_f1_cabRails, LM_FL_f1_lpost, LM_FL_f1_rpost, LM_FL_f1_sw19, LM_FL_f1_sw24, LM_FL_f1_sw30, LM_FL_f1_sw45, LM_FL_f1_sw46, LM_FL_f1_underPF)
Dim BL_FL_f10: BL_FL_f10=Array(LM_FL_f10_BC1, LM_FL_f10_BC2, LM_FL_f10_BC3, LM_FL_f10_BR1, LM_FL_f10_BR3, LM_FL_f10_BS1, LM_FL_f10_BS2, LM_FL_f10_BS3, LM_FL_f10_EIprim, LM_FL_f10_Layer1, LM_FL_f10_Layer2, LM_FL_f10_Layer3, LM_FL_f10_Layer4, LM_FL_f10_Parts, LM_FL_f10_Playfield, LM_FL_f10_cabRails, LM_FL_f10_sw33, LM_FL_f10_sw34, LM_FL_f10_sw35, LM_FL_f10_sw41, LM_FL_f10_sw45)
Dim BL_FL_f114: BL_FL_f114=Array(LM_FL_f114_BC3, LM_FL_f114_BR3, LM_FL_f114_BS2, LM_FL_f114_BS3, LM_FL_f114_EIprim, LM_FL_f114_LRubber1, LM_FL_f114_LRubber2, LM_FL_f114_LRubber3, LM_FL_f114_LRubber4, LM_FL_f114_LUFlipper, LM_FL_f114_Layer1, LM_FL_f114_Layer6, LM_FL_f114_Parts, LM_FL_f114_Playfield, LM_FL_f114_RRubber1, LM_FL_f114_RRubber2, LM_FL_f114_RRubber3, LM_FL_f114_RRubber4, LM_FL_f114_cabRails, LM_FL_f114_gate1, LM_FL_f114_rpost, LM_FL_f114_sw19, LM_FL_f114_sw21, LM_FL_f114_sw22, LM_FL_f114_sw28, LM_FL_f114_sw29, LM_FL_f114_sw30, LM_FL_f114_underPF)
Dim BL_FL_f114a: BL_FL_f114a=Array(LM_FL_f114a_EIprim, LM_FL_f114a_LRubber1, LM_FL_f114a_LRubber2, LM_FL_f114a_LRubber3, LM_FL_f114a_LRubber4, LM_FL_f114a_LUFlipper, LM_FL_f114a_Layer1, LM_FL_f114a_Layer6, LM_FL_f114a_Parts, LM_FL_f114a_Playfield, LM_FL_f114a_RRubber3, LM_FL_f114a_cabRails, LM_FL_f114a_gate1, LM_FL_f114a_lpost, LM_FL_f114a_sw19, LM_FL_f114a_sw21, LM_FL_f114a_sw22, LM_FL_f114a_sw24, LM_FL_f114a_sw30, LM_FL_f114a_underPF)
Dim BL_FL_f2: BL_FL_f2=Array(LM_FL_f2_BC1, LM_FL_f2_BC2, LM_FL_f2_BR1, LM_FL_f2_BR2, LM_FL_f2_BR3, LM_FL_f2_BS1, LM_FL_f2_BS2, LM_FL_f2_BS3, LM_FL_f2_EIprim, LM_FL_f2_LUFlipper, LM_FL_f2_Layer1, LM_FL_f2_Layer2, LM_FL_f2_Layer6, LM_FL_f2_Parts, LM_FL_f2_Playfield, LM_FL_f2_cabRails, LM_FL_f2_gate1, LM_FL_f2_sw27, LM_FL_f2_sw29, LM_FL_f2_sw30, LM_FL_f2_sw41, LM_FL_f2_sw45, LM_FL_f2_sw46, LM_FL_f2_underPF)
Dim BL_FL_f3: BL_FL_f3=Array(LM_FL_f3_BC1, LM_FL_f3_BC2, LM_FL_f3_BC3, LM_FL_f3_BR1, LM_FL_f3_BR2, LM_FL_f3_BS1, LM_FL_f3_BS2, LM_FL_f3_BS3, LM_FL_f3_EIprim, LM_FL_f3_Layer1, LM_FL_f3_Layer2, LM_FL_f3_Parts, LM_FL_f3_Playfield, LM_FL_f3_cabRails, LM_FL_f3_sw23, LM_FL_f3_sw27, LM_FL_f3_sw28, LM_FL_f3_sw29, LM_FL_f3_sw33, LM_FL_f3_sw34, LM_FL_f3_sw35, LM_FL_f3_sw41, LM_FL_f3_sw41a, LM_FL_f3_sw46, LM_FL_f3_underPF)
Dim BL_FL_f4: BL_FL_f4=Array(LM_FL_f4_EIprim, LM_FL_f4_Gate2, LM_FL_f4_Layer1, LM_FL_f4_Layer2, LM_FL_f4_Layer3, LM_FL_f4_Layer5, LM_FL_f4_Parts, LM_FL_f4_Playfield, LM_FL_f4_cabRails, LM_FL_f4_diverter, LM_FL_f4_sw30, LM_FL_f4_sw31, LM_FL_f4_sw32, LM_FL_f4_sw41, LM_FL_f4_sw44, LM_FL_f4_underPF)
Dim BL_FL_f5: BL_FL_f5=Array(LM_FL_f5_EIprim, LM_FL_f5_Layer1, LM_FL_f5_Layer2, LM_FL_f5_Layer3, LM_FL_f5_Parts, LM_FL_f5_Playfield, LM_FL_f5_cabRails, LM_FL_f5_sw41)
Dim BL_FL_f6: BL_FL_f6=Array(LM_FL_f6_BC2, LM_FL_f6_BC3, LM_FL_f6_BR2, LM_FL_f6_BR3, LM_FL_f6_EIprim, LM_FL_f6_LSling1, LM_FL_f6_Layer1, LM_FL_f6_Layer3, LM_FL_f6_Layer6, LM_FL_f6_Layer7, LM_FL_f6_Parts, LM_FL_f6_Playfield, LM_FL_f6_RRubber1, LM_FL_f6_RRubber2, LM_FL_f6_RRubber3, LM_FL_f6_RRubber4, LM_FL_f6_cabRails, LM_FL_f6_lpost, LM_FL_f6_rpost, LM_FL_f6_sw19, LM_FL_f6_sw21, LM_FL_f6_sw22, LM_FL_f6_sw23, LM_FL_f6_sw24, LM_FL_f6_sw30, LM_FL_f6_underPF)
Dim BL_FL_f7: BL_FL_f7=Array(LM_FL_f7_BS2, LM_FL_f7_EIprim, LM_FL_f7_Layer1, LM_FL_f7_Layer2, LM_FL_f7_Layer3, LM_FL_f7_Layer6, LM_FL_f7_Parts, LM_FL_f7_Playfield, LM_FL_f7_cabRails, LM_FL_f7_diverter, LM_FL_f7_gate1, LM_FL_f7_sw23, LM_FL_f7_sw30, LM_FL_f7_sw31, LM_FL_f7_sw43, LM_FL_f7_sw44, LM_FL_f7_underPF)
Dim BL_FL_f8: BL_FL_f8=Array(LM_FL_f8_Layer1, LM_FL_f8_Layer2, LM_FL_f8_Layer3, LM_FL_f8_Parts, LM_FL_f8_Playfield, LM_FL_f8_sw30)
Dim BL_FL_f9: BL_FL_f9=Array(LM_FL_f9_BC2, LM_FL_f9_EIprim, LM_FL_f9_Layer1, LM_FL_f9_Layer2, LM_FL_f9_Layer3, LM_FL_f9_Layer5, LM_FL_f9_Parts, LM_FL_f9_Playfield, LM_FL_f9_cabRails, LM_FL_f9_diverter, LM_FL_f9_sw30, LM_FL_f9_sw31, LM_FL_f9_sw44)
Dim BL_GI: BL_GI=Array(LM_GI_BC1, LM_GI_BC2, LM_GI_BC3, LM_GI_BR1, LM_GI_BR2, LM_GI_BR3, LM_GI_BS1, LM_GI_BS2, LM_GI_BS3, LM_GI_EIprim, LM_GI_LEMK, LM_GI_LRubber1, LM_GI_LRubber2, LM_GI_LRubber3, LM_GI_LRubber4, LM_GI_LSling1, LM_GI_LSling2, LM_GI_LUFlipper, LM_GI_Layer1, LM_GI_Layer2, LM_GI_Layer3, LM_GI_Layer4, LM_GI_Layer6, LM_GI_Layer7, LM_GI_Lockdownbar, LM_GI_Parts, LM_GI_Playfield, LM_GI_REMK, LM_GI_RFlipper, LM_GI_RRubber1, LM_GI_RRubber2, LM_GI_RRubber3, LM_GI_RRubber4, LM_GI_cabRails, LM_GI_diverter, LM_GI_gate1, LM_GI_lpost, LM_GI_rpost, LM_GI_sw19, LM_GI_sw21, LM_GI_sw22, LM_GI_sw23, LM_GI_sw24, LM_GI_sw27, LM_GI_sw28, LM_GI_sw29, LM_GI_sw30, LM_GI_sw31, LM_GI_sw33, LM_GI_sw35, LM_GI_sw41, LM_GI_sw43, LM_GI_sw44, LM_GI_sw45, LM_GI_sw46, LM_GI_underPF)
Dim BL_GIS_GI_1: BL_GIS_GI_1=Array(LM_GIS_GI_1_LEMK, LM_GIS_GI_1_LFlipper, LM_GIS_GI_1_Layer1, LM_GIS_GI_1_Parts, LM_GIS_GI_1_Playfield, LM_GIS_GI_1_RFlipper, LM_GIS_GI_1_underPF)
Dim BL_GIS_GI_10: BL_GIS_GI_10=Array(LM_GIS_GI_10_LRubber1, LM_GIS_GI_10_LRubber2, LM_GIS_GI_10_LRubber3, LM_GIS_GI_10_LRubber4, LM_GIS_GI_10_LUFlipper, LM_GIS_GI_10_Layer1, LM_GIS_GI_10_Parts, LM_GIS_GI_10_Playfield, LM_GIS_GI_10_cabRails, LM_GIS_GI_10_lpost, LM_GIS_GI_10_sw19, LM_GIS_GI_10_sw46, LM_GIS_GI_10_underPF)
Dim BL_GIS_GI_13: BL_GIS_GI_13=Array(LM_GIS_GI_13_BC2, LM_GIS_GI_13_BR2, LM_GIS_GI_13_Layer6, LM_GIS_GI_13_Parts, LM_GIS_GI_13_Playfield, LM_GIS_GI_13_REMK, LM_GIS_GI_13_RRubber1, LM_GIS_GI_13_RRubber2, LM_GIS_GI_13_RRubber3, LM_GIS_GI_13_RRubber4, LM_GIS_GI_13_cabRails, LM_GIS_GI_13_rpost, LM_GIS_GI_13_sw24, LM_GIS_GI_13_underPF)
Dim BL_GIS_GI_15: BL_GIS_GI_15=Array(LM_GIS_GI_15_LRubber1, LM_GIS_GI_15_LRubber2, LM_GIS_GI_15_LRubber3, LM_GIS_GI_15_LRubber4, LM_GIS_GI_15_LUFlipper, LM_GIS_GI_15_Layer1, LM_GIS_GI_15_Parts, LM_GIS_GI_15_Playfield, LM_GIS_GI_15_lpost, LM_GIS_GI_15_sw46, LM_GIS_GI_15_underPF)
Dim BL_GIS_GI_16: BL_GIS_GI_16=Array(LM_GIS_GI_16_LFlipper, LM_GIS_GI_16_Parts, LM_GIS_GI_16_Playfield, LM_GIS_GI_16_REMK, LM_GIS_GI_16_RFlipper, LM_GIS_GI_16_RRubber1, LM_GIS_GI_16_RRubber2, LM_GIS_GI_16_RRubber3, LM_GIS_GI_16_RRubber4, LM_GIS_GI_16_RSling1, LM_GIS_GI_16_RSling2, LM_GIS_GI_16_cabRails, LM_GIS_GI_16_sw14, LM_GIS_GI_16_sw18, LM_GIS_GI_16_sw24, LM_GIS_GI_16_underPF)
Dim BL_GIS_GI_17: BL_GIS_GI_17=Array(LM_GIS_GI_17_LEMK, LM_GIS_GI_17_LFlipper, LM_GIS_GI_17_LRubber1, LM_GIS_GI_17_LRubber2, LM_GIS_GI_17_LRubber3, LM_GIS_GI_17_LRubber4, LM_GIS_GI_17_LSling1, LM_GIS_GI_17_LSling2, LM_GIS_GI_17_Layer1, LM_GIS_GI_17_Parts, LM_GIS_GI_17_Playfield, LM_GIS_GI_17_RFlipper, LM_GIS_GI_17_cabRails, LM_GIS_GI_17_sw14, LM_GIS_GI_17_sw17, LM_GIS_GI_17_sw18, LM_GIS_GI_17_underPF)
Dim BL_GIS_GI_2: BL_GIS_GI_2=Array(LM_GIS_GI_2_LFlipper, LM_GIS_GI_2_Layer1, LM_GIS_GI_2_Parts, LM_GIS_GI_2_Playfield, LM_GIS_GI_2_REMK, LM_GIS_GI_2_RFlipper, LM_GIS_GI_2_RSling1, LM_GIS_GI_2_sw17)
Dim BL_GIS_GI_20: BL_GIS_GI_20=Array(LM_GIS_GI_20_LEMK, LM_GIS_GI_20_LFlipper, LM_GIS_GI_20_LRubber1, LM_GIS_GI_20_LRubber2, LM_GIS_GI_20_LRubber3, LM_GIS_GI_20_LRubber4, LM_GIS_GI_20_LSling1, LM_GIS_GI_20_LSling2, LM_GIS_GI_20_Layer1, LM_GIS_GI_20_Parts, LM_GIS_GI_20_Playfield, LM_GIS_GI_20_RFlipper, LM_GIS_GI_20_cabRails, LM_GIS_GI_20_lpost, LM_GIS_GI_20_rpost, LM_GIS_GI_20_sw17, LM_GIS_GI_20_sw18, LM_GIS_GI_20_sw19, LM_GIS_GI_20_underPF)
Dim BL_GIS_GI_29: BL_GIS_GI_29=Array(LM_GIS_GI_29_LFlipper, LM_GIS_GI_29_Layer1, LM_GIS_GI_29_Parts, LM_GIS_GI_29_Playfield, LM_GIS_GI_29_REMK, LM_GIS_GI_29_RRubber1, LM_GIS_GI_29_RRubber2, LM_GIS_GI_29_RRubber3, LM_GIS_GI_29_RRubber4, LM_GIS_GI_29_RSling1)
Dim BL_IN_L1: BL_IN_L1=Array(LM_IN_L1_Playfield, LM_IN_L1_underPF)
Dim BL_IN_L10: BL_IN_L10=Array(LM_IN_L10_Playfield, LM_IN_L10_underPF)
Dim BL_IN_L11: BL_IN_L11=Array(LM_IN_L11_Playfield, LM_IN_L11_underPF)
Dim BL_IN_L12: BL_IN_L12=Array(LM_IN_L12_Playfield, LM_IN_L12_underPF)
Dim BL_IN_L13: BL_IN_L13=Array(LM_IN_L13_Parts, LM_IN_L13_Playfield, LM_IN_L13_REMK, LM_IN_L13_underPF)
Dim BL_IN_L14: BL_IN_L14=Array(LM_IN_L14_Playfield, LM_IN_L14_underPF)
Dim BL_IN_L15: BL_IN_L15=Array(LM_IN_L15_Layer1, LM_IN_L15_Parts, LM_IN_L15_Playfield, LM_IN_L15_underPF)
Dim BL_IN_L16: BL_IN_L16=Array(LM_IN_L16_LRubber1, LM_IN_L16_LRubber2, LM_IN_L16_LRubber3, LM_IN_L16_LRubber4, LM_IN_L16_Layer1, LM_IN_L16_Parts, LM_IN_L16_Playfield, LM_IN_L16_lpost, LM_IN_L16_underPF)
Dim BL_IN_L17: BL_IN_L17=Array(LM_IN_L17_BC1, LM_IN_L17_BC2, LM_IN_L17_BS2, LM_IN_L17_EIprim, LM_IN_L17_Layer1, LM_IN_L17_Layer2, LM_IN_L17_Parts, LM_IN_L17_Playfield, LM_IN_L17_cabRails, LM_IN_L17_sw27, LM_IN_L17_sw28)
Dim BL_IN_L18: BL_IN_L18=Array(LM_IN_L18_BS2, LM_IN_L18_BS3, LM_IN_L18_EIprim, LM_IN_L18_Layer1, LM_IN_L18_Layer2, LM_IN_L18_Parts, LM_IN_L18_Playfield, LM_IN_L18_cabRails, LM_IN_L18_sw28)
Dim BL_IN_L19: BL_IN_L19=Array(LM_IN_L19_BS2, LM_IN_L19_BS3, LM_IN_L19_EIprim, LM_IN_L19_Layer1, LM_IN_L19_Parts, LM_IN_L19_Playfield, LM_IN_L19_cabRails)
Dim BL_IN_L2: BL_IN_L2=Array(LM_IN_L2_Parts, LM_IN_L2_Playfield, LM_IN_L2_underPF)
Dim BL_IN_L20: BL_IN_L20=Array(LM_IN_L20_BC2, LM_IN_L20_BS2, LM_IN_L20_EIprim, LM_IN_L20_Layer1, LM_IN_L20_Layer2, LM_IN_L20_Parts, LM_IN_L20_Playfield, LM_IN_L20_cabRails, LM_IN_L20_sw27, LM_IN_L20_sw41)
Dim BL_IN_L21: BL_IN_L21=Array(LM_IN_L21_BC2, LM_IN_L21_BS1, LM_IN_L21_BS2, LM_IN_L21_BS3, LM_IN_L21_EIprim, LM_IN_L21_Layer1, LM_IN_L21_Layer2, LM_IN_L21_Parts, LM_IN_L21_Playfield, LM_IN_L21_cabRails, LM_IN_L21_underPF)
Dim BL_IN_L22: BL_IN_L22=Array(LM_IN_L22_BS2, LM_IN_L22_BS3, LM_IN_L22_EIprim, LM_IN_L22_Layer1, LM_IN_L22_Parts, LM_IN_L22_Playfield, LM_IN_L22_cabRails)
Dim BL_IN_L23: BL_IN_L23=Array(LM_IN_L23_BC2, LM_IN_L23_BS1, LM_IN_L23_BS2, LM_IN_L23_EIprim, LM_IN_L23_Layer1, LM_IN_L23_Layer2, LM_IN_L23_Parts, LM_IN_L23_Playfield, LM_IN_L23_sw27, LM_IN_L23_sw41, LM_IN_L23_underPF)
Dim BL_IN_L24: BL_IN_L24=Array(LM_IN_L24_BC1, LM_IN_L24_BC2, LM_IN_L24_BR1, LM_IN_L24_BS1, LM_IN_L24_EIprim, LM_IN_L24_Layer1, LM_IN_L24_Parts, LM_IN_L24_Playfield, LM_IN_L24_cabRails, LM_IN_L24_sw27, LM_IN_L24_sw28)
Dim BL_IN_L25: BL_IN_L25=Array(LM_IN_L25_BC2, LM_IN_L25_BS3, LM_IN_L25_EIprim, LM_IN_L25_Layer1, LM_IN_L25_Parts, LM_IN_L25_Playfield, LM_IN_L25_cabRails)
Dim BL_IN_L26: BL_IN_L26=Array(LM_IN_L26_underPF)
Dim BL_IN_L27: BL_IN_L27=Array(LM_IN_L27_Parts, LM_IN_L27_underPF)
Dim BL_IN_L28: BL_IN_L28=Array(LM_IN_L28_Layer6, LM_IN_L28_Parts, LM_IN_L28_Playfield, LM_IN_L28_RRubber1, LM_IN_L28_RRubber2, LM_IN_L28_RRubber3, LM_IN_L28_RRubber4, LM_IN_L28_sw24, LM_IN_L28_underPF)
Dim BL_IN_L29: BL_IN_L29=Array(LM_IN_L29_Playfield, LM_IN_L29_rpost, LM_IN_L29_underPF)
Dim BL_IN_L3: BL_IN_L3=Array(LM_IN_L3_Parts, LM_IN_L3_Playfield, LM_IN_L3_underPF)
Dim BL_IN_L30: BL_IN_L30=Array(LM_IN_L30_Layer6, LM_IN_L30_Parts, LM_IN_L30_Playfield, LM_IN_L30_RRubber1, LM_IN_L30_RRubber2, LM_IN_L30_RRubber3, LM_IN_L30_RRubber4, LM_IN_L30_rpost, LM_IN_L30_underPF)
Dim BL_IN_L31: BL_IN_L31=Array(LM_IN_L31_Layer6, LM_IN_L31_Parts, LM_IN_L31_Playfield, LM_IN_L31_RRubber1, LM_IN_L31_RRubber2, LM_IN_L31_RRubber3, LM_IN_L31_RRubber4, LM_IN_L31_rpost, LM_IN_L31_underPF)
Dim BL_IN_L32: BL_IN_L32=Array(LM_IN_L32_LFlipper, LM_IN_L32_Parts, LM_IN_L32_Playfield, LM_IN_L32_RFlipper, LM_IN_L32_underPF)
Dim BL_IN_L33: BL_IN_L33=Array(LM_IN_L33_Playfield, LM_IN_L33_underPF)
Dim BL_IN_L34: BL_IN_L34=Array(LM_IN_L34_Layer1, LM_IN_L34_Parts, LM_IN_L34_Playfield, LM_IN_L34_underPF)
Dim BL_IN_L35: BL_IN_L35=Array(LM_IN_L35_Layer1, LM_IN_L35_Playfield, LM_IN_L35_underPF)
Dim BL_IN_L36: BL_IN_L36=Array(LM_IN_L36_Layer1, LM_IN_L36_Parts, LM_IN_L36_Playfield, LM_IN_L36_underPF)
Dim BL_IN_L37: BL_IN_L37=Array(LM_IN_L37_Layer1, LM_IN_L37_Playfield, LM_IN_L37_lpost, LM_IN_L37_underPF)
Dim BL_IN_L38: BL_IN_L38=Array(LM_IN_L38_LRubber1, LM_IN_L38_LRubber2, LM_IN_L38_LRubber3, LM_IN_L38_LRubber4, LM_IN_L38_Layer1, LM_IN_L38_Parts, LM_IN_L38_Playfield, LM_IN_L38_underPF)
Dim BL_IN_L39: BL_IN_L39=Array(LM_IN_L39_Layer1, LM_IN_L39_Parts, LM_IN_L39_Playfield, LM_IN_L39_underPF)
Dim BL_IN_L4: BL_IN_L4=Array(LM_IN_L4_Parts, LM_IN_L4_Playfield, LM_IN_L4_underPF)
Dim BL_IN_L40: BL_IN_L40=Array(LM_IN_L40_Playfield, LM_IN_L40_underPF)
Dim BL_IN_L41: BL_IN_L41=Array(LM_IN_L41_BC1, LM_IN_L41_BC2, LM_IN_L41_BR1, LM_IN_L41_BS1, LM_IN_L41_BS2, LM_IN_L41_BS3, LM_IN_L41_Layer1, LM_IN_L41_Layer2, LM_IN_L41_Parts, LM_IN_L41_Playfield, LM_IN_L41_sw27, LM_IN_L41_sw41, LM_IN_L41_underPF)
Dim BL_IN_L42: BL_IN_L42=Array(LM_IN_L42_BC1, LM_IN_L42_BC2, LM_IN_L42_BC3, LM_IN_L42_BR1, LM_IN_L42_BR2, LM_IN_L42_BR3, LM_IN_L42_BS1, LM_IN_L42_BS3, LM_IN_L42_Layer1, LM_IN_L42_Parts, LM_IN_L42_Playfield, LM_IN_L42_underPF)
Dim BL_IN_L43: BL_IN_L43=Array(LM_IN_L43_BC1, LM_IN_L43_BC2, LM_IN_L43_BR1, LM_IN_L43_BR2, LM_IN_L43_BS1, LM_IN_L43_BS2, LM_IN_L43_BS3, LM_IN_L43_EIprim, LM_IN_L43_Layer1, LM_IN_L43_Parts, LM_IN_L43_Playfield, LM_IN_L43_sw27, LM_IN_L43_sw28, LM_IN_L43_sw29, LM_IN_L43_sw41, LM_IN_L43_underPF)
Dim BL_IN_L44: BL_IN_L44=Array(LM_IN_L44_Playfield, LM_IN_L44_underPF)
Dim BL_IN_L45: BL_IN_L45=Array(LM_IN_L45_Playfield, LM_IN_L45_underPF)
Dim BL_IN_L46: BL_IN_L46=Array(LM_IN_L46_Parts, LM_IN_L46_Playfield, LM_IN_L46_underPF)
Dim BL_IN_L47: BL_IN_L47=Array(LM_IN_L47_Parts, LM_IN_L47_Playfield, LM_IN_L47_underPF)
Dim BL_IN_L48: BL_IN_L48=Array(LM_IN_L48_BC2, LM_IN_L48_Parts, LM_IN_L48_Playfield, LM_IN_L48_sw30, LM_IN_L48_underPF)
Dim BL_IN_L49a: BL_IN_L49a=Array(LM_IN_L49a_EIprim, LM_IN_L49a_Layer1, LM_IN_L49a_Lockdownbar, LM_IN_L49a_Parts, LM_IN_L49a_Playfield, LM_IN_L49a_cabRails, LM_IN_L49a_sw44, LM_IN_L49a_underPF)
Dim BL_IN_L5: BL_IN_L5=Array(LM_IN_L5_Parts, LM_IN_L5_Playfield, LM_IN_L5_underPF)
Dim BL_IN_L50: BL_IN_L50=Array(LM_IN_L50_Parts, LM_IN_L50_underPF)
Dim BL_IN_L51: BL_IN_L51=Array(LM_IN_L51_Playfield, LM_IN_L51_underPF)
Dim BL_IN_L52: BL_IN_L52=Array(LM_IN_L52_Parts, LM_IN_L52_Playfield, LM_IN_L52_underPF)
Dim BL_IN_L53: BL_IN_L53=Array(LM_IN_L53_Parts, LM_IN_L53_Playfield, LM_IN_L53_sw30, LM_IN_L53_underPF)
Dim BL_IN_L54: BL_IN_L54=Array(LM_IN_L54_Layer1, LM_IN_L54_Parts, LM_IN_L54_Playfield, LM_IN_L54_sw30, LM_IN_L54_underPF)
Dim BL_IN_L55: BL_IN_L55=Array(LM_IN_L55_Layer1, LM_IN_L55_Parts, LM_IN_L55_Playfield, LM_IN_L55_underPF)
Dim BL_IN_L56: BL_IN_L56=Array(LM_IN_L56_Parts, LM_IN_L56_Playfield, LM_IN_L56_sw30, LM_IN_L56_underPF)
Dim BL_IN_L57a: BL_IN_L57a=Array(LM_IN_L57a_Layer1, LM_IN_L57a_Layer6, LM_IN_L57a_Parts, LM_IN_L57a_Playfield, LM_IN_L57a_cabRails, LM_IN_L57a_sw23, LM_IN_L57a_underPF)
Dim BL_IN_L6: BL_IN_L6=Array(LM_IN_L6_Parts, LM_IN_L6_Playfield, LM_IN_L6_sw24, LM_IN_L6_underPF)
Dim BL_IN_L7: BL_IN_L7=Array(LM_IN_L7_BC1, LM_IN_L7_BC2, LM_IN_L7_BR1, LM_IN_L7_BS1, LM_IN_L7_BS2, LM_IN_L7_EIprim, LM_IN_L7_Layer1, LM_IN_L7_Layer2, LM_IN_L7_Parts, LM_IN_L7_Playfield, LM_IN_L7_sw41, LM_IN_L7_underPF)
Dim BL_IN_L8: BL_IN_L8=Array(LM_IN_L8_BC1, LM_IN_L8_BC2, LM_IN_L8_BC3, LM_IN_L8_BR1, LM_IN_L8_BR2, LM_IN_L8_BR3, LM_IN_L8_BS1, LM_IN_L8_BS2, LM_IN_L8_BS3, LM_IN_L8_Layer1, LM_IN_L8_Layer2, LM_IN_L8_Parts, LM_IN_L8_Playfield, LM_IN_L8_sw27, LM_IN_L8_sw28, LM_IN_L8_sw41, LM_IN_L8_underPF)
Dim BL_IN_L9: BL_IN_L9=Array(LM_IN_L9_LEMK, LM_IN_L9_LSling1, LM_IN_L9_Parts, LM_IN_L9_Playfield, LM_IN_L9_underPF)
Dim BL_Room: BL_Room=Array(BM_BC1, BM_BC2, BM_BC3, BM_BR1, BM_BR2, BM_BR3, BM_BS1, BM_BS2, BM_BS3, BM_EIprim, BM_Gate2, BM_LEMK, BM_LFlipper, BM_LRubber1, BM_LRubber2, BM_LRubber3, BM_LRubber4, BM_LSling1, BM_LSling2, BM_LUFlipper, BM_Layer1, BM_Layer2, BM_Layer3, BM_Layer4, BM_Layer5, BM_Layer6, BM_Layer7, BM_Lockdownbar, BM_Parts, BM_Playfield, BM_REMK, BM_RFlipper, BM_RRubber1, BM_RRubber2, BM_RRubber3, BM_RRubber4, BM_RSling1, BM_RSling2, BM_cabRails, BM_diverter, BM_gate1, BM_lpost, BM_rpost, BM_sw14, BM_sw15, BM_sw16, BM_sw17, BM_sw18, BM_sw19, BM_sw21, BM_sw22, BM_sw23, BM_sw24, BM_sw27, BM_sw28, BM_sw29, BM_sw30, BM_sw31, BM_sw32, BM_sw33, BM_sw34, BM_sw35, BM_sw36, BM_sw41, BM_sw41a, BM_sw43, BM_sw44, BM_sw45, BM_sw46, BM_sw50, BM_underPF)
' Global arrays
'Dim BG_Bakemap: BG_Bakemap=Array(BM_BC1, BM_BC2, BM_BC3, BM_BR1, BM_BR2, BM_BR3, BM_BS1, BM_BS2, BM_BS3, BM_EIprim, BM_Gate2, BM_LEMK, BM_LFlipper, BM_LRubber1, BM_LRubber2, BM_LRubber3, BM_LRubber4, BM_LSling1, BM_LSling2, BM_LUFlipper, BM_Layer1, BM_Layer2, BM_Layer3, BM_Layer4, BM_Layer5, BM_Layer6, BM_Layer7, BM_Lockdownbar, BM_Parts, BM_Playfield, BM_REMK, BM_RFlipper, BM_RRubber1, BM_RRubber2, BM_RRubber3, BM_RRubber4, BM_RSling1, BM_RSling2, BM_cabRails, BM_diverter, BM_gate1, BM_lpost, BM_rpost, BM_sw14, BM_sw15, BM_sw16, BM_sw17, BM_sw18, BM_sw19, BM_sw21, BM_sw22, BM_sw23, BM_sw24, BM_sw27, BM_sw28, BM_sw29, BM_sw30, BM_sw31, BM_sw32, BM_sw33, BM_sw34, BM_sw35, BM_sw36, BM_sw41, BM_sw41a, BM_sw43, BM_sw44, BM_sw45, BM_sw46, BM_sw50, BM_underPF)
'Dim BG_Lightmap: BG_Lightmap=Array(LM_FL_f1_BC1, LM_FL_f1_BC2, LM_FL_f1_BS1, LM_FL_f1_EIprim, LM_FL_f1_LEMK, LM_FL_f1_LRubber1, LM_FL_f1_LRubber2, LM_FL_f1_LRubber3, LM_FL_f1_LRubber4, LM_FL_f1_LUFlipper, LM_FL_f1_Layer1, LM_FL_f1_Parts, LM_FL_f1_Playfield, LM_FL_f1_RSling1, LM_FL_f1_RSling2, LM_FL_f1_cabRails, LM_FL_f1_lpost, LM_FL_f1_rpost, LM_FL_f1_sw19, LM_FL_f1_sw24, LM_FL_f1_sw30, LM_FL_f1_sw45, LM_FL_f1_sw46, LM_FL_f1_underPF, LM_FL_f10_BC1, LM_FL_f10_BC2, LM_FL_f10_BC3, LM_FL_f10_BR1, LM_FL_f10_BR3, LM_FL_f10_BS1, LM_FL_f10_BS2, LM_FL_f10_BS3, LM_FL_f10_EIprim, LM_FL_f10_Layer1, LM_FL_f10_Layer2, LM_FL_f10_Layer3, LM_FL_f10_Layer4, LM_FL_f10_Parts, LM_FL_f10_Playfield, LM_FL_f10_cabRails, LM_FL_f10_sw33, LM_FL_f10_sw34, LM_FL_f10_sw35, LM_FL_f10_sw41, LM_FL_f10_sw45, LM_FL_f114_BC3, LM_FL_f114_BR3, LM_FL_f114_BS2, LM_FL_f114_BS3, LM_FL_f114_EIprim, LM_FL_f114_LRubber1, LM_FL_f114_LRubber2, LM_FL_f114_LRubber3, LM_FL_f114_LRubber4, LM_FL_f114_LUFlipper, LM_FL_f114_Layer1, LM_FL_f114_Layer6, _
' LM_FL_f114_Parts, LM_FL_f114_Playfield, LM_FL_f114_RRubber1, LM_FL_f114_RRubber2, LM_FL_f114_RRubber3, LM_FL_f114_RRubber4, LM_FL_f114_cabRails, LM_FL_f114_gate1, LM_FL_f114_rpost, LM_FL_f114_sw19, LM_FL_f114_sw21, LM_FL_f114_sw22, LM_FL_f114_sw28, LM_FL_f114_sw29, LM_FL_f114_sw30, LM_FL_f114_underPF, LM_FL_f114a_EIprim, LM_FL_f114a_LRubber1, LM_FL_f114a_LRubber2, LM_FL_f114a_LRubber3, LM_FL_f114a_LRubber4, LM_FL_f114a_LUFlipper, LM_FL_f114a_Layer1, LM_FL_f114a_Layer6, LM_FL_f114a_Parts, LM_FL_f114a_Playfield, LM_FL_f114a_RRubber3, LM_FL_f114a_cabRails, LM_FL_f114a_gate1, LM_FL_f114a_lpost, LM_FL_f114a_sw19, LM_FL_f114a_sw21, LM_FL_f114a_sw22, LM_FL_f114a_sw24, LM_FL_f114a_sw30, LM_FL_f114a_underPF, LM_FL_f2_BC1, LM_FL_f2_BC2, LM_FL_f2_BR1, LM_FL_f2_BR2, LM_FL_f2_BR3, LM_FL_f2_BS1, LM_FL_f2_BS2, LM_FL_f2_BS3, LM_FL_f2_EIprim, LM_FL_f2_LUFlipper, LM_FL_f2_Layer1, LM_FL_f2_Layer2, LM_FL_f2_Layer6, LM_FL_f2_Parts, LM_FL_f2_Playfield, LM_FL_f2_cabRails, LM_FL_f2_gate1, LM_FL_f2_sw27, LM_FL_f2_sw29, LM_FL_f2_sw30, _
' LM_FL_f2_sw41, LM_FL_f2_sw45, LM_FL_f2_sw46, LM_FL_f2_underPF, LM_FL_f3_BC1, LM_FL_f3_BC2, LM_FL_f3_BC3, LM_FL_f3_BR1, LM_FL_f3_BR2, LM_FL_f3_BS1, LM_FL_f3_BS2, LM_FL_f3_BS3, LM_FL_f3_EIprim, LM_FL_f3_Layer1, LM_FL_f3_Layer2, LM_FL_f3_Parts, LM_FL_f3_Playfield, LM_FL_f3_cabRails, LM_FL_f3_sw23, LM_FL_f3_sw27, LM_FL_f3_sw28, LM_FL_f3_sw29, LM_FL_f3_sw33, LM_FL_f3_sw34, LM_FL_f3_sw35, LM_FL_f3_sw41, LM_FL_f3_sw41a, LM_FL_f3_sw46, LM_FL_f3_underPF, LM_FL_f4_EIprim, LM_FL_f4_Gate2, LM_FL_f4_Layer1, LM_FL_f4_Layer2, LM_FL_f4_Layer3, LM_FL_f4_Layer5, LM_FL_f4_Parts, LM_FL_f4_Playfield, LM_FL_f4_cabRails, LM_FL_f4_diverter, LM_FL_f4_sw30, LM_FL_f4_sw31, LM_FL_f4_sw32, LM_FL_f4_sw41, LM_FL_f4_sw44, LM_FL_f4_underPF, LM_FL_f5_EIprim, LM_FL_f5_Layer1, LM_FL_f5_Layer2, LM_FL_f5_Layer3, LM_FL_f5_Parts, LM_FL_f5_Playfield, LM_FL_f5_cabRails, LM_FL_f5_sw41, LM_FL_f6_BC2, LM_FL_f6_BC3, LM_FL_f6_BR2, LM_FL_f6_BR3, LM_FL_f6_EIprim, LM_FL_f6_LSling1, LM_FL_f6_Layer1, LM_FL_f6_Layer3, LM_FL_f6_Layer6, LM_FL_f6_Layer7, _
' LM_FL_f6_Parts, LM_FL_f6_Playfield, LM_FL_f6_RRubber1, LM_FL_f6_RRubber2, LM_FL_f6_RRubber3, LM_FL_f6_RRubber4, LM_FL_f6_cabRails, LM_FL_f6_lpost, LM_FL_f6_rpost, LM_FL_f6_sw19, LM_FL_f6_sw21, LM_FL_f6_sw22, LM_FL_f6_sw23, LM_FL_f6_sw24, LM_FL_f6_sw30, LM_FL_f6_underPF, LM_FL_f7_BS2, LM_FL_f7_EIprim, LM_FL_f7_Layer1, LM_FL_f7_Layer2, LM_FL_f7_Layer3, LM_FL_f7_Layer6, LM_FL_f7_Parts, LM_FL_f7_Playfield, LM_FL_f7_cabRails, LM_FL_f7_diverter, LM_FL_f7_gate1, LM_FL_f7_sw23, LM_FL_f7_sw30, LM_FL_f7_sw31, LM_FL_f7_sw43, LM_FL_f7_sw44, LM_FL_f7_underPF, LM_FL_f8_Layer1, LM_FL_f8_Layer2, LM_FL_f8_Layer3, LM_FL_f8_Parts, LM_FL_f8_Playfield, LM_FL_f8_sw30, LM_FL_f9_BC2, LM_FL_f9_EIprim, LM_FL_f9_Layer1, LM_FL_f9_Layer2, LM_FL_f9_Layer3, LM_FL_f9_Layer5, LM_FL_f9_Parts, LM_FL_f9_Playfield, LM_FL_f9_cabRails, LM_FL_f9_diverter, LM_FL_f9_sw30, LM_FL_f9_sw31, LM_FL_f9_sw44, LM_GI_BC1, LM_GI_BC2, LM_GI_BC3, LM_GI_BR1, LM_GI_BR2, LM_GI_BR3, LM_GI_BS1, LM_GI_BS2, LM_GI_BS3, LM_GI_EIprim, LM_GI_LEMK, LM_GI_LRubber1, _
' LM_GI_LRubber2, LM_GI_LRubber3, LM_GI_LRubber4, LM_GI_LSling1, LM_GI_LSling2, LM_GI_LUFlipper, LM_GI_Layer1, LM_GI_Layer2, LM_GI_Layer3, LM_GI_Layer4, LM_GI_Layer6, LM_GI_Layer7, LM_GI_Lockdownbar, LM_GI_Parts, LM_GI_Playfield, LM_GI_REMK, LM_GI_RFlipper, LM_GI_RRubber1, LM_GI_RRubber2, LM_GI_RRubber3, LM_GI_RRubber4, LM_GI_cabRails, LM_GI_diverter, LM_GI_gate1, LM_GI_lpost, LM_GI_rpost, LM_GI_sw19, LM_GI_sw21, LM_GI_sw22, LM_GI_sw23, LM_GI_sw24, LM_GI_sw27, LM_GI_sw28, LM_GI_sw29, LM_GI_sw30, LM_GI_sw31, LM_GI_sw33, LM_GI_sw35, LM_GI_sw41, LM_GI_sw43, LM_GI_sw44, LM_GI_sw45, LM_GI_sw46, LM_GI_underPF, LM_GIS_GI_1_LEMK, LM_GIS_GI_1_LFlipper, LM_GIS_GI_1_Layer1, LM_GIS_GI_1_Parts, LM_GIS_GI_1_Playfield, LM_GIS_GI_1_RFlipper, LM_GIS_GI_1_underPF, LM_GIS_GI_10_LRubber1, LM_GIS_GI_10_LRubber2, LM_GIS_GI_10_LRubber3, LM_GIS_GI_10_LRubber4, LM_GIS_GI_10_LUFlipper, LM_GIS_GI_10_Layer1, LM_GIS_GI_10_Parts, LM_GIS_GI_10_Playfield, LM_GIS_GI_10_cabRails, LM_GIS_GI_10_lpost, LM_GIS_GI_10_sw19, LM_GIS_GI_10_sw46, _
' LM_GIS_GI_10_underPF, LM_GIS_GI_13_BC2, LM_GIS_GI_13_BR2, LM_GIS_GI_13_Layer6, LM_GIS_GI_13_Parts, LM_GIS_GI_13_Playfield, LM_GIS_GI_13_REMK, LM_GIS_GI_13_RRubber1, LM_GIS_GI_13_RRubber2, LM_GIS_GI_13_RRubber3, LM_GIS_GI_13_RRubber4, LM_GIS_GI_13_cabRails, LM_GIS_GI_13_rpost, LM_GIS_GI_13_sw24, LM_GIS_GI_13_underPF, LM_GIS_GI_15_LRubber1, LM_GIS_GI_15_LRubber2, LM_GIS_GI_15_LRubber3, LM_GIS_GI_15_LRubber4, LM_GIS_GI_15_LUFlipper, LM_GIS_GI_15_Layer1, LM_GIS_GI_15_Parts, LM_GIS_GI_15_Playfield, LM_GIS_GI_15_lpost, LM_GIS_GI_15_sw46, LM_GIS_GI_15_underPF, LM_GIS_GI_16_LFlipper, LM_GIS_GI_16_Parts, LM_GIS_GI_16_Playfield, LM_GIS_GI_16_REMK, LM_GIS_GI_16_RFlipper, LM_GIS_GI_16_RRubber1, LM_GIS_GI_16_RRubber2, LM_GIS_GI_16_RRubber3, LM_GIS_GI_16_RRubber4, LM_GIS_GI_16_RSling1, LM_GIS_GI_16_RSling2, LM_GIS_GI_16_cabRails, LM_GIS_GI_16_sw14, LM_GIS_GI_16_sw18, LM_GIS_GI_16_sw24, LM_GIS_GI_16_underPF, LM_GIS_GI_17_LEMK, LM_GIS_GI_17_LFlipper, LM_GIS_GI_17_LRubber1, LM_GIS_GI_17_LRubber2, LM_GIS_GI_17_LRubber3, _
' LM_GIS_GI_17_LRubber4, LM_GIS_GI_17_LSling1, LM_GIS_GI_17_LSling2, LM_GIS_GI_17_Layer1, LM_GIS_GI_17_Parts, LM_GIS_GI_17_Playfield, LM_GIS_GI_17_RFlipper, LM_GIS_GI_17_cabRails, LM_GIS_GI_17_sw14, LM_GIS_GI_17_sw17, LM_GIS_GI_17_sw18, LM_GIS_GI_17_underPF, LM_GIS_GI_2_LFlipper, LM_GIS_GI_2_Layer1, LM_GIS_GI_2_Parts, LM_GIS_GI_2_Playfield, LM_GIS_GI_2_REMK, LM_GIS_GI_2_RFlipper, LM_GIS_GI_2_RSling1, LM_GIS_GI_2_sw17, LM_GIS_GI_20_LEMK, LM_GIS_GI_20_LFlipper, LM_GIS_GI_20_LRubber1, LM_GIS_GI_20_LRubber2, LM_GIS_GI_20_LRubber3, LM_GIS_GI_20_LRubber4, LM_GIS_GI_20_LSling1, LM_GIS_GI_20_LSling2, LM_GIS_GI_20_Layer1, LM_GIS_GI_20_Parts, LM_GIS_GI_20_Playfield, LM_GIS_GI_20_RFlipper, LM_GIS_GI_20_cabRails, LM_GIS_GI_20_lpost, LM_GIS_GI_20_rpost, LM_GIS_GI_20_sw17, LM_GIS_GI_20_sw18, LM_GIS_GI_20_sw19, LM_GIS_GI_20_underPF, LM_GIS_GI_29_LFlipper, LM_GIS_GI_29_Layer1, LM_GIS_GI_29_Parts, LM_GIS_GI_29_Playfield, LM_GIS_GI_29_REMK, LM_GIS_GI_29_RRubber1, LM_GIS_GI_29_RRubber2, LM_GIS_GI_29_RRubber3, _
' LM_GIS_GI_29_RRubber4, LM_GIS_GI_29_RSling1, LM_IN_L1_Playfield, LM_IN_L1_underPF, LM_IN_L10_Playfield, LM_IN_L10_underPF, LM_IN_L11_Playfield, LM_IN_L11_underPF, LM_IN_L12_Playfield, LM_IN_L12_underPF, LM_IN_L13_Parts, LM_IN_L13_Playfield, LM_IN_L13_REMK, LM_IN_L13_underPF, LM_IN_L14_Playfield, LM_IN_L14_underPF, LM_IN_L15_Layer1, LM_IN_L15_Parts, LM_IN_L15_Playfield, LM_IN_L15_underPF, LM_IN_L16_LRubber1, LM_IN_L16_LRubber2, LM_IN_L16_LRubber3, LM_IN_L16_LRubber4, LM_IN_L16_Layer1, LM_IN_L16_Parts, LM_IN_L16_Playfield, LM_IN_L16_lpost, LM_IN_L16_underPF, LM_IN_L17_BC1, LM_IN_L17_BC2, LM_IN_L17_BS2, LM_IN_L17_EIprim, LM_IN_L17_Layer1, LM_IN_L17_Layer2, LM_IN_L17_Parts, LM_IN_L17_Playfield, LM_IN_L17_cabRails, LM_IN_L17_sw27, LM_IN_L17_sw28, LM_IN_L18_BS2, LM_IN_L18_BS3, LM_IN_L18_EIprim, LM_IN_L18_Layer1, LM_IN_L18_Layer2, LM_IN_L18_Parts, LM_IN_L18_Playfield, LM_IN_L18_cabRails, LM_IN_L18_sw28, LM_IN_L19_BS2, LM_IN_L19_BS3, LM_IN_L19_EIprim, LM_IN_L19_Layer1, LM_IN_L19_Parts, LM_IN_L19_Playfield, _
' LM_IN_L19_cabRails, LM_IN_L2_Parts, LM_IN_L2_Playfield, LM_IN_L2_underPF, LM_IN_L20_BC2, LM_IN_L20_BS2, LM_IN_L20_EIprim, LM_IN_L20_Layer1, LM_IN_L20_Layer2, LM_IN_L20_Parts, LM_IN_L20_Playfield, LM_IN_L20_cabRails, LM_IN_L20_sw27, LM_IN_L20_sw41, LM_IN_L21_BC2, LM_IN_L21_BS1, LM_IN_L21_BS2, LM_IN_L21_BS3, LM_IN_L21_EIprim, LM_IN_L21_Layer1, LM_IN_L21_Layer2, LM_IN_L21_Parts, LM_IN_L21_Playfield, LM_IN_L21_cabRails, LM_IN_L21_underPF, LM_IN_L22_BS2, LM_IN_L22_BS3, LM_IN_L22_EIprim, LM_IN_L22_Layer1, LM_IN_L22_Parts, LM_IN_L22_Playfield, LM_IN_L22_cabRails, LM_IN_L23_BC2, LM_IN_L23_BS1, LM_IN_L23_BS2, LM_IN_L23_EIprim, LM_IN_L23_Layer1, LM_IN_L23_Layer2, LM_IN_L23_Parts, LM_IN_L23_Playfield, LM_IN_L23_sw27, LM_IN_L23_sw41, LM_IN_L23_underPF, LM_IN_L24_BC1, LM_IN_L24_BC2, LM_IN_L24_BR1, LM_IN_L24_BS1, LM_IN_L24_EIprim, LM_IN_L24_Layer1, LM_IN_L24_Parts, LM_IN_L24_Playfield, LM_IN_L24_cabRails, LM_IN_L24_sw27, LM_IN_L24_sw28, LM_IN_L25_BC2, LM_IN_L25_BS3, LM_IN_L25_EIprim, LM_IN_L25_Layer1, LM_IN_L25_Parts, _
' LM_IN_L25_Playfield, LM_IN_L25_cabRails, LM_IN_L26_underPF, LM_IN_L27_Parts, LM_IN_L27_underPF, LM_IN_L28_Layer6, LM_IN_L28_Parts, LM_IN_L28_Playfield, LM_IN_L28_RRubber1, LM_IN_L28_RRubber2, LM_IN_L28_RRubber3, LM_IN_L28_RRubber4, LM_IN_L28_sw24, LM_IN_L28_underPF, LM_IN_L29_Playfield, LM_IN_L29_rpost, LM_IN_L29_underPF, LM_IN_L3_Parts, LM_IN_L3_Playfield, LM_IN_L3_underPF, LM_IN_L30_Layer6, LM_IN_L30_Parts, LM_IN_L30_Playfield, LM_IN_L30_RRubber1, LM_IN_L30_RRubber2, LM_IN_L30_RRubber3, LM_IN_L30_RRubber4, LM_IN_L30_rpost, LM_IN_L30_underPF, LM_IN_L31_Layer6, LM_IN_L31_Parts, LM_IN_L31_Playfield, LM_IN_L31_RRubber1, LM_IN_L31_RRubber2, LM_IN_L31_RRubber3, LM_IN_L31_RRubber4, LM_IN_L31_rpost, LM_IN_L31_underPF, LM_IN_L32_LFlipper, LM_IN_L32_Parts, LM_IN_L32_Playfield, LM_IN_L32_RFlipper, LM_IN_L32_underPF, LM_IN_L33_Playfield, LM_IN_L33_underPF, LM_IN_L34_Layer1, LM_IN_L34_Parts, LM_IN_L34_Playfield, LM_IN_L34_underPF, LM_IN_L35_Layer1, LM_IN_L35_Playfield, LM_IN_L35_underPF, LM_IN_L36_Layer1, _
' LM_IN_L36_Parts, LM_IN_L36_Playfield, LM_IN_L36_underPF, LM_IN_L37_Layer1, LM_IN_L37_Playfield, LM_IN_L37_lpost, LM_IN_L37_underPF, LM_IN_L38_LRubber1, LM_IN_L38_LRubber2, LM_IN_L38_LRubber3, LM_IN_L38_LRubber4, LM_IN_L38_Layer1, LM_IN_L38_Parts, LM_IN_L38_Playfield, LM_IN_L38_underPF, LM_IN_L39_Layer1, LM_IN_L39_Parts, LM_IN_L39_Playfield, LM_IN_L39_underPF, LM_IN_L4_Parts, LM_IN_L4_Playfield, LM_IN_L4_underPF, LM_IN_L40_Playfield, LM_IN_L40_underPF, LM_IN_L41_BC1, LM_IN_L41_BC2, LM_IN_L41_BR1, LM_IN_L41_BS1, LM_IN_L41_BS2, LM_IN_L41_BS3, LM_IN_L41_Layer1, LM_IN_L41_Layer2, LM_IN_L41_Parts, LM_IN_L41_Playfield, LM_IN_L41_sw27, LM_IN_L41_sw41, LM_IN_L41_underPF, LM_IN_L42_BC1, LM_IN_L42_BC2, LM_IN_L42_BC3, LM_IN_L42_BR1, LM_IN_L42_BR2, LM_IN_L42_BR3, LM_IN_L42_BS1, LM_IN_L42_BS3, LM_IN_L42_Layer1, LM_IN_L42_Parts, LM_IN_L42_Playfield, LM_IN_L42_underPF, LM_IN_L43_BC1, LM_IN_L43_BC2, LM_IN_L43_BR1, LM_IN_L43_BR2, LM_IN_L43_BS1, LM_IN_L43_BS2, LM_IN_L43_BS3, LM_IN_L43_EIprim, LM_IN_L43_Layer1, LM_IN_L43_Parts, _
' LM_IN_L43_Playfield, LM_IN_L43_sw27, LM_IN_L43_sw28, LM_IN_L43_sw29, LM_IN_L43_sw41, LM_IN_L43_underPF, LM_IN_L44_Playfield, LM_IN_L44_underPF, LM_IN_L45_Playfield, LM_IN_L45_underPF, LM_IN_L46_Parts, LM_IN_L46_Playfield, LM_IN_L46_underPF, LM_IN_L47_Parts, LM_IN_L47_Playfield, LM_IN_L47_underPF, LM_IN_L48_BC2, LM_IN_L48_Parts, LM_IN_L48_Playfield, LM_IN_L48_sw30, LM_IN_L48_underPF, LM_IN_L49a_EIprim, LM_IN_L49a_Layer1, LM_IN_L49a_Lockdownbar, LM_IN_L49a_Parts, LM_IN_L49a_Playfield, LM_IN_L49a_cabRails, LM_IN_L49a_sw44, LM_IN_L49a_underPF, LM_IN_L5_Parts, LM_IN_L5_Playfield, LM_IN_L5_underPF, LM_IN_L50_Parts, LM_IN_L50_underPF, LM_IN_L51_Playfield, LM_IN_L51_underPF, LM_IN_L52_Parts, LM_IN_L52_Playfield, LM_IN_L52_underPF, LM_IN_L53_Parts, LM_IN_L53_Playfield, LM_IN_L53_sw30, LM_IN_L53_underPF, LM_IN_L54_Layer1, LM_IN_L54_Parts, LM_IN_L54_Playfield, LM_IN_L54_sw30, LM_IN_L54_underPF, LM_IN_L55_Layer1, LM_IN_L55_Parts, LM_IN_L55_Playfield, LM_IN_L55_underPF, LM_IN_L56_Parts, LM_IN_L56_Playfield, _
' LM_IN_L56_sw30, LM_IN_L56_underPF, LM_IN_L57a_Layer1, LM_IN_L57a_Layer6, LM_IN_L57a_Parts, LM_IN_L57a_Playfield, LM_IN_L57a_cabRails, LM_IN_L57a_sw23, LM_IN_L57a_underPF, LM_IN_L6_Parts, LM_IN_L6_Playfield, LM_IN_L6_sw24, LM_IN_L6_underPF, LM_IN_L7_BC1, LM_IN_L7_BC2, LM_IN_L7_BR1, LM_IN_L7_BS1, LM_IN_L7_BS2, LM_IN_L7_EIprim, LM_IN_L7_Layer1, LM_IN_L7_Layer2, LM_IN_L7_Parts, LM_IN_L7_Playfield, LM_IN_L7_sw41, LM_IN_L7_underPF, LM_IN_L8_BC1, LM_IN_L8_BC2, LM_IN_L8_BC3, LM_IN_L8_BR1, LM_IN_L8_BR2, LM_IN_L8_BR3, LM_IN_L8_BS1, LM_IN_L8_BS2, LM_IN_L8_BS3, LM_IN_L8_Layer1, LM_IN_L8_Layer2, LM_IN_L8_Parts, LM_IN_L8_Playfield, LM_IN_L8_sw27, LM_IN_L8_sw28, LM_IN_L8_sw41, LM_IN_L8_underPF, LM_IN_L9_LEMK, LM_IN_L9_LSling1, LM_IN_L9_Parts, LM_IN_L9_Playfield, LM_IN_L9_underPF)
'Dim BG_All: BG_All=Array(BM_BC1, BM_BC2, BM_BC3, BM_BR1, BM_BR2, BM_BR3, BM_BS1, BM_BS2, BM_BS3, BM_EIprim, BM_Gate2, BM_LEMK, BM_LFlipper, BM_LRubber1, BM_LRubber2, BM_LRubber3, BM_LRubber4, BM_LSling1, BM_LSling2, BM_LUFlipper, BM_Layer1, BM_Layer2, BM_Layer3, BM_Layer4, BM_Layer5, BM_Layer6, BM_Layer7, BM_Lockdownbar, BM_Parts, BM_Playfield, BM_REMK, BM_RFlipper, BM_RRubber1, BM_RRubber2, BM_RRubber3, BM_RRubber4, BM_RSling1, BM_RSling2, BM_cabRails, BM_diverter, BM_gate1, BM_lpost, BM_rpost, BM_sw14, BM_sw15, BM_sw16, BM_sw17, BM_sw18, BM_sw19, BM_sw21, BM_sw22, BM_sw23, BM_sw24, BM_sw27, BM_sw28, BM_sw29, BM_sw30, BM_sw31, BM_sw32, BM_sw33, BM_sw34, BM_sw35, BM_sw36, BM_sw41, BM_sw41a, BM_sw43, BM_sw44, BM_sw45, BM_sw46, BM_sw50, BM_underPF, LM_FL_f1_BC1, LM_FL_f1_BC2, LM_FL_f1_BS1, LM_FL_f1_EIprim, LM_FL_f1_LEMK, LM_FL_f1_LRubber1, LM_FL_f1_LRubber2, LM_FL_f1_LRubber3, LM_FL_f1_LRubber4, LM_FL_f1_LUFlipper, LM_FL_f1_Layer1, LM_FL_f1_Parts, LM_FL_f1_Playfield, LM_FL_f1_RSling1, LM_FL_f1_RSling2, _
' LM_FL_f1_cabRails, LM_FL_f1_lpost, LM_FL_f1_rpost, LM_FL_f1_sw19, LM_FL_f1_sw24, LM_FL_f1_sw30, LM_FL_f1_sw45, LM_FL_f1_sw46, LM_FL_f1_underPF, LM_FL_f10_BC1, LM_FL_f10_BC2, LM_FL_f10_BC3, LM_FL_f10_BR1, LM_FL_f10_BR3, LM_FL_f10_BS1, LM_FL_f10_BS2, LM_FL_f10_BS3, LM_FL_f10_EIprim, LM_FL_f10_Layer1, LM_FL_f10_Layer2, LM_FL_f10_Layer3, LM_FL_f10_Layer4, LM_FL_f10_Parts, LM_FL_f10_Playfield, LM_FL_f10_cabRails, LM_FL_f10_sw33, LM_FL_f10_sw34, LM_FL_f10_sw35, LM_FL_f10_sw41, LM_FL_f10_sw45, LM_FL_f114_BC3, LM_FL_f114_BR3, LM_FL_f114_BS2, LM_FL_f114_BS3, LM_FL_f114_EIprim, LM_FL_f114_LRubber1, LM_FL_f114_LRubber2, LM_FL_f114_LRubber3, LM_FL_f114_LRubber4, LM_FL_f114_LUFlipper, LM_FL_f114_Layer1, LM_FL_f114_Layer6, LM_FL_f114_Parts, LM_FL_f114_Playfield, LM_FL_f114_RRubber1, LM_FL_f114_RRubber2, LM_FL_f114_RRubber3, LM_FL_f114_RRubber4, LM_FL_f114_cabRails, LM_FL_f114_gate1, LM_FL_f114_rpost, LM_FL_f114_sw19, LM_FL_f114_sw21, LM_FL_f114_sw22, LM_FL_f114_sw28, LM_FL_f114_sw29, LM_FL_f114_sw30, LM_FL_f114_underPF, _
' LM_FL_f114a_EIprim, LM_FL_f114a_LRubber1, LM_FL_f114a_LRubber2, LM_FL_f114a_LRubber3, LM_FL_f114a_LRubber4, LM_FL_f114a_LUFlipper, LM_FL_f114a_Layer1, LM_FL_f114a_Layer6, LM_FL_f114a_Parts, LM_FL_f114a_Playfield, LM_FL_f114a_RRubber3, LM_FL_f114a_cabRails, LM_FL_f114a_gate1, LM_FL_f114a_lpost, LM_FL_f114a_sw19, LM_FL_f114a_sw21, LM_FL_f114a_sw22, LM_FL_f114a_sw24, LM_FL_f114a_sw30, LM_FL_f114a_underPF, LM_FL_f2_BC1, LM_FL_f2_BC2, LM_FL_f2_BR1, LM_FL_f2_BR2, LM_FL_f2_BR3, LM_FL_f2_BS1, LM_FL_f2_BS2, LM_FL_f2_BS3, LM_FL_f2_EIprim, LM_FL_f2_LUFlipper, LM_FL_f2_Layer1, LM_FL_f2_Layer2, LM_FL_f2_Layer6, LM_FL_f2_Parts, LM_FL_f2_Playfield, LM_FL_f2_cabRails, LM_FL_f2_gate1, LM_FL_f2_sw27, LM_FL_f2_sw29, LM_FL_f2_sw30, LM_FL_f2_sw41, LM_FL_f2_sw45, LM_FL_f2_sw46, LM_FL_f2_underPF, LM_FL_f3_BC1, LM_FL_f3_BC2, LM_FL_f3_BC3, LM_FL_f3_BR1, LM_FL_f3_BR2, LM_FL_f3_BS1, LM_FL_f3_BS2, LM_FL_f3_BS3, LM_FL_f3_EIprim, LM_FL_f3_Layer1, LM_FL_f3_Layer2, LM_FL_f3_Parts, LM_FL_f3_Playfield, LM_FL_f3_cabRails, LM_FL_f3_sw23, _
' LM_FL_f3_sw27, LM_FL_f3_sw28, LM_FL_f3_sw29, LM_FL_f3_sw33, LM_FL_f3_sw34, LM_FL_f3_sw35, LM_FL_f3_sw41, LM_FL_f3_sw41a, LM_FL_f3_sw46, LM_FL_f3_underPF, LM_FL_f4_EIprim, LM_FL_f4_Gate2, LM_FL_f4_Layer1, LM_FL_f4_Layer2, LM_FL_f4_Layer3, LM_FL_f4_Layer5, LM_FL_f4_Parts, LM_FL_f4_Playfield, LM_FL_f4_cabRails, LM_FL_f4_diverter, LM_FL_f4_sw30, LM_FL_f4_sw31, LM_FL_f4_sw32, LM_FL_f4_sw41, LM_FL_f4_sw44, LM_FL_f4_underPF, LM_FL_f5_EIprim, LM_FL_f5_Layer1, LM_FL_f5_Layer2, LM_FL_f5_Layer3, LM_FL_f5_Parts, LM_FL_f5_Playfield, LM_FL_f5_cabRails, LM_FL_f5_sw41, LM_FL_f6_BC2, LM_FL_f6_BC3, LM_FL_f6_BR2, LM_FL_f6_BR3, LM_FL_f6_EIprim, LM_FL_f6_LSling1, LM_FL_f6_Layer1, LM_FL_f6_Layer3, LM_FL_f6_Layer6, LM_FL_f6_Layer7, LM_FL_f6_Parts, LM_FL_f6_Playfield, LM_FL_f6_RRubber1, LM_FL_f6_RRubber2, LM_FL_f6_RRubber3, LM_FL_f6_RRubber4, LM_FL_f6_cabRails, LM_FL_f6_lpost, LM_FL_f6_rpost, LM_FL_f6_sw19, LM_FL_f6_sw21, LM_FL_f6_sw22, LM_FL_f6_sw23, LM_FL_f6_sw24, LM_FL_f6_sw30, LM_FL_f6_underPF, LM_FL_f7_BS2, LM_FL_f7_EIprim, _
' LM_FL_f7_Layer1, LM_FL_f7_Layer2, LM_FL_f7_Layer3, LM_FL_f7_Layer6, LM_FL_f7_Parts, LM_FL_f7_Playfield, LM_FL_f7_cabRails, LM_FL_f7_diverter, LM_FL_f7_gate1, LM_FL_f7_sw23, LM_FL_f7_sw30, LM_FL_f7_sw31, LM_FL_f7_sw43, LM_FL_f7_sw44, LM_FL_f7_underPF, LM_FL_f8_Layer1, LM_FL_f8_Layer2, LM_FL_f8_Layer3, LM_FL_f8_Parts, LM_FL_f8_Playfield, LM_FL_f8_sw30, LM_FL_f9_BC2, LM_FL_f9_EIprim, LM_FL_f9_Layer1, LM_FL_f9_Layer2, LM_FL_f9_Layer3, LM_FL_f9_Layer5, LM_FL_f9_Parts, LM_FL_f9_Playfield, LM_FL_f9_cabRails, LM_FL_f9_diverter, LM_FL_f9_sw30, LM_FL_f9_sw31, LM_FL_f9_sw44, LM_GI_BC1, LM_GI_BC2, LM_GI_BC3, LM_GI_BR1, LM_GI_BR2, LM_GI_BR3, LM_GI_BS1, LM_GI_BS2, LM_GI_BS3, LM_GI_EIprim, LM_GI_LEMK, LM_GI_LRubber1, LM_GI_LRubber2, LM_GI_LRubber3, LM_GI_LRubber4, LM_GI_LSling1, LM_GI_LSling2, LM_GI_LUFlipper, LM_GI_Layer1, LM_GI_Layer2, LM_GI_Layer3, LM_GI_Layer4, LM_GI_Layer6, LM_GI_Layer7, LM_GI_Lockdownbar, LM_GI_Parts, LM_GI_Playfield, LM_GI_REMK, LM_GI_RFlipper, LM_GI_RRubber1, LM_GI_RRubber2, LM_GI_RRubber3, _
' LM_GI_RRubber4, LM_GI_cabRails, LM_GI_diverter, LM_GI_gate1, LM_GI_lpost, LM_GI_rpost, LM_GI_sw19, LM_GI_sw21, LM_GI_sw22, LM_GI_sw23, LM_GI_sw24, LM_GI_sw27, LM_GI_sw28, LM_GI_sw29, LM_GI_sw30, LM_GI_sw31, LM_GI_sw33, LM_GI_sw35, LM_GI_sw41, LM_GI_sw43, LM_GI_sw44, LM_GI_sw45, LM_GI_sw46, LM_GI_underPF, LM_GIS_GI_1_LEMK, LM_GIS_GI_1_LFlipper, LM_GIS_GI_1_Layer1, LM_GIS_GI_1_Parts, LM_GIS_GI_1_Playfield, LM_GIS_GI_1_RFlipper, LM_GIS_GI_1_underPF, LM_GIS_GI_10_LRubber1, LM_GIS_GI_10_LRubber2, LM_GIS_GI_10_LRubber3, LM_GIS_GI_10_LRubber4, LM_GIS_GI_10_LUFlipper, LM_GIS_GI_10_Layer1, LM_GIS_GI_10_Parts, LM_GIS_GI_10_Playfield, LM_GIS_GI_10_cabRails, LM_GIS_GI_10_lpost, LM_GIS_GI_10_sw19, LM_GIS_GI_10_sw46, LM_GIS_GI_10_underPF, LM_GIS_GI_13_BC2, LM_GIS_GI_13_BR2, LM_GIS_GI_13_Layer6, LM_GIS_GI_13_Parts, LM_GIS_GI_13_Playfield, LM_GIS_GI_13_REMK, LM_GIS_GI_13_RRubber1, LM_GIS_GI_13_RRubber2, LM_GIS_GI_13_RRubber3, LM_GIS_GI_13_RRubber4, LM_GIS_GI_13_cabRails, LM_GIS_GI_13_rpost, LM_GIS_GI_13_sw24, _
' LM_GIS_GI_13_underPF, LM_GIS_GI_15_LRubber1, LM_GIS_GI_15_LRubber2, LM_GIS_GI_15_LRubber3, LM_GIS_GI_15_LRubber4, LM_GIS_GI_15_LUFlipper, LM_GIS_GI_15_Layer1, LM_GIS_GI_15_Parts, LM_GIS_GI_15_Playfield, LM_GIS_GI_15_lpost, LM_GIS_GI_15_sw46, LM_GIS_GI_15_underPF, LM_GIS_GI_16_LFlipper, LM_GIS_GI_16_Parts, LM_GIS_GI_16_Playfield, LM_GIS_GI_16_REMK, LM_GIS_GI_16_RFlipper, LM_GIS_GI_16_RRubber1, LM_GIS_GI_16_RRubber2, LM_GIS_GI_16_RRubber3, LM_GIS_GI_16_RRubber4, LM_GIS_GI_16_RSling1, LM_GIS_GI_16_RSling2, LM_GIS_GI_16_cabRails, LM_GIS_GI_16_sw14, LM_GIS_GI_16_sw18, LM_GIS_GI_16_sw24, LM_GIS_GI_16_underPF, LM_GIS_GI_17_LEMK, LM_GIS_GI_17_LFlipper, LM_GIS_GI_17_LRubber1, LM_GIS_GI_17_LRubber2, LM_GIS_GI_17_LRubber3, LM_GIS_GI_17_LRubber4, LM_GIS_GI_17_LSling1, LM_GIS_GI_17_LSling2, LM_GIS_GI_17_Layer1, LM_GIS_GI_17_Parts, LM_GIS_GI_17_Playfield, LM_GIS_GI_17_RFlipper, LM_GIS_GI_17_cabRails, LM_GIS_GI_17_sw14, LM_GIS_GI_17_sw17, LM_GIS_GI_17_sw18, LM_GIS_GI_17_underPF, LM_GIS_GI_2_LFlipper, LM_GIS_GI_2_Layer1, _
' LM_GIS_GI_2_Parts, LM_GIS_GI_2_Playfield, LM_GIS_GI_2_REMK, LM_GIS_GI_2_RFlipper, LM_GIS_GI_2_RSling1, LM_GIS_GI_2_sw17, LM_GIS_GI_20_LEMK, LM_GIS_GI_20_LFlipper, LM_GIS_GI_20_LRubber1, LM_GIS_GI_20_LRubber2, LM_GIS_GI_20_LRubber3, LM_GIS_GI_20_LRubber4, LM_GIS_GI_20_LSling1, LM_GIS_GI_20_LSling2, LM_GIS_GI_20_Layer1, LM_GIS_GI_20_Parts, LM_GIS_GI_20_Playfield, LM_GIS_GI_20_RFlipper, LM_GIS_GI_20_cabRails, LM_GIS_GI_20_lpost, LM_GIS_GI_20_rpost, LM_GIS_GI_20_sw17, LM_GIS_GI_20_sw18, LM_GIS_GI_20_sw19, LM_GIS_GI_20_underPF, LM_GIS_GI_29_LFlipper, LM_GIS_GI_29_Layer1, LM_GIS_GI_29_Parts, LM_GIS_GI_29_Playfield, LM_GIS_GI_29_REMK, LM_GIS_GI_29_RRubber1, LM_GIS_GI_29_RRubber2, LM_GIS_GI_29_RRubber3, LM_GIS_GI_29_RRubber4, LM_GIS_GI_29_RSling1, LM_IN_L1_Playfield, LM_IN_L1_underPF, LM_IN_L10_Playfield, LM_IN_L10_underPF, LM_IN_L11_Playfield, LM_IN_L11_underPF, LM_IN_L12_Playfield, LM_IN_L12_underPF, LM_IN_L13_Parts, LM_IN_L13_Playfield, LM_IN_L13_REMK, LM_IN_L13_underPF, LM_IN_L14_Playfield, LM_IN_L14_underPF, _
' LM_IN_L15_Layer1, LM_IN_L15_Parts, LM_IN_L15_Playfield, LM_IN_L15_underPF, LM_IN_L16_LRubber1, LM_IN_L16_LRubber2, LM_IN_L16_LRubber3, LM_IN_L16_LRubber4, LM_IN_L16_Layer1, LM_IN_L16_Parts, LM_IN_L16_Playfield, LM_IN_L16_lpost, LM_IN_L16_underPF, LM_IN_L17_BC1, LM_IN_L17_BC2, LM_IN_L17_BS2, LM_IN_L17_EIprim, LM_IN_L17_Layer1, LM_IN_L17_Layer2, LM_IN_L17_Parts, LM_IN_L17_Playfield, LM_IN_L17_cabRails, LM_IN_L17_sw27, LM_IN_L17_sw28, LM_IN_L18_BS2, LM_IN_L18_BS3, LM_IN_L18_EIprim, LM_IN_L18_Layer1, LM_IN_L18_Layer2, LM_IN_L18_Parts, LM_IN_L18_Playfield, LM_IN_L18_cabRails, LM_IN_L18_sw28, LM_IN_L19_BS2, LM_IN_L19_BS3, LM_IN_L19_EIprim, LM_IN_L19_Layer1, LM_IN_L19_Parts, LM_IN_L19_Playfield, LM_IN_L19_cabRails, LM_IN_L2_Parts, LM_IN_L2_Playfield, LM_IN_L2_underPF, LM_IN_L20_BC2, LM_IN_L20_BS2, LM_IN_L20_EIprim, LM_IN_L20_Layer1, LM_IN_L20_Layer2, LM_IN_L20_Parts, LM_IN_L20_Playfield, LM_IN_L20_cabRails, LM_IN_L20_sw27, LM_IN_L20_sw41, LM_IN_L21_BC2, LM_IN_L21_BS1, LM_IN_L21_BS2, LM_IN_L21_BS3, LM_IN_L21_EIprim, _
' LM_IN_L21_Layer1, LM_IN_L21_Layer2, LM_IN_L21_Parts, LM_IN_L21_Playfield, LM_IN_L21_cabRails, LM_IN_L21_underPF, LM_IN_L22_BS2, LM_IN_L22_BS3, LM_IN_L22_EIprim, LM_IN_L22_Layer1, LM_IN_L22_Parts, LM_IN_L22_Playfield, LM_IN_L22_cabRails, LM_IN_L23_BC2, LM_IN_L23_BS1, LM_IN_L23_BS2, LM_IN_L23_EIprim, LM_IN_L23_Layer1, LM_IN_L23_Layer2, LM_IN_L23_Parts, LM_IN_L23_Playfield, LM_IN_L23_sw27, LM_IN_L23_sw41, LM_IN_L23_underPF, LM_IN_L24_BC1, LM_IN_L24_BC2, LM_IN_L24_BR1, LM_IN_L24_BS1, LM_IN_L24_EIprim, LM_IN_L24_Layer1, LM_IN_L24_Parts, LM_IN_L24_Playfield, LM_IN_L24_cabRails, LM_IN_L24_sw27, LM_IN_L24_sw28, LM_IN_L25_BC2, LM_IN_L25_BS3, LM_IN_L25_EIprim, LM_IN_L25_Layer1, LM_IN_L25_Parts, LM_IN_L25_Playfield, LM_IN_L25_cabRails, LM_IN_L26_underPF, LM_IN_L27_Parts, LM_IN_L27_underPF, LM_IN_L28_Layer6, LM_IN_L28_Parts, LM_IN_L28_Playfield, LM_IN_L28_RRubber1, LM_IN_L28_RRubber2, LM_IN_L28_RRubber3, LM_IN_L28_RRubber4, LM_IN_L28_sw24, LM_IN_L28_underPF, LM_IN_L29_Playfield, LM_IN_L29_rpost, LM_IN_L29_underPF, _
' LM_IN_L3_Parts, LM_IN_L3_Playfield, LM_IN_L3_underPF, LM_IN_L30_Layer6, LM_IN_L30_Parts, LM_IN_L30_Playfield, LM_IN_L30_RRubber1, LM_IN_L30_RRubber2, LM_IN_L30_RRubber3, LM_IN_L30_RRubber4, LM_IN_L30_rpost, LM_IN_L30_underPF, LM_IN_L31_Layer6, LM_IN_L31_Parts, LM_IN_L31_Playfield, LM_IN_L31_RRubber1, LM_IN_L31_RRubber2, LM_IN_L31_RRubber3, LM_IN_L31_RRubber4, LM_IN_L31_rpost, LM_IN_L31_underPF, LM_IN_L32_LFlipper, LM_IN_L32_Parts, LM_IN_L32_Playfield, LM_IN_L32_RFlipper, LM_IN_L32_underPF, LM_IN_L33_Playfield, LM_IN_L33_underPF, LM_IN_L34_Layer1, LM_IN_L34_Parts, LM_IN_L34_Playfield, LM_IN_L34_underPF, LM_IN_L35_Layer1, LM_IN_L35_Playfield, LM_IN_L35_underPF, LM_IN_L36_Layer1, LM_IN_L36_Parts, LM_IN_L36_Playfield, LM_IN_L36_underPF, LM_IN_L37_Layer1, LM_IN_L37_Playfield, LM_IN_L37_lpost, LM_IN_L37_underPF, LM_IN_L38_LRubber1, LM_IN_L38_LRubber2, LM_IN_L38_LRubber3, LM_IN_L38_LRubber4, LM_IN_L38_Layer1, LM_IN_L38_Parts, LM_IN_L38_Playfield, LM_IN_L38_underPF, LM_IN_L39_Layer1, LM_IN_L39_Parts, _
' LM_IN_L39_Playfield, LM_IN_L39_underPF, LM_IN_L4_Parts, LM_IN_L4_Playfield, LM_IN_L4_underPF, LM_IN_L40_Playfield, LM_IN_L40_underPF, LM_IN_L41_BC1, LM_IN_L41_BC2, LM_IN_L41_BR1, LM_IN_L41_BS1, LM_IN_L41_BS2, LM_IN_L41_BS3, LM_IN_L41_Layer1, LM_IN_L41_Layer2, LM_IN_L41_Parts, LM_IN_L41_Playfield, LM_IN_L41_sw27, LM_IN_L41_sw41, LM_IN_L41_underPF, LM_IN_L42_BC1, LM_IN_L42_BC2, LM_IN_L42_BC3, LM_IN_L42_BR1, LM_IN_L42_BR2, LM_IN_L42_BR3, LM_IN_L42_BS1, LM_IN_L42_BS3, LM_IN_L42_Layer1, LM_IN_L42_Parts, LM_IN_L42_Playfield, LM_IN_L42_underPF, LM_IN_L43_BC1, LM_IN_L43_BC2, LM_IN_L43_BR1, LM_IN_L43_BR2, LM_IN_L43_BS1, LM_IN_L43_BS2, LM_IN_L43_BS3, LM_IN_L43_EIprim, LM_IN_L43_Layer1, LM_IN_L43_Parts, LM_IN_L43_Playfield, LM_IN_L43_sw27, LM_IN_L43_sw28, LM_IN_L43_sw29, LM_IN_L43_sw41, LM_IN_L43_underPF, LM_IN_L44_Playfield, LM_IN_L44_underPF, LM_IN_L45_Playfield, LM_IN_L45_underPF, LM_IN_L46_Parts, LM_IN_L46_Playfield, LM_IN_L46_underPF, LM_IN_L47_Parts, LM_IN_L47_Playfield, LM_IN_L47_underPF, LM_IN_L48_BC2, _
' LM_IN_L48_Parts, LM_IN_L48_Playfield, LM_IN_L48_sw30, LM_IN_L48_underPF, LM_IN_L49a_EIprim, LM_IN_L49a_Layer1, LM_IN_L49a_Lockdownbar, LM_IN_L49a_Parts, LM_IN_L49a_Playfield, LM_IN_L49a_cabRails, LM_IN_L49a_sw44, LM_IN_L49a_underPF, LM_IN_L5_Parts, LM_IN_L5_Playfield, LM_IN_L5_underPF, LM_IN_L50_Parts, LM_IN_L50_underPF, LM_IN_L51_Playfield, LM_IN_L51_underPF, LM_IN_L52_Parts, LM_IN_L52_Playfield, LM_IN_L52_underPF, LM_IN_L53_Parts, LM_IN_L53_Playfield, LM_IN_L53_sw30, LM_IN_L53_underPF, LM_IN_L54_Layer1, LM_IN_L54_Parts, LM_IN_L54_Playfield, LM_IN_L54_sw30, LM_IN_L54_underPF, LM_IN_L55_Layer1, LM_IN_L55_Parts, LM_IN_L55_Playfield, LM_IN_L55_underPF, LM_IN_L56_Parts, LM_IN_L56_Playfield, LM_IN_L56_sw30, LM_IN_L56_underPF, LM_IN_L57a_Layer1, LM_IN_L57a_Layer6, LM_IN_L57a_Parts, LM_IN_L57a_Playfield, LM_IN_L57a_cabRails, LM_IN_L57a_sw23, LM_IN_L57a_underPF, LM_IN_L6_Parts, LM_IN_L6_Playfield, LM_IN_L6_sw24, LM_IN_L6_underPF, LM_IN_L7_BC1, LM_IN_L7_BC2, LM_IN_L7_BR1, LM_IN_L7_BS1, LM_IN_L7_BS2, LM_IN_L7_EIprim, _
' LM_IN_L7_Layer1, LM_IN_L7_Layer2, LM_IN_L7_Parts, LM_IN_L7_Playfield, LM_IN_L7_sw41, LM_IN_L7_underPF, LM_IN_L8_BC1, LM_IN_L8_BC2, LM_IN_L8_BC3, LM_IN_L8_BR1, LM_IN_L8_BR2, LM_IN_L8_BR3, LM_IN_L8_BS1, LM_IN_L8_BS2, LM_IN_L8_BS3, LM_IN_L8_Layer1, LM_IN_L8_Layer2, LM_IN_L8_Parts, LM_IN_L8_Playfield, LM_IN_L8_sw27, LM_IN_L8_sw28, LM_IN_L8_sw41, LM_IN_L8_underPF, LM_IN_L9_LEMK, LM_IN_L9_LSling1, LM_IN_L9_Parts, LM_IN_L9_Playfield, LM_IN_L9_underPF)
' VLM  Arrays - End



'*******************************************
' ZLOA: Load Stuff
'*******************************************

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="esha_la3"
'Const cGameName="esha_l4c" 'competition rom
Const UseSolenoids=2, UseLamps=1, UseGI=0, SSolenoidOn="SolOn", SSolenoidOff="SolOff", SCoin=""
Const UseVPMDMD = true
Const UseVPMModSol = 2

LoadVPM "03060000", "S11.VBS", 3.22

'NoUpperLeftFlipper
'NoUpperRightFlipper


'*******************************************
' ZTIM: Main Timers
'*******************************************

Dim FrameTime, LastGameTime
LastGameTime = 0

Sub FrameTimer_Timer() 'The frame timer interval should be -1, so executes at the display frame rate
  FrameTime = GameTime - LastGameTime
  LastGameTime = GameTime

  BSUpdate
  AnimateBumperSkirts
  RollingUpdate   'update rolling sounds
  DoSTAnim    'handle stand up target animations
  UpdateStandupTargets
  DoDTAnim    'handle drop up target animations
  UpdateDropTargets
  UpdateVRBackglass
' UpdateBallBrightness
End Sub


CorTimer.Interval = 10
Sub CorTimer_Timer(): Cor.Update: End Sub



'*******************************************
' ZINI: Table Initialization and Exiting
'*******************************************

Sub Table1_Init

  ' table initialization
  vpmInit Me
  On Error Resume Next
    With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine = "EARTHSHAKER - Williams 1989"&chr(13)&"VPW"
    .HandleMechanics=0
    .HandleKeyboard=0
    .ShowDMDOnly=1
    .ShowFrame=0
    .ShowTitle=0
        .hidden = 1
        .Games(cGameName).Settings.Value("sound")=1
    '.PuPHide = 1
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With
     On Error Goto 0

  vpmMapLights AllLamps ' Map all lamps to the corresponding ROM output using the value of TimerInterval of each light object

  ' nudging
  vpmNudge.TiltSwitch  = 9
  vpmNudge.Sensitivity = 2
  vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingshot)

    ' physical trough
  set esball1 = sw11.CreateSizedballWithMass(Ballsize / 2,Ballmass)
  set esball2 = sw12.CreateSizedballWithMass(Ballsize / 2,Ballmass)
  set esball3 = sw13.CreateSizedballWithMass(Ballsize / 2,Ballmass)
  gBOT = Array(esball1,esball2,esball3)

  Controller.Switch(11) = 1
  Controller.Switch(12) = 1
  Controller.Switch(13) = 1

   ' captive ball
  CapKicker.CreateBall
  CapKicker.Kick 0,5
  CapKicker.Enabled = false

  ' GI
  PFGI False: PFGI2 False

  ' misc
  sw39wall.collidable=false
  Pdiverter2.Collidable=0

  ' VR
    SetVRBackglass

  'Initialize slings
  RStep = 0:RightSlingShot.Timerenabled=True
  LStep = 0:LeftSlingShot.Timerenabled=True

  GateBlocker.IsDropped = true

End Sub


Sub Table1_Exit
  Controller.Pause = False
  Controller.Stop
End Sub



'*******************************************
'  ZOPT: User Options
'*******************************************

Dim RefractOpt : RefractOpt = 1       '0 - No Refraction (best performance), 1 - Sharp Refractions (improved performance), 2 - Rough Refractions (best visual)
Dim EIMod : EIMod = 1           '0 - Static Earthquake Institute, 1 - Moving Earthquake Institute
Dim OutpostMod : OutpostMod = 1           '0 - Normal, 1 - Hard
Dim ShakerSSF: ShakerSSF = 1        '0 - Disabled, 1 - Enabled
Dim ShakerIntensity: ShakerIntensity = 1  '0 - Low, 1- - Normal, 2 - High
Dim MechVolume : MechVolume = 0.8           'Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Dim BallRollVolume : BallRollVolume = 0.5   'Level of ball rolling volume. Value between 0 and 1
Dim RampRollVolume : RampRollVolume = 0.5   'Level of ramp rolling volume. Value between 0 and 1
Dim VRRoomChoice: VRRoomChoice = 1      '1 - Deluxe Room (Basti), 2 - Minimal Room
Dim LightLevel : LightLevel = 0.5     ' Level of room lighting (0 to 1), where 0 is dark and 100 is brightest


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
  Dim BP

  RefractOpt = Table1.Option("Refraction Setting", 0, 2, 1, 1, 0, Array("No Refraction (best performance)", "Sharp Refractions (improved performance)", "Rough Refractions (best visuals)"))
  SetRefractionProbes RefractOpt

  EIMod = Table1.Option("Earthquake Institute Mod", 0, 1, 1, 1, 0, Array("Building is Static", "Building Moves Up/Down"))
  If EIMod = 1 Then
    For each BP in BP_Layer5: BP.visible = 0: Next
  Else
    InstZ = InstZMax
    For Each BP In BP_EIprim: BP.TransZ = InstZ: next
    For each BP in BP_Layer5: BP.visible = 1: Next
  End If

    ' Outpost Difficulty
    OutpostMod = Table1.Option("Outpost Difficulty", 0, 1, 1, 0, 0, Array("Normal", "Hard"))
  If OutpostMod = 0 Then   ' Normal
    Post_l_norm.collidable = True: rubberwall_l_norm.collidable = True
    Post_r_norm.collidable = True: rubberwall_r_norm.collidable = True
    Post_l_hard.collidable = False: rubberwall_l_hard.collidable = False
    Post_r_hard.collidable = False: rubberwall_r_hard.collidable = False

    For each BP in BP_lpost: BP.transy = 0: Next
    For each BP in BP_LRubber1: BP.visible = False: Next
    For each BP in BP_LRubber2: BP.visible = False: Next
    For each BP in BP_LRubber3: BP.visible = True: Next
    For each BP in BP_LRubber4: BP.visible = True: Next

    For each BP in BP_rpost: BP.transy = 0: Next
    For each BP in BP_RRubber1: BP.visible = False: Next
    For each BP in BP_RRubber2: BP.visible = False: Next
    For each BP in BP_RRubber3: BP.visible = True: Next
    For each BP in BP_RRubber4: BP.visible = True: Next
  Else ' Hard
    Post_l_norm.collidable = False: rubberwall_l_norm.collidable = False
    Post_r_norm.collidable = False: rubberwall_r_norm.collidable = False
    Post_l_hard.collidable = True: rubberwall_l_hard.collidable = True
    Post_r_hard.collidable = True: rubberwall_r_hard.collidable = True

    For each BP in BP_lpost: BP.transy = -11: Next
    For each BP in BP_LRubber1: BP.visible = True: Next
    For each BP in BP_LRubber2: BP.visible = True: Next
    For each BP in BP_LRubber3: BP.visible = False: Next
    For each BP in BP_LRubber4: BP.visible = False: Next

    For each BP in BP_rpost: BP.transy = -13: Next
    For each BP in BP_RRubber1: BP.visible = True: Next
    For each BP in BP_RRubber2: BP.visible = True: Next
    For each BP in BP_RRubber3: BP.visible = False: Next
    For each BP in BP_RRubber4: BP.visible = False: Next
  End If

    ' Shaker SSF
    ShakerSSF = Table1.Option("Shaker SSF", 0, 1, 1, 1, 0, Array("Disabled", "Enabled"))

    ' Shaker SSF Intensity
    ShakerIntensity = Table1.Option("Shaker SSF Intensity", 0, 2, 1, 1, 0, Array("Low", "Normal", "High"))


    ' Sound volumes
    MechVolume = Table1.Option("Mech Volume", 0, 1, 0.01, 0.8, 1)
    BallRollVolume = Table1.Option("Ball Roll Volume", 0, 1, 0.01, 0.5, 1)
  RampRollVolume = Table1.Option("Ramp Roll Volume", 0, 1, 0.01, 0.5, 1)

  ' VR Room Selection
  VRRoomChoice = Table1.Option("VR Room", 1, 3, 1, 1, 0, Array("Deluxe Room", "Minimal Room", "Ultra Minimal"))
  If RenderingMode = 2 Then VRRoom = VRRoomChoice Else VRRoom = 0
  'VRRoom = 1  'Uncomment this to test VR Room on desktop
  EnableVRRoom

  ' Room brightness
  LightLevel = NightDay/100
  'LightLevel = Table1.Option("Table Ambient Light Level", 0, 1, 0.01, .4, 1)
  SetRoomBrightness LightLevel

    If eventId = 3 And dspTriggered Then dspTriggered = False : DisableStaticPreRendering = False : End If
End Sub

Sub SetRefractionProbes(Opt)
  On Error Resume Next
    Select Case Opt
      Case 0:
        BM_BC1.RefractionProbe = "Bumper Caps"
        BM_BC2.RefractionProbe = "Bumper Caps"
        BM_BC3.RefractionProbe = "Bumper Caps"
        BM_Layer1.RefractionProbe = ""
        BM_Layer6.RefractionProbe = ""
      Case 1:
        BM_BC1.RefractionProbe = "Bumper Caps"
        BM_BC2.RefractionProbe = "Bumper Caps"
        BM_BC3.RefractionProbe = "Bumper Caps"
        BM_Layer1.RefractionProbe = "Ramps"
        BM_Layer6.RefractionProbe = "Ramps"
      Case 2:
        BM_BC1.RefractionProbe = "Bumper Caps Rough"
        BM_BC2.RefractionProbe = "Bumper Caps Rough"
        BM_BC3.RefractionProbe = "Bumper Caps Rough"
        BM_Layer1.RefractionProbe = "Ramps Rough"
        BM_Layer6.RefractionProbe = "Ramps Rough"
    End Select
  On Error Goto 0
End Sub



'****************************
'   ZBRI: Room Brightness
'****************************

' Update these arrays if you want to change more materials with room light level
Dim RoomBrightnessMtlArray: RoomBrightnessMtlArray = Array("VLM.Bake.Active","VLM.Bake.Solid","VLM.Bake.RedBumper","VLM.Bake.BlueBumper","VLM.Bake.YellowBumper","Plastic with an image1")


Sub SetRoomBrightness(lvl)
  If lvl > 1 Then lvl = 1
  If lvl < 0 Then lvl = 0
  'lvl = lvl^2

  ' Lighting level
  Dim v: v=(lvl * 240 + 15)/255

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





'*******************************************
'   ZKEY: Key Press Handling
'*******************************************

Sub Table1_KeyDown(ByVal KeyCode)

  If keycode = LeftFlipperKey Then
    Controller.Switch(58)  = 1
    FlipperActivate LeftFlipper, LFPress
    PinCab_Button_Left.X = PinCab_Button_Left.X + 8
  End If

  If keycode = RightFlipperKey Then
    Controller.Switch(57) = 1
    FlipperActivate RightFlipper, RFPress
    PinCab_Button_Right.X = PinCab_Button_Right.X - 8
  End If
'
' If keycode = KeyUpperLeft Then
'   Controller.Switch(60)  = 1
'   FlipperActivate LeftFlipper1, ULFPress
' End If


  If keycode = PlungerKey Then
    Plunger.Pullback
    SoundPlungerPull()
    TimerVRPlunger.Enabled = True
    TimerVRPlunger1.Enabled = False
    Pincab_Plunger.Y = 0
  End If

  If keycode = LeftTiltKey Then Nudge 90, 1:SoundNudgeLeft()
  If keycode = RightTiltKey Then Nudge 270, 1:SoundNudgeRight()
  If keycode = CenterTiltKey Then Nudge 0, 1:SoundNudgeCenter()

  If keycode = keyInsertCoin1 or keycode = keyInsertCoin2 or keycode = keyInsertCoin3 or keycode = keyInsertCoin4 Then
    SoundCoinIn
  End If

  If keycode = StartGameKey Then
        soundStartButton()
  End if

  If KeyDownHandler(keycode) Then Exit Sub
End Sub



Sub Table1_KeyUp(ByVal KeyCode)
  If keycode = LeftFlipperKey Then
    Controller.Switch(58)  = 0
    FlipperDeActivate LeftFlipper, LFPress
    PinCab_Button_Left.X = PinCab_Button_Left.X - 8
  End If

  If keycode = RightFlipperKey Then
    Controller.Switch(57) = 0
    FlipperDeActivate RightFlipper, RFPress
    PinCab_Button_Right.X = PinCab_Button_Right.X + 8
  End If

  If keycode = PlungerKey Then
    Plunger.Fire
    SoundPlungerReleaseBall()
    TimerVRPlunger.Enabled = False
        TimerVRPlunger1.Enabled = True
    Pincab_Plunger.Y = 0
  End If

  If KeyUpHandler(keycode) Then Exit Sub
End Sub






'*******************************************
' ZSOL: Solenoids
'*******************************************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"
SolCallback(sULFlipper) = "SolULFlipper"


SolCallback(01) = "SolOuthole"
SolCallback(02) = "ReleaseBall"
SolCallback(03) = "SolDropReset"
SolCallback(04) = "SolFault"
SolCallback(05) = "VUKLeft"
SolCallback(06) = "VUKBot"
SolCallback(07) = "SolKnocker"
SolCallback(09) = "InstituteDrop"
'SolCallBack(11) = ""
SolCallback(13) = "VUKTop"
SolCallback(22) = "ShakerMotor"


' Solenoid Lights
SolCallback(10) = "PFGI2"
SolModCallBack(12) = "Flash112" 'For the various backglass solenoid flashers
SolModCallBack(14) = "Flash114"
SolCallBack(15) = "PFGI"
SolModCallBack(16) = "Flash116"
SolModCallBack(25) = "Flash125"
SolModCallBack(26) = "Flash126" 'Also controls the backglass solenoid flashers
SolModCallBack(27) = "Flash127" 'Also controls the backglass solenoid flashers
SolModCallBack(28) = "Flash128" 'Also controls the backglass solenoid flashers
SolModCallBack(29) = "Flash129"
SolModCallBack(30) = "Flash130" 'Also controls the backglass solenoid flashers
SolModCallBack(31) = "Flash131" 'Also controls the backglass solenoid flashers
SolModCallBack(32) = "Flash132" 'Also controls the backglass solenoid flashers



' Knocker
Sub SolKnocker(Enabled)
  If Enabled Then
    KnockerSolenoid
  End If
End Sub


'*************************************************************
'   ZFLA: PWM Flasher Stuff
'*************************************************************

const DebugFlashers = false

' Modulated Calls

Sub Flash112(pwm)
  If DebugFlashers then debug.print "Flash112 "&pwm
  IF VRRoom <> 0 Then
    FlBG1.intensityscale = pwm
    FlBG2.intensityscale = pwm
    FlBG3.intensityscale = pwm
    FlBG4.intensityscale = pwm
    FlBG5.intensityscale = pwm
    FlBG6.intensityscale = pwm
    FlBG7.intensityscale = pwm
    FlBG8.intensityscale = pwm
    FlBG9.intensityscale = pwm
    FlBG10.intensityscale = pwm
    FlBG11.intensityscale = pwm
    FlBG12.intensityscale = pwm
   End If
End Sub


Sub Flash114(pwm)
  If DebugFlashers then debug.print "Flash114 "&pwm
  f114.state = pwm
  f114a.state = pwm
End Sub

Sub Flash116(pwm)  ' 16: On Ramp & J Bumper Flasher
  If DebugFlashers then debug.print "Flash116 "&pwm
  f10.state = pwm
End Sub


Sub Flash125(pwm)   ' 01C: Captive Ball Flasher
  If DebugFlashers then debug.print "Flash125 "&pwm
  f6.state = pwm
End Sub


Sub Flash126(pwm)   ' 02C: Center Ramp 1 & Building Flasher
  If DebugFlashers then debug.print "Flash126 "&pwm
  f4.state = pwm
  If VRRoom <> 0 and pwm > 0.5 Then FlBG26.intensityscale = pwm 'Backglass Flasher
End Sub


Sub Flash127(pwm)  ' 03C: Center Ramp 2 & Spinner Flasher
  If DebugFlashers then debug.print "Flash127 "&pwm
  f3.state = pwm
  If VRRoom <> 0 and pwm > 0.5 Then FlBG27.intensityscale = pwm 'Backglass Flasher
End Sub


Sub Flash128(pwm)  ' 05A: Eject Hole Flasher
  If DebugFlashers then debug.print "Flash128 "&pwm
  f2.state = pwm
  If VRRoom <> 0 and pwm > 0.5 Then FlBG28.intensityscale = pwm 'Backglass Flasher
End Sub


Sub Flash129(pwm)  ' 05C: Center Ramp 4 Flasher
  If DebugFlashers then debug.print "Flash129 "&pwm
  f1.state = pwm
End Sub


Sub Flash130(pwm) ' 06C: Right Ramp 1 Flasher
  If DebugFlashers then debug.print "Flash130 "&pwm
  f7.state = pwm
  If VRRoom <> 0 and pwm > 0.5 Then FlBG30.intensityscale = pwm 'Backglass Flasher
End Sub


Sub Flash131(pwm)  ' 07C: Right Ramp 2 Flasher
  If DebugFlashers then debug.print "Flash131 "&pwm
  f8.state = pwm
  If VRRoom <> 0 and pwm > 0.5 Then FlBG31.intensityscale = pwm 'Backglass Flasher

End Sub

Sub Flash132(pwm)
  If DebugFlashers then debug.print "Flash132 "&pwm
  f5.state = pwm
  f9.state = pwm
  If VRRoom <> 0 and pwm > 0.5 Then FlBG32.intensityscale = pwm 'Backglass Flasher
End Sub





' ******************************************************************************************
'      LAMP CALLBACK for the 6 backglass flasher lamps (not the solenoid conrolled ones
' ******************************************************************************************

Set LampCallback = GetRef("UpdateMultipleLamps")

Sub UpdateMultipleLamps()
  IF VRRoom <> 0 Then
  If Controller.Lamp(1) = 0 Then: FlBGL1.visible=0: else: FlBGL1.visible=1
  If Controller.Lamp(2) = 0 Then: FlBGL2.visible=0: else: FlBGL2.visible=1
  If Controller.Lamp(3) = 0 Then: FlBGL3.visible=0: else: FlBGL3.visible=1
  If Controller.Lamp(4) = 0 Then: FlBGL4.visible=0: else: FlBGL4.visible=1
  If Controller.Lamp(5) = 0 Then: FlBGL5.visible=0: else: FlBGL5.visible=1
  If Controller.Lamp(6) = 0 Then: FlBGL6.visible=0: else: FlBGL6.visible=1
  End If
End Sub




'*******************************************
' ZFLP: Flippers
'*******************************************

Const ReflipAngle = 20

Sub SolLFlipper(Enabled)
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

Sub SolULFlipper(Enabled)
    If Enabled Then
    LeftFlipper1.RotateToEnd
    If leftflipper.currentangle < leftflipper.endangle + ReflipAngle Then
      RandomSoundReflipUpLeft LeftFlipper
    Else
      SoundFlipperUpAttackLeft LeftFlipper
      RandomSoundFlipperUpLeft LeftFlipper
    End If
    Else
    LeftFlipper1.RotateToStart
    If LeftFlipper.currentangle < LeftFlipper.startAngle - 5 Then
      RandomSoundFlipperDownLeft LeftFlipper
    End If
    FlipperLeftHitParm = FlipperUpSoundLevel
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
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


'*******************************************
' ZTOY: Solenoid Controlled toys
'*******************************************

Const ShakerSoundLevel = 1
ShakeTimer.Interval = 500

Sub ShakerMotor(enabled)
  If Enabled Then
    If ShakeTimer.Enabled = false Then SoundShaker True
    ShakeTimer.Enabled = false
    ShakeTimer.Enabled = true
  End If
End Sub


Sub ShakeTimer_Timer
  SoundShaker False
  ShakeTimer.Enabled = false
End Sub


Sub SoundShaker(Enabled)
  If ShakerSSF = 0 Then Exit Sub
  'debug.print "SoundShaker  "&Enabled
  If Enabled Then
    Select Case ShakerIntensity
      Case 0
        PlaySoundAtLevelExistingStaticLoop ("SY_TNA_REV02_Shaker_Level_01_RampUp_and_Loop"), ShakerSoundLevel, Bumper2
      Case 1
        PlaySoundAtLevelExistingStaticLoop ("SY_TNA_REV02_Shaker_Level_02_RampUp_and_Loop"), ShakerSoundLevel, Bumper2
      Case 2
        PlaySoundAtLevelExistingStaticLoop ("SY_TNA_REV02_Shaker_Level_03_RampUp_and_Loop"), ShakerSoundLevel, Bumper2
    End Select
  Else
    Select Case ShakerIntensity
      Case 0
        PlaySoundAtLevelStatic ("SY_TNA_REV02_Shaker_Level_01_RampDown_Only"), ShakerSoundLevel, Bumper2
        StopSound "SY_TNA_REV02_Shaker_Level_01_RampUp_and_Loop"
      Case 1
        PlaySoundAtLevelStatic ("SY_TNA_REV02_Shaker_Level_02_RampDown_Only"), ShakerSoundLevel, Bumper2
        StopSound "SY_TNA_REV02_Shaker_Level_02_RampUp_and_Loop"
      Case 2
        PlaySoundAtLevelStatic ("SY_TNA_REV02_Shaker_Level_03_RampDown_Only"), ShakerSoundLevel, Bumper2
        StopSound "SY_TNA_REV02_Shaker_Level_03_RampUp_and_Loop"
    End Select
  End If
End Sub



'*******************************************
' ZGII: General Illumination
'*******************************************
dim gilvl:gilvl = 0
dim gilvl2:gilvl2 = 0

' Lower Playfield GI
Sub PFGI(pwm)
  dim xx
  For each xx in GI:xx.State = pwm: Next
  If pwm >= 0.5 And gilvl < 0.5 Then
    Sound_GI_Relay 1,KnockerPosition
  ElseIf pwm <= 0.4 And gilvl > 0.4 Then
    Sound_GI_Relay 0,KnockerPosition
  End If
  gilvl = pwm
End Sub

' Upper Playfield GI
Sub PFGI2(pwm)
  dim xx
  For each xx in GIU:xx.State = pwm: Next
  If pwm >= 0.5 And gilvl2 < 0.5 Then
    Sound_GI_Relay 1,KnockerPosition
  ElseIf pwm <= 0.4 And gilvl2 > 0.4 Then
    Sound_GI_Relay 0,KnockerPosition
  End If
  gilvl2 = pwm
End Sub



'*****************************************
' ZDRN: Drain, Trough, and Ball Release
'*****************************************

Sub sw11_Hit():   Controller.Switch(11) = 1: UpdateTrough: End Sub
Sub sw11_UnHit(): Controller.Switch(11) = 0: UpdateTrough: End Sub
Sub sw12_Hit():   Controller.Switch(12) = 1: UpdateTrough: End Sub
Sub sw12_UnHit(): Controller.Switch(12) = 0: UpdateTrough: End Sub
Sub sw13_Hit():   Controller.Switch(13) = 1: UpdateTrough: End Sub
Sub sw13_UnHit(): Controller.Switch(13) = 0: UpdateTrough: End Sub


'Ball Release
Sub SolOuthole(enabled)
  If enabled Then
    sw10.kick 60,20
    SoundSaucerKick 1,sw10
  End If
End Sub

Sub ReleaseBall(enabled)
  If enabled Then
    RandomSoundBallRelease sw11
    sw11.kick 60, 12
    UpdateTrough
  End If
End Sub

'Drain
Sub sw10_Hit()
  UpdateTrough
  Controller.Switch(10) = 1
  RandomSoundDrain sw10
End Sub

Sub sw10_UnHit()
  Controller.Switch(10) = 0
End Sub


'Trough
Sub UpdateTrough()
  UpdateTroughTimer.Interval = 100
  UpdateTroughTimer.Enabled = 1
End Sub

Sub UpdateTroughTimer_Timer()
  If sw11.BallCntOver = 0 Then sw12.kick 60, 9
  If sw12.BallCntOver = 0 Then sw13.kick 60, 9
  UpdateTroughTimer.Enabled = 0
End Sub





'*****************************************
' ZSWI: Switches
'*****************************************

' Stand Up Targets
Sub sw19_hit: STHit 19 : End Sub
Sub sw21_hit: STHit 21 : End Sub
Sub sw22_hit: STHit 22 : End Sub
Sub sw23_hit: STHit 23 : End Sub
Sub sw24_hit: STHit 24 : End Sub
Sub sw30_hit: STHit 30 : End Sub

' Drop Targets
 Sub sw27_Hit: DTHit 27 : End Sub
 Sub sw28_Hit: DTHit 28 : End Sub
 Sub sw29_Hit: DTHit 29 : End Sub

' Wire Triggers
Sub sw14_Hit:   Controller.Switch(14)=1: rightInlaneSpeedLimit: End Sub
Sub sw14_UnHit: Controller.Switch(14)=0: rightInlaneSpeedLimit: End Sub
Sub sw15_Hit:   Controller.Switch(15)=1: End Sub
Sub sw15_UnHit: Controller.Switch(15)=0: End Sub
Sub sw16_Hit:   Controller.Switch(16)=1: End Sub
Sub sw16_UnHit: Controller.Switch(16)=0: End Sub
Sub sw17_Hit:   Controller.Switch(17)=1: End Sub
Sub sw17_UnHit: Controller.Switch(17)=0: End Sub
Sub sw18_Hit:   Controller.Switch(18)=1: leftInlaneSpeedLimit: End Sub
Sub sw18_UnHit: Controller.Switch(18)=0: End Sub
Sub sw33_Hit:   controller.switch(33)=1: WireRampOff: End Sub
Sub sw33_unHit: controller.switch(33)=0: End Sub
Sub sw34_Hit:   controller.switch(34)=1: WireRampOff: End Sub
Sub sw34_unHit: controller.switch(34)=0: End Sub
Sub sw35_Hit:   controller.switch(35)=1: WireRampOff: End Sub
Sub sw35_unHit: controller.switch(35)=0: End Sub
Sub sw36_Hit:   controller.switch(36)=1: End Sub
Sub sw36_unHit: controller.switch(36)=0: End Sub

' Gates
Sub sw31_hit: vpmTimer.pulseSw 31 : End Sub
Sub sw32_hit: vpmTimer.pulseSw 32 : End Sub
Sub sw43_hit: vpmTimer.pulseSw 43 : End Sub
Sub sw44_hit: vpmTimer.pulseSw 44 : End Sub

' Subway
Sub sw38_Hit:  Controller.Switch(38)=1 : WireRampOff: playsound "Subway" : End Sub
Sub sw38_UnHit:Controller.Switch(38)=0:End Sub
Sub sw39_Hit:  Controller.Switch(39)=1 : End Sub
Sub sw39_unHit:Controller.Switch(39)=0:End Sub

' Spinners
Sub sw41_Spin: vpmTimer.PulseSw 41 : soundSpinner sw41 : End Sub

' Left Ramp Wire Triggers
Sub sw45_Hit:   controller.switch (45)=1 : End Sub
Sub sw45_unHit: controller.switch (45)=0:End Sub
Sub sw46_Hit:   controller.switch (46)=1 : End Sub
Sub sw46_unHit: controller.switch (46)=0:End Sub

' Shooter lane
Sub sw50_Hit:   controller.switch (50)=1 : End Sub
Sub sw50_unHit: controller.switch (50)=0:End Sub

' Bumpers
Sub Bumper1_Hit : vpmTimer.PulseSw(52) : RandomSoundBumperTop Bumper1: End Sub
Sub Bumper2_Hit : vpmTimer.PulseSw(53) : RandomSoundBumperMiddle Bumper2: End Sub
Sub Bumper3_Hit : vpmTimer.PulseSw(54) : RandomSoundBumperBottom Bumper3: End Sub



Sub leftInlaneSpeedLimit
  'Wylte's implementation
'    debug.print "Spin in: "& activeball.AngMomZ
'    debug.print "Speed in: "& activeball.vely
  If activeball.vely < 0 Then Exit Sub              'don't affect upwards movement
    activeball.AngMomZ = -abs(activeball.AngMomZ) * RndNum(3,6)
    If abs(activeball.AngMomZ) > 60 Then activeball.AngMomZ = 0.8 * activeball.AngMomZ
    If abs(activeball.AngMomZ) > 80 Then activeball.AngMomZ = 0.8 * activeball.AngMomZ
    If activeball.AngMomZ > 100 Then activeball.AngMomZ = RndNum(80,100)
    If activeball.AngMomZ < -100 Then activeball.AngMomZ = RndNum(-80,-100)

    If abs(activeball.vely) > 5 Then activeball.vely = 0.8 * activeball.vely
    If abs(activeball.vely) > 10 Then activeball.vely = 0.8 * activeball.vely
    If abs(activeball.vely) > 15 Then activeball.vely = 0.8 * activeball.vely
    If activeball.vely > 16 Then activeball.vely = RndNum(14,16)
    If activeball.vely < -16 Then activeball.vely = RndNum(-14,-16)
'    debug.print "Spin out: "& activeball.AngMomZ
'    debug.print "Speed out: "& activeball.vely
End Sub


Sub rightInlaneSpeedLimit
  'Wylte's implementation
'    debug.print "Spin in: "& activeball.AngMomZ
'    debug.print "Speed in: "& activeball.vely
  If activeball.vely < 0 Then Exit Sub              'don't affect upwards movement

    activeball.AngMomZ = abs(activeball.AngMomZ) * RndNum(2,4)
    If abs(activeball.AngMomZ) > 60 Then activeball.AngMomZ = 0.8 * activeball.AngMomZ
    If abs(activeball.AngMomZ) > 80 Then activeball.AngMomZ = 0.8 * activeball.AngMomZ
    If activeball.AngMomZ > 100 Then activeball.AngMomZ = RndNum(80,100)
    If activeball.AngMomZ < -100 Then activeball.AngMomZ = RndNum(-80,-100)

  If abs(activeball.vely) > 5 Then activeball.vely = 0.8 * activeball.vely
    If abs(activeball.vely) > 10 Then activeball.vely = 0.8 * activeball.vely
    If abs(activeball.vely) > 15 Then activeball.vely = 0.8 * activeball.vely
    If activeball.vely > 16 Then activeball.vely = RndNum(14,16)
    If activeball.vely < -16 Then activeball.vely = RndNum(-14,-16)
'    debug.print "Spin out: "& activeball.AngMomZ
'    debug.print "Speed out: "& activeball.vely
End Sub



'***********************************
' ZSTA: Stand-up Targets
'***********************************



''''' Start of Roth code

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
Dim ST19, ST21, ST22, ST23, ST24, ST30

'Set array with stand-up target objects
'
'StandupTargetvar = Array(primary, prim, swtich)
'   primary:  vp target to determine target hit
'   prim:    primitive target used for visuals and animation
'          IMPORTANT!!!
'          transy must be used to offset the target animation
'   switch:  ROM switch number
'   animate:  Arrary slot for handling the animation instrucitons, set to 0
'
'You will also need to add a secondary hit object for each stand up (name sw11o, sw12o, and sw13o on the example Table1)
'these are inclined primitives to simulate hitting a bent target and should provide so z velocity on high speed impacts


Set ST19 = (new StandupTarget)(sw19, BM_sw19, 19, 0)
Set ST21 = (new StandupTarget)(sw21, BM_sw21, 21, 0)
Set ST22 = (new StandupTarget)(sw22, BM_sw22, 22, 0)
Set ST23 = (new StandupTarget)(sw23, BM_sw23, 23, 0)
Set ST24 = (new StandupTarget)(sw24, BM_sw24, 24, 0)
Set ST30 = (new StandupTarget)(sw30, BM_sw30, 30, 0)


'Add all the Stand-up Target Arrays to Stand-up Target Animation Array
'   STAnimationArray = Array(ST1, ST2, ....)
Dim STArray
STArray = Array(ST19, ST21, ST22, ST23, ST24, ST30)

'Configure the behavior of Stand-up Targets
Const STAnimStep = 1.5  'vpunits per animation step (control return to Start)
Const STMaxOffset = 9   'max vp units target moves when hit

Const STMass = 0.2    'Mass of the Stand-up Target (between 0 and 1), higher values provide more resistance


'''''' STAND-UP TARGETS FUNCTIONS

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
    vpmTimer.PulseSw switch
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
'*   END STAND-UP TARGETS
'******************************************************





'***********************************
' ZDTA: Drop Targets
'***********************************



Sub UpdateTargetShadows
' ' If the upper GI is off, shadows should be off.
' If LampState(102) = 0 Then
'   SetLamp 105, 0
'   SetLamp 106, 0
'   SetLamp 107, 0
' Else
'   SetLamp 105, 1+Controller.Switch(27)
'   SetLamp 106, 1+Controller.Switch(28)
'   SetLamp 107, 1+Controller.Switch(29)
' End If
End Sub


'Solenoids

Sub SolDropReset(enabled)
  If enabled then
    DTRaise 27
    DTRaise 28
    DTRaise 29
    RandomSoundDropTargetReset BM_sw28
    UpdateTargetShadows
'   Dim xx: For each xx in DTShadows: xx.visible = True: Next
  End If
End Sub


'Actions

Sub DTAction(switchid)
' Select Case switchid
'   Case 2: DTShadows(0).visible = False
'   Case 12: DTShadows(1).visible = False
'   Case 22: DTShadows(2).visible = False
'   Case 32: DTShadows(3).visible = False
' End Select
End Sub

''''''DROP TARGETS INITIALIZATION


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
'         Use the function DTDropped(switchid) to check a target's drop status.

Set DT27 = (new DropTarget)(sw27, sw27a, BM_sw27, 27, 0, False)
Set DT28 = (new DropTarget)(sw28, sw28a, BM_sw28, 28, 0, False)
Set DT29 = (new DropTarget)(sw29, sw29a, BM_sw29, 29, 0, False)


Dim DTArray
DTArray = Array(DT27, DT28, DT29)

'Configure the behavior of Drop Targets.
Const DTDropSpeed = 90 'in milliseconds
Const DTDropUpSpeed = 40 'in milliseconds
Const DTDropUnits = 50 'VP units primitive drops so top of at or below the playfield
Const DTDropUpUnits = 10 'VP units primitive raises above the up position on drops up
Const DTMaxBend = 8 'max degrees primitive rotates when hit
Const DTDropDelay = 20 'time in milliseconds before target drops (due to friction/impact of the ball)
Const DTRaiseDelay = 40 'time in milliseconds before target drops back to normal up position after the solenoid fires to raise the target
Const DTBrickVel = 30 'velocity at which the target will brick, set to '0' to disable brick
Const DTEnableBrick = 0 'Set to 0 to disable bricking, 1 to enable bricking
Const DTMass = 0.2 'Mass of the Drop Target (between 0 and 1), higher values provide more resistance


'''''''DROP TARGETS FUNCTIONS


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
      controller.Switch(Switchid) = 1
      DTAction Switchid
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
          gBOT(b).velz = 10
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
    controller.Switch(Switchid) = 0
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
'*   END DROP TARGETS
'******************************************************





'***********************************
' ZSLG: Slingshots
'***********************************

Dim RStep, Lstep

Sub LeftSlingShot_Slingshot
  LS.VelocityCorrect(ActiveBall)
  vpmTimer.PulseSw 55
  RandomSoundSlingShotLeft EndPoint1LS
  LStep = 0
  LeftSlingShot.TimerInterval = 17
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
  Dim BP
  Dim x1, x2, y: x1 = True:x2 = False:y = 25
  Select Case LStep
    Case 3:x1 = False:x2= True: y = 15
    Case 4:x1 = False:x2 = False:y = 0:LeftSlingShot.TimerEnabled = 0
  End Select

  For Each BP in BP_LSling1 : BP.Visible = x1: Next
  For Each BP in BP_LSling2 : BP.Visible = x2: Next
  For Each BP in BP_LEMK : BP.transx = y: Next

  LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
  RS.VelocityCorrect(ActiveBall)
  vpmTimer.PulseSw 56
  RandomSoundSlingShotRight EndPoint1RS
  RStep = 0
  RightSlingShot.TimerInterval = 17
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
  Dim BP
  Dim x1, x2, y: x1 = True:x2 = False:y = 25
    Select Case RStep
        Case 2:x1 = False:x2 = True:y = 15
        Case 3:x1 = False:x2 = False:y = 0:RightSlingShot.TimerEnabled = 0
    End Select

  For Each BP in BP_RSling1 : BP.Visible = x1: Next
  For Each BP in BP_RSling2 : BP.Visible = x2: Next
  For Each BP in BP_REMK : BP.transx = y: Next

    RStep = RStep + 1
End Sub


'***********************************
' ZVUK: VUKs
'***********************************
Dim KickerBall20, KickerBall37, KickerBall40

Sub KickBall(kball, kangle, kvel, kvelz, kzlift)
  dim rangle
  rangle = PI * (kangle - 90) / 180

  kball.z = kball.z + kzlift
  kball.velz = kvelz
  kball.velx = cos(rangle)*kvel
  kball.vely = sin(rangle)*kvel
End Sub

'********
'Left VUK
'********

Sub sw20_Hit
' msgbox activeball.z
    set KickerBall20 = activeball
    Controller.Switch(20) = 1
    SoundSaucerLock
  sw20.timerinterval=10
  sw20.timerenabled=true
  GateBlocker.IsDropped = false
End Sub

Sub sw20_unHit
  Controller.Switch(20) = 0
  sw20.timerenabled=false
  GateBlocker.IsDropped = true
end sub

sub sw20_timer
' debug.print "vuk left: " & KickerBall20.z
  'prevents ball wiggle in saucer
  KickerBall20.x=sw20.x
  KickerBall20.y=sw20.y
  KickerBall20.z=6.5 'a bit higher so it wont touch the collidable bottom
  KickerBall20.angmomx=0
  KickerBall20.angmomy=0
  KickerBall20.angmomz=0
  KickerBall20.velx=0
  KickerBall20.vely=0
  KickerBall20.velz=0
  If Controller.Switch(20) = 0 Then me.timerenabled = false
end sub

Sub VUKLeft(Enable)
  sw20.timerenabled=false
    If Enable then
    If Controller.Switch(20) <> 0 Then
      KickBall KickerBall20, 250, 15, 5, 10
      SoundSaucerKick 1, sw20
    End If
  End If
End Sub

'********
'Top VUK
'********

Sub sw37_Hit
' msgbox activeball.z
    set KickerBall37 = activeball
    Controller.Switch(37) = 1
    SoundSaucerLock
  WireRampOff
' sw37.timerinterval=10
' sw37.timerenabled=true
End Sub

Sub sw37_unHit
  Controller.Switch(37) = 0
' sw37.timerenabled=false
end sub

'sub sw39_timer
' If Controller.Switch(39) = 0 Then me.timerenabled = false
'end sub

Sub VUKTop(Enable)
' sw37.timerenabled=false
    If Enable then
    If Controller.Switch(37) <> 0 Then
      KickBall KickerBall37, 0, 0, 50, 10
      SoundSaucerKick 1, sw37
    End If
  End If
End Sub

'**********
'Bottom VUK
'**********

Sub sw40_Hit
' msgbox activeball.z
    set KickerBall40 = activeball
    Controller.Switch(40) = 1
    SoundSaucerLock
' sw40.timerinterval=10
' sw40.timerenabled=true
End Sub

Sub sw40_unHit
  Controller.Switch(40) = 0
' sw40.timerenabled=false
end sub

'sub sw40_timer
' If Controller.Switch(40) = 0 Then me.timerenabled = false
'end sub

Sub VUKBot(Enable)
' sw40.timerenabled=false
    If Enable then
    If Controller.Switch(40) <> 0 Then
      KickBall KickerBall40, 0, 0, 50, 10
      SoundSaucerKick 1, sw40
    End If
  End If
End Sub



'***********************************
' ZCND: California Nevada Diverter
'***********************************

dim FaultOpen : FaultOpen = False

Sub SolFault(enabled)
  If enabled Then
    FaultOpen = Not FaultOpen
    If FaultOpen Then
      OpenFaultTimer.enabled=1
      Pdiverter1.Collidable= 0
      Pdiverter2.Collidable= 1
    Else
      CloseFaultTimer.enabled=1
      Pdiverter1.Collidable= 1
      Pdiverter2.Collidable= 0
    End If
  End If
End Sub

Dim BridgeStep: BridgeStep = 0

Sub OpenFaultTimer_Timer
  Dim BP
    Select Case BridgeStep
    Case 0: For Each BP In BP_diverter: BP.TransX = 0: Next   : BridgeStep =1
    Case 1: For Each BP In BP_diverter: BP.TransX = -6: Next  : BridgeStep =2
    Case 2: For Each BP In BP_diverter: BP.TransX = -12: Next : BridgeStep =3
    Case 3: For Each BP In BP_diverter: BP.TransX = -22: Next : BridgeStep =4
        Case 4: For Each BP In BP_diverter: BP.TransX = -32: Next : BridgeStep =5
        Case 5: For Each BP In BP_diverter: BP.TransX = -53: Next : BridgeStep =0 : controller.switch(42) = 1 : OpenFaultTimer.enabled = 0
    End Select
End Sub

Sub CloseFaultTimer_Timer
  Dim BP
    Select Case BridgeStep
    Case 0: For Each BP In BP_diverter: BP.TransX = -53: Next : BridgeStep =1
        Case 1: For Each BP In BP_diverter: BP.TransX = -32: Next : BridgeStep =2
    Case 2: For Each BP In BP_diverter: BP.TransX = -22: Next : BridgeStep =3
    Case 3: For Each BP In BP_diverter: BP.TransX = -12: Next : BridgeStep =4
        Case 4: For Each BP In BP_diverter: BP.TransX = -6: Next  : BridgeStep =5
        Case 5: For Each BP In BP_diverter: BP.TransX = 0: Next   : BridgeStep =0 : controller.switch(42) = 0 : CloseFaultTimer.enabled = 0
    End Select
End Sub



'***********************************
' ZEIA: Earthquake Institute Animation
'***********************************


' In the real production Earthshakers, the Institute building is just a
' static toy.  All it does is flash its lights.  But in the pre-production
' prototypes, the building had a sliding mechanism and a motor that made it
' sink into the playfield and rise back up on certain game events.  The
' ROM code that controls the building movement is still present in the
' production ROMs, so some Earthshaker owners have modded their tables to
' restore the moving building.  This code implements the building animation
' in our simulated version.
'
' To animate the building, we simply move all of the primitives in unison.
' The animation is extremely simple in that all the building does is move straight up and down.
'


' Building travel parameters.  The range of motion is about 3/4 of the building
' height, and it takes about 3 seconds to cover that distance.
dim InstZ, InstDZ, InstDZUp, InstDZDown, InstZMin, InstZMax, InstZRange, InstTravelTime, InstMotorOn, InstZone
InstZMin = -150 : InstZMax = 0   ' travel limits, as Z Translations for the building primitives
InstTravelTime = 3000            ' one-way travel time in milliseconds
InstMotorOn = 0                  ' start with motor turned off

' figure the distance per timer events (total distance divided by total time,
' multiplied by the time interval per event)
InstZRange = InstZMax - InstZMin
InstDZUp = InstZRange / InstTravelTime * InstituteTimer.Interval
InstDZDown = -InstDZUp           ' use the same speed up and down

' start fully up (set Z to max, zone to top (zone 0=bottom, 1=between top and bottom, 2=top))
InstZ = InstZMax
InstZone = 2


Sub InstituteTimer_Timer
  ' Handle the motor
  if InstMotorOn then
    ' move up or down in the current direction
    InstZ = InstZ + InstDZ

    ' check for limits
    if InstZ >= InstZMax then
      ' all the the way up - stop here and reverse directions
      InstZ = InstZMax
      InstDZ = InstDZDown

      ' trip switches to tell the ROM we're in the UP position
      controller.switch(25) = false
      controller.switch(26) = true
      InstZone = 3
    elseif InstZ <= InstZMin then
      ' all the way down - stop here and reverse directions
      InstZ = InstZMin
      InstDZ = InstDZUp

      ' trip switches to tell the ROM we're in the DOWN position
      controller.switch(25) = true
      controller.switch(26) = true
      InstZone = 0
    elseif InstZone = 3 and InstZ < InstZMax - InstZRange/3 then
      ' we were at the top, and now we're 1/3 of the way down - the
      ' original script set both switches off at this position
      controller.switch(25) = false
      controller.switch(26) = false
      InstZone = 2
    elseif InstZone = 2 and InstZ < InstZMin + InstZRange/3 then
      ' crossing 2/3 of the way down - the original script set 25 ON
      ' and 26 OFF at this position
      controller.switch(25) = true
      controller.switch(26) = false
    end if

    Dim BP: For Each BP In BP_EIprim
      BP.TransZ = InstZ
    next
  end if
End Sub


' ROM motor control interface
sub InstituteDrop(enabled)
  If EIMod = 0 Then Exit Sub
  InstMotorOn = enabled
  InstituteTimer.Enabled = enabled
  If enabled Then
    PlaySoundAtLevelStaticLoop "motor", MechVolume, Bumper2
  Else
    StopSound "motor"
  End If
End Sub


' IE Lightmap control

Const EILightOffset = 0   'how dim the lights are when EI is down (fraction)

Sub L17a_Animate
  Dim f: f = L17a.GetInPlayIntensity / L17a.Intensity
  Dim op: op = 100*f*abs(EILightOffset+(1-EILightOffset)*(InstZRange+InstZ)/InstZRange)
  Dim BL: For Each BL In BL_IN_L17: BL.opacity = op: Next
  'debug.print "L17a_Animate f=" & f & " opacity=" & op & " R=" & abs(EILightOffset+(1-EILightOffset)*(InstZRange+InstZ)/InstZRange)
End Sub

Sub L18a_Animate
  Dim f: f = L18a.GetInPlayIntensity / L18a.Intensity
  Dim op: op = 100*f*abs(EILightOffset+(1-EILightOffset)*(InstZRange+InstZ)/InstZRange)
  Dim BL: For Each BL In BL_IN_L18: BL.opacity = op: Next
End Sub

Sub L19a_Animate
  Dim f: f = L19a.GetInPlayIntensity / L19a.Intensity
  Dim op: op = 100*f*abs(EILightOffset+(1-EILightOffset)*(InstZRange+InstZ)/InstZRange)
  Dim BL: For Each BL In BL_IN_L19: BL.opacity = op: Next
End Sub

Sub L20a_Animate
  Dim f: f = L20a.GetInPlayIntensity / L20a.Intensity
  Dim op: op = 100*f*abs(EILightOffset+(1-EILightOffset)*(InstZRange+InstZ)/InstZRange)
  Dim BL: For Each BL In BL_IN_L20: BL.opacity = op: Next
End Sub

Sub L21a_Animate
  Dim f: f = L21a.GetInPlayIntensity / L21a.Intensity
  Dim op: op = 100*f*abs(EILightOffset+(1-EILightOffset)*(InstZRange+InstZ)/InstZRange)
  Dim BL: For Each BL In BL_IN_L21: BL.opacity = op: Next
End Sub

Sub L22a_Animate
  Dim f: f = L22a.GetInPlayIntensity / L22a.Intensity
  Dim op: op = 100*f*abs(EILightOffset+(1-EILightOffset)*(InstZRange+InstZ)/InstZRange)
  Dim BL: For Each BL In BL_IN_L22: BL.opacity = op: Next
End Sub

Sub L23a_Animate
  Dim f: f = L23a.GetInPlayIntensity / L23a.Intensity
  Dim op: op = 100*f*abs(EILightOffset+(1-EILightOffset)*(InstZRange+InstZ)/InstZRange)
  Dim BL: For Each BL In BL_IN_L23: BL.opacity = op: Next
End Sub

Sub L24a_Animate
  Dim f: f = L24a.GetInPlayIntensity / L24a.Intensity
  Dim op: op = 100*f*abs(EILightOffset+(1-EILightOffset)*(InstZRange+InstZ)/InstZRange)
  Dim BL: For Each BL In BL_IN_L24: BL.opacity = op: Next
End Sub

Sub L25a_Animate
  Dim f: f = L25a.GetInPlayIntensity / L25a.Intensity
  Dim op: op = 100*f*abs(EILightOffset+(1-EILightOffset)*(InstZRange+InstZ)/InstZRange)
  Dim BL: For Each BL In BL_IN_L25: BL.opacity = op: Next
End Sub



'***************************************************************
'  ZABS: Ambient ball shadows
'***************************************************************

'Ambient (Room light source)
Const AmbientBSFactor = 0.9    '0 To 1, higher is darker
Const AmbientMovement = 1    '1+ higher means more movement as the ball moves left and right
Const offsetX = 0        'Offset x position under ball (These are if you want to change where the "room" light is for calculating the shadow position,)
Const offsetY = 0        'Offset y position under ball (^^for example 5,5 if the light is in the back left corner)

' *** Trim or extend these to match the number of balls/primitives/flashers on the table!  (will throw errors if there aren't enough objects)
Dim objBallShadow(5)

'Initialization
BSInit

Sub BSInit()
  Dim iii
  'Prepare the shadow objects before play begins
  For iii = 0 To tnob - 1
    Set objBallShadow(iii) = Eval("BallShadow" & iii)
    objBallShadow(iii).material = "BallShadow" & iii
    UpdateMaterial objBallShadow(iii).material,1,0,0,0,0,0,AmbientBSFactor,RGB(0,0,0),0,0,False,True,0,0,0,0
    objBallShadow(iii).Z = 1 + iii / 1000 + 1.04
    objBallShadow(iii).visible = 0
  Next
End Sub


Sub BSUpdate
  Dim s

  'The Magic happens now
  For s = 0 To UBound(gBOT)
    ' *** Normal "ambient light" ball shadow

    '** If on main pf
    If gBOT(s).Z > 20 and gBOT(s).Z < 30 Then
      objBallShadow(s).visible = 1
      objBallShadow(s).X = gBOT(s).X + (gBOT(s).X - (tablewidth / 2)) / (Ballsize / AmbientMovement) + offsetX
      objBallShadow(s).Y = gBOT(s).Y + offsetY
      objBallShadow(s).Z = gBOT(s).Z + s/1000 + 1.04 - 25
    '** Under pf, no shadow
    Else
      objBallShadow(s).visible = 0
    End If
  Next
End Sub




'*****************************************
' ZLED: LEDs
'*****************************************



' Desktop Digits Display
'************************
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


Dim DigitsVR(32)

DigitsVR(0)=Array(D1,D2,D3,D4,D5,D6,D7,D8,D9,D10,D11,D12,D13,D14,D15)
DigitsVR(1)=Array(D16,D17,D18,D19,D20,D21,D22,D23,D24,D25,D26,D27,D28,D29,D30)
DigitsVR(2)=Array(D31,D32,D33,D34,D35,D36,D37,D38,D39,D40,D41,D42,D43,D44,D45)
DigitsVR(3)=Array(D46,D47,D48,D49,D50,D51,D52,D53,D54,D55,D56,D57,D58,D59,D60)
DigitsVR(4)=Array(D61,D62,D63,D64,D65,D66,D67,D68,D69,D70,D71,D72,D73,D74,D75)
DigitsVR(5)=Array(D76,D77,D78,D79,D80,D81,D82,D83,D84,D85,D86,D87,D88,D89,D90)
DigitsVR(6)=Array(D91,D92,D93,D94,D95,D96,D97,D98,D99,D100,D101,D102,D103,D104,D105)
DigitsVR(7)=Array(D106,D107,D108,D109,D110,D111,D112,D113,D114,D115,D116,D117,D118,D119,D120)
DigitsVR(8)=Array(D121,D122,D123,D124,D125,D126,D127,D128,D129,D130,D131,D132,D133,D134,D135)
DigitsVR(9)=Array(D136,D137,D138,D139,D140,D141,D142,D143,D144,D145,D146,D147,D148,D149,D150)
DigitsVR(10)=Array(D151,D152,D153,D154,D155,D156,D157,D158,D159,D160,D161,D162,D163,D164,D165)
DigitsVR(11)=Array(D166,D167,D168,D169,D170,D171,D172,D173,D174,D175,D176,D177,D178,D179,D180)
DigitsVR(12)=Array(D181,D182,D183,D184,D185,D186,D187,D188,D189,D190,D191,D192,D193,D194,D195)
DigitsVR(13)=Array(D196,D197,D198,D199,D200,D201,D202,D203,D204,D205,D206,D207,D208,D209,D210)
DigitsVR(14)=Array(D211,D212,D213,D214,D215,D216,D217,D218,D219,D220,D221,D222,D223,D224,D225)
DigitsVR(15)=Array(D226,D227,D228,D229,D230,D231,D232,D233,D234,D235,D236,D237,D238,D239,D240)

DigitsVR(16)=Array(D241,D242,D243,D244,D245,D246,D247,D248,D249,D250,D251,D252,D253,D254,D255)
DigitsVR(17)=Array(D256,D257,D258,D259,D260,D261,D262,D263,D264,D265,D266,D267,D268,D269,D270)
DigitsVR(18)=Array(D271,D272,D273,D274,D275,D276,D277,D278,D279,D280,D281,D282,D283,D284,D285)
DigitsVR(19)=Array(D286,D287,D288,D289,D290,D291,D292,D293,D294,D295,D296,D297,D298,D299,D300)
DigitsVR(20)=Array(D301,D302,D303,D304,D305,D306,D307,D308,D309,D310,D311,D312,D313,D314,D315)
DigitsVR(21)=Array(D316,D317,D318,D319,D320,D321,D322,D323,D324,D325,D326,D327,D328,D329,D330)
DigitsVR(22)=Array(D331,D332,D333,D334,D335,D336,D337,D338,D339,D340,D341,D342,D343,D344,D345)
DigitsVR(23)=Array(D346,D347,D348,D349,D350,D351,D352,D353,D354,D355,D356,D357,D358,D359,D360)
DigitsVR(24)=Array(D361,D362,D363,D364,D365,D366,D367,D368,D369,D370,D371,D372,D373,D374,D375)
DigitsVR(25)=Array(D376,D377,D378,D379,D380,D381,D382,D383,D384,D385,D386,D387,D388,D389,D390)
DigitsVR(26)=Array(D391,D392,D393,D394,D395,D396,D397,D398,D399,D400,D401,D402,D403,D404,D405)
DigitsVR(27)=Array(D406,D407,D408,D409,D410,D411,D412,D413,D414,D415,D416,D417,D418,D419,D420)
DigitsVR(28)=Array(D421,D422,D423,D424,D425,D426,D427,D428,D429,D430,D431,D432,D433,D434,D435)
DigitsVR(29)=Array(D436,D437,D438,D439,D440,D441,D442,D443,D444,D445,D446,D447,D448,D449,D450)
DigitsVR(30)=Array(D451,D452,D453,D454,D455,D456,D457,D458,D459,D460,D461,D462,D463,D464,D465)
DigitsVR(31)=Array(D466,D467,D468,D469,D470,D471,D472,D473,D474,D475,D476,D477,D478,D479,D480)


Sub InitDigitsVR()
  dim tmp, x, obj
  for x = 0 to uBound(DigitsVR)
    if IsArray(DigitsVR(x) ) then
      For each obj in DigitsVR(x)
        obj.height = obj.height + 18
        FadeDisplay obj, 0
      next
    end If
  Next
End Sub

InitDigitsVR


dim DisplayColor
DisplayColor =  RGB(255,40,1)

Sub DisplayTimer_Timer
  If Not (DesktopMode = True or RenderingMode = 2) Then Exit Sub

    Dim ChgLED, ii, num, chg, stat, obj
    ChgLED=Controller.ChangedLEDs(&Hffffffff, &Hffffffff)

    If Not IsEmpty(ChgLED)Then
    For ii=0 To UBound(chgLED)
      num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)
      If DesktopMode = True and RenderingMode <> 2 Then
        if (num < 32) then
          For Each obj In Digits(num)
            If chg And 1 Then obj.State=stat And 1
            chg=chg\2 : stat=stat\2
          Next
        Else
        end if
      Elseif RenderingMode = 2 then
        'VR display digits
        If (num < 32) then
          For Each obj In DigitsVR(num)
            If chg And 1 Then FadeDisplay obj, stat And 1
            chg=chg\2 : stat=stat\2
          Next
        End If
      End If
    Next
    End If

End Sub


Sub FadeDisplay(object, onoff)
  If OnOff = 1 Then
    object.color = DisplayColor

    If VRRoom=0 Then
      Object.Opacity = 0   'Brightness of the displays are off in non VR Room
    Else
      Object.Opacity = 1500   'Brightness of the displays when they are on in VR Room
    End If
  Else
    Object.Color = RGB(1,1,1)

    If VRRoom=0 Then
      Object.Opacity = 0   'Brightness of the displays when they are off in non VR Room
    Else
      Object.Opacity = 400   'Brightness of the displays when they are off in VR Room
    End If
  End If
End Sub





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
' ZNFF:  FLIPPER CORRECTIONS by nFozzy
'******************************************************



'******************************************************
' Flippers Polarity (Select appropriate sub based on era)
'******************************************************

Dim LF : Set LF = New FlipperPolarity
Dim RF : Set RF = New FlipperPolarity

InitPolarity

'*******************************************
'  Late 80's early 90's

Sub InitPolarity()
  dim x, a : a = Array(LF, RF)
  for each x in a
    x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
    x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
    x.enabled = True
    x.TimeDelay = 60
    x.DebugOn=False ' prints some info in debugger

    x.AddPt "Polarity", 0, 0, 0
    x.AddPt "Polarity", 1, 0.05, - 5
    x.AddPt "Polarity", 2, 0.16, - 5
    x.AddPt "Polarity", 3, 0.22, - 0
    x.AddPt "Polarity", 4, 0.25, - 0
    x.AddPt "Polarity", 5, 0.3, - 2
    x.AddPt "Polarity", 6, 0.4, - 3
    x.AddPt "Polarity", 7, 0.5, - 4.0
    x.AddPt "Polarity", 8, 0.7, - 3.5
    x.AddPt "Polarity", 9, 0.75, - 3.0
    x.AddPt "Polarity", 10, 0.8, - 2.5
    x.AddPt "Polarity", 11, 0.85, - 2.0
    x.AddPt "Polarity", 12, 0.9, - 1.5
    x.AddPt "Polarity", 13, 0.95, - 1.0
    x.AddPt "Polarity", 14, 1, - 0.5
    x.AddPt "Polarity", 15, 1.1, 0
    x.AddPt "Polarity", 16, 1.3, 0

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

Dim LFPress, RFPress, LFCount, RFCount, ULFPress
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

Const LiveCatch = 24 '16
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
' These are data mined bounce curves,
' dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
' Requires tracking ballspeed to calculate COR

Sub zCol_Rubber_LeftSling_Hit: RubbersD.dampen Activeball: End Sub
Sub zCol_Rubber_RightSling_Hit: RubbersD.dampen Activeball: End Sub

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
' END SLINGSHOT CORRECTION FUNCTIONS
'******************************************************






'******************************************************
'   ZBOU: VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************

Const TargetBouncerEnabled = 1    '0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor  = 1  'Level of bounces. Recommmended value of 0.7

Sub TargetBouncer(aBall,defvalue)
  Dim zMultiplier, vel, vratio
  If TargetBouncerEnabled = 1 And aball.z < 30 Then
    '   debug.print "velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz
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
    '   debug.print "---> velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz
    '   debug.print "conservation check: " & BallSpeed(aBall)/vel
  End If
End Sub


'Add targets or posts to the TargetBounce collection if you want to activate the targetbouncer code from them
Sub TargetBounce_Hit(idx)
  TargetBouncer ActiveBall, 1
End Sub




'*******************************************
'  ZSFX: Mechanical Sound effects by Fleep
'*******************************************



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
  PlaySound playsoundparams, 0, min(aVol,1) * MechVolume, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelExistingStatic(playsoundparams, aVol, tableobj)
  PlaySound playsoundparams, 0, min(aVol,1) * MechVolume, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelExistingStaticLoop(playsoundparams, aVol, tableobj)
    PlaySound playsoundparams, -1, min(aVol,1) * MechVolume, AudioPan(tableobj), 0, 0, 1, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticLoop(playsoundparams, aVol, tableobj)
  PlaySound playsoundparams, - 1, min(aVol,1) * MechVolume, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticRandomPitch(playsoundparams, aVol, randomPitch, tableobj)
  PlaySound playsoundparams, 0, min(aVol,1) * MechVolume, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelActiveBall(playsoundparams, aVol)
  PlaySound playsoundparams, 0, min(aVol,1) * MechVolume, AudioPan(ActiveBall), 0, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLevelExistingActiveBall(playsoundparams, aVol)
  PlaySound playsoundparams, 0, min(aVol,1) * MechVolume, AudioPan(ActiveBall), 0, 0, 1, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLeveTimerActiveBall(playsoundparams, aVol, ballvariable)
  PlaySound playsoundparams, 0, min(aVol,1) * MechVolume, AudioPan(ballvariable), 0, 0, 0, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelTimerExistingActiveBall(playsoundparams, aVol, ballvariable)
  PlaySound playsoundparams, 0, min(aVol,1) * MechVolume, AudioPan(ballvariable), 0, 0, 1, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelRoll(playsoundparams, aVol, pitch)
  PlaySound playsoundparams, - 1, min(aVol,1) * MechVolume, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

' Previous Positional Sound Subs

Sub PlaySoundAt(soundname, tableobj)
  PlaySound soundname, 1, 1 * MechVolume, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtVol(soundname, tableobj, aVol)
  PlaySound soundname, 1, min(aVol,1) * MechVolume, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname)
  PlaySoundAt soundname, ActiveBall
End Sub

Sub PlaySoundAtBallVol (Soundname, aVol)
  PlaySound soundname, 1,min(aVol,1) * MechVolume, AudioPan(ActiveBall), 0,0,0, 1, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtBallVolM (Soundname, aVol)
  PlaySound soundname, 1,min(aVol,1) * MechVolume, AudioPan(ActiveBall), 0,0,0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtVolLoops(sound, tableobj, Vol, Loops)
  PlaySound sound, Loops, Vol * MechVolume, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
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
  PlaySound ("Nudge_" & Int(Rnd * 2) + 1), 0, NudgeLeftSoundLevel * MechVolume, - 0.1, 0.25
End Sub

Sub SoundNudgeRight()
  PlaySound ("Nudge_" & Int(Rnd * 2) + 1), 0, NudgeRightSoundLevel * MechVolume, 0.1, 0.25
End Sub

Sub SoundNudgeCenter()
  PlaySound ("Nudge_" & Int(Rnd * 2) + 1), 0, NudgeCenterSoundLevel * MechVolume, 0, 0.25
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

Sub RandomSoundTargetHitSecret(aBall)
  PlaySoundAtLevelStatic SoundFX("Target_Hit_" & Int(Rnd * 4) + 1,DOFTargets), Vol(aBall) * TargetSoundFactor, aBall
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

Sub RandomSoundBallBounceRamp(aBall)
  Select Case Int(Rnd * 9) + 1
    Case 1
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_1"), 1, aBall
    Case 2
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_2"), 1 * 0.5, aBall
    Case 3
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_3"), 1 * 0.8, aBall
    Case 4
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_4"), 1 * 0.5, aBall
    Case 5
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_5"), 1, aBall
    Case 6
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_1"), 1 * 0.2, aBall
    Case 7
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_2"), 1 * 0.2, aBall
    Case 8
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_5"), 1 * 0.2, aBall
    Case 9
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_7"), 1 * 0.3, aBall
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

  PlaySound (snd), 0, CSng(velocity) ^ 2 / 200 * BallWithBallCollisionSoundFactor * MechVolume, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
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

'/////////////////////////////  COIN IN  ////////////////////////////

Sub SoundCoinIn
  Select Case Int(Rnd * 3)
    Case 0
      PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
    Case 1
      PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
    Case 2
      PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
  End Select
End Sub


'******************************************************
'****  END MECHANICAL SOUNDS
'******************************************************




'******************************************************
' ZBRL:  BALL ROLLING AND DROP SOUNDS
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

' ' stop the sound of deleted balls
' For b = UBound(gBOT) + 1 To tnob - 1
'   rolling(b) = False
'   StopSound("BallRoll_" & b)
' Next

  ' exit the sub if no balls on the table
  If UBound(gBOT) =  - 1 Then Exit Sub

  ' play the rolling sound for each ball
  For b = 0 To UBound(gBOT)
    If BallVel(gBOT(b)) > 1 And gBOT(b).z < 30 Then
      rolling(b) = True
      PlaySound ("BallRoll_" & b), - 1, VolPlayfieldRoll(gBOT(b)) * BallRollVolume * MechVolume, AudioPan(gBOT(b)), 0, PitchPlayfieldRoll(gBOT(b)), 1, 0, AudioFade(gBOT(b))
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
Dim RampBalls(4,2)
'x,0 = ball x,1 = ID, 2 = Protection against ending early (minimum amount of updates)

'0,0 is boolean on/off, 0,1 unused for now
RampBalls(0,0) = False

' RampType
' Setup: Set this array to the number Total number of balls that can be tracked at one time + 1.  5 ball multiball then set value to 6
' Description: Array type indexed on BallId and a values used to deterimine what type of ramp the ball is on: False = Wire Ramp, True = Plastic Ramp
Dim RampType(4)

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
          PlaySound("RampLoop" & x), - 1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * MechVolume, AudioPan(RampBalls(x,0)), 0, BallPitchV(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
          StopSound("wireloop" & x)
        Else
          StopSound("RampLoop" & x)
          PlaySound("wireloop" & x), - 1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * MechVolume, AudioPan(RampBalls(x,0)), 0, BallPitch(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
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



'  Ramp Triggers
'***************

Const DebugRampTriggers = False

Sub RampTrigger1_Hit
  if DebugRampTriggers=True Then debug.print "RampTrigger1_Hit"
  WireRampOn True
End Sub
Sub RampTrigger1_UnHit
  if activeball.velx > 0 then WireRampOff
End Sub


Sub RampTrigger2_Hit
  if DebugRampTriggers=True Then debug.print "RampTrigger2_Hit"
  WireRampOn True
End Sub
Sub RampTrigger2_UnHit
  if activeball.vely > 0 then WireRampOff
End Sub


Sub RampTrigger3_Hit
  if DebugRampTriggers=True Then debug.print "RampTrigger3_Hit"
  WireRampOn True
End Sub
Sub RampTrigger3_UnHit
  if activeball.velx < 0 then WireRampOff
End Sub


Sub RampTrigger4_Hit
  if DebugRampTriggers=True Then debug.print "RampTrigger4_Hit"
  WireRampOn False
End Sub


Sub RampTrigger5_Hit
  if DebugRampTriggers=True Then debug.print "RampTrigger5_Hit"
  WireRampOff
  RandomSoundBallBounceRamp activeball
End Sub
Sub RampTrigger5_UnHit
  WireRampOn True
End Sub


Sub RampTrigger6_Hit
  if DebugRampTriggers=True Then debug.print "RampTrigger6_Hit"
  RandomSoundBallBounceRamp activeball
End Sub


Sub RampTrigger7_Hit
  if DebugRampTriggers=True Then debug.print "RampTrigger7_Hit"
  RandomSoundBallBounceRamp activeball
End Sub


Sub RampTrigger8_Hit
  if DebugRampTriggers=True Then debug.print "RampTrigger8_Hit"
  WireRampOff
End Sub


Sub RampTrigger9_Hit
  if DebugRampTriggers=True Then debug.print "RampTrigger9_Hit"
  WireRampOff
End Sub

Sub RampTrigger10_Hit
  if DebugRampTriggers=True Then debug.print "RampTrigger10_Hit"
  WireRampOn False
End Sub



'******************************************************
'**** END RAMP ROLLING SFX
'******************************************************





'*****************************************
'   ZANI: Misc Animations
'*****************************************


' Flipper Animations

Sub LeftFlipper_Animate
  Dim lfa : lfa = LeftFlipper.CurrentAngle
  FlipperLSh.RotZ = LeftFlipper.CurrentAngle
  Dim BP : For Each BP in BP_LFlipper: BP.RotZ = lfa: Next
End Sub

Sub LeftFlipper1_Animate
  Dim lfa : lfa = LeftFlipper1.CurrentAngle
  FlipperLSh1.RotZ = LeftFlipper1.CurrentAngle
  Dim BP : For Each BP in BP_LUFlipper: BP.RotZ = lfa: Next
End Sub

Sub RightFlipper_Animate
  Dim rfa : rfa = RightFlipper.CurrentAngle
    FlipperRSh.RotZ = RightFlipper.CurrentAngle
  Dim BP : For Each BP in BP_RFlipper: BP.RotZ = rfa: Next
End Sub




' Switch Animations
Sub sw14_Animate
  Dim z : z = sw14.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw14 : BP.transz = z: Next
End Sub

Sub sw15_Animate
  Dim z : z = sw15.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw15 : BP.transz = z: Next
End Sub

Sub sw16_Animate
  Dim z : z = sw16.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw16 : BP.transz = z: Next
End Sub

Sub sw17_Animate
  Dim z : z = sw17.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw17 : BP.transz = z: Next
End Sub

Sub sw18_Animate
  Dim z : z = sw18.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw18 : BP.transz = z: Next
End Sub

Sub sw33_Animate
  Dim z : z = sw33.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw33 : BP.transz = z: Next
End Sub

Sub sw34_Animate
  Dim z : z = sw34.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw34 : BP.transz = z: Next
End Sub

Sub sw35_Animate
  Dim z : z = sw35.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw35 : BP.transz = z: Next
End Sub

Sub sw36_Animate
  Dim z : z = sw36.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw36 : BP.transz = z: Next
End Sub

Sub sw45_Animate
  Dim z : z = sw45.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw45 : BP.transz = z: Next
End Sub

Sub sw46_Animate
  Dim z : z = sw46.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw46 : BP.transz = z: Next
End Sub

Sub sw50_Animate
  Dim z : z = sw50.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw50 : BP.transz = z: Next
End Sub



' Gate Animations


Sub Gate1_Animate
    Dim a : a = Gate1.CurrentAngle
    Dim BP : For Each BP in BP_Gate1 : BP.rotx = a: Next
End Sub

Sub Gate2_Animate
    Dim a : a = Gate2.CurrentAngle
    Dim BP : For Each BP in BP_Gate2 : BP.rotx = a: Next
End Sub

Sub sw31_Animate
    Dim a : a = sw31.CurrentAngle
    Dim BP : For Each BP in BP_sw31 : BP.rotx = a: Next
End Sub

Sub sw32_Animate
    Dim a : a = sw32.CurrentAngle
    Dim BP : For Each BP in BP_sw32 : BP.rotx = a: Next
End Sub

Sub sw43_Animate
    Dim a : a = sw43.CurrentAngle
    Dim BP : For Each BP in BP_sw43 : BP.rotx = a: Next
End Sub

Sub sw44_Animate
    Dim a : a = sw44.CurrentAngle
    Dim BP : For Each BP in BP_sw44 : BP.rotx = a: Next
End Sub



' Spinner Animations
Sub sw41_Animate
  Dim spinangle: spinangle = sw41.currentangle
  Dim BP : For Each BP in BP_sw41 : BP.RotX = spinangle: Next
  For Each BP in BP_sw41a : BP.Roty = -10*dSin(spinangle): Next
End Sub




' Bumper Animations

Sub Bumper1_Animate
  Dim z, BP
  z = Bumper1.CurrentRingOffset
  For Each BP in BP_BR1 : BP.transz = z: Next
End Sub

Sub Bumper2_Animate
  Dim z, BP
  z = Bumper2.CurrentRingOffset
  For Each BP in BP_BR2 : BP.transz = z: Next
End Sub

Sub Bumper3_Animate
  Dim z, BP
  z = Bumper3.CurrentRingOffset
  For Each BP in BP_BR3 : BP.transz = z: Next
End Sub

Dim Bumpers : Bumpers = Array(Bumper1, Bumper2, Bumper3)

Sub AnimateBumperSkirts
  dim r, g, s, x, y, b, z
  ' Animate Bumper switch (experimental)
  For r = 0 To 2
    g = 10000.
    For s = 0 to UBound(gBOT)
      If gBOT(s).z < 30 Then ' Ball on playfield
        x = Bumpers(r).x - gBOT(s).x
        y = Bumpers(r).y - gBOT(s).y
        b = x * x + y * y
        If b < g Then g = b
      End If
    Next
    z = 4
    If g < 80 * 80 Then
      z = 1
    End If
    If r = 0 Then For Each x in BP_BS1: x.TransZ = z: Next
    If r = 1 Then For Each x in BP_BS2: x.TransZ = z: Next
    If r = 2 Then For Each x in BP_BS3: x.TransZ = z: Next
  Next
End Sub




' Standup Targets

Sub UpdateStandupTargets
  dim BP, ty

  ty = BM_sw19.transy
  For each BP in BP_sw19 : BP.transy = ty: Next

  ty = BM_sw21.transy
  For each BP in BP_sw21 : BP.transy = ty: Next

  ty = BM_sw22.transy
  For each BP in BP_sw22 : BP.transy = ty: Next

  ty = BM_sw23.transy
  For each BP in BP_sw23 : BP.transy = ty: Next

  ty = BM_sw24.transy
  For each BP in BP_sw24 : BP.transy = ty: Next

  ty = BM_sw30.transy
  For each BP in BP_sw30 : BP.transy = ty: Next

End Sub



' Drop Targets

Sub UpdateDropTargets
  dim BP, tz, rx, ry

  tz = BM_sw27.transz
  rx = BM_sw27.rotx
  ry = BM_sw27.roty
  For each BP in BP_sw27: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw28.transz
  rx = BM_sw28.rotx
  ry = BM_sw28.roty
  For each BP in BP_sw28: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw29.transz
  rx = BM_sw29.rotx
  ry = BM_sw29.roty
  For each BP in BP_sw29: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

End Sub






'*****************************************
' ZVRR: VR Room Stuff
'*****************************************


Sub EnableVRRoom
  Dim VRThings, el, BP
  If VRRoom <> 0 Then
    L58.intensity=0
    L59.intensity=0
    L60.intensity=0
    L61.intensity=0
    L62.intensity=0
    L63.intensity=0
    L64.intensity=0
    PinCab_Backglass.blenddisablelighting = 4
    For Each BP in BP_Lockdownbar: BP.visible = 0: Next
    For Each BP in BP_cabRails: BP.visible = 0: Next

    If VRRoom = 1 Then
      For Each VRThings in VR_Cab: VRThings.Visible = 1:Next
      For Each VRThings in VRMinimalRoom: VRThings.Visible = 0: Next
      For Each VRThings in VRDeluxeRoom: VRThings.Visible = 1: Next
    End If
    If VRRoom = 2 Then
      For Each VRThings in VR_Cab: VRThings.Visible = 1:Next
      For Each VRThings in VRMinimalRoom: VRThings.Visible = 1: Next
      For Each VRThings in VRDeluxeRoom: VRThings.Visible = 0: Next
    End If
    If VRRoom = 3 Then
      For Each VRThings in VR_Cab: VRThings.Visible = 0:Next
      For Each VRThings in VRMinimalRoom: VRThings.Visible = 0: Next
      For Each VRThings in VRDeluxeRoom: VRThings.Visible = 0: Next
      PinCab_Backbox.visible = 1
      PinCab_Backglass.visible = 1
    End If
  Else
    For Each VRThings in VR_Cab: VRThings.visible = 0:Next
    For Each VRThings in VRMinimalRoom: VRThings.visible = 0:Next
    For Each VRThings in VRDeluxeRoom: VRThings.Visible = 0: Next
    If DesktopMode then
      For Each el in DesktopLights: el.visible = 1: Next
      For Each BP in BP_Lockdownbar: BP.visible = 1: Next
      For Each BP in BP_cabRails: BP.visible = 1: Next
    Else
      For Each el in DesktopLights: el.visible = 0: Next
      For Each BP in BP_Lockdownbar: BP.visible = 0: Next
      For Each BP in BP_cabRails: BP.visible = 0: Next
    End If
  End if
End Sub



'Set Up VR Backglass Flashers
' this is for lining up the backglass flashers on top of a backglass image
Sub SetVRBackglass
  Dim obj
  For Each obj In BackGlass
    obj.x = obj.x
    obj.height = - obj.y + 320
    obj.y = -27 'adjusts the distance from the backglass towards the user
  Next
End Sub



Sub UpdateVRBackglass
  If VRRoom <> 0 Then
    f500.visible = L58.state
    f1million.visible = L59.state
    f1million25.visible = L60.state
    f1million5.visible = L61.state
    f1million5eb.visible = L62.state
    f2million.visible = L63.state
    f2million5.visible = L64.state
  End If
End Sub



'*****************************************************************************************************
' VR Plunger Animation
'*****************************************************************************************************

Sub TimerVRPlunger_Timer
  If Pincab_Plunger.Y < 90 then
      Pincab_Plunger.Y = Pincab_Plunger.Y + 5
  End If
End Sub

Sub TimerVRPlunger1_Timer
  Pincab_Plunger.Y = (5* Plunger.Position) - 20
End Sub



'*****************************************************************************************************
' VR BEER!
'*****************************************************************************************************
Sub BeerTimer_Timer()
  If VRRoom > 0 Then
    Randomize(21)
    VRBeerBubble1.z = VRBeerBubble1.z + Rnd(1)*0.5
    if VRBeerBubble1.z > -771 then VRBeerBubble1.z = -955
    VRBeerBubble2.z = VRBeerBubble2.z + Rnd(1)*1
    if VRBeerBubble2.z > -768 then VRBeerBubble2.z = -955
    VRBeerBubble3.z = VRBeerBubble3.z + Rnd(1)*1
    if VRBeerBubble3.z > -768 then VRBeerBubble3.z = -955
    VRBeerBubble4.z = VRBeerBubble4.z + Rnd(1)*0.75
    if VRBeerBubble4.z > -774 then VRBeerBubble4.z = -955
    VRBeerBubble5.z = VRBeerBubble5.z + Rnd(1)*1
    if VRBeerBubble5.z > -771 then VRBeerBubble5.z = -955
    VRBeerBubble6.z = VRBeerBubble6.z + Rnd(1)*1
    if VRBeerBubble6.z > -774 then VRBeerBubble6.z = -955
    VRBeerBubble7.z = VRBeerBubble7.z + Rnd(1)*0.8
    if VRBeerBubble7.z > -768 then VRBeerBubble7.z = -955
    VRBeerBubble8.z = VRBeerBubble8.z + Rnd(1)*1
    if VRBeerBubble8.z > -771 then VRBeerBubble8.z = -955
  End If
End Sub

'*****************************************************************************************************
' VR Lava Lamp
'*****************************************************************************************************
Dim Lspeed(7), Lbob(7), blobAng(7), blobRad(7), blobSiz(7), Blob, Bcnt   ' for VR Lava Lamp

' Lava Lamp code below.  Thank you STEELY!
Bcnt = 0

For Each Blob in Lspeed
  Lbob(Bcnt) = .2
  Lspeed(Bcnt) = Int((5 * Rnd) + 2) * .05
  Bcnt = Bcnt + 1
Next

Sub LavaTimer_Timer()
  If VRRoom > 0 Then
    Bcnt = 0
    For Each Blob in VRLava
      If Blob.TransZ <= VRLavaLampBase.Size_Z * 1.5 Then  'Change blob direction to up
        Lspeed(Bcnt) = Int((5 * Rnd) + 2) * .05   'travel speed
        blobAng(Bcnt) = Int((359 * Rnd) + 1)      'blob location/angle from center
        blobRad(Bcnt) = Int((40 * Rnd) + 10)      'blob distance from center, radius
        blobSiz(Bcnt) = Int((150 * Rnd) + 100)    'blob size
        Blob.Size_x = blobSiz(Bcnt):Blob.Size_y = blobSiz(Bcnt):Blob.Size_z = blobSiz(Bcnt)
        Blob.X = Round(Cos(blobAng(Bcnt)*0.0174533), 3) * blobRad(Bcnt) + VRLavaLampBase.X  'place blob
        Blob.Y = Round(Sin(blobAng(Bcnt)*0.0174533), 3) * blobRad(Bcnt) + VRLavaLampBase.Y
      End If
      If Blob.TransZ => VRLavaLampBase.Size_Z*5 Then    'Change blob direction to down
        blobAng(Bcnt) = Int((359 * Rnd) + 1)      'blob location/angle from center
        blobRad(Bcnt) = Int((40 * Rnd) + 10)      'blob distance from center,radius
        Blob.X = Round(Cos(blobAng(Bcnt)*0.0174533), 3) * blobRad(Bcnt) + VRLavaLampBase.X  'place blob
        Blob.Y = Round(Sin(blobAng(Bcnt)*0.0174533), 3) * blobRad(Bcnt) + VRLavaLampBase.Y
        Lspeed(Bcnt) = Int((8 * Rnd) + 5) * .05:Lspeed(Bcnt) = Lspeed(Bcnt) * -1        'travel speed
      End If

      ' Make blob wobble
      If Blob.Size_x > blobSiz(Bcnt) + blobSiz(Bcnt)*.15 or Blob.Size_x < blobSiz(Bcnt) - blobSiz(Bcnt)*.15  Then
        Lbob(Bcnt) = Lbob(Bcnt) * -1
      End If

      Blob.Size_x = Blob.Size_x + Lbob(Bcnt)
      Blob.Size_y = Blob.Size_y + Lbob(Bcnt)
      Blob.Size_z = Blob.Size_Z - Lbob(Bcnt) * .66
      Blob.TransZ = Blob.TransZ + Lspeed(Bcnt)    'Move blob
      Bcnt = Bcnt + 1
    Next
  End If
End Sub




'*****************************************
' ZCHL: Change Log
'*****************************************
'Based on works by 32Assassin
'000 - Bord, benji & oqqsan - merging WIP versions, mesh playfield, some physics work, subway and ball trough, insert lighting
'001 - Sixtoe - Added more physics objects, added VR room and cabinet including options, added fluppers latest flashers, cut holes in playfield, altered corkscrew mesh to line holes up better with lanes,
'002 - Uncle_Paulie - Added Alphanumeric DMD (and associated brightness adjustments), bonus backglass flashers for VR Rooms
'003 - Sixtoe - Aligned layers for inserts, started to set them up, ran out of time, 99% not hooked up.
'006 - Uncle_Paulie added backglass flahers.  Also fixed z fighting issue with ball shadows.  Changed Z value of 5 ballshadow primatives to 1.01. Animated flipper buttons.
'008 - tomate - all the new collidables imported from blender, new PF imported as reference
'009 - TastyWasps - nFozzy reset to newer physics routines, Fleep sounds
'010 - tomate - prubbers prim fixed
'011 - tomate - various fixes on collidable prims
'012 - TastyWasps - Deleted extra VPX objects on Main layer in prep for bake.
'013 - tomate - first 2k batch added; old prims erased; switches, plunger, flippers, bumpers and such settled in place, some other minor fixes
'014 - tomate - missing VLM materials added
'015 - apophis - Tons of table and script cleanups and updates. Added: Control light setup, GI, flashers, all animations, options menu, room brightness, reflections, Fleep and physics setup. Too much to name...
'016 - Sixtoe - Rebuilt entire physical table, widened holes on spiral prim ramp, adjusted right prim ramp, adjusted playfield mesh, stuff.
'017 - apophis - Made EI lightmaps opacity a function of the EI vertical position
'018 - tomate - pincab rails form VRcabinet layer set as invisible for now, new 2k batch added, drop targets textures updated (thanks redbone!), religned visual rubbers to match with physical rubbers, unwrapped textures for targets, playfield, clear plastics and sidewalls. Adedd some details in blender.
'019 - tomate - a real 2k batch added :P
'020 - apophis - Added Roth DT and ST code.
'021 - Sixtoe - Rebuilt physical ramps and objects to get things working.
'022 - Sixtoe - Rebuilt playfield mesh, converted both VUK's and Kicker Saucer to physical objects, added physical saucer deflector, redid and simplified all collidable primitive ramps, turned down the flippers to match bad cats, turned down the pop bumpers as they were a bit savage, probably some other stuff.
'023 - Benji - enabled depth mask on playfield bakemap
'024 - apophis - Added physical trust post. Updated flipper corrections to mimic work done on Fish Tales.
'025 - tomate - New 4k batch added, several fixes on blender side, unwrapped textures for ramps and Flippers
'026 - apophis - Updates for UseVPMModSol=2 : Removed InitPWM sub, Updated GI subs and FlashXX subs for new physics output, set all light fader models to None. Updated some VR backglass flasher calls.
'027 - tomate - Some fixes at blender side, new 4k batch added
'028 - tomate - Fixed unwrapped textures - right ramp sign still borked
'029 - apophis - Set EILightOffset to 0. Updated flipper triggers and scripts. Fixed diverter. Enabled playfield reflections.
'030 - apophis - reduce plunger strength for better skill shot, reduce ball reflectivity to reduce halo effect
'031 - Sixtoe - Made centre ramp harder, because lol.
'033 - Sixtoe - Changed BM_Layer0 and BM_Layer1 to "hide parts behind", assigned active material to all VLM.Lightmaps prims (which had no material assigned?)
'034 - apophis - Fixed timers causing stutters. Mech plunger adjust set back to 100.
'035 - apophis - Updated solenoid handling to not use core.vbs as it uses 'Execute' which can be very slow. Updated playfield dimensions and glass heights.
'036 - apophis - Uses new core.vbs solenoid callback that support FastFlips and avoids stutters
'037 - apophis - Revert SolCallback2 as it now has a better solution integrated in VPX
'038 - tomate - added new 1k batch with the real dimensions
'039 - Sixtoe - Started repositioning all table assets
'040 - Sixtoe - Finished repositioning all table assets (for now)
'041 - Sixtoe - Tweaked and simplified right side physical ramp
'042 - Sixtoe - Added left and right post hard mode table objects, needs scripting.
'043 - apophis - Added outpost difficulty option scripting. Still needs visual movables. Added rules info.
'044 - tomate - Several fixes at blender side, new 2k batch added, "hide parts behind" checked for BM_playfield
'045 - tomate - After a lot of trials an errors new 2k batch added, getting closer.
'046 - apophis - Tried to fix lower right vuk ramp. Added outpost difficulty mod (left post still broken).
'047 - apophis - Automated cab rail and lockdown bar visibility. Animated spinner wire. Added ramp rolling sound effects. Deleted old ramps.
'048 - tomate - Some fixes at blender side, unwrapped metal walls, 4k batch added
'049 - apophis - Added desktop DMD functionality. Increased live catch window to 24. Reduced slingshot force to 4.0. Fixed spinner wire animation. Fixed outpost option animation. Animated VR plunger.
'050 - Sixtoe - Tweaked "Zone 5" area, added new design physics deflector primitive, redid VR prims
'051 - apophis - Fixed VR plunger position. Set BP_Playfield to hid parts behind and added reflection probe. Enabled reflections for BM_Parts. Made VUKDef invisible. Made desktop cabrails and lockdown invibile in VR.
'052 - TastyWasps - Basti VR Room.  Small VR visual fixes.
'053 - apophis - Fix plunger animation
'054 - apophis - Updated physical bat dimensions per real measurements. Fixed post passing. Scaled up sw37wall a bit so ball sits completely inside hole. Triangulated faces on all collidable ramp prims. Flipper strength 2300. Updated ball roll sfx. Updated desktop alphanumeric colors (thanks HauntFreaks). Added motor sound for EI animation. Added shaker motor sfx and tweak options.
'055 - apophis - Made shaker effect more continuous. Center nudge strength set to 1. Flipper strength 2200. Oiled the spinner.
'056 - tomate - final touches at blender side, file optimized, new 4k batch imported.
'057 - Primetime5k - Added staged flipper support.
'058 - tomate - arrow inserts improved, bumper caps separated, bumper cap materials added, bumper caps refraction probe added
'RC1 - apophis - Added refraction option and EI mod to Tweak menu. Fixed VR room options. Set default game difficulty to 50. Disabled "hide parts behind" for shadow prims. Enabled "hide parts behind" for plastic ramp BMs.
'Release 1.0
'1.0.1 - apophis - Fixed keycode=19 issue. Added wall under plunger lane.
'1.0.2 - apophis - Made plunger material darker. Gate sw43 surface set to 60h. Adjusted Gate1 to prevent stuck balls (Zone 5). Enabled BM_Parts reflections. Deleted blueprint images. Nestmap fixes: removed shadow on california, udpated trustpost rubber color, made flipper rubbers less chalky.
'1.0.3 - Sixtoe - Fixed spiral ramp ball trap
'1.0.4 - apophis - Added left ramp end rubber. Dampener zvel correction. Increased size of playfield support subwall. Flipper elasticity at 0.88
'Release 1.1
'1.1.1 - apophis - Added GateBlocker and logic to prevent stuck balls at shot 5. Removed shadow on back of spinner bakemap (not lightmaps). Updated DisableStaticPreRendering functionality to be compatible with VPX 10.8.1 API

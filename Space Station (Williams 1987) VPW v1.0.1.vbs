'      _____ ____  ___   ____________   ______________  ______________  _   __
'     / ___// __ \/   | / ____/ ____/  / ___/_  __/   |/_  __/  _/ __ \/ | / /
'     \__ \/ /_/ / /| |/ /   / __/     \__ \ / / / /| | / /  / // / / /  |/ /
'    ___/ / ____/ ___ / /___/ /___    ___/ // / / ___ |/ / _/ // /_/ / /|  /
'   /____/_/   /_/  |_\____/_____/   /____//_/ /_/  |_/_/ /___/\____/_/ |_/
'
'    Williams 1987
'    https://www.ipdb.org/machine.cgi?id=2261
'
'
' VPW CREW
' ---------
' Tomate - Commander. Blender modeling.
' Sixtoe - Pilot. VPX table build.
' Apophis - Flight Engineer. Fixing parts.
' MCarter - Payload Specialist. GI wiring.
' DGrimmReaper - EVA Specialist. Added Rawd's VR room.
' Rawd - Special Specialist. Original VR Room.
' bthlonewolf - Mission Specialist. Ball tinting.
' RobbyKingPin - Crew Medical Officer. Fleep sounds.
' nFozzy - Flight Controller. Space Station Toy Code and previous table version.
' iaakki - Machine specialist, side by side testing with a real machine.
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
' ZTOY: Toys
' ZFLP: Flippers
' ZSLG: Slingshot Animations
' ZSSC: Slingshot Corrections
' ZSWI: Switches
' ZVUK: VUKs and Kickers
'   ZBRL: Ball Rolling and Drop Sounds
'   ZRRL: Ramp Rolling Sound Effects
'   ZFLE: Fleep Mechanical Sounds
' ZPHY: General Advice on Physics
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
' ZGIC: MB GI Colors
' ZSHA: Ambient ball shadows
' ZLED: Display LEDs
'
' These sections have not been added yet
'   ZVRR: VR Room / VR Cabinet
'
'*********************************************************************************************************************************

Option Explicit
Randomize
SetLocale 1033      'Forces VBS to use english to stop crashes.

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0


' VLM  Arrays - Start
' Arrays per baked part
Dim BP_BR1: BP_BR1=Array(BM_BR1, LM_Bumpers_BR1, LM_FL_f27_BR1, LM_GIN_BR1, LM_GIG_BR1)
Dim BP_BR2: BP_BR2=Array(BM_BR2, LM_Bumpers_BR2, LM_FL_f15_BR2, LM_FL_f26_BR2, LM_FL_f27_BR2, LM_FL_f28_BR2, LM_FL_f29_BR2, LM_GIN_BR2, LM_GIG_BR2)
Dim BP_BR3: BP_BR3=Array(BM_BR3, LM_Bumpers_BR3, LM_FL_f26_BR3, LM_FL_f27_BR3, LM_GIN_BR3, LM_GIG_BR3, LM_IN_l27_BR3)
Dim BP_BS1: BP_BS1=Array(BM_BS1, LM_Bumpers_BS1, LM_FL_f27_BS1, LM_FL_f28_BS1, LM_GIN_BS1, LM_GIG_BS1)
Dim BP_BS2: BP_BS2=Array(BM_BS2, LM_Bumpers_BS2, LM_FL_f26_BS2, LM_FL_f27_BS2, LM_FL_f28_BS2, LM_GIN_BS2, LM_GIG_BS2)
Dim BP_BS3: BP_BS3=Array(BM_BS3, LM_Bumpers_BS3, LM_FL_f15_BS3, LM_FL_f26_BS3, LM_FL_f27_BS3, LM_FL_f28_BS3, LM_GIN_BS3, LM_GIG_BS3, LM_IN_l27_BS3)
Dim BP_Gate1: BP_Gate1=Array(BM_Gate1, LM_FL_f26_Gate1, LM_GIN_Gate1, LM_GIG_Gate1)
Dim BP_Gate2_001: BP_Gate2_001=Array(BM_Gate2_001, LM_FL_f27_Gate2_001, LM_FL_f28_Gate2_001, LM_GIN_Gate2_001, LM_GIG_Gate2_001)
Dim BP_Gate3: BP_Gate3=Array(BM_Gate3, LM_GIN_Gate3)
Dim BP_Gate4: BP_Gate4=Array(BM_Gate4, LM_FL_f28_Gate4, LM_GIN_Gate4, LM_GIG_Gate4)
Dim BP_Gate_Orbit: BP_Gate_Orbit=Array(BM_Gate_Orbit, LM_FL_f28_Gate_Orbit, LM_GIN_Gate_Orbit, LM_GIG_Gate_Orbit)
Dim BP_LEMK: BP_LEMK=Array(BM_LEMK, LM_GIGS_GI_G_001_LEMK, LM_GIGS_GI_G_002_LEMK, LM_FL_f25_LEMK, LM_FL_f26_LEMK, LM_GINS_gi01_LEMK, LM_GINS_gi02_LEMK)
Dim BP_LSling0_001: BP_LSling0_001=Array(BM_LSling0_001, LM_GIGS_GI_G_001_LSling0_001, LM_GIGS_GI_G_002_LSling0_001, LM_FL_f25_LSling0_001, LM_FL_f26_LSling0_001, LM_FL_f27_LSling0_001, LM_GINS_gi01_LSling0_001, LM_GINS_gi02_LSling0_001, LM_IN_l17_LSling0_001)
Dim BP_LSling1: BP_LSling1=Array(BM_LSling1, LM_GIGS_GI_G_001_LSling1, LM_GIGS_GI_G_002_LSling1, LM_FL_f25_LSling1, LM_FL_f26_LSling1, LM_FL_f27_LSling1, LM_GINS_gi01_LSling1, LM_GINS_gi02_LSling1, LM_IN_l17_LSling1)
Dim BP_LSling2: BP_LSling2=Array(BM_LSling2, LM_GIGS_GI_G_001_LSling2, LM_GIGS_GI_G_002_LSling2, LM_FL_f25_LSling2, LM_FL_f26_LSling2, LM_FL_f27_LSling2, LM_GINS_gi01_LSling2, LM_GINS_gi02_LSling2, LM_IN_l17_LSling2)
Dim BP_Layer0: BP_Layer0=Array(BM_Layer0, LM_GIGS_GI_G_001_Layer0, LM_GIGS_GI_G_002_Layer0, LM_GIGS_GI_G_022_Layer0, LM_GIGS_GI_G_023_Layer0, LM_GIGS_GI_G_024_Layer0, LM_FL_f15_Layer0, LM_FL_f25_Layer0, LM_FL_f26_Layer0, LM_FL_f27_Layer0, LM_FL_f28_Layer0, LM_GINS_gi01_Layer0, LM_GINS_gi02_Layer0, LM_GIN_Layer0, LM_GINS_gi24_Layer0, LM_GINS_gi25_Layer0, LM_GINS_gi26_Layer0, LM_GIG_Layer0, LM_IN_l1_Layer0, LM_IN_l10_Layer0, LM_IN_l11_Layer0, LM_IN_l12_Layer0, LM_IN_l13_Layer0, LM_IN_l17_Layer0, LM_IN_l18_Layer0, LM_IN_l19_Layer0, LM_IN_l2_Layer0, LM_IN_l20_Layer0, LM_IN_l21_Layer0, LM_IN_l22_Layer0, LM_IN_l23_Layer0, LM_IN_l24_Layer0, LM_IN_l25_Layer0, LM_IN_l26_Layer0, LM_IN_l27_Layer0, LM_IN_l28_Layer0, LM_IN_l29_Layer0, LM_IN_l3_Layer0, LM_IN_l30_Layer0, LM_IN_l31_Layer0, LM_IN_l32_Layer0, LM_IN_l34_Layer0, LM_IN_l35_Layer0, LM_IN_l36_Layer0, LM_IN_l4_Layer0, LM_IN_l42_Layer0, LM_IN_l43_Layer0, LM_IN_l44_Layer0, LM_IN_l45_Layer0, LM_IN_l5_Layer0, LM_IN_l50_Layer0, LM_IN_l51_Layer0, LM_IN_l52_Layer0, _
  LM_IN_l53_Layer0, LM_IN_l54_Layer0, LM_IN_l55_Layer0, LM_IN_l56_Layer0, LM_IN_l58_Layer0, LM_IN_l59_Layer0, LM_IN_l6_Layer0, LM_IN_l60_Layer0, LM_IN_l61_Layer0, LM_IN_l62_Layer0, LM_IN_l63_Layer0, LM_IN_l64_Layer0, LM_IN_l7_Layer0, LM_IN_l8_Layer0, LM_IN_l9_Layer0)
Dim BP_Layer1: BP_Layer1=Array(BM_Layer1, LM_GIGS_GI_G_001_Layer1, LM_GIGS_GI_G_002_Layer1, LM_GIGS_GI_G_022_Layer1, LM_GIGS_GI_G_023_Layer1, LM_GIGS_GI_G_024_Layer1, LM_Bumpers_Layer1, LM_FL_f15_Layer1, LM_FL_f25_Layer1, LM_FL_f26_Layer1, LM_FL_f27_Layer1, LM_FL_f28_Layer1, LM_GINS_gi01_Layer1, LM_GINS_gi02_Layer1, LM_GIN_Layer1, LM_GINS_gi24_Layer1, LM_GINS_gi25_Layer1, LM_GINS_gi26_Layer1, LM_GIG_Layer1, LM_IN_l17_Layer1, LM_IN_l2_Layer1, LM_IN_l32_Layer1, LM_IN_l7_Layer1)
Dim BP_Layer2: BP_Layer2=Array(BM_Layer2, LM_Bumpers_Layer2, LM_FL_f15_Layer2, LM_FL_f27_Layer2, LM_FL_f28_Layer2, LM_FL_f29_Layer2, LM_GIN_Layer2, LM_GIG_Layer2, LM_IN_l40_Layer2)
Dim BP_Layer3: BP_Layer3=Array(BM_Layer3, LM_Bumpers_Layer3, LM_FL_f15_Layer3, LM_FL_f27_Layer3, LM_FL_f28_Layer3, LM_FL_f29_Layer3, LM_GIN_Layer3, LM_GIG_Layer3, LM_IN_l14_Layer3, LM_IN_l15_Layer3, LM_IN_l16_Layer3, LM_IN_l28_Layer3, LM_IN_l36_Layer3, LM_IN_l38_Layer3, LM_IN_l39_Layer3, LM_IN_l40_Layer3, LM_IN_l7_Layer3)
Dim BP_LayerCaps: BP_LayerCaps=Array(BM_LayerCaps, LM_Bumpers_LayerCaps, LM_FL_f15_LayerCaps, LM_FL_f27_LayerCaps, LM_FL_f28_LayerCaps, LM_FL_f29_LayerCaps, LM_GIN_LayerCaps, LM_GIG_LayerCaps)
Dim BP_LeftFlipper: BP_LeftFlipper=Array(BM_LeftFlipper, LM_GIGS_GI_G_001_LeftFlipper, LM_FL_f25_LeftFlipper, LM_GINS_gi01_LeftFlipper)
Dim BP_Playfield: BP_Playfield=Array(BM_Playfield, LM_GIGS_GI_G_001_Playfield, LM_GIGS_GI_G_002_Playfield, LM_GIGS_GI_G_020_Playfield, LM_GIGS_GI_G_022_Playfield, LM_GIGS_GI_G_023_Playfield, LM_GIGS_GI_G_024_Playfield, LM_Bumpers_Playfield, LM_FL_f15_Playfield, LM_FL_f25_Playfield, LM_FL_f26_Playfield, LM_FL_f27_Playfield, LM_FL_f28_Playfield, LM_FL_f29_Playfield, LM_GINS_gi01_Playfield, LM_GINS_gi02_Playfield, LM_GIN_Playfield, LM_GINS_gi24_Playfield, LM_GINS_gi25_Playfield, LM_GINS_gi26_Playfield, LM_GIG_Playfield, LM_IN_l1_Playfield, LM_IN_l10_Playfield, LM_IN_l11_Playfield, LM_IN_l12_Playfield, LM_IN_l13_Playfield, LM_IN_l14_Playfield, LM_IN_l15_Playfield, LM_IN_l16_Playfield, LM_IN_l17_Playfield, LM_IN_l18_Playfield, LM_IN_l19_Playfield, LM_IN_l2_Playfield, LM_IN_l20_Playfield, LM_IN_l21_Playfield, LM_IN_l22_Playfield, LM_IN_l23_Playfield, LM_IN_l24_Playfield, LM_IN_l25_Playfield, LM_IN_l26_Playfield, LM_IN_l27_Playfield, LM_IN_l28_Playfield, LM_IN_l29_Playfield, LM_IN_l3_Playfield, LM_IN_l30_Playfield, _
  LM_IN_l31_Playfield, LM_IN_l32_Playfield, LM_IN_l33_Playfield, LM_IN_l34_Playfield, LM_IN_l35_Playfield, LM_IN_l36_Playfield, LM_IN_l38_Playfield, LM_IN_l39_Playfield, LM_IN_l4_Playfield, LM_IN_l40_Playfield, LM_IN_l42_Playfield, LM_IN_l43_Playfield, LM_IN_l44_Playfield, LM_IN_l45_Playfield, LM_IN_l5_Playfield, LM_IN_l50_Playfield, LM_IN_l51_Playfield, LM_IN_l52_Playfield, LM_IN_l53_Playfield, LM_IN_l54_Playfield, LM_IN_l55_Playfield, LM_IN_l56_Playfield, LM_IN_l58_Playfield, LM_IN_l59_Playfield, LM_IN_l6_Playfield, LM_IN_l60_Playfield, LM_IN_l61_Playfield, LM_IN_l62_Playfield, LM_IN_l63_Playfield, LM_IN_l64_Playfield, LM_IN_l7_Playfield, LM_IN_l8_Playfield, LM_IN_l9_Playfield)
Dim BP_REMK: BP_REMK=Array(BM_REMK, LM_GIGS_GI_G_022_REMK, LM_GIGS_GI_G_023_REMK, LM_GIGS_GI_G_024_REMK, LM_FL_f25_REMK, LM_FL_f27_REMK, LM_GINS_gi25_REMK, LM_GINS_gi26_REMK)
Dim BP_RSling0_001: BP_RSling0_001=Array(BM_RSling0_001, LM_GIGS_GI_G_022_RSling0_001, LM_GIGS_GI_G_023_RSling0_001, LM_GIGS_GI_G_024_RSling0_001, LM_FL_f25_RSling0_001, LM_FL_f26_RSling0_001, LM_FL_f27_RSling0_001, LM_GIN_RSling0_001, LM_GINS_gi24_RSling0_001, LM_GINS_gi25_RSling0_001, LM_GINS_gi26_RSling0_001, LM_GIG_RSling0_001, LM_IN_l32_RSling0_001)
Dim BP_RSling1: BP_RSling1=Array(BM_RSling1, LM_GIGS_GI_G_022_RSling1, LM_GIGS_GI_G_023_RSling1, LM_GIGS_GI_G_024_RSling1, LM_FL_f25_RSling1, LM_FL_f27_RSling1, LM_GIN_RSling1, LM_GINS_gi24_RSling1, LM_GINS_gi25_RSling1, LM_GINS_gi26_RSling1, LM_GIG_RSling1, LM_IN_l32_RSling1)
Dim BP_RSling2: BP_RSling2=Array(BM_RSling2, LM_GIGS_GI_G_022_RSling2, LM_GIGS_GI_G_023_RSling2, LM_GIGS_GI_G_024_RSling2, LM_FL_f25_RSling2, LM_FL_f26_RSling2, LM_FL_f27_RSling2, LM_GIN_RSling2, LM_GINS_gi24_RSling2, LM_GINS_gi25_RSling2, LM_GINS_gi26_RSling2)
Dim BP_RightFlipper: BP_RightFlipper=Array(BM_RightFlipper, LM_GIGS_GI_G_023_RightFlipper, LM_GIGS_GI_G_024_RightFlipper, LM_FL_f25_RightFlipper, LM_FL_f27_RightFlipper, LM_GINS_gi26_RightFlipper)
Dim BP_SS_01: BP_SS_01=Array(BM_SS_01, LM_Bumpers_SS_01, LM_FL_f15_SS_01, LM_FL_f27_SS_01, LM_FL_f28_SS_01, LM_FL_f29_SS_01, LM_GIN_SS_01, LM_GIG_SS_01, LM_IN_l28_SS_01)
Dim BP_SS_02: BP_SS_02=Array(BM_SS_02, LM_Bumpers_SS_02, LM_FL_f15_SS_02, LM_FL_f27_SS_02, LM_FL_f28_SS_02, LM_FL_f29_SS_02, LM_GIN_SS_02, LM_GIG_SS_02)
Dim BP_SS_03: BP_SS_03=Array(BM_SS_03, LM_Bumpers_SS_03, LM_FL_f15_SS_03, LM_FL_f27_SS_03, LM_FL_f28_SS_03, LM_FL_f29_SS_03, LM_GIN_SS_03, LM_GIG_SS_03, LM_IN_l28_SS_03)
Dim BP_SS_04: BP_SS_04=Array(BM_SS_04, LM_Bumpers_SS_04, LM_FL_f15_SS_04, LM_FL_f27_SS_04, LM_FL_f28_SS_04, LM_FL_f29_SS_04, LM_GIN_SS_04, LM_GIG_SS_04)
Dim BP_diverter01: BP_diverter01=Array(BM_diverter01, LM_Bumpers_diverter01, LM_FL_f15_diverter01, LM_FL_f27_diverter01, LM_FL_f28_diverter01, LM_FL_f29_diverter01, LM_GIN_diverter01, LM_GIG_diverter01, LM_IN_l28_diverter01, LM_IN_l5_diverter01)
Dim BP_diverter02: BP_diverter02=Array(BM_diverter02, LM_FL_f15_diverter02, LM_FL_f27_diverter02, LM_FL_f28_diverter02, LM_FL_f29_diverter02, LM_GIN_diverter02, LM_GIG_diverter02, LM_IN_l28_diverter02, LM_IN_l5_diverter02)
Dim BP_diverter03: BP_diverter03=Array(BM_diverter03, LM_FL_f15_diverter03, LM_FL_f27_diverter03, LM_FL_f28_diverter03, LM_FL_f29_diverter03, LM_GIN_diverter03, LM_GIG_diverter03, LM_IN_l5_diverter03)
Dim BP_diverter04: BP_diverter04=Array(BM_diverter04, LM_FL_f15_diverter04, LM_FL_f27_diverter04, LM_FL_f28_diverter04, LM_FL_f29_diverter04, LM_GIN_diverter04, LM_GIG_diverter04, LM_IN_l5_diverter04)
Dim BP_domes: BP_domes=Array(BM_domes, LM_GIGS_GI_G_020_domes, LM_GIGS_GI_G_022_domes, LM_GIGS_GI_G_023_domes, LM_GIGS_GI_G_024_domes, LM_Bumpers_domes, LM_FL_f15_domes, LM_FL_f25_domes, LM_FL_f26_domes, LM_FL_f27_domes, LM_FL_f28_domes, LM_FL_f29_domes, LM_GIN_domes, LM_GINS_gi24_domes, LM_GINS_gi25_domes, LM_GINS_gi26_domes, LM_GIG_domes, LM_IN_l1_domes, LM_IN_l12_domes, LM_IN_l13_domes, LM_IN_l18_domes, LM_IN_l19_domes, LM_IN_l20_domes, LM_IN_l21_domes, LM_IN_l22_domes, LM_IN_l23_domes, LM_IN_l24_domes, LM_IN_l28_domes, LM_IN_l29_domes, LM_IN_l30_domes, LM_IN_l31_domes, LM_IN_l32_domes, LM_IN_l35_domes)
Dim BP_parts: BP_parts=Array(BM_parts, LM_GIGS_GI_G_001_parts, LM_GIGS_GI_G_002_parts, LM_GIGS_GI_G_020_parts, LM_GIGS_GI_G_022_parts, LM_GIGS_GI_G_023_parts, LM_GIGS_GI_G_024_parts, LM_Bumpers_parts, LM_FL_f15_parts, LM_FL_f25_parts, LM_FL_f26_parts, LM_FL_f27_parts, LM_FL_f28_parts, LM_FL_f29_parts, LM_GINS_gi01_parts, LM_GINS_gi02_parts, LM_GIN_parts, LM_GINS_gi24_parts, LM_GINS_gi25_parts, LM_GINS_gi26_parts, LM_GIG_parts, LM_IN_l10_parts, LM_IN_l12_parts, LM_IN_l14_parts, LM_IN_l15_parts, LM_IN_l16_parts, LM_IN_l17_parts, LM_IN_l18_parts, LM_IN_l20_parts, LM_IN_l21_parts, LM_IN_l22_parts, LM_IN_l23_parts, LM_IN_l24_parts, LM_IN_l25_parts, LM_IN_l26_parts, LM_IN_l27_parts, LM_IN_l28_parts, LM_IN_l29_parts, LM_IN_l30_parts, LM_IN_l31_parts, LM_IN_l32_parts, LM_IN_l33_parts, LM_IN_l34_parts, LM_IN_l35_parts, LM_IN_l36_parts, LM_IN_l38_parts, LM_IN_l39_parts, LM_IN_l40_parts, LM_IN_l5_parts, LM_IN_l50_parts, LM_IN_l51_parts, LM_IN_l58_parts, LM_IN_l59_parts, LM_IN_l6_parts, LM_IN_l7_parts, LM_IN_l8_parts, _
  LM_IN_l9_parts)
Dim BP_sw17: BP_sw17=Array(BM_sw17, LM_GIGS_GI_G_001_sw17, LM_GIGS_GI_G_002_sw17, LM_GINS_gi01_sw17, LM_GINS_gi02_sw17)
Dim BP_sw18: BP_sw18=Array(BM_sw18, LM_GIGS_GI_G_001_sw18, LM_FL_f25_sw18, LM_FL_f26_sw18, LM_GIN_sw18, LM_GIG_sw18, LM_IN_l18_sw18, LM_IN_l19_sw18)
Dim BP_sw19: BP_sw19=Array(BM_sw19, LM_GIGS_GI_G_001_sw19, LM_FL_f25_sw19, LM_FL_f26_sw19, LM_GIN_sw19, LM_GIG_sw19, LM_IN_l19_sw19, LM_IN_l20_sw19)
Dim BP_sw20: BP_sw20=Array(BM_sw20, LM_FL_f25_sw20, LM_FL_f26_sw20, LM_GIN_sw20, LM_GIG_sw20, LM_IN_l20_sw20, LM_IN_l21_sw20)
Dim BP_sw21: BP_sw21=Array(BM_sw21, LM_FL_f25_sw21, LM_FL_f26_sw21, LM_GIN_sw21, LM_GIG_sw21, LM_IN_l21_sw21, LM_IN_l22_sw21)
Dim BP_sw22: BP_sw22=Array(BM_sw22, LM_FL_f26_sw22, LM_GIN_sw22, LM_GIG_sw22, LM_IN_l22_sw22, LM_IN_l23_sw22)
Dim BP_sw23: BP_sw23=Array(BM_sw23, LM_FL_f26_sw23, LM_GIN_sw23, LM_GIG_sw23, LM_IN_l23_sw23, LM_IN_l24_sw23)
Dim BP_sw24: BP_sw24=Array(BM_sw24, LM_FL_f26_sw24, LM_GIN_sw24, LM_GIG_sw24, LM_IN_l24_sw24)
Dim BP_sw25: BP_sw25=Array(BM_sw25, LM_Bumpers_sw25, LM_FL_f26_sw25, LM_FL_f27_sw25, LM_GIN_sw25, LM_GIG_sw25, LM_IN_l25_sw25)
Dim BP_sw26: BP_sw26=Array(BM_sw26, LM_Bumpers_sw26, LM_FL_f26_sw26, LM_FL_f27_sw26, LM_GIN_sw26, LM_GIG_sw26, LM_IN_l25_sw26, LM_IN_l26_sw26)
Dim BP_sw27: BP_sw27=Array(BM_sw27, LM_Bumpers_sw27, LM_FL_f26_sw27, LM_FL_f27_sw27, LM_GIN_sw27, LM_GIG_sw27, LM_IN_l26_sw27, LM_IN_l27_sw27)
Dim BP_sw28: BP_sw28=Array(BM_sw28, LM_FL_f27_sw28, LM_FL_f28_sw28, LM_GIN_sw28, LM_GIG_sw28, LM_IN_l28_sw28)
Dim BP_sw29: BP_sw29=Array(BM_sw29, LM_FL_f27_sw29, LM_FL_f28_sw29, LM_GIN_sw29, LM_GIG_sw29, LM_IN_l28_sw29, LM_IN_l29_sw29)
Dim BP_sw30: BP_sw30=Array(BM_sw30, LM_FL_f27_sw30, LM_FL_f28_sw30, LM_GIN_sw30, LM_GIG_sw30, LM_IN_l29_sw30, LM_IN_l30_sw30)
Dim BP_sw31: BP_sw31=Array(BM_sw31, LM_FL_f27_sw31, LM_FL_f28_sw31, LM_GIN_sw31, LM_GIG_sw31, LM_IN_l30_sw31, LM_IN_l31_sw31)
Dim BP_sw32: BP_sw32=Array(BM_sw32, LM_GIGS_GI_G_022_sw32, LM_GIGS_GI_G_023_sw32, LM_GIGS_GI_G_024_sw32, LM_GINS_gi24_sw32, LM_GINS_gi25_sw32, LM_GINS_gi26_sw32)
Dim BP_sw33: BP_sw33=Array(BM_sw33, LM_FL_f25_sw33, LM_FL_f26_sw33, LM_FL_f27_sw33, LM_GIN_sw33, LM_GINS_gi24_sw33, LM_GIG_sw33, LM_IN_l33_sw33, LM_IN_l9_sw33)
Dim BP_sw35: BP_sw35=Array(BM_sw35, LM_FL_f25_sw35, LM_GIN_sw35, LM_GIG_sw35, LM_IN_l9_sw35)
Dim BP_sw38: BP_sw38=Array(BM_sw38, LM_GIN_sw38, LM_GIG_sw38)
Dim BP_sw39: BP_sw39=Array(BM_sw39, LM_GIN_sw39, LM_GIG_sw39)
Dim BP_sw40: BP_sw40=Array(BM_sw40, LM_GIN_sw40, LM_GIG_sw40)
Dim BP_sw43: BP_sw43=Array(BM_sw43)
Dim BP_sw45: BP_sw45=Array(BM_sw45, LM_FL_f27_sw45, LM_GIN_sw45, LM_GIG_sw45)
Dim BP_sw57: BP_sw57=Array(BM_sw57, LM_FL_f27_sw57, LM_FL_f28_sw57, LM_GIN_sw57, LM_GIG_sw57, LM_IN_l6_sw57)
Dim BP_sw58: BP_sw58=Array(BM_sw58, LM_FL_f27_sw58, LM_FL_f28_sw58, LM_GIN_sw58, LM_GIG_sw58, LM_IN_l6_sw58)
Dim BP_sw59: BP_sw59=Array(BM_sw59, LM_FL_f27_sw59, LM_GIN_sw59, LM_GIG_sw59, LM_IN_l5_sw59)
Dim BP_underPF: BP_underPF=Array(BM_underPF, LM_GIGS_GI_G_001_underPF, LM_GIGS_GI_G_002_underPF, LM_GIGS_GI_G_020_underPF, LM_GIGS_GI_G_022_underPF, LM_GIGS_GI_G_023_underPF, LM_GIGS_GI_G_024_underPF, LM_Bumpers_underPF, LM_FL_f15_underPF, LM_FL_f25_underPF, LM_FL_f26_underPF, LM_FL_f27_underPF, LM_FL_f28_underPF, LM_FL_f29_underPF, LM_GINS_gi01_underPF, LM_GINS_gi02_underPF, LM_GIN_underPF, LM_GINS_gi24_underPF, LM_GINS_gi25_underPF, LM_GINS_gi26_underPF, LM_GIG_underPF, LM_IN_l1_underPF, LM_IN_l10_underPF, LM_IN_l11_underPF, LM_IN_l12_underPF, LM_IN_l13_underPF, LM_IN_l14_underPF, LM_IN_l15_underPF, LM_IN_l17_underPF, LM_IN_l18_underPF, LM_IN_l19_underPF, LM_IN_l2_underPF, LM_IN_l20_underPF, LM_IN_l21_underPF, LM_IN_l22_underPF, LM_IN_l23_underPF, LM_IN_l24_underPF, LM_IN_l25_underPF, LM_IN_l26_underPF, LM_IN_l27_underPF, LM_IN_l28_underPF, LM_IN_l29_underPF, LM_IN_l3_underPF, LM_IN_l30_underPF, LM_IN_l31_underPF, LM_IN_l32_underPF, LM_IN_l33_underPF, LM_IN_l34_underPF, LM_IN_l35_underPF, LM_IN_l36_underPF, _
  LM_IN_l38_underPF, LM_IN_l39_underPF, LM_IN_l4_underPF, LM_IN_l40_underPF, LM_IN_l42_underPF, LM_IN_l43_underPF, LM_IN_l44_underPF, LM_IN_l45_underPF, LM_IN_l5_underPF, LM_IN_l50_underPF, LM_IN_l51_underPF, LM_IN_l52_underPF, LM_IN_l53_underPF, LM_IN_l54_underPF, LM_IN_l55_underPF, LM_IN_l56_underPF, LM_IN_l58_underPF, LM_IN_l59_underPF, LM_IN_l6_underPF, LM_IN_l60_underPF, LM_IN_l61_underPF, LM_IN_l62_underPF, LM_IN_l63_underPF, LM_IN_l64_underPF, LM_IN_l7_underPF, LM_IN_l8_underPF, LM_IN_l9_underPF)
' Arrays per lighting scenario
Dim BL_Bumpers: BL_Bumpers=Array(LM_Bumpers_BR1, LM_Bumpers_BR2, LM_Bumpers_BR3, LM_Bumpers_BS1, LM_Bumpers_BS2, LM_Bumpers_BS3, LM_Bumpers_Layer1, LM_Bumpers_Layer2, LM_Bumpers_Layer3, LM_Bumpers_LayerCaps, LM_Bumpers_Playfield, LM_Bumpers_SS_01, LM_Bumpers_SS_02, LM_Bumpers_SS_03, LM_Bumpers_SS_04, LM_Bumpers_diverter01, LM_Bumpers_domes, LM_Bumpers_parts, LM_Bumpers_sw25, LM_Bumpers_sw26, LM_Bumpers_sw27, LM_Bumpers_underPF)
Dim BL_FL_f15: BL_FL_f15=Array(LM_FL_f15_BR2, LM_FL_f15_BS3, LM_FL_f15_Layer0, LM_FL_f15_Layer1, LM_FL_f15_Layer2, LM_FL_f15_Layer3, LM_FL_f15_LayerCaps, LM_FL_f15_Playfield, LM_FL_f15_SS_01, LM_FL_f15_SS_02, LM_FL_f15_SS_03, LM_FL_f15_SS_04, LM_FL_f15_diverter01, LM_FL_f15_diverter02, LM_FL_f15_diverter03, LM_FL_f15_diverter04, LM_FL_f15_domes, LM_FL_f15_parts, LM_FL_f15_underPF)
Dim BL_FL_f25: BL_FL_f25=Array(LM_FL_f25_LEMK, LM_FL_f25_LSling0_001, LM_FL_f25_LSling1, LM_FL_f25_LSling2, LM_FL_f25_Layer0, LM_FL_f25_Layer1, LM_FL_f25_LeftFlipper, LM_FL_f25_Playfield, LM_FL_f25_REMK, LM_FL_f25_RSling0_001, LM_FL_f25_RSling1, LM_FL_f25_RSling2, LM_FL_f25_RightFlipper, LM_FL_f25_domes, LM_FL_f25_parts, LM_FL_f25_sw18, LM_FL_f25_sw19, LM_FL_f25_sw20, LM_FL_f25_sw21, LM_FL_f25_sw33, LM_FL_f25_sw35, LM_FL_f25_underPF)
Dim BL_FL_f26: BL_FL_f26=Array(LM_FL_f26_BR2, LM_FL_f26_BR3, LM_FL_f26_BS2, LM_FL_f26_BS3, LM_FL_f26_Gate1, LM_FL_f26_LEMK, LM_FL_f26_LSling0_001, LM_FL_f26_LSling1, LM_FL_f26_LSling2, LM_FL_f26_Layer0, LM_FL_f26_Layer1, LM_FL_f26_Playfield, LM_FL_f26_RSling0_001, LM_FL_f26_RSling2, LM_FL_f26_domes, LM_FL_f26_parts, LM_FL_f26_sw18, LM_FL_f26_sw19, LM_FL_f26_sw20, LM_FL_f26_sw21, LM_FL_f26_sw22, LM_FL_f26_sw23, LM_FL_f26_sw24, LM_FL_f26_sw25, LM_FL_f26_sw26, LM_FL_f26_sw27, LM_FL_f26_sw33, LM_FL_f26_underPF)
Dim BL_FL_f27: BL_FL_f27=Array(LM_FL_f27_BR1, LM_FL_f27_BR2, LM_FL_f27_BR3, LM_FL_f27_BS1, LM_FL_f27_BS2, LM_FL_f27_BS3, LM_FL_f27_Gate2_001, LM_FL_f27_LSling0_001, LM_FL_f27_LSling1, LM_FL_f27_LSling2, LM_FL_f27_Layer0, LM_FL_f27_Layer1, LM_FL_f27_Layer2, LM_FL_f27_Layer3, LM_FL_f27_LayerCaps, LM_FL_f27_Playfield, LM_FL_f27_REMK, LM_FL_f27_RSling0_001, LM_FL_f27_RSling1, LM_FL_f27_RSling2, LM_FL_f27_RightFlipper, LM_FL_f27_SS_01, LM_FL_f27_SS_02, LM_FL_f27_SS_03, LM_FL_f27_SS_04, LM_FL_f27_diverter01, LM_FL_f27_diverter02, LM_FL_f27_diverter03, LM_FL_f27_diverter04, LM_FL_f27_domes, LM_FL_f27_parts, LM_FL_f27_sw25, LM_FL_f27_sw26, LM_FL_f27_sw27, LM_FL_f27_sw28, LM_FL_f27_sw29, LM_FL_f27_sw30, LM_FL_f27_sw31, LM_FL_f27_sw33, LM_FL_f27_sw45, LM_FL_f27_sw57, LM_FL_f27_sw58, LM_FL_f27_sw59, LM_FL_f27_underPF)
Dim BL_FL_f28: BL_FL_f28=Array(LM_FL_f28_BR2, LM_FL_f28_BS1, LM_FL_f28_BS2, LM_FL_f28_BS3, LM_FL_f28_Gate2_001, LM_FL_f28_Gate4, LM_FL_f28_Gate_Orbit, LM_FL_f28_Layer0, LM_FL_f28_Layer1, LM_FL_f28_Layer2, LM_FL_f28_Layer3, LM_FL_f28_LayerCaps, LM_FL_f28_Playfield, LM_FL_f28_SS_01, LM_FL_f28_SS_02, LM_FL_f28_SS_03, LM_FL_f28_SS_04, LM_FL_f28_diverter01, LM_FL_f28_diverter02, LM_FL_f28_diverter03, LM_FL_f28_diverter04, LM_FL_f28_domes, LM_FL_f28_parts, LM_FL_f28_sw28, LM_FL_f28_sw29, LM_FL_f28_sw30, LM_FL_f28_sw31, LM_FL_f28_sw57, LM_FL_f28_sw58, LM_FL_f28_underPF)
Dim BL_FL_f29: BL_FL_f29=Array(LM_FL_f29_BR2, LM_FL_f29_Layer2, LM_FL_f29_Layer3, LM_FL_f29_LayerCaps, LM_FL_f29_Playfield, LM_FL_f29_SS_01, LM_FL_f29_SS_02, LM_FL_f29_SS_03, LM_FL_f29_SS_04, LM_FL_f29_diverter01, LM_FL_f29_diverter02, LM_FL_f29_diverter03, LM_FL_f29_diverter04, LM_FL_f29_domes, LM_FL_f29_parts, LM_FL_f29_underPF)
Dim BL_GIG: BL_GIG=Array(LM_GIG_BR1, LM_GIG_BR2, LM_GIG_BR3, LM_GIG_BS1, LM_GIG_BS2, LM_GIG_BS3, LM_GIG_Gate1, LM_GIG_Gate2_001, LM_GIG_Gate4, LM_GIG_Gate_Orbit, LM_GIG_Layer0, LM_GIG_Layer1, LM_GIG_Layer2, LM_GIG_Layer3, LM_GIG_LayerCaps, LM_GIG_Playfield, LM_GIG_RSling0_001, LM_GIG_RSling1, LM_GIG_SS_01, LM_GIG_SS_02, LM_GIG_SS_03, LM_GIG_SS_04, LM_GIG_diverter01, LM_GIG_diverter02, LM_GIG_diverter03, LM_GIG_diverter04, LM_GIG_domes, LM_GIG_parts, LM_GIG_sw18, LM_GIG_sw19, LM_GIG_sw20, LM_GIG_sw21, LM_GIG_sw22, LM_GIG_sw23, LM_GIG_sw24, LM_GIG_sw25, LM_GIG_sw26, LM_GIG_sw27, LM_GIG_sw28, LM_GIG_sw29, LM_GIG_sw30, LM_GIG_sw31, LM_GIG_sw33, LM_GIG_sw35, LM_GIG_sw38, LM_GIG_sw39, LM_GIG_sw40, LM_GIG_sw45, LM_GIG_sw57, LM_GIG_sw58, LM_GIG_sw59, LM_GIG_underPF)
Dim BL_GIGS_GI_G_001: BL_GIGS_GI_G_001=Array(LM_GIGS_GI_G_001_LEMK, LM_GIGS_GI_G_001_LSling0_001, LM_GIGS_GI_G_001_LSling1, LM_GIGS_GI_G_001_LSling2, LM_GIGS_GI_G_001_Layer0, LM_GIGS_GI_G_001_Layer1, LM_GIGS_GI_G_001_LeftFlipper, LM_GIGS_GI_G_001_Playfield, LM_GIGS_GI_G_001_parts, LM_GIGS_GI_G_001_sw17, LM_GIGS_GI_G_001_sw18, LM_GIGS_GI_G_001_sw19, LM_GIGS_GI_G_001_underPF)
Dim BL_GIGS_GI_G_002: BL_GIGS_GI_G_002=Array(LM_GIGS_GI_G_002_LEMK, LM_GIGS_GI_G_002_LSling0_001, LM_GIGS_GI_G_002_LSling1, LM_GIGS_GI_G_002_LSling2, LM_GIGS_GI_G_002_Layer0, LM_GIGS_GI_G_002_Layer1, LM_GIGS_GI_G_002_Playfield, LM_GIGS_GI_G_002_parts, LM_GIGS_GI_G_002_sw17, LM_GIGS_GI_G_002_underPF)
Dim BL_GIGS_GI_G_020: BL_GIGS_GI_G_020=Array(LM_GIGS_GI_G_020_Playfield, LM_GIGS_GI_G_020_domes, LM_GIGS_GI_G_020_parts, LM_GIGS_GI_G_020_underPF)
Dim BL_GIGS_GI_G_022: BL_GIGS_GI_G_022=Array(LM_GIGS_GI_G_022_Layer0, LM_GIGS_GI_G_022_Layer1, LM_GIGS_GI_G_022_Playfield, LM_GIGS_GI_G_022_REMK, LM_GIGS_GI_G_022_RSling0_001, LM_GIGS_GI_G_022_RSling1, LM_GIGS_GI_G_022_RSling2, LM_GIGS_GI_G_022_domes, LM_GIGS_GI_G_022_parts, LM_GIGS_GI_G_022_sw32, LM_GIGS_GI_G_022_underPF)
Dim BL_GIGS_GI_G_023: BL_GIGS_GI_G_023=Array(LM_GIGS_GI_G_023_Layer0, LM_GIGS_GI_G_023_Layer1, LM_GIGS_GI_G_023_Playfield, LM_GIGS_GI_G_023_REMK, LM_GIGS_GI_G_023_RSling0_001, LM_GIGS_GI_G_023_RSling1, LM_GIGS_GI_G_023_RSling2, LM_GIGS_GI_G_023_RightFlipper, LM_GIGS_GI_G_023_domes, LM_GIGS_GI_G_023_parts, LM_GIGS_GI_G_023_sw32, LM_GIGS_GI_G_023_underPF)
Dim BL_GIGS_GI_G_024: BL_GIGS_GI_G_024=Array(LM_GIGS_GI_G_024_Layer0, LM_GIGS_GI_G_024_Layer1, LM_GIGS_GI_G_024_Playfield, LM_GIGS_GI_G_024_REMK, LM_GIGS_GI_G_024_RSling0_001, LM_GIGS_GI_G_024_RSling1, LM_GIGS_GI_G_024_RSling2, LM_GIGS_GI_G_024_RightFlipper, LM_GIGS_GI_G_024_domes, LM_GIGS_GI_G_024_parts, LM_GIGS_GI_G_024_sw32, LM_GIGS_GI_G_024_underPF)
Dim BL_GIN: BL_GIN=Array(LM_GIN_BR1, LM_GIN_BR2, LM_GIN_BR3, LM_GIN_BS1, LM_GIN_BS2, LM_GIN_BS3, LM_GIN_Gate1, LM_GIN_Gate2_001, LM_GIN_Gate3, LM_GIN_Gate4, LM_GIN_Gate_Orbit, LM_GIN_Layer0, LM_GIN_Layer1, LM_GIN_Layer2, LM_GIN_Layer3, LM_GIN_LayerCaps, LM_GIN_Playfield, LM_GIN_RSling0_001, LM_GIN_RSling1, LM_GIN_RSling2, LM_GIN_SS_01, LM_GIN_SS_02, LM_GIN_SS_03, LM_GIN_SS_04, LM_GIN_diverter01, LM_GIN_diverter02, LM_GIN_diverter03, LM_GIN_diverter04, LM_GIN_domes, LM_GIN_parts, LM_GIN_sw18, LM_GIN_sw19, LM_GIN_sw20, LM_GIN_sw21, LM_GIN_sw22, LM_GIN_sw23, LM_GIN_sw24, LM_GIN_sw25, LM_GIN_sw26, LM_GIN_sw27, LM_GIN_sw28, LM_GIN_sw29, LM_GIN_sw30, LM_GIN_sw31, LM_GIN_sw33, LM_GIN_sw35, LM_GIN_sw38, LM_GIN_sw39, LM_GIN_sw40, LM_GIN_sw45, LM_GIN_sw57, LM_GIN_sw58, LM_GIN_sw59, LM_GIN_underPF)
Dim BL_GINS_gi01: BL_GINS_gi01=Array(LM_GINS_gi01_LEMK, LM_GINS_gi01_LSling0_001, LM_GINS_gi01_LSling1, LM_GINS_gi01_LSling2, LM_GINS_gi01_Layer0, LM_GINS_gi01_Layer1, LM_GINS_gi01_LeftFlipper, LM_GINS_gi01_Playfield, LM_GINS_gi01_parts, LM_GINS_gi01_sw17, LM_GINS_gi01_underPF)
Dim BL_GINS_gi02: BL_GINS_gi02=Array(LM_GINS_gi02_LEMK, LM_GINS_gi02_LSling0_001, LM_GINS_gi02_LSling1, LM_GINS_gi02_LSling2, LM_GINS_gi02_Layer0, LM_GINS_gi02_Layer1, LM_GINS_gi02_Playfield, LM_GINS_gi02_parts, LM_GINS_gi02_sw17, LM_GINS_gi02_underPF)
Dim BL_GINS_gi24: BL_GINS_gi24=Array(LM_GINS_gi24_Layer0, LM_GINS_gi24_Layer1, LM_GINS_gi24_Playfield, LM_GINS_gi24_RSling0_001, LM_GINS_gi24_RSling1, LM_GINS_gi24_RSling2, LM_GINS_gi24_domes, LM_GINS_gi24_parts, LM_GINS_gi24_sw32, LM_GINS_gi24_sw33, LM_GINS_gi24_underPF)
Dim BL_GINS_gi25: BL_GINS_gi25=Array(LM_GINS_gi25_Layer0, LM_GINS_gi25_Layer1, LM_GINS_gi25_Playfield, LM_GINS_gi25_REMK, LM_GINS_gi25_RSling0_001, LM_GINS_gi25_RSling1, LM_GINS_gi25_RSling2, LM_GINS_gi25_domes, LM_GINS_gi25_parts, LM_GINS_gi25_sw32, LM_GINS_gi25_underPF)
Dim BL_GINS_gi26: BL_GINS_gi26=Array(LM_GINS_gi26_Layer0, LM_GINS_gi26_Layer1, LM_GINS_gi26_Playfield, LM_GINS_gi26_REMK, LM_GINS_gi26_RSling0_001, LM_GINS_gi26_RSling1, LM_GINS_gi26_RSling2, LM_GINS_gi26_RightFlipper, LM_GINS_gi26_domes, LM_GINS_gi26_parts, LM_GINS_gi26_sw32, LM_GINS_gi26_underPF)
Dim BL_IN_l1: BL_IN_l1=Array(LM_IN_l1_Layer0, LM_IN_l1_Playfield, LM_IN_l1_domes, LM_IN_l1_underPF)
Dim BL_IN_l10: BL_IN_l10=Array(LM_IN_l10_Layer0, LM_IN_l10_Playfield, LM_IN_l10_parts, LM_IN_l10_underPF)
Dim BL_IN_l11: BL_IN_l11=Array(LM_IN_l11_Layer0, LM_IN_l11_Playfield, LM_IN_l11_underPF)
Dim BL_IN_l12: BL_IN_l12=Array(LM_IN_l12_Layer0, LM_IN_l12_Playfield, LM_IN_l12_domes, LM_IN_l12_parts, LM_IN_l12_underPF)
Dim BL_IN_l13: BL_IN_l13=Array(LM_IN_l13_Layer0, LM_IN_l13_Playfield, LM_IN_l13_domes, LM_IN_l13_underPF)
Dim BL_IN_l14: BL_IN_l14=Array(LM_IN_l14_Layer3, LM_IN_l14_Playfield, LM_IN_l14_parts, LM_IN_l14_underPF)
Dim BL_IN_l15: BL_IN_l15=Array(LM_IN_l15_Layer3, LM_IN_l15_Playfield, LM_IN_l15_parts, LM_IN_l15_underPF)
Dim BL_IN_l16: BL_IN_l16=Array(LM_IN_l16_Layer3, LM_IN_l16_Playfield, LM_IN_l16_parts)
Dim BL_IN_l17: BL_IN_l17=Array(LM_IN_l17_LSling0_001, LM_IN_l17_LSling1, LM_IN_l17_LSling2, LM_IN_l17_Layer0, LM_IN_l17_Layer1, LM_IN_l17_Playfield, LM_IN_l17_parts, LM_IN_l17_underPF)
Dim BL_IN_l18: BL_IN_l18=Array(LM_IN_l18_Layer0, LM_IN_l18_Playfield, LM_IN_l18_domes, LM_IN_l18_parts, LM_IN_l18_sw18, LM_IN_l18_underPF)
Dim BL_IN_l19: BL_IN_l19=Array(LM_IN_l19_Layer0, LM_IN_l19_Playfield, LM_IN_l19_domes, LM_IN_l19_sw18, LM_IN_l19_sw19, LM_IN_l19_underPF)
Dim BL_IN_l2: BL_IN_l2=Array(LM_IN_l2_Layer0, LM_IN_l2_Layer1, LM_IN_l2_Playfield, LM_IN_l2_underPF)
Dim BL_IN_l20: BL_IN_l20=Array(LM_IN_l20_Layer0, LM_IN_l20_Playfield, LM_IN_l20_domes, LM_IN_l20_parts, LM_IN_l20_sw19, LM_IN_l20_sw20, LM_IN_l20_underPF)
Dim BL_IN_l21: BL_IN_l21=Array(LM_IN_l21_Layer0, LM_IN_l21_Playfield, LM_IN_l21_domes, LM_IN_l21_parts, LM_IN_l21_sw20, LM_IN_l21_sw21, LM_IN_l21_underPF)
Dim BL_IN_l22: BL_IN_l22=Array(LM_IN_l22_Layer0, LM_IN_l22_Playfield, LM_IN_l22_domes, LM_IN_l22_parts, LM_IN_l22_sw21, LM_IN_l22_sw22, LM_IN_l22_underPF)
Dim BL_IN_l23: BL_IN_l23=Array(LM_IN_l23_Layer0, LM_IN_l23_Playfield, LM_IN_l23_domes, LM_IN_l23_parts, LM_IN_l23_sw22, LM_IN_l23_sw23, LM_IN_l23_underPF)
Dim BL_IN_l24: BL_IN_l24=Array(LM_IN_l24_Layer0, LM_IN_l24_Playfield, LM_IN_l24_domes, LM_IN_l24_parts, LM_IN_l24_sw23, LM_IN_l24_sw24, LM_IN_l24_underPF)
Dim BL_IN_l25: BL_IN_l25=Array(LM_IN_l25_Layer0, LM_IN_l25_Playfield, LM_IN_l25_parts, LM_IN_l25_sw25, LM_IN_l25_sw26, LM_IN_l25_underPF)
Dim BL_IN_l26: BL_IN_l26=Array(LM_IN_l26_Layer0, LM_IN_l26_Playfield, LM_IN_l26_parts, LM_IN_l26_sw26, LM_IN_l26_sw27, LM_IN_l26_underPF)
Dim BL_IN_l27: BL_IN_l27=Array(LM_IN_l27_BR3, LM_IN_l27_BS3, LM_IN_l27_Layer0, LM_IN_l27_Playfield, LM_IN_l27_parts, LM_IN_l27_sw27, LM_IN_l27_underPF)
Dim BL_IN_l28: BL_IN_l28=Array(LM_IN_l28_Layer0, LM_IN_l28_Layer3, LM_IN_l28_Playfield, LM_IN_l28_SS_01, LM_IN_l28_SS_03, LM_IN_l28_diverter01, LM_IN_l28_diverter02, LM_IN_l28_domes, LM_IN_l28_parts, LM_IN_l28_sw28, LM_IN_l28_sw29, LM_IN_l28_underPF)
Dim BL_IN_l29: BL_IN_l29=Array(LM_IN_l29_Layer0, LM_IN_l29_Playfield, LM_IN_l29_domes, LM_IN_l29_parts, LM_IN_l29_sw29, LM_IN_l29_sw30, LM_IN_l29_underPF)
Dim BL_IN_l3: BL_IN_l3=Array(LM_IN_l3_Layer0, LM_IN_l3_Playfield, LM_IN_l3_underPF)
Dim BL_IN_l30: BL_IN_l30=Array(LM_IN_l30_Layer0, LM_IN_l30_Playfield, LM_IN_l30_domes, LM_IN_l30_parts, LM_IN_l30_sw30, LM_IN_l30_sw31, LM_IN_l30_underPF)
Dim BL_IN_l31: BL_IN_l31=Array(LM_IN_l31_Layer0, LM_IN_l31_Playfield, LM_IN_l31_domes, LM_IN_l31_parts, LM_IN_l31_sw31, LM_IN_l31_underPF)
Dim BL_IN_l32: BL_IN_l32=Array(LM_IN_l32_Layer0, LM_IN_l32_Layer1, LM_IN_l32_Playfield, LM_IN_l32_RSling0_001, LM_IN_l32_RSling1, LM_IN_l32_domes, LM_IN_l32_parts, LM_IN_l32_underPF)
Dim BL_IN_l33: BL_IN_l33=Array(LM_IN_l33_Playfield, LM_IN_l33_parts, LM_IN_l33_sw33, LM_IN_l33_underPF)
Dim BL_IN_l34: BL_IN_l34=Array(LM_IN_l34_Layer0, LM_IN_l34_Playfield, LM_IN_l34_parts, LM_IN_l34_underPF)
Dim BL_IN_l35: BL_IN_l35=Array(LM_IN_l35_Layer0, LM_IN_l35_Playfield, LM_IN_l35_domes, LM_IN_l35_parts, LM_IN_l35_underPF)
Dim BL_IN_l36: BL_IN_l36=Array(LM_IN_l36_Layer0, LM_IN_l36_Layer3, LM_IN_l36_Playfield, LM_IN_l36_parts, LM_IN_l36_underPF)
Dim BL_IN_l38: BL_IN_l38=Array(LM_IN_l38_Layer3, LM_IN_l38_Playfield, LM_IN_l38_parts, LM_IN_l38_underPF)
Dim BL_IN_l39: BL_IN_l39=Array(LM_IN_l39_Layer3, LM_IN_l39_Playfield, LM_IN_l39_parts, LM_IN_l39_underPF)
Dim BL_IN_l4: BL_IN_l4=Array(LM_IN_l4_Layer0, LM_IN_l4_Playfield, LM_IN_l4_underPF)
Dim BL_IN_l40: BL_IN_l40=Array(LM_IN_l40_Layer2, LM_IN_l40_Layer3, LM_IN_l40_Playfield, LM_IN_l40_parts, LM_IN_l40_underPF)
Dim BL_IN_l42: BL_IN_l42=Array(LM_IN_l42_Layer0, LM_IN_l42_Playfield, LM_IN_l42_underPF)
Dim BL_IN_l43: BL_IN_l43=Array(LM_IN_l43_Layer0, LM_IN_l43_Playfield, LM_IN_l43_underPF)
Dim BL_IN_l44: BL_IN_l44=Array(LM_IN_l44_Layer0, LM_IN_l44_Playfield, LM_IN_l44_underPF)
Dim BL_IN_l45: BL_IN_l45=Array(LM_IN_l45_Layer0, LM_IN_l45_Playfield, LM_IN_l45_underPF)
Dim BL_IN_l5: BL_IN_l5=Array(LM_IN_l5_Layer0, LM_IN_l5_Playfield, LM_IN_l5_diverter01, LM_IN_l5_diverter02, LM_IN_l5_diverter03, LM_IN_l5_diverter04, LM_IN_l5_parts, LM_IN_l5_sw59, LM_IN_l5_underPF)
Dim BL_IN_l50: BL_IN_l50=Array(LM_IN_l50_Layer0, LM_IN_l50_Playfield, LM_IN_l50_parts, LM_IN_l50_underPF)
Dim BL_IN_l51: BL_IN_l51=Array(LM_IN_l51_Layer0, LM_IN_l51_Playfield, LM_IN_l51_parts, LM_IN_l51_underPF)
Dim BL_IN_l52: BL_IN_l52=Array(LM_IN_l52_Layer0, LM_IN_l52_Playfield, LM_IN_l52_underPF)
Dim BL_IN_l53: BL_IN_l53=Array(LM_IN_l53_Layer0, LM_IN_l53_Playfield, LM_IN_l53_underPF)
Dim BL_IN_l54: BL_IN_l54=Array(LM_IN_l54_Layer0, LM_IN_l54_Playfield, LM_IN_l54_underPF)
Dim BL_IN_l55: BL_IN_l55=Array(LM_IN_l55_Layer0, LM_IN_l55_Playfield, LM_IN_l55_underPF)
Dim BL_IN_l56: BL_IN_l56=Array(LM_IN_l56_Layer0, LM_IN_l56_Playfield, LM_IN_l56_underPF)
Dim BL_IN_l58: BL_IN_l58=Array(LM_IN_l58_Layer0, LM_IN_l58_Playfield, LM_IN_l58_parts, LM_IN_l58_underPF)
Dim BL_IN_l59: BL_IN_l59=Array(LM_IN_l59_Layer0, LM_IN_l59_Playfield, LM_IN_l59_parts, LM_IN_l59_underPF)
Dim BL_IN_l6: BL_IN_l6=Array(LM_IN_l6_Layer0, LM_IN_l6_Playfield, LM_IN_l6_parts, LM_IN_l6_sw57, LM_IN_l6_sw58, LM_IN_l6_underPF)
Dim BL_IN_l60: BL_IN_l60=Array(LM_IN_l60_Layer0, LM_IN_l60_Playfield, LM_IN_l60_underPF)
Dim BL_IN_l61: BL_IN_l61=Array(LM_IN_l61_Layer0, LM_IN_l61_Playfield, LM_IN_l61_underPF)
Dim BL_IN_l62: BL_IN_l62=Array(LM_IN_l62_Layer0, LM_IN_l62_Playfield, LM_IN_l62_underPF)
Dim BL_IN_l63: BL_IN_l63=Array(LM_IN_l63_Layer0, LM_IN_l63_Playfield, LM_IN_l63_underPF)
Dim BL_IN_l64: BL_IN_l64=Array(LM_IN_l64_Layer0, LM_IN_l64_Playfield, LM_IN_l64_underPF)
Dim BL_IN_l7: BL_IN_l7=Array(LM_IN_l7_Layer0, LM_IN_l7_Layer1, LM_IN_l7_Layer3, LM_IN_l7_Playfield, LM_IN_l7_parts, LM_IN_l7_underPF)
Dim BL_IN_l8: BL_IN_l8=Array(LM_IN_l8_Layer0, LM_IN_l8_Playfield, LM_IN_l8_parts, LM_IN_l8_underPF)
Dim BL_IN_l9: BL_IN_l9=Array(LM_IN_l9_Layer0, LM_IN_l9_Playfield, LM_IN_l9_parts, LM_IN_l9_sw33, LM_IN_l9_sw35, LM_IN_l9_underPF)
Dim BL_Room: BL_Room=Array(BM_BR1, BM_BR2, BM_BR3, BM_BS1, BM_BS2, BM_BS3, BM_Gate1, BM_Gate2_001, BM_Gate3, BM_Gate4, BM_Gate_Orbit, BM_LEMK, BM_LSling0_001, BM_LSling1, BM_LSling2, BM_Layer0, BM_Layer1, BM_Layer2, BM_Layer3, BM_LayerCaps, BM_LeftFlipper, BM_Playfield, BM_REMK, BM_RSling0_001, BM_RSling1, BM_RSling2, BM_RightFlipper, BM_SS_01, BM_SS_02, BM_SS_03, BM_SS_04, BM_diverter01, BM_diverter02, BM_diverter03, BM_diverter04, BM_domes, BM_parts, BM_sw17, BM_sw18, BM_sw19, BM_sw20, BM_sw21, BM_sw22, BM_sw23, BM_sw24, BM_sw25, BM_sw26, BM_sw27, BM_sw28, BM_sw29, BM_sw30, BM_sw31, BM_sw32, BM_sw33, BM_sw35, BM_sw38, BM_sw39, BM_sw40, BM_sw43, BM_sw45, BM_sw57, BM_sw58, BM_sw59, BM_underPF)
' Global arrays
Dim BG_Bakemap: BG_Bakemap=Array(BM_BR1, BM_BR2, BM_BR3, BM_BS1, BM_BS2, BM_BS3, BM_Gate1, BM_Gate2_001, BM_Gate3, BM_Gate4, BM_Gate_Orbit, BM_LEMK, BM_LSling0_001, BM_LSling1, BM_LSling2, BM_Layer0, BM_Layer1, BM_Layer2, BM_Layer3, BM_LayerCaps, BM_LeftFlipper, BM_Playfield, BM_REMK, BM_RSling0_001, BM_RSling1, BM_RSling2, BM_RightFlipper, BM_SS_01, BM_SS_02, BM_SS_03, BM_SS_04, BM_diverter01, BM_diverter02, BM_diverter03, BM_diverter04, BM_domes, BM_parts, BM_sw17, BM_sw18, BM_sw19, BM_sw20, BM_sw21, BM_sw22, BM_sw23, BM_sw24, BM_sw25, BM_sw26, BM_sw27, BM_sw28, BM_sw29, BM_sw30, BM_sw31, BM_sw32, BM_sw33, BM_sw35, BM_sw38, BM_sw39, BM_sw40, BM_sw43, BM_sw45, BM_sw57, BM_sw58, BM_sw59, BM_underPF)
Dim BG_Lightmap: BG_Lightmap=Array(LM_Bumpers_BR1, LM_Bumpers_BR2, LM_Bumpers_BR3, LM_Bumpers_BS1, LM_Bumpers_BS2, LM_Bumpers_BS3, LM_Bumpers_Layer1, LM_Bumpers_Layer2, LM_Bumpers_Layer3, LM_Bumpers_LayerCaps, LM_Bumpers_Playfield, LM_Bumpers_SS_01, LM_Bumpers_SS_02, LM_Bumpers_SS_03, LM_Bumpers_SS_04, LM_Bumpers_diverter01, LM_Bumpers_domes, LM_Bumpers_parts, LM_Bumpers_sw25, LM_Bumpers_sw26, LM_Bumpers_sw27, LM_Bumpers_underPF, LM_FL_f15_BR2, LM_FL_f15_BS3, LM_FL_f15_Layer0, LM_FL_f15_Layer1, LM_FL_f15_Layer2, LM_FL_f15_Layer3, LM_FL_f15_LayerCaps, LM_FL_f15_Playfield, LM_FL_f15_SS_01, LM_FL_f15_SS_02, LM_FL_f15_SS_03, LM_FL_f15_SS_04, LM_FL_f15_diverter01, LM_FL_f15_diverter02, LM_FL_f15_diverter03, LM_FL_f15_diverter04, LM_FL_f15_domes, LM_FL_f15_parts, LM_FL_f15_underPF, LM_FL_f25_LEMK, LM_FL_f25_LSling0_001, LM_FL_f25_LSling1, LM_FL_f25_LSling2, LM_FL_f25_Layer0, LM_FL_f25_Layer1, LM_FL_f25_LeftFlipper, LM_FL_f25_Playfield, LM_FL_f25_REMK, LM_FL_f25_RSling0_001, LM_FL_f25_RSling1, LM_FL_f25_RSling2, _
  LM_FL_f25_RightFlipper, LM_FL_f25_domes, LM_FL_f25_parts, LM_FL_f25_sw18, LM_FL_f25_sw19, LM_FL_f25_sw20, LM_FL_f25_sw21, LM_FL_f25_sw33, LM_FL_f25_sw35, LM_FL_f25_underPF, LM_FL_f26_BR2, LM_FL_f26_BR3, LM_FL_f26_BS2, LM_FL_f26_BS3, LM_FL_f26_Gate1, LM_FL_f26_LEMK, LM_FL_f26_LSling0_001, LM_FL_f26_LSling1, LM_FL_f26_LSling2, LM_FL_f26_Layer0, LM_FL_f26_Layer1, LM_FL_f26_Playfield, LM_FL_f26_RSling0_001, LM_FL_f26_RSling2, LM_FL_f26_domes, LM_FL_f26_parts, LM_FL_f26_sw18, LM_FL_f26_sw19, LM_FL_f26_sw20, LM_FL_f26_sw21, LM_FL_f26_sw22, LM_FL_f26_sw23, LM_FL_f26_sw24, LM_FL_f26_sw25, LM_FL_f26_sw26, LM_FL_f26_sw27, LM_FL_f26_sw33, LM_FL_f26_underPF, LM_FL_f27_BR1, LM_FL_f27_BR2, LM_FL_f27_BR3, LM_FL_f27_BS1, LM_FL_f27_BS2, LM_FL_f27_BS3, LM_FL_f27_Gate2_001, LM_FL_f27_LSling0_001, LM_FL_f27_LSling1, LM_FL_f27_LSling2, LM_FL_f27_Layer0, LM_FL_f27_Layer1, LM_FL_f27_Layer2, LM_FL_f27_Layer3, LM_FL_f27_LayerCaps, LM_FL_f27_Playfield, LM_FL_f27_REMK, LM_FL_f27_RSling0_001, LM_FL_f27_RSling1, LM_FL_f27_RSling2, _
  LM_FL_f27_RightFlipper, LM_FL_f27_SS_01, LM_FL_f27_SS_02, LM_FL_f27_SS_03, LM_FL_f27_SS_04, LM_FL_f27_diverter01, LM_FL_f27_diverter02, LM_FL_f27_diverter03, LM_FL_f27_diverter04, LM_FL_f27_domes, LM_FL_f27_parts, LM_FL_f27_sw25, LM_FL_f27_sw26, LM_FL_f27_sw27, LM_FL_f27_sw28, LM_FL_f27_sw29, LM_FL_f27_sw30, LM_FL_f27_sw31, LM_FL_f27_sw33, LM_FL_f27_sw45, LM_FL_f27_sw57, LM_FL_f27_sw58, LM_FL_f27_sw59, LM_FL_f27_underPF, LM_FL_f28_BR2, LM_FL_f28_BS1, LM_FL_f28_BS2, LM_FL_f28_BS3, LM_FL_f28_Gate2_001, LM_FL_f28_Gate4, LM_FL_f28_Gate_Orbit, LM_FL_f28_Layer0, LM_FL_f28_Layer1, LM_FL_f28_Layer2, LM_FL_f28_Layer3, LM_FL_f28_LayerCaps, LM_FL_f28_Playfield, LM_FL_f28_SS_01, LM_FL_f28_SS_02, LM_FL_f28_SS_03, LM_FL_f28_SS_04, LM_FL_f28_diverter01, LM_FL_f28_diverter02, LM_FL_f28_diverter03, LM_FL_f28_diverter04, LM_FL_f28_domes, LM_FL_f28_parts, LM_FL_f28_sw28, LM_FL_f28_sw29, LM_FL_f28_sw30, LM_FL_f28_sw31, LM_FL_f28_sw57, LM_FL_f28_sw58, LM_FL_f28_underPF, LM_FL_f29_BR2, LM_FL_f29_Layer2, LM_FL_f29_Layer3, _
  LM_FL_f29_LayerCaps, LM_FL_f29_Playfield, LM_FL_f29_SS_01, LM_FL_f29_SS_02, LM_FL_f29_SS_03, LM_FL_f29_SS_04, LM_FL_f29_diverter01, LM_FL_f29_diverter02, LM_FL_f29_diverter03, LM_FL_f29_diverter04, LM_FL_f29_domes, LM_FL_f29_parts, LM_FL_f29_underPF, LM_GIG_BR1, LM_GIG_BR2, LM_GIG_BR3, LM_GIG_BS1, LM_GIG_BS2, LM_GIG_BS3, LM_GIG_Gate1, LM_GIG_Gate2_001, LM_GIG_Gate4, LM_GIG_Gate_Orbit, LM_GIG_Layer0, LM_GIG_Layer1, LM_GIG_Layer2, LM_GIG_Layer3, LM_GIG_LayerCaps, LM_GIG_Playfield, LM_GIG_RSling0_001, LM_GIG_RSling1, LM_GIG_SS_01, LM_GIG_SS_02, LM_GIG_SS_03, LM_GIG_SS_04, LM_GIG_diverter01, LM_GIG_diverter02, LM_GIG_diverter03, LM_GIG_diverter04, LM_GIG_domes, LM_GIG_parts, LM_GIG_sw18, LM_GIG_sw19, LM_GIG_sw20, LM_GIG_sw21, LM_GIG_sw22, LM_GIG_sw23, LM_GIG_sw24, LM_GIG_sw25, LM_GIG_sw26, LM_GIG_sw27, LM_GIG_sw28, LM_GIG_sw29, LM_GIG_sw30, LM_GIG_sw31, LM_GIG_sw33, LM_GIG_sw35, LM_GIG_sw38, LM_GIG_sw39, LM_GIG_sw40, LM_GIG_sw45, LM_GIG_sw57, LM_GIG_sw58, LM_GIG_sw59, LM_GIG_underPF, LM_GIGS_GI_G_001_LEMK, _
  LM_GIGS_GI_G_001_LSling0_001, LM_GIGS_GI_G_001_LSling1, LM_GIGS_GI_G_001_LSling2, LM_GIGS_GI_G_001_Layer0, LM_GIGS_GI_G_001_Layer1, LM_GIGS_GI_G_001_LeftFlipper, LM_GIGS_GI_G_001_Playfield, LM_GIGS_GI_G_001_parts, LM_GIGS_GI_G_001_sw17, LM_GIGS_GI_G_001_sw18, LM_GIGS_GI_G_001_sw19, LM_GIGS_GI_G_001_underPF, LM_GIGS_GI_G_002_LEMK, LM_GIGS_GI_G_002_LSling0_001, LM_GIGS_GI_G_002_LSling1, LM_GIGS_GI_G_002_LSling2, LM_GIGS_GI_G_002_Layer0, LM_GIGS_GI_G_002_Layer1, LM_GIGS_GI_G_002_Playfield, LM_GIGS_GI_G_002_parts, LM_GIGS_GI_G_002_sw17, LM_GIGS_GI_G_002_underPF, LM_GIGS_GI_G_020_Playfield, LM_GIGS_GI_G_020_domes, LM_GIGS_GI_G_020_parts, LM_GIGS_GI_G_020_underPF, LM_GIGS_GI_G_022_Layer0, LM_GIGS_GI_G_022_Layer1, LM_GIGS_GI_G_022_Playfield, LM_GIGS_GI_G_022_REMK, LM_GIGS_GI_G_022_RSling0_001, LM_GIGS_GI_G_022_RSling1, LM_GIGS_GI_G_022_RSling2, LM_GIGS_GI_G_022_domes, LM_GIGS_GI_G_022_parts, LM_GIGS_GI_G_022_sw32, LM_GIGS_GI_G_022_underPF, LM_GIGS_GI_G_023_Layer0, LM_GIGS_GI_G_023_Layer1, LM_GIGS_GI_G_023_Playfield, _
  LM_GIGS_GI_G_023_REMK, LM_GIGS_GI_G_023_RSling0_001, LM_GIGS_GI_G_023_RSling1, LM_GIGS_GI_G_023_RSling2, LM_GIGS_GI_G_023_RightFlipper, LM_GIGS_GI_G_023_domes, LM_GIGS_GI_G_023_parts, LM_GIGS_GI_G_023_sw32, LM_GIGS_GI_G_023_underPF, LM_GIGS_GI_G_024_Layer0, LM_GIGS_GI_G_024_Layer1, LM_GIGS_GI_G_024_Playfield, LM_GIGS_GI_G_024_REMK, LM_GIGS_GI_G_024_RSling0_001, LM_GIGS_GI_G_024_RSling1, LM_GIGS_GI_G_024_RSling2, LM_GIGS_GI_G_024_RightFlipper, LM_GIGS_GI_G_024_domes, LM_GIGS_GI_G_024_parts, LM_GIGS_GI_G_024_sw32, LM_GIGS_GI_G_024_underPF, LM_GIN_BR1, LM_GIN_BR2, LM_GIN_BR3, LM_GIN_BS1, LM_GIN_BS2, LM_GIN_BS3, LM_GIN_Gate1, LM_GIN_Gate2_001, LM_GIN_Gate3, LM_GIN_Gate4, LM_GIN_Gate_Orbit, LM_GIN_Layer0, LM_GIN_Layer1, LM_GIN_Layer2, LM_GIN_Layer3, LM_GIN_LayerCaps, LM_GIN_Playfield, LM_GIN_RSling0_001, LM_GIN_RSling1, LM_GIN_RSling2, LM_GIN_SS_01, LM_GIN_SS_02, LM_GIN_SS_03, LM_GIN_SS_04, LM_GIN_diverter01, LM_GIN_diverter02, LM_GIN_diverter03, LM_GIN_diverter04, LM_GIN_domes, LM_GIN_parts, LM_GIN_sw18, _
  LM_GIN_sw19, LM_GIN_sw20, LM_GIN_sw21, LM_GIN_sw22, LM_GIN_sw23, LM_GIN_sw24, LM_GIN_sw25, LM_GIN_sw26, LM_GIN_sw27, LM_GIN_sw28, LM_GIN_sw29, LM_GIN_sw30, LM_GIN_sw31, LM_GIN_sw33, LM_GIN_sw35, LM_GIN_sw38, LM_GIN_sw39, LM_GIN_sw40, LM_GIN_sw45, LM_GIN_sw57, LM_GIN_sw58, LM_GIN_sw59, LM_GIN_underPF, LM_GINS_gi01_LEMK, LM_GINS_gi01_LSling0_001, LM_GINS_gi01_LSling1, LM_GINS_gi01_LSling2, LM_GINS_gi01_Layer0, LM_GINS_gi01_Layer1, LM_GINS_gi01_LeftFlipper, LM_GINS_gi01_Playfield, LM_GINS_gi01_parts, LM_GINS_gi01_sw17, LM_GINS_gi01_underPF, LM_GINS_gi02_LEMK, LM_GINS_gi02_LSling0_001, LM_GINS_gi02_LSling1, LM_GINS_gi02_LSling2, LM_GINS_gi02_Layer0, LM_GINS_gi02_Layer1, LM_GINS_gi02_Playfield, LM_GINS_gi02_parts, LM_GINS_gi02_sw17, LM_GINS_gi02_underPF, LM_GINS_gi24_Layer0, LM_GINS_gi24_Layer1, LM_GINS_gi24_Playfield, LM_GINS_gi24_RSling0_001, LM_GINS_gi24_RSling1, LM_GINS_gi24_RSling2, LM_GINS_gi24_domes, LM_GINS_gi24_parts, LM_GINS_gi24_sw32, LM_GINS_gi24_sw33, LM_GINS_gi24_underPF, LM_GINS_gi25_Layer0, _
  LM_GINS_gi25_Layer1, LM_GINS_gi25_Playfield, LM_GINS_gi25_REMK, LM_GINS_gi25_RSling0_001, LM_GINS_gi25_RSling1, LM_GINS_gi25_RSling2, LM_GINS_gi25_domes, LM_GINS_gi25_parts, LM_GINS_gi25_sw32, LM_GINS_gi25_underPF, LM_GINS_gi26_Layer0, LM_GINS_gi26_Layer1, LM_GINS_gi26_Playfield, LM_GINS_gi26_REMK, LM_GINS_gi26_RSling0_001, LM_GINS_gi26_RSling1, LM_GINS_gi26_RSling2, LM_GINS_gi26_RightFlipper, LM_GINS_gi26_domes, LM_GINS_gi26_parts, LM_GINS_gi26_sw32, LM_GINS_gi26_underPF, LM_IN_l1_Layer0, LM_IN_l1_Playfield, LM_IN_l1_domes, LM_IN_l1_underPF, LM_IN_l10_Layer0, LM_IN_l10_Playfield, LM_IN_l10_parts, LM_IN_l10_underPF, LM_IN_l11_Layer0, LM_IN_l11_Playfield, LM_IN_l11_underPF, LM_IN_l12_Layer0, LM_IN_l12_Playfield, LM_IN_l12_domes, LM_IN_l12_parts, LM_IN_l12_underPF, LM_IN_l13_Layer0, LM_IN_l13_Playfield, LM_IN_l13_domes, LM_IN_l13_underPF, LM_IN_l14_Layer3, LM_IN_l14_Playfield, LM_IN_l14_parts, LM_IN_l14_underPF, LM_IN_l15_Layer3, LM_IN_l15_Playfield, LM_IN_l15_parts, LM_IN_l15_underPF, LM_IN_l16_Layer3, _
  LM_IN_l16_Playfield, LM_IN_l16_parts, LM_IN_l17_LSling0_001, LM_IN_l17_LSling1, LM_IN_l17_LSling2, LM_IN_l17_Layer0, LM_IN_l17_Layer1, LM_IN_l17_Playfield, LM_IN_l17_parts, LM_IN_l17_underPF, LM_IN_l18_Layer0, LM_IN_l18_Playfield, LM_IN_l18_domes, LM_IN_l18_parts, LM_IN_l18_sw18, LM_IN_l18_underPF, LM_IN_l19_Layer0, LM_IN_l19_Playfield, LM_IN_l19_domes, LM_IN_l19_sw18, LM_IN_l19_sw19, LM_IN_l19_underPF, LM_IN_l2_Layer0, LM_IN_l2_Layer1, LM_IN_l2_Playfield, LM_IN_l2_underPF, LM_IN_l20_Layer0, LM_IN_l20_Playfield, LM_IN_l20_domes, LM_IN_l20_parts, LM_IN_l20_sw19, LM_IN_l20_sw20, LM_IN_l20_underPF, LM_IN_l21_Layer0, LM_IN_l21_Playfield, LM_IN_l21_domes, LM_IN_l21_parts, LM_IN_l21_sw20, LM_IN_l21_sw21, LM_IN_l21_underPF, LM_IN_l22_Layer0, LM_IN_l22_Playfield, LM_IN_l22_domes, LM_IN_l22_parts, LM_IN_l22_sw21, LM_IN_l22_sw22, LM_IN_l22_underPF, LM_IN_l23_Layer0, LM_IN_l23_Playfield, LM_IN_l23_domes, LM_IN_l23_parts, LM_IN_l23_sw22, LM_IN_l23_sw23, LM_IN_l23_underPF, LM_IN_l24_Layer0, LM_IN_l24_Playfield, _
  LM_IN_l24_domes, LM_IN_l24_parts, LM_IN_l24_sw23, LM_IN_l24_sw24, LM_IN_l24_underPF, LM_IN_l25_Layer0, LM_IN_l25_Playfield, LM_IN_l25_parts, LM_IN_l25_sw25, LM_IN_l25_sw26, LM_IN_l25_underPF, LM_IN_l26_Layer0, LM_IN_l26_Playfield, LM_IN_l26_parts, LM_IN_l26_sw26, LM_IN_l26_sw27, LM_IN_l26_underPF, LM_IN_l27_BR3, LM_IN_l27_BS3, LM_IN_l27_Layer0, LM_IN_l27_Playfield, LM_IN_l27_parts, LM_IN_l27_sw27, LM_IN_l27_underPF, LM_IN_l28_Layer0, LM_IN_l28_Layer3, LM_IN_l28_Playfield, LM_IN_l28_SS_01, LM_IN_l28_SS_03, LM_IN_l28_diverter01, LM_IN_l28_diverter02, LM_IN_l28_domes, LM_IN_l28_parts, LM_IN_l28_sw28, LM_IN_l28_sw29, LM_IN_l28_underPF, LM_IN_l29_Layer0, LM_IN_l29_Playfield, LM_IN_l29_domes, LM_IN_l29_parts, LM_IN_l29_sw29, LM_IN_l29_sw30, LM_IN_l29_underPF, LM_IN_l3_Layer0, LM_IN_l3_Playfield, LM_IN_l3_underPF, LM_IN_l30_Layer0, LM_IN_l30_Playfield, LM_IN_l30_domes, LM_IN_l30_parts, LM_IN_l30_sw30, LM_IN_l30_sw31, LM_IN_l30_underPF, LM_IN_l31_Layer0, LM_IN_l31_Playfield, LM_IN_l31_domes, LM_IN_l31_parts, _
  LM_IN_l31_sw31, LM_IN_l31_underPF, LM_IN_l32_Layer0, LM_IN_l32_Layer1, LM_IN_l32_Playfield, LM_IN_l32_RSling0_001, LM_IN_l32_RSling1, LM_IN_l32_domes, LM_IN_l32_parts, LM_IN_l32_underPF, LM_IN_l33_Playfield, LM_IN_l33_parts, LM_IN_l33_sw33, LM_IN_l33_underPF, LM_IN_l34_Layer0, LM_IN_l34_Playfield, LM_IN_l34_parts, LM_IN_l34_underPF, LM_IN_l35_Layer0, LM_IN_l35_Playfield, LM_IN_l35_domes, LM_IN_l35_parts, LM_IN_l35_underPF, LM_IN_l36_Layer0, LM_IN_l36_Layer3, LM_IN_l36_Playfield, LM_IN_l36_parts, LM_IN_l36_underPF, LM_IN_l38_Layer3, LM_IN_l38_Playfield, LM_IN_l38_parts, LM_IN_l38_underPF, LM_IN_l39_Layer3, LM_IN_l39_Playfield, LM_IN_l39_parts, LM_IN_l39_underPF, LM_IN_l4_Layer0, LM_IN_l4_Playfield, LM_IN_l4_underPF, LM_IN_l40_Layer2, LM_IN_l40_Layer3, LM_IN_l40_Playfield, LM_IN_l40_parts, LM_IN_l40_underPF, LM_IN_l42_Layer0, LM_IN_l42_Playfield, LM_IN_l42_underPF, LM_IN_l43_Layer0, LM_IN_l43_Playfield, LM_IN_l43_underPF, LM_IN_l44_Layer0, LM_IN_l44_Playfield, LM_IN_l44_underPF, LM_IN_l45_Layer0, _
  LM_IN_l45_Playfield, LM_IN_l45_underPF, LM_IN_l5_Layer0, LM_IN_l5_Playfield, LM_IN_l5_diverter01, LM_IN_l5_diverter02, LM_IN_l5_diverter03, LM_IN_l5_diverter04, LM_IN_l5_parts, LM_IN_l5_sw59, LM_IN_l5_underPF, LM_IN_l50_Layer0, LM_IN_l50_Playfield, LM_IN_l50_parts, LM_IN_l50_underPF, LM_IN_l51_Layer0, LM_IN_l51_Playfield, LM_IN_l51_parts, LM_IN_l51_underPF, LM_IN_l52_Layer0, LM_IN_l52_Playfield, LM_IN_l52_underPF, LM_IN_l53_Layer0, LM_IN_l53_Playfield, LM_IN_l53_underPF, LM_IN_l54_Layer0, LM_IN_l54_Playfield, LM_IN_l54_underPF, LM_IN_l55_Layer0, LM_IN_l55_Playfield, LM_IN_l55_underPF, LM_IN_l56_Layer0, LM_IN_l56_Playfield, LM_IN_l56_underPF, LM_IN_l58_Layer0, LM_IN_l58_Playfield, LM_IN_l58_parts, LM_IN_l58_underPF, LM_IN_l59_Layer0, LM_IN_l59_Playfield, LM_IN_l59_parts, LM_IN_l59_underPF, LM_IN_l6_Layer0, LM_IN_l6_Playfield, LM_IN_l6_parts, LM_IN_l6_sw57, LM_IN_l6_sw58, LM_IN_l6_underPF, LM_IN_l60_Layer0, LM_IN_l60_Playfield, LM_IN_l60_underPF, LM_IN_l61_Layer0, LM_IN_l61_Playfield, LM_IN_l61_underPF, _
  LM_IN_l62_Layer0, LM_IN_l62_Playfield, LM_IN_l62_underPF, LM_IN_l63_Layer0, LM_IN_l63_Playfield, LM_IN_l63_underPF, LM_IN_l64_Layer0, LM_IN_l64_Playfield, LM_IN_l64_underPF, LM_IN_l7_Layer0, LM_IN_l7_Layer1, LM_IN_l7_Layer3, LM_IN_l7_Playfield, LM_IN_l7_parts, LM_IN_l7_underPF, LM_IN_l8_Layer0, LM_IN_l8_Playfield, LM_IN_l8_parts, LM_IN_l8_underPF, LM_IN_l9_Layer0, LM_IN_l9_Playfield, LM_IN_l9_parts, LM_IN_l9_sw33, LM_IN_l9_sw35, LM_IN_l9_underPF)
Dim BG_All: BG_All=Array(BM_BR1, BM_BR2, BM_BR3, BM_BS1, BM_BS2, BM_BS3, BM_Gate1, BM_Gate2_001, BM_Gate3, BM_Gate4, BM_Gate_Orbit, BM_LEMK, BM_LSling0_001, BM_LSling1, BM_LSling2, BM_Layer0, BM_Layer1, BM_Layer2, BM_Layer3, BM_LayerCaps, BM_LeftFlipper, BM_Playfield, BM_REMK, BM_RSling0_001, BM_RSling1, BM_RSling2, BM_RightFlipper, BM_SS_01, BM_SS_02, BM_SS_03, BM_SS_04, BM_diverter01, BM_diverter02, BM_diverter03, BM_diverter04, BM_domes, BM_parts, BM_sw17, BM_sw18, BM_sw19, BM_sw20, BM_sw21, BM_sw22, BM_sw23, BM_sw24, BM_sw25, BM_sw26, BM_sw27, BM_sw28, BM_sw29, BM_sw30, BM_sw31, BM_sw32, BM_sw33, BM_sw35, BM_sw38, BM_sw39, BM_sw40, BM_sw43, BM_sw45, BM_sw57, BM_sw58, BM_sw59, BM_underPF, LM_Bumpers_BR1, LM_Bumpers_BR2, LM_Bumpers_BR3, LM_Bumpers_BS1, LM_Bumpers_BS2, LM_Bumpers_BS3, LM_Bumpers_Layer1, LM_Bumpers_Layer2, LM_Bumpers_Layer3, LM_Bumpers_LayerCaps, LM_Bumpers_Playfield, LM_Bumpers_SS_01, LM_Bumpers_SS_02, LM_Bumpers_SS_03, LM_Bumpers_SS_04, LM_Bumpers_diverter01, LM_Bumpers_domes, _
  LM_Bumpers_parts, LM_Bumpers_sw25, LM_Bumpers_sw26, LM_Bumpers_sw27, LM_Bumpers_underPF, LM_FL_f15_BR2, LM_FL_f15_BS3, LM_FL_f15_Layer0, LM_FL_f15_Layer1, LM_FL_f15_Layer2, LM_FL_f15_Layer3, LM_FL_f15_LayerCaps, LM_FL_f15_Playfield, LM_FL_f15_SS_01, LM_FL_f15_SS_02, LM_FL_f15_SS_03, LM_FL_f15_SS_04, LM_FL_f15_diverter01, LM_FL_f15_diverter02, LM_FL_f15_diverter03, LM_FL_f15_diverter04, LM_FL_f15_domes, LM_FL_f15_parts, LM_FL_f15_underPF, LM_FL_f25_LEMK, LM_FL_f25_LSling0_001, LM_FL_f25_LSling1, LM_FL_f25_LSling2, LM_FL_f25_Layer0, LM_FL_f25_Layer1, LM_FL_f25_LeftFlipper, LM_FL_f25_Playfield, LM_FL_f25_REMK, LM_FL_f25_RSling0_001, LM_FL_f25_RSling1, LM_FL_f25_RSling2, LM_FL_f25_RightFlipper, LM_FL_f25_domes, LM_FL_f25_parts, LM_FL_f25_sw18, LM_FL_f25_sw19, LM_FL_f25_sw20, LM_FL_f25_sw21, LM_FL_f25_sw33, LM_FL_f25_sw35, LM_FL_f25_underPF, LM_FL_f26_BR2, LM_FL_f26_BR3, LM_FL_f26_BS2, LM_FL_f26_BS3, LM_FL_f26_Gate1, LM_FL_f26_LEMK, LM_FL_f26_LSling0_001, LM_FL_f26_LSling1, LM_FL_f26_LSling2, LM_FL_f26_Layer0, _
  LM_FL_f26_Layer1, LM_FL_f26_Playfield, LM_FL_f26_RSling0_001, LM_FL_f26_RSling2, LM_FL_f26_domes, LM_FL_f26_parts, LM_FL_f26_sw18, LM_FL_f26_sw19, LM_FL_f26_sw20, LM_FL_f26_sw21, LM_FL_f26_sw22, LM_FL_f26_sw23, LM_FL_f26_sw24, LM_FL_f26_sw25, LM_FL_f26_sw26, LM_FL_f26_sw27, LM_FL_f26_sw33, LM_FL_f26_underPF, LM_FL_f27_BR1, LM_FL_f27_BR2, LM_FL_f27_BR3, LM_FL_f27_BS1, LM_FL_f27_BS2, LM_FL_f27_BS3, LM_FL_f27_Gate2_001, LM_FL_f27_LSling0_001, LM_FL_f27_LSling1, LM_FL_f27_LSling2, LM_FL_f27_Layer0, LM_FL_f27_Layer1, LM_FL_f27_Layer2, LM_FL_f27_Layer3, LM_FL_f27_LayerCaps, LM_FL_f27_Playfield, LM_FL_f27_REMK, LM_FL_f27_RSling0_001, LM_FL_f27_RSling1, LM_FL_f27_RSling2, LM_FL_f27_RightFlipper, LM_FL_f27_SS_01, LM_FL_f27_SS_02, LM_FL_f27_SS_03, LM_FL_f27_SS_04, LM_FL_f27_diverter01, LM_FL_f27_diverter02, LM_FL_f27_diverter03, LM_FL_f27_diverter04, LM_FL_f27_domes, LM_FL_f27_parts, LM_FL_f27_sw25, LM_FL_f27_sw26, LM_FL_f27_sw27, LM_FL_f27_sw28, LM_FL_f27_sw29, LM_FL_f27_sw30, LM_FL_f27_sw31, LM_FL_f27_sw33, _
  LM_FL_f27_sw45, LM_FL_f27_sw57, LM_FL_f27_sw58, LM_FL_f27_sw59, LM_FL_f27_underPF, LM_FL_f28_BR2, LM_FL_f28_BS1, LM_FL_f28_BS2, LM_FL_f28_BS3, LM_FL_f28_Gate2_001, LM_FL_f28_Gate4, LM_FL_f28_Gate_Orbit, LM_FL_f28_Layer0, LM_FL_f28_Layer1, LM_FL_f28_Layer2, LM_FL_f28_Layer3, LM_FL_f28_LayerCaps, LM_FL_f28_Playfield, LM_FL_f28_SS_01, LM_FL_f28_SS_02, LM_FL_f28_SS_03, LM_FL_f28_SS_04, LM_FL_f28_diverter01, LM_FL_f28_diverter02, LM_FL_f28_diverter03, LM_FL_f28_diverter04, LM_FL_f28_domes, LM_FL_f28_parts, LM_FL_f28_sw28, LM_FL_f28_sw29, LM_FL_f28_sw30, LM_FL_f28_sw31, LM_FL_f28_sw57, LM_FL_f28_sw58, LM_FL_f28_underPF, LM_FL_f29_BR2, LM_FL_f29_Layer2, LM_FL_f29_Layer3, LM_FL_f29_LayerCaps, LM_FL_f29_Playfield, LM_FL_f29_SS_01, LM_FL_f29_SS_02, LM_FL_f29_SS_03, LM_FL_f29_SS_04, LM_FL_f29_diverter01, LM_FL_f29_diverter02, LM_FL_f29_diverter03, LM_FL_f29_diverter04, LM_FL_f29_domes, LM_FL_f29_parts, LM_FL_f29_underPF, LM_GIG_BR1, LM_GIG_BR2, LM_GIG_BR3, LM_GIG_BS1, LM_GIG_BS2, LM_GIG_BS3, LM_GIG_Gate1, _
  LM_GIG_Gate2_001, LM_GIG_Gate4, LM_GIG_Gate_Orbit, LM_GIG_Layer0, LM_GIG_Layer1, LM_GIG_Layer2, LM_GIG_Layer3, LM_GIG_LayerCaps, LM_GIG_Playfield, LM_GIG_RSling0_001, LM_GIG_RSling1, LM_GIG_SS_01, LM_GIG_SS_02, LM_GIG_SS_03, LM_GIG_SS_04, LM_GIG_diverter01, LM_GIG_diverter02, LM_GIG_diverter03, LM_GIG_diverter04, LM_GIG_domes, LM_GIG_parts, LM_GIG_sw18, LM_GIG_sw19, LM_GIG_sw20, LM_GIG_sw21, LM_GIG_sw22, LM_GIG_sw23, LM_GIG_sw24, LM_GIG_sw25, LM_GIG_sw26, LM_GIG_sw27, LM_GIG_sw28, LM_GIG_sw29, LM_GIG_sw30, LM_GIG_sw31, LM_GIG_sw33, LM_GIG_sw35, LM_GIG_sw38, LM_GIG_sw39, LM_GIG_sw40, LM_GIG_sw45, LM_GIG_sw57, LM_GIG_sw58, LM_GIG_sw59, LM_GIG_underPF, LM_GIGS_GI_G_001_LEMK, LM_GIGS_GI_G_001_LSling0_001, LM_GIGS_GI_G_001_LSling1, LM_GIGS_GI_G_001_LSling2, LM_GIGS_GI_G_001_Layer0, LM_GIGS_GI_G_001_Layer1, LM_GIGS_GI_G_001_LeftFlipper, LM_GIGS_GI_G_001_Playfield, LM_GIGS_GI_G_001_parts, LM_GIGS_GI_G_001_sw17, LM_GIGS_GI_G_001_sw18, LM_GIGS_GI_G_001_sw19, LM_GIGS_GI_G_001_underPF, LM_GIGS_GI_G_002_LEMK, _
  LM_GIGS_GI_G_002_LSling0_001, LM_GIGS_GI_G_002_LSling1, LM_GIGS_GI_G_002_LSling2, LM_GIGS_GI_G_002_Layer0, LM_GIGS_GI_G_002_Layer1, LM_GIGS_GI_G_002_Playfield, LM_GIGS_GI_G_002_parts, LM_GIGS_GI_G_002_sw17, LM_GIGS_GI_G_002_underPF, LM_GIGS_GI_G_020_Playfield, LM_GIGS_GI_G_020_domes, LM_GIGS_GI_G_020_parts, LM_GIGS_GI_G_020_underPF, LM_GIGS_GI_G_022_Layer0, LM_GIGS_GI_G_022_Layer1, LM_GIGS_GI_G_022_Playfield, LM_GIGS_GI_G_022_REMK, LM_GIGS_GI_G_022_RSling0_001, LM_GIGS_GI_G_022_RSling1, LM_GIGS_GI_G_022_RSling2, LM_GIGS_GI_G_022_domes, LM_GIGS_GI_G_022_parts, LM_GIGS_GI_G_022_sw32, LM_GIGS_GI_G_022_underPF, LM_GIGS_GI_G_023_Layer0, LM_GIGS_GI_G_023_Layer1, LM_GIGS_GI_G_023_Playfield, LM_GIGS_GI_G_023_REMK, LM_GIGS_GI_G_023_RSling0_001, LM_GIGS_GI_G_023_RSling1, LM_GIGS_GI_G_023_RSling2, LM_GIGS_GI_G_023_RightFlipper, LM_GIGS_GI_G_023_domes, LM_GIGS_GI_G_023_parts, LM_GIGS_GI_G_023_sw32, LM_GIGS_GI_G_023_underPF, LM_GIGS_GI_G_024_Layer0, LM_GIGS_GI_G_024_Layer1, LM_GIGS_GI_G_024_Playfield, _
  LM_GIGS_GI_G_024_REMK, LM_GIGS_GI_G_024_RSling0_001, LM_GIGS_GI_G_024_RSling1, LM_GIGS_GI_G_024_RSling2, LM_GIGS_GI_G_024_RightFlipper, LM_GIGS_GI_G_024_domes, LM_GIGS_GI_G_024_parts, LM_GIGS_GI_G_024_sw32, LM_GIGS_GI_G_024_underPF, LM_GIN_BR1, LM_GIN_BR2, LM_GIN_BR3, LM_GIN_BS1, LM_GIN_BS2, LM_GIN_BS3, LM_GIN_Gate1, LM_GIN_Gate2_001, LM_GIN_Gate3, LM_GIN_Gate4, LM_GIN_Gate_Orbit, LM_GIN_Layer0, LM_GIN_Layer1, LM_GIN_Layer2, LM_GIN_Layer3, LM_GIN_LayerCaps, LM_GIN_Playfield, LM_GIN_RSling0_001, LM_GIN_RSling1, LM_GIN_RSling2, LM_GIN_SS_01, LM_GIN_SS_02, LM_GIN_SS_03, LM_GIN_SS_04, LM_GIN_diverter01, LM_GIN_diverter02, LM_GIN_diverter03, LM_GIN_diverter04, LM_GIN_domes, LM_GIN_parts, LM_GIN_sw18, LM_GIN_sw19, LM_GIN_sw20, LM_GIN_sw21, LM_GIN_sw22, LM_GIN_sw23, LM_GIN_sw24, LM_GIN_sw25, LM_GIN_sw26, LM_GIN_sw27, LM_GIN_sw28, LM_GIN_sw29, LM_GIN_sw30, LM_GIN_sw31, LM_GIN_sw33, LM_GIN_sw35, LM_GIN_sw38, LM_GIN_sw39, LM_GIN_sw40, LM_GIN_sw45, LM_GIN_sw57, LM_GIN_sw58, LM_GIN_sw59, LM_GIN_underPF, _
  LM_GINS_gi01_LEMK, LM_GINS_gi01_LSling0_001, LM_GINS_gi01_LSling1, LM_GINS_gi01_LSling2, LM_GINS_gi01_Layer0, LM_GINS_gi01_Layer1, LM_GINS_gi01_LeftFlipper, LM_GINS_gi01_Playfield, LM_GINS_gi01_parts, LM_GINS_gi01_sw17, LM_GINS_gi01_underPF, LM_GINS_gi02_LEMK, LM_GINS_gi02_LSling0_001, LM_GINS_gi02_LSling1, LM_GINS_gi02_LSling2, LM_GINS_gi02_Layer0, LM_GINS_gi02_Layer1, LM_GINS_gi02_Playfield, LM_GINS_gi02_parts, LM_GINS_gi02_sw17, LM_GINS_gi02_underPF, LM_GINS_gi24_Layer0, LM_GINS_gi24_Layer1, LM_GINS_gi24_Playfield, LM_GINS_gi24_RSling0_001, LM_GINS_gi24_RSling1, LM_GINS_gi24_RSling2, LM_GINS_gi24_domes, LM_GINS_gi24_parts, LM_GINS_gi24_sw32, LM_GINS_gi24_sw33, LM_GINS_gi24_underPF, LM_GINS_gi25_Layer0, LM_GINS_gi25_Layer1, LM_GINS_gi25_Playfield, LM_GINS_gi25_REMK, LM_GINS_gi25_RSling0_001, LM_GINS_gi25_RSling1, LM_GINS_gi25_RSling2, LM_GINS_gi25_domes, LM_GINS_gi25_parts, LM_GINS_gi25_sw32, LM_GINS_gi25_underPF, LM_GINS_gi26_Layer0, LM_GINS_gi26_Layer1, LM_GINS_gi26_Playfield, LM_GINS_gi26_REMK, _
  LM_GINS_gi26_RSling0_001, LM_GINS_gi26_RSling1, LM_GINS_gi26_RSling2, LM_GINS_gi26_RightFlipper, LM_GINS_gi26_domes, LM_GINS_gi26_parts, LM_GINS_gi26_sw32, LM_GINS_gi26_underPF, LM_IN_l1_Layer0, LM_IN_l1_Playfield, LM_IN_l1_domes, LM_IN_l1_underPF, LM_IN_l10_Layer0, LM_IN_l10_Playfield, LM_IN_l10_parts, LM_IN_l10_underPF, LM_IN_l11_Layer0, LM_IN_l11_Playfield, LM_IN_l11_underPF, LM_IN_l12_Layer0, LM_IN_l12_Playfield, LM_IN_l12_domes, LM_IN_l12_parts, LM_IN_l12_underPF, LM_IN_l13_Layer0, LM_IN_l13_Playfield, LM_IN_l13_domes, LM_IN_l13_underPF, LM_IN_l14_Layer3, LM_IN_l14_Playfield, LM_IN_l14_parts, LM_IN_l14_underPF, LM_IN_l15_Layer3, LM_IN_l15_Playfield, LM_IN_l15_parts, LM_IN_l15_underPF, LM_IN_l16_Layer3, LM_IN_l16_Playfield, LM_IN_l16_parts, LM_IN_l17_LSling0_001, LM_IN_l17_LSling1, LM_IN_l17_LSling2, LM_IN_l17_Layer0, LM_IN_l17_Layer1, LM_IN_l17_Playfield, LM_IN_l17_parts, LM_IN_l17_underPF, LM_IN_l18_Layer0, LM_IN_l18_Playfield, LM_IN_l18_domes, LM_IN_l18_parts, LM_IN_l18_sw18, LM_IN_l18_underPF, _
  LM_IN_l19_Layer0, LM_IN_l19_Playfield, LM_IN_l19_domes, LM_IN_l19_sw18, LM_IN_l19_sw19, LM_IN_l19_underPF, LM_IN_l2_Layer0, LM_IN_l2_Layer1, LM_IN_l2_Playfield, LM_IN_l2_underPF, LM_IN_l20_Layer0, LM_IN_l20_Playfield, LM_IN_l20_domes, LM_IN_l20_parts, LM_IN_l20_sw19, LM_IN_l20_sw20, LM_IN_l20_underPF, LM_IN_l21_Layer0, LM_IN_l21_Playfield, LM_IN_l21_domes, LM_IN_l21_parts, LM_IN_l21_sw20, LM_IN_l21_sw21, LM_IN_l21_underPF, LM_IN_l22_Layer0, LM_IN_l22_Playfield, LM_IN_l22_domes, LM_IN_l22_parts, LM_IN_l22_sw21, LM_IN_l22_sw22, LM_IN_l22_underPF, LM_IN_l23_Layer0, LM_IN_l23_Playfield, LM_IN_l23_domes, LM_IN_l23_parts, LM_IN_l23_sw22, LM_IN_l23_sw23, LM_IN_l23_underPF, LM_IN_l24_Layer0, LM_IN_l24_Playfield, LM_IN_l24_domes, LM_IN_l24_parts, LM_IN_l24_sw23, LM_IN_l24_sw24, LM_IN_l24_underPF, LM_IN_l25_Layer0, LM_IN_l25_Playfield, LM_IN_l25_parts, LM_IN_l25_sw25, LM_IN_l25_sw26, LM_IN_l25_underPF, LM_IN_l26_Layer0, LM_IN_l26_Playfield, LM_IN_l26_parts, LM_IN_l26_sw26, LM_IN_l26_sw27, LM_IN_l26_underPF, _
  LM_IN_l27_BR3, LM_IN_l27_BS3, LM_IN_l27_Layer0, LM_IN_l27_Playfield, LM_IN_l27_parts, LM_IN_l27_sw27, LM_IN_l27_underPF, LM_IN_l28_Layer0, LM_IN_l28_Layer3, LM_IN_l28_Playfield, LM_IN_l28_SS_01, LM_IN_l28_SS_03, LM_IN_l28_diverter01, LM_IN_l28_diverter02, LM_IN_l28_domes, LM_IN_l28_parts, LM_IN_l28_sw28, LM_IN_l28_sw29, LM_IN_l28_underPF, LM_IN_l29_Layer0, LM_IN_l29_Playfield, LM_IN_l29_domes, LM_IN_l29_parts, LM_IN_l29_sw29, LM_IN_l29_sw30, LM_IN_l29_underPF, LM_IN_l3_Layer0, LM_IN_l3_Playfield, LM_IN_l3_underPF, LM_IN_l30_Layer0, LM_IN_l30_Playfield, LM_IN_l30_domes, LM_IN_l30_parts, LM_IN_l30_sw30, LM_IN_l30_sw31, LM_IN_l30_underPF, LM_IN_l31_Layer0, LM_IN_l31_Playfield, LM_IN_l31_domes, LM_IN_l31_parts, LM_IN_l31_sw31, LM_IN_l31_underPF, LM_IN_l32_Layer0, LM_IN_l32_Layer1, LM_IN_l32_Playfield, LM_IN_l32_RSling0_001, LM_IN_l32_RSling1, LM_IN_l32_domes, LM_IN_l32_parts, LM_IN_l32_underPF, LM_IN_l33_Playfield, LM_IN_l33_parts, LM_IN_l33_sw33, LM_IN_l33_underPF, LM_IN_l34_Layer0, LM_IN_l34_Playfield, _
  LM_IN_l34_parts, LM_IN_l34_underPF, LM_IN_l35_Layer0, LM_IN_l35_Playfield, LM_IN_l35_domes, LM_IN_l35_parts, LM_IN_l35_underPF, LM_IN_l36_Layer0, LM_IN_l36_Layer3, LM_IN_l36_Playfield, LM_IN_l36_parts, LM_IN_l36_underPF, LM_IN_l38_Layer3, LM_IN_l38_Playfield, LM_IN_l38_parts, LM_IN_l38_underPF, LM_IN_l39_Layer3, LM_IN_l39_Playfield, LM_IN_l39_parts, LM_IN_l39_underPF, LM_IN_l4_Layer0, LM_IN_l4_Playfield, LM_IN_l4_underPF, LM_IN_l40_Layer2, LM_IN_l40_Layer3, LM_IN_l40_Playfield, LM_IN_l40_parts, LM_IN_l40_underPF, LM_IN_l42_Layer0, LM_IN_l42_Playfield, LM_IN_l42_underPF, LM_IN_l43_Layer0, LM_IN_l43_Playfield, LM_IN_l43_underPF, LM_IN_l44_Layer0, LM_IN_l44_Playfield, LM_IN_l44_underPF, LM_IN_l45_Layer0, LM_IN_l45_Playfield, LM_IN_l45_underPF, LM_IN_l5_Layer0, LM_IN_l5_Playfield, LM_IN_l5_diverter01, LM_IN_l5_diverter02, LM_IN_l5_diverter03, LM_IN_l5_diverter04, LM_IN_l5_parts, LM_IN_l5_sw59, LM_IN_l5_underPF, LM_IN_l50_Layer0, LM_IN_l50_Playfield, LM_IN_l50_parts, LM_IN_l50_underPF, LM_IN_l51_Layer0, _
  LM_IN_l51_Playfield, LM_IN_l51_parts, LM_IN_l51_underPF, LM_IN_l52_Layer0, LM_IN_l52_Playfield, LM_IN_l52_underPF, LM_IN_l53_Layer0, LM_IN_l53_Playfield, LM_IN_l53_underPF, LM_IN_l54_Layer0, LM_IN_l54_Playfield, LM_IN_l54_underPF, LM_IN_l55_Layer0, LM_IN_l55_Playfield, LM_IN_l55_underPF, LM_IN_l56_Layer0, LM_IN_l56_Playfield, LM_IN_l56_underPF, LM_IN_l58_Layer0, LM_IN_l58_Playfield, LM_IN_l58_parts, LM_IN_l58_underPF, LM_IN_l59_Layer0, LM_IN_l59_Playfield, LM_IN_l59_parts, LM_IN_l59_underPF, LM_IN_l6_Layer0, LM_IN_l6_Playfield, LM_IN_l6_parts, LM_IN_l6_sw57, LM_IN_l6_sw58, LM_IN_l6_underPF, LM_IN_l60_Layer0, LM_IN_l60_Playfield, LM_IN_l60_underPF, LM_IN_l61_Layer0, LM_IN_l61_Playfield, LM_IN_l61_underPF, LM_IN_l62_Layer0, LM_IN_l62_Playfield, LM_IN_l62_underPF, LM_IN_l63_Layer0, LM_IN_l63_Playfield, LM_IN_l63_underPF, LM_IN_l64_Layer0, LM_IN_l64_Playfield, LM_IN_l64_underPF, LM_IN_l7_Layer0, LM_IN_l7_Layer1, LM_IN_l7_Layer3, LM_IN_l7_Playfield, LM_IN_l7_parts, LM_IN_l7_underPF, LM_IN_l8_Layer0, _
  LM_IN_l8_Playfield, LM_IN_l8_parts, LM_IN_l8_underPF, LM_IN_l9_Layer0, LM_IN_l9_Playfield, LM_IN_l9_parts, LM_IN_l9_sw33, LM_IN_l9_sw35, LM_IN_l9_underPF)
' VLM  Arrays - End


'******************************************************
'  ZVAR: Constants and Global Variables
'******************************************************
Const DebugFlashers = False
Const DebugGI = False

Const BallSize = 50   'Ball size must be 50
Const BallMass = 1    'Ball mass must be 1
Const tnob = 3      'Total number of balls on the playfield including captive balls.
Const lob = 0     'Total number of locked balls

Dim tablewidth: tablewidth = Table1.width
Dim tableheight: tableheight = Table1.height

Dim BIPL              'Ball in plunger lane
BIPL = False

Dim plungerIM

'  Standard definitions
Const cGameName = "spstn_l5"  'PinMAME ROM name
Const UseSolenoids = 2      '1 = Normal Flippers, 2 = Fastflips
Const UseLamps = 1        '0 = Custom lamp handling, 1 = Built-in VPX handling (using light number in light timer)
Const UseSync = 0
Const HandleMech = 0
Const SSolenoidOn = ""      'Sound sample used for this, obsolete.
Const SSolenoidOff = ""     ' ^
Const SFlipperOn = ""     ' ^
Const SFlipperOff = ""      ' ^
Const SCoin = ""        ' ^

'flag for MB mode, primarily for GI lighting/ball tinting
dim UseMultiballTint : UseMultiballTint = True
dim IsMultiball : isMultiball = False ' track if multiball based on callback

'----- VR Room -----
Dim VRRoomChoice : VRRoomChoice = 1         ' 1 - Mega Room, 2 - Minimal Room

'VRRoom set based on RenderingMode
'Internal DMD in Desktop Mode, using a textbox (must be called before LoadVPM)
Dim UseVPMDMD, DesktopMode, dtlight
DesktopMode = Table1.ShowDT

If Not DesktopMode Then
  For Each dtlight in DesktopLights: dtlight.visible = 0: Next
End If

'----- VR Room Auto-Detect -----
Dim VRRoom, VR_Obj, VRMode
Const VRTest = 0        ' 1 = Testing VR in Live View, 0 = Do not force VR mode.


'Const UseVPMModSol = 1   'Old PWM method. Don't use this
Const UseVPMModSol = 2    'Set to 2 for PWM flashers, inserts, and GI. Requires VPinMame 3.6

'NOTES on UseVPMModSol = 2:
'  - Only supported for S9/S11/DataEast/WPC/Capcom/Whitestar (Sega & Stern)/SAM
'  - All lights on the table must have their Fader model set tp "LED (None)" to get the correct fading effects
'  - When not supported VPM outputs only 0 or 1. Therefore, use VPX "Incandescent" fader for lights

LoadVPM "03060000", "s11.vbs", 3.36  'The "03060000" argument forces user to have VPinMame 3.6

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
  DoSTAnim          'Standup target animations
  UpdateStandupTargets
  DoDTAnim          'Drop target animations
  UpdateDropTargets
  MotorUpdate
  If VRMode = True Then
    VRDisplayTimer
  Else
    UpdateLeds
  End If
End Sub

'The CorTimer interval should be 10. It's sole purpose is to update the Cor (physics) calculations
CorTimer.Interval = 10
Sub CorTimer_Timer(): Cor.Update: End Sub

'******************************************************
'  ZINI: Table Initialization and Exiting
'******************************************************
Dim SSBall1, SSBall2, SSBall3, gBOT

Sub Table1_Init
  vpminit me
  With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
    .SplashInfoLine = "Space Station (Williams 1987)" & vbNewLine & "VPW"
    .HandleKeyboard = 0
    .ShowTitle = 0
    .ShowDMDOnly = 1
    .ShowFrame = 0
    .HandleMechanics = 0
  End With
  On Error Resume Next
  Controller.Run
  If Err Then MsgBox Err.Description
  On Error Goto 0

  'Map all lamps to the corresponding ROM output using the value of TimerInterval of each light object
  vpmMapLights AllLamps     'Make a collection called "AllLamps" and put all the light objects in it.

  'Nudging
  vpmNudge.TiltSwitch=9
  vpmNudge.Sensitivity=4
  vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingshot)

  'Main Timer init
  PinMAMETimer.Interval = PinMAMEInterval
  PinMAMETimer.Enabled = 1

  'Trough - Creates a ball in the kicker switch and gives that ball used an individual name.
  Set SSBall1 = sw11.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set SSBall2 = sw12.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set SSBall3 = sw13.CreateSizedballWithMass(Ballsize/2,Ballmass)

  'Forces the trough switches to "on" at table boot so the game logic knows there are balls in the trough.
  Controller.Switch(11) = 1
  Controller.Switch(12) = 1
  Controller.Switch(13) = 1

  '***Setting up a ball array (collection), must contain all the balls you create on the table.
  gBOT = Array(SSBall1,SSBall2,SSBall3)

  Kickback.PullBack
  StationBlocker.isdropped = 1
  GISelect 1

  LStep = 0: LeftSlingShot.TimerEnabled = 1
  RStep = 0: RightSlingShot.TimerEnabled = 1

  SetBackglass
  InitDigits

  dim LM
  for each LM in BP_Layer3: LM.color = RGB(40,72,157): next
  BP_Layer3(0).color = RGB(255,255,255)

  'Lighting tweaks

  Dim BL :
  For Each BL in BL_GIN : BL.opacity = 150 : Next
  For Each BL in BL_GIG : BL.opacity = 200 : Next
  For Each BL in BL_GINS_gi01 : BL.opacity = 150 : Next
  For Each BL in BL_GINS_gi02 : BL.opacity = 150 : Next
  For Each BL in BL_GINS_gi24 : BL.opacity = 150 : Next
  For Each BL in BL_GINS_gi25 : BL.opacity = 150 : Next
  For Each BL in BL_GINS_gi26 : BL.opacity = 150 : Next
  For Each BL in BL_GIGS_GI_G_001: BL.opacity = 200 : Next
  For Each BL in BL_GIGS_GI_G_002: BL.opacity = 200 : Next
  For Each BL in BL_GIGS_GI_G_020: BL.opacity = 200 : Next
  For Each BL in BL_GIGS_GI_G_022: BL.opacity = 200 : Next
  For Each BL in BL_GIGS_GI_G_023: BL.opacity = 200 : Next
  For Each BL in BL_GIGS_GI_G_024: BL.opacity = 200 : Next
  For Each BL in BL_FL_f15: BL.opacity = 70 : Next
  For Each BL in BL_FL_f26: BL.opacity = 300 : Next
  For Each BL in BL_FL_f25: BL.opacity = 300 : Next
  For Each BL in BL_FL_f27: BL.opacity = 300 : Next
  For Each BL in BL_FL_f28: BL.opacity = 300 : Next

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

Dim LightLevel : LightLevel = 0.25        ' Level of room lighting (0 to 1), where 0 is dark and 100 is brightest
Dim ColorLUT : ColorLUT = 1           ' Color desaturation LUTs: 1 to 12, where 1 is vibrant, 2 is normal and 12 is black'n'white
Dim VolumeDial : VolumeDial = 0.8             ' Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Dim BallRollVolume : BallRollVolume = 1       ' Level of ball rolling volume. Value between 0 and 1
Dim RampRollVolume : RampRollVolume = 0.5     ' Level of ramp rolling volume. Value between 0 and 1
Dim StagedFlippers : StagedFlippers = 0         ' Staged Flippers. 0 = Disabled, 1 = Enabled
Dim mbGI : mbGI = 0

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
  dim v
  dim useTint

  ' Color Saturation
    ColorLUT = Table1.Option("Color Saturation", 1, 12, 1, 1, 0, _
    Array("Vibrant", "Normal", "Desaturated 10%", "Desaturated 20%", "Desaturated 30%", "Desaturated 40%", "Desaturated 50%", _
        "Desaturated 60%", "Desaturated 70%", "Desaturated 80%", "Desaturated 90%", "Black 'n White" ))
  if ColorLUT = 1 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.7vibr80"
  if ColorLUT = 2 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1"
  if ColorLUT = 3 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_90"
  if ColorLUT = 4 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_80"
  if ColorLUT = 5 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_70"
  if ColorLUT = 6 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_60"
  if ColorLUT = 7 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_50"
  if ColorLUT = 8 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_40"
  if ColorLUT = 9 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_30"
  if ColorLUT = 10 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_20"
  if ColorLUT = 11 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_10"
  if ColorLUT = 12 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_00"

  'GI Lamps Color
  mbGI = Table1.Option("Multiball GI Color", 0, 8, 1, 0, 0, Array("Green", "GreenLED", "Blue", "Purple", "Pink", "Red", "Orange", "Yellow", "Random"))
  'Color Temperature
  CurrentColorTempIndex = Table1.Option("Color Temperature", 0, 3, 1, 0, 0, Array("2700k", "3000k", "3500k", "5000k"))
  'Use ball tint
  useTint = Table1.Option("Ball Tint in Multiball", 0, 1, 1, 1, 0, Array("Off","On"))
  if UseTint = 1 then UseMultiballTint = True else UseMultiballTint = False

  SetGiColor mbGI

    ' Sound volumes
    VolumeDial = Table1.Option("Mech Volume", 0, 1, 0.01, 0.8, 1)
    BallRollVolume = Table1.Option("Ball Roll Volume", 0, 1, 0.01, 0.5, 1)
  RampRollVolume = Table1.Option("Ramp Roll Volume", 0, 1, 0.01, 0.5, 1)

  ' Room brightness
' LightLevel = Table1.Option("Table Brightness (Ambient Light Level)", 0, 1, 0.01, .5, 1)
  LightLevel = NightDay/100
  SetRoomBrightness LightLevel   'Uncomment this line for lightmapped tables.

  'Cabinet rails
  v = Table1.Option("Cabinet Rails", 0, 1, 1, 1, 0, Array("Hide", "Show"))
  Primary_SideRailMapRight.visible = v
  Primary_SideRailMapLeft.visible = v
  Primary_LockDownBar.visible = v


    ' Staged Flippers
    'StagedFlippers = Table1.Option("Staged Flippers", 0, 1, 1, 0, 0, Array("Disabled", "Enabled"))

  ' VR Room
    VRRoomChoice = Table1.Option("VR Room", 1, 2, 1, 1, 0, Array("Mega Room", "Minimal Room"))
  SetupRoom

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

Function ExcelMod(number, divisor)
    ExcelMod = number - divisor * Int(number / divisor)
End Function



'******************************************************
'  ZANI: Misc Animations
'******************************************************

' Flipper Animations

Sub LeftFlipper_Animate
  Dim BP
  dim a: a = LeftFlipper.CurrentAngle
  FlipperLSh.RotZ = a
  For Each BP in BP_LeftFlipper
    BP.RotZ = a + 0
  Next
End Sub

Sub RightFlipper_Animate
  Dim BP
  dim a: a = RightFlipper.CurrentAngle
  FlipperRSh.RotZ = a
  For Each BP in BP_RightFlipper
    BP.RotZ = a + 0
  Next
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


' Gate Animations

Sub Gate1_Animate
  Dim a: a = Gate1.currentangle
  Dim BP : For Each BP in BP_Gate1 : BP.RotX = -a: Next
End Sub

Sub Gate2_Animate
  Dim a: a = Gate2.currentangle
  Dim BP : For Each BP in BP_Gate2_001 : BP.RotX = a: Next
End Sub

Sub Gate3_Animate
  Dim a: a = Gate3.currentangle
  Dim BP : For Each BP in BP_Gate3 : BP.RotY = a: Next
End Sub

Sub Gate4_Animate
  Dim a: a = Gate4.currentangle
  Dim BP : For Each BP in BP_Gate4 : BP.RotX = -a: Next
End Sub

Sub Gate_Orbit_Animate
  Dim a: a = Gate_Orbit.currentangle
  Dim BP : For Each BP in BP_Gate_Orbit : BP.RotX = a: Next
End Sub


' Switch Animations

Sub sw17_Animate
  Dim z : z = sw17.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw17 : BP.transz = z: Next
End Sub

Sub sw32_Animate
  Dim z : z = sw32.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw32 : BP.transz = z: Next
End Sub

Sub sw45_Animate
  Dim z : z = sw45.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw45 : BP.transz = z: Next
End Sub



'******************************************************
'   ZBBR: BALL BRIGHTNESS
'******************************************************

Const BallBrightness =  0.8       'Ball brightness - Value between 0 and 1 (0=Dark ... 1=Bright)

' Constants for plunger lane ball darkening.
' You can make a temporary wall in the plunger lane area and use the co-ordinates from the corner control points.
Const PLOffset = 0.5      'Minimum ball brightness scale in plunger lane
Const PLLeft = 865        'X position of punger lane left
Const PLRight = 955       'X position of punger lane right
Const PLTop = 1150        'Y position of punger lane top
Const PLBottom = 1780       'Y position of punger lane bottom
Dim PLGain: PLGain = (1-PLOffset)/(PLTop-PLBottom)


dim ColorTintLocations(1,4) ' used to define locations to fade out tinting effect

' location between the flippers
ColorTintLocations(0,0) = 430 ' x
ColorTintLocations(0,1) = 1704  ' y
ColorTintLocations(0,2) = 450 ' radius
ColorTintLocations(0,3) = 0.4 ' min tint amount
'...add more locations as needed

Function ScaleValueFromDistance(CurrentDistance, MaxDistance, MinScaleValue, MaxScaleValue)
  'scales any value based on distance from a point
    Dim NormalizedDistance, ScaledValue

    ' Calculate the normalized distance (percentage between 0 and 1)
    If MaxDistance <> 0 Then
        NormalizedDistance = CurrentDistance / MaxDistance
    Else
        NormalizedDistance = 0
    End If

    ' Clamp the value between 0 and 1 in case the ball goes outside the defined range
    If NormalizedDistance < 0 Then NormalizedDistance = 0
    If NormalizedDistance > 1 Then NormalizedDistance = 1

    ' Calculate scaled value interpolated between MinScaleValue and MaxScaleValue
    ScaledValue = MinScaleValue + NormalizedDistance * (MaxScaleValue - MinScaleValue)
    ScaleValueFromDistance = ScaledValue
End Function

Function GetBallColorWeight(ball)
  'determine how much weight to use for color tint: 0.00 = none, 1.00 = all
  GetBallColorWeight = 1

  dim weight, balldistance, x

  if ball.z > 35 then
    'arbitrary: when ball is above pf, reduce tint effect
    'using same method as below but distance is on the z axis - ~25 to account for ball height
    balldistance = ball.z - 25
    GetBallColorWeight = 1 - ScaleValueFromDistance(balldistance, 200, 0.5, 1)
    'debug.print "z (" & ball.z & " weight: " & GetBallColorWeight
    exit function
  end if

  for x = 0 to UBound(ColorTintLocations,1)
    balldistance = distance(ball.x,ball.y,ColorTintLocations(x,0),ColorTintLocations(x,1))
    if balldistance <= ColorTintLocations(x,2) then
      GetBallColorWeight = ScaleValueFromDistance(balldistance, ColorTintLocations(x,2), ColorTintLocations(x,3), 1)
      exit function
    end if
  next
End Function

Sub UpdateBallBrightness
  Dim s, b_base, b_r, b_g, b_b, d_w, colorWeight
  b_base = 120 * BallBrightness + 135*gilvl ' orig was 120 and 70

  if IsMultiball and UseMultiballTint then
    For s = 0 To UBound(gBOT)
      ' Handle z direction
      d_w = b_base*(1 - (gBOT(s).z-25)/500)
      If d_w < 30 Then d_w = 30

      'handle special conditions like plunger lane
      If InRect(gBOT(s).x,gBOT(s).y,PLLeft,PLBottom,PLLeft,PLTop,PLRight,PLTop,PLRight,PLBottom) Then
        'ball in plunger lane
        d_w = d_w*(PLOffset+PLGain*(gBOT(s).y-PLBottom))
        colorWeight = 0.2 ' arbitrary 20% coloration when ball in plunger
      Else
        colorWeight = GetBallColorWeight(gBOT(s))
      End If

      ' get base colors for current GI ball
      dim base_c, base_r, base_g, base_b
      base_c = bArray(CurrentGIColorIndex)
      base_r = base_c mod 256
      base_g = (base_c \ 256) mod 256
      base_b = (base_c \ 256 \ 256) mod 256

      ' scale brightness
      b_r = Int((Int(base_r * (d_w/255))*colorWeight) + (Int(d_w)*(1-colorWeight)))
      b_g = Int((Int(base_g * (d_w/255))*colorWeight) + (Int(d_w)*(1-colorWeight)))
      b_b = Int((Int(base_b * (d_w/255))*colorWeight) + (Int(d_w)*(1-colorWeight)))

      If b_r > 255 Then b_r = 255
      If b_g > 255 Then b_g = 255
      If b_b > 255 Then b_b = 255
      gBOT(s).color = b_r + (b_g * 256) + (b_b * 256 * 256)

      'debug.print "color brightness:" & int(d_w/255) & " r:" & b_r & " g:" & b_g & " b:" & b_b & " color weight:" & colorWeight & " - " & 1-colorWeight
    Next
  else
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
  end if

End Sub

'****************************
' ZGIC: GI Colors
'****************************

''''''''' GI Color Change options
' reference: https://andi-siess.de/rgb-to-color-temperature/

Dim c2700k: c2700k = rgb(255, 169, 87)
Dim c3000k: c3000k = rgb(255, 180, 107)
Dim c3500k: c3500k = rgb(255, 196, 137)
Dim c5000k: c5000k = rgb(255, 244, 226)

Dim cRedFull: cRedFull = rgb(255,0,0)
Dim cRed: cRed= rgb(255,5,5)
Dim cRed_Ball: cRed_Ball = rgb(255,135,135)

Dim cPinkFull: cPinkFull = rgb(255,0,225)
Dim cPink: cPink = rgb(255,5,255)
Dim cPink_Ball: cPink_Ball = rgb(225,150,225)

Dim cWhiteFull: cWhiteFull = rgb(255,255,128)
Dim cWhite: cWhite = rgb(255,255,255)
Dim cWhite_Ball: cWhite_Ball = rgb(255,255,255)

Dim cBlueFull: cBlueFull= rgb(0,0,255)
Dim cBlue: cBlue = rgb(81,209,246)
Dim cBlue_Ball: cBlue_Ball = rgb(120,180,190)

Dim cCyanFull: cCyanFull= rgb(0,255,255)
Dim cCyan : cCyan = rgb(5,128,255)
Dim cCyan_Ball : cCyan_Ball = rgb(15,128,255)

Dim cYellowFull:cYellowFull  = rgb(255,255,128)
Dim cYellow: cYellow = rgb(255,255,0)
Dim cYellow_Ball: cYellow_Ball = rgb(200,200,85)

Dim cOrangeFull: cOrangeFull = rgb(255,128,0)
Dim cOrange: cOrange = rgb(255,70,5)
Dim cOrange_Ball: cOrange_Ball = rgb(225,140,65)

Dim cGreenFull: cGreenFull = rgb (2,255,35) '(0,255,64)
Dim cGreen: cGreen = rgb (2,255,35)
Dim cGreen_Ball: cGreen_Ball = rgb (155,205,160)

Dim cGreenLedFull: cGreenLedFull = rgb(5,255,5)
Dim cGreenLed : cGreenLed = rgb (5,255,5)
Dim cGreenLed_Ball : cGreenLed_Ball = rgb (155,225,160)

Dim cPurpleFull: cPurpleFull = rgb(128,0,255)
Dim cPurple: cPurple = rgb(60,5,255)
Dim cPurple_Ball: cPurple_Ball = rgb(180,150,255)

Dim cAmberFull: cAmberFull = rgb(255,197,143)
Dim cAmber: cAmber = rgb(255,197,143)
Dim cAmber_Ball: cAmber_Ball = rgb(255,197,143)

Dim cArray
Dim CurrentColorTempIndex : CurrentColorTempIndex = 1
cArray = Array(c2700k,c3000k,c3500k,c5000k)

Dim gArray,bArray
gArray = Array(cGreen, cGreenLed, cBlue, cPurple, cPink, cRed, cOrange, cYellow)
bArray = Array(cGreen_Ball, cGreenLed_Ball, cBlue_Ball, cPurple_Ball, cPink_Ball, cRed_Ball, cOrange_Ball, cYellow_Ball)

Dim optionArray
optionArray = Array("Green", "Green LED", "Blue", "Purple", "Pink", "Red", "Orange", "Yellow", "Random")

Dim RandomGIColor : RandomGIColor = False
Dim CurrentGIColorIndex : CurrentGIColorIndex = 0 ' keep track of current color for ball tint

sub SetGiColor(c)
  Dim xx, BL, colorOption
  if optionArray(c) = "Random" Then
    RandomGIColor = True
    CurrentGIColorIndex = RandomColorIndex
  Else
    CurrentGIColorIndex = c
  End If
  colorOption = gArray(CurrentGIColorIndex)

  For each xx in GI_Green
    xx.color = colorOption
    xx.colorfull = colorOption
  Next
  For each BL in BL_GIG: BL.color = colorOption : Next
  For each BL in BL_GIGS_GI_G_001: BL.color = colorOption : Next
  For each BL in BL_GIGS_GI_G_002: BL.color = colorOption : Next
  For each BL in BL_GIGS_GI_G_020: BL.color = colorOption : Next
  For each BL in BL_GIGS_GI_G_022: BL.color = colorOption : Next
  For each BL in BL_GIGS_GI_G_023: BL.color = colorOption : Next
  For each BL in BL_GIGS_GI_G_024: BL.color = colorOption : Next

  colorOption = cArray(CurrentColorTempIndex)

  For each xx in GI
    xx.color = colorOption
    xx.colorfull = colorOption
  Next
  For each BL in BL_GIN: BL.color = colorOption : Next
  For each BL in BL_GINS_gi01: BL.color = colorOption : Next
  For each BL in BL_GINS_gi02: BL.color = colorOption : Next
  For each BL in BL_GINS_gi24: BL.color = colorOption : Next
  For each BL in BL_GINS_gi25: BL.color = colorOption : Next
  For each BL in BL_GINS_gi26: BL.color = colorOption : Next
end Sub

Function RandomColorIndex()
  RandomColorIndex = Int((UBound(gArray) - LBound(gArray) + 1) * Rnd + LBound(gArray))
End Function


'****************************
'   ZRBR: Room Brightness
'****************************

'This code only applies to lightmapped tables. It is here for reference.
'NOTE: Objects bightness will be affected by the Day/Night slider only if their blenddisablelighting property is less than 1.
'      Lightmapped table primitives have their blenddisablelighting equal to 1, therefore we need this SetRoomBrightness sub
'      to handle updating their effective ambient brighness.

' Update these arrays if you want to change more materials with room light level
Dim RoomBrightnessMtlArray: RoomBrightnessMtlArray = Array("VLM.Bake.Active","VLM.Bake.Solid","VLM.Bake.Metal","VLM.Bake.Caps","VLM.Bake.UpperPF","VLM.Bake.intDomes","VLM.Bake.extDomes")

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
' ZKEY: Key Press Handling
'******************************************************

Sub Table1_KeyDown(ByVal keycode) '***What to do when a button is pressed***
  If Keycode = LeftFlipperKey Then
    FlipperActivate LeftFlipper, LFPress
    VRFlipperLeft.x = 2110.8 ' VR flipper animation
  End If
  If Keycode = RightFlipperKey Then
    FlipperActivate RightFlipper, RFPress
    AlienCount = AlienCount + 1
    VRFlipperRight.x = 2095.84  ' VR alien counter and moving VR right flipper
    if AlienCount = 40 then MoveAlien()
  End If
  If keycode = PlungerKey Then Plunger.Pullback:TimerVRPlunger.Enabled = True:TimerVRPlunger2.Enabled = False: End If   'added VR Plunger code here
  If keycode = LeftTiltKey Then Nudge 90, 2 : SoundNudgeLeft      ' Sets the nudge angle and power
  If keycode = RightTiltKey Then Nudge 270, 2 : SoundNudgeRight   ' ^
  If keycode = CenterTiltKey Then Nudge 0, 2 : SoundNudgeCenter   ' ^
  If keycode = StartGameKey Then SoundStartButton:Primary_startbuttoninner.y = Primary_startbuttoninner.y -2:Primary_startbutton.y = Primary_startbutton.y -2
  If keycode = AddCreditKey or keycode = AddCreditKey2 Then
    Select Case Int(rnd*3)
      Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
      Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
      Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
    End Select
  End If
  If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal keycode)   '***What to do when a button is released***
  If Keycode = LeftFlipperKey Then FlipperDeActivate LeftFlipper, LFPress: VRFlipperLeft.x = 2100.8   ' VR flipper animation
  If Keycode = RightFlipperKey Then FlipperDeActivate RightFlipper, RFPress: VRFlipperRight.x =2105.84       ' VR flipper animation
  If Keycode = StartGameKey Then
    Controller.Switch(3) = 0
    Primary_startbuttoninner.y = Primary_startbuttoninner.y +2:Primary_startbutton.y = Primary_startbutton.y -2
  End If
  If keycode = PlungerKey Then
    Plunger.Fire
    If BIPL = True Then             'If true then ball in shooter lane, else no ball is shooter lane
      SoundPlungerPullStop()
      SoundPlungerReleaseBall()     'Plunger release sound when there is a ball in shooter lane
    Else
      SoundPlungerPullStop()
      SoundPlungerReleaseNoBall()     'Plunger release sound when there is no ball in shooter lane
    End If
        TimerVRPlunger.Enabled = False
        TimerVRPlunger2.Enabled = True   ' VR Plunger
    VRPlunger.Y = 2149
  End If
  If vpmKeyUp(keycode) Then Exit Sub
End Sub


'******************************************************
' ZSOL: Solenoids & Flashers
'******************************************************

SolCallback(1) = "SolTrough"      'Outhole Kicker (Drain)
SolCallback(2) = "SolRelease"     'Ball Shooter Lane Feeder
SolCallback(3) = "KickerUpperLeft"    'Left Ball Popper (VUK)
SolCallback(4) = "KickerUpperRight"   'Right Ball Popper (VUK)
'SolCallback(5) = ""    'Unused
SolCallback(6) = "ResetDrops3"      '3 Bank Drop Target Raise
SolCallback(7) = "SolKnocker"
SolCallback(8) = "ResetDrops1"      '1 Bank Drop Target Raise

SolModCallback(9) ="GIUpdates"      'GI Playfield
SolCallback(10) = "GiSelect"      'GI Playfield Green
SolCallback(11) = "GiBackbox"     'GI Backbox

SolCallback(12) = "SolACSelectRelay"  'A/C Select Relay
SolCallback(13) = "SolKickBack"     'Left Re-Entry Kickback
'SolCallback(14) = ""         'Unused
SolCallback(16) = "StationMotor"    'Space Station Motor Relay
SolCallback(17) = "RightDockKick"   'Right Dock Kickback
'18 - 22 Slings and Pop Bumpers
SolCallback(32) = "LeftDockKick"    'Right Dock Kickback

'Flashers
SolModcallback(15) = "SolMod15"     'P/f Top Panel Flashers (x3 pf - 28v Lamps)
SolModcallback(25) = "SolMod25"     'Relaunch (x2 pf) + "ON" Flashers (x2 Backbox)
SolModcallback(26) = "SolMod26"     'Left Side (x2 pf) + "SP" Flashers (x2 Backbox)
SolModcallback(27) = "SolMod27"     'Right Side (x2 pf) + "AC" Flashers (x2 Backbox)
SolModcallback(28) = "SolMod28"     'Top Upper Playfield (x2 pf) + "ES" Flashers (x2 Backbox)
SolModcallback(29) = "SolMod29"     'Playfield Top Panel (x2pf) + "TA" Flashers (x2 Backbox)
SolModcallback(30) = "SolMod30"     'Flame + "TI" Flashers (x4 Backbox)
SolModcallback(31) = "SolMod31"     'Station Flashers (x2 Backbox)

SolCallback(sLRFlipper) = "SolRFlipper" 'Right Flipper
SolCallback(sLLFlipper) = "SolLFlipper" 'Left Flipper

'*** Flasher Subs

Sub SolMod15(level)
  If DebugFlashers=True Then debug.print "SolMod15 => "&level
  f15.state = level
  f15b.state = level
  f15c.state = level
End Sub

Sub SolMod25(level)
  If DebugFlashers=True Then debug.print "SolMod25 => "&level
  f25.state = level
  f25b.state = level
End Sub

Sub SolMod26(level)
  If DebugFlashers=True Then debug.print "SolMod26 => "&level
  f26.state = level
  f26b.state = level
End Sub

Sub SolMod27(level)
  If DebugFlashers=True Then debug.print "SolMod27 => "&level
  f27.state = level
  f27b.state = level
End Sub

Sub SolMod28(level)
  If DebugFlashers=True Then debug.print "SolMod28 => "&level
  f28.state = level
  f28b.state = level
End Sub

Sub SolMod29(level)
  If DebugFlashers=True Then debug.print "SolMod29 => "&level
  f29.state = level
  f29b.state = level
End Sub

Sub SolMod30(level)
  If DebugFlashers=True Then debug.print "SolMod30 => "&level
  f30.state = level
End Sub

Sub SolMod31(level)
  If DebugFlashers=True Then debug.print "SolMod31 => "&level
  f31.state = level
End Sub

'******************************************************
' ZDRN: Drain, Trough, and Ball Release
'******************************************************

'********************* TROUGH *************************

Sub sw13_Hit   : Controller.Switch(13) = 1 : UpdateTrough : End Sub
Sub sw13_UnHit : Controller.Switch(13) = 0 : UpdateTrough : End Sub
Sub sw12_Hit   : Controller.Switch(12) = 1 : UpdateTrough : End Sub
Sub sw12_UnHit : Controller.Switch(12) = 0 : UpdateTrough : End Sub
Sub sw11_Hit   : Controller.Switch(11) = 1 : UpdateTrough : End Sub
Sub sw11_UnHit : Controller.Switch(11) = 0 : UpdateTrough : End Sub
Sub sw10_Hit   : Controller.Switch(10) = 1 : UpdateTrough : RandomSoundDrain sw10 : End Sub
Sub sw10_UnHit : Controller.Switch(10) = 0 : UpdateTrough : End Sub

Sub UpdateTrough
  UpdateTroughTimer.Interval = 300
  UpdateTroughTimer.Enabled = 1
End Sub

Sub UpdateTroughTimer_Timer
  If sw12.BallCntOver = 0 Then sw13.kick 57, 10
  If sw11.BallCntOver = 0 Then sw12.kick 57, 10
  UpdateTroughTimer.Enabled = 0
End Sub

'*****************  DRAIN & RELEASE  ******************


Sub SolTrough(enabled)
  If enabled Then
    sw10.kick 57, 10
    RandomSoundDrainKicker sw10
  End If
End Sub

Sub SolRelease(enabled)
  If enabled Then
    sw11.kick 57, 10
    RandomSoundBallRelease sw11
  End If
End Sub


'******************************************************
' ZTOY: TOYS
'******************************************************
'SolCallback(16) = "StationMotor"

Dim MotorOn : MotorOn = False
dim MotorPos : MotorPos = 90
dim MotorSpeed : MotorSpeed = 0.054
dim MotorSoundON:MotorSoundON = 0

Sub StationMotor(aOn)
  Select Case aOn
    Case True : MotorOn = True
    Case False : MotorOn = False : MotorSound False
  End Select
End Sub

Sub MotorUpdate() 'Should go counter-clockwise...
  If MotorOn then
    MotorSound True
    MotorPos = MotorPos + MotorSpeed * FrameTime
    If MotorPos > 360 then MotorPos = MotorPos - 360
    If MotorPos < 0 then MotorPos = MotorPos + 360
    If MotorPos > 180 then Controller.Switch(52) = 1 Else Controller.Switch(52) = 0 End If
    If MotorPos > 90 and  MotorPos < 270 then Controller.Switch(53) = 1 Else Controller.Switch(53) = 0 End If
    UpdateUFO MotorPos
  End If
End Sub

Sub MotorSound(aEnabled)
  If aEnabled Then
    If Not MotorSoundON then PlaySound SoundFX("Motor", DOFGear), -1, 0.1, AudioPan(Phys_Div_NS) , 0, 0, 1, AudioPan(Phys_Div_NS)
  Else
    StopSound "Motor"
  End If
  MotorSoundON = aEnabled
End Sub


'Initialize UFO
InitUFO
Sub InitUFO
  Dim BP
  For Each BP in BP_SS_01: BP.visible = 1: BP.Size_X = 1.005:  BP.Size_Y = 1.005:  BP.Size_Z = 1.005: Next
  For Each BP in BP_SS_02: BP.visible = 1: BP.Size_X = 1.01:  BP.Size_Y = 1.01:  BP.Size_Z = 1.01: Next
  For Each BP in BP_SS_03: BP.visible = 1: BP.Size_X = 1.005:  BP.Size_Y = 1.005:  BP.Size_Z = 1.005: Next
  For Each BP in BP_SS_04: BP.visible = 1: BP.Size_X = 1.01:  BP.Size_Y = 1.01:  BP.Size_Z = 1.01: Next
  BM_SS_01.Size_X = 1: BM_SS_01.Size_Y = 1: BM_SS_01.Size_Z = 1
  BM_SS_02.visible = 0: BM_SS_03.visible = 0: BM_SS_04.visible = 0

  For Each BP in BP_diverter01: BP.visible = 1: BP.Size_X = 1.005:  BP.Size_Y = 1.005:  BP.Size_Z = 1.005: Next
  For Each BP in BP_diverter02: BP.visible = 1: BP.Size_X = 1.01:  BP.Size_Y = 1.01:  BP.Size_Z = 1.01: Next
  For Each BP in BP_diverter03: BP.visible = 1: BP.Size_X = 1.005:  BP.Size_Y = 1.005:  BP.Size_Z = 1.005: Next
  For Each BP in BP_diverter04: BP.visible = 1: BP.Size_X = 1.01:  BP.Size_Y = 1.01:  BP.Size_Z = 1.01: Next
  BM_diverter01.Size_X = 1: BM_diverter01.Size_Y = 1: BM_diverter01.Size_Z = 1
  BM_diverter02.visible = 0: BM_diverter03.visible = 0: BM_diverter04.visible = 0

  UpdateUFO MotorPos
End Sub


'Update UFO
Sub UpdateUFO(aNewPos)
  Dim BP
  Dim Op1, Op2, Op3, Op4

  'Update opacities
  Op1 = 100*max(0,dCOS(aNewPos))
  Op2 = 100*max(0,dCOS(aNewPos-90))
  Op3 = 100*max(0,dCOS(aNewPos-180))
  Op4 = 100*max(0,dCOS(aNewPos-270))

  'Apply updates
  For Each BP in BP_SS_01: BP.RotZ = aNewPos * (-1): BP.opacity = Op1: Next
  BM_SS_01.opacity = 100
  For Each BP in BP_SS_02: BP.RotZ = aNewPos * (-1): BP.opacity = Op2: Next
  For Each BP in BP_SS_03: BP.RotZ = aNewPos * (-1): BP.opacity = Op3: Next
  For Each BP in BP_SS_04: BP.RotZ = aNewPos * (-1): BP.opacity = Op4: Next

  For Each BP in BP_diverter01: BP.RotZ = aNewPos * (-1): BP.opacity = Op1/2: Next
  BM_diverter01.opacity = 100
  For Each BP in BP_diverter02: BP.RotZ = aNewPos * (-1): BP.opacity = Op2/2: Next
  For Each BP in BP_diverter03: BP.RotZ = aNewPos * (-1): BP.opacity = Op3/2: Next
  For Each BP in BP_diverter04: BP.RotZ = aNewPos * (-1): BP.opacity = Op4/2: Next

  'debug.print aNewPos&","&SSOp1&","&SSOp2&","&DivOp1&","&DivOp2&","&DivOp3&","&DivOp4
  'Update collidables
  if AnewPos > 350 then GoLeft : Exit Sub
  If AnewPos < 10 Then GoLeft : Exit Sub
  if anewpos > 80 and anewpos < 100 then GoRight : Exit Sub
  if AnewPos > 170 and AnewPos < 190 then GoLeft : Exit Sub
  if AnewPos > 260 and AnewPos < 280 then GoRight : Exit Sub
  If ANewPos > 290 then GoRight : Exit Sub
  StationBlocker.isdropped = 0
End Sub



Sub GoLeft()
  'debug.print "Go Left (0 180)"
  Phys_Div_NS.collidable=0:Phys_Div_EW.collidable=1
  StationBlocker.isdropped = 1
End Sub

Sub GoRight()
  'debug.print "Go Right"
  Phys_Div_NS.collidable=1:Phys_Div_EW.collidable=0
  StationBlocker.isdropped = 1
End Sub

'******************************************************
' ZFLP: FLIPPERS
'******************************************************

Const ReflipAngle = 20
Const QuickFlipAngle = 20

Sub SolLFlipper(Enabled)
     If Enabled Then
        LF.Fire 'LeftFlipper.RotateToEnd
    FlipperActivate LeftFlipper, LFPress

    If LeftFlipper.currentangle < LeftFlipper.endangle + ReflipAngle Then
      StopAnyFlipperLowerLeftDown()
      RandomSoundFlipperLowerLeftReflip LeftFlipper
    Else
      If BallNearLF = 0 Then
        RandomSoundFlipperLowerLeftUpFullStroke LeftFlipper
      End If
      If BallNearLF = 1 Then
        Select Case Int(Rnd*2)+1
          Case 1 : RandomSoundFlipperLowerLeftUpDampenedStroke LeftFlipper
          Case 2 : RandomSoundFlipperLowerLeftUpFullStroke LeftFlipper
        End Select
      End If
    End If
     Else
    LeftFlipper.RotateToStart
    FlipperDeActivate LeftFlipper, LFPress

    If LeftFlipper.currentangle < LeftFlipper.startAngle - 5 Then
      RandomSoundFlipperLowerLeftDown LeftFlipper
    End If
    If LeftFlipper.currentangle < LeftFlipper.startAngle + QuickFlipAngle and LeftFlipper.currentangle <> LeftFlipper.endangle Then
      StopAnyFlipperLowerLeftUp()
      RandomSoundLowerLeftQuickFlipUp()
    Else
      FlipperLeftLowerHitParm = FlipperUpSoundLevel
    End If
     End If
 End Sub

Sub SolRFlipper(Enabled)
    If enabled then
    FlipperActivate RightFlipper, RFPress
        RF.Fire 'RightFlipper.RotateToEnd

    If RightFlipper.currentangle > RightFlipper.endangle - ReflipAngle Then
      StopAnyFlipperLowerRightDown()
      RandomSoundFlipperLowerRightReflip RightFlipper
    Else
      If BallNearRF = 0 Then
        RandomSoundFlipperLowerRightUpFullStroke RightFlipper
      End If

      If BallNearRF = 1 Then
        Select Case Int(Rnd*2)+1
          Case 1 : RandomSoundFlipperLowerRightUpDampenedStroke RightFlipper
          Case 2 : RandomSoundFlipperLowerRightUpFullStroke RightFlipper
        End Select
      End If
    End If
     Else
        RightFlipper.RotateToStart
    FlipperDeActivate RightFlipper, RFPress

    If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
      RandomSoundFlipperLowerRightDown RightFlipper
    End If
    If RightFlipper.currentangle < RightFlipper.startAngle + QuickFlipAngle and RightFlipper.currentangle <> RightFlipper.endangle Then
      StopAnyFlipperLowerRightUp()
      RandomSoundLowerRightQuickFlipUp()
    Else
      FlipperRightLowerHitParm = FlipperUpSoundLevel
    End If
     End If
 End Sub

'///////////////////////  FLIPPER BATS BALL COLLIDE SOUND  //////////////////////
dim angdamp, veldamp
angdamp = 0.2
veldamp = 0.8

Sub LeftFlipper_Collide(parm)
  CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
  LF.ReProcessBalls ActiveBall
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
  RF.ReProcessBalls ActiveBall
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

'Impulse Plunger
Sub SolAutofire(Enabled)
  If Enabled Then
    PlungerIM.AutoFire
    'SoundPlungerReleaseBall
  End If
End Sub

Const IMPowerSetting = 65
Const IMTime = 0.6
Set plungerIM = New cvpmImpulseP
With plungerIM
    .InitImpulseP Sw43, IMPowerSetting, IMTime
    .Random 0.3
    .CreateEvents "plungerIM"
End With


'************************************************************
' ZSLG: Slingshot Animations
'************************************************************
Dim LStep : LStep = 4 : LeftSlingShot_Timer
Dim RStep : RStep = 4 : RightSlingShot_Timer

Sub RightSlingShot_Slingshot
    RS.VelocityCorrect(Activeball)
    vpmTimer.PulseSw(64)                        'Sling Switch Number
    'RSling1.Visible = 1
    'Sling1.TransY =  - 20                         'Sling Metal Bracket
    RStep = 0
    RightSlingShot.TimerInterval = 10
    RightSlingShot.TimerEnabled = 1
    RandomSoundSlingshotRight Rubber_Post_004
End Sub


Sub RightSlingShot_Timer
    Dim BL
    Dim x1, x2, y: x1 = True:x2 = False:y = 25
    Select Case RStep
        Case 3:x1 = False: x2 = True:y = 15
        Case 4:x1 = False: x2 = False:y = 0: RightSlingShot.TimerEnabled = 0
    End Select
    For Each BL in BP_RSling1 : BL.Visible = x1: Next
    For Each BL in BP_RSling2 : BL.Visible = x2: Next
    For Each BL in BP_REMK : BL.transx = y: Next
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
    LS.VelocityCorrect(Activeball)
    vpmTimer.PulseSw(63)                        'Sling Switch Number
    'LSling1.Visible = 1
    'Sling2.TransY =  - 20                          'Sling Metal Bracket
    LStep = 0
    LeftSlingShot.TimerInterval = 10
    LeftSlingShot.TimerEnabled = 1
    RandomSoundSlingshotLeft Rubber_Post_001
End Sub

Sub LeftSlingShot_Timer
    Dim BL
    Dim x1, x2, y: x1 = True:x2 = False:y = 25
    Select Case LStep
        Case 3:x1 = False: x2= True: y = 15
        Case 4:x1 = False: x2 = False: y = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    For Each BL in BP_LSling1 : BL.Visible = x1: Next
    For Each BL in BP_LSling2 : BL.Visible = x2: Next
    For Each BL in BP_LEMK : BL.transx = y: Next
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

'Rollovers
Sub sw14_Hit:Controller.Switch(14) = 1:End Sub    'Upper Playfield
Sub sw14_UnHit:Controller.Switch(14) = 0:End Sub
Sub sw15_Hit:Controller.Switch(15) = 1:End Sub    'Upper Playfield
Sub sw15_UnHit:Controller.Switch(15) = 0:End Sub
Sub sw16_Hit:Controller.Switch(16) = 1:End Sub    'Upper Playfield
Sub sw16_UnHit:Controller.Switch(16) = 0:End Sub

Sub sw17_Hit:Controller.Switch(17) = 1:RandomSoundOutlaneRollover:End Sub   'Left Inlane
Sub sw17_UnHit:Controller.Switch(17) = 0:End Sub
Sub sw32_Hit:Controller.Switch(32) = 1:RandomSoundOutlaneRollover:End Sub   'Right Inlane
Sub sw32_UnHit:Controller.Switch(32) = 0:End Sub

Sub sw38_Hit:Controller.Switch(38) = 1:End Sub    'USA Rollover
Sub sw38_UnHit:Controller.Switch(38) = 0:End Sub
Sub sw39_Hit:Controller.Switch(39) = 1:End Sub    'USA Rollover
Sub sw39_UnHit:Controller.Switch(39) = 0:End Sub
Sub sw40_Hit:Controller.Switch(40) = 1:End Sub    'USA Rollover
Sub sw40_UnHit:Controller.Switch(40) = 0:End Sub

Sub sw43_Hit:Controller.Switch(43) = 1:BIPL = True:End Sub    'Shooter Lane
Sub sw43_UnHit:Controller.Switch(43) = 0:BIPL = False:End Sub
Sub sw46_Hit:Controller.Switch(46) = 1:End Sub    'Right Lock Ramp Entry
Sub sw46_UnHit:Controller.Switch(46) = 0:End Sub

'Rollunders
Sub sw37_Hit:Controller.Switch(37) = 1:End Sub    'Left Orbit
Sub sw37_UnHit:Controller.Switch(37) = 0:End Sub

Sub sw48_Hit:Controller.Switch(48) = 1:End Sub    'Left Lock Entry
Sub sw48_UnHit:Controller.Switch(48) = 0:End Sub

'Upper Rubber Switches
Sub sw50_rubber_Hit:vpmTimer.PulseSw(50):End Sub    'Lower Right Rubber Switch
Sub sw54_rubber_Hit:vpmTimer.PulseSw(54):End Sub    'Uppper Right Rubber Switch


'************************* Bumpers **************************
Sub Bumper1_Hit(): RandomSoundBumperTop Bumper1: vpmTimer.PulseSw 60: End Sub
Sub Bumper2_Hit(): RandomSoundBumperMiddle Bumper2: vpmTimer.PulseSw 61: End Sub
Sub Bumper3_Hit(): RandomSoundBumperBottom Bumper3: vpmTimer.PulseSw 62: End Sub


'********************* Standup Targets **********************
Sub sw18_hit: STHit 18: End Sub   'S
Sub sw19_hit: STHit 19: End Sub   'H
Sub sw20_hit: STHit 20: End Sub   'U
Sub sw21_hit: STHit 21: End Sub   'T
Sub sw22_hit: STHit 22: End Sub   'T
Sub sw23_hit: STHit 23: End Sub   'L
Sub sw24_hit: STHit 24: End Sub   'E

Sub sw25_hit: STHit 25: End Sub   'S
Sub sw26_hit: STHit 26: End Sub   'T
Sub sw27_hit: STHit 27: End Sub   'A

Sub sw28_hit: STHit 28: End Sub   'T
Sub sw29_hit: STHit 29: End Sub   'I
Sub sw30_hit: STHit 30: End Sub   'O
Sub sw31_hit: STHit 31: End Sub   'N

Sub sw35_hit: STHit 35: End Sub   'Change

'********************** Drop Targets ************************
Sub sw33_Hit: DTHit 33: TargetBouncer Activeball, 1.5: End Sub

Sub ResetDrops1(enabled)
  if enabled then
    RandomSoundDropTargetReset BM_sw33
    DTRaise 33
  end if
End Sub

Sub sw57_Hit: DTHit 57: End Sub
Sub sw58_Hit: DTHit 58: End Sub
Sub sw59_Hit: DTHit 59: End Sub

Sub ResetDrops3(enabled)
  if enabled then
    RandomSoundDropTargetBankReset BM_sw58
    DTRaise 57
    DTRaise 58
    DTRaise 59
  end if
End Sub


'************************************************************
' ZVUK: VUKs and Kickers
'************************************************************

'******************* VUKs **********************

Dim KickerBall41, KickerBall42, KickerBall45, KickerBall47    'Each VUK needs its own "kickerball"

Sub KickBall(kball, kangle, kvel, kvelz, kzlift)  'Defines how KickBall works
  dim rangle
  rangle = PI * (kangle - 90) / 180

  kball.z = kball.z + kzlift
  kball.velz = kvelz
  kball.velx = cos(rangle)*kvel
  kball.vely = sin(rangle)*kvel
End Sub

'Upper Left VUK
Sub sw41_Hit                    'Switch associated with Kicker
  WireRampOn False
    set KickerBall41 = activeball
    Controller.Switch(41) = 1
    SoundSaucerLock
End Sub

Sub sw41_unhit
  KickerBall41 = Empty
  Controller.Switch(41) = 0
End Sub

Sub KickerUpperLeft(Enable)             'Solonoid name associated with kicker.
    If Enable then
    If Not IsEmpty(KickerBall41) Then
      KickBall KickerBall41, 0, 0, 30, 10
      SoundSaucerKick 1, sw41
    Else
      SoundSaucerKick 0, sw41
    End If
  End If
End Sub

'Upper Right VUK
Sub sw42_Hit
  WireRampOn False
    set KickerBall42 = activeball
    Controller.Switch(42) = 1
    SoundSaucerLock
End Sub

Sub sw42_unhit
  KickerBall42 = Empty
  Controller.Switch(42) = 0
End Sub

Sub KickerUpperRight(Enable)
    If Enable then
    If Not IsEmpty(KickerBall42) Then
      KickBall KickerBall42, 270, 0, 60, 10
      SoundSaucerKick 1, sw42
    Else
      SoundSaucerKick 0, sw42
    End If
  End If
End Sub

'Middle Right VUK
Sub sw45_Hit
    set KickerBall45 = activeball
    Controller.Switch(45) = 1
End Sub

Sub sw45_unhit
  KickerBall45 = Empty
  Controller.Switch(45) = 0
End Sub

'Sub sw45k_Hit

'End Sub

'Sub sw45k_unhit

'End Sub

Sub RightDockKick(Enable)
    If Enable then
    If Not IsEmpty(KickerBall45) Then
      KickBall KickerBall45, 0, RndNum(16,24), 0, 0
      SoundSaucerKick 1, sw45
    Else
      SoundSaucerKick 0, sw45
    End If
  End If
End Sub

'Middle Left VUK
Sub sw47_Hit
  WireRampOff
    set KickerBall47 = activeball
    Controller.Switch(47) = 1
End Sub

Sub sw47_unhit
  KickerBall47 = Empty
  Controller.Switch(47) = 0
End Sub

Sub LeftDockKick(Enable)
    If Enable then
    If Not IsEmpty(KickerBall47) Then
      KickBall KickerBall47, 0, RndNum(16,24), 0, 0
      SoundSaucerKick 1, sw47
    Else
      SoundSaucerKick 0, sw47
    End If
  End If
End Sub

'******************* Kicker **********************
Sub SolKickBack(enabled)
    If enabled Then
    SoundSaucerKick 1, Kickback
    Kickback.Fire
  Else
    Kickback.PullBack
  End If
End Sub

'*******************************************
' ZSOL: Other Solenoids
'*******************************************

' Knocker (this sub mimics how you would handle kicker in ROM based tables)
' For this to work, you must create a primitive on the table named KnockerPosition
' SolCallback(XX) = "SolKnocker"  'In ROM based tables, change the solenoid number XX to the correct number for your table.
Sub SolKnocker(Enabled) 'Knocker solenoid
  If enabled Then
    KnockerSolenoid 'Add knocker position object
  End If
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
  Dim b', BOT
' BOT = GetBalls

  ' stop the sound of deleted balls
  For b = UBound(gBOT) + 1 to tnob - 1
    rolling(b) = False
    StopSound(Cartridge_Ball_Roll & "_Ball_Roll_" & b)
  Next

  ' exit the sub if no balls on the table
  If UBound(gBOT) = -1 Then Exit Sub

  ' play the rolling sound for each ball

  For b = 0 to UBound(gBOT)
    If BallVel(gBOT(b)) > 1 AND gBOT(b).z < 30 Then
      rolling(b) = True
      PlaySound (Cartridge_Ball_Roll & "_Ball_Roll_" & b), -1, VolPlayfieldRoll(gBOT(b)) * BallRollVolume * VolumeDial, AudioPan(gBOT(b)), 0, PitchPlayfieldRoll(gBOT(b)), 1, 0, AudioFade(gBOT(b))

    Else
      If rolling(b) = True Then
        StopSound(Cartridge_Ball_Roll & "_Ball_Roll_" & b)
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
dim RampBalls(5,2)
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


'Ramp triggers
Sub RampTrigger1_Hit
  If activeball.vely < 0 Then
    WireRampOn False
  Else
    WireRampOff
  End If
End Sub

Sub RampTrigger2_Hit
  If activeball.vely < 0 Then
    WireRampOn True
  Else
    WireRampOff
  End If
End Sub

Sub RampTrigger3_Hit
    WireRampOff
End Sub

Sub RampTrigger4_Hit
    WireRampOff
End Sub

Sub RampTrigger5_Hit
    WireRampOff
End Sub

Sub RampTrigger5_UnHit
    WireRampOn True
End Sub

Sub RampTrigger6_Hit
    WireRampOff
End Sub

Sub RampTrigger6_UnHit
    WireRampOn True
End Sub

Sub RampTrigger7_Hit
    WireRampOff
End Sub

Sub RampTrigger7_UnHit
    WireRampOn False
End Sub

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

' This part in the script is an entire block that is dedicated to the physics sound system.
' Various scripts and sounds that may be pretty generic and could suit other SYS11 systems, but the most are tailored specifically for the Space Station table.

' Many of the sounds in this package can be added by creating collections and adding the appropriate objects to those collections.
' Create the following new collections:
'  Metals (all metal objects, metal walls, metal posts, metal wire guides)
'  Apron (the apron walls and plunger wall)
'  Walls (all wood or plastic walls)
'  Rollovers (wire rollover triggers, star triggers, or button triggers)
'  Targets (standup or drop targets, these are hit sounds only ... you will want to add separate dropping sounds for drop targets)
'  Gates (plate gates)
'  GatesWire (wire gates)
'    GatesOneWay (one way wire gates)
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

' Updated 2025 by RobbyKingPin. This edited version contains almost exactly the same structure and subroutines as the original Fleep Sound codes but with the addition of the cartrigdes used inside the updated Fleep from 2022.

'/////////////////////////////////  CARTRIDGES  /////////////////////////////////
'
'//  Specify which mechanical sound cartridge to use for each group of elements.
'//  Mechanical sounds naming convention: <CARTRIDGE>_<Soundset_Name>
'//
'//  Cartridge name is composed using the following convention:
'//  <TABLE MANUFACTURER ABBREVIATION>_<TABLE NAME ABBREVIATION>_<SOUNDSET REVISION NUMBER>
'//
'//  General Mechanical Sounds Cartridges:

Const Cartridge_Bumpers         = "WS_PBT_REV01" 'Williams Pinbot Cartridge REV01
Const Cartridge_Slingshots        = "WS_PBT_REV01" 'Williams Pinbot Cartridge REV01
Const Cartridge_Flippers        = "WS_PBT_REV01" 'Williams Pinbot Cartridge REV01
Const Cartridge_Knocker         = "WS_WHD_REV02" 'Williams Whirlwind Cartridge REV02
Const Cartridge_Relays          = "WS_WHD_REV01" 'Williams Whirlwind Cartridge REV01
Const Cartridge_Trough          = "WS_WHD_REV01" 'Williams Whirlwind Cartridge REV01
Const Cartridge_Rollovers       = "WS_WHD_REV01" 'Williams Whirlwind Cartridge REV01
Const Cartridge_Targets         = "WS_WHD_REV01" 'Williams Whirlwind Cartridge REV01
Const Cartridge_Gates         = "WS_WHD_REV01" 'Williams Whirlwind Cartridge REV01
Const Cartridge_Metal_Hits        = "WS_WHD_REV01" 'Williams Whirlwind Cartridge REV01
Const Cartridge_Rubber_Hits       = "WS_WHD_REV01" 'Williams Whirlwind Cartridge REV01
Const Cartridge_Cabinet_Sounds      = "WS_WHD_REV01" 'Williams Whirlwind Cartridge REV01
Const Cartridge_Apron         = "WS_WHD_REV01" 'Williams Whirlwind Cartridge REV01
Const Cartridge_Ball_Roll       = "BY_TOM_REV01" 'Bally Theatre of Magic Cartridge REV01
Const Cartridge_BallBallCollision   = "BY_WDT_REV01" 'Bally WHO Dunnit Cartridge REV01
Const Cartridge_Ball_Drop_Bump      = "WS_WHD_REV01" 'Williams Whirlwind Cartridge REV01

'////////////////////////////  SOUND SOURCE CREDITS  ////////////////////////////
'//  Special thanks go to the following contributors who have provided audio
'//  footage recordings:
'//
'//  Williams Whirlwind - Blackmoor, wrd1972
'//  Williams Diner - Nick Rusis
'//  Spooky Total Nuclear Annihilation - WildDogArcade, Ed and Gary
'//  Bally Theatre of Magic - CalleV, nickbuol
'//  Bally WHO Dunnit - Amazaley1
'//  Williams Pinbot - major_drain_pinball

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
'//  FLIPPER COILS:
'//  Flippers in this table: Lower Left Flipper, Lower Right Flipper, Upper Right Fliiper
Dim FlipperUpSoundLevel, FlipperDownSoundLevel, FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel
Dim FlipperLeftLowerHitParm, FlipperRightLowerHitParm

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
Dim FlipperLeftHitParm, FlipperRightHitParm
Dim SlingshotSoundLevel, BumperSoundFactor, KnockerSoundLevel

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
    AudioFade = CSng(tmp ^ 10)
  Else
    AudioFade = CSng( - (( - tmp) ^ 10) )
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
    AudioPan = CSng(tmp ^ 10)
  Else
    AudioPan = CSng( - (( - tmp) ^ 10) )
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
  PlaySound (Cartridge_Cabinet_Sounds & "_Start_Button"), 0, StartButtonSoundLevel, 0, 0.25
End Sub

Sub SoundNudgeLeft()
  PlaySound ("Nudge_" & Int(Rnd*3)+1), 0, NudgeLeftSoundLevel * VolumeDial, -0.1, 0.25
End Sub

Sub SoundNudgeRight()
  PlaySound ("Nudge_" & Int(Rnd*3)+1), 0, NudgeRightSoundLevel * VolumeDial, 0.1, 0.25
End Sub

Sub SoundNudgeCenter()
  PlaySound ("Nudge_" & Int(Rnd*3)+1), 0, NudgeRightSoundLevel * VolumeDial, 0.1, 0.25
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

'/////////////////////////////  KNOCKER SOLENOID  ////////////////////////////

Sub KnockerSolenoid()
  PlaySoundAtLevelStatic SoundFX(Cartridge_Knocker & "_Knocker_Coil",DOFKnocker), KnockerSoundLevel, KnockerPosition
End Sub

'/////////////////////////////  DRAIN SOUNDS  ////////////////////////////

Sub RandomSoundDrain(drainswitch)
  PlaySoundAtLevelStatic (Cartridge_Trough & "_Outhole_Drain_Hit_" & Int(Rnd*4)+1), DrainSoundLevel, drainswitch
End Sub

Sub RandomSoundDrainKicker(drainswitch)
  PlaySoundAtLevelStatic SoundFX(Cartridge_Trough & "_Outhole_Kicker_" & Int(Rnd*4)+1,DOFContactors), DrainSoundLevel, drainswitch
End Sub

'/////////////////////////////  TROUGH BALL RELEASE SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundBallRelease(drainswitch)
  PlaySoundAtLevelStatic SoundFX(Cartridge_Trough & "_Shooter_Feeder_" & Int(Rnd*6)+1,DOFContactors), BallReleaseSoundLevel, drainswitch
End Sub

'/////////////////////////////  SLINGSHOT SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundSlingshotLeft(sling)
  PlaySoundAtLevelStatic SoundFX(Cartridge_Slingshots & "_Slingshot_Left_" & Int(Rnd*26)+1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

Sub RandomSoundSlingshotRight(sling)
  PlaySoundAtLevelStatic SoundFX(Cartridge_Slingshots & "_Slingshot_Right_" & Int(Rnd*25)+1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

'/////////////////////////////  BUMPER SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundBumperTop(Bump)
  PlaySoundAtLevelStatic SoundFX(Cartridge_Bumpers & "_Jet_Bumper_Left_" & Int(Rnd*22)+1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperMiddle(Bump)
  PlaySoundAtLevelStatic SoundFX(Cartridge_Bumpers & "_Jet_Bumper_Up_" & Int(Rnd*25)+1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperBottom(Bump)
  PlaySoundAtLevelStatic SoundFX(Cartridge_Bumpers & "_Jet_Bumper_Low_" & Int(Rnd*28)+1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

'/////////////////////////////  SPINNER SOUNDS  ////////////////////////////

Sub SoundSpinner(spinnerswitch)
  PlaySoundAtLevelStatic ("Spinner"), SpinnerSoundLevel, spinnerswitch
End Sub

'///////////////////////  FLIPPER BATS SOUND SUBROUTINES  ///////////////////////
'//////////////////////  FLIPPER BATS SOLENOID CORE SOUND  //////////////////////

Sub RandomSoundFlipperLowerLeftUpFullStroke(flipper)
  PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Lower_Left_Up_Full_Stroke_" & RndInt(1,10),DOFFlippers), FlipperLeftLowerHitParm, Flipper
End Sub

Sub RandomSoundFlipperLowerLeftUpDampenedStroke(flipper)
  PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Lower_Left_Up_Dampened_Stroke_" & RndInt(1,23),DOFFlippers), FlipperLeftLowerHitParm * 1.2, Flipper
End Sub

Sub RandomSoundFlipperLowerRightUpFullStroke(flipper)
  PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Lower_Right_Up_Full_Stroke_" & RndInt(1,11),DOFFlippers), FlipperRightLowerHitParm, Flipper
End Sub

Sub RandomSoundFlipperLowerRightUpDampenedStroke(flipper)
  PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Lower_Right_Up_Dampened_Stroke_" & RndInt(1,23),DOFFlippers), FlipperLeftLowerHitParm * 1.2, Flipper
End Sub

Sub RandomSoundFlipperLowerLeftReflip(flipper)
  PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Lower_Left_Reflip_" & RndInt(1,3),DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperLowerRightReflip(flipper)
  PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Lower_Right_Reflip_" & RndInt(1,3),DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperLowerLeftDown(flipper)
  PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Lower_Left_Down_" & RndInt(1,10),DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperLowerRightDown(flipper)
  PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Lower_Right_Down_" & RndInt(1,11),DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

Sub RandomSoundLowerLeftQuickFlipUp()
  PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Lower_Left_QuickFlip_Up_" & RndInt(1,3),DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, LeftFlipper
End Sub

Sub RandomSoundLowerRightQuickFlipUp()
  PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Lower_Right_QuickFlip_Up_" & RndInt(1,3),DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, RightFlipper
End Sub

Sub StopAnyFlipperLowerLeftUp()
  Dim anyFullStrokeSound
  Dim anyDampenedStrokeSound
  For anyFullStrokeSound = 1 to 10
    StopSound(Cartridge_Flippers & "_Flipper_Lower_Left_Up_Full_Stroke_" & anyFullStrokeSound)
  Next
  For anyDampenedStrokeSound = 1 to 23
    StopSound(Cartridge_Flippers & "_Flipper_Lower_Left_Up_Dampened_Stroke_" & anyDampenedStrokeSound)
  Next
End Sub

Sub StopAnyFlipperLowerRightUp()
  Dim anyFullStrokeSound
  Dim anyDampenedStrokeSound
  For anyFullStrokeSound = 1 to 11
    StopSound(Cartridge_Flippers & "_Flipper_Lower_Right_Up_Full_Stroke_" & anyFullStrokeSound)
  Next
  For anyDampenedStrokeSound = 1 to 23
    StopSound(Cartridge_Flippers & "_Flipper_Lower_Right_Up_Dampened_Stroke_" & anyDampenedStrokeSound)
  Next
End Sub

Sub StopAnyFlipperLowerLeftDown()
  Dim anyFullDownSound
  For anyFullDownSound = 1 to 10
    StopSound(Cartridge_Flippers & "_Flipper_Lower_Left_Down_" & anyFullDownSound)
  Next
End Sub

Sub StopAnyFlipperLowerRightDown()
  Dim anyFullDownSound
  For anyFullDownSound = 1 to 10
    StopSound(Cartridge_Flippers & "_Flipper_Lower_Right_Down_" & anyFullDownSound)
  Next
End Sub

Sub RandomSoundRubberFlipper(parm)
  PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Flipper_Hit_" & Int(Rnd*7)+1), parm * RubberFlipperSoundFactor
End Sub

'/////////////////////////////  ROLLOVER SOUNDS  ////////////////////////////

Sub RandomSoundRollover()
  PlaySoundAtLevelActiveBall (Cartridge_Rollovers & "_Rollover_All_" & Int(Rnd*4)+1), RolloverSoundLevel
End Sub

'///////////////////////////  OUTLANE ROLLOVER SOUNDS  //////////////////////////

Sub RandomSoundOutlaneRollover()
  PlaySoundAtLevelActiveBall (Cartridge_Rollovers & "_Rollover_Outlane_" & Int(Rnd*4)+1), RolloverSoundLevel
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
  Select Case Int(Rnd*10)+1
    Case 1 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_1"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 2 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_2"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 3 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_3"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 4 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_4"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 5 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_5"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 6 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_6"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 7 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_7"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 8 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_8"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 9 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_9"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 10 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_10"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6 * voladj
  End Select
End Sub

'/////////////////////////////  RUBBERS AND POSTS - WEAK IMPACTS  ////////////////////////////

Sub RandomSoundRubberWeak()
  PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_" & Int(Rnd*8)+1), Vol(ActiveBall) * RubberWeakSoundFactor
End Sub

'/////////////////////////////  WALL IMPACTS  ////////////////////////////

Sub Walls_Hit(idx)
  RandomSoundWall()
End Sub

Sub RandomSoundWall()
  Dim finalspeed
  finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
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

'/////////////////////////////  METAL TOUCH SOUNDS  ////////////////////////////

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
  PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Arch_Ball_Guide_Hit_Soft_" & Int(Rnd*4)+1), BottomArchBallGuideSoundFactor
End Sub

'/////////////////////////////  BOTTOM ARCH BALL GUIDE - HARD HITS  ////////////////////////////

Sub RandomSoundBottomArchBallGuideHardHit()
  PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Arch_Ball_Guide_Hit_Hard_" & Int(Rnd*3)+1), BottomArchBallGuideSoundFactor * 3
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

'/////////////////////////////  TARGET HIT SOUNDS  ////////////////////////////

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
  PlaySoundAtLevelStatic (Cartridge_Gates & "_Bracket_Gate_1"), GateSoundLevel * 0.5, ActiveBall
End Sub

Sub SoundHeavyGate()
  PlaySoundAtLevelStatic (Cartridge_Gates & "_Bracket_Gate_2"), GateSoundLevel * 0.005, ActiveBall
End Sub

Sub SoundOneWayGate() 'sound of blocked gate
  PlaySoundAtLevelStatic (Cartridge_Gates & "_Oneway_Ball_Gate_" & Int(Rnd*3)+1), GateSoundLevel, ActiveBall
End Sub

Sub Gates_hit(idx)
  SoundHeavyGate
End Sub

Sub GatesWire_hit(idx)
  SoundPlayfieldGate
End Sub

Sub GatesOneWay_Hit(idx)
  SoundOneWayGate
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
      snd = (Cartridge_BallBallCollision & "_BallBall_Collide_1")
    Case 2
      snd = (Cartridge_BallBallCollision & "_BallBall_Collide_2")
    Case 3
      snd = (Cartridge_BallBallCollision & "_BallBall_Collide_3")
    Case 4
      snd = (Cartridge_BallBallCollision & "_BallBall_Collide_4")
    Case 5
      snd = (Cartridge_BallBallCollision & "_BallBall_Collide_5")
    Case 6
      snd = (Cartridge_BallBallCollision & "_BallBall_Collide_6")
    Case 7
      snd = (Cartridge_BallBallCollision & "_BallBall_Collide_7")
  End Select

  PlaySound (snd), 0, CSng(velocity) ^ 2 / 200 * BallWithBallCollisionSoundFactor * VolumeDial, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub


'///////////////////////////  DROP TARGET HIT SOUNDS  ///////////////////////////

Sub RandomSoundDropTargetReset(obj)
  PlaySoundAtLevelStatic SoundFX(Cartridge_Targets & "_Drop_Target_1Bank_Reset_Up_" & Int(Rnd*8)+1,DOFContactors), 1, obj
End Sub

Sub RandomSoundDropTargetBankReset(obj)
  PlaySoundAtLevelStatic SoundFX(Cartridge_Targets & "_Drop_Target_3Bank_Reset_Up_" & Int(Rnd*6)+1,DOFContactors), 1, obj
End Sub

Sub SoundDropTargetDrop(obj)
  PlaySoundAtLevelStatic (Cartridge_Targets & "_Drop_Target_1Bank_Release_Down_" & Int(Rnd*2)+1), 200, obj
End Sub

Sub SoundDropTargetBankDrop(obj)
  PlaySoundAtLevelStatic (Cartridge_Targets & "_Drop_Target_3Bank_Release_Down_" & Int(Rnd*6)+1), 200, obj
End Sub

'/////////////////////////////  GI AND FLASHER RELAYS  ////////////////////////////

Const RelayFlashSoundLevel = 0.015  'volume level; range [0, 1];
Const RelayGISoundLevel = 0.45        'volume level; range [0, 1];
Const RelayBackboxSoundLevel = 0.45    'volume level; range [0, 1];
Const RelayACSelectSoundLevel = 0.3 'volume level; range [0, 1];

'Sub Sound_GI_Relay(toggle, obj)
'    Select Case toggle
'        Case 1
'            PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_Upper_Playfield_GI_Relay_On"), RelayGISoundLevel, obj
'        Case 0
'            PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_Upper_Playfield_GI_Relay_Off"), RelayGISoundLevel, obj
'    End Select
'End Sub

'Sub Sound_GIBackbox_Relay(toggle, obj)
'    Select Case toggle
'        Case 1
'            PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_Lower_Playfield_and_Backbox_GI_Relay_On"), RelayBackboxSoundLevel, obj
'        Case 0
'            PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_Lower_Playfield_and_Backbox_GI_Relay_Off"), RelayBackboxSoundLevel, obj
'    End Select
'End Sub

'Sub Sound_Flash_Relay(toggle, obj)
'    Select Case toggle
'        Case 1
'            PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_Flashers_Relay_On"), RelayFlashSoundLevel, obj
'        Case 0
'            PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_Flashers_Relay_Off"), RelayFlashSoundLevel, obj
'    End Select
'End Sub

'//////////////////////////  SOLENOID A/C SELECT RELAY  /////////////////////////
Sub Sound_Solenoid_ACSelect_Relay(toggle, obj)
    Select Case toggle
        Case 1
            PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_AC_Select_Relay_Side_A"), RelayACSelectSoundLevel, obj
        Case 0
            PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_AC_Select_Relay_Side_C"), RelayACSelectSoundLevel, obj
    End Select
End Sub


'/////////////////////////////////////////////////////////////////
'         End Mechanical Sounds
'/////////////////////////////////////////////////////////////////

'******************************************************
'****  END FLEEP MECHANICAL SOUNDS
'******************************************************

'******************************************************
' ZPHY:  GENERAL ADVICE ON PHYSICS
'******************************************************
'
' It's advised that flipper corrections, dampeners, and general physics settings should all be updated per these
' examples as all of these improvements work together to provide a realistic physics simulation.
'
' Tutorial videos provided by Bord
' Adding nFozzy roth physics : pt1 rubber dampeners         https://youtu.be/AXX3aen06FM?si=Xqd-rcaqTlgEd_wx
' Adding nFozzy roth physics : pt2 flipper physics          https://youtu.be/VSBFuK2RCPE?si=i8ne8Ao2co8rt7fy
' Adding nFozzy roth physics : pt3 other elements           https://youtu.be/JN8HEJapCvs?si=hvgMOk-ej1BEYjJv
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
' | Force         | 12-15    |
' | Hit Threshold | 1.6-2    |
' | Scatter Angle | 2        |
'
' Slingshots
' | Hit Threshold      | 2    |
' | Slingshot Force    | 3-5  |
' | Slingshot Theshold | 2-3  |
' | Elasticity         | 0.85 |
' | Friction           | 0.8  |
' | Scatter Angle      | 1    |






'******************************************************
' ZNFF:  FLIPPER CORRECTIONS by nFozzy
'******************************************************
'
' There are several steps for taking advantage of nFozzy's flipper solution.  At a high level we'll need the following:
' 1. flippers with specific physics settings
' 2. custom triggers for each flipper (TriggerLF, TriggerRF)
' 3. and, special scripting
'
' TriggerLF and RF should now be 27 vp units from the flippers. In addition, 3 degrees should be added to the end angle
' when creating these triggers.
'
' RF.ReProcessBalls Activeball and LF.ReProcessBalls Activeball must be added the flipper_collide subs.
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
' | EOS Torque         | 0.4            | 0.4                   | 0.375                  | 0.375              |
' | EOS Torque Angle   | 4              | 4                     | 6                      | 6                  |
'

'******************************************************
' Flippers Polarity (Select appropriate sub based on era)
'******************************************************

Dim LF : Set LF = New FlipperPolarity
Dim RF : Set RF = New FlipperPolarity
Dim ULF : Set ULF = New FlipperPolarity

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
RightFlipper.timerenabled = True

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
' The Stand Up and Drop Target solutions improve the physics for targets to create more realistic behavior. It allows the ball
' to move through the target enabling the ability to score more than one target with a well placed shot.
' It also handles full target animation, switch handling and deflection on hit. For drop targets there is also a slight lift when
' the drop targets raise, bricking, and popping the ball up if it's over the drop target when it raises.
'
' Add a Timers named DTAnim and STAnim to editor to handle drop & standup target animations, or run them off an always-on 10ms timer (GameTimer)
' DTAnim.interval = 10
' DTAnim.enabled = True

' Sub DTAnim_Timer
'   DoDTAnim
' DoSTAnim
' End Sub

' For each drop target, we'll use two wall objects for physics calculations and one primitive for visuals and
' animation. We will not use target objects.  Place your drop target primitive the same as you would a VP drop target.
' The primitive should have it's pivot point centered on the x and y axis and at or just below the playfield
' level on the z axis. Orientation needs to be set using Rotz and bending deflection using Rotx. You'll find a hooded
' target mesh in this table's example. It uses the same texture map as the VP drop targets.
'
' For each stand up target we'll use a vp target, a laid back collidable primitive, and one primitive for visuals and animation.
' The visual primitive should should have it's pivot point centered on the x and y axis and the z should be at or just below the playfield.
' The target should animate backwards using transy.
'
' To create visual target primitives that work with the stand up and drop target code, follow the below instructions:
' (Other methods will work as well, but this is easy for even non-blender users to do)
' 1) Open a new blank table. Delete everything off the table in editor.
' 2) Copy and paste the VP target from your table into this blank table.
' 3) Place the target at x = 0, y = 0  (upper left hand corner) with an orientation of 0 (target facing the front of the table)
' 4) Under the file menu, select Export "OBJ Mesh"
' 5) Go to "https://threejs.org/editor/". Here you can modify the exported obj file. When you export, it exports your target and also
'    the playfield mesh. You need to delete the playfield mesh here. Under the file menu, chose import, and select the obj you exported
'    from VPX. In the right hand panel, find the Playfield object and click on it and delete. Then use the file menu to Export OBJ.
' 6) In VPX, you can add a primitive and use "Import Mesh" to import the exported obj from the previous step. X,Y,Z scale should be 1.
'    The primitive will use the same target texture as the VP target object.
'
' * Note, each target must have a unique switch number. If they share a same number, add 100 to additional target with that number.
' For example, three targets with switch 32 would use 32, 132, 232 for their switch numbers.
' The 100 and 200 will be removed when setting the switch value for the target.

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
Dim DT33, DT57, DT58, DT59

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

Set DT33 = (new DropTarget)(sw33, sw33a, BM_sw33, 33, 0, false)
Set DT57 = (new DropTarget)(sw57, sw57a, BM_sw57, 57, 0, false)
Set DT58 = (new DropTarget)(sw58, sw58a, BM_sw58, 58, 0, false)
Set DT59 = (new DropTarget)(sw59, sw59a, BM_sw59, 59, 0, false)

Dim DTArray
DTArray = Array(DT33, DT57, DT58, DT59)

'Configure the behavior of Drop Targets.
Const DTDropSpeed = 80 'in milliseconds
Const DTDropUpSpeed = 40 'in milliseconds
Const DTDropUnits = 47 'VP units primitive drops so top of at or below the playfield
Const DTDropUpUnits = 10 'VP units primitive raises above the up position on drops up
Const DTMaxBend = 8 'max degrees primitive rotates when hit
Const DTDropDelay = 20 'time in milliseconds before target drops (due to friction/impact of the ball)
Const DTRaiseDelay = 40 'time in milliseconds before target drops back to normal up position after the solenoid fires to raise the target
Const DTBrickVel = 30 'velocity at which the target will brick, set to '0' to disable brick

Const DTEnableBrick = 0 'Set to 0 to disable bricking, 1 to enable bricking
Const DTHitSound = "" 'Drop Target Hit sound
Const DTDropSound = "DropTarget_Down" 'Drop Target Drop sound
Const DTResetSound = "DropTarget_Up" 'Drop Target reset sound

Const DTMass = 0.1 'Mass of the Drop Target (between 0 and 1), higher values provide more resistance


'******************************************************
'  DROP TARGETS FUNCTIONS
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

Sub UpdateDropTargets
  dim BP, tz, rx, ry

    tz = BM_sw33.transz
  rx = BM_sw33.rotx
  ry = BM_sw33.roty
  For each BP in BP_sw33 : BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

    tz = BM_sw57.transz
  rx = BM_sw57.rotx
  ry = BM_sw57.roty
  For each BP in BP_sw57 : BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

    tz = BM_sw58.transz
  rx = BM_sw58.rotx
  ry = BM_sw58.roty
  For each BP in BP_sw58 : BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

    tz = BM_sw59.transz
  rx = BM_sw59.rotx
  ry = BM_sw59.roty
  For each BP in BP_sw59 : BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next
End Sub

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
Dim ST18, ST19, ST20, ST21, ST22, ST23, ST24, ST25, ST26, ST27, ST28, ST29, ST30, ST31, ST35

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

Set ST18 = (new StandupTarget)(sw18, BM_sw18, 18, 0)
Set ST19 = (new StandupTarget)(sw19, BM_sw19, 19, 0)
Set ST20 = (new StandupTarget)(sw20, BM_sw20, 20, 0)
Set ST21 = (new StandupTarget)(sw21, BM_sw21, 21, 0)
Set ST22 = (new StandupTarget)(sw22, BM_sw22, 22, 0)
Set ST23 = (new StandupTarget)(sw23, BM_sw23, 23, 0)
Set ST24 = (new StandupTarget)(sw24, BM_sw24, 24, 0)
Set ST25 = (new StandupTarget)(sw25, BM_sw25, 25, 0)
Set ST26 = (new StandupTarget)(sw26, BM_sw26, 26, 0)
Set ST27 = (new StandupTarget)(sw27, BM_sw27, 27, 0)
Set ST28 = (new StandupTarget)(sw28, BM_sw28, 28, 0)
Set ST29 = (new StandupTarget)(sw29, BM_sw29, 29, 0)
Set ST30 = (new StandupTarget)(sw30, BM_sw30, 30, 0)
Set ST31 = (new StandupTarget)(sw31, BM_sw31, 31, 0)
Set ST35 = (new StandupTarget)(sw35, BM_sw35, 35, 0)

'Add all the Stand-up Target Arrays to Stand-up Target Animation Array
' STAnimationArray = Array(ST1, ST2, ....)
Dim STArray
STArray = Array(ST18, ST19, ST20, ST21, ST22, ST23, ST24, ST25, ST26, ST27, ST28, ST29, ST30, ST31, ST35)

'Configure the behavior of Stand-up Targets
Const STAnimStep =  1.5     'vpunits per animation step (control return to Start)
Const STMaxOffset = 9       'max vp units target moves when hit

Const STMass = 0.1        'Mass of the Stand-up Target (between 0 and 1), higher values provide more resistance

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

Sub UpdateStandupTargets
  dim BP, ty

    ty = BM_sw18.transy
  For each BP in BP_sw18 : BP.transy = ty: Next

    ty = BM_sw19.transy
  For each BP in BP_sw19 : BP.transy = ty: Next

    ty = BM_sw20.transy
  For each BP in BP_sw20 : BP.transy = ty: Next

    ty = BM_sw21.transy
  For each BP in BP_sw21 : BP.transy = ty: Next

    ty = BM_sw22.transy
  For each BP in BP_sw22 : BP.transy = ty: Next

    ty = BM_sw23.transy
  For each BP in BP_sw23 : BP.transy = ty: Next

    ty = BM_sw24.transy
  For each BP in BP_sw24 : BP.transy = ty: Next

    ty = BM_sw25.transy
  For each BP in BP_sw25 : BP.transy = ty: Next

    ty = BM_sw26.transy
  For each BP in BP_sw26 : BP.transy = ty: Next

    ty = BM_sw27.transy
  For each BP in BP_sw27 : BP.transy = ty: Next

    ty = BM_sw28.transy
  For each BP in BP_sw28 : BP.transy = ty: Next

    ty = BM_sw29.transy
  For each BP in BP_sw29 : BP.transy = ty: Next

    ty = BM_sw30.transy
  For each BP in BP_sw30 : BP.transy = ty: Next

    ty = BM_sw31.transy
  For each BP in BP_sw31 : BP.transy = ty: Next

    ty = BM_sw35.transy
  For each BP in BP_sw35 : BP.transy = ty: Next

End Sub

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
'****  ZGIU:  GI Control
'******************************************************

'**** Global variable to hold current GI light intensity. Not Mandatory, but may help in other lighting subs and timers
'**** This value is updated always when GI state is being updated
dim gilvl

'**** SetRelayGI is called from SolCallback.
sub GIUpdates(aLvl)
    If DebugGI Then debug.print "    GIUpdates value: " & aLvl
    Dim bulb, LM
    For Each bulb in gi_green: bulb.State = aLvl: Next
    For Each bulb in GI: bulb.State = aLvl: Next
    For Each LM in BL_Bumpers: LM.opacity = aLvl*100: Next
    gilvl = aLvl
    Dim GIRelay : GIRelay = 0
    If aLvl >= 0.99 and GIRelay = 0 Then
        Sound_Solenoid_ACSelect_Relay 1, KnockerPosition
        GIRelay = 1
    ElseIf aLvl <= 0.01 and GIRelay = 1 Then
        Sound_Solenoid_ACSelect_Relay 1, KnockerPosition
        GIRelay = 0
    End If
End Sub

sub GISelect(enabled)
  If DebugGI Then debug.print "GISelect value: " & enabled
  Dim BL
  If enabled Then
    IsMultiball = False
    For Each BL in BL_GIG : BL.visible = 0 : Next
    For Each BL in BL_GIGS_GI_G_001 : BL.visible = 0 : Next
    For Each BL in BL_GIGS_GI_G_002 : BL.visible = 0 : Next
    For Each BL in BL_GIGS_GI_G_020 : BL.visible = 0 : Next
    For Each BL in BL_GIGS_GI_G_022 : BL.visible = 0 : Next
    For Each BL in BL_GIGS_GI_G_023 : BL.visible = 0 : Next
    For Each BL in BL_GIGS_GI_G_024 : BL.visible = 0 : Next
    For Each BL in BL_GIN : BL.visible = 1 : Next
    For Each BL in BL_GINS_gi01 : BL.visible = 1 : Next
    For Each BL in BL_GINS_gi02 : BL.visible = 1 : Next
    For Each BL in BL_GINS_gi24 : BL.visible = 1 : Next
    For Each BL in BL_GINS_gi25 : BL.visible = 1 : Next
    For Each BL in BL_GINS_gi26 : BL.visible = 1 : Next
  Else
    IsMultiball = True
    If RandomGIColor Then
      SetGiColor 8
    End If
    For Each BL in BL_GIG : BL.visible = 1 : Next
    For Each BL in BL_GIGS_GI_G_001 : BL.visible = 1 : Next
    For Each BL in BL_GIGS_GI_G_002 : BL.visible = 1 : Next
    For Each BL in BL_GIGS_GI_G_020 : BL.visible = 1 : Next
    For Each BL in BL_GIGS_GI_G_022 : BL.visible = 1 : Next
    For Each BL in BL_GIGS_GI_G_023 : BL.visible = 1 : Next
    For Each BL in BL_GIGS_GI_G_024 : BL.visible = 1 : Next
    For Each BL in BL_GIN : BL.visible = 0 : Next
    For Each BL in BL_GINS_gi01 : BL.visible = 0 : Next
    For Each BL in BL_GINS_gi02 : BL.visible = 0 : Next
    For Each BL in BL_GINS_gi24 : BL.visible = 0 : Next
    For Each BL in BL_GINS_gi25 : BL.visible = 0 : Next
    For Each BL in BL_GINS_gi26 : BL.visible = 0 : Next
  End If
  For Each BL in BP_LSling1 : BL.Visible = 0: Next
    For Each BL in BP_LSling2 : BL.Visible = 0: Next
  For Each BL in BP_RSling1 : BL.Visible = 0: Next
    For Each BL in BP_RSling2 : BL.Visible = 0: Next
end sub


Sub GiBackbox(enabled)
    If DebugGI Then debug.print "GiBackbox value: " & enabled
    If Enabled Then
        GI_BG.state = 1
        Sound_Solenoid_ACSelect_Relay 1, KnockerPosition
    Else
        GI_BG.state = 0
        Sound_Solenoid_ACSelect_Relay 1, KnockerPosition
    End If
End Sub

Sub SolACSelectRelay(enabled)
    If DebugGI Then debug.print "SolACSelectRelay value: " & enabled
    If Enabled Then
        Sound_Solenoid_ACSelect_Relay 1, KnockerPosition
    Else
        Sound_Solenoid_ACSelect_Relay 0, KnockerPosition
    End If
End Sub

'******************************************************
'****  END GI Control
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
Dim objBallShadow(3)

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


'***************************************************************
' ZLED: Display LEDs
'***************************************************************


'Desktop
Dim Digits(28)
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
Digits(14) = Array(b00, b02, b05, b06, b04, b01, b03, b07)
Digits(15) = Array(b10, b12, b15, b16, b14, b11, b13)
Digits(16) = Array(b20, b22, b25, b26, b24, b21, b23)
Digits(17) = Array(b30, b32, b35, b36, b34, b31, b33, b37)
Digits(18) = Array(b40, b42, b45, b46, b44, b41, b43)
Digits(19) = Array(b50, b52, b55, b56, b54, b51, b53)
Digits(20) = Array(b60, b62, b65, b66, b64, b61, b63)
Digits(21) = Array(b70, b72, b75, b76, b74, b71, b73, b77)
Digits(22) = Array(b80, b82, b85, b86, b84, b81, b83)
Digits(23) = Array(b90, b92, b95, b96, b94, b91, b93)
Digits(24) = Array(ba0, ba2, ba5, ba6, ba4, ba1, ba3, ba7)
Digits(25) = Array(bb0, bb2, bb5, bb6, bb4, bb1, bb3)
Digits(26) = Array(bc0, bc2, bc5, bc6, bc4, bc1, bc3)
Digits(27) = Array(bd0, bd2, bd5, bd6, bd4, bd1, bd3)


InitLeds
Sub InitLeds
  Dim ii, jj, dig
  For ii = 0 to 27
    dig = Digits(ii)
    For jj = 0 to Ubound(dig)
      dig(jj).ColorFull = RGB(155,62,0)
    Next
  Next
End Sub


Sub UpdateLeds
    Dim ChgLED, ii, jj, num, chg, stat, obj, b, x
    ChgLED = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
        For ii = 0 To UBound(chgLED)
            num = chgLED(ii, 0):chg = chgLED(ii, 1):stat = chgLED(ii, 2)
      For Each obj In Digits(num)
        If chg And 1 Then obj.State = stat And 1
        chg = chg \ 2:stat = stat \ 2
      Next
        Next
    End If
End Sub


Sub SetupRoom
  If RenderingMode = 2 or Table1.ShowFSS = True or VRTest = 1 Then
    VRMode = True
    For Each dtlight in DesktopLights: dtlight.visible = 0: Next
    For Each VR_Obj in VRLedDigits:VR_Obj.Visible = 1:Next
    For Each VR_Obj in VRCabinet : VR_Obj.Visible = 1 : Next
    For Each VR_Obj in BackglassFlashers : VR_Obj.Visible = 1 : Next
    For Each VR_Obj in BackglassFlashers_Top : VR_Obj.Visible = 1 : Next
    For Each VR_Obj in VRMinimalRoom : VR_Obj.Visible = 1 : Next
    Primary_LockDownBar.Visible = 1
    Primary_SideRailMapRight.Visible = 1
    Primary_SideRailMapLeft.Visible = 1
    If VRRoomChoice = 1 Then
      For Each VR_Obj in VRMinimalRoom : VR_Obj.Visible = 0 : Next
      For Each VR_Obj in VRMegaRoom : VR_Obj.Visible = 1 : Next
      Plaque.Sidevisible = 1
      Rock1Timer.enabled = True
      if AlienOn = true Then EyeballTimer.enabled = True
      SkyTimer.enabled = True
      StarTimer.enabled = True
      LightTest.enabled = True
      VRFire.visible = False  ' ????  work?
      ShipControlTimer.enabled = True  ' do regardless
      FSTimer.enabled = True
      If ShipOn = True Then
        ShipTimer.enabled = True
        FireTimer.enabled = True
        ShipControlTimer.enabled = false
        VRFire.visible = True
        WobbleTimer.enabled = True
      End If
    Else
      For Each VR_Obj in VRMinimalRoom : VR_Obj.Visible = 1 : Next
      For Each VR_Obj in VRMegaRoom : VR_Obj.Visible = 0 : Next
      Plaque.Sidevisible = 0
      Rock1Timer.enabled = False
      SkyTimer.enabled = False
      StarTimer.enabled = False
      LightTest.enabled = False
      ShipTimer.enabled = False
      FireTimer.enabled = False
      ShipControlTimer.enabled = False
      WobbleTimer.enabled = False
      VRFire.visible = False
      FSTimer.enabled = False
    End If
  Else
    VRMode = False
    For Each VR_Obj in VRLedDigits:VR_Obj.Visible = 0:Next
    For Each VR_Obj in VRCabinet : VR_Obj.Visible = 0 : Next
    For Each VR_Obj in BackglassFlashers : VR_Obj.Visible = 0 : Next
    For Each VR_Obj in BackglassFlashers_Top : VR_Obj.Visible = 0 : Next
    For Each VR_Obj in VRMinimalRoom : VR_Obj.Visible = 0 : Next
    For Each VR_Obj in VRMegaRoom : VR_Obj.Visible = 0 : Next
    If DesktopMode Then
      'Primary_LockDownBar.Visible = 1
      'Primary_SideRailMapRight.Visible = 1
      'Primary_SideRailMapLeft.Visible = 1
      Primary_SideRailMapRight.z = 198-25
      Primary_SideRailMapLeft.z = 333-30
      Primary_LockDownBar.z = -41-25
    End If
    VRBackwall.Visible = False
    VRBackwall.SideVisible = False
    Plaque.Sidevisible = 0
    Rock1Timer.enabled = False
    SkyTimer.enabled = False
    StarTimer.enabled = False
    LightTest.enabled = False
    ShipTimer.enabled = False
    FireTimer.enabled = False
    ShipControlTimer.enabled = False
    WobbleTimer.enabled = False
    VRFire.visible = False
    FSTimer.enabled = False
  End If
End Sub



'***************************************************************************
' VR BG Display
'***************************************************************************
Dim DisplayColor

Dim VRDigits(32)
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
VRDigits(14)=Array(D211,D212,D213,D214,D215,D216,D217,D218)
VRDigits(15)=Array(D219,D220,D221,D222,D223,D224,D225,D226)
VRDigits(16)=Array(D227,D228,D229,D230,D231,D232,D233,D234)
VRDigits(17)=Array(D235,D236,D237,D238,D239,D240,D241,D242)
VRDigits(18)=Array(D243,D244,D245,D246,D247,D248,D249,D250)
VRDigits(19)=Array(D251,D252,D253,D254,D255,D256,D257,D258)
VRDigits(20)=Array(D259,D260,D261,D262,D263,D264,D265,D266)
VRDigits(21)=Array(D267,D268,D269,D270,D271,D272,D273,D274)
VRDigits(22)=Array(D275,D276,D277,D278,D279,D280,D281,D282)
VRDigits(23)=Array(D283,D284,D285,D286,D287,D288,D289,D290)
VRDigits(24)=Array(D291,D292,D293,D294,D295,D296,D297,D298)
VRDigits(25)=Array(D299,D300,D301,D302,D303,D304,D305,D306)
VRDigits(26)=Array(D307,D308,D309,D310,D311,D312,D313,D314)
VRDigits(27)=Array(D315,D316,D317,D318,D319,D320,D321,D322)

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
    object.color = RGB(255,51,0)
    Object.Opacity = 100
  Else
    Object.Color = RGB(1,1,1)
    Object.Opacity = 1
  End If
End Sub

Sub InitDigits()
  dim tmp, x, obj
  for x = 0 to uBound(VRDigits)
    if IsArray(VRDigits(x) ) then
      For each obj in VRDigits(x)
        obj.height = obj.height + 0
        FadeDisplay obj, 0
      next
    end If
  Next
  DisplayColor = RGB(255,51,0)
End Sub

'VR Stuff Below.. ****************************************************************************************************

'******************************************************
'*******  Set Up Backglass Flashers *******
'******************************************************
' this is for lining up the backglass flashers on top of a backglass image

Sub SetBackglass()
  Dim obj
  For Each obj In BackglassFlashers
    obj.x = obj.x
    obj.height = - obj.y + 650
    obj.y = -72 'adjusts the distance from the backglass towards the user
  Next
  For Each obj In BackglassFlashers_Top
    obj.x = obj.x
    obj.height = - obj.y + 650
    obj.y = -76 'adjusts the distance from the backglass towards the user
  Next
End Sub

Dim move1:move1=0.02   'For VR Skybox

Sub SkyTimer_Timer()
  Sky.ObjRotZ=Sky.ObjRotZ+move1
End Sub

Dim move2:move2=0.06   'For VR Meteors
Dim move3:move3=0.08
Dim MoveShip: MoveShip = 0.2
Dim MoveShip2: MoveShip2 = 0.25

Sub Rock1Timer_Timer()
  Rock1.ObjRotZ=Rock1.ObjRotZ+move2
  Rock1.ObjRotX=Rock1.ObjRotX+move3
  Rock2.ObjRotZ=Rock2.ObjRotZ+move2
  Rock2.ObjRotX=Rock2.ObjRotX+move3
  Rock4.ObjRotZ=Rock4.ObjRotZ+move3
  Rock4.ObjRotX=Rock4.ObjRotX+move2
  Rock7.ObjRotZ=Rock7.ObjRotZ+move3
  Rock7.ObjRotX=Rock7.ObjRotX+move2
  FallingStar1.ObjRotZ=FallingStar1.ObjRotZ+move3
  FallingStar1.ObjRotX=FallingStar1.ObjRotX+move3
  FallingStar2.ObjRotZ=FallingStar2.ObjRotZ+move3
  FallingStar2.ObjRotX=FallingStar2.ObjRotX+move3
  SpaceStation.ObjRotZ=SpaceStation.ObjRotZ+move2
  FlyingShip.ObjRotZ=FlyingShip.ObjRotZ - 0.08
  FlyingShip.z = FlyingShip.z + MoveShip
  FlyingShip.TransX = FlyingShip.TransX + MoveShip2
  if FlyingShip.TransX =< -2500 then Moveship2 = 0.25
  if FlyingShip.TransX => -1050 then Moveship2 = -0.25
  if FlyingShip.z => 2200 then Moveship = -0.2
  if FlyingShip.z =< 300 then Moveship = 0.2
  Moon.ObjRotZ=Moon.ObjRotZ+move2
End Sub

Dim move4:move4=22
Dim move5:move5=-24

Sub StarTimer_Timer()
  FallingStar1.Y=FallingStar1.Y+move4
  FallingStar2.Y=FallingStar2.Y+move5
  If FallingStar1.Y>=50000 then Randomize (21): FallingStar1.Y = -42000: FallingStar1.X = 15000 + rnd(1)*-25000 : FallingStar1.Size_x = 1300 * rnd(1) +320 : FallingStar1.Size_y = 1300 * rnd(1) +320: FallingStar1.Size_z = 1300 * rnd(1) +550' Randomize X position, x,y,z size here
  If FallingStar2.Y<=-40000 then Randomize (5): FallingStar2.Y = 50000: FallingStar2.X = 15000 + rnd(1)*-25000 : FallingStar2.Size_x = 1300 * rnd(1) +320 : FallingStar2.Size_y = 1300 * rnd(1) +320: FallingStar2.Size_z = 1300 * rnd(1) +550' Randomize X position, x,y,z size here
End Sub

'Moving Alien for VR...
dim AlienCount: AlienCount = 0

Dim AlienOn
AlienOn = false

Sub MoveAlien()
  Alien.Z = 790: Alienshadow.z = -790: Alienshadow1.z = -790: Alienshadow2.z = -790: Alien2.Z = -6500: Lefteye.Z = 1238: Righteye.z=1186
  EyeballTimer.enabled = true
  AlienOn = true
End Sub

Dim move7: move7 = 2
Sub EyeballTimer_Timer
  LeftEye.ObjRotZ=Lefteye.ObjRotZ+move7
  RightEye.ObjRotZ=Righteye.ObjRotZ+move7
  if Lefteye.ObjRotz >= 30 then move7 = -2
  if Lefteye.ObjRotz <= -10 then move7 = 2
End Sub

' plunger code below.
Sub TimerVRPlunger_Timer()
  If VRPlunger.Y < 2264 then
       VRPlunger.Y = VRPlunger.Y + 5
  End If
End Sub

Sub TimerVRPlunger2_Timer()
  VRPlunger.Y = 2149 + (5* Plunger.Position) -20
End Sub

Sub LightTest_Timer()
  if RedUFO.visible = false then
    RedUFO.visible = true
    UFOMiddle.disablelighting = 1
  Else
    RedUFO.visible = false
    UFOMiddle.disablelighting = 0
  end if
End Sub

Dim ShipSpeed
ShipSpeed = 4

Sub ShipTimer_Timer()
  Moonbase.z = Moonbase.z + ShipSpeed
  VRFire.z = VRFire.z + ShipSpeed
  if Moonbase.z  = 14496 Then
    ShipSpeed = -4
  end if
  if Moonbase.z  = -788 Then
    Moonbase.z = -784
    ShipSpeed = 4
    VRFire.visible = false
    Firetimer.enabled = false
    ShipTimer.enabled = false
    ShipOn = false
    ShipControlTimer.enabled = true
  end if
  if Moonbase.z  > 0 Then
    WobbleTimer.enabled = True
  Else
    WobbleTimer.enabled = False
    Moonbase.Roty = -34 ' bring ship back flat
    Moonbase.Rotz = 0
    Moonbase.Rotx = 92
  end if
End Sub

Dim VRFireCounter
VRFireCounter = 1
VRFire. visible = false  '  I do this because otherwise the first image on animation shows whits.

Sub FireTimer_Timer()
  VRFire.Image = "Fire_" & VRFireCounter
  VRFireCounter = VRFireCounter + 1
  If VRFireCounter > 59 Then
    VRFireCounter = 0
  end If
End Sub

Dim ShipOn
ShipOn = false

Sub ShipControlTimer_Timer()
  Firetimer.enabled = True
  ShipTimer.enabled = True
  ShipControlTimer.enabled = false
  VRFire.visible = true
  ShipOn = true
End Sub

Dim WobbleSpeed
Dim WobbleSpeed2
Dim WobbleSpeed3
WobbleSpeed = 0.04
WobbleSpeed2 = 0.02

Sub WobbleTimer_Timer()
' Roty -34 is landing..
  Moonbase.Roty = Moonbase.Roty + WobbleSpeed
  Moonbase.Rotz = Moonbase.Rotz + WobbleSpeed2
  Moonbase.Rotx = Moonbase.Rotx + WobbleSpeed
  if Moonbase.Roty > -30 then
    WobbleSpeed = -0.04
    WobbleSpeed2 = -0.02
  end If
  if Moonbase.Roty < -38 then
    WobbleSpeed = 0.04
    WobbleSpeed2 = 0.02
  end If
End Sub

'FlyingShip.image = "SF_Fighter3"

Dim ShipColour
ShipColour = 1

Sub FSTimer_timer()
  If ShipColour = 1 then FlyingShip.image = "SF_Fighter2":ShipColour = 2: exit sub
  If ShipColour = 2 then FlyingShip.image = "SF_Fighter3":ShipColour = 3: exit sub
  If ShipColour = 3 then FlyingShip.image = "SF_Fighter1":ShipColour = 1: exit sub
End Sub

Sub BGVRBulbTimer_Timer()
  BGBulb1.disablelighting = Controller.Lamp(47)  ' This matches up with Joels Flasher1
  BGBulb2.disablelighting = Controller.Lamp(46)  ' This matches up with Joels Flasher2
  BGBulb3.disablelighting = Controller.Lamp(37)  ' This matches up with Joels Flasher3
  BGBulb4.disablelighting = Controller.Lamp(57)  ' This matches up with Joels Flasher4
  BGBulb5.disablelighting = Controller.Lamp(41)  ' This matches up with Joels Flasher5
  BGBulb6.disablelighting = Controller.Lamp(48)  ' This matches up with Joels Flasher6
End Sub


'Changelog
'=========
'001 - tomate - blank VPX table with correct dimensions created
'002 - tomate - first 512 batch to place the objects on VPX
'003 - Sixtoe - Built... Everything? Crashes. Sucks to be me.
'004 - apophis - Ground control to Major Sixtoe. Fixed crashes, standup targets, some collections. Trough, kickback, GI still not working correctly.
'005 - Sixtoe - I think my spaceship knows which way to go. Hooked up space station & placeholder prim, Fixed trough, added settings to flippers, slings and pop bumpers, adjusted flipper triggers, kickback fixed, turned everything invisible, added more blocker walls, loads of small fixes and tweaks. Still no GI. Fully playable now?
'006 - tomate - New 2k batch added, GI doesnt work yet - No flashers added
'007 - mcarter78 - Hook up modulated GI with coil driven color toggle
'008 - tomate - New 2k batch added, inserts fixed, added Flupper LUT and desaturated options (see 2°, 3° and 4° option into game menu), changed material for upper PF color and added option menu for MB GI color
'009 - mcarter78 - Add more MB GI color options and random option.  Fix for GI level after multiball end.
'010 - tomate - Extra LUT options added, l30 fixed, tweaked blue MB color
'011 - tomate -  new 2k batch added with a lot of blender fixes, movable objects collection set (not hooked up yet), flashers and bumper lights added
'012 - apophis - simplified the GI code
'013 - tomate - new 2k batch added, new flashers, more LUT options added
'014 - tomate - new 2k batch added, tweaked flashers materisl, improved upperPF material, fixed upper PF inserts and bulbs
'015 - Sixtoe - Hooked up movables (flippers, slings, targets, spacestation and diverter) but there are visual issues (space station lighting and drop targets?), made upper pf hole fractionally bigger, added plunger texture and tweaked power
'016 - apophis - Fixed some fleep stuff. Fixed some animation stuff. Set up ramp rolling sfx triggers. Made bumper lights follow GI. Adjusted flipper strength. Removed old space station toy prim.
'017 - tomate - unwrapped PF, added diverter and space station stages, tried to fix upper PF front wall but failed
'018 - apophis - Fix motor sfx issue. Trying to get the SS and Diverter animations to work. Still broken.
'019 - tomate - New batch added, several fixes at blender side (finally upperPF looks correct), new diverter/SS stages added
'020 - apophis - SS and Diverter animations working. Added desktop alphanumeric display.
'021 - Reverted
'022 - Sixtoe - Tweaked and randomised kickers, including kickback, hard coded sling timer and reduced to 10 (still bugs out occasionally?), split shooter wall and changed to wood, retooled orbit exit and gate lower rubber, made physical drop target smaller, made cone rubbers smaller, swapped to correct Const EOSTnew (thanks RobbyKingPin)
'023 - tomate - tweaked upperPF and bumpers caps materials. Added missing materials to room brightness array
'024 - tomate - new 2k batch added, new flahser domes, unwrapped textures for flippers and targets, fixed PF texture, changed value for DTDropUnits (thanks Apophis!)
'025 - apophis - Made desktop LED display only visible in desktop mode.
'026 - tomate - flasher domes fixed
'027 - tomate - new 2k batch, inserts color corrected, performance improvements
'028 - Sixtoe - new 4k batch
'029 - tomate - added missing renders, separated GINS and GIGS for raytraced shadows, modified upper PF hole to get some ball flirting, fixed slingshots default position (thanks bthlonewolf!)
'030 - tomate - added "Vibrant" LUT for the ones that prefer to play with AGX instead Reinhard filter
'031 - tomate - diverted fixed
'032 - tomate - PF texture fixed
'033 - apophis - Sling lightmap visibility bug fixed
'034 - DGrimmReaper - Ported over VR from Rawd's Table (still need to hook up Backglass lights)
'035 - RobbyKingPin - Created a custom Fleep Sound package made from assets and codes used from Bride of Pinbot. Customized to be more identical to the original Fleep codes to make future conversions a bit more easy.
'036 - DGrimmReaper - Hooked Up VR Backglass Lights
'037 - tomate - new natural green MB option added, lighting tweaks, ball reflection decreased, bloom strength increased
'038 - tomate - re-rendered GI lighting, now swappable by script (still need to add this as an option in f12 menu)
'039 - DGrimmReaper - cleaned up images and converted larger images to webp
'040 - RobbyKingPin - Fixed an issue with the Rothbauer droptarget codes. Hooked up relay sounds for flashers and GI. Removed a duplicate subroutine for the GI color.
'041 - tomate - right flipper texture fixed
'042 - bthlonewolf - ball tint in multiball, added color temp settings to tie into existing code
'043 - apophis - Cab rails visible in desktop mode. Small tweak to bumper positions. Hooked up GI Backbox solenoid to VR backbox GI. Adjusted cRed_Ball and cGreen_Ball. Updated VR plaque. Some code cleanup.
'044 - apophis - Fixed random GI color option. Added info card.
'045 - bthlonewolf - added tint scaling in multiball, option in tweak to disable
'RC1 - apophis - Animated: bumpers, gates, switches. Fixed VR backglass positions. Adjusted rail position in desktop mode. Added VPW logo layer. Updated table credits. Tweaked ball tint colors. Updated relay sounds for flashers and gi and rubber sound (thanks RKP). Updated upperpf prim (thanks Sixtoe).
'RC2 - bthlonewolf - Tint z-axis fix, gate animations/speed fixed
'RC3 - tomate - Inserts glitches fixed (Thanks Mecha Enron), wall right and wall left added to Metal collection (thanks Studlygoorite)
'RC4 - bthlonewolf - gate fix for real this time
'RC5 - Mecha_Enron - Create option for siderails/lockbar
'RC6 - RobbyKingPin - Reverted back to the old balldrop sounds.
'RC7 - tomate - PF reflections fixed, blockdown bar position adjusted for desktop mode, ball rolling sounds adjusted
'RC8 - bthlonewolf - tilt switch changed to 9
'RC9 - Sixtoe - Removed stray rubber, tweaked plunger area, tweaked centre targets, removed collidable from many VR objects, updated credits, tidied file.

'RELEASE 1.0
'1.0.1 - tomate - space station toy material changed to avoid artifacts, right flipper texture fixed



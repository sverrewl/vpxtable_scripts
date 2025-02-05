'Lethal Weapon 3 (Data East 1992)
'https://www.ipdb.org/machine.cgi?id=1433


'VPW Lethal Weapons (v2.0)
'tomate - Blender toolkit renders. New primitives and table fixes. Updated physical ramps.
'apophis - Render integration and script optimizations, phyiscs updates, tweak menu options
'iaakki - Back panel mirror
'mcarter78 - ramp fixes
'DGrimmReaper - VR room fixes
'fluffhead - ball rolling sfx tweaks
'Wylte - some physics tweaks
'Niwak - Blender toolkit and pinball parts library
'Testers - AstroNasty, Studlygoorite, Darth Vito, Colvert, Flint Beastwood
'
'
'VPW Lethal Weapons (v1.0)
'Graphics: tomate
'Inserts: Sheltemke, oqqsan
'Scripting: oqqsan, apophis, tomate
'nFozzys physics: tomate, iaakki
'Fleep sounds: apophis
'Ramp Sounds: fluffhead35
'Flashes & Lighting Overhaul: tomate, Sixtoe
'Shadows: Wylte
'GI and Drop Target Shadows HauntFreaks
'VR Room and fixes: Sixtoe, Leojreimroc, 360 Room by pattyg234.
'Miscellaneous tweaks: Sixtoe, tomate, oqqsan
'Testing: Bord, Rik, oqqsan, VPW team
'Playfield and plastics: EBisLit
'Original VPX table by Javier & 32Assassin


'*********************************************************************************************************************************
' === TABLE OF CONTENTS  ===
'
' You can quickly jump to a section by searching the four letter tag (ZXXX)
'
' ZVAR: Constants and Global Variable
'   ZVLM: VLM Arrays
' ZTIM: Timers
' ZINI: Table Initialization and Exiting
'   ZOPT: User Options
'   ZMAT: General Math Functions
' ZANI: Misc Animations
'   ZBBR: Ball Brightness
'   ZRBR: Room Brightness
' ZKEY: Key Press Handling
' ZSOL: Solenoids
' ZFLA: Flashers
'   ZGIU: GI updates
' ZFLP: Flippers
' ZVUK: VUKs and Kickers
' ZDRN: Drain, Trough, and Ball Release
' ZSLG: Slingshot Animations
' ZSSC: Slingshot Corrections
' ZSWI: Switches
'   ZBRL: Ball Rolling and Drop Sounds
'   ZRRL: Ramp Rolling Sound Effects
' ZSHA: Ambient Ball Shadows
'   ZFLE: Fleep Mechanical Sounds
' ZNFF: Flipper Corrections
'   ZDMP: Rubber Dampeners
'   ZBOU: VPW TargetBouncer for targets and posts
'   ZRDT: Drop Targets
' ZRST: Stand-Up Targets
'   ZVRR: VR Room / VR Cabinet
'
'*********************************************************************************************************************************



Option Explicit
Randomize
SetLocale 1033

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0


'******************************************************
'  ZVAR: Constants and Global Variables
'******************************************************

Const Ballsize = 50             'Ball size must be 50
Const BallMass = 1              'Ball mass must be 1
Const tnob = 3
Const lob = 0

Dim tablewidth: tablewidth = Table1.width
Dim tableheight: tableheight = Table1.height
Dim bsLock, bsLock2, plungerIM
Dim LWBall1, LWBall2, LWBall3, gBOT

Dim DesktopMode, UseVPMDMD, VRRoom: VRRoom = 0
DesktopMode = Table1.ShowDT
If RenderingMode = 2 Then UseVPMDMD = True Else UseVPMDMD = DesktopMode

'  Standard definitions
Const cGameName = "lw3_301"   'PinMAME ROM name
Const UseVPMModSol = 2
Const UseSolenoids = 2      '1 = Normal Flippers, 2 = Fastflips
Const UseLamps = 1        '0 = Custom lamp handling, 1 = Built-in VPX handling (using light number in light timer)
Const UseSync = 0
Const HandleMech = 0
Const SSolenoidOn = ""      'Sound sample used for this, obsolete.
Const SSolenoidOff = ""     ' ^
Const SFlipperOn = ""     ' ^
Const SFlipperOff = ""      ' ^
Const SCoin = ""        ' ^

LoadVPM "03060000", "DE.VBS", 3.36



'******************************************************
'  ZVLM: VLM Arrays
'******************************************************


' VLM  Arrays - Start
' Arrays per baked part
Dim BP_BR1: BP_BR1=Array(BM_BR1, LM_FL_F126_BR1, LM_FL_F127_BR1, LM_FL_F128_BR1, LM_FL_F129_BR1, LM_FL_F130_BR1, LM_FL_F132_BR1, LM_FL_F109_BR1, LM_IN_L34_BR1, LM_IN_L42_BR1, LM_IN_L43_BR1, LM_FL_L44_BR1, LM_IN_L60_BR1, LM_IN_L61_BR1, LM_GIS_GI17_BR1, LM_GIS_GI22_BR1, LM_GI_BR1, LM_IN_l55_BR1, LM_IN_l56_BR1)
Dim BP_BR2: BP_BR2=Array(BM_BR2, LM_FL_F126_BR2, LM_FL_F127_BR2, LM_FL_F128_BR2, LM_FL_F130_BR2, LM_FL_F132_BR2, LM_FL_F109_BR2, LM_IN_L34_BR2, LM_IN_L43_BR2, LM_FL_L44_BR2, LM_IN_L60_BR2, LM_GIS_GI17_BR2, LM_GIS_GI19_BR2, LM_GIS_GI18_BR2, LM_GIS_GI20_BR2, LM_GIS_GI22_BR2, LM_GIS_GI26_BR2, LM_GI_BR2)
Dim BP_BR3: BP_BR3=Array(BM_BR3, LM_FL_F126_BR3, LM_FL_F127_BR3, LM_FL_F128_BR3, LM_FL_F129_BR3, LM_FL_F130_BR3, LM_FL_F132_BR3, LM_FL_F109_BR3, LM_IN_L34_BR3, LM_FL_L44_BR3, LM_IN_L60_BR3, LM_IN_L61_BR3, LM_GIS_GI17_BR3, LM_GIS_GI19_BR3, LM_GIS_GI26_BR3, LM_GI_BR3)
Dim BP_Gate1: BP_Gate1=Array(BM_Gate1, LM_FL_F125_Gate1, LM_FL_F126_Gate1, LM_FL_F127_Gate1, LM_GIS_GI01_Gate1, LM_GIS_GI12_Gate1, LM_GIS_GI18_Gate1, LM_GI_Gate1)
Dim BP_Gate5: BP_Gate5=Array(BM_Gate5, LM_FL_F126_Gate5, LM_FL_F129_Gate5, LM_FL_F130_Gate5, LM_FL_F109_Gate5, LM_FL_L44_Gate5, LM_IN_L61_Gate5, LM_GI_Gate5)
Dim BP_LEMK: BP_LEMK=Array(BM_LEMK, LM_GIS_GI06_LEMK, LM_GIS_GI08_LEMK, LM_GIS_GI14_LEMK)
Dim BP_LFlipper: BP_LFlipper=Array(BM_LFlipper, LM_IN_L17_LFlipper, LM_IN_L49_LFlipper, LM_IN_L50_LFlipper, LM_GIS_GI01_LFlipper, LM_GIS_GI06_LFlipper, LM_GIS_GI08_LFlipper, LM_GIS_GI10_LFlipper, LM_GIS_GI14_LFlipper)
Dim BP_LFlipperU: BP_LFlipperU=Array(BM_LFlipperU, LM_FL_F127_LFlipperU, LM_IN_L17_LFlipperU, LM_IN_L49_LFlipperU, LM_IN_L50_LFlipperU, LM_GIS_GI01_LFlipperU, LM_GIS_GI06_LFlipperU, LM_GIS_GI08_LFlipperU, LM_GIS_GI10_LFlipperU, LM_GIS_GI14_LFlipperU, LM_IN_l51_LFlipperU)
Dim BP_LSling1: BP_LSling1=Array(BM_LSling1, LM_FL_F126_LSling1, LM_GIS_GI01_LSling1, LM_GIS_GI06_LSling1, LM_GIS_GI08_LSling1, LM_GIS_GI10_LSling1, LM_GIS_GI12_LSling1, LM_GIS_GI14_LSling1)
Dim BP_LSling2: BP_LSling2=Array(BM_LSling2, LM_FL_F126_LSling2, LM_GIS_GI01_LSling2, LM_GIS_GI06_LSling2, LM_GIS_GI08_LSling2, LM_GIS_GI10_LSling2, LM_GIS_GI12_LSling2, LM_GIS_GI14_LSling2)
Dim BP_Layer1: BP_Layer1=Array(BM_Layer1, LM_FL_F125_Layer1, LM_FL_F126_Layer1, LM_FL_F127_Layer1, LM_FL_F128_Layer1, LM_FL_F132_Layer1, LM_FL_F109_Layer1, LM_IN_L18_Layer1, LM_IN_L24_Layer1, LM_IN_L27_Layer1, LM_IN_L41_Layer1, LM_IN_L42_Layer1, LM_IN_L43_Layer1, LM_FL_L44_Layer1, LM_IN_L45_Layer1, LM_IN_L46_Layer1, LM_IN_L47_Layer1, LM_IN_L48_Layer1, LM_IN_L49_Layer1, LM_IN_L61_Layer1, LM_GIS_GI01_Layer1, LM_GIS_GI06_Layer1, LM_GIS_GI08_Layer1, LM_GIS_GI10_Layer1, LM_GIS_GI12_Layer1, LM_GIS_GI14_Layer1, LM_GIS_GI18_Layer1, LM_GIS_GI20_Layer1, LM_GIS_GI22_Layer1, LM_GI_Layer1, LM_IN_l10_Layer1, LM_IN_l23_Layer1, LM_IN_l53_Layer1, LM_IN_L9_Layer1)
Dim BP_Layer2: BP_Layer2=Array(BM_Layer2, LM_FL_F125_Layer2, LM_FL_F126_Layer2, LM_FL_F132_Layer2, LM_IN_L18_Layer2, LM_IN_L22_Layer2, LM_IN_L27_Layer2, LM_FL_L44_Layer2, LM_IN_L48_Layer2, LM_IN_L49_Layer2, LM_IN_L50_Layer2, LM_GIS_GI17_Layer2, LM_GIS_GI19_Layer2, LM_GIS_GI01_Layer2, LM_GIS_GI06_Layer2, LM_GIS_GI08_Layer2, LM_GIS_GI10_Layer2, LM_GIS_GI12_Layer2, LM_GIS_GI14_Layer2, LM_GIS_GI18_Layer2, LM_GIS_GI20_Layer2, LM_GIS_GI22_Layer2, LM_GIS_GI30_Layer2, LM_GIS_GI32_Layer2, LM_GI_Layer2, LM_IN_l10_Layer2, LM_IN_l40_Layer2, LM_IN_l52_Layer2, LM_IN_l53_Layer2, LM_IN_l58_Layer2, LM_IN_L9_Layer2)
Dim BP_Parts: BP_Parts=Array(BM_Parts, LM_FL_F125_Parts, LM_FL_F116_Parts, LM_FL_F126_Parts, LM_FL_F127_Parts, LM_FL_F128_Parts, LM_FL_F129_Parts, LM_FL_F130_Parts, LM_FL_F131_Parts, LM_FL_F132_Parts, LM_FL_F109_Parts, LM_IN_L1_Parts, LM_IN_L17_Parts, LM_IN_L18_Parts, LM_IN_L19_Parts, LM_IN_L2_Parts, LM_IN_L21_Parts, LM_IN_L22_Parts, LM_IN_L24_Parts, LM_IN_L27_Parts, LM_IN_L3_Parts, LM_IN_L34_Parts, LM_IN_L41_Parts, LM_IN_L42_Parts, LM_IN_L43_Parts, LM_FL_L44_Parts, LM_IN_L45_Parts, LM_IN_L46_Parts, LM_IN_L47_Parts, LM_IN_L48_Parts, LM_IN_L49_Parts, LM_IN_L50_Parts, LM_IN_L60_Parts, LM_IN_L61_Parts, LM_GIS_GI17_Parts, LM_GIS_GI19_Parts, LM_GIS_GI01_Parts, LM_GIS_GI06_Parts, LM_GIS_GI08_Parts, LM_GIS_GI10_Parts, LM_GIS_GI12_Parts, LM_GIS_GI14_Parts, LM_GIS_GI18_Parts, LM_GIS_GI20_Parts, LM_GIS_GI22_Parts, LM_GIS_GI26_Parts, LM_GIS_GI30_Parts, LM_GIS_GI32_Parts, LM_GI_Parts, LM_IN_l10_Parts, LM_IN_l11_Parts, LM_IN_L16_Parts, LM_IN_l20_Parts, LM_IN_l23_Parts, LM_IN_l25_Parts, LM_IN_l26_Parts, LM_IN_l28_Parts, _
  LM_IN_l29_Parts, LM_IN_l30_Parts, LM_IN_l35_Parts, LM_IN_l36_Parts, LM_IN_l37_Parts, LM_IN_l38_Parts, LM_IN_l39_Parts, LM_IN_l40_Parts, LM_IN_l51_Parts, LM_IN_l52_Parts, LM_IN_l53_Parts, LM_IN_l54_Parts, LM_IN_l55_Parts, LM_IN_l56_Parts, LM_IN_l58_Parts, LM_IN_l62_Parts, LM_IN_l63_Parts, LM_IN_l64_Parts, LM_IN_L7_Parts, LM_IN_L9_Parts)
Dim BP_Playfield: BP_Playfield=Array(BM_Playfield, LM_FL_F125_Playfield, LM_FL_F126_Playfield, LM_FL_F127_Playfield, LM_FL_F128_Playfield, LM_FL_F129_Playfield, LM_FL_F130_Playfield, LM_FL_F131_Playfield, LM_FL_F132_Playfield, LM_FL_F109_Playfield, LM_IN_L1_Playfield, LM_IN_L18_Playfield, LM_IN_L19_Playfield, LM_IN_L2_Playfield, LM_IN_L21_Playfield, LM_IN_L22_Playfield, LM_IN_L24_Playfield, LM_IN_L27_Playfield, LM_IN_L3_Playfield, LM_IN_L34_Playfield, LM_IN_L4_Playfield, LM_IN_L41_Playfield, LM_IN_L42_Playfield, LM_IN_L43_Playfield, LM_FL_L44_Playfield, LM_IN_L45_Playfield, LM_IN_L46_Playfield, LM_IN_L47_Playfield, LM_IN_L48_Playfield, LM_IN_L49_Playfield, LM_IN_L5_Playfield, LM_IN_L6_Playfield, LM_IN_L60_Playfield, LM_IN_L61_Playfield, LM_GIS_GI17_Playfield, LM_GIS_GI19_Playfield, LM_GIS_GI01_Playfield, LM_GIS_GI06_Playfield, LM_GIS_GI08_Playfield, LM_GIS_GI10_Playfield, LM_GIS_GI12_Playfield, LM_GIS_GI14_Playfield, LM_GIS_GI18_Playfield, LM_GIS_GI20_Playfield, LM_GIS_GI22_Playfield, LM_GIS_GI26_Playfield, _
  LM_GIS_GI30_Playfield, LM_GIS_GI32_Playfield, LM_GI_Playfield, LM_IN_l20_Playfield, LM_IN_l23_Playfield, LM_IN_l25_Playfield, LM_IN_l56_Playfield, LM_IN_L7_Playfield)
Dim BP_REMK: BP_REMK=Array(BM_REMK, LM_GIS_GI01_REMK, LM_GIS_GI10_REMK, LM_GIS_GI12_REMK)
Dim BP_RFlipper: BP_RFlipper=Array(BM_RFlipper, LM_FL_F127_RFlipper, LM_IN_L17_RFlipper, LM_GIS_GI01_RFlipper, LM_GIS_GI06_RFlipper, LM_GIS_GI08_RFlipper, LM_GIS_GI10_RFlipper, LM_GIS_GI12_RFlipper, LM_IN_l52_RFlipper, LM_IN_l53_RFlipper)
Dim BP_RFlipperU: BP_RFlipperU=Array(BM_RFlipperU, LM_IN_L1_RFlipperU, LM_IN_L17_RFlipperU, LM_GIS_GI01_RFlipperU, LM_GIS_GI06_RFlipperU, LM_GIS_GI08_RFlipperU, LM_GIS_GI10_RFlipperU, LM_GIS_GI12_RFlipperU, LM_IN_l51_RFlipperU, LM_IN_l52_RFlipperU, LM_IN_l53_RFlipperU)
Dim BP_RSling1: BP_RSling1=Array(BM_RSling1, LM_FL_F126_RSling1, LM_GIS_GI01_RSling1, LM_GIS_GI06_RSling1, LM_GIS_GI08_RSling1, LM_GIS_GI10_RSling1, LM_GIS_GI12_RSling1, LM_GIS_GI14_RSling1, LM_IN_l28_RSling1, LM_IN_l29_RSling1)
Dim BP_RSling2: BP_RSling2=Array(BM_RSling2, LM_FL_F126_RSling2, LM_GIS_GI01_RSling2, LM_GIS_GI06_RSling2, LM_GIS_GI08_RSling2, LM_GIS_GI10_RSling2, LM_GIS_GI12_RSling2, LM_GIS_GI14_RSling2, LM_IN_l28_RSling2)
Dim BP_Rpost: BP_Rpost=Array(BM_Rpost, LM_FL_F126_Rpost, LM_FL_F109_Rpost, LM_IN_L18_Rpost, LM_FL_L44_Rpost, LM_GIS_GI01_Rpost, LM_GIS_GI10_Rpost, LM_GIS_GI12_Rpost)
Dim BP_Sw14: BP_Sw14=Array(BM_Sw14, LM_IN_L7_Sw14)
Dim BP_Sw21: BP_Sw21=Array(BM_Sw21, LM_FL_F126_Sw21, LM_FL_F128_Sw21, LM_FL_F130_Sw21, LM_FL_F109_Sw21, LM_IN_L42_Sw21, LM_IN_L43_Sw21, LM_FL_L44_Sw21, LM_GI_Sw21)
Dim BP_Sw22: BP_Sw22=Array(BM_Sw22, LM_FL_F130_Sw22, LM_GI_Sw22)
Dim BP_Sw28: BP_Sw28=Array(BM_Sw28, LM_GIS_GI06_Sw28, LM_GIS_GI08_Sw28, LM_GIS_GI14_Sw28)
Dim BP_Sw29: BP_Sw29=Array(BM_Sw29, LM_GIS_GI06_Sw29, LM_GIS_GI08_Sw29, LM_GIS_GI14_Sw29)
Dim BP_Sw36: BP_Sw36=Array(BM_Sw36, LM_GIS_GI01_Sw36, LM_GIS_GI10_Sw36, LM_GIS_GI12_Sw36)
Dim BP_Sw37: BP_Sw37=Array(BM_Sw37, LM_GIS_GI01_Sw37, LM_GIS_GI10_Sw37, LM_GIS_GI12_Sw37)
Dim BP_Sw41: BP_Sw41=Array(BM_Sw41, LM_FL_F126_Sw41, LM_FL_F109_Sw41, LM_FL_L44_Sw41, LM_GI_Sw41)
Dim BP_Sw42: BP_Sw42=Array(BM_Sw42, LM_FL_F126_Sw42, LM_FL_F129_Sw42, LM_FL_F130_Sw42, LM_FL_L44_Sw42, LM_GI_Sw42)
Dim BP_Sw43: BP_Sw43=Array(BM_Sw43, LM_FL_F126_Sw43, LM_FL_F129_Sw43, LM_FL_F130_Sw43, LM_FL_L44_Sw43, LM_GI_Sw43)
Dim BP_Sw47: BP_Sw47=Array(BM_Sw47, LM_FL_F127_Sw47, LM_FL_F128_Sw47, LM_FL_F132_Sw47, LM_FL_F109_Sw47, LM_IN_L41_Sw47, LM_IN_L42_Sw47, LM_IN_L43_Sw47, LM_FL_L44_Sw47, LM_IN_L45_Sw47, LM_GIS_GI26_Sw47, LM_GI_Sw47)
Dim BP_Sw48: BP_Sw48=Array(BM_Sw48, LM_FL_F130_Sw48, LM_FL_F132_Sw48, LM_IN_L61_Sw48, LM_GIS_GI22_Sw48, LM_GIS_GI26_Sw48, LM_GI_Sw48)
Dim BP_Sw54: BP_Sw54=Array(BM_Sw54, LM_FL_F128_Sw54, LM_FL_F109_Sw54, LM_GI_Sw54)
Dim BP_Sw55: BP_Sw55=Array(BM_Sw55)
Dim BP_sw17: BP_sw17=Array(BM_sw17, LM_FL_F126_sw17, LM_FL_F127_sw17, LM_FL_F128_sw17, LM_FL_F109_sw17, LM_IN_L43_sw17, LM_IN_L45_sw17, LM_IN_L46_sw17, LM_GIS_GI26_sw17, LM_GI_sw17)
Dim BP_sw18: BP_sw18=Array(BM_sw18, LM_FL_F126_sw18, LM_FL_F127_sw18, LM_FL_F109_sw18, LM_IN_L43_sw18, LM_IN_L45_sw18, LM_IN_L46_sw18, LM_IN_L47_sw18, LM_GIS_GI26_sw18, LM_GI_sw18)
Dim BP_sw19: BP_sw19=Array(BM_sw19, LM_FL_F126_sw19, LM_IN_L45_sw19, LM_IN_L46_sw19, LM_IN_L47_sw19, LM_IN_L48_sw19, LM_GIS_GI26_sw19, LM_GI_sw19)
Dim BP_sw20: BP_sw20=Array(BM_sw20, LM_FL_F126_sw20, LM_IN_L45_sw20, LM_IN_L46_sw20, LM_IN_L47_sw20, LM_IN_L48_sw20, LM_GIS_GI17_sw20, LM_GIS_GI26_sw20, LM_GI_sw20, LM_IN_l40_sw20)
Dim BP_sw25: BP_sw25=Array(BM_sw25, LM_FL_F126_sw25, LM_FL_F127_sw25, LM_FL_F131_sw25, LM_FL_F132_sw25, LM_FL_F109_sw25, LM_IN_L19_sw25, LM_GIS_GI19_sw25, LM_GIS_GI18_sw25, LM_GIS_GI20_sw25, LM_GIS_GI26_sw25, LM_GI_sw25, LM_IN_l20_sw25)
Dim BP_sw26: BP_sw26=Array(BM_sw26, LM_FL_F126_sw26, LM_FL_F127_sw26, LM_FL_F128_sw26, LM_FL_F131_sw26, LM_FL_F132_sw26, LM_FL_F109_sw26, LM_IN_L19_sw26, LM_IN_L21_sw26, LM_GIS_GI19_sw26, LM_GIS_GI18_sw26, LM_GIS_GI20_sw26, LM_GIS_GI26_sw26, LM_GI_sw26, LM_IN_l20_sw26)
Dim BP_sw27: BP_sw27=Array(BM_sw27, LM_FL_F126_sw27, LM_FL_F127_sw27, LM_FL_F132_sw27, LM_IN_L19_sw27, LM_IN_L21_sw27, LM_GIS_GI19_sw27, LM_GIS_GI18_sw27, LM_GIS_GI20_sw27, LM_GIS_GI22_sw27, LM_GIS_GI26_sw27, LM_GI_sw27, LM_IN_l20_sw27, LM_IN_l62_sw27, LM_IN_l63_sw27)
Dim BP_sw32: BP_sw32=Array(BM_sw32, LM_FL_F126_sw32, LM_FL_F130_sw32, LM_FL_L44_sw32, LM_GI_sw32)
Dim BP_sw33: BP_sw33=Array(BM_sw33, LM_FL_F127_sw33, LM_FL_F132_sw33, LM_IN_L22_sw33, LM_IN_L24_sw33, LM_GIS_GI12_sw33, LM_GIS_GI18_sw33, LM_GIS_GI20_sw33, LM_GIS_GI22_sw33, LM_GIS_GI26_sw33, LM_GI_sw33, LM_IN_l23_sw33)
Dim BP_sw34: BP_sw34=Array(BM_sw34, LM_FL_F127_sw34, LM_FL_F132_sw34, LM_IN_L22_sw34, LM_IN_L24_sw34, LM_GIS_GI10_sw34, LM_GIS_GI12_sw34, LM_GIS_GI18_sw34, LM_GIS_GI20_sw34, LM_GIS_GI22_sw34, LM_GIS_GI26_sw34, LM_GI_sw34, LM_IN_l23_sw34)
Dim BP_sw35: BP_sw35=Array(BM_sw35, LM_FL_F127_sw35, LM_FL_F132_sw35, LM_IN_L22_sw35, LM_IN_L24_sw35, LM_GIS_GI10_sw35, LM_GIS_GI12_sw35, LM_GIS_GI18_sw35, LM_GIS_GI20_sw35, LM_GIS_GI22_sw35, LM_GIS_GI26_sw35, LM_GI_sw35, LM_IN_l23_sw35)
Dim BP_sw39: BP_sw39=Array(BM_sw39, LM_FL_F126_sw39, LM_GIS_GI06_sw39, LM_GIS_GI08_sw39, LM_GIS_GI14_sw39, LM_GIS_GI30_sw39, LM_GIS_GI32_sw39, LM_GI_sw39, LM_IN_l10_sw39, LM_IN_l25_sw39, LM_IN_l38_sw39, LM_IN_l39_sw39, LM_IN_L9_sw39)
Dim BP_sw49: BP_sw49=Array(BM_sw49, LM_FL_F126_sw49, LM_FL_F127_sw49, LM_FL_F109_sw49, LM_IN_L42_sw49, LM_IN_L43_sw49, LM_FL_L44_sw49, LM_GIS_GI26_sw49, LM_GI_sw49)
Dim BP_sw50: BP_sw50=Array(BM_sw50, LM_FL_F126_sw50, LM_FL_F130_sw50, LM_FL_F109_sw50, LM_GI_sw50)
Dim BP_underPF: BP_underPF=Array(BM_underPF, LM_FL_F125_underPF, LM_FL_F116_underPF, LM_FL_F126_underPF, LM_FL_F127_underPF, LM_FL_F128_underPF, LM_FL_F129_underPF, LM_FL_F130_underPF, LM_FL_F131_underPF, LM_FL_F132_underPF, LM_FL_F109_underPF, LM_IN_L1_underPF, LM_IN_L17_underPF, LM_IN_L18_underPF, LM_IN_L19_underPF, LM_IN_L2_underPF, LM_IN_L21_underPF, LM_IN_L22_underPF, LM_IN_L24_underPF, LM_IN_L27_underPF, LM_IN_L3_underPF, LM_IN_L34_underPF, LM_IN_L4_underPF, LM_IN_L41_underPF, LM_IN_L42_underPF, LM_IN_L43_underPF, LM_FL_L44_underPF, LM_IN_L45_underPF, LM_IN_L46_underPF, LM_IN_L47_underPF, LM_IN_L48_underPF, LM_IN_L49_underPF, LM_IN_L5_underPF, LM_IN_L50_underPF, LM_IN_L6_underPF, LM_IN_L60_underPF, LM_GIS_GI17_underPF, LM_GIS_GI19_underPF, LM_GIS_GI01_underPF, LM_GIS_GI06_underPF, LM_GIS_GI08_underPF, LM_GIS_GI10_underPF, LM_GIS_GI12_underPF, LM_GIS_GI14_underPF, LM_GIS_GI18_underPF, LM_GIS_GI20_underPF, LM_GIS_GI22_underPF, LM_GIS_GI26_underPF, LM_GIS_GI30_underPF, LM_GIS_GI32_underPF, LM_GI_underPF, _
  LM_IN_l10_underPF, LM_IN_l11_underPF, LM_IN_l12_underPF, LM_IN_l13_underPF, LM_IN_l14_underPF, LM_IN_l15_underPF, LM_IN_L16_underPF, LM_IN_l20_underPF, LM_IN_l23_underPF, LM_IN_l25_underPF, LM_IN_l26_underPF, LM_IN_l28_underPF, LM_IN_l29_underPF, LM_IN_l30_underPF, LM_IN_l31_underPF, LM_IN_l32_underPF, LM_IN_l35_underPF, LM_IN_l36_underPF, LM_IN_l37_underPF, LM_IN_l38_underPF, LM_IN_l39_underPF, LM_IN_l40_underPF, LM_IN_l51_underPF, LM_IN_l52_underPF, LM_IN_l53_underPF, LM_IN_l54_underPF, LM_IN_l55_underPF, LM_IN_l56_underPF, LM_IN_l58_underPF, LM_IN_l59_underPF, LM_IN_l62_underPF, LM_IN_l63_underPF, LM_IN_l64_underPF, LM_IN_L7_underPF, LM_IN_L8_underPF, LM_IN_L9_underPF)
'' Arrays per lighting scenario
'Dim BL_FL_F109: BL_FL_F109=Array(LM_FL_F109_BR1, LM_FL_F109_BR2, LM_FL_F109_BR3, LM_FL_F109_Gate5, LM_FL_F109_Layer1, LM_FL_F109_Parts, LM_FL_F109_Playfield, LM_FL_F109_Rpost, LM_FL_F109_Sw21, LM_FL_F109_Sw41, LM_FL_F109_Sw47, LM_FL_F109_Sw54, LM_FL_F109_sw17, LM_FL_F109_sw18, LM_FL_F109_sw25, LM_FL_F109_sw26, LM_FL_F109_sw49, LM_FL_F109_sw50, LM_FL_F109_underPF)
'Dim BL_FL_F116: BL_FL_F116=Array(LM_FL_F116_Parts, LM_FL_F116_underPF)
'Dim BL_FL_F125: BL_FL_F125=Array(LM_FL_F125_Gate1, LM_FL_F125_Layer1, LM_FL_F125_Layer2, LM_FL_F125_Parts, LM_FL_F125_Playfield, LM_FL_F125_underPF)
'Dim BL_FL_F126: BL_FL_F126=Array(LM_FL_F126_BR1, LM_FL_F126_BR2, LM_FL_F126_BR3, LM_FL_F126_Gate1, LM_FL_F126_Gate5, LM_FL_F126_LSling1, LM_FL_F126_LSling2, LM_FL_F126_Layer1, LM_FL_F126_Layer2, LM_FL_F126_Parts, LM_FL_F126_Playfield, LM_FL_F126_RSling1, LM_FL_F126_RSling2, LM_FL_F126_Rpost, LM_FL_F126_Sw21, LM_FL_F126_Sw41, LM_FL_F126_Sw42, LM_FL_F126_Sw43, LM_FL_F126_sw17, LM_FL_F126_sw18, LM_FL_F126_sw19, LM_FL_F126_sw20, LM_FL_F126_sw25, LM_FL_F126_sw26, LM_FL_F126_sw27, LM_FL_F126_sw32, LM_FL_F126_sw39, LM_FL_F126_sw49, LM_FL_F126_sw50, LM_FL_F126_underPF)
'Dim BL_FL_F127: BL_FL_F127=Array(LM_FL_F127_BR1, LM_FL_F127_BR2, LM_FL_F127_BR3, LM_FL_F127_Gate1, LM_FL_F127_LFlipperU, LM_FL_F127_Layer1, LM_FL_F127_Parts, LM_FL_F127_Playfield, LM_FL_F127_RFlipper, LM_FL_F127_Sw47, LM_FL_F127_sw17, LM_FL_F127_sw18, LM_FL_F127_sw25, LM_FL_F127_sw26, LM_FL_F127_sw27, LM_FL_F127_sw33, LM_FL_F127_sw34, LM_FL_F127_sw35, LM_FL_F127_sw49, LM_FL_F127_underPF)
'Dim BL_FL_F128: BL_FL_F128=Array(LM_FL_F128_BR1, LM_FL_F128_BR2, LM_FL_F128_BR3, LM_FL_F128_Layer1, LM_FL_F128_Parts, LM_FL_F128_Playfield, LM_FL_F128_Sw21, LM_FL_F128_Sw47, LM_FL_F128_Sw54, LM_FL_F128_sw17, LM_FL_F128_sw26, LM_FL_F128_underPF)
'Dim BL_FL_F129: BL_FL_F129=Array(LM_FL_F129_BR1, LM_FL_F129_BR3, LM_FL_F129_Gate5, LM_FL_F129_Parts, LM_FL_F129_Playfield, LM_FL_F129_Sw42, LM_FL_F129_Sw43, LM_FL_F129_underPF)
'Dim BL_FL_F130: BL_FL_F130=Array(LM_FL_F130_BR1, LM_FL_F130_BR2, LM_FL_F130_BR3, LM_FL_F130_Gate5, LM_FL_F130_Parts, LM_FL_F130_Playfield, LM_FL_F130_Sw21, LM_FL_F130_Sw22, LM_FL_F130_Sw42, LM_FL_F130_Sw43, LM_FL_F130_Sw48, LM_FL_F130_sw32, LM_FL_F130_sw50, LM_FL_F130_underPF)
'Dim BL_FL_F131: BL_FL_F131=Array(LM_FL_F131_Parts, LM_FL_F131_Playfield, LM_FL_F131_sw25, LM_FL_F131_sw26, LM_FL_F131_underPF)
'Dim BL_FL_F132: BL_FL_F132=Array(LM_FL_F132_BR1, LM_FL_F132_BR2, LM_FL_F132_BR3, LM_FL_F132_Layer1, LM_FL_F132_Layer2, LM_FL_F132_Parts, LM_FL_F132_Playfield, LM_FL_F132_Sw47, LM_FL_F132_Sw48, LM_FL_F132_sw25, LM_FL_F132_sw26, LM_FL_F132_sw27, LM_FL_F132_sw33, LM_FL_F132_sw34, LM_FL_F132_sw35, LM_FL_F132_underPF)
'Dim BL_FL_L44: BL_FL_L44=Array(LM_FL_L44_BR1, LM_FL_L44_BR2, LM_FL_L44_BR3, LM_FL_L44_Gate5, LM_FL_L44_Layer1, LM_FL_L44_Layer2, LM_FL_L44_Parts, LM_FL_L44_Playfield, LM_FL_L44_Rpost, LM_FL_L44_Sw21, LM_FL_L44_Sw41, LM_FL_L44_Sw42, LM_FL_L44_Sw43, LM_FL_L44_Sw47, LM_FL_L44_sw32, LM_FL_L44_sw49, LM_FL_L44_underPF)
'Dim BL_GI: BL_GI=Array(LM_GI_BR1, LM_GI_BR2, LM_GI_BR3, LM_GI_Gate1, LM_GI_Gate5, LM_GI_Layer1, LM_GI_Layer2, LM_GI_Parts, LM_GI_Playfield, LM_GI_Sw21, LM_GI_Sw22, LM_GI_Sw41, LM_GI_Sw42, LM_GI_Sw43, LM_GI_Sw47, LM_GI_Sw48, LM_GI_Sw54, LM_GI_sw17, LM_GI_sw18, LM_GI_sw19, LM_GI_sw20, LM_GI_sw25, LM_GI_sw26, LM_GI_sw27, LM_GI_sw32, LM_GI_sw33, LM_GI_sw34, LM_GI_sw35, LM_GI_sw39, LM_GI_sw49, LM_GI_sw50, LM_GI_underPF)
'Dim BL_GIS_GI01: BL_GIS_GI01=Array(LM_GIS_GI01_Gate1, LM_GIS_GI01_LFlipper, LM_GIS_GI01_LFlipperU, LM_GIS_GI01_LSling1, LM_GIS_GI01_LSling2, LM_GIS_GI01_Layer1, LM_GIS_GI01_Layer2, LM_GIS_GI01_Parts, LM_GIS_GI01_Playfield, LM_GIS_GI01_REMK, LM_GIS_GI01_RFlipper, LM_GIS_GI01_RFlipperU, LM_GIS_GI01_RSling1, LM_GIS_GI01_RSling2, LM_GIS_GI01_Rpost, LM_GIS_GI01_Sw36, LM_GIS_GI01_Sw37, LM_GIS_GI01_underPF)
'Dim BL_GIS_GI06: BL_GIS_GI06=Array(LM_GIS_GI06_LEMK, LM_GIS_GI06_LFlipper, LM_GIS_GI06_LFlipperU, LM_GIS_GI06_LSling1, LM_GIS_GI06_LSling2, LM_GIS_GI06_Layer1, LM_GIS_GI06_Layer2, LM_GIS_GI06_Parts, LM_GIS_GI06_Playfield, LM_GIS_GI06_RFlipper, LM_GIS_GI06_RFlipperU, LM_GIS_GI06_RSling1, LM_GIS_GI06_RSling2, LM_GIS_GI06_Sw28, LM_GIS_GI06_Sw29, LM_GIS_GI06_sw39, LM_GIS_GI06_underPF)
'Dim BL_GIS_GI08: BL_GIS_GI08=Array(LM_GIS_GI08_LEMK, LM_GIS_GI08_LFlipper, LM_GIS_GI08_LFlipperU, LM_GIS_GI08_LSling1, LM_GIS_GI08_LSling2, LM_GIS_GI08_Layer1, LM_GIS_GI08_Layer2, LM_GIS_GI08_Parts, LM_GIS_GI08_Playfield, LM_GIS_GI08_RFlipper, LM_GIS_GI08_RFlipperU, LM_GIS_GI08_RSling1, LM_GIS_GI08_RSling2, LM_GIS_GI08_Sw28, LM_GIS_GI08_Sw29, LM_GIS_GI08_sw39, LM_GIS_GI08_underPF)
'Dim BL_GIS_GI10: BL_GIS_GI10=Array(LM_GIS_GI10_LFlipper, LM_GIS_GI10_LFlipperU, LM_GIS_GI10_LSling1, LM_GIS_GI10_LSling2, LM_GIS_GI10_Layer1, LM_GIS_GI10_Layer2, LM_GIS_GI10_Parts, LM_GIS_GI10_Playfield, LM_GIS_GI10_REMK, LM_GIS_GI10_RFlipper, LM_GIS_GI10_RFlipperU, LM_GIS_GI10_RSling1, LM_GIS_GI10_RSling2, LM_GIS_GI10_Rpost, LM_GIS_GI10_Sw36, LM_GIS_GI10_Sw37, LM_GIS_GI10_sw34, LM_GIS_GI10_sw35, LM_GIS_GI10_underPF)
'Dim BL_GIS_GI12: BL_GIS_GI12=Array(LM_GIS_GI12_Gate1, LM_GIS_GI12_LSling1, LM_GIS_GI12_LSling2, LM_GIS_GI12_Layer1, LM_GIS_GI12_Layer2, LM_GIS_GI12_Parts, LM_GIS_GI12_Playfield, LM_GIS_GI12_REMK, LM_GIS_GI12_RFlipper, LM_GIS_GI12_RFlipperU, LM_GIS_GI12_RSling1, LM_GIS_GI12_RSling2, LM_GIS_GI12_Rpost, LM_GIS_GI12_Sw36, LM_GIS_GI12_Sw37, LM_GIS_GI12_sw33, LM_GIS_GI12_sw34, LM_GIS_GI12_sw35, LM_GIS_GI12_underPF)
'Dim BL_GIS_GI14: BL_GIS_GI14=Array(LM_GIS_GI14_LEMK, LM_GIS_GI14_LFlipper, LM_GIS_GI14_LFlipperU, LM_GIS_GI14_LSling1, LM_GIS_GI14_LSling2, LM_GIS_GI14_Layer1, LM_GIS_GI14_Layer2, LM_GIS_GI14_Parts, LM_GIS_GI14_Playfield, LM_GIS_GI14_RSling1, LM_GIS_GI14_RSling2, LM_GIS_GI14_Sw28, LM_GIS_GI14_Sw29, LM_GIS_GI14_sw39, LM_GIS_GI14_underPF)
'Dim BL_GIS_GI17: BL_GIS_GI17=Array(LM_GIS_GI17_BR1, LM_GIS_GI17_BR2, LM_GIS_GI17_BR3, LM_GIS_GI17_Layer2, LM_GIS_GI17_Parts, LM_GIS_GI17_Playfield, LM_GIS_GI17_sw20, LM_GIS_GI17_underPF)
'Dim BL_GIS_GI18: BL_GIS_GI18=Array(LM_GIS_GI18_BR2, LM_GIS_GI18_Gate1, LM_GIS_GI18_Layer1, LM_GIS_GI18_Layer2, LM_GIS_GI18_Parts, LM_GIS_GI18_Playfield, LM_GIS_GI18_sw25, LM_GIS_GI18_sw26, LM_GIS_GI18_sw27, LM_GIS_GI18_sw33, LM_GIS_GI18_sw34, LM_GIS_GI18_sw35, LM_GIS_GI18_underPF)
'Dim BL_GIS_GI19: BL_GIS_GI19=Array(LM_GIS_GI19_BR2, LM_GIS_GI19_BR3, LM_GIS_GI19_Layer2, LM_GIS_GI19_Parts, LM_GIS_GI19_Playfield, LM_GIS_GI19_sw25, LM_GIS_GI19_sw26, LM_GIS_GI19_sw27, LM_GIS_GI19_underPF)
'Dim BL_GIS_GI20: BL_GIS_GI20=Array(LM_GIS_GI20_BR2, LM_GIS_GI20_Layer1, LM_GIS_GI20_Layer2, LM_GIS_GI20_Parts, LM_GIS_GI20_Playfield, LM_GIS_GI20_sw25, LM_GIS_GI20_sw26, LM_GIS_GI20_sw27, LM_GIS_GI20_sw33, LM_GIS_GI20_sw34, LM_GIS_GI20_sw35, LM_GIS_GI20_underPF)
'Dim BL_GIS_GI22: BL_GIS_GI22=Array(LM_GIS_GI22_BR1, LM_GIS_GI22_BR2, LM_GIS_GI22_Layer1, LM_GIS_GI22_Layer2, LM_GIS_GI22_Parts, LM_GIS_GI22_Playfield, LM_GIS_GI22_Sw48, LM_GIS_GI22_sw27, LM_GIS_GI22_sw33, LM_GIS_GI22_sw34, LM_GIS_GI22_sw35, LM_GIS_GI22_underPF)
'Dim BL_GIS_GI26: BL_GIS_GI26=Array(LM_GIS_GI26_BR2, LM_GIS_GI26_BR3, LM_GIS_GI26_Parts, LM_GIS_GI26_Playfield, LM_GIS_GI26_Sw47, LM_GIS_GI26_Sw48, LM_GIS_GI26_sw17, LM_GIS_GI26_sw18, LM_GIS_GI26_sw19, LM_GIS_GI26_sw20, LM_GIS_GI26_sw25, LM_GIS_GI26_sw26, LM_GIS_GI26_sw27, LM_GIS_GI26_sw33, LM_GIS_GI26_sw34, LM_GIS_GI26_sw35, LM_GIS_GI26_sw49, LM_GIS_GI26_underPF)
'Dim BL_GIS_GI30: BL_GIS_GI30=Array(LM_GIS_GI30_Layer2, LM_GIS_GI30_Parts, LM_GIS_GI30_Playfield, LM_GIS_GI30_sw39, LM_GIS_GI30_underPF)
'Dim BL_GIS_GI32: BL_GIS_GI32=Array(LM_GIS_GI32_Layer2, LM_GIS_GI32_Parts, LM_GIS_GI32_Playfield, LM_GIS_GI32_sw39, LM_GIS_GI32_underPF)
'Dim BL_IN_L1: BL_IN_L1=Array(LM_IN_L1_Parts, LM_IN_L1_Playfield, LM_IN_L1_RFlipperU, LM_IN_L1_underPF)
'Dim BL_IN_L16: BL_IN_L16=Array(LM_IN_L16_Parts, LM_IN_L16_underPF)
'Dim BL_IN_L17: BL_IN_L17=Array(LM_IN_L17_LFlipper, LM_IN_L17_LFlipperU, LM_IN_L17_Parts, LM_IN_L17_RFlipper, LM_IN_L17_RFlipperU, LM_IN_L17_underPF)
'Dim BL_IN_L18: BL_IN_L18=Array(LM_IN_L18_Layer1, LM_IN_L18_Layer2, LM_IN_L18_Parts, LM_IN_L18_Playfield, LM_IN_L18_Rpost, LM_IN_L18_underPF)
'Dim BL_IN_L19: BL_IN_L19=Array(LM_IN_L19_Parts, LM_IN_L19_Playfield, LM_IN_L19_sw25, LM_IN_L19_sw26, LM_IN_L19_sw27, LM_IN_L19_underPF)
'Dim BL_IN_L2: BL_IN_L2=Array(LM_IN_L2_Parts, LM_IN_L2_Playfield, LM_IN_L2_underPF)
'Dim BL_IN_L21: BL_IN_L21=Array(LM_IN_L21_Parts, LM_IN_L21_Playfield, LM_IN_L21_sw26, LM_IN_L21_sw27, LM_IN_L21_underPF)
'Dim BL_IN_L22: BL_IN_L22=Array(LM_IN_L22_Layer2, LM_IN_L22_Parts, LM_IN_L22_Playfield, LM_IN_L22_sw33, LM_IN_L22_sw34, LM_IN_L22_sw35, LM_IN_L22_underPF)
'Dim BL_IN_L24: BL_IN_L24=Array(LM_IN_L24_Layer1, LM_IN_L24_Parts, LM_IN_L24_Playfield, LM_IN_L24_sw33, LM_IN_L24_sw34, LM_IN_L24_sw35, LM_IN_L24_underPF)
'Dim BL_IN_L27: BL_IN_L27=Array(LM_IN_L27_Layer1, LM_IN_L27_Layer2, LM_IN_L27_Parts, LM_IN_L27_Playfield, LM_IN_L27_underPF)
'Dim BL_IN_L3: BL_IN_L3=Array(LM_IN_L3_Parts, LM_IN_L3_Playfield, LM_IN_L3_underPF)
'Dim BL_IN_L34: BL_IN_L34=Array(LM_IN_L34_BR1, LM_IN_L34_BR2, LM_IN_L34_BR3, LM_IN_L34_Parts, LM_IN_L34_Playfield, LM_IN_L34_underPF)
'Dim BL_IN_L4: BL_IN_L4=Array(LM_IN_L4_Playfield, LM_IN_L4_underPF)
'Dim BL_IN_L41: BL_IN_L41=Array(LM_IN_L41_Layer1, LM_IN_L41_Parts, LM_IN_L41_Playfield, LM_IN_L41_Sw47, LM_IN_L41_underPF)
'Dim BL_IN_L42: BL_IN_L42=Array(LM_IN_L42_BR1, LM_IN_L42_Layer1, LM_IN_L42_Parts, LM_IN_L42_Playfield, LM_IN_L42_Sw21, LM_IN_L42_Sw47, LM_IN_L42_sw49, LM_IN_L42_underPF)
'Dim BL_IN_L43: BL_IN_L43=Array(LM_IN_L43_BR1, LM_IN_L43_BR2, LM_IN_L43_Layer1, LM_IN_L43_Parts, LM_IN_L43_Playfield, LM_IN_L43_Sw21, LM_IN_L43_Sw47, LM_IN_L43_sw17, LM_IN_L43_sw18, LM_IN_L43_sw49, LM_IN_L43_underPF)
'Dim BL_IN_L45: BL_IN_L45=Array(LM_IN_L45_Layer1, LM_IN_L45_Parts, LM_IN_L45_Playfield, LM_IN_L45_Sw47, LM_IN_L45_sw17, LM_IN_L45_sw18, LM_IN_L45_sw19, LM_IN_L45_sw20, LM_IN_L45_underPF)
'Dim BL_IN_L46: BL_IN_L46=Array(LM_IN_L46_Layer1, LM_IN_L46_Parts, LM_IN_L46_Playfield, LM_IN_L46_sw17, LM_IN_L46_sw18, LM_IN_L46_sw19, LM_IN_L46_sw20, LM_IN_L46_underPF)
'Dim BL_IN_L47: BL_IN_L47=Array(LM_IN_L47_Layer1, LM_IN_L47_Parts, LM_IN_L47_Playfield, LM_IN_L47_sw18, LM_IN_L47_sw19, LM_IN_L47_sw20, LM_IN_L47_underPF)
'Dim BL_IN_L48: BL_IN_L48=Array(LM_IN_L48_Layer1, LM_IN_L48_Layer2, LM_IN_L48_Parts, LM_IN_L48_Playfield, LM_IN_L48_sw19, LM_IN_L48_sw20, LM_IN_L48_underPF)
'Dim BL_IN_L49: BL_IN_L49=Array(LM_IN_L49_LFlipper, LM_IN_L49_LFlipperU, LM_IN_L49_Layer1, LM_IN_L49_Layer2, LM_IN_L49_Parts, LM_IN_L49_Playfield, LM_IN_L49_underPF)
'Dim BL_IN_L5: BL_IN_L5=Array(LM_IN_L5_Playfield, LM_IN_L5_underPF)
'Dim BL_IN_L50: BL_IN_L50=Array(LM_IN_L50_LFlipper, LM_IN_L50_LFlipperU, LM_IN_L50_Layer2, LM_IN_L50_Parts, LM_IN_L50_underPF)
'Dim BL_IN_L6: BL_IN_L6=Array(LM_IN_L6_Playfield, LM_IN_L6_underPF)
'Dim BL_IN_L60: BL_IN_L60=Array(LM_IN_L60_BR1, LM_IN_L60_BR2, LM_IN_L60_BR3, LM_IN_L60_Parts, LM_IN_L60_Playfield, LM_IN_L60_underPF)
'Dim BL_IN_L61: BL_IN_L61=Array(LM_IN_L61_BR1, LM_IN_L61_BR3, LM_IN_L61_Gate5, LM_IN_L61_Layer1, LM_IN_L61_Parts, LM_IN_L61_Playfield, LM_IN_L61_Sw48)
'Dim BL_IN_L7: BL_IN_L7=Array(LM_IN_L7_Parts, LM_IN_L7_Playfield, LM_IN_L7_Sw14, LM_IN_L7_underPF)
'Dim BL_IN_L8: BL_IN_L8=Array(LM_IN_L8_underPF)
'Dim BL_IN_L9: BL_IN_L9=Array(LM_IN_L9_Layer1, LM_IN_L9_Layer2, LM_IN_L9_Parts, LM_IN_L9_sw39, LM_IN_L9_underPF)
'Dim BL_IN_l10: BL_IN_l10=Array(LM_IN_l10_Layer1, LM_IN_l10_Layer2, LM_IN_l10_Parts, LM_IN_l10_sw39, LM_IN_l10_underPF)
'Dim BL_IN_l11: BL_IN_l11=Array(LM_IN_l11_Parts, LM_IN_l11_underPF)
'Dim BL_IN_l12: BL_IN_l12=Array(LM_IN_l12_underPF)
'Dim BL_IN_l13: BL_IN_l13=Array(LM_IN_l13_underPF)
'Dim BL_IN_l14: BL_IN_l14=Array(LM_IN_l14_underPF)
'Dim BL_IN_l15: BL_IN_l15=Array(LM_IN_l15_underPF)
'Dim BL_IN_l20: BL_IN_l20=Array(LM_IN_l20_Parts, LM_IN_l20_Playfield, LM_IN_l20_sw25, LM_IN_l20_sw26, LM_IN_l20_sw27, LM_IN_l20_underPF)
'Dim BL_IN_l23: BL_IN_l23=Array(LM_IN_l23_Layer1, LM_IN_l23_Parts, LM_IN_l23_Playfield, LM_IN_l23_sw33, LM_IN_l23_sw34, LM_IN_l23_sw35, LM_IN_l23_underPF)
'Dim BL_IN_l25: BL_IN_l25=Array(LM_IN_l25_Parts, LM_IN_l25_Playfield, LM_IN_l25_sw39, LM_IN_l25_underPF)
'Dim BL_IN_l26: BL_IN_l26=Array(LM_IN_l26_Parts, LM_IN_l26_underPF)
'Dim BL_IN_l28: BL_IN_l28=Array(LM_IN_l28_Parts, LM_IN_l28_RSling1, LM_IN_l28_RSling2, LM_IN_l28_underPF)
'Dim BL_IN_l29: BL_IN_l29=Array(LM_IN_l29_Parts, LM_IN_l29_RSling1, LM_IN_l29_underPF)
'Dim BL_IN_l30: BL_IN_l30=Array(LM_IN_l30_Parts, LM_IN_l30_underPF)
'Dim BL_IN_l31: BL_IN_l31=Array(LM_IN_l31_underPF)
'Dim BL_IN_l32: BL_IN_l32=Array(LM_IN_l32_underPF)
'Dim BL_IN_l35: BL_IN_l35=Array(LM_IN_l35_Parts, LM_IN_l35_underPF)
'Dim BL_IN_l36: BL_IN_l36=Array(LM_IN_l36_Parts, LM_IN_l36_underPF)
'Dim BL_IN_l37: BL_IN_l37=Array(LM_IN_l37_Parts, LM_IN_l37_underPF)
'Dim BL_IN_l38: BL_IN_l38=Array(LM_IN_l38_Parts, LM_IN_l38_sw39, LM_IN_l38_underPF)
'Dim BL_IN_l39: BL_IN_l39=Array(LM_IN_l39_Parts, LM_IN_l39_sw39, LM_IN_l39_underPF)
'Dim BL_IN_l40: BL_IN_l40=Array(LM_IN_l40_Layer2, LM_IN_l40_Parts, LM_IN_l40_sw20, LM_IN_l40_underPF)
'Dim BL_IN_l51: BL_IN_l51=Array(LM_IN_l51_LFlipperU, LM_IN_l51_Parts, LM_IN_l51_RFlipperU, LM_IN_l51_underPF)
'Dim BL_IN_l52: BL_IN_l52=Array(LM_IN_l52_Layer2, LM_IN_l52_Parts, LM_IN_l52_RFlipper, LM_IN_l52_RFlipperU, LM_IN_l52_underPF)
'Dim BL_IN_l53: BL_IN_l53=Array(LM_IN_l53_Layer1, LM_IN_l53_Layer2, LM_IN_l53_Parts, LM_IN_l53_RFlipper, LM_IN_l53_RFlipperU, LM_IN_l53_underPF)
'Dim BL_IN_l54: BL_IN_l54=Array(LM_IN_l54_Parts, LM_IN_l54_underPF)
'Dim BL_IN_l55: BL_IN_l55=Array(LM_IN_l55_BR1, LM_IN_l55_Parts, LM_IN_l55_underPF)
'Dim BL_IN_l56: BL_IN_l56=Array(LM_IN_l56_BR1, LM_IN_l56_Parts, LM_IN_l56_Playfield, LM_IN_l56_underPF)
'Dim BL_IN_l58: BL_IN_l58=Array(LM_IN_l58_Layer2, LM_IN_l58_Parts, LM_IN_l58_underPF)
'Dim BL_IN_l59: BL_IN_l59=Array(LM_IN_l59_underPF)
'Dim BL_IN_l62: BL_IN_l62=Array(LM_IN_l62_Parts, LM_IN_l62_sw27, LM_IN_l62_underPF)
'Dim BL_IN_l63: BL_IN_l63=Array(LM_IN_l63_Parts, LM_IN_l63_sw27, LM_IN_l63_underPF)
'Dim BL_IN_l64: BL_IN_l64=Array(LM_IN_l64_Parts, LM_IN_l64_underPF)
'Dim BL_Room: BL_Room=Array(BM_BR1, BM_BR2, BM_BR3, BM_Gate1, BM_Gate5, BM_LEMK, BM_LFlipper, BM_LFlipperU, BM_LSling1, BM_LSling2, BM_Layer1, BM_Layer2, BM_Parts, BM_Playfield, BM_REMK, BM_RFlipper, BM_RFlipperU, BM_RSling1, BM_RSling2, BM_Rpost, BM_Sw14, BM_Sw21, BM_Sw22, BM_Sw28, BM_Sw29, BM_Sw36, BM_Sw37, BM_Sw41, BM_Sw42, BM_Sw43, BM_Sw47, BM_Sw48, BM_Sw54, BM_Sw55, BM_sw17, BM_sw18, BM_sw19, BM_sw20, BM_sw25, BM_sw26, BM_sw27, BM_sw32, BM_sw33, BM_sw34, BM_sw35, BM_sw39, BM_sw49, BM_sw50, BM_underPF)
'' Global arrays
'Dim BG_Bakemap: BG_Bakemap=Array(BM_BR1, BM_BR2, BM_BR3, BM_Gate1, BM_Gate5, BM_LEMK, BM_LFlipper, BM_LFlipperU, BM_LSling1, BM_LSling2, BM_Layer1, BM_Layer2, BM_Parts, BM_Playfield, BM_REMK, BM_RFlipper, BM_RFlipperU, BM_RSling1, BM_RSling2, BM_Rpost, BM_Sw14, BM_Sw21, BM_Sw22, BM_Sw28, BM_Sw29, BM_Sw36, BM_Sw37, BM_Sw41, BM_Sw42, BM_Sw43, BM_Sw47, BM_Sw48, BM_Sw54, BM_Sw55, BM_sw17, BM_sw18, BM_sw19, BM_sw20, BM_sw25, BM_sw26, BM_sw27, BM_sw32, BM_sw33, BM_sw34, BM_sw35, BM_sw39, BM_sw49, BM_sw50, BM_underPF)
'Dim BG_Lightmap: BG_Lightmap=Array(LM_FL_F109_BR1, LM_FL_F109_BR2, LM_FL_F109_BR3, LM_FL_F109_Gate5, LM_FL_F109_Layer1, LM_FL_F109_Parts, LM_FL_F109_Playfield, LM_FL_F109_Rpost, LM_FL_F109_Sw21, LM_FL_F109_Sw41, LM_FL_F109_Sw47, LM_FL_F109_Sw54, LM_FL_F109_sw17, LM_FL_F109_sw18, LM_FL_F109_sw25, LM_FL_F109_sw26, LM_FL_F109_sw49, LM_FL_F109_sw50, LM_FL_F109_underPF, LM_FL_F116_Parts, LM_FL_F116_underPF, LM_FL_F125_Gate1, LM_FL_F125_Layer1, LM_FL_F125_Layer2, LM_FL_F125_Parts, LM_FL_F125_Playfield, LM_FL_F125_underPF, LM_FL_F126_BR1, LM_FL_F126_BR2, LM_FL_F126_BR3, LM_FL_F126_Gate1, LM_FL_F126_Gate5, LM_FL_F126_LSling1, LM_FL_F126_LSling2, LM_FL_F126_Layer1, LM_FL_F126_Layer2, LM_FL_F126_Parts, LM_FL_F126_Playfield, LM_FL_F126_RSling1, LM_FL_F126_RSling2, LM_FL_F126_Rpost, LM_FL_F126_Sw21, LM_FL_F126_Sw41, LM_FL_F126_Sw42, LM_FL_F126_Sw43, LM_FL_F126_sw17, LM_FL_F126_sw18, LM_FL_F126_sw19, LM_FL_F126_sw20, LM_FL_F126_sw25, LM_FL_F126_sw26, LM_FL_F126_sw27, LM_FL_F126_sw32, LM_FL_F126_sw39, LM_FL_F126_sw49, _
' LM_FL_F126_sw50, LM_FL_F126_underPF, LM_FL_F127_BR1, LM_FL_F127_BR2, LM_FL_F127_BR3, LM_FL_F127_Gate1, LM_FL_F127_LFlipperU, LM_FL_F127_Layer1, LM_FL_F127_Parts, LM_FL_F127_Playfield, LM_FL_F127_RFlipper, LM_FL_F127_Sw47, LM_FL_F127_sw17, LM_FL_F127_sw18, LM_FL_F127_sw25, LM_FL_F127_sw26, LM_FL_F127_sw27, LM_FL_F127_sw33, LM_FL_F127_sw34, LM_FL_F127_sw35, LM_FL_F127_sw49, LM_FL_F127_underPF, LM_FL_F128_BR1, LM_FL_F128_BR2, LM_FL_F128_BR3, LM_FL_F128_Layer1, LM_FL_F128_Parts, LM_FL_F128_Playfield, LM_FL_F128_Sw21, LM_FL_F128_Sw47, LM_FL_F128_Sw54, LM_FL_F128_sw17, LM_FL_F128_sw26, LM_FL_F128_underPF, LM_FL_F129_BR1, LM_FL_F129_BR3, LM_FL_F129_Gate5, LM_FL_F129_Parts, LM_FL_F129_Playfield, LM_FL_F129_Sw42, LM_FL_F129_Sw43, LM_FL_F129_underPF, LM_FL_F130_BR1, LM_FL_F130_BR2, LM_FL_F130_BR3, LM_FL_F130_Gate5, LM_FL_F130_Parts, LM_FL_F130_Playfield, LM_FL_F130_Sw21, LM_FL_F130_Sw22, LM_FL_F130_Sw42, LM_FL_F130_Sw43, LM_FL_F130_Sw48, LM_FL_F130_sw32, LM_FL_F130_sw50, LM_FL_F130_underPF, LM_FL_F131_Parts, _
' LM_FL_F131_Playfield, LM_FL_F131_sw25, LM_FL_F131_sw26, LM_FL_F131_underPF, LM_FL_F132_BR1, LM_FL_F132_BR2, LM_FL_F132_BR3, LM_FL_F132_Layer1, LM_FL_F132_Layer2, LM_FL_F132_Parts, LM_FL_F132_Playfield, LM_FL_F132_Sw47, LM_FL_F132_Sw48, LM_FL_F132_sw25, LM_FL_F132_sw26, LM_FL_F132_sw27, LM_FL_F132_sw33, LM_FL_F132_sw34, LM_FL_F132_sw35, LM_FL_F132_underPF, LM_FL_L44_BR1, LM_FL_L44_BR2, LM_FL_L44_BR3, LM_FL_L44_Gate5, LM_FL_L44_Layer1, LM_FL_L44_Layer2, LM_FL_L44_Parts, LM_FL_L44_Playfield, LM_FL_L44_Rpost, LM_FL_L44_Sw21, LM_FL_L44_Sw41, LM_FL_L44_Sw42, LM_FL_L44_Sw43, LM_FL_L44_Sw47, LM_FL_L44_sw32, LM_FL_L44_sw49, LM_FL_L44_underPF, LM_GI_BR1, LM_GI_BR2, LM_GI_BR3, LM_GI_Gate1, LM_GI_Gate5, LM_GI_Layer1, LM_GI_Layer2, LM_GI_Parts, LM_GI_Playfield, LM_GI_Sw21, LM_GI_Sw22, LM_GI_Sw41, LM_GI_Sw42, LM_GI_Sw43, LM_GI_Sw47, LM_GI_Sw48, LM_GI_Sw54, LM_GI_sw17, LM_GI_sw18, LM_GI_sw19, LM_GI_sw20, LM_GI_sw25, LM_GI_sw26, LM_GI_sw27, LM_GI_sw32, LM_GI_sw33, LM_GI_sw34, LM_GI_sw35, LM_GI_sw39, LM_GI_sw49, LM_GI_sw50, _
' LM_GI_underPF, LM_GIS_GI01_Gate1, LM_GIS_GI01_LFlipper, LM_GIS_GI01_LFlipperU, LM_GIS_GI01_LSling1, LM_GIS_GI01_LSling2, LM_GIS_GI01_Layer1, LM_GIS_GI01_Layer2, LM_GIS_GI01_Parts, LM_GIS_GI01_Playfield, LM_GIS_GI01_REMK, LM_GIS_GI01_RFlipper, LM_GIS_GI01_RFlipperU, LM_GIS_GI01_RSling1, LM_GIS_GI01_RSling2, LM_GIS_GI01_Rpost, LM_GIS_GI01_Sw36, LM_GIS_GI01_Sw37, LM_GIS_GI01_underPF, LM_GIS_GI06_LEMK, LM_GIS_GI06_LFlipper, LM_GIS_GI06_LFlipperU, LM_GIS_GI06_LSling1, LM_GIS_GI06_LSling2, LM_GIS_GI06_Layer1, LM_GIS_GI06_Layer2, LM_GIS_GI06_Parts, LM_GIS_GI06_Playfield, LM_GIS_GI06_RFlipper, LM_GIS_GI06_RFlipperU, LM_GIS_GI06_RSling1, LM_GIS_GI06_RSling2, LM_GIS_GI06_Sw28, LM_GIS_GI06_Sw29, LM_GIS_GI06_sw39, LM_GIS_GI06_underPF, LM_GIS_GI08_LEMK, LM_GIS_GI08_LFlipper, LM_GIS_GI08_LFlipperU, LM_GIS_GI08_LSling1, LM_GIS_GI08_LSling2, LM_GIS_GI08_Layer1, LM_GIS_GI08_Layer2, LM_GIS_GI08_Parts, LM_GIS_GI08_Playfield, LM_GIS_GI08_RFlipper, LM_GIS_GI08_RFlipperU, LM_GIS_GI08_RSling1, LM_GIS_GI08_RSling2, LM_GIS_GI08_Sw28, _
' LM_GIS_GI08_Sw29, LM_GIS_GI08_sw39, LM_GIS_GI08_underPF, LM_GIS_GI10_LFlipper, LM_GIS_GI10_LFlipperU, LM_GIS_GI10_LSling1, LM_GIS_GI10_LSling2, LM_GIS_GI10_Layer1, LM_GIS_GI10_Layer2, LM_GIS_GI10_Parts, LM_GIS_GI10_Playfield, LM_GIS_GI10_REMK, LM_GIS_GI10_RFlipper, LM_GIS_GI10_RFlipperU, LM_GIS_GI10_RSling1, LM_GIS_GI10_RSling2, LM_GIS_GI10_Rpost, LM_GIS_GI10_Sw36, LM_GIS_GI10_Sw37, LM_GIS_GI10_sw34, LM_GIS_GI10_sw35, LM_GIS_GI10_underPF, LM_GIS_GI12_Gate1, LM_GIS_GI12_LSling1, LM_GIS_GI12_LSling2, LM_GIS_GI12_Layer1, LM_GIS_GI12_Layer2, LM_GIS_GI12_Parts, LM_GIS_GI12_Playfield, LM_GIS_GI12_REMK, LM_GIS_GI12_RFlipper, LM_GIS_GI12_RFlipperU, LM_GIS_GI12_RSling1, LM_GIS_GI12_RSling2, LM_GIS_GI12_Rpost, LM_GIS_GI12_Sw36, LM_GIS_GI12_Sw37, LM_GIS_GI12_sw33, LM_GIS_GI12_sw34, LM_GIS_GI12_sw35, LM_GIS_GI12_underPF, LM_GIS_GI14_LEMK, LM_GIS_GI14_LFlipper, LM_GIS_GI14_LFlipperU, LM_GIS_GI14_LSling1, LM_GIS_GI14_LSling2, LM_GIS_GI14_Layer1, LM_GIS_GI14_Layer2, LM_GIS_GI14_Parts, LM_GIS_GI14_Playfield, _
' LM_GIS_GI14_RSling1, LM_GIS_GI14_RSling2, LM_GIS_GI14_Sw28, LM_GIS_GI14_Sw29, LM_GIS_GI14_sw39, LM_GIS_GI14_underPF, LM_GIS_GI17_BR1, LM_GIS_GI17_BR2, LM_GIS_GI17_BR3, LM_GIS_GI17_Layer2, LM_GIS_GI17_Parts, LM_GIS_GI17_Playfield, LM_GIS_GI17_sw20, LM_GIS_GI17_underPF, LM_GIS_GI18_BR2, LM_GIS_GI18_Gate1, LM_GIS_GI18_Layer1, LM_GIS_GI18_Layer2, LM_GIS_GI18_Parts, LM_GIS_GI18_Playfield, LM_GIS_GI18_sw25, LM_GIS_GI18_sw26, LM_GIS_GI18_sw27, LM_GIS_GI18_sw33, LM_GIS_GI18_sw34, LM_GIS_GI18_sw35, LM_GIS_GI18_underPF, LM_GIS_GI19_BR2, LM_GIS_GI19_BR3, LM_GIS_GI19_Layer2, LM_GIS_GI19_Parts, LM_GIS_GI19_Playfield, LM_GIS_GI19_sw25, LM_GIS_GI19_sw26, LM_GIS_GI19_sw27, LM_GIS_GI19_underPF, LM_GIS_GI20_BR2, LM_GIS_GI20_Layer1, LM_GIS_GI20_Layer2, LM_GIS_GI20_Parts, LM_GIS_GI20_Playfield, LM_GIS_GI20_sw25, LM_GIS_GI20_sw26, LM_GIS_GI20_sw27, LM_GIS_GI20_sw33, LM_GIS_GI20_sw34, LM_GIS_GI20_sw35, LM_GIS_GI20_underPF, LM_GIS_GI22_BR1, LM_GIS_GI22_BR2, LM_GIS_GI22_Layer1, LM_GIS_GI22_Layer2, LM_GIS_GI22_Parts, _
' LM_GIS_GI22_Playfield, LM_GIS_GI22_Sw48, LM_GIS_GI22_sw27, LM_GIS_GI22_sw33, LM_GIS_GI22_sw34, LM_GIS_GI22_sw35, LM_GIS_GI22_underPF, LM_GIS_GI26_BR2, LM_GIS_GI26_BR3, LM_GIS_GI26_Parts, LM_GIS_GI26_Playfield, LM_GIS_GI26_Sw47, LM_GIS_GI26_Sw48, LM_GIS_GI26_sw17, LM_GIS_GI26_sw18, LM_GIS_GI26_sw19, LM_GIS_GI26_sw20, LM_GIS_GI26_sw25, LM_GIS_GI26_sw26, LM_GIS_GI26_sw27, LM_GIS_GI26_sw33, LM_GIS_GI26_sw34, LM_GIS_GI26_sw35, LM_GIS_GI26_sw49, LM_GIS_GI26_underPF, LM_GIS_GI30_Layer2, LM_GIS_GI30_Parts, LM_GIS_GI30_Playfield, LM_GIS_GI30_sw39, LM_GIS_GI30_underPF, LM_GIS_GI32_Layer2, LM_GIS_GI32_Parts, LM_GIS_GI32_Playfield, LM_GIS_GI32_sw39, LM_GIS_GI32_underPF, LM_IN_L1_Parts, LM_IN_L1_Playfield, LM_IN_L1_RFlipperU, LM_IN_L1_underPF, LM_IN_L16_Parts, LM_IN_L16_underPF, LM_IN_L17_LFlipper, LM_IN_L17_LFlipperU, LM_IN_L17_Parts, LM_IN_L17_RFlipper, LM_IN_L17_RFlipperU, LM_IN_L17_underPF, LM_IN_L18_Layer1, LM_IN_L18_Layer2, LM_IN_L18_Parts, LM_IN_L18_Playfield, LM_IN_L18_Rpost, LM_IN_L18_underPF, LM_IN_L19_Parts, _
' LM_IN_L19_Playfield, LM_IN_L19_sw25, LM_IN_L19_sw26, LM_IN_L19_sw27, LM_IN_L19_underPF, LM_IN_L2_Parts, LM_IN_L2_Playfield, LM_IN_L2_underPF, LM_IN_L21_Parts, LM_IN_L21_Playfield, LM_IN_L21_sw26, LM_IN_L21_sw27, LM_IN_L21_underPF, LM_IN_L22_Layer2, LM_IN_L22_Parts, LM_IN_L22_Playfield, LM_IN_L22_sw33, LM_IN_L22_sw34, LM_IN_L22_sw35, LM_IN_L22_underPF, LM_IN_L24_Layer1, LM_IN_L24_Parts, LM_IN_L24_Playfield, LM_IN_L24_sw33, LM_IN_L24_sw34, LM_IN_L24_sw35, LM_IN_L24_underPF, LM_IN_L27_Layer1, LM_IN_L27_Layer2, LM_IN_L27_Parts, LM_IN_L27_Playfield, LM_IN_L27_underPF, LM_IN_L3_Parts, LM_IN_L3_Playfield, LM_IN_L3_underPF, LM_IN_L34_BR1, LM_IN_L34_BR2, LM_IN_L34_BR3, LM_IN_L34_Parts, LM_IN_L34_Playfield, LM_IN_L34_underPF, LM_IN_L4_Playfield, LM_IN_L4_underPF, LM_IN_L41_Layer1, LM_IN_L41_Parts, LM_IN_L41_Playfield, LM_IN_L41_Sw47, LM_IN_L41_underPF, LM_IN_L42_BR1, LM_IN_L42_Layer1, LM_IN_L42_Parts, LM_IN_L42_Playfield, LM_IN_L42_Sw21, LM_IN_L42_Sw47, LM_IN_L42_sw49, LM_IN_L42_underPF, LM_IN_L43_BR1, LM_IN_L43_BR2, _
' LM_IN_L43_Layer1, LM_IN_L43_Parts, LM_IN_L43_Playfield, LM_IN_L43_Sw21, LM_IN_L43_Sw47, LM_IN_L43_sw17, LM_IN_L43_sw18, LM_IN_L43_sw49, LM_IN_L43_underPF, LM_IN_L45_Layer1, LM_IN_L45_Parts, LM_IN_L45_Playfield, LM_IN_L45_Sw47, LM_IN_L45_sw17, LM_IN_L45_sw18, LM_IN_L45_sw19, LM_IN_L45_sw20, LM_IN_L45_underPF, LM_IN_L46_Layer1, LM_IN_L46_Parts, LM_IN_L46_Playfield, LM_IN_L46_sw17, LM_IN_L46_sw18, LM_IN_L46_sw19, LM_IN_L46_sw20, LM_IN_L46_underPF, LM_IN_L47_Layer1, LM_IN_L47_Parts, LM_IN_L47_Playfield, LM_IN_L47_sw18, LM_IN_L47_sw19, LM_IN_L47_sw20, LM_IN_L47_underPF, LM_IN_L48_Layer1, LM_IN_L48_Layer2, LM_IN_L48_Parts, LM_IN_L48_Playfield, LM_IN_L48_sw19, LM_IN_L48_sw20, LM_IN_L48_underPF, LM_IN_L49_LFlipper, LM_IN_L49_LFlipperU, LM_IN_L49_Layer1, LM_IN_L49_Layer2, LM_IN_L49_Parts, LM_IN_L49_Playfield, LM_IN_L49_underPF, LM_IN_L5_Playfield, LM_IN_L5_underPF, LM_IN_L50_LFlipper, LM_IN_L50_LFlipperU, LM_IN_L50_Layer2, LM_IN_L50_Parts, LM_IN_L50_underPF, LM_IN_L6_Playfield, LM_IN_L6_underPF, LM_IN_L60_BR1, _
' LM_IN_L60_BR2, LM_IN_L60_BR3, LM_IN_L60_Parts, LM_IN_L60_Playfield, LM_IN_L60_underPF, LM_IN_L61_BR1, LM_IN_L61_BR3, LM_IN_L61_Gate5, LM_IN_L61_Layer1, LM_IN_L61_Parts, LM_IN_L61_Playfield, LM_IN_L61_Sw48, LM_IN_L7_Parts, LM_IN_L7_Playfield, LM_IN_L7_Sw14, LM_IN_L7_underPF, LM_IN_L8_underPF, LM_IN_L9_Layer1, LM_IN_L9_Layer2, LM_IN_L9_Parts, LM_IN_L9_sw39, LM_IN_L9_underPF, LM_IN_l10_Layer1, LM_IN_l10_Layer2, LM_IN_l10_Parts, LM_IN_l10_sw39, LM_IN_l10_underPF, LM_IN_l11_Parts, LM_IN_l11_underPF, LM_IN_l12_underPF, LM_IN_l13_underPF, LM_IN_l14_underPF, LM_IN_l15_underPF, LM_IN_l20_Parts, LM_IN_l20_Playfield, LM_IN_l20_sw25, LM_IN_l20_sw26, LM_IN_l20_sw27, LM_IN_l20_underPF, LM_IN_l23_Layer1, LM_IN_l23_Parts, LM_IN_l23_Playfield, LM_IN_l23_sw33, LM_IN_l23_sw34, LM_IN_l23_sw35, LM_IN_l23_underPF, LM_IN_l25_Parts, LM_IN_l25_Playfield, LM_IN_l25_sw39, LM_IN_l25_underPF, LM_IN_l26_Parts, LM_IN_l26_underPF, LM_IN_l28_Parts, LM_IN_l28_RSling1, LM_IN_l28_RSling2, LM_IN_l28_underPF, LM_IN_l29_Parts, LM_IN_l29_RSling1, _
' LM_IN_l29_underPF, LM_IN_l30_Parts, LM_IN_l30_underPF, LM_IN_l31_underPF, LM_IN_l32_underPF, LM_IN_l35_Parts, LM_IN_l35_underPF, LM_IN_l36_Parts, LM_IN_l36_underPF, LM_IN_l37_Parts, LM_IN_l37_underPF, LM_IN_l38_Parts, LM_IN_l38_sw39, LM_IN_l38_underPF, LM_IN_l39_Parts, LM_IN_l39_sw39, LM_IN_l39_underPF, LM_IN_l40_Layer2, LM_IN_l40_Parts, LM_IN_l40_sw20, LM_IN_l40_underPF, LM_IN_l51_LFlipperU, LM_IN_l51_Parts, LM_IN_l51_RFlipperU, LM_IN_l51_underPF, LM_IN_l52_Layer2, LM_IN_l52_Parts, LM_IN_l52_RFlipper, LM_IN_l52_RFlipperU, LM_IN_l52_underPF, LM_IN_l53_Layer1, LM_IN_l53_Layer2, LM_IN_l53_Parts, LM_IN_l53_RFlipper, LM_IN_l53_RFlipperU, LM_IN_l53_underPF, LM_IN_l54_Parts, LM_IN_l54_underPF, LM_IN_l55_BR1, LM_IN_l55_Parts, LM_IN_l55_underPF, LM_IN_l56_BR1, LM_IN_l56_Parts, LM_IN_l56_Playfield, LM_IN_l56_underPF, LM_IN_l58_Layer2, LM_IN_l58_Parts, LM_IN_l58_underPF, LM_IN_l59_underPF, LM_IN_l62_Parts, LM_IN_l62_sw27, LM_IN_l62_underPF, LM_IN_l63_Parts, LM_IN_l63_sw27, LM_IN_l63_underPF, LM_IN_l64_Parts, _
' LM_IN_l64_underPF)
'Dim BG_All: BG_All=Array(BM_BR1, BM_BR2, BM_BR3, BM_Gate1, BM_Gate5, BM_LEMK, BM_LFlipper, BM_LFlipperU, BM_LSling1, BM_LSling2, BM_Layer1, BM_Layer2, BM_Parts, BM_Playfield, BM_REMK, BM_RFlipper, BM_RFlipperU, BM_RSling1, BM_RSling2, BM_Rpost, BM_Sw14, BM_Sw21, BM_Sw22, BM_Sw28, BM_Sw29, BM_Sw36, BM_Sw37, BM_Sw41, BM_Sw42, BM_Sw43, BM_Sw47, BM_Sw48, BM_Sw54, BM_Sw55, BM_sw17, BM_sw18, BM_sw19, BM_sw20, BM_sw25, BM_sw26, BM_sw27, BM_sw32, BM_sw33, BM_sw34, BM_sw35, BM_sw39, BM_sw49, BM_sw50, BM_underPF, LM_FL_F109_BR1, LM_FL_F109_BR2, LM_FL_F109_BR3, LM_FL_F109_Gate5, LM_FL_F109_Layer1, LM_FL_F109_Parts, LM_FL_F109_Playfield, LM_FL_F109_Rpost, LM_FL_F109_Sw21, LM_FL_F109_Sw41, LM_FL_F109_Sw47, LM_FL_F109_Sw54, LM_FL_F109_sw17, LM_FL_F109_sw18, LM_FL_F109_sw25, LM_FL_F109_sw26, LM_FL_F109_sw49, LM_FL_F109_sw50, LM_FL_F109_underPF, LM_FL_F116_Parts, LM_FL_F116_underPF, LM_FL_F125_Gate1, LM_FL_F125_Layer1, LM_FL_F125_Layer2, LM_FL_F125_Parts, LM_FL_F125_Playfield, LM_FL_F125_underPF, LM_FL_F126_BR1, _
' LM_FL_F126_BR2, LM_FL_F126_BR3, LM_FL_F126_Gate1, LM_FL_F126_Gate5, LM_FL_F126_LSling1, LM_FL_F126_LSling2, LM_FL_F126_Layer1, LM_FL_F126_Layer2, LM_FL_F126_Parts, LM_FL_F126_Playfield, LM_FL_F126_RSling1, LM_FL_F126_RSling2, LM_FL_F126_Rpost, LM_FL_F126_Sw21, LM_FL_F126_Sw41, LM_FL_F126_Sw42, LM_FL_F126_Sw43, LM_FL_F126_sw17, LM_FL_F126_sw18, LM_FL_F126_sw19, LM_FL_F126_sw20, LM_FL_F126_sw25, LM_FL_F126_sw26, LM_FL_F126_sw27, LM_FL_F126_sw32, LM_FL_F126_sw39, LM_FL_F126_sw49, LM_FL_F126_sw50, LM_FL_F126_underPF, LM_FL_F127_BR1, LM_FL_F127_BR2, LM_FL_F127_BR3, LM_FL_F127_Gate1, LM_FL_F127_LFlipperU, LM_FL_F127_Layer1, LM_FL_F127_Parts, LM_FL_F127_Playfield, LM_FL_F127_RFlipper, LM_FL_F127_Sw47, LM_FL_F127_sw17, LM_FL_F127_sw18, LM_FL_F127_sw25, LM_FL_F127_sw26, LM_FL_F127_sw27, LM_FL_F127_sw33, LM_FL_F127_sw34, LM_FL_F127_sw35, LM_FL_F127_sw49, LM_FL_F127_underPF, LM_FL_F128_BR1, LM_FL_F128_BR2, LM_FL_F128_BR3, LM_FL_F128_Layer1, LM_FL_F128_Parts, LM_FL_F128_Playfield, LM_FL_F128_Sw21, LM_FL_F128_Sw47, _
' LM_FL_F128_Sw54, LM_FL_F128_sw17, LM_FL_F128_sw26, LM_FL_F128_underPF, LM_FL_F129_BR1, LM_FL_F129_BR3, LM_FL_F129_Gate5, LM_FL_F129_Parts, LM_FL_F129_Playfield, LM_FL_F129_Sw42, LM_FL_F129_Sw43, LM_FL_F129_underPF, LM_FL_F130_BR1, LM_FL_F130_BR2, LM_FL_F130_BR3, LM_FL_F130_Gate5, LM_FL_F130_Parts, LM_FL_F130_Playfield, LM_FL_F130_Sw21, LM_FL_F130_Sw22, LM_FL_F130_Sw42, LM_FL_F130_Sw43, LM_FL_F130_Sw48, LM_FL_F130_sw32, LM_FL_F130_sw50, LM_FL_F130_underPF, LM_FL_F131_Parts, LM_FL_F131_Playfield, LM_FL_F131_sw25, LM_FL_F131_sw26, LM_FL_F131_underPF, LM_FL_F132_BR1, LM_FL_F132_BR2, LM_FL_F132_BR3, LM_FL_F132_Layer1, LM_FL_F132_Layer2, LM_FL_F132_Parts, LM_FL_F132_Playfield, LM_FL_F132_Sw47, LM_FL_F132_Sw48, LM_FL_F132_sw25, LM_FL_F132_sw26, LM_FL_F132_sw27, LM_FL_F132_sw33, LM_FL_F132_sw34, LM_FL_F132_sw35, LM_FL_F132_underPF, LM_FL_L44_BR1, LM_FL_L44_BR2, LM_FL_L44_BR3, LM_FL_L44_Gate5, LM_FL_L44_Layer1, LM_FL_L44_Layer2, LM_FL_L44_Parts, LM_FL_L44_Playfield, LM_FL_L44_Rpost, LM_FL_L44_Sw21, LM_FL_L44_Sw41, _
' LM_FL_L44_Sw42, LM_FL_L44_Sw43, LM_FL_L44_Sw47, LM_FL_L44_sw32, LM_FL_L44_sw49, LM_FL_L44_underPF, LM_GI_BR1, LM_GI_BR2, LM_GI_BR3, LM_GI_Gate1, LM_GI_Gate5, LM_GI_Layer1, LM_GI_Layer2, LM_GI_Parts, LM_GI_Playfield, LM_GI_Sw21, LM_GI_Sw22, LM_GI_Sw41, LM_GI_Sw42, LM_GI_Sw43, LM_GI_Sw47, LM_GI_Sw48, LM_GI_Sw54, LM_GI_sw17, LM_GI_sw18, LM_GI_sw19, LM_GI_sw20, LM_GI_sw25, LM_GI_sw26, LM_GI_sw27, LM_GI_sw32, LM_GI_sw33, LM_GI_sw34, LM_GI_sw35, LM_GI_sw39, LM_GI_sw49, LM_GI_sw50, LM_GI_underPF, LM_GIS_GI01_Gate1, LM_GIS_GI01_LFlipper, LM_GIS_GI01_LFlipperU, LM_GIS_GI01_LSling1, LM_GIS_GI01_LSling2, LM_GIS_GI01_Layer1, LM_GIS_GI01_Layer2, LM_GIS_GI01_Parts, LM_GIS_GI01_Playfield, LM_GIS_GI01_REMK, LM_GIS_GI01_RFlipper, LM_GIS_GI01_RFlipperU, LM_GIS_GI01_RSling1, LM_GIS_GI01_RSling2, LM_GIS_GI01_Rpost, LM_GIS_GI01_Sw36, LM_GIS_GI01_Sw37, LM_GIS_GI01_underPF, LM_GIS_GI06_LEMK, LM_GIS_GI06_LFlipper, LM_GIS_GI06_LFlipperU, LM_GIS_GI06_LSling1, LM_GIS_GI06_LSling2, LM_GIS_GI06_Layer1, LM_GIS_GI06_Layer2, _
' LM_GIS_GI06_Parts, LM_GIS_GI06_Playfield, LM_GIS_GI06_RFlipper, LM_GIS_GI06_RFlipperU, LM_GIS_GI06_RSling1, LM_GIS_GI06_RSling2, LM_GIS_GI06_Sw28, LM_GIS_GI06_Sw29, LM_GIS_GI06_sw39, LM_GIS_GI06_underPF, LM_GIS_GI08_LEMK, LM_GIS_GI08_LFlipper, LM_GIS_GI08_LFlipperU, LM_GIS_GI08_LSling1, LM_GIS_GI08_LSling2, LM_GIS_GI08_Layer1, LM_GIS_GI08_Layer2, LM_GIS_GI08_Parts, LM_GIS_GI08_Playfield, LM_GIS_GI08_RFlipper, LM_GIS_GI08_RFlipperU, LM_GIS_GI08_RSling1, LM_GIS_GI08_RSling2, LM_GIS_GI08_Sw28, LM_GIS_GI08_Sw29, LM_GIS_GI08_sw39, LM_GIS_GI08_underPF, LM_GIS_GI10_LFlipper, LM_GIS_GI10_LFlipperU, LM_GIS_GI10_LSling1, LM_GIS_GI10_LSling2, LM_GIS_GI10_Layer1, LM_GIS_GI10_Layer2, LM_GIS_GI10_Parts, LM_GIS_GI10_Playfield, LM_GIS_GI10_REMK, LM_GIS_GI10_RFlipper, LM_GIS_GI10_RFlipperU, LM_GIS_GI10_RSling1, LM_GIS_GI10_RSling2, LM_GIS_GI10_Rpost, LM_GIS_GI10_Sw36, LM_GIS_GI10_Sw37, LM_GIS_GI10_sw34, LM_GIS_GI10_sw35, LM_GIS_GI10_underPF, LM_GIS_GI12_Gate1, LM_GIS_GI12_LSling1, LM_GIS_GI12_LSling2, LM_GIS_GI12_Layer1, _
' LM_GIS_GI12_Layer2, LM_GIS_GI12_Parts, LM_GIS_GI12_Playfield, LM_GIS_GI12_REMK, LM_GIS_GI12_RFlipper, LM_GIS_GI12_RFlipperU, LM_GIS_GI12_RSling1, LM_GIS_GI12_RSling2, LM_GIS_GI12_Rpost, LM_GIS_GI12_Sw36, LM_GIS_GI12_Sw37, LM_GIS_GI12_sw33, LM_GIS_GI12_sw34, LM_GIS_GI12_sw35, LM_GIS_GI12_underPF, LM_GIS_GI14_LEMK, LM_GIS_GI14_LFlipper, LM_GIS_GI14_LFlipperU, LM_GIS_GI14_LSling1, LM_GIS_GI14_LSling2, LM_GIS_GI14_Layer1, LM_GIS_GI14_Layer2, LM_GIS_GI14_Parts, LM_GIS_GI14_Playfield, LM_GIS_GI14_RSling1, LM_GIS_GI14_RSling2, LM_GIS_GI14_Sw28, LM_GIS_GI14_Sw29, LM_GIS_GI14_sw39, LM_GIS_GI14_underPF, LM_GIS_GI17_BR1, LM_GIS_GI17_BR2, LM_GIS_GI17_BR3, LM_GIS_GI17_Layer2, LM_GIS_GI17_Parts, LM_GIS_GI17_Playfield, LM_GIS_GI17_sw20, LM_GIS_GI17_underPF, LM_GIS_GI18_BR2, LM_GIS_GI18_Gate1, LM_GIS_GI18_Layer1, LM_GIS_GI18_Layer2, LM_GIS_GI18_Parts, LM_GIS_GI18_Playfield, LM_GIS_GI18_sw25, LM_GIS_GI18_sw26, LM_GIS_GI18_sw27, LM_GIS_GI18_sw33, LM_GIS_GI18_sw34, LM_GIS_GI18_sw35, LM_GIS_GI18_underPF, LM_GIS_GI19_BR2, _
' LM_GIS_GI19_BR3, LM_GIS_GI19_Layer2, LM_GIS_GI19_Parts, LM_GIS_GI19_Playfield, LM_GIS_GI19_sw25, LM_GIS_GI19_sw26, LM_GIS_GI19_sw27, LM_GIS_GI19_underPF, LM_GIS_GI20_BR2, LM_GIS_GI20_Layer1, LM_GIS_GI20_Layer2, LM_GIS_GI20_Parts, LM_GIS_GI20_Playfield, LM_GIS_GI20_sw25, LM_GIS_GI20_sw26, LM_GIS_GI20_sw27, LM_GIS_GI20_sw33, LM_GIS_GI20_sw34, LM_GIS_GI20_sw35, LM_GIS_GI20_underPF, LM_GIS_GI22_BR1, LM_GIS_GI22_BR2, LM_GIS_GI22_Layer1, LM_GIS_GI22_Layer2, LM_GIS_GI22_Parts, LM_GIS_GI22_Playfield, LM_GIS_GI22_Sw48, LM_GIS_GI22_sw27, LM_GIS_GI22_sw33, LM_GIS_GI22_sw34, LM_GIS_GI22_sw35, LM_GIS_GI22_underPF, LM_GIS_GI26_BR2, LM_GIS_GI26_BR3, LM_GIS_GI26_Parts, LM_GIS_GI26_Playfield, LM_GIS_GI26_Sw47, LM_GIS_GI26_Sw48, LM_GIS_GI26_sw17, LM_GIS_GI26_sw18, LM_GIS_GI26_sw19, LM_GIS_GI26_sw20, LM_GIS_GI26_sw25, LM_GIS_GI26_sw26, LM_GIS_GI26_sw27, LM_GIS_GI26_sw33, LM_GIS_GI26_sw34, LM_GIS_GI26_sw35, LM_GIS_GI26_sw49, LM_GIS_GI26_underPF, LM_GIS_GI30_Layer2, LM_GIS_GI30_Parts, LM_GIS_GI30_Playfield, LM_GIS_GI30_sw39, _
' LM_GIS_GI30_underPF, LM_GIS_GI32_Layer2, LM_GIS_GI32_Parts, LM_GIS_GI32_Playfield, LM_GIS_GI32_sw39, LM_GIS_GI32_underPF, LM_IN_L1_Parts, LM_IN_L1_Playfield, LM_IN_L1_RFlipperU, LM_IN_L1_underPF, LM_IN_L16_Parts, LM_IN_L16_underPF, LM_IN_L17_LFlipper, LM_IN_L17_LFlipperU, LM_IN_L17_Parts, LM_IN_L17_RFlipper, LM_IN_L17_RFlipperU, LM_IN_L17_underPF, LM_IN_L18_Layer1, LM_IN_L18_Layer2, LM_IN_L18_Parts, LM_IN_L18_Playfield, LM_IN_L18_Rpost, LM_IN_L18_underPF, LM_IN_L19_Parts, LM_IN_L19_Playfield, LM_IN_L19_sw25, LM_IN_L19_sw26, LM_IN_L19_sw27, LM_IN_L19_underPF, LM_IN_L2_Parts, LM_IN_L2_Playfield, LM_IN_L2_underPF, LM_IN_L21_Parts, LM_IN_L21_Playfield, LM_IN_L21_sw26, LM_IN_L21_sw27, LM_IN_L21_underPF, LM_IN_L22_Layer2, LM_IN_L22_Parts, LM_IN_L22_Playfield, LM_IN_L22_sw33, LM_IN_L22_sw34, LM_IN_L22_sw35, LM_IN_L22_underPF, LM_IN_L24_Layer1, LM_IN_L24_Parts, LM_IN_L24_Playfield, LM_IN_L24_sw33, LM_IN_L24_sw34, LM_IN_L24_sw35, LM_IN_L24_underPF, LM_IN_L27_Layer1, LM_IN_L27_Layer2, LM_IN_L27_Parts, _
' LM_IN_L27_Playfield, LM_IN_L27_underPF, LM_IN_L3_Parts, LM_IN_L3_Playfield, LM_IN_L3_underPF, LM_IN_L34_BR1, LM_IN_L34_BR2, LM_IN_L34_BR3, LM_IN_L34_Parts, LM_IN_L34_Playfield, LM_IN_L34_underPF, LM_IN_L4_Playfield, LM_IN_L4_underPF, LM_IN_L41_Layer1, LM_IN_L41_Parts, LM_IN_L41_Playfield, LM_IN_L41_Sw47, LM_IN_L41_underPF, LM_IN_L42_BR1, LM_IN_L42_Layer1, LM_IN_L42_Parts, LM_IN_L42_Playfield, LM_IN_L42_Sw21, LM_IN_L42_Sw47, LM_IN_L42_sw49, LM_IN_L42_underPF, LM_IN_L43_BR1, LM_IN_L43_BR2, LM_IN_L43_Layer1, LM_IN_L43_Parts, LM_IN_L43_Playfield, LM_IN_L43_Sw21, LM_IN_L43_Sw47, LM_IN_L43_sw17, LM_IN_L43_sw18, LM_IN_L43_sw49, LM_IN_L43_underPF, LM_IN_L45_Layer1, LM_IN_L45_Parts, LM_IN_L45_Playfield, LM_IN_L45_Sw47, LM_IN_L45_sw17, LM_IN_L45_sw18, LM_IN_L45_sw19, LM_IN_L45_sw20, LM_IN_L45_underPF, LM_IN_L46_Layer1, LM_IN_L46_Parts, LM_IN_L46_Playfield, LM_IN_L46_sw17, LM_IN_L46_sw18, LM_IN_L46_sw19, LM_IN_L46_sw20, LM_IN_L46_underPF, LM_IN_L47_Layer1, LM_IN_L47_Parts, LM_IN_L47_Playfield, LM_IN_L47_sw18, _
' LM_IN_L47_sw19, LM_IN_L47_sw20, LM_IN_L47_underPF, LM_IN_L48_Layer1, LM_IN_L48_Layer2, LM_IN_L48_Parts, LM_IN_L48_Playfield, LM_IN_L48_sw19, LM_IN_L48_sw20, LM_IN_L48_underPF, LM_IN_L49_LFlipper, LM_IN_L49_LFlipperU, LM_IN_L49_Layer1, LM_IN_L49_Layer2, LM_IN_L49_Parts, LM_IN_L49_Playfield, LM_IN_L49_underPF, LM_IN_L5_Playfield, LM_IN_L5_underPF, LM_IN_L50_LFlipper, LM_IN_L50_LFlipperU, LM_IN_L50_Layer2, LM_IN_L50_Parts, LM_IN_L50_underPF, LM_IN_L6_Playfield, LM_IN_L6_underPF, LM_IN_L60_BR1, LM_IN_L60_BR2, LM_IN_L60_BR3, LM_IN_L60_Parts, LM_IN_L60_Playfield, LM_IN_L60_underPF, LM_IN_L61_BR1, LM_IN_L61_BR3, LM_IN_L61_Gate5, LM_IN_L61_Layer1, LM_IN_L61_Parts, LM_IN_L61_Playfield, LM_IN_L61_Sw48, LM_IN_L7_Parts, LM_IN_L7_Playfield, LM_IN_L7_Sw14, LM_IN_L7_underPF, LM_IN_L8_underPF, LM_IN_L9_Layer1, LM_IN_L9_Layer2, LM_IN_L9_Parts, LM_IN_L9_sw39, LM_IN_L9_underPF, LM_IN_l10_Layer1, LM_IN_l10_Layer2, LM_IN_l10_Parts, LM_IN_l10_sw39, LM_IN_l10_underPF, LM_IN_l11_Parts, LM_IN_l11_underPF, LM_IN_l12_underPF, _
' LM_IN_l13_underPF, LM_IN_l14_underPF, LM_IN_l15_underPF, LM_IN_l20_Parts, LM_IN_l20_Playfield, LM_IN_l20_sw25, LM_IN_l20_sw26, LM_IN_l20_sw27, LM_IN_l20_underPF, LM_IN_l23_Layer1, LM_IN_l23_Parts, LM_IN_l23_Playfield, LM_IN_l23_sw33, LM_IN_l23_sw34, LM_IN_l23_sw35, LM_IN_l23_underPF, LM_IN_l25_Parts, LM_IN_l25_Playfield, LM_IN_l25_sw39, LM_IN_l25_underPF, LM_IN_l26_Parts, LM_IN_l26_underPF, LM_IN_l28_Parts, LM_IN_l28_RSling1, LM_IN_l28_RSling2, LM_IN_l28_underPF, LM_IN_l29_Parts, LM_IN_l29_RSling1, LM_IN_l29_underPF, LM_IN_l30_Parts, LM_IN_l30_underPF, LM_IN_l31_underPF, LM_IN_l32_underPF, LM_IN_l35_Parts, LM_IN_l35_underPF, LM_IN_l36_Parts, LM_IN_l36_underPF, LM_IN_l37_Parts, LM_IN_l37_underPF, LM_IN_l38_Parts, LM_IN_l38_sw39, LM_IN_l38_underPF, LM_IN_l39_Parts, LM_IN_l39_sw39, LM_IN_l39_underPF, LM_IN_l40_Layer2, LM_IN_l40_Parts, LM_IN_l40_sw20, LM_IN_l40_underPF, LM_IN_l51_LFlipperU, LM_IN_l51_Parts, LM_IN_l51_RFlipperU, LM_IN_l51_underPF, LM_IN_l52_Layer2, LM_IN_l52_Parts, LM_IN_l52_RFlipper, _
' LM_IN_l52_RFlipperU, LM_IN_l52_underPF, LM_IN_l53_Layer1, LM_IN_l53_Layer2, LM_IN_l53_Parts, LM_IN_l53_RFlipper, LM_IN_l53_RFlipperU, LM_IN_l53_underPF, LM_IN_l54_Parts, LM_IN_l54_underPF, LM_IN_l55_BR1, LM_IN_l55_Parts, LM_IN_l55_underPF, LM_IN_l56_BR1, LM_IN_l56_Parts, LM_IN_l56_Playfield, LM_IN_l56_underPF, LM_IN_l58_Layer2, LM_IN_l58_Parts, LM_IN_l58_underPF, LM_IN_l59_underPF, LM_IN_l62_Parts, LM_IN_l62_sw27, LM_IN_l62_underPF, LM_IN_l63_Parts, LM_IN_l63_sw27, LM_IN_l63_underPF, LM_IN_l64_Parts, LM_IN_l64_underPF)
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
  BSUpdate
  UpdateBallBrightness
  RollingUpdate       'Update rolling sounds
  DoSTAnim          'Standup target animations
  UpdateStandupTargets
  DoDTAnim          'Drop target animations
  UpdateDropTargets
  'AnimateBumperSkirts
End Sub

'The CorTimer interval should be 10. It's sole purpose is to update the Cor (physics) calculations
CorTimer.Interval = 10
Sub CorTimer_Timer(): Cor.Update: End Sub


'******************************************************
'  ZINI: Table Initialization and Exiting
'******************************************************

Sub Table1_Init

  vpmInit Me
  On Error Resume Next
    With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine = "Lethal Weapon 3 Data East 1992"&chr(13)&"VPW"
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

  vpmMapLights AllLamps

    vpmNudge.TiltSwitch = 1
  vpmNudge.Sensitivity = 2
    vpmNudge.tiltobj = Array(LeftSlingShot,RightSlingShot,Bumper1,Bumper2,Bumper3)

  'Trough
  Set LWBall1 = sw13.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set LWBall2 = sw12.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set LWBall3 = sw11.CreateSizedballWithMass(Ballsize/2,Ballmass)
  gBOT = Array(LWBall1,LWBall2,LWBall3)
  Controller.Switch(11) = 1
  Controller.Switch(12) = 1
  Controller.Switch(13) = 1

  'Auto plunger
  Set plungerIM = New cvpmImpulseP
  With plungerIM
    .InitImpulseP swplunger, 43, 0.6
    .Random 0.3
    .switch 14
    .InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("",DOFContactors)
    .CreateEvents "plungerIM"
  End With

    Set bsLock = new cvpmBallStack
  bsLock.InitSaucer sw40, 40, 165, 17
  'bsLock.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
  bsLock.KickForceVar = 5
  bsLock.KickAngleVar=5

    Set bsLock2 = new cvpmBallStack
  bsLock2.InitSaucer sw32, 32, 280, 17
  'bsLock2.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
  bsLock2.KickForceVar = 5
  bsLock2.KickAngleVar=5

  'Initialize kickback
  kickback.pullback

  'Initialize slings
  RStep = 0:RightSlingShot.Timerenabled=True
  LStep = 0:LeftSlingShot.Timerenabled=True

  'Initialize VR beacon
  SolRotateBeacons False

End Sub

Sub Table_Paused: Controller.Pause = 1: End Sub
Sub Table_unPaused: Controller.Pause = 0: End Sub
Sub Table_exit
  Controller.Pause = False
  Controller.Stop
End Sub


'*******************************************
'  ZOPT: User Options
'*******************************************

Dim MirrorMod : MirrorMod = 1           ' 0 = No Mirror, 1 = Mirror
Dim OutpostMod : OutpostMod = 1               ' 0 - Easy, 1 - Hard
Dim VRRoomChoice : VRRoomChoice = 1             ' 1 - 360 Room, 2 - Minimal Room, 3 - Ultra Minimal
Dim VRFlashingBackglass : VRFlashingBackglass = 1   ' 0 - Static VR Backglass, 1 - Flashing VR Backglass
Dim LightLevel : LightLevel = 0.25          ' Level of room lighting (0 to 1), where 0 is dark and 100 is brightest
Dim ColorLUT : ColorLUT = 1             ' Color desaturation LUTs: 1 to 11, where 1 is normal and 11 is black'n'white
Dim VolumeDial : VolumeDial = 0.8               ' Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Dim BallRollVolume : BallRollVolume = 0.5       ' Level of ball rolling volume. Value between 0 and 1
Dim RampRollVolume : RampRollVolume = 0.5       ' Level of ramp rolling volume. Value between 0 and 1


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

  Dim BP, LM

   ' Mirror Backpanel Option
    MirrorMod = Table1.Option("Mirrored Backpanel", 0, 1, 1, 1, 0, Array("Disabled", "Enabled"))
  If MirrorMod = 1 Then
    For each LM in MirrorLMs: LM.ReflectionEnabled = True: Next
    BM_Mirror1.ReflectionProbe = "Mirror1"
    BM_Mirror2.ReflectionProbe = "Mirror2"
    BM_Mirror1.visible = true
    BM_Mirror2.visible = true
  Else
    For each LM in MirrorLMs: LM.ReflectionEnabled = False: Next
    BM_Mirror1.ReflectionProbe = ""
    BM_Mirror2.ReflectionProbe = ""
    BM_Mirror1.visible = false
    BM_Mirror2.visible = false
  End If

    ' Outpost Difficulty
    OutpostMod = Table1.Option("Outpost Difficulty", 0, 2, 1, 1, 0, Array("Easy", "Normal", "Hard"))
  If OutpostMod = 0 Then
    RpostEasy.collidable = True: RpostMed.collidable = False: RpostHard.collidable = False
    For each BP in BP_Rpost: BP.TransX = 16: Next
  ElseIf OutpostMod = 1 Then
    RpostEasy.collidable = False: RpostMed.collidable = True: RpostHard.collidable = False
    For each BP in BP_Rpost: BP.TransX = 8: Next
  Else
    RpostEasy.collidable = False: RpostMed.collidable = False: RpostHard.collidable = True
    For each BP in BP_Rpost: BP.TransX = 0: Next
  End If

    ' VR Room
    VRRoomChoice = Table1.Option("VR Room", 1, 3, 1, 1, 0, Array("360 Room", "Minimal Room", "Ultra Minimal"))
  If RenderingMode = 2 Then VRRoom = VRRoomChoice Else VRRoom = 0
  'VRRoom = VRRoomChoice  'uncomment to test VR Room

  ' VR Backglass
  VRFlashingBackglass = Table1.Option("VR Backglass", 0, 1, 1, 1, 0, Array("Static", "Flashing"))
  SetupVRRoom

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

    ' Sound volumes
    VolumeDial = Table1.Option("Mech Volume", 0, 1, 0.01, 0.8, 1)
    BallRollVolume = Table1.Option("Ball Roll Volume", 0, 1, 0.01, 0.5, 1)
  RampRollVolume = Table1.Option("Ramp Roll Volume", 0, 1, 0.01, 0.5, 1)

  ' Room brightness
' LightLevel = Table1.Option("Table Brightness (Ambient Light Level)", 0, 1, 0.01, .5, 1)
  LightLevel = NightDay/100
  SetRoomBrightness LightLevel   'Uncomment this line for lightmapped tables.

    If eventId = 3 And dspTriggered Then dspTriggered = False : DisableStaticPreRendering = False : End If
End Sub


'
'Sub PrintLMRefEn  'for debugging only
' Dim LM, count: count=0
' For each LM in MirrorLMs
'   If LM.ReflectionEnabled Then
'     debug.print LM.Name
'     count=count+1
'   End If
' Next
' debug.print count
'End Sub



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

' Flipper Animations

Sub LeftFlipper_Animate
  Dim a : a = LeftFlipper.CurrentAngle
  FlipperLSh.RotZ = a

  Dim v, BP
  v = 255.0 * (122.0 -  LeftFlipper.CurrentAngle) / (122.0 -  70.0)

  For each BP in BP_LFlipper
    BP.Rotz = a
    BP.visible = v < 128.0
  Next
  For each BP in BP_LFlipperU
    BP.Rotz = a
    BP.visible = v >= 128.0
  Next
End Sub


Sub RightFlipper_Animate
  Dim a : a = RightFlipper.CurrentAngle
  FlipperRSh.RotZ = a

  Dim v, BP
  v = 255.0 * (-122.0 -  RightFlipper.CurrentAngle) / (-122.0 +  70.0)

  For each BP in BP_RFlipper
    BP.Rotz = a
    BP.visible = v < 128.0
  Next
  For each BP in BP_RFlipperU
    BP.Rotz = a
    BP.visible = v >= 128.0
  Next
End Sub


' Switch Animations

Sub sw14_Animate
  Dim z : z = sw14.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw14 : BP.transz = z: Next
End Sub

Sub sw21_Animate
  Dim z : z = sw21.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw21 : BP.transz = z: Next
End Sub

Sub sw22_Animate
  Dim z : z = sw22.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw22 : BP.transz = z: Next
End Sub

Sub sw28_Animate
  Dim z : z = sw28.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw28 : BP.transz = z: Next
End Sub

Sub sw29_Animate
  Dim z : z = sw29.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw29 : BP.transz = z: Next
End Sub

Sub sw36_Animate
  Dim z : z = sw36.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw36 : BP.transz = z: Next
End Sub

Sub sw37_Animate
  Dim z : z = sw37.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw37 : BP.transz = z: Next
End Sub

Sub sw41_Animate
  Dim z : z = sw41.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw41 : BP.transz = z: Next
End Sub

Sub sw42_Animate
  Dim z : z = sw42.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw42 : BP.transz = z: Next
End Sub

Sub sw43_Animate
  Dim z : z = sw43.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw43 : BP.transz = z: Next
End Sub

Sub sw54_Animate
  Dim z : z = sw54.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw54 : BP.transz = z: Next
End Sub

Sub sw55_Animate
  Dim z : z = sw55.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw55 : BP.transz = z: Next
End Sub


' Gate Animations

Sub sw49_Animate
    Dim a : a = sw49.CurrentAngle
    Dim BP : For Each BP in BP_sw49 : BP.rotx = -a: Next
End Sub

Sub sw50_Animate
  Dim z : z = sw50.CurrentAngle
  Dim BP : For Each BP in BP_sw50
    If z <= 0 then BP.objrotz = z
  Next
End Sub

Sub Gate1_Animate
    Dim a : a = Gate1.CurrentAngle
    Dim BP : For Each BP in BP_Gate1 : BP.rotx = a: Next
End Sub

Sub Gate5_Animate
    Dim a : a = Gate5.CurrentAngle
    Dim BP : For Each BP in BP_Gate5 : BP.rotx = a: Next
End Sub



' Spinner Animations

Sub sw47_Animate
  Dim spinangle:spinangle = sw47.currentangle
  Dim BP : For Each BP in BP_sw47 : BP.RotX = spinangle: Next
End Sub

Sub sw48_Animate
  Dim spinangle:spinangle = sw48.currentangle
  Dim BP : For Each BP in BP_sw48 : BP.RotX = spinangle: Next
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

'Dim Bumpers : Bumpers = Array(Bumper1, Bumper2, Bumper3)
'
'Sub AnimateBumperSkirts
' dim r, g, s, x, y, b, z
' ' Animate Bumper switch (experimental)
' For r = 0 To 2
'   g = 10000.
'   For s = 0 to UBound(gBOT)
'     If gBOT(s).z < 30 Then ' Ball on playfield
'       x = Bumpers(r).x - gBOT(s).x
'       y = Bumpers(r).y - gBOT(s).y
'       b = x * x + y * y
'       If b < g Then g = b
'     End If
'   Next
'   z = 4
'   If g < 80 * 80 Then
'     z = 1
'   End If
'   If r = 0 Then For Each x in BP_BS1: x.Z = z: Next
'   If r = 1 Then For Each x in BP_BS2: x.Z = z: Next
'   If r = 2 Then For Each x in BP_BS3: x.Z = z: Next
' Next
'End Sub



'******************************************************
'   ZBBR: BALL BRIGHTNESS
'******************************************************

Const BallBrightness =  1       'Ball brightness - Value between 0 and 1 (0=Dark ... 1=Bright)

' Constants for plunger lane ball darkening.
' You can make a temporary wall in the plunger lane area and use the co-ordinates from the corner control points.
Const PLOffset = 0.5      'Minimum ball brightness scale in plunger lane
Const PLLeft = 860        'X position of punger lane left
Const PLRight = 930       'X position of punger lane right
Const PLTop = 1420        'Y position of punger lane top
Const PLBottom = 1940       'Y position of punger lane bottom
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
' ZKEY: Key Press Handling
'******************************************************


Sub Table1_KeyDown(ByVal KeyCode)
  If keycode = LeftFlipperKey Then
    FlipperActivate LeftFlipper, LFPress
    PinCab_LeftFlipperButton.X = PinCab_LeftFlipperButton.X + 10
  End If
  If keycode = RightFlipperKey Then
    FlipperActivate RightFlipper, RFPress
    PinCab_RightFlipperButton.X = PinCab_RightFlipperButton.X - 10
  End If
  If keycode = LeftFlipperKey Then Controller.Switch(15) = 1
  If keycode = RightFlipperKey Then Controller.Switch(16) = 1
  If keycode = PlungerKey Then Controller.Switch(9) = 1 : SoundPlungerPull()
  If keycode = LockBarKey Then Controller.Switch(9) = 1  'Lockdown Bar button support
  if keycode = StartGameKey then
    SoundStartButton
    VRLaunchButton.Y = VRLaunchButton.Y - 4
    Flasher1.Y = Flasher1.Y - 4
  End if
    If keycode = LeftTiltKey Then Nudge 90, 1: SoundNudgeLeft
    If keycode = RightTiltKey Then Nudge 270, 1: SoundNudgeRight
    If keycode = CenterTiltKey Then Nudge 0, 1: SoundNudgeCenter
    If keycode = keyInsertCoin1 or keycode = keyInsertCoin2 or keycode = keyInsertCoin3 or keycode = keyInsertCoin4 Then
        Select Case Int(rnd*3)
            Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
            Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
            Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
        End Select
    End If
  If vpmKeyDown(keycode) Then Exit Sub
End Sub


Sub Table1_KeyUp(ByVal KeyCode)
  If keycode = LeftFlipperKey Then
    FlipperDeActivate LeftFlipper, LFPress
    PinCab_LeftFlipperButton.X = PinCab_LeftFlipperButton.X - 10
  End If
  If keycode = RightFlipperKey Then
    FlipperDeActivate RightFlipper, RFPress
    PinCab_RightFlipperButton.X = PinCab_RightFlipperButton.X + 10
  End If
  If keycode = LeftFlipperKey Then Controller.Switch(15) = 0
  If keycode = RightFlipperKey Then Controller.Switch(16) = 0
  If keycode = PlungerKey Then Controller.Switch(9) = 0
  If keycode = LockBarKey Then Controller.Switch(9) = 0
  if keycode = StartGameKey then
    SoundStartButton
    VRLaunchButton.Y = VRLaunchButton.Y + 4
    Flasher1.Y = Flasher1.Y + 4
  End if
  If vpmKeyUp(keycode) Then Exit Sub
End Sub


'******************************************************
' ZSOL: Solenoids & Flashers
'******************************************************


SolCallback(1)= "SolTrough"
SolCallback(2)= "SolRelease"
'3 Not Used
SolCallback(4)= "bsLock.SolOut"
SolCallback(5)= "bsLock2.SolOut"
SolCallback(6)= "dtMSolDropUp"
SolCallback(7)= "dtRSolDropUp"
SolCallback(8)= "SolKnocker"
SolModCallback(9)= "Sol09"      'IWSC Buding  X4 Lights
'10 PPB board not needed in PINMAME
SolModCallback(11)= "SolRelayGI" 'GI Relay
SolCallback(12)= "Autofire"
'13 Not Used
SolCallback(14)= "SolRotateBeacons"   'Mars Light aka Beacon
SolCallback(15)= "VukTopPop"
SolModCallback(16)= "Sol16"         'X4 Lights            'BG - Top Right Copter
'SolCallback(17)="vpmSolSound ""jet1"","
'SolCallback(18)="vpmSolSound ""jet1"","
'SolCallback(19)="vpmSolSound ""jet1"","
'SolCallback(20)="vpmSolSound ""rsling"","
'SolCallback(21)="vpmSolSound ""rsling"","
SolCallback(22)="SolKickback"
'23-24 Not Used
SolModCallback(25)="Flash1R"        '1R PF          'BG - Mel Top - Bottom Right Building Left
SolModCallback(26)="Flash2R"        '2R Fluppers Dome
SolModCallback(27)="Flash3R"        '3R Fluppers Dome, PF 'BG - Bottom Left Rails
SolModCallback(28)="Flash4R"        '4R Fluppers Dome
SolModCallback(29)="Flash5R"        '5R Fluppers Dome, PF 'BG - Mel Bottom - Bottom Right Building Bottom
SolModCallback(30)="Flash6R"        '6R Fluppers Dome, PF
SolModCallback(31)="Flash7R"        '7R PF          'BG - Left 1 - Bottom Right Building Top
SolModCallback(32)="Flash8R"        '8R PF          'BG - Left 2
SolModCallback(32)="Flash8R"        '8R Fluppers Dome

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"


' Some Solenoids
'*******************

Sub Autofire(enabled)
    If enabled Then
        PlungerIM.AutoFire
    SoundPlungerReleaseBall()
    End If
End Sub


Sub SolKickback(Enabled)
  If Enabled Then
    SoundSaucerKick 1, kickback
    kickback.fire
  Else
    kickback.pullback
  End If
End Sub


Sub SolKnocker(Enabled)
    If enabled Then
        KnockerSolenoid 'Add knocker position object
    End If
End Sub



'******************************************************
' ZFLA: Flashers
'******************************************************

Sub Sol09(level)
  F109.state = level
  F109a.state = level
End Sub

Sub Sol16(level)
  F116.state = level
End Sub

Sub Flash1R(level)
  F125.state = level
End Sub

Sub Flash2R(level) '2R
  F126.state = level
  F126a.state = level
End Sub

Sub Flash3R(level) '3R
  F127.state = level
  F127a.state = level
End Sub

Sub Flash4R(level) '4R
  F128.state = level
  F128a.state = level
  F128b.state = level
End Sub

Sub Flash5R(level) '5R
  F129.state = level
  F129a.state = level
End Sub

Sub Flash6R(level) '6R
  F130.state = level
  F130a.state = level
  F130b.state = level
  F130c.state = level
End Sub

Sub Flash7R(level)
  F131.state = level
End Sub

Sub Flash8R(level) '8R
  F132.state = level
  F132a.state = level
 End Sub



'******************************************************
'   ZGIU: GI updates
'******************************************************

dim gilvl: gilvl=0

'Playfield GI
Sub SolRelayGI(level)
' debug.print "SolRelayGI value: " & level
  dim bulb: For each bulb in GI: bulb.State = level: Next

  ' Handle spotlight bulbs
  pSpotBulbs.opacity = level
  pSpotBulbs.blenddisablelighting = 50*level

  ' Relay sound
  If level >= 0.5 And gilvl < 0.5 Then
    Sound_GI_Relay 1, KnockerPosition 'Note: Bumper1 is just used for sound positioning. Can be anywhere that makes sense.
  ElseIf level <= 0.4 And gilvl > 0.4 Then
    Sound_GI_Relay 0, KnockerPosition
  End If
End Sub




'******************************************************
' ZFLP: FLIPPERS
'******************************************************

Const ReflipAngle = 20

Sub SolLFlipper(Enabled)
     If Enabled Then
      LF.fire 'LeftFlipper.RotateToEnd
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
      RF.fire 'RightFlipper.RotateToEnd
        If RightFlipper.currentangle > RightFlipper.endangle - ReflipAngle Then
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




'******************************************************
' ZVUK: VUKs and Kickers
'******************************************************


 ' kickers
Sub sw32_Hit: bsLock2.AddBall 0 : SoundSaucerLock: End Sub
Sub sw32_UnHit: SoundSaucerKick 1, sw32: End Sub
Sub sw40_Hit: bsLock.AddBall 0 : SoundSaucerLock: End Sub
Sub sw40_UnHit: SoundSaucerKick 1, sw40: End Sub


'Top Raising VUK sw31
Dim RaiseBall

Sub sw31_Hit
  Set RaiseBall = ActiveBall
  Controller.switch(31) = 1
  SoundSaucerLock
End Sub

Sub VukTopPop(Enabled)
  If (Enabled and Not IsEmpty(RaiseBall)) Then
    SoundSaucerKick 1, sw31
    sw31.KickZ 0, 30, 90, 50
    Controller.switch (31) = 0
  End If
End Sub



'******************************************************
' ZDRN: Drain, Trough, and Ball Release
'******************************************************

' Trough

Sub sw13_Hit   : Controller.Switch(13) = 1 : UpdateTrough : End Sub
Sub sw13_UnHit : Controller.Switch(13) = 0 : UpdateTrough : End Sub
Sub sw12_Hit   : Controller.Switch(12) = 1 : UpdateTrough : End Sub
Sub sw12_UnHit : Controller.Switch(12) = 0 : UpdateTrough : End Sub
Sub sw11_Hit   : Controller.Switch(11) = 1 : UpdateTrough : End Sub
Sub sw11_UnHit : Controller.Switch(11) = 0 : UpdateTrough : End Sub
Sub sw10_Hit   : Controller.Switch(10) = 1 : UpdateTrough : RandomSoundDrain sw10 :End Sub
Sub sw10_UnHit : Controller.Switch(10) = 0 : UpdateTrough : End Sub


Sub UpdateTrough
  UpdateTroughTimer.Interval = 200
  UpdateTroughTimer.Enabled = 1
End Sub

Sub UpdateTroughTimer_Timer
  If sw13.BallCntOver = 0 Then sw12.kick 60, 10
  If sw12.BallCntOver = 0 Then sw11.kick 60, 10
  If sw11.BallCntOver = 0 Then sw10.kick 60, 15
  UpdateTroughTimer.Enabled = 0
End Sub


' Drain and Release

Sub SolTrough(enabled)
  If enabled Then
    sw10.kick 60, 30
  End If
End Sub

Sub SolRelease(enabled)
  If enabled Then
    sw13.kick 60, 10
    RandomSoundBallRelease sw13
  End If
End Sub




'************************************************************
' ZSLG: Slingshot Animations
'************************************************************


Dim RStep, Lstep

Sub RightSlingShot_Slingshot
  RS.VelocityCorrect(Activeball)
  vpmTimer.PulseSw 38
    RandomSoundSlingshotRight Sw37
  RStep = 0
  RightSlingShot.TimerInterval = 17
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
  Dim BL
  Dim x1, x2, y: x1 = True:x2 = False:y = 25
    Select Case RStep
        Case 2:x1 = False:x2 = True:y = 15
        Case 3:x1 = False:x2 = False:y = 0:RightSlingShot.TimerEnabled = 0
    End Select
  For Each BL in BP_RSling1 : BL.Visible = x1: Next
  For Each BL in BP_RSling2 : BL.Visible = x2: Next
  For Each BL in BP_REMK : BL.transx = y: Next
    RStep = RStep + 1
End Sub


Sub LeftSlingShot_Slingshot
  LS.VelocityCorrect(Activeball)
  vpmTimer.PulseSw 30
    RandomSoundSlingshotLeft Sw29
  LStep = 0
  LeftSlingShot.TimerInterval = 17
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
  Dim BL
  Dim x1, x2, y: x1 = True:x2 = False:y = 25
  Select Case LStep
    Case 3:x1 = False:x2= True: y = 15
    Case 4:x1 = False:x2 = False:y = 0:LeftSlingShot.TimerEnabled = 0
  End Select
  For Each BL in BP_LSling1 : BL.Visible = x1: Next
  For Each BL in BP_LSling2 : BL.Visible = x2: Next
  For Each BL in BP_LEMK : BL.transx = y: Next
  LStep = LStep + 1
End Sub




'******************************************************
' ZSSC: SLINGSHOT CORRECTION FUNCTIONS by apophis
'******************************************************

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


'Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 44 : RandomSoundBumperTop Bumper1 : End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 45 : RandomSoundBumperMiddle Bumper2 : End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 46 : RandomSoundBumperBottom Bumper3 : End Sub

'Spinner
Sub Sw47_Spin:vpmTimer.PulseSw 47:End sub
Sub Sw48_Spin:vpmTimer.PulseSw 48:End sub

'Center Ramp Gates
Sub Sw49_Hit:vpmTimer.PulseSw 49: End Sub
Sub Sw50_Hit:vpmTimer.PulseSw 50: End sub

'10 Point
Sub Sw52_Hit:Controller.Switch(52)=1: End Sub
Sub Sw52_UnHit:Controller.Switch(52)=0: End Sub


'Wire Triggers
'Sub Sw14_Hit: : End Sub 'Coded to impulse plunger
'Sub Sw14_UnHit:Controller.Switch(14)=0: End Sub
Sub Sw21_Hit:Controller.Switch(21)=1: End Sub
Sub Sw21_UnHit:Controller.Switch(21)=0: End Sub
Sub Sw22_Hit:Controller.Switch(22)=1: End Sub
Sub Sw22_UnHit:Controller.Switch(22)=0: End Sub
Sub Sw28_Hit:Controller.Switch(28)=1: End Sub
Sub Sw28_UnHit:Controller.Switch(28)=0: End Sub
Sub Sw29_Hit:Controller.Switch(29)=1: leftInlaneSpeedLimit: End Sub
Sub Sw29_UnHit:Controller.Switch(29)=0: End Sub
Sub Sw36_Hit:Controller.Switch(36)=1: End Sub
Sub Sw36_UnHit:Controller.Switch(36)=0: End Sub
Sub Sw37_Hit:Controller.Switch(37)=1: rightInlaneSpeedLimit: End Sub
Sub Sw37_UnHit:Controller.Switch(37)=0: End Sub
Sub Sw41_Hit:Controller.Switch(41)=1: End Sub
Sub Sw41_UnHit:Controller.Switch(41)=0: End Sub
Sub Sw42_Hit:Controller.Switch(42)=1: End Sub
Sub Sw42_UnHit:Controller.Switch(42)=0: End Sub
Sub Sw43_Hit:Controller.Switch(43)=1: End Sub
Sub Sw43_UnHit:Controller.Switch(43)=0: End Sub
Sub Sw54_Hit:Controller.Switch(54)=1: End Sub
Sub Sw54_UnHit:Controller.Switch(54)=0: End Sub
Sub Sw55_Hit:Controller.Switch(55)=1: End Sub
Sub Sw55_UnHit:Controller.Switch(55)=0: End Sub

' Inlane switch speedlimit code

Sub leftInlaneSpeedLimit
  'Wylte's implementation
'    debug.print "Spin in: "& activeball.AngMomZ
'    debug.print "Speed in: "& activeball.vely
  if activeball.vely < 0 then exit sub              'don't affect upwards movement
    activeball.AngMomZ = -abs(activeball.AngMomZ) * RndNum(3,6)
    If abs(activeball.AngMomZ) > 60 Then activeball.AngMomZ = 0.8 * activeball.AngMomZ
    If abs(activeball.AngMomZ) > 80 Then activeball.AngMomZ = 0.8 * activeball.AngMomZ
    If activeball.AngMomZ > 100 Then activeball.AngMomZ = RndNum(80,100)
    If activeball.AngMomZ < -100 Then activeball.AngMomZ = RndNum(-80,-100)

    if abs(activeball.vely) > 5 then activeball.vely = 0.8 * activeball.vely
    if abs(activeball.vely) > 10 then activeball.vely = 0.8 * activeball.vely
    if abs(activeball.vely) > 15 then activeball.vely = 0.8 * activeball.vely
    if activeball.vely > 16 then activeball.vely = RndNum(14,16)
    if activeball.vely < -16 then activeball.vely = RndNum(-14,-16)
'    debug.print "Spin out: "& activeball.AngMomZ
'    debug.print "Speed out: "& activeball.vely
End Sub


Sub rightInlaneSpeedLimit
  'Wylte's implementation
'    debug.print "Spin in: "& activeball.AngMomZ
'    debug.print "Speed in: "& activeball.vely
  if activeball.vely < 0 then exit sub              'don't affect upwards movement
    activeball.AngMomZ = abs(activeball.AngMomZ) * RndNum(2,4)
    If abs(activeball.AngMomZ) > 60 Then activeball.AngMomZ = 0.8 * activeball.AngMomZ
    If abs(activeball.AngMomZ) > 80 Then activeball.AngMomZ = 0.8 * activeball.AngMomZ
    If activeball.AngMomZ > 100 Then activeball.AngMomZ = RndNum(80,100)
    If activeball.AngMomZ < -100 Then activeball.AngMomZ = RndNum(-80,-100)

  if abs(activeball.vely) > 5 then activeball.vely = 0.8 * activeball.vely
    if abs(activeball.vely) > 10 then activeball.vely = 0.8 * activeball.vely
    if abs(activeball.vely) > 15 then activeball.vely = 0.8 * activeball.vely
    if activeball.vely > 16 then activeball.vely = RndNum(14,16)
    if activeball.vely < -16 then activeball.vely = RndNum(-14,-16)
'    debug.print "Spin out: "& activeball.AngMomZ
'    debug.print "Speed out: "& activeball.vely
End Sub


Sub swplunger_UnHit: SoundPlungerReleaseBall(): End Sub
Sub sw14_UnHit: SoundPlungerReleaseBall(): End Sub



'Stand Up targets
Sub Sw17_Hit: STHit 17: End Sub
Sub Sw18_Hit: STHit 18: End Sub
Sub Sw19_Hit: STHit 19: End Sub
Sub Sw20_Hit: STHit 20: End Sub
Sub Sw39_Hit: STHit 39: End Sub


' Drop Targets
Sub sw25_hit: DTHit 25: End Sub
Sub sw26_hit: DTHit 26: End Sub
Sub sw27_hit: DTHit 27: End Sub
Sub sw33_hit: DTHit 33: End Sub
Sub sw34_hit: DTHit 34: End Sub
Sub sw35_hit: DTHit 35: End Sub


Sub dtMSolDropUp(Enabled)
    If Enabled Then
        DTRaise 25
        DTRaise 26
        DTRaise 27
    RandomSoundDropTargetReset BM_sw26
        'reset drop shadows
    S25.visible=True
    S26.visible=True
    S27.visible=True
    End If
End Sub

Sub dtRSolDropUp(Enabled)
    If Enabled Then
        DTRaise 33
        DTRaise 34
        DTRaise 35
    RandomSoundDropTargetReset BM_sw34
        'reset drop shadows
    S33.visible=True
    S34.visible=True
    S35.visible=True
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
dim RampBalls(4,2)
'x,0 = ball x,1 = ID, 2 = Protection against ending early (minimum amount of updates)
'0,0 is boolean on/off, 0,1 unused for now
RampBalls(0,0) = False

' RampType
'     Setup: Set this array to the number Total number of balls that can be tracked at one time + 1.  5 ball multiball then set value to 6
'     Description: Array type indexed on BallId and a values used to deterimine what type of ramp the ball is on: False = Wire Ramp, True = Plastic Ramp
dim RampType(4)

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

Sub WRT000_UnHit
  If activeball.vely < 0 Then
    WireRampOn True
  Else
    WireRampOff
  End If
End Sub

Sub WRT001_Hit
  WireRampOn False
  activeball.angmomx = 0
  activeball.angmomy = 0
  activeball.angmomz = 0
End Sub

Sub WRT002_Hit
  WireRampOff
  'RandomSoundRampStop WRT002
End Sub

Sub WRT002_UnHit
  RandomSoundDelayedBallDropOnPlayfield ActiveBall
End Sub

Sub WRT003_Hit: WireRampOn False: End Sub

Sub WRT004_Hit
  WireRampOff
  'RandomSoundRampStop WRT004
End Sub

Sub WRT004_UnHit
  RandomSoundDelayedBallDropOnPlayfield ActiveBall
End Sub

Sub RandomSoundRampStop(obj)
  Select Case Int(rnd*3)
    Case 0: PlaySoundAtVol "wireramp_stop1", obj, 0.7*VolumeDial:PlaySoundAtLevelActiveBall ("Rubber_Strong_1"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
    Case 1: PlaySoundAtVol "wireramp_stop2", obj, 0.7*VolumeDial:PlaySoundAtLevelActiveBall ("Rubber_Strong_2"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
    Case 2: PlaySoundAtVol "wireramp_stop3", obj, 0.7*VolumeDial:PlaySoundAtLevelActiveBall ("Rubber_1_Hard"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
  End Select
End Sub



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
  PlaySoundAtLevelStatic ("Plunger_Pull_1"), PlungerPullSoundLevel, swPlunger
End Sub

Sub SoundPlungerReleaseBall()
  PlaySoundAtLevelStatic ("Plunger_Release_Ball"), PlungerReleaseSoundLevel, swPlunger
End Sub

Sub SoundPlungerReleaseNoBall()
  PlaySoundAtLevelStatic ("Plunger_Release_No_Ball"), PlungerReleaseSoundLevel, swPlunger
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
' ZNFF:  FLIPPER CORRECTIONS by nFozzy
'******************************************************

Dim LF : Set LF = New FlipperPolarity
Dim RF : Set RF = New FlipperPolarity

InitPolarity

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
'   ZBOU: VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************

Const TargetBouncerEnabled = 1    '0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.7   'Level of bounces. Recommmended value of 0.7-1.0

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
'   ZRDT:  DROP TARGETS by Rothbauerw
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
Dim DT25, DT26, DT27, DT33, DT34, DT35


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

Set DT25 = (new DropTarget)(sw25, sw25a, BM_sw25, 25, 0, false)
Set DT26 = (new DropTarget)(sw26, sw26a, BM_sw26, 26, 0, false)
Set DT27 = (new DropTarget)(sw27, sw27a, BM_sw27, 27, 0, false)
Set DT33 = (new DropTarget)(sw33, sw33a, BM_sw33, 33, 0, false)
Set DT34 = (new DropTarget)(sw34, sw34a, BM_sw34, 34, 0, false)
Set DT35 = (new DropTarget)(sw35, sw35a, BM_sw35, 35, 0, false)

Dim DTArray
DTArray = Array(DT25, DT26, DT27, DT33, DT34, DT35)

'Configure the behavior of Drop Targets.
Const DTDropSpeed = 80 'in milliseconds
Const DTDropUpSpeed = 40 'in milliseconds
Const DTDropUnits = 44 'VP units primitive drops so top of at or below the playfield
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



'  DROP TARGETS FUNCTIONS
'*************************

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

  tz = BM_sw25.transz
  rx = BM_sw25.rotx
  ry = BM_sw25.roty
  For each BP in BP_sw25: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw26.transz
  rx = BM_sw26.rotx
  ry = BM_sw26.roty
  For each BP in BP_sw26: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw27.transz
  rx = BM_sw27.rotx
  ry = BM_sw27.roty
  For each BP in BP_sw27: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw33.transz
  rx = BM_sw33.rotx
  ry = BM_sw33.roty
  For each BP in BP_sw33: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw34.transz
  rx = BM_sw34.rotx
  ry = BM_sw34.roty
  For each BP in BP_sw34: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw35.transz
  rx = BM_sw35.rotx
  ry = BM_sw35.roty
  For each BP in BP_sw35: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next
End Sub



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
Dim ST17, ST18, ST19, ST20, ST39

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

Set ST17 = (new StandupTarget)(sw17, BM_sw17, 17, 0)
Set ST18 = (new StandupTarget)(sw18, BM_sw18, 18, 0)
Set ST19 = (new StandupTarget)(sw19, BM_sw19, 19, 0)
Set ST20 = (new StandupTarget)(sw20, BM_sw20, 20, 0)
Set ST39 = (new StandupTarget)(sw39, BM_sw39, 39, 0)

'Add all the Stand-up Target Arrays to Stand-up Target Animation Array
' STAnimationArray = Array(ST1, ST2, ....)
Dim STArray
STArray = Array(ST17, ST18, ST19, ST20, ST39)

'Configure the behavior of Stand-up Targets
Const STAnimStep =  1.5     'vpunits per animation step (control return to Start)
Const STMaxOffset = 9       'max vp units target moves when hit

Const STMass = 0.2        'Mass of the Stand-up Target (between 0 and 1), higher values provide more resistance


' STAND-UP TARGETS FUNCTIONS
'*****************************

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

    ty = BM_sw17.transy
  For each BP in BP_sw17 : BP.transy = ty: Next

    ty = BM_sw18.transy
  For each BP in BP_sw18 : BP.transy = ty: Next

    ty = BM_sw19.transy
  For each BP in BP_sw19 : BP.transy = ty: Next

    ty = BM_sw20.transy
  For each BP in BP_sw20 : BP.transy = ty: Next

End Sub



'******************************************************
'   ZVRR: VR Room / VR Cabinet
'******************************************************


'VR Mode

Sub SetupVRRoom
  Dim VRThings, xx, BP

  If VRRoom > 0 Then
    scoretext.visible = 0
    DMD.visible = 1
    PinCab_Backglass.blenddisablelighting = 3
    'For each BP in BP_PinCab_Rails: BP.visible = 1: Next
    If VRRoom = 1 Then
      for each VRThings in BigRoom:VRThings.visible = 1:Next
      for each VRThings in VRCab:VRThings.visible = 1:Next
      for each VRThings in VRMin:VRThings.visible = 0:Next
    End If
    If VRRoom = 2 Then
      for each VRThings in BigRoom:VRThings.visible = 0:Next
      for each VRThings in VRCab:VRThings.visible = 1:Next
      for each VRThings in VRMin:VRThings.visible = 1:Next
    End If
    If VRRoom = 3 Then
      for each VRThings in BigRoom:VRThings.visible = 0:Next
      for each VRThings in VRCab:VRThings.visible = 0:Next
      for each VRThings in VRMin:VRThings.visible = 0:Next
      PinCab_Backbox.visible = 1
      PinCab_Backglass.visible = 1
    End If
    If VRFlashingBackglass = 1 Then
      For each xx in BGGI: xx.visible = 1: Next
      PinCab_Backglass.visible = 0
      BGDark.visible = 1
    Else
      For each xx in BGGI: xx.visible = 0: Next
    End If
  Else
    for each VRThings in BigRoom:VRThings.visible = 0:Next
    for each VRThings in VRCab:VRThings.visible = 0:Next
    for each VRThings in VRMin:VRThings.visible = 0:Next
    if DesktopMode then
      scoretext.visible = 1
      PinCab_Rails.visible = 1
'     For each BP in BP_PinCab_Rails: BP.visible = 1: Next
    else
      scoretext.visible = 0
      PinCab_Rails.visible = 0
'     For each BP in BP_PinCab_Rails: BP.visible = 0: Next
    End if
  End If


End Sub

SetBackglass
Sub SetBackglass()
  Dim obj

  For Each obj In BackglassMid
    obj.x = obj.x + 3
    obj.height = - obj.y + 335
    obj.y = -43 'adjusts the distance from the backglass towards the user
    obj.rotx=-89
  Next

  For Each obj In BackglassArea
    obj.x = obj.x + 3
    obj.height = - obj.y + 335
    obj.y = -50 'adjusts the distance from the backglass towards the user
    obj.rotx=-89
  Next

End Sub



' VR Toys

Dim BeaconPos:BeaconPos = 0

Sub BeaconTimer_Timer
  BeaconPos = BeaconPos + 3
  if BeaconPos = 360 then BeaconPos = 0
  BeaconBlueInt.RotY = BeaconPos+90
    BeaconBlue.BlendDisableLighting=.3 * abs(sin((BeaconPos+90+90) * 6.28 / 360))
  BeaconFB.RotY = BeaconPos + 90
    if BeaconPos+90 > 270 then BeaconFb.IntensityScale = -1 else BeaconFb.IntensityScale = 2
End Sub

Sub SolRotateBeacons(Enabled)
    If Enabled then
    BeaconTimer.Enabled = true
        BeaconBlue.image = "dome3_blue_lit"
    'PlaySound SoundFX("fx_relay",DOFContactors)
    Else
    BeaconTimer.Enabled = false
        BeaconBlue.image = "dome3_blue"
    if BeaconPos+90 > 270 then BeaconFb.IntensityScale = -.01 else BeaconFb.IntensityScale = .01
    End If
End Sub



'1.2.1 Wylte - Added FlipperCoilRampupMode options, lowered LiveCatch 24 > 16
'1.2.2 fluffhead - Added ability to adjust volume multiplication for a specific ball on ramps.  Set the pigtail ramp to 8 for volume multiplication.
'1.2.3 tomate - OFF plastics texture fixed, left hand ramp stopper friction fixed (thanks Sixtoe!)
'1.2.4 fluffhead - Changed WireRampOnSndX from 8 to 2 to make sounds better after more testing.
'1.2.5 apophis - Notes:
'  - Reorganized layers. Reorg scripts, add TOC
'  - Update all light faders to None.
'  - UseVPMModSol=2. UseLamps=1. Updated GI and Flasher subs.
'  - Set up insert lights to use vpmMapLights (TimerInterval of each light set to ROM light index. Added to AllLamps collection)
'  - Added Inlane switch speed limit code.
'  - Remove Lamp fading code. Removed flupper dome code.
'  - Add ball brightness and room brightness subs
'  - Update ball rolling and ramp rolling subs. Updated Fleep script.
'  - Updated flipper triggers. Update physics scripts.
'  - Removed old dynamic shadows in favor of new built-in dynamic shadows. Updated ambient shadow code.
'  - Added physical trough. Removed/updated all ball destroying scripts.
'  - Added Roth standup and drop targets.
'  - Update animations using _animate subs. Prepare for toolkit animations.
'  - Added Tweak menu, including desat LUTs
'  - Added table rules card and description with markdown formatting
'  - Update physical walls
'  - Clean up table (Delete old stuff)
'  - Clean up image manager
'1.2.6 apophis - Added LUT images.
'1.2.7 tommate & apophis - New 2k batch. (Has some issues)
'1.2.8 apophis - Set all lights to hidden. Hide physical stuff. Set up GI shadows. Updated all animations. Added Outpost Difficulty option. Aligned some physical objects with bakemaps.
'1.2.9 apophis - New 1k-ish bake. Fixed GI08 name.
'1.2.10 apophis - New 2k bake. Fixed L27 name.
'1.2.11 tomate - New physical ramps added, VUK tuned (still need work)
'1.2.12 tomate - changed lamps, GI and flashers to incandescent
'1.2.13 mcarter78 - moved ramp end triggers onto physical ramps to prevent triggering from below, randomized ramp end sounds, added delayed ball drop sound at ramp end
'1.2.14 iaakki - Mirror backpanel
'1.2.15 tomate - tons of fixes at blender side, 4k batch added, hide parts behind set for BM_playfield.
'1.2.16 apophis - Added mirror option to tweak menu (mirror not working yet). Dampener velz correction. Flippers at 3000 strength. Slope 6.3 deeg. FIxed SetBackglass issue.
'1.2.17 DGrimmReaper - VR Fixes.  Animated flipper buttons.
'1.2.18 apophis - Fixed mirror visibility (still need to add LMs). Fixed VR cab visibility in DT mode. Flippers at 2900 strength. Slingshots at 4.5.
'1.2.19 apophis - Aligned physical flippers to visual flippers. Set activeball angmom to zero at top of spiral ramp. Enabled playfield reflections.
'1.2.20 tomate - re-worked wireramps and metal walls in blender, new batch imported.
'1.2.21 DGrimmReaper - Lockdownbar and rails added.  Additional VR fixes. Anitmated start button.  Made rails visible in desktop. dome3_blue_lit was missing from images
'1.2.22 apophis - Fixed stutters. Fixed VUK vertical kick. Lowered ball in plunger lane.
'1.2.23 iaakki - Added mirror backpanel parts to MirrorLMs collection. Hid ramp410
'1.0 RC1 apophis - Fixed group GI control light assignment. Updated DisableStaticPreRendering functionality to be compatible with VPX 10.8.1 API.
'1.0 RC2 apophis - Tried to protect drain against multiball balls. Tried to fortify right wall against stuck balls. Adjusted playfield slope to 6.4 and friction to 0.2, and made slings less sensitive near posts (thanks AstroNasty). Made apron physical wall 300 tall. Flipper strength to 2800. Deleted old visible target primitives.
'1.0 RC3 apophis - Flipper strength back to 2900. Mirror on by default.
'1.0 RC4 apophis - Spotlight bulb hack added.
'1.0 RC5 apophis - Fixed black E in Data East. Added screenshot
'Release 2.0

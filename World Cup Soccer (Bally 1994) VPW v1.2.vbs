' __          ______  _____  _      _____     _____ _    _ _____
' \ \        / / __ \|  __ \| |    |  __ \   / ____| |  | |  __ \
'  \ \  /\  / / |  | | |__) | |    | |  | | | |    | |  | | |__) |
'   \ \/  \/ /| |  | |  _  /| |    | |  | | | |    | |  | |  ___/
'    \  /\  / | |__| | | \ \| |____| |__| | | |____| |__| | |
'     \/  \/___\____/|_|  \_\______|_____/ _ \_____|\____/|_|
'         / ____|                         ( ) _ \| || |
'        | (___   ___   ___ ___ ___ _ __  |/ (_) | || |_
'         \___ \ / _ \ / __/ __/ _ \ '__|   \__, |__   _|
'         ____) | (_) | (_| (_|  __/ |        / /   | |
'        |_____/ \___/ \___\___\___|_|       /_/    |_|
'
' World Cup Soccer '94 (Bally 1994)
' https://www.ipdb.org/machine.cgi?gid=2811
'
' Team VPW
' ----------
' mcarter78 - Captain
' Sixtoe -  Forward
' tomate - Striker
' apophis - Winger
' DaRdog - Midfielder
' DGrimmReaper - Defender
' RobbyKingPin - Goalkeeper
'
' Based on recreation by Knorr and Clark Kent
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
'   ZFLD: Flupper Domes
'   ZFLB: Flupper Bumpers
' ZSHA: Ambient Ball Shadows
'   ZVRS: VR Stuff
'
Option Explicit
Randomize
SetLocale 1033      'Forces VBS to use english to stop crashes.

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0


' VLM  Arrays - Start
' Arrays per baked part
Dim BP_BR1: BP_BR1=Array(BM_BR1, LM_GIRight_GI041_BR1, LM_GILeft_GI014_BR1, LM_GITop_BR1, LM_FL_f20_BR1, LM_FL_f22_BR1, LM_FL_f25_BR1, LM_IN_l47_BR1)
Dim BP_BR2: BP_BR2=Array(BM_BR2, LM_GITop_BR2, LM_FL_f17_BR2, LM_FL_f20_BR2, LM_FL_f21_BR2, LM_FL_f25_BR2, LM_FL_f28_BR2, LM_IN_l47_BR2)
Dim BP_BR3: BP_BR3=Array(BM_BR3, LM_GIRight_GI041_BR3, LM_GITop_BR3, LM_FL_f20_BR3, LM_FL_f21_BR3, LM_FL_f22_BR3, LM_FL_f25_BR3, LM_FL_f28_BR3, LM_IN_l53_BR3, LM_IN_l54_BR3)
Dim BP_BS1: BP_BS1=Array(BM_BS1, LM_GIRight_GI041_BS1, LM_GITop_BS1, LM_FL_f20_BS1, LM_FL_f22_BS1, LM_IN_l47_BS1)
Dim BP_BS2: BP_BS2=Array(BM_BS2, LM_GITop_BS2, LM_FL_f17_BS2, LM_FL_f20_BS2, LM_FL_f28_BS2)
Dim BP_BS3: BP_BS3=Array(BM_BS3, LM_GITop_BS3, LM_FL_f20_BS3, LM_FL_f22_BS3, LM_IN_l53_BS3)
Dim BP_LEMK: BP_LEMK=Array(BM_LEMK, LM_GILeft_GI013_LEMK)
Dim BP_LFlip: BP_LFlip=Array(BM_LFlip, LM_GILeft_GI010_LFlip, LM_GILeft_GI011_LFlip, LM_GILeft_GI012_LFlip, LM_GILeft_GI013_LFlip, LM_FL_f27_LFlip)
Dim BP_LFlipU: BP_LFlipU=Array(BM_LFlipU, LM_GIRight_GI001_LFlipU, LM_GIRight_GI004_LFlipU, LM_GILeft_GI010_LFlipU, LM_GILeft_GI011_LFlipU, LM_GILeft_GI012_LFlipU, LM_GILeft_GI013_LFlipU, LM_FL_f27_LFlipU)
Dim BP_LSling0: BP_LSling0=Array(BM_LSling0, LM_GIRight_GI003_LSling0, LM_GILeft_GI012_LSling0, LM_GILeft_GI013_LSling0)
Dim BP_LSling1: BP_LSling1=Array(BM_LSling1, LM_GIRight_GI003_LSling1, LM_GILeft_GI012_LSling1, LM_GILeft_GI013_LSling1, LM_GILeft_GI014_LSling1)
Dim BP_LSling2: BP_LSling2=Array(BM_LSling2, LM_GIRight_GI003_LSling2, LM_GILeft_GI012_LSling2, LM_GILeft_GI013_LSling2)
Dim BP_Layer0: BP_Layer0=Array(BM_Layer0, LM_GIRight_GI041_Layer0, LM_GILeft_GI014_Layer0, LM_GITop_Layer0, LM_FL_f20_Layer0, LM_FL_f21_Layer0, LM_FL_f22_Layer0, LM_FL_f25_Layer0, LM_IN_l47_Layer0)
Dim BP_Layer1: BP_Layer1=Array(BM_Layer1, LM_GIRight_GI003_Layer1, LM_GIRight_GI004_Layer1, LM_GIRight_GI041_Layer1, LM_GILeft_GI012_Layer1, LM_GILeft_GI013_Layer1, LM_GILeft_GI014_Layer1, LM_GITop_Layer1, LM_FL_f17_Layer1, LM_FL_f18_Layer1, LM_FL_f19_Layer1, LM_FL_f20_Layer1, LM_FL_f21_Layer1, LM_FL_f22_Layer1, LM_FL_f25_Layer1, LM_FL_f26_Layer1, LM_FL_f27_Layer1, LM_FL_f28_Layer1, LM_IN_l42_Layer1, LM_IN_l43_Layer1, LM_IN_l46_Layer1, LM_IN_l47_Layer1, LM_IN_l48_Layer1, LM_IN_l53_Layer1, LM_IN_l54_Layer1, LM_IN_l57_Layer1, LM_IN_l58_Layer1, LM_IN_l64_Layer1, LM_IN_l68_Layer1, LM_IN_l73_Layer1, LM_IN_l74_Layer1, LM_IN_l76_Layer1, LM_IN_l77_Layer1, LM_IN_l78_Layer1, LM_IN_l86_Layer1)
Dim BP_Layer2: BP_Layer2=Array(BM_Layer2, LM_GIRight_GI041_Layer2, LM_GILeft_GI014_Layer2, LM_GITop_Layer2, LM_FL_f17_Layer2, LM_FL_f20_Layer2, LM_FL_f21_Layer2, LM_FL_f22_Layer2, LM_FL_f25_Layer2, LM_FL_f28_Layer2, LM_IN_l47_Layer2, LM_IN_l53_Layer2, LM_IN_l54_Layer2, LM_IN_l67_Layer2)
Dim BP_Layer3: BP_Layer3=Array(BM_Layer3, LM_GIRight_GI003_Layer3, LM_GIRight_GI004_Layer3, LM_GIRight_GI041_Layer3, LM_GILeft_GI014_Layer3, LM_GITop_Layer3, LM_FL_f17_Layer3, LM_FL_f18_Layer3, LM_FL_f19_Layer3, LM_FL_f21_Layer3, LM_FL_f22_Layer3, LM_FL_f25_Layer3, LM_FL_f27_Layer3, LM_FL_f28_Layer3, LM_IN_l44_Layer3, LM_IN_l45_Layer3, LM_IN_l47_Layer3, LM_IN_l51_Layer3, LM_IN_l68_Layer3, LM_IN_l77_Layer3, LM_IN_l85_Layer3, LM_IN_l86_Layer3)
Dim BP_Layer4: BP_Layer4=Array(BM_Layer4, LM_GIRight_GI003_Layer4, LM_GIRight_GI004_Layer4, LM_GIRight_GI041_Layer4, LM_GILeft_GI012_Layer4, LM_GILeft_GI013_Layer4, LM_GILeft_GI014_Layer4, LM_GITop_Layer4, LM_FL_f17_Layer4, LM_FL_f18_Layer4, LM_FL_f19_Layer4, LM_FL_f20_Layer4, LM_FL_f21_Layer4, LM_FL_f22_Layer4, LM_FL_f25_Layer4, LM_FL_f26_Layer4, LM_FL_f27_Layer4, LM_FL_f28_Layer4, LM_IN_l13_Layer4, LM_IN_l15_Layer4, LM_IN_l16_Layer4, LM_IN_l17_Layer4, LM_IN_l18_Layer4, LM_IN_l21_Layer4, LM_IN_l22_Layer4, LM_IN_l23_Layer4, LM_IN_l24_Layer4, LM_IN_l25_Layer4, LM_IN_l27_Layer4, LM_IN_l31_Layer4, LM_IN_l32_Layer4, LM_IN_l41_Layer4, LM_IN_l42_Layer4, LM_IN_l43_Layer4, LM_IN_l44_Layer4, LM_IN_l45_Layer4, LM_IN_l46_Layer4, LM_IN_l47_Layer4, LM_IN_l48_Layer4, LM_IN_l52_Layer4, LM_IN_l53_Layer4, LM_IN_l62_Layer4, LM_IN_l63_Layer4, LM_IN_l64_Layer4, LM_IN_l66_Layer4, LM_IN_l67_Layer4, LM_IN_l68_Layer4, LM_IN_l71_Layer4, LM_IN_l74_Layer4, LM_IN_l75_Layer4, LM_IN_l76_Layer4, LM_IN_l77_Layer4, LM_IN_l78_Layer4, _
  LM_IN_l86_Layer4)
Dim BP_Layer5: BP_Layer5=Array(BM_Layer5, LM_GIRight_GI004_Layer5, LM_GIRight_GI041_Layer5, LM_GILeft_GI012_Layer5, LM_GILeft_GI014_Layer5, LM_GITop_Layer5, LM_FL_f17_Layer5, LM_FL_f18_Layer5, LM_FL_f19_Layer5, LM_FL_f20_Layer5, LM_FL_f21_Layer5, LM_FL_f22_Layer5, LM_FL_f25_Layer5, LM_FL_f26_Layer5, LM_FL_f28_Layer5, LM_IN_l44_Layer5, LM_IN_l47_Layer5, LM_IN_l48_Layer5, LM_IN_l51_Layer5, LM_IN_l52_Layer5, LM_IN_l53_Layer5, LM_IN_l54_Layer5, LM_IN_l67_Layer5, LM_IN_l68_Layer5, LM_IN_l71_Layer5, LM_IN_l76_Layer5, LM_IN_l77_Layer5, LM_IN_l78_Layer5, LM_IN_l85_Layer5, LM_IN_l86_Layer5)
Dim BP_Layer6: BP_Layer6=Array(BM_Layer6, LM_GIRight_GI041_Layer6, LM_GILeft_GI014_Layer6, LM_GITop_Layer6, LM_FL_f17_Layer6, LM_FL_f18_Layer6, LM_FL_f19_Layer6, LM_FL_f21_Layer6, LM_FL_f22_Layer6, LM_FL_f25_Layer6, LM_FL_f26_Layer6, LM_FL_f28_Layer6, LM_IN_l31_Layer6, LM_IN_l32_Layer6, LM_IN_l43_Layer6, LM_IN_l64_Layer6, LM_IN_l71_Layer6, LM_IN_l72_Layer6, LM_IN_l76_Layer6, LM_IN_l77_Layer6, LM_IN_l78_Layer6)
Dim BP_Layer7: BP_Layer7=Array(BM_Layer7, LM_GIRight_GI001_Layer7, LM_GIRight_GI002_Layer7, LM_GIRight_GI003_Layer7, LM_GIRight_GI004_Layer7, LM_GIRight_GI041_Layer7, LM_GILeft_GI012_Layer7, LM_GILeft_GI013_Layer7, LM_GILeft_GI014_Layer7, LM_GITop_Layer7, LM_FL_f19_Layer7, LM_FL_f22_Layer7, LM_FL_f25_Layer7, LM_FL_f27_Layer7, LM_FL_f28_Layer7, LM_IN_l21_Layer7, LM_IN_l22_Layer7, LM_IN_l23_Layer7, LM_IN_l24_Layer7, LM_IN_l31_Layer7, LM_IN_l32_Layer7, LM_IN_l33_Layer7, LM_IN_l34_Layer7, LM_IN_l35_Layer7, LM_IN_l36_Layer7, LM_IN_l44_Layer7, LM_IN_l45_Layer7, LM_IN_l46_Layer7, LM_IN_l47_Layer7, LM_IN_l55_Layer7, LM_IN_l56_Layer7, LM_IN_l57_Layer7, LM_IN_l58_Layer7, LM_IN_l68_Layer7, LM_IN_l71_Layer7, LM_IN_l77_Layer7, LM_IN_l78_Layer7, LM_IN_l85_Layer7, LM_IN_l86_Layer7)
Dim BP_Layer8: BP_Layer8=Array(BM_Layer8, LM_GIRight_GI041_Layer8, LM_FL_f19_Layer8, LM_IN_l55_Layer8, LM_IN_l68_Layer8, LM_IN_l85_Layer8, LM_IN_l86_Layer8)
Dim BP_LockPin: BP_LockPin=Array(BM_LockPin, LM_GILeft_GI013_LockPin, LM_GILeft_GI014_LockPin, LM_IN_l18_LockPin)
Dim BP_Lpost: BP_Lpost=Array(BM_Lpost, LM_GIRight_GI041_Lpost, LM_GILeft_GI012_Lpost, LM_GILeft_GI013_Lpost, LM_GILeft_GI014_Lpost)
Dim BP_Playfield: BP_Playfield=Array(BM_Playfield, LM_GIRight_GI001_Playfield, LM_GIRight_GI002_Playfield, LM_GIRight_GI003_Playfield, LM_GIRight_GI004_Playfield, LM_GIRight_GI041_Playfield, LM_GILeft_GI010_Playfield, LM_GILeft_GI011_Playfield, LM_GILeft_GI012_Playfield, LM_GILeft_GI013_Playfield, LM_GILeft_GI014_Playfield, LM_GITop_Playfield, LM_FL_f17_Playfield, LM_FL_f18_Playfield, LM_FL_f19_Playfield, LM_FL_f20_Playfield, LM_FL_f21_Playfield, LM_FL_f22_Playfield, LM_FL_f25_Playfield, LM_FL_f26_Playfield, LM_FL_f27_Playfield, LM_FL_f28_Playfield, LM_IN_l11_Playfield, LM_IN_l12_Playfield, LM_IN_l13_Playfield, LM_IN_l14_Playfield, LM_IN_l15_Playfield, LM_IN_l16_Playfield, LM_IN_l17_Playfield, LM_IN_l18_Playfield, LM_IN_l21_Playfield, LM_IN_l22_Playfield, LM_IN_l23_Playfield, LM_IN_l24_Playfield, LM_IN_l25_Playfield, LM_IN_l26_Playfield, LM_IN_l27_Playfield, LM_IN_l31_Playfield, LM_IN_l32_Playfield, LM_IN_l35_Playfield, LM_IN_l36_Playfield, LM_IN_l41_Playfield, LM_IN_l42_Playfield, LM_IN_l44_Playfield, _
  LM_IN_l45_Playfield, LM_IN_l46_Playfield, LM_IN_l47_Playfield, LM_IN_l51_Playfield, LM_IN_l52_Playfield, LM_IN_l53_Playfield, LM_IN_l54_Playfield, LM_IN_l55_Playfield, LM_IN_l56_Playfield, LM_IN_l57_Playfield, LM_IN_l58_Playfield, LM_IN_l61_Playfield, LM_IN_l62_Playfield, LM_IN_l63_Playfield, LM_IN_l64_Playfield, LM_IN_l65_Playfield, LM_IN_l66_Playfield, LM_IN_l67_Playfield, LM_IN_l68_Playfield, LM_IN_l71_Playfield, LM_IN_l72_Playfield, LM_IN_l73_Playfield, LM_IN_l74_Playfield, LM_IN_l75_Playfield, LM_IN_l76_Playfield, LM_IN_l77_Playfield, LM_IN_l78_Playfield, LM_IN_l81_Playfield, LM_IN_l85_Playfield, LM_IN_l86_Playfield)
Dim BP_REMK: BP_REMK=Array(BM_REMK, LM_GIRight_GI003_REMK, LM_GIRight_GI004_REMK)
Dim BP_RFlip: BP_RFlip=Array(BM_RFlip, LM_GIRight_GI001_RFlip, LM_GIRight_GI002_RFlip, LM_GIRight_GI003_RFlip, LM_GIRight_GI004_RFlip, LM_GILeft_GI013_RFlip, LM_FL_f27_RFlip)
Dim BP_RFlipU: BP_RFlipU=Array(BM_RFlipU, LM_GIRight_GI001_RFlipU, LM_GIRight_GI002_RFlipU, LM_GIRight_GI003_RFlipU, LM_GIRight_GI004_RFlipU, LM_GILeft_GI010_RFlipU, LM_GILeft_GI013_RFlipU)
Dim BP_RSling0: BP_RSling0=Array(BM_RSling0, LM_GIRight_GI003_RSling0, LM_GIRight_GI004_RSling0, LM_GIRight_GI041_RSling0, LM_GILeft_GI012_RSling0)
Dim BP_RSling1: BP_RSling1=Array(BM_RSling1, LM_GIRight_GI003_RSling1, LM_GIRight_GI004_RSling1, LM_GIRight_GI041_RSling1, LM_GILeft_GI012_RSling1, LM_GILeft_GI013_RSling1, LM_IN_l56_RSling1)
Dim BP_RSling2: BP_RSling2=Array(BM_RSling2, LM_GIRight_GI003_RSling2, LM_GIRight_GI004_RSling2, LM_GIRight_GI041_RSling2, LM_GILeft_GI012_RSling2)
Dim BP_Rivet_1_008: BP_Rivet_1_008=Array(BM_Rivet_1_008, LM_GITop_Rivet_1_008, LM_FL_f17_Rivet_1_008, LM_FL_f28_Rivet_1_008)
Dim BP_Rpost: BP_Rpost=Array(BM_Rpost, LM_GIRight_GI003_Rpost, LM_GIRight_GI004_Rpost, LM_GIRight_GI041_Rpost, LM_FL_f27_Rpost, LM_IN_l68_Rpost, LM_IN_l85_Rpost, LM_IN_l86_Rpost)
Dim BP_diverterside01: BP_diverterside01=Array(BM_diverterside01, LM_GITop_diverterside01, LM_FL_f17_diverterside01, LM_FL_f18_diverterside01, LM_FL_f21_diverterside01, LM_FL_f22_diverterside01, LM_FL_f28_diverterside01)
Dim BP_diverterside02: BP_diverterside02=Array(BM_diverterside02, LM_GITop_diverterside02, LM_FL_f17_diverterside02, LM_FL_f18_diverterside02, LM_FL_f21_diverterside02, LM_FL_f25_diverterside02, LM_FL_f28_diverterside02)
Dim BP_gate002: BP_gate002=Array(BM_gate002, LM_GITop_gate002, LM_FL_f17_gate002, LM_IN_l66_gate002)
Dim BP_gate003: BP_gate003=Array(BM_gate003, LM_GITop_gate003, LM_IN_l67_gate003)
Dim BP_gate004: BP_gate004=Array(BM_gate004, LM_GIRight_GI041_gate004, LM_FL_f19_gate004, LM_FL_f22_gate004, LM_FL_f25_gate004)
Dim BP_gate006: BP_gate006=Array(BM_gate006, LM_GITop_gate006, LM_FL_f22_gate006, LM_FL_f25_gate006, LM_IN_l76_gate006)
Dim BP_gate007: BP_gate007=Array(BM_gate007, LM_GIRight_GI041_gate007, LM_GITop_gate007, LM_FL_f22_gate007)
Dim BP_goalkeeper: BP_goalkeeper=Array(BM_goalkeeper, LM_GIRight_GI041_goalkeeper, LM_GILeft_GI014_goalkeeper, LM_GITop_goalkeeper, LM_FL_f17_goalkeeper, LM_FL_f18_goalkeeper, LM_FL_f21_goalkeeper, LM_FL_f22_goalkeeper, LM_FL_f25_goalkeeper, LM_FL_f28_goalkeeper, LM_IN_l31_goalkeeper, LM_IN_l51_goalkeeper, LM_IN_l52_goalkeeper, LM_IN_l53_goalkeeper, LM_IN_l54_goalkeeper, LM_IN_l72_goalkeeper, LM_IN_l76_goalkeeper, LM_IN_l78_goalkeeper, LM_IN_l81_goalkeeper)
Dim BP_parts: BP_parts=Array(BM_parts, LM_GIRight_GI001_parts, LM_GIRight_GI002_parts, LM_GIRight_GI003_parts, LM_GIRight_GI004_parts, LM_GIRight_GI041_parts, LM_GILeft_GI010_parts, LM_GILeft_GI011_parts, LM_GILeft_GI012_parts, LM_GILeft_GI013_parts, LM_GILeft_GI014_parts, LM_GITop_parts, LM_FL_f17_parts, LM_FL_f18_parts, LM_FL_f19_parts, LM_FL_f20_parts, LM_FL_f21_parts, LM_FL_f22_parts, LM_FL_f25_parts, LM_FL_f26_parts, LM_FL_f27_parts, LM_FL_f28_parts, LM_IN_l11_parts, LM_IN_l12_parts, LM_IN_l13_parts, LM_IN_l14_parts, LM_IN_l15_parts, LM_IN_l16_parts, LM_IN_l17_parts, LM_IN_l18_parts, LM_IN_l21_parts, LM_IN_l22_parts, LM_IN_l23_parts, LM_IN_l24_parts, LM_IN_l25_parts, LM_IN_l26_parts, LM_IN_l27_parts, LM_IN_l28_parts, LM_IN_l31_parts, LM_IN_l32_parts, LM_IN_l33_parts, LM_IN_l34_parts, LM_IN_l35_parts, LM_IN_l36_parts, LM_IN_l37_parts, LM_IN_l38_parts, LM_IN_l41_parts, LM_IN_l42_parts, LM_IN_l43_parts, LM_IN_l44_parts, LM_IN_l45_parts, LM_IN_l46_parts, LM_IN_l47_parts, LM_IN_l48_parts, LM_IN_l51_parts, _
  LM_IN_l52_parts, LM_IN_l53_parts, LM_IN_l54_parts, LM_IN_l55_parts, LM_IN_l56_parts, LM_IN_l57_parts, LM_IN_l58_parts, LM_IN_l61_parts, LM_IN_l62_parts, LM_IN_l63_parts, LM_IN_l64_parts, LM_IN_l65_parts, LM_IN_l66_parts, LM_IN_l67_parts, LM_IN_l68_parts, LM_IN_l71_parts, LM_IN_l72_parts, LM_IN_l73_parts, LM_IN_l74_parts, LM_IN_l75_parts, LM_IN_l76_parts, LM_IN_l77_parts, LM_IN_l78_parts, LM_IN_l81_parts, LM_IN_l82_parts, LM_IN_l83_parts, LM_IN_l84_parts, LM_IN_l85_parts, LM_IN_l86_parts)
Dim BP_plunger: BP_plunger=Array(BM_plunger, LM_FL_f19_plunger)
Dim BP_soccerBall_001: BP_soccerBall_001=Array(BM_soccerBall_001, LM_GIRight_GI041_soccerBall_001, LM_GILeft_GI013_soccerBall_001, LM_GILeft_GI014_soccerBall_001, LM_GITop_soccerBall_001, LM_FL_f18_soccerBall_001, LM_FL_f19_soccerBall_001, LM_FL_f20_soccerBall_001, LM_FL_f21_soccerBall_001, LM_FL_f22_soccerBall_001, LM_FL_f25_soccerBall_001, LM_FL_f26_soccerBall_001, LM_FL_f28_soccerBall_001, LM_IN_l31_soccerBall_001, LM_IN_l32_soccerBall_001, LM_IN_l76_soccerBall_001, LM_IN_l77_soccerBall_001, LM_IN_l78_soccerBall_001)
Dim BP_soccerBall_002: BP_soccerBall_002=Array(BM_soccerBall_002, LM_GIRight_GI041_soccerBall_002, LM_GILeft_GI013_soccerBall_002, LM_GILeft_GI014_soccerBall_002, LM_GITop_soccerBall_002, LM_FL_f18_soccerBall_002, LM_FL_f19_soccerBall_002, LM_FL_f20_soccerBall_002, LM_FL_f21_soccerBall_002, LM_FL_f22_soccerBall_002, LM_FL_f25_soccerBall_002, LM_FL_f26_soccerBall_002, LM_FL_f28_soccerBall_002, LM_IN_l31_soccerBall_002, LM_IN_l76_soccerBall_002, LM_IN_l77_soccerBall_002, LM_IN_l78_soccerBall_002)
Dim BP_soccerBall_003: BP_soccerBall_003=Array(BM_soccerBall_003, LM_GIRight_GI041_soccerBall_003, LM_GILeft_GI013_soccerBall_003, LM_GILeft_GI014_soccerBall_003, LM_GITop_soccerBall_003, LM_FL_f18_soccerBall_003, LM_FL_f19_soccerBall_003, LM_FL_f20_soccerBall_003, LM_FL_f21_soccerBall_003, LM_FL_f22_soccerBall_003, LM_FL_f25_soccerBall_003, LM_FL_f26_soccerBall_003, LM_FL_f28_soccerBall_003, LM_IN_l76_soccerBall_003, LM_IN_l77_soccerBall_003, LM_IN_l78_soccerBall_003, LM_IN_l82_soccerBall_003, LM_IN_l85_soccerBall_003)
Dim BP_sw15: BP_sw15=Array(BM_sw15, LM_GILeft_GI013_sw15, LM_GILeft_GI014_sw15)
Dim BP_sw16: BP_sw16=Array(BM_sw16, LM_GIRight_GI041_sw16, LM_GITop_sw16, LM_FL_f18_sw16, LM_FL_f22_sw16, LM_FL_f25_sw16, LM_IN_l52_sw16, LM_IN_l54_sw16)
Dim BP_sw17: BP_sw17=Array(BM_sw17, LM_GIRight_GI003_sw17, LM_GIRight_GI004_sw17, LM_GIRight_GI041_sw17, LM_FL_f27_sw17)
Dim BP_sw18: BP_sw18=Array(BM_sw18, LM_GIRight_GI002_sw18, LM_GIRight_GI004_sw18, LM_GIRight_GI041_sw18, LM_FL_f27_sw18)
Dim BP_sw25: BP_sw25=Array(BM_sw25, LM_GIRight_GI041_sw25, LM_GILeft_GI012_sw25, LM_GILeft_GI013_sw25, LM_GILeft_GI014_sw25, LM_GITop_sw25, LM_FL_f22_sw25, LM_FL_f25_sw25, LM_IN_l31_sw25, LM_IN_l32_sw25, LM_IN_l77_sw25, LM_IN_l78_sw25)
Dim BP_sw26: BP_sw26=Array(BM_sw26, LM_GILeft_GI014_sw26)
Dim BP_sw27: BP_sw27=Array(BM_sw27, LM_GILeft_GI014_sw27, LM_GITop_sw27, LM_FL_f25_sw27, LM_FL_f26_sw27, LM_FL_f28_sw27, LM_IN_l46_sw27, LM_IN_l62_sw27)
Dim BP_sw28: BP_sw28=Array(BM_sw28, LM_GIRight_GI004_sw28, LM_GILeft_GI013_sw28, LM_GILeft_GI014_sw28)
Dim BP_sw37: BP_sw37=Array(BM_sw37, LM_GIRight_GI003_sw37, LM_GIRight_GI004_sw37, LM_GIRight_GI041_sw37, LM_GILeft_GI014_sw37, LM_IN_l85_sw37, LM_IN_l86_sw37)
Dim BP_sw38: BP_sw38=Array(BM_sw38, LM_FL_f19_sw38)
Dim BP_sw47: BP_sw47=Array(BM_sw47, LM_GITop_sw47, LM_FL_f25_sw47)
Dim BP_sw61: BP_sw61=Array(BM_sw61, LM_GITop_sw61, LM_FL_f21_sw61, LM_IN_l81_sw61)
Dim BP_sw62: BP_sw62=Array(BM_sw62, LM_GITop_sw62, LM_FL_f21_sw62, LM_IN_l82_sw62)
Dim BP_sw63: BP_sw63=Array(BM_sw63, LM_IN_l83_sw63)
Dim BP_sw64: BP_sw64=Array(BM_sw64, LM_IN_l84_sw64)
Dim BP_sw66: BP_sw66=Array(BM_sw66)
Dim BP_sw67: BP_sw67=Array(BM_sw67, LM_GILeft_GI014_sw67, LM_IN_l76_sw67)
Dim BP_sw71: BP_sw71=Array(BM_sw71, LM_GITop_sw71)
Dim BP_sw74: BP_sw74=Array(BM_sw74, LM_FL_f19_sw74)
Dim BP_sw76: BP_sw76=Array(BM_sw76, LM_GIRight_GI003_sw76)
Dim BP_sw77: BP_sw77=Array(BM_sw77, LM_GILeft_GI014_sw77)
Dim BP_sw78: BP_sw78=Array(BM_sw78, LM_GILeft_GI014_sw78)
Dim BP_sw86: BP_sw86=Array(BM_sw86, LM_FL_f27_sw86)
Dim BP_sw87: BP_sw87=Array(BM_sw87, LM_GITop_sw87)
Dim BP_sw88: BP_sw88=Array(BM_sw88, LM_GITop_sw88)
Dim BP_underPF: BP_underPF=Array(BM_underPF, LM_GIRight_GI001_underPF, LM_GIRight_GI002_underPF, LM_GIRight_GI003_underPF, LM_GIRight_GI004_underPF, LM_GIRight_GI041_underPF, LM_GILeft_GI012_underPF, LM_GILeft_GI014_underPF, LM_GITop_underPF, LM_FL_f17_underPF, LM_FL_f18_underPF, LM_FL_f19_underPF, LM_FL_f20_underPF, LM_FL_f21_underPF, LM_FL_f22_underPF, LM_FL_f25_underPF, LM_FL_f26_underPF, LM_FL_f27_underPF, LM_FL_f28_underPF, LM_IN_l11_underPF, LM_IN_l12_underPF, LM_IN_l13_underPF, LM_IN_l14_underPF, LM_IN_l15_underPF, LM_IN_l16_underPF, LM_IN_l17_underPF, LM_IN_l18_underPF, LM_IN_l21_underPF, LM_IN_l22_underPF, LM_IN_l23_underPF, LM_IN_l24_underPF, LM_IN_l25_underPF, LM_IN_l26_underPF, LM_IN_l27_underPF, LM_IN_l28_underPF, LM_IN_l31_underPF, LM_IN_l32_underPF, LM_IN_l33_underPF, LM_IN_l34_underPF, LM_IN_l35_underPF, LM_IN_l36_underPF, LM_IN_l37_underPF, LM_IN_l38_underPF, LM_IN_l41_underPF, LM_IN_l42_underPF, LM_IN_l43_underPF, LM_IN_l44_underPF, LM_IN_l45_underPF, LM_IN_l46_underPF, LM_IN_l47_underPF, _
  LM_IN_l48_underPF, LM_IN_l51_underPF, LM_IN_l52_underPF, LM_IN_l53_underPF, LM_IN_l54_underPF, LM_IN_l55_underPF, LM_IN_l56_underPF, LM_IN_l57_underPF, LM_IN_l58_underPF, LM_IN_l61_underPF, LM_IN_l62_underPF, LM_IN_l63_underPF, LM_IN_l64_underPF, LM_IN_l65_underPF, LM_IN_l66_underPF, LM_IN_l67_underPF, LM_IN_l72_underPF, LM_IN_l73_underPF, LM_IN_l74_underPF, LM_IN_l75_underPF, LM_IN_l77_underPF, LM_IN_l78_underPF, LM_IN_l81_underPF, LM_IN_l84_underPF, LM_IN_l85_underPF, LM_IN_l86_underPF)
' Arrays per lighting scenario
Dim BL_FL_f17: BL_FL_f17=Array(LM_FL_f17_BR2, LM_FL_f17_BS2, LM_FL_f17_Layer1, LM_FL_f17_Layer2, LM_FL_f17_Layer3, LM_FL_f17_Layer4, LM_FL_f17_Layer5, LM_FL_f17_Layer6, LM_FL_f17_Playfield, LM_FL_f17_Rivet_1_008, LM_FL_f17_diverterside01, LM_FL_f17_diverterside02, LM_FL_f17_gate002, LM_FL_f17_goalkeeper, LM_FL_f17_parts, LM_FL_f17_underPF)
Dim BL_FL_f18: BL_FL_f18=Array(LM_FL_f18_Layer1, LM_FL_f18_Layer3, LM_FL_f18_Layer4, LM_FL_f18_Layer5, LM_FL_f18_Layer6, LM_FL_f18_Playfield, LM_FL_f18_diverterside01, LM_FL_f18_diverterside02, LM_FL_f18_goalkeeper, LM_FL_f18_parts, LM_FL_f18_soccerBall_001, LM_FL_f18_soccerBall_002, LM_FL_f18_soccerBall_003, LM_FL_f18_sw16, LM_FL_f18_underPF)
Dim BL_FL_f19: BL_FL_f19=Array(LM_FL_f19_Layer1, LM_FL_f19_Layer3, LM_FL_f19_Layer4, LM_FL_f19_Layer5, LM_FL_f19_Layer6, LM_FL_f19_Layer7, LM_FL_f19_Layer8, LM_FL_f19_Playfield, LM_FL_f19_gate004, LM_FL_f19_parts, LM_FL_f19_plunger, LM_FL_f19_soccerBall_001, LM_FL_f19_soccerBall_002, LM_FL_f19_soccerBall_003, LM_FL_f19_sw38, LM_FL_f19_sw74, LM_FL_f19_underPF)
Dim BL_FL_f20: BL_FL_f20=Array(LM_FL_f20_BR1, LM_FL_f20_BR2, LM_FL_f20_BR3, LM_FL_f20_BS1, LM_FL_f20_BS2, LM_FL_f20_BS3, LM_FL_f20_Layer0, LM_FL_f20_Layer1, LM_FL_f20_Layer2, LM_FL_f20_Layer4, LM_FL_f20_Layer5, LM_FL_f20_Playfield, LM_FL_f20_parts, LM_FL_f20_soccerBall_001, LM_FL_f20_soccerBall_002, LM_FL_f20_soccerBall_003, LM_FL_f20_underPF)
Dim BL_FL_f21: BL_FL_f21=Array(LM_FL_f21_BR2, LM_FL_f21_BR3, LM_FL_f21_Layer0, LM_FL_f21_Layer1, LM_FL_f21_Layer2, LM_FL_f21_Layer3, LM_FL_f21_Layer4, LM_FL_f21_Layer5, LM_FL_f21_Layer6, LM_FL_f21_Playfield, LM_FL_f21_diverterside01, LM_FL_f21_diverterside02, LM_FL_f21_goalkeeper, LM_FL_f21_parts, LM_FL_f21_soccerBall_001, LM_FL_f21_soccerBall_002, LM_FL_f21_soccerBall_003, LM_FL_f21_sw61, LM_FL_f21_sw62, LM_FL_f21_underPF)
Dim BL_FL_f22: BL_FL_f22=Array(LM_FL_f22_BR1, LM_FL_f22_BR3, LM_FL_f22_BS1, LM_FL_f22_BS3, LM_FL_f22_Layer0, LM_FL_f22_Layer1, LM_FL_f22_Layer2, LM_FL_f22_Layer3, LM_FL_f22_Layer4, LM_FL_f22_Layer5, LM_FL_f22_Layer6, LM_FL_f22_Layer7, LM_FL_f22_Playfield, LM_FL_f22_diverterside01, LM_FL_f22_gate004, LM_FL_f22_gate006, LM_FL_f22_gate007, LM_FL_f22_goalkeeper, LM_FL_f22_parts, LM_FL_f22_soccerBall_001, LM_FL_f22_soccerBall_002, LM_FL_f22_soccerBall_003, LM_FL_f22_sw16, LM_FL_f22_sw25, LM_FL_f22_underPF)
Dim BL_FL_f25: BL_FL_f25=Array(LM_FL_f25_BR1, LM_FL_f25_BR2, LM_FL_f25_BR3, LM_FL_f25_Layer0, LM_FL_f25_Layer1, LM_FL_f25_Layer2, LM_FL_f25_Layer3, LM_FL_f25_Layer4, LM_FL_f25_Layer5, LM_FL_f25_Layer6, LM_FL_f25_Layer7, LM_FL_f25_Playfield, LM_FL_f25_diverterside02, LM_FL_f25_gate004, LM_FL_f25_gate006, LM_FL_f25_goalkeeper, LM_FL_f25_parts, LM_FL_f25_soccerBall_001, LM_FL_f25_soccerBall_002, LM_FL_f25_soccerBall_003, LM_FL_f25_sw16, LM_FL_f25_sw25, LM_FL_f25_sw27, LM_FL_f25_sw47, LM_FL_f25_underPF)
Dim BL_FL_f26: BL_FL_f26=Array(LM_FL_f26_Layer1, LM_FL_f26_Layer4, LM_FL_f26_Layer5, LM_FL_f26_Layer6, LM_FL_f26_Playfield, LM_FL_f26_parts, LM_FL_f26_soccerBall_001, LM_FL_f26_soccerBall_002, LM_FL_f26_soccerBall_003, LM_FL_f26_sw27, LM_FL_f26_underPF)
Dim BL_FL_f27: BL_FL_f27=Array(LM_FL_f27_LFlip, LM_FL_f27_LFlipU, LM_FL_f27_Layer1, LM_FL_f27_Layer3, LM_FL_f27_Layer4, LM_FL_f27_Layer7, LM_FL_f27_Playfield, LM_FL_f27_RFlip, LM_FL_f27_Rpost, LM_FL_f27_parts, LM_FL_f27_sw17, LM_FL_f27_sw18, LM_FL_f27_sw86, LM_FL_f27_underPF)
Dim BL_FL_f28: BL_FL_f28=Array(LM_FL_f28_BR2, LM_FL_f28_BR3, LM_FL_f28_BS2, LM_FL_f28_Layer1, LM_FL_f28_Layer2, LM_FL_f28_Layer3, LM_FL_f28_Layer4, LM_FL_f28_Layer5, LM_FL_f28_Layer6, LM_FL_f28_Layer7, LM_FL_f28_Playfield, LM_FL_f28_Rivet_1_008, LM_FL_f28_diverterside01, LM_FL_f28_diverterside02, LM_FL_f28_goalkeeper, LM_FL_f28_parts, LM_FL_f28_soccerBall_001, LM_FL_f28_soccerBall_002, LM_FL_f28_soccerBall_003, LM_FL_f28_sw27, LM_FL_f28_underPF)
'Dim BL_GILeft_GI010: BL_GILeft_GI010=Array(LM_GILeft_GI010_LFlip, LM_GILeft_GI010_LFlipU, LM_GILeft_GI010_Playfield, LM_GILeft_GI010_RFlipU, LM_GILeft_GI010_parts)
'Dim BL_GILeft_GI011: BL_GILeft_GI011=Array(LM_GILeft_GI011_LFlip, LM_GILeft_GI011_LFlipU, LM_GILeft_GI011_Playfield, LM_GILeft_GI011_parts)
'Dim BL_GILeft_GI012: BL_GILeft_GI012=Array(LM_GILeft_GI012_LFlip, LM_GILeft_GI012_LFlipU, LM_GILeft_GI012_LSling0, LM_GILeft_GI012_LSling1, LM_GILeft_GI012_LSling2, LM_GILeft_GI012_Layer1, LM_GILeft_GI012_Layer4, LM_GILeft_GI012_Layer5, LM_GILeft_GI012_Layer7, LM_GILeft_GI012_Lpost, LM_GILeft_GI012_Playfield, LM_GILeft_GI012_RSling0, LM_GILeft_GI012_RSling1, LM_GILeft_GI012_RSling2, LM_GILeft_GI012_parts, LM_GILeft_GI012_sw25, LM_GILeft_GI012_underPF)
'Dim BL_GILeft_GI013: BL_GILeft_GI013=Array(LM_GILeft_GI013_LEMK, LM_GILeft_GI013_LFlip, LM_GILeft_GI013_LFlipU, LM_GILeft_GI013_LSling0, LM_GILeft_GI013_LSling1, LM_GILeft_GI013_LSling2, LM_GILeft_GI013_Layer1, LM_GILeft_GI013_Layer4, LM_GILeft_GI013_Layer7, LM_GILeft_GI013_LockPin, LM_GILeft_GI013_Lpost, LM_GILeft_GI013_Playfield, LM_GILeft_GI013_RFlip, LM_GILeft_GI013_RFlipU, LM_GILeft_GI013_RSling1, LM_GILeft_GI013_parts, LM_GILeft_GI013_soccerBall_001, LM_GILeft_GI013_soccerBall_002, LM_GILeft_GI013_soccerBall_003, LM_GILeft_GI013_sw15, LM_GILeft_GI013_sw25, LM_GILeft_GI013_sw28)
'Dim BL_GILeft_GI014: BL_GILeft_GI014=Array(LM_GILeft_GI014_BR1, LM_GILeft_GI014_LSling1, LM_GILeft_GI014_Layer0, LM_GILeft_GI014_Layer1, LM_GILeft_GI014_Layer2, LM_GILeft_GI014_Layer3, LM_GILeft_GI014_Layer4, LM_GILeft_GI014_Layer5, LM_GILeft_GI014_Layer6, LM_GILeft_GI014_Layer7, LM_GILeft_GI014_LockPin, LM_GILeft_GI014_Lpost, LM_GILeft_GI014_Playfield, LM_GILeft_GI014_goalkeeper, LM_GILeft_GI014_parts, LM_GILeft_GI014_soccerBall_001, LM_GILeft_GI014_soccerBall_002, LM_GILeft_GI014_soccerBall_003, LM_GILeft_GI014_sw15, LM_GILeft_GI014_sw25, LM_GILeft_GI014_sw26, LM_GILeft_GI014_sw27, LM_GILeft_GI014_sw28, LM_GILeft_GI014_sw37, LM_GILeft_GI014_sw67, LM_GILeft_GI014_sw77, LM_GILeft_GI014_sw78, LM_GILeft_GI014_underPF)
'Dim BL_GIRight_GI001: BL_GIRight_GI001=Array(LM_GIRight_GI001_LFlipU, LM_GIRight_GI001_Layer7, LM_GIRight_GI001_Playfield, LM_GIRight_GI001_RFlip, LM_GIRight_GI001_RFlipU, LM_GIRight_GI001_parts, LM_GIRight_GI001_underPF)
'Dim BL_GIRight_GI002: BL_GIRight_GI002=Array(LM_GIRight_GI002_Layer7, LM_GIRight_GI002_Playfield, LM_GIRight_GI002_RFlip, LM_GIRight_GI002_RFlipU, LM_GIRight_GI002_parts, LM_GIRight_GI002_sw18, LM_GIRight_GI002_underPF)
'Dim BL_GIRight_GI003: BL_GIRight_GI003=Array(LM_GIRight_GI003_LSling0, LM_GIRight_GI003_LSling1, LM_GIRight_GI003_LSling2, LM_GIRight_GI003_Layer1, LM_GIRight_GI003_Layer3, LM_GIRight_GI003_Layer4, LM_GIRight_GI003_Layer7, LM_GIRight_GI003_Playfield, LM_GIRight_GI003_REMK, LM_GIRight_GI003_RFlip, LM_GIRight_GI003_RFlipU, LM_GIRight_GI003_RSling0, LM_GIRight_GI003_RSling1, LM_GIRight_GI003_RSling2, LM_GIRight_GI003_Rpost, LM_GIRight_GI003_parts, LM_GIRight_GI003_sw17, LM_GIRight_GI003_sw37, LM_GIRight_GI003_sw76, LM_GIRight_GI003_underPF)
'Dim BL_GIRight_GI004: BL_GIRight_GI004=Array(LM_GIRight_GI004_LFlipU, LM_GIRight_GI004_Layer1, LM_GIRight_GI004_Layer3, LM_GIRight_GI004_Layer4, LM_GIRight_GI004_Layer5, LM_GIRight_GI004_Layer7, LM_GIRight_GI004_Playfield, LM_GIRight_GI004_REMK, LM_GIRight_GI004_RFlip, LM_GIRight_GI004_RFlipU, LM_GIRight_GI004_RSling0, LM_GIRight_GI004_RSling1, LM_GIRight_GI004_RSling2, LM_GIRight_GI004_Rpost, LM_GIRight_GI004_parts, LM_GIRight_GI004_sw17, LM_GIRight_GI004_sw18, LM_GIRight_GI004_sw28, LM_GIRight_GI004_sw37, LM_GIRight_GI004_underPF)
'Dim BL_GIRight_GI041: BL_GIRight_GI041=Array(LM_GIRight_GI041_BR1, LM_GIRight_GI041_BR3, LM_GIRight_GI041_BS1, LM_GIRight_GI041_Layer0, LM_GIRight_GI041_Layer1, LM_GIRight_GI041_Layer2, LM_GIRight_GI041_Layer3, LM_GIRight_GI041_Layer4, LM_GIRight_GI041_Layer5, LM_GIRight_GI041_Layer6, LM_GIRight_GI041_Layer7, LM_GIRight_GI041_Layer8, LM_GIRight_GI041_Lpost, LM_GIRight_GI041_Playfield, LM_GIRight_GI041_RSling0, LM_GIRight_GI041_RSling1, LM_GIRight_GI041_RSling2, LM_GIRight_GI041_Rpost, LM_GIRight_GI041_gate004, LM_GIRight_GI041_gate007, LM_GIRight_GI041_goalkeeper, LM_GIRight_GI041_parts, LM_GIRight_GI041_soccerBall_001, LM_GIRight_GI041_soccerBall_002, LM_GIRight_GI041_soccerBall_003, LM_GIRight_GI041_sw16, LM_GIRight_GI041_sw17, LM_GIRight_GI041_sw18, LM_GIRight_GI041_sw25, LM_GIRight_GI041_sw37, LM_GIRight_GI041_underPF)
'Dim BL_GITop: BL_GITop=Array(LM_GITop_BR1, LM_GITop_BR2, LM_GITop_BR3, LM_GITop_BS1, LM_GITop_BS2, LM_GITop_BS3, LM_GITop_Layer0, LM_GITop_Layer1, LM_GITop_Layer2, LM_GITop_Layer3, LM_GITop_Layer4, LM_GITop_Layer5, LM_GITop_Layer6, LM_GITop_Layer7, LM_GITop_Playfield, LM_GITop_Rivet_1_008, LM_GITop_diverterside01, LM_GITop_diverterside02, LM_GITop_gate002, LM_GITop_gate003, LM_GITop_gate006, LM_GITop_gate007, LM_GITop_goalkeeper, LM_GITop_parts, LM_GITop_soccerBall_001, LM_GITop_soccerBall_002, LM_GITop_soccerBall_003, LM_GITop_sw16, LM_GITop_sw25, LM_GITop_sw27, LM_GITop_sw47, LM_GITop_sw61, LM_GITop_sw62, LM_GITop_sw71, LM_GITop_sw87, LM_GITop_sw88, LM_GITop_underPF)
'Dim BL_IN_l11: BL_IN_l11=Array(LM_IN_l11_Playfield, LM_IN_l11_parts, LM_IN_l11_underPF)
'Dim BL_IN_l12: BL_IN_l12=Array(LM_IN_l12_Playfield, LM_IN_l12_parts, LM_IN_l12_underPF)
'Dim BL_IN_l13: BL_IN_l13=Array(LM_IN_l13_Layer4, LM_IN_l13_Playfield, LM_IN_l13_parts, LM_IN_l13_underPF)
'Dim BL_IN_l14: BL_IN_l14=Array(LM_IN_l14_Playfield, LM_IN_l14_parts, LM_IN_l14_underPF)
'Dim BL_IN_l15: BL_IN_l15=Array(LM_IN_l15_Layer4, LM_IN_l15_Playfield, LM_IN_l15_parts, LM_IN_l15_underPF)
'Dim BL_IN_l16: BL_IN_l16=Array(LM_IN_l16_Layer4, LM_IN_l16_Playfield, LM_IN_l16_parts, LM_IN_l16_underPF)
'Dim BL_IN_l17: BL_IN_l17=Array(LM_IN_l17_Layer4, LM_IN_l17_Playfield, LM_IN_l17_parts, LM_IN_l17_underPF)
'Dim BL_IN_l18: BL_IN_l18=Array(LM_IN_l18_Layer4, LM_IN_l18_LockPin, LM_IN_l18_Playfield, LM_IN_l18_parts, LM_IN_l18_underPF)
'Dim BL_IN_l21: BL_IN_l21=Array(LM_IN_l21_Layer4, LM_IN_l21_Layer7, LM_IN_l21_Playfield, LM_IN_l21_parts, LM_IN_l21_underPF)
'Dim BL_IN_l22: BL_IN_l22=Array(LM_IN_l22_Layer4, LM_IN_l22_Layer7, LM_IN_l22_Playfield, LM_IN_l22_parts, LM_IN_l22_underPF)
'Dim BL_IN_l23: BL_IN_l23=Array(LM_IN_l23_Layer4, LM_IN_l23_Layer7, LM_IN_l23_Playfield, LM_IN_l23_parts, LM_IN_l23_underPF)
'Dim BL_IN_l24: BL_IN_l24=Array(LM_IN_l24_Layer4, LM_IN_l24_Layer7, LM_IN_l24_Playfield, LM_IN_l24_parts, LM_IN_l24_underPF)
'Dim BL_IN_l25: BL_IN_l25=Array(LM_IN_l25_Layer4, LM_IN_l25_Playfield, LM_IN_l25_parts, LM_IN_l25_underPF)
'Dim BL_IN_l26: BL_IN_l26=Array(LM_IN_l26_Playfield, LM_IN_l26_parts, LM_IN_l26_underPF)
'Dim BL_IN_l27: BL_IN_l27=Array(LM_IN_l27_Layer4, LM_IN_l27_Playfield, LM_IN_l27_parts, LM_IN_l27_underPF)
'Dim BL_IN_l28: BL_IN_l28=Array(LM_IN_l28_parts, LM_IN_l28_underPF)
'Dim BL_IN_l31: BL_IN_l31=Array(LM_IN_l31_Layer4, LM_IN_l31_Layer6, LM_IN_l31_Layer7, LM_IN_l31_Playfield, LM_IN_l31_goalkeeper, LM_IN_l31_parts, LM_IN_l31_soccerBall_001, LM_IN_l31_soccerBall_002, LM_IN_l31_sw25, LM_IN_l31_underPF)
'Dim BL_IN_l32: BL_IN_l32=Array(LM_IN_l32_Layer4, LM_IN_l32_Layer6, LM_IN_l32_Layer7, LM_IN_l32_Playfield, LM_IN_l32_parts, LM_IN_l32_soccerBall_001, LM_IN_l32_sw25, LM_IN_l32_underPF)
'Dim BL_IN_l33: BL_IN_l33=Array(LM_IN_l33_Layer7, LM_IN_l33_parts, LM_IN_l33_underPF)
'Dim BL_IN_l34: BL_IN_l34=Array(LM_IN_l34_Layer7, LM_IN_l34_parts, LM_IN_l34_underPF)
'Dim BL_IN_l35: BL_IN_l35=Array(LM_IN_l35_Layer7, LM_IN_l35_Playfield, LM_IN_l35_parts, LM_IN_l35_underPF)
'Dim BL_IN_l36: BL_IN_l36=Array(LM_IN_l36_Layer7, LM_IN_l36_Playfield, LM_IN_l36_parts, LM_IN_l36_underPF)
'Dim BL_IN_l37: BL_IN_l37=Array(LM_IN_l37_parts, LM_IN_l37_underPF)
'Dim BL_IN_l38: BL_IN_l38=Array(LM_IN_l38_parts, LM_IN_l38_underPF)
'Dim BL_IN_l41: BL_IN_l41=Array(LM_IN_l41_Layer4, LM_IN_l41_Playfield, LM_IN_l41_parts, LM_IN_l41_underPF)
'Dim BL_IN_l42: BL_IN_l42=Array(LM_IN_l42_Layer1, LM_IN_l42_Layer4, LM_IN_l42_Playfield, LM_IN_l42_parts, LM_IN_l42_underPF)
'Dim BL_IN_l43: BL_IN_l43=Array(LM_IN_l43_Layer1, LM_IN_l43_Layer4, LM_IN_l43_Layer6, LM_IN_l43_parts, LM_IN_l43_underPF)
'Dim BL_IN_l44: BL_IN_l44=Array(LM_IN_l44_Layer3, LM_IN_l44_Layer4, LM_IN_l44_Layer5, LM_IN_l44_Layer7, LM_IN_l44_Playfield, LM_IN_l44_parts, LM_IN_l44_underPF)
'Dim BL_IN_l45: BL_IN_l45=Array(LM_IN_l45_Layer3, LM_IN_l45_Layer4, LM_IN_l45_Layer7, LM_IN_l45_Playfield, LM_IN_l45_parts, LM_IN_l45_underPF)
'Dim BL_IN_l46: BL_IN_l46=Array(LM_IN_l46_Layer1, LM_IN_l46_Layer4, LM_IN_l46_Layer7, LM_IN_l46_Playfield, LM_IN_l46_parts, LM_IN_l46_sw27, LM_IN_l46_underPF)
'Dim BL_IN_l47: BL_IN_l47=Array(LM_IN_l47_BR1, LM_IN_l47_BR2, LM_IN_l47_BS1, LM_IN_l47_Layer0, LM_IN_l47_Layer1, LM_IN_l47_Layer2, LM_IN_l47_Layer3, LM_IN_l47_Layer4, LM_IN_l47_Layer5, LM_IN_l47_Layer7, LM_IN_l47_Playfield, LM_IN_l47_parts, LM_IN_l47_underPF)
'Dim BL_IN_l48: BL_IN_l48=Array(LM_IN_l48_Layer1, LM_IN_l48_Layer4, LM_IN_l48_Layer5, LM_IN_l48_parts, LM_IN_l48_underPF)
'Dim BL_IN_l51: BL_IN_l51=Array(LM_IN_l51_Layer3, LM_IN_l51_Layer5, LM_IN_l51_Playfield, LM_IN_l51_goalkeeper, LM_IN_l51_parts, LM_IN_l51_underPF)
'Dim BL_IN_l52: BL_IN_l52=Array(LM_IN_l52_Layer4, LM_IN_l52_Layer5, LM_IN_l52_Playfield, LM_IN_l52_goalkeeper, LM_IN_l52_parts, LM_IN_l52_sw16, LM_IN_l52_underPF)
'Dim BL_IN_l53: BL_IN_l53=Array(LM_IN_l53_BR3, LM_IN_l53_BS3, LM_IN_l53_Layer1, LM_IN_l53_Layer2, LM_IN_l53_Layer4, LM_IN_l53_Layer5, LM_IN_l53_Playfield, LM_IN_l53_goalkeeper, LM_IN_l53_parts, LM_IN_l53_underPF)
'Dim BL_IN_l54: BL_IN_l54=Array(LM_IN_l54_BR3, LM_IN_l54_Layer1, LM_IN_l54_Layer2, LM_IN_l54_Layer5, LM_IN_l54_Playfield, LM_IN_l54_goalkeeper, LM_IN_l54_parts, LM_IN_l54_sw16, LM_IN_l54_underPF)
'Dim BL_IN_l55: BL_IN_l55=Array(LM_IN_l55_Layer7, LM_IN_l55_Layer8, LM_IN_l55_Playfield, LM_IN_l55_parts, LM_IN_l55_underPF)
'Dim BL_IN_l56: BL_IN_l56=Array(LM_IN_l56_Layer7, LM_IN_l56_Playfield, LM_IN_l56_RSling1, LM_IN_l56_parts, LM_IN_l56_underPF)
'Dim BL_IN_l57: BL_IN_l57=Array(LM_IN_l57_Layer1, LM_IN_l57_Layer7, LM_IN_l57_Playfield, LM_IN_l57_parts, LM_IN_l57_underPF)
'Dim BL_IN_l58: BL_IN_l58=Array(LM_IN_l58_Layer1, LM_IN_l58_Layer7, LM_IN_l58_Playfield, LM_IN_l58_parts, LM_IN_l58_underPF)
'Dim BL_IN_l61: BL_IN_l61=Array(LM_IN_l61_Playfield, LM_IN_l61_parts, LM_IN_l61_underPF)
'Dim BL_IN_l62: BL_IN_l62=Array(LM_IN_l62_Layer4, LM_IN_l62_Playfield, LM_IN_l62_parts, LM_IN_l62_sw27, LM_IN_l62_underPF)
'Dim BL_IN_l63: BL_IN_l63=Array(LM_IN_l63_Layer4, LM_IN_l63_Playfield, LM_IN_l63_parts, LM_IN_l63_underPF)
'Dim BL_IN_l64: BL_IN_l64=Array(LM_IN_l64_Layer1, LM_IN_l64_Layer4, LM_IN_l64_Layer6, LM_IN_l64_Playfield, LM_IN_l64_parts, LM_IN_l64_underPF)
'Dim BL_IN_l65: BL_IN_l65=Array(LM_IN_l65_Playfield, LM_IN_l65_parts, LM_IN_l65_underPF)
'Dim BL_IN_l66: BL_IN_l66=Array(LM_IN_l66_Layer4, LM_IN_l66_Playfield, LM_IN_l66_gate002, LM_IN_l66_parts, LM_IN_l66_underPF)
'Dim BL_IN_l67: BL_IN_l67=Array(LM_IN_l67_Layer2, LM_IN_l67_Layer4, LM_IN_l67_Layer5, LM_IN_l67_Playfield, LM_IN_l67_gate003, LM_IN_l67_parts, LM_IN_l67_underPF)
'Dim BL_IN_l68: BL_IN_l68=Array(LM_IN_l68_Layer1, LM_IN_l68_Layer3, LM_IN_l68_Layer4, LM_IN_l68_Layer5, LM_IN_l68_Layer7, LM_IN_l68_Layer8, LM_IN_l68_Playfield, LM_IN_l68_Rpost, LM_IN_l68_parts)
'Dim BL_IN_l71: BL_IN_l71=Array(LM_IN_l71_Layer4, LM_IN_l71_Layer5, LM_IN_l71_Layer6, LM_IN_l71_Layer7, LM_IN_l71_Playfield, LM_IN_l71_parts)
'Dim BL_IN_l72: BL_IN_l72=Array(LM_IN_l72_Layer6, LM_IN_l72_Playfield, LM_IN_l72_goalkeeper, LM_IN_l72_parts, LM_IN_l72_underPF)
'Dim BL_IN_l73: BL_IN_l73=Array(LM_IN_l73_Layer1, LM_IN_l73_Playfield, LM_IN_l73_parts, LM_IN_l73_underPF)
'Dim BL_IN_l74: BL_IN_l74=Array(LM_IN_l74_Layer1, LM_IN_l74_Layer4, LM_IN_l74_Playfield, LM_IN_l74_parts, LM_IN_l74_underPF)
'Dim BL_IN_l75: BL_IN_l75=Array(LM_IN_l75_Layer4, LM_IN_l75_Playfield, LM_IN_l75_parts, LM_IN_l75_underPF)
'Dim BL_IN_l76: BL_IN_l76=Array(LM_IN_l76_Layer1, LM_IN_l76_Layer4, LM_IN_l76_Layer5, LM_IN_l76_Layer6, LM_IN_l76_Playfield, LM_IN_l76_gate006, LM_IN_l76_goalkeeper, LM_IN_l76_parts, LM_IN_l76_soccerBall_001, LM_IN_l76_soccerBall_002, LM_IN_l76_soccerBall_003, LM_IN_l76_sw67)
'Dim BL_IN_l77: BL_IN_l77=Array(LM_IN_l77_Layer1, LM_IN_l77_Layer3, LM_IN_l77_Layer4, LM_IN_l77_Layer5, LM_IN_l77_Layer6, LM_IN_l77_Layer7, LM_IN_l77_Playfield, LM_IN_l77_parts, LM_IN_l77_soccerBall_001, LM_IN_l77_soccerBall_002, LM_IN_l77_soccerBall_003, LM_IN_l77_sw25, LM_IN_l77_underPF)
'Dim BL_IN_l78: BL_IN_l78=Array(LM_IN_l78_Layer1, LM_IN_l78_Layer4, LM_IN_l78_Layer5, LM_IN_l78_Layer6, LM_IN_l78_Layer7, LM_IN_l78_Playfield, LM_IN_l78_goalkeeper, LM_IN_l78_parts, LM_IN_l78_soccerBall_001, LM_IN_l78_soccerBall_002, LM_IN_l78_soccerBall_003, LM_IN_l78_sw25, LM_IN_l78_underPF)
'Dim BL_IN_l81: BL_IN_l81=Array(LM_IN_l81_Playfield, LM_IN_l81_goalkeeper, LM_IN_l81_parts, LM_IN_l81_sw61, LM_IN_l81_underPF)
'Dim BL_IN_l82: BL_IN_l82=Array(LM_IN_l82_parts, LM_IN_l82_soccerBall_003, LM_IN_l82_sw62)
'Dim BL_IN_l83: BL_IN_l83=Array(LM_IN_l83_parts, LM_IN_l83_sw63)
'Dim BL_IN_l84: BL_IN_l84=Array(LM_IN_l84_parts, LM_IN_l84_sw64, LM_IN_l84_underPF)
'Dim BL_IN_l85: BL_IN_l85=Array(LM_IN_l85_Layer3, LM_IN_l85_Layer5, LM_IN_l85_Layer7, LM_IN_l85_Layer8, LM_IN_l85_Playfield, LM_IN_l85_Rpost, LM_IN_l85_parts, LM_IN_l85_soccerBall_003, LM_IN_l85_sw37, LM_IN_l85_underPF)
'Dim BL_IN_l86: BL_IN_l86=Array(LM_IN_l86_Layer1, LM_IN_l86_Layer3, LM_IN_l86_Layer4, LM_IN_l86_Layer5, LM_IN_l86_Layer7, LM_IN_l86_Layer8, LM_IN_l86_Playfield, LM_IN_l86_Rpost, LM_IN_l86_parts, LM_IN_l86_sw37, LM_IN_l86_underPF)
'Dim BL_Room: BL_Room=Array(BM_BR1, BM_BR2, BM_BR3, BM_BS1, BM_BS2, BM_BS3, BM_LEMK, BM_LFlip, BM_LFlipU, BM_LSling0, BM_LSling1, BM_LSling2, BM_Layer0, BM_Layer1, BM_Layer2, BM_Layer3, BM_Layer4, BM_Layer5, BM_Layer6, BM_Layer7, BM_Layer8, BM_LockPin, BM_Lpost, BM_Playfield, BM_REMK, BM_RFlip, BM_RFlipU, BM_RSling0, BM_RSling1, BM_RSling2, BM_Rivet_1_008, BM_Rpost, BM_diverterside01, BM_diverterside02, BM_gate002, BM_gate003, BM_gate004, BM_gate006, BM_gate007, BM_goalkeeper, BM_parts, BM_plunger, BM_soccerBall_001, BM_soccerBall_002, BM_soccerBall_003, BM_sw15, BM_sw16, BM_sw17, BM_sw18, BM_sw25, BM_sw26, BM_sw27, BM_sw28, BM_sw37, BM_sw38, BM_sw47, BM_sw61, BM_sw62, BM_sw63, BM_sw64, BM_sw66, BM_sw67, BM_sw71, BM_sw74, BM_sw76, BM_sw77, BM_sw78, BM_sw86, BM_sw87, BM_sw88, BM_underPF)
'' Global arrays
'Dim BG_Bakemap: BG_Bakemap=Array(BM_BR1, BM_BR2, BM_BR3, BM_BS1, BM_BS2, BM_BS3, BM_LEMK, BM_LFlip, BM_LFlipU, BM_LSling0, BM_LSling1, BM_LSling2, BM_Layer0, BM_Layer1, BM_Layer2, BM_Layer3, BM_Layer4, BM_Layer5, BM_Layer6, BM_Layer7, BM_Layer8, BM_LockPin, BM_Lpost, BM_Playfield, BM_REMK, BM_RFlip, BM_RFlipU, BM_RSling0, BM_RSling1, BM_RSling2, BM_Rivet_1_008, BM_Rpost, BM_diverterside01, BM_diverterside02, BM_gate002, BM_gate003, BM_gate004, BM_gate006, BM_gate007, BM_goalkeeper, BM_parts, BM_plunger, BM_soccerBall_001, BM_soccerBall_002, BM_soccerBall_003, BM_sw15, BM_sw16, BM_sw17, BM_sw18, BM_sw25, BM_sw26, BM_sw27, BM_sw28, BM_sw37, BM_sw38, BM_sw47, BM_sw61, BM_sw62, BM_sw63, BM_sw64, BM_sw66, BM_sw67, BM_sw71, BM_sw74, BM_sw76, BM_sw77, BM_sw78, BM_sw86, BM_sw87, BM_sw88, BM_underPF)
'Dim BG_Lightmap: BG_Lightmap=Array(LM_FL_f17_BR2, LM_FL_f17_BS2, LM_FL_f17_Layer1, LM_FL_f17_Layer2, LM_FL_f17_Layer3, LM_FL_f17_Layer4, LM_FL_f17_Layer5, LM_FL_f17_Layer6, LM_FL_f17_Playfield, LM_FL_f17_Rivet_1_008, LM_FL_f17_diverterside01, LM_FL_f17_diverterside02, LM_FL_f17_gate002, LM_FL_f17_goalkeeper, LM_FL_f17_parts, LM_FL_f17_underPF, LM_FL_f18_Layer1, LM_FL_f18_Layer3, LM_FL_f18_Layer4, LM_FL_f18_Layer5, LM_FL_f18_Layer6, LM_FL_f18_Playfield, LM_FL_f18_diverterside01, LM_FL_f18_diverterside02, LM_FL_f18_goalkeeper, LM_FL_f18_parts, LM_FL_f18_soccerBall_001, LM_FL_f18_soccerBall_002, LM_FL_f18_soccerBall_003, LM_FL_f18_sw16, LM_FL_f18_underPF, LM_FL_f19_Layer1, LM_FL_f19_Layer3, LM_FL_f19_Layer4, LM_FL_f19_Layer5, LM_FL_f19_Layer6, LM_FL_f19_Layer7, LM_FL_f19_Layer8, LM_FL_f19_Playfield, LM_FL_f19_gate004, LM_FL_f19_parts, LM_FL_f19_plunger, LM_FL_f19_soccerBall_001, LM_FL_f19_soccerBall_002, LM_FL_f19_soccerBall_003, LM_FL_f19_sw38, LM_FL_f19_sw74, LM_FL_f19_underPF, LM_FL_f20_BR1, LM_FL_f20_BR2, _
' LM_FL_f20_BR3, LM_FL_f20_BS1, LM_FL_f20_BS2, LM_FL_f20_BS3, LM_FL_f20_Layer0, LM_FL_f20_Layer1, LM_FL_f20_Layer2, LM_FL_f20_Layer4, LM_FL_f20_Layer5, LM_FL_f20_Playfield, LM_FL_f20_parts, LM_FL_f20_soccerBall_001, LM_FL_f20_soccerBall_002, LM_FL_f20_soccerBall_003, LM_FL_f20_underPF, LM_FL_f21_BR2, LM_FL_f21_BR3, LM_FL_f21_Layer0, LM_FL_f21_Layer1, LM_FL_f21_Layer2, LM_FL_f21_Layer3, LM_FL_f21_Layer4, LM_FL_f21_Layer5, LM_FL_f21_Layer6, LM_FL_f21_Playfield, LM_FL_f21_diverterside01, LM_FL_f21_diverterside02, LM_FL_f21_goalkeeper, LM_FL_f21_parts, LM_FL_f21_soccerBall_001, LM_FL_f21_soccerBall_002, LM_FL_f21_soccerBall_003, LM_FL_f21_sw61, LM_FL_f21_sw62, LM_FL_f21_underPF, LM_FL_f22_BR1, LM_FL_f22_BR3, LM_FL_f22_BS1, LM_FL_f22_BS3, LM_FL_f22_Layer0, LM_FL_f22_Layer1, LM_FL_f22_Layer2, LM_FL_f22_Layer3, LM_FL_f22_Layer4, LM_FL_f22_Layer5, LM_FL_f22_Layer6, LM_FL_f22_Layer7, LM_FL_f22_Playfield, LM_FL_f22_diverterside01, LM_FL_f22_gate004, LM_FL_f22_gate006, LM_FL_f22_gate007, LM_FL_f22_goalkeeper, _
' LM_FL_f22_parts, LM_FL_f22_soccerBall_001, LM_FL_f22_soccerBall_002, LM_FL_f22_soccerBall_003, LM_FL_f22_sw16, LM_FL_f22_sw25, LM_FL_f22_underPF, LM_FL_f25_BR1, LM_FL_f25_BR2, LM_FL_f25_BR3, LM_FL_f25_Layer0, LM_FL_f25_Layer1, LM_FL_f25_Layer2, LM_FL_f25_Layer3, LM_FL_f25_Layer4, LM_FL_f25_Layer5, LM_FL_f25_Layer6, LM_FL_f25_Layer7, LM_FL_f25_Playfield, LM_FL_f25_diverterside02, LM_FL_f25_gate004, LM_FL_f25_gate006, LM_FL_f25_goalkeeper, LM_FL_f25_parts, LM_FL_f25_soccerBall_001, LM_FL_f25_soccerBall_002, LM_FL_f25_soccerBall_003, LM_FL_f25_sw16, LM_FL_f25_sw25, LM_FL_f25_sw27, LM_FL_f25_sw47, LM_FL_f25_underPF, LM_FL_f26_Layer1, LM_FL_f26_Layer4, LM_FL_f26_Layer5, LM_FL_f26_Layer6, LM_FL_f26_Playfield, LM_FL_f26_parts, LM_FL_f26_soccerBall_001, LM_FL_f26_soccerBall_002, LM_FL_f26_soccerBall_003, LM_FL_f26_sw27, LM_FL_f26_underPF, LM_FL_f27_LFlip, LM_FL_f27_LFlipU, LM_FL_f27_Layer1, LM_FL_f27_Layer3, LM_FL_f27_Layer4, LM_FL_f27_Layer7, LM_FL_f27_Playfield, LM_FL_f27_RFlip, LM_FL_f27_Rpost, LM_FL_f27_parts, _
' LM_FL_f27_sw17, LM_FL_f27_sw18, LM_FL_f27_sw86, LM_FL_f27_underPF, LM_FL_f28_BR2, LM_FL_f28_BR3, LM_FL_f28_BS2, LM_FL_f28_Layer1, LM_FL_f28_Layer2, LM_FL_f28_Layer3, LM_FL_f28_Layer4, LM_FL_f28_Layer5, LM_FL_f28_Layer6, LM_FL_f28_Layer7, LM_FL_f28_Playfield, LM_FL_f28_Rivet_1_008, LM_FL_f28_diverterside01, LM_FL_f28_diverterside02, LM_FL_f28_goalkeeper, LM_FL_f28_parts, LM_FL_f28_soccerBall_001, LM_FL_f28_soccerBall_002, LM_FL_f28_soccerBall_003, LM_FL_f28_sw27, LM_FL_f28_underPF, LM_GILeft_GI010_LFlip, LM_GILeft_GI010_LFlipU, LM_GILeft_GI010_Playfield, LM_GILeft_GI010_RFlipU, LM_GILeft_GI010_parts, LM_GILeft_GI011_LFlip, LM_GILeft_GI011_LFlipU, LM_GILeft_GI011_Playfield, LM_GILeft_GI011_parts, LM_GILeft_GI012_LFlip, LM_GILeft_GI012_LFlipU, LM_GILeft_GI012_LSling0, LM_GILeft_GI012_LSling1, LM_GILeft_GI012_LSling2, LM_GILeft_GI012_Layer1, LM_GILeft_GI012_Layer4, LM_GILeft_GI012_Layer5, LM_GILeft_GI012_Layer7, LM_GILeft_GI012_Lpost, LM_GILeft_GI012_Playfield, LM_GILeft_GI012_RSling0, LM_GILeft_GI012_RSling1, _
' LM_GILeft_GI012_RSling2, LM_GILeft_GI012_parts, LM_GILeft_GI012_sw25, LM_GILeft_GI012_underPF, LM_GILeft_GI013_LEMK, LM_GILeft_GI013_LFlip, LM_GILeft_GI013_LFlipU, LM_GILeft_GI013_LSling0, LM_GILeft_GI013_LSling1, LM_GILeft_GI013_LSling2, LM_GILeft_GI013_Layer1, LM_GILeft_GI013_Layer4, LM_GILeft_GI013_Layer7, LM_GILeft_GI013_LockPin, LM_GILeft_GI013_Lpost, LM_GILeft_GI013_Playfield, LM_GILeft_GI013_RFlip, LM_GILeft_GI013_RFlipU, LM_GILeft_GI013_RSling1, LM_GILeft_GI013_parts, LM_GILeft_GI013_soccerBall_001, LM_GILeft_GI013_soccerBall_002, LM_GILeft_GI013_soccerBall_003, LM_GILeft_GI013_sw15, LM_GILeft_GI013_sw25, LM_GILeft_GI013_sw28, LM_GILeft_GI014_BR1, LM_GILeft_GI014_LSling1, LM_GILeft_GI014_Layer0, LM_GILeft_GI014_Layer1, LM_GILeft_GI014_Layer2, LM_GILeft_GI014_Layer3, LM_GILeft_GI014_Layer4, LM_GILeft_GI014_Layer5, LM_GILeft_GI014_Layer6, LM_GILeft_GI014_Layer7, LM_GILeft_GI014_LockPin, LM_GILeft_GI014_Lpost, LM_GILeft_GI014_Playfield, LM_GILeft_GI014_goalkeeper, LM_GILeft_GI014_parts, _
' LM_GILeft_GI014_soccerBall_001, LM_GILeft_GI014_soccerBall_002, LM_GILeft_GI014_soccerBall_003, LM_GILeft_GI014_sw15, LM_GILeft_GI014_sw25, LM_GILeft_GI014_sw26, LM_GILeft_GI014_sw27, LM_GILeft_GI014_sw28, LM_GILeft_GI014_sw37, LM_GILeft_GI014_sw67, LM_GILeft_GI014_sw77, LM_GILeft_GI014_sw78, LM_GILeft_GI014_underPF, LM_GIRight_GI001_LFlipU, LM_GIRight_GI001_Layer7, LM_GIRight_GI001_Playfield, LM_GIRight_GI001_RFlip, LM_GIRight_GI001_RFlipU, LM_GIRight_GI001_parts, LM_GIRight_GI001_underPF, LM_GIRight_GI002_Layer7, LM_GIRight_GI002_Playfield, LM_GIRight_GI002_RFlip, LM_GIRight_GI002_RFlipU, LM_GIRight_GI002_parts, LM_GIRight_GI002_sw18, LM_GIRight_GI002_underPF, LM_GIRight_GI003_LSling0, LM_GIRight_GI003_LSling1, LM_GIRight_GI003_LSling2, LM_GIRight_GI003_Layer1, LM_GIRight_GI003_Layer3, LM_GIRight_GI003_Layer4, LM_GIRight_GI003_Layer7, LM_GIRight_GI003_Playfield, LM_GIRight_GI003_REMK, LM_GIRight_GI003_RFlip, LM_GIRight_GI003_RFlipU, LM_GIRight_GI003_RSling0, LM_GIRight_GI003_RSling1, _
' LM_GIRight_GI003_RSling2, LM_GIRight_GI003_Rpost, LM_GIRight_GI003_parts, LM_GIRight_GI003_sw17, LM_GIRight_GI003_sw37, LM_GIRight_GI003_sw76, LM_GIRight_GI003_underPF, LM_GIRight_GI004_LFlipU, LM_GIRight_GI004_Layer1, LM_GIRight_GI004_Layer3, LM_GIRight_GI004_Layer4, LM_GIRight_GI004_Layer5, LM_GIRight_GI004_Layer7, LM_GIRight_GI004_Playfield, LM_GIRight_GI004_REMK, LM_GIRight_GI004_RFlip, LM_GIRight_GI004_RFlipU, LM_GIRight_GI004_RSling0, LM_GIRight_GI004_RSling1, LM_GIRight_GI004_RSling2, LM_GIRight_GI004_Rpost, LM_GIRight_GI004_parts, LM_GIRight_GI004_sw17, LM_GIRight_GI004_sw18, LM_GIRight_GI004_sw28, LM_GIRight_GI004_sw37, LM_GIRight_GI004_underPF, LM_GIRight_GI041_BR1, LM_GIRight_GI041_BR3, LM_GIRight_GI041_BS1, LM_GIRight_GI041_Layer0, LM_GIRight_GI041_Layer1, LM_GIRight_GI041_Layer2, LM_GIRight_GI041_Layer3, LM_GIRight_GI041_Layer4, LM_GIRight_GI041_Layer5, LM_GIRight_GI041_Layer6, LM_GIRight_GI041_Layer7, LM_GIRight_GI041_Layer8, LM_GIRight_GI041_Lpost, LM_GIRight_GI041_Playfield, _
' LM_GIRight_GI041_RSling0, LM_GIRight_GI041_RSling1, LM_GIRight_GI041_RSling2, LM_GIRight_GI041_Rpost, LM_GIRight_GI041_gate004, LM_GIRight_GI041_gate007, LM_GIRight_GI041_goalkeeper, LM_GIRight_GI041_parts, LM_GIRight_GI041_soccerBall_001, LM_GIRight_GI041_soccerBall_002, LM_GIRight_GI041_soccerBall_003, LM_GIRight_GI041_sw16, LM_GIRight_GI041_sw17, LM_GIRight_GI041_sw18, LM_GIRight_GI041_sw25, LM_GIRight_GI041_sw37, LM_GIRight_GI041_underPF, LM_GITop_BR1, LM_GITop_BR2, LM_GITop_BR3, LM_GITop_BS1, LM_GITop_BS2, LM_GITop_BS3, LM_GITop_Layer0, LM_GITop_Layer1, LM_GITop_Layer2, LM_GITop_Layer3, LM_GITop_Layer4, LM_GITop_Layer5, LM_GITop_Layer6, LM_GITop_Layer7, LM_GITop_Playfield, LM_GITop_Rivet_1_008, LM_GITop_diverterside01, LM_GITop_diverterside02, LM_GITop_gate002, LM_GITop_gate003, LM_GITop_gate006, LM_GITop_gate007, LM_GITop_goalkeeper, LM_GITop_parts, LM_GITop_soccerBall_001, LM_GITop_soccerBall_002, LM_GITop_soccerBall_003, LM_GITop_sw16, LM_GITop_sw25, LM_GITop_sw27, LM_GITop_sw47, LM_GITop_sw61, _
' LM_GITop_sw62, LM_GITop_sw71, LM_GITop_sw87, LM_GITop_sw88, LM_GITop_underPF, LM_IN_l11_Playfield, LM_IN_l11_parts, LM_IN_l11_underPF, LM_IN_l12_Playfield, LM_IN_l12_parts, LM_IN_l12_underPF, LM_IN_l13_Layer4, LM_IN_l13_Playfield, LM_IN_l13_parts, LM_IN_l13_underPF, LM_IN_l14_Playfield, LM_IN_l14_parts, LM_IN_l14_underPF, LM_IN_l15_Layer4, LM_IN_l15_Playfield, LM_IN_l15_parts, LM_IN_l15_underPF, LM_IN_l16_Layer4, LM_IN_l16_Playfield, LM_IN_l16_parts, LM_IN_l16_underPF, LM_IN_l17_Layer4, LM_IN_l17_Playfield, LM_IN_l17_parts, LM_IN_l17_underPF, LM_IN_l18_Layer4, LM_IN_l18_LockPin, LM_IN_l18_Playfield, LM_IN_l18_parts, LM_IN_l18_underPF, LM_IN_l21_Layer4, LM_IN_l21_Layer7, LM_IN_l21_Playfield, LM_IN_l21_parts, LM_IN_l21_underPF, LM_IN_l22_Layer4, LM_IN_l22_Layer7, LM_IN_l22_Playfield, LM_IN_l22_parts, LM_IN_l22_underPF, LM_IN_l23_Layer4, LM_IN_l23_Layer7, LM_IN_l23_Playfield, LM_IN_l23_parts, LM_IN_l23_underPF, LM_IN_l24_Layer4, LM_IN_l24_Layer7, LM_IN_l24_Playfield, LM_IN_l24_parts, LM_IN_l24_underPF, _
' LM_IN_l25_Layer4, LM_IN_l25_Playfield, LM_IN_l25_parts, LM_IN_l25_underPF, LM_IN_l26_Playfield, LM_IN_l26_parts, LM_IN_l26_underPF, LM_IN_l27_Layer4, LM_IN_l27_Playfield, LM_IN_l27_parts, LM_IN_l27_underPF, LM_IN_l28_parts, LM_IN_l28_underPF, LM_IN_l31_Layer4, LM_IN_l31_Layer6, LM_IN_l31_Layer7, LM_IN_l31_Playfield, LM_IN_l31_goalkeeper, LM_IN_l31_parts, LM_IN_l31_soccerBall_001, LM_IN_l31_soccerBall_002, LM_IN_l31_sw25, LM_IN_l31_underPF, LM_IN_l32_Layer4, LM_IN_l32_Layer6, LM_IN_l32_Layer7, LM_IN_l32_Playfield, LM_IN_l32_parts, LM_IN_l32_soccerBall_001, LM_IN_l32_sw25, LM_IN_l32_underPF, LM_IN_l33_Layer7, LM_IN_l33_parts, LM_IN_l33_underPF, LM_IN_l34_Layer7, LM_IN_l34_parts, LM_IN_l34_underPF, LM_IN_l35_Layer7, LM_IN_l35_Playfield, LM_IN_l35_parts, LM_IN_l35_underPF, LM_IN_l36_Layer7, LM_IN_l36_Playfield, LM_IN_l36_parts, LM_IN_l36_underPF, LM_IN_l37_parts, LM_IN_l37_underPF, LM_IN_l38_parts, LM_IN_l38_underPF, LM_IN_l41_Layer4, LM_IN_l41_Playfield, LM_IN_l41_parts, LM_IN_l41_underPF, LM_IN_l42_Layer1, _
' LM_IN_l42_Layer4, LM_IN_l42_Playfield, LM_IN_l42_parts, LM_IN_l42_underPF, LM_IN_l43_Layer1, LM_IN_l43_Layer4, LM_IN_l43_Layer6, LM_IN_l43_parts, LM_IN_l43_underPF, LM_IN_l44_Layer3, LM_IN_l44_Layer4, LM_IN_l44_Layer5, LM_IN_l44_Layer7, LM_IN_l44_Playfield, LM_IN_l44_parts, LM_IN_l44_underPF, LM_IN_l45_Layer3, LM_IN_l45_Layer4, LM_IN_l45_Layer7, LM_IN_l45_Playfield, LM_IN_l45_parts, LM_IN_l45_underPF, LM_IN_l46_Layer1, LM_IN_l46_Layer4, LM_IN_l46_Layer7, LM_IN_l46_Playfield, LM_IN_l46_parts, LM_IN_l46_sw27, LM_IN_l46_underPF, LM_IN_l47_BR1, LM_IN_l47_BR2, LM_IN_l47_BS1, LM_IN_l47_Layer0, LM_IN_l47_Layer1, LM_IN_l47_Layer2, LM_IN_l47_Layer3, LM_IN_l47_Layer4, LM_IN_l47_Layer5, LM_IN_l47_Layer7, LM_IN_l47_Playfield, LM_IN_l47_parts, LM_IN_l47_underPF, LM_IN_l48_Layer1, LM_IN_l48_Layer4, LM_IN_l48_Layer5, LM_IN_l48_parts, LM_IN_l48_underPF, LM_IN_l51_Layer3, LM_IN_l51_Layer5, LM_IN_l51_Playfield, LM_IN_l51_goalkeeper, LM_IN_l51_parts, LM_IN_l51_underPF, LM_IN_l52_Layer4, LM_IN_l52_Layer5, LM_IN_l52_Playfield, _
' LM_IN_l52_goalkeeper, LM_IN_l52_parts, LM_IN_l52_sw16, LM_IN_l52_underPF, LM_IN_l53_BR3, LM_IN_l53_BS3, LM_IN_l53_Layer1, LM_IN_l53_Layer2, LM_IN_l53_Layer4, LM_IN_l53_Layer5, LM_IN_l53_Playfield, LM_IN_l53_goalkeeper, LM_IN_l53_parts, LM_IN_l53_underPF, LM_IN_l54_BR3, LM_IN_l54_Layer1, LM_IN_l54_Layer2, LM_IN_l54_Layer5, LM_IN_l54_Playfield, LM_IN_l54_goalkeeper, LM_IN_l54_parts, LM_IN_l54_sw16, LM_IN_l54_underPF, LM_IN_l55_Layer7, LM_IN_l55_Layer8, LM_IN_l55_Playfield, LM_IN_l55_parts, LM_IN_l55_underPF, LM_IN_l56_Layer7, LM_IN_l56_Playfield, LM_IN_l56_RSling1, LM_IN_l56_parts, LM_IN_l56_underPF, LM_IN_l57_Layer1, LM_IN_l57_Layer7, LM_IN_l57_Playfield, LM_IN_l57_parts, LM_IN_l57_underPF, LM_IN_l58_Layer1, LM_IN_l58_Layer7, LM_IN_l58_Playfield, LM_IN_l58_parts, LM_IN_l58_underPF, LM_IN_l61_Playfield, LM_IN_l61_parts, LM_IN_l61_underPF, LM_IN_l62_Layer4, LM_IN_l62_Playfield, LM_IN_l62_parts, LM_IN_l62_sw27, LM_IN_l62_underPF, LM_IN_l63_Layer4, LM_IN_l63_Playfield, LM_IN_l63_parts, LM_IN_l63_underPF, _
' LM_IN_l64_Layer1, LM_IN_l64_Layer4, LM_IN_l64_Layer6, LM_IN_l64_Playfield, LM_IN_l64_parts, LM_IN_l64_underPF, LM_IN_l65_Playfield, LM_IN_l65_parts, LM_IN_l65_underPF, LM_IN_l66_Layer4, LM_IN_l66_Playfield, LM_IN_l66_gate002, LM_IN_l66_parts, LM_IN_l66_underPF, LM_IN_l67_Layer2, LM_IN_l67_Layer4, LM_IN_l67_Layer5, LM_IN_l67_Playfield, LM_IN_l67_gate003, LM_IN_l67_parts, LM_IN_l67_underPF, LM_IN_l68_Layer1, LM_IN_l68_Layer3, LM_IN_l68_Layer4, LM_IN_l68_Layer5, LM_IN_l68_Layer7, LM_IN_l68_Layer8, LM_IN_l68_Playfield, LM_IN_l68_Rpost, LM_IN_l68_parts, LM_IN_l71_Layer4, LM_IN_l71_Layer5, LM_IN_l71_Layer6, LM_IN_l71_Layer7, LM_IN_l71_Playfield, LM_IN_l71_parts, LM_IN_l72_Layer6, LM_IN_l72_Playfield, LM_IN_l72_goalkeeper, LM_IN_l72_parts, LM_IN_l72_underPF, LM_IN_l73_Layer1, LM_IN_l73_Playfield, LM_IN_l73_parts, LM_IN_l73_underPF, LM_IN_l74_Layer1, LM_IN_l74_Layer4, LM_IN_l74_Playfield, LM_IN_l74_parts, LM_IN_l74_underPF, LM_IN_l75_Layer4, LM_IN_l75_Playfield, LM_IN_l75_parts, LM_IN_l75_underPF, LM_IN_l76_Layer1, _
' LM_IN_l76_Layer4, LM_IN_l76_Layer5, LM_IN_l76_Layer6, LM_IN_l76_Playfield, LM_IN_l76_gate006, LM_IN_l76_goalkeeper, LM_IN_l76_parts, LM_IN_l76_soccerBall_001, LM_IN_l76_soccerBall_002, LM_IN_l76_soccerBall_003, LM_IN_l76_sw67, LM_IN_l77_Layer1, LM_IN_l77_Layer3, LM_IN_l77_Layer4, LM_IN_l77_Layer5, LM_IN_l77_Layer6, LM_IN_l77_Layer7, LM_IN_l77_Playfield, LM_IN_l77_parts, LM_IN_l77_soccerBall_001, LM_IN_l77_soccerBall_002, LM_IN_l77_soccerBall_003, LM_IN_l77_sw25, LM_IN_l77_underPF, LM_IN_l78_Layer1, LM_IN_l78_Layer4, LM_IN_l78_Layer5, LM_IN_l78_Layer6, LM_IN_l78_Layer7, LM_IN_l78_Playfield, LM_IN_l78_goalkeeper, LM_IN_l78_parts, LM_IN_l78_soccerBall_001, LM_IN_l78_soccerBall_002, LM_IN_l78_soccerBall_003, LM_IN_l78_sw25, LM_IN_l78_underPF, LM_IN_l81_Playfield, LM_IN_l81_goalkeeper, LM_IN_l81_parts, LM_IN_l81_sw61, LM_IN_l81_underPF, LM_IN_l82_parts, LM_IN_l82_soccerBall_003, LM_IN_l82_sw62, LM_IN_l83_parts, LM_IN_l83_sw63, LM_IN_l84_parts, LM_IN_l84_sw64, LM_IN_l84_underPF, LM_IN_l85_Layer3, LM_IN_l85_Layer5, _
' LM_IN_l85_Layer7, LM_IN_l85_Layer8, LM_IN_l85_Playfield, LM_IN_l85_Rpost, LM_IN_l85_parts, LM_IN_l85_soccerBall_003, LM_IN_l85_sw37, LM_IN_l85_underPF, LM_IN_l86_Layer1, LM_IN_l86_Layer3, LM_IN_l86_Layer4, LM_IN_l86_Layer5, LM_IN_l86_Layer7, LM_IN_l86_Layer8, LM_IN_l86_Playfield, LM_IN_l86_Rpost, LM_IN_l86_parts, LM_IN_l86_sw37, LM_IN_l86_underPF)
'Dim BG_All: BG_All=Array(BM_BR1, BM_BR2, BM_BR3, BM_BS1, BM_BS2, BM_BS3, BM_LEMK, BM_LFlip, BM_LFlipU, BM_LSling0, BM_LSling1, BM_LSling2, BM_Layer0, BM_Layer1, BM_Layer2, BM_Layer3, BM_Layer4, BM_Layer5, BM_Layer6, BM_Layer7, BM_Layer8, BM_LockPin, BM_Lpost, BM_Playfield, BM_REMK, BM_RFlip, BM_RFlipU, BM_RSling0, BM_RSling1, BM_RSling2, BM_Rivet_1_008, BM_Rpost, BM_diverterside01, BM_diverterside02, BM_gate002, BM_gate003, BM_gate004, BM_gate006, BM_gate007, BM_goalkeeper, BM_parts, BM_plunger, BM_soccerBall_001, BM_soccerBall_002, BM_soccerBall_003, BM_sw15, BM_sw16, BM_sw17, BM_sw18, BM_sw25, BM_sw26, BM_sw27, BM_sw28, BM_sw37, BM_sw38, BM_sw47, BM_sw61, BM_sw62, BM_sw63, BM_sw64, BM_sw66, BM_sw67, BM_sw71, BM_sw74, BM_sw76, BM_sw77, BM_sw78, BM_sw86, BM_sw87, BM_sw88, BM_underPF, LM_FL_f17_BR2, LM_FL_f17_BS2, LM_FL_f17_Layer1, LM_FL_f17_Layer2, LM_FL_f17_Layer3, LM_FL_f17_Layer4, LM_FL_f17_Layer5, LM_FL_f17_Layer6, LM_FL_f17_Playfield, LM_FL_f17_Rivet_1_008, LM_FL_f17_diverterside01, _
' LM_FL_f17_diverterside02, LM_FL_f17_gate002, LM_FL_f17_goalkeeper, LM_FL_f17_parts, LM_FL_f17_underPF, LM_FL_f18_Layer1, LM_FL_f18_Layer3, LM_FL_f18_Layer4, LM_FL_f18_Layer5, LM_FL_f18_Layer6, LM_FL_f18_Playfield, LM_FL_f18_diverterside01, LM_FL_f18_diverterside02, LM_FL_f18_goalkeeper, LM_FL_f18_parts, LM_FL_f18_soccerBall_001, LM_FL_f18_soccerBall_002, LM_FL_f18_soccerBall_003, LM_FL_f18_sw16, LM_FL_f18_underPF, LM_FL_f19_Layer1, LM_FL_f19_Layer3, LM_FL_f19_Layer4, LM_FL_f19_Layer5, LM_FL_f19_Layer6, LM_FL_f19_Layer7, LM_FL_f19_Layer8, LM_FL_f19_Playfield, LM_FL_f19_gate004, LM_FL_f19_parts, LM_FL_f19_plunger, LM_FL_f19_soccerBall_001, LM_FL_f19_soccerBall_002, LM_FL_f19_soccerBall_003, LM_FL_f19_sw38, LM_FL_f19_sw74, LM_FL_f19_underPF, LM_FL_f20_BR1, LM_FL_f20_BR2, LM_FL_f20_BR3, LM_FL_f20_BS1, LM_FL_f20_BS2, LM_FL_f20_BS3, LM_FL_f20_Layer0, LM_FL_f20_Layer1, LM_FL_f20_Layer2, LM_FL_f20_Layer4, LM_FL_f20_Layer5, LM_FL_f20_Playfield, LM_FL_f20_parts, LM_FL_f20_soccerBall_001, LM_FL_f20_soccerBall_002, _
' LM_FL_f20_soccerBall_003, LM_FL_f20_underPF, LM_FL_f21_BR2, LM_FL_f21_BR3, LM_FL_f21_Layer0, LM_FL_f21_Layer1, LM_FL_f21_Layer2, LM_FL_f21_Layer3, LM_FL_f21_Layer4, LM_FL_f21_Layer5, LM_FL_f21_Layer6, LM_FL_f21_Playfield, LM_FL_f21_diverterside01, LM_FL_f21_diverterside02, LM_FL_f21_goalkeeper, LM_FL_f21_parts, LM_FL_f21_soccerBall_001, LM_FL_f21_soccerBall_002, LM_FL_f21_soccerBall_003, LM_FL_f21_sw61, LM_FL_f21_sw62, LM_FL_f21_underPF, LM_FL_f22_BR1, LM_FL_f22_BR3, LM_FL_f22_BS1, LM_FL_f22_BS3, LM_FL_f22_Layer0, LM_FL_f22_Layer1, LM_FL_f22_Layer2, LM_FL_f22_Layer3, LM_FL_f22_Layer4, LM_FL_f22_Layer5, LM_FL_f22_Layer6, LM_FL_f22_Layer7, LM_FL_f22_Playfield, LM_FL_f22_diverterside01, LM_FL_f22_gate004, LM_FL_f22_gate006, LM_FL_f22_gate007, LM_FL_f22_goalkeeper, LM_FL_f22_parts, LM_FL_f22_soccerBall_001, LM_FL_f22_soccerBall_002, LM_FL_f22_soccerBall_003, LM_FL_f22_sw16, LM_FL_f22_sw25, LM_FL_f22_underPF, LM_FL_f25_BR1, LM_FL_f25_BR2, LM_FL_f25_BR3, LM_FL_f25_Layer0, LM_FL_f25_Layer1, LM_FL_f25_Layer2, _
' LM_FL_f25_Layer3, LM_FL_f25_Layer4, LM_FL_f25_Layer5, LM_FL_f25_Layer6, LM_FL_f25_Layer7, LM_FL_f25_Playfield, LM_FL_f25_diverterside02, LM_FL_f25_gate004, LM_FL_f25_gate006, LM_FL_f25_goalkeeper, LM_FL_f25_parts, LM_FL_f25_soccerBall_001, LM_FL_f25_soccerBall_002, LM_FL_f25_soccerBall_003, LM_FL_f25_sw16, LM_FL_f25_sw25, LM_FL_f25_sw27, LM_FL_f25_sw47, LM_FL_f25_underPF, LM_FL_f26_Layer1, LM_FL_f26_Layer4, LM_FL_f26_Layer5, LM_FL_f26_Layer6, LM_FL_f26_Playfield, LM_FL_f26_parts, LM_FL_f26_soccerBall_001, LM_FL_f26_soccerBall_002, LM_FL_f26_soccerBall_003, LM_FL_f26_sw27, LM_FL_f26_underPF, LM_FL_f27_LFlip, LM_FL_f27_LFlipU, LM_FL_f27_Layer1, LM_FL_f27_Layer3, LM_FL_f27_Layer4, LM_FL_f27_Layer7, LM_FL_f27_Playfield, LM_FL_f27_RFlip, LM_FL_f27_Rpost, LM_FL_f27_parts, LM_FL_f27_sw17, LM_FL_f27_sw18, LM_FL_f27_sw86, LM_FL_f27_underPF, LM_FL_f28_BR2, LM_FL_f28_BR3, LM_FL_f28_BS2, LM_FL_f28_Layer1, LM_FL_f28_Layer2, LM_FL_f28_Layer3, LM_FL_f28_Layer4, LM_FL_f28_Layer5, LM_FL_f28_Layer6, LM_FL_f28_Layer7, _
' LM_FL_f28_Playfield, LM_FL_f28_Rivet_1_008, LM_FL_f28_diverterside01, LM_FL_f28_diverterside02, LM_FL_f28_goalkeeper, LM_FL_f28_parts, LM_FL_f28_soccerBall_001, LM_FL_f28_soccerBall_002, LM_FL_f28_soccerBall_003, LM_FL_f28_sw27, LM_FL_f28_underPF, LM_GILeft_GI010_LFlip, LM_GILeft_GI010_LFlipU, LM_GILeft_GI010_Playfield, LM_GILeft_GI010_RFlipU, LM_GILeft_GI010_parts, LM_GILeft_GI011_LFlip, LM_GILeft_GI011_LFlipU, LM_GILeft_GI011_Playfield, LM_GILeft_GI011_parts, LM_GILeft_GI012_LFlip, LM_GILeft_GI012_LFlipU, LM_GILeft_GI012_LSling0, LM_GILeft_GI012_LSling1, LM_GILeft_GI012_LSling2, LM_GILeft_GI012_Layer1, LM_GILeft_GI012_Layer4, LM_GILeft_GI012_Layer5, LM_GILeft_GI012_Layer7, LM_GILeft_GI012_Lpost, LM_GILeft_GI012_Playfield, LM_GILeft_GI012_RSling0, LM_GILeft_GI012_RSling1, LM_GILeft_GI012_RSling2, LM_GILeft_GI012_parts, LM_GILeft_GI012_sw25, LM_GILeft_GI012_underPF, LM_GILeft_GI013_LEMK, LM_GILeft_GI013_LFlip, LM_GILeft_GI013_LFlipU, LM_GILeft_GI013_LSling0, LM_GILeft_GI013_LSling1, LM_GILeft_GI013_LSling2, _
' LM_GILeft_GI013_Layer1, LM_GILeft_GI013_Layer4, LM_GILeft_GI013_Layer7, LM_GILeft_GI013_LockPin, LM_GILeft_GI013_Lpost, LM_GILeft_GI013_Playfield, LM_GILeft_GI013_RFlip, LM_GILeft_GI013_RFlipU, LM_GILeft_GI013_RSling1, LM_GILeft_GI013_parts, LM_GILeft_GI013_soccerBall_001, LM_GILeft_GI013_soccerBall_002, LM_GILeft_GI013_soccerBall_003, LM_GILeft_GI013_sw15, LM_GILeft_GI013_sw25, LM_GILeft_GI013_sw28, LM_GILeft_GI014_BR1, LM_GILeft_GI014_LSling1, LM_GILeft_GI014_Layer0, LM_GILeft_GI014_Layer1, LM_GILeft_GI014_Layer2, LM_GILeft_GI014_Layer3, LM_GILeft_GI014_Layer4, LM_GILeft_GI014_Layer5, LM_GILeft_GI014_Layer6, LM_GILeft_GI014_Layer7, LM_GILeft_GI014_LockPin, LM_GILeft_GI014_Lpost, LM_GILeft_GI014_Playfield, LM_GILeft_GI014_goalkeeper, LM_GILeft_GI014_parts, LM_GILeft_GI014_soccerBall_001, LM_GILeft_GI014_soccerBall_002, LM_GILeft_GI014_soccerBall_003, LM_GILeft_GI014_sw15, LM_GILeft_GI014_sw25, LM_GILeft_GI014_sw26, LM_GILeft_GI014_sw27, LM_GILeft_GI014_sw28, LM_GILeft_GI014_sw37, LM_GILeft_GI014_sw67, _
' LM_GILeft_GI014_sw77, LM_GILeft_GI014_sw78, LM_GILeft_GI014_underPF, LM_GIRight_GI001_LFlipU, LM_GIRight_GI001_Layer7, LM_GIRight_GI001_Playfield, LM_GIRight_GI001_RFlip, LM_GIRight_GI001_RFlipU, LM_GIRight_GI001_parts, LM_GIRight_GI001_underPF, LM_GIRight_GI002_Layer7, LM_GIRight_GI002_Playfield, LM_GIRight_GI002_RFlip, LM_GIRight_GI002_RFlipU, LM_GIRight_GI002_parts, LM_GIRight_GI002_sw18, LM_GIRight_GI002_underPF, LM_GIRight_GI003_LSling0, LM_GIRight_GI003_LSling1, LM_GIRight_GI003_LSling2, LM_GIRight_GI003_Layer1, LM_GIRight_GI003_Layer3, LM_GIRight_GI003_Layer4, LM_GIRight_GI003_Layer7, LM_GIRight_GI003_Playfield, LM_GIRight_GI003_REMK, LM_GIRight_GI003_RFlip, LM_GIRight_GI003_RFlipU, LM_GIRight_GI003_RSling0, LM_GIRight_GI003_RSling1, LM_GIRight_GI003_RSling2, LM_GIRight_GI003_Rpost, LM_GIRight_GI003_parts, LM_GIRight_GI003_sw17, LM_GIRight_GI003_sw37, LM_GIRight_GI003_sw76, LM_GIRight_GI003_underPF, LM_GIRight_GI004_LFlipU, LM_GIRight_GI004_Layer1, LM_GIRight_GI004_Layer3, LM_GIRight_GI004_Layer4, _
' LM_GIRight_GI004_Layer5, LM_GIRight_GI004_Layer7, LM_GIRight_GI004_Playfield, LM_GIRight_GI004_REMK, LM_GIRight_GI004_RFlip, LM_GIRight_GI004_RFlipU, LM_GIRight_GI004_RSling0, LM_GIRight_GI004_RSling1, LM_GIRight_GI004_RSling2, LM_GIRight_GI004_Rpost, LM_GIRight_GI004_parts, LM_GIRight_GI004_sw17, LM_GIRight_GI004_sw18, LM_GIRight_GI004_sw28, LM_GIRight_GI004_sw37, LM_GIRight_GI004_underPF, LM_GIRight_GI041_BR1, LM_GIRight_GI041_BR3, LM_GIRight_GI041_BS1, LM_GIRight_GI041_Layer0, LM_GIRight_GI041_Layer1, LM_GIRight_GI041_Layer2, LM_GIRight_GI041_Layer3, LM_GIRight_GI041_Layer4, LM_GIRight_GI041_Layer5, LM_GIRight_GI041_Layer6, LM_GIRight_GI041_Layer7, LM_GIRight_GI041_Layer8, LM_GIRight_GI041_Lpost, LM_GIRight_GI041_Playfield, LM_GIRight_GI041_RSling0, LM_GIRight_GI041_RSling1, LM_GIRight_GI041_RSling2, LM_GIRight_GI041_Rpost, LM_GIRight_GI041_gate004, LM_GIRight_GI041_gate007, LM_GIRight_GI041_goalkeeper, LM_GIRight_GI041_parts, LM_GIRight_GI041_soccerBall_001, LM_GIRight_GI041_soccerBall_002, _
' LM_GIRight_GI041_soccerBall_003, LM_GIRight_GI041_sw16, LM_GIRight_GI041_sw17, LM_GIRight_GI041_sw18, LM_GIRight_GI041_sw25, LM_GIRight_GI041_sw37, LM_GIRight_GI041_underPF, LM_GITop_BR1, LM_GITop_BR2, LM_GITop_BR3, LM_GITop_BS1, LM_GITop_BS2, LM_GITop_BS3, LM_GITop_Layer0, LM_GITop_Layer1, LM_GITop_Layer2, LM_GITop_Layer3, LM_GITop_Layer4, LM_GITop_Layer5, LM_GITop_Layer6, LM_GITop_Layer7, LM_GITop_Playfield, LM_GITop_Rivet_1_008, LM_GITop_diverterside01, LM_GITop_diverterside02, LM_GITop_gate002, LM_GITop_gate003, LM_GITop_gate006, LM_GITop_gate007, LM_GITop_goalkeeper, LM_GITop_parts, LM_GITop_soccerBall_001, LM_GITop_soccerBall_002, LM_GITop_soccerBall_003, LM_GITop_sw16, LM_GITop_sw25, LM_GITop_sw27, LM_GITop_sw47, LM_GITop_sw61, LM_GITop_sw62, LM_GITop_sw71, LM_GITop_sw87, LM_GITop_sw88, LM_GITop_underPF, LM_IN_l11_Playfield, LM_IN_l11_parts, LM_IN_l11_underPF, LM_IN_l12_Playfield, LM_IN_l12_parts, LM_IN_l12_underPF, LM_IN_l13_Layer4, LM_IN_l13_Playfield, LM_IN_l13_parts, LM_IN_l13_underPF, _
' LM_IN_l14_Playfield, LM_IN_l14_parts, LM_IN_l14_underPF, LM_IN_l15_Layer4, LM_IN_l15_Playfield, LM_IN_l15_parts, LM_IN_l15_underPF, LM_IN_l16_Layer4, LM_IN_l16_Playfield, LM_IN_l16_parts, LM_IN_l16_underPF, LM_IN_l17_Layer4, LM_IN_l17_Playfield, LM_IN_l17_parts, LM_IN_l17_underPF, LM_IN_l18_Layer4, LM_IN_l18_LockPin, LM_IN_l18_Playfield, LM_IN_l18_parts, LM_IN_l18_underPF, LM_IN_l21_Layer4, LM_IN_l21_Layer7, LM_IN_l21_Playfield, LM_IN_l21_parts, LM_IN_l21_underPF, LM_IN_l22_Layer4, LM_IN_l22_Layer7, LM_IN_l22_Playfield, LM_IN_l22_parts, LM_IN_l22_underPF, LM_IN_l23_Layer4, LM_IN_l23_Layer7, LM_IN_l23_Playfield, LM_IN_l23_parts, LM_IN_l23_underPF, LM_IN_l24_Layer4, LM_IN_l24_Layer7, LM_IN_l24_Playfield, LM_IN_l24_parts, LM_IN_l24_underPF, LM_IN_l25_Layer4, LM_IN_l25_Playfield, LM_IN_l25_parts, LM_IN_l25_underPF, LM_IN_l26_Playfield, LM_IN_l26_parts, LM_IN_l26_underPF, LM_IN_l27_Layer4, LM_IN_l27_Playfield, LM_IN_l27_parts, LM_IN_l27_underPF, LM_IN_l28_parts, LM_IN_l28_underPF, LM_IN_l31_Layer4, _
' LM_IN_l31_Layer6, LM_IN_l31_Layer7, LM_IN_l31_Playfield, LM_IN_l31_goalkeeper, LM_IN_l31_parts, LM_IN_l31_soccerBall_001, LM_IN_l31_soccerBall_002, LM_IN_l31_sw25, LM_IN_l31_underPF, LM_IN_l32_Layer4, LM_IN_l32_Layer6, LM_IN_l32_Layer7, LM_IN_l32_Playfield, LM_IN_l32_parts, LM_IN_l32_soccerBall_001, LM_IN_l32_sw25, LM_IN_l32_underPF, LM_IN_l33_Layer7, LM_IN_l33_parts, LM_IN_l33_underPF, LM_IN_l34_Layer7, LM_IN_l34_parts, LM_IN_l34_underPF, LM_IN_l35_Layer7, LM_IN_l35_Playfield, LM_IN_l35_parts, LM_IN_l35_underPF, LM_IN_l36_Layer7, LM_IN_l36_Playfield, LM_IN_l36_parts, LM_IN_l36_underPF, LM_IN_l37_parts, LM_IN_l37_underPF, LM_IN_l38_parts, LM_IN_l38_underPF, LM_IN_l41_Layer4, LM_IN_l41_Playfield, LM_IN_l41_parts, LM_IN_l41_underPF, LM_IN_l42_Layer1, LM_IN_l42_Layer4, LM_IN_l42_Playfield, LM_IN_l42_parts, LM_IN_l42_underPF, LM_IN_l43_Layer1, LM_IN_l43_Layer4, LM_IN_l43_Layer6, LM_IN_l43_parts, LM_IN_l43_underPF, LM_IN_l44_Layer3, LM_IN_l44_Layer4, LM_IN_l44_Layer5, LM_IN_l44_Layer7, LM_IN_l44_Playfield, _
' LM_IN_l44_parts, LM_IN_l44_underPF, LM_IN_l45_Layer3, LM_IN_l45_Layer4, LM_IN_l45_Layer7, LM_IN_l45_Playfield, LM_IN_l45_parts, LM_IN_l45_underPF, LM_IN_l46_Layer1, LM_IN_l46_Layer4, LM_IN_l46_Layer7, LM_IN_l46_Playfield, LM_IN_l46_parts, LM_IN_l46_sw27, LM_IN_l46_underPF, LM_IN_l47_BR1, LM_IN_l47_BR2, LM_IN_l47_BS1, LM_IN_l47_Layer0, LM_IN_l47_Layer1, LM_IN_l47_Layer2, LM_IN_l47_Layer3, LM_IN_l47_Layer4, LM_IN_l47_Layer5, LM_IN_l47_Layer7, LM_IN_l47_Playfield, LM_IN_l47_parts, LM_IN_l47_underPF, LM_IN_l48_Layer1, LM_IN_l48_Layer4, LM_IN_l48_Layer5, LM_IN_l48_parts, LM_IN_l48_underPF, LM_IN_l51_Layer3, LM_IN_l51_Layer5, LM_IN_l51_Playfield, LM_IN_l51_goalkeeper, LM_IN_l51_parts, LM_IN_l51_underPF, LM_IN_l52_Layer4, LM_IN_l52_Layer5, LM_IN_l52_Playfield, LM_IN_l52_goalkeeper, LM_IN_l52_parts, LM_IN_l52_sw16, LM_IN_l52_underPF, LM_IN_l53_BR3, LM_IN_l53_BS3, LM_IN_l53_Layer1, LM_IN_l53_Layer2, LM_IN_l53_Layer4, LM_IN_l53_Layer5, LM_IN_l53_Playfield, LM_IN_l53_goalkeeper, LM_IN_l53_parts, LM_IN_l53_underPF, _
' LM_IN_l54_BR3, LM_IN_l54_Layer1, LM_IN_l54_Layer2, LM_IN_l54_Layer5, LM_IN_l54_Playfield, LM_IN_l54_goalkeeper, LM_IN_l54_parts, LM_IN_l54_sw16, LM_IN_l54_underPF, LM_IN_l55_Layer7, LM_IN_l55_Layer8, LM_IN_l55_Playfield, LM_IN_l55_parts, LM_IN_l55_underPF, LM_IN_l56_Layer7, LM_IN_l56_Playfield, LM_IN_l56_RSling1, LM_IN_l56_parts, LM_IN_l56_underPF, LM_IN_l57_Layer1, LM_IN_l57_Layer7, LM_IN_l57_Playfield, LM_IN_l57_parts, LM_IN_l57_underPF, LM_IN_l58_Layer1, LM_IN_l58_Layer7, LM_IN_l58_Playfield, LM_IN_l58_parts, LM_IN_l58_underPF, LM_IN_l61_Playfield, LM_IN_l61_parts, LM_IN_l61_underPF, LM_IN_l62_Layer4, LM_IN_l62_Playfield, LM_IN_l62_parts, LM_IN_l62_sw27, LM_IN_l62_underPF, LM_IN_l63_Layer4, LM_IN_l63_Playfield, LM_IN_l63_parts, LM_IN_l63_underPF, LM_IN_l64_Layer1, LM_IN_l64_Layer4, LM_IN_l64_Layer6, LM_IN_l64_Playfield, LM_IN_l64_parts, LM_IN_l64_underPF, LM_IN_l65_Playfield, LM_IN_l65_parts, LM_IN_l65_underPF, LM_IN_l66_Layer4, LM_IN_l66_Playfield, LM_IN_l66_gate002, LM_IN_l66_parts, LM_IN_l66_underPF, _
' LM_IN_l67_Layer2, LM_IN_l67_Layer4, LM_IN_l67_Layer5, LM_IN_l67_Playfield, LM_IN_l67_gate003, LM_IN_l67_parts, LM_IN_l67_underPF, LM_IN_l68_Layer1, LM_IN_l68_Layer3, LM_IN_l68_Layer4, LM_IN_l68_Layer5, LM_IN_l68_Layer7, LM_IN_l68_Layer8, LM_IN_l68_Playfield, LM_IN_l68_Rpost, LM_IN_l68_parts, LM_IN_l71_Layer4, LM_IN_l71_Layer5, LM_IN_l71_Layer6, LM_IN_l71_Layer7, LM_IN_l71_Playfield, LM_IN_l71_parts, LM_IN_l72_Layer6, LM_IN_l72_Playfield, LM_IN_l72_goalkeeper, LM_IN_l72_parts, LM_IN_l72_underPF, LM_IN_l73_Layer1, LM_IN_l73_Playfield, LM_IN_l73_parts, LM_IN_l73_underPF, LM_IN_l74_Layer1, LM_IN_l74_Layer4, LM_IN_l74_Playfield, LM_IN_l74_parts, LM_IN_l74_underPF, LM_IN_l75_Layer4, LM_IN_l75_Playfield, LM_IN_l75_parts, LM_IN_l75_underPF, LM_IN_l76_Layer1, LM_IN_l76_Layer4, LM_IN_l76_Layer5, LM_IN_l76_Layer6, LM_IN_l76_Playfield, LM_IN_l76_gate006, LM_IN_l76_goalkeeper, LM_IN_l76_parts, LM_IN_l76_soccerBall_001, LM_IN_l76_soccerBall_002, LM_IN_l76_soccerBall_003, LM_IN_l76_sw67, LM_IN_l77_Layer1, LM_IN_l77_Layer3, _
' LM_IN_l77_Layer4, LM_IN_l77_Layer5, LM_IN_l77_Layer6, LM_IN_l77_Layer7, LM_IN_l77_Playfield, LM_IN_l77_parts, LM_IN_l77_soccerBall_001, LM_IN_l77_soccerBall_002, LM_IN_l77_soccerBall_003, LM_IN_l77_sw25, LM_IN_l77_underPF, LM_IN_l78_Layer1, LM_IN_l78_Layer4, LM_IN_l78_Layer5, LM_IN_l78_Layer6, LM_IN_l78_Layer7, LM_IN_l78_Playfield, LM_IN_l78_goalkeeper, LM_IN_l78_parts, LM_IN_l78_soccerBall_001, LM_IN_l78_soccerBall_002, LM_IN_l78_soccerBall_003, LM_IN_l78_sw25, LM_IN_l78_underPF, LM_IN_l81_Playfield, LM_IN_l81_goalkeeper, LM_IN_l81_parts, LM_IN_l81_sw61, LM_IN_l81_underPF, LM_IN_l82_parts, LM_IN_l82_soccerBall_003, LM_IN_l82_sw62, LM_IN_l83_parts, LM_IN_l83_sw63, LM_IN_l84_parts, LM_IN_l84_sw64, LM_IN_l84_underPF, LM_IN_l85_Layer3, LM_IN_l85_Layer5, LM_IN_l85_Layer7, LM_IN_l85_Layer8, LM_IN_l85_Playfield, LM_IN_l85_Rpost, LM_IN_l85_parts, LM_IN_l85_soccerBall_003, LM_IN_l85_sw37, LM_IN_l85_underPF, LM_IN_l86_Layer1, LM_IN_l86_Layer3, LM_IN_l86_Layer4, LM_IN_l86_Layer5, LM_IN_l86_Layer7, LM_IN_l86_Layer8, _
' LM_IN_l86_Playfield, LM_IN_l86_Rpost, LM_IN_l86_parts, LM_IN_l86_sw37, LM_IN_l86_underPF)
'' VLM  Arrays - End



'******************************************************
'  ZVAR: Constants and Global Variables
'******************************************************
Const BallSize = 50   'Ball size must be 50
Const BallMass = 1    'Ball mass must be 1
Const tnob = 5      'Total number of balls on the playfield including captive balls.
Const lob = 1     'Total number of locked balls

Dim tablewidth: tablewidth = Table1.width
Dim tableheight: tableheight = Table1.height

'  Standard definitions
Const cGameName = "wcs_l2"    'PinMAME ROM name
Const UseSolenoids = 2      '1 = Normal Flippers, 2 = Fastflips
Const UseLamps = 1        '0 = Custom lamp handling, 1 = Built-in VPX handling (using light number in light timer)
Const UseSync = 0
Const HandleMech = 0
Const SSolenoidOn = ""      'Sound sample used for this, obsolete.
Const SSolenoidOff = ""     ' ^
Const SFlipperOn = ""     ' ^
Const SFlipperOff = ""      ' ^
Const SCoin = ""        ' ^

Const VRTest = False

Dim UseVPMDMD: UseVPMDMD = Table1.ShowDT Or (RenderingMode = 2) ' DMD for Desktop and VR
Dim DesktopMode: DesktopMode = Table1.ShowDT

Const UseVPMModSol = 2    'Set to 2 for PWM flashers, inserts, and GI. Requires VPinMame 3.6

'NOTES on UseVPMModSol = 2:
'  - Only supported for S9/S11/DataEast/WPC/Capcom/Whitestar (Sega & Stern)/SAM
'  - All lights on the table must have their Fader model set tp "LED (None)" to get the correct fading effects
'  - When not supported VPM outputs only 0 or 1. Therefore, use VPX "Incandescent" fader for lights

LoadVPM "03060000", "wpc.vbs", 3.46  'The "03060000" argument forces user to have VPinMame 3.6

'******************************************************
'  ZTIM: Timers
'******************************************************

'The FrameTimer interval should be -1, so executes at the display frame rate
'The frame timer should be used to update anything visual, like some animations, shadows, etc.
'However, a lot of animations will be handled in their respective _animate subroutines.

Dim FrameTime, lastgametime
lastgametime = 0

FrameTimer.Interval = -1
Sub FrameTimer_Timer()
  FrameTime = gametime - lastgametime 'Calculate FrameTime as some animuations could use this
  lastgametime = gametime 'Count frametime
  'Add animation stuff here
  BSUpdate
  UpdateBallBrightness
  RollingUpdate       'Update rolling sounds
  DoSTAnim          'Standup target animations
  UpdateStandupTargets
  UpdateSoccerBall
  If GoalieEnabled Then UpdateGoalie
End Sub



'The CorTimer interval should be 10. It's sole purpose is to update the Cor (physics) calculations
CorTimer.Interval = 10
Sub CorTimer_Timer(): Cor.Update: End Sub

' Plunger animation timers
Sub TimerPlunger_Timer
  If VRCab_Plunger.TransZ < 90 then
    VRCab_Plunger.TransZ = VRCab_Plunger.TransZ + FrameTime*0.2
    VRCab_CustomShooter.TransY = VRCab_Plunger.TransZ '+2373.47
  End If
  Dim BP
  For Each BP in BP_plunger : BP.Transy = VRCab_Plunger.TransZ : Next
End Sub
Sub TimerPlunger2_Timer
  VRCab_Plunger.TransZ = (6.0* Plunger.Position) - 20
  VRCab_CustomShooter.TransY = (6.0* Plunger.Position) - 20 '+2373.47
  Dim BP
  For Each BP in BP_plunger : BP.Transy = VRCab_Plunger.TransZ : Next
End Sub


' Hack to return Narnia ball back in play
Sub Narnia_Timer
    Dim b
  For b = 0 to UBound(gBOT)
    if gBOT(b).z < -500 Then
      'msgbox "Ball " &b& " in Narnia X: " & gBOT(b).x &" Y: "&gBOT(b).y & " Z: "&gBOT(b).z
      'debug.print "Move narnia ball ("& gBOT(b).x &" Y: "&gBOT(b).y & " Z: "&gBOT(b).z&") to upper left vuk"
      gBOT(b).x = 333
      gBOT(b).y = 250
      gBOT(b).z = -58
    end if
  next
end sub

'******************************************************
'  ZINI: Table Initialization and Exiting
'******************************************************
Dim Mag1, Mag2, PlungerIM, ttBall, mBall, mGoalie
Dim WCSBall1, WCSBall2, WCSBall3, WCSBall4, WCSBall5, gBOT

Sub Table1_Init
  vpminit me
  With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
    .SplashInfoLine = "World Cup Soccer '94 (Bally 1994)" & vbNewLine & "VPW"
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
  vpmNudge.TiltSwitch=14
  vpmNudge.Sensitivity=4
  vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingshot)

  'Main Timer init
  PinMAMETimer.Interval = PinMAMEInterval
  PinMAMETimer.Enabled = 1

  'Trough - Creates a ball in the kicker switch and gives that ball used an individual name.
  Set WCSBall1 = sw31.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set WCSBall2 = sw32.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set WCSBall3 = sw33.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set WCSBall4 = sw34.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set WCSBall5 = sw35.CreateSizedballWithMass(Ballsize/2,Ballmass)

  'Forces the trough switches to "on" at table boot so the game logic knows there are balls in the trough.
  Controller.Switch(31) = 1
  Controller.Switch(32) = 1
  Controller.Switch(33) = 1
  Controller.Switch(34) = 1
  Controller.Switch(35) = 1

  'Close coin door
  Controller.Switch(22) = 1

  '***Setting up a ball array (collection), must contain all the balls you create on the table.
  gBOT = Array(WCSBall1,WCSBall2,WCSBall3,WCSBall4,WCSBall5)

  '***Magnets
  Set mag1= New cvpmMagnet
  With mag1
    .InitMagnet MagnaGoalie, 16   'Magnet playfield object name and magnet strength.
    .GrabCenter = False     'False = Grab ball in magnet centre and stop all movement.
    .solenoid=33        'Solonoid Number
    .CreateEvents "mag1"
  End With

  Set mag2= New cvpmMagnet
  With mag2
    .InitMagnet LockMagnet, 40
    .GrabCenter = True
    .solenoid=35
    .CreateEvents "mag2"
  End With

  Set ttBall = New cvpmTurnTable
  With ttBall
    .InitTurnTable SoccerBall, 240
    .SpinUp = 0
    .SpinDown = 0
    .CreateEvents "ttBall"
  End With

  'Initialize other stuff
  KickBackPlunger.Pullback    'Pulls back and prepares the kickback kicker ready to fire.
  LStep = 0 : LeftSlingShot.TimerEnabled = 1
  RStep = 0 : RightSlingShot.TimerEnabled = 1
  diverterwall_open.collidable = 0
  diverterwall_closed.collidable = 1
  wall_sw42.collidable = False
  SetBackglass
  UpdateGoalie

  'Color adjustments
  LM_GITop_Layer2.color = RGB(0,255,255) 'tweak to bumper cap LM
  Dim BL: For each BL in BL_FL_f17: BL.color = RGB(255,0,0): Next
  For each BL in BL_FL_f19: BL.color = RGB(255,160,0): Next
  For each BL in BL_FL_f26: BL.color = RGB(255,160,0): Next
  For each BL in BL_FL_f28: BL.color = RGB(255,160,0): Next

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

'----- VR Room -----
Dim VRRoom
Dim VRRoomChoice : VRRoomChoice = 1       ' 1 - MEGA, 2 - Minimal Room, 3 - VR Cab only
Dim CabRails : CabRails = 1           ' 0=Hide, 1=Show
Dim CustomShooter : CustomShooter = 1       ' 0=Hide, 1=Show
Dim LightLevel : LightLevel = 0.25        ' Level of room lighting (0 to 1), where 0 is dark and 100 is brightest
Dim ColorLUT : ColorLUT = 1           ' Color desaturation LUTs: 1 to 11, where 1 is normal and 11 is black'n'white
Dim VolumeDial : VolumeDial = 0.8             ' Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Dim BallRollVolume : BallRollVolume = 0.5     ' Level of ball rolling volume. Value between 0 and 1
Dim RampRollVolume : RampRollVolume = 0.5     ' Level of ramp rolling volume. Value between 0 and 1


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
  Dim BP, v

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

  ' Playfield Reflections
  v = Table1.Option("Playfield Reflections", 0, 2, 1, 1, 0, Array("Off", "Clean", "Rough"))
  Select Case v
    Case 0: playfield_mesh.ReflectionProbe = "": BM_Playfield.ReflectionProbe = ""
    Case 1: playfield_mesh.ReflectionProbe = "Playfield Reflections": BM_Playfield.ReflectionProbe = "Playfield Reflections"
    Case 2: playfield_mesh.ReflectionProbe = "Playfield Reflections Rough": BM_Playfield.ReflectionProbe = "Playfield Reflections Rough"
  End Select

  ' Room brightness
' LightLevel = Table1.Option("Table Brightness (Ambient Light Level)", 0, 1, 0.01, .5, 1)
  LightLevel = NightDay/100
  SetRoomBrightness LightLevel   'Uncomment this line for lightmapped tables.

  'Outlane difficulty
  v = Table1.Option("Outlane Post Difficulty", 1, 3, 1, 2, 0, Array("Easy", "Normal", "Hard"))
  Select Case v
    Case 1
      zCol_Rubber_Lpost.x = 101 : zCol_Rubber_Lpost.y = 1390
      zCol_Rubber_Rpost.x = 840 : zCol_Rubber_Rpost.y = 1427
    Case 2
      zCol_Rubber_Lpost.x = 101 : zCol_Rubber_Lpost.y = 1382
      zCol_Rubber_Rpost.x = 840 : zCol_Rubber_Rpost.y = 1419
    Case 3
      zCol_Rubber_Lpost.x = 101 : zCol_Rubber_Lpost.y = 1373
      zCol_Rubber_Rpost.x = 840 : zCol_Rubber_Rpost.y = 1410
  End Select
  For Each BP in BP_Lpost: BP.x = zCol_Rubber_Lpost.x: BP.y = zCol_Rubber_Lpost.y: Next
  For Each BP in BP_Rpost: BP.x = zCol_Rubber_Rpost.x: BP.y = zCol_Rubber_Rpost.y: Next

  ' Cab Side Rails
  CabRails = Table1.Option("Cab Side Rails", 0, 1, 1, 1, 0, Array("Hide", "Show"))

  ' Custom Shooter Knob
  CustomShooter = Table1.Option("VR Custom Shooter Knob", 0, 1, 1, 1, 0, Array("Hide", "Show"))

  ' VRRoom
  VRRoomChoice = Table1.Option("VR Room", 1, 4, 1, 1, 0, Array("VPW Stadium", "Minimal", "Cab Only", "Mixed Reality"))
  LoadVRRoom

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
  Dim a: a = LeftFlipper.CurrentAngle       ' store flipper angle in a
  Dim max_angle, min_angle, mid_angle       ' min and max angles from the flipper
  max_angle = LeftFlipper.StartAngle                  ' flipper down angle
  min_angle = LeftFlipper.EndAngle
  mid_angle = (max_angle-min_angle)/2 + min_angle ' bake map switch point angle
  FlipperLSh.RotZ = a               ' set flipper shadow angle

  Dim BP :
  For Each BP in BP_LFlip
    BP.RotZ = a                 ' rotate the maps
    BP.visible = a > mid_angle
  Next
  For Each BP in BP_LFlipU
    BP.RotZ = a                 ' rotate the maps
    BP.visible = a < mid_angle
  Next
End Sub

Sub RightFlipper_Animate
  Dim a: a = RightFlipper.CurrentAngle        ' store flipper angle in a
  Dim max_angle, min_angle, mid_angle       ' min and max angles from the flipper
  max_angle = RightFlipper.StartAngle               ' flipper down angle
  min_angle = RightFlipper.EndAngle                 ' flipper up angle
  mid_angle = (max_angle-min_angle)/2 + min_angle ' bake map switch point angle
  FlipperRSh.RotZ = a               ' set flipper shadow angle

  Dim BP :
  For Each BP in BP_RFlip
    BP.RotZ = a                 ' rotate the maps
    BP.visible = a < mid_angle
  Next
  For Each BP in BP_RFlipU
    BP.visible = a > mid_angle
    BP.RotZ = a                 ' rotate the maps
  Next
End Sub



'******************************************************
'   ZBBR: BALL BRIGHTNESS
'******************************************************

Const BallBrightness =  1       'Ball brightness - Value between 0 and 1 (0=Dark ... 1=Bright)

' Constants for plunger lane ball darkening.
' You can make a temporary wall in the plunger lane area and use the co-ordinates from the corner control points.
Const PLOffset = 0.5      'Minimum ball brightness scale in plunger lane
Const PLLeft = 1000       'X position of punger lane left
Const PLRight = 1060      'X position of punger lane right
Const PLTop = 1225        'Y position of punger lane top
Const PLBottom = 1900       'Y position of punger lane bottom
Dim PLGain: PLGain = (1-PLOffset)/(PLTop-PLBottom)

Sub UpdateBallBrightness
  Dim s, b_base, b_r, b_g, b_b, d_w
  b_base = 120 * BallBrightness + 135*gi0lvl  ' orig was 120 and 70

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
Dim RoomBrightnessMtlArray: RoomBrightnessMtlArray = Array("VLM.Bake.Active","VLM.Bake.Solid","VLM.Bake.BumperCaps")

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
  If Keycode = LeftFlipperKey Then FlipperActivate LeftFlipper, LFPress: VRCab_FlipperLeft.transx = 6
  If Keycode = RightFlipperKey Then FlipperActivate RightFlipper, RFPress: VRCab_FlipperRight.transx = -6
  If keycode = PlungerKey Then
    Plunger.Pullback
    SoundPlungerPull
    TimerPlunger.Enabled = True
    TimerPlunger2.Enabled = False
    VRCab_Plunger.TransZ = 0
  End If
  If keycode = LeftTiltKey Then Nudge 90, 1 : SoundNudgeLeft      ' Sets the nudge angle and power
  If keycode = RightTiltKey Then Nudge 270, 1 : SoundNudgeRight   ' ^
  If keycode = CenterTiltKey Then Nudge 0, 1 : SoundNudgeCenter   ' ^
  If keycode = StartGameKey Then SoundStartButton: VRCab_StartButton.transy = 5: VRCab_StartButton2.transy = 5
  If keycode = AddCreditKey or keycode = AddCreditKey2 Then
    Select Case Int(rnd*3)
      Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
      Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
      Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
    End Select
  End If
  If keycode = LeftMagnaSave Then Controller.Switch(12) = 1: VRCab_FlipperLeftMagna.transx = -6
  If keycode = keyAddBall Or Keycode = KeyFront Or keycode = LockBarKey Then Controller.Switch(23) = 1: VRCab_BuyInButton.transy = 8: VRCab_BuyinButton2.transy = 8
  If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal keycode)   '***What to do when a button is released***
  If Keycode = LeftFlipperKey Then FlipperDeActivate LeftFlipper, LFPress: VRCab_FlipperLeft.transx = 0
  If Keycode = RightFlipperKey Then FlipperDeActivate RightFlipper, RFPress: VRCab_FlipperRight.transx = 0
  If Keycode = StartGameKey Then Controller.Switch(16) = 0: VRCab_StartButton.transy = 0: VRCab_StartButton2.transy = 0
  If keycode = PlungerKey Then
    Plunger.Fire
    SoundPlungerReleaseBall Plunger
    TimerPlunger.Enabled = False
    TimerPlunger2.Enabled = True
  End If
  If keycode = LeftMagnaSave Then Controller.Switch(12) = 0: VRCab_FlipperLeftMagna.transx = 0
  If keycode = keyAddBall Or Keycode = KeyFront Or keycode = LockBarKey Then Controller.Switch(23) = 0: VRCab_BuyInButton.transy = 0: VRCab_BuyinButton2.transy = 0
  If vpmKeyUp(keycode) Then Exit Sub
End Sub

'******************************************************
'****  ZGIU:  GI Control
'******************************************************

'**** Global variable to hold current GI light intensity. Not Mandatory, but may help in other lighting subs and timers
'**** This value is updated always when GI state is being updated
dim gilvl

'**** Use these for GI strings and stepped GI, and comment out the SetRelayGI
'**** This example table uses Relay for GI control, we don't need these at all

Set GICallback2 = GetRef("GIUpdates2")    'use this for stepped/modulated GI

'GIupdates2 is called always when some event happens to GI channel.
Sub GIUpdates2(aNr, aLvl)
' debug.print "GIUpdates2 nr: " & aNr & " value: " & aLvl
  Dim bulb

  Select Case aNr 'Strings are selected here

    Case 0:  'GI String 1 (Left)
      ' Update the state for each GI light. The state will be a float value between 0 and 1.
      For Each bulb in GILeft: bulb.State = aLvl: Next

      ' If the GI has an associated Relay sound, this can be played
      'If aLvl >= 0.5 And gilvl < 0.5 Then
      ' Sound_GI_Relay 1, Bumper1 'Note: Bumper1 is just used for sound positioning. Can be anywhere that makes sense.
      'ElseIf aLvl <= 0.4 And gilvl > 0.4 Then
      ' Sound_GI_Relay 0, Bumper1
      'End If

      'You may add any other GI related effects here. Like if you want to make some toy to appear more bright, set it like this:
      'Primitive001.blenddisablelighting = 1.2 * aLvl + 0.2 'This will result to DL brightness between 0.2 - 1.4 for ON/OFF states

      gilvl = aLvl    'Storing the latest GI fading state into global variable, so one can use it elsewhere too.
    Case 1:  'GI String 2 (Right)
      ' Update the state for each GI light. The state will be a float value between 0 and 1.
      For Each bulb in GIRight: bulb.State = aLvl: Next

      ' If the GI has an associated Relay sound, this can be played
      'If aLvl >= 0.5 And gilvl < 0.5 Then
      ' Sound_GI_Relay 1, Bumper1 'Note: Bumper1 is just used for sound positioning. Can be anywhere that makes sense.
      'ElseIf aLvl <= 0.4 And gilvl > 0.4 Then
      ' Sound_GI_Relay 0, Bumper1
      'End If

      'You may add any other GI related effects here. Like if you want to make some toy to appear more bright, set it like this:
      'Primitive001.blenddisablelighting = 1.2 * aLvl + 0.2 'This will result to DL brightness between 0.2 - 1.4 for ON/OFF states

      gilvl = aLvl    'Storing the latest GI fading state into global variable, so one can use it elsewhere too.
    Case 4:  'GI String 5 (Upper)
      ' Update the state for each GI light. The state will be a float value between 0 and 1.
      For Each bulb in GITop: bulb.State = aLvl: Next

      ' If the GI has an associated Relay sound, this can be played
      'If aLvl >= 0.5 And gilvl < 0.5 Then
      ' Sound_GI_Relay 1, Bumper1 'Note: Bumper1 is just used for sound positioning. Can be anywhere that makes sense.
      'ElseIf aLvl <= 0.4 And gilvl > 0.4 Then
      ' Sound_GI_Relay 0, Bumper1
      'End If

      'You may add any other GI related effects here. Like if you want to make some toy to appear more bright, set it like this:
      'Primitive001.blenddisablelighting = 1.2 * aLvl + 0.2 'This will result to DL brightness between 0.2 - 1.4 for ON/OFF states

      gilvl = aLvl    'Storing the latest GI fading state into global variable, so one can use it elsewhere too.
    Case 2:  'GI String 3 (Backbox)
    Case 3:  'GI String 4 (Backbox)

  End Select

End Sub

'******************************************************
'****  END GI Control
'******************************************************


'******************************************************
' ZSOL: Solenoids & Flashers
'******************************************************

SolCallback(1)  = "SolGoalPopper"
SolCallback(2)  = "SolTVPopper"
SolCallback(3)  = "SolKickback"
SolCallback(4)  = "SolLockRelease"
SolCallback(5)  = "SolUpperEject"
SolCallback(6)  = "SolTrough"
SolCallback(7)  = "SolKnocker"
SolCallback(8)  = "SolRampDiverter"
SolCallback(14) = "SolRightEject"
SolCallback(15) = "SolLeftEject"
SolCallback(16) = "SolRampDiverterHold"

'***Playfield Flashers (Converted to Lamps)
SolModCallback(17) = "SolFlash17"   'Flasher - Goal cage top
SolModCallback(18) = "SolFlash18"   'Flasher - Goal
SolModCallback(19) = "SolFlash19"   'Flasher - Skill shot
SolModCallback(20) = "SolFlash20"   'Flasher - Bumpers
SolModCallback(21) = "SolFlash21"   'Flasher - Goalie
SolCallback(21) = "SolGoalie"       'Mech - Goalie
SolModCallback(22) = "SolFlash22"   'Flasher - Spinning ball
'SolCallBack(23) = "ttBall.SolMotorState True,"     '"ball_clockwise - for spinnerDisk"
'SolCallBack(24) = "ttBall.SolMotorState False,"      '"ball_counter_clockwise - for spinnerDisk"
SolCallBack(23) = "SolBallMotorCW"      '"ball_clockwise - for spinnerDisk"
SolCallBack(24) = "SolBallMotorCCW"     '"ball_counter_clockwise - for spinnerDisk"
SolModCallback(25) = "SolFlash25"   'Flasher - Left ramp entrance
SolModCallback(26) = "SolFlash26"   'Flasher - Lock area
SolModCallback(27) = "SolFlash27"   'Flasher - Flipper lanes
SolModCallback(28) = "SolFlash28"   'Flasher - Ramp rear

'SolCallback(33) = "SolMag1"
'SolCallback(35) = "SolMag2"

SolCallback(sLRFlipper) = "SolRFlipper" 'Right Flipper
SolCallback(sLLFlipper) = "SolLFlipper" 'Left Flipper

Sub SolKnocker(Enable)
  If Enable Then KnockerSolenoid
End Sub

Sub SolLockRelease(Enable)
  Dim BP
  If Enable Then
    LockPost.collidable = False
    For Each BP in BP_LockPin : BP.z = BP.z - 50 : Next
  Else
    LockPost.collidable = True
    For Each BP in BP_LockPin : BP.z = BP.z + 50 : Next
  End If
End Sub


'*** Flasher Subs

Sub SolFlash17(level)
  f17.state = level
End Sub

Sub SolFlash18(level)
  f18.state = level
End Sub

Sub SolFlash19(level)
  f19.state = level
End Sub

Sub SolFlash20(level)
  f20.state = level
End Sub

Sub SolFlash21(level)
  f21.state = level
End Sub

Sub SolFlash22(level)
  f22.state = level
  f22a.state = level
End Sub

Sub SolFlash25(level)
  f25.state = level
End Sub

Sub SolFlash26(level)
  f26.state = level
End Sub

Sub SolFlash27(level)
  f27.state = level
  f27a.state = level
End Sub

Sub SolFlash28(level)
  f28.state = level
  f28a.state = level
End Sub



'******************************************************
' ZDRN: Drain, Trough, and Ball Release
'******************************************************

'********************* TROUGH *************************

Sub sw31_Hit   : Controller.Switch(31) = 1 : UpdateTrough : End Sub
Sub sw31_UnHit : Controller.Switch(31) = 0 : UpdateTrough : End Sub
Sub sw32_Hit   : Controller.Switch(32) = 1 : UpdateTrough : End Sub
Sub sw32_UnHit : Controller.Switch(12) = 0 : UpdateTrough : End Sub
Sub sw33_Hit   : Controller.Switch(33) = 1 : UpdateTrough : End Sub
Sub sw33_UnHit : Controller.Switch(33) = 0 : UpdateTrough : End Sub
Sub sw34_Hit   : Controller.Switch(34) = 1 : UpdateTrough : End Sub
Sub sw34_UnHit : Controller.Switch(34) = 0 : UpdateTrough : End Sub
Sub sw35_Hit    : Controller.Switch(35)  = 1 : UpdateTrough : End Sub
Sub sw35_UnHit  : Controller.Switch(35)  = 0 : UpdateTrough : End Sub
Sub Drain_Hit : RandomSoundDrain Drain : UpdateTrough : End Sub

Sub UpdateTrough
  UpdateTroughTimer.Interval = 100
  UpdateTroughTimer.Enabled = 1
End Sub

Sub UpdateTroughTimer_Timer
  If sw31.BallCntOver = 0 Then sw32.kick 57, 10
  If sw32.BallCntOver = 0 Then sw33.kick 57, 10
  If sw33.BallCntOver = 0 Then sw34.kick 57, 10
  If sw34.BallCntOver = 0 Then sw35.kick 57, 10
  If sw35.BallCntOver = 0 Then Drain.kick 57, 10
  Me.Enabled = 0
End Sub


'*****************  DRAIN & RELEASE  ******************

Sub SolTrough(enabled)
  If enabled Then
    vpmTimer.PulseSw 36
    sw31.kick 57, 10
    RandomSoundBallRelease sw31
  End If
End Sub


'******************************************************
' ZFLP: FLIPPERS
'******************************************************
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
Dim LStep, RStep

Sub RightSlingShot_Slingshot
  RS.VelocityCorrect(Activeball)
  vpmTimer.PulseSw(85)            'Sling Switch Number
  'RSling1.Visible = 1
  'Sling1.TransY =  - 20            'Sling Metal Bracket
  RStep = 0
  RightSlingShot.TimerEnabled = 1
  RightSlingShot.TimerInterval = 10
  RandomSoundSlingshotRight zCol_Rubber_Post021
End Sub

Sub RightSlingShot_Timer
  Dim BP
  Dim x1, x2, y: x1 = False:x2 = True:y = 20
    Select Case RStep
        Case 2:x1 = True:x2 = False:y = 10
        Case 3:x1 = False:x2 = False:y = 0:RightSlingShot.TimerEnabled = 0
    End Select

  For Each BP in BP_RSling1 : BP.Visible = x1: Next
  For Each BP in BP_RSling2 : BP.Visible = x2: Next
  For Each BP in BP_remk : BP.transx = y: Next

  RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  LS.VelocityCorrect(Activeball)
  vpmTimer.PulseSw(84)            'Sling Switch Number
  'LSling1.Visible = 1
  'Sling2.TransY =  - 20              'Sling Metal Bracket
  LStep = 0
  LeftSlingShot.TimerEnabled = 1
  LeftSlingShot.TimerInterval = 10
  RandomSoundSlingshotLeft zCol_Rubber_Post042
End Sub

Sub LeftSlingShot_Timer
  Dim BP
  Dim x1, x2, y: x1 = False:x2 = True:y = 20
    Select Case LStep
        Case 2:x1 = True:x2 = False:y = 10
        Case 3:x1 = False:x2 = False:y = 0:LeftSlingShot.TimerEnabled = 0
    End Select

  For Each BP in BP_LSling1 : BP.Visible = x1: Next
  For Each BP in BP_LSling2 : BP.Visible = x2: Next
  For Each BP in BP_lemk : BP.transx = y: Next

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

' The following sub are needed, however they exist in the ZMAT maths section of the script. Uncomment below if needed
'Dim PI: PI = 4*Atn(1)
'Function dSin(degrees)
' dsin = sin(degrees * Pi/180)
'End Function
'Function dCos(degrees)
' dcos = cos(degrees * Pi/180)
'End Function
'
'Function RotPoint(x,y,angle)
' dim rx, ry
' rx = x*dCos(angle) - y*dSin(angle)
' ry = x*dSin(angle) + y*dCos(angle)
' RotPoint = Array(rx,ry)
'End Function

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

Sub Gate001_Animate
  Dim a : a = Gate001.CurrentAngle
' Dim BP : For Each BP in BP_Gate001 : BP.rotx = a: Next
End Sub

Sub Gate002_Animate
  Dim a : a = Gate002.CurrentAngle
  Dim BP : For Each BP in BP_Gate002 : BP.rotx = a: Next
End Sub

Sub Gate003_Animate
  Dim a : a = Gate003.CurrentAngle
  Dim BP : For Each BP in BP_Gate003 : BP.rotx = a: Next
End Sub

Sub Gate004_Animate
  Dim a : a = Gate004.CurrentAngle
  Dim BP : For Each BP in BP_Gate004 : BP.rotx = a: Next
End Sub

Sub Gate006_Animate
  Dim a : a = Gate006.CurrentAngle
  Dim BP : For Each BP in BP_Gate006 : BP.rotx = a: Next
End Sub

Sub Gate007_Animate
  Dim a : a = Gate007.CurrentAngle
  Dim BP : For Each BP in BP_Gate007 : BP.rotx = a: Next
End Sub


'************************************************************
' ZSWI: SWITCHES
'************************************************************

Sub sw27_Spin():vpmTimer.PulseSw 27:SoundSpinner sw27:End Sub

Sub sw27_Animate
    Dim BP, a, b, offset
    a = sw27.currentangle

    For Each BP in BP_sw27: BP.RotX = -a: Next
End Sub

Sub sw65_Hit() : vpmTimer.PulseSw 65 : End Sub

Sub sw76_Hit()
  Dim BP
  For Each BP in BP_sw76 : BP.rotz = BP.rotz + 45 : Next
  Controller.Switch(76) = 1
End Sub

Sub sw76_UnHit
  Dim BP
  For Each BP in BP_sw76 : BP.rotz = BP.rotz - 45 : Next
  Controller.Switch(76) = 0
End Sub

Sub sw77_Hit()
  Dim BP
  For Each BP in BP_sw77 : BP.rotz = BP.rotz + 45 : Next
  Controller.Switch(77) = 1
End Sub

Sub sw77_UnHit
  Dim BP
  For Each BP in BP_sw77 : BP.rotz = BP.rotz - 45 : Next
  Controller.Switch(77) = 0
End Sub

'TARGETS

Sub sw16_Hit
  STHit 16
End Sub

Sub sw25_Hit
  STHit 25
End Sub

Sub sw28_Hit
  STHit 28
End Sub

Sub sw37_Hit
  STHit 37
End Sub

Sub sw66_Hit
  STHit 66
End Sub

Sub sw67_Hit
  STHit 67
End Sub


'************************* Bumpers **************************
Sub Bumper1_Hit()
  RandomSoundBumperTop Bumper1
  vpmTimer.PulseSw 82

  Dim BP
  For Each BP in BP_BS1 : BP.roty=skirtAY(me,Activeball) : Next
  For Each BP in BP_BS1 : BP.rotx=skirtAX(me,Activeball) : Next

  me.timerinterval = 150
  me.timerenabled=1
End Sub

sub Bumper1_timer
  Dim BP
  For Each BP in BP_Bs1 : BP.roty=0 : Next
  For Each BP in BP_Bs1 : BP.rotx=0 : Next
  me.timerenabled=0
end sub

Sub Bumper2_Hit()
  RandomSoundBumperMiddle Bumper2
  vpmTimer.PulseSw 81

  Dim BP
  For Each BP in BP_BS2 : BP.roty=skirtAY(me,Activeball) : Next
  For Each BP in BP_BS2 : BP.rotx=skirtAX(me,Activeball) : Next

  me.timerinterval = 150
  me.timerenabled=1
End Sub

sub Bumper2_timer
  Dim BP
  For Each BP in BP_Bs2 : BP.roty=0 : Next
  For Each BP in BP_Bs2 : BP.rotx=0 : Next
  me.timerenabled=0
end sub

Sub Bumper3_Hit()
  RandomSoundBumperBottom Bumper3
  vpmTimer.PulseSw 83

  Dim BP
  For Each BP in BP_BS3 : BP.roty=skirtAY(me,Activeball) : Next
  For Each BP in BP_BS3 : BP.rotx=skirtAX(me,Activeball) : Next

  me.timerinterval = 150
  me.timerenabled=1
End Sub

sub Bumper3_timer
  Dim BP
  For Each BP in BP_Bs3 : BP.roty=0 : Next
  For Each BP in BP_Bs3 : BP.rotx=0 : Next
  me.timerenabled=0
end sub

Sub Bumper1_Animate
  Dim z: z = Bumper1.CurrentRingOffset

  Dim BP
  For Each BP in BP_Br1 : BP.transz = z : Next
End Sub

Sub Bumper2_Animate
  Dim z: z = Bumper2.CurrentRingOffset

  Dim BP
  For Each BP in BP_Br2 : BP.transz = z : Next
End Sub

Sub Bumper3_Animate
  Dim z: z = Bumper3.CurrentRingOffset

  Dim BP
  For Each BP in BP_Br3 : BP.transz = z : Next
End Sub

'******************************************************
'     SKIRT ANIMATION FUNCTIONS
'******************************************************
' NOTE: set bumper object timer to around 150-175 in order to be able
' to actually see the animaation, adjust to your liking

'Const PI = 3.1415926
Const SkirtTilt=5   'angle of skirt tilting in degrees

Function SkirtAX(bumper, bumperball)
  skirtAX=cos(skirtA(bumper,bumperball))*(SkirtTilt)    'x component of angle
  if (bumper.y<bumperball.y) then skirtAX=skirtAX*-1  'adjust for ball hit bottom half
End Function

Function SkirtAY(bumper, bumperball)
  skirtAY=sin(skirtA(bumper,bumperball))*(SkirtTilt)    'y component of angle
  if (bumper.x>bumperball.x) then skirtAY=skirtAY*-1  'adjust for ball hit left half
End Function

Function SkirtA(bumper, bumperball)
  dim hitx, hity, dx, dy
  hitx=bumperball.x
  hity=bumperball.y

  dy=Abs(hity-bumper.y)         'y offset ball at hit to center of bumper
  if dy=0 then dy=0.0000001
  dx=Abs(hitx-bumper.x)         'x offset ball at hit to center of bumper
  skirtA=(atn(dx/dy)) '/(PI/180)      'angle in radians to ball from center of Bumper1
End Function


'************************ Rollovers *************************
Sub sw38_Hit(): Controller.Switch(38) = 1: AnimateWireHeight BP_sw38, 1, false :  End Sub 'Plunger Lane Rollover Switch
Sub sw38_UnHit: Controller.Switch(38) = 0: AnimateWireHeight BP_sw38, 0, false :  End Sub
Sub sw87_Hit(): Controller.Switch(87) = 1: AnimateWireHeight BP_sw87, 1, false : End Sub  'Top Left Rollover Switch
Sub sw87_UnHit: Controller.Switch(87) = 0: AnimateWireHeight BP_sw87, 0, false : End Sub
Sub sw88_Hit(): Controller.Switch(88) = 1: AnimateWireHeight BP_sw88, 1, false : End Sub  'Top Right Rollover Switch
Sub sw88_UnHit: Controller.Switch(88) = 0: AnimateWireHeight BP_sw88, 0, false : End Sub
Sub sw15_Hit(): Controller.Switch(15) = 1: AnimateWireHeight BP_sw15, 1, false : leftInlaneSpeedLimit: End Sub  'Left Inlane Rollover Switch
Sub sw15_UnHit: Controller.Switch(15) = 0: AnimateWireHeight BP_sw15, 0, false : End Sub
Sub sw86_Hit(): Controller.Switch(86) = 1: AnimateWireHeight BP_sw86, 1, false : End Sub  'Left Outlane Rollover Switch / Kickback
Sub sw86_UnHit: Controller.Switch(86) = 0: AnimateWireHeight BP_sw86, 0, false : End Sub
Sub sw26_Hit(): Controller.Switch(26) = 1: AnimateWireHeight BP_sw26, 1, false : End Sub  'Left Outlane Rollover Switch / Kickback Upper
Sub sw26_UnHit: Controller.Switch(26) = 0: AnimateWireHeight BP_sw26, 0, false : End Sub
Sub sw18_Hit(): Controller.Switch(18) = 1: AnimateWireHeight BP_sw18, 1, false : End Sub  'Right Outlane Rollover Switch
Sub sw18_UnHit: Controller.Switch(18) = 0: AnimateWireHeight BP_sw18, 0, false : End Sub
Sub sw17_Hit(): Controller.Switch(17) = 1: AnimateWireHeight BP_sw17, 1, false : rightInlaneSpeedLimit: End Sub 'Right Inlane Rollover Switch
Sub sw17_UnHit: Controller.Switch(17) = 0: AnimateWireHeight BP_sw17, 0, false : End Sub
Sub sw47_Hit(): Controller.Switch(47) = 1: AnimateWireHeight BP_sw47, 1, false : End Sub  'Left Orbit Rollover Switch
Sub sw47_UnHit: Controller.Switch(47) = 0: AnimateWireHeight BP_sw47, 0, false : End Sub

Sub sw61_Hit()
  Controller.Switch(61) = 1
  AnimateWireHeight BP_sw61, 1, true
  'playfield_button1.collidable = 0
End Sub

Sub sw61_UnHit()
  Controller.Switch(61) = 0
  AnimateWireHeight BP_sw61, 0, true
  'playfield_button1.collidable = 1
End Sub

Sub sw62_Hit()
  Controller.Switch(62) = 1
  AnimateWireHeight BP_sw62, 1, true
  'playfield_button2.collidable = 0
End Sub

Sub sw62_UnHit()
  Controller.Switch(62) = 0
  AnimateWireHeight BP_sw62, 0, true
  'playfield_button2.collidable = 1
End Sub

Sub sw63_Hit()
  Controller.Switch(63) = 1
  AnimateWireHeight BP_sw63, 1, true
  'playfield_button3.collidable = 0
End Sub

Sub sw63_UnHit()
  Controller.Switch(63) = 0
  AnimateWireHeight BP_sw63, 0, true
  'playfield_button3.collidable = 1
End Sub

Sub sw64_Hit()
  Controller.Switch(64) = 1
  AnimateWireHeight BP_sw64, 1, true
  'playfield_button4.collidable = 0
End Sub

Sub sw64_UnHit()
  Controller.Switch(64) = 0
  AnimateWireHeight BP_sw64, 0, true
  'playfield_button4.collidable = 1
End Sub

'************************* Coin Toss ************************
Sub sw51_Hit(): Controller.Switch(51) = 1:End Sub 'Coin toss lower
Sub sw52_Hit(): Controller.Switch(52) = 1:End Sub 'Coin toss middle
Sub sw53_Hit(): Controller.Switch(53) = 1:End Sub 'Coin toss upper

Sub toss_unhit_hit()
  Controller.Switch(51) = 0
  Controller.Switch(52) = 0
  Controller.Switch(53) = 0
  Debug.Print "Toss Unhit"
End Sub

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

Sub AnimateWireHeight(group, action, button) ' Action = 1 - to drop, 0 to raise)
  Dim drop : drop = -13
  if button then drop = -2
  Dim BP
  If action = 1 Then
    For Each BP in group : BP.transz = drop : Next
  Else
    For Each BP in group : BP.transz = 0 : Next
  End If
End Sub



'********************* Standup Targets **********************

Sub sw41_Hit(): Controller.Switch(41) = 1: End Sub    'Goal Switch
Sub sw41_UnHit: Controller.Switch(41) = 0: End Sub

'************************************************************
' ZVUK: VUKs and Kickers
'************************************************************

' TODO: Fix VUK angles/velocity

'******************* VUKs **********************

Dim KickerBall42, KickerBall45, KickerBall54, KickerBall55, KickerBall56    'Each VUK needs its own "kickerball"

Sub KickBall(kball, kangle, kvel, kvelz, kzlift)  'Defines how KickBall works
  dim rangle
  rangle = PI * (kangle - 90) / 180

  kball.z = kball.z + kzlift
  kball.velz = kvelz
  kball.velx = cos(rangle)*kvel
  kball.vely = sin(rangle)*kvel
End Sub

'Goal VUK
Sub sw42_Hit                    'Switch associated with Kicker
    set KickerBall42 = activeball
    Controller.Switch(42) = 1
    SoundSaucerLock
  wall_sw42.collidable = True
End Sub

Sub sw42_unhit
  KickerBall42 = Empty
  Controller.Switch(42) = 0
End Sub

Sub sw42_exit_Hit
  wall_sw42.collidable = False
End Sub

Sub SolGoalPopper(Enable)             'Solonoid name associated with kicker.
    If Enable then
    If Not IsEmpty(KickerBall42) Then
      KickBall KickerBall42, 0, 0, 40, 0
      SoundSaucerKick 1, sw42
      WireRampOn False
    End If
  End If
End Sub

'TV VUK
Sub sw45_Hit                    'Switch associated with Kicker
    set KickerBall45 = activeball
    Controller.Switch(45) = 1
    SoundSaucerLock
End Sub

Sub SolTVPopper(Enable)             'Solonoid name associated with kicker.
    If Enable then
    If Controller.Switch(45) <> 0 Then
      KickBall KickerBall45, 20, 50, 0, 0
      SoundSaucerKick 1, sw45
      Controller.Switch(45) = 0
    End If
  End If
End Sub

Sub sw54_Hit
    set KickerBall54 = activeball
    Controller.Switch(54) = 1
    SoundSaucerLock
End Sub

Sub SolRightEject(Enable)
    If Enable then
    If Controller.Switch(54) <> 0 Then
      KickBall KickerBall54, 210, 12, 5, 10
      SoundSaucerKick 1, sw54
      Controller.Switch(54) = 0
    End If
  End If
End Sub

Sub sw55_Hit
    set KickerBall55 = activeball
    Controller.Switch(55) = 1
    SoundSaucerLock
End Sub

Sub SolUpperEject(Enable)
    If Enable then
    If Controller.Switch(55) <> 0 Then
      KickBall KickerBall55, 3, 20, 5, 15
      SoundSaucerKick 1, sw55
      Controller.Switch(55) = 0
    End If
  End If
End Sub

Sub sw56_Hit
    set KickerBall56 = activeball
    Controller.Switch(56) = 1
    SoundSaucerLock
End Sub

Sub SolLeftEject(Enable)
    If Enable then
    If Controller.Switch(56) <> 0 Then
      KickBall KickerBall56, 160, 15, 5, 10
      SoundSaucerKick 1, sw56
      Controller.Switch(56) = 0
    End If
  End If
End Sub

'Impulse Plunger
Sub SolAutofire(Enabled)
  If Enabled Then
    PlungerIM.AutoFire
    SoundPlungerReleaseBall Plunger
  End If
End Sub

Const IMPowerSetting = 85
Const IMTime = 0.6
Set plungerIM = New cvpmImpulseP
With plungerIM
    .InitImpulseP swPlunger, IMPowerSetting, IMTime
    .Random 0.3
    .CreateEvents "plungerIM"
End With

Sub SolKickBack(enabled)
    If enabled Then
    SoundPlungerReleaseBall KickBackPlunger
    KickbackPlunger.Fire
  Else
    KickbackPlunger.PullBack
  End If
End Sub

'*********
'DIVERTER
'*********

Dim LockPower, LockHold

Sub SolRampDiverter(enabled)  'Lock Diverter
  LockPower = enabled
  If enabled Then
    DiverterFlipper001.rotatetoend
    diverterwall_open.collidable = 1
    diverterwall_closed.collidable = 0
    PlaySoundAt SoundFX ("fx_diverter_open", DOFContactors), RampTrigger7
  End If
  If Not enabled AND Not LockHold Then
    DiverterFlipper001.rotatetostart
    diverterwall_open.collidable = 0
    diverterwall_closed.collidable = 1
    PlaySoundAt SoundFX ("fx_diverter_open", DOFContactors), RampTrigger7
  End If
End Sub

Sub SolRampDiverterHold(enabled)
  LockHold = enabled
  If Not enabled AND Not LockPower Then
    DiverterFlipper001.rotatetostart
    diverterwall_open.collidable = 0
    diverterwall_closed.collidable = 1
  End If
End Sub

Sub DiverterFlipper001_Animate()
  Dim a : a = DiverterFlipper001.CurrentAngle
  Dim max_angle, min_angle, mid_angle       ' min and max angles from the flipper
  max_angle = DiverterFlipper001.StartAngle   ' flipper down angle
  min_angle = DiverterFlipper001.EndAngle
  mid_angle = (max_angle-min_angle)/2 + min_angle ' bake map switch point angle

  Dim BP :
  For Each BP in BP_diverterside01
    BP.RotZ = -(a - 172)              ' rotate the maps
    BP.visible = a < mid_angle
  Next
  For Each BP in BP_diverterside02
    BP.RotZ = -(a - 172)              ' rotate the maps
    BP.visible = a > mid_angle
  Next
End Sub

'********************
'SOCCERBALL_ANIMATION
'********************

Dim BallCW: BallCW = False
Dim BallCCW: BallCCW = False
Dim BallDir: BallDir = 0 '1=CC, -1=CCW
Dim BallAngSpeed: BallAngSpeed = 0
Dim BallAngle: BallAngle = 0
Dim BlurBall1: BlurBall1 = 0
Dim BlurBall2: BlurBall2 = 0

Const BallMaxSpeed = 1200   'deg per second
Const BallAccel = 900       'deg per second^2


Sub UpdateSoccerBall()
  Dim SubAngle

  'Update Speed
  If BallCW Then
    BallDir = 1
  ElseIf BallCCW Then
    BallDir = -1
    End If
  If BallCW OR BallCCW Then
    If abs(BallAngSpeed) < BallMaxSpeed Then BallAngSpeed = BallAngSpeed + (FrameTime/1000)*BallAccel*BallDir
  Else
    If abs(BallAngSpeed) > 0 Then BallAngSpeed = BallAngSpeed - (FrameTime/1000)*BallAccel*BallDir
  End If
  If BallAngSpeed > BallMaxSpeed Then BallAngSpeed = BallMaxSpeed
  If BallAngSpeed < -BallMaxSpeed Then BallAngSpeed = -BallMaxSpeed
  If abs(BallAngSpeed) < 3 Then BallAngSpeed = 0

  'Update Angle
  BallAngle = BallAngle + (FrameTime/1000)*BallAngSpeed
  If BallAngle > 360 Then BallAngle = BallAngle - 360
  If BallAngle < 0 Then BallAngle = BallAngle + 360
  SubAngle = BallAngle Mod 72

  'Handle sound effect
  if BallAngSpeed <> 0 then
    'Debug.Print("BallAngSpeed: " & BallAngSpeed)
    'PlaySound(string, int loopcount, float volume, float pan, float randompitch, int pitch, bool useexisting, bool restart, float front_rear_fade)
    If BallAngSpeed >= 1050 Then
      stopsound "fx_motor2"
      stopsound "fx_motor3"
      stopsound "fx_motor4"
      stopsound "fx_motor5"
      stopsound "fx_motor6"
      stopsound "fx_motor7"
      playsound "fx_motor1", -1, 0.01, 0, 0, 1, 1, 0
    ElseIf BallAngSpeed < 1050 And BallAngSpeed >= 900 Then
      stopsound "fx_motor1"
      stopsound "fx_motor3"
      stopsound "fx_motor4"
      stopsound "fx_motor5"
      stopsound "fx_motor6"
      stopsound "fx_motor7"
      playsound "fx_motor2", -1, 0.01, 0, 0, 1, 1, 0
    ElseIf BallAngSpeed < 900 And BallAngSpeed >= 750 Then
      stopsound "fx_motor1"
      stopsound "fx_motor2"
      stopsound "fx_motor4"
      stopsound "fx_motor5"
      stopsound "fx_motor6"
      stopsound "fx_motor7"
      playsound "fx_motor3", -1, 0.01, 0, 0, 1, 1, 0
    ElseIf BallAngSpeed < 750 And BallAngSpeed >= 600 Then
      stopsound "fx_motor1"
      stopsound "fx_motor2"
      stopsound "fx_motor3"
      stopsound "fx_motor5"
      stopsound "fx_motor6"
      stopsound "fx_motor7"
      playsound "fx_motor4", -1, 0.01, 0, 0, 1, 1, 0
    ElseIf BallAngSpeed < 600 And BallAngSpeed >= 450 Then
      stopsound "fx_motor1"
      stopsound "fx_motor2"
      stopsound "fx_motor3"
      stopsound "fx_motor4"
      stopsound "fx_motor6"
      stopsound "fx_motor7"
      playsound "fx_motor5", -1, 0.01, 0, 0, 1, 1, 0
    ElseIf BallAngSpeed < 450 And BallAngSpeed >= 200 Then
      stopsound "fx_motor1"
      stopsound "fx_motor2"
      stopsound "fx_motor3"
      stopsound "fx_motor4"
      stopsound "fx_motor5"
      stopsound "fx_motor7"
      playsound "fx_motor6", -1, 0.01, 0, 0, 1, 1, 0
    Else
      stopsound "fx_motor1"
      stopsound "fx_motor2"
      stopsound "fx_motor3"
      stopsound "fx_motor4"
      stopsound "fx_motor5"
      stopsound "fx_motor6"
      playsound "fx_motor7", -1, 0.01, 0, 0, 1, 1, 0
    End If
  Else
    stopsound "fx_motor1"'
    stopsound "fx_motor2"
    stopsound "fx_motor3"
    stopsound "fx_motor4"
    stopsound "fx_motor5"
    stopsound "fx_motor6"
    stopsound "fx_motor7"
  End if
  if BallAngSpeed = 0 then stopsound "fx_motor7"'

  'Handle turn table
  ttball.Speed = BallAngSpeed*0.03

  'Texture swap for blurred ball
  If BallAngSpeed > BallMaxSpeed*0.75 And BlurBall1 = 0 Then
    blurred_ball.image = "ballBlur1"
    blurred_ball.visible = 1
    BM_soccerBall_001.visible = 0
    BlurBall1 = 1
  End If
  If BallAngSpeed > BallMaxSpeed*0.9 And BlurBall2 = 0 Then
    blurred_ball.image = "ballBlur2"
    BlurBall2 = 1
  End If
  If BallAngSpeed < BallMaxSpeed*0.9 And BlurBall2 = 1 Then
    blurred_ball.image = "ballBlur1"
    BlurBall2 = 0
  End If
  If BallAngSpeed < BallMaxSpeed*0.75 And BlurBall1 = 1 Then
    blurred_ball.visible = 0
    BM_soccerBall_001.visible = 1
    BlurBall1 = 0
  End If


  If BallAngSpeed=0 Then Exit Sub
  'Debug.Print "BallDir="&BallDir&" BallAngSpeed="&BallAngSpeed&" SubAngle="&SubAngle&" BallCW="&BallCW&" BallCCW="&BallCCW
  UpdateSoccerBallAnim(SubAngle)
End Sub


Sub UpdateSoccerBallAnim(SubAngle)
  Dim BP, Op1, Op2, Op3, a
  'Calculate new opacities
  If SubAngle >= 0 And SubAngle < 24 Then
    Op1 = (1 - SubAngle/24) * 100
    Op2 = (SubAngle/24) * 100
    Op3 = 0
  ElseIf SubAngle >= 24 And SubAngle < 48 Then
    Op1 = 0
    Op2 = (1 - (SubAngle-24)/24) * 100
    Op3 = ((SubAngle-24)/24) * 100
  Else
    Op1 = ((SubAngle-48)/24) * 100
    Op2 = 0
    Op3 = (1 - (SubAngle-48)/24) * 100
  End If
  'debug.print "SubAngle="&SubAngle&" Op1="&Op1&" Op2="&Op2&" Op3="&Op3

  'Update animations
  For Each BP In BP_soccerBall_001
    If SubAngle >= 48 And SubAngle < 72 Then
      BP.RotZ = SubAngle-72
      a = SubAngle-72
    Else
      BP.RotZ = SubAngle
      a = SubAngle
    End If
    BP.Opacity = Op1
  Next
  BM_soccerBall_001.Opacity = 100
  For Each BP In BP_soccerBall_002
    BP.RotZ = SubAngle
    BP.Opacity = Op2
  Next
  For Each BP In BP_soccerBall_003
    BP.RotZ = SubAngle
    BP.Opacity = Op3
  Next
  blurred_ball.RotZ = SubAngle
End Sub


'Init soccer ball
InitSoccerBall
Sub InitSoccerBall
  Dim BP
  'Scale to prevent z-fighting amongst the three positions
  For Each BP in BP_soccerBall_001 : BP.size_x = 1.000: BP.size_y = 1.000: BP.size_z = 1.000: BP.z = BP.z * 1.000: BP.visible=1: Next
  For Each BP in BP_soccerBall_002 : BP.size_x = 1.005: BP.size_y = 1.005: BP.size_z = 1.005: BP.z = BP.z * 1.005: BP.visible=1: Next
  For Each BP in BP_soccerBall_003 : BP.size_x = 1.010: BP.size_y = 1.010: BP.size_z = 1.010: BP.z = BP.z * 1.010: BP.visible=1: Next
  'BM_soccerBall_001.size_x = 0.995: BM_soccerBall_001.size_y = 0.995:  BM_soccerBall_001.size_z = 0.995
  blurred_ball.size_x = 99.5: blurred_ball.size_y = 99.5: blurred_ball.size_z = 99.5
  'Only one bakemap will be used
  BM_soccerBall_001.Visible = 1
  BM_soccerBall_002.Visible = 0
  BM_soccerBall_003.Visible = 0
  blurred_ball.visible = 0
  UpdateSoccerBallAnim(0)
End Sub


'Turn table effect
Sub SoccerBall_Hit(): ttBall.AddBall ActiveBall: End Sub
Sub SoccerBall_UnHit(): ttBall.RemoveBall ActiveBall: End Sub
Sub ttBall_timer(): ttBall.Update(): End Sub


'Sound
Sub LoopMotorSoundTimer_Timer
  LoopMotorSoundTimer.Enabled = False
End Sub

'Solenoids
Sub SolBallMotorCW(Enabled)
  BallCW = Enabled
  ttBall.SolMotorState True,Enabled
End Sub

Sub SolBallMotorCCW(Enabled)
  BallCCW = Enabled
  ttBall.SolMotorState False,Enabled
End Sub




'****************
'GOALIEANIMATION
'****************
Const GoalieSpeed = 0.0047
Dim GoalieWalls, NumGoalieWalls, GoalieTheta, GSinTheta, GPos, GNewPos, GLastPos, GoalieEnabled
GoalieEnabled = False
GNewPos = 0
GLastPos = 0

GoalieWalls = Array(GT040,GT039,GT038,GT037,GT036,GT035,GT034,GT033,GT032,GT031,GT030,GT029,GT028,GT027,GT026,GT025,GT024,GT023,GT022,GT021,GT020,GT019,GT018,GT017,GT016,GT015,GT014,GT013,GT012,GT011,GT010,GT009,GT008,GT007,GT006)
NumGoalieWalls = UBound(GoalieWalls)

InitGoalie
Sub InitGoalie
  Dim BP
  For Each BP in BP_goalkeeper : BP.x = 376 : BP.y = 280 : BP.z = -10 :Next
  GoalieTheta = PI/2
  Controller.Switch(44) = True
End Sub

Sub SolGoalie(Enabled)
  GoalieEnabled = Enabled
  If Enabled Then
    playsound "fx_goaliedrive", -1, 0.2, 0, 0, 1, 1, 0
  Else
    stopsound "fx_goaliedrive"
  End If
End Sub


'Sub UpdateGoalie(aNewPos,aSpeed,aLastPos)
Sub UpdateGoalie
  'Calculate new goalie position
  GoalieTheta = GoalieTheta + GoalieSpeed*FrameTime
  if GoalieTheta >= 2*PI Then GoalieTheta = GoalieTheta - 2*PI
  GSinTheta = 1.3*Sin(GoalieTheta)
  If GSinTheta < -1 Then GSinTheta = -1
  If GSinTheta > 1 Then GSinTheta = 1

  'Actuate switches
  If GSinTheta < 0.0 and Controller.Switch(43) = False Then Controller.Switch(43) = True
  If GSinTheta > 0.0 and Controller.Switch(43) = True Then Controller.Switch(43) = False
  If GSinTheta > 0.1 and Controller.Switch(44) = False Then Controller.Switch(44) = True
  If GSinTheta < 0.1 and Controller.Switch(44) = True Then Controller.Switch(44) = False

  'Handle collidable targets
  GPos = Int(NumGoalieWalls*0.5*(GSinTheta+1))
  GLastPos = GNewPos
  GNewPos = GPos
  If GNewPos <> GLastPos Then
    GoalieWalls(GNewPos).collidable = 1: GoalieWalls(GLastPos).collidable = 0
    'GoalieWalls(GNewPos).visible = 1:    GoalieWalls(GLastPos).visible = 0    'visualize physical target, for debugging
  End If

  'Animate goalie
  Dim BP: For Each BP in BP_goalkeeper : BP.transx = 50*GSinTheta : BP.rotY = 10*GSinTheta : Next
End Sub


'Shake Goalie when hit

Sub GWalls_Hit(idx)
  vpmTimer.PulseSw 48
  ShakeGoalie
  RandomSoundWall()
End Sub

Sub ShakeGoalie()
  Dim BP: For Each BP in BP_goalkeeper : BP.transy = -8 : Next
  ShakeGoalieTimer.Enabled = 1
End Sub

Sub ShakeGoalieTimer_Timer()
  Dim BP: For Each BP in BP_goalkeeper : BP.transy = 0 : Next
  ShakeGoalieTimer.Enabled = 0
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
Sub sw71_Hit
  vpmTimer.PulseSw 71
  Dim BP
  For Each BP in BP_sw71 : BP.rotx = BP.rotx + 45 : Next
End Sub

Sub sw71_UnHit
  Dim BP
  For Each BP in BP_sw71 : BP.rotx = BP.rotx - 45 : Next
End Sub

Sub sw72_Hit
  vpmTimer.PulseSw 72
  If activeball.vely < 0 Then
    WireRampOn True
  Else
    WireRampOff
  End If
End Sub

Sub sw74_Hit
  vpmTimer.PulseSw 74
  Dim BP
  For Each BP in BP_sw74 : BP.rotx = BP.rotx + 45 : Next
End Sub

Sub sw74_UnHit
  Dim BP
  For Each BP in BP_sw74 : BP.rotx = BP.rotx - 45 : Next
End Sub

Sub sw75_Hit
  vpmTimer.PulseSw 75
  If activeball.vely < 0 Then
    WireRampOn True
  Else
    WireRampOff
  End If
End Sub

Sub sw78_Hit
  vpmTimer.PulseSw 78
  Dim BP
  For Each BP in BP_sw78 : BP.rotx = BP.rotx + 45 : Next
End Sub

Sub sw78_UnHit
  Dim BP
  For Each BP in BP_sw78 : BP.rotx = BP.rotx - 45 : Next
End Sub

Sub RampTrigger1_Hit
  RandomSoundDelayedBallDropOnPlayfield ActiveBall
    WireRampOff
End Sub

Sub RampTrigger2_Hit
  RandomSoundDelayedBallDropOnPlayfield ActiveBall
    WireRampOff
End Sub

Sub RampTrigger4_Hit
  RandomSoundDelayedBallDropOnPlayfield ActiveBall
    WireRampOff
End Sub

Sub RampTrigger5_Hit
  RandomSoundDelayedBallDropOnPlayfield ActiveBall
    WireRampOff
End Sub

Sub RampTrigger6_Hit
  RandomSoundDelayedBallDropOnPlayfield ActiveBall
    WireRampOff
End Sub

Sub RampTrigger7_Hit
  RandomSoundDelayedBallDropOnPlayfield ActiveBall
End Sub

Sub RampTrigger8_Hit
  WireRampOff
  RandomSoundRampStop RampTrigger8
  RandomSoundDelayedBallDropOnPlayfield ActiveBall
End Sub

Sub RampTrigger001_Hit
    WireRampOn False
End Sub

Sub RampTrigger002_Hit
  RandomSoundDelayedBallDropOnPlayfield ActiveBall
    WireRampOff
End Sub


'Sub RampTrigger5_Hit
' If activeball.vely < 0 Then
'   WireRampOn True
' Else
'   WireRampOff
' End If
'End Sub

'Sub RampTrigger6_Hit
' WireRampOff
'End Sub

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

Sub SoundPlungerReleaseBall(pl)
  PlaySoundAtLevelStatic ("Plunger_Release_Ball"), PlungerReleaseSoundLevel, pl
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
'****  END FLEEP MECHANICAL SOUNDS
'******************************************************







'******************************************************
' ZPHY:  GNEREAL ADVICE ON PHYSICS
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

InitPolarity

'
''*******************************************
'' Late 70's to early 80's
'
'Sub InitPolarity()
'   dim x, a : a = Array(LF, RF)
' for each x in a
'   x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
'   x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
'   x.enabled = True
'   x.TimeDelay = 80
'   x.DebugOn=False ' prints some info in debugger
'
'
'        x.AddPt "Polarity", 0, 0, 0
'        x.AddPt "Polarity", 1, 0.05, - 2.7
'        x.AddPt "Polarity", 2, 0.16, - 2.7
'        x.AddPt "Polarity", 3, 0.22, - 0
'        x.AddPt "Polarity", 4, 0.25, - 0
'        x.AddPt "Polarity", 5, 0.3, - 1
'        x.AddPt "Polarity", 6, 0.4, - 2
'        x.AddPt "Polarity", 7, 0.5, - 2.7
'        x.AddPt "Polarity", 8, 0.65, - 1.8
'        x.AddPt "Polarity", 9, 0.75, - 0.5
'        x.AddPt "Polarity", 10, 0.81, - 0.5
'        x.AddPt "Polarity", 11, 0.88, 0
'        x.AddPt "Polarity", 12, 1.3, 0
'
'   x.AddPt "Velocity", 0, 0, 0.85
'   x.AddPt "Velocity", 1, 0.15, 0.85
'   x.AddPt "Velocity", 2, 0.2, 0.9
'   x.AddPt "Velocity", 3, 0.23, 0.95
'   x.AddPt "Velocity", 4, 0.41, 0.95
'   x.AddPt "Velocity", 5, 0.53, 0.95 '0.982
'   x.AddPt "Velocity", 6, 0.62, 1.0
'   x.AddPt "Velocity", 7, 0.702, 0.968
'   x.AddPt "Velocity", 8, 0.95,  0.968
'   x.AddPt "Velocity", 9, 1.03,  0.945
'   x.AddPt "Velocity", 10, 1.5,  0.945
'
' Next
'
' ' SetObjects arguments: 1: name of object 2: flipper object: 3: Trigger object around flipper
'    LF.SetObjects "LF", LeftFlipper, TriggerLF
'    RF.SetObjects "RF", RightFlipper, TriggerRF
'End Sub
'
'
'
''*******************************************
'' Mid 80's
'
'Sub InitPolarity()
'   dim x, a : a = Array(LF, RF)
' for each x in a
'   x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
'   x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
'   x.enabled = True
'   x.TimeDelay = 80
'   x.DebugOn=False ' prints some info in debugger
'
'   x.AddPt "Polarity", 0, 0, 0
'   x.AddPt "Polarity", 1, 0.05, - 3.7
'   x.AddPt "Polarity", 2, 0.16, - 3.7
'   x.AddPt "Polarity", 3, 0.22, - 0
'   x.AddPt "Polarity", 4, 0.25, - 0
'   x.AddPt "Polarity", 5, 0.3, - 2
'   x.AddPt "Polarity", 6, 0.4, - 3
'   x.AddPt "Polarity", 7, 0.5, - 3.7
'   x.AddPt "Polarity", 8, 0.65, - 2.3
'   x.AddPt "Polarity", 9, 0.75, - 1.5
'   x.AddPt "Polarity", 10, 0.81, - 1
'   x.AddPt "Polarity", 11, 0.88, 0
'   x.AddPt "Polarity", 12, 1.3, 0
'
'   x.AddPt "Velocity", 0, 0, 0.85
'   x.AddPt "Velocity", 1, 0.15, 0.85
'   x.AddPt "Velocity", 2, 0.2, 0.9
'   x.AddPt "Velocity", 3, 0.23, 0.95
'   x.AddPt "Velocity", 4, 0.41, 0.95
'   x.AddPt "Velocity", 5, 0.53, 0.95 '0.982
'   x.AddPt "Velocity", 6, 0.62, 1.0
'   x.AddPt "Velocity", 7, 0.702, 0.968
'   x.AddPt "Velocity", 8, 0.95,  0.968
'   x.AddPt "Velocity", 9, 1.03,  0.945
'   x.AddPt "Velocity", 10, 1.5,  0.945
'
' Next
'
' ' SetObjects arguments: 1: name of object 2: flipper object: 3: Trigger object around flipper
'    LF.SetObjects "LF", LeftFlipper, TriggerLF
'    RF.SetObjects "RF", RightFlipper, TriggerRF
'End Sub
'
''*******************************************
''  Late 80's early 90's
'
'Sub InitPolarity()
' dim x, a : a = Array(LF, RF)
' for each x in a
'   x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
'   x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
'   x.enabled = True
'   x.TimeDelay = 60
'   x.DebugOn=False ' prints some info in debugger
'
'   x.AddPt "Polarity", 0, 0, 0
'   x.AddPt "Polarity", 1, 0.05, - 5
'   x.AddPt "Polarity", 2, 0.16, - 5
'   x.AddPt "Polarity", 3, 0.22, - 0
'   x.AddPt "Polarity", 4, 0.25, - 0
'   x.AddPt "Polarity", 5, 0.3, - 2
'   x.AddPt "Polarity", 6, 0.4, - 3
'   x.AddPt "Polarity", 7, 0.5, - 4.0
'   x.AddPt "Polarity", 8, 0.7, - 3.5
'   x.AddPt "Polarity", 9, 0.75, - 3.0
'   x.AddPt "Polarity", 10, 0.8, - 2.5
'   x.AddPt "Polarity", 11, 0.85, - 2.0
'   x.AddPt "Polarity", 12, 0.9, - 1.5
'   x.AddPt "Polarity", 13, 0.95, - 1.0
'   x.AddPt "Polarity", 14, 1, - 0.5
'   x.AddPt "Polarity", 15, 1.1, 0
'   x.AddPt "Polarity", 16, 1.3, 0
'
'   x.AddPt "Velocity", 0, 0, 0.85
'   x.AddPt "Velocity", 1, 0.15, 0.85
'   x.AddPt "Velocity", 2, 0.2, 0.9
'   x.AddPt "Velocity", 3, 0.23, 0.95
'   x.AddPt "Velocity", 4, 0.41, 0.95
'   x.AddPt "Velocity", 5, 0.53, 0.95 '0.982
'   x.AddPt "Velocity", 6, 0.62, 1.0
'   x.AddPt "Velocity", 7, 0.702, 0.968
'   x.AddPt "Velocity", 8, 0.95,  0.968
'   x.AddPt "Velocity", 9, 1.03,  0.945
'   x.AddPt "Velocity", 10, 1.5,  0.945

' Next
'
' ' SetObjects arguments: 1: name of object 2: flipper object: 3: Trigger object around flipper
' LF.SetObjects "LF", LeftFlipper, TriggerLF
' RF.SetObjects "RF", RightFlipper, TriggerRF
'End Sub

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
'   Const EOSReturn = 0.035  'mid 80's to early 90's
Const EOSReturn = 0.025  'mid 90's and later

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


Dim DTArray

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
Dim ST16, ST25, ST28, ST37, ST66, ST67

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

' TODO: Uncomment once target prims exist

Set ST16 = (new StandupTarget)(sw16, BM_sw16, 16, 0)
Set ST25 = (new StandupTarget)(sw25, BM_sw25, 25, 0)
Set ST28 = (new StandupTarget)(sw28, BM_sw28, 28, 0)
Set ST37 = (new StandupTarget)(sw37, BM_sw37, 37, 0)
Set ST66 = (new StandupTarget)(sw66, BM_sw66, 66, 0)
Set ST67 = (new StandupTarget)(sw67, BM_sw67, 67, 0)

'Add all the Stand-up Target Arrays to Stand-up Target Animation Array
' STAnimationArray = Array(ST1, ST2, ....)
Dim STArray
STArray = Array(ST16, ST25, ST28, ST37, ST66, ST67)

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


Sub UpdateStandupTargets
  dim BP, ty

    ty = BM_sw16.transy
  For each BP in BP_sw16 : BP.transy = ty: Next

    ty = BM_sw25.transy
  For each BP in BP_sw25 : BP.transy = ty: Next

  ty = BM_sw28.transy
  For each BP in BP_sw28 : BP.transy = ty: Next

    ty = BM_sw37.transy
  For each BP in BP_sw37 : BP.transy = ty: Next

    ty = BM_sw67.transy
  For each BP in BP_sw67 : BP.transy = ty: Next

  ty = BM_sw66.transy
  For each BP in BP_sw66 : BP.transy = ty: Next

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






'****************************
' ZGIU: GI Updates
'****************************

Set GiCallback2 = GetRef("GIUpdate2")

dim gi0lvl:gi0lvl = 0
dim gi1lvl:gi1lvl = 0
dim gi2lvl:gi2lvl = 0


Sub GIUpdate2(no, pwm)
  'debug.print "GIUpdate2 no="&no&" level="&pwm
  Dim bulb
  Select Case no
    Case 0 :
      For each bulb in GILeft: bulb.State = pwm: Next
      'GIRelaySound pwm,gi0lvl
      gi0lvl = pwm
    Case 1 :
      For each bulb in GIRight: bulb.State = pwm: Next
      'GIRelaySound pwm,gi1lvl
      gi1lvl = pwm
    Case 4 :
      For each bulb in GITop: bulb.State = pwm: Next
      'GIRelaySound pwm,gi2lvl
      gi2lvl = pwm
  End Select
End Sub


Sub GIRelaySound(pwm, gilvl)
  If pwm >= 0.5 And gilvl < 0.5 Then
    Sound_GI_Relay 1, Bumper1
  ElseIf pwm <= 0.4 And gilvl > 0.4 Then
    Sound_GI_Relay 0, Bumper1
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
' ZVRS: VR Stuff
'******************************************************
Sub LoadVRRoom
  DIM VRThings
  for each VRThings in VRCab:VRThings.visible = 0:Next
  for each VRThings in VRMin:VRThings.visible = 0:Next
  for each VRThings in VRMega:VRThings.visible = 0:Next
  for each VRThings in VRBackglass:VRThings.visible = 0:Next
  VRCab_Rails.visible = CabRails
  BGDark.visible = 0
  TimerPlunger2.Enabled = True


  If RenderingMode = 2 or VRTest Then
    VRRoom = VRRoomChoice
    VRCab_Rails.visible = True 'VR
    VRCab_CustomShooter.Visible = CustomShooter
    TimerPlunger2.Enabled = True
  Else
    VRRoom = 0
  End If

  If VRRoom = 1 Then
    for each VRThings in VRCab:VRThings.visible = 1:Next
    for each VRThings in VRMin:VRThings.visible = 0:Next
    for each VRThings in VRMega:VRThings.visible = 1:Next
    for each VRThings in VRSphere:VRThings.visible = 0:Next
    for each VRThings in VRBackglass:VRThings.visible = 1:Next
    BGDark.visible = 1
  End If
  If VRRoom = 2 Then
    for each VRThings in VRCab:VRThings.visible = 1:Next
    for each VRThings in VRMin:VRThings.visible = 1:Next
    for each VRThings in VRMega:VRThings.visible = 0:Next
    for each VRThings in VRSphere:VRThings.visible = 0:Next
    for each VRThings in VRBackglass:VRThings.visible = 1:Next
    BGDark.visible = 1
  End If
  If VRRoom = 3 Then
    for each VRThings in VRCab:VRThings.visible = 1:Next
    for each VRThings in VRMin:VRThings.visible = 0:Next
    for each VRThings in VRMega:VRThings.visible = 0:Next
    for each VRThings in VRSphere:VRThings.visible = 0:Next
    for each VRThings in VRBackglass:VRThings.visible = 1:Next
    BGDark.visible = 1
  End If
  If VRRoom = 4 Then
    for each VRThings in VRCab:VRThings.visible = 1:Next
    for each VRThings in VRMin:VRThings.visible = 0:Next
    for each VRThings in VRMega:VRThings.visible = 0:Next
    for each VRThings in VRSphere:VRThings.visible = 1:Next
    for each VRThings in VRBackglass:VRThings.visible = 1:Next
    BGDark.visible = 1
  End If
End Sub

Sub SetBackglass()
    Dim obj

    For Each obj In VRBackglass
        'obj.x = obj.x + 3
        obj.height = - obj.y + 340
        obj.y = 80 'adjusts the distance from the backglass towards the user
        obj.rotx=-90
        obj.color = RGB(255,180,80)
    Next

    BGDark.height = - BGDark.y + 340
    BGDark.y = 65
    BGDark.rotx=-90

End Sub

Sub L88_animate
  If l88.state > 0.5 Then
    VRCab_StartButton2.disablelighting = 1
  Else
    VRCab_StartButton2.disablelighting = 0
  End if
End Sub

Sub L87_animate
  If l87.state > 0.5 Then
    VRCab_BuyInButton2.disablelighting = 1
  Else
    VRCab_BuyInButton2.disablelighting = 0
  End if
End Sub

' Changelog
' =========
' 0.001 - mcarter78 - Initial build - posts, walls, flippers, kickers, targets, all lights, initial scripting
' 0.002 - RobbyKingPin - Adjusted flipper dimensions and reshaped flippertriggers. Added rubber dampeners on posts, pegs and sleeves. Small corrections on nFozzy settings inside the script. Added Fleep parts into collections
' 0.003 - mcarter78 - Add plastic ramp prims & playfield mesh from tomate.  Adding ramp triggers WIP
' 0.004 - Sixtoe - Added skillshot primitive ramp, added physical wire ramp, redid playfield mesh, changed all kickers to physical, added goalie ramp, added goalie kicker and scoop, added blueprint image
' 0.005 - mcarter78 - Added and adjusted height for remaining ramp triggers.  Added skillshot triggers, WIP soccer ball & goalie mechs
' 0.006 - mcarter78 - Added ramp diverter & kickback plunger, minor scripting fixes
' 0.007 - tomate - ramp diverter modified sighly, some flashers adedd/moved to their right position, most of th gates still missing
' 0.008 - mcarter78 - Add lock area with magnet & switches, add missing ramp switches & missing gates
' 0.009 - tomate - most of the objects set as not visibles, first 10% test bumpy batch added
' 0.010 - tomate - new 2k batch adedd, color lut from flupper for AGX filter added
' 0.011 - Sixtoe - Turned a few more things non-visible, removed some redundant materials and images, rebuilt subway kicker so it works now
' 0.012 - mcarter78 - Set up GI, Animate flippers, slings, bumpers & rollovers
' 0.013 - tomate - Blender changes: GI setup, some changes to flasher and GI placement to match the real table, added flippers up position. VPX: new 2k batch added
' 0.014 - mcarter78 - Add animations for gates, ramp switches, soccer ball (wip), add new GI to collection, finish flipper animation
' 0.015 - mcarter78 - Add swPlunger for autoplunge, Add missing GI lights for bumpers
' 0.016 - tomate - blender fixes, unwrapped textures, new 2k batch added
' 0.017 - mcarter78 - Fix diverter operation and animation, add new right ramp mesh
' 0.018 - apophis - Soccerball animation updates. It is not all working yet.
' 0.019 - apophis - Soccerball animation working.
' 0.020 - mcarter78 - Implementation/Animation for goalie mech, finish standup targets implementation, add bottom sling posts to NoTargetBouncer
' 0.021 - mcarter78 - Fix goalie targets size/shape, add Y rotation to visual goalie to match real thing
' 0.022 - Sixtoe - Split skill ramp into frame and pins, changed pins to sleeve material, remade playfield mesh with physical button holes, added physical playfield buttons to physics layer (not hooked up to anything yet)
' 0.023 - mcarter78 - Hook up and animate playfield buttons, fix plunger alignment/strength, fix kickback visuals, tweak kicker angles, fix a couple stuck ball spots
' 0.024 - mcarter78 - Fix goalie walls (thanks DGrimmReaper, add wall sound to goalie hit, add metal sound to goal hit
' 0.025 - mcarter78 - Add wall between soccer ball and standup target to prevent stuck ball
' 0.026 - mcarter78 - Fix diverter orientation and power/hold logic
' 0.027 - mcarter78 - Another attempt to fix stuck ball at top lanes
' 0.028 - apophis - Updated bumper cap material and render probe.
' 0.029 - DGrimmReaper - Initial VR Setup (Not fully ready but playable)
' 0.030 - DGrimmReaper - VR Extra Ball button added, most buttons animations added (Left Magna and Extra Ball still needed).  Still to do: VRBackglass
' 0.031 - mcarter78 - Hook up left magnasave button and EB buy-in button (also allows player to use lockdown bar button for buy-in)
' 0.032 - DGrimmReaper - Thanks to DaRdog for HD Cab Art and Mega Room. Animated Buttons for Left Magna and Extra Ball.  Still to do: VRBackglass
' 0.033 - mcarter78 - Remove duplicate GT036 in GoalieWalls array
' 0.034 - mcarter78 - Fix positioning for gates, add under pf wall to prevent ball from sinking around soccer ball
' 0.035 - apophis - Reworked the goalie mech.
' 0.036 - Sixtoe - Experiment with table buttons, remade goalie ramp run and changed goalie subway, remade kickers, replaced playfield mesh with one I didn't mess up.
' 0.037 - DaRdog - Slimmed down 60mb
' 0.038 - apophis - Reworked the flipper angles and triggers. Fixed flipper shadows. Changed ball image. Added outlane difficulty option. Fixed GWalls_Hit.
' 0.039 - tomate - new 2k batch added, several fixes at blender side
' 0.040 - mcarter78 - Fix flippers position/angles & trigger shapes, add goal roof and ramp posts/pegs, enable mechanical plunger, add scoops to metals collection
' 0.041 - tomate - new 4k batch added, Pf set as hide part behind and layer0 set as not visible
' 0.042 - mcarter78 - Add sounds for soccer ball & goalie, fix lock magnet & inlane guides alignment, add KnockerPosition and enable Knocker, hook up and animate spinner, fix ramp end sounds, animate plunger, remove trough gate from sounds collection, tune kicker angles, animate diverter
' 0.043 - tomate - new 4k batch added, new soccer ball texture unwrapped, center target texture fixed, a few others improvements, desktop POV fixed
' 0.044 - Sixtoe - Added GI hack for left sling gi light, added physical deflector plates behind all kickers, rebuilt extra ball hole area, changed pop bumper size and position slightly, refactored most walls, added blocker areas, modified vuk scoop angle, changed physics material of all sleeves, added external faces to both ramp primitives, rebuilt lock area, added goalie net primitive, numerous other tweaks and fixes
' 0.045 - apophis - Added tint to bumper caps. Tweaked playfield button physics. Fixed goalie going into pf when moving. Wired up ball blurring effect. Fixed diverter animation. Added color to insert ball reflections. Removed reference blueprint image.
' 0.046 - apophis - Optimized ball blurring method.
' 0.047 - apophis - Improved blurring method (two blur images).
' 0.048 - tomate - Ball blur images changed
' 0.049 - Sixtoe - Changed throwin ramp primitive, changed ball physical object, removed collidable from ballblur prim, added collidable plastics, rebuilt goalie area AGAIN and adjusted VUK (should be fine now, might need a feed gate for multiple balls?), rebuilt lock area, fixed spinner direction, fixed some ball traps, adjusted vr cab and rails, automated rail visibility for VR and desktop, probably some other stuff I've forgotten.
' 0.050 - DGrimmReaper - VR Backglass bulbs and flashers added (thanks to Hauntfreaks for image and bulb/flasher placements)
' 0.051 - Sixtoe - Removed diverter bar protection island from lower ramp, converted diverters to primitives to stop them fouling lower ramps, tweaked free kick kicker, added some sound to the throw in ramp and area, hooked up trigger end top of upper ramp hole, added another option to ball buy in, rotated sling0 vlm's 180 degrees, made backbox rear solid, added another plastic protector.
' 0.052 - apophis - Refactored goalie mech so that it runs smooth. Optimized some frametimer stuff. Fix cab mode rails. Fixed custom shooter knob option. Enabled ball shadows near slings.
' 0.053 - DGrimmReaper - BP_Plunger animation for button press.  Pitch image reduced to 4000x4000 and VRCabinet adjusted for coin toss on front of cab.
' 0.054 - tomate - missing flasher light added, splitted GI lights reduces, flasher power reduced, missing GI light on the left slinshot added, gi013 light hack removed, bumper caps material changed, ramps material improved, coin toss plastic improved
' 0.055 - apophis - Fixed plunger animation (for cab). Minor color adjustment on flasher LMs. Fixed visual sling rubbers. Fixed bumper caps refractions. Added playfield reflections option.
' 0.056 - Sixtoe - Fixed some trigger heights, added blocker wall and extra switch for sw42 goalie vuk and logic to stop multiple balls blocking VUK, unified some of the walls because there's something wrong with me.
' RC1 - apophis - Tweaks to goalie animation. Added cab rail option.
' RC2 - apophis - Adjusted magnet strengths and trigger sizes.
' RC3 - apophis - Reworked kicker sw42 logic to be more robust (hopefully). Made lock magnet grab center. Updated FS POV.
' RC4 - apophis - Fortified ramp near lock pin. Fixed tilt switch assignment.
' RC5 - apophis - Fixed goalie switches (thanks Astro).
' Release 1.0
' 1.0.1 - apophis - Fixed ball motor sfx issue. Initialze slingshots properly.
' 1.0.2 - apophis - Added Narnia ball check
' 1.0.3 - Sixtoe - Fixed coin toss switches, fixed ball trap behind goal net, tested and tweaked goalie vuk, tweaked coin toss frame
' Release 1.1
' 1.1.2 - Sixtoe - Redid pf mesh a bit, tweaked some physics pbjects
' 1.1.3 - Sixtoe - Redid loads of physics items, integrated ext2k's changes, remade hinges and messed about with rails, might play harder now
' 1.1.4 - Ext2k - Updates to VR
' 1.1.5 - DGrimmReaper - Removed backbox reflections, reduced the cab image and converted to webp, removed old cabinet images no longer being used.
' 1.1.6 - apophis - Increased flipper strength slightly. Prep file for release.
' Release 1.2

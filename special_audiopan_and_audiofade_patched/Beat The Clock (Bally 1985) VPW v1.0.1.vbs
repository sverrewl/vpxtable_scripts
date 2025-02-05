'           ___  _______ ______  __  __         _______   ____  _______ __
'          / _ )/ __/ _ /_  __/ / /_/ /  ___   / ___/ /  / __ \/ ___/ //_/
'         / _  / _// __ |/ /   / __/ _ \/ -_) / /__/ /__/ /_/ / /__/ ,<
'        /____/___/_/ |_/_/    \__/_//_/\__/  \___/____/\____/\___/_/|_|
'
'
'                Bally Midway Manufacturing Company (1985)
'
'                https://www.ipdb.org/machine.cgi?id=212
'
'
' VPW TEAM
' Apophis - Project lead, 3D modeling/rendering, physics, sounds, table fixes, and more
' ArmyAviation - Playfield and plastics redraw, initial table build
' TastyWasps - VR backglass scripting
' Primetime5k - Extensive testing
' HauntFreaks - Backglass assets
' DarthVito - VR posters
' Thanks to Bord, Flux, and Tomate for providing blender models and support!
'
'
'
' PLEASE NOTE:
' - All game options can be adjusted by pressing F12 and then left flipper button during the game. Press start button to save your options.
'
'
'
'
'*********************************************************************************************************************************
' === TABLE OF CONTENTS  ===
'
' You can quickly jump to a section by searching the four letter tag (ZXXX)
'
'  ZCON: Constants and Global Variables
'  ZVLM: VLM Arrays
'  ZTIM: Timers
'  ZINI: Table Initialization and Exiting
'  ZOPT: User Options
'  ZBRI: Room Brightness
'  ZSOL: Solenoid Callbacks
'  ZFLP: Flippers
'  ZGII: GI
'  ZDRN: Drain, Trough, and Ball Release
'  ZKEY: Key Press Handling
'  ZTRI: Triggers
'  ZBMP: Bumpers
'  ZSLG: Slingshots
'  ZLED: Digital Display
'  ZBRL: Ball Rolling and Drop Sounds
'  ZSHA: Ambient ball shadows
'  ZNFF: Flipper Corrections
'  ZDMP: Rubber Dampeners
'  ZBOU: Target Bouncer
'  ZSSC: Slingshot Corrections
'  ZFLE: Fleep Mechanical Sounds
'  ZRDT: Roth Drop Targets
'  ZRST: Roth Stand-up Targets
'  ZANI: Misc Animations
'  ZVRR: Setup VR Room
'  ZVRB: VR Backglass
'  ZCHG: Change Log
'
'*********************************************************************************************************************************



Option Explicit
Randomize
SetLocale 1033




'*******************************************
'  ZCON: Constants and Global Variables
'*******************************************

Const BallSize = 50
Const BallMass = 1
Dim tablewidth: tablewidth = Table1.width
Dim tableheight: tableheight = Table1.height
Const tnob = 1 ' total number of balls
Const lob = 0 ' total number of locked balls
Const UseVPMDMD = True 'For tournament mode

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="beatclc2"  'Flasher support
'Const cGameName="beatclck"
Const UseSolenoids=2,UseLamps=1,UseGI=0,sCoin=""

LoadVPM "01120100", "BALLY.VBS", 3.02

Dim VRMode, DesktopMode: DesktopMode = Table1.ShowDT
Const ForceVR = False   'Forces VR room to be active


'*******************************************
' ZVLM: VLM Arrays
'*******************************************


' VLM Arrays - Start
' Arrays per baked part
Dim BP_BR1: BP_BR1=Array(BM_BR1, LM_GI_GI07_BR1, LM_GI_GI08_BR1)
Dim BP_BR2: BP_BR2=Array(BM_BR2)
Dim BP_BR3: BP_BR3=Array(BM_BR3)
Dim BP_BS1: BP_BS1=Array(BM_BS1, LM_FL_F6a_BS1, LM_FL_F6b_BS1, LM_GI_G13_BS1, LM_GI_GI00_BS1, LM_GI_GI07_BS1, LM_GI_GI08_BS1, LM_GI_GI09_BS1)
Dim BP_BS2: BP_BS2=Array(BM_BS2, LM_FL_F6a_BS2, LM_FL_F6b_BS2, LM_GI_G13_BS2, LM_GI_GI00_BS2, LM_GI_GI07_BS2, LM_GI_GI11_BS2)
Dim BP_BS3: BP_BS3=Array(BM_BS3, LM_GI_GI00_BS3, LM_GI_GI07_BS3, LM_GI_GI08_BS3, LM_GI_GI09_BS3)
Dim BP_Gate001: BP_Gate001=Array(BM_Gate001, LM_L_l92_Gate001)
Dim BP_Gate1: BP_Gate1=Array(BM_Gate1, LM_GI_GI02_Gate1, LM_GI_GI03_Gate1)
Dim BP_Gate2: BP_Gate2=Array(BM_Gate2, LM_FL_F6b_Gate2, LM_FL_F6d_Gate2, LM_GI_GI00_Gate2)
Dim BP_Gate3: BP_Gate3=Array(BM_Gate3, LM_GI_G13_Gate3, LM_GI_GI00_Gate3, LM_GI_GI11_Gate3)
Dim BP_Gate4: BP_Gate4=Array(BM_Gate4, LM_FL_F6a_Gate4, LM_FL_F6c_Gate4)
Dim BP_LFlipper: BP_LFlipper=Array(BM_LFlipper, LM_GI_GI01_LFlipper, LM_GI_GI02_LFlipper, LM_GI_GI03_LFlipper, LM_L_l62_LFlipper)
Dim BP_LFlipperU: BP_LFlipperU=Array(BM_LFlipperU, LM_GI_GI01_LFlipperU, LM_GI_GI02_LFlipperU, LM_GI_GI03_LFlipperU, LM_L_l04_LFlipperU, LM_L_l35_LFlipperU, LM_L_l36_LFlipperU, LM_L_l62_LFlipperU)
Dim BP_LSling1: BP_LSling1=Array(BM_LSling1, LM_GI_GI01_LSling1, LM_GI_GI03_LSling1, LM_GI_GI04_LSling1)
Dim BP_LSling2: BP_LSling2=Array(BM_LSling2, LM_GI_GI01_LSling2, LM_GI_GI03_LSling2, LM_GI_GI04_LSling2)
Dim BP_LSlingArm: BP_LSlingArm=Array(BM_LSlingArm, LM_GI_GI03_LSlingArm)
Dim BP_Layer1: BP_Layer1=Array(BM_Layer1, LM_FL_F6a_Layer1, LM_FL_F6b_Layer1, LM_FL_F6c_Layer1, LM_FL_F6d_Layer1, LM_GI_G13_Layer1, LM_GI_GI00_Layer1, LM_GI_GI07_Layer1, LM_GI_GI08_Layer1, LM_GI_GI09_Layer1, LM_GI_GI11_Layer1, LM_L_l16_Layer1, LM_L_l49_Layer1, LM_L_l50_Layer1, LM_L_l59_Layer1, LM_L_l63_Layer1, LM_L_l65_Layer1, LM_L_l75_Layer1, LM_L_l78_Layer1, LM_L_l79_Layer1, LM_L_l81_Layer1, LM_L_l83_Layer1, LM_L_l91_Layer1, LM_L_l92_Layer1, LM_L_l94_Layer1, LM_L_l95_Layer1)
Dim BP_Parts: BP_Parts=Array(BM_Parts, LM_FL_F6a_Parts, LM_FL_F6b_Parts, LM_FL_F6c_Parts, LM_FL_F6d_Parts, LM_GI_G13_Parts, LM_GI_GI01_Parts, LM_GI_GI02_Parts, LM_GI_GI03_Parts, LM_GI_GI04_Parts, LM_GI_GI00_Parts, LM_GI_GI07_Parts, LM_GI_GI08_Parts, LM_GI_GI09_Parts, LM_GI_GI11_Parts, LM_L_L02_Parts, LM_L_l04_Parts, LM_L_l05_Parts, LM_L_l14_Parts, LM_L_l16_Parts, LM_L_l20_Parts, LM_L_l35_Parts, LM_L_l36_Parts, LM_L_l45_Parts, LM_L_l49_Parts, LM_L_l50_Parts, LM_L_l51_Parts, LM_L_l52_Parts, LM_L_l53_Parts, LM_L_l54_Parts, LM_L_l58_Parts, LM_L_l59_Parts, LM_L_l60_Parts, LM_L_l62_Parts, LM_L_l63_Parts, LM_L_l65_Parts, LM_L_l68_Parts, LM_L_l74_Parts, LM_L_l75_Parts, LM_L_l76_Parts, LM_L_l78_Parts, LM_L_l79_Parts, LM_L_l81_Parts, LM_L_l83_Parts, LM_L_l84_Parts, LM_L_l85_Parts, LM_L_l90_Parts, LM_L_l91_Parts, LM_L_l92_Parts, LM_L_l94_Parts, LM_L_l95_Parts)
Dim BP_PincabRails: BP_PincabRails=Array(BM_PincabRails)
Dim BP_Playfield: BP_Playfield=Array(BM_Playfield, LM_FL_F6a_Playfield, LM_FL_F6b_Playfield, LM_FL_F6c_Playfield, LM_FL_F6d_Playfield, LM_GI_G13_Playfield, LM_GI_GI01_Playfield, LM_GI_GI02_Playfield, LM_GI_GI03_Playfield, LM_GI_GI04_Playfield, LM_GI_GI00_Playfield, LM_GI_GI07_Playfield, LM_GI_GI08_Playfield, LM_GI_GI09_Playfield, LM_GI_GI11_Playfield, LM_L_L02_Playfield, LM_L_l04_Playfield, LM_L_l05_Playfield, LM_L_l06_Playfield, LM_L_l07_Playfield, LM_L_l08_Playfield, LM_L_l09_Playfield, LM_L_l10_Playfield, LM_L_l11_Playfield, LM_L_l12_Playfield, LM_L_l13_Playfield, LM_L_l14_Playfield, LM_L_l15_Playfield, LM_L_l16_Playfield, LM_L_l20_Playfield, LM_L_l21_Playfield, LM_L_l22_Playfield, LM_L_l23_Playfield, LM_L_l24_Playfield, LM_L_l25_Playfield, LM_L_l26_Playfield, LM_L_l27_Playfield, LM_L_l28_Playfield, LM_L_l29_Playfield, LM_L_l30_Playfield, LM_L_l31_Playfield, LM_L_l35_Playfield, LM_L_l36_Playfield, LM_L_l37_Playfield, LM_L_l38_Playfield, LM_L_l39_Playfield, LM_L_l40_Playfield, LM_L_l41_Playfield, _
  LM_L_l42_Playfield, LM_L_l43_Playfield, LM_L_l44_Playfield, LM_L_l45_Playfield, LM_L_l46_Playfield, LM_L_l47_Playfield, LM_L_l49_Playfield, LM_L_l50_Playfield, LM_L_l51_Playfield, LM_L_l52_Playfield, LM_L_l53_Playfield, LM_L_l54_Playfield, LM_L_l58_Playfield, LM_L_l59_Playfield, LM_L_l60_Playfield, LM_L_l62_Playfield, LM_L_l63_Playfield, LM_L_l65_Playfield, LM_L_l66_Playfield, LM_L_l67_Playfield, LM_L_l68_Playfield, LM_L_l69_Playfield, LM_L_l74_Playfield, LM_L_l75_Playfield, LM_L_l76_Playfield, LM_L_l77_Playfield, LM_L_l78_Playfield, LM_L_l79_Playfield, LM_L_l81_Playfield, LM_L_l82_Playfield, LM_L_l83_Playfield, LM_L_l84_Playfield, LM_L_l85_Playfield, LM_L_l89_Playfield, LM_L_l90_Playfield, LM_L_l91_Playfield, LM_L_l92_Playfield, LM_L_l93_Playfield, LM_L_l94_Playfield, LM_L_l95_Playfield)
Dim BP_RFlipper: BP_RFlipper=Array(BM_RFlipper, LM_GI_GI01_RFlipper, LM_GI_GI02_RFlipper, LM_GI_GI04_RFlipper, LM_L_l62_RFlipper)
Dim BP_RFlipper1: BP_RFlipper1=Array(BM_RFlipper1, LM_GI_GI04_RFlipper1, LM_GI_GI00_RFlipper1, LM_L_l89_RFlipper1, LM_L_l90_RFlipper1)
Dim BP_RFlipper1U: BP_RFlipper1U=Array(BM_RFlipper1U, LM_GI_GI04_RFlipper1U, LM_GI_GI00_RFlipper1U, LM_L_l89_RFlipper1U)
Dim BP_RFlipperU: BP_RFlipperU=Array(BM_RFlipperU, LM_GI_GI01_RFlipperU, LM_GI_GI02_RFlipperU, LM_GI_GI04_RFlipperU, LM_L_l05_RFlipperU, LM_L_l20_RFlipperU, LM_L_l35_RFlipperU, LM_L_l62_RFlipperU)
Dim BP_RSling1: BP_RSling1=Array(BM_RSling1, LM_GI_GI01_RSling1, LM_GI_GI03_RSling1, LM_GI_GI04_RSling1, LM_GI_GI00_RSling1, LM_L_l93_RSling1)
Dim BP_RSling2: BP_RSling2=Array(BM_RSling2, LM_GI_GI01_RSling2, LM_GI_GI03_RSling2, LM_GI_GI04_RSling2, LM_GI_GI00_RSling2, LM_L_l93_RSling2)
Dim BP_RSlingArm: BP_RSlingArm=Array(BM_RSlingArm, LM_GI_GI01_RSlingArm, LM_GI_GI03_RSlingArm, LM_GI_GI04_RSlingArm, LM_GI_GI00_RSlingArm)
Dim BP_UnderPF: BP_UnderPF=Array(BM_UnderPF, LM_FL_F6a_UnderPF, LM_FL_F6b_UnderPF, LM_FL_F6c_UnderPF, LM_FL_F6d_UnderPF, LM_GI_G13_UnderPF, LM_GI_GI01_UnderPF, LM_GI_GI02_UnderPF, LM_GI_GI03_UnderPF, LM_GI_GI04_UnderPF, LM_GI_GI00_UnderPF, LM_GI_GI07_UnderPF, LM_GI_GI08_UnderPF, LM_GI_GI09_UnderPF, LM_GI_GI11_UnderPF, LM_L_L02_UnderPF, LM_L_l04_UnderPF, LM_L_l05_UnderPF, LM_L_l06_UnderPF, LM_L_l07_UnderPF, LM_L_l08_UnderPF, LM_L_l09_UnderPF, LM_L_l10_UnderPF, LM_L_l11_UnderPF, LM_L_l12_UnderPF, LM_L_l13_UnderPF, LM_L_l14_UnderPF, LM_L_l15_UnderPF, LM_L_l16_UnderPF, LM_L_l20_UnderPF, LM_L_l21_UnderPF, LM_L_l22_UnderPF, LM_L_l23_UnderPF, LM_L_l24_UnderPF, LM_L_l25_UnderPF, LM_L_l26_UnderPF, LM_L_l27_UnderPF, LM_L_l28_UnderPF, LM_L_l29_UnderPF, LM_L_l30_UnderPF, LM_L_l31_UnderPF, LM_L_l35_UnderPF, LM_L_l36_UnderPF, LM_L_l37_UnderPF, LM_L_l38_UnderPF, LM_L_l39_UnderPF, LM_L_l40_UnderPF, LM_L_l41_UnderPF, LM_L_l42_UnderPF, LM_L_l43_UnderPF, LM_L_l44_UnderPF, LM_L_l45_UnderPF, LM_L_l46_UnderPF, LM_L_l47_UnderPF, _
  LM_L_l49_UnderPF, LM_L_l50_UnderPF, LM_L_l51_UnderPF, LM_L_l52_UnderPF, LM_L_l53_UnderPF, LM_L_l54_UnderPF, LM_L_l58_UnderPF, LM_L_l59_UnderPF, LM_L_l60_UnderPF, LM_L_l62_UnderPF, LM_L_l63_UnderPF, LM_L_l65_UnderPF, LM_L_l66_UnderPF, LM_L_l67_UnderPF, LM_L_l68_UnderPF, LM_L_l69_UnderPF, LM_L_l74_UnderPF, LM_L_l75_UnderPF, LM_L_l76_UnderPF, LM_L_l77_UnderPF, LM_L_l78_UnderPF, LM_L_l79_UnderPF, LM_L_l81_UnderPF, LM_L_l82_UnderPF, LM_L_l83_UnderPF, LM_L_l84_UnderPF, LM_L_l85_UnderPF, LM_L_l89_UnderPF, LM_L_l90_UnderPF, LM_L_l91_UnderPF, LM_L_l92_UnderPF, LM_L_l93_UnderPF, LM_L_l94_UnderPF, LM_L_l95_UnderPF)
Dim BP_sw1: BP_sw1=Array(BM_sw1, LM_GI_GI00_sw1)
Dim BP_sw13: BP_sw13=Array(BM_sw13, LM_GI_GI02_sw13, LM_L_l45_sw13)
Dim BP_sw16: BP_sw16=Array(BM_sw16, LM_GI_GI01_sw16, LM_L_l14_sw16)
Dim BP_sw17: BP_sw17=Array(BM_sw17, LM_GI_GI00_sw17, LM_GI_GI07_sw17, LM_GI_GI08_sw17)
Dim BP_sw18: BP_sw18=Array(BM_sw18, LM_GI_GI00_sw18, LM_GI_GI07_sw18, LM_GI_GI08_sw18)
Dim BP_sw19: BP_sw19=Array(BM_sw19, LM_GI_GI00_sw19, LM_GI_GI07_sw19, LM_GI_GI08_sw19, LM_GI_GI09_sw19)
Dim BP_sw2: BP_sw2=Array(BM_sw2, LM_GI_GI00_sw2)
Dim BP_sw20: BP_sw20=Array(BM_sw20, LM_GI_GI00_sw20, LM_GI_GI07_sw20, LM_GI_GI08_sw20, LM_GI_GI09_sw20)
Dim BP_sw21: BP_sw21=Array(BM_sw21, LM_FL_F6a_sw21, LM_GI_GI00_sw21, LM_GI_GI07_sw21, LM_GI_GI08_sw21, LM_GI_GI09_sw21, LM_GI_GI11_sw21)
Dim BP_sw23: BP_sw23=Array(BM_sw23, LM_L_l74_sw23)
Dim BP_sw24: BP_sw24=Array(BM_sw24, LM_GI_GI01_sw24, LM_L_l90_sw24)
Dim BP_sw3: BP_sw3=Array(BM_sw3, LM_GI_GI00_sw3, LM_GI_GI07_sw3)
Dim BP_sw30: BP_sw30=Array(BM_sw30, LM_GI_GI01_sw30, LM_GI_GI03_sw30, LM_GI_GI04_sw30, LM_GI_GI00_sw30, LM_L_l10_sw30, LM_L_l89_sw30)
Dim BP_sw31: BP_sw31=Array(BM_sw31, LM_GI_GI01_sw31, LM_GI_GI03_sw31, LM_GI_GI04_sw31, LM_GI_GI00_sw31, LM_L_l58_sw31, LM_L_l74_sw31)
Dim BP_sw32: BP_sw32=Array(BM_sw32, LM_FL_F6b_sw32, LM_GI_G13_sw32, LM_L_l16_sw32, LM_L_l95_sw32)
Dim BP_sw33: BP_sw33=Array(BM_sw33, LM_FL_F6a_sw33, LM_FL_F6b_sw33, LM_GI_G13_sw33, LM_GI_GI00_sw33, LM_GI_GI08_sw33, LM_GI_GI11_sw33, LM_L_l49_sw33, LM_L_l95_sw33)
Dim BP_sw34: BP_sw34=Array(BM_sw34, LM_FL_F6b_sw34, LM_GI_G13_sw34, LM_L_l49_sw34, LM_L_l65_sw34)
Dim BP_sw35: BP_sw35=Array(BM_sw35, LM_FL_F6b_sw35, LM_GI_G13_sw35, LM_GI_GI00_sw35, LM_L_l65_sw35, LM_L_l81_sw35)
Dim BP_sw36: BP_sw36=Array(BM_sw36, LM_FL_F6b_sw36, LM_GI_G13_sw36, LM_GI_GI00_sw36, LM_GI_GI07_sw36, LM_L_l50_sw36, LM_L_l65_sw36, LM_L_l81_sw36)
Dim BP_sw37: BP_sw37=Array(BM_sw37, LM_FL_F6b_sw37, LM_GI_G13_sw37, LM_L_l92_sw37)
Dim BP_sw38: BP_sw38=Array(BM_sw38, LM_FL_F6b_sw38, LM_GI_G13_sw38, LM_L_l76_sw38)
Dim BP_sw39: BP_sw39=Array(BM_sw39, LM_FL_F6a_sw39, LM_GI_GI11_sw39, LM_L_l60_sw39)
Dim BP_sw4: BP_sw4=Array(BM_sw4, LM_GI_GI00_sw4, LM_GI_GI07_sw4)
Dim BP_sw40: BP_sw40=Array(BM_sw40, LM_FL_F6a_sw40, LM_GI_GI07_sw40, LM_GI_GI11_sw40, LM_L_l91_sw40)
Dim BP_sw45: BP_sw45=Array(BM_sw45, LM_GI_GI00_sw45, LM_L_l78_sw45, LM_L_l95_sw45)
Dim BP_sw5: BP_sw5=Array(BM_sw5, LM_GI_GI00_sw5, LM_GI_GI07_sw5, LM_GI_GI08_sw5)
Dim BP_sw7: BP_sw7=Array(BM_sw7, LM_GI_GI00_sw7, LM_GI_GI07_sw7, LM_GI_GI08_sw7, LM_GI_GI09_sw7)
' Arrays per lighting scenario
Dim BL_FL_F6a: BL_FL_F6a=Array(LM_FL_F6a_BS1, LM_FL_F6a_BS2, LM_FL_F6a_Gate4, LM_FL_F6a_Layer1, LM_FL_F6a_Parts, LM_FL_F6a_Playfield, LM_FL_F6a_UnderPF, LM_FL_F6a_sw21, LM_FL_F6a_sw33, LM_FL_F6a_sw39, LM_FL_F6a_sw40)
Dim BL_FL_F6b: BL_FL_F6b=Array(LM_FL_F6b_BS1, LM_FL_F6b_BS2, LM_FL_F6b_Gate2, LM_FL_F6b_Layer1, LM_FL_F6b_Parts, LM_FL_F6b_Playfield, LM_FL_F6b_UnderPF, LM_FL_F6b_sw32, LM_FL_F6b_sw33, LM_FL_F6b_sw34, LM_FL_F6b_sw35, LM_FL_F6b_sw36, LM_FL_F6b_sw37, LM_FL_F6b_sw38)
Dim BL_FL_F6c: BL_FL_F6c=Array(LM_FL_F6c_Gate4, LM_FL_F6c_Layer1, LM_FL_F6c_Parts, LM_FL_F6c_Playfield, LM_FL_F6c_UnderPF)
Dim BL_FL_F6d: BL_FL_F6d=Array(LM_FL_F6d_Gate2, LM_FL_F6d_Layer1, LM_FL_F6d_Parts, LM_FL_F6d_Playfield, LM_FL_F6d_UnderPF)
Dim BL_GI_G13: BL_GI_G13=Array(LM_GI_G13_BS1, LM_GI_G13_BS2, LM_GI_G13_Gate3, LM_GI_G13_Layer1, LM_GI_G13_Parts, LM_GI_G13_Playfield, LM_GI_G13_UnderPF, LM_GI_G13_sw32, LM_GI_G13_sw33, LM_GI_G13_sw34, LM_GI_G13_sw35, LM_GI_G13_sw36, LM_GI_G13_sw37, LM_GI_G13_sw38)
Dim BL_GI_GI00: BL_GI_GI00=Array(LM_GI_GI00_BS1, LM_GI_GI00_BS2, LM_GI_GI00_BS3, LM_GI_GI00_Gate2, LM_GI_GI00_Gate3, LM_GI_GI00_Layer1, LM_GI_GI00_Parts, LM_GI_GI00_Playfield, LM_GI_GI00_RFlipper1, LM_GI_GI00_RFlipper1U, LM_GI_GI00_RSling1, LM_GI_GI00_RSling2, LM_GI_GI00_RSlingArm, LM_GI_GI00_UnderPF, LM_GI_GI00_sw1, LM_GI_GI00_sw17, LM_GI_GI00_sw18, LM_GI_GI00_sw19, LM_GI_GI00_sw2, LM_GI_GI00_sw20, LM_GI_GI00_sw21, LM_GI_GI00_sw3, LM_GI_GI00_sw30, LM_GI_GI00_sw31, LM_GI_GI00_sw33, LM_GI_GI00_sw35, LM_GI_GI00_sw36, LM_GI_GI00_sw4, LM_GI_GI00_sw45, LM_GI_GI00_sw5, LM_GI_GI00_sw7)
Dim BL_GI_GI01: BL_GI_GI01=Array(LM_GI_GI01_LFlipper, LM_GI_GI01_LFlipperU, LM_GI_GI01_LSling1, LM_GI_GI01_LSling2, LM_GI_GI01_Parts, LM_GI_GI01_Playfield, LM_GI_GI01_RFlipper, LM_GI_GI01_RFlipperU, LM_GI_GI01_RSling1, LM_GI_GI01_RSling2, LM_GI_GI01_RSlingArm, LM_GI_GI01_UnderPF, LM_GI_GI01_sw16, LM_GI_GI01_sw24, LM_GI_GI01_sw30, LM_GI_GI01_sw31)
Dim BL_GI_GI02: BL_GI_GI02=Array(LM_GI_GI02_Gate1, LM_GI_GI02_LFlipper, LM_GI_GI02_LFlipperU, LM_GI_GI02_Parts, LM_GI_GI02_Playfield, LM_GI_GI02_RFlipper, LM_GI_GI02_RFlipperU, LM_GI_GI02_UnderPF, LM_GI_GI02_sw13)
Dim BL_GI_GI03: BL_GI_GI03=Array(LM_GI_GI03_Gate1, LM_GI_GI03_LFlipper, LM_GI_GI03_LFlipperU, LM_GI_GI03_LSling1, LM_GI_GI03_LSling2, LM_GI_GI03_LSlingArm, LM_GI_GI03_Parts, LM_GI_GI03_Playfield, LM_GI_GI03_RSling1, LM_GI_GI03_RSling2, LM_GI_GI03_RSlingArm, LM_GI_GI03_UnderPF, LM_GI_GI03_sw30, LM_GI_GI03_sw31)
Dim BL_GI_GI04: BL_GI_GI04=Array(LM_GI_GI04_LSling1, LM_GI_GI04_LSling2, LM_GI_GI04_Parts, LM_GI_GI04_Playfield, LM_GI_GI04_RFlipper, LM_GI_GI04_RFlipper1, LM_GI_GI04_RFlipper1U, LM_GI_GI04_RFlipperU, LM_GI_GI04_RSling1, LM_GI_GI04_RSling2, LM_GI_GI04_RSlingArm, LM_GI_GI04_UnderPF, LM_GI_GI04_sw30, LM_GI_GI04_sw31)
Dim BL_GI_GI07: BL_GI_GI07=Array(LM_GI_GI07_BR1, LM_GI_GI07_BS1, LM_GI_GI07_BS2, LM_GI_GI07_BS3, LM_GI_GI07_Layer1, LM_GI_GI07_Parts, LM_GI_GI07_Playfield, LM_GI_GI07_UnderPF, LM_GI_GI07_sw17, LM_GI_GI07_sw18, LM_GI_GI07_sw19, LM_GI_GI07_sw20, LM_GI_GI07_sw21, LM_GI_GI07_sw3, LM_GI_GI07_sw36, LM_GI_GI07_sw4, LM_GI_GI07_sw40, LM_GI_GI07_sw5, LM_GI_GI07_sw7)
Dim BL_GI_GI08: BL_GI_GI08=Array(LM_GI_GI08_BR1, LM_GI_GI08_BS1, LM_GI_GI08_BS3, LM_GI_GI08_Layer1, LM_GI_GI08_Parts, LM_GI_GI08_Playfield, LM_GI_GI08_UnderPF, LM_GI_GI08_sw17, LM_GI_GI08_sw18, LM_GI_GI08_sw19, LM_GI_GI08_sw20, LM_GI_GI08_sw21, LM_GI_GI08_sw33, LM_GI_GI08_sw5, LM_GI_GI08_sw7)
Dim BL_GI_GI09: BL_GI_GI09=Array(LM_GI_GI09_BS1, LM_GI_GI09_BS3, LM_GI_GI09_Layer1, LM_GI_GI09_Parts, LM_GI_GI09_Playfield, LM_GI_GI09_UnderPF, LM_GI_GI09_sw19, LM_GI_GI09_sw20, LM_GI_GI09_sw21, LM_GI_GI09_sw7)
Dim BL_GI_GI11: BL_GI_GI11=Array(LM_GI_GI11_BS2, LM_GI_GI11_Gate3, LM_GI_GI11_Layer1, LM_GI_GI11_Parts, LM_GI_GI11_Playfield, LM_GI_GI11_UnderPF, LM_GI_GI11_sw21, LM_GI_GI11_sw33, LM_GI_GI11_sw39, LM_GI_GI11_sw40)
Dim BL_L_L02: BL_L_L02=Array(LM_L_L02_Parts, LM_L_L02_Playfield, LM_L_L02_UnderPF)
Dim BL_L_l04: BL_L_l04=Array(LM_L_l04_LFlipperU, LM_L_l04_Parts, LM_L_l04_Playfield, LM_L_l04_UnderPF)
Dim BL_L_l05: BL_L_l05=Array(LM_L_l05_Parts, LM_L_l05_Playfield, LM_L_l05_RFlipperU, LM_L_l05_UnderPF)
Dim BL_L_l06: BL_L_l06=Array(LM_L_l06_Playfield, LM_L_l06_UnderPF)
Dim BL_L_l07: BL_L_l07=Array(LM_L_l07_Playfield, LM_L_l07_UnderPF)
Dim BL_L_l08: BL_L_l08=Array(LM_L_l08_Playfield, LM_L_l08_UnderPF)
Dim BL_L_l09: BL_L_l09=Array(LM_L_l09_Playfield, LM_L_l09_UnderPF)
Dim BL_L_l10: BL_L_l10=Array(LM_L_l10_Playfield, LM_L_l10_UnderPF, LM_L_l10_sw30)
Dim BL_L_l11: BL_L_l11=Array(LM_L_l11_Playfield, LM_L_l11_UnderPF)
Dim BL_L_l12: BL_L_l12=Array(LM_L_l12_Playfield, LM_L_l12_UnderPF)
Dim BL_L_l13: BL_L_l13=Array(LM_L_l13_Playfield, LM_L_l13_UnderPF)
Dim BL_L_l14: BL_L_l14=Array(LM_L_l14_Parts, LM_L_l14_Playfield, LM_L_l14_UnderPF, LM_L_l14_sw16)
Dim BL_L_l15: BL_L_l15=Array(LM_L_l15_Playfield, LM_L_l15_UnderPF)
Dim BL_L_l16: BL_L_l16=Array(LM_L_l16_Layer1, LM_L_l16_Parts, LM_L_l16_Playfield, LM_L_l16_UnderPF, LM_L_l16_sw32)
Dim BL_L_l20: BL_L_l20=Array(LM_L_l20_Parts, LM_L_l20_Playfield, LM_L_l20_RFlipperU, LM_L_l20_UnderPF)
Dim BL_L_l21: BL_L_l21=Array(LM_L_l21_Playfield, LM_L_l21_UnderPF)
Dim BL_L_l22: BL_L_l22=Array(LM_L_l22_Playfield, LM_L_l22_UnderPF)
Dim BL_L_l23: BL_L_l23=Array(LM_L_l23_Playfield, LM_L_l23_UnderPF)
Dim BL_L_l24: BL_L_l24=Array(LM_L_l24_Playfield, LM_L_l24_UnderPF)
Dim BL_L_l25: BL_L_l25=Array(LM_L_l25_Playfield, LM_L_l25_UnderPF)
Dim BL_L_l26: BL_L_l26=Array(LM_L_l26_Playfield, LM_L_l26_UnderPF)
Dim BL_L_l27: BL_L_l27=Array(LM_L_l27_Playfield, LM_L_l27_UnderPF)
Dim BL_L_l28: BL_L_l28=Array(LM_L_l28_Playfield, LM_L_l28_UnderPF)
Dim BL_L_l29: BL_L_l29=Array(LM_L_l29_Playfield, LM_L_l29_UnderPF)
Dim BL_L_l30: BL_L_l30=Array(LM_L_l30_Playfield, LM_L_l30_UnderPF)
Dim BL_L_l31: BL_L_l31=Array(LM_L_l31_Playfield, LM_L_l31_UnderPF)
Dim BL_L_l35: BL_L_l35=Array(LM_L_l35_LFlipperU, LM_L_l35_Parts, LM_L_l35_Playfield, LM_L_l35_RFlipperU, LM_L_l35_UnderPF)
Dim BL_L_l36: BL_L_l36=Array(LM_L_l36_LFlipperU, LM_L_l36_Parts, LM_L_l36_Playfield, LM_L_l36_UnderPF)
Dim BL_L_l37: BL_L_l37=Array(LM_L_l37_Playfield, LM_L_l37_UnderPF)
Dim BL_L_l38: BL_L_l38=Array(LM_L_l38_Playfield, LM_L_l38_UnderPF)
Dim BL_L_l39: BL_L_l39=Array(LM_L_l39_Playfield, LM_L_l39_UnderPF)
Dim BL_L_l40: BL_L_l40=Array(LM_L_l40_Playfield, LM_L_l40_UnderPF)
Dim BL_L_l41: BL_L_l41=Array(LM_L_l41_Playfield, LM_L_l41_UnderPF)
Dim BL_L_l42: BL_L_l42=Array(LM_L_l42_Playfield, LM_L_l42_UnderPF)
Dim BL_L_l43: BL_L_l43=Array(LM_L_l43_Playfield, LM_L_l43_UnderPF)
Dim BL_L_l44: BL_L_l44=Array(LM_L_l44_Playfield, LM_L_l44_UnderPF)
Dim BL_L_l45: BL_L_l45=Array(LM_L_l45_Parts, LM_L_l45_Playfield, LM_L_l45_UnderPF, LM_L_l45_sw13)
Dim BL_L_l46: BL_L_l46=Array(LM_L_l46_Playfield, LM_L_l46_UnderPF)
Dim BL_L_l47: BL_L_l47=Array(LM_L_l47_Playfield, LM_L_l47_UnderPF)
Dim BL_L_l49: BL_L_l49=Array(LM_L_l49_Layer1, LM_L_l49_Parts, LM_L_l49_Playfield, LM_L_l49_UnderPF, LM_L_l49_sw33, LM_L_l49_sw34)
Dim BL_L_l50: BL_L_l50=Array(LM_L_l50_Layer1, LM_L_l50_Parts, LM_L_l50_Playfield, LM_L_l50_UnderPF, LM_L_l50_sw36)
Dim BL_L_l51: BL_L_l51=Array(LM_L_l51_Parts, LM_L_l51_Playfield, LM_L_l51_UnderPF)
Dim BL_L_l52: BL_L_l52=Array(LM_L_l52_Parts, LM_L_l52_Playfield, LM_L_l52_UnderPF)
Dim BL_L_l53: BL_L_l53=Array(LM_L_l53_Parts, LM_L_l53_Playfield, LM_L_l53_UnderPF)
Dim BL_L_l54: BL_L_l54=Array(LM_L_l54_Parts, LM_L_l54_Playfield, LM_L_l54_UnderPF)
Dim BL_L_l58: BL_L_l58=Array(LM_L_l58_Parts, LM_L_l58_Playfield, LM_L_l58_UnderPF, LM_L_l58_sw31)
Dim BL_L_l59: BL_L_l59=Array(LM_L_l59_Layer1, LM_L_l59_Parts, LM_L_l59_Playfield, LM_L_l59_UnderPF)
Dim BL_L_l60: BL_L_l60=Array(LM_L_l60_Parts, LM_L_l60_Playfield, LM_L_l60_UnderPF, LM_L_l60_sw39)
Dim BL_L_l62: BL_L_l62=Array(LM_L_l62_LFlipper, LM_L_l62_LFlipperU, LM_L_l62_Parts, LM_L_l62_Playfield, LM_L_l62_RFlipper, LM_L_l62_RFlipperU, LM_L_l62_UnderPF)
Dim BL_L_l63: BL_L_l63=Array(LM_L_l63_Layer1, LM_L_l63_Parts, LM_L_l63_Playfield, LM_L_l63_UnderPF)
Dim BL_L_l65: BL_L_l65=Array(LM_L_l65_Layer1, LM_L_l65_Parts, LM_L_l65_Playfield, LM_L_l65_UnderPF, LM_L_l65_sw34, LM_L_l65_sw35, LM_L_l65_sw36)
Dim BL_L_l66: BL_L_l66=Array(LM_L_l66_Playfield, LM_L_l66_UnderPF)
Dim BL_L_l67: BL_L_l67=Array(LM_L_l67_Playfield, LM_L_l67_UnderPF)
Dim BL_L_l68: BL_L_l68=Array(LM_L_l68_Parts, LM_L_l68_Playfield, LM_L_l68_UnderPF)
Dim BL_L_l69: BL_L_l69=Array(LM_L_l69_Playfield, LM_L_l69_UnderPF)
Dim BL_L_l74: BL_L_l74=Array(LM_L_l74_Parts, LM_L_l74_Playfield, LM_L_l74_UnderPF, LM_L_l74_sw23, LM_L_l74_sw31)
Dim BL_L_l75: BL_L_l75=Array(LM_L_l75_Layer1, LM_L_l75_Parts, LM_L_l75_Playfield, LM_L_l75_UnderPF)
Dim BL_L_l76: BL_L_l76=Array(LM_L_l76_Parts, LM_L_l76_Playfield, LM_L_l76_UnderPF, LM_L_l76_sw38)
Dim BL_L_l77: BL_L_l77=Array(LM_L_l77_Playfield, LM_L_l77_UnderPF)
Dim BL_L_l78: BL_L_l78=Array(LM_L_l78_Layer1, LM_L_l78_Parts, LM_L_l78_Playfield, LM_L_l78_UnderPF, LM_L_l78_sw45)
Dim BL_L_l79: BL_L_l79=Array(LM_L_l79_Layer1, LM_L_l79_Parts, LM_L_l79_Playfield, LM_L_l79_UnderPF)
Dim BL_L_l81: BL_L_l81=Array(LM_L_l81_Layer1, LM_L_l81_Parts, LM_L_l81_Playfield, LM_L_l81_UnderPF, LM_L_l81_sw35, LM_L_l81_sw36)
Dim BL_L_l82: BL_L_l82=Array(LM_L_l82_Playfield, LM_L_l82_UnderPF)
Dim BL_L_l83: BL_L_l83=Array(LM_L_l83_Layer1, LM_L_l83_Parts, LM_L_l83_Playfield, LM_L_l83_UnderPF)
Dim BL_L_l84: BL_L_l84=Array(LM_L_l84_Parts, LM_L_l84_Playfield, LM_L_l84_UnderPF)
Dim BL_L_l85: BL_L_l85=Array(LM_L_l85_Parts, LM_L_l85_Playfield, LM_L_l85_UnderPF)
Dim BL_L_l89: BL_L_l89=Array(LM_L_l89_Playfield, LM_L_l89_RFlipper1, LM_L_l89_RFlipper1U, LM_L_l89_UnderPF, LM_L_l89_sw30)
Dim BL_L_l90: BL_L_l90=Array(LM_L_l90_Parts, LM_L_l90_Playfield, LM_L_l90_RFlipper1, LM_L_l90_UnderPF, LM_L_l90_sw24)
Dim BL_L_l91: BL_L_l91=Array(LM_L_l91_Layer1, LM_L_l91_Parts, LM_L_l91_Playfield, LM_L_l91_UnderPF, LM_L_l91_sw40)
Dim BL_L_l92: BL_L_l92=Array(LM_L_l92_Gate001, LM_L_l92_Layer1, LM_L_l92_Parts, LM_L_l92_Playfield, LM_L_l92_UnderPF, LM_L_l92_sw37)
Dim BL_L_l93: BL_L_l93=Array(LM_L_l93_Playfield, LM_L_l93_RSling1, LM_L_l93_RSling2, LM_L_l93_UnderPF)
Dim BL_L_l94: BL_L_l94=Array(LM_L_l94_Layer1, LM_L_l94_Parts, LM_L_l94_Playfield, LM_L_l94_UnderPF)
Dim BL_L_l95: BL_L_l95=Array(LM_L_l95_Layer1, LM_L_l95_Parts, LM_L_l95_Playfield, LM_L_l95_UnderPF, LM_L_l95_sw32, LM_L_l95_sw33, LM_L_l95_sw45)
Dim BL_Room: BL_Room=Array(BM_BR1, BM_BR2, BM_BR3, BM_BS1, BM_BS2, BM_BS3, BM_Gate001, BM_Gate1, BM_Gate2, BM_Gate3, BM_Gate4, BM_LFlipper, BM_LFlipperU, BM_LSling1, BM_LSling2, BM_LSlingArm, BM_Layer1, BM_Parts, BM_PincabRails, BM_Playfield, BM_RFlipper, BM_RFlipper1, BM_RFlipper1U, BM_RFlipperU, BM_RSling1, BM_RSling2, BM_RSlingArm, BM_UnderPF, BM_sw1, BM_sw13, BM_sw16, BM_sw17, BM_sw18, BM_sw19, BM_sw2, BM_sw20, BM_sw21, BM_sw23, BM_sw24, BM_sw3, BM_sw30, BM_sw31, BM_sw32, BM_sw33, BM_sw34, BM_sw35, BM_sw36, BM_sw37, BM_sw38, BM_sw39, BM_sw4, BM_sw40, BM_sw45, BM_sw5, BM_sw7)
' Global arrays
Dim BG_Bakemap: BG_Bakemap=Array(BM_BR1, BM_BR2, BM_BR3, BM_BS1, BM_BS2, BM_BS3, BM_Gate001, BM_Gate1, BM_Gate2, BM_Gate3, BM_Gate4, BM_LFlipper, BM_LFlipperU, BM_LSling1, BM_LSling2, BM_LSlingArm, BM_Layer1, BM_Parts, BM_PincabRails, BM_Playfield, BM_RFlipper, BM_RFlipper1, BM_RFlipper1U, BM_RFlipperU, BM_RSling1, BM_RSling2, BM_RSlingArm, BM_UnderPF, BM_sw1, BM_sw13, BM_sw16, BM_sw17, BM_sw18, BM_sw19, BM_sw2, BM_sw20, BM_sw21, BM_sw23, BM_sw24, BM_sw3, BM_sw30, BM_sw31, BM_sw32, BM_sw33, BM_sw34, BM_sw35, BM_sw36, BM_sw37, BM_sw38, BM_sw39, BM_sw4, BM_sw40, BM_sw45, BM_sw5, BM_sw7)
Dim BG_Lightmap: BG_Lightmap=Array(LM_FL_F6a_BS1, LM_FL_F6a_BS2, LM_FL_F6a_Gate4, LM_FL_F6a_Layer1, LM_FL_F6a_Parts, LM_FL_F6a_Playfield, LM_FL_F6a_UnderPF, LM_FL_F6a_sw21, LM_FL_F6a_sw33, LM_FL_F6a_sw39, LM_FL_F6a_sw40, LM_FL_F6b_BS1, LM_FL_F6b_BS2, LM_FL_F6b_Gate2, LM_FL_F6b_Layer1, LM_FL_F6b_Parts, LM_FL_F6b_Playfield, LM_FL_F6b_UnderPF, LM_FL_F6b_sw32, LM_FL_F6b_sw33, LM_FL_F6b_sw34, LM_FL_F6b_sw35, LM_FL_F6b_sw36, LM_FL_F6b_sw37, LM_FL_F6b_sw38, LM_FL_F6c_Gate4, LM_FL_F6c_Layer1, LM_FL_F6c_Parts, LM_FL_F6c_Playfield, LM_FL_F6c_UnderPF, LM_FL_F6d_Gate2, LM_FL_F6d_Layer1, LM_FL_F6d_Parts, LM_FL_F6d_Playfield, LM_FL_F6d_UnderPF, LM_GI_G13_BS1, LM_GI_G13_BS2, LM_GI_G13_Gate3, LM_GI_G13_Layer1, LM_GI_G13_Parts, LM_GI_G13_Playfield, LM_GI_G13_UnderPF, LM_GI_G13_sw32, LM_GI_G13_sw33, LM_GI_G13_sw34, LM_GI_G13_sw35, LM_GI_G13_sw36, LM_GI_G13_sw37, LM_GI_G13_sw38, LM_GI_GI00_BS1, LM_GI_GI00_BS2, LM_GI_GI00_BS3, LM_GI_GI00_Gate2, LM_GI_GI00_Gate3, LM_GI_GI00_Layer1, LM_GI_GI00_Parts, LM_GI_GI00_Playfield, _
  LM_GI_GI00_RFlipper1, LM_GI_GI00_RFlipper1U, LM_GI_GI00_RSling1, LM_GI_GI00_RSling2, LM_GI_GI00_RSlingArm, LM_GI_GI00_UnderPF, LM_GI_GI00_sw1, LM_GI_GI00_sw17, LM_GI_GI00_sw18, LM_GI_GI00_sw19, LM_GI_GI00_sw2, LM_GI_GI00_sw20, LM_GI_GI00_sw21, LM_GI_GI00_sw3, LM_GI_GI00_sw30, LM_GI_GI00_sw31, LM_GI_GI00_sw33, LM_GI_GI00_sw35, LM_GI_GI00_sw36, LM_GI_GI00_sw4, LM_GI_GI00_sw45, LM_GI_GI00_sw5, LM_GI_GI00_sw7, LM_GI_GI01_LFlipper, LM_GI_GI01_LFlipperU, LM_GI_GI01_LSling1, LM_GI_GI01_LSling2, LM_GI_GI01_Parts, LM_GI_GI01_Playfield, LM_GI_GI01_RFlipper, LM_GI_GI01_RFlipperU, LM_GI_GI01_RSling1, LM_GI_GI01_RSling2, LM_GI_GI01_RSlingArm, LM_GI_GI01_UnderPF, LM_GI_GI01_sw16, LM_GI_GI01_sw24, LM_GI_GI01_sw30, LM_GI_GI01_sw31, LM_GI_GI02_Gate1, LM_GI_GI02_LFlipper, LM_GI_GI02_LFlipperU, LM_GI_GI02_Parts, LM_GI_GI02_Playfield, LM_GI_GI02_RFlipper, LM_GI_GI02_RFlipperU, LM_GI_GI02_UnderPF, LM_GI_GI02_sw13, LM_GI_GI03_Gate1, LM_GI_GI03_LFlipper, LM_GI_GI03_LFlipperU, LM_GI_GI03_LSling1, LM_GI_GI03_LSling2, _
  LM_GI_GI03_LSlingArm, LM_GI_GI03_Parts, LM_GI_GI03_Playfield, LM_GI_GI03_RSling1, LM_GI_GI03_RSling2, LM_GI_GI03_RSlingArm, LM_GI_GI03_UnderPF, LM_GI_GI03_sw30, LM_GI_GI03_sw31, LM_GI_GI04_LSling1, LM_GI_GI04_LSling2, LM_GI_GI04_Parts, LM_GI_GI04_Playfield, LM_GI_GI04_RFlipper, LM_GI_GI04_RFlipper1, LM_GI_GI04_RFlipper1U, LM_GI_GI04_RFlipperU, LM_GI_GI04_RSling1, LM_GI_GI04_RSling2, LM_GI_GI04_RSlingArm, LM_GI_GI04_UnderPF, LM_GI_GI04_sw30, LM_GI_GI04_sw31, LM_GI_GI07_BR1, LM_GI_GI07_BS1, LM_GI_GI07_BS2, LM_GI_GI07_BS3, LM_GI_GI07_Layer1, LM_GI_GI07_Parts, LM_GI_GI07_Playfield, LM_GI_GI07_UnderPF, LM_GI_GI07_sw17, LM_GI_GI07_sw18, LM_GI_GI07_sw19, LM_GI_GI07_sw20, LM_GI_GI07_sw21, LM_GI_GI07_sw3, LM_GI_GI07_sw36, LM_GI_GI07_sw4, LM_GI_GI07_sw40, LM_GI_GI07_sw5, LM_GI_GI07_sw7, LM_GI_GI08_BR1, LM_GI_GI08_BS1, LM_GI_GI08_BS3, LM_GI_GI08_Layer1, LM_GI_GI08_Parts, LM_GI_GI08_Playfield, LM_GI_GI08_UnderPF, LM_GI_GI08_sw17, LM_GI_GI08_sw18, LM_GI_GI08_sw19, LM_GI_GI08_sw20, LM_GI_GI08_sw21, LM_GI_GI08_sw33, _
  LM_GI_GI08_sw5, LM_GI_GI08_sw7, LM_GI_GI09_BS1, LM_GI_GI09_BS3, LM_GI_GI09_Layer1, LM_GI_GI09_Parts, LM_GI_GI09_Playfield, LM_GI_GI09_UnderPF, LM_GI_GI09_sw19, LM_GI_GI09_sw20, LM_GI_GI09_sw21, LM_GI_GI09_sw7, LM_GI_GI11_BS2, LM_GI_GI11_Gate3, LM_GI_GI11_Layer1, LM_GI_GI11_Parts, LM_GI_GI11_Playfield, LM_GI_GI11_UnderPF, LM_GI_GI11_sw21, LM_GI_GI11_sw33, LM_GI_GI11_sw39, LM_GI_GI11_sw40, LM_L_L02_Parts, LM_L_L02_Playfield, LM_L_L02_UnderPF, LM_L_l04_LFlipperU, LM_L_l04_Parts, LM_L_l04_Playfield, LM_L_l04_UnderPF, LM_L_l05_Parts, LM_L_l05_Playfield, LM_L_l05_RFlipperU, LM_L_l05_UnderPF, LM_L_l06_Playfield, LM_L_l06_UnderPF, LM_L_l07_Playfield, LM_L_l07_UnderPF, LM_L_l08_Playfield, LM_L_l08_UnderPF, LM_L_l09_Playfield, LM_L_l09_UnderPF, LM_L_l10_Playfield, LM_L_l10_UnderPF, LM_L_l10_sw30, LM_L_l11_Playfield, LM_L_l11_UnderPF, LM_L_l12_Playfield, LM_L_l12_UnderPF, LM_L_l13_Playfield, LM_L_l13_UnderPF, LM_L_l14_Parts, LM_L_l14_Playfield, LM_L_l14_UnderPF, LM_L_l14_sw16, LM_L_l15_Playfield, LM_L_l15_UnderPF, _
  LM_L_l16_Layer1, LM_L_l16_Parts, LM_L_l16_Playfield, LM_L_l16_UnderPF, LM_L_l16_sw32, LM_L_l20_Parts, LM_L_l20_Playfield, LM_L_l20_RFlipperU, LM_L_l20_UnderPF, LM_L_l21_Playfield, LM_L_l21_UnderPF, LM_L_l22_Playfield, LM_L_l22_UnderPF, LM_L_l23_Playfield, LM_L_l23_UnderPF, LM_L_l24_Playfield, LM_L_l24_UnderPF, LM_L_l25_Playfield, LM_L_l25_UnderPF, LM_L_l26_Playfield, LM_L_l26_UnderPF, LM_L_l27_Playfield, LM_L_l27_UnderPF, LM_L_l28_Playfield, LM_L_l28_UnderPF, LM_L_l29_Playfield, LM_L_l29_UnderPF, LM_L_l30_Playfield, LM_L_l30_UnderPF, LM_L_l31_Playfield, LM_L_l31_UnderPF, LM_L_l35_LFlipperU, LM_L_l35_Parts, LM_L_l35_Playfield, LM_L_l35_RFlipperU, LM_L_l35_UnderPF, LM_L_l36_LFlipperU, LM_L_l36_Parts, LM_L_l36_Playfield, LM_L_l36_UnderPF, LM_L_l37_Playfield, LM_L_l37_UnderPF, LM_L_l38_Playfield, LM_L_l38_UnderPF, LM_L_l39_Playfield, LM_L_l39_UnderPF, LM_L_l40_Playfield, LM_L_l40_UnderPF, LM_L_l41_Playfield, LM_L_l41_UnderPF, LM_L_l42_Playfield, LM_L_l42_UnderPF, LM_L_l43_Playfield, LM_L_l43_UnderPF, _
  LM_L_l44_Playfield, LM_L_l44_UnderPF, LM_L_l45_Parts, LM_L_l45_Playfield, LM_L_l45_UnderPF, LM_L_l45_sw13, LM_L_l46_Playfield, LM_L_l46_UnderPF, LM_L_l47_Playfield, LM_L_l47_UnderPF, LM_L_l49_Layer1, LM_L_l49_Parts, LM_L_l49_Playfield, LM_L_l49_UnderPF, LM_L_l49_sw33, LM_L_l49_sw34, LM_L_l50_Layer1, LM_L_l50_Parts, LM_L_l50_Playfield, LM_L_l50_UnderPF, LM_L_l50_sw36, LM_L_l51_Parts, LM_L_l51_Playfield, LM_L_l51_UnderPF, LM_L_l52_Parts, LM_L_l52_Playfield, LM_L_l52_UnderPF, LM_L_l53_Parts, LM_L_l53_Playfield, LM_L_l53_UnderPF, LM_L_l54_Parts, LM_L_l54_Playfield, LM_L_l54_UnderPF, LM_L_l58_Parts, LM_L_l58_Playfield, LM_L_l58_UnderPF, LM_L_l58_sw31, LM_L_l59_Layer1, LM_L_l59_Parts, LM_L_l59_Playfield, LM_L_l59_UnderPF, LM_L_l60_Parts, LM_L_l60_Playfield, LM_L_l60_UnderPF, LM_L_l60_sw39, LM_L_l62_LFlipper, LM_L_l62_LFlipperU, LM_L_l62_Parts, LM_L_l62_Playfield, LM_L_l62_RFlipper, LM_L_l62_RFlipperU, LM_L_l62_UnderPF, LM_L_l63_Layer1, LM_L_l63_Parts, LM_L_l63_Playfield, LM_L_l63_UnderPF, LM_L_l65_Layer1, _
  LM_L_l65_Parts, LM_L_l65_Playfield, LM_L_l65_UnderPF, LM_L_l65_sw34, LM_L_l65_sw35, LM_L_l65_sw36, LM_L_l66_Playfield, LM_L_l66_UnderPF, LM_L_l67_Playfield, LM_L_l67_UnderPF, LM_L_l68_Parts, LM_L_l68_Playfield, LM_L_l68_UnderPF, LM_L_l69_Playfield, LM_L_l69_UnderPF, LM_L_l74_Parts, LM_L_l74_Playfield, LM_L_l74_UnderPF, LM_L_l74_sw23, LM_L_l74_sw31, LM_L_l75_Layer1, LM_L_l75_Parts, LM_L_l75_Playfield, LM_L_l75_UnderPF, LM_L_l76_Parts, LM_L_l76_Playfield, LM_L_l76_UnderPF, LM_L_l76_sw38, LM_L_l77_Playfield, LM_L_l77_UnderPF, LM_L_l78_Layer1, LM_L_l78_Parts, LM_L_l78_Playfield, LM_L_l78_UnderPF, LM_L_l78_sw45, LM_L_l79_Layer1, LM_L_l79_Parts, LM_L_l79_Playfield, LM_L_l79_UnderPF, LM_L_l81_Layer1, LM_L_l81_Parts, LM_L_l81_Playfield, LM_L_l81_UnderPF, LM_L_l81_sw35, LM_L_l81_sw36, LM_L_l82_Playfield, LM_L_l82_UnderPF, LM_L_l83_Layer1, LM_L_l83_Parts, LM_L_l83_Playfield, LM_L_l83_UnderPF, LM_L_l84_Parts, LM_L_l84_Playfield, LM_L_l84_UnderPF, LM_L_l85_Parts, LM_L_l85_Playfield, LM_L_l85_UnderPF, LM_L_l89_Playfield, _
  LM_L_l89_RFlipper1, LM_L_l89_RFlipper1U, LM_L_l89_UnderPF, LM_L_l89_sw30, LM_L_l90_Parts, LM_L_l90_Playfield, LM_L_l90_RFlipper1, LM_L_l90_UnderPF, LM_L_l90_sw24, LM_L_l91_Layer1, LM_L_l91_Parts, LM_L_l91_Playfield, LM_L_l91_UnderPF, LM_L_l91_sw40, LM_L_l92_Gate001, LM_L_l92_Layer1, LM_L_l92_Parts, LM_L_l92_Playfield, LM_L_l92_UnderPF, LM_L_l92_sw37, LM_L_l93_Playfield, LM_L_l93_RSling1, LM_L_l93_RSling2, LM_L_l93_UnderPF, LM_L_l94_Layer1, LM_L_l94_Parts, LM_L_l94_Playfield, LM_L_l94_UnderPF, LM_L_l95_Layer1, LM_L_l95_Parts, LM_L_l95_Playfield, LM_L_l95_UnderPF, LM_L_l95_sw32, LM_L_l95_sw33, LM_L_l95_sw45)
Dim BG_All: BG_All=Array(BM_BR1, BM_BR2, BM_BR3, BM_BS1, BM_BS2, BM_BS3, BM_Gate001, BM_Gate1, BM_Gate2, BM_Gate3, BM_Gate4, BM_LFlipper, BM_LFlipperU, BM_LSling1, BM_LSling2, BM_LSlingArm, BM_Layer1, BM_Parts, BM_PincabRails, BM_Playfield, BM_RFlipper, BM_RFlipper1, BM_RFlipper1U, BM_RFlipperU, BM_RSling1, BM_RSling2, BM_RSlingArm, BM_UnderPF, BM_sw1, BM_sw13, BM_sw16, BM_sw17, BM_sw18, BM_sw19, BM_sw2, BM_sw20, BM_sw21, BM_sw23, BM_sw24, BM_sw3, BM_sw30, BM_sw31, BM_sw32, BM_sw33, BM_sw34, BM_sw35, BM_sw36, BM_sw37, BM_sw38, BM_sw39, BM_sw4, BM_sw40, BM_sw45, BM_sw5, BM_sw7, LM_FL_F6a_BS1, LM_FL_F6a_BS2, LM_FL_F6a_Gate4, LM_FL_F6a_Layer1, LM_FL_F6a_Parts, LM_FL_F6a_Playfield, LM_FL_F6a_UnderPF, LM_FL_F6a_sw21, LM_FL_F6a_sw33, LM_FL_F6a_sw39, LM_FL_F6a_sw40, LM_FL_F6b_BS1, LM_FL_F6b_BS2, LM_FL_F6b_Gate2, LM_FL_F6b_Layer1, LM_FL_F6b_Parts, LM_FL_F6b_Playfield, LM_FL_F6b_UnderPF, LM_FL_F6b_sw32, LM_FL_F6b_sw33, LM_FL_F6b_sw34, LM_FL_F6b_sw35, LM_FL_F6b_sw36, LM_FL_F6b_sw37, LM_FL_F6b_sw38, LM_FL_F6c_Gate4, _
  LM_FL_F6c_Layer1, LM_FL_F6c_Parts, LM_FL_F6c_Playfield, LM_FL_F6c_UnderPF, LM_FL_F6d_Gate2, LM_FL_F6d_Layer1, LM_FL_F6d_Parts, LM_FL_F6d_Playfield, LM_FL_F6d_UnderPF, LM_GI_G13_BS1, LM_GI_G13_BS2, LM_GI_G13_Gate3, LM_GI_G13_Layer1, LM_GI_G13_Parts, LM_GI_G13_Playfield, LM_GI_G13_UnderPF, LM_GI_G13_sw32, LM_GI_G13_sw33, LM_GI_G13_sw34, LM_GI_G13_sw35, LM_GI_G13_sw36, LM_GI_G13_sw37, LM_GI_G13_sw38, LM_GI_GI00_BS1, LM_GI_GI00_BS2, LM_GI_GI00_BS3, LM_GI_GI00_Gate2, LM_GI_GI00_Gate3, LM_GI_GI00_Layer1, LM_GI_GI00_Parts, LM_GI_GI00_Playfield, LM_GI_GI00_RFlipper1, LM_GI_GI00_RFlipper1U, LM_GI_GI00_RSling1, LM_GI_GI00_RSling2, LM_GI_GI00_RSlingArm, LM_GI_GI00_UnderPF, LM_GI_GI00_sw1, LM_GI_GI00_sw17, LM_GI_GI00_sw18, LM_GI_GI00_sw19, LM_GI_GI00_sw2, LM_GI_GI00_sw20, LM_GI_GI00_sw21, LM_GI_GI00_sw3, LM_GI_GI00_sw30, LM_GI_GI00_sw31, LM_GI_GI00_sw33, LM_GI_GI00_sw35, LM_GI_GI00_sw36, LM_GI_GI00_sw4, LM_GI_GI00_sw45, LM_GI_GI00_sw5, LM_GI_GI00_sw7, LM_GI_GI01_LFlipper, LM_GI_GI01_LFlipperU, LM_GI_GI01_LSling1, _
  LM_GI_GI01_LSling2, LM_GI_GI01_Parts, LM_GI_GI01_Playfield, LM_GI_GI01_RFlipper, LM_GI_GI01_RFlipperU, LM_GI_GI01_RSling1, LM_GI_GI01_RSling2, LM_GI_GI01_RSlingArm, LM_GI_GI01_UnderPF, LM_GI_GI01_sw16, LM_GI_GI01_sw24, LM_GI_GI01_sw30, LM_GI_GI01_sw31, LM_GI_GI02_Gate1, LM_GI_GI02_LFlipper, LM_GI_GI02_LFlipperU, LM_GI_GI02_Parts, LM_GI_GI02_Playfield, LM_GI_GI02_RFlipper, LM_GI_GI02_RFlipperU, LM_GI_GI02_UnderPF, LM_GI_GI02_sw13, LM_GI_GI03_Gate1, LM_GI_GI03_LFlipper, LM_GI_GI03_LFlipperU, LM_GI_GI03_LSling1, LM_GI_GI03_LSling2, LM_GI_GI03_LSlingArm, LM_GI_GI03_Parts, LM_GI_GI03_Playfield, LM_GI_GI03_RSling1, LM_GI_GI03_RSling2, LM_GI_GI03_RSlingArm, LM_GI_GI03_UnderPF, LM_GI_GI03_sw30, LM_GI_GI03_sw31, LM_GI_GI04_LSling1, LM_GI_GI04_LSling2, LM_GI_GI04_Parts, LM_GI_GI04_Playfield, LM_GI_GI04_RFlipper, LM_GI_GI04_RFlipper1, LM_GI_GI04_RFlipper1U, LM_GI_GI04_RFlipperU, LM_GI_GI04_RSling1, LM_GI_GI04_RSling2, LM_GI_GI04_RSlingArm, LM_GI_GI04_UnderPF, LM_GI_GI04_sw30, LM_GI_GI04_sw31, LM_GI_GI07_BR1, _
  LM_GI_GI07_BS1, LM_GI_GI07_BS2, LM_GI_GI07_BS3, LM_GI_GI07_Layer1, LM_GI_GI07_Parts, LM_GI_GI07_Playfield, LM_GI_GI07_UnderPF, LM_GI_GI07_sw17, LM_GI_GI07_sw18, LM_GI_GI07_sw19, LM_GI_GI07_sw20, LM_GI_GI07_sw21, LM_GI_GI07_sw3, LM_GI_GI07_sw36, LM_GI_GI07_sw4, LM_GI_GI07_sw40, LM_GI_GI07_sw5, LM_GI_GI07_sw7, LM_GI_GI08_BR1, LM_GI_GI08_BS1, LM_GI_GI08_BS3, LM_GI_GI08_Layer1, LM_GI_GI08_Parts, LM_GI_GI08_Playfield, LM_GI_GI08_UnderPF, LM_GI_GI08_sw17, LM_GI_GI08_sw18, LM_GI_GI08_sw19, LM_GI_GI08_sw20, LM_GI_GI08_sw21, LM_GI_GI08_sw33, LM_GI_GI08_sw5, LM_GI_GI08_sw7, LM_GI_GI09_BS1, LM_GI_GI09_BS3, LM_GI_GI09_Layer1, LM_GI_GI09_Parts, LM_GI_GI09_Playfield, LM_GI_GI09_UnderPF, LM_GI_GI09_sw19, LM_GI_GI09_sw20, LM_GI_GI09_sw21, LM_GI_GI09_sw7, LM_GI_GI11_BS2, LM_GI_GI11_Gate3, LM_GI_GI11_Layer1, LM_GI_GI11_Parts, LM_GI_GI11_Playfield, LM_GI_GI11_UnderPF, LM_GI_GI11_sw21, LM_GI_GI11_sw33, LM_GI_GI11_sw39, LM_GI_GI11_sw40, LM_L_L02_Parts, LM_L_L02_Playfield, LM_L_L02_UnderPF, LM_L_l04_LFlipperU, LM_L_l04_Parts, _
  LM_L_l04_Playfield, LM_L_l04_UnderPF, LM_L_l05_Parts, LM_L_l05_Playfield, LM_L_l05_RFlipperU, LM_L_l05_UnderPF, LM_L_l06_Playfield, LM_L_l06_UnderPF, LM_L_l07_Playfield, LM_L_l07_UnderPF, LM_L_l08_Playfield, LM_L_l08_UnderPF, LM_L_l09_Playfield, LM_L_l09_UnderPF, LM_L_l10_Playfield, LM_L_l10_UnderPF, LM_L_l10_sw30, LM_L_l11_Playfield, LM_L_l11_UnderPF, LM_L_l12_Playfield, LM_L_l12_UnderPF, LM_L_l13_Playfield, LM_L_l13_UnderPF, LM_L_l14_Parts, LM_L_l14_Playfield, LM_L_l14_UnderPF, LM_L_l14_sw16, LM_L_l15_Playfield, LM_L_l15_UnderPF, LM_L_l16_Layer1, LM_L_l16_Parts, LM_L_l16_Playfield, LM_L_l16_UnderPF, LM_L_l16_sw32, LM_L_l20_Parts, LM_L_l20_Playfield, LM_L_l20_RFlipperU, LM_L_l20_UnderPF, LM_L_l21_Playfield, LM_L_l21_UnderPF, LM_L_l22_Playfield, LM_L_l22_UnderPF, LM_L_l23_Playfield, LM_L_l23_UnderPF, LM_L_l24_Playfield, LM_L_l24_UnderPF, LM_L_l25_Playfield, LM_L_l25_UnderPF, LM_L_l26_Playfield, LM_L_l26_UnderPF, LM_L_l27_Playfield, LM_L_l27_UnderPF, LM_L_l28_Playfield, LM_L_l28_UnderPF, LM_L_l29_Playfield, _
  LM_L_l29_UnderPF, LM_L_l30_Playfield, LM_L_l30_UnderPF, LM_L_l31_Playfield, LM_L_l31_UnderPF, LM_L_l35_LFlipperU, LM_L_l35_Parts, LM_L_l35_Playfield, LM_L_l35_RFlipperU, LM_L_l35_UnderPF, LM_L_l36_LFlipperU, LM_L_l36_Parts, LM_L_l36_Playfield, LM_L_l36_UnderPF, LM_L_l37_Playfield, LM_L_l37_UnderPF, LM_L_l38_Playfield, LM_L_l38_UnderPF, LM_L_l39_Playfield, LM_L_l39_UnderPF, LM_L_l40_Playfield, LM_L_l40_UnderPF, LM_L_l41_Playfield, LM_L_l41_UnderPF, LM_L_l42_Playfield, LM_L_l42_UnderPF, LM_L_l43_Playfield, LM_L_l43_UnderPF, LM_L_l44_Playfield, LM_L_l44_UnderPF, LM_L_l45_Parts, LM_L_l45_Playfield, LM_L_l45_UnderPF, LM_L_l45_sw13, LM_L_l46_Playfield, LM_L_l46_UnderPF, LM_L_l47_Playfield, LM_L_l47_UnderPF, LM_L_l49_Layer1, LM_L_l49_Parts, LM_L_l49_Playfield, LM_L_l49_UnderPF, LM_L_l49_sw33, LM_L_l49_sw34, LM_L_l50_Layer1, LM_L_l50_Parts, LM_L_l50_Playfield, LM_L_l50_UnderPF, LM_L_l50_sw36, LM_L_l51_Parts, LM_L_l51_Playfield, LM_L_l51_UnderPF, LM_L_l52_Parts, LM_L_l52_Playfield, LM_L_l52_UnderPF, LM_L_l53_Parts, _
  LM_L_l53_Playfield, LM_L_l53_UnderPF, LM_L_l54_Parts, LM_L_l54_Playfield, LM_L_l54_UnderPF, LM_L_l58_Parts, LM_L_l58_Playfield, LM_L_l58_UnderPF, LM_L_l58_sw31, LM_L_l59_Layer1, LM_L_l59_Parts, LM_L_l59_Playfield, LM_L_l59_UnderPF, LM_L_l60_Parts, LM_L_l60_Playfield, LM_L_l60_UnderPF, LM_L_l60_sw39, LM_L_l62_LFlipper, LM_L_l62_LFlipperU, LM_L_l62_Parts, LM_L_l62_Playfield, LM_L_l62_RFlipper, LM_L_l62_RFlipperU, LM_L_l62_UnderPF, LM_L_l63_Layer1, LM_L_l63_Parts, LM_L_l63_Playfield, LM_L_l63_UnderPF, LM_L_l65_Layer1, LM_L_l65_Parts, LM_L_l65_Playfield, LM_L_l65_UnderPF, LM_L_l65_sw34, LM_L_l65_sw35, LM_L_l65_sw36, LM_L_l66_Playfield, LM_L_l66_UnderPF, LM_L_l67_Playfield, LM_L_l67_UnderPF, LM_L_l68_Parts, LM_L_l68_Playfield, LM_L_l68_UnderPF, LM_L_l69_Playfield, LM_L_l69_UnderPF, LM_L_l74_Parts, LM_L_l74_Playfield, LM_L_l74_UnderPF, LM_L_l74_sw23, LM_L_l74_sw31, LM_L_l75_Layer1, LM_L_l75_Parts, LM_L_l75_Playfield, LM_L_l75_UnderPF, LM_L_l76_Parts, LM_L_l76_Playfield, LM_L_l76_UnderPF, LM_L_l76_sw38, _
  LM_L_l77_Playfield, LM_L_l77_UnderPF, LM_L_l78_Layer1, LM_L_l78_Parts, LM_L_l78_Playfield, LM_L_l78_UnderPF, LM_L_l78_sw45, LM_L_l79_Layer1, LM_L_l79_Parts, LM_L_l79_Playfield, LM_L_l79_UnderPF, LM_L_l81_Layer1, LM_L_l81_Parts, LM_L_l81_Playfield, LM_L_l81_UnderPF, LM_L_l81_sw35, LM_L_l81_sw36, LM_L_l82_Playfield, LM_L_l82_UnderPF, LM_L_l83_Layer1, LM_L_l83_Parts, LM_L_l83_Playfield, LM_L_l83_UnderPF, LM_L_l84_Parts, LM_L_l84_Playfield, LM_L_l84_UnderPF, LM_L_l85_Parts, LM_L_l85_Playfield, LM_L_l85_UnderPF, LM_L_l89_Playfield, LM_L_l89_RFlipper1, LM_L_l89_RFlipper1U, LM_L_l89_UnderPF, LM_L_l89_sw30, LM_L_l90_Parts, LM_L_l90_Playfield, LM_L_l90_RFlipper1, LM_L_l90_UnderPF, LM_L_l90_sw24, LM_L_l91_Layer1, LM_L_l91_Parts, LM_L_l91_Playfield, LM_L_l91_UnderPF, LM_L_l91_sw40, LM_L_l92_Gate001, LM_L_l92_Layer1, LM_L_l92_Parts, LM_L_l92_Playfield, LM_L_l92_UnderPF, LM_L_l92_sw37, LM_L_l93_Playfield, LM_L_l93_RSling1, LM_L_l93_RSling2, LM_L_l93_UnderPF, LM_L_l94_Layer1, LM_L_l94_Parts, LM_L_l94_Playfield, _
  LM_L_l94_UnderPF, LM_L_l95_Layer1, LM_L_l95_Parts, LM_L_l95_Playfield, LM_L_l95_UnderPF, LM_L_l95_sw32, LM_L_l95_sw33, LM_L_l95_sw45)
' VLM Arrays - End




'*******************************************
' ZTIM:Timers
'*******************************************

' FIXME for the time being, the cor timer interval must be 10 ms (so below 60FPS framerate)
CorTimer.Interval = 10
Sub CorTimer_Timer(): Cor.Update: End Sub


Sub FrameTimer_Timer() 'The frame timer interval should be -1, so executes at the display frame rate
  BSUpdate
  AnimateBumperSkirts
  RollingUpdate   'update rolling sounds
  DoDTAnim    'handle drop target animations
  UpdateDropTargets
  DoSTAnim    'handle stand up target animations
  UpdateStandupTargets
  If VRMode = True Then
    VRBackglassTimer ' Update BG digits and BG flashers
  End If
End Sub



'*******************************************
' ZINI: Table Initialization and Exiting
'*******************************************


Dim BTCBall1, gBOT

Sub Table1_Init
  vpmInit Me
  On Error Resume Next
    With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine = "Beat The Clock'"&chr(13)&"VPW"
    .HandleMechanics=0
    .HandleKeyboard=0
    .ShowDMDOnly=1
    .ShowFrame=0
    .ShowTitle=0
    .hidden = 1
        .Games(cGameName).Settings.Value("sound")=1
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With
     On Error Goto 0

  vpmMapLights AllLamps ' Map all lamps to the corresponding ROM output using the value of TimerInterval of each light object

  PinMAMETimer.Interval=PinMAMEInterval
  PinMAMETimer.Enabled=1

  vpmNudge.TiltSwitch=15
  vpmNudge.Sensitivity=5
  vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingshot)

    'Ball initializations need for physical trough
  Set BTCBall1 = sw8.CreateSizedballWithMass(Ballsize / 2,Ballmass)
  gBOT = Array(BTCBall1)
  Controller.Switch(8) = 1

  ' Initizations
  SetGI 0
  SetupVRRoom
  SetupVRBackglass

  'Initialize slings
  RStep = 0:RightSlingShot.Timerenabled=True
  LStep = 0:LeftSlingShot.Timerenabled=True

End Sub


Sub Table1_Exit()
  Controller.Pause = False
  Controller.Stop
End Sub


'*******************************************
'  ZOPT: User Options
'*******************************************


Dim LightLevel : LightLevel = 0.5     ' Level of room lighting (0 to 1), where 0 is dark and 100 is brightest
Dim VolumeDial : VolumeDial = 0.8           'Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Dim BallRollVolume : BallRollVolume = 0.5   'Level of ball rolling volume. Value between 0 and 1

' Called when options are tweaked by the player.
' - 0: game has started, good time to load options and adjust accordingly
' - 1: an option has changed
' - 2: options have been reseted
' - 3: player closed the tweak UI, good time to update staticly prerendered parts
' Table1.Option arguments are:
' - option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional arry of literal strings
Sub Table1_OptionEvent(ByVal eventId)
    If eventId = 1 Then DisableStaticPreRendering = True

    ' Sound volumes
    VolumeDial = Table1.Option("Mech Volume", 0, 1, 0.01, 0.8, 1)
    BallRollVolume = Table1.Option("Ball Roll Volume", 0, 1, 0.01, 0.5, 1)

    ' Room Brightness
    LightLevel = NightDay/100
    SetRoomBrightness LightLevel

  SetupVRRoom

    If eventId = 3 Then DisableStaticPreRendering = False
End Sub





'****************************
'   ZBRI: Room Brightness
'****************************

' Update these arrays if you want to change more materials with room light level
Dim RoomBrightnessMtlArray: RoomBrightnessMtlArray = Array("VLM.Bake.Active","VLM.Bake.Solid","VLM.Bake.Metal","Plastic with an image")


Sub SetRoomBrightness(lvl)
  If lvl > 1 Then lvl = 1
  If lvl < 0 Then lvl = 0
  lvl = lvl^2

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




'********************************************************
' ZSOL: Solenoid Callbacks
'********************************************************

SolCallBack(6)="Flasher6" 'Flasher 6
SolCallBack(7)="Flasher7" 'Flasher 7
SolCallBack(8)="SolOut46" 'Left Saucer
SolCallBack(9)="SolOut47" 'Middle Saucer
SolCallback(10)="SolDropUpDTR" 'Single Drop Target Reset
SolCallback(11)="SolDropUpDTL" '1-6 Drop Target Reset
SolCallback(14)="SolRelease"  'BallRelease
SolCallback(15)="vpmSolSound SoundFX(""Knocker_1"",DOFKnocker)," 'Knocker
SolCallback(17)="SolGIOn"


Sub SolGIOn(Enabled)
  If Enabled Then
    SetGI 1
  Else
    SetGI 0
  End If
End Sub

Sub Flasher6(Enabled)
  If Enabled Then
    F6a.state = 1
    F6b.state = 1
    F6c.state = 1
    F6d.state = 1
  Else
    F6a.state = 0
    F6b.state = 0
    F6c.state = 0
    F6d.state = 0
  End If
End Sub

Sub Flasher7(Enabled)
End Sub

Sub SolOut46(Enabled)
  If Enabled Then
    sw46.kick 178+rnd*4, 15+rnd*6
    SoundSaucerKick 1,sw47
  End If
End Sub

Sub SolOut47(Enabled)
  If Enabled Then
    sw47.kick 178+rnd*4, 17+rnd*6
    SoundSaucerKick 1,sw47
  End If
End Sub

Sub SolDropUpDTL(Enabled)
  If Enabled Then
    DTRaise 1
    DTRaise 2
    DTRaise 3
    DTRaise 4
    DTRaise 5
    DTRaise 7
    RandomSoundDropTargetReset BM_sw1
    RandomSoundDropTargetReset BM_sw7
  End If
End Sub

Sub SolDropUpDTR(Enabled)
  If Enabled Then
    DTRaise 45
    RandomSoundDropTargetReset BM_sw45
  End If
End Sub




'*******************************************
' ZFLP: Flippers
'*******************************************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

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
    RightFlipper1.RotateToEnd
    If rightflipper.currentangle > rightflipper.endangle - ReflipAngle Then
      RandomSoundReflipUpRight RightFlipper
    Else
      SoundFlipperUpAttackRight RightFlipper
      RandomSoundFlipperUpRight RightFlipper
    End If
  Else
    FlipperDeActivate RightFlipper, RFPress
    RightFlipper.RotateToStart
    RightFlipper1.RotateToStart
    If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
      RandomSoundFlipperDownRight RightFlipper
    End If
    FlipperRightHitParm = FlipperUpSoundLevel
  End If
End Sub

Sub LeftFlipper_Collide(parm)
  CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
  LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
  CheckLiveCatch Activeball, RightFlipper, RFCount, parm
  RightFlipperCollide parm
End Sub


'****************************************************************
' ZGII: GI
'****************************************************************

'Playfield GI

Sub SetGI(s)
  dim xx
  For each xx in GI:xx.State = s: Next
  ' VR Backglass simple flasher GI
  If VRMode = True Then
    If s = 0 Then
      BGDark.Visible = 1
      BGBright.Visible = 0
    Else
      BGDark.Visible = 1
      BGBright.Visible = 1
    End If
  End If
End Sub



'*******************************************
' ZDRN: Drain, Trough, and Ball Release
'*******************************************
' It is best practice to never destroy balls. This leads to more stable and accurate pinball game simulations.
' The following code supports a "physical trough" where balls are not destroyed.
' To use this,
'   - The trough geometry needs to be modeled with walls, and a set of kickers needs to be added to
'  the trough. The number of kickers depends on the number of physical balls on the table.
'   - A timer called "UpdateTroughTimer" needs to be added to the table. It should have an interval of 300 and be initially disabled.
'   - The balls need to be created within the Table1_Init sub. A global ball array (gBOT) can be created and used throughout the script


' TROUGH DRAIN & RELEASE
Sub sw8_Hit
  Controller.Switch(8) = 1
  RandomSoundDrain sw8
End Sub

Sub SolRelease(enabled)
  If enabled Then
    Controller.Switch(8) = 0
    sw8.kick 60, 20
    RandomSoundBallRelease sw8
  End If
End Sub



'*******************************************
' ZKEY: Key Press Handling
'*******************************************

Sub Table1_KeyDown(ByVal keycode)
  If keycode = LeftFlipperKey and VRMode = True Then VR_ButtonLeft.TransX = 8 ' VRroom
  If keycode = RightFlipperKey and VRMode = True Then VR_ButtonRight.TransX = -8 ' VRroom
  If keycode = StartGameKey and VRMode = True Then VR_ButtonStart.TransY = -8 ' VRroom

  If keycode = StartGameKey Then SoundStartButton
  If keycode = RightFlipperKey Then Controller.Switch(12) = 1
  If keycode = LeftTiltKey Then : Nudge 90, 1 : SoundNudgeLeft :End If
  If keycode = RightTiltKey Then : Nudge 270, 1 : SoundNudgeRight : End If
  If keycode = CenterTiltKey Then : Nudge 0, 1 : SoundNudgeCenter : End If
  If keycode = PlungerKey Then
    Plunger.PullBack
    SoundPlungerPull
    TimerVRPlunger.Enabled = True
    TimerVRPlunger2.Enabled = False
  End If
  If keycode = keyInsertCoin1 or keycode = keyInsertCoin2 or keycode = keyInsertCoin3 or keycode = keyInsertCoin4 or keycode = AddCreditKey Or keycode = AddCreditKey2 Then
    Select Case Int(Rnd * 3)
      Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
      Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
      Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
    End Select
  End If
  If KeyDownHandler(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal keycode)
  If keycode = LeftFlipperKey and VRMode = True Then VR_ButtonLeft.TransX = 0 ' VRroom
  If keycode = RightFlipperKey and VRMode = True Then VR_ButtonRight.TransX = 0 ' VRroom
  If keycode = StartGameKey and VRMode = True Then VR_ButtonStart.TransY = 0 ' VRroom

  If KeyCode = PlungerKey Then
    Plunger.Fire
    SoundPlungerReleaseBall
    TimerVRPlunger.Enabled = False
    TimerVRPlunger2.Enabled = True
    VR_Shooter.Y = 0
  End If
  If keycode = RightFlipperKey Then Controller.Switch(12) = 0
  If KeyUpHandler(keycode) Then Exit Sub
End Sub



'*******************************************
' ZTRI: Triggers
'*******************************************


' Saucers

Sub sw46_Hit:   Controller.Switch(46)=1: SoundSaucerLock: End Sub
Sub sw46_UnHit: Controller.Switch(46)=0: End Sub
Sub sw47_Hit:   Controller.Switch(47)=1: SoundSaucerLock: End Sub
Sub sw47_UnHit: Controller.Switch(47)=0: End Sub

' Drop Targets

Sub sw1_Hit:DTHit 1:End Sub
Sub sw2_Hit:DTHit 2:End Sub
Sub sw3_Hit:DTHit 3:End Sub
Sub sw4_Hit:DTHit 4:End Sub
Sub sw5_Hit:DTHit 5:End Sub
Sub sw7_Hit:DTHit 7:End Sub
Sub sw45_Hit:DTHit 45:End Sub

' Stand Up Targets

Sub sw17_Hit:STHit 17:End Sub
Sub sw18_Hit:STHit 18:End Sub
Sub sw19_Hit:STHit 19:End Sub
Sub sw20_Hit:STHit 20:End Sub
Sub sw21_Hit:STHit 21:End Sub


Sub sw33_Hit:STHit 33:End Sub
Sub sw34_Hit:STHit 34:End Sub
Sub sw35_Hit:STHit 35:End Sub
Sub sw36_Hit:STHit 36:End Sub

' Star Triggers

Sub sw30_Hit:  Controller.Switch(30)=1:End Sub
Sub sw30_unHit:Controller.Switch(30)=0:End Sub
Sub sw31_Hit:  Controller.Switch(31)=1:End Sub
Sub sw31_unHit:Controller.Switch(31)=0:End Sub
Sub sw32_Hit:  Controller.Switch(32)=1: Me.TimerEnabled=1:L16.State=0:End Sub    'why L16 here? -apophis
Sub sw32_unHit:Controller.Switch(32)=0:End Sub
Sub sw32_Timer: Me.TimerEnabled=0:L16.State=1:End Sub


' Wire Triggers

Sub sw13_Hit:Controller.Switch(13)=1: End Sub
Sub sw13_unHit:Controller.Switch(13)=0:End Sub
Sub sw16_Hit:Controller.Switch(16)=1: End Sub
Sub sw16_unHit:Controller.Switch(16)=0:End Sub
Sub sw23_Hit:Controller.Switch(23)=1: End Sub
Sub sw23_unHit:Controller.Switch(23)=0:End Sub
Sub sw24_Hit:Controller.Switch(24)=1: End Sub
Sub sw24_unHit:Controller.Switch(24)=0:End Sub
Sub sw37_Hit:Controller.Switch(37)=1: End Sub
Sub sw37_unHit:Controller.Switch(37)=0: End Sub
Sub sw38_Hit:Controller.Switch(38)=1: End Sub
Sub sw38_unHit:Controller.Switch(38)=0: End Sub
Sub sw39_Hit:Controller.Switch(39)=1: End Sub
Sub sw39_unHit:Controller.Switch(39)=0: End Sub
Sub sw40_Hit:Controller.Switch(40)=1: End Sub
Sub sw40_unHit:Controller.Switch(40)=0: End Sub


'*******************************************
' ZBMP: Bumpers
'*******************************************

Sub Bumper1_Hit:vpmTimer.PulseSw 25 : RandomSoundBumperTop Bumper1 : End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 26 : RandomSoundBumperMiddle Bumper2 : End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 27 : RandomSoundBumperBottom Bumper3 : End Sub



'****************************************************************
' ZSLG: Slingshots
'****************************************************************

Dim RStep, Lstep

Sub RightSlingShot_Slingshot
  RS.VelocityCorrect(ActiveBall)
  vpmTimer.PulseSw 29
    RandomSoundSlingshotRight BM_RSlingArm
    RStep = 0
  RightSlingShot.TimerInterval = 17
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
  Dim BP
  Dim x1, x2, y: x1 = False:x2 = True:y = 15
    Select Case RStep
        Case 2:x1 = True:x2 = False:y = 9
        Case 3:x1 = False:x2 = False:y = 0:RightSlingShot.TimerEnabled = 0
    End Select

  For Each BP in BP_RSling1 : BP.Visible = x1: Next
  For Each BP in BP_RSling2 : BP.Visible = x2: Next
  For Each BP in BP_RSlingArm : BP.transx = y: Next

    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  LS.VelocityCorrect(ActiveBall)
  vpmTimer.PulseSw 28
    RandomSoundSlingshotLeft BM_LSlingArm
    LStep = 0
  LeftSlingShot.TimerInterval = 17
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
  Dim BP
  Dim x1, x2, y: x1 = False:x2 = True:y = 15
    Select Case LStep
        Case 2:x1 = True:x2 = False:y = 9
        Case 3:x1 = False:x2 = False:y = 0:LeftSlingShot.TimerEnabled = 0
    End Select

  For Each BP in BP_LSling1 : BP.Visible = x1: Next
  For Each BP in BP_LSling2 : BP.Visible = x2: Next
  For Each BP in BP_LSlingArm : BP.transx = y: Next

    LStep = LStep + 1
End Sub


'**********************************************************************************************************
'  ZLED: Digital Display
'**********************************************************************************************************

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

Sub DisplayTimer_Timer
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
' ZBRL:  BALL ROLLING AND DROP SOUNDS
'******************************************************

' Be sure to call RollingUpdate in a timer with a 10ms interval see the GameTimer_Timer() sub

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
'****  END BALL ROLLING AND DROP SOUNDS
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
' ZNFF:  FLIPPER CORRECTIONS by nFozzy
'******************************************************


'******************************************************
' Flippers Polarity (Select appropriate sub based on era)
'******************************************************

Dim LF
Set LF = New FlipperPolarity
Dim RF
Set RF = New FlipperPolarity
Dim RF2
Set RF2 = New FlipperPolarity

InitPolarity

'
''*******************************************
''  Late 80's early 90's
'
Sub InitPolarity()
  dim x, a : a = Array(LF, RF, RF2)
  for each x in a
    x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
    x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
    x.enabled = True
    x.TimeDelay = 60
    x.DebugOn=False ' prints some info in debugger

    x.AddPt "Polarity", 0, 0, 0
    x.AddPt "Polarity", 1, 0.05, - 5
    x.AddPt "Polarity", 2, 0.4, - 5
    x.AddPt "Polarity", 3, 0.6, - 4.5
    x.AddPt "Polarity", 4, 0.65, - 4.0
    x.AddPt "Polarity", 5, 0.7, - 3.5
    x.AddPt "Polarity", 6, 0.75, - 3.0
    x.AddPt "Polarity", 7, 0.8, - 2.5
    x.AddPt "Polarity", 8, 0.85, - 2.0
    x.AddPt "Polarity", 9, 0.9, - 1.5
    x.AddPt "Polarity", 10, 0.95, - 1.0
    x.AddPt "Polarity", 11, 1, - 0.5
    x.AddPt "Polarity", 12, 1.1, 0
    x.AddPt "Polarity", 13, 1.3, 0

    x.AddPt "Velocity", 0, 0, 1
    x.AddPt "Velocity", 1, 0.16, 1.06
    x.AddPt "Velocity", 2, 0.41, 1.05
    x.AddPt "Velocity", 3, 0.53, 1 '0.982
    x.AddPt "Velocity", 4, 0.702, 0.968
    x.AddPt "Velocity", 5, 0.95,  0.968
    x.AddPt "Velocity", 6, 1.03,  0.945
  Next

' ' SetObjects arguments: 1: name of object 2: flipper object: 3: Trigger object around flipper
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
'   Otherwise it should function exactly the same as before

Class FlipperPolarity
  Public DebugOn, Enabled
  Private FlipAt    'Timer variable (IE 'flip at 723,530ms...)
  Public TimeDelay    'delay before trigger turns off and polarity is disabled
  private Flipper, FlipperStart, FlipperEnd, FlipperEndY, LR, PartialFlipCoef
  Private Balls(20), balldata(20)
  private Name

  dim PolarityIn, PolarityOut
  dim VelocityIn, VelocityOut
  dim YcoefIn, YcoefOut
  Public Sub Class_Initialize
    redim PolarityIn(0) : redim PolarityOut(0) : redim VelocityIn(0) : redim VelocityOut(0) : redim YcoefIn(0) : redim YcoefOut(0)
    Enabled = True : TimeDelay = 50 : LR = 1:  dim x : for x = 0 to uBound(balls) : balls(x) = Empty : set Balldata(x) = new SpoofBall : next
  End Sub

  Public Sub SetObjects(aName, aFlipper, aTrigger)

    if typename(aName) <> "String" then msgbox "FlipperPolarity: .SetObjects error: first argument must be a string (and name of Object). Found:" & typename(aName) end if
    if typename(aFlipper) <> "Flipper" then msgbox "FlipperPolarity: .SetObjects error: second argument must be a flipper. Found:" & typename(aFlipper) end if
    if typename(aTrigger) <> "Trigger" then msgbox "FlipperPolarity: .SetObjects error: third argument must be a trigger. Found:" & typename(aTrigger) end if
    if aFlipper.EndAngle > aFlipper.StartAngle then LR = -1 Else LR = 1 End If
    Name = aName
    Set Flipper = aFlipper : FlipperStart = aFlipper.x
    FlipperEnd = Flipper.Length * sin((Flipper.StartAngle / 57.295779513082320876798154814105)) + Flipper.X ' big floats for degree to rad conversion
    FlipperEndY = Flipper.Length * cos(Flipper.StartAngle / 57.295779513082320876798154814105)*-1 + Flipper.Y

    dim str : str = "sub " & aTrigger.name & "_Hit() : " & aName & ".AddBall ActiveBall : End Sub'"
    ExecuteGlobal(str)
    str = "sub " & aTrigger.name & "_UnHit() : " & aName & ".PolarityCorrect ActiveBall : End Sub'"
    ExecuteGlobal(str)

  End Sub

  Public Property Let EndPoint(aInput) :  : End Property ' Legacy: just no op

  Public Sub AddPt(aChooseArray, aIDX, aX, aY) 'Index #, X position, (in) y Position (out)
    Select Case aChooseArray
      case "Polarity" : ShuffleArrays PolarityIn, PolarityOut, 1 : PolarityIn(aIDX) = aX : PolarityOut(aIDX) = aY : ShuffleArrays PolarityIn, PolarityOut, 0
      Case "Velocity" : ShuffleArrays VelocityIn, VelocityOut, 1 :VelocityIn(aIDX) = aX : VelocityOut(aIDX) = aY : ShuffleArrays VelocityIn, VelocityOut, 0
      Case "Ycoef" : ShuffleArrays YcoefIn, YcoefOut, 1 :YcoefIn(aIDX) = aX : YcoefOut(aIDX) = aY : ShuffleArrays YcoefIn, YcoefOut, 0
    End Select
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
  Private Function FlipperOn() : if gameTime < FlipAt+TimeDelay then FlipperOn = True : End If : End Function   'Timer shutoff for polaritycorrect

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
          if ballpos > 0.65 then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)                'find safety coefficient 'ycoef' data
        end if
      Next

      If BallPos = 0 Then 'no ball data meaning the ball is entering and exiting pretty close to the same position, use current values.
        BallPos = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
        if ballpos > 0.65 then  Ycoef = LinearEnvelope(aBall.Y, YcoefIn, YcoefOut)                        'find safety coefficient 'ycoef' data
      End If

      'Velocity correction
      if not IsEmpty(VelocityIn(0) ) then
        Dim VelCoef
        VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)

        if partialflipcoef < 1 then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)

        if Enabled then aBall.Velx = aBall.Velx*VelCoef
        if Enabled then aBall.Vely = aBall.Vely*VelCoef
      End If

      'Polarity Correction (optional now)
      if not IsEmpty(PolarityIn(0) ) then
        dim AddX : AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR

        if Enabled then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef)
      End If
      if DebugOn then debug.print "PolarityCorrect" & " " & Name & " @ " & gametime & " " & Round(BallPos*100) & "%" & " AddX:" & Round(AddX,2) & " Vel%:" & Round(VelCoef*100)
    End If
    RemoveBall aBall
  End Sub
End Class

'******************************************************
'  FLIPPER POLARITY AND RUBBER DAMPENER SUPPORTING FUNCTIONS
'******************************************************

' Used for flipper correction and rubber dampeners
Sub ShuffleArray(ByRef aArray, byVal offset) 'shuffle 1d array
  dim x, aCount : aCount = 0
  redim a(uBound(aArray) )
  for x = 0 to uBound(aArray)   'Shuffle objects in a temp array
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
  redim aArray(aCount-1+offset)   'Resize original array
  for x = 0 to aCount-1       'set objects back into original array
    if IsObject(a(x)) then
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
  BallSpeed = SQR(ball.VelX^2 + ball.VelY^2 + ball.VelZ^2)
End Function

' Used for flipper correction and rubber dampeners
Function PSlope(Input, X1, Y1, X2, Y2)    'Set up line via two points, no clamping. Input X, output Y
  dim x, y, b, m : x = input : m = (Y2 - Y1) / (X2 - X1) : b = Y2 - m*X2
  Y = M*x+b
  PSlope = Y
End Function

' Used for flipper correction
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

' Used for flipper correction and rubber dampeners
Function LinearEnvelope(xInput, xKeyFrame, yLvl)
  dim y 'Y output
  dim L 'Line
  dim ii : for ii = 1 to uBound(xKeyFrame)    'find active line
    if xInput <= xKeyFrame(ii) then L = ii : exit for : end if
  Next
  if xInput > xKeyFrame(uBound(xKeyFrame) ) then L = uBound(xKeyFrame)    'catch line overrun
  Y = pSlope(xInput, xKeyFrame(L-1), yLvl(L-1), xKeyFrame(L), yLvl(L) )

  if xInput <= xKeyFrame(lBound(xKeyFrame) ) then Y = yLvl(lBound(xKeyFrame) )     'Clamp lower
  if xInput >= xKeyFrame(uBound(xKeyFrame) ) then Y = yLvl(uBound(xKeyFrame) )    'Clamp upper

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
  Dim BOT
  BOT = GetBalls

  If Flipper1.currentangle = Endangle1 And EOSNudge1 <> 1 Then
    EOSNudge1 = 1
    '   debug.print Flipper1.currentangle &" = "& Endangle1 &"--"& Flipper2.currentangle &" = "& EndAngle2
    If Flipper2.currentangle = EndAngle2 Then
      For b = 0 To UBound(BOT)
        If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper1) Then
          'Debug.Print "ball in flip1. exit"
          Exit Sub
        End If
      Next
      For b = 0 To UBound(BOT)
        If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper2) Then
          BOT(b).velx = BOT(b).velx / 1.3
          BOT(b).vely = BOT(b).vely - 0.5
        End If
      Next
    End If
  Else
    If Abs(Flipper1.currentangle) > Abs(EndAngle1) + 30 Then EOSNudge1 = 0
  End If
End Sub

'*****************
' Maths
'*****************

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

Function max(a,b)
  if a > b then
    max = a
  Else
    max = b
  end if
end Function

Function min(a,b)
  if a > b then
    min = b
  Else
    min = a
  end if
end Function

'*************************************************
'  Check ball distance from Flipper for Rem
'*************************************************

Function Distance(ax,ay,bx,by)
  Distance = Sqr((ax - bx) ^ 2 + (ay - by) ^ 2)
End Function

Function DistancePL(px,py,ax,ay,bx,by) 'Distance between a point and a line where point is px,py
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
'Const EOSReturn = 0.055  'EM's
Const EOSReturn = 0.045  'late 70's to mid 80's
'Const EOSReturn = 0.035  'mid 80's to early 90's
'Const EOSReturn = 0.025  'mid 90's and later

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
    Dim b, BOT
    BOT = GetBalls

    For b = 0 To UBound(BOT)
      If Distance(BOT(b).x, BOT(b).y, Flipper.x, Flipper.y) < 55 Then 'check for cradle
        If BOT(b).vely >= - 0.4 Then BOT(b).vely =  - 0.4
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

Const LiveDistanceMin = 30  'minimum distance in vp units from flipper base live catch dampening will occur
Const LiveDistanceMax = 114 'maximum distance in vp units from flipper base live catch dampening will occur (tip protection)

Sub CheckLiveCatch(ball, Flipper, FCount, parm) 'Experimental new live catch
  Dim Dir
  Dir = Flipper.startangle / Abs(Flipper.startangle)  '-1 for Right Flipper
  Dim LiveCatchBounce                                                           'If live catch is not perfect, it won't freeze ball totally
  Dim CatchTime
  CatchTime = GameTime - FCount

  If CatchTime <= LiveCatch And parm > 6 And Abs(Flipper.x - ball.x) > LiveDistanceMin And Abs(Flipper.x - ball.x) < LiveDistanceMax Then
    If CatchTime <= LiveCatch * 0.5 Then                        'Perfect catch only when catch time happens in the beginning of the window
      LiveCatchBounce = 0
    Else
      LiveCatchBounce = Abs((LiveCatch / 2) - CatchTime)    'Partial catch when catch happens a bit late
    End If

    If LiveCatchBounce = 0 And ball.velx * Dir > 0 Then ball.velx = 0
    ball.vely = LiveCatchBounce * (32 / LiveCatch) ' Multiplier for inaccuracy bounce
    ball.angmomx = 0
    ball.angmomy = 0
    ball.angmomz = 0
  Else
    If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 1 Then FlippersD.Dampenf Activeball, parm
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
'   ZBOU: VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************

Const TargetBouncerEnabled = 1    '0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.9  'Level of bounces. Recommmended value of 0.7

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
' Tutorial vides by Apophis
' Audio : Adding Fleep Part 1       https://youtu.be/rG35JVHxtx4
' Audio : Adding Fleep Part 2       https://youtu.be/dk110pWMxGo
' Audio : Adding Fleep Part 3       https://youtu.be/ESXWGJZY_EI


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
  PlaySound playsoundparams, 0, min(1,aVol) * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelExistingStatic(playsoundparams, aVol, tableobj)
  PlaySound playsoundparams, 0, min(1,aVol) * VolumeDial, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticLoop(playsoundparams, aVol, tableobj)
  PlaySound playsoundparams, - 1, min(1,aVol) * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticRandomPitch(playsoundparams, aVol, randomPitch, tableobj)
  PlaySound playsoundparams, 0, min(1,aVol) * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelActiveBall(playsoundparams, aVol)
  PlaySound playsoundparams, 0, min(1,aVol) * VolumeDial, AudioPan(ActiveBall), 0, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLevelExistingActiveBall(playsoundparams, aVol)
  PlaySound playsoundparams, 0, min(1,aVol) * VolumeDial, AudioPan(ActiveBall), 0, 0, 1, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLeveTimerActiveBall(playsoundparams, aVol, ballvariable)
  PlaySound playsoundparams, 0, min(1,aVol) * VolumeDial, AudioPan(ballvariable), 0, 0, 0, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelTimerExistingActiveBall(playsoundparams, aVol, ballvariable)
  PlaySound playsoundparams, 0, min(1,aVol) * VolumeDial, AudioPan(ballvariable), 0, 0, 1, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelRoll(playsoundparams, aVol, pitch)
  PlaySound playsoundparams, - 1, min(1,aVol) * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

' Previous Positional Sound Subs

Sub PlaySoundAt(soundname, tableobj)
  PlaySound soundname, 1, 1 * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtVol(soundname, tableobj, aVol)
  PlaySound soundname, 1, min(1,aVol) * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname)
  PlaySoundAt soundname, ActiveBall
End Sub

Sub PlaySoundAtBallVol (Soundname, aVol)
  Playsound soundname, 1,min(1,aVol) * VolumeDial, AudioPan(ActiveBall), 0,0,0, 1, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtBallVolM (Soundname, aVol)
  Playsound soundname, 1,min(1,aVol) * VolumeDial, AudioPan(ActiveBall), 0,0,0, 0, AudioFade(ActiveBall)
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
  finalspeed = Sqr(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
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
  finalspeed = Sqr(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
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
  finalspeed = Sqr(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
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
  If Abs(cor.ballvelx(activeball.id) < 4) And cor.ballvely(activeball.id) > 7 Then
    RandomSoundBottomArchBallGuideHardHit()
  Else
    RandomSoundBottomArchBallGuide
  End If
End Sub

'/////////////////////////////  FLIPPER BALL GUIDE  ////////////////////////////

Sub RandomSoundFlipperBallGuide()
  Dim finalspeed
  finalspeed = Sqr(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
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
  finalspeed = Sqr(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 10 Then
    RandomSoundTargetHitStrong()
    RandomSoundBallBouncePlayfieldSoft Activeball
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
  PlaySoundAtLevelStatic ("Gate_FastTrigger_" & Int(Rnd * 2) + 1), GateSoundLevel, Activeball
End Sub

Sub SoundHeavyGate()
  PlaySoundAtLevelStatic ("Gate_2"), GateSoundLevel, Activeball
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
  If Activeball.velx > 1 Then SoundPlayfieldGate
  StopSound "Arch_L1"
  StopSound "Arch_L2"
  StopSound "Arch_L3"
  StopSound "Arch_L4"
End Sub

Sub Arch1_unhit()
  If activeball.velx <  - 8 Then
    RandomSoundRightArch
  End If
End Sub

Sub Arch2_hit()
  If Activeball.velx < 1 Then SoundPlayfieldGate
  StopSound "Arch_R1"
  StopSound "Arch_R2"
  StopSound "Arch_R3"
  StopSound "Arch_R4"
End Sub

Sub Arch2_unhit()
  If activeball.velx > 10 Then
    RandomSoundLeftArch
  End If
End Sub

'/////////////////////////////  SAUCERS (KICKER HOLES)  ////////////////////////////

Sub SoundSaucerLock()
  PlaySoundAtLevelStatic ("Saucer_Enter_" & Int(Rnd * 2) + 1), SaucerLockSoundLevel, Activeball
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
'****  FLEEP MECHANICAL SOUNDS
'******************************************************




'******************************************************
'   ZRDT:  DROP TARGETS by Rothbauerw
'******************************************************


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
Dim DT1, DT2, DT3, DT4, DT5, DT7, DT45

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

Set DT1 = (new DropTarget)(sw1, sw1a, BM_sw1, 1, 0, False)
Set DT2 = (new DropTarget)(sw2, sw2a, BM_sw2, 2, 0, False)
Set DT3 = (new DropTarget)(sw3, sw3a, BM_sw3, 3, 0, False)
Set DT4 = (new DropTarget)(sw4, sw4a, BM_sw4, 4, 0, False)
Set DT5 = (new DropTarget)(sw5, sw5a, BM_sw5, 5, 0, False)
Set DT7 = (new DropTarget)(sw7, sw7a, BM_sw7, 7, 0, False)
Set DT45 = (new DropTarget)(sw45, sw45a, BM_sw45, 45, 0, False)

Dim DTArray
DTArray = Array(DT1, DT2, DT3, DT4, DT5, DT7, DT45)

'Configure the behavior of Drop Targets.
Const DTDropSpeed = 90 'in milliseconds
Const DTDropUpSpeed = 40 'in milliseconds
Const DTDropUnits = 44 'VP units primitive drops so top of at or below the playfield
Const DTDropUpUnits = 10 'VP units primitive raises above the up position on drops up
Const DTMaxBend = 8 'max degrees primitive rotates when hit
Const DTDropDelay = 20 'time in milliseconds before target drops (due to friction/impact of the ball)
Const DTRaiseDelay = 40 'time in milliseconds before target drops back to normal up position after the solenoid fires to raise the target
Const DTBrickVel = 30 'velocity at which the target will brick, set to '0' to disable brick
Const DTEnableBrick = 0 'Set to 0 to disable bricking, 1 to enable bricking
Const DTMass = 0.2 'Mass of the Drop Target (between 0 and 1), higher values provide more resistance



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
      'DTAction Switchid
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




'******************************************************
' ZRST: STAND-UP TARGETS by Rothbauerw
'******************************************************

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
Dim ST17, ST18, ST19, ST20, ST21, ST33, ST34, ST35, ST36

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

Set ST17 = (new StandupTarget)(sw17, BM_sw17, 17, 0)
Set ST18 = (new StandupTarget)(sw18, BM_sw18, 18, 0)
Set ST19 = (new StandupTarget)(sw19, BM_sw19, 19, 0)
Set ST20 = (new StandupTarget)(sw20, BM_sw20, 20, 0)
Set ST21 = (new StandupTarget)(sw21, BM_sw21, 21, 0)
Set ST33 = (new StandupTarget)(sw33, BM_sw33, 33, 0)
Set ST34 = (new StandupTarget)(sw34, BM_sw34, 34, 0)
Set ST35 = (new StandupTarget)(sw35, BM_sw35, 35, 0)
Set ST36 = (new StandupTarget)(sw36, BM_sw36, 36, 0)

'Add all the Stand-up Target Arrays to Stand-up Target Animation Array
'   STAnimationArray = Array(ST1, ST2, ....)
Dim STArray
STArray = Array(ST17, ST18, ST19, ST20, ST21, ST33, ST34, ST35, ST36)

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
'****   END STAND-UP TARGETS
'******************************************************






'******************************************************
'   ZANI: Misc Animations
'******************************************************


''''' Flippers


Sub LeftFlipper_Animate
  Dim a : a = LeftFlipper.CurrentAngle
  FlipperLSh.RotZ = a

  Dim v, BP
  v = 255.0 * (122.0 - LeftFlipper.CurrentAngle) / (122.0 -  73.0)

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
  v = 255.0 * (-122.0 - RightFlipper.CurrentAngle) / (-122.0 +  73.0)

  For each BP in BP_RFlipper
    BP.Rotz = a
    BP.visible = v < 128.0
  Next
  For each BP in BP_RFlipperU
    BP.Rotz = a
    BP.visible = v >= 128.0
  Next
End Sub

Sub RightFlipper1_Animate
  Dim a : a = RightFlipper1.CurrentAngle
  FlipperRSh001.RotZ = a

  Dim v, BP
  v = 255.0 * (198.0 - RightFlipper1.CurrentAngle) / (198.0 -  253.0)

  For each BP in BP_RFlipper1
    BP.Rotz = a
    BP.visible = v < 128.0
  Next
  For each BP in BP_RFlipper1U
    BP.Rotz = a
    BP.visible = v >= 128.0
  Next
End Sub



''''' Bumper Animations

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
    If r = 0 Then For Each x in BP_BS1: x.transZ = tz: Next
    If r = 1 Then For Each x in BP_BS2: x.transZ = tz: Next
    If r = 2 Then For Each x in BP_BS3: x.transZ = tz: Next
  Next
End Sub



''''' Gate Animations

Sub Gate1_Animate
  Dim a : a = Gate1.CurrentAngle
  Dim BP : For Each BP in BP_Gate1 : BP.rotx = a: Next
End Sub

Sub Gate2_Animate
  Dim a : a = Gate2.CurrentAngle
  Dim BP : For Each BP in BP_Gate2 : BP.rotx = a: Next
End Sub

Sub Gate3_Animate
  Dim a : a = Gate3.CurrentAngle
  Dim BP : For Each BP in BP_Gate3 : BP.rotx = a: Next
End Sub

Sub Gate4_Animate
  Dim a : a = Gate4.CurrentAngle
  Dim BP : For Each BP in BP_Gate4 : BP.rotx = a: Next
End Sub

Sub Gate001_Animate
  Dim a : a = Gate001.CurrentAngle
  Dim BP : For Each BP in BP_Gate001 : BP.rotx = a: Next
End Sub



''''' Switch Animations

Sub sw13_Animate
  Dim z : z = sw13.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw13 : BP.transz = z: Next
End Sub

Sub sw16_Animate
  Dim z : z = sw16.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw16 : BP.transz = z: Next
End Sub

Sub sw23_Animate
  Dim z : z = sw23.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw23 : BP.transz = z: Next
End Sub

Sub sw24_Animate
  Dim z : z = sw24.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw24 : BP.transz = z: Next
End Sub

Sub sw30_Animate
  Dim z : z = sw30.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw30 : BP.transz = z: Next
End Sub

Sub sw31_Animate
  Dim z : z = sw31.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw31 : BP.transz = z: Next
End Sub

Sub sw32_Animate
  Dim z : z = sw32.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw32 : BP.transz = z: Next
End Sub

Sub sw37_Animate
  Dim z : z = sw37.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw37 : BP.transz = z: Next
End Sub

Sub sw38_Animate
  Dim z : z = sw38.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw38 : BP.transz = z: Next
End Sub

Sub sw39_Animate
  Dim z : z = sw39.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw39 : BP.transz = z: Next
End Sub

Sub sw40_Animate
  Dim z : z = sw40.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw40 : BP.transz = z: Next
End Sub


''''' Standup Targets

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

  ty = BM_sw21.transy
  For each BP in BP_sw21 : BP.transy = ty: Next

  ty = BM_sw33.transy
  For each BP in BP_sw33 : BP.transy = ty: Next

  ty = BM_sw34.transy
  For each BP in BP_sw34 : BP.transy = ty: Next

  ty = BM_sw35.transy
  For each BP in BP_sw35 : BP.transy = ty: Next

  ty = BM_sw36.transy
  For each BP in BP_sw36 : BP.transy = ty: Next

End Sub



''''' Drop Targets

Sub UpdateDropTargets
  dim BP, tz, rx, ry

  tz = BM_sw1.transz
  rx = BM_sw1.rotx
  ry = BM_sw1.roty
  For each BP in BP_sw1: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw2.transz
  rx = BM_sw2.rotx
  ry = BM_sw2.roty
  For each BP in BP_sw2: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw3.transz
  rx = BM_sw3.rotx
  ry = BM_sw3.roty
  For each BP in BP_sw3: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw4.transz
  rx = BM_sw4.rotx
  ry = BM_sw4.roty
  For each BP in BP_sw4: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw5.transz
  rx = BM_sw5.rotx
  ry = BM_sw5.roty
  For each BP in BP_sw5: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw7.transz
  rx = BM_sw7.rotx
  ry = BM_sw7.roty
  For each BP in BP_sw7: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw45.transz
  rx = BM_sw45.rotx
  ry = BM_sw45.roty
  For each BP in BP_sw45: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

End Sub


'******************************
' ZVRR - Setup VR Room
'******************************

Sub SetupVRRoom
  dim VRThing, BP
  For Each BP in BP_PincabRails : BP.Visible = 0: Next
  If RenderingMode = 2 or ForceVR = True Then
    VRMode = True
    For Each VRThing in VRThings: VRThing.visible = 1: Next
  Else
    VRMode = False
    For Each VRThing in VRThings: VRThing.visible = 0: Next
    If DesktopMode = True Then For Each BP in BP_PincabRails : BP.Visible = 1: Next
  End If
End Sub


Sub TimerVRPlunger_Timer
  If VR_Shooter.Y < 100 then
    VR_Shooter.Y = VR_Shooter.Y + 5
  End If
End Sub

Sub TimerVRPlunger2_Timer
  VR_Shooter.Y = (5* Plunger.Position) - 20
End Sub


'******************************
' ZVRB - Setup VR Backglass
'******************************

Dim xoff,yoff1, yoff2, yoff3, yoff4, yoff5, zoff,xrot,zscale, xcen,ycen

Sub SetupVRBackglass()

  xoff = -20
  yoff1 = 33 ' this is where you adjust the forward/backward position for player 1 score
  yoff2 = 23 ' this is where you adjust the forward/backward position for player 2 score
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

dim DisplayColor

DisplayColor =  RGB(255,40,1)

Sub VRBackglassTimer
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




'*******************************************
'  ZCHG: Change Log
'*******************************************

' pre-78 - armyaviation - Initial table development, including playfield redraw
' 78 - apophis - Updated flippers to mid-80's Bally style. Fixed some physics script issues. Added slingshot corrections
' 79 - apophis - Fleep sounds. Rolling sounds. Physical trough. Dynamic shadows. GI light repositioning.
' 80 - apophis - Lampz. Roth DTs and STs.
' 81 - apophis - Added lane change switch. Physics tuning.
' 82 - apophis - Added flashers
' 83 - apophis - Fixed upper right gate stuck ball issue.
' 84 - apophis - Fixed insert lamp assignments. Increased bumper strength. Fixed WallULeft physics material.
' 87 - apophis - Initial toolkit script implementation. Removed Lampz. Removed old LUT crap. Ambient shadows only.
' 88 - apophis - Test toolkit import. Alignment of physical objects. Added all animations. Updated roth target code.
' 90 - apophis - 2k batch imported.
' 91 - apophis - 4k batch imported. Different env hdri
' 92 - apophis - New 4k batch imported. Original env hdri. LUT removed. File cleanups.
' 93 - apophis - New 4k batch imported.
' 94 - apophis - New 4k batch imported.
' 95 - apophis - Meshes and nestmaps remade after influence map fix from Flux.
' 96 - apophis - Added tweak menu. Added some VR room stuffv not working yet). Inc bumpert strength a little. Update gate physics params.
' 97 - TastyWasps - Basic VR digits and VR backglass added.  Placements will depend on new VR cabinet so X/Y/Z placements will be made later.
' 98 - apophis - Updated VR room and cab
' 99 - apophis - Updated VR cab (added keys, coin in plastics, new texture)
' 100 - apophis - Updated desktop backdrop
' 101 - apophis - Fixed playfield reflections
' 102 - apophis - Changed tonemap. Removed stray kickers. Lowered plunger pull speed. VR: Fixed button smoothness. Backglass position tweaks. Removed swinging key. Added posters on wall (thanks DarthVito)
' 103 - apophis - Adjusted VR metal material.
' Release v1.0
' 1.0.1 - apophis - Fixed timers causing stutters.



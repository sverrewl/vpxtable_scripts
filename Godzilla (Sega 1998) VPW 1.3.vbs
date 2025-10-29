'          ________           .___      .__.__  .__
'         /  _____/  ____   __| _/______|__|  | |  | _____
'        /   \  ___ /  _ \ / __ |\___   /  |  | |  | \__  \
'        \    \_\  (  <_> ) /_/ | /    /|  |  |_|  |__/ __ \_
'         \______  /\____/\____ |/_____ \__|____/____(____  /
'                \/            \/      \/                 \/
'
' SEGA Pinball, 1998
'
' https://www.ipdb.org/machine.cgi?id=4443
'
'
' VPW TEAM
' --------
' TastyWasps - Project Lead, nFozzy, Fleep, toolkit script-side integration, VR room
' Apophis - Project Lead, toolkit script-side integration, various scripting tweaks, mentoring
' Tomate - Graphics modeling and rendering
' Sixtoe - Table full scratch build and various scripting, VR support
' Iaakki - Scripting tweaks
' EBisLit - Playfield, plastics, inserts & 3d scans with measurements
' ClarkKent - Playfield Stitch
' Niwak - Blender Toolkit
' HauntFreaks - Backglass image
' Testing - PinStratsDan, BountyBob, Bietekwiet, Studlygoorite, Smaug, Bord, VPW Team
' Toxie - VPX and pushing out the newest 10.8 as a beta so we could release.


'//////////////////////////////////////////////////////////////////////
' THIS TABLE INCLUDES THE IN-GAME OPTIONS MENU
' To open the menu, press F12
'//////////////////////////////////////////////////////////////////////


Option Explicit
Randomize
SetLocale 1033

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0


'*******************************************
' Constants and Global Variables
'*******************************************
Const BallSize = 50     ' Ball size must be 50
Const BallMass = 1      ' Ball mass must be 1
Const tnob = 7        ' Total number of balls
Const lob = 0       ' Total number of locked balls

Dim tablewidth: tablewidth = Table1.width
Dim tableheight: tableheight = Table1.height

Const cGameName="godzilla",UseSolenoids=2,UseLamps=1,UseGI=0, SCoin=""
Const UseVPMModSol = 2


' VLM Arrays - Start
' Arrays per baked part
Dim BP_BS1: BP_BS1=Array(BM_BS1, LM_GI_BS1, LM_inserts_L45_BS1, LM_inserts_L46_BS1, LM_inserts_L47_BS1, LM_flashers_f07_BS1, LM_flashers_f1_BS1, LM_flashers_f19_BS1, LM_flashers_f1a_BS1, LM_flashers_f1b_BS1, LM_flashers_f2_BS1, LM_flashers_f2a_BS1, LM_flashers_f8_BS1, LM_flashers_f8a_BS1)
Dim BP_BS2: BP_BS2=Array(BM_BS2, LM_GI_BS2, LM_inserts_L45_BS2, LM_inserts_L46_BS2, LM_inserts_L47_BS2, LM_flashers_f1_BS2, LM_flashers_f19_BS2, LM_flashers_f1b_BS2, LM_flashers_f1c_BS2, LM_flashers_f2_BS2, LM_flashers_f2a_BS2, LM_flashers_f4_BS2, LM_flashers_f6_BS2, LM_flashers_f8_BS2, LM_flashers_f8a_BS2)
Dim BP_BS3: BP_BS3=Array(BM_BS3, LM_GI_BS3, LM_inserts_L45_BS3, LM_inserts_L46_BS3, LM_inserts_L47_BS3, LM_flashers_f19_BS3, LM_flashers_f1c_BS3, LM_flashers_f3_BS3, LM_flashers_f6_BS3, LM_flashers_f8_BS3)
Dim BP_BallReleaseGate: BP_BallReleaseGate=Array(BM_BallReleaseGate)
Dim BP_Bumper1: BP_Bumper1=Array(BM_Bumper1, LM_GI_Bumper1, LM_inserts_L45_Bumper1, LM_inserts_L46_Bumper1, LM_inserts_L47_Bumper1, LM_flashers_f1_Bumper1, LM_flashers_f19_Bumper1, LM_flashers_f1a_Bumper1, LM_flashers_f1b_Bumper1, LM_flashers_f2_Bumper1, LM_flashers_f2a_Bumper1)
Dim BP_Bumper2: BP_Bumper2=Array(BM_Bumper2, LM_GI_Bumper2, LM_inserts_L45_Bumper2, LM_inserts_L46_Bumper2, LM_inserts_L47_Bumper2, LM_flashers_f19_Bumper2, LM_flashers_f1a_Bumper2, LM_flashers_f1b_Bumper2, LM_flashers_f1c_Bumper2, LM_flashers_f2a_Bumper2, LM_flashers_f8_Bumper2)
Dim BP_Bumper3: BP_Bumper3=Array(BM_Bumper3, LM_GI_Bumper3, LM_inserts_L45_Bumper3, LM_inserts_L46_Bumper3, LM_inserts_L47_Bumper3, LM_flashers_f1c_Bumper3)
Dim BP_Gate1: BP_Gate1=Array(BM_Gate1, LM_GI_Gate1, LM_inserts_L45_Gate1, LM_flashers_f1_Gate1, LM_flashers_f1a_Gate1)
Dim BP_Gate2: BP_Gate2=Array(BM_Gate2, LM_GI_Gate2, LM_inserts_L45_Gate2, LM_inserts_L46_Gate2, LM_inserts_L47_Gate2, LM_flashers_f1_Gate2, LM_flashers_f2a_Gate2, LM_flashers_f8_Gate2, LM_flashers_f8a_Gate2)
Dim BP_LEMK: BP_LEMK=Array(BM_LEMK, LM_GI_splitted_GI_14_LEMK, LM_GI_LEMK, LM_GI_splitted_GI_22_LEMK, LM_GI_splitted_GI_3_LEMK, LM_GI_splitted_GI_6_LEMK)
Dim BP_LFlipper: BP_LFlipper=Array(BM_LFlipper, LM_GI_splitted_GI_14_LFlipper, LM_GI_splitted_GI_16_LFlipper, LM_GI_splitted_GI_2_LFlipper, LM_GI_LFlipper, LM_GI_splitted_GI_22_LFlipper, LM_GI_splitted_GI_3_LFlipper, LM_GI_splitted_GI_6_LFlipper, LM_GI_splitted_GI_8_LFlipper)
Dim BP_LSling1: BP_LSling1=Array(BM_LSling1, LM_GI_splitted_GI_14_LSling1, LM_GI_splitted_GI_22_LSling1)
Dim BP_LSling2: BP_LSling2=Array(BM_LSling2, LM_GI_splitted_GI_14_LSling2, LM_GI_splitted_GI_22_LSling2, LM_GI_splitted_GI_3_LSling2, LM_GI_splitted_GI_6_LSling2, LM_inserts_L1_LSling2, LM_flashers_f18_LSling2, LM_flashers_f4_LSling2, LM_flashers_f7lb_LSling2, LM_flashers_f7lt_LSling2)
Dim BP_Layer0: BP_Layer0=Array(BM_Layer0, LM_GI_Layer0, LM_inserts_L38_Layer0, LM_inserts_L46_Layer0, LM_flashers_f1_Layer0, LM_flashers_f6_Layer0, LM_flashers_f8_Layer0, LM_flashers_f8a_Layer0)
Dim BP_Layer1: BP_Layer1=Array(BM_Layer1, LM_GI_splitted_GI_10_Layer1, LM_GI_splitted_GI_12_Layer1, LM_GI_splitted_GI_14_Layer1, LM_GI_splitted_GI_16_Layer1, LM_GI_splitted_GI_2_Layer1, LM_GI_Layer1, LM_GI_splitted_GI_22_Layer1, LM_GI_splitted_GI_24_Layer1, LM_GI_splitted_GI_26_Layer1, LM_GI_splitted_GI_3_Layer1, LM_GI_splitted_GI_30_Layer1, LM_GI_splitted_GI_6_Layer1, LM_GI_splitted_GI_8_Layer1, LM_inserts_L12_Layer1, LM_inserts_L13_Layer1, LM_inserts_L14_Layer1, LM_inserts_L18_Layer1, LM_inserts_L29_Layer1, LM_inserts_L20_Layer1, LM_inserts_L21_Layer1, LM_inserts_L22_Layer1, LM_inserts_L23_Layer1, LM_inserts_L25_Layer1, LM_inserts_L26_Layer1, LM_inserts_L27_Layer1, LM_inserts_L30_Layer1, LM_inserts_L31_Layer1, LM_inserts_L36_Layer1, LM_inserts_L37_Layer1, LM_inserts_L38_Layer1, LM_inserts_L45_Layer1, LM_inserts_L46_Layer1, LM_inserts_L47_Layer1, LM_inserts_L49_Layer1, LM_inserts_L50_Layer1, LM_inserts_L51_Layer1, LM_inserts_L52_Layer1, LM_inserts_L54_Layer1, LM_flashers_f07_Layer1, LM_flashers_f1_Layer1, _
  LM_flashers_f18_Layer1, LM_flashers_f19_Layer1, LM_flashers_f1a_Layer1, LM_flashers_f1b_Layer1, LM_flashers_f1c_Layer1, LM_flashers_f2_Layer1, LM_flashers_f20a_Layer1, LM_flashers_f20b_Layer1, LM_flashers_f2a_Layer1, LM_flashers_f3_Layer1, LM_flashers_f4_Layer1, LM_flashers_f5_Layer1, LM_flashers_f6_Layer1, LM_flashers_f8_Layer1, LM_flashers_f8a_Layer1)
Dim BP_Playfield: BP_Playfield=Array(BM_Playfield, LM_GI_splitted_GI_10_Playfield, LM_GI_splitted_GI_12_Playfield, LM_GI_splitted_GI_14_Playfield, LM_GI_splitted_GI_16_Playfield, LM_GI_splitted_GI_2_Playfield, LM_GI_Playfield, LM_GI_splitted_GI_22_Playfield, LM_GI_splitted_GI_24_Playfield, LM_GI_splitted_GI_26_Playfield, LM_GI_splitted_GI_3_Playfield, LM_GI_splitted_GI_30_Playfield, LM_GI_splitted_GI_6_Playfield, LM_GI_splitted_GI_8_Playfield, LM_inserts_L15_Playfield, LM_inserts_L16_Playfield, LM_inserts_L17_Playfield, LM_inserts_L23_Playfield, LM_inserts_L24_Playfield, LM_inserts_L28_Playfield, LM_inserts_L2_Playfield, LM_inserts_L32_Playfield, LM_inserts_L33_Playfield, LM_inserts_L34_Playfield, LM_inserts_L35_Playfield, LM_inserts_L36_Playfield, LM_inserts_L37_Playfield, LM_inserts_L38_Playfield, LM_inserts_L39_Playfield, LM_inserts_L41_Playfield, LM_inserts_L42_Playfield, LM_inserts_L43_Playfield, LM_inserts_L44_Playfield, LM_inserts_L45_Playfield, LM_inserts_L46_Playfield, LM_inserts_L47_Playfield, _
  LM_inserts_L53_Playfield, LM_inserts_L54_Playfield, LM_inserts_L55_Playfield, LM_inserts_L56_Playfield, LM_inserts_L5_Playfield, LM_inserts_L7_Playfield, LM_flashers_f07_Playfield, LM_flashers_f1_Playfield, LM_flashers_f18_Playfield, LM_flashers_f19_Playfield, LM_flashers_f1c_Playfield, LM_flashers_f2_Playfield, LM_flashers_f20a_Playfield, LM_flashers_f20b_Playfield, LM_flashers_f2a_Playfield, LM_flashers_f3_Playfield, LM_flashers_f4_Playfield, LM_flashers_f5_Playfield, LM_flashers_f6_Playfield, LM_flashers_f7lt_Playfield, LM_flashers_f8_Playfield, LM_flashers_f8a_Playfield)
Dim BP_REMK: BP_REMK=Array(BM_REMK, LM_GI_splitted_GI_16_REMK, LM_GI_splitted_GI_2_REMK, LM_GI_splitted_GI_24_REMK, LM_GI_splitted_GI_8_REMK, LM_flashers_f18_REMK)
Dim BP_RFlipper: BP_RFlipper=Array(BM_RFlipper, LM_GI_splitted_GI_14_RFlipper, LM_GI_splitted_GI_16_RFlipper, LM_GI_splitted_GI_2_RFlipper, LM_GI_RFlipper, LM_GI_splitted_GI_24_RFlipper, LM_GI_splitted_GI_3_RFlipper, LM_GI_splitted_GI_6_RFlipper, LM_GI_splitted_GI_8_RFlipper)
Dim BP_RSling1: BP_RSling1=Array(BM_RSling1, LM_GI_splitted_GI_16_RSling1, LM_GI_splitted_GI_24_RSling1, LM_flashers_f18_RSling1)
Dim BP_RSling2: BP_RSling2=Array(BM_RSling2, LM_GI_splitted_GI_16_RSling2, LM_GI_splitted_GI_2_RSling2, LM_GI_splitted_GI_24_RSling2, LM_GI_splitted_GI_8_RSling2, LM_flashers_f18_RSling2, LM_flashers_f19_RSling2)
Dim BP_lPost: BP_lPost=Array(BM_lPost, LM_GI_splitted_GI_12_lPost, LM_GI_lPost, LM_GI_splitted_GI_22_lPost, LM_GI_splitted_GI_30_lPost)
Dim BP_parts: BP_parts=Array(BM_parts, LM_GI_splitted_GI_10_parts, LM_GI_splitted_GI_12_parts, LM_GI_splitted_GI_14_parts, LM_GI_splitted_GI_16_parts, LM_GI_splitted_GI_2_parts, LM_GI_parts, LM_GI_splitted_GI_22_parts, LM_GI_splitted_GI_24_parts, LM_GI_splitted_GI_26_parts, LM_GI_splitted_GI_3_parts, LM_GI_splitted_GI_30_parts, LM_GI_splitted_GI_6_parts, LM_GI_splitted_GI_8_parts, LM_inserts_L12_parts, LM_inserts_L13_parts, LM_inserts_L14_parts, LM_inserts_L17_parts, LM_inserts_L18_parts, LM_inserts_L19_parts, LM_inserts_L29_parts, LM_inserts_L1_parts, LM_inserts_L20_parts, LM_inserts_L21_parts, LM_inserts_L22_parts, LM_inserts_L23_parts, LM_inserts_L24_parts, LM_inserts_L25_parts, LM_inserts_L26_parts, LM_inserts_L27_parts, LM_inserts_L28_parts, LM_inserts_L30_parts, LM_inserts_L31_parts, LM_inserts_L32_parts, LM_inserts_L35_parts, LM_inserts_L36_parts, LM_inserts_L37_parts, LM_inserts_L38_parts, LM_inserts_L45_parts, LM_inserts_L46_parts, LM_inserts_L47_parts, LM_inserts_L49_parts, LM_inserts_L50_parts, _
  LM_inserts_L51_parts, LM_inserts_L52_parts, LM_inserts_L53_parts, LM_inserts_L55_parts, LM_inserts_L6_parts, LM_flashers_f07_parts, LM_flashers_f1_parts, LM_flashers_f18_parts, LM_flashers_f19_parts, LM_flashers_f1a_parts, LM_flashers_f1b_parts, LM_flashers_f1c_parts, LM_flashers_f2_parts, LM_flashers_f20a_parts, LM_flashers_f20b_parts, LM_flashers_f2a_parts, LM_flashers_f3_parts, LM_flashers_f4_parts, LM_flashers_f5_parts, LM_flashers_f6_parts, LM_flashers_f8_parts, LM_flashers_f8a_parts)
Dim BP_rPost: BP_rPost=Array(BM_rPost, LM_GI_splitted_GI_10_rPost, LM_GI_rPost, LM_GI_splitted_GI_24_rPost, LM_GI_splitted_GI_26_rPost, LM_flashers_f19_rPost)
Dim BP_sw16: BP_sw16=Array(BM_sw16)
Dim BP_sw17: BP_sw17=Array(BM_sw17, LM_GI_splitted_GI_12_sw17, LM_GI_sw17, LM_GI_splitted_GI_30_sw17, LM_inserts_L26_sw17, LM_flashers_f18_sw17)
Dim BP_sw18: BP_sw18=Array(BM_sw18, LM_GI_splitted_GI_12_sw18, LM_GI_sw18, LM_GI_splitted_GI_30_sw18, LM_inserts_L27_sw18, LM_flashers_f18_sw18)
Dim BP_sw19: BP_sw19=Array(BM_sw19, LM_GI_splitted_GI_12_sw19, LM_GI_sw19, LM_GI_splitted_GI_30_sw19, LM_flashers_f07_sw19, LM_flashers_f18_sw19)
Dim BP_sw20: BP_sw20=Array(BM_sw20, LM_GI_sw20, LM_GI_splitted_GI_24_sw20, LM_GI_splitted_GI_3_sw20, LM_GI_splitted_GI_30_sw20, LM_flashers_f18_sw20, LM_flashers_f19_sw20, LM_flashers_f4_sw20, LM_flashers_f5_sw20)
Dim BP_sw21: BP_sw21=Array(BM_sw21, LM_GI_splitted_GI_10_sw21, LM_GI_sw21, LM_GI_splitted_GI_26_sw21, LM_inserts_L30_sw21, LM_flashers_f19_sw21, LM_flashers_f4_sw21, LM_flashers_f5_sw21)
Dim BP_sw22: BP_sw22=Array(BM_sw22, LM_GI_splitted_GI_10_sw22, LM_GI_sw22, LM_GI_splitted_GI_26_sw22, LM_inserts_L31_sw22, LM_flashers_f19_sw22, LM_flashers_f4_sw22, LM_flashers_f5_sw22)
Dim BP_sw23: BP_sw23=Array(BM_sw23, LM_GI_splitted_GI_10_sw23, LM_GI_sw23, LM_GI_splitted_GI_26_sw23, LM_flashers_f19_sw23, LM_flashers_f4_sw23, LM_flashers_f5_sw23)
Dim BP_sw24: BP_sw24=Array(BM_sw24, LM_GI_splitted_GI_2_sw24, LM_GI_sw24, LM_GI_splitted_GI_22_sw24, LM_GI_splitted_GI_26_sw24, LM_flashers_f19_sw24, LM_flashers_f4_sw24, LM_flashers_f5_sw24)
Dim BP_sw26: BP_sw26=Array(BM_sw26, LM_GI_sw26, LM_flashers_f2_sw26, LM_flashers_f2a_sw26)
Dim BP_sw28: BP_sw28=Array(BM_sw28, LM_GI_sw28)
Dim BP_sw29: BP_sw29=Array(BM_sw29, LM_GI_splitted_GI_16_sw29, LM_GI_splitted_GI_2_sw29, LM_GI_sw29, LM_GI_splitted_GI_22_sw29, LM_GI_splitted_GI_6_sw29, LM_GI_splitted_GI_8_sw29, LM_inserts_L41_sw29, LM_flashers_f07_sw29, LM_flashers_f19_sw29, LM_flashers_f3_sw29, LM_flashers_f4_sw29, LM_flashers_f5_sw29, LM_flashers_f6_sw29)
Dim BP_sw30: BP_sw30=Array(BM_sw30, LM_GI_splitted_GI_14_sw30, LM_GI_splitted_GI_16_sw30, LM_GI_sw30, LM_GI_splitted_GI_22_sw30, LM_GI_splitted_GI_6_sw30, LM_GI_splitted_GI_8_sw30, LM_flashers_f19_sw30, LM_flashers_f3_sw30, LM_flashers_f4_sw30, LM_flashers_f5_sw30, LM_flashers_f6_sw30)
Dim BP_sw31: BP_sw31=Array(BM_sw31, LM_GI_splitted_GI_14_sw31, LM_GI_splitted_GI_16_sw31, LM_GI_sw31, LM_inserts_L44_sw31, LM_flashers_f19_sw31, LM_flashers_f3_sw31, LM_flashers_f4_sw31, LM_flashers_f5_sw31, LM_flashers_f6_sw31)
Dim BP_sw32: BP_sw32=Array(BM_sw32, LM_GI_splitted_GI_16_sw32, LM_GI_sw32, LM_flashers_f19_sw32, LM_flashers_f3_sw32, LM_flashers_f4_sw32, LM_flashers_f5_sw32, LM_flashers_f6_sw32)
Dim BP_sw41: BP_sw41=Array(BM_sw41, LM_GI_sw41, LM_flashers_f19_sw41)
Dim BP_sw42: BP_sw42=Array(BM_sw42, LM_GI_sw42, LM_flashers_f3_sw42)
Dim BP_sw43: BP_sw43=Array(BM_sw43, LM_GI_sw43, LM_inserts_L38_sw43, LM_flashers_f6_sw43)
Dim BP_sw44: BP_sw44=Array(BM_sw44, LM_GI_sw44)
Dim BP_sw45: BP_sw45=Array(BM_sw45, LM_GI_sw45, LM_inserts_L45_sw45, LM_inserts_L47_sw45, LM_flashers_f07_sw45, LM_flashers_f2_sw45, LM_flashers_f20a_sw45)
Dim BP_sw46: BP_sw46=Array(BM_sw46, LM_GI_sw46, LM_inserts_L23_sw46, LM_inserts_L45_sw46, LM_flashers_f07_sw46, LM_flashers_f18_sw46, LM_flashers_f3_sw46, LM_flashers_f6_sw46)
Dim BP_sw47: BP_sw47=Array(BM_sw47, LM_GI_sw47)
Dim BP_sw47_001: BP_sw47_001=Array(BM_sw47_001, LM_GI_sw47_001)
Dim BP_sw57: BP_sw57=Array(BM_sw57, LM_GI_splitted_GI_14_sw57, LM_GI_splitted_GI_22_sw57, LM_GI_splitted_GI_3_sw57, LM_GI_splitted_GI_6_sw57)
Dim BP_sw58: BP_sw58=Array(BM_sw58, LM_GI_splitted_GI_14_sw58, LM_GI_splitted_GI_22_sw58, LM_GI_splitted_GI_3_sw58, LM_GI_splitted_GI_6_sw58, LM_flashers_f18_sw58)
Dim BP_sw60: BP_sw60=Array(BM_sw60, LM_GI_splitted_GI_16_sw60, LM_GI_splitted_GI_2_sw60, LM_GI_sw60, LM_GI_splitted_GI_24_sw60, LM_GI_splitted_GI_8_sw60)
Dim BP_sw61: BP_sw61=Array(BM_sw61, LM_GI_splitted_GI_16_sw61, LM_GI_splitted_GI_2_sw61, LM_GI_splitted_GI_24_sw61, LM_GI_splitted_GI_8_sw61, LM_flashers_f19_sw61)
Dim BP_toys: BP_toys=Array(BM_toys, LM_GI_splitted_GI_12_toys, LM_GI_toys, LM_GI_splitted_GI_30_toys, LM_inserts_L13_toys, LM_inserts_L14_toys, LM_inserts_L18_toys, LM_inserts_L20_toys, LM_inserts_L21_toys, LM_inserts_L22_toys, LM_inserts_L23_toys, LM_inserts_L34_toys, LM_inserts_L35_toys, LM_inserts_L41_toys, LM_inserts_L42_toys, LM_inserts_L43_toys, LM_inserts_L45_toys, LM_inserts_L46_toys, LM_inserts_L47_toys, LM_inserts_L54_toys, LM_inserts_L55_toys, LM_inserts_L56_toys, LM_flashers_f07_toys, LM_flashers_f1_toys, LM_flashers_f18_toys, LM_flashers_f19_toys, LM_flashers_f1a_toys, LM_flashers_f1b_toys, LM_flashers_f1c_toys, LM_flashers_f2_toys, LM_flashers_f20a_toys, LM_flashers_f20b_toys, LM_flashers_f2a_toys, LM_flashers_f3_toys, LM_flashers_f4_toys, LM_flashers_f5_toys, LM_flashers_f6_toys, LM_flashers_f8_toys, LM_flashers_f8a_toys)
Dim BP_underPF: BP_underPF=Array(BM_underPF, LM_GI_splitted_GI_10_underPF, LM_GI_splitted_GI_14_underPF, LM_GI_splitted_GI_16_underPF, LM_GI_underPF, LM_GI_splitted_GI_22_underPF, LM_GI_splitted_GI_24_underPF, LM_GI_splitted_GI_26_underPF, LM_GI_splitted_GI_30_underPF, LM_inserts_L10_underPF, LM_inserts_L11_underPF, LM_inserts_L12_underPF, LM_inserts_L13_underPF, LM_inserts_L14_underPF, LM_inserts_L15_underPF, LM_inserts_L16_underPF, LM_inserts_L17_underPF, LM_inserts_L18_underPF, LM_inserts_L19_underPF, LM_inserts_L29_underPF, LM_inserts_L1_underPF, LM_inserts_L20_underPF, LM_inserts_L21_underPF, LM_inserts_L22_underPF, LM_inserts_L23_underPF, LM_inserts_L24_underPF, LM_inserts_L25_underPF, LM_inserts_L26_underPF, LM_inserts_L27_underPF, LM_inserts_L2_underPF, LM_inserts_L30_underPF, LM_inserts_L31_underPF, LM_inserts_L33_underPF, LM_inserts_L34_underPF, LM_inserts_L35_underPF, LM_inserts_L36_underPF, LM_inserts_L37_underPF, LM_inserts_L38_underPF, LM_inserts_L39_underPF, LM_inserts_L3_underPF, _
  LM_inserts_L41_underPF, LM_inserts_L42_underPF, LM_inserts_L43_underPF, LM_inserts_L44_underPF, LM_inserts_L49_underPF, LM_inserts_L4_underPF, LM_inserts_L50_underPF, LM_inserts_L51_underPF, LM_inserts_L52_underPF, LM_inserts_L53_underPF, LM_inserts_L54_underPF, LM_inserts_L55_underPF, LM_inserts_L56_underPF, LM_inserts_L5_underPF, LM_inserts_L6_underPF, LM_inserts_L7_underPF, LM_inserts_L9_underPF, LM_flashers_f07_underPF, LM_flashers_f18_underPF, LM_flashers_f19_underPF, LM_flashers_f1a_underPF, LM_flashers_f1b_underPF, LM_flashers_f1c_underPF, LM_flashers_f20a_underPF, LM_flashers_f3_underPF, LM_flashers_f4_underPF, LM_flashers_f5_underPF, LM_flashers_f6_underPF, LM_flashers_f7lb_underPF, LM_flashers_f7lt_underPF, LM_flashers_f7rb_underPF, LM_flashers_f7rt_underPF)
'' Arrays per lighting scenario
'Dim BL_Backbox_light: BL_Backbox_light=Array(BM_BS1, BM_BS2, BM_BS3, BM_BallReleaseGate, BM_Bumper1, BM_Bumper2, BM_Bumper3, BM_Gate1, BM_Gate2, BM_LEMK, BM_LFlipper, BM_LSling1, BM_LSling2, BM_Layer0, BM_Layer1, BM_Playfield, BM_REMK, BM_RFlipper, BM_RSling1, BM_RSling2, BM_lPost, BM_parts, BM_rPost, BM_sw16, BM_sw17, BM_sw18, BM_sw19, BM_sw20, BM_sw21, BM_sw22, BM_sw23, BM_sw24, BM_sw26, BM_sw28, BM_sw29, BM_sw30, BM_sw31, BM_sw32, BM_sw41, BM_sw42, BM_sw43, BM_sw44, BM_sw45, BM_sw46, BM_sw47, BM_sw47_001, BM_sw57, BM_sw58, BM_sw60, BM_sw61, BM_toys, BM_underPF)
'Dim BL_GI: BL_GI=Array(LM_GI_BS1, LM_GI_BS2, LM_GI_BS3, LM_GI_Bumper1, LM_GI_Bumper2, LM_GI_Bumper3, LM_GI_Gate1, LM_GI_Gate2, LM_GI_LEMK, LM_GI_LFlipper, LM_GI_Layer0, LM_GI_Layer1, LM_GI_Playfield, LM_GI_RFlipper, LM_GI_lPost, LM_GI_parts, LM_GI_rPost, LM_GI_sw17, LM_GI_sw18, LM_GI_sw19, LM_GI_sw20, LM_GI_sw21, LM_GI_sw22, LM_GI_sw23, LM_GI_sw24, LM_GI_sw26, LM_GI_sw28, LM_GI_sw29, LM_GI_sw30, LM_GI_sw31, LM_GI_sw32, LM_GI_sw41, LM_GI_sw42, LM_GI_sw43, LM_GI_sw44, LM_GI_sw45, LM_GI_sw46, LM_GI_sw47, LM_GI_sw47_001, LM_GI_sw60, LM_GI_toys, LM_GI_underPF)
'Dim BL_GI_splitted_GI_10: BL_GI_splitted_GI_10=Array(LM_GI_splitted_GI_10_Layer1, LM_GI_splitted_GI_10_Playfield, LM_GI_splitted_GI_10_parts, LM_GI_splitted_GI_10_rPost, LM_GI_splitted_GI_10_sw21, LM_GI_splitted_GI_10_sw22, LM_GI_splitted_GI_10_sw23, LM_GI_splitted_GI_10_underPF)
'Dim BL_GI_splitted_GI_12: BL_GI_splitted_GI_12=Array(LM_GI_splitted_GI_12_Layer1, LM_GI_splitted_GI_12_Playfield, LM_GI_splitted_GI_12_lPost, LM_GI_splitted_GI_12_parts, LM_GI_splitted_GI_12_sw17, LM_GI_splitted_GI_12_sw18, LM_GI_splitted_GI_12_sw19, LM_GI_splitted_GI_12_toys)
'Dim BL_GI_splitted_GI_14: BL_GI_splitted_GI_14=Array(LM_GI_splitted_GI_14_LEMK, LM_GI_splitted_GI_14_LFlipper, LM_GI_splitted_GI_14_LSling1, LM_GI_splitted_GI_14_LSling2, LM_GI_splitted_GI_14_Layer1, LM_GI_splitted_GI_14_Playfield, LM_GI_splitted_GI_14_RFlipper, LM_GI_splitted_GI_14_parts, LM_GI_splitted_GI_14_sw30, LM_GI_splitted_GI_14_sw31, LM_GI_splitted_GI_14_sw57, LM_GI_splitted_GI_14_sw58, LM_GI_splitted_GI_14_underPF)
'Dim BL_GI_splitted_GI_16: BL_GI_splitted_GI_16=Array(LM_GI_splitted_GI_16_LFlipper, LM_GI_splitted_GI_16_Layer1, LM_GI_splitted_GI_16_Playfield, LM_GI_splitted_GI_16_REMK, LM_GI_splitted_GI_16_RFlipper, LM_GI_splitted_GI_16_RSling1, LM_GI_splitted_GI_16_RSling2, LM_GI_splitted_GI_16_parts, LM_GI_splitted_GI_16_sw29, LM_GI_splitted_GI_16_sw30, LM_GI_splitted_GI_16_sw31, LM_GI_splitted_GI_16_sw32, LM_GI_splitted_GI_16_sw60, LM_GI_splitted_GI_16_sw61, LM_GI_splitted_GI_16_underPF)
'Dim BL_GI_splitted_GI_2: BL_GI_splitted_GI_2=Array(LM_GI_splitted_GI_2_LFlipper, LM_GI_splitted_GI_2_Layer1, LM_GI_splitted_GI_2_Playfield, LM_GI_splitted_GI_2_REMK, LM_GI_splitted_GI_2_RFlipper, LM_GI_splitted_GI_2_RSling2, LM_GI_splitted_GI_2_parts, LM_GI_splitted_GI_2_sw24, LM_GI_splitted_GI_2_sw29, LM_GI_splitted_GI_2_sw60, LM_GI_splitted_GI_2_sw61)
'Dim BL_GI_splitted_GI_22: BL_GI_splitted_GI_22=Array(LM_GI_splitted_GI_22_LEMK, LM_GI_splitted_GI_22_LFlipper, LM_GI_splitted_GI_22_LSling1, LM_GI_splitted_GI_22_LSling2, LM_GI_splitted_GI_22_Layer1, LM_GI_splitted_GI_22_Playfield, LM_GI_splitted_GI_22_lPost, LM_GI_splitted_GI_22_parts, LM_GI_splitted_GI_22_sw24, LM_GI_splitted_GI_22_sw29, LM_GI_splitted_GI_22_sw30, LM_GI_splitted_GI_22_sw57, LM_GI_splitted_GI_22_sw58, LM_GI_splitted_GI_22_underPF)
'Dim BL_GI_splitted_GI_24: BL_GI_splitted_GI_24=Array(LM_GI_splitted_GI_24_Layer1, LM_GI_splitted_GI_24_Playfield, LM_GI_splitted_GI_24_REMK, LM_GI_splitted_GI_24_RFlipper, LM_GI_splitted_GI_24_RSling1, LM_GI_splitted_GI_24_RSling2, LM_GI_splitted_GI_24_parts, LM_GI_splitted_GI_24_rPost, LM_GI_splitted_GI_24_sw20, LM_GI_splitted_GI_24_sw60, LM_GI_splitted_GI_24_sw61, LM_GI_splitted_GI_24_underPF)
'Dim BL_GI_splitted_GI_26: BL_GI_splitted_GI_26=Array(LM_GI_splitted_GI_26_Layer1, LM_GI_splitted_GI_26_Playfield, LM_GI_splitted_GI_26_parts, LM_GI_splitted_GI_26_rPost, LM_GI_splitted_GI_26_sw21, LM_GI_splitted_GI_26_sw22, LM_GI_splitted_GI_26_sw23, LM_GI_splitted_GI_26_sw24, LM_GI_splitted_GI_26_underPF)
'Dim BL_GI_splitted_GI_3: BL_GI_splitted_GI_3=Array(LM_GI_splitted_GI_3_LEMK, LM_GI_splitted_GI_3_LFlipper, LM_GI_splitted_GI_3_LSling2, LM_GI_splitted_GI_3_Layer1, LM_GI_splitted_GI_3_Playfield, LM_GI_splitted_GI_3_RFlipper, LM_GI_splitted_GI_3_parts, LM_GI_splitted_GI_3_sw20, LM_GI_splitted_GI_3_sw57, LM_GI_splitted_GI_3_sw58)
'Dim BL_GI_splitted_GI_30: BL_GI_splitted_GI_30=Array(LM_GI_splitted_GI_30_Layer1, LM_GI_splitted_GI_30_Playfield, LM_GI_splitted_GI_30_lPost, LM_GI_splitted_GI_30_parts, LM_GI_splitted_GI_30_sw17, LM_GI_splitted_GI_30_sw18, LM_GI_splitted_GI_30_sw19, LM_GI_splitted_GI_30_sw20, LM_GI_splitted_GI_30_toys, LM_GI_splitted_GI_30_underPF)
'Dim BL_GI_splitted_GI_6: BL_GI_splitted_GI_6=Array(LM_GI_splitted_GI_6_LEMK, LM_GI_splitted_GI_6_LFlipper, LM_GI_splitted_GI_6_LSling2, LM_GI_splitted_GI_6_Layer1, LM_GI_splitted_GI_6_Playfield, LM_GI_splitted_GI_6_RFlipper, LM_GI_splitted_GI_6_parts, LM_GI_splitted_GI_6_sw29, LM_GI_splitted_GI_6_sw30, LM_GI_splitted_GI_6_sw57, LM_GI_splitted_GI_6_sw58)
'Dim BL_GI_splitted_GI_8: BL_GI_splitted_GI_8=Array(LM_GI_splitted_GI_8_LFlipper, LM_GI_splitted_GI_8_Layer1, LM_GI_splitted_GI_8_Playfield, LM_GI_splitted_GI_8_REMK, LM_GI_splitted_GI_8_RFlipper, LM_GI_splitted_GI_8_RSling2, LM_GI_splitted_GI_8_parts, LM_GI_splitted_GI_8_sw29, LM_GI_splitted_GI_8_sw30, LM_GI_splitted_GI_8_sw60, LM_GI_splitted_GI_8_sw61)
Dim BL_flashers_f07: BL_flashers_f07=Array(LM_flashers_f07_BS1, LM_flashers_f07_Layer1, LM_flashers_f07_Playfield, LM_flashers_f07_parts, LM_flashers_f07_sw19, LM_flashers_f07_sw29, LM_flashers_f07_sw45, LM_flashers_f07_sw46, LM_flashers_f07_toys, LM_flashers_f07_underPF)
Dim BL_flashers_f1: BL_flashers_f1=Array(LM_flashers_f1_BS1, LM_flashers_f1_BS2, LM_flashers_f1_Bumper1, LM_flashers_f1_Gate1, LM_flashers_f1_Gate2, LM_flashers_f1_Layer0, LM_flashers_f1_Layer1, LM_flashers_f1_Playfield, LM_flashers_f1_parts, LM_flashers_f1_toys)
Dim BL_flashers_f18: BL_flashers_f18=Array(LM_flashers_f18_LSling2, LM_flashers_f18_Layer1, LM_flashers_f18_Playfield, LM_flashers_f18_REMK, LM_flashers_f18_RSling1, LM_flashers_f18_RSling2, LM_flashers_f18_parts, LM_flashers_f18_sw17, LM_flashers_f18_sw18, LM_flashers_f18_sw19, LM_flashers_f18_sw20, LM_flashers_f18_sw46, LM_flashers_f18_sw58, LM_flashers_f18_toys, LM_flashers_f18_underPF)
Dim BL_flashers_f19: BL_flashers_f19=Array(LM_flashers_f19_BS1, LM_flashers_f19_BS2, LM_flashers_f19_BS3, LM_flashers_f19_Bumper1, LM_flashers_f19_Bumper2, LM_flashers_f19_Layer1, LM_flashers_f19_Playfield, LM_flashers_f19_RSling2, LM_flashers_f19_parts, LM_flashers_f19_rPost, LM_flashers_f19_sw20, LM_flashers_f19_sw21, LM_flashers_f19_sw22, LM_flashers_f19_sw23, LM_flashers_f19_sw24, LM_flashers_f19_sw29, LM_flashers_f19_sw30, LM_flashers_f19_sw31, LM_flashers_f19_sw32, LM_flashers_f19_sw41, LM_flashers_f19_sw61, LM_flashers_f19_toys, LM_flashers_f19_underPF)
Dim BL_flashers_f1a: BL_flashers_f1a=Array(LM_flashers_f1a_BS1, LM_flashers_f1a_Bumper1, LM_flashers_f1a_Bumper2, LM_flashers_f1a_Gate1, LM_flashers_f1a_Layer1, LM_flashers_f1a_parts, LM_flashers_f1a_toys, LM_flashers_f1a_underPF)
Dim BL_flashers_f1b: BL_flashers_f1b=Array(LM_flashers_f1b_BS1, LM_flashers_f1b_BS2, LM_flashers_f1b_Bumper1, LM_flashers_f1b_Bumper2, LM_flashers_f1b_Layer1, LM_flashers_f1b_parts, LM_flashers_f1b_toys, LM_flashers_f1b_underPF)
Dim BL_flashers_f1c: BL_flashers_f1c=Array(LM_flashers_f1c_BS2, LM_flashers_f1c_BS3, LM_flashers_f1c_Bumper2, LM_flashers_f1c_Bumper3, LM_flashers_f1c_Layer1, LM_flashers_f1c_Playfield, LM_flashers_f1c_parts, LM_flashers_f1c_toys, LM_flashers_f1c_underPF)
Dim BL_flashers_f2: BL_flashers_f2=Array(LM_flashers_f2_BS1, LM_flashers_f2_BS2, LM_flashers_f2_Bumper1, LM_flashers_f2_Layer1, LM_flashers_f2_Playfield, LM_flashers_f2_parts, LM_flashers_f2_sw26, LM_flashers_f2_sw45, LM_flashers_f2_toys)
Dim BL_flashers_f20a: BL_flashers_f20a=Array(LM_flashers_f20a_Layer1, LM_flashers_f20a_Playfield, LM_flashers_f20a_parts, LM_flashers_f20a_sw45, LM_flashers_f20a_toys, LM_flashers_f20a_underPF)
Dim BL_flashers_f20b: BL_flashers_f20b=Array(LM_flashers_f20b_Layer1, LM_flashers_f20b_Playfield, LM_flashers_f20b_parts, LM_flashers_f20b_toys)
Dim BL_flashers_f2a: BL_flashers_f2a=Array(LM_flashers_f2a_BS1, LM_flashers_f2a_BS2, LM_flashers_f2a_Bumper1, LM_flashers_f2a_Bumper2, LM_flashers_f2a_Gate2, LM_flashers_f2a_Layer1, LM_flashers_f2a_Playfield, LM_flashers_f2a_parts, LM_flashers_f2a_sw26, LM_flashers_f2a_toys)
Dim BL_flashers_f3: BL_flashers_f3=Array(LM_flashers_f3_BS3, LM_flashers_f3_Layer1, LM_flashers_f3_Playfield, LM_flashers_f3_parts, LM_flashers_f3_sw29, LM_flashers_f3_sw30, LM_flashers_f3_sw31, LM_flashers_f3_sw32, LM_flashers_f3_sw42, LM_flashers_f3_sw46, LM_flashers_f3_toys, LM_flashers_f3_underPF)
Dim BL_flashers_f4: BL_flashers_f4=Array(LM_flashers_f4_BS2, LM_flashers_f4_LSling2, LM_flashers_f4_Layer1, LM_flashers_f4_Playfield, LM_flashers_f4_parts, LM_flashers_f4_sw20, LM_flashers_f4_sw21, LM_flashers_f4_sw22, LM_flashers_f4_sw23, LM_flashers_f4_sw24, LM_flashers_f4_sw29, LM_flashers_f4_sw30, LM_flashers_f4_sw31, LM_flashers_f4_sw32, LM_flashers_f4_toys, LM_flashers_f4_underPF)
Dim BL_flashers_f5: BL_flashers_f5=Array(LM_flashers_f5_Layer1, LM_flashers_f5_Playfield, LM_flashers_f5_parts, LM_flashers_f5_sw20, LM_flashers_f5_sw21, LM_flashers_f5_sw22, LM_flashers_f5_sw23, LM_flashers_f5_sw24, LM_flashers_f5_sw29, LM_flashers_f5_sw30, LM_flashers_f5_sw31, LM_flashers_f5_sw32, LM_flashers_f5_toys, LM_flashers_f5_underPF)
Dim BL_flashers_f6: BL_flashers_f6=Array(LM_flashers_f6_BS2, LM_flashers_f6_BS3, LM_flashers_f6_Layer0, LM_flashers_f6_Layer1, LM_flashers_f6_Playfield, LM_flashers_f6_parts, LM_flashers_f6_sw29, LM_flashers_f6_sw30, LM_flashers_f6_sw31, LM_flashers_f6_sw32, LM_flashers_f6_sw43, LM_flashers_f6_sw46, LM_flashers_f6_toys, LM_flashers_f6_underPF)
Dim BL_flashers_f7lb: BL_flashers_f7lb=Array(LM_flashers_f7lb_LSling2, LM_flashers_f7lb_underPF)
Dim BL_flashers_f7lt: BL_flashers_f7lt=Array(LM_flashers_f7lt_LSling2, LM_flashers_f7lt_Playfield, LM_flashers_f7lt_underPF)
Dim BL_flashers_f7rb: BL_flashers_f7rb=Array(LM_flashers_f7rb_underPF)
Dim BL_flashers_f7rt: BL_flashers_f7rt=Array(LM_flashers_f7rt_underPF)
Dim BL_flashers_f8: BL_flashers_f8=Array(LM_flashers_f8_BS1, LM_flashers_f8_BS2, LM_flashers_f8_BS3, LM_flashers_f8_Bumper2, LM_flashers_f8_Gate2, LM_flashers_f8_Layer0, LM_flashers_f8_Layer1, LM_flashers_f8_Playfield, LM_flashers_f8_parts, LM_flashers_f8_toys)
Dim BL_flashers_f8a: BL_flashers_f8a=Array(LM_flashers_f8a_BS1, LM_flashers_f8a_BS2, LM_flashers_f8a_Gate2, LM_flashers_f8a_Layer0, LM_flashers_f8a_Layer1, LM_flashers_f8a_Playfield, LM_flashers_f8a_parts, LM_flashers_f8a_toys)
'Dim BL_inserts_L1: BL_inserts_L1=Array(LM_inserts_L1_LSling2, LM_inserts_L1_parts, LM_inserts_L1_underPF)
'Dim BL_inserts_L10: BL_inserts_L10=Array(LM_inserts_L10_underPF)
'Dim BL_inserts_L11: BL_inserts_L11=Array(LM_inserts_L11_underPF)
'Dim BL_inserts_L12: BL_inserts_L12=Array(LM_inserts_L12_Layer1, LM_inserts_L12_parts, LM_inserts_L12_underPF)
'Dim BL_inserts_L13: BL_inserts_L13=Array(LM_inserts_L13_Layer1, LM_inserts_L13_parts, LM_inserts_L13_toys, LM_inserts_L13_underPF)
'Dim BL_inserts_L14: BL_inserts_L14=Array(LM_inserts_L14_Layer1, LM_inserts_L14_parts, LM_inserts_L14_toys, LM_inserts_L14_underPF)
'Dim BL_inserts_L15: BL_inserts_L15=Array(LM_inserts_L15_Playfield, LM_inserts_L15_underPF)
'Dim BL_inserts_L16: BL_inserts_L16=Array(LM_inserts_L16_Playfield, LM_inserts_L16_underPF)
'Dim BL_inserts_L17: BL_inserts_L17=Array(LM_inserts_L17_Playfield, LM_inserts_L17_parts, LM_inserts_L17_underPF)
'Dim BL_inserts_L18: BL_inserts_L18=Array(LM_inserts_L18_Layer1, LM_inserts_L18_parts, LM_inserts_L18_toys, LM_inserts_L18_underPF)
'Dim BL_inserts_L19: BL_inserts_L19=Array(LM_inserts_L19_parts, LM_inserts_L19_underPF)
'Dim BL_inserts_L2: BL_inserts_L2=Array(LM_inserts_L2_Playfield, LM_inserts_L2_underPF)
'Dim BL_inserts_L20: BL_inserts_L20=Array(LM_inserts_L20_Layer1, LM_inserts_L20_parts, LM_inserts_L20_toys, LM_inserts_L20_underPF)
'Dim BL_inserts_L21: BL_inserts_L21=Array(LM_inserts_L21_Layer1, LM_inserts_L21_parts, LM_inserts_L21_toys, LM_inserts_L21_underPF)
'Dim BL_inserts_L22: BL_inserts_L22=Array(LM_inserts_L22_Layer1, LM_inserts_L22_parts, LM_inserts_L22_toys, LM_inserts_L22_underPF)
'Dim BL_inserts_L23: BL_inserts_L23=Array(LM_inserts_L23_Layer1, LM_inserts_L23_Playfield, LM_inserts_L23_parts, LM_inserts_L23_sw46, LM_inserts_L23_toys, LM_inserts_L23_underPF)
'Dim BL_inserts_L24: BL_inserts_L24=Array(LM_inserts_L24_Playfield, LM_inserts_L24_parts, LM_inserts_L24_underPF)
'Dim BL_inserts_L25: BL_inserts_L25=Array(LM_inserts_L25_Layer1, LM_inserts_L25_parts, LM_inserts_L25_underPF)
'Dim BL_inserts_L26: BL_inserts_L26=Array(LM_inserts_L26_Layer1, LM_inserts_L26_parts, LM_inserts_L26_sw17, LM_inserts_L26_underPF)
'Dim BL_inserts_L27: BL_inserts_L27=Array(LM_inserts_L27_Layer1, LM_inserts_L27_parts, LM_inserts_L27_sw18, LM_inserts_L27_underPF)
'Dim BL_inserts_L28: BL_inserts_L28=Array(LM_inserts_L28_Playfield, LM_inserts_L28_parts)
'Dim BL_inserts_L29: BL_inserts_L29=Array(LM_inserts_L29_Layer1, LM_inserts_L29_parts, LM_inserts_L29_underPF)
'Dim BL_inserts_L3: BL_inserts_L3=Array(LM_inserts_L3_underPF)
'Dim BL_inserts_L30: BL_inserts_L30=Array(LM_inserts_L30_Layer1, LM_inserts_L30_parts, LM_inserts_L30_sw21, LM_inserts_L30_underPF)
'Dim BL_inserts_L31: BL_inserts_L31=Array(LM_inserts_L31_Layer1, LM_inserts_L31_parts, LM_inserts_L31_sw22, LM_inserts_L31_underPF)
'Dim BL_inserts_L32: BL_inserts_L32=Array(LM_inserts_L32_Playfield, LM_inserts_L32_parts)
'Dim BL_inserts_L33: BL_inserts_L33=Array(LM_inserts_L33_Playfield, LM_inserts_L33_underPF)
'Dim BL_inserts_L34: BL_inserts_L34=Array(LM_inserts_L34_Playfield, LM_inserts_L34_toys, LM_inserts_L34_underPF)
'Dim BL_inserts_L35: BL_inserts_L35=Array(LM_inserts_L35_Playfield, LM_inserts_L35_parts, LM_inserts_L35_toys, LM_inserts_L35_underPF)
'Dim BL_inserts_L36: BL_inserts_L36=Array(LM_inserts_L36_Layer1, LM_inserts_L36_Playfield, LM_inserts_L36_parts, LM_inserts_L36_underPF)
'Dim BL_inserts_L37: BL_inserts_L37=Array(LM_inserts_L37_Layer1, LM_inserts_L37_Playfield, LM_inserts_L37_parts, LM_inserts_L37_underPF)
'Dim BL_inserts_L38: BL_inserts_L38=Array(LM_inserts_L38_Layer0, LM_inserts_L38_Layer1, LM_inserts_L38_Playfield, LM_inserts_L38_parts, LM_inserts_L38_sw43, LM_inserts_L38_underPF)
'Dim BL_inserts_L39: BL_inserts_L39=Array(LM_inserts_L39_Playfield, LM_inserts_L39_underPF)
'Dim BL_inserts_L4: BL_inserts_L4=Array(LM_inserts_L4_underPF)
'Dim BL_inserts_L41: BL_inserts_L41=Array(LM_inserts_L41_Playfield, LM_inserts_L41_sw29, LM_inserts_L41_toys, LM_inserts_L41_underPF)
'Dim BL_inserts_L42: BL_inserts_L42=Array(LM_inserts_L42_Playfield, LM_inserts_L42_toys, LM_inserts_L42_underPF)
'Dim BL_inserts_L43: BL_inserts_L43=Array(LM_inserts_L43_Playfield, LM_inserts_L43_toys, LM_inserts_L43_underPF)
'Dim BL_inserts_L44: BL_inserts_L44=Array(LM_inserts_L44_Playfield, LM_inserts_L44_sw31, LM_inserts_L44_underPF)
'Dim BL_inserts_L45: BL_inserts_L45=Array(LM_inserts_L45_BS1, LM_inserts_L45_BS2, LM_inserts_L45_BS3, LM_inserts_L45_Bumper1, LM_inserts_L45_Bumper2, LM_inserts_L45_Bumper3, LM_inserts_L45_Gate1, LM_inserts_L45_Gate2, LM_inserts_L45_Layer1, LM_inserts_L45_Playfield, LM_inserts_L45_parts, LM_inserts_L45_sw45, LM_inserts_L45_sw46, LM_inserts_L45_toys)
'Dim BL_inserts_L46: BL_inserts_L46=Array(LM_inserts_L46_BS1, LM_inserts_L46_BS2, LM_inserts_L46_BS3, LM_inserts_L46_Bumper1, LM_inserts_L46_Bumper2, LM_inserts_L46_Bumper3, LM_inserts_L46_Gate2, LM_inserts_L46_Layer0, LM_inserts_L46_Layer1, LM_inserts_L46_Playfield, LM_inserts_L46_parts, LM_inserts_L46_toys)
'Dim BL_inserts_L47: BL_inserts_L47=Array(LM_inserts_L47_BS1, LM_inserts_L47_BS2, LM_inserts_L47_BS3, LM_inserts_L47_Bumper1, LM_inserts_L47_Bumper2, LM_inserts_L47_Bumper3, LM_inserts_L47_Gate2, LM_inserts_L47_Layer1, LM_inserts_L47_Playfield, LM_inserts_L47_parts, LM_inserts_L47_sw45, LM_inserts_L47_toys)
'Dim BL_inserts_L49: BL_inserts_L49=Array(LM_inserts_L49_Layer1, LM_inserts_L49_parts, LM_inserts_L49_underPF)
'Dim BL_inserts_L5: BL_inserts_L5=Array(LM_inserts_L5_Playfield, LM_inserts_L5_underPF)
'Dim BL_inserts_L50: BL_inserts_L50=Array(LM_inserts_L50_Layer1, LM_inserts_L50_parts, LM_inserts_L50_underPF)
'Dim BL_inserts_L51: BL_inserts_L51=Array(LM_inserts_L51_Layer1, LM_inserts_L51_parts, LM_inserts_L51_underPF)
'Dim BL_inserts_L52: BL_inserts_L52=Array(LM_inserts_L52_Layer1, LM_inserts_L52_parts, LM_inserts_L52_underPF)
'Dim BL_inserts_L53: BL_inserts_L53=Array(LM_inserts_L53_Playfield, LM_inserts_L53_parts, LM_inserts_L53_underPF)
'Dim BL_inserts_L54: BL_inserts_L54=Array(LM_inserts_L54_Layer1, LM_inserts_L54_Playfield, LM_inserts_L54_toys, LM_inserts_L54_underPF)
'Dim BL_inserts_L55: BL_inserts_L55=Array(LM_inserts_L55_Playfield, LM_inserts_L55_parts, LM_inserts_L55_toys, LM_inserts_L55_underPF)
'Dim BL_inserts_L56: BL_inserts_L56=Array(LM_inserts_L56_Playfield, LM_inserts_L56_toys, LM_inserts_L56_underPF)
'Dim BL_inserts_L6: BL_inserts_L6=Array(LM_inserts_L6_parts, LM_inserts_L6_underPF)
'Dim BL_inserts_L7: BL_inserts_L7=Array(LM_inserts_L7_Playfield, LM_inserts_L7_underPF)
'Dim BL_inserts_L9: BL_inserts_L9=Array(LM_inserts_L9_underPF)
'' Global arrays
'Dim BG_Bakemap: BG_Bakemap=Array(BM_BS1, BM_BS2, BM_BS3, BM_BallReleaseGate, BM_Bumper1, BM_Bumper2, BM_Bumper3, BM_Gate1, BM_Gate2, BM_LEMK, BM_LFlipper, BM_LSling1, BM_LSling2, BM_Layer0, BM_Layer1, BM_Playfield, BM_REMK, BM_RFlipper, BM_RSling1, BM_RSling2, BM_lPost, BM_parts, BM_rPost, BM_sw16, BM_sw17, BM_sw18, BM_sw19, BM_sw20, BM_sw21, BM_sw22, BM_sw23, BM_sw24, BM_sw26, BM_sw28, BM_sw29, BM_sw30, BM_sw31, BM_sw32, BM_sw41, BM_sw42, BM_sw43, BM_sw44, BM_sw45, BM_sw46, BM_sw47, BM_sw47_001, BM_sw57, BM_sw58, BM_sw60, BM_sw61, BM_toys, BM_underPF)
'Dim BG_Lightmap: BG_Lightmap=Array(LM_GI_BS1, LM_GI_BS2, LM_GI_BS3, LM_GI_Bumper1, LM_GI_Bumper2, LM_GI_Bumper3, LM_GI_Gate1, LM_GI_Gate2, LM_GI_LEMK, LM_GI_LFlipper, LM_GI_Layer0, LM_GI_Layer1, LM_GI_Playfield, LM_GI_RFlipper, LM_GI_lPost, LM_GI_parts, LM_GI_rPost, LM_GI_sw17, LM_GI_sw18, LM_GI_sw19, LM_GI_sw20, LM_GI_sw21, LM_GI_sw22, LM_GI_sw23, LM_GI_sw24, LM_GI_sw26, LM_GI_sw28, LM_GI_sw29, LM_GI_sw30, LM_GI_sw31, LM_GI_sw32, LM_GI_sw41, LM_GI_sw42, LM_GI_sw43, LM_GI_sw44, LM_GI_sw45, LM_GI_sw46, LM_GI_sw47, LM_GI_sw47_001, LM_GI_sw60, LM_GI_toys, LM_GI_underPF, LM_GI_splitted_GI_10_Layer1, LM_GI_splitted_GI_10_Playfield, LM_GI_splitted_GI_10_parts, LM_GI_splitted_GI_10_rPost, LM_GI_splitted_GI_10_sw21, LM_GI_splitted_GI_10_sw22, LM_GI_splitted_GI_10_sw23, LM_GI_splitted_GI_10_underPF, LM_GI_splitted_GI_12_Layer1, LM_GI_splitted_GI_12_Playfield, LM_GI_splitted_GI_12_lPost, LM_GI_splitted_GI_12_parts, LM_GI_splitted_GI_12_sw17, LM_GI_splitted_GI_12_sw18, LM_GI_splitted_GI_12_sw19, _
' LM_GI_splitted_GI_12_toys, LM_GI_splitted_GI_14_LEMK, LM_GI_splitted_GI_14_LFlipper, LM_GI_splitted_GI_14_LSling1, LM_GI_splitted_GI_14_LSling2, LM_GI_splitted_GI_14_Layer1, LM_GI_splitted_GI_14_Playfield, LM_GI_splitted_GI_14_RFlipper, LM_GI_splitted_GI_14_parts, LM_GI_splitted_GI_14_sw30, LM_GI_splitted_GI_14_sw31, LM_GI_splitted_GI_14_sw57, LM_GI_splitted_GI_14_sw58, LM_GI_splitted_GI_14_underPF, LM_GI_splitted_GI_16_LFlipper, LM_GI_splitted_GI_16_Layer1, LM_GI_splitted_GI_16_Playfield, LM_GI_splitted_GI_16_REMK, LM_GI_splitted_GI_16_RFlipper, LM_GI_splitted_GI_16_RSling1, LM_GI_splitted_GI_16_RSling2, LM_GI_splitted_GI_16_parts, LM_GI_splitted_GI_16_sw29, LM_GI_splitted_GI_16_sw30, LM_GI_splitted_GI_16_sw31, LM_GI_splitted_GI_16_sw32, LM_GI_splitted_GI_16_sw60, LM_GI_splitted_GI_16_sw61, LM_GI_splitted_GI_16_underPF, LM_GI_splitted_GI_2_LFlipper, LM_GI_splitted_GI_2_Layer1, LM_GI_splitted_GI_2_Playfield, LM_GI_splitted_GI_2_REMK, LM_GI_splitted_GI_2_RFlipper, LM_GI_splitted_GI_2_RSling2, _
' LM_GI_splitted_GI_2_parts, LM_GI_splitted_GI_2_sw24, LM_GI_splitted_GI_2_sw29, LM_GI_splitted_GI_2_sw60, LM_GI_splitted_GI_2_sw61, LM_GI_splitted_GI_22_LEMK, LM_GI_splitted_GI_22_LFlipper, LM_GI_splitted_GI_22_LSling1, LM_GI_splitted_GI_22_LSling2, LM_GI_splitted_GI_22_Layer1, LM_GI_splitted_GI_22_Playfield, LM_GI_splitted_GI_22_lPost, LM_GI_splitted_GI_22_parts, LM_GI_splitted_GI_22_sw24, LM_GI_splitted_GI_22_sw29, LM_GI_splitted_GI_22_sw30, LM_GI_splitted_GI_22_sw57, LM_GI_splitted_GI_22_sw58, LM_GI_splitted_GI_22_underPF, LM_GI_splitted_GI_24_Layer1, LM_GI_splitted_GI_24_Playfield, LM_GI_splitted_GI_24_REMK, LM_GI_splitted_GI_24_RFlipper, LM_GI_splitted_GI_24_RSling1, LM_GI_splitted_GI_24_RSling2, LM_GI_splitted_GI_24_parts, LM_GI_splitted_GI_24_rPost, LM_GI_splitted_GI_24_sw20, LM_GI_splitted_GI_24_sw60, LM_GI_splitted_GI_24_sw61, LM_GI_splitted_GI_24_underPF, LM_GI_splitted_GI_26_Layer1, LM_GI_splitted_GI_26_Playfield, LM_GI_splitted_GI_26_parts, LM_GI_splitted_GI_26_rPost, LM_GI_splitted_GI_26_sw21, _
' LM_GI_splitted_GI_26_sw22, LM_GI_splitted_GI_26_sw23, LM_GI_splitted_GI_26_sw24, LM_GI_splitted_GI_26_underPF, LM_GI_splitted_GI_3_LEMK, LM_GI_splitted_GI_3_LFlipper, LM_GI_splitted_GI_3_LSling2, LM_GI_splitted_GI_3_Layer1, LM_GI_splitted_GI_3_Playfield, LM_GI_splitted_GI_3_RFlipper, LM_GI_splitted_GI_3_parts, LM_GI_splitted_GI_3_sw20, LM_GI_splitted_GI_3_sw57, LM_GI_splitted_GI_3_sw58, LM_GI_splitted_GI_30_Layer1, LM_GI_splitted_GI_30_Playfield, LM_GI_splitted_GI_30_lPost, LM_GI_splitted_GI_30_parts, LM_GI_splitted_GI_30_sw17, LM_GI_splitted_GI_30_sw18, LM_GI_splitted_GI_30_sw19, LM_GI_splitted_GI_30_sw20, LM_GI_splitted_GI_30_toys, LM_GI_splitted_GI_30_underPF, LM_GI_splitted_GI_6_LEMK, LM_GI_splitted_GI_6_LFlipper, LM_GI_splitted_GI_6_LSling2, LM_GI_splitted_GI_6_Layer1, LM_GI_splitted_GI_6_Playfield, LM_GI_splitted_GI_6_RFlipper, LM_GI_splitted_GI_6_parts, LM_GI_splitted_GI_6_sw29, LM_GI_splitted_GI_6_sw30, LM_GI_splitted_GI_6_sw57, LM_GI_splitted_GI_6_sw58, LM_GI_splitted_GI_8_LFlipper, _
' LM_GI_splitted_GI_8_Layer1, LM_GI_splitted_GI_8_Playfield, LM_GI_splitted_GI_8_REMK, LM_GI_splitted_GI_8_RFlipper, LM_GI_splitted_GI_8_RSling2, LM_GI_splitted_GI_8_parts, LM_GI_splitted_GI_8_sw29, LM_GI_splitted_GI_8_sw30, LM_GI_splitted_GI_8_sw60, LM_GI_splitted_GI_8_sw61, LM_flashers_f07_BS1, LM_flashers_f07_Layer1, LM_flashers_f07_Playfield, LM_flashers_f07_parts, LM_flashers_f07_sw19, LM_flashers_f07_sw29, LM_flashers_f07_sw45, LM_flashers_f07_sw46, LM_flashers_f07_toys, LM_flashers_f07_underPF, LM_flashers_f1_BS1, LM_flashers_f1_BS2, LM_flashers_f1_Bumper1, LM_flashers_f1_Gate1, LM_flashers_f1_Gate2, LM_flashers_f1_Layer0, LM_flashers_f1_Layer1, LM_flashers_f1_Playfield, LM_flashers_f1_parts, LM_flashers_f1_toys, LM_flashers_f18_LSling2, LM_flashers_f18_Layer1, LM_flashers_f18_Playfield, LM_flashers_f18_REMK, LM_flashers_f18_RSling1, LM_flashers_f18_RSling2, LM_flashers_f18_parts, LM_flashers_f18_sw17, LM_flashers_f18_sw18, LM_flashers_f18_sw19, LM_flashers_f18_sw20, LM_flashers_f18_sw46, _
' LM_flashers_f18_sw58, LM_flashers_f18_toys, LM_flashers_f18_underPF, LM_flashers_f19_BS1, LM_flashers_f19_BS2, LM_flashers_f19_BS3, LM_flashers_f19_Bumper1, LM_flashers_f19_Bumper2, LM_flashers_f19_Layer1, LM_flashers_f19_Playfield, LM_flashers_f19_RSling2, LM_flashers_f19_parts, LM_flashers_f19_rPost, LM_flashers_f19_sw20, LM_flashers_f19_sw21, LM_flashers_f19_sw22, LM_flashers_f19_sw23, LM_flashers_f19_sw24, LM_flashers_f19_sw29, LM_flashers_f19_sw30, LM_flashers_f19_sw31, LM_flashers_f19_sw32, LM_flashers_f19_sw41, LM_flashers_f19_sw61, LM_flashers_f19_toys, LM_flashers_f19_underPF, LM_flashers_f1a_BS1, LM_flashers_f1a_Bumper1, LM_flashers_f1a_Bumper2, LM_flashers_f1a_Gate1, LM_flashers_f1a_Layer1, LM_flashers_f1a_parts, LM_flashers_f1a_toys, LM_flashers_f1a_underPF, LM_flashers_f1b_BS1, LM_flashers_f1b_BS2, LM_flashers_f1b_Bumper1, LM_flashers_f1b_Bumper2, LM_flashers_f1b_Layer1, LM_flashers_f1b_parts, LM_flashers_f1b_toys, LM_flashers_f1b_underPF, LM_flashers_f1c_BS2, LM_flashers_f1c_BS3, _
' LM_flashers_f1c_Bumper2, LM_flashers_f1c_Bumper3, LM_flashers_f1c_Layer1, LM_flashers_f1c_Playfield, LM_flashers_f1c_parts, LM_flashers_f1c_toys, LM_flashers_f1c_underPF, LM_flashers_f2_BS1, LM_flashers_f2_BS2, LM_flashers_f2_Bumper1, LM_flashers_f2_Layer1, LM_flashers_f2_Playfield, LM_flashers_f2_parts, LM_flashers_f2_sw26, LM_flashers_f2_sw45, LM_flashers_f2_toys, LM_flashers_f20a_Layer1, LM_flashers_f20a_Playfield, LM_flashers_f20a_parts, LM_flashers_f20a_sw45, LM_flashers_f20a_toys, LM_flashers_f20a_underPF, LM_flashers_f20b_Layer1, LM_flashers_f20b_Playfield, LM_flashers_f20b_parts, LM_flashers_f20b_toys, LM_flashers_f2a_BS1, LM_flashers_f2a_BS2, LM_flashers_f2a_Bumper1, LM_flashers_f2a_Bumper2, LM_flashers_f2a_Gate2, LM_flashers_f2a_Layer1, LM_flashers_f2a_Playfield, LM_flashers_f2a_parts, LM_flashers_f2a_sw26, LM_flashers_f2a_toys, LM_flashers_f3_BS3, LM_flashers_f3_Layer1, LM_flashers_f3_Playfield, LM_flashers_f3_parts, LM_flashers_f3_sw29, LM_flashers_f3_sw30, LM_flashers_f3_sw31, _
' LM_flashers_f3_sw32, LM_flashers_f3_sw42, LM_flashers_f3_sw46, LM_flashers_f3_toys, LM_flashers_f3_underPF, LM_flashers_f4_BS2, LM_flashers_f4_LSling2, LM_flashers_f4_Layer1, LM_flashers_f4_Playfield, LM_flashers_f4_parts, LM_flashers_f4_sw20, LM_flashers_f4_sw21, LM_flashers_f4_sw22, LM_flashers_f4_sw23, LM_flashers_f4_sw24, LM_flashers_f4_sw29, LM_flashers_f4_sw30, LM_flashers_f4_sw31, LM_flashers_f4_sw32, LM_flashers_f4_toys, LM_flashers_f4_underPF, LM_flashers_f5_Layer1, LM_flashers_f5_Playfield, LM_flashers_f5_parts, LM_flashers_f5_sw20, LM_flashers_f5_sw21, LM_flashers_f5_sw22, LM_flashers_f5_sw23, LM_flashers_f5_sw24, LM_flashers_f5_sw29, LM_flashers_f5_sw30, LM_flashers_f5_sw31, LM_flashers_f5_sw32, LM_flashers_f5_toys, LM_flashers_f5_underPF, LM_flashers_f6_BS2, LM_flashers_f6_BS3, LM_flashers_f6_Layer0, LM_flashers_f6_Layer1, LM_flashers_f6_Playfield, LM_flashers_f6_parts, LM_flashers_f6_sw29, LM_flashers_f6_sw30, LM_flashers_f6_sw31, LM_flashers_f6_sw32, LM_flashers_f6_sw43, LM_flashers_f6_sw46, _
' LM_flashers_f6_toys, LM_flashers_f6_underPF, LM_flashers_f7lb_LSling2, LM_flashers_f7lb_underPF, LM_flashers_f7lt_LSling2, LM_flashers_f7lt_Playfield, LM_flashers_f7lt_underPF, LM_flashers_f7rb_underPF, LM_flashers_f7rt_underPF, LM_flashers_f8_BS1, LM_flashers_f8_BS2, LM_flashers_f8_BS3, LM_flashers_f8_Bumper2, LM_flashers_f8_Gate2, LM_flashers_f8_Layer0, LM_flashers_f8_Layer1, LM_flashers_f8_Playfield, LM_flashers_f8_parts, LM_flashers_f8_toys, LM_flashers_f8a_BS1, LM_flashers_f8a_BS2, LM_flashers_f8a_Gate2, LM_flashers_f8a_Layer0, LM_flashers_f8a_Layer1, LM_flashers_f8a_Playfield, LM_flashers_f8a_parts, LM_flashers_f8a_toys, LM_inserts_L1_LSling2, LM_inserts_L1_parts, LM_inserts_L1_underPF, LM_inserts_L10_underPF, LM_inserts_L11_underPF, LM_inserts_L12_Layer1, LM_inserts_L12_parts, LM_inserts_L12_underPF, LM_inserts_L13_Layer1, LM_inserts_L13_parts, LM_inserts_L13_toys, LM_inserts_L13_underPF, LM_inserts_L14_Layer1, LM_inserts_L14_parts, LM_inserts_L14_toys, LM_inserts_L14_underPF, LM_inserts_L15_Playfield, _
' LM_inserts_L15_underPF, LM_inserts_L16_Playfield, LM_inserts_L16_underPF, LM_inserts_L17_Playfield, LM_inserts_L17_parts, LM_inserts_L17_underPF, LM_inserts_L18_Layer1, LM_inserts_L18_parts, LM_inserts_L18_toys, LM_inserts_L18_underPF, LM_inserts_L19_parts, LM_inserts_L19_underPF, LM_inserts_L2_Playfield, LM_inserts_L2_underPF, LM_inserts_L20_Layer1, LM_inserts_L20_parts, LM_inserts_L20_toys, LM_inserts_L20_underPF, LM_inserts_L21_Layer1, LM_inserts_L21_parts, LM_inserts_L21_toys, LM_inserts_L21_underPF, LM_inserts_L22_Layer1, LM_inserts_L22_parts, LM_inserts_L22_toys, LM_inserts_L22_underPF, LM_inserts_L23_Layer1, LM_inserts_L23_Playfield, LM_inserts_L23_parts, LM_inserts_L23_sw46, LM_inserts_L23_toys, LM_inserts_L23_underPF, LM_inserts_L24_Playfield, LM_inserts_L24_parts, LM_inserts_L24_underPF, LM_inserts_L25_Layer1, LM_inserts_L25_parts, LM_inserts_L25_underPF, LM_inserts_L26_Layer1, LM_inserts_L26_parts, LM_inserts_L26_sw17, LM_inserts_L26_underPF, LM_inserts_L27_Layer1, LM_inserts_L27_parts, _
' LM_inserts_L27_sw18, LM_inserts_L27_underPF, LM_inserts_L28_Playfield, LM_inserts_L28_parts, LM_inserts_L29_Layer1, LM_inserts_L29_parts, LM_inserts_L29_underPF, LM_inserts_L3_underPF, LM_inserts_L30_Layer1, LM_inserts_L30_parts, LM_inserts_L30_sw21, LM_inserts_L30_underPF, LM_inserts_L31_Layer1, LM_inserts_L31_parts, LM_inserts_L31_sw22, LM_inserts_L31_underPF, LM_inserts_L32_Playfield, LM_inserts_L32_parts, LM_inserts_L33_Playfield, LM_inserts_L33_underPF, LM_inserts_L34_Playfield, LM_inserts_L34_toys, LM_inserts_L34_underPF, LM_inserts_L35_Playfield, LM_inserts_L35_parts, LM_inserts_L35_toys, LM_inserts_L35_underPF, LM_inserts_L36_Layer1, LM_inserts_L36_Playfield, LM_inserts_L36_parts, LM_inserts_L36_underPF, LM_inserts_L37_Layer1, LM_inserts_L37_Playfield, LM_inserts_L37_parts, LM_inserts_L37_underPF, LM_inserts_L38_Layer0, LM_inserts_L38_Layer1, LM_inserts_L38_Playfield, LM_inserts_L38_parts, LM_inserts_L38_sw43, LM_inserts_L38_underPF, LM_inserts_L39_Playfield, LM_inserts_L39_underPF, _
' LM_inserts_L4_underPF, LM_inserts_L41_Playfield, LM_inserts_L41_sw29, LM_inserts_L41_toys, LM_inserts_L41_underPF, LM_inserts_L42_Playfield, LM_inserts_L42_toys, LM_inserts_L42_underPF, LM_inserts_L43_Playfield, LM_inserts_L43_toys, LM_inserts_L43_underPF, LM_inserts_L44_Playfield, LM_inserts_L44_sw31, LM_inserts_L44_underPF, LM_inserts_L45_BS1, LM_inserts_L45_BS2, LM_inserts_L45_BS3, LM_inserts_L45_Bumper1, LM_inserts_L45_Bumper2, LM_inserts_L45_Bumper3, LM_inserts_L45_Gate1, LM_inserts_L45_Gate2, LM_inserts_L45_Layer1, LM_inserts_L45_Playfield, LM_inserts_L45_parts, LM_inserts_L45_sw45, LM_inserts_L45_sw46, LM_inserts_L45_toys, LM_inserts_L46_BS1, LM_inserts_L46_BS2, LM_inserts_L46_BS3, LM_inserts_L46_Bumper1, LM_inserts_L46_Bumper2, LM_inserts_L46_Bumper3, LM_inserts_L46_Gate2, LM_inserts_L46_Layer0, LM_inserts_L46_Layer1, LM_inserts_L46_Playfield, LM_inserts_L46_parts, LM_inserts_L46_toys, LM_inserts_L47_BS1, LM_inserts_L47_BS2, LM_inserts_L47_BS3, LM_inserts_L47_Bumper1, LM_inserts_L47_Bumper2, _
' LM_inserts_L47_Bumper3, LM_inserts_L47_Gate2, LM_inserts_L47_Layer1, LM_inserts_L47_Playfield, LM_inserts_L47_parts, LM_inserts_L47_sw45, LM_inserts_L47_toys, LM_inserts_L49_Layer1, LM_inserts_L49_parts, LM_inserts_L49_underPF, LM_inserts_L5_Playfield, LM_inserts_L5_underPF, LM_inserts_L50_Layer1, LM_inserts_L50_parts, LM_inserts_L50_underPF, LM_inserts_L51_Layer1, LM_inserts_L51_parts, LM_inserts_L51_underPF, LM_inserts_L52_Layer1, LM_inserts_L52_parts, LM_inserts_L52_underPF, LM_inserts_L53_Playfield, LM_inserts_L53_parts, LM_inserts_L53_underPF, LM_inserts_L54_Layer1, LM_inserts_L54_Playfield, LM_inserts_L54_toys, LM_inserts_L54_underPF, LM_inserts_L55_Playfield, LM_inserts_L55_parts, LM_inserts_L55_toys, LM_inserts_L55_underPF, LM_inserts_L56_Playfield, LM_inserts_L56_toys, LM_inserts_L56_underPF, LM_inserts_L6_parts, LM_inserts_L6_underPF, LM_inserts_L7_Playfield, LM_inserts_L7_underPF, LM_inserts_L9_underPF)
'Dim BG_All: BG_All=Array(BM_BS1, BM_BS2, BM_BS3, BM_BallReleaseGate, BM_Bumper1, BM_Bumper2, BM_Bumper3, BM_Gate1, BM_Gate2, BM_LEMK, BM_LFlipper, BM_LSling1, BM_LSling2, BM_Layer0, BM_Layer1, BM_Playfield, BM_REMK, BM_RFlipper, BM_RSling1, BM_RSling2, BM_lPost, BM_parts, BM_rPost, BM_sw16, BM_sw17, BM_sw18, BM_sw19, BM_sw20, BM_sw21, BM_sw22, BM_sw23, BM_sw24, BM_sw26, BM_sw28, BM_sw29, BM_sw30, BM_sw31, BM_sw32, BM_sw41, BM_sw42, BM_sw43, BM_sw44, BM_sw45, BM_sw46, BM_sw47, BM_sw47_001, BM_sw57, BM_sw58, BM_sw60, BM_sw61, BM_toys, BM_underPF, LM_GI_BS1, LM_GI_BS2, LM_GI_BS3, LM_GI_Bumper1, LM_GI_Bumper2, LM_GI_Bumper3, LM_GI_Gate1, LM_GI_Gate2, LM_GI_LEMK, LM_GI_LFlipper, LM_GI_Layer0, LM_GI_Layer1, LM_GI_Playfield, LM_GI_RFlipper, LM_GI_lPost, LM_GI_parts, LM_GI_rPost, LM_GI_sw17, LM_GI_sw18, LM_GI_sw19, LM_GI_sw20, LM_GI_sw21, LM_GI_sw22, LM_GI_sw23, LM_GI_sw24, LM_GI_sw26, LM_GI_sw28, LM_GI_sw29, LM_GI_sw30, LM_GI_sw31, LM_GI_sw32, LM_GI_sw41, LM_GI_sw42, LM_GI_sw43, LM_GI_sw44, LM_GI_sw45, LM_GI_sw46, _
' LM_GI_sw47, LM_GI_sw47_001, LM_GI_sw60, LM_GI_toys, LM_GI_underPF, LM_GI_splitted_GI_10_Layer1, LM_GI_splitted_GI_10_Playfield, LM_GI_splitted_GI_10_parts, LM_GI_splitted_GI_10_rPost, LM_GI_splitted_GI_10_sw21, LM_GI_splitted_GI_10_sw22, LM_GI_splitted_GI_10_sw23, LM_GI_splitted_GI_10_underPF, LM_GI_splitted_GI_12_Layer1, LM_GI_splitted_GI_12_Playfield, LM_GI_splitted_GI_12_lPost, LM_GI_splitted_GI_12_parts, LM_GI_splitted_GI_12_sw17, LM_GI_splitted_GI_12_sw18, LM_GI_splitted_GI_12_sw19, LM_GI_splitted_GI_12_toys, LM_GI_splitted_GI_14_LEMK, LM_GI_splitted_GI_14_LFlipper, LM_GI_splitted_GI_14_LSling1, LM_GI_splitted_GI_14_LSling2, LM_GI_splitted_GI_14_Layer1, LM_GI_splitted_GI_14_Playfield, LM_GI_splitted_GI_14_RFlipper, LM_GI_splitted_GI_14_parts, LM_GI_splitted_GI_14_sw30, LM_GI_splitted_GI_14_sw31, LM_GI_splitted_GI_14_sw57, LM_GI_splitted_GI_14_sw58, LM_GI_splitted_GI_14_underPF, LM_GI_splitted_GI_16_LFlipper, LM_GI_splitted_GI_16_Layer1, LM_GI_splitted_GI_16_Playfield, LM_GI_splitted_GI_16_REMK, _
' LM_GI_splitted_GI_16_RFlipper, LM_GI_splitted_GI_16_RSling1, LM_GI_splitted_GI_16_RSling2, LM_GI_splitted_GI_16_parts, LM_GI_splitted_GI_16_sw29, LM_GI_splitted_GI_16_sw30, LM_GI_splitted_GI_16_sw31, LM_GI_splitted_GI_16_sw32, LM_GI_splitted_GI_16_sw60, LM_GI_splitted_GI_16_sw61, LM_GI_splitted_GI_16_underPF, LM_GI_splitted_GI_2_LFlipper, LM_GI_splitted_GI_2_Layer1, LM_GI_splitted_GI_2_Playfield, LM_GI_splitted_GI_2_REMK, LM_GI_splitted_GI_2_RFlipper, LM_GI_splitted_GI_2_RSling2, LM_GI_splitted_GI_2_parts, LM_GI_splitted_GI_2_sw24, LM_GI_splitted_GI_2_sw29, LM_GI_splitted_GI_2_sw60, LM_GI_splitted_GI_2_sw61, LM_GI_splitted_GI_22_LEMK, LM_GI_splitted_GI_22_LFlipper, LM_GI_splitted_GI_22_LSling1, LM_GI_splitted_GI_22_LSling2, LM_GI_splitted_GI_22_Layer1, LM_GI_splitted_GI_22_Playfield, LM_GI_splitted_GI_22_lPost, LM_GI_splitted_GI_22_parts, LM_GI_splitted_GI_22_sw24, LM_GI_splitted_GI_22_sw29, LM_GI_splitted_GI_22_sw30, LM_GI_splitted_GI_22_sw57, LM_GI_splitted_GI_22_sw58, LM_GI_splitted_GI_22_underPF, _
' LM_GI_splitted_GI_24_Layer1, LM_GI_splitted_GI_24_Playfield, LM_GI_splitted_GI_24_REMK, LM_GI_splitted_GI_24_RFlipper, LM_GI_splitted_GI_24_RSling1, LM_GI_splitted_GI_24_RSling2, LM_GI_splitted_GI_24_parts, LM_GI_splitted_GI_24_rPost, LM_GI_splitted_GI_24_sw20, LM_GI_splitted_GI_24_sw60, LM_GI_splitted_GI_24_sw61, LM_GI_splitted_GI_24_underPF, LM_GI_splitted_GI_26_Layer1, LM_GI_splitted_GI_26_Playfield, LM_GI_splitted_GI_26_parts, LM_GI_splitted_GI_26_rPost, LM_GI_splitted_GI_26_sw21, LM_GI_splitted_GI_26_sw22, LM_GI_splitted_GI_26_sw23, LM_GI_splitted_GI_26_sw24, LM_GI_splitted_GI_26_underPF, LM_GI_splitted_GI_3_LEMK, LM_GI_splitted_GI_3_LFlipper, LM_GI_splitted_GI_3_LSling2, LM_GI_splitted_GI_3_Layer1, LM_GI_splitted_GI_3_Playfield, LM_GI_splitted_GI_3_RFlipper, LM_GI_splitted_GI_3_parts, LM_GI_splitted_GI_3_sw20, LM_GI_splitted_GI_3_sw57, LM_GI_splitted_GI_3_sw58, LM_GI_splitted_GI_30_Layer1, LM_GI_splitted_GI_30_Playfield, LM_GI_splitted_GI_30_lPost, LM_GI_splitted_GI_30_parts, LM_GI_splitted_GI_30_sw17, _
' LM_GI_splitted_GI_30_sw18, LM_GI_splitted_GI_30_sw19, LM_GI_splitted_GI_30_sw20, LM_GI_splitted_GI_30_toys, LM_GI_splitted_GI_30_underPF, LM_GI_splitted_GI_6_LEMK, LM_GI_splitted_GI_6_LFlipper, LM_GI_splitted_GI_6_LSling2, LM_GI_splitted_GI_6_Layer1, LM_GI_splitted_GI_6_Playfield, LM_GI_splitted_GI_6_RFlipper, LM_GI_splitted_GI_6_parts, LM_GI_splitted_GI_6_sw29, LM_GI_splitted_GI_6_sw30, LM_GI_splitted_GI_6_sw57, LM_GI_splitted_GI_6_sw58, LM_GI_splitted_GI_8_LFlipper, LM_GI_splitted_GI_8_Layer1, LM_GI_splitted_GI_8_Playfield, LM_GI_splitted_GI_8_REMK, LM_GI_splitted_GI_8_RFlipper, LM_GI_splitted_GI_8_RSling2, LM_GI_splitted_GI_8_parts, LM_GI_splitted_GI_8_sw29, LM_GI_splitted_GI_8_sw30, LM_GI_splitted_GI_8_sw60, LM_GI_splitted_GI_8_sw61, LM_flashers_f07_BS1, LM_flashers_f07_Layer1, LM_flashers_f07_Playfield, LM_flashers_f07_parts, LM_flashers_f07_sw19, LM_flashers_f07_sw29, LM_flashers_f07_sw45, LM_flashers_f07_sw46, LM_flashers_f07_toys, LM_flashers_f07_underPF, LM_flashers_f1_BS1, LM_flashers_f1_BS2, _
' LM_flashers_f1_Bumper1, LM_flashers_f1_Gate1, LM_flashers_f1_Gate2, LM_flashers_f1_Layer0, LM_flashers_f1_Layer1, LM_flashers_f1_Playfield, LM_flashers_f1_parts, LM_flashers_f1_toys, LM_flashers_f18_LSling2, LM_flashers_f18_Layer1, LM_flashers_f18_Playfield, LM_flashers_f18_REMK, LM_flashers_f18_RSling1, LM_flashers_f18_RSling2, LM_flashers_f18_parts, LM_flashers_f18_sw17, LM_flashers_f18_sw18, LM_flashers_f18_sw19, LM_flashers_f18_sw20, LM_flashers_f18_sw46, LM_flashers_f18_sw58, LM_flashers_f18_toys, LM_flashers_f18_underPF, LM_flashers_f19_BS1, LM_flashers_f19_BS2, LM_flashers_f19_BS3, LM_flashers_f19_Bumper1, LM_flashers_f19_Bumper2, LM_flashers_f19_Layer1, LM_flashers_f19_Playfield, LM_flashers_f19_RSling2, LM_flashers_f19_parts, LM_flashers_f19_rPost, LM_flashers_f19_sw20, LM_flashers_f19_sw21, LM_flashers_f19_sw22, LM_flashers_f19_sw23, LM_flashers_f19_sw24, LM_flashers_f19_sw29, LM_flashers_f19_sw30, LM_flashers_f19_sw31, LM_flashers_f19_sw32, LM_flashers_f19_sw41, LM_flashers_f19_sw61, _
' LM_flashers_f19_toys, LM_flashers_f19_underPF, LM_flashers_f1a_BS1, LM_flashers_f1a_Bumper1, LM_flashers_f1a_Bumper2, LM_flashers_f1a_Gate1, LM_flashers_f1a_Layer1, LM_flashers_f1a_parts, LM_flashers_f1a_toys, LM_flashers_f1a_underPF, LM_flashers_f1b_BS1, LM_flashers_f1b_BS2, LM_flashers_f1b_Bumper1, LM_flashers_f1b_Bumper2, LM_flashers_f1b_Layer1, LM_flashers_f1b_parts, LM_flashers_f1b_toys, LM_flashers_f1b_underPF, LM_flashers_f1c_BS2, LM_flashers_f1c_BS3, LM_flashers_f1c_Bumper2, LM_flashers_f1c_Bumper3, LM_flashers_f1c_Layer1, LM_flashers_f1c_Playfield, LM_flashers_f1c_parts, LM_flashers_f1c_toys, LM_flashers_f1c_underPF, LM_flashers_f2_BS1, LM_flashers_f2_BS2, LM_flashers_f2_Bumper1, LM_flashers_f2_Layer1, LM_flashers_f2_Playfield, LM_flashers_f2_parts, LM_flashers_f2_sw26, LM_flashers_f2_sw45, LM_flashers_f2_toys, LM_flashers_f20a_Layer1, LM_flashers_f20a_Playfield, LM_flashers_f20a_parts, LM_flashers_f20a_sw45, LM_flashers_f20a_toys, LM_flashers_f20a_underPF, LM_flashers_f20b_Layer1, _
' LM_flashers_f20b_Playfield, LM_flashers_f20b_parts, LM_flashers_f20b_toys, LM_flashers_f2a_BS1, LM_flashers_f2a_BS2, LM_flashers_f2a_Bumper1, LM_flashers_f2a_Bumper2, LM_flashers_f2a_Gate2, LM_flashers_f2a_Layer1, LM_flashers_f2a_Playfield, LM_flashers_f2a_parts, LM_flashers_f2a_sw26, LM_flashers_f2a_toys, LM_flashers_f3_BS3, LM_flashers_f3_Layer1, LM_flashers_f3_Playfield, LM_flashers_f3_parts, LM_flashers_f3_sw29, LM_flashers_f3_sw30, LM_flashers_f3_sw31, LM_flashers_f3_sw32, LM_flashers_f3_sw42, LM_flashers_f3_sw46, LM_flashers_f3_toys, LM_flashers_f3_underPF, LM_flashers_f4_BS2, LM_flashers_f4_LSling2, LM_flashers_f4_Layer1, LM_flashers_f4_Playfield, LM_flashers_f4_parts, LM_flashers_f4_sw20, LM_flashers_f4_sw21, LM_flashers_f4_sw22, LM_flashers_f4_sw23, LM_flashers_f4_sw24, LM_flashers_f4_sw29, LM_flashers_f4_sw30, LM_flashers_f4_sw31, LM_flashers_f4_sw32, LM_flashers_f4_toys, LM_flashers_f4_underPF, LM_flashers_f5_Layer1, LM_flashers_f5_Playfield, LM_flashers_f5_parts, LM_flashers_f5_sw20, _
' LM_flashers_f5_sw21, LM_flashers_f5_sw22, LM_flashers_f5_sw23, LM_flashers_f5_sw24, LM_flashers_f5_sw29, LM_flashers_f5_sw30, LM_flashers_f5_sw31, LM_flashers_f5_sw32, LM_flashers_f5_toys, LM_flashers_f5_underPF, LM_flashers_f6_BS2, LM_flashers_f6_BS3, LM_flashers_f6_Layer0, LM_flashers_f6_Layer1, LM_flashers_f6_Playfield, LM_flashers_f6_parts, LM_flashers_f6_sw29, LM_flashers_f6_sw30, LM_flashers_f6_sw31, LM_flashers_f6_sw32, LM_flashers_f6_sw43, LM_flashers_f6_sw46, LM_flashers_f6_toys, LM_flashers_f6_underPF, LM_flashers_f7lb_LSling2, LM_flashers_f7lb_underPF, LM_flashers_f7lt_LSling2, LM_flashers_f7lt_Playfield, LM_flashers_f7lt_underPF, LM_flashers_f7rb_underPF, LM_flashers_f7rt_underPF, LM_flashers_f8_BS1, LM_flashers_f8_BS2, LM_flashers_f8_BS3, LM_flashers_f8_Bumper2, LM_flashers_f8_Gate2, LM_flashers_f8_Layer0, LM_flashers_f8_Layer1, LM_flashers_f8_Playfield, LM_flashers_f8_parts, LM_flashers_f8_toys, LM_flashers_f8a_BS1, LM_flashers_f8a_BS2, LM_flashers_f8a_Gate2, LM_flashers_f8a_Layer0, _
' LM_flashers_f8a_Layer1, LM_flashers_f8a_Playfield, LM_flashers_f8a_parts, LM_flashers_f8a_toys, LM_inserts_L1_LSling2, LM_inserts_L1_parts, LM_inserts_L1_underPF, LM_inserts_L10_underPF, LM_inserts_L11_underPF, LM_inserts_L12_Layer1, LM_inserts_L12_parts, LM_inserts_L12_underPF, LM_inserts_L13_Layer1, LM_inserts_L13_parts, LM_inserts_L13_toys, LM_inserts_L13_underPF, LM_inserts_L14_Layer1, LM_inserts_L14_parts, LM_inserts_L14_toys, LM_inserts_L14_underPF, LM_inserts_L15_Playfield, LM_inserts_L15_underPF, LM_inserts_L16_Playfield, LM_inserts_L16_underPF, LM_inserts_L17_Playfield, LM_inserts_L17_parts, LM_inserts_L17_underPF, LM_inserts_L18_Layer1, LM_inserts_L18_parts, LM_inserts_L18_toys, LM_inserts_L18_underPF, LM_inserts_L19_parts, LM_inserts_L19_underPF, LM_inserts_L2_Playfield, LM_inserts_L2_underPF, LM_inserts_L20_Layer1, LM_inserts_L20_parts, LM_inserts_L20_toys, LM_inserts_L20_underPF, LM_inserts_L21_Layer1, LM_inserts_L21_parts, LM_inserts_L21_toys, LM_inserts_L21_underPF, LM_inserts_L22_Layer1, _
' LM_inserts_L22_parts, LM_inserts_L22_toys, LM_inserts_L22_underPF, LM_inserts_L23_Layer1, LM_inserts_L23_Playfield, LM_inserts_L23_parts, LM_inserts_L23_sw46, LM_inserts_L23_toys, LM_inserts_L23_underPF, LM_inserts_L24_Playfield, LM_inserts_L24_parts, LM_inserts_L24_underPF, LM_inserts_L25_Layer1, LM_inserts_L25_parts, LM_inserts_L25_underPF, LM_inserts_L26_Layer1, LM_inserts_L26_parts, LM_inserts_L26_sw17, LM_inserts_L26_underPF, LM_inserts_L27_Layer1, LM_inserts_L27_parts, LM_inserts_L27_sw18, LM_inserts_L27_underPF, LM_inserts_L28_Playfield, LM_inserts_L28_parts, LM_inserts_L29_Layer1, LM_inserts_L29_parts, LM_inserts_L29_underPF, LM_inserts_L3_underPF, LM_inserts_L30_Layer1, LM_inserts_L30_parts, LM_inserts_L30_sw21, LM_inserts_L30_underPF, LM_inserts_L31_Layer1, LM_inserts_L31_parts, LM_inserts_L31_sw22, LM_inserts_L31_underPF, LM_inserts_L32_Playfield, LM_inserts_L32_parts, LM_inserts_L33_Playfield, LM_inserts_L33_underPF, LM_inserts_L34_Playfield, LM_inserts_L34_toys, LM_inserts_L34_underPF, _
' LM_inserts_L35_Playfield, LM_inserts_L35_parts, LM_inserts_L35_toys, LM_inserts_L35_underPF, LM_inserts_L36_Layer1, LM_inserts_L36_Playfield, LM_inserts_L36_parts, LM_inserts_L36_underPF, LM_inserts_L37_Layer1, LM_inserts_L37_Playfield, LM_inserts_L37_parts, LM_inserts_L37_underPF, LM_inserts_L38_Layer0, LM_inserts_L38_Layer1, LM_inserts_L38_Playfield, LM_inserts_L38_parts, LM_inserts_L38_sw43, LM_inserts_L38_underPF, LM_inserts_L39_Playfield, LM_inserts_L39_underPF, LM_inserts_L4_underPF, LM_inserts_L41_Playfield, LM_inserts_L41_sw29, LM_inserts_L41_toys, LM_inserts_L41_underPF, LM_inserts_L42_Playfield, LM_inserts_L42_toys, LM_inserts_L42_underPF, LM_inserts_L43_Playfield, LM_inserts_L43_toys, LM_inserts_L43_underPF, LM_inserts_L44_Playfield, LM_inserts_L44_sw31, LM_inserts_L44_underPF, LM_inserts_L45_BS1, LM_inserts_L45_BS2, LM_inserts_L45_BS3, LM_inserts_L45_Bumper1, LM_inserts_L45_Bumper2, LM_inserts_L45_Bumper3, LM_inserts_L45_Gate1, LM_inserts_L45_Gate2, LM_inserts_L45_Layer1, LM_inserts_L45_Playfield, _
' LM_inserts_L45_parts, LM_inserts_L45_sw45, LM_inserts_L45_sw46, LM_inserts_L45_toys, LM_inserts_L46_BS1, LM_inserts_L46_BS2, LM_inserts_L46_BS3, LM_inserts_L46_Bumper1, LM_inserts_L46_Bumper2, LM_inserts_L46_Bumper3, LM_inserts_L46_Gate2, LM_inserts_L46_Layer0, LM_inserts_L46_Layer1, LM_inserts_L46_Playfield, LM_inserts_L46_parts, LM_inserts_L46_toys, LM_inserts_L47_BS1, LM_inserts_L47_BS2, LM_inserts_L47_BS3, LM_inserts_L47_Bumper1, LM_inserts_L47_Bumper2, LM_inserts_L47_Bumper3, LM_inserts_L47_Gate2, LM_inserts_L47_Layer1, LM_inserts_L47_Playfield, LM_inserts_L47_parts, LM_inserts_L47_sw45, LM_inserts_L47_toys, LM_inserts_L49_Layer1, LM_inserts_L49_parts, LM_inserts_L49_underPF, LM_inserts_L5_Playfield, LM_inserts_L5_underPF, LM_inserts_L50_Layer1, LM_inserts_L50_parts, LM_inserts_L50_underPF, LM_inserts_L51_Layer1, LM_inserts_L51_parts, LM_inserts_L51_underPF, LM_inserts_L52_Layer1, LM_inserts_L52_parts, LM_inserts_L52_underPF, LM_inserts_L53_Playfield, LM_inserts_L53_parts, LM_inserts_L53_underPF, _
' LM_inserts_L54_Layer1, LM_inserts_L54_Playfield, LM_inserts_L54_toys, LM_inserts_L54_underPF, LM_inserts_L55_Playfield, LM_inserts_L55_parts, LM_inserts_L55_toys, LM_inserts_L55_underPF, LM_inserts_L56_Playfield, LM_inserts_L56_toys, LM_inserts_L56_underPF, LM_inserts_L6_parts, LM_inserts_L6_underPF, LM_inserts_L7_Playfield, LM_inserts_L7_underPF, LM_inserts_L9_underPF)
'' VLM Arrays - End



'*************************************************************
' VR Room Auto-Detect
'*************************************************************
Dim VR_Obj, UseVPMDMD

Dim DesktopMode: DesktopMode = Table1.ShowDT

If RenderingMode = 2 Then
  Pincab_Bottom.Visible = 1
  For Each VR_Obj in VRCabinet : VR_Obj.Visible = 1 : Next
  For Each VR_Obj in VRThings : VR_Obj.Visible = 1 : Next
  UseVPMDMD = True
Else
  Pincab_Bottom.Visible = 0
  For Each VR_Obj in VRCabinet : VR_Obj.Visible = 0 : Next
  For Each VR_Obj in VRThings : VR_Obj.Visible = 0 : Next
  UseVPMDMD = DesktopMode
End If

LoadVPM "03060000","SEGA.VBS",3.57


'**************
' Timers
'**************


' FIXME for the time being, the cor timer interval must be 10 ms (so below 60FPS framerate)
CorTimer.Interval = 10
Sub CorTimer_Timer(): Cor.Update: End Sub


Sub FrameTimer_Timer() 'The frame timer interval should be -1, so executes at the display frame rate
  UpdateBallBrightness
  BSUpdate
  AnimateBumperSkirts
  RollingUpdate   ' Update rolling sounds
  DoSTAnim
  UpdateStandupTargets
End Sub


'**************
' Table Init
'**************

Dim Mag1, Mag2, Mag3, Mag4
Dim GZBall1, GZBall2, GZBall3, GZBall4, GZCaptiveBall1, GZCaptiveBall2, GZCaptiveBall3, gBOT

Sub Table1_Init
  vpmInit Me
  On Error Resume Next
    With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine = "Godzilla, Sega 1998"&chr(13)&"VPW"
    .HandleMechanics=0
    .HandleKeyboard=0
    .ShowDMDOnly=1
    .ShowFrame=0
    .ShowTitle=0
    .hidden = 0
        '.Games(cGameName).Settings.Value("sound")=1
    '.PuPHide = 1
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With
     On Error Goto 0

  'Nudging
  vpmNudge.TiltSwitch = 56
  vpmNudge.Sensitivity = 4
  vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingshot)

  PinMAMETimer.Interval = PinMAMEInterval
  PinMAMETimer.Enabled = 1

  vpmMapLights AllLamps ' Map all lamps to the corresponding ROM output using the value of TimerInterval of each light object

  'Trough
  Set GZBall1 = sw14.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set GZBall2 = sw13.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set GZBall3 = sw12.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set GZBall4 = sw11.CreateSizedballWithMass(Ballsize/2,Ballmass)

  Controller.Switch(11) = 1
  Controller.Switch(12) = 1
  Controller.Switch(13) = 1
  Controller.Switch(14) = 1

  'Captive Ball
  Set GZCaptiveBall1 = captiveBall1.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set GZCaptiveBall2 = captiveBall2.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set GZCaptiveBall3 = captiveBall3.CreateSizedballWithMass(Ballsize/2,Ballmass)

  gBOT = Array(GZCaptiveBall1,GZCaptiveBall2,GZCaptiveBall3,GZBall1,GZBall2,GZBall3,GZBall4)

  vpmTimer.AddTimer 300, "captiveball1.kick 180,1 '"
  vpmTimer.AddTimer 310, "captiveball1.enabled= 0 '"
  vpmTimer.AddTimer 300, "captiveball2.kick 180,1 '"
  vpmTimer.AddTimer 310, "captiveball2.enabled= 0 '"
  vpmTimer.AddTimer 300, "captiveball3.kick 180,1 '"
  vpmTimer.AddTimer 310, "captiveball3.enabled= 0 '"

  'Initialize slings
  RStep = 0:RightSlingShot.Timerenabled=True
  LStep = 0:LeftSlingShot.Timerenabled=True

  'Magnets
  Set mag1= New cvpmMagnet
  With mag1
    .InitMagnet MagTop, 16
    .GrabCenter = True
    .solenoid=3
    .CreateEvents "mag1"
  End With

  Set mag2= New cvpmMagnet
  With mag2
    .InitMagnet MagMid, 16
    .GrabCenter = False
    .solenoid=4
    .CreateEvents "mag2"
  End With

  Set mag3= New cvpmMagnet
  With mag3
    .InitMagnet MagBot, 16
    .GrabCenter = False
    .solenoid=5
    .CreateEvents "mag3"
  End With

  Set mag4= New cvpmMagnet
  With mag4
    .InitMagnet MagLeft, 16
    .GrabCenter = True
    .solenoid=14
    .CreateEvents "mag4"
  End With

  Diverter1.IsDropped = 0
  Diverter2.IsDropped = 1
End Sub


Sub Table1_Paused:Controller.Pause = 1:End Sub
Sub Table1_unPaused:Controller.Pause = 0:End Sub
Sub Table1_Exit:Controller.Stop:End Sub




'*******************************************
'  ZOPT: User Options
'*******************************************

Dim RefractOpt : RefractOpt = 1         ' 0 - No Refraction (best performance), 1 - Sharp Refractions (improved performance), 2 - Rough Refractions (best visual)
Dim LightLevel : LightLevel = 0.5       ' Level of room lighting (0 to 1), where 0 is dark and 100 is brightest
Dim ColorLUT : ColorLUT = 1           ' Color desaturation LUTs: 1 to 11, where 1 is normal and 11 is black'n'white
Dim VolumeDial : VolumeDial = 0.8             ' Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Dim BallRollVolume : BallRollVolume = 0.5     ' Level of ball rolling volume. Value between 0 and 1
Dim RampRollVolume : RampRollVolume = 0.5     ' Level of ramp rolling volume. Value between 0 and 1
Dim OutPostMod : OutPostMod = 1         ' Difficulty : 0 = Easy, 1 = Medium, 2 = Hard
Dim DirtyRamp : DirtyRamp = 1         ' 0 - Clean, 1 - Adds dirt to ramp
Dim EnableSSFShaker: EnableSSFShaker = 1    ' Enables SSF shaker for Godzilla sound fx on target hits.  0 for DOF.
Dim ShakerIntensity: ShakerIntensity = 2    ' Low, Normal, High - 1,2,3
Dim VPMDMDVisible: VPMDMDVisible = 1      ' 0 - Not Visible, 1 - Visible,
Dim VPMDMDColor: VPMDMDColor = 1        ' 0 - Orange, 1 - Green


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

  RefractOpt = Table1.Option("Refraction Setting", 0, 2, 1, 2, 0, Array("No Refraction (best performance)", "Sharp Refractions (improved performance)", "Rough Refractions (best visuals)"))
  SetRefractionProbes RefractOpt

  'Difficulty
  OutPostMod = Table1.Option("Outlane Difficulty", 0, 2, 1, 1, 0, Array("Easy", "Medium", "Hard"))
  Dim BP
  Select Case OutPostMod
    Case 0 ' Easy
      lout_hard.collidable = false: lout_medium.collidable = false: lout_easy.collidable = true
      rout_hard.collidable = false: rout_medium.collidable = false: rout_easy.collidable = true
      For Each BP in BP_lPost: BP.y = 1448.5: Next
      For Each BP in BP_rPost: BP.y = 1448.4: Next
    Case 1 ' Medium
      lout_hard.collidable = false: lout_medium.collidable = true: lout_easy.collidable = false
      rout_hard.collidable = false: rout_medium.collidable = true: rout_easy.collidable = false
      For Each BP in BP_lPost: BP.y = 1436.8: Next
      For Each BP in BP_rPost: BP.y = 1436.9: Next
    Case Else ' Hard
      lout_hard.collidable = true: lout_medium.collidable = false: lout_easy.collidable = false
      rout_hard.collidable = true: rout_medium.collidable = false: rout_easy.collidable = false
      For Each BP in BP_lPost: BP.y = 1425: Next
      For Each BP in BP_rPost: BP.y = 1424.5: Next
  End Select

  'Shaker Intensity
  EnableSSFShaker = Table1.Option("Enable SSF Shaker", 0, 1, 1, 1, 0, Array("Disabled", "Enabled"))

  'Shaker Intensity
  ShakerIntensity = Table1.Option("Shaker Intensity", 1, 3, 1, 2, 0, Array("Low", "Normal", "High"))

  'Ramp dirt mod
  DirtyRamp = Table1.Option("Ramp Dirt", 0, 1, 1, 0, 0, Array("Clean", "Dirty"))
  If DirtyRamp = 1 Then
    RampDirt.visible = true
  Else
    RampDirt.visible = false
  End If

  'Desktop DMD
  VPMDMDVisible = Table1.Option("Desktop DMD", 0, 1, 1, 1, 0, Array("Not Visible", "Visible"))
  If VPMDMDVisible = 1 Then
    Scoretext.Visible = 1
  Else
    Scoretext.Visible = 0
  End If

  'Desktop DMD
  VPMDMDColor = Table1.Option("Desktop DMD Color", 0, 1, 1, 1, 0, Array("Orange", "Green"))
  If VPMDMDColor = 1 Then
    Scoretext.Fontcolor = rgb(0, 128, 0)
    DMD.color = rgb(0, 128, 0)
  Else
    Scoretext.Fontcolor = rgb(255, 165, 0)
    DMD.color = rgb(255, 165, 0)
  End If


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
  SetRoomBrightness LightLevel

    If eventId = 3 And dspTriggered Then dspTriggered = False : DisableStaticPreRendering = False : End If
End Sub


Sub SetRefractionProbes(Opt)
  On Error Resume Next
    Select Case Opt
      Case 0:
        BM_Layer1.RefractionProbe = ""
      Case 1:
        BM_Layer1.RefractionProbe = "Ramp Refractions"
      Case 2:
        BM_Layer1.RefractionProbe = "Ramp Refractions Rough"
    End Select
  On Error Goto 0
End Sub



'******************************************************
'   ZBBR: BALL BRIGHTNESS
'******************************************************

Const BallBrightness =  1        'Ball brightness - Value between 0 and 1 (0=Dark ... 1=Bright)
Const PLOffset = 0.5
Dim PLGain: PLGain = (1-PLOffset)/(1260-2000)

Sub UpdateBallBrightness
  Dim s, b_base, b_r, b_g, b_b, d_w
  b_base = 120 * BallBrightness + 135*gilvl ' orig was 120 and 70

  For s = 0 To UBound(gBOT)
    ' Handle z direction
    d_w = b_base*(1 - (gBOT(s).z-25)/500)
    If d_w < 30 Then d_w = 30
    ' Handle plunger lane
    If InRect(gBOT(s).x,gBOT(s).y,870,2000,870,1260,930,1260,930,2000) Then
      d_w = d_w*(PLOffset+PLGain*(gBOT(s).y-2000))
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
'    Room Brightness
'****************************

' Update these arrays if you want to change more materials with room light level
Dim RoomBrightnessMtlArray: RoomBrightnessMtlArray = Array("VLM.Bake.Active","VLM.Bake.Solid","VLM.Bake.Ramp","rampMaterial")

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


'*************************************************************
' Solenoid Call Backs
'*************************************************************
SolCallBack(1) = "SolRelease"
SolCallBack(2) = "Auto_Plunger"
SolCallBack(6) = "SolShaker"
SolCallback(17) = "SolRampDiverter"

'****************
' Solenoid Lights
'****************
SolModCallBack(7)="Flash07"     'Flash Spinner x2
SolModCallBack(18)="Flash18"    'Flash LT
SolModCallBack(19)="Flash19"    'Flash RT
SolModCallBack(20)="Flash20"    'Flash Inner LT Orbit x2
SolModCallBack(25)="Flash1"     'F1 Flash Pops x4
SolModCallBack(26)="Flash2"     'F2 Flash Top LT x2
SolModCallBack(27)="Flash3"     'F3 Flash CTR Playfield
SolModCallBack(28)="Flash4"     'F4 Flash Ramp
SolModCallBack(29)="Flash5"     'F5 Flash Lite Godz
SolModCallBack(30)="Flash6"     'F6 Flash Score Does
SolModCallBack(31)="Flash7"     'F7 Flash Slings x4
SolModCallBack(32)="Flash8"     'F8 Flash Top RT x2

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"


'****************
' Modulated Calls
'****************

const DebugFlashers = false


Sub Flash07(pwm)
  If DebugFlashers then debug.print "Flash07 "&pwm
  Dim p: p = pwm / 255.0
  f07.state = p
  f07b.state = p
End Sub
Sub f07_Animate
  AdjustBulbTint f07, BL_flashers_f07
End Sub


f18_mayo.intensityscale = 0
Sub Flash18(pwm)        'Dome
  If DebugFlashers then debug.print "Flash18 "&pwm
  f18.state = pwm
  f18_mayo.intensityscale = pwm
End Sub
Sub f18_Animate
  AdjustBulbTint f18, BL_flashers_f18
End Sub


f19_mayo.intensityscale = 0
Sub Flash19(pwm)        'Dome
  If DebugFlashers then debug.print "Flash19 "&pwm
  f19.state = pwm
  f19_mayo.intensityscale = pwm
End Sub
Sub f19_Animate
  AdjustBulbTint f19, BL_flashers_f19
End Sub


Sub Flash20(pwm)
  If DebugFlashers then debug.print "Flash20 "&pwm
  f20a.state = pwm
  f20b.state = pwm
End Sub
Sub f20a_Animate
  AdjustBulbTint f20a, BL_flashers_f20a
  AdjustBulbTint f20b, BL_flashers_f20b
End Sub


Sub Flash1(pwm)
  If DebugFlashers then debug.print "Flash1 "&pwm
  f1a.state = pwm
  f1b.state = pwm
  f1c.state = pwm
  f1.state = pwm
End Sub
Sub f1_Animate
  AdjustBulbTint f1, BL_flashers_f1
  AdjustBulbTint f1a, BL_flashers_f1a
  AdjustBulbTint f1b, BL_flashers_f1b
  AdjustBulbTint f1c, BL_flashers_f1c
End Sub


f2_mayo.intensityscale = 0
f2a_mayo.intensityscale = 0
Sub Flash2(pwm)
  If DebugFlashers then debug.print "Flash2 "&pwm
  f2.state = pwm
  f2a.state = pwm
  f2_mayo.intensityscale = pwm
  f2a_mayo.intensityscale = pwm
End Sub
Sub f2_Animate
  AdjustBulbTint f2, BL_flashers_f2
  AdjustBulbTint f2a, BL_flashers_f2a
End Sub


Sub Flash3(pwm)
  If DebugFlashers then debug.print "Flash3 "&pwm
  f3.state = pwm
End Sub
Sub f3_Animate
  AdjustBulbTint f3, BL_flashers_f3
End Sub


f4_mayo.intensityscale = 0
Sub Flash4(pwm)
  If DebugFlashers then debug.print "Flash4 "&pwm
  f4.state = pwm
  f4_mayo.intensityscale = pwm * 0.2
End Sub
Sub f4_Animate
  AdjustBulbTint f4, BL_flashers_f4
End Sub


f5_mayo.intensityscale = 0
Sub Flash5(pwm)
  If DebugFlashers then debug.print "Flash5 "&pwm
  f5.state = pwm
  f5_mayo.intensityscale = pwm * 0.2
End Sub
Sub f5_Animate
  AdjustBulbTint f5, BL_flashers_f5
End Sub


f6_mayo.intensityscale = 0
Sub Flash6(pwm)
  If DebugFlashers then debug.print "Flash6 "&pwm
  f6.state = pwm
  f6_mayo.intensityscale = pwm * 0.5
End Sub
Sub f6_Animate
  AdjustBulbTint f6, BL_flashers_f6
End Sub


Sub Flash7(pwm)
  If DebugFlashers then debug.print "Flash7 "&pwm
  f7lb.state = pwm
  f7lt.state = pwm
  f7rb.state = pwm
  f7rt.state = pwm
End Sub
Sub f71b_Animate
  AdjustBulbTint f7lb, BL_flashers_f7lb
  AdjustBulbTint f7lt, BL_flashers_f7lt
  AdjustBulbTint f7rb, BL_flashers_f7rb
  AdjustBulbTint f7rt, BL_flashers_f7rt
End Sub


f8_mayo.intensityscale = 0
Sub Flash8(pwm)       'Dome
  If DebugFlashers then debug.print "Flash8 "&pwm
  f8.state = pwm
  f8a.state = pwm
  f8_mayo.intensityscale = pwm
End Sub
Sub f8_Animate
  AdjustBulbTint f8, BL_flashers_f8
  AdjustBulbTint f8a, BL_flashers_f8a
End Sub


' This simulates the extra red light that an incandescent bulb emits as it cools down. (experimental)
Sub AdjustBulbTint(light, lightmaps)
    Dim p : p = light.GetInPlayIntensity / (light.Intensity * light.IntensityScale)
    Dim r : r = 1.0
    Dim g : g = p^0.15
    Dim b : b = (p - 0.04) / 0.96 : If b < 0 Then b = 0
    Dim i : i = .2126 * r^2.2 + .7152 * g^2.2 + .0722 * b^2.2
    Dim v : v = RGB(Int(255 * r), Int(255 * g), Int(255 * b))
    dim x : For Each x in lightmaps : x.color = v : x.opacity = 100.0 / i : Next
End Sub


'****************
' Flippers
'****************
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

Sub LeftFlipper_Collide(parm)
  CheckLiveCatch ActiveBall, LeftFlipper, LFCount, parm
  LF.ReProcessBalls ActiveBall
  LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
  CheckLiveCatch ActiveBall, RightFlipper, RFCount, parm
  RF.ReProcessBalls ActiveBall
  RightFlipperCollide parm
End Sub

Sub LeftFlipper_Animate
  Dim lfa : lfa = LeftFlipper.CurrentAngle
  FlipperLSh.RotZ = LeftFlipper.CurrentAngle
  Dim BP : For Each BP in BP_LFlipper: BP.RotZ = lfa: Next
End Sub

Sub RightFlipper_Animate
  Dim rfa : rfa = RightFlipper.CurrentAngle
    FlipperRSh.RotZ = RightFlipper.CurrentAngle
  Dim BP : For Each BP in BP_RFlipper: BP.RotZ = rfa: Next
End Sub


'******************************************************
' Trough
'******************************************************
Sub sw11_Hit():Controller.Switch(11) = 1:UpdateTrough:End Sub 'first in trough
Sub sw11_UnHit():Controller.Switch(11) = 0:End Sub
Sub sw12_Hit():Controller.Switch(12) = 1:End Sub
Sub sw12_UnHit():Controller.Switch(12) = 0:End Sub
Sub sw13_Hit():Controller.Switch(13) = 1:End Sub
Sub sw13_UnHit():Controller.Switch(13) = 0:End Sub
Sub sw14_Hit():Controller.Switch(14) = 1:End Sub
Sub sw14_UnHit():Controller.Switch(14) = 0:UpdateTrough:End Sub 'last in trough

Dim jjx

Sub UpdateTrough()
    jjx = 1
  UpdateTroughTimer.Interval = 200
  UpdateTroughTimer.Enabled = 1
End Sub

Sub UpdateTroughTimer_Timer()
  If sw14.BallCntOver = 0 Then sw13.kick 60, 9
  If sw13.BallCntOver = 0 Then sw12.kick 60, 9
  If sw12.BallCntOver = 0 Then sw11.kick 60, 9
  if jjx = 5 then
        Me.Enabled = 0
    else
        jjx = jjx + 1
    end if
End Sub




'******************************************************
' Drain & Release
'******************************************************
Sub Drain_Hit()
  RandomSoundDrain drain
  UpdateTrough
  vpmTimer.AddTimer 500, "Drain.kick 60, 20'"
End Sub

Sub SolRelease(enabled)
  If enabled Then
    vpmTimer.PulseSw 15
    sw14.kick 60, 9
    RandomSoundBallRelease sw14
  End If
End Sub

Sub Auto_Plunger(Enabled)
  If Enabled Then
    PlungerIM.AutoFire
  End If
End Sub

Sub SolRampDiverter(enabled)
  If enabled then
    RandomSoundRampDiverterDivert DiverterPosition
    RandomSoundRampDiverterHold 1, DiverterPosition
    Diverter1.IsDropped = 1
    Diverter2.IsDropped = 0
  Else
    RandomSoundRampDiverterBack DiverterPosition
    RandomSoundRampDiverterHold 0, DiverterPosition
    Diverter1.IsDropped = 0
    Diverter2.IsDropped = 1
  End If
End Sub

Sub SolShaker(Enabled)
  if Enabled then
    SoundShakerSSF 1
  Else
    SoundShakerSSF 0
  End if
end sub

'**************
' GI Functions
'**************
dim gilvl:gilvl = 0

set GICallback2 = GetRef("UpdateGI")
Sub UpdateGI(no, level)
  For each xx in GI:xx.State = level: Next
  If level >= 0.5 And gilvl < 0.5 Then
    dim xx
    gilvl = 1
'   PlaySound "fx_relay"
    Sound_GI_Relay 1, GIRelayPosition
  ElseIf level <= 0.4 And gilvl > 0.4 Then
    gilvl = 0
'   PlaySound "fx_relay"
    Sound_GI_Relay 0, GIRelayPosition
  End If
  gilvl = level
End Sub

'******************************************************
' Key Down / Key Ups
'******************************************************

Sub Table1_KeyDown(ByVal KeyCode)

  If keycode = LeftFlipperKey Then VR_CabFlipperLeft.X = VR_CabFlipperLeft.X + 10
  If keycode = RightFlipperKey Then VR_CabFlipperRight.X = VR_CabFlipperRight.X - 10

  If keycode = LeftFlipperKey Then FlipperActivate LeftFlipper, LFPress
  If keycode = RightFlipperKey Then FlipperActivate RightFlipper, RFPress
  If keycode = PlungerKey or keycode = LockBarKey Then
    Controller.Switch(53)=1
    TimerPlunger.Enabled = True
    TimerPlunger2.Enabled = False
  End If

  If keycode = LeftTiltKey Then Nudge 90, 2:SoundNudgeLeft()
  If keycode = RightTiltKey Then Nudge 270, 2:SoundNudgeRight()
  If keycode = CenterTiltKey Then Nudge 0, 1:SoundNudgeCenter()
  If keycode = StartGameKey Then soundStartButton()

  If keycode = keyInsertCoin1 or keycode = keyInsertCoin2 or keycode = keyInsertCoin3 or keycode = keyInsertCoin4 Then
    Select Case Int(rnd*3)
      Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
      Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
      Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
    End Select
  End If

  If KeyDownHandler(keycode) Then Exit Sub

End Sub

Sub Table1_KeyUp(ByVal KeyCode)

  If keycode = LeftFlipperKey Then VR_CabFlipperLeft.X = VR_CabFlipperLeft.X - 10
  If keycode = RightFlipperKey Then VR_CabFlipperRight.X = VR_CabFlipperRight.X + 10

  If keycode = LeftFlipperKey Then FlipperDeActivate LeftFlipper, LFPress
  If keycode = RightFlipperKey Then FlipperDeActivate RightFlipper, RFPress

  If KeyCode = PlungerKey or Keycode = LockBarKey Then
    Controller.Switch(53)=0
  End If

  If KeyUpHandler(keycode) Then Exit Sub

End Sub

'******************************************************
' Plunger
'******************************************************

' Impulse Plunger
Dim PlungerIM
Const IMPowerSetting = 55
Const IMTime = 0.6
Set plungerIM = New cvpmImpulseP
With plungerIM
  .InitImpulseP swplunger, IMPowerSetting, IMTime
  .Random 0.3
  .InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
  .CreateEvents "plungerIM"
End With

'******************************************************
' Switches
'******************************************************

'Stand Up Targets
Sub sw17_hit:STHit 17 : End Sub
Sub sw18_hit:STHit 18 : End Sub
Sub sw19_hit:STHit 19 : End Sub

Sub sw21_hit:STHit 21 : End Sub
Sub sw22_hit:STHit 22 : End Sub
Sub sw23_hit:STHit 23 : End Sub

Sub sw29_hit:STHit 29: End Sub
Sub sw30_hit:STHit 30: End Sub
Sub sw31_hit:STHit 31: End Sub
Sub sw32_hit:STHit 32: End Sub

Sub sw20_Hit:STHit 20 : End Sub
Sub sw24_Hit:STHit 24 : End Sub

'Wire Triggers
Sub sw16_Hit : Controller.Switch(16)=1:End Sub
Sub sw16_Unhit : Controller.Switch(16)=0:End Sub
Sub sw41_Hit : Controller.Switch(41)=1:End Sub
Sub sw41_Unhit : Controller.Switch(41)=0:End Sub
Sub sw42_Hit : Controller.Switch(42)=1:End Sub
Sub sw42_Unhit : Controller.Switch(42)=0:End Sub
Sub sw43_Hit : Controller.Switch(43)=1:End Sub
Sub sw43_Unhit : Controller.Switch(43)=0:End Sub
Sub sw44_Hit : Controller.Switch(44)=1:End Sub
Sub sw44_Unhit : Controller.Switch(44)=0:End Sub
Sub sw47_Hit : Controller.Switch(47)=1:End Sub
Sub sw47_Unhit : Controller.Switch(47)=0:End Sub
Sub sw48_Hit : Controller.Switch(48)=1:End Sub
Sub sw48_Unhit : Controller.Switch(48)=0:End Sub
Sub sw57_Hit : Controller.Switch(57)=1:End Sub
Sub sw57_Unhit : Controller.Switch(57)=0:End Sub
Sub sw58_Hit : Controller.Switch(58)=1: leftInlaneSpeedLimit: End Sub
Sub sw58_Unhit : Controller.Switch(58)=0:End Sub
Sub sw60_Hit : Controller.Switch(60)=1:End Sub
Sub sw60_Unhit : Controller.Switch(60)=0:End Sub
Sub sw61_Hit : Controller.Switch(61)=1: rightInlaneSpeedLimit: End Sub
Sub sw61_Unhit : Controller.Switch(61)=0:End Sub

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


' Switch Animations

' In-Lane / Out-Lane Rollovers
Sub sw57_Animate
  Dim z : z = sw57.CurrentAnimOffset
  Dim BL : For Each BL in BP_sw57 : BL.transz = z: Next
End Sub

Sub sw58_Animate
  Dim z : z = sw58.CurrentAnimOffset
  Dim BL : For Each BL in BP_sw58 : BL.transz = z: Next
End Sub

Sub sw61_Animate
  Dim z : z = sw61.CurrentAnimOffset
  Dim BL : For Each BL in BP_sw61 : BL.transz = z: Next
End Sub

Sub sw60_Animate
  Dim z : z = sw60.CurrentAnimOffset
  Dim BL : For Each BL in BP_sw60 : BL.transz = z: Next
End Sub

Sub sw47_Animate
  Dim z : z = sw47.CurrentAnimOffset
  Dim BL : For Each BL in BP_sw47 : BL.transz = z: Next
End Sub

Sub sw48_Animate
  Dim z : z = sw48.CurrentAnimOffset
  Dim BL : For Each BL in BP_sw47_001 : BL.transz = z: Next
End Sub

Sub sw41_Animate
  Dim z : z = sw41.CurrentAnimOffset
  Dim BL : For Each BL in BP_sw41: BL.transz = z: Next
End Sub

Sub sw42_Animate
  Dim z : z = sw42.CurrentAnimOffset
  Dim BL : For Each BL in BP_sw42 : BL.transz = z: Next
End Sub

Sub sw43_Animate
  Dim z : z = sw43.CurrentAnimOffset
  Dim BL : For Each BL in BP_sw43 : BL.transz = z: Next
End Sub

Sub sw44_Animate
  Dim z : z = sw44.CurrentAnimOffset
  Dim BL : For Each BL in BP_sw44 : BL.transz = z: Next
End Sub

' Ramp Gate Triggers
Sub sw26_Hit:vpmTimer.PulseSw 26:End Sub
Sub sw28_Hit:vpmTimer.PulseSw 28:End Sub

' Gate Triggers
Sub sw45_Hit:vpmTimer.PulseSw 45:End Sub

' Gate Animations
Sub sw26_Animate
    Dim a : a = sw26.CurrentAngle
    Dim BL : For Each BL in BP_sw26 : BL.rotx = -a: Next 'this needs reverse direction
End Sub

Sub sw28_Animate
    Dim a : a = sw28.CurrentAngle
    Dim BL : For Each BL in BP_sw28 : BL.rotx = -a: Next 'this needs reverse direction
End Sub

Sub sw45_Animate
  Dim a : a = sw45.CurrentAngle
  Dim BL : For Each BL in BP_sw45 : BL.rotx = -a: Next 'this needs reverse direction
End Sub

Sub Gate1_Animate
  Dim a : a = Gate1.CurrentAngle
  Dim BL : For Each BL in BP_Gate1 : BL.rotx = a: Next
End Sub

Sub Gate2_Animate
  Dim a : a = Gate2.CurrentAngle
  Dim BL : For Each BL in BP_Gate2 : BL.rotx = a: Next
End Sub

' Spinners
Sub sw46_Spin:vpmTimer.PulseSw 46:SoundSpinner sw46:End Sub

' Spinner Animations
Sub sw46_Animate
  Dim spinangle:spinangle = sw46.currentangle
  Dim BL : For Each BL in BP_sw46 : BL.RotX = -spinangle: Next
  SpinnerShadow.size_y = abs(sin( (spinangle+180) * (2*PI/360)) * 5)
End Sub

'**************
' Bumper Animations
'**************

' Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw(49):RandomSoundBumperTop Bumper1:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw(50):RandomSoundBumperMiddle Bumper2:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw(51):RandomSoundBumperBottom Bumper3:End Sub

' Bumper Animations
Sub Bumper1_Animate
  Dim z, BL
  z = Bumper1.CurrentRingOffset
  For Each BL in BP_Bumper1 : BL.transz = z: Next
End Sub

Sub Bumper2_Animate
  Dim z, BL
  z = Bumper2.CurrentRingOffset
  For Each BL in BP_Bumper2 : BL.transz = z: Next
End Sub

Sub Bumper3_Animate
  Dim z, BL
  z = Bumper3.CurrentRingOffset
  For Each BL in BP_Bumper3 : BL.transz = z: Next
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
    If r = 0 Then For Each x in BP_BS1: x.Z = z: Next
    If r = 1 Then For Each x in BP_BS2: x.Z = z: Next
    If r = 2 Then For Each x in BP_BS3: x.Z = z: Next
  Next
End Sub


'**************
' Ramp Triggers
'**************
Sub TriggerSexyRamp001_Hit
  WireRampOn True
  RampHelper2Hit
End Sub

Sub TriggerSexyRamp001_UnHit
  If Activeball.vely > 0 Then WireRampOff
End Sub

Sub TriggerSexyRamp002_Hit
  WireRampOff
End Sub

Sub TriggerSexyRamp003_Hit
  WireRampOff
End Sub

Sub TriggerSexyRamp004_Hit
  WireRampOff
End Sub

Sub TriggerSexyRamp005_Hit
  WireRampOff
End Sub

Sub TriggerRightRamp001_Hit
  WireRampOn True
  RampHelper1Hit
End Sub

Sub TriggerRightRamp001_UnHit
  If Activeball.velx < 0 Then WireRampOff
End Sub

Sub TriggerRightRamp002_Hit
  WireRampOff
End Sub


'*****************************************************************
' Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'*****************************************************************

Dim RStep, Lstep

Sub LeftSlingShot_Slingshot
  LS.VelocityCorrect(ActiveBall)
  vpmTimer.PulseSw 59
  RandomSoundSlingShotLeft BM_LEMK
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

Sub RightSlingShot_Slingshot
  RS.VelocityCorrect(ActiveBall)
  vpmTimer.PulseSw 62
  RandomSoundSlingShotRight BM_REMK
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


'******************************************************
' ZNFF:  FLIPPER CORRECTIONS by nFozzy
'******************************************************
'
' There are several steps for taking advantage of nFozzy's flipper solution.  At a high level we'll need the following:
' 1. flippers with specific physics settings
' 2. custom triggers for each flipper (TriggerLF, TriggerRF)
' 3. an object or point to tell the script where the tip of the flipper is at rest (EndPointLp, EndPointRp)
' 4. and, special scripting
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
' | EOS Torque         | 0.3            | 0.3                   | 0.275                  | 0.275              |
' | EOS Torque Angle   | 4              | 4                     | 6                      | 6                  |
'

'******************************************************
' Flippers Polarity (Select appropriate sub based on era)
'******************************************************

Dim LF
Set LF = New FlipperPolarity
Dim RF
Set RF = New FlipperPolarity

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

Sub dPosts_Hit(idx)
  RubbersD.dampen ActiveBall
  TargetBouncer ActiveBall, 1
End Sub

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
Const TargetBouncerFactor  = 0.7  'Level of bounces. Recommmended value of 0.7

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
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelExistingStatic(playsoundparams, aVol, tableobj)
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelExistingStaticLoop(playsoundparams, aVol, tableobj)
    PlaySound playsoundparams, -1, min(aVol,1) * VolumeDial, AudioPan(tableobj), 0, 0, 1, 1, AudioFade(tableobj)
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
  Playsound soundname, 1,min(aVol,1) * VolumeDial, AudioPan(ActiveBall), 0,0,0, 1, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtBallVolM (Soundname, aVol)
  Playsound soundname, 1,min(aVol,1) * VolumeDial, AudioPan(ActiveBall), 0,0,0, 0, AudioFade(ActiveBall)
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
' debug.print "Roll"
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

  FlipperCradleCollision ball1, ball2, velocity

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

'/////////////////////////////  RAMP COLLISIONS  ////////////////////////////

dim RRHit1_volume, RRHit2_volume, RRHit3_volume, RRHit4_volume, RRHit5_volume, RRHit6_volume
dim LRHit1_volume, LRHit2_volume, LRHit3_volume, LRHit4_volume, LRHit5_volume

Dim LeftRampSoundLevel
Dim RightRampSoundLevel
Dim RampFallbackSoundLevel

Dim FlapSoundLevel

'///////////////////////-----Ramps-----///////////////////////
'///////////////////////-----Plastic Ramps-----///////////////////////
LeftRampSoundLevel = 0.1                        'volume level; range [0, 1]
RightRampSoundLevel = 0.1                       'volume level; range [0, 1]
RampFallbackSoundLevel = 0.2                      'volume level; range [0, 1]

'///////////////////////-----Ramp Flaps-----///////////////////////
FlapSoundLevel = 0.8                          'volume level; range [0, 1]



'/////////////////////////////  PLASTIC RIGHT RAMP SOUNDS  ////////////////////////////
sub RRHit1_Hit()
  RRHit1_volume = RightRampSoundLevel
  PlaySoundAtLevelStatic ("TOM_C_Ramp_1_Improved2"), RRHit1_volume, RRHit1
end Sub

Sub RRHit2_Hit()
  RRHit2_volume = RightRampSoundLevel
  RRHit1.TimerInterval = 5
  RRHit1.TimerEnabled = 1
  PlaySoundAtLevelStatic ("TOM_C_Ramp_2_Improved2"), RRHit2_volume, RRHit2
End Sub

Sub RRHit3_Hit()
  RRHit3_volume = RightRampSoundLevel
  RRHit2.TimerInterval = 5
  RRHit2.TimerEnabled = 1
  PlaySoundAtLevelStatic ("TOM_C_Ramp_3_Improved2"), RRHit3_volume, RRHit3
End Sub

Sub RRHit4_Hit()
  RRHit4_volume = RightRampSoundLevel
  RRHit3.TimerInterval = 5
  RRHit3.TimerEnabled = 1
  PlaySoundAtLevelStatic ("TOM_C_Ramp_4_Improved2"), RRHit4_volume, RRHit4
End Sub

Sub RRHit5_Hit()
  RRHit5_volume = RightRampSoundLevel
  RRHit4.TimerInterval = 5
  RRHit4.TimerEnabled = 1
  PlaySoundAtLevelStatic ("TOM_C_Ramp_5_Improved2"), RRHit5_volume, RRHit5
End Sub

Sub RRHit6_Hit()
  RRHit6_volume = RightRampSoundLevel
  RRHit5.TimerInterval = 5
  RRHit5.TimerEnabled = 1
  PlaySoundAtLevelStatic ("TOM_C_Ramp_6_Improved2"), RRHit6_volume, RRHit6
End Sub

'/////////////////////////////  PLASTIC RIGHT RAMP SOUND TIMERS  ////////////////////////////
Sub RRHit1_Timer()
' debug.print "1: " & RRHit1_volume
  If RRHit1_volume > 0 Then
    RRHit1_volume = RRHit1_volume - 0.05
    PlaySoundAtLevelExistingStatic ("TOM_C_Ramp_1_Improved2"), RRHit1_volume, RRHit1
  Else
    Me.TimerEnabled = 0
  End If
End Sub

Sub RRHit2_Timer()
' debug.print "2: " & RRHit2_volume
  If RRHit2_volume > 0 Then
    RRHit2_volume = RRHit2_volume - 0.05
    PlaySoundAtLevelExistingStatic ("TOM_C_Ramp_2_Improved2"), RRHit2_volume, RRHit2
  Else
    Me.TimerEnabled = 0
  End If
End Sub

Sub RRHit3_Timer()
' debug.print "3: " & RRHit3_volume
  If RRHit3_volume > 0 Then
    RRHit3_volume = RRHit3_volume - 0.05
    PlaySoundAtLevelExistingStatic ("TOM_C_Ramp_3_Improved2"), RRHit3_volume, RRHit3
  Else
    Me.TimerEnabled = 0
  End If
End Sub

Sub RRHit4_Timer()
' debug.print "4: " & RRHit4_volume
  If RRHit4_volume > 0 Then
    RRHit4_volume = RRHit4_volume - 0.05
    PlaySoundAtLevelExistingStatic ("TOM_C_Ramp_4_Improved2"), RRHit4_volume, RRHit4
  Else
    Me.TimerEnabled = 0
  End If
End Sub

Sub RRHit5_Timer()
' debug.print "5: " & RRHit5_volume
  If RRHit5_volume > 0 Then
    RRHit5_volume = RRHit5_volume - 0.05
    PlaySoundAtLevelExistingStatic ("TOM_C_Ramp_5_Improved2"), RRHit5_volume, RRHit5
  Else
    Me.TimerEnabled = 0
  End If
End Sub



'/////////////////////////////  PLASTIC LEFT RAMP SOUNDS  ////////////////////////////

sub LRHit1_Hit()
  LRHit1_volume = LeftRampSoundLevel
  PlaySoundAtLevelStatic ("TOM_C_Ramp_1_Improved2"), LRHit1_volume, LRHit1
end Sub

Sub LRHit2_Hit()
  LRHit2_volume = LeftRampSoundLevel
  LRHit1.TimerInterval = 5
  LRHit1.TimerEnabled = 1
  PlaySoundAtLevelStatic ("TOM_C_Ramp_2_Improved2"), LRHit2_volume, LRHit2
End Sub

Sub LRHit3_Hit()
  LRHit3_volume = LeftRampSoundLevel
  LRHit2.TimerInterval = 5
  LRHit2.TimerEnabled = 1
  PlaySoundAtLevelStatic ("TOM_C_Ramp_3_Improved2"), LRHit3_volume, LRHit3
End Sub

Sub LRHit4_Hit()
  LRHit4_volume = LeftRampSoundLevel
  LRHit3.TimerInterval = 5
  LRHit3.TimerEnabled = 1
  PlaySoundAtLevelStatic ("TOM_C_Ramp_4_Improved2"), LRHit4_volume, LRHit4
End Sub

Sub LRHit5_Hit()
  LRHit5_volume = LeftRampSoundLevel
  LRHit4.TimerInterval = 5
  LRHit4.TimerEnabled = 1
  PlaySoundAtLevelStatic ("TOM_C_Ramp_5_Improved2"), LRHit5_volume, LRHit5
End Sub


'Plate by the snake ramp
'Sub LRHitPlate_Hit()
' PlaySoundAtLevelStatic ("TOM_R_Ramp_1"), LeftRampSoundLevel, LRHitPlate 'metal plate
'End Sub
'
''before snake ramp vortex
'Sub SRHit1_Hit()
' PlaySoundAtLevelStatic ("TOM_C_Ramp_2_Improved2"), LeftRampSoundLevel, SRHit1
'end sub

sub Diverter2_hit()
' debug.print "diverter2"
  PlaySoundAtLevelStatic ("TOM_R_Ramp_3"), LeftRampSoundLevel * 2, activeball
end sub


'/////////////////////////////  PLASTIC LEFT RAMP SOUND TIMERS ////////////////////////////
Sub LRHit1_Timer()
' debug.print "1: " & LRHit1_volume
  If LRHit1_volume > 0 Then
    LRHit1_volume = LRHit1_volume - 0.05
    PlaySoundAtLevelExistingStatic ("TOM_C_Ramp_1_Improved2"), LRHit1_volume, LRHit1
  Else
    Me.TimerEnabled = 0
  End If
End Sub

Sub LRHit2_Timer()
' debug.print "2: " & LRHit2_volume
  If LRHit2_volume > 0 Then
    LRHit2_volume = LRHit2_volume - 0.05
    PlaySoundAtLevelExistingStatic ("TOM_C_Ramp_2_Improved2"), LRHit2_volume, LRHit2
  Else
    Me.TimerEnabled = 0
  End If
End Sub

Sub LRHit3_Timer()
' debug.print "3: " & LRHit3_volume
  If LRHit3_volume > 0 Then
    LRHit3_volume = LRHit3_volume - 0.05
    PlaySoundAtLevelExistingStatic ("TOM_C_Ramp_3_Improved2"), LRHit3_volume, LRHit3
  Else
    Me.TimerEnabled = 0
  End If
End Sub

Sub LRHit4_Timer()
' debug.print "4: " & LRHit4_volume
  If LRHit4_volume > 0 Then
    LRHit4_volume = LRHit4_volume - 0.05
    PlaySoundAtLevelExistingStatic ("TOM_C_Ramp_4_Improved2"), LRHit4_volume, LRHit4
  Else
    Me.TimerEnabled = 0
  End If
End Sub


'/////////////////////////////  PLASTIC RAMPS FLAPS  ////////////////////////////
'/////////////////////////////  PLASTIC RAMPS FLAPS - EVENTS  ////////////////////////////
Sub RampHelper1Hit()
' debug.print "TriggerRightRamp001: " & ActiveBall.VelX
  If (ActiveBall.VelX < 0) Then
    'ball is traveling down the playfield
    RandomSoundRampFlapDown()

  ElseIf (ActiveBall.VelX > 0) Then
    'ball is traveling up the playfield
    RandomSoundRampFlapUp()

  End If
End Sub

Sub RampHelper2Hit()
' debug.print "TriggerSexyRamp001: " & ActiveBall.VelY
  If (ActiveBall.VelY > 0) Then
    'ball is traveling down the playfield
    RandomSoundRampFlapDown()
  ElseIf (ActiveBall.VelY < 0) Then
    RandomSoundRampFlapUp()

  End If
End Sub



'/////////////////////////////  PLASTIC RAMPS FLAPS - SOUNDS  ////////////////////////////
Sub RandomSoundRampFlapUp()
' debug.print "flap up"
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySoundAtLevelActiveBall ("TOM_Ramp_Flap_Up_1"), FlapSoundLevel
    Case 2 : PlaySoundAtLevelActiveBall ("TOM_Ramp_Flap_Up_2"), FlapSoundLevel
    Case 3 : PlaySoundAtLevelActiveBall ("TOM_Ramp_Flap_Up_3"), FlapSoundLevel
  End Select
End Sub

Sub RandomSoundRampFlapDown()
' debug.print "flap down"
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySoundAtLevelActiveBall ("TOM_Ramp_Flap_Down_1"), FlapSoundLevel
    Case 2 : PlaySoundAtLevelActiveBall ("TOM_Ramp_Flap_Down_2"), FlapSoundLevel
    Case 3 : PlaySoundAtLevelActiveBall ("TOM_Ramp_Flap_Down_3"), FlapSoundLevel
  End Select
End Sub



'//  SHAKER:
Dim ShakerSoundLevel

ShakerSoundLevel = 1                          'volume level; range [0, 1]
Const Cartridge_Table_Specifics     = "SY_TNA_REV02" 'Spooky Total Nuclear Annihilation Cartridge REV02

'//////////////////////////////////  SHAKER  ////////////////////////////////////


Sub SoundShakerSSF(toggle)
  If EnableSSFShaker = 1 Then
    Select Case toggle
      Case 1
        Select Case ShakerIntensity
          Case 1
            PlaySoundAtLevelExistingStaticLoop (Cartridge_Table_Specifics & "_Shaker_Level_01_RampUp_and_Loop"), ShakerSoundLevel, ShakerPosition
          Case 2
            PlaySoundAtLevelExistingStaticLoop (Cartridge_Table_Specifics & "_Shaker_Level_02_RampUp_and_Loop"), ShakerSoundLevel, ShakerPosition
          Case 3
            PlaySoundAtLevelExistingStaticLoop (Cartridge_Table_Specifics & "_Shaker_Level_03_RampUp_and_Loop"), ShakerSoundLevel, ShakerPosition
        End Select
      Case 0
        Select Case shakerintensity
          Case 1
            PlaySoundAtLevelStatic (Cartridge_Table_Specifics & "_Shaker_Level_01_RampDown_Only"), ShakerSoundLevel, ShakerPosition
            StopSound Cartridge_Table_Specifics & "_Shaker_Level_01_RampUp_and_Loop"
          Case 2
            PlaySoundAtLevelStatic (Cartridge_Table_Specifics & "_Shaker_Level_02_RampDown_Only"), ShakerSoundLevel, ShakerPosition
            StopSound Cartridge_Table_Specifics & "_Shaker_Level_02_RampUp_and_Loop"
          Case 3
            PlaySoundAtLevelStatic (Cartridge_Table_Specifics & "_Shaker_Level_03_RampDown_Only"), ShakerSoundLevel, ShakerPosition
            StopSound Cartridge_Table_Specifics & "_Shaker_Level_03_RampUp_and_Loop"
        End Select
    End Select
  End If
End Sub


' diverter ssf sounds reworked from Diner SSF sounds - iaakki
'////////////////////////////////////////////////////////////////////////////////
'////          Mechanical Sounds, by Fleep                                   ////
'////                     Last Updated: January, 2022                        ////
'////////////////////////////////////////////////////////////////////////////////
'
'/////////////////////////////////  CARTRIDGES  /////////////////////////////////
'
'//  Specify which mechanical sound cartridge to use for each group of elements.
'//  Mechanical sounds naming convention: <CARTRIDGE>_<Soundset_Name>
'//
'//  Cartridge name is composed using the following convention:
'//  <TABLE MANUFACTURER ABBREVIATION>_<TABLE NAME ABBREVIATION>_<SOUNDSET REVISION NUMBER>
'//
'//  General Mechanical Sounds Cartridges:

Const Cartridge_Diverters       = "WS_DNR_REV01" 'Williams Diner Cartridge REV01

Dim Solenoid_Diverter_Enabled_SoundLevel, Solenoid_Diverter_Hold_SoundLevel, Solenoid_Diverter_Disabled_SoundLevel

Solenoid_Diverter_Enabled_SoundLevel = 0.5
Solenoid_Diverter_Hold_SoundLevel = 0.05
Solenoid_Diverter_Disabled_SoundLevel = 0.2

'///////////////////////  RAMP DIVERTER SOLENOID - DIVERT  //////////////////////
Sub RandomSoundRampDiverterDivert(obj)
  PlaySoundAtLevelStatic (Cartridge_Diverters & "_Diverter_Divert_" & Int(Rnd*4)+1), Solenoid_Diverter_Enabled_SoundLevel, obj
End Sub

'///////////////////////  RAMP DIVERTER SOLENOID - BACK  ///////////////////////
Sub RandomSoundRampDiverterBack(obj)
  PlaySoundAtLevelStatic (Cartridge_Diverters & "_Diverter_Back_" & Int(Rnd*4)+1), Solenoid_Diverter_Disabled_SoundLevel, obj
End Sub

'////////////////////  RAMP DIVERTER SOLENOID - MAGNET SOUND  ///////////////////
Sub RandomSoundRampDiverterHold(toggle, obj)
  Select Case toggle
    Case 1
      PlaySoundAtLevelStaticLoop SoundFX(Cartridge_Diverters & "_Diverter_Hold_Loop",DOFShaker), Solenoid_Diverter_Hold_SoundLevel, obj
    Case 0
      StopSound Cartridge_Diverters & "_Diverter_Hold_Loop"
  End Select
End Sub

'******************************************************
'****  END FLEEP MECHANICAL SOUNDS
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
  '   Dim BOT
  '   BOT = GetBalls

  ' stop the sound of deleted balls
  For b = UBound(gBOT) + 1 To tnob - 1
    ' Comment the next line if you are not implementing Dyanmic Ball Shadows
    'If AmbientBallShadowOn = 0 Then BallShadowA(b).visible = 0
    rolling(b) = False
    StopSound("BallRoll_" & b)
  Next

  ' exit the sub if no balls on the table
  If UBound(gBOT) =  - 1 Then Exit Sub

  ' play the rolling sound for each ball
  For b = 0 To UBound(gBOT)
    If BallVel(gBOT(b)) > 1 And gBOT(b).z < 30 And Not InRect(gBOT(b).x,gBOT(b).y,579,2015,845,1870,845,1959,590,2090) Then
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

    ' "Static" Ball Shadows
    ' Comment the next If block, if you are not implementing the Dynamic Ball Shadows
    'If AmbientBallShadowOn = 0 Then
    ' If gBOT(b).Z > 30 Then
    '   BallShadowA(b).height = gBOT(b).z - BallSize / 4    'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping the ramp
    ' Else
    '   BallShadowA(b).height = 0.1
    ' End If
    ' BallShadowA(b).Y = gBOT(b).Y + offsetY
    ' BallShadowA(b).X = gBOT(b).X + offsetX
    ' BallShadowA(b).visible = 1
    'End If

    'Fix stuck ball on Left Ramp dips
        If gBOT(b).z > 94 and gBOT(b).z < 100 and gBOT(b).x < 70 and gBOT(b).y < 1175 and gBOT(b).y > 1110 and gBOT(b).vely < 0.1 then
            gBOT(b).vely = gBOT(b).vely - 0.1
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
'     * Set RampBalls and RampType variable to Total Number of Balls
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
Dim RampBalls(8,2)
'x,0 = ball x,1 = ID, 2 = Protection against ending early (minimum amount of updates)

'0,0 is boolean on/off, 0,1 unused for now
RampBalls(0,0) = False

' RampType
' Setup: Set this array to the number Total number of balls that can be tracked at one time + 1.  5 ball multiball then set value to 6
' Description: Array type indexed on BallId and a values used to deterimine what type of ramp the ball is on: False = Wire Ramp, True = Plastic Ramp
Dim RampType(8)

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
      Debug.print "WireRampOn error, ball queue is full: " & vbNewLine & _
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

Sub RampRollUpdate() ' Timer update
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

'******************************************************
'**** END RAMP ROLLING SFX
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
Dim objBallShadow(7)

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
' ZRST: STAND-UP TARGETS by Rothbauerw
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
Dim ST17, ST18, ST19, ST20, ST21, ST22, ST23, ST24, ST29, ST30, ST31, ST32

'Set array with stand-up target objects
'
'StandupTargetvar = Array(primary, prim, switch)
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
Set ST22 = (new StandupTarget)(sw22, BM_sw22, 22, 0)
Set ST23 = (new StandupTarget)(sw23, BM_sw23, 23, 0)
Set ST24 = (new StandupTarget)(sw24, BM_sw24, 24, 0)
Set ST29 = (new StandupTarget)(sw29, BM_sw29, 29, 0)
Set ST30 = (new StandupTarget)(sw30, BM_sw30, 30, 0)
Set ST31 = (new StandupTarget)(sw31, BM_sw31, 31, 0)
Set ST32 = (new StandupTarget)(sw32, BM_sw32, 32, 0)

'Add all the Stand-up Target Arrays to Stand-up Target Animation Array
'   STAnimationArray = Array(ST1, ST2, ....)
Dim STArray
STArray = Array(ST17, ST18, ST19, ST20, ST21, ST22, ST23, ST24, ST29, ST30, ST31, ST32)

'Configure the behavior of Stand-up Targets
Const STAnimStep = 1.5  'vpunits per animation step (control return to Start)
Const STMaxOffset = 9   'max vp units target moves when hit

Const STMass = 0.2    'Mass of the Stand-up Target (between 0 and 1), higher values provide more resistance

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
    prim.transy = - STMaxOffset
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

    ty = BM_sw22.transy
  For each BP in BP_sw22 : BP.transy = ty: Next

    ty = BM_sw23.transy
  For each BP in BP_sw23 : BP.transy = ty: Next

    ty = BM_sw24.transy
  For each BP in BP_sw24 : BP.transy = ty: Next

    ty = BM_sw29.transy
  For each BP in BP_sw29 : BP.transy = ty: Next

    ty = BM_sw30.transy
  For each BP in BP_sw30 : BP.transy = ty: Next

    ty = BM_sw31.transy
  For each BP in BP_sw31 : BP.transy = ty: Next

    ty = BM_sw32.transy
  For each BP in BP_sw32 : BP.transy = ty: Next

End Sub

'******************************************************
'****   END STAND-UP TARGETS
'******************************************************


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




' Make bubbles random speed
Sub BeerTimer_Timer()

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

End Sub






'0.001 - Sixtoe - Initial Table Construction
'0.002 - TastyWasps / Apophis - Solenoid / Flasher Lighting
'0.003 - TastyWasps / Apophis - GI Lighting
'0.008 - TastyWasps / Apophis - nFozzy Physics
'0.009 - TastyWasps / Apophis - Fleep Sounds
'0.010 - Sixtoe - Corrected ramp switch heights, filled in all playfield gaps and reinforced areas, added and renamed extra outlane difficulty posts, changed fleep sling sounds to use physics primitive.
'0.011 - tomate - new ramp primitive added --> right ramp height raised to match with the plastics height, left exit fixed, central hole fixed (needs testing)
'0.012 - Sixtoe - Simplified primitive ramp, changed right ramp switch height, adjusted primitive ramp in numerous places, redid flipper trigger areas, tweaked layout, raised apron area to stop flyballs getting trapped on it.
'0.014 - TastyWasps - Initial VR Room
'0.015 - Sixtoe - Sorting out flashers and flasher assignments, including adding missing flashers and missing doubled up flashers, minor GI alterations, added of L28 and L32 lamps behind left and right targets.
'0.016 - tomate - first 1k test batch added, missing VPX materials aplied, all the VPX objects made invisble
'0.017 - Sixtoe - All the rest of the VPX objects made invisible ;)
'0.018 - tomate - Playfield_mesh added, unused textures deleted
'0.019 - tomate - Reimported batch with new toolkit version, splited into two prims main ramp and clear plastic (Layer0 and Layer1)
'0.020 - TastyWasps - Started implementing movable code. Added standup target code.
'0.021 - apophis - Updated movable code anticipation of blender fixes. Updated standup target code. Cleaned up some timers. Fixed flasher light names. Update light fader models. Set up raytraced shadows (added GI_3 light)
'0.022 - tomate - new 2k batch added, light room code added (thanks Apophis!)
'0.023 - tomate - new 2k batch added, optimized meshes added, unwrapped textures for PF, toys and targets, normal maps for Zilla, bumpers sockets moved to movable collection, main ramp and sona lastic are now transparent
'0.024 - apophis - Added SetLocale. Fixed/added standup target animations. Added frametimer. Animated bumper skirts. Added ball brightness code. Added ramp refractions. Added playfield reflections.
'0.025 - apophis - Reduced ST collidable prim heights.
'0.026 - apophis - Finished implementing ST code. Added inlane speed limit hack.
'0.027 - tomate - 4k batch added
'0.028 - apophis - Added outpost difficulty option. Fixed slingshot animation. Added/fixed ramp sfx triggers. Added ramp material. Adjusted playfield reflections.
'0.029 - iaakki - added more ramp ssf sounds, there are few spares that can be used for that diverter area
'0.030 - iaakki - Added diverter2 hit sound, Added GI relay sounds (no other relays found), reversed the animation of sw45, debug print cleanup
'0.031 - apophis - Added playfield image so that it reflects off of balls. Made MagTop and MagLeft grab center true.
'0.032 - tomate - More fixes in blender side, new 4k batch added
'0.033 - iaakki - SSF shaker added. Disable it if you use DOF shaker
'0.034 - apophis - Re-fixed: Render probe assignments to ramp and playfield.
'0.035 - TastyWasps - FlexDMD options menu
'0.035 - TastyWasps - VR touchups
'0.037 - apophis - Set playfield "hide parts behind". Force VPM 3.6 and PWM flashers. Minor option menu cleanup.
'0.038 - Sixtoe - Various VR tweaks,VR DMD enabled.
'0.039 - tomate - An extra layer of dust and scratches added over the main ramp
'0.040 - apophis - Hid VR stuff when not in VR. Hooked rampMaterial up to room lighting. Made ramp dirt an option.
'0.041 - TastyWasps - Fire button and beer bubbles animated (VR)
'0.042 - TastyWasps - Insert reflections fixed.  VR grills and DMD placement. POVs.
'0.043 - iaakki - disabled rolling sounds for trough balls and modified trough operations, SSF Shaker default level change, minor changes to ball reflection colors for godzilla footprints, diverter sounds added
'0.044 - apophis - Added and tuned new experimental cradle collision flipper trick.
'0.045 - tomate - Left outlane normals fixed
'0.046 - iaakki - cradle collition velocity filter added, insert ball reflection saturations reduced
'0.047 - TastyWasps - Hauntfreaks backglass added for VR. Cabinet POV too. VPMDMD object added. Desktop BG updated with minimal Godzilla logo.
'0.048 - tomate - Inserts depth corrected
'0.049 - apophis - Fixed post passes. Re-fixed: Ramp material assignment. Render probe assignments to ramp and playfield.
'0.050 - Sixtoe - Physical ramp tweaked.
'0.051 - Sixtoe - Fixed gate animations for big ramp, fixed l39 look with iaakki and added visible lighting for l28 and l32 so a re-render isn't needed.
'0.052 - apophis - Slowed down the GI fade-up speed to match PAPA video better. Re-added VPMDMD object.
'0.053 - apophis - Hooked up VR room to light level option. Fixed VR keydown/up button issue. Fixed VR DMD color issue.
'0.054 - TastyWasps - Added VPMDMD options to Flex Menu for on/off and color choice.
'v1.0 -  RELEASE
'v1.0.1 - apophis - BM_Playfield "Hide parts behind" checked. Fixed VolumeDial adjustment. Fixed AnimateBumperSkirts.
'v1.0.2 - iaakki - put some grease onto glass
'v1.0.3 - apophis - Added fix for stuck ball on Left Ramp dips (thanks rothbauerw).
'v1.0.4 - apophis - Set playfield reflection probe roughness to 0.
'v1.1 -  RELEASE
'v1.1.1 - apophis - Fixed spinner rotation direction. Fixed switch animations. Utilizing new DisableStaticPrerendering when options menu is active (requires VPX build 1666 or higher)
'v1.1.2 - apophis - Made PF reflection probe fully dynamic. Enabled reflection of Parts.
'v1.1.3 - apophis - Disabled reflection of Halo lights l28a and l32a
'v1.1.4 - Sixtoe - Added captive ball area roof to stop balls falling in it.
'v1.1.5 - apophis - Updated default desktop POV. Converted options menu to tweak UI. Reworked ball and room brightness stuff.
'v1.2 -  RELEASE
'v1.2.1 - Niwak - Update for PinMame 3.6 physic outputs
'v1.2.2 - apophis - Fixed timers causing stutters.
'v1.2.3 - apophis - Flipper shadow fix. Add table info and rules to Tweak menu. Added screenshot.
'v1.2.4 - apophis - Using GICallback2 now. Added refraction option to tweak menu. Updated flipper physics scripts and triggers. Updated DisableStaticPreRendering functionality to be compatible with VPX 10.8.1 API.
'v1.2.5 - apophis - Added temperature tint on flasher domes (thanks Niwak)
'v1.2.6 - apophis - Added ball shadows. Updated VR backglass image (thanks HauntFreaks).
'v1.3 -  RELEASE

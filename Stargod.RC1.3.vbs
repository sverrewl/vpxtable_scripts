'*************************************************************

' Stargod (Zaccaria 1980)

' Full rebuild based my own mod which in turn was based on a table by Goldchicco / 32Assassin

' A few months in the making due to having to leann how to use Blender and the awesome

' Credits:
' RetroRitchie: Blender rebuild, Playfield Image cleaup and partial redraw, plastics clean up, Asteroid VR room, VPW physics implementation,
' Niwak: Lighmapper tool and pinball parts library

' sorry if i missed anyone!

'Thanks to: VPW for the tutorials and example Table
'Movieguru for upscaled PF imagery
'Fluffhead35 and Wylte for some great tips and guidance earlier on.
'PinstratsDan, DGrimmReaper, Cliffy, StudlyGoorite, Apophis, Bord, Rajo Joey, Darth Vito, somatik, Rothbauerw. BountyBob, Devious626, Thalamus, Gerdang, Colvert for testing
'Apophis, Bord, TastyWasps, Sixtoe for general advice and help
'Hauntfreaks for the updated backglass and some lighting tips
'Rajo Joey for the English Apron Instructons
'Tomate for tips and sharing his Godzilla Blender file for learning
'Devious626 for the authentic cabinet art


'**********************************************************************************************************

Option Explicit
Randomize

SetLocale 1033

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const TableVersion = "RC1.3"


' VLM  Arrays - Start
' Arrays per baked part
Dim BP_Bumper1_Ring_001: BP_Bumper1_Ring_001=Array(BM_Bumper1_Ring_001, LM_Lamps_L27_Bumper1_Ring_001, LM_FL_Bumper1_Ring_001)
Dim BP_Bumper_Socket: BP_Bumper_Socket=Array(BM_Bumper_Socket, LM_GI_Bumper_Socket, LM_DY_GI_006_Bumper_Socket, LM_DY_GL_Special_T_Bumper_Socke, LM_FL_Bumper_Socket)
Dim BP_Gate1: BP_Gate1=Array(BM_Gate1, LM_FL_Gate1)
Dim BP_Gate2_001: BP_Gate2_001=Array(BM_Gate2_001, LM_GI_Gate2_001, LM_DY_GI_006_Gate2_001)
Dim BP_Gate3_001: BP_Gate3_001=Array(BM_Gate3_001, LM_GI_Gate3_001, LM_DY_GL_Special_T_Gate3_001)
Dim BP_L1Sling1: BP_L1Sling1=Array(BM_L1Sling1, LM_GI_L1Sling1, LM_DY_GI_009_L1Sling1, LM_FL_L1Sling1)
Dim BP_L1Sling2: BP_L1Sling2=Array(BM_L1Sling2, LM_GI_L1Sling2, LM_DY_GI_009_L1Sling2, LM_FL_L1Sling2)
Dim BP_LFP: BP_LFP=Array(BM_LFP, LM_GI_LFP, LM_DY_GI_011_LFP, LM_DY_GI_012_LFP)
Dim BP_LFP_UP_001: BP_LFP_UP_001=Array(BM_LFP_UP_001, LM_GI_LFP_UP_001, LM_DY_GI_011_LFP_UP_001, LM_DY_GI_012_LFP_UP_001)
Dim BP_LSling1: BP_LSling1=Array(BM_LSling1, LM_GI_LSling1, LM_DY_GI_011_LSling1, LM_DY_GI_012_LSling1, LM_FL_LSling1)
Dim BP_LSling2: BP_LSling2=Array(BM_LSling2, LM_GI_LSling2, LM_DY_GI_011_LSling2, LM_FL_LSling2)
Dim BP_Lockdown: BP_Lockdown=Array(BM_Lockdown)
Dim BP_Parts: BP_Parts=Array(BM_Parts, LM_GI_Parts, LM_DY_GI_006_Parts, LM_DY_GI_009_Parts, LM_DY_GI_010_Parts, LM_DY_GI_011_Parts, LM_DY_GI_012_Parts, LM_DY_GI_013_Parts, LM_DY_GI_Special_B_Parts, LM_DY_GL_Special_T_Parts, LM_Lamps_L13_Parts, LM_Lamps_L2_Parts, LM_Lamps_L20_Parts, LM_Lamps_L24_Parts, LM_Lamps_L25_Parts, LM_Lamps_L27_Parts, LM_Lamps_L3_Parts, LM_Lamps_L4_Parts, LM_Lamps_L5_Parts, LM_Lamps_L58_Parts, LM_Lamps_L59_Parts, LM_FL_Parts)
Dim BP_Plastics: BP_Plastics=Array(BM_Plastics, LM_GI_Plastics, LM_DY_GI_006_Plastics, LM_DY_GI_009_Plastics, LM_DY_GI_010_Plastics, LM_DY_GI_011_Plastics, LM_DY_GI_012_Plastics, LM_DY_GI_013_Plastics, LM_DY_GI_Special_B_Plastics, LM_DY_GL_Special_T_Plastics, LM_Lamps_L2_Plastics, LM_Lamps_L24_Plastics, LM_Lamps_L27_Plastics, LM_Lamps_L3_Plastics, LM_Lamps_L30_Plastics, LM_Lamps_L4_Plastics, LM_Lamps_L53_Plastics, LM_FL_Plastics)
Dim BP_Playfield: BP_Playfield=Array(BM_Playfield, LM_GI_Playfield, LM_DY_GI_006_Playfield, LM_DY_GI_009_Playfield, LM_DY_GI_010_Playfield, LM_DY_GI_011_Playfield, LM_DY_GI_012_Playfield, LM_DY_GI_013_Playfield, LM_DY_GI_Special_B_Playfield, LM_DY_GL_Special_T_Playfield, LM_Lamps_L1_Playfield, LM_Lamps_L14_Playfield, LM_Lamps_L15_Playfield, LM_Lamps_L16_Playfield, LM_Lamps_L17_Playfield, LM_Lamps_L2_Playfield, LM_Lamps_L20_Playfield, LM_Lamps_L21_Playfield, LM_Lamps_L22_Playfield, LM_Lamps_L24_Playfield, LM_Lamps_L25_Playfield, LM_Lamps_L3_Playfield, LM_Lamps_L30_Playfield, LM_Lamps_L33_Playfield, LM_Lamps_L34_Playfield, LM_Lamps_L35_Playfield, LM_Lamps_L36_Playfield, LM_Lamps_L37_Playfield, LM_Lamps_L38_Playfield, LM_Lamps_L39_Playfield, LM_Lamps_L4_Playfield, LM_Lamps_L40_Playfield, LM_Lamps_L41_Playfield, LM_Lamps_L42_Playfield, LM_Lamps_L44_Playfield, LM_Lamps_L5_Playfield, LM_Lamps_L52_Playfield, LM_Lamps_L53_Playfield, LM_FL_Playfield)
Dim BP_R1Sling1: BP_R1Sling1=Array(BM_R1Sling1, LM_GI_R1Sling1, LM_DY_GL_Special_T_R1Sling1, LM_FL_R1Sling1)
Dim BP_R1Sling2: BP_R1Sling2=Array(BM_R1Sling2, LM_GI_R1Sling2, LM_DY_GL_Special_T_R1Sling2, LM_FL_R1Sling2)
Dim BP_RFP: BP_RFP=Array(BM_RFP, LM_GI_RFP, LM_DY_GI_011_RFP, LM_DY_GI_012_RFP)
Dim BP_RFP_UP: BP_RFP_UP=Array(BM_RFP_UP, LM_GI_RFP_UP, LM_DY_GI_011_RFP_UP, LM_DY_GI_012_RFP_UP)
Dim BP_RSling1: BP_RSling1=Array(BM_RSling1, LM_GI_RSling1, LM_DY_GI_012_RSling1, LM_DY_GI_013_RSling1, LM_FL_RSling1)
Dim BP_RSling2: BP_RSling2=Array(BM_RSling2, LM_GI_RSling2, LM_DY_GI_011_RSling2, LM_DY_GI_012_RSling2, LM_FL_RSling2)
Dim BP_Rails_VR: BP_Rails_VR=Array(BM_Rails_VR, LM_GI_Rails_VR, LM_FL_Rails_VR)
Dim BP_SlingArmL: BP_SlingArmL=Array(BM_SlingArmL, LM_GI_SlingArmL, LM_DY_GI_011_SlingArmL, LM_FL_SlingArmL)
Dim BP_SlingArmL1: BP_SlingArmL1=Array(BM_SlingArmL1, LM_GI_SlingArmL1, LM_DY_GI_009_SlingArmL1, LM_FL_SlingArmL1)
Dim BP_SlingArmR: BP_SlingArmR=Array(BM_SlingArmR, LM_GI_SlingArmR, LM_DY_GI_012_SlingArmR)
Dim BP_SlingArmR1: BP_SlingArmR1=Array(BM_SlingArmR1, LM_GI_SlingArmR1, LM_DY_GL_Special_T_SlingArmR1, LM_FL_SlingArmR1)
Dim BP_Spinner1_Wire: BP_Spinner1_Wire=Array(BM_Spinner1_Wire, LM_GI_Spinner1_Wire, LM_FL_Spinner1_Wire)
Dim BP_Spinner1_Wire_001: BP_Spinner1_Wire_001=Array(BM_Spinner1_Wire_001, LM_GI_Spinner1_Wire_001, LM_DY_GI_Special_B_Spinner1_Wir, LM_FL_Spinner1_Wire_001)
Dim BP_Trigger7: BP_Trigger7=Array(BM_Trigger7, LM_GI_Trigger7, LM_DY_GI_009_Trigger7, LM_FL_Trigger7)
Dim BP_UnderPlastic: BP_UnderPlastic=Array(BM_UnderPlastic, LM_GI_UnderPlastic, LM_DY_GI_009_UnderPlastic, LM_Lamps_L56_UnderPlastic, LM_Lamps_L57_UnderPlastic, LM_Lamps_L58_UnderPlastic, LM_Lamps_L59_UnderPlastic, LM_FL_UnderPlastic)
Dim BP_apron_card_center_1: BP_apron_card_center_1=Array(BM_apron_card_center_1)
Dim BP_apron_card_center_2: BP_apron_card_center_2=Array(BM_apron_card_center_2)
Dim BP_apron_card_center_3: BP_apron_card_center_3=Array(BM_apron_card_center_3)
Dim BP_apron_cards_eng: BP_apron_cards_eng=Array(BM_apron_cards_eng)
Dim BP_apron_cards_it: BP_apron_cards_it=Array(BM_apron_cards_it)
Dim BP_psw32: BP_psw32=Array(BM_psw32, LM_DY_GI_013_psw32, LM_DY_GI_Special_B_psw32, LM_FL_psw32)
Dim BP_psw39: BP_psw39=Array(BM_psw39, LM_GI_psw39, LM_DY_GL_Special_T_psw39, LM_FL_psw39)
Dim BP_sw24p: BP_sw24p=Array(BM_sw24p, LM_GI_sw24p, LM_DY_GI_009_sw24p, LM_DY_GI_010_sw24p, LM_DY_GI_011_sw24p, LM_DY_GI_012_sw24p, LM_FL_sw24p)
Dim BP_sw25p: BP_sw25p=Array(BM_sw25p, LM_GI_sw25p, LM_DY_GI_009_sw25p, LM_DY_GI_010_sw25p, LM_DY_GI_011_sw25p, LM_DY_GI_012_sw25p, LM_DY_GI_013_sw25p, LM_FL_sw25p)
Dim BP_sw26p: BP_sw26p=Array(BM_sw26p, LM_GI_sw26p, LM_DY_GI_009_sw26p, LM_DY_GI_010_sw26p, LM_DY_GI_012_sw26p, LM_DY_GI_013_sw26p, LM_FL_sw26p)
Dim BP_sw27p: BP_sw27p=Array(BM_sw27p, LM_GI_sw27p, LM_DY_GI_009_sw27p, LM_DY_GI_010_sw27p, LM_DY_GI_012_sw27p, LM_DY_GI_013_sw27p, LM_FL_sw27p)
Dim BP_sw40p: BP_sw40p=Array(BM_sw40p, LM_GI_sw40p, LM_FL_sw40p)
Dim BP_sw41p: BP_sw41p=Array(BM_sw41p, LM_GI_sw41p, LM_FL_sw41p)
Dim BP_sw42p: BP_sw42p=Array(BM_sw42p, LM_GI_sw42p, LM_DY_GL_Special_T_sw42p, LM_FL_sw42p)
Dim BP_sw43p: BP_sw43p=Array(BM_sw43p, LM_GI_sw43p, LM_DY_GL_Special_T_sw43p, LM_FL_sw43p)
Dim BP_sw44p: BP_sw44p=Array(BM_sw44p, LM_GI_sw44p, LM_DY_GL_Special_T_sw44p, LM_FL_sw44p)
Dim BP_sw45p: BP_sw45p=Array(BM_sw45p, LM_GI_sw45p, LM_DY_GI_006_sw45p, LM_DY_GL_Special_T_sw45p, LM_FL_sw45p)
Dim BP_sw46p: BP_sw46p=Array(BM_sw46p, LM_GI_sw46p, LM_DY_GI_006_sw46p, LM_DY_GL_Special_T_sw46p, LM_FL_sw46p)
Dim BP_sw47p: BP_sw47p=Array(BM_sw47p, LM_GI_sw47p, LM_DY_GI_006_sw47p, LM_FL_sw47p)
Dim BP_sw_LIL: BP_sw_LIL=Array(BM_sw_LIL, LM_GI_sw_LIL, LM_DY_GI_010_sw_LIL, LM_DY_GI_011_sw_LIL)
Dim BP_sw_LOL: BP_sw_LOL=Array(BM_sw_LOL, LM_GI_sw_LOL, LM_DY_GI_011_sw_LOL)
Dim BP_sw_RIL: BP_sw_RIL=Array(BM_sw_RIL, LM_GI_sw_RIL, LM_DY_GI_012_sw_RIL, LM_DY_GI_013_sw_RIL)
Dim BP_sw_ROL: BP_sw_ROL=Array(BM_sw_ROL, LM_GI_sw_ROL, LM_DY_GI_012_sw_ROL)
Dim BP_sw_TRoll1: BP_sw_TRoll1=Array(BM_sw_TRoll1, LM_GI_sw_TRoll1, LM_DY_GI_006_sw_TRoll1, LM_DY_GL_Special_T_sw_TRoll1)
Dim BP_sw_TRoll2: BP_sw_TRoll2=Array(BM_sw_TRoll2, LM_GI_sw_TRoll2, LM_DY_GI_006_sw_TRoll2, LM_DY_GL_Special_T_sw_TRoll2)
Dim BP_sw_TRoll3: BP_sw_TRoll3=Array(BM_sw_TRoll3, LM_GI_sw_TRoll3, LM_DY_GI_006_sw_TRoll3, LM_DY_GL_Special_T_sw_TRoll3)
' Arrays per lighting scenario
Dim BL_DY_GI_006: BL_DY_GI_006=Array(LM_DY_GI_006_Bumper_Socket, LM_DY_GI_006_Gate2_001, LM_DY_GI_006_Parts, LM_DY_GI_006_Plastics, LM_DY_GI_006_Playfield, LM_DY_GI_006_sw45p, LM_DY_GI_006_sw46p, LM_DY_GI_006_sw47p, LM_DY_GI_006_sw_TRoll1, LM_DY_GI_006_sw_TRoll2, LM_DY_GI_006_sw_TRoll3)
Dim BL_DY_GI_009: BL_DY_GI_009=Array(LM_DY_GI_009_L1Sling1, LM_DY_GI_009_L1Sling2, LM_DY_GI_009_Parts, LM_DY_GI_009_Plastics, LM_DY_GI_009_Playfield, LM_DY_GI_009_SlingArmL1, LM_DY_GI_009_Trigger7, LM_DY_GI_009_UnderPlastic, LM_DY_GI_009_sw24p, LM_DY_GI_009_sw25p, LM_DY_GI_009_sw26p, LM_DY_GI_009_sw27p)
Dim BL_DY_GI_010: BL_DY_GI_010=Array(LM_DY_GI_010_Parts, LM_DY_GI_010_Plastics, LM_DY_GI_010_Playfield, LM_DY_GI_010_sw24p, LM_DY_GI_010_sw25p, LM_DY_GI_010_sw26p, LM_DY_GI_010_sw27p, LM_DY_GI_010_sw_LIL)
Dim BL_DY_GI_011: BL_DY_GI_011=Array(LM_DY_GI_011_LFP, LM_DY_GI_011_LFP_UP_001, LM_DY_GI_011_LSling1, LM_DY_GI_011_LSling2, LM_DY_GI_011_Parts, LM_DY_GI_011_Plastics, LM_DY_GI_011_Playfield, LM_DY_GI_011_RFP, LM_DY_GI_011_RFP_UP, LM_DY_GI_011_RSling2, LM_DY_GI_011_SlingArmL, LM_DY_GI_011_sw24p, LM_DY_GI_011_sw25p, LM_DY_GI_011_sw_LIL, LM_DY_GI_011_sw_LOL)
Dim BL_DY_GI_012: BL_DY_GI_012=Array(LM_DY_GI_012_LFP, LM_DY_GI_012_LFP_UP_001, LM_DY_GI_012_LSling1, LM_DY_GI_012_Parts, LM_DY_GI_012_Plastics, LM_DY_GI_012_Playfield, LM_DY_GI_012_RFP, LM_DY_GI_012_RFP_UP, LM_DY_GI_012_RSling1, LM_DY_GI_012_RSling2, LM_DY_GI_012_SlingArmR, LM_DY_GI_012_sw24p, LM_DY_GI_012_sw25p, LM_DY_GI_012_sw26p, LM_DY_GI_012_sw27p, LM_DY_GI_012_sw_RIL, LM_DY_GI_012_sw_ROL)
Dim BL_DY_GI_013: BL_DY_GI_013=Array(LM_DY_GI_013_Parts, LM_DY_GI_013_Plastics, LM_DY_GI_013_Playfield, LM_DY_GI_013_RSling1, LM_DY_GI_013_psw32, LM_DY_GI_013_sw25p, LM_DY_GI_013_sw26p, LM_DY_GI_013_sw27p, LM_DY_GI_013_sw_RIL)
Dim BL_DY_GI_Special_B: BL_DY_GI_Special_B=Array(LM_DY_GI_Special_B_Parts, LM_DY_GI_Special_B_Plastics, LM_DY_GI_Special_B_Playfield, LM_DY_GI_Special_B_Spinner1_Wir, LM_DY_GI_Special_B_psw32)
Dim BL_DY_GL_Special_T: BL_DY_GL_Special_T=Array(LM_DY_GL_Special_T_Bumper_Socke, LM_DY_GL_Special_T_Gate3_001, LM_DY_GL_Special_T_Parts, LM_DY_GL_Special_T_Plastics, LM_DY_GL_Special_T_Playfield, LM_DY_GL_Special_T_R1Sling1, LM_DY_GL_Special_T_R1Sling2, LM_DY_GL_Special_T_SlingArmR1, LM_DY_GL_Special_T_psw39, LM_DY_GL_Special_T_sw42p, LM_DY_GL_Special_T_sw43p, LM_DY_GL_Special_T_sw44p, LM_DY_GL_Special_T_sw45p, LM_DY_GL_Special_T_sw46p, LM_DY_GL_Special_T_sw_TRoll1, LM_DY_GL_Special_T_sw_TRoll2, LM_DY_GL_Special_T_sw_TRoll3)
Dim BL_FL: BL_FL=Array(LM_FL_Bumper_Socket, LM_FL_Bumper1_Ring_001, LM_FL_Gate1, LM_FL_L1Sling1, LM_FL_L1Sling2, LM_FL_LSling1, LM_FL_LSling2, LM_FL_Parts, LM_FL_Plastics, LM_FL_Playfield, LM_FL_R1Sling1, LM_FL_R1Sling2, LM_FL_RSling1, LM_FL_RSling2, LM_FL_Rails_VR, LM_FL_SlingArmL, LM_FL_SlingArmL1, LM_FL_SlingArmR1, LM_FL_Spinner1_Wire, LM_FL_Spinner1_Wire_001, LM_FL_Trigger7, LM_FL_UnderPlastic, LM_FL_psw32, LM_FL_psw39, LM_FL_sw24p, LM_FL_sw25p, LM_FL_sw26p, LM_FL_sw27p, LM_FL_sw40p, LM_FL_sw41p, LM_FL_sw42p, LM_FL_sw43p, LM_FL_sw44p, LM_FL_sw45p, LM_FL_sw46p, LM_FL_sw47p)
Dim BL_GI: BL_GI=Array(LM_GI_Bumper_Socket, LM_GI_Gate2_001, LM_GI_Gate3_001, LM_GI_L1Sling1, LM_GI_L1Sling2, LM_GI_LFP, LM_GI_LFP_UP_001, LM_GI_LSling1, LM_GI_LSling2, LM_GI_Parts, LM_GI_Plastics, LM_GI_Playfield, LM_GI_R1Sling1, LM_GI_R1Sling2, LM_GI_RFP, LM_GI_RFP_UP, LM_GI_RSling1, LM_GI_RSling2, LM_GI_Rails_VR, LM_GI_SlingArmL, LM_GI_SlingArmL1, LM_GI_SlingArmR, LM_GI_SlingArmR1, LM_GI_Spinner1_Wire, LM_GI_Spinner1_Wire_001, LM_GI_Trigger7, LM_GI_UnderPlastic, LM_GI_psw39, LM_GI_sw24p, LM_GI_sw25p, LM_GI_sw26p, LM_GI_sw27p, LM_GI_sw40p, LM_GI_sw41p, LM_GI_sw42p, LM_GI_sw43p, LM_GI_sw44p, LM_GI_sw45p, LM_GI_sw46p, LM_GI_sw47p, LM_GI_sw_LIL, LM_GI_sw_LOL, LM_GI_sw_RIL, LM_GI_sw_ROL, LM_GI_sw_TRoll1, LM_GI_sw_TRoll2, LM_GI_sw_TRoll3)
Dim BL_Lamps_L1: BL_Lamps_L1=Array(LM_Lamps_L1_Playfield)
Dim BL_Lamps_L13: BL_Lamps_L13=Array(LM_Lamps_L13_Parts)
Dim BL_Lamps_L14: BL_Lamps_L14=Array(LM_Lamps_L14_Playfield)
Dim BL_Lamps_L15: BL_Lamps_L15=Array(LM_Lamps_L15_Playfield)
Dim BL_Lamps_L16: BL_Lamps_L16=Array(LM_Lamps_L16_Playfield)
Dim BL_Lamps_L17: BL_Lamps_L17=Array(LM_Lamps_L17_Playfield)
Dim BL_Lamps_L2: BL_Lamps_L2=Array(LM_Lamps_L2_Parts, LM_Lamps_L2_Plastics, LM_Lamps_L2_Playfield)
Dim BL_Lamps_L20: BL_Lamps_L20=Array(LM_Lamps_L20_Parts, LM_Lamps_L20_Playfield)
Dim BL_Lamps_L21: BL_Lamps_L21=Array(LM_Lamps_L21_Playfield)
Dim BL_Lamps_L22: BL_Lamps_L22=Array(LM_Lamps_L22_Playfield)
Dim BL_Lamps_L24: BL_Lamps_L24=Array(LM_Lamps_L24_Parts, LM_Lamps_L24_Plastics, LM_Lamps_L24_Playfield)
Dim BL_Lamps_L25: BL_Lamps_L25=Array(LM_Lamps_L25_Parts, LM_Lamps_L25_Playfield)
Dim BL_Lamps_L27: BL_Lamps_L27=Array(LM_Lamps_L27_Bumper1_Ring_001, LM_Lamps_L27_Parts, LM_Lamps_L27_Plastics)
Dim BL_Lamps_L3: BL_Lamps_L3=Array(LM_Lamps_L3_Parts, LM_Lamps_L3_Plastics, LM_Lamps_L3_Playfield)
Dim BL_Lamps_L30: BL_Lamps_L30=Array(LM_Lamps_L30_Plastics, LM_Lamps_L30_Playfield)
Dim BL_Lamps_L33: BL_Lamps_L33=Array(LM_Lamps_L33_Playfield)
Dim BL_Lamps_L34: BL_Lamps_L34=Array(LM_Lamps_L34_Playfield)
Dim BL_Lamps_L35: BL_Lamps_L35=Array(LM_Lamps_L35_Playfield)
Dim BL_Lamps_L36: BL_Lamps_L36=Array(LM_Lamps_L36_Playfield)
Dim BL_Lamps_L37: BL_Lamps_L37=Array(LM_Lamps_L37_Playfield)
Dim BL_Lamps_L38: BL_Lamps_L38=Array(LM_Lamps_L38_Playfield)
Dim BL_Lamps_L39: BL_Lamps_L39=Array(LM_Lamps_L39_Playfield)
Dim BL_Lamps_L4: BL_Lamps_L4=Array(LM_Lamps_L4_Parts, LM_Lamps_L4_Plastics, LM_Lamps_L4_Playfield)
Dim BL_Lamps_L40: BL_Lamps_L40=Array(LM_Lamps_L40_Playfield)
Dim BL_Lamps_L41: BL_Lamps_L41=Array(LM_Lamps_L41_Playfield)
Dim BL_Lamps_L42: BL_Lamps_L42=Array(LM_Lamps_L42_Playfield)
Dim BL_Lamps_L44: BL_Lamps_L44=Array(LM_Lamps_L44_Playfield)
Dim BL_Lamps_L5: BL_Lamps_L5=Array(LM_Lamps_L5_Parts, LM_Lamps_L5_Playfield)
Dim BL_Lamps_L52: BL_Lamps_L52=Array(LM_Lamps_L52_Playfield)
Dim BL_Lamps_L53: BL_Lamps_L53=Array(LM_Lamps_L53_Plastics, LM_Lamps_L53_Playfield)
Dim BL_Lamps_L56: BL_Lamps_L56=Array(LM_Lamps_L56_UnderPlastic)
Dim BL_Lamps_L57: BL_Lamps_L57=Array(LM_Lamps_L57_UnderPlastic)
Dim BL_Lamps_L58: BL_Lamps_L58=Array(LM_Lamps_L58_Parts, LM_Lamps_L58_UnderPlastic)
Dim BL_Lamps_L59: BL_Lamps_L59=Array(LM_Lamps_L59_Parts, LM_Lamps_L59_UnderPlastic)
Dim BL_World: BL_World=Array(BM_Bumper_Socket, BM_Bumper1_Ring_001, BM_Gate1, BM_Gate2_001, BM_Gate3_001, BM_L1Sling1, BM_L1Sling2, BM_LFP, BM_LFP_UP_001, BM_LSling1, BM_LSling2, BM_Lockdown, BM_Parts, BM_Plastics, BM_Playfield, BM_R1Sling1, BM_R1Sling2, BM_RFP, BM_RFP_UP, BM_RSling1, BM_RSling2, BM_Rails_VR, BM_SlingArmL, BM_SlingArmL1, BM_SlingArmR, BM_SlingArmR1, BM_Spinner1_Wire, BM_Spinner1_Wire_001, BM_Trigger7, BM_UnderPlastic, BM_apron_card_center_1, BM_apron_card_center_2, BM_apron_card_center_3, BM_apron_cards_eng, BM_apron_cards_it, BM_psw32, BM_psw39, BM_sw24p, BM_sw25p, BM_sw26p, BM_sw27p, BM_sw40p, BM_sw41p, BM_sw42p, BM_sw43p, BM_sw44p, BM_sw45p, BM_sw46p, BM_sw47p, BM_sw_LIL, BM_sw_LOL, BM_sw_RIL, BM_sw_ROL, BM_sw_TRoll1, BM_sw_TRoll2, BM_sw_TRoll3)
' Global arrays
Dim BG_Bakemap: BG_Bakemap=Array(BM_Bumper_Socket, BM_Bumper1_Ring_001, BM_Gate1, BM_Gate2_001, BM_Gate3_001, BM_L1Sling1, BM_L1Sling2, BM_LFP, BM_LFP_UP_001, BM_LSling1, BM_LSling2, BM_Lockdown, BM_Parts, BM_Plastics, BM_Playfield, BM_R1Sling1, BM_R1Sling2, BM_RFP, BM_RFP_UP, BM_RSling1, BM_RSling2, BM_Rails_VR, BM_SlingArmL, BM_SlingArmL1, BM_SlingArmR, BM_SlingArmR1, BM_Spinner1_Wire, BM_Spinner1_Wire_001, BM_Trigger7, BM_UnderPlastic, BM_apron_card_center_1, BM_apron_card_center_2, BM_apron_card_center_3, BM_apron_cards_eng, BM_apron_cards_it, BM_psw32, BM_psw39, BM_sw24p, BM_sw25p, BM_sw26p, BM_sw27p, BM_sw40p, BM_sw41p, BM_sw42p, BM_sw43p, BM_sw44p, BM_sw45p, BM_sw46p, BM_sw47p, BM_sw_LIL, BM_sw_LOL, BM_sw_RIL, BM_sw_ROL, BM_sw_TRoll1, BM_sw_TRoll2, BM_sw_TRoll3)
Dim BG_Lightmap: BG_Lightmap=Array(LM_DY_GI_006_Bumper_Socket, LM_DY_GI_006_Gate2_001, LM_DY_GI_006_Parts, LM_DY_GI_006_Plastics, LM_DY_GI_006_Playfield, LM_DY_GI_006_sw45p, LM_DY_GI_006_sw46p, LM_DY_GI_006_sw47p, LM_DY_GI_006_sw_TRoll1, LM_DY_GI_006_sw_TRoll2, LM_DY_GI_006_sw_TRoll3, LM_DY_GI_009_L1Sling1, LM_DY_GI_009_L1Sling2, LM_DY_GI_009_Parts, LM_DY_GI_009_Plastics, LM_DY_GI_009_Playfield, LM_DY_GI_009_SlingArmL1, LM_DY_GI_009_Trigger7, LM_DY_GI_009_UnderPlastic, LM_DY_GI_009_sw24p, LM_DY_GI_009_sw25p, LM_DY_GI_009_sw26p, LM_DY_GI_009_sw27p, LM_DY_GI_010_Parts, LM_DY_GI_010_Plastics, LM_DY_GI_010_Playfield, LM_DY_GI_010_sw24p, LM_DY_GI_010_sw25p, LM_DY_GI_010_sw26p, LM_DY_GI_010_sw27p, LM_DY_GI_010_sw_LIL, LM_DY_GI_011_LFP, LM_DY_GI_011_LFP_UP_001, LM_DY_GI_011_LSling1, LM_DY_GI_011_LSling2, LM_DY_GI_011_Parts, LM_DY_GI_011_Plastics, LM_DY_GI_011_Playfield, LM_DY_GI_011_RFP, LM_DY_GI_011_RFP_UP, LM_DY_GI_011_RSling2, LM_DY_GI_011_SlingArmL, LM_DY_GI_011_sw24p, LM_DY_GI_011_sw25p, LM_DY_GI_011_sw_LIL, _
  LM_DY_GI_011_sw_LOL, LM_DY_GI_012_LFP, LM_DY_GI_012_LFP_UP_001, LM_DY_GI_012_LSling1, LM_DY_GI_012_Parts, LM_DY_GI_012_Plastics, LM_DY_GI_012_Playfield, LM_DY_GI_012_RFP, LM_DY_GI_012_RFP_UP, LM_DY_GI_012_RSling1, LM_DY_GI_012_RSling2, LM_DY_GI_012_SlingArmR, LM_DY_GI_012_sw24p, LM_DY_GI_012_sw25p, LM_DY_GI_012_sw26p, LM_DY_GI_012_sw27p, LM_DY_GI_012_sw_RIL, LM_DY_GI_012_sw_ROL, LM_DY_GI_013_Parts, LM_DY_GI_013_Plastics, LM_DY_GI_013_Playfield, LM_DY_GI_013_RSling1, LM_DY_GI_013_psw32, LM_DY_GI_013_sw25p, LM_DY_GI_013_sw26p, LM_DY_GI_013_sw27p, LM_DY_GI_013_sw_RIL, LM_DY_GI_Special_B_Parts, LM_DY_GI_Special_B_Plastics, LM_DY_GI_Special_B_Playfield, LM_DY_GI_Special_B_Spinner1_Wir, LM_DY_GI_Special_B_psw32, LM_DY_GL_Special_T_Bumper_Socke, LM_DY_GL_Special_T_Gate3_001, LM_DY_GL_Special_T_Parts, LM_DY_GL_Special_T_Plastics, LM_DY_GL_Special_T_Playfield, LM_DY_GL_Special_T_R1Sling1, LM_DY_GL_Special_T_R1Sling2, LM_DY_GL_Special_T_SlingArmR1, LM_DY_GL_Special_T_psw39, LM_DY_GL_Special_T_sw42p, _
  LM_DY_GL_Special_T_sw43p, LM_DY_GL_Special_T_sw44p, LM_DY_GL_Special_T_sw45p, LM_DY_GL_Special_T_sw46p, LM_DY_GL_Special_T_sw_TRoll1, LM_DY_GL_Special_T_sw_TRoll2, LM_DY_GL_Special_T_sw_TRoll3, LM_FL_Bumper_Socket, LM_FL_Bumper1_Ring_001, LM_FL_Gate1, LM_FL_L1Sling1, LM_FL_L1Sling2, LM_FL_LSling1, LM_FL_LSling2, LM_FL_Parts, LM_FL_Plastics, LM_FL_Playfield, LM_FL_R1Sling1, LM_FL_R1Sling2, LM_FL_RSling1, LM_FL_RSling2, LM_FL_Rails_VR, LM_FL_SlingArmL, LM_FL_SlingArmL1, LM_FL_SlingArmR1, LM_FL_Spinner1_Wire, LM_FL_Spinner1_Wire_001, LM_FL_Trigger7, LM_FL_UnderPlastic, LM_FL_psw32, LM_FL_psw39, LM_FL_sw24p, LM_FL_sw25p, LM_FL_sw26p, LM_FL_sw27p, LM_FL_sw40p, LM_FL_sw41p, LM_FL_sw42p, LM_FL_sw43p, LM_FL_sw44p, LM_FL_sw45p, LM_FL_sw46p, LM_FL_sw47p, LM_GI_Bumper_Socket, LM_GI_Gate2_001, LM_GI_Gate3_001, LM_GI_L1Sling1, LM_GI_L1Sling2, LM_GI_LFP, LM_GI_LFP_UP_001, LM_GI_LSling1, LM_GI_LSling2, LM_GI_Parts, LM_GI_Plastics, LM_GI_Playfield, LM_GI_R1Sling1, LM_GI_R1Sling2, LM_GI_RFP, LM_GI_RFP_UP, LM_GI_RSling1, _
  LM_GI_RSling2, LM_GI_Rails_VR, LM_GI_SlingArmL, LM_GI_SlingArmL1, LM_GI_SlingArmR, LM_GI_SlingArmR1, LM_GI_Spinner1_Wire, LM_GI_Spinner1_Wire_001, LM_GI_Trigger7, LM_GI_UnderPlastic, LM_GI_psw39, LM_GI_sw24p, LM_GI_sw25p, LM_GI_sw26p, LM_GI_sw27p, LM_GI_sw40p, LM_GI_sw41p, LM_GI_sw42p, LM_GI_sw43p, LM_GI_sw44p, LM_GI_sw45p, LM_GI_sw46p, LM_GI_sw47p, LM_GI_sw_LIL, LM_GI_sw_LOL, LM_GI_sw_RIL, LM_GI_sw_ROL, LM_GI_sw_TRoll1, LM_GI_sw_TRoll2, LM_GI_sw_TRoll3, LM_Lamps_L1_Playfield, LM_Lamps_L13_Parts, LM_Lamps_L14_Playfield, LM_Lamps_L15_Playfield, LM_Lamps_L16_Playfield, LM_Lamps_L17_Playfield, LM_Lamps_L2_Parts, LM_Lamps_L2_Plastics, LM_Lamps_L2_Playfield, LM_Lamps_L20_Parts, LM_Lamps_L20_Playfield, LM_Lamps_L21_Playfield, LM_Lamps_L22_Playfield, LM_Lamps_L24_Parts, LM_Lamps_L24_Plastics, LM_Lamps_L24_Playfield, LM_Lamps_L25_Parts, LM_Lamps_L25_Playfield, LM_Lamps_L27_Bumper1_Ring_001, LM_Lamps_L27_Parts, LM_Lamps_L27_Plastics, LM_Lamps_L3_Parts, LM_Lamps_L3_Plastics, LM_Lamps_L3_Playfield, _
  LM_Lamps_L30_Plastics, LM_Lamps_L30_Playfield, LM_Lamps_L33_Playfield, LM_Lamps_L34_Playfield, LM_Lamps_L35_Playfield, LM_Lamps_L36_Playfield, LM_Lamps_L37_Playfield, LM_Lamps_L38_Playfield, LM_Lamps_L39_Playfield, LM_Lamps_L4_Parts, LM_Lamps_L4_Plastics, LM_Lamps_L4_Playfield, LM_Lamps_L40_Playfield, LM_Lamps_L41_Playfield, LM_Lamps_L42_Playfield, LM_Lamps_L44_Playfield, LM_Lamps_L5_Parts, LM_Lamps_L5_Playfield, LM_Lamps_L52_Playfield, LM_Lamps_L53_Plastics, LM_Lamps_L53_Playfield, LM_Lamps_L56_UnderPlastic, LM_Lamps_L57_UnderPlastic, LM_Lamps_L58_Parts, LM_Lamps_L58_UnderPlastic, LM_Lamps_L59_Parts, LM_Lamps_L59_UnderPlastic)
Dim BG_All: BG_All=Array(BM_Bumper_Socket, BM_Bumper1_Ring_001, BM_Gate1, BM_Gate2_001, BM_Gate3_001, BM_L1Sling1, BM_L1Sling2, BM_LFP, BM_LFP_UP_001, BM_LSling1, BM_LSling2, BM_Lockdown, BM_Parts, BM_Plastics, BM_Playfield, BM_R1Sling1, BM_R1Sling2, BM_RFP, BM_RFP_UP, BM_RSling1, BM_RSling2, BM_Rails_VR, BM_SlingArmL, BM_SlingArmL1, BM_SlingArmR, BM_SlingArmR1, BM_Spinner1_Wire, BM_Spinner1_Wire_001, BM_Trigger7, BM_UnderPlastic, BM_apron_card_center_1, BM_apron_card_center_2, BM_apron_card_center_3, BM_apron_cards_eng, BM_apron_cards_it, BM_psw32, BM_psw39, BM_sw24p, BM_sw25p, BM_sw26p, BM_sw27p, BM_sw40p, BM_sw41p, BM_sw42p, BM_sw43p, BM_sw44p, BM_sw45p, BM_sw46p, BM_sw47p, BM_sw_LIL, BM_sw_LOL, BM_sw_RIL, BM_sw_ROL, BM_sw_TRoll1, BM_sw_TRoll2, BM_sw_TRoll3, LM_DY_GI_006_Bumper_Socket, LM_DY_GI_006_Gate2_001, LM_DY_GI_006_Parts, LM_DY_GI_006_Plastics, LM_DY_GI_006_Playfield, LM_DY_GI_006_sw45p, LM_DY_GI_006_sw46p, LM_DY_GI_006_sw47p, LM_DY_GI_006_sw_TRoll1, LM_DY_GI_006_sw_TRoll2, LM_DY_GI_006_sw_TRoll3, _
  LM_DY_GI_009_L1Sling1, LM_DY_GI_009_L1Sling2, LM_DY_GI_009_Parts, LM_DY_GI_009_Plastics, LM_DY_GI_009_Playfield, LM_DY_GI_009_SlingArmL1, LM_DY_GI_009_Trigger7, LM_DY_GI_009_UnderPlastic, LM_DY_GI_009_sw24p, LM_DY_GI_009_sw25p, LM_DY_GI_009_sw26p, LM_DY_GI_009_sw27p, LM_DY_GI_010_Parts, LM_DY_GI_010_Plastics, LM_DY_GI_010_Playfield, LM_DY_GI_010_sw24p, LM_DY_GI_010_sw25p, LM_DY_GI_010_sw26p, LM_DY_GI_010_sw27p, LM_DY_GI_010_sw_LIL, LM_DY_GI_011_LFP, LM_DY_GI_011_LFP_UP_001, LM_DY_GI_011_LSling1, LM_DY_GI_011_LSling2, LM_DY_GI_011_Parts, LM_DY_GI_011_Plastics, LM_DY_GI_011_Playfield, LM_DY_GI_011_RFP, LM_DY_GI_011_RFP_UP, LM_DY_GI_011_RSling2, LM_DY_GI_011_SlingArmL, LM_DY_GI_011_sw24p, LM_DY_GI_011_sw25p, LM_DY_GI_011_sw_LIL, LM_DY_GI_011_sw_LOL, LM_DY_GI_012_LFP, LM_DY_GI_012_LFP_UP_001, LM_DY_GI_012_LSling1, LM_DY_GI_012_Parts, LM_DY_GI_012_Plastics, LM_DY_GI_012_Playfield, LM_DY_GI_012_RFP, LM_DY_GI_012_RFP_UP, LM_DY_GI_012_RSling1, LM_DY_GI_012_RSling2, LM_DY_GI_012_SlingArmR, LM_DY_GI_012_sw24p, _
  LM_DY_GI_012_sw25p, LM_DY_GI_012_sw26p, LM_DY_GI_012_sw27p, LM_DY_GI_012_sw_RIL, LM_DY_GI_012_sw_ROL, LM_DY_GI_013_Parts, LM_DY_GI_013_Plastics, LM_DY_GI_013_Playfield, LM_DY_GI_013_RSling1, LM_DY_GI_013_psw32, LM_DY_GI_013_sw25p, LM_DY_GI_013_sw26p, LM_DY_GI_013_sw27p, LM_DY_GI_013_sw_RIL, LM_DY_GI_Special_B_Parts, LM_DY_GI_Special_B_Plastics, LM_DY_GI_Special_B_Playfield, LM_DY_GI_Special_B_Spinner1_Wir, LM_DY_GI_Special_B_psw32, LM_DY_GL_Special_T_Bumper_Socke, LM_DY_GL_Special_T_Gate3_001, LM_DY_GL_Special_T_Parts, LM_DY_GL_Special_T_Plastics, LM_DY_GL_Special_T_Playfield, LM_DY_GL_Special_T_R1Sling1, LM_DY_GL_Special_T_R1Sling2, LM_DY_GL_Special_T_SlingArmR1, LM_DY_GL_Special_T_psw39, LM_DY_GL_Special_T_sw42p, LM_DY_GL_Special_T_sw43p, LM_DY_GL_Special_T_sw44p, LM_DY_GL_Special_T_sw45p, LM_DY_GL_Special_T_sw46p, LM_DY_GL_Special_T_sw_TRoll1, LM_DY_GL_Special_T_sw_TRoll2, LM_DY_GL_Special_T_sw_TRoll3, LM_FL_Bumper_Socket, LM_FL_Bumper1_Ring_001, LM_FL_Gate1, LM_FL_L1Sling1, LM_FL_L1Sling2, LM_FL_LSling1, _
  LM_FL_LSling2, LM_FL_Parts, LM_FL_Plastics, LM_FL_Playfield, LM_FL_R1Sling1, LM_FL_R1Sling2, LM_FL_RSling1, LM_FL_RSling2, LM_FL_Rails_VR, LM_FL_SlingArmL, LM_FL_SlingArmL1, LM_FL_SlingArmR1, LM_FL_Spinner1_Wire, LM_FL_Spinner1_Wire_001, LM_FL_Trigger7, LM_FL_UnderPlastic, LM_FL_psw32, LM_FL_psw39, LM_FL_sw24p, LM_FL_sw25p, LM_FL_sw26p, LM_FL_sw27p, LM_FL_sw40p, LM_FL_sw41p, LM_FL_sw42p, LM_FL_sw43p, LM_FL_sw44p, LM_FL_sw45p, LM_FL_sw46p, LM_FL_sw47p, LM_GI_Bumper_Socket, LM_GI_Gate2_001, LM_GI_Gate3_001, LM_GI_L1Sling1, LM_GI_L1Sling2, LM_GI_LFP, LM_GI_LFP_UP_001, LM_GI_LSling1, LM_GI_LSling2, LM_GI_Parts, LM_GI_Plastics, LM_GI_Playfield, LM_GI_R1Sling1, LM_GI_R1Sling2, LM_GI_RFP, LM_GI_RFP_UP, LM_GI_RSling1, LM_GI_RSling2, LM_GI_Rails_VR, LM_GI_SlingArmL, LM_GI_SlingArmL1, LM_GI_SlingArmR, LM_GI_SlingArmR1, LM_GI_Spinner1_Wire, LM_GI_Spinner1_Wire_001, LM_GI_Trigger7, LM_GI_UnderPlastic, LM_GI_psw39, LM_GI_sw24p, LM_GI_sw25p, LM_GI_sw26p, LM_GI_sw27p, LM_GI_sw40p, LM_GI_sw41p, LM_GI_sw42p, LM_GI_sw43p, _
  LM_GI_sw44p, LM_GI_sw45p, LM_GI_sw46p, LM_GI_sw47p, LM_GI_sw_LIL, LM_GI_sw_LOL, LM_GI_sw_RIL, LM_GI_sw_ROL, LM_GI_sw_TRoll1, LM_GI_sw_TRoll2, LM_GI_sw_TRoll3, LM_Lamps_L1_Playfield, LM_Lamps_L13_Parts, LM_Lamps_L14_Playfield, LM_Lamps_L15_Playfield, LM_Lamps_L16_Playfield, LM_Lamps_L17_Playfield, LM_Lamps_L2_Parts, LM_Lamps_L2_Plastics, LM_Lamps_L2_Playfield, LM_Lamps_L20_Parts, LM_Lamps_L20_Playfield, LM_Lamps_L21_Playfield, LM_Lamps_L22_Playfield, LM_Lamps_L24_Parts, LM_Lamps_L24_Plastics, LM_Lamps_L24_Playfield, LM_Lamps_L25_Parts, LM_Lamps_L25_Playfield, LM_Lamps_L27_Bumper1_Ring_001, LM_Lamps_L27_Parts, LM_Lamps_L27_Plastics, LM_Lamps_L3_Parts, LM_Lamps_L3_Plastics, LM_Lamps_L3_Playfield, LM_Lamps_L30_Plastics, LM_Lamps_L30_Playfield, LM_Lamps_L33_Playfield, LM_Lamps_L34_Playfield, LM_Lamps_L35_Playfield, LM_Lamps_L36_Playfield, LM_Lamps_L37_Playfield, LM_Lamps_L38_Playfield, LM_Lamps_L39_Playfield, LM_Lamps_L4_Parts, LM_Lamps_L4_Plastics, LM_Lamps_L4_Playfield, LM_Lamps_L40_Playfield, _
  LM_Lamps_L41_Playfield, LM_Lamps_L42_Playfield, LM_Lamps_L44_Playfield, LM_Lamps_L5_Parts, LM_Lamps_L5_Playfield, LM_Lamps_L52_Playfield, LM_Lamps_L53_Plastics, LM_Lamps_L53_Playfield, LM_Lamps_L56_UnderPlastic, LM_Lamps_L57_UnderPlastic, LM_Lamps_L58_Parts, LM_Lamps_L58_UnderPlastic, LM_Lamps_L59_Parts, LM_Lamps_L59_UnderPlastic)
' VLM  Arrays - End



' VLM VR Arrays - Start
' Arrays per baked part
Dim BP_VRVR: BP_VRVR=Array(VRBM_VR, VRLM_GI_VR, VRLM_DY_GI_006_VR, VRLM_DY_GL_Special_T_VR)
' Arrays per lighting scenario
Dim BL_VRDY_GI_006: BL_VRDY_GI_006=Array(VRLM_DY_GI_006_VR)
Dim BL_VRDY_GL_Special_T: BL_VRDY_GL_Special_T=Array(VRLM_DY_GL_Special_T_VR)
Dim BL_VRGI: BL_VRGI=Array(VRLM_GI_VR)
Dim BL_VRWorld: BL_VRWorld=Array(VRBM_VR)
' Global arrays
Dim BGVR_Bakemap: BGVR_Bakemap=Array(VRBM_VR)
Dim BGVR_Lightmap: BGVR_Lightmap=Array(VRLM_DY_GI_006_VR, VRLM_DY_GL_Special_T_VR, VRLM_GI_VR)
Dim BGVR_All: BGVR_All=Array(VRBM_VR, VRLM_DY_GI_006_VR, VRLM_DY_GL_Special_T_VR, VRLM_GI_VR)
' VLM VR Arrays - End



'*******************************************
'  Global Vars
'*******************************************


Const UsingROM  = True        'The UsingROM flag is to indicate code that requires ROM usage.
Const BallSize  = 50        'Sets the ball size
Const BallMass = 1          'Ball mass must be 1
Const tnob = 1            'Total number of balls
Const lob = 0           'Locked balls
Const cGameName = "stargoda"


Dim tablewidth : tablewidth = Table1.width
Dim tableheight : tableheight = Table1.height

'Dim BIPL : BIPL = False

LoadVPM "01200000","zac1.vbs",2.0

Const UseSolenoids  = 1
Const UseLamps    = 1
Const UseSync     = 0
Const HandleMech  = 0
Const UseGI     = 0
Const SCoin = ""
Const TargetBouncerEnabled = 1    '0 = normal standup targets, 1 = bouncy targets, 2 = orig TargetBouncer
Const TargetBouncerFactor = 0.7   'Level of bounces. Recommmended value of 0.7 when TargetBouncerEnabled



'Internal DMD in Desktop Mode, using a textbox (must be called before LoadVPM)
Dim VR, UseVPMDMD, CabinetMode, DesktopMode: DesktopMode = Table1.ShowDT
If RenderingMode = 2 Then  VR = True Else VR = False
If VR Then UseVPMDMD = True Else UseVPMDMD = DesktopMode
If Not DesktopMode and VR = False Then CabinetMode=1 Else CabinetMode=0

'Override for DT VR Testing
'VR = true


'*******************************************
' ZTIM:Timers
'*******************************************

Sub GameTimer_Timer()   'The game timer interval; should be 10 ms
  Cor.Update        'update ball tracking (this sometimes goes in the RDampen_Timer sub)


  'DTShadows
End Sub

' The frame timer interval is -1, so executes at the display frame rate
dim FrameTime, InitFrameTime : InitFrameTime = 0

Sub FrameTimer_Timer() 'The frame timer interval should be -1, so executes at the display frame rate
  FrameTime = GameTime - InitFrameTime
  InitFrameTime = GameTime  'Count frametime
  BGUpdate
  BSUpdate
  'Options_UpdateDMD
  UpdateBallBrightness
  AnimateBumperSkirts
  DoDTAnim      'handle drop target animations
  DoSTAnim      'handle stand up target animations
  UpdateStandupTargets
  UpdateDropTargets
RollingUpdate     'update rolling sounds
End Sub



'*******************************************
'  ZOPT: User Options
'*******************************************

Dim LightLevel : LightLevel = 0.5     ' Level of room lighting (0 to 1), where 0 is dark and 100 is brightest
Dim VolumeDial : VolumeDial = 0.8           'Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Dim BallRollVolume : BallRollVolume = 0.5   'Level of ball rolling volume. Value between 0 and 1
Dim RampRollVolume : RampRollVolume = 0.5     ' Level of ramp rolling volume. Value between 0 and 1
Dim Saturation : Saturation = 2
Dim InstructionCardLang: InstructionCardLang = 1
Dim Aproncard: Aproncard = 1
Dim VRMotion : VRMotion = 2
Dim GIColor : GIColor = 0
Dim TopSpecialColor : TopSpecialColor = 0
Dim Rails : Rails = 1
Dim VRRoom

' Called when options are tweaked by the player.
' - 0: game has started, good time to load options and adjust accordingly
' - 1: an option has changed
' - 2: options have been reseted
' - 3: player closed the tweak UI, good time to update staticly prerendered parts
' Table1.Option arguments are:
' - option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional arry of literal strings
Dim DspTriggered : DspTriggered = False
Sub Table1_OptionEvent(ByVal eventId)
Dim v,r

   If eventId = 1 And Not DspTriggered Then DspTriggered = True : DisableStaticPreRendering = True : End If

    'Sound volumes
    VolumeDial = Table1.Option("Mech Volume", 0, 1, 0.01, 0.8, 1)
    BallRollVolume = Table1.Option("Ball Roll Volume", 0, 1, 0.01, 0.5, 1)

  'Room Brightness
    LightLevel = NightDay/100
  SetRoomBrightness LightLevel   'Uncomment this line for lightmapped tables.

  'GI Bulb Color
  GIColor = Table1.Option("GI Color and Strength", 0, 7, 1, 3, 0, Array("2700k","3000k","3500k","4000k","White","Amber","Yellow","Orange"))
  Select Case GIColor
  Case 0
  SetGIColor cArray(0)
  Case 1
  SetGIColor cArray(1)
  Case 2
  SetGIColor cArray(2)
  Case 3
  SetGIColor cArray(3)
  Case 4
  SetGIColor cArray(4)
  Case 5
  SetGIColor cArray(5)
  Case 6
  SetGIColor cArray(6)
  Case 7
  SetGIColor cArray(7)
  End Select

  TopSpecialColor = Table1.Option("Top Special Flasher Color", 0, 2, 1, 0, 0, Array("Orange (Default)","Red","Pink"))

  'Saturation
  Saturation = Table1.Option("Adjust Saturation", 0, 3, 1, 1, 0, Array("+10%", "0","-10%","-20%"))
  SetSaturation

  'Top Special Flasher Bulb Color
  Select Case TopSpecialColor
  Case 0
  SetSpecialTopColor TopSpecialArray(0)
  Case 1
  SetSpecialTopColor TopSpecialArray(1)
  Case 2
  SetSpecialTopColor TopSpecialArray(2)
  End Select

  'Arpon Card Center
  Aproncard = Table1.Option("Apron Center Card", 0, 2, 1, 0, 0, Array("Zaccaria Logo","Balls Per Game", "Star God Logo"))
  SetApronCard

  'Apron Instruction Card Lang
  InstructionCardLang = Table1.Option("Instruction Card Language", 0, 1, 1, 0, 0, Array("English","Italian"))
  SetInstructionCard

  'Rails
  r = Table1.Option("Rails and Lockdown Visibility", 0, 1, 1, 1, 0, Array("Hide","Show"))
  SetRails(r)

  'VR Room Type
  VRRoom = Table1.Option("VRRoom", 0, 1, 1, 0, 0, Array("Minimal","Asteroid"))
  SetupVRRoom

  'VR Motion Control
  VRMotion = Table1.Option("VR Motion", 0, 3, 1, 1, 0, Array("Off","Slow", "Medium", "Vomit Inducing"))

If eventId = 3 And Not DspTriggered Then DspTriggered = False : DisableStaticPreRendering = False : End If

End Sub



Sub SetInstructionCard
  '*********************
  'SCORECARDS
  '*********************
  if InstructionCardLang = 0 Then BM_apron_cards_it.visible = 0 : BM_apron_cards_eng.visible = 1
  if InstructionCardLang = 1 Then BM_apron_cards_eng.visible = 0 : BM_apron_cards_it.visible = 1
End Sub


Sub SetApronCard
  '*********************
  'CENTER CARD
  '*********************
  if Aproncard = 0 Then BM_apron_card_center_1.visible = 1 : BM_apron_card_center_2.visible = 0 : BM_apron_card_center_3.visible = 0
  if Aproncard = 1 Then BM_apron_card_center_2.visible = 1 : BM_apron_card_center_1.visible = 0 : BM_apron_card_center_3.visible = 0
  if Aproncard = 2 Then BM_apron_card_center_3.visible = 1 : BM_apron_card_center_1.visible = 0 : BM_apron_card_center_2.visible = 0
End Sub


''''''''' GI Color Change options
' reference: https://andi-siess.de/rgb-to-color-temperature/

Sub SetRails(r)

Dim RailItems

if r = 0 OR VR = True Then

    for each RailItems in BP_Rails_VR
      RailItems.visible = 0
    Next

    for each RailItems in BP_Lockdown
      RailItems.visible = 0
    Next




elseif r = 1 then
  for each RailItems in BP_Rails_VR
    RailItems.visible = 1
    Next

for each RailItems in BP_Lockdown
      RailItems.visible = 1
    Next

end if

End Sub

Dim cWhiteFull: cWhiteFull = rgb(255,255,128)
Dim cWhite: cWhite = rgb(255,255,255)
Dim c2700k: c2700k = rgb(255, 169, 87)
Dim c3000k: c3000k = rgb(255, 180, 107)
Dim c3500k: c3500k = rgb(255, 196, 137)
Dim c4000k: c4000k = rgb(255, 209, 163)
Dim cAmber: cAmber = rgb(255,197,143)
Dim cYellow: cYellow = rgb(255,255,0)
Dim cOrange: cOrange = rgb(255,70,5)
Dim cPink: cPink = rgb(255,5,255)
Dim cRed: cRed= rgb(255,5,5)

Dim cArray
cArray = Array(c2700k,c3000k, c3500k,c4000k,cWhite,cAmber,cYellow,cOrange)

Dim TopSpecialArray
TopSpecialArray = Array(cOrange,cRed, cPink)

sub SetGIColor(c)
  Dim xx, BL
  For each xx in BL_GI: xx.color = c: Next
  For each xx in GI: xx.color = c: Next
  For each BL in BL_DY_GI_006: BL.color = c: Next
  For each BL in BL_DY_GI_009: BL.color = c: Next
  For each BL in BL_DY_GI_010: BL.color = c: Next
  For each BL in BL_DY_GI_011: BL.color = c: Next
  For each BL in BL_DY_GI_012: BL.color = c: Next
  For each BL in BL_DY_GI_013: BL.color = c: Next

end Sub

sub SetSpecialTopColor(c)
  Dim BL
  For each BL in BL_DY_GL_Special_T: BL.color = c: Next
  GL_Special_T.color = c
end Sub


Sub SetSaturation
if  Saturation = 0 Then Table1.ColorGradeImage = "colorgradelut256x16+10"
  if Saturation = 1 Then Table1.ColorGradeImage = ""
  if Saturation = 2 Then Table1.ColorGradeImage = "colorgradelut256x16-10"
  if Saturation = 3 Then Table1.ColorGradeImage = "colorgradelut256x16-20"
  if Saturation = 4 Then Table1.ColorGradeImage = "colorgradelut256x16-30"
End Sub


''''''''''''''''''' Start Color Changing GI
'Dim R, G, B
'Dim CCGIStep, xxGiCC
Dim RedRGB, GreenRGB, BlueRGB, CCGIStep, xxGiCC, xxCCGI

RedRGB = 255
GreenRGB = 0
BlueRGB = 0

Sub CCGI_timer ()
  If RedRGB < 0 then RedRGB = 0 end If
  If GreenRGB < 0 then GreenRGB = 0 End If
  If BlueRGB < 0 then BlueRGB = 0 End If
  If RedRGB > 255 then RedRGB = 255 End If
  If GreenRGB > 255 then GreenRGB = 255 End If
  If BlueRGB > 255 then BlueRGB = 255 End If

  If CCGIStep > 0 and CCGIStep < 255 Then
    GreenRGB = GreenRGB + 1
  End If
    If CCGIStep > 255 and CCGIStep < 510 Then
    RedRGB = RedRGB - 1

  End If
  If CCGIStep > 510 and CCGIStep < 765 Then
    BlueRGB = BlueRGB + 1

  End If
  If CCGIStep > 765 and CCGIStep < 1020 Then
    GreenRGB = GreenRGB - 1

  End If
  If CCGIStep > 1020 and CCGIStep < 1275 Then
    RedRGB = RedRGB + 1

  End If
  If CCGIStep > 1275 and CCGIStep < 1530 Then
    BlueRGB = BlueRGB - 1
  End If

  If CCGIStep = 1530 then CCGIStep = 0 End If

  CCGIStep = CCGIStep + 1

End Sub

''''''''''''''''''' End Color Changing GI

'**********************************************************************************************************
'Initiate Table
'**********************************************************************************************************


Dim SGBall1, gBOT

Sub Table1_Init
  vpmInit Me
  On Error Resume Next
    With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine = "Star God ( Zaccaria ) 1980"&chr(13)&""
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

   'Ball initializations need for physical trough
  Set SGBall1 = sw16.CreateSizedballWithMass(Ballsize / 2,Ballmass)
  gBOT = Array(SGBall1)
  Controller.Switch(16) = 1


  'Nudge
  vpmNudge.TiltSwitch=1
    vpmNudge.Sensitivity=4
    vpmNudge.TiltObj=Array(Bumper1,LeftSlingShot,RightSlingShot)

  PinMAMETimer.Interval=PinMAMEInterval
  PinMAMETimer.Enabled= 1

  vpmMapLights AllLamps ' Map all lamps to the corresponding ROM output using the value of TimerInterval of each light object



  'Initialize slings
  RStep = 0:RightSlingShot.Timerenabled=True
  LStep = 0:LeftSlingShot.Timerenabled=True
  RStep1 = 0:RightSlingShot1.Timerenabled=True
  LStep1 = 0:LeftSlingShot1.Timerenabled=True



  If DesktopMode and VR = False Then 'Show Desktop components
  for each xx in BGlights : xx.state = 1 : next
  Else
  for each xx in BGlights : xx.state = 0 : next
  End if


' Room Brightness
  SetRoomBrightness LightLevel/100

End Sub

Sub Table1_Paused:Controller.Pause = 1:End Sub
Sub Table1_unPaused:Controller.Pause = 0:End Sub
Sub Table1_Exit()
  Controller.Pause = False
  Controller.Stop
End Sub


Sub BGUpdate
  dim l
  For Each l in BGlights
    If l.GetInPlayIntensity > 0 Then
      l.Bulb = 0
    Else
      l.Bulb = 1
    End If
  Next

IF Controller.Lamp(60) = 0 Then: l60bg.visible = 0 : Else : l60bg.visible = 1 'p1
IF Controller.Lamp(61) = 0 Then: l61bg.visible = 0 : Else : l61bg.visible = 1 'p2
IF Controller.Lamp(62) = 0 Then: l62bg.visible = 0 : Else : l62bg.visible = 1 'p3
IF Controller.Lamp(63) = 0 Then: l63bg.visible = 0 : Else : l63bg.visible = 1 'p4
IF Controller.Lamp(27) = 0 Then: l27bg.visible = 0 : Else : l27bg.visible = 1 'Bumper/logo

IF Controller.Lamp(19) = 0 Then: l19bg.visible = 0 : Else : l19bg.visible = 1 'Superbonus
IF Controller.Lamp(1) = 0 Then: l1bg.visible = 0 : Else : l1bg.visible = 1 'Bonus Ball

IF Controller.Lamp(51) = 0 Then: l51bg.visible = 0 : Else : l51bg.visible = 1 'Tilt
IF Controller.Lamp(7) = 0 Then: l7bg.visible = 0 : Else : l7bg.visible = 1 'Game Over

End Sub


'*************************************************************
'Solenoid Call backs
'**********************************************************************************************************

SolCallback(2)  ="SolDTTopBankReset"
SolCallback(6)  ="SolDTBottomBankReset"
SolCallback(11) ="SolRelease"
SolCallback(12) = "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(3)  = "SolDT40"
SolCallback(5)  = "SolDT41"
SolCallback(13) = "SolDT42"
SolCallback(14) = "SolDT43"
SolCallback(15) = "SolDT44"
SolCallback(16) = "SolDT45"
SolCallback(17) = "SolDT46"
SolCallback(18) = "SolDT47"

'Playfield GI
'*****GI Lights On
dim xx
For each xx in GI:xx.State = 1: Next

Dim gilvl   'General Illumination light state tracked for Dynamic Ball Shadows
gilvl = 1



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


 'TROUGH DRAIN & RELEASE
Sub SolRelease(enabled)
  If enabled Then
    Controller.Switch(16) = 0
    sw16.kick 60, 20
    RandomSoundBallRelease sw16
  End If
End Sub




'**********************************************************************************************************
'Keys
'**********************************************************************************************************

'KeyDown

Sub Table1_KeyDown(ByVal KeyCode)

  If keycode = LeftFlipperKey Then
    FlipperActivate LeftFlipper, LFPress
  End If

  If keycode = RightFlipperKey Then
    FlipperActivate RightFlipper, RFPress
  End If



If keycode = PlungerKey Then
    Plunger.PullBack
    SoundPlungerPull
End If


If VRRoom <> 0 Then

  If keycode = LeftFlipperKey Then
    Primary_flipper_button_left.x = Primary_flipper_button_left.x + 4
  End IF

  If KeyCode = RightFlipperKey Then
    Primary_flipper_button_right.x = Primary_flipper_button_right.x - 4
  End If

  If keycode = StartGameKey Then
    PinCab_Button1i.y = -375.0651 - 4
  End If


  If keycode = PlungerKey Then
    TimerVRPlunger.Enabled = True
    TimerVRPlunger2.Enabled = False
  End If
End If


  If keycode = AddCreditKey or keycode = AddCreditKey2 Then

    Select Case Int(rnd*3)
      Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
      Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
      Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
    End Select
  End If

  If keycode = StartGameKey Then SoundStartButton
  If keycode = LeftTiltKey Then Nudge 90, 1.5 : SoundNudgeLeft
  If keycode = RightTiltKey Then Nudge 270, 1.5 : SoundNudgeRight
  If keycode = CenterTiltKey Then Nudge 0, 1.5 : SoundNudgeCenter
  If keycode = MechanicalTilt Then SoundNudgeCenter()



  If KeyDownHandler(keycode) Then Exit Sub
End Sub

'KeyUp

Sub Table1_KeyUp(ByVal KeyCode)



  If KeyCode = PlungerKey Then

    Plunger.Fire
    SoundPlungerReleaseBall
'
'     Plunger.Fire
'     If BIPL Then
'       SoundPlungerReleaseBall()     'Plunger release sound when there is a ball in shooter lane
'     Else
'       SoundPlungerReleaseNoBall()     'Plunger release sound when there is no ball in shooter lane
'     End If
    End If


If VRRoom <> 0 Then
    If keycode = PlungerKey Then
      TimerVRPlunger.Enabled = False
      TimerVRPlunger2.Enabled = True
    End If
    If keycode = RightFlipperKey Then
      Primary_flipper_button_right.x = 2103.377
    End If
    If keycode = LeftFlipperKey Then
      Primary_flipper_button_left.x = 2103.801
    End If
    If keycode = StartGameKey Then
      PinCab_Button1i.y = -375.0651
    End If
  End If


  If keycode = LeftFlipperKey Then
    FlipperDeActivate LeftFlipper, LFPress
  End If
  If keycode = RightFlipperKey Then
    FlipperDeActivate RightFlipper, RFPress
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
Dim RoomBrightnessMtlArray: RoomBrightnessMtlArray = Array("VLM.Bake.Active","VLM.Bake.Solid", "Plastic with an image")

Sub SetRoomBrightness(lvl)
  If lvl > 1 Then lvl = 1
  If lvl < 0 Then lvl = 0

  'Lighting level
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
      objBallShadow(s).X = gBOT(s).X + (gBOT(s).X - (tablewidth / 2)) / (BallSize / AmbientMovement) + offsetX
      objBallShadow(s).Y = gBOT(s).Y + offsetY
      'objBallShadow(s).Z = gBOT(s).Z + s/1000 + 1.04 - 25

    '** No shadow if ball is off the main playfield (this may need to be adjusted per table)
    Else
      objBallShadow(s).visible = 0
    End If
  Next
End Sub


'******************************************************
'   ZBBR: BALL BRIGHTNESS
'******************************************************

Const BallBrightness =  0.7       'Ball brightness - Value between 0 and 1 (0=Dark ... 1=Bright)

' Constants for plunger lane ball darkening.
' You can make a temporary wall in the plunger lane area and use the co-ordinates from the corner control points.
Const PLOffset = 0.2      'Minimum ball brightness scale in plunger lane
Const PLLeft = 862      'X position of punger lane left
Const PLRight = 924       'X position of punger lane right
Const PLTop = 1082        'Y position of punger lane top
Const PLBottom = 1980       'Y position of punger lane bottom
Dim PLGain: PLGain = (1-PLOffset)/(PLTop-PLBottom)

Sub UpdateBallBrightness
  Dim s, b_base, b_r, b_g, b_b, d_w
  b_base = 100 * BallBrightness + 70*gilvl  ' orig was 120 and 70

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





'**********************************************************************************************************
'Flippers
'**********************************************************************************************************

Const ReflipAngle = 20

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
     If Enabled Then
    LF.Fire
    If LeftFlipper.currentangle < LeftFlipper.endangle + ReflipAngle Then
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

' Flipper collide subs
Sub LeftFlipper_Collide(parm)
  CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
  LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
  CheckLiveCatch Activeball, RightFlipper, RFCount, parm
  RightFlipperCollide parm
End Sub


' Flipper Animations
Sub LeftFlipper_Animate
  Dim a : a = LeftFlipper.CurrentAngle
  'FlipperLSh.RotZ = a

  Dim v, BP
  v = 255.0 * (122.0 -  LeftFlipper.CurrentAngle) / (122.0 -  75.0)

  For each BP in BP_LFP
    BP.Rotz = a
    BP.visible = v < 128.0
    lfshadow_up.visible = v < 128.0
    lfshadow_down.visible = v > 128.0
  Next
  For each BP in BP_LFP_UP_001
    BP.Rotz = a
    BP.visible = v >= 128.0
    lfshadow_down.visible = v < 128.0
    lfshadow_up.visible = v > 128.0
  Next
End Sub

Sub RightFlipper_Animate
  Dim a : a = RightFlipper.CurrentAngle
  'FlipperRSh.RotZ = a

  Dim v, BP
  v = 255.0 * (-122.0 -  RightFlipper.CurrentAngle) / (-122.0 +  75.0)

  For each BP in BP_RFP
    BP.Rotz = a
    BP.visible = v < 128.0
rfshadow_up.visible = v < 128.0
    rfshadow_down.visible = v > 128.0
  Next
  For each BP in BP_RFP_UP
    BP.Rotz = a
    BP.Rotz = a
    BP.visible = v >= 128.0
rfshadow_down.visible = v < 128.0
    rfshadow_up.visible = v > 128.0
  Next

End Sub

'*******************************************
'STAND UP TARGETS:
'*******************************************

Sub sw32_hit:STHit 34:End Sub
Sub sw39_hit:STHit 39:End Sub

'*******************************************
'DROP TARGETS
'*******************************************

'Lower Bank
Sub sw24_Hit : DTHit 24 : End Sub
Sub sw25_Hit : DTHit 25 : End Sub
Sub sw26_Hit : DTHit 26 : End Sub
Sub sw27_Hit : DTHit 27 : End Sub

'Upper Bank
Sub sw40_Hit : DTHit 40 : End Sub
Sub sw41_Hit : DTHit 41 : End Sub
Sub sw42_Hit : DTHit 42 : End Sub
Sub sw43_Hit : DTHit 43 : End Sub
Sub sw44_Hit : DTHit 44 : End Sub
Sub sw45_Hit : DTHit 45 : End Sub
Sub sw46_Hit : DTHit 46 : End Sub
Sub sw47_Hit : DTHit 47 : End Sub


' Reset top right drop target bank
Sub SolDTTopBankReset(enabled)
  dim xx
  if enabled then
    RandomSoundDropTargetReset BM_sw40p
    DTRaise 40
    DTRaise 41
    DTRaise 42
    DTRaise 43
    DTRaise 44
    DTRaise 45
    DTRaise 46
    DTRaise 47
    For each xx in DTShadowB: xx.visible = True: Next
  end if
End Sub

' Reset bottom left drop target bank
Sub SolDTBottomBankReset(enabled)
  dim xx
  if enabled then
    RandomSoundDropTargetReset BM_sw24p
    DTRaise 24
    DTRaise 25
    DTRaise 26
    DTRaise 27
    For each xx in DTShadows: xx.visible = True: Next
  end if
End Sub


Sub SolDT47(Enabled)
  if enabled Then
    if not DTDropped(47) Then
    DTDrop 47
    end if
  end if
end sub


Sub SolDT40(Enabled)
'if enabled AND Not DTDropped(40)Then
  if enabled Then
    if not DTDropped(40) Then
    DTDrop 40
    end if
  end if
end sub


Sub SolDT41(Enabled)
'if enabled AND Not DTDropped(40)Then
  if enabled Then
    if not DTDropped(41) Then
    DTDrop 41
    end if
  end if
end sub


Sub SolDT42(Enabled)
'if enabled AND Not DTDropped(40)Then
  if enabled Then
    if not DTDropped(42) Then
    DTDrop 42
    end if
  end if
end sub


Sub SolDT43(Enabled)
'if enabled AND Not DTDropped(40)Then
  if enabled Then
    if not DTDropped(43) Then
    DTDrop 43
    end if
  end if
end sub


Sub SolDT44(Enabled)
'if enabled AND Not DTDropped(40)Then
  if enabled Then
    if not DTDropped(44) Then
    DTDrop 44
    end if
  end if
end sub

Sub SolDT45(Enabled)
'if enabled AND Not DTDropped(40)Then
  if enabled Then
    if not DTDropped(45) Then
    DTDrop 45
    end if
  end if
end sub

Sub SolDT46(Enabled)
'if enabled AND Not DTDropped(40)Then
  if enabled Then
    if not DTDropped(46) Then
    DTDrop 46
    end if
  end if
end sub

Sub UpdateDropTargets
  dim BP, tz, rx, ry

  tz = BM_sw24p.transz
  rx = BM_sw24p.rotx
  ry = BM_sw24p.roty
  For each BP in BP_sw24p: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw25p.transz
  rx = BM_sw25p.rotx
  ry = BM_sw25p.roty
  For each BP in BP_sw25p: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw26p.transz
  rx = BM_sw26p.rotx
  ry = BM_sw26p.roty
  For each BP in BP_sw26p: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw27p.transz
  rx = BM_sw27p.rotx
  ry = BM_sw27p.roty
  For each BP in BP_sw27p: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw40p.transz
  rx = BM_sw40p.rotx
  ry = BM_sw40p.roty
  For each BP in BP_sw40p: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw41p.transz
  rx = BM_sw41p.rotx
  ry = BM_sw41p.roty
  For each BP in BP_sw41p: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw42p.transz
  rx = BM_sw42p.rotx
  ry = BM_sw42p.roty
  For each BP in BP_sw42p: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw43p.transz
  rx = BM_sw43p.rotx
  ry = BM_sw43p.roty
  For each BP in BP_sw43p: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw44p.transz
  rx = BM_sw44p.rotx
  ry = BM_sw44p.roty
  For each BP in BP_sw44p: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw45p.transz
  rx = BM_sw45p.rotx
  ry = BM_sw45p.roty
  For each BP in BP_sw45p: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw46p.transz
  rx = BM_sw46p.rotx
  ry = BM_sw46p.roty
  For each BP in BP_sw46p: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw47p.transz
  rx = BM_sw47p.rotx
  ry = BM_sw47p.roty
  For each BP in BP_sw47p: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next


End Sub


'Actions

Sub DTAction(switchid)
  Select Case switchid
    'Bottom Bank
    Case 24: DTShadows(0).visible = False : DTShadows(1).visible = False
    Case 25: DTShadows(2).visible = False : DTShadows(3).visible = False
    Case 26: DTShadows(4).visible = False : DTShadows(5).visible = False
    Case 27: DTShadows(6).visible = False : DTShadows(7).visible = False
    'Top Bank
    Case 40: DTShadowB(0).visible = False
    Case 41: DTShadowB(1).visible = False
    Case 42: DTShadowB(2).visible = False
    Case 43: DTShadowB(3).visible = False
    Case 44: DTShadowB(4).visible = False
    Case 45: DTShadowB(5).visible = False
    Case 46: DTShadowB(6).visible = False
    Case 47: DTShadowB(7).visible = False
  End Select
End Sub



'**********************************************************************************************************
'RELEASE
'**********************************************************************************************************
'Sub BallRelease_UnHit : RandomSoundBallRelease BallRelease: End Sub

'**********************************************************************************************************
'DRAIN
'**********************************************************************************************************
Sub sw16_Hit:Controller.Switch(16) = 1:RandomSoundDrain sw16:End Sub
Sub sw16_unhit:Controller.Switch(16)=0:end sub
'Sub sw11_Hit:bsTrough.addball me : RandomSoundDrain Drain: End Sub

'**********************************************************************************************************
'SCORING RUBBERS
'**********************************************************************************************************
Sub Rubber5_Hit:vpmTimer.PulseSw 23:PlaySound "sensor":End Sub
Sub Rubber8_Hit:vpmTimer.PulseSw 23:PlaySound "sensor":End Sub
Sub Rubber9_Hit:vpmTimer.PulseSw 28:PlaySound "sensor":End Sub
Sub Rubber17_Hit:vpmTimer.PulseSw 28:PlaySound "sensor":End Sub
Sub Rubber18_Hit:vpmTimer.PulseSw 28:PlaySound "sensor":End Sub

'**********************************************************************************************************
'BUMPERS
'**********************************************************************************************************
Sub Bumper1_Hit
  vpmTimer.PulseSw(35)
  RandomSoundBumperTop Bumper1
End Sub

' Bumper Aminations
Sub Bumper1_Animate
  Dim z, BP
  z = Bumper1.CurrentRingOffset
  For Each BP in BP_Bumper1_Ring_001 : BP.transz = z: Next
End Sub

Dim Bumpers : Bumpers = Array(Bumper1)

Sub AnimateBumperSkirts
  dim r, g, s, x, y, b, tz


  ' Animate Bumper switch (experimental)
  For r = 0 To 0
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
    If r = 0 Then For Each x in BP_Bumper_Socket: x.transZ = tz: Next
  Next
End Sub



'**********************************************************************************************************
'ROLLOVERS
'**********************************************************************************************************
Sub sw_TRoll1_Hit:Controller.Switch(36)=1: End Sub
Sub sw_TRoll1_unHit:Controller.Switch(36)=0:End Sub
Sub sw_TRoll2_Hit:Controller.Switch(37)=1:End Sub
Sub sw_TRoll2_unHit:Controller.Switch(37)=0:End Sub
Sub sw_TRoll3_Hit:Controller.Switch(38)=1:End Sub
Sub sw_TRoll3_unHit:Controller.Switch(38)=0:End Sub
Sub sw_LOL_Hit:Controller.Switch(21)=1:End Sub
Sub sw_LOL_unHit:Controller.Switch(21)=0:End Sub
Sub sw_LIL_Hit:Controller.Switch(22)=1: leftInlaneSpeedLimit :End Sub
Sub sw_LIL_unHit:Controller.Switch(22)=0:End Sub
Sub sw_RIL_Hit:Controller.Switch(19)=1: rightInlaneSpeedLimit :End Sub
Sub sw_RIL_unHit:Controller.Switch(19)=0:End Sub
Sub sw_ROL_Hit:Controller.Switch(20)=1:End Sub
Sub sw_ROL_unHit:Controller.Switch(20)=0:End Sub
Sub Trigger7_Hit:Controller.Switch(29)=1:End Sub
Sub Trigger7_unHit:Controller.Switch(29)=0:End Sub


'Rollover Animations
' In-Lane / Out-Lane Rollovers
Sub sw_LOL_Animate
  Dim z : z = sw_LOL.CurrentAnimOffset
  Dim BL : For Each BL in BP_sw_LOL : BL.transz = z: Next
End Sub

Sub sw_LIL_Animate
  Dim z : z = sw_LIL.CurrentAnimOffset
  Dim BL : For Each BL in BP_sw_LIL : BL.transz = z: Next
End Sub

Sub sw_ROL_Animate
  Dim z : z = sw_ROL.CurrentAnimOffset
  Dim BL : For Each BL in BP_sw_ROL : BL.transz = z: Next
End Sub

Sub sw_RIL_Animate
  Dim z : z = sw_RIL.CurrentAnimOffset
  Dim BL : For Each BL in BP_sw_RIL : BL.transz = z: Next
End Sub

Sub sw_Troll1_Animate
  Dim z : z = sw_Troll1.CurrentAnimOffset
  Dim BL : For Each BL in BP_sw_Troll1 : BL.transz = z: Next
End Sub

Sub sw_Troll2_Animate
  Dim z : z = sw_Troll2.CurrentAnimOffset
  Dim BL : For Each BL in BP_sw_Troll2 : BL.transz = z: Next
End Sub

Sub sw_Troll3_Animate
  Dim z : z = sw_Troll3.CurrentAnimOffset
  Dim BL : For Each BL in BP_sw_Troll3 : BL.transz = z: Next
End Sub

Sub Trigger7_Animate
  Dim z : z = Trigger7.CurrentAnimOffset
  Dim BL : For Each BL in BP_Trigger7 : BL.transz = z: Next
End Sub



'**********************************************************************************************************
'Spinners
'**********************************************************************************************************
Sub Spinner1_Spin:vpmTimer.PulseSw 30:SoundSpinner Spinner1: End Sub
Sub Spinner2_Spin:vpmTimer.PulseSw 31:SoundSpinner Spinner2:End Sub


' Spinner Animations
Sub Spinner1_Animate
  Dim spinangle:spinangle = Spinner1.currentangle
  Dim BL
  For Each BL in BP_Spinner1_Wire : BL.RotX = spinangle: Next
End Sub

Sub Spinner2_Animate
  Dim spinangle:spinangle = Spinner2.currentangle
  Dim BL
For Each BL in BP_Spinner1_Wire_001 : BL.RotX = spinangle: Next
  'Spinner001Shadow.size_y = abs(sin( (spinangle+180) * (2*PI/360)) * 5)
End Sub


'**********************************************************************************************************
'Gates
'**********************************************************************************************************
Sub Gate1_Hit():PlaySound "gate":End Sub
Sub Gate2_Hit():PlaySound "gate":End Sub
Sub Gate4_Hit():PlaySound "gate":End Sub


Sub Gate1_Animate
  Dim a : a = Gate1.CurrentAngle
  Dim BP : For Each BP in BP_Gate1 : BP.rotx = a: Next
End Sub

Sub Gate2_Animate
  Dim a : a = Gate2.CurrentAngle
  Dim BP : For Each BP in BP_Gate2_001 : BP.rotx = a: Next
End Sub

Sub Gate3_Animate
  Dim a : a = Gate3.CurrentAngle
  Dim BP : For Each BP in BP_Gate3_001 : BP.rotx = a: Next

End Sub

'****************************************************************
' ZSLG: Slingshots
'****************************************************************

' RStep and LStep are the variables that increment the animation
Dim RStep, LStep, RStep1, LStep1

Sub LeftSlingShot_Slingshot
  LS.VelocityCorrect(ActiveBall)
  vpmTimer.PulseSw 17
  RandomSoundSlingShotLeft BM_SlingArmL
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
  For Each BL in BP_SlingArmL : BL.transx = y: Next

  LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
  RS.VelocityCorrect(ActiveBall)
  vpmTimer.PulseSw 18
  RandomSoundSlingShotRight BM_SlingArmR
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
  For Each BL in BP_SlingArmR : BL.transx = y: Next

    RStep = RStep + 1
End Sub


Sub LeftSlingShot1_Slingshot
  LS.VelocityCorrect(ActiveBall)
  vpmTimer.PulseSw 34
  RandomSoundSlingShotLeft BM_SlingArmL1
  LStep1 = 0
  LeftSlingShot1.TimerInterval = 17
    LeftSlingShot1.TimerEnabled = 1
End Sub

Sub LeftSlingShot1_Timer
  Dim BL
  Dim x1, x2, y: x1 = True:x2 = False:y = 25
  Select Case LStep1
    Case 3:x1 = False:x2= True: y = 15
    Case 4:x1 = False:x2 = False:y = 0:LeftSlingShot1.TimerEnabled = 0
  End Select

  For Each BL in BP_L1Sling1 : BL.Visible = x1: Next
  For Each BL in BP_L1Sling2 : BL.Visible = x2: Next
  For Each BL in BP_SlingArmL1 : BL.transx = y: Next

  LStep1 = LStep1 + 1
End Sub

Sub RightSlingShot1_Slingshot
  RS.VelocityCorrect(ActiveBall)
  vpmTimer.PulseSw 33
  RandomSoundSlingShotRight BM_SlingArmR1
  RStep1 = 0
  RightSlingShot1.TimerInterval = 17
    RightSlingShot1.TimerEnabled = 1
End Sub

Sub RightSlingShot1_Timer
  Dim BL
  Dim x1, x2, y: x1 = True:x2 = False:y = 25
    Select Case RStep1
        Case 2:x1 = False:x2 = True:y = 15
        Case 3:x1 = False:x2 = False:y = 0:RightSlingShot1.TimerEnabled = 0
    End Select

  For Each BL in BP_R1Sling1 : BL.Visible = x1: Next
  For Each BL in BP_R1Sling2 : BL.Visible = x2: Next
  For Each BL in BP_SlingArmR1 : BL.transx = y: Next

    RStep1 = RStep1 + 1
End Sub



'Sub TestSlingShot_Slingshot
' TS.VelocityCorrect(ActiveBall)
'End Sub



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
    ' Comment the next line if you are not implementing Dyanmic Ball Shadows
    'If AmbientBallShadowOn = 0 Then BallShadowA(b).visible = 0
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

    ' "Static" Ball Shadows
    ' Comment the next If block, if you are not implementing the Dynamic Ball Shadows
'   If AmbientBallShadowOn = 0 Then
'     If BOT(b).Z > 30 Then
'       BallShadowA(b).height = BOT(b).z - BallSize / 4   'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping the ramp
'     Else
'       BallShadowA(b).height = 0.1
'     End If
'     BallShadowA(b).Y = BOT(b).Y + offsetY
'     BallShadowA(b).X = BOT(b).X + offsetX
'     BallShadowA(b).visible = 1
'   End If
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
' Tutorial vides by Apophis
' Audio : Adding Fleep Part 1       https://youtu.be/rG35JVHxtx4
' Audio : Adding Fleep Part 2       https://youtu.be/dk110pWMxGo
' Audio : Adding Fleep Part 3       https://youtu.be/ESXWGJZY_EI


'///////////////////////////////  SOUNDS PARAMETERS  //////////////////////////////
Dim GlobalSoundLevel, CoinSoundLevel, PlungerReleaseSoundLevel, PlungerPullSoundLevel, NudgeLeftSoundLevel
Dim NudgeRightSoundLevel, NudgeCenterSoundLevel, StartButtonSoundLevel, RollingSoundFactor

CoinSoundLevel = 0.5            'volume level; range [0, 1]
NudgeLeftSoundLevel = 1        'volume level; range [0, 1]
NudgeRightSoundLevel = 1        'volume level; range [0, 1]
NudgeCenterSoundLevel = 1        'volume level; range [0, 1]
StartButtonSoundLevel = 0.1      'volume level; range [0, 1]
PlungerReleaseSoundLevel = 0.8 '1 wjr   'volume level; range [0, 1]
PlungerPullSoundLevel = 1        'volume level; range [0, 1]
RollingSoundFactor = 1 / 5

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
SlingshotSoundLevel = 0.8           'volume level; range [0, 1]
BumperSoundFactor = 3           'volume multiplier; must not be zero
KnockerSoundLevel = 0.5              'volume level; range [0, 1]

'///////////////////////-----Ball Drops, Bumps and Collisions-----///////////////////////
Dim RubberStrongSoundFactor, RubberWeakSoundFactor, RubberFlipperSoundFactor,BallWithBallCollisionSoundFactor
Dim BallBouncePlayfieldSoftFactor, BallBouncePlayfieldHardFactor, PlasticRampDropToPlayfieldSoundLevel, WireRampDropToPlayfieldSoundLevel, DelayedBallDropOnPlayfieldSoundLevel
Dim WallImpactSoundFactor, MetalImpactSoundFactor, SubwaySoundLevel, SubwayEntrySoundLevel, ScoopEntrySoundLevel
Dim SaucerLockSoundLevel, SaucerKickSoundLevel

BallWithBallCollisionSoundFactor = 3      'volume multiplier; must not be zero
RubberStrongSoundFactor = 0.070 / 4      'volume multiplier; must not be zero
RubberWeakSoundFactor = 0.070 / 5        'volume multiplier; must not be zero
RubberFlipperSoundFactor = 0.075 / 5      'volume multiplier; must not be zero
BallBouncePlayfieldSoftFactor = 0.025      'volume multiplier; must not be zero
BallBouncePlayfieldHardFactor = 0.025      'volume multiplier; must not be zero
DelayedBallDropOnPlayfieldSoundLevel = 0.8    'volume level; range [0, 1]
WallImpactSoundFactor = 0.075          'volume multiplier; must not be zero
MetalImpactSoundFactor = 0.045 / 3
SaucerLockSoundLevel = 0.5
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

DrainSoundLevel = 0.5          'volume level; range [0, 1]
BallReleaseSoundLevel = 0.5        'volume level; range [0, 1]
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
  PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelExistingStatic(playsoundparams, aVol, tableobj)
  PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticLoop(playsoundparams, aVol, tableobj)
  PlaySound playsoundparams, - 1, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticRandomPitch(playsoundparams, aVol, randomPitch, tableobj)
  PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelActiveBall(playsoundparams, aVol)
  PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ActiveBall), 0, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLevelExistingActiveBall(playsoundparams, aVol)
  PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ActiveBall), 0, 0, 1, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLeveTimerActiveBall(playsoundparams, aVol, ballvariable)
  PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ballvariable), 0, 0, 0, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelTimerExistingActiveBall(playsoundparams, aVol, ballvariable)
  PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ballvariable), 0, 0, 1, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelRoll(playsoundparams, aVol, pitch)
  PlaySound playsoundparams, - 1, aVol * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

' Previous Positional Sound Subs

Sub PlaySoundAt(soundname, tableobj)
  PlaySound soundname, 1, 1 * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtVol(soundname, tableobj, aVol)
  PlaySound soundname, 1, aVol * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname)
  PlaySoundAt soundname, ActiveBall
End Sub

Sub PlaySoundAtBallVol (Soundname, aVol)
  PlaySound soundname, 1,aVol * VolumeDial, AudioPan(ActiveBall), 0,0,0, 1, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtBallVolM (Soundname, aVol)
  PlaySound soundname, 1,aVol * VolumeDial, AudioPan(ActiveBall), 0,0,0, 0, AudioFade(ActiveBall)
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
  'additional conditional as the DT will hit without ball velocity via the rom
  if not finalspeed = empty then
  finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
  If finalspeed > 10 Then
    RandomSoundTargetHitStrong()
    RandomSoundBallBouncePlayfieldSoft ActiveBall
  Else
    RandomSoundTargetHitWeak()
  End If
end if
End Sub

Sub Targets_Hit (idx)
  TargetBouncer ActiveBall, 1
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



' ZPHY:  GNEREAL ADVICE ON PHYSICS
'******************************************************
'
' It's advised that flipper corrections, dampeners, and general physics settings should all be updated per these
' examples as all of these improvements work together to provide a realistic physics simulation.
'
' Tutorial videos provided by Bord
' Adding nFozzy roth physics : pt1 rubber dampeners         https://youtu.be/AXX3aen06FM
' Adding nFozzy roth physics : pt2 flipper physics          https://youtu.be/VSBFuK2RCPE
' Adding nFozzy roth physics : pt3 other elements           https://youtu.be/JN8HEJapCvs
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
' | Force         | 9.5-10.5 |
' | Hit Threshold | 1.6-2    |
' | Scatter Angle | 2        |
'
' Slingshots
' | Hit Threshold      | 2    |
' | Slingshot Force    | 4-5  |
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
'   x.AddPt "Polarity", 0, 0, 0
'   x.AddPt "Polarity", 1, 0.05, - 2.7
'   x.AddPt "Polarity", 2, 0.33, - 2.7
'   x.AddPt "Polarity", 3, 0.37, - 2.7
'   x.AddPt "Polarity", 4, 0.41, - 2.7
'   x.AddPt "Polarity", 5, 0.45, - 2.7
'   x.AddPt "Polarity", 6, 0.576, - 2.7
'   x.AddPt "Polarity", 7, 0.66, - 1.8
'   x.AddPt "Polarity", 8, 0.743, - 0.5
'   x.AddPt "Polarity", 9, 0.81, - 0.5
'   x.AddPt "Polarity", 10, 0.88, 0
'
'   x.AddPt "Velocity", 0, 0, 1
'   x.AddPt "Velocity", 1, 0.16, 1.06
'   x.AddPt "Velocity", 2, 0.41, 1.05
'   x.AddPt "Velocity", 3, 0.53, 1 '0.982
'   x.AddPt "Velocity", 4, 0.702, 0.968
'   x.AddPt "Velocity", 5, 0.95,  0.968
'   x.AddPt "Velocity", 6, 1.03, 0.945
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
'   x.AddPt "Polarity", 2, 0.33, - 3.7
'   x.AddPt "Polarity", 3, 0.37, - 3.7
'   x.AddPt "Polarity", 4, 0.41, - 3.7
'   x.AddPt "Polarity", 5, 0.45, - 3.7
'   x.AddPt "Polarity", 6, 0.576,- 3.7
'   x.AddPt "Polarity", 7, 0.66, - 2.3
'   x.AddPt "Polarity", 8, 0.743, - 1.5
'   x.AddPt "Polarity", 9, 0.81, - 1
'   x.AddPt "Polarity", 10, 0.88, 0
'
'   x.AddPt "Velocity", 0, 0, 1
'   x.AddPt "Velocity", 1, 0.16, 1.06
'   x.AddPt "Velocity", 2, 0.41, 1.05
'   x.AddPt "Velocity", 3, 0.53, 1 '0.982
'   x.AddPt "Velocity", 4, 0.702, 0.968
'   x.AddPt "Velocity", 5, 0.95,  0.968
'   x.AddPt "Velocity", 6, 1.03, 0.945
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
'   x.AddPt "Polarity", 2, 0.4, - 5
'   x.AddPt "Polarity", 3, 0.6, - 4.5
'   x.AddPt "Polarity", 4, 0.65, - 4.0
'   x.AddPt "Polarity", 5, 0.7, - 3.5
'   x.AddPt "Polarity", 6, 0.75, - 3.0
'   x.AddPt "Polarity", 7, 0.8, - 2.5
'   x.AddPt "Polarity", 8, 0.85, - 2.0
'   x.AddPt "Polarity", 9, 0.9, - 1.5
'   x.AddPt "Polarity", 10, 0.95, - 1.0
'   x.AddPt "Polarity", 11, 1, - 0.5
'   x.AddPt "Polarity", 12, 1.1, 0
'   x.AddPt "Polarity", 13, 1.3, 0
'
'   x.AddPt "Velocity", 0, 0, 1
'   x.AddPt "Velocity", 1, 0.16, 1.06
'   x.AddPt "Velocity", 2, 0.41, 1.05
'   x.AddPt "Velocity", 3, 0.53, 1 '0.982
'   x.AddPt "Velocity", 4, 0.702, 0.968
'   x.AddPt "Velocity", 5, 0.95,  0.968
'   x.AddPt "Velocity", 6, 1.03,  0.945
' Next
'
' ' SetObjects arguments: 1: name of object 2: flipper object: 3: Trigger object around flipper
' LF.SetObjects "LF", LeftFlipper, TriggerLF
' RF.SetObjects "RF", RightFlipper, TriggerRF
'End Sub

'*******************************************
' Late 70's to early 80's

Sub InitPolarity()
   dim x, a : a = Array(LF, RF)
  for each x in a
    x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
    x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
    x.enabled = True
    x.TimeDelay = 80
    x.DebugOn=False ' prints some info in debugger

    x.AddPt "Polarity", 0, 0, 0
    x.AddPt "Polarity", 1, 0.05, - 2.7
    x.AddPt "Polarity", 2, 0.33, - 2.7
    x.AddPt "Polarity", 3, 0.37, - 2.7
    x.AddPt "Polarity", 4, 0.41, - 2.7
    x.AddPt "Polarity", 5, 0.45, - 2.7
    x.AddPt "Polarity", 6, 0.576, - 2.7
    x.AddPt "Polarity", 7, 0.66, - 1.8
    x.AddPt "Polarity", 8, 0.743, - 0.5
    x.AddPt "Polarity", 9, 0.81, - 0.5
    x.AddPt "Polarity", 10, 0.88, 0

    x.AddPt "Velocity", 0, 0, 1
    x.AddPt "Velocity", 1, 0.16, 1.06
    x.AddPt "Velocity", 2, 0.41, 1.05
    x.AddPt "Velocity", 3, 0.53, 1 '0.982
    x.AddPt "Velocity", 4, 0.702, 0.968
    x.AddPt "Velocity", 5, 0.95,  0.968
    x.AddPt "Velocity", 6, 1.03, 0.945
  Next

  ' SetObjects arguments: 1: name of object 2: flipper object: 3: Trigger object around flipper
    LF.SetObjects "LF", LeftFlipper, TriggerLF
    RF.SetObjects "RF", RightFlipper, TriggerRF
End Sub

'' Flipper trigger hit subs
'Sub TriggerLF_Hit()
' LF.Addball activeball
'End Sub
'Sub TriggerLF_UnHit()
' LF.PolarityCorrect activeball
'End Sub
'Sub TriggerRF_Hit()
' RF.Addball activeball
'End Sub
'Sub TriggerRF_UnHit()
' RF.PolarityCorrect activeball
'End Sub

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
  Private Flipper, FlipperStart, FlipperEnd, FlipperEndY, LR, PartialFlipCoef
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
      If Not IsEmpty(balls(x) ) Then
        pos = pSlope(Balls(x).x, FlipperStart, 0, FlipperEnd, 1)
      End If
    Next
  End Property

  Public Sub ProcessBalls() 'save data of balls in flipper range
    FlipAt = GameTime
    Dim x
    For x = 0 To UBound(balls)
      If Not IsEmpty(balls(x) ) Then
        balldata(x).Data = balls(x)
      End If
    Next
    PartialFlipCoef = ((Flipper.StartAngle - Flipper.CurrentAngle) / (Flipper.StartAngle - Flipper.EndAngle))
    PartialFlipCoef = abs(PartialFlipCoef-1)
  End Sub
  'Timer shutoff for polaritycorrect
  Private Function FlipperOn()
    If GameTime < FlipAt+TimeDelay Then
      FlipperOn = True
    End If
  End Function

  Public Sub PolarityCorrect(aBall)
    If FlipperOn() Then
      Dim tmp, BallPos, x, IDX, Ycoef
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
          If ballpos > 0.65 Then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)                'find safety coefficient 'ycoef' data
        End If
      Next

      If BallPos = 0 Then 'no ball data meaning the ball is entering and exiting pretty close to the same position, use current values.
        BallPos = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
        If ballpos > 0.65 Then  Ycoef = LinearEnvelope(aBall.Y, YcoefIn, YcoefOut)                        'find safety coefficient 'ycoef' data
      End If

      'Velocity correction
      If Not IsEmpty(VelocityIn(0) ) Then
        Dim VelCoef
        VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)

        If partialflipcoef < 1 Then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)

        If Enabled Then aBall.Velx = aBall.Velx*VelCoef
        If Enabled Then aBall.Vely = aBall.Vely*VelCoef
      End If

      'Polarity Correction (optional now)
      If Not IsEmpty(PolarityIn(0) ) Then
        Dim AddX
        AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR

        If Enabled Then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef)
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
Const EOSTnew = 1 'EM's to late 80's
'Const EOSTnew = 0.8 '90's and later
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

Const LiveDistanceMin = 30  'minimum distance In vp units from flipper base live catch dampening will occur
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

' Note, cor.update must be called in a 10 ms timer. The example table uses the GameTimer for this purpose, but sometimes a dedicated timer call RDampen is used.
'
'Sub RDampen_Timer
' Cor.Update
'End Sub

'******************************************************
'****  END PHYSICS DAMPENERS
'******************************************************



'******************************************************
'   ZBOU: VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************


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

Dim LS1
Set LS1 = New SlingshotCorrection
Dim RS1
Set RS1 = New SlingshotCorrection

InitSlingCorrection

Sub InitSlingCorrection
  LS.Object = LeftSlingshot
  LS.EndPoint1 = EndPoint1LS
  LS.EndPoint2 = EndPoint2LS

  RS.Object = RightSlingshot
  RS.EndPoint1 = EndPoint1RS
  RS.EndPoint2 = EndPoint2RS


  LS1.Object = LeftSlingshot1
  LS1.EndPoint1 = EndPoint1LS1
  LS1.EndPoint2 = EndPoint2LS1

  RS1.Object = RightSlingshot1
  RS1.EndPoint1 = EndPoint1RS1
  RS1.EndPoint2 = EndPoint2RS1

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

' The following sub are needed, however they may exist somewhere else in the script. Uncomment below if needed
'Dim PI: PI = 4*Atn(1)
'Function dSin(degrees)
' dsin = sin(degrees * Pi/180)
'End Function
'Function dCos(degrees)
' dcos = cos(degrees * Pi/180)
'End Function

Function RotPoint(x,y,angle)
  dim rx, ry
  rx = x*dCos(angle) - y*dSin(angle)
  ry = x*dSin(angle) + y*dCos(angle)
  RotPoint = Array(rx,ry)
End Function

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



'******************************************************
'   ZRDT:  DROP TARGETS by Rothbauerw
'******************************************************
' This solution improves the physics for drop targets to create more realistic behavior. It allows the ball
' to move through the target enabling the ability to score more than one target with a well placed shot.
' It also handles full drop target animation, including deflection on hit and a slight lift when the drop
' targets raise, switch handling, bricking, and popping the ball up if it's over the drop target when it raises.
'
'Add a Timer named DTAnim to editor to handle drop & standup target animations, or run them off an always-on 10ms timer (GameTimer)
'DTAnim.interval = 10
'DTAnim.enabled = True

'Sub DTAnim_Timer
' DoDTAnim
' DoSTAnim
'End Sub

' For each drop target, we'll use two wall objects for physics calculations and one primitive for visuals and
' animation. We will not use target objects.  Place your drop target primitive the same as you would a VP drop target.
' The primitive should have it's pivot point centered on the x and y axis and at or just below the playfield
' level on the z axis. Orientation needs to be set using Rotz and bending deflection using Rotx. You'll find a hooded
' target mesh in this table's example. It uses the same texture map as the VP drop targets.

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
Dim DT24, DT25, DT26, DT27, DT40, DT41, DT42, DT43, DT44, DT45, DT46, DT47

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


Set DT24 = (new DropTarget)(sw24, sw24a, BM_sw24p,24, 0, false)
Set DT25 = (new DropTarget)(sw25, sw25a, BM_sw25p ,25, 0, false)
Set DT26 = (new DropTarget)(sw26, sw26a, BM_sw26p ,26, 0, false)
Set DT27 = (new DropTarget)(sw27, sw27a, BM_sw27p ,27, 0, false)

Set DT40 = (new DropTarget)(sw40, sw40a, BM_sw40p,40, 0, false)
Set DT41 = (new DropTarget)(sw41, sw41a, BM_sw41p ,41, 0, false)
Set DT42 = (new DropTarget)(sw42, sw42a, BM_sw42p ,42, 0, false)
Set DT43 = (new DropTarget)(sw43, sw43a, BM_sw43p ,43, 0, false)
Set DT44 = (new DropTarget)(sw44, sw44a, BM_sw44p ,44, 0, false)
Set DT45 = (new DropTarget)(sw45, sw45a, BM_sw45p,45, 0, false)
Set DT46 = (new DropTarget)(sw46, sw46a, BM_sw46p, 46, 0, false)
Set DT47 = (new DropTarget)(sw47, sw47a, BM_sw47p, 47, 0, false)


Dim DTArray
DTArray = Array(DT24, DT25, DT26, DT27, DT40, DT41, DT42, DT43, DT44, DT45, DT46, DT47)


'Dim DTArray0, DTArray1, DTArray2, DTArray3, DTArray4, DTArray5
'ReDim DTArray0(UBound(DTArray)), DTArray1(UBound(DTArray)), DTArray2(UBound(DTArray)), DTArray3(UBound(DTArray)), DTArray4(UBound(DTArray)), DTArray5(UBound(DTArray))
'
'Dim DTIdx
'For DTIdx = 0 To UBound(DTArray)
' Set DTArray0(DTIdx) = DTArray(DTIdx)(0)
' Set DTArray1(DTIdx) = DTArray(DTIdx)(1)
' Set DTArray2(DTIdx) = DTArray(DTIdx)(2)
' DTArray3(DTIdx) = DTArray(DTIdx)(3)
' DTArray4(DTIdx) = DTArray(DTIdx)(4)
' DTArray5(DTIdx) = DTArray(DTIdx)(5)
'Next


'Configure the behavior of Drop Targets.
Const DTDropSpeed = 90 'in milliseconds
Const DTDropUpSpeed = 40 'in milliseconds
Const DTDropUnits = 44 'VP units primitive drops so top of at or below the playfield
Const DTDropUpUnits = 10 'VP units primitive raises above the up position on drops up
Const DTMaxBend = 4 'max degrees primitive rotates when hit
Const DTDropDelay = 20 'time in milliseconds before target drops (due to friction/impact of the ball)
Const DTRaiseDelay = 40 'time in milliseconds before target drops back to normal up position after the solenoid fires to raise the target
Const DTBrickVel = 30 'velocity at which the target will brick, set to '0' to disable brick
Const DTEnableBrick = 0 'Set to 0 to disable bricking, 1 to enable bricking
Const DTMass = 0.5 'Mass of the Drop Target (between 0 and 1), higher values provide more resistance

'******************************************************
'  DROP TARGETS FUNCTIONS
'******************************************************

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
    If prim.z >=0 then
      prim.rotx = DTMaxBend * Cos(rangle)
    Else
      prim.rotx = DTMaxBend * Cos(rangle) '+ 11
    End If
    prim.roty = DTMaxBend * Sin(rangle)
    DTAnimate = animate
    Exit Function
  ElseIf (animate = 1 Or animate = 4) And animtime > DTDropDelay Then
    primary.collidable = 0
    If animate = 1 Then secondary.collidable = 1 Else secondary.collidable = 0
    If prim.z >=0 then
      prim.rotx = DTMaxBend * Cos(rangle)
    Else
      prim.rotx = DTMaxBend * Cos(rangle) '+ 11
    End If
    prim.roty = DTMaxBend * Sin(rangle)
    animate = 2
    SoundDropTargetDrop prim
  End If

  If animate = 2 Then
    transz = (animtime - DTDropDelay) / DTDropSpeed * DTDropUnits *  - 1
    If prim.transz >  - DTDropUnits  Then
      prim.transz = transz
    End If

    If prim.z >=0 then
      prim.rotx = DTMaxBend * Cos(rangle) / 2
    Else
      prim.rotx = DTMaxBend * Cos(rangle) / 2 '+ 11
    End If
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
    If prim.z >=0 then
      prim.rotx = DTMaxBend * Cos(rangle)
    Else
      prim.rotx = DTMaxBend * Cos(rangle)' + 11
    End If
    prim.roty = DTMaxBend * Sin(rangle)
  ElseIf animate = 3 And animtime > DTDropDelay Then
    primary.collidable = 1
    secondary.collidable = 0
    If prim.z >=0 then
      prim.rotx = 0
    Else
      prim.rotx = 11
    End If
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
      If prim.z >=0 then
        prim.rotx = 0
      Else
        prim.rotx = 0
      End If
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

'Used for drop targets
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

'Function RotPoint(x,y,angle)
  'Dim rx, ry
  'rx = x * dCos(angle) - y * dSin(angle)
  'ry = x * dSin(angle) + y * dCos(angle)
  'RotPoint = Array(rx,ry)
'End Function

'******************************************************
'****  END DROP TARGETS
'******************************************************


'******************************************************
'****  INLANE SPEED LIMIT
'******************************************************

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

'******************************************************
'****  END INLANE SPEED LIMIT
'******************************************************

'******************************************************
' ZRST: STAND-UP TARGETS by Rothbauerw
'******************************************************



'Define a variable for each stand-up target
Dim ST32, ST39

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

ST32 = Array(sw32, BM_psw32, 32, 0)
ST39 = Array(sw39, BM_psw39, 39, 0)

'Set ST32 = (new StandupTarget)(sw32, BM_psw32, 32, 0)
'Set ST39 = (new StandupTarget)(sw39, BM_psw39, 39, 0)

'Add all the Stand-up Target Arrays to Stand-up Target Animation Array
'   STAnimationArray = Array(ST1, ST2, ....)
Dim STArray
STArray = Array(ST32, ST39)

Dim STArray0, STArray1, STArray2, STArray3
ReDim STArray0(UBound(STArray)), STArray1(UBound(STArray)), STArray2(UBound(STArray)), STArray3(UBound(STArray))

Dim STIdx
For STIdx = 0 To UBound(STArray)
  Set STArray0(STIdx) = STArray(STIdx)(0)
  Set STArray1(STIdx) = STArray(STIdx)(1)
  STArray2(STIdx) = STArray(STIdx)(2)
  STArray3(STIdx) = STArray(STIdx)(3)
Next

'Configure the behavior of Stand-up Targets
Const STAnimStep = 1.5  'vpunits per animation step (control return to Start)
Const STMaxOffset = 9   'max vp units target moves when hit

Const STMass = 0.4    'Mass of the Stand-up Target (between 0 and 1), higher values provide more resistance



'******************************************************
'       STAND-UP TARGETS FUNCTIONS
'******************************************************


Sub STHit(switch)
  Dim i
  i = STArrayID(switch)

  PlayTargetSound
  STArray3(i) = STCheckHit(ActiveBall,STArray0(i))

  If STArray3(i) <> 0 Then
    DTBallPhysics ActiveBall, STArray0(i).orientation, STMass
  End If
  DoSTAnim
End Sub

Function STArrayID(switch)
  Dim i
  For i = 0 To UBound(STArray)
    If STArray2(i) = switch Then
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
    STArray3(i) = STAnimate(STArray0(i),STArray1(i),STArray2(i),STArray3(i))
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

End Sub


Sub UpdateStandupTargets
  dim BP, ty
    ty = BM_psw32.transy
  For each BP in BP_psw32 : BP.transy = ty: Next

    ty = BM_psw39.transy
  For each BP in BP_psw39 : BP.transy = ty: Next

End Sub

'******************************************************
'****   END STAND-UP TARGETS
'******************************************************



''**********************************************************************************************************
''Digital Display
''**********************************************************************************************************
'


Dim Digits(34)
Digits(0)=Array(D16a,D17a,D18a,D19a,D20a,D21a,D22a)
Digits(1)=Array(D23a,D24a,D25a,D26a,D27a,D28a,D29a)
Digits(2)=Array(D30a,D31a,D32a,D33a,D34a,D35a,D36a)
Digits(3)=Array(D38a,D39a,D40a,D41a,D42a,D43a,D44a)
Digits(4)=Array(D45a,D46a,D47a,D48a,D49a,D50a,D51a)
Digits(5)=Array(D52a,D53a,D54a,D55a,D56a,D57a,D58a)

Digits(6)=Array(D16a1,D17a1,D18a1,D19a1,D20a1,D21a1,D22a1)
Digits(7)=Array(D23a1,D24a1,D25a1,D26a1,D27a1,D28a1,D29a1)
Digits(8)=Array(D30a1,D31a1,D32a1,D33a1,D34a1,D35a1,D36a1)
Digits(9)=Array(D38a1,D39a1,D40a1,D41a1,D42a1,D43a1,D44a1)
Digits(10)=Array(D45a1,D46a1,D47a1,D48a1,D49a1,D50a1,D51a1)
Digits(11)=Array(D52a1,D53a1,D54a1,D55a1,D56a1,D57a1,D58a1)

Digits(12)=Array(D16a2,D17a2,D18a2,D19a2,D20a2,D21a2,D22a2)
Digits(13)=Array(D23a2,D24a2,D25a2,D26a2,D27a2,D28a2,D29a2)
Digits(14)=Array(D30a2,D31a2,D32a2,D33a2,D34a2,D35a2,D36a2)
Digits(15)=Array(D38a2,D39a2,D40a2,D41a2,D42a2,D43a2,D44a2)
Digits(16)=Array(D45a2,D46a2,D47a2,D48a2,D49a2,D50a2,D51a2)
Digits(17)=Array(D52a2,D53a2,D54a2,D55a2,D56a2,D57a2,D58a2)

Digits(18)=Array(D16a3,D17a3,D18a3,D19a3,D20a3,D21a3,D22a3)
Digits(19)=Array(D23a3,D24a3,D25a3,D26a3,D27a3,D28a3,D29a3)
Digits(20)=Array(D30a3,D31a3,D32a3,D33a3,D34a3,D35a3,D36a3)
Digits(21)=Array(D38a3,D39a3,D40a3,D41a3,D42a3,D43a3,D44a3)
Digits(22)=Array(D45a3,D46a3,D47a3,D48a3,D49a3,D50a3,D51a3)
Digits(23)=Array(D52a3,D53a3,D54a3,D55a3,D56a3,D57a3,D58a3)

Digits(24)=Array(D16a4,D17a4,D18a4,D19a4,D20a4,D21a4,D22a4)
Digits(25)=Array(D23a4,D24a4,D25a4,D26a4,D27a4,D28a4,D29a4)
Digits(26)=Array(D30a4,D31a4,D32a4,D33a4,D34a4,D35a4,D36a4)
Digits(27)=Array(D38a4,D39a4,D40a4,D41a4,D42a4,D43a4,D44a4)

Digits(28)=Array(D45a4,D46a4,D47a4,D48a4,D49a4,D50a4,D51a4)
Digits(29)=Array(D52a4,D53a4,D54a4,D55a4,D56a4,D57a4,D58a4)

Digits(30)=Array(D16a5,D17a5,D18a5,D19a5,D20a5,D21a5,D22a5)
Digits(31)=Array(D23a5,D24a5,D25a5,D26a5,D27a5,D28a5,D29a5)
Digits(32)=Array(D30a5,D31a5,D32a5,D33a5,D34a5,D35a5,D36a5)
Digits(33)=Array(D38a5,D39a5,D40a5,D41a5,D42a5,D43a5,D44a5)

Sub DisplayTimer_Timer
If VR = False And DesktopMode Then
  Dim ChgLED,ii,num,chg,stat,obj
  ChgLed = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
If Not IsEmpty(ChgLED) Then
    For ii = 0 To UBound(chgLED)
      num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
      if (num < 34) then
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



'*******************************************
' VR Room / VR Cabinet
'*******************************************

DIM VRThings, VLM_Item
If VR Then
SetBackglass()
bgdark.visible=1
bgLit.visible=1
  for each VRThings in VRCabinet:VRThings.visible=1:Next
  for each VRThings in VRBackglassDigits:VRThings.visible=1:Next
  for each VRThings in VRBackglass:VRThings.visible=1:Next

  For Each VLM_Item in BGVR_All
  VLM_Item.visible = 1
  Next




else
  for each VRThings in VRCabinet:VRThings.visible=0:Next
  'for each VRThings in VRRoomMin:VRThings.visible=0:Next
  for each VRThings in VRBackglassDigits:VRThings.visible=0:Next
  for each VRThings in VRBackglass:VRThings.visible=0:Next
  bgdark.visible=0
  bgLit.visible=0
For Each VLM_Item in BGVR_All
  VLM_Item.visible = 0
  Next


SetupVRRoom

end if





Sub SetupVRRoom

  IF VR Then

  Dim VRThings
  If VRRoom = 1 Then
    for each VRThings in VR_Moon
    VRThings.visible = 1
    next

    for each VRThings in VRRoomMin
    VRThings.visible = 0
    next

End If

  If VRRoom = 0 Then
    for each VRThings in VRRoomMin
    VRThings.visible = 1
    next

    for each VRThings in VR_Moon
    VRThings.visible = 0
    next



  End If
End If





End Sub


Dim VRBGObj

Sub SetBackglass()


  For Each VRBGObj In VRBackglass
    VRBGObj.x = VRBGObj.x
    VRBGObj.height = - VRBGObj.y + 148
    VRBGObj.y = 0 'adjusts the distance from the backglass towards the user
    VRBGObj.Rotx = -90
  Next

  For Each VRBGObj In VRBackglassDigits
    VRBGObj.x = VRBGObj.x
    VRBGObj.height = - VRBGObj.y + 130
    VRBGObj.y = 0 'adjusts the distance from the backglass towards the user
    VRBGObj.Rotx = -90
  Next



End Sub

Dim xoff,yoff1, yoff2, yoff3, yoff4, yoff5,zoff,xrot,zscale, xcen,ycen

'******************* VR Plunger **********************

Sub TimerVRPlunger_Timer
  If VR_Primary_plunger.Y < 2216 then
       VR_Primary_plunger.Y = VR_Primary_plunger.Y + 3.3
  End If
End Sub

Sub TimerVRPlunger2_Timer
  VR_Primary_plunger.Y = 2128.745 + (3.3* Plunger.Position) -20
End Sub

Sub VRNebulaTimer_Timer
If VRRoom = 1 Then
Dim VRAsteroid

if VRMotion <> 0 then
VRNebulaTimer.Enabled = 1

For Each VRAsteroid In VR_Asteroids
    VRAsteroid.ObjRotZ = VRAsteroid.ObjRotZ-0.015
    VRAsteroid.ObjRotX = VRAsteroid.ObjRotX-0.045
Next

'VR_Nebula.ObjRotZ = VR_Nebula.ObjRotZ+0.015
'VR_Nebula.ObjRotX = VR_Nebula.ObjRotX+0.045

if VRMotion = 1 Then
VR_Nebula.ObjRotZ = VR_Nebula.ObjRotZ+0.003
VR_Nebula.ObjRotX = VR_Nebula.ObjRotX+0.003

VR_Nebula001.ObjRotZ = VR_Nebula001.ObjRotZ+0.005
VR_Nebula001.ObjRotX = VR_Nebula001.ObjRotX+0.005

elseif VRMotion = 2 Then
VR_Nebula.ObjRotZ = VR_Nebula.ObjRotZ+0.008
VR_Nebula.ObjRotX = VR_Nebula.ObjRotX+0.012

VR_Nebula001.ObjRotZ = VR_Nebula001.ObjRotZ+0.010
VR_Nebula001.ObjRotX = VR_Nebula001.ObjRotX+0.014

elseif VRMotion = 3 Then
VR_Nebula.ObjRotZ = VR_Nebula.ObjRotZ+0.018
VR_Nebula.ObjRotX = VR_Nebula.ObjRotX+0.018

VR_Nebula001.ObjRotZ = VR_Nebula001.ObjRotZ+0.030
VR_Nebula001.ObjRotX = VR_Nebula001.ObjRotX+0.030
end if

End If


End If
End Sub




Dim VRDigits(34)
'P1
VRDigits(0) = Array(LED1x0,LED1x1,LED1x2,LED1x3,LED1x4,LED1x5,LED1x6)
VRDigits(1) = Array(LED2x0,LED2x1,LED2x2,LED2x3,LED2x4,LED2x5,LED2x6)
VRDigits(2) = Array(LED3x0,LED3x1,LED3x2,LED3x3,LED3x4,LED3x5,LED3x6)
VRDigits(3) = Array(LED4x0,LED4x1,LED4x2,LED4x3,LED4x4,LED4x5,LED4x6)
VRDigits(4) = Array(LED5x0,LED5x1,LED5x2,LED5x3,LED5x4,LED5x5,LED5x6)
VRDigits(5) = Array(LED6x0,LED6x1,LED6x2,LED6x3,LED6x4,LED6x5,LED6x6)

'P2
VRDigits(6) = Array(LED8x0,LED8x1,LED8x2,LED8x3,LED8x4,LED8x5,LED8x6)
VRDigits(7) = Array(LED9x0,LED9x1,LED9x2,LED9x3,LED9x4,LED9x5,LED9x6)
VRDigits(8) = Array(LED10x0,LED10x1,LED10x2,LED10x3,LED10x4,LED10x5,LED10x6)
VRDigits(9) = Array(LED11x0,LED11x1,LED11x2,LED11x3,LED11x4,LED11x5,LED11x6)
VRDigits(10) = Array(LED12x0,LED12x1,LED12x2,LED12x3,LED12x4,LED12x5,LED12x6)
VRDigits(11) = Array(LED13x0,LED13x1,LED13x2,LED13x3,LED13x4,LED13x5,LED13x6)

'P3
VRDigits(12) = Array(LED1x000,LED1x001,LED1x002,LED1x003,LED1x004,LED1x005,LED1x006)
VRDigits(13) = Array(LED1x100,LED1x101,LED1x102,LED1x103,LED1x104,LED1x105,LED1x106)
VRDigits(14) = Array(LED1x200,LED1x201,LED1x202,LED1x203,LED1x204,LED1x205,LED1x206)
VRDigits(15) = Array(LED1x300,LED1x301,LED1x302,LED1x303,LED1x304,LED1x305,LED1x306)
VRDigits(16) = Array(LED1x400,LED1x401,LED1x402,LED1x403,LED1x404,LED1x405,LED1x406)
VRDigits(17) = Array(LED1x500,LED1x501,LED1x502,LED1x503,LED1x504,LED1x505,LED1x506)

'P4
VRDigits(18) = Array(LED2x000,LED2x001,LED2x002,LED2x003,LED2x004,LED2x005,LED2x006)
VRDigits(19) = Array(LED2x100,LED2x101,LED2x102,LED2x103,LED2x104,LED2x105,LED2x106)
VRDigits(20) = Array(LED2x200,LED2x201,LED2x202,LED2x203,LED2x204,LED2x205,LED2x206)
VRDigits(21) = Array(LED2x300,LED2x301,LED2x302,LED2x303,LED2x304,LED2x305,LED2x306)
VRDigits(22) = Array(LED2x400,LED2x401,LED2x402,LED2x403,LED2x404,LED2x405,LED2x406)
VRDigits(23) = Array(LED2x500,LED2x501,LED2x502,LED2x503,LED2x504,LED2x505,LED2x506)


'CREDITS
VRDigits(24) = Array(ledCr1x0,ledCr1x1,ledCr1x2,ledCr1x3,ledCr1x4,ledCr1x5,ledCr1x6)
VRDigits(25) = Array(ledCr2x0,ledCr2x1,ledCr2x2,ledCr2x3,ledCr2x4,ledCr2x5,ledCr2x6)
'MATCH / BALLS
VRDigits(26) = Array(LEDcx500,LEDcx501,LEDcx502,LEDcx503,LEDcx504,LEDcx505,LEDcx506)
VRDigits(27) = Array(LEDdx600,LEDdx601,LEDdx602,LEDdx603,LEDdx604,LEDdx605,LEDdx606)

'HS
VRDigits(28) = Array(LED3x000,LED3x001,LED3x002,LED3x003,LED3x004,LED3x005,LED3x006)
VRDigits(29) = Array(LED3x100,LED3x101,LED3x102,LED3x103,LED3x104,LED3x105,LED3x106)
VRDigits(30) = Array(LED3x200,LED3x201,LED3x202,LED3x203,LED3x204,LED3x205,LED3x206)
VRDigits(31) = Array(LED3x300,LED3x301,LED3x302,LED3x303,LED3x304,LED3x305,LED3x306)
VRDigits(32) = Array(LED3x400,LED3x401,LED3x402,LED3x403,LED3x404,LED3x405,LED3x406)
VRDigits(33) = Array(LED3x500,LED3x501,LED3x502,LED3x503,LED3x504,LED3x505,LED3x506)

dim DisplayColor
DisplayColor =  RGB(255,40,1)

Sub VRDisplayTimer_timer
  If VR Then
    Dim ChgLED, ii, jj, num, chg, stat, obj, b, x
    ChgLED=Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
       For ii=0 To UBound(chgLED)
        num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)
if (num < 34) then
          For Each obj In VRDigits(num)

  '                   If chg And 1 Then obj.visible=stat And 1    'if you use the object color for off; turn the display object visible to not visible on the playfield, and uncomment this line out.
             If chg And 1 Then FadeDisplay obj, stat And 1
             chg=chg\2 : stat=stat\2
          Next
else
end if
      Next

    End If
  UpdateVRLamps
  End If
 End Sub

Sub FadeDisplay(object, onoff)
  If OnOff = 1 Then
    object.color = DisplayColor
    Object.Opacity = 1000
  Else
    Object.Color = RGB(1,1,1)
    Object.Opacity = 50
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

InitDigits


Sub UpdateVRLamps
  ' need to add image based backglass lamps for VR
IF Controller.Lamp(1) = 0 Then: vrBG_BonusBall.visible = 0 : Else : vrBG_BonusBall.visible = 1 'Bonus Ball
IF Controller.Lamp(19) = 0 Then: vrBG_SuperBonus.visible = 0 : Else : vrBG_SuperBonus.visible = 1 'Superbonus
IF Controller.Lamp(8) = 0 Then: vrBG_Match.visible = 0 : Else : vrBG_Match.visible = 1 'Match
IF Controller.Lamp(7) = 0 Then: vrBG_GameOver.visible = 0 : Else : vrBG_GameOver.visible = 1 'Game Over
IF Controller.Lamp(51) = 0 Then: vrBG_Tilt.visible = 0 : Else : vrBG_Tilt.visible = 1 'Tilt
IF Controller.Lamp(23) = 0 Then: vrBG_BallsToPlay.visible = 0 : Else : vrBG_BallsToPlay.visible = 1 'Ball in Play
IF Controller.Lamp(47) = 0 Then: vrBG_P1sm.visible = 0 : Else : vrBG_P1sm.visible = 1 'p1 Bottom Circle
IF Controller.Lamp(48) = 0 Then: vrBG_P2sm.visible = 0 : Else : vrBG_P2sm.visible = 1 'p2 Bottom Circle
IF Controller.Lamp(49) = 0 Then: vrBG_P3sm.visible = 0 : Else : vrBG_P3sm.visible = 1 'p3 Bottom Circle
IF Controller.Lamp(50) = 0 Then: vrBG_P4sm.visible = 0 : Else : vrBG_P4sm.visible = 1 'p4 Bottom Circle

IF Controller.Lamp(60) = 0 Then: vrBG_P1.visible = 0 : Else : vrBG_P1.visible = 1 'p1
IF Controller.Lamp(61) = 0 Then: vrBG_P2.visible = 0 : Else : vrBG_P2.visible = 1 'p2
IF Controller.Lamp(62) = 0 Then: vrBG_P3.visible = 0 : Else : vrBG_P3.visible = 1 'p3
IF Controller.Lamp(63) = 0 Then: vrBG_P4.visible = 0 : Else : vrBG_P4.visible = 1 'p4

End Sub





Sub UpdateDTLamps
  ' need to add image based backglass lamps for VR
IF Controller.Lamp(60) = 0 Then: l60bg.visible = 0 : Else : l60bg.visible = 1 'Bonus Ball

End Sub



' Culture neutral string to double conversion (handles situation where you don't know how the string was written)
Function CNCDbl(str)
    Dim strt, Sep, i
    If IsNumeric(str) Then
        CNCDbl = CDbl(str)
    Else
        Sep = Mid(CStr(0.5), 2, 1)
        Select Case Sep
        Case "."
            i = InStr(1, str, ",")
        Case ","
            i = InStr(1, str, ".")
        End Select
        If i = 0 Then
            CNCDbl = Empty
        Else
            strt = Mid(str, 1, i - 1) & Sep & Mid(str, i + 1)
            If IsNumeric(strt) Then
                CNCDbl = CDbl(strt)
            Else
                CNCDbl = Empty
            End If
        End If
    End If
End Function




' ** NEW for VP 10.8 **

'v 0 - 1.9:
'- playfield and plastics cleanup, recolor and redraw - (no more baked-in shadows or reflections)
'- Table rebuilt in Blender with a few remodelled assets but mostly using VLM and Niwak's parts library
'- A few remodels of old vpx wall-created objects
'- Gameplay fixes:-
'     a. correct top bank targets should now drop under the following conditions:
'       1. LIT bottom left inlane rollover will drop top left target
'       2. LIT bottom right inlane rollover will drop top right target
'     b. Multiplayer: Dropped targets in top bank should now work correctly, remembering each player's progress
' Removed old LUT and options and replaced with VPW style DMD options for:
'     1. brightness
'     2. Saturation/LUT
'     3. Bumper strength
'- general physics and gameplay tweaks

'***************************************

'2.0   added VR Room 2 work-in-progress with animated sky sphere and asteroids, added cord and outlet stolen from HH
'2.1.1 updated slings, sling animation, VR room 2 options for motion, uder xmultplier plastics blender adjusts
'2.1.2 updated apron, credit light, added 3d instruction cards, changed lower special stand up to square
'2.1.3 fully rebuilt apron. added flex dmd options for Italian or English instruction cards. Updated centre apron card. Thx RajoJoey for ENG instructions
'2.1.6 added asteroid platform and adjustments to VR room 2, added 3 flex dmd options for center apron card.
'2.1.7 updated VR room 2 nebula for parallax
'2.1.8 added playfield reflections, fixed VR DMD
'2.1.9 rebaked - updated bumper internal bulb pos. upated bumper light to no longer catch on bottom left lane guide plastics. fixed type and position on right apron card for ENG. instructions.
'2.2.0 rebaked - lowered bumper, gates and spinner pasts/bracket very slightly as they were a little too high
'2.2.1 micro adjustments to collidables to match updated parts primitive
'2.2.2 small adjust plastics refractions, animated bumper skirt
'2.2.3 testing
'2.2.4 tidied up bumper bake, added multple GI color options to the FlexDMD Menu, cleaned up a couple playfeld bits
'2.2.5 Added physical trough, changed all old BOT calls to gBOT

'RC0.0.9: further plastics image cleanup. Re-lit. Rebuilt for Blender 4.1. removed old FlexDMD options and replaced with F12 version. Some table remodelling (raised plastics a bit. Plastics posts have added nuts underneath).
'RC0.0.91: Added partial baked VR cabinet with new accurate cab art ' THANKS DEVIOUS626
'RC1: Updated VR Backbox, added VR specific rails.
'RC1.1: Testing
'RC1.2:








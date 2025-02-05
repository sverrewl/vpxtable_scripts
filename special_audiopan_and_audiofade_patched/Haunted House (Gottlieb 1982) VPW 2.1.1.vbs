'==============================================================================================='
'                                               '
'                     Haunted House                               '
'                                      Gottlieb (1982)                                    '
'                  http://www.ipdb.org/machine.cgi?id=1133                      '
'                                               '
'                                  Created by: VPW                                  '
'                                               '
'==============================================================================================='
'
'
' VPW TEAM
' ----------
' Cyberpez - Table maker, blender builder, and seance medium
' Apophis - Script toiler, undertaker
' Rothbauerw - Physics master, magician
' Tomate - Blender, baker, chef
' Hauntfreaks - Sideshow artist
' Sixtoe - Dust collector
'
'
'//////////////////////////////////////////////////////////////////////////////////////////////////////
' THIS TABLE INCLUDES THE IN-GAME OPTIONS MENU
' To open the menu, press both magnasave buttons at the same time in the game before plunging the ball.
'//////////////////////////////////////////////////////////////////////////////////////////////////////
'
' REQUIRED: VPX 10.8 beta 6 or later
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
'   ZBBR: Ball Brightness
'   ZBRI: Room Brightness
' ZDIP: DIP Switches
' ZKEY: Key Press Handling
' ZSOL: Solenoids
' ZGII: GI
' ZGLO: Glow Balls
' ZFLP: Flippers
' ZRBR: Rubbers
' ZSWI: Switches
' ZGAT: Gates
' ZKIC: Kickers
' ZSTA: Stand-up Targets
' ZKTA: Kicking Targets
' ZDTA: Drop Targets
' ZSLG: Slingshots
' ZBMP: Bumpers
' ZKNO: Knocker
' ZTRP: Trap Door
' ZSEC: Secret Target
' ZDRN: Drain, Trough, and Ball Release
'   ZBRL: Ball Rolling and Drop Sounds
'   ZABS: Ambient ball shadows
'   ZRRL: Ramp Rolling Sound effects
' ZSFX: Mechanical Sound effects
' ZLED: LEDs
'   ZMAT: General Math Functions
' ZNFF: Flipper Corrections
'   ZDMP: Rubber Dampeners
' ZBOU: VPW TargetBouncer
'   ZANI: Misc Animations
'   ZOPT: Table Options
' ZVRR: VR Room Stuff
' ZCHL: Change Log
'
'************************************************************************************************



Option Explicit
Randomize
SetLocale 1033

Const TableVersion = "2.1"


'*******************************************
' ZCON: Constants and Global Variables
'*******************************************

'*****************************
'Rom Version Selector
'*****************************
'1 "hh"   ' 6-digit
'2 "hh7"   ' 7-digit
Dim RomSet: RomSet = 2

Const ShowBGinDesktopMode = False  'Set to True if you want backglass showing (and reflecting) in desktop mode

'Ball Size and Mass
Const BallSize = 50
Const BallMass = 1

Dim HHBall, gBOT
Dim I, x, obj
Dim plungerIM

Dim tablewidth: tablewidth = HH.width
Dim tableheight: tableheight = HH.height

Dim DesktopMode: DesktopMode = HH.ShowDT
Dim VR_Room: If RenderingMode = 2 Then VR_Room = 1 Else VR_Room = 0

Const tnob = 1 ' total number of balls
Const lob = 0 ' total number of locked balls


'*******************************************
' ZVLM: VLM Arrays
'*******************************************

' VLM Arrays - Start
' Arrays per baked part
Dim BP_BRing1: BP_BRing1=Array(BM_BRing1, LM_BLight_BL2_BRing1)
Dim BP_BRing2: BP_BRing2=Array(BM_BRing2, LM_BLight_BL1_BRing2, LM_GI_MPF_giML1_BRing2, LM_GI_MPF_giML2_BRing2)
Dim BP_BRing3: BP_BRing3=Array(BM_BRing3, LM_BLight_BL3_BRing3, LM_GI_UPF_giUL1_BRing3, LM_GI_UPF_gUL2_BRing3, LM_GI_UPF_giUL6_BRing3, LM_inserts_UPF_l21_BRing3, LM_inserts_UPF_l2_BRing3)
Dim BP_BRing4: BP_BRing4=Array(BM_BRing4, LM_BLight_BL4_BRing4, LM_GI_LPF_giLL1_BRing4, LM_inserts_LPF_l26_BRing4, LM_inserts_MPF_l39_BRing4)
Dim BP_BSkirt1: BP_BSkirt1=Array(BM_BSkirt1, LM_BLight_BL2_BSkirt1, LM_GI_LPF_giLL1_BSkirt1, LM_GI_MPF_giML1_BSkirt1, LM_inserts_MPF_l30_BSkirt1)
Dim BP_BSkirt2: BP_BSkirt2=Array(BM_BSkirt2, LM_BLight_BL1_BSkirt2, LM_GI_MPF_giML2_BSkirt2)
Dim BP_BSkirt3: BP_BSkirt3=Array(BM_BSkirt3, LM_BLight_BL3_BSkirt3, LM_GI_UPF_giUL1_BSkirt3, LM_GI_UPF_gUL2_BSkirt3, LM_GI_UPF_giUL3_BSkirt3, LM_GI_UPF_giUL6_BSkirt3, LM_inserts_UPF_l21_BSkirt3, LM_inserts_UPF_l2_BSkirt3)
Dim BP_BSkirt4: BP_BSkirt4=Array(BM_BSkirt4, LM_BLight_BL4_BSkirt4, LM_GI_LPF_giLL1_BSkirt4, LM_inserts_LPF_l26_BSkirt4, LM_inserts_MPF_l8_BSkirt4)
Dim BP_ICLeft: BP_ICLeft=Array(BM_ICLeft)
Dim BP_ICRight: BP_ICRight=Array(BM_ICRight)
Dim BP_LPFR1a: BP_LPFR1a=Array(BM_LPFR1a, LM_GI_LPF_giLL1_LPFR1a)
Dim BP_LPFR1b: BP_LPFR1b=Array(BM_LPFR1b, LM_GI_LPF_giLL1_LPFR1b, LM_inserts_LPF_l25_LPFR1b, LM_inserts_MPF_l30_LPFR1b, LM_inserts_MPF_l31_LPFR1b)
Dim BP_LPFR2a: BP_LPFR2a=Array(BM_LPFR2a, LM_BLight_BL4_LPFR2a, LM_GI_LPF_giLL1_LPFR2a)
Dim BP_LPFR2b: BP_LPFR2b=Array(BM_LPFR2b, LM_BLight_BL4_LPFR2b, LM_GI_LPF_giLL1_LPFR2b)
Dim BP_LSlingPos: BP_LSlingPos=Array(BM_LSlingPos, LM_GI_LPF_giLL1_LSlingPos)
Dim BP_Layer1: BP_Layer1=Array(BM_Layer1)
Dim BP_Layer2: BP_Layer2=Array(BM_Layer2, LM_GI_LPF_giLL1_Layer2, LM_inserts_LPF_l25_Layer2, LM_inserts_MPF_l30_Layer2, LM_inserts_MPF_l31_Layer2)
Dim BP_Layer3: BP_Layer3=Array(BM_Layer3, LM_inserts_UPF_l21_Layer3, LM_inserts_MPF_l51_Layer3)
Dim BP_MPFR1a: BP_MPFR1a=Array(BM_MPFR1a, LM_GI_MPF_giML3_MPFR1a, LM_GI_MPF_giML4_MPFR1a, LM_inserts_MPF_l8_MPFR1a)
Dim BP_MPFR1b: BP_MPFR1b=Array(BM_MPFR1b, LM_GI_MPF_giML3_MPFR1b, LM_GI_MPF_giML4_MPFR1b, LM_inserts_MPF_l39_MPFR1b, LM_inserts_MPF_l47_MPFR1b, LM_inserts_MPF_l8_MPFR1b)
Dim BP_PFLower: BP_PFLower=Array(BM_PFLower, LM_BLight_BL1_PFLower, LM_BLight_BL2_PFLower, LM_BLight_BL4_PFLower, LM_GI_LPF_giLL1_PFLower, LM_GI_MPF_giML1_PFLower, LM_GI_MPF_giML2_PFLower, LM_GI_MPF_giML3_PFLower, LM_inserts_MPF_l18_PFLower, LM_inserts_MPF_l19_PFLower, LM_inserts_MPF_l20_PFLower, LM_inserts_LPF_l24_PFLower, LM_inserts_LPF_l25_PFLower, LM_inserts_MPF_l30_PFLower, LM_inserts_MPF_l31_PFLower, LM_inserts_MPF_l35_PFLower, LM_inserts_MPF_l36_PFLower, LM_inserts_MPF_l37_PFLower, LM_inserts_MPF_l38_PFLower, LM_inserts_MPF_l39_PFLower, LM_inserts_MPF_l46_PFLower, LM_inserts_MPF_l51_PFLower, LM_inserts_MPF_l8_PFLower)
Dim BP_PFMain: BP_PFMain=Array(BM_PFMain, LM_BLight_BL1_PFMain, LM_BLight_BL2_PFMain, LM_BLight_BL4_PFMain, LM_GI_LPF_giLL1_PFMain, LM_GI_MPF_giML1_PFMain, LM_GI_MPF_giML2_PFMain, LM_GI_MPF_giML3_PFMain, LM_GI_MPF_giML4_PFMain, LM_GI_MPF_giML5_PFMain, LM_GI_MPF_giML6_PFMain, LM_GI_UPF_giUL1_PFMain, LM_GI_UPF_gUL2_PFMain, LM_GI_UPF_giUL3_PFMain, LM_GI_UPF_giUL4_PFMain, LM_GI_UPF_giUL5_PFMain, LM_GI_UPF_giUL6_PFMain, LM_inserts_MPF_l18_PFMain, LM_inserts_MPF_l19_PFMain, LM_inserts_MPF_l20_PFMain, LM_inserts_UPF_l21_PFMain, LM_inserts_MPF_l27_PFMain, LM_inserts_MPF_l28_PFMain, LM_inserts_MPF_l29_PFMain, LM_inserts_UPF_l2_PFMain, LM_inserts_MPF_l35_PFMain, LM_inserts_MPF_l36_PFMain, LM_inserts_MPF_l37_PFMain, LM_inserts_MPF_l38_PFMain, LM_inserts_MPF_l39_PFMain, LM_inserts_MPF_l3_PFMain, LM_inserts_UPF_l40_PFMain, LM_inserts_UPF_l41_PFMain, LM_inserts_UPF_l42_PFMain, LM_inserts_MPF_l46_PFMain, LM_inserts_MPF_l47_PFMain, LM_inserts_MPF_l51_PFMain, LM_inserts_MPF_l8_PFMain)
Dim BP_PFUpper: BP_PFUpper=Array(BM_PFUpper, LM_BLight_BL3_PFUpper, LM_GI_UPF_giUL1_PFUpper, LM_GI_UPF_gUL2_PFUpper, LM_GI_UPF_giUL3_PFUpper, LM_GI_UPF_giUL4_PFUpper, LM_GI_UPF_giUL5_PFUpper, LM_GI_UPF_giUL6_PFUpper, LM_inserts_UPF_l21_PFUpper, LM_inserts_UPF_l2_PFUpper, LM_inserts_UPF_l40_PFUpper, LM_inserts_UPF_l41_PFUpper, LM_inserts_UPF_l42_PFUpper)
Dim BP_PLF: BP_PLF=Array(BM_PLF)
Dim BP_PLF2: BP_PLF2=Array(BM_PLF2, LM_BLight_BL2_PLF2)
Dim BP_PLF2up: BP_PLF2up=Array(BM_PLF2up, LM_BLight_BL2_PLF2up)
Dim BP_PLFup: BP_PLFup=Array(BM_PLFup)
Dim BP_PRF: BP_PRF=Array(BM_PRF)
Dim BP_PRF2: BP_PRF2=Array(BM_PRF2, LM_GI_MPF_giML3_PRF2)
Dim BP_PRF2up: BP_PRF2up=Array(BM_PRF2up, LM_GI_MPF_giML3_PRF2up)
Dim BP_PRFup: BP_PRFup=Array(BM_PRFup)
Dim BP_PURF: BP_PURF=Array(BM_PURF, LM_BLight_BL3_PURF, LM_GI_UPF_giUL3_PURF, LM_GI_UPF_giUL4_PURF, LM_GI_UPF_giUL5_PURF, LM_GI_UPF_giUL6_PURF, LM_inserts_UPF_l21_PURF, LM_inserts_UPF_l2_PURF, LM_inserts_UPF_l40_PURF, LM_inserts_UPF_l41_PURF, LM_inserts_UPF_l42_PURF)
Dim BP_PURFup: BP_PURFup=Array(BM_PURFup, LM_BLight_BL3_PURFup, LM_GI_UPF_giUL1_PURFup, LM_GI_UPF_gUL2_PURFup, LM_GI_UPF_giUL3_PURFup, LM_GI_UPF_giUL4_PURFup, LM_GI_UPF_giUL5_PURFup, LM_GI_UPF_giUL6_PURFup, LM_inserts_UPF_l21_PURFup, LM_inserts_UPF_l2_PURFup, LM_inserts_UPF_l40_PURFup, LM_inserts_UPF_l41_PURFup, LM_inserts_UPF_l42_PURFup)
Dim BP_Parts: BP_Parts=Array(BM_Parts, LM_BLight_BL1_Parts, LM_BLight_BL2_Parts, LM_BLight_BL3_Parts, LM_BLight_BL4_Parts, LM_GI_LPF_giLL1_Parts, LM_GI_MPF_giML1_Parts, LM_GI_MPF_giML2_Parts, LM_GI_MPF_giML3_Parts, LM_GI_MPF_giML4_Parts, LM_GI_MPF_giML5_Parts, LM_GI_MPF_giML6_Parts, LM_GI_UPF_giUL1_Parts, LM_GI_UPF_gUL2_Parts, LM_GI_UPF_giUL3_Parts, LM_GI_UPF_giUL4_Parts, LM_GI_UPF_giUL5_Parts, LM_GI_UPF_giUL6_Parts, LM_inserts_MPF_l18_Parts, LM_inserts_UPF_l21_Parts, LM_inserts_LPF_l24_Parts, LM_inserts_LPF_l25_Parts, LM_inserts_LPF_l26_Parts, LM_inserts_MPF_l27_Parts, LM_inserts_MPF_l28_Parts, LM_inserts_MPF_l29_Parts, LM_inserts_UPF_l2_Parts, LM_inserts_MPF_l30_Parts, LM_inserts_MPF_l33_Parts, LM_inserts_MPF_l34_Parts, LM_inserts_MPF_l35_Parts, LM_inserts_MPF_l36_Parts, LM_inserts_MPF_l37_Parts, LM_inserts_MPF_l38_Parts, LM_inserts_MPF_l39_Parts, LM_inserts_MPF_l3_Parts, LM_inserts_UPF_l40_Parts, LM_inserts_UPF_l41_Parts, LM_inserts_UPF_l42_Parts, LM_inserts_LPF_l45_Parts, LM_inserts_MPF_l46_Parts, _
  LM_inserts_MPF_l47_Parts, LM_inserts_MPF_l51_Parts, LM_inserts_MPF_l8_Parts)
Dim BP_PartsLP: BP_PartsLP=Array(BM_PartsLP, LM_BLight_BL2_PartsLP, LM_BLight_BL4_PartsLP, LM_GI_LPF_giLL1_PartsLP, LM_GI_MPF_giML1_PartsLP, LM_GI_MPF_giML2_PartsLP, LM_GI_MPF_giML3_PartsLP, LM_GI_MPF_giML4_PartsLP, LM_inserts_MPF_l19_PartsLP, LM_inserts_LPF_l25_PartsLP, LM_inserts_LPF_l26_PartsLP, LM_inserts_MPF_l30_PartsLP, LM_inserts_MPF_l31_PartsLP, LM_inserts_MPF_l35_PartsLP, LM_inserts_MPF_l36_PartsLP, LM_inserts_MPF_l37_PartsLP, LM_inserts_MPF_l38_PartsLP, LM_inserts_MPF_l39_PartsLP, LM_inserts_LPF_l45_PartsLP, LM_inserts_MPF_l8_PartsLP)
Dim BP_SideRails: BP_SideRails=Array(BM_SideRails, LM_GI_UPF_giUL5_SideRails)
Dim BP_TWKicker1: BP_TWKicker1=Array(BM_TWKicker1)
Dim BP_UPFR1a: BP_UPFR1a=Array(BM_UPFR1a, LM_GI_UPF_giUL1_UPFR1a, LM_GI_UPF_gUL2_UPFR1a, LM_GI_UPF_giUL3_UPFR1a, LM_GI_UPF_giUL6_UPFR1a, LM_inserts_UPF_l21_UPFR1a)
Dim BP_UPFR1b: BP_UPFR1b=Array(BM_UPFR1b, LM_GI_UPF_giUL1_UPFR1b, LM_GI_UPF_gUL2_UPFR1b, LM_GI_UPF_giUL3_UPFR1b, LM_GI_UPF_giUL6_UPFR1b, LM_inserts_UPF_l21_UPFR1b)
Dim BP_UPFR2a: BP_UPFR2a=Array(BM_UPFR2a, LM_BLight_BL3_UPFR2a, LM_GI_UPF_giUL1_UPFR2a, LM_GI_UPF_gUL2_UPFR2a, LM_GI_UPF_giUL3_UPFR2a, LM_GI_UPF_giUL6_UPFR2a)
Dim BP_UPFR2b: BP_UPFR2b=Array(BM_UPFR2b, LM_BLight_BL3_UPFR2b, LM_GI_UPF_giUL1_UPFR2b, LM_GI_UPF_gUL2_UPFR2b, LM_GI_UPF_giUL3_UPFR2b, LM_GI_UPF_giUL6_UPFR2b)
Dim BP_UPFR2c: BP_UPFR2c=Array(BM_UPFR2c, LM_BLight_BL3_UPFR2c, LM_GI_UPF_giUL1_UPFR2c, LM_GI_UPF_gUL2_UPFR2c, LM_GI_UPF_giUL3_UPFR2c, LM_GI_UPF_giUL6_UPFR2c, LM_inserts_UPF_l21_UPFR2c)
Dim BP_UPFR2d: BP_UPFR2d=Array(BM_UPFR2d, LM_BLight_BL3_UPFR2d, LM_GI_UPF_giUL1_UPFR2d, LM_GI_UPF_gUL2_UPFR2d, LM_GI_UPF_giUL3_UPFR2d, LM_GI_UPF_giUL6_UPFR2d)
Dim BP_UPFR3a: BP_UPFR3a=Array(BM_UPFR3a, LM_BLight_BL3_UPFR3a, LM_GI_UPF_giUL1_UPFR3a, LM_GI_UPF_gUL2_UPFR3a, LM_GI_UPF_giUL3_UPFR3a, LM_GI_UPF_giUL6_UPFR3a)
Dim BP_UPFR3b: BP_UPFR3b=Array(BM_UPFR3b, LM_BLight_BL3_UPFR3b, LM_GI_UPF_giUL1_UPFR3b, LM_GI_UPF_gUL2_UPFR3b, LM_GI_UPF_giUL3_UPFR3b, LM_GI_UPF_giUL6_UPFR3b)
Dim BP_UPFR4a: BP_UPFR4a=Array(BM_UPFR4a, LM_BLight_BL3_UPFR4a, LM_GI_UPF_giUL1_UPFR4a, LM_GI_UPF_gUL2_UPFR4a, LM_GI_UPF_giUL3_UPFR4a, LM_GI_UPF_giUL4_UPFR4a)
Dim BP_UPFR4b: BP_UPFR4b=Array(BM_UPFR4b, LM_BLight_BL3_UPFR4b, LM_GI_UPF_giUL1_UPFR4b, LM_GI_UPF_gUL2_UPFR4b, LM_GI_UPF_giUL3_UPFR4b, LM_GI_UPF_giUL4_UPFR4b)
Dim BP_USlingPos: BP_USlingPos=Array(BM_USlingPos, LM_GI_MPF_giML4_USlingPos, LM_inserts_MPF_l39_USlingPos)
Dim BP_dust: BP_dust=Array(BM_dust, LM_GI_LPF_giLL1_dust, LM_inserts_MPF_l18_dust, LM_inserts_MPF_l19_dust, LM_inserts_MPF_l20_dust, LM_inserts_MPF_l35_dust, LM_inserts_MPF_l36_dust, LM_inserts_MPF_l37_dust, LM_inserts_MPF_l38_dust, LM_inserts_MPF_l39_dust)
Dim BP_lockDownBar: BP_lockDownBar=Array(BM_lockDownBar)
Dim BP_pInfoCard: BP_pInfoCard=Array(BM_pInfoCard)
Dim BP_pLLF: BP_pLLF=Array(BM_pLLF, LM_GI_LPF_giLL1_pLLF, LM_inserts_MPF_l18_pLLF, LM_inserts_MPF_l19_pLLF, LM_inserts_MPF_l20_pLLF, LM_inserts_MPF_l31_pLLF, LM_inserts_MPF_l35_pLLF, LM_inserts_MPF_l36_pLLF)
Dim BP_pLLFup: BP_pLLFup=Array(BM_pLLFup, LM_GI_LPF_giLL1_pLLFup, LM_inserts_MPF_l18_pLLFup, LM_inserts_MPF_l19_pLLFup, LM_inserts_MPF_l31_pLLFup, LM_inserts_MPF_l35_pLLFup, LM_inserts_MPF_l36_pLLFup)
Dim BP_pLRF: BP_pLRF=Array(BM_pLRF, LM_BLight_BL4_pLRF, LM_GI_LPF_giLL1_pLRF, LM_inserts_MPF_l18_pLRF, LM_inserts_MPF_l35_pLRF, LM_inserts_MPF_l36_pLRF, LM_inserts_MPF_l37_pLRF, LM_inserts_MPF_l38_pLRF, LM_inserts_MPF_l51_pLRF)
Dim BP_pLRFUup: BP_pLRFUup=Array(BM_pLRFUup, LM_BLight_BL4_pLRFUup, LM_GI_LPF_giLL1_pLRFUup, LM_inserts_MPF_l18_pLRFUup, LM_inserts_MPF_l35_pLRFUup, LM_inserts_MPF_l36_pLRFUup, LM_inserts_MPF_l37_pLRFUup, LM_inserts_MPF_l38_pLRFUup, LM_inserts_MPF_l51_pLRFUup)
Dim BP_pSSpring: BP_pSSpring=Array(BM_pSSpring, LM_GI_LPF_giLL1_pSSpring, LM_GI_MPF_giML3_pSSpring)
Dim BP_pSSpring2: BP_pSSpring2=Array(BM_pSSpring2, LM_GI_LPF_giLL1_pSSpring2, LM_GI_MPF_giML4_pSSpring2)
Dim BP_pSW00: BP_pSW00=Array(BM_pSW00, LM_GI_LPF_giLL1_pSW00)
Dim BP_pSW02: BP_pSW02=Array(BM_pSW02, LM_BLight_BL3_pSW02, LM_GI_UPF_giUL1_pSW02, LM_GI_UPF_gUL2_pSW02, LM_GI_UPF_giUL3_pSW02, LM_GI_UPF_giUL6_pSW02)
Dim BP_pSW03: BP_pSW03=Array(BM_pSW03, LM_GI_UPF_giUL4_pSW03, LM_GI_UPF_giUL5_pSW03, LM_GI_UPF_giUL6_pSW03, LM_inserts_UPF_l2_pSW03)
Dim BP_pSW04: BP_pSW04=Array(BM_pSW04, LM_BLight_BL2_pSW04, LM_GI_LPF_giLL1_pSW04, LM_GI_MPF_giML1_pSW04)
Dim BP_pSW05: BP_pSW05=Array(BM_pSW05, LM_BLight_BL1_pSW05, LM_GI_MPF_giML1_pSW05, LM_GI_MPF_giML2_pSW05, LM_inserts_MPF_l18_pSW05, LM_inserts_MPF_l19_pSW05, LM_inserts_MPF_l51_pSW05)
Dim BP_pSW10: BP_pSW10=Array(BM_pSW10, LM_GI_LPF_giLL1_pSW10)
Dim BP_pSW11: BP_pSW11=Array(BM_pSW11, LM_GI_LPF_giLL1_pSW11)
Dim BP_pSW12: BP_pSW12=Array(BM_pSW12, LM_BLight_BL3_pSW12, LM_GI_UPF_giUL1_pSW12, LM_GI_UPF_gUL2_pSW12, LM_GI_UPF_giUL3_pSW12, LM_GI_UPF_giUL6_pSW12)
Dim BP_pSW13: BP_pSW13=Array(BM_pSW13, LM_GI_UPF_giUL4_pSW13, LM_GI_UPF_giUL5_pSW13, LM_GI_UPF_giUL6_pSW13, LM_inserts_UPF_l40_pSW13, LM_inserts_UPF_l41_pSW13)
Dim BP_pSW14: BP_pSW14=Array(BM_pSW14, LM_BLight_BL2_pSW14, LM_GI_MPF_giML1_pSW14, LM_GI_MPF_giML2_pSW14)
Dim BP_pSW15: BP_pSW15=Array(BM_pSW15, LM_GI_MPF_giML3_pSW15, LM_GI_UPF_giUL6_pSW15, LM_inserts_UPF_l21_pSW15, LM_inserts_MPF_l51_pSW15)
Dim BP_pSW20: BP_pSW20=Array(BM_pSW20, LM_GI_LPF_giLL1_pSW20, LM_inserts_MPF_l30_pSW20)
Dim BP_pSW22: BP_pSW22=Array(BM_pSW22, LM_BLight_BL3_pSW22, LM_GI_UPF_giUL1_pSW22, LM_GI_UPF_gUL2_pSW22, LM_GI_UPF_giUL3_pSW22, LM_GI_UPF_giUL6_pSW22)
Dim BP_pSW23: BP_pSW23=Array(BM_pSW23, LM_GI_UPF_giUL4_pSW23, LM_GI_UPF_giUL5_pSW23, LM_inserts_UPF_l41_pSW23)
Dim BP_pSW24: BP_pSW24=Array(BM_pSW24, LM_GI_MPF_giML1_pSW24, LM_GI_MPF_giML2_pSW24)
Dim BP_pSW25: BP_pSW25=Array(BM_pSW25, LM_GI_UPF_giUL6_pSW25, LM_inserts_UPF_l21_pSW25)
Dim BP_pSW30: BP_pSW30=Array(BM_pSW30, LM_GI_LPF_giLL1_pSW30, LM_inserts_MPF_l30_pSW30)
Dim BP_pSW32: BP_pSW32=Array(BM_pSW32, LM_BLight_BL3_pSW32, LM_GI_UPF_giUL1_pSW32, LM_GI_UPF_gUL2_pSW32, LM_GI_UPF_giUL3_pSW32, LM_GI_UPF_giUL6_pSW32)
Dim BP_pSW33: BP_pSW33=Array(BM_pSW33, LM_GI_UPF_giUL4_pSW33, LM_GI_UPF_giUL5_pSW33, LM_inserts_UPF_l41_pSW33, LM_inserts_UPF_l42_pSW33)
Dim BP_pSW40: BP_pSW40=Array(BM_pSW40, LM_GI_LPF_giLL1_pSW40, LM_inserts_MPF_l30_pSW40)
Dim BP_pSW50: BP_pSW50=Array(BM_pSW50, LM_GI_LPF_giLL1_pSW50)
Dim BP_pSW55: BP_pSW55=Array(BM_pSW55, LM_GI_MPF_giML3_pSW55, LM_inserts_MPF_l8_pSW55)
Dim BP_pSW60: BP_pSW60=Array(BM_pSW60, LM_GI_LPF_giLL1_pSW60, LM_inserts_LPF_l25_pSW60, LM_inserts_MPF_l30_pSW60, LM_inserts_MPF_l31_pSW60)
Dim BP_pSWSecret: BP_pSWSecret=Array(BM_pSWSecret, LM_GI_MPF_giML1_pSWSecret, LM_GI_MPF_giML2_pSWSecret, LM_inserts_MPF_l18_pSWSecret, LM_inserts_UPF_l21_pSWSecret, LM_inserts_MPF_l35_pSWSecret, LM_inserts_MPF_l51_pSWSecret)
Dim BP_pTrapDoor: BP_pTrapDoor=Array(BM_pTrapDoor, LM_GI_MPF_giML4_pTrapDoor, LM_GI_UPF_giUL6_pTrapDoor, LM_inserts_UPF_l21_pTrapDoor, LM_inserts_MPF_l47_pTrapDoor)
Dim BP_pULF: BP_pULF=Array(BM_pULF, LM_GI_UPF_giUL1_pULF, LM_GI_UPF_gUL2_pULF, LM_GI_UPF_giUL3_pULF, LM_GI_UPF_giUL6_pULF, LM_inserts_UPF_l21_pULF)
Dim BP_pULFup: BP_pULFup=Array(BM_pULFup, LM_GI_UPF_giUL1_pULFup, LM_GI_UPF_gUL2_pULFup, LM_GI_UPF_giUL3_pULFup, LM_GI_UPF_giUL6_pULFup, LM_inserts_UPF_l21_pULFup)
Dim BP_p_gate1: BP_p_gate1=Array(BM_p_gate1, LM_GI_MPF_giML6_p_gate1)
Dim BP_p_gate2: BP_p_gate2=Array(BM_p_gate2, LM_GI_UPF_gUL2_p_gate2, LM_GI_UPF_giUL6_p_gate2, LM_inserts_UPF_l21_p_gate2)
Dim BP_p_gate3: BP_p_gate3=Array(BM_p_gate3, LM_GI_MPF_giML4_p_gate3, LM_inserts_MPF_l47_p_gate3)
Dim BP_p_gate4: BP_p_gate4=Array(BM_p_gate4, LM_GI_LPF_giLL1_p_gate4, LM_GI_MPF_giML3_p_gate4)
Dim BP_psw21: BP_psw21=Array(BM_psw21, LM_GI_LPF_giLL1_psw21)
Dim BP_psw34: BP_psw34=Array(BM_psw34, LM_BLight_BL2_psw34, LM_GI_LPF_giLL1_psw34, LM_inserts_MPF_l46_psw34)
Dim BP_psw54: BP_psw54=Array(BM_psw54)
Dim BP_underPFL: BP_underPFL=Array(BM_underPFL, LM_GI_LPF_giLL1_underPFL, LM_inserts_LPF_l23_underPFL, LM_inserts_LPF_l24_underPFL, LM_inserts_LPF_l25_underPFL, LM_inserts_LPF_l26_underPFL, LM_inserts_LPF_l45_underPFL, LM_inserts_MPF_l51_underPFL)
Dim BP_underPFM: BP_underPFM=Array(BM_underPFM, LM_BLight_BL4_underPFM, LM_GI_LPF_giLL1_underPFM, LM_GI_MPF_giML1_underPFM, LM_GI_MPF_giML2_underPFM, LM_GI_MPF_giML3_underPFM, LM_GI_MPF_giML4_underPFM, LM_GI_MPF_giML5_underPFM, LM_GI_UPF_giUL6_underPFM, LM_inserts_MPF_l18_underPFM, LM_inserts_MPF_l19_underPFM, LM_inserts_MPF_l20_underPFM, LM_inserts_UPF_l21_underPFM, LM_inserts_MPF_l22_underPFM, LM_inserts_MPF_l27_underPFM, LM_inserts_MPF_l28_underPFM, LM_inserts_MPF_l29_underPFM, LM_inserts_MPF_l30_underPFM, LM_inserts_MPF_l31_underPFM, LM_inserts_MPF_l32_underPFM, LM_inserts_MPF_l33_underPFM, LM_inserts_MPF_l34_underPFM, LM_inserts_MPF_l35_underPFM, LM_inserts_MPF_l36_underPFM, LM_inserts_MPF_l37_underPFM, LM_inserts_MPF_l38_underPFM, LM_inserts_MPF_l39_underPFM, LM_inserts_MPF_l3_underPFM, LM_inserts_MPF_l46_underPFM, LM_inserts_MPF_l47_underPFM, LM_inserts_MPF_l51_underPFM, LM_inserts_MPF_l8_underPFM)
Dim BP_underPFU: BP_underPFU=Array(BM_underPFU, LM_BLight_BL3_underPFU, LM_inserts_UPF_l2_underPFU, LM_inserts_UPF_l40_underPFU, LM_inserts_UPF_l41_underPFU, LM_inserts_UPF_l42_underPFU, LM_inserts_UPF_l44_underPFU)
' Arrays per lighting scenario
Dim BL_BLight_BL1: BL_BLight_BL1=Array(LM_BLight_BL1_BRing2, LM_BLight_BL1_BSkirt2, LM_BLight_BL1_PFLower, LM_BLight_BL1_PFMain, LM_BLight_BL1_Parts, LM_BLight_BL1_pSW05)
Dim BL_BLight_BL2: BL_BLight_BL2=Array(LM_BLight_BL2_BRing1, LM_BLight_BL2_BSkirt1, LM_BLight_BL2_PFLower, LM_BLight_BL2_PFMain, LM_BLight_BL2_PLF2, LM_BLight_BL2_PLF2up, LM_BLight_BL2_Parts, LM_BLight_BL2_PartsLP, LM_BLight_BL2_pSW04, LM_BLight_BL2_pSW14, LM_BLight_BL2_psw34)
Dim BL_BLight_BL3: BL_BLight_BL3=Array(LM_BLight_BL3_BRing3, LM_BLight_BL3_BSkirt3, LM_BLight_BL3_PFUpper, LM_BLight_BL3_PURF, LM_BLight_BL3_PURFup, LM_BLight_BL3_Parts, LM_BLight_BL3_UPFR2a, LM_BLight_BL3_UPFR2b, LM_BLight_BL3_UPFR2c, LM_BLight_BL3_UPFR2d, LM_BLight_BL3_UPFR3a, LM_BLight_BL3_UPFR3b, LM_BLight_BL3_UPFR4a, LM_BLight_BL3_UPFR4b, LM_BLight_BL3_pSW02, LM_BLight_BL3_pSW12, LM_BLight_BL3_pSW22, LM_BLight_BL3_pSW32, LM_BLight_BL3_underPFU)
Dim BL_BLight_BL4: BL_BLight_BL4=Array(LM_BLight_BL4_BRing4, LM_BLight_BL4_BSkirt4, LM_BLight_BL4_LPFR2a, LM_BLight_BL4_LPFR2b, LM_BLight_BL4_PFLower, LM_BLight_BL4_PFMain, LM_BLight_BL4_Parts, LM_BLight_BL4_PartsLP, LM_BLight_BL4_pLRF, LM_BLight_BL4_pLRFUup, LM_BLight_BL4_underPFM)
Dim BL_GI_LPF_giLL1: BL_GI_LPF_giLL1=Array(LM_GI_LPF_giLL1_BRing4, LM_GI_LPF_giLL1_BSkirt1, LM_GI_LPF_giLL1_BSkirt4, LM_GI_LPF_giLL1_LPFR1a, LM_GI_LPF_giLL1_LPFR1b, LM_GI_LPF_giLL1_LPFR2a, LM_GI_LPF_giLL1_LPFR2b, LM_GI_LPF_giLL1_LSlingPos, LM_GI_LPF_giLL1_Layer2, LM_GI_LPF_giLL1_PFLower, LM_GI_LPF_giLL1_PFMain, LM_GI_LPF_giLL1_Parts, LM_GI_LPF_giLL1_PartsLP, LM_GI_LPF_giLL1_dust, LM_GI_LPF_giLL1_pLLF, LM_GI_LPF_giLL1_pLLFup, LM_GI_LPF_giLL1_pLRF, LM_GI_LPF_giLL1_pLRFUup, LM_GI_LPF_giLL1_pSSpring, LM_GI_LPF_giLL1_pSSpring2, LM_GI_LPF_giLL1_pSW00, LM_GI_LPF_giLL1_pSW04, LM_GI_LPF_giLL1_pSW10, LM_GI_LPF_giLL1_pSW11, LM_GI_LPF_giLL1_pSW20, LM_GI_LPF_giLL1_pSW30, LM_GI_LPF_giLL1_pSW40, LM_GI_LPF_giLL1_pSW50, LM_GI_LPF_giLL1_pSW60, LM_GI_LPF_giLL1_p_gate4, LM_GI_LPF_giLL1_psw21, LM_GI_LPF_giLL1_psw34, LM_GI_LPF_giLL1_underPFL, LM_GI_LPF_giLL1_underPFM)
Dim BL_GI_MPF_giML1: BL_GI_MPF_giML1=Array(LM_GI_MPF_giML1_BRing2, LM_GI_MPF_giML1_BSkirt1, LM_GI_MPF_giML1_PFLower, LM_GI_MPF_giML1_PFMain, LM_GI_MPF_giML1_Parts, LM_GI_MPF_giML1_PartsLP, LM_GI_MPF_giML1_pSW04, LM_GI_MPF_giML1_pSW05, LM_GI_MPF_giML1_pSW14, LM_GI_MPF_giML1_pSW24, LM_GI_MPF_giML1_pSWSecret, LM_GI_MPF_giML1_underPFM)
Dim BL_GI_MPF_giML2: BL_GI_MPF_giML2=Array(LM_GI_MPF_giML2_BRing2, LM_GI_MPF_giML2_BSkirt2, LM_GI_MPF_giML2_PFLower, LM_GI_MPF_giML2_PFMain, LM_GI_MPF_giML2_Parts, LM_GI_MPF_giML2_PartsLP, LM_GI_MPF_giML2_pSW05, LM_GI_MPF_giML2_pSW14, LM_GI_MPF_giML2_pSW24, LM_GI_MPF_giML2_pSWSecret, LM_GI_MPF_giML2_underPFM)
Dim BL_GI_MPF_giML3: BL_GI_MPF_giML3=Array(LM_GI_MPF_giML3_MPFR1a, LM_GI_MPF_giML3_MPFR1b, LM_GI_MPF_giML3_PFLower, LM_GI_MPF_giML3_PFMain, LM_GI_MPF_giML3_PRF2, LM_GI_MPF_giML3_PRF2up, LM_GI_MPF_giML3_Parts, LM_GI_MPF_giML3_PartsLP, LM_GI_MPF_giML3_pSSpring, LM_GI_MPF_giML3_pSW15, LM_GI_MPF_giML3_pSW55, LM_GI_MPF_giML3_p_gate4, LM_GI_MPF_giML3_underPFM)
Dim BL_GI_MPF_giML4: BL_GI_MPF_giML4=Array(LM_GI_MPF_giML4_MPFR1a, LM_GI_MPF_giML4_MPFR1b, LM_GI_MPF_giML4_PFMain, LM_GI_MPF_giML4_Parts, LM_GI_MPF_giML4_PartsLP, LM_GI_MPF_giML4_USlingPos, LM_GI_MPF_giML4_pSSpring2, LM_GI_MPF_giML4_pTrapDoor, LM_GI_MPF_giML4_p_gate3, LM_GI_MPF_giML4_underPFM)
Dim BL_GI_MPF_giML5: BL_GI_MPF_giML5=Array(LM_GI_MPF_giML5_PFMain, LM_GI_MPF_giML5_Parts, LM_GI_MPF_giML5_underPFM)
Dim BL_GI_MPF_giML6: BL_GI_MPF_giML6=Array(LM_GI_MPF_giML6_PFMain, LM_GI_MPF_giML6_Parts, LM_GI_MPF_giML6_p_gate1)
Dim BL_GI_UPF_gUL2: BL_GI_UPF_gUL2=Array(LM_GI_UPF_gUL2_BRing3, LM_GI_UPF_gUL2_BSkirt3, LM_GI_UPF_gUL2_PFMain, LM_GI_UPF_gUL2_PFUpper, LM_GI_UPF_gUL2_PURFup, LM_GI_UPF_gUL2_Parts, LM_GI_UPF_gUL2_UPFR1a, LM_GI_UPF_gUL2_UPFR1b, LM_GI_UPF_gUL2_UPFR2a, LM_GI_UPF_gUL2_UPFR2b, LM_GI_UPF_gUL2_UPFR2c, LM_GI_UPF_gUL2_UPFR2d, LM_GI_UPF_gUL2_UPFR3a, LM_GI_UPF_gUL2_UPFR3b, LM_GI_UPF_gUL2_UPFR4a, LM_GI_UPF_gUL2_UPFR4b, LM_GI_UPF_gUL2_pSW02, LM_GI_UPF_gUL2_pSW12, LM_GI_UPF_gUL2_pSW22, LM_GI_UPF_gUL2_pSW32, LM_GI_UPF_gUL2_pULF, LM_GI_UPF_gUL2_pULFup, LM_GI_UPF_gUL2_p_gate2)
Dim BL_GI_UPF_giUL1: BL_GI_UPF_giUL1=Array(LM_GI_UPF_giUL1_BRing3, LM_GI_UPF_giUL1_BSkirt3, LM_GI_UPF_giUL1_PFMain, LM_GI_UPF_giUL1_PFUpper, LM_GI_UPF_giUL1_PURFup, LM_GI_UPF_giUL1_Parts, LM_GI_UPF_giUL1_UPFR1a, LM_GI_UPF_giUL1_UPFR1b, LM_GI_UPF_giUL1_UPFR2a, LM_GI_UPF_giUL1_UPFR2b, LM_GI_UPF_giUL1_UPFR2c, LM_GI_UPF_giUL1_UPFR2d, LM_GI_UPF_giUL1_UPFR3a, LM_GI_UPF_giUL1_UPFR3b, LM_GI_UPF_giUL1_UPFR4a, LM_GI_UPF_giUL1_UPFR4b, LM_GI_UPF_giUL1_pSW02, LM_GI_UPF_giUL1_pSW12, LM_GI_UPF_giUL1_pSW22, LM_GI_UPF_giUL1_pSW32, LM_GI_UPF_giUL1_pULF, LM_GI_UPF_giUL1_pULFup)
Dim BL_GI_UPF_giUL3: BL_GI_UPF_giUL3=Array(LM_GI_UPF_giUL3_BSkirt3, LM_GI_UPF_giUL3_PFMain, LM_GI_UPF_giUL3_PFUpper, LM_GI_UPF_giUL3_PURF, LM_GI_UPF_giUL3_PURFup, LM_GI_UPF_giUL3_Parts, LM_GI_UPF_giUL3_UPFR1a, LM_GI_UPF_giUL3_UPFR1b, LM_GI_UPF_giUL3_UPFR2a, LM_GI_UPF_giUL3_UPFR2b, LM_GI_UPF_giUL3_UPFR2c, LM_GI_UPF_giUL3_UPFR2d, LM_GI_UPF_giUL3_UPFR3a, LM_GI_UPF_giUL3_UPFR3b, LM_GI_UPF_giUL3_UPFR4a, LM_GI_UPF_giUL3_UPFR4b, LM_GI_UPF_giUL3_pSW02, LM_GI_UPF_giUL3_pSW12, LM_GI_UPF_giUL3_pSW22, LM_GI_UPF_giUL3_pSW32, LM_GI_UPF_giUL3_pULF, LM_GI_UPF_giUL3_pULFup)
Dim BL_GI_UPF_giUL4: BL_GI_UPF_giUL4=Array(LM_GI_UPF_giUL4_PFMain, LM_GI_UPF_giUL4_PFUpper, LM_GI_UPF_giUL4_PURF, LM_GI_UPF_giUL4_PURFup, LM_GI_UPF_giUL4_Parts, LM_GI_UPF_giUL4_UPFR4a, LM_GI_UPF_giUL4_UPFR4b, LM_GI_UPF_giUL4_pSW03, LM_GI_UPF_giUL4_pSW13, LM_GI_UPF_giUL4_pSW23, LM_GI_UPF_giUL4_pSW33)
Dim BL_GI_UPF_giUL5: BL_GI_UPF_giUL5=Array(LM_GI_UPF_giUL5_PFMain, LM_GI_UPF_giUL5_PFUpper, LM_GI_UPF_giUL5_PURF, LM_GI_UPF_giUL5_PURFup, LM_GI_UPF_giUL5_Parts, LM_GI_UPF_giUL5_SideRails, LM_GI_UPF_giUL5_pSW03, LM_GI_UPF_giUL5_pSW13, LM_GI_UPF_giUL5_pSW23, LM_GI_UPF_giUL5_pSW33)
Dim BL_GI_UPF_giUL6: BL_GI_UPF_giUL6=Array(LM_GI_UPF_giUL6_BRing3, LM_GI_UPF_giUL6_BSkirt3, LM_GI_UPF_giUL6_PFMain, LM_GI_UPF_giUL6_PFUpper, LM_GI_UPF_giUL6_PURF, LM_GI_UPF_giUL6_PURFup, LM_GI_UPF_giUL6_Parts, LM_GI_UPF_giUL6_UPFR1a, LM_GI_UPF_giUL6_UPFR1b, LM_GI_UPF_giUL6_UPFR2a, LM_GI_UPF_giUL6_UPFR2b, LM_GI_UPF_giUL6_UPFR2c, LM_GI_UPF_giUL6_UPFR2d, LM_GI_UPF_giUL6_UPFR3a, LM_GI_UPF_giUL6_UPFR3b, LM_GI_UPF_giUL6_pSW02, LM_GI_UPF_giUL6_pSW03, LM_GI_UPF_giUL6_pSW12, LM_GI_UPF_giUL6_pSW13, LM_GI_UPF_giUL6_pSW15, LM_GI_UPF_giUL6_pSW22, LM_GI_UPF_giUL6_pSW25, LM_GI_UPF_giUL6_pSW32, LM_GI_UPF_giUL6_pTrapDoor, LM_GI_UPF_giUL6_pULF, LM_GI_UPF_giUL6_pULFup, LM_GI_UPF_giUL6_p_gate2, LM_GI_UPF_giUL6_underPFM)
Dim BL_Lit_Room: BL_Lit_Room=Array(BM_BRing1, BM_BRing2, BM_BRing3, BM_BRing4, BM_BSkirt1, BM_BSkirt2, BM_BSkirt3, BM_BSkirt4, BM_ICLeft, BM_ICRight, BM_LPFR1a, BM_LPFR1b, BM_LPFR2a, BM_LPFR2b, BM_LSlingPos, BM_Layer1, BM_Layer2, BM_Layer3, BM_MPFR1a, BM_MPFR1b, BM_PFLower, BM_PFMain, BM_PFUpper, BM_PLF, BM_PLF2, BM_PLF2up, BM_PLFup, BM_PRF, BM_PRF2, BM_PRF2up, BM_PRFup, BM_PURF, BM_PURFup, BM_Parts, BM_PartsLP, BM_SideRails, BM_TWKicker1, BM_UPFR1a, BM_UPFR1b, BM_UPFR2a, BM_UPFR2b, BM_UPFR2c, BM_UPFR2d, BM_UPFR3a, BM_UPFR3b, BM_UPFR4a, BM_UPFR4b, BM_USlingPos, BM_dust, BM_lockDownBar, BM_pInfoCard, BM_pLLF, BM_pLLFup, BM_pLRF, BM_pLRFUup, BM_pSSpring, BM_pSSpring2, BM_pSW00, BM_pSW02, BM_pSW03, BM_pSW04, BM_pSW05, BM_pSW10, BM_pSW11, BM_pSW12, BM_pSW13, BM_pSW14, BM_pSW15, BM_pSW20, BM_pSW22, BM_pSW23, BM_pSW24, BM_pSW25, BM_pSW30, BM_pSW32, BM_pSW33, BM_pSW40, BM_pSW50, BM_pSW55, BM_pSW60, BM_pSWSecret, BM_pTrapDoor, BM_pULF, BM_pULFup, BM_p_gate1, BM_p_gate2, BM_p_gate3, BM_p_gate4, BM_psw21, BM_psw34, _
  BM_psw54, BM_underPFL, BM_underPFM, BM_underPFU)
Dim BL_inserts_LPF_l23: BL_inserts_LPF_l23=Array(LM_inserts_LPF_l23_underPFL)
Dim BL_inserts_LPF_l24: BL_inserts_LPF_l24=Array(LM_inserts_LPF_l24_PFLower, LM_inserts_LPF_l24_Parts, LM_inserts_LPF_l24_underPFL)
Dim BL_inserts_LPF_l25: BL_inserts_LPF_l25=Array(LM_inserts_LPF_l25_LPFR1b, LM_inserts_LPF_l25_Layer2, LM_inserts_LPF_l25_PFLower, LM_inserts_LPF_l25_Parts, LM_inserts_LPF_l25_PartsLP, LM_inserts_LPF_l25_pSW60, LM_inserts_LPF_l25_underPFL)
Dim BL_inserts_LPF_l26: BL_inserts_LPF_l26=Array(LM_inserts_LPF_l26_BRing4, LM_inserts_LPF_l26_BSkirt4, LM_inserts_LPF_l26_Parts, LM_inserts_LPF_l26_PartsLP, LM_inserts_LPF_l26_underPFL)
Dim BL_inserts_LPF_l45: BL_inserts_LPF_l45=Array(LM_inserts_LPF_l45_Parts, LM_inserts_LPF_l45_PartsLP, LM_inserts_LPF_l45_underPFL)
Dim BL_inserts_MPF_l18: BL_inserts_MPF_l18=Array(LM_inserts_MPF_l18_PFLower, LM_inserts_MPF_l18_PFMain, LM_inserts_MPF_l18_Parts, LM_inserts_MPF_l18_dust, LM_inserts_MPF_l18_pLLF, LM_inserts_MPF_l18_pLLFup, LM_inserts_MPF_l18_pLRF, LM_inserts_MPF_l18_pLRFUup, LM_inserts_MPF_l18_pSW05, LM_inserts_MPF_l18_pSWSecret, LM_inserts_MPF_l18_underPFM)
Dim BL_inserts_MPF_l19: BL_inserts_MPF_l19=Array(LM_inserts_MPF_l19_PFLower, LM_inserts_MPF_l19_PFMain, LM_inserts_MPF_l19_PartsLP, LM_inserts_MPF_l19_dust, LM_inserts_MPF_l19_pLLF, LM_inserts_MPF_l19_pLLFup, LM_inserts_MPF_l19_pSW05, LM_inserts_MPF_l19_underPFM)
Dim BL_inserts_MPF_l20: BL_inserts_MPF_l20=Array(LM_inserts_MPF_l20_PFLower, LM_inserts_MPF_l20_PFMain, LM_inserts_MPF_l20_dust, LM_inserts_MPF_l20_pLLF, LM_inserts_MPF_l20_underPFM)
Dim BL_inserts_MPF_l22: BL_inserts_MPF_l22=Array(LM_inserts_MPF_l22_underPFM)
Dim BL_inserts_MPF_l27: BL_inserts_MPF_l27=Array(LM_inserts_MPF_l27_PFMain, LM_inserts_MPF_l27_Parts, LM_inserts_MPF_l27_underPFM)
Dim BL_inserts_MPF_l28: BL_inserts_MPF_l28=Array(LM_inserts_MPF_l28_PFMain, LM_inserts_MPF_l28_Parts, LM_inserts_MPF_l28_underPFM)
Dim BL_inserts_MPF_l29: BL_inserts_MPF_l29=Array(LM_inserts_MPF_l29_PFMain, LM_inserts_MPF_l29_Parts, LM_inserts_MPF_l29_underPFM)
Dim BL_inserts_MPF_l3: BL_inserts_MPF_l3=Array(LM_inserts_MPF_l3_PFMain, LM_inserts_MPF_l3_Parts, LM_inserts_MPF_l3_underPFM)
Dim BL_inserts_MPF_l30: BL_inserts_MPF_l30=Array(LM_inserts_MPF_l30_BSkirt1, LM_inserts_MPF_l30_LPFR1b, LM_inserts_MPF_l30_Layer2, LM_inserts_MPF_l30_PFLower, LM_inserts_MPF_l30_Parts, LM_inserts_MPF_l30_PartsLP, LM_inserts_MPF_l30_pSW20, LM_inserts_MPF_l30_pSW30, LM_inserts_MPF_l30_pSW40, LM_inserts_MPF_l30_pSW60, LM_inserts_MPF_l30_underPFM)
Dim BL_inserts_MPF_l31: BL_inserts_MPF_l31=Array(LM_inserts_MPF_l31_LPFR1b, LM_inserts_MPF_l31_Layer2, LM_inserts_MPF_l31_PFLower, LM_inserts_MPF_l31_PartsLP, LM_inserts_MPF_l31_pLLF, LM_inserts_MPF_l31_pLLFup, LM_inserts_MPF_l31_pSW60, LM_inserts_MPF_l31_underPFM)
Dim BL_inserts_MPF_l32: BL_inserts_MPF_l32=Array(LM_inserts_MPF_l32_underPFM)
Dim BL_inserts_MPF_l33: BL_inserts_MPF_l33=Array(LM_inserts_MPF_l33_Parts, LM_inserts_MPF_l33_underPFM)
Dim BL_inserts_MPF_l34: BL_inserts_MPF_l34=Array(LM_inserts_MPF_l34_Parts, LM_inserts_MPF_l34_underPFM)
Dim BL_inserts_MPF_l35: BL_inserts_MPF_l35=Array(LM_inserts_MPF_l35_PFLower, LM_inserts_MPF_l35_PFMain, LM_inserts_MPF_l35_Parts, LM_inserts_MPF_l35_PartsLP, LM_inserts_MPF_l35_dust, LM_inserts_MPF_l35_pLLF, LM_inserts_MPF_l35_pLLFup, LM_inserts_MPF_l35_pLRF, LM_inserts_MPF_l35_pLRFUup, LM_inserts_MPF_l35_pSWSecret, LM_inserts_MPF_l35_underPFM)
Dim BL_inserts_MPF_l36: BL_inserts_MPF_l36=Array(LM_inserts_MPF_l36_PFLower, LM_inserts_MPF_l36_PFMain, LM_inserts_MPF_l36_Parts, LM_inserts_MPF_l36_PartsLP, LM_inserts_MPF_l36_dust, LM_inserts_MPF_l36_pLLF, LM_inserts_MPF_l36_pLLFup, LM_inserts_MPF_l36_pLRF, LM_inserts_MPF_l36_pLRFUup, LM_inserts_MPF_l36_underPFM)
Dim BL_inserts_MPF_l37: BL_inserts_MPF_l37=Array(LM_inserts_MPF_l37_PFLower, LM_inserts_MPF_l37_PFMain, LM_inserts_MPF_l37_Parts, LM_inserts_MPF_l37_PartsLP, LM_inserts_MPF_l37_dust, LM_inserts_MPF_l37_pLRF, LM_inserts_MPF_l37_pLRFUup, LM_inserts_MPF_l37_underPFM)
Dim BL_inserts_MPF_l38: BL_inserts_MPF_l38=Array(LM_inserts_MPF_l38_PFLower, LM_inserts_MPF_l38_PFMain, LM_inserts_MPF_l38_Parts, LM_inserts_MPF_l38_PartsLP, LM_inserts_MPF_l38_dust, LM_inserts_MPF_l38_pLRF, LM_inserts_MPF_l38_pLRFUup, LM_inserts_MPF_l38_underPFM)
Dim BL_inserts_MPF_l39: BL_inserts_MPF_l39=Array(LM_inserts_MPF_l39_BRing4, LM_inserts_MPF_l39_MPFR1b, LM_inserts_MPF_l39_PFLower, LM_inserts_MPF_l39_PFMain, LM_inserts_MPF_l39_Parts, LM_inserts_MPF_l39_PartsLP, LM_inserts_MPF_l39_USlingPos, LM_inserts_MPF_l39_dust, LM_inserts_MPF_l39_underPFM)
Dim BL_inserts_MPF_l46: BL_inserts_MPF_l46=Array(LM_inserts_MPF_l46_PFLower, LM_inserts_MPF_l46_PFMain, LM_inserts_MPF_l46_Parts, LM_inserts_MPF_l46_psw34, LM_inserts_MPF_l46_underPFM)
Dim BL_inserts_MPF_l47: BL_inserts_MPF_l47=Array(LM_inserts_MPF_l47_MPFR1b, LM_inserts_MPF_l47_PFMain, LM_inserts_MPF_l47_Parts, LM_inserts_MPF_l47_pTrapDoor, LM_inserts_MPF_l47_p_gate3, LM_inserts_MPF_l47_underPFM)
Dim BL_inserts_MPF_l51: BL_inserts_MPF_l51=Array(LM_inserts_MPF_l51_Layer3, LM_inserts_MPF_l51_PFLower, LM_inserts_MPF_l51_PFMain, LM_inserts_MPF_l51_Parts, LM_inserts_MPF_l51_pLRF, LM_inserts_MPF_l51_pLRFUup, LM_inserts_MPF_l51_pSW05, LM_inserts_MPF_l51_pSW15, LM_inserts_MPF_l51_pSWSecret, LM_inserts_MPF_l51_underPFL, LM_inserts_MPF_l51_underPFM)
Dim BL_inserts_MPF_l8: BL_inserts_MPF_l8=Array(LM_inserts_MPF_l8_BSkirt4, LM_inserts_MPF_l8_MPFR1a, LM_inserts_MPF_l8_MPFR1b, LM_inserts_MPF_l8_PFLower, LM_inserts_MPF_l8_PFMain, LM_inserts_MPF_l8_Parts, LM_inserts_MPF_l8_PartsLP, LM_inserts_MPF_l8_pSW55, LM_inserts_MPF_l8_underPFM)
Dim BL_inserts_UPF_l2: BL_inserts_UPF_l2=Array(LM_inserts_UPF_l2_BRing3, LM_inserts_UPF_l2_BSkirt3, LM_inserts_UPF_l2_PFMain, LM_inserts_UPF_l2_PFUpper, LM_inserts_UPF_l2_PURF, LM_inserts_UPF_l2_PURFup, LM_inserts_UPF_l2_Parts, LM_inserts_UPF_l2_pSW03, LM_inserts_UPF_l2_underPFU)
Dim BL_inserts_UPF_l21: BL_inserts_UPF_l21=Array(LM_inserts_UPF_l21_BRing3, LM_inserts_UPF_l21_BSkirt3, LM_inserts_UPF_l21_Layer3, LM_inserts_UPF_l21_PFMain, LM_inserts_UPF_l21_PFUpper, LM_inserts_UPF_l21_PURF, LM_inserts_UPF_l21_PURFup, LM_inserts_UPF_l21_Parts, LM_inserts_UPF_l21_UPFR1a, LM_inserts_UPF_l21_UPFR1b, LM_inserts_UPF_l21_UPFR2c, LM_inserts_UPF_l21_pSW15, LM_inserts_UPF_l21_pSW25, LM_inserts_UPF_l21_pSWSecret, LM_inserts_UPF_l21_pTrapDoor, LM_inserts_UPF_l21_pULF, LM_inserts_UPF_l21_pULFup, LM_inserts_UPF_l21_p_gate2, LM_inserts_UPF_l21_underPFM)
Dim BL_inserts_UPF_l40: BL_inserts_UPF_l40=Array(LM_inserts_UPF_l40_PFMain, LM_inserts_UPF_l40_PFUpper, LM_inserts_UPF_l40_PURF, LM_inserts_UPF_l40_PURFup, LM_inserts_UPF_l40_Parts, LM_inserts_UPF_l40_pSW13, LM_inserts_UPF_l40_underPFU)
Dim BL_inserts_UPF_l41: BL_inserts_UPF_l41=Array(LM_inserts_UPF_l41_PFMain, LM_inserts_UPF_l41_PFUpper, LM_inserts_UPF_l41_PURF, LM_inserts_UPF_l41_PURFup, LM_inserts_UPF_l41_Parts, LM_inserts_UPF_l41_pSW13, LM_inserts_UPF_l41_pSW23, LM_inserts_UPF_l41_pSW33, LM_inserts_UPF_l41_underPFU)
Dim BL_inserts_UPF_l42: BL_inserts_UPF_l42=Array(LM_inserts_UPF_l42_PFMain, LM_inserts_UPF_l42_PFUpper, LM_inserts_UPF_l42_PURF, LM_inserts_UPF_l42_PURFup, LM_inserts_UPF_l42_Parts, LM_inserts_UPF_l42_pSW33, LM_inserts_UPF_l42_underPFU)
Dim BL_inserts_UPF_l44: BL_inserts_UPF_l44=Array(LM_inserts_UPF_l44_underPFU)
' Global arrays
Dim BG_Bakemap: BG_Bakemap=Array(BM_BRing1, BM_BRing2, BM_BRing3, BM_BRing4, BM_BSkirt1, BM_BSkirt2, BM_BSkirt3, BM_BSkirt4, BM_ICLeft, BM_ICRight, BM_LPFR1a, BM_LPFR1b, BM_LPFR2a, BM_LPFR2b, BM_LSlingPos, BM_Layer1, BM_Layer2, BM_Layer3, BM_MPFR1a, BM_MPFR1b, BM_PFLower, BM_PFMain, BM_PFUpper, BM_PLF, BM_PLF2, BM_PLF2up, BM_PLFup, BM_PRF, BM_PRF2, BM_PRF2up, BM_PRFup, BM_PURF, BM_PURFup, BM_Parts, BM_PartsLP, BM_SideRails, BM_TWKicker1, BM_UPFR1a, BM_UPFR1b, BM_UPFR2a, BM_UPFR2b, BM_UPFR2c, BM_UPFR2d, BM_UPFR3a, BM_UPFR3b, BM_UPFR4a, BM_UPFR4b, BM_USlingPos, BM_dust, BM_lockDownBar, BM_pInfoCard, BM_pLLF, BM_pLLFup, BM_pLRF, BM_pLRFUup, BM_pSSpring, BM_pSSpring2, BM_pSW00, BM_pSW02, BM_pSW03, BM_pSW04, BM_pSW05, BM_pSW10, BM_pSW11, BM_pSW12, BM_pSW13, BM_pSW14, BM_pSW15, BM_pSW20, BM_pSW22, BM_pSW23, BM_pSW24, BM_pSW25, BM_pSW30, BM_pSW32, BM_pSW33, BM_pSW40, BM_pSW50, BM_pSW55, BM_pSW60, BM_pSWSecret, BM_pTrapDoor, BM_pULF, BM_pULFup, BM_p_gate1, BM_p_gate2, BM_p_gate3, BM_p_gate4, BM_psw21, BM_psw34, _
  BM_psw54, BM_underPFL, BM_underPFM, BM_underPFU)
Dim BG_Lightmap: BG_Lightmap=Array(LM_BLight_BL1_BRing2, LM_BLight_BL1_BSkirt2, LM_BLight_BL1_PFLower, LM_BLight_BL1_PFMain, LM_BLight_BL1_Parts, LM_BLight_BL1_pSW05, LM_BLight_BL2_BRing1, LM_BLight_BL2_BSkirt1, LM_BLight_BL2_PFLower, LM_BLight_BL2_PFMain, LM_BLight_BL2_PLF2, LM_BLight_BL2_PLF2up, LM_BLight_BL2_Parts, LM_BLight_BL2_PartsLP, LM_BLight_BL2_pSW04, LM_BLight_BL2_pSW14, LM_BLight_BL2_psw34, LM_BLight_BL3_BRing3, LM_BLight_BL3_BSkirt3, LM_BLight_BL3_PFUpper, LM_BLight_BL3_PURF, LM_BLight_BL3_PURFup, LM_BLight_BL3_Parts, LM_BLight_BL3_UPFR2a, LM_BLight_BL3_UPFR2b, LM_BLight_BL3_UPFR2c, LM_BLight_BL3_UPFR2d, LM_BLight_BL3_UPFR3a, LM_BLight_BL3_UPFR3b, LM_BLight_BL3_UPFR4a, LM_BLight_BL3_UPFR4b, LM_BLight_BL3_pSW02, LM_BLight_BL3_pSW12, LM_BLight_BL3_pSW22, LM_BLight_BL3_pSW32, LM_BLight_BL3_underPFU, LM_BLight_BL4_BRing4, LM_BLight_BL4_BSkirt4, LM_BLight_BL4_LPFR2a, LM_BLight_BL4_LPFR2b, LM_BLight_BL4_PFLower, LM_BLight_BL4_PFMain, LM_BLight_BL4_Parts, LM_BLight_BL4_PartsLP, LM_BLight_BL4_pLRF, _
  LM_BLight_BL4_pLRFUup, LM_BLight_BL4_underPFM, LM_GI_LPF_giLL1_BRing4, LM_GI_LPF_giLL1_BSkirt1, LM_GI_LPF_giLL1_BSkirt4, LM_GI_LPF_giLL1_LPFR1a, LM_GI_LPF_giLL1_LPFR1b, LM_GI_LPF_giLL1_LPFR2a, LM_GI_LPF_giLL1_LPFR2b, LM_GI_LPF_giLL1_LSlingPos, LM_GI_LPF_giLL1_Layer2, LM_GI_LPF_giLL1_PFLower, LM_GI_LPF_giLL1_PFMain, LM_GI_LPF_giLL1_Parts, LM_GI_LPF_giLL1_PartsLP, LM_GI_LPF_giLL1_dust, LM_GI_LPF_giLL1_pLLF, LM_GI_LPF_giLL1_pLLFup, LM_GI_LPF_giLL1_pLRF, LM_GI_LPF_giLL1_pLRFUup, LM_GI_LPF_giLL1_pSSpring, LM_GI_LPF_giLL1_pSSpring2, LM_GI_LPF_giLL1_pSW00, LM_GI_LPF_giLL1_pSW04, LM_GI_LPF_giLL1_pSW10, LM_GI_LPF_giLL1_pSW11, LM_GI_LPF_giLL1_pSW20, LM_GI_LPF_giLL1_pSW30, LM_GI_LPF_giLL1_pSW40, LM_GI_LPF_giLL1_pSW50, LM_GI_LPF_giLL1_pSW60, LM_GI_LPF_giLL1_p_gate4, LM_GI_LPF_giLL1_psw21, LM_GI_LPF_giLL1_psw34, LM_GI_LPF_giLL1_underPFL, LM_GI_LPF_giLL1_underPFM, LM_GI_MPF_giML1_BRing2, LM_GI_MPF_giML1_BSkirt1, LM_GI_MPF_giML1_PFLower, LM_GI_MPF_giML1_PFMain, LM_GI_MPF_giML1_Parts, LM_GI_MPF_giML1_PartsLP, _
  LM_GI_MPF_giML1_pSW04, LM_GI_MPF_giML1_pSW05, LM_GI_MPF_giML1_pSW14, LM_GI_MPF_giML1_pSW24, LM_GI_MPF_giML1_pSWSecret, LM_GI_MPF_giML1_underPFM, LM_GI_MPF_giML2_BRing2, LM_GI_MPF_giML2_BSkirt2, LM_GI_MPF_giML2_PFLower, LM_GI_MPF_giML2_PFMain, LM_GI_MPF_giML2_Parts, LM_GI_MPF_giML2_PartsLP, LM_GI_MPF_giML2_pSW05, LM_GI_MPF_giML2_pSW14, LM_GI_MPF_giML2_pSW24, LM_GI_MPF_giML2_pSWSecret, LM_GI_MPF_giML2_underPFM, LM_GI_MPF_giML3_MPFR1a, LM_GI_MPF_giML3_MPFR1b, LM_GI_MPF_giML3_PFLower, LM_GI_MPF_giML3_PFMain, LM_GI_MPF_giML3_PRF2, LM_GI_MPF_giML3_PRF2up, LM_GI_MPF_giML3_Parts, LM_GI_MPF_giML3_PartsLP, LM_GI_MPF_giML3_pSSpring, LM_GI_MPF_giML3_pSW15, LM_GI_MPF_giML3_pSW55, LM_GI_MPF_giML3_p_gate4, LM_GI_MPF_giML3_underPFM, LM_GI_MPF_giML4_MPFR1a, LM_GI_MPF_giML4_MPFR1b, LM_GI_MPF_giML4_PFMain, LM_GI_MPF_giML4_Parts, LM_GI_MPF_giML4_PartsLP, LM_GI_MPF_giML4_USlingPos, LM_GI_MPF_giML4_pSSpring2, LM_GI_MPF_giML4_pTrapDoor, LM_GI_MPF_giML4_p_gate3, LM_GI_MPF_giML4_underPFM, LM_GI_MPF_giML5_PFMain, _
  LM_GI_MPF_giML5_Parts, LM_GI_MPF_giML5_underPFM, LM_GI_MPF_giML6_PFMain, LM_GI_MPF_giML6_Parts, LM_GI_MPF_giML6_p_gate1, LM_GI_UPF_gUL2_BRing3, LM_GI_UPF_gUL2_BSkirt3, LM_GI_UPF_gUL2_PFMain, LM_GI_UPF_gUL2_PFUpper, LM_GI_UPF_gUL2_PURFup, LM_GI_UPF_gUL2_Parts, LM_GI_UPF_gUL2_UPFR1a, LM_GI_UPF_gUL2_UPFR1b, LM_GI_UPF_gUL2_UPFR2a, LM_GI_UPF_gUL2_UPFR2b, LM_GI_UPF_gUL2_UPFR2c, LM_GI_UPF_gUL2_UPFR2d, LM_GI_UPF_gUL2_UPFR3a, LM_GI_UPF_gUL2_UPFR3b, LM_GI_UPF_gUL2_UPFR4a, LM_GI_UPF_gUL2_UPFR4b, LM_GI_UPF_gUL2_pSW02, LM_GI_UPF_gUL2_pSW12, LM_GI_UPF_gUL2_pSW22, LM_GI_UPF_gUL2_pSW32, LM_GI_UPF_gUL2_pULF, LM_GI_UPF_gUL2_pULFup, LM_GI_UPF_gUL2_p_gate2, LM_GI_UPF_giUL1_BRing3, LM_GI_UPF_giUL1_BSkirt3, LM_GI_UPF_giUL1_PFMain, LM_GI_UPF_giUL1_PFUpper, LM_GI_UPF_giUL1_PURFup, LM_GI_UPF_giUL1_Parts, LM_GI_UPF_giUL1_UPFR1a, LM_GI_UPF_giUL1_UPFR1b, LM_GI_UPF_giUL1_UPFR2a, LM_GI_UPF_giUL1_UPFR2b, LM_GI_UPF_giUL1_UPFR2c, LM_GI_UPF_giUL1_UPFR2d, LM_GI_UPF_giUL1_UPFR3a, LM_GI_UPF_giUL1_UPFR3b, LM_GI_UPF_giUL1_UPFR4a, _
  LM_GI_UPF_giUL1_UPFR4b, LM_GI_UPF_giUL1_pSW02, LM_GI_UPF_giUL1_pSW12, LM_GI_UPF_giUL1_pSW22, LM_GI_UPF_giUL1_pSW32, LM_GI_UPF_giUL1_pULF, LM_GI_UPF_giUL1_pULFup, LM_GI_UPF_giUL3_BSkirt3, LM_GI_UPF_giUL3_PFMain, LM_GI_UPF_giUL3_PFUpper, LM_GI_UPF_giUL3_PURF, LM_GI_UPF_giUL3_PURFup, LM_GI_UPF_giUL3_Parts, LM_GI_UPF_giUL3_UPFR1a, LM_GI_UPF_giUL3_UPFR1b, LM_GI_UPF_giUL3_UPFR2a, LM_GI_UPF_giUL3_UPFR2b, LM_GI_UPF_giUL3_UPFR2c, LM_GI_UPF_giUL3_UPFR2d, LM_GI_UPF_giUL3_UPFR3a, LM_GI_UPF_giUL3_UPFR3b, LM_GI_UPF_giUL3_UPFR4a, LM_GI_UPF_giUL3_UPFR4b, LM_GI_UPF_giUL3_pSW02, LM_GI_UPF_giUL3_pSW12, LM_GI_UPF_giUL3_pSW22, LM_GI_UPF_giUL3_pSW32, LM_GI_UPF_giUL3_pULF, LM_GI_UPF_giUL3_pULFup, LM_GI_UPF_giUL4_PFMain, LM_GI_UPF_giUL4_PFUpper, LM_GI_UPF_giUL4_PURF, LM_GI_UPF_giUL4_PURFup, LM_GI_UPF_giUL4_Parts, LM_GI_UPF_giUL4_UPFR4a, LM_GI_UPF_giUL4_UPFR4b, LM_GI_UPF_giUL4_pSW03, LM_GI_UPF_giUL4_pSW13, LM_GI_UPF_giUL4_pSW23, LM_GI_UPF_giUL4_pSW33, LM_GI_UPF_giUL5_PFMain, LM_GI_UPF_giUL5_PFUpper, LM_GI_UPF_giUL5_PURF, _
  LM_GI_UPF_giUL5_PURFup, LM_GI_UPF_giUL5_Parts, LM_GI_UPF_giUL5_SideRails, LM_GI_UPF_giUL5_pSW03, LM_GI_UPF_giUL5_pSW13, LM_GI_UPF_giUL5_pSW23, LM_GI_UPF_giUL5_pSW33, LM_GI_UPF_giUL6_BRing3, LM_GI_UPF_giUL6_BSkirt3, LM_GI_UPF_giUL6_PFMain, LM_GI_UPF_giUL6_PFUpper, LM_GI_UPF_giUL6_PURF, LM_GI_UPF_giUL6_PURFup, LM_GI_UPF_giUL6_Parts, LM_GI_UPF_giUL6_UPFR1a, LM_GI_UPF_giUL6_UPFR1b, LM_GI_UPF_giUL6_UPFR2a, LM_GI_UPF_giUL6_UPFR2b, LM_GI_UPF_giUL6_UPFR2c, LM_GI_UPF_giUL6_UPFR2d, LM_GI_UPF_giUL6_UPFR3a, LM_GI_UPF_giUL6_UPFR3b, LM_GI_UPF_giUL6_pSW02, LM_GI_UPF_giUL6_pSW03, LM_GI_UPF_giUL6_pSW12, LM_GI_UPF_giUL6_pSW13, LM_GI_UPF_giUL6_pSW15, LM_GI_UPF_giUL6_pSW22, LM_GI_UPF_giUL6_pSW25, LM_GI_UPF_giUL6_pSW32, LM_GI_UPF_giUL6_pTrapDoor, LM_GI_UPF_giUL6_pULF, LM_GI_UPF_giUL6_pULFup, LM_GI_UPF_giUL6_p_gate2, LM_GI_UPF_giUL6_underPFM, LM_inserts_LPF_l23_underPFL, LM_inserts_LPF_l24_PFLower, LM_inserts_LPF_l24_Parts, LM_inserts_LPF_l24_underPFL, LM_inserts_LPF_l25_LPFR1b, LM_inserts_LPF_l25_Layer2, _
  LM_inserts_LPF_l25_PFLower, LM_inserts_LPF_l25_Parts, LM_inserts_LPF_l25_PartsLP, LM_inserts_LPF_l25_pSW60, LM_inserts_LPF_l25_underPFL, LM_inserts_LPF_l26_BRing4, LM_inserts_LPF_l26_BSkirt4, LM_inserts_LPF_l26_Parts, LM_inserts_LPF_l26_PartsLP, LM_inserts_LPF_l26_underPFL, LM_inserts_LPF_l45_Parts, LM_inserts_LPF_l45_PartsLP, LM_inserts_LPF_l45_underPFL, LM_inserts_MPF_l18_PFLower, LM_inserts_MPF_l18_PFMain, LM_inserts_MPF_l18_Parts, LM_inserts_MPF_l18_dust, LM_inserts_MPF_l18_pLLF, LM_inserts_MPF_l18_pLLFup, LM_inserts_MPF_l18_pLRF, LM_inserts_MPF_l18_pLRFUup, LM_inserts_MPF_l18_pSW05, LM_inserts_MPF_l18_pSWSecret, LM_inserts_MPF_l18_underPFM, LM_inserts_MPF_l19_PFLower, LM_inserts_MPF_l19_PFMain, LM_inserts_MPF_l19_PartsLP, LM_inserts_MPF_l19_dust, LM_inserts_MPF_l19_pLLF, LM_inserts_MPF_l19_pLLFup, LM_inserts_MPF_l19_pSW05, LM_inserts_MPF_l19_underPFM, LM_inserts_MPF_l20_PFLower, LM_inserts_MPF_l20_PFMain, LM_inserts_MPF_l20_dust, LM_inserts_MPF_l20_pLLF, LM_inserts_MPF_l20_underPFM, _
  LM_inserts_MPF_l22_underPFM, LM_inserts_MPF_l27_PFMain, LM_inserts_MPF_l27_Parts, LM_inserts_MPF_l27_underPFM, LM_inserts_MPF_l28_PFMain, LM_inserts_MPF_l28_Parts, LM_inserts_MPF_l28_underPFM, LM_inserts_MPF_l29_PFMain, LM_inserts_MPF_l29_Parts, LM_inserts_MPF_l29_underPFM, LM_inserts_MPF_l3_PFMain, LM_inserts_MPF_l3_Parts, LM_inserts_MPF_l3_underPFM, LM_inserts_MPF_l30_BSkirt1, LM_inserts_MPF_l30_LPFR1b, LM_inserts_MPF_l30_Layer2, LM_inserts_MPF_l30_PFLower, LM_inserts_MPF_l30_Parts, LM_inserts_MPF_l30_PartsLP, LM_inserts_MPF_l30_pSW20, LM_inserts_MPF_l30_pSW30, LM_inserts_MPF_l30_pSW40, LM_inserts_MPF_l30_pSW60, LM_inserts_MPF_l30_underPFM, LM_inserts_MPF_l31_LPFR1b, LM_inserts_MPF_l31_Layer2, LM_inserts_MPF_l31_PFLower, LM_inserts_MPF_l31_PartsLP, LM_inserts_MPF_l31_pLLF, LM_inserts_MPF_l31_pLLFup, LM_inserts_MPF_l31_pSW60, LM_inserts_MPF_l31_underPFM, LM_inserts_MPF_l32_underPFM, LM_inserts_MPF_l33_Parts, LM_inserts_MPF_l33_underPFM, LM_inserts_MPF_l34_Parts, LM_inserts_MPF_l34_underPFM, _
  LM_inserts_MPF_l35_PFLower, LM_inserts_MPF_l35_PFMain, LM_inserts_MPF_l35_Parts, LM_inserts_MPF_l35_PartsLP, LM_inserts_MPF_l35_dust, LM_inserts_MPF_l35_pLLF, LM_inserts_MPF_l35_pLLFup, LM_inserts_MPF_l35_pLRF, LM_inserts_MPF_l35_pLRFUup, LM_inserts_MPF_l35_pSWSecret, LM_inserts_MPF_l35_underPFM, LM_inserts_MPF_l36_PFLower, LM_inserts_MPF_l36_PFMain, LM_inserts_MPF_l36_Parts, LM_inserts_MPF_l36_PartsLP, LM_inserts_MPF_l36_dust, LM_inserts_MPF_l36_pLLF, LM_inserts_MPF_l36_pLLFup, LM_inserts_MPF_l36_pLRF, LM_inserts_MPF_l36_pLRFUup, LM_inserts_MPF_l36_underPFM, LM_inserts_MPF_l37_PFLower, LM_inserts_MPF_l37_PFMain, LM_inserts_MPF_l37_Parts, LM_inserts_MPF_l37_PartsLP, LM_inserts_MPF_l37_dust, LM_inserts_MPF_l37_pLRF, LM_inserts_MPF_l37_pLRFUup, LM_inserts_MPF_l37_underPFM, LM_inserts_MPF_l38_PFLower, LM_inserts_MPF_l38_PFMain, LM_inserts_MPF_l38_Parts, LM_inserts_MPF_l38_PartsLP, LM_inserts_MPF_l38_dust, LM_inserts_MPF_l38_pLRF, LM_inserts_MPF_l38_pLRFUup, LM_inserts_MPF_l38_underPFM, LM_inserts_MPF_l39_BRing4, _
  LM_inserts_MPF_l39_MPFR1b, LM_inserts_MPF_l39_PFLower, LM_inserts_MPF_l39_PFMain, LM_inserts_MPF_l39_Parts, LM_inserts_MPF_l39_PartsLP, LM_inserts_MPF_l39_USlingPos, LM_inserts_MPF_l39_dust, LM_inserts_MPF_l39_underPFM, LM_inserts_MPF_l46_PFLower, LM_inserts_MPF_l46_PFMain, LM_inserts_MPF_l46_Parts, LM_inserts_MPF_l46_psw34, LM_inserts_MPF_l46_underPFM, LM_inserts_MPF_l47_MPFR1b, LM_inserts_MPF_l47_PFMain, LM_inserts_MPF_l47_Parts, LM_inserts_MPF_l47_pTrapDoor, LM_inserts_MPF_l47_p_gate3, LM_inserts_MPF_l47_underPFM, LM_inserts_MPF_l51_Layer3, LM_inserts_MPF_l51_PFLower, LM_inserts_MPF_l51_PFMain, LM_inserts_MPF_l51_Parts, LM_inserts_MPF_l51_pLRF, LM_inserts_MPF_l51_pLRFUup, LM_inserts_MPF_l51_pSW05, LM_inserts_MPF_l51_pSW15, LM_inserts_MPF_l51_pSWSecret, LM_inserts_MPF_l51_underPFL, LM_inserts_MPF_l51_underPFM, LM_inserts_MPF_l8_BSkirt4, LM_inserts_MPF_l8_MPFR1a, LM_inserts_MPF_l8_MPFR1b, LM_inserts_MPF_l8_PFLower, LM_inserts_MPF_l8_PFMain, LM_inserts_MPF_l8_Parts, LM_inserts_MPF_l8_PartsLP, _
  LM_inserts_MPF_l8_pSW55, LM_inserts_MPF_l8_underPFM, LM_inserts_UPF_l2_BRing3, LM_inserts_UPF_l2_BSkirt3, LM_inserts_UPF_l2_PFMain, LM_inserts_UPF_l2_PFUpper, LM_inserts_UPF_l2_PURF, LM_inserts_UPF_l2_PURFup, LM_inserts_UPF_l2_Parts, LM_inserts_UPF_l2_pSW03, LM_inserts_UPF_l2_underPFU, LM_inserts_UPF_l21_BRing3, LM_inserts_UPF_l21_BSkirt3, LM_inserts_UPF_l21_Layer3, LM_inserts_UPF_l21_PFMain, LM_inserts_UPF_l21_PFUpper, LM_inserts_UPF_l21_PURF, LM_inserts_UPF_l21_PURFup, LM_inserts_UPF_l21_Parts, LM_inserts_UPF_l21_UPFR1a, LM_inserts_UPF_l21_UPFR1b, LM_inserts_UPF_l21_UPFR2c, LM_inserts_UPF_l21_pSW15, LM_inserts_UPF_l21_pSW25, LM_inserts_UPF_l21_pSWSecret, LM_inserts_UPF_l21_pTrapDoor, LM_inserts_UPF_l21_pULF, LM_inserts_UPF_l21_pULFup, LM_inserts_UPF_l21_p_gate2, LM_inserts_UPF_l21_underPFM, LM_inserts_UPF_l40_PFMain, LM_inserts_UPF_l40_PFUpper, LM_inserts_UPF_l40_PURF, LM_inserts_UPF_l40_PURFup, LM_inserts_UPF_l40_Parts, LM_inserts_UPF_l40_pSW13, LM_inserts_UPF_l40_underPFU, LM_inserts_UPF_l41_PFMain, _
  LM_inserts_UPF_l41_PFUpper, LM_inserts_UPF_l41_PURF, LM_inserts_UPF_l41_PURFup, LM_inserts_UPF_l41_Parts, LM_inserts_UPF_l41_pSW13, LM_inserts_UPF_l41_pSW23, LM_inserts_UPF_l41_pSW33, LM_inserts_UPF_l41_underPFU, LM_inserts_UPF_l42_PFMain, LM_inserts_UPF_l42_PFUpper, LM_inserts_UPF_l42_PURF, LM_inserts_UPF_l42_PURFup, LM_inserts_UPF_l42_Parts, LM_inserts_UPF_l42_pSW33, LM_inserts_UPF_l42_underPFU, LM_inserts_UPF_l44_underPFU)
Dim BG_All: BG_All=Array(BM_BRing1, BM_BRing2, BM_BRing3, BM_BRing4, BM_BSkirt1, BM_BSkirt2, BM_BSkirt3, BM_BSkirt4, BM_ICLeft, BM_ICRight, BM_LPFR1a, BM_LPFR1b, BM_LPFR2a, BM_LPFR2b, BM_LSlingPos, BM_Layer1, BM_Layer2, BM_Layer3, BM_MPFR1a, BM_MPFR1b, BM_PFLower, BM_PFMain, BM_PFUpper, BM_PLF, BM_PLF2, BM_PLF2up, BM_PLFup, BM_PRF, BM_PRF2, BM_PRF2up, BM_PRFup, BM_PURF, BM_PURFup, BM_Parts, BM_PartsLP, BM_SideRails, BM_TWKicker1, BM_UPFR1a, BM_UPFR1b, BM_UPFR2a, BM_UPFR2b, BM_UPFR2c, BM_UPFR2d, BM_UPFR3a, BM_UPFR3b, BM_UPFR4a, BM_UPFR4b, BM_USlingPos, BM_dust, BM_lockDownBar, BM_pInfoCard, BM_pLLF, BM_pLLFup, BM_pLRF, BM_pLRFUup, BM_pSSpring, BM_pSSpring2, BM_pSW00, BM_pSW02, BM_pSW03, BM_pSW04, BM_pSW05, BM_pSW10, BM_pSW11, BM_pSW12, BM_pSW13, BM_pSW14, BM_pSW15, BM_pSW20, BM_pSW22, BM_pSW23, BM_pSW24, BM_pSW25, BM_pSW30, BM_pSW32, BM_pSW33, BM_pSW40, BM_pSW50, BM_pSW55, BM_pSW60, BM_pSWSecret, BM_pTrapDoor, BM_pULF, BM_pULFup, BM_p_gate1, BM_p_gate2, BM_p_gate3, BM_p_gate4, BM_psw21, BM_psw34, BM_psw54, _
  BM_underPFL, BM_underPFM, BM_underPFU, LM_BLight_BL1_BRing2, LM_BLight_BL1_BSkirt2, LM_BLight_BL1_PFLower, LM_BLight_BL1_PFMain, LM_BLight_BL1_Parts, LM_BLight_BL1_pSW05, LM_BLight_BL2_BRing1, LM_BLight_BL2_BSkirt1, LM_BLight_BL2_PFLower, LM_BLight_BL2_PFMain, LM_BLight_BL2_PLF2, LM_BLight_BL2_PLF2up, LM_BLight_BL2_Parts, LM_BLight_BL2_PartsLP, LM_BLight_BL2_pSW04, LM_BLight_BL2_pSW14, LM_BLight_BL2_psw34, LM_BLight_BL3_BRing3, LM_BLight_BL3_BSkirt3, LM_BLight_BL3_PFUpper, LM_BLight_BL3_PURF, LM_BLight_BL3_PURFup, LM_BLight_BL3_Parts, LM_BLight_BL3_UPFR2a, LM_BLight_BL3_UPFR2b, LM_BLight_BL3_UPFR2c, LM_BLight_BL3_UPFR2d, LM_BLight_BL3_UPFR3a, LM_BLight_BL3_UPFR3b, LM_BLight_BL3_UPFR4a, LM_BLight_BL3_UPFR4b, LM_BLight_BL3_pSW02, LM_BLight_BL3_pSW12, LM_BLight_BL3_pSW22, LM_BLight_BL3_pSW32, LM_BLight_BL3_underPFU, LM_BLight_BL4_BRing4, LM_BLight_BL4_BSkirt4, LM_BLight_BL4_LPFR2a, LM_BLight_BL4_LPFR2b, LM_BLight_BL4_PFLower, LM_BLight_BL4_PFMain, LM_BLight_BL4_Parts, LM_BLight_BL4_PartsLP, LM_BLight_BL4_pLRF, _
  LM_BLight_BL4_pLRFUup, LM_BLight_BL4_underPFM, LM_GI_LPF_giLL1_BRing4, LM_GI_LPF_giLL1_BSkirt1, LM_GI_LPF_giLL1_BSkirt4, LM_GI_LPF_giLL1_LPFR1a, LM_GI_LPF_giLL1_LPFR1b, LM_GI_LPF_giLL1_LPFR2a, LM_GI_LPF_giLL1_LPFR2b, LM_GI_LPF_giLL1_LSlingPos, LM_GI_LPF_giLL1_Layer2, LM_GI_LPF_giLL1_PFLower, LM_GI_LPF_giLL1_PFMain, LM_GI_LPF_giLL1_Parts, LM_GI_LPF_giLL1_PartsLP, LM_GI_LPF_giLL1_dust, LM_GI_LPF_giLL1_pLLF, LM_GI_LPF_giLL1_pLLFup, LM_GI_LPF_giLL1_pLRF, LM_GI_LPF_giLL1_pLRFUup, LM_GI_LPF_giLL1_pSSpring, LM_GI_LPF_giLL1_pSSpring2, LM_GI_LPF_giLL1_pSW00, LM_GI_LPF_giLL1_pSW04, LM_GI_LPF_giLL1_pSW10, LM_GI_LPF_giLL1_pSW11, LM_GI_LPF_giLL1_pSW20, LM_GI_LPF_giLL1_pSW30, LM_GI_LPF_giLL1_pSW40, LM_GI_LPF_giLL1_pSW50, LM_GI_LPF_giLL1_pSW60, LM_GI_LPF_giLL1_p_gate4, LM_GI_LPF_giLL1_psw21, LM_GI_LPF_giLL1_psw34, LM_GI_LPF_giLL1_underPFL, LM_GI_LPF_giLL1_underPFM, LM_GI_MPF_giML1_BRing2, LM_GI_MPF_giML1_BSkirt1, LM_GI_MPF_giML1_PFLower, LM_GI_MPF_giML1_PFMain, LM_GI_MPF_giML1_Parts, LM_GI_MPF_giML1_PartsLP, _
  LM_GI_MPF_giML1_pSW04, LM_GI_MPF_giML1_pSW05, LM_GI_MPF_giML1_pSW14, LM_GI_MPF_giML1_pSW24, LM_GI_MPF_giML1_pSWSecret, LM_GI_MPF_giML1_underPFM, LM_GI_MPF_giML2_BRing2, LM_GI_MPF_giML2_BSkirt2, LM_GI_MPF_giML2_PFLower, LM_GI_MPF_giML2_PFMain, LM_GI_MPF_giML2_Parts, LM_GI_MPF_giML2_PartsLP, LM_GI_MPF_giML2_pSW05, LM_GI_MPF_giML2_pSW14, LM_GI_MPF_giML2_pSW24, LM_GI_MPF_giML2_pSWSecret, LM_GI_MPF_giML2_underPFM, LM_GI_MPF_giML3_MPFR1a, LM_GI_MPF_giML3_MPFR1b, LM_GI_MPF_giML3_PFLower, LM_GI_MPF_giML3_PFMain, LM_GI_MPF_giML3_PRF2, LM_GI_MPF_giML3_PRF2up, LM_GI_MPF_giML3_Parts, LM_GI_MPF_giML3_PartsLP, LM_GI_MPF_giML3_pSSpring, LM_GI_MPF_giML3_pSW15, LM_GI_MPF_giML3_pSW55, LM_GI_MPF_giML3_p_gate4, LM_GI_MPF_giML3_underPFM, LM_GI_MPF_giML4_MPFR1a, LM_GI_MPF_giML4_MPFR1b, LM_GI_MPF_giML4_PFMain, LM_GI_MPF_giML4_Parts, LM_GI_MPF_giML4_PartsLP, LM_GI_MPF_giML4_USlingPos, LM_GI_MPF_giML4_pSSpring2, LM_GI_MPF_giML4_pTrapDoor, LM_GI_MPF_giML4_p_gate3, LM_GI_MPF_giML4_underPFM, LM_GI_MPF_giML5_PFMain, _
  LM_GI_MPF_giML5_Parts, LM_GI_MPF_giML5_underPFM, LM_GI_MPF_giML6_PFMain, LM_GI_MPF_giML6_Parts, LM_GI_MPF_giML6_p_gate1, LM_GI_UPF_gUL2_BRing3, LM_GI_UPF_gUL2_BSkirt3, LM_GI_UPF_gUL2_PFMain, LM_GI_UPF_gUL2_PFUpper, LM_GI_UPF_gUL2_PURFup, LM_GI_UPF_gUL2_Parts, LM_GI_UPF_gUL2_UPFR1a, LM_GI_UPF_gUL2_UPFR1b, LM_GI_UPF_gUL2_UPFR2a, LM_GI_UPF_gUL2_UPFR2b, LM_GI_UPF_gUL2_UPFR2c, LM_GI_UPF_gUL2_UPFR2d, LM_GI_UPF_gUL2_UPFR3a, LM_GI_UPF_gUL2_UPFR3b, LM_GI_UPF_gUL2_UPFR4a, LM_GI_UPF_gUL2_UPFR4b, LM_GI_UPF_gUL2_pSW02, LM_GI_UPF_gUL2_pSW12, LM_GI_UPF_gUL2_pSW22, LM_GI_UPF_gUL2_pSW32, LM_GI_UPF_gUL2_pULF, LM_GI_UPF_gUL2_pULFup, LM_GI_UPF_gUL2_p_gate2, LM_GI_UPF_giUL1_BRing3, LM_GI_UPF_giUL1_BSkirt3, LM_GI_UPF_giUL1_PFMain, LM_GI_UPF_giUL1_PFUpper, LM_GI_UPF_giUL1_PURFup, LM_GI_UPF_giUL1_Parts, LM_GI_UPF_giUL1_UPFR1a, LM_GI_UPF_giUL1_UPFR1b, LM_GI_UPF_giUL1_UPFR2a, LM_GI_UPF_giUL1_UPFR2b, LM_GI_UPF_giUL1_UPFR2c, LM_GI_UPF_giUL1_UPFR2d, LM_GI_UPF_giUL1_UPFR3a, LM_GI_UPF_giUL1_UPFR3b, LM_GI_UPF_giUL1_UPFR4a, _
  LM_GI_UPF_giUL1_UPFR4b, LM_GI_UPF_giUL1_pSW02, LM_GI_UPF_giUL1_pSW12, LM_GI_UPF_giUL1_pSW22, LM_GI_UPF_giUL1_pSW32, LM_GI_UPF_giUL1_pULF, LM_GI_UPF_giUL1_pULFup, LM_GI_UPF_giUL3_BSkirt3, LM_GI_UPF_giUL3_PFMain, LM_GI_UPF_giUL3_PFUpper, LM_GI_UPF_giUL3_PURF, LM_GI_UPF_giUL3_PURFup, LM_GI_UPF_giUL3_Parts, LM_GI_UPF_giUL3_UPFR1a, LM_GI_UPF_giUL3_UPFR1b, LM_GI_UPF_giUL3_UPFR2a, LM_GI_UPF_giUL3_UPFR2b, LM_GI_UPF_giUL3_UPFR2c, LM_GI_UPF_giUL3_UPFR2d, LM_GI_UPF_giUL3_UPFR3a, LM_GI_UPF_giUL3_UPFR3b, LM_GI_UPF_giUL3_UPFR4a, LM_GI_UPF_giUL3_UPFR4b, LM_GI_UPF_giUL3_pSW02, LM_GI_UPF_giUL3_pSW12, LM_GI_UPF_giUL3_pSW22, LM_GI_UPF_giUL3_pSW32, LM_GI_UPF_giUL3_pULF, LM_GI_UPF_giUL3_pULFup, LM_GI_UPF_giUL4_PFMain, LM_GI_UPF_giUL4_PFUpper, LM_GI_UPF_giUL4_PURF, LM_GI_UPF_giUL4_PURFup, LM_GI_UPF_giUL4_Parts, LM_GI_UPF_giUL4_UPFR4a, LM_GI_UPF_giUL4_UPFR4b, LM_GI_UPF_giUL4_pSW03, LM_GI_UPF_giUL4_pSW13, LM_GI_UPF_giUL4_pSW23, LM_GI_UPF_giUL4_pSW33, LM_GI_UPF_giUL5_PFMain, LM_GI_UPF_giUL5_PFUpper, LM_GI_UPF_giUL5_PURF, _
  LM_GI_UPF_giUL5_PURFup, LM_GI_UPF_giUL5_Parts, LM_GI_UPF_giUL5_SideRails, LM_GI_UPF_giUL5_pSW03, LM_GI_UPF_giUL5_pSW13, LM_GI_UPF_giUL5_pSW23, LM_GI_UPF_giUL5_pSW33, LM_GI_UPF_giUL6_BRing3, LM_GI_UPF_giUL6_BSkirt3, LM_GI_UPF_giUL6_PFMain, LM_GI_UPF_giUL6_PFUpper, LM_GI_UPF_giUL6_PURF, LM_GI_UPF_giUL6_PURFup, LM_GI_UPF_giUL6_Parts, LM_GI_UPF_giUL6_UPFR1a, LM_GI_UPF_giUL6_UPFR1b, LM_GI_UPF_giUL6_UPFR2a, LM_GI_UPF_giUL6_UPFR2b, LM_GI_UPF_giUL6_UPFR2c, LM_GI_UPF_giUL6_UPFR2d, LM_GI_UPF_giUL6_UPFR3a, LM_GI_UPF_giUL6_UPFR3b, LM_GI_UPF_giUL6_pSW02, LM_GI_UPF_giUL6_pSW03, LM_GI_UPF_giUL6_pSW12, LM_GI_UPF_giUL6_pSW13, LM_GI_UPF_giUL6_pSW15, LM_GI_UPF_giUL6_pSW22, LM_GI_UPF_giUL6_pSW25, LM_GI_UPF_giUL6_pSW32, LM_GI_UPF_giUL6_pTrapDoor, LM_GI_UPF_giUL6_pULF, LM_GI_UPF_giUL6_pULFup, LM_GI_UPF_giUL6_p_gate2, LM_GI_UPF_giUL6_underPFM, LM_inserts_LPF_l23_underPFL, LM_inserts_LPF_l24_PFLower, LM_inserts_LPF_l24_Parts, LM_inserts_LPF_l24_underPFL, LM_inserts_LPF_l25_LPFR1b, LM_inserts_LPF_l25_Layer2, _
  LM_inserts_LPF_l25_PFLower, LM_inserts_LPF_l25_Parts, LM_inserts_LPF_l25_PartsLP, LM_inserts_LPF_l25_pSW60, LM_inserts_LPF_l25_underPFL, LM_inserts_LPF_l26_BRing4, LM_inserts_LPF_l26_BSkirt4, LM_inserts_LPF_l26_Parts, LM_inserts_LPF_l26_PartsLP, LM_inserts_LPF_l26_underPFL, LM_inserts_LPF_l45_Parts, LM_inserts_LPF_l45_PartsLP, LM_inserts_LPF_l45_underPFL, LM_inserts_MPF_l18_PFLower, LM_inserts_MPF_l18_PFMain, LM_inserts_MPF_l18_Parts, LM_inserts_MPF_l18_dust, LM_inserts_MPF_l18_pLLF, LM_inserts_MPF_l18_pLLFup, LM_inserts_MPF_l18_pLRF, LM_inserts_MPF_l18_pLRFUup, LM_inserts_MPF_l18_pSW05, LM_inserts_MPF_l18_pSWSecret, LM_inserts_MPF_l18_underPFM, LM_inserts_MPF_l19_PFLower, LM_inserts_MPF_l19_PFMain, LM_inserts_MPF_l19_PartsLP, LM_inserts_MPF_l19_dust, LM_inserts_MPF_l19_pLLF, LM_inserts_MPF_l19_pLLFup, LM_inserts_MPF_l19_pSW05, LM_inserts_MPF_l19_underPFM, LM_inserts_MPF_l20_PFLower, LM_inserts_MPF_l20_PFMain, LM_inserts_MPF_l20_dust, LM_inserts_MPF_l20_pLLF, LM_inserts_MPF_l20_underPFM, _
  LM_inserts_MPF_l22_underPFM, LM_inserts_MPF_l27_PFMain, LM_inserts_MPF_l27_Parts, LM_inserts_MPF_l27_underPFM, LM_inserts_MPF_l28_PFMain, LM_inserts_MPF_l28_Parts, LM_inserts_MPF_l28_underPFM, LM_inserts_MPF_l29_PFMain, LM_inserts_MPF_l29_Parts, LM_inserts_MPF_l29_underPFM, LM_inserts_MPF_l3_PFMain, LM_inserts_MPF_l3_Parts, LM_inserts_MPF_l3_underPFM, LM_inserts_MPF_l30_BSkirt1, LM_inserts_MPF_l30_LPFR1b, LM_inserts_MPF_l30_Layer2, LM_inserts_MPF_l30_PFLower, LM_inserts_MPF_l30_Parts, LM_inserts_MPF_l30_PartsLP, LM_inserts_MPF_l30_pSW20, LM_inserts_MPF_l30_pSW30, LM_inserts_MPF_l30_pSW40, LM_inserts_MPF_l30_pSW60, LM_inserts_MPF_l30_underPFM, LM_inserts_MPF_l31_LPFR1b, LM_inserts_MPF_l31_Layer2, LM_inserts_MPF_l31_PFLower, LM_inserts_MPF_l31_PartsLP, LM_inserts_MPF_l31_pLLF, LM_inserts_MPF_l31_pLLFup, LM_inserts_MPF_l31_pSW60, LM_inserts_MPF_l31_underPFM, LM_inserts_MPF_l32_underPFM, LM_inserts_MPF_l33_Parts, LM_inserts_MPF_l33_underPFM, LM_inserts_MPF_l34_Parts, LM_inserts_MPF_l34_underPFM, _
  LM_inserts_MPF_l35_PFLower, LM_inserts_MPF_l35_PFMain, LM_inserts_MPF_l35_Parts, LM_inserts_MPF_l35_PartsLP, LM_inserts_MPF_l35_dust, LM_inserts_MPF_l35_pLLF, LM_inserts_MPF_l35_pLLFup, LM_inserts_MPF_l35_pLRF, LM_inserts_MPF_l35_pLRFUup, LM_inserts_MPF_l35_pSWSecret, LM_inserts_MPF_l35_underPFM, LM_inserts_MPF_l36_PFLower, LM_inserts_MPF_l36_PFMain, LM_inserts_MPF_l36_Parts, LM_inserts_MPF_l36_PartsLP, LM_inserts_MPF_l36_dust, LM_inserts_MPF_l36_pLLF, LM_inserts_MPF_l36_pLLFup, LM_inserts_MPF_l36_pLRF, LM_inserts_MPF_l36_pLRFUup, LM_inserts_MPF_l36_underPFM, LM_inserts_MPF_l37_PFLower, LM_inserts_MPF_l37_PFMain, LM_inserts_MPF_l37_Parts, LM_inserts_MPF_l37_PartsLP, LM_inserts_MPF_l37_dust, LM_inserts_MPF_l37_pLRF, LM_inserts_MPF_l37_pLRFUup, LM_inserts_MPF_l37_underPFM, LM_inserts_MPF_l38_PFLower, LM_inserts_MPF_l38_PFMain, LM_inserts_MPF_l38_Parts, LM_inserts_MPF_l38_PartsLP, LM_inserts_MPF_l38_dust, LM_inserts_MPF_l38_pLRF, LM_inserts_MPF_l38_pLRFUup, LM_inserts_MPF_l38_underPFM, LM_inserts_MPF_l39_BRing4, _
  LM_inserts_MPF_l39_MPFR1b, LM_inserts_MPF_l39_PFLower, LM_inserts_MPF_l39_PFMain, LM_inserts_MPF_l39_Parts, LM_inserts_MPF_l39_PartsLP, LM_inserts_MPF_l39_USlingPos, LM_inserts_MPF_l39_dust, LM_inserts_MPF_l39_underPFM, LM_inserts_MPF_l46_PFLower, LM_inserts_MPF_l46_PFMain, LM_inserts_MPF_l46_Parts, LM_inserts_MPF_l46_psw34, LM_inserts_MPF_l46_underPFM, LM_inserts_MPF_l47_MPFR1b, LM_inserts_MPF_l47_PFMain, LM_inserts_MPF_l47_Parts, LM_inserts_MPF_l47_pTrapDoor, LM_inserts_MPF_l47_p_gate3, LM_inserts_MPF_l47_underPFM, LM_inserts_MPF_l51_Layer3, LM_inserts_MPF_l51_PFLower, LM_inserts_MPF_l51_PFMain, LM_inserts_MPF_l51_Parts, LM_inserts_MPF_l51_pLRF, LM_inserts_MPF_l51_pLRFUup, LM_inserts_MPF_l51_pSW05, LM_inserts_MPF_l51_pSW15, LM_inserts_MPF_l51_pSWSecret, LM_inserts_MPF_l51_underPFL, LM_inserts_MPF_l51_underPFM, LM_inserts_MPF_l8_BSkirt4, LM_inserts_MPF_l8_MPFR1a, LM_inserts_MPF_l8_MPFR1b, LM_inserts_MPF_l8_PFLower, LM_inserts_MPF_l8_PFMain, LM_inserts_MPF_l8_Parts, LM_inserts_MPF_l8_PartsLP, _
  LM_inserts_MPF_l8_pSW55, LM_inserts_MPF_l8_underPFM, LM_inserts_UPF_l2_BRing3, LM_inserts_UPF_l2_BSkirt3, LM_inserts_UPF_l2_PFMain, LM_inserts_UPF_l2_PFUpper, LM_inserts_UPF_l2_PURF, LM_inserts_UPF_l2_PURFup, LM_inserts_UPF_l2_Parts, LM_inserts_UPF_l2_pSW03, LM_inserts_UPF_l2_underPFU, LM_inserts_UPF_l21_BRing3, LM_inserts_UPF_l21_BSkirt3, LM_inserts_UPF_l21_Layer3, LM_inserts_UPF_l21_PFMain, LM_inserts_UPF_l21_PFUpper, LM_inserts_UPF_l21_PURF, LM_inserts_UPF_l21_PURFup, LM_inserts_UPF_l21_Parts, LM_inserts_UPF_l21_UPFR1a, LM_inserts_UPF_l21_UPFR1b, LM_inserts_UPF_l21_UPFR2c, LM_inserts_UPF_l21_pSW15, LM_inserts_UPF_l21_pSW25, LM_inserts_UPF_l21_pSWSecret, LM_inserts_UPF_l21_pTrapDoor, LM_inserts_UPF_l21_pULF, LM_inserts_UPF_l21_pULFup, LM_inserts_UPF_l21_p_gate2, LM_inserts_UPF_l21_underPFM, LM_inserts_UPF_l40_PFMain, LM_inserts_UPF_l40_PFUpper, LM_inserts_UPF_l40_PURF, LM_inserts_UPF_l40_PURFup, LM_inserts_UPF_l40_Parts, LM_inserts_UPF_l40_pSW13, LM_inserts_UPF_l40_underPFU, LM_inserts_UPF_l41_PFMain, _
  LM_inserts_UPF_l41_PFUpper, LM_inserts_UPF_l41_PURF, LM_inserts_UPF_l41_PURFup, LM_inserts_UPF_l41_Parts, LM_inserts_UPF_l41_pSW13, LM_inserts_UPF_l41_pSW23, LM_inserts_UPF_l41_pSW33, LM_inserts_UPF_l41_underPFU, LM_inserts_UPF_l42_PFMain, LM_inserts_UPF_l42_PFUpper, LM_inserts_UPF_l42_PURF, LM_inserts_UPF_l42_PURFup, LM_inserts_UPF_l42_Parts, LM_inserts_UPF_l42_pSW33, LM_inserts_UPF_l42_underPFU, LM_inserts_UPF_l44_underPFU)
' VLM Arrays - End



'*******************************************
' ZLOA: Load Stuff
'*******************************************

'Check the selected ROM version
Dim cGameName

If RomSet = 1 then
  cGameName="hh"
  If HH.ShowFSS = True or RenderingMode = 2 then
    DisplayTimerFSS6.Enabled = True
    DisplayTimer7.Enabled = False
    DisplayTimer6.Enabled = False
  Else
    DisplayTimer6.Enabled = true
  End If
End If
If RomSet = 2 then
  cGameName="hh7"
  If HH.ShowFSS = True or RenderingMode = 2 then
    DisplayTimerFSS7.Enabled = True
    DisplayTimer7.Enabled = False
    DisplayTimer6.Enabled = False
  Else
    DisplayTimer7.Enabled = true
  End If
End If

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the Controller.vbs file in order to run this table (installed with the VPX package in the scripts folder)"
On Error Goto 0


Const UseSolenoids    = 2
Const UseLamps      = 1
Const UseGI         = 0
Const UseSync       = 1
Const HandleMech    = 0
Const SSolenoidOn     = "SOL_On"
Const SSolenoidOff    = "SOL_Off"
Const SCoin       = ""
'Const SKnocker       = "Knocker"


' load game controller
LoadVPM "01560000", "sys80.VBS", 3.37



'*******************************************
' ZTIM: Main Timers
'*******************************************

' FIXME for the time being, the cor timer interval must be 10 ms (so below 60FPS framerate)
CorTimer.Interval = 10
Sub CorTimer_Timer(): Cor.Update: End Sub

Sub FrameTimer_Timer() 'The frame timer interval should be -1, so executes at the display frame rate
  BSUpdate
  AnimateBumperSkirts
  UpdateBallBrightness
  RollingUpdate   'update rolling sounds
  DoSTAnim    'handle stand up target animations
  UpdateStandupTargets
  DoDTAnim    'handle drop up target animations
  UpdateDropTargets
  Options_UpdateDMD
End Sub

'*******************************************
' ZINI: Table Initialization and Exiting
'*******************************************

Sub HH_Init
  ' table initialization
  vpmInit Me
  With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
    .SplashInfoLine = "Haunted House" & vbNewLine & "Gottlieb 1982" & vbNewLine & "VPW"
    '.Settings.Value("dmd_red") = 0
    '.Settings.Value("dmd_green") = 194
    '.Settings.Value("dmd_blue") = 0
    '.Games(cGameName).Settings.Value("rol")=0
    .Games(cGameName).Settings.value("sound") = 1 ' - Test table sounds...  disables ROM sounds
    .HandleKeyboard = 0
    .ShowTitle = 0
    .ShowDMDOnly = 1
    .ShowFrame = 0
    .HandleMechanics = 0
    .Hidden = 1
    On Error Resume Next
    Controller.SolMask(0) = 0
    vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - Then add the timer to renable all the solenoids after 2 seconds
    .Run GetPlayerHWnd
    If Err Then MsgBox Err.Description
    On Error Goto 0
  End With

  vpmMapLights AllLamps ' Map all lamps to the corresponding ROM output using the value of TimerInterval of each light object

  ' basic pinmame timer
  PinMAMETimer.Interval = PinMAMEInterval
  PinMAMETimer.Enabled  = True

  ' nudging
  vpmNudge.TiltSwitch   = 57
  vpmNudge.Sensitivity  = 3
  vpmNudge.TiltObj    = Array(Bumper1,Bumper2,Bumper3,Bumper4,UpperSlingShot,LowerSlingShot)

'     ' Upper Drop Targets
'     Set udtDrop = new cvpmDropTarget
'     With udtDrop
'       .Initdrop Array(sw02,sw12,sw22,sw32), Array(2,12,22,32)
''        .InitSnd "fx_resetdrop","fx_droptarget"
'      End With
'
'     ' Lower Drop Targets
'     Set ldtDrop = new cvpmDropTarget
'     With ldtDrop
'       .Initdrop Array(sw00,sw10,sw20,sw30,sw40), Array(0,10,20,30,40)
''        .InitSnd "fx_resetdrop","fx_droptarget"
'      End With

  Const IMPowerSetting = 25 'Right Side Help Kicker Power
  Const IMTime = 0.3        ' Time in seconds for Full Plunge
  Set plungerIM = New cvpmImpulseP
        plungerIM.InitImpulseP sw65, IMPowerSetting, IMTime
        plungerIM.Random 0.3
        plungerIM.switch 65
        plungerIM.InitExitSnd SoundFX("SOL_On",DOFContactors), SoundFX("SOL_Off",DOFContactors)
        plungerIM.CreateEvents "plungerIM"

  'Ball initializations need for physical trough
  Set HHBall = Kicker1.CreateSizedballWithMass(Ballsize / 2,Ballmass)
  gBOT = Array(HHBall)
  Controller.Switch(46) = 1

  ' Other initializations
  Backdrop_Init

  ' Load Options
  Options_Load

  ' Room Brightness
  SetRoomBrightness LightLevel/100

  ' Slingshot init
  LowerSlingshotStep = 0
  LowerSlingshot.TimerEnabled = 1
  UpperSlingshotStep = 0
  UpperSlingShot.TimerEnabled = 1

End Sub


Sub HH_Exit
  Controller.Stop
End Sub


'  Setup Desktop
Sub Backdrop_Init
  Dim bdl
  If DesktopMode = True then
    l24.Y = 1060
    l25.Y = 1060
    l26.Y = 1060
    l25b.Y = 1141.842
    l23.Y = 1184.212
    l26b.Y = 1303.198
    l45.Y = 1415.443
    For each bdl in BackdropLights: bdl.visible = true:Next
  Else
    l24.Y = 1066
    l25.Y = 1066
    l26.Y = 1066
    l25b.Y = 1153.842
    l23.Y = 1196.212
    l26b.Y = 1317.198
    l45.Y = 1428.776
    For each bdl in BackdropLights: bdl.visible = false:Next
  End If
End Sub



'******************************************************
'   ZBBR: BALL BRIGHTNESS
'******************************************************

Const BallBrightness =  1        'Ball brightness - Value between 0 and 1 (0=Dark ... 1=Bright)
Const LPFGIGain = 0.2
'Const PLOffset = 0.2
'Dim PLGain: PLGain = (1-PLOffset)/(1260-2000)

Sub UpdateBallBrightness
  Dim s, b_base, b_r, b_g, b_b, d_w
  b_base = 90 * BallBrightness + 65*LightLevel/100 + 100  ' orig was 120 and 70

  For s = 0 To UBound(gBOT)
    d_w = b_base
    ' Handle z direction and lower PG gi state
    If gBOT(s).z < 0 Then
      d_w = d_w * (LPFGIGain+(1-LPFGIGain)*GIstate)
    End If
    If d_w < 30 Then d_w = 30
'   ' Handle plunger lane
'   If InRect(gBOT(s).x,gBOT(s).y,870,2000,870,1260,930,1260,930,2000) Then
'     d_w = d_w*(PLOffset+PLGain*(gBOT(s).y-2000))
'   End If
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
'   ZBRI: Room Brightness
'****************************


' Update these arrays if you want to change more materials with room light level
Dim RoomBrightnessMtlArray: RoomBrightnessMtlArray = Array("VLM.Bake.Active","VLM.Bake.Solid","VLM.Bake.ActiveLPF","VLM.Bake.SolidLPF", _
    "_noXtraShading","_noXtraShadingVR","_noXtraShadingVRTrans","_noXtraShadingBlackGlossy","VR_Metal","VR_ClearRed","VR_WhitePlastic", _
    "VR_ClearGreen",  "VR_Stone",  "VR_DeadTree",  "VR_Plank_Fence",  "VR_House",  "VR_Plastic_Gray_Dark",  "Plastic with an image")

Dim SavedMtlColorArray:     SavedMtlColorArray     = Array(0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0)


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

SaveMtlColors
Sub SaveMtlColors
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
'   ZDIP: DIP Switches
'*******************************************


'Inkochnito Dip switches
Set vpmShowDips = GetRef("editDips")

Sub editDips
  Dim vpmDips:Set vpmDips = New cvpmDips
  With vpmDips
    .AddForm 700, 400, "Haunted House - DIP switches"
    .AddFrame 2, 2, 190, "Maximum credits", 49152, Array("8 credits", 0, "10 credits", 32768, "15 credits", &H00004000, "25 credits", 49152)                                                                                   'dip 15&16
    .AddFrame 2, 78, 190, "Coin chute 1 and 2 control", &H00002000, Array("Seperate", 0, "Same", &H00002000)                                                                                                                   'dip 14
    .AddFrame 2, 124, 190, "Playfield special", &H00200000, Array("Replay", 0, "Extra Ball", &H00200000)                                                                                                                       'dip 22
    .AddFrame 2, 170, 190, "Tilt penalty", &H10000000, Array("game over", 0, "ball in play only", &H10000000)                                                                                                                  'dip 29
    .AddFrame 2, 216, 190, "High score to date awards", &H00C00000, Array("not displayed and no award", 0, "displayed and no award", &H00800000, "displayed and 2 credits", &H00400000, "displayed and 3 credits", &H00C00000) 'dip 23&24
    .AddFrameExtra 205, 2, 190, "Attract Sound", &H000c, Array("Off", 0, "every 10 seconds", &H0004, "every 2 minutes", &H0008, "every 4 minutes", &H000C)                                                                     'sounddip 3&4
    .AddFrame 205, 78, 190, "Balls per game", &H00010000, Array("5 balls", 0, "3 balls", &H00010000)                                                                                                                           'dip 17
    .AddFrame 205, 124, 190, "Replay limit", &H00040000, Array("No limit", 0, "One per ball", &H00040000)                                                                                                                      'dip 19
    .AddFrame 205, 170, 190, "Novelty mode", &H00080000, Array("Normal game mode", 0, "50,000 points per special/extra ball", &H00080000)                                                                                      'dip 20
    .AddFrame 205, 216, 190, "Game mode", &H00100000, Array("Replay", 0, "Extra ball", &H00100000)                                                                                                                             'dip 21
    .AddFrame 205, 262, 190, "3rd coin chute credits control", &H00001000, Array("No effect", 0, "Add 9", &H00001000)                                                                                                          'dip 13
    .AddChk 7, 292, 100, Array("Background sound32", &H80000000)                                                                                                                                                               'dip 32
    .AddChk 7, 308, 100, Array("Match feature", &H00020000)                                                                                                                                                                    'dip 18
    .AddChk 7, 328, 100, Array("Credits displayed", &H08000000)                                                                                                                                                                'dip 28
    .AddChk 7, 348, 100, Array("Coin switch tune", &H04000000)                                                                                                                                                                 'dip 27
    .AddChk 207, 308, 100, Array("Attract features", &H20000000)                                                                                                                                                               'dip 30
    .AddChkExtra 207, 328, 140, Array("Background sound", &H0010)                                                                                                                                                              'sounddip 5
    .AddChk 207, 348, 100, Array("Must be on", &H03000000)                                                                                                                                                                     'dip 25&26
    .AddLabel 50, 370, 300, 20, "After hitting OK, press F3 to reset game with new settings."
  End With
  Dim extra:extra = Controller.Dip(4) + Controller.Dip(5) * 256
  extra = vpmDips.ViewDipsExtra(extra)
  Controller.Dip(4) = extra And 255:Controller.Dip(5) = (extra And 65280) \ 256 And 255
End Sub

Sub SetDefaultDips
  'msgbox "Set Default Dips"
  Controller.Dip(0) = 0
  Controller.Dip(1) = 192
  Controller.Dip(2) = 195
  Controller.Dip(3) = 191
  Controller.Dip(4) = 20
  Controller.Dip(5) = 0
End Sub

Sub ResetDips
  InitDips = 1
  SaveValue cGameName, "SETDIPS", InitDips
End Sub

'Sub DipVals
' debug.print "Controller.Dip(0) = "&Controller.Dip(0)
' debug.print "Controller.Dip(1) = "&Controller.Dip(1)
' debug.print "Controller.Dip(2) = "&Controller.Dip(2)
' debug.print "Controller.Dip(3) = "&Controller.Dip(3)
' debug.print "Controller.Dip(4) = "&Controller.Dip(4)
' debug.print "Controller.Dip(5) = "&Controller.Dip(5)
'End Sub

'*******************************************
'  ZKEY: Key Press Handling
'*******************************************

Sub HH_KeyDown(ByVal keycode)
  VR_Keydown keycode
  If bInOptions Then
    Options_KeyDown keycode
    Exit Sub
  End If
  If keycode = LeftMagnaSave And InRect(HHBall.x,HHBall.y,1037,1689,1037,2137,500,2137,439,2045) Then
    If bOptionsMagna Then Options_Open() Else bOptionsMagna = True
  ElseIf keycode = RightMagnaSave And InRect(HHBall.x,HHBall.y,1139,1689,1139,2137,500,2137,439,2045) Then
    If bOptionsMagna Then Options_Open() Else bOptionsMagna = True
  End If

  If FlipperKeyMod = 1 then
    If keycode = LeftMagnaSave then
      SolLUFlipper 1
    End If
    If keycode = RightMagnaSave then
      SolRUFlipper 1
    End If
  End If
  If keycode = LeftTiltKey Then Nudge 90, 1: SoundNudgeLeft: End If
  If keycode = RightTiltKey Then Nudge 270, 1: SoundNudgeRight: End If
  If keycode = CenterTiltKey Then Nudge 0, 1: SoundNudgeCenter: End If
  If keycode = PlungerKey Then Plunger.Pullback: SoundPlungerPull: End If
  If keycode = StartGameKey Then SoundStartButton
  If keycode = AddCreditKey Or keycode = AddCreditKey2 Then SoundCoinIn

  If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub HH_KeyUp(ByVal keycode)
  VR_Keyup keycode
  If keycode = LeftMagnaSave And Not bInOptions Then bOptionsMagna = False
  If keycode = RightMagnaSave And Not bInOptions Then bOptionsMagna = False

  If FlipperKeyMod = 1 then
    If keycode = LeftMagnaSave then
      SolLUFlipper 0
    End If
    If keycode = RightMagnaSave then
      SolRUFlipper 0
    End If
  End If
  If keycode = PlungerKey Then PlungerRelease

  If vpmKeyUp(KeyCode) Then Exit Sub
End Sub

Sub PlungerRelease
  Plunger.Fire
  If InRect(HHBall.x,HHBall.y,1029,1844,1029,2026,1129,2026,1129,1844) Then
    SoundPlungerReleaseBall()
  Else
    SoundPlungerReleaseNoBall()
  End If
End Sub



'*******************************************
'  ZSOL: Solenoids
'*******************************************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"


SolCallback(1) = "Kick46"
SolCallback(2) = "UpKick"
SolCallback(5) = "UpperDropsUp"
SolCallback(6) = "BS"
SolCallback(8) = "PlayKnocker"
SolCallback(9) = "KickBallToLane"
SolCallback(10) = "GameOver"
SolCallback(11) = "SolGi"

Dim GameInPlay:GameInPlay=False

Sub GameOver(enabled)
  If enabled Then
    GameInPlay = True
  Else
    GameInPlay = False
  End If
End Sub


' Lights working as solenoids

Set LampCallback = GetRef("UpdateMultipleLamps")

Dim Old12, Old13, Old15, Old16,Old17
Old12=0:Old13=0:Old15=0:Old16=0:Old17=0

Sub UpdateMultipleLamps
  If Controller.Lamp(12) <> Old12 Then
    If Controller.Lamp(12) Then
      BasementUpKick
    End If
    Old12 = Controller.Lamp(12)
  End If

  If Controller.Lamp(13) <> Old13 Then
    If Controller.Lamp(13) Then
      LowerDropsUp
    End If
    Old13 = Controller.Lamp(13)
  End If

  If Controller.Lamp(15) <> Old15 Then
    If Controller.Lamp(15) Then
      Kick65
    End If
    Old15 = Controller.Lamp(15)
  End If

  If Controller.Lamp(16) <> Old16 Then
    If Controller.Lamp(16) Then
      TrapDoorR.Collidable = 0
      TrapDoorDir = 2
      TrapDoorTimer.Enabled = 1
      Playsoundat SoundFX("SOL_on",DOFContactors),BM_pTrapDoor
    Else
      TrapDoorR.Collidable = 1
      TrapDoorDir = -2
      TrapDoorTimer.Enabled = 1
      Playsoundat SoundFX("SOL_off",DOFContactors),BM_pTrapDoor
    End If
    Old16 = Controller.Lamp(16)
  End If

  If Controller.Lamp(17) <> Old17 Then
    If Controller.Lamp(17) Then
      SolGIOn
    Else
      SolGiOff
    End If
    Old17 = Controller.Lamp(17)
  End If
End Sub



'*******************************************
'  ZGII: GI
'*******************************************

dim GIstate: GIstate=0

Sub SolGi(Enabled)
  If Enabled Then
    Sound_GI_Relay 1,KnockerPosition
  Else
    Sound_GI_Relay 0,KnockerPosition
  End If
End Sub


Sub SolGIOff 'lower playfield off, upper pf on
  Dim xx, i
  GIstate = 0
  'Upper and Main
  For each xx in GIUpper: xx.State = 1: Next
  For each xx in GIMain: xx.State = 1: Next
  For i = 0 to 3
    Select Case i
      Case 0: If Not DT02.IsDropped Then DTShadows(0).visible = True
      Case 1: If Not DT12.IsDropped Then DTShadows(1).visible = True
      Case 2: If Not DT22.IsDropped Then DTShadows(2).visible = True
      Case 3: If Not DT32.IsDropped Then DTShadows(3).visible = True
    End Select
  Next
  'Lower
  For each xx in GILower: xx.State = 0: Next
End Sub


Sub SolGiOn 'lower playfield on, upper pf off
  Dim xx
  GIstate = 1
  'Upper and Main
  For each xx in GIUpper: xx.State = 0: Next
  For each xx in GIMain: xx.State = 0: Next
  For each xx in DTShadows: xx.visible = False: Next
  'Lower
  For each xx in GILower: xx.State = 1: Next
End Sub


''''''''' GI Color Change options
' reference: https://andi-siess.de/rgb-to-color-temperature/

Dim c2700k: c2700k = rgb(255, 169, 87)
Dim c3000k: c3000k = rgb(255, 180, 107)

Dim cRedFull: cRedFull = rgb(255,0,0)
Dim cRed: cRed= rgb(255,5,5)
Dim cPinkFull: cPinkFull = rgb(255,0,225)
Dim cPink: cPink = rgb(255,5,255)
Dim cWhiteFull: cWhiteFull = rgb(255,255,128)
Dim cWhite: cWhite = rgb(255,255,255)
Dim cBlueFull: cBlueFull= rgb(0,0,255)
Dim cBlue: cBlue = rgb(5,5,255)
Dim cCyanFull: cBlueFull= rgb(0,255,255)
Dim cCyan : cCyan = rgb(5,128,255)
Dim cYellowFull:cYellowFull  = rgb(255,255,128)
Dim cYellow: cYellow = rgb(255,255,0)
Dim cOrangeFull: cOrangeFull = rgb(255,128,0)
Dim cOrange: cOrange = rgb(255,70,5)
Dim cGreenFull: cGreenFull = rgb(0,255,0)
Dim cGreen: cGreen = rgb(5,255,5)
Dim cPurpleFull: cPurpleFull = rgb(128,0,255)
Dim cPurple: cPurple = rgb(100,5,128)
Dim cAmberFull: cAmberFull = rgb(255,197,143)
Dim cAmber: cAmber = rgb(255,197,143)

Dim cArray
cArray = Array(c2700k,cRed,cOrange,cGreen,cCyan,cPurple)

sub SetLowerGiColor(c)
  Dim xx, BL
  'Lower
  For each xx in GILower: xx.color = c: xx.colorfull = c: Next
  For each BL in BL_GI_LPF_giLL1: BL.color = c: Next
end Sub

sub SetMainGiColor(c)
  Dim xx, BL
  'Main
  For each xx in GIMain: xx.color = c: xx.colorfull = c: Next
  For each BL in BL_GI_MPF_giML1: BL.color = c: Next
  For each BL in BL_GI_MPF_giML2: BL.color = c: Next
  For each BL in BL_GI_MPF_giML3: BL.color = c: Next
  For each BL in BL_GI_MPF_giML4: BL.color = c: Next
  For each BL in BL_GI_MPF_giML5: BL.color = c: Next
  For each BL in BL_GI_MPF_giML6: BL.color = c: Next
end Sub

sub SetUpperGiColor(c)
  Dim xx, BL
  'Upper
  For each xx in GIUpper: xx.color = c: xx.colorfull = c: Next
  L21.color = c: L21.colorfull = c
  L21b.color = c: L21b.colorfull = c
  For each BL in BL_GI_UPF_gUL2: BL.color = c: Next
  For each BL in BL_GI_UPF_giUL1: BL.color = c: Next
  For each BL in BL_GI_UPF_giUL3: BL.color = c: Next
  For each BL in BL_GI_UPF_giUL4: BL.color = c: Next
  For each BL in BL_GI_UPF_giUL5: BL.color = c: Next
  For each BL in BL_GI_UPF_giUL6: BL.color = c: Next
  For each BL in BL_inserts_UPF_l21: BL.color = c: Next
end Sub

Sub SetUpperGiColorSpooky
  Dim xx, BL
  'Upper
  For each xx in GIUpper: xx.color = cPurple: xx.colorfull = cPurple: Next
  giUL6.color = cGreen: giUL6.colorfull = cGreen
  L21.color = cGreen: L21.colorfull = cGreen
  L21b.color = cGreen: L21b.colorfull = cGreen
  For each BL in BL_GI_UPF_gUL2: BL.color = cPurple: Next
  For each BL in BL_GI_UPF_giUL1: BL.color = cPurple: Next
  For each BL in BL_GI_UPF_giUL3: BL.color = cPurple: Next
  For each BL in BL_GI_UPF_giUL4: BL.color = cPurple: Next
  For each BL in BL_GI_UPF_giUL5: BL.color = cPurple: Next
  For each BL in BL_GI_UPF_giUL6: BL.color = cGreen: Next
  For each BL in BL_inserts_UPF_l21: BL.color = cGreen: Next
End Sub

Sub SetBumperColor(c)
  Dim BL
  'Bumpers
  BL1.color = c
  BL2.color = c
  BL3.color = c
  BL4.color = c
  For each BL in BL_BLight_BL1: BL.color = c: Next
  For each BL in BL_BLight_BL2: BL.color = c: Next
  For each BL in BL_BLight_BL3: BL.color = c: Next
  For each BL in BL_BLight_BL4: BL.color = c: Next
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

  If CCposts = 1 then
    giML5.Color = rgb(RedRGB,GreenRGB,BlueRGB)
    giML5.ColorFull = rgb(RedRGB,GreenRGB,BlueRGB)
    giML6.Color = rgb(RedRGB,GreenRGB,BlueRGB)
    giML6.ColorFull = rgb(RedRGB,GreenRGB,BlueRGB)

    For each xxCCGI in BL_GI_MPF_giML5: xxCCGI.color = rgb(RedRGB,GreenRGB,BlueRGB): Next
    For each xxCCGI in BL_GI_MPF_giML6: xxCCGI.color = rgb(RedRGB,GreenRGB,BlueRGB): Next
  End If
  MaterialColor "CCBallMaterial1", rgb(BlueRGB,GreenRGB,RedRGB)

End Sub

''''''''''''''''''' End Color Changing GI



'*******************************************
'  ZGLO: Glow Balls
'*******************************************


Sub MoveGlowBalls_Timer()
  pPinball1.z = HHBall.z:pPinball1.y = HHBall.y:pPinball1.x = HHBall.x
End Sub




'*******************************************
'  ZFLP: Flippers
'*******************************************

Const ReflipAngle = 20

Sub SolLFlipper(Enabled)
  If Enabled Then
    If FlipperKeyMod = 2 Then
      SolLUFlipper 1
    End if

    FlipperActivate LeftFlipper, LFPress
    FlipperActivate LeftFlipper2, LFPress2
    LF.Fire  'LeftFlipper.RotateToEnd
    LF2.Fire  'LeftFlipper2.RotateToEnd
    ' Sound effects
    If leftflipper.currentangle < leftflipper.endangle + ReflipAngle Then
      RandomSoundReflipUpLeft LeftFlipper
    Else
      SoundFlipperUpAttackLeft LeftFlipper
      RandomSoundFlipperUpLeft LeftFlipper
    End If
  Else
    If FlipperKeyMod = 2 Then
      SolLUFlipper 0
    End if

    FlipperDeActivate LeftFlipper, LFPress
    FlipperDeActivate LeftFlipper2, LFPress2
    LeftFlipper.RotateToStart
    LeftFlipper2.RotateToStart
    ' Sound effects
    If LeftFlipper.currentangle < LeftFlipper.startAngle - 5 Then
      RandomSoundFlipperDownLeft LeftFlipper
    End If
    FlipperLeftHitParm = FlipperUpSoundLevel
  End If
End Sub

Sub SolLUFlipper(enabled)
  If Enabled and GameInPlay Then
    If Controller.Lamp(17) Then
      FlipperActivate FlipperLL, LFPress4
      FlipperLL.RotateToEnd
    End If
    FlipperActivate FlipperUL, LFPress3
    ULF.Fire  'FlipperUL.RotateToEnd
    DOF 301, 1 ': debug.print "DOF 301, 1"

    If FlipperUL.currentangle < FlipperUL.endangle + ReflipAngle Then
      RandomSoundReflipUpLeft FlipperUL
    Else
      SoundFlipperUpAttackLeft FlipperUL
      RandomSoundFlipperUpLeft FlipperUL
    End If

  Else
    FlipperDeActivate FlipperLL, LFPress4
    FlipperLL.RotateToStart
    FlipperDeActivate FlipperUL, LFPress3
    FlipperUL.RotateToStart
    DOF 301, 0 ': debug.print "DOF 301, 0"

    If FlipperUL.currentangle < FlipperUL.startAngle - 5 Then
      RandomSoundFlipperDownLeft FlipperUL
    End If
  End If
End Sub

Sub SolRFlipper(Enabled)
  If Enabled Then
    If FlipperKeyMod = 2 Then
      SolRUFlipper 1
    End if

    FlipperActivate RightFlipper, RFPress
    FlipperActivate RightFlipper2, RFPress2
    RF.Fire  'RightFlipper.RotateToEnd
    RF2.Fire  'RightFlipper2.RotateToEnd

    ' Sound effects
    If rightflipper.currentangle > rightflipper.endangle - ReflipAngle Then
      RandomSoundReflipUpRight RightFlipper
    Else
      SoundFlipperUpAttackRight RightFlipper
      RandomSoundFlipperUpRight RightFlipper
    End If
  Else
    If FlipperKeyMod = 2 Then
      SolRUFlipper 0
    End if

    FlipperDeActivate RightFlipper, RFPress
    FlipperDeActivate RightFlipper2, RFPress2
    RightFlipper.RotateToStart
    RightFlipper2.RotateToStart

    ' Sound effects
    If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
      RandomSoundFlipperDownRight RightFlipper
    End If
    FlipperRightHitParm = FlipperUpSoundLevel
  End If
End Sub

Sub SolRUFlipper(enabled)
  If Enabled and GameInPlay Then
    If Controller.Lamp(17) Then
      FlipperActivate FlipperLR, RFPress4
      FlipperLR.RotateToEnd
    End If
    FlipperActivate FlipperUR, RFPress3
    URF.Fire
    DOF 302, 1 ': debug.print "DOF 302, 1"

    If FlipperUR.currentangle < FlipperUR.endangle + ReflipAngle Then
      RandomSoundReflipUpRight FlipperUR
    Else
      SoundFlipperUpAttackRight FlipperUR
      RandomSoundFlipperUpRight FlipperUR
    End If

  Else
    FlipperDeActivate FlipperLR, RFPress4
    FlipperLR.RotateToStart
    FlipperDeActivate FlipperUR, RFPress3
    FlipperUR.RotateToStart
    DOF 302, 0 ': debug.print "DOF 302, 0"

    If FlipperUR.currentangle < FlipperUR.startAngle - 5 Then
      RandomSoundFlipperDownRight FlipperUR
    End If
  End If
End Sub


Sub LeftFlipper_Collide(parm)
  CheckLiveCatch ActiveBall, LeftFlipper, LFCount, parm
  LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
  CheckLiveCatch ActiveBall, RightFlipper, RFCount, parm
  RightFlipperCollide parm
End Sub

Sub LeftFlipper2_Collide(parm)
  CheckLiveCatch ActiveBall, LeftFlipper2, LFCount2, parm
  LeftFlipperCollide parm
End Sub

Sub RightFlipper2_Collide(parm)
  CheckLiveCatch ActiveBall, RightFlipper2, RFCount2, parm
  RightFlipperCollide parm
End Sub

Sub FlipperUL_Collide(parm)
  CheckLiveCatch ActiveBall, FlipperUL, LFCount3, parm
  LeftFlipperCollide parm
End Sub

Sub FlipperUR_Collide(parm)
  CheckLiveCatch ActiveBall, FlipperUR, RFCount3, parm
  RightFlipperCollide parm
End Sub

Sub FlipperLL_Collide(parm)
  CheckLiveCatch ActiveBall, FlipperLL, LFCount4, parm
  LeftFlipperCollide parm
End Sub

Sub FlipperLR_Collide(parm)
  CheckLiveCatch ActiveBall, FlipperLR, RFCount4, parm
  RightFlipperCollide parm
End Sub






'*******************************************
' ZRBR: Rubbers
'*******************************************

'''' Scoring Rubbers

''upper
Sub sw42a_Hit
  vpmTimer.PulseSw(42)
  Dim BP
  For each BP in BP_UPFR1b: BP.visible = true: Next
  For each BP in BP_UPFR1a: BP.visible = false: Next
  sw42a.TimerEnabled = True
End Sub
Sub sw42a_Timer
  Dim BP
  For each BP in BP_UPFR1a: BP.visible = true: Next
  For each BP in BP_UPFR1b: BP.visible = false: Next
  sw42a.TimerEnabled = False
End Sub

Sub sw42b_Hit:vpmTimer.PulseSw(42):End Sub

''main
Sub sw66b_hit()
  vpmTimer.PulseSw(66)
  DOF 105, DOFPulse
End Sub

''Lower
Sub sw51a_Hit
  vpmTimer.PulseSw(51)
  Dim BP
  For each BP in BP_LPFR2b: BP.visible = true: Next
  For each BP in BP_LPFR2a: BP.visible = false: Next
  sw51a.TimerEnabled = True
End Sub
Sub sw51a_Timer
  Dim BP
  For each BP in BP_LPFR2a: BP.visible = true: Next
  For each BP in BP_LPFR2b: BP.visible = false: Next
  sw51a.TimerEnabled = False
End Sub


Sub sw51b_Hit:vpmTimer.PulseSw(51):End Sub





'Animated Rubbers UPF

Sub animUPFR3_hit
  Dim BP
  For each BP in BP_UPFR3b: BP.visible = true: Next
  For each BP in BP_UPFR3a: BP.visible = false: Next
  animUPFR3.TimerEnabled = True
End Sub
Sub animUPFR3_Timer
  Dim BP
  For each BP in BP_UPFR3a: BP.visible = true: Next
  For each BP in BP_UPFR3b: BP.visible = false: Next
  animUPFR3.TimerEnabled = False
End Sub

Sub animUPFR4_hit
  Dim BP
  For each BP in BP_UPFR4b: BP.visible = true: Next
  For each BP in BP_UPFR4a: BP.visible = false: Next
  animUPFR4.TimerEnabled = True
End Sub
Sub animUPFR4_Timer
  Dim BP
  For each BP in BP_UPFR4a: BP.visible = true: Next
  For each BP in BP_UPFR4b: BP.visible = false: Next
  animUPFR4.TimerEnabled = False
End Sub

Sub animUPFR2b_hit
  Dim BP
  For each BP in BP_UPFR2b: BP.visible = true: Next
  For each BP in BP_UPFR2a: BP.visible = false: Next
  animUPFR2b.TimerEnabled = True
End Sub
Sub animUPFR2b_Timer
  Dim BP
  For each BP in BP_UPFR2a: BP.visible = true: Next
  For each BP in BP_UPFR2b: BP.visible = false: Next
  animUPFR2b.TimerEnabled = False
End Sub

Sub animUPFR2c_hit
  Dim BP
  For each BP in BP_UPFR2c: BP.visible = true: Next
  For each BP in BP_UPFR2a: BP.visible = false: Next
  animUPFR2c.TimerEnabled = True
End Sub
Sub animUPFR2c_Timer
  Dim BP
  For each BP in BP_UPFR2a: BP.visible = true: Next
  For each BP in BP_UPFR2c: BP.visible = false: Next
  animUPFR2c.TimerEnabled = False
End Sub

Sub animUPFR2d_hit
  Dim BP
  For each BP in BP_UPFR2d: BP.visible = true: Next
  For each BP in BP_UPFR2a: BP.visible = false: Next
  animUPFR2d.TimerEnabled = True
End Sub
Sub animUPFR2d_Timer
  Dim BP
  For each BP in BP_UPFR2a: BP.visible = true: Next
  For each BP in BP_UPFR2d: BP.visible = false: Next
  animUPFR2d.TimerEnabled = False
End Sub




'*******************************************
'  ZSWI: Switches
'*******************************************

Sub Sw04_Hit(): Controller.Switch(04)=1: End Sub
Sub Sw04_UnHit(): Controller.Switch(04)=0: End Sub

Sub Sw21_Hit(): Controller.Switch(21)=1: End Sub
Sub Sw21_UnHit(): Controller.Switch(21)=0: End Sub

Sub Sw34_Hit(): Controller.Switch(34)=1: End Sub
Sub Sw34_UnHit():Controller.Switch(34)=0: End Sub

Sub Sw54_Hit():Controller.Switch(54)=1: End Sub
Sub Sw54_UnHit(): Controller.Switch(54)=0: End Sub

Sub Sw64_Hit(): Controller.Switch(64)=1: Gate3.Damping=0.85: sw64step=0: me.timerenabled = True: End Sub
Sub Sw64_UnHit(): Controller.Switch(64)=0: End Sub

Sub Sw35_Hit(): Controller.Switch(35)=1: Gate4.Damping=0.85: sw35step=0: me.timerenabled = True: End Sub
Sub Sw35_UnHit(): Controller.Switch(35)=0: End Sub

'''''''switches for lower throughs

Sub sw06_Hit():Controller.Switch(6) = 1:End Sub
Sub sw06_UnHit():Controller.Switch(6) = 0:End Sub

Sub sw16_Hit():Controller.Switch(16) = 1:End Sub
Sub sw16_UnHit():Controller.Switch(16) = 0:End Sub

Sub sw26_Hit():Controller.Switch(26) = 1:End Sub
Sub sw26_UnHit():Controller.Switch(26) = 0:End Sub

Sub sw36_Hit():Controller.Switch(36) = 1:End Sub
Sub sw36_UnHit():Controller.Switch(36) = 0:End Sub



'*******************************************
'  ZGAT: Gates
'*******************************************


Dim sw64step, sw35step

Sub Sw64_timer()
  Select case sw64step
    Case 0:'Gate3.Damping = .9:'Gate3.GravityFactor = 5
    Case 1:Gate3.Damping = .9:Gate3.GravityFactor = 3
    Case 2:Gate3.Damping = .95
    Case 3:Gate3.Damping = .99999
    Case 4:
    Case 35: me.timerenabled = false:sw64step = 0:Gate3.Damping = .85:Gate3.GravityFactor = 5
  End Select
  sw64step = sw64step + 1
End Sub

Sub Sw35_timer()
  Select case sw35step
    Case 0:'Gate4.Damping = .9:'Gate4.GravityFactor = 5
    Case 1:Gate4.Damping = .9:Gate4.GravityFactor = 3
    Case 2:Gate4.Damping = .95
    Case 3:Gate4.Damping = .99999
    Case 4:
    Case 35: me.timerenabled = false:sw35step = 0:Gate4.Damping = .85:Gate4.GravityFactor = 5
  End Select
  sw35step = sw35step + 1
End Sub



'*******************************************
'  ZKIC: Kickers
'*******************************************

''''  Up Kick


Sub sw45_hit():Controller.Switch(45) = 1: SoundSaucerLock: End Sub
Sub sw45_unhit():Controller.Switch(45) = 0:End Sub

Sub UpKick(Enabled)
  If Enabled Then
    sw45.Kick 0,35,1.56
    SoundSaucerKick 1,sw45
  End If
End Sub


''''  Kick 46

Dim SW46Step, sw46ark

Sub sw46_hit():Controller.Switch(46) = 1: SoundSaucerLock: End Sub
Sub sw46_unhit():Controller.Switch(46) = 0:End Sub

Sub Kick46(Enabled)
  If Enabled Then
    If sw46ark = 1 Then
    Else
      sw46ark = 1
      SW46Step = 0
      sw46.timerenabled = true
      SoundSaucerKick 1,sw46
    End If
  End If
End Sub

Sub SW46_Timer()
  Select Case SW46Step
    Case 0: TWKicker1.transY = 2
    Case 1: TWKicker1.transY = 5:HHBall.velz = 12:HHBall.vely = 7:'HHBall.velx = 2:'HHBall.vely = 10
    Case 2: TWKicker1.transY = 10:HHBall.velx = 3
    Case 3: TWKicker1.transY = 2
    Case 4: TWKicker1.transY = 0:
    Case 5:
    Case 9: Controller.Switch(46) = 0
    Case 10:Me.TimerEnabled = 0:SW46Step = 0:sw46ark = 0
  End Select
  SW46Step = SW46Step + 1
  Dim BP: For each BP in BP_TWKicker1: BP.transz = TWKicker1.transY: Next
End Sub


''''  Kick 65

Sub sw65_hit():Controller.Switch(65) = 1:End Sub
Sub sw65_unhit():Controller.Switch(65) = 0:End Sub

Sub Kick65()
  PlungerIM.AutoFire
End Sub



'''' Basement Up Kick


Sub sw31_hit(): Controller.Switch(31) = 1: SoundSaucerLock: End Sub
Sub sw31_unhit(): Controller.Switch(31) = 0:End Sub

Sub BasementUpKick()
  sw31.timerenabled = 1
End Sub

Dim sw31step, sw31ka, sw31ks

Sub sw31_timer()
    Select Case sw31step
        Case 0: SoundSaucerKick 1,sw31
        Case 1:
                sw31ka = Rnd() * 30 - 12.5
                sw31ks = 65 + sw31ka/2 + Rnd() * 10
                sw31.Kickxyz sw31ka, sw31ks,1.50,-7,-1,20
                sw31.timerEnabled = 0
                sw31step = -1
    End Select
    sw31step = sw31step + 1
End Sub

' Basement Special

Sub sw41_hit():Controller.Switch(41) = 1: SoundSaucerLock: End Sub
Sub sw41_unhit():Controller.Switch(41) = 0:End Sub

Sub BS(enabled)
    SoundSaucerKick 1,sw41
    sw41.Kickz 90,15,0.7,25
End Sub


'*******************************************
'  ZSTA: Stand-up Targets
'*******************************************


'Upper
Sub sw03_Hit: StHit 3 : End Sub
Sub sw13_Hit: StHit 13: End Sub
Sub sw23_Hit: StHit 23: End Sub
Sub sw33_Hit: StHit 33: End Sub

'Main
Sub sw05_Hit: STHit 5 : End Sub
Sub sw15_Hit: STHit 15: End Sub
Sub sw55_Hit: STHit 55: End Sub

'Lower
Sub sw50_Hit: STHit 50: End Sub
Sub sw60_Hit: STHit 60: End Sub

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
Dim ST03, ST13, ST23, ST33  'Upper
Dim ST05, ST15, ST55    'Main
Dim ST50, ST60        'Lower

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


Set ST03 = (new StandupTarget)(sw03, BM_pSW03, 3, 0)
Set ST13 = (new StandupTarget)(sw13, BM_pSW13, 13, 0)
Set ST23 = (new StandupTarget)(sw23, BM_pSW23, 23, 0)
Set ST33 = (new StandupTarget)(sw33, BM_pSW33, 33, 0)

Set ST05 = (new StandupTarget)(sw05, BM_pSW05, 5, 0)
Set ST15 = (new StandupTarget)(sw15, BM_pSW15, 15, 0)
Set ST55 = (new StandupTarget)(sw55, BM_pSW55, 55, 0)

Set ST50 = (new StandupTarget)(sw50, BM_pSW50, 50, 0)
Set ST60 = (new StandupTarget)(sw60, BM_pSW60, 60, 0)

'Add all the Stand-up Target Arrays to Stand-up Target Animation Array
'   STAnimationArray = Array(ST1, ST2, ....)
Dim STArray
STArray = Array(ST03, ST13, ST23, ST33, ST05, ST15, ST55, ST50, ST60)

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

'*******************************************
'  ZKTA: KICKING TARGETS
'*******************************************

'Kicking Standups
Sub sw11_Hit: KTHit 11: End Sub
Sub sw14_Hit: KTHit 14: End Sub
Sub sw24_Hit: KTHit 24: End Sub
Sub sw25_Hit: KTHit 25: End Sub

'Kick Objects
Sub pSW11col_hit: KTKick 11: End Sub
Sub pSW14col_hit: KTKick 14: End Sub
Sub pSW24col_hit: KTKick 24: End Sub
Sub pSW25col_hit: KTKick 25: End Sub

Class KickingTarget
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

'Define a variable for each kicking target
Dim KT11, KT14, KT24, KT25  'Kicking

'Set array with kicking target objects
'
'KickingTargetvar = Array(primary, prim, swtich)
'   primary:  vp target to determine target hit
'   prim:    primitive target used for visuals and animation
'          IMPORTANT!!!
'          transy must be used to offset the target animation
'   switch:  ROM switch number
'   animate:  Arrary slot for handling the animation instrucitons, set to 0
'


Set KT11 = (new KickingTarget)(sw11, BM_pSW11, 11, 0)
Set KT14 = (new KickingTarget)(sw14, BM_pSW14, 14, 0)
Set KT24 = (new KickingTarget)(sw24, BM_pSW24, 24, 0)
Set KT25 = (new KickingTarget)(sw25, BM_pSW25, 25, 0)

'Add all the Kicking Target Arrays to Kicking Target Animation Array
'   KTAnimationArray = Array(KT1, KT2, ....)
Dim KTArray
KTArray = Array(KT11, KT14, KT24, KT25)

' kSpring   - strength of the target spring (non-kicking)
' kKickVel  - velocity imparted on the ball with the target solenoid fires
' kDist   - distance the target must be displaced before the switch registers and the solenoid fires
' kWidth  - total width of the target
Dim kSpring, kKickVel, kDist, kWidth
kSpring = 15
kKickvel = 30
kDist = 10
kWidth = 66

'''''' KICKING TARGETS FUNCTIONS

Sub KTHit(switch)
  Dim i
  i = KTArrayID(switch)

  PlayTargetSound
  KTArray(i).animate = STCheckHit(ActiveBall,KTArray(i).primary)

  Activeball.velx = cor.BallVelx(Activeball.id)
  Activeball.vely = cor.BallVely(Activeball.id)

  DoKTAnim
End Sub

Sub KTKick(switch)
  Dim i
  i = KTArrayID(switch)

  KTArray(i).animate = 2

  Activeball.velx = cor.BallVelx(Activeball.id)
  Activeball.vely = cor.BallVely(Activeball.id)

  DoKTAnim
End Sub


Function KTArrayID(switch)
  Dim i
  For i = 0 To UBound(KTArray)
    If KTArray(i).sw = switch Then
      KTArrayID = i
      Exit Function
    End If
  Next
End Function

Sub KTAnim_Timer
  DoKTAnim
  dim bangle
End Sub

Sub DoKTAnim()
  Dim i
  For i = 0 To UBound(KTArray)
    KTArray(i).animate = KTAnimate(KTArray(i).primary,KTArray(i).prim,KTArray(i).sw,KTArray(i).animate)
  Next
End Sub

Function KTAnimate(primary, prim, switch,  animate)
  KTAnimate = animate

  If animate = 0  Then
    primary.uservalue = 0
    KTAnimate = 0
    primary.collidable = 1
    Exit Function
  ElseIf primary.uservalue = 0 Then
    primary.uservalue = GameTime
  End If

  Dim animtime, btdist, btwidth, btangle, bangle, angle, nposx, nposy, tdist, kparavel, kballvel
  tdist = 31.31

  angle = primary.orientation

  animtime = GameTime - primary.uservalue
  primary.uservalue = GameTime

  nposx = primary.x + prim.transy * dCos(angle + 90)
  nposy = primary.y + prim.transy * dSin(angle + 90)

  btwidth = DistancePL(HHBall.x,HHBall.y,primary.x,primary.y,primary.x+dcos(angle+90),primary.y+dsin(angle+90))
  btdist = DistancePL(HHBall.x,HHBall.y,nposx,nposy,nposx+dcos(angle),nposy+dsin(angle))

  kballvel = sqr(HHBall.velx^2 + HHBall.vely^2)

  if kballvel <> 0 Then
    bangle = arcCos(HHball.velx/kballvel)*180/PI
  Else
    bangle = 0
  End If

  If HHBall.vely < 0 Then
    bangle = bangle * -1
  End If

  btangle = bangle - angle

  prim.transy = prim.transy + btdist - tdist

  'debug.print btangle & " " & btwidth & " " & btdist &  " " & prim.transy

  If animate = 1 Then
    primary.collidable = 0

    If btdist < tdist and btwidth < kWidth/2 + 25 Then
      if abs(prim.transy) >= kDist Then
        vpmTimer.PulseSw switch
        'fire Solenoid
        RandomSoundSlingshotLeft primary
        kparavel = Sqr(HHBall.velx^2 + HHBall.vely^2)*dCos(btangle)
        HHBall.velx = dcos(angle)*kparavel - (kKickVel * dsin(angle))
        HHBall.vely = dsin(angle)*kparavel + (kKickVel * dcos(angle))
        KTAnimate = 3
        'debug.print HHBall.velx & " fire " & HHBall.vely & " kpara " & kparavel
      Else
        HHBall.velx = HHBall.velx - (kSpring * dsin(angle) * abs(prim.transy) * animtime/1000)
        HHBall.vely = HHBall.vely + (kSpring * dcos(angle) * abs(prim.transy) * animtime/1000)
        'debug.print HHBall.velx & " nofire " & HHBall.vely
        'debug.print - (kSpring * dsin(angle) * abs(prim.transy) * animtime/1000) & " delta " & + (kSpring * dcos(angle) * abs(prim.transy) * animtime/1000)
      End If
    Elseif btdist > tdist then
      HHBall.velx = HHBall.velx - (kSpring * dsin(angle) * abs(prim.transy) * animtime/1000)
      HHBall.vely = HHBall.vely + (kSpring * dcos(angle) * abs(prim.transy) * animtime/1000)
      'debug.print HHBall.velx & " nofire " & HHBall.vely
      'debug.print - (kSpring * dsin(angle) * abs(prim.transy) * animtime/1000) & " delta " & + (kSpring * dcos(angle) * abs(prim.transy) * animtime/1000)

      if prim.transy >= 0 Then
        prim.transy = 0
        KTAnimate = 0
        Exit Function
      end If
    Else
      prim.transy = prim.transy + 1
      if prim.transy >= 0 Then
        prim.transy = 0
        KTAnimate = 0
        Exit Function
      end If
    End If
  Elseif animate = 2 Then
    vpmTimer.PulseSw switch
    'fire Solenoid
    RandomSoundSlingshotLeft primary
    kparavel = Sqr(HHBall.velx^2 + HHBall.vely^2)*dCos(btangle)
    HHBall.velx = dcos(angle)*kparavel - (kKickVel * dsin(angle))
    HHBall.vely = dsin(angle)*kparavel + (kKickVel * dcos(angle))
    KTAnimate = 3
    'debug.print HHBall.velx & " fire2 " & HHBall.vely & " kpara " & kparavel
  Elseif animate = 3 Then
    'debug.print HHBall.velx & " fire3 " & HHBall.vely & " tang " & dCos(btangle) * kballvel

    if prim.transy >= 0 Then
      prim.transy = 0
      KTAnimate = 0
      Exit Function
    end If
  End If
End Function

'******************************************************
'*   END KICKING TARGETS
'******************************************************


'*******************************************
'  ZDTA: Drop Targets
'*******************************************

'Solenoids

Sub UpperDropsUp(enabled)
  If enabled then
    DTRaise 2
    DTRaise 12
    DTRaise 22
    DTRaise 32
    RandomSoundDropTargetReset BM_pSW22
    Dim xx: For each xx in DTShadows: xx.visible = True: Next
  End If
End Sub

Sub LowerDropsUp()
  DTRaise 0
  DTRaise 10
  DTRaise 20
  DTRaise 30
  DTRaise 40
  RandomSoundDropTargetReset BM_pSW20
End Sub


'Upper

Sub sw02_Hit: DTHit 2 : End Sub
Sub sw12_Hit: DTHit 12: End Sub
Sub sw22_Hit: DTHit 22: End Sub
Sub sw32_Hit: DTHit 32: End Sub

'Lower

Sub sw00_Hit: DTHit 0 : End Sub
Sub sw10_Hit: DTHit 10: End Sub
Sub sw20_Hit: DTHit 20: End Sub
Sub sw30_Hit: DTHit 30: End Sub
Sub sw40_Hit: DTHit 40: End Sub


'Actions

Sub DTAction(switchid)
  Select Case switchid
    Case 2: DTShadows(0).visible = False
    Case 12: DTShadows(1).visible = False
    Case 22: DTShadows(2).visible = False
    Case 32: DTShadows(3).visible = False
  End Select
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
Dim DT02, DT12, DT22, DT32      'Upper
Dim DT00, DT10, DT20, DT30, DT40    'Lower

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

Set DT02 = (new DropTarget)(sw02, sw02a, BM_psw02, 2, 0, False)
Set DT12 = (new DropTarget)(sw12, sw12a, BM_psw12, 12, 0, False)
Set DT22 = (new DropTarget)(sw22, sw22a, BM_psw22, 22, 0, False)
Set DT32 = (new DropTarget)(sw32, sw32a, BM_psw32, 32, 0, False)

Set DT00 = (new DropTarget)(sw00, sw00a, BM_psw00, 0, 0, False)
Set DT10 = (new DropTarget)(sw10, sw10a, BM_psw10, 10, 0, False)
Set DT20 = (new DropTarget)(sw20, sw20a, BM_psw20, 20, 0, False)
Set DT30 = (new DropTarget)(sw30, sw30a, BM_psw30, 30, 0, False)
Set DT40 = (new DropTarget)(sw40, sw40a, BM_psw40, 40, 0, False)


Dim DTArray
DTArray = Array(DT02, DT12, DT22, DT32, DT00, DT10, DT20, DT30, DT40)

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




'*******************************************
'  ZSLG: Slingshots
'*******************************************


Dim LowerSlingshotStep, UpperSlingShotStep

''Upper
Sub UpperSlingShot_Slingshot
  vpmTimer.PulseSw 66
  UpperSlingshotStep = 0
  UpperSlingShot.TimerEnabled = 1
  DOF 104, DOFPulse
  RandomSoundSlingshotRight UpperSlingPos
End Sub

Sub UpperSlingshot_Timer
  Dim va, vb, x,  BP
  Select Case UpperSlingshotStep
    Case 0:va = false: vb = true: x = 15
    Case 3:va = true: vb = false: x = 0:UpperSlingshot.TimerEnabled = 0
  End Select

  For each BP in BP_USlingPos: BP.transx = x: Next
  For each BP in BP_MPFR1b: BP.visible = vb: Next
  'For each BP in BP_MPFR1a: BP.visible = va: Next

  UpperSlingshotStep = UpperSlingshotStep + 1
End Sub

''Lower
Sub LowerSlingShot_Slingshot
  vpmTimer.PulseSw 51
  LowerSlingshotStep = 0
  LowerSlingshot.TimerEnabled = 1
  DOF 102, DOFPulse
  RandomSoundSlingshotLeft LowerSlingPos
End Sub

Sub LowerSlingshot_Timer
  Dim va, vb, y,  BP
  Select Case LowerSlingshotStep
    Case 0:va = false: vb = true: x = 20
    Case 3:va = true: vb = false: x = 0:LowerSlingshot.TimerEnabled = 0
  End Select

  For each BP in BP_LSlingPos: BP.transy = x: Next
  For each BP in BP_LPFR1b: BP.visible = vb: Next
  'For each BP in BP_LPFR1a: BP.visible = va: Next

  LowerSlingshotStep = LowerSlingshotStep + 1
End Sub




'*******************************************
'  ZBMP: Bumpers
'*******************************************


Sub Bumper2_hit() 'Main Top
  vpmTimer.PulseSw(44)
  RandomSoundBumperMiddle Bumper2
End Sub

Sub Bumper1_hit() 'Main Bottom
  vpmTimer.PulseSw(44)
  RandomSoundBumperBottom Bumper1
End Sub

Sub Bumper3_hit() 'Upper
  vpmTimer.PulseSw(43)
  RandomSoundBumperTop Bumper3
End Sub

Sub Bumper4_hit() 'Lower
  vpmTimer.PulseSw(1)
  RandomSoundBumperMiddle Bumper4
End Sub



'*******************************************
'  ZKNO: Knocker
'*******************************************

Sub PlayKnocker(Enabled)
  DelayKnocker.enabled = 1
End Sub

Dim DelayKnockerStep

Sub DelayKnocker_timer()
  Select Case DelayKnockerStep
    Case 0:
    Case 1: KnockerSolenoid
    Case 2:
    Case 3:
    Case 4:
    Case 5:
    Case 6:
    Case 7:
    Case 8:
    Case 9:
    Case 20: DelayKnocker.Enabled = 0: DelayKnockerStep = 0
  End Select
  DelayKnockerStep = DelayKnockerStep + 1
End Sub



'*******************************************
'  ZTRP: Trap Door
'*******************************************


Dim TrapDoorDir, TrapDoorAngle
Const TrapDoorMin = 0
Const TrapDoorMax = 30

Sub TrapDoorTimer_timer ()
  TrapDoorAngle = -BM_pTrapDoor.RotX
  TrapDoorAngle = TrapDoorAngle + TrapDoorDir
  If TrapDoorAngle >= TrapDoorMax then
    TrapDoorTimer.Enabled = 0
    TrapDoorAngle = TrapDoorMax
  End If
  If TrapDoorAngle <= TrapDoorMin then
    TrapDoorTimer.Enabled = 0
    TrapDoorAngle = TrapDoorMin
  End If
  Dim BP: For each BP in BP_pTrapDoor: BP. RotX = -TrapDoorAngle: Next
End Sub


'*******************************************
'  ZDRN: Drain, Trough, and Ball Release
'*******************************************


'TROUGH
Sub Kicker1_hit: controller.switch(67) = 1: End Sub
Sub Kicker1_unhit: controller.switch(67) = 0: End Sub


' DRAIN & RELEASE
Sub Drain_Hit
  RandomSoundDrain Drain
  Drain.Kick 80,20
End Sub

Sub KickBallToLane(enabled)
  If enabled Then
    Kicker1.Kick 90,12
    RandomSoundBallRelease Kicker1
  End If
End Sub




'*****************************************
'  ZBRL: Ball Rolling and Drop Sounds
'*****************************************

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
    If BallVel(gBOT(b)) > 1 And Not InRect(gBOT(b).x,gBOT(b).y,459,2045,1029,1751,1029,1873,495,2141) Then
      rolling(b) = True
      PlaySound ("BallRoll_" & b), - 1, VolPlayfieldRoll(gBOT(b)) * BallRollVolume * MechVolume, AudioPan(gBOT(b)), 0, PitchPlayfieldRoll(gBOT(b)), 1, 0, AudioFade(gBOT(b))
    Else
      If rolling(b) = True Then
        StopSound("BallRoll_" & b)
        rolling(b) = False
      End If
    End If

    ' Ball Drop Sounds
    If gBOT(b).VelZ <  - 1 And (gBOT(b).z < 55 And gBOT(b).z > 27) or (gBOT(b).z < (0.1971 * gBOT(b).Y - 525.8 + 55) and gBOT(b).z > (0.1971 * gBOT(b).Y - 525.8 + 27)) Then 'height adjust for ball drop sounds
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

    ' ********* Secret Target Code
    If InRect(gBOT(b).x,gBOT(b).y,483,605,543,605,543,732,483,732) And gBOT(b).z > -10 Then
      pSecretTarget.ObjrotX = (732 - gBOT(b).y)*90/35
      if pSecretTarget.ObjRotX > 90 Then pSecretTarget.ObjRotX = 90

      If SecretTargetHit = 0 Then
        SecretTargetTimer.enabled = 1
        RandomSoundTargetHitSecret gBOT(b)
      End If
      SecretTargetHit = 1
    Elseif SecretTargetHit = 1 and gBOT(b).z <= -10 Then
      SecretTargetStep = 8
      SecretTargetHit = 0
    Elseif SecretTargetHit = 1 and gBOT(b).z > 20 Then
      SecretTargetHit = 0
      SecretTargetStep = 17
    End If
    ' ********* End Target Code
  Next
End Sub

'******************************************************
'****  END BALL ROLLING AND DROP SOUNDS
'******************************************************

'*******************************************
'  ZSEC: Secret Target
'*******************************************


Dim SecretTargetHit,SecretTargetStep
SecretTargetHit = 0
SecretTargetStep = 0

Sub SecretTargetTimer_timer()
  Select Case SecretTargetStep
    Case 7:SecretTargetStep = 6
    Case 8:pSecretTarget.ObjRotX = 90:Playsoundat ("SecretPassageSound"),pSecretTarget'PlaySoundAtBall "SecretPassageSound"
    Case 9:pSecretTarget.ObjRotX = 75
    Case 10:pSecretTarget.ObjRotX = 60
    Case 11:pSecretTarget.ObjRotX = 45
    Case 12:pSecretTarget.ObjRotX = 30
    Case 13:pSecretTarget.ObjRotX = 15
    Case 14:pSecretTarget.ObjRotX = 0
    Case 15:pSecretTarget.ObjRotX = -8
    Case 16:pSecretTarget.ObjRotX = 8
    Case 17:pSecretTarget.ObjRotX = 0:Me.Enabled = 0:SecretTargetStep = 0
  End Select
  dim BP: For each BP in BP_pSWSecret: BP.ObjRotX = pSecretTarget.ObjRotX :Next
  SecretTargetStep = SecretTargetStep + 1
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
    objBallShadow(iii).Z = 1 + iii / 1000 + 1.04
    objBallShadow(iii).visible = 0
  Next
End Sub


Sub BSUpdate
  Dim s

' 'Hide shadow of deleted balls
' For s = UBound(gBOT) + 1 To tnob - 1
'   objBallShadow(s).visible = 0
' Next
'
' If UBound(gBOT) < lob Then Exit Sub 'No balls in play, exit

  'The Magic happens now
  For s = lob To UBound(gBOT)
    ' *** Normal "ambient light" ball shadow

    'Primitive shadow on playfield, flasher shadow in ramps
    '** If on main and upper pf
    If gBOT(s).Z > 20  and not (InRect(gBOT(s).x,gBOT(s).y,330,0,1114,0,1114,600,333,600) and gBOT(s).Z < 30) And BallMod<3 Then
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


'****************************************************************
'****  END VPW DYNAMIC BALL SHADOWS by Iakki, Apophis, and Wylte
'****************************************************************




'*****************************************
'  ZRRL: Ramp Rolling Sound effects
'*****************************************


'RampHelpers

' Main

Sub Ramp3HelperA_hit
  If activeball.vely < 0 Then PlaySoundAtBallVol "RampLoop1",RampRollVolume
End Sub

Sub Ramp3HelperA_unhit
  If activeball.vely > 0 Then StopSound("RampLoop1")
End Sub

Sub Ramp3HelperB_hit
  If activeball.vely > 0 Then PlaySoundAtBallVol "RampLoop1",RampRollVolume
End Sub

Sub Ramp3HelperB_unhit
  If activeball.vely < 0 Then StopSound("RampLoop1")
End Sub


' Lower

Sub Ramp1HelperA_hit
  PlaySoundAtBallVol "RampLoop1",RampRollVolume
End Sub

Sub Ramp1HelperB_hit
  StopSound("RampLoop1")
End Sub


Sub Ramp4HelperA_hit
  PlaySoundAtBallVol "RampLoop1",RampRollVolume
End Sub

Sub Ramp4HelperB_hit
  StopSound("RampLoop1")
End Sub


Sub Ramp5HelperA_hit
  PlaySoundAtBallVol "RampLoop1",RampRollVolume
End Sub

Sub Ramp5HelperB_hit
  StopSound("RampLoop1")
End Sub

Sub Ramp6HelperA_hit
  PlaySoundAtBallVol "RampLoop1",RampRollVolume
End Sub

Sub Ramp6HelperB_hit
  StopSound("RampLoop1")
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





'*******************************************
'  ZLED: LEDs
'*******************************************


'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'     LL    EEEEEE  DDDD    ,,   SSSSS
'         LL    EE    DD  DD    ,,  SS
'     LL    EE    DD   DD    ,   SS
'     LL    EEEE  DD   DD        SS
'     LL    EE    DD  DD          SS
'     LLLLLL  EEEEEE  DDDD      SSSSS
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$



'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'   6 Didget Array
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Dim LED(33)
LED(0)=Array(d121,d122,d123,d124,d125,d126,d127,LXM,d128)
LED(1)=Array(d131,d132,d133,d134,d135,d136,d137,LXM,d138)
LED(2)=Array(d141,d142,d143,d144,d145,d146,d147,LXM,d148)
LED(3)=Array(d151,d152,d153,d154,d155,d156,d157,LXM,d158)
LED(4)=Array(d161,d162,d163,d164,d165,d166,d167,LXM,d168)
LED(5)=Array(d171,d172,d173,d174,d175,d176,d177,LXM,d178)

LED(6)=Array(d221,d222,d223,d224,d225,d226,d227,LXM,d228)
LED(7)=Array(d231,d232,d233,d234,d235,d236,d237,LXM,d238)
LED(8)=Array(d241,d242,d243,d244,d245,d246,d247,LXM,d248)
LED(9)=Array(d251,d252,d253,d254,d255,d256,d257,LXM,d258)
LED(10)=Array(d261,d262,d263,d264,d265,d266,d267,LXM,d268)
LED(11)=Array(d271,d272,d273,d274,d275,d276,d277,LXM,d278)

LED(12)=Array(d321a,d322a,d323a,d324a,d325a,d326a,d327a,LXM,d328a)
LED(13)=Array(d331a,d332a,d333a,d334a,d335a,d336a,d337a,LXM,d338a)
LED(14)=Array(d341,d342,d343,d344,d345,d346,d347,LXM,d348)
LED(15)=Array(d351,d352,d353,d354,d355,d356,d357,LXM,d358)
LED(16)=Array(d361,d362,d363,d364,d365,d366,d367,LXM,d368)
LED(17)=Array(d371,d372,d373,d374,d375,d376,d377,LXM,d378)

LED(18)=Array(d421,d422,d423,d424,d425,d426,d427,LXM,d428)
LED(19)=Array(d431,d432,d433,d434,d435,d436,d437,LXM,d438)
LED(20)=Array(d441,d442,d443,d444,d445,d446,d447,LXM,d448)
LED(21)=Array(d451,d452,d453,d454,d455,d456,d457,LXM,d458)
LED(22)=Array(d461,d462,d463,d464,d465,d466,d467,LXM,d468)
LED(23)=Array(d471,d472,d473,d474,d475,d476,d477,LXM,d478)


LED(24)=Array(d511,d512,d513,d514,d515,d516,d517,LXM,d518)
LED(25)=Array(d521,d522,d523,d524,d525,d526,d527,LXM,d528)
LED(26)=Array(d611,d612,d613,d614,d615,d616,d617,LXM,d618)
LED(27)=Array(d621,d622,d623,d624,d625,d626,d627,LXM,d628)


LED(28)=Array(D281,D282,D283,D284,D285,D286,D287,LXM,D288)
LED(29)=Array(D291,D292,D293,D294,D295,D296,D297,LXM,D298)
LED(30)=Array(D301,D302,D303,D304,D305,D306,D307,LXM,D308)
LED(31)=Array(D311,D312,D313,D314,D315,D316,D317,LXM,D318)
LED(32)=Array(D321,D322,D323,D324,D325,D326,D327,LXM,D328)
LED(33)=Array(D331,D332,D333,D334,D335,D336,D337,LXM,D338)



Sub DisplayTimer6_Timer
  Dim ChgLED, ii, num, chg, stat, obj
  ChgLED = Controller.ChangedLEDs (&Hffffffff, &Hffffffff)
  If Not IsEmpty (ChgLED) Then
    For ii = 0 To UBound (chgLED)
      num = chgLED (ii, 0) : chg = chgLED (ii, 1) : stat = chgLED (ii, 2)
      For Each obj In LED (num)
        If chg And 1 Then obj.State = stat And 1
        chg = chg \ 2 : stat = stat \ 2
      Next
    Next
  End If
End Sub


'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'   7 Didget Array
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Dim LED7(38)
LED7(0)=Array(d111,d112,d113,d114,d115,d116,d117,LXM,d118)
LED7(1)=Array(d121,d122,d123,d124,d125,d126,d127,LXM,d128)
LED7(2)=Array(d131,d132,d133,d134,d135,d136,d137,LXM,d138)
LED7(3)=Array(d141,d142,d143,d144,d145,d146,d147,LXM,d148)
LED7(4)=Array(d151,d152,d153,d154,d155,d156,d157,LXM,d158)
LED7(5)=Array(d161,d162,d163,d164,d165,d166,d167,LXM,d168)
LED7(6)=Array(d171,d172,d173,d174,d175,d176,d177,LXM,d178)

LED7(7)=Array(d211,d212,d213,d214,d215,d216,d217,LXM,d218)
LED7(8)=Array(d221,d222,d223,d224,d225,d226,d227,LXM,d228)
LED7(9)=Array(d231,d232,d233,d234,d235,d236,d237,LXM,d238)
LED7(10)=Array(d241,d242,d243,d244,d245,d246,d247,LXM,d248)
LED7(11)=Array(d251,d252,d253,d254,d255,d256,d257,LXM,d258)
LED7(12)=Array(d261,d262,d263,d264,d265,d266,d267,LXM,d268)
LED7(13)=Array(d271,d272,d273,d274,d275,d276,d277,LXM,d278)

LED7(14)=Array(d311a,d312a,d313a,d314a,d315a,d316a,d317a,LXM,d318a)
LED7(15)=Array(d321a,d322a,d323a,d324a,d325a,d326a,d327a,LXM,d328a)
LED7(16)=Array(d331a,d332a,d333a,d334a,d335a,d336a,d337a,LXM,d338a)
LED7(17)=Array(d341,d342,d343,d344,d345,d346,d347,LXM,d348)
LED7(18)=Array(d351,d352,d353,d354,d355,d356,d357,LXM,d358)
LED7(19)=Array(d361,d362,d363,d364,d365,d366,d367,LXM,d368)
LED7(20)=Array(d371,d372,d373,d374,d375,d376,d377,LXM,d378)

LED7(21)=Array(d411,d412,d413,d414,d415,d416,d417,LXM,d418)
LED7(22)=Array(d421,d422,d423,d424,d425,d426,d427,LXM,d428)
LED7(23)=Array(d431,d432,d433,d434,d435,d436,d437,LXM,d438)
LED7(24)=Array(d441,d442,d443,d444,d445,d446,d447,LXM,d448)
LED7(25)=Array(d451,d452,d453,d454,d455,d456,d457,LXM,d458)
LED7(26)=Array(d461,d462,d463,d464,d465,d466,d467,LXM,d468)
LED7(27)=Array(d471,d472,d473,d474,d475,d476,d477,LXM,d478)

LED7(28)=Array(d511,d512,d513,d514,d515,d516,d517,LXM,d518)   'was 24 -- 26
LED7(29)=Array(d521,d522,d523,d524,d525,d526,d527,LXM,d528)   'was 25 -- 27
LED7(30)=Array(d611,d612,d613,d614,d615,d616,d617,LXM,d618)   'was 26 -- 24
LED7(31)=Array(d621,d622,d623,d624,d625,d626,d627,LXM,d628)   'was 27 -- 25

LED7(32)=Array(D281,D282,D283,D284,D285,D286,D287,LXM,D288)
LED7(33)=Array(D291,D292,D293,D294,D295,D296,D297,LXM,D298)
LED7(34)=Array(D301,D302,D303,D304,D305,D306,D307,LXM,D308)
LED7(35)=Array(D311,D312,D313,D314,D315,D316,D317,LXM,D318)
LED7(36)=Array(D321,D322,D323,D324,D325,D326,D327,LXM,D328)
LED7(37)=Array(D331,D332,D333,D334,D335,D336,D337,LXM,D338)

Sub DisplayTimer7_Timer
  Dim ChgLED, ii, num, chg, stat, obj
  ChgLED = Controller.ChangedLEDs (&Hffffffff, &Hffffffff)
  If Not IsEmpty (ChgLED) Then
    For ii = 0 To UBound (chgLED)
      num = chgLED (ii, 0) : chg = chgLED (ii, 1) : stat = chgLED (ii, 2)
      For Each obj In LED7 (num)
        If chg And 1 Then obj.State = stat And 1
        chg = chg \ 2 : stat = stat \ 2
      Next
    Next
  End If
End Sub




Dim Digits(34)

'Digits(0) = Array(LED1x0,LED1x1,LED1x2,LED1x3,LED1x4,LED1x5,LED1x6, n1, LED1x8)
Digits(0) = Array(LED2x0,LED2x1,LED2x2,LED2x3,LED2x4,LED2x5,LED2x6, n1, LED2x8)
Digits(1) = Array(LED3x0,LED3x1,LED3x2,LED3x3,LED3x4,LED3x5,LED3x6, n1, LED3x8)
Digits(2) = Array(LED4x0,LED4x1,LED4x2,LED4x3,LED4x4,LED4x5,LED4x6, n1, LED4x8)
Digits(3) = Array(LED5x0,LED5x1,LED5x2,LED5x3,LED5x4,LED5x5,LED5x6, n1, LED5x8)
Digits(4) = Array(LED6x0,LED6x1,LED6x2,LED6x3,LED6x4,LED6x5,LED6x6, n1, LED6x8)
Digits(5) = Array(LED7x0,LED7x1,LED7x2,LED7x3,LED7x4,LED7x5,LED7x6, n1, LED7x8)


'Digits(7) = Array(LED8x0,LED8x1,LED8x2,LED8x3,LED8x4,LED8x5,LED8x6, n1, LED8x8)
Digits(6) = Array(LED9x0,LED9x1,LED9x2,LED9x3,LED9x4,LED9x5,LED9x6, n1, LED9x8)
Digits(7) = Array(LED10x0,LED10x1,LED10x2,LED10x3,LED10x4,LED10x5,LED10x6, n1, LED10x8)
Digits(8) = Array(LED11x0,LED11x1,LED11x2,LED11x3,LED11x4,LED11x5,LED11x6, n1, LED11x8)
Digits(9) = Array(LED12x0,LED12x1,LED12x2,LED12x3,LED12x4,LED12x5,LED12x6, n1, LED12x8)
Digits(10) = Array(LED13x0,LED13x1,LED13x2,LED13x3,LED13x4,LED13x5,LED13x6, n1, LED13x8)
Digits(11) = Array(LED14x0,LED14x1,LED14x2,LED14x3,LED14x4,LED14x5,LED14x6, n1, LED14x8)

'Digits(14) = Array(LED1x000,LED1x001,LED1x002,LED1x003,LED1x004,LED1x005,LED1x006, n1, LED1x008)
Digits(12) = Array(LED1x100,LED1x101,LED1x102,LED1x103,LED1x104,LED1x105,LED1x106, n1, LED1x108)
Digits(13) = Array(LED1x200,LED1x201,LED1x202,LED1x203,LED1x204,LED1x205,LED1x206, n1, LED1x208)
Digits(14) = Array(LED1x300,LED1x301,LED1x302,LED1x303,LED1x304,LED1x305,LED1x306, n1, LED1x308)
Digits(15) = Array(LED1x400,LED1x401,LED1x402,LED1x403,LED1x404,LED1x405,LED1x406, n1, LED1x408)
Digits(16) = Array(LED1x500,LED1x501,LED1x502,LED1x503,LED1x504,LED1x505,LED1x506, n1, LED1x508)
Digits(17) = Array(LED1x600,LED1x601,LED1x602,LED1x603,LED1x604,LED1x605,LED1x606, n1, LED1x608)


'Digits(21) = Array(LED2x000,LED2x001,LED2x002,LED2x003,LED2x004,LED2x005,LED2x006, n1, LED2x008)
Digits(18) = Array(LED2x100,LED2x101,LED2x102,LED2x103,LED2x104,LED2x105,LED2x106, n1, LED2x108)
Digits(19) = Array(LED2x200,LED2x201,LED2x202,LED2x203,LED2x204,LED2x205,LED2x206, n1, LED2x208)
Digits(20) = Array(LED2x300,LED2x301,LED2x302,LED2x303,LED2x304,LED2x305,LED2x306, n1, LED2x308)
Digits(21) = Array(LED2x400,LED2x401,LED2x402,LED2x403,LED2x404,LED2x405,LED2x406, n1, LED2x408)
Digits(22) = Array(LED2x500,LED2x501,LED2x502,LED2x503,LED2x504,LED2x505,LED2x506, n1, LED2x508)
Digits(23) = Array(LED2x600,LED2x601,LED2x602,LED2x603,LED2x604,LED2x605,LED2x606, n1, LED2x608)

Digits(24) = Array(LEDax300,LEDax301,LEDax302,LEDax303,LEDax304,LEDax305,LEDax306, n1, LEDax308)
Digits(25) = Array(LEDbx400,LEDbx401,LEDbx402,LEDbx403,LEDbx404,LEDbx405,LEDbx406, n1, LEDbx408)

Digits(26) = Array(LEDcx500,LEDcx501,LEDcx502,LEDcx503,LEDcx504,LEDcx505,LEDcx506, n1, LEDcx508)
Digits(27) = Array(LEDdx600,LEDdx601,LEDdx602,LEDdx603,LEDdx604,LEDdx605,LEDdx606, n1, LEDdx608)

Digits(28)=Array(D281,D282,D283,D284,D285,D286,D287,LXM,D288)
Digits(29)=Array(D291,D292,D293,D294,D295,D296,D297,LXM,D298)
Digits(30)=Array(D301,D302,D303,D304,D305,D306,D307,LXM,D308)
Digits(31)=Array(D311,D312,D313,D314,D315,D316,D317,LXM,D318)
Digits(32)=Array(D321,D322,D323,D324,D325,D326,D327,LXM,D328)
Digits(33)=Array(D331,D332,D333,D334,D335,D336,D337,LXM,D338)



Sub DisplayTimerFSS6_Timer
  Dim ChgLED,ii,num,chg,stat,obj
  ChgLed = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
  If Not IsEmpty(ChgLED) Then
    'If DesktopMode = True Then
    For ii = 0 To UBound(chgLED)
      num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
      if (num < 28) then
        For Each obj In Digits(num)
          If chg And 1 Then obj.visible = stat And 1
          chg = chg\2 : stat = stat\2
        Next
      else
        For Each obj In Digits(num)
          If chg And 1 Then obj.state = stat And 1
          chg = chg\2 : stat = stat\2
        Next
      end if
    next
    'end if
  end if
End Sub


Dim Digits7(38)

Digits7(0) = Array(LED1x0,LED1x1,LED1x2,LED1x3,LED1x4,LED1x5,LED1x6, n1, LED1x8)
Digits7(1) = Array(LED2x0,LED2x1,LED2x2,LED2x3,LED2x4,LED2x5,LED2x6, n1, LED2x8)
Digits7(2) = Array(LED3x0,LED3x1,LED3x2,LED3x3,LED3x4,LED3x5,LED3x6, n1, LED3x8)
Digits7(3) = Array(LED4x0,LED4x1,LED4x2,LED4x3,LED4x4,LED4x5,LED4x6, n1, LED4x8)
Digits7(4) = Array(LED5x0,LED5x1,LED5x2,LED5x3,LED5x4,LED5x5,LED5x6, n1, LED5x8)
Digits7(5) = Array(LED6x0,LED6x1,LED6x2,LED6x3,LED6x4,LED6x5,LED6x6, n1, LED6x8)
Digits7(6) = Array(LED7x0,LED7x1,LED7x2,LED7x3,LED7x4,LED7x5,LED7x6, n1, LED7x8)


Digits7(7) = Array(LED8x0,LED8x1,LED8x2,LED8x3,LED8x4,LED8x5,LED8x6, n1, LED8x8)
Digits7(8) = Array(LED9x0,LED9x1,LED9x2,LED9x3,LED9x4,LED9x5,LED9x6, n1, LED9x8)
Digits7(9) = Array(LED10x0,LED10x1,LED10x2,LED10x3,LED10x4,LED10x5,LED10x6, n1, LED10x8)
Digits7(10) = Array(LED11x0,LED11x1,LED11x2,LED11x3,LED11x4,LED11x5,LED11x6, n1, LED11x8)
Digits7(11) = Array(LED12x0,LED12x1,LED12x2,LED12x3,LED12x4,LED12x5,LED12x6, n1, LED12x8)
Digits7(12) = Array(LED13x0,LED13x1,LED13x2,LED13x3,LED13x4,LED13x5,LED13x6, n1, LED13x8)
Digits7(13) = Array(LED14x0,LED14x1,LED14x2,LED14x3,LED14x4,LED14x5,LED14x6, n1, LED14x8)

Digits7(14) = Array(LED1x000,LED1x001,LED1x002,LED1x003,LED1x004,LED1x005,LED1x006, n1, LED1x008)
Digits7(15) = Array(LED1x100,LED1x101,LED1x102,LED1x103,LED1x104,LED1x105,LED1x106, n1, LED1x108)
Digits7(16) = Array(LED1x200,LED1x201,LED1x202,LED1x203,LED1x204,LED1x205,LED1x206, n1, LED1x208)
Digits7(17) = Array(LED1x300,LED1x301,LED1x302,LED1x303,LED1x304,LED1x305,LED1x306, n1, LED1x308)
Digits7(18) = Array(LED1x400,LED1x401,LED1x402,LED1x403,LED1x404,LED1x405,LED1x406, n1, LED1x408)
Digits7(19) = Array(LED1x500,LED1x501,LED1x502,LED1x503,LED1x504,LED1x505,LED1x506, n1, LED1x508)
Digits7(20) = Array(LED1x600,LED1x601,LED1x602,LED1x603,LED1x604,LED1x605,LED1x606, n1, LED1x608)


Digits7(21) = Array(LED2x000,LED2x001,LED2x002,LED2x003,LED2x004,LED2x005,LED2x006, n1, LED2x008)
Digits7(22) = Array(LED2x100,LED2x101,LED2x102,LED2x103,LED2x104,LED2x105,LED2x106, n1, LED2x108)
Digits7(23) = Array(LED2x200,LED2x201,LED2x202,LED2x203,LED2x204,LED2x205,LED2x206, n1, LED2x208)
Digits7(24) = Array(LED2x300,LED2x301,LED2x302,LED2x303,LED2x304,LED2x305,LED2x306, n1, LED2x308)
Digits7(25) = Array(LED2x400,LED2x401,LED2x402,LED2x403,LED2x404,LED2x405,LED2x406, n1, LED2x408)
Digits7(26) = Array(LED2x500,LED2x501,LED2x502,LED2x503,LED2x504,LED2x505,LED2x506, n1, LED2x508)
Digits7(27) = Array(LED2x600,LED2x601,LED2x602,LED2x603,LED2x604,LED2x605,LED2x606, n1, LED2x608)

Digits7(28) = Array(LEDax300,LEDax301,LEDax302,LEDax303,LEDax304,LEDax305,LEDax306, n1, LEDax308)
Digits7(29) = Array(LEDbx400,LEDbx401,LEDbx402,LEDbx403,LEDbx404,LEDbx405,LEDbx406, n1, LEDbx408)

Digits7(30) = Array(LEDcx500,LEDcx501,LEDcx502,LEDcx503,LEDcx504,LEDcx505,LEDcx506, n1, LEDcx508)
Digits7(31) = Array(LEDdx600,LEDdx601,LEDdx602,LEDdx603,LEDdx604,LEDdx605,LEDdx606, n1, LEDdx608)

Digits7(32)=Array(D281,D282,D283,D284,D285,D286,D287,LXM,D288)
Digits7(33)=Array(D291,D292,D293,D294,D295,D296,D297,LXM,D298)
Digits7(34)=Array(D301,D302,D303,D304,D305,D306,D307,LXM,D308)
Digits7(35)=Array(D311,D312,D313,D314,D315,D316,D317,LXM,D318)
Digits7(36)=Array(D321,D322,D323,D324,D325,D326,D327,LXM,D328)
Digits7(37)=Array(D331,D332,D333,D334,D335,D336,D337,LXM,D338)


Sub DisplayTimerFSS7_Timer
  Dim ChgLED,ii,num,chg,stat,obj
  ChgLed = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
  If Not IsEmpty(ChgLED) Then
    'If DesktopMode = True Then
    For ii = 0 To UBound(chgLED)
      num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
      if (num < 32) then
        For Each obj In Digits7(num)
          If chg And 1 Then obj.visible = stat And 1
          chg = chg\2 : stat = stat\2
        Next
      else
        For Each obj In Digits7(num)
          If chg And 1 Then obj.state = stat And 1
          chg = chg\2 : stat = stat\2
        Next
      end if
    next
    'end if
  end if
End Sub




'**********************************
' ZMAT: General Math Functions
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

Dim LF: Set LF = New FlipperPolarity
Dim RF: Set RF = New FlipperPolarity
Dim LF2: Set LF2 = New FlipperPolarity
Dim RF2: Set RF2 = New FlipperPolarity
Dim ULF: Set ULF = New FlipperPolarity
Dim URF: Set URF = New FlipperPolarity

InitPolarity


'*******************************************
' Late 70's to early 80's

Sub InitPolarity()
   dim x, a : a = Array(LF, RF, LF2, RF2, ULF, URF)
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
    LF2.SetObjects "LF2", LeftFlipper2, TriggerLF2
    RF2.SetObjects "RF2", RightFlipper2, TriggerRF2
    ULF.SetObjects "ULF", FlipperUL, TriggerULF
    URF.SetObjects "URF", FlipperUR, TriggerURF
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
' To add the flipper tricks you must
'  - Include a call to FlipperCradleCollision from within OnBallBallCollision subroutine
'  - Include a call the CheckLiveCatch from the LeftFlipper_Collide and RightFlipper_Collide subroutines

RightFlipper.timerinterval = 1
Rightflipper.timerenabled = True

Sub RightFlipper_timer()
  FlipperTricks LeftFlipper, LFPress, LFCount, LFEndAngle, LFState
  FlipperTricks RightFlipper, RFPress, RFCount, RFEndAngle, RFState
  FlipperTricks LeftFlipper2, LFPress2, LFCount2, LFEndAngle2, LFState2
  FlipperTricks RightFlipper2, RFPress2, RFCount2, RFEndAngle2, RFState2
  FlipperTricks FlipperUL, LFPress3, LFCount3, LFEndAngle3, LFState3
  FlipperTricks FlipperUR, RFPress3, RFCount3, RFEndAngle3, RFState3
  'FlipperTricks FlipperLL, LFPress4, LFCount4, LFEndAngle4, LFState4
  'FlipperTricks FlipperLR, RFPress4, RFCount4, RFEndAngle4, RFState4

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

Function AnglePPP(ax,ay,bx,by,cx,cy) 'Angle for three points where a is the vertex
  Dim Pab, Pac, Pbc

  Pab = sqr((ax - bx)^2 + (ay - by)^2)
  Pac = sqr((ax - cx)^2 + (ay - cy)^2)
  Pbc = sqr((bx - cx)^2 + (by - cy)^2)

  AnglePPP = ArcCos((Pab^2 + Pac^2 - Pbc^2)/(2 * Pab * Pac)) * 180 / PI
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
Dim LFPress2, RFPress2, LFCount2, RFCount2
Dim LFPress3, RFPress3, LFCount3, RFCount3
Dim LFPress4, RFPress4, LFCount4, RFCount4

Dim LFState, RFState
Dim LFState2, RFState2
Dim LFState3, RFState3
Dim LFState4, RFState4

Dim RFEndAngle, LFEndAngle
Dim RFEndAngle2, LFEndAngle2
Dim RFEndAngle3, LFEndAngle3
Dim RFEndAngle4, LFEndAngle4

Dim EOST, EOSA,Frampup, FElasticity,FReturn


Const FlipperCoilRampupMode = 0 '0 = fast, 1 = medium, 2 = slow (tap passes should work)

LFState = 1
RFState = 1
EOST = leftflipper.eostorque
EOSA = leftflipper.eostorqueangle
Frampup = LeftFlipper.rampup
FElasticity = LeftFlipper.elasticity
FReturn = LeftFlipper.return
Const EOSTnew = 1.5 'EM's to late 80's
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
'Const EOSReturn = 0.055  'EM's
Const EOSReturn = 0.045  'late 70's to mid 80's
'Const EOSReturn = 0.035  'mid 80's to early 90's
'Const EOSReturn = 0.025  'mid 90's and later

LFEndAngle = Leftflipper.endangle
RFEndAngle = RightFlipper.endangle
LFEndAngle2 = Leftflipper2.endangle
RFEndAngle2 = RightFlipper2.endangle
LFEndAngle3 = FlipperUL.endangle
RFEndAngle3 = FlipperUR.endangle
LFEndAngle4 = FlipperLL.endangle
RFEndAngle4 = FlipperLR.endangle


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

    For Each b In allBalls
      If b.id >= HighestID Then highestID = b.id
    Next

    If UBound(ballvel) < highestID Then ReDim ballvel(highestID)  'set bounds
    If UBound(ballvelx) < highestID Then ReDim ballvelx(highestID)  'set bounds
    If UBound(ballvely) < highestID Then ReDim ballvely(highestID)  'set bounds

    For Each b In allBalls
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
Const TargetBouncerFactor = 0.7  'Level of bounces. Recommmended value of 0.7

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
'   ZANI: Misc Animations
'******************************************************


''''' Flippers


Sub LeftFlipper_Animate
  Dim a : a = LeftFlipper.CurrentAngle
  FlipperLSh.RotZ = a

  Dim v, BP
  v = 255.0 * (122.0 -  LeftFlipper.CurrentAngle) / (122.0 -  75.0)

  For each BP in BP_PLF
    BP.Rotz = a
    BP.visible = v < 128.0
  Next
  For each BP in BP_PLFup
    BP.Rotz = a
    BP.visible = v >= 128.0
  Next
End Sub

Sub RightFlipper_Animate
  Dim a : a = RightFlipper.CurrentAngle
  FlipperRSh.RotZ = a

  Dim v, BP
  v = 255.0 * (-122.0 -  RightFlipper.CurrentAngle) / (-122.0 +  75.0)

  For each BP in BP_PRF
    BP.Rotz = a
    BP.visible = v < 128.0
  Next
  For each BP in BP_PRFup
    BP.Rotz = a
    BP.visible = v >= 128.0
  Next
End Sub

Sub LeftFlipper2_Animate
  Dim a : a = LeftFlipper2.CurrentAngle
  FlipperLSh2.RotZ = a

  Dim v, BP
  v = 255.0 * (121.0 -  LeftFlipper2.CurrentAngle) / (121.0 -  75.0)

  For each BP in BP_PLF2
    BP.Rotz = a
    BP.visible = v < 128.0
  Next
  For each BP in BP_PLF2up
    BP.Rotz = a
    BP.visible = v >= 128.0
  Next
End Sub

Sub RightFlipper2_Animate
  Dim a : a = RightFlipper2.CurrentAngle
  FlipperRSh2.RotZ = a

  Dim v, BP
  v = 255.0 * (-121.0 -  RightFlipper2.CurrentAngle) / (-121.0 +  75.0)

  For each BP in BP_PRF2
    BP.Rotz = a
    BP.visible = v < 128.0
  Next
  For each BP in BP_PRF2up
    BP.Rotz = a
    BP.visible = v >= 128.0
  Next
End Sub


'TODO: add upper PF flipper shadows?
Sub FlipperUL_Animate
  Dim a : a = FlipperUL.CurrentAngle
' FlipperULSh.RotZ = a

  Dim v, BP
  v = 255.0 * (-126.0 -  FlipperUL.CurrentAngle) / (-126.0 +  80.0)

  For each BP in BP_PULF
    BP.Rotz = a
    BP.visible = v < 128.0
  Next
  For each BP in BP_PULFup
    BP.Rotz = a
    BP.visible = v >= 128.0
  Next
End Sub

Sub FlipperUR_Animate
  Dim a : a = FlipperUR.CurrentAngle
' FlipperURSh.RotZ = a

  Dim v, BP
  v = 255.0 * (115.0 -  FlipperUR.CurrentAngle) / (115.0 -  78.0)

  For each BP in BP_PURF
    BP.Rotz = a
    BP.visible = v < 128.0
  Next
  For each BP in BP_PURFup
    BP.Rotz = a
    BP.visible = v >= 128.0
  Next
End Sub



Sub FlipperLL_Animate
  Dim a : a = FlipperLL.CurrentAngle

  Dim v, BP
  v = 255.0 * (55.0 -  FlipperLL.CurrentAngle) / (55.0 - 101.0)

  For each BP in BP_PLLF
    BP.Rotz = a
    BP.visible = v < 128.0
  Next
  For each BP in BP_PLLFup
    BP.Rotz = a
    BP.visible = v >= 128.0
  Next
End Sub

Sub FlipperLR_Animate
  Dim a : a = FlipperLR.CurrentAngle

  Dim v, BP
  v = 255.0 * (-60.0 -  FlipperLR.CurrentAngle) / (-60.0 +  106.0)

  For each BP in BP_pLRF
    BP.Rotz = a
    BP.visible = v < 128.0
  Next
  For each BP in BP_pLRFUup
    BP.Rotz = a
    BP.visible = v >= 128.0
  Next
End Sub




''''' Bumper Animations

Sub Bumper1_Animate
  Dim z, BP
  z = Bumper1.CurrentRingOffset
  For Each BP in BP_BRing1 : BP.transz = z: Next
End Sub

Sub Bumper2_Animate
  Dim z, BP
  z = Bumper2.CurrentRingOffset
  For Each BP in BP_BRing2 : BP.transz = z: Next
End Sub

Sub Bumper3_Animate
  Dim z, BP
  z = Bumper3.CurrentRingOffset
  For Each BP in BP_BRing3 : BP.transz = z: Next
End Sub

Sub Bumper4_Animate
  Dim z, BP
  z = Bumper4.CurrentRingOffset
  For Each BP in BP_BRing4 : BP.transz = z: Next
End Sub

Dim Bumpers : Bumpers = Array(Bumper1, Bumper2, Bumper3, Bumper4)

Sub AnimateBumperSkirts
  dim r, g, s, x, y, b, tz
  ' Animate Bumper switch (experimental)
  For r = 0 To 3
    g = 10000.
    For s = 0 to UBound(gBOT)
      If r<3 OR (r=3 and gBOT(s).z < 0)  Then  'deal with lower pf bumper
        x = Bumpers(r).x - gBOT(s).x
        y = Bumpers(r).y - gBOT(s).y
        b = x * x + y * y
        If b < g Then g = b
      End If
    Next
    tz = 0
    If g < 80 * 80 Then
      tz = 4
    End If
    If r = 0 Then For Each x in BP_BSkirt1: x.transZ = tz: Next
    If r = 1 Then For Each x in BP_BSkirt2: x.transZ = tz: Next
    If r = 2 Then For Each x in BP_BSkirt3: x.transZ = tz: Next
    If r = 3 Then For Each x in BP_BSkirt4: x.transZ = tz: Next
  Next
End Sub



''''' Gate Animations

Sub Gate1_Animate
  Dim a : a = Gate1.CurrentAngle
  Dim BP : For Each BP in BP_p_Gate1 : BP.rotx = a: Next
End Sub

Sub Gate2_Animate
  Dim a : a = Gate2.CurrentAngle
  Dim BP : For Each BP in BP_p_Gate2 : BP.rotx = a: Next
End Sub

Sub Gate3_Animate
  Dim a : a = Gate3.CurrentAngle
  Dim BP : For Each BP in BP_p_Gate3 : BP.rotx = a: Next
  For Each BP in BP_pSSpring2 : BP.rotx = -a*4/90: Next
End Sub

Sub Gate4_Animate
  Dim a : a = Gate4.CurrentAngle
  Dim BP : For Each BP in BP_p_Gate4 : BP.rotx = a: Next
  For Each BP in BP_pSSpring : BP.rotx = -a*4/90: Next
End Sub



''''' Switch Animations

Sub sw04_Animate
  Dim z : z = sw04.CurrentAnimOffset
  Dim BP : For Each BP in BP_psw04 : BP.transz = z: Next
End Sub

Sub sw21_Animate
  Dim z : z = sw21.CurrentAnimOffset
  Dim BP : For Each BP in BP_psw21 : BP.transz = z: Next
End Sub

Sub sw34_Animate
  Dim z : z = sw34.CurrentAnimOffset
  Dim BP : For Each BP in BP_psw34 : BP.transz = z: Next
End Sub

Sub sw54_Animate
  Dim z : z = sw54.CurrentAnimOffset
  Dim BP : For Each BP in BP_psw54 : BP.transz = z: Next
End Sub



''''' Standup Targets

Sub UpdateStandupTargets
  dim BP, ty

  ty = BM_pSW03.transy
  For each BP in BP_psw03 : BP.transy = ty: Next

  ty = BM_pSW13.transy
  For each BP in BP_psw13 : BP.transy = ty: Next

  ty = BM_pSW23.transy
  For each BP in BP_psw23 : BP.transy = ty: Next

  ty = BM_pSW33.transy
  For each BP in BP_psw33 : BP.transy = ty: Next

  ty = BM_pSW05.transy
  For each BP in BP_psw05 : BP.transy = ty: Next

  ty = BM_pSW15.transy
  For each BP in BP_psw15 : BP.transy = ty: Next

  ty = BM_pSW55.transy
  For each BP in BP_psw55 : BP.transy = ty: Next

  ty = BM_pSW50.transy
  For each BP in BP_psw50 : BP.transy = ty: Next

  ty = BM_pSW60.transy
  For each BP in BP_psw60 : BP.transy = ty: Next

  ty = BM_pSW11.transy
  For each BP in BP_psw11 : BP.transy = ty: Next

  ty = BM_pSW14.transy
  For each BP in BP_psw14 : BP.transy = ty: Next

  ty = BM_pSW24.transy
  For each BP in BP_psw24 : BP.transy = ty: Next

  ty = BM_pSW25.transy
  For each BP in BP_psw25 : BP.transy = ty: Next

End Sub



''''' Drop Targets

Sub UpdateDropTargets
  dim BP, tz, rx, ry

  tz = BM_psw02.transz
  rx = BM_psw02.rotx
  ry = BM_psw02.roty
  For each BP in BP_psw02: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_psw12.transz
  rx = BM_psw12.rotx
  ry = BM_psw12.roty
  For each BP in BP_psw12 : BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_psw22.transz
  rx = BM_psw22.rotx
  ry = BM_psw22.roty
  For each BP in BP_psw22 : BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_psw32.transz
  rx = BM_psw32.rotx
  ry = BM_psw32.roty
  For each BP in BP_psw32 : BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_psw00.transz
  rx = BM_psw00.rotx
  ry = BM_psw00.roty
  For each BP in BP_psw00 : BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_psw10.transz
  rx = BM_psw10.rotx
  ry = BM_psw10.roty
  For each BP in BP_psw10 : BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_psw20.transz
  rx = BM_psw20.rotx
  ry = BM_psw20.roty
  For each BP in BP_psw20 : BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_psw30.transz
  rx = BM_psw30.rotx
  ry = BM_psw30.roty
  For each BP in BP_psw30 : BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_psw40.transz
  rx = BM_psw40.rotx
  ry = BM_psw40.roty
  For each BP in BP_psw40 : BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next
End Sub




'***********************************************************************
'*   ZOPT: Table Options
'***********************************************************************


Dim LightLevel : LightLevel = 50      ' Level of room lighting (0 to 100), where 0 is dark and 100 is brightest
Dim UpperGIColor : UpperGIColor = 1   ' 0=Random, 1=Normal, 2=Red, 3=Orange, 4=Green, 5=Blue, 6=Purple
Dim MainGIColor : MainGIColor = 1     ' 0=Random, 1=Normal, 2=Red, 3=Orange, 4=Green, 5=Blue, 6=Purple
Dim LowerGIColor : LowerGIColor = 1   ' 0=Random, 1=Normal, 2=Red, 3=Orange, 4=Green, 5=Blue, 6=Purple
Dim BumperLightColor : BumperLightColor = 1 ' 0=Random, 1=Normal, 2=Red, 3=Orange, 4=Green, 5=Blue, 6=Purple
Dim CCPosts : CCPosts = 0         ' Color changing posts (only for two post top left of Table), 0=Not changing, 1=Color changing
Dim BallMod : BallMod = 0         ' 0=Normal Ball, 1=Marbled ball Green / Red, 2=Marbled ball Green / Yellow, 3 to 11 = Glowball
Dim FlipperKeyMod : FlipperKeyMod = 1   ' 1 = Standard flipper buttons (normal main lower and upper manga ), 2=LazyMan flippers (only normal button for all flippers)
Dim InstCards : InstCards = 0       ' 0 = Standard , 1=Mod1, 2=Mod2 3Ball, 3=Mod2 5ball
Dim InfoCard : InfoCard = 1       ' 0 = Dont show, 1 = Show
Dim Dust : Dust = 1           ' 0 = Dont show, 1 = Show

Dim MechVolume : MechVolume = 0.8     'Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Dim BallRollVolume : BallRollVolume = 0.5   'Level of ball rolling volume. Value between 0 and 1
Dim RampRollVolume : RampRollVolume = 0.5   'Level of ramp rolling volume. Value between 0 and 1

Dim InitDips: InitDips = 0

' Base options
Const Opt_Light = 0
Const Opt_UGIcolor = 1
Const Opt_MGIcolor = 2
Const Opt_LGIcolor = 3
Const Opt_BLColor = 4
Const Opt_CCPosts = 5
Const Opt_Ball = 6
Const Opt_FlipKeys = 7
Const Opt_InstCards = 8
Const Opt_InfoCard = 9
Const Opt_Dust = 10

Const Opt_Volume = 11
Const Opt_Volume_Ramp = 12
Const Opt_Volume_Ball = 13
Const Opt_Info_1 = 14
Const Opt_Info_2 = 15

Const NOptions = 16

Const FlexDMD_RenderMode_DMD_GRAY_2 = 0
Const FlexDMD_RenderMode_DMD_GRAY_4 = 1
Const FlexDMD_RenderMode_DMD_RGB = 2
Const FlexDMD_RenderMode_SEG_2x16Alpha = 3
Const FlexDMD_RenderMode_SEG_2x20Alpha = 4
Const FlexDMD_RenderMode_SEG_2x7Alpha_2x7Num = 5
Const FlexDMD_RenderMode_SEG_2x7Alpha_2x7Num_4x1Num = 6
Const FlexDMD_RenderMode_SEG_2x7Num_2x7Num_4x1Num = 7
Const FlexDMD_RenderMode_SEG_2x7Num_2x7Num_10x1Num = 8
Const FlexDMD_RenderMode_SEG_2x7Num_2x7Num_4x1Num_gen7 = 9
Const FlexDMD_RenderMode_SEG_2x7Num10_2x7Num10_4x1Num = 10
Const FlexDMD_RenderMode_SEG_2x6Num_2x6Num_4x1Num = 11
Const FlexDMD_RenderMode_SEG_2x6Num10_2x6Num10_4x1Num = 12
Const FlexDMD_RenderMode_SEG_4x7Num10 = 13
Const FlexDMD_RenderMode_SEG_6x4Num_4x1Num = 14
Const FlexDMD_RenderMode_SEG_2x7Num_4x1Num_1x16Alpha = 15
Const FlexDMD_RenderMode_SEG_1x16Alpha_1x16Num_1x7Num = 16

Const FlexDMD_Align_TopLeft = 0
Const FlexDMD_Align_Top = 1
Const FlexDMD_Align_TopRight = 2
Const FlexDMD_Align_Left = 3
Const FlexDMD_Align_Center = 4
Const FlexDMD_Align_Right = 5
Const FlexDMD_Align_BottomLeft = 6
Const FlexDMD_Align_Bottom = 7
Const FlexDMD_Align_BottomRight = 8

Const FlexDMD_Scaling_Fit = 0
Const FlexDMD_Scaling_Fill = 1
Const FlexDMD_Scaling_FillX = 2
Const FlexDMD_Scaling_FillY = 3
Const FlexDMD_Scaling_Stretch = 4
Const FlexDMD_Scaling_StretchX = 5
Const FlexDMD_Scaling_StretchY = 6
Const FlexDMD_Scaling_None = 7

Const FlexDMD_Interpolation_Linear = 0
Const FlexDMD_Interpolation_ElasticIn = 1
Const FlexDMD_Interpolation_ElasticOut = 2
Const FlexDMD_Interpolation_ElasticInOut = 3
Const FlexDMD_Interpolation_QuadIn = 4
Const FlexDMD_Interpolation_QuadOut = 5
Const FlexDMD_Interpolation_QuadInOut = 6
Const FlexDMD_Interpolation_CubeIn = 7
Const FlexDMD_Interpolation_CubeOut = 8
Const FlexDMD_Interpolation_CubeInOut = 9
Const FlexDMD_Interpolation_QuartIn = 10
Const FlexDMD_Interpolation_QuartOut = 11
Const FlexDMD_Interpolation_QuartInOut = 12
Const FlexDMD_Interpolation_QuintIn = 13
Const FlexDMD_Interpolation_QuintOut = 14
Const FlexDMD_Interpolation_QuintInOut = 15
Const FlexDMD_Interpolation_SineIn = 16
Const FlexDMD_Interpolation_SineOut = 17
Const FlexDMD_Interpolation_SineInOut = 18
Const FlexDMD_Interpolation_BounceIn = 19
Const FlexDMD_Interpolation_BounceOut = 20
Const FlexDMD_Interpolation_BounceInOut = 21
Const FlexDMD_Interpolation_CircIn = 22
Const FlexDMD_Interpolation_CircOut = 23
Const FlexDMD_Interpolation_CircInOut = 24
Const FlexDMD_Interpolation_ExpoIn = 25
Const FlexDMD_Interpolation_ExpoOut = 26
Const FlexDMD_Interpolation_ExpoInOut = 27
Const FlexDMD_Interpolation_BackIn = 28
Const FlexDMD_Interpolation_BackOut = 29
Const FlexDMD_Interpolation_BackInOut = 30

Dim OptionDMD: Set OptionDMD = Nothing
Dim bOptionsMagna, bInOptions : bOptionsMagna = False
Dim OptPos, OptSelected, OptN, OptTop, OptBot, OptSel
Dim OptFontHi, OptFontLo

Sub Options_Open
  bOptionsMagna = False
  On Error Resume Next
  Set OptionDMD = CreateObject("FlexDMD.FlexDMD")
  On Error Goto 0
  If OptionDMD is Nothing Then
    Debug.Print "FlexDMD is not installed"
    Debug.Print "Option UI can not be opened"
    MsgBox "You need to install FlexDMD to access table options"
    Exit Sub
  End If
  If HH.ShowDT Then OptionDMDFlasher.RotX = -25
  bInOptions = True
  OptPos = 0
  OptSelected = False
  OptionDMD.Show = False
  OptionDMD.RenderMode = FlexDMD_RenderMode_DMD_GRAY_4
  OptionDMD.Width = 128
  OptionDMD.Height = 32
  OptionDMD.Clear = True
  OptionDMD.Run = True
  Dim a, scene, font
  Set scene = OptionDMD.NewGroup("Scene")
  Set OptFontHi = OptionDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", vbWhite, vbWhite, 0)
  Set OptFontLo = OptionDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", RGB(100, 100, 100), RGB(100, 100, 100), 0)
  Set OptSel = OptionDMD.NewGroup("Sel")
  Set a = OptionDMD.NewLabel(">", OptFontLo, ">>>")
  a.SetAlignedPosition 1, 16, FlexDMD_Align_Left
  OptSel.AddActor a
  Set a = OptionDMD.NewLabel(">", OptFontLo, "<<<")
  a.SetAlignedPosition 127, 16, FlexDMD_Align_Right
  OptSel.AddActor a
  scene.AddActor OptSel
  OptSel.SetBounds 0, 0, 128, 32
  OptSel.Visible = False

  Set a = OptionDMD.NewLabel("Info1", OptFontLo, "MAGNA EXIT/ENTER")
  a.SetAlignedPosition 1, 32, FlexDMD_Align_BottomLeft
  scene.AddActor a
  Set a = OptionDMD.NewLabel("Info2", OptFontLo, "FLIPPER SELECT")
  a.SetAlignedPosition 127, 32, FlexDMD_Align_BottomRight
  scene.AddActor a
  Set OptN = OptionDMD.NewLabel("Pos", OptFontLo, "LINE 1")
  Set OptTop = OptionDMD.NewLabel("Top", OptFontLo, "LINE 1")
  Set OptBot = OptionDMD.NewLabel("Bottom", OptFontLo, "LINE 2")
  scene.AddActor OptN
  scene.AddActor OptTop
  scene.AddActor OptBot
  Options_OnOptChg
  OptionDMD.LockRenderThread
  OptionDMD.Stage.AddActor scene
  OptionDMD.UnlockRenderThread
  OptionDMDFlasher.Visible = True
  DisableStaticPrerendering = True
End Sub

Sub Options_UpdateDMD
  If OptionDMD is Nothing Then Exit Sub
  Dim DMDp: DMDp = OptionDMD.DmdPixels
  If Not IsEmpty(DMDp) Then
    OptionDMDFlasher.DMDWidth = OptionDMD.Width
    OptionDMDFlasher.DMDHeight = OptionDMD.Height
    OptionDMDFlasher.DMDPixels = DMDp
  End If
End Sub

Sub Options_Close
  bInOptions = False
  OptionDMDFlasher.Visible = False
  If OptionDMD is Nothing Then Exit Sub
  OptionDMD.Run = False
  Set OptionDMD = Nothing
  DisableStaticPrerendering = False
End Sub

Function Options_OnOffText(opt)
  If opt Then
    Options_OnOffText = "ON"
  Else
    Options_OnOffText = "OFF"
  End If
End Function

Sub Options_OnOptChg
  If OptionDMD is Nothing Then Exit Sub
  OptionDMD.LockRenderThread
  OptN.Text = (OptPos+1) & "/" & (NOptions)
  If OptSelected Then
    OptTop.Font = OptFontLo
    OptBot.Font = OptFontHi
    OptSel.Visible = True
  Else
    OptTop.Font = OptFontHi
    OptBot.Font = OptFontLo
    OptSel.Visible = False
  End If
  If OptPos = Opt_Light Then
    OptTop.Text = "LIGHT LEVEL"
    OptBot.Text = "LEVEL " & LightLevel
    SaveValue cGameName, "LIGHT", LightLevel
  ElseIf OptPos = Opt_UGIcolor Then
    OptTop.Text = "UPPER GI COLOR"
    If UpperGIColor = 0 Then OptBot.Text = "RANDOM"
    If UpperGIColor = 1 Then OptBot.Text = "NORMAL"
    If UpperGIColor = 2 Then OptBot.Text = "RED"
    If UpperGIColor = 3 Then OptBot.Text = "ORANGE"
    If UpperGIColor = 4 Then OptBot.Text = "GREEN"
    If UpperGIColor = 5 Then OptBot.Text = "BLUE"
    If UpperGIColor = 6 Then OptBot.Text = "PURPLE"
    If UpperGIColor = 7 Then OptBot.Text = "SPOOKY"
    SaveValue cGameName, "UGICOLOR", UpperGIColor
  ElseIf OptPos = Opt_MGIcolor Then
    OptTop.Text = "MAIN GI COLOR"
    If MainGIColor = 0 Then OptBot.Text = "RANDOM"
    If MainGIColor = 1 Then OptBot.Text = "NORMAL"
    If MainGIColor = 2 Then OptBot.Text = "RED"
    If MainGIColor = 3 Then OptBot.Text = "ORANGE"
    If MainGIColor = 4 Then OptBot.Text = "GREEN"
    If MainGIColor = 5 Then OptBot.Text = "BLUE"
    If MainGIColor = 6 Then OptBot.Text = "PURPLE"
    SaveValue cGameName, "MGICOLOR", MainGIColor
  ElseIf OptPos = Opt_LGIcolor Then
    OptTop.Text = "LOWER GI COLOR"
    If LowerGIColor = 0 Then OptBot.Text = "RANDOM"
    If LowerGIColor = 1 Then OptBot.Text = "NORMAL"
    If LowerGIColor = 2 Then OptBot.Text = "RED"
    If LowerGIColor = 3 Then OptBot.Text = "ORANGE"
    If LowerGIColor = 4 Then OptBot.Text = "GREEN"
    If LowerGIColor = 5 Then OptBot.Text = "BLUE"
    If LowerGIColor = 6 Then OptBot.Text = "PURPLE"
    SaveValue cGameName, "LGICOLOR", LowerGIColor
  ElseIf OptPos = Opt_BLColor Then
    OptTop.Text = "BUMPER LIGHT COLOR"
    If BumperLightColor = 0 Then OptBot.Text = "RANDOM"
    If BumperLightColor = 1 Then OptBot.Text = "NORMAL"
    If BumperLightColor = 2 Then OptBot.Text = "RED"
    If BumperLightColor = 3 Then OptBot.Text = "ORANGE"
    If BumperLightColor = 4 Then OptBot.Text = "GREEN"
    If BumperLightColor = 5 Then OptBot.Text = "BLUE"
    If BumperLightColor = 6 Then OptBot.Text = "PURPLE"
    SaveValue cGameName, "BLCOLOR", BumperLightColor
  ElseIf OptPos = Opt_CCPosts Then
    OptTop.Text = "COLOR CHANGING POSTS"
    If CCPosts = 1 Then
      OptBot.Text = "ON"
    Else
      OptBot.Text = "OFF"
    End If
    SaveValue cGameName, "CCPOSTS", CCPosts
  ElseIf OptPos = Opt_Ball Then
    OptTop.Text = "BALL"
    If BallMod = 0 Then OptBot.Text = "NORMAL"
    If BallMod = 1 Then OptBot.Text = "GREEN RED MARBLE"
    If BallMod = 2 Then OptBot.Text = "GREEN YELLOW MARBLE"
    If BallMod = 3 Then OptBot.Text = "RED GLOWBALL"
    If BallMod = 4 Then OptBot.Text = "ORANGE GLOWBALL"
    If BallMod = 5 Then OptBot.Text = "YELLOW GLOWBALL"
    If BallMod = 6 Then OptBot.Text = "GREEN1 GLOWBALL"
    If BallMod = 7 Then OptBot.Text = "GREEN2 GLOWBALL"
    If BallMod = 8 Then OptBot.Text = "BLUE GLOWBALL"
    If BallMod = 9 Then OptBot.Text = "PINK GLOWBALL"
    If BallMod = 10 Then OptBot.Text = "PURPLE GLOWBALL"
    If BallMod = 11 Then OptBot.Text = "MULTICOLOR GLOWBALL"
    SaveValue cGameName, "BALLMOD", BallMod
  ElseIf OptPos = Opt_FlipKeys Then
    OptTop.Text = "FLIPPER MODE"
    If FlipperKeyMod = 2 Then
      OptBot.Text = "NO MAGNAS"
    Else
      OptBot.Text = "ORIGINAL (WITH MAGNAS)"
    End If
    SaveValue cGameName, "LAZYFLIPS", FlipperKeyMod
  ElseIf OptPos = Opt_InstCards Then
    OptTop.Text = "INSTRUCTION CARDS"
    If InstCards = 0 Then OptBot.Text = "STANDARD"
    If InstCards = 1 Then OptBot.Text = "MOD 1"
    If InstCards = 2 Then OptBot.Text = "MOD 2 - 3 BALL"
    If InstCards = 3 Then OptBot.Text = "MOD 2 - 5 BALL"
    SaveValue cGameName, "INSTCARDS", InstCards
  ElseIf OptPos = Opt_InfoCard Then
    OptTop.Text = "INFO CARD"
    If InfoCard = 1 Then
      OptBot.Text = "ON"
    Else
      OptBot.Text = "OFF"
    End If
    SaveValue cGameName, "INFOCARD", InfoCard
  ElseIf OptPos = Opt_Dust Then
    OptTop.Text = "PLAYFIELD DUST"
    If Dust = 1 Then
      OptBot.Text = "ON"
    Else
      OptBot.Text = "OFF"
    End If
    SaveValue cGameName, "DUST", Dust
  ElseIf OptPos = Opt_Volume Then
    OptTop.Text = "MECH VOLUME"
    OptBot.Text = "LEVEL " & CInt(MechVolume * 100)
    SaveValue cGameName, "VOLUME", MechVolume
  ElseIf OptPos = Opt_Volume_Ramp Then
    OptTop.Text = "RAMP VOLUME"
    OptBot.Text = "LEVEL " & CInt(RampRollVolume * 100)
    SaveValue cGameName, "RAMPVOLUME", RampRollVolume
  ElseIf OptPos = Opt_Volume_Ball Then
    OptTop.Text = "BALL VOLUME"
    OptBot.Text = "LEVEL " & CInt(BallRollVolume * 100)
    SaveValue cGameName, "BALLVOLUME", BallRollVolume
  ElseIf OptPos = Opt_Info_1 Then
    OptTop.Text = "VPX " & VersionMajor & "." & VersionMinor & "." & VersionRevision
    OptBot.Text = "HAUNTED HOUSE " & TableVersion
  ElseIf OptPos = Opt_Info_2 Then
    OptTop.Text = "RENDER MODE"
    If RenderingMode = 0 Then OptBot.Text = "DEFAULT"
    If RenderingMode = 1 Then OptBot.Text = "STEREO 3D"
    If RenderingMode = 2 Then OptBot.Text = "VR"
  End If
  OptTop.Pack
  OptTop.SetAlignedPosition 127, 1, FlexDMD_Align_TopRight
  OptBot.SetAlignedPosition 64, 16, FlexDMD_Align_Center
  OptionDMD.UnlockRenderThread
  UpdateMods
End Sub

Sub Options_Toggle(amount)
  If OptionDMD is Nothing Then Exit Sub
  If OptPos = Opt_Light Then
    LightLevel = LightLevel + amount * 10
    If LightLevel < 0 Then LightLevel = 100
    If LightLevel > 100 Then LightLevel = 0
  ElseIf OptPos = Opt_UGIcolor Then
    UpperGIColor = UpperGIColor + amount * 1
    If UpperGIColor < 0 Then UpperGIColor = 7
    If UpperGIColor > 7 Then UpperGIColor = 0
  ElseIf OptPos = Opt_MGIcolor Then
    MainGIColor = MainGIColor + amount * 1
    If MainGIColor < 0 Then MainGIColor = 6
    If MainGIColor > 6 Then MainGIColor = 0
  ElseIf OptPos = Opt_LGIcolor Then
    LowerGIColor = LowerGIColor + amount * 1
    If LowerGIColor < 0 Then LowerGIColor = 6
    If LowerGIColor > 6 Then LowerGIColor = 0
  ElseIf OptPos = Opt_BLColor Then
    BumperLightColor = BumperLightColor + amount * 1
    If BumperLightColor < 0 Then BumperLightColor = 6
    If BumperLightColor > 6 Then BumperLightColor = 0
  ElseIf OptPos = Opt_CCPosts Then
    If CCPosts = 1 Then CCPosts = 0 Else CCPosts = 1
  ElseIf OptPos = Opt_Ball Then
    BallMod = BallMod + amount * 1
    If BallMod < 0 Then BallMod = 11
    If BallMod > 11 Then BallMod = 0
  ElseIf OptPos = Opt_FlipKeys Then
    If FlipperKeyMod = 2 Then FlipperKeyMod = 1 Else FlipperKeyMod = 2
  ElseIf OptPos = Opt_InstCards Then
    InstCards = InstCards + amount * 1
    If InstCards < 0 Then InstCards = 3
    If InstCards > 3 Then InstCards = 0
  ElseIf OptPos = Opt_InfoCard Then
    If InfoCard = 1 Then InfoCard = 0 Else InfoCard = 1
  ElseIf OptPos = Opt_Dust Then
    If Dust = 1 Then Dust = 0 Else Dust = 1
  ElseIf OptPos = Opt_Volume Then
    MechVolume = MechVolume + amount * 0.1
    If MechVolume < 0 Then MechVolume = 1
    If MechVolume > 1 Then MechVolume = 0
  ElseIf OptPos = Opt_Volume_Ramp Then
    RampRollVolume = RampRollVolume + amount * 0.1
    If RampRollVolume < 0 Then RampRollVolume = 1
    If RampRollVolume > 1 Then RampRollVolume = 0
  ElseIf OptPos = Opt_Volume_Ball Then
    BallRollVolume = BallRollVolume + amount * 0.1
    If BallRollVolume < 0 Then BallRollVolume = 1
    If BallRollVolume > 1 Then BallRollVolume = 0
  End If
End Sub

Sub Options_KeyDown(ByVal keycode)
  If OptSelected Then
    If keycode = LeftMagnaSave Then ' Exit / Cancel
      OptSelected = False
    ElseIf keycode = RightMagnaSave Then ' Enter / Select
      OptSelected = False
    ElseIf keycode = LeftFlipperKey Then ' Next / +
      Options_Toggle  -1
    ElseIf keycode = RightFlipperKey Then ' Prev / -
      Options_Toggle  1
    End If
  Else
    If keycode = LeftMagnaSave Then ' Exit / Cancel
      Options_Close
    ElseIf keycode = RightMagnaSave Then ' Enter / Select
      If OptPos < Opt_Info_1 Then OptSelected = True
    ElseIf keycode = LeftFlipperKey Then ' Next / +
      OptPos = OptPos - 1
      If OptPos < 0 Then OptPos = NOptions - 1
    ElseIf keycode = RightFlipperKey Then ' Prev / -
      OptPos = OptPos + 1
      If OptPos >= NOPtions Then OptPos = 0
    End If
  End If
  Options_OnOptChg
End Sub

Sub Options_Load
  Dim x
  x = LoadValue(cGameName, "LIGHT") : If x <> "" Then LightLevel = CInt(x) Else LightLevel = 50
  x = LoadValue(cGameName, "UGICOLOR") : If x <> "" Then UpperGIColor = CInt(x) Else UpperGIColor = 1
  x = LoadValue(cGameName, "MGICOLOR") : If x <> "" Then MainGIColor = CInt(x) Else MainGIColor = 1
  x = LoadValue(cGameName, "LGICOLOR") : If x <> "" Then LowerGIColor = CInt(x) Else LowerGIColor = 1
  x = LoadValue(cGameName, "BLCOLOR") : If x <> "" Then BumperLightColor = CInt(x) Else BumperLightColor = 1
  x = LoadValue(cGameName, "CCPOSTS") : If x <> "" Then CCPosts = CInt(x) Else CCPosts = 0
  x = LoadValue(cGameName, "BALLMOD") : If x <> "" Then BallMod = CInt(x) Else BallMod = 0
  x = LoadValue(cGameName, "LAZYFLIPS") : If x <> "" Then FlipperKeyMod = CInt(x) Else FlipperKeyMod = 1
  x = LoadValue(cGameName, "INSTCARDS") : If x <> "" Then InstCards = CInt(x) Else InstCards = 0
  x = LoadValue(cGameName, "INFOCARD") : If x <> "" Then InfoCard = CInt(x) Else InfoCard = 1
  x = LoadValue(cGameName, "DUST") : If x <> "" Then Dust = CInt(x) Else Dust = 1
  x = LoadValue(cGameName, "VOLUME") : If x <> "" Then MechVolume = CNCDbl(x) Else MechVolume = 0.8
  x = LoadValue(cGameName, "RAMPVOLUME") : If x <> "" Then RampRollVolume = CNCDbl(x) Else RampRollVolume = 0.5
  x = LoadValue(cGameName, "BALLVOLUME") : If x <> "" Then BallRollVolume = CNCDbl(x) Else BallRollVolume = 0.5
  x = LoadValue(cGameName, "SETDIPS") : If x <> "" Then InitDips = CInt(x) Else InitDips = 1
  UpdateMods
End Sub


Sub UpdateMods
  Dim BP

  'One time initialization of DIPs (requres restart)
  If InitDips = 1 Then
    SetDefaultDips
    InitDips = 0
    SaveValue cGameName, "SETDIPS", InitDips
  End If

  ' Room Brightness
  SetRoomBrightness LightLevel/100

  ' Upper GI Color
  If UpperGIColor = 0 Then
    SetUpperGiColor cArray(RndInt(0,5))
  ElseIf UpperGIColor = 7 Then
    SetUpperGiColorSpooky
  Else
    SetUpperGiColor cArray(UpperGIColor-1)
  End If

  ' Main GI Color
  If MainGIColor = 0 Then
    SetMainGiColor cArray(RndInt(0,5))
  Else
    SetMainGiColor cArray(MainGIColor-1)
  End If

  ' Lower GI Color
  If LowerGIColor = 0 Then
    SetLowerGiColor cArray(RndInt(0,5))
  Else
    SetLowerGiColor cArray(LowerGIColor-1)
  End If

  ' Bumper Light Color
  If BumperLightColor = 0 Then
    SetBumperColor cArray(RndInt(0,5))
  Else
    SetBumperColor cArray(BumperLightColor-1)
  End If


  ' BallMod's
  If BallMod = 0 Then
    HHBall.FrontDecal = "balldecal_ScratchesWeathered2"
    HHBall.image = "ball_Test8B"
    HHBall.visible = True
    HHBall.bulbintensityscale = 1
    pPinball1.visible = False
    MoveGlowBalls.enabled = False
  ElseIf BallMod = 1 Then
    HHBall.FrontDecal = "marbledball RedGreen"
    HHBall.image = "VR_Black"
    HHBall.visible = True
    HHBall.bulbintensityscale = 0
    pPinball1.visible = False
    MoveGlowBalls.enabled = False
  ElseIf BallMod = 2 Then
    HHBall.FrontDecal = "marbledball YellowGreen"
    HHBall.image = "VR_Black"
    HHBall.visible = True
    HHBall.bulbintensityscale = 0
    pPinball1.visible = False
    MoveGlowBalls.enabled = False
  Else 'Glowball
    'CCGI2.Enabled = true
    MoveGlowBalls.enabled = True
    HHBall.visible = False
    pPinball1.visible = True
    pPinball1.material = "CCBallMaterial1"
    Select Case BallMod
      Case 3: pPinball1.material = "AcrylicBLRed"
      Case 4: pPinball1.material = "AcrylicBLOrange"
      Case 5: pPinball1.material = "AcrylicBLYellow"
      Case 6: pPinball1.material = "AcrylicBLGreen1"
      Case 7: pPinball1.material = "AcrylicBLGreen2"
      Case 8: pPinball1.material = "AcrylicBLBlue"
      Case 9: pPinball1.material = "AcrylicBLPink"
      Case 10: pPinball1.material = "AcrylicBLPurple"
    End Select
  End If


  ' Instruction Cards
  If InstCards = 1 Then
    BM_ICLeft.visible = False
    BM_ICRight.visible = False
    pLeftInstuctionCard.visible = True
    pRightInstuctionCard.visible = True
    pLeftInstuctionCard.Image = "HHICLeftC2"
    pRightInstuctionCard.Image = "HHICRightC2"
  ElseIf InstCards = 2 Then
    BM_ICLeft.visible = False
    BM_ICRight.visible = True
    pLeftInstuctionCard.visible = True
    pRightInstuctionCard.visible = False
    pLeftInstuctionCard.Image = "HHICLeft3ball"
  ElseIf InstCards = 3 Then
    BM_ICLeft.visible = False
    BM_ICRight.visible = True
    pLeftInstuctionCard.visible = True
    pRightInstuctionCard.visible = False
    pLeftInstuctionCard.Image = "HHICLeft5ball"
  Else
    BM_ICLeft.visible = True
    BM_ICRight.visible = True
    pLeftInstuctionCard.visible = False
    pRightInstuctionCard.visible = False
  End If


  ' Info Cards
  If InfoCard = 1 Then
    BM_pInfoCard.visible = False
    pInfoCard.visible = True
  Else
    BM_pInfoCard.visible = False
    pInfoCard.visible = False
  End If

  ' Playfield Dust
  If Dust = 1 Then
    DustFlasher.visible = True
    For each BP in BP_dust: BP.visible = False: Next
  Else
    DustFlasher.visible = False
    For each BP in BP_dust: BP.visible = False: Next
  End If


  ' Side rails (always for VR, and if not in cabinet mode)
  If Not DesktopMode And VR_Room = 0 Then 'Cab Mode
    For each BP in BP_SideRails: BP.visible = False: Next
    For each BP in BP_lockDownBar: BP.visible = False: Next
  Else
    For each BP in BP_SideRails: BP.visible = True: Next
    For each BP in BP_lockDownBar: BP.visible = True: Next
  End If

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



'*****************************************************************************************************
'  ZVRR: VR ROOM STUFF
'*****************************************************************************************************


'*****************************************************************************************************
' VR CABINET ANIMATIONS
'*****************************************************************************************************

Sub VR_Plunger_Timer
  If VR_Primary_plunger.transy > -135 then
      VR_Primary_plunger.transy = VR_Primary_plunger.transy - 5
    'VR_Primary_plunger.transz = VR_Primary_plunger.transy * dSin(6.5)
  End If
End Sub

Sub VR_Plunger2_Timer
  VR_Primary_plunger.transy = (-5* Plunger.Position)
  'VR_Primary_plunger.transz = VR_Primary_plunger.transy * dSin(6.5)
End Sub

Sub VR_Keydown(keycode)
  If keycode = LeftFlipperKey Then
    VR_CabFlipperLeft.transx = 10 ' Animate VR Left flipperbutton
  End If
  If keycode = RightFlipperKey Then
    VR_CabFlipperRight.transx = -10 ' Animate VR Right flipperbutton
  End If

  If keycode = LeftMagnaSave Then
    VR_CabMagnaLeft.transx = 10 ' Animate VR Left flipperbutton
  End If
  If keycode = RightMagnaSave Then
    VR_CabMagnaRight.transx = -10 ' Animate VR Right flipperbutton
  End If
  If keycode = PlungerKey Then
    VR_Plunger.Enabled = True
    VR_Plunger2.Enabled = False
  End If
  If Keycode = StartGameKey Then
    VR_Cab_StartButton.transx = -5    ' Animate VR Startbutton
  End If
End Sub

Sub VR_Keyup(keycode)
  If keycode = LeftFlipperKey Then
    VR_CabFlipperLeft.transx = 0  ' Animate VR Left flipperbutton
  End If
  If keycode = RightFlipperKey Then
    VR_CabFlipperRight.transx = 0 ' Animate VR Right flipperbutton
  End If

  If keycode = LeftMagnaSave Then
    VR_CabMagnaLeft.transx = 0  ' Animate VR Left flipperbutton
  End If
  If keycode = RightMagnaSave Then
    VR_CabMagnaRight.transx = 0 ' Animate VR Right flipperbutton
  End If
  If keycode = PlungerKey Then
    VR_Plunger.Enabled = False
    VR_Plunger2.Enabled = True
  End If
  If Keycode = StartGameKey Then
    VR_Cab_StartButton.transx = 0   ' Animate VR Startbutton
  End If
End Sub

'*****************************************************************************************************
' VR ROOM
'*****************************************************************************************************

Dim VRThings, vrx, VRLShadows(30), VRLBDL(60), VRLSCount, VRLBDLCount

if VR_Room = 1 Then
  for each VRThings in VRBackglass:VRThings.visible = 1:Next
  for each VRThings in VRCabinet:VRThings.visible = 1:Next
  for each VRThings in VRRoom:VRThings.visible = 1:Next
  vrx = 0
  For Each VRThings in VRRoomShadows
    Set VRLShadows(vrx) = VRThings
    vrx = vrx + 1
  Next
  VRLSCount = vrx

  vrx = 0
  For Each VRThings in VRRoomBDL
    Set VRLBDL(vrx) = VRThings
    vrx = vrx + 1
  Next
  VRLBDLCount = vrx
  SkyLightTimer1.enabled = true
  VRTimer.enabled = true
  center_digits
Else
  If ShowBGinDesktopMode Then
    center_digits
    for each VRThings in VRBackglass:VRThings.visible = 1:Next
  Else
    for each VRThings in VRBackglass:VRThings.visible = 0:Next
  End If
  for each VRThings in VRCabinet:VRThings.visible = 0:Next
  for each VRThings in VRRoom:VRThings.visible = 0:Next
  SkyLightTimer1.enabled = false
  SkyLightTimer2.enabled = false
  SkyLightTimer3.enabled = false
  VRTimer.enabled = false
End If

' *******************     LOL - They aren't in Collections... :) *****************************************************
Sub LightningOn()
  For vrx = 0 to VRLSCount - 1
    VRLShadows(vrx).visible = True
  Next
  For vrx = 0 to VRLBDLCount - 1
    VRLBDL(vrx).Blenddisablelighting = 0.05
  Next
End Sub

Sub LightningOff()
  For vrx = 0 to VRLSCount - 1
    VRLShadows(vrx).visible = False
  Next
  For vrx = 0 to VRLBDLCount - 1
    VRLBDL(vrx).Blenddisablelighting = 0
  Next
End Sub

Sub SkyLightTimer1_timer()
  SkyLightTimer2.interval = 40 + rnd(1)*120  ' random between 40 and 160ms
  SkyLightTimer3.interval = 2000 + rnd(1)*1000  ' random between 2 and 3 seconds
  VR_Sphere1.visible = false
  VR_Sphere2.visible = true
  LightningOn
  SkyLightTimer1.enabled = false
  SkyLightTimer2.enabled = true
  SkyLightTimer3.enabled = true
End Sub

Sub SkyLightTimer2_timer()
  SkyLightTimer1.interval = 10000 + rnd(1)*15000  ' random between 10 and 25 seconds
  VR_Sphere1.visible = true
  VR_Sphere2.visible = false
  LightningOff
  SkyLightTimer1.enabled = true
  SkyLightTimer2.enabled = false
End Sub

Sub SkyLightTimer3_timer()
  select case Int(rnd*4):
    Case 0: PlaySoundAt "thunder8", VR_TreesGroup
    Case 1: PlaySoundAt "thunder9", VR_TreesGroup
    Case 2: PlaySoundAt "thunder3", VR_TreesGroup
    Case 3: PlaySoundAt "thunder6", VR_TreesGroup
  End Select
  SkyLightTimer3.enabled = false
End Sub

Dim GhostDir, GhostSpeed, llcount, bgull, bgulm, bgulr, bgulb, bgurr, bgurm, bgurl
GhostSpeed = 15

Sub VRTimer_timer()
  If VR_Room = 1 Then
    VR_Clouds.ObjRotZ=VR_Clouds.ObjRotZ+0.01
    VR_Clouds2.ObjRotZ=VR_Clouds2.ObjRotZ-0.01
  End If

  If VR_Ghost.transy = 1 Then
    VR_Ghost.transy = 0
    VR_Ghost.x = -40000
    GhostDir = 1
  End If

  If VR_Ghost.x > 40000 Then
    GhostDir = -1
    VR_Ghost.Roty = 90
  End If
  If VR_Ghost.x < -40000 Then
    GhostDir=1
    VR_Ghost.Roty = -90
  End If

  VR_Ghost.x = VR_Ghost.x + GhostDir*GhostSpeed


  'Backglass
  If GameInPlay or controller.Lamp(1) Then
    BG_Rampdown VR_BG_GO, 100, 0
    BG_Rampdown VR_BG_Match, 100, 0
    BG_Rampup VR_BG_BIP, 100, 0
  Else
    BG_Rampup VR_BG_GO, 100, 0
    BG_Rampup VR_BG_Match, 100, 0
    BG_Rampdown VR_BG_BIP, 100, 0
  End If

  If controller.Lamp(1) = true Then
    BG_Rampup VR_BG_Tilt, 100, 0
  Else
    BG_Rampdown VR_BG_Tilt, 100, 0
  End if

  If controller.Lamp(3) = true Then
    BG_Rampup VR_BG_SA, 100, 0
  Else
    BG_Rampdown VR_BG_SA, 100, 0
  End if

  If controller.Lamp(10) = true Then
    BG_Rampup VR_BG_HSTD, 100, 0
  Else
    BG_Rampdown VR_BG_HSTD, 100, 0
  End if

  Select Case llcount
    Case 0: bgull=1
    Case 11: bgull=0:bgulm=1
    Case 22: bgulm=0:bgulr=1
    Case 33: bgulr=0
    Case 37: bgurr=1
    Case 48: bgurr=0:bgurm=1
    Case 59: bgurm=0:bgurl=1
    Case 70: bgurl=0
    Case 73: bgull=1
    Case 84: bgull=0:bgulb=1
    Case 95: bgulb=0
    Case 140: llcount = -1
  End Select

  llcount = llcount + 1

  If bgull = 1 then BG_Rampup VR_BG_ULL, 100,  0 else BG_Rampdown VR_BG_ULL, 100, 0 end if
  If bgulm = 1 then BG_Rampup VR_BG_ULM, 100,  0 else BG_Rampdown VR_BG_ULM, 100, 0 end if
  If bgulr = 1 then BG_Rampup VR_BG_ULR, 100,  0 else BG_Rampdown VR_BG_ULR, 100, 0 end if
  If bgulb = 1 then BG_Rampup VR_BG_ULB, 100,  0 else BG_Rampdown VR_BG_ULB, 100, 0 end if
  If bgurl = 1 then BG_Rampup VR_BG_URL, 100,  0 else BG_Rampdown VR_BG_URL, 100, 0 end if
  If bgurm = 1 then BG_Rampup VR_BG_URM, 100,  0 else BG_Rampdown VR_BG_URM, 100, 0 end if
  If bgurr = 1 then BG_Rampup VR_BG_URR, 100,  0 else BG_Rampdown VR_BG_URR, 100, 0 end if
End Sub

Sub BG_Rampup(obj, max, min)
  If obj.opacity < max Then
    obj.opacity = obj.opacity + (max-min)/3
  End If
  If obj.opacity > max Then obj.opacity = max
End Sub

Sub BG_Rampdown(obj, max, min)
  If obj.opacity > min Then
    obj.opacity = obj.opacity - (max-min)/4
  End If
  If obj.opacity < min Then obj.opacity = min
End Sub

'*****************************************************************************************************
' VR BACKGLASS
'*****************************************************************************************************

Sub center_digits()
  Dim xoff,yoff,zoff,xrot

  xoff =480
  yoff =15
  zoff =238
  xrot = -85

  Dim ii,xx,yy,yfact,xfact,obj,xcen,ycen,zscale

  zscale = 0.0000001

  xcen =(1032 /2) - (74 / 2)
  ycen = (1020 /2 ) + (194 /2)

  yfact =0 'y fudge factor (ycen was wrong so fix)
  xfact =0

  If RomSet = 1 Then
    for ii = 0 to 27
      For Each obj In Digits(ii)
      xx = obj.x

      obj.x = (xoff - xcen) + xx + xfact
      yy = obj.y ' get the yoffset before it is changed
      obj.y =yoff

      If(yy < 0.) then
        yy = yy * -1
      end if

      obj.height =( zoff - ycen) + yy - (yy * zscale) + yfact

      obj.rotx = xrot
      Next
    Next
  Elseif Romset = 2 Then
    for ii = 0 to 31
      For Each obj In Digits7(ii)
      xx = obj.x

      obj.x = (xoff - xcen) + xx + xfact
      yy = obj.y ' get the yoffset before it is changed
      obj.y =yoff

      If(yy < 0.) then
        yy = yy * -1
      end if

      obj.height =( zoff - ycen) + yy - (yy * zscale) + yfact

      obj.rotx = xrot
      Next
    Next
  End If

  For each obj in VRBackglass
    obj.rotx = 175
    obj.x = 2264.278
    obj.y = 1008.072
    obj.z = -575
  Next
  VR_BG_Dark.y = VR_BG_Dark.y - 0.1
end sub



'*****************************************
' ZCHL: CHANGELOG
'*****************************************
'
' 1.1.2 - cyberpez
'   Initial table update based on v1.0
' 1.1.3 - apophis
'   Lots of script clean up and rework
'   Table object clean up and layer organization
'   Updated all physics materials and scripts, added flipper corrections, added rubber dampening
'   Added Roth stand-up and drop target objects and code
'   Updated and simplified trough code
'   Reworked insert lights in prep for toolkit. Deleted old code related to lighting
'   GI light cleanup in prep for toolkit
'   Add Fleep mechanical sound effects. Deleted old sound code and effects
'   Add ball ambient shadows, flipper shadows
'   Updated ball rolling sound effects
'   Added ramp rolling sound effects
'   Set up some animation code in prep for toolkit
'   Timer and collections clean up
'   Updated ball image
' 1.1.4 - rothbauerw
'   Added new code for kicking targets
'   Fixed ball shadow height
'   Hooked up flippers to the correct key codes for the option to split between flipper buttons and magna's. Also set UseSolenoid = 2.
'   Add drop sounds to the cellar
'   Adjusted physics around the right kick back so it can enter and exit cleanly
'   Added plastics and other physics objects to prevent stuck balls
'   Removed kicker code from ball rolling sub, it's no longer needed
'   reactivated the animation code for the secret target
' 1.1.4b - cyberpez
'   Initial 2k bake import
'   Fixed some physics walls
' 1.1.5 - apophis
'   Added VLM materials. Made all collidables invisible. Set all bakemaps to non-static.
'   Hid all lights. Added bumper lights. Fixed some light names. Enabled ray traced shadows.
'   Added room light level and initial ball brightness code
'   Added initial GI color mods
'   Animated: switches, gates, bumpers, flippers
' 1.1.5b - cyberpez
'   Second Batch (2k)
'   GlowBalls and ColotChangingPosts
' 1.1.5d - rothbauerw
'   Tuned the LPF drop sounds, adjusted all the VP objects on the lower playfield for the altered LPF ramp,
'   and adjust both kickers to align with the lower playfield and allow the ball to drop into the kickers.
'   Also modified the secret target code for more accurate animation and soundFX.
'   Ajusted LPF flippers so they are rotate correctly.
' 1.1.5e - cyberpez
'   Third Batch (2k)
' 1.1.6 - apophis
'   Set up two-position flipper animations. Fixed lower flippers transformations
'   Fixed some issues with GI color options
'   Set all Flashers lightmaps to be controlled by giLL1 (should be done in blender)
'   Set "hide parts behind" for all playfield bakemaps
'   Cleaned up image manager
'   Animated: Standup, kicking, drop, and secret targets. Trapdoor. Rubbers. Slingshots.
' 1.1.6a - cyberpez
'   Fourth Batch (2k)
' 1.1.6b - apophis
'   Fifth Batch (2k)
' 1.1.7 - apophis
'   Re-fixes: Set "hide parts behind" for all playfield bakemaps. Lower PF flipper orientations.
'       Fixed upper pf inserts (L2 was wrong). Swapped L3 and L8 positions on main pf. (fixed in blender too)
'   Hid ball shadow when under the upper pf. Adjust ball brightness with gi state when on lower pf.
'   Added Flex option menu. Updated GI color options.
' 1.1.8 - tomate
'   some fixes at blender side, new 2k batch added
' 1.1.9 - apophis
'   Re-fixes: Set "hide parts behind" for all playfield bakemaps. Lower PF flipper orientations.
'   Added reflection probes to Main playfield. Fixed flipper shadows.
'   Added giML8 light to GI Main layer and collection
'   Fixed trap door ramp opening angle and visibility issues
'   Made "LazyMan" flippers option default to OFF
' 1.2.0 - tomate
'   Re-fixes: Set "hide parts behind" for all playfield bakemaps.
'   Some fixes at blender side, New 4k batch added
' 1.2.1 - apophis
'   Re-fixes: Lower PF flipper orientations. Added reflection probes to Main playfield.
'   Added Info Card and Instruction Card options to menu (pInfoCard, pICLeft, pICRight)
'   Animated scoring springs (pSSpring, pSSpring2). Added TWKicker1 animation. Added missing gate animations
' 1.2.2 - apophis
'   Set sw31wall top height to -410 and sw41wall top height to -225. Set Const EOSTnew = 1.5 (thanks rothbauerw)
'   Added missing sw21 animation. Added missing animated rubber animation.
'   Added drop target shadows animations on upper playfield DTs
'   Set PlayfieldReflectionScale=0 for marble ball mods
'   Delayed SecretTarget bounce back after ball falls thru
'   Made ball shadows not show up if glowball being used
'   Fixed lower PF DT animations by adding 11 deg to their RotX values.
' 1.2.4 - tomate
'   Re-fixes: Set "hide parts behind" for all playfield bakemaps.
'   Baked inserts, new screws and nuts, lower playfield collection separated, some other fixes, New 4k batch added
' 1.2.5 - rothbauerw
'   Added VR Backglass and VR Scoring
' 1.2.6 - apophis
'   Re-fixes: Lower PF object orientations. Added reflection probes to Main playfield.
'   Assinged lower pf objects with new materials. Brightness adjusted in material manager.
'   Fixed DB issues with underPFs and Layers
'   Updated flipper option text
'   Fixed ambient ball shadows
'   Added automatic cabinet mode (hiding rails and lockdown bar)
'   Hooked up VR room brightness with LightLevel option.
' 1.2.7 - Sixtoe
'   Added dust to main pf window
' 1.2.8 - apophis
'   Fixed upper PF inserts in front of standups.
'     Set dust opacity to 200 and made it optional.
'   Made option menu not open when ball is in plunger lane.
'   Not adjusting VR_Clouds and VR_Ghost materials with light level anymore. Set their base colors to 50% of previous value.
'       Fixed playfield reflections
'   Made backglass reflect on playfield in desktop mode. Do we want this?
' 1.2.9 - rothbauerw
'   Stuck ball and lost ball fixes. No standalone fix yet.
' 1.3.0 - rothbauerw
'   Secret target rotation updates.
' 1.3.1 - rothbauerw
'   Implemented the CorTracker fix from jsm.
'   Added more robust secretTarget animation handling
'   Updated VR backglass lightning timing. Found a good video that showed clearly the pattern and timing. It's now much more accurate.
'   The secret target may need some minor update once the new pivot point is in, but I tried to account for that already.
' 1.3.2 - apophis
'   Added "spooky" upper GI color option
'   Updated glowball option to include all colors
'   Added const ShowBGinDesktopMode for those who really want to see the backglass reflections in desktop mode
'   Added DOF commands for magnasave keys (E301, E302)
'   Added SetDefaultDips for one-time initialization of DIPs (requres restart)
' 1.3.3 - rothbauerw
'   Fixed a material issue in the VR Room, not sure where that cropped up.
' 1.3.4 - rothbauerw
'   Added UPF ceiling and adjusted the LPF saucer kickout
' 1.3.5 - tomate
'   New 4k bake that includes updated GI, insert cups, main pf plastics and left scoop height adjustments,
'       upper pf light names assignment fixes, clear tube put on new layer, dust layer, and more.
' RC1 - apophis
'   Re-fixes: Hide parts behind for all playfields. Lower PF object orientations and material assignments.
'     Added reflections to upper playfield. Fixed some DB issues. Disabled DisplayTimer6 in the editor.
' RC2 - apophis
'   Fixed zfighting on the lower pf lightmaps
'   Ramp002 and secret target code tweaks (thanks rothbauerw)
' RC3 - rothbauerw
'   Final fix for the secret target
' RC4 - apophis
'   Fixed red glowball material (thanks cyberpez)
'   Fixed marble ball image.
' RC5 - apophis
'   Updated sw31 kicker per rothbauerw & nFozzy suggestion
'   Fixed ball bulbintensityscale issue
'   Enabled reflections of Parts
' RC6 - apophis
'   L2 lightmaps linked to upper pf GI
'   Fixed magna key DOF logic
' RC7 - rothbauerw
'   Tightened tolerances on the kickback targets
' RC8 - apophis
'   Made lower PF bakemaps slightly darker
'     Fixed bumper skirt animations
'   Fixed texture on right rail (thanks tomate)
' Release 2.0
' 2.0.1 - rothbauerw
'   Updated left scoop kick randomness. This also fixes standalone issue.
'   Fixed double scoring issue in lower PF scoop
' 2.0.2 - apophis
'   Updated ball brightness in LFP when GI is off
'   Fixed issue with the backglass object reflections. Had render probes wrongly assigned to them
'   Utilizing new DisableStaticPrerendering when options menu is active (requires VPX build 1666 or higher)
' 2.0.3 - apophis
'   Removed wrongly assigned render probe to VR Cab primitives. (thanks rothbauerw)
' 2.0.4 - apophis
'   Removed the customization in the Roth DT code that handled lower playfield. Instead, added 11 deg to objrotx for all LPF DTs
'   Made drop targets drop down a little more.
' 2.0.5 - apophis
'   Updated DT code to account for collision edge case (rothbauerw).
' Release 2.1
' 2.1.1 - apophis
'     Fixed timers causing stutters.
'
' ***********************************************************************************************
' ***********************************************************************************************


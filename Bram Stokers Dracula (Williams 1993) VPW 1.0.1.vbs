'
'                     +@
'             ...:::..#@@:                                                                 .
'        +@@@@@@@@@@@@@@@@@*                                                              .@#
'     -++-.       .+@@@@@@@@@*.                                                           @@#
'   =%.             +@@@@@@@@@@:                                                          @@@+
'                   %@@@@@@#:@@@@.                                                       :@@@@%:
'                   %@@@@@@:  :@@.                                                        @@@@@
'                   #@@@@@@.   =@@@@@+      --.                       .%@@@@@@=           -@@@*.
'                   .@@@@@@     .@@@@@=    -@@@.   -=                :@@@@@@@@:     *@@*   @@@*      +@@
'                    @@@@@@      *@@@@@-   =@@@+#@@@%.     :  %-    :@@@@@@@@:  :#. @@@@.  @@@+     =@@@%
'                    %@@@@@      +@@@@@@:-+@@@@@#@@@@.   =@@@@@@@@- *@@@@@@@@   %@# @@@@. .@@@=   :@++@@@
'                    +@@@@@      :@@@@@@@-   =@@@@@@@: :@@@. @@@:   @@@@@@@-@  .%@. @@@+  :@@@- .*@#.@@@=
'                    :@@@@@     =@@@@@@*     :@@@.=@@@-@@@* *@@@-  :@@@@*=:.@   #@ -@@@=  :@@@:.@@@*-@@@
'                    .@@@@@      %@@@@@:     -@@@:     @@@ :@@@@.  :@@@=   :@   -@+@+@@.  :@@@..@@@+=@@+
'                     @@%@%     .@@@@@-      :@@@      *@# #@@@%   -@@@.        .@@@:@@   *@@@ .@@@+@@@:
'                     @@#@#    .@@@@@=       .@@@       @@.@@@@*    @@@.         @@*.@:   -@@@ .@@@%@@#
'                     @@-@=    *@@@@.         #@*       :@@@+ %+    #@@:             @.   .@@*  *@@@@@-
'             .-*#%+. +@#@.  -@@@@@...        :@=       .@@+ .@@.   :@@=             #.    @@=  -@@@:%:
'           .@@-      +@@@   @@@@.+@@-        :@-        #+  .@@-    :@%             *     @@-   ... =-
'        :%@+.        +@=@ =@@@:    ::       *@@@             :@      =@*                  #@:        =
'     +@@@+           +@+@%@@.                =@@             .%.      :@%.                +@.
'   :@@@@.            #@@@@-                  .@.                        .@@=.             =%
'  .%@@+             .@@*@                     +                            .-****@:       =*
'  .@@@:          :*@++*:#                                                                 -=
'    .%@-.....-*=.    +=                                                                   ..
'        ...          =+
'                     +@
'
'
'                                           Bram Stoker's Dracula
'                                   IPD No. 3072 / April, 1993 / 4 Players
'
' VPW TEAM
' ----------
' FrankEnstein - 3D modeling and rendering. Blood sucking.
' Apophis - Scripting, physics, VR rooms and associated trickery. Blood bathing.
' Sixtoe - Table building and research. Bloody hell.
' Stavcas - Initial table build and coding. Blood donating.
' Schlabber34 - 3D models of the ramps. Blood chilling.
' Kevv - 3D scans of the sculpts, numerous measurements and photos of real machine. Bleeder.
' DGrimmReaper - Initial VR minimal room work. Blood letting.
' HauntFreaks - Backglass image and supporting b2s. Blood tasting.
'
' Special thanks to Landers and Dorsola for their Mist magnet script (it was used with minor modifications) and to ClarkKent for the playfield scan.
' Thanks to all the previous authors that worked on this title: JPSalas, ICPjuggla, Dozer316, Thalamus
'
'
' REQUIRED: VPX 10.8 RC4 or later
'           VPinMame 3.6.0.998 or later
'




'//////////////////////////////////////////////////////////////////////////////////////////////////////
' THIS TABLE INCLUDES THE IN-GAME OPTIONS MENU
' To open the menu, press F12
'//////////////////////////////////////////////////////////////////////////////////////////////////////



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
' ZKEY: Key Press Handling
'   ZOPT: Table Options
'   ZBBR: Ball Brightness
'   ZBRI: Room Brightness
' ZGIU: GI Updates
' ZGIC: GI Colors
' ZSOL: Solenoids
' ZFLP: Flippers
' ZSLG: Slingshots
' ZBMP: Bumpers
' ZSWI: Switches
' ZINL: Inlane speed limit code
' ZDRN: Drain, Trough, and Ball Release
' ZVUK: VUKs
' ZMST: Mist Magnet
'   ZMAT: General Math Functions
' ZFTR:  Flipper Tricks
'   ZNFF: Flipper Corrections
'   ZDMP: Rubber Dampeners
' ZSSC: Slingshot Corrections
'   ZBOU: VPW TargetBouncer
'   ZRRL: Ramp Rolling Sound effects
' ZSFX: Mechanical Sound effects
' ZBRL: Ball Rolling and Drop Sounds
'   ZABS: Ambient ball shadows
' ZSTA: Stand-up Targets
' ZDTA: Drop Targets
'   ZANI: Misc Animations
' ZVRR: VR Room Stuff
' ZCHL: Change Log
'
'************************************************************************************************



'*******************************************
' ZCON: Constants and Global Variables
'*******************************************

Option Explicit
Randomize

' ROM
Const cGameName = "drac_l1"
'Const cGameName = "drac_l2c"  'L-2C Competition MOD

' Standard definitions
Const UseVPMModSol = 2
Const UseSolenoids = 2
Const UseLamps = 1
Const UseSync = 0
Const HandleMech = 0

' Standard Balls
Const BallSize = 50
Const BallMass = 1
Const tnob = 4            'Total number of balls the table can hold
Const lob = 0           'Locked balls

' Standard Sounds
Const SSolenoidOn = ""
Const SSolenoidOff = ""
Const SFlipperOn = ""
Const SFlipperOff = ""
Const SCoin = ""

' Table stuff
Dim DesktopMode:DesktopMode = Table1.ShowDT
dim tablewidth: tablewidth = Table1.width
dim tableheight: tableheight = Table1.height

' Table variables
Dim VarHidden, UseVPMDMD
Dim mMagnet, dtLDrop, x
Dim BSDBall1, BSDBall2, BSDBall3, BSDBall4, gBOT

Const TestVRDT = False


'*******************************************
' ZVLM: VLM Arrays
'*******************************************



' VLM  Arrays - Start
' Arrays per baked part
Dim BP_BmpsBtm: BP_BmpsBtm=Array(BM_BmpsBtm, LM_FL_F18_BmpsBtm, LM_FL_F19a_BmpsBtm, LM_FL_F20_Dome_BmpsBtm, LM_FL_F22_BmpsBtm, LM_FL_F23_BmpsBtm, LM_L_L68_BmpsBtm, LM_L_Bmp_BmpsBtm, LM_GI_gim04_BmpsBtm, LM_GI_gim05_BmpsBtm, LM_GI_git01_BmpsBtm, LM_L_l84_BmpsBtm)
Dim BP_BmpsTop: BP_BmpsTop=Array(BM_BmpsTop, LM_FL_F18_BmpsTop, LM_FL_F19_BmpsTop, LM_FL_F19a_BmpsTop, LM_FL_F20_BmpsTop, LM_FL_F20_Dome_BmpsTop, LM_FL_F21_BmpsTop, LM_FL_F22_BmpsTop, LM_FL_F23_BmpsTop, LM_FL_F24_BmpsTop, LM_L_L68_BmpsTop, LM_L_Bmp_BmpsTop, LM_GI_gim01_BmpsTop, LM_GI_gim04_BmpsTop, LM_GI_gim05_BmpsTop, LM_GI_git01_BmpsTop, LM_L_l25_BmpsTop, LM_L_l28_BmpsTop, LM_L_l35_BmpsTop, LM_L_l83_BmpsTop, LM_L_l84_BmpsTop, LM_L_l85_BmpsTop)
Dim BP_Bumper1_Ring: BP_Bumper1_Ring=Array(BM_Bumper1_Ring, LM_FL_F18_Bumper1_Ring, LM_FL_F19_Bumper1_Ring, LM_FL_F19a_Bumper1_Ring, LM_FL_F21_Bumper1_Ring, LM_FL_F21a_Bumper1_Ring, LM_FL_F22_Bumper1_Ring, LM_FL_F23_Bumper1_Ring, LM_FL_F24_Bumper1_Ring, LM_L_Bmp_Bumper1_Ring, LM_GI_gim01_Bumper1_Ring, LM_GI_gim04_Bumper1_Ring, LM_GI_gim05_Bumper1_Ring, LM_L_l85_Bumper1_Ring)
Dim BP_Bumper1_Socket: BP_Bumper1_Socket=Array(BM_Bumper1_Socket, LM_FL_F18_Bumper1_Socket, LM_FL_F19a_Bumper1_Socket, LM_FL_F22_Bumper1_Socket, LM_FL_F23_Bumper1_Socket, LM_L_Bmp_Bumper1_Socket, LM_GI_gim04_Bumper1_Socket, LM_GI_gim05_Bumper1_Socket, LM_GI_git01_Bumper1_Socket)
Dim BP_Bumper2_Ring: BP_Bumper2_Ring=Array(BM_Bumper2_Ring, LM_FL_F18_Bumper2_Ring, LM_FL_F19a_Bumper2_Ring, LM_FL_F22_Bumper2_Ring, LM_FL_F23_Bumper2_Ring, LM_L_L68_Bumper2_Ring, LM_L_Bmp_Bumper2_Ring, LM_GI_gim04_Bumper2_Ring, LM_GI_git01_Bumper2_Ring, LM_L_l83_Bumper2_Ring, LM_L_l84_Bumper2_Ring)
Dim BP_Bumper2_Socket: BP_Bumper2_Socket=Array(BM_Bumper2_Socket, LM_FL_F18_Bumper2_Socket, LM_FL_F19a_Bumper2_Socket, LM_FL_F20_Bumper2_Socket, LM_FL_F22_Bumper2_Socket, LM_FL_F23_Bumper2_Socket, LM_L_L61_Bumper2_Socket, LM_L_Bmp_Bumper2_Socket, LM_GI_gim04_Bumper2_Socket, LM_GI_git01_Bumper2_Socket, LM_L_l58_Bumper2_Socket, LM_L_l84_Bumper2_Socket)
Dim BP_Bumper3_Ring: BP_Bumper3_Ring=Array(BM_Bumper3_Ring, LM_FL_F18_Bumper3_Ring, LM_FL_F19a_Bumper3_Ring, LM_FL_F22_Bumper3_Ring, LM_FL_F23_Bumper3_Ring, LM_L_L68_Bumper3_Ring, LM_L_Bmp_Bumper3_Ring, LM_GI_gim05_Bumper3_Ring, LM_GI_git01_Bumper3_Ring)
Dim BP_Bumper3_Socket: BP_Bumper3_Socket=Array(BM_Bumper3_Socket, LM_FL_F18_Bumper3_Socket, LM_FL_F19a_Bumper3_Socket, LM_FL_F22_Bumper3_Socket, LM_FL_F23_Bumper3_Socket, LM_L_Bmp_Bumper3_Socket, LM_GI_gim05_Bumper3_Socket, LM_GI_git01_Bumper3_Socket)
Dim BP_Diverter: BP_Diverter=Array(BM_Diverter, LM_FL_F17_Dome_Diverter, LM_GI_git01_Diverter, LM_L_l58_Diverter)
Dim BP_DiverterU: BP_DiverterU=Array(BM_DiverterU, LM_FL_F17_Dome_DiverterU, LM_L_L63_DiverterU, LM_GI_git01_DiverterU, LM_L_l58_DiverterU, LM_L_l64_DiverterU)
Dim BP_ExtPlasticScrew1: BP_ExtPlasticScrew1=Array(BM_ExtPlasticScrew1, LM_FL_F21a_ExtPlasticScrew1, LM_GI_gib02_ExtPlasticScrew1)
Dim BP_ExtPlasticScrew2: BP_ExtPlasticScrew2=Array(BM_ExtPlasticScrew2, LM_FL_F19_ExtPlasticScrew2, LM_GI_gib08_ExtPlasticScrew2)
Dim BP_ExtPlasticScrew3: BP_ExtPlasticScrew3=Array(BM_ExtPlasticScrew3, LM_FL_F19_ExtPlasticScrew3, LM_FL_F19a_ExtPlasticScrew3, LM_FL_F21_ExtPlasticScrew3, LM_FL_F21a_ExtPlasticScrew3, LM_GI_gib01_ExtPlasticScrew3)
Dim BP_ExtPlasticScrew4: BP_ExtPlasticScrew4=Array(BM_ExtPlasticScrew4, LM_FL_F21a_ExtPlasticScrew4, LM_GI_gim01_ExtPlasticScrew4)
Dim BP_ExtPlasticSpacer1: BP_ExtPlasticSpacer1=Array(BM_ExtPlasticSpacer1)
Dim BP_ExtPlasticSpacer2: BP_ExtPlasticSpacer2=Array(BM_ExtPlasticSpacer2, LM_GI_gib03_ExtPlasticSpacer2)
Dim BP_ExtPlasticSpacer3: BP_ExtPlasticSpacer3=Array(BM_ExtPlasticSpacer3, LM_GI_gib01_ExtPlasticSpacer3)
Dim BP_ExtPlasticSpacer4: BP_ExtPlasticSpacer4=Array(BM_ExtPlasticSpacer4, LM_GI_gib05_ExtPlasticSpacer4)
Dim BP_ExtraPlasticL: BP_ExtraPlasticL=Array(BM_ExtraPlasticL, LM_GI_gib01_ExtraPlasticL, LM_GI_gib02_ExtraPlasticL, LM_GI_gib03_ExtraPlasticL, LM_GI_gib04_ExtraPlasticL, LM_GI_gib08_ExtraPlasticL, LM_L_l38_ExtraPlasticL)
Dim BP_ExtraPlasticR: BP_ExtraPlasticR=Array(BM_ExtraPlasticR, LM_FL_F21a_ExtraPlasticR, LM_GI_gib03_ExtraPlasticR, LM_GI_gib05_ExtraPlasticR, LM_GI_gib06_ExtraPlasticR, LM_GI_gim01_ExtraPlasticR)
Dim BP_Gate4: BP_Gate4=Array(BM_Gate4)
Dim BP_Gate6: BP_Gate6=Array(BM_Gate6, LM_FL_F17_Dome_Gate6, LM_L_L63_Gate6, LM_GI_git01_Gate6)
Dim BP_LFlip: BP_LFlip=Array(BM_LFlip, LM_GI_gib02_LFlip, LM_GI_gib04_LFlip, LM_L_l55_LFlip)
Dim BP_LFlipU: BP_LFlipU=Array(BM_LFlipU, LM_GI_gib02_LFlipU, LM_GI_gib04_LFlipU, LM_L_l43_LFlipU, LM_L_l44_LFlipU)
Dim BP_LPostRubber_p1: BP_LPostRubber_p1=Array(BM_LPostRubber_p1, LM_FL_F22_LPostRubber_p1, LM_GI_gib01_LPostRubber_p1, LM_L_l38_LPostRubber_p1)
Dim BP_LPostRubber_p2: BP_LPostRubber_p2=Array(BM_LPostRubber_p2, LM_FL_F22_LPostRubber_p2, LM_GI_gib01_LPostRubber_p2, LM_L_l37_LPostRubber_p2, LM_L_l38_LPostRubber_p2)
Dim BP_LPostRubber_p3: BP_LPostRubber_p3=Array(BM_LPostRubber_p3, LM_FL_F22_LPostRubber_p3, LM_GI_gib01_LPostRubber_p3, LM_L_l37_LPostRubber_p3, LM_L_l38_LPostRubber_p3)
Dim BP_LPost_p3: BP_LPost_p3=Array(BM_LPost_p3, LM_FL_F21a_LPost_p3, LM_GI_gib01_LPost_p3, LM_GI_gim01_LPost_p3, LM_L_l37_LPost_p3)
Dim BP_LRampLever: BP_LRampLever=Array(BM_LRampLever, LM_FL_F21a_LRampLever)
Dim BP_LRamp_DP: BP_LRamp_DP=Array(BM_LRamp_DP)
Dim BP_LRamp_UP: BP_LRamp_UP=Array(BM_LRamp_UP, LM_FL_F21a_LRamp_UP)
Dim BP_LSling1: BP_LSling1=Array(BM_LSling1, LM_FL_F21_LSling1, LM_FL_F21a_LSling1, LM_GI_gib01_LSling1, LM_GI_gib02_LSling1)
Dim BP_LSling2: BP_LSling2=Array(BM_LSling2, LM_FL_F21a_LSling2, LM_GI_gib01_LSling2, LM_GI_gib02_LSling2)
Dim BP_LSlingArm: BP_LSlingArm=Array(BM_LSlingArm, LM_FL_F21a_LSlingArm, LM_GI_gib01_LSlingArm, LM_GI_gib02_LSlingArm, LM_GI_gib03_LSlingArm, LM_GI_gib04_LSlingArm)
Dim BP_LSling_001: BP_LSling_001=Array(BM_LSling_001, LM_FL_F19a_LSling_001, LM_FL_F21_LSling_001, LM_FL_F21a_LSling_001, LM_GI_gib01_LSling_001, LM_GI_gib02_LSling_001, LM_GI_gib03_LSling_001, LM_GI_gib04_LSling_001, LM_L_l38_LSling_001)
Dim BP_Layer0: BP_Layer0=Array(BM_Layer0, LM_FL_F19a_Layer0, LM_FL_F21_Layer0, LM_FL_F21a_Layer0, LM_FL_F22_Layer0, LM_FL_F23_Layer0, LM_L_Bmp_Layer0, LM_GI_gim01_Layer0, LM_GI_git01_Layer0, LM_L_l16_Layer0, LM_L_l18_Layer0, LM_L_l24_Layer0, LM_L_l26_Layer0, LM_L_l27_Layer0, LM_L_l28_Layer0, LM_L_l31_Layer0, LM_L_l32_Layer0, LM_L_l33_Layer0, LM_L_l34_Layer0, LM_L_l35_Layer0, LM_L_l81_Layer0)
Dim BP_Layer1: BP_Layer1=Array(BM_Layer1, LM_FL_F17_Dome_Layer1, LM_FL_F20_Layer1, LM_FL_F21a_Layer1, LM_FL_F23_Layer1, LM_GI_git01_Layer1, LM_L_l16_Layer1)
Dim BP_Layer2: BP_Layer2=Array(BM_Layer2, LM_FL_F17_Dome_Layer2, LM_FL_F18_Layer2, LM_FL_F19a_Layer2, LM_FL_F20_Layer2, LM_FL_F20_Dome_Layer2, LM_FL_F22_Layer2, LM_FL_F23_Layer2, LM_L_L61_Layer2, LM_L_L62_Layer2, LM_L_L63_Layer2, LM_L_Bmp_Layer2, LM_GI_gim05_Layer2, LM_GI_git01_Layer2, LM_L_l16_Layer2, LM_L_l21_Layer2, LM_L_l22_Layer2, LM_L_l51_Layer2, LM_L_l52_Layer2, LM_L_l58_Layer2, LM_L_l64_Layer2, LM_L_l65_Layer2, LM_L_l66_Layer2, LM_L_l83_Layer2, LM_L_l84_Layer2)
Dim BP_Layer3: BP_Layer3=Array(BM_Layer3, LM_FL_F17_Dome_Layer3, LM_FL_F18_Layer3, LM_FL_F19_Layer3, LM_FL_F19a_Layer3, LM_FL_F20_Layer3, LM_FL_F20_Dome_Layer3, LM_FL_F21_Layer3, LM_FL_F21a_Layer3, LM_FL_F22_Layer3, LM_FL_F23_Layer3, LM_FL_F24_Layer3, LM_L_L61_Layer3, LM_L_L62_Layer3, LM_L_L63_Layer3, LM_L_L68_Layer3, LM_L_Bmp_Layer3, LM_GI_gib05_Layer3, LM_GI_gim01_Layer3, LM_GI_gim04_Layer3, LM_GI_gim05_Layer3, LM_GI_git01_Layer3, LM_L_l16_Layer3, LM_L_l17_Layer3, LM_L_l18_Layer3, LM_L_l22_Layer3, LM_L_l23_Layer3, LM_L_l27_Layer3, LM_L_l31_Layer3, LM_L_l32_Layer3, LM_L_l33_Layer3, LM_L_l34_Layer3, LM_L_l35_Layer3, LM_L_l41_Layer3, LM_L_l42_Layer3, LM_L_l48_Layer3, LM_L_l58_Layer3, LM_L_l64_Layer3, LM_L_l65_Layer3, LM_L_l67_Layer3, LM_L_l72_Layer3, LM_L_l83_Layer3, LM_L_l84_Layer3, LM_L_l85_Layer3)
Dim BP_Layer4: BP_Layer4=Array(BM_Layer4, LM_FL_F17_Dome_Layer4, LM_FL_F19a_Layer4, LM_FL_F20_Layer4, LM_FL_F20_Dome_Layer4, LM_FL_F23_Layer4, LM_L_Bmp_Layer4, LM_GI_git01_Layer4, LM_L_l16_Layer4, LM_L_l18_Layer4, LM_L_l21_Layer4, LM_L_l22_Layer4, LM_L_l26_Layer4, LM_L_l27_Layer4, LM_L_l28_Layer4, LM_L_l33_Layer4, LM_L_l35_Layer4, LM_L_l51_Layer4, LM_L_l52_Layer4, LM_L_l64_Layer4, LM_L_l65_Layer4, LM_L_l84_Layer4)
Dim BP_Lgate: BP_Lgate=Array(BM_Lgate, LM_FL_F19_Lgate, LM_FL_F23_Lgate)
Dim BP_PF: BP_PF=Array(BM_PF, LM_FL_F17_Dome_PF, LM_FL_F18_PF, LM_FL_F19_PF, LM_FL_F19a_PF, LM_FL_F20_PF, LM_FL_F20_Dome_PF, LM_FL_F21_PF, LM_FL_F21a_PF, LM_FL_F22_PF, LM_FL_F23_PF, LM_FL_F24_PF, LM_L_L61_PF, LM_L_L62_PF, LM_L_L63_PF, LM_L_L68_PF, LM_L_Bmp_PF, LM_GI_gib01_PF, LM_GI_gib02_PF, LM_GI_gib03_PF, LM_GI_gib04_PF, LM_GI_gib05_PF, LM_GI_gib06_PF, LM_GI_gib08_PF, LM_GI_gim01_PF, LM_GI_gim04_PF, LM_GI_gim05_PF, LM_GI_git01_PF, LM_L_l16_PF, LM_L_l17_PF, LM_L_l18_PF, LM_L_l23_PF, LM_L_l24_PF, LM_L_l25_PF, LM_L_l26_PF, LM_L_l27_PF, LM_L_l28_PF, LM_L_l31_PF, LM_L_l32_PF, LM_L_l33_PF, LM_L_l34_PF, LM_L_l35_PF, LM_L_l36_PF, LM_L_l37_PF, LM_L_l38_PF, LM_L_l41_PF, LM_L_l42_PF, LM_L_l43_PF, LM_L_l44_PF, LM_L_l45_PF, LM_L_l46_PF, LM_L_l47_PF, LM_L_l48_PF, LM_L_l54_PF, LM_L_l55_PF, LM_L_l56_PF, LM_L_l57_PF, LM_L_l58_PF, LM_L_l64_PF, LM_L_l65_PF, LM_L_l66_PF, LM_L_l67_PF, LM_L_l71_PF, LM_L_l72_PF, LM_L_l73_PF, LM_L_l74_PF, LM_L_l75_PF, LM_L_l76_PF, LM_L_l77_PF, LM_L_l78_PF, LM_L_l81_PF, LM_L_l82_PF, LM_L_l83_PF, _
  LM_L_l84_PF, LM_L_l85_PF, LM_L_l86_PF, LM_L_ml1_PF, LM_L_ml10_PF, LM_L_ml11_PF, LM_L_ml12_PF, LM_L_ml13_PF, LM_L_ml2_PF, LM_L_ml3_PF, LM_L_ml4_PF, LM_L_ml5_PF, LM_L_ml6_PF, LM_L_ml7_PF, LM_L_ml8_PF, LM_L_ml9_PF)
Dim BP_Parts: BP_Parts=Array(BM_Parts, LM_FL_F17_Dome_Parts, LM_FL_F18_Parts, LM_FL_F19_Parts, LM_FL_F19a_Parts, LM_FL_F20_Parts, LM_FL_F20_Dome_Parts, LM_FL_F21_Parts, LM_FL_F21a_Parts, LM_FL_F22_Parts, LM_FL_F23_Parts, LM_FL_F24_Parts, LM_L_L61_Parts, LM_L_L62_Parts, LM_L_L63_Parts, LM_L_L68_Parts, LM_L_Bmp_Parts, LM_GI_gib01_Parts, LM_GI_gib02_Parts, LM_GI_gib03_Parts, LM_GI_gib04_Parts, LM_GI_gib05_Parts, LM_GI_gib06_Parts, LM_GI_gib08_Parts, LM_GI_gim01_Parts, LM_GI_gim04_Parts, LM_GI_gim05_Parts, LM_GI_git01_Parts, LM_L_l16_Parts, LM_L_l17_Parts, LM_L_l18_Parts, LM_L_l21_Parts, LM_L_l22_Parts, LM_L_l24_Parts, LM_L_l25_Parts, LM_L_l26_Parts, LM_L_l27_Parts, LM_L_l28_Parts, LM_L_l31_Parts, LM_L_l32_Parts, LM_L_l33_Parts, LM_L_l34_Parts, LM_L_l35_Parts, LM_L_l37_Parts, LM_L_l38_Parts, LM_L_l41_Parts, LM_L_l42_Parts, LM_L_l48_Parts, LM_L_l51_Parts, LM_L_l52_Parts, LM_L_l55_Parts, LM_L_l56_Parts, LM_L_l57_Parts, LM_L_l58_Parts, LM_L_l64_Parts, LM_L_l65_Parts, LM_L_l66_Parts, LM_L_l73_Parts, LM_L_l74_Parts, _
  LM_L_l75_Parts, LM_L_l76_Parts, LM_L_l77_Parts, LM_L_l78_Parts, LM_L_l81_Parts, LM_L_l83_Parts, LM_L_l84_Parts, LM_L_l85_Parts, LM_L_l86_Parts, LM_L_ml1_Parts, LM_L_ml11_Parts)
Dim BP_PartsAbove: BP_PartsAbove=Array(BM_PartsAbove, LM_FL_F17_Dome_PartsAbove, LM_FL_F18_PartsAbove, LM_FL_F19_PartsAbove, LM_FL_F19a_PartsAbove, LM_FL_F20_PartsAbove, LM_FL_F20_Dome_PartsAbove, LM_FL_F21_PartsAbove, LM_FL_F21a_PartsAbove, LM_FL_F22_PartsAbove, LM_FL_F23_PartsAbove, LM_FL_F24_PartsAbove, LM_L_L61_PartsAbove, LM_L_L62_PartsAbove, LM_L_L63_PartsAbove, LM_GI_gib05_PartsAbove, LM_GI_gim01_PartsAbove, LM_GI_git01_PartsAbove, LM_L_l31_PartsAbove, LM_L_l32_PartsAbove, LM_L_l33_PartsAbove, LM_L_l34_PartsAbove, LM_L_l41_PartsAbove, LM_L_l51_PartsAbove, LM_L_l58_PartsAbove, LM_L_l64_PartsAbove, LM_L_l65_PartsAbove, LM_L_l66_PartsAbove, LM_L_l83_PartsAbove, LM_L_l84_PartsAbove)
Dim BP_RFlip: BP_RFlip=Array(BM_RFlip, LM_GI_gib06_RFlip, LM_GI_gib08_RFlip, LM_L_l55_RFlip)
Dim BP_RFlipU: BP_RFlipU=Array(BM_RFlipU, LM_GI_gib06_RFlipU, LM_GI_gib08_RFlipU, LM_L_l46_RFlipU, LM_L_l47_RFlipU)
Dim BP_RGate: BP_RGate=Array(BM_RGate, LM_FL_F21_RGate, LM_FL_F23_RGate, LM_FL_F24_RGate, LM_GI_git01_RGate)
Dim BP_RPostRubber_p1: BP_RPostRubber_p1=Array(BM_RPostRubber_p1, LM_FL_F19_RPostRubber_p1, LM_FL_F21_RPostRubber_p1, LM_FL_F21a_RPostRubber_p1, LM_GI_gim01_RPostRubber_p1)
Dim BP_RPostRubber_p2: BP_RPostRubber_p2=Array(BM_RPostRubber_p2, LM_FL_F19_RPostRubber_p2, LM_FL_F21_RPostRubber_p2, LM_FL_F21a_RPostRubber_p2, LM_GI_gim01_RPostRubber_p2, LM_L_l41_RPostRubber_p2)
Dim BP_RPostRubber_p3: BP_RPostRubber_p3=Array(BM_RPostRubber_p3, LM_FL_F19_RPostRubber_p3, LM_FL_F21_RPostRubber_p3, LM_FL_F21a_RPostRubber_p3, LM_GI_gim01_RPostRubber_p3, LM_L_l41_RPostRubber_p3)
Dim BP_RPost_p3: BP_RPost_p3=Array(BM_RPost_p3, LM_FL_F21_RPost_p3, LM_FL_F21a_RPost_p3, LM_GI_gib05_RPost_p3, LM_GI_gim01_RPost_p3, LM_L_l42_RPost_p3)
Dim BP_RRamp_DP: BP_RRamp_DP=Array(BM_RRamp_DP, LM_FL_F17_Dome_RRamp_DP, LM_FL_F19a_RRamp_DP, LM_FL_F23_RRamp_DP, LM_L_Bmp_RRamp_DP, LM_GI_git01_RRamp_DP, LM_L_l16_RRamp_DP)
Dim BP_RRamp_UP: BP_RRamp_UP=Array(BM_RRamp_UP, LM_FL_F17_Dome_RRamp_UP, LM_FL_F19a_RRamp_UP, LM_L_Bmp_RRamp_UP, LM_GI_git01_RRamp_UP, LM_L_l16_RRamp_UP, LM_L_l27_RRamp_UP)
Dim BP_RSling: BP_RSling=Array(BM_RSling, LM_FL_F21a_RSling, LM_GI_gib03_RSling, LM_GI_gib05_RSling, LM_GI_gib06_RSling, LM_GI_gib08_RSling, LM_GI_gim01_RSling, LM_L_l41_RSling)
Dim BP_RSling1: BP_RSling1=Array(BM_RSling1, LM_FL_F19_RSling1, LM_GI_gib05_RSling1, LM_GI_gib06_RSling1)
Dim BP_RSling2: BP_RSling2=Array(BM_RSling2, LM_FL_F21a_RSling2, LM_GI_gib03_RSling2, LM_GI_gib05_RSling2, LM_GI_gib06_RSling2, LM_GI_gib08_RSling2)
Dim BP_RSlingArm: BP_RSlingArm=Array(BM_RSlingArm, LM_FL_F21a_RSlingArm, LM_GI_gib03_RSlingArm, LM_GI_gib05_RSlingArm, LM_GI_gib06_RSlingArm, LM_GI_gib08_RSlingArm)
Dim BP_RailL: BP_RailL=Array(BM_RailL)
Dim BP_RailR: BP_RailR=Array(BM_RailR)
Dim BP_Sculpts: BP_Sculpts=Array(BM_Sculpts, LM_FL_F18_Sculpts, LM_FL_F19_Sculpts, LM_FL_F19a_Sculpts, LM_FL_F20_Sculpts, LM_FL_F20_Dome_Sculpts, LM_FL_F21_Sculpts, LM_FL_F21a_Sculpts, LM_FL_F22_Sculpts, LM_FL_F23_Sculpts, LM_FL_F24_Sculpts, LM_GI_gim01_Sculpts, LM_GI_gim04_Sculpts, LM_GI_git01_Sculpts, LM_L_l18_Sculpts, LM_L_l24_Sculpts, LM_L_l26_Sculpts, LM_L_l28_Sculpts, LM_L_l33_Sculpts, LM_L_l34_Sculpts, LM_L_l35_Sculpts, LM_L_l37_Sculpts, LM_L_l38_Sculpts, LM_L_l41_Sculpts, LM_L_l51_Sculpts, LM_L_l52_Sculpts, LM_L_l57_Sculpts, LM_L_l73_Sculpts, LM_L_l74_Sculpts, LM_L_l75_Sculpts, LM_L_l81_Sculpts)
Dim BP_UnderPF: BP_UnderPF=Array(BM_UnderPF, LM_FL_F18_UnderPF, LM_FL_F19_UnderPF, LM_FL_F19a_UnderPF, LM_FL_F20_UnderPF, LM_FL_F21_UnderPF, LM_FL_F21a_UnderPF, LM_FL_F22_UnderPF, LM_FL_F23_UnderPF, LM_FL_F24_UnderPF, LM_L_L61_UnderPF, LM_L_L62_UnderPF, LM_L_L63_UnderPF, LM_L_Bmp_UnderPF, LM_GI_gib01_UnderPF, LM_GI_gim01_UnderPF, LM_GI_gim04_UnderPF, LM_GI_gim05_UnderPF, LM_GI_git01_UnderPF, LM_L_l16_UnderPF, LM_L_l18_UnderPF, LM_L_l21_UnderPF, LM_L_l22_UnderPF, LM_L_l23_UnderPF, LM_L_l24_UnderPF, LM_L_l25_UnderPF, LM_L_l26_UnderPF, LM_L_l27_UnderPF, LM_L_l28_UnderPF, LM_L_l31_UnderPF, LM_L_l32_UnderPF, LM_L_l33_UnderPF, LM_L_l34_UnderPF, LM_L_l35_UnderPF, LM_L_l36_UnderPF, LM_L_l37_UnderPF, LM_L_l38_UnderPF, LM_L_l41_UnderPF, LM_L_l42_UnderPF, LM_L_l43_UnderPF, LM_L_l44_UnderPF, LM_L_l45_UnderPF, LM_L_l46_UnderPF, LM_L_l47_UnderPF, LM_L_l48_UnderPF, LM_L_l51_UnderPF, LM_L_l52_UnderPF, LM_L_l54_UnderPF, LM_L_l55_UnderPF, LM_L_l57_UnderPF, LM_L_l58_UnderPF, LM_L_l64_UnderPF, LM_L_l65_UnderPF, LM_L_l66_UnderPF, _
  LM_L_l67_UnderPF, LM_L_l71_UnderPF, LM_L_l72_UnderPF, LM_L_l73_UnderPF, LM_L_l74_UnderPF, LM_L_l75_UnderPF, LM_L_l76_UnderPF, LM_L_l77_UnderPF, LM_L_l78_UnderPF, LM_L_l82_UnderPF, LM_L_l83_UnderPF, LM_L_l84_UnderPF, LM_L_l85_UnderPF, LM_L_l86_UnderPF)
Dim BP_lockdownbar: BP_lockdownbar=Array(BM_lockdownbar)
Dim BP_sw15: BP_sw15=Array(BM_sw15, LM_FL_F20_sw15, LM_FL_F20_Dome_sw15)
Dim BP_sw16: BP_sw16=Array(BM_sw16, LM_FL_F20_sw16, LM_FL_F20_Dome_sw16, LM_FL_F24_sw16, LM_GI_git01_sw16)
Dim BP_sw25: BP_sw25=Array(BM_sw25, LM_GI_git01_sw25)
Dim BP_sw26: BP_sw26=Array(BM_sw26, LM_FL_F20_Dome_sw26, LM_GI_git01_sw26)
Dim BP_sw27: BP_sw27=Array(BM_sw27, LM_FL_F22_sw27, LM_GI_git01_sw27)
Dim BP_sw28: BP_sw28=Array(BM_sw28)
Dim BP_sw31: BP_sw31=Array(BM_sw31, LM_FL_F21a_sw31)
Dim BP_sw35: BP_sw35=Array(BM_sw35, LM_GI_gib02_sw35, LM_GI_gib03_sw35)
Dim BP_sw36: BP_sw36=Array(BM_sw36, LM_GI_gib01_sw36, LM_GI_gib02_sw36, LM_GI_gib04_sw36)
Dim BP_sw37: BP_sw37=Array(BM_sw37, LM_FL_F21a_sw37, LM_GI_gib03_sw37, LM_GI_gib05_sw37, LM_GI_gib06_sw37)
Dim BP_sw38: BP_sw38=Array(BM_sw38, LM_FL_F21a_sw38)
Dim BP_sw66: BP_sw66=Array(BM_sw66, LM_FL_F18_sw66, LM_FL_F19_sw66, LM_FL_F19a_sw66, LM_FL_F22_sw66, LM_FL_F24_sw66, LM_GI_gim01_sw66, LM_GI_gim04_sw66, LM_L_l73_sw66, LM_L_l74_sw66)
Dim BP_sw67: BP_sw67=Array(BM_sw67, LM_FL_F18_sw67, LM_FL_F19_sw67, LM_FL_F19a_sw67, LM_FL_F22_sw67, LM_FL_F24_sw67, LM_GI_gim01_sw67, LM_L_l73_sw67, LM_L_l74_sw67, LM_L_l75_sw67)
Dim BP_sw68: BP_sw68=Array(BM_sw68, LM_FL_F18_sw68, LM_FL_F19_sw68, LM_FL_F19a_sw68, LM_FL_F22_sw68, LM_FL_F24_sw68, LM_GI_gim01_sw68, LM_L_l73_sw68, LM_L_l74_sw68, LM_L_l75_sw68)
Dim BP_sw84: BP_sw84=Array(BM_sw84, LM_GI_git01_sw84)
Dim BP_sw85: BP_sw85=Array(BM_sw85, LM_GI_git01_sw85)
Dim BP_sw86: BP_sw86=Array(BM_sw86, LM_FL_F18_sw86, LM_FL_F19a_sw86, LM_FL_F21_sw86, LM_FL_F21a_sw86, LM_FL_F22_sw86, LM_FL_F23_sw86, LM_FL_F24_sw86, LM_GI_gim01_sw86, LM_GI_gim04_sw86, LM_GI_gim05_sw86, LM_L_l76_sw86, LM_L_l77_sw86)
Dim BP_sw87: BP_sw87=Array(BM_sw87, LM_FL_F18_sw87, LM_FL_F19a_sw87, LM_FL_F21_sw87, LM_FL_F21a_sw87, LM_FL_F22_sw87, LM_FL_F23_sw87, LM_GI_gim01_sw87, LM_GI_gim04_sw87, LM_GI_gim05_sw87, LM_L_l76_sw87, LM_L_l77_sw87, LM_L_l78_sw87)
Dim BP_sw88: BP_sw88=Array(BM_sw88, LM_FL_F19a_sw88, LM_FL_F21_sw88, LM_FL_F21a_sw88, LM_FL_F22_sw88, LM_FL_F23_sw88, LM_GI_gim01_sw88, LM_GI_gim04_sw88, LM_GI_gim05_sw88, LM_L_l25_sw88, LM_L_l77_sw88, LM_L_l78_sw88)
' Arrays per lighting scenario
Dim BL_FL_F17_Dome: BL_FL_F17_Dome=Array(LM_FL_F17_Dome_Diverter, LM_FL_F17_Dome_DiverterU, LM_FL_F17_Dome_Gate6, LM_FL_F17_Dome_Layer1, LM_FL_F17_Dome_Layer2, LM_FL_F17_Dome_Layer3, LM_FL_F17_Dome_Layer4, LM_FL_F17_Dome_PF, LM_FL_F17_Dome_Parts, LM_FL_F17_Dome_PartsAbove, LM_FL_F17_Dome_RRamp_DP, LM_FL_F17_Dome_RRamp_UP)
Dim BL_FL_F18: BL_FL_F18=Array(LM_FL_F18_BmpsBtm, LM_FL_F18_BmpsTop, LM_FL_F18_Bumper1_Ring, LM_FL_F18_Bumper1_Socket, LM_FL_F18_Bumper2_Ring, LM_FL_F18_Bumper2_Socket, LM_FL_F18_Bumper3_Ring, LM_FL_F18_Bumper3_Socket, LM_FL_F18_Layer2, LM_FL_F18_Layer3, LM_FL_F18_PF, LM_FL_F18_Parts, LM_FL_F18_PartsAbove, LM_FL_F18_Sculpts, LM_FL_F18_UnderPF, LM_FL_F18_sw66, LM_FL_F18_sw67, LM_FL_F18_sw68, LM_FL_F18_sw86, LM_FL_F18_sw87)
Dim BL_FL_F19: BL_FL_F19=Array(LM_FL_F19_BmpsTop, LM_FL_F19_Bumper1_Ring, LM_FL_F19_ExtPlasticScrew2, LM_FL_F19_ExtPlasticScrew3, LM_FL_F19_Layer3, LM_FL_F19_Lgate, LM_FL_F19_PF, LM_FL_F19_Parts, LM_FL_F19_PartsAbove, LM_FL_F19_RPostRubber_p1, LM_FL_F19_RPostRubber_p2, LM_FL_F19_RPostRubber_p3, LM_FL_F19_RSling1, LM_FL_F19_Sculpts, LM_FL_F19_UnderPF, LM_FL_F19_sw66, LM_FL_F19_sw67, LM_FL_F19_sw68)
Dim BL_FL_F19a: BL_FL_F19a=Array(LM_FL_F19a_BmpsBtm, LM_FL_F19a_BmpsTop, LM_FL_F19a_Bumper1_Ring, LM_FL_F19a_Bumper1_Socket, LM_FL_F19a_Bumper2_Ring, LM_FL_F19a_Bumper2_Socket, LM_FL_F19a_Bumper3_Ring, LM_FL_F19a_Bumper3_Socket, LM_FL_F19a_ExtPlasticScrew3, LM_FL_F19a_LSling_001, LM_FL_F19a_Layer0, LM_FL_F19a_Layer2, LM_FL_F19a_Layer3, LM_FL_F19a_Layer4, LM_FL_F19a_PF, LM_FL_F19a_Parts, LM_FL_F19a_PartsAbove, LM_FL_F19a_RRamp_DP, LM_FL_F19a_RRamp_UP, LM_FL_F19a_Sculpts, LM_FL_F19a_UnderPF, LM_FL_F19a_sw66, LM_FL_F19a_sw67, LM_FL_F19a_sw68, LM_FL_F19a_sw86, LM_FL_F19a_sw87, LM_FL_F19a_sw88)
Dim BL_FL_F20: BL_FL_F20=Array(LM_FL_F20_BmpsTop, LM_FL_F20_Bumper2_Socket, LM_FL_F20_Layer1, LM_FL_F20_Layer2, LM_FL_F20_Layer3, LM_FL_F20_Layer4, LM_FL_F20_PF, LM_FL_F20_Parts, LM_FL_F20_PartsAbove, LM_FL_F20_Sculpts, LM_FL_F20_UnderPF, LM_FL_F20_sw15, LM_FL_F20_sw16)
Dim BL_FL_F20_Dome: BL_FL_F20_Dome=Array(LM_FL_F20_Dome_BmpsBtm, LM_FL_F20_Dome_BmpsTop, LM_FL_F20_Dome_Layer2, LM_FL_F20_Dome_Layer3, LM_FL_F20_Dome_Layer4, LM_FL_F20_Dome_PF, LM_FL_F20_Dome_Parts, LM_FL_F20_Dome_PartsAbove, LM_FL_F20_Dome_Sculpts, LM_FL_F20_Dome_sw15, LM_FL_F20_Dome_sw16, LM_FL_F20_Dome_sw26)
Dim BL_FL_F21: BL_FL_F21=Array(LM_FL_F21_BmpsTop, LM_FL_F21_Bumper1_Ring, LM_FL_F21_ExtPlasticScrew3, LM_FL_F21_LSling_001, LM_FL_F21_LSling1, LM_FL_F21_Layer0, LM_FL_F21_Layer3, LM_FL_F21_PF, LM_FL_F21_Parts, LM_FL_F21_PartsAbove, LM_FL_F21_RGate, LM_FL_F21_RPostRubber_p1, LM_FL_F21_RPostRubber_p2, LM_FL_F21_RPostRubber_p3, LM_FL_F21_RPost_p3, LM_FL_F21_Sculpts, LM_FL_F21_UnderPF, LM_FL_F21_sw86, LM_FL_F21_sw87, LM_FL_F21_sw88)
Dim BL_FL_F21a: BL_FL_F21a=Array(LM_FL_F21a_Bumper1_Ring, LM_FL_F21a_ExtPlasticScrew1, LM_FL_F21a_ExtPlasticScrew3, LM_FL_F21a_ExtPlasticScrew4, LM_FL_F21a_ExtraPlasticR, LM_FL_F21a_LPost_p3, LM_FL_F21a_LRampLever, LM_FL_F21a_LRamp_UP, LM_FL_F21a_LSling_001, LM_FL_F21a_LSling1, LM_FL_F21a_LSling2, LM_FL_F21a_LSlingArm, LM_FL_F21a_Layer0, LM_FL_F21a_Layer1, LM_FL_F21a_Layer3, LM_FL_F21a_PF, LM_FL_F21a_Parts, LM_FL_F21a_PartsAbove, LM_FL_F21a_RPostRubber_p1, LM_FL_F21a_RPostRubber_p2, LM_FL_F21a_RPostRubber_p3, LM_FL_F21a_RPost_p3, LM_FL_F21a_RSling, LM_FL_F21a_RSling2, LM_FL_F21a_RSlingArm, LM_FL_F21a_Sculpts, LM_FL_F21a_UnderPF, LM_FL_F21a_sw31, LM_FL_F21a_sw37, LM_FL_F21a_sw38, LM_FL_F21a_sw86, LM_FL_F21a_sw87, LM_FL_F21a_sw88)
Dim BL_FL_F22: BL_FL_F22=Array(LM_FL_F22_BmpsBtm, LM_FL_F22_BmpsTop, LM_FL_F22_Bumper1_Ring, LM_FL_F22_Bumper1_Socket, LM_FL_F22_Bumper2_Ring, LM_FL_F22_Bumper2_Socket, LM_FL_F22_Bumper3_Ring, LM_FL_F22_Bumper3_Socket, LM_FL_F22_LPostRubber_p1, LM_FL_F22_LPostRubber_p2, LM_FL_F22_LPostRubber_p3, LM_FL_F22_Layer0, LM_FL_F22_Layer2, LM_FL_F22_Layer3, LM_FL_F22_PF, LM_FL_F22_Parts, LM_FL_F22_PartsAbove, LM_FL_F22_Sculpts, LM_FL_F22_UnderPF, LM_FL_F22_sw27, LM_FL_F22_sw66, LM_FL_F22_sw67, LM_FL_F22_sw68, LM_FL_F22_sw86, LM_FL_F22_sw87, LM_FL_F22_sw88)
Dim BL_FL_F23: BL_FL_F23=Array(LM_FL_F23_BmpsBtm, LM_FL_F23_BmpsTop, LM_FL_F23_Bumper1_Ring, LM_FL_F23_Bumper1_Socket, LM_FL_F23_Bumper2_Ring, LM_FL_F23_Bumper2_Socket, LM_FL_F23_Bumper3_Ring, LM_FL_F23_Bumper3_Socket, LM_FL_F23_Layer0, LM_FL_F23_Layer1, LM_FL_F23_Layer2, LM_FL_F23_Layer3, LM_FL_F23_Layer4, LM_FL_F23_Lgate, LM_FL_F23_PF, LM_FL_F23_Parts, LM_FL_F23_PartsAbove, LM_FL_F23_RGate, LM_FL_F23_RRamp_DP, LM_FL_F23_Sculpts, LM_FL_F23_UnderPF, LM_FL_F23_sw86, LM_FL_F23_sw87, LM_FL_F23_sw88)
Dim BL_FL_F24: BL_FL_F24=Array(LM_FL_F24_BmpsTop, LM_FL_F24_Bumper1_Ring, LM_FL_F24_Layer3, LM_FL_F24_PF, LM_FL_F24_Parts, LM_FL_F24_PartsAbove, LM_FL_F24_RGate, LM_FL_F24_Sculpts, LM_FL_F24_UnderPF, LM_FL_F24_sw16, LM_FL_F24_sw66, LM_FL_F24_sw67, LM_FL_F24_sw68, LM_FL_F24_sw86)
Dim BL_GI_gib01: BL_GI_gib01=Array(LM_GI_gib01_ExtPlasticScrew3, LM_GI_gib01_ExtPlasticSpacer3, LM_GI_gib01_ExtraPlasticL, LM_GI_gib01_LPostRubber_p1, LM_GI_gib01_LPostRubber_p2, LM_GI_gib01_LPostRubber_p3, LM_GI_gib01_LPost_p3, LM_GI_gib01_LSling_001, LM_GI_gib01_LSling1, LM_GI_gib01_LSling2, LM_GI_gib01_LSlingArm, LM_GI_gib01_PF, LM_GI_gib01_Parts, LM_GI_gib01_UnderPF, LM_GI_gib01_sw36)
Dim BL_GI_gib02: BL_GI_gib02=Array(LM_GI_gib02_ExtPlasticScrew1, LM_GI_gib02_ExtraPlasticL, LM_GI_gib02_LFlip, LM_GI_gib02_LFlipU, LM_GI_gib02_LSling_001, LM_GI_gib02_LSling1, LM_GI_gib02_LSling2, LM_GI_gib02_LSlingArm, LM_GI_gib02_PF, LM_GI_gib02_Parts, LM_GI_gib02_sw35, LM_GI_gib02_sw36)
Dim BL_GI_gib03: BL_GI_gib03=Array(LM_GI_gib03_ExtPlasticSpacer2, LM_GI_gib03_ExtraPlasticL, LM_GI_gib03_ExtraPlasticR, LM_GI_gib03_LSling_001, LM_GI_gib03_LSlingArm, LM_GI_gib03_PF, LM_GI_gib03_Parts, LM_GI_gib03_RSling, LM_GI_gib03_RSling2, LM_GI_gib03_RSlingArm, LM_GI_gib03_sw35, LM_GI_gib03_sw37)
Dim BL_GI_gib04: BL_GI_gib04=Array(LM_GI_gib04_ExtraPlasticL, LM_GI_gib04_LFlip, LM_GI_gib04_LFlipU, LM_GI_gib04_LSling_001, LM_GI_gib04_LSlingArm, LM_GI_gib04_PF, LM_GI_gib04_Parts, LM_GI_gib04_sw36)
Dim BL_GI_gib05: BL_GI_gib05=Array(LM_GI_gib05_ExtPlasticSpacer4, LM_GI_gib05_ExtraPlasticR, LM_GI_gib05_Layer3, LM_GI_gib05_PF, LM_GI_gib05_Parts, LM_GI_gib05_PartsAbove, LM_GI_gib05_RPost_p3, LM_GI_gib05_RSling, LM_GI_gib05_RSling1, LM_GI_gib05_RSling2, LM_GI_gib05_RSlingArm, LM_GI_gib05_sw37)
Dim BL_GI_gib06: BL_GI_gib06=Array(LM_GI_gib06_ExtraPlasticR, LM_GI_gib06_PF, LM_GI_gib06_Parts, LM_GI_gib06_RFlip, LM_GI_gib06_RFlipU, LM_GI_gib06_RSling, LM_GI_gib06_RSling1, LM_GI_gib06_RSling2, LM_GI_gib06_RSlingArm, LM_GI_gib06_sw37)
Dim BL_GI_gib08: BL_GI_gib08=Array(LM_GI_gib08_ExtPlasticScrew2, LM_GI_gib08_ExtraPlasticL, LM_GI_gib08_PF, LM_GI_gib08_Parts, LM_GI_gib08_RFlip, LM_GI_gib08_RFlipU, LM_GI_gib08_RSling, LM_GI_gib08_RSling2, LM_GI_gib08_RSlingArm)
Dim BL_GI_gim01: BL_GI_gim01=Array(LM_GI_gim01_BmpsTop, LM_GI_gim01_Bumper1_Ring, LM_GI_gim01_ExtPlasticScrew4, LM_GI_gim01_ExtraPlasticR, LM_GI_gim01_LPost_p3, LM_GI_gim01_Layer0, LM_GI_gim01_Layer3, LM_GI_gim01_PF, LM_GI_gim01_Parts, LM_GI_gim01_PartsAbove, LM_GI_gim01_RPostRubber_p1, LM_GI_gim01_RPostRubber_p2, LM_GI_gim01_RPostRubber_p3, LM_GI_gim01_RPost_p3, LM_GI_gim01_RSling, LM_GI_gim01_Sculpts, LM_GI_gim01_UnderPF, LM_GI_gim01_sw66, LM_GI_gim01_sw67, LM_GI_gim01_sw68, LM_GI_gim01_sw86, LM_GI_gim01_sw87, LM_GI_gim01_sw88)
Dim BL_GI_gim04: BL_GI_gim04=Array(LM_GI_gim04_BmpsBtm, LM_GI_gim04_BmpsTop, LM_GI_gim04_Bumper1_Ring, LM_GI_gim04_Bumper1_Socket, LM_GI_gim04_Bumper2_Ring, LM_GI_gim04_Bumper2_Socket, LM_GI_gim04_Layer3, LM_GI_gim04_PF, LM_GI_gim04_Parts, LM_GI_gim04_Sculpts, LM_GI_gim04_UnderPF, LM_GI_gim04_sw66, LM_GI_gim04_sw86, LM_GI_gim04_sw87, LM_GI_gim04_sw88)
Dim BL_GI_gim05: BL_GI_gim05=Array(LM_GI_gim05_BmpsBtm, LM_GI_gim05_BmpsTop, LM_GI_gim05_Bumper1_Ring, LM_GI_gim05_Bumper1_Socket, LM_GI_gim05_Bumper3_Ring, LM_GI_gim05_Bumper3_Socket, LM_GI_gim05_Layer2, LM_GI_gim05_Layer3, LM_GI_gim05_PF, LM_GI_gim05_Parts, LM_GI_gim05_UnderPF, LM_GI_gim05_sw86, LM_GI_gim05_sw87, LM_GI_gim05_sw88)
Dim BL_GI_git01: BL_GI_git01=Array(LM_GI_git01_BmpsBtm, LM_GI_git01_BmpsTop, LM_GI_git01_Bumper1_Socket, LM_GI_git01_Bumper2_Ring, LM_GI_git01_Bumper2_Socket, LM_GI_git01_Bumper3_Ring, LM_GI_git01_Bumper3_Socket, LM_GI_git01_Diverter, LM_GI_git01_DiverterU, LM_GI_git01_Gate6, LM_GI_git01_Layer0, LM_GI_git01_Layer1, LM_GI_git01_Layer2, LM_GI_git01_Layer3, LM_GI_git01_Layer4, LM_GI_git01_PF, LM_GI_git01_Parts, LM_GI_git01_PartsAbove, LM_GI_git01_RGate, LM_GI_git01_RRamp_DP, LM_GI_git01_RRamp_UP, LM_GI_git01_Sculpts, LM_GI_git01_UnderPF, LM_GI_git01_sw16, LM_GI_git01_sw25, LM_GI_git01_sw26, LM_GI_git01_sw27, LM_GI_git01_sw84, LM_GI_git01_sw85)
Dim BL_L_L61: BL_L_L61=Array(LM_L_L61_Bumper2_Socket, LM_L_L61_Layer2, LM_L_L61_Layer3, LM_L_L61_PF, LM_L_L61_Parts, LM_L_L61_PartsAbove, LM_L_L61_UnderPF)
Dim BL_L_L62: BL_L_L62=Array(LM_L_L62_Layer2, LM_L_L62_Layer3, LM_L_L62_PF, LM_L_L62_Parts, LM_L_L62_PartsAbove, LM_L_L62_UnderPF)
Dim BL_L_L63: BL_L_L63=Array(LM_L_L63_DiverterU, LM_L_L63_Gate6, LM_L_L63_Layer2, LM_L_L63_Layer3, LM_L_L63_PF, LM_L_L63_Parts, LM_L_L63_PartsAbove, LM_L_L63_UnderPF)
Dim BL_L_L68: BL_L_L68=Array(LM_L_L68_BmpsBtm, LM_L_L68_BmpsTop, LM_L_L68_Bumper2_Ring, LM_L_L68_Bumper3_Ring, LM_L_L68_Layer3, LM_L_L68_PF, LM_L_L68_Parts)
Dim BL_L_l16: BL_L_l16=Array(LM_L_l16_Layer0, LM_L_l16_Layer1, LM_L_l16_Layer2, LM_L_l16_Layer3, LM_L_l16_Layer4, LM_L_l16_PF, LM_L_l16_Parts, LM_L_l16_RRamp_DP, LM_L_l16_RRamp_UP, LM_L_l16_UnderPF)
Dim BL_L_l17: BL_L_l17=Array(LM_L_l17_Layer3, LM_L_l17_PF, LM_L_l17_Parts)
Dim BL_L_l18: BL_L_l18=Array(LM_L_l18_Layer0, LM_L_l18_Layer3, LM_L_l18_Layer4, LM_L_l18_PF, LM_L_l18_Parts, LM_L_l18_Sculpts, LM_L_l18_UnderPF)
Dim BL_L_l21: BL_L_l21=Array(LM_L_l21_Layer2, LM_L_l21_Layer4, LM_L_l21_Parts, LM_L_l21_UnderPF)
Dim BL_L_l22: BL_L_l22=Array(LM_L_l22_Layer2, LM_L_l22_Layer3, LM_L_l22_Layer4, LM_L_l22_Parts, LM_L_l22_UnderPF)
Dim BL_L_l23: BL_L_l23=Array(LM_L_l23_Layer3, LM_L_l23_PF, LM_L_l23_UnderPF)
Dim BL_L_l24: BL_L_l24=Array(LM_L_l24_Layer0, LM_L_l24_PF, LM_L_l24_Parts, LM_L_l24_Sculpts, LM_L_l24_UnderPF)
Dim BL_L_l25: BL_L_l25=Array(LM_L_l25_BmpsTop, LM_L_l25_PF, LM_L_l25_Parts, LM_L_l25_UnderPF, LM_L_l25_sw88)
Dim BL_L_l26: BL_L_l26=Array(LM_L_l26_Layer0, LM_L_l26_Layer4, LM_L_l26_PF, LM_L_l26_Parts, LM_L_l26_Sculpts, LM_L_l26_UnderPF)
Dim BL_L_l27: BL_L_l27=Array(LM_L_l27_Layer0, LM_L_l27_Layer3, LM_L_l27_Layer4, LM_L_l27_PF, LM_L_l27_Parts, LM_L_l27_RRamp_UP, LM_L_l27_UnderPF)
Dim BL_L_l28: BL_L_l28=Array(LM_L_l28_BmpsTop, LM_L_l28_Layer0, LM_L_l28_Layer4, LM_L_l28_PF, LM_L_l28_Parts, LM_L_l28_Sculpts, LM_L_l28_UnderPF)
Dim BL_L_l31: BL_L_l31=Array(LM_L_l31_Layer0, LM_L_l31_Layer3, LM_L_l31_PF, LM_L_l31_Parts, LM_L_l31_PartsAbove, LM_L_l31_UnderPF)
Dim BL_L_l32: BL_L_l32=Array(LM_L_l32_Layer0, LM_L_l32_Layer3, LM_L_l32_PF, LM_L_l32_Parts, LM_L_l32_PartsAbove, LM_L_l32_UnderPF)
Dim BL_L_l33: BL_L_l33=Array(LM_L_l33_Layer0, LM_L_l33_Layer3, LM_L_l33_Layer4, LM_L_l33_PF, LM_L_l33_Parts, LM_L_l33_PartsAbove, LM_L_l33_Sculpts, LM_L_l33_UnderPF)
Dim BL_L_l34: BL_L_l34=Array(LM_L_l34_Layer0, LM_L_l34_Layer3, LM_L_l34_PF, LM_L_l34_Parts, LM_L_l34_PartsAbove, LM_L_l34_Sculpts, LM_L_l34_UnderPF)
Dim BL_L_l35: BL_L_l35=Array(LM_L_l35_BmpsTop, LM_L_l35_Layer0, LM_L_l35_Layer3, LM_L_l35_Layer4, LM_L_l35_PF, LM_L_l35_Parts, LM_L_l35_Sculpts, LM_L_l35_UnderPF)
Dim BL_L_l36: BL_L_l36=Array(LM_L_l36_PF, LM_L_l36_UnderPF)
Dim BL_L_l37: BL_L_l37=Array(LM_L_l37_LPostRubber_p2, LM_L_l37_LPostRubber_p3, LM_L_l37_LPost_p3, LM_L_l37_PF, LM_L_l37_Parts, LM_L_l37_Sculpts, LM_L_l37_UnderPF)
Dim BL_L_l38: BL_L_l38=Array(LM_L_l38_ExtraPlasticL, LM_L_l38_LPostRubber_p1, LM_L_l38_LPostRubber_p2, LM_L_l38_LPostRubber_p3, LM_L_l38_LSling_001, LM_L_l38_PF, LM_L_l38_Parts, LM_L_l38_Sculpts, LM_L_l38_UnderPF)
Dim BL_L_l41: BL_L_l41=Array(LM_L_l41_Layer3, LM_L_l41_PF, LM_L_l41_Parts, LM_L_l41_PartsAbove, LM_L_l41_RPostRubber_p2, LM_L_l41_RPostRubber_p3, LM_L_l41_RSling, LM_L_l41_Sculpts, LM_L_l41_UnderPF)
Dim BL_L_l42: BL_L_l42=Array(LM_L_l42_Layer3, LM_L_l42_PF, LM_L_l42_Parts, LM_L_l42_RPost_p3, LM_L_l42_UnderPF)
Dim BL_L_l43: BL_L_l43=Array(LM_L_l43_LFlipU, LM_L_l43_PF, LM_L_l43_UnderPF)
Dim BL_L_l44: BL_L_l44=Array(LM_L_l44_LFlipU, LM_L_l44_PF, LM_L_l44_UnderPF)
Dim BL_L_l45: BL_L_l45=Array(LM_L_l45_PF, LM_L_l45_UnderPF)
Dim BL_L_l46: BL_L_l46=Array(LM_L_l46_PF, LM_L_l46_RFlipU, LM_L_l46_UnderPF)
Dim BL_L_l47: BL_L_l47=Array(LM_L_l47_PF, LM_L_l47_RFlipU, LM_L_l47_UnderPF)
Dim BL_L_l48: BL_L_l48=Array(LM_L_l48_Layer3, LM_L_l48_PF, LM_L_l48_Parts, LM_L_l48_UnderPF)
Dim BL_L_l51: BL_L_l51=Array(LM_L_l51_Layer2, LM_L_l51_Layer4, LM_L_l51_Parts, LM_L_l51_PartsAbove, LM_L_l51_Sculpts, LM_L_l51_UnderPF)
Dim BL_L_l52: BL_L_l52=Array(LM_L_l52_Layer2, LM_L_l52_Layer4, LM_L_l52_Parts, LM_L_l52_Sculpts, LM_L_l52_UnderPF)
Dim BL_L_l54: BL_L_l54=Array(LM_L_l54_PF, LM_L_l54_UnderPF)
Dim BL_L_l55: BL_L_l55=Array(LM_L_l55_LFlip, LM_L_l55_PF, LM_L_l55_Parts, LM_L_l55_RFlip, LM_L_l55_UnderPF)
Dim BL_L_l56: BL_L_l56=Array(LM_L_l56_PF, LM_L_l56_Parts)
Dim BL_L_l57: BL_L_l57=Array(LM_L_l57_PF, LM_L_l57_Parts, LM_L_l57_Sculpts, LM_L_l57_UnderPF)
Dim BL_L_l58: BL_L_l58=Array(LM_L_l58_Bumper2_Socket, LM_L_l58_Diverter, LM_L_l58_DiverterU, LM_L_l58_Layer2, LM_L_l58_Layer3, LM_L_l58_PF, LM_L_l58_Parts, LM_L_l58_PartsAbove, LM_L_l58_UnderPF)
Dim BL_L_l64: BL_L_l64=Array(LM_L_l64_DiverterU, LM_L_l64_Layer2, LM_L_l64_Layer3, LM_L_l64_Layer4, LM_L_l64_PF, LM_L_l64_Parts, LM_L_l64_PartsAbove, LM_L_l64_UnderPF)
Dim BL_L_l65: BL_L_l65=Array(LM_L_l65_Layer2, LM_L_l65_Layer3, LM_L_l65_Layer4, LM_L_l65_PF, LM_L_l65_Parts, LM_L_l65_PartsAbove, LM_L_l65_UnderPF)
Dim BL_L_l66: BL_L_l66=Array(LM_L_l66_Layer2, LM_L_l66_PF, LM_L_l66_Parts, LM_L_l66_PartsAbove, LM_L_l66_UnderPF)
Dim BL_L_l67: BL_L_l67=Array(LM_L_l67_Layer3, LM_L_l67_PF, LM_L_l67_UnderPF)
Dim BL_L_l71: BL_L_l71=Array(LM_L_l71_PF, LM_L_l71_UnderPF)
Dim BL_L_l72: BL_L_l72=Array(LM_L_l72_Layer3, LM_L_l72_PF, LM_L_l72_UnderPF)
Dim BL_L_l73: BL_L_l73=Array(LM_L_l73_PF, LM_L_l73_Parts, LM_L_l73_Sculpts, LM_L_l73_UnderPF, LM_L_l73_sw66, LM_L_l73_sw67, LM_L_l73_sw68)
Dim BL_L_l74: BL_L_l74=Array(LM_L_l74_PF, LM_L_l74_Parts, LM_L_l74_Sculpts, LM_L_l74_UnderPF, LM_L_l74_sw66, LM_L_l74_sw67, LM_L_l74_sw68)
Dim BL_L_l75: BL_L_l75=Array(LM_L_l75_PF, LM_L_l75_Parts, LM_L_l75_Sculpts, LM_L_l75_UnderPF, LM_L_l75_sw67, LM_L_l75_sw68)
Dim BL_L_l76: BL_L_l76=Array(LM_L_l76_PF, LM_L_l76_Parts, LM_L_l76_UnderPF, LM_L_l76_sw86, LM_L_l76_sw87)
Dim BL_L_l77: BL_L_l77=Array(LM_L_l77_PF, LM_L_l77_Parts, LM_L_l77_UnderPF, LM_L_l77_sw86, LM_L_l77_sw87, LM_L_l77_sw88)
Dim BL_L_l78: BL_L_l78=Array(LM_L_l78_PF, LM_L_l78_Parts, LM_L_l78_UnderPF, LM_L_l78_sw87, LM_L_l78_sw88)
Dim BL_L_l81: BL_L_l81=Array(LM_L_l81_Layer0, LM_L_l81_PF, LM_L_l81_Parts, LM_L_l81_Sculpts)
Dim BL_L_l82: BL_L_l82=Array(LM_L_l82_PF, LM_L_l82_UnderPF)
Dim BL_L_l83: BL_L_l83=Array(LM_L_l83_BmpsTop, LM_L_l83_Bumper2_Ring, LM_L_l83_Layer2, LM_L_l83_Layer3, LM_L_l83_PF, LM_L_l83_Parts, LM_L_l83_PartsAbove, LM_L_l83_UnderPF)
Dim BL_L_l84: BL_L_l84=Array(LM_L_l84_BmpsBtm, LM_L_l84_BmpsTop, LM_L_l84_Bumper2_Ring, LM_L_l84_Bumper2_Socket, LM_L_l84_Layer2, LM_L_l84_Layer3, LM_L_l84_Layer4, LM_L_l84_PF, LM_L_l84_Parts, LM_L_l84_PartsAbove, LM_L_l84_UnderPF)
Dim BL_L_l85: BL_L_l85=Array(LM_L_l85_BmpsTop, LM_L_l85_Bumper1_Ring, LM_L_l85_Layer3, LM_L_l85_PF, LM_L_l85_Parts, LM_L_l85_UnderPF)
Dim BL_L_l86: BL_L_l86=Array(LM_L_l86_PF, LM_L_l86_Parts, LM_L_l86_UnderPF)
Dim BL_L_ml1: BL_L_ml1=Array(LM_L_ml1_PF, LM_L_ml1_Parts)
Dim BL_L_ml10: BL_L_ml10=Array(LM_L_ml10_PF)
Dim BL_L_ml11: BL_L_ml11=Array(LM_L_ml11_PF, LM_L_ml11_Parts)
Dim BL_L_ml12: BL_L_ml12=Array(LM_L_ml12_PF)
Dim BL_L_ml13: BL_L_ml13=Array(LM_L_ml13_PF)
Dim BL_L_ml2: BL_L_ml2=Array(LM_L_ml2_PF)
Dim BL_L_ml3: BL_L_ml3=Array(LM_L_ml3_PF)
Dim BL_L_ml4: BL_L_ml4=Array(LM_L_ml4_PF)
Dim BL_L_ml5: BL_L_ml5=Array(LM_L_ml5_PF)
Dim BL_L_ml6: BL_L_ml6=Array(LM_L_ml6_PF)
Dim BL_L_ml7: BL_L_ml7=Array(LM_L_ml7_PF)
Dim BL_L_ml8: BL_L_ml8=Array(LM_L_ml8_PF)
Dim BL_L_ml9: BL_L_ml9=Array(LM_L_ml9_PF)
Dim BL_L_Bmp: BL_L_Bmp=Array(LM_L_Bmp_BmpsBtm, LM_L_Bmp_BmpsTop, LM_L_Bmp_Bumper1_Ring, LM_L_Bmp_Bumper1_Socket, LM_L_Bmp_Bumper2_Ring, LM_L_Bmp_Bumper2_Socket, LM_L_Bmp_Bumper3_Ring, LM_L_Bmp_Bumper3_Socket, LM_L_Bmp_Layer0, LM_L_Bmp_Layer2, LM_L_Bmp_Layer3, LM_L_Bmp_Layer4, LM_L_Bmp_PF, LM_L_Bmp_Parts, LM_L_Bmp_RRamp_DP, LM_L_Bmp_RRamp_UP, LM_L_Bmp_UnderPF)
Dim BL_World: BL_World=Array(BM_BmpsBtm, BM_BmpsTop, BM_Bumper1_Ring, BM_Bumper1_Socket, BM_Bumper2_Ring, BM_Bumper2_Socket, BM_Bumper3_Ring, BM_Bumper3_Socket, BM_Diverter, BM_DiverterU, BM_ExtPlasticScrew1, BM_ExtPlasticScrew2, BM_ExtPlasticScrew3, BM_ExtPlasticScrew4, BM_ExtPlasticSpacer1, BM_ExtPlasticSpacer2, BM_ExtPlasticSpacer3, BM_ExtPlasticSpacer4, BM_ExtraPlasticL, BM_ExtraPlasticR, BM_Gate4, BM_Gate6, BM_LFlip, BM_LFlipU, BM_LPostRubber_p1, BM_LPostRubber_p2, BM_LPostRubber_p3, BM_LPost_p3, BM_LRampLever, BM_LRamp_DP, BM_LRamp_UP, BM_LSling_001, BM_LSling1, BM_LSling2, BM_LSlingArm, BM_Layer0, BM_Layer1, BM_Layer2, BM_Layer3, BM_Layer4, BM_Lgate, BM_PF, BM_Parts, BM_PartsAbove, BM_RFlip, BM_RFlipU, BM_RGate, BM_RPostRubber_p1, BM_RPostRubber_p2, BM_RPostRubber_p3, BM_RPost_p3, BM_RRamp_DP, BM_RRamp_UP, BM_RSling, BM_RSling1, BM_RSling2, BM_RSlingArm, BM_RailL, BM_RailR, BM_Sculpts, BM_UnderPF, BM_lockdownbar, BM_sw15, BM_sw16, BM_sw25, BM_sw26, BM_sw27, BM_sw28, BM_sw31, BM_sw35, BM_sw36, BM_sw37, _
  BM_sw38, BM_sw66, BM_sw67, BM_sw68, BM_sw84, BM_sw85, BM_sw86, BM_sw87, BM_sw88)
' Global arrays
Dim BG_Bakemap: BG_Bakemap=Array(BM_BmpsBtm, BM_BmpsTop, BM_Bumper1_Ring, BM_Bumper1_Socket, BM_Bumper2_Ring, BM_Bumper2_Socket, BM_Bumper3_Ring, BM_Bumper3_Socket, BM_Diverter, BM_DiverterU, BM_ExtPlasticScrew1, BM_ExtPlasticScrew2, BM_ExtPlasticScrew3, BM_ExtPlasticScrew4, BM_ExtPlasticSpacer1, BM_ExtPlasticSpacer2, BM_ExtPlasticSpacer3, BM_ExtPlasticSpacer4, BM_ExtraPlasticL, BM_ExtraPlasticR, BM_Gate4, BM_Gate6, BM_LFlip, BM_LFlipU, BM_LPostRubber_p1, BM_LPostRubber_p2, BM_LPostRubber_p3, BM_LPost_p3, BM_LRampLever, BM_LRamp_DP, BM_LRamp_UP, BM_LSling_001, BM_LSling1, BM_LSling2, BM_LSlingArm, BM_Layer0, BM_Layer1, BM_Layer2, BM_Layer3, BM_Layer4, BM_Lgate, BM_PF, BM_Parts, BM_PartsAbove, BM_RFlip, BM_RFlipU, BM_RGate, BM_RPostRubber_p1, BM_RPostRubber_p2, BM_RPostRubber_p3, BM_RPost_p3, BM_RRamp_DP, BM_RRamp_UP, BM_RSling, BM_RSling1, BM_RSling2, BM_RSlingArm, BM_RailL, BM_RailR, BM_Sculpts, BM_UnderPF, BM_lockdownbar, BM_sw15, BM_sw16, BM_sw25, BM_sw26, BM_sw27, BM_sw28, BM_sw31, BM_sw35, BM_sw36, _
  BM_sw37, BM_sw38, BM_sw66, BM_sw67, BM_sw68, BM_sw84, BM_sw85, BM_sw86, BM_sw87, BM_sw88)
Dim BG_Lightmap: BG_Lightmap=Array(LM_FL_F17_Dome_Diverter, LM_FL_F17_Dome_DiverterU, LM_FL_F17_Dome_Gate6, LM_FL_F17_Dome_Layer1, LM_FL_F17_Dome_Layer2, LM_FL_F17_Dome_Layer3, LM_FL_F17_Dome_Layer4, LM_FL_F17_Dome_PF, LM_FL_F17_Dome_Parts, LM_FL_F17_Dome_PartsAbove, LM_FL_F17_Dome_RRamp_DP, LM_FL_F17_Dome_RRamp_UP, LM_FL_F18_BmpsBtm, LM_FL_F18_BmpsTop, LM_FL_F18_Bumper1_Ring, LM_FL_F18_Bumper1_Socket, LM_FL_F18_Bumper2_Ring, LM_FL_F18_Bumper2_Socket, LM_FL_F18_Bumper3_Ring, LM_FL_F18_Bumper3_Socket, LM_FL_F18_Layer2, LM_FL_F18_Layer3, LM_FL_F18_PF, LM_FL_F18_Parts, LM_FL_F18_PartsAbove, LM_FL_F18_Sculpts, LM_FL_F18_UnderPF, LM_FL_F18_sw66, LM_FL_F18_sw67, LM_FL_F18_sw68, LM_FL_F18_sw86, LM_FL_F18_sw87, LM_FL_F19_BmpsTop, LM_FL_F19_Bumper1_Ring, LM_FL_F19_ExtPlasticScrew2, LM_FL_F19_ExtPlasticScrew3, LM_FL_F19_Layer3, LM_FL_F19_Lgate, LM_FL_F19_PF, LM_FL_F19_Parts, LM_FL_F19_PartsAbove, LM_FL_F19_RPostRubber_p1, LM_FL_F19_RPostRubber_p2, LM_FL_F19_RPostRubber_p3, LM_FL_F19_RSling1, LM_FL_F19_Sculpts, _
  LM_FL_F19_UnderPF, LM_FL_F19_sw66, LM_FL_F19_sw67, LM_FL_F19_sw68, LM_FL_F19a_BmpsBtm, LM_FL_F19a_BmpsTop, LM_FL_F19a_Bumper1_Ring, LM_FL_F19a_Bumper1_Socket, LM_FL_F19a_Bumper2_Ring, LM_FL_F19a_Bumper2_Socket, LM_FL_F19a_Bumper3_Ring, LM_FL_F19a_Bumper3_Socket, LM_FL_F19a_ExtPlasticScrew3, LM_FL_F19a_LSling_001, LM_FL_F19a_Layer0, LM_FL_F19a_Layer2, LM_FL_F19a_Layer3, LM_FL_F19a_Layer4, LM_FL_F19a_PF, LM_FL_F19a_Parts, LM_FL_F19a_PartsAbove, LM_FL_F19a_RRamp_DP, LM_FL_F19a_RRamp_UP, LM_FL_F19a_Sculpts, LM_FL_F19a_UnderPF, LM_FL_F19a_sw66, LM_FL_F19a_sw67, LM_FL_F19a_sw68, LM_FL_F19a_sw86, LM_FL_F19a_sw87, LM_FL_F19a_sw88, LM_FL_F20_BmpsTop, LM_FL_F20_Bumper2_Socket, LM_FL_F20_Layer1, LM_FL_F20_Layer2, LM_FL_F20_Layer3, LM_FL_F20_Layer4, LM_FL_F20_PF, LM_FL_F20_Parts, LM_FL_F20_PartsAbove, LM_FL_F20_Sculpts, LM_FL_F20_UnderPF, LM_FL_F20_sw15, LM_FL_F20_sw16, LM_FL_F20_Dome_BmpsBtm, LM_FL_F20_Dome_BmpsTop, LM_FL_F20_Dome_Layer2, LM_FL_F20_Dome_Layer3, LM_FL_F20_Dome_Layer4, LM_FL_F20_Dome_PF, _
  LM_FL_F20_Dome_Parts, LM_FL_F20_Dome_PartsAbove, LM_FL_F20_Dome_Sculpts, LM_FL_F20_Dome_sw15, LM_FL_F20_Dome_sw16, LM_FL_F20_Dome_sw26, LM_FL_F21_BmpsTop, LM_FL_F21_Bumper1_Ring, LM_FL_F21_ExtPlasticScrew3, LM_FL_F21_LSling_001, LM_FL_F21_LSling1, LM_FL_F21_Layer0, LM_FL_F21_Layer3, LM_FL_F21_PF, LM_FL_F21_Parts, LM_FL_F21_PartsAbove, LM_FL_F21_RGate, LM_FL_F21_RPostRubber_p1, LM_FL_F21_RPostRubber_p2, LM_FL_F21_RPostRubber_p3, LM_FL_F21_RPost_p3, LM_FL_F21_Sculpts, LM_FL_F21_UnderPF, LM_FL_F21_sw86, LM_FL_F21_sw87, LM_FL_F21_sw88, LM_FL_F21a_Bumper1_Ring, LM_FL_F21a_ExtPlasticScrew1, LM_FL_F21a_ExtPlasticScrew3, LM_FL_F21a_ExtPlasticScrew4, LM_FL_F21a_ExtraPlasticR, LM_FL_F21a_LPost_p3, LM_FL_F21a_LRampLever, LM_FL_F21a_LRamp_UP, LM_FL_F21a_LSling_001, LM_FL_F21a_LSling1, LM_FL_F21a_LSling2, LM_FL_F21a_LSlingArm, LM_FL_F21a_Layer0, LM_FL_F21a_Layer1, LM_FL_F21a_Layer3, LM_FL_F21a_PF, LM_FL_F21a_Parts, LM_FL_F21a_PartsAbove, LM_FL_F21a_RPostRubber_p1, LM_FL_F21a_RPostRubber_p2, LM_FL_F21a_RPostRubber_p3, _
  LM_FL_F21a_RPost_p3, LM_FL_F21a_RSling, LM_FL_F21a_RSling2, LM_FL_F21a_RSlingArm, LM_FL_F21a_Sculpts, LM_FL_F21a_UnderPF, LM_FL_F21a_sw31, LM_FL_F21a_sw37, LM_FL_F21a_sw38, LM_FL_F21a_sw86, LM_FL_F21a_sw87, LM_FL_F21a_sw88, LM_FL_F22_BmpsBtm, LM_FL_F22_BmpsTop, LM_FL_F22_Bumper1_Ring, LM_FL_F22_Bumper1_Socket, LM_FL_F22_Bumper2_Ring, LM_FL_F22_Bumper2_Socket, LM_FL_F22_Bumper3_Ring, LM_FL_F22_Bumper3_Socket, LM_FL_F22_LPostRubber_p1, LM_FL_F22_LPostRubber_p2, LM_FL_F22_LPostRubber_p3, LM_FL_F22_Layer0, LM_FL_F22_Layer2, LM_FL_F22_Layer3, LM_FL_F22_PF, LM_FL_F22_Parts, LM_FL_F22_PartsAbove, LM_FL_F22_Sculpts, LM_FL_F22_UnderPF, LM_FL_F22_sw27, LM_FL_F22_sw66, LM_FL_F22_sw67, LM_FL_F22_sw68, LM_FL_F22_sw86, LM_FL_F22_sw87, LM_FL_F22_sw88, LM_FL_F23_BmpsBtm, LM_FL_F23_BmpsTop, LM_FL_F23_Bumper1_Ring, LM_FL_F23_Bumper1_Socket, LM_FL_F23_Bumper2_Ring, LM_FL_F23_Bumper2_Socket, LM_FL_F23_Bumper3_Ring, LM_FL_F23_Bumper3_Socket, LM_FL_F23_Layer0, LM_FL_F23_Layer1, LM_FL_F23_Layer2, LM_FL_F23_Layer3, LM_FL_F23_Layer4, _
  LM_FL_F23_Lgate, LM_FL_F23_PF, LM_FL_F23_Parts, LM_FL_F23_PartsAbove, LM_FL_F23_RGate, LM_FL_F23_RRamp_DP, LM_FL_F23_Sculpts, LM_FL_F23_UnderPF, LM_FL_F23_sw86, LM_FL_F23_sw87, LM_FL_F23_sw88, LM_FL_F24_BmpsTop, LM_FL_F24_Bumper1_Ring, LM_FL_F24_Layer3, LM_FL_F24_PF, LM_FL_F24_Parts, LM_FL_F24_PartsAbove, LM_FL_F24_RGate, LM_FL_F24_Sculpts, LM_FL_F24_UnderPF, LM_FL_F24_sw16, LM_FL_F24_sw66, LM_FL_F24_sw67, LM_FL_F24_sw68, LM_FL_F24_sw86, LM_GI_gib01_ExtPlasticScrew3, LM_GI_gib01_ExtPlasticSpacer3, LM_GI_gib01_ExtraPlasticL, LM_GI_gib01_LPostRubber_p1, LM_GI_gib01_LPostRubber_p2, LM_GI_gib01_LPostRubber_p3, LM_GI_gib01_LPost_p3, LM_GI_gib01_LSling_001, LM_GI_gib01_LSling1, LM_GI_gib01_LSling2, LM_GI_gib01_LSlingArm, LM_GI_gib01_PF, LM_GI_gib01_Parts, LM_GI_gib01_UnderPF, LM_GI_gib01_sw36, LM_GI_gib02_ExtPlasticScrew1, LM_GI_gib02_ExtraPlasticL, LM_GI_gib02_LFlip, LM_GI_gib02_LFlipU, LM_GI_gib02_LSling_001, LM_GI_gib02_LSling1, LM_GI_gib02_LSling2, LM_GI_gib02_LSlingArm, LM_GI_gib02_PF, LM_GI_gib02_Parts, _
  LM_GI_gib02_sw35, LM_GI_gib02_sw36, LM_GI_gib03_ExtPlasticSpacer2, LM_GI_gib03_ExtraPlasticL, LM_GI_gib03_ExtraPlasticR, LM_GI_gib03_LSling_001, LM_GI_gib03_LSlingArm, LM_GI_gib03_PF, LM_GI_gib03_Parts, LM_GI_gib03_RSling, LM_GI_gib03_RSling2, LM_GI_gib03_RSlingArm, LM_GI_gib03_sw35, LM_GI_gib03_sw37, LM_GI_gib04_ExtraPlasticL, LM_GI_gib04_LFlip, LM_GI_gib04_LFlipU, LM_GI_gib04_LSling_001, LM_GI_gib04_LSlingArm, LM_GI_gib04_PF, LM_GI_gib04_Parts, LM_GI_gib04_sw36, LM_GI_gib05_ExtPlasticSpacer4, LM_GI_gib05_ExtraPlasticR, LM_GI_gib05_Layer3, LM_GI_gib05_PF, LM_GI_gib05_Parts, LM_GI_gib05_PartsAbove, LM_GI_gib05_RPost_p3, LM_GI_gib05_RSling, LM_GI_gib05_RSling1, LM_GI_gib05_RSling2, LM_GI_gib05_RSlingArm, LM_GI_gib05_sw37, LM_GI_gib06_ExtraPlasticR, LM_GI_gib06_PF, LM_GI_gib06_Parts, LM_GI_gib06_RFlip, LM_GI_gib06_RFlipU, LM_GI_gib06_RSling, LM_GI_gib06_RSling1, LM_GI_gib06_RSling2, LM_GI_gib06_RSlingArm, LM_GI_gib06_sw37, LM_GI_gib08_ExtPlasticScrew2, LM_GI_gib08_ExtraPlasticL, LM_GI_gib08_PF, _
  LM_GI_gib08_Parts, LM_GI_gib08_RFlip, LM_GI_gib08_RFlipU, LM_GI_gib08_RSling, LM_GI_gib08_RSling2, LM_GI_gib08_RSlingArm, LM_GI_gim01_BmpsTop, LM_GI_gim01_Bumper1_Ring, LM_GI_gim01_ExtPlasticScrew4, LM_GI_gim01_ExtraPlasticR, LM_GI_gim01_LPost_p3, LM_GI_gim01_Layer0, LM_GI_gim01_Layer3, LM_GI_gim01_PF, LM_GI_gim01_Parts, LM_GI_gim01_PartsAbove, LM_GI_gim01_RPostRubber_p1, LM_GI_gim01_RPostRubber_p2, LM_GI_gim01_RPostRubber_p3, LM_GI_gim01_RPost_p3, LM_GI_gim01_RSling, LM_GI_gim01_Sculpts, LM_GI_gim01_UnderPF, LM_GI_gim01_sw66, LM_GI_gim01_sw67, LM_GI_gim01_sw68, LM_GI_gim01_sw86, LM_GI_gim01_sw87, LM_GI_gim01_sw88, LM_GI_gim04_BmpsBtm, LM_GI_gim04_BmpsTop, LM_GI_gim04_Bumper1_Ring, LM_GI_gim04_Bumper1_Socket, LM_GI_gim04_Bumper2_Ring, LM_GI_gim04_Bumper2_Socket, LM_GI_gim04_Layer3, LM_GI_gim04_PF, LM_GI_gim04_Parts, LM_GI_gim04_Sculpts, LM_GI_gim04_UnderPF, LM_GI_gim04_sw66, LM_GI_gim04_sw86, LM_GI_gim04_sw87, LM_GI_gim04_sw88, LM_GI_gim05_BmpsBtm, LM_GI_gim05_BmpsTop, LM_GI_gim05_Bumper1_Ring, _
  LM_GI_gim05_Bumper1_Socket, LM_GI_gim05_Bumper3_Ring, LM_GI_gim05_Bumper3_Socket, LM_GI_gim05_Layer2, LM_GI_gim05_Layer3, LM_GI_gim05_PF, LM_GI_gim05_Parts, LM_GI_gim05_UnderPF, LM_GI_gim05_sw86, LM_GI_gim05_sw87, LM_GI_gim05_sw88, LM_GI_git01_BmpsBtm, LM_GI_git01_BmpsTop, LM_GI_git01_Bumper1_Socket, LM_GI_git01_Bumper2_Ring, LM_GI_git01_Bumper2_Socket, LM_GI_git01_Bumper3_Ring, LM_GI_git01_Bumper3_Socket, LM_GI_git01_Diverter, LM_GI_git01_DiverterU, LM_GI_git01_Gate6, LM_GI_git01_Layer0, LM_GI_git01_Layer1, LM_GI_git01_Layer2, LM_GI_git01_Layer3, LM_GI_git01_Layer4, LM_GI_git01_PF, LM_GI_git01_Parts, LM_GI_git01_PartsAbove, LM_GI_git01_RGate, LM_GI_git01_RRamp_DP, LM_GI_git01_RRamp_UP, LM_GI_git01_Sculpts, LM_GI_git01_UnderPF, LM_GI_git01_sw16, LM_GI_git01_sw25, LM_GI_git01_sw26, LM_GI_git01_sw27, LM_GI_git01_sw84, LM_GI_git01_sw85, LM_L_L61_Bumper2_Socket, LM_L_L61_Layer2, LM_L_L61_Layer3, LM_L_L61_PF, LM_L_L61_Parts, LM_L_L61_PartsAbove, LM_L_L61_UnderPF, LM_L_L62_Layer2, LM_L_L62_Layer3, LM_L_L62_PF, _
  LM_L_L62_Parts, LM_L_L62_PartsAbove, LM_L_L62_UnderPF, LM_L_L63_DiverterU, LM_L_L63_Gate6, LM_L_L63_Layer2, LM_L_L63_Layer3, LM_L_L63_PF, LM_L_L63_Parts, LM_L_L63_PartsAbove, LM_L_L63_UnderPF, LM_L_L68_BmpsBtm, LM_L_L68_BmpsTop, LM_L_L68_Bumper2_Ring, LM_L_L68_Bumper3_Ring, LM_L_L68_Layer3, LM_L_L68_PF, LM_L_L68_Parts, LM_L_l16_Layer0, LM_L_l16_Layer1, LM_L_l16_Layer2, LM_L_l16_Layer3, LM_L_l16_Layer4, LM_L_l16_PF, LM_L_l16_Parts, LM_L_l16_RRamp_DP, LM_L_l16_RRamp_UP, LM_L_l16_UnderPF, LM_L_l17_Layer3, LM_L_l17_PF, LM_L_l17_Parts, LM_L_l18_Layer0, LM_L_l18_Layer3, LM_L_l18_Layer4, LM_L_l18_PF, LM_L_l18_Parts, LM_L_l18_Sculpts, LM_L_l18_UnderPF, LM_L_l21_Layer2, LM_L_l21_Layer4, LM_L_l21_Parts, LM_L_l21_UnderPF, LM_L_l22_Layer2, LM_L_l22_Layer3, LM_L_l22_Layer4, LM_L_l22_Parts, LM_L_l22_UnderPF, LM_L_l23_Layer3, LM_L_l23_PF, LM_L_l23_UnderPF, LM_L_l24_Layer0, LM_L_l24_PF, LM_L_l24_Parts, LM_L_l24_Sculpts, LM_L_l24_UnderPF, LM_L_l25_BmpsTop, LM_L_l25_PF, LM_L_l25_Parts, LM_L_l25_UnderPF, LM_L_l25_sw88, _
  LM_L_l26_Layer0, LM_L_l26_Layer4, LM_L_l26_PF, LM_L_l26_Parts, LM_L_l26_Sculpts, LM_L_l26_UnderPF, LM_L_l27_Layer0, LM_L_l27_Layer3, LM_L_l27_Layer4, LM_L_l27_PF, LM_L_l27_Parts, LM_L_l27_RRamp_UP, LM_L_l27_UnderPF, LM_L_l28_BmpsTop, LM_L_l28_Layer0, LM_L_l28_Layer4, LM_L_l28_PF, LM_L_l28_Parts, LM_L_l28_Sculpts, LM_L_l28_UnderPF, LM_L_l31_Layer0, LM_L_l31_Layer3, LM_L_l31_PF, LM_L_l31_Parts, LM_L_l31_PartsAbove, LM_L_l31_UnderPF, LM_L_l32_Layer0, LM_L_l32_Layer3, LM_L_l32_PF, LM_L_l32_Parts, LM_L_l32_PartsAbove, LM_L_l32_UnderPF, LM_L_l33_Layer0, LM_L_l33_Layer3, LM_L_l33_Layer4, LM_L_l33_PF, LM_L_l33_Parts, LM_L_l33_PartsAbove, LM_L_l33_Sculpts, LM_L_l33_UnderPF, LM_L_l34_Layer0, LM_L_l34_Layer3, LM_L_l34_PF, LM_L_l34_Parts, LM_L_l34_PartsAbove, LM_L_l34_Sculpts, LM_L_l34_UnderPF, LM_L_l35_BmpsTop, LM_L_l35_Layer0, LM_L_l35_Layer3, LM_L_l35_Layer4, LM_L_l35_PF, LM_L_l35_Parts, LM_L_l35_Sculpts, LM_L_l35_UnderPF, LM_L_l36_PF, LM_L_l36_UnderPF, LM_L_l37_LPostRubber_p2, LM_L_l37_LPostRubber_p3, _
  LM_L_l37_LPost_p3, LM_L_l37_PF, LM_L_l37_Parts, LM_L_l37_Sculpts, LM_L_l37_UnderPF, LM_L_l38_ExtraPlasticL, LM_L_l38_LPostRubber_p1, LM_L_l38_LPostRubber_p2, LM_L_l38_LPostRubber_p3, LM_L_l38_LSling_001, LM_L_l38_PF, LM_L_l38_Parts, LM_L_l38_Sculpts, LM_L_l38_UnderPF, LM_L_l41_Layer3, LM_L_l41_PF, LM_L_l41_Parts, LM_L_l41_PartsAbove, LM_L_l41_RPostRubber_p2, LM_L_l41_RPostRubber_p3, LM_L_l41_RSling, LM_L_l41_Sculpts, LM_L_l41_UnderPF, LM_L_l42_Layer3, LM_L_l42_PF, LM_L_l42_Parts, LM_L_l42_RPost_p3, LM_L_l42_UnderPF, LM_L_l43_LFlipU, LM_L_l43_PF, LM_L_l43_UnderPF, LM_L_l44_LFlipU, LM_L_l44_PF, LM_L_l44_UnderPF, LM_L_l45_PF, LM_L_l45_UnderPF, LM_L_l46_PF, LM_L_l46_RFlipU, LM_L_l46_UnderPF, LM_L_l47_PF, LM_L_l47_RFlipU, LM_L_l47_UnderPF, LM_L_l48_Layer3, LM_L_l48_PF, LM_L_l48_Parts, LM_L_l48_UnderPF, LM_L_l51_Layer2, LM_L_l51_Layer4, LM_L_l51_Parts, LM_L_l51_PartsAbove, LM_L_l51_Sculpts, LM_L_l51_UnderPF, LM_L_l52_Layer2, LM_L_l52_Layer4, LM_L_l52_Parts, LM_L_l52_Sculpts, LM_L_l52_UnderPF, LM_L_l54_PF, _
  LM_L_l54_UnderPF, LM_L_l55_LFlip, LM_L_l55_PF, LM_L_l55_Parts, LM_L_l55_RFlip, LM_L_l55_UnderPF, LM_L_l56_PF, LM_L_l56_Parts, LM_L_l57_PF, LM_L_l57_Parts, LM_L_l57_Sculpts, LM_L_l57_UnderPF, LM_L_l58_Bumper2_Socket, LM_L_l58_Diverter, LM_L_l58_DiverterU, LM_L_l58_Layer2, LM_L_l58_Layer3, LM_L_l58_PF, LM_L_l58_Parts, LM_L_l58_PartsAbove, LM_L_l58_UnderPF, LM_L_l64_DiverterU, LM_L_l64_Layer2, LM_L_l64_Layer3, LM_L_l64_Layer4, LM_L_l64_PF, LM_L_l64_Parts, LM_L_l64_PartsAbove, LM_L_l64_UnderPF, LM_L_l65_Layer2, LM_L_l65_Layer3, LM_L_l65_Layer4, LM_L_l65_PF, LM_L_l65_Parts, LM_L_l65_PartsAbove, LM_L_l65_UnderPF, LM_L_l66_Layer2, LM_L_l66_PF, LM_L_l66_Parts, LM_L_l66_PartsAbove, LM_L_l66_UnderPF, LM_L_l67_Layer3, LM_L_l67_PF, LM_L_l67_UnderPF, LM_L_l71_PF, LM_L_l71_UnderPF, LM_L_l72_Layer3, LM_L_l72_PF, LM_L_l72_UnderPF, LM_L_l73_PF, LM_L_l73_Parts, LM_L_l73_Sculpts, LM_L_l73_UnderPF, LM_L_l73_sw66, LM_L_l73_sw67, LM_L_l73_sw68, LM_L_l74_PF, LM_L_l74_Parts, LM_L_l74_Sculpts, LM_L_l74_UnderPF, LM_L_l74_sw66, _
  LM_L_l74_sw67, LM_L_l74_sw68, LM_L_l75_PF, LM_L_l75_Parts, LM_L_l75_Sculpts, LM_L_l75_UnderPF, LM_L_l75_sw67, LM_L_l75_sw68, LM_L_l76_PF, LM_L_l76_Parts, LM_L_l76_UnderPF, LM_L_l76_sw86, LM_L_l76_sw87, LM_L_l77_PF, LM_L_l77_Parts, LM_L_l77_UnderPF, LM_L_l77_sw86, LM_L_l77_sw87, LM_L_l77_sw88, LM_L_l78_PF, LM_L_l78_Parts, LM_L_l78_UnderPF, LM_L_l78_sw87, LM_L_l78_sw88, LM_L_l81_Layer0, LM_L_l81_PF, LM_L_l81_Parts, LM_L_l81_Sculpts, LM_L_l82_PF, LM_L_l82_UnderPF, LM_L_l83_BmpsTop, LM_L_l83_Bumper2_Ring, LM_L_l83_Layer2, LM_L_l83_Layer3, LM_L_l83_PF, LM_L_l83_Parts, LM_L_l83_PartsAbove, LM_L_l83_UnderPF, LM_L_l84_BmpsBtm, LM_L_l84_BmpsTop, LM_L_l84_Bumper2_Ring, LM_L_l84_Bumper2_Socket, LM_L_l84_Layer2, LM_L_l84_Layer3, LM_L_l84_Layer4, LM_L_l84_PF, LM_L_l84_Parts, LM_L_l84_PartsAbove, LM_L_l84_UnderPF, LM_L_l85_BmpsTop, LM_L_l85_Bumper1_Ring, LM_L_l85_Layer3, LM_L_l85_PF, LM_L_l85_Parts, LM_L_l85_UnderPF, LM_L_l86_PF, LM_L_l86_Parts, LM_L_l86_UnderPF, LM_L_ml1_PF, LM_L_ml1_Parts, LM_L_ml10_PF, LM_L_ml11_PF, _
  LM_L_ml11_Parts, LM_L_ml12_PF, LM_L_ml13_PF, LM_L_ml2_PF, LM_L_ml3_PF, LM_L_ml4_PF, LM_L_ml5_PF, LM_L_ml6_PF, LM_L_ml7_PF, LM_L_ml8_PF, LM_L_ml9_PF, LM_L_Bmp_BmpsBtm, LM_L_Bmp_BmpsTop, LM_L_Bmp_Bumper1_Ring, LM_L_Bmp_Bumper1_Socket, LM_L_Bmp_Bumper2_Ring, LM_L_Bmp_Bumper2_Socket, LM_L_Bmp_Bumper3_Ring, LM_L_Bmp_Bumper3_Socket, LM_L_Bmp_Layer0, LM_L_Bmp_Layer2, LM_L_Bmp_Layer3, LM_L_Bmp_Layer4, LM_L_Bmp_PF, LM_L_Bmp_Parts, LM_L_Bmp_RRamp_DP, LM_L_Bmp_RRamp_UP, LM_L_Bmp_UnderPF)
Dim BG_All: BG_All=Array(BM_BmpsBtm, BM_BmpsTop, BM_Bumper1_Ring, BM_Bumper1_Socket, BM_Bumper2_Ring, BM_Bumper2_Socket, BM_Bumper3_Ring, BM_Bumper3_Socket, BM_Diverter, BM_DiverterU, BM_ExtPlasticScrew1, BM_ExtPlasticScrew2, BM_ExtPlasticScrew3, BM_ExtPlasticScrew4, BM_ExtPlasticSpacer1, BM_ExtPlasticSpacer2, BM_ExtPlasticSpacer3, BM_ExtPlasticSpacer4, BM_ExtraPlasticL, BM_ExtraPlasticR, BM_Gate4, BM_Gate6, BM_LFlip, BM_LFlipU, BM_LPostRubber_p1, BM_LPostRubber_p2, BM_LPostRubber_p3, BM_LPost_p3, BM_LRampLever, BM_LRamp_DP, BM_LRamp_UP, BM_LSling_001, BM_LSling1, BM_LSling2, BM_LSlingArm, BM_Layer0, BM_Layer1, BM_Layer2, BM_Layer3, BM_Layer4, BM_Lgate, BM_PF, BM_Parts, BM_PartsAbove, BM_RFlip, BM_RFlipU, BM_RGate, BM_RPostRubber_p1, BM_RPostRubber_p2, BM_RPostRubber_p3, BM_RPost_p3, BM_RRamp_DP, BM_RRamp_UP, BM_RSling, BM_RSling1, BM_RSling2, BM_RSlingArm, BM_RailL, BM_RailR, BM_Sculpts, BM_UnderPF, BM_lockdownbar, BM_sw15, BM_sw16, BM_sw25, BM_sw26, BM_sw27, BM_sw28, BM_sw31, BM_sw35, BM_sw36, BM_sw37, _
  BM_sw38, BM_sw66, BM_sw67, BM_sw68, BM_sw84, BM_sw85, BM_sw86, BM_sw87, BM_sw88, LM_FL_F17_Dome_Diverter, LM_FL_F17_Dome_DiverterU, LM_FL_F17_Dome_Gate6, LM_FL_F17_Dome_Layer1, LM_FL_F17_Dome_Layer2, LM_FL_F17_Dome_Layer3, LM_FL_F17_Dome_Layer4, LM_FL_F17_Dome_PF, LM_FL_F17_Dome_Parts, LM_FL_F17_Dome_PartsAbove, LM_FL_F17_Dome_RRamp_DP, LM_FL_F17_Dome_RRamp_UP, LM_FL_F18_BmpsBtm, LM_FL_F18_BmpsTop, LM_FL_F18_Bumper1_Ring, LM_FL_F18_Bumper1_Socket, LM_FL_F18_Bumper2_Ring, LM_FL_F18_Bumper2_Socket, LM_FL_F18_Bumper3_Ring, LM_FL_F18_Bumper3_Socket, LM_FL_F18_Layer2, LM_FL_F18_Layer3, LM_FL_F18_PF, LM_FL_F18_Parts, LM_FL_F18_PartsAbove, LM_FL_F18_Sculpts, LM_FL_F18_UnderPF, LM_FL_F18_sw66, LM_FL_F18_sw67, LM_FL_F18_sw68, LM_FL_F18_sw86, LM_FL_F18_sw87, LM_FL_F19_BmpsTop, LM_FL_F19_Bumper1_Ring, LM_FL_F19_ExtPlasticScrew2, LM_FL_F19_ExtPlasticScrew3, LM_FL_F19_Layer3, LM_FL_F19_Lgate, LM_FL_F19_PF, LM_FL_F19_Parts, LM_FL_F19_PartsAbove, LM_FL_F19_RPostRubber_p1, LM_FL_F19_RPostRubber_p2, LM_FL_F19_RPostRubber_p3, _
  LM_FL_F19_RSling1, LM_FL_F19_Sculpts, LM_FL_F19_UnderPF, LM_FL_F19_sw66, LM_FL_F19_sw67, LM_FL_F19_sw68, LM_FL_F19a_BmpsBtm, LM_FL_F19a_BmpsTop, LM_FL_F19a_Bumper1_Ring, LM_FL_F19a_Bumper1_Socket, LM_FL_F19a_Bumper2_Ring, LM_FL_F19a_Bumper2_Socket, LM_FL_F19a_Bumper3_Ring, LM_FL_F19a_Bumper3_Socket, LM_FL_F19a_ExtPlasticScrew3, LM_FL_F19a_LSling_001, LM_FL_F19a_Layer0, LM_FL_F19a_Layer2, LM_FL_F19a_Layer3, LM_FL_F19a_Layer4, LM_FL_F19a_PF, LM_FL_F19a_Parts, LM_FL_F19a_PartsAbove, LM_FL_F19a_RRamp_DP, LM_FL_F19a_RRamp_UP, LM_FL_F19a_Sculpts, LM_FL_F19a_UnderPF, LM_FL_F19a_sw66, LM_FL_F19a_sw67, LM_FL_F19a_sw68, LM_FL_F19a_sw86, LM_FL_F19a_sw87, LM_FL_F19a_sw88, LM_FL_F20_BmpsTop, LM_FL_F20_Bumper2_Socket, LM_FL_F20_Layer1, LM_FL_F20_Layer2, LM_FL_F20_Layer3, LM_FL_F20_Layer4, LM_FL_F20_PF, LM_FL_F20_Parts, LM_FL_F20_PartsAbove, LM_FL_F20_Sculpts, LM_FL_F20_UnderPF, LM_FL_F20_sw15, LM_FL_F20_sw16, LM_FL_F20_Dome_BmpsBtm, LM_FL_F20_Dome_BmpsTop, LM_FL_F20_Dome_Layer2, LM_FL_F20_Dome_Layer3, _
  LM_FL_F20_Dome_Layer4, LM_FL_F20_Dome_PF, LM_FL_F20_Dome_Parts, LM_FL_F20_Dome_PartsAbove, LM_FL_F20_Dome_Sculpts, LM_FL_F20_Dome_sw15, LM_FL_F20_Dome_sw16, LM_FL_F20_Dome_sw26, LM_FL_F21_BmpsTop, LM_FL_F21_Bumper1_Ring, LM_FL_F21_ExtPlasticScrew3, LM_FL_F21_LSling_001, LM_FL_F21_LSling1, LM_FL_F21_Layer0, LM_FL_F21_Layer3, LM_FL_F21_PF, LM_FL_F21_Parts, LM_FL_F21_PartsAbove, LM_FL_F21_RGate, LM_FL_F21_RPostRubber_p1, LM_FL_F21_RPostRubber_p2, LM_FL_F21_RPostRubber_p3, LM_FL_F21_RPost_p3, LM_FL_F21_Sculpts, LM_FL_F21_UnderPF, LM_FL_F21_sw86, LM_FL_F21_sw87, LM_FL_F21_sw88, LM_FL_F21a_Bumper1_Ring, LM_FL_F21a_ExtPlasticScrew1, LM_FL_F21a_ExtPlasticScrew3, LM_FL_F21a_ExtPlasticScrew4, LM_FL_F21a_ExtraPlasticR, LM_FL_F21a_LPost_p3, LM_FL_F21a_LRampLever, LM_FL_F21a_LRamp_UP, LM_FL_F21a_LSling_001, LM_FL_F21a_LSling1, LM_FL_F21a_LSling2, LM_FL_F21a_LSlingArm, LM_FL_F21a_Layer0, LM_FL_F21a_Layer1, LM_FL_F21a_Layer3, LM_FL_F21a_PF, LM_FL_F21a_Parts, LM_FL_F21a_PartsAbove, LM_FL_F21a_RPostRubber_p1, _
  LM_FL_F21a_RPostRubber_p2, LM_FL_F21a_RPostRubber_p3, LM_FL_F21a_RPost_p3, LM_FL_F21a_RSling, LM_FL_F21a_RSling2, LM_FL_F21a_RSlingArm, LM_FL_F21a_Sculpts, LM_FL_F21a_UnderPF, LM_FL_F21a_sw31, LM_FL_F21a_sw37, LM_FL_F21a_sw38, LM_FL_F21a_sw86, LM_FL_F21a_sw87, LM_FL_F21a_sw88, LM_FL_F22_BmpsBtm, LM_FL_F22_BmpsTop, LM_FL_F22_Bumper1_Ring, LM_FL_F22_Bumper1_Socket, LM_FL_F22_Bumper2_Ring, LM_FL_F22_Bumper2_Socket, LM_FL_F22_Bumper3_Ring, LM_FL_F22_Bumper3_Socket, LM_FL_F22_LPostRubber_p1, LM_FL_F22_LPostRubber_p2, LM_FL_F22_LPostRubber_p3, LM_FL_F22_Layer0, LM_FL_F22_Layer2, LM_FL_F22_Layer3, LM_FL_F22_PF, LM_FL_F22_Parts, LM_FL_F22_PartsAbove, LM_FL_F22_Sculpts, LM_FL_F22_UnderPF, LM_FL_F22_sw27, LM_FL_F22_sw66, LM_FL_F22_sw67, LM_FL_F22_sw68, LM_FL_F22_sw86, LM_FL_F22_sw87, LM_FL_F22_sw88, LM_FL_F23_BmpsBtm, LM_FL_F23_BmpsTop, LM_FL_F23_Bumper1_Ring, LM_FL_F23_Bumper1_Socket, LM_FL_F23_Bumper2_Ring, LM_FL_F23_Bumper2_Socket, LM_FL_F23_Bumper3_Ring, LM_FL_F23_Bumper3_Socket, LM_FL_F23_Layer0, LM_FL_F23_Layer1, _
  LM_FL_F23_Layer2, LM_FL_F23_Layer3, LM_FL_F23_Layer4, LM_FL_F23_Lgate, LM_FL_F23_PF, LM_FL_F23_Parts, LM_FL_F23_PartsAbove, LM_FL_F23_RGate, LM_FL_F23_RRamp_DP, LM_FL_F23_Sculpts, LM_FL_F23_UnderPF, LM_FL_F23_sw86, LM_FL_F23_sw87, LM_FL_F23_sw88, LM_FL_F24_BmpsTop, LM_FL_F24_Bumper1_Ring, LM_FL_F24_Layer3, LM_FL_F24_PF, LM_FL_F24_Parts, LM_FL_F24_PartsAbove, LM_FL_F24_RGate, LM_FL_F24_Sculpts, LM_FL_F24_UnderPF, LM_FL_F24_sw16, LM_FL_F24_sw66, LM_FL_F24_sw67, LM_FL_F24_sw68, LM_FL_F24_sw86, LM_GI_gib01_ExtPlasticScrew3, LM_GI_gib01_ExtPlasticSpacer3, LM_GI_gib01_ExtraPlasticL, LM_GI_gib01_LPostRubber_p1, LM_GI_gib01_LPostRubber_p2, LM_GI_gib01_LPostRubber_p3, LM_GI_gib01_LPost_p3, LM_GI_gib01_LSling_001, LM_GI_gib01_LSling1, LM_GI_gib01_LSling2, LM_GI_gib01_LSlingArm, LM_GI_gib01_PF, LM_GI_gib01_Parts, LM_GI_gib01_UnderPF, LM_GI_gib01_sw36, LM_GI_gib02_ExtPlasticScrew1, LM_GI_gib02_ExtraPlasticL, LM_GI_gib02_LFlip, LM_GI_gib02_LFlipU, LM_GI_gib02_LSling_001, LM_GI_gib02_LSling1, LM_GI_gib02_LSling2, _
  LM_GI_gib02_LSlingArm, LM_GI_gib02_PF, LM_GI_gib02_Parts, LM_GI_gib02_sw35, LM_GI_gib02_sw36, LM_GI_gib03_ExtPlasticSpacer2, LM_GI_gib03_ExtraPlasticL, LM_GI_gib03_ExtraPlasticR, LM_GI_gib03_LSling_001, LM_GI_gib03_LSlingArm, LM_GI_gib03_PF, LM_GI_gib03_Parts, LM_GI_gib03_RSling, LM_GI_gib03_RSling2, LM_GI_gib03_RSlingArm, LM_GI_gib03_sw35, LM_GI_gib03_sw37, LM_GI_gib04_ExtraPlasticL, LM_GI_gib04_LFlip, LM_GI_gib04_LFlipU, LM_GI_gib04_LSling_001, LM_GI_gib04_LSlingArm, LM_GI_gib04_PF, LM_GI_gib04_Parts, LM_GI_gib04_sw36, LM_GI_gib05_ExtPlasticSpacer4, LM_GI_gib05_ExtraPlasticR, LM_GI_gib05_Layer3, LM_GI_gib05_PF, LM_GI_gib05_Parts, LM_GI_gib05_PartsAbove, LM_GI_gib05_RPost_p3, LM_GI_gib05_RSling, LM_GI_gib05_RSling1, LM_GI_gib05_RSling2, LM_GI_gib05_RSlingArm, LM_GI_gib05_sw37, LM_GI_gib06_ExtraPlasticR, LM_GI_gib06_PF, LM_GI_gib06_Parts, LM_GI_gib06_RFlip, LM_GI_gib06_RFlipU, LM_GI_gib06_RSling, LM_GI_gib06_RSling1, LM_GI_gib06_RSling2, LM_GI_gib06_RSlingArm, LM_GI_gib06_sw37, LM_GI_gib08_ExtPlasticScrew2, _
  LM_GI_gib08_ExtraPlasticL, LM_GI_gib08_PF, LM_GI_gib08_Parts, LM_GI_gib08_RFlip, LM_GI_gib08_RFlipU, LM_GI_gib08_RSling, LM_GI_gib08_RSling2, LM_GI_gib08_RSlingArm, LM_GI_gim01_BmpsTop, LM_GI_gim01_Bumper1_Ring, LM_GI_gim01_ExtPlasticScrew4, LM_GI_gim01_ExtraPlasticR, LM_GI_gim01_LPost_p3, LM_GI_gim01_Layer0, LM_GI_gim01_Layer3, LM_GI_gim01_PF, LM_GI_gim01_Parts, LM_GI_gim01_PartsAbove, LM_GI_gim01_RPostRubber_p1, LM_GI_gim01_RPostRubber_p2, LM_GI_gim01_RPostRubber_p3, LM_GI_gim01_RPost_p3, LM_GI_gim01_RSling, LM_GI_gim01_Sculpts, LM_GI_gim01_UnderPF, LM_GI_gim01_sw66, LM_GI_gim01_sw67, LM_GI_gim01_sw68, LM_GI_gim01_sw86, LM_GI_gim01_sw87, LM_GI_gim01_sw88, LM_GI_gim04_BmpsBtm, LM_GI_gim04_BmpsTop, LM_GI_gim04_Bumper1_Ring, LM_GI_gim04_Bumper1_Socket, LM_GI_gim04_Bumper2_Ring, LM_GI_gim04_Bumper2_Socket, LM_GI_gim04_Layer3, LM_GI_gim04_PF, LM_GI_gim04_Parts, LM_GI_gim04_Sculpts, LM_GI_gim04_UnderPF, LM_GI_gim04_sw66, LM_GI_gim04_sw86, LM_GI_gim04_sw87, LM_GI_gim04_sw88, LM_GI_gim05_BmpsBtm, _
  LM_GI_gim05_BmpsTop, LM_GI_gim05_Bumper1_Ring, LM_GI_gim05_Bumper1_Socket, LM_GI_gim05_Bumper3_Ring, LM_GI_gim05_Bumper3_Socket, LM_GI_gim05_Layer2, LM_GI_gim05_Layer3, LM_GI_gim05_PF, LM_GI_gim05_Parts, LM_GI_gim05_UnderPF, LM_GI_gim05_sw86, LM_GI_gim05_sw87, LM_GI_gim05_sw88, LM_GI_git01_BmpsBtm, LM_GI_git01_BmpsTop, LM_GI_git01_Bumper1_Socket, LM_GI_git01_Bumper2_Ring, LM_GI_git01_Bumper2_Socket, LM_GI_git01_Bumper3_Ring, LM_GI_git01_Bumper3_Socket, LM_GI_git01_Diverter, LM_GI_git01_DiverterU, LM_GI_git01_Gate6, LM_GI_git01_Layer0, LM_GI_git01_Layer1, LM_GI_git01_Layer2, LM_GI_git01_Layer3, LM_GI_git01_Layer4, LM_GI_git01_PF, LM_GI_git01_Parts, LM_GI_git01_PartsAbove, LM_GI_git01_RGate, LM_GI_git01_RRamp_DP, LM_GI_git01_RRamp_UP, LM_GI_git01_Sculpts, LM_GI_git01_UnderPF, LM_GI_git01_sw16, LM_GI_git01_sw25, LM_GI_git01_sw26, LM_GI_git01_sw27, LM_GI_git01_sw84, LM_GI_git01_sw85, LM_L_L61_Bumper2_Socket, LM_L_L61_Layer2, LM_L_L61_Layer3, LM_L_L61_PF, LM_L_L61_Parts, LM_L_L61_PartsAbove, LM_L_L61_UnderPF, _
  LM_L_L62_Layer2, LM_L_L62_Layer3, LM_L_L62_PF, LM_L_L62_Parts, LM_L_L62_PartsAbove, LM_L_L62_UnderPF, LM_L_L63_DiverterU, LM_L_L63_Gate6, LM_L_L63_Layer2, LM_L_L63_Layer3, LM_L_L63_PF, LM_L_L63_Parts, LM_L_L63_PartsAbove, LM_L_L63_UnderPF, LM_L_L68_BmpsBtm, LM_L_L68_BmpsTop, LM_L_L68_Bumper2_Ring, LM_L_L68_Bumper3_Ring, LM_L_L68_Layer3, LM_L_L68_PF, LM_L_L68_Parts, LM_L_l16_Layer0, LM_L_l16_Layer1, LM_L_l16_Layer2, LM_L_l16_Layer3, LM_L_l16_Layer4, LM_L_l16_PF, LM_L_l16_Parts, LM_L_l16_RRamp_DP, LM_L_l16_RRamp_UP, LM_L_l16_UnderPF, LM_L_l17_Layer3, LM_L_l17_PF, LM_L_l17_Parts, LM_L_l18_Layer0, LM_L_l18_Layer3, LM_L_l18_Layer4, LM_L_l18_PF, LM_L_l18_Parts, LM_L_l18_Sculpts, LM_L_l18_UnderPF, LM_L_l21_Layer2, LM_L_l21_Layer4, LM_L_l21_Parts, LM_L_l21_UnderPF, LM_L_l22_Layer2, LM_L_l22_Layer3, LM_L_l22_Layer4, LM_L_l22_Parts, LM_L_l22_UnderPF, LM_L_l23_Layer3, LM_L_l23_PF, LM_L_l23_UnderPF, LM_L_l24_Layer0, LM_L_l24_PF, LM_L_l24_Parts, LM_L_l24_Sculpts, LM_L_l24_UnderPF, LM_L_l25_BmpsTop, LM_L_l25_PF, _
  LM_L_l25_Parts, LM_L_l25_UnderPF, LM_L_l25_sw88, LM_L_l26_Layer0, LM_L_l26_Layer4, LM_L_l26_PF, LM_L_l26_Parts, LM_L_l26_Sculpts, LM_L_l26_UnderPF, LM_L_l27_Layer0, LM_L_l27_Layer3, LM_L_l27_Layer4, LM_L_l27_PF, LM_L_l27_Parts, LM_L_l27_RRamp_UP, LM_L_l27_UnderPF, LM_L_l28_BmpsTop, LM_L_l28_Layer0, LM_L_l28_Layer4, LM_L_l28_PF, LM_L_l28_Parts, LM_L_l28_Sculpts, LM_L_l28_UnderPF, LM_L_l31_Layer0, LM_L_l31_Layer3, LM_L_l31_PF, LM_L_l31_Parts, LM_L_l31_PartsAbove, LM_L_l31_UnderPF, LM_L_l32_Layer0, LM_L_l32_Layer3, LM_L_l32_PF, LM_L_l32_Parts, LM_L_l32_PartsAbove, LM_L_l32_UnderPF, LM_L_l33_Layer0, LM_L_l33_Layer3, LM_L_l33_Layer4, LM_L_l33_PF, LM_L_l33_Parts, LM_L_l33_PartsAbove, LM_L_l33_Sculpts, LM_L_l33_UnderPF, LM_L_l34_Layer0, LM_L_l34_Layer3, LM_L_l34_PF, LM_L_l34_Parts, LM_L_l34_PartsAbove, LM_L_l34_Sculpts, LM_L_l34_UnderPF, LM_L_l35_BmpsTop, LM_L_l35_Layer0, LM_L_l35_Layer3, LM_L_l35_Layer4, LM_L_l35_PF, LM_L_l35_Parts, LM_L_l35_Sculpts, LM_L_l35_UnderPF, LM_L_l36_PF, LM_L_l36_UnderPF, _
  LM_L_l37_LPostRubber_p2, LM_L_l37_LPostRubber_p3, LM_L_l37_LPost_p3, LM_L_l37_PF, LM_L_l37_Parts, LM_L_l37_Sculpts, LM_L_l37_UnderPF, LM_L_l38_ExtraPlasticL, LM_L_l38_LPostRubber_p1, LM_L_l38_LPostRubber_p2, LM_L_l38_LPostRubber_p3, LM_L_l38_LSling_001, LM_L_l38_PF, LM_L_l38_Parts, LM_L_l38_Sculpts, LM_L_l38_UnderPF, LM_L_l41_Layer3, LM_L_l41_PF, LM_L_l41_Parts, LM_L_l41_PartsAbove, LM_L_l41_RPostRubber_p2, LM_L_l41_RPostRubber_p3, LM_L_l41_RSling, LM_L_l41_Sculpts, LM_L_l41_UnderPF, LM_L_l42_Layer3, LM_L_l42_PF, LM_L_l42_Parts, LM_L_l42_RPost_p3, LM_L_l42_UnderPF, LM_L_l43_LFlipU, LM_L_l43_PF, LM_L_l43_UnderPF, LM_L_l44_LFlipU, LM_L_l44_PF, LM_L_l44_UnderPF, LM_L_l45_PF, LM_L_l45_UnderPF, LM_L_l46_PF, LM_L_l46_RFlipU, LM_L_l46_UnderPF, LM_L_l47_PF, LM_L_l47_RFlipU, LM_L_l47_UnderPF, LM_L_l48_Layer3, LM_L_l48_PF, LM_L_l48_Parts, LM_L_l48_UnderPF, LM_L_l51_Layer2, LM_L_l51_Layer4, LM_L_l51_Parts, LM_L_l51_PartsAbove, LM_L_l51_Sculpts, LM_L_l51_UnderPF, LM_L_l52_Layer2, LM_L_l52_Layer4, LM_L_l52_Parts, _
  LM_L_l52_Sculpts, LM_L_l52_UnderPF, LM_L_l54_PF, LM_L_l54_UnderPF, LM_L_l55_LFlip, LM_L_l55_PF, LM_L_l55_Parts, LM_L_l55_RFlip, LM_L_l55_UnderPF, LM_L_l56_PF, LM_L_l56_Parts, LM_L_l57_PF, LM_L_l57_Parts, LM_L_l57_Sculpts, LM_L_l57_UnderPF, LM_L_l58_Bumper2_Socket, LM_L_l58_Diverter, LM_L_l58_DiverterU, LM_L_l58_Layer2, LM_L_l58_Layer3, LM_L_l58_PF, LM_L_l58_Parts, LM_L_l58_PartsAbove, LM_L_l58_UnderPF, LM_L_l64_DiverterU, LM_L_l64_Layer2, LM_L_l64_Layer3, LM_L_l64_Layer4, LM_L_l64_PF, LM_L_l64_Parts, LM_L_l64_PartsAbove, LM_L_l64_UnderPF, LM_L_l65_Layer2, LM_L_l65_Layer3, LM_L_l65_Layer4, LM_L_l65_PF, LM_L_l65_Parts, LM_L_l65_PartsAbove, LM_L_l65_UnderPF, LM_L_l66_Layer2, LM_L_l66_PF, LM_L_l66_Parts, LM_L_l66_PartsAbove, LM_L_l66_UnderPF, LM_L_l67_Layer3, LM_L_l67_PF, LM_L_l67_UnderPF, LM_L_l71_PF, LM_L_l71_UnderPF, LM_L_l72_Layer3, LM_L_l72_PF, LM_L_l72_UnderPF, LM_L_l73_PF, LM_L_l73_Parts, LM_L_l73_Sculpts, LM_L_l73_UnderPF, LM_L_l73_sw66, LM_L_l73_sw67, LM_L_l73_sw68, LM_L_l74_PF, LM_L_l74_Parts, _
  LM_L_l74_Sculpts, LM_L_l74_UnderPF, LM_L_l74_sw66, LM_L_l74_sw67, LM_L_l74_sw68, LM_L_l75_PF, LM_L_l75_Parts, LM_L_l75_Sculpts, LM_L_l75_UnderPF, LM_L_l75_sw67, LM_L_l75_sw68, LM_L_l76_PF, LM_L_l76_Parts, LM_L_l76_UnderPF, LM_L_l76_sw86, LM_L_l76_sw87, LM_L_l77_PF, LM_L_l77_Parts, LM_L_l77_UnderPF, LM_L_l77_sw86, LM_L_l77_sw87, LM_L_l77_sw88, LM_L_l78_PF, LM_L_l78_Parts, LM_L_l78_UnderPF, LM_L_l78_sw87, LM_L_l78_sw88, LM_L_l81_Layer0, LM_L_l81_PF, LM_L_l81_Parts, LM_L_l81_Sculpts, LM_L_l82_PF, LM_L_l82_UnderPF, LM_L_l83_BmpsTop, LM_L_l83_Bumper2_Ring, LM_L_l83_Layer2, LM_L_l83_Layer3, LM_L_l83_PF, LM_L_l83_Parts, LM_L_l83_PartsAbove, LM_L_l83_UnderPF, LM_L_l84_BmpsBtm, LM_L_l84_BmpsTop, LM_L_l84_Bumper2_Ring, LM_L_l84_Bumper2_Socket, LM_L_l84_Layer2, LM_L_l84_Layer3, LM_L_l84_Layer4, LM_L_l84_PF, LM_L_l84_Parts, LM_L_l84_PartsAbove, LM_L_l84_UnderPF, LM_L_l85_BmpsTop, LM_L_l85_Bumper1_Ring, LM_L_l85_Layer3, LM_L_l85_PF, LM_L_l85_Parts, LM_L_l85_UnderPF, LM_L_l86_PF, LM_L_l86_Parts, LM_L_l86_UnderPF, _
  LM_L_ml1_PF, LM_L_ml1_Parts, LM_L_ml10_PF, LM_L_ml11_PF, LM_L_ml11_Parts, LM_L_ml12_PF, LM_L_ml13_PF, LM_L_ml2_PF, LM_L_ml3_PF, LM_L_ml4_PF, LM_L_ml5_PF, LM_L_ml6_PF, LM_L_ml7_PF, LM_L_ml8_PF, LM_L_ml9_PF, LM_L_Bmp_BmpsBtm, LM_L_Bmp_BmpsTop, LM_L_Bmp_Bumper1_Ring, LM_L_Bmp_Bumper1_Socket, LM_L_Bmp_Bumper2_Ring, LM_L_Bmp_Bumper2_Socket, LM_L_Bmp_Bumper3_Ring, LM_L_Bmp_Bumper3_Socket, LM_L_Bmp_Layer0, LM_L_Bmp_Layer2, LM_L_Bmp_Layer3, LM_L_Bmp_Layer4, LM_L_Bmp_PF, LM_L_Bmp_Parts, LM_L_Bmp_RRamp_DP, LM_L_Bmp_RRamp_UP, LM_L_Bmp_UnderPF)
' VLM  Arrays - End



' VLM T Arrays - Start
' Arrays per baked part
Dim BP_TCandles: BP_TCandles=Array(TBM_Candles, TLM_L_fla1_Candles, TLM_L_fla2_Candles, TLM_L_fla3_Candles, TLM_L_fla4_Candles, TLM_L_fla5_Candles, TLM_L_fla6_Candles)
Dim BP_TFlame1: BP_TFlame1=Array(TBM_Flame1, TLM_L_fla1_Flame1, TLM_L_fla3_Flame1, TLM_L_fla4_Flame1, TLM_L_fla5_Flame1)
Dim BP_TFlame2: BP_TFlame2=Array(TBM_Flame2, TLM_L_fla1_Flame2, TLM_L_fla2_Flame2, TLM_L_fla5_Flame2, TLM_L_fla6_Flame2)
Dim BP_TFlame3: BP_TFlame3=Array(TBM_Flame3, TLM_L_fla1_Flame3, TLM_L_fla3_Flame3, TLM_L_fla4_Flame3, TLM_L_fla5_Flame3)
Dim BP_TFlame4: BP_TFlame4=Array(TBM_Flame4, TLM_L_fla1_Flame4, TLM_L_fla3_Flame4, TLM_L_fla4_Flame4)
Dim BP_TFlame5: BP_TFlame5=Array(TBM_Flame5, TLM_L_fla1_Flame5, TLM_L_fla3_Flame5, TLM_L_fla5_Flame5)
Dim BP_TFlame6: BP_TFlame6=Array(TBM_Flame6, TLM_L_fla2_Flame6, TLM_L_fla6_Flame6)
Dim BP_THead: BP_THead=Array(TBM_Head, TLM_L_fla1_Head, TLM_L_fla2_Head, TLM_L_fla3_Head, TLM_L_fla4_Head, TLM_L_fla5_Head, TLM_L_fla6_Head, TLM_L_l61_Head, TLM_L_l62_Head, TLM_L_l63_Head)
Dim BP_TWall: BP_TWall=Array(TBM_Wall, TLM_L_fla1_Wall, TLM_L_fla2_Wall, TLM_L_fla3_Wall, TLM_L_fla4_Wall, TLM_L_fla5_Wall, TLM_L_fla6_Wall)
' Arrays per lighting scenario
Dim BL_TL_fla1: BL_TL_fla1=Array(TLM_L_fla1_Candles, TLM_L_fla1_Flame1, TLM_L_fla1_Flame2, TLM_L_fla1_Flame3, TLM_L_fla1_Flame4, TLM_L_fla1_Flame5, TLM_L_fla1_Head, TLM_L_fla1_Wall)
Dim BL_TL_fla2: BL_TL_fla2=Array(TLM_L_fla2_Candles, TLM_L_fla2_Flame2, TLM_L_fla2_Flame6, TLM_L_fla2_Head, TLM_L_fla2_Wall)
Dim BL_TL_fla3: BL_TL_fla3=Array(TLM_L_fla3_Candles, TLM_L_fla3_Flame1, TLM_L_fla3_Flame3, TLM_L_fla3_Flame4, TLM_L_fla3_Flame5, TLM_L_fla3_Head, TLM_L_fla3_Wall)
Dim BL_TL_fla4: BL_TL_fla4=Array(TLM_L_fla4_Candles, TLM_L_fla4_Flame1, TLM_L_fla4_Flame3, TLM_L_fla4_Flame4, TLM_L_fla4_Head, TLM_L_fla4_Wall)
Dim BL_TL_fla5: BL_TL_fla5=Array(TLM_L_fla5_Candles, TLM_L_fla5_Flame1, TLM_L_fla5_Flame2, TLM_L_fla5_Flame3, TLM_L_fla5_Flame5, TLM_L_fla5_Head, TLM_L_fla5_Wall)
Dim BL_TL_fla6: BL_TL_fla6=Array(TLM_L_fla6_Candles, TLM_L_fla6_Flame2, TLM_L_fla6_Flame6, TLM_L_fla6_Head, TLM_L_fla6_Wall)
Dim BL_TL_l61: BL_TL_l61=Array(TLM_L_l61_Head)
Dim BL_TL_l62: BL_TL_l62=Array(TLM_L_l62_Head)
Dim BL_TL_l63: BL_TL_l63=Array(TLM_L_l63_Head)
Dim BL_TWorld: BL_TWorld=Array(TBM_Candles, TBM_Flame1, TBM_Flame2, TBM_Flame3, TBM_Flame4, TBM_Flame5, TBM_Flame6, TBM_Head, TBM_Wall)
' Global arrays
Dim BGT_Bakemap: BGT_Bakemap=Array(TBM_Candles, TBM_Flame1, TBM_Flame2, TBM_Flame3, TBM_Flame4, TBM_Flame5, TBM_Flame6, TBM_Head, TBM_Wall)
Dim BGT_Lightmap: BGT_Lightmap=Array(TLM_L_fla1_Candles, TLM_L_fla1_Flame1, TLM_L_fla1_Flame2, TLM_L_fla1_Flame3, TLM_L_fla1_Flame4, TLM_L_fla1_Flame5, TLM_L_fla1_Head, TLM_L_fla1_Wall, TLM_L_fla2_Candles, TLM_L_fla2_Flame2, TLM_L_fla2_Flame6, TLM_L_fla2_Head, TLM_L_fla2_Wall, TLM_L_fla3_Candles, TLM_L_fla3_Flame1, TLM_L_fla3_Flame3, TLM_L_fla3_Flame4, TLM_L_fla3_Flame5, TLM_L_fla3_Head, TLM_L_fla3_Wall, TLM_L_fla4_Candles, TLM_L_fla4_Flame1, TLM_L_fla4_Flame3, TLM_L_fla4_Flame4, TLM_L_fla4_Head, TLM_L_fla4_Wall, TLM_L_fla5_Candles, TLM_L_fla5_Flame1, TLM_L_fla5_Flame2, TLM_L_fla5_Flame3, TLM_L_fla5_Flame5, TLM_L_fla5_Head, TLM_L_fla5_Wall, TLM_L_fla6_Candles, TLM_L_fla6_Flame2, TLM_L_fla6_Flame6, TLM_L_fla6_Head, TLM_L_fla6_Wall, TLM_L_l61_Head, TLM_L_l62_Head, TLM_L_l63_Head)
Dim BGT_All: BGT_All=Array(TBM_Candles, TBM_Flame1, TBM_Flame2, TBM_Flame3, TBM_Flame4, TBM_Flame5, TBM_Flame6, TBM_Head, TBM_Wall, TLM_L_fla1_Candles, TLM_L_fla1_Flame1, TLM_L_fla1_Flame2, TLM_L_fla1_Flame3, TLM_L_fla1_Flame4, TLM_L_fla1_Flame5, TLM_L_fla1_Head, TLM_L_fla1_Wall, TLM_L_fla2_Candles, TLM_L_fla2_Flame2, TLM_L_fla2_Flame6, TLM_L_fla2_Head, TLM_L_fla2_Wall, TLM_L_fla3_Candles, TLM_L_fla3_Flame1, TLM_L_fla3_Flame3, TLM_L_fla3_Flame4, TLM_L_fla3_Flame5, TLM_L_fla3_Head, TLM_L_fla3_Wall, TLM_L_fla4_Candles, TLM_L_fla4_Flame1, TLM_L_fla4_Flame3, TLM_L_fla4_Flame4, TLM_L_fla4_Head, TLM_L_fla4_Wall, TLM_L_fla5_Candles, TLM_L_fla5_Flame1, TLM_L_fla5_Flame2, TLM_L_fla5_Flame3, TLM_L_fla5_Flame5, TLM_L_fla5_Head, TLM_L_fla5_Wall, TLM_L_fla6_Candles, TLM_L_fla6_Flame2, TLM_L_fla6_Flame6, TLM_L_fla6_Head, TLM_L_fla6_Wall, TLM_L_l61_Head, TLM_L_l62_Head, TLM_L_l63_Head)
' VLM T Arrays - End




' VLM B Arrays - Start
' Arrays per baked part
Dim BP_BBG: BP_BBG=Array(BBM_BG, BLM_Lbg_F17_Dome_BG, BLM_Lbg_F18_BG, BLM_Lbg_F19_BG, BLM_Lbg_F20_BG, BLM_Lbg_F21_BG, BLM_Lbg_F22_BG, BLM_Lbg_F23_BG, BLM_Lbg_F24_BG, BLM_Lbg_F26_BG, BLM_Lbg_gib07_BG, BLM_Lbg_git01_BG, BLM_Lbg_gim01_BG, BLM_Lbg_gibg4_BG, BLM_Lbg_gibg5_BG)
' Arrays per lighting scenario
Dim BL_BLbg_F17_Dome: BL_BLbg_F17_Dome=Array(BLM_Lbg_F17_Dome_BG)
Dim BL_BLbg_F18: BL_BLbg_F18=Array(BLM_Lbg_F18_BG)
Dim BL_BLbg_F19: BL_BLbg_F19=Array(BLM_Lbg_F19_BG)
Dim BL_BLbg_F20: BL_BLbg_F20=Array(BLM_Lbg_F20_BG)
Dim BL_BLbg_F21: BL_BLbg_F21=Array(BLM_Lbg_F21_BG)
Dim BL_BLbg_F22: BL_BLbg_F22=Array(BLM_Lbg_F22_BG)
Dim BL_BLbg_F23: BL_BLbg_F23=Array(BLM_Lbg_F23_BG)
Dim BL_BLbg_F24: BL_BLbg_F24=Array(BLM_Lbg_F24_BG)
Dim BL_BLbg_F26: BL_BLbg_F26=Array(BLM_Lbg_F26_BG)
Dim BL_BLbg_gib07: BL_BLbg_gib07=Array(BLM_Lbg_gib07_BG)
Dim BL_BLbg_gibg4: BL_BLbg_gibg4=Array(BLM_Lbg_gibg4_BG)
Dim BL_BLbg_gibg5: BL_BLbg_gibg5=Array(BLM_Lbg_gibg5_BG)
Dim BL_BLbg_gim01: BL_BLbg_gim01=Array(BLM_Lbg_gim01_BG)
Dim BL_BLbg_git01: BL_BLbg_git01=Array(BLM_Lbg_git01_BG)
Dim BL_BWorld: BL_BWorld=Array(BBM_BG)
' Global arrays
Dim BGB_Bakemap: BGB_Bakemap=Array(BBM_BG)
Dim BGB_Lightmap: BGB_Lightmap=Array(BLM_Lbg_F17_Dome_BG, BLM_Lbg_F18_BG, BLM_Lbg_F19_BG, BLM_Lbg_F20_BG, BLM_Lbg_F21_BG, BLM_Lbg_F22_BG, BLM_Lbg_F23_BG, BLM_Lbg_F24_BG, BLM_Lbg_F26_BG, BLM_Lbg_gib07_BG, BLM_Lbg_gibg4_BG, BLM_Lbg_gibg5_BG, BLM_Lbg_gim01_BG, BLM_Lbg_git01_BG)
Dim BGB_All: BGB_All=Array(BBM_BG, BLM_Lbg_F17_Dome_BG, BLM_Lbg_F18_BG, BLM_Lbg_F19_BG, BLM_Lbg_F20_BG, BLM_Lbg_F21_BG, BLM_Lbg_F22_BG, BLM_Lbg_F23_BG, BLM_Lbg_F24_BG, BLM_Lbg_F26_BG, BLM_Lbg_gib07_BG, BLM_Lbg_gibg4_BG, BLM_Lbg_gibg5_BG, BLM_Lbg_gim01_BG, BLM_Lbg_git01_BG)
' VLM B Arrays - End



'*******************************************
' ZLOA: Load Stuff
'*******************************************

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

If RenderingMode=2 Then
  UseVPMDMD = True
Else
  UseVPMDMD = DesktopMode
End If

LoadVPM "03060000", "WPC.VBS", 3.26

NoUpperRightFlipper   'Disables secondary flipper controls as they're used for mist gates
NoUpperLeftFlipper




'*******************************************
' ZTIM: Main Timers
'*******************************************

Sub FrameTimer_timer()
  BSUpdate
  RollingUpdate   'update rolling sounds
  UpdateBallBrightness
  'AnimateBumperSkirts
  DoSTAnim    'handle stand up target animations
  UpdateStandupTargets
  DoDTAnim    'handle drop up target animations
  UpdateDropTargets
End Sub


' FIXME for the time being, the cor timer interval must be 10 ms (so below 60FPS framerate)
CorTimer.Interval = 10
Sub CorTimer_Timer(): Cor.Update: End Sub



'*******************************************
' ZINI: Table Initialization and Exiting
'*******************************************

Sub Table1_Init
  vpmInit Me
  vpmMapLights AllLamps

  With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
    .SplashInfoLine = "Bram Stoker's Dracula, Williams, 1993" & vbNewLine & "VPW"
    .Games(cGameName).Settings.Value("rol") = 0 'rotated left
    .HandleMechanics = 0
    .HandleKeyboard = 0
    .ShowDMDOnly = 1
    .ShowFrame = 0
    .ShowTitle = 0
    .Hidden = 0
    On Error Resume Next
    .Run GetPlayerHWnd
    If Err Then MsgBox Err.Description
    On Error Goto 0
    .Switch(22) = 1 'close coin door
    .Switch(24) = 1 'and keep it closed
  End With

  vpmNudge.TiltSwitch = 14
  vpmNudge.Sensitivity = 5
  vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingShot, RightSlingShot)

  ' Trough
  Set BSDBall1 = sw44.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set BSDBall2 = sw43.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set BSDBall3 = sw42.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set BSDBall4 = sw41.CreateSizedballWithMass(Ballsize/2,Ballmass)
  gBOT = Array(BSDBall1,BSDBall2,BSDBall3,BSDBall4)

  Controller.Switch(44) = 1
  Controller.Switch(43) = 1
  Controller.Switch(42) = 1
  Controller.Switch(41) = 1

  ' Mist Magnet
  Set mMagnet = New cMagnet
  ' mMagnet.InitMagnet Magnet, 5
  mMagnet.InitMagnet Magnet, 0.8
  mMagnet.Size = 65
  MagnetPos = 0
  SetMagnetPosition
    Controller.Switch(81) = (MagnetPos > 490)
    Controller.Switch(83) = (MagnetPos < 10)

  Plunger.Pullback
  Wdivert.Isdropped = 1

  ' Other inits
  LStep=0:LeftSlingShot.TimerEnabled = 1
  RStep=0:RightSlingShot.TimerEnabled = 1

  LanternBulbs.blenddisablelighting = 2
  CoinSlots.blenddisablelighting = 25

End Sub

Sub Table1_Paused: Controller.Pause = 1: End Sub
Sub Table1_unPaused: Controller.Pause = 0: End Sub
Sub Table1_Exit: Controller.Stop: End Sub



'*******************************************
' ZKEY: Key Press Handling
'*******************************************

Sub Table1_KeyDown(ByVal Keycode)
  If Keycode = LeftFlipperKey Then
    FlipperActivate LeftFlipper, LFPress
    LeftBut.X = LeftBut.X + 10
  End If
  If Keycode = RightFlipperKey Then
    FlipperActivate RightFlipper, RFPress
    RightBut.X = RightBut.X - 10
  End If
  If keycode = LeftTiltKey Then Nudge 90, 2:SoundNudgeLeft
  If keycode = RightTiltKey Then Nudge 270, 2:SoundNudgeRight
  If keycode = CenterTiltKey Then Nudge 0, 2:SoundNudgeCenter
  If keycode = keyInsertCoin1 or keycode = keyInsertCoin2 or keycode = keyInsertCoin3 or keycode = keyInsertCoin4 Then SoundCoinIn
  If keycode = PlungerKey or keycode = LockBarKey Then
    Plunger.FireSpeed = 95 + 5*rnd
    Controller.Switch(34) = 1
    LaunchBut.Y = LaunchBut.Y - 7
    Flasher1.Y = Flasher1.Y - 7
    soundStartButton
  End If
  if keycode=StartGameKey then soundStartButton
  If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal Keycode)
  If Keycode = LeftFlipperKey Then
    FlipperDeActivate LeftFlipper, LFPress
    LeftBut.X = LeftBut.X - 10
  End If
  If Keycode = RightFlipperKey Then
    FlipperDeActivate RightFlipper, RFPress
    RightBut.X = RightBut.X + 10
  End If
  If keycode = PlungerKey or keycode = LockBarKey Then
    Controller.Switch(34) = 0
    LaunchBut.Y = LaunchBut.Y + 7
    Flasher1.Y = Flasher1.Y + 7
  End If
  If vpmKeyUp(keycode) Then Exit Sub
End Sub




'*******************************************
'   ZOPT: Table Options
'*******************************************

Dim LightLevel : LightLevel = 0.25        ' Level of room lighting (0 to 1), where 0 is dark and 100 is brightest
Dim ColorLUT : ColorLUT = 1           ' Color desaturation LUTs: 1 to 11, where 1 is normal and 11 is black'n'white
Dim VolumeDial : VolumeDial = 0.8             ' Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Dim BallRollVolume : BallRollVolume = 0.5     ' Level of ball rolling volume. Value between 0 and 1
Dim RampRollVolume : RampRollVolume = 0.5     ' Level of ramp rolling volume. Value between 0 and 1
Dim VRRoomChoice : VRRoomChoice = 1       '1 = Minimal Room, 2 = Mega Room
Dim VRPosterLeft : VRPosterLeft = 1
Dim VRPosterRight : VRPosterRight = 2
Dim VRRoom


' Called when options are tweaked by the player.
' - 0: game has started, good time to load options and adjust accordingly
' - 1: an option has changed
' - 2: options have been reseted
' - 3: player closed the tweak UI, good time to update staticly prerendered parts
' Table1.Option arguments are:
' - option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional arry of literal strings
Sub Table1_OptionEvent(ByVal eventId)
    If eventId = 1 Then DisableStaticPreRendering = True
    Dim v, BP

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

  'Desktop DMD
  v = Table1.Option("Desktop DMD", 0, 1, 1, 1, 0, Array("Not Visible", "Visible"))
  TextBox.Visible = v

  'Desktop DMD
  v = Table1.Option("Desktop DMD Color", 0, 6, 1, 0, 0, Array("Red", "Orange", "Yellow", "Green", "Blue", "Purple", "White"))
  Select Case v
    Case 0: TextBox.FontColor = RGB(255,0,0)
    Case 1: TextBox.FontColor = RGB(255,127,0)
    Case 2: TextBox.FontColor = RGB(255,255,0)
    Case 3: TextBox.FontColor = RGB(0,255,0)
    Case 4: TextBox.FontColor = RGB(20,20,255)
    Case 5: TextBox.FontColor = RGB(127,0,255)
    Case 6: TextBox.FontColor = RGB(255,255,255)
  End Select

  'Top GI Color
  v = Table1.Option("Top GI Color", 0, 7, 1, 0, 0, Array("Normal", "Red", "Orange", "Green", "Cyan", "Blue", "Purple", "White"))
  SetTopGiColor v

  'Middle GI Color
  v = Table1.Option("Middle GI Color", 0, 7, 1, 0, 0, Array("Normal", "Red", "Orange", "Green", "Cyan", "Blue", "Purple", "White"))
  SetMidGiColor v

  'Bottom GI Color
  v = Table1.Option("Bottom GI Color", 0, 7, 1, 0, 0, Array("Normal", "Red", "Orange", "Green", "Cyan", "Blue", "Purple", "White"))
  SetBotGiColor v

  'Flashers Color
  v = Table1.Option("Flashers Color", 0, 7, 1, 0, 0, Array("Normal", "Red", "Orange", "Green", "Cyan", "Blue", "Purple", "White"))
  SetFlasherColor v

    ' Toggle Rails
  If RenderingMode <> 2 Then
    v = Table1.Option("Show Rails (Desktop/Cab)", 0, 1, 1, 1, 0, Array("Hide", "Show"))
    For Each BP in BP_RailL : BP.Visible = v: Next
    For Each BP in BP_RailR: BP.Visible = v: Next
  End If

    ' Toggle Lockdown Bar
  If RenderingMode <> 2 Then
    v = Table1.Option("Show Lockbar (Desktop/Cab)", 0, 1, 1, 1, 0, Array("Hide", "Show"))
    For Each BP in BP_lockdownbar : BP.Visible = v: Next
  End If

    ' Toggle Extra Plastics
  v = Table1.Option("Extra Plastics", 0, 1, 1, 1, 0, Array("Hide", "Show"))
  For Each BP in BP_ExtraPlasticL : BP.Visible = v: Next
  For Each BP in BP_ExtraPlasticR : BP.Visible = v: Next
  For Each BP in BP_ExtPlasticSpacer1: BP.Visible = v: Next
  For Each BP in BP_ExtPlasticSpacer2: BP.Visible = v: Next
  For Each BP in BP_ExtPlasticSpacer3: BP.Visible = v: Next
  For Each BP in BP_ExtPlasticSpacer4: BP.Visible = v: Next
  Dim offsetScrew:
  Select Case v
    Case 0
      offsetScrew = -43
    Case 1
      offsetScrew = 0
  End Select
  For Each BP in BP_ExtPlasticScrew1: BP.transZ = offsetScrew: Next
  For Each BP in BP_ExtPlasticScrew2: BP.transZ = offsetScrew: Next
  For Each BP in BP_ExtPlasticScrew3: BP.transZ = offsetScrew: Next
  For Each BP in BP_ExtPlasticScrew4: BP.transZ = offsetScrew: Next

    ' Toggle Reflections
  v = Table1.Option("Playfield Reflections", 0, 2, 1, 1, 0, Array("None (Vampire)", "Partial", "Full"))
  ReflectionToggle(v)

    ' Glass dirt
  v = Table1.Option("Glass dirt", 0, 4, 1, 2, 0, Array("Hide", "25%", "50%", "75%", "100%"))
  F_Glass_spec.Visible = v>0
  F_Glass_dif.Visible = v>0
  Select Case v
    Case 1: F_Glass_spec.Opacity = 50000*0.25: F_Glass_dif.Opacity = 150*0.25
    Case 2: F_Glass_spec.Opacity = 50000*0.50: F_Glass_dif.Opacity = 150*0.50
    Case 3: F_Glass_spec.Opacity = 50000*0.75: F_Glass_dif.Opacity = 150*0.75
    Case 4: F_Glass_spec.Opacity = 50000*1.00: F_Glass_dif.Opacity = 150*1.00
  End Select

    ' Toggle Refraction
  v = Table1.Option("Ramps Refractions", 0, 1, 1, 1, 0, Array("Off", "On"))
  Select Case v
    Case 0
      BM_Layer0.RefractionProbe = ""
      BM_Layer1.RefractionProbe = ""
      BM_Layer2.RefractionProbe = ""
      BM_Layer3.RefractionProbe = ""
      BM_Layer4.RefractionProbe = ""
    Case 1
      BM_Layer0.RefractionProbe = "Refraction Pass 2"
      BM_Layer1.RefractionProbe = "Refraction Pass 2"
      BM_Layer2.RefractionProbe = "Refraction Pass 2"
      BM_Layer3.RefractionProbe = "Refraction Pass 2"
      BM_Layer4.RefractionProbe = "Refraction Pass 2"
  End Select

    ' Outlane Difficulty
  v = Table1.Option("Outlanes difficulty", 0, 2, 1, 1, 0, Array("Easy", "Medium", "Hard"))
  SetOutlaneDifficulty(v)

    ' Sound volumes
    VolumeDial = Table1.Option("Mech Volume", 0, 1, 0.01, 0.8, 1)
    BallRollVolume = Table1.Option("Ball Roll Volume", 0, 1, 0.01, 0.5, 1)
  RampRollVolume = Table1.Option("Ramp Roll Volume", 0, 1, 0.01, 0.5, 1)

  ' Room brightness
' LightLevel = Table1.Option("Table Brightness (Ambient Light Level)", 0, 1, 0.01, .5, 1)
  LightLevel = NightDay/100
  SetRoomBrightness LightLevel

    If eventId = 3 Then DisableStaticPreRendering = False

  'VR Room
  VRRoomChoice = Table1.Option("VR Room", 1, 2, 1, 2, 0, Array("Minimal", "Dracula"))
  If RenderingMode = 2 OR TestVRDT=True Then
    VRRoom = VRRoomChoice
    For Each BP in BP_RailL : BP.Visible = 0: Next
    For Each BP in BP_RailR: BP.Visible = 0: Next
    For Each BP in BP_lockdownbar : BP.Visible = 0: Next
  Else VRRoom = 0
  End If

  'Minimal Room Left Poster"
  VRPosterLeft = Table1.Option("Minimal Room Left Poster", 0, 4, 1, 1, 0, Array("No Poster", "Option 1", "Option 2", "Option 3", "Option 4"))
  If VRPosterLeft > 0 Then
    PosterLeft.Image = "BSD_poster"&VRPosterLeft
  End If

  'Minimal Room Right Posters
  VRPosterRight = Table1.Option("Minimal Room Right Poster", 0, 4, 1, 2, 0, Array("No Poster", "Option 1", "Option 2", "Option 3", "Option 4"))
  If VRPosterRight > 0 Then
    PosterRight.Image = "BSD_poster"&VRPosterRight
  End If

  SetupRoom

End Sub

Function SetOutlaneDifficulty(d)
  Dim BP, v1, v2, v3, y
  y = 1378.664
  Select Case d
    Case 0:
      v1 = False : v2 = False: v3 = True
      For Each BP in BP_RPost_p3 : BP.transy = 0: Next
      For Each BP in BP_LPost_p3 : BP.transy = 0: Next
    Case 1:
      v1 = False: v2 =True:  v3 = False
      For Each BP in BP_RPost_p3 : BP.transy = -10: Next
      For Each BP in BP_LPost_p3 : BP.transy = -10: Next
    Case 2:
      v1 = True:  v2 = False: v3 = False
      For Each BP in BP_RPost_p3 : BP.transy = -20: Next
      For Each BP in BP_LPost_p3 : BP.transy = -20: Next
  End Select
  ' set the BPs
  For Each BP in BP_RPostRubber_p1 : BP.Visible = v1: Next
  For Each BP in BP_RPostRubber_p2 : BP.Visible = v2: Next
  For Each BP in BP_RPostRubber_p3 : BP.Visible = v3: Next
  For Each BP in BP_LPostRubber_p1 : BP.Visible = v1: Next
  For Each BP in BP_LPostRubber_p2 : BP.Visible = v2: Next
  For Each BP in BP_LPostRubber_p3 : BP.Visible = v3: Next
  ' set the physics
  Outlane_L_Easy_Post.collidable = v3: Outlane_L_Easy_Rub.collidable = v3
  Outlane_L_Med_Post.collidable  = v2:  Outlane_L_Med_Rub.collidable = v2
  Outlane_L_Hard_Post.collidable = v1: Outlane_L_Hard_Rub.collidable = v1
  Outlane_R_Easy_Post.collidable = v3: Outlane_R_Easy_Rub.collidable = v3
  Outlane_R_Med_Post.collidable  = v2:  Outlane_R_Med_Rub.collidable = v2
  Outlane_R_Hard_Post.collidable = v1: Outlane_R_Hard_Rub.collidable = v1
End Function

Function ReflectionToggle(state)
  Dim ReflObjectsArray: ReflObjectsArray = Array(BP_Parts, BP_LFlip, BP_LFlipU, BP_RFlip, BP_RFlipU, BP_BmpsBtm, BP_Bumper2_Ring, BP_Bumper2_Socket, BP_sw86, BP_sw87, BP_sw88)
  Dim ReflPrtObjsArray: ReflPrtObjsArray = Array(BM_Parts, BM_LFlip, BM_LFlipU, BM_RFlip, BM_RFlipU, BM_BmpsBtm, BM_Bumper2_Ring, BM_Bumper2_Socket, BM_sw86, BM_sw87, BM_sw88)
  Dim IBP, BP, v1, v2
  Select Case state
    Case 0: v1 = false: v2 = false  'No reflections
    Case 1: v1 = false: v2 = true   'Only reflect static prims
    Case 2: v1 = true:  v2 = true   'Reflect everything (that matters as defined in ReflObjectsArray)
  End Select

  For Each IBP in ReflObjectsArray
    For Each BP in IBP
      BP.ReflectionEnabled = v1
    Next
  Next

  For Each BP in ReflPrtObjsArray
    BP.ReflectionEnabled = v2
  Next

  BM_Parts.ReflectionEnabled = v2
  BM_PartsAbove.ReflectionEnabled = v2
  BM_PF.ReflectionEnabled = False
End Function


'******************************************************
'   ZBBR: Ball Brightness
'******************************************************

Const BallBrightness =  1        'Ball brightness - Value between 0 and 1 (0=Dark ... 1=Bright)
Const PLOffset = 0.5
Dim PLGain: PLGain = (1-PLOffset)/(1260-2000)

Sub UpdateBallBrightness
  Dim s, b_base, b_r, b_g, b_b, d_w
  b_base = 120 * BallBrightness + 135 ' orig was 120 and 70

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
'   ZBRI: Room Brightness
'****************************

' Update these arrays if you want to change more materials with room light level
Dim RoomBrightnessMtlArray: RoomBrightnessMtlArray = Array("VLM.Bake.Active","VLM.Bake.Solid", "VLM.Bake.Bumper", "VLM.Bake.Ramp", "VLM.Bake.RampEdge", "VLM.Bake.CoffinLid", "VR.Metals")

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
      For each bulb in GIBOT: bulb.State = pwm: Next
      'GIRelaySound pwm,gi0lvl
      gi0lvl = pwm
    Case 1 :
      For each bulb in GITOP: bulb.State = pwm: Next
      'GIRelaySound pwm,gi1lvl
      gi1lvl = pwm
    Case 2 :
      For each bulb in GIMID: bulb.State = pwm: Next
      'GIRelaySound pwm,gi2lvl
      gi2lvl = pwm
    Case 3 : gibg4.state = pwm
    Case 4 : gibg5.state = pwm
  End Select
End Sub


Sub GIRelaySound(pwm, gilvl)
  If pwm >= 0.5 And gilvl < 0.5 Then
    Sound_GI_Relay 1, Bumper1
  ElseIf pwm <= 0.4 And gilvl > 0.4 Then
    Sound_GI_Relay 0, Bumper1
  End If
End Sub



'****************************
' ZGIC: GI Colors
'****************************

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
Dim cCyanFull: cCyanFull= rgb(0,255,255)
Dim cCyan : cCyan = rgb(5,128,255)
Dim cYellowFull:cYellowFull  = rgb(255,255,128)
Dim cYellow: cYellow = rgb(255,255,0)
Dim cOrangeFull: cOrangeFull = rgb(255,128,0)
Dim cOrange: cOrange = rgb(255,70,5)
Dim cGreenFull: cGreenFull = rgb(0,255,0)
Dim cGreen: cGreen = rgb(5,255,5)
Dim cPurpleFull: cPurpleFull = rgb(128,0,255)
Dim cPurple: cPurple = rgb(60,5,255)
Dim cAmberFull: cAmberFull = rgb(255,197,143)
Dim cAmber: cAmber = rgb(255,197,143)

Dim cArray
cArray = Array(c2700k,cRed,cOrange,cGreen,cCyanFull,cBlue,cPurple,cWhite)

Sub SetBotGiColor(c)
  Dim xx, BL
  For each xx in GIBot: xx.color = cArray(c): xx.colorfull = cArray(c): Next
  For each BL in BL_GI_gib01: BL.color = cArray(c): Next
  For each BL in BL_GI_gib02: BL.color = cArray(c): Next
  'For each BL in BL_GI_gib03: BL.color = cArray(c): Next
  For each BL in BL_GI_gib04: BL.color = cArray(c): Next
  For each BL in BL_GI_gib05: BL.color = cArray(c): Next
  For each BL in BL_GI_gib06: BL.color = cArray(c): Next
  'For each BL in BL_GI_gib07: BL.color = cArray(c): Next
  For each BL in BL_GI_gib08: BL.color = cArray(c): Next
End Sub

Sub SetMidGiColor(c)
  Dim xx, BL
  For each xx in GIMid: xx.color = cArray(c): xx.colorfull = cArray(c): Next
  gim001.color = cArray(0): gim001.colorfull = cArray(0)
  For each BL in BL_GI_gim01: BL.color = cArray(c): Next
  For each BL in BL_GI_gim04: BL.color = cArray(c): Next
  For each BL in BL_GI_gim05: BL.color = cArray(c): Next
  'For each BL in BL_GI_gim08: BL.color = cArray(c): Next
  'For each BL in BL_GI_gim09: BL.color = cArray(c): Next
End Sub

Sub SetTopGiColor(c)
  Dim xx, BL
  For each xx in GITop: xx.color = cArray(c): xx.colorfull = cArray(c): Next
  For each BL in BL_GI_git01: BL.color = cArray(c): Next
End Sub


Sub SetFlasherColor(c)
  Dim BL
  F17_Dome.color = cArray(c): F17_Dome.colorfull = cArray(c)
  F19.color = cArray(c): F19.colorfull = cArray(c)
  F19a.color = cArray(c): F19a.colorfull = cArray(c)
  F20.color = cArray(c): F20.colorfull = cArray(c)
  F20_Dome.color = cArray(c): F20_Dome.colorfull = cArray(c)
  F21.color = cArray(c): F21.colorfull = cArray(c)
  F21a.color = cArray(c): F21a.colorfull = cArray(c)
  F24.color = cArray(c): F24.colorfull = cArray(c)
  For each BL in BL_FL_F17_Dome: BL.color = cArray(c): Next
  For each BL in BL_FL_F19: BL.color = cArray(c): Next
  For each BL in BL_FL_F19a: BL.color = cArray(c): Next
  For each BL in BL_FL_F20: BL.color = cArray(c): Next
  For each BL in BL_FL_F20_Dome: BL.color = cArray(c): Next
  For each BL in BL_FL_F21: BL.color = cArray(c): Next
  For each BL in BL_FL_F21a: BL.color = cArray(c): Next
  For each BL in BL_FL_F24: BL.color = cArray(c): Next
End Sub





'****************************
' ZSOL: Solenoids
'****************************

SolCallback(1) = "AutoPlunger"
SolCallback(2) = "CoffinPopper"       'VUK
SolCallback(3) = "CastlePopper"       'VUK
SolCallback(4) = "SolRRampDown"
SolCallback(5) = "CryptPopper"        'VUK
SolCallback(6) = "WirerampPopper"     'VUK
SolCallback(7) = "SolKnocker"
SolCallback(8) = "SolShooterRamp"
SolCallback(14) = "SolRRampUp"
SolCallback(16) = "SolRelease"

SolCallback(25) = "SolDTUp"
SolCallback(27) = "SolMistMagnet"
SolCallback(33) = "solTopDiverter"
SolCallback(34) = "SolRGate"
SolCallback(35) = "CastleLockPost"      'Castle Lock Post
SolCallback(36) = "SolLGate"

' Flashers
SolModCallback(17) = "Flasher17"      'Top Right Corner Flasher (#906), Dracula Flasher (x2) (#906)
SolModCallback(18) = "Flasher18"      'Jackpot Flasher (#906), Stoker Flasher (#906)
SolModCallback(19) = "Flasher19"      '3 Bank Flasher (#89), House Flasher (#906)
SolModCallback(20) = "Flasher20"      'Top Left Corner Flasher (#89, #906), Mina Flasher (#906)
SolModCallback(21) = "Flasher21"      'Castle Flasher (#89), Helsing Flasher (#906)
SolModCallback(22) = "Flasher22"      'Left Ramp Flasher (#906), Left Logo Flasher (#906)
SolModCallback(23) = "Flasher23"      'Right Ramp Flasher (#906), Right Logo Flasher (#906)
SolModCallback(24) = "Flasher24"      'Asylum Flasher (#89), Renfield Flasher (#906)
SolModCallback(26) = "Flasher26"      'Speaker panel

' Flippers
SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"


'  Misc
'**********

Sub AutoPlunger(Enabled)
  If Enabled Then
    Plunger.FireSpeed = 95 + 5*rnd
    Plunger.Fire
  Else
    Plunger.PullBack
  End If
End Sub


Sub SolKnocker(Enabled)
  If enabled Then
    KnockerSolenoid 'Add knocker position object
  End If
End Sub



' Drop targets
'******************

Sub SolDTUp(Enabled)
  If Enabled Then
    DTRaise 15
    RandomSoundDropTargetReset BM_sw15
  End If
End Sub




' Castle Lock Post
'******************

clpost.TimerInterval = 280
clpost.TimerEnabled = False

Sub CastleLockPost(Enabled)
  If Enabled Then
    clpost.isdropped = 0
    SoundSaucerKick 0,sw53
  Else
    clpost.isdropped = 1
    clpost.TimerEnabled = True
  End If
End Sub

Sub clpost_Timer
  clpost.isdropped = 0
  clpost.TimerEnabled = False
  SoundSaucerKick 0,sw53
End Sub

Sub clpost_hit
  RandomSoundMetal
End Sub



' Moving Ramp
'*************

Controller.Switch(77) = False
'RightRamp.Collidable = True

Sub SolRRampUp(Enabled)
  If Enabled Then
  FRRamp.rotatetoend
    Controller.Switch(77) = True
    RightRamp.Collidable = False
  SoundSaucerKick 0,Bumper3
  End If
End Sub

Sub SolRRampDown(Enabled)
  If Enabled Then
  FRRamp.rotatetostart
    Controller.Switch(77) = False
    RightRamp.Collidable = True
  SoundSaucerKick 0,Bumper3
  End If
End sub



' Shooter Ramp
'*************
Sub SolShooterRamp(Enabled)
  If Enabled Then
    sramp2.collidable = 0
  FLRamp.rotatetoend
  SoundSaucerKick 0,RampTrigger008
  Else
    sramp2.collidable = 1
    FLRamp.rotatetostart
  SoundSaucerKick 0,RampTrigger008
  End If
End Sub




' Top Ramp Diverter
'******************

Sub SolTopDiverter(Enabled)
  If Enabled Then
    wDivert.isdropped = 0
    FDivert.rotatetostart
  Else
    WDivert.isdropped = 1
    FDivert.rotatetoend
  End If
End sub



' Mist Gates
'******************

Sub SolLGate(Enabled)
  If Enabled then
    LGate.open = 1
    Wall_LO.isdropped = 1
  else
    LGate.open = 0
    Wall_LO.isdropped = 0
  End If
End Sub

Sub SolRGate(Enabled)
  If Enabled then
    RGate.open = 1
    RGateWall.IsDropped = True
  else
    RGate.open = 0
    RGateWall.IsDropped = False
  End If
End Sub




' Flashers
'**********

const DebugFlashers = False

Sub Flasher17(pwm)
  If DebugFlashers then debug.print "Flasher17 "&pwm
  F17_Dome.State = pwm
End Sub

Sub Flasher18(pwm)
  If DebugFlashers then debug.print "Flasher18 "&pwm
  F18.State = pwm
End Sub

Sub Flasher19(pwm)
  If DebugFlashers then debug.print "Flasher19 "&pwm
  F19.State = pwm
  F19a.State = pwm
End Sub

Sub Flasher20(pwm)
  If DebugFlashers then debug.print "Flasher20 "&pwm
  F20.State = pwm
  F20_Dome.State = pwm
End Sub

Sub Flasher21(pwm)
  If DebugFlashers then debug.print "Flasher21 "&pwm
  F21.State = pwm
  F21a.State = pwm
End Sub

Sub Flasher22(pwm)
  If DebugFlashers then debug.print "Flasher22 "&pwm
  F22.State = pwm
End Sub

Sub Flasher23(pwm)
  If DebugFlashers then debug.print "Flasher23 "&pwm
  F23.State = pwm
End Sub

Sub Flasher24(pwm)
  If DebugFlashers then debug.print "Flasher24 "&pwm
  F24.State = pwm
End Sub

Sub Flasher26(pwm)
  If DebugFlashers then debug.print "Flasher26 "&pwm
  F26.State = pwm
End Sub




'****************************
' ZFLP: Flippers
'****************************


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



'****************************
' ZSLG: Slingshots
'****************************

' Slings & div switches
Dim LStep
Dim RStep

Sub LeftSlingShot_Slingshot
  LS.VelocityCorrect(Activeball)
  RandomSoundSlingshotLeft zCol_Rubber_Post003
  vpmTimer.PulseSw 64
  LStep = 0
  LeftSlingShot_Timer
  LeftSlingShot.TimerEnabled = 1
  LeftSlingShot.TimerInterval = 17
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
  For Each BP in BP_LSlingArm : BP.transx = y: Next

    LStep = LStep + 1
End Sub


Sub RightSlingShot_Slingshot
  RS.VelocityCorrect(Activeball)
  RandomSoundSlingshotRight zCol_Rubber_Post004
  vpmTimer.PulseSw 65
  RStep = 0
  RightSlingShot_Timer
  RightSlingShot.TimerEnabled = 1
  RightSlingShot.TimerInterval = 17
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
  For Each BP in BP_RSlingArm : BP.transx = y: Next

    RStep = RStep + 1
End Sub



'****************************
' ZBMP: Bumpers
'****************************

'**
'            SKIRT ANIMATION FUNCTIONS
'**
' NOTE: set bumper object timer to around 150-175 in order to be able
' to actually see the animation, adjust to your liking

'Const PI = 3.1415926
Const SkirtTilt = 6       'angle of skirt tilting in degrees

Function SkirtAX(bumper, bumperball)
    skirtAX=cos(skirtA(bumper,bumperball))*(SkirtTilt)        'x component of angle
    if (bumper.y<bumperball.y) then    skirtAX=-skirtAX    'adjust for ball hit bottom half
End Function

Function SkirtAY(bumper, bumperball)
    skirtAY=sin(skirtA(bumper,bumperball))*(SkirtTilt)        'y component of angle
    if (bumper.x>bumperball.x) then    skirtAY=-skirtAY    'adjust for ball hit left half
End Function

Function SkirtA(bumper, bumperball)
    dim hitx, hity, dx, dy
    hitx=bumperball.x
    hity=bumperball.y

    dy=Abs(hity-bumper.y)                    'y offset ball at hit to center of bumper
    if dy=0 then dy=0.0000001
    dx=Abs(hitx-bumper.x)                    'x offset ball at hit to center of bumper
    skirtA=(atn(dx/dy)) '/(PI/180)            'angle in radians to ball from center of Bumper1
End Function


Sub Bumper1_Hit()
  vpmTimer.PulseSw 63
  RandomSoundBumperBottom Bumper1

  'VLM skirt movable script
  Dim BP
  For each BP in BP_Bumper1_Socket
    BP.roty = skirtAY(me,Activeball)
    BP.rotx = skirtAX(me,Activeball)
  Next
  me.TimerEnabled = 1
End Sub

Sub Bumper1_timer
  'VLM movable script
  Dim BP
  For each BP in BP_Bumper1_Socket
    BP.roty = 0
    BP.rotx = 0
  Next
  me.TimerEnabled = 0
End Sub

Sub Bumper2_Hit()
  vpmTimer.PulseSw 61
  RandomSoundBumperTop Bumper2
  'VLM skirt movable script
  Dim BP
  For each BP in BP_Bumper2_Socket
    BP.roty = skirtAY(me,Activeball)
    BP.rotx = skirtAX(me,Activeball)
  Next
  me.TimerEnabled = 1
End Sub

Sub Bumper2_timer
  'VLM movable script
  Dim BP
  For each BP in BP_Bumper2_Socket
    BP.roty = 0
    BP.rotx = 0
  Next
  me.TimerEnabled = 0
End Sub

Sub Bumper3_Hit()
  vpmTimer.PulseSw 62
  RandomSoundBumperTop Bumper3
  'VLM skirt movable script
  Dim BP
  For each BP in BP_Bumper3_Socket
    BP.roty = skirtAY(me,Activeball)
    BP.rotx = skirtAX(me,Activeball)
  Next
  me.TimerEnabled = 1
End Sub

Sub Bumper3_timer
  'VLM movable script
  Dim BP
  For each BP in BP_Bumper3_Socket
    BP.roty = 0
    BP.rotx = 0
  Next
  me.TimerEnabled = 0
End Sub




'****************************
' ZSWI: Switches
'****************************


' Rollovers & Ramp Switches
Sub sw35_Hit():Controller.Switch(35) = 1:End Sub
Sub sw35_UnHit():Controller.Switch(35) = 0:End Sub

Sub sw36_Hit():Controller.Switch(36) = 1: leftInlaneSpeedLimit: End Sub
Sub sw36_UnHit():Controller.Switch(36) = 0:End Sub

Sub sw37_Hit():Controller.Switch(37) = 1: rightInlaneSpeedLimit: End Sub
Sub sw37_UnHit():Controller.Switch(37) = 0:End Sub

Sub sw38_Hit():Controller.Switch(38) = 1:End Sub
Sub sw38_UnHit():Controller.Switch(38) = 0:End Sub

Sub sw25_Hit():Controller.Switch(25) = 1:End Sub
Sub sw25_UnHit():Controller.Switch(25) = 0:End Sub

Sub sw26_Hit():Controller.Switch(26) = 1:End Sub
Sub sw26_UnHit():Controller.Switch(26) = 0:End Sub

Sub sw27_Hit():Controller.Switch(27) = 1:End Sub
Sub sw27_UnHit():Controller.Switch(27) = 0:End Sub

Sub sw16_Hit():Controller.Switch(16) = 1:End Sub
Sub sw16_UnHit():Controller.Switch(16) = 0:End Sub

Sub sw28_Hit():Controller.Switch(28) = 1: WireRampOff: End Sub
Sub sw28_UnHit():Controller.Switch(28) = 0:WireRampOn False: End Sub

Sub sw84_Hit():Controller.Switch(84) = 1:End Sub
Sub sw84_UnHit():Controller.Switch(84) = 0:End Sub

Sub sw85_Hit():Controller.Switch(85) = 1:WireRampOff: End Sub
Sub sw85_UnHit():Controller.Switch(85) = 0:WireRampOn False: End Sub

Sub sw31_Hit():Controller.Switch(31) = 1:End Sub
Sub sw31_UnHit():Controller.Switch(31) = 0:End Sub

Sub sw51_Hit():Controller.Switch(51) = 1:End Sub
Sub sw51_UnHit():Controller.Switch(51) = 0:End Sub

Sub sw52_Hit():Controller.Switch(52) = 1:End Sub
Sub sw52_UnHit():Controller.Switch(52) = 0:End Sub

Sub sw73_Hit():Controller.Switch(73) = 1:End Sub
Sub sw73_UnHit():Controller.Switch(73) = 0:End Sub

Sub sw17_Hit():Controller.Switch(17) = 1:End Sub
Sub sw17_UnHit():Controller.Switch(17) = 0:End Sub

Sub sw58_Hit():vpmTimer.PulseSw 58:End Sub


' Castle Lock
Sub sw53_Hit():Controller.Switch(53) = 1:End Sub    'Castle Lock 1
Sub sw53_UnHit():Controller.Switch(53) = 0: WireRampOff: End Sub
Sub sw54_Hit():Controller.Switch(54) = 1:End Sub    'Castle Lock 2
Sub sw54_UnHit():Controller.Switch(54) = 0:End Sub
Sub sw57_Hit():Controller.Switch(57) = 1:End Sub    'Castle Lock 3
Sub sw57_UnHit():Controller.Switch(57) = 0:End Sub



'****************************
' ZINL: Inlane speed limit code
'****************************

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



'****************************
' ZDRN: Drain, Trough, and Ball Release
'****************************


Sub sw41_Hit():Controller.Switch(41) = 1:UpdateTrough:End Sub
Sub sw41_UnHit():Controller.Switch(41) = 0:UpdateTrough:End Sub
Sub sw42_Hit():Controller.Switch(42) = 1:UpdateTrough:End Sub
Sub sw42_UnHit():Controller.Switch(42) = 0:UpdateTrough:End Sub
Sub sw43_Hit():Controller.Switch(43) = 1:UpdateTrough:End Sub
Sub sw43_UnHit():Controller.Switch(43) = 0:UpdateTrough:End Sub
Sub sw44_Hit():Controller.Switch(44) = 1:UpdateTrough:End Sub
Sub sw44_UnHit():Controller.Switch(44) = 0:UpdateTrough:End Sub


Sub UpdateTrough
  UpdateTroughTimer.Interval = 100
  UpdateTroughTimer.Enabled = 1
End Sub

Sub UpdateTroughTimer_Timer
  If sw41.BallCntOver = 0 Then sw42.kick 57, 9
  If sw42.BallCntOver = 0 Then sw43.kick 57, 9
  If sw43.BallCntOver = 0 Then sw44.kick 57, 9
  Me.Enabled = 0
End Sub

Sub sw48_Hit()
  UpdateTrough
  vpmTimer.PulseSw 48
  vpmTimer.AddTimer 500, "sw48.kick 60, 20'"
  RandomSoundDrain sw48
End Sub

Sub SolRelease(enabled)
  If enabled Then
    'vpmTimer.PulseSw 17
    sw41.kick 45, 9
    RandomSoundBallRelease sw41
  End If
End Sub


'****************************
' ZVUK: VUKs
'****************************

Dim KickerBall55, KickerBall56, KickerBall71, KickerBall72


Sub KickBall(kball, kangle, kvel, kvelz, kzlift)
  dim rangle
  rangle = PI * (kangle - 90) / 180

  kball.z = kball.z + kzlift
  kball.velz = kvelz
  kball.velx = cos(rangle)*kvel
  kball.vely = sin(rangle)*kvel
End Sub


'Wireramp Popper
Sub sw55_Hit
  set KickerBall55 = activeball
  Controller.Switch(55) = 1
    SoundSaucerLock
End Sub

Sub sw55_UnHit
  Controller.Switch(55) = 0
End Sub

Sub WirerampPopper(Enable)
  If Enable then
    If Controller.Switch(55) <> 0 Then
      KickBall KickerBall55, 0, 0, 35, 50
      SoundSaucerKick 1, sw55
    Else
      SoundSaucerKick 0, sw55
    End If
  End If
End Sub



'Crypt Popper
Sub sw56_Hit
  set KickerBall56 = activeball
  Controller.Switch(56) = 1
  mMagnet.RemoveBall ActiveBall
  SoundSaucerLock
End Sub

Sub sw56_UnHit
  Controller.Switch(56) = 0
End Sub

Sub CryptPopper(Enable)
  If Enable then
    If Controller.Switch(56) <> 0 Then
      KickBall KickerBall56, 2*(rnd-0.5), 2*(rnd-0.5), 22+rnd*8, 0
      SoundSaucerKick 1, sw56
    Else
      SoundSaucerKick 0, sw56
    End If
  End If
End Sub



'Castle Popper
Sub sw71_Hit
  set KickerBall71 = activeball
  Controller.Switch(71) = 1
  mMagnet.RemoveBall ActiveBall
  SoundSaucerLock
End Sub

Sub sw71_UnHit
  Controller.Switch(71) = 0
End Sub

Sub CastlePopper(Enable)
  If Enable then
    If Controller.Switch(71) <> 0 Then
      KickBall KickerBall71, 23+rnd*2, 38+rnd*5, 0, 0
      SoundSaucerKick 1, sw71
    Else
      SoundSaucerKick 0, sw71
    End If
  End If
End Sub



'Coffin Popper
Sub sw72_Hit
  set KickerBall72 = activeball
  Controller.Switch(72) = 1
  SoundSaucerLock
End Sub

Sub sw72_UnHit
  Controller.Switch(72) = 0
End Sub

Sub CoffinPopper(Enable)
  If Enable then
    If Controller.Switch(72) <> 0 Then
      KickBall KickerBall72, 90, 0, 50, 0
      SoundSaucerKick 1, sw72
    Else
      SoundSaucerKick 0, sw72
    End If
  End If
End Sub



'****************************
' ZMST: Mist Magnet
'****************************
'
'          Mist Magnet
'    taken from Lander's table
' with only a small modification
'-------------------------------
' Magnet Simulator Class
' (07/10/2001 Dorsola)
' Modified for Dracula (08/16/2001) by Dorsola
'-------------------------------

class cMagnet
  Private cX, cY, cStrength, cRange
  private cBalls, cClaimed
  private cTempX, cTempY

  Private Sub Class_Initialize()
      set cBalls = CreateObject("Scripting.Dictionary")
      cRange = 1
      cStrength = 0
  End Sub

  Public Sub InitMagnet(aTrigger, inStrength)
      cX = aTrigger.X
      cY = aTrigger.Y
      cRange = aTrigger.Radius
      cStrength = inStrength
  End Sub

  Public Sub MoveTo(inX, inY)
      cX = inX
      cY = inY
  End Sub

  Public Property Get X:X = cX:End Property
  Public Property Get Y:Y = cY:End Property
  Public Property Get Strength:Strength = cStrength:End Property
  Public Property Get Size:Size = cRange:End Property
  Public Property Get Range:Range = cRange:End Property
  Public Property Get Balls:Balls = cBalls.Keys:End Property

  Public Property Let X(inX):cX = inX:End Property
  Public Property Let Y(inY):cY = inY:End Property
  Public Property Let Strength(inStrength):cStrength = inStrength:End Property
  Public Property Let Size(inSize):cRange = inSize:End Property
  Public Property Let Range(inSize):cRange = inSize:End Property

  Public Sub AddBall(aBall)
      cBalls.Item(aBall) = 0
  End Sub

  Public Sub RemoveBall(aBall)
      ' This function tags balls for removal, but does not remove them.
      ' Another sub will be called to remove tagged objects from the dictionary.
      If cBalls.Exists(aBall) then
          if cClaimed then
              cBalls.Item(aBall) = 1
          else
              cBalls.Remove(aBall)
          end if
      end if
  End Sub

  Public Sub Claim():cClaimed = True:End Sub

  Public Sub Release()
      cClaimed = False
      Dim tempobj
      for each tempobj in cBalls.Keys
          if cBalls.Item(tempobj) = 1 then cBalls.Remove(tempobj)
      next
  End Sub

  Public Sub ProcessBalls()
      Dim tempObj
      for each tempObj in cBalls.Keys:AttractBall tempObj:next
  End Sub

  Public Function GetDist(aBall)
      on error resume next
      if aBall is Nothing then
          GetDist = 100000
      else
          cTempX = aBall.X - cX
          cTempY = aBall.Y - cY
          GetDist = Sqr(cTempX * cTempX + cTempY * cTempY)
          if Err then GetDist = 100000
      end if
  End Function

  Public Sub AttractBall(aBall)
      if aBall is Nothing then Exit Sub
      Dim Dist
      Dist = GetDist(aBall)
      if Dist> cRange then Exit Sub

      ' Attract ball toward magnet center (cX,cY).

      ' Attraction force is determined by distance from center, and strength of magnet.

      Dim Force, Ratio
      Ratio = Dist / (1.5 * cRange)

      ' TODO: Figure out how to dampen the force when ball is near center and
      ' at low velocity, so that balls don't jitter on the magnets.
      ' Also shore up instability on moving magnet.

      Force = cStrength * exp(-0.2 / Ratio) / (Ratio * Ratio * 56)
      aBall.VelX = (aBall.VelX - cTempX * Force / Dist) * 0.985
      aBall.VelY = (aBall.VelY - cTempY * Force / Dist) * 0.985
  End Sub
End Class

'-----------------------------------------------
' Mist Multiball - courtesy of Dorsola
'-----------------------------------------------

' Method: Since any ball that can block the Mist opto is necessarily in the Mist Magnet's trigger area,
' we automatically have access to all balls in this range.  We can therefore check each ball's position
' against a line equation and see if it happens to be blocking the opto, and set the switch accordingly.
' This requires a timer loop.

Dim MagnetOn
MagnetOn = false

Sub MistTimer_Timer()
  ' Endpoints of the line are (126,1220) and (870,885)
  ' Slope: m = (y2-y1)/(x2-x1) = -0.45
  ' Y-intercept: b = y1 - m*x1 = 1277

  mMagnet.Claim

  Dim obj, CheckState, x, TargetY
  CheckState = 0
  on error resume next
  for each obj in mMagnet.Balls
      ' y = mx+b (m=slope, b=yint)
      TargetY = (-0.45) * obj.X +  1277
      if(obj.Y> TargetY - 60) and(obj.Y <TargetY + 60) then CheckState = 1' :magnet_marker001.x = motorx1:magnet_marker001.y = motory1 + 25:magnet_marker002.x = motorx1:magnet_marker002.y = motory1 - 25:magnet_marker003.x = motorx2:magnet_marker003.y = motory2 + 25:magnet_marker004.x = motorx2:magnet_marker004.y = motory2 + 25
  next
  on error goto 0

  Controller.Switch(82) = CheckState

  if MagnetOn then mMagnet.ProcessBalls

  mMagnet.Release
End Sub

Sub SolMistMagnet(enabled)
  MagnetOn = enabled
End Sub

Sub Magnet_Hit()
  mMagnet.AddBall ActiveBall
End Sub

Sub Magnet_UnHit()
  mMagnet.RemoveBall ActiveBall
End Sub

'------------------------
' Handle the Mist Motor
'------------------------
' Method: Treat motor's position as a number from right to left (0-500)
' and compute its position based on the line equation given above.

const motorx1 = 850
const motorx2 = 126
const motorxrange = 739
const motory1 = 885
const motory2 = 1220
const motoryrange = -335
const motorslope = -0.44
const motoryint = 1238

' Endpoints of the line are (126,1220) and (870,885)
' Slope: m = (y2-y1)/(x2-x1) = -0.45
' Y-intercept: b = y1 - m*x1 = 1277

Dim MagnetPos, MagnetDir
MagnetPos = 0:MagnetDir = 0

' Coding for MagnetDir: 0 = left, 1 = right, toggle at endpoints.

Sub MotorTimer_Timer()

  if Controller.Solenoid(28) then
    if MagnetDir = 0 then
      'MagnetPos = MagnetPos + 1
      MagnetPos = MagnetPos + .4 ' Thalamus
      'mist lights
      Select Case MagnetPos \ 33
        Case 0:ml1.State = 0:ml2.State = 0:ml3.State = 0:ml4.State = 0
          ml5.State = 0:ml6.State = 0:ml7.State = 0:ml8.State = 0
          ml9.State = 0:ml10.State = 0:ml11.State = 0:ml12.State = 0:ml13.State = 0
        Case 1:ml13.State = 1
        Case 2:ml12.State = 1:ml13.State = 0
        Case 3:ml11.State = 1:ml12.State = 0
        Case 4:ml10.State = 1:ml11.State = 0
        Case 5:ml9.State = 1:ml10.State = 0
        Case 6:ml8.State = 1:ml9.State = 0
        Case 7:ml7.State = 1:ml8.State = 0
        Case 8:ml6.State = 1:ml7.State = 0
        Case 9:ml5.State = 1:ml6.State = 0
        Case 10:ml4.State = 1:ml5.State = 0
        Case 11:ml3.State = 1:ml4.State = 0
        Case 12:ml2.State = 1:ml3.State = 0
        Case 13:ml1.State = 1:ml2.State = 0
        Case 14:ml1.State = 0
      End Select

      if MagnetPos >= 500 then
        MagnetPos = 500
        MagnetDir = 1
      end if
    else
      'MagnetPos = MagnetPos - 1
      MagnetPos = MagnetPos - .4 ' Thalamus
      Select Case MagnetPos \ 33
        Case 0:ml1.State = 0:ml2.State = 0:ml3.State = 0:ml4.State = 0
          ml5.State = 0:ml6.State = 0:ml7.State = 0:ml8.State = 0
          ml9.State = 0:ml10.State = 0:ml11.State = 0:ml12.State = 0:ml13.State = 0
        Case 1:ml13.State = 0
        Case 2:ml12.State = 0:ml13.State = 1
        Case 3:ml11.State = 0:ml12.State = 1
        Case 4:ml10.State = 0:ml11.State = 1
        Case 5:ml9.State = 0:ml10.State = 1
        Case 6:ml8.State = 0:ml9.State = 1
        Case 7:ml7.State = 0:ml8.State = 1
        Case 8:ml6.State = 0:ml7.State = 1
        Case 9:ml5.State = 0:ml6.State = 1
        Case 10:ml4.State = 0:ml5.State = 1
        Case 11:ml3.State = 0:ml4.State = 1
        Case 12:ml2.State = 0:ml3.State = 1
        Case 13:ml1.State = 0:ml2.State = 1
        Case 14:ml1.State = 1
      End Select
      if MagnetPos <= 0 then
        MagnetPos = 0
        MagnetDir = 0
      end if
    end if

    SetMagnetPosition
    Controller.Switch(81) = (MagnetPos> 490)
    Controller.Switch(83) = (MagnetPos <10)
  end if
End Sub

Sub SetMagnetPosition()
  mMagnet.X = motorx1 -(motorxrange * (MagnetPos / 500) )
  mMagnet.Y = motorslope * mMagnet.X + motoryint
  If MagnetPos MOD 33 = 0 Then
    MotorTimer.Interval = 80
  Else
    MotorTimer.Interval = 8
  End If
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
' ZFTR:  Flipper Tricks
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

    If Flipper2.currentangle = EndAngle2 Then
      For b = 0 To UBound(gBOT)
        If FlipperTrigger(gBOT(b).x, gBOT(b).y, Flipper1) Then
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


' End - Check ball distance from Flipper for Rem
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
Const EOSTnew = 1.2 '90's and later
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
    Dim b
    For b = 0 To UBound(gBOT)
      If Distance(gBOT(b).x, gBOT(b).y, Flipper.x, Flipper.y) < 57 Then 'check for cradle
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
' ZNFF:  FLIPPER CORRECTIONS by nFozzy
'******************************************************

Dim LF : Set LF = New FlipperPolarity
Dim RF : Set RF = New FlipperPolarity

InitPolarity

''*******************************************
''  Late 80's early 90's

Sub InitPolarity()
  dim x, a : a = Array(LF, RF)
  for each x in a
    x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
    x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
    x.enabled = True
    x.TimeDelay = 60
    x.DebugOn=False ' prints some info in debugger

    ' From example table v1.5.12
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

    ' From Fish Tales v1.0
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
    x.AddPt "Velocity", 5, 0.53, 0.95
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
      'debug.print "PolarityCorrect" & " " & GameTime & " " & Round(BallPos*100) & "%" & " AddX:" & Round(AddX,2) & " Vel%:" & Round(VelCoef*100) & " " & NoCorrection & " " & PartialFlipcoef
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
'   ZDMP: RUBBER DAMPENERS
'******************************************************
'
' These are data mined bounce curves,
' dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
' Requires tracking ballspeed to calculate COR

Sub dPosts_Hit(idx)
  RubbersD.dampen Activeball
  TargetBouncer Activeball, 1
End Sub

Sub dSleeves_Hit(idx)
  SleevesD.Dampen Activeball
End Sub

Sub dPost_leftsling_hit: RubbersD.dampen Activeball: End Sub
Sub dPost_rightsling_hit: RubbersD.dampen Activeball: End Sub

dim RubbersD : Set RubbersD = new Dampener        'frubber
RubbersD.name = "Rubbers"
RubbersD.debugOn = False        'shows info in textbox "TBPout"
RubbersD.Print = False        'debug, reports in debugger (in vel, out cor)
'cor bounce curve (linear)
'for best results, try to match in-game velocity as closely as possible to the desired curve
'RubbersD.addpoint 0, 0, 0.935        'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 0, 0, 1.1        'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 1, 3.77, 0.97
RubbersD.addpoint 2, 5.76, 0.967        'dont take this as gospel. if you can data mine rubber elasticitiy, please help!
RubbersD.addpoint 3, 15.84, 0.874
RubbersD.addpoint 4, 56, 0.64        'there's clamping so interpolate up to 56 at least

dim SleevesD : Set SleevesD = new Dampener        'this is just rubber but cut down to 85%...
SleevesD.name = "Sleeves"
SleevesD.debugOn = False        'shows info in textbox "TBPout"
SleevesD.Print = False        'debug, reports in debugger (in vel, out cor)
SleevesD.CopyCoef RubbersD, 0.85

'######################### Add new FlippersD Profile
'#########################    Adjust these values to increase or lessen the elasticity

dim FlippersD : Set FlippersD = new Dampener
FlippersD.name = "Flippers"
FlippersD.debugOn = False
FlippersD.Print = False
FlippersD.addpoint 0, 0, 1.1
FlippersD.addpoint 1, 3.77, 0.99
FlippersD.addpoint 2, 6, 0.99

Class Dampener
  Public Print, debugOn 'tbpOut.text
  public name, Threshold         'Minimum threshold. Useful for Flippers, which don't have a hit threshold.
  Public ModIn, ModOut
  Private Sub Class_Initialize : redim ModIn(0) : redim Modout(0): End Sub

  Public Sub AddPoint(aIdx, aX, aY)
    ShuffleArrays ModIn, ModOut, 1 : ModIn(aIDX) = aX : ModOut(aIDX) = aY : ShuffleArrays ModIn, ModOut, 0
    if gametime > 100 then Report
  End Sub

  public sub Dampen(aBall)
    if threshold then if BallSpeed(aBall) < threshold then exit sub end if end if
    dim RealCOR, DesiredCOR, str, coef
    DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut ) + (rnd-0.5)*0.04  'add some randomness to the COR (apophis)
    RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id)+0.0001)
    coef = desiredcor / realcor
    if debugOn then str = name & " in vel:" & round(cor.ballvel(aBall.id),2 ) & vbnewline & "desired cor: " & round(desiredcor,4) & vbnewline & _
    "actual cor: " & round(realCOR,4) & vbnewline & "ballspeed coef: " & round(coef, 3) & vbnewline
    if Print then debug.print Round(cor.ballvel(aBall.id),2) & ", " & round(desiredcor,3)

    aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
    if debugOn then TBPout.text = str
  End Sub

  public sub Dampenf(aBall, parm) 'Rubberizer is handle here
    dim RealCOR, DesiredCOR, str, coef
    DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
    RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id)+0.0001)
    coef = desiredcor / realcor
    If abs(aball.velx) < 2 and aball.vely < 0 and aball.vely > -3.75 then
      aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
    End If
  End Sub

  Public Sub CopyCoef(aObj, aCoef) 'alternative addpoints, copy with coef
    dim x : for x = 0 to uBound(aObj.ModIn)
      addpoint x, aObj.ModIn(x), aObj.ModOut(x)*aCoef
    Next
  End Sub


  Public Sub Report()         'debug, reports all coords in tbPL.text
    if not debugOn then exit sub
    dim a1, a2 : a1 = ModIn : a2 = ModOut
    dim str, x : for x = 0 to uBound(a1) : str = str & x & ": " & round(a1(x),4) & ", " & round(a2(x),4) & vbnewline : next
    TBPout.text = str
  End Sub

End Class



'******************************************************
'****  END PHYSICS DAMPENERS
'******************************************************



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


'******************************************************
'   ZBOU: VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************

Const TargetBouncerEnabled = 1    '0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.9   'Level of bounces. Recommmended value of 0.7

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
'   ZRRL: RAMP ROLLING SFX
'******************************************************

' Ramp sound effect triggers

Sub RampTrigger001_Hit
  If RightRamp.Collidable = True Then WireRampOn True  'plastic ramp
End Sub

Sub RampTrigger001_UnHit
  If activeball.vely > 0 Then WireRampOff
End Sub

Sub RampTrigger002_Hit
  WireRampOn True  'plastic ramp
End Sub

Sub RampTrigger002_UnHit
  If activeball.vely > 0 Then WireRampOff
End Sub

'Sub RampTrigger003_Hit  'Replaced with sw85
' WireRampOff
'End Sub

Sub RampTrigger003_UnHit
  WireRampOn False  'wire ramp
End Sub

Sub RampTrigger004_Hit
  WireRampOff
End Sub

Sub RampTrigger005_Hit
  WireRampOff
End Sub

Sub RampTrigger006_Hit
  WireRampOff
End Sub

Sub RampTrigger007_Hit
  WireRampOn False  'wire ramp
End Sub

Sub RampTrigger008_Hit
  If activeball.z > 35 Then
    WireRampOn True  'plastic ramp
  End If
End Sub

Sub RampTrigger009_UnHit
  WireRampOff
End Sub

Sub RampTrigger010_UnHit
  WireRampOn False  'wire ramp
End Sub

Sub RampTrigger011_UnHit
  WireRampOn False  'wire ramp
End Sub



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
Dim RampBalls(5,2)
'x,0 = ball x,1 = ID, 2 = Protection against ending early (minimum amount of updates)

'0,0 is boolean on/off, 0,1 unused for now
RampBalls(0,0) = False

' RampType
' Setup: Set this array to the number Total number of balls that can be tracked at one time + 1.  5 ball multiball then set value to 6
' Description: Array type indexed on BallId and a values used to deterimine what type of ramp the ball is on: False = Wire Ramp, True = Plastic Ramp
Dim RampType(5)

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





'******************************************************
' ZSFX: Mechanical Sound effects
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
  PlaySound soundname, 1,min(1,aVol) * VolumeDial, AudioPan(ActiveBall), 0,0,0, 1, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtBallVolM (Soundname, aVol)
  PlaySound soundname, 1,min(1,aVol) * VolumeDial, AudioPan(ActiveBall), 0,0,0, 0, AudioFade(ActiveBall)
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

'/////////////////////////////  COIN IN  ////////////////////////////

Sub SoundCoinIn
  Select Case Int(rnd*3)
    Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
    Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
    Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
  End Select
End Sub

'/////////////////////////////////////////////////////////////////
'         End Mechanical Sounds
'/////////////////////////////////////////////////////////////////





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
  '   Dim BOT
  '   BOT = GetBalls

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




'*******************************************
'   ZABS: Ambient ball shadows
'*******************************************

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
    objBallShadow(iii).Z = 1 + iii / 1000 + 1.04
    objBallShadow(iii).visible = 0
  Next
End Sub


Sub BSUpdate
  Dim s

  'The Magic happens now
  For s = lob To UBound(gBOT)
    ' *** Normal "ambient light" ball shadow

    'Primitive shadow on playfield, flasher shadow in ramps
    '** If on main pf
    If gBOT(s).Z  < 30 and gBOT(s).Z  > 20 Then
      objBallShadow(s).visible = 1
      objBallShadow(s).X = gBOT(s).X + (gBOT(s).X - (tablewidth / 2)) / (Ballsize / AmbientMovement) + offsetX
      objBallShadow(s).Y = gBOT(s).Y + offsetY
      'objBallShadow(s).Z = gBOT(s).Z + s/1000 + 1.04 - 25

    '** Under and over pf, no shadow
    Else
      objBallShadow(s).visible = 0
    End If
  Next
End Sub







'*******************************************
' ZSTA: Stand-up Targets
'*******************************************

'sw 66, 67, 68, 86, 87, 88

' Standup Targets
Sub sw66_Hit(): STHit 66: End Sub
Sub sw67_Hit(): STHit 67: End Sub
Sub sw68_Hit(): STHit 68: End Sub
Sub sw86_Hit(): STHit 86: End Sub
Sub sw87_Hit(): STHit 87: End Sub
Sub sw88_Hit(): STHit 88: End Sub


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
Dim ST66, ST67, ST68, ST86, ST87, ST88

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
'You will also need to add a secondary hit object for each stand up (name sw11o, sw12o, and sw13o on the example Table)
'these are inclined primitives to simulate hitting a bent target and should provide so z velocity on high speed impacts


Set ST66 =  (new StandupTarget)(sw66, BM_sw66, 66, 0)
Set ST67 =  (new StandupTarget)(sw67, BM_sw67, 67, 0)
Set ST68 =  (new StandupTarget)(sw68, BM_sw68, 68, 0)
Set ST86 =  (new StandupTarget)(sw86, BM_sw86, 86, 0)
Set ST87 =  (new StandupTarget)(sw87, BM_sw87, 87, 0)
Set ST88 =  (new StandupTarget)(sw88, BM_sw88, 88, 0)


'Add all the Stand-up Target Arrays to Stand-up Target Animation Array
' STAnimationArray = Array(ST1, ST2, ....)
Dim STArray
STArray = Array(ST66, ST67, ST68, ST86, ST87, ST88)

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

    ty = BM_sw66.transy
  For each BP in BP_sw66 : BP.transy = ty: Next

    ty = BM_sw67.transy
  For each BP in BP_sw67 : BP.transy = ty: Next

    ty = BM_sw68.transy
  For each BP in BP_sw68 : BP.transy = ty: Next

    ty = BM_sw86.transy
  For each BP in BP_sw86 : BP.transy = ty: Next

    ty = BM_sw87.transy
  For each BP in BP_sw87 : BP.transy = ty: Next

    ty = BM_sw88.transy
  For each BP in BP_sw88 : BP.transy = ty: Next

End Sub


'*******************************************
' ZDTA: Drop Targets
'*******************************************

'sw 15

Sub sw15_Hit(): DTHit 15: End Sub



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
Dim DT15

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

Set DT15 = (new DropTarget)(sw15, sw15a, BM_sw15, 15, 0, false)


Dim DTArray
DTArray = Array(DT15)

'Configure the behavior of Drop Targets.
Const DTDropSpeed = 80 'in milliseconds
Const DTDropUpSpeed = 40 'in milliseconds
Const DTDropUnits = 55 'VP units primitive drops so top of at or below the playfield
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
      controller.Switch(Switchid) = 1
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
      'Dim gBOT: gBOT = GetBalls

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



Sub UpdateDropTargets
  dim BP, tz, rx, ry

    tz = BM_sw15.transz
  rx = BM_sw15.rotx
  ry = BM_sw15.roty
  For each BP in BP_sw15 : BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next
End Sub




'*******************************************
'   ZANI: Misc Animations
'*******************************************


' Bumpers

Sub Bumper1_Animate
  Dim z, BP
  z = Bumper1.CurrentRingOffset
  For Each BP in BP_Bumper1_Ring : BP.transz = z: Next
End Sub

Sub Bumper2_Animate
  Dim z, BP
  z = Bumper2.CurrentRingOffset
  For Each BP in BP_Bumper2_Ring : BP.transz = z: Next
End Sub

Sub Bumper3_Animate
  Dim z, BP
  z = Bumper3.CurrentRingOffset
  For Each BP in BP_Bumper3_Ring : BP.transz = z: Next
End Sub


Dim Bumpers : Bumpers = Array(Bumper1, Bumper2, Bumper3)

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
''    If r = 0 Then For Each x in BP_Bumper1_Socket: x.Z = z: Next
''    If r = 1 Then For Each x in BP_Bumper2_Socket: x.Z = z: Next
''    If r = 2 Then For Each x in BP_Bumper3_Socket: x.Z = z: Next
' Next
'End Sub


' Gates

Sub Gate4_Animate
  Dim a: a = Gate4.CurrentAngle
  Dim BP : For Each BP in BP_Gate4: BP.RotX = a: Next
End Sub

Sub Gate6_Animate
  Dim a: a = Gate6.CurrentAngle
  Dim BP : For Each BP in BP_Gate6: BP.RotX = a: Next
End Sub

Sub Lgate_Animate
  Dim a: a = Lgate.CurrentAngle
  Dim BP : For Each BP in BP_Lgate: BP.RotX = a: Next
End Sub

Sub Rgate_Animate
  Dim a: a = Rgate.CurrentAngle
  Dim BP : For Each BP in BP_Rgate: BP.RotX = a: Next
End Sub


' Switches

Sub sw16_Animate
  Dim z : z = sw16.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw16 : BP.transz = z: Next
End Sub

Sub sw25_Animate
  Dim z : z = sw25.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw25 : BP.transz = z: Next
End Sub

Sub sw26_Animate
  Dim z : z = sw26.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw26 : BP.transz = z: Next
End Sub

Sub sw27_Animate
  Dim z : z = sw27.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw27 : BP.transz = z: Next
End Sub

Sub sw28_Animate 'on ramp
  Dim z : z = sw28.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw28 : BP.transz = -z: Next
End Sub

Sub sw31_Animate
  Dim z : z = sw31.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw31 : BP.transz = z: Next
End Sub

Sub sw35_Animate
  Dim z : z = sw35.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw35 : BP.transz = z: Next
End Sub

Sub sw36_Animate
  Dim z : z = sw36.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw36 : BP.transz = z: Next
End Sub

Sub sw37_Animate
  Dim z : z = sw37.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw37 : BP.transz = z: Next
End Sub

Sub sw38_Animate
  Dim z : z = sw38.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw38 : BP.transz = z: Next
End Sub

Sub sw84_Animate 'on ramp
  Dim z : z = sw84.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw84 : BP.transz = -z: Next
End Sub

Sub sw85_Animate 'castle lock off ramp
  Dim z : z = sw85.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw85 : BP.rotZ = z * 3: Next
End Sub

Sub l87_Animate
  LaunchBut.blenddisablelighting = l87.GetInPlayIntensity * 1
End Sub

Sub l88_Animate
  StartBut.blenddisablelighting = l88.GetInPlayIntensity * 2
End Sub

' Diverter

Sub FDivert_animate
  Dim a: a = FDivert.CurrentAngle
  Dim max_angle, min_angle, mid_angle
  max_angle = FDivert.StartAngle
  min_angle = FDivert.EndAngle
  mid_angle = (max_angle-min_angle)/2 + min_angle ' bake map switch point angle

  Dim BP
  For Each BP in BP_Diverter
    BP.transy = a
    BP.visible = a > mid_angle
  Next
  For Each BP in BP_DiverterU
    BP.transy = a
    BP.visible = a < mid_angle
  Next
End Sub


' Ramps

Sub FLRamp_animate
  Dim a: a = FLRamp.CurrentAngle
  Dim max_angle, min_angle, mid_angle
  max_angle = FLRamp.StartAngle
  min_angle = FLRamp.EndAngle
  mid_angle = (max_angle-min_angle)/2 + min_angle ' bake map switch point angle

  Dim BP
  For Each BP in BP_LRamp_DP
    BP.objrotx = a
    BP.visible = a > mid_angle
  Next
  For Each BP in BP_LRamp_UP
    BP.objrotx = a
    BP.visible = a < mid_angle
  Next
End Sub

Sub FRRamp_animate
  Dim a: a = FRRamp.CurrentAngle
  Dim max_angle, min_angle, mid_angle
  max_angle = FRRamp.StartAngle
  min_angle = FRRamp.EndAngle
  mid_angle = (max_angle-min_angle)/2 + min_angle ' bake map switch point angle

  Dim BP
  For Each BP in BP_RRamp_DP
    BP.objrotx = a
    BP.visible = a > mid_angle
  Next
  For Each BP in BP_RRamp_UP
    BP.objrotx = a
    BP.visible = a < mid_angle
  Next
End Sub





' Flipper bake maps and shadows

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




'*******************************************
' ZVRR: VR Room Stuff
'*******************************************


Sub SetupRoom
  DIM VRThings
  CandleFlicker.Enabled = 0
  If VRRoom > 0 Then
    for each VRThings in VR_Cab:VRThings.visible = 1:Next
    If VRRoom = 1 Then
      for each VRThings in VR_Min:VRThings.visible = 1:Next
      for each VRThings in VRMegaRoom:VRThings.visible = 0:Next
      For each VRThings in BGT_All:VRThings.Visible = 0: Next
      For each VRThings in BGB_All:VRThings.Visible = 1: Next
      If VRPosterLeft > 0 Then PosterLeft.Visible = 1 Else PosterLeft.Visible = 0
      If VRPosterRight > 0 Then PosterRight.Visible = 1 Else PosterRight.Visible = 0
    End If
    If VRRoom = 2 Then
      CandleFlicker.Enabled = 1
      for each VRThings in VR_Min:VRThings.visible = 0:Next
      for each VRThings in VRMegaRoom:VRThings.visible = 1:Next
      For each VRThings in BGT_All:VRThings.Visible = 1: Next
      For each VRThings in BGB_All:VRThings.Visible = 1: Next
'     PinCab_Metal_Rear.visible=False
    End If
  Else
    for each VRThings in VR_Cab:VRThings.visible = 0:Next
    for each VRThings in VRMegaRoom:VRThings.visible = 0:Next
    for each VRThings in VR_Min:VRThings.visible = 0:Next
    For each VRThings in BGT_All:VRThings.Visible = 0: Next
    For each VRThings in BGB_All:VRThings.Visible = 0: Next
  End If
End Sub



' Candle flame animations

InitTopper
Sub InitTopper
  Dim BP
  For each BP in BP_TWall: BP.transy = -2: BP.transx = 5: BP.Size_x = 0.99: Next
  For each BP in BP_THead: BP.transy = -2: Next
End Sub

Const FlameOp = 200
Const FlickerAmp = 0.3
Const FlickerThresh = 0.06

Sub CandleFlicker_Timer
  UpdateCandle fla1,BP_TFlame1,0
  UpdateCandle fla2,BP_TFlame2,0
  UpdateCandle fla3,BP_TFlame3,1
  UpdateCandle fla4,BP_TFlame4,2
  UpdateCandle fla5,BP_TFlame5,1
  UpdateCandle fla6,BP_TFlame6,2
End Sub

Sub UpdateCandle(light,BP_Array,gistring)
  dim f0, f1, gistate, BP
  gistate = 1
  Select Case gistring
    Case 0: if gi0lvl < 0.3 then gistate = 0
    Case 1: if gi1lvl < 0.3 then gistate = 0
    Case 2: if gi2lvl < 0.3 then gistate = 0
  End Select
  If gistate = 1 Then
    f0 = light.GetInPlayIntensity / light.Intensity
    f1 = (1-FlickerAmp) + FlickerAmp*rnd
    If abs(f1-f0) > FlickerThresh Then
      light.state = f1
      For each BP in BP_Array
        BP.roty = 20*(f1-f0)
        BP.rotx = 20*(f1-f0)
      Next
    End If
  Else
    light.state = 0
  End If
End Sub

Sub AnimateCandle(light,BP_Array)
  dim f, BP
  f = light.GetInPlayIntensity / light.Intensity
  For each BP in BP_Array
    If f > 0.3 Then
      BP.opacity = FlameOp*(0.9 + 0.1)*f
    Else
      BP.opacity = 0
    End If
    BP.size_z = 0.6 + 0.4*f
    BP.size_x = 1 + 0.2*(1-f)
    BP.size_y = 1 + 0.2*(1-f)
    BP.roty = BP.roty*0.9
    BP.rotx = BP.rotx*0.9
  Next
End Sub

Sub fla1_animate: AnimateCandle fla1,BP_TFlame1: End Sub
Sub fla2_animate: AnimateCandle fla2,BP_TFlame2: End Sub
Sub fla3_animate: AnimateCandle fla3,BP_TFlame3: End Sub
Sub fla4_animate: AnimateCandle fla4,BP_TFlame4: End Sub
Sub fla5_animate: AnimateCandle fla5,BP_TFlame5: End Sub
Sub fla6_animate: AnimateCandle fla6,BP_TFlame6: End Sub



' Red eye blink

dim bEyesBlinking: bEyesBlinking = False
Sub EyeBlink_Timer
  If bEyesBlinking = False Then
    bEyesBlinking = True
    EyeBlink.Interval = 80
    Eyes.visible = False
  Else
    bEyesBlinking = False
    EyeBlink.Interval = RndInt(1000,15000)
    Eyes.visible = True
  End If
End Sub



'*******************************************
' ZCHL: Change Log
'*******************************************


'004 - stavcas - new pf scans, reorganized rubbers and added physics prims accordingly
'005 - stavcas - fixed ramps
'007 - stavcas - fixed kicker ang triggers locations
'008 - stavcas - lights and GI
'009 - stavcas - ramp prims and some big misc prims relocated
'014 - stavcas - worked on right ramp (solved thanks to fluffhead!)
'015 - stavcas - cleaned script to most basic.
'016 - Sixtoe - Added playfield mesh, added subways, added and repositioned walls (fixed the shooter into mist thing, still plays up a little), redid all ramps especially ramp interfaces, removed old script and assets, ramps prepared for new kickers and castle multiball opto setup, table quick and dirty relight, fixed flippers opening mist gates, probably some other stuff
'017 - stavcas - added trough in table init, updated sw names, added release and drain subs
'018 - Sixtoe - Fixed trough, mist preload, solonoids and kickers hooked up to new code, castle lock switches and post added, some other things
'019 - stavcas - fixed castle VUK,WireRamp popper VUK - added similar "periscope" to guide the ball onto the ramp (accroding to example from LOTR by VPW)
'020 - stavcas - adjusted trigger areas, added nfozzy flipper subs and corrections, dPosts and dSleeves collections
'021 - stavcas - added missing TargetBouncer sub
'022 - stavcas - added Fleep Mech sounds
'023 - stavcas - fixed all VUKs, the periscope on the wire ramp VUK needs to be rotated, a few holes were fixed
'024 - stavcas - added insert text and wired first prim
'025 - stavcas - add wall to fix ball disapearing on the top right on launch
'026 - stavcas - Added Lampz
'027 - apophis - Fixed some physics materials. Removed a bunch of duplicate physics objects. Updated Fleep code. Added sling corrections. Worked on kickers and VUks (still needs work).
'028 - apophis - More physics tuning (incl kickers). Tried (and failed) to fixed coffin vuk ramp. Added coffin lock deflector. Shortened flipper bats to correct length. Fixed KnockerPosition error. Other stuff.
'029 - apophis - Fixed mist magnet. Fixed coffin vuk. The coffin lock pin mech still needs to be built.
'030 - Sixtoe - Layout reinforced and tidied up, 1st Pass.
'031 - Sixtoe - Fixed layout issue I caused, added deflector plate to left hole, raised physics object posts to 25, fixed ramp entrances, changed main ramp as it's wrong, fixed ramp lock issue caused by new ramp position, made lots of things invisible, changed plastic ramp right exit hole, tweaked target locations, tweaked physics.
'032 - apophis - Updated hole deflectors. PF slope set to 6 deg. PF fric to 0.25. Fixed stuck subway ball. Fixed castle lock post timing. Cleaned up some sound effects. Added ramp rolling sounds. Other minor physics tweaks.
'033 - Stavcas - Added bumpers
'034 - Sixtoe - New playfield from CK
'035 - Sixtoe - Changed dimensions of playfield, stripped and rebuilt table, still in progress
'036 - apophis - Massive script reorg and cleanup. Added: new PWM stuff, Tweak menu, Desat LUTs, Room and ball brightness code, new ball image, ambient shadows, flipper shadows. Updated flippers, corrections, and tricks (per Fish Tales). Minor phyiscal wall tweaks.
'037 - Sixtoe - Reprofiled ramps
'038 - apophis - More randomness on asylum vuk. Fixed Gate6 position. Mist magnet strength set to 3. Faster bumper ring speed, strength at 15. Slings strength at 4. Sleeve friction set to 0.3
'039 - apophis - Adjusted center nudge. Adjusted inlane elasticity.
'040 - apophis - Updated CheckLiveCatch. Reduced mist mag strength to 1. Tweaked right plunger ramp geometry.
'041 - apophis - Timer cleanup. Included UpdateBallBrightness. Flipper trigger shape tweak.
'042 - apophis - Uses new core.vbs solenoid callback that support FastFlips and avoids stutters. Requires VPX 10.8.0-1904, VPM 3.6-862 (or later)
'043 - apophis - Revert SolCallback2 as it now has a better solution integrated in VPX
'044 - apophis - Corrected playfield dimensions and realigned everything.
'045 - Sixtoe - Adjusted all the ramps to the new size, made new left kickout scoop
'046 - apophis - Added a wall to prevent plunged ball from clipping bottom of ramp over plunger lane (causing ball to slow down).
'047 - FrankEnstein - Added apron geo and textures. Added instruction/credits card. Added insert to back wall. Fixed standup targets size. Added standin: castle, graveyard and drac from the old table. Added bumper caps and bumper lights. Added green insert to back wall with decal. Added second level decal plastics on slings. Added inlane guard metal layer
'048 - apophis - Separate some of the GI lights off the 3 strings to get dynamic shadows.
'049 - FrankEnstein - Make sure there is a light inside the coffin exit hole. Make a mask object for the transparent bit of the movable coffin lock ramp. Make mask separators for the extra slingshot plastics. Separate all hardware off the transparent masks and make the transparent ramp layers non-opaque. Change the side wall decals and coffin decal to UV bake mode (might need a revisit of the UV layout if there are overlaps). Make top and bottom movable ramp positions and set to render separated with hide from others. Add side rails and lockbar. Hid vpx lights. Renamed GI lights to match lightmaps.
'050 - FrankEnstein - Added coffin decals (placeholder that will need to be revisited at some point). Added ramp decals. Added coffin tombstone frame. Replaced standup targets with library ones and assigned proper colors. Added top right corner decal. Added placeholder village model from Kevv’s scan (needs cleaner bottom scan). Added back wall and decal for it.
'054 - FrankEnstein - Added light inside coffin exit. Added translucence map to drac’s eyes inside the coffin. Added layer separators for all transparent objects. Separated ramps into bottom path area and curved area, and applied different VPX materials. Redid the UVs on all ramps for better coverage. Added side rails and lockbar. Added raytraced shadows to some major lights
'055 - FrankEnstein - Fixed the names for flasher 17 and 20 that were named 17_Dome and 20_Dome in blender. Made a tighter layer separator for the coffin lid. RRamp material made smoky transparent plastic. Applied the transforms to the launch ramp. Hope it fixes the skewed output for VPX. Fixed the lane guards targets
'060 - FrankEnstein - Updated the bumper to TZ setup. Fixed the coffin lid and using coffin vpx material. Optimized the bake but back to baking at 50%. Added under ramp pf GI. Removed L68 that was in middle of bumpers with no physical counterpart. Added glass dirt (optional via F12). Rails and lockbar removable (optional via F12). Refractions optional via F12. Reflections optional via F12 (will optimize later with subset array, once shapes settle down). Fixed inserts depth (finally!). Added help card to F12. GI Lights baked white
'061 - Sixtoe - New playfield mesh (old one's holes were too narrow), aadded L68 back to light layer, added moon GI to GI layer (mid-GI), completely redid physical ramps to match up with graphics ramps, switch heights need more work, drac movable ramp needs more work, getting caught on something
'062 - apophis - Options updates: Reflections, Desktop DMD vis and color, GI colors. Added animations: diverter, bumper rings & sockets, gates, switches, slings & slingarms, LRamp, RRamp, drop target, standup targets. Fixed right ramp blockage. Cleaned-up timers.
'065 - FrankEnstein - Added extra detail to coffin lock ramp and UV mapped all metal rails. Restored mist lights and 3 other under playfield lights (drac’s face, rats and jet bumpers center). Side guards updated material to match reference. Added level to right side guard. Decoupled the interior coffin lights from top GI (should respond to own triggers now). Added bumper skirt rotation script. Fixed the hanging switches direction and numbers on top right 84 vs 38. Baked at full rez. Brighter dracula inside coffin
'066 - apophis - Fixed sw85 location. Deleted unused spinner objects. Included AdjustBulbTint effect for flashers. Fixed mist ball magnet fails.
'067 - FrankEnstein - Moved sw73 to match the ramp opto.
'070 - apophis - Fixed some ramp rolling sounds. Add flasher color option, and removed the AdjustBulbTint effect. Added visual blocker under PF near mystery subway.
'071 - FrankEnstein - Adding animation for sw85 and moved the trigger to match the physical part location.  Adjusted the skill shot lights l61, l62 and l63 to match the graphic on the back wall decal. When extra plastics are hidden, support spacer are also removed and screws are translated down
'072 - DGrimmReaper - VR minimal setup
'073 - apophis - Added glass dirt opacity options. Made standup target backwalls invisible. Adjusted some VR room things. VR room/cab DB set to 0.
'074 - DGrimmReaper - Hooked up start button and Launch button lights
'075 - Sixtoe - Tweaked plunger area and dropped the entire thing down, made plunger non visible, tweaked outlane rubbers, decreased playfield friction to 0.23, increased slope to 6.25 as per manual, removed lip on entry ramp, removed old sound plunger lane ramp sound trigger and merged with other ramp sound trigger, added hard positions for left and right inlane rubbers but didn't hook them up.
'076 - FrankEnstein - Added 3 option for outlanes difficulty: Easy, Med, Hard with matching visuals and physics adjustments. Updated all sculpts with the latest ones from Kevv. Cleaned up scans, re-fitted and re-UV-ed for optimal texture resolution on visible parts and removed any geometry that would never be seen. Sculpts are baked unwrapped to please the VR peeping Toms. Optimized other maps to stay within budget with the added unwraps. Fixed floating back wall post
'077 - apophis - Put new outpost rubber posts in dPosts collection. Default dirty glass set to 50%.
'078 - FrankEnstein - Full reflections limited to Parts only (other prims don't contribute enough). one less refraction probe, reusing the bumpers one for ramps. Disabled ambocc.
'079 - FrankEnstein - Reduced the poly count for sculpts from 48k to 12k each. Removed some GI lights from RT set. Less vibrant blue on apron. Added texture to apron and interior sidewalls.
'080 - apophis - Increased bumper strength to 4.5. Decreased flipper strength to 3150. PF frction to 0.22. Increased targets elasticity to 0.75 and scatter angle to 2. Increased TargetBouncerFactor to 0.9. Added some randomness to dampener COR (+/- 0.02).
'081 - FrankEnstein - A better reflection selection function. Fixed layer1 refraction issue by making it additive (hack).
'082 - FrankEnstein - letters on the VIDEO inserts block a bit of light so they still read when lit. Apron blue rebalanced a bit more.
'083 - apophis - Added animated VR topper. Always on for now. Will hook up to VR room settings later.
'084 - apophis - Added mega room. Still a WIP, some things need visual tweaking.
'085 - apophis - Updated VR meta room and topper. Toned down flipper and sling strengths a little.
'086 - apophis - Added VR backglass and speaker panel. Baked the VR cab. Made topper head fangs more fangy. Updated VR minimal room. Removed unused images. Fixed stuck ball issue.
'087 - apophis - Fixed minimal room brightness. Red eyes down the hall now occationally blink. Make candle flame animation look more realistic.
'088 - apophis - Tweaked candle flame animation (again). Added movie poster options to the minimal room.
'089 - FrankEnstein - VR mega room has glass windows and vista view of unsuspecting village full of dinner
'090 - apophis - fixed the green tint in fl22/23 flashers
'091 - apophis - Enabled apron and inlane object hit events. Set mist mag strength to 0.8. Fixed RampTrigger001 logic. Fixed RampTrigger007 height.
'092 - FrankEnstein - Moved the flasher fix into the render, fixed the right sling rubber mesh
'RC1 - apophis - Updated table info and comments.

' Release 1.0

' LOVE NEVER DIES



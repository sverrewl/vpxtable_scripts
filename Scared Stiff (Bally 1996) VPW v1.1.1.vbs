'                                                                                                                                                  +@@@@#
'       @@@@        =-                                                                                                                   *@@@@@@#*+=-::=@@
'      @*  +#=@@@@@@#%@  @@@          @       .=                                                                                @@@@@@%+#@@@*::::::::::-@.@
'     @*              @@@@+-@%*%@@@@@#@@  #@@@%@@ @@@@@@@%                                                             @@@@@@@@*=-::::::%@@@+::::::::::-@ .@
'    @+               =@@*             @  @+    @@@*      :#@. @%@@@@@@@@@@@@@@@@@.                           @@@@@@@@*--::#@@@-::::::::*@ @+::::::::::=@  +#
'    @-         :     +@#              @ @#     @@@         .@ @:       @@@       *@                 @@@@@@%#*+-:-@@@#:::::*@:@-::::::::#@ @+::::::---+#@  #@
'   .*         *@@@@  %@+      @@      @@@=     -@@#         @@@@       @@*        :@       %@@@@@%#@@@=:::::::::-@ ##:::::*@.@=::::+#%@@@ *+:::::-@@    @ %@
'    *       @@@   # :@@@*     @+@@@  :@@+      @@@*   %@@    @@@:  @@@.@@@+  @-    @#   @@*--:::::-@@%-:::::::-=+@ :%:::::+@ @=::::@@=   #=*:::::-@*@    #@@
'    @@@.    .@    @ @  @*    +@   @= @+@#      @@@*   @ @@-  %@@  .@     @=  @@@*  @@  @#::::=*#%%@@@@@#@=:::-@%  %@%:::::=@.@=::::@# =  -@*::::::@: @@@@
'      @=     -@   %@@ @%.    @    @*.@ @+  @@  *@@.   @  @#  .@%   @@@@  @: -@ @%  +@  @*:::-@. -.:  @@@@+:::-@ *+-@@-::::+@ @=::::#@@@@@@@#::::::#@@@@
'      @+       @#     @+     @     @@#@%  @@@.  @@@   @  @*   @%      @ @%  =@  @  =@  @=::::::::::+@@@@@+:::-@  @@@%:::::+@ @=::::::::#@@@*::::::::::%@.
'      @%@@      .@   @@+     @        @*  @ @%  @@@=  @:@@+   @%   @@%@ -#  +@ @@. :@  @@@%@@#-::::+@ %@@*:::-@  @@@%:::::+@ @=::::::::%#:@*::::::::::@=@@
'          @.      %@  @@     @    :@  @-  @  @. =@@=  .*      @%   @@   @%  +@@@*  +@   @@@=-##::::+@  @@*:::-@  @@%%:::::=@ @+::::::::*@.@*::::::::::%. @@
'      *%@ @%@%      @: @*    @    @@@%#   @@%#   @@*        @@@%   @@@@@@@*  *.    @-  @*==+++-::::#@ -@@*:::-@  @@*@-::::=@ @*::::-@@@@@=**::::::-==-@.  @
'      %.:    :@=     @@@*    @    %=@@#          -@%   @@%  @@@#       @@@+       *@   @+::::::::::#@  @@+:::-@  *@+@:::::+@ @*:::::@:@  *@*:::::-@@   @  @
'     @=  @     @-     @@*    @    # @@*   :+@:%  -@@:  @ +@+:@@*     . @@%+@@-@@@@@     @@..*#@%***@# -@@*:::-@  @@-@-::::=@.%*:::::@  :@@@#::::::@-@   -%@
'    @%         %#     *@#    ++  @+ @@*  @    %: =@#: *   ## @@*:@@@@@@@@@@              @@@@*.       =@@@@%#*@  @@.@-::::=@ @*::::-@:  @@@#::::::%: @@@@@:
'    .@%.   #*   %.    + @     @%#=  @@+  @    @= -@@*@@    @#@@#@                                %@@@@@#  @@   =+@@ @@@@%**@ @#:::::@   @ :#::::::@. @@
'     @+     @@@@+     @ @=          @@+  %    @@@@ @@       @                                               .@@@@-   :@:    #%@*:%@@@@  @  @-:::::@  @@
'     @=               @ @+        @@@@=*@                                                                                @@@@@@@@     @ @  @@@@@#+@  %@
'      @@.            #  @@    .  @   +#%                                                                                          +@@@@@@   @@     #@@@
'       @=           .@   @%  %= +                                                                                                             @@@@@@@@#
'       +-@@%=    : :@     @#*-                                    Scared Stiff (Bally 1996)
'         =@ +=-@@@@       %@@
'            #=@                                           https://www.ipdb.org/machine.cgi?id=3915
'
'
'
' VPW STIFFS
' ----------
' Sixtoe - Scratch table build and scripting.
' Tomate - 3D modeling and rendering.
' MCarter78 - 3D scanning, modeling and scripting.
' Apophis - Scripting and VR support.
' Cyberpez - Playfield Cutting, VR backglass.
' DGrimmReaper - Custom VR ball shooter.
' FrankEnstein - Custom instruction card.
'
' TABLE OF CONTENTS
' -----------------
'   ZVAR: Constants and Global Variables
'   ZTIM: Timers
'   ZINI: Table Initialization and Exiting
'   ZOPT: User Options
'   ZMAT: General Math Functions
'   ZBBR: Ball Brightness
'   ZRBR: Room Brightness
' ZKEY: Key Press Handling
' ZSOL: Solenoids & Flashers
' ZDRN: Drain, Trough, and Ball Release
' ZFLP: Flippers
' ZANI: Misc Animations
' ZSLG: Slingshot Animations
' ZBOO: Boogie Monsters
' ZSSC: Slingshot Corrections
' ZBMP: Bumpers
' ZSWI: Switches
' ZFRG: Frogs
' ZVUK: VUKs and Kickers
' ZSHA: Ambient ball shadows
' ZBRL: Ball Rolling and Drop Sounds
'   ZRRL: Ramp Rolling SFX
'   ZFLE: Fleep Mechanical Sounds
' ZNFF: Flipper Corrections and Tricks
'   ZDMP: Rubber Dampeners
' ZRST: Stand Up Targets
'   ZBOU: VPW TargetBouncer
' ZGIU: GI Control
'   ZVBG: VR Backglass
'   ZVRR: VR Room
'   ZPFW: Playfield Spider Wheel

Option Explicit
Randomize
SetLocale 1033      'Forces VBS to use english text encoding to stop crashes.

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

'******************************************************
'  ZVAR: Constants and Global Variables
'******************************************************
Const VRTest = false
Const BallSize = 50   'Ball size must be 50
Const BallMass = 1    'Ball mass must be 1
Const tnob = 4      'Total number of balls on the playfield including captive balls.
Const lob = 1     'Total number of locked balls

Dim tablewidth: tablewidth = Table1.width
Dim tableheight: tableheight = Table1.height
Dim gilvl

'  Standard definitions
Const cGameName = "SS_15"   'PinMAME ROM name
Const UseSolenoids = 2      '1 = Normal Flippers, 2 = Fastflips
Const UseLamps = 1        '0 = Custom lamp handling, 1 = Built-in VPX handling (using light number in light timer)
Const UseSync = 0
Const HandleMech = 0
Const SSolenoidOn = ""      'Sound sample used for this, obsolete.
Const SSolenoidOff = ""     ' ^
Const SFlipperOn = ""     ' ^
Const SFlipperOff = ""      ' ^
Const SCoin = ""        ' ^

Const UseVPMModSol = 2    'Set to 2 for PWM flashers, inserts, and GI. Requires VPinMame 3.6
Dim DesktopMode: DesktopMode = Table1.ShowDT
Dim UseVPMDMD: UseVPMDMD = DesktopMode Or (RenderingMode = 2) ' DMD for Desktop and VR

LoadVPM "03060000", "WPC.VBS", 3.36

NoUpperRightFlipper
NoUpperLeftFlipper

' VLM  Arrays - Start
' Arrays per baked part
Dim BP_Autoplunger: BP_Autoplunger=Array(BM_Autoplunger, LM_F_f36_Autoplunger)
Dim BP_BFlasher1: BP_BFlasher1=Array(BM_BFlasher1, LM_F_f17_BFlasher1, LM_F_f18_BFlasher1, LM_F_f22_BFlasher1, LM_GIC_BFlasher1, LM_GIU_BFlasher1)
Dim BP_BFlasher2: BP_BFlasher2=Array(BM_BFlasher2, LM_F_f18_BFlasher2, LM_GIU_BFlasher2)
Dim BP_BFlasher3: BP_BFlasher3=Array(BM_BFlasher3, LM_F_f17_BFlasher3, LM_F_f19_BFlasher3, LM_F_f28_BFlasher3)
Dim BP_BRing1: BP_BRing1=Array(BM_BRing1, LM_F_f17_BRing1, LM_F_f18_BRing1, LM_F_f19_BRing1, LM_F_f22_BRing1, LM_F_f23_BRing1, LM_F_f23a_BRing1, LM_F_f25_BRing1, LM_F_f28_BRing1, LM_GIC_BRing1, LM_GIU_BRing1)
Dim BP_BRing2: BP_BRing2=Array(BM_BRing2, LM_F_f17_BRing2, LM_F_f18_BRing2, LM_F_f19_BRing2, LM_F_f23a_BRing2, LM_F_f24_BRing2, LM_F_f28_BRing2, LM_GIC_BRing2, LM_GIU_BRing2)
Dim BP_BRing3: BP_BRing3=Array(BM_BRing3, LM_F_f18_BRing3, LM_F_f19_BRing3, LM_F_f23a_BRing3, LM_F_f28_BRing3, LM_GIC_BRing3)
Dim BP_BSocket1: BP_BSocket1=Array(BM_BSocket1, LM_F_f17_BSocket1, LM_F_f18_BSocket1, LM_F_f19_BSocket1, LM_F_f21_BSocket1, LM_F_f22_BSocket1, LM_F_f25_BSocket1, LM_GIC_BSocket1, LM_GIU_BSocket1)
Dim BP_BSocket2: BP_BSocket2=Array(BM_BSocket2, LM_F_f17_BSocket2, LM_F_f18_BSocket2, LM_F_f19_BSocket2, LM_F_f21_BSocket2, LM_F_f25_BSocket2, LM_GIC_BSocket2, LM_GIU_BSocket2)
Dim BP_BSocket3: BP_BSocket3=Array(BM_BSocket3, LM_F_f17_BSocket3, LM_F_f18_BSocket3, LM_F_f19_BSocket3, LM_F_f21_BSocket3, LM_F_f23a_BSocket3, LM_F_f28_BSocket3, LM_GIC_BSocket3, LM_GIU_BSocket3)
Dim BP_CoffinDiverter: BP_CoffinDiverter=Array(BM_CoffinDiverter, LM_F_f22_CoffinDiverter, LM_F_f23_CoffinDiverter, LM_F_f23a_CoffinDiverter, LM_F_f24_CoffinDiverter, LM_GIC_CoffinDiverter, LM_GIU_CoffinDiverter, LM_L_l56a_CoffinDiverter, LM_L_l58_CoffinDiverter)
Dim BP_Gate001: BP_Gate001=Array(BM_Gate001, LM_F_f28_Gate001, LM_GIL_gi_l_011_Gate001, LM_GIL_gi_l_012_Gate001, LM_GIC_Gate001)
Dim BP_Gate002: BP_Gate002=Array(BM_Gate002, LM_F_f27_Gate002, LM_GIU_Gate002)
Dim BP_Gate003: BP_Gate003=Array(BM_Gate003, LM_F_f24_Gate003, LM_GIC_Gate003)
Dim BP_Gate005: BP_Gate005=Array(BM_Gate005, LM_F_f28_Gate005)
Dim BP_Gate007: BP_Gate007=Array(BM_Gate007, LM_F_f22_Gate007, LM_F_f23_Gate007, LM_F_f24_Gate007, LM_GIU_Gate007, LM_L_l33_Gate007, LM_L_l34_Gate007)
Dim BP_GateLoop: BP_GateLoop=Array(BM_GateLoop, LM_F_f22_GateLoop, LM_GIU_GateLoop)
Dim BP_LBoogie01: BP_LBoogie01=Array(BM_LBoogie01, LM_F_f26_LBoogie01, LM_GIL_gi_l_003_LBoogie01, LM_GIL_gi_l_004_LBoogie01)
Dim BP_LBoogie02: BP_LBoogie02=Array(BM_LBoogie02, LM_F_f26_LBoogie02, LM_F_f35_LBoogie02, LM_L_l53_LBoogie02)
Dim BP_LSilng0: BP_LSilng0=Array(BM_LSilng0, LM_F_f26_LSilng0, LM_F_f35_LSilng0, LM_F_f36_LSilng0, LM_GIL_gi_l_001_LSilng0, LM_GIL_gi_l_002_LSilng0)
Dim BP_LSling1: BP_LSling1=Array(BM_LSling1, LM_F_f26_LSling1, LM_GIL_gi_l_001_LSling1, LM_GIL_gi_l_002_LSling1, LM_L_l53_LSling1)
Dim BP_LSling2: BP_LSling2=Array(BM_LSling2, LM_GIL_gi_l_001_LSling2, LM_GIL_gi_l_002_LSling2)
Dim BP_LSlingArmU: BP_LSlingArmU=Array(BM_LSlingArmU, LM_F_f23_LSlingArmU, LM_F_f23a_LSlingArmU, LM_F_f24_LSlingArmU, LM_F_f28_LSlingArmU, LM_GIC_LSlingArmU, LM_GIU_LSlingArmU)
Dim BP_LSlingU0: BP_LSlingU0=Array(BM_LSlingU0, LM_F_f23a_LSlingU0, LM_GIU_LSlingU0)
Dim BP_LSlingU1: BP_LSlingU1=Array(BM_LSlingU1, LM_F_f23a_LSlingU1, LM_GIU_LSlingU1)
Dim BP_LSlingU2: BP_LSlingU2=Array(BM_LSlingU2, LM_F_f23a_LSlingU2, LM_GIC_LSlingU2, LM_GIU_LSlingU2)
Dim BP_Layer1: BP_Layer1=Array(BM_Layer1, LM_F_f17_Layer1, LM_F_f18_Layer1, LM_F_f19_Layer1, LM_F_f21_Layer1, LM_F_f22_Layer1, LM_F_f23_Layer1, LM_F_f23a_Layer1, LM_F_f24_Layer1, LM_F_f25_Layer1, LM_F_f27_Layer1, LM_F_f28_Layer1, LM_F_f36_Layer1, LM_GIC_Layer1, LM_GIU_Layer1, LM_L_l23_Layer1, LM_L_l31_Layer1, LM_L_l32_Layer1, LM_L_l33_Layer1, LM_L_l34_Layer1, LM_L_l44_Layer1, LM_L_l84_Layer1, LM_L_l85_Layer1, LM_L_l86_Layer1)
Dim BP_Layer2: BP_Layer2=Array(BM_Layer2, LM_F_f17_Layer2, LM_F_f18_Layer2, LM_F_f19_Layer2, LM_F_f20_Layer2, LM_F_f21_Layer2, LM_F_f22_Layer2, LM_F_f23_Layer2, LM_F_f23a_Layer2, LM_F_f24_Layer2, LM_F_f25_Layer2, LM_F_f26_Layer2, LM_F_f27_Layer2, LM_F_f28_Layer2, LM_F_f36_Layer2, LM_GIL_gi_l_001_Layer2, LM_GIL_gi_l_002_Layer2, LM_GIL_gi_l_005_Layer2, LM_GIL_gi_l_006_Layer2, LM_GIL_gi_l_007_Layer2, LM_GIL_gi_l_008_Layer2, LM_GIL_gi_l_011_Layer2, LM_GIL_gi_l_012_Layer2, LM_GIC_Layer2, LM_GIU_Layer2, LM_L_l18_Layer2, LM_L_l23_Layer2, LM_L_l28_Layer2, LM_L_l31_Layer2, LM_L_l32_Layer2, LM_L_l33_Layer2, LM_L_l34_Layer2, LM_L_l36_Layer2, LM_L_l37_Layer2, LM_L_l44_Layer2, LM_L_l45_Layer2, LM_L_l47_Layer2, LM_L_l48_Layer2, LM_L_l51_Layer2, LM_L_l56a_Layer2, LM_L_l62_Layer2, LM_L_l84_Layer2, LM_L_l85_Layer2, LM_L_l86_Layer2)
Dim BP_Layer3: BP_Layer3=Array(BM_Layer3, LM_F_f18_Layer3, LM_F_f19_Layer3, LM_F_f21_Layer3, LM_F_f22_Layer3, LM_F_f23_Layer3, LM_F_f23a_Layer3, LM_F_f24_Layer3, LM_F_f25_Layer3, LM_F_f26_Layer3, LM_F_f27_Layer3, LM_F_f28_Layer3, LM_F_f35_Layer3, LM_F_f36_Layer3, LM_GIL_gi_l_001_Layer3, LM_GIL_gi_l_002_Layer3, LM_GIL_gi_l_003_Layer3, LM_GIL_gi_l_005_Layer3, LM_GIL_gi_l_007_Layer3, LM_GIL_gi_l_008_Layer3, LM_GIL_gi_l_009_Layer3, LM_GIL_gi_l_010_Layer3, LM_GIL_gi_l_011_Layer3, LM_GIL_gi_l_012_Layer3, LM_GIC_Layer3, LM_GIU_Layer3, LM_L_l18_Layer3, LM_L_l23_Layer3, LM_L_l25_Layer3, LM_L_l26_Layer3, LM_L_l27_Layer3, LM_L_l31_Layer3, LM_L_l32_Layer3, LM_L_l33_Layer3, LM_L_l34_Layer3, LM_L_l41_Layer3, LM_L_l42_Layer3, LM_L_l44_Layer3, LM_L_l46_Layer3, LM_L_l53_Layer3, LM_L_l54_Layer3, LM_L_l56_Layer3, LM_L_l56a_Layer3, LM_L_l57_Layer3, LM_L_l58_Layer3, LM_L_l62_Layer3, LM_L_l63_Layer3, LM_L_l84_Layer3)
Dim BP_Lflipper: BP_Lflipper=Array(BM_Lflipper, LM_F_f20_Lflipper, LM_F_f20a_Lflipper, LM_F_f26_Lflipper, LM_GIL_gi_l_001_Lflipper, LM_GIL_gi_l_002_Lflipper, LM_GIL_gi_l_003_Lflipper, LM_GIL_gi_l_004_Lflipper, LM_GIL_gi_l_005_Lflipper, LM_GIL_gi_l_006_Lflipper, LM_GIL_gi_l_007_Lflipper, LM_L_l55_Lflipper)
Dim BP_LflipperU: BP_LflipperU=Array(BM_LflipperU, LM_F_f20_LflipperU, LM_F_f20a_LflipperU, LM_F_f26_LflipperU, LM_GIL_gi_l_001_LflipperU, LM_GIL_gi_l_002_LflipperU, LM_GIL_gi_l_003_LflipperU, LM_GIL_gi_l_004_LflipperU, LM_GIL_gi_l_005_LflipperU, LM_GIL_gi_l_006_LflipperU, LM_GIL_gi_l_007_LflipperU, LM_L_l55_LflipperU)
Dim BP_Parts: BP_Parts=Array(BM_Parts, LM_F_f17_Parts, LM_F_f18_Parts, LM_F_f19_Parts, LM_F_f20_Parts, LM_F_f20a_Parts, LM_F_f21_Parts, LM_F_f22_Parts, LM_F_f23_Parts, LM_F_f23a_Parts, LM_F_f24_Parts, LM_F_f25_Parts, LM_F_f26_Parts, LM_F_f27_Parts, LM_F_f28_Parts, LM_F_f35_Parts, LM_F_f36_Parts, LM_GIL_gi_l_001_Parts, LM_GIL_gi_l_002_Parts, LM_GIL_gi_l_003_Parts, LM_GIL_gi_l_004_Parts, LM_GIL_gi_l_005_Parts, LM_GIL_gi_l_006_Parts, LM_GIL_gi_l_007_Parts, LM_GIL_gi_l_008_Parts, LM_GIL_gi_l_009_Parts, LM_GIL_gi_l_010_Parts, LM_GIL_gi_l_011_Parts, LM_GIL_gi_l_012_Parts, LM_GIC_Parts, LM_GIU_Parts, LM_L_l15_Parts, LM_L_l18_Parts, LM_L_l23_Parts, LM_L_l31_Parts, LM_L_l32_Parts, LM_L_l33_Parts, LM_L_l34_Parts, LM_L_l35_Parts, LM_L_l37_Parts, LM_L_l41_Parts, LM_L_l44_Parts, LM_L_l45_Parts, LM_L_l46_Parts, LM_L_l47_Parts, LM_L_l48_Parts, LM_L_l51_Parts, LM_L_l53_Parts, LM_L_l54_Parts, LM_L_l56_Parts, LM_L_l56a_Parts, LM_L_l57_Parts, LM_L_l58_Parts, LM_L_l84_Parts, LM_L_l85_Parts, LM_L_l86_Parts)
Dim BP_Playfield: BP_Playfield=Array(BM_Playfield, LM_F_f17_Playfield, LM_F_f18_Playfield, LM_F_f19_Playfield, LM_F_f20_Playfield, LM_F_f20a_Playfield, LM_F_f21_Playfield, LM_F_f22_Playfield, LM_F_f23_Playfield, LM_F_f23a_Playfield, LM_F_f24_Playfield, LM_F_f25_Playfield, LM_F_f26_Playfield, LM_F_f27_Playfield, LM_F_f28_Playfield, LM_F_f35_Playfield, LM_F_f36_Playfield, LM_GIL_gi_l_001_Playfield, LM_GIL_gi_l_002_Playfield, LM_GIL_gi_l_003_Playfield, LM_GIL_gi_l_004_Playfield, LM_GIL_gi_l_005_Playfield, LM_GIL_gi_l_006_Playfield, LM_GIL_gi_l_007_Playfield, LM_GIL_gi_l_008_Playfield, LM_GIL_gi_l_009_Playfield, LM_GIL_gi_l_010_Playfield, LM_GIL_gi_l_011_Playfield, LM_GIL_gi_l_012_Playfield, LM_GIC_Playfield, LM_GIU_Playfield, LM_L_l11_Playfield, LM_L_l12_Playfield, LM_L_l13_Playfield, LM_L_l14_Playfield, LM_L_l15_Playfield, LM_L_l16_Playfield, LM_L_l17_Playfield, LM_L_l21_Playfield, LM_L_l22_Playfield, LM_L_l23_Playfield, LM_L_l24_Playfield, LM_L_l25_Playfield, LM_L_l26_Playfield, LM_L_l27_Playfield, _
  LM_L_l28_Playfield, LM_L_l33_Playfield, LM_L_l34_Playfield, LM_L_l35_Playfield, LM_L_l36_Playfield, LM_L_l37_Playfield, LM_L_l38_Playfield, LM_L_l41_Playfield, LM_L_l42_Playfield, LM_L_l43_Playfield, LM_L_l45_Playfield, LM_L_l46_Playfield, LM_L_l47_Playfield, LM_L_l48_Playfield, LM_L_l51_Playfield, LM_L_l52_Playfield, LM_L_l53_Playfield, LM_L_l54_Playfield, LM_L_l55_Playfield, LM_L_l56_Playfield, LM_L_l57_Playfield, LM_L_l58_Playfield, LM_L_l61_Playfield, LM_L_l62_Playfield, LM_L_l63_Playfield, LM_L_l84_Playfield, LM_L_l85_Playfield, LM_L_l86_Playfield)
Dim BP_PrLeaper1: BP_PrLeaper1=Array(BM_PrLeaper1, LM_F_f23_PrLeaper1, LM_F_f23a_PrLeaper1, LM_F_f24_PrLeaper1, LM_GIC_PrLeaper1, LM_GIU_PrLeaper1, LM_L_l41_PrLeaper1, LM_L_l57_PrLeaper1)
Dim BP_PrLeaper1rod: BP_PrLeaper1rod=Array(BM_PrLeaper1rod, LM_F_f22_PrLeaper1rod, LM_F_f23_PrLeaper1rod, LM_F_f23a_PrLeaper1rod, LM_F_f24_PrLeaper1rod, LM_GIC_PrLeaper1rod)
Dim BP_PrLeaper2: BP_PrLeaper2=Array(BM_PrLeaper2, LM_F_f18_PrLeaper2, LM_F_f23_PrLeaper2, LM_F_f23a_PrLeaper2, LM_F_f24_PrLeaper2, LM_F_f28_PrLeaper2, LM_GIC_PrLeaper2, LM_GIU_PrLeaper2, LM_L_l21_PrLeaper2, LM_L_l22_PrLeaper2, LM_L_l24_PrLeaper2, LM_L_l42_PrLeaper2)
Dim BP_PrLeaper2rod: BP_PrLeaper2rod=Array(BM_PrLeaper2rod, LM_F_f23_PrLeaper2rod, LM_F_f23a_PrLeaper2rod, LM_F_f24_PrLeaper2rod, LM_GIL_gi_l_007_PrLeaper2rod, LM_GIL_gi_l_012_PrLeaper2rod, LM_GIC_PrLeaper2rod)
Dim BP_PrLeaper3: BP_PrLeaper3=Array(BM_PrLeaper3, LM_F_f18_PrLeaper3, LM_F_f23a_PrLeaper3, LM_F_f24_PrLeaper3, LM_F_f28_PrLeaper3, LM_GIC_PrLeaper3, LM_GIU_PrLeaper3, LM_L_l36_PrLeaper3)
Dim BP_PrLeaper3rod: BP_PrLeaper3rod=Array(BM_PrLeaper3rod, LM_F_f24_PrLeaper3rod, LM_F_f35_PrLeaper3rod, LM_F_f36_PrLeaper3rod, LM_GIL_gi_l_001_PrLeaper3rod, LM_GIC_PrLeaper3rod, LM_L_l54_PrLeaper3rod)
Dim BP_RBoggie01: BP_RBoggie01=Array(BM_RBoggie01, LM_F_f20a_RBoggie01, LM_F_f26_RBoggie01, LM_GIL_gi_l_005_RBoggie01, LM_GIL_gi_l_006_RBoggie01, LM_GIL_gi_l_007_RBoggie01)
Dim BP_RBoggie02: BP_RBoggie02=Array(BM_RBoggie02, LM_F_f26_RBoggie02, LM_F_f36_RBoggie02)
Dim BP_RSling0: BP_RSling0=Array(BM_RSling0, LM_F_f35_RSling0, LM_GIL_gi_l_007_RSling0, LM_GIL_gi_l_008_RSling0)
Dim BP_RSling1: BP_RSling1=Array(BM_RSling1, LM_GIL_gi_l_008_RSling1, LM_L_l51_RSling1)
Dim BP_RSling2: BP_RSling2=Array(BM_RSling2, LM_GIL_gi_l_007_RSling2, LM_GIL_gi_l_008_RSling2)
Dim BP_Rflipper: BP_Rflipper=Array(BM_Rflipper, LM_F_f20_Rflipper, LM_F_f20a_Rflipper, LM_F_f26_Rflipper, LM_GIL_gi_l_003_Rflipper, LM_GIL_gi_l_004_Rflipper, LM_GIL_gi_l_005_Rflipper, LM_GIL_gi_l_006_Rflipper, LM_GIL_gi_l_007_Rflipper, LM_GIL_gi_l_008_Rflipper, LM_L_l55_Rflipper)
Dim BP_RflipperU: BP_RflipperU=Array(BM_RflipperU, LM_F_f20_RflipperU, LM_F_f20a_RflipperU, LM_F_f26_RflipperU, LM_GIL_gi_l_003_RflipperU, LM_GIL_gi_l_004_RflipperU, LM_GIL_gi_l_005_RflipperU, LM_GIL_gi_l_006_RflipperU, LM_GIL_gi_l_007_RflipperU, LM_GIL_gi_l_008_RflipperU, LM_L_l55_RflipperU)
Dim BP_cDoorClose: BP_cDoorClose=Array(BM_cDoorClose, LM_F_f23a_cDoorClose, LM_F_f24_cDoorClose, LM_F_f26_cDoorClose, LM_F_f35_cDoorClose, LM_GIL_gi_l_009_cDoorClose, LM_GIL_gi_l_010_cDoorClose, LM_GIC_cDoorClose, LM_L_l54_cDoorClose)
Dim BP_cDoorOpen: BP_cDoorOpen=Array(BM_cDoorOpen, LM_F_f24_cDoorOpen, LM_F_f35_cDoorOpen, LM_GIL_gi_l_009_cDoorOpen, LM_GIL_gi_l_010_cDoorOpen, LM_GIC_cDoorOpen, LM_L_l54_cDoorOpen)
Dim BP_cadaver: BP_cadaver=Array(BM_cadaver, LM_F_f35_cadaver, LM_F_f36_cadaver, LM_GIL_gi_l_009_cadaver, LM_GIC_cadaver, LM_L_l54_cadaver)
Dim BP_crate_door: BP_crate_door=Array(BM_crate_door, LM_F_f18_crate_door, LM_F_f23a_crate_door, LM_F_f24_crate_door, LM_F_f28_crate_door, LM_GIC_crate_door, LM_GIU_crate_door, LM_L_l23_crate_door, LM_L_l46_crate_door, LM_L_l47_crate_door)
Dim BP_crate_doorU: BP_crate_doorU=Array(BM_crate_doorU, LM_F_f18_crate_doorU, LM_F_f23a_crate_doorU, LM_F_f28_crate_doorU, LM_GIC_crate_doorU, LM_GIU_crate_doorU, LM_L_l46_crate_doorU)
Dim BP_lPost: BP_lPost=Array(BM_lPost, LM_GIL_gi_l_001_lPost, LM_GIL_gi_l_009_lPost, LM_GIL_gi_l_010_lPost)
Dim BP_lemk: BP_lemk=Array(BM_lemk, LM_F_f20_lemk, LM_F_f26_lemk, LM_F_f35_lemk, LM_GIL_gi_l_001_lemk, LM_GIL_gi_l_002_lemk, LM_GIL_gi_l_003_lemk, LM_GIL_gi_l_004_lemk, LM_GIL_gi_l_005_lemk, LM_GIL_gi_l_006_lemk, LM_GIL_gi_l_007_lemk, LM_GIL_gi_l_008_lemk, LM_L_l53_lemk)
Dim BP_lockdownbar: BP_lockdownbar=Array(BM_lockdownbar, LM_F_f27_lockdownbar, LM_GIU_lockdownbar)
Dim BP_rPost: BP_rPost=Array(BM_rPost, LM_GIL_gi_l_008_rPost, LM_GIL_gi_l_011_rPost, LM_GIL_gi_l_012_rPost)
Dim BP_rampFlap: BP_rampFlap=Array(BM_rampFlap, LM_F_f28_rampFlap, LM_GIC_rampFlap)
Dim BP_remk: BP_remk=Array(BM_remk, LM_F_f20_remk, LM_F_f20a_remk, LM_F_f26_remk, LM_F_f36_remk, LM_GIL_gi_l_001_remk, LM_GIL_gi_l_002_remk, LM_GIL_gi_l_003_remk, LM_GIL_gi_l_004_remk, LM_GIL_gi_l_005_remk, LM_GIL_gi_l_006_remk, LM_GIL_gi_l_007_remk, LM_GIL_gi_l_008_remk, LM_GIL_gi_l_012_remk, LM_L_l51_remk)
Dim BP_sideblades: BP_sideblades=Array(BM_sideblades, LM_F_f17_sideblades, LM_F_f18_sideblades, LM_F_f19_sideblades, LM_F_f21_sideblades, LM_F_f22_sideblades, LM_F_f23_sideblades, LM_F_f23a_sideblades, LM_F_f24_sideblades, LM_F_f25_sideblades, LM_F_f26_sideblades, LM_F_f27_sideblades, LM_F_f28_sideblades, LM_F_f35_sideblades, LM_F_f36_sideblades, LM_GIL_gi_l_001_sideblades, LM_GIL_gi_l_002_sideblades, LM_GIL_gi_l_003_sideblades, LM_GIL_gi_l_004_sideblades, LM_GIL_gi_l_006_sideblades, LM_GIL_gi_l_007_sideblades, LM_GIL_gi_l_008_sideblades, LM_GIL_gi_l_009_sideblades, LM_GIL_gi_l_010_sideblades, LM_GIL_gi_l_011_sideblades, LM_GIL_gi_l_012_sideblades, LM_GIC_sideblades, LM_GIU_sideblades, LM_L_l35_sideblades, LM_L_l54_sideblades)
Dim BP_sw16: BP_sw16=Array(BM_sw16, LM_F_f35_sw16, LM_GIL_gi_l_001_sw16, LM_GIL_gi_l_002_sw16, LM_GIL_gi_l_003_sw16, LM_GIL_gi_l_004_sw16, LM_GIL_gi_l_008_sw16)
Dim BP_sw17: BP_sw17=Array(BM_sw17, LM_GIL_gi_l_005_sw17, LM_GIL_gi_l_006_sw17, LM_GIL_gi_l_007_sw17, LM_GIL_gi_l_008_sw17, LM_GIL_gi_l_011_sw17, LM_GIL_gi_l_012_sw17)
Dim BP_sw18: BP_sw18=Array(BM_sw18)
Dim BP_sw25: BP_sw25=Array(BM_sw25, LM_F_f22_sw25, LM_GIU_sw25)
Dim BP_sw26: BP_sw26=Array(BM_sw26, LM_GIL_gi_l_001_sw26, LM_GIL_gi_l_002_sw26, LM_GIL_gi_l_003_sw26)
Dim BP_sw27: BP_sw27=Array(BM_sw27, LM_F_f36_sw27, LM_GIL_gi_l_008_sw27, LM_GIL_gi_l_011_sw27, LM_GIL_gi_l_012_sw27)
Dim BP_sw28: BP_sw28=Array(BM_sw28, LM_F_f26_sw28, LM_F_f36_sw28, LM_GIL_gi_l_001_sw28, LM_GIL_gi_l_005_sw28, LM_GIL_gi_l_006_sw28, LM_GIL_gi_l_007_sw28, LM_GIL_gi_l_008_sw28, LM_GIL_gi_l_011_sw28, LM_GIL_gi_l_012_sw28)
Dim BP_sw58: BP_sw58=Array(BM_sw58, LM_F_f23_sw58, LM_F_f23a_sw58, LM_F_f24_sw58, LM_F_f27_sw58, LM_GIU_sw58)
Dim BP_sw61: BP_sw61=Array(BM_sw61, LM_F_f26_sw61, LM_F_f35_sw61, LM_GIL_gi_l_001_sw61, LM_GIL_gi_l_003_sw61, LM_GIL_gi_l_008_sw61, LM_GIL_gi_l_009_sw61, LM_GIL_gi_l_010_sw61, LM_GIC_sw61, LM_L_l25_sw61, LM_L_l26_sw61, LM_L_l27_sw61)
Dim BP_sw62: BP_sw62=Array(BM_sw62, LM_F_f26_sw62, LM_F_f35_sw62, LM_GIL_gi_l_001_sw62, LM_GIL_gi_l_002_sw62, LM_GIL_gi_l_003_sw62, LM_GIL_gi_l_008_sw62, LM_GIL_gi_l_009_sw62, LM_GIL_gi_l_010_sw62, LM_L_l25_sw62, LM_L_l26_sw62, LM_L_l27_sw62)
Dim BP_sw63: BP_sw63=Array(BM_sw63, LM_F_f26_sw63, LM_F_f35_sw63, LM_GIL_gi_l_001_sw63, LM_GIL_gi_l_002_sw63, LM_GIL_gi_l_003_sw63, LM_GIL_gi_l_004_sw63, LM_GIL_gi_l_008_sw63, LM_GIL_gi_l_009_sw63, LM_GIL_gi_l_010_sw63, LM_L_l25_sw63, LM_L_l26_sw63, LM_L_l27_sw63)
Dim BP_sw64: BP_sw64=Array(BM_sw64, LM_F_f23_sw64, LM_F_f23a_sw64, LM_F_f24_sw64, LM_GIC_sw64, LM_L_l41_sw64)
Dim BP_sw65: BP_sw65=Array(BM_sw65, LM_F_f23a_sw65, LM_F_f24_sw65, LM_GIC_sw65, LM_L_l24_sw65)
Dim BP_sw66: BP_sw66=Array(BM_sw66, LM_F_f26_sw66, LM_GIC_sw66, LM_L_l36_sw66)
Dim BP_sw68: BP_sw68=Array(BM_sw68, LM_F_f21_sw68, LM_F_f22_sw68, LM_GIU_sw68)
Dim BP_sw71: BP_sw71=Array(BM_sw71, LM_GIU_sw71)
Dim BP_sw72: BP_sw72=Array(BM_sw72, LM_GIU_sw72, LM_L_l31_sw72, LM_L_l32_sw72)
Dim BP_sw73: BP_sw73=Array(BM_sw73, LM_GIU_sw73)
Dim BP_sw74: BP_sw74=Array(BM_sw74, LM_F_f23_sw74, LM_F_f23a_sw74, LM_F_f24_sw74, LM_F_f27_sw74, LM_GIU_sw74)
Dim BP_underPF: BP_underPF=Array(BM_underPF, LM_F_f20_underPF, LM_F_f20a_underPF, LM_F_f21_underPF, LM_F_f22_underPF, LM_F_f23_underPF, LM_F_f23a_underPF, LM_F_f24_underPF, LM_F_f25_underPF, LM_F_f26_underPF, LM_F_f35_underPF, LM_F_f36_underPF, LM_GIL_gi_l_001_underPF, LM_GIL_gi_l_011_underPF, LM_GIL_gi_l_012_underPF, LM_GIC_underPF, LM_GIU_underPF, LM_L_l11_underPF, LM_L_l12_underPF, LM_L_l13_underPF, LM_L_l14_underPF, LM_L_l15_underPF, LM_L_l16_underPF, LM_L_l17_underPF, LM_L_l21_underPF, LM_L_l22_underPF, LM_L_l23_underPF, LM_L_l24_underPF, LM_L_l25_underPF, LM_L_l26_underPF, LM_L_l27_underPF, LM_L_l28_underPF, LM_L_l35_underPF, LM_L_l36_underPF, LM_L_l37_underPF, LM_L_l38_underPF, LM_L_l41_underPF, LM_L_l42_underPF, LM_L_l43_underPF, LM_L_l45_underPF, LM_L_l46_underPF, LM_L_l47_underPF, LM_L_l48_underPF, LM_L_l51_underPF, LM_L_l52_underPF, LM_L_l53_underPF, LM_L_l54_underPF, LM_L_l55_underPF, LM_L_l56_underPF, LM_L_l57_underPF, LM_L_l58_underPF, LM_L_l61_underPF, LM_L_l62_underPF, LM_L_l63_underPF, _
  LM_L_l84_underPF, LM_L_l85_underPF, LM_L_l86_underPF)
' Arrays per lighting scenario
'Dim BL_F_f17: BL_F_f17=Array(LM_F_f17_BFlasher1, LM_F_f17_BFlasher3, LM_F_f17_BRing1, LM_F_f17_BRing2, LM_F_f17_BSocket1, LM_F_f17_BSocket2, LM_F_f17_BSocket3, LM_F_f17_Layer1, LM_F_f17_Layer2, LM_F_f17_Parts, LM_F_f17_Playfield, LM_F_f17_sideblades)
'Dim BL_F_f18: BL_F_f18=Array(LM_F_f18_BFlasher1, LM_F_f18_BFlasher2, LM_F_f18_BRing1, LM_F_f18_BRing2, LM_F_f18_BRing3, LM_F_f18_BSocket1, LM_F_f18_BSocket2, LM_F_f18_BSocket3, LM_F_f18_Layer1, LM_F_f18_Layer2, LM_F_f18_Layer3, LM_F_f18_Parts, LM_F_f18_Playfield, LM_F_f18_PrLeaper2, LM_F_f18_PrLeaper3, LM_F_f18_crate_door, LM_F_f18_crate_doorU, LM_F_f18_sideblades)
'Dim BL_F_f19: BL_F_f19=Array(LM_F_f19_BFlasher3, LM_F_f19_BRing1, LM_F_f19_BRing2, LM_F_f19_BRing3, LM_F_f19_BSocket1, LM_F_f19_BSocket2, LM_F_f19_BSocket3, LM_F_f19_Layer1, LM_F_f19_Layer2, LM_F_f19_Layer3, LM_F_f19_Parts, LM_F_f19_Playfield, LM_F_f19_sideblades)
'Dim BL_F_f20: BL_F_f20=Array(LM_F_f20_Layer2, LM_F_f20_Lflipper, LM_F_f20_LflipperU, LM_F_f20_Parts, LM_F_f20_Playfield, LM_F_f20_Rflipper, LM_F_f20_RflipperU, LM_F_f20_lemk, LM_F_f20_remk, LM_F_f20_underPF)
'Dim BL_F_f20a: BL_F_f20a=Array(LM_F_f20a_Lflipper, LM_F_f20a_LflipperU, LM_F_f20a_Parts, LM_F_f20a_Playfield, LM_F_f20a_RBoggie01, LM_F_f20a_Rflipper, LM_F_f20a_RflipperU, LM_F_f20a_remk, LM_F_f20a_underPF)
'Dim BL_F_f21: BL_F_f21=Array(LM_F_f21_BSocket1, LM_F_f21_BSocket2, LM_F_f21_BSocket3, LM_F_f21_Layer1, LM_F_f21_Layer2, LM_F_f21_Layer3, LM_F_f21_Parts, LM_F_f21_Playfield, LM_F_f21_sideblades, LM_F_f21_sw68, LM_F_f21_underPF)
'Dim BL_F_f22: BL_F_f22=Array(LM_F_f22_BFlasher1, LM_F_f22_BRing1, LM_F_f22_BSocket1, LM_F_f22_CoffinDiverter, LM_F_f22_Gate007, LM_F_f22_GateLoop, LM_F_f22_Layer1, LM_F_f22_Layer2, LM_F_f22_Layer3, LM_F_f22_Parts, LM_F_f22_Playfield, LM_F_f22_PrLeaper1rod, LM_F_f22_sideblades, LM_F_f22_sw25, LM_F_f22_sw68, LM_F_f22_underPF)
'Dim BL_F_f23: BL_F_f23=Array(LM_F_f23_BRing1, LM_F_f23_CoffinDiverter, LM_F_f23_Gate007, LM_F_f23_LSlingArmU, LM_F_f23_Layer1, LM_F_f23_Layer2, LM_F_f23_Layer3, LM_F_f23_Parts, LM_F_f23_Playfield, LM_F_f23_PrLeaper1, LM_F_f23_PrLeaper1rod, LM_F_f23_PrLeaper2, LM_F_f23_PrLeaper2rod, LM_F_f23_sideblades, LM_F_f23_sw58, LM_F_f23_sw64, LM_F_f23_sw74, LM_F_f23_underPF)
'Dim BL_F_f23a: BL_F_f23a=Array(LM_F_f23a_BRing1, LM_F_f23a_BRing2, LM_F_f23a_BRing3, LM_F_f23a_BSocket3, LM_F_f23a_CoffinDiverter, LM_F_f23a_LSlingArmU, LM_F_f23a_LSlingU0, LM_F_f23a_LSlingU1, LM_F_f23a_LSlingU2, LM_F_f23a_Layer1, LM_F_f23a_Layer2, LM_F_f23a_Layer3, LM_F_f23a_Parts, LM_F_f23a_Playfield, LM_F_f23a_PrLeaper1, LM_F_f23a_PrLeaper1rod, LM_F_f23a_PrLeaper2, LM_F_f23a_PrLeaper2rod, LM_F_f23a_PrLeaper3, LM_F_f23a_cDoorClose, LM_F_f23a_crate_door, LM_F_f23a_crate_doorU, LM_F_f23a_sideblades, LM_F_f23a_sw58, LM_F_f23a_sw64, LM_F_f23a_sw65, LM_F_f23a_sw74, LM_F_f23a_underPF)
'Dim BL_F_f24: BL_F_f24=Array(LM_F_f24_BRing2, LM_F_f24_CoffinDiverter, LM_F_f24_Gate003, LM_F_f24_Gate007, LM_F_f24_LSlingArmU, LM_F_f24_Layer1, LM_F_f24_Layer2, LM_F_f24_Layer3, LM_F_f24_Parts, LM_F_f24_Playfield, LM_F_f24_PrLeaper1, LM_F_f24_PrLeaper1rod, LM_F_f24_PrLeaper2, LM_F_f24_PrLeaper2rod, LM_F_f24_PrLeaper3, LM_F_f24_PrLeaper3rod, LM_F_f24_cDoorClose, LM_F_f24_cDoorOpen, LM_F_f24_crate_door, LM_F_f24_sideblades, LM_F_f24_sw58, LM_F_f24_sw64, LM_F_f24_sw65, LM_F_f24_sw74, LM_F_f24_underPF)
'Dim BL_F_f25: BL_F_f25=Array(LM_F_f25_BRing1, LM_F_f25_BSocket1, LM_F_f25_BSocket2, LM_F_f25_Layer1, LM_F_f25_Layer2, LM_F_f25_Layer3, LM_F_f25_Parts, LM_F_f25_Playfield, LM_F_f25_sideblades, LM_F_f25_underPF)
'Dim BL_F_f26: BL_F_f26=Array(LM_F_f26_LBoogie01, LM_F_f26_LBoogie02, LM_F_f26_LSilng0, LM_F_f26_LSling1, LM_F_f26_Layer2, LM_F_f26_Layer3, LM_F_f26_Lflipper, LM_F_f26_LflipperU, LM_F_f26_Parts, LM_F_f26_Playfield, LM_F_f26_RBoggie01, LM_F_f26_RBoggie02, LM_F_f26_Rflipper, LM_F_f26_RflipperU, LM_F_f26_cDoorClose, LM_F_f26_lemk, LM_F_f26_remk, LM_F_f26_sideblades, LM_F_f26_sw28, LM_F_f26_sw61, LM_F_f26_sw62, LM_F_f26_sw63, LM_F_f26_sw66, LM_F_f26_underPF)
'Dim BL_F_f27: BL_F_f27=Array(LM_F_f27_Gate002, LM_F_f27_Layer1, LM_F_f27_Layer2, LM_F_f27_Layer3, LM_F_f27_Parts, LM_F_f27_Playfield, LM_F_f27_lockdownbar, LM_F_f27_sideblades, LM_F_f27_sw58, LM_F_f27_sw74)
'Dim BL_F_f28: BL_F_f28=Array(LM_F_f28_BFlasher3, LM_F_f28_BRing1, LM_F_f28_BRing2, LM_F_f28_BRing3, LM_F_f28_BSocket3, LM_F_f28_Gate001, LM_F_f28_Gate005, LM_F_f28_LSlingArmU, LM_F_f28_Layer1, LM_F_f28_Layer2, LM_F_f28_Layer3, LM_F_f28_Parts, LM_F_f28_Playfield, LM_F_f28_PrLeaper2, LM_F_f28_PrLeaper3, LM_F_f28_crate_door, LM_F_f28_crate_doorU, LM_F_f28_rampFlap, LM_F_f28_sideblades)
'Dim BL_F_f35: BL_F_f35=Array(LM_F_f35_LBoogie02, LM_F_f35_LSilng0, LM_F_f35_Layer3, LM_F_f35_Parts, LM_F_f35_Playfield, LM_F_f35_PrLeaper3rod, LM_F_f35_RSling0, LM_F_f35_cDoorClose, LM_F_f35_cDoorOpen, LM_F_f35_cadaver, LM_F_f35_lemk, LM_F_f35_sideblades, LM_F_f35_sw16, LM_F_f35_sw61, LM_F_f35_sw62, LM_F_f35_sw63, LM_F_f35_underPF)
'Dim BL_F_f36: BL_F_f36=Array(LM_F_f36_Autoplunger, LM_F_f36_LSilng0, LM_F_f36_Layer1, LM_F_f36_Layer2, LM_F_f36_Layer3, LM_F_f36_Parts, LM_F_f36_Playfield, LM_F_f36_PrLeaper3rod, LM_F_f36_RBoggie02, LM_F_f36_cadaver, LM_F_f36_remk, LM_F_f36_sideblades, LM_F_f36_sw27, LM_F_f36_sw28, LM_F_f36_underPF)
'Dim BL_GIC: BL_GIC=Array(LM_GIC_BFlasher1, LM_GIC_BRing1, LM_GIC_BRing2, LM_GIC_BRing3, LM_GIC_BSocket1, LM_GIC_BSocket2, LM_GIC_BSocket3, LM_GIC_CoffinDiverter, LM_GIC_Gate001, LM_GIC_Gate003, LM_GIC_LSlingArmU, LM_GIC_LSlingU2, LM_GIC_Layer1, LM_GIC_Layer2, LM_GIC_Layer3, LM_GIC_Parts, LM_GIC_Playfield, LM_GIC_PrLeaper1, LM_GIC_PrLeaper1rod, LM_GIC_PrLeaper2, LM_GIC_PrLeaper2rod, LM_GIC_PrLeaper3, LM_GIC_PrLeaper3rod, LM_GIC_cDoorClose, LM_GIC_cDoorOpen, LM_GIC_cadaver, LM_GIC_crate_door, LM_GIC_crate_doorU, LM_GIC_rampFlap, LM_GIC_sideblades, LM_GIC_sw61, LM_GIC_sw64, LM_GIC_sw65, LM_GIC_sw66, LM_GIC_underPF)
'Dim BL_GIL_gi_l_001: BL_GIL_gi_l_001=Array(LM_GIL_gi_l_001_LSilng0, LM_GIL_gi_l_001_LSling1, LM_GIL_gi_l_001_LSling2, LM_GIL_gi_l_001_Layer2, LM_GIL_gi_l_001_Layer3, LM_GIL_gi_l_001_Lflipper, LM_GIL_gi_l_001_LflipperU, LM_GIL_gi_l_001_Parts, LM_GIL_gi_l_001_Playfield, LM_GIL_gi_l_001_PrLeaper3rod, LM_GIL_gi_l_001_lPost, LM_GIL_gi_l_001_lemk, LM_GIL_gi_l_001_remk, LM_GIL_gi_l_001_sideblades, LM_GIL_gi_l_001_sw16, LM_GIL_gi_l_001_sw26, LM_GIL_gi_l_001_sw28, LM_GIL_gi_l_001_sw61, LM_GIL_gi_l_001_sw62, LM_GIL_gi_l_001_sw63, LM_GIL_gi_l_001_underPF)
'Dim BL_GIL_gi_l_002: BL_GIL_gi_l_002=Array(LM_GIL_gi_l_002_LSilng0, LM_GIL_gi_l_002_LSling1, LM_GIL_gi_l_002_LSling2, LM_GIL_gi_l_002_Layer2, LM_GIL_gi_l_002_Layer3, LM_GIL_gi_l_002_Lflipper, LM_GIL_gi_l_002_LflipperU, LM_GIL_gi_l_002_Parts, LM_GIL_gi_l_002_Playfield, LM_GIL_gi_l_002_lemk, LM_GIL_gi_l_002_remk, LM_GIL_gi_l_002_sideblades, LM_GIL_gi_l_002_sw16, LM_GIL_gi_l_002_sw26, LM_GIL_gi_l_002_sw62, LM_GIL_gi_l_002_sw63)
'Dim BL_GIL_gi_l_003: BL_GIL_gi_l_003=Array(LM_GIL_gi_l_003_LBoogie01, LM_GIL_gi_l_003_Layer3, LM_GIL_gi_l_003_Lflipper, LM_GIL_gi_l_003_LflipperU, LM_GIL_gi_l_003_Parts, LM_GIL_gi_l_003_Playfield, LM_GIL_gi_l_003_Rflipper, LM_GIL_gi_l_003_RflipperU, LM_GIL_gi_l_003_lemk, LM_GIL_gi_l_003_remk, LM_GIL_gi_l_003_sideblades, LM_GIL_gi_l_003_sw16, LM_GIL_gi_l_003_sw26, LM_GIL_gi_l_003_sw61, LM_GIL_gi_l_003_sw62, LM_GIL_gi_l_003_sw63)
'Dim BL_GIL_gi_l_004: BL_GIL_gi_l_004=Array(LM_GIL_gi_l_004_LBoogie01, LM_GIL_gi_l_004_Lflipper, LM_GIL_gi_l_004_LflipperU, LM_GIL_gi_l_004_Parts, LM_GIL_gi_l_004_Playfield, LM_GIL_gi_l_004_Rflipper, LM_GIL_gi_l_004_RflipperU, LM_GIL_gi_l_004_lemk, LM_GIL_gi_l_004_remk, LM_GIL_gi_l_004_sideblades, LM_GIL_gi_l_004_sw16, LM_GIL_gi_l_004_sw63)
'Dim BL_GIL_gi_l_005: BL_GIL_gi_l_005=Array(LM_GIL_gi_l_005_Layer2, LM_GIL_gi_l_005_Layer3, LM_GIL_gi_l_005_Lflipper, LM_GIL_gi_l_005_LflipperU, LM_GIL_gi_l_005_Parts, LM_GIL_gi_l_005_Playfield, LM_GIL_gi_l_005_RBoggie01, LM_GIL_gi_l_005_Rflipper, LM_GIL_gi_l_005_RflipperU, LM_GIL_gi_l_005_lemk, LM_GIL_gi_l_005_remk, LM_GIL_gi_l_005_sw17, LM_GIL_gi_l_005_sw28)
'Dim BL_GIL_gi_l_006: BL_GIL_gi_l_006=Array(LM_GIL_gi_l_006_Layer2, LM_GIL_gi_l_006_Lflipper, LM_GIL_gi_l_006_LflipperU, LM_GIL_gi_l_006_Parts, LM_GIL_gi_l_006_Playfield, LM_GIL_gi_l_006_RBoggie01, LM_GIL_gi_l_006_Rflipper, LM_GIL_gi_l_006_RflipperU, LM_GIL_gi_l_006_lemk, LM_GIL_gi_l_006_remk, LM_GIL_gi_l_006_sideblades, LM_GIL_gi_l_006_sw17, LM_GIL_gi_l_006_sw28)
'Dim BL_GIL_gi_l_007: BL_GIL_gi_l_007=Array(LM_GIL_gi_l_007_Layer2, LM_GIL_gi_l_007_Layer3, LM_GIL_gi_l_007_Lflipper, LM_GIL_gi_l_007_LflipperU, LM_GIL_gi_l_007_Parts, LM_GIL_gi_l_007_Playfield, LM_GIL_gi_l_007_PrLeaper2rod, LM_GIL_gi_l_007_RBoggie01, LM_GIL_gi_l_007_RSling0, LM_GIL_gi_l_007_RSling2, LM_GIL_gi_l_007_Rflipper, LM_GIL_gi_l_007_RflipperU, LM_GIL_gi_l_007_lemk, LM_GIL_gi_l_007_remk, LM_GIL_gi_l_007_sideblades, LM_GIL_gi_l_007_sw17, LM_GIL_gi_l_007_sw28)
'Dim BL_GIL_gi_l_008: BL_GIL_gi_l_008=Array(LM_GIL_gi_l_008_Layer2, LM_GIL_gi_l_008_Layer3, LM_GIL_gi_l_008_Parts, LM_GIL_gi_l_008_Playfield, LM_GIL_gi_l_008_RSling0, LM_GIL_gi_l_008_RSling1, LM_GIL_gi_l_008_RSling2, LM_GIL_gi_l_008_Rflipper, LM_GIL_gi_l_008_RflipperU, LM_GIL_gi_l_008_lemk, LM_GIL_gi_l_008_rPost, LM_GIL_gi_l_008_remk, LM_GIL_gi_l_008_sideblades, LM_GIL_gi_l_008_sw16, LM_GIL_gi_l_008_sw17, LM_GIL_gi_l_008_sw27, LM_GIL_gi_l_008_sw28, LM_GIL_gi_l_008_sw61, LM_GIL_gi_l_008_sw62, LM_GIL_gi_l_008_sw63)
'Dim BL_GIL_gi_l_009: BL_GIL_gi_l_009=Array(LM_GIL_gi_l_009_Layer3, LM_GIL_gi_l_009_Parts, LM_GIL_gi_l_009_Playfield, LM_GIL_gi_l_009_cDoorClose, LM_GIL_gi_l_009_cDoorOpen, LM_GIL_gi_l_009_cadaver, LM_GIL_gi_l_009_lPost, LM_GIL_gi_l_009_sideblades, LM_GIL_gi_l_009_sw61, LM_GIL_gi_l_009_sw62, LM_GIL_gi_l_009_sw63)
'Dim BL_GIL_gi_l_010: BL_GIL_gi_l_010=Array(LM_GIL_gi_l_010_Layer3, LM_GIL_gi_l_010_Parts, LM_GIL_gi_l_010_Playfield, LM_GIL_gi_l_010_cDoorClose, LM_GIL_gi_l_010_cDoorOpen, LM_GIL_gi_l_010_lPost, LM_GIL_gi_l_010_sideblades, LM_GIL_gi_l_010_sw61, LM_GIL_gi_l_010_sw62, LM_GIL_gi_l_010_sw63)
'Dim BL_GIL_gi_l_011: BL_GIL_gi_l_011=Array(LM_GIL_gi_l_011_Gate001, LM_GIL_gi_l_011_Layer2, LM_GIL_gi_l_011_Layer3, LM_GIL_gi_l_011_Parts, LM_GIL_gi_l_011_Playfield, LM_GIL_gi_l_011_rPost, LM_GIL_gi_l_011_sideblades, LM_GIL_gi_l_011_sw17, LM_GIL_gi_l_011_sw27, LM_GIL_gi_l_011_sw28, LM_GIL_gi_l_011_underPF)
'Dim BL_GIL_gi_l_012: BL_GIL_gi_l_012=Array(LM_GIL_gi_l_012_Gate001, LM_GIL_gi_l_012_Layer2, LM_GIL_gi_l_012_Layer3, LM_GIL_gi_l_012_Parts, LM_GIL_gi_l_012_Playfield, LM_GIL_gi_l_012_PrLeaper2rod, LM_GIL_gi_l_012_rPost, LM_GIL_gi_l_012_remk, LM_GIL_gi_l_012_sideblades, LM_GIL_gi_l_012_sw17, LM_GIL_gi_l_012_sw27, LM_GIL_gi_l_012_sw28, LM_GIL_gi_l_012_underPF)
'Dim BL_GIU: BL_GIU=Array(LM_GIU_BFlasher1, LM_GIU_BFlasher2, LM_GIU_BRing1, LM_GIU_BRing2, LM_GIU_BSocket1, LM_GIU_BSocket2, LM_GIU_BSocket3, LM_GIU_CoffinDiverter, LM_GIU_Gate002, LM_GIU_Gate007, LM_GIU_GateLoop, LM_GIU_LSlingArmU, LM_GIU_LSlingU0, LM_GIU_LSlingU1, LM_GIU_LSlingU2, LM_GIU_Layer1, LM_GIU_Layer2, LM_GIU_Layer3, LM_GIU_Parts, LM_GIU_Playfield, LM_GIU_PrLeaper1, LM_GIU_PrLeaper2, LM_GIU_PrLeaper3, LM_GIU_crate_door, LM_GIU_crate_doorU, LM_GIU_lockdownbar, LM_GIU_sideblades, LM_GIU_sw25, LM_GIU_sw58, LM_GIU_sw68, LM_GIU_sw71, LM_GIU_sw72, LM_GIU_sw73, LM_GIU_sw74, LM_GIU_underPF)
'Dim BL_L_l11: BL_L_l11=Array(LM_L_l11_Playfield, LM_L_l11_underPF)
'Dim BL_L_l12: BL_L_l12=Array(LM_L_l12_Playfield, LM_L_l12_underPF)
'Dim BL_L_l13: BL_L_l13=Array(LM_L_l13_Playfield, LM_L_l13_underPF)
'Dim BL_L_l14: BL_L_l14=Array(LM_L_l14_Playfield, LM_L_l14_underPF)
'Dim BL_L_l15: BL_L_l15=Array(LM_L_l15_Parts, LM_L_l15_Playfield, LM_L_l15_underPF)
'Dim BL_L_l16: BL_L_l16=Array(LM_L_l16_Playfield, LM_L_l16_underPF)
'Dim BL_L_l17: BL_L_l17=Array(LM_L_l17_Playfield, LM_L_l17_underPF)
'Dim BL_L_l18: BL_L_l18=Array(LM_L_l18_Layer2, LM_L_l18_Layer3, LM_L_l18_Parts)
'Dim BL_L_l21: BL_L_l21=Array(LM_L_l21_Playfield, LM_L_l21_PrLeaper2, LM_L_l21_underPF)
'Dim BL_L_l22: BL_L_l22=Array(LM_L_l22_Playfield, LM_L_l22_PrLeaper2, LM_L_l22_underPF)
'Dim BL_L_l23: BL_L_l23=Array(LM_L_l23_Layer1, LM_L_l23_Layer2, LM_L_l23_Layer3, LM_L_l23_Parts, LM_L_l23_Playfield, LM_L_l23_crate_door, LM_L_l23_underPF)
'Dim BL_L_l24: BL_L_l24=Array(LM_L_l24_Playfield, LM_L_l24_PrLeaper2, LM_L_l24_sw65, LM_L_l24_underPF)
'Dim BL_L_l25: BL_L_l25=Array(LM_L_l25_Layer3, LM_L_l25_Playfield, LM_L_l25_sw61, LM_L_l25_sw62, LM_L_l25_sw63, LM_L_l25_underPF)
'Dim BL_L_l26: BL_L_l26=Array(LM_L_l26_Layer3, LM_L_l26_Playfield, LM_L_l26_sw61, LM_L_l26_sw62, LM_L_l26_sw63, LM_L_l26_underPF)
'Dim BL_L_l27: BL_L_l27=Array(LM_L_l27_Layer3, LM_L_l27_Playfield, LM_L_l27_sw61, LM_L_l27_sw62, LM_L_l27_sw63, LM_L_l27_underPF)
'Dim BL_L_l28: BL_L_l28=Array(LM_L_l28_Layer2, LM_L_l28_Playfield, LM_L_l28_underPF)
'Dim BL_L_l31: BL_L_l31=Array(LM_L_l31_Layer1, LM_L_l31_Layer2, LM_L_l31_Layer3, LM_L_l31_Parts, LM_L_l31_sw72)
'Dim BL_L_l32: BL_L_l32=Array(LM_L_l32_Layer1, LM_L_l32_Layer2, LM_L_l32_Layer3, LM_L_l32_Parts, LM_L_l32_sw72)
'Dim BL_L_l33: BL_L_l33=Array(LM_L_l33_Gate007, LM_L_l33_Layer1, LM_L_l33_Layer2, LM_L_l33_Layer3, LM_L_l33_Parts, LM_L_l33_Playfield)
'Dim BL_L_l34: BL_L_l34=Array(LM_L_l34_Gate007, LM_L_l34_Layer1, LM_L_l34_Layer2, LM_L_l34_Layer3, LM_L_l34_Parts, LM_L_l34_Playfield)
'Dim BL_L_l35: BL_L_l35=Array(LM_L_l35_Parts, LM_L_l35_Playfield, LM_L_l35_sideblades, LM_L_l35_underPF)
'Dim BL_L_l36: BL_L_l36=Array(LM_L_l36_Layer2, LM_L_l36_Playfield, LM_L_l36_PrLeaper3, LM_L_l36_sw66, LM_L_l36_underPF)
'Dim BL_L_l37: BL_L_l37=Array(LM_L_l37_Layer2, LM_L_l37_Parts, LM_L_l37_Playfield, LM_L_l37_underPF)
'Dim BL_L_l38: BL_L_l38=Array(LM_L_l38_Playfield, LM_L_l38_underPF)
'Dim BL_L_l41: BL_L_l41=Array(LM_L_l41_Layer3, LM_L_l41_Parts, LM_L_l41_Playfield, LM_L_l41_PrLeaper1, LM_L_l41_sw64, LM_L_l41_underPF)
'Dim BL_L_l42: BL_L_l42=Array(LM_L_l42_Layer3, LM_L_l42_Playfield, LM_L_l42_PrLeaper2, LM_L_l42_underPF)
'Dim BL_L_l43: BL_L_l43=Array(LM_L_l43_Playfield, LM_L_l43_underPF)
'Dim BL_L_l44: BL_L_l44=Array(LM_L_l44_Layer1, LM_L_l44_Layer2, LM_L_l44_Layer3, LM_L_l44_Parts)
'Dim BL_L_l45: BL_L_l45=Array(LM_L_l45_Layer2, LM_L_l45_Parts, LM_L_l45_Playfield, LM_L_l45_underPF)
'Dim BL_L_l46: BL_L_l46=Array(LM_L_l46_Layer3, LM_L_l46_Parts, LM_L_l46_Playfield, LM_L_l46_crate_door, LM_L_l46_crate_doorU, LM_L_l46_underPF)
'Dim BL_L_l47: BL_L_l47=Array(LM_L_l47_Layer2, LM_L_l47_Parts, LM_L_l47_Playfield, LM_L_l47_crate_door, LM_L_l47_underPF)
'Dim BL_L_l48: BL_L_l48=Array(LM_L_l48_Layer2, LM_L_l48_Parts, LM_L_l48_Playfield, LM_L_l48_underPF)
'Dim BL_L_l51: BL_L_l51=Array(LM_L_l51_Layer2, LM_L_l51_Parts, LM_L_l51_Playfield, LM_L_l51_RSling1, LM_L_l51_remk, LM_L_l51_underPF)
'Dim BL_L_l52: BL_L_l52=Array(LM_L_l52_Playfield, LM_L_l52_underPF)
'Dim BL_L_l53: BL_L_l53=Array(LM_L_l53_LBoogie02, LM_L_l53_LSling1, LM_L_l53_Layer3, LM_L_l53_Parts, LM_L_l53_Playfield, LM_L_l53_lemk, LM_L_l53_underPF)
'Dim BL_L_l54: BL_L_l54=Array(LM_L_l54_Layer3, LM_L_l54_Parts, LM_L_l54_Playfield, LM_L_l54_PrLeaper3rod, LM_L_l54_cDoorClose, LM_L_l54_cDoorOpen, LM_L_l54_cadaver, LM_L_l54_sideblades, LM_L_l54_underPF)
'Dim BL_L_l55: BL_L_l55=Array(LM_L_l55_Lflipper, LM_L_l55_LflipperU, LM_L_l55_Playfield, LM_L_l55_Rflipper, LM_L_l55_RflipperU, LM_L_l55_underPF)
'Dim BL_L_l56: BL_L_l56=Array(LM_L_l56_Layer3, LM_L_l56_Parts, LM_L_l56_Playfield, LM_L_l56_underPF)
'Dim BL_L_l56a: BL_L_l56a=Array(LM_L_l56a_CoffinDiverter, LM_L_l56a_Layer2, LM_L_l56a_Layer3, LM_L_l56a_Parts)
'Dim BL_L_l57: BL_L_l57=Array(LM_L_l57_Layer3, LM_L_l57_Parts, LM_L_l57_Playfield, LM_L_l57_PrLeaper1, LM_L_l57_underPF)
'Dim BL_L_l58: BL_L_l58=Array(LM_L_l58_CoffinDiverter, LM_L_l58_Layer3, LM_L_l58_Parts, LM_L_l58_Playfield, LM_L_l58_underPF)
'Dim BL_L_l61: BL_L_l61=Array(LM_L_l61_Playfield, LM_L_l61_underPF)
'Dim BL_L_l62: BL_L_l62=Array(LM_L_l62_Layer2, LM_L_l62_Layer3, LM_L_l62_Playfield, LM_L_l62_underPF)
'Dim BL_L_l63: BL_L_l63=Array(LM_L_l63_Layer3, LM_L_l63_Playfield, LM_L_l63_underPF)
'Dim BL_L_l84: BL_L_l84=Array(LM_L_l84_Layer1, LM_L_l84_Layer2, LM_L_l84_Layer3, LM_L_l84_Parts, LM_L_l84_Playfield, LM_L_l84_underPF)
'Dim BL_L_l85: BL_L_l85=Array(LM_L_l85_Layer1, LM_L_l85_Layer2, LM_L_l85_Parts, LM_L_l85_Playfield, LM_L_l85_underPF)
'Dim BL_L_l86: BL_L_l86=Array(LM_L_l86_Layer1, LM_L_l86_Layer2, LM_L_l86_Parts, LM_L_l86_Playfield, LM_L_l86_underPF)
'Dim BL_World: BL_World=Array(BM_Autoplunger, BM_BFlasher1, BM_BFlasher2, BM_BFlasher3, BM_BRing1, BM_BRing2, BM_BRing3, BM_BSocket1, BM_BSocket2, BM_BSocket3, BM_CoffinDiverter, BM_Gate001, BM_Gate002, BM_Gate003, BM_Gate005, BM_Gate007, BM_GateLoop, BM_LBoogie01, BM_LBoogie02, BM_LSilng0, BM_LSling1, BM_LSling2, BM_LSlingArmU, BM_LSlingU0, BM_LSlingU1, BM_LSlingU2, BM_Layer1, BM_Layer2, BM_Layer3, BM_Lflipper, BM_LflipperU, BM_Parts, BM_Playfield, BM_PrLeaper1, BM_PrLeaper1rod, BM_PrLeaper2, BM_PrLeaper2rod, BM_PrLeaper3, BM_PrLeaper3rod, BM_RBoggie01, BM_RBoggie02, BM_RSling0, BM_RSling1, BM_RSling2, BM_Rflipper, BM_RflipperU, BM_cDoorClose, BM_cDoorOpen, BM_cadaver, BM_crate_door, BM_crate_doorU, BM_lPost, BM_lemk, BM_lockdownbar, BM_rPost, BM_rampFlap, BM_remk, BM_sideblades, BM_sw16, BM_sw17, BM_sw18, BM_sw25, BM_sw26, BM_sw27, BM_sw28, BM_sw58, BM_sw61, BM_sw62, BM_sw63, BM_sw64, BM_sw65, BM_sw66, BM_sw68, BM_sw71, BM_sw72, BM_sw73, BM_sw74, BM_underPF)
' Global arrays
'Dim BG_Bakemap: BG_Bakemap=Array(BM_Autoplunger, BM_BFlasher1, BM_BFlasher2, BM_BFlasher3, BM_BRing1, BM_BRing2, BM_BRing3, BM_BSocket1, BM_BSocket2, BM_BSocket3, BM_CoffinDiverter, BM_Gate001, BM_Gate002, BM_Gate003, BM_Gate005, BM_Gate007, BM_GateLoop, BM_LBoogie01, BM_LBoogie02, BM_LSilng0, BM_LSling1, BM_LSling2, BM_LSlingArmU, BM_LSlingU0, BM_LSlingU1, BM_LSlingU2, BM_Layer1, BM_Layer2, BM_Layer3, BM_Lflipper, BM_LflipperU, BM_Parts, BM_Playfield, BM_PrLeaper1, BM_PrLeaper1rod, BM_PrLeaper2, BM_PrLeaper2rod, BM_PrLeaper3, BM_PrLeaper3rod, BM_RBoggie01, BM_RBoggie02, BM_RSling0, BM_RSling1, BM_RSling2, BM_Rflipper, BM_RflipperU, BM_cDoorClose, BM_cDoorOpen, BM_cadaver, BM_crate_door, BM_crate_doorU, BM_lPost, BM_lemk, BM_lockdownbar, BM_rPost, BM_rampFlap, BM_remk, BM_sideblades, BM_sw16, BM_sw17, BM_sw18, BM_sw25, BM_sw26, BM_sw27, BM_sw28, BM_sw58, BM_sw61, BM_sw62, BM_sw63, BM_sw64, BM_sw65, BM_sw66, BM_sw68, BM_sw71, BM_sw72, BM_sw73, BM_sw74, BM_underPF)
'Dim BG_Lightmap: BG_Lightmap=Array(LM_F_f17_BFlasher1, LM_F_f17_BFlasher3, LM_F_f17_BRing1, LM_F_f17_BRing2, LM_F_f17_BSocket1, LM_F_f17_BSocket2, LM_F_f17_BSocket3, LM_F_f17_Layer1, LM_F_f17_Layer2, LM_F_f17_Parts, LM_F_f17_Playfield, LM_F_f17_sideblades, LM_F_f18_BFlasher1, LM_F_f18_BFlasher2, LM_F_f18_BRing1, LM_F_f18_BRing2, LM_F_f18_BRing3, LM_F_f18_BSocket1, LM_F_f18_BSocket2, LM_F_f18_BSocket3, LM_F_f18_Layer1, LM_F_f18_Layer2, LM_F_f18_Layer3, LM_F_f18_Parts, LM_F_f18_Playfield, LM_F_f18_PrLeaper2, LM_F_f18_PrLeaper3, LM_F_f18_crate_door, LM_F_f18_crate_doorU, LM_F_f18_sideblades, LM_F_f19_BFlasher3, LM_F_f19_BRing1, LM_F_f19_BRing2, LM_F_f19_BRing3, LM_F_f19_BSocket1, LM_F_f19_BSocket2, LM_F_f19_BSocket3, LM_F_f19_Layer1, LM_F_f19_Layer2, LM_F_f19_Layer3, LM_F_f19_Parts, LM_F_f19_Playfield, LM_F_f19_sideblades, LM_F_f20_Layer2, LM_F_f20_Lflipper, LM_F_f20_LflipperU, LM_F_f20_Parts, LM_F_f20_Playfield, LM_F_f20_Rflipper, LM_F_f20_RflipperU, LM_F_f20_lemk, LM_F_f20_remk, LM_F_f20_underPF, _
' LM_F_f20a_Lflipper, LM_F_f20a_LflipperU, LM_F_f20a_Parts, LM_F_f20a_Playfield, LM_F_f20a_RBoggie01, LM_F_f20a_Rflipper, LM_F_f20a_RflipperU, LM_F_f20a_remk, LM_F_f20a_underPF, LM_F_f21_BSocket1, LM_F_f21_BSocket2, LM_F_f21_BSocket3, LM_F_f21_Layer1, LM_F_f21_Layer2, LM_F_f21_Layer3, LM_F_f21_Parts, LM_F_f21_Playfield, LM_F_f21_sideblades, LM_F_f21_sw68, LM_F_f21_underPF, LM_F_f22_BFlasher1, LM_F_f22_BRing1, LM_F_f22_BSocket1, LM_F_f22_CoffinDiverter, LM_F_f22_Gate007, LM_F_f22_GateLoop, LM_F_f22_Layer1, LM_F_f22_Layer2, LM_F_f22_Layer3, LM_F_f22_Parts, LM_F_f22_Playfield, LM_F_f22_PrLeaper1rod, LM_F_f22_sideblades, LM_F_f22_sw25, LM_F_f22_sw68, LM_F_f22_underPF, LM_F_f23_BRing1, LM_F_f23_CoffinDiverter, LM_F_f23_Gate007, LM_F_f23_LSlingArmU, LM_F_f23_Layer1, LM_F_f23_Layer2, LM_F_f23_Layer3, LM_F_f23_Parts, LM_F_f23_Playfield, LM_F_f23_PrLeaper1, LM_F_f23_PrLeaper1rod, LM_F_f23_PrLeaper2, LM_F_f23_PrLeaper2rod, LM_F_f23_sideblades, LM_F_f23_sw58, LM_F_f23_sw64, LM_F_f23_sw74, LM_F_f23_underPF, _
' LM_F_f23a_BRing1, LM_F_f23a_BRing2, LM_F_f23a_BRing3, LM_F_f23a_BSocket3, LM_F_f23a_CoffinDiverter, LM_F_f23a_LSlingArmU, LM_F_f23a_LSlingU0, LM_F_f23a_LSlingU1, LM_F_f23a_LSlingU2, LM_F_f23a_Layer1, LM_F_f23a_Layer2, LM_F_f23a_Layer3, LM_F_f23a_Parts, LM_F_f23a_Playfield, LM_F_f23a_PrLeaper1, LM_F_f23a_PrLeaper1rod, LM_F_f23a_PrLeaper2, LM_F_f23a_PrLeaper2rod, LM_F_f23a_PrLeaper3, LM_F_f23a_cDoorClose, LM_F_f23a_crate_door, LM_F_f23a_crate_doorU, LM_F_f23a_sideblades, LM_F_f23a_sw58, LM_F_f23a_sw64, LM_F_f23a_sw65, LM_F_f23a_sw74, LM_F_f23a_underPF, LM_F_f24_BRing2, LM_F_f24_CoffinDiverter, LM_F_f24_Gate003, LM_F_f24_Gate007, LM_F_f24_LSlingArmU, LM_F_f24_Layer1, LM_F_f24_Layer2, LM_F_f24_Layer3, LM_F_f24_Parts, LM_F_f24_Playfield, LM_F_f24_PrLeaper1, LM_F_f24_PrLeaper1rod, LM_F_f24_PrLeaper2, LM_F_f24_PrLeaper2rod, LM_F_f24_PrLeaper3, LM_F_f24_PrLeaper3rod, LM_F_f24_cDoorClose, LM_F_f24_cDoorOpen, LM_F_f24_crate_door, LM_F_f24_sideblades, LM_F_f24_sw58, LM_F_f24_sw64, LM_F_f24_sw65, LM_F_f24_sw74, _
' LM_F_f24_underPF, LM_F_f25_BRing1, LM_F_f25_BSocket1, LM_F_f25_BSocket2, LM_F_f25_Layer1, LM_F_f25_Layer2, LM_F_f25_Layer3, LM_F_f25_Parts, LM_F_f25_Playfield, LM_F_f25_sideblades, LM_F_f25_underPF, LM_F_f26_LBoogie01, LM_F_f26_LBoogie02, LM_F_f26_LSilng0, LM_F_f26_LSling1, LM_F_f26_Layer2, LM_F_f26_Layer3, LM_F_f26_Lflipper, LM_F_f26_LflipperU, LM_F_f26_Parts, LM_F_f26_Playfield, LM_F_f26_RBoggie01, LM_F_f26_RBoggie02, LM_F_f26_Rflipper, LM_F_f26_RflipperU, LM_F_f26_cDoorClose, LM_F_f26_lemk, LM_F_f26_remk, LM_F_f26_sideblades, LM_F_f26_sw28, LM_F_f26_sw61, LM_F_f26_sw62, LM_F_f26_sw63, LM_F_f26_sw66, LM_F_f26_underPF, LM_F_f27_Gate002, LM_F_f27_Layer1, LM_F_f27_Layer2, LM_F_f27_Layer3, LM_F_f27_Parts, LM_F_f27_Playfield, LM_F_f27_lockdownbar, LM_F_f27_sideblades, LM_F_f27_sw58, LM_F_f27_sw74, LM_F_f28_BFlasher3, LM_F_f28_BRing1, LM_F_f28_BRing2, LM_F_f28_BRing3, LM_F_f28_BSocket3, LM_F_f28_Gate001, LM_F_f28_Gate005, LM_F_f28_LSlingArmU, LM_F_f28_Layer1, LM_F_f28_Layer2, LM_F_f28_Layer3, LM_F_f28_Parts, _
' LM_F_f28_Playfield, LM_F_f28_PrLeaper2, LM_F_f28_PrLeaper3, LM_F_f28_crate_door, LM_F_f28_crate_doorU, LM_F_f28_rampFlap, LM_F_f28_sideblades, LM_F_f35_LBoogie02, LM_F_f35_LSilng0, LM_F_f35_Layer3, LM_F_f35_Parts, LM_F_f35_Playfield, LM_F_f35_PrLeaper3rod, LM_F_f35_RSling0, LM_F_f35_cDoorClose, LM_F_f35_cDoorOpen, LM_F_f35_cadaver, LM_F_f35_lemk, LM_F_f35_sideblades, LM_F_f35_sw16, LM_F_f35_sw61, LM_F_f35_sw62, LM_F_f35_sw63, LM_F_f35_underPF, LM_F_f36_Autoplunger, LM_F_f36_LSilng0, LM_F_f36_Layer1, LM_F_f36_Layer2, LM_F_f36_Layer3, LM_F_f36_Parts, LM_F_f36_Playfield, LM_F_f36_PrLeaper3rod, LM_F_f36_RBoggie02, LM_F_f36_cadaver, LM_F_f36_remk, LM_F_f36_sideblades, LM_F_f36_sw27, LM_F_f36_sw28, LM_F_f36_underPF, LM_GIC_BFlasher1, LM_GIC_BRing1, LM_GIC_BRing2, LM_GIC_BRing3, LM_GIC_BSocket1, LM_GIC_BSocket2, LM_GIC_BSocket3, LM_GIC_CoffinDiverter, LM_GIC_Gate001, LM_GIC_Gate003, LM_GIC_LSlingArmU, LM_GIC_LSlingU2, LM_GIC_Layer1, LM_GIC_Layer2, LM_GIC_Layer3, LM_GIC_Parts, LM_GIC_Playfield, LM_GIC_PrLeaper1, _
' LM_GIC_PrLeaper1rod, LM_GIC_PrLeaper2, LM_GIC_PrLeaper2rod, LM_GIC_PrLeaper3, LM_GIC_PrLeaper3rod, LM_GIC_cDoorClose, LM_GIC_cDoorOpen, LM_GIC_cadaver, LM_GIC_crate_door, LM_GIC_crate_doorU, LM_GIC_rampFlap, LM_GIC_sideblades, LM_GIC_sw61, LM_GIC_sw64, LM_GIC_sw65, LM_GIC_sw66, LM_GIC_underPF, LM_GIL_gi_l_001_LSilng0, LM_GIL_gi_l_001_LSling1, LM_GIL_gi_l_001_LSling2, LM_GIL_gi_l_001_Layer2, LM_GIL_gi_l_001_Layer3, LM_GIL_gi_l_001_Lflipper, LM_GIL_gi_l_001_LflipperU, LM_GIL_gi_l_001_Parts, LM_GIL_gi_l_001_Playfield, LM_GIL_gi_l_001_PrLeaper3rod, LM_GIL_gi_l_001_lPost, LM_GIL_gi_l_001_lemk, LM_GIL_gi_l_001_remk, LM_GIL_gi_l_001_sideblades, LM_GIL_gi_l_001_sw16, LM_GIL_gi_l_001_sw26, LM_GIL_gi_l_001_sw28, LM_GIL_gi_l_001_sw61, LM_GIL_gi_l_001_sw62, LM_GIL_gi_l_001_sw63, LM_GIL_gi_l_001_underPF, LM_GIL_gi_l_002_LSilng0, LM_GIL_gi_l_002_LSling1, LM_GIL_gi_l_002_LSling2, LM_GIL_gi_l_002_Layer2, LM_GIL_gi_l_002_Layer3, LM_GIL_gi_l_002_Lflipper, LM_GIL_gi_l_002_LflipperU, LM_GIL_gi_l_002_Parts, _
' LM_GIL_gi_l_002_Playfield, LM_GIL_gi_l_002_lemk, LM_GIL_gi_l_002_remk, LM_GIL_gi_l_002_sideblades, LM_GIL_gi_l_002_sw16, LM_GIL_gi_l_002_sw26, LM_GIL_gi_l_002_sw62, LM_GIL_gi_l_002_sw63, LM_GIL_gi_l_003_LBoogie01, LM_GIL_gi_l_003_Layer3, LM_GIL_gi_l_003_Lflipper, LM_GIL_gi_l_003_LflipperU, LM_GIL_gi_l_003_Parts, LM_GIL_gi_l_003_Playfield, LM_GIL_gi_l_003_Rflipper, LM_GIL_gi_l_003_RflipperU, LM_GIL_gi_l_003_lemk, LM_GIL_gi_l_003_remk, LM_GIL_gi_l_003_sideblades, LM_GIL_gi_l_003_sw16, LM_GIL_gi_l_003_sw26, LM_GIL_gi_l_003_sw61, LM_GIL_gi_l_003_sw62, LM_GIL_gi_l_003_sw63, LM_GIL_gi_l_004_LBoogie01, LM_GIL_gi_l_004_Lflipper, LM_GIL_gi_l_004_LflipperU, LM_GIL_gi_l_004_Parts, LM_GIL_gi_l_004_Playfield, LM_GIL_gi_l_004_Rflipper, LM_GIL_gi_l_004_RflipperU, LM_GIL_gi_l_004_lemk, LM_GIL_gi_l_004_remk, LM_GIL_gi_l_004_sideblades, LM_GIL_gi_l_004_sw16, LM_GIL_gi_l_004_sw63, LM_GIL_gi_l_005_Layer2, LM_GIL_gi_l_005_Layer3, LM_GIL_gi_l_005_Lflipper, LM_GIL_gi_l_005_LflipperU, LM_GIL_gi_l_005_Parts, _
' LM_GIL_gi_l_005_Playfield, LM_GIL_gi_l_005_RBoggie01, LM_GIL_gi_l_005_Rflipper, LM_GIL_gi_l_005_RflipperU, LM_GIL_gi_l_005_lemk, LM_GIL_gi_l_005_remk, LM_GIL_gi_l_005_sw17, LM_GIL_gi_l_005_sw28, LM_GIL_gi_l_006_Layer2, LM_GIL_gi_l_006_Lflipper, LM_GIL_gi_l_006_LflipperU, LM_GIL_gi_l_006_Parts, LM_GIL_gi_l_006_Playfield, LM_GIL_gi_l_006_RBoggie01, LM_GIL_gi_l_006_Rflipper, LM_GIL_gi_l_006_RflipperU, LM_GIL_gi_l_006_lemk, LM_GIL_gi_l_006_remk, LM_GIL_gi_l_006_sideblades, LM_GIL_gi_l_006_sw17, LM_GIL_gi_l_006_sw28, LM_GIL_gi_l_007_Layer2, LM_GIL_gi_l_007_Layer3, LM_GIL_gi_l_007_Lflipper, LM_GIL_gi_l_007_LflipperU, LM_GIL_gi_l_007_Parts, LM_GIL_gi_l_007_Playfield, LM_GIL_gi_l_007_PrLeaper2rod, LM_GIL_gi_l_007_RBoggie01, LM_GIL_gi_l_007_RSling0, LM_GIL_gi_l_007_RSling2, LM_GIL_gi_l_007_Rflipper, LM_GIL_gi_l_007_RflipperU, LM_GIL_gi_l_007_lemk, LM_GIL_gi_l_007_remk, LM_GIL_gi_l_007_sideblades, LM_GIL_gi_l_007_sw17, LM_GIL_gi_l_007_sw28, LM_GIL_gi_l_008_Layer2, LM_GIL_gi_l_008_Layer3, LM_GIL_gi_l_008_Parts, _
' LM_GIL_gi_l_008_Playfield, LM_GIL_gi_l_008_RSling0, LM_GIL_gi_l_008_RSling1, LM_GIL_gi_l_008_RSling2, LM_GIL_gi_l_008_Rflipper, LM_GIL_gi_l_008_RflipperU, LM_GIL_gi_l_008_lemk, LM_GIL_gi_l_008_rPost, LM_GIL_gi_l_008_remk, LM_GIL_gi_l_008_sideblades, LM_GIL_gi_l_008_sw16, LM_GIL_gi_l_008_sw17, LM_GIL_gi_l_008_sw27, LM_GIL_gi_l_008_sw28, LM_GIL_gi_l_008_sw61, LM_GIL_gi_l_008_sw62, LM_GIL_gi_l_008_sw63, LM_GIL_gi_l_009_Layer3, LM_GIL_gi_l_009_Parts, LM_GIL_gi_l_009_Playfield, LM_GIL_gi_l_009_cDoorClose, LM_GIL_gi_l_009_cDoorOpen, LM_GIL_gi_l_009_cadaver, LM_GIL_gi_l_009_lPost, LM_GIL_gi_l_009_sideblades, LM_GIL_gi_l_009_sw61, LM_GIL_gi_l_009_sw62, LM_GIL_gi_l_009_sw63, LM_GIL_gi_l_010_Layer3, LM_GIL_gi_l_010_Parts, LM_GIL_gi_l_010_Playfield, LM_GIL_gi_l_010_cDoorClose, LM_GIL_gi_l_010_cDoorOpen, LM_GIL_gi_l_010_lPost, LM_GIL_gi_l_010_sideblades, LM_GIL_gi_l_010_sw61, LM_GIL_gi_l_010_sw62, LM_GIL_gi_l_010_sw63, LM_GIL_gi_l_011_Gate001, LM_GIL_gi_l_011_Layer2, LM_GIL_gi_l_011_Layer3, LM_GIL_gi_l_011_Parts, _
' LM_GIL_gi_l_011_Playfield, LM_GIL_gi_l_011_rPost, LM_GIL_gi_l_011_sideblades, LM_GIL_gi_l_011_sw17, LM_GIL_gi_l_011_sw27, LM_GIL_gi_l_011_sw28, LM_GIL_gi_l_011_underPF, LM_GIL_gi_l_012_Gate001, LM_GIL_gi_l_012_Layer2, LM_GIL_gi_l_012_Layer3, LM_GIL_gi_l_012_Parts, LM_GIL_gi_l_012_Playfield, LM_GIL_gi_l_012_PrLeaper2rod, LM_GIL_gi_l_012_rPost, LM_GIL_gi_l_012_remk, LM_GIL_gi_l_012_sideblades, LM_GIL_gi_l_012_sw17, LM_GIL_gi_l_012_sw27, LM_GIL_gi_l_012_sw28, LM_GIL_gi_l_012_underPF, LM_GIU_BFlasher1, LM_GIU_BFlasher2, LM_GIU_BRing1, LM_GIU_BRing2, LM_GIU_BSocket1, LM_GIU_BSocket2, LM_GIU_BSocket3, LM_GIU_CoffinDiverter, LM_GIU_Gate002, LM_GIU_Gate007, LM_GIU_GateLoop, LM_GIU_LSlingArmU, LM_GIU_LSlingU0, LM_GIU_LSlingU1, LM_GIU_LSlingU2, LM_GIU_Layer1, LM_GIU_Layer2, LM_GIU_Layer3, LM_GIU_Parts, LM_GIU_Playfield, LM_GIU_PrLeaper1, LM_GIU_PrLeaper2, LM_GIU_PrLeaper3, LM_GIU_crate_door, LM_GIU_crate_doorU, LM_GIU_lockdownbar, LM_GIU_sideblades, LM_GIU_sw25, LM_GIU_sw58, LM_GIU_sw68, LM_GIU_sw71, LM_GIU_sw72, _
' LM_GIU_sw73, LM_GIU_sw74, LM_GIU_underPF, LM_L_l11_Playfield, LM_L_l11_underPF, LM_L_l12_Playfield, LM_L_l12_underPF, LM_L_l13_Playfield, LM_L_l13_underPF, LM_L_l14_Playfield, LM_L_l14_underPF, LM_L_l15_Parts, LM_L_l15_Playfield, LM_L_l15_underPF, LM_L_l16_Playfield, LM_L_l16_underPF, LM_L_l17_Playfield, LM_L_l17_underPF, LM_L_l18_Layer2, LM_L_l18_Layer3, LM_L_l18_Parts, LM_L_l21_Playfield, LM_L_l21_PrLeaper2, LM_L_l21_underPF, LM_L_l22_Playfield, LM_L_l22_PrLeaper2, LM_L_l22_underPF, LM_L_l23_Layer1, LM_L_l23_Layer2, LM_L_l23_Layer3, LM_L_l23_Parts, LM_L_l23_Playfield, LM_L_l23_crate_door, LM_L_l23_underPF, LM_L_l24_Playfield, LM_L_l24_PrLeaper2, LM_L_l24_sw65, LM_L_l24_underPF, LM_L_l25_Layer3, LM_L_l25_Playfield, LM_L_l25_sw61, LM_L_l25_sw62, LM_L_l25_sw63, LM_L_l25_underPF, LM_L_l26_Layer3, LM_L_l26_Playfield, LM_L_l26_sw61, LM_L_l26_sw62, LM_L_l26_sw63, LM_L_l26_underPF, LM_L_l27_Layer3, LM_L_l27_Playfield, LM_L_l27_sw61, LM_L_l27_sw62, LM_L_l27_sw63, LM_L_l27_underPF, LM_L_l28_Layer2, _
' LM_L_l28_Playfield, LM_L_l28_underPF, LM_L_l31_Layer1, LM_L_l31_Layer2, LM_L_l31_Layer3, LM_L_l31_Parts, LM_L_l31_sw72, LM_L_l32_Layer1, LM_L_l32_Layer2, LM_L_l32_Layer3, LM_L_l32_Parts, LM_L_l32_sw72, LM_L_l33_Gate007, LM_L_l33_Layer1, LM_L_l33_Layer2, LM_L_l33_Layer3, LM_L_l33_Parts, LM_L_l33_Playfield, LM_L_l34_Gate007, LM_L_l34_Layer1, LM_L_l34_Layer2, LM_L_l34_Layer3, LM_L_l34_Parts, LM_L_l34_Playfield, LM_L_l35_Parts, LM_L_l35_Playfield, LM_L_l35_sideblades, LM_L_l35_underPF, LM_L_l36_Layer2, LM_L_l36_Playfield, LM_L_l36_PrLeaper3, LM_L_l36_sw66, LM_L_l36_underPF, LM_L_l37_Layer2, LM_L_l37_Parts, LM_L_l37_Playfield, LM_L_l37_underPF, LM_L_l38_Playfield, LM_L_l38_underPF, LM_L_l41_Layer3, LM_L_l41_Parts, LM_L_l41_Playfield, LM_L_l41_PrLeaper1, LM_L_l41_sw64, LM_L_l41_underPF, LM_L_l42_Layer3, LM_L_l42_Playfield, LM_L_l42_PrLeaper2, LM_L_l42_underPF, LM_L_l43_Playfield, LM_L_l43_underPF, LM_L_l44_Layer1, LM_L_l44_Layer2, LM_L_l44_Layer3, LM_L_l44_Parts, LM_L_l45_Layer2, LM_L_l45_Parts, LM_L_l45_Playfield, _
' LM_L_l45_underPF, LM_L_l46_Layer3, LM_L_l46_Parts, LM_L_l46_Playfield, LM_L_l46_crate_door, LM_L_l46_crate_doorU, LM_L_l46_underPF, LM_L_l47_Layer2, LM_L_l47_Parts, LM_L_l47_Playfield, LM_L_l47_crate_door, LM_L_l47_underPF, LM_L_l48_Layer2, LM_L_l48_Parts, LM_L_l48_Playfield, LM_L_l48_underPF, LM_L_l51_Layer2, LM_L_l51_Parts, LM_L_l51_Playfield, LM_L_l51_RSling1, LM_L_l51_remk, LM_L_l51_underPF, LM_L_l52_Playfield, LM_L_l52_underPF, LM_L_l53_LBoogie02, LM_L_l53_LSling1, LM_L_l53_Layer3, LM_L_l53_Parts, LM_L_l53_Playfield, LM_L_l53_lemk, LM_L_l53_underPF, LM_L_l54_Layer3, LM_L_l54_Parts, LM_L_l54_Playfield, LM_L_l54_PrLeaper3rod, LM_L_l54_cDoorClose, LM_L_l54_cDoorOpen, LM_L_l54_cadaver, LM_L_l54_sideblades, LM_L_l54_underPF, LM_L_l55_Lflipper, LM_L_l55_LflipperU, LM_L_l55_Playfield, LM_L_l55_Rflipper, LM_L_l55_RflipperU, LM_L_l55_underPF, LM_L_l56_Layer3, LM_L_l56_Parts, LM_L_l56_Playfield, LM_L_l56_underPF, LM_L_l56a_CoffinDiverter, LM_L_l56a_Layer2, LM_L_l56a_Layer3, LM_L_l56a_Parts, LM_L_l57_Layer3, _
' LM_L_l57_Parts, LM_L_l57_Playfield, LM_L_l57_PrLeaper1, LM_L_l57_underPF, LM_L_l58_CoffinDiverter, LM_L_l58_Layer3, LM_L_l58_Parts, LM_L_l58_Playfield, LM_L_l58_underPF, LM_L_l61_Playfield, LM_L_l61_underPF, LM_L_l62_Layer2, LM_L_l62_Layer3, LM_L_l62_Playfield, LM_L_l62_underPF, LM_L_l63_Layer3, LM_L_l63_Playfield, LM_L_l63_underPF, LM_L_l84_Layer1, LM_L_l84_Layer2, LM_L_l84_Layer3, LM_L_l84_Parts, LM_L_l84_Playfield, LM_L_l84_underPF, LM_L_l85_Layer1, LM_L_l85_Layer2, LM_L_l85_Parts, LM_L_l85_Playfield, LM_L_l85_underPF, LM_L_l86_Layer1, LM_L_l86_Layer2, LM_L_l86_Parts, LM_L_l86_Playfield, LM_L_l86_underPF)
'Dim BG_All: BG_All=Array(BM_Autoplunger, BM_BFlasher1, BM_BFlasher2, BM_BFlasher3, BM_BRing1, BM_BRing2, BM_BRing3, BM_BSocket1, BM_BSocket2, BM_BSocket3, BM_CoffinDiverter, BM_Gate001, BM_Gate002, BM_Gate003, BM_Gate005, BM_Gate007, BM_GateLoop, BM_LBoogie01, BM_LBoogie02, BM_LSilng0, BM_LSling1, BM_LSling2, BM_LSlingArmU, BM_LSlingU0, BM_LSlingU1, BM_LSlingU2, BM_Layer1, BM_Layer2, BM_Layer3, BM_Lflipper, BM_LflipperU, BM_Parts, BM_Playfield, BM_PrLeaper1, BM_PrLeaper1rod, BM_PrLeaper2, BM_PrLeaper2rod, BM_PrLeaper3, BM_PrLeaper3rod, BM_RBoggie01, BM_RBoggie02, BM_RSling0, BM_RSling1, BM_RSling2, BM_Rflipper, BM_RflipperU, BM_cDoorClose, BM_cDoorOpen, BM_cadaver, BM_crate_door, BM_crate_doorU, BM_lPost, BM_lemk, BM_lockdownbar, BM_rPost, BM_rampFlap, BM_remk, BM_sideblades, BM_sw16, BM_sw17, BM_sw18, BM_sw25, BM_sw26, BM_sw27, BM_sw28, BM_sw58, BM_sw61, BM_sw62, BM_sw63, BM_sw64, BM_sw65, BM_sw66, BM_sw68, BM_sw71, BM_sw72, BM_sw73, BM_sw74, BM_underPF, LM_F_f17_BFlasher1, LM_F_f17_BFlasher3, _
' LM_F_f17_BRing1, LM_F_f17_BRing2, LM_F_f17_BSocket1, LM_F_f17_BSocket2, LM_F_f17_BSocket3, LM_F_f17_Layer1, LM_F_f17_Layer2, LM_F_f17_Parts, LM_F_f17_Playfield, LM_F_f17_sideblades, LM_F_f18_BFlasher1, LM_F_f18_BFlasher2, LM_F_f18_BRing1, LM_F_f18_BRing2, LM_F_f18_BRing3, LM_F_f18_BSocket1, LM_F_f18_BSocket2, LM_F_f18_BSocket3, LM_F_f18_Layer1, LM_F_f18_Layer2, LM_F_f18_Layer3, LM_F_f18_Parts, LM_F_f18_Playfield, LM_F_f18_PrLeaper2, LM_F_f18_PrLeaper3, LM_F_f18_crate_door, LM_F_f18_crate_doorU, LM_F_f18_sideblades, LM_F_f19_BFlasher3, LM_F_f19_BRing1, LM_F_f19_BRing2, LM_F_f19_BRing3, LM_F_f19_BSocket1, LM_F_f19_BSocket2, LM_F_f19_BSocket3, LM_F_f19_Layer1, LM_F_f19_Layer2, LM_F_f19_Layer3, LM_F_f19_Parts, LM_F_f19_Playfield, LM_F_f19_sideblades, LM_F_f20_Layer2, LM_F_f20_Lflipper, LM_F_f20_LflipperU, LM_F_f20_Parts, LM_F_f20_Playfield, LM_F_f20_Rflipper, LM_F_f20_RflipperU, LM_F_f20_lemk, LM_F_f20_remk, LM_F_f20_underPF, LM_F_f20a_Lflipper, LM_F_f20a_LflipperU, LM_F_f20a_Parts, LM_F_f20a_Playfield, _
' LM_F_f20a_RBoggie01, LM_F_f20a_Rflipper, LM_F_f20a_RflipperU, LM_F_f20a_remk, LM_F_f20a_underPF, LM_F_f21_BSocket1, LM_F_f21_BSocket2, LM_F_f21_BSocket3, LM_F_f21_Layer1, LM_F_f21_Layer2, LM_F_f21_Layer3, LM_F_f21_Parts, LM_F_f21_Playfield, LM_F_f21_sideblades, LM_F_f21_sw68, LM_F_f21_underPF, LM_F_f22_BFlasher1, LM_F_f22_BRing1, LM_F_f22_BSocket1, LM_F_f22_CoffinDiverter, LM_F_f22_Gate007, LM_F_f22_GateLoop, LM_F_f22_Layer1, LM_F_f22_Layer2, LM_F_f22_Layer3, LM_F_f22_Parts, LM_F_f22_Playfield, LM_F_f22_PrLeaper1rod, LM_F_f22_sideblades, LM_F_f22_sw25, LM_F_f22_sw68, LM_F_f22_underPF, LM_F_f23_BRing1, LM_F_f23_CoffinDiverter, LM_F_f23_Gate007, LM_F_f23_LSlingArmU, LM_F_f23_Layer1, LM_F_f23_Layer2, LM_F_f23_Layer3, LM_F_f23_Parts, LM_F_f23_Playfield, LM_F_f23_PrLeaper1, LM_F_f23_PrLeaper1rod, LM_F_f23_PrLeaper2, LM_F_f23_PrLeaper2rod, LM_F_f23_sideblades, LM_F_f23_sw58, LM_F_f23_sw64, LM_F_f23_sw74, LM_F_f23_underPF, LM_F_f23a_BRing1, LM_F_f23a_BRing2, LM_F_f23a_BRing3, LM_F_f23a_BSocket3, _
' LM_F_f23a_CoffinDiverter, LM_F_f23a_LSlingArmU, LM_F_f23a_LSlingU0, LM_F_f23a_LSlingU1, LM_F_f23a_LSlingU2, LM_F_f23a_Layer1, LM_F_f23a_Layer2, LM_F_f23a_Layer3, LM_F_f23a_Parts, LM_F_f23a_Playfield, LM_F_f23a_PrLeaper1, LM_F_f23a_PrLeaper1rod, LM_F_f23a_PrLeaper2, LM_F_f23a_PrLeaper2rod, LM_F_f23a_PrLeaper3, LM_F_f23a_cDoorClose, LM_F_f23a_crate_door, LM_F_f23a_crate_doorU, LM_F_f23a_sideblades, LM_F_f23a_sw58, LM_F_f23a_sw64, LM_F_f23a_sw65, LM_F_f23a_sw74, LM_F_f23a_underPF, LM_F_f24_BRing2, LM_F_f24_CoffinDiverter, LM_F_f24_Gate003, LM_F_f24_Gate007, LM_F_f24_LSlingArmU, LM_F_f24_Layer1, LM_F_f24_Layer2, LM_F_f24_Layer3, LM_F_f24_Parts, LM_F_f24_Playfield, LM_F_f24_PrLeaper1, LM_F_f24_PrLeaper1rod, LM_F_f24_PrLeaper2, LM_F_f24_PrLeaper2rod, LM_F_f24_PrLeaper3, LM_F_f24_PrLeaper3rod, LM_F_f24_cDoorClose, LM_F_f24_cDoorOpen, LM_F_f24_crate_door, LM_F_f24_sideblades, LM_F_f24_sw58, LM_F_f24_sw64, LM_F_f24_sw65, LM_F_f24_sw74, LM_F_f24_underPF, LM_F_f25_BRing1, LM_F_f25_BSocket1, LM_F_f25_BSocket2, _
' LM_F_f25_Layer1, LM_F_f25_Layer2, LM_F_f25_Layer3, LM_F_f25_Parts, LM_F_f25_Playfield, LM_F_f25_sideblades, LM_F_f25_underPF, LM_F_f26_LBoogie01, LM_F_f26_LBoogie02, LM_F_f26_LSilng0, LM_F_f26_LSling1, LM_F_f26_Layer2, LM_F_f26_Layer3, LM_F_f26_Lflipper, LM_F_f26_LflipperU, LM_F_f26_Parts, LM_F_f26_Playfield, LM_F_f26_RBoggie01, LM_F_f26_RBoggie02, LM_F_f26_Rflipper, LM_F_f26_RflipperU, LM_F_f26_cDoorClose, LM_F_f26_lemk, LM_F_f26_remk, LM_F_f26_sideblades, LM_F_f26_sw28, LM_F_f26_sw61, LM_F_f26_sw62, LM_F_f26_sw63, LM_F_f26_sw66, LM_F_f26_underPF, LM_F_f27_Gate002, LM_F_f27_Layer1, LM_F_f27_Layer2, LM_F_f27_Layer3, LM_F_f27_Parts, LM_F_f27_Playfield, LM_F_f27_lockdownbar, LM_F_f27_sideblades, LM_F_f27_sw58, LM_F_f27_sw74, LM_F_f28_BFlasher3, LM_F_f28_BRing1, LM_F_f28_BRing2, LM_F_f28_BRing3, LM_F_f28_BSocket3, LM_F_f28_Gate001, LM_F_f28_Gate005, LM_F_f28_LSlingArmU, LM_F_f28_Layer1, LM_F_f28_Layer2, LM_F_f28_Layer3, LM_F_f28_Parts, LM_F_f28_Playfield, LM_F_f28_PrLeaper2, LM_F_f28_PrLeaper3, _
' LM_F_f28_crate_door, LM_F_f28_crate_doorU, LM_F_f28_rampFlap, LM_F_f28_sideblades, LM_F_f35_LBoogie02, LM_F_f35_LSilng0, LM_F_f35_Layer3, LM_F_f35_Parts, LM_F_f35_Playfield, LM_F_f35_PrLeaper3rod, LM_F_f35_RSling0, LM_F_f35_cDoorClose, LM_F_f35_cDoorOpen, LM_F_f35_cadaver, LM_F_f35_lemk, LM_F_f35_sideblades, LM_F_f35_sw16, LM_F_f35_sw61, LM_F_f35_sw62, LM_F_f35_sw63, LM_F_f35_underPF, LM_F_f36_Autoplunger, LM_F_f36_LSilng0, LM_F_f36_Layer1, LM_F_f36_Layer2, LM_F_f36_Layer3, LM_F_f36_Parts, LM_F_f36_Playfield, LM_F_f36_PrLeaper3rod, LM_F_f36_RBoggie02, LM_F_f36_cadaver, LM_F_f36_remk, LM_F_f36_sideblades, LM_F_f36_sw27, LM_F_f36_sw28, LM_F_f36_underPF, LM_GIC_BFlasher1, LM_GIC_BRing1, LM_GIC_BRing2, LM_GIC_BRing3, LM_GIC_BSocket1, LM_GIC_BSocket2, LM_GIC_BSocket3, LM_GIC_CoffinDiverter, LM_GIC_Gate001, LM_GIC_Gate003, LM_GIC_LSlingArmU, LM_GIC_LSlingU2, LM_GIC_Layer1, LM_GIC_Layer2, LM_GIC_Layer3, LM_GIC_Parts, LM_GIC_Playfield, LM_GIC_PrLeaper1, LM_GIC_PrLeaper1rod, LM_GIC_PrLeaper2, LM_GIC_PrLeaper2rod, _
' LM_GIC_PrLeaper3, LM_GIC_PrLeaper3rod, LM_GIC_cDoorClose, LM_GIC_cDoorOpen, LM_GIC_cadaver, LM_GIC_crate_door, LM_GIC_crate_doorU, LM_GIC_rampFlap, LM_GIC_sideblades, LM_GIC_sw61, LM_GIC_sw64, LM_GIC_sw65, LM_GIC_sw66, LM_GIC_underPF, LM_GIL_gi_l_001_LSilng0, LM_GIL_gi_l_001_LSling1, LM_GIL_gi_l_001_LSling2, LM_GIL_gi_l_001_Layer2, LM_GIL_gi_l_001_Layer3, LM_GIL_gi_l_001_Lflipper, LM_GIL_gi_l_001_LflipperU, LM_GIL_gi_l_001_Parts, LM_GIL_gi_l_001_Playfield, LM_GIL_gi_l_001_PrLeaper3rod, LM_GIL_gi_l_001_lPost, LM_GIL_gi_l_001_lemk, LM_GIL_gi_l_001_remk, LM_GIL_gi_l_001_sideblades, LM_GIL_gi_l_001_sw16, LM_GIL_gi_l_001_sw26, LM_GIL_gi_l_001_sw28, LM_GIL_gi_l_001_sw61, LM_GIL_gi_l_001_sw62, LM_GIL_gi_l_001_sw63, LM_GIL_gi_l_001_underPF, LM_GIL_gi_l_002_LSilng0, LM_GIL_gi_l_002_LSling1, LM_GIL_gi_l_002_LSling2, LM_GIL_gi_l_002_Layer2, LM_GIL_gi_l_002_Layer3, LM_GIL_gi_l_002_Lflipper, LM_GIL_gi_l_002_LflipperU, LM_GIL_gi_l_002_Parts, LM_GIL_gi_l_002_Playfield, LM_GIL_gi_l_002_lemk, LM_GIL_gi_l_002_remk, _
' LM_GIL_gi_l_002_sideblades, LM_GIL_gi_l_002_sw16, LM_GIL_gi_l_002_sw26, LM_GIL_gi_l_002_sw62, LM_GIL_gi_l_002_sw63, LM_GIL_gi_l_003_LBoogie01, LM_GIL_gi_l_003_Layer3, LM_GIL_gi_l_003_Lflipper, LM_GIL_gi_l_003_LflipperU, LM_GIL_gi_l_003_Parts, LM_GIL_gi_l_003_Playfield, LM_GIL_gi_l_003_Rflipper, LM_GIL_gi_l_003_RflipperU, LM_GIL_gi_l_003_lemk, LM_GIL_gi_l_003_remk, LM_GIL_gi_l_003_sideblades, LM_GIL_gi_l_003_sw16, LM_GIL_gi_l_003_sw26, LM_GIL_gi_l_003_sw61, LM_GIL_gi_l_003_sw62, LM_GIL_gi_l_003_sw63, LM_GIL_gi_l_004_LBoogie01, LM_GIL_gi_l_004_Lflipper, LM_GIL_gi_l_004_LflipperU, LM_GIL_gi_l_004_Parts, LM_GIL_gi_l_004_Playfield, LM_GIL_gi_l_004_Rflipper, LM_GIL_gi_l_004_RflipperU, LM_GIL_gi_l_004_lemk, LM_GIL_gi_l_004_remk, LM_GIL_gi_l_004_sideblades, LM_GIL_gi_l_004_sw16, LM_GIL_gi_l_004_sw63, LM_GIL_gi_l_005_Layer2, LM_GIL_gi_l_005_Layer3, LM_GIL_gi_l_005_Lflipper, LM_GIL_gi_l_005_LflipperU, LM_GIL_gi_l_005_Parts, LM_GIL_gi_l_005_Playfield, LM_GIL_gi_l_005_RBoggie01, LM_GIL_gi_l_005_Rflipper, _
' LM_GIL_gi_l_005_RflipperU, LM_GIL_gi_l_005_lemk, LM_GIL_gi_l_005_remk, LM_GIL_gi_l_005_sw17, LM_GIL_gi_l_005_sw28, LM_GIL_gi_l_006_Layer2, LM_GIL_gi_l_006_Lflipper, LM_GIL_gi_l_006_LflipperU, LM_GIL_gi_l_006_Parts, LM_GIL_gi_l_006_Playfield, LM_GIL_gi_l_006_RBoggie01, LM_GIL_gi_l_006_Rflipper, LM_GIL_gi_l_006_RflipperU, LM_GIL_gi_l_006_lemk, LM_GIL_gi_l_006_remk, LM_GIL_gi_l_006_sideblades, LM_GIL_gi_l_006_sw17, LM_GIL_gi_l_006_sw28, LM_GIL_gi_l_007_Layer2, LM_GIL_gi_l_007_Layer3, LM_GIL_gi_l_007_Lflipper, LM_GIL_gi_l_007_LflipperU, LM_GIL_gi_l_007_Parts, LM_GIL_gi_l_007_Playfield, LM_GIL_gi_l_007_PrLeaper2rod, LM_GIL_gi_l_007_RBoggie01, LM_GIL_gi_l_007_RSling0, LM_GIL_gi_l_007_RSling2, LM_GIL_gi_l_007_Rflipper, LM_GIL_gi_l_007_RflipperU, LM_GIL_gi_l_007_lemk, LM_GIL_gi_l_007_remk, LM_GIL_gi_l_007_sideblades, LM_GIL_gi_l_007_sw17, LM_GIL_gi_l_007_sw28, LM_GIL_gi_l_008_Layer2, LM_GIL_gi_l_008_Layer3, LM_GIL_gi_l_008_Parts, LM_GIL_gi_l_008_Playfield, LM_GIL_gi_l_008_RSling0, LM_GIL_gi_l_008_RSling1, _
' LM_GIL_gi_l_008_RSling2, LM_GIL_gi_l_008_Rflipper, LM_GIL_gi_l_008_RflipperU, LM_GIL_gi_l_008_lemk, LM_GIL_gi_l_008_rPost, LM_GIL_gi_l_008_remk, LM_GIL_gi_l_008_sideblades, LM_GIL_gi_l_008_sw16, LM_GIL_gi_l_008_sw17, LM_GIL_gi_l_008_sw27, LM_GIL_gi_l_008_sw28, LM_GIL_gi_l_008_sw61, LM_GIL_gi_l_008_sw62, LM_GIL_gi_l_008_sw63, LM_GIL_gi_l_009_Layer3, LM_GIL_gi_l_009_Parts, LM_GIL_gi_l_009_Playfield, LM_GIL_gi_l_009_cDoorClose, LM_GIL_gi_l_009_cDoorOpen, LM_GIL_gi_l_009_cadaver, LM_GIL_gi_l_009_lPost, LM_GIL_gi_l_009_sideblades, LM_GIL_gi_l_009_sw61, LM_GIL_gi_l_009_sw62, LM_GIL_gi_l_009_sw63, LM_GIL_gi_l_010_Layer3, LM_GIL_gi_l_010_Parts, LM_GIL_gi_l_010_Playfield, LM_GIL_gi_l_010_cDoorClose, LM_GIL_gi_l_010_cDoorOpen, LM_GIL_gi_l_010_lPost, LM_GIL_gi_l_010_sideblades, LM_GIL_gi_l_010_sw61, LM_GIL_gi_l_010_sw62, LM_GIL_gi_l_010_sw63, LM_GIL_gi_l_011_Gate001, LM_GIL_gi_l_011_Layer2, LM_GIL_gi_l_011_Layer3, LM_GIL_gi_l_011_Parts, LM_GIL_gi_l_011_Playfield, LM_GIL_gi_l_011_rPost, LM_GIL_gi_l_011_sideblades, _
' LM_GIL_gi_l_011_sw17, LM_GIL_gi_l_011_sw27, LM_GIL_gi_l_011_sw28, LM_GIL_gi_l_011_underPF, LM_GIL_gi_l_012_Gate001, LM_GIL_gi_l_012_Layer2, LM_GIL_gi_l_012_Layer3, LM_GIL_gi_l_012_Parts, LM_GIL_gi_l_012_Playfield, LM_GIL_gi_l_012_PrLeaper2rod, LM_GIL_gi_l_012_rPost, LM_GIL_gi_l_012_remk, LM_GIL_gi_l_012_sideblades, LM_GIL_gi_l_012_sw17, LM_GIL_gi_l_012_sw27, LM_GIL_gi_l_012_sw28, LM_GIL_gi_l_012_underPF, LM_GIU_BFlasher1, LM_GIU_BFlasher2, LM_GIU_BRing1, LM_GIU_BRing2, LM_GIU_BSocket1, LM_GIU_BSocket2, LM_GIU_BSocket3, LM_GIU_CoffinDiverter, LM_GIU_Gate002, LM_GIU_Gate007, LM_GIU_GateLoop, LM_GIU_LSlingArmU, LM_GIU_LSlingU0, LM_GIU_LSlingU1, LM_GIU_LSlingU2, LM_GIU_Layer1, LM_GIU_Layer2, LM_GIU_Layer3, LM_GIU_Parts, LM_GIU_Playfield, LM_GIU_PrLeaper1, LM_GIU_PrLeaper2, LM_GIU_PrLeaper3, LM_GIU_crate_door, LM_GIU_crate_doorU, LM_GIU_lockdownbar, LM_GIU_sideblades, LM_GIU_sw25, LM_GIU_sw58, LM_GIU_sw68, LM_GIU_sw71, LM_GIU_sw72, LM_GIU_sw73, LM_GIU_sw74, LM_GIU_underPF, LM_L_l11_Playfield, LM_L_l11_underPF, _
' LM_L_l12_Playfield, LM_L_l12_underPF, LM_L_l13_Playfield, LM_L_l13_underPF, LM_L_l14_Playfield, LM_L_l14_underPF, LM_L_l15_Parts, LM_L_l15_Playfield, LM_L_l15_underPF, LM_L_l16_Playfield, LM_L_l16_underPF, LM_L_l17_Playfield, LM_L_l17_underPF, LM_L_l18_Layer2, LM_L_l18_Layer3, LM_L_l18_Parts, LM_L_l21_Playfield, LM_L_l21_PrLeaper2, LM_L_l21_underPF, LM_L_l22_Playfield, LM_L_l22_PrLeaper2, LM_L_l22_underPF, LM_L_l23_Layer1, LM_L_l23_Layer2, LM_L_l23_Layer3, LM_L_l23_Parts, LM_L_l23_Playfield, LM_L_l23_crate_door, LM_L_l23_underPF, LM_L_l24_Playfield, LM_L_l24_PrLeaper2, LM_L_l24_sw65, LM_L_l24_underPF, LM_L_l25_Layer3, LM_L_l25_Playfield, LM_L_l25_sw61, LM_L_l25_sw62, LM_L_l25_sw63, LM_L_l25_underPF, LM_L_l26_Layer3, LM_L_l26_Playfield, LM_L_l26_sw61, LM_L_l26_sw62, LM_L_l26_sw63, LM_L_l26_underPF, LM_L_l27_Layer3, LM_L_l27_Playfield, LM_L_l27_sw61, LM_L_l27_sw62, LM_L_l27_sw63, LM_L_l27_underPF, LM_L_l28_Layer2, LM_L_l28_Playfield, LM_L_l28_underPF, LM_L_l31_Layer1, LM_L_l31_Layer2, LM_L_l31_Layer3, _
' LM_L_l31_Parts, LM_L_l31_sw72, LM_L_l32_Layer1, LM_L_l32_Layer2, LM_L_l32_Layer3, LM_L_l32_Parts, LM_L_l32_sw72, LM_L_l33_Gate007, LM_L_l33_Layer1, LM_L_l33_Layer2, LM_L_l33_Layer3, LM_L_l33_Parts, LM_L_l33_Playfield, LM_L_l34_Gate007, LM_L_l34_Layer1, LM_L_l34_Layer2, LM_L_l34_Layer3, LM_L_l34_Parts, LM_L_l34_Playfield, LM_L_l35_Parts, LM_L_l35_Playfield, LM_L_l35_sideblades, LM_L_l35_underPF, LM_L_l36_Layer2, LM_L_l36_Playfield, LM_L_l36_PrLeaper3, LM_L_l36_sw66, LM_L_l36_underPF, LM_L_l37_Layer2, LM_L_l37_Parts, LM_L_l37_Playfield, LM_L_l37_underPF, LM_L_l38_Playfield, LM_L_l38_underPF, LM_L_l41_Layer3, LM_L_l41_Parts, LM_L_l41_Playfield, LM_L_l41_PrLeaper1, LM_L_l41_sw64, LM_L_l41_underPF, LM_L_l42_Layer3, LM_L_l42_Playfield, LM_L_l42_PrLeaper2, LM_L_l42_underPF, LM_L_l43_Playfield, LM_L_l43_underPF, LM_L_l44_Layer1, LM_L_l44_Layer2, LM_L_l44_Layer3, LM_L_l44_Parts, LM_L_l45_Layer2, LM_L_l45_Parts, LM_L_l45_Playfield, LM_L_l45_underPF, LM_L_l46_Layer3, LM_L_l46_Parts, LM_L_l46_Playfield, _
' LM_L_l46_crate_door, LM_L_l46_crate_doorU, LM_L_l46_underPF, LM_L_l47_Layer2, LM_L_l47_Parts, LM_L_l47_Playfield, LM_L_l47_crate_door, LM_L_l47_underPF, LM_L_l48_Layer2, LM_L_l48_Parts, LM_L_l48_Playfield, LM_L_l48_underPF, LM_L_l51_Layer2, LM_L_l51_Parts, LM_L_l51_Playfield, LM_L_l51_RSling1, LM_L_l51_remk, LM_L_l51_underPF, LM_L_l52_Playfield, LM_L_l52_underPF, LM_L_l53_LBoogie02, LM_L_l53_LSling1, LM_L_l53_Layer3, LM_L_l53_Parts, LM_L_l53_Playfield, LM_L_l53_lemk, LM_L_l53_underPF, LM_L_l54_Layer3, LM_L_l54_Parts, LM_L_l54_Playfield, LM_L_l54_PrLeaper3rod, LM_L_l54_cDoorClose, LM_L_l54_cDoorOpen, LM_L_l54_cadaver, LM_L_l54_sideblades, LM_L_l54_underPF, LM_L_l55_Lflipper, LM_L_l55_LflipperU, LM_L_l55_Playfield, LM_L_l55_Rflipper, LM_L_l55_RflipperU, LM_L_l55_underPF, LM_L_l56_Layer3, LM_L_l56_Parts, LM_L_l56_Playfield, LM_L_l56_underPF, LM_L_l56a_CoffinDiverter, LM_L_l56a_Layer2, LM_L_l56a_Layer3, LM_L_l56a_Parts, LM_L_l57_Layer3, LM_L_l57_Parts, LM_L_l57_Playfield, LM_L_l57_PrLeaper1, LM_L_l57_underPF, _
' LM_L_l58_CoffinDiverter, LM_L_l58_Layer3, LM_L_l58_Parts, LM_L_l58_Playfield, LM_L_l58_underPF, LM_L_l61_Playfield, LM_L_l61_underPF, LM_L_l62_Layer2, LM_L_l62_Layer3, LM_L_l62_Playfield, LM_L_l62_underPF, LM_L_l63_Layer3, LM_L_l63_Playfield, LM_L_l63_underPF, LM_L_l84_Layer1, LM_L_l84_Layer2, LM_L_l84_Layer3, LM_L_l84_Parts, LM_L_l84_Playfield, LM_L_l84_underPF, LM_L_l85_Layer1, LM_L_l85_Layer2, LM_L_l85_Parts, LM_L_l85_Playfield, LM_L_l85_underPF, LM_L_l86_Layer1, LM_L_l86_Layer2, LM_L_l86_Parts, LM_L_l86_Playfield, LM_L_l86_underPF)
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
  NudgeAnim
  SoundCmdListener
End Sub

'The CorTimer interval should be 10. It's sole purpose is to update the Cor (physics) calculations
CorTimer.Interval = 10
Sub CorTimer_Timer(): Cor.Update: End Sub


' Plunger animation timers
Sub TimerPlunger_Timer
  If PinCab_Shooter.TransZ < 90 then
    PinCab_Shooter.TransZ = PinCab_Shooter.TransZ + FrameTime*0.2
    PinCab_CustomShooter.Y = 2453.4 + PinCab_Shooter.TransZ
  End If
End Sub
Sub TimerPlunger2_Timer
  PinCab_Shooter.TransZ = (6.0* Plunger.Position) - 20
  PinCab_CustomShooter.Y = 2453.4 + (6.0* Plunger.Position) - 20
End Sub

'******************************************************
'  ZINI: Table Initialization and Exiting
'******************************************************
Dim PlungerIM
Dim SSBall1, SSBall2, SSBall3, SSBall4, gBOT, mSpider

Sub Table1_Init
  vpminit me
  With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
    .SplashInfoLine = "Scared Stiff (Bally 1996)" & vbNewLine & "VPW"
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
  vpmNudge.TiltSwitch = 14
  vpmNudge.Sensitivity = 4
  vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingshot, TopSlingShot)


  'Spider
  Set mSpider = new cvpmMech
  With mSpider
    .Length = 200
    .Steps = 48
    .mType = vpmMechStepSol + vpmMechCircle + vpmMechLinear + vpmMechFast
    .Sol1 = 39
    .Sol2 = 40
    .Addsw 12, 0, 0
    .Callback = GetRef("UpdateSpider")
    .Start
  End With


  'Main Timer init
  PinMAMETimer.Interval = PinMAMEInterval
  PinMAMETimer.Enabled = 1

  'Trough - Creates a ball in the kicker switch and gives that ball used an individual name.
  Set SSBall1 = sw32.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set SSBall2 = sw33.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set SSBall3 = sw34.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set SSBall4 = sw35.CreateSizedballWithMass(Ballsize/2,Ballmass)
  gBOT = Array(SSBall1,SSBall2,SSBall3,SSBall4)

  'Forces the trough switches to "on" at table boot so the game logic knows there are balls in the trough.
  Controller.Switch(32) = 1
  Controller.Switch(33) = 1
  Controller.Switch(34) = 1
  Controller.Switch(35) = 1

  BM_LSLing1.Visible = 0
  BM_LSLing2.Visible = 0
  BM_RSLing1.Visible = 0
  BM_RSLing2.Visible = 0
  BM_LSLingU1.Visible = 0
  BM_LSLingU2.Visible = 0

  DiverterWall001.IsDropped = 1
  Gate008.collidable = 0

  Dim BP
  For Each BP in BP_PrLeaper1rod : BP.z = 60 : Next
  For Each BP in BP_PrLeaper2rod : BP.z = 60 : Next
  For Each BP in BP_PrLeaper3rod : BP.z = 60 : Next
  For Each BP in BP_sw18 : BP.z = -23.5 : Next
  for each BP in BP_underPF: BP.transz = -5: Next 'little fix

  InitCoffin
  InitBackglass

End Sub


Sub Table_Paused:Controller.Pause = 1:End Sub
Sub Table_unPaused:Controller.Pause = 0:End Sub

' Thalamus - table is named Table1 - not Table
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


Sub InitCoffin
  Dim BP
  For Each BP in BP_cDoorClose
    BP.rotX = -33
    BP.rotY = 0
    BP.rotZ = 0
    BP.ObjRotX = 0
    BP.ObjRotY = 0
    BP.ObjRotZ = -22.5
    BP.visible = 1
  Next
  For Each BP in BP_cDoorOpen
    BP.rotX = -33
    BP.rotY = 0
    BP.rotZ = 0
    BP.ObjRotX = 0
    BP.ObjRotY = 0
    BP.ObjRotZ = -22.5
    BP.visible = 0
  Next
End Sub


'*******************************************
'  ZOPT: User Options
'*******************************************

Dim VolumeDial : VolumeDial = 0.8             ' Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Dim BallRollVolume : BallRollVolume = 0.5     ' Level of ball rolling volume. Value between 0 and 1
Dim RampRollVolume : RampRollVolume = 0.5     ' Level of ramp rolling volume. Value between 0 and 1
Dim DTBackglass : DTBackglass = 1         ' 0=Hide, 1=Show
Dim UsePFWheel : UsePFWheel = 0         ' 0=Dont Use, 1=Use
Dim CustomShooter : CustomShooter = 1       ' 0=Hide, 1=Show
Dim CustomVRTopper: CustomVRTopper = 1      ' 0=Hide, 1=Show
Dim VRPosterLeft : VRPosterLeft = 1
Dim VRPosterRight : VRPosterRight = 2
Dim VRRoomChoice, VRRoom
Dim ColorLUT : ColorLUT = 1           ' Color desaturation LUTs: 1 to 11, where 1 is normal and 11 is black'n'white

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
  Dim v

  ' Desktop DMD
  v = Table1.Option("Desktop DMD", 0, 1, 1, 1, 0, Array("Hide", "Show"))
  ScoreText.visible = v

  ' Desktop Backglass
  DTBackglass = Table1.Option("Desktop Backglass", 0, 1, 1, 1, 0, Array("Hide", "Show"))

  ' Playfield Spider Wheel
  UsePFWheel = Table1.Option("Playfield Spider Wheel", 0, 1, 1, 0, 0, Array("Disable", "Enable"))

  ' Outlane difficulty
  v = Table1.Option("Outlane Difficulty", 0, 2, 1, 1, 0, Array("Easy", "Medium", "Hard"))
  SetOutlanePositions v

  ' Cabinet mode
  v = Table1.Option("Cabinet Mode", 0, 1, 1, 0, 0, Array("Disabled", "Enabled"))
  SetCabinetMode v

  ' Art Blades
  v = Table1.Option("Art Blades", 0, 1, 1, 0, 0, Array("Disabled", "Enabled"))
  SetArtBlades v

  ' Playfield Reflections
  v = Table1.Option("Playfield Reflections", 0, 2, 1, 1, 0, Array("Off", "Clean", "Rough"))
  Select Case v
    Case 0: playfield_mesh.ReflectionProbe = "": BM_Playfield.ReflectionProbe = ""
    Case 1: playfield_mesh.ReflectionProbe = "Playfield Reflections": BM_Playfield.ReflectionProbe = "Playfield Reflections"
    Case 2: playfield_mesh.ReflectionProbe = "Playfield Reflections Rough": BM_Playfield.ReflectionProbe = "Playfield Reflections Rough"
  End Select

    ' Playfield Reflections Scope
  v = Table1.Option("Playfield Reflections Scope", 0, 2, 1, 1, 0, Array("Reflect Ball Only", "Reflect Ball & Parts", "Reflect Ball, Parts & Lighting"))
  ReflectionToggle(v)

  ' Ramp Refractions
  v = Table1.Option("Ramp Refractions", 0, 2, 1, 1, 0, Array("Off", "Clean", "Rough"))
  SetRefractions v

  ' VRRoom
  VRRoomChoice = Table1.Option("VR Room", 1, 2, 1, 1, 0, Array("Minimal", "Cab Only"))

  ' Custom Shooter Knob
  CustomShooter = Table1.Option("VR Custom Shooter Knob", 0, 1, 1, 1, 0, Array("Hide", "Show"))

  ' Custom VR Topper
  CustomVRTopper = Table1.Option("VR Custom Topper", 0, 1, 1, 1, 0, Array("Hide", "Show"))

  'Minimal Room Left Poster
  VRPosterLeft = Table1.Option("VR Minimal Room Left Poster", 0, 4, 1, 3, 0, Array("No Poster", "Option 1", "Option 2", "Option 3", "Option 4"))
  If VRPosterLeft > 0 Then
    PosterLeft.Image = "SS_poster"&VRPosterLeft
  End If

  'Minimal Room Right Posters
  VRPosterRight = Table1.Option("VR Minimal Room Right Poster", 0, 4, 1, 4, 0, Array("No Poster", "Option 1", "Option 2", "Option 3", "Option 4"))
  If VRPosterRight > 0 Then
    PosterRight.Image = "SS_poster"&VRPosterRight
  End If

  SetRoom

    ' Sound volumes
    VolumeDial = Table1.Option("Mech Volume", 0, 1, 0.01, 0.8, 1)
    BallRollVolume = Table1.Option("Ball Roll Volume", 0, 1, 0.01, 0.5, 1)
  RampRollVolume = Table1.Option("Ramp Roll Volume", 0, 1, 0.01, 0.5, 1)

  ' Room brightness
  SetRoomBrightness NightDay/100

  ' Color Saturation
    ColorLUT = Table1.Option("Color Saturation", 0, 11, 1, 0, 0, _
    Array("Normal", "Vibrant", "Desaturated 10%", "Desaturated 20%", "Desaturated 30%", "Desaturated 40%", "Desaturated 50%", _
        "Desaturated 60%", "Desaturated 70%", "Desaturated 80%", "Desaturated 90%", "Black 'n White"))
  if ColorLUT = 0 Then Table1.ColorGradeImage = "colorgradelut256x16_1to1"
  if ColorLUT = 1 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.70vibr50sat100"
  if ColorLUT = 2 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.70vibr50sat90"
  if ColorLUT = 3 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.70vibr50sat80"
  if ColorLUT = 4 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.70vibr50sat70"
  if ColorLUT = 5 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.70vibr50sat60"
  if ColorLUT = 6 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.70vibr50sat50"
  if ColorLUT = 7 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.70vibr50sat40"
  if ColorLUT = 8 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.70vibr50sat30"
  if ColorLUT = 9 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.70vibr50sat20"
  if ColorLUT = 10 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.70vibr50sat10"
  if ColorLUT = 11 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.70vibr50sat00"

    If eventId = 3 And dspTriggered Then dspTriggered = False : DisableStaticPreRendering = False : End If

End Sub


Sub SetOutlanePositions(v)
  Dim BP
  Select Case v
    Case 0
      zCol_LeftPost.x  = 43.0:  zCol_LeftPost.y  = 1448:  For Each BP in BP_lPost: BP.x = zCol_LeftPost.x: BP.y = zCol_LeftPost.y: Next
      zCol_RightPost.x = 828.7: zCol_RightPost.y = 1448:  For Each BP in BP_rPost: BP.x = zCol_RightPost.x: BP.y = zCol_RightPost.y: Next
    Case 1
      zCol_LeftPost.x  = 43.0:  zCol_LeftPost.y  = 1436:  For Each BP in BP_lPost: BP.x = zCol_LeftPost.x: BP.y = zCol_LeftPost.y: Next
      zCol_RightPost.x = 828.7: zCol_RightPost.y = 1436:  For Each BP in BP_rPost: BP.x = zCol_RightPost.x: BP.y = zCol_RightPost.y: Next
    Case 2
      zCol_LeftPost.x  = 43.0:  zCol_LeftPost.y  = 1425:  For Each BP in BP_lPost: BP.x = zCol_LeftPost.x: BP.y = zCol_LeftPost.y: Next
      zCol_RightPost.x = 828.7: zCol_RightPost.y = 1425:  For Each BP in BP_rPost: BP.x = zCol_RightPost.x: BP.y = zCol_RightPost.y: Next
  End Select
End Sub

Sub SetCabinetMode(v)
  Dim visible : visible = 0
  If v = 0 then visible = 1 : End If
  Dim BP
  For Each BP in BP_lockdownbar : BP.visible = visible : Next
End Sub

Sub SetArtBlades(v)
  Dim BP
  For Each BP in BP_sideblades : BP.visible = v : Next
End Sub

Sub SetRefractions(v)
  Select Case v
    Case 0: BM_Layer1.RefractionProbe = "": BM_Layer2.RefractionProbe = "": BM_Layer3.RefractionProbe = ""
    Case 1: BM_Layer1.RefractionProbe = "Refraction Ramp 1": BM_Layer2.RefractionProbe = "Refraction Ramp 2": BM_Layer3.RefractionProbe = "Refraction Ramp 3"
    Case 2: BM_Layer1.RefractionProbe = "Refraction Ramp 1 Rough": BM_Layer2.RefractionProbe = "Refraction Ramp 2 Rough": BM_Layer3.RefractionProbe = "Refraction Ramp 3 Rough"
  End Select
End Sub

Function ReflectionToggle(state)
  Dim CompletReflectObjectsArray: CompletReflectObjectsArray = Array(BP_Parts, BP_Layer2, BP_Layer3, BP_LFlipper, BP_LFlipperU, BP_RFlipper, BP_RFlipperU)
  Dim PartialReflectObjectsArray: PartialReflectObjectsArray = Array(BM_Parts, BM_Layer2, BM_Layer3, BM_LFlipper, BM_LFlipperU, BM_RFlipper, BM_RFlipperU)
  Dim IBP, BP, v1, v2
  Select Case state
    Case 0: v1 = false: v2 = false 'Only reflect the ball
    Case 1: v1 = false: v2 = true  'Only reflect a selection of bakemaps
    Case 2: v1 = true:  v2 = true  'Reflect everything (that matters as defined in ReflObjectsArray)
  End Select

  For Each IBP in CompletReflectObjectsArray
    For Each BP in IBP
      If InStr(BP.name, "_L_") > 0 or InStr(BP.name, "_F_") > 0 Then  'only reflect the GI lighting
        BP.ReflectionEnabled = false
      Else
        BP.ReflectionEnabled = v1
      End If
    Next
  Next

  For Each BP in PartialReflectObjectsArray
    BP.ReflectionEnabled = v2
  Next

  BM_Playfield.ReflectionEnabled = False
End Function


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
Dim RoomBrightnessMtlArray: RoomBrightnessMtlArray = Array("VLM.Bake.Active","VLM.Bake.Solid","spider")

Sub SetRoomBrightness(lvl)
  If lvl > 1 Then lvl = 1
  If lvl < 0 Then lvl = 0

  ' Lighting level
  Dim v: v=(lvl * 245 + 10)/255

  Dim i: For i = 0 to UBound(RoomBrightnessMtlArray)
    ModulateMaterialBaseColor RoomBrightnessMtlArray(i), i, v
  Next

  Dim v2: v2 = lvl * 155 + 100
  BGDark_Layer1.color = RGB(v2,v2,v2)
  BGDark_Layer2a.color = RGB(v2,v2,v2)
  BGDark_Layer2b.color = RGB(v2,v2,v2)
  BGDark_Layer3.color = RGB(v2,v2,v2)

  v2 = lvl * 100 + 20
  PFWeb.color = RGB(v2,v2,v2)
  VRTop_Back.color = RGB(v2,v2,v2)
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
    PinCab_FlipperLeft.transx = 6
    End If
  If Keycode = RightFlipperKey Then
    FlipperActivate RightFlipper, RFPress
    PinCab_FlipperRight.transx = -6
  End If
  If keycode = PlungerKey Then
    Plunger.Pullback
    SoundPlungerPull
    TimerPlunger.Enabled = True
    TimerPlunger2.Enabled = False
    PinCab_Shooter.TransZ = 0
  End If
  If keycode = LeftTiltKey Then Nudge 90, 1 : SoundNudgeLeft: BoogRandNudge     ' Sets the nudge angle and power
  If keycode = RightTiltKey Then Nudge 270, 1 : SoundNudgeRight: BoogRandNudge    ' ^
  If keycode = CenterTiltKey Then Nudge 0, 1 : SoundNudgeCenter: BoogRandNudge    ' ^
  If keycode = StartGameKey Then SoundStartButton
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
  If Keycode = LeftFlipperKey Then
    FlipperDeActivate LeftFlipper, LFPress
    PinCab_FlipperLeft.transx = 0
  End If
  If Keycode = RightFlipperKey Then
    FlipperDeActivate RightFlipper, RFPress
    PinCab_FlipperRight.transx = 0
  End If
  If Keycode = StartGameKey Then Controller.Switch(16) = 0
  If keycode = PlungerKey Then
    Plunger.Fire
    SoundPlungerReleaseBall
    TimerPlunger.Enabled = False
    TimerPlunger2.Enabled = True
  End If
  If vpmKeyUp(keycode) Then Exit Sub
End Sub


'******************************************************
' ZSOL: Solenoids & Flashers
'******************************************************

SolCallback(1) = "AutoPlunge"
SolCallback(2) = "LoopGate"
SolCallback(3) = "scoop_right"
SolCallback(4) = "CoffinPopper"
SolCallback(5) = "CoffinDoor"
SolCallback(6) = "scoop_topleft"
SolCallback(7) = "vpmSolSound SoundFX(""Knocker_1"",DOFKnocker),"
'SolCallback(8) = "CratePostPower"  'Not Required, just for inital power surge
SolCallback(9) = "SolRelease"
SolCallback(10) = "SolLeftSling"
SolCallback(11) = "SolRightSling"
'SolCallBack(12) = "Centre Jet"   'Not Required
'SolCallBack(13) = "Upper Jet"    'Not Required
'SolCallBack(14) = "Lower Jet"    'Not Required
'SolCallBack(15) = "Upper Sling"  'Not Required
SolCallback(16) = "CratePostHold"

SolCallback(33) = "DiverterPower"
SolCallback(34) = "DiverterHold"


'*** Flashers
SolModCallback(17) = "SolFlash17" 'Top Jet Flash (Pop Bumper 1)
SolModCallback(18) = "SolFlash18" 'Mid Jet Flash (Pop Bumper 2)
SolModCallback(19) = "SolFlash19" 'Low Jet Flash (Pop Bumper 3)
SolModCallback(20) = "SolFlash20" 'Bolts Flasher (Playfield)
SolModCallback(21) = "SolFlash21" 'Skull Flasher Left (Blue Skull Flasher)
SolModCallback(22) = "SolFlash22" 'Upper Right Flasher (Orb)
SolModCallback(23) = "SolFlash23" 'Left Ramp Flasher (Inside Boney Skull)
SolModCallback(24) = "SolFlash24" 'Mid Left Flasher (Orb)
SolModCallback(25) = "SolFlash25" 'Skull Flasher Right (White Skull Flasher)
SolModCallback(26) = "SolFlash26" 'Center TV Flasher (Playfield)
SolModCallback(27) = "SolFlash27" 'Upper Left Flasher (Orb)
SolModCallback(28) = "SolFlash28" 'Center Right Flasher (Orb)
SolModCallback(35) = "SolFlash35" 'Lower Left Flasher (Orb)
SolModCallback(36) = "SolFlash36" 'Lower Right Flasher (Orb)

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
  f20a.state = level
End Sub

Sub SolFlash21(level)
  f21.state = level
End Sub

Sub SolFlash22(level)
  f22.state = level
End Sub

Sub SolFlash23(level)
  f23.state = level
  f23a.state = level
End Sub

Sub SolFlash24(level)
  f24.state = level
End Sub

Sub SolFlash25(level)
  f25.state = level
End Sub

Sub SolFlash26(level)
  f26.state = level
End Sub

Sub SolFlash27(level)
  f27.state = level
End Sub

Sub SolFlash28(level)
  f28.state = level
End Sub

Sub SolFlash35(level)
  f35.state = level
End Sub

Sub SolFlash36(level)
  f36.state = level
End Sub

SolCallback(sLRFlipper) = "SolRFlipper" 'Right Flipper
SolCallback(sLLFlipper) = "SolLFlipper" 'Left Flipper


'******************************************************
' ZDRN: Drain, Trough, and Ball Release
'******************************************************

'********************* TROUGH *************************
Sub sw32_Hit   : Controller.Switch(32) = 1 : UpdateTrough : End Sub
Sub sw32_UnHit : Controller.Switch(32) = 0 : UpdateTrough : End Sub
Sub sw33_Hit   : Controller.Switch(33) = 1 : UpdateTrough : End Sub
Sub sw33_UnHit : Controller.Switch(33) = 0 : UpdateTrough : End Sub
Sub sw34_Hit   : Controller.Switch(34) = 1 : UpdateTrough : End Sub
Sub sw34_UnHit : Controller.Switch(34) = 0 : UpdateTrough : End Sub
Sub sw35_Hit   : Controller.Switch(35) = 1 : UpdateTrough : End Sub
Sub sw35_UnHit : Controller.Switch(35) = 0 : UpdateTrough : End Sub

Sub UpdateTrough
  UpdateTroughTimer.Interval = 300
  UpdateTroughTimer.Enabled = 1
End Sub

Sub UpdateTroughTimer_Timer
  If sw32.BallCntOver = 0 Then sw33.kick 57, 10
  If sw33.BallCntOver = 0 Then sw34.kick 57, 10
  If sw34.BallCntOver = 0 Then sw35.kick 57, 10
  Me.Enabled = 0
End Sub


Sub Drain_Hit()
  RandomSoundDrain drain
  UpdateTrough
  vpmTimer.AddTimer 500, "Drain.kick 70, 10'"
End Sub

Sub SolRelease(enabled)
  If enabled Then
    vpmTimer.PulseSw 31
    sw32.kick 60, 9
    RandomSoundBallRelease sw32
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

Gate008.collidable = 0
'Impulse Plunger
Sub AutoPlunge(Enabled)
  If Enabled Then
    Gate008.collidable = 1
    PlungerIM.AutoFire
    SoundSaucerKick 1, sw18
    Gate008.timerenabled = 1
  End If
End Sub

Sub Gate008_Timer
  Gate008.collidable = 0
  Gate008.timerenabled = 0
End Sub

Sub Gate008_Animate
  Dim a : a = Gate008.currentangle
  Dim BP
  For Each BP in BP_Autoplunger : BP.rotx = a : Next
End Sub

Const IMPowerSetting = 45
Const IMTime = 0.6
Set plungerIM = New cvpmImpulseP
With plungerIM
    .InitImpulseP swplunger, IMPowerSetting, IMTime
    .Random 0.3
    .CreateEvents "plungerIM"
End With

'******************************************************
' ZANI: Misc Animations
'******************************************************

'spider
Sub UpdateSpider(aNewPos, aSpeed, aLastPos)
  pSpider.RotY = aNewPos * 360/48 + 200
  pSpider2.ObjRotz = aNewPos * 360/48 + 200
End Sub


Sub LeftFlipper_Animate
  Dim a: a = LeftFlipper.CurrentAngle       ' store flipper angle in a
  Dim max_angle, min_angle, mid_angle       ' min and max angles from the flipper
  max_angle = LeftFlipper.StartAngle                  ' flipper down angle
  min_angle = LeftFlipper.EndAngle
  mid_angle = (max_angle-min_angle)/2 + min_angle ' bake map switch point angle
  FlipperLSh.RotZ = a               ' set flipper shadow angle

  Dim BP :
  For Each BP in BP_LFlipper
    BP.RotZ = a                 ' rotate the maps
    BP.visible = a > mid_angle
  Next
  For Each BP in BP_LFlipperU
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
  For Each BP in BP_RFlipper
    BP.RotZ = a                 ' rotate the maps
    BP.visible = a < mid_angle
  Next
  For Each BP in BP_RFlipperU
    BP.visible = a > mid_angle
    BP.RotZ = a                 ' rotate the maps
  Next
End Sub

Sub LockFlipper_Animate
  Dim BP
  For Each BP in BP_CoffinDiverter : BP.rotz = LockFlipper.CurrentAngle + 180 : Next
End Sub

Sub Gate001_Animate
  Dim BP
  For Each BP in BP_Gate001 : BP.rotx = Gate001.CurrentAngle : Next
End Sub

Sub Gate002_Animate
  Dim BP
  For Each BP in BP_Gate002 : BP.rotx = Gate002.CurrentAngle : Next
End Sub

Sub Gate003_Animate
  Dim BP
  For Each BP in BP_Gate003 : BP.rotx = Gate003.CurrentAngle : Next
End Sub

Sub Gate005_Animate
  Dim BP
  For Each BP in BP_Gate005 : BP.rotx = Gate005.CurrentAngle : Next
End Sub

Dim CrateAnimating : CrateAnimating = False
Sub crate_pin_Hit()
  If Not CrateAnimating Then
    CrateAnimating = True
    Dim BP
    For Each BP in BP_crate_door : BP.rotx = BP.rotx - 2 : Next
    Gate006.TimerEnabled = 1
  End If
End Sub

Sub Gate006_Timer()
  Dim BP
  For Each BP in BP_crate_door : BP.rotx = BP.rotx + 2 : Next
  Gate006.TimerEnabled = 0
  CrateAnimating = False
End Sub

Sub Gate006_Animate
  Dim a: a = Gate006.CurrentAngle / 2     ' store gate angle in a
  if a < 0 then a = 0
' Dim max_angle, min_angle, mid_angle       ' min and max angles from the flipper
' max_angle = Gate006.CloseAngle / 2              ' flipper down angle
' min_angle = Gate006.OpenAngle / 2               ' flipper up angle
' mid_angle = (max_angle-min_angle)/2 + min_angle ' bake map switch point angle

  Dim BP
  For Each BP in BP_crate_door
    BP.Rotx = -a                ' rotate the maps
    BP.visible = 1 'a < mid_angle
  Next
  For Each BP in BP_crate_doorU
    BP.visible = 0 'a > mid_angle
    'BP.Rotx = -a                 ' rotate the maps
  Next
End Sub

Sub GateLoop_Animate
  Dim BP
  For Each BP in BP_GateLoop : BP.rotx = GateLoop.CurrentAngle : Next
End Sub

Sub Gate007_Animate
  Dim BP
  For Each BP in BP_Gate007 : BP.rotx = Gate007.CurrentAngle : Next
End Sub


'************************************************************
' ZSLG: Slingshot Animations
'************************************************************
Dim LStep : LStep = 0 : LeftSlingShot.TimerEnabled = 1
Dim RStep : RStep = 0 : RightSlingShot.TimerEnabled = 1
Dim TStep : TStep = 0 : TopSlingShot.TimerEnabled = 1
Dim RightSlingWasHit : RightSlingWasHit = false
Dim LeftSlingWasHit : LeftSlingWasHit = false

Sub SolRightSling(Enabled)
  If RightSlingWasHit = False Then
    If Enabled Then
      'debug.print "SolRightSling "&Enabled
      RStep = 0
      RightSlingShot.TimerEnabled = 1
      RightSlingShot.TimerInterval = 17
      RandomSoundSlingshotRight BM_remk
      BoogRSlingNudge
    End If
  End If
End Sub

Sub RightSlingShot_Slingshot
  RightSlingWasHit = True
  RS.VelocityCorrect(Activeball)
  vpmTimer.PulseSw(52)            'Sling Switch Number
  RStep = 0
  RightSlingShot.TimerEnabled = 1
  RightSlingShot.TimerInterval = 17
  RandomSoundSlingshotRight BM_remk
  BoogRSlingNudge
End Sub

Sub RightSlingShot_Timer
  Dim BP
  Dim x1, x2, y: x1 = False:x2 = True:y = 20
    Select Case RStep
        Case 2:x1 = True:x2 = False:y = 10
        Case 3:x1 = False:x2 = False:y = 0: RightSlingWasHit = False: RightSlingShot.TimerEnabled = 0
    End Select

  For Each BP in BP_RSling2 : BP.Visible = x1: Next
  For Each BP in BP_RSling1 : BP.Visible = x2: Next
  For Each BP in BP_remk : BP.transx = y: Next
  For Each BP in BP_RBoggie01 : BP.transx = y: Next
  For Each BP in BP_RBoggie02 : BP.transx = y: Next

  RStep = RStep + 1
End Sub



Sub SolLeftSling(Enabled)
  If LeftSlingWasHit = False Then
    If Enabled Then
      'debug.print "SolLeftSling "&Enabled
      LStep = 0
      LeftSlingShot.TimerEnabled = 1
      LeftSlingShot.TimerInterval = 17
      RandomSoundSlingshotLeft BM_lemk
      BoogLSlingNudge
    End If
  End If
End Sub

Sub LeftSlingShot_Slingshot
  LeftSlingWasHit = True
  LS.VelocityCorrect(Activeball)
  vpmTimer.PulseSw(51)            'Sling Switch Number
  LStep = 0
  LeftSlingShot.TimerEnabled = 1
  LeftSlingShot.TimerInterval = 17
  RandomSoundSlingshotLeft BM_lemk
  BoogLSlingNudge
End Sub

Sub LeftSlingShot_Timer
  Dim BP
  Dim x1, x2, y: x1 = False:x2 = True:y = 20
    Select Case LStep
        Case 2:x1 = True:x2 = False:y = 10
        Case 3:x1 = False:x2 = False:y = 0:LeftSlingWasHit = False: LeftSlingShot.TimerEnabled = 0
    End Select

  For Each BP in BP_LSling2 : BP.Visible = x1: Next
  For Each BP in BP_LSling1 : BP.Visible = x2: Next
  For Each BP in BP_lemk : BP.transx = y: Next
  For Each BP in BP_LBoogie01 : BP.transx = y: Next
  For Each BP in BP_LBoogie02 : BP.transx = y: Next

    LStep = LStep + 1
End Sub

Sub TopSlingShot_Slingshot
  TS.VelocityCorrect(Activeball)
  vpmTimer.PulseSw(56)            'Sling Switch Number
  TStep = 0       'Initialize Step to 0
  TopSlingShot.TimerEnabled = 1
  TopSlingShot.TimerInterval = 10
  RandomSoundSlingshotLeft BM_LSlingArmU
End Sub

Sub TopSlingShot_Timer
  Dim BP
  Dim x1, x2, y: x1 = False:x2 = True:y = 20
    Select Case TStep
        Case 2:x1 = True:x2 = False:y = 10
        Case 3:x1 = False:x2 = False:y = 0:TopSlingShot.TimerEnabled = 0
    End Select

  For Each BP in BP_LSlingU2 : BP.Visible = x1: Next
  For Each BP in BP_LSlingU1 : BP.Visible = x2: Next
  For Each BP in BP_LSlingArmU : BP.transx = y: Next

  TStep = TStep + 1
End Sub


'************************************************************
' ZBOO:  Boogie Monsters  (aka little fuckers)
'************************************************************

' Wobble Script to animate objects when the table is nudged. The below shows how to animate a single axis. Additional scripting is needed to animate for another axis.
'
'  You'll need variables for each object to:
'   1 - track time interval (objtime)
'   2 - two to track velocity (objvel, objvel2)
'   3 - track the displacement (objdisp)
'   4 - define the maximum movement of the object (objmaxmovement)
'   5 - define the acceleration of the object, or how fast it moves (objacc)
'   6 - define the how fast the object movement decays (objdecay)
'
'  NudgeAnim should be placed in a frame timer (interval = -1)
'   - NudgeAnim should be used to read any analog nudges, and if so, apply an impulse to the object
'   - NudgeAnim should also call the animation sub for the object
'
'  A sub should be created to apply an impulse to the object when the table is nudged. It will be placed in the NudgeAnim sub for analog nudges and the nudge key presses for digital nudge
'
' Sub ObjNudge(namp, ndir)
'   CalcVelTime namp - rnd * namp * 0.1 , ndir, objvel, objtime, objvel2, objmaxmovement, objacc  'adds some randomness to the impulse
' End Sub
'
'   Where:
'  - namp is the amplitude of the impulse
'  - ndir is the direction of the impulse
'
'  A sub should be created to update the displacement of the object. It will be placed in the NudgeAnim sub.
'
' Sub ObjUpdate()
'   CalcDisplacement objvel, objtime, objdisp, objvel2, objmaxmovement, objacc, objdecay
'   obj.transy = - objdisp
' End Sub


Dim BoogTimeY_L1: BoogTimeY_L1=0
Dim BoogTimeY_L2: BoogTimeY_L2=0
Dim BoogTimeY_R1: BoogTimeY_R1=0
Dim BoogTimeY_R2: BoogTimeY_R2=0

Dim BoogVelY_L1, BoogVelY2_L1: BoogVelY_L1 = 0: BoogVelY2_L1 = 0
Dim BoogVelY_L2, BoogVelY2_L2: BoogVelY_L2 = 0: BoogVelY2_L2 = 0
Dim BoogVelY_R1, BoogVelY2_R1: BoogVelY_R1 = 0: BoogVelY2_R1 = 0
Dim BoogVelY_R2, BoogVelY2_R2: BoogVelY_R2 = 0: BoogVelY2_R2 = 0

Dim BoogDispY_L1: BoogDispY_L1 = 0
Dim BoogDispY_L2: BoogDispY_L2 = 0
Dim BoogDispY_R1: BoogDispY_R1 = 0
Dim BoogDispY_R2: BoogDispY_R2 = 0

Const BoogMaxYMovement = 90
Const BoogDecay = 0.94
Const BoogAcc = 35000
Const BoogNudgeGain = 3

Sub NudgeAnim() 'Call from FrameTimer
    Dim X, Y
    NudgeSensorStatus X, Y
  'debug.print "X = "&X&"  Y = "&Y

  'Add impulse to any objects that wobble around the y-axis
    If ABS(X) > 0.05 Then
    BoogNudge Abs(x)*BoogNudgeGain, x/Abs(x), BoogVelY_L1, BoogTimeY_L1, BoogVelY2_L1, BoogMaxYMovement, BoogAcc
    BoogNudge Abs(x)*BoogNudgeGain, x/Abs(x), BoogVelY_L2, BoogTimeY_L2, BoogVelY2_L2, BoogMaxYMovement, BoogAcc
    BoogNudge Abs(x)*BoogNudgeGain, x/Abs(x), BoogVelY_R1, BoogTimeY_R1, BoogVelY2_R1, BoogMaxYMovement, BoogAcc
    BoogNudge Abs(x)*BoogNudgeGain, x/Abs(x), BoogVelY_R2, BoogTimeY_R2, BoogVelY2_R2, BoogMaxYMovement, BoogAcc
    End If

  'Add impulse to any objects that wobble around the x-axis
    If ABS(Y) > 0.05 Then
    BoogNudge Abs(y)*BoogNudgeGain, y/Abs(y), BoogVelY_L1, BoogTimeY_L1, BoogVelY2_L1, BoogMaxYMovement, BoogAcc
    BoogNudge Abs(y)*BoogNudgeGain, y/Abs(y), BoogVelY_L2, BoogTimeY_L2, BoogVelY2_L2, BoogMaxYMovement, BoogAcc
    BoogNudge Abs(y)*BoogNudgeGain, y/Abs(y), BoogVelY_R1, BoogTimeY_R1, BoogVelY2_R1, BoogMaxYMovement, BoogAcc
    BoogNudge Abs(y)*BoogNudgeGain, y/Abs(y), BoogVelY_R2, BoogTimeY_R2, BoogVelY2_R2, BoogMaxYMovement, BoogAcc
    End If

    BoogUpdate
End Sub

Sub BoogNudge(namp, ndir, Vel, nTime, Vel2, MaxMovement, Acc)
  CalcVelTime namp - rnd * namp * 0.1 , ndir, Vel, nTime, Vel2, MaxMovement, Acc  'adds some randomness to the impulse
End Sub

Sub BoogUpdate()
  Dim BP
  CalcDisplacement BoogVelY_L1, BoogTimeY_L1, BoogDispY_L1, BoogVelY2_L1, BoogMaxYMovement, BoogAcc, BoogDecay
  CalcDisplacement BoogVelY_L2, BoogTimeY_L2, BoogDispY_L2, BoogVelY2_L2, BoogMaxYMovement, BoogAcc, BoogDecay
  CalcDisplacement BoogVelY_R1, BoogTimeY_R1, BoogDispY_R1, BoogVelY2_R1, BoogMaxYMovement, BoogAcc, BoogDecay
  CalcDisplacement BoogVelY_R2, BoogTimeY_R2, BoogDispY_R2, BoogVelY2_R2, BoogMaxYMovement, BoogAcc, BoogDecay

  BM_LBoogie01.roty = BoogDispY_L1
  BM_LBoogie02.roty = BoogDispY_L2
  BM_RBoggie01.roty = BoogDispY_R1
  BM_RBoggie02.roty = BoogDispY_R2

  'debug.print "BoogDispX = "&BoogDispX&"  BoogDispY = "&BoogDispY
  For each BP in BP_LBoogie01: BP.roty = BM_LBoogie01.roty: Next
  For each BP in BP_LBoogie02: BP.roty = BM_LBoogie02.roty: Next
  For each BP in BP_RBoggie01: BP.roty = BM_RBoggie01.roty: Next
  For each BP in BP_RBoggie02: BP.roty = BM_RBoggie02.roty: Next
End Sub

Sub BoogRandNudge
  BoogNudge 0.5-(0.3*Rnd),-1, BoogVelY_L1, BoogTimeY_L1, BoogVelY2_L1, BoogMaxYMovement, BoogAcc
  BoogNudge 0.5-(0.3*Rnd),-1, BoogVelY_L2, BoogTimeY_L2, BoogVelY2_L2, BoogMaxYMovement, BoogAcc
  BoogNudge 0.5-(0.3*Rnd),-1, BoogVelY_R1, BoogTimeY_R1, BoogVelY2_R1, BoogMaxYMovement, BoogAcc
  BoogNudge 0.5-(0.3*Rnd),-1, BoogVelY_R2, BoogTimeY_R2, BoogVelY2_R2, BoogMaxYMovement, BoogAcc
End Sub

Sub BoogLSlingNudge
  BoogNudge 0.7-(0.2*Rnd),-1, BoogVelY_L1, BoogTimeY_L1, BoogVelY2_L1, BoogMaxYMovement, BoogAcc
  BoogNudge 0.7-(0.2*Rnd),-1, BoogVelY_L2, BoogTimeY_L2, BoogVelY2_L2, BoogMaxYMovement, BoogAcc
End Sub

Sub BoogRSlingNudge
  BoogNudge 0.7-(0.2*Rnd),-1, BoogVelY_R1, BoogTimeY_R1, BoogVelY2_R1, BoogMaxYMovement, BoogAcc
  BoogNudge 0.7-(0.2*Rnd),-1, BoogVelY_R2, BoogTimeY_R2, BoogVelY2_R2, BoogMaxYMovement, BoogAcc
End Sub

' Wobble Supporting Functions
'*****************************

' Calculate the change in velocity of the object due to the impulse from the nudge
Sub CalcVelTime(simpulse, sidir, svel, stime, nvel, smaxmovement, sacc)
  if simpulse > 1 then simpulse = 1

  if svel = 0 Then
    svel = simpulse * sidir * Vel0(sMaxMovement, sacc)
    stime = GameTime
    nvel = 0
  else
    nvel = simpulse * sidir * Vel0(sMaxMovement, sacc)
  end If
End Sub

Function Vel0(displacement, acceleration)
  Vel0 = SQR(acceleration* ABS(displacement) * 2)
End Function

' Calculate the current displacement of the object
Sub CalcDisplacement(svel, stime, sdisp, nvel, smaxmovement, sacc, sdecay)
  dim velM, accM , stimef, stimec

  stimec = Gametime - stime
  stimef = TimeF(MaxDisplacement(svel, sacc), sacc)

  If stimec > stimef Then
    stimec = stimec - stimef
    stime = Gametime - stimec

    If nvel <> 0 Then
      if abs(nvel) > abs(svel) then
        svel = -sgn(svel)*abs(nvel)
      Else
        svel = -svel * sdecay
      End If

      nvel = 0
    Else
      svel = -svel * sdecay
    End If

    stimef = TimeF(MaxDisplacement(svel, sacc), sacc)

    If stimec > stimef Then
      svel = 0
      sdisp  = 0
    End If
  End If

  velM = svel * stimec / 1000
  accM = (sacc * (stimec / 1000)^2) / 2

  If svel < 0 Then
    sdisp = velM + accM
  ElseIf svel > 0 Then
    sdisp = velM - accM
  End If
End Sub

Function MaxDisplacement(velocity, acceleration)
  MaxDisplacement = Velocity^2/(2*acceleration)
End Function

Function TimeF(displacement, acceleration)
  TimeF = 1000*2*SQR(Abs(displacement)*acceleration*2)/acceleration
End Function





'******************************************************
' ZSSC: SLINGSHOT CORRECTION FUNCTIONS by apophis
'******************************************************
' To add these slingshot corrections:
'  - On the table, add the endpoint primitives that define the two ends of the Slingshot
'  - Initialize the SlingshotCorrection objects in InitSlingCorrection
'  - Call the .VelocityCorrect methods from the respective _Slingshot event sub

Dim LS: Set LS = New SlingshotCorrection
Dim RS: Set RS = New SlingshotCorrection
Dim TS: Set TS = New SlingshotCorrection

InitSlingCorrection

Sub InitSlingCorrection
  LS.Object = LeftSlingshot
  LS.EndPoint1 = EndPoint1LS
  LS.EndPoint2 = EndPoint2LS

  RS.Object = RightSlingshot
  RS.EndPoint1 = EndPoint1RS
  RS.EndPoint2 = EndPoint2RS

  TS.Object = TopSlingshot
  TS.EndPoint1 = EndPoint1TS
  TS.EndPoint2 = EndPoint2TS

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
  a = Array(LS, RS, TS)
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


'************************************************************
' ZBMP: Bumpers
'************************************************************


'************************* Bumpers **************************
Sub Bumper1_Hit()
  RandomSoundBumperTop Bumper1
  vpmTimer.PulseSw 53

  Dim BP
  For Each BP in BP_BSocket1 : BP.roty=skirtAY(me,Activeball) : Next
  For Each BP in BP_BSocket1 : BP.rotx=skirtAX(me,Activeball) : Next

  me.timerinterval = 150
  me.timerenabled=1
End Sub

sub Bumper1_timer
  Dim BP
  For Each BP in BP_BSocket1 : BP.roty=0 : Next
  For Each BP in BP_BSocket1 : BP.rotx=0 : Next
  me.timerenabled=0
end sub

Sub Bumper2_Hit()
  RandomSoundBumperMiddle Bumper2
  vpmTimer.PulseSw 54

  Dim BP
  For Each BP in BP_BSocket2 : BP.roty=skirtAY(me,Activeball) : Next
  For Each BP in BP_BSocket2 : BP.rotx=skirtAX(me,Activeball) : Next

  me.timerinterval = 150
  me.timerenabled=1
End Sub

sub Bumper2_timer
  Dim BP
  For Each BP in BP_BSocket2 : BP.roty=0 : Next
  For Each BP in BP_BSocket2 : BP.rotx=0 : Next
  me.timerenabled=0
end sub

Sub Bumper3_Hit()
  RandomSoundBumperBottom Bumper3
  vpmTimer.PulseSw 55

  Dim BP
  For Each BP in BP_BSocket3 : BP.roty=skirtAY(me,Activeball) : Next
  For Each BP in BP_BSocket3 : BP.rotx=skirtAX(me,Activeball) : Next

  me.timerinterval = 150
  me.timerenabled=1
End Sub

sub Bumper3_timer
  Dim BP
  For Each BP in BP_BSocket3 : BP.roty=0 : Next
  For Each BP in BP_BSocket3 : BP.rotx=0 : Next
  me.timerenabled=0
end sub

Sub Bumper1_Animate
  Dim z: z = Bumper1.CurrentRingOffset

  Dim BP
  For Each BP in BP_BRing1 : BP.transz = z : Next
  For Each BP in BP_BFlasher1 : BP.transz = z : Next
End Sub

Sub Bumper2_Animate
  Dim z: z = Bumper2.CurrentRingOffset

  Dim BP
  For Each BP in BP_BRing2 : BP.transz = z : Next
  For Each BP in BP_BFlasher2 : BP.transz = z : Next
End Sub

Sub Bumper3_Animate
  Dim z: z = Bumper3.CurrentRingOffset

  Dim BP
  For Each BP in BP_BRing3 : BP.transz = z : Next
  For Each BP in BP_BFlasher3 : BP.transz = z : Next
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


'************************************************************
' ZSWI: Switches
'************************************************************


'************************ Rollovers *************************
Sub sw16_Hit(): Controller.Switch(16) = 1: AnimateWire BP_sw16, 1 : End Sub   'Rollover
Sub sw16_UnHit: Controller.Switch(16) = 0: AnimateWire BP_sw16, 0 : End Sub
Sub sw17_Hit(): Controller.Switch(17) = 1: AnimateWire BP_sw17, 1 : rightInlaneSpeedLimit : End Sub   'Rollover
Sub sw17_UnHit: Controller.Switch(17) = 0: AnimateWire BP_sw17, 0 : End Sub
Sub sw18_Hit(): Controller.Switch(18) = 1: AnimateWire BP_sw18, 1 : End Sub   'Rollover
Sub sw18_UnHit: Controller.Switch(18) = 0: AnimateWire BP_sw18, 0 : End Sub
Sub sw25_Hit(): Controller.Switch(25) = 1: AnimateWire BP_sw25, 1 : End Sub   'Rollover
Sub sw25_UnHit: Controller.Switch(25) = 0: AnimateWire BP_sw25, 0 : End Sub
Sub sw26_Hit(): Controller.Switch(26) = 1: AnimateWire BP_sw26, 1 : leftInlaneSpeedLimit : End Sub    'Rollover
Sub sw26_UnHit: Controller.Switch(26) = 0: AnimateWire BP_sw26, 0 : End Sub
Sub sw27_Hit(): Controller.Switch(27) = 1: AnimateWire BP_sw27, 1 : End Sub   'Rollover
Sub sw27_UnHit: Controller.Switch(27) = 0: AnimateWire BP_sw27, 0 : End Sub
Sub sw58_Hit(): Controller.Switch(58) = 1: AnimateWire BP_sw58, 1 : End Sub   'Rollover
Sub sw58_UnHit: Controller.Switch(58) = 0: AnimateWire BP_sw58, 0 : End Sub
Sub sw68_Hit(): Controller.Switch(68) = 1: AnimateWire BP_sw68, 1 : End Sub   'Rollover
Sub sw68_UnHit: Controller.Switch(68) = 0: AnimateWire BP_sw68, 0 : End Sub
Sub sw71_Hit(): Controller.Switch(71) = 1: AnimateWire BP_sw71, 1 : End Sub   'Rollover
Sub sw71_UnHit: Controller.Switch(71) = 0: AnimateWire BP_sw71, 0 : End Sub
Sub sw72_Hit(): Controller.Switch(72) = 1: AnimateWire BP_sw72, 1 : End Sub   'Rollover
Sub sw72_UnHit: Controller.Switch(72) = 0: AnimateWire BP_sw72, 0 : End Sub
Sub sw73_Hit(): Controller.Switch(73) = 1: AnimateWire BP_sw73, 1 : End Sub   'Rollover
Sub sw73_UnHit: Controller.Switch(73) = 0: AnimateWire BP_sw73, 0 : End Sub
Sub sw74_Hit(): Controller.Switch(74) = 1: AnimateWire BP_sw74, 1 : End Sub   'Rollover
Sub sw74_UnHit: Controller.Switch(74) = 0: AnimateWire BP_sw74, 0 : End Sub

Sub AnimateWire(group, action) ' Action = 1 - to drop, 0 to raise)
  Dim BP
  If action = 1 Then
    For Each BP in group : BP.transz = -13 : Next
  Else
    For Each BP in group : BP.transz = 0 : Next
  End If
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


'31 - Trough Eject
'32 - Trough 1
'33 - Trough 2
'34 - Trough 3
'35 - Trough 4
'36 - Right VUK Switch
'37 - Left VUK Switch

'38 - Subway Crate Opto
Sub sw38_Hit(): Controller.Switch(38) = 1: End Sub    'Subway Crate Opto
Sub sw38_UnHit: Controller.Switch(38) = 0: End Sub

'Coffin Trough Left
'Sub sw41_Hit(): Controller.Switch(41) = 1: End Sub   'Coffin Trough Left
'Sub sw41_UnHit: Controller.Switch(41) = 0: End Sub
Sub sw42_Hit(): Controller.Switch(42) = 1: End Sub    'Coffin Trough Centre
Sub sw42_UnHit: Controller.Switch(42) = 0: End Sub
Sub sw43_Hit(): Controller.Switch(43) = 1: End Sub    'Coffin Trough Right
Sub sw43_UnHit: Controller.Switch(43) = 0: End Sub

'Ramp Triggers
Sub sw44_Hit(): Controller.Switch(44) = 1: End Sub    'Ramp Left Enter
Sub sw44_UnHit: Controller.Switch(44) = 0: End Sub
Sub sw45_Hit(): Controller.Switch(45) = 1: End Sub    'Ramp Right Enter
Sub sw45_UnHit: Controller.Switch(45) = 0: End Sub
Sub sw46_Hit(): Controller.Switch(46) = 1: End Sub    'Ramp Left Made
Sub sw46_UnHit: Controller.Switch(46) = 0: End Sub
Sub sw47_Hit(): Controller.Switch(47) = 1: End Sub    'Ramp Right Made
Sub sw47_UnHit: Controller.Switch(47) = 0: End Sub

'48 - Coffin Entrance
Sub sw48_Hit(): Controller.Switch(48) = 1: ScoopHit: End Sub
Sub sw48_UnHit: Controller.Switch(48) = 0: End Sub

'51 sling left
'52 sling right
'53 Upper Pop
'54 Centre Pop
'55 Lower Pop
'56 sling upper

'57 - Crate Magnetic Sensor
Sub sw57_Hit(): Controller.Switch(57) = 1: End Sub
Sub sw57_UnHit: Controller.Switch(57) = 0: End Sub

Sub sw67_Hit(): Controller.Switch(67) = 1: End Sub    'Top Right Rubber Switch
Sub sw67_UnHit: Controller.Switch(67) = 0: End Sub


'********************* Standup Targets **********************
Sub sw28_hit: STHit 28: End Sub

Sub sw61_hit: STHit 61: End Sub

Sub sw62_hit: STHit 62: End Sub

Sub sw63_hit: STHit 63: End Sub



'************************************************************
' ZFRG: Frogs
'************************************************************

dim FrogDir1, frogdir2, frogdir3, Frog1Vel, Frog2Vel, Frog3Vel
frogdir1 = 1 :frogdir2 = 1 :frogdir3 = 1

Const FrogSpeed = 1.3
Const FrogSpeedDown = 0.6
Const FrogRotSpeed = 0.01

Sub sw64_Hit        'Left leaper
  STHit 64
  Frog1Vel = BallSpeed(activeball) * 10
  RandomSoundMetal3
  me.TimerEnabled = 1
End Sub

Sub sw65_Hit        'Center Leaper
  STHit 65
  Frog2Vel = BallSpeed(activeball) * 10
  RandomSoundMetal3
  me.TimerEnabled = 1
End Sub

Sub sw66_Hit        'Right Leaper
  STHit 66
  Frog3Vel = BallSpeed(activeball) * 10
  RandomSoundMetal3
  me.TimerEnabled = 1
End Sub


Dim Dir1, chdir1, updown1, slowmo
Dir1 = 1
updown1 = 1
ChDir1 = 0
Sub sw64_Timer()
  Dim maxHeight
  If Frog1Vel > 100 Then
    maxHeight = 200
  Else
    maxHeight = Frog1Vel * 2
  End If

  dim BP
  If updown1 = -1 AND ChDir1 = 0 Then ChDir1 = 1
  If ChDir1 = 1 Then
    If BM_PrLeaper1.Z >= maxHeight Then RandomSoundMetal : ChDir1 = 2
    If BM_PrLeaper1.Z >= 180 AND BM_PrLeaper1.Z < maxHeight Then RandomSoundMetal : ChDir1 = 2
    If BM_PrLeaper1.Z >= 160 AND BM_PrLeaper1.Z < 180 Then RandomSoundMetal : ChDir1 = 2
  End If
  For Each BP in BP_PrLeaper1 : BP.Z = dSin(dir1) * 500 + 60 : Next
  For Each BP in BP_PrLeaper1rod : BP.Z = dSin(dir1) * 500 + 60 : Next

  if ActiveBall.velX > 0 Then
    frogdir1 = -1
  Else
    frogdir1 = 1
  end if

  For Each BP in BP_PrLeaper1 : BP.RotZ = BP.RotZ + (FrogRotSpeed * Frog1Vel * frogdir1) * (FrameTime/7) : Next 'simple rotation
  If BM_PrLeaper1.Z >= maxHeight Then updown1 = -1
  If dir1 >= 80 Then updown1 = -1
  If updown1 = 1 Then  ' Up
    dir1 = dir1 + dCos(dir1) * FrogSpeed * (FrameTime/7)
  Else  ' Down (updown1 = -1)
    dir1 = dir1 + dCos(dir1) * updown1 * FrogSpeedDown * (FrameTime/7)
  End If
  If BM_PrLeaper1.Z <= 60 Then
    For Each BP in BP_PrLeaper1 : BP.Z = 60 : Next
    For Each BP in BP_PrLeaper1rod : BP.Z = 60 : Next
    Me.TimerEnabled = 0
    Dir1 = 1
    ChDir1 = 0
    updown1 = 1
  End If
End Sub

Dim Dir2, chdir2, updown2
Dir2 = 1
updown2 = 1
ChDir2 = 0
Sub sw65_Timer()
  Dim maxHeight
  If Frog2Vel > 100 Then
    maxHeight = 200
  Else
    maxHeight = Frog2Vel * 2
  End If

  dim BP
  If updown2 = -1 AND ChDir2 = 0 Then ChDir2 = 1
  If ChDir2 = 1 Then
    If BM_PrLeaper2.Z >= maxHeight Then RandomSoundMetal : ChDir2 = 2
    If BM_PrLeaper2.Z >= 180 AND BM_PrLeaper2.Z < maxHeight Then RandomSoundMetal : ChDir2 = 2
    If BM_PrLeaper2.Z >= 160 AND BM_PrLeaper2.Z < 180 Then RandomSoundMetal : ChDir2 = 2
  End If
  For Each BP in BP_PrLeaper2 : BP.Z = dSin(dir2) * 500 + 60 : Next
  For Each BP in BP_PrLeaper2rod : BP.Z = dSin(dir2) * 500 + 60 : Next

  if ActiveBall.velX > 0 Then
    frogdir2 = -1
  Else
    frogdir2 = 1
  end if

  For Each BP in BP_PrLeaper2 : BP.RotZ = BP.RotZ + (FrogRotSpeed * Frog2Vel * frogdir2) * (FrameTime/7) : Next 'simple rotation
  If BM_PrLeaper2.Z >= maxHeight Then updown2 = -1
  If dir2 >= 80 Then updown2 = -1
  If updown2 = 1 Then  ' Up
    dir2 = dir2 + dCos(dir2) * FrogSpeed * (FrameTime/7)
  Else  ' Down (updown2 = -1)
    dir2 = dir2 + dCos(dir3) * updown2 * FrogSpeedDown * (FrameTime/7)
  End If
  If BM_PrLeaper2.Z <= 60 Then
    For Each BP in BP_PrLeaper2 : BP.Z = 60 : Next
    For Each BP in BP_PrLeaper2rod : BP.Z = 60 : Next
    Me.TimerEnabled = 0
    Dir2 = 1
    ChDir2 = 0
    updown2 = 1
  End If
End Sub

Dim Dir3, chdir3, updown3
Dir3 = 1
updown3 = 1
ChDir3 = 0
Sub sw66_Timer()
  Dim maxHeight
  If Frog3Vel > 100 Then
    maxHeight = 200
  Else
    maxHeight = Frog3Vel * 2
  End If

  dim BP
  If updown3 = -1 AND ChDir3 = 0 Then ChDir3 = 1
  If ChDir3 = 1 Then
    If BM_PrLeaper3.Z >= maxHeight Then RandomSoundMetal : ChDir3 = 2
    If BM_PrLeaper3.Z >= 180 AND BM_PrLeaper3.Z < maxHeight Then RandomSoundMetal : ChDir3 = 2
    If BM_PrLeaper3.Z >= 160 AND BM_PrLeaper3.Z < 180 Then RandomSoundMetal : ChDir3 = 2
  End If
  For Each BP in BP_PrLeaper3 : BP.Z = dSin(dir3) * 500 + 60 : Next
  For Each BP in BP_PrLeaper3rod : BP.Z = dSin(dir3) * 500 + 60 : Next

  if ActiveBall.velX > 0 Then
    frogdir3 = -1
  Else
    frogdir3 = 1
  end if

  For Each BP in BP_PrLeaper3 : BP.RotZ = BP.RotZ + (FrogRotSpeed * Frog3Vel * frogdir3) * (FrameTime/7) : Next 'simple rotation
  If BM_PrLeaper3.Z >= maxHeight Then updown3 = -1
  If dir3 >= 80 Then updown3 = -1
  If updown3 = 1 Then  ' Up
    dir3 = dir3 + dCos(dir3) * FrogSpeed * (FrameTime/7)
  Else  ' Down (updown3 = -1)
    dir3 = dir3 + dCos(dir3) * updown3 * FrogSpeedDown * (FrameTime/7)
  End If
  If BM_PrLeaper3.Z <= 60 Then
    For Each BP in BP_PrLeaper3 : BP.Z = 60 : Next
    For Each BP in BP_PrLeaper3rod : BP.Z = 60 : Next
    Me.TimerEnabled = 0
    Dir3 = 1
    ChDir3 = 0
    updown3 = 1
  End If
End Sub

'************************************************************
' ZVUK: VUKs and Kickers
'************************************************************

Dim KickerBall36, KickerBall36h, KickerBall37, KickerBall41   'Each VUK needs its own "kickerball"

Sub KickBall(kball, kangle, kvel, kvelz, kzlift)  'Defines how KickBall works
  dim rangle
  rangle = PI * (kangle - 90) / 180

  kball.z = kball.z + kzlift
  kball.velz = kvelz
  kball.velx = cos(rangle)*kvel
  kball.vely = sin(rangle)*kvel
End Sub

'Coffin Scoop
Sub sw41_Hit                    'Switch associated with Kicker
    set KickerBall41 = activeball
    Controller.Switch(41) = 1
    SoundSaucerLock
End Sub

Sub CoffinPopper(Enable)              'Solonoid name associated with kicker.
    If Enable then
    If Controller.Switch(41) <> 0 Then
      KickBall KickerBall41, 168, 1, 30, 100
      SoundSaucerKick 1, sw41
      Controller.Switch(41) = 0
    End If
  End If
End Sub

'Left Scoop
Sub sw37_Hit                    'Switch associated with Kicker
    set KickerBall37 = activeball
    Controller.Switch(37) = 1
    ScoopHit
End Sub

Sub scoop_topleft(Enable)             'Solonoid name associated with kicker.
    If Enable then
    If Controller.Switch(37) <> 0 Then
      KickBall KickerBall37, 0, 3, RndNum(35,45), 50
      SoundSaucerKick 1, sw37
      Controller.Switch(37) = 0
    End If
  End If
End Sub

'Right Scoop
dim sw36sfx: sw36sfx = True
Sub sw36_Hit
    set KickerBall36 = activeball
    Controller.Switch(36) = 1
  If sw36sfx = True Then
    sw36sfx = False
    vpmTimer.AddTimer 500, "sw36sfx = True'"
    ScoopHit
  End If
End Sub

Sub scoop_right(Enable)
    If Enable then
    If Controller.Switch(36) <> 0 Then
      Subway_UpVukTrap.collidable = 0
      KickBall KickerBall36, 0, 0, 30, 105
      SoundSaucerKick 1, sw36
      Controller.Switch(36) = 0
      sw36.TimerEnabled = 1
      Dim BP
      For Each BP in BP_rampFlap : BP.rotx = BP.rotx - 80 : Next
    End If
  End If
End Sub

Sub sw36_Timer
  Dim BP
  If KickerBall36.Z >= 109 Then
    If UsePFWheel = 1 Then HidePFWheel
    Subway_UpVukTrap.collidable = 1
    For Each BP in BP_rampFlap : BP.rotx = BP.rotx + 80 : Next
    sw36.TimerEnabled = 0
  ElseIf KickerBall36.Y >= 1200 Then
    If UsePFWheel = 1 Then HidePFWheel
    Subway_UpVukTrap.collidable = 1
    For Each BP in BP_rampFlap : BP.rotx = BP.rotx + 80 : Next
    sw36.TimerEnabled = 0
  End If
end sub

'******************* Kicker / Trapdoor **********************

Sub LoopGate(Enabled)
  GateLoop.Open = Enabled
End Sub

Sub CratePostHold(Enabled)
  Crate_Pin.Collidable = Not Enabled
End Sub

Dim LockPower, LockHold

Sub DiverterPower(enabled)  'Lock Diverter
  LockPower = enabled
  If enabled Then
    LockFlipper.RotateToEnd
    If LockFlipper.CurrentAngle < 207 then RandomSoundMetal2 BM_CoffinDiverter
    DiverterWall.isDropped = 1
    DiverterWall001.isDropped = 0
  End If
  If Not enabled AND Not LockHold Then
    LockFlipper.RotateToStart
    RandomSoundMetal2 BM_CoffinDiverter
    DiverterWall.isDropped = 0
    DiverterWall001.isDropped = 1
  End If
End Sub

Sub DiverterHold(enabled)
  LockHold = enabled
  If Not enabled AND Not LockPower Then
    LockFlipper.RotateToStart
    RandomSoundMetal2 BM_CoffinDiverter
    DiverterWall.isDropped = 0
    DiverterWall001.isDropped = 1
  End If
End Sub

Sub SolKickBack(enabled)
    If enabled Then
    Kickback.Fire
  Else
    Kickback.PullBack
  End If
End Sub


Sub CoffinDoor(Enabled)
  If Enabled Then
    CoffinFlipper.RotateToEnd
  Else
    CoffinFlipper.RotateToStart
  End If
End Sub

Sub CoffinFlipper_animate
  Dim BP, a
  a = CoffinFlipper.CurrentAngle
  For each BP in BP_cDoorClose
    BP.roty = a
    BP.visible = (a > -45)
  Next
  For each BP in BP_cDoorOpen
    BP.roty = a
    BP.visible = (a <= -45)
  Next
End Sub

Sub CadaverGate_animate
  Dim BP, a
  a = CadaverGate.CurrentAngle / 3
  if a > 20 then a = 45
  If a < 0 then a = 0
  For each BP in BP_cadaver
    BP.rotx = a
  Next
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
Dim objBallShadow(4)

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
Sub RampTrigger1_Hit
  RandomSoundMetal
  If activeball.vely < 0 Then
    WireRampOn True
  Else
    WireRampOff
  End If
End Sub

Sub RampTrigger2_Hit
  RandomSoundMetal
  If activeball.vely < 0 Then
    WireRampOn True
  Else
    WireRampOff
  End If
End Sub

Sub RampTrigger3_Hit
    if abs(activeball.AngMomZ) > 70 then activeball.AngMomZ = 50
    activeball.AngMomZ = -abs(activeball.AngMomZ) * 3
    WireRampOff
  RandomSoundMetal
  RandomSoundDelayedBallDropOnPlayfield ActiveBall
End Sub

Sub RampTrigger4_Hit
    if abs(activeball.AngMomZ) > 70 then activeball.AngMomZ = 50
    activeball.AngMomZ = abs(activeball.AngMomZ) * 3
  RandomSoundMetal
    WireRampOff
  RandomSoundDelayedBallDropOnPlayfield ActiveBall
End Sub

Sub RampTrigger5_Hit
  WireRampOff
  RandomSoundDelayedBallDropOnPlayfield ActiveBall
End Sub

Sub RampTrigger6_Hit
  WireRampOff
  RandomSoundDelayedBallDropOnPlayfield ActiveBall
End Sub

Sub RampTrigger7_Hit
  WireRampOn True
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

Sub RandomSoundMetal2(tableobj)
  PlaySoundAtVol ("Metal_Touch_" & Int(Rnd * 13) + 1), tableobj, MetalImpactSoundFactor
End Sub

Sub RandomSoundMetal3()
  PlaySoundAtLevelActiveBall "metalhit2", Vol(ActiveBall) * MetalImpactSoundFactor
End Sub

'/////////////////////////////  METAL - EVENTS  ////////////////////////////

Sub Metals_Hit (idx)
  RandomSoundMetal
End Sub

Sub Target_Metal_Hit (idx)
  RandomSoundMetal3
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

Sub ScoopHit()
  Select Case Int(Rnd*2)+1
    Case 1 : PlaySoundAtLevelActiveBall ("ScoopHit1"), VolumeDial
    Case 2 : PlaySoundAtLevelActiveBall ("ScoopHit2"), VolumeDial
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
'Const EOSReturn = 0.055  'EM's
'Const EOSReturn = 0.045  'late 70's to mid 80's
'Const EOSReturn = 0.035  'mid 80's to early 90's
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
Dim ST28, ST61, ST62, ST63, ST64, ST65, ST66

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

Set ST28 = (new StandupTarget)(sw28, BM_sw28, 28, 0)
Set ST61 = (new StandupTarget)(sw61, BM_sw61, 61, 0)
Set ST62 = (new StandupTarget)(sw62, BM_sw62, 62, 0)
Set ST63 = (new StandupTarget)(sw63, BM_sw63, 63, 0)
Set ST64 = (new StandupTarget)(sw64, BM_sw64, 64, 0)
Set ST65 = (new StandupTarget)(sw65, BM_sw65, 65, 0)
Set ST66 = (new StandupTarget)(sw66, BM_sw66, 66, 0)

'Add all the Stand-up Target Arrays to Stand-up Target Animation Array
' STAnimationArray = Array(ST1, ST2, ....)
Dim STArray
STArray = Array(ST28, ST61, ST62, ST63, ST64, ST65, ST66)

Sub UpdateStandupTargets
  dim BP, ty

    ty = BM_sw28.transy
  For each BP in BP_sw28 : BP.transy = ty: Next

    ty = BM_sw61.transy
  For each BP in BP_sw61 : BP.transy = ty: Next

    ty = BM_sw62.transy
  For each BP in BP_sw62 : BP.transy = ty: Next

    ty = BM_sw63.transy
  For each BP in BP_sw63 : BP.transy = ty: Next

    ty = BM_sw64.transy
  For each BP in BP_sw64 : BP.transy = ty / 2: Next
  'For each BP in BP_PrLeaper1 : BP.transz = ty * 10 : Next

    ty = BM_sw65.transy
  For each BP in BP_sw65 : BP.transy = ty / 2: Next
  'For each BP in BP_PrLeaper2 : BP.transz = ty * 10 : Next

  ty = BM_sw66.transy
  For each BP in BP_sw66 : BP.transy = ty / 2: Next
  'For each BP in BP_PrLeaper3 : BP.transz = ty * 10 : Next

End Sub

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


'******************************************************
'***  END STAND-UP TARGETS
'******************************************************

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


'**** Use these for GI strings and stepped GI, and comment out the SetRelayGI
'**** This example table uses Relay for GI control, we don't need these at all

Set GICallback2 = GetRef("GIUpdates2")    'use this for stepped/modulated GI

'GIupdates2 is called always when some event happens to GI channel.
Sub GIUpdates2(aNr, aLvl)
' debug.print "GIUpdates2 nr: " & aNr & " value: " & aLvl
  Dim bulb

  Select Case aNr 'Strings are selected here

    Case 0:  'GI String 1 (Upper)
      ' Update the state for each GI light. The state will be a float value between 0 and 1.
      For Each bulb in GI_Upper: bulb.State = aLvl: Next

      ' If the GI has an associated Relay sound, this can be played
      'If aLvl >= 0.5 And gilvl < 0.5 Then
      ' Sound_GI_Relay 1, Bumper1 'Note: Bumper1 is just used for sound positioning. Can be anywhere that makes sense.
      'ElseIf aLvl <= 0.4 And gilvl > 0.4 Then
      ' Sound_GI_Relay 0, Bumper1
      'End If

      'You may add any other GI related effects here. Like if you want to make some toy to appear more bright, set it like this:
      'Primitive001.blenddisablelighting = 1.2 * aLvl + 0.2 'This will result to DL brightness between 0.2 - 1.4 for ON/OFF states

      gilvl = aLvl    'Storing the latest GI fading state into global variable, so one can use it elsewhere too.
    Case 1:  'GI String 2 (Center)
      ' Update the state for each GI light. The state will be a float value between 0 and 1.
      For Each bulb in GI_Mid: bulb.State = aLvl: Next

      ' If the GI has an associated Relay sound, this can be played
      'If aLvl >= 0.5 And gilvl < 0.5 Then
      ' Sound_GI_Relay 1, Bumper1 'Note: Bumper1 is just used for sound positioning. Can be anywhere that makes sense.
      'ElseIf aLvl <= 0.4 And gilvl > 0.4 Then
      ' Sound_GI_Relay 0, Bumper1
      'End If

      'You may add any other GI related effects here. Like if you want to make some toy to appear more bright, set it like this:
      'Primitive001.blenddisablelighting = 1.2 * aLvl + 0.2 'This will result to DL brightness between 0.2 - 1.4 for ON/OFF states

      gilvl = aLvl    'Storing the latest GI fading state into global variable, so one can use it elsewhere too.
    Case 2:  'GI String 3 (Lower)
      ' Update the state for each GI light. The state will be a float value between 0 and 1.
      For Each bulb in GI_Lower: bulb.State = aLvl: Next

      ' If the GI has an associated Relay sound, this can be played
      'If aLvl >= 0.5 And gilvl < 0.5 Then
      ' Sound_GI_Relay 1, Bumper1 'Note: Bumper1 is just used for sound positioning. Can be anywhere that makes sense.
      'ElseIf aLvl <= 0.4 And gilvl > 0.4 Then
      ' Sound_GI_Relay 0, Bumper1
      'End If

      'You may add any other GI related effects here. Like if you want to make some toy to appear more bright, set it like this:
      'Primitive001.blenddisablelighting = 1.2 * aLvl + 0.2 'This will result to DL brightness between 0.2 - 1.4 for ON/OFF states

      gilvl = aLvl    'Storing the latest GI fading state into global variable, so one can use it elsewhere too.
    Case 3:  'GI String 3 (Backbox)
      For Each bulb in GI_4: bulb.State = aLvl: Next
      gilvl = aLvl
    Case 4:  'GI String 4 (Backbox)
      For Each bulb in GI_5: bulb.State = aLvl: Next
      gilvl = aLvl

  End Select

End Sub

'******************************************************
'****  END GI Control
'******************************************************




'******************************************************
'   ZVBG: VR Backglass
'******************************************************


Sub InitBackglass()
    Dim obj
  Dim xBG, yBG1, yBG2, yBG3, hBG, v

  'Set default backglass positions
  xBG = 0
  yBG1 = -170
  yBG2 = yBG1 + 30
  yBG3 = yBG2 + 30
  hBG = 640 'height
  v = 1

  'Relocate depending on play mode
  If Not (RenderingMode = 2 or VRTest = True) Then
    If DesktopMode = True Then
      xBG = 1200
      yBG1 = 400
      yBG2 = yBG1 + 30
      yBG3 = yBG2 + 30
      hBG = -580 'height
      v = 1
    Else
      v = 0
    End If
  End If

    'Layer 1
    For Each obj In BG_Layer1
        obj.height = - obj.y + hBG
    obj.x = obj.x + xBG
        obj.y = yBG1
        obj.rotx = -90
    obj.visible = v
    Next
    BGDark_Layer1.y = yBG1-1
  BGDark_Layer1.visible = v
  pSpider.x = xBG + 500
  pSpider.y = yBG1 + 10
  pSpider.z = 1230 +(hBG-500)
  pSpider.visible = v

    'Layer 2
  For Each obj In BG_Layer2
        obj.height = - obj.y + hBG
    obj.x = obj.x + xBG
        obj.y = yBG2
        obj.rotx = -90
    obj.visible = v
    Next
    BGDark_Layer2a.y = yBG2-10
  BGDark_Layer2a.visible = v
    BGDark_Layer2b.y = yBG2-10
  BGDark_Layer2b.visible = v

    'Layer 3
  For Each obj In BG_Layer3
        obj.height = - obj.y + hBG
    obj.x = obj.x + xBG
        obj.y = yBG3
        obj.rotx = -90
    obj.visible = v
    Next
  BGDark_Layer3.y = yBG3-10
    BGDark_Layer3.visible = v

  'Bulb temp color
  For Each obj In BG_bulbs
        obj.color = RGB(255,180,80)
    Next

End Sub



'*****  VR Backbox Display Driver   ******



'******************************************************
'   ZVRR: VR Room
'******************************************************

Sub SetRoom
  Dim VRThings
  TimerPlunger2.Enabled = False

  If RenderingMode = 2 or VRTest Then
    VRRoom = VRRoomChoice
    BM_lockdownbar.visible = 0
    For Each VRThings in VRBackGlass: VRThings.visible = 1: Next
    For Each VRThings in VRCabinet: VRThings.visible = 1: Next
    For Each VRThings in VRTopper: VRThings.visible = CustomVRTopper: Next
    PinCab_CustomShooter.Visible = CustomShooter
    TimerPlunger2.Enabled = True
  Else
    VRRoom = 0
    If DesktopMode = True Then
      For Each VRThings in VRBackGlass: VRThings.visible = DTBackglass: Next
    Else
      For Each VRThings in VRBackGlass: VRThings.visible = 0: Next
    End If
    For Each VRThings in VRCabinet: VRThings.visible = 0: Next
    For Each VRThings in VRTopper: VRThings.visible = 0: Next
    for each VRThings in VRMin:VRThings.visible = 0: Next
  End If

  'Set up the room
  If VRRoom = 1 Then
    for each VRThings in VRMin:VRThings.visible = 1: Next
    If VRPosterLeft > 0 Then PosterLeft.Visible = 1 Else PosterLeft.Visible = 0
    If VRPosterRight > 0 Then PosterRight.Visible = 1 Else PosterRight.Visible = 0
  ElseIf VRRoom = 2 Then
    for each VRThings in VRMin:VRThings.visible = 0: Next
  End If

End Sub



'********************************************************
'  ZPFW: Playfield Spider Wheel
'********************************************************


' Sound commands
' 0BF8 = start of spider
' 012C = end of spider (sometimes)

Const hex01 = 1
Const hex0B = 11
Const hexF8 = 248
Const hex2C = 44


Dim SpiderRunning : SpiderRunning = False
Dim LastSnd : LastSnd = 0
Dim SndCmdStr

Sub SoundCmdListener
    Dim NewSounds,ii,Snd
    NewSounds=Controller.NewSoundCommands
    If Not IsEmpty(NewSounds) Then

    ' Listen for specific commands
        For ii=0 To UBound(NewSounds)
            Snd=NewSounds(ii,0)
            If LastSnd = hex0B Then
        If Snd = hexF8 Then
          SpiderRunning = True
          If UsePFWheel = 1 Then ShowPFWheel
          'debug.print "SpiderRunning = True"
        End If
      End If
'     If LastSnd = hex01 Then
'       If Snd = hex2C Then
'         SpiderRunning = False
'         If UsePFWheel = 1 Then HidePFWheel      'HidePFWheel handled when sw36 vuk is kicked
'         'debug.print "SpiderRunning = False"
'       End If
'            End If
      LastSnd = Snd
        Next
    End If
End Sub


HidePFWheel
Sub HidePFWheel
  Dim WheelThings
  For each WheelThings in PFWheel: WheelThings.visible=0: Next
End Sub

Sub ShowPFWheel
  Dim WheelThings
  For each WheelThings in PFWheel: WheelThings.visible=1: Next
End Sub







'Changelog
'=========
'001 - sixtoe - "Just in case" Upload - Unplayable
'002 - sixtoe - Hit a script wall, please send help
'003 - sixtoe - IT'S ALIVE
'004 - sixtoe - Fully Playable
'005 - tomate - First 2k batch added, most of the VPX objects set as not visible,VLM materials applied
'006 - mcarter78 - Add room brightness, ball & plunger textures, make lights hidden, hide a few other trigger objects
'007 - mcarter78 - Movables animation WIP (flippers, gates, crate, targets, bumpers, slings)
'008 - tomate - dozens of fixes in Blender, not so many in VPX
'009 - mcarter78 - Animate Leapers & diverter, fix GI string #s in update callback, Add rules card, cabinet mode & art blades to option menu
'010 - mcarter78 - Reduce height of collidable targets under ramps causing stuck balls
'011 - mcarter78 - Increase bumper ring speed
'012 - mcarter78 - add collections/sfx for walls, metals & apron, add ramp rolling, crate hit, and autoplunger sounds
'013 - mcarter78 - animate closed crate hit & ramp flap from spider VUK, hide ramp triggers
'014 - tomate - some axis fixed in blender, open/close coffin prims added, fixed coffin position in VPX (thanks Apophis!), bumpers and gates elements re-named in blender and changed in the script, backgorund set to black for desktop mode
'015 - apophis - Animated boogie arms. Updated desktop POV. Fixed GI rtx shadows. New ball image.
'016 - tomate - crate and boogies textures fixed, set "hide parts behind" for PF object
'017 - apophis - Adjusted slingshot strength and threshold. Fixed flipper shadows. Added ball shadows. Animated coffin door and cadaver. Added options: desktop DMD visibility, outlane difficulty, ramp refractions
'018 - apophis - Fixed coffin diverter logic.
'019 - tomate - 4k batch added
'020 - tomate - some fixes on blender side, new 4k batch added, VPX unused textures deleted
'021 - apophis - Increased flipper strength to 2500. Increased upper gates elasticity to 0.85
'022 - mcarter78 - Animate autoplunger & leaper rods, less bouncy gates
'023 - cyberpez - backglass first draft
'024 - tomate - new 4k batch added, previous issues fixed (cadaver, top lanes and insert), collidalbe low poly right ramp updated to match visual ramp
'025 - mcarter78 - Fix autoplunger gate, open up the crate more when entering
'026 - tomate - new 4k batch added, targets textures fixed, GIL on group0 issue fixed
'027 - apophis - Backglass second draft. Made BG visible in desktop mode.
'028 - apophis - Added VR Cab. Made desktop backglass optional. Light blocker below playfield. Adjusted visible crate opening angle.
'029 - Sixtoe - Crate VUK rejigged and randomised, various rubber tweaks, redid wall and rubber layout in blocker format, made l53 visible temporarily.
'030 - tomate - new 4k batch added, flippers and l53 fixed, vpx l53 back to hide, changed PF render probe to dynamic, add a hole on PF blocker to see the visual VUK
'031 - tomate - added default LUT and saturation option menu
'032 - apophis - Fixed tilt switch assignment. Changed default LUT is 1to1. Made insert lights reflect on ball. Changed default cab POV. Added info card to tweak menu (thanks FrankEnstein).
'033 - apophis - Improved post passes. Added inlane slowdown code. Added custom shooter knob option (thanks DGrimmReaper). Added minimal VR room with movie poster options.
'034 - apophis - Fixed spider b2s issues. Disabled reflections of some items.
'RC1 - apophis - Fixed boogie man analog nudging. Updated some backglass flashers (thanks cyberpez). Updated custom rules card in tweak menu (thanks FrankEnstein). Fixed knocker sfx. Updated frog target hit sfx and scoop entrance sfx. Fixed VR poster visible bug.
'RC2 - mcarter78 - Only animate crate door on hit if not already currently animating
'RC3 - apophis - Solved crate door issue? Added playfield reflection options. Lowered plunger speed.
'RC4 - apophis - Accounted for frame rate in frog speed calcs. Adjusted down frog leap speed. Slowed down right vuk eject and ramps (slightly). Slingshots strength set to 4.2 from 4.5
'RC5 - mcater78 - lowered ball height threshold for vuk trapdoor to  >= 109 from > 110 and added a Y value failsafe
'RC6 - apophis - Updated (optimized) the reflection scope option. Improved spider wheel icon images (thanks HauntFreaks). Chnaged z lift on sw36 kicker from 100 to 105.
'RC7 - Sixtoe - Added metal hit to leaper targets, refactored entire cadaver coffin exit area,
'Release 1.0
'1.0.1 - apophis - Added playfield spider wheel option. Added ramp rolling sound upon spider vuk eject.
'1.0.2 - apophis - Added ability to use "Boogieman Dance" ROM feature. Fixed issue with new PF spider wheel option. Fixed hole in VR sideblades near lockdown bar.
'1.0.3 - apophis - Fix VR plunger pullback animation. Fixed hole in back of VR backbox.
'1.0.4 - apophis - Fix sling animation steps.
'1.0.5 - tomate - leapers' descent speed tuned down
'1.0.6 - retro27 - added custom VR topper
'1.0.7 - apophis - added lights to custom VR topper. Updated EOSReturn.
'Release 1.1
'1.1.1 - apophis - Fixed cab pov issue. Added improved topper image and lights (thanks Retro27).
'Release 1.1.1

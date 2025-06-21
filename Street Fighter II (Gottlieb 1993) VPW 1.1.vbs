'
'                              ......
'                       ..::--:     *-
'                    .-=.           .+++                                ...:.:--======.
'                 .-:               .@%:..              ..   --=+++**=.    **.      :+.
'               .=.           ..-=*@@@@@%=     ...:-=:    .=%     .=#  :=*#@@@@.  %@@@@@:
'              =.         =@@@@@@@@@@@@* .-+-  -#+         %.  +@@@@.    .@@@@:  @@@@@@:
'            ::       .%@@@@@@@@+.   -+.       :#*   #@=  +:     :@-   :=+--#.  @@@@=                                         -=---::::::::...
'           ::      .@@@@@@+*=--+*-. -=      .%@@@: ==  *@@.  =%@@*  +@@%@@@-  %@@@.                                        :#.                         .-+*#=.
'           -:       *%-..         ::#-.:   :@@@@= :%+  @@=   ..-@      .-=:  +@@@.            ..:                        #=:....                      =@@@@@*:
'           -:                      *=. *  .@@@@+ .#@*  =@:    .%=.:+@@@@@+. -@@@:..:=+=:. ..::==+*=. .......:=.          .-@@@@@@@%:   -@@@    -@@@@@@@@@@+-:
'            -:       :+@@@@#       @@-=   %@@@+  +@@@-  @@+@@@@@@@@@@@@@@%=%%+--:.     .:#@*-:.               .:.         -------+-   .@@@-   .@@@****+:.
'             -%@@@@@@@@@@@@+      *@@+:  #@@@=  -@@@+:+#*=:--:*@%+==+#+::.        .%@@@@@@@+          .--      ::               :-    %@@=   .@@#
'              .=#@@@*+=-=:.     .#@@@+. -@@@.-*#++=:.     .--#@%%#+++   .=*#@@@. .@@@@@@@@%#+.     %@@@@@@+     +=             :=.   *@@*.   #@@.
'                     =.:      .*@@@@::=@#*+-.           .=#%@@@@-  -=+@@@@@@@@-  %%=.     -==-    %@@@@@@@+     #@-           .-:   =@@%:   =@@:
'                 -:-.       :#@@@@*--.            .=*%@@@@@@@@@    .:-%@@+. +*. -@@  .--#@%#=    *@@@.  -=     +@@+           :-   :@@@-   .@@+
'            ::.-.        :#@@%+:            .-*#@@@@@@@@@@%##=..=.:- .:-:  +*.  @@.   . +@@#.   +@@@: --:     +@@@=          .=   .#@@=.   @@%
'      ::.:-.         .=#@%=               .@@@@@@+. :=-+.    :-=++=  +@#: .@-  %@=  =++=-==.   -@@@%:*.     :%@@@+          .=    +@@#:   #@@.
' .++:             -+@@@%         .-=.     -+.-##-  -*+.   :==%@@@@.       %%. :@#  =%+*@@@-    #@#:      :*@@@@@=           -.   -@@@-   -@@-
'   :=.  ..   .:=@@@@@@=      .-#@@@@:       .=%:  =@=   +@@@@@#-#%       :@-  @@      -@#=             %@@@@@@*            .-   .%@@=.  .@@*
'   :++=*..:*@@@@@@@@%.*. .-#@@@@@+.      :+@@@:  =@=   #@%.    .@   *@+  %*. *@- +@@@@@-       ::       :@@@-              *    +@@+:   %@@.
'  ..+--%@@@@@@@=. . .*:*@@@@@@@#.       @@@@@=  -@*   .@@@.  .-@*  =@#  *@- .@@#:@@@@:      *@@@@@:       .=:.            *.   =@@%-   +@@-
'  ..  =@@@*.       .+-.-@@%:.-+:=@%:   @@@@:=  :%@*         .@@@+ =@@*:=@+  @@@+@*@*#*-    *@@@@@@@+          -+-        -:   :@@@=.  .@@*
'    ::.               .:     . -=*:   #@@@ :.  +@@@=       .@@@=--@@@*#@#. :@@-     .:    +@@@=    @@:             .:-+#:=   .*@@=:   #@@:
'                             .#%#-   -@@@ .%* =@@* :@@@%.  @@@=.%@@@# .#@#:@@@.    #     =@@@+      .@@-               +%    -@@#=   =@@-
'                                .=: .@@@   :*.*@@.   =@@@*@@@%    =*    .+#@@-    +*###%@@@@*         -@@@=          :#@:   :@@@=.  .%@*.
'                                  *@@@@:    .*@@*       .@@@@            :#@#.      :@@@@@@#.           =@@@@@*:    +@@*    =@@=-   +@@:
'                                   *@@#       =%.        ..=               @:       :------.              :*@@@@@@@@@@@.   .@@@+   .%@+
'                                                                                                             =*:                    :-----==++=-.
'                                                                                                            -+=*%%%####***+==-:..             =+===:
'                                                                                                             :%@@@@@@@@@@@@@@@@@@@@@@@@@@@@%#@@@@@#+-
'                                                                                                                .....:::::---==++**##%%@@@@@@@@@.
'
' Street Fighter II
' Premier/Gottlieb 1993
' https://www.ipdb.org/machine.cgi?id=2403

' VPW TEAM
' Gedankekojote97 - Blender, VR rooms, table build
' Apophis - Scripts, phyiscs, animations, table
' CainArg - Graphics cleanup
' Sixtoe - Table support

' Reference: VPX8 table by jpsalas


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
' ZDRN: Drain, Trough, and Ball Release
' ZDIV: Diverters
' ZCHF: Chunli flipper
' ZCAR: Car Crunch
' ZFLP: Flippers
' ZSLG: Slingshot Animations
' ZSSC: Slingshot Corrections
' ZVUK: VUKs and Kickers
' ZSWI: Switches
'   ZBRL: Ball Rolling and Drop Sounds
'   ZRRL: Ramp Rolling Sound Effects
'   ZFLE: Fleep Mechanical Sounds
' ZNFF: Flipper Corrections
'   ZDMP: Rubber Dampeners
'   ZBOU: VPW TargetBouncer for targets and posts
'   ZRDT: Drop Targets
' ZRST: Stand-Up Targets
'   ZGIU: GI updates
' ZKNO: Knocker Sound
'   ZINS: Inserts Special Cases
' ZSHA: Ambient Ball Shadows
'   ZTPR: Update Topper
' ZVRR: VR Room
' ZDBG: Debug Code
'   ZLOG: Change Log
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
Const TestVRonDT = False

Const BallSize = 50   'Ball size must be 50
Const BallMass = 1    'Ball mass must be 1
Const tnob = 4      'Total number of balls on the playfield including captive balls.
Const lob = 0     'Total number of locked balls

'  Standard definitions
Const cGameName = "sfight2"   'PinMAME ROM name
Const UseVPMModSol = 2
Const UseSolenoids = 2      '1 = Normal Flippers, 2 = Fastflips
Const UseLamps = 1        '0 = Custom lamp handling, 1 = Built-in VPX handling (using light number in light timer)
Const UseSync = 0
Const HandleMech = 0
Const SSolenoidOn = ""      'Sound sample used for this, obsolete.
Const SSolenoidOff = ""     ' ^
Const SFlipperOn = ""     ' ^
Const SFlipperOff = ""      ' ^
Const SCoin = ""        '
Const UseGI = 1
Const swLCoin = 0
Const swRCoin = 1
Const swCCoin = 2
Const swCoinShuttle = 3
Const swStartButton = 4
Const swTournament = 5
Const swFrontDoor = 6

'Internal DMD in Desktop Mode, using a textbox (must be called before LoadVPM)
Dim UseVPMDMD, VRRoom, DesktopMode, VarHidden, UseVPMColoredDMD
DesktopMode = Table1.ShowDT
If RenderingMode = 2 OR TestVRonDT=True Then UseVPMDMD = True Else UseVPMDMD = DesktopMode
If DesktopMode = true then
    UseVPMColoredDMD = true
    VarHidden = 1
Else
    UseVPMColoredDMD = False
    VarHidden = 0
End If

LoadVPM "03060000", "gts3.VBS", 3.10

' Globals
Dim tablewidth: tablewidth = Table1.width
Dim tableheight: tableheight = Table1.height
Dim PlungerIM
Dim SFBall1, SFBall2, SFBall3, CarBall, gBOT
Dim bsTrough, bsLKicker, bsL2Kicker, bsL3Kicker, bsRLowKicker, x
Dim LowerFlipper:LowerFlipper = False


'******************************************************
'  ZVLM: VLM Arrays
'******************************************************

' VLM  Arrays - Start
' Arrays per baked part
Dim BP_Bumper1Ring: BP_Bumper1Ring=Array(BM_Bumper1Ring, LM_F_F15_Bumper1Ring, LM_F_F17_Bumper1Ring, LM_I_L35_Bumper1Ring, LM_I_L36_Bumper1Ring, LM_I_L37_Bumper1Ring, LM_I_L45_Bumper1Ring, LM_F_L57_Bumper1Ring, LM_Gi_Bumper1Ring)
Dim BP_Car: BP_Car=Array(BM_Car, LM_CC_L02_Car, LM_CC_L03_Car, LM_CC_L04_Car, LM_CC_L05_Car, LM_CC_L06_Car, LM_CC_L07_Car, LM_I_L31_Car, LM_I_L46_Car, LM_I_L60_Car)
Dim BP_Diverter1: BP_Diverter1=Array(BM_Diverter1, LM_F_F17_Diverter1, LM_Gi2_gi03_Diverter1)
Dim BP_Diverter2: BP_Diverter2=Array(BM_Diverter2, LM_F_F20_Diverter2, LM_Gi2_gi12_Diverter2)
Dim BP_Gate1: BP_Gate1=Array(BM_Gate1, LM_F_F16_Gate1, LM_Gi_Gate1)
Dim BP_LFlip: BP_LFlip=Array(BM_LFlip, LM_F_F17_LFlip, LM_F_F18_LFlip, LM_F_F19_LFlip, LM_F_F23_LFlip, LM_I_L0_LFlip, LM_I_L14_LFlip, LM_Gi2_gi01_LFlip, LM_Gi2_gi02_LFlip, LM_Gi2_gi03_LFlip, LM_Gi_LFlip, LM_Gi2_gi13_LFlip, LM_Gi2_gi14_LFlip)
Dim BP_LFlip1: BP_LFlip1=Array(BM_LFlip1, LM_F_F16_LFlip1, LM_F_F17_LFlip1, LM_F_F18_LFlip1, LM_F_F19_LFlip1, LM_F_F20_LFlip1, LM_Gi2_gi02_LFlip1, LM_Gi2_gi03_LFlip1, LM_Gi_LFlip1)
Dim BP_LFlip1U: BP_LFlip1U=Array(BM_LFlip1U, LM_F_F16_LFlip1U, LM_F_F17_LFlip1U, LM_F_F18_LFlip1U, LM_F_F19_LFlip1U, LM_F_F20_LFlip1U, LM_F_F23_LFlip1U, LM_I_L40_LFlip1U, LM_I_L50_LFlip1U, LM_I_L51_LFlip1U, LM_I_L52_LFlip1U, LM_Gi2_gi01_LFlip1U, LM_Gi2_gi02_LFlip1U, LM_Gi2_gi03_LFlip1U, LM_Gi_LFlip1U, LM_Gi2_gi12_LFlip1U, LM_Gi2_gi13_LFlip1U, LM_Gi2_gi14_LFlip1U)
Dim BP_LFlipU: BP_LFlipU=Array(BM_LFlipU, LM_F_F17_LFlipU, LM_F_F18_LFlipU, LM_F_F19_LFlipU, LM_F_F23_LFlipU, LM_I_L0_LFlipU, LM_I_L14_LFlipU, LM_Gi2_gi01_LFlipU, LM_Gi2_gi02_LFlipU, LM_Gi2_gi03_LFlipU, LM_Gi_LFlipU, LM_Gi2_gi14_LFlipU)
Dim BP_LPF: BP_LPF=Array(BM_LPF, LM_CC_L02_LPF, LM_CC_L03_LPF, LM_CC_L04_LPF, LM_CC_L05_LPF, LM_CC_L06_LPF, LM_CC_L07_LPF, LM_I_L30_LPF, LM_I_L31_LPF, LM_I_L46_LPF, LM_I_L60_LPF)
Dim BP_LRampArm: BP_LRampArm=Array(BM_LRampArm, LM_F_F18_LRampArm, LM_I_L41_LRampArm, LM_I_L42_LRampArm, LM_Gi_LRampArm)
Dim BP_LRampArmU: BP_LRampArmU=Array(BM_LRampArmU, LM_F_F18_LRampArmU, LM_I_L53_LRampArmU, LM_Gi_LRampArmU)
Dim BP_LRampU: BP_LRampU=Array(BM_LRampU, LM_F_F18_LRampU, LM_I_L41_LRampU, LM_I_L42_LRampU, LM_I_L43_LRampU, LM_I_L44_LRampU, LM_I_L53_LRampU, LM_Gi_LRampU)
Dim BP_LRamp_: BP_LRamp_=Array(BM_LRamp_, LM_F_F18_LRamp_, LM_I_L43_LRamp_, LM_Gi_LRamp_)
Dim BP_LSling: BP_LSling=Array(BM_LSling, LM_F_F19_LSling, LM_F_F23_LSling, LM_I_L11_LSling, LM_I_L15_LSling, LM_Gi2_gi01_LSling, LM_Gi2_gi02_LSling, LM_Gi2_gi03_LSling, LM_Gi2_gi13_LSling)
Dim BP_LSling1: BP_LSling1=Array(BM_LSling1, LM_F_F19_LSling1, LM_I_L15_LSling1, LM_Gi2_gi01_LSling1, LM_Gi2_gi02_LSling1, LM_Gi2_gi03_LSling1)
Dim BP_LSling2: BP_LSling2=Array(BM_LSling2, LM_F_F19_LSling2, LM_I_L15_LSling2, LM_Gi2_gi01_LSling2, LM_Gi2_gi02_LSling2, LM_Gi2_gi03_LSling2)
Dim BP_Layer1: BP_Layer1=Array(BM_Layer1, LM_CC_L02_Layer1, LM_CC_L03_Layer1, LM_CC_L04_Layer1, LM_CC_L05_Layer1, LM_CC_L06_Layer1, LM_CC_L07_Layer1, LM_I_L60_Layer1)
Dim BP_Layer2: BP_Layer2=Array(BM_Layer2, LM_CC_L02_Layer2, LM_CC_L03_Layer2, LM_CC_L04_Layer2, LM_CC_L05_Layer2, LM_CC_L06_Layer2, LM_CC_L07_Layer2, LM_I_L60_Layer2)
Dim BP_Layer3: BP_Layer3=Array(BM_Layer3, LM_F_F15_Layer3, LM_F_F16_Layer3, LM_F_F17_Layer3, LM_F_F18_Layer3, LM_F_F20_Layer3, LM_F_F22_Layer3, LM_I_L34_Layer3, LM_I_L35_Layer3, LM_I_L36_Layer3, LM_I_L37_Layer3, LM_I_L42_Layer3, LM_I_L47_Layer3, LM_F_L57_Layer3, LM_Gi2_gi03_Layer3, LM_Gi_Layer3)
Dim BP_Layer4: BP_Layer4=Array(BM_Layer4, LM_F_F21_Layer4, LM_F_F22_Layer4, LM_Gi_Layer4)
Dim BP_Layer5: BP_Layer5=Array(BM_Layer5, LM_F_F15_Layer5, LM_F_F16_Layer5, LM_F_F18_Layer5, LM_F_F21_Layer5, LM_F_F22_Layer5, LM_I_L35_Layer5, LM_I_L36_Layer5, LM_I_L37_Layer5, LM_I_L43_Layer5, LM_I_L44_Layer5, LM_I_L55_Layer5, LM_I_L56_Layer5, LM_F_L57_Layer5, LM_Gi_Layer5)
Dim BP_Layer6: BP_Layer6=Array(BM_Layer6, LM_F_F15_Layer6, LM_F_F16_Layer6, LM_F_F18_Layer6, LM_Gi_Layer6)
Dim BP_Mov3: BP_Mov3=Array(BM_Mov3, LM_F_F15_Mov3, LM_F_F16_Mov3, LM_F_F18_Mov3, LM_F_F21_Mov3, LM_F_F22_Mov3, LM_Gi_Mov3)
Dim BP_PF: BP_PF=Array(BM_PF, LM_F_F15_PF, LM_F_F16_PF, LM_F_F17_PF, LM_F_F18_PF, LM_F_F19_PF, LM_F_F20_PF, LM_F_F21_PF, LM_F_F22_PF, LM_F_F23_PF, LM_CC_L02_PF, LM_CC_L03_PF, LM_CC_L04_PF, LM_CC_L05_PF, LM_CC_L06_PF, LM_CC_L07_PF, LM_I_L0_PF, LM_I_L10_PF, LM_I_L11_PF, LM_I_L12_PF, LM_I_L13_PF, LM_I_L14_PF, LM_I_L15_PF, LM_I_L16_PF, LM_I_L17_PF, LM_I_L20_PF, LM_I_L21_PF, LM_I_L22_PF, LM_I_L23_PF, LM_I_L24_PF, LM_I_L25_PF, LM_I_L26_PF, LM_I_L27_PF, LM_I_L30_PF, LM_I_L31_PF, LM_I_L32_PF, LM_I_L33_PF, LM_I_L34_PF, LM_I_L35_PF, LM_I_L36_PF, LM_I_L37_PF, LM_I_L40_PF, LM_I_L41_PF, LM_I_L42_PF, LM_I_L43_PF, LM_I_L44_PF, LM_I_L45_PF, LM_I_L46_PF, LM_I_L47_PF, LM_I_L50_PF, LM_I_L51_PF, LM_I_L52_PF, LM_I_L53_PF, LM_I_L54_PF, LM_I_L55_PF, LM_I_L56_PF, LM_F_L57_PF, LM_Gi2_gi01_PF, LM_Gi2_gi02_PF, LM_Gi2_gi03_PF, LM_Gi_PF, LM_Gi2_gi11_PF, LM_Gi2_gi12_PF, LM_Gi2_gi13_PF, LM_Gi2_gi14_PF)
Dim BP_Parts: BP_Parts=Array(BM_Parts, LM_F_F15_Parts, LM_F_F16_Parts, LM_F_F17_Parts, LM_F_F18_Parts, LM_F_F19_Parts, LM_F_F20_Parts, LM_F_F21_Parts, LM_F_F22_Parts, LM_F_F23_Parts, LM_CC_L02_Parts, LM_CC_L03_Parts, LM_CC_L04_Parts, LM_CC_L05_Parts, LM_CC_L06_Parts, LM_CC_L07_Parts, LM_I_L0_Parts, LM_I_L10_Parts, LM_I_L11_Parts, LM_I_L12_Parts, LM_I_L13_Parts, LM_I_L14_Parts, LM_I_L15_Parts, LM_I_L16_Parts, LM_I_L17_Parts, LM_I_L20_Parts, LM_I_L21_Parts, LM_I_L22_Parts, LM_I_L23_Parts, LM_I_L24_Parts, LM_I_L25_Parts, LM_I_L26_Parts, LM_I_L27_Parts, LM_I_L30_Parts, LM_I_L31_Parts, LM_I_L32_Parts, LM_I_L33_Parts, LM_I_L34_Parts, LM_I_L35_Parts, LM_I_L36_Parts, LM_I_L37_Parts, LM_I_L40_Parts, LM_I_L41_Parts, LM_I_L42_Parts, LM_I_L43_Parts, LM_I_L44_Parts, LM_I_L45_Parts, LM_I_L46_Parts, LM_I_L47_Parts, LM_I_L52_Parts, LM_I_L53_Parts, LM_I_L54_Parts, LM_I_L55_Parts, LM_I_L56_Parts, LM_F_L57_Parts, LM_I_L60_Parts, LM_A_L61_Parts, LM_A_L62_Parts, LM_A_L63_Parts, LM_A_L64_Parts, LM_A_L65_Parts, LM_A_L66_Parts, _
  LM_A_L67_Parts, LM_Gi2_gi01_Parts, LM_Gi2_gi02_Parts, LM_Gi2_gi03_Parts, LM_Gi_Parts, LM_Gi2_gi11_Parts, LM_Gi2_gi12_Parts, LM_Gi2_gi13_Parts, LM_Gi2_gi14_Parts)
Dim BP_Plunger: BP_Plunger=Array(BM_Plunger)
Dim BP_RFlip: BP_RFlip=Array(BM_RFlip, LM_F_F15_RFlip, LM_F_F18_RFlip, LM_F_F19_RFlip, LM_F_F23_RFlip, LM_I_L0_RFlip, LM_I_L14_RFlip, LM_Gi2_gi01_RFlip, LM_Gi2_gi02_RFlip, LM_Gi2_gi03_RFlip, LM_Gi_RFlip, LM_Gi2_gi13_RFlip, LM_Gi2_gi14_RFlip)
Dim BP_RFlip1: BP_RFlip1=Array(BM_RFlip1, LM_CC_L02_RFlip1, LM_CC_L03_RFlip1, LM_CC_L04_RFlip1, LM_CC_L05_RFlip1, LM_CC_L06_RFlip1, LM_I_L60_RFlip1)
Dim BP_RFlip1U: BP_RFlip1U=Array(BM_RFlip1U, LM_CC_L02_RFlip1U, LM_CC_L03_RFlip1U, LM_CC_L04_RFlip1U, LM_CC_L05_RFlip1U, LM_CC_L06_RFlip1U, LM_CC_L07_RFlip1U, LM_I_L60_RFlip1U)
Dim BP_RFlipU: BP_RFlipU=Array(BM_RFlipU, LM_F_F15_RFlipU, LM_F_F18_RFlipU, LM_F_F19_RFlipU, LM_F_F23_RFlipU, LM_I_L0_RFlipU, LM_I_L14_RFlipU, LM_Gi2_gi01_RFlipU, LM_Gi2_gi03_RFlipU, LM_Gi_RFlipU, LM_Gi2_gi13_RFlipU, LM_Gi2_gi14_RFlipU)
Dim BP_RRamp: BP_RRamp=Array(BM_RRamp, LM_F_F20_RRamp, LM_F_F21_RRamp, LM_Gi2_gi02_RRamp, LM_Gi_RRamp, LM_Gi2_gi11_RRamp, LM_Gi2_gi12_RRamp)
Dim BP_RRampArm: BP_RRampArm=Array(BM_RRampArm, LM_F_F20_RRampArm, LM_F_F21_RRampArm, LM_I_L27_RRampArm, LM_Gi_RRampArm, LM_Gi2_gi12_RRampArm)
Dim BP_RRampArmU: BP_RRampArmU=Array(BM_RRampArmU, LM_F_F20_RRampArmU, LM_Gi_RRampArmU)
Dim BP_RRampU: BP_RRampU=Array(BM_RRampU, LM_F_F20_RRampU, LM_F_F21_RRampU, LM_Gi_RRampU, LM_Gi2_gi11_RRampU)
Dim BP_RSling: BP_RSling=Array(BM_RSling, LM_F_F17_RSling, LM_F_F18_RSling, LM_F_F19_RSling, LM_F_F20_RSling, LM_F_F23_RSling, LM_I_L12_RSling, LM_I_L20_RSling, LM_I_L21_RSling, LM_Gi2_gi02_RSling, LM_Gi2_gi12_RSling, LM_Gi2_gi13_RSling, LM_Gi2_gi14_RSling)
Dim BP_RSling1: BP_RSling1=Array(BM_RSling1, LM_F_F18_RSling1, LM_F_F19_RSling1, LM_F_F20_RSling1, LM_F_F23_RSling1, LM_I_L12_RSling1, LM_I_L20_RSling1, LM_I_L21_RSling1, LM_Gi2_gi12_RSling1, LM_Gi2_gi13_RSling1, LM_Gi2_gi14_RSling1)
Dim BP_RSling2: BP_RSling2=Array(BM_RSling2, LM_F_F17_RSling2, LM_F_F18_RSling2, LM_F_F20_RSling2, LM_F_F23_RSling2, LM_I_L12_RSling2, LM_I_L20_RSling2, LM_I_L21_RSling2, LM_Gi2_gi12_RSling2, LM_Gi2_gi13_RSling2, LM_Gi2_gi14_RSling2)
Dim BP_SideRails: BP_SideRails=Array(BM_SideRails, LM_F_F15_SideRails, LM_F_F16_SideRails, LM_F_F18_SideRails, LM_F_F19_SideRails, LM_F_F21_SideRails, LM_F_F22_SideRails, LM_Gi_SideRails)
Dim BP_Sling1: BP_Sling1=Array(BM_Sling1, LM_F_F19_Sling1, LM_CC_L02_Sling1, LM_Gi2_gi01_Sling1, LM_Gi2_gi02_Sling1)
Dim BP_Sling2: BP_Sling2=Array(BM_Sling2, LM_F_F23_Sling2, LM_CC_L05_Sling2, LM_I_L20_Sling2, LM_Gi2_gi13_Sling2, LM_Gi2_gi14_Sling2)
Dim BP_SpinFlip1: BP_SpinFlip1=Array(BM_SpinFlip1, LM_F_F15_SpinFlip1, LM_F_F16_SpinFlip1, LM_F_F22_SpinFlip1, LM_F_L57_SpinFlip1, LM_Gi_SpinFlip1)
Dim BP_SpinFlip2: BP_SpinFlip2=Array(BM_SpinFlip2, LM_F_F15_SpinFlip2, LM_F_F16_SpinFlip2, LM_F_F21_SpinFlip2, LM_F_F22_SpinFlip2, LM_Gi_SpinFlip2)
Dim BP_SpinFlip3: BP_SpinFlip3=Array(BM_SpinFlip3, LM_F_F15_SpinFlip3, LM_F_F16_SpinFlip3, LM_Gi_SpinFlip3)
Dim BP_SpinFlip4: BP_SpinFlip4=Array(BM_SpinFlip4, LM_F_F15_SpinFlip4, LM_F_F16_SpinFlip4, LM_F_L57_SpinFlip4, LM_Gi_SpinFlip4)
Dim BP_WireDiverter1: BP_WireDiverter1=Array(BM_WireDiverter1, LM_F_F21_WireDiverter1, LM_F_F22_WireDiverter1, LM_Gi_WireDiverter1)
Dim BP_sw102: BP_sw102=Array(BM_sw102, LM_I_L36_sw102, LM_I_L54_sw102, LM_I_L55_sw102, LM_I_L56_sw102, LM_F_L57_sw102, LM_Gi_sw102)
Dim BP_sw103: BP_sw103=Array(BM_sw103, LM_F_F17_sw103, LM_I_L33_sw103, LM_I_L34_sw103, LM_I_L35_sw103, LM_I_L43_sw103, LM_I_L46_sw103, LM_I_L47_sw103, LM_F_L57_sw103, LM_Gi_sw103, LM_Gi2_gi13_sw103)
Dim BP_sw104: BP_sw104=Array(BM_sw104, LM_F_F17_sw104, LM_F_F18_sw104, LM_F_F20_sw104, LM_I_L30_sw104, LM_I_L31_sw104, LM_Gi_sw104)
Dim BP_sw105: BP_sw105=Array(BM_sw105, LM_F_F17_sw105, LM_F_F19_sw105, LM_CC_L03_sw105, LM_I_L11_sw105, LM_Gi2_gi02_sw105, LM_Gi2_gi03_sw105)
Dim BP_sw106: BP_sw106=Array(BM_sw106, LM_F_F23_sw106, LM_CC_L06_sw106, LM_I_L12_sw106, LM_Gi2_gi12_sw106, LM_Gi2_gi13_sw106)
Dim BP_sw112: BP_sw112=Array(BM_sw112, LM_I_L35_sw112, LM_I_L36_sw112, LM_I_L37_sw112, LM_I_L55_sw112, LM_I_L56_sw112, LM_F_L57_sw112, LM_Gi_sw112)
Dim BP_sw113: BP_sw113=Array(BM_sw113, LM_F_F17_sw113, LM_I_L33_sw113, LM_I_L34_sw113, LM_I_L35_sw113, LM_I_L46_sw113, LM_I_L47_sw113, LM_F_L57_sw113, LM_Gi_sw113)
Dim BP_sw114: BP_sw114=Array(BM_sw114, LM_Gi_sw114)
Dim BP_sw115: BP_sw115=Array(BM_sw115)
Dim BP_sw116: BP_sw116=Array(BM_sw116)
Dim BP_sw20: BP_sw20=Array(BM_sw20, LM_Gi_sw20)
Dim BP_sw22: BP_sw22=Array(BM_sw22)
Dim BP_sw30: BP_sw30=Array(BM_sw30, LM_I_L35_sw30, LM_Gi_sw30)
Dim BP_sw32: BP_sw32=Array(BM_sw32, LM_Gi_sw32)
Dim BP_sw92: BP_sw92=Array(BM_sw92, LM_F_F21_sw92, LM_I_L36_sw92, LM_I_L37_sw92, LM_I_L44_sw92, LM_I_L45_sw92, LM_I_L54_sw92, LM_I_L55_sw92, LM_I_L56_sw92, LM_F_L57_sw92, LM_Gi_sw92)
Dim BP_sw93: BP_sw93=Array(BM_sw93, LM_F_F17_sw93, LM_I_L43_sw93, LM_I_L44_sw93, LM_I_L46_sw93, LM_F_L57_sw93, LM_Gi_sw93)
Dim BP_sw94: BP_sw94=Array(BM_sw94, LM_F_F18_sw94, LM_I_L30_sw94, LM_Gi_sw94)
Dim BP_sw95: BP_sw95=Array(BM_sw95, LM_F_F17_sw95, LM_F_F19_sw95, LM_I_L10_sw95, LM_Gi2_gi02_sw95, LM_Gi2_gi03_sw95)
Dim BP_sw96: BP_sw96=Array(BM_sw96, LM_F_F23_sw96, LM_I_L13_sw96, LM_Gi2_gi12_sw96, LM_Gi2_gi13_sw96)
' Arrays per lighting scenario
Dim BL_A_L61: BL_A_L61=Array(LM_A_L61_Parts)
Dim BL_A_L62: BL_A_L62=Array(LM_A_L62_Parts)
Dim BL_A_L63: BL_A_L63=Array(LM_A_L63_Parts)
Dim BL_A_L64: BL_A_L64=Array(LM_A_L64_Parts)
Dim BL_A_L65: BL_A_L65=Array(LM_A_L65_Parts)
Dim BL_A_L66: BL_A_L66=Array(LM_A_L66_Parts)
Dim BL_A_L67: BL_A_L67=Array(LM_A_L67_Parts)
Dim BL_CC_L02: BL_CC_L02=Array(LM_CC_L02_Car, LM_CC_L02_LPF, LM_CC_L02_Layer1, LM_CC_L02_Layer2, LM_CC_L02_PF, LM_CC_L02_Parts, LM_CC_L02_RFlip1, LM_CC_L02_RFlip1U, LM_CC_L02_Sling1)
Dim BL_CC_L03: BL_CC_L03=Array(LM_CC_L03_Car, LM_CC_L03_LPF, LM_CC_L03_Layer1, LM_CC_L03_Layer2, LM_CC_L03_PF, LM_CC_L03_Parts, LM_CC_L03_RFlip1, LM_CC_L03_RFlip1U, LM_CC_L03_sw105)
Dim BL_CC_L04: BL_CC_L04=Array(LM_CC_L04_Car, LM_CC_L04_LPF, LM_CC_L04_Layer1, LM_CC_L04_Layer2, LM_CC_L04_PF, LM_CC_L04_Parts, LM_CC_L04_RFlip1, LM_CC_L04_RFlip1U)
Dim BL_CC_L05: BL_CC_L05=Array(LM_CC_L05_Car, LM_CC_L05_LPF, LM_CC_L05_Layer1, LM_CC_L05_Layer2, LM_CC_L05_PF, LM_CC_L05_Parts, LM_CC_L05_RFlip1, LM_CC_L05_RFlip1U, LM_CC_L05_Sling2)
Dim BL_CC_L06: BL_CC_L06=Array(LM_CC_L06_Car, LM_CC_L06_LPF, LM_CC_L06_Layer1, LM_CC_L06_Layer2, LM_CC_L06_PF, LM_CC_L06_Parts, LM_CC_L06_RFlip1, LM_CC_L06_RFlip1U, LM_CC_L06_sw106)
Dim BL_CC_L07: BL_CC_L07=Array(LM_CC_L07_Car, LM_CC_L07_LPF, LM_CC_L07_Layer1, LM_CC_L07_Layer2, LM_CC_L07_PF, LM_CC_L07_Parts, LM_CC_L07_RFlip1U)
Dim BL_F_F15: BL_F_F15=Array(LM_F_F15_Bumper1Ring, LM_F_F15_Layer3, LM_F_F15_Layer5, LM_F_F15_Layer6, LM_F_F15_Mov3, LM_F_F15_PF, LM_F_F15_Parts, LM_F_F15_RFlip, LM_F_F15_RFlipU, LM_F_F15_SideRails, LM_F_F15_SpinFlip1, LM_F_F15_SpinFlip2, LM_F_F15_SpinFlip3, LM_F_F15_SpinFlip4)
Dim BL_F_F16: BL_F_F16=Array(LM_F_F16_Gate1, LM_F_F16_LFlip1, LM_F_F16_LFlip1U, LM_F_F16_Layer3, LM_F_F16_Layer5, LM_F_F16_Layer6, LM_F_F16_Mov3, LM_F_F16_PF, LM_F_F16_Parts, LM_F_F16_SideRails, LM_F_F16_SpinFlip1, LM_F_F16_SpinFlip2, LM_F_F16_SpinFlip3, LM_F_F16_SpinFlip4)
Dim BL_F_F17: BL_F_F17=Array(LM_F_F17_Bumper1Ring, LM_F_F17_Diverter1, LM_F_F17_LFlip, LM_F_F17_LFlip1, LM_F_F17_LFlip1U, LM_F_F17_LFlipU, LM_F_F17_Layer3, LM_F_F17_PF, LM_F_F17_Parts, LM_F_F17_RSling, LM_F_F17_RSling2, LM_F_F17_sw103, LM_F_F17_sw104, LM_F_F17_sw105, LM_F_F17_sw113, LM_F_F17_sw93, LM_F_F17_sw95)
Dim BL_F_F18: BL_F_F18=Array(LM_F_F18_LFlip, LM_F_F18_LFlip1, LM_F_F18_LFlip1U, LM_F_F18_LFlipU, LM_F_F18_LRamp_, LM_F_F18_LRampArm, LM_F_F18_LRampArmU, LM_F_F18_LRampU, LM_F_F18_Layer3, LM_F_F18_Layer5, LM_F_F18_Layer6, LM_F_F18_Mov3, LM_F_F18_PF, LM_F_F18_Parts, LM_F_F18_RFlip, LM_F_F18_RFlipU, LM_F_F18_RSling, LM_F_F18_RSling1, LM_F_F18_RSling2, LM_F_F18_SideRails, LM_F_F18_sw104, LM_F_F18_sw94)
Dim BL_F_F19: BL_F_F19=Array(LM_F_F19_LFlip, LM_F_F19_LFlip1, LM_F_F19_LFlip1U, LM_F_F19_LFlipU, LM_F_F19_LSling, LM_F_F19_LSling1, LM_F_F19_LSling2, LM_F_F19_PF, LM_F_F19_Parts, LM_F_F19_RFlip, LM_F_F19_RFlipU, LM_F_F19_RSling, LM_F_F19_RSling1, LM_F_F19_SideRails, LM_F_F19_Sling1, LM_F_F19_sw105, LM_F_F19_sw95)
Dim BL_F_F20: BL_F_F20=Array(LM_F_F20_Diverter2, LM_F_F20_LFlip1, LM_F_F20_LFlip1U, LM_F_F20_Layer3, LM_F_F20_PF, LM_F_F20_Parts, LM_F_F20_RRamp, LM_F_F20_RRampArm, LM_F_F20_RRampArmU, LM_F_F20_RRampU, LM_F_F20_RSling, LM_F_F20_RSling1, LM_F_F20_RSling2, LM_F_F20_sw104)
Dim BL_F_F21: BL_F_F21=Array(LM_F_F21_Layer4, LM_F_F21_Layer5, LM_F_F21_Mov3, LM_F_F21_PF, LM_F_F21_Parts, LM_F_F21_RRamp, LM_F_F21_RRampArm, LM_F_F21_RRampU, LM_F_F21_SideRails, LM_F_F21_SpinFlip2, LM_F_F21_WireDiverter1, LM_F_F21_sw92)
Dim BL_F_F22: BL_F_F22=Array(LM_F_F22_Layer3, LM_F_F22_Layer4, LM_F_F22_Layer5, LM_F_F22_Mov3, LM_F_F22_PF, LM_F_F22_Parts, LM_F_F22_SideRails, LM_F_F22_SpinFlip1, LM_F_F22_SpinFlip2, LM_F_F22_WireDiverter1)
Dim BL_F_F23: BL_F_F23=Array(LM_F_F23_LFlip, LM_F_F23_LFlip1U, LM_F_F23_LFlipU, LM_F_F23_LSling, LM_F_F23_PF, LM_F_F23_Parts, LM_F_F23_RFlip, LM_F_F23_RFlipU, LM_F_F23_RSling, LM_F_F23_RSling1, LM_F_F23_RSling2, LM_F_F23_Sling2, LM_F_F23_sw106, LM_F_F23_sw96)
Dim BL_F_L57: BL_F_L57=Array(LM_F_L57_Bumper1Ring, LM_F_L57_Layer3, LM_F_L57_Layer5, LM_F_L57_PF, LM_F_L57_Parts, LM_F_L57_SpinFlip1, LM_F_L57_SpinFlip4, LM_F_L57_sw102, LM_F_L57_sw103, LM_F_L57_sw112, LM_F_L57_sw113, LM_F_L57_sw92, LM_F_L57_sw93)
Dim BL_Gi: BL_Gi=Array(LM_Gi_Bumper1Ring, LM_Gi_Gate1, LM_Gi_LFlip, LM_Gi_LFlip1, LM_Gi_LFlip1U, LM_Gi_LFlipU, LM_Gi_LRamp_, LM_Gi_LRampArm, LM_Gi_LRampArmU, LM_Gi_LRampU, LM_Gi_Layer3, LM_Gi_Layer4, LM_Gi_Layer5, LM_Gi_Layer6, LM_Gi_Mov3, LM_Gi_PF, LM_Gi_Parts, LM_Gi_RFlip, LM_Gi_RFlipU, LM_Gi_RRamp, LM_Gi_RRampArm, LM_Gi_RRampArmU, LM_Gi_RRampU, LM_Gi_SideRails, LM_Gi_SpinFlip1, LM_Gi_SpinFlip2, LM_Gi_SpinFlip3, LM_Gi_SpinFlip4, LM_Gi_WireDiverter1, LM_Gi_sw102, LM_Gi_sw103, LM_Gi_sw104, LM_Gi_sw112, LM_Gi_sw113, LM_Gi_sw114, LM_Gi_sw20, LM_Gi_sw30, LM_Gi_sw32, LM_Gi_sw92, LM_Gi_sw93, LM_Gi_sw94)
Dim BL_Gi2_gi01: BL_Gi2_gi01=Array(LM_Gi2_gi01_LFlip, LM_Gi2_gi01_LFlip1U, LM_Gi2_gi01_LFlipU, LM_Gi2_gi01_LSling, LM_Gi2_gi01_LSling1, LM_Gi2_gi01_LSling2, LM_Gi2_gi01_PF, LM_Gi2_gi01_Parts, LM_Gi2_gi01_RFlip, LM_Gi2_gi01_RFlipU, LM_Gi2_gi01_Sling1)
Dim BL_Gi2_gi02: BL_Gi2_gi02=Array(LM_Gi2_gi02_LFlip, LM_Gi2_gi02_LFlip1, LM_Gi2_gi02_LFlip1U, LM_Gi2_gi02_LFlipU, LM_Gi2_gi02_LSling, LM_Gi2_gi02_LSling1, LM_Gi2_gi02_LSling2, LM_Gi2_gi02_PF, LM_Gi2_gi02_Parts, LM_Gi2_gi02_RFlip, LM_Gi2_gi02_RRamp, LM_Gi2_gi02_RSling, LM_Gi2_gi02_Sling1, LM_Gi2_gi02_sw105, LM_Gi2_gi02_sw95)
Dim BL_Gi2_gi03: BL_Gi2_gi03=Array(LM_Gi2_gi03_Diverter1, LM_Gi2_gi03_LFlip, LM_Gi2_gi03_LFlip1, LM_Gi2_gi03_LFlip1U, LM_Gi2_gi03_LFlipU, LM_Gi2_gi03_LSling, LM_Gi2_gi03_LSling1, LM_Gi2_gi03_LSling2, LM_Gi2_gi03_Layer3, LM_Gi2_gi03_PF, LM_Gi2_gi03_Parts, LM_Gi2_gi03_RFlip, LM_Gi2_gi03_RFlipU, LM_Gi2_gi03_sw105, LM_Gi2_gi03_sw95)
Dim BL_Gi2_gi11: BL_Gi2_gi11=Array(LM_Gi2_gi11_PF, LM_Gi2_gi11_Parts, LM_Gi2_gi11_RRamp, LM_Gi2_gi11_RRampU)
Dim BL_Gi2_gi12: BL_Gi2_gi12=Array(LM_Gi2_gi12_Diverter2, LM_Gi2_gi12_LFlip1U, LM_Gi2_gi12_PF, LM_Gi2_gi12_Parts, LM_Gi2_gi12_RRamp, LM_Gi2_gi12_RRampArm, LM_Gi2_gi12_RSling, LM_Gi2_gi12_RSling1, LM_Gi2_gi12_RSling2, LM_Gi2_gi12_sw106, LM_Gi2_gi12_sw96)
Dim BL_Gi2_gi13: BL_Gi2_gi13=Array(LM_Gi2_gi13_LFlip, LM_Gi2_gi13_LFlip1U, LM_Gi2_gi13_LSling, LM_Gi2_gi13_PF, LM_Gi2_gi13_Parts, LM_Gi2_gi13_RFlip, LM_Gi2_gi13_RFlipU, LM_Gi2_gi13_RSling, LM_Gi2_gi13_RSling1, LM_Gi2_gi13_RSling2, LM_Gi2_gi13_Sling2, LM_Gi2_gi13_sw103, LM_Gi2_gi13_sw106, LM_Gi2_gi13_sw96)
Dim BL_Gi2_gi14: BL_Gi2_gi14=Array(LM_Gi2_gi14_LFlip, LM_Gi2_gi14_LFlip1U, LM_Gi2_gi14_LFlipU, LM_Gi2_gi14_PF, LM_Gi2_gi14_Parts, LM_Gi2_gi14_RFlip, LM_Gi2_gi14_RFlipU, LM_Gi2_gi14_RSling, LM_Gi2_gi14_RSling1, LM_Gi2_gi14_RSling2, LM_Gi2_gi14_Sling2)
Dim BL_I_L0: BL_I_L0=Array(LM_I_L0_LFlip, LM_I_L0_LFlipU, LM_I_L0_PF, LM_I_L0_Parts, LM_I_L0_RFlip, LM_I_L0_RFlipU)
Dim BL_I_L10: BL_I_L10=Array(LM_I_L10_PF, LM_I_L10_Parts, LM_I_L10_sw95)
Dim BL_I_L11: BL_I_L11=Array(LM_I_L11_LSling, LM_I_L11_PF, LM_I_L11_Parts, LM_I_L11_sw105)
Dim BL_I_L12: BL_I_L12=Array(LM_I_L12_PF, LM_I_L12_Parts, LM_I_L12_RSling, LM_I_L12_RSling1, LM_I_L12_RSling2, LM_I_L12_sw106)
Dim BL_I_L13: BL_I_L13=Array(LM_I_L13_PF, LM_I_L13_Parts, LM_I_L13_sw96)
Dim BL_I_L14: BL_I_L14=Array(LM_I_L14_LFlip, LM_I_L14_LFlipU, LM_I_L14_PF, LM_I_L14_Parts, LM_I_L14_RFlip, LM_I_L14_RFlipU)
Dim BL_I_L15: BL_I_L15=Array(LM_I_L15_LSling, LM_I_L15_LSling1, LM_I_L15_LSling2, LM_I_L15_PF, LM_I_L15_Parts)
Dim BL_I_L16: BL_I_L16=Array(LM_I_L16_PF, LM_I_L16_Parts)
Dim BL_I_L17: BL_I_L17=Array(LM_I_L17_PF, LM_I_L17_Parts)
Dim BL_I_L20: BL_I_L20=Array(LM_I_L20_PF, LM_I_L20_Parts, LM_I_L20_RSling, LM_I_L20_RSling1, LM_I_L20_RSling2, LM_I_L20_Sling2)
Dim BL_I_L21: BL_I_L21=Array(LM_I_L21_PF, LM_I_L21_Parts, LM_I_L21_RSling, LM_I_L21_RSling1, LM_I_L21_RSling2)
Dim BL_I_L22: BL_I_L22=Array(LM_I_L22_PF, LM_I_L22_Parts)
Dim BL_I_L23: BL_I_L23=Array(LM_I_L23_PF, LM_I_L23_Parts)
Dim BL_I_L24: BL_I_L24=Array(LM_I_L24_PF, LM_I_L24_Parts)
Dim BL_I_L25: BL_I_L25=Array(LM_I_L25_PF, LM_I_L25_Parts)
Dim BL_I_L26: BL_I_L26=Array(LM_I_L26_PF, LM_I_L26_Parts)
Dim BL_I_L27: BL_I_L27=Array(LM_I_L27_PF, LM_I_L27_Parts, LM_I_L27_RRampArm)
Dim BL_I_L30: BL_I_L30=Array(LM_I_L30_LPF, LM_I_L30_PF, LM_I_L30_Parts, LM_I_L30_sw104, LM_I_L30_sw94)
Dim BL_I_L31: BL_I_L31=Array(LM_I_L31_Car, LM_I_L31_LPF, LM_I_L31_PF, LM_I_L31_Parts, LM_I_L31_sw104)
Dim BL_I_L32: BL_I_L32=Array(LM_I_L32_PF, LM_I_L32_Parts)
Dim BL_I_L33: BL_I_L33=Array(LM_I_L33_PF, LM_I_L33_Parts, LM_I_L33_sw103, LM_I_L33_sw113)
Dim BL_I_L34: BL_I_L34=Array(LM_I_L34_Layer3, LM_I_L34_PF, LM_I_L34_Parts, LM_I_L34_sw103, LM_I_L34_sw113)
Dim BL_I_L35: BL_I_L35=Array(LM_I_L35_Bumper1Ring, LM_I_L35_Layer3, LM_I_L35_Layer5, LM_I_L35_PF, LM_I_L35_Parts, LM_I_L35_sw103, LM_I_L35_sw112, LM_I_L35_sw113, LM_I_L35_sw30)
Dim BL_I_L36: BL_I_L36=Array(LM_I_L36_Bumper1Ring, LM_I_L36_Layer3, LM_I_L36_Layer5, LM_I_L36_PF, LM_I_L36_Parts, LM_I_L36_sw102, LM_I_L36_sw112, LM_I_L36_sw92)
Dim BL_I_L37: BL_I_L37=Array(LM_I_L37_Bumper1Ring, LM_I_L37_Layer3, LM_I_L37_Layer5, LM_I_L37_PF, LM_I_L37_Parts, LM_I_L37_sw112, LM_I_L37_sw92)
Dim BL_I_L40: BL_I_L40=Array(LM_I_L40_LFlip1U, LM_I_L40_PF, LM_I_L40_Parts)
Dim BL_I_L41: BL_I_L41=Array(LM_I_L41_LRampArm, LM_I_L41_LRampU, LM_I_L41_PF, LM_I_L41_Parts)
Dim BL_I_L42: BL_I_L42=Array(LM_I_L42_LRampArm, LM_I_L42_LRampU, LM_I_L42_Layer3, LM_I_L42_PF, LM_I_L42_Parts)
Dim BL_I_L43: BL_I_L43=Array(LM_I_L43_LRamp_, LM_I_L43_LRampU, LM_I_L43_Layer5, LM_I_L43_PF, LM_I_L43_Parts, LM_I_L43_sw103, LM_I_L43_sw93)
Dim BL_I_L44: BL_I_L44=Array(LM_I_L44_LRampU, LM_I_L44_Layer5, LM_I_L44_PF, LM_I_L44_Parts, LM_I_L44_sw92, LM_I_L44_sw93)
Dim BL_I_L45: BL_I_L45=Array(LM_I_L45_Bumper1Ring, LM_I_L45_PF, LM_I_L45_Parts, LM_I_L45_sw92)
Dim BL_I_L46: BL_I_L46=Array(LM_I_L46_Car, LM_I_L46_LPF, LM_I_L46_PF, LM_I_L46_Parts, LM_I_L46_sw103, LM_I_L46_sw113, LM_I_L46_sw93)
Dim BL_I_L47: BL_I_L47=Array(LM_I_L47_Layer3, LM_I_L47_PF, LM_I_L47_Parts, LM_I_L47_sw103, LM_I_L47_sw113)
Dim BL_I_L50: BL_I_L50=Array(LM_I_L50_LFlip1U, LM_I_L50_PF)
Dim BL_I_L51: BL_I_L51=Array(LM_I_L51_LFlip1U, LM_I_L51_PF)
Dim BL_I_L52: BL_I_L52=Array(LM_I_L52_LFlip1U, LM_I_L52_PF, LM_I_L52_Parts)
Dim BL_I_L53: BL_I_L53=Array(LM_I_L53_LRampArmU, LM_I_L53_LRampU, LM_I_L53_PF, LM_I_L53_Parts)
Dim BL_I_L54: BL_I_L54=Array(LM_I_L54_PF, LM_I_L54_Parts, LM_I_L54_sw102, LM_I_L54_sw92)
Dim BL_I_L55: BL_I_L55=Array(LM_I_L55_Layer5, LM_I_L55_PF, LM_I_L55_Parts, LM_I_L55_sw102, LM_I_L55_sw112, LM_I_L55_sw92)
Dim BL_I_L56: BL_I_L56=Array(LM_I_L56_Layer5, LM_I_L56_PF, LM_I_L56_Parts, LM_I_L56_sw102, LM_I_L56_sw112, LM_I_L56_sw92)
Dim BL_I_L60: BL_I_L60=Array(LM_I_L60_Car, LM_I_L60_LPF, LM_I_L60_Layer1, LM_I_L60_Layer2, LM_I_L60_Parts, LM_I_L60_RFlip1, LM_I_L60_RFlip1U)
Dim BL_Room: BL_Room=Array(BM_Bumper1Ring, BM_Car, BM_Diverter1, BM_Diverter2, BM_Gate1, BM_LFlip, BM_LFlip1, BM_LFlip1U, BM_LFlipU, BM_LPF, BM_LRamp_, BM_LRampArm, BM_LRampArmU, BM_LRampU, BM_LSling, BM_LSling1, BM_LSling2, BM_Layer1, BM_Layer2, BM_Layer3, BM_Layer4, BM_Layer5, BM_Layer6, BM_Mov3, BM_PF, BM_Parts, BM_Plunger, BM_RFlip, BM_RFlip1, BM_RFlip1U, BM_RFlipU, BM_RRamp, BM_RRampArm, BM_RRampArmU, BM_RRampU, BM_RSling, BM_RSling1, BM_RSling2, BM_SideRails, BM_Sling1, BM_Sling2, BM_SpinFlip1, BM_SpinFlip2, BM_SpinFlip3, BM_SpinFlip4, BM_WireDiverter1, BM_sw102, BM_sw103, BM_sw104, BM_sw105, BM_sw106, BM_sw112, BM_sw113, BM_sw114, BM_sw115, BM_sw116, BM_sw20, BM_sw22, BM_sw30, BM_sw32, BM_sw92, BM_sw93, BM_sw94, BM_sw95, BM_sw96)
'' Global arrays
'Dim BG_Bakemap: BG_Bakemap=Array(BM_Bumper1Ring, BM_Car, BM_Diverter1, BM_Diverter2, BM_Gate1, BM_LFlip, BM_LFlip1, BM_LFlip1U, BM_LFlipU, BM_LPF, BM_LRamp_, BM_LRampArm, BM_LRampArmU, BM_LRampU, BM_LSling, BM_LSling1, BM_LSling2, BM_Layer1, BM_Layer2, BM_Layer3, BM_Layer4, BM_Layer5, BM_Layer6, BM_Mov3, BM_PF, BM_Parts, BM_Plunger, BM_RFlip, BM_RFlip1, BM_RFlip1U, BM_RFlipU, BM_RRamp, BM_RRampArm, BM_RRampArmU, BM_RRampU, BM_RSling, BM_RSling1, BM_RSling2, BM_SideRails, BM_Sling1, BM_Sling2, BM_SpinFlip1, BM_SpinFlip2, BM_SpinFlip3, BM_SpinFlip4, BM_WireDiverter1, BM_sw102, BM_sw103, BM_sw104, BM_sw105, BM_sw106, BM_sw112, BM_sw113, BM_sw114, BM_sw115, BM_sw116, BM_sw20, BM_sw22, BM_sw30, BM_sw32, BM_sw92, BM_sw93, BM_sw94, BM_sw95, BM_sw96)
'Dim BG_Lightmap: BG_Lightmap=Array(LM_A_L61_Parts, LM_A_L62_Parts, LM_A_L63_Parts, LM_A_L64_Parts, LM_A_L65_Parts, LM_A_L66_Parts, LM_A_L67_Parts, LM_CC_L02_Car, LM_CC_L02_LPF, LM_CC_L02_Layer1, LM_CC_L02_Layer2, LM_CC_L02_PF, LM_CC_L02_Parts, LM_CC_L02_RFlip1, LM_CC_L02_RFlip1U, LM_CC_L02_Sling1, LM_CC_L03_Car, LM_CC_L03_LPF, LM_CC_L03_Layer1, LM_CC_L03_Layer2, LM_CC_L03_PF, LM_CC_L03_Parts, LM_CC_L03_RFlip1, LM_CC_L03_RFlip1U, LM_CC_L03_sw105, LM_CC_L04_Car, LM_CC_L04_LPF, LM_CC_L04_Layer1, LM_CC_L04_Layer2, LM_CC_L04_PF, LM_CC_L04_Parts, LM_CC_L04_RFlip1, LM_CC_L04_RFlip1U, LM_CC_L05_Car, LM_CC_L05_LPF, LM_CC_L05_Layer1, LM_CC_L05_Layer2, LM_CC_L05_PF, LM_CC_L05_Parts, LM_CC_L05_RFlip1, LM_CC_L05_RFlip1U, LM_CC_L05_Sling2, LM_CC_L06_Car, LM_CC_L06_LPF, LM_CC_L06_Layer1, LM_CC_L06_Layer2, LM_CC_L06_PF, LM_CC_L06_Parts, LM_CC_L06_RFlip1, LM_CC_L06_RFlip1U, LM_CC_L06_sw106, LM_CC_L07_Car, LM_CC_L07_LPF, LM_CC_L07_Layer1, LM_CC_L07_Layer2, LM_CC_L07_PF, LM_CC_L07_Parts, LM_CC_L07_RFlip1U, LM_F_F15_Bumper1Ring, _
' LM_F_F15_Layer3, LM_F_F15_Layer5, LM_F_F15_Layer6, LM_F_F15_Mov3, LM_F_F15_PF, LM_F_F15_Parts, LM_F_F15_RFlip, LM_F_F15_RFlipU, LM_F_F15_SideRails, LM_F_F15_SpinFlip1, LM_F_F15_SpinFlip2, LM_F_F15_SpinFlip3, LM_F_F15_SpinFlip4, LM_F_F16_Gate1, LM_F_F16_LFlip1, LM_F_F16_LFlip1U, LM_F_F16_Layer3, LM_F_F16_Layer5, LM_F_F16_Layer6, LM_F_F16_Mov3, LM_F_F16_PF, LM_F_F16_Parts, LM_F_F16_SideRails, LM_F_F16_SpinFlip1, LM_F_F16_SpinFlip2, LM_F_F16_SpinFlip3, LM_F_F16_SpinFlip4, LM_F_F17_Bumper1Ring, LM_F_F17_Diverter1, LM_F_F17_LFlip, LM_F_F17_LFlip1, LM_F_F17_LFlip1U, LM_F_F17_LFlipU, LM_F_F17_Layer3, LM_F_F17_PF, LM_F_F17_Parts, LM_F_F17_RSling, LM_F_F17_RSling2, LM_F_F17_sw103, LM_F_F17_sw104, LM_F_F17_sw105, LM_F_F17_sw113, LM_F_F17_sw93, LM_F_F17_sw95, LM_F_F18_LFlip, LM_F_F18_LFlip1, LM_F_F18_LFlip1U, LM_F_F18_LFlipU, LM_F_F18_LRamp_, LM_F_F18_LRampArm, LM_F_F18_LRampArmU, LM_F_F18_LRampU, LM_F_F18_Layer3, LM_F_F18_Layer5, LM_F_F18_Layer6, LM_F_F18_Mov3, LM_F_F18_PF, LM_F_F18_Parts, LM_F_F18_RFlip, _
' LM_F_F18_RFlipU, LM_F_F18_RSling, LM_F_F18_RSling1, LM_F_F18_RSling2, LM_F_F18_SideRails, LM_F_F18_sw104, LM_F_F18_sw94, LM_F_F19_LFlip, LM_F_F19_LFlip1, LM_F_F19_LFlip1U, LM_F_F19_LFlipU, LM_F_F19_LSling, LM_F_F19_LSling1, LM_F_F19_LSling2, LM_F_F19_PF, LM_F_F19_Parts, LM_F_F19_RFlip, LM_F_F19_RFlipU, LM_F_F19_RSling, LM_F_F19_RSling1, LM_F_F19_SideRails, LM_F_F19_Sling1, LM_F_F19_sw105, LM_F_F19_sw95, LM_F_F20_Diverter2, LM_F_F20_LFlip1, LM_F_F20_LFlip1U, LM_F_F20_Layer3, LM_F_F20_PF, LM_F_F20_Parts, LM_F_F20_RRamp, LM_F_F20_RRampArm, LM_F_F20_RRampArmU, LM_F_F20_RRampU, LM_F_F20_RSling, LM_F_F20_RSling1, LM_F_F20_RSling2, LM_F_F20_sw104, LM_F_F21_Layer4, LM_F_F21_Layer5, LM_F_F21_Mov3, LM_F_F21_PF, LM_F_F21_Parts, LM_F_F21_RRamp, LM_F_F21_RRampArm, LM_F_F21_RRampU, LM_F_F21_SideRails, LM_F_F21_SpinFlip2, LM_F_F21_WireDiverter1, LM_F_F21_sw92, LM_F_F22_Layer3, LM_F_F22_Layer4, LM_F_F22_Layer5, LM_F_F22_Mov3, LM_F_F22_PF, LM_F_F22_Parts, LM_F_F22_SideRails, LM_F_F22_SpinFlip1, LM_F_F22_SpinFlip2, _
' LM_F_F22_WireDiverter1, LM_F_F23_LFlip, LM_F_F23_LFlip1U, LM_F_F23_LFlipU, LM_F_F23_LSling, LM_F_F23_PF, LM_F_F23_Parts, LM_F_F23_RFlip, LM_F_F23_RFlipU, LM_F_F23_RSling, LM_F_F23_RSling1, LM_F_F23_RSling2, LM_F_F23_Sling2, LM_F_F23_sw106, LM_F_F23_sw96, LM_F_L57_Bumper1Ring, LM_F_L57_Layer3, LM_F_L57_Layer5, LM_F_L57_PF, LM_F_L57_Parts, LM_F_L57_SpinFlip1, LM_F_L57_SpinFlip4, LM_F_L57_sw102, LM_F_L57_sw103, LM_F_L57_sw112, LM_F_L57_sw113, LM_F_L57_sw92, LM_F_L57_sw93, LM_Gi_Bumper1Ring, LM_Gi_Gate1, LM_Gi_LFlip, LM_Gi_LFlip1, LM_Gi_LFlip1U, LM_Gi_LFlipU, LM_Gi_LRamp_, LM_Gi_LRampArm, LM_Gi_LRampArmU, LM_Gi_LRampU, LM_Gi_Layer3, LM_Gi_Layer4, LM_Gi_Layer5, LM_Gi_Layer6, LM_Gi_Mov3, LM_Gi_PF, LM_Gi_Parts, LM_Gi_RFlip, LM_Gi_RFlipU, LM_Gi_RRamp, LM_Gi_RRampArm, LM_Gi_RRampArmU, LM_Gi_RRampU, LM_Gi_SideRails, LM_Gi_SpinFlip1, LM_Gi_SpinFlip2, LM_Gi_SpinFlip3, LM_Gi_SpinFlip4, LM_Gi_WireDiverter1, LM_Gi_sw102, LM_Gi_sw103, LM_Gi_sw104, LM_Gi_sw112, LM_Gi_sw113, LM_Gi_sw114, LM_Gi_sw20, LM_Gi_sw30, LM_Gi_sw32, _
' LM_Gi_sw92, LM_Gi_sw93, LM_Gi_sw94, LM_Gi2_gi01_LFlip, LM_Gi2_gi01_LFlip1U, LM_Gi2_gi01_LFlipU, LM_Gi2_gi01_LSling, LM_Gi2_gi01_LSling1, LM_Gi2_gi01_LSling2, LM_Gi2_gi01_PF, LM_Gi2_gi01_Parts, LM_Gi2_gi01_RFlip, LM_Gi2_gi01_RFlipU, LM_Gi2_gi01_Sling1, LM_Gi2_gi02_LFlip, LM_Gi2_gi02_LFlip1, LM_Gi2_gi02_LFlip1U, LM_Gi2_gi02_LFlipU, LM_Gi2_gi02_LSling, LM_Gi2_gi02_LSling1, LM_Gi2_gi02_LSling2, LM_Gi2_gi02_PF, LM_Gi2_gi02_Parts, LM_Gi2_gi02_RFlip, LM_Gi2_gi02_RRamp, LM_Gi2_gi02_RSling, LM_Gi2_gi02_Sling1, LM_Gi2_gi02_sw105, LM_Gi2_gi02_sw95, LM_Gi2_gi03_Diverter1, LM_Gi2_gi03_LFlip, LM_Gi2_gi03_LFlip1, LM_Gi2_gi03_LFlip1U, LM_Gi2_gi03_LFlipU, LM_Gi2_gi03_LSling, LM_Gi2_gi03_LSling1, LM_Gi2_gi03_LSling2, LM_Gi2_gi03_Layer3, LM_Gi2_gi03_PF, LM_Gi2_gi03_Parts, LM_Gi2_gi03_RFlip, LM_Gi2_gi03_RFlipU, LM_Gi2_gi03_sw105, LM_Gi2_gi03_sw95, LM_Gi2_gi11_PF, LM_Gi2_gi11_Parts, LM_Gi2_gi11_RRamp, LM_Gi2_gi11_RRampU, LM_Gi2_gi12_Diverter2, LM_Gi2_gi12_LFlip1U, LM_Gi2_gi12_PF, LM_Gi2_gi12_Parts, LM_Gi2_gi12_RRamp, _
' LM_Gi2_gi12_RRampArm, LM_Gi2_gi12_RSling, LM_Gi2_gi12_RSling1, LM_Gi2_gi12_RSling2, LM_Gi2_gi12_sw106, LM_Gi2_gi12_sw96, LM_Gi2_gi13_LFlip, LM_Gi2_gi13_LFlip1U, LM_Gi2_gi13_LSling, LM_Gi2_gi13_PF, LM_Gi2_gi13_Parts, LM_Gi2_gi13_RFlip, LM_Gi2_gi13_RFlipU, LM_Gi2_gi13_RSling, LM_Gi2_gi13_RSling1, LM_Gi2_gi13_RSling2, LM_Gi2_gi13_Sling2, LM_Gi2_gi13_sw103, LM_Gi2_gi13_sw106, LM_Gi2_gi13_sw96, LM_Gi2_gi14_LFlip, LM_Gi2_gi14_LFlip1U, LM_Gi2_gi14_LFlipU, LM_Gi2_gi14_PF, LM_Gi2_gi14_Parts, LM_Gi2_gi14_RFlip, LM_Gi2_gi14_RFlipU, LM_Gi2_gi14_RSling, LM_Gi2_gi14_RSling1, LM_Gi2_gi14_RSling2, LM_Gi2_gi14_Sling2, LM_I_L0_LFlip, LM_I_L0_LFlipU, LM_I_L0_PF, LM_I_L0_Parts, LM_I_L0_RFlip, LM_I_L0_RFlipU, LM_I_L10_PF, LM_I_L10_Parts, LM_I_L10_sw95, LM_I_L11_LSling, LM_I_L11_PF, LM_I_L11_Parts, LM_I_L11_sw105, LM_I_L12_PF, LM_I_L12_Parts, LM_I_L12_RSling, LM_I_L12_RSling1, LM_I_L12_RSling2, LM_I_L12_sw106, LM_I_L13_PF, LM_I_L13_Parts, LM_I_L13_sw96, LM_I_L14_LFlip, LM_I_L14_LFlipU, LM_I_L14_PF, LM_I_L14_Parts, LM_I_L14_RFlip, _
' LM_I_L14_RFlipU, LM_I_L15_LSling, LM_I_L15_LSling1, LM_I_L15_LSling2, LM_I_L15_PF, LM_I_L15_Parts, LM_I_L16_PF, LM_I_L16_Parts, LM_I_L17_PF, LM_I_L17_Parts, LM_I_L20_PF, LM_I_L20_Parts, LM_I_L20_RSling, LM_I_L20_RSling1, LM_I_L20_RSling2, LM_I_L20_Sling2, LM_I_L21_PF, LM_I_L21_Parts, LM_I_L21_RSling, LM_I_L21_RSling1, LM_I_L21_RSling2, LM_I_L22_PF, LM_I_L22_Parts, LM_I_L23_PF, LM_I_L23_Parts, LM_I_L24_PF, LM_I_L24_Parts, LM_I_L25_PF, LM_I_L25_Parts, LM_I_L26_PF, LM_I_L26_Parts, LM_I_L27_PF, LM_I_L27_Parts, LM_I_L27_RRampArm, LM_I_L30_LPF, LM_I_L30_PF, LM_I_L30_Parts, LM_I_L30_sw104, LM_I_L30_sw94, LM_I_L31_Car, LM_I_L31_LPF, LM_I_L31_PF, LM_I_L31_Parts, LM_I_L31_sw104, LM_I_L32_PF, LM_I_L32_Parts, LM_I_L33_PF, LM_I_L33_Parts, LM_I_L33_sw103, LM_I_L33_sw113, LM_I_L34_Layer3, LM_I_L34_PF, LM_I_L34_Parts, LM_I_L34_sw103, LM_I_L34_sw113, LM_I_L35_Bumper1Ring, LM_I_L35_Layer3, LM_I_L35_Layer5, LM_I_L35_PF, LM_I_L35_Parts, LM_I_L35_sw103, LM_I_L35_sw112, LM_I_L35_sw113, LM_I_L35_sw30, LM_I_L36_Bumper1Ring, _
' LM_I_L36_Layer3, LM_I_L36_Layer5, LM_I_L36_PF, LM_I_L36_Parts, LM_I_L36_sw102, LM_I_L36_sw112, LM_I_L36_sw92, LM_I_L37_Bumper1Ring, LM_I_L37_Layer3, LM_I_L37_Layer5, LM_I_L37_PF, LM_I_L37_Parts, LM_I_L37_sw112, LM_I_L37_sw92, LM_I_L40_LFlip1U, LM_I_L40_PF, LM_I_L40_Parts, LM_I_L41_LRampArm, LM_I_L41_LRampU, LM_I_L41_PF, LM_I_L41_Parts, LM_I_L42_LRampArm, LM_I_L42_LRampU, LM_I_L42_Layer3, LM_I_L42_PF, LM_I_L42_Parts, LM_I_L43_LRamp_, LM_I_L43_LRampU, LM_I_L43_Layer5, LM_I_L43_PF, LM_I_L43_Parts, LM_I_L43_sw103, LM_I_L43_sw93, LM_I_L44_LRampU, LM_I_L44_Layer5, LM_I_L44_PF, LM_I_L44_Parts, LM_I_L44_sw92, LM_I_L44_sw93, LM_I_L45_Bumper1Ring, LM_I_L45_PF, LM_I_L45_Parts, LM_I_L45_sw92, LM_I_L46_Car, LM_I_L46_LPF, LM_I_L46_PF, LM_I_L46_Parts, LM_I_L46_sw103, LM_I_L46_sw113, LM_I_L46_sw93, LM_I_L47_Layer3, LM_I_L47_PF, LM_I_L47_Parts, LM_I_L47_sw103, LM_I_L47_sw113, LM_I_L50_LFlip1U, LM_I_L50_PF, LM_I_L51_LFlip1U, LM_I_L51_PF, LM_I_L52_LFlip1U, LM_I_L52_PF, LM_I_L52_Parts, LM_I_L53_LRampArmU, LM_I_L53_LRampU, _
' LM_I_L53_PF, LM_I_L53_Parts, LM_I_L54_PF, LM_I_L54_Parts, LM_I_L54_sw102, LM_I_L54_sw92, LM_I_L55_Layer5, LM_I_L55_PF, LM_I_L55_Parts, LM_I_L55_sw102, LM_I_L55_sw112, LM_I_L55_sw92, LM_I_L56_Layer5, LM_I_L56_PF, LM_I_L56_Parts, LM_I_L56_sw102, LM_I_L56_sw112, LM_I_L56_sw92, LM_I_L60_Car, LM_I_L60_LPF, LM_I_L60_Layer1, LM_I_L60_Layer2, LM_I_L60_Parts, LM_I_L60_RFlip1, LM_I_L60_RFlip1U)
'Dim BG_All: BG_All=Array(BM_Bumper1Ring, BM_Car, BM_Diverter1, BM_Diverter2, BM_Gate1, BM_LFlip, BM_LFlip1, BM_LFlip1U, BM_LFlipU, BM_LPF, BM_LRamp_, BM_LRampArm, BM_LRampArmU, BM_LRampU, BM_LSling, BM_LSling1, BM_LSling2, BM_Layer1, BM_Layer2, BM_Layer3, BM_Layer4, BM_Layer5, BM_Layer6, BM_Mov3, BM_PF, BM_Parts, BM_Plunger, BM_RFlip, BM_RFlip1, BM_RFlip1U, BM_RFlipU, BM_RRamp, BM_RRampArm, BM_RRampArmU, BM_RRampU, BM_RSling, BM_RSling1, BM_RSling2, BM_SideRails, BM_Sling1, BM_Sling2, BM_SpinFlip1, BM_SpinFlip2, BM_SpinFlip3, BM_SpinFlip4, BM_WireDiverter1, BM_sw102, BM_sw103, BM_sw104, BM_sw105, BM_sw106, BM_sw112, BM_sw113, BM_sw114, BM_sw115, BM_sw116, BM_sw20, BM_sw22, BM_sw30, BM_sw32, BM_sw92, BM_sw93, BM_sw94, BM_sw95, BM_sw96, LM_A_L61_Parts, LM_A_L62_Parts, LM_A_L63_Parts, LM_A_L64_Parts, LM_A_L65_Parts, LM_A_L66_Parts, LM_A_L67_Parts, LM_CC_L02_Car, LM_CC_L02_LPF, LM_CC_L02_Layer1, LM_CC_L02_Layer2, LM_CC_L02_PF, LM_CC_L02_Parts, LM_CC_L02_RFlip1, LM_CC_L02_RFlip1U, LM_CC_L02_Sling1, LM_CC_L03_Car, _
' LM_CC_L03_LPF, LM_CC_L03_Layer1, LM_CC_L03_Layer2, LM_CC_L03_PF, LM_CC_L03_Parts, LM_CC_L03_RFlip1, LM_CC_L03_RFlip1U, LM_CC_L03_sw105, LM_CC_L04_Car, LM_CC_L04_LPF, LM_CC_L04_Layer1, LM_CC_L04_Layer2, LM_CC_L04_PF, LM_CC_L04_Parts, LM_CC_L04_RFlip1, LM_CC_L04_RFlip1U, LM_CC_L05_Car, LM_CC_L05_LPF, LM_CC_L05_Layer1, LM_CC_L05_Layer2, LM_CC_L05_PF, LM_CC_L05_Parts, LM_CC_L05_RFlip1, LM_CC_L05_RFlip1U, LM_CC_L05_Sling2, LM_CC_L06_Car, LM_CC_L06_LPF, LM_CC_L06_Layer1, LM_CC_L06_Layer2, LM_CC_L06_PF, LM_CC_L06_Parts, LM_CC_L06_RFlip1, LM_CC_L06_RFlip1U, LM_CC_L06_sw106, LM_CC_L07_Car, LM_CC_L07_LPF, LM_CC_L07_Layer1, LM_CC_L07_Layer2, LM_CC_L07_PF, LM_CC_L07_Parts, LM_CC_L07_RFlip1U, LM_F_F15_Bumper1Ring, LM_F_F15_Layer3, LM_F_F15_Layer5, LM_F_F15_Layer6, LM_F_F15_Mov3, LM_F_F15_PF, LM_F_F15_Parts, LM_F_F15_RFlip, LM_F_F15_RFlipU, LM_F_F15_SideRails, LM_F_F15_SpinFlip1, LM_F_F15_SpinFlip2, LM_F_F15_SpinFlip3, LM_F_F15_SpinFlip4, LM_F_F16_Gate1, LM_F_F16_LFlip1, LM_F_F16_LFlip1U, LM_F_F16_Layer3, LM_F_F16_Layer5, _
' LM_F_F16_Layer6, LM_F_F16_Mov3, LM_F_F16_PF, LM_F_F16_Parts, LM_F_F16_SideRails, LM_F_F16_SpinFlip1, LM_F_F16_SpinFlip2, LM_F_F16_SpinFlip3, LM_F_F16_SpinFlip4, LM_F_F17_Bumper1Ring, LM_F_F17_Diverter1, LM_F_F17_LFlip, LM_F_F17_LFlip1, LM_F_F17_LFlip1U, LM_F_F17_LFlipU, LM_F_F17_Layer3, LM_F_F17_PF, LM_F_F17_Parts, LM_F_F17_RSling, LM_F_F17_RSling2, LM_F_F17_sw103, LM_F_F17_sw104, LM_F_F17_sw105, LM_F_F17_sw113, LM_F_F17_sw93, LM_F_F17_sw95, LM_F_F18_LFlip, LM_F_F18_LFlip1, LM_F_F18_LFlip1U, LM_F_F18_LFlipU, LM_F_F18_LRamp_, LM_F_F18_LRampArm, LM_F_F18_LRampArmU, LM_F_F18_LRampU, LM_F_F18_Layer3, LM_F_F18_Layer5, LM_F_F18_Layer6, LM_F_F18_Mov3, LM_F_F18_PF, LM_F_F18_Parts, LM_F_F18_RFlip, LM_F_F18_RFlipU, LM_F_F18_RSling, LM_F_F18_RSling1, LM_F_F18_RSling2, LM_F_F18_SideRails, LM_F_F18_sw104, LM_F_F18_sw94, LM_F_F19_LFlip, LM_F_F19_LFlip1, LM_F_F19_LFlip1U, LM_F_F19_LFlipU, LM_F_F19_LSling, LM_F_F19_LSling1, LM_F_F19_LSling2, LM_F_F19_PF, LM_F_F19_Parts, LM_F_F19_RFlip, LM_F_F19_RFlipU, LM_F_F19_RSling, _
' LM_F_F19_RSling1, LM_F_F19_SideRails, LM_F_F19_Sling1, LM_F_F19_sw105, LM_F_F19_sw95, LM_F_F20_Diverter2, LM_F_F20_LFlip1, LM_F_F20_LFlip1U, LM_F_F20_Layer3, LM_F_F20_PF, LM_F_F20_Parts, LM_F_F20_RRamp, LM_F_F20_RRampArm, LM_F_F20_RRampArmU, LM_F_F20_RRampU, LM_F_F20_RSling, LM_F_F20_RSling1, LM_F_F20_RSling2, LM_F_F20_sw104, LM_F_F21_Layer4, LM_F_F21_Layer5, LM_F_F21_Mov3, LM_F_F21_PF, LM_F_F21_Parts, LM_F_F21_RRamp, LM_F_F21_RRampArm, LM_F_F21_RRampU, LM_F_F21_SideRails, LM_F_F21_SpinFlip2, LM_F_F21_WireDiverter1, LM_F_F21_sw92, LM_F_F22_Layer3, LM_F_F22_Layer4, LM_F_F22_Layer5, LM_F_F22_Mov3, LM_F_F22_PF, LM_F_F22_Parts, LM_F_F22_SideRails, LM_F_F22_SpinFlip1, LM_F_F22_SpinFlip2, LM_F_F22_WireDiverter1, LM_F_F23_LFlip, LM_F_F23_LFlip1U, LM_F_F23_LFlipU, LM_F_F23_LSling, LM_F_F23_PF, LM_F_F23_Parts, LM_F_F23_RFlip, LM_F_F23_RFlipU, LM_F_F23_RSling, LM_F_F23_RSling1, LM_F_F23_RSling2, LM_F_F23_Sling2, LM_F_F23_sw106, LM_F_F23_sw96, LM_F_L57_Bumper1Ring, LM_F_L57_Layer3, LM_F_L57_Layer5, LM_F_L57_PF, _
' LM_F_L57_Parts, LM_F_L57_SpinFlip1, LM_F_L57_SpinFlip4, LM_F_L57_sw102, LM_F_L57_sw103, LM_F_L57_sw112, LM_F_L57_sw113, LM_F_L57_sw92, LM_F_L57_sw93, LM_Gi_Bumper1Ring, LM_Gi_Gate1, LM_Gi_LFlip, LM_Gi_LFlip1, LM_Gi_LFlip1U, LM_Gi_LFlipU, LM_Gi_LRamp_, LM_Gi_LRampArm, LM_Gi_LRampArmU, LM_Gi_LRampU, LM_Gi_Layer3, LM_Gi_Layer4, LM_Gi_Layer5, LM_Gi_Layer6, LM_Gi_Mov3, LM_Gi_PF, LM_Gi_Parts, LM_Gi_RFlip, LM_Gi_RFlipU, LM_Gi_RRamp, LM_Gi_RRampArm, LM_Gi_RRampArmU, LM_Gi_RRampU, LM_Gi_SideRails, LM_Gi_SpinFlip1, LM_Gi_SpinFlip2, LM_Gi_SpinFlip3, LM_Gi_SpinFlip4, LM_Gi_WireDiverter1, LM_Gi_sw102, LM_Gi_sw103, LM_Gi_sw104, LM_Gi_sw112, LM_Gi_sw113, LM_Gi_sw114, LM_Gi_sw20, LM_Gi_sw30, LM_Gi_sw32, LM_Gi_sw92, LM_Gi_sw93, LM_Gi_sw94, LM_Gi2_gi01_LFlip, LM_Gi2_gi01_LFlip1U, LM_Gi2_gi01_LFlipU, LM_Gi2_gi01_LSling, LM_Gi2_gi01_LSling1, LM_Gi2_gi01_LSling2, LM_Gi2_gi01_PF, LM_Gi2_gi01_Parts, LM_Gi2_gi01_RFlip, LM_Gi2_gi01_RFlipU, LM_Gi2_gi01_Sling1, LM_Gi2_gi02_LFlip, LM_Gi2_gi02_LFlip1, LM_Gi2_gi02_LFlip1U, _
' LM_Gi2_gi02_LFlipU, LM_Gi2_gi02_LSling, LM_Gi2_gi02_LSling1, LM_Gi2_gi02_LSling2, LM_Gi2_gi02_PF, LM_Gi2_gi02_Parts, LM_Gi2_gi02_RFlip, LM_Gi2_gi02_RRamp, LM_Gi2_gi02_RSling, LM_Gi2_gi02_Sling1, LM_Gi2_gi02_sw105, LM_Gi2_gi02_sw95, LM_Gi2_gi03_Diverter1, LM_Gi2_gi03_LFlip, LM_Gi2_gi03_LFlip1, LM_Gi2_gi03_LFlip1U, LM_Gi2_gi03_LFlipU, LM_Gi2_gi03_LSling, LM_Gi2_gi03_LSling1, LM_Gi2_gi03_LSling2, LM_Gi2_gi03_Layer3, LM_Gi2_gi03_PF, LM_Gi2_gi03_Parts, LM_Gi2_gi03_RFlip, LM_Gi2_gi03_RFlipU, LM_Gi2_gi03_sw105, LM_Gi2_gi03_sw95, LM_Gi2_gi11_PF, LM_Gi2_gi11_Parts, LM_Gi2_gi11_RRamp, LM_Gi2_gi11_RRampU, LM_Gi2_gi12_Diverter2, LM_Gi2_gi12_LFlip1U, LM_Gi2_gi12_PF, LM_Gi2_gi12_Parts, LM_Gi2_gi12_RRamp, LM_Gi2_gi12_RRampArm, LM_Gi2_gi12_RSling, LM_Gi2_gi12_RSling1, LM_Gi2_gi12_RSling2, LM_Gi2_gi12_sw106, LM_Gi2_gi12_sw96, LM_Gi2_gi13_LFlip, LM_Gi2_gi13_LFlip1U, LM_Gi2_gi13_LSling, LM_Gi2_gi13_PF, LM_Gi2_gi13_Parts, LM_Gi2_gi13_RFlip, LM_Gi2_gi13_RFlipU, LM_Gi2_gi13_RSling, LM_Gi2_gi13_RSling1, LM_Gi2_gi13_RSling2, _
' LM_Gi2_gi13_Sling2, LM_Gi2_gi13_sw103, LM_Gi2_gi13_sw106, LM_Gi2_gi13_sw96, LM_Gi2_gi14_LFlip, LM_Gi2_gi14_LFlip1U, LM_Gi2_gi14_LFlipU, LM_Gi2_gi14_PF, LM_Gi2_gi14_Parts, LM_Gi2_gi14_RFlip, LM_Gi2_gi14_RFlipU, LM_Gi2_gi14_RSling, LM_Gi2_gi14_RSling1, LM_Gi2_gi14_RSling2, LM_Gi2_gi14_Sling2, LM_I_L0_LFlip, LM_I_L0_LFlipU, LM_I_L0_PF, LM_I_L0_Parts, LM_I_L0_RFlip, LM_I_L0_RFlipU, LM_I_L10_PF, LM_I_L10_Parts, LM_I_L10_sw95, LM_I_L11_LSling, LM_I_L11_PF, LM_I_L11_Parts, LM_I_L11_sw105, LM_I_L12_PF, LM_I_L12_Parts, LM_I_L12_RSling, LM_I_L12_RSling1, LM_I_L12_RSling2, LM_I_L12_sw106, LM_I_L13_PF, LM_I_L13_Parts, LM_I_L13_sw96, LM_I_L14_LFlip, LM_I_L14_LFlipU, LM_I_L14_PF, LM_I_L14_Parts, LM_I_L14_RFlip, LM_I_L14_RFlipU, LM_I_L15_LSling, LM_I_L15_LSling1, LM_I_L15_LSling2, LM_I_L15_PF, LM_I_L15_Parts, LM_I_L16_PF, LM_I_L16_Parts, LM_I_L17_PF, LM_I_L17_Parts, LM_I_L20_PF, LM_I_L20_Parts, LM_I_L20_RSling, LM_I_L20_RSling1, LM_I_L20_RSling2, LM_I_L20_Sling2, LM_I_L21_PF, LM_I_L21_Parts, LM_I_L21_RSling, _
' LM_I_L21_RSling1, LM_I_L21_RSling2, LM_I_L22_PF, LM_I_L22_Parts, LM_I_L23_PF, LM_I_L23_Parts, LM_I_L24_PF, LM_I_L24_Parts, LM_I_L25_PF, LM_I_L25_Parts, LM_I_L26_PF, LM_I_L26_Parts, LM_I_L27_PF, LM_I_L27_Parts, LM_I_L27_RRampArm, LM_I_L30_LPF, LM_I_L30_PF, LM_I_L30_Parts, LM_I_L30_sw104, LM_I_L30_sw94, LM_I_L31_Car, LM_I_L31_LPF, LM_I_L31_PF, LM_I_L31_Parts, LM_I_L31_sw104, LM_I_L32_PF, LM_I_L32_Parts, LM_I_L33_PF, LM_I_L33_Parts, LM_I_L33_sw103, LM_I_L33_sw113, LM_I_L34_Layer3, LM_I_L34_PF, LM_I_L34_Parts, LM_I_L34_sw103, LM_I_L34_sw113, LM_I_L35_Bumper1Ring, LM_I_L35_Layer3, LM_I_L35_Layer5, LM_I_L35_PF, LM_I_L35_Parts, LM_I_L35_sw103, LM_I_L35_sw112, LM_I_L35_sw113, LM_I_L35_sw30, LM_I_L36_Bumper1Ring, LM_I_L36_Layer3, LM_I_L36_Layer5, LM_I_L36_PF, LM_I_L36_Parts, LM_I_L36_sw102, LM_I_L36_sw112, LM_I_L36_sw92, LM_I_L37_Bumper1Ring, LM_I_L37_Layer3, LM_I_L37_Layer5, LM_I_L37_PF, LM_I_L37_Parts, LM_I_L37_sw112, LM_I_L37_sw92, LM_I_L40_LFlip1U, LM_I_L40_PF, LM_I_L40_Parts, LM_I_L41_LRampArm, LM_I_L41_LRampU, _
' LM_I_L41_PF, LM_I_L41_Parts, LM_I_L42_LRampArm, LM_I_L42_LRampU, LM_I_L42_Layer3, LM_I_L42_PF, LM_I_L42_Parts, LM_I_L43_LRamp_, LM_I_L43_LRampU, LM_I_L43_Layer5, LM_I_L43_PF, LM_I_L43_Parts, LM_I_L43_sw103, LM_I_L43_sw93, LM_I_L44_LRampU, LM_I_L44_Layer5, LM_I_L44_PF, LM_I_L44_Parts, LM_I_L44_sw92, LM_I_L44_sw93, LM_I_L45_Bumper1Ring, LM_I_L45_PF, LM_I_L45_Parts, LM_I_L45_sw92, LM_I_L46_Car, LM_I_L46_LPF, LM_I_L46_PF, LM_I_L46_Parts, LM_I_L46_sw103, LM_I_L46_sw113, LM_I_L46_sw93, LM_I_L47_Layer3, LM_I_L47_PF, LM_I_L47_Parts, LM_I_L47_sw103, LM_I_L47_sw113, LM_I_L50_LFlip1U, LM_I_L50_PF, LM_I_L51_LFlip1U, LM_I_L51_PF, LM_I_L52_LFlip1U, LM_I_L52_PF, LM_I_L52_Parts, LM_I_L53_LRampArmU, LM_I_L53_LRampU, LM_I_L53_PF, LM_I_L53_Parts, LM_I_L54_PF, LM_I_L54_Parts, LM_I_L54_sw102, LM_I_L54_sw92, LM_I_L55_Layer5, LM_I_L55_PF, LM_I_L55_Parts, LM_I_L55_sw102, LM_I_L55_sw112, LM_I_L55_sw92, LM_I_L56_Layer5, LM_I_L56_PF, LM_I_L56_Parts, LM_I_L56_sw102, LM_I_L56_sw112, LM_I_L56_sw92, LM_I_L60_Car, LM_I_L60_LPF, _
' LM_I_L60_Layer1, LM_I_L60_Layer2, LM_I_L60_Parts, LM_I_L60_RFlip1, LM_I_L60_RFlip1U)
' VLM  Arrays - End




' VLM T Arrays - Start
' Arrays per baked part
Dim BP_TTMCar: BP_TTMCar=Array(TBM_TMCar, TLM_TFla_F01_TMCar, TLM_TFla_F02_TMCar, TLM_TFla_F17_TMCar, TLM_TFla_F18_TMCar, TLM_TFla_F19_TMCar, TLM_TFla_F20_TMCar, TLM_TFla_F22_TMCar, TLM_TFla_F23_TMCar, TLM_TCar_TMCar, TLM_TGi_TMCar)
Dim BP_TTMRyu: BP_TTMRyu=Array(TBM_TMRyu, TLM_TFla_F01_TMRyu, TLM_TFla_F17_TMRyu, TLM_TFla_F18_TMRyu, TLM_TFla_F19_TMRyu, TLM_TCar_TMRyu, TLM_TGi_TMRyu)
Dim BP_TTParts: BP_TTParts=Array(TBM_TParts, TLM_TFla_F01_TParts, TLM_TFla_F02_TParts, TLM_TFla_F17_TParts, TLM_TFla_F18_TParts, TLM_TFla_F19_TParts, TLM_TFla_F20_TParts, TLM_TFla_F22_TParts, TLM_TFla_F23_TParts, TLM_TCar_TParts, TLM_TGi_TParts)
' Arrays per lighting scenario
Dim BL_TLit_Room: BL_TLit_Room=Array(TBM_TMCar, TBM_TMRyu, TBM_TParts)
Dim BL_TTCar: BL_TTCar=Array(TLM_TCar_TMCar, TLM_TCar_TMRyu, TLM_TCar_TParts)
Dim BL_TTFla_F01: BL_TTFla_F01=Array(TLM_TFla_F01_TMCar, TLM_TFla_F01_TMRyu, TLM_TFla_F01_TParts)
Dim BL_TTFla_F02: BL_TTFla_F02=Array(TLM_TFla_F02_TMCar, TLM_TFla_F02_TParts)
Dim BL_TTFla_F17: BL_TTFla_F17=Array(TLM_TFla_F17_TMCar, TLM_TFla_F17_TMRyu, TLM_TFla_F17_TParts)
Dim BL_TTFla_F18: BL_TTFla_F18=Array(TLM_TFla_F18_TMCar, TLM_TFla_F18_TMRyu, TLM_TFla_F18_TParts)
Dim BL_TTFla_F19: BL_TTFla_F19=Array(TLM_TFla_F19_TMCar, TLM_TFla_F19_TMRyu, TLM_TFla_F19_TParts)
Dim BL_TTFla_F20: BL_TTFla_F20=Array(TLM_TFla_F20_TMCar, TLM_TFla_F20_TParts)
Dim BL_TTFla_F22: BL_TTFla_F22=Array(TLM_TFla_F22_TMCar, TLM_TFla_F22_TParts)
Dim BL_TTFla_F23: BL_TTFla_F23=Array(TLM_TFla_F23_TMCar, TLM_TFla_F23_TParts)
Dim BL_TTGi: BL_TTGi=Array(TLM_TGi_TMCar, TLM_TGi_TMRyu, TLM_TGi_TParts)
'' Global arrays
'Dim BGT_Bakemap: BGT_Bakemap=Array(TBM_TMCar, TBM_TMRyu, TBM_TParts)
'Dim BGT_Lightmap: BGT_Lightmap=Array(TLM_TCar_TMCar, TLM_TCar_TMRyu, TLM_TCar_TParts, TLM_TFla_F01_TMCar, TLM_TFla_F01_TMRyu, TLM_TFla_F01_TParts, TLM_TFla_F02_TMCar, TLM_TFla_F02_TParts, TLM_TFla_F17_TMCar, TLM_TFla_F17_TMRyu, TLM_TFla_F17_TParts, TLM_TFla_F18_TMCar, TLM_TFla_F18_TMRyu, TLM_TFla_F18_TParts, TLM_TFla_F19_TMCar, TLM_TFla_F19_TMRyu, TLM_TFla_F19_TParts, TLM_TFla_F20_TMCar, TLM_TFla_F20_TParts, TLM_TFla_F22_TMCar, TLM_TFla_F22_TParts, TLM_TFla_F23_TMCar, TLM_TFla_F23_TParts, TLM_TGi_TMCar, TLM_TGi_TMRyu, TLM_TGi_TParts)
Dim BGT_All: BGT_All=Array(TBM_TMCar, TBM_TMRyu, TBM_TParts, TLM_TCar_TMCar, TLM_TCar_TMRyu, TLM_TCar_TParts, TLM_TFla_F01_TMCar, TLM_TFla_F01_TMRyu, TLM_TFla_F01_TParts, TLM_TFla_F02_TMCar, TLM_TFla_F02_TParts, TLM_TFla_F17_TMCar, TLM_TFla_F17_TMRyu, TLM_TFla_F17_TParts, TLM_TFla_F18_TMCar, TLM_TFla_F18_TMRyu, TLM_TFla_F18_TParts, TLM_TFla_F19_TMCar, TLM_TFla_F19_TMRyu, TLM_TFla_F19_TParts, TLM_TFla_F20_TMCar, TLM_TFla_F20_TParts, TLM_TFla_F22_TMCar, TLM_TFla_F22_TParts, TLM_TFla_F23_TMCar, TLM_TFla_F23_TParts, TLM_TGi_TMCar, TLM_TGi_TMRyu, TLM_TGi_TParts)
' VLM T Arrays - End



' VLM B Arrays - Start
' Arrays per baked part
Dim BP_BParts: BP_BParts=Array(BBM_Parts, BLM_BGi_Parts, BLM_BFlasher_L120_Parts, BLM_BFlasher_L121_Parts, BLM_BFlasher_L122_Parts, BLM_BFlasher_L123_Parts, BLM_BFlasher_L124_Parts, BLM_BFlasher_L125_Parts, BLM_BFlasher_L126_Parts, BLM_BFlasher_L127_Parts)
' Arrays per lighting scenario
Dim BL_BBFlasher_L120: BL_BBFlasher_L120=Array(BLM_BFlasher_L120_Parts)
Dim BL_BBFlasher_L121: BL_BBFlasher_L121=Array(BLM_BFlasher_L121_Parts)
Dim BL_BBFlasher_L122: BL_BBFlasher_L122=Array(BLM_BFlasher_L122_Parts)
Dim BL_BBFlasher_L123: BL_BBFlasher_L123=Array(BLM_BFlasher_L123_Parts)
Dim BL_BBFlasher_L124: BL_BBFlasher_L124=Array(BLM_BFlasher_L124_Parts)
Dim BL_BBFlasher_L125: BL_BBFlasher_L125=Array(BLM_BFlasher_L125_Parts)
Dim BL_BBFlasher_L126: BL_BBFlasher_L126=Array(BLM_BFlasher_L126_Parts)
Dim BL_BBFlasher_L127: BL_BBFlasher_L127=Array(BLM_BFlasher_L127_Parts)
Dim BL_BBGi: BL_BBGi=Array(BLM_BGi_Parts)
Dim BL_BBRoom: BL_BBRoom=Array(BBM_Parts)
' Global arrays
Dim BGB_Bakemap: BGB_Bakemap=Array(BBM_Parts)
Dim BGB_Lightmap: BGB_Lightmap=Array(BLM_BFlasher_L120_Parts, BLM_BFlasher_L121_Parts, BLM_BFlasher_L122_Parts, BLM_BFlasher_L123_Parts, BLM_BFlasher_L124_Parts, BLM_BFlasher_L125_Parts, BLM_BFlasher_L126_Parts, BLM_BFlasher_L127_Parts, BLM_BGi_Parts)
Dim BGB_All: BGB_All=Array(BBM_Parts, BLM_BFlasher_L120_Parts, BLM_BFlasher_L121_Parts, BLM_BFlasher_L122_Parts, BLM_BFlasher_L123_Parts, BLM_BFlasher_L124_Parts, BLM_BFlasher_L125_Parts, BLM_BFlasher_L126_Parts, BLM_BFlasher_L127_Parts, BLM_BGi_Parts)
' VLM B Arrays - End


'******************************************************
'  ZTIM: Timers
'******************************************************

'The FrameTimer interval should be -1, so executes at the display frame rate
'The frame timer should be used to update anything visual, like some animations, shadows, etc.
'However, a lot of animations will be handled in their respective _animate subroutines.

Dim FrameTime, InitFrameTime
InitFrameTime = 0

FrameTimer.Interval = -1
FrameTimer.Enabled = true
Sub FrameTimer_Timer()
  FrameTime = gametime - InitFrameTime 'Calculate FrameTime as some animuations could use this
  InitFrameTime = gametime  'Count frametime
  'Add animation stuff here
  BSUpdate
  UpdateBallBrightness
  RollingUpdate       'Update rolling sounds
  UpdateChunli
  UpdateCarReset
  DoSTAnim
  'ChangedSwitches 'debug code
End Sub

'The CorTimer interval should be 10. It's sole purpose is to update the Cor (physics) calculations
CorTimer.Interval = 10
CorTimer.Enabled = true
Sub CorTimer_Timer(): Cor.Update: End Sub



'******************************************************
'  ZINI: Table Initialization and Exiting
'******************************************************


Sub Table1_Init
    vpmInit Me
  vpmMapLights AllLamps

    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Street Fighter 2" & vbNewLine & "VPW"
        '.Games(cGameName).Settings.Value("rol") = 0
        .Games(cGameName).Settings.Value("sound") = 1 '1 enabled rom sound
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = VarHidden
        .Dip(0) = (0 * 1 + 0 * 2 + 0 * 4 + 0 * 8 + 0 * 16 + 0 * 32 + 1 * 64 + 1 * 128) '01-08
        .Dip(1) = (0 * 1 + 0 * 2 + 0 * 4 + 0 * 8 + 0 * 16 + 1 * 32 + 1 * 64 + 1 * 128) '09-16
        .Dip(2) = (0 * 1 + 0 * 2 + 0 * 4 + 0 * 8 + 0 * 16 + 1 * 32 + 1 * 64 + 1 * 128) '17-24
        .Dip(3) = (1 * 1 + 1 * 2 + 1 * 4 + 0 * 8 + 1 * 16 + 0 * 32 + 1 * 64 + 1 * 128) '25-32
        On Error Resume Next
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
        On Error Goto 0
    End With

  'Nudging
    vpmNudge.TiltSwitch = 151
    vpmNudge.Sensitivity = 5
  vpmNudge.TiltObj = Array(Bumper1, LeftSlingshot, RightSlingshot)

  'Trough
  Set SFBall1 = sw31o.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set SFBall2 = sw31a.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set SFBall3 = sw31b.CreateSizedballWithMass(Ballsize/2,Ballmass)
  'Set SFBall3 = sw21.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Controller.Switch(31) = 1
  Controller.Switch(21) = 1
  gBOT = Array(SFBall1,SFBall2,SFBall3)

    'Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

    'Init car
    Set CarBall = KickerBottom.Createball
    KickerBottom.kick 0, 0
  KickerBottom.Enabled = False
  CarBall.color = RGB(34,34,34)

  'Init passive wire diverter
  WireDiverter1.IsDropped = False
  WireDiverter2.IsDropped = True
  WDiverterF.rotatetostart

  GIState False  'False means GI is on
    SolBG False

End Sub



'*******************************************
'  ZOPT: User Options
'*******************************************

Dim LightLevel : LightLevel = 0.25        ' Level of room lighting (0 to 1), where 0 is dark and 100 is brightest
Dim VolumeDial : VolumeDial = 0.8             ' Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Dim BallRollVolume : BallRollVolume = 0.5     ' Level of ball rolling volume. Value between 0 and 1
Dim RampRollVolume : RampRollVolume = 0.5     ' Level of ramp rolling volume. Value between 0 and 1
Dim VRRoomChoice: VRRoomChoice = 2
Dim RailVis: RailVis=1
Dim BackVis: BackVis=1

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

  'Siderail Visibility
  RailVis = Table1.Option("Siderails", 0, 1, 1, 1, 0, Array("Not Visible", "Visible"))

  'Back Top Visibility
  If RenderingMode <> 2 And TestVRonDT = False Then BackVis = Table1.Option("Back Top", 0, 1, 1, 1, 0, Array("Not Visible", "Visible"))

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

  'Playfield Reflections
  v = Table1.Option("Playfield Reflections", 0, 3, 1, 1, 0, Array("Off", "Ball Only", "Static", "Dynamic"))
  ReflectionToggle(v)

  'VR Room
  VRRoomChoice = Table1.Option("VR Room", 0, 2, 1, 2, 0, Array("Ultra Minimal", "Minimal", "Old Arcade"))
  SetupRoom

  'Desktop DMD
  v = Table1.Option("Desktop DMD", 0, 1, 1, 1, 0, Array("Not Visible", "Visible"))
  Scoretext.Visible = v

    ' Sound volumes
    VolumeDial = Table1.Option("Mech Volume", 0, 1, 0.01, 0.8, 1)
    BallRollVolume = Table1.Option("Ball Roll Volume", 0, 1, 0.01, 0.5, 1)
  RampRollVolume = Table1.Option("Ramp Roll Volume", 0, 1, 0.01, 0.5, 1)

  ' Room brightness
  LightLevel = NightDay/100
  SetRoomBrightness LightLevel   'Uncomment this line for lightmapped tables.

    If eventId = 3 And dspTriggered Then dspTriggered = False : DisableStaticPreRendering = False : End If
End Sub

Function ReflectionToggle(state)
  Dim ReflFullObjArray: ReflFullObjArray = Array(BP_Parts, BP_LFlip, BP_LFlipU, BP_LFlip1, BP_LFlip1U, BP_RFlip, BP_RFlipU, BP_RFlip1, BP_RFlip1U, BP_sw93, BP_sw103, BP_sw113)
  Dim ReflPartObjArray: ReflPartObjArray = Array(BM_Parts, BM_LFlip, BM_LFlipU, BM_LFlip1, BM_LFlip1U, BM_RFlip, BM_RFlipU, BM_RFlip1, BM_RFlip1U, BM_sw93, BM_sw103, BM_sw113)
  Dim IBP, BP, v1, v2
  Select Case state
    Case 0: 'No reflections
      v1 = false
      v2 = false
      BM_PF.ReflectionProbe = ""
      BM_Layer2.ReflectionProbe = ""
    Case 1: 'Ball Only reflection
      v1 = false
      v2 = false
      BM_PF.ReflectionProbe = "Playfield Reflections"
      BM_Layer2.ReflectionProbe = "Playfield Reflections"
    Case 2: 'Only reflect static prims (that matters as defined in ReflPartObjArray)
      v1 = false
      v2 = true
      BM_PF.ReflectionProbe = "Playfield Reflections"
      BM_Layer2.ReflectionProbe = "Playfield Reflections"
    Case 3: 'Reflect everything (that matters as defined in ReflFullObjArray)
      v1 = true
      v2 = true
      BM_PF.ReflectionProbe = "Playfield Reflections"
      BM_Layer2.ReflectionProbe = "Playfield Reflections"
  End Select

  For Each IBP in ReflFullObjArray
    For Each BP in IBP
      BP.ReflectionEnabled = v1
    Next
  Next

  For Each BP in ReflPartObjArray
    BP.ReflectionEnabled = v2
  Next

  BM_PF.ReflectionEnabled = False
  BM_Layer2.ReflectionEnabled = False
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

Function dArcSin(x)
  If X = 1 Then
    dArcSin = 90
  ElseIf x = -1 Then
    dArcSin = -90
  Else
    dArcSin = Atn(X / Sqr(-X * X + 1))*180/PI
  End If
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

' Flippers
Sub LeftFlipper_Animate
  Dim a: a = LeftFlipper.CurrentAngle
  FlipperLSh.RotZ = a

  Dim v, BP :
  v = 255.0 * (124.0 - LeftFlipper.CurrentAngle) / (124.0 - 67.0)

  For Each BP in BP_Lflip
    BP.RotZ = a
    BP.visible = v < 128.0
  Next
  For Each BP in BP_LflipU
    BP.RotZ = a
    BP.visible = v >= 128.0
  Next
End Sub

Sub RightFlipper_Animate
  Dim a: a = RightFlipper.CurrentAngle
  FlipperRSh.RotZ = a

  Dim v, BP :
  v = 255.0 * (-124.0 - RightFlipper.CurrentAngle) / (-124.0 + 87.2)

  For Each BP in BP_Rflip
    BP.RotZ = a
    BP.visible = v < 128.0
  Next
  For Each BP in BP_RflipU
    BP.RotZ = a
    BP.visible = v >= 128.0
  Next
End Sub


Sub LeftFlipper1_Animate
  Dim a: a = LeftFlipper1.CurrentAngle
  FlipperLSh1.RotZ = a

  Dim v, BP :
  v = 255.0 * (149.6 - LeftFlipper1.CurrentAngle) / (149.6 - 87.2)

  For Each BP in BP_Lflip1
    BP.RotZ = a
    BP.visible = v < 128.0
  Next
  For Each BP in BP_Lflip1U
    BP.RotZ = a
    BP.visible = v >= 128.0
  Next
End Sub


Sub RightFlipper1_Animate
  Dim a: a = RightFlipper1.CurrentAngle
  'FlipperRSh1.RotZ = a

  Dim v, BP :
  v = 255.0 * (-119.4 - RightFlipper1.CurrentAngle) / (-119.4 + 57)

  For Each BP in BP_Rflip1
    BP.RotZ = a
    BP.visible = v < 128.0
  Next
  For Each BP in BP_Rflip1U
    BP.RotZ = a
    BP.visible = v >= 128.0
  Next
End Sub


' Gates
Sub Gate1_Animate
  Dim a: a = Gate1.CurrentAngle
  Dim BP : For Each BP in BP_Gate1: BP.RotX = a: Next
End Sub


' Bumpers
Sub Bumper1_Animate
  Dim z, BP
  z = Bumper1.CurrentRingOffset
  For Each BP in BP_Bumper1Ring : BP.transz = z: Next
End Sub


' Diverters
Sub Diverter1F_Animate
  Dim z, BP
  z = Diverter1F.CurrentAngle
  For Each BP in BP_Diverter1 : BP.transz = z: Next
  'debug.print "Diverter1F_Animate "&z
End Sub

Sub Diverter2F_Animate
  Dim z, BP
  z = Diverter2F.CurrentAngle
  For Each BP in BP_Diverter2 : BP.transz = z: Next
  'debug.print "Diverter2F_Animate "&z
End Sub

Sub WDiverterF_Animate
  Dim a, BP
  a = WDiverterF.CurrentAngle
  For Each BP in BP_WireDiverter1 : BP.rotz = a: Next
  'debug.print "WDiverterF_Animate "&a
End Sub



' Ramps
Sub LRampF_Animate
  Dim a: a = LRampF.CurrentAngle

  Dim v, BP :
  v = 255.0 * LRampF.CurrentAngle / 40
  'debug.print "LRampF_Animate "&v

  For Each BP in BP_LRamp_
    BP.ObjRotX = a*(15/40)
    BP.visible = v < 128.0
  Next
  For Each BP in BP_LRampArm
    BP.RotX = a - 27
    BP.visible = v < 128.0
  Next
  For Each BP in BP_LRampU
    BP.ObjRotX = a*(15/40)
    BP.visible = v >= 128.0
  Next
  For Each BP in BP_LRampArmU
    BP.RotX = a - 27
    BP.visible = v >= 128.0
  Next
End Sub

Sub RRampF_Animate
  Dim a: a = RRampF.CurrentAngle

  Dim v, BP :
  v = 255.0 * RRampF.CurrentAngle / 40
  'debug.print "RRampF_Animate "&v

  For Each BP in BP_RRamp
    BP.ObjRotX = a*(15/40)
    BP.visible = v < 128.0
  Next
  For Each BP in BP_RRampArm
    BP.RotX = a - 27
    BP.visible = v < 128.0
  Next
  For Each BP in BP_RRampU
    BP.ObjRotX = a*(15/40)
    BP.visible = v >= 128.0
  Next
  For Each BP in BP_RRampArmU
    BP.RotX = a - 27
    BP.visible = v >= 128.0
  Next
End Sub


InitRampBPs
Sub InitRampBPs  'Fix for ramp arms
  Dim BP
  For Each BP in BP_RRampArm : BP.objrotx = 0: Next
  For Each BP in BP_RRampArmU : BP.objrotx = 0: Next
  For Each BP in BP_LRampArm : BP.objrotx = 0: Next
  For Each BP in BP_LRampArmU : BP.objrotx = 0: Next
End Sub



' Switches
Sub sw22_Animate
  Dim z : z = sw22.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw22 : BP.transz = z: Next
End Sub

Sub sw95_Animate
  Dim z : z = sw95.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw95 : BP.transz = z: Next
End Sub

Sub sw96_Animate
  Dim z : z = sw96.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw96 : BP.transz = z: Next
End Sub

Sub sw105_Animate
  Dim z : z = sw105.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw105 : BP.transz = z: Next
End Sub

Sub sw106_Animate
  Dim z : z = sw106.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw106 : BP.transz = z: Next
End Sub

Sub sw114_Animate
  Dim z : z = sw114.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw114 : BP.transz = z: Next
End Sub

Sub sw115_Animate
  Dim z : z = sw115.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw115 : BP.transz = z: Next
End Sub

Sub sw116_Animate
  Dim z : z = sw116.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw116 : BP.transz = z: Next
End Sub




'******************************************************
'   ZBBR: BALL BRIGHTNESS
'******************************************************

Const BallBrightness =  1       'Ball brightness - Value between 0 and 1 (0=Dark ... 1=Bright)

' Constants for plunger lane ball darkening.
' You can make a temporary wall in the plunger lane area and use the co-ordinates from the corner control points.
Const PLOffset = 0.5      'Minimum ball brightness scale in plunger lane
Const PLLeft = 860          'X position of punger lane left
Const PLRight = 920       'X position of punger lane right
Const PLTop = 1225        'Y position of punger lane top
Const PLBottom = 1900       'Y position of punger lane bottom
Dim PLGain: PLGain = (1-PLOffset)/(PLTop-PLBottom)

Sub UpdateBallBrightness
  Dim s, b_base, b_r, b_g, b_b, d_w
  b_base = 120 * BallBrightness + 135 ' orig was 120 and 70

  For s = 0 To UBound(gBOT)
    ' Handle z direction
    If gBOT(s).z >= 24 Then ' above PF
      d_w = b_base*(1 - (gBOT(s).z-25)/500)
    Else  ' below PF
      d_w = b_base*(1 + (gBOT(s).z-25)/100)
    End If
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

' this is for updating the lower playfield ball brightness
Sub L02_animate
  dim c: c = 34 + L02.state*180
  c = Int(c)
  If c > 255 Then c = 255
  'debug.print "LPF ball c="&c
  CarBall.color = c + (c * 256) + (c * 256 * 256)
End Sub


'****************************
'   ZRBR: Room Brightness
'****************************

'This code only applies to lightmapped tables. It is here for reference.
'NOTE: Objects bightness will be affected by the Day/Night slider only if their blenddisablelighting property is less than 1.
'      Lightmapped table primitives have their blenddisablelighting equal to 1, therefore we need this SetRoomBrightness sub
'      to handle updating their effective ambient brighness.

' Update these arrays if you want to change more materials with room light level
Dim RoomBrightnessMtlArray: RoomBrightnessMtlArray = Array("VLM.Bake.Active","VLM.Bake.Solid","VLM.Bake.Solid1")

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

Sub table1_KeyDown(ByVal Keycode)
  If Keycode = LeftFlipperKey Then
    MovButtonLeft.TransX = 8
    Controller.Switch(82) = 1
  End If
    If keycode = RightFlipperkey then
    MovButtonRight.TransX = -8
    Controller.Switch(83) = 1
  End If
    If LowerFlipper Then
        If keycode = RightFlipperkey then
            RightFlipper1.RotateToEnd
      UpdateTopper
        End If
    End If
    If keycode = PlungerKey Then
    SoundPlungerPull
    Plunger.Pullback
    TimerVRPlunger.Enabled = True
    TimerVRPlunger1.Enabled = False
    MovPlunger.TransY = 0
  End If
    If keycode = RightMagnaSave Then Controller.Switch(5) = 1  'Tournament Mode
    If keycode = LeftTiltKey Then Nudge 90, 1: SoundNudgeLeft
    If keycode = RightTiltKey Then Nudge 270, 1: SoundNudgeRight
    If keycode = CenterTiltKey Then Nudge 0, 1: SoundNudgeCenter
  If keycode = StartGameKey Then
    SoundStartButton
    MovStart.TransY = -8
    MovStart1.TransY = -8
  End If
  If keycode = AddCreditKey or keycode = AddCreditKey2 Then SoundCoinIn
    If vpmKeyDown(keycode) Then Exit Sub
End Sub



Sub table1_KeyUp(ByVal Keycode)
  If Keycode = LeftFlipperKey Then
    Controller.Switch(82) = 0
    MovButtonLeft.TransX = 0
  End If
  If Keycode = RightFlipperKey Then
    Controller.Switch(83) = 0
    RightFlipper1.RotatetoStart
    MovButtonRight.TransX = 0
  End If
    If keycode = PlungerKey Then
    If Controller.Switch(22) = 1 Then
      SoundPlungerReleaseBall
    Else
      SoundPlungerReleaseNoBall
    End If
    Plunger.Fire
    TimerVRPlunger.Enabled = False
    TimerVRPlunger1.Enabled = True
  End If
    If keycode = RightMagnaSave Then Controller.Switch(5) = 0
  If keycode = StartGameKey Then
    MovStart.TransY = 0
    MovStart1.TransY = 0
  End If
    If vpmKeyUp(keycode) Then Exit Sub
End Sub


'******************************************************
' ZSOL: Solenoids
'******************************************************

'SolCallback(1)  = ""     'Pop bumper
'SolCallback(2)  = ""     'Left sling
'SolCallback(3)  = ""     'Right sling
SolCallback(4) = "SolOutSw33"
SolCallback(5) = "SolOutSw34"
SolCallback(6) = "SolDiv1"
SolCallback(7) = "SolDiv2"
SolCallback(8) = "SolDiv3"
SolCallback(9) = "SolDiv4"
SolCallback(10) = "SolOutSw13"
SolCallback(11) = "SolOutSw14"
'12 NOT USED
'13 NOT USED
SolCallback(14) = "CarReset"
'Flashers
SolModCallBack(15) = "Flasher15"
SolModCallBack(16) = "Flasher16"
SolModCallBack(17) = "Flasher17"
SolModCallBack(18) = "Flasher18"
SolModCallBack(19) = "Flasher19"
SolModCallBack(20) = "Flasher20"
SolModCallBack(21) = "Flasher21"
SolModCallBack(22) = "Flasher22"
SolModCallBack(23) = "Flasher23"

'SolCallback(24) ' this is the underplayfield flipper
SolCallback(25) = "SolSpinChunli"

SolCallback(26) = "SolBG"
SolCallback(27) = "vpmSolSound ""SolOn"","
SolCallback(28) = "SolRelease"
SolCallback(29) = "SolTrough"
SolCallback(30) = "Knocker"
SolCallback(31) = "GIState"
'SolCallback(32) = "GameOn"

'******************************************************
' ZFLA: Flashers
'******************************************************

const DebugFlashers = false

Sub Flasher15(pwm)
  If DebugFlashers then debug.print "Flasher15 "&pwm
  f15.state = pwm
End Sub
Sub f15_Animate
  AdjustBulbTint f15, BL_F_f15
End Sub

Sub Flasher16(pwm)
  If DebugFlashers then debug.print "Flasher16 "&pwm
  f16.state = pwm
End Sub
Sub f16_Animate
  AdjustBulbTint f16, BL_F_f16
End Sub

Sub Flasher17(pwm)
  If DebugFlashers then debug.print "Flasher17 "&pwm
  f17.state = pwm
End Sub
Sub f17_Animate
  AdjustBulbTint f17, BL_F_f17
End Sub

Sub Flasher18(pwm)
  If DebugFlashers then debug.print "Flasher18 "&pwm
  f18.state = pwm
End Sub
Sub f18_Animate
  AdjustBulbTint f18, BL_F_f18
End Sub

Sub Flasher19(pwm)
  If DebugFlashers then debug.print "Flasher19 "&pwm
  f19.state = pwm
End Sub
Sub f19_Animate
  AdjustBulbTint f19, BL_F_f19
End Sub

Sub Flasher20(pwm)
  If DebugFlashers then debug.print "Flasher20 "&pwm
  f20.state = pwm
End Sub
Sub f20_Animate
  AdjustBulbTint f20, BL_F_f20
End Sub

Sub Flasher21(pwm)
  If DebugFlashers then debug.print "Flasher21 "&pwm
  f21.state = pwm
End Sub
Sub f21_Animate
  AdjustBulbTint f21, BL_F_f21
End Sub

Sub Flasher22(pwm)
  If DebugFlashers then debug.print "Flasher22 "&pwm
  f22.state = pwm
End Sub
Sub f22_Animate
  AdjustBulbTint f22, BL_F_f22
End Sub

Sub Flasher23(pwm)
  If DebugFlashers then debug.print "Flasher23 "&pwm
  f23.state = pwm
End Sub
Sub f23_Animate
  AdjustBulbTint f23, BL_F_f23
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


'******************************************************
' ZDRN: Drain, Trough, and Ball Release
'******************************************************

'********************* TROUGH *************************

Sub sw31o_UnHit : UpdateTrough : End Sub
Sub sw31o_Hit   : UpdateTrough : End Sub
Sub sw31a_UnHit : Controller.Switch(31) = 0 : UpdateTrough : End Sub
Sub sw31a_Hit   : Controller.Switch(31) = 1 : UpdateTrough : End Sub
Sub sw31b_UnHit : Controller.Switch(21) = 0 : UpdateTrough : Checksw21: End Sub
Sub sw31b_Hit   : UpdateTrough : End Sub
Sub sw21_UnHit  : UpdateTrough : End Sub
Sub sw21_Hit    : Controller.Switch(21) = 1 : UpdateTrough : RandomSoundDrain sw21 : End Sub

Sub Checksw21  'Hacky fix for trough issue
  if sw21.BallCntOver > 0 Then Controller.Switch(21) = 1
End Sub

Sub UpdateTrough
  UpdateTroughTimer.Interval = 100
  UpdateTroughTimer.Enabled = 1
End Sub

Sub UpdateTroughTimer_Timer
  If sw31o.BallCntOver = 0 Then sw31a.kick 57, 10
  If sw31a.BallCntOver = 0 Then sw31b.kick 57, 10
  UpdateTroughTimer.Enabled = 0
End Sub

'*****************  DRAIN & RELEASE  ******************

Sub SolTrough(enabled)
  If enabled Then
    sw21.kick 57, 20
    SoundSaucerKick 1,sw21
  End If
End Sub

Sub SolRelease(enabled)
  If enabled Then
    sw31o.kick 57, 10
    RandomSoundBallRelease sw31o
  End If
End Sub




'******************************************************
' ZFLP: FLIPPERS
'******************************************************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"
SolCallback(sULFlipper) = "SolULFlipper"
SolCallback(24) = "SolRFlipper1"

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
    If rightflipper.currentangle > rightflipper.endangle - ReflipAngle Then
      RandomSoundReflipUpRight RightFlipper
    Else
      SoundFlipperUpAttackRight RightFlipper
      RandomSoundFlipperUpRight RightFlipper
    End If
  Else
    FlipperDeActivate RightFlipper, RFPress
    RightFlipper.RotateToStart
    If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
      RandomSoundFlipperDownRight RightFlipper
    End If
    FlipperRightHitParm = FlipperUpSoundLevel
  End If
End Sub

Sub SolULFlipper(Enabled)
  If Enabled Then
    FlipperActivate LeftFlipper1, LFPress1
    LeftFlipper1.RotateToEnd
    If leftflipper1.currentangle < leftflipper1.endangle + ReflipAngle Then
      RandomSoundReflipUpLeft LeftFlipper1
    Else
      SoundFlipperUpAttackLeft LeftFlipper1
      RandomSoundFlipperUpLeft LeftFlipper1
    End If
  Else
    FlipperDeActivate LeftFlipper1, LFPress1
    LeftFlipper1.RotateToStart
    If LeftFlipper1.currentangle < LeftFlipper1.startAngle - 5 Then
      RandomSoundFlipperDownLeft LeftFlipper1
    End If
    FlipperLeftHitParm = FlipperUpSoundLevel
  End If
End Sub

Sub SolRFlipper1(Enabled)  'Car flipper
    LowerFlipper = Enabled
  UpdateCar.Enabled = Enabled
End Sub


'Flipper collide subs
Sub LeftFlipper1_Collide(parm)
  LeftFlipperCollide parm
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



'******************************************************
' ZDIV: Diverters
'******************************************************
Dim RRDown: RRDown = True
Dim LRDown: LRDown = True


Sub SolDiv1(enabled)
  Dim BP
  SoundSaucerKick 0, BM_Diverter1
    if enabled then
        Diverter1.isdropped = 1
    Diverter1F.RotateToEnd
    else
        Diverter1.isdropped = 0
    Diverter1F.RotateToStart
    End if
End Sub

Sub SolDiv2(enabled)
  SoundSaucerKick 0, BM_Diverter2
    if enabled then
        Diverter2.isdropped = 1
    Diverter2F.RotateToEnd
    else
        Diverter2.isdropped = 0
    Diverter2F.RotateToStart
    End if
End Sub

Sub SolDiv3(enabled)
  SoundSaucerKick 0, BM_LRamp_
    if enabled then
    LRDown = False
    LRampF.RotateToEnd
        LeftRampDown.Collidable = False
    LeftRampUp.Collidable = True
    else
    LRDown = True
    LRampF.RotateToStart
        LeftRampDown.Collidable = True
    LeftRampUp.Collidable = False
    End if
End Sub

Sub SolDiv4(enabled)
  SoundSaucerKick 0, BM_RRamp
    if enabled then
    RRDown = False
    RRampF.RotateToEnd
        RightRampDown.Collidable = False
    RightRampUp.Collidable = True
    else
    RRDown = True
    RRampF.RotateToStart
        RightRampDown.Collidable = True
    RightRampUp.Collidable = False
    End if
End Sub


' Passive wire diverter

Sub TriggerWD1_Hit
  WireDiverter1.IsDropped = True
  WireDiverter2.IsDropped = False
  WDiverterF.RotateToEnd
End Sub

Sub TriggerWD1_UnHit
  If Activeball.vely > 0 Then
    WireDiverter1.IsDropped = False
    WireDiverter2.IsDropped = True
    WDiverterF.RotateToStart
  End If
End Sub

Sub TriggerWD2_Unhit
  WireDiverter1.IsDropped = False
  WireDiverter2.IsDropped = True
  WDiverterF.RotateToStart
End Sub


'******************************************************
' ZCHF: Chunli flipper
'******************************************************
' apophis   FIXME

Dim ChunliAngle, SpinChunli, ChunliDelta
Dim LMs_SpinFlip1, LMs_SpinFlip2, LMs_SpinFlip3, LMs_SpinFlip4

InitChunli
Sub InitChunli
  Dim i
  ChunliAngle = 180
  ChunliDelta = 0
  SpinChunli = False
  ReDim LMs_SpinFlip1(UBound(BP_SpinFlip1)-1)
  ReDim LMs_SpinFlip2(UBound(BP_SpinFlip2)-1)
  ReDim LMs_SpinFlip3(UBound(BP_SpinFlip3)-1)
  ReDim LMs_SpinFlip4(UBound(BP_SpinFlip4)-1)
  For i = 0 to (UBound(BP_SpinFlip1)-1): Set LMs_SpinFlip1(i) = BP_SpinFlip1(i+1): Next
  For i = 0 to (UBound(BP_SpinFlip2)-1): Set LMs_SpinFlip2(i) = BP_SpinFlip2(i+1): Next
  For i = 0 to (UBound(BP_SpinFlip3)-1): Set LMs_SpinFlip3(i) = BP_SpinFlip3(i+1): Next
  For i = 0 to (UBound(BP_SpinFlip4)-1): Set LMs_SpinFlip4(i) = BP_SpinFlip4(i+1): Next
  BM_SpinFlip1.Size_X = BM_SpinFlip1.Size_X*0.99  'mitigate zfighting
  BM_SpinFlip1.Size_Z = BM_SpinFlip1.Size_Z*0.998 'mitigate zfighting
  BM_SpinFlip1.visible = true
  BM_SpinFlip2.visible = false
  BM_SpinFlip3.visible = false
  BM_SpinFlip4.visible = false
  DrawChunli ChunliAngle
End Sub

Sub SolSpinChunli(Enabled)
  ChunliDelta = 0.64
  SpinChunli = Enabled
  If Enabled=True Then
    PlaySoundAtLevelStaticLoop "motor", VolumeDial*0.05, BM_SpinFlip1
  Else
    StopSound "motor"
  End If
End Sub

Sub UpdateChunli
  If ChunliDelta > 0.01 Then
    If SpinChunli = True Then
      ChunliDelta = 0.64   'Spin rate based on measurements from videos
    Else
      ChunliDelta = ChunliDelta - FrameTime*0.64/1300  'Deceleration
      If ChunliDelta < 0.01 Then ChunliDelta = 0
    End If
    ChunliAngle = ChunliAngle + ChunliDelta*FrameTime
    If SpinningFlipper.startangle >= 360 Then ChunliAngle = ChunliAngle - 360
    SpinningFlipper.startangle = ChunliAngle
    DrawChunli ChunliAngle
  End If
End Sub


Sub DrawChunli(a)
  Dim LM, v1, v2, v3, v4, op13, op24
  v1 = ((a > 91) and (a < 269))
  v2 = ((a > 1) and (a < 179))
  v3 = ((a > 271) or (a < 89))
  v4 = ((a > 181) and (a < 359))
  op13 = abs(100*dCos(a))
  op24 = abs(100*dSin(a))
  BM_SpinFlip1.ObjRotZ = a
  For each LM in LMs_SpinFlip1
    LM.Visible = v1
    LM.Opacity = op13
    LM.ObjRotZ = a
  Next
  For each LM in LMs_SpinFlip2
    LM.Visible = v2
    LM.Opacity = op24
    LM.ObjRotZ = a
  Next
  For each LM in LMs_SpinFlip3
    LM.Visible = v3
    LM.Opacity = op13
    LM.ObjRotZ = a
  Next
  For each LM in LMs_SpinFlip4
    LM.Visible = v4
    LM.Opacity = op24
    LM.ObjRotZ = a
  Next
' debug.print "ChunliAngle = "&a
' debug.print "SpinFlip1  Visible = "&v1&"  Opacity = "&op13
' debug.print "SpinFlip2  Visible = "&v2&"  Opacity = "&op24
' debug.print "SpinFlip3  Visible = "&v3&"  Opacity = "&op13
' debug.print "SpinFlip4  Visible = "&v4&"  Opacity = "&op24
End Sub




'******************************************************
' ZCAR: Car Crunch
'******************************************************
' apophis

Const LenBall2Car = 75   'primitive y size of 100, origin in center
Const CrunchDist = 110   'total crunch distance
Const deltay_0 = 44      'crunch distance when car is at start of crunch, when ball is at max vel of 60
Const deltay_f = 38       'crunch distance when car is at end of crunch, when ball is at max vel of 60
Const vy_reset = 7       'velocity of reset (vpuints/msec)
Const CORCrunch = 0.2    'coefficient of restitution for the crunch hits
Const yBMCarInit = -52   'Update this to set the visual cars initial position

Dim xCarInit, yCarInit, kCrunch, mCrunch, bCarReset

InitCar
Sub InitCar
  xCarInit = CarP.x  'Home position
  yCarInit = CarP.y
  mCrunch = deltay_0/60
  kCrunch = (deltay_0 - deltay_f)/CrunchDist
  bCarReset = False
  AnimateCar
End Sub

UpdateCar.Interval = -1
UpdateCar.Enabled = False

Sub UpdateCar_Timer
  dim dx, dy, deltay, dCrunch, ynew, vy, F
  dx = CarP.x - CarBall.x
  dy = CarP.y - CarBall.y + LenBall2Car
  vy = cor.ballvely(CarBall.ID)
  If dy >= 0 And vy < 0 Then   'collision
    RandomSoundMetal2 CarBall
    dCrunch = yCarInit - CarP.y
    F = kCrunch*dCrunch
    deltay = max(-vy*mCrunch - F, 0)
    CarP.y = CarP.y - deltay
    AnimateCar
    CarBall.y = CarP.y + LenBall2Car
    CarBall.vely = -CORCrunch*vy
    dyCrunch = -vy*0.1
    CrunchBounce.Enabled = True
    'debug.print "dCrunch = "&dCrunch
    If deltay>0 And dCrunch<CrunchDist Then vpmTimer.PulseSw 15
    If dCrunch>=CrunchDist Then
      vpmTimer.PulseSw 81
      F01.state = 2: F01.TimerInterval = 2000: F01.TimerEnabled = 1
      debug.print "Car Crunch Win"
    End If
  End If
End Sub

CrunchBounce.Interval = 40
CrunchBounce.Enabled = False
Dim dyCrunch: dyCrunch = 0

Sub CrunchBounce_Timer
  CarP.y = CarP.y + dyCrunch
  AnimateCar
  dyCrunch = dyCrunch*(-0.5)
  If abs(dyCrunch) < 0.05 then
    CrunchBounce.Enabled = False
  End If
End Sub

Sub CarReset(enabled)
    If enabled Then
        bCarReset = True
    SoundSaucerKick 0,CarP
    End If
End Sub

Sub UpdateCarReset
  If Not bCarReset Then Exit Sub
  CarP.y = CarP.y + vy_reset*FrameTime
  If CarP.y >= yCarInit Then
    CarP.y = yCarInit
    bCarReset = False
    dyCrunch = 10
    CrunchBounce.Enabled = True
  End If
  AnimateCar
End Sub

Sub AnimateCar
  Dim BP, dy
  dy = yCarInit - CarP.y
  For each BP in BP_Car: BP.transy = yBMCarInit - dy: Next
  CarShadow.y = CarP.y - 40
End Sub





'************************************************************
' ZSLG: Slingshot Animations
'************************************************************


Dim LStep : LStep = 0 : LeftSlingShot.TimerEnabled = 1
Dim RStep : RStep = 0 : RightSlingShot.TimerEnabled = 1

Sub RightSlingShot_Slingshot
  RS.VelocityCorrect(Activeball)
  vpmTimer.PulseSw(12)            'Sling Switch Number
  RStep = 0
  RightSlingShot_Timer
  RightSlingShot.TimerEnabled = 1
  RightSlingShot.TimerInterval = 17
  RandomSoundSlingshotRight sw106
End Sub

Sub RightSlingShot_Timer
  Dim BL
  Dim x1, x2, y: x1 = True:x2 = False:y = -20
    Select Case RStep
        Case 2:x1 = False:x2 = True:y = -10
        Case 3:x1 = False:x2 = False:y = 0:RightSlingShot.TimerEnabled = 0
    End Select

  For Each BL in BP_RSling1 : BL.Visible = x1: Next
  For Each BL in BP_RSling2 : BL.Visible = x2: Next
  For Each BL in BP_Sling2 : BL.transy = y: Next

    RStep = RStep + 1
End Sub


Sub LeftSlingShot_Slingshot
  LS.VelocityCorrect(Activeball)
  vpmTimer.PulseSw(11)            'Sling Switch Number
  LStep = 0
  LeftSlingShot_Timer
  LeftSlingShot.TimerEnabled = 1
  LeftSlingShot.TimerInterval = 17
  RandomSoundSlingshotLeft sw105
End Sub

Sub LeftSlingShot_Timer
  Dim BL
  Dim x1, x2, y: x1 = True:x2 = False:y = -20
    Select Case LStep
        Case 3:x1 = False:x2 = True:y = -10
        Case 4:x1 = False:x2 = False:y = 0:LeftSlingShot.TimerEnabled = 0
    End Select

  For Each BL in BP_LSling1 : BL.Visible = x1: Next
  For Each BL in BP_LSling2 : BL.Visible = x2: Next
  For Each BL in BP_Sling1 : BL.transy = y: Next

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
  AddSlingsPt 0, 0.00, - 5
  AddSlingsPt 1, 0.45, - 6
  AddSlingsPt 2, 0.48,  0
  AddSlingsPt 3, 0.52,  0
  AddSlingsPt 4, 0.55,  6
  AddSlingsPt 5, 1.00,  5
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
' ZVUK: VUKs and Kickers
'************************************************************


Sub sw13_Hit: Controller.Switch(13) = 1: End Sub
Sub sw13_UnHit: Controller.Switch(13) = 0: End Sub

Sub sw14_Hit: Controller.Switch(14) = 1: End Sub
Sub sw14_UnHit: Controller.Switch(14) = 0: End Sub

Sub sw33_Hit
  Controller.Switch(33) = 1
  WireRampOff
  if activeball.velz < -6 then SoundScoopFall
End Sub
Sub sw33_UnHit: Controller.Switch(33) = 0: End Sub

Sub sw34_Hit
  Controller.Switch(34) = 1
  WireRampOff
  if activeball.velz < -6 then SoundScoopFall
End Sub
Sub sw34_UnHit: Controller.Switch(34) = 0: End Sub

Sub sw100_Hit: vpmTimer.PulseSw 100: End Sub


Sub SolOutSw13(Enabled)
  If Enabled Then
    sw13.kick 200,14
    SoundScoopKick
  End If
End Sub

Sub SolOutSw14(Enabled)
  If Enabled Then
    sw14.kick 200,14
    SoundScoopKick
  End If
End Sub

Sub SolOutSw33(Enabled)
  If Enabled Then
    sw33.kickz 0,40,90,30
    SoundScoopKick
  End If
End Sub

Sub SolOutSw34(Enabled)
  If Enabled Then
    sw34.kickz 0,40,90,30
    SoundScoopKick
  End If
End Sub




'************************************************************
' ZSWI: SWITCHES
'************************************************************

'************************* Bumpers **************************

Sub Bumper1_Hit(): RandomSoundBumperTop Bumper1: vpmTimer.PulseSw 10: End Sub

'************************ Rollovers *************************

Sub sw22_Hit:Controller.Switch(22) = 1:End Sub
Sub sw22_UnHit:Controller.Switch(22) = 0:End Sub

Sub sw32_Hit:Controller.Switch(32) = 1:End Sub
Sub sw32_UnHit:Controller.Switch(32) = 0:End Sub

Sub sw80_Hit:Controller.Switch(80) = 1: WireRampOff: WireRampOn True: End Sub
Sub sw80_UnHit:Controller.Switch(80) = 0:End Sub

Sub sw90_Hit:Controller.Switch(90) = 1: WireRampOn True: End Sub
Sub sw90_UnHit:Controller.Switch(90) = 0:End Sub

Sub sw91_Hit:Controller.Switch(91) = 1: WireRampOn False: End Sub
Sub sw91_UnHit:Controller.Switch(91) = 0:End Sub

Sub sw95_Hit:Controller.Switch(95) = 1:End Sub
Sub sw95_UnHit:Controller.Switch(95) = 0:End Sub

Sub sw96_Hit:Controller.Switch(96) = 1:End Sub
Sub sw96_UnHit:Controller.Switch(96) = 0:End Sub

Sub sw101_Hit:Controller.Switch(101) = 1: WireRampOn False: End Sub
Sub sw101_UnHit:Controller.Switch(101) = 0:End Sub

Sub sw105_Hit:Controller.Switch(105) = 1:End Sub
Sub sw105_UnHit:Controller.Switch(105) = 0:End Sub

Sub sw106_Hit:Controller.Switch(106) = 1:End Sub
Sub sw106_UnHit:Controller.Switch(106) = 0:End Sub

Sub sw110_Hit:Controller.Switch(110) = 1:End Sub
Sub sw110_UnHit:Controller.Switch(110) = 0: WireRampOff: End Sub

Sub sw111_Hit:Controller.Switch(111) = 1:End Sub
Sub sw111_UnHit:Controller.Switch(111) = 0:End Sub

Sub sw114_Hit:Controller.Switch(114) = 1:End Sub
Sub sw114_UnHit:Controller.Switch(114) = 0:End Sub

Sub sw115_Hit:Controller.Switch(115) = 1:End Sub
Sub sw115_UnHit:Controller.Switch(115) = 0:End Sub

Sub sw116_Hit:Controller.Switch(116) = 1:End Sub
Sub sw116_UnHit:Controller.Switch(116) = 0:End Sub

'********************** Ramp Switches ***********************

'********************* Standup Targets **********************

Sub sw92_Hit:STHit 92:End Sub
Sub sw93_Hit:STHit 93:End Sub
Sub sw94_Hit:STHit 94:End Sub
Sub sw102_Hit:STHit 102:End Sub
Sub sw103_Hit:STHit 103:End Sub
Sub sw104_Hit:STHit 104:End Sub
Sub sw112_Hit:STHit 112:End Sub
Sub sw113_Hit:STHit 113:End Sub


Sub sw20_Hit
  vpmTimer.PulseSw 30  'Switches 20 and 30 need to be swapped. Need to confirm -FIXME
  sw20.TimerEnabled = True
  Dim BP: For each BP in BP_sw20: BP.transy = -8: Next
End Sub

Sub sw20_Timer
  sw20.TimerEnabled = False
  Dim BP: For each BP in BP_sw20: BP.transy = 0: Next
End Sub

Sub sw30_Hit
  vpmTimer.PulseSw 20  'Switches 20 and 30 need to be swapped. Need to confirm -FIXME
  sw30.TimerEnabled = True
  Dim BP: For each BP in BP_sw30: BP.transy = -8: Next
End Sub

Sub sw30_Timer
  sw30.TimerEnabled = False
  Dim BP: For each BP in BP_sw30: BP.transy = 0: Next
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

Sub RampTrigger001_Hit: WireRampOn False: End Sub
Sub RampTrigger001_UnHit
  If Activeball.vely > 0 Then WireRampOff
End Sub

Sub RampTrigger002_Hit: WireRampOff: End Sub
Sub RampTrigger003_Hit: WireRampOff: End Sub
Sub RampTrigger004_Hit: WireRampOff: End Sub

Sub RampTrigger005_Hit
  If LRDown=True Then WireRampOn True
End Sub
Sub RampTrigger005_UnHit
  If LRDown=True And Activeball.vely>0 Then WireRampOff
End Sub

Sub RampTrigger006_Hit
  If RRDown=True Then WireRampOn True
End Sub
Sub RampTrigger006_UnHit
  If RRDown=True And Activeball.vely>0 Then WireRampOff
End Sub

Sub RampTrigger007_Hit: WireRampOn False :End Sub

Sub RampTrigger008_Hit: SoundScoopFall :End Sub

Sub RampTrigger009_Hit: SoundScoopFall :End Sub


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
  PlaySoundAtLevelStatic SoundFX("Knocker",DOFKnocker), KnockerSoundLevel, KnockerPosition
End Sub

'/////////////////////////////  DRAIN SOUNDS  ////////////////////////////

Sub RandomSoundDrain(drainswitch)
  PlaySoundAtLevelStatic ("Drain_" & Int(Rnd * 4) + 1), DrainSoundLevel, drainswitch
End Sub

'/////////////////////////////  TROUGH BALL RELEASE SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundBallRelease(drainswitch)
  PlaySoundAtLevelStatic SoundFX("BallRelease" & Int(Rnd * 7) + 1,DOFContactors), BallReleaseSoundLevel, drainswitch
End Sub

'/////////////////////////////  SLINGSHOT SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundSlingshotLeft(sling)
  PlaySoundAtLevelStatic SoundFX("Sling_L" & Int(Rnd * 7) + 1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

Sub RandomSoundSlingshotRight(sling)
  PlaySoundAtLevelStatic SoundFX("Sling_R" & Int(Rnd * 7) + 1,DOFContactors), SlingshotSoundLevel, Sling
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

Sub SoundScoopFall
      PlaySoundAtBallVol "z_RCT_Scoop_Fall_Skillshot", VolumeDial*0.1
End Sub

Sub SoundScoopKick
      PlaySoundAtBallVol "z_RCT_Scoop", VolumeDial*0.3
End Sub


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

Sub RandomSoundMetal2(ball)
  PlaySoundAtLevelStatic ("Metal_Touch_" & Int(Rnd * 13) + 1), Vol(ball) * MetalImpactSoundFactor, ball
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
  Cor.Update
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
Const RelayGISoundLevel = 1.08    'volume level; range [0, 1];

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


'///////////////////////////  COIN SOUNDS  ///////////////////////////

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
'****  END FLEEP MECHANICAL SOUNDS
'******************************************************

'******************************************************
' ZNFF:  FLIPPER CORRECTIONS by nFozzy
'******************************************************
'

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
  FlipperTricks LeftFlipper1, LFPress1, LFCount1, LFEndAngle1, LFState1
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

Function Distance2Obj(obj1, obj2)
  Distance2Obj = SQR((obj1.x - obj2.x)^2 + (obj1.y - obj2.y)^2)
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

Dim LFPress1, LFCount1, LFEndAngle1, LFState1

Const FlipperCoilRampupMode = 0 '0 = fast, 1 = medium, 2 = slow (tap passes should work)

LFState1 = 1
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
Const EOSReturn = 0.035  'mid 80's to early 90's
'Const EOSReturn = 0.025  'mid 90's and later

LFEndAngle1 = Leftflipper1.endangle
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
Sub NoTargetBounce_Hit
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
'****  END PHYSICS DAMPENERS
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


'******************************************************
' ZRST: STAND-UP TARGET INITIALIZATION
'******************************************************


Class StandupTarget
  Private m_primary, m_prims, m_sw, m_animate

  Public Property Get Primary(): Set Primary = m_primary: End Property
  Public Property Let Primary(input): Set m_primary = input: End Property

  Public Property Get Prims(): Prims = m_prims: End Property

  Public Property Get Sw(): Sw = m_sw: End Property
  Public Property Let Sw(input): m_sw = input: End Property

  Public Property Get Animate(): Animate = m_animate: End Property
  Public Property Let Animate(input): m_animate = input: End Property

  Public default Function init(primary, prims, sw, animate)
    Set m_primary = primary
    m_prims = prims
    m_sw = sw
    m_animate = animate

    Set Init = Me
  End Function
End Class

'Define a variable for each stand-up target
Dim ST92, ST93, ST94, ST102, ST103, ST104, ST112, ST113

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

Set ST92 = (new StandupTarget)(sw92, BP_sw92, 92, 0)
Set ST93 = (new StandupTarget)(sw93, BP_sw93, 93, 0)
Set ST94 = (new StandupTarget)(sw94, BP_sw94, 94, 0)
Set ST102 = (new StandupTarget)(sw102, BP_sw102, 102, 0)
Set ST103 = (new StandupTarget)(sw103, BP_sw103, 103, 0)
Set ST104 = (new StandupTarget)(sw104, BP_sw104, 104, 0)
Set ST112 = (new StandupTarget)(sw112, BP_sw112, 112, 0)
Set ST113 = (new StandupTarget)(sw113, BP_sw113, 113, 0)

'Add all the Stand-up Target Arrays to Stand-up Target Animation Array
'   STAnimationArray = Array(ST1, ST2, ....)
Dim STArray
STArray = Array(ST92, ST93, ST94, ST102, ST103, ST104, ST112, ST113)

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
    STArray(i).animate = STAnimate(STArray(i).primary,STArray(i).prims,STArray(i).sw,STArray(i).animate)
  Next
End Sub

Function STAnimate(primary, prims, switch,  animate)
  Dim animtime, a, y

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
    For Each a In prims: a.transy =  - STMaxOffset: Next
    vpmTimer.PulseSw switch
    STAnimate = 2
  ElseIf animate = 2 Then
    y = prims(0).transy + STAnimStep
    If y >= 0 Then
      y = 0
      primary.collidable = 1
      STAnimate = 0
    Else
      STAnimate = 2
    End If
    For Each a In prims: a.transy = y: Next
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

'******************************************************
'  ZGIU:  GI Control
'******************************************************

dim gilvl

Sub GIState(Enabled)
  If Enabled=False Then
    gilvl = 1
    Sound_Flash_Relay 1 , GiRelay
  Else
    gilvl = 0
    Sound_Flash_Relay 0 , GiRelay
  End If
  Dim bulb: For Each bulb in GI: bulb.State = gilvl: Next
End Sub


Sub SolBG(Enabled)
If Enabled=False Then
  gibg.state = 1
    Sound_Gi_Relay 1 , BackglassRelay
Else
  gibg.state = 0
    Sound_Gi_Relay 0 , BackglassRelay
End If
End Sub

'******************************************************
'  ZKNO:  Knocker Sound
'******************************************************

Sub Knocker(Enabled)
If Enabled Then
KnockerSolenoid
End If
End Sub

'******************************************************
'  ZINS:  Inserts Special Cases
'******************************************************

' Trick to make the white inserts brighter when they Flasher
Const nIns = 0.3

Sub l17a_animate: l17.state = l17a.state^nIns : End Sub
Sub l22a_animate: l22.state = l22a.state^nIns : End Sub
Sub l23a_animate: l23.state = l23a.state^nIns : End Sub
Sub l26a_animate: l26.state = l26a.state^nIns : End Sub
Sub l27a_animate: l27.state = l27a.state^nIns : End Sub
Sub l34a_animate: l34.state = l34a.state^nIns : End Sub
Sub l37a_animate: l37.state = l37a.state^nIns : End Sub
Sub l42a_animate: l42.state = l42a.state^nIns : End Sub
Sub l45a_animate: l45.state = l45a.state^nIns : End Sub
Sub l47a_animate: l47.state = l47a.state^nIns : End Sub
Sub l52a_animate: l52.state = l52a.state^nIns : End Sub
Sub l53a_animate: l53.state = l53a.state^nIns : End Sub


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


'***************************************************************
'   ZTPR: Update Topper
'***************************************************************

Dim TopperRyu: TopperRyu = True
Dim TCarAng: TCarAng=0

InitTopper
Sub InitTopper
  dim BP, BL
  For each BP in BP_TTMCar: BP.rotx = 9: Next
  For each BL in BL_TTFla_F01: BL.opacity = 150: Next
  For each BL in BL_TTCar: BL.opacity = 150: Next
End Sub

Sub UpdateTopper
  F01.state = 2: F01.TimerInterval = 300: F01.TimerEnabled = 1
  If TopperRyu Then
    TMRyuF.RotateToEnd: TMRyuF.TimerEnabled = 1
    TCarAng = 5: CarTopper.Enabled = 1
  Else
    F02.state = 2: F02.TimerEnabled = 1
    TCarAng = -5: CarTopper.Enabled = 1
  End If
  TopperRyu = Not TopperRyu
End Sub

Sub F01_Timer: F01.state = 0: F01.TimerEnabled = 0: End Sub
Sub F02_Timer: F02.state = 0: F02.TimerEnabled = 0: End Sub
Sub TMRyuF_Timer:  TMRyuF.RotateToStart:  TMRyuF.TimerEnabled = 0:  End Sub
Sub TMCarF1_Timer: TMCarF1.RotateToStart: TMCarF1.TimerEnabled = 0: End Sub
Sub TMCarF2_Timer: TMCarF2.RotateToStart: TMCarF2.TimerEnabled = 0: End Sub

Sub TMRyuF_animate
  dim BP
  For each BP in BP_TTMRyu: BP.objrotz = TMRyuF.currentangle: Next
End Sub


Sub CarTopper_Timer
  dim BP
  For each BP in BP_TTMCar: BP.roty = TCarAng: Next
  TCarAng = -0.7*TCarAng
  If abs(TCarAng) < 0.1 Then
    TCarAng = 0
    CarTopper.Enabled = 0
  End If
End Sub



'***************************************************************
' ZVRR: VR Room
'***************************************************************

Sub SetupRoom
  Dim VRThing, BP
  If RenderingMode = 2 OR TestVRonDT = True Then
    If VRRoomChoice = 2 Then    ' Old Arcade
      For Each BP in BP_SideRails: BP.visible = 0: Next
      For Each BP in BP_Mov3: BP.visible = 0: Next
      For Each VRThing in BGT_All: VRThing.visible = 1: Next
      For Each VRThing in BGB_All: VRThing.visible = 1: Next
      For Each VRThing in VRCab: VRThing.visible = 1: Next
      For Each VRThing in VRMega: VRThing.visible = 1: Next
      MinRoom.visible = 0
    ElseIf VRRoomChoice = 1 Then    ' Minimal
      For Each BP in BP_SideRails: BP.visible = 0: Next
      For Each BP in BP_Mov3: BP.visible = 0: Next
      For Each VRThing in BGT_All: VRThing.visible = 0: Next
      For Each VRThing in BGB_All: VRThing.visible = 1: Next
      For Each VRThing in VRCab: VRThing.visible = 1: Next
      For Each VRThing in VRMega: VRThing.visible = 0: Next
      MinRoom.visible = 1
    Else          ' Ultra Minimal
      For Each BP in BP_SideRails: BP.visible = RailVis: Next
      For Each BP in BP_Mov3: BP.visible = 1: Next
      For Each VRThing in BGT_All: VRThing.visible = 0: Next
      For Each VRThing in BGB_All: VRThing.visible = 0: Next
      For Each VRThing in VRCab: VRThing.visible = 0: Next
      For Each VRThing in VRMega: VRThing.visible = 0: Next
      MinRoom.visible = 0
      DMD1.visible = 1
    End If
  Else
    For Each BP in BP_SideRails: BP.visible = RailVis: Next
    For Each BP in BP_Mov3: BP.visible = BackVis: Next
    For Each VRThing in BGT_All: VRThing.visible = 0: Next
    For Each VRThing in BGB_All: VRThing.visible = 0: Next
    For Each VRThing in VRCab: VRThing.visible = 0: Next
    For Each VRThing in VRMega: VRThing.visible = 0: Next
    MinRoom.visible = 0
  End If
End Sub

Sub L1_animate  'start button
  dim s : s = L1.state
  MovStart1.opacity = 100*s
  MovStart1.blenddisablelighting = 4*s
End Sub


' VR Plunger code

Sub TimerVRPlunger_Timer
  If MovPlunger.TransY < 90 then
    MovPlunger.TransY = MovPlunger.TransY + 5
    'MovPlunger.TransY = (5* Plunger.Position) -20
    Dim BP: For each BP in BP_Plunger: BP.TransY = MovPlunger.TransY: Next
  End If
End Sub

Sub TimerVRPlunger1_Timer
  MovPlunger.TransY = (5* Plunger.Position) -20
  Dim BP: For each BP in BP_Plunger: BP.TransY = MovPlunger.TransY: Next
End Sub



'***************************************************************
' ZDBG: Debug Code
'***************************************************************

Dim LastSwitch(150)
OnSwtiches
Sub OnSwtiches
  Dim c
  For c = 0 to 149
    If Controller.Switch(c) <> 0 Then
      debug.print "Controller.Switch("&c&") = "&Controller.Switch(c)
    End If
    LastSwitch(c) = Controller.Switch(c)
  Next
End Sub

Sub ChangedSwitches
  Dim c
  For c = 0 to 149
    If Controller.Switch(c) <> LastSwitch(c) Then
      debug.print "CHANGED -- Controller.Switch("&c&") = "&Controller.Switch(c)
    End If
    LastSwitch(c) = Controller.Switch(c)
  Next
End Sub




'******************************************************
'  ZLOG: Change Log
'******************************************************
' 01 - Gedankekojote97 - Initial build and scripting
' 02 - Gedankekojote97 - Added some new meshes for subway, kicker, spinner
' 03 - apophis -  Lots of script updates including PWM implementation. Fixed physical trough, diverets, and many other things.
' 04 - apophis - Fixed L60 blinking. Fixed slingshot sound issue.
' 05 - apophis - Completely reworked car crunch mech
' 06 - apophis - Fix some car crunch mech issues. Added some missing sfx. Added ramp roll sfx triggers.
' 07 - Gedankekojote97 - Fixed deflectors, adjusted trigger positions, adjusted power of the kickers (not final tweak)
' 08 - Gedankekojote97 - Imported 4k bake
' 09 - apophis - Made all physical objects and lights invisible. Adjusted Diverter1 and Diverter2 wall shapes. Configured GI for rtx shadows.
'                Animated: slingshots, flippers, diverters, ramps and arms, switches, standup targets, car, Chunli diverter
' 10 - apophis - Updated Chunli angles.
' 11 - apophis - Added 2k bake. Updated car crash animation. Fixed car crash lights. Fixed flasher script. Fixed diverter animations. Fix standup target code. Added insert brightness option.
' 12 - apophis - Added new 2k bake. Removed insert brightness option. Added PF reflections. Set some L60 LMs to invisible. Fixed Wall041.
' 13 - apophis - Added new 4k bake. Adjusted car init position and return speed. Re-fix: Set some L60 LMs to invisible.
' 14 - apophis - Reduced flipper strengths and PF slope. Adjusted car crunch parameters. Decelerate Chunli flipper when turned off. Added Chunli motor sfx. Fixed stuck ball in upper PF pocket. Added desktop backdrop image and DMD option. Updated desktop POV.
' 15 - Gedankekojote97 - New Backdrop with fixes. Adjusted/Fixed Car crash.
' 16 - Gedankekojote97 - Fixed ball stuck after launch/next to kickout right.
' 17 - Sixtoe - Added spiral piledriver physics ramp, fixed ball stuck issue top left of upper playfield
' 18 - apophis - Fixed staged flippers. Adjusted sling strength. Adjust white insert brightness curves.
' 19 - apophis - Fixed trough switches.
' 20 - Gedankekojote97 - Added Gi Relay Sound. Switched sound on plastic ramp entrance from metal wire to plastic ramp.
' 21 - Gedankekojote97 - Added new lightmapped topper. Not coded yet.
' 22 - apophis - Fixed GI callback. Coded topper animations.
' 23 - Gedankekojote97 - Updated lightmapped topper.
' 24 - apophis - Revised topper animations.
' 25 - apophis - Slighlty revised topper animations.
' 26 - apophis - Fixed chunli flipper z-fighting. Adjusted motor sound volume. Deleted unused PF and blueprint images from image manager.
' 27 - Gedankekojote97 - Added backglass renders
' 28 - apophis - Wired up backglass render.
' 29 - apophis - Wired up backglass GI lighting correctly.
' 30 - Gedankekojote97 - Added VR room models and bakes.
' RC1 - apophis - Added rules card. Added minimal room. VR Room and siderail options added. Rewired VR backglass gi light. Added VR DMD. Animated VR buttons and plunger, desktop plunger. Updated GI relay sound.
' RC2 - apophis - Added siderail option to ultra min room. Added dirty glass option (thanks FrankEnstein!). Added shadow under car. Fixed stuck ball on launch ramps. Fixed VR room zfighting. Added subway fall sfx. Maybe fixed the trough?
' RC3 - apophis - Fixed dirty glass height and angle, 50% default. Added visual blocker under plunger lane. Fixed zfighting and color of vr cab under backbox.
' RC4 - apophis - Fixed visual blocker wall.
' RC5 - apophis - Fixed more zfighting in arcade VR room. Lowered VR cab slightly and added another visual blocker near plunger. Lowered center nudge strength.
' Release 1.0
' 1.0.1 - apophis - Added optional playflield reflections.
' 1.0.2 - Gedankekojote97 - Changed Car Crash pf from collidable meshes to VPX Walls, changed scoop fall and kickout sounds.
' 1.0.3 - apophis - Added back top visibility option for cab and destop.
' 1.0.4 - Gedankekojote97 - Fixed visible black boxes, adjusted glass high.
' 1.0.5 - apophis - Darkened all balls that go under the PF. Animated car crunch brightness based on car cruch GI.
' 1.0.6 - Gedankekojote97 - Fixed BG Gi and added flasher sound
' 1.0.7 - FrankEnstein - Reworked the reflection options. None, Partial (reflects "world" of select objects"), Full (also reflects lightmaps of select objects)
' 1.0.8 - Gedankekojote97 - Added knocker sounds.
' Release 1.1

' Thalamus : Exit in a clean and proper way
Sub Table1_exit
  Controller.Pause = False
  Controller.Stop
End Sub


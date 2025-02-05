'    _______   __     _______       _______       __      _____  ___    _______       _______       __        _______
'   |   _  "\ |" \   /" _   "|     |   _  "\     /""\    (\"   \|"  \  /" _   "|     |   _  "\     /""\      /"      \
'   (. |_)  :)||  | (: ( \___)     (. |_)  :)   /    \   |.\\   \    |(: ( \___)     (. |_)  :)   /    \    |:        |
'   |:     \/ |:  |  \/ \          |:     \/   /' /\  \  |: \.   \\  | \/ \          |:     \/   /' /\  \   |_____/   )
'   (|  _  \\ |.  |  //  \ ___     (|  _  \\  //  __'  \ |.  \    \. | //  \ ___     (|  _  \\  //  __'  \   //      /
'   |: |_)  :)/\  |\(:   _(  _|    |: |_)  :)/   /  \\  \|    \    \ |(:   _(  _|    |: |_)  :)/   /  \\  \ |:  __   \
'   (_______/(__\_|_)\_______)     (_______/(___/    \___)\___|\____\) \_______)     (_______/(___/    \___)|__|  \___)
'
' Capcom 1996
'
' https://www.ipdb.org/machine.cgi?id=4001
'
'
' Your VPW Big Bang Bartenders
' ----------------------------
' Gedankekojote97: 3D rebuild and rendering with light mapper, nFozzy physics and Fleep sounds
' Apophis: Light mapper script integration, tons of script updates, table rebuild #savetheballs
' Rawd: New RGB VR room and VR cab setup
' Leojreimroc: VR backglass
' ClarkKent: 2D and 3D assets, physics tweaks
' Jsm174: VPX standalone support
' Toxie: VPinMAME support
' Sixtoe: Table cleanup
'
'
' Special Thanks To
' -----------------
' Niwak for creating the amazing blender toolkit!
' Steely for the VR backbox hinge model and lava lamp code
' UncleWilly, Jimmyfingers, Grizz, and Rom, the original team that made BBB in VP9 and FP
' Ninuzzu and ClarkKent for their VPX conversion
' The VPX and VPinMAME developers. Without them the party would have never started.


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
' ZKEY: Key Press Handling
' ZDRN: Drain, Trough, and Ball Release
' ZSOL: Solenoids & Flashers
' ZFLP: Flippers
'   ZGAT: Gates
'   ZLLK: LowerLock And Post Handling
'   ZLKB: Left Kickback
'   ZRDV: Ramp diverters
'   ZALL: Alien Lock Mech Handler
' ZSLG: Slingshot Animations
' ZBMP: Bumper Animations
'   ZSWI: Switches
'   ZKIC: Kickers
'   ZDNC: Dancer Animation
'   ZBOU: VPW TargetBouncer for targets and posts
'   ZDMP: Rubber Dampeners
' ZRRL: Ramp Rolling Sound Effects
'   ZBRL: Ball Rolling and Drop Sounds
'   ZFLE: Fleep Mechanical Sounds
' ZSSC: Slingshot Corrections
' ZSHA: Ball shadows
'   ZRDT: Drop Targets
' ZRST: Stand-Up Targets
'   ZBLL: Blacklight Lamp
'   ZANI: Misc Animations
' ZVLM: VPX Light Mapper
' ZVRR: VR Room
'
'*********************************************************************************************************************************

' CHANGE LOG
' 1.0.1 - apophis - New 4k batch imported.
' 1.0.2 - apophis - Notes:
'  - Updated flipper triggers and scripts
'  - UseLamps = 1. UseVPMModSol = 2.
'  - Removed Lampz.
'  - Light set up for vpmMapLights (faders, timerinteval, visibility)
'  - Fixed timers causing stutters.
'  - Replaced VRBG flasher code with control light links.
'  - Updated blacklight activation script.
'  - Added UpdateBallBrightness and updated SetRoomBrightness
'  - Remove old dynamic shadows. Set up new shadows.
'  - Updated animations.
'  - Update roth code
'  - Update Fleep code
'  - Add Tweak options menu
'  - Added rules card to tweak menu (thanks carny_priest)
' 1.0.3 - apophis - New 1k bake. Flipper animation update.
' 1.0.4 - apophis - New 4k bake. Improved tube ramp. Fixed LMod effect. Fix plunger control light L01.
' 1.0.5 - Gedankekojote97 - blend file updates:
'   - New Droptargets from Capcom
'   - Fixed Ramp Flap and upper gate
'   - New Materials for Apron, Siderails, Metals, Plastics, Ramp, Ramp Flap
'   - Adjusted material for tube
'   - Fixed some shading issues (Wireramps and other meshes)
'   - Fixed some broken meshes
'   - New Environment
' 1.0.5 - apophis - New 4k bake. Added outpost difficulty mod. Fixed diverter animation.
' 2.0 RC1 - apophis - New 4k bake: fixed plastic material transparency. Fix slingshot sfx issue. Fixed issue with white objects in VR.
' 2.0 RC2 - apophis - Made alien bakemaps a little brighter with new VPX material. Adjusted Nestmap0 to make orange tube rings more vibrant color.
' 2.0 RC2.1 - Sixtoe - Adjusted and centred pop bumpers
' 2.0 RC3 - apophis - Adjusted wire ramp colors (ClarkKent). Boosted pop bumper lights (20% brighter). Tweaked ball brightness. Fixed desktop siderail visibilty in VR. New optional desktop backdrop.
' 2.0 RC4 - apophis - Adjusted droptargets drop height. Fixed VR siderail issue (for real this time). Enabled reflections for BM_Parts.
' 2.0 RC5 - apophis - Minor color correction on purple wire ramps, insert fix, and impact sound volume adjustments (ClarkKent).
' 2.0 Release
' 2.0.1 - apophis - Fix VR backglass drift issue. Fixed glass scratches option. Made all lightmaps 30% brighter.
' 2.0.2 - apophis - Reverted lightmap brightness.
' 2.1 Release


Option Explicit
Randomize
SetLocale 1033




'******************************************************
'  ZVAR: Constants and Global Variables
'******************************************************


'VR Init Dims
Dim Stuff
Dim VRRoom
'end VR Dims

Const BallSize = 50
Const BallMass = 1

Dim DesktopMode:DesktopMode = Table.ShowDT
Dim UseVPMDMD:UseVPMDMD = DesktopMode

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const UseVPMModSol = 2

LoadVPM "03060000", "Capcom.VBS", 3.10

Const cGameName = "bbb109"
Const UseSolenoids = 2
Const UseLamps = 1
Const UseSync = 1
Const HandleMech = 0

'Standard Sounds
Const SSolenoidOn = "Solenoid"
Const SSolenoidOff = ""
Const SCoin = ""


Dim xx
Dim BBBall1, BBBall2, BBBall3, BBBall4, BBCapBall1, BBCapBall2, BBCapBall3, BBCapBall4, gBOT
Dim bsRHole

Const tnob = 8
Const lob = 4
Dim tablewidth: tablewidth = Table.width
Dim tableheight: tableheight = Table.height




'******************************************************
'  ZVLM: VLM Arrays
'******************************************************

' VLM  Arrays - Start
' Arrays per baked part
Dim BP_Alien1: BP_Alien1=Array(BM_Alien1, LM_Flashers_F22_Alien1, LM_Flashers_F25_Alien1, LM_Gi_L35_Alien1, LM_Gi_L36_Alien1, LM_Gi_L37_Alien1, LM_Flashers_L44_Alien1, LM_Flashers_L62_Alien1)
Dim BP_Alien2: BP_Alien2=Array(BM_Alien2, LM_Flashers_F22_Alien2, LM_Flashers_F25_Alien2, LM_Gi_L37_Alien2, LM_Gi_L39_Alien2)
Dim BP_Bumper1Ring: BP_Bumper1Ring=Array(BM_Bumper1Ring, LM_Flashers_F22_Bumper1Ring, LM_Flashers_F23_Bumper1Ring, LM_Flashers_F24_Bumper1Ring, LM_Inserts_L121_Bumper1Ring, LM_Gi_L122_Bumper1Ring, LM_Gi_L123_Bumper1Ring, LM_Gi_L124_Bumper1Ring, LM_Inserts_L126_Bumper1Ring, LM_Inserts_L127_Bumper1Ring, LM_Gi_L26_Bumper1Ring, LM_Gi_L33_Bumper1Ring, LM_Gi_L34_Bumper1Ring, LM_Gi_L35_Bumper1Ring, LM_Gi_L40_Bumper1Ring)
Dim BP_Bumper2Ring: BP_Bumper2Ring=Array(BM_Bumper2Ring, LM_Flashers_F22_Bumper2Ring, LM_Flashers_F23_Bumper2Ring, LM_Flashers_L104_Bumper2Ring, LM_Inserts_L121_Bumper2Ring, LM_Gi_L122_Bumper2Ring, LM_Gi_L123_Bumper2Ring, LM_Inserts_L127_Bumper2Ring, LM_Gi_L26_Bumper2Ring, LM_Gi_L27_Bumper2Ring, LM_Gi_L34_Bumper2Ring, LM_Gi_L35_Bumper2Ring, LM_Gi_L36_Bumper2Ring, LM_Gi_L40_Bumper2Ring, LM_Flashers_L52_Bumper2Ring, LM_Flashers_L53_Bumper2Ring, LM_Flashers_L54_Bumper2Ring)
Dim BP_Bumper3Ring: BP_Bumper3Ring=Array(BM_Bumper3Ring, LM_Flashers_F22_Bumper3Ring, LM_Flashers_F23_Bumper3Ring, LM_Flashers_L104_Bumper3Ring, LM_Gi_L123_Bumper3Ring, LM_Inserts_L127_Bumper3Ring, LM_Gi_L26_Bumper3Ring, LM_Gi_L27_Bumper3Ring, LM_Flashers_L54_Bumper3Ring)
Dim BP_DiverterLeft: BP_DiverterLeft=Array(BM_DiverterLeft, LM_Flashers_F21_DiverterLeft, LM_Flashers_F22_DiverterLeft, LM_Gi_L29_DiverterLeft)
Dim BP_DiverterRight: BP_DiverterRight=Array(BM_DiverterRight, LM_Flashers_F21_DiverterRight, LM_Flashers_F22_DiverterRight, LM_Flashers_F23_DiverterRight, LM_Gi_L29_DiverterRight, LM_Inserts_L49_DiverterRight, LM_Flashers_L62_DiverterRight)
Dim BP_GateL: BP_GateL=Array(BM_GateL, LM_Flashers_F22_GateL, LM_Flashers_F23_GateL, LM_Gi_L34_GateL, LM_Gi_L35_GateL, LM_Inserts_L49_GateL, LM_Flashers_L53_GateL, LM_Flashers_L54_GateL)
Dim BP_GateR: BP_GateR=Array(BM_GateR, LM_Flashers_F22_GateR, LM_Flashers_F23_GateR, LM_Inserts_L101_GateR, LM_Gi_L35_GateR, LM_Inserts_L51_GateR, LM_Flashers_L54_GateR)
Dim BP_LFlipper: BP_LFlipper=Array(BM_LFlipper, LM_Flashers_F26_LFlipper, LM_Gi_L14_LFlipper, LM_Inserts_L67_LFlipper, LM_Inserts_L68_LFlipper, LM_Inserts_L72_LFlipper)
Dim BP_LFlipperU: BP_LFlipperU=Array(BM_LFlipperU, LM_Flashers_F26_LFlipperU, LM_Gi_L12_LFlipperU, LM_Gi_L14_LFlipperU, LM_Inserts_L67_LFlipperU, LM_Inserts_L68_LFlipperU, LM_Inserts_L70_LFlipperU, LM_Inserts_L71_LFlipperU)
Dim BP_Layer1: BP_Layer1=Array(BM_Layer1, LM_Flashers_F21_Layer1, LM_Flashers_F22_Layer1, LM_Flashers_F23_Layer1, LM_Flashers_F24_Layer1, LM_Flashers_F25_Layer1, LM_Flashers_F26_Layer1, LM_LMods_Layer1, LM_Gi_L10_Layer1, LM_Inserts_L102_Layer1, LM_Inserts_L103_Layer1, LM_Inserts_L105_Layer1, LM_Inserts_L115_Layer1, LM_Inserts_L116_Layer1, LM_Flashers_L119_Layer1, LM_Flashers_L120_Layer1, LM_Gi_L122_Layer1, LM_Gi_L123_Layer1, LM_Gi_L124_Layer1, LM_Inserts_L126_Layer1, LM_Inserts_L127_Layer1, LM_Gi_L128_Layer1, LM_Gi_L25_Layer1, LM_Gi_L26_Layer1, LM_Gi_L27_Layer1, LM_Gi_L28_Layer1, LM_Gi_L29_Layer1, LM_Inserts_L30_Layer1, LM_Inserts_L31_Layer1, LM_Inserts_L32_Layer1, LM_Gi_L33_Layer1, LM_Gi_L34_Layer1, LM_Gi_L35_Layer1, LM_Gi_L36_Layer1, LM_Gi_L37_Layer1, LM_Gi_L39_Layer1, LM_Gi_L40_Layer1, LM_Inserts_L41_Layer1, LM_Inserts_L42_Layer1, LM_Inserts_L43_Layer1, LM_Flashers_L44_Layer1, LM_Flashers_L45_Layer1, LM_Inserts_L51_Layer1, LM_Flashers_L52_Layer1, LM_Flashers_L53_Layer1, LM_Flashers_L54_Layer1, _
  LM_Flashers_L62_Layer1, LM_Inserts_L81_Layer1, LM_Inserts_L82_Layer1, LM_Gi_L9_Layer1, LM_Inserts_L97_Layer1, LM_Inserts_l57_Layer1, LM_Inserts_l58_Layer1, LM_Inserts_l59_Layer1)
Dim BP_Layer2: BP_Layer2=Array(BM_Layer2, LM_Flashers_F21_Layer2, LM_Flashers_F22_Layer2, LM_Flashers_F23_Layer2, LM_Flashers_F25_Layer2, LM_Inserts_L101_Layer2, LM_Inserts_L121_Layer2, LM_Gi_L123_Layer2, LM_Gi_L124_Layer2, LM_Inserts_L126_Layer2, LM_Gi_L25_Layer2, LM_Gi_L26_Layer2, LM_Gi_L27_Layer2, LM_Gi_L28_Layer2, LM_Gi_L29_Layer2, LM_Inserts_L30_Layer2, LM_Inserts_L31_Layer2, LM_Inserts_L32_Layer2, LM_Gi_L33_Layer2, LM_Gi_L34_Layer2, LM_Gi_L35_Layer2, LM_Gi_L36_Layer2, LM_Gi_L37_Layer2, LM_Gi_L39_Layer2, LM_Inserts_L43_Layer2, LM_Flashers_L44_Layer2, LM_Flashers_L45_Layer2, LM_Inserts_L49_Layer2, LM_Inserts_L50_Layer2, LM_Inserts_L51_Layer2, LM_Flashers_L52_Layer2, LM_Flashers_L53_Layer2, LM_Flashers_L54_Layer2, LM_Flashers_L62_Layer2, LM_Inserts_l57_Layer2, LM_Inserts_l58_Layer2, LM_Inserts_l59_Layer2)
Dim BP_Layer3: BP_Layer3=Array(BM_Layer3, LM_Flashers_F21_Layer3, LM_Flashers_F22_Layer3, LM_Flashers_F23_Layer3, LM_Inserts_L121_Layer3, LM_Gi_L122_Layer3, LM_Gi_L123_Layer3, LM_Gi_L124_Layer3, LM_Gi_L128_Layer3, LM_Gi_L25_Layer3, LM_Gi_L26_Layer3, LM_Gi_L27_Layer3, LM_Gi_L28_Layer3, LM_Gi_L29_Layer3, LM_Inserts_L31_Layer3, LM_Inserts_L32_Layer3, LM_Gi_L33_Layer3, LM_Gi_L34_Layer3, LM_Gi_L35_Layer3, LM_Gi_L36_Layer3, LM_Gi_L40_Layer3, LM_Inserts_L49_Layer3, LM_Inserts_L50_Layer3, LM_Flashers_L52_Layer3, LM_Flashers_L53_Layer3, LM_Flashers_L54_Layer3, LM_Flashers_L62_Layer3, LM_Inserts_l58_Layer3, LM_Inserts_l59_Layer3)
Dim BP_Lsling1: BP_Lsling1=Array(BM_Lsling1, LM_Flashers_F26_Lsling1, LM_Gi_L12_Lsling1, LM_Gi_L14_Lsling1, LM_Inserts_L66_Lsling1, LM_Inserts_L88_Lsling1)
Dim BP_Lsling2: BP_Lsling2=Array(BM_Lsling2, LM_Flashers_F26_Lsling2, LM_Gi_L12_Lsling2, LM_Gi_L14_Lsling2, LM_Inserts_L66_Lsling2, LM_Inserts_L88_Lsling2)
Dim BP_MissionLockPin: BP_MissionLockPin=Array(BM_MissionLockPin, LM_Flashers_F26_MissionLockPin, LM_Gi_L12_MissionLockPin, LM_Gi_L14_MissionLockPin, LM_Gi_L22_MissionLockPin, LM_Gi_L23_MissionLockPin, LM_Gi_L24_MissionLockPin, LM_Inserts_L65_MissionLockPin, LM_Inserts_L66_MissionLockPin, LM_Inserts_L68_MissionLockPin, LM_Inserts_L69_MissionLockPin, LM_Inserts_L70_MissionLockPin, LM_Gi_L9_MissionLockPin)
Dim BP_Parts: BP_Parts=Array(BM_Parts, LM_Flashers_F21_Parts, LM_Flashers_F22_Parts, LM_Flashers_F23_Parts, LM_Flashers_F24_Parts, LM_Flashers_F25_Parts, LM_Flashers_F26_Parts, LM_LMods_Parts, LM_Flashers_L01_Parts, LM_Gi_L10_Parts, LM_Inserts_L100_Parts, LM_Inserts_L101_Parts, LM_Inserts_L102_Parts, LM_Inserts_L103_Parts, LM_Flashers_L104_Parts, LM_Inserts_L105_Parts, LM_Inserts_L106_Parts, LM_Inserts_L107_Parts, LM_Inserts_L108_Parts, LM_Inserts_L109_Parts, LM_Gi_L11_Parts, LM_Inserts_L111_Parts, LM_Inserts_L112_Parts, LM_Inserts_L113_Parts, LM_Inserts_L114_Parts, LM_Inserts_L115_Parts, LM_Inserts_L116_Parts, LM_Inserts_L117_Parts, LM_Inserts_L118_Parts, LM_Flashers_L119_Parts, LM_Gi_L12_Parts, LM_Flashers_L120_Parts, LM_Inserts_L121_Parts, LM_Gi_L122_Parts, LM_Gi_L123_Parts, LM_Gi_L124_Parts, LM_Inserts_L126_Parts, LM_Inserts_L127_Parts, LM_Gi_L128_Parts, LM_Gi_L14_Parts, LM_Gi_L17_Parts, LM_Gi_L18_Parts, LM_Inserts_L19_Parts, LM_Inserts_L20_Parts, LM_Gi_L21_Parts, LM_Gi_L22_Parts, LM_Gi_L23_Parts, _
  LM_Gi_L24_Parts, LM_Gi_L25_Parts, LM_Gi_L26_Parts, LM_Gi_L27_Parts, LM_Gi_L28_Parts, LM_Gi_L29_Parts, LM_Inserts_L30_Parts, LM_Inserts_L31_Parts, LM_Inserts_L32_Parts, LM_Gi_L33_Parts, LM_Gi_L34_Parts, LM_Gi_L35_Parts, LM_Gi_L36_Parts, LM_Gi_L37_Parts, LM_Gi_L39_Parts, LM_Gi_L40_Parts, LM_Inserts_L41_Parts, LM_Inserts_L42_Parts, LM_Inserts_L43_Parts, LM_Flashers_L44_Parts, LM_Flashers_L45_Parts, LM_Inserts_L49_Parts, LM_Inserts_L50_Parts, LM_Inserts_L51_Parts, LM_Flashers_L52_Parts, LM_Flashers_L53_Parts, LM_Flashers_L54_Parts, LM_Flashers_L62_Parts, LM_Inserts_L65_Parts, LM_Inserts_L66_Parts, LM_Inserts_L67_Parts, LM_Inserts_L68_Parts, LM_Inserts_L69_Parts, LM_Inserts_L70_Parts, LM_Inserts_L71_Parts, LM_Inserts_L73_Parts, LM_Inserts_L74_Parts, LM_Inserts_L75_Parts, LM_Inserts_L76_Parts, LM_Inserts_L77_Parts, LM_Inserts_L78_Parts, LM_Inserts_L79_Parts, LM_Inserts_L80_Parts, LM_Inserts_L81_Parts, LM_Inserts_L82_Parts, LM_Inserts_L83_Parts, LM_Inserts_L84_Parts, LM_Inserts_L85_Parts, LM_Inserts_L86_Parts, _
  LM_Inserts_L87_Parts, LM_Inserts_L88_Parts, LM_Inserts_L89_Parts, LM_Gi_L9_Parts, LM_Inserts_L90_Parts, LM_Inserts_L91_Parts, LM_Inserts_L92_Parts, LM_Inserts_L93_Parts, LM_Inserts_L94_Parts, LM_Inserts_L95_Parts, LM_Inserts_L97_Parts, LM_Inserts_L98_Parts, LM_Inserts_L99_Parts, LM_Inserts_l57_Parts, LM_Inserts_l58_Parts, LM_Inserts_l59_Parts)
Dim BP_Playfield: BP_Playfield=Array(BM_Playfield, LM_Flashers_F21_Playfield, LM_Flashers_F22_Playfield, LM_Flashers_F23_Playfield, LM_Flashers_F24_Playfield, LM_Flashers_F25_Playfield, LM_Flashers_F26_Playfield, LM_Gi_L10_Playfield, LM_Inserts_L100_Playfield, LM_Inserts_L101_Playfield, LM_Inserts_L102_Playfield, LM_Inserts_L103_Playfield, LM_Flashers_L104_Playfield, LM_Inserts_L105_Playfield, LM_Inserts_L106_Playfield, LM_Inserts_L107_Playfield, LM_Inserts_L108_Playfield, LM_Inserts_L109_Playfield, LM_Gi_L11_Playfield, LM_Inserts_L110_Playfield, LM_Inserts_L111_Playfield, LM_Inserts_L112_Playfield, LM_Inserts_L113_Playfield, LM_Inserts_L114_Playfield, LM_Inserts_L115_Playfield, LM_Inserts_L116_Playfield, LM_Inserts_L117_Playfield, LM_Inserts_L118_Playfield, LM_Flashers_L119_Playfield, LM_Gi_L12_Playfield, LM_Flashers_L120_Playfield, LM_Inserts_L121_Playfield, LM_Gi_L122_Playfield, LM_Gi_L123_Playfield, LM_Gi_L124_Playfield, LM_Inserts_L126_Playfield, LM_Inserts_L127_Playfield, LM_Gi_L128_Playfield, _
  LM_Gi_L14_Playfield, LM_Gi_L17_Playfield, LM_Gi_L18_Playfield, LM_Inserts_L19_Playfield, LM_Inserts_L20_Playfield, LM_Gi_L21_Playfield, LM_Gi_L22_Playfield, LM_Gi_L23_Playfield, LM_Gi_L24_Playfield, LM_Gi_L25_Playfield, LM_Gi_L26_Playfield, LM_Gi_L27_Playfield, LM_Gi_L28_Playfield, LM_Gi_L29_Playfield, LM_Inserts_L30_Playfield, LM_Inserts_L31_Playfield, LM_Inserts_L32_Playfield, LM_Gi_L33_Playfield, LM_Gi_L34_Playfield, LM_Gi_L35_Playfield, LM_Gi_L36_Playfield, LM_Gi_L37_Playfield, LM_Gi_L39_Playfield, LM_Gi_L40_Playfield, LM_Inserts_L41_Playfield, LM_Inserts_L42_Playfield, LM_Inserts_L43_Playfield, LM_Flashers_L45_Playfield, LM_Inserts_L49_Playfield, LM_Inserts_L50_Playfield, LM_Inserts_L51_Playfield, LM_Flashers_L52_Playfield, LM_Flashers_L53_Playfield, LM_Flashers_L54_Playfield, LM_Flashers_L62_Playfield, LM_Inserts_L65_Playfield, LM_Inserts_L66_Playfield, LM_Inserts_L67_Playfield, LM_Inserts_L68_Playfield, LM_Inserts_L69_Playfield, LM_Inserts_L70_Playfield, LM_Inserts_L71_Playfield, _
  LM_Inserts_L72_Playfield, LM_Inserts_L73_Playfield, LM_Inserts_L74_Playfield, LM_Inserts_L75_Playfield, LM_Inserts_L76_Playfield, LM_Inserts_L77_Playfield, LM_Inserts_L78_Playfield, LM_Inserts_L79_Playfield, LM_Inserts_L80_Playfield, LM_Inserts_L81_Playfield, LM_Inserts_L82_Playfield, LM_Inserts_L83_Playfield, LM_Inserts_L84_Playfield, LM_Inserts_L85_Playfield, LM_Inserts_L86_Playfield, LM_Inserts_L87_Playfield, LM_Inserts_L88_Playfield, LM_Inserts_L89_Playfield, LM_Gi_L9_Playfield, LM_Inserts_L90_Playfield, LM_Inserts_L91_Playfield, LM_Inserts_L92_Playfield, LM_Inserts_L93_Playfield, LM_Inserts_L94_Playfield, LM_Inserts_L95_Playfield, LM_Inserts_L97_Playfield, LM_Inserts_L98_Playfield, LM_Inserts_L99_Playfield)
Dim BP_RFlipper: BP_RFlipper=Array(BM_RFlipper, LM_Flashers_F26_RFlipper, LM_Gi_L14_RFlipper, LM_Gi_L21_RFlipper, LM_Gi_L22_RFlipper, LM_Gi_L23_RFlipper, LM_Gi_L24_RFlipper, LM_Inserts_L68_RFlipper, LM_Inserts_L72_RFlipper, LM_Inserts_L80_RFlipper)
Dim BP_RFlipper1: BP_RFlipper1=Array(BM_RFlipper1, LM_Flashers_F24_RFlipper1, LM_Gi_L128_RFlipper1, LM_Gi_L17_RFlipper1)
Dim BP_RFlipper1U: BP_RFlipper1U=Array(BM_RFlipper1U, LM_Flashers_F24_RFlipper1U, LM_Gi_L128_RFlipper1U, LM_Gi_L17_RFlipper1U)
Dim BP_RFlipperU: BP_RFlipperU=Array(BM_RFlipperU, LM_Flashers_F26_RFlipperU, LM_Gi_L21_RFlipperU, LM_Gi_L22_RFlipperU, LM_Gi_L23_RFlipperU, LM_Gi_L24_RFlipperU, LM_Inserts_L68_RFlipperU, LM_Inserts_L70_RFlipperU, LM_Inserts_L71_RFlipperU, LM_Inserts_L80_RFlipperU)
Dim BP_RPost: BP_RPost=Array(BM_RPost, LM_Gi_L18_RPost, LM_Gi_L21_RPost, LM_Gi_L22_RPost, LM_Gi_L23_RPost, LM_Inserts_L76_RPost)
Dim BP_Rsling1: BP_Rsling1=Array(BM_Rsling1, LM_Flashers_F26_Rsling1, LM_Gi_L21_Rsling1, LM_Gi_L22_Rsling1, LM_Gi_L23_Rsling1, LM_Gi_L24_Rsling1, LM_Inserts_L77_Rsling1, LM_Inserts_L79_Rsling1)
Dim BP_Rsling2: BP_Rsling2=Array(BM_Rsling2, LM_Flashers_F26_Rsling2, LM_Gi_L21_Rsling2, LM_Gi_L22_Rsling2, LM_Gi_L23_Rsling2, LM_Gi_L24_Rsling2, LM_Inserts_L77_Rsling2, LM_Inserts_L79_Rsling2)
Dim BP_SLING1: BP_SLING1=Array(BM_SLING1, LM_Flashers_F26_SLING1, LM_Gi_L21_SLING1, LM_Gi_L22_SLING1, LM_Gi_L23_SLING1, LM_Gi_L24_SLING1, LM_Inserts_L78_SLING1, LM_Inserts_L79_SLING1)
Dim BP_SLING2: BP_SLING2=Array(BM_SLING2, LM_Flashers_F26_SLING2, LM_Gi_L12_SLING2, LM_Gi_L14_SLING2)
Dim BP_Siderails: BP_Siderails=Array(BM_Siderails, LM_Flashers_F21_Siderails, LM_Inserts_L31_Siderails)
Dim BP_Spinner1: BP_Spinner1=Array(BM_Spinner1, LM_Flashers_F22_Spinner1, LM_Flashers_F23_Spinner1, LM_Inserts_L103_Spinner1, LM_Inserts_l57_Spinner1)
Dim BP_Spinner2: BP_Spinner2=Array(BM_Spinner2, LM_Flashers_F22_Spinner2, LM_Flashers_F23_Spinner2, LM_Gi_L36_Spinner2, LM_Flashers_L62_Spinner2)
Dim BP_SpinnerPlate: BP_SpinnerPlate=Array(BM_SpinnerPlate, LM_Flashers_F22_SpinnerPlate, LM_Gi_L25_SpinnerPlate, LM_Inserts_L30_SpinnerPlate)
Dim BP_dancer: BP_dancer=Array(BM_dancer, LM_Flashers_F21_dancer, LM_Flashers_F22_dancer, LM_Flashers_F23_dancer, LM_Gi_L122_dancer, LM_Gi_L28_dancer, LM_Gi_L29_dancer, LM_Gi_L33_dancer, LM_Gi_L34_dancer, LM_Inserts_L49_dancer, LM_Flashers_L52_dancer, LM_Flashers_L53_dancer, LM_Flashers_L54_dancer)
Dim BP_sw17p: BP_sw17p=Array(BM_sw17p, LM_Flashers_F26_sw17p, LM_Gi_L10_sw17p, LM_Gi_L11_sw17p, LM_Gi_L12_sw17p, LM_Gi_L14_sw17p, LM_Gi_L21_sw17p, LM_Inserts_L65_sw17p, LM_Inserts_L82_sw17p, LM_Inserts_L83_sw17p, LM_Inserts_L84_sw17p, LM_Inserts_L85_sw17p, LM_Inserts_L86_sw17p, LM_Inserts_L87_sw17p, LM_Inserts_L88_sw17p, LM_Inserts_L89_sw17p, LM_Gi_L9_sw17p, LM_Inserts_L90_sw17p, LM_Inserts_L91_sw17p, LM_Inserts_L93_sw17p, LM_Inserts_L94_sw17p, LM_Inserts_L98_sw17p, LM_Inserts_L99_sw17p)
Dim BP_sw18p: BP_sw18p=Array(BM_sw18p, LM_Flashers_F26_sw18p, LM_Gi_L10_sw18p, LM_Inserts_L100_sw18p, LM_Gi_L11_sw18p, LM_Inserts_L112_sw18p, LM_Gi_L12_sw18p, LM_Gi_L14_sw18p, LM_Inserts_L81_sw18p, LM_Inserts_L82_sw18p, LM_Inserts_L83_sw18p, LM_Inserts_L84_sw18p, LM_Inserts_L85_sw18p, LM_Inserts_L86_sw18p, LM_Inserts_L87_sw18p, LM_Inserts_L89_sw18p, LM_Gi_L9_sw18p, LM_Inserts_L90_sw18p, LM_Inserts_L91_sw18p, LM_Inserts_L93_sw18p, LM_Inserts_L94_sw18p, LM_Inserts_L97_sw18p, LM_Inserts_L98_sw18p, LM_Inserts_L99_sw18p)
Dim BP_sw19p: BP_sw19p=Array(BM_sw19p, LM_Gi_L10_sw19p, LM_Inserts_L100_sw19p, LM_Gi_L11_sw19p, LM_Gi_L12_sw19p, LM_Gi_L14_sw19p, LM_Inserts_L81_sw19p, LM_Inserts_L82_sw19p, LM_Inserts_L83_sw19p, LM_Inserts_L84_sw19p, LM_Inserts_L85_sw19p, LM_Inserts_L86_sw19p, LM_Inserts_L89_sw19p, LM_Gi_L9_sw19p, LM_Inserts_L90_sw19p, LM_Inserts_L91_sw19p, LM_Inserts_L93_sw19p, LM_Inserts_L94_sw19p, LM_Inserts_L97_sw19p, LM_Inserts_L98_sw19p, LM_Inserts_L99_sw19p)
Dim BP_sw20p: BP_sw20p=Array(BM_sw20p, LM_Gi_L10_sw20p, LM_Inserts_L100_sw20p, LM_Gi_L11_sw20p, LM_Gi_L12_sw20p, LM_Inserts_L81_sw20p, LM_Inserts_L82_sw20p, LM_Inserts_L83_sw20p, LM_Inserts_L84_sw20p, LM_Inserts_L85_sw20p, LM_Inserts_L86_sw20p, LM_Inserts_L89_sw20p, LM_Gi_L9_sw20p, LM_Inserts_L90_sw20p, LM_Inserts_L93_sw20p, LM_Inserts_L97_sw20p, LM_Inserts_L98_sw20p, LM_Inserts_L99_sw20p)
Dim BP_sw21p: BP_sw21p=Array(BM_sw21p, LM_Inserts_L98_sw21p)
Dim BP_sw22p: BP_sw22p=Array(BM_sw22p, LM_Inserts_L100_sw22p, LM_Inserts_L99_sw22p)
Dim BP_sw23p: BP_sw23p=Array(BM_sw23p, LM_Flashers_F23_sw23p, LM_Inserts_L100_sw23p, LM_Inserts_L101_sw23p, LM_Inserts_L102_sw23p, LM_Gi_L25_sw23p)
Dim BP_sw26: BP_sw26=Array(BM_sw26, LM_Flashers_F22_sw26, LM_Gi_L28_sw26, LM_Inserts_L32_sw26)
Dim BP_sw27: BP_sw27=Array(BM_sw27, LM_Flashers_F22_sw27, LM_Gi_L27_sw27)
Dim BP_sw28: BP_sw28=Array(BM_sw28, LM_Flashers_F22_sw28, LM_Flashers_F23_sw28, LM_Gi_L33_sw28, LM_Gi_L34_sw28, LM_Inserts_L49_sw28, LM_Flashers_L52_sw28, LM_Flashers_L53_sw28, LM_Flashers_L54_sw28)
Dim BP_sw29: BP_sw29=Array(BM_sw29, LM_Flashers_F22_sw29, LM_Flashers_F23_sw29, LM_Gi_L34_sw29, LM_Gi_L35_sw29, LM_Flashers_L54_sw29)
Dim BP_sw30: BP_sw30=Array(BM_sw30, LM_Flashers_F22_sw30, LM_Flashers_F23_sw30, LM_Gi_L124_sw30, LM_Gi_L35_sw30, LM_Gi_L36_sw30, LM_Gi_L37_sw30, LM_Flashers_L52_sw30, LM_Flashers_L53_sw30, LM_Flashers_L54_sw30, LM_Flashers_L62_sw30)
Dim BP_sw43: BP_sw43=Array(BM_sw43)
Dim BP_sw44: BP_sw44=Array(BM_sw44, LM_Flashers_F26_sw44, LM_Gi_L12_sw44, LM_Gi_L14_sw44, LM_Gi_L24_sw44, LM_Inserts_L65_sw44, LM_Inserts_L68_sw44)
Dim BP_sw45: BP_sw45=Array(BM_sw45, LM_Flashers_F26_sw45, LM_Gi_L12_sw45, LM_Gi_L14_sw45, LM_Inserts_L65_sw45, LM_Inserts_L87_sw45)
Dim BP_sw49p: BP_sw49p=Array(BM_sw49p, LM_Flashers_F22_sw49p, LM_Flashers_F23_sw49p, LM_Flashers_F24_sw49p, LM_Inserts_L100_sw49p, LM_Inserts_L101_sw49p, LM_Inserts_L102_sw49p, LM_Inserts_L103_sw49p, LM_Inserts_L105_sw49p, LM_Inserts_L106_sw49p, LM_Inserts_L107_sw49p, LM_Inserts_L108_sw49p, LM_Inserts_L109_sw49p, LM_Inserts_L110_sw49p, LM_Inserts_L111_sw49p, LM_Inserts_L112_sw49p, LM_Inserts_L113_sw49p, LM_Inserts_L114_sw49p, LM_Inserts_L115_sw49p, LM_Inserts_L118_sw49p, LM_Gi_L122_sw49p, LM_Inserts_L127_sw49p, LM_Gi_L25_sw49p, LM_Gi_L26_sw49p, LM_Gi_L27_sw49p, LM_Inserts_L98_sw49p, LM_Inserts_L99_sw49p)
Dim BP_sw50p: BP_sw50p=Array(BM_sw50p, LM_Flashers_F22_sw50p, LM_Flashers_F23_sw50p, LM_Flashers_F24_sw50p, LM_Inserts_L100_sw50p, LM_Inserts_L101_sw50p, LM_Inserts_L102_sw50p, LM_Inserts_L103_sw50p, LM_Inserts_L105_sw50p, LM_Inserts_L106_sw50p, LM_Inserts_L107_sw50p, LM_Inserts_L108_sw50p, LM_Inserts_L109_sw50p, LM_Inserts_L110_sw50p, LM_Inserts_L111_sw50p, LM_Inserts_L112_sw50p, LM_Inserts_L113_sw50p, LM_Inserts_L114_sw50p, LM_Inserts_L115_sw50p, LM_Inserts_L117_sw50p, LM_Inserts_L118_sw50p, LM_Flashers_L120_sw50p, LM_Gi_L122_sw50p, LM_Gi_L123_sw50p, LM_Inserts_L127_sw50p, LM_Gi_L25_sw50p, LM_Gi_L26_sw50p, LM_Gi_L27_sw50p, LM_Inserts_L99_sw50p)
Dim BP_sw51p: BP_sw51p=Array(BM_sw51p, LM_Flashers_F22_sw51p, LM_Flashers_F23_sw51p, LM_Flashers_F24_sw51p, LM_Inserts_L100_sw51p, LM_Inserts_L101_sw51p, LM_Inserts_L102_sw51p, LM_Inserts_L103_sw51p, LM_Inserts_L105_sw51p, LM_Inserts_L106_sw51p, LM_Inserts_L107_sw51p, LM_Inserts_L108_sw51p, LM_Inserts_L109_sw51p, LM_Inserts_L110_sw51p, LM_Inserts_L111_sw51p, LM_Inserts_L112_sw51p, LM_Inserts_L113_sw51p, LM_Inserts_L114_sw51p, LM_Inserts_L115_sw51p, LM_Inserts_L116_sw51p, LM_Inserts_L117_sw51p, LM_Inserts_L118_sw51p, LM_Flashers_L120_sw51p, LM_Gi_L122_sw51p, LM_Gi_L123_sw51p, LM_Inserts_L127_sw51p, LM_Gi_L26_sw51p, LM_Gi_L27_sw51p, LM_Inserts_L99_sw51p)
Dim BP_sw52p: BP_sw52p=Array(BM_sw52p, LM_Flashers_F22_sw52p, LM_Flashers_F23_sw52p, LM_Flashers_F24_sw52p, LM_Inserts_L101_sw52p, LM_Inserts_L102_sw52p, LM_Inserts_L103_sw52p, LM_Inserts_L113_sw52p, LM_Inserts_L114_sw52p, LM_Inserts_L115_sw52p, LM_Gi_L122_sw52p, LM_Gi_L26_sw52p, LM_Gi_L27_sw52p)
Dim BP_sw53p: BP_sw53p=Array(BM_sw53p, LM_Flashers_F22_sw53p, LM_Flashers_F23_sw53p, LM_Flashers_F24_sw53p, LM_Inserts_L102_sw53p, LM_Inserts_L105_sw53p, LM_Inserts_L106_sw53p, LM_Inserts_L113_sw53p, LM_Inserts_L114_sw53p, LM_Inserts_L115_sw53p, LM_Gi_L122_sw53p, LM_Gi_L123_sw53p, LM_Gi_L26_sw53p, LM_Gi_L27_sw53p)
Dim BP_sw58p: BP_sw58p=Array(BM_sw58p, LM_Flashers_F22_sw58p, LM_Flashers_F23_sw58p, LM_Inserts_L121_sw58p, LM_Gi_L124_sw58p, LM_Inserts_L126_sw58p, LM_Inserts_L127_sw58p, LM_Gi_L26_sw58p, LM_Gi_L35_sw58p, LM_Gi_L36_sw58p, LM_Gi_L37_sw58p, LM_Gi_L40_sw58p, LM_Inserts_L42_sw58p, LM_Inserts_L43_sw58p, LM_Inserts_L50_sw58p, LM_Inserts_L51_sw58p, LM_Flashers_L54_sw58p)
Dim BP_sw59: BP_sw59=Array(BM_sw59, LM_Flashers_F23_sw59, LM_Gi_L40_sw59)
Dim BP_sw60: BP_sw60=Array(BM_sw60, LM_Flashers_F25_sw60)
Dim BP_sw65: BP_sw65=Array(BM_sw65, LM_Gi_L21_sw65, LM_Gi_L22_sw65, LM_Gi_L23_sw65, LM_Gi_L24_sw65, LM_Inserts_L78_sw65)
Dim BP_sw66: BP_sw66=Array(BM_sw66, LM_Gi_L21_sw66, LM_Gi_L22_sw66, LM_Gi_L23_sw66, LM_Gi_L24_sw66, LM_Inserts_L78_sw66)
Dim BP_sw77: BP_sw77=Array(BM_sw77)
Dim BP_sw78: BP_sw78=Array(BM_sw78)
Dim BP_sw79: BP_sw79=Array(BM_sw79)
Dim BP_sw80: BP_sw80=Array(BM_sw80, LM_Inserts_L116_sw80, LM_Flashers_L120_sw80)
' Arrays per lighting scenario
Dim BL_Flashers_F21: BL_Flashers_F21=Array(LM_Flashers_F21_DiverterLeft, LM_Flashers_F21_DiverterRight, LM_Flashers_F21_Layer1, LM_Flashers_F21_Layer2, LM_Flashers_F21_Layer3, LM_Flashers_F21_Parts, LM_Flashers_F21_Playfield, LM_Flashers_F21_Siderails, LM_Flashers_F21_dancer)
Dim BL_Flashers_F22: BL_Flashers_F22=Array(LM_Flashers_F22_Alien1, LM_Flashers_F22_Alien2, LM_Flashers_F22_Bumper1Ring, LM_Flashers_F22_Bumper2Ring, LM_Flashers_F22_Bumper3Ring, LM_Flashers_F22_DiverterLeft, LM_Flashers_F22_DiverterRight, LM_Flashers_F22_GateL, LM_Flashers_F22_GateR, LM_Flashers_F22_Layer1, LM_Flashers_F22_Layer2, LM_Flashers_F22_Layer3, LM_Flashers_F22_Parts, LM_Flashers_F22_Playfield, LM_Flashers_F22_Spinner1, LM_Flashers_F22_Spinner2, LM_Flashers_F22_SpinnerPlate, LM_Flashers_F22_dancer, LM_Flashers_F22_sw26, LM_Flashers_F22_sw27, LM_Flashers_F22_sw28, LM_Flashers_F22_sw29, LM_Flashers_F22_sw30, LM_Flashers_F22_sw49p, LM_Flashers_F22_sw50p, LM_Flashers_F22_sw51p, LM_Flashers_F22_sw52p, LM_Flashers_F22_sw53p, LM_Flashers_F22_sw58p)
Dim BL_Flashers_F23: BL_Flashers_F23=Array(LM_Flashers_F23_Bumper1Ring, LM_Flashers_F23_Bumper2Ring, LM_Flashers_F23_Bumper3Ring, LM_Flashers_F23_DiverterRight, LM_Flashers_F23_GateL, LM_Flashers_F23_GateR, LM_Flashers_F23_Layer1, LM_Flashers_F23_Layer2, LM_Flashers_F23_Layer3, LM_Flashers_F23_Parts, LM_Flashers_F23_Playfield, LM_Flashers_F23_Spinner1, LM_Flashers_F23_Spinner2, LM_Flashers_F23_dancer, LM_Flashers_F23_sw23p, LM_Flashers_F23_sw28, LM_Flashers_F23_sw29, LM_Flashers_F23_sw30, LM_Flashers_F23_sw49p, LM_Flashers_F23_sw50p, LM_Flashers_F23_sw51p, LM_Flashers_F23_sw52p, LM_Flashers_F23_sw53p, LM_Flashers_F23_sw58p, LM_Flashers_F23_sw59)
Dim BL_Flashers_F24: BL_Flashers_F24=Array(LM_Flashers_F24_Bumper1Ring, LM_Flashers_F24_Layer1, LM_Flashers_F24_Parts, LM_Flashers_F24_Playfield, LM_Flashers_F24_RFlipper1, LM_Flashers_F24_RFlipper1U, LM_Flashers_F24_sw49p, LM_Flashers_F24_sw50p, LM_Flashers_F24_sw51p, LM_Flashers_F24_sw52p, LM_Flashers_F24_sw53p)
Dim BL_Flashers_F25: BL_Flashers_F25=Array(LM_Flashers_F25_Alien1, LM_Flashers_F25_Alien2, LM_Flashers_F25_Layer1, LM_Flashers_F25_Layer2, LM_Flashers_F25_Parts, LM_Flashers_F25_Playfield, LM_Flashers_F25_sw60)
Dim BL_Flashers_F26: BL_Flashers_F26=Array(LM_Flashers_F26_LFlipper, LM_Flashers_F26_LFlipperU, LM_Flashers_F26_Layer1, LM_Flashers_F26_Lsling1, LM_Flashers_F26_Lsling2, LM_Flashers_F26_MissionLockPin, LM_Flashers_F26_Parts, LM_Flashers_F26_Playfield, LM_Flashers_F26_RFlipper, LM_Flashers_F26_RFlipperU, LM_Flashers_F26_Rsling1, LM_Flashers_F26_Rsling2, LM_Flashers_F26_SLING1, LM_Flashers_F26_SLING2, LM_Flashers_F26_sw17p, LM_Flashers_F26_sw18p, LM_Flashers_F26_sw44, LM_Flashers_F26_sw45)
Dim BL_Flashers_L01: BL_Flashers_L01=Array(LM_Flashers_L01_Parts)
Dim BL_Flashers_L104: BL_Flashers_L104=Array(LM_Flashers_L104_Bumper2Ring, LM_Flashers_L104_Bumper3Ring, LM_Flashers_L104_Parts, LM_Flashers_L104_Playfield)
Dim BL_Flashers_L119: BL_Flashers_L119=Array(LM_Flashers_L119_Layer1, LM_Flashers_L119_Parts, LM_Flashers_L119_Playfield)
Dim BL_Flashers_L120: BL_Flashers_L120=Array(LM_Flashers_L120_Layer1, LM_Flashers_L120_Parts, LM_Flashers_L120_Playfield, LM_Flashers_L120_sw50p, LM_Flashers_L120_sw51p, LM_Flashers_L120_sw80)
Dim BL_Flashers_L44: BL_Flashers_L44=Array(LM_Flashers_L44_Alien1, LM_Flashers_L44_Layer1, LM_Flashers_L44_Layer2, LM_Flashers_L44_Parts)
Dim BL_Flashers_L45: BL_Flashers_L45=Array(LM_Flashers_L45_Layer1, LM_Flashers_L45_Layer2, LM_Flashers_L45_Parts, LM_Flashers_L45_Playfield)
Dim BL_Flashers_L52: BL_Flashers_L52=Array(LM_Flashers_L52_Bumper2Ring, LM_Flashers_L52_Layer1, LM_Flashers_L52_Layer2, LM_Flashers_L52_Layer3, LM_Flashers_L52_Parts, LM_Flashers_L52_Playfield, LM_Flashers_L52_dancer, LM_Flashers_L52_sw28, LM_Flashers_L52_sw30)
Dim BL_Flashers_L53: BL_Flashers_L53=Array(LM_Flashers_L53_Bumper2Ring, LM_Flashers_L53_GateL, LM_Flashers_L53_Layer1, LM_Flashers_L53_Layer2, LM_Flashers_L53_Layer3, LM_Flashers_L53_Parts, LM_Flashers_L53_Playfield, LM_Flashers_L53_dancer, LM_Flashers_L53_sw28, LM_Flashers_L53_sw30)
Dim BL_Flashers_L54: BL_Flashers_L54=Array(LM_Flashers_L54_Bumper2Ring, LM_Flashers_L54_Bumper3Ring, LM_Flashers_L54_GateL, LM_Flashers_L54_GateR, LM_Flashers_L54_Layer1, LM_Flashers_L54_Layer2, LM_Flashers_L54_Layer3, LM_Flashers_L54_Parts, LM_Flashers_L54_Playfield, LM_Flashers_L54_dancer, LM_Flashers_L54_sw28, LM_Flashers_L54_sw29, LM_Flashers_L54_sw30, LM_Flashers_L54_sw58p)
Dim BL_Flashers_L62: BL_Flashers_L62=Array(LM_Flashers_L62_Alien1, LM_Flashers_L62_DiverterRight, LM_Flashers_L62_Layer1, LM_Flashers_L62_Layer2, LM_Flashers_L62_Layer3, LM_Flashers_L62_Parts, LM_Flashers_L62_Playfield, LM_Flashers_L62_Spinner2, LM_Flashers_L62_sw30)
Dim BL_Gi_L10: BL_Gi_L10=Array(LM_Gi_L10_Layer1, LM_Gi_L10_Parts, LM_Gi_L10_Playfield, LM_Gi_L10_sw17p, LM_Gi_L10_sw18p, LM_Gi_L10_sw19p, LM_Gi_L10_sw20p)
Dim BL_Gi_L11: BL_Gi_L11=Array(LM_Gi_L11_Parts, LM_Gi_L11_Playfield, LM_Gi_L11_sw17p, LM_Gi_L11_sw18p, LM_Gi_L11_sw19p, LM_Gi_L11_sw20p)
Dim BL_Gi_L12: BL_Gi_L12=Array(LM_Gi_L12_LFlipperU, LM_Gi_L12_Lsling1, LM_Gi_L12_Lsling2, LM_Gi_L12_MissionLockPin, LM_Gi_L12_Parts, LM_Gi_L12_Playfield, LM_Gi_L12_SLING2, LM_Gi_L12_sw17p, LM_Gi_L12_sw18p, LM_Gi_L12_sw19p, LM_Gi_L12_sw20p, LM_Gi_L12_sw44, LM_Gi_L12_sw45)
Dim BL_Gi_L122: BL_Gi_L122=Array(LM_Gi_L122_Bumper1Ring, LM_Gi_L122_Bumper2Ring, LM_Gi_L122_Layer1, LM_Gi_L122_Layer3, LM_Gi_L122_Parts, LM_Gi_L122_Playfield, LM_Gi_L122_dancer, LM_Gi_L122_sw49p, LM_Gi_L122_sw50p, LM_Gi_L122_sw51p, LM_Gi_L122_sw52p, LM_Gi_L122_sw53p)
Dim BL_Gi_L123: BL_Gi_L123=Array(LM_Gi_L123_Bumper1Ring, LM_Gi_L123_Bumper2Ring, LM_Gi_L123_Bumper3Ring, LM_Gi_L123_Layer1, LM_Gi_L123_Layer2, LM_Gi_L123_Layer3, LM_Gi_L123_Parts, LM_Gi_L123_Playfield, LM_Gi_L123_sw50p, LM_Gi_L123_sw51p, LM_Gi_L123_sw53p)
Dim BL_Gi_L124: BL_Gi_L124=Array(LM_Gi_L124_Bumper1Ring, LM_Gi_L124_Layer1, LM_Gi_L124_Layer2, LM_Gi_L124_Layer3, LM_Gi_L124_Parts, LM_Gi_L124_Playfield, LM_Gi_L124_sw30, LM_Gi_L124_sw58p)
Dim BL_Gi_L128: BL_Gi_L128=Array(LM_Gi_L128_Layer1, LM_Gi_L128_Layer3, LM_Gi_L128_Parts, LM_Gi_L128_Playfield, LM_Gi_L128_RFlipper1, LM_Gi_L128_RFlipper1U)
Dim BL_Gi_L14: BL_Gi_L14=Array(LM_Gi_L14_LFlipper, LM_Gi_L14_LFlipperU, LM_Gi_L14_Lsling1, LM_Gi_L14_Lsling2, LM_Gi_L14_MissionLockPin, LM_Gi_L14_Parts, LM_Gi_L14_Playfield, LM_Gi_L14_RFlipper, LM_Gi_L14_SLING2, LM_Gi_L14_sw17p, LM_Gi_L14_sw18p, LM_Gi_L14_sw19p, LM_Gi_L14_sw44, LM_Gi_L14_sw45)
Dim BL_Gi_L17: BL_Gi_L17=Array(LM_Gi_L17_Parts, LM_Gi_L17_Playfield, LM_Gi_L17_RFlipper1, LM_Gi_L17_RFlipper1U)
Dim BL_Gi_L18: BL_Gi_L18=Array(LM_Gi_L18_Parts, LM_Gi_L18_Playfield, LM_Gi_L18_RPost)
Dim BL_Gi_L21: BL_Gi_L21=Array(LM_Gi_L21_Parts, LM_Gi_L21_Playfield, LM_Gi_L21_RFlipper, LM_Gi_L21_RFlipperU, LM_Gi_L21_RPost, LM_Gi_L21_Rsling1, LM_Gi_L21_Rsling2, LM_Gi_L21_SLING1, LM_Gi_L21_sw17p, LM_Gi_L21_sw65, LM_Gi_L21_sw66)
Dim BL_Gi_L22: BL_Gi_L22=Array(LM_Gi_L22_MissionLockPin, LM_Gi_L22_Parts, LM_Gi_L22_Playfield, LM_Gi_L22_RFlipper, LM_Gi_L22_RFlipperU, LM_Gi_L22_RPost, LM_Gi_L22_Rsling1, LM_Gi_L22_Rsling2, LM_Gi_L22_SLING1, LM_Gi_L22_sw65, LM_Gi_L22_sw66)
Dim BL_Gi_L23: BL_Gi_L23=Array(LM_Gi_L23_MissionLockPin, LM_Gi_L23_Parts, LM_Gi_L23_Playfield, LM_Gi_L23_RFlipper, LM_Gi_L23_RFlipperU, LM_Gi_L23_RPost, LM_Gi_L23_Rsling1, LM_Gi_L23_Rsling2, LM_Gi_L23_SLING1, LM_Gi_L23_sw65, LM_Gi_L23_sw66)
Dim BL_Gi_L24: BL_Gi_L24=Array(LM_Gi_L24_MissionLockPin, LM_Gi_L24_Parts, LM_Gi_L24_Playfield, LM_Gi_L24_RFlipper, LM_Gi_L24_RFlipperU, LM_Gi_L24_Rsling1, LM_Gi_L24_Rsling2, LM_Gi_L24_SLING1, LM_Gi_L24_sw44, LM_Gi_L24_sw65, LM_Gi_L24_sw66)
Dim BL_Gi_L25: BL_Gi_L25=Array(LM_Gi_L25_Layer1, LM_Gi_L25_Layer2, LM_Gi_L25_Layer3, LM_Gi_L25_Parts, LM_Gi_L25_Playfield, LM_Gi_L25_SpinnerPlate, LM_Gi_L25_sw23p, LM_Gi_L25_sw49p, LM_Gi_L25_sw50p)
Dim BL_Gi_L26: BL_Gi_L26=Array(LM_Gi_L26_Bumper1Ring, LM_Gi_L26_Bumper2Ring, LM_Gi_L26_Bumper3Ring, LM_Gi_L26_Layer1, LM_Gi_L26_Layer2, LM_Gi_L26_Layer3, LM_Gi_L26_Parts, LM_Gi_L26_Playfield, LM_Gi_L26_sw49p, LM_Gi_L26_sw50p, LM_Gi_L26_sw51p, LM_Gi_L26_sw52p, LM_Gi_L26_sw53p, LM_Gi_L26_sw58p)
Dim BL_Gi_L27: BL_Gi_L27=Array(LM_Gi_L27_Bumper2Ring, LM_Gi_L27_Bumper3Ring, LM_Gi_L27_Layer1, LM_Gi_L27_Layer2, LM_Gi_L27_Layer3, LM_Gi_L27_Parts, LM_Gi_L27_Playfield, LM_Gi_L27_sw27, LM_Gi_L27_sw49p, LM_Gi_L27_sw50p, LM_Gi_L27_sw51p, LM_Gi_L27_sw52p, LM_Gi_L27_sw53p)
Dim BL_Gi_L28: BL_Gi_L28=Array(LM_Gi_L28_Layer1, LM_Gi_L28_Layer2, LM_Gi_L28_Layer3, LM_Gi_L28_Parts, LM_Gi_L28_Playfield, LM_Gi_L28_dancer, LM_Gi_L28_sw26)
Dim BL_Gi_L29: BL_Gi_L29=Array(LM_Gi_L29_DiverterLeft, LM_Gi_L29_DiverterRight, LM_Gi_L29_Layer1, LM_Gi_L29_Layer2, LM_Gi_L29_Layer3, LM_Gi_L29_Parts, LM_Gi_L29_Playfield, LM_Gi_L29_dancer)
Dim BL_Gi_L33: BL_Gi_L33=Array(LM_Gi_L33_Bumper1Ring, LM_Gi_L33_Layer1, LM_Gi_L33_Layer2, LM_Gi_L33_Layer3, LM_Gi_L33_Parts, LM_Gi_L33_Playfield, LM_Gi_L33_dancer, LM_Gi_L33_sw28)
Dim BL_Gi_L34: BL_Gi_L34=Array(LM_Gi_L34_Bumper1Ring, LM_Gi_L34_Bumper2Ring, LM_Gi_L34_GateL, LM_Gi_L34_Layer1, LM_Gi_L34_Layer2, LM_Gi_L34_Layer3, LM_Gi_L34_Parts, LM_Gi_L34_Playfield, LM_Gi_L34_dancer, LM_Gi_L34_sw28, LM_Gi_L34_sw29)
Dim BL_Gi_L35: BL_Gi_L35=Array(LM_Gi_L35_Alien1, LM_Gi_L35_Bumper1Ring, LM_Gi_L35_Bumper2Ring, LM_Gi_L35_GateL, LM_Gi_L35_GateR, LM_Gi_L35_Layer1, LM_Gi_L35_Layer2, LM_Gi_L35_Layer3, LM_Gi_L35_Parts, LM_Gi_L35_Playfield, LM_Gi_L35_sw29, LM_Gi_L35_sw30, LM_Gi_L35_sw58p)
Dim BL_Gi_L36: BL_Gi_L36=Array(LM_Gi_L36_Alien1, LM_Gi_L36_Bumper2Ring, LM_Gi_L36_Layer1, LM_Gi_L36_Layer2, LM_Gi_L36_Layer3, LM_Gi_L36_Parts, LM_Gi_L36_Playfield, LM_Gi_L36_Spinner2, LM_Gi_L36_sw30, LM_Gi_L36_sw58p)
Dim BL_Gi_L37: BL_Gi_L37=Array(LM_Gi_L37_Alien1, LM_Gi_L37_Alien2, LM_Gi_L37_Layer1, LM_Gi_L37_Layer2, LM_Gi_L37_Parts, LM_Gi_L37_Playfield, LM_Gi_L37_sw30, LM_Gi_L37_sw58p)
Dim BL_Gi_L39: BL_Gi_L39=Array(LM_Gi_L39_Alien2, LM_Gi_L39_Layer1, LM_Gi_L39_Layer2, LM_Gi_L39_Parts, LM_Gi_L39_Playfield)
Dim BL_Gi_L40: BL_Gi_L40=Array(LM_Gi_L40_Bumper1Ring, LM_Gi_L40_Bumper2Ring, LM_Gi_L40_Layer1, LM_Gi_L40_Layer3, LM_Gi_L40_Parts, LM_Gi_L40_Playfield, LM_Gi_L40_sw58p, LM_Gi_L40_sw59)
Dim BL_Gi_L9: BL_Gi_L9=Array(LM_Gi_L9_Layer1, LM_Gi_L9_MissionLockPin, LM_Gi_L9_Parts, LM_Gi_L9_Playfield, LM_Gi_L9_sw17p, LM_Gi_L9_sw18p, LM_Gi_L9_sw19p, LM_Gi_L9_sw20p)
Dim BL_Inserts_L100: BL_Inserts_L100=Array(LM_Inserts_L100_Parts, LM_Inserts_L100_Playfield, LM_Inserts_L100_sw18p, LM_Inserts_L100_sw19p, LM_Inserts_L100_sw20p, LM_Inserts_L100_sw22p, LM_Inserts_L100_sw23p, LM_Inserts_L100_sw49p, LM_Inserts_L100_sw50p, LM_Inserts_L100_sw51p)
Dim BL_Inserts_L101: BL_Inserts_L101=Array(LM_Inserts_L101_GateR, LM_Inserts_L101_Layer2, LM_Inserts_L101_Parts, LM_Inserts_L101_Playfield, LM_Inserts_L101_sw23p, LM_Inserts_L101_sw49p, LM_Inserts_L101_sw50p, LM_Inserts_L101_sw51p, LM_Inserts_L101_sw52p)
Dim BL_Inserts_L102: BL_Inserts_L102=Array(LM_Inserts_L102_Layer1, LM_Inserts_L102_Parts, LM_Inserts_L102_Playfield, LM_Inserts_L102_sw23p, LM_Inserts_L102_sw49p, LM_Inserts_L102_sw50p, LM_Inserts_L102_sw51p, LM_Inserts_L102_sw52p, LM_Inserts_L102_sw53p)
Dim BL_Inserts_L103: BL_Inserts_L103=Array(LM_Inserts_L103_Layer1, LM_Inserts_L103_Parts, LM_Inserts_L103_Playfield, LM_Inserts_L103_Spinner1, LM_Inserts_L103_sw49p, LM_Inserts_L103_sw50p, LM_Inserts_L103_sw51p, LM_Inserts_L103_sw52p)
Dim BL_Inserts_L105: BL_Inserts_L105=Array(LM_Inserts_L105_Layer1, LM_Inserts_L105_Parts, LM_Inserts_L105_Playfield, LM_Inserts_L105_sw49p, LM_Inserts_L105_sw50p, LM_Inserts_L105_sw51p, LM_Inserts_L105_sw53p)
Dim BL_Inserts_L106: BL_Inserts_L106=Array(LM_Inserts_L106_Parts, LM_Inserts_L106_Playfield, LM_Inserts_L106_sw49p, LM_Inserts_L106_sw50p, LM_Inserts_L106_sw51p, LM_Inserts_L106_sw53p)
Dim BL_Inserts_L107: BL_Inserts_L107=Array(LM_Inserts_L107_Parts, LM_Inserts_L107_Playfield, LM_Inserts_L107_sw49p, LM_Inserts_L107_sw50p, LM_Inserts_L107_sw51p)
Dim BL_Inserts_L108: BL_Inserts_L108=Array(LM_Inserts_L108_Parts, LM_Inserts_L108_Playfield, LM_Inserts_L108_sw49p, LM_Inserts_L108_sw50p, LM_Inserts_L108_sw51p)
Dim BL_Inserts_L109: BL_Inserts_L109=Array(LM_Inserts_L109_Parts, LM_Inserts_L109_Playfield, LM_Inserts_L109_sw49p, LM_Inserts_L109_sw50p, LM_Inserts_L109_sw51p)
Dim BL_Inserts_L110: BL_Inserts_L110=Array(LM_Inserts_L110_Playfield, LM_Inserts_L110_sw49p, LM_Inserts_L110_sw50p, LM_Inserts_L110_sw51p)
Dim BL_Inserts_L111: BL_Inserts_L111=Array(LM_Inserts_L111_Parts, LM_Inserts_L111_Playfield, LM_Inserts_L111_sw49p, LM_Inserts_L111_sw50p, LM_Inserts_L111_sw51p)
Dim BL_Inserts_L112: BL_Inserts_L112=Array(LM_Inserts_L112_Parts, LM_Inserts_L112_Playfield, LM_Inserts_L112_sw18p, LM_Inserts_L112_sw49p, LM_Inserts_L112_sw50p, LM_Inserts_L112_sw51p)
Dim BL_Inserts_L113: BL_Inserts_L113=Array(LM_Inserts_L113_Parts, LM_Inserts_L113_Playfield, LM_Inserts_L113_sw49p, LM_Inserts_L113_sw50p, LM_Inserts_L113_sw51p, LM_Inserts_L113_sw52p, LM_Inserts_L113_sw53p)
Dim BL_Inserts_L114: BL_Inserts_L114=Array(LM_Inserts_L114_Parts, LM_Inserts_L114_Playfield, LM_Inserts_L114_sw49p, LM_Inserts_L114_sw50p, LM_Inserts_L114_sw51p, LM_Inserts_L114_sw52p, LM_Inserts_L114_sw53p)
Dim BL_Inserts_L115: BL_Inserts_L115=Array(LM_Inserts_L115_Layer1, LM_Inserts_L115_Parts, LM_Inserts_L115_Playfield, LM_Inserts_L115_sw49p, LM_Inserts_L115_sw50p, LM_Inserts_L115_sw51p, LM_Inserts_L115_sw52p, LM_Inserts_L115_sw53p)
Dim BL_Inserts_L116: BL_Inserts_L116=Array(LM_Inserts_L116_Layer1, LM_Inserts_L116_Parts, LM_Inserts_L116_Playfield, LM_Inserts_L116_sw51p, LM_Inserts_L116_sw80)
Dim BL_Inserts_L117: BL_Inserts_L117=Array(LM_Inserts_L117_Parts, LM_Inserts_L117_Playfield, LM_Inserts_L117_sw50p, LM_Inserts_L117_sw51p)
Dim BL_Inserts_L118: BL_Inserts_L118=Array(LM_Inserts_L118_Parts, LM_Inserts_L118_Playfield, LM_Inserts_L118_sw49p, LM_Inserts_L118_sw50p, LM_Inserts_L118_sw51p)
Dim BL_Inserts_L121: BL_Inserts_L121=Array(LM_Inserts_L121_Bumper1Ring, LM_Inserts_L121_Bumper2Ring, LM_Inserts_L121_Layer2, LM_Inserts_L121_Layer3, LM_Inserts_L121_Parts, LM_Inserts_L121_Playfield, LM_Inserts_L121_sw58p)
Dim BL_Inserts_L126: BL_Inserts_L126=Array(LM_Inserts_L126_Bumper1Ring, LM_Inserts_L126_Layer1, LM_Inserts_L126_Layer2, LM_Inserts_L126_Parts, LM_Inserts_L126_Playfield, LM_Inserts_L126_sw58p)
Dim BL_Inserts_L127: BL_Inserts_L127=Array(LM_Inserts_L127_Bumper1Ring, LM_Inserts_L127_Bumper2Ring, LM_Inserts_L127_Bumper3Ring, LM_Inserts_L127_Layer1, LM_Inserts_L127_Parts, LM_Inserts_L127_Playfield, LM_Inserts_L127_sw49p, LM_Inserts_L127_sw50p, LM_Inserts_L127_sw51p, LM_Inserts_L127_sw58p)
Dim BL_Inserts_L19: BL_Inserts_L19=Array(LM_Inserts_L19_Parts, LM_Inserts_L19_Playfield)
Dim BL_Inserts_L20: BL_Inserts_L20=Array(LM_Inserts_L20_Parts, LM_Inserts_L20_Playfield)
Dim BL_Inserts_L30: BL_Inserts_L30=Array(LM_Inserts_L30_Layer1, LM_Inserts_L30_Layer2, LM_Inserts_L30_Parts, LM_Inserts_L30_Playfield, LM_Inserts_L30_SpinnerPlate)
Dim BL_Inserts_L31: BL_Inserts_L31=Array(LM_Inserts_L31_Layer1, LM_Inserts_L31_Layer2, LM_Inserts_L31_Layer3, LM_Inserts_L31_Parts, LM_Inserts_L31_Playfield, LM_Inserts_L31_Siderails)
Dim BL_Inserts_L32: BL_Inserts_L32=Array(LM_Inserts_L32_Layer1, LM_Inserts_L32_Layer2, LM_Inserts_L32_Layer3, LM_Inserts_L32_Parts, LM_Inserts_L32_Playfield, LM_Inserts_L32_sw26)
Dim BL_Inserts_L41: BL_Inserts_L41=Array(LM_Inserts_L41_Layer1, LM_Inserts_L41_Parts, LM_Inserts_L41_Playfield)
Dim BL_Inserts_L42: BL_Inserts_L42=Array(LM_Inserts_L42_Layer1, LM_Inserts_L42_Parts, LM_Inserts_L42_Playfield, LM_Inserts_L42_sw58p)
Dim BL_Inserts_L43: BL_Inserts_L43=Array(LM_Inserts_L43_Layer1, LM_Inserts_L43_Layer2, LM_Inserts_L43_Parts, LM_Inserts_L43_Playfield, LM_Inserts_L43_sw58p)
Dim BL_Inserts_L49: BL_Inserts_L49=Array(LM_Inserts_L49_DiverterRight, LM_Inserts_L49_GateL, LM_Inserts_L49_Layer2, LM_Inserts_L49_Layer3, LM_Inserts_L49_Parts, LM_Inserts_L49_Playfield, LM_Inserts_L49_dancer, LM_Inserts_L49_sw28)
Dim BL_Inserts_L50: BL_Inserts_L50=Array(LM_Inserts_L50_Layer2, LM_Inserts_L50_Layer3, LM_Inserts_L50_Parts, LM_Inserts_L50_Playfield, LM_Inserts_L50_sw58p)
Dim BL_Inserts_L51: BL_Inserts_L51=Array(LM_Inserts_L51_GateR, LM_Inserts_L51_Layer1, LM_Inserts_L51_Layer2, LM_Inserts_L51_Parts, LM_Inserts_L51_Playfield, LM_Inserts_L51_sw58p)
Dim BL_Inserts_L65: BL_Inserts_L65=Array(LM_Inserts_L65_MissionLockPin, LM_Inserts_L65_Parts, LM_Inserts_L65_Playfield, LM_Inserts_L65_sw17p, LM_Inserts_L65_sw44, LM_Inserts_L65_sw45)
Dim BL_Inserts_L66: BL_Inserts_L66=Array(LM_Inserts_L66_Lsling1, LM_Inserts_L66_Lsling2, LM_Inserts_L66_MissionLockPin, LM_Inserts_L66_Parts, LM_Inserts_L66_Playfield)
Dim BL_Inserts_L67: BL_Inserts_L67=Array(LM_Inserts_L67_LFlipper, LM_Inserts_L67_LFlipperU, LM_Inserts_L67_Parts, LM_Inserts_L67_Playfield)
Dim BL_Inserts_L68: BL_Inserts_L68=Array(LM_Inserts_L68_LFlipper, LM_Inserts_L68_LFlipperU, LM_Inserts_L68_MissionLockPin, LM_Inserts_L68_Parts, LM_Inserts_L68_Playfield, LM_Inserts_L68_RFlipper, LM_Inserts_L68_RFlipperU, LM_Inserts_L68_sw44)
Dim BL_Inserts_L69: BL_Inserts_L69=Array(LM_Inserts_L69_MissionLockPin, LM_Inserts_L69_Parts, LM_Inserts_L69_Playfield)
Dim BL_Inserts_L70: BL_Inserts_L70=Array(LM_Inserts_L70_LFlipperU, LM_Inserts_L70_MissionLockPin, LM_Inserts_L70_Parts, LM_Inserts_L70_Playfield, LM_Inserts_L70_RFlipperU)
Dim BL_Inserts_L71: BL_Inserts_L71=Array(LM_Inserts_L71_LFlipperU, LM_Inserts_L71_Parts, LM_Inserts_L71_Playfield, LM_Inserts_L71_RFlipperU)
Dim BL_Inserts_L72: BL_Inserts_L72=Array(LM_Inserts_L72_LFlipper, LM_Inserts_L72_Playfield, LM_Inserts_L72_RFlipper)
Dim BL_Inserts_L73: BL_Inserts_L73=Array(LM_Inserts_L73_Parts, LM_Inserts_L73_Playfield)
Dim BL_Inserts_L74: BL_Inserts_L74=Array(LM_Inserts_L74_Parts, LM_Inserts_L74_Playfield)
Dim BL_Inserts_L75: BL_Inserts_L75=Array(LM_Inserts_L75_Parts, LM_Inserts_L75_Playfield)
Dim BL_Inserts_L76: BL_Inserts_L76=Array(LM_Inserts_L76_Parts, LM_Inserts_L76_Playfield, LM_Inserts_L76_RPost)
Dim BL_Inserts_L77: BL_Inserts_L77=Array(LM_Inserts_L77_Parts, LM_Inserts_L77_Playfield, LM_Inserts_L77_Rsling1, LM_Inserts_L77_Rsling2)
Dim BL_Inserts_L78: BL_Inserts_L78=Array(LM_Inserts_L78_Parts, LM_Inserts_L78_Playfield, LM_Inserts_L78_SLING1, LM_Inserts_L78_sw65, LM_Inserts_L78_sw66)
Dim BL_Inserts_L79: BL_Inserts_L79=Array(LM_Inserts_L79_Parts, LM_Inserts_L79_Playfield, LM_Inserts_L79_Rsling1, LM_Inserts_L79_Rsling2, LM_Inserts_L79_SLING1)
Dim BL_Inserts_L80: BL_Inserts_L80=Array(LM_Inserts_L80_Parts, LM_Inserts_L80_Playfield, LM_Inserts_L80_RFlipper, LM_Inserts_L80_RFlipperU)
Dim BL_Inserts_L81: BL_Inserts_L81=Array(LM_Inserts_L81_Layer1, LM_Inserts_L81_Parts, LM_Inserts_L81_Playfield, LM_Inserts_L81_sw18p, LM_Inserts_L81_sw19p, LM_Inserts_L81_sw20p)
Dim BL_Inserts_L82: BL_Inserts_L82=Array(LM_Inserts_L82_Layer1, LM_Inserts_L82_Parts, LM_Inserts_L82_Playfield, LM_Inserts_L82_sw17p, LM_Inserts_L82_sw18p, LM_Inserts_L82_sw19p, LM_Inserts_L82_sw20p)
Dim BL_Inserts_L83: BL_Inserts_L83=Array(LM_Inserts_L83_Parts, LM_Inserts_L83_Playfield, LM_Inserts_L83_sw17p, LM_Inserts_L83_sw18p, LM_Inserts_L83_sw19p, LM_Inserts_L83_sw20p)
Dim BL_Inserts_L84: BL_Inserts_L84=Array(LM_Inserts_L84_Parts, LM_Inserts_L84_Playfield, LM_Inserts_L84_sw17p, LM_Inserts_L84_sw18p, LM_Inserts_L84_sw19p, LM_Inserts_L84_sw20p)
Dim BL_Inserts_L85: BL_Inserts_L85=Array(LM_Inserts_L85_Parts, LM_Inserts_L85_Playfield, LM_Inserts_L85_sw17p, LM_Inserts_L85_sw18p, LM_Inserts_L85_sw19p, LM_Inserts_L85_sw20p)
Dim BL_Inserts_L86: BL_Inserts_L86=Array(LM_Inserts_L86_Parts, LM_Inserts_L86_Playfield, LM_Inserts_L86_sw17p, LM_Inserts_L86_sw18p, LM_Inserts_L86_sw19p, LM_Inserts_L86_sw20p)
Dim BL_Inserts_L87: BL_Inserts_L87=Array(LM_Inserts_L87_Parts, LM_Inserts_L87_Playfield, LM_Inserts_L87_sw17p, LM_Inserts_L87_sw18p, LM_Inserts_L87_sw45)
Dim BL_Inserts_L88: BL_Inserts_L88=Array(LM_Inserts_L88_Lsling1, LM_Inserts_L88_Lsling2, LM_Inserts_L88_Parts, LM_Inserts_L88_Playfield, LM_Inserts_L88_sw17p)
Dim BL_Inserts_L89: BL_Inserts_L89=Array(LM_Inserts_L89_Parts, LM_Inserts_L89_Playfield, LM_Inserts_L89_sw17p, LM_Inserts_L89_sw18p, LM_Inserts_L89_sw19p, LM_Inserts_L89_sw20p)
Dim BL_Inserts_L90: BL_Inserts_L90=Array(LM_Inserts_L90_Parts, LM_Inserts_L90_Playfield, LM_Inserts_L90_sw17p, LM_Inserts_L90_sw18p, LM_Inserts_L90_sw19p, LM_Inserts_L90_sw20p)
Dim BL_Inserts_L91: BL_Inserts_L91=Array(LM_Inserts_L91_Parts, LM_Inserts_L91_Playfield, LM_Inserts_L91_sw17p, LM_Inserts_L91_sw18p, LM_Inserts_L91_sw19p)
Dim BL_Inserts_L92: BL_Inserts_L92=Array(LM_Inserts_L92_Parts, LM_Inserts_L92_Playfield)
Dim BL_Inserts_L93: BL_Inserts_L93=Array(LM_Inserts_L93_Parts, LM_Inserts_L93_Playfield, LM_Inserts_L93_sw17p, LM_Inserts_L93_sw18p, LM_Inserts_L93_sw19p, LM_Inserts_L93_sw20p)
Dim BL_Inserts_L94: BL_Inserts_L94=Array(LM_Inserts_L94_Parts, LM_Inserts_L94_Playfield, LM_Inserts_L94_sw17p, LM_Inserts_L94_sw18p, LM_Inserts_L94_sw19p)
Dim BL_Inserts_L95: BL_Inserts_L95=Array(LM_Inserts_L95_Parts, LM_Inserts_L95_Playfield)
Dim BL_Inserts_L97: BL_Inserts_L97=Array(LM_Inserts_L97_Layer1, LM_Inserts_L97_Parts, LM_Inserts_L97_Playfield, LM_Inserts_L97_sw18p, LM_Inserts_L97_sw19p, LM_Inserts_L97_sw20p)
Dim BL_Inserts_L98: BL_Inserts_L98=Array(LM_Inserts_L98_Parts, LM_Inserts_L98_Playfield, LM_Inserts_L98_sw17p, LM_Inserts_L98_sw18p, LM_Inserts_L98_sw19p, LM_Inserts_L98_sw20p, LM_Inserts_L98_sw21p, LM_Inserts_L98_sw49p)
Dim BL_Inserts_L99: BL_Inserts_L99=Array(LM_Inserts_L99_Parts, LM_Inserts_L99_Playfield, LM_Inserts_L99_sw17p, LM_Inserts_L99_sw18p, LM_Inserts_L99_sw19p, LM_Inserts_L99_sw20p, LM_Inserts_L99_sw22p, LM_Inserts_L99_sw49p, LM_Inserts_L99_sw50p, LM_Inserts_L99_sw51p)
Dim BL_Inserts_l57: BL_Inserts_l57=Array(LM_Inserts_l57_Layer1, LM_Inserts_l57_Layer2, LM_Inserts_l57_Parts, LM_Inserts_l57_Spinner1)
Dim BL_Inserts_l58: BL_Inserts_l58=Array(LM_Inserts_l58_Layer1, LM_Inserts_l58_Layer2, LM_Inserts_l58_Layer3, LM_Inserts_l58_Parts)
Dim BL_Inserts_l59: BL_Inserts_l59=Array(LM_Inserts_l59_Layer1, LM_Inserts_l59_Layer2, LM_Inserts_l59_Layer3, LM_Inserts_l59_Parts)
Dim BL_LMods: BL_LMods=Array(LM_LMods_Layer1, LM_LMods_Parts)
Dim BL_Lit_Room: BL_Lit_Room=Array(BM_Alien1, BM_Alien2, BM_Bumper1Ring, BM_Bumper2Ring, BM_Bumper3Ring, BM_DiverterLeft, BM_DiverterRight, BM_GateL, BM_GateR, BM_LFlipper, BM_LFlipperU, BM_Layer1, BM_Layer2, BM_Layer3, BM_Lsling1, BM_Lsling2, BM_MissionLockPin, BM_Parts, BM_Playfield, BM_RFlipper, BM_RFlipper1, BM_RFlipper1U, BM_RFlipperU, BM_RPost, BM_Rsling1, BM_Rsling2, BM_SLING1, BM_SLING2, BM_Siderails, BM_Spinner1, BM_Spinner2, BM_SpinnerPlate, BM_dancer, BM_sw17p, BM_sw18p, BM_sw19p, BM_sw20p, BM_sw21p, BM_sw22p, BM_sw23p, BM_sw26, BM_sw27, BM_sw28, BM_sw29, BM_sw30, BM_sw43, BM_sw44, BM_sw45, BM_sw49p, BM_sw50p, BM_sw51p, BM_sw52p, BM_sw53p, BM_sw58p, BM_sw59, BM_sw60, BM_sw65, BM_sw66, BM_sw77, BM_sw78, BM_sw79, BM_sw80)
' Global arrays
Dim BG_Bakemap: BG_Bakemap=Array(BM_Alien1, BM_Alien2, BM_Bumper1Ring, BM_Bumper2Ring, BM_Bumper3Ring, BM_DiverterLeft, BM_DiverterRight, BM_GateL, BM_GateR, BM_LFlipper, BM_LFlipperU, BM_Layer1, BM_Layer2, BM_Layer3, BM_Lsling1, BM_Lsling2, BM_MissionLockPin, BM_Parts, BM_Playfield, BM_RFlipper, BM_RFlipper1, BM_RFlipper1U, BM_RFlipperU, BM_RPost, BM_Rsling1, BM_Rsling2, BM_SLING1, BM_SLING2, BM_Siderails, BM_Spinner1, BM_Spinner2, BM_SpinnerPlate, BM_dancer, BM_sw17p, BM_sw18p, BM_sw19p, BM_sw20p, BM_sw21p, BM_sw22p, BM_sw23p, BM_sw26, BM_sw27, BM_sw28, BM_sw29, BM_sw30, BM_sw43, BM_sw44, BM_sw45, BM_sw49p, BM_sw50p, BM_sw51p, BM_sw52p, BM_sw53p, BM_sw58p, BM_sw59, BM_sw60, BM_sw65, BM_sw66, BM_sw77, BM_sw78, BM_sw79, BM_sw80)
Dim BG_Lightmap: BG_Lightmap=Array(LM_Flashers_F21_DiverterLeft, LM_Flashers_F21_DiverterRight, LM_Flashers_F21_Layer1, LM_Flashers_F21_Layer2, LM_Flashers_F21_Layer3, LM_Flashers_F21_Parts, LM_Flashers_F21_Playfield, LM_Flashers_F21_Siderails, LM_Flashers_F21_dancer, LM_Flashers_F22_Alien1, LM_Flashers_F22_Alien2, LM_Flashers_F22_Bumper1Ring, LM_Flashers_F22_Bumper2Ring, LM_Flashers_F22_Bumper3Ring, LM_Flashers_F22_DiverterLeft, LM_Flashers_F22_DiverterRight, LM_Flashers_F22_GateL, LM_Flashers_F22_GateR, LM_Flashers_F22_Layer1, LM_Flashers_F22_Layer2, LM_Flashers_F22_Layer3, LM_Flashers_F22_Parts, LM_Flashers_F22_Playfield, LM_Flashers_F22_Spinner1, LM_Flashers_F22_Spinner2, LM_Flashers_F22_SpinnerPlate, LM_Flashers_F22_dancer, LM_Flashers_F22_sw26, LM_Flashers_F22_sw27, LM_Flashers_F22_sw28, LM_Flashers_F22_sw29, LM_Flashers_F22_sw30, LM_Flashers_F22_sw49p, LM_Flashers_F22_sw50p, LM_Flashers_F22_sw51p, LM_Flashers_F22_sw52p, LM_Flashers_F22_sw53p, LM_Flashers_F22_sw58p, LM_Flashers_F23_Bumper1Ring, _
  LM_Flashers_F23_Bumper2Ring, LM_Flashers_F23_Bumper3Ring, LM_Flashers_F23_DiverterRight, LM_Flashers_F23_GateL, LM_Flashers_F23_GateR, LM_Flashers_F23_Layer1, LM_Flashers_F23_Layer2, LM_Flashers_F23_Layer3, LM_Flashers_F23_Parts, LM_Flashers_F23_Playfield, LM_Flashers_F23_Spinner1, LM_Flashers_F23_Spinner2, LM_Flashers_F23_dancer, LM_Flashers_F23_sw23p, LM_Flashers_F23_sw28, LM_Flashers_F23_sw29, LM_Flashers_F23_sw30, LM_Flashers_F23_sw49p, LM_Flashers_F23_sw50p, LM_Flashers_F23_sw51p, LM_Flashers_F23_sw52p, LM_Flashers_F23_sw53p, LM_Flashers_F23_sw58p, LM_Flashers_F23_sw59, LM_Flashers_F24_Bumper1Ring, LM_Flashers_F24_Layer1, LM_Flashers_F24_Parts, LM_Flashers_F24_Playfield, LM_Flashers_F24_RFlipper1, LM_Flashers_F24_RFlipper1U, LM_Flashers_F24_sw49p, LM_Flashers_F24_sw50p, LM_Flashers_F24_sw51p, LM_Flashers_F24_sw52p, LM_Flashers_F24_sw53p, LM_Flashers_F25_Alien1, LM_Flashers_F25_Alien2, LM_Flashers_F25_Layer1, LM_Flashers_F25_Layer2, LM_Flashers_F25_Parts, LM_Flashers_F25_Playfield, LM_Flashers_F25_sw60, _
  LM_Flashers_F26_LFlipper, LM_Flashers_F26_LFlipperU, LM_Flashers_F26_Layer1, LM_Flashers_F26_Lsling1, LM_Flashers_F26_Lsling2, LM_Flashers_F26_MissionLockPin, LM_Flashers_F26_Parts, LM_Flashers_F26_Playfield, LM_Flashers_F26_RFlipper, LM_Flashers_F26_RFlipperU, LM_Flashers_F26_Rsling1, LM_Flashers_F26_Rsling2, LM_Flashers_F26_SLING1, LM_Flashers_F26_SLING2, LM_Flashers_F26_sw17p, LM_Flashers_F26_sw18p, LM_Flashers_F26_sw44, LM_Flashers_F26_sw45, LM_Flashers_L01_Parts, LM_Flashers_L104_Bumper2Ring, LM_Flashers_L104_Bumper3Ring, LM_Flashers_L104_Parts, LM_Flashers_L104_Playfield, LM_Flashers_L119_Layer1, LM_Flashers_L119_Parts, LM_Flashers_L119_Playfield, LM_Flashers_L120_Layer1, LM_Flashers_L120_Parts, LM_Flashers_L120_Playfield, LM_Flashers_L120_sw50p, LM_Flashers_L120_sw51p, LM_Flashers_L120_sw80, LM_Flashers_L44_Alien1, LM_Flashers_L44_Layer1, LM_Flashers_L44_Layer2, LM_Flashers_L44_Parts, LM_Flashers_L45_Layer1, LM_Flashers_L45_Layer2, LM_Flashers_L45_Parts, LM_Flashers_L45_Playfield, _
  LM_Flashers_L52_Bumper2Ring, LM_Flashers_L52_Layer1, LM_Flashers_L52_Layer2, LM_Flashers_L52_Layer3, LM_Flashers_L52_Parts, LM_Flashers_L52_Playfield, LM_Flashers_L52_dancer, LM_Flashers_L52_sw28, LM_Flashers_L52_sw30, LM_Flashers_L53_Bumper2Ring, LM_Flashers_L53_GateL, LM_Flashers_L53_Layer1, LM_Flashers_L53_Layer2, LM_Flashers_L53_Layer3, LM_Flashers_L53_Parts, LM_Flashers_L53_Playfield, LM_Flashers_L53_dancer, LM_Flashers_L53_sw28, LM_Flashers_L53_sw30, LM_Flashers_L54_Bumper2Ring, LM_Flashers_L54_Bumper3Ring, LM_Flashers_L54_GateL, LM_Flashers_L54_GateR, LM_Flashers_L54_Layer1, LM_Flashers_L54_Layer2, LM_Flashers_L54_Layer3, LM_Flashers_L54_Parts, LM_Flashers_L54_Playfield, LM_Flashers_L54_dancer, LM_Flashers_L54_sw28, LM_Flashers_L54_sw29, LM_Flashers_L54_sw30, LM_Flashers_L54_sw58p, LM_Flashers_L62_Alien1, LM_Flashers_L62_DiverterRight, LM_Flashers_L62_Layer1, LM_Flashers_L62_Layer2, LM_Flashers_L62_Layer3, LM_Flashers_L62_Parts, LM_Flashers_L62_Playfield, LM_Flashers_L62_Spinner2, LM_Flashers_L62_sw30, _
  LM_Gi_L10_Layer1, LM_Gi_L10_Parts, LM_Gi_L10_Playfield, LM_Gi_L10_sw17p, LM_Gi_L10_sw18p, LM_Gi_L10_sw19p, LM_Gi_L10_sw20p, LM_Gi_L11_Parts, LM_Gi_L11_Playfield, LM_Gi_L11_sw17p, LM_Gi_L11_sw18p, LM_Gi_L11_sw19p, LM_Gi_L11_sw20p, LM_Gi_L12_LFlipperU, LM_Gi_L12_Lsling1, LM_Gi_L12_Lsling2, LM_Gi_L12_MissionLockPin, LM_Gi_L12_Parts, LM_Gi_L12_Playfield, LM_Gi_L12_SLING2, LM_Gi_L12_sw17p, LM_Gi_L12_sw18p, LM_Gi_L12_sw19p, LM_Gi_L12_sw20p, LM_Gi_L12_sw44, LM_Gi_L12_sw45, LM_Gi_L122_Bumper1Ring, LM_Gi_L122_Bumper2Ring, LM_Gi_L122_Layer1, LM_Gi_L122_Layer3, LM_Gi_L122_Parts, LM_Gi_L122_Playfield, LM_Gi_L122_dancer, LM_Gi_L122_sw49p, LM_Gi_L122_sw50p, LM_Gi_L122_sw51p, LM_Gi_L122_sw52p, LM_Gi_L122_sw53p, LM_Gi_L123_Bumper1Ring, LM_Gi_L123_Bumper2Ring, LM_Gi_L123_Bumper3Ring, LM_Gi_L123_Layer1, LM_Gi_L123_Layer2, LM_Gi_L123_Layer3, LM_Gi_L123_Parts, LM_Gi_L123_Playfield, LM_Gi_L123_sw50p, LM_Gi_L123_sw51p, LM_Gi_L123_sw53p, LM_Gi_L124_Bumper1Ring, LM_Gi_L124_Layer1, LM_Gi_L124_Layer2, LM_Gi_L124_Layer3, _
  LM_Gi_L124_Parts, LM_Gi_L124_Playfield, LM_Gi_L124_sw30, LM_Gi_L124_sw58p, LM_Gi_L128_Layer1, LM_Gi_L128_Layer3, LM_Gi_L128_Parts, LM_Gi_L128_Playfield, LM_Gi_L128_RFlipper1, LM_Gi_L128_RFlipper1U, LM_Gi_L14_LFlipper, LM_Gi_L14_LFlipperU, LM_Gi_L14_Lsling1, LM_Gi_L14_Lsling2, LM_Gi_L14_MissionLockPin, LM_Gi_L14_Parts, LM_Gi_L14_Playfield, LM_Gi_L14_RFlipper, LM_Gi_L14_SLING2, LM_Gi_L14_sw17p, LM_Gi_L14_sw18p, LM_Gi_L14_sw19p, LM_Gi_L14_sw44, LM_Gi_L14_sw45, LM_Gi_L17_Parts, LM_Gi_L17_Playfield, LM_Gi_L17_RFlipper1, LM_Gi_L17_RFlipper1U, LM_Gi_L18_Parts, LM_Gi_L18_Playfield, LM_Gi_L18_RPost, LM_Gi_L21_Parts, LM_Gi_L21_Playfield, LM_Gi_L21_RFlipper, LM_Gi_L21_RFlipperU, LM_Gi_L21_RPost, LM_Gi_L21_Rsling1, LM_Gi_L21_Rsling2, LM_Gi_L21_SLING1, LM_Gi_L21_sw17p, LM_Gi_L21_sw65, LM_Gi_L21_sw66, LM_Gi_L22_MissionLockPin, LM_Gi_L22_Parts, LM_Gi_L22_Playfield, LM_Gi_L22_RFlipper, LM_Gi_L22_RFlipperU, LM_Gi_L22_RPost, LM_Gi_L22_Rsling1, LM_Gi_L22_Rsling2, LM_Gi_L22_SLING1, LM_Gi_L22_sw65, LM_Gi_L22_sw66, _
  LM_Gi_L23_MissionLockPin, LM_Gi_L23_Parts, LM_Gi_L23_Playfield, LM_Gi_L23_RFlipper, LM_Gi_L23_RFlipperU, LM_Gi_L23_RPost, LM_Gi_L23_Rsling1, LM_Gi_L23_Rsling2, LM_Gi_L23_SLING1, LM_Gi_L23_sw65, LM_Gi_L23_sw66, LM_Gi_L24_MissionLockPin, LM_Gi_L24_Parts, LM_Gi_L24_Playfield, LM_Gi_L24_RFlipper, LM_Gi_L24_RFlipperU, LM_Gi_L24_Rsling1, LM_Gi_L24_Rsling2, LM_Gi_L24_SLING1, LM_Gi_L24_sw44, LM_Gi_L24_sw65, LM_Gi_L24_sw66, LM_Gi_L25_Layer1, LM_Gi_L25_Layer2, LM_Gi_L25_Layer3, LM_Gi_L25_Parts, LM_Gi_L25_Playfield, LM_Gi_L25_SpinnerPlate, LM_Gi_L25_sw23p, LM_Gi_L25_sw49p, LM_Gi_L25_sw50p, LM_Gi_L26_Bumper1Ring, LM_Gi_L26_Bumper2Ring, LM_Gi_L26_Bumper3Ring, LM_Gi_L26_Layer1, LM_Gi_L26_Layer2, LM_Gi_L26_Layer3, LM_Gi_L26_Parts, LM_Gi_L26_Playfield, LM_Gi_L26_sw49p, LM_Gi_L26_sw50p, LM_Gi_L26_sw51p, LM_Gi_L26_sw52p, LM_Gi_L26_sw53p, LM_Gi_L26_sw58p, LM_Gi_L27_Bumper2Ring, LM_Gi_L27_Bumper3Ring, LM_Gi_L27_Layer1, LM_Gi_L27_Layer2, LM_Gi_L27_Layer3, LM_Gi_L27_Parts, LM_Gi_L27_Playfield, LM_Gi_L27_sw27, LM_Gi_L27_sw49p, _
  LM_Gi_L27_sw50p, LM_Gi_L27_sw51p, LM_Gi_L27_sw52p, LM_Gi_L27_sw53p, LM_Gi_L28_Layer1, LM_Gi_L28_Layer2, LM_Gi_L28_Layer3, LM_Gi_L28_Parts, LM_Gi_L28_Playfield, LM_Gi_L28_dancer, LM_Gi_L28_sw26, LM_Gi_L29_DiverterLeft, LM_Gi_L29_DiverterRight, LM_Gi_L29_Layer1, LM_Gi_L29_Layer2, LM_Gi_L29_Layer3, LM_Gi_L29_Parts, LM_Gi_L29_Playfield, LM_Gi_L29_dancer, LM_Gi_L33_Bumper1Ring, LM_Gi_L33_Layer1, LM_Gi_L33_Layer2, LM_Gi_L33_Layer3, LM_Gi_L33_Parts, LM_Gi_L33_Playfield, LM_Gi_L33_dancer, LM_Gi_L33_sw28, LM_Gi_L34_Bumper1Ring, LM_Gi_L34_Bumper2Ring, LM_Gi_L34_GateL, LM_Gi_L34_Layer1, LM_Gi_L34_Layer2, LM_Gi_L34_Layer3, LM_Gi_L34_Parts, LM_Gi_L34_Playfield, LM_Gi_L34_dancer, LM_Gi_L34_sw28, LM_Gi_L34_sw29, LM_Gi_L35_Alien1, LM_Gi_L35_Bumper1Ring, LM_Gi_L35_Bumper2Ring, LM_Gi_L35_GateL, LM_Gi_L35_GateR, LM_Gi_L35_Layer1, LM_Gi_L35_Layer2, LM_Gi_L35_Layer3, LM_Gi_L35_Parts, LM_Gi_L35_Playfield, LM_Gi_L35_sw29, LM_Gi_L35_sw30, LM_Gi_L35_sw58p, LM_Gi_L36_Alien1, LM_Gi_L36_Bumper2Ring, LM_Gi_L36_Layer1, LM_Gi_L36_Layer2, _
  LM_Gi_L36_Layer3, LM_Gi_L36_Parts, LM_Gi_L36_Playfield, LM_Gi_L36_Spinner2, LM_Gi_L36_sw30, LM_Gi_L36_sw58p, LM_Gi_L37_Alien1, LM_Gi_L37_Alien2, LM_Gi_L37_Layer1, LM_Gi_L37_Layer2, LM_Gi_L37_Parts, LM_Gi_L37_Playfield, LM_Gi_L37_sw30, LM_Gi_L37_sw58p, LM_Gi_L39_Alien2, LM_Gi_L39_Layer1, LM_Gi_L39_Layer2, LM_Gi_L39_Parts, LM_Gi_L39_Playfield, LM_Gi_L40_Bumper1Ring, LM_Gi_L40_Bumper2Ring, LM_Gi_L40_Layer1, LM_Gi_L40_Layer3, LM_Gi_L40_Parts, LM_Gi_L40_Playfield, LM_Gi_L40_sw58p, LM_Gi_L40_sw59, LM_Gi_L9_Layer1, LM_Gi_L9_MissionLockPin, LM_Gi_L9_Parts, LM_Gi_L9_Playfield, LM_Gi_L9_sw17p, LM_Gi_L9_sw18p, LM_Gi_L9_sw19p, LM_Gi_L9_sw20p, LM_Inserts_L100_Parts, LM_Inserts_L100_Playfield, LM_Inserts_L100_sw18p, LM_Inserts_L100_sw19p, LM_Inserts_L100_sw20p, LM_Inserts_L100_sw22p, LM_Inserts_L100_sw23p, LM_Inserts_L100_sw49p, LM_Inserts_L100_sw50p, LM_Inserts_L100_sw51p, LM_Inserts_L101_GateR, LM_Inserts_L101_Layer2, LM_Inserts_L101_Parts, LM_Inserts_L101_Playfield, LM_Inserts_L101_sw23p, LM_Inserts_L101_sw49p, _
  LM_Inserts_L101_sw50p, LM_Inserts_L101_sw51p, LM_Inserts_L101_sw52p, LM_Inserts_L102_Layer1, LM_Inserts_L102_Parts, LM_Inserts_L102_Playfield, LM_Inserts_L102_sw23p, LM_Inserts_L102_sw49p, LM_Inserts_L102_sw50p, LM_Inserts_L102_sw51p, LM_Inserts_L102_sw52p, LM_Inserts_L102_sw53p, LM_Inserts_L103_Layer1, LM_Inserts_L103_Parts, LM_Inserts_L103_Playfield, LM_Inserts_L103_Spinner1, LM_Inserts_L103_sw49p, LM_Inserts_L103_sw50p, LM_Inserts_L103_sw51p, LM_Inserts_L103_sw52p, LM_Inserts_L105_Layer1, LM_Inserts_L105_Parts, LM_Inserts_L105_Playfield, LM_Inserts_L105_sw49p, LM_Inserts_L105_sw50p, LM_Inserts_L105_sw51p, LM_Inserts_L105_sw53p, LM_Inserts_L106_Parts, LM_Inserts_L106_Playfield, LM_Inserts_L106_sw49p, LM_Inserts_L106_sw50p, LM_Inserts_L106_sw51p, LM_Inserts_L106_sw53p, LM_Inserts_L107_Parts, LM_Inserts_L107_Playfield, LM_Inserts_L107_sw49p, LM_Inserts_L107_sw50p, LM_Inserts_L107_sw51p, LM_Inserts_L108_Parts, LM_Inserts_L108_Playfield, LM_Inserts_L108_sw49p, LM_Inserts_L108_sw50p, LM_Inserts_L108_sw51p, _
  LM_Inserts_L109_Parts, LM_Inserts_L109_Playfield, LM_Inserts_L109_sw49p, LM_Inserts_L109_sw50p, LM_Inserts_L109_sw51p, LM_Inserts_L110_Playfield, LM_Inserts_L110_sw49p, LM_Inserts_L110_sw50p, LM_Inserts_L110_sw51p, LM_Inserts_L111_Parts, LM_Inserts_L111_Playfield, LM_Inserts_L111_sw49p, LM_Inserts_L111_sw50p, LM_Inserts_L111_sw51p, LM_Inserts_L112_Parts, LM_Inserts_L112_Playfield, LM_Inserts_L112_sw18p, LM_Inserts_L112_sw49p, LM_Inserts_L112_sw50p, LM_Inserts_L112_sw51p, LM_Inserts_L113_Parts, LM_Inserts_L113_Playfield, LM_Inserts_L113_sw49p, LM_Inserts_L113_sw50p, LM_Inserts_L113_sw51p, LM_Inserts_L113_sw52p, LM_Inserts_L113_sw53p, LM_Inserts_L114_Parts, LM_Inserts_L114_Playfield, LM_Inserts_L114_sw49p, LM_Inserts_L114_sw50p, LM_Inserts_L114_sw51p, LM_Inserts_L114_sw52p, LM_Inserts_L114_sw53p, LM_Inserts_L115_Layer1, LM_Inserts_L115_Parts, LM_Inserts_L115_Playfield, LM_Inserts_L115_sw49p, LM_Inserts_L115_sw50p, LM_Inserts_L115_sw51p, LM_Inserts_L115_sw52p, LM_Inserts_L115_sw53p, LM_Inserts_L116_Layer1, _
  LM_Inserts_L116_Parts, LM_Inserts_L116_Playfield, LM_Inserts_L116_sw51p, LM_Inserts_L116_sw80, LM_Inserts_L117_Parts, LM_Inserts_L117_Playfield, LM_Inserts_L117_sw50p, LM_Inserts_L117_sw51p, LM_Inserts_L118_Parts, LM_Inserts_L118_Playfield, LM_Inserts_L118_sw49p, LM_Inserts_L118_sw50p, LM_Inserts_L118_sw51p, LM_Inserts_L121_Bumper1Ring, LM_Inserts_L121_Bumper2Ring, LM_Inserts_L121_Layer2, LM_Inserts_L121_Layer3, LM_Inserts_L121_Parts, LM_Inserts_L121_Playfield, LM_Inserts_L121_sw58p, LM_Inserts_L126_Bumper1Ring, LM_Inserts_L126_Layer1, LM_Inserts_L126_Layer2, LM_Inserts_L126_Parts, LM_Inserts_L126_Playfield, LM_Inserts_L126_sw58p, LM_Inserts_L127_Bumper1Ring, LM_Inserts_L127_Bumper2Ring, LM_Inserts_L127_Bumper3Ring, LM_Inserts_L127_Layer1, LM_Inserts_L127_Parts, LM_Inserts_L127_Playfield, LM_Inserts_L127_sw49p, LM_Inserts_L127_sw50p, LM_Inserts_L127_sw51p, LM_Inserts_L127_sw58p, LM_Inserts_L19_Parts, LM_Inserts_L19_Playfield, LM_Inserts_L20_Parts, LM_Inserts_L20_Playfield, LM_Inserts_L30_Layer1, _
  LM_Inserts_L30_Layer2, LM_Inserts_L30_Parts, LM_Inserts_L30_Playfield, LM_Inserts_L30_SpinnerPlate, LM_Inserts_L31_Layer1, LM_Inserts_L31_Layer2, LM_Inserts_L31_Layer3, LM_Inserts_L31_Parts, LM_Inserts_L31_Playfield, LM_Inserts_L31_Siderails, LM_Inserts_L32_Layer1, LM_Inserts_L32_Layer2, LM_Inserts_L32_Layer3, LM_Inserts_L32_Parts, LM_Inserts_L32_Playfield, LM_Inserts_L32_sw26, LM_Inserts_L41_Layer1, LM_Inserts_L41_Parts, LM_Inserts_L41_Playfield, LM_Inserts_L42_Layer1, LM_Inserts_L42_Parts, LM_Inserts_L42_Playfield, LM_Inserts_L42_sw58p, LM_Inserts_L43_Layer1, LM_Inserts_L43_Layer2, LM_Inserts_L43_Parts, LM_Inserts_L43_Playfield, LM_Inserts_L43_sw58p, LM_Inserts_L49_DiverterRight, LM_Inserts_L49_GateL, LM_Inserts_L49_Layer2, LM_Inserts_L49_Layer3, LM_Inserts_L49_Parts, LM_Inserts_L49_Playfield, LM_Inserts_L49_dancer, LM_Inserts_L49_sw28, LM_Inserts_L50_Layer2, LM_Inserts_L50_Layer3, LM_Inserts_L50_Parts, LM_Inserts_L50_Playfield, LM_Inserts_L50_sw58p, LM_Inserts_L51_GateR, LM_Inserts_L51_Layer1, _
  LM_Inserts_L51_Layer2, LM_Inserts_L51_Parts, LM_Inserts_L51_Playfield, LM_Inserts_L51_sw58p, LM_Inserts_L65_MissionLockPin, LM_Inserts_L65_Parts, LM_Inserts_L65_Playfield, LM_Inserts_L65_sw17p, LM_Inserts_L65_sw44, LM_Inserts_L65_sw45, LM_Inserts_L66_Lsling1, LM_Inserts_L66_Lsling2, LM_Inserts_L66_MissionLockPin, LM_Inserts_L66_Parts, LM_Inserts_L66_Playfield, LM_Inserts_L67_LFlipper, LM_Inserts_L67_LFlipperU, LM_Inserts_L67_Parts, LM_Inserts_L67_Playfield, LM_Inserts_L68_LFlipper, LM_Inserts_L68_LFlipperU, LM_Inserts_L68_MissionLockPin, LM_Inserts_L68_Parts, LM_Inserts_L68_Playfield, LM_Inserts_L68_RFlipper, LM_Inserts_L68_RFlipperU, LM_Inserts_L68_sw44, LM_Inserts_L69_MissionLockPin, LM_Inserts_L69_Parts, LM_Inserts_L69_Playfield, LM_Inserts_L70_LFlipperU, LM_Inserts_L70_MissionLockPin, LM_Inserts_L70_Parts, LM_Inserts_L70_Playfield, LM_Inserts_L70_RFlipperU, LM_Inserts_L71_LFlipperU, LM_Inserts_L71_Parts, LM_Inserts_L71_Playfield, LM_Inserts_L71_RFlipperU, LM_Inserts_L72_LFlipper, LM_Inserts_L72_Playfield, _
  LM_Inserts_L72_RFlipper, LM_Inserts_L73_Parts, LM_Inserts_L73_Playfield, LM_Inserts_L74_Parts, LM_Inserts_L74_Playfield, LM_Inserts_L75_Parts, LM_Inserts_L75_Playfield, LM_Inserts_L76_Parts, LM_Inserts_L76_Playfield, LM_Inserts_L76_RPost, LM_Inserts_L77_Parts, LM_Inserts_L77_Playfield, LM_Inserts_L77_Rsling1, LM_Inserts_L77_Rsling2, LM_Inserts_L78_Parts, LM_Inserts_L78_Playfield, LM_Inserts_L78_SLING1, LM_Inserts_L78_sw65, LM_Inserts_L78_sw66, LM_Inserts_L79_Parts, LM_Inserts_L79_Playfield, LM_Inserts_L79_Rsling1, LM_Inserts_L79_Rsling2, LM_Inserts_L79_SLING1, LM_Inserts_L80_Parts, LM_Inserts_L80_Playfield, LM_Inserts_L80_RFlipper, LM_Inserts_L80_RFlipperU, LM_Inserts_L81_Layer1, LM_Inserts_L81_Parts, LM_Inserts_L81_Playfield, LM_Inserts_L81_sw18p, LM_Inserts_L81_sw19p, LM_Inserts_L81_sw20p, LM_Inserts_L82_Layer1, LM_Inserts_L82_Parts, LM_Inserts_L82_Playfield, LM_Inserts_L82_sw17p, LM_Inserts_L82_sw18p, LM_Inserts_L82_sw19p, LM_Inserts_L82_sw20p, LM_Inserts_L83_Parts, LM_Inserts_L83_Playfield, _
  LM_Inserts_L83_sw17p, LM_Inserts_L83_sw18p, LM_Inserts_L83_sw19p, LM_Inserts_L83_sw20p, LM_Inserts_L84_Parts, LM_Inserts_L84_Playfield, LM_Inserts_L84_sw17p, LM_Inserts_L84_sw18p, LM_Inserts_L84_sw19p, LM_Inserts_L84_sw20p, LM_Inserts_L85_Parts, LM_Inserts_L85_Playfield, LM_Inserts_L85_sw17p, LM_Inserts_L85_sw18p, LM_Inserts_L85_sw19p, LM_Inserts_L85_sw20p, LM_Inserts_L86_Parts, LM_Inserts_L86_Playfield, LM_Inserts_L86_sw17p, LM_Inserts_L86_sw18p, LM_Inserts_L86_sw19p, LM_Inserts_L86_sw20p, LM_Inserts_L87_Parts, LM_Inserts_L87_Playfield, LM_Inserts_L87_sw17p, LM_Inserts_L87_sw18p, LM_Inserts_L87_sw45, LM_Inserts_L88_Lsling1, LM_Inserts_L88_Lsling2, LM_Inserts_L88_Parts, LM_Inserts_L88_Playfield, LM_Inserts_L88_sw17p, LM_Inserts_L89_Parts, LM_Inserts_L89_Playfield, LM_Inserts_L89_sw17p, LM_Inserts_L89_sw18p, LM_Inserts_L89_sw19p, LM_Inserts_L89_sw20p, LM_Inserts_L90_Parts, LM_Inserts_L90_Playfield, LM_Inserts_L90_sw17p, LM_Inserts_L90_sw18p, LM_Inserts_L90_sw19p, LM_Inserts_L90_sw20p, LM_Inserts_L91_Parts, _
  LM_Inserts_L91_Playfield, LM_Inserts_L91_sw17p, LM_Inserts_L91_sw18p, LM_Inserts_L91_sw19p, LM_Inserts_L92_Parts, LM_Inserts_L92_Playfield, LM_Inserts_L93_Parts, LM_Inserts_L93_Playfield, LM_Inserts_L93_sw17p, LM_Inserts_L93_sw18p, LM_Inserts_L93_sw19p, LM_Inserts_L93_sw20p, LM_Inserts_L94_Parts, LM_Inserts_L94_Playfield, LM_Inserts_L94_sw17p, LM_Inserts_L94_sw18p, LM_Inserts_L94_sw19p, LM_Inserts_L95_Parts, LM_Inserts_L95_Playfield, LM_Inserts_L97_Layer1, LM_Inserts_L97_Parts, LM_Inserts_L97_Playfield, LM_Inserts_L97_sw18p, LM_Inserts_L97_sw19p, LM_Inserts_L97_sw20p, LM_Inserts_L98_Parts, LM_Inserts_L98_Playfield, LM_Inserts_L98_sw17p, LM_Inserts_L98_sw18p, LM_Inserts_L98_sw19p, LM_Inserts_L98_sw20p, LM_Inserts_L98_sw21p, LM_Inserts_L98_sw49p, LM_Inserts_L99_Parts, LM_Inserts_L99_Playfield, LM_Inserts_L99_sw17p, LM_Inserts_L99_sw18p, LM_Inserts_L99_sw19p, LM_Inserts_L99_sw20p, LM_Inserts_L99_sw22p, LM_Inserts_L99_sw49p, LM_Inserts_L99_sw50p, LM_Inserts_L99_sw51p, LM_Inserts_l57_Layer1, LM_Inserts_l57_Layer2, _
  LM_Inserts_l57_Parts, LM_Inserts_l57_Spinner1, LM_Inserts_l58_Layer1, LM_Inserts_l58_Layer2, LM_Inserts_l58_Layer3, LM_Inserts_l58_Parts, LM_Inserts_l59_Layer1, LM_Inserts_l59_Layer2, LM_Inserts_l59_Layer3, LM_Inserts_l59_Parts, LM_LMods_Layer1, LM_LMods_Parts)
Dim BG_All: BG_All=Array(BM_Alien1, BM_Alien2, BM_Bumper1Ring, BM_Bumper2Ring, BM_Bumper3Ring, BM_DiverterLeft, BM_DiverterRight, BM_GateL, BM_GateR, BM_LFlipper, BM_LFlipperU, BM_Layer1, BM_Layer2, BM_Layer3, BM_Lsling1, BM_Lsling2, BM_MissionLockPin, BM_Parts, BM_Playfield, BM_RFlipper, BM_RFlipper1, BM_RFlipper1U, BM_RFlipperU, BM_RPost, BM_Rsling1, BM_Rsling2, BM_SLING1, BM_SLING2, BM_Siderails, BM_Spinner1, BM_Spinner2, BM_SpinnerPlate, BM_dancer, BM_sw17p, BM_sw18p, BM_sw19p, BM_sw20p, BM_sw21p, BM_sw22p, BM_sw23p, BM_sw26, BM_sw27, BM_sw28, BM_sw29, BM_sw30, BM_sw43, BM_sw44, BM_sw45, BM_sw49p, BM_sw50p, BM_sw51p, BM_sw52p, BM_sw53p, BM_sw58p, BM_sw59, BM_sw60, BM_sw65, BM_sw66, BM_sw77, BM_sw78, BM_sw79, BM_sw80, LM_Flashers_F21_DiverterLeft, LM_Flashers_F21_DiverterRight, LM_Flashers_F21_Layer1, LM_Flashers_F21_Layer2, LM_Flashers_F21_Layer3, LM_Flashers_F21_Parts, LM_Flashers_F21_Playfield, LM_Flashers_F21_Siderails, LM_Flashers_F21_dancer, LM_Flashers_F22_Alien1, LM_Flashers_F22_Alien2, _
  LM_Flashers_F22_Bumper1Ring, LM_Flashers_F22_Bumper2Ring, LM_Flashers_F22_Bumper3Ring, LM_Flashers_F22_DiverterLeft, LM_Flashers_F22_DiverterRight, LM_Flashers_F22_GateL, LM_Flashers_F22_GateR, LM_Flashers_F22_Layer1, LM_Flashers_F22_Layer2, LM_Flashers_F22_Layer3, LM_Flashers_F22_Parts, LM_Flashers_F22_Playfield, LM_Flashers_F22_Spinner1, LM_Flashers_F22_Spinner2, LM_Flashers_F22_SpinnerPlate, LM_Flashers_F22_dancer, LM_Flashers_F22_sw26, LM_Flashers_F22_sw27, LM_Flashers_F22_sw28, LM_Flashers_F22_sw29, LM_Flashers_F22_sw30, LM_Flashers_F22_sw49p, LM_Flashers_F22_sw50p, LM_Flashers_F22_sw51p, LM_Flashers_F22_sw52p, LM_Flashers_F22_sw53p, LM_Flashers_F22_sw58p, LM_Flashers_F23_Bumper1Ring, LM_Flashers_F23_Bumper2Ring, LM_Flashers_F23_Bumper3Ring, LM_Flashers_F23_DiverterRight, LM_Flashers_F23_GateL, LM_Flashers_F23_GateR, LM_Flashers_F23_Layer1, LM_Flashers_F23_Layer2, LM_Flashers_F23_Layer3, LM_Flashers_F23_Parts, LM_Flashers_F23_Playfield, LM_Flashers_F23_Spinner1, LM_Flashers_F23_Spinner2, _
  LM_Flashers_F23_dancer, LM_Flashers_F23_sw23p, LM_Flashers_F23_sw28, LM_Flashers_F23_sw29, LM_Flashers_F23_sw30, LM_Flashers_F23_sw49p, LM_Flashers_F23_sw50p, LM_Flashers_F23_sw51p, LM_Flashers_F23_sw52p, LM_Flashers_F23_sw53p, LM_Flashers_F23_sw58p, LM_Flashers_F23_sw59, LM_Flashers_F24_Bumper1Ring, LM_Flashers_F24_Layer1, LM_Flashers_F24_Parts, LM_Flashers_F24_Playfield, LM_Flashers_F24_RFlipper1, LM_Flashers_F24_RFlipper1U, LM_Flashers_F24_sw49p, LM_Flashers_F24_sw50p, LM_Flashers_F24_sw51p, LM_Flashers_F24_sw52p, LM_Flashers_F24_sw53p, LM_Flashers_F25_Alien1, LM_Flashers_F25_Alien2, LM_Flashers_F25_Layer1, LM_Flashers_F25_Layer2, LM_Flashers_F25_Parts, LM_Flashers_F25_Playfield, LM_Flashers_F25_sw60, LM_Flashers_F26_LFlipper, LM_Flashers_F26_LFlipperU, LM_Flashers_F26_Layer1, LM_Flashers_F26_Lsling1, LM_Flashers_F26_Lsling2, LM_Flashers_F26_MissionLockPin, LM_Flashers_F26_Parts, LM_Flashers_F26_Playfield, LM_Flashers_F26_RFlipper, LM_Flashers_F26_RFlipperU, LM_Flashers_F26_Rsling1, _
  LM_Flashers_F26_Rsling2, LM_Flashers_F26_SLING1, LM_Flashers_F26_SLING2, LM_Flashers_F26_sw17p, LM_Flashers_F26_sw18p, LM_Flashers_F26_sw44, LM_Flashers_F26_sw45, LM_Flashers_L01_Parts, LM_Flashers_L104_Bumper2Ring, LM_Flashers_L104_Bumper3Ring, LM_Flashers_L104_Parts, LM_Flashers_L104_Playfield, LM_Flashers_L119_Layer1, LM_Flashers_L119_Parts, LM_Flashers_L119_Playfield, LM_Flashers_L120_Layer1, LM_Flashers_L120_Parts, LM_Flashers_L120_Playfield, LM_Flashers_L120_sw50p, LM_Flashers_L120_sw51p, LM_Flashers_L120_sw80, LM_Flashers_L44_Alien1, LM_Flashers_L44_Layer1, LM_Flashers_L44_Layer2, LM_Flashers_L44_Parts, LM_Flashers_L45_Layer1, LM_Flashers_L45_Layer2, LM_Flashers_L45_Parts, LM_Flashers_L45_Playfield, LM_Flashers_L52_Bumper2Ring, LM_Flashers_L52_Layer1, LM_Flashers_L52_Layer2, LM_Flashers_L52_Layer3, LM_Flashers_L52_Parts, LM_Flashers_L52_Playfield, LM_Flashers_L52_dancer, LM_Flashers_L52_sw28, LM_Flashers_L52_sw30, LM_Flashers_L53_Bumper2Ring, LM_Flashers_L53_GateL, LM_Flashers_L53_Layer1, _
  LM_Flashers_L53_Layer2, LM_Flashers_L53_Layer3, LM_Flashers_L53_Parts, LM_Flashers_L53_Playfield, LM_Flashers_L53_dancer, LM_Flashers_L53_sw28, LM_Flashers_L53_sw30, LM_Flashers_L54_Bumper2Ring, LM_Flashers_L54_Bumper3Ring, LM_Flashers_L54_GateL, LM_Flashers_L54_GateR, LM_Flashers_L54_Layer1, LM_Flashers_L54_Layer2, LM_Flashers_L54_Layer3, LM_Flashers_L54_Parts, LM_Flashers_L54_Playfield, LM_Flashers_L54_dancer, LM_Flashers_L54_sw28, LM_Flashers_L54_sw29, LM_Flashers_L54_sw30, LM_Flashers_L54_sw58p, LM_Flashers_L62_Alien1, LM_Flashers_L62_DiverterRight, LM_Flashers_L62_Layer1, LM_Flashers_L62_Layer2, LM_Flashers_L62_Layer3, LM_Flashers_L62_Parts, LM_Flashers_L62_Playfield, LM_Flashers_L62_Spinner2, LM_Flashers_L62_sw30, LM_Gi_L10_Layer1, LM_Gi_L10_Parts, LM_Gi_L10_Playfield, LM_Gi_L10_sw17p, LM_Gi_L10_sw18p, LM_Gi_L10_sw19p, LM_Gi_L10_sw20p, LM_Gi_L11_Parts, LM_Gi_L11_Playfield, LM_Gi_L11_sw17p, LM_Gi_L11_sw18p, LM_Gi_L11_sw19p, LM_Gi_L11_sw20p, LM_Gi_L12_LFlipperU, LM_Gi_L12_Lsling1, LM_Gi_L12_Lsling2, _
  LM_Gi_L12_MissionLockPin, LM_Gi_L12_Parts, LM_Gi_L12_Playfield, LM_Gi_L12_SLING2, LM_Gi_L12_sw17p, LM_Gi_L12_sw18p, LM_Gi_L12_sw19p, LM_Gi_L12_sw20p, LM_Gi_L12_sw44, LM_Gi_L12_sw45, LM_Gi_L122_Bumper1Ring, LM_Gi_L122_Bumper2Ring, LM_Gi_L122_Layer1, LM_Gi_L122_Layer3, LM_Gi_L122_Parts, LM_Gi_L122_Playfield, LM_Gi_L122_dancer, LM_Gi_L122_sw49p, LM_Gi_L122_sw50p, LM_Gi_L122_sw51p, LM_Gi_L122_sw52p, LM_Gi_L122_sw53p, LM_Gi_L123_Bumper1Ring, LM_Gi_L123_Bumper2Ring, LM_Gi_L123_Bumper3Ring, LM_Gi_L123_Layer1, LM_Gi_L123_Layer2, LM_Gi_L123_Layer3, LM_Gi_L123_Parts, LM_Gi_L123_Playfield, LM_Gi_L123_sw50p, LM_Gi_L123_sw51p, LM_Gi_L123_sw53p, LM_Gi_L124_Bumper1Ring, LM_Gi_L124_Layer1, LM_Gi_L124_Layer2, LM_Gi_L124_Layer3, LM_Gi_L124_Parts, LM_Gi_L124_Playfield, LM_Gi_L124_sw30, LM_Gi_L124_sw58p, LM_Gi_L128_Layer1, LM_Gi_L128_Layer3, LM_Gi_L128_Parts, LM_Gi_L128_Playfield, LM_Gi_L128_RFlipper1, LM_Gi_L128_RFlipper1U, LM_Gi_L14_LFlipper, LM_Gi_L14_LFlipperU, LM_Gi_L14_Lsling1, LM_Gi_L14_Lsling2, LM_Gi_L14_MissionLockPin, _
  LM_Gi_L14_Parts, LM_Gi_L14_Playfield, LM_Gi_L14_RFlipper, LM_Gi_L14_SLING2, LM_Gi_L14_sw17p, LM_Gi_L14_sw18p, LM_Gi_L14_sw19p, LM_Gi_L14_sw44, LM_Gi_L14_sw45, LM_Gi_L17_Parts, LM_Gi_L17_Playfield, LM_Gi_L17_RFlipper1, LM_Gi_L17_RFlipper1U, LM_Gi_L18_Parts, LM_Gi_L18_Playfield, LM_Gi_L18_RPost, LM_Gi_L21_Parts, LM_Gi_L21_Playfield, LM_Gi_L21_RFlipper, LM_Gi_L21_RFlipperU, LM_Gi_L21_RPost, LM_Gi_L21_Rsling1, LM_Gi_L21_Rsling2, LM_Gi_L21_SLING1, LM_Gi_L21_sw17p, LM_Gi_L21_sw65, LM_Gi_L21_sw66, LM_Gi_L22_MissionLockPin, LM_Gi_L22_Parts, LM_Gi_L22_Playfield, LM_Gi_L22_RFlipper, LM_Gi_L22_RFlipperU, LM_Gi_L22_RPost, LM_Gi_L22_Rsling1, LM_Gi_L22_Rsling2, LM_Gi_L22_SLING1, LM_Gi_L22_sw65, LM_Gi_L22_sw66, LM_Gi_L23_MissionLockPin, LM_Gi_L23_Parts, LM_Gi_L23_Playfield, LM_Gi_L23_RFlipper, LM_Gi_L23_RFlipperU, LM_Gi_L23_RPost, LM_Gi_L23_Rsling1, LM_Gi_L23_Rsling2, LM_Gi_L23_SLING1, LM_Gi_L23_sw65, LM_Gi_L23_sw66, LM_Gi_L24_MissionLockPin, LM_Gi_L24_Parts, LM_Gi_L24_Playfield, LM_Gi_L24_RFlipper, LM_Gi_L24_RFlipperU, _
  LM_Gi_L24_Rsling1, LM_Gi_L24_Rsling2, LM_Gi_L24_SLING1, LM_Gi_L24_sw44, LM_Gi_L24_sw65, LM_Gi_L24_sw66, LM_Gi_L25_Layer1, LM_Gi_L25_Layer2, LM_Gi_L25_Layer3, LM_Gi_L25_Parts, LM_Gi_L25_Playfield, LM_Gi_L25_SpinnerPlate, LM_Gi_L25_sw23p, LM_Gi_L25_sw49p, LM_Gi_L25_sw50p, LM_Gi_L26_Bumper1Ring, LM_Gi_L26_Bumper2Ring, LM_Gi_L26_Bumper3Ring, LM_Gi_L26_Layer1, LM_Gi_L26_Layer2, LM_Gi_L26_Layer3, LM_Gi_L26_Parts, LM_Gi_L26_Playfield, LM_Gi_L26_sw49p, LM_Gi_L26_sw50p, LM_Gi_L26_sw51p, LM_Gi_L26_sw52p, LM_Gi_L26_sw53p, LM_Gi_L26_sw58p, LM_Gi_L27_Bumper2Ring, LM_Gi_L27_Bumper3Ring, LM_Gi_L27_Layer1, LM_Gi_L27_Layer2, LM_Gi_L27_Layer3, LM_Gi_L27_Parts, LM_Gi_L27_Playfield, LM_Gi_L27_sw27, LM_Gi_L27_sw49p, LM_Gi_L27_sw50p, LM_Gi_L27_sw51p, LM_Gi_L27_sw52p, LM_Gi_L27_sw53p, LM_Gi_L28_Layer1, LM_Gi_L28_Layer2, LM_Gi_L28_Layer3, LM_Gi_L28_Parts, LM_Gi_L28_Playfield, LM_Gi_L28_dancer, LM_Gi_L28_sw26, LM_Gi_L29_DiverterLeft, LM_Gi_L29_DiverterRight, LM_Gi_L29_Layer1, LM_Gi_L29_Layer2, LM_Gi_L29_Layer3, LM_Gi_L29_Parts, _
  LM_Gi_L29_Playfield, LM_Gi_L29_dancer, LM_Gi_L33_Bumper1Ring, LM_Gi_L33_Layer1, LM_Gi_L33_Layer2, LM_Gi_L33_Layer3, LM_Gi_L33_Parts, LM_Gi_L33_Playfield, LM_Gi_L33_dancer, LM_Gi_L33_sw28, LM_Gi_L34_Bumper1Ring, LM_Gi_L34_Bumper2Ring, LM_Gi_L34_GateL, LM_Gi_L34_Layer1, LM_Gi_L34_Layer2, LM_Gi_L34_Layer3, LM_Gi_L34_Parts, LM_Gi_L34_Playfield, LM_Gi_L34_dancer, LM_Gi_L34_sw28, LM_Gi_L34_sw29, LM_Gi_L35_Alien1, LM_Gi_L35_Bumper1Ring, LM_Gi_L35_Bumper2Ring, LM_Gi_L35_GateL, LM_Gi_L35_GateR, LM_Gi_L35_Layer1, LM_Gi_L35_Layer2, LM_Gi_L35_Layer3, LM_Gi_L35_Parts, LM_Gi_L35_Playfield, LM_Gi_L35_sw29, LM_Gi_L35_sw30, LM_Gi_L35_sw58p, LM_Gi_L36_Alien1, LM_Gi_L36_Bumper2Ring, LM_Gi_L36_Layer1, LM_Gi_L36_Layer2, LM_Gi_L36_Layer3, LM_Gi_L36_Parts, LM_Gi_L36_Playfield, LM_Gi_L36_Spinner2, LM_Gi_L36_sw30, LM_Gi_L36_sw58p, LM_Gi_L37_Alien1, LM_Gi_L37_Alien2, LM_Gi_L37_Layer1, LM_Gi_L37_Layer2, LM_Gi_L37_Parts, LM_Gi_L37_Playfield, LM_Gi_L37_sw30, LM_Gi_L37_sw58p, LM_Gi_L39_Alien2, LM_Gi_L39_Layer1, LM_Gi_L39_Layer2, _
  LM_Gi_L39_Parts, LM_Gi_L39_Playfield, LM_Gi_L40_Bumper1Ring, LM_Gi_L40_Bumper2Ring, LM_Gi_L40_Layer1, LM_Gi_L40_Layer3, LM_Gi_L40_Parts, LM_Gi_L40_Playfield, LM_Gi_L40_sw58p, LM_Gi_L40_sw59, LM_Gi_L9_Layer1, LM_Gi_L9_MissionLockPin, LM_Gi_L9_Parts, LM_Gi_L9_Playfield, LM_Gi_L9_sw17p, LM_Gi_L9_sw18p, LM_Gi_L9_sw19p, LM_Gi_L9_sw20p, LM_Inserts_L100_Parts, LM_Inserts_L100_Playfield, LM_Inserts_L100_sw18p, LM_Inserts_L100_sw19p, LM_Inserts_L100_sw20p, LM_Inserts_L100_sw22p, LM_Inserts_L100_sw23p, LM_Inserts_L100_sw49p, LM_Inserts_L100_sw50p, LM_Inserts_L100_sw51p, LM_Inserts_L101_GateR, LM_Inserts_L101_Layer2, LM_Inserts_L101_Parts, LM_Inserts_L101_Playfield, LM_Inserts_L101_sw23p, LM_Inserts_L101_sw49p, LM_Inserts_L101_sw50p, LM_Inserts_L101_sw51p, LM_Inserts_L101_sw52p, LM_Inserts_L102_Layer1, LM_Inserts_L102_Parts, LM_Inserts_L102_Playfield, LM_Inserts_L102_sw23p, LM_Inserts_L102_sw49p, LM_Inserts_L102_sw50p, LM_Inserts_L102_sw51p, LM_Inserts_L102_sw52p, LM_Inserts_L102_sw53p, LM_Inserts_L103_Layer1, _
  LM_Inserts_L103_Parts, LM_Inserts_L103_Playfield, LM_Inserts_L103_Spinner1, LM_Inserts_L103_sw49p, LM_Inserts_L103_sw50p, LM_Inserts_L103_sw51p, LM_Inserts_L103_sw52p, LM_Inserts_L105_Layer1, LM_Inserts_L105_Parts, LM_Inserts_L105_Playfield, LM_Inserts_L105_sw49p, LM_Inserts_L105_sw50p, LM_Inserts_L105_sw51p, LM_Inserts_L105_sw53p, LM_Inserts_L106_Parts, LM_Inserts_L106_Playfield, LM_Inserts_L106_sw49p, LM_Inserts_L106_sw50p, LM_Inserts_L106_sw51p, LM_Inserts_L106_sw53p, LM_Inserts_L107_Parts, LM_Inserts_L107_Playfield, LM_Inserts_L107_sw49p, LM_Inserts_L107_sw50p, LM_Inserts_L107_sw51p, LM_Inserts_L108_Parts, LM_Inserts_L108_Playfield, LM_Inserts_L108_sw49p, LM_Inserts_L108_sw50p, LM_Inserts_L108_sw51p, LM_Inserts_L109_Parts, LM_Inserts_L109_Playfield, LM_Inserts_L109_sw49p, LM_Inserts_L109_sw50p, LM_Inserts_L109_sw51p, LM_Inserts_L110_Playfield, LM_Inserts_L110_sw49p, LM_Inserts_L110_sw50p, LM_Inserts_L110_sw51p, LM_Inserts_L111_Parts, LM_Inserts_L111_Playfield, LM_Inserts_L111_sw49p, LM_Inserts_L111_sw50p, _
  LM_Inserts_L111_sw51p, LM_Inserts_L112_Parts, LM_Inserts_L112_Playfield, LM_Inserts_L112_sw18p, LM_Inserts_L112_sw49p, LM_Inserts_L112_sw50p, LM_Inserts_L112_sw51p, LM_Inserts_L113_Parts, LM_Inserts_L113_Playfield, LM_Inserts_L113_sw49p, LM_Inserts_L113_sw50p, LM_Inserts_L113_sw51p, LM_Inserts_L113_sw52p, LM_Inserts_L113_sw53p, LM_Inserts_L114_Parts, LM_Inserts_L114_Playfield, LM_Inserts_L114_sw49p, LM_Inserts_L114_sw50p, LM_Inserts_L114_sw51p, LM_Inserts_L114_sw52p, LM_Inserts_L114_sw53p, LM_Inserts_L115_Layer1, LM_Inserts_L115_Parts, LM_Inserts_L115_Playfield, LM_Inserts_L115_sw49p, LM_Inserts_L115_sw50p, LM_Inserts_L115_sw51p, LM_Inserts_L115_sw52p, LM_Inserts_L115_sw53p, LM_Inserts_L116_Layer1, LM_Inserts_L116_Parts, LM_Inserts_L116_Playfield, LM_Inserts_L116_sw51p, LM_Inserts_L116_sw80, LM_Inserts_L117_Parts, LM_Inserts_L117_Playfield, LM_Inserts_L117_sw50p, LM_Inserts_L117_sw51p, LM_Inserts_L118_Parts, LM_Inserts_L118_Playfield, LM_Inserts_L118_sw49p, LM_Inserts_L118_sw50p, LM_Inserts_L118_sw51p, _
  LM_Inserts_L121_Bumper1Ring, LM_Inserts_L121_Bumper2Ring, LM_Inserts_L121_Layer2, LM_Inserts_L121_Layer3, LM_Inserts_L121_Parts, LM_Inserts_L121_Playfield, LM_Inserts_L121_sw58p, LM_Inserts_L126_Bumper1Ring, LM_Inserts_L126_Layer1, LM_Inserts_L126_Layer2, LM_Inserts_L126_Parts, LM_Inserts_L126_Playfield, LM_Inserts_L126_sw58p, LM_Inserts_L127_Bumper1Ring, LM_Inserts_L127_Bumper2Ring, LM_Inserts_L127_Bumper3Ring, LM_Inserts_L127_Layer1, LM_Inserts_L127_Parts, LM_Inserts_L127_Playfield, LM_Inserts_L127_sw49p, LM_Inserts_L127_sw50p, LM_Inserts_L127_sw51p, LM_Inserts_L127_sw58p, LM_Inserts_L19_Parts, LM_Inserts_L19_Playfield, LM_Inserts_L20_Parts, LM_Inserts_L20_Playfield, LM_Inserts_L30_Layer1, LM_Inserts_L30_Layer2, LM_Inserts_L30_Parts, LM_Inserts_L30_Playfield, LM_Inserts_L30_SpinnerPlate, LM_Inserts_L31_Layer1, LM_Inserts_L31_Layer2, LM_Inserts_L31_Layer3, LM_Inserts_L31_Parts, LM_Inserts_L31_Playfield, LM_Inserts_L31_Siderails, LM_Inserts_L32_Layer1, LM_Inserts_L32_Layer2, LM_Inserts_L32_Layer3, _
  LM_Inserts_L32_Parts, LM_Inserts_L32_Playfield, LM_Inserts_L32_sw26, LM_Inserts_L41_Layer1, LM_Inserts_L41_Parts, LM_Inserts_L41_Playfield, LM_Inserts_L42_Layer1, LM_Inserts_L42_Parts, LM_Inserts_L42_Playfield, LM_Inserts_L42_sw58p, LM_Inserts_L43_Layer1, LM_Inserts_L43_Layer2, LM_Inserts_L43_Parts, LM_Inserts_L43_Playfield, LM_Inserts_L43_sw58p, LM_Inserts_L49_DiverterRight, LM_Inserts_L49_GateL, LM_Inserts_L49_Layer2, LM_Inserts_L49_Layer3, LM_Inserts_L49_Parts, LM_Inserts_L49_Playfield, LM_Inserts_L49_dancer, LM_Inserts_L49_sw28, LM_Inserts_L50_Layer2, LM_Inserts_L50_Layer3, LM_Inserts_L50_Parts, LM_Inserts_L50_Playfield, LM_Inserts_L50_sw58p, LM_Inserts_L51_GateR, LM_Inserts_L51_Layer1, LM_Inserts_L51_Layer2, LM_Inserts_L51_Parts, LM_Inserts_L51_Playfield, LM_Inserts_L51_sw58p, LM_Inserts_L65_MissionLockPin, LM_Inserts_L65_Parts, LM_Inserts_L65_Playfield, LM_Inserts_L65_sw17p, LM_Inserts_L65_sw44, LM_Inserts_L65_sw45, LM_Inserts_L66_Lsling1, LM_Inserts_L66_Lsling2, LM_Inserts_L66_MissionLockPin, _
  LM_Inserts_L66_Parts, LM_Inserts_L66_Playfield, LM_Inserts_L67_LFlipper, LM_Inserts_L67_LFlipperU, LM_Inserts_L67_Parts, LM_Inserts_L67_Playfield, LM_Inserts_L68_LFlipper, LM_Inserts_L68_LFlipperU, LM_Inserts_L68_MissionLockPin, LM_Inserts_L68_Parts, LM_Inserts_L68_Playfield, LM_Inserts_L68_RFlipper, LM_Inserts_L68_RFlipperU, LM_Inserts_L68_sw44, LM_Inserts_L69_MissionLockPin, LM_Inserts_L69_Parts, LM_Inserts_L69_Playfield, LM_Inserts_L70_LFlipperU, LM_Inserts_L70_MissionLockPin, LM_Inserts_L70_Parts, LM_Inserts_L70_Playfield, LM_Inserts_L70_RFlipperU, LM_Inserts_L71_LFlipperU, LM_Inserts_L71_Parts, LM_Inserts_L71_Playfield, LM_Inserts_L71_RFlipperU, LM_Inserts_L72_LFlipper, LM_Inserts_L72_Playfield, LM_Inserts_L72_RFlipper, LM_Inserts_L73_Parts, LM_Inserts_L73_Playfield, LM_Inserts_L74_Parts, LM_Inserts_L74_Playfield, LM_Inserts_L75_Parts, LM_Inserts_L75_Playfield, LM_Inserts_L76_Parts, LM_Inserts_L76_Playfield, LM_Inserts_L76_RPost, LM_Inserts_L77_Parts, LM_Inserts_L77_Playfield, LM_Inserts_L77_Rsling1, _
  LM_Inserts_L77_Rsling2, LM_Inserts_L78_Parts, LM_Inserts_L78_Playfield, LM_Inserts_L78_SLING1, LM_Inserts_L78_sw65, LM_Inserts_L78_sw66, LM_Inserts_L79_Parts, LM_Inserts_L79_Playfield, LM_Inserts_L79_Rsling1, LM_Inserts_L79_Rsling2, LM_Inserts_L79_SLING1, LM_Inserts_L80_Parts, LM_Inserts_L80_Playfield, LM_Inserts_L80_RFlipper, LM_Inserts_L80_RFlipperU, LM_Inserts_L81_Layer1, LM_Inserts_L81_Parts, LM_Inserts_L81_Playfield, LM_Inserts_L81_sw18p, LM_Inserts_L81_sw19p, LM_Inserts_L81_sw20p, LM_Inserts_L82_Layer1, LM_Inserts_L82_Parts, LM_Inserts_L82_Playfield, LM_Inserts_L82_sw17p, LM_Inserts_L82_sw18p, LM_Inserts_L82_sw19p, LM_Inserts_L82_sw20p, LM_Inserts_L83_Parts, LM_Inserts_L83_Playfield, LM_Inserts_L83_sw17p, LM_Inserts_L83_sw18p, LM_Inserts_L83_sw19p, LM_Inserts_L83_sw20p, LM_Inserts_L84_Parts, LM_Inserts_L84_Playfield, LM_Inserts_L84_sw17p, LM_Inserts_L84_sw18p, LM_Inserts_L84_sw19p, LM_Inserts_L84_sw20p, LM_Inserts_L85_Parts, LM_Inserts_L85_Playfield, LM_Inserts_L85_sw17p, LM_Inserts_L85_sw18p, _
  LM_Inserts_L85_sw19p, LM_Inserts_L85_sw20p, LM_Inserts_L86_Parts, LM_Inserts_L86_Playfield, LM_Inserts_L86_sw17p, LM_Inserts_L86_sw18p, LM_Inserts_L86_sw19p, LM_Inserts_L86_sw20p, LM_Inserts_L87_Parts, LM_Inserts_L87_Playfield, LM_Inserts_L87_sw17p, LM_Inserts_L87_sw18p, LM_Inserts_L87_sw45, LM_Inserts_L88_Lsling1, LM_Inserts_L88_Lsling2, LM_Inserts_L88_Parts, LM_Inserts_L88_Playfield, LM_Inserts_L88_sw17p, LM_Inserts_L89_Parts, LM_Inserts_L89_Playfield, LM_Inserts_L89_sw17p, LM_Inserts_L89_sw18p, LM_Inserts_L89_sw19p, LM_Inserts_L89_sw20p, LM_Inserts_L90_Parts, LM_Inserts_L90_Playfield, LM_Inserts_L90_sw17p, LM_Inserts_L90_sw18p, LM_Inserts_L90_sw19p, LM_Inserts_L90_sw20p, LM_Inserts_L91_Parts, LM_Inserts_L91_Playfield, LM_Inserts_L91_sw17p, LM_Inserts_L91_sw18p, LM_Inserts_L91_sw19p, LM_Inserts_L92_Parts, LM_Inserts_L92_Playfield, LM_Inserts_L93_Parts, LM_Inserts_L93_Playfield, LM_Inserts_L93_sw17p, LM_Inserts_L93_sw18p, LM_Inserts_L93_sw19p, LM_Inserts_L93_sw20p, LM_Inserts_L94_Parts, _
  LM_Inserts_L94_Playfield, LM_Inserts_L94_sw17p, LM_Inserts_L94_sw18p, LM_Inserts_L94_sw19p, LM_Inserts_L95_Parts, LM_Inserts_L95_Playfield, LM_Inserts_L97_Layer1, LM_Inserts_L97_Parts, LM_Inserts_L97_Playfield, LM_Inserts_L97_sw18p, LM_Inserts_L97_sw19p, LM_Inserts_L97_sw20p, LM_Inserts_L98_Parts, LM_Inserts_L98_Playfield, LM_Inserts_L98_sw17p, LM_Inserts_L98_sw18p, LM_Inserts_L98_sw19p, LM_Inserts_L98_sw20p, LM_Inserts_L98_sw21p, LM_Inserts_L98_sw49p, LM_Inserts_L99_Parts, LM_Inserts_L99_Playfield, LM_Inserts_L99_sw17p, LM_Inserts_L99_sw18p, LM_Inserts_L99_sw19p, LM_Inserts_L99_sw20p, LM_Inserts_L99_sw22p, LM_Inserts_L99_sw49p, LM_Inserts_L99_sw50p, LM_Inserts_L99_sw51p, LM_Inserts_l57_Layer1, LM_Inserts_l57_Layer2, LM_Inserts_l57_Parts, LM_Inserts_l57_Spinner1, LM_Inserts_l58_Layer1, LM_Inserts_l58_Layer2, LM_Inserts_l58_Layer3, LM_Inserts_l58_Parts, LM_Inserts_l59_Layer1, LM_Inserts_l59_Layer2, LM_Inserts_l59_Layer3, LM_Inserts_l59_Parts, LM_LMods_Layer1, LM_LMods_Parts)
' VLM  Arrays - End



'******************************************************
'  ZTIM: Timers
'******************************************************


Sub RDampen_Timer
  Cor.Update            'update ball tracking
End Sub

' The frame timer interval is -1, so executes at the display frame rate
Sub FrameTimer_Timer()
  BSUpdate
  UpdateBallBrightness
  RollingUpdate         'update rolling sounds
  DoDTAnim
  UpdateDropTargets
  DoSTAnim
  UpdateStandupTargets
End Sub


'******************************************************
'  ZINI: Table Initialization and Exiting
'******************************************************

Sub Table_Init
  vpmInit Me
  With Controller
    .GameName = cGameName
    .SplashInfoLine = "Big Bang Bar, Capcom 1996"
    .HandleMechanics = 0
    .HandleKeyboard = 0
    .ShowDMDOnly = 1
    .ShowFrame = 0
    .ShowTitle = 0
    .Hidden = DesktopMode

'        If ColorDMD = 0 Then
'            .Games(cGameName).Settings.Value("dmd_colorize") = 0
'        ElseIf ColorDMD = 1 Then
'            .Games(cGameName).Settings.Value("dmd_colorize") = 1
'            '100%
'            .Games(cGameName).Settings.Value("dmd_red") = 0
'            .Games(cGameName).Settings.Value("dmd_green") = 255
'            .Games(cGameName).Settings.Value("dmd_blue") = 66
'            '66%
'            .Games(cGameName).Settings.Value("dmd_red66") = 0
'            .Games(cGameName).Settings.Value("dmd_green66") = 195
'            .Games(cGameName).Settings.Value("dmd_blue66") = 235
'            '33%
'            .Games(cGameName).Settings.Value("dmd_red33") = 125
'            .Games(cGameName).Settings.Value("dmd_green33") = 15
'            .Games(cGameName).Settings.Value("dmd_blue33") = 172
'            '0%
'            .Games(cGameName).Settings.Value("dmd_red0") = 0
'            .Games(cGameName).Settings.Value("dmd_green0") = 14
'            .Games(cGameName).Settings.Value("dmd_blue0") = 28
'        End If

  End With
  On Error Resume Next
  Controller.Run
  If Err Then MsgBox Err.Description
  On Error Goto 0

  vpmMapLights AllLamps

  'Nudging
  vpmNudge.TiltSwitch=10
  vpmNudge.Sensitivity=5
  vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingshot)

  'Trough
  Set BBBall1 = sw39.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set BBBall2 = sw38.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set BBBall3 = sw37.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set BBBall4 = sw36.CreateSizedballWithMass(Ballsize/2,Ballmass)

  Controller.Switch(39) = 1
  Controller.Switch(38) = 1
  Controller.Switch(37) = 1
  Controller.Switch(36) = 1

  '***Captive Ball Creation
  Set BBCapBall1 = CapCreate1.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set BBCapBall2 = CapCreate2.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set BBCapBall3 = CapCreate3.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set BBCapBall4 = CapCreate4.CreateSizedballWithMass(Ballsize/2,Ballmass)
  BBCapBall1.FrontDecal = ""
    BBCapBall2.FrontDecal = ""
    BBCapBall3.FrontDecal = ""
    BBCapBall4.FrontDecal = ""

  '***Setting up a ball array(collection), giving each ball used an individual name.
  gBOT = Array(BBCapBall1,BBCapBall2,BBCapBall3,BBCapBall4,BBBall1,BBBall2,BBBall3,BBBall4)

  vpmTimer.AddTimer 200, "CapCreate1.kick 180,1 '"    'Creates a ball from the "CapCreate1" kicker
  vpmTimer.AddTimer 210, "CapCreate1.enabled= 0 '"    'Permenantly Disables Captive Ball Kicker
  vpmTimer.AddTimer 200, "CapCreate2.kick 180,1 '"    'Creates a ball from the "CapCreate2" kicker
  vpmTimer.AddTimer 210, "CapCreate2.enabled= 0 '"    'Permenantly Disables Captive Ball Kicker
  vpmTimer.AddTimer 300, "CapCreate3.kick 180,1 '"    'Creates a ball from the "CapCreate3" kicker
  vpmTimer.AddTimer 310, "CapCreate3.enabled= 0 '"    'Permenantly Disables Captive Ball Kicker
  vpmTimer.AddTimer 300, "CapCreate4.kick 180,1 '"    'Creates a ball from the "CapCreate4" kicker
  vpmTimer.AddTimer 310, "CapCreate4.enabled= 0 '"    'Permenantly Disables Captive Ball Kicker

  '********Right Hole***********
  Set bsRHole = New cvpmBallStack
  With bsRHole
    .InitSaucer sw67, 67, 222, 25
    .KickZ = 0.4
    .InitExitSnd SoundFX("Saucer_Kick",DOFContactors), SoundFX("Solenoid",DOFContactors)
    .KickForceVar = 2
  End With

  '********KickBack Init
  kickback.PullBack

  '********Diverters Init
  DivLR.IsDropped=1
  DivTube1.isDropped=1
  DivAlienL.IsDropped=0
  DivAlienR.isDropped=0


  RStep = 0:RightSlingShot.Timerenabled=True
  LStep = 0:LeftSlingShot.Timerenabled=True

  SetRoomBrightness LightLevel

  Setbackglass

End Sub

Sub Table_Paused:Controller.Pause = 1:End Sub
Sub Table_unPaused:Controller.Pause = 0:End Sub



Dim VRLightUp: VRLightUp = False
Dim VRRGBON: VRRGBON = False

Dim RGBToggled: RGBToggled = false




'*******************************************
'  ZOPT: User Options
'*******************************************

Dim OutpostMod: OutpostMod = 1            ' 0 = Easy, 1 = Normal, 2 = Hard
Dim BlacklightOn : BlacklightOn = 1         ' 0 = Blacklight Off, 1 = Blacklight On
Dim VRRoomChoice : VRRoomChoice = 1             ' 1 = RGB Room  2 = Ultra-Minimal Room
Dim GlassScratches : GlassScratches = 0         ' 0 = Glass Scratches OFF  1 = Glass scratches ON
Dim HideMessages : HideMessages = 0         ' 0 = Show Magnasave button Instructions  1 = Never show instructions
Dim RGBSpeed : RGBSpeed = 1             ' Speed the RBG lights change colors in RGB Room. Value between 1 and 50. (1=Slow, 50=Dance Party)
Dim LightLevel : LightLevel = 0.25          ' Level of room lighting (0 to 1), where 0 is dark and 100 is brightest
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
Sub Table_OptionEvent(ByVal eventId)
    If eventId = 1 Then DisableStaticPreRendering = True
  Dim BP

    ' Outpost Difficulty
    OutpostMod = Table.Option("Outpost Difficulty", 0, 2, 1, 1, 0, Array("Easy", "Normal", "Hard"))
  If OutpostMod = 0 Then      'Easy
    zCol_RPost.y = 1390 + 13
    For Each BP in BP_RPost: BP.transy = 13: Next
  ElseIf OutpostMod = 1 Then    'Normal
    zCol_RPost.y = 1390
    For Each BP in BP_RPost: BP.transy = 0: Next
  Else              'Hard
    zCol_RPost.y = 1390 - 13
    For Each BP in BP_RPost: BP.transy = -13: Next
  End If

    ' VR Room
  If RenderingMode = 2 Then
    VRRoomChoice = Table.Option("VR Room", 1, 2, 1, 1, 0, Array("RGB Room", "Ultra Minimal"))
    VRRoom = VRRoomChoice
    ' VR Glass scratches
    GlassScratches = Table.Option("VR Glass Scratches", 0, 1, 1, 0, 0, Array("Off", "On"))

    ' VR Magnasave Instructions
    HideMessages = Table.Option("VR Magnasave Instructions", 0, 1, 1, 0, 0, Array("Show", "Hide"))

    ' VR RGB Speed
    RGBSpeed = Table.Option("VR RGB Speed (1 to 50)", 1, 50, 1, 1, 0)
  Else
    VRRoom = 0
  End If
  'VRRoom = VRRoomChoice 'uncomment to test VR on desktop

  ' Blacklight
  BlacklightOn = Table.Option("Enable Blacklight", 0, 1, 1, 1, 0, Array("Disabled", "Enabled"))

    ' Sound volumes
    VolumeDial = Table.Option("Mech Volume", 0, 1, 0.01, 0.8, 1)
    BallRollVolume = Table.Option("Ball Roll Volume", 0, 1, 0.01, 0.5, 1)
  RampRollVolume = Table.Option("Ramp Roll Volume", 0, 1, 0.01, 0.5, 1)

  ' Room brightness
  If RenderingMode = 2 Then
    LightLevel = 0.4
  Else
'   LightLevel = NightDay/100
    LightLevel = Table.Option("Table Brightness (Ambient Light Level)", 0, 1, 0.01, .5, 1)
    SetRoomBrightness LightLevel   'Uncomment this line for lightmapped tables.
  End If

  SetupRoom

    If eventId = 3 Then DisableStaticPreRendering = False
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
'   ZBBR: BALL BRIGHTNESS
'******************************************************

Const BallBrightness =  0.8       'Ball brightness - Value between 0 and 1 (0=Dark ... 1=Bright)

' Constants for plunger lane ball darkening.
' You can make a temporary wall in the plunger lane area and use the co-ordinates from the corner control points.
Const PLOffset = 0.5      'Minimum ball brightness scale in plunger lane
Const PLLeft = 860        'X position of punger lane left
Const PLRight = 930       'X position of punger lane right
Const PLTop = 1100        'Y position of punger lane top
Const PLBottom = 1940       'Y position of punger lane bottom
Dim PLGain: PLGain = (1-PLOffset)/(PLTop-PLBottom)

Sub UpdateBallBrightness
  Dim s, b_base, b_r, b_g, b_b, d_w
  b_base = 120 * BallBrightness + 135 ' orig was 120 and 70

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




'***************************************
' ZKEY: KEYS
'***************************************

Sub Table_KeyDown(ByVal keycode)

  'VR Specific
  If keycode = LeftMagnasave and VRRoom = 1 and BLON = false and VRLightUp = false then
    VRBrightnessUP.enabled = false
    VRBrightnessDOWN.enabled = true

    if HideMessages = 0 then
      VRMessageLeft.visible = true
      VRMessageRight.visible = true
      VRMessageTimer.interval = 3000
      VRMessageTimer.enabled = true
    End if
  End If

  If keycode = LeftMagnasave and VRRoom = 1 and BLON = false and VRLightUp = true then
    VRBrightnessUP.enabled = true
    VRBrightnessDOWN.enabled = false

    if HideMessages = 0 then
      VRMessageLeft.visible = true
      VRMessageRight.visible = true
      VRMessageTimer.interval = 3000
      VRMessageTimer.enabled = true
    End if
  end if

  If keycode = RightMagnasave and VRRoom = 1 and BLON = false and VRRGBON = false then
    GIRGB.enabled = True

    if HideMessages = 0 then
      VRMessageLeft.visible = true
      VRMessageRight.visible = true
      VRMessageTimer.interval = 3000
      VRMessageTimer.enabled = true
    End if
  End If

  If keycode = RightMagnasave and VRRoom = 1 and BLON = false and VRRGBON = true then
    GIRGB.enabled = False

    if HideMessages = 0 then
      VRMessageLeft.visible = true
      VRMessageRight.visible = true
      VRMessageTimer.interval = 3000
      VRMessageTimer.enabled = true
    End if
  end If

  If Keycode = StartGameKey and VRRoom = 1 then
    PinCab_Start_Button.y = PinCab_Start_Button.y -4
  End If

  If keycode = RightFlipperKey and VRRoom = 1 then VR_FlipRightLIT.x = VR_FlipRightLIT.x -4: VR_FlipRightDARK.x = VR_FlipRightDARK.x -4
  If keycode = LeftFlipperKey and VRRoom = 1 then VR_FlipperLeftLIT.x = VR_FlipperLeftLIT.x +4: VR_FlipperLeftDARK.x = VR_FlipperLeftDARK.x +4
  ' End VR Specific


  If keycode = keyInsertCoin1 or keycode = keyInsertCoin2 or keycode = keyInsertCoin3 or keycode = keyInsertCoin4 Then
    Select Case Int(rnd*3)
      Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
      Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
      Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
    End Select
  End If
  If keycode = RightFlipperKey Then FlipperActivate RightFlipper, RFPress End If
  If keycode = LeftFlipperKey Then FlipperActivate LeftFlipper, LFPress End If

  If keycode = PlungerKey Then
    Plunger.PullBack:SoundPlungerPull()
    if VRRoom = 1 then
      TimerVRPlunger.enabled = true
      TimerVRPlunger2.enabled = False
    End if
  End if


  If keycode = LeftTiltKey Then Nudge 90, 1:SoundNudgeLeft()
  If keycode = RightTiltKey Then Nudge 270, 1:SoundNudgeRight()
  If keycode = CenterTiltKey Then Nudge 0, 2:SoundNudgeCenter()
  If vpmKeyDown(keycode) Then Exit Sub
End Sub



Sub Table_KeyUp(ByVal keycode)

    'VR Specific
  If keycode = LeftMagnasave and VRRoom = 1 Then
    VRBrightnessDOWN.enabled = false
    VRBrightnessUP.enabled = false
    If VRLightUp = true then
      VRLightUp = false
      Else
      VRLightUp = true
    end If
  End If

  If keycode = RightMagnasave and VRRoom = 1 Then
      RGBToggled = true  ' The user has toggled from something other than default White. Used later when setting color back to default after BL mode
    if VRRGBON = false then
      VRRGBON = true
      Else
      VRRGBON = False
    end If
  end If
  If Keycode = StartGameKey and VRRoom = 1 then
    PinCab_Start_Button.y = PinCab_Start_Button.y +4
  End If
  If keycode = RightFlipperKey and VRRoom = 1 then VR_FlipRightLIT.x = VR_FlipRightLIT.x +4: VR_FlipRightDARK.x = VR_FlipRightDARK.x +4
  If keycode = LeftFlipperKey and VRRoom = 1 then VR_FlipperLeftLIT.x = VR_FlipperLeftLIT.x -4: VR_FlipperLeftDARK.x = VR_FlipperLeftDARK.x -4
  ' End VR Specific

  If keycode = RightFlipperKey Then FlipperDeActivate RightFlipper, RFPress End If
  If keycode = LeftFlipperKey Then FlipperDeActivate LeftFlipper, LFPress End If

  If keycode = PlungerKey Then
    Plunger.Fire : SoundPlungerReleaseBall()
    If VRRoom = 1 then
      TimerVRPlunger.enabled = false
      TimerVRPlunger2.enabled = true
    End if
  End If

  If vpmKeyUp(keycode) Then Exit Sub
End Sub



'******************************************************
' ZDRN: Drain, Trough, and Ball Release
'******************************************************

'********************* TROUGH *************************

Sub sw36_Hit   : Controller.Switch(36) = 1 : UpdateTrough : End Sub
Sub sw36_UnHit : Controller.Switch(36) = 0 : UpdateTrough : End Sub
Sub sw37_Hit   : Controller.Switch(37) = 1 : UpdateTrough : End Sub
Sub sw37_UnHit : Controller.Switch(37) = 0 : UpdateTrough : End Sub
Sub sw38_Hit   : Controller.Switch(38) = 1 : UpdateTrough : End Sub
Sub sw38_UnHit : Controller.Switch(38) = 0 : UpdateTrough : End Sub
Sub sw39_Hit   : Controller.Switch(39) = 1 : UpdateTrough : End Sub
Sub sw39_UnHit : Controller.Switch(39) = 0 : UpdateTrough : End Sub
Sub sw35_Hit   : Controller.Switch(35)  = 1 : UpdateTrough : RandomSoundDrain sw35 : End Sub
Sub sw35_UnHit : Controller.Switch(35)  = 0 : UpdateTrough : End Sub

Sub UpdateTrough
  UpdateTroughTimer.Interval = 100
  UpdateTroughTimer.Enabled = 1
End Sub

Sub UpdateTroughTimer_Timer
  If sw36.BallCntOver = 0 Then sw37.kick 57, 10
  If sw37.BallCntOver = 0 Then sw38.kick 57, 10
  If sw38.BallCntOver = 0 Then sw39.kick 57, 10
  Me.Enabled = 0
End Sub


'*****************  DRAIN & RELEASE  ******************

Sub SolTrough(enabled)
  If enabled Then
    sw35.kick 57, 20
  End If
End Sub

Sub SolRelease(enabled)
  If enabled Then
    sw36.kick 90, 10
    RandomSoundBallRelease sw36
  End If
End Sub


'******************************************
'  ZSOL: Solenoids & Flashers
'******************************************

SolCallback(1) = "SolTrough"                'OutHole
SolCallback(2) = "SolRelease"               'Ball Release
SolCallback(3) = "vpmSolSound SoundFX(""knocker"",DOFContactors),"      'Knocker
'SolCallback(4) = "vpmSolSound SoundFX(""""),"    'LeftSling
'SolCallback(5) = "vpmSolSound SoundFX(""""),"    'RightSling
SolCallback(6) = "SolKickBack"                  'KickBack
SolCallback(7) = "sol4Bank"                   '4-Bank drop target
SolCallback(8) = "SolLowerLockPin"                'Lower Lock Post
'SolCallback(9) = "SolLFlipper"                 'Left Flipper
'SolCallback(10) = "SolRFlipper"                  'Right Flipper
'SolCallback(11) = "SolURFlipper"               'Upper Flipper
SolCallback(12) = "bsRHole.SolOut"                'Eject Hole
SolCallback(13) = "SolLRDIvert"                 'Island Diverter
SolCallback(14) = "SolRDivert1"                 'Ramp Diverter 1
SolCallback(15) = "SolRDivert2"                 'Ramp Diverter 2
SolCallback(16) = "SolRDivert3"                 'Alien Lock Post
SolCallback(17) = "sol3Bank"                  '3-Bank drop target
'SolCallback(18) = "vpmSolSound SoundFX("""",DOFContactors),"     'Left Bumper
'SolCallback(19) = "vpmSolSound SoundFX("""",DOFContactors),"   'Middle Bumper
'SolCallback(20) = "vpmSolSound SoundFX("""",DOFContactors),"   'Right Bumper
SolModCallback(21)  ="Flash1"               'BackBox Left Flasher
SolModCallback(22)  ="Flash2"               'Tube Dancer Flasher
SolModCallback(23)  ="Flash3"               'Dance Floor Flasher
SolModCallback(24)  ="Flash4"               'Eject Hole Flasher
SolModCallback(25)  ="Flash5"               'Alien Lock Flasher
SolModCallback(26)  ="Flash6"               'Lower Lock Flasher
SolCallback(27) = "GateLeft"                  'Loop Left Gate
SolCallback(28) = "GateRight"                 'Loop Right Gate
SolCallback(29) = "sol1Bank"                  '1-Bank drop target
SolCallback(30) = "solDancer"                 'Tube Dancer Motor
SolCallback(31) = "SolAlienForward"               'Alien Motor Forward
SolCallback(32)  ="SolAlienReverse"               'ALien Motor Reverse

SolCallback(51)  ="SolGameOn" 'Game on/off control


' Flasher Callbacks
' Add solenoid sound effects? FIXME?

sub Flash1(level)
  F21.state = level
End Sub

sub Flash2(level)
  F22.state = level
End Sub

sub Flash3(level)
  F23.state = level
End Sub

sub Flash4(level)
  F24.state = level
End Sub

sub Flash5(level)
  F25.state = level
End Sub

sub Flash6(level)
  F26.state = level
End Sub




'******************************************
'  ZFLP: Flippers
'******************************************

'******Flipper fix stuff till emulation is fixed

dim GameStateOn
Sub SolGameOn(Enabled)
  If Enabled Then
    GameStateOn = True
  Else
    GameStateOn = False
  End If
End Sub

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"


' FLIPPERS
'******************************************

Const ReflipAngle = 20

' Flipper Solenoid Callbacks (these subs mimics how you would handle flippers in ROM based tables)
Sub SolLFlipper(Enabled)
  If Enabled and GameStateOn = True Then
    LF.Fire  'leftflipper.rotatetoend
    Controller.Switch(33) = 1
    If leftflipper.currentangle < leftflipper.endangle + ReflipAngle Then
      RandomSoundReflipUpLeft LeftFlipper
    Else
      SoundFlipperUpAttackLeft LeftFlipper
      RandomSoundFlipperUpLeft LeftFlipper
    End If
  Else
    Controller.Switch(33) = 0
    LeftFlipper.RotateToStart
    If LeftFlipper.currentangle < LeftFlipper.startAngle - 5 Then
      RandomSoundFlipperDownLeft LeftFlipper
    End If
    FlipperLeftHitParm = FlipperUpSoundLevel
  End If
End Sub

Sub SolRFlipper(Enabled)
  If Enabled and GameStateOn = True Then
    RF.Fire : RightFlipper1.RotateToEnd'rightflipper.rotatetoend
    Controller.Switch(34) = 1
    Controller.Switch(68) = 1
    If rightflipper.currentangle > rightflipper.endangle - ReflipAngle Then
      RandomSoundReflipUpRight RightFlipper
    Else
      SoundFlipperUpAttackRight RightFlipper
      RandomSoundFlipperUpRight RightFlipper
    End If
  Else
    Controller.Switch(34) = 0
    Controller.Switch(68) = 0
    RightFlipper1.RotateToStart
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
  LF.ReProcessBalls ActiveBall
  LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
  CheckLiveCatch Activeball, RightFlipper, RFCount, parm
  RF.ReProcessBalls ActiveBall
  RightFlipperCollide parm
End Sub




dim LF : Set LF = New FlipperPolarity
dim RF : Set RF = New FlipperPolarity

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




'******************************************
' ZGAT: Gates
'******************************************

dim GateStateL, GateStateR : GateStateL = 0 : GateStateR = 0
'
Sub GateLeft(enabled)  : If enabled then : GateL.Open = true : GateStateL = 1 : vpmTimer.AddTimer 1000, "GateCloseL" : End If : End Sub
Sub GateCloseL(aSw)    : GateL.Open = false  : GateStateL = 0 : End Sub

Sub GateRight(enabled) : If enabled then : GateR.Open = true : GateStateR = 1 : vpmTimer.AddTimer 1000, "GateCloseR" : End If : End Sub
Sub GateCloseR(aSw)    : GateR.Open = false : GateStateR = 0 : End Sub



'******************************************
' ZLLK: LowerLock And Post Handling
'******************************************

dim postdwn : postdwn = 0
Sub SolLowerLockPin(Enabled)
  dim BP
  if Enabled then
    If MissionLockPin.IsDropped=0 Then:Playsound SoundFX("diverter",DOFContactors):End If
    If postdwn = 0 Then vpmTimer.AddTimer 300, "PostUp'"
    MissionLockPin.IsDropped=1
    postdwn = 1
    For each BP in BP_MissionLockPin : BP.transz = -60: Next
  End If
End Sub

Sub PostUp
  postdwn = 0
  MissionLockPin.IsDropped=0
  dim BP: For each BP in BP_MissionLockPin : BP.transz = 0: Next
End Sub

Sub MissionLockPin_hit: RandomSoundMetal: End Sub


'*******Bottom Lock Switches
Sub sw46_Hit(): Controller.Switch(46) = 1:End Sub
Sub sw46_UnHit: Controller.Switch(46) = 0:End Sub

Sub sw47_Hit(): Controller.Switch(47) = 1:End Sub
Sub sw47_UnHit: Controller.Switch(47) = 0:End Sub

Sub sw48_Hit(): Controller.Switch(48) = 1:End Sub
Sub sw48_UnHit: Controller.Switch(48) = 0:End Sub


'******************************************
' ZLKB: Left Kickback
'******************************************

Sub SolKickBack(enabled)
  If enabled then
    PlaySoundAt "KickBack2", kickback
    kickback.Fire
  else
    kickback.PullBack
  End If
End Sub

'******************************************
' ZRDV: Ramp diverters
'******************************************

Sub SolLRDIvert(Enabled)
  If Enabled then
    DivLR.IsDropped=0
  Else
    DivLR.IsDropped=1
  End If
End Sub

Sub SolRDivert1(Enabled)
  If Enabled then
    DivTubef.RotateToEnd
    DivTube.isDropped=1
    DivTube1.isDropped=0
'   PlaySoundAt "diverter" , Diverter2Sound
    vpmTimer.AddTimer 3000, "UnRDivert"
  End If
End Sub
Sub UnRDivert(aSw)
  DivTubef.RotateToStart
  DivTube.isDropped=0
  DivTube1.isDropped=1
End Sub

Sub SolRDivert2(Enabled)
  If Enabled then
    DivTube2f.RotateToEnd
    DivTube2.isDropped=1
'   PlaySoundAt "diverter" , Diverter1Sound
  Else
    DivTube2f.RotateToStart
    DivTube2.isDropped=0
  End If
End Sub

Sub SolRDivert3(Enabled)
  If Enabled then
    AliensLockPin.IsDropped=1
'   PlaySoundAt "diverter" , Diverter3Sound
  Else
    AliensLockPin.IsDropped=0
  End If
End Sub

'******************************************
' ZALL: Alien Lock Mech Handler
'******************************************

'Note : The Startup timer is a workaround for motor overshoot caused by no power control option in VPM for the solenoids.
' - It sends an extra quarter cycle of switch changes with no animation before animation engages at actual overshoot position)
' - This prevents the alien mech from looping 2-3 times every time it starts due to position loss caused by overshoot in the
' - forward direction.  The alien mech has also been modified to correct a mis-position problem that happens sometimes after
' - Looped In Space mode ends.  The correction ensures the alien mech always faces home when it stops.  Since during the game
' - it only ever stops at the home position, this makes it fully functional, although the operator test mode doesn't function
' - as one would expect it to (making full loops instead of small movements), but that has no bearing on normal gameplay.
dim startstatus, checkstart : startstatus = 0 : checkstart = 0

Sub SolAlienForward(Enabled)
  If enabled then
    forward = 1 : ALockTimer.enabled = 1
  Else
    forward = 0
  End If
End Sub

Sub SolAlienReverse(Enabled)
  If enabled then
    reverse = 1 : ALockTimer.enabled = 1
  Else
    reverse = 0
  End If
End Sub

Sub ALockStartup_Timer()
  If direction = 0 then
    If checkstart = 0 and startstatus = 1 then
      vpmTimer.AddTimer 1000, "CheckStatus'"
    End If
    checkstart = 1
  Else
    checkstart = 0
  End If
End Sub

Sub CheckStatus
  if checkstart = 1 Then
    startstatus = 0
    OldPos = 31
    NewPos = 0
  End If
End Sub

dim NewPos, forward, reverse, OldPos : OldPos = 31 : NewPos = 0 : forward = 0 : reverse = 0
Const RadiusL = 40
Const RadiusR = 40

dim direction, AngAlien1, AngAlien2
Sub ALockTimer_timer()
  if reverse = 1 then : direction = -1 : end if : If forward = 1 then direction = 1 : end if
  NewPos = OldPos + direction
  If (NewPos > OldPos) and NewPos >= 32 Then NewPos = 0
  If (NewPos < OldPos) and NewPos  <  0 Then NewPos = 31
  If startstatus = 0 and forward = 1 then
    if NewPos >=  7   then startstatus = 1
    Select Case NewPos
      Case 0 : controller.switch(57) = 0 ' Home Notch 1 (Reset Trick)
      Case 1 : controller.switch(57) = 1
      Case 2 : controller.switch(57) = 0 ' Home Notch 2 (Rest Trick)
      Case 3 : controller.switch(57) = 1
      Case 7 : controller.switch(57) = 1
    End Select
  Else
    ' If OldPos <> NewPos Then PlaySound Soundfx ("Motor")
    ' Handle Optos
    If forward = 1 or reverse = 1 Then
      Select Case NewPos
        Case 0 : controller.switch(57) = 0 ' Home Notch 1
        Case 1 : controller.switch(57) = 1
        Case 2 : controller.switch(57) = 0 ' Home Notch 2
        Case 3 : controller.switch(57) = 1
        Case 7 : controller.switch(57) = 1
        Case 8 : controller.switch(57) = 0 ' Quarter Turn
        Case 9 : controller.switch(57) = 1
        Case 15: controller.switch(57) = 1
        Case 16: controller.switch(57) = 0 ' Half Turn
        Case 17: controller.switch(57) = 1
        Case 23: controller.switch(57) = 1
        Case 24: controller.switch(57) = 0 ' Three Quarters Turn
        Case 25: controller.switch(57) = 1
        Case 31: controller.switch(57) = 1
      End Select
    End If
    ' Animate Aliens

    Select Case NewPos
      case 0  : AngAlien1=90        : AngAlien2=-70
      case 1  : AngAlien1=79.0909   : AngAlien2=-59.0909
      case 2  : AngAlien1=68.1818   : AngAlien2=-48.1818
      case 3  : AngAlien1=57.2727   : AngAlien2=-37.2727
      case 4  : AngAlien1=46.3636   : AngAlien2=-26.3636
      case 5  : AngAlien1=35.4545   : AngAlien2=-15.4545
      case 6  : AngAlien1=24.5455   : AngAlien2=-4.5455
      case 7  : AngAlien1=13.6364   : AngAlien2=6.3636
      case 8  : AngAlien1=2.7273    : AngAlien2=17.2727
      case 9  : AngAlien1=-8.1818   : AngAlien2=28.1818
      case 10 : AngAlien1=-19.0909  : AngAlien2=39.0909
      case 11 : AngAlien1=-30       : AngAlien2=50
      case 12 : AngAlien1=-40.9091  : AngAlien2=60.9091
      case 13 : AngAlien1=-51.8182  : AngAlien2=71.8182
      case 14 : AngAlien1=-62.7273  : AngAlien2=82.7273
      case 15 : AngAlien1=-73.6364  : AngAlien2=93.6364
      case 16 : AngAlien1=-84.5455  : AngAlien2=104.5455
      case 17 : AngAlien1=-95.4545  : AngAlien2=115.4545
      case 18 : AngAlien1=-106.3636 : AngAlien2=126.3636
      case 19 : AngAlien1=-117.2727 : AngAlien2=137.2727
      case 20 : AngAlien1=-128.1818 : AngAlien2=148.1818
      case 21 : AngAlien1=-139.0909 : AngAlien2=159.0909
      case 22 : AngAlien1=-150      : AngAlien2=170
      case 23 : AngAlien1=-160.9091 : AngAlien2=180.9091
      case 24 : AngAlien1=-171.8182 : AngAlien2=191.8182
      case 25 : AngAlien1=-182.7273 : AngAlien2=202.7273
      case 26 : AngAlien1=-193.6364 : AngAlien2=213.6364
      case 27 : AngAlien1=-204.5455 : AngAlien2=224.5455
      case 28 : AngAlien1=-215.4545 : AngAlien2=235.4545
      case 29 : AngAlien1=-226.3636 : AngAlien2=246.3636
      case 30 : AngAlien1=-237.2727 : AngAlien2=257.2727
      case 31 : AngAlien1=-248.1818 : AngAlien2=268.1818
      case 32 : AngAlien1=-259.0909 : AngAlien2=279.0909
    End Select

    dim BP
    For each BP in BP_Alien1 : BP.RotZ = AngAlien1: Next
    For each BP in BP_Alien2 : BP.RotZ = AngAlien2: Next

    'Eject Lock
    If (NewPos >= 12 and NewPos <= 15) and (AlienLBall > 0 or AlienRBall > 0) Then
      AlienEjectR : AlienEjectL
    End If
    'Rotate the balls
    If AlienLBall>0 Then
      BallAlienL.X = BM_Alien1.X + RadiusL*sin((AngAlien1+57)*(Pi/180))
      BallAlienL.Y = BM_Alien1.Y - RadiusL*cos((AngAlien1+57)*(Pi/180))
      If NewPos > 19 Then AlienLBall=0
    End If

    If AlienRBall>0 Then
      BallAlienR.X = BM_Alien2.X + RadiusR*sin((AngAlien2+52)*(Pi/180))
      BallAlienR.Y = BM_Alien2.Y - RadiusR*cos((AngAlien2+52)*(Pi/180))
      If NewPos > 18 Then AlienRBall=0
    End If

  End If
  OldPos = NewPos
  If direction = 1 Then
    If forward = 0 and reverse = 0 and NewPos >= 6 and NewPos <= 9 Then : Direction = 0 : ALockTimer.enabled = False : StopSound "Motor": End If
  Else
    If forward = 0 and reverse = 0 Then : Direction = 0 : ALockTimer.enabled = False : End If
  End If
End Sub


dim BallAlienL, BallAlienR
dim AlienLBall, AlienRBall
AlienLBall = 0 : AlienRBall = 0

' Eject Ball Subs
Sub AlienEjectL
  DivAlienL.IsDropped = 1
  DivAlienL.Timerenabled = True
  'PlaySoundAt "KickandWire",sw61
End Sub

Sub AlienEjectR
  DivAlienR.IsDropped = 1
  DivAlienR.Timerenabled = True
  'PlaySoundAt "KickandWire",sw62
End Sub

Sub DivAlienL_Timer
  AlienLBall = 0
  DivAlienL.IsDropped = 0
  DivAlienL.Timerenabled = False
End Sub

Sub DivAlienR_Timer
  AlienRBall = 0
  DivAlienR.IsDropped = 0
  DivAlienR.Timerenabled = False
End Sub



'*************************************************
' ZSLG: SlingShots Animations
'*************************************************

Dim RStep, Lstep

Sub RightSlingShot_Slingshot
  RS.VelocityCorrect(ActiveBall)
  RandomSoundSlingshotRight sw65
' RSling.Visible = 0
' RSling1.Visible = 1
' sling1.TransZ = -20
  RStep = 0
  RightSlingShot_Timer
  RightSlingShot.TimerEnabled = 1
  vpmTimer.PulseSw 42
End Sub

Sub RightSlingShot_Timer
  Dim y1, y2, ty, BP
    Select Case RStep
    Case 0: y2=0: y1=1: ty=-20
        Case 2: y2=1: y1=0: ty=-10
        Case 3: y2=0: y1=0: ty=0   : RightSlingShot.TimerEnabled = 0
    End Select
  For each BP in BP_Rsling1 : BP.visible = y1 : Next
  For each BP in BP_Rsling2 : BP.visible = y2 : Next
  For each BP in BP_Sling1 : BP.transy = ty : Next
    RStep = RStep + 1
End Sub


Sub LeftSlingShot_Slingshot
  LS.VelocityCorrect(ActiveBall)
  RandomSoundSlingshotLeft sw45
' LSling.Visible = 0
' LSling1.Visible = 1
' sling2.TransZ = -20
  LStep = 0
  LeftSlingShot_Timer
  LeftSlingShot.TimerEnabled = 1
  vpmTimer.PulseSw 41
End Sub

Sub LeftSlingShot_Timer
  Dim y1, y2, ty, BP
    Select Case LStep
    Case 0: y2=0: y1=1: ty=-20
        Case 2: y2=1: y1=0: ty=-10
        Case 3: y2=0: y1=0: ty=0   : LeftSlingShot.TimerEnabled = 0
    End Select
  For each BP in BP_Lsling1 : BP.visible = y1 : Next
  For each BP in BP_Lsling2 : BP.visible = y2 : Next
  For each BP in BP_Sling2 : BP.transy = ty : Next
    LStep = LStep + 1
End Sub

'************************************************
' ZBMP: Bumper Animations
'************************************************

Dim dirRing1:dirRing1 = -1
Dim dirRing2:dirRing2 = -1
Dim dirRing3:dirRing3 = -1

Sub Bumper1_Hit
  vpmTimer.PulseSw 56
  Bumper1.TimerEnabled = 1
  RandomSoundBumperTop Bumper1
End Sub

Sub Bumper2_Hit
  vpmTimer.PulseSw 54
  Bumper2.TimerEnabled = 1
  RandomSoundBumperMiddle Bumper2
End Sub

Sub Bumper3_Hit
  vpmTimer.PulseSw 55
  Bumper3.TimerEnabled = 1
  RandomSoundBumperBottom Bumper3
End Sub

Sub Bumper1_timer()
  Dim z, BP
  z = BM_Bumper1Ring.transz + (5 * dirRing1)
  If z <= -40 Then dirRing1 = 1
  If z >= 0 Then
    dirRing1 = -1
    z = 0
    Bumper1.TimerEnabled = 0
  End If
  For each BP in BP_Bumper1Ring : BP.transz = z : Next
End Sub

Sub Bumper2_timer()
  Dim z, BP
  z = BM_Bumper2Ring.transz + (5 * dirRing2)
  If z <= -40 Then dirRing2 = 1
  If z >= 0 Then
    dirRing2 = -1
    z = 0
    Bumper2.TimerEnabled = 0
  End If
  For each BP in BP_Bumper2Ring : BP.transz = z : Next
End Sub

Sub Bumper3_timer()
  Dim z, BP
  z = BM_Bumper3Ring.transz + (5 * dirRing3)
  If z <= -40 Then dirRing3 = 1
  If z >= 0 Then
    dirRing3 = -1
    z = 0
    Bumper3.TimerEnabled = 0
  End If
  For each BP in BP_Bumper3Ring : BP.transz = z : Next
End Sub

'************************************************
' ZSWI: Switches
'************************************************

'*******Rollover switches
Sub sw26_Hit: Controller.Switch(26) = 1:  End Sub
Sub sw26_UnHit: Controller.Switch(26) = 0:  End Sub

Sub sw27_Hit: Controller.Switch(27) = 1:  End Sub
Sub sw27_UnHit: Controller.Switch(27) = 0:  End Sub

Sub sw28_Hit: Controller.Switch(28) = 1:  End Sub
Sub sw28_UnHit: Controller.Switch(28) = 0:  End Sub

Sub sw29_Hit: Controller.Switch(29) = 1:  End Sub
Sub sw29_UnHit: Controller.Switch(29) = 0:  End Sub

Sub sw30_Hit: Controller.Switch(30) = 1:  End Sub
Sub sw30_UnHit: Controller.Switch(30) = 0:  End Sub

Sub sw43_Hit: Controller.Switch(43) = 1:  L01.state = 2: End Sub
Sub sw43_UnHit: Controller.Switch(43) = 0:  PlaysoundAt "Launch_from_Shooter_Lane", sw43 :  End Sub

Sub sw44_Hit: Controller.Switch(44) = 1:  End Sub
Sub sw44_UnHit: Controller.Switch(44) = 0:  End Sub

Sub sw45_Hit: Controller.Switch(45) = 1:  End Sub
Sub sw45_UnHit: Controller.Switch(45) = 0:  End Sub

Sub sw59_Hit: Controller.Switch(59) = 1:  End Sub
Sub sw59_UnHit: Controller.Switch(59) = 0:  End Sub

Sub sw60_Hit: Controller.Switch(60) = 1:  End Sub
Sub sw60_UnHit: Controller.Switch(60) = 0:  End Sub

Sub sw65_Hit: Controller.Switch(65) = 1:  End Sub
Sub sw65_UnHit: Controller.Switch(65) = 0:  End Sub

Sub sw66_Hit: Controller.Switch(66) = 1:  End Sub
Sub sw66_UnHit: Controller.Switch(66) = 0:  End Sub

'*******Ramp Switches
Sub sw24_Hit(): Controller.Switch(24) = 1:  End Sub
Sub sw24_UnHit: Controller.Switch(24) = 0:  End Sub

Sub sw32_Hit(): Controller.Switch(32) = 1:  End Sub
Sub sw32_UnHit: Controller.Switch(32) = 0:  End Sub

Sub sw69_Hit(): Controller.Switch(69) = 1:  End Sub
Sub sw69_UnHit: Controller.Switch(69) = 0:  End Sub

Sub sw70_Hit(): Controller.Switch(70) = 1:  End Sub
Sub sw70_UnHit: Controller.Switch(70) = 0:  End Sub

Sub sw71_Hit(): Controller.Switch(71) = 1:  End Sub
Sub sw71_UnHit: Controller.Switch(71) = 0:  End Sub

Sub sw31_Hit(): Controller.Switch(31) = 1:  End Sub
Sub sw31_UnHit: Controller.Switch(31) = 0:  End Sub

Sub sw61_Hit():Controller.Switch(61) = 1: WireRampOff: SoundSaucerLock: AlienLBall = 1: Set BallAlienL = ActiveBall: End Sub
Sub sw61_UnHit:Controller.Switch(61) = 0: End Sub

Sub sw62_Hit():Controller.Switch(62) = 1: WireRampOff: SoundSaucerLock: AlienRBall = 1: Set BallAlienR = ActiveBall: End Sub
Sub sw62_UnHit:Controller.Switch(62) = 0: End Sub

'*******Captive Ball Switches
Sub sw77_Hit(): Controller.Switch(77) = 1: End Sub
Sub sw77_UnHit: Controller.Switch(77) = 0: End Sub

Sub sw78_Hit(): Controller.Switch(78) = 1: End Sub
Sub sw78_UnHit: Controller.Switch(78) = 0: End Sub

Sub sw79_Hit(): Controller.Switch(79) = 1: End Sub
Sub sw79_UnHit: Controller.Switch(79) = 0: End Sub

Sub sw80_Hit(): Controller.Switch(80) = 1: End Sub
Sub sw80_UnHit: Controller.Switch(80) = 0: End Sub



'*******Spinner
Sub sw25_Spin : vpmTimer.PulseSw 25: SoundSpinner sw25: End Sub
Sub Spinner2_Spin : SoundSpinner Spinner2 : End Sub
Sub Spinner1_Spin : SoundSpinner Spinner1 : End Sub

'*******StandUp Targets
Sub sw21_Hit():STHit 21:End Sub
Sub sw22_Hit():STHit 22:End Sub
Sub sw23_Hit():STHit 23:End Sub
Sub sw52_Hit():STHit 52:End Sub
Sub sw53_Hit():STHit 53:End Sub

'******************************************
' ZKIC: Kickers
'******************************************


Sub sw67_Hit:SoundSaucerLock:bsRHole.AddBall Me:End Sub


'******************************************
' ZDNC: DANCER ANIMATION
'******************************************

Const DanceAmp = 6
Dim dance

sub solDancer(enabled)
  if enabled then
    dancerT.enabled=1
  else
    dancerT.enabled=0
  end if
end sub

sub dancerT_timer()

  dance = dance + 9
  If dance >= 360 Then dance = 0 End If

  dim BP: For each BP in BP_dancer
    BP.rotx = DanceAmp*dSin(dance)
    BP.roty = DanceAmp*dCos(dance)
  Next

end sub



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
' ZRRL: Ramp Rolling Sound Effects
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
dim RampBalls(6,2)
'x,0 = ball x,1 = ID, 2 = Protection against ending early (minimum amount of updates)
'0,0 is boolean on/off, 0,1 unused for now
RampBalls(0,0) = False

' RampType
'     Setup: Set this array to the number Total number of balls that can be tracked at one time + 1.  5 ball multiball then set value to 6
'     Description: Array type indexed on BallId and a values used to deterimine what type of ramp the ball is on: False = Wire Ramp, True = Plastic Ramp
dim RampType(6)

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


'//////////////////////////////////////////////////////////////////////
'// RAMP TRIGGERS
'//////////////////////////////////////////////////////////////////////

Sub ramptrigger01_hit()
  WireRampOn True 'Play Plastic Ramp Sound
End Sub

Sub ramptrigger01_unhit()
  If activeball.vely>0 Then WireRampOff
End Sub

Sub ramptrigger001_hit()
  WireRampOn True 'Play Plastic Ramp Sound
End Sub

Sub ramptrigger02_hit()
  WireRampOff ' Turn off the Plastic Ramp Sound
End Sub

Sub ramptrigger02_unhit()
  WireRampOn False ' On Wire Ramp Pay Wire Ramp Sound
End Sub

Sub ramptrigger0002_unhit()
  WireRampOn False ' On Wire Ramp Pay Wire Ramp Sound
End Sub

Sub ramptrigger0002_hit()
  WireRampOff ' Turn off the Plastic Ramp Sound
End Sub

Sub ramptrigger00002_unhit()
  WireRampOn False ' On Wire Ramp Pay Wire Ramp Sound
End Sub

Sub ramptrigger03_hit()
  WireRampOff ' Exiting Wire Ramp Stop Playing Sound
End Sub

Sub ramptrigger03_unhit()
  PlaySoundAt "WireRamp_Stop", ramptrigger03
End Sub


Sub ramptrigger003_hit()
  WireRampOff ' Exiting Wire Ramp Stop Playing Sound
End Sub

Sub ramptrigger003_unhit()
  PlaySoundAt "WireRamp_Stop", ramptrigger003
End Sub

Sub ramptrigger002_unhit()
  PlaySoundAt "WireRamp_Stop", ramptrigger002
End Sub



'******************************************************
'   ZBRL: Ball Rolling and Drop Sounds
'******************************************************

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
  Dim b


  ' stop the sound of deleted balls
  For b = UBound(gBOT) + 1 to tnob
    rolling(b) = False
    StopSound("BallRoll_" & b)
  Next

  ' exit the sub if no balls on the table
  If UBound(gBOT) = -1 Then Exit Sub

  ' play the rolling sound for each ball

  For b = lob to UBound(gBOT)
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
RubberFlipperSoundFactor = 0.3 / 5      'volume multiplier; must not be zero
BallBouncePlayfieldSoftFactor = 0.025      'volume multiplier; must not be zero
BallBouncePlayfieldHardFactor = 0.025      'volume multiplier; must not be zero
DelayedBallDropOnPlayfieldSoundLevel = 0.8    'volume level; range [0, 1]
WallImpactSoundFactor = 0.075          'volume multiplier; must not be zero
MetalImpactSoundFactor = 0.3 / 3
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
ArchSoundFactor = 0.3 / 5      'volume multiplier; must not be zero

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

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "Table" is the name of the table
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

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "Table" is the name of the table
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
  PlaySoundAtLevelStatic SoundFX("Knocker_1",DOFKnocker), KnockerSoundLevel, l123
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



'*********************************************************************
'  Extra Sounds
'*********************************************************************

Sub ShooterEnd_Hit()
  'plungerbulb.state=0
  L01.state = 0
  PlaysoundAt "Ball_Bounce_Playfield_Soft_1", ShooterEnd
End Sub


'******************************************************
' ZSSC: SLINGSHOT CORRECTION FUNCTIONS by apophis
'******************************************************
' To add these slingshot corrections:
'   - On the table, add the endpoint primitives that define the two ends of the Slingshot
' - Initialize the SlingshotCorrection objects in InitSlingCorrection
'   - Call the .VelocityCorrect methods from the respective _Slingshot event sub


dim LS : Set LS = New SlingshotCorrection
dim RS : Set RS = New SlingshotCorrection

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
  AddSlingsPt 0, 0.00,  -4
  AddSlingsPt 1, 0.45,  -7
  AddSlingsPt 2, 0.48,  0
  AddSlingsPt 3, 0.52,  0
  AddSlingsPt 4, 0.55,  7
  AddSlingsPt 5, 1.00,  4

End Sub


Sub AddSlingsPt(idx, aX, aY)        'debugger wrapper for adjusting flipper script in-game
  dim a : a = Array(LS, RS)
  dim x : for each x in a
    x.addpoint idx, aX, aY
  Next
End Sub

Class SlingshotCorrection
  Public DebugOn, Enabled
  private Slingshot, SlingX1, SlingX2, SlingY1, SlingY2

  Public ModIn, ModOut
  Private Sub Class_Initialize : redim ModIn(0) : redim Modout(0): Enabled = True : End Sub

  Public Property let Object(aInput) : Set Slingshot = aInput : End Property
  Public Property Let EndPoint1(aInput) : SlingX1 = aInput.x: SlingY1 = aInput.y: End Property
  Public Property Let EndPoint2(aInput) : SlingX2 = aInput.x: SlingY2 = aInput.y: End Property

  Public Sub AddPoint(aIdx, aX, aY)
    ShuffleArrays ModIn, ModOut, 1 : ModIn(aIDX) = aX : ModOut(aIDX) = aY : ShuffleArrays ModIn, ModOut, 0
    If gametime > 100 then Report
  End Sub

  Public Sub Report()         'debug, reports all coords in tbPL.text
    If not debugOn then exit sub
    dim a1, a2 : a1 = ModIn : a2 = ModOut
    dim str, x : for x = 0 to uBound(a1) : str = str & x & ": " & round(a1(x),4) & ", " & round(a2(x),4) & vbnewline : next
    TBPout.text = str
  End Sub


  Public Sub VelocityCorrect(aBall)
    dim BallPos, XL, XR, YL, YR

    'Assign right and left end points
    If SlingX1 < SlingX2 Then
      XL = SlingX1 : YL = SlingY1 : XR = SlingX2 : YR = SlingY2
    Else
      XL = SlingX2 : YL = SlingY2 : XR = SlingX1 : YR = SlingY1
    End If

    'Find BallPos = % on Slingshot
    If Not IsEmpty(aBall.id) Then
      If ABS(XR-XL) > ABS(YR-YL) Then
        BallPos = PSlope(aBall.x, XL, 0, XR, 1)
      Else
        BallPos = PSlope(aBall.y, YL, 0, YR, 1)
      End If
      If BallPos < 0 Then BallPos = 0
      If BallPos > 1 Then BallPos = 1
    End If

    'Velocity angle correction
    If not IsEmpty(ModIn(0) ) then
      Dim Angle, RotVxVy
      Angle = LinearEnvelope(BallPos, ModIn, ModOut)
      'debug.print " BallPos=" & BallPos &" Angle=" & Angle
      'debug.print " BEFORE: aBall.Velx=" & aBall.Velx &" aBall.Vely" & aBall.Vely
      RotVxVy = RotPoint(aBall.Velx,aBall.Vely,Angle)
      If Enabled then aBall.Velx = RotVxVy(0)
      If Enabled then aBall.Vely = RotVxVy(1)
      'debug.print " AFTER: aBall.Velx=" & aBall.Velx &" aBall.Vely" & aBall.Vely
      'debug.print " "
    End If
  End Sub

End Class





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
Dim objBallShadow(10)

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






'******************************************
' ZRDT: Drop Targets
'******************************************


Sub sol4Bank(Enabled)
  If Enabled Then
    DTRaise 20
    DTRaise 19
    DTRaise 18
    DTRaise 17
    RandomSoundDropTargetReset BM_sw19p
  End If
End Sub

Sub sol3Bank(Enabled)
  If Enabled Then
    DTRaise 51
    DTRaise 50
    DTRaise 49
    RandomSoundDropTargetReset BM_sw50p
  End If
End Sub

Sub sol1Bank(Enabled)
  If Enabled Then
    DTRaise 58
    RandomSoundDropTargetReset BM_sw58p
  End If
End Sub

Sub sw20_hit
  DTHit 20
End Sub

Sub sw19_hit
  DTHit 19
End Sub

Sub sw18_hit
  DTHit 18
End Sub

Sub sw17_hit
  DTHit 17
End Sub

Sub sw51_hit
  DTHit 51
End Sub

Sub sw50_hit
  DTHit 50
End Sub

Sub sw49_hit
  DTHit 49
End Sub

Sub sw58_hit
  DTHit 58
End Sub



'  DROP TARGETS INITIALIZATION
'******************************


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
Dim DT58,DT51,DT50,DT49,DT20,DT19,DT18,DT17

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

Set DT58 = (new DropTarget)(sw58, sw58a, BM_sw58p, 58, 0, false)
Set DT51 = (new DropTarget)(sw51, sw51a, BM_sw51p, 51, 0, false)
Set DT50 = (new DropTarget)(sw50, sw50a, BM_sw50p, 50, 0, false)
Set DT49 = (new DropTarget)(sw49, sw49a, BM_sw49p, 49, 0, false)
Set DT20 = (new DropTarget)(sw20, sw20a, BM_sw20p, 20, 0, false)
Set DT19 = (new DropTarget)(sw19, sw19a, BM_sw19p, 19, 0, false)
Set DT18 = (new DropTarget)(sw18, sw18a, BM_sw18p, 18, 0, false)
Set DT17 = (new DropTarget)(sw17, sw17a, BM_sw17p, 17, 0, false)

Dim DTArray
DTArray = Array(DT58,DT51,DT50,DT49,DT20,DT19,DT18,DT17)


'Configure the behavior of Drop Targets.
Const DTDropSpeed = 80 'in milliseconds
Const DTDropUpSpeed = 40 'in milliseconds
Const DTDropUnits = 49 'VP units primitive drops so top of at or below the playfield
Const DTDropUpUnits = 10 'VP units primitive raises above the up position on drops up
Const DTMaxBend = 8 'max degrees primitive rotates when hit
Const DTDropDelay = 20 'time in milliseconds before target drops (due to friction/impact of the ball)
Const DTRaiseDelay = 40 'time in milliseconds before target drops back to normal up position after the solenoid fires to raise the target
Const DTBrickVel = 30 'velocity at which the target will brick, set to '0' to disable brick

Const DTEnableBrick = 0 'Set to 0 to disable bricking, 1 to enable bricking
Const DTHitSound = "" 'Drop Target Hit sound
Const DTDropSound = "DropTarget_Down" 'Drop Target Drop sound
Const DTResetSound = "DropTarget_Up" 'Drop Target reset sound

Const DTMass = 0.3 'Mass of the Drop Target (between 0 and 1), higher values provide more resistance



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


'******************************************************
'****  END DROP TARGETS
'******************************************************




'******************************************************
'  ZRST:  STAND-UP TARGET INITIALIZATION
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
Dim ST21, ST22, ST23, ST52, ST53

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
'You will also need to add a secondary hit object for each stand up (name sw11o, sw12o, and sw13o on the example Table)
'these are inclined primitives to simulate hitting a bent target and should provide so z velocity on high speed impacts


Set ST21 = (new StandupTarget)(sw21, BM_sw21p, 21, 0)
Set ST22 = (new StandupTarget)(sw22, BM_sw22p, 22, 0)
Set ST23 = (new StandupTarget)(sw23, BM_sw23p, 23, 0)
Set ST52 = (new StandupTarget)(sw52, BM_sw52p, 52, 0)
Set ST53 = (new StandupTarget)(sw53, BM_sw53p, 53, 0)

'Add all the Stand-up Target Arrays to Stand-up Target Animation Array
'   STAnimationArray = Array(ST1, ST2, ....)
Dim STArray
STArray = Array(ST21, ST22, ST23, ST52, ST53)

'Configure the behavior of Stand-up Targets
Const STAnimStep = 1.5  'vpunits per animation step (control return to Start)
Const STMaxOffset = 9   'max vp units target moves when hit

Const STMass = 0.2    'Mass of the Stand-up Target (between 0 and 1), higher values provide more resistance


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


'******************************************************
'   END STAND-UP TARGETS
'******************************************************





'******************************************************
'  ZBLL:  Blacklight Lamp
'******************************************************

Dim BLON: BLON = false

Sub L62_animate
  'Blacklight
  If BlacklightOn = 1 Then
    If L62.state < 0.4 And BLON = True Then
      BLON = False
      table.ColorGradeImage = "ColorGradeLUT256x16_1to1"
      If VRroom = 1 then ' Only do the following if we are in the VR RGB room..
        VRCabinetLIT.image = "NewBBBcabDark3"
        PinCab_BackboxNEWLIT.image ="NewBackBoxDARK1"
        DartboardLIT.image ="DartboardWhite3"
        Topper2LITON.visible = False
        FrontBulb2.visible = False

        SetGiRGBColor
        For Each Stuff in RoomLit: Stuff.Opacity = RBL:next 'sets VRroom brightness back to user setting..
        LightLevel = RBL/500 ' forces RoomBrightness along with it
        SetRoomBrightness LightLevel ' sets room brightness
        if RBL < 12 then   ' If User setting was below 12, this sets the walls to black, but the room items with a bit of opacity to be seen from the ambient light
          LeftWallLIGHT.Opacity = 0
          RightWallLIGHT.Opacity = 0
          BackWallLIT.Opacity = 0
          FrontWallLIght.Opacity = 0
          CeilingLIT.Opacity = 0
          if VRCabinetLit.Opacity < 25 then VRCabinetLIT.Opacity = 25 ' Keeps the cab lit from ambient light
          if PinCab_BackboxNEWLIT.Opacity < 25 then PinCab_BackboxNEWLIT.Opacity = 25 ' Keeps the cab lit from ambient light
        End if
       End if
    ElseIf L62.state > 0.5 And BLON = False Then
      If GameStateOn = true then
        BLON = True
        table.ColorGradeImage = "ColorGradeLUT256x16_ConSatBlueHalf"
        If VRroom = 1 then ' Only do the following if we are in the VR RGB room..
          VRCabinetLIT.image = "NewBBBcabNeon"
          PinCab_BackboxNEWLIT.image ="NewBackBoxNEON1"
          DartboardLIT.image ="DartboardNEON"
          Topper2LITON.visible = True
          FrontBulb2.visible = True

          SetGiRGBColor
          For Each Stuff in RoomLit: Stuff.Opacity = 200:next  ' Forcing Room lighting to 200 for Blacklight Mode
          LightLevel = 200/500 ' forces RoomBrightness along with it
          SetRoomBrightness LightLevel ' sets room brightness
        End If
      End If
    End If
  End If
End Sub




'******************************************************
'   ZANI: Misc Animations
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

Sub RightFlipper1_Animate
  Dim a : a = RightFlipper1.CurrentAngle
  FlipperRSh1.RotZ = a

  Dim v, BP
  v = 255.0 * (-122.0 -  RightFlipper1.CurrentAngle) / (-122.0 +  70.0)

  For each BP in BP_RFlipper1
    BP.Rotz = a
    BP.visible = v < 128.0
  Next
  For each BP in BP_RFlipper1U
    BP.Rotz = a
    BP.visible = v >= 128.0
  Next
End Sub



' Switch Animations

Sub sw26_Animate
  Dim z : z = sw26.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw26 : BP.transz = z: Next
End Sub

Sub sw27_Animate
  Dim z : z = sw27.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw27 : BP.transz = z: Next
End Sub

Sub sw28_Animate
  Dim z : z = sw28.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw28 : BP.transz = z: Next
End Sub

Sub sw29_Animate
  Dim z : z = sw29.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw29 : BP.transz = z: Next
End Sub

Sub sw30_Animate
  Dim z : z = sw30.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw30: BP.transz = z: Next
End Sub

Sub sw43_Animate
  Dim z : z = sw43.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw43: BP.transz = z: Next
End Sub

Sub sw44_Animate
  Dim z : z = sw44.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw44: BP.transz = z: Next
End Sub

Sub sw45_Animate
  Dim z : z = sw45.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw45: BP.transz = z: Next
End Sub

Sub sw59_Animate
  Dim z : z = sw59.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw59: BP.transz = z: Next
End Sub

Sub sw60_Animate
  Dim z : z = sw60.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw60: BP.transz = z: Next
End Sub

Sub sw65_Animate
  Dim z : z = sw65.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw65: BP.transz = z: Next
End Sub

Sub sw66_Animate
  Dim z : z = sw66.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw66: BP.transz = z: Next
End Sub

Sub sw77_Animate
  Dim z : z = sw77.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw77: BP.transz = z: Next
End Sub

Sub sw78_Animate
  Dim z : z = sw78.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw78: BP.transz = z: Next
End Sub

Sub sw79_Animate
  Dim z : z = sw79.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw79: BP.transz = z: Next
End Sub

Sub sw80_Animate
  Dim z : z = sw80.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw80: BP.transz = z: Next
End Sub


' Gate Animations

Sub GateL_Animate
    Dim a : a = GateL.CurrentAngle
    Dim BP : For Each BP in BP_GateL : BP.rotx = a: Next
End Sub

Sub GateR_Animate
    Dim a : a = GateR.CurrentAngle
    Dim BP : For Each BP in BP_GateR : BP.rotx = a: Next
End Sub

Sub Spinner1_Animate
    Dim a : a = Spinner1.CurrentAngle
    Dim BP : For Each BP in BP_Spinner1 : BP.rotx = a: Next
End Sub

Sub Spinner2_Animate
    Dim a : a = Spinner2.CurrentAngle
    Dim BP : For Each BP in BP_Spinner2 : BP.rotx = a: Next
End Sub



' Spinner Animations

Sub sw25_Animate
  Dim spinangle:spinangle = sw25.currentangle
  Dim BP : For Each BP in BP_SpinnerPlate : BP.RotX = spinangle: Next
End Sub



' Diverters Animations

Sub DivTubef_Animate
  Dim a:a = DivTubef.currentangle
  Dim BP : For Each BP in BP_DiverterLeft : BP.ObjRotZ = a: Next
End Sub

Sub DivTube2f_Animate
  Dim a:a = DivTube2f.currentangle
  Dim BP : For Each BP in BP_DiverterRight : BP.ObjRotZ = a: Next
End Sub


Sub UpdateDropTargets
  dim BP, tz, rx, ry

    tz = BM_sw17p.transz
  rx = BM_sw17p.rotx
  ry = BM_sw17p.roty
  For each BP in BP_sw17p : BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

    tz = BM_sw18p.transz
  rx = BM_sw18p.rotx
  ry = BM_sw18p.roty
  For each BP in BP_sw18p : BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

    tz = BM_sw19p.transz
  rx = BM_sw19p.rotx
  ry = BM_sw19p.roty
  For each BP in BP_sw19p : BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

    tz = BM_sw20p.transz
  rx = BM_sw20p.rotx
  ry = BM_sw20p.roty
  For each BP in BP_sw20p : BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

    tz = BM_sw49p.transz
  rx = BM_sw49p.rotx
  ry = BM_sw49p.roty
  For each BP in BP_sw49p : BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

    tz = BM_sw50p.transz
  rx = BM_sw50p.rotx
  ry = BM_sw50p.roty
  For each BP in BP_sw50p : BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

    tz = BM_sw51p.transz
  rx = BM_sw51p.rotx
  ry = BM_sw51p.roty
  For each BP in BP_sw51p : BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

    tz = BM_sw58p.transz
  rx = BM_sw58p.rotx
  ry = BM_sw58p.roty
  For each BP in BP_sw58p : BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

End Sub


Sub UpdateStandupTargets
  dim BP, ty

    ty = BM_sw52p.transy
  For each BP in BP_sw52p : BP.transy = ty: Next

    ty = BM_sw53p.transy
  For each BP in BP_sw53p : BP.transy = ty: Next

    ty = BM_sw21p.transy
  For each BP in BP_sw21p : BP.transy = ty: Next

    ty = BM_sw22p.transy
  For each BP in BP_sw22p : BP.transy = ty: Next

    ty = BM_sw23p.transy
  For each BP in BP_sw23p : BP.transy = ty: Next

End Sub





'******************************************************
' ZVRR:  VR Room
'******************************************************


'VRRoom Initialize *******************

Sub SetupRoom
  Dim BP

  If VRRoom = 0 then
    for each Stuff in EverythingRoom: Stuff.visible = false: next
    for each Stuff in FullVRCab: Stuff.visible = false: next
    for each Stuff in Clock: Stuff.visible = false: next

    LavaTimer.enabled = false
    ClockTimer.enabled = false
    TimerVRPlunger2.enabled = false
  End If

  If VRRoom = 1 then
    for each Stuff in EverythingRoom: Stuff.visible = true: next
    for each Stuff in FullVRCab: Stuff.visible = true: next
    for each Stuff in Clock: Stuff.visible = true: next

    'StartVR Timers..
    LavaTimer.enabled = true
    ClockTimer.enabled = true
    TimerVRPlunger2.enabled = true
    BM_Siderails.visible = false ' desktop rails and lockbar
    if GlassScratches = 1 then
      VRGlassImpurities.visible = true
    Else
      VRGlassImpurities.visible = false
    End If
    ' set lighting.. DONT TOUCH *******
    For Each Stuff in RoomLit: Stuff.Blenddisablelighting = 0.1:Next
    For Each Stuff in RoomLit: Stuff.Opacity = 250:next
    NEONWorkshop.blenddisablelighting = 1
    NewNeonVISUAL.blenddisablelighting = 1
    VR_CoinInsert2.blenddisablelighting = 5
    VR_CoinInsert1.blenddisablelighting = 5
    Topper2LITON.blenddisablelighting = 190
    LightLevel = 0.4 'forced in VR RGB room as it's tied to the room VRroom lighting.
    '**********************************
  End If

  'set backglass, backbox and rails on for 'Ultra-minimal room'
  If VRRoom = 2 then
    for each Stuff in EverythingRoom: Stuff.visible = false: next
    for each Stuff in FullVRCab: Stuff.visible = false: next
    for each Stuff in Clock: Stuff.visible = false: next

    LavaTimer.enabled = false
    ClockTimer.enabled = false
    TimerVRPlunger2.enabled = false

    PinCab_BackboxNEWLIT.visible = true
    PinCab_BackboxNEWDARK.visible = true
    LockDownBarLIT.visible = true
    LockDownBarDARK.visible = true
    SideRailMapRightLIT.visible = true
    SideRailMapRightDARK.visible = true
    SideRailMapLeftLIT.visible = true
    SideRailMapLeftDARK.visible = true
    BM_Siderails.visible = false ' desktop rails and lockbar
    VRdisplay.visible = true 'DMD
    VRGlassImpurities.visible = false
  End if

  If VRRoom > 0 then
    if HideMessages = 0 then
      VRMessageLeft.visible = true
      VRMessageRight.visible = true
      VRMessageTimer.enabled = true
    Else
      VRMessageLeft.visible = false
      VRMessageRight.visible = false
      VRMessageTimer.enabled = false
    End If
  Else
    VRMessageLeft.visible = false
    VRMessageRight.visible = false
    VRMessageTimer.enabled = false
  End If


  'Desktop Side rails
  If VRRoom > 0 OR Not DesktopMode Then
    For each BP in BP_Siderails : BP.visible = False: Next
  Else
    For each BP in BP_Siderails : BP.visible = True: Next
  End If

End Sub


' VR Start button
Sub L03_animate
  if VRRoom = 1 then PinCab_Start_Button.BlendDisableLighting = 100*(l03.GetInPlayIntensity / l03.Intensity)
End Sub



Dim RBL '(RoomBrightnessLevel)
RBL = 250 ' Default room brightness

Sub VRBrightnessDOWN_timer

  For Each Stuff in RoomLit: Stuff.Opacity = Stuff.Opacity - 2:Next
  RBL = Plant2LIT.opacity  'Using this to find the current room brightness level and set new RBL.

  LightLevel = RBL/500
  SetRoomBrightness LightLevel

  If  Plant2LIT.opacity < 12 Then
    For Each Stuff in RoomLit: Stuff.Opacity = 12:next  ' Keep room items somewhat visible. not compleltey black, while still turning lights to off.
    LeftWallLIGHT.Opacity = 0
    RightWallLIGHT.Opacity = 0
    BackWallLIT.Opacity = 0
    FrontWallLIght.Opacity = 0
    CeilingLIT.Opacity = 0
    if VRCabinetLit.Opacity < 25 then VRCabinetLIT.Opacity = 25
    if PinCab_BackboxNEWLIT.Opacity < 25 then PinCab_BackboxNEWLIT.Opacity = 25
    end if
end Sub

Sub VRBrightnessUP_timer

    if LeftWallLIGHT.Opacity < 800 then  ' limit room brightness
    For Each Stuff in RoomLit: Stuff.Opacity = Stuff.Opacity + 2:Next
    RBL = Plant2LIT.opacity  'Using this to find the current room brightness level and set new RBL.

    LightLevel = RBL/500
    SetRoomBrightness LightLevel

    if LeftWallLIGHT.Opacity = 0 Then  ' get the walls back in line with the other lighting opacity..
      LeftWallLIGHT.Opacity = 12
      RightWallLIGHT.Opacity = 12
      BackWallLIT.Opacity = 12
      FrontWallLIght.Opacity = 12
      CeilingLIT.Opacity = 12
    End if

    if VRCabinetLit.Opacity < 25 then VRCabinetLIT.Opacity = 25
    if PinCab_BackboxNEWLIT.Opacity < 25 then PinCab_BackboxNEWLIT.Opacity = 25
    end if
end Sub

'----------GI rainbow---------
dim giRed, giGreen, GIBlue
dim giRedDir, giGreenDir, GIBlueDir
dim RGBCase, RGBInc, RGBMin, RGBMax

RGBCase = 0
RGBInc = 0
RGBMin = 1
RGBMax = 254

giRed = RGBMax
giGreen = RGBMax
giBlue = RGBMax

giRedDir = 0
giGreenDir = 0
giBlueDir = 0


sub GIRGB_timer
    Select Case RGBCase
        Case 0: giRedDir=0:  giGreenDir=-1: giBlueDir=-1
        Case 1: giRedDir=0:  giGreenDir=1:  giBlueDir=0
        Case 2: giRedDir=-1: giGreenDir=0:  giBlueDir=0
        Case 3: giRedDir=0:  giGreenDir=0:  giBlueDir=1
        Case 4: giRedDir=0:  giGreenDir=-1: giBlueDir=0
        Case 5: giRedDir=1:  giGreenDir=0:  giBlueDir=0
        Case 6: giRedDir=0:  giGreenDir=0:  giBlueDir=-1
        Case 7: giRedDir=0:  giGreenDir=1:  giBlueDir=1
    End Select

    giRed = giRed + giRedDir * RGBSpeed
    giGreen = giGreen + giGreenDir * RGBSpeed
    giBlue = giBlue + giBlueDir * RGBSpeed
    RGBInc = 0

    If giRed < RGBMin then:   giRed = RGBMin:   RGBInc=1: end If
    If giRed > RGBMax then:   giRed = RGBMax:   RGBInc=1: end If
    If giGreen < RGBMin then: giGreen = RGBMin: RGBInc=1: end If
    If giGreen > RGBMax then: giGreen = RGBMax: RGBInc=1: end If
    If giBlue < RGBMin then:  giBlue = RGBMin:  RGBInc=1: end If
    If giBlue > RGBMax then:  giBlue = RGBMax:  RGBInc=1: end If

    RGBCase = RGBCase + RGBInc
    If RGBCase > 7 then RGBCase=0

    SetGiRGBColor
end sub

sub SetGiRGBColor
    dim c, x

    if BLON = true then
    c = RGB(0,32,255)
    Else
        if RGBToggled = true then
            c = RGB(giRed,giGreen,GIBlue) ' The user has toggled the VR room lights, go back to last user setting.
        else
            c = RGB(255,255,255) ' The user has never toggled the VR room lights, so we will default back to white.
        end if
    end If

    For Each x in RoomLit:x.color = c:Next
end Sub
'----------GI rainbow End ----------


' Lavalamp code below.  Thank you STEELY!  Steely had some fun here... Thanks Tom! ***************
'********************************************************************************************************************************************************************************************************************
'********************************************************************************************************************************************************************************************************************

Dim Lspeed(7), Lbob(7), blobAng(7), blobRad(7), blobSiz(7), Blob, Bcnt   ' for VR Lavalamp

For Each Blob in Lspeed
  Lbob(Bcnt) = .2
  Lspeed(Bcnt) = Int((5 * Rnd) + 2) * .05
  Bcnt = Bcnt + 1
Next

Sub LavaTimer_Timer()
  Bcnt = 0
  For Each Blob in Lava
    If Blob.TransZ <= LavaBase.Size_Z * 1.5 Then  'Change blob direction to up
      Lspeed(Bcnt) = Int((5 * Rnd) + 2) * .05   'travel speed
      blobAng(Bcnt) = Int((359 * Rnd) + 1)      'blob location/angle from center
      blobRad(Bcnt) = Int((40 * Rnd) + 10)      'blob distance from center, radius
      blobSiz(Bcnt) = Int((150 * Rnd) + 100)    'blob size
      Blob.Size_x = blobSiz(Bcnt):Blob.Size_y = blobSiz(Bcnt):Blob.Size_z = blobSiz(Bcnt)
      Blob.X = Round(Cos(blobAng(Bcnt)*0.0174533), 3) * blobRad(Bcnt) + LavaBase.X  'place blob
      Blob.Y = Round(Sin(blobAng(Bcnt)*0.0174533), 3) * blobRad(Bcnt) + LavaBase.Y
    End If

    If Blob.TransZ => LavaBase.Size_Z*5 Then    'Change blob direction to down
      blobAng(Bcnt) = Int((359 * Rnd) + 1)      'blob location/angle from center
      blobRad(Bcnt) = Int((40 * Rnd) + 10)      'blob distance from center,radius
      Blob.X = Round(Cos(blobAng(Bcnt)*0.0174533), 3) * blobRad(Bcnt) + LavaBase.X  'place blob
      Blob.Y = Round(Sin(blobAng(Bcnt)*0.0174533), 3) * blobRad(Bcnt) + LavaBase.Y
      Lspeed(Bcnt) = Int((8 * Rnd) + 5) * .05:Lspeed(Bcnt) = Lspeed(Bcnt) * -1        'travel speed
    End If

    'Make blob wobble
    If Blob.Size_x > blobSiz(Bcnt) + blobSiz(Bcnt)*.15 or Blob.Size_x < blobSiz(Bcnt) - blobSiz(Bcnt)*.15  Then Lbob(Bcnt) = Lbob(Bcnt) * -1
    Blob.Size_x = Blob.Size_x + Lbob(Bcnt)
    Blob.Size_y = Blob.Size_y + Lbob(Bcnt)
    Blob.Size_z = Blob.Size_Z - Lbob(Bcnt) * .66
    Blob.TransZ = Blob.TransZ + Lspeed(Bcnt)    'Move blob
    Bcnt = Bcnt + 1
  Next

'spins ceiling fan
FanDARK.objrotz = FanDARK.objrotz + 3
FanLit.objrotz = FanLit.objrotz + 3

End Sub

' ***************** VR Clock code below ******************
Dim CurrentMinute ' for VR clock
Sub ClockTimer_Timer()

  dim n
    n=Hour(now) MOD 12
  hour1.imagea="digit" & CStr(n \ 10)
    hour2.imagea="digit" & CStr(n mod 10)
  if n = 0 then hour1.visible = true:hour1.imagea="digit1": hour2.imagea="digit2"
  if n = 10 then hour1.visible = true:hour1.imagea="digit1": hour2.imagea="digit0"
  if n = 11 then hour1.visible = true:hour1.imagea="digit1": hour2.imagea="digit1"
  if n = 12 then hour1.visible = true:hour1.imagea="digit1": hour2.imagea="digit2"
  if n = 22 then hour1.visible = true:hour1.imagea="digit1": hour2.imagea="digit0"
  if n = 23 then hour1.visible = true:hour1.imagea="digit1": hour2.imagea="digit1"
  if n = 1 or n =2 or n =3 or n =4 or n =5 or n =6 or n =7 or n =8 or n =9 or n =13 or n =14 or n =15 or n =16 or n =17 or n =18 or n =19 or n =20 or n =21 then hour1.visible = false
  n=Minute(now)
  minute1.imagea="digit" & CStr(n \ 10)
    minute2.imagea="digit" & CStr(n mod 10)
End Sub
 ' ********************** END CLOCK CODE   *********************************

Sub VRMessageTimer_timer
  VRMessageLeft.visible = false
  VRMessageRight.visible = false
end Sub

Sub TimerVRPlunger_Timer
  if VRPlungerLIT.Y < 20 then
    VRPlungerLIT.Y = VRPlungerLIT.y +3.5
    VRPlungerDARK.Y = VRPlungerDARK.Y +3.5
  end if
End Sub

Sub TimerVRPlunger2_Timer
  VRPlungerLIT.Y = -121 + (5* Plunger.Position) -20  ' This follows our dummy plunger position for analog plunger hardware users.
  VRPlungerDARK.Y = -121 + (5* Plunger.Position) -20
end sub


'******************************************************
'*******  Set VR Backglass Flashers *******
'******************************************************

Sub SetBackglass()
  Dim obj

  For Each obj In VRBackglass
    obj.x = obj.x
    obj.Rotx = -85.8
    obj.height = - obj.y + 255
    obj.y = -75 'adjusts the distance from the backglass towards the user
  Next

End Sub

' Thalamus : Exit in a clean and proper way
Sub Table1_exit
  Controller.Pause = False
  Controller.Stop
End Sub

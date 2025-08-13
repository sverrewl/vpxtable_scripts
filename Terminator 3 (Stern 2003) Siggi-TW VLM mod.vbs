' Terminator 3: Rise of the Machines / IPD No. 4787 / June, 2003 / 4 Players
' VPX by jpsalas 2022, version 4.0.0
' Artistic Re-Skin by Siggi 2022
' Physics/Sound/VR upgrade by TastyWasps 2024 / VR Assets by Solters, DaRdog81 & DarthVito

' apophis updates:
' - UseVPMModSol=2
' - removed aGiFlashers and associated objects
' -



'001 - iaakki - fixed top lane switch perimeters, various animations, RPG animation, swapped some sounds,
'002 - bord - separated meshes for animation. New bakes on brackets and wires.
'003 - iaakki - added animations for gates, posts, gun and slings. Fixed minor flaw in right sling physics
'004 - bord - fix RPG reflections
'004a - apophis physics fixes
'005 - bord - render fixes
'006 - DGrimmReaper - VR fixes
'007 - last minute fixes (RC?)

Option Explicit
Randomize

Const BallSize = 50
Const BallMass = 1
Const RPGBallSize = 40
Dim TableWidth: TableWidth = Table1.Width
Dim TableHeight: TableHeight = Table1.Height
Const tnob = 6

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

'********************
' Table Options
'********************

'----- Target Bouncer Levels -----
Const TargetBouncerEnabled = 1    ' 0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.7   ' Level of bounces. Recommmended value of 0.7

'----- General Sound Options -----
Const VolumeDial = 0.80       ' Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Const BallRollVolume = 0.5      ' Level of ball rolling volume. Value between 0 and 1
Const RampRollVolume = 0.5      ' Level of ramp rolling volume. Value between 0 and 1

'----- Ball Shadow Options -----
Const DynamicBallShadowsOn = 0    ' 0 = no dynamic ball shadow ("triangles" near slings and such), 1 = enable dynamic ball shadow
Const AmbientBallShadowOn = 1   ' 0 = Static shadow under ball ("flasher" image, like JP's)
'                 ' 1 = Moving ball shadow ("primitive" object, like ninuzzu's) - This is the only one that shows up on the pf when in ramps and fades when close to lights!
'                 ' 2 = flasher image shadow, but it moves like ninuzzu's
'----- VR Room Options -----
Const VRArnold = 1          ' 1 = Show Terminator 3D Model, 0 = Do Not Show Terminator 3D Model
Const VRCyborg = 1          ' 1 = Show Cyborg 3D Model, 0 = Do Not Show Cyborg 3D Model
Const VRTest = 0          ' 1 = Test VR in Live Editor, 0 = Off
'Const VRSideBlades = 1       ' 1 = Art Side Blades, 0 = Black side blades


'********************
' Automate VR Mode
'********************

Dim VRMode, VR_Obj

If RenderingMode = 2 or VRTest = 1 Then
  VRMode = True
  Pincab_Backglass.BlendDisableLighting = 1.75
  ApronBlocker.Visible = 1
  ApronBlocker.SideVisible = 1
  BM_lrail.Visible = 0
' lrail001.Visible = 0
  BM_rrail.Visible = 0
' rrail001.Visible = 0
  wall34.Visible = 0
  wall34.SideVisible = 0
  wall35.Visible = 0
  wall35.SideVisible = 0
  Flasher7.Visible = 0
  Flasher8.Visible = 0
  Table1.PlayfieldReflectionStrength = 5
  For Each VR_Obj in VRCabinet : VR_Obj.Visible = 1 : Next
  For Each VR_Obj in VRRoom : VR_Obj.Visible = 1 : Next

  If VRArnold = 0 Then
    Terminator001.Visible = 0
    Terminator002.Visible = 0
    Terminator003.Visible = 0
    Terminator004.Visible = 0
    Terminator005.Visible = 0
  End If

  If VRCyborg = 0 Then
    Terminator006.Visible = 0
    Terminator007.Visible = 0
    Terminator008.Visible = 0
    RedEyeLeft.Visible = 0
    RedEyeRight.Visible = 0
  End If

' If VRSideBlades = 1 Then
'   Pincab_Blades.Image = "Pincab_Side_Blades"
' Else
'   Pincab_Blades.Image = "Pincab_Blades_New"
' End If

Else
  For Each VR_Obj in VRCabinet : VR_Obj.Visible = 0 : Next
  For Each VR_Obj in VRRoom : VR_Obj.Visible = 0 : Next
  VRMode = False
End If




'********************
'Standard definitions
'********************
Const UseVPMModSol = 2
Const UseSolenoids = 2
Const UseLamps = 1
Const UseGI = 0
Const UseSync = 0
Const HandleMech = 0

' Standard Sounds
Const SSolenoidOn = "fx_SolenoidOn"
Const SSolenoidOff = "fx_SolenoidOff"
Const SCoin = ""



' ROM language
Const cGameName = "term3" ' English
'Const cGameName = "term3f" ' French
'Const cGameName = "term3g" ' German
'Const cGameName = "term3i" ' Italian
'Const cGameName = "term3l" ' Spanish


Dim VarHidden, UseVPMColoredDMD, x
If Table1.ShowDT = true then
    UseVPMColoredDMD = true
    VarHidden = 1
    For each x in aReels:x.Visible = 1:Next
Else
    UseVPMColoredDMD = False
    VarHidden = 0
    For each x in aReels:x.Visible = 0:Next
End If

if B2SOn = true then VarHidden = 1


LoadVPM "03060000", "SEGA.VBS", 3.26





' VLM  Arrays - Start
' Arrays per baked part
Dim BP_Bumper001_Ring: BP_Bumper001_Ring=Array(BM_Bumper001_Ring, LM_gi_gi1_Bumper001_Ring, LM_ins_l71_Bumper001_Ring)
Dim BP_Bumper002_Ring: BP_Bumper002_Ring=Array(BM_Bumper002_Ring, LM_flash_F26_Bumper002_Ring, LM_flash_f30_Bumper002_Ring, LM_gi_gi1_Bumper002_Ring, LM_gi_gi027_Bumper002_Ring, LM_gi_gi36_Bumper002_Ring, LM_ins_l71_Bumper002_Ring)
Dim BP_Bumper003_Ring: BP_Bumper003_Ring=Array(BM_Bumper003_Ring, LM_flash_f30_Bumper003_Ring, LM_gi_gi1_Bumper003_Ring, LM_gi_gi027_Bumper003_Ring, LM_gi_gi029_Bumper003_Ring, LM_ins_l71_Bumper003_Ring)
Dim BP_Gate001_Wire: BP_Gate001_Wire=Array(BM_Gate001_Wire, LM_gi_gi1_Gate001_Wire, LM_gi_gi027_Gate001_Wire, LM_gi_gi36_Gate001_Wire, LM_ins_l38_Gate001_Wire)
Dim BP_Gate002_Wire: BP_Gate002_Wire=Array(BM_Gate002_Wire, LM_gi_gi029_Gate002_Wire, LM_gi_gi031_Gate002_Wire)
Dim BP_Gate003_Wire: BP_Gate003_Wire=Array(BM_Gate003_Wire, LM_flash_f31_Gate003_Wire)
Dim BP_Gate004_Wire: BP_Gate004_Wire=Array(BM_Gate004_Wire, LM_flash_f30_Gate004_Wire, LM_gi_gi1_Gate004_Wire, LM_ins_l71_Gate004_Wire)
Dim BP_Gate005_Wire: BP_Gate005_Wire=Array(BM_Gate005_Wire)
Dim BP_LFlip: BP_LFlip=Array(BM_LFlip, LM_flash_f27_LFlip, LM_flash_f32_LFlip, LM_gi_gi007_LFlip, LM_gi_gi008_LFlip, LM_ins_l17_LFlip, LM_ins_l37a_LFlip, LM_ins_l74_LFlip, LM_ins_l76_LFlip, LM_ins_l80_LFlip)
Dim BP_Lemk: BP_Lemk=Array(BM_Lemk, LM_gi_gi005_Lemk, LM_gi_gi006_Lemk, LM_gi_gi007_Lemk, LM_gi_gi008_Lemk, LM_ins_l37_Lemk)
Dim BP_Overlay: BP_Overlay=Array(BM_Overlay, LM_flash_F26_Overlay, LM_flash_f27_Overlay, LM_flash_f30_Overlay, LM_flash_f31_Overlay, LM_flash_f32_Overlay, LM_gi_gi003_Overlay, LM_gi_gi005_Overlay, LM_gi_gi006_Overlay, LM_gi_gi007_Overlay, LM_gi_gi004_Overlay, LM_gi_gi001_Overlay, LM_gi_gi1_Overlay, LM_gi_gi002_Overlay, LM_gi_gi008_Overlay, LM_ins_l18_Overlay, LM_ins_l19_Overlay, LM_ins_l23_Overlay, LM_ins_l26_Overlay, LM_ins_l27_Overlay, LM_ins_l33_Overlay, LM_ins_l34_Overlay, LM_ins_l35_Overlay, LM_ins_l37_Overlay, LM_ins_l41_Overlay, LM_ins_l42_Overlay, LM_ins_l43_Overlay, LM_ins_l44_Overlay, LM_ins_l45_Overlay, LM_ins_l46_Overlay, LM_ins_l47_Overlay, LM_ins_l48_Overlay, LM_ins_l49_Overlay, LM_ins_l50_Overlay, LM_ins_l51_Overlay, LM_ins_l57_Overlay, LM_ins_l58_Overlay, LM_ins_l59_Overlay, LM_ins_l64_Overlay, LM_ins_l71_Overlay, LM_ins_l72_Overlay, LM_ins_l8_Overlay, LM_ins_l9_Overlay)
Dim BP_Parts: BP_Parts=Array(BM_Parts, LM_flash_F26_Parts, LM_flash_f27_Parts, LM_flash_f28_Parts, LM_flash_f29_Parts, LM_flash_f30_Parts, LM_flash_f31_Parts, LM_flash_f32_Parts, LM_gi_gi003_Parts, LM_gi_gi005_Parts, LM_gi_gi006_Parts, LM_gi_gi007_Parts, LM_gi_gi004_Parts, LM_gi_gi001_Parts, LM_gi_gi1_Parts, LM_gi_gi002_Parts, LM_gi_gi008_Parts, LM_gi_gi010_Parts, LM_gi_gi011_Parts, LM_gi_gi013_Parts, LM_gi_gi014_Parts, LM_gi_gi025_Parts, LM_gi_gi027_Parts, LM_gi_gi029_Parts, LM_gi_gi031_Parts, LM_gi_gi034_Parts, LM_gi_gi36_Parts, LM_ins_l10_Parts, LM_ins_l11_Parts, LM_ins_l12_Parts, LM_ins_l13_Parts, LM_ins_l14_Parts, LM_ins_l15_Parts, LM_ins_l16_Parts, LM_ins_l17_Parts, LM_ins_l18_Parts, LM_ins_l19_Parts, LM_ins_l20_Parts, LM_ins_l21_Parts, LM_ins_l22_Parts, LM_ins_l23_Parts, LM_ins_l24_Parts, LM_ins_l25_Parts, LM_ins_l26_Parts, LM_ins_l27_Parts, LM_ins_l28_Parts, LM_ins_l29_Parts, LM_ins_l30_Parts, LM_ins_l31_Parts, LM_ins_l32_Parts, LM_ins_l33_Parts, LM_ins_l34_Parts, LM_ins_l35_Parts, LM_ins_l36_Parts, _
  LM_ins_l37_Parts, LM_ins_l38_Parts, LM_ins_l39_Parts, LM_ins_l40_Parts, LM_ins_l41_Parts, LM_ins_l42_Parts, LM_ins_l43_Parts, LM_ins_l44_Parts, LM_ins_l45_Parts, LM_ins_l46_Parts, LM_ins_l47_Parts, LM_ins_l48_Parts, LM_ins_l49_Parts, LM_ins_l50_Parts, LM_ins_l51_Parts, LM_ins_l52_Parts, LM_ins_l53_Parts, LM_ins_l54_Parts, LM_ins_l55_Parts, LM_ins_l56_Parts, LM_ins_l57_Parts, LM_ins_l58_Parts, LM_ins_l59_Parts, LM_ins_l60_Parts, LM_ins_l61_Parts, LM_ins_l62_Parts, LM_ins_l64_Parts, LM_ins_l65_Parts, LM_ins_l66_Parts, LM_ins_l67_Parts, LM_ins_l68_Parts, LM_ins_l69_Parts, LM_ins_l6_Parts, LM_ins_l71_Parts, LM_ins_l72_Parts, LM_ins_l76_Parts, LM_ins_l77_Parts, LM_ins_l7_Parts, LM_ins_l8_Parts, LM_ins_l9_Parts)
Dim BP_Playfield: BP_Playfield=Array(BM_Playfield, LM_flash_F26_Playfield, LM_flash_f27_Playfield, LM_flash_f28_Playfield, LM_flash_f30_Playfield, LM_flash_f31_Playfield, LM_flash_f32_Playfield, LM_gi_gi003_Playfield, LM_gi_gi005_Playfield, LM_gi_gi006_Playfield, LM_gi_gi007_Playfield, LM_gi_gi004_Playfield, LM_gi_gi001_Playfield, LM_gi_gi1_Playfield, LM_gi_gi002_Playfield, LM_gi_gi008_Playfield, LM_gi_gi010_Playfield, LM_gi_gi011_Playfield, LM_gi_gi013_Playfield, LM_gi_gi014_Playfield, LM_gi_gi025_Playfield, LM_gi_gi027_Playfield, LM_gi_gi029_Playfield, LM_gi_gi031_Playfield, LM_gi_gi034_Playfield, LM_gi_gi36_Playfield, LM_ins_l10_Playfield, LM_ins_l11_Playfield, LM_ins_l12_Playfield, LM_ins_l13_Playfield, LM_ins_l14_Playfield, LM_ins_l15_Playfield, LM_ins_l16_Playfield, LM_ins_l17_Playfield, LM_ins_l18_Playfield, LM_ins_l19_Playfield, LM_ins_l1_Playfield, LM_ins_l20_Playfield, LM_ins_l21_Playfield, LM_ins_l22_Playfield, LM_ins_l23_Playfield, LM_ins_l24_Playfield, LM_ins_l25_Playfield, LM_ins_l26_Playfield, _
  LM_ins_l27_Playfield, LM_ins_l28_Playfield, LM_ins_l29_Playfield, LM_ins_l2_Playfield, LM_ins_l30_Playfield, LM_ins_l31_Playfield, LM_ins_l32_Playfield, LM_ins_l33_Playfield, LM_ins_l34_Playfield, LM_ins_l35_Playfield, LM_ins_l36_Playfield, LM_ins_l37_Playfield, LM_ins_l37a_Playfield, LM_ins_l37b_Playfield, LM_ins_l38_Playfield, LM_ins_l39_Playfield, LM_ins_l3_Playfield, LM_ins_l40_Playfield, LM_ins_l41_Playfield, LM_ins_l42_Playfield, LM_ins_l43_Playfield, LM_ins_l44_Playfield, LM_ins_l45_Playfield, LM_ins_l46_Playfield, LM_ins_l47_Playfield, LM_ins_l48_Playfield, LM_ins_l49_Playfield, LM_ins_l4_Playfield, LM_ins_l50_Playfield, LM_ins_l51_Playfield, LM_ins_l52_Playfield, LM_ins_l55_Playfield, LM_ins_l56_Playfield, LM_ins_l57_Playfield, LM_ins_l58_Playfield, LM_ins_l59_Playfield, LM_ins_l5_Playfield, LM_ins_l60_Playfield, LM_ins_l61_Playfield, LM_ins_l62_Playfield, LM_ins_l65_Playfield, LM_ins_l66_Playfield, LM_ins_l67_Playfield, LM_ins_l68_Playfield, LM_ins_l69_Playfield, LM_ins_l6_Playfield, _
  LM_ins_l72_Playfield, LM_ins_l73_Playfield, LM_ins_l74_Playfield, LM_ins_l75_Playfield, LM_ins_l76_Playfield, LM_ins_l77_Playfield, LM_ins_l78_Playfield, LM_ins_l7_Playfield, LM_ins_l80_Playfield, LM_ins_l8_Playfield, LM_ins_l9_Playfield)
Dim BP_Post1: BP_Post1=Array(BM_Post1, LM_flash_F26_Post1, LM_flash_f30_Post1, LM_gi_gi1_Post1, LM_ins_l71_Post1)
Dim BP_Post2: BP_Post2=Array(BM_Post2, LM_flash_F26_Post2, LM_flash_f30_Post2, LM_flash_f31_Post2, LM_gi_gi1_Post2, LM_gi_gi027_Post2, LM_ins_l38_Post2, LM_ins_l71_Post2)
Dim BP_RFlip: BP_RFlip=Array(BM_RFlip, LM_flash_f27_RFlip, LM_flash_f32_RFlip, LM_gi_gi001_RFlip, LM_gi_gi002_RFlip, LM_ins_l25_RFlip, LM_ins_l37a_RFlip, LM_ins_l37b_RFlip, LM_ins_l75_RFlip, LM_ins_l77_RFlip, LM_ins_l80_RFlip)
Dim BP_RPGun: BP_RPGun=Array(BM_RPGun, LM_flash_f27_RPGun, LM_flash_f32_RPGun)
Dim BP_Remk: BP_Remk=Array(BM_Remk, LM_gi_gi001_Remk, LM_gi_gi002_Remk)
Dim BP_Wall049: BP_Wall049=Array(BM_Wall049, LM_flash_f27_Wall049, LM_gi_gi003_Wall049, LM_gi_gi004_Wall049, LM_gi_gi001_Wall049, LM_ins_l35_Wall049, LM_ins_l36_Wall049, LM_ins_l77_Wall049)
Dim BP_Wall91: BP_Wall91=Array(BM_Wall91, LM_gi_gi005_Wall91, LM_gi_gi006_Wall91, LM_ins_l33_Wall91, LM_ins_l34_Wall91, LM_ins_l37_Wall91, LM_ins_l76_Wall91)
Dim BP_lrail: BP_lrail=Array(BM_lrail)
Dim BP_lsling: BP_lsling=Array(BM_lsling, LM_gi_gi007_lsling, LM_gi_gi008_lsling, LM_ins_l34_lsling)
Dim BP_lsling_001: BP_lsling_001=Array(BM_lsling_001, LM_flash_f31_lsling_001, LM_flash_f32_lsling_001, LM_gi_gi007_lsling_001, LM_gi_gi008_lsling_001, LM_ins_l18_lsling_001, LM_ins_l34_lsling_001)
Dim BP_lsling_002: BP_lsling_002=Array(BM_lsling_002, LM_gi_gi007_lsling_002, LM_gi_gi008_lsling_002, LM_ins_l34_lsling_002)
Dim BP_rrail: BP_rrail=Array(BM_rrail, LM_gi_gi011_rrail)
Dim BP_rsling: BP_rsling=Array(BM_rsling, LM_gi_gi001_rsling, LM_gi_gi002_rsling, LM_ins_l27_rsling, LM_ins_l35_rsling)
Dim BP_rsling_001: BP_rsling_001=Array(BM_rsling_001, LM_flash_f32_rsling_001, LM_gi_gi001_rsling_001, LM_gi_gi002_rsling_001, LM_ins_l25_rsling_001, LM_ins_l26_rsling_001, LM_ins_l27_rsling_001, LM_ins_l35_rsling_001)
Dim BP_rsling_002: BP_rsling_002=Array(BM_rsling_002, LM_flash_f27_rsling_002, LM_gi_gi005_rsling_002, LM_gi_gi006_rsling_002, LM_gi_gi007_rsling_002, LM_gi_gi001_rsling_002, LM_gi_gi002_rsling_002, LM_ins_l26_rsling_002, LM_ins_l35_rsling_002)
Dim BP_sw24: BP_sw24=Array(BM_sw24, LM_gi_gi001_sw24)
Dim BP_sw25: BP_sw25=Array(BM_sw25, LM_flash_F26_sw25, LM_gi_gi1_sw25, LM_ins_l47_sw25, LM_ins_l57_sw25, LM_ins_l59_sw25, LM_ins_l71_sw25)
Dim BP_sw28: BP_sw28=Array(BM_sw28, LM_flash_F26_sw28, LM_flash_f30_sw28, LM_gi_gi1_sw28)
Dim BP_sw29: BP_sw29=Array(BM_sw29, LM_flash_F26_sw29, LM_flash_f30_sw29, LM_gi_gi1_sw29, LM_ins_l71_sw29)
Dim BP_sw37: BP_sw37=Array(BM_sw37, LM_flash_f30_sw37, LM_gi_gi1_sw37, LM_gi_gi027_sw37, LM_gi_gi36_sw37, LM_ins_l38_sw37, LM_ins_l71_sw37)
Dim BP_sw38: BP_sw38=Array(BM_sw38, LM_flash_f30_sw38, LM_flash_f31_sw38, LM_flash_f32_sw38, LM_gi_gi001_sw38, LM_gi_gi1_sw38, LM_gi_gi027_sw38, LM_gi_gi029_sw38, LM_ins_l39_sw38, LM_ins_l71_sw38)
Dim BP_sw39: BP_sw39=Array(BM_sw39, LM_flash_f30_sw39, LM_gi_gi1_sw39, LM_gi_gi027_sw39, LM_gi_gi029_sw39, LM_gi_gi031_sw39, LM_ins_l40_sw39, LM_ins_l71_sw39)
Dim BP_sw46: BP_sw46=Array(BM_sw46, LM_gi_gi001_sw46, LM_gi_gi1_sw46, LM_ins_l71_sw46)
Dim BP_sw57: BP_sw57=Array(BM_sw57, LM_gi_gi006_sw57, LM_gi_gi007_sw57, LM_ins_l37_sw57)
Dim BP_sw58: BP_sw58=Array(BM_sw58, LM_flash_f32_sw58, LM_gi_gi005_sw58, LM_gi_gi006_sw58, LM_gi_gi007_sw58, LM_gi_gi008_sw58, LM_ins_l19_sw58, LM_ins_l37_sw58)
Dim BP_sw60: BP_sw60=Array(BM_sw60, LM_flash_f27_sw60, LM_gi_gi004_sw60, LM_gi_gi001_sw60)
Dim BP_sw61: BP_sw61=Array(BM_sw61, LM_flash_f27_sw61, LM_flash_f32_sw61, LM_gi_gi003_sw61, LM_gi_gi004_sw61, LM_gi_gi001_sw61, LM_gi_gi002_sw61, LM_ins_l26_sw61, LM_ins_l27_sw61)
' Arrays per lighting scenario
Dim BL_World: BL_World=Array(BM_Bumper001_Ring, BM_Bumper002_Ring, BM_Bumper003_Ring, BM_Gate001_Wire, BM_Gate002_Wire, BM_Gate003_Wire, BM_Gate004_Wire, BM_Gate005_Wire, BM_LFlip, BM_Lemk, BM_Overlay, BM_Parts, BM_Playfield, BM_Post1, BM_Post2, BM_RFlip, BM_RPGun, BM_Remk, BM_Wall049, BM_Wall91, BM_lrail, BM_lsling, BM_lsling_001, BM_lsling_002, BM_rrail, BM_rsling, BM_rsling_001, BM_rsling_002, BM_sw24, BM_sw25, BM_sw28, BM_sw29, BM_sw37, BM_sw38, BM_sw39, BM_sw46, BM_sw57, BM_sw58, BM_sw60, BM_sw61)
Dim BL_flash_F26: BL_flash_F26=Array(LM_flash_F26_Bumper002_Ring, LM_flash_F26_Overlay, LM_flash_F26_Parts, LM_flash_F26_Playfield, LM_flash_F26_Post1, LM_flash_F26_Post2, LM_flash_F26_sw25, LM_flash_F26_sw28, LM_flash_F26_sw29)
Dim BL_flash_f27: BL_flash_f27=Array(LM_flash_f27_LFlip, LM_flash_f27_Overlay, LM_flash_f27_Parts, LM_flash_f27_Playfield, LM_flash_f27_RFlip, LM_flash_f27_RPGun, LM_flash_f27_Wall049, LM_flash_f27_rsling_002, LM_flash_f27_sw60, LM_flash_f27_sw61)
Dim BL_flash_f28: BL_flash_f28=Array(LM_flash_f28_Parts, LM_flash_f28_Playfield)
Dim BL_flash_f29: BL_flash_f29=Array(LM_flash_f29_Parts)
Dim BL_flash_f30: BL_flash_f30=Array(LM_flash_f30_Bumper002_Ring, LM_flash_f30_Bumper003_Ring, LM_flash_f30_Gate004_Wire, LM_flash_f30_Overlay, LM_flash_f30_Parts, LM_flash_f30_Playfield, LM_flash_f30_Post1, LM_flash_f30_Post2, LM_flash_f30_sw28, LM_flash_f30_sw29, LM_flash_f30_sw37, LM_flash_f30_sw38, LM_flash_f30_sw39)
Dim BL_flash_f31: BL_flash_f31=Array(LM_flash_f31_Gate003_Wire, LM_flash_f31_Overlay, LM_flash_f31_Parts, LM_flash_f31_Playfield, LM_flash_f31_Post2, LM_flash_f31_lsling_001, LM_flash_f31_sw38)
Dim BL_flash_f32: BL_flash_f32=Array(LM_flash_f32_LFlip, LM_flash_f32_Overlay, LM_flash_f32_Parts, LM_flash_f32_Playfield, LM_flash_f32_RFlip, LM_flash_f32_RPGun, LM_flash_f32_lsling_001, LM_flash_f32_rsling_001, LM_flash_f32_sw38, LM_flash_f32_sw58, LM_flash_f32_sw61)
Dim BL_gi_gi001: BL_gi_gi001=Array(LM_gi_gi001_Overlay, LM_gi_gi001_Parts, LM_gi_gi001_Playfield, LM_gi_gi001_RFlip, LM_gi_gi001_Remk, LM_gi_gi001_Wall049, LM_gi_gi001_rsling, LM_gi_gi001_rsling_001, LM_gi_gi001_rsling_002, LM_gi_gi001_sw24, LM_gi_gi001_sw38, LM_gi_gi001_sw46, LM_gi_gi001_sw60, LM_gi_gi001_sw61)
Dim BL_gi_gi002: BL_gi_gi002=Array(LM_gi_gi002_Overlay, LM_gi_gi002_Parts, LM_gi_gi002_Playfield, LM_gi_gi002_RFlip, LM_gi_gi002_Remk, LM_gi_gi002_rsling, LM_gi_gi002_rsling_001, LM_gi_gi002_rsling_002, LM_gi_gi002_sw61)
Dim BL_gi_gi003: BL_gi_gi003=Array(LM_gi_gi003_Overlay, LM_gi_gi003_Parts, LM_gi_gi003_Playfield, LM_gi_gi003_Wall049, LM_gi_gi003_sw61)
Dim BL_gi_gi004: BL_gi_gi004=Array(LM_gi_gi004_Overlay, LM_gi_gi004_Parts, LM_gi_gi004_Playfield, LM_gi_gi004_Wall049, LM_gi_gi004_sw60, LM_gi_gi004_sw61)
Dim BL_gi_gi005: BL_gi_gi005=Array(LM_gi_gi005_Lemk, LM_gi_gi005_Overlay, LM_gi_gi005_Parts, LM_gi_gi005_Playfield, LM_gi_gi005_Wall91, LM_gi_gi005_rsling_002, LM_gi_gi005_sw58)
Dim BL_gi_gi006: BL_gi_gi006=Array(LM_gi_gi006_Lemk, LM_gi_gi006_Overlay, LM_gi_gi006_Parts, LM_gi_gi006_Playfield, LM_gi_gi006_Wall91, LM_gi_gi006_rsling_002, LM_gi_gi006_sw57, LM_gi_gi006_sw58)
Dim BL_gi_gi007: BL_gi_gi007=Array(LM_gi_gi007_LFlip, LM_gi_gi007_Lemk, LM_gi_gi007_Overlay, LM_gi_gi007_Parts, LM_gi_gi007_Playfield, LM_gi_gi007_lsling, LM_gi_gi007_lsling_001, LM_gi_gi007_lsling_002, LM_gi_gi007_rsling_002, LM_gi_gi007_sw57, LM_gi_gi007_sw58)
Dim BL_gi_gi008: BL_gi_gi008=Array(LM_gi_gi008_LFlip, LM_gi_gi008_Lemk, LM_gi_gi008_Overlay, LM_gi_gi008_Parts, LM_gi_gi008_Playfield, LM_gi_gi008_lsling, LM_gi_gi008_lsling_001, LM_gi_gi008_lsling_002, LM_gi_gi008_sw58)
Dim BL_gi_gi010: BL_gi_gi010=Array(LM_gi_gi010_Parts, LM_gi_gi010_Playfield)
Dim BL_gi_gi011: BL_gi_gi011=Array(LM_gi_gi011_Parts, LM_gi_gi011_Playfield, LM_gi_gi011_rrail)
Dim BL_gi_gi013: BL_gi_gi013=Array(LM_gi_gi013_Parts, LM_gi_gi013_Playfield)
Dim BL_gi_gi014: BL_gi_gi014=Array(LM_gi_gi014_Parts, LM_gi_gi014_Playfield)
Dim BL_gi_gi025: BL_gi_gi025=Array(LM_gi_gi025_Parts, LM_gi_gi025_Playfield)
Dim BL_gi_gi027: BL_gi_gi027=Array(LM_gi_gi027_Bumper002_Ring, LM_gi_gi027_Bumper003_Ring, LM_gi_gi027_Gate001_Wire, LM_gi_gi027_Parts, LM_gi_gi027_Playfield, LM_gi_gi027_Post2, LM_gi_gi027_sw37, LM_gi_gi027_sw38, LM_gi_gi027_sw39)
Dim BL_gi_gi029: BL_gi_gi029=Array(LM_gi_gi029_Bumper003_Ring, LM_gi_gi029_Gate002_Wire, LM_gi_gi029_Parts, LM_gi_gi029_Playfield, LM_gi_gi029_sw38, LM_gi_gi029_sw39)
Dim BL_gi_gi031: BL_gi_gi031=Array(LM_gi_gi031_Gate002_Wire, LM_gi_gi031_Parts, LM_gi_gi031_Playfield, LM_gi_gi031_sw39)
Dim BL_gi_gi034: BL_gi_gi034=Array(LM_gi_gi034_Parts, LM_gi_gi034_Playfield)
Dim BL_gi_gi1: BL_gi_gi1=Array(LM_gi_gi1_Bumper001_Ring, LM_gi_gi1_Bumper002_Ring, LM_gi_gi1_Bumper003_Ring, LM_gi_gi1_Gate001_Wire, LM_gi_gi1_Gate004_Wire, LM_gi_gi1_Overlay, LM_gi_gi1_Parts, LM_gi_gi1_Playfield, LM_gi_gi1_Post1, LM_gi_gi1_Post2, LM_gi_gi1_sw25, LM_gi_gi1_sw28, LM_gi_gi1_sw29, LM_gi_gi1_sw37, LM_gi_gi1_sw38, LM_gi_gi1_sw39, LM_gi_gi1_sw46)
Dim BL_gi_gi36: BL_gi_gi36=Array(LM_gi_gi36_Bumper002_Ring, LM_gi_gi36_Gate001_Wire, LM_gi_gi36_Parts, LM_gi_gi36_Playfield, LM_gi_gi36_sw37)
Dim BL_ins_l1: BL_ins_l1=Array(LM_ins_l1_Playfield)
Dim BL_ins_l10: BL_ins_l10=Array(LM_ins_l10_Parts, LM_ins_l10_Playfield)
Dim BL_ins_l11: BL_ins_l11=Array(LM_ins_l11_Parts, LM_ins_l11_Playfield)
Dim BL_ins_l12: BL_ins_l12=Array(LM_ins_l12_Parts, LM_ins_l12_Playfield)
Dim BL_ins_l13: BL_ins_l13=Array(LM_ins_l13_Parts, LM_ins_l13_Playfield)
Dim BL_ins_l14: BL_ins_l14=Array(LM_ins_l14_Parts, LM_ins_l14_Playfield)
Dim BL_ins_l15: BL_ins_l15=Array(LM_ins_l15_Parts, LM_ins_l15_Playfield)
Dim BL_ins_l16: BL_ins_l16=Array(LM_ins_l16_Parts, LM_ins_l16_Playfield)
Dim BL_ins_l17: BL_ins_l17=Array(LM_ins_l17_LFlip, LM_ins_l17_Parts, LM_ins_l17_Playfield)
Dim BL_ins_l18: BL_ins_l18=Array(LM_ins_l18_Overlay, LM_ins_l18_Parts, LM_ins_l18_Playfield, LM_ins_l18_lsling_001)
Dim BL_ins_l19: BL_ins_l19=Array(LM_ins_l19_Overlay, LM_ins_l19_Parts, LM_ins_l19_Playfield, LM_ins_l19_sw58)
Dim BL_ins_l2: BL_ins_l2=Array(LM_ins_l2_Playfield)
Dim BL_ins_l20: BL_ins_l20=Array(LM_ins_l20_Parts, LM_ins_l20_Playfield)
Dim BL_ins_l21: BL_ins_l21=Array(LM_ins_l21_Parts, LM_ins_l21_Playfield)
Dim BL_ins_l22: BL_ins_l22=Array(LM_ins_l22_Parts, LM_ins_l22_Playfield)
Dim BL_ins_l23: BL_ins_l23=Array(LM_ins_l23_Overlay, LM_ins_l23_Parts, LM_ins_l23_Playfield)
Dim BL_ins_l24: BL_ins_l24=Array(LM_ins_l24_Parts, LM_ins_l24_Playfield)
Dim BL_ins_l25: BL_ins_l25=Array(LM_ins_l25_Parts, LM_ins_l25_Playfield, LM_ins_l25_RFlip, LM_ins_l25_rsling_001)
Dim BL_ins_l26: BL_ins_l26=Array(LM_ins_l26_Overlay, LM_ins_l26_Parts, LM_ins_l26_Playfield, LM_ins_l26_rsling_001, LM_ins_l26_rsling_002, LM_ins_l26_sw61)
Dim BL_ins_l27: BL_ins_l27=Array(LM_ins_l27_Overlay, LM_ins_l27_Parts, LM_ins_l27_Playfield, LM_ins_l27_rsling, LM_ins_l27_rsling_001, LM_ins_l27_sw61)
Dim BL_ins_l28: BL_ins_l28=Array(LM_ins_l28_Parts, LM_ins_l28_Playfield)
Dim BL_ins_l29: BL_ins_l29=Array(LM_ins_l29_Parts, LM_ins_l29_Playfield)
Dim BL_ins_l3: BL_ins_l3=Array(LM_ins_l3_Playfield)
Dim BL_ins_l30: BL_ins_l30=Array(LM_ins_l30_Parts, LM_ins_l30_Playfield)
Dim BL_ins_l31: BL_ins_l31=Array(LM_ins_l31_Parts, LM_ins_l31_Playfield)
Dim BL_ins_l32: BL_ins_l32=Array(LM_ins_l32_Parts, LM_ins_l32_Playfield)
Dim BL_ins_l33: BL_ins_l33=Array(LM_ins_l33_Overlay, LM_ins_l33_Parts, LM_ins_l33_Playfield, LM_ins_l33_Wall91)
Dim BL_ins_l34: BL_ins_l34=Array(LM_ins_l34_Overlay, LM_ins_l34_Parts, LM_ins_l34_Playfield, LM_ins_l34_Wall91, LM_ins_l34_lsling, LM_ins_l34_lsling_001, LM_ins_l34_lsling_002)
Dim BL_ins_l35: BL_ins_l35=Array(LM_ins_l35_Overlay, LM_ins_l35_Parts, LM_ins_l35_Playfield, LM_ins_l35_Wall049, LM_ins_l35_rsling, LM_ins_l35_rsling_001, LM_ins_l35_rsling_002)
Dim BL_ins_l36: BL_ins_l36=Array(LM_ins_l36_Parts, LM_ins_l36_Playfield, LM_ins_l36_Wall049)
Dim BL_ins_l37: BL_ins_l37=Array(LM_ins_l37_Lemk, LM_ins_l37_Overlay, LM_ins_l37_Parts, LM_ins_l37_Playfield, LM_ins_l37_Wall91, LM_ins_l37_sw57, LM_ins_l37_sw58)
Dim BL_ins_l37a: BL_ins_l37a=Array(LM_ins_l37a_LFlip, LM_ins_l37a_Playfield, LM_ins_l37a_RFlip)
Dim BL_ins_l37b: BL_ins_l37b=Array(LM_ins_l37b_Playfield, LM_ins_l37b_RFlip)
Dim BL_ins_l38: BL_ins_l38=Array(LM_ins_l38_Gate001_Wire, LM_ins_l38_Parts, LM_ins_l38_Playfield, LM_ins_l38_Post2, LM_ins_l38_sw37)
Dim BL_ins_l39: BL_ins_l39=Array(LM_ins_l39_Parts, LM_ins_l39_Playfield, LM_ins_l39_sw38)
Dim BL_ins_l4: BL_ins_l4=Array(LM_ins_l4_Playfield)
Dim BL_ins_l40: BL_ins_l40=Array(LM_ins_l40_Parts, LM_ins_l40_Playfield, LM_ins_l40_sw39)
Dim BL_ins_l41: BL_ins_l41=Array(LM_ins_l41_Overlay, LM_ins_l41_Parts, LM_ins_l41_Playfield)
Dim BL_ins_l42: BL_ins_l42=Array(LM_ins_l42_Overlay, LM_ins_l42_Parts, LM_ins_l42_Playfield)
Dim BL_ins_l43: BL_ins_l43=Array(LM_ins_l43_Overlay, LM_ins_l43_Parts, LM_ins_l43_Playfield)
Dim BL_ins_l44: BL_ins_l44=Array(LM_ins_l44_Overlay, LM_ins_l44_Parts, LM_ins_l44_Playfield)
Dim BL_ins_l45: BL_ins_l45=Array(LM_ins_l45_Overlay, LM_ins_l45_Parts, LM_ins_l45_Playfield)
Dim BL_ins_l46: BL_ins_l46=Array(LM_ins_l46_Overlay, LM_ins_l46_Parts, LM_ins_l46_Playfield)
Dim BL_ins_l47: BL_ins_l47=Array(LM_ins_l47_Overlay, LM_ins_l47_Parts, LM_ins_l47_Playfield, LM_ins_l47_sw25)
Dim BL_ins_l48: BL_ins_l48=Array(LM_ins_l48_Overlay, LM_ins_l48_Parts, LM_ins_l48_Playfield)
Dim BL_ins_l49: BL_ins_l49=Array(LM_ins_l49_Overlay, LM_ins_l49_Parts, LM_ins_l49_Playfield)
Dim BL_ins_l5: BL_ins_l5=Array(LM_ins_l5_Playfield)
Dim BL_ins_l50: BL_ins_l50=Array(LM_ins_l50_Overlay, LM_ins_l50_Parts, LM_ins_l50_Playfield)
Dim BL_ins_l51: BL_ins_l51=Array(LM_ins_l51_Overlay, LM_ins_l51_Parts, LM_ins_l51_Playfield)
Dim BL_ins_l52: BL_ins_l52=Array(LM_ins_l52_Parts, LM_ins_l52_Playfield)
Dim BL_ins_l53: BL_ins_l53=Array(LM_ins_l53_Parts)
Dim BL_ins_l54: BL_ins_l54=Array(LM_ins_l54_Parts)
Dim BL_ins_l55: BL_ins_l55=Array(LM_ins_l55_Parts, LM_ins_l55_Playfield)
Dim BL_ins_l56: BL_ins_l56=Array(LM_ins_l56_Parts, LM_ins_l56_Playfield)
Dim BL_ins_l57: BL_ins_l57=Array(LM_ins_l57_Overlay, LM_ins_l57_Parts, LM_ins_l57_Playfield, LM_ins_l57_sw25)
Dim BL_ins_l58: BL_ins_l58=Array(LM_ins_l58_Overlay, LM_ins_l58_Parts, LM_ins_l58_Playfield)
Dim BL_ins_l59: BL_ins_l59=Array(LM_ins_l59_Overlay, LM_ins_l59_Parts, LM_ins_l59_Playfield, LM_ins_l59_sw25)
Dim BL_ins_l6: BL_ins_l6=Array(LM_ins_l6_Parts, LM_ins_l6_Playfield)
Dim BL_ins_l60: BL_ins_l60=Array(LM_ins_l60_Parts, LM_ins_l60_Playfield)
Dim BL_ins_l61: BL_ins_l61=Array(LM_ins_l61_Parts, LM_ins_l61_Playfield)
Dim BL_ins_l62: BL_ins_l62=Array(LM_ins_l62_Parts, LM_ins_l62_Playfield)
Dim BL_ins_l64: BL_ins_l64=Array(LM_ins_l64_Overlay, LM_ins_l64_Parts)
Dim BL_ins_l65: BL_ins_l65=Array(LM_ins_l65_Parts, LM_ins_l65_Playfield)
Dim BL_ins_l66: BL_ins_l66=Array(LM_ins_l66_Parts, LM_ins_l66_Playfield)
Dim BL_ins_l67: BL_ins_l67=Array(LM_ins_l67_Parts, LM_ins_l67_Playfield)
Dim BL_ins_l68: BL_ins_l68=Array(LM_ins_l68_Parts, LM_ins_l68_Playfield)
Dim BL_ins_l69: BL_ins_l69=Array(LM_ins_l69_Parts, LM_ins_l69_Playfield)
Dim BL_ins_l7: BL_ins_l7=Array(LM_ins_l7_Parts, LM_ins_l7_Playfield)
Dim BL_ins_l71: BL_ins_l71=Array(LM_ins_l71_Bumper001_Ring, LM_ins_l71_Bumper002_Ring, LM_ins_l71_Bumper003_Ring, LM_ins_l71_Gate004_Wire, LM_ins_l71_Overlay, LM_ins_l71_Parts, LM_ins_l71_Post1, LM_ins_l71_Post2, LM_ins_l71_sw25, LM_ins_l71_sw29, LM_ins_l71_sw37, LM_ins_l71_sw38, LM_ins_l71_sw39, LM_ins_l71_sw46)
Dim BL_ins_l72: BL_ins_l72=Array(LM_ins_l72_Overlay, LM_ins_l72_Parts, LM_ins_l72_Playfield)
Dim BL_ins_l73: BL_ins_l73=Array(LM_ins_l73_Playfield)
Dim BL_ins_l74: BL_ins_l74=Array(LM_ins_l74_LFlip, LM_ins_l74_Playfield)
Dim BL_ins_l75: BL_ins_l75=Array(LM_ins_l75_Playfield, LM_ins_l75_RFlip)
Dim BL_ins_l76: BL_ins_l76=Array(LM_ins_l76_LFlip, LM_ins_l76_Parts, LM_ins_l76_Playfield, LM_ins_l76_Wall91)
Dim BL_ins_l77: BL_ins_l77=Array(LM_ins_l77_Parts, LM_ins_l77_Playfield, LM_ins_l77_RFlip, LM_ins_l77_Wall049)
Dim BL_ins_l78: BL_ins_l78=Array(LM_ins_l78_Playfield)
Dim BL_ins_l8: BL_ins_l8=Array(LM_ins_l8_Overlay, LM_ins_l8_Parts, LM_ins_l8_Playfield)
Dim BL_ins_l80: BL_ins_l80=Array(LM_ins_l80_LFlip, LM_ins_l80_Playfield, LM_ins_l80_RFlip)
Dim BL_ins_l9: BL_ins_l9=Array(LM_ins_l9_Overlay, LM_ins_l9_Parts, LM_ins_l9_Playfield)
' Global arrays
Dim BG_Bakemap: BG_Bakemap=Array(BM_Bumper001_Ring, BM_Bumper002_Ring, BM_Bumper003_Ring, BM_Gate001_Wire, BM_Gate002_Wire, BM_Gate003_Wire, BM_Gate004_Wire, BM_Gate005_Wire, BM_LFlip, BM_Lemk, BM_Overlay, BM_Parts, BM_Playfield, BM_Post1, BM_Post2, BM_RFlip, BM_RPGun, BM_Remk, BM_Wall049, BM_Wall91, BM_lrail, BM_lsling, BM_lsling_001, BM_lsling_002, BM_rrail, BM_rsling, BM_rsling_001, BM_rsling_002, BM_sw24, BM_sw25, BM_sw28, BM_sw29, BM_sw37, BM_sw38, BM_sw39, BM_sw46, BM_sw57, BM_sw58, BM_sw60, BM_sw61)
Dim BG_Lightmap: BG_Lightmap=Array(LM_flash_F26_Bumper002_Ring, LM_flash_F26_Overlay, LM_flash_F26_Parts, LM_flash_F26_Playfield, LM_flash_F26_Post1, LM_flash_F26_Post2, LM_flash_F26_sw25, LM_flash_F26_sw28, LM_flash_F26_sw29, LM_flash_f27_LFlip, LM_flash_f27_Overlay, LM_flash_f27_Parts, LM_flash_f27_Playfield, LM_flash_f27_RFlip, LM_flash_f27_RPGun, LM_flash_f27_Wall049, LM_flash_f27_rsling_002, LM_flash_f27_sw60, LM_flash_f27_sw61, LM_flash_f28_Parts, LM_flash_f28_Playfield, LM_flash_f29_Parts, LM_flash_f30_Bumper002_Ring, LM_flash_f30_Bumper003_Ring, LM_flash_f30_Gate004_Wire, LM_flash_f30_Overlay, LM_flash_f30_Parts, LM_flash_f30_Playfield, LM_flash_f30_Post1, LM_flash_f30_Post2, LM_flash_f30_sw28, LM_flash_f30_sw29, LM_flash_f30_sw37, LM_flash_f30_sw38, LM_flash_f30_sw39, LM_flash_f31_Gate003_Wire, LM_flash_f31_Overlay, LM_flash_f31_Parts, LM_flash_f31_Playfield, LM_flash_f31_Post2, LM_flash_f31_lsling_001, LM_flash_f31_sw38, LM_flash_f32_LFlip, LM_flash_f32_Overlay, LM_flash_f32_Parts, _
  LM_flash_f32_Playfield, LM_flash_f32_RFlip, LM_flash_f32_RPGun, LM_flash_f32_lsling_001, LM_flash_f32_rsling_001, LM_flash_f32_sw38, LM_flash_f32_sw58, LM_flash_f32_sw61, LM_gi_gi001_Overlay, LM_gi_gi001_Parts, LM_gi_gi001_Playfield, LM_gi_gi001_RFlip, LM_gi_gi001_Remk, LM_gi_gi001_Wall049, LM_gi_gi001_rsling, LM_gi_gi001_rsling_001, LM_gi_gi001_rsling_002, LM_gi_gi001_sw24, LM_gi_gi001_sw38, LM_gi_gi001_sw46, LM_gi_gi001_sw60, LM_gi_gi001_sw61, LM_gi_gi002_Overlay, LM_gi_gi002_Parts, LM_gi_gi002_Playfield, LM_gi_gi002_RFlip, LM_gi_gi002_Remk, LM_gi_gi002_rsling, LM_gi_gi002_rsling_001, LM_gi_gi002_rsling_002, LM_gi_gi002_sw61, LM_gi_gi003_Overlay, LM_gi_gi003_Parts, LM_gi_gi003_Playfield, LM_gi_gi003_Wall049, LM_gi_gi003_sw61, LM_gi_gi004_Overlay, LM_gi_gi004_Parts, LM_gi_gi004_Playfield, LM_gi_gi004_Wall049, LM_gi_gi004_sw60, LM_gi_gi004_sw61, LM_gi_gi005_Lemk, LM_gi_gi005_Overlay, LM_gi_gi005_Parts, LM_gi_gi005_Playfield, LM_gi_gi005_Wall91, LM_gi_gi005_rsling_002, LM_gi_gi005_sw58, LM_gi_gi006_Lemk, _
  LM_gi_gi006_Overlay, LM_gi_gi006_Parts, LM_gi_gi006_Playfield, LM_gi_gi006_Wall91, LM_gi_gi006_rsling_002, LM_gi_gi006_sw57, LM_gi_gi006_sw58, LM_gi_gi007_LFlip, LM_gi_gi007_Lemk, LM_gi_gi007_Overlay, LM_gi_gi007_Parts, LM_gi_gi007_Playfield, LM_gi_gi007_lsling, LM_gi_gi007_lsling_001, LM_gi_gi007_lsling_002, LM_gi_gi007_rsling_002, LM_gi_gi007_sw57, LM_gi_gi007_sw58, LM_gi_gi008_LFlip, LM_gi_gi008_Lemk, LM_gi_gi008_Overlay, LM_gi_gi008_Parts, LM_gi_gi008_Playfield, LM_gi_gi008_lsling, LM_gi_gi008_lsling_001, LM_gi_gi008_lsling_002, LM_gi_gi008_sw58, LM_gi_gi010_Parts, LM_gi_gi010_Playfield, LM_gi_gi011_Parts, LM_gi_gi011_Playfield, LM_gi_gi011_rrail, LM_gi_gi013_Parts, LM_gi_gi013_Playfield, LM_gi_gi014_Parts, LM_gi_gi014_Playfield, LM_gi_gi025_Parts, LM_gi_gi025_Playfield, LM_gi_gi027_Bumper002_Ring, LM_gi_gi027_Bumper003_Ring, LM_gi_gi027_Gate001_Wire, LM_gi_gi027_Parts, LM_gi_gi027_Playfield, LM_gi_gi027_Post2, LM_gi_gi027_sw37, LM_gi_gi027_sw38, LM_gi_gi027_sw39, LM_gi_gi029_Bumper003_Ring, _
  LM_gi_gi029_Gate002_Wire, LM_gi_gi029_Parts, LM_gi_gi029_Playfield, LM_gi_gi029_sw38, LM_gi_gi029_sw39, LM_gi_gi031_Gate002_Wire, LM_gi_gi031_Parts, LM_gi_gi031_Playfield, LM_gi_gi031_sw39, LM_gi_gi034_Parts, LM_gi_gi034_Playfield, LM_gi_gi1_Bumper001_Ring, LM_gi_gi1_Bumper002_Ring, LM_gi_gi1_Bumper003_Ring, LM_gi_gi1_Gate001_Wire, LM_gi_gi1_Gate004_Wire, LM_gi_gi1_Overlay, LM_gi_gi1_Parts, LM_gi_gi1_Playfield, LM_gi_gi1_Post1, LM_gi_gi1_Post2, LM_gi_gi1_sw25, LM_gi_gi1_sw28, LM_gi_gi1_sw29, LM_gi_gi1_sw37, LM_gi_gi1_sw38, LM_gi_gi1_sw39, LM_gi_gi1_sw46, LM_gi_gi36_Bumper002_Ring, LM_gi_gi36_Gate001_Wire, LM_gi_gi36_Parts, LM_gi_gi36_Playfield, LM_gi_gi36_sw37, LM_ins_l1_Playfield, LM_ins_l10_Parts, LM_ins_l10_Playfield, LM_ins_l11_Parts, LM_ins_l11_Playfield, LM_ins_l12_Parts, LM_ins_l12_Playfield, LM_ins_l13_Parts, LM_ins_l13_Playfield, LM_ins_l14_Parts, LM_ins_l14_Playfield, LM_ins_l15_Parts, LM_ins_l15_Playfield, LM_ins_l16_Parts, LM_ins_l16_Playfield, LM_ins_l17_LFlip, LM_ins_l17_Parts, _
  LM_ins_l17_Playfield, LM_ins_l18_Overlay, LM_ins_l18_Parts, LM_ins_l18_Playfield, LM_ins_l18_lsling_001, LM_ins_l19_Overlay, LM_ins_l19_Parts, LM_ins_l19_Playfield, LM_ins_l19_sw58, LM_ins_l2_Playfield, LM_ins_l20_Parts, LM_ins_l20_Playfield, LM_ins_l21_Parts, LM_ins_l21_Playfield, LM_ins_l22_Parts, LM_ins_l22_Playfield, LM_ins_l23_Overlay, LM_ins_l23_Parts, LM_ins_l23_Playfield, LM_ins_l24_Parts, LM_ins_l24_Playfield, LM_ins_l25_Parts, LM_ins_l25_Playfield, LM_ins_l25_RFlip, LM_ins_l25_rsling_001, LM_ins_l26_Overlay, LM_ins_l26_Parts, LM_ins_l26_Playfield, LM_ins_l26_rsling_001, LM_ins_l26_rsling_002, LM_ins_l26_sw61, LM_ins_l27_Overlay, LM_ins_l27_Parts, LM_ins_l27_Playfield, LM_ins_l27_rsling, LM_ins_l27_rsling_001, LM_ins_l27_sw61, LM_ins_l28_Parts, LM_ins_l28_Playfield, LM_ins_l29_Parts, LM_ins_l29_Playfield, LM_ins_l3_Playfield, LM_ins_l30_Parts, LM_ins_l30_Playfield, LM_ins_l31_Parts, LM_ins_l31_Playfield, LM_ins_l32_Parts, LM_ins_l32_Playfield, LM_ins_l33_Overlay, LM_ins_l33_Parts, _
  LM_ins_l33_Playfield, LM_ins_l33_Wall91, LM_ins_l34_Overlay, LM_ins_l34_Parts, LM_ins_l34_Playfield, LM_ins_l34_Wall91, LM_ins_l34_lsling, LM_ins_l34_lsling_001, LM_ins_l34_lsling_002, LM_ins_l35_Overlay, LM_ins_l35_Parts, LM_ins_l35_Playfield, LM_ins_l35_Wall049, LM_ins_l35_rsling, LM_ins_l35_rsling_001, LM_ins_l35_rsling_002, LM_ins_l36_Parts, LM_ins_l36_Playfield, LM_ins_l36_Wall049, LM_ins_l37_Lemk, LM_ins_l37_Overlay, LM_ins_l37_Parts, LM_ins_l37_Playfield, LM_ins_l37_Wall91, LM_ins_l37_sw57, LM_ins_l37_sw58, LM_ins_l37a_LFlip, LM_ins_l37a_Playfield, LM_ins_l37a_RFlip, LM_ins_l37b_Playfield, LM_ins_l37b_RFlip, LM_ins_l38_Gate001_Wire, LM_ins_l38_Parts, LM_ins_l38_Playfield, LM_ins_l38_Post2, LM_ins_l38_sw37, LM_ins_l39_Parts, LM_ins_l39_Playfield, LM_ins_l39_sw38, LM_ins_l4_Playfield, LM_ins_l40_Parts, LM_ins_l40_Playfield, LM_ins_l40_sw39, LM_ins_l41_Overlay, LM_ins_l41_Parts, LM_ins_l41_Playfield, LM_ins_l42_Overlay, LM_ins_l42_Parts, LM_ins_l42_Playfield, LM_ins_l43_Overlay, LM_ins_l43_Parts, _
  LM_ins_l43_Playfield, LM_ins_l44_Overlay, LM_ins_l44_Parts, LM_ins_l44_Playfield, LM_ins_l45_Overlay, LM_ins_l45_Parts, LM_ins_l45_Playfield, LM_ins_l46_Overlay, LM_ins_l46_Parts, LM_ins_l46_Playfield, LM_ins_l47_Overlay, LM_ins_l47_Parts, LM_ins_l47_Playfield, LM_ins_l47_sw25, LM_ins_l48_Overlay, LM_ins_l48_Parts, LM_ins_l48_Playfield, LM_ins_l49_Overlay, LM_ins_l49_Parts, LM_ins_l49_Playfield, LM_ins_l5_Playfield, LM_ins_l50_Overlay, LM_ins_l50_Parts, LM_ins_l50_Playfield, LM_ins_l51_Overlay, LM_ins_l51_Parts, LM_ins_l51_Playfield, LM_ins_l52_Parts, LM_ins_l52_Playfield, LM_ins_l53_Parts, LM_ins_l54_Parts, LM_ins_l55_Parts, LM_ins_l55_Playfield, LM_ins_l56_Parts, LM_ins_l56_Playfield, LM_ins_l57_Overlay, LM_ins_l57_Parts, LM_ins_l57_Playfield, LM_ins_l57_sw25, LM_ins_l58_Overlay, LM_ins_l58_Parts, LM_ins_l58_Playfield, LM_ins_l59_Overlay, LM_ins_l59_Parts, LM_ins_l59_Playfield, LM_ins_l59_sw25, LM_ins_l6_Parts, LM_ins_l6_Playfield, LM_ins_l60_Parts, LM_ins_l60_Playfield, LM_ins_l61_Parts, _
  LM_ins_l61_Playfield, LM_ins_l62_Parts, LM_ins_l62_Playfield, LM_ins_l64_Overlay, LM_ins_l64_Parts, LM_ins_l65_Parts, LM_ins_l65_Playfield, LM_ins_l66_Parts, LM_ins_l66_Playfield, LM_ins_l67_Parts, LM_ins_l67_Playfield, LM_ins_l68_Parts, LM_ins_l68_Playfield, LM_ins_l69_Parts, LM_ins_l69_Playfield, LM_ins_l7_Parts, LM_ins_l7_Playfield, LM_ins_l71_Bumper001_Ring, LM_ins_l71_Bumper002_Ring, LM_ins_l71_Bumper003_Ring, LM_ins_l71_Gate004_Wire, LM_ins_l71_Overlay, LM_ins_l71_Parts, LM_ins_l71_Post1, LM_ins_l71_Post2, LM_ins_l71_sw25, LM_ins_l71_sw29, LM_ins_l71_sw37, LM_ins_l71_sw38, LM_ins_l71_sw39, LM_ins_l71_sw46, LM_ins_l72_Overlay, LM_ins_l72_Parts, LM_ins_l72_Playfield, LM_ins_l73_Playfield, LM_ins_l74_LFlip, LM_ins_l74_Playfield, LM_ins_l75_Playfield, LM_ins_l75_RFlip, LM_ins_l76_LFlip, LM_ins_l76_Parts, LM_ins_l76_Playfield, LM_ins_l76_Wall91, LM_ins_l77_Parts, LM_ins_l77_Playfield, LM_ins_l77_RFlip, LM_ins_l77_Wall049, LM_ins_l78_Playfield, LM_ins_l8_Overlay, LM_ins_l8_Parts, LM_ins_l8_Playfield, _
  LM_ins_l80_LFlip, LM_ins_l80_Playfield, LM_ins_l80_RFlip, LM_ins_l9_Overlay, LM_ins_l9_Parts, LM_ins_l9_Playfield)
Dim BG_All: BG_All=Array(BM_Bumper001_Ring, BM_Bumper002_Ring, BM_Bumper003_Ring, BM_Gate001_Wire, BM_Gate002_Wire, BM_Gate003_Wire, BM_Gate004_Wire, BM_Gate005_Wire, BM_LFlip, BM_Lemk, BM_Overlay, BM_Parts, BM_Playfield, BM_Post1, BM_Post2, BM_RFlip, BM_RPGun, BM_Remk, BM_Wall049, BM_Wall91, BM_lrail, BM_lsling, BM_lsling_001, BM_lsling_002, BM_rrail, BM_rsling, BM_rsling_001, BM_rsling_002, BM_sw24, BM_sw25, BM_sw28, BM_sw29, BM_sw37, BM_sw38, BM_sw39, BM_sw46, BM_sw57, BM_sw58, BM_sw60, BM_sw61, LM_flash_F26_Bumper002_Ring, LM_flash_F26_Overlay, LM_flash_F26_Parts, LM_flash_F26_Playfield, LM_flash_F26_Post1, LM_flash_F26_Post2, LM_flash_F26_sw25, LM_flash_F26_sw28, LM_flash_F26_sw29, LM_flash_f27_LFlip, LM_flash_f27_Overlay, LM_flash_f27_Parts, LM_flash_f27_Playfield, LM_flash_f27_RFlip, LM_flash_f27_RPGun, LM_flash_f27_Wall049, LM_flash_f27_rsling_002, LM_flash_f27_sw60, LM_flash_f27_sw61, LM_flash_f28_Parts, LM_flash_f28_Playfield, LM_flash_f29_Parts, LM_flash_f30_Bumper002_Ring, _
  LM_flash_f30_Bumper003_Ring, LM_flash_f30_Gate004_Wire, LM_flash_f30_Overlay, LM_flash_f30_Parts, LM_flash_f30_Playfield, LM_flash_f30_Post1, LM_flash_f30_Post2, LM_flash_f30_sw28, LM_flash_f30_sw29, LM_flash_f30_sw37, LM_flash_f30_sw38, LM_flash_f30_sw39, LM_flash_f31_Gate003_Wire, LM_flash_f31_Overlay, LM_flash_f31_Parts, LM_flash_f31_Playfield, LM_flash_f31_Post2, LM_flash_f31_lsling_001, LM_flash_f31_sw38, LM_flash_f32_LFlip, LM_flash_f32_Overlay, LM_flash_f32_Parts, LM_flash_f32_Playfield, LM_flash_f32_RFlip, LM_flash_f32_RPGun, LM_flash_f32_lsling_001, LM_flash_f32_rsling_001, LM_flash_f32_sw38, LM_flash_f32_sw58, LM_flash_f32_sw61, LM_gi_gi001_Overlay, LM_gi_gi001_Parts, LM_gi_gi001_Playfield, LM_gi_gi001_RFlip, LM_gi_gi001_Remk, LM_gi_gi001_Wall049, LM_gi_gi001_rsling, LM_gi_gi001_rsling_001, LM_gi_gi001_rsling_002, LM_gi_gi001_sw24, LM_gi_gi001_sw38, LM_gi_gi001_sw46, LM_gi_gi001_sw60, LM_gi_gi001_sw61, LM_gi_gi002_Overlay, LM_gi_gi002_Parts, LM_gi_gi002_Playfield, LM_gi_gi002_RFlip, _
  LM_gi_gi002_Remk, LM_gi_gi002_rsling, LM_gi_gi002_rsling_001, LM_gi_gi002_rsling_002, LM_gi_gi002_sw61, LM_gi_gi003_Overlay, LM_gi_gi003_Parts, LM_gi_gi003_Playfield, LM_gi_gi003_Wall049, LM_gi_gi003_sw61, LM_gi_gi004_Overlay, LM_gi_gi004_Parts, LM_gi_gi004_Playfield, LM_gi_gi004_Wall049, LM_gi_gi004_sw60, LM_gi_gi004_sw61, LM_gi_gi005_Lemk, LM_gi_gi005_Overlay, LM_gi_gi005_Parts, LM_gi_gi005_Playfield, LM_gi_gi005_Wall91, LM_gi_gi005_rsling_002, LM_gi_gi005_sw58, LM_gi_gi006_Lemk, LM_gi_gi006_Overlay, LM_gi_gi006_Parts, LM_gi_gi006_Playfield, LM_gi_gi006_Wall91, LM_gi_gi006_rsling_002, LM_gi_gi006_sw57, LM_gi_gi006_sw58, LM_gi_gi007_LFlip, LM_gi_gi007_Lemk, LM_gi_gi007_Overlay, LM_gi_gi007_Parts, LM_gi_gi007_Playfield, LM_gi_gi007_lsling, LM_gi_gi007_lsling_001, LM_gi_gi007_lsling_002, LM_gi_gi007_rsling_002, LM_gi_gi007_sw57, LM_gi_gi007_sw58, LM_gi_gi008_LFlip, LM_gi_gi008_Lemk, LM_gi_gi008_Overlay, LM_gi_gi008_Parts, LM_gi_gi008_Playfield, LM_gi_gi008_lsling, LM_gi_gi008_lsling_001, _
  LM_gi_gi008_lsling_002, LM_gi_gi008_sw58, LM_gi_gi010_Parts, LM_gi_gi010_Playfield, LM_gi_gi011_Parts, LM_gi_gi011_Playfield, LM_gi_gi011_rrail, LM_gi_gi013_Parts, LM_gi_gi013_Playfield, LM_gi_gi014_Parts, LM_gi_gi014_Playfield, LM_gi_gi025_Parts, LM_gi_gi025_Playfield, LM_gi_gi027_Bumper002_Ring, LM_gi_gi027_Bumper003_Ring, LM_gi_gi027_Gate001_Wire, LM_gi_gi027_Parts, LM_gi_gi027_Playfield, LM_gi_gi027_Post2, LM_gi_gi027_sw37, LM_gi_gi027_sw38, LM_gi_gi027_sw39, LM_gi_gi029_Bumper003_Ring, LM_gi_gi029_Gate002_Wire, LM_gi_gi029_Parts, LM_gi_gi029_Playfield, LM_gi_gi029_sw38, LM_gi_gi029_sw39, LM_gi_gi031_Gate002_Wire, LM_gi_gi031_Parts, LM_gi_gi031_Playfield, LM_gi_gi031_sw39, LM_gi_gi034_Parts, LM_gi_gi034_Playfield, LM_gi_gi1_Bumper001_Ring, LM_gi_gi1_Bumper002_Ring, LM_gi_gi1_Bumper003_Ring, LM_gi_gi1_Gate001_Wire, LM_gi_gi1_Gate004_Wire, LM_gi_gi1_Overlay, LM_gi_gi1_Parts, LM_gi_gi1_Playfield, LM_gi_gi1_Post1, LM_gi_gi1_Post2, LM_gi_gi1_sw25, LM_gi_gi1_sw28, LM_gi_gi1_sw29, LM_gi_gi1_sw37, LM_gi_gi1_sw38, _
  LM_gi_gi1_sw39, LM_gi_gi1_sw46, LM_gi_gi36_Bumper002_Ring, LM_gi_gi36_Gate001_Wire, LM_gi_gi36_Parts, LM_gi_gi36_Playfield, LM_gi_gi36_sw37, LM_ins_l1_Playfield, LM_ins_l10_Parts, LM_ins_l10_Playfield, LM_ins_l11_Parts, LM_ins_l11_Playfield, LM_ins_l12_Parts, LM_ins_l12_Playfield, LM_ins_l13_Parts, LM_ins_l13_Playfield, LM_ins_l14_Parts, LM_ins_l14_Playfield, LM_ins_l15_Parts, LM_ins_l15_Playfield, LM_ins_l16_Parts, LM_ins_l16_Playfield, LM_ins_l17_LFlip, LM_ins_l17_Parts, LM_ins_l17_Playfield, LM_ins_l18_Overlay, LM_ins_l18_Parts, LM_ins_l18_Playfield, LM_ins_l18_lsling_001, LM_ins_l19_Overlay, LM_ins_l19_Parts, LM_ins_l19_Playfield, LM_ins_l19_sw58, LM_ins_l2_Playfield, LM_ins_l20_Parts, LM_ins_l20_Playfield, LM_ins_l21_Parts, LM_ins_l21_Playfield, LM_ins_l22_Parts, LM_ins_l22_Playfield, LM_ins_l23_Overlay, LM_ins_l23_Parts, LM_ins_l23_Playfield, LM_ins_l24_Parts, LM_ins_l24_Playfield, LM_ins_l25_Parts, LM_ins_l25_Playfield, LM_ins_l25_RFlip, LM_ins_l25_rsling_001, LM_ins_l26_Overlay, LM_ins_l26_Parts, _
  LM_ins_l26_Playfield, LM_ins_l26_rsling_001, LM_ins_l26_rsling_002, LM_ins_l26_sw61, LM_ins_l27_Overlay, LM_ins_l27_Parts, LM_ins_l27_Playfield, LM_ins_l27_rsling, LM_ins_l27_rsling_001, LM_ins_l27_sw61, LM_ins_l28_Parts, LM_ins_l28_Playfield, LM_ins_l29_Parts, LM_ins_l29_Playfield, LM_ins_l3_Playfield, LM_ins_l30_Parts, LM_ins_l30_Playfield, LM_ins_l31_Parts, LM_ins_l31_Playfield, LM_ins_l32_Parts, LM_ins_l32_Playfield, LM_ins_l33_Overlay, LM_ins_l33_Parts, LM_ins_l33_Playfield, LM_ins_l33_Wall91, LM_ins_l34_Overlay, LM_ins_l34_Parts, LM_ins_l34_Playfield, LM_ins_l34_Wall91, LM_ins_l34_lsling, LM_ins_l34_lsling_001, LM_ins_l34_lsling_002, LM_ins_l35_Overlay, LM_ins_l35_Parts, LM_ins_l35_Playfield, LM_ins_l35_Wall049, LM_ins_l35_rsling, LM_ins_l35_rsling_001, LM_ins_l35_rsling_002, LM_ins_l36_Parts, LM_ins_l36_Playfield, LM_ins_l36_Wall049, LM_ins_l37_Lemk, LM_ins_l37_Overlay, LM_ins_l37_Parts, LM_ins_l37_Playfield, LM_ins_l37_Wall91, LM_ins_l37_sw57, LM_ins_l37_sw58, LM_ins_l37a_LFlip, LM_ins_l37a_Playfield, _
  LM_ins_l37a_RFlip, LM_ins_l37b_Playfield, LM_ins_l37b_RFlip, LM_ins_l38_Gate001_Wire, LM_ins_l38_Parts, LM_ins_l38_Playfield, LM_ins_l38_Post2, LM_ins_l38_sw37, LM_ins_l39_Parts, LM_ins_l39_Playfield, LM_ins_l39_sw38, LM_ins_l4_Playfield, LM_ins_l40_Parts, LM_ins_l40_Playfield, LM_ins_l40_sw39, LM_ins_l41_Overlay, LM_ins_l41_Parts, LM_ins_l41_Playfield, LM_ins_l42_Overlay, LM_ins_l42_Parts, LM_ins_l42_Playfield, LM_ins_l43_Overlay, LM_ins_l43_Parts, LM_ins_l43_Playfield, LM_ins_l44_Overlay, LM_ins_l44_Parts, LM_ins_l44_Playfield, LM_ins_l45_Overlay, LM_ins_l45_Parts, LM_ins_l45_Playfield, LM_ins_l46_Overlay, LM_ins_l46_Parts, LM_ins_l46_Playfield, LM_ins_l47_Overlay, LM_ins_l47_Parts, LM_ins_l47_Playfield, LM_ins_l47_sw25, LM_ins_l48_Overlay, LM_ins_l48_Parts, LM_ins_l48_Playfield, LM_ins_l49_Overlay, LM_ins_l49_Parts, LM_ins_l49_Playfield, LM_ins_l5_Playfield, LM_ins_l50_Overlay, LM_ins_l50_Parts, LM_ins_l50_Playfield, LM_ins_l51_Overlay, LM_ins_l51_Parts, LM_ins_l51_Playfield, LM_ins_l52_Parts, _
  LM_ins_l52_Playfield, LM_ins_l53_Parts, LM_ins_l54_Parts, LM_ins_l55_Parts, LM_ins_l55_Playfield, LM_ins_l56_Parts, LM_ins_l56_Playfield, LM_ins_l57_Overlay, LM_ins_l57_Parts, LM_ins_l57_Playfield, LM_ins_l57_sw25, LM_ins_l58_Overlay, LM_ins_l58_Parts, LM_ins_l58_Playfield, LM_ins_l59_Overlay, LM_ins_l59_Parts, LM_ins_l59_Playfield, LM_ins_l59_sw25, LM_ins_l6_Parts, LM_ins_l6_Playfield, LM_ins_l60_Parts, LM_ins_l60_Playfield, LM_ins_l61_Parts, LM_ins_l61_Playfield, LM_ins_l62_Parts, LM_ins_l62_Playfield, LM_ins_l64_Overlay, LM_ins_l64_Parts, LM_ins_l65_Parts, LM_ins_l65_Playfield, LM_ins_l66_Parts, LM_ins_l66_Playfield, LM_ins_l67_Parts, LM_ins_l67_Playfield, LM_ins_l68_Parts, LM_ins_l68_Playfield, LM_ins_l69_Parts, LM_ins_l69_Playfield, LM_ins_l7_Parts, LM_ins_l7_Playfield, LM_ins_l71_Bumper001_Ring, LM_ins_l71_Bumper002_Ring, LM_ins_l71_Bumper003_Ring, LM_ins_l71_Gate004_Wire, LM_ins_l71_Overlay, LM_ins_l71_Parts, LM_ins_l71_Post1, LM_ins_l71_Post2, LM_ins_l71_sw25, LM_ins_l71_sw29, LM_ins_l71_sw37, _
  LM_ins_l71_sw38, LM_ins_l71_sw39, LM_ins_l71_sw46, LM_ins_l72_Overlay, LM_ins_l72_Parts, LM_ins_l72_Playfield, LM_ins_l73_Playfield, LM_ins_l74_LFlip, LM_ins_l74_Playfield, LM_ins_l75_Playfield, LM_ins_l75_RFlip, LM_ins_l76_LFlip, LM_ins_l76_Parts, LM_ins_l76_Playfield, LM_ins_l76_Wall91, LM_ins_l77_Parts, LM_ins_l77_Playfield, LM_ins_l77_RFlip, LM_ins_l77_Wall049, LM_ins_l78_Playfield, LM_ins_l8_Overlay, LM_ins_l8_Parts, LM_ins_l8_Playfield, LM_ins_l80_LFlip, LM_ins_l80_Playfield, LM_ins_l80_RFlip, LM_ins_l9_Overlay, LM_ins_l9_Parts, LM_ins_l9_Playfield)
' VLM  Arrays - End



'************
' Table init
'************

Dim bsTrough, cbCaptive, bsTX, bsVuk, dtBank, GunMech, plungerIM

Sub table1_Init
    vpmInit me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Terminator 3 - Stern 2003" & vbNewLine & "VPX table by JPSalas, Siggi, TastyWasps"
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = VarHidden
        .Games(cGameName).Settings.Value("rol") = 0   '1= rotated display, 0= normal
        .Games(cGameName).Settings.Value("sound") = 1 '1 enabled rom sound
        '.SetDisplayPosition 0,0,GetPlayerHWnd 'uncomment if you can't see the dmd
        On Error Resume Next
        Controller.SolMask(0) = 0
        vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the Timer to renable all the solenoids after 2 seconds
        Controller.Run GetPlayerHWnd
        On Error Goto 0
    End With

  vpmMapLights AllLamps

    ' Nudging
    vpmNudge.TiltSwitch = 56
    vpmNudge.Sensitivity = 3
    vpmNudge.TiltObj = Array(Bumper001, Bumper002, Bumper003, LeftSlingshot, RightSlingshot)

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0, 14, 13, 12, 11, 0, 0, 0
        .InitKick BallRelease, 90, 4
        .InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .Balls = 4
        .IsTrough = 1
    End With

    'Right VUK
    Set bsVuk = new cvpmballstack
    With bsVuk
 .KickForceVar = 3
 .KickAngleVar = 3
        .InitSw 0, 36, 0, 0, 0, 0, 0, 0
        .InitKick sw36a, 180, 8
        .InitExitSnd SoundFX("Saucer_Kick", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    End With

    'TX VUK
    Set bsTX = new cvpmballstack
    With bsTX
 .KickForceVar = 3
 .KickAngleVar = 3
        .InitSw 0, 33, 34, 35, 0, 0, 0, 0
        .InitKick sw31a, 180, 12
        .InitExitSnd SoundFX("Saucer_Kick", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    End With

    'Drop target
    Set dtBank = New cvpmDropTarget
    With dtBank
        .InitDrop sw25, 25
        .InitSnd SoundFX("fx_droptarget", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
        .CreateEvents "dtBank"
    End With

    'Captive Ball
    Set cbCaptive = New cvpmCaptiveBall
    With cbCaptive
        .InitCaptive CapTrigger1, CapWall1, CapKicker1, 358
        .NailedBalls = 0
        .ForceTrans = .9
        .MinForce = 3.5
        .CreateEvents "cbCaptive"
        .Start
    End With

    ' Impulse Plunger
    Const IMPowerSetting = 60 'Plunger Power
    Const IMTime = 0.6        ' Time in seconds for Full Plunge
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP sw16, IMPowerSetting, IMTime
        .Random 0.4
        .switch 16
        .InitExitSnd SoundFX("fx_plunger", DOFContactors), SoundFX("fx_plunger", DOFContactors)
        .CreateEvents "plungerIM"
    End With

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1
    RealTime.Enabled = 1

    ' Init GPG
    RPGKicker.Createsizedball(RPGBallSize / 2)

    ' Init Kickback
    KickBack.Pullback

    Div1.Isdropped = 1
    Post1.Isdropped = 1
    Post2.Isdropped = 1

    LoadLUT

  Flasher126 0
  Flasher127 0
  Flasher128 0
  Flasher129 0
  Flasher130 0
  Flasher131 0
  Flasher132 0

End Sub

'******************
' RealTime Updates
'******************

Sub RealTime_Timer
End Sub

Sub LeftFlipper_Animate
  Dim a : a = LeftFlipper.CurrentAngle
  Dim BP
  For Each BP in BP_LFlip: BP.RotZ = a-121: Next
End Sub

Sub RightFlipper_Animate
  Dim a : a = RightFlipper.CurrentAngle
  Dim BP
  For Each BP in BP_RFlip: BP.RotZ = a+121: Next
End Sub


'******************************************************
'  ZANI: Misc Animations
'******************************************************

' Switches

' Rollovers

Sub sw24_Animate
  Dim z : z = sw24.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw24 : BP.transz = z: Next
End Sub

Sub sw28_Animate
  Dim z : z = sw28.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw28 : BP.transz = z: Next
End Sub

Sub sw29_Animate
  Dim z : z = sw29.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw29 : BP.transz = z: Next
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

Sub sw46_Animate
  Dim z : z = sw46.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw46 : BP.transz = z: Next
End Sub

Sub sw57_Animate
  Dim z : z = sw57.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw57 : BP.transz = z: Next
End Sub

Sub sw58_Animate
  Dim z : z = sw58.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw58 : BP.transz = z: Next
End Sub

Sub sw60_Animate
  Dim z : z = sw60.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw60 : BP.transz = z: Next
End Sub

Sub sw61_Animate
  Dim z : z = sw61.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw61 : BP.transz = z: Next
End Sub

' Gates_hit
Sub Gate001_Animate
  Dim a: a = Gate001.CurrentAngle
  Dim BP : For Each BP in BP_Gate001_Wire: BP.RotX = a: Next
End Sub

Sub Gate002_Animate
  Dim a: a = Gate002.CurrentAngle
  Dim BP : For Each BP in BP_Gate002_Wire: BP.RotX = a: Next
End Sub

Sub Gate003_Animate
  Dim a: a = -Gate003.CurrentAngle
  Dim BP : For Each BP in BP_Gate003_Wire: BP.RotX = a: Next
End Sub

Sub Gate004_Animate
  Dim a: a = Gate004.CurrentAngle
  Dim BP : For Each BP in BP_Gate004_Wire: BP.RotX = a: Next
End Sub

Sub Gate005_Animate
  Dim a: a = Gate005.CurrentAngle
  Dim BP : For Each BP in BP_Gate005_Wire: BP.RotX = a: Next
End Sub



' Bumpers
Sub bumper001_Animate
  Dim z, BP
  z = bumper001.CurrentRingOffset
  For Each BP in BP_Bumper001_Ring : BP.transz = z: Next
End Sub

Sub bumper002_Animate
  Dim z, BP
  z = bumper002.CurrentRingOffset
  For Each BP in BP_Bumper002_Ring : BP.transz = z: Next
End Sub

Sub bumper003_Animate
  Dim z, BP
  z = bumper003.CurrentRingOffset
  For Each BP in BP_Bumper003_Ring : BP.transz = z: Next
End Sub



'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)

    If keycode = PlungerKey Then
    PlaySoundat "fx_solenoidOn", sw16
    controller.switch(55) = True
  End If

  If keycode = LeftFlipperKey Then
    FlipperActivate LeftFlipper, LFPress
    Pincab_LeftFlipperButton.X = Pincab_LeftFlipperButton.X + 8
  End If

  If keycode = RightFlipperKey Then
    FlipperActivate RightFlipper, RFPress
    Pincab_RightFlipperButton.X = Pincab_RightFlipperButton.X - 8
  End If

    If keycode = LeftTiltKey Then Nudge 90, 3::SoundNudgeLeft()
    If keycode = RightTiltKey Then Nudge 270, 3::SoundNudgeRight()
    If keycode = CenterTiltKey Then Nudge 0, 5:SoundNudgeCenter()

    If keycode = RightMagnaSave Then
        controller.switch(53) = true
        If bLutActive Then NextLUT:End If
    End If

  If keycode = keyInsertCoin1 or keycode = keyInsertCoin2 or keycode = keyInsertCoin3 or keycode = keyInsertCoin4 Then
    Select Case Int(rnd*3)
      Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
      Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
      Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
    End Select
  End If

  If keycode = StartGameKey Then soundStartButton()

    If vpmKeyDown(keycode) Then Exit Sub

End Sub

Sub table1_KeyUp(ByVal Keycode)

    If keycode = PlungerKey Then
    PlaySoundAt "fx_solenoidOff", sw16
    controller.switch(55) = False
  End If

  If keycode = LeftFlipperKey Then
    FlipperDeActivate LeftFlipper, LFPress
    Pincab_LeftFlipperButton.X = Pincab_LeftFlipperButton.X - 8
  End If

  If keycode = RightFlipperKey Then
    FlipperDeActivate RightFlipper, RFPress
    Pincab_RightFlipperButton.X = Pincab_RightFlipperButton.X + 8
  End If

    If keycode = RightMagnaSave Then controller.switch(53) = true

    If vpmKeyUp(keycode)Then Exit Sub

End Sub

'***************************
'   LUT - Darkness control
'***************************

Dim bLutActive, LUTImage

Sub LoadLUT
    bLutActive = False
    x = LoadValue(cGameName, "LUTImage")
    If(x <> "")Then LUTImage = x Else LUTImage = 0
    UpdateLUT
End Sub

Sub SaveLUT
    SaveValue cGameName, "LUTImage", LUTImage
End Sub

Sub NextLUT:LUTImage = (LUTImage + 1)MOD 15:UpdateLUT:SaveLUT:Lutbox.text = "level of darkness " & LUTImage + 1:End Sub

Sub UpdateLUT
    Select Case LutImage
        Case 0:table1.ColorGradeImage = "LUT0":GiIntensity = 1:ChangeGIIntensity 1
        Case 1:table1.ColorGradeImage = "LUT1":GiIntensity = 1.05:ChangeGIIntensity 1
        Case 2:table1.ColorGradeImage = "LUT2":GiIntensity = 1.1:ChangeGIIntensity 1
        Case 3:table1.ColorGradeImage = "LUT3":GiIntensity = 1.15:ChangeGIIntensity 1
        Case 4:table1.ColorGradeImage = "LUT4":GiIntensity = 1.2:ChangeGIIntensity 1
        Case 5:table1.ColorGradeImage = "LUT5":GiIntensity = 1.25:ChangeGIIntensity 1
        Case 6:table1.ColorGradeImage = "LUT6":GiIntensity = 1.3:ChangeGIIntensity 1
        Case 7:table1.ColorGradeImage = "LUT7":GiIntensity = 1.35:ChangeGIIntensity 1
        Case 8:table1.ColorGradeImage = "LUT8":GiIntensity = 1.4:ChangeGIIntensity 1
        Case 9:table1.ColorGradeImage = "LUT9":GiIntensity = 1.45:ChangeGIIntensity 1
        Case 10:table1.ColorGradeImage = "LUT10":GiIntensity = 1.5:ChangeGIIntensity 1
        Case 11:table1.ColorGradeImage = "LUT11":GiIntensity = 1.55:ChangeGIIntensity 1
        Case 12:table1.ColorGradeImage = "LUT12":GiIntensity = 1.6:ChangeGIIntensity 1
        Case 13:table1.ColorGradeImage = "LUT13":GiIntensity = 1.65:ChangeGIIntensity 1
        Case 14:table1.ColorGradeImage = "LUT14":GiIntensity = 1.7:ChangeGIIntensity 1
    End Select
End Sub

Dim GiIntensity
GiIntensity = 1   'used by the LUT changing to increase the GI lights when the table is darker

Sub ChangeGiIntensity(factor) 'changes the intensity scale
    Dim bulb
    For each bulb in aGiLights
        bulb.IntensityScale = GiIntensity * factor
    Next
End Sub

'*********
' Switches
'*********

' Slings
Dim LStep, RStep

Sub LeftSlingShot_Slingshot
  dim BL
  LS.VelocityCorrect(ActiveBall)
  RandomSoundSlingshotLeft sw58
    DOF 101, DOFPulse
'    LeftSling004.Visible = 1
    'Lemk.RotX = 26
    LStep = 0 : LeftSlingShot_Timer ' Initialize Step to 0
    vpmTimer.PulseSw 59
  For Each BL in BP_lsling : BL.Visible = False: Next
  LeftSlingShot.TimerInterval = 17
    LeftSlingShot.TimerEnabled = 1
End Sub

'BP_lsling_001
dim BLinit
For Each BLinit in BP_lsling_001 : BLinit.Visible = False: Next
For Each BLinit in BP_lsling_002 : BLinit.Visible = False: Next
For Each BLinit in BP_rsling_001 : BLinit.Visible = False: Next
For Each BLinit in BP_rsling_002 : BLinit.Visible = False: Next

Sub LeftSlingShot_Timer
  Dim BL
  Dim x1, x2, y
  x1 = True:x2 = False:y = -22

    Select Case LStep
    Case 1:y = -22
        Case 3:x1 = False:x2 = True:y = -10
        Case 4:x1 = False:x2 = False:y = 0
         For Each BL in BP_lsling : BL.Visible = True: Next
         LeftSlingShot.TimerEnabled = 0
    End Select

  For Each BL in BP_lsling_001 : BL.Visible = x1: Next
  For Each BL in BP_lsling_002 : BL.Visible = x2: Next
  For Each BL in BP_LEMK : BL.transy = y: Next

    LStep = LStep + 1
End Sub


Sub RightSlingShot_Slingshot
  dim BL
  RS.VelocityCorrect(ActiveBall)
    RandomSoundSlingshotRight sw61
    DOF 102, DOFPulse
'    RightSling004.Visible = 1
'    Remk.RotX = 26
    RStep = 0 : RightSlingShot_Timer
    vpmTimer.PulseSw 62
  For Each BL in BP_rsling : BL.Visible = False: Next
  RightSlingShot.TimerInterval = 17
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
  Dim BL
  Dim x1, x2, y
  x1 = True:x2 = False:y = -22

    Select Case RStep
    Case 1:y = -22
        Case 3:x1 = False:x2 = True:y = -10
        Case 4:x1 = False:x2 = False:y = 0
         For Each BL in BP_rsling : BL.Visible = True: Next
         RightSlingShot.TimerEnabled = 0
    End Select

  For Each BL in BP_rsling_001 : BL.Visible = x1: Next
  For Each BL in BP_rsling_002 : BL.Visible = x2: Next
  For Each BL in BP_REMK : BL.transy = y: Next

    RStep = RStep + 1
End Sub

'Sub RightSlingShot_Timer_old
'    Select Case RStep
'        Case 1:RightSLing004.Visible = 0:RightSLing003.Visible = 1:Remk.RotX = 14
'        Case 2:RightSLing003.Visible = 0:RightSLing002.Visible = 1:Remk.RotX = 2
'        Case 3:RightSLing002.Visible = 0:Remk.RotX = -10:RightSlingShot.TimerEnabled = 0
'    End Select
'    RStep = RStep + 1
'End Sub

' Bumpers
Sub Bumper001_Hit:vpmTimer.PulseSw 51:RandomSoundBumperBottom Bumper001:End Sub
Sub Bumper002_Hit:vpmTimer.PulseSw 49:RandomSoundBumperTop Bumper002:End Sub
Sub Bumper003_Hit:vpmTimer.PulseSw 50:RandomSoundBumperMiddle Bumper003:End Sub

' Drain & holes
Sub Drain_Hit
  RandomSoundDrain Drain
  bsTrough.AddBall Me
End Sub

Sub sw31_Hit:SoundSaucerLock:vpmTimer.PulseSw 31:bsTX.AddBall Me:End Sub
Sub sw31b_Hit:SoundSaucerLock:bsTX.AddBall Me:End Sub
Sub sw36_Hit:SoundSaucerLock:bsVuk.AddBall Me:End Sub

' Rollovers
Sub sw24_Hit:Controller.Switch(24) = 1:End Sub
Sub sw24_UnHit:Controller.Switch(24) = 0:End Sub

Sub sw28_Hit:Controller.Switch(28) = 1:End Sub
Sub sw28_UnHit:Controller.Switch(28) = 0:End Sub

Sub sw29_Hit:Controller.Switch(29) = 1:End Sub
Sub sw29_UnHit:Controller.Switch(29) = 0:End Sub

Sub sw30_Hit:Controller.Switch(30) = 1:End Sub
Sub sw30_UnHit:Controller.Switch(30) = 0:End Sub

Sub sw32_Hit:Controller.Switch(32) = 1:End Sub
Sub sw32_UnHit:Controller.Switch(32) = 0:End Sub

Sub sw37_Hit:Controller.Switch(37) = 1:End Sub
Sub sw37_UnHit:Controller.Switch(37) = 0:End Sub

Sub sw38_Hit:Controller.Switch(38) = 1:End Sub
Sub sw38_UnHit:Controller.Switch(38) = 0:End Sub

Sub sw39_Hit:Controller.Switch(39) = 1:End Sub
Sub sw39_UnHit:Controller.Switch(39) = 0:End Sub

Sub sw40_Hit:Controller.Switch(40) = 1:End Sub
Sub sw40_UnHit:Controller.Switch(40) = 0:End Sub

Sub sw46_Hit:Controller.Switch(46) = 1:End Sub
Sub sw46_UnHit:Controller.Switch(46) = 0:End Sub

Sub sw52_Hit:Controller.Switch(52) = 1:End Sub
Sub sw52_UnHit:Controller.Switch(52) = 0:End Sub

Sub sw57_Hit:Controller.Switch(57) = 1:End Sub
Sub sw57_UnHit:Controller.Switch(57) = 0:End Sub

Sub sw58_Hit:Controller.Switch(58) = 1:End Sub
Sub sw58_UnHit:Controller.Switch(58) = 0:End Sub

Sub sw60_Hit:Controller.Switch(60) = 1:End Sub
Sub sw60_UnHit:Controller.Switch(60) = 0:End Sub

Sub sw61_Hit:Controller.Switch(61) = 1:End Sub
Sub sw61_UnHit:Controller.Switch(61) = 0:End Sub

' Targets
Sub sw10_Hit:STHit 10:End Sub
Sub sw17_Hit:STHit 17:End Sub
Sub sw18_Hit:STHit 18:End Sub
Sub sw19_Hit:STHit 19:End Sub
Sub sw20_Hit:STHit 20:End Sub
Sub sw21_Hit:STHit 21:End Sub
Sub sw22_Hit:STHit 22:End Sub
Sub sw23_Hit:vpmTimer.PulseSw 23:End Sub
Sub sw25_Hit:vpmTimer.PulseSw 25:End Sub
Sub sw41_Hit:vpmTimer.PulseSw 41:End Sub
Sub sw42_Hit:vpmTimer.PulseSw 42:End Sub
Sub sw43_Hit:vpmTimer.PulseSw 43:End Sub
Sub sw44_Hit:vpmTimer.PulseSw 44:End Sub
Sub sw45_Hit:vpmTimer.PulseSw 45:End Sub

Dim CaptiveBallSoundFlag: CaptiveBallSoundFlag = 0

Sub CaptiveBallTrigger_Hit
  If CaptiveBallSoundFlag = 0 Then
    PlaySoundAt "Ball_Collide_1", CaptiveBallTrigger
    CaptiveBallSoundFlag = 1
  End If
End Sub

Sub CaptiveBallTrigger_Timer
  CaptiveBallSoundFlag = 0
End Sub

Sub LeftRampStart_Hit
  WireRampOn False
End Sub

Sub RightRampStart_Hit
  WireRampOn False
End Sub

Sub MiddleRampStart_Hit
  WireRampOn False
End Sub

Sub LaunchRampStart_Hit
  WireRampOn False
End Sub

Sub LockRampStart_Hit
  WireRampOn False
End Sub

Sub UpperRampStart_Hit
  WireRampOn False
End Sub

Sub PipeRampStart_Hit
  WireRampOn False
End Sub

'*********
'Solenoids
'*********
Solcallback(1) = "bstrough.solout"
Solcallback(2) = "Auto_plunger"
Solcallback(3) = "dtbank.soldropup"
SolModCallback(4) = "Flasher104" 'RPG Gi
Solcallback(5) = "RPGKick"      ' backbox kicker
Solcallback(8) = "dtBank.solhit 1,"
'9, 10, 11 bumpers
Solcallback(12) = "Solkickback" 'kickback
Solcallback(13) = "bsVuk.solout"
Solcallback(14) = "bsTX.solout"
'15, 16 flippers
'17, 18 slingshots
'19 not used
Solcallback(20) = "RPGMotor" 'RPG Motor
Solcallback(21) = "SolDiv"   'backpanel diverter
Solcallback(22) = "SolPost1" 'Left up post
Solcallback(23) = "SolPost2" 'Center up post
'24 optional
'25 not used
SolModCallback(26) = "Flasher126" 'TX flashers
SolModCallback(27) = "Flasher127" 'backbox left
SolModCallback(28) = "Flasher128" 'backbox right
SolModCallback(29) = "Flasher129" 'Super JP
SolModCallback(30) = "Flasher130" 'backpanel flashers
SolModCallback(31) = "Flasher131" 'middle flashers
SolModCallback(32) = "Flasher132" 'lower flashers

Sub Auto_Plunger(Enabled)
    If Enabled Then
        PlungerIM.AutoFire
    End If
End Sub

Sub SolDiv(Enabled)
  if Enabled then
    PlaySoundAt "TOM_Diverter_UP_2", sw31
  Else
    PlaySoundAt "TOM_Diverter_DOWN_2", sw31
  end if
    Div1.IsDropped = Not Enabled
    Div1a.IsDropped = Enabled
End Sub

Sub SolPost1(Enabled)
'    PlaySoundAt "fx_Diverter", sw31
    If Enabled then
        Post1.IsDropped = 0
    PlaySoundAt "TOM_Diverter_UP_2", sw31
    for each bp2 in BP_Post1
      bp2.transz = 0
    Next
    Else
        Post1.IsDropped = 1
    PlaySoundAt "TOM_Diverter_DOWN_2", sw31
    for each bp2 in BP_Post1
      bp2.transz = -55
    Next
    End If
End Sub

Sub SolPost2(Enabled)
    PlaySoundAt "fx_Diverter", sw36
    If Enabled Then
        Post2.IsDropped = 0
    PlaySoundAt "TOM_Diverter_DOWN_2", sw31
        Post2a.IsDropped = 1
    for each bp2 in BP_Post2
      bp2.transz = -55
    Next
    Else
        Post2.IsDropped = 1
    PlaySoundAt "TOM_Diverter_UP_2", sw31
        Post2a.IsDropped = 0
    for each bp2 in BP_Post2
      bp2.transz = 55
    Next
    End If
End Sub

'RPG

Sub RPGKick(Enabled)
    PlaySoundAt "Saucer_Kick", RPGKicker
    RPGKicker.Kick RPGun.Rotz - 182, 41 'values tested that one can actually aim and hit all the shots // iaakki
End Sub

Sub RPGMotor(Enabled)
    RPGTimer.Enabled = Enabled
End Sub

' RPG animation // iaakki
Dim RPGDir 'Direction and speed
RPGDir = -1

dim RPGobjrot, RPGMin, RPGMax, RPGPOS
RPGobjrot = 0
RPGMin = -275
RPGMax = -310

dim BP2

RPGPOS = RPGun.rotz

for each bp2 in BP_RPGun
  bp2.rotz = RPGPOS
  bp2.z = 35        'gun height
Next


Sub RPGTimer_Timer()
    RPGPOS = RPGPOS + RPGDir
  RPGun.Rotz = RPGPOS
  for each bp2 in BP_RPGun
    bp2.rotz = RPGPOS
  Next
    If RPGun.Rotz < RPGMax Then RPGDir = 1
    If RPGun.Rotz > RPGMin Then RPGDir = -1

End Sub

Sub Solkickback(Enabled)
    If Enabled Then
        PlaySoundAt "fx_plunger", kickback
        Kickback.Fire
    Else
        Kickback.Pullback
    End If
End Sub

'*******************
'  Flipper Subs
'*******************

SolCallback(sLLFlipper) = "SolLFlipper"
SolCallback(sLRFlipper) = "SolRFlipper"

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
  CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
  LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
  CheckLiveCatch Activeball, RightFlipper, RFCount, parm
  RightFlipperCollide parm
End Sub

Sub MiddleRampAccelerate_Hit
  ActiveBall.VelY = ActiveBall.VelY * 1.15
End Sub


'*******************
'  Flashers
'*******************

Sub Flasher104(level):  debug.print "Flasher104 "&level : End Sub

Sub Flasher126(level):  F26.state = level: End Sub
Sub Flasher127(level):  F27.state = level: End Sub
Sub Flasher128(level):  F28.state = level: End Sub
Sub Flasher129(level):  F29.state = level: End Sub
Sub Flasher130(level):  F30.state = level: End Sub
Sub Flasher131(level):  F31.state = level: End Sub
Sub Flasher132(level):  F32.state = level: End Sub


'********************
' GI
'********************

Set GiCallback2 = GetRef("UpdateGI")

Sub UpdateGI(no, level)
' debug.print "UpdateGI no="&no&" level="&level
    For each x in aGiLights
        x.State = level
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
    x.AddPt "Polarity", 1, 0.05, -5.5
    x.AddPt "Polarity", 2, 0.4, -5.5
    x.AddPt "Polarity", 3, 0.6, -5.0
    x.AddPt "Polarity", 4, 0.65, -4.5
    x.AddPt "Polarity", 5, 0.7, -4.0
    x.AddPt "Polarity", 6, 0.75, -3.5
    x.AddPt "Polarity", 7, 0.8, -3.0
    x.AddPt "Polarity", 8, 0.85, -2.5
    x.AddPt "Polarity", 9, 0.9,-2.0
    x.AddPt "Polarity", 10, 0.95, -1.5
    x.AddPt "Polarity", 11, 1, -1.0
    x.AddPt "Polarity", 12, 1.05, -0.5
    x.AddPt "Polarity", 13, 1.1, 0
    x.AddPt "Polarity", 14, 1.3, 0

    x.AddPt "Velocity", 0, 0,    1
    x.AddPt "Velocity", 1, 0.160, 1.06
    x.AddPt "Velocity", 2, 0.410, 1.05
    x.AddPt "Velocity", 3, 0.530, 1'0.982
    x.AddPt "Velocity", 4, 0.702, 0.968
    x.AddPt "Velocity", 5, 0.95,  0.968
    x.AddPt "Velocity", 6, 1.03,  0.945
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
  Dim gBOT
  gBOT = GetBalls

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
'Const EOSTnew = 1 'EM's to late 80's
Const EOSTnew = 0.8 '90's and later
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
    Dim b, gBOT
    gBOT = GetBalls

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
Sub RDampen_Timer
  Cor.Update
End Sub

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

'' The following sub are needed, however they may exist somewhere else in the script. Uncomment below if needed
'Dim PI: PI = 4*Atn(1)
'Function dSin(degrees)
' dsin = sin(degrees * Pi/180)
'End Function
'Function dCos(degrees)
' dcos = cos(degrees * Pi/180)
'End Function
'
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
RubberFlipperSoundFactor = 0.375 / 5      'volume multiplier; must not be zero
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
  On Error Resume Next
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

'******************************************************
'****  END FLEEP MECHANICAL SOUNDS
'******************************************************


'******************************************************
' ZBRL:  BALL ROLLING AND DROP SOUNDS
'******************************************************

' Be sure to call RollingUpdate in a timer with a 10ms interval see the GameTimer_Timer() sub
Sub GameTimer_Timer
  RollingUpdate
  DoSTAnim ' Animate Stand-Up Targets
End Sub

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
  Dim gBOT
  gBOT = GetBalls

  ' stop the sound of deleted balls
  For b = UBound(gBOT) + 1 To tnob - 1
    ' Comment the next line if you are not implementing Dyanmic Ball Shadows
    If AmbientBallShadowOn = 0 Then BallShadowA(b).visible = 0
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
    If AmbientBallShadowOn = 0 Then
      If gBOT(b).Z > 30 Then
        BallShadowA(b).height = gBOT(b).z - BallSize / 4    'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping the ramp
      Else
        BallShadowA(b).height = 0.1
      End If
      BallShadowA(b).Y = gBOT(b).Y + offsetY
      BallShadowA(b).X = gBOT(b).X + offsetX
      BallShadowA(b).visible = 1
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
  On Error Resume Next
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
' ZRST: STAND-UP TARGETS by Rothbauerw
'******************************************************

'Define a variable for each stand-up target
Dim ST17, ST18, ST19, ST20, ST21, ST22, ST10

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

ST17 = Array(sw17, sw17p, 17, 0)
ST18 = Array(sw18, sw18p, 18, 0)
ST19 = Array(sw19, sw19p, 19, 0)
ST20 = Array(sw20, sw20p, 20, 0)
ST21 = Array(sw21, sw21p, 21, 0)
ST22 = Array(sw22, sw22p, 22, 0)
ST10 = Array(sw10, sw10p, 10, 0)


'Add all the Stand-up Target Arrays to Stand-up Target Animation Array
'   STAnimationArray = Array(ST1, ST2, ....)
Dim STArray
STArray = Array(ST17, ST18, ST19, ST20, ST21, ST22, ST10)

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
Const STAnimStep = 0.5  'vpunits per animation step (control return to Start)
Const STMaxOffset = 6   'max vp units target moves when hit

Const STMaxRotStep = 0.05 'dont ask... Well never mind.. These were all in parts object. Cannot animate,.,
Const STMaxRot = -0.6

Const STMass = 0.2    'Mass of the Stand-up Target (between 0 and 1), higher values provide more resistance

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
'   debug.print "animate1: " & prim.transy
    prim.transy = prim.transy - STMaxOffset
    prim.rotz = prim.rotz - STMaxRot
    vpmTimer.PulseSw switch
    STAnimate = 2
'   debug.print "animate1 ---> " & prim.transy
    Exit Function
  ElseIf animate = 2 Then
'   debug.print "animate2"
    prim.transy = prim.transy + STAnimStep
    prim.rotz = prim.rotz + STMaxRotStep
    If prim.transy >= 0 Then
      prim.transy = 0
      prim.rotz = 0
      primary.collidable = 1
      STAnimate = 0
      Exit Function
    Else
      STAnimate = 2
    End If
  End If
End Function


Sub STAction(Switch)
  Select Case Switch
    Case 11
      Addscore 1000
      Flash1 True 'Demo of the flasher
      vpmTimer.AddTimer 150,"Flash1 False'"   'Disable the flash after short time, just like a ROM would do

    Case 12
      Addscore 1000
      Flash2 True 'Demo of the flasher
      vpmTimer.AddTimer 150,"Flash2 False'"   'Disable the flash after short time, just like a ROM would do

    Case 13
      Addscore 1000
      Flash3 True 'Demo of the flasher
      vpmTimer.AddTimer 150,"Flash3 False'"   'Disable the flash after short time, just like a ROM would do
  End Select
End Sub

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
'****   END STAND-UP TARGETS
'******************************************************

'***************************************************************
' ZSHA: VPW DYNAMIC BALL SHADOWS by Iakki, Apophis, and Wylte
'***************************************************************

' *** Timer sub
' The "DynamicBSUpdate" sub should be called by a timer with an interval of -1 (framerate)
Sub FrameTimer_Timer()
  If DynamicBallShadowsOn Or AmbientBallShadowOn Then DynamicBSUpdate 'update ball shadows
End Sub

' *** These are usually defined elsewhere (ballrolling), but activate here if necessary
Const lob = 0
Const gilvl = 1
'Dim tablewidth: tablewidth =

' *** Required Functions, enable these if they are not already present elswhere in your table
Function max(a,b)
  If a > b Then
    max = a
  Else
    max = b
  End If
End Function

'Function Distance(ax,ay,bx,by)
' Distance = SQR((ax - bx)^2 + (ay - by)^2)
'End Function

'Dim PI: PI = 4*Atn(1)

'Function Atn2(dy, dx)
' If dx > 0 Then
'   Atn2 = Atn(dy / dx)
' ElseIf dx < 0 Then
'   If dy = 0 Then
'     Atn2 = pi
'   Else
'     Atn2 = Sgn(dy) * (pi - Atn(Abs(dy / dx)))
'   end if
' ElseIf dx = 0 Then
'   if dy = 0 Then
'     Atn2 = 0
'   else
'     Atn2 = Sgn(dy) * pi / 2
'   end if
' End If
'End Function

'Function AnglePP(ax,ay,bx,by)
' AnglePP = Atn2((by - ay),(bx - ax))*180/PI
'End Function

'****** End Part B:  Code and Functions ******


'****** Part C:  The Magic ******

' *** These define the appearance of shadows in your table  ***

'Ambient (Room light source)
Const AmbientBSFactor = 0.9  '0 To 1, higher is darker
Const AmbientMovement = 1    '1+ higher means more movement as the ball moves left and right
Const offsetX = 0        'Offset x position under ball (These are if you want to change where the "room" light is for calculating the shadow position,)
Const offsetY = 5        'Offset y position under ball (^^for example 5,5 if the light is in the back left corner)

'Dynamic (Table light sources)
Const DynamicBSFactor = 0.90  '0 To 1, higher is darker
Const Wideness = 20      'Sets how wide the dynamic ball shadows can get (20 +5 thinness is technically most accurate for lights at z ~25 hitting a 50 unit ball)
Const Thinness = 5        'Sets minimum as ball moves away from source

' *** Trim or extend these to match the number of balls/primitives/flashers on the table!  (will throw errors if there aren't enough objects)
Dim objrtx1(6), objrtx2(6)
Dim objBallShadow(6)
Dim OnPF(6)
Dim BallShadowA
BallShadowA = Array (BallShadowA0,BallShadowA1,BallShadowA2,BallShadowA3,BallShadowA4, BallShadowA5, BallShadowA6)
Dim DSSources(30), numberofsources', DSGISide(30) 'Adapted for TZ with GI left / GI right

' *** The Shadow Dictionary
Dim bsDict
Set bsDict = New cvpmDictionary
Const bsNone = "None"
Const bsWire = "Wire"
Const bsRamp = "Ramp"
Const bsRampClear = "Clear"

'Initialization
DynamicBSInit

Sub DynamicBSInit()
  Dim iii, source

  'Prepare the shadow objects before play begins
  For iii = 0 To tnob - 1
    Set objrtx1(iii) = Eval("RtxBallShadow" & iii)
    objrtx1(iii).material = "RtxBallShadow" & iii
    objrtx1(iii).z = 1 + iii / 1000 + 0.01  'Separate z for layering without clipping
    objrtx1(iii).visible = 0

    Set objrtx2(iii) = Eval("RtxBall2Shadow" & iii)
    objrtx2(iii).material = "RtxBallShadow2_" & iii
    objrtx2(iii).z = 1 + iii / 1000 + 0.02
    objrtx2(iii).visible = 0

    Set objBallShadow(iii) = Eval("BallShadow" & iii)
    objBallShadow(iii).material = "BallShadow" & iii
    UpdateMaterial objBallShadow(iii).material,1,0,0,0,0,0,AmbientBSFactor,RGB(0,0,0),0,0,False,True,0,0,0,0
    objBallShadow(iii).Z = 1 + iii / 1000 + 0.04
    objBallShadow(iii).visible = 0

    BallShadowA(iii).Opacity = 100 * AmbientBSFactor
    BallShadowA(iii).visible = 0
  Next

  iii = 0

  For Each Source In DynamicSources
    DSSources(iii) = Array(Source.x, Source.y)
    '   If Instr(Source.name , "Left") > 0 Then DSGISide(iii) = 0 Else DSGISide(iii) = 1  'Adapted for TZ with GI left / GI right
    iii = iii + 1
  Next
  numberofsources = iii
End Sub

Sub BallOnPlayfieldNow(onPlayfield, ballNum)  'Whether a ball is currently on the playfield. Only update certain things once, save some cycles
  If onPlayfield Then
    Dim gBOT: gBOT=getballs
    OnPF(ballNum) = True
    bsRampOff gBOT(ballNum).ID
    '   debug.print "Back on PF"
    UpdateMaterial objBallShadow(ballNum).material,1,0,0,0,0,0,AmbientBSFactor,RGB(0,0,0),0,0,False,True,0,0,0,0
    objBallShadow(ballNum).size_x = 5
    objBallShadow(ballNum).size_y = 4.5
    objBallShadow(ballNum).visible = 1
    BallShadowA(ballNum).visible = 0
    BallShadowA(ballNum).Opacity = 100 * AmbientBSFactor
  Else
    OnPF(ballNum) = False
    '   debug.print "Leaving PF"
  End If
End Sub

Sub DynamicBSUpdate
  Dim falloff 'Max distance to light sources, can be changed dynamically if you have a reason
  falloff = 150
  Dim ShadowOpacity1, ShadowOpacity2
  Dim s, LSd, iii
  Dim dist1, dist2, src1, src2
  Dim bsRampType
  Dim gBOT: gBOT=getballs

  'Hide shadow of deleted balls
  For s = UBound(gBOT) + 1 To tnob - 1
    objrtx1(s).visible = 0
    objrtx2(s).visible = 0
    objBallShadow(s).visible = 0
    BallShadowA(s).visible = 0
  Next

  If UBound(gBOT) < lob Then Exit Sub 'No balls in play, exit

  'The Magic happens now
  For s = lob To UBound(gBOT)
    ' *** Normal "ambient light" ball shadow
    'Layered from top to bottom. If you had an upper pf at for example 80 units and ramps even above that, your Elseif segments would be z>110; z<=110 And z>100; z<=100 And z>30; z<=30 And z>20; Else (under 20)

    'Primitive shadow on playfield, flasher shadow in ramps
    If AmbientBallShadowOn = 1 Then
      '** Above the playfield
      If gBOT(s).Z > 30 Then
        If OnPF(s) Then BallOnPlayfieldNow False, s   'One-time update
        bsRampType = getBsRampType(gBOT(s).id)
        '   debug.print bsRampType

        If Not bsRampType = bsRamp Then 'Primitive visible on PF
          objBallShadow(s).visible = 1
          objBallShadow(s).X = gBOT(s).X + (gBOT(s).X - (tablewidth / 2)) / (Ballsize / AmbientMovement) + offsetX
          objBallShadow(s).Y = gBOT(s).Y + offsetY
          objBallShadow(s).size_x = 5 * ((gBOT(s).Z + BallSize) / 80) 'Shadow gets larger and more diffuse as it moves up
          objBallShadow(s).size_y = 4.5 * ((gBOT(s).Z + BallSize) / 80)
          UpdateMaterial objBallShadow(s).material,1,0,0,0,0,0,AmbientBSFactor * (30 / (gBOT(s).Z)),RGB(0,0,0),0,0,False,True,0,0,0,0
        Else 'Opaque, no primitive below
          objBallShadow(s).visible = 0
        End If

        If bsRampType = bsRampClear Or bsRampType = bsRamp Then 'Flasher visible on opaque ramp
          BallShadowA(s).visible = 1
          BallShadowA(s).X = gBOT(s).X + offsetX
          BallShadowA(s).Y = gBOT(s).Y + offsetY + BallSize / 10
          BallShadowA(s).height = gBOT(s).z - BallSize / 4 + s / 1000 'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping the ramp
          If bsRampType = bsRampClear Then BallShadowA(s).Opacity = 50 * AmbientBSFactor
        ElseIf bsRampType = bsWire Or bsRampType = bsNone Then 'Turn it off on wires or falling out of a ramp
          BallShadowA(s).visible = 0
        End If

        '** On pf, primitive only
      ElseIf gBOT(s).Z <= 30 And gBOT(s).Z > 20 Then
        If Not OnPF(s) Then BallOnPlayfieldNow True, s
        objBallShadow(s).X = gBOT(s).X + (gBOT(s).X - (tablewidth / 2)) / (Ballsize / AmbientMovement) + offsetX
        objBallShadow(s).Y = gBOT(s).Y + offsetY
        '   objBallShadow(s).Z = gBOT(s).Z + s/1000 + 0.04    'Uncomment (and adjust If/Elseif height logic) if you want the primitive shadow on an upper/split pf

        '** Under pf, flasher shadow only
      Else
        If OnPF(s) Then BallOnPlayfieldNow False, s
        objBallShadow(s).visible = 0
        BallShadowA(s).visible = 1
        BallShadowA(s).X = gBOT(s).X + offsetX
        BallShadowA(s).Y = gBOT(s).Y + offsetY
        BallShadowA(s).height = gBOT(s).z - BallSize / 4 + s / 1000
      End If

      'Flasher shadow everywhere
    ElseIf AmbientBallShadowOn = 2 Then
      If gBOT(s).Z > 30 Then 'In a ramp
        BallShadowA(s).X = gBOT(s).X + offsetX
        BallShadowA(s).Y = gBOT(s).Y + offsetY + BallSize / 10
        BallShadowA(s).height = gBOT(s).z - BallSize / 4 + s / 1000 'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping the ramp
      ElseIf gBOT(s).Z <= 30 And gBOT(s).Z > 20 Then 'On pf
        BallShadowA(s).visible = 1
        BallShadowA(s).X = gBOT(s).X + (gBOT(s).X - (tablewidth / 2)) / (Ballsize / AmbientMovement) + offsetX
        BallShadowA(s).Y = gBOT(s).Y + offsetY
        BallShadowA(s).height = 1.04 + s / 1000
      Else 'Under pf
        BallShadowA(s).X = gBOT(s).X + offsetX
        BallShadowA(s).Y = gBOT(s).Y + offsetY
        BallShadowA(s).height = gBOT(s).z - BallSize / 4 + s / 1000
      End If
    End If

    ' *** Dynamic shadows
    If DynamicBallShadowsOn Then
      If gBOT(s).Z < 30 And gBOT(s).X < 850 Then 'Parameters for where the shadows can show, here they are not visible above the table (no upper pf) or in the plunger lane
        dist1 = falloff
        dist2 = falloff
        For iii = 0 To numberofsources - 1 'Search the 2 nearest influencing lights
          LSd = Distance(gBOT(s).x, gBOT(s).y, DSSources(iii)(0), DSSources(iii)(1)) 'Calculating the Linear distance to the Source
          If LSd < falloff And gilvl > 0 Then
            '   If LSd < dist2 And ((DSGISide(iii) = 0 And Lampz.State(100)>0) Or (DSGISide(iii) = 1 And Lampz.State(104)>0)) Then  'Adapted for TZ with GI left / GI right
            dist2 = dist1
            dist1 = LSd
            src2 = src1
            src1 = iii
          End If
        Next
        ShadowOpacity1 = 0
        If dist1 < falloff Then
          objrtx1(s).visible = 1
          objrtx1(s).X = gBOT(s).X
          objrtx1(s).Y = gBOT(s).Y
          '   objrtx1(s).Z = gBOT(s).Z - 25 + s/1000 + 0.01 'Uncomment if you want to add shadows to an upper/lower pf
          objrtx1(s).rotz = AnglePP(DSSources(src1)(0), DSSources(src1)(1), gBOT(s).X, gBOT(s).Y) + 90
          ShadowOpacity1 = 1 - dist1 / falloff
          objrtx1(s).size_y = Wideness * ShadowOpacity1 + Thinness
          UpdateMaterial objrtx1(s).material,1,0,0,0,0,0,ShadowOpacity1 * DynamicBSFactor ^ 3,RGB(0,0,0),0,0,False,True,0,0,0,0
        Else
          objrtx1(s).visible = 0
        End If
        ShadowOpacity2 = 0
        If dist2 < falloff Then
          objrtx2(s).visible = 1
          objrtx2(s).X = gBOT(s).X
          objrtx2(s).Y = gBOT(s).Y + offsetY
          '   objrtx2(s).Z = gBOT(s).Z - 25 + s/1000 + 0.02 'Uncomment if you want to add shadows to an upper/lower pf
          objrtx2(s).rotz = AnglePP(DSSources(src2)(0), DSSources(src2)(1), gBOT(s).X, gBOT(s).Y) + 90
          ShadowOpacity2 = 1 - dist2 / falloff
          objrtx2(s).size_y = Wideness * ShadowOpacity2 + Thinness
          UpdateMaterial objrtx2(s).material,1,0,0,0,0,0,ShadowOpacity2 * DynamicBSFactor ^ 3,RGB(0,0,0),0,0,False,True,0,0,0,0
        Else
          objrtx2(s).visible = 0
        End If
        If AmbientBallShadowOn = 1 Then
          'Fades the ambient shadow (primitive only) when it's close to a light
          UpdateMaterial objBallShadow(s).material,1,0,0,0,0,0,AmbientBSFactor * (1 - max(ShadowOpacity1, ShadowOpacity2)),RGB(0,0,0),0,0,False,True,0,0,0,0
        Else
          BallShadowA(s).Opacity = 100 * AmbientBSFactor * (1 - max(ShadowOpacity1, ShadowOpacity2))
        End If
      Else 'Hide dynamic shadows everywhere else, just in case
        objrtx2(s).visible = 0
        objrtx1(s).visible = 0
      End If
    End If
  Next
End Sub

' *** Ramp type definitions

Sub bsRampOnWire()
  If bsDict.Exists(ActiveBall.ID) Then
    bsDict.Item(ActiveBall.ID) = bsWire
  Else
    bsDict.Add ActiveBall.ID, bsWire
  End If
End Sub

Sub bsRampOn()
  If bsDict.Exists(ActiveBall.ID) Then
    bsDict.Item(ActiveBall.ID) = bsRamp
  Else
    bsDict.Add ActiveBall.ID, bsRamp
  End If
End Sub

Sub bsRampOnClear()
  If bsDict.Exists(ActiveBall.ID) Then
    bsDict.Item(ActiveBall.ID) = bsRampClear
  Else
    bsDict.Add ActiveBall.ID, bsRampClear
  End If
End Sub

Sub bsRampOff(idx)
  If bsDict.Exists(idx) Then
    bsDict.Item(idx) = bsNone
  End If
End Sub

Function getBsRampType(id)
  Dim retValue
  If bsDict.Exists(id) Then
    retValue = bsDict.Item(id)
  Else
    retValue = bsNone
  End If
  getBsRampType = retValue
End Function

'****************************************************************
'****  END VPW DYNAMIC BALL SHADOWS by Iakki, Apophis, and Wylte
'****************************************************************

' Thalamus : Exit in a clean and proper way
Sub Table1_exit
  Controller.Pause = False
  Controller.Stop
End Sub

'Spider-Man Vault Edition
'Stern 2016
'Table Recreation by Alessio for Visual Pinball 10

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Includeds fix by DjRobX
' Added InitVpmFFlipsSAM

' Thalamus 2018-09-16 : Improved directional sounds

'Siggi's Spiderman Classic Edition: VPW Mod


Option Explicit
Randomize
SetLocale 1033


'************************************************************************

'GI Color Mod - Choose your own custom color for GI.
'Primary Colors
'Red = 255, 0, 0
'Green = 0, 255, 0
'Blue = 0, 0, 255
'Incandescent = 255, 197, 143
'Refer to https://rgbcolorcode.com for customized color codes

'Enter RGB values below for "BULB" color
GIColorRed       =  255
GIColorGreen     =  195
GIColorBlue      =  100
dim GIColorRedOrig : GIColorRedOrig =  GIColorRed
dim GIColorGreenOrig : GIColorGreenOrig = GIColorGreen
dim GIColorBlueOrig : GIColorBlueOrig = GIColorBlue

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Dim Ballsize,BallMass
BallSize = 50
BallMass = (Ballsize^3)/125000

Dim tablewidth
tablewidth = Table1.width
Dim tableheight
tableheight = Table1.height
Dim BIP             'Balls in play
BIP = 0
Dim BIPL            'Ball in plunger lane
BIPL = False

Dim DesktopMode:DesktopMode = Table1.ShowDT
Dim UseVPMDMD
If RenderingMode = 2 Then UseVPMDMD = True Else UseVPMDMD = DesktopMode

' Option dims
Dim VolumeDial, BallRollVolume, RampRollVolume, StagedFlippers, VRRoom, VRRoomChoice, CabinetMode, LightLevel, ArtSides


'********************
'Standard definitions
'********************
Const UseVPMModSol = 2
Const UseSolenoids = 1
Const UseLamps = 0
Const UseSync = 0
Const HandleMech = 0

'Standard Sounds
Const SSolenoidOn = ""
Const SSolenoidOff = ""
Const SFlipperOn = ""
Const SFlipperOff = ""
Const SCoin = ""

LoadVPM "03060000", "sam.VBS", 3.10

' VLM  Arrays - Start
' Arrays per baked part
Dim BP_Layer0: BP_Layer0=Array(BM_Layer0, LM_FL_Flasherlight1_Layer0, LM_FL_Flasherlight2_Layer0, LM_FL_Flasherlight3_Layer0, LM_FL_Flasherlight4_Layer0, LM_IN_L60_Layer0, LM_IN_L61_Layer0, LM_IN_L62_Layer0, LM_GI_Layer0, LM_IN_MissioneSandman1_Layer0, LM_IN_MissioneSandman2_Layer0, LM_IN_MissioneVenom1_Layer0, LM_IN_MissioneVenom2_Layer0, LM_IN_MissioneVenom3_Layer0)
Dim BP_Layer1: BP_Layer1=Array(BM_Layer1, LM_FL_Flasherlight1_Layer1, LM_FL_Flasherlight2_Layer1, LM_FL_Flasherlight3_Layer1, LM_FL_Flasherlight4_Layer1, LM_IN_L60_Layer1, LM_IN_L61_Layer1, LM_IN_L62_Layer1, LM_GI_Layer1, LM_IN_MissioneSandman1_Layer1, LM_IN_MissioneSandman2_Layer1, LM_IN_MissioneSandman3_Layer1, LM_IN_MissioneVenom1_Layer1, LM_IN_MissioneVenom2_Layer1)
Dim BP_Layer2: BP_Layer2=Array(BM_Layer2, LM_FL_Flasherlight1_Layer2, LM_FL_Flasherlight2_Layer2, LM_FL_Flasherlight3_Layer2, LM_FL_Flasherlight4_Layer2, LM_GI_Layer2, LM_IN_MissioneSandman1_Layer2, LM_IN_MissioneSandman2_Layer2, LM_IN_MissioneVenom1_Layer2, LM_IN_MissioneVenom2_Layer2, LM_IN_MissioneVenom3_Layer2)
Dim BP_Layer3: BP_Layer3=Array(BM_Layer3, LM_FL_Flasherlight1_Layer3, LM_FL_Flasherlight2_Layer3, LM_FL_Flasherlight3_Layer3, LM_FL_Flasherlight4_Layer3, LM_IN_L60_Layer3, LM_GI_Layer3, LM_IN_MissioneSandman1_Layer3, LM_IN_MissioneSandman2_Layer3, LM_IN_MissioneVenom1_Layer3, LM_IN_MissioneVenom2_Layer3, LM_IN_MissioneVenom3_Layer3)
Dim BP_Octopus: BP_Octopus=Array(BM_Octopus, LM_FL_Flasherlight4_Octopus, LM_GI_Octopus)
Dim BP_Overlay: BP_Overlay=Array(BM_Overlay)
Dim BP_Parts: BP_Parts=Array(BM_Parts, LM_FL_Flasherlight1_Parts, LM_FL_Flasherlight2_Parts, LM_FL_Flasherlight3_Parts, LM_FL_Flasherlight4_Parts, LM_IN_L60_Parts, LM_IN_L61_Parts, LM_IN_L62_Parts, LM_GIS_LA008_Parts, LM_GIS_LA010_Parts, LM_GIS_LA012_Parts, LM_GIS_LA013_Parts, LM_GIS_LA019_Parts, LM_GIS_LA020_Parts, LM_GIS_LA021_Parts, LM_GIS_LA022_Parts, LM_GIS_LA023_Parts, LM_GIS_LA028_Parts, LM_GI_Parts, LM_IN_MissioneSandman1_Parts, LM_IN_MissioneSandman2_Parts, LM_IN_MissioneSandman3_Parts, LM_IN_MissioneVenom1_Parts, LM_IN_MissioneVenom2_Parts, LM_IN_MissioneVenom3_Parts)
Dim BP_Parts_not_baked: BP_Parts_not_baked=Array(BM_Parts_not_baked)
Dim BP_Playfield: BP_Playfield=Array(BM_Playfield, LM_FL_Flasherlight1_Playfield, LM_FL_Flasherlight2_Playfield, LM_FL_Flasherlight3_Playfield, LM_FL_Flasherlight4_Playfield, LM_GIS_LA008_Playfield, LM_GIS_LA010_Playfield, LM_GIS_LA012_Playfield, LM_GIS_LA013_Playfield, LM_GIS_LA019_Playfield, LM_GIS_LA020_Playfield, LM_GIS_LA021_Playfield, LM_GIS_LA022_Playfield, LM_GIS_LA023_Playfield, LM_GIS_LA028_Playfield, LM_GI_Playfield)
Dim BP_Sandman: BP_Sandman=Array(BM_Sandman, LM_FL_Flasherlight1_Sandman, LM_FL_Flasherlight2_Sandman, LM_FL_Flasherlight3_Sandman, LM_FL_Flasherlight4_Sandman, LM_IN_L60_Sandman, LM_IN_L62_Sandman, LM_GI_Sandman)
Dim BP_Venom: BP_Venom=Array(BM_Venom, LM_FL_Flasherlight1_Venom, LM_FL_Flasherlight3_Venom, LM_FL_Flasherlight4_Venom, LM_GI_Venom, LM_IN_MissioneVenom3_Venom)
Dim BP_goblin: BP_goblin=Array(BM_goblin, LM_GIS_LA008_goblin, LM_GIS_LA010_goblin, LM_GI_goblin)
' Arrays per lighting scenario
Dim BL_FL_Flasherlight1: BL_FL_Flasherlight1=Array(LM_FL_Flasherlight1_Layer0, LM_FL_Flasherlight1_Layer1, LM_FL_Flasherlight1_Layer2, LM_FL_Flasherlight1_Layer3, LM_FL_Flasherlight1_Parts, LM_FL_Flasherlight1_Playfield, LM_FL_Flasherlight1_Sandman, LM_FL_Flasherlight1_Venom)
Dim BL_FL_Flasherlight2: BL_FL_Flasherlight2=Array(LM_FL_Flasherlight2_Layer0, LM_FL_Flasherlight2_Layer1, LM_FL_Flasherlight2_Layer2, LM_FL_Flasherlight2_Layer3, LM_FL_Flasherlight2_Parts, LM_FL_Flasherlight2_Playfield, LM_FL_Flasherlight2_Sandman)
Dim BL_FL_Flasherlight3: BL_FL_Flasherlight3=Array(LM_FL_Flasherlight3_Layer0, LM_FL_Flasherlight3_Layer1, LM_FL_Flasherlight3_Layer2, LM_FL_Flasherlight3_Layer3, LM_FL_Flasherlight3_Parts, LM_FL_Flasherlight3_Playfield, LM_FL_Flasherlight3_Sandman, LM_FL_Flasherlight3_Venom)
Dim BL_FL_Flasherlight4: BL_FL_Flasherlight4=Array(LM_FL_Flasherlight4_Layer0, LM_FL_Flasherlight4_Layer1, LM_FL_Flasherlight4_Layer2, LM_FL_Flasherlight4_Layer3, LM_FL_Flasherlight4_Octopus, LM_FL_Flasherlight4_Parts, LM_FL_Flasherlight4_Playfield, LM_FL_Flasherlight4_Sandman, LM_FL_Flasherlight4_Venom)
Dim BL_GI: BL_GI=Array(LM_GI_Layer0, LM_GI_Layer1, LM_GI_Layer2, LM_GI_Layer3, LM_GI_Octopus, LM_GI_Parts, LM_GI_Playfield, LM_GI_Sandman, LM_GI_Venom, LM_GI_goblin)
Dim BL_GIS_LA008: BL_GIS_LA008=Array(LM_GIS_LA008_Parts, LM_GIS_LA008_Playfield, LM_GIS_LA008_goblin)
Dim BL_GIS_LA010: BL_GIS_LA010=Array(LM_GIS_LA010_Parts, LM_GIS_LA010_Playfield, LM_GIS_LA010_goblin)
Dim BL_GIS_LA012: BL_GIS_LA012=Array(LM_GIS_LA012_Parts, LM_GIS_LA012_Playfield)
Dim BL_GIS_LA013: BL_GIS_LA013=Array(LM_GIS_LA013_Parts, LM_GIS_LA013_Playfield)
Dim BL_GIS_LA019: BL_GIS_LA019=Array(LM_GIS_LA019_Parts, LM_GIS_LA019_Playfield)
Dim BL_GIS_LA020: BL_GIS_LA020=Array(LM_GIS_LA020_Parts, LM_GIS_LA020_Playfield)
Dim BL_GIS_LA021: BL_GIS_LA021=Array(LM_GIS_LA021_Parts, LM_GIS_LA021_Playfield)
Dim BL_GIS_LA022: BL_GIS_LA022=Array(LM_GIS_LA022_Parts, LM_GIS_LA022_Playfield)
Dim BL_GIS_LA023: BL_GIS_LA023=Array(LM_GIS_LA023_Parts, LM_GIS_LA023_Playfield)
Dim BL_GIS_LA028: BL_GIS_LA028=Array(LM_GIS_LA028_Parts, LM_GIS_LA028_Playfield)
Dim BL_IN_L60: BL_IN_L60=Array(LM_IN_L60_Layer0, LM_IN_L60_Layer1, LM_IN_L60_Layer3, LM_IN_L60_Parts, LM_IN_L60_Sandman)
Dim BL_IN_L61: BL_IN_L61=Array(LM_IN_L61_Layer0, LM_IN_L61_Layer1, LM_IN_L61_Parts)
Dim BL_IN_L62: BL_IN_L62=Array(LM_IN_L62_Layer0, LM_IN_L62_Layer1, LM_IN_L62_Parts, LM_IN_L62_Sandman)
Dim BL_IN_MissioneSandman1: BL_IN_MissioneSandman1=Array(LM_IN_MissioneSandman1_Layer0, LM_IN_MissioneSandman1_Layer1, LM_IN_MissioneSandman1_Layer2, LM_IN_MissioneSandman1_Layer3, LM_IN_MissioneSandman1_Parts)
Dim BL_IN_MissioneSandman2: BL_IN_MissioneSandman2=Array(LM_IN_MissioneSandman2_Layer0, LM_IN_MissioneSandman2_Layer1, LM_IN_MissioneSandman2_Layer2, LM_IN_MissioneSandman2_Layer3, LM_IN_MissioneSandman2_Parts)
Dim BL_IN_MissioneSandman3: BL_IN_MissioneSandman3=Array(LM_IN_MissioneSandman3_Layer1, LM_IN_MissioneSandman3_Parts)
Dim BL_IN_MissioneVenom1: BL_IN_MissioneVenom1=Array(LM_IN_MissioneVenom1_Layer0, LM_IN_MissioneVenom1_Layer1, LM_IN_MissioneVenom1_Layer2, LM_IN_MissioneVenom1_Layer3, LM_IN_MissioneVenom1_Parts)
Dim BL_IN_MissioneVenom2: BL_IN_MissioneVenom2=Array(LM_IN_MissioneVenom2_Layer0, LM_IN_MissioneVenom2_Layer1, LM_IN_MissioneVenom2_Layer2, LM_IN_MissioneVenom2_Layer3, LM_IN_MissioneVenom2_Parts)
Dim BL_IN_MissioneVenom3: BL_IN_MissioneVenom3=Array(LM_IN_MissioneVenom3_Layer0, LM_IN_MissioneVenom3_Layer2, LM_IN_MissioneVenom3_Layer3, LM_IN_MissioneVenom3_Parts, LM_IN_MissioneVenom3_Venom)
Dim BL_world: BL_world=Array(BM_Layer0, BM_Layer1, BM_Layer2, BM_Layer3, BM_Octopus, BM_Overlay, BM_Parts, BM_Parts_not_baked, BM_Playfield, BM_Sandman, BM_Venom, BM_goblin)
' Global arrays
Dim BG_Bakemap: BG_Bakemap=Array(BM_Layer0, BM_Layer1, BM_Layer2, BM_Layer3, BM_Octopus, BM_Overlay, BM_Parts, BM_Parts_not_baked, BM_Playfield, BM_Sandman, BM_Venom, BM_goblin)
Dim BG_Lightmap: BG_Lightmap=Array(LM_FL_Flasherlight1_Layer0, LM_FL_Flasherlight1_Layer1, LM_FL_Flasherlight1_Layer2, LM_FL_Flasherlight1_Layer3, LM_FL_Flasherlight1_Parts, LM_FL_Flasherlight1_Playfield, LM_FL_Flasherlight1_Sandman, LM_FL_Flasherlight1_Venom, LM_FL_Flasherlight2_Layer0, LM_FL_Flasherlight2_Layer1, LM_FL_Flasherlight2_Layer2, LM_FL_Flasherlight2_Layer3, LM_FL_Flasherlight2_Parts, LM_FL_Flasherlight2_Playfield, LM_FL_Flasherlight2_Sandman, LM_FL_Flasherlight3_Layer0, LM_FL_Flasherlight3_Layer1, LM_FL_Flasherlight3_Layer2, LM_FL_Flasherlight3_Layer3, LM_FL_Flasherlight3_Parts, LM_FL_Flasherlight3_Playfield, LM_FL_Flasherlight3_Sandman, LM_FL_Flasherlight3_Venom, LM_FL_Flasherlight4_Layer0, LM_FL_Flasherlight4_Layer1, LM_FL_Flasherlight4_Layer2, LM_FL_Flasherlight4_Layer3, LM_FL_Flasherlight4_Octopus, LM_FL_Flasherlight4_Parts, LM_FL_Flasherlight4_Playfield, LM_FL_Flasherlight4_Sandman, LM_FL_Flasherlight4_Venom, LM_GI_Layer0, LM_GI_Layer1, LM_GI_Layer2, LM_GI_Layer3, LM_GI_Octopus, LM_GI_Parts, _
  LM_GI_Playfield, LM_GI_Sandman, LM_GI_Venom, LM_GI_goblin, LM_GIS_LA008_Parts, LM_GIS_LA008_Playfield, LM_GIS_LA008_goblin, LM_GIS_LA010_Parts, LM_GIS_LA010_Playfield, LM_GIS_LA010_goblin, LM_GIS_LA012_Parts, LM_GIS_LA012_Playfield, LM_GIS_LA013_Parts, LM_GIS_LA013_Playfield, LM_GIS_LA019_Parts, LM_GIS_LA019_Playfield, LM_GIS_LA020_Parts, LM_GIS_LA020_Playfield, LM_GIS_LA021_Parts, LM_GIS_LA021_Playfield, LM_GIS_LA022_Parts, LM_GIS_LA022_Playfield, LM_GIS_LA023_Parts, LM_GIS_LA023_Playfield, LM_GIS_LA028_Parts, LM_GIS_LA028_Playfield, LM_IN_L60_Layer0, LM_IN_L60_Layer1, LM_IN_L60_Layer3, LM_IN_L60_Parts, LM_IN_L60_Sandman, LM_IN_L61_Layer0, LM_IN_L61_Layer1, LM_IN_L61_Parts, LM_IN_L62_Layer0, LM_IN_L62_Layer1, LM_IN_L62_Parts, LM_IN_L62_Sandman, LM_IN_MissioneSandman1_Layer0, LM_IN_MissioneSandman1_Layer1, LM_IN_MissioneSandman1_Layer2, LM_IN_MissioneSandman1_Layer3, LM_IN_MissioneSandman1_Parts, LM_IN_MissioneSandman2_Layer0, LM_IN_MissioneSandman2_Layer1, LM_IN_MissioneSandman2_Layer2, _
  LM_IN_MissioneSandman2_Layer3, LM_IN_MissioneSandman2_Parts, LM_IN_MissioneSandman3_Layer1, LM_IN_MissioneSandman3_Parts, LM_IN_MissioneVenom1_Layer0, LM_IN_MissioneVenom1_Layer1, LM_IN_MissioneVenom1_Layer2, LM_IN_MissioneVenom1_Layer3, LM_IN_MissioneVenom1_Parts, LM_IN_MissioneVenom2_Layer0, LM_IN_MissioneVenom2_Layer1, LM_IN_MissioneVenom2_Layer2, LM_IN_MissioneVenom2_Layer3, LM_IN_MissioneVenom2_Parts, LM_IN_MissioneVenom3_Layer0, LM_IN_MissioneVenom3_Layer2, LM_IN_MissioneVenom3_Layer3, LM_IN_MissioneVenom3_Parts, LM_IN_MissioneVenom3_Venom)
Dim BG_All: BG_All=Array(BM_Layer0, BM_Layer1, BM_Layer2, BM_Layer3, BM_Octopus, BM_Overlay, BM_Parts, BM_Parts_not_baked, BM_Playfield, BM_Sandman, BM_Venom, BM_goblin, LM_FL_Flasherlight1_Layer0, LM_FL_Flasherlight1_Layer1, LM_FL_Flasherlight1_Layer2, LM_FL_Flasherlight1_Layer3, LM_FL_Flasherlight1_Parts, LM_FL_Flasherlight1_Playfield, LM_FL_Flasherlight1_Sandman, LM_FL_Flasherlight1_Venom, LM_FL_Flasherlight2_Layer0, LM_FL_Flasherlight2_Layer1, LM_FL_Flasherlight2_Layer2, LM_FL_Flasherlight2_Layer3, LM_FL_Flasherlight2_Parts, LM_FL_Flasherlight2_Playfield, LM_FL_Flasherlight2_Sandman, LM_FL_Flasherlight3_Layer0, LM_FL_Flasherlight3_Layer1, LM_FL_Flasherlight3_Layer2, LM_FL_Flasherlight3_Layer3, LM_FL_Flasherlight3_Parts, LM_FL_Flasherlight3_Playfield, LM_FL_Flasherlight3_Sandman, LM_FL_Flasherlight3_Venom, LM_FL_Flasherlight4_Layer0, LM_FL_Flasherlight4_Layer1, LM_FL_Flasherlight4_Layer2, LM_FL_Flasherlight4_Layer3, LM_FL_Flasherlight4_Octopus, LM_FL_Flasherlight4_Parts, LM_FL_Flasherlight4_Playfield, _
  LM_FL_Flasherlight4_Sandman, LM_FL_Flasherlight4_Venom, LM_GI_Layer0, LM_GI_Layer1, LM_GI_Layer2, LM_GI_Layer3, LM_GI_Octopus, LM_GI_Parts, LM_GI_Playfield, LM_GI_Sandman, LM_GI_Venom, LM_GI_goblin, LM_GIS_LA008_Parts, LM_GIS_LA008_Playfield, LM_GIS_LA008_goblin, LM_GIS_LA010_Parts, LM_GIS_LA010_Playfield, LM_GIS_LA010_goblin, LM_GIS_LA012_Parts, LM_GIS_LA012_Playfield, LM_GIS_LA013_Parts, LM_GIS_LA013_Playfield, LM_GIS_LA019_Parts, LM_GIS_LA019_Playfield, LM_GIS_LA020_Parts, LM_GIS_LA020_Playfield, LM_GIS_LA021_Parts, LM_GIS_LA021_Playfield, LM_GIS_LA022_Parts, LM_GIS_LA022_Playfield, LM_GIS_LA023_Parts, LM_GIS_LA023_Playfield, LM_GIS_LA028_Parts, LM_GIS_LA028_Playfield, LM_IN_L60_Layer0, LM_IN_L60_Layer1, LM_IN_L60_Layer3, LM_IN_L60_Parts, LM_IN_L60_Sandman, LM_IN_L61_Layer0, LM_IN_L61_Layer1, LM_IN_L61_Parts, LM_IN_L62_Layer0, LM_IN_L62_Layer1, LM_IN_L62_Parts, LM_IN_L62_Sandman, LM_IN_MissioneSandman1_Layer0, LM_IN_MissioneSandman1_Layer1, LM_IN_MissioneSandman1_Layer2, LM_IN_MissioneSandman1_Layer3, _
  LM_IN_MissioneSandman1_Parts, LM_IN_MissioneSandman2_Layer0, LM_IN_MissioneSandman2_Layer1, LM_IN_MissioneSandman2_Layer2, LM_IN_MissioneSandman2_Layer3, LM_IN_MissioneSandman2_Parts, LM_IN_MissioneSandman3_Layer1, LM_IN_MissioneSandman3_Parts, LM_IN_MissioneVenom1_Layer0, LM_IN_MissioneVenom1_Layer1, LM_IN_MissioneVenom1_Layer2, LM_IN_MissioneVenom1_Layer3, LM_IN_MissioneVenom1_Parts, LM_IN_MissioneVenom2_Layer0, LM_IN_MissioneVenom2_Layer1, LM_IN_MissioneVenom2_Layer2, LM_IN_MissioneVenom2_Layer3, LM_IN_MissioneVenom2_Parts, LM_IN_MissioneVenom3_Layer0, LM_IN_MissioneVenom3_Layer2, LM_IN_MissioneVenom3_Layer3, LM_IN_MissioneVenom3_Parts, LM_IN_MissioneVenom3_Venom)
' VLM  Arrays - End


'************
' Table init.
'************

'Const cGameName = "sman_261" 'ENGLISH Normal VERSION

Const cGameName = "smanve_101" 'ENGLISH Vault Edition VERSION

'Const cGameName = "smanve_101c" 'ENGLISH  Vault Edition VERSION DMD COLORED

'Const cGameName = "sman_210ai" 'ITALIAN Normal VERSION

Const tnob = 4 ' total number of balls

'Variables
Dim bsTrough, bsSandman, bsDocOck, x
Dim mag1
Dim PlungerIM
Dim Attendi
Dim DocMagnet

Dim SMBall1, SMBall2, SMBall3, SMBall4, gBOT

Sub Table1_Init
    vpmInit Me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Spider-Man Classic Vault Edition(Stern 2016)" & vbNewLine & "VPW Mod"
        .Games(cGameName).Settings.Value("rol") = 0
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = DesktopMode
        On Error Resume Next
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
        On Error Goto 0
    End With

    Set SMBall1 = Trough1.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set SMBall2 = Trough2.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set SMBall3 = Trough3.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set SMBall4 = Trough4.CreateSizedballWithMass(Ballsize/2,Ballmass)

  gBOT = Array(SMBall1, SMBall2, SMBall3, SMBall4)

  Controller.Switch(18) = 1
  Controller.Switch(19) = 1
  Controller.Switch(20) = 1
  Controller.Switch(21) = 1

  'Sandman VUK
  Set bsSandman = New cvpmSaucer
  bsSandman.InitKicker sw59, 59, 0, 45, 1.56 'ORIGINALE: 59,0,35,1.56
  bsSandman.InitExitVariance 3,4

  'Doc Ock VUK
  Set bsDocOck = New cvpmSaucer
  bsDocOck.InitKicker sw36, 36, 0, 45, 1.56   'ORIGINALE: 36,0,35,1.56
  bsDocOck.InitExitVariance 3,4

  'Doc Ock Magmet
  Set DocMagnet = New cvpmMagnet
  DocMagnet.InitMagnet DocOckMagnet, 50
  DocMagnet.Solenoid = 3
  'DocMagnet.GrabCenter = True
  DocMagnet.CreateEvents "DocMagnet"

  'Loop Diverter
  diverter.IsDropped = 1

    'Nudging
  vpmNudge.TiltSwitch=-7
    vpmNudge.Sensitivity=3
    vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingshot)


    ' Impulse Plunger
    Const IMPowerSetting = 55
    Const IMTime = 0.6
    Set plungerIM = New cvpmImpulseP
    With plungerIM
    .InitImpulseP swplunger, IMPowerSetting, IMTime
    .Switch 23
    .Random 0.3
    .InitExitSnd SoundFX("solenoid", DOFContactors), ""
        .CreateEvents "plungerIM"
    End With
  Attendi=1
  PausaAnimazione.Enabled=1
  Controller.Switch(50)=1
  Controller.Switch(53)=1
  Controller.Switch(57)=1
  BankAlto=1
  SandmanPronto=0
  BIP=0
  SandmanAlto=0
  PallaBucoOctopus=0
  PallaBucoSandman=0
  PallaSuMagnete=0
  'AlzaSandman
  sw36.Enabled=0
  sw59.Enabled=0
  'SetGIColor
  InitVpmFFlipsSAM

  solLSling 1
  solRSling 1
 End Sub


Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub
Sub table1_exit():Controller.Stop:End Sub


''*******************************************
''  ZOPT: User Options
''*******************************************


' Table1.Option arguments are:
' - option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional arry of literal strings

' Called when options are tweaked by the player.
' - 0: game has started, good time to load options and adjust accordingly
' - 1: an option has changed
' - 2: options have been reseted
' - 3: player closed the tweak UI, good time to update staticly prerendered parts
Sub Table1_OptionEvent(ByVal eventId)
    If eventId = 1 Then DisableStaticPreRendering = True

    ' VR Room
  If RenderingMode = 2 Then
    VRRoomChoice = Table1.Option("VR Room", 1, 3, 1, 2, 0, Array("Minimal", "360", "Ultra Minimal"))
    VRRoom = VRRoomChoice
  Else
    VRRoom = 0
  End If

    ' Cabinet Mode
    If RenderingMode = 2 Then
    CabinetMode = 0
  Else
    CabinetMode = Table1.Option("Cabinet Mode", 0, 1, 1, 0, 0, Array("Off", "On"))
  End If

    ' Art Sideblades
    'ArtSides = Table1.Option("Art Sideblades", 0, 1, 1, 0, 0, Array("Off", "On"))
  ArtSides = 0

    ' Sound volumes
    VolumeDial = Table1.Option("Mech Volume", 0, 1, 0.01, 0.8, 1)
    RampRollVolume = Table1.Option("Ramp Volume", 0, 1, 0.01, 0.5, 1)
    BallRollVolume = Table1.Option("Ball Roll Volume", 0, 1, 0.01, 0.5, 1)

  ' Staged flippers
  StagedFlippers = Table1.Option("Staged Flippers", 0, 1, 1, 0, 0, Array("Off", "On"))

  SetupRoom

    If eventId = 3 Then DisableStaticPreRendering = False
End Sub



Sub SetupRoom

  If CabinetMode = 1 Then
    PinCab_Rails.visible = 0

  Else
    PinCab_Rails.visible = 1

  End If

  If Artsides = 1 then
'   PinCab_WoodBlades.visible = 0
'   PinCab_ArtBlades.visible = 0
'   PinCab_ArtBlades.image = "artSidesON"
'   PinCab_ArtBlades.blenddisablelighting = 0.8
  Else
'   PinCab_ArtBlades.visible = 0
'   PinCab_WoodBlades.visible = 0
'   PinCab_WoodBlades.image = "woodSidesON"
'   PinCab_WoodBlades.blenddisablelighting = 0.8
  End if

  Dim VRThings
  If VRRoom > 0 Then
    ScoreText.visible = 0
    If VRRoom = 1 Then
      for each VRThings in VRCab:VRThings.visible = 1:Next
      for each VRThings in VRStuff:VRThings.visible = 1:Next
      for each VRThings in VRMega:VRThings.visible = 0:Next
    End If
    If VRRoom = 2 Then
      for each VRThings in VRCab:VRThings.visible = 1:Next
      for each VRThings in VRStuff:VRThings.visible = 0:Next
      for each VRThings in VRMega:VRThings.visible = 1:Next
            VR_Logo.visible = False
        End If
    If VRRoom = 3 Then
      for each VRThings in VRCab:VRThings.visible = 0:Next
      for each VRThings in VRStuff:VRThings.visible = 0:Next
            for each VRThings in VRMega:VRThings.visible = 0:Next
      PinCab_Backglass.visible = 1
      PinCab_Backbox.visible = 1
      PinCab_Backbox.image = "Pincab_Backbox_Min"
      DMD.visible = 1
    End If
  Else
      for each VRThings in VRCab:VRThings.visible = 0:Next
      for each VRThings in VRStuff:VRThings.visible = 0:Next
      for each VRThings in VRMega:VRThings.visible = 0:Next
      If DesktopMode then ScoreText.visible = 1 else ScoreText.visible = 0 End If
  End if

End Sub






'**********
'Timer Code
'**********

Sub GameTimer_Timer()
  Cor.Update
End Sub

Sub FrameTimer_Timer()
  BSUpdate
  GatesTimer
    UpdateFlipperLogos
  RollingSound
  LampTimer
  VR_Primary_plunger.Y = -50 + (5* Plunger.Position) -20
End Sub


'************************************
' PAUSA ANIMAZIONE DOPO AVVIO TAVOLO
'************************************

Sub PausaAnimazione_Timer()
  Attendi=0
  sw63.IsDropped=1
  Me.Enabled=0
End Sub

Dim bulb


'********************************************
'  Depotenzia magnete dopo presa della palla
'********************************************

Dim PallaSuMagnete

Sub PresenzaSuMagnete_Hit() 'CON PALLA SUL MAGNETE RIDUCE LA POTENZA DELLA CALAMITA
  If DocMagnet.MagnetON= True Then DepotenziaMagnete.Enabled=True
  PallaSuMagnete=1
End Sub


Sub PresenzaSuMagnete_UnHit() 'IN USCITA DAL MAGNETE VIENE RIPRISTINATA LA POTENZA DELLA CALAMITA
  If DocMagnet.MagnetON= False Then DocMagnet.InitMagnet DocOckMagnet, 50
  PallaSuMagnete=0
End Sub


Sub DepotenziaMagnete_timer() 'TIMER PER RIDURRE LA POTENZA DEL MAGNETE
  DocMagnet.MagnetON= False
  DocMagnet.InitMagnet DocOckMagnet, 2
  DocMagnet.MagnetON= True
  Me.Enabled = 0
End Sub


'************************************
'  Lancio palla dopo sgancio magnete
'************************************

Dim Pausa

Sub solDocMagnet(enabled)
    MagnetOffTimer.Enabled = Not enabled
    If enabled Then DocMagnet.MagnetOn = True
End Sub

Sub MagnetOffTimer_Timer
    Dim ball
    For Each ball In DocMagnet.Balls
        With ball
    'RilanciaPallaMagnete.Enabled=1
      .VelX = 15: .VelY = -8: Pausa= 1  'ERA VelX = 15: .VelY = -7
          'If BIP=1 Then .VelX = 15: .VelY = -15: Pausa= 1 'REGOLARE VELY PER AUMENTARE IL LANCIO DELLA PALLA (NEGATIVO PER ANDARE VERSO L'ALTO)
      'If BIP>1 Then .VelX = 15: .VelY = -15: Pausa= 1 'REGOLARE VELY PER AUMENTARE IL LANCIO DELLA PALLA (NEGATIVO PER ANDARE VERSO L'ALTO)
        End With
    Next
    Me.Enabled = False:DocMagnet.MagnetOn = False
End Sub

Sub DocOckMagnet_UnHit()
  If Pausa=1 Then
  DocMagnet.MagnetON= False
  DocMagnet.InitMagnet DocOckMagnet, 0
  DocOckMagnet.Enabled=0
  MagneteDisabilitato.Enabled=1
  End If
End Sub

Sub MagneteDisabilitato_Timer()
  DocMagnet.InitMagnet DocOckMagnet, 50
  Pausa= 0
  DocOckMagnet.Enabled=1
  Me.Enabled=0
End Sub

'**********
' Keys
'**********

Sub Table1_KeyDown(ByVal Keycode)

  If keycode = PlungerKey Then
    Plunger.PullBack
        SoundPlungerPull
  End If

  ' If Keycode = RightFlipperKey then
  '   Controller.Switch(86)=1  ' UR Flipper
  ' End If

     If keycode = LeftTiltKey Then
    Nudge 90, 1
    SoundNudgeLeft
  End If

  If keycode = RightTiltKey Then
    Nudge 270, 1
    SoundNudgeRight
  End If

  If keycode = CenterTiltKey Then
    Nudge 0, 1
    SoundNudgeCenter
  End If

    If keycode = LeftFlipperKey Then
    FlipperActivate LeftFlipper, LFPress
    Pincab_LeftFlipperButton.X = Pincab_LeftFlipperButton.X + 8
  End If

  If keycode = RightFlipperKey Then
    FlipperActivate RightFlipper, RFPress
    Pincab_RightFlipperButton.X = Pincab_RightFlipperButton.X - 8
        If StagedFlippers = 0 Then
            FlipperActivate RightFlipper2, URFPress
    End If
  End If

  If keycode = KeyUpperRight Then
    FlipperActivate RightFlipper2, URFPress
  End If

    If keycode = StartGameKey Then
    SoundStartButton
  End If

  If keycode = keyInsertCoin1 or keycode = keyInsertCoin2 or keycode = keyInsertCoin3 or keycode = keyInsertCoin4 Then 'Use this for ROM based games
    Select Case Int(Rnd * 3)
      Case 0
        PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
      Case 1
        PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
      Case 2
        PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
    End Select
  End If


  If vpmKeyDown(Keycode) Then Exit Sub
 End Sub

Sub Table1_KeyUp(ByVal Keycode)
  If keycode = PlungerKey Then
    Plunger.Fire
    If BIPL = 1 Then
      SoundPlungerReleaseBall()     'Plunger release sound when there is a ball in shooter lane
    Else
      SoundPlungerReleaseNoBall()     'Plunger release sound when there is no ball in shooter lane
    End If
  End If

    If keycode = LeftFlipperKey Then
        FlipperDeActivate LeftFlipper, LFPress
    Pincab_LeftFlipperButton.X = Pincab_LeftFlipperButton.X - 8
    End If

    If keycode = RightFlipperKey Then
        FlipperDeActivate RightFlipper, RFPress
    Pincab_RightFlipperButton.X = Pincab_RightFlipperButton.X + 8
        If StagedFlippers = 0 Then
            FlipperDeActivate RightFlipper2, URFPress
        End If
    End If

  If vpmKeyUp(Keycode) Then Exit Sub
 End Sub

'Realtime updates

Sub GatesTimer()
  GateSWsx.RotZ= -Gate2.currentangle
  GateSWdx.RotZ= -Gate3.currentangle
  GateP0.RotX = -Gate4.currentangle + 90
  GateP1.RotX = -Gate5.currentangle + 90
  GateP2.RotX = -LeftRampEnd.currentangle + 90    'gate V shape
  GateP3.RotX = -LeftRampStart.currentangle + 90
  GateP4.RotX = -RightRampEnd.currentangle +90    'gate V shape
  GateP5.RotX = -RightRampStart.currentangle +90
End Sub

Sub UpdateFlipperLogos
    LFLogo.RotY = LeftFlipper.CurrentAngle
    RFLogo.RotY = RightFlipper.CurrentAngle
  flipperr1.RotY = RightFlipper2.CurrentAngle
    FlipperLSh.RotZ = LeftFlipper.CurrentAngle
  FlipperRSh.RotZ = RightFlipper.CurrentAngle
    FlipperRSh001.RotZ = RightFlipper2.CurrentAngle
End Sub

'Solenoids
SolCallback(1) = "ReleaseBall"
SolCallback(2) = "solAutofire"
SolCallback(3) = "solDocMagnet"
'SolModCallback(3) = "SetModLampmm 0, 200,"

SolCallback(4) = "solDocVUK"
SolCallback(5) = "solDocMotor"
'SolCallback(6) = "ShakerMotor"
SolCallback(7) = "Gate2.open ="
SolCallback(8) = "Gate3.open ="
SolCallback(9) = "solLBump"
SolCallback(10) = "solRBump"
SolCallback(11) = "solBBump"
SolCallback(12) = "solSandVUK"
SolCallback(13) = "solSandMotor"
SolCallback(14) = "solURFlipper"
SolCallback(15) = "SolLFlipper" '"solLFlipper"
SolCallback(16) = "SolRFlipper" '"solRFlipper"
SolCallback(17) = "solLSling"
SolCallback(18) = "solRSling"
SolCallback(19) = "ShakeGoblin"
SolCallback(20) = "sol3Bank"
SolCallback(22) = "solDivert"
SolCallback(24) = "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolModCallBack(21) = "FlashSol21" '"SetModLamp 21," 'DomeRed RedFlasherMid
SolModCallBack(23) = "SetModLamp 23,"
SolModCallBack(25) = "SetModLamp 25,"
SolModCallBack(26) = "SetModLamp 26,"
SolModCallback(27) = "FlashSol27" '"SetModLamp 27, " 'DomeYellow YellowFlasherTop
SolModCallback(28) = "SetModLamp 28, "
SolModCallBack(29) = "FlashSol29" ' "SetModLamp 29," 'DomerBlue BlueFlasherTop
SolModCallBack(30) = "FlashSol30" '"SetModLamp 30," 'DomerRed RedFlasherTop
SolModCallBack(31) = "SetModLamp 31,"


'Solenoid Functions
Sub solTrough(Enabled)
  If Enabled Then
    bsTrough.ExitSol_On
    vpmTimer.PulseSw 22
  End If
 End Sub

Sub solAutofire(Enabled)
  If Enabled Then
    PlungerIM.AutoFire
        RandomSoundAutoPlunger
  End If
 End Sub


Sub solDocMotor(Enabled)
  If Enabled Then
    If sw63.IsDropped Then
    Controller.Switch(58) = 0
    Controller.Switch(57) = 1
    sw63.IsDropped=0
    AbbassaOctopus
  Else
    Controller.Switch(57) = 0
    Controller.Switch(58) = 1
    If Attendi=0 Then AlzaOctopus
    End If
  End If
 End Sub

Sub solSandMotor(Enabled)
  If Enabled Then
    If sw42.IsDropped Then
    Controller.Switch(54) =0
    Controller.Switch(53) =1
    sw42.IsDropped=0
    AbbassaSandman
  Else
    Controller.Switch(53) =0
    Controller.Switch(54) =1
    AlzaSandman
    End If
  End If
 End Sub

Sub Sol3Bank(Enabled)
  If Enabled Then
    If sw11.IsDropped Then
    Controller.Switch(49)=0
    Controller.Switch(50)=1
    ParetiBankSu
    AlzaBank
  Else
    Controller.Switch(50)=0
    Controller.Switch(49)=1
    AbbassaBank
    End If
  End If
End Sub

Sub solSandVUK(Enabled) 'GESTIONE SOLENOIDE Sandman
  If Enabled Then
    bsSandman.ExitSol_On
    SolenoideSandmanAbilitato
    Playsound "solenoid" '
  End If
 End Sub

Sub SolenoideSandmanAbilitato
  SolenoideSandman.TransZ= -50
  SolenoideUscitaSandman.enabled=1
End Sub

Sub SolenoideUscitaSandman_Timer 'GESTIONE PISTONE LANCIO PALLA Octopus
  SolenoideSandman.TransZ= -60
  Playsound "solenoid" '
  Me.Enabled=0
End Sub

Sub solDocVUK(Enabled) 'GESTIONE SOLENOIDE Octopus
  If Enabled Then
    bsDocOck.ExitSol_On
    SolenoideOctopusAbilitato
    Playsound "solenoid" ' TODO
  End If
 End Sub

Sub SolenoideOctopusAbilitato
  SolenoideOctopus.TransZ= -50
  SolenoideUscitaOctopus.enabled=1
End Sub

Sub SolenoideUscitaOctopus_Timer 'GESTIONE PISTONE LANCIO PALLA Octopus
  SolenoideOctopus.TransZ= -59
  PlaysoundAtVol "solenoid", SolenoideOctopus, 1
  Me.Enabled=0
End Sub

Dim LockAttivo

Sub solDivert(Enabled)
  If Enabled Then
    Diverter.IsDropped = 0
    Playsound SoundFX("Diverter",DOFContactors) ' TODO
    LockAttivo = 1
  Else
    Diverter.IsDropped = 1
    Playsound SoundFX("Diverter",DOFContactors) ' TODO
    LockAttivo=0
  End If
 End Sub

'Drains and Kickers
Sub drain_Hit()
  BIP = BIP - 1
  RandomSoundDrain drain
    UpdateTrough
 End Sub

 Sub ReleaseBall(enabled)
  If enabled Then
    Trough4.kick 60, 12
        vpmTimer.PulseSw 23
    UpdateTrough
        RandomSoundBallRelease Trough4
  End If
End Sub

Sub Trough1_Hit   : Controller.Switch(18) = 1 : UpdateTrough : End Sub
Sub Trough1_UnHit : Controller.Switch(18) = 0 : UpdateTrough : End Sub
Sub Trough2_Hit   : Controller.Switch(19) = 1 : UpdateTrough : End Sub
Sub Trough2_UnHit : Controller.Switch(19) = 0 : UpdateTrough : End Sub
Sub Trough3_Hit   : Controller.Switch(20) = 1 : UpdateTrough : End Sub
Sub Trough3_UnHit : Controller.Switch(20) = 0 : UpdateTrough : End Sub
Sub Trough4_Hit   : Controller.Switch(21) = 1 : UpdateTrough : End Sub
Sub Trough4_UnHit : Controller.Switch(21) = 0 : UpdateTrough : End Sub


Sub UpdateTrough
  UpdateTroughTimer.Interval = 300
  UpdateTroughTimer.Enabled = True
End Sub

Sub UpdateTroughTimer_Timer
  If Trough4.BallCntOver = 0 Then Trough3.kick 57, 10
  If Trough3.BallCntOver = 0 Then Trough2.kick 57, 10
  If Trough2.BallCntOver = 0 Then Trough1.kick 57, 10
  If Trough1.BallCntOver = 0 Then Drain.kick 57, 10
  Me.Enabled = False
End Sub

' Sub BallRelease_UnHit()
'   BIP = BIP + 1
'     RandomSoundBallRelease BallRelease
'  End Sub

'************************************************
'************Slingshots Animation****************
'************************************************

Dim RStep, Lstep

Sub LeftSlingShot_Slingshot
  LS.VelocityCorrect(ActiveBall)
  vpmTimer.PulseSw 26
End Sub

Sub RightSlingShot_Slingshot
  RS.VelocityCorrect(ActiveBall)
  vpmTimer.PulseSw 27
End Sub

Sub solLSling(enabled)
  If enabled then
    RandomSoundSlingshotLeft(SLING1)
    LSling.Visible = 0
    LSling1.Visible = 1
    sling1.TransZ = -27
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
  End If
End Sub

Sub solRSling(enabled)
  If enabled then
    RandomSoundSlingshotRight(SLING2)
    RSling.Visible = 0
    RSling1.Visible = 1
    sling2.TransZ = -27
    RStep = 0
    RightSlingShot.TimerEnabled = 1
  End If
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling1.TransZ = -15
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling1.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling2.TransZ = -15
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling2.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

'************************************************
'**************Bumpers Animation*****************
'************************************************

Dim dirRing1:dirRing1 = -1
Dim dirRing2:dirRing2 = -1
Dim dirRing3:dirRing3 = -1

Sub Bumper1_Hit:vpmTimer.PulseSw 30:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 31:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 32:End Sub

'
Sub solLBump (enabled): If enabled then Bumper1.TimerEnabled = 1: RandomSoundBumperTop(Bumper1):End If: End Sub
Sub solRBump (enabled): If enabled then Bumper2.TimerEnabled = 1: RandomSoundBumperMiddle(Bumper2) : End If: End Sub
Sub solBBump (enabled): If enabled then Bumper3.TimerEnabled = 1: RandomSoundBumperBottom(Bumper3) : End If: End Sub

Sub Bumper1_timer()
  BumperRing1.Z = BumperRing1.Z + (5 * dirRing1)
  If BumperRing1.Z <= -40 Then dirRing1 = 1
  If BumperRing1.Z >= 0 Then
    dirRing1 = -1
    BumperRing1.Z = 0
    Me.TimerEnabled = 0
  End If
End Sub

Sub Bumper2_timer()
  BumperRing2.Z = BumperRing2.Z + (5 * dirRing2)
  If BumperRing2.Z <= -40 Then dirRing2 = 1
  If BumperRing2.Z >= 0 Then
    dirRing2 = -1
    BumperRing2.Z = 0
    Me.TimerEnabled = 0
  End If
End Sub

Sub Bumper3_timer()
  BumperRing3.Z = BumperRing3.Z + (5 * dirRing3)
  If BumperRing3.Z <= -40 Then dirRing3 = 1
  If BumperRing3.Z >= 0 Then
    dirRing3 = -1
    BumperRing3.Z = 0
    Me.TimerEnabled = 0
  End If
End Sub


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


'Rollovers
Sub sw23_Hit:Controller.Switch(23) = 1:End Sub
Sub sw23_UnHit:Controller.Switch(23) = 0: End Sub

'Lower Lanes
Sub sw24_Hit:Controller.Switch(24) = 1:End Sub
Sub sw24_UnHit:Controller.Switch(24) = 0:End Sub
Sub sw25_Hit:Controller.Switch(25) = 1:leftInlaneSpeedLimit:End Sub
Sub sw25_UnHit:Controller.Switch(25) = 0:End Sub
Sub sw28_Hit:Controller.Switch(28) = 1:rightInlaneSpeedLimit:End Sub
Sub sw28_UnHit:Controller.Switch(28) = 0:End Sub
Sub sw29_Hit:Controller.Switch(29) = 1:End Sub
Sub sw29_UnHit:Controller.Switch(29) = 0:End Sub

'Upper Lanes
Sub sw8_Hit:Controller.Switch(8) = 1
End Sub
Sub sw8_UnHit:Controller.Switch(8) = 0:End Sub
Sub sw33_Hit:Controller.Switch(33) = 1:End Sub
Sub sw33_UnHit:Controller.Switch(33) = 0:End Sub
Sub sw34_Hit:Controller.Switch(34) = 1:End Sub
Sub sw34_UnHit:Controller.Switch(34) = 0:End Sub
Sub sw35_Hit:Controller.Switch(35) = 1:End Sub
Sub sw35_UnHit:Controller.Switch(35) = 0:End Sub

'Right
Sub sw37_Hit:Controller.Switch(37) = 1:End Sub
Sub sw37_UnHit:Controller.Switch(37) = 0:End Sub
Sub sw38_Hit:Controller.Switch(38) = 1
End Sub
Sub sw38_UnHit:Controller.Switch(38) = 0:End Sub

'Right Under Flipper
Sub sw46_Hit:Controller.Switch(46) = 1:End Sub
Sub sw46_UnHit:Controller.Switch(46) = 0:End Sub

'Spinner
Sub sw7_Spin:vpmTimer.PulseSw 7:SoundSpinner sw7: End Sub

'Right Ramp
Sub sw44_Hit:Controller.Switch(44) = 1:End Sub

Sub sw44_UnHit:Controller.Switch(44) = 0:End Sub

Sub sw45_Hit:Controller.Switch(45) = 1
End Sub
Sub sw45_UnHit:Controller.Switch(45) = 0:End Sub

'Left Ramp
Sub sw47_Hit:Controller.Switch(47) = 1:End Sub
Sub sw47_UnHit:Controller.Switch(47) = 0:End Sub

Sub sw48_Hit:Controller.Switch(48) = 1
End Sub

Sub sw48_UnHit:Controller.Switch(48) = 0:End Sub

'Venom da rampa venom
Sub sw43_Hit()
  Controller.Switch(43)=1
  FotocellulaVenom.Visible=0
End Sub

Sub sw43_UnHit()
  Controller.Switch(43)=0
  TimerFotocellula3.enabled=1
End Sub

Sub TimerFotocellula3_Timer()
  FotocellulaVenom.Visible=1
  Me.Enabled=0
End Sub

'Doc Ock
Sub S63_Hit:Controller.Switch(63)=1:End Sub

Sub S63_unHit:Controller.Switch(63)=0:End Sub

'Sandman
Sub s42_Hit:Controller.Switch(42)=1:End Sub
Sub s42_unHit:Controller.Switch(42)=0:End Sub
'Lock Opto
Sub sw6_Hit:vpmTimer.PulseSw 6
  sw6p.TransX = -5:sw6.TimerEnabled = 1:End Sub

Sub sw6_Timer:sw6p.TransX = 0: Me.TimerEnabled = 0: End Sub

Dim SandmanPronto

Sub SandmanOK_Hit
  SandmanPronto=1
  If ActiveBall.velY > 2  Then    'ball is going up
  sw59.enabled=1
  Else
  sw59.enabled=0
  End If
End Sub

Sub OctopusOK_Hit
  If ActiveBall.velY > 2  Then    'ball is going up
  sw36.enabled=1
  Else
  sw36.enabled=0
  End If
End Sub

'Sandman Optos
Dim zMultiplier

Sub sw9_Hit
  Select Case Int(Rnd * 4) + 1
        Case 1: zMultiplier = 2.55
        Case 2: zMultiplier = 2.2
        Case 3: zMultiplier = 1.9
        Case 4: zMultiplier = 1.7
    End Select
    activeball.velz = activeball.velz*zMultiplier
  If SandmanPronto=0 Then vpmTimer.PulseSw 9
  Bank.TransX = 5
  BankColpito.Enabled = 1
End Sub

Sub sw10_Hit
  Select Case Int(Rnd * 4) + 1
        Case 1: zMultiplier = 2.55
        Case 2: zMultiplier = 2.2
        Case 3: zMultiplier = 1.9
        Case 4: zMultiplier = 1.7
    End Select
    activeball.velz = activeball.velz*zMultiplier
  If SandmanPronto=0 Then vpmTimer.PulseSw 10
  Bank.TransX = 5
  BankColpito.Enabled = 1
End Sub

Sub sw11_Hit
  Select Case Int(Rnd * 4) + 1
        Case 1: zMultiplier = 2.55
        Case 2: zMultiplier = 2.2
        Case 3: zMultiplier = 1.9
        Case 4: zMultiplier = 1.7
    End Select
    activeball.velz = activeball.velz*zMultiplier
  If SandmanPronto=0 Then vpmTimer.PulseSw 11
  Bank.TransX = 5
  BankColpito.Enabled = 1
End Sub

Sub BankColpito_Timer()
  Bank.TransX = -0
  Me.Enabled = 0
End Sub

Sub sw12_Hit:vpmTimer.PulseSw 12
  sw12p.TransX = -5:sw12.TimerEnabled = 1:End Sub

Sub sw12_Timer:sw12p.TransX = 0:Me.TimerEnabled = 0: End Sub

Sub sw13_Hit:vpmTimer.PulseSw 13
  sw13p.TransX = -5:sw13.TimerEnabled = 1:End Sub

Sub sw13_Timer:sw13p.TransX = 0:Me.TimerEnabled = 0:End Sub

'Green Goblin Optos
Sub sw1_Hit:vpmTimer.PulseSw 1
  sw1p.TransX = -5: ModeTimer.Enabled = 1 : sw1.TimerEnabled = 1:End Sub

Sub sw1_Timer:sw1p.TransX = 0:Me.TimerEnabled = 0:End Sub

Sub sw2_Hit:vpmTimer.PulseSw 2
  sw2p.TransX = -5: sw2.TimerEnabled = 1:End Sub

Sub sw2_Timer:sw2p.TransX = 0:Me.TimerEnabled = 0:End Sub

Sub sw3_Hit:vpmTimer.PulseSw 3
  sw3p.TransX = -5: sw3.TimerEnabled = 1:End Sub

Sub sw3_Timer:sw3p.TransX = 0:Me.TimerEnabled = 0:End Sub

Sub sw4_Hit:vpmTimer.PulseSw 4
  sw4p.TransX = -5: sw4.TimerEnabled = 1:End Sub

Sub sw4_Timer:sw4p.TransX = 0:Me.TimerEnabled = 0:End Sub

Sub sw5_Hit:vpmTimer.PulseSw 5
  sw5p.TransX = -5: sw5.TimerEnabled = 1:End Sub

Sub sw5_Timer:sw5p.TransX = 0:Me.TimerEnabled = 0:End Sub

'Right 3Bank Optos
Sub sw39_Hit:vpmTimer.PulseSw 39
  sw39p.TransX = -5: sw39.TimerEnabled = 1:End Sub

Sub sw39_Timer:sw39p.TransX = 0:sw39.TimerEnabled = 0:End Sub

Sub sw40_Hit:vpmTimer.PulseSw 40
  sw40p.TransX = -5: sw40.TimerEnabled = 1:End Sub

Sub sw40_Timer:sw40p.TransX = 0:sw40.TimerEnabled = 0:End Sub

Sub sw41_Hit:vpmTimer.PulseSw 41
  sw41p.TransX = -5: sw41.TimerEnabled = 1:End Sub

Sub sw41_Timer:sw41p.TransX = 0:sw41.TimerEnabled = 0:End Sub

'Sub PallaBloccata_Hit: vpmTimer.PulseSw 9:vpmTimer.PulseSw 10:vpmTimer.PulseSw 11:End Sub

'Switch 14
Sub RPostColl21_Hit():vpmTimer.PulseSw 14:End Sub

'Sandman VUK

Sub MuroSandman_Hit()
  Playsound "Parete" 'TODO 'Walls?
End Sub

Sub sw59Abilitato_Hit()
  If PallaBucoSandman=0 Then
      sw59.Enabled=1
  End If
End Sub

Sub sw59Abilitato_UnHit()
  If PallaBucoSandman=0 Then
      sw59.Enabled=1
  End If
End Sub

Sub sw59_UnHit()  'Uscita buco sandman
  SandmanPronto=0
  PallaBucoSandman=0
  SoundSaucerKick 1, sw59
  sw59.Enabled=0
End Sub

Sub sw59_Hit()  'Entrata buco sandman
    SoundSaucerLock
  bsSandman.AddBall Me
  PallaBucoSandman=1
  ModeTimer.Enabled = 1
End Sub

'DocOck VUK

Dim PallaBucoOctopus, PallaBucoSandman

Sub MuroOctopus_Hit()
  Playsound "Parete" ' TODO
End Sub

Sub sw36Abilitato_UnHit()
  sw36.Enabled=1
End Sub

Sub sw36Abilitato_Hit()
  SoundSaucerLock
End Sub

Sub sw36_UnHit()  'Buco doc uscita
  PallaBucoOctopus=0
    SoundSaucerKick 1, sw36
  sw36.Enabled=0
End Sub

Sub sw36_Hit()  'Buco doc entrata
  bsDocOck.AddBall Me
  PallaBucoOctopus=1
 End Sub

'******************************************************
'       NFOZZY'S FLIPPERS
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

Sub SolURFlipper(Enabled)
    If Enabled Then
        RightFlipper2.RotateToEnd

        If RightFlipper2.currentangle < RightFlipper2.endangle + ReflipAngle Then
            RandomSoundReflipUpLeft RightFlipper2
        Else
            SoundFlipperUpAttackLeft RightFlipper2
            RandomSoundFlipperUpLeft RightFlipper2
        End If
    Else
        RightFlipper2.RotateToStart
        If RightFlipper2.currentangle < RightFlipper2.startAngle - 5 Then
            RandomSoundFlipperDownLeft RightFlipper2
        End If
        FlipperLeftHitParm = FlipperUpSoundLevel
    End If
End Sub



'******************************************************
'       FLIPPER AND RUBBER CORRECTION
'******************************************************

Dim LF : Set LF = New FlipperPolarity
Dim RF : Set RF = New FlipperPolarity

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

Dim LFPress, RFPress, URFPress, LFCount, RFCount
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




'*********************************
' GESTIONE MOVIMENTAZIONE GOBLIN
'*********************************


Dim GoblinPos

Sub ShakeGoblin(aValue)
  'debug.print "shakevalue:" & aValue
  GoblinPos = 8
  if aValue = True then
    GoblinShakeTimer.Enabled = 1
  end If
End Sub

Sub GoblinShakeTimer_Timer
  Playsound "shake" 'TODO
'    Goblin.TransZ = GoblinPos
' GoblinBracket.TransZ = GoblinPos
  Dim BP: For Each BP in BP_goblin: BP.TransZ = GoblinPos: Next
    If GoblinPos = 0 Then GoblinShakeTimer.Enabled = 0:Exit Sub
    If GoblinPos < 0 Then
        GoblinPos = ABS(GoblinPos) - 1
    Else
        GoblinPos = - GoblinPos + 1
    End If
End Sub


Sub ModeGiColor(aNr, aValue)
  if aValue = 1 or aValue = 0 Then
    ModeTimer.Enabled = 1
  end If
end Sub

Sub ModeTimer_Timer()

  dim ar,ag,ab,sum,ii:ar=0:ag=0:ab=0

  ar = round1(Lampz.state(74)) + round1(Lampz.state(76))
  ag = round1(Lampz.state(74)) + round1(Lampz.state(77))
  ab = round1(Lampz.state(75)) + round1(Lampz.state(76))

  sum = ar+ag+ab
  'debug.print "red: " & ar & " green: " & ag & " blue: " & ab
  if sum > 0 and sum < 6 then
    GIColorRed=ar/sum*180+20
    GIColorGreen=ag/sum*180+20
    GIColorBlue=ab/sum*180+20
  else
    GIColorRed       =  GIColorRedOrig
    GIColorGreen     =  GIColorGreenOrig
    GIColorBlue      =  GIColorBlueOrig
  end if
  SetGIColor
  Me.Enabled=0
End Sub

Function round1(a)
  if a >= 0.5 Then
    round1 = 1
  Else
    round1 = 0
  End If
End Function


'*********************************
' GESTIONE MOVIMENTAZIONE OCTOPUS
'*********************************

Dim OctopusDir, OctopusPos

Sub sw63_Hit()
  PlaysoundAtVol "bersaglio", BM_Octopus, 1
' Octopus.TransZ = 5
' BracketOctopus.TransZ = 5
  Dim BP: For Each BP in BP_Octopus: BP.TransY = -5: Next
  OctopusColpito.Enabled = 1
End Sub

Sub Fotocellula1Disattiva_Hit()
  FotocellulaOctopus.Visible=0
End Sub

Sub Fotocellula1Disattiva_UnHit()
  TimerFotocellula1.enabled=1
End Sub

Sub TimerFotocellula1_Timer()
  FotocellulaOctopus.Visible=1
  Me.Enabled=0
End Sub

Sub Fotocellula2Disattiva_Hit()
  FotocellulaSandmanDx.Visible=0
  FotocellulaSandmanSx.Visible=0
End Sub

Sub Fotocellula2Disattiva_UnHit()
  TimerFotocellula2.enabled=1
End Sub

Sub TimerFotocellula2_Timer()
  FotocellulaSandmanDx.Visible=1
  FotocellulaSandmanSx.Visible=1
  Me.Enabled=0
End Sub




Sub OctopusColpito_Timer()
' Octopus.TransZ = 0
' BracketOctopus.TransZ = 0
  Dim BP: For Each BP in BP_Octopus: BP.TransY = 0: Next
  Me.Enabled = 0
End Sub

Sub AlzaOctopus
  OctopusDir = -1 ' removing 1 will make Oct go up
  OctopusTimer.Enabled = 1
End Sub

Sub AbbassaOctopus
  OctopusDir = 1 ' adding 1 will make Oct to go down by one step
  OctopusTimer.Enabled = 1
End Sub


Sub OctopusTimer_Timer
  PlaysoundAtVol "Motor", BM_Octopus, 1
' Octopus.TransY = -OctopusPos
' BracketOctopus.TransY = -OctopusPos
  Dim BP: For Each BP in BP_Octopus: BP.TransZ = -OctopusPos: Next
  OctopusPos = OctopusPos + OctopusDir
  If OctopusPos < 0 Then OctopusPos=0: sw63.IsDropped=1: Me.Enabled = 0
  If OctopusPos > 50 Then OctopusPos = 50: Me.Enabled = 0
  ModeTimer.Enabled=1
End Sub


'*********************************
' GESTIONE MOVIMENTAZIONE SANDMAN
'*********************************

Dim SandmanDir, SandmanPos, SandmanAlto

Sub sw42_Hit()
  PlaysoundAtVol "bersaglio", BM_Sandman, 1
' Sandman.TransZ = 5
' BracketSandman.TransZ = 5
  Dim BP: For Each BP in BP_Sandman: BP.TransY = -5: Next
  SandmanColpito.Enabled = 1
End Sub

Sub SandmanColpito_Timer()  'hit
' Sandman.TransZ = 0
' BracketSandman.TransZ = 0
  Dim BP: For Each BP in BP_Sandman: BP.TransY = 0: Next
  Me.Enabled = 0
End Sub

Sub AlzaSandman 'rise sandman
  SandmanDir = 1 ' removing 1 will make Oct go up
  SandmanTimer.Enabled = 1
  SandmanAlto=1
End Sub

Sub AbbassaSandman  'drop sandman
  SandmanDir = -1 ' adding 1 will make Oct to go down by one step
  SandmanTimer.Enabled = 1
  SandmanAlto=0
End Sub


Sub SandmanTimer_Timer
  PlaysoundAtVol "Motor", BM_Sandman, 1
' Sandman.TransY = SandmanPos
' BracketSandman.TransY = SandmanPos
  Dim BP: For Each BP in BP_Sandman: BP.TransZ = SandmanPos: Next
  SandmanPos = SandmanPos + SandmanDir
  If SandmanPos < 0 Then SandmanPos=0: Me.Enabled = 0
  If SandmanPos > 50 Then SandmanPos = 50: sw42.IsDropped=1: Me.Enabled = 0
  ModeTimer.Enabled=1
End Sub


'***********************************
' GESTIONE MOVIMENTAZIONE MURO BABK
'***********************************

Dim BankDir, BankPos

Sub AlzaBank
  BankDir = -1
  BankTimer.Enabled = 1
End Sub

Sub AbbassaBank
  BankDir = 1
  BankTimer.Enabled = 1
End Sub

Sub BankTimer_Timer
  Bank.TransY = -BankPos
  BankPos = BankPos + BankDir
  If BankPos < 0 Then BankPos = 0: Me.Enabled = 0
  If BankPos > 52 Then BankPos = 52: ParetiBankGiu: Me.Enabled = 0
  ModeTimer.Enabled=1
End Sub

Dim BankAlto

Sub PallaBloccata_Hit()
  If BankAlto=1 AND Attendi=0 Then
  Controller.Switch(50)=0
  Controller.Switch(49)=1
  AbbassaBank
  ParetiBankGiu
  SbloccaPallaBank.Enabled= True
  End If
End Sub

Sub SbloccaPallaBank_Timer()
  If Attendi=0 Then
  Controller.Switch(49)=0
  Controller.Switch(50)=1
  AlzaBank
  ParetiBankSu
  Me.Enabled=0
  End If
End Sub

Sub ParetiBankSu
  sw9.IsDropped=0
  sw10.IsDropped=0
  sw11.IsDropped=0
  sandblock.IsDropped=0
  BankAlto=1
End Sub

Sub ParetiBankGiu
  sw9.IsDropped=1
  sw10.IsDropped=1
  sw11.IsDropped=1
  sandblock.IsDropped=1
  BankAlto=0
End Sub


' Sub ShooterEnd_UnHit():If activeball.z > 30  Then vpmTimer.AddTimer 150, "BallHitSound":End If:End Sub

' Sub BallHitSound(dummy):PlaySound "ballhit":End Sub

Sub RampEnd1_Hit()
    'AccendiVenom
    ActiveBall.VelZ = -2
    ActiveBall.VelY = 0
    ActiveBall.VelX = 0
    RandomSoundRampStop RampEnd1
    ModeTimer.Enabled=1
 End Sub

 Sub RampEnd1_UnHit()
    RandomSoundBallBouncePlayfieldHard ActiveBall
 End Sub

Sub RampEnd2_Hit()
    ActiveBall.VelZ = -2
    ActiveBall.VelY = 0
    ActiveBall.VelX = 0
    RandomSoundRampStop RampEnd2
    ModeTimer.Enabled=1
End Sub

Sub RampEnd2_UnHit()
    RandomSoundDelayedBallDropOnPlayfield ActiveBall
End Sub

Sub RHP1_Hit()
    WireRampOn True
End Sub

Sub RHP2_Hit
  WireRampOff
End Sub

Sub RHP2_UnHit()
    WireRampOn False
End Sub

Sub LHP1_Hit()
    WireRampOn True
End Sub

Sub LHP2_Hit()
    WireRampOn True
End Sub

Sub LHP3_Hit() 'Venom da rampa sinistra
  WireRampOff
    WireRampOn False
End Sub

Sub VenomOK_UnHit()
'    Debug.Print("Hit ramp trigger")
  WireRampOff
    WireRampOn False
End Sub

Sub DocVUKExit_Hit()
    WireRampOn False
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

Sub RightFlipper2_Collide(parm)
  RandomSoundRubberFlipper(parm)
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

Dim AutoPlungerSoundLevel
AutoPlungerSoundLevel = 1                       'volume level; range [0, 1]

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

Sub Apron1_Hit (idx)
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

'////////////////////////////  AUTO-PLUNGER SOUNDS  /////////////////////////////
Sub RandomSoundAutoPlunger()
    PlaySoundAtLevelStatic SoundFX("SY_TNA_REV01_Auto_Launch_Coil_" & Int(Rnd*5)+1, DOFContactors), AutoPlungerSoundLevel, swplunger
End Sub

Sub RandomSoundAutoPlungerDOF(AutoPlunger, DOFevent, DOFstate)
    PlaySoundAtLevelStatic SoundFXDOF("SY_TNA_REV01_Auto_Launch_Coil_" & Int(Rnd*5)+1, DOFevent, DOFstate, DOFContactors), AutoPlungerSoundLevel, AutoPlunger
End Sub

'/////////////////////////////////////////////////////////////////
'         End Mechanical Sounds
'/////////////////////////////////////////////////////////////////

'******************************************************
'****  FLEEP MECHANICAL SOUNDS
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
Dim RampBalls(6,2)
'x,0 = ball x,1 = ID, 2 = Protection against ending early (minimum amount of updates)

'0,0 is boolean on/off, 0,1 unused for now
RampBalls(0,0) = False

' RampType
' Setup: Set this array to the number Total number of balls that can be tracked at one time + 1.  5 ball multiball then set value to 6
' Description: Array type indexed on BallId and a values used to deterimine what type of ramp the ball is on: False = Wire Ramp, True = Plastic Ramp
Dim RampType(6)

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

Sub RandomSoundRampStop(obj)
  Select Case Int(rnd*3)
    Case 0: PlaySoundAtVol "wireramp_stop", obj, 0.2*VolumeDial:PlaySoundAtLevelActiveBall ("Rubber_Strong_1"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
    Case 1: PlaySoundAtVol "wireramp_stop2", obj, 0.2*VolumeDial:PlaySoundAtLevelActiveBall ("Rubber_Strong_2"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
    Case 2: PlaySoundAtVol "wireramp_stop3", obj, 0.2*VolumeDial:PlaySoundAtLevelActiveBall ("Rubber_1_Hard"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
  End Select
End Sub

'******************************************************
'**** END RAMP ROLLING SFX
'******************************************************



'******************************************************
'   BALL ROLLING AND DROP SOUNDS
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

Sub RollingSound()
  Dim b

  ' stop the sound of deleted balls
  For b = UBound(gBOT) + 1 to tnob
    rolling(b) = False
    StopSound("BallRoll_" & b)
  Next

  ' exit the sub if no balls on the table
  If UBound(gBOT) = -1 Then Exit Sub

  ' play the rolling sound for each ball

  For b = 0 to UBound(gBOT)
    If BallVel(gBOT(b)) > 1 AND gBOT(b).z < 30 Then
      rolling(b) = True
      PlaySound ("BallRoll_" & b), -1, VolPlayfieldRoll(gBOT(b)) * 1.1 * VolumeDial, AudioPan(gBOT(b)), 0, PitchPlayfieldRoll(gBOT(b)), 1, 0, AudioFade(gBOT(b))

    Else
      If rolling(b) = True Then
        StopSound("BallRoll_" & b)
        rolling(b) = False
      End If
    End If

    '***Ball Drop Sounds***
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
'     RUBBER CORRECTION
'******************************************************


Sub dPosts_Hit(idx)
  RubbersD.dampen ActiveBall
  TargetBouncer ActiveBall, 1
End Sub

Sub dSleeves_Hit(idx)
  SleevesD.Dampen ActiveBall
  TargetBouncer ActiveBall, 0.7
End Sub

Sub dPostsNoTB_Hit(idx)
    RubbersD.dampen ActiveBall
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


'******************************************************
'   ZBOU: VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************

Const TargetBouncerEnabled = 1    '0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 1  'Level of bounces. Recommmended value of 0.7

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


'*******************************************************
' End nFozzy Dampening'
'******************************************************




'******************************************************
' Flasher domes
'******************************************************
Sub FlashSol29(pwm) ' blue
  Flasherlight1.state = pwm
End Sub

Sub FlashSol30(pwm) ' top red
  Flasherlight2.state = pwm
End Sub

Sub FlashSol27(pwm) ' top red
  Flasherlight3.state = pwm
End Sub

Sub FlashSol21(pwm) ' top red
  Flasherlight4.state = pwm
End Sub



'***************************************
'***Begin nFozzy lamp handling***
'***************************************

Dim NullFader : set NullFader = new NullFadingObject
Dim Lampz : Set Lampz = New VPMLampUpdater

InitLampsNF              ' Setup lamp assignments

Sub LampTimer()
  Dim x, chglamp
  chglamp = Controller.ChangedLamps
  If Not IsEmpty(chglamp) Then
    For x = 0 To UBound(chglamp) 'nmbr = chglamp(x, 0), state = chglamp(x, 1)
      Lampz.state(chglamp(x, 0)) = chglamp(x, 1)/255.0
    Next
  End If
End Sub


Sub UpdateLightMap(lightmap, intensity, ByVal aLvl)
   lightmap.Opacity = aLvl * intensity
End Sub

Sub DisableLighting(pri, DLintensity, ByVal aLvl) 'cp's script  DLintensity = disabled lighting intesity
  pri.blenddisablelighting = aLvl * DLintensity
End Sub

Sub SetModLamp(id, val)
  Lampz.state(id) = val
End Sub


Sub InitLampsNF()

  Lampz.MassAssign(3) = l3
  Lampz.Callback(3) = "DisableLighting p3on, 75,"


  Lampz.MassAssign(4) = l4
  Lampz.Callback(4) = "DisableLighting p4on, 275,"


  Lampz.MassAssign(5) = l5
  Lampz.Callback(5) = "DisableLighting p5on, 275,"


  Lampz.MassAssign(6) = l6
  Lampz.Callback(6) = "DisableLighting p6on, 275,"


  Lampz.MassAssign(7) = l7
  Lampz.Callback(7) = "DisableLighting p7on, 275,"


  Lampz.MassAssign(8) = l8
  Lampz.Callback(8) = "DisableLighting p8on, 275,"


  Lampz.MassAssign(9) = l9
  Lampz.Callback(9) = "DisableLighting p9on, 275,"


  Lampz.MassAssign(10) = l10
  Lampz.Callback(10) = "DisableLighting p10on, 275,"


  Lampz.MassAssign(11) = l11
  Lampz.Callback(11) = "DisableLighting p11on, 275,"


  Lampz.MassAssign(12) = l12
  Lampz.Callback(12) = "DisableLighting p12on, 975,"


  Lampz.MassAssign(13) = l13
  Lampz.Callback(13) = "DisableLighting p13on, 275,"


  Lampz.MassAssign(14) = l14
  Lampz.Callback(14) = "DisableLighting p14on, 975,"


  Lampz.MassAssign(15) = l15
  Lampz.Callback(15) = "DisableLighting p15on, 75,"


  Lampz.MassAssign(16) = l16
  Lampz.Callback(16) = "DisableLighting p16on, 75,"


  Lampz.MassAssign(17) = l17
  Lampz.Callback(17) = "DisableLighting p17on, 75,"
  Lampz.MassAssign(17) = l17r

  Lampz.MassAssign(18) = l18
  Lampz.Callback(18) = "DisableLighting p18on, 75,"
  Lampz.MassAssign(18) = l18r

  Lampz.MassAssign(19) = l19
  Lampz.Callback(19) = "DisableLighting p19on, 75,"
  Lampz.MassAssign(19) = l19r

  Lampz.MassAssign(20) = l20
  Lampz.Callback(20) = "DisableLighting p20on, 75,"
  Lampz.MassAssign(20) = l20r

  Lampz.MassAssign(21) = l21
  Lampz.Callback(21) = "DisableLighting p21on, 75,"
  Lampz.MassAssign(21) = l21r

  Lampz.MassAssign(22) = l22
  Lampz.Callback(22) = "DisableLighting p22on, 75,"


  Lampz.MassAssign(23) = l23
  Lampz.Callback(23) = "DisableLighting p23on, 75,"

  Lampz.MassAssign(24) = l24
  Lampz.Callback(24) = "DisableLighting p24on, 75,"


  Lampz.MassAssign(25) = l25
  Lampz.Callback(25) = "DisableLighting p25on, 75,"


  Lampz.MassAssign(26) = l26
  Lampz.Callback(26) = "DisableLighting p26on, 75,"


  Lampz.MassAssign(27) = l27
  Lampz.Callback(27) = "DisableLighting p27on, 75,"


  Lampz.MassAssign(28) = l28
  Lampz.Callback(28) = "DisableLighting p28on, 75,"
  Lampz.MassAssign(28) = l28r

  Lampz.MassAssign(29) = l29
  Lampz.Callback(29) = "DisableLighting p29on, 75,"


  Lampz.MassAssign(30) = l30
  Lampz.Callback(30) = "DisableLighting p30on, 75,"


  Lampz.MassAssign(31) = l31
  Lampz.Callback(31) = "DisableLighting p31on, 75,"


  Lampz.MassAssign(32) = l32
  Lampz.Callback(32) = "DisableLighting p32on, 75,"


  Lampz.MassAssign(33) = l33
  Lampz.Callback(33) = "DisableLighting p33on, 75,"


  Lampz.MassAssign(34) = l34
  Lampz.Callback(34) = "DisableLighting p34on, 75,"


  Lampz.MassAssign(35) = l35
  Lampz.Callback(35) = "DisableLighting p35on, 75,"

  Lampz.MassAssign(36) = l36
  Lampz.Callback(36) = "DisableLighting p36on, 75,"

  Lampz.MassAssign(37) = l37
  Lampz.Callback(37) = "DisableLighting p37on, 75,"


  Lampz.MassAssign(38) = l38
  Lampz.Callback(38) = "DisableLighting p38on, 75,"


  Lampz.MassAssign(39) = l39
  Lampz.Callback(39) = "DisableLighting p39on, 75,"


  Lampz.MassAssign(40) = l40
  Lampz.Callback(40) = "DisableLighting p40on, 75,"


  Lampz.MassAssign(41) = l41
  Lampz.Callback(41) = "DisableLighting p41on, 75,"
  Lampz.MassAssign(41) = l41r

  Lampz.MassAssign(42) = l42
  Lampz.Callback(42) = "DisableLighting p42on, 75,"


  Lampz.MassAssign(43) = l43
  Lampz.Callback(43) = "DisableLighting p43on, 75,"


  Lampz.MassAssign(44) = l44
  Lampz.Callback(44) = "DisableLighting p44on, 75,"
  Lampz.MassAssign(44) = l44r

  Lampz.MassAssign(45) = l45
  Lampz.Callback(45) = "DisableLighting p45on, 75,"


  Lampz.MassAssign(46) = l46
  Lampz.Callback(46) = "DisableLighting p46on, 75,"


  Lampz.MassAssign(47) = l47
  Lampz.Callback(47) = "DisableLighting p47on, 75,"


  Lampz.MassAssign(48) = l48
  Lampz.Callback(48) = "DisableLighting p48on, 75,"


  Lampz.MassAssign(49) = l49
  Lampz.Callback(49) = "DisableLighting p49on, 75,"


  Lampz.MassAssign(50) = l50
  Lampz.Callback(50) = "DisableLighting p50on, 75,"


  Lampz.MassAssign(51) = l51
  Lampz.Callback(51) = "DisableLighting p51on, 75,"


  Lampz.MassAssign(52) = l52
  Lampz.Callback(52) = "DisableLighting p52on, 75,"
  Lampz.MassAssign(52) = l52r

  Lampz.MassAssign(53) = l53
  Lampz.Callback(53) = "DisableLighting p53on, 75,"
  Lampz.MassAssign(53) = l53r


  Lampz.MassAssign(54) = l54
  Lampz.Callback(54) = "DisableLighting p54on, 75, "
  Lampz.MassAssign(54) = l54r

  Lampz.MassAssign(57) = l57
  Lampz.Callback(57) = "DisableLighting p57on, 75,"


  Lampz.MassAssign(58) = l58
  Lampz.Callback(58) = "DisableLighting p58on, 75,"


  Lampz.MassAssign(59) = l59
  Lampz.Callback(59) = "DisableLighting p59on, 75,"

  Lampz.MassAssign(60) = L60
  Lampz.Callback(60) = "UpdateLightMap LM_IN_L60_Layer0, 100.0, "
  Lampz.Callback(60) = "UpdateLightMap LM_IN_L60_Layer1, 100.0, "
  Lampz.Callback(60) = "UpdateLightMap LM_IN_L60_Layer3, 100.0, "
  Lampz.Callback(60) = "UpdateLightMap LM_IN_L60_Parts, 100.0, "
  Lampz.Callback(60) = "UpdateLightMap LM_IN_L60_Sandman, 100.0, "

  Lampz.MassAssign(61) = L61
  Lampz.Callback(61) = "UpdateLightMap LM_IN_L61_Layer0, 100.0, "
  Lampz.Callback(61) = "UpdateLightMap LM_IN_L61_Layer1, 100.0, "
  Lampz.Callback(61) = "UpdateLightMap LM_IN_L61_Parts, 100.0, "

  Lampz.MassAssign(62) = L62
  Lampz.Callback(62) = "UpdateLightMap LM_IN_L62_Layer0, 100.0, "
  Lampz.Callback(62) = "UpdateLightMap LM_IN_L62_Layer1, 100.0, "
  Lampz.Callback(62) = "UpdateLightMap LM_IN_L62_Parts, 100.0, "
  Lampz.Callback(62) = "UpdateLightMap LM_IN_L62_Sandman, 100.0, "


  Lampz.MassAssign(63) = l63
  Lampz.Callback(63) = "DisableLighting p63on, 75,"
  Lampz.MassAssign(63) = l63r

  Lampz.MassAssign(64) = l64
  Lampz.Callback(64) = "DisableLighting p64on, 75,"

  Lampz.MassAssign(65) = l65
  Lampz.Callback(65) = "DisableLighting p65on, 75,"

  Lampz.MassAssign(66) = MissioneVenom3
  Lampz.Callback(66) = "UpdateLightMap LM_IN_MissioneVenom3_Layer0, 100.0, "
  Lampz.Callback(66) = "UpdateLightMap LM_IN_MissioneVenom3_Layer2, 100.0, "
  Lampz.Callback(66) = "UpdateLightMap LM_IN_MissioneVenom3_Layer3, 100.0, "
  Lampz.Callback(66) = "UpdateLightMap LM_IN_MissioneVenom3_Parts, 100.0, "
  Lampz.Callback(66) = "UpdateLightMap LM_IN_MissioneVenom3_Venom, 100.0, "

  Lampz.MassAssign(67) = MissioneVenom2
  Lampz.Callback(67) = "UpdateLightMap LM_IN_MissioneVenom2_Layer0, 100.0, "
  Lampz.Callback(67) = "UpdateLightMap LM_IN_MissioneVenom2_Layer2, 100.0, "
  Lampz.Callback(67) = "UpdateLightMap LM_IN_MissioneVenom2_Layer1, 100.0, "
  Lampz.Callback(67) = "UpdateLightMap LM_IN_MissioneVenom2_Layer3, 100.0, "
  Lampz.Callback(67) = "UpdateLightMap LM_IN_MissioneVenom2_Parts, 100.0, "

  Lampz.MassAssign(68) = MissioneVenom1
  Lampz.Callback(68) = "UpdateLightMap LM_IN_MissioneVenom1_Layer0, 100.0, "
  Lampz.Callback(68) = "UpdateLightMap LM_IN_MissioneVenom1_Layer1, 100.0, "
  Lampz.Callback(68) = "UpdateLightMap LM_IN_MissioneVenom1_Layer2, 100.0, "
  Lampz.Callback(68) = "UpdateLightMap LM_IN_MissioneVenom1_Layer3, 100.0, "
  Lampz.Callback(68) = "UpdateLightMap LM_IN_MissioneVenom1_Parts, 100.0, "


  Lampz.MassAssign(69) = MissioneSandman3
  Lampz.Callback(69) = "UpdateLightMap LM_IN_MissioneSandman3_Layer1, 100.0, "
  Lampz.Callback(69) = "UpdateLightMap LM_IN_MissioneSandman3_Parts, 100.0, "

  Lampz.MassAssign(70) = MissioneSandman2
  Lampz.Callback(70) = "UpdateLightMap LM_IN_MissioneSandman2_Layer0, 100.0, "
  Lampz.Callback(70) = "UpdateLightMap LM_IN_MissioneSandman2_Layer1, 100.0, "
  Lampz.Callback(70) = "UpdateLightMap LM_IN_MissioneSandman2_Layer2, 100.0, "
  Lampz.Callback(70) = "UpdateLightMap LM_IN_MissioneSandman2_Layer3, 100.0, "
  Lampz.Callback(70) = "UpdateLightMap LM_IN_MissioneSandman2_Parts, 100.0, "

  Lampz.MassAssign(71) = MissioneSandman1
  Lampz.Callback(71) = "UpdateLightMap LM_IN_MissioneSandman1_Layer0, 100.0, "
  Lampz.Callback(71) = "UpdateLightMap LM_IN_MissioneSandman1_Layer1, 100.0, "
  Lampz.Callback(71) = "UpdateLightMap LM_IN_MissioneSandman1_Layer2, 100.0, "
  Lampz.Callback(71) = "UpdateLightMap LM_IN_MissioneSandman1_Layer3, 100.0, "
  Lampz.Callback(71) = "UpdateLightMap LM_IN_MissioneSandman1_Parts, 100.0, "


  Lampz.MassAssign(72) = l72
  Lampz.Callback(72) = "DisableLighting p72on, 75,"
  Lampz.MassAssign(72) = l72r

  'Lampz.MassAssign(74) = spotSandman
  Lampz.Callback(74) = "ModeGiColor 74,"
  'Lampz.Callback(74) = "DisableLighting Sandman, 0.5,"

  'Lampz.MassAssign(75) = spotVenom
  Lampz.Callback(75) = "ModeGiColor 75,"
  'Lampz.Callback(75) = "DisableLighting Venom, 0.5,"

  'Lampz.MassAssign(76) = spotGoblin
  Lampz.Callback(76) = "ModeGiColor 76,"
  'Lampz.Callback(76) = "DisableLighting Goblin, 0.5,"

  'Lampz.MassAssign(77) = spotOcto
  Lampz.Callback(77) = "ModeGiColor 77,"
  'Lampz.Callback(77) = "DisableLighting Octopus, 1,"

  Lampz.MassAssign(78) = l78
  Lampz.Callback(78) = "DisableLighting p78on, 75,"

  Lampz.MassAssign(123)=f23
  Lampz.MassAssign(123)=F23a

  Lampz.MassAssign(125)=f25
  Lampz.MassAssign(125)=F25a

  Lampz.MassAssign(126)=f26

  Lampz.MassAssign(128)=FlasherGoblin

  Lampz.MassAssign(131)=FlasherBumpers

  dim ii
  For each ii in GI:ii.IntensityScale = 0.3:Next
  For each ii in GI_PF:ii.IntensityScale = 1:Next
  'For each ii in GI_PF:ii.Falloff = 400:Next


  'This just turns state of any lamps to 1
  Lampz.Init

  'Turn off all lamps on startup
  Dim x: For x = 0 to 150: Lampz.State(x) = 0: Next


End Sub


Function ColtoArray(aDict)  'converts a collection to an indexed array. Indexes will come out random probably.
  redim a(999)
  dim count : count = 0
  dim x  : for each x in aDict : set a(Count) = x : count = count + 1 : Next
  redim preserve a(count-1) : ColtoArray = a
End Function


'*********************************************************************************************************************************************************
'Begin lamp helper functions
'*********************************************************************************************************************************************************

'***************************************
'GI On/Off
'***************************************
Const MaxGILvl = 0.75

dim gilvl:gilvl = 0

Set GICallback = GetRef("GIUpdate")
Set GICallback2 = GetRef("GIUpdate2")

Sub GIUpdate(no, enabled)
  'debug.print "GIUpdate no="&no&" enabled="&enabled
End Sub

Sub GIUpdate2(no, level)
  'debug.print "GIUpdate2 no="&no&" level="&level
  Dim bulb
  For each bulb in GI: bulb.state = level/MaxGILvl: Next
  SetGIColor
  If level >= 0.5 And gilvl < 0.5 Then
    DOF 201, DOFOn
    PinCab_Backglass.image = "BackglassImageOn"
    LFLogo.blenddisablelighting=2
    RFLogo.blenddisablelighting=2
    flipperr1.blenddisablelighting=2
  ElseIf level <= 0.4 And gilvl > 0.4 Then
    DOF 201, DOFOff
    PinCab_Backglass.image = "BackglassImage"
    LFLogo.blenddisablelighting=1
    RFLogo.blenddisablelighting=1
    flipperr1.blenddisablelighting=1
  End if
  gilvl = level
End Sub


'***GI Color Mod***
Dim GIxx, ColorModRed, ColorModRedFull, ColorModGreen, ColorModGreenFull, ColorModBlue, ColorModBlueFull
Dim GIColorRed, GIColorGreen, GIColorBlue, GIColorFullRed, GIColorFullGreen, GIColorFullBlue, GIColor
Sub SetGIColor ()
' for each GIxx in GILighting
'   GIxx.Color = rgb(GIColorRed, GIColorGreen, GIColorBlue)
'   GIxx.ColorFull = rgb(GIColorFullRed, GIColorFullGreen, GIColorFullBlue)
' next
  const ColorMultiplier = 0.8
  const FullColorMultiplier = 0.95
' debug.print GIColorRed &" - "& GIColorGreen &" - "& GIColorBlue

  for each GIxx in GI
    GIxx.Color = rgb(GIColorRed*ColorMultiplier, GIColorGreen*ColorMultiplier, GIColorBlue*ColorMultiplier)
    GIxx.ColorFull = rgb(GIColorRed*FullColorMultiplier, GIColorGreen*FullColorMultiplier, GIColorBlue*FullColorMultiplier)
    GIColor = GIxx.ColorFull
  next
  FlasherRGBGI.Color = GIColor

  'Lightmap color changes
  Dim BL
  For Each BL in BL_GI: BL.color = GIColor: Next
  For Each BL in BL_GIS_LA008: BL.color = GIColor: Next
  For Each BL in BL_GIS_LA010: BL.color = GIColor: Next
  For Each BL in BL_GIS_LA012: BL.color = GIColor: Next
  For Each BL in BL_GIS_LA013: BL.color = GIColor: Next
  For Each BL in BL_GIS_LA019: BL.color = GIColor: Next
  For Each BL in BL_GIS_LA020: BL.color = GIColor: Next
  For Each BL in BL_GIS_LA021: BL.color = GIColor: Next
  For Each BL in BL_GIS_LA022: BL.color = GIColor: Next
  For Each BL in BL_GIS_LA023: BL.color = GIColor: Next
  For Each BL in BL_GIS_LA028: BL.color = GIColor: Next

End Sub


'*********************************************************************************************************************************************************
'End lamp helper functions
'*********************************************************************************************************************************************************


'====================
'Class jungle nf
'====================

'No-op object instead of adding more conditionals to the main loop
'It also prevents errors if empty lamp numbers are called, and it's only one object
'should be g2g?
Class NullFadingObject
  Public Property Let IntensityScale(input)

  End Property
End Class

Class VPMLampUpdater
  Public Name
  Public Obj(150), OnOff(150)
  Private UseCallback(150), cCallback(150)

  Sub Class_Initialize()
    Name = "VPMLampUpdater" 'NEEDS TO BE CHANGED IF THERE'S MULTIPLE OF THESE OBJECTS, OTHERWISE CALLBACKS WILL INTERFERE WITH EACH OTHER!!
    Dim x : For x = 0 to uBound(OnOff)
        OnOff(x) = 0
      Set Obj(x) = NullFader
    Next
  End Sub

  Public Property Let Callback(idx, String)
    UseCallBack(idx) = True
    'cCallback(idx) = String 'old execute method
    'New method: build wrapper subs using ExecuteGlobal, then call them
    cCallback(idx) = cCallback(idx) & "___" & String  'multiple strings dilineated by 3x _

    dim tmp : tmp = Split(cCallback(idx), "___")

    dim str, x : for x = 0 to uBound(tmp) 'build proc contents
      'If Not tmp(x)="" then str = str & "  " & tmp(x) & " aLVL" & "  '" & x & vbnewline  'more verbose
      If Not tmp(x)="" then str = str & tmp(x) & " aLVL:"
    Next

    dim out : out = "Sub " & name & idx & "(aLvl):" & str & "End Sub"
    'if idx = 132 then msgbox out 'debug
    ExecuteGlobal Out
  End Property

  'Mass assign, Builds arrays where necessary
  'Sub MassAssign(aIdx, aInput)
  Public Property Let MassAssign(aIdx, aInput)
    If typename(obj(aIdx)) = "NullFadingObject" Then 'if empty, use Set
      if IsArray(aInput) then
        obj(aIdx) = aInput
      Else
        Set obj(aIdx) = aInput
      end if
    Else
      Obj(aIdx) = AppendArray(obj(aIdx), aInput)
    end if
  End Property

  Sub SetLamp(aIdx, aOn) : state(aIdx) = aOn : End Sub

  Public Property Let state(ByVal x, input)
    Dim xx
    OnOff(x) = input
    If IsArray(obj(x)) Then
      For Each xx In obj(x)
        xx.IntensityScale = input
        'debug.print x&"  obj.Intensityscale = " & input
      Next
    Else
      obj(x).Intensityscale = input
      'debug.print "obj("&x&").Intensityscale = " & input
    End if
    'If UseCallBack(x) then execute cCallback(x) & " " & (Lvl(x)) 'Callback
    If UseCallBack(x) then Proc name & x,input
  End Property

  Public Property Get state(idx) : state = OnOff(idx) : end Property

  Public Sub TurnOnStates() 'If obj contains any light objects, set their states to 1 (Fading is our job!)
    dim idx : for idx = 0 to uBound(obj)
      if IsArray(obj(idx)) then
        dim x, tmp : tmp = obj(idx) 'set tmp to array in order to access it
        for x = 0 to uBound(tmp)
          if typename(tmp(x)) = "Light" then DisableState tmp(x)' : debugstr = debugstr & tmp(x).name & " state'd" & vbnewline
          tmp(x).intensityscale = 0.001 ' this can prevent init stuttering
        Next
      Else
        if typename(obj(idx)) = "Light" then DisableState obj(idx)' : debugstr = debugstr & obj(idx).name & " state'd (not array)" & vbnewline
        obj(idx).intensityscale = 0.001 ' this can prevent init stuttering
      end if
    Next
  End Sub
  Private Sub DisableState(ByRef aObj) : aObj.FadeSpeedUp = 1000 : aObj.State = 1 : End Sub 'turn state to 1

  Public Sub Init() 'Just runs TurnOnStates right now
    TurnOnStates
  End Sub

End Class


'Helper functions
Sub Proc(String, Callback)  'proc using a string and one argument
  'On Error Resume Next
  Dim p
  Set P = GetRef(String)
  P Callback
  If err.number = 13 Then  MsgBox "Proc error! No such procedure: " & vbNewLine & String
  If err.number = 424 Then MsgBox "Proc error! No such Object"
End Sub

Function AppendArray(ByVal aArray, aInput)  'append one value, object, or Array onto the End of a 1 dimensional array
  If IsArray(aInput) Then 'Input is an array...
    Dim tmp
    tmp = aArray
    If Not IsArray(aArray) Then 'if not array, create an array
      tmp = aInput
    Else          'Append existing array with aInput array
      ReDim Preserve tmp(UBound(aArray) + UBound(aInput) + 1) 'If existing array, increase bounds by uBound of incoming array
      Dim x
      For x = 0 To UBound(aInput)
        If IsObject(aInput(x)) Then
          Set tmp(x + UBound(aArray) + 1 ) = aInput(x)
        Else
          tmp(x + UBound(aArray) + 1 ) = aInput(x)
        End If
      Next
      AppendArray = tmp  'return new array
    End If
  Else 'Input is NOT an array...
    If Not IsArray(aArray) Then 'if not array, create an array
      aArray = Array(aArray, aInput)
    Else
      ReDim Preserve aArray(UBound(aArray) + 1) 'If array, increase bounds by 1
      If IsObject(aInput) Then
        Set aArray(UBound(aArray)) = aInput
      Else
        aArray(UBound(aArray)) = aInput
      End If
    End If
    AppendArray = aArray 'return new array
  End If
End Function


'***********************class jungle**************



dim GameOnFF
GameOnFF = 0



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
  Dim s: For s = 0 To UBound(gBOT)
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


'VPW edits
' 00X - benji/iaakki - various versions
' 006 - iaakki - NF script updated one more time, GI resolved, solenoid light control added, insert materials redone and adjusted, insert text flasher should be done
' 009 - iaakki - GI rework and various tweaks. Sidewalls are borked
' 010 - Benji - Added reflective metal texture to walls (only reflect with SSR Enabled) and color graded all the graphics. Added preliminary beveled edges to bottom half of plastics
' 011 - iaakki - RGB GI implemented, various fixes
' 014 - iaakki - pf&text images reworked, inserts made more shallow, GI adjusted different when mode on
' 015 - iaakki - villain spots created, heads material swapped
' 016 - Benji - Adjusted bevels and baked GI onto playfield (experimental) old playfield still in image manager
' 017 - iaakki - redid rgb gi, some cleanup
' 018.6 - Sixtoe - Added VR room and associated switches, replaced most flashers & hooked up flash prims to lighting, added global rgb gi flasher, tweaked spotlights and added character prims to lighting system, aligned ramps, unified timers, added ball shadow, repalced green goblin for comic version, reverted playfield, aligned apron wall and made visible, raised Apron, changed trigger shapes, removed numerous old incorrect walls and lights, raised metal screws and fittings up slightly
' 018.7 - Tomate - replaced wire run prims and textures
' 019 - Sixtoe - Made more holes in playfield for triggers aand targets, added wood drop sides, added new wall for sw43a image, added backglass to gi system, hooked up light refllection flashers on sandman and added some to goblin and lock lights.
' 020 - Benji - Added color grade to  playfield from 019
' 021 - Benji reintroduced more yellows to overall color grade
' 022 - tomate - spiderwebs and diverters redone, new textures for metals, diverters and spiderwebs. Improved wire ramp textures
' 023 - Sixtoe - fixed some depth bias issues, fixed rear back wood, fixed pop bumper flashers, fixed a few flashers
' 024 - tomate - Some tweaks to the textures, shadows in the apron and reduction of the file weight by lowering the resolution of some textures
' 025 - iaakki - Latest NF flips code, cabinetmode, fixed desktop mode, small fix to default pov, readjusted flips and fixed trigger areas
' 026 - iaakki - RGB GI brightness adjusted
' 029 - oqq -  added some missing variables. swapped out 1 at ballwithball colission ( not sure if its the right one ) . Bumper sound ... svapped out Vol(activeball) with 0.2 at RandomSoundBumperTop
' 031 - tomate - add Goblin's bracket and some work on Goblin texture
' RC1 - iaakki - default options set, flipnudge check fixed, sandman standup target bank physics reworked
' RC2 - tomate - new prims and brackets for all villains, Sandman texture fixed so it doesn't collide with the ramp, add some thickness to central plastic, shadows added to apron
' RC3 - tomate - new bumper rings prims and new textures
' RC4 - Sixtoe - redid sandman and ock "holder" prims and textures, fixed goblin lighting, fixed spider sense lighting, unlinked difficulty rubbers from main prim and repositioned correctly, removed floating screw, various depth bias issues corrected, cleaned up unused images
' RC5 - iaakki - Some inserts fixed
' RC6 - iaakki - flupper domes
' RC7 - iaakki - dome flasher adjust, info fields updated, script cleanup
' RC8 - oqq/tomate - add movement to goblin bracket and separate prims
' RC9 - Sixtoe - Added invisible walls to stop ball being lost under top right ramp and hopping over centre targets, updated laser sensor lights, split cover plastics and adjusted position and materials so they're in the right place, added ultra minimal vr room options, colour corrected DMD
' RC10 - Sixtoe - Fixed what I broke because I can't code at 2am, various tweaks
' v1.0.0 - iaakki/sixtoe - Final tune and testing
' v1.0.1 - sixtoe - made plastic protector over green goblin collidable and increased size of protector wall under venom ramp to prevent balls being trapped
' v1.0.2 - apophis - Major physics overhaul (new scripts, flipper triggers, materials, etc). New PWM lamp and flasher support added (not optional).
' v1.0.3 - mcarter78 - Added updated Fleep sounds & Ramp Rolling sounds, Flipper shadows, TweakUI menu, Updated flipper subs with staged flippers
' v1.0.4 - mcarter78 - Fix ramp transition sound for Venom ramp, add autoplunger sound
' v1.0.5 - mcarter78 - Add physical trough and lower apron, tone down wild gate animations, fix sounds for gates and apron
' v1.0.6 - mcarter78 - Fix for stuck ball on orbit wall on short plunge, move plunger & rest down and increase rest angle
' v1.0.7 - mcarter78 - Fix staged flippers implementation, add invisible walls to prevent airballs over apron
' v1.0.8 - mcarter78 - Remove targetbouncer from lower sling posts, another fix for staged flippers
' v1.0.9 - apophis - Added rules to tweak menu.
' v1.0.10 - apophis - Automated VR room. Added other options to tweak menu.
' v1.0.11 - apophis - Fixed automated VR room. Added ramp refractions. Added blocker walls above sandman targets. Flipper strength 2900.
' v1.0.12 - Sixtoe - Loads of layout work on physical table and ramps, made pop bumpers much smaller, added missing metal pin around lock area, adjusted sandman entrance targets and materials,
' v1.0.13 - Sixtoe - Tweaked position of sw6
' v1.0.14 - Sixtoe - Tweaked venom ramp, made orbit post smaller, turned sw6 invisible again, rebuild right physical ramp to take "hump" out of it, added profiled ramp roof to clear vuk loop,
' v1.0.15 - Sixtoe - Made physical area invisible, changed spider sense area layout, reduced complexity of metalsp prim
' v1.0.16 - tomate - new plastic ramps prims and textures, new apron prim and textures, new wireRamps prims and textures, new sidewalls textures, new backwall textures, new plastic textures, new bats baked texture, on/off textures animation to match GI, wall003 and 004 on layer "Visual prims" erased becouse they was duplicated
' v1.0.17 - tomate - new black wooden sidewalls added, config sideblades selection in main menu, physical ramp hided (fix the missing laneguide texture), changed ball settings (provided by HayJay)
' v1.0.18 - tomate - huge 4k batch added, tons of things improved at blender side, new side ramp prim added, unused textures erased
' v1.0.19 - apophis - Wired up all the new visuals in the script. Deleted old table stuff. Updated ball shadows. Removed art sideblade option.
' v1.0.20 - tomate - New 4k bake: more primitives into the bake process, metal walls, bumpers, some blue rubbers. Increase flasher brightness. Added bumper lights.
' v2.0 (RC1) - apophis - Wired up some new lightmaps in the script. Repaired "Visual Prims" and "Insert Prims" layers.
' v2.0 (RC2) - tomate - added new prims into the batch, unwrapped metal wall textures, hidded no needed VPX objects, added extra details to center middle clear plastic
' v2.0 (RC3) - tomate - erased some non-existent posts and moved a blue rubber to the correct position in Sandman area, right spinner set as visible, added roof to Venom ramp
' v2.0 (RC4) - apophis - Fixed octpus & sandman animations. Adjusted gate animations. Added 360 VR Room (Thanks DaRDog!). Fixed trigger visibilities. Made most bakemaps reflections enabled.
' v2.0 (RC5) - apophis - Remerging stuff.
' v2.0 (RC6) - tomate - Burned out backwall bulb replaced :P
' v2.0 (RC7) - DGrimmReaper - VR Animated Flipper Buttons
' v2.0 (RC8) - iaakki - Insert tray adjustments to prevent edges glowing through.
' v2.0 (RC9) - apophis - Removed speck from Nestmap0. Darkened Insert_Text.
' v2.0 Release
' v2.0.1 - tomate - Now spidey has a "left hand", switches set as visible, erased unused textures, file cleaned locked
' v2.0.2 - apophis - Made spinner and apron walls visible. Made BM_Playfield material not active. Removed speck from insert in Nestmap0. Fixed issue where ball sometimes collided with gate on wire ramp near Venom ramp. Made flipper shadows visible and not hide parts behind.
' v2.1 Release

Option Explicit
Randomize
' Dozer/85Vett 2016-02-09
' Original version
' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Thalamus 2018-08-16 : Improved directional sounds
' Added InitVpmFFlipsSAM
' Retro27 2021-10-09
' Added VR Room
' RobbyKingPin 2023-09-27
' Added nFozzy/RothbauerW physics and Fleep Sound
' RobbyKingPin 2023-10-27
' Improved lights and graphics
' Retro27 2023-11-18
' Added new playfields for Standard Edition and LE, changed speaker plastics
' Retro27 2023-11-30
' Sorted the problem with the flippers, making  changes to the plastics, and a flashing speakers and posters on the playfield
' RobbyKingPin 2024-11-30
' Fixed an issue with the center lock pins not dropping at the start of the game

'*******************************************
'  ZOPT: User Options
'*******************************************

'----- Shadow Options -----
Const DynamicBallShadowsOn = 1      '0 = no dynamic ball shadow ("triangles" near slings and such), 1 = enable dynamic ball shadow
Const AmbientBallShadowOn = 1     '0 = Static shadow under ball ("flasher" image, like JP's), 1 = Moving ball shadow ("primitive" object, like ninuzzu's) - This is the only one that behaves like a true shadow!, 2 = flasher image shadow, but it moves like ninuzzu's

'----- General Sound Options -----
Const VolumeDial = 0.8              'Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Const BallRollVolume = 0.5          'Level of ball rolling volume. Value between 0 and 1
Const RampRollVolume = 0.5          'Level of ramp rolling volume. Value between 0 and 1

'*******************************************
'  ZCON: Constants and Global Variables
'*******************************************

Const BallSize = 50             'Ball diameter in VPX units; must be 50
Const BallMass = 1                'Ball mass must be 1
Const tnob = 8                    'Total number of balls the table can hold
Const lob = 0                   'Locked balls
Const cGameName = "rsn_110h"    'Limited Edition
const cSingleLFlip = 0
const cSingleRFlip = 0

'----- VR Room Auto-Detect -----
Dim VRRoom, VR_Obj

If RenderingMode = 2 Then
  Textbox1.visible = 0
  For Each VR_Obj in VRCabinet : VR_Obj.Visible = 1 : Next
  For Each VR_Obj in VRMinimalRoom : VR_Obj.Visible = 1 : Next
  UseVPMDMD = False
  VarHidden = 0
    PinCab_Rails.Visible = 1
Else
  For Each VR_Obj in VRCabinet : VR_Obj.Visible = 0 : Next
  For Each VR_Obj in VRMinimalRoom : VR_Obj.Visible = 0 : Next
End If

Dim tablewidth
tablewidth = Table1.width
Dim tableheight
tableheight = Table1.height
Dim BIP                     'Balls in play
BIP = 0
Dim BIPL                      'Ball in plunger lane
BIPL = False

'******************************************
'   Table Options
'******************************************

'***  Table Select  ***
'***  0 = Limited Edition Cabinet, 1 = Standard Cabinet, 2 = Promotional Cabinet ***
Const VR_Cabinet = 1

'***  Speakers Mod ***
'***    0 = Standard Speakers, 1 = Custom Speakers ***
Const CustomSpeakers = 0

'***  DMD Reflection ***
'***    0 = Disables DMD Reflection, 1 = Enables DMD Reflection ***
Const DMDreflection = 0

'***  Scratched Glass ***
'***    0 = Disables Scratched Glass, 1 = Enables Scratched Glass ***
Const Glassreflection = 0

'***  Side Blades ***
'***  0 = Black Side Blades,  1 = Blue RetroRefurbs Side Blades,  2 = Union Jack Side Blades, 3 = Album Covers Side Blades ***
Const Customcabsides = 0

'***  Instruction Cards ***
'***  0 = Standard Inst Stern Cards ,  1 = German Inst Stern Cards,  2 = Custom Inst Cards ***
Const Custominst = 1


'******************************************
'  VR Options
'******************************************

'*** VR Room ***
'*** 1 = Minimal Room, 2 = Ultra Minimal Room
Const VRRoomChoice = 0

'***  B2S Backglass ***
'***  0 = Disables B2S Backglass, 1 = Enables B2S Backglass ***
Const VRBackglass = 0

'***  VR  Cabinet Tournament Topper ***
'***    0 = Disables Topper, 1 = Enables Topper ***
Const VRtopper = 0

'***  VR  Cabinet DMD Decals ***
'***  0 = Standard DMD Decal,  1 = Blue RetroRefurbs DMD Decal,  2 = Union Jack DMD Decal,  3 = Amplifier RetroRefurbs DMD Decal***
Const Customdmddecal = 0

'***  VR Room Logo ***
'***  0 = Disables VR Room Logo, 1 = Enables VR Room ***
Const VRlogo = 0

'******************************************

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const UseVPMColoredDMD = 1
'----- Desktop Visible Elements -----
Dim VarHidden, UseVPMDMD

LoadVPM "01560000", "sam.VBS", 3.10

Const UseSolenoids = 1,UseLamps = 1,UseSync = 1,UseGI = 0,SSolenoidOn = "SolOn",SSolenoidOff = "SolOff",SFlipperOn = "FlipperUp",SFlipperOff = "FlipperDown",sCoin = "coin3"

SolCallback(1)        = "SolTrough"
SolCallback(2)        = "SolAutofire"
SolCallback(3)        = "SolCenterLock"
SolCallback(4)        = "SolCenterLockLatch"
SolCallback(5)        = "MagLeft"
SolCallback(6)        = "SolControlGate"
SolCallback(7)        = "MagRight"
'SolCallBack(9)         = "vpmSolSound ""bumper2"","
'SolCallBack(10)      = "vpmSolSound ""bumper2"","
'SolCallBack(11)      = "vpmSolSound ""bumper2"","
'SolCallback(13)            = "vpmSolSound ""left_slingshot_new"","
'SolCallback(14)            = "vpmSolSound ""right_slingshot_new"","

SolCallback(15)       = "SolLFlipper"
SolCallback(16)       = "SolRFlipper"

SolCallback(17)       = "SolLUD"
SolCallback(18)             = "MickRelayLeft2"
SolCallback(19)             = "MickRelayRight2"

SolCallBack(20)         = "Sol20"
SolCallBack(21)         = "Sol21"
SolCallBack(22)         = "Sol22"
SolCallBack(23)         = "Sol23"
SolCallBack(25)         = "Sol25"
SolCallBack(26)         = "Sol26"
SolCallBack(27)         = "Sol27"
SolCallBack(28)         = "Sol28"
SolCallBack(29)         = "Sol29"
SolCallBack(31)         = "Sol31"

SolCallback(30)       = "SolCUD"
SolCallback(32)       = "SolRUD"

Dim bsTrough, x, xx, Mag1, Mag2, object

Sub Table1_Init

' Thalamus : Was missing 'vpminit me'
  vpminit me

    Controller.GameName= cGameName
  Controller.SplashInfoLine="Rolling Stones (Stern 2011)"
  Controller.Games(cGameName).Settings.Value("rol")=0
  Controller.ShowTitle=0
  Controller.ShowDMDOnly=1
  Controller.ShowFrame=0
  Controller.HandleMechanics=0
    'Controller.Hidden = varhidden
  On Error Resume Next
    Controller.Run
    If Err Then MsgBox Err.Description
  On Error Goto 0
  PinMAMETimer.Interval=PinMAMEInterval:PinMAMETimer.Enabled=1:vpmNudge.TiltSwitch=-7
  vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingshot)
  vpmNudge.Sensitivity=5

    'Set bsTrough=New cvpmBallStack
  'bsTrough.InitSw 0,21,20,19,18,17,0,0
  'bsTrough.InitSw 0,21,20,19,18,0,0,0
  'bsTrough.InitKick BallRelease,110,5
  'bsTrough.InitExitSnd "BallRel","Solon"
  'bsTrough.KickBalls = 1
  'bsTrough.Balls=5

    Set mag1= New cvpmMagnet
    mag1.InitMagnet Magnet1, 26 '16
    mag1.GrabCenter = False

    Set mag2= New cvpmMagnet
    mag2.InitMagnet Magnet2, 26 '16
    mag2.GrabCenter = False

    Plunger1.Pullback
    Controller.Switch(71) = 1
    W17.isdropped = 1
  W18.isdropped = 1
  W19.isdropped = 1
  W20.isdropped = 1
  W21.isdropped = 1
    LUD.isdropped = 1
    RUD.isdropped = 1
    CUD.isdropped = 1
    MW1.isdropped = 1
  MW2.isdropped = 1
  MW3.isdropped = 1
  MW4.isdropped = 1
  MW5.isdropped = 1
  MW6.isdropped = 1
  MW_PARK.isdropped = 1
    Mick_PRIM.RotZ = 37
  MW1a.isdropped = 1
  MW2a.isdropped = 1
  MW3a.isdropped = 1
  MW4a.isdropped = 1
  MW5a.isdropped = 1
  MW6a.isdropped = 1
  MW_PARKa.isdropped = 1
    between1.isdropped = 1
    between2.isdropped = 1
    between3.isdropped = 1
    between4.isdropped = 1
    between5.isdropped = 1
    between6.isdropped = 1
    GI_On
  InitVpmFFlipsSAM

setup_backglass()


End Sub


'**********************************************************************************************************
'VR Room stuff
'**********************************************************************************************************

Sub TimerVRPlunger_Timer
  If Primary_new_plunger.Y < 1000 then
  Primary_new_plunger.Y = Primary_new_plunger.Y + 5
  End If
End Sub

Sub TimerVRPlunger2_Timer
  'debug.print plunger.position
  'Primary_new_plunger.Y = 1080.2957 + (5* Plunger.Position) -20
  Primary_new_plunger.Y = 2180 + (5* Plunger.Position) -20
End Sub

'Cabinet Side Artwork

Select Case Customcabsides
    Case 0
        SideCab.image = "CabSidesOriginal"
  Case 1
    SideCab.image = "CabSides"
  Case 2
    SideCab.image = "CabSides_1"
  Case 3
    SideCab.image = "CabSides_2"
End Select

'Cabinet Instruction Cards

Select Case Custominst
    Case 0
        Card1.image = "Inst_755-51B8-12-Y"
  if VRtopper = 1 Then
    Card2.image = "Inst_755-5400-09-Y"
    Else
    Card2.image = "Inst_755-5400-11-Y"
  End If
  Case 1
    Card1.image = "Inst_755-51B1-03-Y"
    if VRtopper = 1 Then
    Card2.image = "Inst_755-5400-09-Y"
    Else
    Card2.image = "Inst_755-5400-11-Y"
  End If
  Case 2
  If VR_Cabinet = 0 then
    Card1.image = "Inst_Custom_L_LE"
    Card2.image = "Inst_Custom_R_LE"
    Else
    Card1.image = "Inst_Custom_L_Standard"
    Card2.image = "Inst_Custom_R_Standard"
  End If
End Select


Select Case CustomSpeakers
  Case 0
    RampAL.visible = True
    RampAR.visible = True
  Case 1
    Amp_Left.visible = True
    Amp_Right.visible = True
End Select

'DMD Reflection
Primary_DMD_reflection.visible = DMDreflection

'Scratched Glass
Primary_Glass_scratches.visible = Glassreflection

'VR Logo
VR_Room_Logo.visible = VRlogo

'VR Topper
Primary_topper.visible = VRtopper

'DMD Decal
Select Case Customdmddecal
  Case 0
    PinCab_CustDMD.image = "Pincab_DMD_Decal"
  Case 1
    PinCab_CustDMD.image = "Pincab_DMD_Decal_1"
  Case 2
    PinCab_CustDMD.image = "Pincab_DMD_Decal_2"
  Case 3
    PinCab_CustDMD.image = "Pincab_DMD_Decal_3"
End Select

'B2S Backglass
PinCab_Backglass.visible = VRBackglass

Select Case VRBackglass
  Case 0
    Flasher005.imageA = "_defglowtrans"
    Flasher005.imageB = "_defglowtrans"
    Flasher006.imageA = "_defglowtrans"
    Flasher006.imageB = "_defglowtrans"
    Flasher007.imageA = "_defglowtrans"
    Flasher007.imageB = "_defglowtrans"
    Flasher008.imageA = "_defglowtrans"
    Flasher008.imageB = "_defglowtrans"
  Case 1
    Flasher005.imageA = "<none>"
    Flasher005.imageB = "<none>"
    Flasher006.imageA = "<none>"
    Flasher006.imageB = "<none>"
    Flasher007.imageA = "<none>"
    Flasher007.imageB = "<none>"
    Flasher008.imageA = "<none>"
    Flasher008.imageB = "<none>"
    BGHigh.visible = 0
    BGDark.visible = 0
    BGHigh1.visible = 0
End Select

'Standard Canbinet
Select Case VR_Cabinet
  Case 0 'LE Cabinet
    Table1.image = "Playfield_LE"
    PinCab_Backbox.image = "Pincab_backbox_LE"
    PinCab_Grills.image = "Pincab_DMD_Grill_Stern_RS"
    PinCab_Rails.material = "Metal0.2"
    VR_LegsFront.material = "Metal0.2"
    VR_LegsBack.material = "Metal0.2"
    PinCab_Metals.material = "Metal0.2"
    Ramp55.image = "apron_LE"
  Case 1 'Standard Cabinet
    Table1.image = "Playfield_Standard"
    PinCab_Backbox.image = "Pincab_backbox"
    PinCab_Grills.image = "Pincab_DMD_Grill_Stern"
    PinCab_Rails.material = "Metal_Black_Powdercoat"
    VR_LegsFront.material = "Metal_Black_Powdercoat"
    VR_LegsBack.material = "Metal_Black_Powdercoat"
    PinCab_Metals.material = "Metal_Black_Powdercoat"
    VRFlipperButtonLeft1.visible=0
    VRButtonRings1.visible=0
    VRFlipperButtonRight1.visible=0
    LELeftRailBlade.visible =0
    LERightRailBlade.visible =0
    LeftRailscrew1.visible =0
    LeftRailscrew2.visible =0
    LeftRailscrew3.visible =0
    RightRailscrew1.visible =0
    RightRailscrew2.visible =0
    RightRailscrew3.visible =0
    CUD.visible =0
    CUD.sidevisible =0
    PostCUD.visible =0
    CUD_LAMP.visible =0
    CUD.collidable =0
    RUD.visible =0
    RUD.sidevisible =0
    LUD.visible =0
    LUD.sidevisible =0
    F29.visible =0
    F29a.visible =0
    F29b.visible =0
    F29c.visible =0
    F29d.visible =0
    F29e.visible =0
    Ramp55.image = "apron_Standard"
  Case 2 'Promo Cabinet
    Table1.image = "Playfield_Standard_alt"
    PinCab_Backbox.image = "Pincab_backbox"
    PinCab_Grills.image = "Pincab_DMD_Grill_Stern"
    PinCab_Rails.material = "Metal_Black_Powdercoat"
    VR_LegsFront.material = "Metal_Black_Powdercoat"
    VR_LegsBack.material = "Metal_Black_Powdercoat"
    PinCab_Metals.material = "Metal_Black_Powdercoat"
    VRFlipperButtonLeft1.visible=0
    VRButtonRings1.visible=0
    VRFlipperButtonRight1.visible=0
    LELeftRailBlade.visible =0
    LERightRailBlade.visible =0
    LeftRailscrew1.visible =0
    LeftRailscrew2.visible =0
    LeftRailscrew3.visible =0
    RightRailscrew1.visible =0
    RightRailscrew2.visible =0
    RightRailscrew3.visible =0
    CUD.visible =0
    CUD.sidevisible =0
    PostCUD.visible =0
    CUD_LAMP.visible =0
    CUD.collidable =0
    RUD.visible =0
    RUD.sidevisible =0
    LUD.visible =0
    LUD.sidevisible =0
    F29.visible =0
    F29a.visible =0
    F29b.visible =0
    F29c.visible =0
    F29d.visible =0
    F29e.visible =0
    Ramp55.image = "apron_Standard"
End Select

' ***************************************************************************
' ***************************************************************************
'                    BASIC FSS(DMD,SS,EM) SETUP CODE
' ****************************************************************************

Dim xoff,yoff,zoff,xrot,zscale, xcen,ycen

Sub setup_backglass()
  xoff =470
  yoff =-20
  zoff =938
  xrot = -90

  bgdark.x = xoff +25
  bgdark.y = yoff -40
  bgdark.height = zoff  +250
  bgdark.rotx = xrot

  bgHigh.x = xoff +25
  bgHigh.y = yoff -40
  bgHigh.height = zoff  +250
  bgHigh.rotx = xrot

  bgHigh1.x = xoff +25
  bgHigh1.y = yoff -40
  bgHigh1.height = zoff +250
  bgHigh1.rotx = xrot

  center_graphix()

End Sub


Dim BGArr
BGArr=Array (Flasher005,Flasher006,Flasher007,Flasher008)


Sub center_graphix()
  Dim xx,yy,yfact,xfact,xobj
  zscale = 0.00000001

  xcen =(1068 /2) - (72 / 2)
  ycen = (1128 /2 ) + (250 /2)

  yfact =0 'y fudge factor (ycen was wrong so fix)
  xfact =0

  For Each xobj In BGArr
    xx =xobj.x

    xobj.x = (xoff -xcen) + xx +xfact
    yy = xobj.y ' get the yoffset before it is changed
    xobj.y =yoff

    If(yy < 0.) Then
    yy = yy * -1
    End If

    xobj.height =( zoff - ycen) + yy - (yy * zscale) + yfact

    xobj.rotx = xrot
    xobj.visible =0 ' for testing
  Next
End Sub

Sub Table1_KeyDown(ByVal KeyCode)
  If keycode = PlungerKey Then
    Plunger.Pullback
    SoundPlungerPull
  End If
  If keycode = LeftFlipperKey Then
    FlipperActivate LeftFlipper, LFPress
    Controller.Switch(87) = 1
    VRFlipperButtonLeft.X = VRFlipperButtonLeft.X +8
  End If

  If keycode = RightFlipperKey Then
    FlipperActivate RightFlipper, RFPress
    Controller.Switch(85) = 1
    VRFlipperButtonRight.X = VRFlipperButtonRight.X -8
  End If
  If keycode = LeftMagnaSave Then Controller.Switch(88)=1 : VRFlipperButtonLeft1.X = VRFlipperButtonLeft1.X +8
  If keycode = RightMagnaSave Then Controller.Switch(86)=1 : VRFlipperButtonRight1.X = VRFlipperButtonRight1.X -8
  If keycode = StartGameKey Then Controller.Switch(14)=1 :VRStartButton.Y = VRStartButton.Y -5 :VRStartButton2.Y = VRStartButton2.Y -5
  'If keycode=3 Then Controller.Switch(88)=1:Controller.Switch(86)=1
  If keycode = keyfront Then Controller.Switch(15)=1 :VRTourneyButton.Y =VRTourneyButton.Y -5
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
  If keycode = MechanicalTilt Then
    SoundNudgeCenter() 'Send the Tilting command to the ROM (usually by pulsing a Switch), or run the tilting code for an orginal table
  End If
  If keycode = AddCreditKey Or keycode = AddCreditKey2 Then
    Select Case Int(Rnd * 3)
      Case 0
        PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
      Case 1
        PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
      Case 2
        PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
    End Select
  End If
  If vpmKeyDown(KeyCode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
  If KeyCode = PlungerKey Then
    Plunger.Fire
    If BIPL = 1 Then
      SoundPlungerReleaseBall()   'Plunger release sound when there is a ball in shooter lane
    Else
      SoundPlungerReleaseNoBall() 'Plunger release sound when there is no ball in shooter lane
    End If
  End If
  If keycode = LeftFlipperKey Then
    FlipperDeActivate LeftFlipper, LFPress
    Controller.Switch(87) = 0
    VRFlipperButtonLeft.X = VRFlipperButtonLeft.X -8
  End If

  If keycode = RightFlipperKey Then
    FlipperDeActivate RightFlipper, RFPress
    Controller.Switch(85) = 0
    VRFlipperButtonRight.X = VRFlipperButtonRight.X +8
  End If
  'If KeyCode = LeftFlipperKey Then Controller.Switch(73) = 0 :VRFlipperButtonLeft.X = VRFlipperButtonLeft.X -8
  'If KeyCode = RightFlipperKey Then Controller.Switch(75) = 0 :VRFlipperButtonRight.X = VRFlipperButtonRight.X +8
    If keycode = LeftMagnaSave Then Controller.Switch(88) = 0 :VRFlipperButtonLeft1.X = VRFlipperButtonLeft1.X -8
    If keycode = RightMagnaSave Then Controller.Switch(86) = 0 :VRFlipperButtonRight1.X = VRFlipperButtonRight1.X +8
  If keycode = StartGameKey Then Controller.Switch(14) = 0 :VRStartButton.Y = VRStartButton.Y +5 :VRStartButton2.Y = VRStartButton2.Y +5
    'If keycode = 3 Then Controller.Switch(88) = 0:Controller.Switch(86)=0
  If keycode = keyfront Then Controller.Switch(15) = 0 :VRTourneyButton.Y =VRTourneyButton.Y +5
  If vpmKeyUp(KeyCode) Then Exit Sub
End Sub

Set MotorCallback=GetRef("UpdateMultipleLamps")

Sub UpdateMultipleLamps
  FGI1.visible = Light40.State
  FGI2.visible = Light40.State
  FGI3.visible = Light40.State
  FGI4.visible = Light40.State
  FGI5.visible = Light40.State

  Flasher1.visible = Light40.State
  Flasher2.visible = Light40.State
  Flasher3.visible = Light40.State

  F1.visible = Light40.State
  F2.visible = Light26.State
End Sub

Set Lights(3) = L3
Set Lights(4) = L4
Set Lights(5) = L5
Set Lights(6) = L6
Set Lights(7) = L7
Set Lights(8) = L8
Set Lights(9) = L9
Set Lights(10) = L10
Set Lights(11) = L11
Set Lights(12) = L12
Set Lights(13) = L13
Set Lights(14) = L14
Set Lights(15) = L15
Set Lights(16) = L16
Set Lights(17) = L17
Set Lights(18) = L18
Set Lights(19) = L19
Set Lights(20) = L20
Set Lights(21) = L21
Set Lights(22) = L22
Set Lights(23) = L23
Set Lights(24) = L24
Set Lights(25) = L25
Set Lights(26) = L26
Set Lights(27) = L27
Set Lights(28) = L28
Set Lights(29) = L29
Set Lights(30) = L30
Set Lights(31) = L31
Set Lights(32) = L32
Set Lights(33) = L33
Set Lights(34) = L34
Set Lights(35) = L35
Set Lights(36) = L36
Set Lights(37) = L37
Set Lights(38) = L38
Set Lights(39) = L39
Set Lights(40) = L40
Set Lights(41) = L41
Set Lights(42) = L42
Set Lights(43) = L43
Set Lights(44) = L44
Set Lights(45) = L45
Set Lights(46) = L46
Set Lights(47) = L47
Set Lights(48) = L48
Set Lights(49) = L49
Set Lights(50) = L50
Set Lights(51) = L51
Set Lights(52) = L52
Set Lights(53) = L53
Set Lights(58) = L58
Set Lights(60) = L60
Set Lights(61) = L61
Set Lights(62) = L62

' For VR Start and Tournament Buttons...
Set Lights(1) = L01
Set Lights(2) = L02

'Switch Subs

Sub SW50_Hit()
  Controller.Switch(50) = 1
End Sub

Sub SW50_UnHit()
  Controller.Switch(50) = 0
End Sub

Sub SW23_Hit()
  BIPL = True
  Controller.Switch(23) = 1
  Gi_Off
End Sub

Sub Gi_Start_Hit()
  Gi_On
End Sub

Sub SW23_UnHit()
  BIPL = False
  Controller.Switch(23) = 0
  If Controller.Switch(71) = False And ActiveBall.id = 666 Then
    Controller.Switch(71) = True
  End If
End Sub

Sub SW6_Hit():Controller.Switch(6) = 1:End Sub
Sub SW6_UnHit():Controller.Switch(6) = 0:End Sub
Sub SW7_Hit():Controller.Switch(7) = 1:End Sub
Sub SW7_UnHit():Controller.Switch(7) = 0:End Sub
Sub SW8_Hit():Controller.Switch(8) = 1:End Sub
Sub SW8_UnHit():Controller.Switch(8) = 0:End Sub
Sub SW9_Hit():Controller.Switch(9) = 1:End Sub
Sub SW9_UnHit():Controller.Switch(9) = 0:End Sub

Sub SW45_Hit():vpmTimer.PulseSw 45:Gi_Off:End Sub
Sub SW48_Hit():vpmTimer.PulseSw 48:End Sub

Sub SW29_Hit():vpmTimer.PulseSw 29:End Sub
Sub SW28_Hit():vpmTimer.PulseSw 28:End Sub

Sub SW25_Hit():vpmTimer.PulseSw 25:End Sub
Sub SW24_Hit():vpmTimer.PulseSw 24:End Sub
Sub SW41_Spin() : vpmTimer.PulseSw 41 : SoundSpinner SW41 : End Sub

Sub SW42_Hit():vpmTimer.PulseSw 42:Gi_Off:End Sub
Sub SW43_Hit():vpmTimer.PulseSw 43:End Sub

Sub SW44_Hit():vpmTimer.PulseSw 44:End Sub

Sub SW53_Hit():Controller.Switch(53) = 1:gflash = 1:gi_flash.enabled = 1:gi_flash_stop.enabled = 1:End Sub
Sub SW53_UnHit():Controller.Switch(53) = 0:End Sub

Sub SW46_Hit():Controller.Switch(46) = 1:End Sub
Sub SW46_UnHit():Controller.Switch(46) = 0:End Sub

Sub SW51_Hit():TargetBouncer ActiveBall, 1:vpmTimer.PulseSw 51:Sw51p.transx = -10:me.timerenabled = 1:End Sub
Sub SW51_Timer():Sw51p.transx = 0:me.timerenabled = 0:End Sub

Sub SW56_Hit():TargetBouncer ActiveBall, 1:vpmTimer.PulseSw 56:Sw56p.transx = -10:me.timerenabled = 1:End Sub
Sub SW56_Timer():Sw56p.transx = 0:me.timerenabled = 0:End Sub

Sub SW57_Hit():TargetBouncer ActiveBall, 1:vpmTimer.PulseSw 57:Sw57p.transx = -10:me.timerenabled = 1:End Sub
Sub SW57_Timer():Sw57p.transx = 0:me.timerenabled = 0:End Sub

Sub SW1_Hit():TargetBouncer ActiveBall, 1:vpmTimer.PulseSw 1:Sw1p.transx = -10:me.timerenabled = 1:End Sub
Sub SW1_Timer():Sw1p.transx = 0:me.timerenabled = 0:End Sub

Sub SW2_Hit():TargetBouncer ActiveBall, 1:vpmTimer.PulseSw 2:Sw2p.transx = -10:me.timerenabled = 1:End Sub
Sub SW2_Timer():Sw2p.transx = 0:me.timerenabled = 0:End Sub

Sub SW3_Hit():TargetBouncer ActiveBall, 1:vpmTimer.PulseSw 3:Sw3p.transx = -10:me.timerenabled = 1:End Sub
Sub SW3_Timer():Sw3p.transx = 0:me.timerenabled = 0:End Sub

Sub SW10_Hit():TargetBouncer ActiveBall, 1:vpmTimer.PulseSw 10:Sw10p.transx = -10:me.timerenabled = 1:End Sub
Sub SW10_Timer():Sw10p.transx = 0:me.timerenabled = 0:End Sub

Sub SW11_Hit():TargetBouncer ActiveBall, 1:vpmTimer.PulseSw 11:Sw11p.transx = -10:me.timerenabled = 1:End Sub
Sub SW11_Timer():Sw11p.transx = 0:me.timerenabled = 0:End Sub

Sub SW12_Hit():TargetBouncer ActiveBall, 1:vpmTimer.PulseSw 12:Sw12p.transx = -10:me.timerenabled = 1:End Sub
Sub SW12_Timer():Sw12p.transx = 0:me.timerenabled = 0:End Sub

Sub SW54_Hit():TargetBouncer ActiveBall, 1:vpmTimer.PulseSw 54:Sw54p.transx = -10:me.timerenabled = 1:gflash = 1:gi_flash.enabled = 1:gi_flash_stop.enabled = 1:End Sub
Sub SW54_Timer():Sw54p.transx = 0:me.timerenabled = 0:End Sub

Sub SW55_Hit():TargetBouncer ActiveBall, 1:vpmTimer.PulseSw 55:Sw55p.transx = -10:me.timerenabled = 1:gflash = 1:gi_flash.enabled = 1:gi_flash_stop.enabled = 1:End Sub
Sub SW55_Timer():Sw55p.transx = 0:me.timerenabled = 0:End Sub


Sub Magnet1_Hit()
  If NOT ActiveBall.id = 666 Then
    Mag1.AddBall ActiveBall
  End If
End Sub

Sub Magnet1_UnHit()
  Mag1.RemoveBall ActiveBall
End Sub

Sub Magnet2_Hit()
  If NOT ActiveBall.id = 666 Then
    Mag2.AddBall ActiveBall
  End If
End Sub

Sub Magnet2_UnHit()
  Mag2.RemoveBall ActiveBall
End Sub

'Subs for Various Things.

Sub SolTrough(Enabled)
  If Enabled Then
    SW21.kick 37,30
    vpmTimer.PulseSw 22
        RandomSoundBallRelease sw21
        BIP = BIP + 1
  End If
 End Sub

Sub Bumper1_Hit
  RandomSoundBumperTop Bumper1
  'B1L1.State = 1:B1L2. State = 1
  Me.TimerEnabled = 1
    vpmTimer.PulseSw 30
End Sub

Sub Bumper1_Timer
  'B1L1.State = 0:B1L2. State = 0
  Me.Timerenabled = 0
End Sub

Sub Bumper2_Hit
  RandomSoundBumperMiddle Bumper2
  'B2L1.State = 1:B2L2. State = 1
  Me.TimerEnabled = 1
    vpmTimer.PulseSw 31
End Sub

Sub Bumper2_Timer
  'B2L1.State = 0:B2L2. State = 0
  Me.Timerenabled = 0
End Sub

Sub Bumper3_Hit
  RandomSoundBumperBottom Bumper3
  'B3L1.State = 1:B3L2. State = 1
  Me.TimerEnabled = 1
    vpmTimer.PulseSw 32
End Sub

Sub Bumper3_Timer
  'B3L1.State = 0:B3L2. State = 0
  Me.Timerenabled = 0
End Sub

'*******************************************
' ZFLP: Flippers
'*******************************************

Const ReflipAngle = 20

' Flipper Solenoid Callbacks (these subs mimics how you would handle flippers in ROM based tables)
Sub SolLFlipper(Enabled) 'Left flipper solenoid callback
  If Enabled Then
    LF.Fire  'leftflipper.rotatetoend

    If Leftflipper.currentangle < LeftFlipper.endangle + ReflipAngle Then
      RandomSoundReflipUpLeft LeftFlipper
    Else
      SoundFlipperUpAttackLeft LeftFlipper
      RandomSoundFlipperUpLeft LeftFlipper
    End If
  Else
    LeftFlipper.RotateToStart
    If LeftFlipper.currentangle < LeftFlipper.startAngle - 5 Then
      'RandomSoundFlipperDownLeft LeftFlipper
    End If
    FlipperLeftHitParm = FlipperUpSoundLevel
  End If
End Sub

Sub SolRFlipper(Enabled) 'Right flipper solenoid callback
  If Enabled Then
    RF.Fire 'rightflipper.rotatetoend

    If Rightflipper.currentangle > RightFlipper.endangle - ReflipAngle Then
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
  CheckLiveCatch ActiveBall, LeftFlipper, LFCount, parm
  LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
  CheckLiveCatch ActiveBall, RightFlipper, RFCount, parm
  RightFlipperCollide parm
End Sub

Sub FlipperVisualUpdate 'This subroutine updates the flipper shadows and visual primitives
  'FlipperLSh.RotZ = LeftFlipper.CurrentAngle
  'FlipperRSh.RotZ = RightFlipper.CurrentAngle
  LFLogo.RotY = LeftFlipper.CurrentAngle
    RFlogo.RotY = RightFlipper.CurrentAngle
  If L60.state = 1 Then P60.blenddisablelighting = 10
  If L60.state = 0 Then P60.blenddisablelighting = 0
  If L61.state = 1 Then P61.blenddisablelighting = 10
  If L61.state = 0 Then P61.blenddisablelighting = 0
  If L62.state = 1 Then P62.blenddisablelighting = 10
  If L62.state = 0 Then P62.blenddisablelighting = 0
End Sub

Sub SolLUD(enabled)
  If enabled Then
    LUD.isdropped = 0
    Else
    LUD.isdropped = 1
  End If
  If enabled and VR_Cabinet = 0 then
    PlaySound ("Apron_Soft_4")
  End If
End Sub

Sub postlud_hit()
  If LUD.isdropped = 0 Then ActiveBall.Velz = 10 End If
end sub

Sub SolRUD(enabled)
  If enabled Then
    RUD.isdropped = 0
    Else
    RUD.isdropped = 1
  End If
  If enabled and VR_Cabinet = 0 then
    PlaySound ("Apron_Soft_4")
  End If
End Sub

Sub postrud_hit()
  If RUD.isdropped = 0 Then ActiveBall.Velz = 10 End If
end sub

Sub SolCUD(enabled)
  If enabled Then
    CUD.isdropped = 0
    CUD_Lamp.State = 2
  Else
    CUD.isdropped = 1
    CUD_Lamp.State = 0
  End If
End Sub

Sub postcud_hit()
  If CUD.isdropped = 0 Then ActiveBall.Velz = 10 End If
End Sub

Sub MagLeft(enabled)
  If enabled Then
    Mag1.MagnetOn = 1
  Else
    Mag1.MagnetOn = 0
  End If
End Sub

Sub MagRight(enabled)
  If enabled Then
    Mag2.MagnetOn = 1
  Else
    Mag2.MagnetOn = 0
  End If
End Sub

Sub SolControlGate(enabled)
  If enabled Then
    UCG.open = 1
  Else
    UCG.open = 0
  End If
End Sub

Sub SolCenterLock(enabled)
  If enabled Then
    BL1.isdropped = 0
    BL2.isdropped = 0
  End If
End Sub

Sub SolCenterLockLatch(enabled)
  If enabled Then
    BL1.isdropped = 0
    BL2.isdropped = 0
  Else
    BL1.isdropped = 1
    BL2.isdropped = 1
  End If
End Sub

Sub SolAutoFire(enabled)
  If enabled Then
    Plunger1.Fire
    SoundPlungerReleaseBall
  Else
    Plunger1.Pullback
  End If
End Sub

Sub Drain_Hit()
  If ActiveBall.id = 666 Then
    me.destroyball
    Set wball = Kicker_Load.createball
    wball.image = "Powerball2"
    wball.id = 666
    Kicker_Load.kick 45,5
    RandomSoundDrain Drain
    BIP = BIP - 1
    Drain.enabled = 0
    TDrain.enabled = 1
    If BIP < 1 Then
      Gi_Off
    End If
  Else
    me.destroyball
    Kicker_Load.createball
    Kicker_Load.kick 45,5
    RandomSoundDrain Drain
    BIP = BIP - 1
    Drain.enabled = 0
    TDrain.enabled = 1
    If BIP < 1 Then
      Gi_Off
    End If
  End If
End Sub

Sub TDrain_Timer()
  Drain.enabled = 1
  me.enabled = 0
End Sub

'Mick Left / Right Direction Handler.

Sub MickRelayLeft2(enabled)
  If enabled Then
    'LL.State = 1
    micky_left.enabled = 1
  Else
    'LL.State = 0
    micky_left.enabled = 0
  End If
End Sub

Sub MickRelayRight2(enabled)
  If enabled Then
    micky_right.enabled = 1
    'LR.State = 1
  Else
    'LR.State = 0
    micky_right.enabled = 0
  End If
End Sub

Sub micky_right_timer()
  Mick_Prim.RotZ = Mick_Prim.RotZ + 1
  If Mick_Prim.RotZ>36 Then Mick_Prim.RotZ=36
End Sub

Sub micky_left_timer()
  Mick_Prim.RotZ = Mick_Prim.RotZ - 1
  If Mick_Prim.RotZ<-27 Then Mick_Prim.RotZ=-27
End Sub

'Shake the Primitive when the ball hits the wall tied with  Micks position.

Dim mshake

Sub micky_shake_timer()
  Select Case mshake
    Case 1:Mick_Prim.TransY = -15:mshake = 2
    Case 2:Mick_Prim.TransY = 0:mshake = 3
    Case 3:Mick_Prim.TransY = -7:mshake = 4
    Case 4:Mick_Prim.TransY = 0:me.enabled = 0
  End Select
End Sub

'Mick on a Stick (TM) :) Watchdog Timer.

Sub Move_Mick_Timer()

  If Mick_Prim.RotZ => 34 And Mick_Prim.RotZ <= 36 Then
    Controller.Switch(33) = 1
    MW1.isdropped = 0
    MW1a.isdropped = 0
  Else
    Controller.Switch(33) = 0
    MW1.isdropped = 1
    MW1a.isdropped = 1
  End If

  If Mick_Prim.RotZ => 27 And Mick_Prim.RotZ <= 33 Then
    between1.isdropped = 0
  Else
    between1.isdropped = 1
  End If

  If Mick_Prim.RotZ => 24 And Mick_Prim.RotZ <= 26 Then
    Controller.Switch(34) = 1
    MW2.isdropped = 0
    MW2a.isdropped = 0
  Else
    Controller.Switch(34) = 0
    MW2.isdropped = 1
    MW2a.isdropped = 1
  End If

  If Mick_Prim.RotZ => 17 And Mick_Prim.RotZ <= 23 Then
    between2.isdropped = 0
  Else
    between2.isdropped = 1
  End If

  If Mick_Prim.RotZ => 14 And Mick_Prim.RotZ <= 16 Then
    Controller.Switch(35) = 1
    MW3.isdropped = 0
    MW3a.isdropped = 0
  Else
    Controller.Switch(35) = 0
    MW3.isdropped = 1
    MW3a.isdropped = 1
  End If

  If Mick_Prim.RotZ => 6 And Mick_Prim.RotZ <= 13 Then
    between3.isdropped = 0
  Else
    between3.isdropped = 1
  End If

  If Mick_Prim.RotZ => 3  And Mick_Prim.RotZ <= 5 Then
    Controller.Switch(36) = 1
    MW4.isdropped = 0
    MW4a.isdropped = 0
  Else
    Controller.Switch(36) = 0
    MW4.isdropped = 1
    MW4a.isdropped = 1
  End If

  If Mick_Prim.RotZ => -5 And Mick_Prim.RotZ <= 2 Then
    between4.isdropped = 0
  Else
    between4.isdropped = 1
  End If

  If Mick_Prim.RotZ => -7 And Mick_Prim.RotZ <= -6 Then
    Controller.Switch(39) = 1
    MW_Park.isdropped = 0
    MW_Parka.isdropped = 0
  Else
    Controller.Switch(39) = 0
    MW_Park.isdropped = 1
    MW_Parka.isdropped = 1
  End If

  If Mick_Prim.RotZ => -14 And Mick_Prim.RotZ <= -8 Then
    between5.isdropped = 0
  Else
    between5.isdropped = 1
  End If

  If Mick_Prim.RotZ => -17 And Mick_Prim.RotZ <= -15 Then
    Controller.Switch(37) = 1
    MW5.isdropped = 0
    MW5a.isdropped = 0
  Else
    Controller.Switch(37) = 0
    MW5.isdropped = 1
    MW5a.isdropped = 1
  End If

  If Mick_Prim.RotZ => -24 And Mick_Prim.RotZ <= -18 Then
    between6.isdropped = 0
  Else
    between6.isdropped = 1
  End If

  If Mick_Prim.RotZ => -27 And Mick_Prim.RotZ <= -25 Then
    Controller.Switch(38) = 1
    MW6.isdropped = 0
    MW6a.isdropped = 0
  Else
    Controller.Switch(38) = 0
    MW6.isdropped = 1
    MW6a.isdropped = 1
  End If

End Sub

'Mick on a Stick Target Handler.

Sub MW1_Hit()
  vpmTimer.PulseSw 72
  mshake = 1
  micky_shake.enabled = 1
  gflash = 1:gi_flash.enabled = 1:gi_flash_stop.enabled = 1
End Sub

Sub MW2_Hit()
  vpmTimer.PulseSw 72
  mshake = 1
  micky_shake.enabled = 1
  gflash = 1:gi_flash.enabled = 1:gi_flash_stop.enabled = 1
End Sub

Sub MW3_Hit()
  vpmTimer.PulseSw 72
  mshake = 1
  micky_shake.enabled = 1
  gflash = 1:gi_flash.enabled = 1:gi_flash_stop.enabled = 1
End Sub

Sub MW4_Hit()
  vpmTimer.PulseSw 72
  mshake = 1
  micky_shake.enabled = 1
  gflash = 1:gi_flash.enabled = 1:gi_flash_stop.enabled = 1
End Sub

Sub MW5_Hit()
  vpmTimer.PulseSw 72
  mshake = 1
  micky_shake.enabled = 1
  gflash = 1:gi_flash.enabled = 1:gi_flash_stop.enabled = 1
End Sub

Sub MW_Park_Hit()
  vpmTimer.PulseSw 72
  mshake = 1
  micky_shake.enabled = 1
  gflash = 1:gi_flash.enabled = 1:gi_flash_stop.enabled = 1
End Sub

Sub MW6_Hit()
  vpmTimer.PulseSw 72
  mshake = 1
  micky_shake.enabled = 1
  gflash = 1:gi_flash.enabled = 1:gi_flash_stop.enabled = 1
End Sub

'Flasher Handlers - Launches A Timer to work around the constant state ON SAM issue.

Dim f20time

Sub Sol20(enabled)
  If enabled Then
    F20.state = 1
    Mick_PRIM.image = "Mick_Bright"
    f20time = 1
    Fade20.enabled = 1
    Flasher005.visible =1
  Else
    F20.state = 0
    Fade20.enabled = 0
    Mick_PRIM.image = "Mick"
    Flasher005.visible = 0
  End If
End Sub

Sub Fade20_Timer()
  Select Case f20time
    Case 1:
      F20.state = 0
      Mick_PRIM.image = "Mick"
      f20time = 2
      Flasher005.visible = 0
    Case 2:
      F20.state = 1
      Mick_PRIM.image = "Mick_Bright"
      f20time = 1
      Flasher005.visible = 1
  End Select
End Sub

Dim f21time

Sub Sol21(enabled)
  If enabled Then
    F21.state = 1
    Mick_PRIM.image = "Mick_Bright"
    f21time = 1
    Fade21.enabled = 1
    Flasher005.visible = 1
  Else
    F21.state = 0
    Mick_PRIM.image = "Mick"
    Fade21.enabled = 0
    Flasher005.visible = 0
  End If
End Sub

Sub Fade21_Timer()
  Select Case f21time
    Case 1:
      F21.state = 0
      Mick_PRIM.image = "Mick"
      f21time = 2
      Flasher005.visible = 0
    Case 2:
      F21.state = 1
      Mick_PRIM.image = "Mick_Bright"
      f21time = 1
      Flasher005.visible = 1
  End Select
End Sub

Sub Sol31(enabled)
  If enabled Then
    f31.state = 1
    Flasher008.visible = 1
  Else
    f31.state = 0
    Flasher008.visible = 0
  End If
End Sub

Sub Sol26(enabled)
  If enabled Then
    Light26.state = 1
    Light26a.state = 1
    RAMPAR.Image = "amp_right_bright"
    RAMPAL.Image = "amp_left_bright"
    RS_Posters.Image = "Posters_Light"
    Amp_Left.Image = "Marshall_Amp_bright"
    Amp_Right.Image = "Marshall_Amp_bright"
    Flasher006.visible = 1
  Else
    Light26.state = 0
    Light26a.state = 0
    RAMPAR.Image = "amp_right"
    RAMPAL.Image = "amp_left"
    RS_Posters.Image = "Posters_Dark"
    Amp_Left.Image = "Marshall_Amp"
    Amp_Right.Image = "Marshall_Amp"
    Flasher006.visible = 0
  End If
End Sub

Dim f25time

Sub Sol25(enabled)
  If enabled Then
    F25.state = 1
    F25a.state = 1
    F25pf.state = 1
    Mick_PRIM.image = "Mick_Blue"
    f25time = 1
    Fade25.enabled = 1
    Flasher007.visible = 1
  Else
    F25.state = 0
    F25a.state = 0
    F25pf.state = 0
    Mick_PRIM.image = "Mick"
    Fade25.enabled = 0
    Flasher007.visible = 0
  End If
End Sub

Sub Fade25_Timer()
  Select Case f25time
    Case 1:
      F25.state = 0
      F25a.state = 0
      F25pf.state = 0
      Mick_PRIM.image = "Mick"
      f25time = 2
      Flasher005.visible = 0
    Case 2:
      F25.state = 1
      F25a.state = 1
      F25pf.state = 1
      Mick_PRIM.image = "Mick_Blue"
      f25time = 1
      Flasher005.visible = 1
  End Select
End Sub

Dim f27time

Sub Sol27(enabled)
  If enabled Then
    F27.state = 1
    F27a.state = 1
    F27pf.state = 1
    Mick_PRIM.image = "Mick_Red"
    f27time = 1
    Fade27.enabled = 1
    Flasher008.visible = 1
  Else
    F27.state = 0
    F27a.state = 0
    F27pf.state = 0
    Mick_PRIM.image = "Mick"
    Fade27.enabled = 0
    Flasher008.visible = 0
  End If
End Sub

Sub Fade27_Timer()
  Select Case f27time
    Case 1:
      F27.state = 0
      F27a.state = 0
      F27pf.state = 0
      Mick_PRIM.image = "Mick"
      f27time = 2
      Flasher005.visible = 0
    Case 2:
      F27.state = 1
      F27a.state = 1
      F27pf.state = 1
      Mick_PRIM.image = "Mick_Red"
      f27time = 1
      Flasher005.visible = 1
  End Select
End Sub

Dim f29time

Sub Sol29(enabled)
  If enabled Then
    f29.state = 1
    f29a.state = 1
    f29b.state = 1
    f29c.state = 1
    f29d.state = 1
    f29e.state = 1
    f29time = 1
    Fade29.enabled = 1
  Else
    f29.state = 0
    f29a.state = 0
    f29b.state = 0
    f29c.state = 0
    f29d.state = 0
    f29e.state = 0
    Fade29.enabled = 0
  End If
End Sub

Sub Fade29_Timer()
  Select Case f29time
    Case 1:
      f29.state = 0
      f29a.state = 0
      f29b.state = 0
      f29c.state = 0
      f29d.state = 0
      f29e.state = 0
      f29time = 2
    Case 2:
      f29.state = 1
      f29a.state = 1
      f29b.state = 1
      f29c.state = 1
      f29d.state = 1
      f29e.state = 1
      f29time = 1
  End Select
End Sub

Dim f28time

Sub Sol28(enabled)
  If enabled Then
    F28.state = 1
    F28a.state = 1
    F28pf.state = 1
    Mick_PRIM.image = "Mick_Blue"
    f28time = 1
    Fade28.enabled = 1
    Flasher006.visible = 1
  Else
    F28.state = 0
    F28a.state = 0
    F28pf.state = 0
    Mick_PRIM.image = "Mick"
    Fade28.enabled = 0
    Flasher006.visible = 0
  End If
End Sub

Sub Fade28_Timer()
  Select Case f28time
    Case 1:
      F28.state = 0
      F28a.state = 0
      F28pf.state = 0
      Mick_PRIM.image = "Mick"
      f28time = 2
      Flasher005.visible = 0
    Case 2:
      F28.state = 1
      F28a.state = 1
      F28pf.state = 1
      Mick_PRIM.image = "Mick_Blue"
      f28time = 1
      Flasher005.visible = 1
    End Select
End Sub

Dim f22time

Sub Sol22(enabled)
  If enabled Then
    Light22.state = 1
    Fl142p.BlendDisableLighting=1
    Light22a.state = 1
    Charlie.Image = "drummer_bright"
    Keith.Image = "guitar1b_bright"
    Mick_PRIM.image = "Mick_Bright"
    RampAL.image = "amp_left_bright"
    f22time = 1
    Fade22.enabled = 1
    Flasher005.visible = 1
    Flasher006.visible = 1
    Flasher007.visible = 1
    Flasher008.visible = 1
  Else
    Light22.state = 0
    Light22a.state = 0
    Fl142p.BlendDisableLighting=0
    Fade22.enabled = 0
    Charlie.Image = "drummer"
    Keith.Image = "guitar1b"
    Mick_PRIM.image = "Mick"
    RampAL.image = "amp_left"
    Flasher005.visible = 0
    Flasher006.visible = 0
    Flasher007.visible = 0
    Flasher008.visible = 0
  End If
End Sub

Sub Fade22_Timer()
  Select Case f22time
    Case 1:
      Light22.state = 0
      Light22a.state = 0
      Charlie.Image = "drummer"
      Keith.Image = "guitar1b"
      Mick_PRIM.image = "Mick"
      RampAL.image = "amp_left_bright"
      f22time = 2
      Flasher005.visible = 0
      Flasher006.visible = 0
      Flasher007.visible = 0
      Flasher008.visible = 0
    Case 2:
      Light22.state = 1
      Light22a.state = 1
      Charlie.Image = "drummer_bright"
      Keith.Image = "guitar1b_bright"
      Mick_PRIM.image = "Mick_Bright"
      RampAL.image = "amp_left"
      f22time = 1
      Flasher005.visible = 1
      Flasher006.visible = 1
      Flasher007.visible = 1
      Flasher008.visible = 1
  End Select
End Sub

Dim f23time

Sub Sol23(enabled)
  If enabled Then
    Light23.state = 1
    Light23a.state = 1
    Light23a1.state = 1
    f23time = 1
    Fade23.enabled = 1
    Mick_PRIM.image = "Mick_Bright"
    Ronny.image = "Guitar2_Bright"
    RampAR.image = "amp_right_bright"
    Flasher005.visible = 1
  Else
    Light23.state = 0
    Light23a.state = 0
    Light23a1.state = 0
    Fade23.enabled = 0
    Mick_PRIM.image = "Mick"
    Ronny.image = "Guitar2"
    RampAR.image = "amp_right"
    Flasher005.visible = 0
  End If
End Sub

Sub Fade23_Timer()
  Select Case f23time
    Case 1:
      Light23.state = 0
      Light23a.state = 0
      Light23a1.state = 0
      Ronny.image = "Guitar2"
      RampAR.image = "amp_right"
      f23time = 2
      Flasher007.visible = 0
    Case 2:
      Light23.state = 1
      Light23a.state = 1
      Light23a1.state = 1
      f23time = 1
      Ronny.image = "Guitar2_Bright"
      RampAR.image = "amp_right_bright"
      Flasher007.visible = 1
  End Select
End Sub

'*******************************************
'  Ramp Triggers
'*******************************************
Sub ramptrigger01_hit()
  WireRampOn True  'Play Plastic Ramp Sound
  bsRampOnClear    'Shadow on ramp and pf below
End Sub

Sub RRD_Hit()
  WireRampOff  'Turn off the Plastic Ramp Sound
  bsRampOnWire  'Shadow only on pf
  PlaySoundAtVol "Ball_Bounce", RRD, 1
  gi_on
End Sub

Sub ramptrigger02_hit()
  WireRampOn True  'Play Plastic Ramp Sound
  bsRampOnClear    'Shadow on ramp and pf below
End Sub

Sub LRD_Hit()
  WireRampOff  'Turn off the Plastic Ramp Sound
  bsRampOnWire  'Shadow only on pf
  PlaySoundAtVol "Ball_Bounce", LRD, 1
  gi_on
End Sub

' Virtual Spring - The real game has a small spring just to the left of the back entrance to the ball lock.
' If the ball is moving fast enough it will overpower and pass by the spring into the top lanes, if it is moving too slow the spring will
' knock the ball to the right down the back entrance to the ball lock.

Dim virtdir
Sub Virt_Setup_Hit()
  virtdir = 1
  virt_reset.enabled = 1
End Sub

Sub Virt_Reset_Timer()
  virtdir = 0
End Sub

Sub Virt_Spring_Hit()
  If ActiveBall.VelX > -10 And Virtdir = 1 Then
    ActiveBall.VelX = 2
  End If
End Sub

'Manual Trough System (Switches) to handle the White Ceramic Ball
'and activate the White Ball Detection Opto.

Sub SW21_Hit()
  If ActiveBall.id = 666 Then
    Controller.Switch(71) = 0
  End If
  Controller.Switch(21) = 1
  W21.isdropped = 0
End Sub

Sub SW21_Unhit()
  W21.isdropped = 1
  Controller.Switch(21) = 0
End Sub

Sub SW20_Hit()
  W20.isdropped = 0
  Controller.Switch(20) = 1
End Sub

Sub SW20_Unhit()
  W20.isdropped = 1
  Controller.Switch(20) = 0
End Sub

Sub SW19_Hit()
  W19.isdropped = 0
  Controller.Switch(19) = 1
End Sub

Sub SW19_Unhit()
  W19.isdropped = 1
  Controller.Switch(19) = 0
End Sub

Sub SW18_Hit()
  W18.isdropped = 0
  Controller.Switch(18) = 1
End Sub

Sub SW18_Unhit()
  W18.isdropped = 1
  Controller.Switch(18) = 0
End Sub

Sub SW17_Hit()
  'W17.isdropped = 0
  Controller.Switch(17) = 1
End Sub

Sub SW17_Unhit()
  'W17.isdropped = 1
  Controller.Switch(17) = 0
End Sub

'Intial Load Timer to populate the Trough.

Dim wball,tball
tball = 1

Sub load_trough_timer()
  Select Case tball
    Case 1:Kicker_Load.createball:Kicker_Load.kick 45,5:tball = 2
    Case 2:Kicker_Load.createball:Kicker_Load.kick 45,5:tball = 3
    Case 3:Set wball = Kicker_Load.createball:wball.image = "powerball2":wball.id = 666:Kicker_Load.kick 45,5:tball = 4
    Case 4:Kicker_Load.createball:Kicker_Load.kick 45,5:tball = 5
    Case 5:Kicker_Load.createball:Kicker_Load.kick 45,5:tball = 6
    Case 6:me.enabled = 0
  End Select
End Sub

'****************************************************************
' ZGII: GI
'****************************************************************

Dim gilvl   'General Illumination light state tracked for Dynamic Ball Shadows
gilvl = 1

Sub ToggleGI(Enabled)
  Dim xx
  If enabled Then
    For Each xx In GI
      xx.state = 1
    Next
    'PFShadowsGION.visible = 1
    gilvl = 1
  Else
    For Each xx In GI
      xx.state = 0
    Next
    'PFShadowsGION.visible = 0
    GITimer.enabled = True
    gilvl = 0
  End If
  Sound_GI_Relay enabled, bumper1
End Sub

Sub GITimer_Timer()
  Me.enabled = False
  ToggleGI 1
End Sub

'*****GI Lights On/Off Subs

dim rsxx

For each rsxx in GI:rsxx.State = 1: Next

Sub GI_on()
'For each xx in Inserts:xx.Intensity = 7:next
  For each rsxx in GI:rsxx.State = 1: Next
  'DOF 100,1
    LFLogo.image = "flipper-l2"
    RFLogo.image = "flipper-r2"
    Mick_PRIM.image = "Mick_Bright"
    Flasher1.opacity = 60
    Flasher2.opacity = 60
    Flasher3.opacity = 60
    RS_Posters.blenddisablelighting = 0.2
    Wall4.blenddisablelighting = 0.2
    SideCab.blenddisablelighting = 0.2
End Sub

Sub GI_off()
'For each xx in Inserts:xx.Intensity = 15:next
  For each rsxx in GI:rsxx.State = 0: Next
  'DOF 100,0
    LFLogo.image = "flipper-l2d"
    RFLogo.image = "flipper-r2d"
    Mick_PRIM.image = "Mick_Dark"
    Flasher1.opacity=0
    Flasher2.opacity=0
    Flasher3.opacity=0
    RS_Posters.blenddisablelighting=0.1
    Wall4.blenddisablelighting=0.1
    Sidecab.blenddisablelighting=0.1
End Sub

Dim gflash

Sub GI_Flash_Timer()
  Select Case gflash
    Case 1:gi_on:gflash = 2
    Case 2:gi_off:gflash = 1
  End Select
End Sub

Sub Gi_Flash_Stop_Timer()
  gi_flash.enabled = 0
  gi_on
  me.enabled = 0
End Sub

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
  RandomSoundSlingshotRight Sling1
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
    RightSlingShot.TimerInterval = 10
    vpmTimer.PulseSw 27
  'gi1.State = 0:Gi2.State = 0
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  RandomSoundSlingshotLeft Sling2
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
    LeftSlingShot.TimerInterval = 10
    vpmTimer.PulseSw 26
  'gi3.State = 0:Gi4.State = 0
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

' Loads all the stuff for bubble level...

Dim bubble
Set bubble=Kicker003.CreateSizedBall(5)

'bubble.ReflectionEnabled=False
Kicker003.Kick 0, 0
Kicker003.Enabled=False
bubble.Image="bubble"


' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub

Sub VRStartButtonTimer_timer()
  If L01.state = 1 Then
    VRStartButton.disableLighting = 1
    VRStartButton2.disableLighting = 1
  Else
    VRStartButton.disableLighting = 0
    VRStartButton2.disableLighting = 0
  End If
  If L02.state = 1 Then
    VRTourneyButton.disableLighting = 1
  Else
    VRTourneyButton.disableLighting = 0
  End If
End Sub

Sub KeyShakeTimer_Timer()
  'VR_KeyShake.rotz= 0 - staypuft.roty
  'VR_KeyShake.roty= 0 - staypuft.roty/8
End Sub

'*******************************************
' ZTIM:Timers
'*******************************************

Sub GameTimer_Timer() 'The game timer interval; should be 10 ms
  Cor.Update    'update ball tracking (this sometimes goes in the RDampen_Timer sub)
  RollingUpdate   'update rolling sounds
End Sub

Dim FrameTime, InitFrameTime
InitFrameTime = 0
Sub FrameTimer_Timer() 'The frame timer interval should be -1, so executes at the display frame rate
  FrameTime = GameTime - InitFrameTime
  InitFrameTime = GameTime  'Count frametime
  FlipperVisualUpdate    'update flipper shadows and primitives
  If DynamicBallShadowsOn Or AmbientBallShadowOn Then DynamicBSUpdate 'update ball shadows
End Sub

'***************************************************************
' ZSHA: VPW DYNAMIC BALL SHADOWS by Iaakki, Apophis, and Wylte
'***************************************************************

'****** INSTRUCTIONS please read ******

'****** Part A:  Table Elements ******
'
' Import the "bsrtx8" and "ballshadow" images
' Import the shadow materials file (3 sets included) (you can also export the 3 sets from this table to create the same file)
' Copy in the BallShadowA flasher set and the sets of primitives named BallShadow#, RtxBallShadow#, and RtxBall2Shadow#
' * Count from 0 up, with at least as many objects each as there can be balls, including locked balls.  You'll get an "eval" warning if tnob is higher
' * Warning:  If merging with another system (JP's ballrolling), you may need to check tnob math and add an extra BallShadowA# flasher (out of range error)
' Ensure you have a timer with a -1 interval that is always running
' Set plastic ramps DB to *less* than the ambient shadows (-11000) if you want to see the pf shadow through the ramp
' Place triggers at the start of each ramp *type* (solid, clear, wire) and one at the end if it doesn't return to the base pf
' * These can share duties as triggers for RampRolling sounds

' Create a collection called DynamicSources that includes all light sources you want to cast ball shadows
' It's recommended that you be selective in which lights go in this collection, as there are limitations:
' 1. The shadows can "pass through" solid objects and other light sources, so be mindful of where the lights would actually able to cast shadows
' 2. If there are more than two equidistant sources, the shadows can suddenly switch on and off, so places like top and bottom lanes need attention
' 3. At this time the shadows get the light on/off from tracking gilvl, so if you have lights you want shadows for that are on at different times you will need to either:
' a) remove this restriction (shadows think lights are always On)
' b) come up with a custom solution (see TZ example in script)
' After confirming the shadows work in general, use ball control to move around and look for any weird behavior

'****** End Part A:  Table Elements ******


'****** Part B:  Code and Functions ******

' *** Timer sub
' The "DynamicBSUpdate" sub should be called by a timer with an interval of -1 (framerate)
' Example timer sub:

'Sub FrameTimer_Timer()
' If DynamicBallShadowsOn Or AmbientBallShadowOn Then DynamicBSUpdate 'update ball shadows
'End Sub

' *** These are usually defined elsewhere (ballrolling), but activate here if necessary
'Const tnob = 10 ' total number of balls
'Const lob = 0  'locked balls on start; might need some fiddling depending on how your locked balls are done
'Dim tablewidth: tablewidth = Table1.width
'Dim tableheight: tableheight = Table1.height

' *** User Options - Uncomment here or move to top for easy access by players
'----- Shadow Options -----
'Const DynamicBallShadowsOn = 1   '0 = no dynamic ball shadow ("triangles" near slings and such), 1 = enable dynamic ball shadow
'Const AmbientBallShadowOn = 1    '0 = Static shadow under ball ("flasher" image, like JP's)
'                 '1 = Moving ball shadow ("primitive" object, like ninuzzu's) - This is the only one that shows up on the pf when in ramps and fades when close to lights!
'                 '2 = flasher image shadow, but it moves like ninuzzu's

' *** The following segment goes within the RollingUpdate sub, so that if Ambient...=0 and Dynamic...=0 the entire DynamicBSUpdate sub can be skipped for max performance
' ** Change gBOT to BOT if using existing getballs code
' ** Double commented lines commonly found there included for reference:

''  ' stop the sound of deleted balls
''  For b = UBound(gBOT) + 1 to tnob
'   If AmbientBallShadowOn = 0 Then BallShadowA(b).visible = 0
''    ...rolling(b) = False
''    ...StopSound("BallRoll_" & b)
''  Next
''
'' ...rolling and drop sounds...
''
''    If DropCount(b) < 5 Then
''      DropCount(b) = DropCount(b) + 1
''    End If
''
'   ' "Static" Ball Shadows
'   If AmbientBallShadowOn = 0 Then
'     BallShadowA(b).visible = 1
'     BallShadowA(b).X = gBOT(b).X + offsetX
'     If gBOT(b).Z > 30 Then
'       BallShadowA(b).height=gBOT(b).z - BallSize/4 + b/1000 'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping the ramp
'       BallShadowA(b).Y = gBOT(b).Y + offsetY + BallSize/10
'     Else
'       BallShadowA(b).height=gBOT(b).z - BallSize/2 + 1.04 + b/1000
'       BallShadowA(b).Y = gBOT(b).Y + offsetY
'     End If
'   End If

' *** Place this inside the table init, just after trough balls are added to gBOT
'
' Add balls to shadow dictionary
' For Each xx in gBOT
'   bsDict.Add xx.ID, bsNone
' Next

' *** Example RampShadow trigger subs:

'Sub ClearRampStart_hit()
' bsRampOnClear     'Shadow on ramp and pf below
'End Sub

'Sub SolidRampStart_hit()
' bsRampOn        'Shadow on ramp only
'End Sub

'Sub WireRampStart_hit()
' bsRampOnWire      'Shadow only on pf
'End Sub

'Sub RampEnd_hit()
' bsRampOff ActiveBall.ID 'Back to default shadow behavior
'End Sub


'' *** Required Functions, enable these if they are not already present elswhere in your table
'Function max(a,b)
' If a > b Then
'   max = a
' Else
'   max = b
' End If
'End Function

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
Dim objrtx1(8), objrtx2(8)
Dim objBallShadow(8)
Dim OnPF(5)
Dim BallShadowA
BallShadowA = Array (BallShadowA0,BallShadowA1,BallShadowA2,BallShadowA3,BallShadowA4,BallShadowA5,BallShadowA6,BallShadowA7)
Dim DSSources(4), numberofsources', DSGISide(30) 'Adapted for TZ with GI left / GI right

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
  Dim BOT: BOT = getballs 'Uncomment if you're destroying balls - Not recommended! #SaveTheBalls
  If onPlayfield Then
    OnPF(ballNum) = True
    bsRampOff BOT(ballNum).ID
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
  Dim BOT: BOT = getballs 'Uncomment if you're destroying balls - Not recommended! #SaveTheBalls

  'Hide shadow of deleted balls
  For s = UBound(BOT) + 1 To tnob - 1
    objrtx1(s).visible = 0
    objrtx2(s).visible = 0
    objBallShadow(s).visible = 0
    BallShadowA(s).visible = 0
  Next

  If UBound(BOT) < lob Then Exit Sub 'No balls in play, exit

  'The Magic happens now
  For s = lob To UBound(BOT)
    ' *** Normal "ambient light" ball shadow
    'Layered from top to bottom. If you had an upper pf at for example 80 units and ramps even above that, your Elseif segments would be z>110; z<=110 And z>100; z<=100 And z>30; z<=30 And z>20; Else (under 20)

    'Primitive shadow on playfield, flasher shadow in ramps
    If AmbientBallShadowOn = 1 Then
      '** Above the playfield
      If BOT(s).Z > 30 Then
        If OnPF(s) Then BallOnPlayfieldNow False, s   'One-time update
        bsRampType = getBsRampType(BOT(s).id)
        '   debug.print bsRampType

        If Not bsRampType = bsRamp Then 'Primitive visible on PF
          objBallShadow(s).visible = 1
          objBallShadow(s).X = BOT(s).X + (BOT(s).X - (tablewidth / 2)) / (Ballsize / AmbientMovement) + offsetX
          objBallShadow(s).Y = BOT(s).Y + offsetY
          objBallShadow(s).size_x = 5 * ((BOT(s).Z + BallSize) / 80) 'Shadow gets larger and more diffuse as it moves up
          objBallShadow(s).size_y = 4.5 * ((BOT(s).Z + BallSize) / 80)
          UpdateMaterial objBallShadow(s).material,1,0,0,0,0,0,AmbientBSFactor * (30 / (BOT(s).Z)),RGB(0,0,0),0,0,False,True,0,0,0,0
        Else 'Opaque, no primitive below
          objBallShadow(s).visible = 0
        End If

        If bsRampType = bsRampClear Or bsRampType = bsRamp Then 'Flasher visible on opaque ramp
          BallShadowA(s).visible = 1
          BallShadowA(s).X = BOT(s).X + offsetX
          BallShadowA(s).Y = BOT(s).Y + offsetY + BallSize / 10
          BallShadowA(s).height = BOT(s).z - BallSize / 4 + s / 1000 'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping the ramp
          If bsRampType = bsRampClear Then BallShadowA(s).Opacity = 50 * AmbientBSFactor
        ElseIf bsRampType = bsWire Or bsRampType = bsNone Then 'Turn it off on wires or falling out of a ramp
          BallShadowA(s).visible = 0
        End If

        '** On pf, primitive only
      ElseIf BOT(s).Z <= 30 And BOT(s).Z > 20 Then
        If Not OnPF(s) Then BallOnPlayfieldNow True, s
        objBallShadow(s).X = BOT(s).X + (BOT(s).X - (tablewidth / 2)) / (Ballsize / AmbientMovement) + offsetX
        objBallShadow(s).Y = BOT(s).Y + offsetY
        '   objBallShadow(s).Z = gBOT(s).Z + s/1000 + 0.04    'Uncomment (and adjust If/Elseif height logic) if you want the primitive shadow on an upper/split pf

        '** Under pf, flasher shadow only
      Else
        If OnPF(s) Then BallOnPlayfieldNow False, s
        objBallShadow(s).visible = 0
        BallShadowA(s).visible = 1
        BallShadowA(s).X = BOT(s).X + offsetX
        BallShadowA(s).Y = BOT(s).Y + offsetY
        BallShadowA(s).height = BOT(s).z - BallSize / 4 + s / 1000
      End If

      'Flasher shadow everywhere
    ElseIf AmbientBallShadowOn = 2 Then
      If gBOT(s).Z > 30 Then 'In a ramp
        BallShadowA(s).X = BOT(s).X + offsetX
        BallShadowA(s).Y = BOT(s).Y + offsetY + BallSize / 10
        BallShadowA(s).height = BOT(s).z - BallSize / 4 + s / 1000 'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping the ramp
      ElseIf BOT(s).Z <= 30 And BOT(s).Z > 20 Then 'On pf
        BallShadowA(s).visible = 1
        BallShadowA(s).X = BOT(s).X + (gBOT(s).X - (tablewidth / 2)) / (Ballsize / AmbientMovement) + offsetX
        BallShadowA(s).Y = BOT(s).Y + offsetY
        BallShadowA(s).height = 1.04 + s / 1000
      Else 'Under pf
        BallShadowA(s).X = BOT(s).X + offsetX
        BallShadowA(s).Y = BOT(s).Y + offsetY
        BallShadowA(s).height = BOT(s).z - BallSize / 4 + s / 1000
      End If
    End If

    ' *** Dynamic shadows
    If DynamicBallShadowsOn Then
      If BOT(s).Z < 30 And BOT(s).X < 850 Then 'Parameters for where the shadows can show, here they are not visible above the table (no upper pf) or in the plunger lane
        dist1 = falloff
        dist2 = falloff
        For iii = 0 To numberofsources - 1 'Search the 2 nearest influencing lights
          LSd = Distance(BOT(s).x, BOT(s).y, DSSources(iii)(0), DSSources(iii)(1)) 'Calculating the Linear distance to the Source
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
          objrtx1(s).X = BOT(s).X
          objrtx1(s).Y = BOT(s).Y
          '   objrtx1(s).Z = gBOT(s).Z - 25 + s/1000 + 0.01 'Uncomment if you want to add shadows to an upper/lower pf
          objrtx1(s).rotz = AnglePP(DSSources(src1)(0), DSSources(src1)(1), BOT(s).X, BOT(s).Y) + 90
          ShadowOpacity1 = 1 - dist1 / falloff
          objrtx1(s).size_y = Wideness * ShadowOpacity1 + Thinness
          UpdateMaterial objrtx1(s).material,1,0,0,0,0,0,ShadowOpacity1 * DynamicBSFactor ^ 3,RGB(0,0,0),0,0,False,True,0,0,0,0
        Else
          objrtx1(s).visible = 0
        End If
        ShadowOpacity2 = 0
        If dist2 < falloff Then
          objrtx2(s).visible = 1
          objrtx2(s).X = BOT(s).X
          objrtx2(s).Y = BOT(s).Y + offsetY
          '   objrtx2(s).Z = BOT(s).Z - 25 + s/1000 + 0.02 'Uncomment if you want to add shadows to an upper/lower pf
          objrtx2(s).rotz = AnglePP(DSSources(src2)(0), DSSources(src2)(1), BOT(s).X, BOT(s).Y) + 90
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
'****  END VPW DYNAMIC BALL SHADOWS by Iaakki, Apophis, and Wylte
'****************************************************************

'******************************************************
' ZPHY:  GENERAL ADVICE ON PHYSICS
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
' To add the flipper tricks you must
'  - Include a call to FlipperCradleCollision from within OnBallBallCollision subroutine
'  - Include a call the CheckLiveCatch from the LeftFlipper_Collide and RightFlipper_Collide subroutines

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
  Dim BOT
  BOT = GetBalls

  ' stop the sound of deleted balls
  For b = UBound(BOT) + 1 To tnob - 1
    ' Comment the next line if you are not implementing Dyanmic Ball Shadows
    If AmbientBallShadowOn = 0 Then BallShadowA(b).visible = 0
    rolling(b) = False
    StopSound("BallRoll_" & b)
  Next

  ' exit the sub if no balls on the table
  If UBound(BOT) =  - 1 Then Exit Sub

  ' play the rolling sound for each ball
  For b = 0 To UBound(BOT)
    If BallVel(BOT(b)) > 1 And BOT(b).z < 30 Then
      rolling(b) = True
      PlaySound ("BallRoll_" & b), - 1, VolPlayfieldRoll(BOT(b)) * BallRollVolume * VolumeDial, AudioPan(BOT(b)), 0, PitchPlayfieldRoll(BOT(b)), 1, 0, AudioFade(BOT(b))
    Else
      If rolling(b) = True Then
        StopSound("BallRoll_" & b)
        rolling(b) = False
      End If
    End If

    ' Ball Drop Sounds
    If BOT(b).VelZ <  - 1 And BOT(b).z < 55 And BOT(b).z > 27 Then 'height adjust for ball drop sounds
      If DropCount(b) >= 5 Then
        DropCount(b) = 0
        If BOT(b).velz >  - 7 Then
          RandomSoundBallBouncePlayfieldSoft BOT(b)
        Else
          RandomSoundBallBouncePlayfieldHard BOT(b)
        End If
      End If
    End If

    If DropCount(b) < 5 Then
      DropCount(b) = DropCount(b) + 1
    End If

    ' "Static" Ball Shadows
    ' Comment the next If block, if you are not implementing the Dynamic Ball Shadows
    'If AmbientBallShadowOn = 0 Then
    ' If BOT(b).Z > 30 Then
    '   BallShadowA(b).height = BOT(b).z - BallSize / 4   'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping the ramp
    ' Else
    '   BallShadowA(b).height = 0.1
    ' End If
    ' BallShadowA(b).Y = BOT(b).Y + offsetY
    ' BallShadowA(b).X = BOT(b).X + offsetX
    ' BallShadowA(b).visible = 1
    'End If
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

'******************************************************
'****  FLEEP MECHANICAL SOUNDS
'******************************************************

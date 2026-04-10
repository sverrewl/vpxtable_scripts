'Genesis - Gottlieb 1986
'VP912 table by jpsalas
'3D VP10 conversion by nFozzy, DOF additions by arngrim
'Created by ICPjuggla. This is a MOD of nFozzy's MOD of JPSalas - VP9 table
'Version 1.11

Option Explicit
Randomize

Dim FlipperCoilRampupMode

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

' Options
' Volume devided by - lower gets higher sound

Const VolDiv = 10    ' Lower number, louder ballrolling/collition sound
Const VolCol = 1      ' Ball collition divider ( voldiv/volcol )

' The rest of the values are multipliers
'
'  .5 = lower volume
' 1.5 = higher volume

Const VolBump   = 1.5    ' Bumpers volume.
Const VolGates  = 1.5    ' Gates volume.
Const VolMetal  = 1.5    ' Metals volume.
Const VolPo     = 1.5    ' Rubber posts volume.
Const VolPlast  = 1.5    ' Plastics volume.
Const VolWood   = 1.5    ' Woods volume.
Const VolFlip   = 1.5    ' Flipper volume.

Const Ballsize = 52
Const BallMass = 1.5

'***********  Set the Flippers Type *********************************

' FlipperCoilRampupMode
' Flipper coil ramp behavior in-game:
' Either of the following modes may feel more natural
' Depends on various factors such as system specifications, playfield monitor, GPU settings, flipper keys vs. leaf switches

' 0 - Static - flipper coil-ramp up is static; Underpowered systems may need to use this mode.
' 1 - Dynamic - flipper coil-ramp up changes dynamically for a better simulation of tap pass capabilities. Requires a fast system - otherwise may introduce a possible flipper lag

FlipperCoilRampupMode = 1

'*********************************


' ***************************************************************************
'                          DAY & NIGHT FUNCTIONS AND DATASETS
' ****************************************************************************
' PGI (Persistent General illumination)
'-1= 0.5x  emmissive de-block increase
' 0= (default)full emmisive blocking
' 1= Half of full
' 2= Quarter of full
' 3= 8'th of full
' 4= 16'th of full
' 5= 32'nd of full
' 6 = full emmision

Const GILevel = 5
Const FLLevel = 3
Const BTLevel = 3

Const cUSESHADOW = 1
Const cUSECOLORGRADE = 1
Const cUSECCUPBOOST = 1
Const cUSELITEBOOST = 0
Const cUSEBACKGLASS = 1
Const cDYNPINBALL = 2 ' 1:(inverted) easy to play light to Dark, 2: (normal) ball dark to light matches table
Const cDYNGAMEBLADES =0
Const cSHADOWBLADES =0
Const cUSEBGLASSHIGH =0
Const cUSETOPPER =0
Const cUSEBACKLIGHT =1
Const cUSEBGLASSCCX =0
Const cUSEBGLASSREFL =0
Const cUSETRANSLIST =1 ' manage transparent light to dark object list
Const cUSEOPAQUELIST =1 ' 0=OFF, 1=ON, 2=switched(ON/OFF) opaque transparent light to dark object list
Const cUSESEMITRANSLIST =1 ' manage semi-transparent light to dark object list


Dim nxx, DNS
DNS = Genesis.NightDay

Dim DivLevel: DivLevel = 35
Dim DNSVal: DNSVal = Round(DNS/10)
Dim DNShift: DNShift = 1

'Dim OPSValues: OPSValues=Array (100,50,20,10 ,5,4 ,3 ,2 ,1, 0,0)
'Dim DNSValues: DNSValues=Array (1.0,0.5,0.1,0.05,0.01,0.005,0.001,0.0005,0.0001, 0.00005,0.00000)
Dim OPSValues: OPSValues=Array (100,50,20,10 ,9,8 ,7 ,6 ,5, 4,3,2,1,0)
Dim OPS2: OPS2  =Array (1.0,0.95,0.90,0.85,0.80,0.75,0.70,0.65,0.60,0.55,0.50,0.50)
Dim OPSValues2: OPSValues2=Array (100*OPS2(DNSVal),50*OPS2(DNSVal),20*OPS2(DNSVal),10*OPS2(DNSVal),9*OPS2(DNSVal),8*OPS2(DNSVal),7*OPS2(DNSVal),6*OPS2(DNSVal),5*OPS2(DNSVal), 4*OPS2(DNSVal),3*OPS2(DNSVal),2*OPS2(DNSVal),1*OPS2(DNSVal),0*OPS2(DNSVal))
Dim OPS3: OPS3 =Array (1.0,0.93,0.85,0.78,0.70,0.65,0.55,0.48,0.40,0.33,0.25,0.25)
Dim OPSValues3: OPSValues3=Array (100*OPS3(DNSVal),50*OPS3(DNSVal),20*OPS3(DNSVal),10*OPS3(DNSVal),9*OPS3(DNSVal),8*OPS3(DNSVal),7*OPS3(DNSVal),6*OPS3(DNSVal),5*OPS3(DNSVal),4*OPS3(DNSVal),3*OPS3(DNSVal),2*OPS3(DNSVal),1*OPS3(DNSVal),0*OPS3(DNSVal))
Dim OPS4: OPS4 =Array (1.0,0.91,0.82,0.74,0.65,0.56,0.47,0.38,0.30,0.21,0.12,0.12)
Dim OPSValues4: OPSValues4=Array (100*OPS4(DNSVal),50*OPS4(DNSVal),20*OPS4(DNSVal),10*OPS4(DNSVal),9*OPS4(DNSVal),8*OPS4(DNSVal),7*OPS4(DNSVal),6*OPS4(DNSVal),5*OPS4(DNSVal),4*OPS4(DNSVal),3*OPS4(DNSVal),2*OPS4(DNSVal),1*OPS4(DNSVal),0*OPS4(DNSVal))
Dim OPS5: OPS5 =Array (1.0,0.91,0.83,0.72,0.63,0.55,0.44,0.35,0.25,0.16,0.07,0.07)
Dim OPSValues5: OPSValues5=Array (100*OPS5(DNSVal),50*OPS5(DNSVal),20*OPS5(DNSVal),10*OPS5(DNSVal),9*OPS5(DNSVal),8*OPS5(DNSVal),7*OPS5(DNSVal),6*OPS5(DNSVal),5*OPS5(DNSVal),4*OPS5(DNSVal),3*OPS5(DNSVal),2*OPS5(DNSVal),1*OPS5(DNSVal),0*OPS5(DNSVal))
Dim OPS6: OPS6 =Array (1.0,0.90,0.81,0.71,0.62,0.52,0.42,0.33,0.23,0.14,0.04,0.04)
Dim OPSValues6: OPSValues6=Array (100*OPS6(DNSVal),50*OPS6(DNSVal),20*OPS6(DNSVal),10*OPS5(DNSVal),9*OPS6(DNSVal),8*OPS6(DNSVal),7*OPS6(DNSVal),6*OPS6(DNSVal),5*OPS6(DNSVal),4*OPS6(DNSVal),3*OPS6(DNSVal),2*OPS6(DNSVal),1*OPS6(DNSVal),0*OPS6(DNSVal))

Dim OPS1: OPS1 = 0.01 '0.25
Dim OPSValues1: OPSValues1=Array (100*OPS1,50*OPS1,20*OPS1,10*OPS1,9*OPS1,8*OPS1,7*OPS1,6*OPS1,5*OPS1,4*OPS1,3*OPS1,2*OPS1,1*OPS1,0*OPS1)

Dim DNSValues: DNSValues=Array (1.0,0.5,0.1,0.05,0.01,0.005,0.001,0.0005,0.0001, 0.00005,0.00001, 0.000005, 0.000001)
Dim SysDNSVal: SysDNSVal=Array (1.0,0.9,0.8,0.7,0.6,0.5,0.5,0.5,0.5, 0.5,0.5)
Dim DivValues: DivValues =Array (1,2,4,8,16,32,32,32,32, 32,32)
Dim DivValues2: DivValues2 =Array (1,1.5,2,2.5,3,3.5,4,4.5,5, 5.5,6)
Dim DivValues3: DivValues3 =Array (1,1.1,1.2,1.3,1.4,1.5,1.6,1.7,1.8,1.9,2.0,2.1)

Dim DIV4: DIV4 =Array (1.0,0.95,0.90,0.85,0.80,0.75,0.70,0.65,0.60,0.55,0.50,0.50)
Dim DivValues4: DivValues4 =Array (1*DIV4(DNSVal),1.1*DIV4(DNSVal),1.2*DIV4(DNSVal),1.3*DIV4(DNSVal),1.4*DIV4(DNSVal),1.5*DIV4(DNSVal),1.6*DIV4(DNSVal),1.7*DIV4(DNSVal),1.8*DIV4(DNSVal),1.9*DIV4(DNSVal),2.0*DIV4(DNSVal),2.1*DIV4(DNSVal))
Dim DIV5: DIV5 =Array (1.0,0.93,0.85,0.78,0.70,0.65,0.55,0.48,0.40,0.33,0.25,0.25)
Dim DivValues5: DivValues5 =Array (1*DIV5(DNSVal),1.1*DIV5(DNSVal),1.2*DIV5(DNSVal),1.3*DIV5(DNSVal),1.4*DIV5(DNSVal),1.5*DIV5(DNSVal),1.6*DIV5(DNSVal),1.7*DIV5(DNSVal),1.8*DIV5(DNSVal),1.9*DIV5(DNSVal),2.0*DIV5(DNSVal),2.1*DIV5(DNSVal))
Dim DIV6: DIV6 =Array (1.0,0.91,0.82,0.74,0.65,0.56,0.47,0.38,0.30,0.21,0.12,0.12)
Dim DivValues6: DivValues6 =Array (1*DIV6(DNSVal),1.1*DIV6(DNSVal),1.2*DIV6(DNSVal),1.3*DIV6(DNSVal),1.4*DIV6(DNSVal),1.5*DIV6(DNSVal),1.6*DIV6(DNSVal),1.7*DIV6(DNSVal),1.8*DIV6(DNSVal),1.9*DIV6(DNSVal),2.0*DIV6(DNSVal),2.1*DIV6(DNSVal))
Dim DIV7: DIV7 =Array (1.0,0.91,0.83,0.72,0.63,0.55,0.44,0.35,0.25,0.16,0.07,0.07)
Dim DivValues7: DivValues7 =Array (1*DIV7(DNSVal),1.1*DIV7(DNSVal),1.2*DIV7(DNSVal),1.3*DIV7(DNSVal),1.4*DIV7(DNSVal),1.5*DIV7(DNSVal),1.6*DIV7(DNSVal),1.7*DIV7(DNSVal),1.8*DIV7(DNSVal),1.9*DIV7(DNSVal),2.0*DIV7(DNSVal),2.1*DIV7(DNSVal))
Dim DIV8: DIV8 =Array (1.0,0.90,0.81,0.71,0.62,0.52,0.42,0.33,0.23,0.14,0.04,0.04)
Dim DivValues8: DivValues8 =Array (1*DIV8(DNSVal),1.1*DIV8(DNSVal),1.2*DIV8(DNSVal),1.3*DIV8(DNSVal),1.4*DIV8(DNSVal),1.5*DIV8(DNSVal),1.6*DIV8(DNSVal),1.7*DIV8(DNSVal),1.8*DIV8(DNSVal),1.9*DIV8(DNSVal),2.0*DIV8(DNSVal),2.1*DIV8(DNSVal))

Dim MUL1: MUL1 =Array (1.0,2.0,3.0,4.0,5.0,6.0,7.0,8.0,9.0,10.0,11.0)' 1.1
Dim MulValues1: MulValues1 =Array (1*MUL1(DNSVal),1.1*MUL1(DNSVal),1.2*MUL1(DNSVal),1.3*MUL1(DNSVal),1.4*MUL1(DNSVal),1.5*MUL1(DNSVal),1.6*MUL1(DNSVal),1.7*MUL1(DNSVal),1.8*MUL1(DNSVal),1.9*MUL1(DNSVal),2.0*MUL1(DNSVal),2.1*MUL1(DNSVal))


Dim RValUP: RValUP=Array (30,60,90,120,150,180,210,240,255,255,255,255)
Dim GValUP: GValUP=Array (30,60,90,120,150,180,210,240,255,255,255,255)
Dim BValUP: BValUP=Array (30,60,90,120,150,180,210,240,255,255,255,255)
Dim RValDN: RValDN=Array (255,210,180,150,120,90,60,30,10,10,10)
Dim GValDN: GValDN=Array (255,210,180,150,120,90,60,30,10,10,10)
Dim BValDN: BValDN=Array (255,210,180,150,120,90,60,30,10,10,10)
Dim FValUP: FValUP=Array (35,40,45,50,55,60,65,70,75,80,85,90,95,100,105)
Dim FValDN: FValDN=Array (100,85,80,75,70,65,60,55,50,45,40,35,30)
Dim MVSAdd: MVSAdd=Array (0.9,0.9,0.8,0.8,0.7,0.7,0.6,0.6,0.5,0.5,0.4,0.3,0.2,0.1)
Dim ReflDN: ReflDN=Array (60,55,50,45,40,35,30,28,26,24,22,20,19,18,16,15,14,13,12,11,10)
Dim DarkUP: DarkUP=Array (1,2,3,4,5,6,6,6,6,6,6,6,6,6,6,6,6,6)

' PLAYFIELD GENERAL OPERATIONAL and LOCALALIZED GI ILLUMINATION
Dim aAllFlashers: aAllFlashers=Array(BGFlasher0,BGFlasher001,BGFlasher002,BGFlasher003,BGFlasher004,BGFlasher005, _
BGFlasher006,BGFlasher007,BGFlasher008,BGFlasher009,BGFlasher010,BGFlasher011,BGFlasher012,BGFlasher013,BGFlasher014)
Dim aGiLights: aGiLights=array(gi5,gi6,gi7,gi8,gi9,gi10,gi11,gi12,gi13,gi14,gi14_1,gi14_2,gi14_3,gi15)
Dim BloomLights: BloomLights=array(gi_ambient,fl1,fl2,fl3,fr1,fr2,fr3,gi1,gi2,gi3,gi4,_
Rl1,Rl2,Rl3,Rl4,Rl5,ll1,ll2,ll3,ll4,ll5,cl1,cl2,cl3,cl4,cl5)


Dim TargetDropGi: TargetDropGi = array()

' PLAYFIELD GLOBAL INTENSITY ILLUMINATION FLASHERS
Flasher1.opacity = OPSValues(DNSVal + DNShift) / DivValues(DNSVal)
Flasher1.intensityscale = DNSValues(DNSVal + DNShift) /DivValues(DNSVal)
Flasher2.opacity = OPSValues(DNSVal + DNShift) /DivValues(DNSVal)
Flasher2.intensityscale = DNSValues(DNSVal + DNShift) /DivValues(DNSVal)
Flasher3.opacity = OPSValues(DNSVal + DNShift) /DivValues(DNSVal)
Flasher3.intensityscale = DNSValues(DNSVal + DNShift) /DivValues(DNSVal)
'Flasher4.opacity = OPSValues(DNSVal + DNShift) /DivValues(DNSVal)
'Flasher4.intensityscale = DNSValues(DNSVal + DNShift) /DivValues(DNSVal)

if cUSELITEBOOST then
Dim LiteUP: LiteUP=Array (10,6,5,4,4,4,4,3,3,3,2,2,1,0.5,0.4,0.3,0.2,0.1)
Flasher5.Color = RGB(RValUP(DNSVal)/LiteUP(DNSVal),GValUP(DNSVal)/LiteUP(DNSVal),BValUP(DNSVal)/LiteUP(DNSVal))
end if 'cUSELITEBOOST

'BACKBOX & BACKGLASS ILLUMINATION
if cUSEBACKGLASS then
BGDark.ModulateVsAdd = MVSAdd(DNSVal)
BGDark.Color = RGB(RValUP(DNSVal),GValUP(DNSVal),BValUP(DNSVal))
BGDark.Amount = FValUP(DNSVal) / DivValues(DNSVal)
BGHigh.Color = RGB(RValDN(DNSVal),GValDN(DNSVal),BValDN(DNSVal))
BGHigh.Amount = FValDN(DNSVal)  / DivValues(DNSVal)
BGHigh1.Color = RGB(RValDN(DNSVal),GValDN(DNSVal),BValDN(DNSVal))
BGHigh1.Amount = FValDN(DNSVal) / DivValues(DNSVal)

if cUSETOPPER > 0 then ' add (TopHigh,TopHigh1,TopHigh2,) to aAllFlashers array, or enable code for to adjust brightness
TopDark.ModulateVsAdd = MVSAdd(DNSVal)
TopDark.Color = RGB(RValUP(DNSVal)-25,GValUP(DNSVal)-25,BValUP(DNSVal)-25)
'TopHigh.Color = RGB(RValDN(DNSVal),GValDN(DNSVal),BValDN(DNSVal))

if cUSETOPPER > 1 then
TopHigh.amount = TopHigh.amount / DivValues3(DNSVal)
TopHigh.opacity = TopHigh.opacity * OPSValues(DNSVal) / DivLevel
if cUSETOPPER > 2 then
TopHigh1.amount = TopHigh1.amount / DivValues3(DNSVal)
TopHigh1.opacity = TopHigh1.opacity * OPSValues(DNSVal) / DivLevel
if cUSETOPPER > 3 then
TopHigh2.amount = TopHigh2.amount / DivValues3(DNSVal)
TopHigh2.opacity = TopHigh2.opacity * OPSValues(DNSVal) / DivLevel
end if 'cUSETOPPER 3
end if 'cUSETOPPER 2
end if 'cUSETOPPER 1
end if 'cUSETOPPER 0

if cUSEBACKLIGHT then
Const nBACKBoost = 3
FloodLightLeft.Color = RGB(RValDN(DNSVal),GValDN(DNSVal),BValDN(DNSVal))
FloodLightLeft.Amount = FValDN(DNSVal)  * nBACKBoost
'FloodLightLeft.opacity = OPSValues(DNSVal+ DNShift) /DivValues(DNSVal)
FloodLightRight.Color = RGB(RValDN(DNSVal),GValDN(DNSVal),BValDN(DNSVal))
FloodLightRight.Amount = FValDN(DNSVal) * nBACKBoost
'FloodLightRight.opacity = OPSValues(DNSVal+ DNShift) /DivValues(DNSVal)
end if 'cUSEBACKLIGHT

BGframe.ModulateVsAdd = MVSAdd(DNSVal) * 0.2
BGframe.Color = RGB(RValUP(DNSVal),GValUP(DNSVal),BValUP(DNSVal))
BGframe.Amount = FValUP(DNSVal) * DivValues(DNSVal)

' Use on arcade backgrounds
'Dim DarkDN: DarkDN=Array (1.5,1.0,0.8,0.75,0.7,0.68,0.66,0.64,0.62,0.60,0.59,0.58,0.57,0.56,0.55,0.54,0.53,0.52)

'Flasher12.opacity = OPSValues(DNSVal)
'Flasher12.intensityscale = DarkDN(DNSVal) * 4.0
'Flasher12.Color = RGB(RValUP(DNSVal) / 2.0,GValUP(DNSVal)/ 2.0,BValUP(DNSVal)/ 2.0)
'Flasher19.opacity = OPSValues(DNSVal)
'Flasher19.intensityscale = DarkDN(DNSVal ) * 4.0
'Flasher19.Color = RGB(RValUP(DNSVal)/ 2.0,GValUP(DNSVal)/ 2.0,BValUP(DNSVal)/ 2.0)

BGHigh.intensityscale = 0.6
BGHigh1.intensityscale = 0.6
BGDark.intensityscale = 0.5
BGframe.intensityscale = 0.4
bggrill.intensityscale = 0.3

'Add BGHigh2 to aAllFlashers array if not using BGFrameMaskFill1 & BGFrameMaskFill5 in cUSEBGLASSREFL otherwise leave it out.
BGHigh2.intensityscale = 1.8

'special case for EM backglass, must set BGDark to 1.0
if 0 then
Dim EMBiasValues: EMBiasValues=Array (800,300,100,-100,300,-600,-900,-1200,-1600,-1800,-1900,-2000)
Dim EMOPSValues: EMOPSValues=Array (100,90,80,79 ,78,76 ,75 ,74 ,73, 72,71,70,70,70)

BGFrameMaskFill.opacity =EMOPSValues(DNSVal)'
BGFrameMaskFill.DepthBias = EMBiasValues(DNSVal)' * 100.0
end if

'special case for EM/SS playfield plastics, see: voltan for settings
'FlPlasticsMask,World, Filter:Overlay, AB:Off,Amt:1,Op:60,Bias:-10000,Mod:0.9:RGB:255,255,255
'Plastics,World, Filter:Screen , AB:Off, Amt:100,Op:100,Bias:-10000,Mod:1.0 RGB:255,255,255
'Lights Falloff:100,FPower:5,Intensity:80, Clr:63x3,Full:232x3,Bias:-10000,enable=1,Height:60,Mod:1,Trans:0.5
'shadows,World, Filter:screen, AB:on, Amt:400, Op:800, Bias:-10000, Mod:1, RGB:RGB:255,255,255
'NOTE: 'Plastics mask' must be below 'plastics' in drawing order.
if 0 then
Const OPM=0.9
Const OPS=1.0
Dim EMOPMValues
Dim EMOPShValues

if Genesis.ShowFSS = true then
EMOPMValues=Array (97,97 ,95,85 ,80 ,75 ,73, 72,71,70,70,70)
EMOPShValues=Array (95,95 ,85,85 ,80 ,80 ,75, 75,70,70,70,70)

FlPlasticsMask.opacity =EMOPMValues(DNSVal) * OPM
shadows.opacity =EMOPShValues(DNSVal) * OPS ' bord's style rendered lights/shadows

Else
EMOPMValues=Array (80,79 ,78,76 ,75 ,74 ,73, 72,71,70,70,70)
EMOPShValues=Array (800,650 ,500,400 ,300 ,200 ,100, 90,80,70,70,70)

FlPlasticsMask.opacity =EMOPMValues(DNSVal) * (OPM + 0.2)
shadows.opacity =EMOPShValues(DNSVal) * OPS ' bord's style rendered lights/shadows
end if
end if

if cUSEBGLASSHIGH then ' GBGHigh3 is _defhorz flasher, Additive, AMT:100, OP:2, DB:8000, MOD:1

BGHigh3.intensityscale = 0.2
Dim ReflUP: ReflUP=Array (0.3,0.4,0.5,0.55,0.6,0.65,0.7,0.75,0.8,0.85,0.9)
BGDark.intensityscale = ReflUP(DNSVal)

end if 'cUSEBGLASSHIGH

Dim nUSEBGLASSREFL: nUSEBGLASSREFL =Array(1,1,1,0,1,0)
'Dim nUSEBGLASSREFL  Array=(0,0,0,0,0,0)

Dim tMaskFill: tMaskFill = Array(BGFrameMaskFill1,BGFrameMaskFill2,BGFrameMaskFill3,Empty,BGFrameMaskFill5,Empty)
' Fill with: BGFrameMaskFill1,BGFrameMaskFill2,BGFrameMaskFill3,BGFrameMaskFill4,BGFrameMaskFill5,PFRearBoxFill1
if cUSEBGLASSREFL > 0 then

Dim ReflFill: ReflFill=Array (10,10,20,20,30,30,30,30,40,40,40)
Dim ReflImg: ReflImg=Array (160,200,300,400,500,600,700,800,900,1000,1000)
Dim FillMag: FillMag=Array (0.8,0.8,0.7,0.7,0.6,0.6,0.55,0.55,0.54,0.54,0.53)

'playfield glass reflections
If nUSEBGLASSREFL(1) = 1 and nUSEBGLASSREFL(2) = 1 then
if NOT IsEmpty(tMaskFill(1)) then: tMaskFill(1).opacity= ReflFill(DNSVal)*2.5: end if' IMG:full table, def glass trans, FLTR:screen,AMT:900,OP:50, AB=off, RGB=0,0,32, DB=-10000,MOD=1
if NOT IsEmpty(tMaskFill(2)) then: tMaskFill(2).opacity= ReflFill(DNSVal)*3.5: end if' IMG:top half, def glass trans, FLTR:screen,AMT:900,OP:50, AB=off, RGB=0,0,32, DB=-10000,MOD=1
  If nUSEBGLASSREFL(3) = 1 then
  if NOT IsEmpty(tMaskFill(3)) then: tMaskFill(3).opacity= ReflImg(DNSVal): end if' IMG:backglass,   FLTR:screen,AMT:300,OP:1000,AB=on, RBG=16,16,32, DB=-1000,MOD=0.3
  else
  if NOT IsEmpty(tMaskFill(3)) then: tMaskFill(3).visible =0: end if
  end if
if NOT IsEmpty(tMaskFill(1)) then: tMaskFill(1).intensityscale= 0.5: end if
if NOT IsEmpty(tMaskFill(2)) then: tMaskFill(2).intensityscale= 0.5: end if
Else
  if NOT IsEmpty(tMaskFill(1)) then: tMaskFill(1).visible =0: end if
  if NOT IsEmpty(tMaskFill(2)) then: tMaskFill(2).visible =0: end if
end if

If nUSEBGLASSREFL(0) = 1 OR nUSEBGLASSREFL(4) = 1 then
'backbox glass reflections
  If nUSEBGLASSREFL(0) = 1 then
  if NOT IsEmpty(tMaskFill(0)) then: tMaskFill(0).opacity= ReflFill(DNSVal) +40: end if ' IMG:black mask fill, FLTR:screen,AMT:500,OP:90, AB=off, RGB=0,0,32, DB=-10000,MOD=1
  Else
  if NOT IsEmpty(tMaskFill(0)) then: tMaskFill(0).visible =0: end if
  end if
  'background reflective image usually __default_screen_space_reflection.png
  If nUSEBGLASSREFL(4) = 1 then
  if NOT IsEmpty(tMaskFill(4)) then: tMaskFill(4).opacity= ReflImg(DNSVal) / 3 : end if' IMG:Def SS Refl, FLTR:screen,AMT:100,OP:1000,AB=on, RBG=10,10,10, DB=-10000,MOD=0.5
  Else
  if NOT IsEmpty(tMaskFill(4)) then: tMaskFill(4).visible =0: end if
  end if
else
if NOT IsEmpty(tMaskFill(0)) then: tMaskFill(0).visible =0: end if
if NOT IsEmpty(tMaskFill(4)) then: tMaskFill(4).visible =0: end if
end if

If nUSEBGLASSREFL(5) = 1 then
' playfield rear EM/SS vertical rececessed shadow
if NOT IsEmpty(tMaskFill(5)) then: tMaskFill(5).opacity= ReflFill(DNSVal)*4: end if' IMG:full table, def glass trans, FLTR:screen,AMT:900,OP:50, AB=off, RGB=0,0,32, DB=-10000,MOD=1
if NOT IsEmpty(tMaskFill(5)) then: tMaskFill(5).intensityscale= 0.5: end if
Else
if NOT IsEmpty(tMaskFill(5)) then: tMaskFill(5).visible =0: end if
end If

Else

if NOT IsEmpty(tMaskFill(0)) then: tMaskFill(0).visible =0: end if
if NOT IsEmpty(tMaskFill(1)) then: tMaskFill(1).visible =0: end if
if NOT IsEmpty(tMaskFill(2)) then: tMaskFill(2).visible =0: end if
if NOT IsEmpty(tMaskFill(3)) then: tMaskFill(3).visible =0: end if
if NOT IsEmpty(tMaskFill(4)) then: tMaskFill(4).visible =0: end if
if NOT IsEmpty(tMaskFill(5)) then: tMaskFill(5).visible =0: end if
end if 'cUSEBGLASSREFL

end If ' cUSEBACKGLASS


' shadow code.
'PF=flasher(Screen,AMT:1,OP:50,DB:0,MOD:1,RGB=255)
'PS=flasher(Screen,AMT:1700,OP:30,DB:-10000,MOD:1,RGB=7,7,7)
'DS =flasher (Screen ,AMT:1,OP:80,DB:0,MOD:1,RGB=255)
Dim tShadows: tShadows = Array(OCPFShadow,OCPFShadow1,Empty,Empty, OCPFDarkShadow)
'Dim tShadows: tShadows = Array(OCPFShadow,OCPFShadow1,OCPFShadow2,OCPFShadow3,OCPFDarkShadow )
if cUSESHADOW then
'LOWER PF Shadows
Const cPFShDiv = 0.95 ' higher = less opaque, smaller= more opaque
Dim PFShDiv: PFShDiv=      Array (1.39*cPFShDiv,1.36*cPFShDiv,1.33*cPFShDiv,1.29*cPFShDiv,1.26*cPFShDiv,1.23*cPFShDiv,1.19*cPFShDiv,1.16*cPFShDiv,1.13*cPFShDiv,1.1*cPFShDiv,1.0*cPFShDiv)
Dim PFShValues: PFShValues=Array (50,80,90,91,92,93,94,95,96,97,98,99,100,101,101,101,101,101)
Dim PFDrkShValues: PFDrkShValues=Array (100,99,90,80,70,60,50,40,30,20,10,5,0)
Dim PFLitShValues: PFLitShValues=Array (5,10,20,30,40,50,60,70,80,90,100)
Dim PFLitShSlopes: PFLitShSlopes=Array ( 3.5,3.5,3.0,2.6,2.2,2.0,1.8,1.6,1.4,1.2,1.0)
Dim PFDrkShSlopes: PFDrkShSlopes=Array (3.0,2.8,2.6,2.4,2.2,2.0,1.8,1.6,1.4,1.2,1.0)

'transparent plastic edging (the edging must be on seperate flasher)
'OFF state usage:
'FlPlasticsEdge.opacity = Round(PFLitShValues(DNSVal) * PFDrkShSlopes(DNSVal))
'ON state usage:
'FlPlasticsEdge.opacity = Round(PFLitShValues(DNSVal) * PFLitShSlopes(DNSVal))


  if NOT IsEmpty(tShadows(0)) then
  tShadows(0).visible =1
  tShadows(0).opacity = PFShValues(DNSVal) / PFShDiv(DNSVal) ' set height to ~0.5 just above PF
  end if
  if NOT IsEmpty(tShadows(1)) then
  tShadows(1).visible =1
  tShadows(1).opacity = PFShValues(DNSVal) / 1.9 ' set height to ~60 just above plastics
  end if
' REVERSE Shadow
  if NOT IsEmpty(tShadows(4)) then
  tShadows(4).visible =1
  if Genesis.ShowFSS = true then
  'tShadows(4).opacity = Round(PFDrkShValues(DNSVal) /1.05)' (using bords lights/shadows) set height to ~60 just above plastics
  tShadows(4).opacity = Round(PFDrkShValues(DNSVal) /1.2)' (default)set height to ~60 just above plastics
  else 'if  Table1.ShowFS = true then
  tShadows(4).opacity = Round(PFDrkShValues(DNSVal) /1.3)' set height to ~60 just above plastics
  end if
  end if

' UPPER PF Shadows
Const cPFShDiv2 = 1.3 ' higher = less opaque, smaller= more opaque
Dim PFShDiv2: PFShDiv2=      Array (1.39*cPFShDiv2,1.36*cPFShDiv2,1.33*cPFShDiv2,1.29*cPFShDiv2,1.26*cPFShDiv2,1.23*cPFShDiv2,1.19*cPFShDiv2,1.16*cPFShDiv2,1.13*cPFShDiv2,1.1*cPFShDiv2,1.0*cPFShDiv2)
Dim PFShValues2: PFShValues2=Array (50,80,90,91,92,93,94,95,96,97,98,99,100,101,101,101,101,101)
  if NOT IsEmpty(tShadows(2)) then
  tShadows(2).visible =1
  tShadows(2).opacity = PFShValues2(DNSVal) / PFShDiv2(DNSVal) ' height ~60
  end if
  if NOT IsEmpty(tShadows(3)) then
  tShadows(3).visible =1
  tShadows(3).opacity = PFShValues2(DNSVal) / 1.7 ' height ~110
  end if

else
  if NOT IsEmpty(tShadows(0)) then: tShadows(0).visible =0: end if
  if NOT IsEmpty(tShadows(1)) then: tShadows(1).visible =0: end if
  if NOT IsEmpty(tShadows(2)) then: tShadows(2).visible =0: end if
  if NOT IsEmpty(tShadows(3)) then: tShadows(3).visible =0: end if
  if NOT IsEmpty(tShadows(4)) then: tShadows(4).visible =0: end if
end If 'cUSESHADOW

if cSHADOWBLADES then
' import shadowblade materials.mat to work, duplicate or import _default_gameblades and scale apropriately
Dim ShadeBladeVals: ShadeBladeVals=Array ("Shadow with an image0","Shadow with an image1",_
"Shadow with an image2","Shadow with an image3", "Shadow with an image3",_
"Shadow with an image4","Shadow with an image4",_
"Shadow with an image5","Shadow with an image5","Shadow with an image5","Shadow with an image5","Shadow with an image5")

ShadowBlades.material = ShadeBladeVals(DNSVal)
end if 'cSHADOWBLADES

Dim BlueVals: BlueVals=Array ("bluelist9","bluelist9",_
"bluelist8","bluelist8", "bluelist7","bluelist7", _
"bluelist6","bluelist6", "bluelist5","bluelist4", _
"bluelist4","bluelist4","bluelist4","bluelist4")
Dim RedVals: RedVals=Array ("redlist9","redlist9",_
"redlist8","redlist8", "redlist7","redlist7", _
"redlist6","redlist6", "redlist5","redlist4", _
"redlist4","redlist4","redlist4","redlist4")
Dim YellowVals: YellowVals=Array ("yellowlist9","yellowlist9",_
"yellowlist8","yellowlist8", "yellowlist7","yellowlist7", _
"yellowlist6","yellowlist6", "yellowlist5","yellowlist4", _
"yellowlist4","yellowlist4","yellowlist4","yellowlist4")
Dim OrangeVals: OrangeVals=Array ("orangelist9","orangelist9",_
"orangelist8","orangelist8", "orangelist7","orangelist7", _
"orangelist6","orangelist6", "orangelist5","orangelist4", _
"orangelist4","orangelist4","orangelist4","orangelist4")
Dim WhiteVals: WhiteVals=Array ("whitelist9","whitelist9",_
"whitelist8","whitelist8", "whitelist7","whitelist7", _
"whitelist6","whitelist6", "whitelist5","whitelist4", _
"whitelist4","whitelist4","whitelist4","whitelist4")
Dim UserVals: UserVals=Array ("violetlist9","violetlist9",_
"violetlist8","violetlist8", "violetlist7","violetlist7", _
"violetlist6","violetlist6", "violetlist5","violetlist4", _
"violetlist4","violetlist4","violetlist4","violetlist4")
Dim GrayVals: GrayVals=Array ("graylist9","graylist9",_
"graylist8","graylist8", "graylist7","graylist7", _
"graylist6","graylist6", "graylist5","graylist4", _
"graylist4","graylist4","graylist4","graylist4")
Dim GreenVals: GreenVals=Array ("greenlist9","greenlist9",_
"greenlist8","greenlist8", "greenlist7","greenlist7", _
"greenlist6","greenlist6", "greenlist5","greenlist4", _
"greenlist4","greenlist4","greenlist4","greenlist4")
Dim BlackVals: BlackVals=Array ("blacklist9","blacklist9",_
"blacklist8","blacklist8", "blacklist7","blacklist7", _
"blacklist6","blacklist6", "blacklist5","blacklist4", _
"blacklist4","blacklist4","blacklist4","blacklist4")

Dim OpaqueBias: OpaqueBias=Array (20,10,0,-10,-20,-30,-40, -50,-60, -70,-80,-90)



' cUSETRANSLIST & cUSEOPAQUELIST: fill arrays with elements you wish to control the ambient light
Dim TransList: TransList=Array (Empty,Empty,Empty)
Dim OpaqueList: OpaqueList=Array (LFLogo,RFLogo,apron,Ramp10,Ramp37,Bumper1,Bumper2,Bumper3,Bumper4,Empty)
Dim SemiTransList: SemiTransList=Array (Primitive1,Primitive2,Primitive3, Primitive4,dt1,dt2,dt3, Empty)

Dim OpaqueColorList(10) '0=white,1=black,2=red,3=green,4=blue,5=yellow,6=orange, 7=black


OpaqueColorList(0) = Array(Empty) ' WHITE
OpaqueColorList(1) = Array(Empty) ' GRAY
OpaqueColorList(2) = Array(Empty) ' RED
OpaqueColorList(3) = Array(Empty) ' GREEN
OpaqueColorList(4) = Array(Empty) ' BLUE
OpaqueColorList(5) = Array(Empty) ' YELLOW
OpaqueColorList(6) = Array(Empty) ' ORANGE
OpaqueColorList(7) = Array(Empty) ' BLACK
OpaqueColorList(8) = Array(Empty) ' UNUSED
OpaqueColorList(9) = Array(Primitive_LeftRamp,Primitive_RightRamp, Empty) ' UNUSED

Sub SetOpaqueColorListObj(nidx, nobj)
  Select Case color
    Case 0 nobj.material = WhiteVals(DNSVal)
    Case 1 nobj.material = GrayVals(DNSVal)
    Case 2 nobj.material = RedVals(DNSVal)
    Case 3 nobj.material = GreenVals(DNSVal)
    Case 4 nobj.material = BlueVals(DNSVal)
    Case 5 nobj.material = YellowVals(DNSVal)
    Case 6 nobj.material = OrangeVals(DNSVal)
    Case 7 nobj.material = BlackVals(DNSVal)
    Case 8 nobj.material = UserVals(DNSVal)
    Case 9 nobj.material = UserVals(DNSVal)
   End Select
  nobj.depthBias = OpaqueBias(DNSVal)
End Sub

'WHITE
Dim nii
for nii=0 to 9
for Each nxx in OpaqueColorList(nii)
  if NOT IsEmpty(nxx) then
  Select Case nii
    Case 0 nxx.material = WhiteVals(DNSVal)
    Case 1 nxx.material = GrayVals(DNSVal)
    Case 2 nxx.material = RedVals(DNSVal)
    Case 3 nxx.material = GreenVals(DNSVal)
    Case 4 nxx.material = BlueVals(DNSVal)
    Case 5 nxx.material = YellowVals(DNSVal)
    Case 6 nxx.material = OrangeVals(DNSVal)
    Case 7 nxx.material = BlackVals(DNSVal)
    Case 8 nxx.material = UserVals(DNSVal)
    Case 9 nxx.material = UserVals(DNSVal)
   End Select
  nxx.depthBias = OpaqueBias(DNSVal)
  end if
Next
Next

Sub SetTransListObj(obj)
obj.material = TransVals(DNSVal)
obj.depthBias = TransBias(DNSVal)
End Sub

if cUSETRANSLIST > 0 then ' import default_translist.mat materials,
Dim TransBias: TransBias=Array (-1000,-2000,-3000, -4000,-5000,-6000,-7000, -8000,-9000, -10000,-10000,-10000)
Dim TransVals: TransVals=Array ("translist9","translist9",_
"translist8","translist8", "translist7","translist7", _
"translist6","translist6", "translist5","translist5", _
"translist4","translist4")

For each nxx in TransList: if NOT IsEmpty(nxx) then:nxx.material = TransVals(DNSVal):end if:Next
For each nxx in TransList: if NOT IsEmpty(nxx) then:nxx.depthBias = TransBias(DNSVal):end if: Next

end if 'cUSETRANSLIST > 0

if cUSEOPAQUELIST > 0 then ' import default_opaquelist.mat materials,
Dim OpaqueVals: OpaqueVals=Array ("opaquelist9","opaquelist9",_
"opaquelist8","opaquelist8", "opaquelist7","opaquelist7", _
"opaquelist6","opaquelist6", "opaquelist5","opaquelist4", _
"opaquelist4","opaquelist4")

Const USEWHITE = 1
Const USEBLUE = 2
Const USERED = 3
Const USEYELLOW = 4
Const USEORANGE = 5
Const USEGRAY = 6
Const USEBLACK = 7

Dim MAT1: MAT1 = USEWHITE
Dim MAT2: MAT2 = USEBLUE

Dim BMAT0: BMAT0 = USEWHITE ' CAP
Dim BMAT1: BMAT1 = USEWHITE ' BASE
Dim BMAT2: BMAT2 = USEWHITE 'SKIRT
Dim BMAT3: BMAT3 = USEWHITE ' RING

'handle walls , generic flippers, rubbers, spinners,
Sub SetOpaqueListObj(obj)
if TypeName(obj) = "Wall" Then
obj.topmaterial = OpaqueVals(DNSVal)
obj.sidematerial = OpaqueVals(DNSVal)
elseif TypeName(obj) = "Bumper" Then

  if obj.capvisible = true Then
  if BMAT0 = USEWHITE then: obj.capmaterial = WhiteVals(DNSVal)
  if BMAT0 = USERED then: obj.capmaterial = RedVals(DNSVal)
  if BMAT0 = USEYELLOW then: obj.capmaterial = YellowVals(DNSVal)
  end if
  if obj.basevisible = true Then
  if BMAT1 = USEWHITE then: obj.basematerial = WhiteVals(DNSVal)
  end if
  if obj.ringvisible = true Then
  end if
  if obj.skirtvisible = true then
  if BMAT2 = USEWHITE then: obj.skirtmaterial = WhiteVals(DNSVal)
  if BMAT2 = USEBLUE then: obj.skirtmaterial = BlueVals(DNSVal)
  if BMAT2 = USERED then: obj.skirtmaterial = RedVals(DNSVal)
  if BMAT2 = USEYELLOW then: obj.skirtmaterial = YellowVals(DNSVal)
  if BMAT2 = USEORANGE then: obj.skirtmaterial = OrangeVals(DNSVal)
  end If
elseif TypeName(obj) = "Flipper" Then
  if MAT1 = USEWHITE then: obj.material = WhiteVals(DNSVal)
  if MAT1 = USEYELLOW then: obj.material = YellowVals(DNSVal)
  if MAT1 = USEBLACK then: obj.material = BlackVals(DNSVal)

  if MAT2 = USEBLUE then: obj.rubbermaterial = BlueVals(DNSVal)
  if MAT2 = USERED then: obj.rubbermaterial = RedVals(DNSVal)
  if MAT2 = USEYELLOW then: obj.rubbermaterial = YellowVals(DNSVal)
  if MAT2 = USEORANGE then: obj.rubbermaterial = OrangeVals(DNSVal)
  if MAT2 = USEGRAY then: obj.rubbermaterial = GrayVals(DNSVal)
  if MAT2 = USEBLACK then: obj.rubbermaterial = BlackVals(DNSVal)
elseif TypeName(obj) = "Rubber" OR TypeName(obj) = "Spinner" OR TypeName(obj) = "Ramp" Then
obj.material = OpaqueVals(DNSVal)
else ' --> and primitives, hit targets, drop targets
obj.material = OpaqueVals(DNSVal)
obj.depthBias = OpaqueBias(DNSVal)
end if
end sub

For each nxx in OpaqueList
if NOT IsEmpty(nxx) then
 SetOpaqueListObj(nxx)
end if
Next


'                                      on/off,DNS
Dim OpaqueHiLowList(3,12) '0=object, 1= material off, 2= material on
Dim LightUpDn(2)

if cUSEOPAQUELIST > 1 then

Const USEZERO = 1
Const OP4 ="opaquelist4"
Const OP5 ="opaquelist5"
Const OP6 ="opaquelist6"
Const OP7 ="opaquelist7"
Const OP8 ="opaquelist8"
Const OP9 ="opaquelist9"

' opaque object affected by light state list
OpaqueHiLowList(0,0) = Array(Empty ,Empty,Empty)


' lights off state
OpaqueHiLowList(1,0) = Array(OP9,OP9,OP9,OP9,OP9,OP9,OP9,OP9,OP9,OP9,OP9,OP9,OP9,OP9,OP9)
OpaqueHiLowList(1,1) = Array(OP9,OP9,OP9,OP9,OP9,OP9,OP9,OP9,OP9,OP9,OP9,OP9,OP9,OP9,OP9)
OpaqueHiLowList(1,2) = Array(OP8,OP8,OP8,OP8,OP8,OP8,OP8,OP8,OP8,OP8,OP8,OP8,OP8,OP8,OP8)
OpaqueHiLowList(1,3) = Array(OP8,OP8,OP8,OP8,OP8,OP8,OP8,OP8,OP8,OP8,OP8,OP8,OP8,OP8,OP8)
OpaqueHiLowList(1,4) = Array(OP7,OP7,OP7,OP7,OP7,OP7,OP7,OP7,OP7,OP7,OP7,OP7,OP7,OP7,OP7)
OpaqueHiLowList(1,5) = Array(OP7,OP7,OP7,OP7,OP7,OP7,OP7,OP7,OP7,OP7,OP7,OP7,OP7,OP7,OP7)
OpaqueHiLowList(1,6) = Array(OP6,OP6,OP6,OP6,OP6,OP6,OP6,OP6,OP6,OP6,OP6,OP6,OP6,OP6,OP6)
OpaqueHiLowList(1,7) = Array(OP6,OP6,OP6,OP6,OP6,OP6,OP6,OP6,OP6,OP6,OP6,OP6,OP6,OP6,OP6)
OpaqueHiLowList(1,8) = Array(OP5,OP5,OP5,OP5,OP5,OP5,OP5,OP5,OP5,OP5,OP5,OP5,OP5,OP5,OP5)
OpaqueHiLowList(1,9) = Array(OP4,OP4,OP4,OP4,OP4,OP4,OP4,OP4,OP4,OP4,OP4,OP4,OP4,OP4,OP4)
OpaqueHiLowList(1,10)= Array(OP4,OP4,OP4,OP4,OP4,OP4,OP4,OP4,OP4,OP4,OP4,OP4,OP4,OP4,OP4)
OpaqueHiLowList(1,11)= Array(OP4,OP4,OP4,OP4,OP4,OP4,OP4,OP4,OP4,OP4,OP4,OP4,OP4,OP4,OP4)
' lights on state
OpaqueHiLowList(2,0) = Array(OP8,OP8,OP8,OP8,OP8,OP8,OP8,OP8,OP8,OP8,OP8,OP8,OP8,OP8,OP8)
OpaqueHiLowList(2,1) = Array(OP8,OP8,OP8,OP8,OP8,OP8,OP8,OP8,OP8,OP8,OP8,OP8,OP8,OP8,OP8)
OpaqueHiLowList(2,2) = Array(OP7,OP7,OP7,OP7,OP7,OP7,OP7,OP7,OP7,OP7,OP7,OP7,OP7,OP7,OP7)
OpaqueHiLowList(2,3) = Array(OP7,OP7,OP7,OP7,OP7,OP7,OP7,OP7,OP7,OP7,OP7,OP7,OP7,OP7,OP7)
OpaqueHiLowList(2,4) = Array(OP6,OP6,OP6,OP6,OP6,OP6,OP6,OP6,OP6,OP6,OP6,OP6,OP6,OP6,OP6)
OpaqueHiLowList(2,5) = Array(OP6,OP6,OP6,OP6,OP6,OP6,OP6,OP6,OP6,OP6,OP6,OP6,OP6,OP6,OP6)
OpaqueHiLowList(2,6) = Array(OP5,OP5,OP5,OP5,OP5,OP5,OP5,OP5,OP5,OP5,OP5,OP5,OP5,OP5,OP5)
OpaqueHiLowList(2,7) = Array(OP5,OP5,OP5,OP5,OP5,OP5,OP5,OP5,OP5,OP5,OP5,OP5,OP5,OP5,OP5)
OpaqueHiLowList(2,8) = Array(OP4,OP4,OP4,OP4,OP4,OP4,OP4,OP4,OP4,OP4,OP4,OP4,OP4,OP4,OP4)
OpaqueHiLowList(2,9) = Array(OP4,OP4,OP4,OP4,OP4,OP4,OP4,OP4,OP4,OP4,OP4,OP4,OP4,OP4,OP4)
OpaqueHiLowList(2,10)= Array(OP4,OP4,OP4,OP4,OP4,OP4,OP4,OP4,OP4,OP4,OP4,OP4,OP4,OP4,OP4)
OpaqueHiLowList(2,11)= Array(OP4,OP4,OP4,OP4,OP4,OP4,OP4,OP4,OP4,OP4,OP4,OP4,OP4,OP4,OP4)


LightUpDn(0) = Array (0.1,0.2,0.3,0.4,0.5,0.6,0.7,0.8,0.9,1.0,1.0)
LightUpDn(1) = Array (0.6,0.5,0.5,0.5,0.55,0.65,0.8,0.9,1.0,1.1,1.1)

fnSwitchLightState 0, Array(Empty)

end if 'cUSEOPAQUELIST > 1

Sub fnSwitchLightState(onoff, iscalelist)
Dim xx, cnt, ar
if cUSEOPAQUELIST > 1 then
cnt =0
  if onoff = 0 then  ' ----------- OFF --------------

  ar =LightUpDn(0)
    for Each xx in iscalelist
      if NOT IsEmpty(xx) then: xx.intensityscale =  ar(DNSVal)
    Next

    for Each xx in OpaqueHiLowList(0,0)
      if NOT IsEmpty(xx) then

        if TypeName(xx) = "Bumper" Then
        'MsgBox "bumper"

          if xx.capvisible = true Then
            if BMAT0 = USEWHITE then: xx.capmaterial = WhiteVals(DNSVal)
          end if
          if xx.basevisible = true Then
            if BMAT1 = USEWHITE then: xx.basematerial = WhiteVals(DNSVal)
          end if
          if xx.ringvisible = true Then
          end if
          if xx.skirtvisible = true then
            if BMAT2 = USEWHITE then: xx.skirtmaterial = WhiteVals(DNSVal)
            if BMAT2 = USEBLUE then: xx.skirtmaterial = BlueVals(DNSVal)
            if BMAT2 = USERED then: xx.skirtmaterial = RedVals(DNSVal)
          end If

        else
        ar = OpaqueHiLowList(1,DNSVal)
        xx.material = ar(cnt)
        end if
      end if
      if USEZERO Then
      cnt = 0
      else
      cnt = cnt+1
      end if
    Next

  Else ' ----------- ON --------------

  ar =LightUpDn(1)
    for Each xx in iscalelist
      if NOT IsEmpty(xx) then: xx.intensityscale =  ar(DNSVal)
    Next

    for Each xx in OpaqueHiLowList(0,0)
      if NOT IsEmpty(xx) then
        if TypeName(xx) = "Bumper" Then

          if xx.capvisible = true Then
            if BMAT0 = USEWHITE then: xx.capmaterial = WhiteVals(DNSVal+4)
          end if
          if xx.basevisible = true Then
            if BMAT1 = USEWHITE then: xx.basematerial = WhiteVals(DNSVal+4)
          end if
          if xx.ringvisible = true Then
          end if
          if xx.skirtvisible = true then
            if BMAT2 = USEWHITE then: xx.skirtmaterial = WhiteVals(DNSVal+4)
            if BMAT2 = USEBLUE then: xx.skirtmaterial = BlueVals(DNSVal+4)
            if BMAT2 = USERED then: xx.skirtmaterial = RedVals(DNSVal+4)
          end If
        else

        ar = OpaqueHiLowList(2,DNSVal)
        xx.material = ar(cnt)
        end if
      end if

      if USEZERO Then
      cnt = 0
      else
      cnt = cnt+1
      end if
    Next

  end if ' ON/OFF
end if 'cUSEOPAQUELIST > 1
end sub



end if 'cUSEOPAQUELIST > 0

if cUSESEMITRANSLIST > 0 then
Dim SemiTransBias: SemiTransBias=Array ( -100, -200,-300,-400,-500, -600,-700, -800,-900,-1000,-1000)
Dim SemiTransVals: SemiTransVals=Array ("semitranslist9","semitranslist9",_
"semitranslist8","semitranslist8", "semitranslist7","semitranslist7", _
"semitranslist6","semitranslist6", "semitranslist5","semitranslist5", _
"semitranslist4","semitranslist4")


Sub SetSemiTransListObj(obj)
obj.material = SemiTransVals(DNSVal)
obj.depthBias = SemiTransBias(DNSVal)
End Sub

For each nxx in SemiTransList: if NOT IsEmpty(nxx) then:nxx.material = SemiTransVals(DNSVal):end if:Next
For each nxx in SemiTransList: if NOT IsEmpty(nxx) then:nxx.depthBias = SemiTransBias(DNSVal): end if: Next

end if 'cUSESEMITRANSLIST > 0


' color grade transition
if cUSECOLORGRADE then
Const nUSEREDREDUXFIX = 1

Const nCGDiv = 1.0 ' increase when gets to understaturated at higher light settings
Const rSHIFT = 0 ' increase too saturated at lower light settings
Dim ColValues

if nUSEREDREDUXFIX = 0 then
ColValues=Array ("_ColorGrade_9_Lite","_ColorGrade_8_Lite","_ColorGrade_7_Lite","_ColorGrade_6_Lite", _
"_ColorGrade_5_Lite","_ColorGrade_4_Lite","_ColorGrade_3_Lite","_ColorGrade_2_Lite","_ColorGrade_1_Lite","_ColorGrade_0_Lite","_ColorGrade_0_Lite","_ColorGrade_0_Lite",_
"_ColorGrade_0_Lite","_ColorGrade_0_Lite","_ColorGrade_0_Lite","_ColorGrade_0_Lite","_ColorGrade_0_Lite","_ColorGrade_0_Lite")
else
ColValues=Array ("_ColorGrade_RedReduxFix9","_ColorGrade_RedReduxFix8","_ColorGrade_RedReduxFix7","_ColorGrade_RedReduxFix6", _
"_ColorGrade_RedReduxFix5","_ColorGrade_RedReduxFix4","_ColorGrade_RedReduxFix3","_ColorGrade_RedReduxFix2","_ColorGrade_RedReduxFix1","_ColorGrade_RedReduxFix0","_ColorGrade_RedReduxFix0","_ColorGrade_RedReduxFix0",_
"_ColorGrade_RedReduxFix0","_ColorGrade_RedReduxFix0","_ColorGrade_RedReduxFix0","_ColorGrade_RedReduxFix0","_ColorGrade_RedReduxFix0","_ColorGrade_RedReduxFix0")
end if
Genesis.ColorGradeImage = ColValues((int)((DNSVal+rSHIFT)/nCGDiv))
end if 'cUSECOLORGRADE


Genesis.PlayfieldReflectionStrength = ReflDN(DNSVal + 4)

If GILevel = 0 then ' default = Real
For each nxx in aGiLights:nxx.intensity = nxx.intensity * SysDNSVal(DNSVal) /DivValues3(DNSVal) :Next
elseif GILevel = 1 then ' half of real reduced
For each nxx in aGiLights:nxx.intensity = nxx.intensity * SysDNSVal(DNSVal) /DivValues4(DNSVal) :Next
elseif GILevel = 2 then' 1/4 of real reduced
For each nxx in aGiLights:nxx.intensity = nxx.intensity * SysDNSVal(DNSVal) /DivValues5(DNSVal) :Next
elseif GILevel = 3 then' 1/8'th of real reduced
For each nxx in aGiLights:nxx.intensity = nxx.intensity * SysDNSVal(DNSVal) /DivValues6(DNSVal) :Next
elseif GILevel = 4 then' 1/16'th of real reduced
For each nxx in aGiLights:nxx.intensity = nxx.intensity * SysDNSVal(DNSVal) /DivValues7(DNSVal) :Next
elseif GILevel = 5 then' 1/32'th of real reduced
For each nxx in aGiLights:nxx.intensity = nxx.intensity * SysDNSVal(DNSVal) /DivValues8(DNSVal) :Next
end if


If FLLevel = 0 then ' default = Real
For each nxx in aAllFlashers:nxx.amount = nxx.amount / DivValues3(DNSVal):Next
For each nxx in aAllFlashers:nxx.opacity = nxx.opacity * OPSValues(DNSVal) / DivLevel:Next
elseif FLLevel = 1 then ' half of real reduced
For each nxx in aAllFlashers:nxx.amount = nxx.amount / DivValues4(DNSVal):Next
For each nxx in aAllFlashers:nxx.opacity = nxx.opacity * OPSValues2(DNSVal) / DivLevel:Next
elseif FLLevel = 2 then ' 1/4 of real reduced
For each nxx in aAllFlashers:nxx.amount = nxx.amount / DivValues5(DNSVal):Next
For each nxx in aAllFlashers:nxx.opacity = nxx.opacity * OPSValues3(DNSVal) / DivLevel:Next
elseif FLLevel = 3 then ' 1/8'th of real reduced
For each nxx in aAllFlashers:nxx.amount = nxx.amount / DivValues6(DNSVal):Next
For each nxx in aAllFlashers:nxx.opacity = nxx.opacity * OPSValues4(DNSVal) / DivLevel:Next
elseif FLLevel = 4 then ' 1/16'th of real reduced
For each nxx in aAllFlashers:nxx.amount = nxx.amount / DivValues7(DNSVal):Next
For each nxx in aAllFlashers:nxx.opacity = nxx.opacity * OPSValues5(DNSVal) / DivLevel:Next
elseif FLLevel = 5 then ' 1/32'th of real reduced
For each nxx in aAllFlashers:nxx.amount = nxx.amount / DivValues8(DNSVal):Next
For each nxx in aAllFlashers:nxx.opacity = nxx.opacity * OPSValues6(DNSVal) / DivLevel:Next
end If

If BTLevel = 0 then ' default = Real
For each nxx in BloomLights:nxx.intensity = nxx.intensity *SysDNSVal(DNSVal)/DivValues3(DNSVal):Next
For each nxx in TargetDropGi:nxx.intensity = nxx.intensity *SysDNSVal(DNSVal)/DivValues3(DNSVal):Next
elseif BTLevel = 1 then ' half of real reduced
For each nxx in BloomLights:nxx.intensity = nxx.intensity *SysDNSVal(DNSVal)/DivValues4(DNSVal):Next
For each nxx in TargetDropGi:nxx.intensity = nxx.intensity *SysDNSVal(DNSVal)/DivValues4(DNSVal):Next
elseif BTLevel = 2 then' 1/4 of real reduced
For each nxx in BloomLights:nxx.intensity = nxx.intensity *SysDNSVal(DNSVal)/DivValues5(DNSVal):Next
For each nxx in TargetDropGi:nxx.intensity = nxx.intensity *SysDNSVal(DNSVal)/DivValues5(DNSVal):Next
elseif BTLevel = 3 then' 1/8'th of real reduced
For each nxx in BloomLights:nxx.intensity = nxx.intensity *SysDNSVal(DNSVal)/DivValues6(DNSVal):Next
For each nxx in TargetDropGi:nxx.intensity = nxx.intensity *SysDNSVal(DNSVal)/DivValues6(DNSVal):Next
elseif BTLevel = 4 then' 1/16'th of real reduced
For each nxx in BloomLights:nxx.intensity = nxx.intensity *SysDNSVal(DNSVal)/DivValues7(DNSVal):Next
For each nxx in TargetDropGi:nxx.intensity = nxx.intensity *SysDNSVal(DNSVal)/DivValues7(DNSVal):Next
elseif BTLevel = 5 then' 1/32'th of real reduced
For each nxx in BloomLights:nxx.intensity = nxx.intensity *SysDNSVal(DNSVal)/DivValues8(DNSVal):Next
For each nxx in TargetDropGi:nxx.intensity = nxx.intensity *SysDNSVal(DNSVal)/DivValues8(DNSVal):Next
end if


Const DynPinOffset = 3
Dim DynPinVals
Dim DynPinRefVals
Dim DynPinRefSclVals

if cDYNPINBALL=1 then
DynPinVals=Array ("_defballHDR0","_defballHDR0", _
"_defballHDR1","_defballHDR1", "_defballHDR1", _
"_defballHDR2","_defballHDR2", _
"_defballHDR3","_defballHDR3","_defballHDR4","_defballHDR4","_defballHDR4","_defballHDR4","_defballHDR4")

Genesis.ballImage = DynPinVals(DNSVal)

Genesis.BallReflection = 10
Genesis.BallPlayfieldReflectionScale = 0.01

elseif cDYNPINBALL=2 then
DynPinVals=Array ("_defballHDR4","_defballHDR4", _
"_defballHDR3","_defballHDR3", "_defballHDR2", _
"_defballHDR2","_defballHDR1", _
"_defballHDR1","_defballHDR1","_defballHDR0","_defballHDR0","_defballHDR0","_defballHDR0","_defballHDR0")

DynPinRefVals=Array (200, 180,160,140,120,100,80,60,40,20,10,10)
DynPinRefSclVals = Array(0.1,0.09,0.08,0.07,0.06,0.05, 0.04,0.03,0.02, 0.01,0.01)

Genesis.ballImage = DynPinVals(DNSVal+DynPinOffset)
Genesis.BallReflection = DynPinRefVals(DNSVal)
Genesis.BallPlayfieldReflectionScale = DynPinRefSclVals(DNSVal)

end if 'cDYNPINBALL


if cDYNGAMEBLADES then
Dim DynBladeVals: DynBladeVals=Array ("_defgbladeswoodblack0","_defgbladeswoodblack0",_
"_defgbladeswoodblack1","_defgbladeswoodblack1", "_defgbladeswoodblack1",_
"_defgbladeswoodblack2","_defgbladeswoodblack2",_
"_defgbladeswoodblack3","_defgbladeswoodblack3","_defgbladeswoodblack3","_defgbladeswoodblack3","_defgbladeswoodblack3")

Primitive5.image = DynBladeVals(DNSVal)
Primitive6.image = DynBladeVals(DNSVal)
end if 'cDYNGAMEBLADES


Sub SetSMLiDNS(object, enabled)
  If enabled > 0 then
  object.intensity = enabled * SysDNSVal(DNSVal) /DivValues2(DNSVal)
  Else
  object.intensity =0
  end if
End Sub

Sub SetSMFlDNS(object, enabled)

  If enabled > 0 then
  object.opacity = enabled / DivValues2(DNSVal)
  else
  object.opacity = 0
  end if
End Sub


Sub SetSLiDNS(object, enabled)
  If enabled then
  object.intensity = 1 * SysDNSVal(DNSVal) /DivValues2(DNSVal)
  Else
  object.intensity =0
  end if
End Sub

Sub SetSFlDNS(object, enabled)

  If enabled then
  object.opacity = 1 * OPSValues(DNSVal) / DivLevel
  else
  object.opacity = 0
  end if
End Sub


Sub SetDNSFlash(nr, object)
    Select Case LightState(nr)
        Case 0 'off
            FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown(nr)
            If FlashLevel(nr) < FlashMin(nr) Then
                FlashLevel(nr) = FlashMin(nr)
                LightState(nr) = -1 'completely off, so stop the fading loop
            End if
            Object.IntensityScale = FlashLevel(nr) * SysDNSVal(DNSVal) /DivValues2(DNSVal)
        Case 1 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
            If FlashLevel(nr) > FlashMax(nr) Then
                FlashLevel(nr) = FlashMax(nr)
                LightState(nr) = -1 'completely on, so stop the fading loop
            End if
            Object.IntensityScale = FlashLevel(nr) * SysDNSVal(DNSVal) /DivValues2(DNSVal)
    End Select
End Sub


Sub SetDNSFlashm(nr, object) 'multiple flashers, it just sets the intensity
    Object.IntensityScale = FlashLevel(nr) * SysDNSVal(DNSVal) /DivValues2(DNSVal)
End Sub

Sub SetDNSFlashex(object, value) 'it just sets the intensityscale for non system lights
    Object.IntensityScale = value * SysDNSVal(DNSVal) /DivValues2(DNSVal)
End Sub

' usage: fnColorGradeImage step, nCGDiv
Sub fnColorGradeImage(step, div)
Dim ii
  if cUSECOLORGRADE then
    'table1.ColorGradeImage = ColValues((int)(DNSVal/nCGDiv))

      If step = -1 Then
      Genesis.ColorGradeImage =ColValues((int)((DNSVal+rSHIFT)/nCGDiv)) ' default state
      else
      if (DNSVal - 4) < 0 then '0,1,2
        if step+1 > 5 then
        ii =2
        elseif  step+1 > 2 then
        ii =1
        Else
        ii =0
        end if
      Genesis.ColorGradeImage = ColValues((int)(ii/div)) ' very dark
      Else
        If DNSVal <= 7 then  '3,4,5
          if step+1 > 5 then
          ii =5
          elseif  step+1 > 2 then
          ii =4
          Else
          ii =3
          end if
        Genesis.ColorGradeImage = ColValues((int)(ii/div)) ' average dark
        else

          if step+1 > 5 then
          ii =8
          elseif  step+1 > 2 then
          ii =7
          Else
          ii =6
          end if
        Genesis.ColorGradeImage = ColValues((int)(ii/div))
        end if
      end if
      end if
  end if
End Sub

'**************************************************************************************
'               [CC Color Correction] Textures must be CC corrected to work
'**************************************************************************************
Const UseLightRamp = 1
Const UseMixGtoB = 1
Dim FlGlobals : FlGlobals=Array (1.0,1.5,2.0,2.5,3.0,3.5,4.0,4.5,5.0,5.0,5.0)
Dim FlGlobRamp : FlGlobRamp=Array (1.0,1.1,1.1,1.2,1.2,1.3,1.3,1.4,1.4,1.5,1.5)

Dim DivGlobal
if UseLightRamp = 0 then
DivGlobal = FlGlobals(DNSVal) * 1.0 ' multiply between 1.0 to 0.75, 1.0 = default
else
DivGlobal = FlGlobals(DNSVal) * FlGlobRamp(DNSVal) * 1.0
end if


Dim DivClrCor: DivClrCor = 1.1

Dim BlueDiv:BlueDiv  = 1.0* DivGlobal
Dim RedDiv:RedDiv  = 0.9* DivGlobal
Dim RedGreenDiv:RedGreenDiv  = 1.0* DivGlobal
Dim FullSpecDiv:FullSpecDiv  = 0.1* DivGlobal
Dim CCFull: CCFull = 1.4 * DivClrCor
Dim CCBlue: CCBlue = 1.2 * DivClrCor
Dim CCRedGreen: CCRedGreen = 0.9 * DivClrCor
Dim CCRed: CCRed = 0.9 * DivClrCor

' original values depricated
'Dim FlValues : FlValues=Array (1.0,1.0,1.0,1.0,1.0,1.0,1.0,1.0)
'Dim AMTValues: AMTValues=Array (500,3000,4000,8000,16000,32000,64000,128000,128000,128000,128000,128000)
'Dim BiasValues: BiasValues=Array (-100,-90,-80,-70,-60,-50,-40,-30,-20,-10,-9,-8)
'Dim MixBlueChan: MixBlueChan=Array (8.0*BlueDiv, 16.0*BlueDiv, 34.0*BlueDiv, 64.0*BlueDiv, 72.0*BlueDiv, 80.0*BlueDiv, 84.0*BlueDiv, 82.0*BlueDiv, 80.0*BlueDiv, 78.0*BlueDiv, 76.0*BlueDiv)
'Dim MixRedChan: MixRedChan=Array (16.0 * RedDiv, 16.0* RedDiv, 34.0* RedDiv, 64.0* RedDiv, 72.0* RedDiv, 80.0* RedDiv, 84.0* RedDiv, 82.0* RedDiv, 80.0* RedDiv, 78.0* RedDiv, 76.0* RedDiv)
'Dim MixRedGreenChan: MixRedGreenChan=Array (8.0*RedGreenDiv, 16.0*RedGreenDiv, 34.0*RedGreenDiv, 64.0*RedGreenDiv, 72.0*RedGreenDiv, 80.0*RedGreenDiv, 84.0*RedGreenDiv, 82.0*RedGreenDiv, 80.0*RedGreenDiv, 78.0*RedGreenDiv, 76.0*RedGreenDiv)
'Dim MixFullSpectrum: MixFullSpectrum=Array (8.0*FullSpecDiv, 16.0*FullSpecDiv, 34.0*FullSpecDiv, 64.0*FullSpecDiv, 72.0*FullSpecDiv, 73.0*FullSpecDiv, 74.0*FullSpecDiv, 74.0*FullSpecDiv, 74.0*FullSpecDiv, 74.0*FullSpecDiv, 74.0*FullSpecDiv)
'Dim CCFSValues: CCFSValues=Array (100.0*CCFull, 1.0*CCFull, 0.1*CCFull, 0.01*CCFull, 0.001*CCFull, 0.0001*CCFull, 0.0001*CCFull, 0.0001*CCFull, 0.0001*CCFull,0.0001*CCFull,0.0001*CCFull)
'Dim CCBlueValues: CCBlueValues=Array (0.9*CCBlue, 1.0*CCBlue, 1.2*CCBlue, 1.4*CCBlue, 1.5*CCBlue, 1.55*CCBlue, 1.56*CCBlue, 1.565*CCBlue, 1.568*CCBlue, 1.569*CCBlue, 1.569*CCBlue)
'Dim MixGtoB : MixGtoB=Array (1,1,2,2,3,4,5,6,7,8,9,10)

'new values
Dim FlValues : FlValues             =Array (1.0,1.0,1.0,1.0,1.0,1.0,1.0,1.0)
Dim AMTValues: AMTValues            =Array (500,3000,4000,8000,16000,32000,64000,80000,90000,100000,128000,128000)
Dim BiasValues: BiasValues          =Array (-100,-90,-80,-70,-60,-50,-40,-30,-20,-10,-9,-8)
Dim MixBlueChan: MixBlueChan        =Array (8.0*BlueDiv    , 16.0*BlueDiv    , 19.0*BlueDiv    , 20.0*BlueDiv    , 30.0*BlueDiv    , 60.0*BlueDiv    , 21.0*BlueDiv    , 52.0*BlueDiv    , 70.5*BlueDiv    , 100.0*BlueDiv    , 100.0*BlueDiv)
Dim MixRedChan: MixRedChan          =Array (8.0*RedDiv     , 16.0* RedDiv    , 19.0* RedDiv    , 20.0* RedDiv    , 30.0* RedDiv    , 60.0* RedDiv    , 21.0* RedDiv    , 52.0* RedDiv    , 70.5* RedDiv    , 100.0* RedDiv    , 100.0* RedDiv)
'                                                            |                                                                       |                                                                       |
Dim MixRedGreenChan: MixRedGreenChan=Array (8.0*RedGreenDiv, 16.0*RedGreenDiv, 19.0*RedGreenDiv, 20.0*RedGreenDiv, 30.0*RedGreenDiv, 60.0*RedGreenDiv, 21.0*RedGreenDiv, 52.0*RedGreenDiv, 70.5*RedGreenDiv, 100.0*RedGreenDiv, 100.0*RedGreenDiv)
Dim MixFullSpectrum: MixFullSpectrum=Array (8.0*FullSpecDiv, 16.0*FullSpecDiv, 19.0*FullSpecDiv, 20.0*FullSpecDiv, 30.0*FullSpecDiv, 60.0*FullSpecDiv, 20.0*FullSpecDiv, 52.0*FullSpecDiv, 70.5*FullSpecDiv, 100.0*FullSpecDiv, 100.0*FullSpecDiv)
Dim CCFSValues: CCFSValues          =Array (100.0*CCFull   , 1.0*CCFull      , 0.1*CCFull      , 0.01*CCFull     , 0.001*CCFull    , 0.0001*CCFull   , 0.0001*CCFull   , 0.0001*CCFull   , 0.0001*CCFull   , 0.0001*CCFull    , 0.0001*CCFull)
Dim CCBlueValues: CCBlueValues      =Array (0.9*CCBlue     , 1.0*CCBlue      , 1.2*CCBlue      , 1.4*CCBlue      , 1.5*CCBlue      , 1.55*CCBlue     , 1.56*CCBlue     , 1.565*CCBlue    , 1.568*CCBlue    , 1.569*CCBlue     , 1.569*CCBlue)
Dim MixGtoB : MixGtoB=Array (1,1,2,2,3,4,5,6,7,8,9,10)


Dim FlData : FlData=Array (Flasher1,Flasher2,Flasher3,Flasher4,FlBGBlueSpec,FlBGRedGreenSpec,Empty,Empty)

'Full blue channel
FlData(0).amount = AMTValues(DNSVal ) * DivValues3(DNSVal)
FlData(0).DepthBias = BiasValues(DNSVal) * 100.0
FlData(0).intensityscale = FlData(0).intensityscale * MixBlueChan(DNSVal) '0.0008'
FlData(0).opacity = CCBlueValues(DNSVal) '* CCFSValues(DNSVal)'

'Full blue channel
if UseMixGtoB then
FlData(0).Color = RGB(0, MixGtoB(DNSVal), 20)
end if

'Half red channel
FlData(1).amount = AMTValues(DNSVal ) * DivValues3(DNSVal)
FlData(1).DepthBias = BiasValues(DNSVal) * 100.0
FlData(1).intensityscale = FlData(1).intensityscale * (MixRedChan(DNSVal) * SysDNSVal(DNSVal)) '0.0008'
FlData(1).opacity = CCRed
'Full red/green channel
FlData(2).amount = AMTValues(DNSVal ) * DivValues3(DNSVal)
FlData(2).DepthBias = BiasValues(DNSVal) * 100.0
FlData(2).intensityscale = FlData(2).intensityscale * MixRedGreenChan(DNSVal) '0.0008'
FlData(2).opacity = CCRedGreen
'Full Spectrum
FlData(3).amount = AMTValues(DNSVal ) * (DivValues3(DNSVal) * 0.1)
FlData(3).DepthBias = BiasValues(DNSVal) * 100.0
'FlData(3).intensityscale = FlData(3).intensityscale * MixFullSpectrum(DNSVal) '0.0008'
'FlData(3).opacity = FlData(3).opacity * CCFSValues(DNSVal)

Const USEBGBLUE = 1
Const USEBGREDGREEN = 1
Const nDAMPNBLUE = 1
Const nDAMPNREDGREEN = 0

Dim nbChan: nbChan = 1.0
Dim bCCXDN : bCCXDN=Array (0.1*nbChan,0.2*nbChan,0.3*nbChan,0.4*nbChan,0.5*nbChan,0.6*nbChan,0.7*nbChan,0.8*nbChan,0.9*nbChan,1.0*nbChan,1.0*nbChan,1.0*nbChan,1.0*nbChan,1.0*nbChan)
Dim nrgChan: nrgChan = 1.0
Dim rgCCXDN : rgCCXDN=Array (0.1*nrgChan,0.2*nrgChan,0.3*nrgChan,0.4*nrgChan,0.5*nrgChan,0.6*nrgChan,0.7*nrgChan,0.8*nrgChan,0.9*nrgChan,1.0*nrgChan,1.0*nrgChan,1.0*nrgChan,1.0*nrgChan,1.0*nrgChan)
Dim nnn

if cUSEBGLASSCCX then

' when too much blue influence dampen blue chan
if nDAMPNBLUE then
nnn=0
  For each nxx in MixBlueChan
  MixBlueChan(nnn) = MixBlueChan(nnn) * bCCXDN(nnn)
  nnn = nnn +1
  Next
end If 'nDAMPNBLUE

' when too much red/green influence dampen blue chan
if nDAMPNREDGREEN then
nnn=0
  For each nxx in MixRedGreenChan
  MixRedGreenChan(nnn) = MixRedGreenChan(nnn) * rgCCXDN(nnn)
  nnn = nnn +1
  Next
end if 'nDAMPNREDGREEN

' add (FlBGBlueSpec,FlBGRedGreenSpec) to FlData Array and set indexes correctly
Dim BGCCXVals : BGCCXVals=Array (1.6,1.5,1.4,1.3,1.1,1.03,1.01,1.008,1.005,1.003 , 1.001, 1.0005)
Dim nOff: nOff = 0

if USEBGBLUE then
FlData(4+nOff).visible = 1
end if
if USEBGREDGREEN then
FlData(5+nOff).visible = 1
end if

FlData(4+nOff).opacity = OPSValues(DNSVal + DNShift) / DivValues(DNSVal)
FlData(4+nOff).intensityscale = DNSValues(DNSVal + DNShift) /DivValues(DNSVal)


FlData(5+nOff).opacity = OPSValues(DNSVal + DNShift) /DivValues(DNSVal)
FlData(5+nOff).intensityscale = DNSValues(DNSVal + DNShift) /DivValues(DNSVal)


' blue chan
FlData(4+nOff).amount = AMTValues(DNSVal ) * (DivValues3(DNSVal) * BGCCXVals(DNSVal) )' 0.001)
FlData(4+nOff).DepthBias = BiasValues(DNSVal) * 100.0
FlData(4+nOff).intensityscale = FlData(4+nOff).intensityscale * MixBlueChan(DNSVal) '0.0008'
FlData(4+nOff).opacity = CCBlueValues(DNSVal)  '* CCFSValues(DNSVal)'
' red green chan
FlData(5+nOff).amount = AMTValues(DNSVal ) * (DivValues3(DNSVal) * BGCCXVals(DNSVal) ) '0.001)
FlData(5+nOff).DepthBias = BiasValues(DNSVal) * 100.0
FlData(5+nOff).intensityscale = FlData(5+nOff).intensityscale * MixRedGreenChan(DNSVal) '0.0008'
FlData(5+nOff).opacity = CCRedGreen

end if 'cUSEBGLASSCCX

if cUSECCUPBOOST then
Dim DivCCUP: DivCCUP = 0.5

'original values depricated
'Dim CCUPValues : CCUPValues=Array (1.0*DivCCUP,1.0*DivCCUP,1.0*DivCCUP,1.5*DivCCUP,2.5*DivCCUP,5.0*DivCCUP,10.0*DivCCUP,20.0*DivCCUP,30.0*DivCCUP,40.0*DivCCUP,40.0*DivCCUP)
'new values
Dim CCUPValues : CCUPValues=Array (1.0*DivCCUP,1.0*DivCCUP,1.0*DivCCUP,1.5*DivCCUP,2.5*DivCCUP,5.0*DivCCUP,10.0*DivCCUP,12.0*DivCCUP,20.0*DivCCUP,40.0*DivCCUP,40.0*DivCCUP)


FlData(3).intensityscale = FlData(3).intensityscale * MixFullSpectrum(DNSVal) * CCUPValues(DNSVal)'0.0008'
FlData(3).opacity = FlData(3).opacity * CCFSValues(DNSVal) * CCUPValues(DNSVal)
Else
FlData(3).intensityscale = FlData(3).intensityscale * MixFullSpectrum(DNSVal) '0.0008'
FlData(3).opacity = FlData(3).opacity * CCFSValues(DNSVal)
end if 'cUSECCUPBOOST


FlValues(0) = FlData(0).intensityscale
FlValues(1) = FlData(1).intensityscale
FlValues(2) = FlData(2).intensityscale
FlValues(3) = FlData(3).intensityscale

' ***************************************************************************
' TABLE surface and reflexion shifting for CC and D&N
' ***************************************************************************
Const Forward = 150 ' 150
Const Height = 200 '250
Const RotX = -5 ' -5
'Full
FlData(0).y = FlData(0).y + Forward
FlData(0).Height = Height
FlData(0).RotX = Rotx
' upper half
FlData(1).y = FlData(1).y + Forward
FlData(1).Height = Height + 50
FlData(1).RotX = Rotx
'full
FlData(2).y = FlData(2).y + Forward
FlData(2).Height = Height
FlData(2).RotX = Rotx
' half
FlData(3).y = FlData(3).y + Forward
FlData(3).Height = Height+ 50
FlData(3).RotX = Rotx

if NOT IsEmpty(FlData(4)) then
  if cUSEBGLASSCCX > 0 then
    if  4+nOff > 4 then
    FlData(4).y = FlData(4).y + Forward
    FlData(4).Height = Height
    FlData(4).RotX = Rotx
    end if
  else
  FlData(4).y = FlData(4).y + Forward
  FlData(4).Height = Height
  FlData(4).RotX = Rotx
  end if
End if

'BGMirror.y = FlData().y + Forward
'BGMirror.Height = Height
'BGMirror.RotX = Rotx
' ***************************************************************************

Dim UseVPMDMD
Dim DesktopMode:DesktopMode = Genesis.ShowDT

If DesktopMode = True Then 'Show Desktop components
Ramp10.visible = 1
Ramp5.visible = 1
Ramp37.visible = 1
Ramp38.visible = 1
Prim_Sidewalls.visible=1
Else
Ramp10.visible = 0
Ramp5.visible = 0
Ramp37.visible = 0
Ramp38.visible = 0
Prim_Sidewalls.visible=0
End if

UseVPMDMD = 0


dim fullscreendisplay
''===================
'\\\\\\\\\\\\\\\\\\\\\
'fullscreen display options
'2 = fullscreen display (best for single screen setups)
'1 = movable pinmame DMD window
'0 = No fullscreen display (use B2S instead)

fullscreendisplay = 0

'/////////////////////
'====================

dim Dropfix

'Set this to 1 if you are using VP10 1.0 and the drop targets aren't resetting properly
Dropfix = 0



 'setup display
dim xx
If DesktopMode then
  For each xx in Display:xx.X = xx.X - 150: xx.Y = xx.Y - 400: xx.rotX = -55: xx.height = xx.height + 320: Next
  For each xx in Display2:xx.Y = xx.y - 20: xx.X = xx.X - 6: xx.height = xx.height - 30: Next
end if


'Load VPM and scripts

LoadVPM "01560000", "sys80.vbs", 3.36

Const cGameName = "genesis"
Const UseSolenoids = 2
Const UseLamps = 0

'Standard sounds
Const SSolenoidOn = "Solenoid"    'Solenoid activates
Const SSolenoidOff = ""           'Solenoid deactivates
Const SCoin = "Coin"              'Coin inserted

Dim bsTrough, bsArmsLock, bsLegsLock, dtM', plungerIM
Dim x, Balls', bump1, bump2, bump3, bump4
Dim Last12, Current12, Last13, Current13, Last14, Current14

Set MotorCallback = GetRef("RollingUpdate") 'realtime updates - rolling sound

'**********
'Table Init
'**********

Sub Genesis_Init
  vpmInit Me

  With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
    .SplashInfoLine = "Genesis (Gottlieb 1986)" & vbNewLine & "VP9 table by JPSalas"
'   .Games(cGameName).Settings.Value("dmd_red") = 0
'   .Games(cGameName).Settings.Value("dmd_green") = 128
'   .Games(cGameName).Settings.Value("dmd_blue") = 255
'   .Games(cGameName).Settings.Value("rol") = 0
    .HandleKeyboard = 0
    .ShowTitle = 0
'   .ShowDMDOnly = 1
    .ShowFrame = 0
    .HandleMechanics = 0
    .Hidden = 1
    '.SetDisplayPosition 0,0,GetPlayerHWnd 'if you can't see the DMD then uncomment this line
    On Error Resume Next
    Controller.SolMask(0) = 0
    vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the timer to renable all the solenoids after 2 seconds
    Controller.Run GetPlayerHwnd
    If Err Then MsgBox Err.Description
    On Error Goto 0
  End With

  ' Nudging
  vpmNudge.TiltSwitch = 57 'swTilt
  vpmNudge.Sensitivity = 5
  vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, Bumper4, LeftSlingshot, RightSlingshot)

  'Saucers Declaration
  Set bsArmsLock = New cvpmSaucer
  with bsArmsLock
    .InitKicker Armslock, 43, 185, 8, 8 ' LeftKickout '160, 5 ,8
    .InitExitVariance 1, 0
    .InitSounds "kicker_enter", SoundFX("Solenoid",DOFContactors), SoundFX("Popper",DOFContactors)
    .createevents "bsArmsLock", Armslock
  end with

  Set bsLegsLock = New cvpmSaucer
  with bsLegsLock
    .InitKicker Legslock, 73, 346, 18, 8  ' RightKickout '320, 16, 20
    .InitExitVariance 2, 2
    .InitSounds "kicker_enter", SoundFX("Solenoid",DOFContactors), SoundFX("Popper",DOFContactors)
    .createevents "bsLegsLock", LegsLock
  end with


  Set dtM = New cvpmDropTarget
  dtM.InitDrop Array(dt1, dt2, dt3), Array(41, 51, 61)
  dtM.InitSnd SoundFX("droptarget",DOFContactors), SoundFX("resetdrop",DOFContactors)
  dtM.CreateEvents "dtM"


  'Trough Declaration
  Set bsTrough = New cvpmBallStack
  bsTrough.InitSw 55, 0, 74, 0, 0, 0, 0, 0
  bsTrough.InitKick Ballrelease, 80, 6
  bsTrough.InitExitSnd SoundFX("ballrel",DOFContactors), SoundFX("Solenoid",DOFContactors)
  bsTrough.Balls = 2

    ' New style Trough that didn't quite work out
'    Set bsTrough = New cvpmTrough
'    With bsTrough
'        .size = 2
'        .initSwitches Array(55, 74)
'        .Initexit BallRelease, 80, 6
'        .InitEntrySounds "drain", "Solenoid", "Solenoid"
'        .InitExitSounds "Solenoid", "ballrel"
'        .Balls = 2
'    End With

  'Init Target Walls animation
  RightKick.IsDropped = 1:LeftKick.IsDropped = 1
  RightKick2.IsDropped = 0:LeftKick2.IsDropped = 0

  'Init Robot Lights
  StopRobotLights

  'Other variables
  Last12 = 0
  Current12 = 0
  Last13 = 0
  Current13 = 0
  Last14 = 0
  Current14 = 0

  PinMAMETimer.Interval = PinMAMEInterval
  PinMAMETimer.Enabled = 1
  'StartShake

  SolGI 0:SolLeft 0:SolRight 0

  'display Option
  if not DesktopMode and fullscreendisplay <> 2 then
  For each xx in Display:xx.visible = 0: Next
  Displaytimer.enabled = 0
  end If

setup_backglass()

TextBox001.visible = 0

End Sub


Dim dloop: dloop =0

sub delay()
Dim nloop

for nloop =0 to 10
  nloop = nloop + 1
  Next
dloop = dloop +1
lalocks()
end sub


Sub StartTest ()
Test.enabled = 1
End Sub

'Legs + Arms Lock
Sub lalocks()

vpmTimer.PulseSw 43
vpmTimer.PulseSw 73
end Sub

Sub Test_Timer ()

  'If l24.state = 1 then
  If bl24State = 1 then
  me.enabled = 0
  'exit sub
  else
  'Left Lane brain
    if dloop = 0 then
    Controller.switch(53) = 1
    delay
    elseif dloop = 1 then
    Controller.switch(53) = 0
    delay
    elseif dloop = 2 then
    Controller.switch(53) = 1
    delay
    elseif dloop = 3 then
    Controller.switch(53) = 0
    delay
    elseif dloop = 4 then
    Controller.switch(53) = 1
    delay
    elseif dloop = 5 then
    Controller.switch(53) = 0
    delay
    elseif dloop = 6 then
    Controller.switch(53) = 1
    delay
    elseif dloop = 7 then
    Controller.switch(53) = 0
    delay
    elseif dloop = 8 then


    'Body
    Controller.switch(42) = 1
    delay
    elseif dloop = 9 then
    Controller.switch(42) = 0
    delay
    elseif dloop = 10 then
    Controller.switch(52) = 1
    delay
    elseif dloop = 11 then
    Controller.switch(52) = 0
    delay
    elseif dloop = 12 then
    Controller.switch(62) = 1
    delay
    elseif dloop = 13 then
    Controller.switch(62) = 0
    delay
    elseif dloop = 14 then
    Controller.switch(42) = 1
    delay
    elseif dloop = 15 then
    Controller.switch(42) = 0
    delay
    elseif dloop = 16 then


    'right lane sw63 brain
    Controller.switch(63) = 1
    delay
    elseif dloop = 17 then
    Controller.switch(63) = 0
    delay
    elseif dloop = 18 then
    Controller.switch(63) = 1
    delay
    elseif dloop = 19 then
    Controller.switch(63) = 0
    delay
    elseif dloop = 20 then
    Controller.switch(63) = 1
    delay
    elseif dloop = 21 then
    Controller.switch(63) = 0
    delay
    elseif dloop = 22 then
    Controller.switch(63) = 1
    delay
    elseif dloop = 23 then
    Controller.switch(63) = 0
    delay
    dloop =0
    end if 'dloop
  end if

End Sub

' ***************************************************************************
'                    BASIC FSS(DMD,SS,EM) SETUP CODE
' ****************************************************************************

const USEEM = 0
const USESS = 0
Dim xoff,yoff,zoff,xrot,zscale, xcen,ycen

sub setup_backglass()

xoff =475
yoff =0
zoff =650
xrot = -90


bgdark.x = xoff
bgdark.y = yoff
bgdark.height = zoff
bgdark.rotx = xrot

bgHigh.x = xoff
bgHigh.y = yoff
bgHigh.height = zoff
bgHigh.rotx = xrot

bgHigh1.x = xoff
bgHigh1.y = yoff
bgHigh1.height = zoff
bgHigh1.rotx = xrot

bgHigh2.x = xoff
bgHigh2.y = yoff
bgHigh2.height = zoff
bgHigh2.rotx = xrot

bgGrill.x = xoff
bgGrill.y = yoff
bgGrill.height = zoff
bgGrill.rotx = xrot

bgFrame.x = xoff
bgFrame.y = yoff
bgFrame.height = zoff+8
bgFrame.rotx = xrot

bgFrameMask.x = xoff
bgFrameMask.y = yoff
bgFrameMask.height = zoff+8
bgFrameMask.rotx = xrot

bgFrameMaskFill.x = xoff
bgFrameMaskFill.y = yoff
bgFrameMaskFill.height = zoff
bgFrameMaskFill.rotx = xrot

if cUSEBGLASSREFL then

' the mask fill image
if NOT IsEmpty(tMaskFill(0)) then
BGFrameMaskFill1.x = xoff
BGFrameMaskFill1.y = yoff
BGFrameMaskFill1.height = zoff
BGFrameMaskFill1.rotx = xrot
end If
' the background reflective image usually __default_screen_space_reflection.png
if NOT IsEmpty(tMaskFill(4)) then
BGFrameMaskFill5.x = xoff
BGFrameMaskFill5.y = yoff
BGFrameMaskFill5.height = zoff-100 ' adjust value to correct vertical
BGFrameMaskFill5.rotx = xrot
end if
end if 'cUSEBGLASSREFL

if cUSEBGLASSCCX then
FlBGRedGreenSpec.x = xoff
FlBGRedGreenSpec.y = yoff
FlBGRedGreenSpec.height = zoff
FlBGRedGreenSpec.rotx = xrot

FlBGBlueSpec.x = xoff
FlBGBlueSpec.y = yoff
FlBGBlueSpec.height = zoff
FlBGBlueSpec.rotx = xrot
end if ' cUSEBGLASSCCX


center_graphix()
center_digits()


end sub


Dim BGArr
BGArr=Array (BGFlasher0,BGFlasher001,BGFlasher002,BGFlasher003,BGFlasher004,BGFlasher005,_
BGFlasher006,BGFlasher007,BGFlasher008,BGFlasher009,BGFlasher010,BGFlasher011,BGFlasher012,_
BGFlasher013,BGFlasher014)


Sub center_graphix()
Dim xx,yy,yfact,xfact,xobj
zscale = 0.0000001

xcen =(977 /2) - (52 / 2)
ycen = (1000 /2 ) + (200 /2)


yfact =0 'y fudge factor (ycen was wrong so fix)
xfact =0

  For Each xobj In BGArr
  xx =xobj.x

  xobj.x = (xoff -xcen) + xx +xfact
  yy = xobj.y ' get the yoffset before it is changed
  xobj.y =yoff

    If(yy < 0.) then
    yy = yy * -1
    end if


  xobj.height =( zoff - ycen) + yy - (yy * zscale) + yfact

  xobj.rotx = xrot
  xobj.visible =1 ' for testing
  Next
end sub


Sub center_digits()

zscale = 0.0000001

xcen =(977 /2) - (52 / 2)
ycen = (1000 /2 ) + (200 /2)
Dim ii
Dim xx
Dim yy
Dim yfact
Dim xfact
Dim obj
yfact =0 'y fudge factor (ycen was wrong so fix)
xfact =0


for ii =0 to 39
  For Each obj In Digits(ii)
  xx =obj.x

  obj.x = (xoff -xcen) + xx +xfact
  yy = obj.y ' get the yoffset before it is changed
  obj.y =yoff

    If(yy < 0.) then
    yy = yy * -1
    end if

  obj.height =( zoff - ycen) + yy - (yy * zscale) + yfact

  obj.rotx = xrot
  Next
  Next
end sub


Dim Digits(40)
 Digits(0)=Array(ax00, ax05, ax0c, ax0d, ax08, ax01, ax06, ax0f, ax02, ax03, ax04, ax07, ax0b, ax0a, ax09, ax0e)
 Digits(1)=Array(ax10, ax15, ax1c, ax1d, ax18, ax11, ax16, ax1f, ax12, ax13, ax14, ax17, ax1b, ax1a, ax19, ax1e)
 Digits(2)=Array(ax20, ax25, ax2c, ax2d, ax28, ax21, ax26, ax2f, ax22, ax23, ax24, ax27, ax2b, ax2a, ax29, ax2e)
 Digits(3)=Array(ax30, ax35, ax3c, ax3d, ax38, ax31, ax36, ax3f, ax32, ax33, ax34, ax37, ax3b, ax3a, ax39, ax3e)
 Digits(4)=Array(ax40, ax45, ax4c, ax4d, ax48, ax41, ax46, ax4f, ax42, ax43, ax44, ax47, ax4b, ax4a, ax49, ax4e)
 Digits(5)=Array(ax50, ax55, ax5c, ax5d, ax58, ax51, ax56, ax5f, ax52, ax53, ax54, ax57, ax5b, ax5a, ax59, ax5e)
 Digits(6)=Array(ax60, ax65, ax6c, ax6d, ax68, ax61, ax66, ax6f, ax62, ax63, ax64, ax67, ax6b, ax6a, ax69, ax6e)
 Digits(7)=Array(ax70, ax75, ax7c, ax7d, ax78, ax71, ax76, ax7f, ax72, ax73, ax74, ax77, ax7b, ax7a, ax79, ax7e)
 Digits(8)=Array(ax80, ax85, ax8c, ax8d, ax88, ax81, ax86, ax8f, ax82, ax83, ax84, ax87, ax8b, ax8a, ax89, ax8e)
 Digits(9)=Array(ax90, ax95, ax9c, ax9d, ax98, ax91, ax96, ax9f, ax92, ax93, ax94, ax97, ax9b, ax9a, ax99, ax9e)
 Digits(10)=Array(axa0, axa5, axac, axad, axa8, axa1, axa6, axaf, axa2, axa3, axa4, axa7, axab, axaa, axa9, axae)
 Digits(11)=Array(axb0, axb5, axbc, axbd, axb8, axb1, axb6, axbf, axb2, axb3, axb4, axb7, axbb, axba, axb9, axbe)
 Digits(12)=Array(axc0, axc5, axcc, axcd, axc8, axc1, axc6, axcf, axc2, axc3, axc4, axc7, axcb, axca, axc9, axce)
 Digits(13)=Array(axd0, axd5, axdc, axdd, axd8, axd1, axd6, axdf, axd2, axd3, axd4, axd7, axdb, axda, axd9, axde)
 Digits(14)=Array(axe0, axe5, axec, axed, axe8, axe1, axe6, axef, axe2, axe3, axe4, axe7, axeb, axea, axe9, axee)
 Digits(15)=Array(axf0, axf5, axfc, axfd, axf8, axf1, axf6, axff, axf2, axf3, axf4, axf7, axfb, axfa, axf9, axfe)
 Digits(16)=Array(bx00, bx05, bx0c, bx0d, bx08, bx01, bx06, bx0f, bx02, bx03, bx04, bx07, bx0b, bx0a, bx09, bx0e)
 Digits(17)=Array(bx10, bx15, bx1c, bx1d, bx18, bx11, bx16, bx1f, bx12, bx13, bx14, bx17, bx1b, bx1a, bx19, bx1e)
 Digits(18)=Array(bx20, bx25, bx2c, bx2d, bx28, bx21, bx26, bx2f, bx22, bx23, bx24, bx27, bx2b, bx2a, bx29, bx2e)
 Digits(19)=Array(bx30, bx35, bx3c, bx3d, bx38, bx31, bx36, bx3f, bx32, bx33, bx34, bx37, bx3b, bx3a, bx39, bx3e)

 Digits(20)=Array(bx40, bx45, bx4c, bx4d, bx48, bx41, bx46, bx4f, bx42, bx43, bx44, bx47, bx4b, bx4a, bx49, bx4e)
 Digits(21)=Array(bx50, bx55, bx5c, bx5d, bx58, bx51, bx56, bx5f, bx52, bx53, bx54, bx57, bx5b, bx5a, bx59, bx5e)
 Digits(22)=Array(bx60, bx65, bx6c, bx6d, bx68, bx61, bx66, bx6f, bx62, bx63, bx64, bx67, bx6b, bx6a, bx69, bx6e)
 Digits(23)=Array(bx70, bx75, bx7c, bx7d, bx78, bx71, bx76, bx7f, bx72, bx73, bx74, bx77, bx7b, bx7a, bx79, bx7e)
 Digits(24)=Array(bx80, bx85, bx8c, bx8d, bx88, bx81, bx86, bx8f, bx82, bx83, bx84, bx87, bx8b, bx8a, bx89, bx8e)
 Digits(25)=Array(bx90, bx95, bx9c, bx9d, bx98, bx91, bx96, bx9f, bx92, bx93, bx94, bx97, bx9b, bx9a, bx99, bx9e)
 Digits(26)=Array(bxa0, bxa5, bxac, bxad, bxa8, bxa1, bxa6, bxaf, bxa2, bxa3, bxa4, bxa7, bxab, bxaa, bxa9, bxae)
 Digits(27)=Array(bxb0, bxb5, bxbc, bxbd, bxb8, bxb1, bxb6, bxbf, bxb2, bxb3, bxb4, bxb7, bxbb, bxba, bxb9, bxbe)
 Digits(28)=Array(bxc0, bxc5, bxcc, bxcd, bxc8, bxc1, bxc6, bxcf, bxc2, bxc3, bxc4, bxc7, bxcb, bxca, bxc9, bxce)
 Digits(29)=Array(bxd0, bxd5, bxdc, bxdd, bxd8, bxd1, bxd6, bxdf, bxd2, bxd3, bxd4, bxd7, bxdb, bxda, bxd9, bxde)
 Digits(30)=Array(bxe0, bxe5, bxec, bxed, bxe8, bxe1, bxe6, bxef, bxe2, bxe3, bxe4, bxe7, bxeb, bxea, bxe9, bxee)
 Digits(31)=Array(bxf0, bxf5, bxfc, bxfd, bxf8, bxf1, bxf6, bxff, bxf2, bxf3, bxf4, bxf7, bxfb, bxfa, bxf9, bxfe)


 Digits(32)=Array(byg1, byg2, byg3 ,byg4, byg5, byg6, byg7, bygg, byg8, byg9, byga, bygb, bygc, bygd, byge, bygf)
 Digits(33)=Array(byh1, byh2, byh3 ,byh4, byh5, byh6, byh7, byhg, byh8, byh9, byha, byhb, byhc, byhd, byhe, byhf)
 Digits(34)=Array(bya1, bya2, bya3, bya4, bya5, bya6, bya7, byag, bya8, bya9, byaa, byab, byac, byad, byae, byaf)
 Digits(35)=Array(byb1, byb2, byb3, byb4, byb5, byb6, byb7, bybg, byb8, byb9, byba, bybb, bybc, bybd, bybe, bybf)
 Digits(36)=Array(byc1, byc2, byc3, byc4, byc5, byc6, byc7, bycg, byc8, byc9, byca, bycb, bycc, bycd, byce, bycf)
 Digits(37)=Array(byd1, byd2, byd3, byd4, byd5, byd6, byd7, bydg, byd8, byd9, byda, bydb, bydc, bydd, byde, bydf)
 Digits(38)=Array(bye1, bye2, bye3, bye4, bye5, bye6, bye7, byeg, bye8, bye9, byea, byeb, byec, byed, byee, byef)
 Digits(39)=Array(byf1, byf2, byf3, byf4, byf5, byf6, byf7, byfg, byf8, byf9, byfa, byfb, byfc, byfd, byfe, byff)


 Sub DisplayTimer_Timer
    Dim ChgLED, ii, jj, num, chg, stat, obj, b, x
    ChgLED=Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
    'If DesktopMode = True Then
       For ii=0 To UBound(chgLED)
          num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)
      'if (num < 32) then
      if (num < 41) then
              For Each obj In Digits(num)
                   If chg And 1 Then obj.visible=stat And 1
                   chg=chg\2 : stat=stat\2
                  Next
      Else
             end if
        Next
     'end if
    End If
 End Sub


Sub Genesis_Paused:Controller.Pause = 1:End Sub
Sub Genesis_unPaused:Controller.Pause = 0:End Sub

'***********************
' keys
'***********************



Sub Genesis_KeyDown(ByVal keycode)
  If keycode = PlungerKey Then Plunger.Pullback
  If keycode = LeftTiltKey Then PlaySound SoundFX("nudge_left",0)
  If keycode = RightTiltKey Then PlaySound SoundFX("nudge_right",0)
  If keycode = CenterTiltKey Then PlaySound SoundFX("nudge_forward",0)

  If keycode = LeftFlipperKey Then LFPress = 1
  If keycode = RightFlipperKey Then rfpress = 1

  If vpmKeyDown(keycode) Then Exit Sub
End Sub


Sub Genesis_KeyUp(ByVal keycode)
  If keycode = PlungerKey Then
    Plunger.Fire
    If(BallinPlunger = 1) then 'the ball is in the plunger lane
      PlaySound "Plunger2"
    else
      PlaySound "Plunger"
    end if
  End If

'nfozzy physics'
  If keycode = LeftFlipperKey Then
    lfpress = 0
    leftflipper.eostorqueangle = EOSA
    leftflipper.eostorque = EOST
  End If
  If keycode = RightFlipperKey Then
    rfpress = 0
    rightflipper.eostorqueangle = EOSA
    rightflipper.eostorque = EOST
  End If

  If vpmKeyUp(KeyCode) Then Exit Sub
End Sub

'***********************
'JP's Alpha Ramp Plunger
'***********************
Dim BallinPlunger

Sub swPlunger_Hit:BallinPlunger = 1:End Sub                            'in this sub you may add a switch, for example Controller.Switch(14) = 1

Sub swPlunger_UnHit:BallinPlunger = 0:End Sub                          'in this sub you may add a switch, for example Controller.Switch(14) = 0


'*******************
'Solenoids Callback
'*******************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

'SolCallback(1) = "" 'Varitarget ??
SolCallback(2) = "bsLegsLock.SolOut"
SolCallback(4) = "SolLeft"
SolCallback(5) = "bsArmsLock.SolOut"
SolCallback(6) = "DropDelaysub"
'SolCallback(6) = "dtM.SolDropUp"
SolCallback(7) = "SolRight"
SolCallback(8) = "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
'SolCallback(8) = "VpmSolSound""knocker"","
SolCallback(9) = "SolOuthole"
SolCallback(10) = "SolGI"

'dim Drop targets
dim drop1, drop2, drop3
drop1 = dt1.isdropped
drop2 = dt2.isdropped
drop3 = dt3.isdropped



'Drop Delay
Sub DropDelaysub(enabled)
  If Dropfix = 1 then
    DropDelay.Enabled = 1
  Else
    dtM.DropSol_On
    drop1 = 0
    drop2 = 0
    drop3 = 0
    updateGI
  End If
End Sub

Sub DropDelay_Timer()
  dtM.DropSol_On
  me.interval = 20
  me.enabled = 0
  drop1 = 0
  drop2 = 0
  drop3 = 0
  updateGI
End Sub

'Solenoids Subs

Sub SolGI(Enabled)
  If Enabled Then
' textbox1.text = Enabled
    gi1.state=1:gi2.state=1:gi3.state=1:gi4.state=1:gi5.state=1:gi6.state=1:gi7.state=1:gi8.state=1:gi9.state=1:gi10.state=1:gi11.state=1:gi12.state=1:gi13.state=1:gi_ambient.state=1:gi15.state=1
    UpdateGi
  Else
' textbox1.text = Enabled
    gi1.state=0:gi2.state=0:gi3.state=0:gi4.state=0:gi5.state=0:gi6.state=0:gi7.state=0:gi8.state=0:gi9.state=0:gi10.state=0:gi11.state=0:gi12.state=0:gi13.state=0:gi_ambient.state=0:gi15.state=0
    UpdateGi
  End If
End Sub

Sub UpdateGI
gi14.state = gi7.state
if drop1 = 1 then gi14_1.state = gi7.state else gi14_1.state = 0 end if
if drop2 = 1 then gi14_2.state = gi7.state else gi14_2.state = 0 end if
if drop3 = 1 then gi14_3.state = gi7.state else gi14_3.state = 0 end if

end sub


'Sub timer1_timer() 'check DTs Debug
'textbox4.text = dt1.isdropped & " " & dt2.isdropped & " " & dt3.isdropped
'textbox5.text = drop1 & " " & drop2 & " " & drop3
'updategi

'End Sub

Sub SolLeft(Enabled)
  If Enabled Then
    Fl2.state = 2:fl3.state = 2:playsound "lswitch", 0, 0.01, 0, 0.1
'   textbox1.text = "ON"
  Else
    Fl2.state=0:fl3.state=0
'   textbox1.text = "OFF"
  End If
End Sub



Sub SolRight(Enabled)
  If Enabled Then
    fr2.state = 2
    fr3.State = 2
    playsound "lswitch", 0, 0.001, 0
'   textbox2.text = "ON"
  Else
    fr2.state = 0
    fr3.State = 0
'   textbox2.text = "OFF"
  End If
End Sub

Sub SolOuthole(enabled)
  if enabled then
    bsTrough.EntrySol_On
'   bsTrough.ExitSol_On
  end if
End Sub

'*************
'Robots Lights
'*************

Dim RobotLightStep, RobotLightsOn, EndIt

RobotLightStep = 0:RobotLightsOn = 0

Sub StartRobotLights
  ll1.state=2:rl1.state=2:ll2.state=2:rl2.state=2:ll3.state=2:rl3.state=2:ll4.state=2:rl4.state=2:ll5.state=2:rl5.state=2
  cl1.state = 2: cl2.state = 2: cl3.state = 2: cl4.state = 2: cl5.state = 2
End Sub

'Sub maybestoprobotlights ' I think this prevents the lightshow from ending early during the robot reveal sequence
' If CurrentRot=0 then LightSeqTimer.Enabled=1 End If 'lightseqtimer judges if the lights should be on or not..
' If CurrentRot<0 then
' StopRobotLights
' End If
'end Sub


Sub StopRobotLights
  ll1.state=0:rl1.state=0:ll2.state=0:rl2.state=0:ll3.state=0:rl3.state=0:ll4.state=0:rl4.state=0:ll5.state=0:rl5.state=0
  cl1.state = 0: cl2.state = 0: cl3.state = 0: cl4.state = 0: cl5.state = 0
End Sub

'Sub RobotLights_Timer  'replaced by blink pattern 'interval was 70
' Select Case RobotLightStep
'   Case 0:Ll1.State=1:Rl1.state=1
'   Case 1:Ll2.state=1:Rl2.state=1
'   Case 2:ll1.state=0:rl1.state=0:ll3.state=1:rl3.state=1
'   Case 3:ll2.state=0:rl2.state=0:ll4.state=1:rl4.state=1
'   Case 4:ll3.state=0:rl3.state=0:ll5.state=1:rl5.state=1
'   Case 5:ll4.state=0:rl4.state=0
'   Case 6:ll5.state=0:rl5.state=0
'   Case 66:ll1.state=0:rl1.state=0:ll2.state=0:rl2.state=0:ll3.state=0:rl3.state=0:ll4.state=0:rl4.state=0:ll5.state=0:rl5.state=0
'   Case 67:RobotLights.Enabled=0:RobotLightStep=1
' End Select
' RobotLightStep = RobotLightStep + 1
' If RobotLightStep = 7 Then RobotLightStep = 0
'End Sub

'****************
' Robot Animation

'****************
dim bDoAnimation: bDoAnimation = 0
dim StartRotation
dim EndRotation
Dim CurrentRot

StartRotation=0
EndRotation=-360

Sub StartRobotAnimation
  CurrentRot=0
  bDoAnimation = 0
  SpinTimer.Enabled=1
  'MsgBox "Start Robot Ani"
End Sub

Sub SpinTimer_Timer()
  If Currentrot=EndRotation then
    currentrot=StartRotation 'back to 0
    me.Enabled=0
    StopRobotLights
    'MsgBox "End Animation"
    Exit Sub
  End If

  If currentrot > EndRotation then
  currentrot=currentrot-0.25
  End If

  Goldy.roty=CurrentRot
  Goldy2.roty=CurrentRot
End Sub

' ****************************************************
' FLIPPERS with Multiple Flip Up/Flip Down Sounds
' ****************************************************

Const ReflipAngle = 20

Sub SolLFlipper(Enabled)
    If Enabled Then
    If leftflipper.currentangle < leftflipper.endangle + ReflipAngle Then
            PlaySoundAtVol SoundFX("TOM_Calle_ReFlip_L0" & Int(Rnd*3)+1, DOFFlippers), LeftFlipper, VolFlip
    Else
      PlaySoundAtVol SoundFX("TOM_Calle_Flipper_Attack-L01", DOFFlippers), LeftFlipper, VolFlip
            PlaySoundAtVol SoundFX("TOM_Calle_Flipper_L0" & Int(Rnd*9)+1, DOFFlippers), LeftFlipper, VolFlip
    End If
    LF.Fire
    Else
        PlaySoundAtVol SoundFX("WD_TOM_Flipper_Left_Down_" & Int(Rnd*7)+1, DOFFlippers), LeftFlipper, VolFlip
    LeftFlipper.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
    If rightflipper.currentangle < rightflipper.endangle + ReflipAngle Then
            PlaySoundAtVol SoundFX("TOM_Calle_ReFlip_R0" & Int(Rnd*3)+1, DOFFlippers), RightFlipper, VolFlip
    Else
      PlaySoundAtVol SoundFX("TOM_Calle_Flipper_Attack-R01", DOFFlippers), RightFlipper, VolFlip
            PlaySoundAtVol SoundFX("TOM_Calle_Flipper_R0" & Int(Rnd*11)+1, DOFFlippers), RightFlipper, VolFlip
    End If
    RF.Fire
    Else
        PlaySoundAtVol SoundFX("WD_TOM_Flipper_Right_Down_" & Int(Rnd*8)+1, DOFFlippers), RightFlipper, VolFlip
    RightFlipper.RotateToStart
    End If
End Sub


'******************************************************
'       FLIPPER AND RUBBER CORRECTION
'******************************************************

dim LF : Set LF = New FlipperPolarity
dim RF : Set RF = New FlipperPolarity

InitPolarity

Sub InitPolarity()
  dim x, a : a = Array(LF, RF)
  for each x in a
    'safety coefficient (diminishes polarity correction only)
    'x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 0  'don't mess with these
    x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
    x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1

    x.enabled = True
    'x.DebugOn = True : stickL.visible = True : tbpl.visible = True : vpmSolFlipsTEMP.DebugOn = True
    x.TimeDelay = 60
  Next

  'rf.report "Velocity"
  addpt "Velocity", 0, 0,   1
  addpt "Velocity", 1, 0.1,   1.07
  addpt "Velocity", 2, 0.2,   1.15
  addpt "Velocity", 3, 0.3,   1.25
  addpt "Velocity", 4, 0.41, 1.05
  addpt "Velocity", 5, 0.65,  1.0'0.982
  addpt "Velocity", 6, 0.702, 0.968
  addpt "Velocity", 7, 0.95,  0.968
  addpt "Velocity", 8, 1.03,  0.945

  AddPt "Polarity", 0, 0, 0
  AddPt "Polarity", 1, 0.05, -5.5
  AddPt "Polarity", 2, 0.4, -5.5
  AddPt "Polarity", 3, 0.8, -5.5
  AddPt "Polarity", 4, 0.85, -5.25
  AddPt "Polarity", 5, 0.9, -4.25
  AddPt "Polarity", 6, 0.95, -3.75
  AddPt "Polarity", 7, 1, -3.25
  AddPt "Polarity", 8, 1.05, -2.25
  AddPt "Polarity", 9, 1.1, -1.5
  AddPt "Polarity", 10, 1.15, -1
  AddPt "Polarity", 11, 1.2, -0.5
  AddPt "Polarity", 12, 1.25, 0
  AddPt "Polarity", 13, 1.3, 0

  LF.Object = LeftFlipper
  LF.EndPoint = EndPointLp  'you can use just a coordinate, or an object with a .x property. Using a couple of simple primitive objects
  RF.Object = RightFlipper
  RF.EndPoint = EndPointRp
End Sub

Sub AddPt(aStr, idx, aX, aY)  'debugger wrapper for adjusting flipper script in-game
  dim a : a = Array(LF, RF)
  dim x : for each x in a
    x.addpoint aStr, idx, aX, aY
  Next
End Sub

'Trigger Hit - .AddBall activeball
'Trigger UnHit - .PolarityCorrect activeball

Sub TriggerLF_Hit() : LF.Addball activeball : End Sub
Sub TriggerLF_UnHit() : LF.PolarityCorrect activeball : End Sub
Sub TriggerRF_Hit() : RF.Addball activeball : End Sub
Sub TriggerRF_UnHit() : RF.PolarityCorrect activeball : End Sub

'Methods:
'.TimeDelay - Delay before trigger shuts off automatically. Default = 80 (ms)
'.AddPoint - "Polarity", "Velocity", "Ycoef" coordinate points. Use one of these 3 strings, keep coordinates sequential. x = %position on the flipper, y = output
'.Object - set to flipper reference. Optional.
'.StartPoint - set start point coord. Unnecessary, if .object is used.

'Called with flipper -
'ProcessBalls - catches ball data.
' - OR -
'.Fire - fires flipper.rotatetoend automatically + processballs. Requires .Object to be set to the flipper.

Class FlipperPolarity
  Public DebugOn, Enabled
  Private FlipAt  'Timer variable (IE 'flip at 723,530ms...)
  Public TimeDelay  'delay before trigger turns off and polarity is disabled TODO set time!
  private Flipper, FlipperStart,FlipperEnd, FlipperEndY, LR, PartialFlipCoef
  Private Balls(20), balldata(20)

  dim PolarityIn, PolarityOut
  dim VelocityIn, VelocityOut
  dim YcoefIn, YcoefOut
  Public Sub Class_Initialize
    redim PolarityIn(0) : redim PolarityOut(0) : redim VelocityIn(0) : redim VelocityOut(0) : redim YcoefIn(0) : redim YcoefOut(0)
    Enabled = True : TimeDelay = 50 : LR = 1:  dim x : for x = 0 to uBound(balls) : balls(x) = Empty : set Balldata(x) = new SpoofBall : next
  End Sub

  Public Property let Object(aInput) : Set Flipper = aInput : StartPoint = Flipper.x : End Property
  Public Property Let StartPoint(aInput) : if IsObject(aInput) then FlipperStart = aInput.x else FlipperStart = aInput : end if : End Property
  Public Property Get StartPoint : StartPoint = FlipperStart : End Property
  Public Property Let EndPoint(aInput) : FlipperEnd = aInput.x: FlipperEndY = aInput.y: End Property
  Public Property Get EndPoint : EndPoint = FlipperEnd : End Property
  Public Property Get EndPointY: EndPointY = FlipperEndY : End Property


  Public Sub AddPoint(aChooseArray, aIDX, aX, aY) 'Index #, X position, (in) y Position (out)
    Select Case aChooseArray
      case "Polarity" : ShuffleArrays PolarityIn, PolarityOut, 1 : PolarityIn(aIDX) = aX : PolarityOut(aIDX) = aY : ShuffleArrays PolarityIn, PolarityOut, 0
      Case "Velocity" : ShuffleArrays VelocityIn, VelocityOut, 1 :VelocityIn(aIDX) = aX : VelocityOut(aIDX) = aY : ShuffleArrays VelocityIn, VelocityOut, 0
      Case "Ycoef" : ShuffleArrays YcoefIn, YcoefOut, 1 :YcoefIn(aIDX) = aX : YcoefOut(aIDX) = aY : ShuffleArrays YcoefIn, YcoefOut, 0
    End Select
    if gametime > 100 then Report aChooseArray
  End Sub

  Public Sub Report(aChooseArray)   'debug, reports all coords in tbPL.text
    if not DebugOn then exit sub
    dim a1, a2 : Select Case aChooseArray
      case "Polarity" : a1 = PolarityIn : a2 = PolarityOut
      Case "Velocity" : a1 = VelocityIn : a2 = VelocityOut
      Case "Ycoef" : a1 = YcoefIn : a2 = YcoefOut
      case else :tbpl.text = "wrong string" : exit sub
    End Select
    dim str, x : for x = 0 to uBound(a1) : str = str & aChooseArray & " x: " & round(a1(x),4) & ", " & round(a2(x),4) & vbnewline : next
    tbpl.text = str
  End Sub

  Public Sub AddBall(aBall) : dim x : for x = 0 to uBound(balls) : if IsEmpty(balls(x)) then set balls(x) = aBall : exit sub :end if : Next  : End Sub

  Private Sub RemoveBall(aBall)
    dim x : for x = 0 to uBound(balls)
      if TypeName(balls(x) ) = "IBall" then
        if aBall.ID = Balls(x).ID Then
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
    dim x : for x = 0 to uBound(balls)
      if not IsEmpty(balls(x) ) then
        pos = pSlope(Balls(x).x, FlipperStart, 0, FlipperEnd, 1)
      End If
    Next
  End Property

  Public Sub ProcessBalls() 'save data of balls in flipper range
    FlipAt = GameTime
    dim x : for x = 0 to uBound(balls)
      if not IsEmpty(balls(x) ) then
        balldata(x).Data = balls(x)
      End If
    Next
    PartialFlipCoef = ((Flipper.StartAngle - Flipper.CurrentAngle) / (Flipper.StartAngle - Flipper.EndAngle))
    PartialFlipCoef = abs(PartialFlipCoef-1)
  End Sub
  Private Function FlipperOn() : if gameTime < FlipAt+TimeDelay then FlipperOn = True : End If : End Function 'Timer shutoff for polaritycorrect

  Public Sub PolarityCorrect(aBall)
    if FlipperOn() then
      dim tmp, BallPos, x, IDX, Ycoef : Ycoef = 1

      'y safety Exit
      if aBall.VelY > -8 then 'ball going down
        RemoveBall aBall
        exit Sub
      end if

      'Find balldata. BallPos = % on Flipper
      for x = 0 to uBound(Balls)
        if aBall.id = BallData(x).id AND not isempty(BallData(x).id) then
          idx = x
          BallPos = PSlope(BallData(x).x, FlipperStart, 0, FlipperEnd, 1)
          if ballpos > 0.65 then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)        'find safety coefficient 'ycoef' data
        end if
      Next

      If BallPos = 0 Then 'no ball data meaning the ball is entering and exiting pretty close to the same position, use current values.
        BallPos = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
        if ballpos > 0.65 then  Ycoef = LinearEnvelope(aBall.Y, YcoefIn, YcoefOut)            'find safety coefficient 'ycoef' data
      End If

      'Velocity correction
      if not IsEmpty(VelocityIn(0) ) then
        Dim VelCoef
   :      VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)

        if partialflipcoef < 1 then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)

        if Enabled then aBall.Velx = aBall.Velx*VelCoef
        if Enabled then aBall.Vely = aBall.Vely*VelCoef
      End If

      'Polarity Correction (optional now)
      if not IsEmpty(PolarityIn(0) ) then
        If StartPoint > EndPoint then LR = -1 'Reverse polarity if left flipper
        dim AddX : AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR

        if Enabled then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef)
        'debug.print BallPos & " " & AddX & " " & Ycoef & " "& PartialFlipcoef & " "& VelCoef
        'playsound "fx_knocker"
      End If
    End If
    RemoveBall aBall
  End Sub
End Class


'================================
'Helper Functions


Sub ShuffleArray(ByRef aArray, byVal offset) 'shuffle 1d array
  dim x, aCount : aCount = 0
  redim a(uBound(aArray) )
  for x = 0 to uBound(aArray) 'Shuffle objects in a temp array
    if not IsEmpty(aArray(x) ) Then
      if IsObject(aArray(x)) then
        Set a(aCount) = aArray(x)
      Else
        a(aCount) = aArray(x)
      End If
      aCount = aCount + 1
    End If
  Next
  if offset < 0 then offset = 0
  redim aArray(aCount-1+offset) 'Resize original array
  for x = 0 to aCount-1   'set objects back into original array
    if IsObject(a(x)) then
      Set aArray(x) = a(x)
    Else
      aArray(x) = a(x)
    End If
  Next
End Sub

Sub ShuffleArrays(aArray1, aArray2, offset)
  ShuffleArray aArray1, offset
  ShuffleArray aArray2, offset
End Sub


Function BallSpeed(ball) 'Calculates the ball speed
    BallSpeed = SQR(ball.VelX^2 + ball.VelY^2 + ball.VelZ^2)
End Function

Function PSlope(Input, X1, Y1, X2, Y2)  'Set up line via two points, no clamping. Input X, output Y
  dim x, y, b, m : x = input : m = (Y2 - Y1) / (X2 - X1) : b = Y2 - m*X2
  Y = M*x+b
  PSlope = Y
End Function

Function NullFunctionZ(aEnabled):End Function '1 argument null function placeholder  TODO move me or replac eme

Class spoofball
  Public X, Y, Z, VelX, VelY, VelZ, ID, Mass, Radius
  Public Property Let Data(aBall)
    With aBall
      x = .x : y = .y : z = .z : velx = .velx : vely = .vely : velz = .velz
      id = .ID : mass = .mass : radius = .radius
    end with
  End Property
  Public Sub Reset()
    x = Empty : y = Empty : z = Empty  : velx = Empty : vely = Empty : velz = Empty
    id = Empty : mass = Empty : radius = Empty
  End Sub
End Class

Function Distance(ax,ay,bx,by)
  Distance = SQR((ax - bx)^2 + (ay - by)^2)
End Function

Function DistancePL(px,py,ax,ay,bx,by) ' Distance between a point and a line where point is px,py
  DistancePL = ABS((by - ay)*px - (bx - ax) * py + bx*ay - by*ax)/Distance(ax,ay,bx,by)
End Function

'///////////////////////////////// Flipper Tricks Physics //////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////

RightFlipper.timerinterval=1
Rightflipper.timerenabled=True

sub RightFlipper_timer()
  FlipperTricks LeftFlipper, LFPress, LFCount, LFEndAngle, LFState
  FlipperTricks RightFlipper, RFPress, RFCount, RFEndAngle, RFState
end sub

dim LFPress, RFPress, LFCount, RFCount
dim LFState, RFState
dim EOST, EOSA,Frampup, FElasticity,FReturn
dim RFEndAngle, LFEndAngle

EOST = leftflipper.eostorque
EOSA = leftflipper.eostorqueangle
Frampup = LeftFlipper.rampup
FElasticity = LeftFlipper.elasticity
FReturn = LeftFlipper.return
Const EOSTnew = 0.8   '1.0 FEOST
Const EOSAnew = 1   '0.2
Const EOSRampup = 0 '0.5
Dim SOSRampup
  Select Case FlipperCoilRampupMode
    Case 0:
      SOSRampup = 2.5
    Case 1:
      SOSRampup = 8.5
  End Select
Const LiveCatch = 8
Const LiveElasticity = 0.45
Const SOSEM = 0.815
Const EOSReturn = 0.025

LFEndAngle = Leftflipper.endangle
RFEndAngle = RightFlipper.endangle

Sub FlipperActivate(Flipper, FlipperPress)
  FlipperPress = 1
  Flipper.Elasticity = FElasticity

  Flipper.eostorque = EOST            'new
  Flipper.eostorqueangle = EOSA         'new
End Sub

Sub FlipperDeactivate(Flipper, FlipperPress)
  FlipperPress = 0
  Flipper.eostorqueangle = EOSA
  Flipper.eostorque = EOST*EOSReturn/FReturn  'EOST


  If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 0.1 Then
    Dim BOT, b
    BOT = GetBalls

    For b = 0 to UBound(BOT)
      If Distance(BOT(b).x, BOT(b).y, Flipper.x, Flipper.y) < 55 Then 'check for cradle
        If BOT(b).vely >= -0.4 Then BOT(b).vely = -0.4
      End If
    Next
  End If
End Sub

Sub FlipperTricks (Flipper, FlipperPress, FCount, FEndAngle, FState)
  Dim Dir
  Dir = Flipper.startangle/Abs(Flipper.startangle)  '-1 for Right Flipper

  If Abs(Flipper.currentangle) > Abs(Flipper.startangle) - 0.05 Then
    If FState <> 1 Then
      Flipper.rampup = SOSRampup
      Flipper.endangle = FEndAngle - 3*Dir
      Flipper.Elasticity = FElasticity * SOSEM
      FCount = 0
      FState = 1
    End If
  ElseIf Abs(Flipper.currentangle) <= Abs(Flipper.endangle) and FlipperPress = 1 then
    if FCount = 0 Then FCount = GameTime

'   if GameTime - FCount < LiveCatch Then
'     Flipper.Elasticity = LiveElasticity
'   elseif GameTime - FCount < LiveCatch * 2 Then
'     Flipper.Elasticity = 0.1
'   Else
'     Flipper.Elasticity = FElasticity
'   end if

    If FState <> 2 Then
      Flipper.eostorqueangle = EOSAnew
      Flipper.eostorque = EOSTnew
      Flipper.rampup = EOSRampup
      Flipper.endangle = FEndAngle
      FState = 2
    End If
  Elseif Abs(Flipper.currentangle) > Abs(Flipper.endangle) + 0.01 and FlipperPress = 1 Then
    If FState <> 3 Then
      Flipper.eostorque = EOST            'EOST
      Flipper.eostorqueangle = EOSA       'EOSA
      Flipper.rampup = Frampup
      Flipper.Elasticity = FElasticity
      FState = 3
    End If

  End If
End Sub

Const LiveDistanceMin = 30  'minimum distance in vp units from flipper base live catch dampening will occur
Const LiveDistanceMax = 114  'maximum distance in vp units from flipper base live catch dampening will occur (tip protection)

Sub CheckLiveCatch(ball, Flipper, FCount, parm) 'Experimental new live catch
  Dim Dir
  Dir = Flipper.startangle/Abs(Flipper.startangle)  '-1 for Right Flipper

  if GameTime - FCount < LiveCatch and parm > 6 and ABS(Flipper.x - ball.x) > LiveDistanceMin and ABS(Flipper.x - ball.x) < LiveDistanceMax Then
    If ball.velx * Dir > 0 Then ball.velx = 0
    ball.vely = 0
    ball.angmomx= 0
    ball.angmomy= 0
    ball.angmomz= 0
  End If
End Sub

Sub LeftFlipper_Collide(parm)
  CheckDampen Activeball, LeftFlipper, parm
  CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
  RandomSoundFlipper()
  'LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
  CheckDampen Activeball, RightFlipper, parm
  CheckLiveCatch Activeball, RightFlipper, RFCount, parm
  RandomSoundFlipper()
  'RightFlipperCollide parm
End Sub

dim angdamp, veldamp
angdamp = 0.2
veldamp = 0.8

Sub CheckDampen(ball, Flipper, parm)
  Dim Dir
  Dir = Flipper.startangle/Abs(Flipper.startangle)  '-1 for Right Flipper

  If parm > 4 and ABS(Flipper.x - ball.x) < LiveDistanceMin And  Flipper.currentangle = Flipper.endangle Then
    ball.angmomx=ball.angmomx*angdamp
    ball.angmomy=ball.angmomy*angdamp
    ball.angmomz=ball.angmomz*angdamp
    If  ball.velx*Dir > 0 Then ball.velx = ball.velx * veldamp
  End If
End Sub

'*****************************************************************************************************
'*******************************************************************************************************
'END nFOZZY FLIPPERS'
'Set MotorCallback = GetRef("RealTimeUpdates")

Sub FlipperTimer_Timer()
  UpdateLeftFlipperLogo
  UpdateRightFlipperLogo
End Sub

Sub UpdateLeftFlipperLogo()
  LFLogo.RotY = LeftFlipper.CurrentAngle
End Sub
Sub UpdateRightFlipperLogo()
  RFLogo.RotY = RightFlipper.CurrentAngle
End Sub


'************
'Varitarget
'omg
'************

'Varitarget primitive version:
'range: all the way forward rotX -9 & -128
'all the way back rotx 12 & -20
dim variball

dim Y 'variangle
dim CFA 'FlipperAngle
CFA = Flipper2.currentangle

Y = -9

sub Invari_hit()
  me.timerenabled = 0
  me.timerinterval = 500
  VariChecker.enabled = 1
end sub

sub Invari_unhit()
  me.timerinterval = 500
  me.timerenabled = 1
' VariChecker.enabled = 0

end sub

sub invari_timer()
  if y < -8.5 then varichecker.enabled = 0: me.timerenabled = 0 end if
end sub


'sub  set variball = BallcntOver
Sub Varichecker_timer()
  CFA = Flipper2.currentangle
  y = ((7 * CFA) / 36) + (143 / 9)
  Varitarget.rotX = Y
' textbox2.text = Y
' textbox3.text = flipper2.currentangle & flipper2.startangle & flipper2.endangle


'I am bad at maths


End sub

dim v1, v2, v3, v4
v1 = 0:v2 = 0:v3=0:v4=0
'Varitimer is 200ms


Sub Varitarget1_Hit
  If ActiveBall.VelY <0 Then
    Controller.Switch(40) = 1
    V1 = 1
    V2 = 0
  end if
End Sub

Sub Varitarget1_UnHit
  If ActiveBall.VelY> 0 Then VariTimer.Enabled = 1
End Sub


Sub Varitarget2_Hit
  If ActiveBall.VelY <0 Then
    Controller.Switch(50) = 1
    V2 = 1
    V3 = 0
  end if
End Sub

Sub Varitarget3_Hit
  If ActiveBall.VelY <0 Then
    Controller.Switch(60) = 1
    V3 = 1
    V4 = 0
  end if
End Sub

Sub Varitarget4_Hit
  Controller.Switch(70) = 1
    if bl24State = 1 then
    bDoAnimation = 1
    end if
End Sub

Sub VariTimer_Timer
  If v4 = 0 Then
    Controller.Switch(70) = 0
    V4 = 1
    V3 = 0
    Exit Sub
  End If

  If v3 = 0 Then
    Controller.Switch(60) = 0
    V3 = 1
    V2 = 0
    Exit Sub
  End If

  If v2 = 0 Then
    Controller.Switch(50) = 0
    V2 = 1
    V1 = 0
    Exit Sub
  End If

  Controller.Switch(40) = 0
  VariTimer.Enabled = 0
End Sub


'Triggers

Sub LeftKick_Timer:LeftKick.TimerEnabled = 0:LeftKick2.IsDropped = 0:LeftKick.IsDropped = 1:End Sub

Sub LeftKick2_Slingshot():vpmTimer.PulseSw(45):LeftKick.IsDropped = 0:LeftKick2.IsDropped = 1:LeftKick.TimerEnabled = 1:PlaySound SoundFx("slingshot",DOFContactors),0, 0.8, -0.08, 0.05:End Sub 'PlaySound "name",loopcount,volume,pan,randompitch

Sub RightKick_Timer:RightKick.TimerEnabled = 0:RightKick2.IsDropped = 0:RightKick.IsDropped = 1:End Sub

Sub RightKick2_Slingshot():vpmTimer.PulseSw(65):RightKick.IsDropped = 0:RightKick2.IsDropped = 1:RightKick.TimerEnabled = 1:PlaySound SoundFx("slingshot",DOFContactors),0, 0.8, 0.08, 0.05:End Sub 'PlaySound "name",loopcount,volume,pan,randompitch

Sub LeftLane_Hit():Playsound "sensor":Controller.switch(53) = 1:End Sub
Sub LeftLane_UnHit():Controller.switch(53) = 0:End Sub


Sub Spinner_Spin():Playsound "spinner":End Sub

Sub Toplane1_Hit():Playsound "sensor":Controller.switch(42) = 1:End Sub
Sub Toplane1_UnHit():Controller.switch(42) = 0:End Sub
Sub Toplane2_Hit():Playsound "sensor":Controller.switch(52) = 1:End Sub
Sub Toplane2_UnHit():Controller.switch(52) = 0:End Sub
Sub Toplane3_Hit():Playsound "sensor":Controller.switch(62) = 1:End Sub
Sub Toplane3_UnHit():Controller.switch(62) = 0:End Sub

Sub LeftInLane_Hit():Playsound "sensor":Controller.switch(44) = 1:DOF 101, DOFOn:End Sub
Sub LeftInLane_UnHit():Controller.switch(44) = 0:DOF 101, DOFOff:End Sub
Sub LeftOutlane_Hit():Playsound "sensor":Controller.switch(54) = 1:End Sub
Sub LeftOutlane_UnHit():Controller.switch(54) = 0:End Sub
Sub RightInlane_Hit():Playsound "sensor":Controller.switch(44) = 1:DOF 102, DOFOn:End Sub
Sub RightInlane_UnHit():Controller.switch(44) = 0:DOF 102, DOFOff:End Sub
Sub RightOutlane_Hit():Playsound "sensor":Controller.switch(64) = 1:End Sub
Sub RightOutlane_UnHit():Controller.switch(64) = 0:End Sub

'One Way Switch
Dim TopDown
TopDown=False

Sub OneWaySwitch_Hit()
  OneWayTimer.Enabled=1
  TopDown=True
End Sub

Sub OneWayTimer_Timer()
  TopDown=False
  OneWayTimer.Enabled=0
End Sub

Sub sw63_Hit()
  If TopDown=False then Controller.switch(63) = 1':playsound "Diverter"
End Sub

Sub sw63_UnHit() 'extra switch juice
  me.timerenabled=1
End Sub

Sub sw63_Timer() 'extra switch juice
  Controller.switch(63) = 0
  me.Timerenabled=0
End Sub

'Drop-Targets
Sub dt1_hit():dtM.Hit 1:End Sub
Sub dt1_dropped():drop1 = 1:updategi:End Sub

Sub dt2_hit():dtM.Hit 2:End Sub
Sub dt2_dropped():drop2 = 1:updategi:End Sub

Sub dt3_hit():dtM.Hit 3:End Sub
Sub dt3_dropped():drop3 = 1:updategi:End Sub


'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub LeftSlingShot_Slingshot
  vpmTimer.PulseSw(72)
    PlaySound SoundFXDOF("slingshot",112,DOFPulse,DOFContactors),0,1,-0.05,0.05
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -42
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -25
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0':gi3.State = 1:Gi4.State = 1
    End Select
    LStep = LStep + 1
End Sub

'Right Slingshot

Sub RightSlingShot_Slingshot
  vpmTimer.PulseSw(72)
    PlaySound SoundFXDOF("slingshot",113,DOFPulse,DOFContactors), 0, 1, 0.05, 0.05
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -42
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -25
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0':gi1.State = 1:Gi2.State = 1
    End Select
    RStep = RStep + 1
End Sub

'-==============================

' Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 71:PlaySoundAtVol SoundFX("TOM_Bumpers_Reworked_v2_Middle_1",DOFContactors), Bumper1, VolBump:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 71:PlaySoundAtVol SoundFX("TOM_Bumpers_Reworked_v2_Middle_2",DOFContactors), Bumper2, VolBump:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 71:PlaySoundAtVol SoundFX("TOM_Bumpers_Reworked_v2_Middle_3",DOFContactors), Bumper3, VolBump:End Sub
Sub Bumper4_Hit:vpmTimer.PulseSw 71:PlaySoundAtVol SoundFX("TOM_Bumpers_Reworked_v2_Middle_1",DOFContactors), Bumper3, VolBump:End Sub

'Bumpers
'Sub Bumper1_Hit():vpmTimer.PulseSw(71):PlaySound "bumper1":End Sub
'
'Sub Bumper2_Hit():vpmTimer.PulseSw(71):PlaySound "bumper2":End Sub
'
'Sub Bumper3_Hit():vpmTimer.PulseSw(71):PlaySound "bumper3":End Sub
'
'Sub Bumper4_Hit():vpmTimer.PulseSw(71):PlaySound "bumper2":End Sub
'




'Outhole

Sub Drain_Hit():Playsound "drain":bsTrough.AddBall Me:End Sub
'Sub Drain_Hit():Playsound "drain":me.destroyball:End Sub 'Debug

'Ramps Top

Sub Ramp1_Hit():PlaySound "Ramp_Hit1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0:End Sub 'PlaySound "name",loopcount,volume,pan,randompitch
Sub Ramp2_Hit():PlaySound "Ramp_Hit1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0:End Sub

Sub Ramp3_Hit():PlaySound "Ramp_Hit1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0:End Sub 'PlaySound "name",loopcount,volume,pan,randompitch
Sub Ramp4_Hit():PlaySound "Ramp_Hit1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0:End Sub

Sub RHelp1_Hit():Playsound "ramp_hit3",0,1,-0.06:DOF 107, DOFPulse:End Sub
Sub RHelp2_Hit():Playsound "ramp_hit3",0,1,0.06:DOF 108, DOFPulse:End Sub

' Holes

Sub ArmsLock_Hit:PlaySound "kicker_enter":vpmTimer.PulseSw 43:End Sub
Sub LegsLock_Hit:PlaySound "kicker_enter":vpmTimer.PulseSw 73:End Sub

'***************
' Special lights
'***************

Set LampCallback = GetRef("UpdateMultipleLamps")

Sub UpdateMultipleLamps

  ' Robot animation

  if bDoAnimation = 1 Then
  bDoAnimation = 2
      'MsgBox "Do Animation"
      StartRobotLights
      StartRobotAnimation
  bDoAnimation = 0
  end if
  'Last12 = Current12

  ' Robot lights

' Current13 = l13.State
' if Current13 <> Last13 Then
'   if Current13 = 1 then
'     StartRobotLights
'     LightSeqTimer.Enabled=0
'     LightSeqTimer.Interval=1000
'   else                'StopRobotLights
'   maybestoprobotlights
'   'If CurrentRot<0 then StopRobotLights else LightSeqTimer.enabled=1
'   end if
' end if
' Last13 = Current13


  'Check BallTrough
  Current14 = l14.State
  if Current14 <> Last14 Then
    if Current14 = 1 then
      if bsTrough.Balls then bsTrough.ExitSol_On
    end if
  end if
  Last14 = Current14
End Sub

  'Robot Light Sequence Protector
'Sub  LightSeqTimer_Timer()
' StopRobotLights
' Me.Enabled=0
'End Sub


'================VP10 Fading Lamps Script

Dim LampState(200), FadingLevel(200)
Dim FlashSpeedUp(200), FlashSpeedDown(200), FlashMin(200), FlashMax(200), FlashLevel(200)

InitLamps()
LampTimer.Interval = 10
LampTimer.Enabled = 1


Sub LampTimer_Timer()
    Dim chgLamp, num, chg, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp) Then
        For ii = 0 To UBound(chgLamp)
            LampState(chgLamp(ii, 0) ) = chgLamp(ii, 1)       'keep the real state in an array
            FadingLevel(chgLamp(ii, 0) ) = chgLamp(ii, 1) + 4 'actual fading step
        Next

    End If

    UpdateLamps
End Sub

Dim bl24StateOff: bl24StateOff =0
Dim bl24StateOn: bl24StateOn =0
Dim bl24State: bl24State =0

Sub UpdateLamps
  NFadeL 2, l2 'FadeL
  NFadeL 3, l3 'FadeL
  NFadeL 5, l5 'FadeL
  NFadeL 6, l6 'FadeL
  NFadeL 7, l7 'FadeL
  NFadeL 8, l8 'FadeL
  NFadeL 9, l9 'FadeL
  NFadeL 10, l10 'FadeL
  NFadeL 11, l11 'FadeL
  NFadeL 12, l12

  NFadeL 13, l13
  'NFadeLS 13, l13 'start robot flash lights

  NFadeL 14, l14 'check balltrough
  NFadeL 15, l15 'FadeL
  NFadeL 16, l16 'FadeL
  NFadeL 17, l17 'FadeL
  NFadeL 18, l18 'FadeL
  NFadeL 19, l19 'FadeL

' NFadeL 20, l20 'FadeL
' NFadeL 21, l21 'FadeL
' NFadeL 22, l22 'FadeL
' NFadeL 23, l23 'FadeL
  NFadeL 24, l24 'FadeL

  If l24.state <> 0 then
  bl24StateOn = bl24StateOn + 2
  else
  bl24StateOff = bl24StateOff + 1
  end If


    if (bl24StateOff - bl24StateOn) < 20 then
    bl24State = 1
    'TextBox001.text = "State Test ON:" & bl24StateOn &" OFF:" & bl24StateOff & " on-off:" & bl24StateOff - bl24StateOn
    'TextBox001.text = "State Test ON:" & bl24StateOn &" OFF:" & bl24StateOff & " on-off:" & bl24StateOff - bl24StateOn
    end if

    'if TextBox001.visible = 1 then
    TextBox001.text = "State Test: " & bl24State
    'end If

    If bl24StateOff > 500 then
    bl24State = 0
    bl24StateOff = 21
    bl24StateOn =0
    end if


  NFadeLm 20, l20 'FadeL
  NFadeLm 21, l21 'FadeL
  NFadeLm 22, l22 'FadeL
  NFadeLm 23, l23 'FadeL
  NFadeLm 20, l20a 'FadeL
  NFadeLm 21, l21a 'FadeL
  NFadeLm 22, l22a 'FadeL
  NFadeLm 23, l23a 'FadeL

' NFadeL 25, l25 'FadeL
' NFadeL 26, l26 'FadeL
' NFadeL 27, l27 'FadeL
  NFadeLwf2 25, l25, l25F, l25F2 'FadeL
  NFadeLwf2 26, l26, l26F, l26F2 'FadeL
  NFadeLwf2 27, l27, l27F, l27F2 'FadeL

  Flash 28, A_RMS
  Flash 29, AR_MS
  Flash 30, ARM_S
  Flash 31, ARMS_
  Flash 32, B_RAIN
  Flash 33, BR_AIN
  Flash 34, BRA_IN
  Flash 35, BRAI_N
  Flash 36, BRAIN_
  Flash 37, B_ODY
  Flash 38, BO_DY
  Flash 39, BOD_Y
  Flash 40, BODY_
  Flash 41, L_EGS
  Flash 42, LE_GS
  Flash 43, LEG_S
  Flash 44, LEGS_

  NFadeL 45, l45 'FadeL
  NFadeL 46, l46 'FadeL
  NFadeLm 47, l47b 'FadeLm
  NFadeL 47, l47 'FadeL
  NFadeLm 51, l51b'FadeLm
  NFadeL 51, l51 'FadeL
End Sub

Sub InitLamps()
    Dim x
    For x = 0 to 200
        LampState(x) = 0        ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4      ' used to track the fading state
        FlashSpeedUp(x) = 0.2   ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.1 ' slower speed when turning off the flasher
        FlashMax(x) = 1         ' the maximum value when on, usually 1
        FlashMin(x) = 0.05         ' the minimum value when off, usually 0
        FlashLevel(x) = 0       ' the intensity of the flashers, usually from 0 to 1
    Next
End Sub

Sub AllLampsOff
    Dim x
    For x = 0 to 200
        SetLamp x, 0
    Next
End Sub

Sub SetLamp(nr, value)
    If value <> LampState(nr) Then
        LampState(nr) = abs(value)
        FadingLevel(nr) = abs(value) + 4
    End If
End Sub

'Walls

' Lights: used for VP10 standard lights, the fading is handled by VP itself

Sub NFadeL(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.state = 0:FadingLevel(nr) = 0
        Case 5:object.state = 1:FadingLevel(nr) = 1
    End Select
End Sub

'Special - Lights Robot Lights
Sub NFadeLS(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.state = 0:FadingLevel(nr) = 0:LightSeqTimer.enabled = 1
        Case 5:object.state = 1:FadingLevel(nr) = 1:StartRobotLights:LightSeqTimer.interval = 300
    End Select
End Sub

Sub LightSeqTimer_Timer()
  StopRobotLights
  me.enabled = 0
end sub

'LightSeqTimer
'StartRobotLights
'StopRobotLights


Sub NFadeLm(nr, object) ' used for multiple lights
    Select Case FadingLevel(nr)
        Case 4:object.state = 0
        Case 5:object.state = 1
    End Select
End Sub

Sub FadeLm(nr, a, b) 'Old
  Select Case LampState(nr)
    Case 2:b.state = 0
    Case 3:b.state = 1
    Case 4:a.state = 0
    Case 5:b.state = 1
    Case 6:a.state = 1
  End Select
End Sub

'Lights, Ramps & Primitives used as 4 step fading lights
'a,b,c,d are the images used from on to off

Sub FadeObj(nr, object, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 4:object.image = b:FadingLevel(nr) = 6                   'fading to off...
        Case 5:object.image = a:FadingLevel(nr) = 1                   'ON
        Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1             'wait
        Case 9:object.image = c:FadingLevel(nr) = FadingLevel(nr) + 1 'fading...
        Case 10, 11, 12:FadingLevel(nr) = FadingLevel(nr) + 1         'wait
        Case 13:object.image = d:FadingLevel(nr) = 0                  'Off
    End Select
End Sub

Sub FadeObjm(nr, object, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 4:object.image = b
        Case 5:object.image = a
        Case 9:object.image = c
        Case 13:object.image = d
    End Select
End Sub

Sub NFadeObj(nr, object, a, b)
    Select Case FadingLevel(nr)
        Case 4:object.image = b:FadingLevel(nr) = 0 'off
        Case 5:object.image = a:FadingLevel(nr) = 1 'on
    End Select
End Sub

Sub NFadeObjm(nr, object, a, b)
    Select Case FadingLevel(nr)
        Case 4:object.image = b
        Case 5:object.image = a
    End Select
End Sub

' Flasher objects

Sub Flash(nr, object)
    Select Case FadingLevel(nr)
        Case 4 'off
            FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown(nr)
            If FlashLevel(nr) < FlashMin(nr) Then
                FlashLevel(nr) = FlashMin(nr)
                FadingLevel(nr) = 0 'completely off
            End if
            Object.IntensityScale = FlashLevel(nr)
        Case 5 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
            If FlashLevel(nr) > FlashMax(nr) Then
                FlashLevel(nr) = FlashMax(nr)
                FadingLevel(nr) = 1 'completely on
            End if
            Object.IntensityScale = FlashLevel(nr)
    End Select
        '   Object.IntensityScale = 1
End Sub

Sub Flashm(nr, object) 'multiple flashers, it just sets the flashlevel
    Object.IntensityScale = FlashLevel(nr)
         ' Object.IntensityScale = 1
End Sub

Sub NFadeLwF(nr, object1, object2)
    Select Case FadingLevel(nr)
'   Case 0:If object2.intensityscale > 0 then Object2.intensityscale = object2.intensityscale + (object1.fadespeeddown * -1) *2 end if
'   Case 1:If object2.intensityscale < 1 then Object2.intensityscale = object2.intensityscale + (object1.fadespeedup) *2 end if

    Case 0:object2.intensityscale = 0
    Case 1:object2.intensityscale = 1
        Case 4:object1.state = 0:FadingLevel(nr) = 16
        Case 5:object1.state = 1:FadingLevel(nr) = 6':TextBox4.text = object1.fadespeedup 'to 6
'0.1 up, 0.1 down
    Case 6, 7, 8, 9, 10, 11, 12, 13, 14:If object2.intensityscale < 1 then Object2.intensityscale = object2.intensityscale + 0.1 end if:FadingLevel(nr) = FadingLevel(nr) + 1
    Case 15:If object2.intensityscale < 1 then Object2.intensityscale = object2.intensityscale + 0.1 end if:FadingLevel(nr) = 1':TextBox4.text = "Case 11"
    Case 16, 17, 18, 19, 20, 21, 22, 23, 24:If object2.intensityscale > 0 then Object2.intensityscale = object2.intensityscale + -0.1 end if:FadingLevel(nr) = FadingLevel(nr) + 1
    Case 25:If object2.intensityscale > 0 then Object2.intensityscale = object2.intensityscale + -0.1 end if:FadingLevel(nr) = 0':TextBox4.text = "Case 26"
    End Select
End Sub

Sub NFadeLwF2(nr, object1, object2, object3)  'one light two flashers
    Select Case FadingLevel(nr)

    Case 0:object2.intensityscale = 0:object3.intensityscale = object2.intensityscale
    Case 1:object2.intensityscale = 1:object3.intensityscale = object2.intensityscale
        Case 4:object1.state = 0:FadingLevel(nr) = 16
        Case 5:object1.state = 1:FadingLevel(nr) = 6':TextBox4.text = object1.fadespeedup 'to 6
'0.1 up, 0.1 down
    Case 6, 7, 8, 9, 10, 11, 12, 13, 14:If object2.intensityscale < 1 then Object2.intensityscale = object2.intensityscale + 0.1 end if:object3.intensityscale = object2.intensityscale: FadingLevel(nr) = FadingLevel(nr) + 1
    Case 15:If object2.intensityscale < 1 then Object2.intensityscale = object2.intensityscale + 0.1 end if:FadingLevel(nr) = 1:object3.intensityscale = object2.intensityscale':TextBox4.text = "Case 11"
    Case 16, 17, 18, 19, 20, 21, 22, 23, 24:If object2.intensityscale > 0 then Object2.intensityscale = object2.intensityscale + -0.1 end if:object3.intensityscale = object2.intensityscale :FadingLevel(nr) = FadingLevel(nr) + 1
    Case 25:If object2.intensityscale > 0 then Object2.intensityscale = object2.intensityscale + -0.1 end if:FadingLevel(nr) = 0:object3.intensityscale = object2.intensityscale':TextBox4.text = "Case 26"

  end select
End Sub

' Desktop Objects: Reels & texts (you may also use lights on the desktop)

' Reels

Sub FadeR(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.SetValue 1:FadingLevel(nr) = 6                   'fading to off...
        Case 5:object.SetValue 0:FadingLevel(nr) = 1                   'ON
        Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1              'wait
        Case 9:object.SetValue 2:FadingLevel(nr) = FadingLevel(nr) + 1 'fading...
        Case 10, 11, 12:FadingLevel(nr) = FadingLevel(nr) + 1          'wait
        Case 13:object.SetValue 3:FadingLevel(nr) = 0                  'Off
    End Select
End Sub

Sub FadeRm(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.SetValue 1
        Case 5:object.SetValue 0
        Case 9:object.SetValue 2
        Case 3:object.SetValue 3
    End Select
End Sub

'Texts

Sub NFadeT(nr, object, message)
    Select Case FadingLevel(nr)
        Case 4:object.Text = "":FadingLevel(nr) = 0
        Case 5:object.Text = message:FadingLevel(nr) = 1
    End Select
End Sub

Sub NFadeTm(nr, object, b)
    Select Case FadingLevel(nr)
        Case 4:object.Text = ""
        Case 5:object.Text = message
    End Select
End Sub



' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************


'===============================================================

' Extra Sounds


Sub PlasticRamps_Hit (idx)
  PlaySound "flip_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
End Sub

Sub Targets_Hit (idx)
  PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
End Sub

Sub Metals_Thin_Hit (idx)
  PlaySound "metalhit_thin", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub Metals_Medium_Hit (idx)
  PlaySound "metalhit_medium", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub Metals_Hit (idx) 'Inlanes & shooter lane
  PlaySound "metalhit2", 0, Vol(ActiveBall)*2, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub Gates_Hit (idx)
  PlaySound "gate4", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub RubberBands_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 20 then
        PlaySound "TOM_Rubber_1_Hard", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End Sub

Sub RubberSlings_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 10 then
        PlaySound "TOM_Loud_Rubber_" & Int(Rnd*7)+1, 0, 1.3*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End Sub

Sub RubberPosts_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
         PlaySound "TOM_Loud_Rubber_" & Int(Rnd*7)+1, 0, 1.3*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    RandomSoundRubber()
  End If
End Sub

Sub RandomSoundRubber()
    '                 dim finalspeed
    ' finalspeed=SQR((activeball.velx ^2) + (activeball.vely ^2))
    ' If finalspeed > 20 then
         PlaySound "TOM_Loud_Rubber_" & Int(Rnd*7)+1, 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    ' End if
    ' If finalspeed >= 6 AND finalspeed <= 20 then
    '     PlaySound "TOM_Loud_Rubber_" & Int(Rnd*7)+1, 0, .5*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    ' End If
End Sub

' Thalamus : This sub is used twice - this means ... this one IS NOT USED
' I moved the sound to the previous sub that seems to be the most advanced one
' Sub LeftFlipper_Collide(parm)
'   RandomSoundFlipper()
' End Sub
'
' Sub RightFlipper_Collide(parm)
'   RandomSoundFlipper()
' End Sub

Sub RandomSoundFlipper()
  PlaySound "TOM_Rubber_Flipper_Loud_" & Int(Rnd*7)+1, 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub


'Animated rubbers
'Sub Rubber_Straightb8_hit()
' Rubber_Straightb8.size_x = 90
' Rubberanim.enabled = 1
'End Sub

'Sub Rubber_Straightb14_hit()
' Rubber_Straightb14.size_x = 90
' Rubberanim.enabled = 1
'End Sub

'Sub Rubber_Straightb10_hit()
' Rubber_Straightb10.size_x = 90
' Rubberanim.enabled = 1
'End Sub
'Sub Rubber_Straightb5_hit()
' Rubber_Straightb5.size_x = 90
' Rubberanim.enabled = 1
'End Sub

'Sub Rubberanim_timer()
' Rubber_Straightb8.size_x = 100
' Rubber_Straightb14.size_x = 100
' Rubber_Straightb10.size_x = 100
' Rubber_Straightb5.size_x = 100
' me.enabled = 0
'End Sub

' Eala's rutine
' Dim Digits(40)
' Digits(0)=Array(f01_1, f02_1, f03_1, f04_1, f05_1, f06_1, f07_1, f08_1, f09_1, f10_1, f11_1, f12_1, f13_1, f14_1, f15_1, f16_1)
' Digits(1)=Array(f01_2, f02_2, f03_2, f04_2, f05_2, f06_2, f07_2, f08_2, f09_2, f10_2, f11_2, f12_2, f13_2, f14_2, f15_2, f16_2)
' Digits(2)=Array(f01_3, f02_3, f03_3, f04_3, f05_3, f06_3, f07_3, f08_3, f09_3, f10_3, f11_3, f12_3, f13_3, f14_3, f15_3, f16_3)
' Digits(3)=Array(f01_4, f02_4, f03_4, f04_4, f05_4, f06_4, f07_4, f08_4, f09_4, f10_4, f11_4, f12_4, f13_4, f14_4, f15_4, f16_4)
' Digits(4)=Array(f01_5, f02_5, f03_5, f04_5, f05_5, f06_5, f07_5, f08_5, f09_5, f10_5, f11_5, f12_5, f13_5, f14_5, f15_5, f16_5)
' Digits(5)=Array(f01_6, f02_6, f03_6, f04_6, f05_6, f06_6, f07_6, f08_6, f09_6, f10_6, f11_6, f12_6, f13_6, f14_6, f15_6, f16_6)
' Digits(6)=Array(f01_7, f02_7, f03_7, f04_7, f05_7, f06_7, f07_7, f08_7, f09_7, f10_7, f11_7, f12_7, f13_7, f14_7, f15_7, f16_7)
' Digits(7)=Array(f01_8, f02_8, f03_8, f04_8, f05_8, f06_8, f07_8, f08_8, f09_8, f10_8, f11_8, f12_8, f13_8, f14_8, f15_8, f16_8)
' Digits(8)=Array(f01_9, f02_9, f03_9, f04_9, f05_9, f06_9, f07_9, f08_9, f09_9, f10_9, f11_9, f12_9, f13_9, f14_9, f15_9, f16_9)
' Digits(9)=Array(f01_10, f02_10, f03_10, f04_10, f05_10, f06_10, f07_10, f08_10, f09_10, f10_10, f11_10, f12_10, f13_10, f14_10, f15_10, f16_10)
' Digits(10)=Array(f01_11, f02_11, f03_11, f04_11, f05_11, f06_11, f07_11, f08_11, f09_11, f10_11, f11_11, f12_11, f13_11, f14_11, f15_11, f16_11)
' Digits(11)=Array(f01_12, f02_12, f03_12, f04_12, f05_12, f06_12, f07_12, f08_12, f09_12, f10_12, f11_12, f12_12, f13_12, f14_12, f15_12, f16_12)
' Digits(12)=Array(f01_13, f02_13, f03_13, f04_13, f05_13, f06_13, f07_13, f08_13, f09_13, f10_13, f11_13, f12_13, f13_13, f14_13, f15_13, f16_13)
' Digits(13)=Array(f01_14, f02_14, f03_14, f04_14, f05_14, f06_14, f07_14, f08_14, f09_14, f10_14, f11_14, f12_14, f13_14, f14_14, f15_14, f16_14)
' Digits(14)=Array(f01_15, f02_15, f03_15, f04_15, f05_15, f06_15, f07_15, f08_15, f09_15, f10_15, f11_15, f12_15, f13_15, f14_15, f15_15, f16_15)
' Digits(15)=Array(f01_16, f02_16, f03_16, f04_16, f05_16, f06_16, f07_16, f08_16, f09_16, f10_16, f11_16, f12_16, f13_16, f14_16, f15_16, f16_16)
' Digits(16)=Array(f01_17, f02_17, f03_17, f04_17, f05_17, f06_17, f07_17, f08_17, f09_17, f10_17, f11_17, f12_17, f13_17, f14_17, f15_17, f16_17)
' Digits(17)=Array(f01_18, f02_18, f03_18, f04_18, f05_18, f06_18, f07_18, f08_18, f09_18, f10_18, f11_18, f12_18, f13_18, f14_18, f15_18, f16_18)
' Digits(18)=Array(f01_19, f02_19, f03_19, f04_19, f05_19, f06_19, f07_19, f08_19, f09_19, f10_19, f11_19, f12_19, f13_19, f14_19, f15_19, f16_19)
' Digits(19)=Array(f01_20, f02_20, f03_20, f04_20, f05_20, f06_20, f07_20, f08_20, f09_20, f10_20, f11_20, f12_20, f13_20, f14_20, f15_20, f16_20)
' Digits(20)=Array(f01_21, f02_21, f03_21, f04_21, f05_21, f06_21, f07_21, f08_21, f09_21, f10_21, f11_21, f12_21, f13_21, f14_21, f15_21, f16_21)
' Digits(21)=Array(f01_22, f02_22, f03_22, f04_22, f05_22, f06_22, f07_22, f08_22, f09_22, f10_22, f11_22, f12_22, f13_22, f14_22, f15_22, f16_22)
' Digits(22)=Array(f01_23, f02_23, f03_23, f04_23, f05_23, f06_23, f07_23, f08_23, f09_23, f10_23, f11_23, f12_23, f13_23, f14_23, f15_23, f16_23)
' Digits(23)=Array(f01_24, f02_24, f03_24, f04_24, f05_24, f06_24, f07_24, f08_24, f09_24, f10_24, f11_24, f12_24, f13_24, f14_24, f15_24, f16_24)
' Digits(24)=Array(f01_25, f02_25, f03_25, f04_25, f05_25, f06_25, f07_25, f08_25, f09_25, f10_25, f11_25, f12_25, f13_25, f14_25, f15_25, f16_25)
' Digits(25)=Array(f01_26, f02_26, f03_26, f04_26, f05_26, f06_26, f07_26, f08_26, f09_26, f10_26, f11_26, f12_26, f13_26, f14_26, f15_26, f16_26)
' Digits(26)=Array(f01_27, f02_27, f03_27, f04_27, f05_27, f06_27, f07_27, f08_27, f09_27, f10_27, f11_27, f12_27, f13_27, f14_27, f15_27, f16_27)
' Digits(27)=Array(f01_28, f02_28, f03_28, f04_28, f05_28, f06_28, f07_28, f08_28, f09_28, f10_28, f11_28, f12_28, f13_28, f14_28, f15_28, f16_28)
' Digits(28)=Array(f01_29, f02_29, f03_29, f04_29, f05_29, f06_29, f07_29, f08_29, f09_29, f10_29, f11_29, f12_29, f13_29, f14_29, f15_29, f16_29)
' Digits(29)=Array(f01_30, f02_30, f03_30, f04_30, f05_30, f06_30, f07_30, f08_30, f09_30, f10_30, f11_30, f12_30, f13_30, f14_30, f15_30, f16_30)
' Digits(30)=Array(f01_31, f02_31, f03_31, f04_31, f05_31, f06_31, f07_31, f08_31, f09_31, f10_31, f11_31, f12_31, f13_31, f14_31, f15_31, f16_31)
' Digits(31)=Array(f01_32, f02_32, f03_32, f04_32, f05_32, f06_32, f07_32, f08_32, f09_32, f10_32, f11_32, f12_32, f13_32, f14_32, f15_32, f16_32)
' Digits(32)=Array(f01_33, f02_33, f03_33, f04_33, f05_33, f06_33, f07_33, f08_33, f09_33, f10_33, f11_33, f12_33, f13_33, f14_33, f15_33, f16_33)
' Digits(33)=Array(f01_34, f02_34, f03_34, f04_34, f05_34, f06_34, f07_34, f08_34, f09_34, f10_34, f11_34, f12_34, f13_34, f14_34, f15_34, f16_34)
' Digits(34)=Array(f01_35, f02_35, f03_35, f04_35, f05_35, f06_35, f07_35, f08_35, f09_35, f10_35, f11_35, f12_35, f13_35, f14_35, f15_35, f16_35)
' Digits(35)=Array(f01_36, f02_36, f03_36, f04_36, f05_36, f06_36, f07_36, f08_36, f09_36, f10_36, f11_36, f12_36, f13_36, f14_36, f15_36, f16_36)
' Digits(36)=Array(f01_37, f02_37, f03_37, f04_37, f05_37, f06_37, f07_37, f08_37, f09_37, f10_37, f11_37, f12_37, f13_37, f14_37, f15_37, f16_37)
' Digits(37)=Array(f01_38, f02_38, f03_38, f04_38, f05_38, f06_38, f07_38, f08_38, f09_38, f10_38, f11_38, f12_38, f13_38, f14_38, f15_38, f16_38)
' Digits(38)=Array(f01_39, f02_39, f03_39, f04_39, f05_39, f06_39, f07_39, f08_39, f09_39, f10_39, f11_39, f12_39, f13_39, f14_39, f15_39, f16_39)
' Digits(39)=Array(f01_40, f02_40, f03_40, f04_40, f05_40, f06_40, f07_40, f08_40, f09_40, f10_40, f11_40, f12_40, f13_40, f14_40, f15_40, f16_40)



 Sub Displaytimer_Timer_
    Dim ChgLED, ii, jj, num, chg, stat, obj, b, x
    ChgLED=Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
       For ii=0 To UBound(chgLED)
          num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)
          For Each obj In Digits(num)
             If chg And 1 Then obj.visible=stat And 1
'             If chg And 1 Then obj.State=stat And 1
             chg=chg\2 : stat=stat\2
          Next
       Next
    End If
 End Sub



'Gottlieb Genesis
'added by Inkochnito
Sub editDips
  Dim vpmDips:Set vpmDips = New cvpmDips
  With vpmDips
    .AddForm 700, 400, "Genesis - DIP switches"
    .AddFrame 2, 4, 190, "Maximum credits", 49152, Array("8 credits", 0, "10 credits", 32768, "15 credits", &H00004000, "20 credits", 49152)                                                                                  'dip 15&16
    .AddFrame 2, 80, 190, "Coin chute 1 and 2 control", &H00002000, Array("seperate", 0, "same", &H00002000)                                                                                                                  'dip 14
    .AddFrame 2, 126, 190, "Playfield special", &H00200000, Array("replay", 0, "extra ball", &H00200000)                                                                                                                      'dip 22
    .AddFrame 2, 172, 190, "High games to date control", &H00000020, Array("no effect", 0, "reset high games 2-5 on power off", &H00000020)                                                                                   'dip 6
    .AddFrame 2, 218, 190, "Completing drop target sequence", &H00000080, Array("adds a letter to most complete part", 0, "spots a letter to each part", &H00000080)                                                          'dip 8
    .AddFrame 2, 264, 190, "Special lights after", &H40000000, Array("Hitting 'Lifeforce' when flashing", 0, "completing all body parts", &H40000000)                                                                         'dip 31
    .AddFrame 2, 310, 190, "Extra ball after completing", &H80000000, Array("4 body parts during the same ball", 0, "3 body parts during the same ball", &H80000000)                                                          'dip 32
    .AddFrame 205, 4, 190, "High game to date awards", &H00C00000, Array("not displayed and no award", 0, "displayed and no award", &H00800000, "displayed and 2 replays", &H00400000, "displayed and 3 replays", &H00C00000) 'dip 23&24
    .AddFrame 205, 80, 190, "Balls per game", &H01000000, Array("5 balls", 0, "3 balls", &H01000000)                                                                                                                          'dip 25
    .AddFrame 205, 126, 190, "Replay limit", &H04000000, Array("no limit", 0, "one per game", &H04000000)                                                                                                                     'dip 27
    .AddFrame 205, 172, 190, "Novelty", &H08000000, Array("normal", 0, "extra ball and replay scores points", &H08000000)                                                                                                     'dip 28
    .AddFrame 205, 218, 190, "Game mode", &H10000000, Array("replay", 0, "extra ball", &H10000000)                                                                                                                            'dip 29
    .AddFrame 205, 264, 190, "3rd coin chute credits control", &H20000000, Array("no effect", 0, "add 9", &H20000000)                                                                                                         'dip 30
    .AddChk 205, 316, 180, Array("Match feature", &H02000000)                                                                                                                                                                 'dip 26
    .AddChk 205, 331, 190, Array("Attract sound", &H00000040)                                                                                                                                                                 'dip 7
    .AddLabel 50, 360, 300, 20, "After hitting OK, press F3 to reset game with new settings."
    .ViewDips
  End With
End Sub

Set vpmShowDips = GetRef("editDips")

' Rules
Sub Rules()
  Dim Msg(32)
  Msg(0) = "Genesis - Gottlieb 1986" &Chr(10) &Chr(10)
  Msg(1) = ""
  Msg(2) = "SPECIAL: Completing all Body Parts lights LIFEFORCE"
  Msg(3) = "  Hitting the Vari-Target all the way back lights SPECIAL"
  Msg(4) = ""
  Msg(5) = "EXTRA BALL: Completing 3 Body Parts lights EXTRA BALL"
  Msg(6) = "  Completing next Body Part awards EXTRA BALL"
  Msg(7) = ""
  Msg(8) = "SCORING MULTIPLIER: Completing Body Parts when needed"
  Msg(9) = "  advances Scoreing Multiplier"
  Msg(10) = ""
  Msg(11) = "MULTI-MULTIPLIER: Scoring Multiplier is doubled during Multi-Ball play"
  Msg(12) = ""
  Msg(13) = "BODY PARTS LETTERS: Letters awarded bt hitting Vari-target"
  Msg(14) = "  all the way back or by scoring Flashing Targets or Sequences."
  Msg(15) = "  Return Rollovers flash Vari-Target for a period of time."
  Msg(16) = "  Hitting Vari-Target all the way back when flashing"
  Msg(17) = "  awards a letter in all Body Parts."
  Msg(18) = "  Completing the Drop Target Sequence (1-2-3) awards"
  Msg(19) = "  a letter in the most complete Body Part."
  Msg(20) = ""
  Msg(21) = "LIFEFORCE: Expose Robot by hitting Vari-target all the way back"
  Msg(22) = "  when LIFEFORCE is flashing."
  Msg(23) = ""
  Msg(24) = "MULTIBALL: Completing Body Part when needed enables ramp for capture"
  Msg(25) = ""
  Msg(26) = "ENTERING INITIALS: Enter letter by presing Flippers and Credit Button."
  Msg(27) = ""
  Msg(28) = ""

  For X = 1 To 28
    Msg(0) = Msg(0) + Msg(X) &Chr(13)
  Next

  MsgBox Msg(0), , "         Instructions and Rule Card"
End Sub

' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX
' PlaySoundAtVol sound, 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
' *******************************************************************************************************

' Play a sound, depending on the X,Y position of the table element (especially cool for surround speaker setups, otherwise stereo panning only)
' parameters (defaults): loopcount (1), volume (1), randompitch (0), pitch (0), useexisting (0), restart (1))
' Note that this will not work (currently) for walls/slingshots as these do not feature a simple, single X,Y position

Sub PlayXYSound(soundname, tableobj, loopcount, volume, randompitch, pitch, useexisting, restart)
  PlaySound soundname, loopcount, volume, AudioPan(tableobj), randompitch, pitch, useexisting, restart, AudioFade(tableobj)
End Sub

' Set position as table object (Use object or light but NOT wall) and Vol to 1

Sub PlaySoundAt(soundname, tableobj)
  PlaySound soundname, 1, 1, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'Set all as per ball position & speed.

Sub PlaySoundAtBall(soundname)
  PlaySoundAt soundname, ActiveBall
End Sub

'Set position as table object and Vol manually.

Sub PlaySoundAtVol(sound, tableobj, Volum)
  PlaySound sound, 1, Volum, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'Set all as per ball position & speed, but Vol Multiplier may be used eg; PlaySoundAtBallVol "sound",3

Sub PlaySoundAtBallVol(sound, VolMult)
  PlaySound sound, 0, Vol(ActiveBall) * VolMult, Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub

'Set position as bumperX and Vol manually.

Sub PlaySoundAtBumperVol(sound, tableobj, Vol)
  PlaySound sound, 1, Vol, Pan(tableobj), 0,0,1, 1, AudioFade(tableobj)
End Sub

'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / Genesis.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / Genesis.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / Genesis.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioFade(ball) ' Can this be together with the above function ?
  Dim tmp
  tmp = ball.y * 2 / Genesis.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
  Vol = Csng(BallVel(ball) ^2 / VolDiv)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
  Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
  BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

'*********** ROLLING SOUND *********************************
Const tnob = 5            ' total number of balls : 4 (trough) + 1 (fake ball used to animate Austin toy)
Const fakeballs = 0         ' number of balls created on table start (rolling sound will be skipped)
ReDim rolling(tnob)
InitRolling

Sub InitRolling:Dim i:For i=0 to (tnob-1):rolling(i) = False:Next:End Sub

Sub RollingUpdate()
    Dim BOT, b, ballpitch
    BOT = GetBalls
  ' stop the sound of deleted balls
  If UBound(BOT)<(tnob -1) Then
    For b = (UBound(BOT) + 1) to (tnob-1)
      rolling(b) = False
      StopSound("TOM_BallRoll_" & b)
    Next
  End If
  ' exit the Sub if no balls on the table
    If UBound(BOT) = fakeballs-1 Then Exit Sub
  ' play the rolling sound for each ball

    For b = 0 to UBound(BOT)
      If BallVel(BOT(b) ) > 1 Then
        rolling(b) = True
        if BOT(b).z < 30 Then ' Ball on playfield
          PlaySound("TOM_BallRoll_" & b), -1, Vol(BOT(b))*.1, Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        Else ' Ball on raised ramp
          PlaySound("TOM_BallRoll_" & b), -1, Vol(BOT(b) )*.1, Pan(BOT(b) ), 0, Pitch(BOT(b) )+50000, 1, 0, AudioFade(BOT(b) )
        End If
      Else
        If rolling(b) = True Then
          StopSound("TOM_BallRoll_" & b)
          rolling(b) = False
        End If
      End If
    Next
End Sub

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / (VolDiv/VolCol), Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

' Thalamus : Exit in a clean and proper way
Sub Genesis_exit()
  Controller.Pause = False
  Controller.Stop
End Sub

'******************************************************
'     RUBBER CORRECTION
'******************************************************

'****************************************************************************
'PHYSICS DAMPENERS

'These are data mined bounce curves,
'dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
'Requires tracking ballspeed to calculate COR

Sub RDampen_Timer()
Cor.Update
End Sub



Sub dPosts_Hit(idx)
  RandomSoundRubber()
  RubbersD.dampen Activeball
End Sub

Sub dSleeves_Hit(idx)
  SleevesD.Dampen Activeball
End Sub

dim RubbersD : Set RubbersD = new Dampener  'frubber
RubbersD.name = "Rubbers"
RubbersD.debugOn = False  'shows info in textbox "TBPout"
RubbersD.Print = False  'debug, reports in debugger (in vel, out cor)
'cor bounce curve (linear)
'for best results, try to match in-game velocity as closely as possible to the desired curve
'RubbersD.addpoint 0, 0, 0.935  'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 0, 0, 0.935 '0.96 'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 1, 3.77, 0.935 '0.96
RubbersD.addpoint 2, 5.76, 0.942 '0.967 'dont take this as gospel. if you can data mine rubber elasticitiy, please help!
RubbersD.addpoint 3, 15.84, 0.874
RubbersD.addpoint 4, 56, 0.64 'there's clamping so interpolate up to 56 at least

dim SleevesD : Set SleevesD = new Dampener  'this is just rubber but cut down to 85%...
SleevesD.name = "Sleeves"
SleevesD.debugOn = False  'shows info in textbox "TBPout"
SleevesD.Print = False  'debug, reports in debugger (in vel, out cor)
SleevesD.CopyCoef RubbersD, 0.85

Class Dampener
  Public Print, debugOn 'tbpOut.text
  public name, Threshold  'Minimum threshold. Useful for Flippers, which don't have a hit threshold.
  Public ModIn, ModOut
  Private Sub Class_Initialize : redim ModIn(0) : redim Modout(0): End Sub

  Public Sub AddPoint(aIdx, aX, aY)
    ShuffleArrays ModIn, ModOut, 1 : ModIn(aIDX) = aX : ModOut(aIDX) = aY : ShuffleArrays ModIn, ModOut, 0
    if gametime > 100 then Report
  End Sub

  public sub Dampen(aBall)
    if threshold then if BallSpeed(aBall) < threshold then exit sub end if end if
    dim RealCOR, DesiredCOR, str, coef
    DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
    RealCOR = BallSpeed(aBall) / cor.ballvel(aBall.id)
    coef = desiredcor / realcor
    if debugOn then str = name & " in vel:" & round(cor.ballvel(aBall.id),2 ) & vbnewline & "desired cor: " & round(desiredcor,4) & vbnewline & _
    "actual cor: " & round(realCOR,4) & vbnewline & "ballspeed coef: " & round(coef, 3) & vbnewline
    if Print then debug.print Round(cor.ballvel(aBall.id),2) & ", " & round(desiredcor,3)

' Thalamus - patched : ' Thalamus - patched :     aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
    aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
' Thalamus - patched :     aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
    aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
    'playsound "fx_knocker"
    if debugOn then TBPout.text = str
  End Sub

  Public Sub CopyCoef(aObj, aCoef) 'alternative addpoints, copy with coef
    dim x : for x = 0 to uBound(aObj.ModIn)
      addpoint x, aObj.ModIn(x), aObj.ModOut(x)*aCoef
    Next
  End Sub


  Public Sub Report()   'debug, reports all coords in tbPL.text
    if not debugOn then exit sub
    dim a1, a2 : a1 = ModIn : a2 = ModOut
    dim str, x : for x = 0 to uBound(a1) : str = str & x & ": " & round(a1(x),4) & ", " & round(a2(x),4) & vbnewline : next
    TBPout.text = str
  End Sub


End Class

'Tracks ball velocity for judging bounce calculations & angle
'apologies to JimmyFingers is this is what his script does. I know his tracks ball velocity too but idk how it works in particular
dim cor : set cor = New CoRTracker
cor.debugOn = False
'cor.update() - put this on a low interval timer
Class CoRTracker
  public DebugOn 'tbpIn.text
  public ballvel

  Private Sub Class_Initialize : redim ballvel(0) : End Sub
  'TODO this would be better if it didn't do the sorting every ms, but instead every time it's pulled for COR stuff
  Public Sub Update() 'tracks in-ball-velocity
    dim str, b, AllBalls, highestID : allBalls = getballs
    'if uBound(allballs) < 0 then if DebugOn then str = "no balls" : TBPin.text = str : exit Sub else exit sub end if: end if
    for each b in allballs
      if b.id >= HighestID then highestID = b.id
    Next

    if uBound(ballvel) < highestID then redim ballvel(highestID)  'set bounds

    for each b in allballs
      ballvel(b.id) = BallSpeed(b)
'     if DebugOn then
'       dim s, bs 'debug spacer, ballspeed
'       bs = round(BallSpeed(b),1)
'       if bs < 10 then s = " " else s = "" end if
'       str = str & b.id & ": " & s & bs & vbnewline
'       'str = str & b.id & ": " & s & bs & "z:" & b.z & vbnewline
'     end if
    Next
    'if DebugOn then str = "ubound ballvels: " & ubound(ballvel) & vbnewline & str : if TBPin.text <> str then TBPin.text = str : end if
  End Sub
End Class

Function LinearEnvelope(xInput, xKeyFrame, yLvl)
  dim y 'Y output
  dim L 'Line
  dim ii : for ii = 1 to uBound(xKeyFrame)  'find active line
    if xInput <= xKeyFrame(ii) then L = ii : exit for : end if
  Next
  if xInput > xKeyFrame(uBound(xKeyFrame) ) then L = uBound(xKeyFrame)  'catch line overrun
  Y = pSlope(xInput, xKeyFrame(L-1), yLvl(L-1), xKeyFrame(L), yLvl(L) )

  'Clamp if on the boundry lines
  'if L=1 and Y < yLvl(LBound(yLvl) ) then Y = yLvl(lBound(yLvl) )
  'if L=uBound(xKeyFrame) and Y > yLvl(uBound(yLvl) ) then Y = yLvl(uBound(yLvl) )
  'clamp 2.0
  if xInput <= xKeyFrame(lBound(xKeyFrame) ) then Y = yLvl(lBound(xKeyFrame) )  'Clamp lower
  if xInput >= xKeyFrame(uBound(xKeyFrame) ) then Y = yLvl(uBound(xKeyFrame) )  'Clamp upper

  LinearEnvelope = Y
End Function



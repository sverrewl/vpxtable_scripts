''************ Big Bamg Bar
''************ Capcom, 1996
''************ vp9 version by unclewilly and jimmyfingers, artwork by grizz
''************ original 3dmodels by rom extracted from his Future Pinball release
''************ vpx conversion by ninuzzu

Option Explicit
Randomize

' Thalamus 2019 February : Improved directional sounds
' !! NOTE : Table not verified yet !!

' Options
' Volume devided by - lower gets higher sound

Const VolDiv = 2000    ' Lower number, louder ballrolling/collition sound
Const VolCol = 10      ' Ball collition divider ( voldiv/volcol )

' The rest of the values are multipliers
'
'  .5 = lower volume
' 1.5 = higher volume

Const VolBump   = 2    ' Bumpers volume.
Const VolRol    = 1    ' Rollovers volume.
Const VolGates  = 1    ' Gates volume.
Const VolMetal  = 1    ' Metals volume.
Const VolRB     = 1    ' Rubber bands volume.
Const VolRH     = 1    ' Rubber hits volume.
Const VolPo     = 1    ' Rubber posts volume.
Const VolPi     = 1    ' Rubber pins volume.
Const VolPlast  = 1    ' Plastics volume.
Const VolTarg   = 1    ' Targets volume.
Const VolWood   = 1    ' Woods volume.
Const VolKick   = 1    ' Kicker volume.
Const VolSpin   = 1.5  ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.


'Const path ="HKEY_CURRENT_USER\SOFTWARE\Visual Pinball\Controller\"
'Dim sString : sString = "Const dir = """ &  path &  """" & vbCrLf & "Dim oShell: Set oShell = CreateObject(""WScript.Shell"")" & vbCrLf &"If Err.number <> 0 Then" & vbCrLf & "oShell.RegWrite dir & ""ForceDisableB2S"",1, ""REG_DWORD"""& vbCrLf & "end if" & vbCrLf& "Set oShell = nothing"& vbCrLf
'Dim objFSO : Set objFSO=CreateObject("Scripting.FileSystemObject")
'Dim objFile:  Set objFile = objFSO.CreateTextFile("NoB2S.vbs",True)
'objFile.Write sString & vbCrLf
'objFile.Close

'ExecuteGlobal GetTextFile("NoB2S.vbs")
'If Err Then MsgBox "Unable to run NoB2S"

 '******************* Options *********************
' DMD/Backglass Controller Setting
Dim cController:cController = 3   '0=Use value defined in cController.txt, 1=VPinMAME, 2=UVP server, 3=B2S server, 4=B2S with DOF (disable VP mech sounds)
'*************************************************


Const dir = "HKEY_CURRENT_USER\SOFTWARE\Visual Pinball\Controller\"
Dim oShell: Set oShell = CreateObject("WScript.Shell")

Const UseVPMColoredDMD = 1

if Table.ShowFSS = true or Table.ShowDT = true then
oShell.RegWrite dir & "ForceDisableB2S",1, "REG_DWORD"

cController = 0
  if Table.ShowFSS = true then
  ScoreText.text = " "
  ScoreText.visible = 0
  Else
  ScoreText.text = "DMD"
  ScoreText.visible = 1
  end if

else
oShell.RegWrite dir & "ForceDisableB2S",0, "REG_DWORD"
ScoreText.text = " "
ScoreText.visible = 0
cController = 3
end if
Set oShell = nothing


Dim DesktopMode:DesktopMode = Table.ShowDT
Dim UseVPMDMD

if Table.ShowFSS = true then
UseVPMDMD = 0
Else
UseVPMDMD = DesktopMode
end if

const SingleScreenFS = 0
Const IntegratedDMDType = 0

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

Const GILevel = 1
Const FLLevel = 1
Const BTLevel = 1


Const cUSESHADOW = 1
Const cUSECOLORGRADE = 1
Const cUSECCUPBOOST = 1
Const cUSELITEBOOST = 1
Const cUSEBACKGLASS = 1
Const cDYNPINBALL = 1
Const cDYNGAMEBLADES =1
Const cSHADOWBLADES =1
Const cUSEBGLASSHIGH =0
Const cUSETOPPER =1
Const cUSEBACKLIGHT =1
Const cUSEBGLASSCCX =0
Const cUSEBGLASSREFL =1
Const cUSEDUALRAMPS =0

Dim nxx, DNS
DNS = Table.NightDay

Dim DivLevel: DivLevel = 35
Dim DNSVal: DNSVal = Round(DNS/10)
Dim DNShift: DNShift = 2

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
Dim BValUP: BValUP=Array (30,60,90,120,150,180,210,240,255,255,255,255,255,255,255,255)
Dim RValDN: RValDN=Array (255,210,180,150,120,90,60,30,10,10,10)
Dim GValDN: GValDN=Array (255,210,180,150,120,90,60,30,10,10,10)
Dim BValDN: BValDN=Array (255,210,180,150,120,90,60,30,10,10,10)
Dim FValUP: FValUP=Array (35,40,45,50,55,60,65,70,75,80,85,90,95,100,105)
Dim FValDN: FValDN=Array (100,85,80,75,70,65,60,55,50,45,40,35,30)
Dim MVSAdd: MVSAdd=Array (0.9,0.9,0.8,0.8,0.7,0.7,0.6,0.6,0.5,0.5,0.4,0.3,0.2,0.1)
Dim ReflDN: ReflDN=Array (60,55,50,45,40,35,30,28,26,24,22,20,19,18,16,15,14,13,12,11,10)




' PLAYFIELD GENERAL OPERATIONAL and LOCALALIZED GI ILLUMINATION
Dim aAllFlashers: aAllFlashers=Array(TopHigh,TopHigh1,TopHigh2, BBBLeftgirl,BBBRightgirl,FlBbBTitle,Flasher4,Flasher5,Flasher6,Flasher7,Flasher8,Flasher9,Flasher10,Flasher11,Flasher12,_
Flasher1,Flasher2,Flasher13,Flasher14,Flasher15,F21a,F24a,F25a,l62,l62a,l62b)
Dim aGiLights: aGiLights=array(l9,l10,l11,l12,l14,l17,l18,l21,l22,l23,l23a,l24,l25,l26,l27,l28,l29,l38,l39,l40,l128, f26)
Dim BloomLights: BloomLights=array(plungerbulb,l104,l119,l120,l122a,l123a,l124a,_
f21,F22,F23,F24,F25)
Dim TargetDropGi: TargetDropGi = array()

' PLAYFIELD GLOBAL INTENSITY ILLUMINATION FLASHERS
Flasher16.opacity = OPSValues(DNSVal + DNShift) / DivValues(DNSVal)
Flasher16.intensityscale = DNSValues(DNSVal + DNShift) /DivValues(DNSVal)
Flasher17.opacity = OPSValues(DNSVal + DNShift) /DivValues(DNSVal)
Flasher17.intensityscale = DNSValues(DNSVal + DNShift) /DivValues(DNSVal)
Flasher19.opacity = OPSValues(DNSVal + DNShift) /DivValues(DNSVal)
Flasher19.intensityscale = DNSValues(DNSVal + DNShift) /DivValues(DNSVal)
'Flasher18.opacity = OPSValues(DNSVal + DNShift) /DivValues(DNSVal)
'Flasher18.intensityscale = DNSValues(DNSVal + DNShift) /DivValues(DNSVal)

if cUSELITEBOOST then
Dim LiteUP: LiteUP=Array (10,6,5,4,4,4,4,3,3,3,2,2,1,0.5,0.4,0.3,0.2,0.1)
Flasher20.Color = RGB(RValUP(DNSVal)/LiteUP(DNSVal),GValUP(DNSVal)/LiteUP(DNSVal),BValUP(DNSVal+4)/LiteUP(DNSVal))
end if

'BACKBOX & BACKGLASS ILLUMINATION
if cUSEBACKGLASS then
BBBbgDark.ModulateVsAdd = MVSAdd(DNSVal+2)
BBBbgDark.Color = RGB(RValUP(DNSVal),GValUP(DNSVal),BValUP(DNSVal))
BBBbgDark.Amount = FValUP(DNSVal) / DivValues(DNSVal)
BBBbgHigh.Color = RGB(RValDN(DNSVal),GValDN(DNSVal),BValDN(DNSVal))
BBBbgHigh.Amount = FValDN(DNSVal)  / DivValues(DNSVal)

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
Const nBACKBoost = 5
FloodLightLeft.Color = RGB(RValDN(DNSVal),GValDN(DNSVal),BValDN(DNSVal))
FloodLightLeft.Amount = FValDN(DNSVal)  * nBACKBoost
'FloodLightLeft.opacity = OPSValues(DNSVal+ DNShift) /DivValues(DNSVal)
FloodLightRight.Color = RGB(RValDN(DNSVal),GValDN(DNSVal),BValDN(DNSVal))
FloodLightRight.Amount = FValDN(DNSVal) * nBACKBoost
'FloodLightRight.opacity = OPSValues(DNSVal+ DNShift) /DivValues(DNSVal)
end if 'cUSEBACKLIGHT

'FlBgNeonLamp.Color = RGB(RValDN(DNSVal),GValDN(DNSVal),BValDN(DNSVal))
FlBgNeonLamp.Amount = (FValDN(DNSVal) * 10.0) / DivValues2(DNSVal)
'FlBgNeonLamp1.Color = RGB(RValDN(DNSVal),GValDN(DNSVal),BValDN(DNSVal))
FlBgNeonLamp1.Amount = (FValDN(DNSVal) * 10.0) / DivValues2(DNSVal)

BBBbgFrame.ModulateVsAdd = MVSAdd(DNSVal) * 0.2
BBBbgFrame.Color = RGB(RValUP(DNSVal),GValUP(DNSVal),BValUP(DNSVal))
BBBbgFrame.Amount = FValUP(DNSVal) * DivValues(DNSVal)

BBBbgHigh.intensityscale = 0.7
FlBgNeonLamp.intensityscale = 0.6
FlBgNeonLamp1.intensityscale = 0.6
BBBbgDark.intensityscale = 0.3
BBBbgFrame.intensityscale = 0.3
BBBbgHigh1.intensityscale = 1.0


if cUSEBGLASSHIGH then ' GBGHigh3 is _defhorz flasher, Additive, AMT:100, OP:2, DB:8000, MOD:1

BGHigh3.intensityscale = 0.2
Dim ReflUP: ReflUP=Array (0.3,0.4,0.5,0.55,0.6,0.65,0.7,0.75,0.8,0.85,0.9)
BGDark.intensityscale = ReflUP(DNSVal)

end if 'cUSEBGLASSHIGH


Dim tMaskFill: tMaskFill = Array(Empty,BGFrameMaskFill2,BGFrameMaskFill3,Empty,Empty,Empty)
' Fill with: BGFrameMaskFill1,BGFrameMaskFill2,BGFrameMaskFill3,BGFrameMaskFill4,BGFrameMaskFill5,PFRearBoxFill1
if cUSEBGLASSREFL then

Dim ReflFill: ReflFill=Array (10,10,20,20,30,30,30,30,40,40,40)
Dim ReflImg: ReflImg=Array (160,200,300,400,500,600,700,800,900,1000,1000)
Dim FillMag: FillMag=Array (0.8,0.8,0.7,0.7,0.6,0.6,0.55,0.55,0.54,0.54,0.53)

'playfield glass reflections
if NOT IsEmpty(tMaskFill(1)) then: tMaskFill(1).opacity= ReflFill(DNSVal)*2: end if' IMG:full table, def glass trans, FLTR:screen,AMT:900,OP:50, AB=off, RGB=0,0,32, DB=-10000,MOD=1
if NOT IsEmpty(tMaskFill(2)) then: tMaskFill(2).opacity= ReflFill(DNSVal)*3: end if' IMG:top half, def glass trans, FLTR:screen,AMT:900,OP:50, AB=off, RGB=0,0,32, DB=-10000,MOD=1
if NOT IsEmpty(tMaskFill(3)) then: tMaskFill(3).opacity= ReflImg(DNSVal): end if' IMG:backglass,   FLTR:screen,AMT:300,OP:1000,AB=on, RBG=16,16,32, DB=-1000,MOD=0.3

if NOT IsEmpty(tMaskFill(1)) then: tMaskFill(1).intensityscale= 0.5: end if
if NOT IsEmpty(tMaskFill(2)) then: tMaskFill(2).intensityscale= 0.5: end if
'backbox glass reflections
if NOT IsEmpty(tMaskFill(0)) then: tMaskFill(0).opacity= ReflFill(DNSVal) +40: end if ' IMG:black mask fill, FLTR:screen,AMT:500,OP:90, AB=off, RGB=0,0,32, DB=-10000,MOD=1
'background reflective image usually __default_screen_space_reflection.png
if NOT IsEmpty(tMaskFill(4)) then: tMaskFill(4).opacity= ReflImg(DNSVal) / 3 : end if' IMG:Def SS Refl, FLTR:screen,AMT:100,OP:1000,AB=on, RBG=10,10,10, DB=-10000,MOD=0.5

' playfield rear EM/SS vertical rececessed shadow
if NOT IsEmpty(tMaskFill(5)) then: tMaskFill(5).opacity= ReflFill(DNSVal)*4: end if' IMG:full table, def glass trans, FLTR:screen,AMT:900,OP:50, AB=off, RGB=0,0,32, DB=-10000,MOD=1
if NOT IsEmpty(tMaskFill(5)) then: tMaskFill(5).intensityscale= 0.5: end if

Else

if NOT IsEmpty(tMaskFill(0)) then: tMaskFill(0).visible =0: end if
if NOT IsEmpty(tMaskFill(1)) then: tMaskFill(1).visible =0: end if
if NOT IsEmpty(tMaskFill(2)) then: tMaskFill(2).visible =0: end if
if NOT IsEmpty(tMaskFill(3)) then: tMaskFill(3).visible =0: end if
if NOT IsEmpty(tMaskFill(4)) then: tMaskFill(4).visible =0: end if
end if 'cUSEBGLASSREFL


end If ' cUSEBACKGLASS


' shadow code.
'PF=flasher(Screen,AMT:1,OP:50,DB:0,MOD:1,RGB=255)
'PS=flasher(Screen,AMT:1700,OP:30,DB:-10000,MOD:1,RGB=7,7,7)
'DS =flasher (Screen ,AMT:1,OP:80,DB:0,MOD:1,RGB=255)
Dim tShadows: tShadows = Array(OCPFShadow,OCPFShadow1,OCPFShadow2,OCPFShadow3, OCPFDarkShadow)
'Dim tShadows: tShadows = Array(OCPFShadow,OCPFShadow1,OCPFShadow2,OCPFShadow3,OCPFDarkShadow )
if cUSESHADOW then
'LOWER PF Shadows
Const cPFShDiv = 1.2 ' higher = less opaque, smaller= more opaque
Dim PFShDiv: PFShDiv=      Array (1.39*cPFShDiv,1.36*cPFShDiv,1.33*cPFShDiv,1.29*cPFShDiv,1.26*cPFShDiv,1.23*cPFShDiv,1.19*cPFShDiv,1.16*cPFShDiv,1.13*cPFShDiv,1.1*cPFShDiv,1.0*cPFShDiv)
Dim PFShValues: PFShValues=Array (50,80,90,91,92,93,94,95,96,97,98,99,100,101,101,101,101,101)
Dim PFDrkShValues: PFDrkShValues=Array (100,99,90,80,70,60,50,40,30,20,10,5,0)

  if NOT IsEmpty(tShadows(0)) then
  tShadows(0).visible =1
  tShadows(0).opacity = PFShValues(DNSVal) / PFShDiv(DNSVal) ' set height to ~0.5 just above PF
  end if
  if NOT IsEmpty(tShadows(1)) then
  tShadows(1).visible =1
  tShadows(1).opacity = PFShValues(DNSVal) / 1.6 ' set height to ~60 just above plastics
  end if
' REVERSE Shadow
  if NOT IsEmpty(tShadows(4)) then
  tShadows(4).visible =1
  tShadows(4).opacity = Round(PFDrkShValues(DNSVal) /1.2)' set height to ~60 just above plastics
  end if

' UPPER PF Shadows
Const cPFShDiv2 = 1.4 ' higher = less opaque, smaller= more opaque
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
end if

' color grade transition
if cUSECOLORGRADE then
Const nCGDiv = 1.0 ' increase when gets to understaturated at higher light settings
Const rSHIFT = 0 ' increase too saturated at lower light settings
Dim ColValues: ColValues=Array ("_ColorGrade_9_Lite","_ColorGrade_8_Lite","_ColorGrade_7_Lite","_ColorGrade_6_Lite", _
"_ColorGrade_5_Lite","_ColorGrade_4_Lite","_ColorGrade_3_Lite","_ColorGrade_2_Lite","_ColorGrade_1_Lite","_ColorGrade_0_Lite","_ColorGrade_0_Lite","_ColorGrade_0_Lite",_
"_ColorGrade_0_Lite","_ColorGrade_0_Lite","_ColorGrade_0_Lite","_ColorGrade_0_Lite","_ColorGrade_0_Lite","_ColorGrade_0_Lite")

table.ColorGradeImage = ColValues((int)((DNSVal+rSHIFT)/nCGDiv))
end if 'cUSECOLORGRADE


Dim TubeVal: TubeVal=Array ("Ramps-PlasticO","Ramps-Plastic1","Ramps-Plastic2","Ramps-Plastic3", _
"Ramps-Plastic4","Ramps-Plastic5","Ramps-Plastic6","Ramps-Plastic7","Ramps-Plastic8","Ramps-Plastic9", _
"Ramps-Plastic9","Ramps-Plastic9")

Primitive2.material = TubeVal(DNSVal)


Table.PlayfieldReflectionStrength = ReflDN(DNSVal + 4)

'For each nxx in aGiLights:nxx.intensity = nxx.intensity * SysDNSVal(DNSVal) /DivValues3(DNSVal) :Next
'For each nxx in aAllFlashers:nxx.amount = nxx.amount / DivValues3(DNSVal):Next
'For each nxx in aAllFlashers:nxx.opacity = nxx.opacity * OPSValues(DNSVal) / DivLevel:Next
'For each nxx in BloomLights:nxx.intensity = nxx.intensity *SysDNSVal(DNSVal)/DivValues3(DNSVal):Next
'For each nxx in TargetDropGi:nxx.intensity = nxx.intensity *SysDNSVal(DNSVal)/DivValues3(DNSVal):Next

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


if cDYNPINBALL then
Dim DynPinVals: DynPinVals=Array ("_defballHDR0","_defballHDR0", _
"_defballHDR1","_defballHDR1", "_defballHDR1", _
"_defballHDR2","_defballHDR2", _
"_defballHDR3","_defballHDR3","_defballHDR3","_defballHDR3","_defballHDR3")

Table.ballImage = DynPinVals(DNSVal)
end if

if cDYNGAMEBLADES then
Dim DynBladeVals: DynBladeVals=Array ("_defgbladeswoodblack0","_defgbladeswoodblack0",_
"_defgbladeswoodblack1","_defgbladeswoodblack1", "_defgbladeswoodblack1",_
"_defgbladeswoodblack2","_defgbladeswoodblack2",_
"_defgbladeswoodblack3","_defgbladeswoodblack3","_defgbladeswoodblack3","_defgbladeswoodblack3","_defgbladeswoodblack3")

Primitive5.image = DynBladeVals(DNSVal)
Primitive6.image = DynBladeVals(DNSVal)
end if

'**************************************************************************************
'               [CC Color Correction] Textures must be CC corrected to work
'**************************************************************************************

Const UseMixGtoB = 1
Dim FlGlobals : FlGlobals=Array (1.0,1.5,2.0,2.5,3.0,3.5,4.0,4.5,5.0,5.0,5.0)

Dim DivGlobal: DivGlobal = FlGlobals(DNSVal) * 1.0 ' multiply between 1.0 to 0.75, 1.0 = default
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

Dim FlData : FlData=Array (Flasher17,Flasher16,Flasher19,Flasher18,Flasher20,Empty,Empty,Empty)

'half blue channel
FlData(0).amount = AMTValues(DNSVal ) * DivValues3(DNSVal)
FlData(0).DepthBias = BiasValues(DNSVal) * 100.0
FlData(0).intensityscale = FlData(0).intensityscale * MixBlueChan(DNSVal) '0.0008'
FlData(0).opacity = CCBlueValues(DNSVal) '* CCFSValues(DNSVal)'

'Full blue channel
if UseMixGtoB then
FlData(0).Color = RGB(0, MixGtoB(DNSVal), 10)
end if

'full red channel
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

'Full Spectrum
FlData(4).amount = AMTValues(DNSVal ) * (DivValues3(DNSVal) * 0.1)
FlData(4).DepthBias = BiasValues(DNSVal) * 100.0
'FlData(4).intensityscale = FlData(4).intensityscale * MixFullSpectrum(DNSVal) '0.0008'
'FlData(4).opacity = FlData(4).opacity * CCFSValues(DNSVal)

Const USEBGBLUE = 0
Const USEBGREDGREEN = 0
Const nDAMPNBLUE = 0
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
Dim DivCCUP: DivCCUP = 0.8

'original values depricated
'Dim CCUPValues : CCUPValues=Array (1.0*DivCCUP,1.0*DivCCUP,1.0*DivCCUP,1.5*DivCCUP,2.5*DivCCUP,5.0*DivCCUP,10.0*DivCCUP,20.0*DivCCUP,30.0*DivCCUP,40.0*DivCCUP,40.0*DivCCUP)
'new values
Dim CCUPValues : CCUPValues=Array (1.0*DivCCUP,1.0*DivCCUP,1.0*DivCCUP,1.5*DivCCUP,2.5*DivCCUP,5.0*DivCCUP,10.0*DivCCUP,12.0*DivCCUP,20.0*DivCCUP,40.0*DivCCUP,40.0*DivCCUP)

FlData(3).intensityscale = FlData(3).intensityscale * MixFullSpectrum(DNSVal) * CCUPValues(DNSVal)'0.0008'
FlData(3).opacity = FlData(3).opacity * CCFSValues(DNSVal) * CCUPValues(DNSVal)
FlData(4).intensityscale = FlData(4).intensityscale * MixFullSpectrum(DNSVal) * CCUPValues(DNSVal)'0.0008'
FlData(4).opacity = FlData(4).opacity * CCFSValues(DNSVal) * CCUPValues(DNSVal)
'FlData(4).intensityscale = FlData(4).intensityscale * MixFullSpectrum(DNSVal) '0.0008'
'FlData(4).opacity = FlData(4).opacity * CCFSValues(DNSVal)
Else
FlData(3).intensityscale = FlData(3).intensityscale * MixFullSpectrum(DNSVal) '0.0008'
FlData(3).opacity = FlData(3).opacity * CCFSValues(DNSVal)
FlData(4).intensityscale = FlData(4).intensityscale * MixFullSpectrum(DNSVal) '0.0008'
FlData(4).opacity = FlData(4).opacity * CCFSValues(DNSVal)
end if


FlValues(0) = FlData(0).intensityscale
FlValues(1) = FlData(1).intensityscale
FlValues(2) = FlData(2).intensityscale
FlValues(3) = FlData(3).intensityscale


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

' ***************************************************************************
'
' ****************************************************************************




   LoadVPM "01560000", "Capcom.VBS", 3.10


Dim cNewController
Sub LoadVPM(VPMver, VBSfile, VBSver)
  Dim FileObj, ControllerFile, TextStr


  On Error Resume Next
  If ScriptEngineMajorVersion < 5 Then MsgBox "VB Script Engine 5.0 or higher required"
  ExecuteGlobal GetTextFile(VBSfile)
  If Err Then MsgBox "Unable to open " & VBSfile & ". Ensure that it is in the same folder as this table. " & vbNewLine & Err.Description
  cNewController = 1
  If cController = 0 then
    Set FileObj=CreateObject("Scripting.FileSystemObject")
    If Not FileObj.FolderExists(UserDirectory) then
      Msgbox "Visual Pinball\User directory does not exist. Defaulting to vPinMame"
    ElseIf Not FileObj.FileExists(UserDirectory & "cController.txt") then
      Set ControllerFile=FileObj.CreateTextFile(UserDirectory & "cController.txt",True)
      ControllerFile.WriteLine 1: ControllerFile.Close
    Else
      Set ControllerFile=FileObj.GetFile(UserDirectory & "cController.txt")
      Set TextStr=ControllerFile.OpenAsTextStream(1,0)
      If (TextStr.AtEndOfStream=True) then
        Set ControllerFile=FileObj.CreateTextFile(UserDirectory & "cController.txt",True)
        ControllerFile.WriteLine 1: ControllerFile.Close
      Else
        cNewController=Textstr.ReadLine: TextStr.Close
      End If
    End If
  Else
    cNewController = cController
  End If

  Select Case cNewController
    Case 1
      Set Controller = CreateObject("VPinMAME.Controller")
      If Err Then MsgBox "Can't Load VPinMAME." & vbNewLine & Err.Description
      If VPMver>"" Then If Controller.Version < VPMver Or Err Then MsgBox "VPinMAME ver " & VPMver & " required."
      If VPinMAMEDriverVer < VBSver Or Err Then MsgBox VBSFile & " ver " & VBSver & " or higher required."
    Case 2
      Set Controller = CreateObject("UltraVP.BackglassServ")
    Case 3,4
      Set Controller = CreateObject("B2S.Server")
  End Select
  On Error Goto 0


End Sub

'*************************************************************
'Toggle DOF sounds on/off based on cController value
'*************************************************************
Dim ToggleMechSounds
Function SoundFX (sound)
    If cNewController= 4 and ToggleMechSounds = 0 Then
        SoundFX = ""
    Else
        SoundFX = sound
    End If
End Function

Sub DOF(dofevent, dofstate)
  If cNewController>2 Then
    If dofstate = 2 Then
      Controller.B2SSetData dofevent, 1:Controller.B2SSetData dofevent, 0
    Else
      Controller.B2SSetData dofevent, dofstate
    End If
  End If
End Sub

  Const cGameName = "bbb109"
  Const UseSolenoids = 1
  Const UseLamps = 0
  Const UseSync = 1
  Const HandleMech = 0

  'Standard Sounds
  Const SSolenoidOn = "Solenoid"
  Const SSolenoidOff = ""
  Const SCoin = "Coin"

'******Variables***************
  Dim xx
  Dim bsTrough, DTBank4, DTBank3, DTBank1, bsRHole
  Dim captb1, captb2, captb3, captb4
  Dim bsTHelp, bsSw46, bsSw47, bsSw48
  Dim bsALL, bsALR
  Dim xoff, yoff, zoff, xrot, zscale, xcen, ycen

'********Table Init************
  Sub Table_Init
  With Controller
    .GameName = cGameName
    .SplashInfoLine = "Big Bang Bar, Capcom 1996"
    .HandleMechanics = 0
    .HandleKeyboard = 0
    .ShowDMDOnly = 1
    .ShowFrame = 0
    .ShowTitle = 0
    .Hidden = DesktopMode
  End With
    On Error Resume Next
    Controller.Run
    If Err Then MsgBox Err.Description
      On Error Goto 0

  'Nudging
      vpmNudge.TiltSwitch=10
      vpmNudge.Sensitivity=5
      vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingshot)

'********Trough***************
  Set bsTrough = New cvpmBallStack
  With bsTrough
    .InitSw 35,36,37,38,39,0,0,0
    .InitKick BallRelease, 90, 8
    .InitExitSnd SoundFX("ballrelease"), SoundFX("Solenoid")
    .Balls = 4
  End With

'********Right Hole***********
  Set bsRHole = New cvpmBallStack
  With bsRHole
    .InitSaucer sw67, 67, 220, 18
    .KickZ = 0.4
    .InitExitSnd SoundFX("scoop_exit"), SoundFX("Solenoid")
    .KickForceVar = 2
  End With

     'Ball lock stacks
     Set bsSw46 = New cvpmBallStack
     With bsSw46
         .InitSw 0, 0, 0, 0, 0, 0, 0, 0
     End With

     Set bsSw47 = New cvpmBallStack
     With bsSw47
         .InitSw 0, 0, 0, 0, 0, 0, 0, 0
     End With

     Set bsSw48 = New cvpmBallStack
     With bsSw48
         .InitSw 0, 0, 0, 0, 0, 0, 0, 0
     End With

     Set bsTHelp = New cvpmBallStack  'To help if all 4 balls get caught in lower lock
     With bsTHelp
         .InitSw 0, 0, 0, 0, 0, 0, 0, 0
     End With

'********Alien Locks ***********
     Set bsALL = New cvpmBallStack
     With bsALL
         .InitSw 0, 0, 0, 0, 0, 0, 0, 0
         .InitKick sw61Out, 160, 1
     End With

     Set bsALR = New cvpmBallStack
     With bsALR
         .InitSw 0, 0, 0, 0, 0, 0, 0, 0
         .InitKick sw62Out, 190, 1
     End With

'********DropTargets
    Set DTBank4 = New cvpmDropTarget
      With DTBank4
      .InitDrop Array(sw17,sw18,sw19,sw20), Array(17,18,19,20)
        .Initsnd "", SoundFX("resetdrop")
       End With

    Set DTBank3 = New cvpmDropTarget
      With DTBank3
      .InitDrop Array(sw49,sw50,sw51), Array(49,50,51)
        .Initsnd "", SoundFX("resetdrop")
       End With

    Set DTBank1 = New cvpmDropTarget
      With DTBank1
      .InitDrop sw58, 58
        .Initsnd "", SoundFX("resetdrop")
       End With

'*******Captive balls
  ' Initialize Captive Ball System
  Controller.Switch(77) = True : Controller.Switch(78) = False : Controller.Switch(79) = True : Controller.Switch(80) = False
  BP1 = True : BP2 = True : BP3 = True : BP4 = True : BP1R = False : BP2R = False : BP3L = False : BP4L = False
  capmove1.enabled = true : capmove3.enabled = true : capmove3L.enabled = false: capmove1R.enabled = false
  capmove2.enabled = true : capmove4.enabled = true : capmove4L.enabled = true : capmove2R.enabled = true
  CreateBalls

'********Main Timer init
  PinMAMETimer.Interval = PinMAMEInterval
  PinMAMETimer.Enabled = 1

'********KickBack Init
  kickback.PullBack

'********Diverters Init
  DivLR.IsDropped=1

  setup_backglass()


  End Sub

Dim topoff

Sub setup_backglass()

xoff = 460
yoff =0
zoff = 740
xrot = -90

  if Table.ShowFSS = true then
  'Table.ColorGradeImage = "_ColorGrade_8"
  BBBbgDark.x = xoff
  BBBbgDark.y = yoff
  BBBbgDark.height = zoff
  BBBbgDark.rotx = xrot

  BBBbgHigh.x = xoff
  BBBbgHigh.y = yoff
  BBBbgHigh.height = zoff
  BBBbgHigh.rotx = xrot

  BBBbgHigh1.x = xoff
  BBBbgHigh1.y = yoff
  BBBbgHigh1.height = zoff
  BBBbgHigh1.rotx = xrot

  BBBbgFrame.x = xoff
  BBBbgFrame.y = yoff
  BBBbgFrame.height = zoff
  BBBbgFrame.rotx = xrot

  BBBbgFrameMask.x = xoff
  BBBbgFrameMask.y = yoff
  BBBbgFrameMask.height = zoff
  BBBbgFrameMask.rotx = xrot


' the topper
  topoff = 520
  If Table.inclination >= 65.0 then

    TopDark.x = xoff
    TopDark.y = yoff
    TopDark.height = zoff +topoff
    TopDark.rotx = xrot

    TopHigh.x = xoff
    TopHigh.y = yoff
    TopHigh.height = zoff +topoff
    TopHigh.rotx = xrot

    TopHigh1.x = xoff
    TopHigh1.y = yoff
    TopHigh1.height = zoff +topoff
    TopHigh1.rotx = xrot


    TopHigh2.x = xoff
    TopHigh2.y = yoff
    TopHigh2.height = zoff +(topoff-100)
    TopHigh2.rotx = xrot

  else
  end if


  DMD.x = xoff +20
  DMD.y = yoff
  DMD.height = zoff -280
  DMD.rotx = xrot

  DMD1.x = xoff +20
  DMD1.y = yoff +400
  DMD1.height = zoff -450
  DMD1.rotx = xrot +(-180)

  center_graphics()

  Else

  BBBbgDark.visible =0
  BBBbgHigh.visible =0
  BBBbgFrame.visible =0
  DMD.visible =0

  DMD1.x = xoff +20
  DMD1.y = yoff +400
  DMD1.height = zoff -450
  DMD1.rotx = xrot +(-180)

    if Table.ShowDT = true then
    'Table.ColorGradeImage = "ColorGradeLUT256x16_ConSat"
    DMD1.visible =0
    else
    'Table.ColorGradeImage = "_ColorGrade_8"
    end if

  hide_elem(Graphix(0))
  hide_elem(Graphix(1))


  end if

end sub


Sub hide_elem(elem)
Dim objx

  For Each objx In elem ' hide all elements
  objx.visible =0
  Next
end sub


Sub show_elem(elem)
Dim objx

  For Each objx In elem ' hide all elements
  objx.visible =1
  Next
end sub



Dim Graphix(2)

Graphix(0)  =Array(FlBbBTitle, BBBRightgirl,BBBLeftgirl, Flasher3)
Graphix(1)  =Array(FlBgNeonLamp, FlBgNeonLamp1 )

Sub center_graphics()
zscale = 0.09

xcen =(1042 /2) - (40 / 2)
ycen = (987 /2 ) + (150 /2)
Dim ii
Dim xx
Dim yy
Dim yfact
Dim xfact
Dim obj
yfact =50 'y fudge factor (ycen was wrong so fix)
xfact =0


for ii =0 to 1
  For Each obj In Graphix(ii)
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


   Sub Table_Paused:Controller.Pause = 1:End Sub
   Sub Table_unPaused:Controller.Pause = 0:End Sub
   Sub Table_Exit:Controller.Stop():End Sub

'***************************************
'        KEYS
'***************************************

 Sub Table_KeyDown(ByVal keycode)
  If Keycode = LeftFlipperKey then
    Controller.Switch(33) = 1
  End If
  If Keycode = RightFlipperKey then
    Controller.Switch(34) = 1: Controller.Switch(68) = 1
  End If
    If keycode = PlungerKey Then PlaySoundAtVol "PlungerPull", Plunger, 1:Plunger.Pullback
     If keycode = LeftTiltKey Then Playsound SoundFX("fx_nudge")
    If keycode = RightTiltKey Then Playsound SoundFX("fx_nudge")
    If keycode = CenterTiltKey Then Playsound SoundFX("nudge")
    If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub Table_KeyUp(ByVal keycode)
  If vpmKeyUp(keycode) Then Exit Sub
  If Keycode = LeftFlipperKey then
    Controller.Switch(33) = 0
  End If
  If Keycode = RightFlipperKey then
    Controller.Switch(34) = 0 : Controller.Switch(68) = 0
  End If
    If keycode = PlungerKey Then StopSound "PlungerPull":PlaySoundAtVol "Plunger", Plunger, 1:Plunger.Fire
End Sub

'******************************************
'Solenoids
'******************************************

  SolCallback(1) = "bsTrough.SolIn"               'OutHole
  SolCallback(2) = "bsTrough.SolOut"                'Ball Release
  SolCallback(3) = "vpmSolSound SoundFX(""knocker""),"      'Knocker
  SolCallback(4) = "vpmSolSound SoundFX(""SlingshotLeft""),"    'LeftSling
  SolCallback(5) = "vpmSolSound SoundFX(""SlingshotRight""),"   'RightSling
  SolCallback(6) = "SolKickBack"                  'KickBack
  SolCallback(7) = "sol4Bank"                   '4-Bank drop target
  SolCallback(8) = "SolLowerLockPin"                'Lower Lock Post
' SolCallback(9) = "SolLFlipper"                  'Left Flipper
' SolCallback(10) = "SolRFlipper"                 'Right Flipper
' SolCallback(11) = "SolURFlipper"                'Upper Flipper
  SolCallback(12) = "bsRHole.SolOut"                'Eject Hole
  SolCallback(13) = "SolLRDIvert"                 'Island Diverter
  SolCallback(14) = "SolRDivert1"                 'Ramp Diverter 1
  SolCallback(15) = "SolRDivert2"                 'Ramp Diverter 2
  SolCallback(16) = "SolRDivert3"                 'Alien Lock Post
  SolCallback(17) = "sol3Bank"                  '3-Bank drop target
  SolCallback(18) = "vpmSolSound SoundFX(""fx_bumper1""),"    'Left Bumper
  SolCallback(19) = "vpmSolSound SoundFX(""fx_bumper2""),"    'Middle Bumper
  SolCallback(20) = "vpmSolSound SoundFX(""fx_bumper3""),"    'Right Bumper
  SolCallback(21)  ="setlamp 161,"                'BackBox Left Flasher
  SolCallback(22)  ="setlamp 162,"                'Tube Dancer Flasher
  SolCallback(23)  ="setlamp 163,"                'Dance Floor Flasher
  SolCallback(24)  ="setlamp 164,"                'Eject Hole Flasher
  SolCallback(25)  ="setlamp 165,"                'Alien Lock Flasher
  SolCallback(26)  ="setlamp 166,"                'Lower Lock Flasher
  SolCallback(27) = "GateLeft"                  'Loop Left Gate
  SolCallback(28) = "GateRight"                 'Loop Right Gate
  SolCallback(29) = "sol1Bank"                  '1-Bank drop target
  SolCallback(30) = "solDancer"                 'Tube Dancer Motor
  SolCallback(31) = "SolAlienForward"               'Alien Motor Forward
  SolCallback(32)  ="SolAlienReverse"               'ALien Motor Reverse

'******************************************
'        Flippers
'******************************************

'******Flipper fix stuff till emulation is fixed

dim GameOnFF:GameOnFF = 0
dim BIPFF:BIPFF = 0
'
  Sub FFTimer_Timer()
  If bsTrough.Balls = 4 or RightSlingshot.SlingshotStrength = 0 or BIPFF = 0 Then
    GameOnFF = 0
  else
    GameOnFF = 1
  end if
  End Sub

 SolCallback(sLRFlipper) = "SolRFlipper"
 SolCallback(sLLFlipper) = "SolLFlipper"

  Sub SolLFlipper(Enabled)
    If Enabled and GameOnFF = 1 Then
    PlaySoundAtVol SoundFX("fx_flipperup"), LeftFlipper, VolFlip
    LeftFlipper.RotateToEnd
     Else
     PlaySoundAtVol SoundFX("fx_flipperdown"), LeftFlipper, VolFlip
     LeftFlipper.RotateToStart
     End If
  End Sub

  Sub SolRFlipper(Enabled)
     If Enabled and GameOnFF = 1 Then
    PlaySoundAtVol SoundFX("fx_flipperup"), RightFlipper, VolFlip
     RightFlipper.RotateToEnd
     RightFlipper1.RotateToEnd
     Else
     PlaySoundAtVol SoundFX("fx_flipperdown"), RightFlipper, VolFlip
     RightFlipper.RotateToStart
     RightFlipper1.RotateToStart
     End If
  End Sub

'******************************************
'Drop Targets
'******************************************

Dim tdir

  Sub sw17_Hit:DTBank4.Hit 1:tdir = -1:me.TimerEnabled = 1:End Sub
  Sub sw18_Hit:DTBank4.Hit 2:tdir = -1:me.TimerEnabled = 1:End Sub
  Sub sw19_Hit:DTBank4.Hit 3:tdir = -1:me.TimerEnabled = 1:End Sub
  Sub sw20_Hit:DTBank4.Hit 4:tdir = -1:me.TimerEnabled = 1:End Sub

  Sub sw49_Hit:DTBank3.Hit 1:tdir = -1:me.TimerEnabled = 1:End Sub
  Sub sw50_Hit:DTBank3.Hit 2:tdir = -1:me.TimerEnabled = 1:End Sub
  Sub sw51_Hit:DTBank3.Hit 3:tdir = -1:me.TimerEnabled = 1:End Sub
  Sub sw58_Hit:DTBank1.Hit 1:tdir = -1:me.TimerEnabled = 1:End Sub

   Sub sol4Bank(Enabled)
    If Enabled Then
      tdir=1
      sw17.TimerEnabled = 1:sw18.TimerEnabled = 1:sw19.TimerEnabled = 1:sw20.TimerEnabled = 1
      DTBank4.DropSol_On
    End if
   End Sub

   Sub sol3Bank(Enabled)
    If Enabled Then
      tdir=1
      sw49.TimerEnabled = 1:sw50.TimerEnabled = 1:sw51.TimerEnabled = 1
      DTBank3.DropSol_On
    End if
   End Sub

   Sub sol1Bank(Enabled)
    If Enabled Then
      tDir = 1
      sw58.TimerEnabled = 1
      DTBank1.DropSol_On
    End if
   End Sub

  Sub sw17_Timer()
  sw17p.z=sw17p.z+1*tdir
  If sw17p.z < -50 then sw17p.z =-50:me.timerenabled = 0
  If sw17p.z>0 then sw17p.z=0:me.timerenabled = 0
  End Sub

  Sub sw18_Timer()
  sw18p.z=sw18p.z+1*tdir
  If sw18p.z < -50 then sw18p.z =-50:me.timerenabled = 0
  If sw18p.z>0 then sw18p.z=0:me.timerenabled = 0
  End Sub

  Sub sw19_Timer()
  sw19p.z=sw19p.z+1*tdir
  If sw19p.z < -50 then sw19p.z =-50:me.timerenabled = 0
  If sw19p.z>0 then sw19p.z=0:me.timerenabled = 0
  End Sub

  Sub sw20_Timer()
  sw20p.z=sw20p.z+1*tdir
  If sw20p.z < -50 then sw20p.z =-50:me.timerenabled = 0
  If sw20p.z>0 then sw20p.z=0:me.timerenabled = 0
  End Sub

  Sub sw49_Timer()
  sw49p.z=sw49p.z+1*tdir
  If sw49p.z < -50 then sw49p.z =-50:me.timerenabled = 0
  If sw49p.z>0 then sw49p.z=0:me.timerenabled = 0
  End Sub

  Sub sw50_Timer()
  sw50p.z=sw50p.z+1*tdir
  If sw50p.z < -50 then sw50p.z =-50:me.timerenabled = 0
  If sw50p.z>0 then sw50p.z=0:me.timerenabled = 0
  End Sub

  Sub sw51_Timer()
  sw51p.z=sw51p.z+1*tdir
  If sw51p.z < -50 then sw51p.z =-50:me.timerenabled = 0
  If sw51p.z>0 then sw51p.z=0:me.timerenabled = 0
  End Sub

  Sub sw58_Timer()
  sw58p.z=sw58p.z+1*tdir
  If sw58p.z < -50 then sw58p.z =-50:me.timerenabled = 0
  If sw58p.z>0 then sw58p.z=0:me.timerenabled = 0
  End Sub

'******************************************
'        DANCER ANIMATION
'******************************************
dim dance

sub solDancer(enabled)
if enabled then
dancerT.enabled=1
else
dancerT.enabled=0
if dancer1.visible=0 then
dancer1.visible=1
dancer2.visible=0
dancer3.visible=0
dancer4.visible=0
dancer5.visible=0
dancer6.visible=0
dancer7.visible=0
dancer8.visible=0
dancer9.visible=0
dancer10.visible=0
end if
end if
end sub

sub dancerT_timer()
select case dance
case 1:
dancer1.visible=1
dancer2.visible=0
dancer3.visible=0
dancer4.visible=0
dancer5.visible=0
dancer6.visible=0
dancer7.visible=0
dancer8.visible=0
dancer9.visible=0
dancer10.visible=0
mechpost.rotx=0

case 2
dancer1.visible=0
dancer2.visible=1
dancer3.visible=0
dancer4.visible=0
dancer5.visible=0
dancer6.visible=0
dancer7.visible=0
dancer8.visible=0
dancer9.visible=0
dancer10.visible=0
mechpost.rotx=2

case 3
dancer1.visible=0
dancer2.visible=0
dancer3.visible=1
dancer4.visible=0
dancer5.visible=0
dancer6.visible=0
dancer7.visible=0
dancer8.visible=0
dancer9.visible=0
dancer10.visible=0
mechpost.rotx=3

case 4
dancer1.visible=0
dancer2.visible=0
dancer3.visible=0
dancer4.visible=1
dancer5.visible=0
dancer6.visible=0
dancer7.visible=0
dancer8.visible=0
dancer9.visible=0
dancer10.visible=0
mechpost.rotx=6

case 5
dancer1.visible=0
dancer2.visible=0
dancer3.visible=0
dancer4.visible=0
dancer5.visible=1
dancer6.visible=0
dancer7.visible=0
dancer8.visible=0
dancer9.visible=0
dancer10.visible=0
mechpost.rotx=4

case 6
dancer1.visible=0
dancer2.visible=0
dancer3.visible=0
dancer4.visible=0
dancer5.visible=0
dancer6.visible=1
dancer7.visible=0
dancer8.visible=0
dancer9.visible=0
dancer10.visible=0
mechpost.rotx=3

case 7
dancer1.visible=0
dancer2.visible=0
dancer3.visible=0
dancer4.visible=0
dancer5.visible=0
dancer6.visible=0
dancer7.visible=1
dancer8.visible=0
dancer9.visible=0
dancer10.visible=0
mechpost.rotx=6

case 8
dancer1.visible=0
dancer2.visible=0
dancer3.visible=0
dancer4.visible=0
dancer5.visible=0
dancer6.visible=0
dancer7.visible=0
dancer8.visible=1
dancer9.visible=0
mechpost.rotx=4

case 9
dancer1.visible=0
dancer2.visible=0
dancer3.visible=0
dancer4.visible=0
dancer5.visible=0
dancer6.visible=0
dancer7.visible=0
dancer8.visible=0
dancer9.visible=1
mechpost.rotx=2

case 10
dancer1.visible=0
dancer2.visible=0
dancer3.visible=0
dancer4.visible=0
dancer5.visible=0
dancer6.visible=0
dancer7.visible=0
dancer8.visible=0
dancer9.visible=0
dancer10.visible=1
mechpost.rotx=0
dance=1
End select
dance=dance+1
end sub

'******************************************
'         Gates
'******************************************

 dim GateStateL, GateStateR : GateStateL = 0 : GateStateR = 0
 '
 Sub GateLeft(enabled)  : If enabled then : GateL.Open = true : GateStateL = 1 : vpmTimer.AddTimer 1000, "GateCloseL" : End If : End Sub
 Sub GateCloseL(aSw)    : GateL.Open = false  : GateStateL = 0 : End Sub

 Sub GateRight(enabled) : If enabled then : GateR.Open = true : GateStateR = 1 : vpmTimer.AddTimer 1000, "GateCloseR" : End If : End Sub
 Sub GateCloseR(aSw)    : GateR.Open = false : GateStateR = 0 : End Sub


'******************************************
'     Captive ball script
'   based on pacdudes script
'******************************************

  dim mBallDir, mBallCos, mBallSin
  dim mTrig, mWall, mKickers, mVelX, mVelY, mVelX2, mVelY2
  dim ForceTransL, ForceTransR, MinForce
  dim BP1,BP2,BP3,BP4,BP1R,BP2R,BP3L,BP4L

'   Default Settings
  ForceTransL = 0.4 : ForceTransR = 0.6 : MinForce = 5 : mBallDir = 10
  mBallCos = Cos(mBallDir * 3.1415927/180) : mBallSin = Sin(mBallDir * 3.1415927/180)

' Trigger Leads (Left/Right Side)
  Sub CapLLead_Hit() : mVelX  = activeBall.VelX : mVelY  = activeBall.VelY : End Sub
  Sub CapRLead_Hit() : mVelX2 = activeBall.VelX : mVelY2 = activeBall.VelY : End Sub

'   Left Side Captive Ball Hits (Up Motion)
  Sub CapLHit_Hit()
    dim force, dx, dy
    dX = ActiveBall.X - Captive1.X : dY = ActiveBall.Y - Captive1.Y
    force = -ForceTransL * (dY * mVelY + dX * mVelX) * (dY * mBallCos + dX * mBallSin) / (dX*dX + dY*dY)
    If force < 1 Then Exit Sub
    If force < MinForce Then force = MinForce
    If BP3L Then : BP3L = False : ForceTransL = 0.2 :PlaySoundAtVol "fx_collide", ActiveBall, 1: Captive3L.DestroyBall : CapMove3L.CreateBall : CapMove3L.Kick mBallDir, force : Else :_
    If BP4L Then : BP4L = False : ForceTransL = 0.3 : Controller.Switch(78) = False :PlaySoundAtVol "fx_collide", ActiveBall, 1: Captive4L.DestroyBall : CapMove4L.CreateBall : CapMove3L.enabled = False : CapMove4L.Kick mBallDir, force : Else :_
    If BP2  Then : BP2  = False : ForceTransL = 0.4:PlaySoundAtVol "fx_collide", ActiveBall, 1: Captive2.DestroyBall : CapMove2.CreateBall : CapMove4L.enabled = False : CapMove2.Kick mBallDir, force : Else :_
    If BP1  Then : BP1  = False : Controller.Switch(77) = False :PlaySoundAtVol "fx_collide", ActiveBall, 1: Captive1.DestroyBall : CapMove1.CreateBall : CapMove2.enabled = False : CapMove1.Kick mBallDir, force :_
    End If : End If : End If : End If
  End Sub

'   Right Side Captive Ball Hits (Up Motion)
  Sub CapRHit_Hit()
    dim force2, dx2, dy2
    dX2 = ActiveBall.X - Captive3.X : dY2 = ActiveBall.Y - Captive3.Y
    force2 = -ForceTransR * (dY2 * mVelY2 + dX2 * mVelX2) * (dY2 * mBallCos + dX2 * mBallSin) / (dX2*dX2 + dY2*dY2)
    If force2 < 1 Then Exit Sub
    If force2 < MinForce Then force2 = MinForce
    If BP1R Then : BP1R = False : ForceTransR = 0.2 :PlaySoundAtVol "fx_collide", ActiveBall, 1: Captive1R.DestroyBall : CapMove1R.CreateBall : CapMove1R.Kick mBallDir, force2 : Else :_
    If BP2R Then : BP2R = False : ForceTransR = 0.3 : Controller.Switch(80) = False :PlaySoundAtVol "fx_collide", ActiveBall, 1: Captive2R.DestroyBall : CapMove2R.CreateBall : CapMove1R.enabled = False : CapMove2R.Kick mBallDir, force2 : Else :_
    If BP4  Then : BP4  = False : ForceTransR = 0.4 :PlaySoundAtVol "fx_collide", ActiveBall, 1: Captive4.DestroyBall : CapMove4.CreateBall : CapMove2R.enabled = False : CapMove4.Kick mBallDir, force2 : Else :_
    If BP3  Then : BP3  = False : Controller.Switch(79) = False :PlaySoundAtVol "fx_collide", ActiveBall, 1: Captive3.DestroyBall : CapMove3.CreateBall : CapMove4.enabled = False : CapMove3.Kick mBallDir, force2 :_
    End If : End If : End If : End If
  End Sub

'   Left Side Return Hits (Down Motion)
  Sub CapMove1_Hit()
  If BP1 = True Then
   CapMove1.DestroyBall : CapMove2_Hit
  Else
    BP1  = True : Controller.Switch(77) = True : CapMove1.DestroyBall : Captive1.CreateBall : CapMove2.enabled = True
    End If
  End Sub
  Sub CapMove2_Hit()
  If BP2 = True Then
   CapMove2.DestroyBall : CapMove4L_Hit
  Else
    BP2  = True : ForceTransL = 0.4 : CapMove2.DestroyBall : Captive2.CreateBall : CapMove4L.enabled = True
    end if
  End Sub
  Sub CapMove4L_Hit()
  If BP4L = True Then
   CapMove4L.DestroyBall : CapMove3L_Hit
  Else
    BP4L = True : ForceTransL = 0.3 : Controller.Switch(78) = True : CapMove4L.DestroyBall : Captive4L.CreateBall : CapMove3L.enabled = True
    End If
  End Sub
  Sub CapMove3L_Hit()
    BP3L = True : ForceTransL = 0.2 : CapMove3L.DestroyBall : Captive3L.CreateBall
  End Sub

' Right Side Return Hits (Down Motion)
  Sub CapMove3_Hit()
  If BP3 = True Then
   CapMove3.DestroyBall : CapMove4_Hit
  Else
    BP3  = True : Controller.Switch(79) = True : CapMove3.DestroyBall : Captive3.CreateBall : CapMove4.enabled = True
    End If
  End Sub
  Sub CapMove4_Hit()
  If BP4 = True Then
   CapMove4.DestroyBall : CapMove2R_Hit
  Else
    BP4  = True : ForceTransR = 0.4 : CapMove4.DestroyBall : Captive4.CreateBall : CapMove2R.enabled = True
    End If
  End Sub
  Sub CapMove2R_Hit()
  If BP2R = True Then
   CapMove2R.DestroyBall : CapMove1R_Hit
  Else
    BP2R = True : ForceTransR = 0.3 : Controller.Switch(80) = True : CapMove2R.DestroyBall : Captive2R.CreateBall : CapMove1R.enabled = True
    End If
  End Sub
  Sub CapMove1R_Hit()
    BP1R = True : ForceTransR = 0.2 : CapMove1R.DestroyBall : Captive1R.CreateBall
  End Sub

'   Destroy any Existing Balls and Create the 4 Captive Balls (either real or reel-based)
  Sub CreateBalls()
  Captive1.DestroyBall  : Captive2.DestroyBall  : Captive3.DestroyBall  : Captive4.DestroyBall
  Captive4L.DestroyBall : Captive3L.DestroyBall : Captive1R.DestroyBall : Captive2R.DestroyBall
    If BP1  Then : Captive1.CreateBall  : End If : If BP2  Then : Captive2.CreateBall  : End If : If BP4L Then : Captive4L.CreateBall : End If
    If BP3L Then : Captive3L.CreateBall : End If : If BP3  Then : Captive3.CreateBall  : End If : If BP4  Then : Captive4.CreateBall  : End If
    If BP2R Then : Captive2R.CreateBall : End If : If BP1R Then : Captive1R.CreateBall : End If
End Sub

'******************************************
'      LowerLock And Post Handling
'******************************************

dim postdwn : postdwn = 0
  Sub SolLowerLockPin(Enabled)
  if enabled then
    If lockPin.IsDropped=0 Then:PlaysoundAtVol SoundFX("diverter"), sw45,1:End If
    If postdwn = 0 Then vpmTimer.AddTimer 500, "PostUp"
    lockPin.IsDropped=1:postdwn = 1
    If bsSw48.Balls>0 Then
      sw48H.DestroyBall:sw48.CreateBall:bsSw48.SolOut 1:sw48.Kick 210, 1
      If bsSw48.Balls = 0 Then
        controller.switch(48) = 0
      End If
    End If
    If bsSw47.Balls>0 Then
      sw47H.DestroyBall:sw47.CreateBall:bsSw47.SolOut 1:sw47.Kick 180, 1
      If bsSw47.Balls = 0 Then
        controller.switch(47) = 0
      End If
    End If
    If bsSw46.Balls>0 Then
      sw46H.DestroyBall:sw46.CreateBall:bsSw46.SolOut 1:sw46.Kick 180, 1
      If bsSw46.Balls = 0 Then
        controller.switch(46) = 0
      End If
    End If
    If bsTHelp.Balls>0 Then
      THelpH.DestroyBall:THelp.CreateBall:bsTHelp.SolOut 1:THelp.Kick 180, 1
    End If
  End If
  End Sub

  Sub PostUp(aSw)
  postdwn = 0:lockPin.IsDropped=0
  End Sub

  Sub swTHelp_Hit()
  Dim BSpeed
' BIPFF = BIPFF - 1
  If bsTHelp.Balls>0 Then
    swTHelp.DestroyBall:sw46_Hit
  Else
    BSpeed = ActiveBall.VelY
    If bsSw46.Balls = 0 then
      swTHelp.Kick 180, BSpeed
    Else
      swTHelp.DestroyBall:swTHelpH.CreateBall:bsTHelp.AddBall 1
    End If
  End If
  End Sub

  Sub sw46_Hit()
  Dim BSpeed
  If bsSw46.Balls>0 Then
    sw46.DestroyBall:sw48_Hit
  Else
    BSpeed = ActiveBall.VelY
    If bsSw47.Balls = 0 then
      vpmTimer.PulseSw 46
      sw46.Kick 180, BSpeed
    Else
      controller.switch(46) = 1
      sw46.DestroyBall:sw46H.CreateBall:bsSw46.AddBall 1
    End If
  End If
  End Sub

  Sub sw47_Hit()
  Dim BSpeed
  If bsSw47.Balls>0 Then
    sw47.DestroyBall:sw46_Hit
  Else
    BSpeed = ActiveBall.VelY
    If bsSw48.Balls = 0 then
      vpmTimer.PulseSw 47
      sw47.Kick 180, BSpeed
    Else
      controller.switch(47) = 1
      sw47.DestroyBall:sw47H.CreateBall:bsSw47.AddBall 1
    End If
  End If
  End Sub

  Sub sw48_Hit()
  Dim BSpeed
  BSpeed = ActiveBall.VelY
  If bsSw48.Balls>0 and bsSw47.Balls>0 and bsSw46.Balls>0 Then
    sw48.Kick 210, BSpeed
  Else
    If bsSw48.Balls > 0 then
      sw48.DestroyBall:sw47_Hit
    Else
      controller.switch(48) = 1
      sw48.DestroyBall:sw48H.CreateBall:bsSw48.AddBall 1
    End If
  End If
  End Sub

'  Sub sw48_UnHit():BIPFF = BIPFF + 1:End Sub

'******************************************
'      Left Kickback
'******************************************

  Sub SolKickBack(enabled)
  If enabled then
    PlaySound SoundFX("Solenoid")
    kickback.Fire
  else
    kickback.PullBack
  End If
  End Sub

'******************************************
'      Ramp diverters
'******************************************

  Sub SolLRDIvert(Enabled)
  If Enabled then
    debug.print "Multiball div on"
    DivLR.IsDropped=0
  PlaySound SoundFX("Diverter")
  Else
    debug.print "Multiball div off"
    DivLR.IsDropped=1
  End If
  End Sub

  Sub SolRDivert1(Enabled)
  If Enabled then
    debug.print "Tube div on"
    DivTubef.RotateToEnd
    DivTube.isDropped=1
    PlaySound SoundFX("Diverter")
    vpmTimer.AddTimer 2000, "UnRDivert"
  End If
  End Sub
  Sub UnRDivert(aSw)
    debug.print "Tube div off"
    DivTubef.RotateToStart
    DivTube.isDropped=0
  End Sub

Sub SolRDivert2(Enabled)
  If Enabled then
    debug.print "Aliens div on"
    DivTube2f.RotateToEnd
    DivTube2.isDropped=1
  PlaySound SoundFX("Diverter")
  Else
    debug.print "Aliens div off"
    DivTube2f.RotateToStart
    DivTube2.isDropped=0
  End If
  End Sub

  Sub SolRDivert3(Enabled)
  If Enabled then
    debug.print "L Alien div on"
    DivLock.IsDropped=1
  PlaySound SoundFX("Diverter")
  Else
    debug.print "L Alien div on"
    DivLock.IsDropped=0
  End If
  End Sub

'******************************************
' Alien Lock Mech Handler
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

Sub ALockStartup_Timer() : If direction = 0 then : If checkstart = 0 and startstatus = 1 then : vpmTimer.AddTimer 1000, "CheckStatus" : End If : checkstart = 1 : Else : checkstart = 0 : End If : End Sub
Sub CheckStatus(aSw) : if checkstart = 1 Then : startstatus = 0 : OldPos = 31 : NewPos = 0 : End If : End Sub
dim NewPos, forward, reverse, OldPos : OldPos = 31 : NewPos = 0 : forward = 0 : reverse = 0
'dim RAlien:RAlien = Array(RM0,RM1,RM2,RM3,RM4,RM5,RM6,RM7,RM8,RM9)
'dim LAlien:LAlien = Array(RM10,RM19,RM18,RM17,RM16,RM15,RM14,RM13,RM12,RM11)

dim direction
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
    Case  2 : Alien1.rotZ=-255:Alien2.rotZ=345    'AlienReel.setvalue 8
    Case  3 : Alien1.rotZ=-270:Alien2.rotZ=0      'AlienReel.setvalue 9
    Case  5 : Alien1.rotZ=-290:Alien2.rotZ=20     'AlienReel.setvalue 9
    Case  6 : Alien1.rotZ=-315:Alien2.rotZ=45     'AlienReel.setvalue 0
    Case  9 : Alien1.rotZ=-330:Alien2.rotZ=60     'AlienReel.setvalue 0
    Case 10 : Alien1.rotZ=0:Alien2.rotZ=90      'AlienReel.setvalue 1
    Case 11 : Alien1.rotZ=-15:Alien2.rotZ=105     'AlienReel.setvalue 1
    Case 12 : Alien1.rotZ=-30:Alien2.rotZ=120     'AlienReel.setvalue 2
    Case 15 : Alien1.rotZ=-45:Alien2.rotZ=135     'AlienReel.setvalue 2
    Case 16 : Alien1.rotZ=-60:Alien2.rotZ=150     'AlienReel.setvalue 3
    Case 18 : Alien1.rotZ=-75:Alien2.rotZ=165     'AlienReel.setvalue 3
    Case 19 : Alien1.rotZ=-90:Alien2.rotZ=180     'AlienReel.setvalue 4
    Case 21 : Alien1.rotZ=-120:Alien2.rotZ=200    'AlienReel.setvalue 4
    Case 22 : Alien1.rotZ=-135:Alien2.rotZ=225    'AlienReel.setvalue 5
    Case 24 : Alien1.rotZ=-160:Alien2.rotZ=250    'AlienReel.setvalue 5
    Case 25 : Alien1.rotZ=-180:Alien2.rotZ=270    'AlienReel.setvalue 6
    Case 27 : Alien1.rotZ=-195:Alien2.rotZ=285    'AlienReel.setvalue 6
    Case 28 : Alien1.rotZ=-210:Alien2.rotZ=300    'AlienReel.setvalue 7
    Case 30 : Alien1.rotZ=-225:Alien2.rotZ=315    'AlienReel.setvalue 7
    Case 31 : Alien1.rotZ=-240:Alien2.rotZ=330    'AlienReel.setvalue 8

'   Case  2 : RAlien(7).IsDropped=1:RAlien(9).IsDropped=1:RAlien(8).IsDropped=0  'AlienReel.setvalue 8
'         LAlien(7).IsDropped=1:LAlien(9).IsDropped=1:LAlien(8).IsDropped=0
'   Case  3 : RAlien(8).IsDropped=1:RAlien(0).IsDropped=1:RAlien(9).IsDropped=0   'AlienReel.setvalue 9
'         LAlien(8).IsDropped=1:LAlien(0).IsDropped=1:LAlien(9).IsDropped=0
'   Case  5 : RAlien(0).IsDropped=1:RAlien(8).IsDropped=1:RAlien(9).IsDropped=0   'AlienReel.setvalue 9
'         LAlien(0).IsDropped=1:LAlien(8).IsDropped=1:LAlien(9).IsDropped=0
'   Case  6 : RAlien(9).IsDropped=1:RAlien(1).IsDropped=1:RAlien(0).IsDropped=0   'AlienReel.setvalue 0
'         LAlien(9).IsDropped=1:LAlien(1).IsDropped=1:LAlien(0).IsDropped=0
'   Case  9 : RAlien(9).IsDropped=1:RAlien(1).IsDropped=1:RAlien(0).IsDropped=0   'AlienReel.setvalue 0
'         LAlien(9).IsDropped=1:LAlien(1).IsDropped=1:LAlien(0).IsDropped=0
'   Case 10 : RAlien(2).IsDropped=1:RAlien(0).IsDropped=1:RAlien(1).IsDropped=0   'AlienReel.setvalue 1
'         LAlien(2).IsDropped=1:LAlien(0).IsDropped=1:LAlien(1).IsDropped=0
'   Case 11 : RAlien(2).IsDropped=1:RAlien(0).IsDropped=1:RAlien(1).IsDropped=0   'AlienReel.setvalue 1
'         LAlien(2).IsDropped=1:LAlien(0).IsDropped=1:LAlien(1).IsDropped=0
'   Case 12 : RAlien(1).IsDropped=1:RAlien(3).IsDropped=1:RAlien(2).IsDropped=0   'AlienReel.setvalue 2
'         LAlien(1).IsDropped=1:LAlien(3).IsDropped=1:LAlien(2).IsDropped=0
'   Case 15 : RAlien(1).IsDropped=1:RAlien(3).IsDropped=1:RAlien(2).IsDropped=0   'AlienReel.setvalue 2
'         LAlien(1).IsDropped=1:LAlien(3).IsDropped=1:LAlien(2).IsDropped=0
'   Case 16 : RAlien(2).IsDropped=1:RAlien(4).IsDropped=1:RAlien(3).IsDropped=0   'AlienReel.setvalue 3
'         LAlien(2).IsDropped=1:LAlien(4).IsDropped=1:LAlien(3).IsDropped=0
'   Case 18 : RAlien(2).IsDropped=1:RAlien(4).IsDropped=1:RAlien(3).IsDropped=0   'AlienReel.setvalue 3
'         LAlien(2).IsDropped=1:LAlien(4).IsDropped=1:LAlien(3).IsDropped=0
'   Case 19 : RAlien(3).IsDropped=1:RAlien(5).IsDropped=1:RAlien(4).IsDropped=0   'AlienReel.setvalue 4
'         LAlien(3).IsDropped=1:LAlien(5).IsDropped=1:LAlien(4).IsDropped=0
'   Case 21 : RAlien(3).IsDropped=1:RAlien(5).IsDropped=1:RAlien(4).IsDropped=0   'AlienReel.setvalue 4
'         LAlien(3).IsDropped=1:LAlien(5).IsDropped=1:LAlien(4).IsDropped=0
'   Case 22 : RAlien(4).IsDropped=1:RAlien(6).IsDropped=1:RAlien(5).IsDropped=0   'AlienReel.setvalue 5
'         LAlien(4).IsDropped=1:LAlien(6).IsDropped=1:LAlien(5).IsDropped=0
'   Case 24 : RAlien(4).IsDropped=1:RAlien(6).IsDropped=1:RAlien(5).IsDropped=0   'AlienReel.setvalue 5
'         LAlien(4).IsDropped=1:LAlien(6).IsDropped=1:LAlien(5).IsDropped=0
'   Case 25 : RAlien(5).IsDropped=1:RAlien(7).IsDropped=1:RAlien(6).IsDropped=0   'AlienReel.setvalue 6
'         LAlien(5).IsDropped=1:LAlien(7).IsDropped=1:LAlien(6).IsDropped=0
'   Case 27 : RAlien(5).IsDropped=1:RAlien(7).IsDropped=1:RAlien(6).IsDropped=0   'AlienReel.setvalue 6
'         LAlien(5).IsDropped=1:LAlien(7).IsDropped=1:LAlien(6).IsDropped=0
'   Case 28 : RAlien(6).IsDropped=1:RAlien(8).IsDropped=1:RAlien(7).IsDropped=0   'AlienReel.setvalue 7
'         LAlien(6).IsDropped=1:LAlien(8).IsDropped=1:LAlien(7).IsDropped=0
'   Case 30 : RAlien(6).IsDropped=1:RAlien(8).IsDropped=1:RAlien(7).IsDropped=0   'AlienReel.setvalue 7
'         LAlien(6).IsDropped=1:LAlien(8).IsDropped=1:LAlien(7).IsDropped=0
'   Case 31 : RAlien(7).IsDropped=1:RAlien(9).IsDropped=1:RAlien(8).IsDropped=0   'AlienReel.setvalue 8
'         LAlien(7).IsDropped=1:LAlien(9).IsDropped=1:LAlien(8).IsDropped=0
  End Select

  'Eject Lock
  If (NewPos >= 12 and NewPos <= 15) and (bsALL.balls > 0 or bsALR.balls > 0) Then
    AlienEjectR : vpmTimer.Addtimer 300, "AlienEjectL"
  End If
End If
  OldPos = NewPos
  If direction = 1 Then
    If forward = 0 and reverse = 0 and NewPos >= 6 and NewPos <= 9 Then : Direction = 0 : ALockTimer.enabled = False : StopSound "Motor": End If
  Else
    If forward = 0 and reverse = 0 Then : Direction = 0 : ALockTimer.enabled = False : End If
  End If
End Sub

  ' Eject Ball Subs
Sub AlienEjectL(aSw) : bsALL.SolOut 1 : End Sub
Sub AlienEjectR()    : bsALR.SolOut 1 : End Sub

 Sub sw61Out_UnHit():BIPFF = BIPFF + 1:End Sub
 Sub sw62Out_UnHit():BIPFF = BIPFF + 1:End Sub

'******************************************
'     Drains and Kickers
'******************************************

  Sub Drain_Hit():PlaySoundAtVol "Drain", ActiveBall, 1:BIPFF = BIPFF - 1:bsTrough.AddBall Me:End Sub
  Sub BallRelease_UnHit():BIPFF = BIPFF + 1:End Sub
  Sub sw61_Hit():BIPFF = BIPFF - 1:bsALL.AddBall Me:End Sub
  Sub sw62_Hit():BIPFF = BIPFF - 1:bsALR.AddBall Me:End Sub
  Sub sw67_Hit:PlaysoundAtVol "kicker_enter", ActiveBall, 1:bsRHole.AddBall Me:End Sub

'******************************************
' JF's Positional SlingShot Strength
'******************************************

Dim RStep, Lstep

 Dim LeftSlingYMiddle, RightSlingYMiddle
 Dim LeftSlingYRange, RightSlingYRange
 Dim LSYPosFactor, RSYPosFactor
 Dim LSBallAngle, LSAngleFactorY
 Dim RSBallAngle, RSAngleFactorY
 Dim SlingYPosEffectStrength:SlingYPosEffectStrength=1.67     'Use value between 1 and 2; 1 is light effect and 2 is strong effect

 Const LeftSlingYTop=1480.926     'Enter values from table control points for top and bottom Y coordinate for left and right sling
 Const LeftSlingYBottom=1673.57
 Const RightSlingYTop=1485.802
 Const RightSlingYBottom=1678.224

 LeftSlingYMiddle=(LeftSlingYBottom+LeftSlingYTop)/2
 RightSlingYMiddle=(RightSlingYBottom+RightSlingYTop)/2
 LeftSlingYRange=(LeftSlingYBottom-LeftSlingYTop)/2
 RightSlingYRange=(RightSlingYBottom-RightSlingYTop)/2


 Sub LeftSlingShot_Slingshot
  'Positional SlingShot Strength Routine Related
  GetAngle Cint(ActiveBall.VelX),Cint(ActiveBall.VelY), LSBallAngle
  LSAngleFactorY=Round(ABS(cos(LSBallAngle)),3)
  LSYPosFactor=(ABS(LeftSlingYMiddle-activeball.y))/(LeftSlingYRange/SlingYPosEffectStrength)
  DampenXY 2, (.3+LSYPosFactor), 5, 2, (.5+LSYPosFactor*LSAngleFactorY), 5
  'End Positional SlingShot Strength Routine Related
  vpmTimer.PulseSw 41:LStep = 0:Me.TimerEnabled = 1
  LSling.Visible = 0
  LSling1.Visible = 1
  sling1.TransZ = -27
  End Sub

 Sub RightSlingShot_Slingshot
  'Positional SlingShot Strength Routine Related
  GetAngle Cint(ActiveBall.VelX),Cint(ActiveBall.VelY), RSBallAngle
  RSAngleFactorY=Round(ABS(cos(RSBallAngle)),3)
  RSYPosFactor=(ABS(RightSlingYMiddle-activeball.y))/(RightSlingYRange/SlingYPosEffectStrength)
  DampenXY 2, (.3+RSYPosFactor), 5, 2, (.5+RSYPosFactor*RSAngleFactorY), 5
  'End Positional SlingShot Strength Routine Related
  vpmTimer.PulseSw 42:RStep = 0:Me.TimerEnabled = 1
  RSling.Visible = 0
  RSling1.Visible = 1
  sling2.TransZ = -27
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
'     Bumpers Animation
'************************************************

Dim dirRing1:dirRing1 = -1
Dim dirRing2:dirRing2 = -1
Dim dirRing3:dirRing3 = -1

Sub Bumper1_Hit:vpmTimer.PulseSw 56:me.TimerEnabled = 1:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 54:me.TimerEnabled = 1:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 55:me.TimerEnabled = 1:End Sub

Sub Bumper1_timer()
  Bumper1Ring.Z = Bumper1Ring.Z + (5 * dirRing1)
  If Bumper1Ring.Z <= -40 Then dirRing1 = 1
  If Bumper1Ring.Z >= 0 Then
    dirRing1 = -1
    Bumper1Ring.Z = 0
    Me.TimerEnabled = 0
  End If
End Sub

Sub Bumper2_timer()
  Bumper2Ring.Z = Bumper2Ring.Z + (5 * dirRing2)
  If Bumper2Ring.Z <= -40 Then dirRing2 = 1
  If Bumper2Ring.Z >= 0 Then
    dirRing2 = -1
    Bumper2Ring.Z = 0
    Me.TimerEnabled = 0
  End If
End Sub

Sub Bumper3_timer()
  Bumper3Ring.Z = Bumper3Ring.Z + (5 * dirRing3)
  If Bumper3Ring.Z <= -40 Then dirRing3 = 1
  If Bumper3Ring.Z >= 0 Then
    dirRing3 = -1
    Bumper3Ring.Z = 0
    Me.TimerEnabled = 0
  End If
End Sub

'************************************************
'     Switches
'************************************************

'*******Rollover switches
 Sub sw26_Hit:Controller.Switch(26) = 1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub
 Sub sw26_UnHit:Controller.Switch(26) = 0:End Sub

 Sub sw27_Hit:Controller.Switch(27) = 1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub
 Sub sw27_UnHit:Controller.Switch(27) = 0:End Sub

 Sub sw28_Hit:Controller.Switch(28) = 1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub
 Sub sw28_UnHit:Controller.Switch(28) = 0:End Sub

 Sub sw29_Hit:Controller.Switch(29) = 1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub
 Sub sw29_UnHit:Controller.Switch(29) = 0:End Sub

 Sub sw30_Hit:Controller.Switch(30) = 1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub
 Sub sw30_UnHit:Controller.Switch(30) = 0:End Sub

 Sub sw43_Hit:Controller.Switch(43) = 1:PlaySoundAtVol "rollover", ActiveBall, 1:plungerbulb.state=2:End Sub
 Sub sw43_UnHit:Controller.Switch(43) = 0:End Sub

 Sub sw44_Hit:Controller.Switch(44) = 1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub
 Sub sw44_UnHit:Controller.Switch(44) = 0:End Sub

 Sub sw45_Hit:Controller.Switch(45) = 1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub
 Sub sw45_UnHit:Controller.Switch(45) = 0:End Sub

 Sub sw59_Hit:Controller.Switch(59) = 1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub
 Sub sw59_UnHit:Controller.Switch(59) = 0:End Sub

 Sub sw60_Hit:Controller.Switch(60) = 1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub
 Sub sw60_UnHit:Controller.Switch(60) = 0:End Sub

 Sub sw65_Hit:Controller.Switch(65) = 1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub
 Sub sw65_UnHit:Controller.Switch(65) = 0:End Sub

 Sub sw66_Hit:Controller.Switch(66) = 1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub
 Sub sw66_UnHit:Controller.Switch(66) = 0:End Sub

'*******Ramp Switches
  Sub sw24_Hit()::vpmTimer.PulseSw 24:PlaySoundAtVol "Gate", ActiveBall, 1:End Sub

 Sub sw32_Hit():Controller.Switch(32) = 1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub
 Sub sw32_UnHit:Controller.Switch(32) = 0:End Sub

 Sub sw69_Hit:Controller.Switch(69) = 1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub
 Sub sw69_UnHit:Controller.Switch(69) = 0:End Sub

 Sub sw70_Hit:Controller.Switch(70) = 1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub
 Sub sw70_UnHit:Controller.Switch(70) = 0:End Sub

 Sub sw71_Hit:Controller.Switch(71) = 1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub
 Sub sw71_UnHit:Controller.Switch(71) = 0:End Sub

 Sub sw31_Hit:Controller.Switch(31) = 1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub
 Sub sw31_UnHit:Controller.Switch(31) = 0:End Sub

  Sub sw61a_Hit()::vpmTimer.PulseSw 61:End Sub
  Sub sw62a_Hit()::vpmTimer.PulseSw 62:End Sub

'*******Spinner
  Sub sw25_Spin():vpmTimer.PulseSw 25:PlaySoundAtVol "fx_spinner", sw25, VolFlip:End Sub

'*******StandUp Targets
  Sub sw21_Hit():vpmTimer.PulseSw 21:sw21p.transX=-5:me.TimerEnabled=1:End Sub
  Sub sw21_Timer():sw21p.transX=0:ME.TimerEnabled=0:End Sub

  Sub sw22_Hit():vpmTimer.PulseSw 22:sw22p.transX=-5:me.TimerEnabled=1:End Sub
  Sub sw22_Timer():sw22p.transX=0:ME.TimerEnabled=0:End Sub

  Sub sw23_Hit():vpmTimer.PulseSw 23:sw23p.transX=-5:me.TimerEnabled=1:End Sub
  Sub sw23_Timer():sw23p.transX=0:ME.TimerEnabled=0:End Sub

  Sub sw52_Hit():vpmTimer.PulseSw 52:sw52p.transX=-5:me.TimerEnabled=1:End Sub
  Sub sw52_Timer():sw52p.transX=0:ME.TimerEnabled=0:End Sub

  Sub sw53_Hit():vpmTimer.PulseSw 53:sw53p.transX=-5:me.TimerEnabled=1:End Sub
  Sub sw53_Timer():sw53p.transX=0:ME.TimerEnabled=0:End Sub


'***************************************************
'       JP's VP10 Fading Lamps & Flashers
'       Based on PD's Fading Light System
' SetLamp 0 is Off
' SetLamp 1 is On
' fading for non opacity objects is 4 steps
'***************************************************

Dim LampState(200), FadingLevel(200)
Dim FlashSpeedUp(200), FlashSpeedDown(200), FlashMin(200), FlashMax(200), FlashLevel(200)

Dim cnt
cnt =100

InitLamps()             ' turn off the lights and flashers and reset them to the default parameters
LampTimer.Interval = 10 'lamp fading speed
LampTimer.Enabled = 1

' Lamp & Flasher Timers

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

 Sub UpdateLamps
'GI
'If LampState(9) = 0 and LampState(10) = 0 and lampstate(11)= 0 and lampstate (25)=0 then GILeft.state=0
'If LampState(9) = 1 and LampState(10) = 1 and lampstate(11)= 1 and lampstate (25)=1 then GILeft.state=1
'If LampState(12) = 0 and LampState(14) = 0  then GIBLeft.state=0
'If LampState(12) = 1 and LampState(14) = 1  then GIBLeft.state=1
'If LampState(17) = 0 and LampState(18) = 0 then GIRight.state=0
'If LampState(17) = 1 and LampState(18) = 1 then GIRight.state=1



If LampState(21) = 0 and LampState(22) = 0 and lampstate(23)= 0 and lampstate(24)= 0 and _
LampState(9) = 0 and LampState(10) = 0 and lampstate(11)= 0 and lampstate (25)=0 and _
LampState(12) = 0 and LampState(14) = 0 and LampState(17) = 0 and LampState(18) = 0 and _
LampState(33) = 0 and LampState(34) = 0 and lampstate(35)= 0 and lampstate (36)=0 and _
LampState(57) = 0 and LampState(58) = 0 and LampState(59) =0 then

  If cnt = 0 then
  FlBgNeonLamp.visible=0
  FlBgNeonLamp1.visible=0
  FlBbBTitle.visible =0
  Flasher1.visible =0
  Flasher2.visible =0
  Flasher13.visible =0
  Flasher14.visible =0
  Flasher15.visible =0
  cnt = 100
  else
  cnt = cnt -1
  end if
End if

'If LampState(21) = 1 and LampState(22) = 1 and lampstate(23)= 1 and lampstate(24)= 1 then
If LampState(33) = 1 and LampState(34) = 1 and lampstate(35)= 1 and lampstate (36)=1 then
FlBgNeonLamp.visible=1
FlBgNeonLamp1.visible=1
FlBbBTitle.visible =1
Flasher1.visible =1
Flasher2.visible =1
Flasher13.visible =1
Flasher14.visible =1
Flasher15.visible =1
cnt =100
End if


'If LampState(33) = 0 and LampState(34) = 0 and lampstate(35)= 0 and lampstate (36)=0 then GITop.state=0
'If LampState(33) = 1 and LampState(34) = 1 and lampstate(35)= 1 and lampstate (36)=1 then GITop.state=1

  NFadeLm  9, l9      ' 4-Bank GI 1
  Flash 9, Flasher6
  NFadeL 10, l10      ' 4-Bank GI 2
  NFadeLm 11, l11     ' 4-Bank GI 3
  Flash 11, Flasher7
  NFadeL 12, l12      ' Left Sling GI 1


  NFadeL 14, l14      ' Left Flipper GI 1
  NFadeLm 17, l17     ' Upper Right Flipper GI 1
  Flash 17, Flasher9
  NFadeLm 18, l18     ' Eject Hole GI 1
  Flash 18, Flasher8
  NFadeLm 21, l21     ' Right Sling GI 1

  Flash 21,BBBLeftgirl 'FlBgNeonLamp
  NFadeLm 22, l22     ' Right Sling GI 2
  Flash 22, BBBRightgirl 'FlBgNeonLamp1
  NFadeLm 23, l23     ' Right Flipper GI 1
  NFadeL 23, l23a     ' Left Flipper GI 2
  NFadeL 24, l24      ' Right Flipper GI 2
  NFadeLm 25, l25     ' Tube GI 1
  Flash 25, Flasher5
  NFadeL 26, l26      ' Tube GI 2
  NFadeL 27, l27      ' Tube GI 3
  NFadeLm 28, l28     ' Tube GI 4
  Flash 28, Flasher4
  NFadeL 29, l29      ' Tube GI 5
  NFadeLm 33, l33     ' Hoot GI 1
  NFadeLm 33, l33a
  NFadeLm 34, l34     ' Hoot GI 2
  NFadeLm 33, l33b
  NFadeLm 35, l35     ' Hoot GI 3
  NFadeLm 33, l33c
  NFadeLm 36, l36     ' Hoot GI 4
  NFadeLm 33, l33d
  NFadeL 37, l37      ' Alien GI 1
  NFadeL 38, l38      ' Alien GI 2
  NFadeLm 39, l39     ' Alien GI 3
  Flash 39, Flasher11
  NFadeLm 40, l40     ' Captive GI 1
  Flash 40, Flasher12
  NFadeLm 128, l128   ' Upper Right Flipper GI 2
  Flash 128, Flasher10

'Inserts
  NFadeL 19, l19      ' Spaceship 1
  NFadeL 20, l20      ' Spaceship 2
  NFadeL 30, l30      ' Left Orbit Chase 1
  NFadeL 31, l31      ' Left Orbit Chase 2
  NFadeL 32, l32      ' Left Orbit Chase 3
  NFadeL 41, l41      ' Right Orbit Chase 1
  NFadeL 42, l42      ' Right Orbit Chase 2
  NFadeL 43, l43      ' Right Orbit Chase 3
  NFadeLm 49, l49     ' Rollover B
  Flash 49, l49r
  NFadeLm 50, l50     ' Rollover A
  Flash 50, l50r
  NFadeLm 51, l51     ' Rollover R
  Flash 51, l51r
  NFadeL 65, l65      ' Bonus 2X
  NFadeL 66, l66      ' Bonus 3X
  NFadeL 67, l67      ' Mode: Underground
  NFadeL 68, l68      ' Mode: Big Bang
  NFadeL 69, l69      ' Mode: Bar Brawl
  NFadeL 70, l70      ' Mode: Ball Busters
  NFadeL 71, l71      ' Mode: Looped
  NFadeL 72, l72      ' Shoot Again
  NFadeL 73, l73      ' Mode: Babescanner
  NFadeL 74, l74      ' Mode: Waitress
  NFadeL 75, l75      ' Shoot: Cosmic Dartz
  NFadeLm 76, l76     ' Special (Outlane R)
  Flash 76, l76r
  NFadeL 77, l77      ' Inlane Right
  NFadeL 78, l78      ' Bonus 5X
  NFadeL 79, l79      ' Bonus 4X
  NFadeL 80, l80      ' Mode: Tube Dancer
  NFadeL 81, l81      ' Shoot: Left Orbit
  NFadeL 82, l82      ' Shoot: Babescanner
  NFadeL 83, l83      ' 4-Bank Mars
  NFadeL 84, l84      ' 4-Bank Pythos
  NFadeL 85, l85      ' 4-Bank Venus
  NFadeL 86, l86      ' 4-Bank Mercury
  NFadeLm 87, l87     ' Free Shot (Outlane L)
  Flash 87, l87r
  NFadeL 88, l88      ' Inlane Left
  NFadeL 89, l89      ' Mode: Cosmic Dartz
  NFadeL 90, l90      ' Mode: Tour de Bar
  NFadeL 91, l91      ' Mode: Mosh
  NFadeL 92, l92      ' Mode: Happy Hour
  NFadeL 93, l93      ' Mode: Extra Ball
  NFadeL 94, l94      ' Mode: Get Lucky
  NFadeL 95, l95      ' Mode: Loonapalooza
  NFadeL 97, l97      ' Ramp Jackpot
  NFadeL 98, l98      ' Ramp Standup Left
  NFadeL 99, l99      ' Ramp Standup Right
  NFadeL 100, l100    ' Ramp Standup Side
  NFadeL 101, l101    ' Double Jackpot
  NFadeL 102, l102    ' Shoot: Tour de Bar
  NFadeL 103, l103    ' Shoot: Underground 1
  NFadeL 105, l105    ' Captive Left 4
  NFadeL 106, l106    ' Captive Left 3
  NFadeL 107, l107    ' Captive Left 2
  NFadeL 108, l108    ' Captive Left 1
  NFadeL 109, l109    ' Captive Right 4
  NFadeL 110, l110    ' Captive Right 3
  NFadeL 111, l111    ' Captive Right 2
  NFadeL 112, l112    ' Captive Right 1
  NFadeL 113, l113    ' 3-Bank Uranus
  NFadeL 114, l114    ' 3-Bank Neptune
  NFadeL 115, l115    ' 3-Bank Pluto
  NFadeL 116, l116    ' Shoot: Right Orbit
  NFadeL 117, l117    ' DJ Eyes
  NFadeL 118, l118    ' Shoot: Lunapalooza
  NFadeL 121, l121    ' Shoot: Underground 2
  NFadeLm 122, l122   ' Star Bumper Left
  NFadeL 122, l122a
  NFadeLm 123, l123   ' Star Bumper Medium
  NFadeL 123, l123a
  NFadeLm 124, l124   ' Star Bumper Right
  NFadeL 124, l124a
  NFadeL 125, l125    ' Dance Floor
  NFadeL 126, l126    ' Shoot: Extra Ball
  NFadeL 127, l127    ' Shoot: Big Bang

' Ramp Arrows
  NFadeLm 57, L57     ' Low
  FadeObj 57, l57p, "l_on", "l_b", "l_a", "l_off"
  NFadeLm 58, L58     ' Mid
  FadeObj 58, l58p, "l_on", "l_b", "l_a", "l_off"
  NFadeLm 59, L59     ' Top
  FadeObj 59, l59p, "l_on", "l_b", "l_a", "l_off"

' Neon Light
  FadeObjm 62, NeonLamp, "BlackLOn", "BlackLb", "BlackLa", "BlackL"
  Flashm 62, l62
  Flashm 62, l62a
  Flash 62, l62b

'Bulbs
  Flash 44, l44     ' Alien Lock Left Bulb
  Flash 45, l45     ' Alien Lock Right Bulb
  Flash 52, l52     ' Tube Sign X-Ball Bulb
  Flash 53, l53     ' Tube Sign 10 Million Bulb
  Flash 54, l54     ' Tube Sign Jackpot Bulb
  NFadeL 104, l104    ' Qualify Mode Bulb
  NFadeL 119, l119    ' Island: Lock Ready Bulb
  NFadeL 120, l120    ' Island: Mode Ready Bulb

'Flashers
  NFadeLm 161, f21      'BackBox Left Flasher
  Flash 161, f21a
  NFadeL 162, f22     'Tube Dancer Flasher
  NFadeL 163, f23     'Dance Floor Flasher
  NFadeLm 164, f24      'Eject Hole Flasher
  Flash 164, f24a
  NFadeLm 165, f25      'Alien Lock Flasher
  Flash 165, f25a
  NFadeL 166, F26     'Lower Lock Flasher

 End Sub

Sub InitLamps()
    Dim x
    For x = 0 to 200
        LampState(x) = 0         ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4       ' used to track the fading state
        FlashSpeedUp(x) = 0.5    ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.35 ' slower speed when turning off the flasher
        FlashMax(x) = 1          ' the maximum value when on, usually 1
        FlashMin(x) = 0          ' the minimum value when off, usually 0
        FlashLevel(x) = 0        ' the intensity of the flashers, usually from 0 to 1
    Next
End Sub

Sub SetLamp(nr, value)
    If value <> LampState(nr) Then
        LampState(nr) = abs(value)
        FadingLevel(nr) = abs(value) + 4
    End If
End Sub

' Lights: used for VP10 standard lights, the fading is handled by VP itself

Sub NFadeL(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.state = 0:FadingLevel(nr) = 0
        Case 5:object.state = 1:FadingLevel(nr) = 1
    End Select
End Sub

Sub NFadeLm(nr, object) ' used for multiple lights
    Select Case FadingLevel(nr)
        Case 4:object.state = 0
        Case 5:object.state = 1
    End Select
End Sub

'Lights, Ramps & Primitives used as 4 step fading lights
'a,b,c,d are the images used from on to off

Sub FadeObj(nr, object, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 4:object.image = b:FadingLevel(nr) = 6                   'fading to off...
        Case 5:object.image = a:FadingLevel(nr) = 1                   'ON
        Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1            'wait
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
End Sub

Sub Flashm(nr, object) 'multiple flashers, it just sets the flashlevel
    Object.IntensityScale = FlashLevel(nr)
End Sub

' Desktop Objects: Reels & texts (you may also use lights on the desktop)

' Reels

Sub FadeR(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.SetValue 1:FadingLevel(nr) = 6                   'fading to off...
        Case 5:object.SetValue 0:FadingLevel(nr) = 1                   'ON
        Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1             'wait
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

'******************
' RealTime Updates
'******************
Set MotorCallback = GetRef("GameTimer")

Sub GameTimer
    UpdateMechs
End Sub

Sub UpdateMechs
GateSWsx.RotZ= -GateL.currentangle
GateSWdx.RotZ= -GateR.currentangle
sw24p.rotx=-Spinner1.currentangle + 90
DivTubeP.rotz=DivTubef.currentangle
DivTube2P.rotz=DivTube2f.currentangle
LeftFlipperP.roty=LeftFlipper.currentangle
RightFlipperP.roty=RightFlipper.currentangle
RightUFlipperP.roty=RightFlipper1.currentangle
End Sub

' *********************************************************************
'                      Ramp Helpers
' *********************************************************************

sub TubeHelper_hit(): If ActiveBall.VelY < -1 Then activeball.vely= -1 : end if:end sub

Sub LRHelpers_Hit(idx)
  If ActiveBall.VelY < 0 AND ActiveBall.VelY  > -18 Then
    ActiveBall.VelY=ActiveBall.VelY/1.08
    ActiveBall.VelX=ActiveBall.VelX/1.08
  End If
  If ActiveBall.VelY > 0 Then
    ActiveBall.VelY=ActiveBall.VelY/1.1
    ActiveBall.VelX=ActiveBall.VelX/1.1
  End If
End Sub


' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Dim Xin,Yin,rAngle,Radit,wAngle,Pi
Pi = Round(4*Atn(1),6)          '3.1415926535897932384626433832795

 Sub GetAngle(Xin, Yin, wAngle)
  If Sgn(Xin) = 0 Then
    If Sgn(Yin) = 1 Then rAngle = 3 * Pi/2 Else rAngle = Pi/2
    If Sgn(Yin) = 0 Then rAngle = 0
  Else
    rAngle = atn(-Yin/Xin)
  End If
  If sgn(Xin) = -1 Then Radit = Pi Else Radit = 0
  If sgn(Xin) = 1 and sgn(Yin) = 1 Then Radit = 2 * Pi
  wAngle = round((Radit + rAngle),4)
 End Sub

Sub DampenXY (dtx,dfx,rx, dty, dfy, ry)  'dt is threshold speed, df is dampen factor 0 to 1 (higher more dampening), r is randomness
  Dim dfxRandomness
  Dim dfyRandomness
  rx=cint(rx)
  ry=cint(ry)
  dfxRandomness=INT(RND*(2*rx+1))
  dfyRandomness=INT(RND*(2*ry+1))
  dfx=dfx+(rx-dfxRandomness)*.01
  dfy=dfy+(ry-dfyRandomness)*.01
  If ABS(activeball.velx) > dtx Then activeball.velx=activeball.velx*(1-dfx*(ABS(activeball.velx)/100))
  If ABS(activeball.vely) > dty Then activeball.vely=activeball.vely*(1-dfy*(ABS(activeball.vely)/100))
End Sub


' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX and Rothbauerw
' PlaySound sound, 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
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

Sub PlaySoundAtVol(sound, tableobj, Volume)
  PlaySound sound, 1, Volume, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'Set all as per ball position & speed, but Vol Multiplier may be used eg; PlaySoundAtBallVol "sound",3

Sub PlaySoundAtBallVol(sound, VolMult)
  PlaySound sound, 0, Vol(ActiveBall) * VolMult, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub

'Set position as bumperX and Vol manually.

Sub PlaySoundAtBumperVol(sound, tableobj, Vol)
  PlaySound sound, 1, Vol, AudioPan(tableobj), 0,0,1, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBOTBallZ(sound, BOT)
    PlaySound sound, 0, ABS(BOT.velz)/17, Pan(BOT), 0, Pitch(BOT), 1, 0, AudioFade(BOT)
End Sub

'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

Function RndNum(min, max)
    RndNum = Int(Rnd() * (max-min + 1) ) + min ' Sets a random number between min and max
End Function

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = tableobj.y * 2 / table1.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = tableobj.x * 2 / table1.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = ball.x * 2 / table1.width-1
  If tmp > 0 Then
    Pan = Csng(tmp ^10)
  Else
    Pan = Csng(-((- tmp) ^10) )
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

Function BallVelZ(ball) 'Calculates the ball speed in the -Z
    BallVelZ = INT((ball.VelZ) * -1 )
End Function

Function VolZ(ball) ' Calculates the Volume of the sound based on the ball speed in the Z
    VolZ = Csng(BallVelZ(ball) ^2 / 200)*1.2
End Function

'*** Determines if a Points (px,py) is inside a 4 point polygon A-D in Clockwise/CCW order

Function InRect(px,py,ax,ay,bx,by,cx,cy,dx,dy)
  Dim AB, BC, CD, DA
  AB = (bx*py) - (by*px) - (ax*py) + (ay*px) + (ax*by) - (ay*bx)
  BC = (cx*py) - (cy*px) - (bx*py) + (by*px) + (bx*cy) - (by*cx)
  CD = (dx*py) - (dy*px) - (cx*py) + (cy*px) + (cx*dy) - (cy*dx)
  DA = (ax*py) - (ay*px) - (dx*py) + (dy*px) + (dx*ay) - (dy*ax)

  If (AB <= 0 AND BC <=0 AND CD <= 0 AND DA <= 0) Or (AB >= 0 AND BC >=0 AND CD >= 0 AND DA >= 0) Then
    InRect = True
  Else
    InRect = False
  End If
End Function

' *********************************************************************
'                      Extra Sounds
' *********************************************************************

Sub LWireStart_Hit():PlaysoundAtVol "WireRamp", ActiveBall, 1:End Sub

Sub RWireStart_Hit():PlaysoundAtVol "WireRamp", ActiveBall, 1:End Sub

Sub RWireStart1_Hit():PlaysoundAtVol "WireRamp", ActiveBall, 1:End Sub

Sub ShooterEnd_Hit():If activeball.z > 30  Then vpmTimer.AddTimer 150, "BallHitSound":plungerbulb.state=0:End If:End Sub

Sub BallHitSound(dummy):PlaySound "ballhit":End Sub

Sub LWireEnd_Hit()
     vpmTimer.AddTimer 150, "BallHitSound":
   StopSound "WireRamp"
 End Sub

Sub RWireEnd_Hit()
     vpmTimer.AddTimer 150, "BallHitSound":
   StopSound "WireRamp"
 End Sub

Sub Targets_Hit (idx)
  PlaySound SoundFX("target"), 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub TargetBankWalls_Hit (idx)
  PlaySound SoundFX("droptarget"), 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
  PlaySound "gate4", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Rubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 20 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End Sub

Sub Posts_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    RandomSoundRubber()
  End If
End Sub

Sub RandomSoundRubber()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

Sub LeftFlipper_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RightFlipper_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RightFlipper1_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 8 ' total number of balls
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub RollingTimer_Timer()
    Dim BOT, b
    BOT = GetBalls

  ' stop the sound of deleted balls
    For b = UBound(BOT) + 1 to tnob
        rolling(b) = False
        StopSound("fx_ballrolling" & b)
    Next

  ' exit the sub if no balls on the table
    If UBound(BOT) = -1 Then Exit Sub

  ' play the rolling sound for each ball
    For b = 0 to UBound(BOT)
      If BallVel(BOT(b) ) > 1 Then
        rolling(b) = True
        if BOT(b).z < 30 Then ' Ball on playfield
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), AudioPan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        Else ' Ball on raised ramp
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*.5, AudioPan(BOT(b) ), 0, Pitch(BOT(b) )+50000, 1, 0, AudioFade(BOT(b) )
        End If
      Else
        If rolling(b) = True Then
          StopSound("fx_ballrolling" & b)
          rolling(b) = False
        End If
      End If
    Next
End Sub

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / (VolDiv/VolCol), AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub


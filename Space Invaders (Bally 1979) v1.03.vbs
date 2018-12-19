Option Explicit
Randomize

' Thalamus 2018-11-01 : Improved directional sounds
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


On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the Controller.vbs file in order to run this table (installed with the VPX package in the scripts folder)"
On Error Goto 0

dim FastFlips ' cFastFlips lib
Const nDebug = 0
' ***************************************************************************
'                          DAY & NIGHT FUNCTIONS AND DATASETS
' ****************************************************************************

Dim nxx, DNS
DNS = Table1.NightDay
'Dim OPSValues: OPSValues=Array (100,50,20,10 ,5,4 ,3 ,2 ,1, 0,0)
'Dim DNSValues: DNSValues=Array (1.0,0.5,0.1,0.05,0.01,0.005,0.001,0.0005,0.0001, 0.00005,0.00000)
Dim OPSValues: OPSValues=Array (100,50,20,10 ,9,8 ,7 ,6 ,5, 4,3,2,1,0)
Dim DNSValues: DNSValues=Array (1.0,0.5,0.1,0.05,0.01,0.005,0.001,0.0005,0.0001, 0.00005,0.00001, 0.000005, 0.000001)
Dim SysDNSVal: SysDNSVal=Array (1.0,0.9,0.8,0.7,0.6,0.5,0.5,0.5,0.5, 0.5,0.5)
Dim DivValues: DivValues =Array (1,2,4,8,16,32,32,32,32, 32,32)
Dim DivValues2: DivValues2 =Array (1,1.5,2,2.5,3,3.5,4,4.5,5, 5.5,6)
Dim DivValues3: DivValues3 =Array (1,1.1,1.2,1.3,1.4,1.5,1.6,1.7,1.8,1.9,2.0,2.1)
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

Dim DivLevel: DivLevel = 95
Dim DNSVal: DNSVal = Round(DNS/10)
Dim DNShift: DNShift = 1

' PLAYFIELD GENERAL OPERATIONAL and LOCALALIZED GI ILLUMINATION
Dim aAllFlashers: aAllFlashers=Array(SIBGL1Pos0,SIBGL2Pos0,SIBGL3Pos0,SIBGL4Pos0,SIBGL1Pos1,SIBGL2Pos1,SIBGL3Pos1,SIBGL4Pos1,_
SIBGL1Pos2,SIBGL2Pos2,SIBGL3Pos2,SIBGL4Pos2,SIBGL1Pos3,SIBGL2Pos3,SIBGL3Pos3,SIBGL4Pos3,SIBGL1Pos4,SIBGL2Pos4,SIBGL3Pos4,SIBGL4Pos4,_
SIBGL1Pos1h,SIBGL2Pos1h,SIBGL3Pos1h,SIBGL4Pos1h,SIBGL1Pos2h,SIBGL2Pos2h,SIBGL3Pos2h,SIBGL4Pos2h,SIBGL1Pos3h,SIBGL2Pos3h,SIBGL3Pos3h,_
SIBGL4Pos3h,SIBGL1Pos4h,SIBGL2Pos4h,SIBGL3Pos4h,SIBGL4Pos4h,_
Flasher28,Flasher29,Flasher30,Flasher31,Flasher32,Flasher33,Flasher34,Flasher35,Flasher36,Flasher37,Flasher38,Flasher39)
Dim aGiLights: aGiLights=array(Light1,Light58,Light59,Light60,Light61,Light62,Light63,Light64,Light65,Light66,Light67,Light68,Light69,_
Light70,Light71,Light72,Light73,Light74,Light75,Light76,Light77,Light78,Light79,Light80,Light81,Light82,Light83,Light84,Light85,_
Light86,Light87,Light88,Light89,Light90,Light91,Light92,Light93,Light94,Light95,Light96,Light97,Light98,Light99,Light100,Light101,_
Light102,Light103,Light104,Light105,Light106,Light107,Light108)
Dim BloomLights: BloomLights=array()
Dim TargetDropGi: TargetDropGi = array()

' PLAYFIELD GLOBAL INTENSITY ILLUMINATION FLASHERS
Flasher3.opacity = OPSValues(DNSVal + DNShift) / DivValues(DNSVal)
Flasher3.intensityscale = DNSValues(DNSVal + DNShift) /DivValues(DNSVal)
Flasher4.opacity = OPSValues(DNSVal + DNShift) /DivValues(DNSVal)
Flasher4.intensityscale = DNSValues(DNSVal + DNShift) /DivValues(DNSVal)
Flasher5.opacity = OPSValues(DNSVal + DNShift) /DivValues(DNSVal)
Flasher5.intensityscale = DNSValues(DNSVal + DNShift) /DivValues(DNSVal)
'Flasher2.opacity = OPSValues(DNSVal + DNShift) /DivValues(DNSVal)
'Flasher2.intensityscale = DNSValues(DNSVal + DNShift) /DivValues(DNSVal)


Dim DivValues4: DivValues4 =Array (1,2,3,2.9,2.8,2.7,2.6,2.5,2.4,2.3,2.2,2.1,2.0)

'BACKBOX & BACKGLASS ILLUMINATION
SIBGDark.ModulateVsAdd = MVSAdd(DNSVal)
SIBGDark.Color = RGB(abs((RValUP(DNSVal)-23)/DivValues4(DNSVal)),abs((GValUP(DNSVal)-23)/DivValues4(DNSVal)),abs((BValUP(DNSVal)-23)/DivValues4(DNSVal)))
SIBGDark.Amount = FValUP(DNSVal) / DivValues3(DNSVal)
SIBGHigh.Color = RGB(RValDN(DNSVal),GValDN(DNSVal),BValDN(DNSVal))
SIBGHigh.Amount = FValDN(DNSVal)  / DivValues(DNSVal)
SIBGHigh1.Color = RGB(RValDN(DNSVal),GValDN(DNSVal),BValDN(DNSVal))
SIBGHigh1.Amount = FValDN(DNSVal) / DivValues(DNSVal)
SIBGHigh2.Color = RGB(RValDN(DNSVal),GValDN(DNSVal),BValDN(DNSVal))
SIBGHigh2.Amount = FValDN(DNSVal) / DivValues(DNSVal)

SIBGframe.ModulateVsAdd = MVSAdd(DNSVal+4) * 0.2
SIBGframe.Color = RGB(RValUP(DNSVal+4),GValUP(DNSVal+4),BValUP(DNSVal+4))
SIBGframe.Amount = FValUP(DNSVal+4) * DivValues(DNSVal)

SIBGHigh.intensityscale = 1.0
SIBGHigh1.intensityscale = 1.0
SIBGHigh2.intensityscale = 1.0
SIBGDark.intensityscale = 1.0
SIBGframe.intensityscale = 1.0


Table1.PlayfieldReflectionStrength = ReflDN(DNSVal + 4)

Dim aAltFlashers: aAltFlashers=Array(Flasher7,Flasher8,Flasher9,Flasher10,Flasher11,Flasher12,Flasher13,Flasher14,_
Flasher15,Flasher16,Flasher17,Flasher18, Flasher19,Flasher20,Flasher21,Flasher22,Flasher23,Flasher24,Flasher25,Flasher26)

For each nxx in aGiLights:nxx.intensity = nxx.intensity * SysDNSVal(DNSVal) /DivValues3(DNSVal) :Next
For each nxx in aAllFlashers:nxx.amount = nxx.amount / DivValues3(DNSVal):Next
For each nxx in aAllFlashers:nxx.opacity = nxx.opacity * OPSValues(DNSVal) / DivLevel:Next
'For each nxx in BloomLights:nxx.intensity = nxx.intensity *SysDNSVal(DNSVal)/DivValues3(DNSVal):Next
'For each nxx in TargetDropGi:nxx.intensity = nxx.intensity *SysDNSVal(DNSVal)/DivValues3(DNSVal):Next

For each nxx in aAltFlashers:nxx.amount = nxx.amount / DivValues3(DNSVal):Next

' Special considerations for Haunt Freaks wide body POV settings
	if table1.scalex < 1.2 then
	'Flasher5.opacity
	'Flasher5.amount
	If Table1.ShowDT = false and Table1.ShowFSS = false  then

	Flasher28.intensityscale = Flasher28.intensityscale  /4.0
	Flasher29.intensityscale = Flasher29.intensityscale  /4.0
	Flasher30.intensityscale = Flasher30.intensityscale  /4.0
	Flasher31.intensityscale = Flasher31.intensityscale  /4.0

	Light86.intensity =  Light86.intensity / 4.0
	Light87.intensity =  Light87.intensity / 4.0
	Light88.intensity =  Light88.intensity / 4.0
	Light89.intensity =  Light89.intensity / 4.0

	Light101.intensity =  Light101.intensity / 2.0
	Light102.intensity =  Light102.intensity / 2.0
	Light103.intensity =  Light103.intensity / 2.0
	Light104.intensity =  Light104.intensity / 2.0

	Light100.intensity =  Light100.intensity / 2.0
	Light58.intensity =  Light58.intensity / 2.0
	Light79.intensity =  Light79.intensity / 2.0
	Light99.intensity =  Light99.intensity / 3.0

	end if
	end If

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
' ***************************************************************************

Const cGameName = "spaceinv" 'Standard rom (don't use)
Const cGameName2= "spaceinb"  '7 digit rom

Const UseSolenoids = 1
Const UseLamps = 0
Const UseGI = 0
Const UseSync = 0
Const HandleMech = 0

LoadVPM "01560000", "Bally.VBS", 3.26
'LoadVPM"01000100","BALLY.VBS",1.2



SolCallback(7) = "bsTrough.SolOut"
SolCallback(6) = "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(8) = "dtBank.SolDropUp"
SolCallback(9) = "dtSingle.SolDropUp"
SolCallBack(19) = "FastFlips.TiltSol" ' cFastFlips lib
 '****FlipperSubs replaced by cfastflips
'SolCallback(sLRFlipper) = "SolRFlipper"
'SolCallback(sLLFlipper) = "SolLFlipper"


Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors), LeftFlipper, VolFlip
         PlaySoundAtVol "fx_Flipperup", LeftFlipper1, VolFlip
		LeftFlipper.RotateToEnd
		LeftFlipper1.RotateToEnd
		controller.switch(35) = 1
     Else
		if LeftFlipper.CurrentAngle = LeftFlipper.StartAngle then
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors), LeftFlipper, VolFlip
         PlaySoundAtVol "fx_Flipperdown", LeftFlipper1, VolFlip
		end if
		LeftFlipper.RotateToStart
		LeftFlipper1.RotateToStart
		controller.switch(35) = 0
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors), RightFlipper, VolFlip
         PlaySoundAtVol "fx_Flipperup", RightFlipper1, VolFlip
		RightFlipper.RotateToEnd
		RightFlipper1.RotateToEnd
		controller.switch(33) = 1
     Else
		if LeftFlipper.CurrentAngle = LeftFlipper.StartAngle then
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors), RightFlipper, VolFlip
         PlaySoundAtVol "fx_Flipperdown", RightFlipper1, VolFlip
		end if
		RightFlipper.RotateToStart
		RightFlipper1.RotateToStart
		controller.switch(33) = 0
     End If
End Sub

Sub UpdateFlipper_Timer
	Primitive17.ObjRotZ = LeftFlipper.CurrentAngle -120 +1
	Primitive15.ObjRotZ = LeftFlipper.CurrentAngle -120 +1

	Primitive18.ObjRotZ  = RightFlipper.CurrentAngle +120 +1
	Primitive16.ObjRotZ  = RightFlipper.CurrentAngle  +120 +1

	If nDebug >= 1 then
		If(RightFlipper.CurrentAngle <= -121.0) then ' -71.0
		'TextBox12.text = RightFlipper.CurrentAngle
		LPFDebugWallR1.isdropped=1
		LPFDebugWallR2.isdropped=1
		LPFDebugWallR3.isdropped=1
		LPFDebugWallR4.isdropped=1
		elseif(RightFlipper.CurrentAngle <= -100.0)  then '-80.0
		'TextBox12.text = RightFlipper.CurrentAngle
		LPFDebugWallR1.isdropped=0
		LPFDebugWallR2.isdropped=1
		LPFDebugWallR3.isdropped=1
		LPFDebugWallR4.isdropped=1
		elseif(RightFlipper.CurrentAngle <= -90.0)  then '-90.0
		'TextBox12.text = RightFlipper.CurrentAngle
		LPFDebugWallR1.isdropped=0
		LPFDebugWallR2.isdropped=0
		LPFDebugWallR3.isdropped=1
		LPFDebugWallR4.isdropped=1
		elseif(RightFlipper.CurrentAngle <= -80.0)  then '-100.0
		'TextBox12.text = RightFlipper.CurrentAngle
		LPFDebugWallR1.isdropped=0
		LPFDebugWallR2.isdropped=0
		LPFDebugWallR3.isdropped=0
		LPFDebugWallR4.isdropped=1
		elseif(RightFlipper.CurrentAngle <= -71.0)  then '-123.0
		'TextBox12.text = RightFlipper.CurrentAngle
		LPFDebugWallR1.isdropped=0
		LPFDebugWallR2.isdropped=0
		LPFDebugWallR3.isdropped=0
		LPFDebugWallR4.isdropped=0
		end if


		If(LeftFlipper.CurrentAngle >= 121.0) then ' -71.0
		'TextBox12.text = LeftFlipper.CurrentAngle
		LPFDebugWallL1.isdropped=1
		LPFDebugWallL2.isdropped=1
		LPFDebugWallL3.isdropped=1
		LPFDebugWallL4.isdropped=1
		elseif(LeftFlipper.CurrentAngle >= 100.0)  then '-80.0
		'TextBox12.text = LeftFlipper.CurrentAngle
		LPFDebugWallL1.isdropped=0
		LPFDebugWallL2.isdropped=1
		LPFDebugWallL3.isdropped=1
		LPFDebugWallL4.isdropped=1
		elseif(LeftFlipper.CurrentAngle >= 90.0)  then '-90.0
		'TextBox12.text = LeftFlipper.CurrentAngle
		LPFDebugWallL1.isdropped=0
		LPFDebugWallL2.isdropped=0
		LPFDebugWallL3.isdropped=1
		LPFDebugWallL4.isdropped=1
		elseif(LeftFlipper.CurrentAngle >= 80.0)  then '-100.0
		'TextBox12.text = LeftFlipper.CurrentAngle
		LPFDebugWallL1.isdropped=0
		LPFDebugWallL2.isdropped=0
		LPFDebugWallL3.isdropped=0
		LPFDebugWallL4.isdropped=1
		elseif(LeftFlipper.CurrentAngle >= 71.0)  then '-123.0
		'TextBox12.text =LeftFlipper.CurrentAngle
		LPFDebugWallL1.isdropped=0
		LPFDebugWallL2.isdropped=0
		LPFDebugWallL3.isdropped=0
		LPFDebugWallL4.isdropped=0
		end if

	end if

End Sub


Dim EnableBallControl
EnableBallControl = false 'Change to true to enable manual ball control (or press C in-game) via the arrow keys and B (boost movement) keys

'Standard Sounds
Const SSolenoidOn = "Solenoid"
Const SSolenoidOff = ""
Const SCoin = "Coin"

Dim bsTrough, dtBank, dtSingle, IMPowerSetting

Sub Table1_Init

	vpmInit Me
	On Error Resume Next

		With Controller
		.GameName = cGameName2
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
		.SplashInfoLine = "Space Invaders (Bally 1979)"&chr(13)&"Ported to VPX by Arconovum"
		.HandleMechanics=0
		.HandleKeyboard=0
		.ShowDMDOnly=1
		.ShowFrame=0
		.ShowTitle=0
        .hidden = 1
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With
     On Error Goto 0


'Nudging
       vpmNudge.TiltSwitch = 7
       vpmNudge.Sensitivity = 4
		vpmNudge.TiltObj = Array(BumperCenter, BumperLeft, BumperRight, LeftSlingshot, RightSlingShot)


'**Trough

       Set bsTrough = New cvpmBallStack
       With bsTrough
           .InitSw 0,8,0,0,0,0,0,0
  		  .InitKick BallRelease, 90, 4
			.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
			.InitEntrySnd SoundFX("Drain",DOFContactors), SoundFX("Solenoid",DOFContactors)
           .Balls = 1
       End With


 	CapKicker.CreateBall
 	CapKicker.Kick 0,0
	CapKicker.enabled = 0


 '**Targets
       set dtBank = new cvpmdroptarget
       With dtBank
           .InitDrop Array(sw1, sw2, sw3), Array(1, 2, 3)
           '.Initsnd "droptargetL", "resetdropL"
			.Initsnd SoundFX("droptargetL",DOFContactors),SoundFX("resetdropL",DOFContactors)
           '.CreateEvents "dtBank"
       End With



       set dtSingle = new cvpmdroptarget
       With dtSingle
           .InitDrop sw34, 34
           '.Initsnd "droptargetR", "resetdropR"
			.Initsnd SoundFX("droptargetR",DOFContactors),SoundFX("resetdropR",DOFContactors)
           '.CreateEvents "dtSingle"
       End With


 if 0 then
  '**Impulse Plunger
       Const IMPowerSetting = 55 ' Plunger Power
       Const IMTime = 0.6        ' Time in seconds for Full Plunge
       Set plungerIM = New cvpmImpulseP
       With plungerIM
           .InitImpulseP ImpPlunger, IMPowerSetting, IMTime
           .Random 0.3
           .InitExitSnd "plunger2", "plunger"
           .CreateEvents "plungerIM"
       End With
 end if
  '**Main Timer init
       PinMAMETimer.Interval = PinMAMEInterval
       PinMAMETimer.Enabled = 1

	if nDebug >= 1 Then
	DebugWall1.isdropped =0
	DebugWall2.isdropped =0

	LPFDebugPost3.isdropped=0

	LPFDebugWallR1.isdropped=0
	LPFDebugWallR2.isdropped=0
	LPFDebugWallR3.isdropped=0
	LPFDebugWallR4.isdropped=0

	LPFDebugWallL1.isdropped=0
	LPFDebugWallL2.isdropped=0
	LPFDebugWallL3.isdropped=0
	LPFDebugWallL4.isdropped=0
	Else
	DebugWall1.isdropped =1
	DebugWall2.isdropped =1
	LPFDebugPost3.isdropped=1

	LPFDebugWallR1.isdropped=1
	LPFDebugWallR2.isdropped=1
	LPFDebugWallR3.isdropped=1
	LPFDebugWallR4.isdropped=1

	LPFDebugWallL1.isdropped=1
	LPFDebugWallL2.isdropped=1
	LPFDebugWallL3.isdropped=1
	LPFDebugWallL4.isdropped=1

	end if

	'....cfastFlips....
Set FastFlips = new cFastFlips
with FastFlips
	.CallBackL = "SolLFlipper"	'Point these to flipper subs
	.CallBackR = "SolRFlipper"	'...
	'.CallBackUL = "SolULflipper"'...(upper flippers, if needed)
	'.CallBackUR = "SolURflipper"'...
	.TiltObjects = True 'Optional, if True calls vpmnudge.solgameon automatically. IF YOU GET A LINE 1 ERROR, DISABLE THIS! (or setup vpmNudge.TiltObj!)
	.InitDelay "FastFlips", 100			'Optional, if > 0 adds some compensation for solenoid jitter (occasional problem on Bram Stoker's Dracula)
	.DebugOn = False		'Debug, always-on flippers. Call FastFlips.DebugOn True or False in debugger to enable/disable.
end with

	setup_backglass()

  End Sub


Sub Table1_Paused: Controller.Pause = 1:End Sub
Sub Table1_UnPaused: Controller.Pause = 0:End Sub
Sub Table1_Exit(): Controller.Stop:End Sub


' ***************************************************************************
'                    BASIC FSS(DMD,SS,EM) SETUP CODE
' ****************************************************************************
Dim xtmp
if table1.showFSS = true Then
Ramp1773.visible = 1
Ramp2541.visible = 1
end if

if table1.showFSS = False Then
SIBGDark.visible = 0
sibgHigh.visible = 0
sibgHigh1.visible = 0
SIBGHigh2.visible = 0
SIBGFrame.visible = 0

	'For Each xtmp In BGArr
	'xtmp.visible = 0
	'Next
Flasher7.visible = 0
Flasher8.visible = 0
Flasher19.visible = 0
Flasher20.visible = 0
Flasher21.visible = 0
Flasher22.visible = 0
Flasher23.visible = 0
Flasher24.visible = 0
Flasher25.visible = 0
Flasher26.visible = 0
Flasher9.visible = 0
Flasher10.visible = 0
Flasher11.visible = 0
Flasher12.visible = 0
Flasher13.visible = 0
Flasher14.visible = 0
Flasher15.visible = 0
Flasher16.visible = 0
Flasher17.visible = 0
Flasher18.visible = 0

end if


const L1OFF = 0
const L2OFF = 10
const L3OFF = 20
const L4OFF = 30

const L0DEPTH = -10
const L1DEPTH = -20
const L2DEPTH = -30
const L3DEPTH = -40


const USEEM = 0
const USESS = 0
Dim xoff,yoff,zoff,xrot,zscale, xcen,ycen

sub setup_backglass()

xoff =620
yoff =0
zoff =670
xrot = -90


sibgdark.x = xoff
sibgdark.y = yoff
sibgdark.height = zoff
sibgdark.rotx = xrot

sibgHigh.x = xoff
sibgHigh.y = yoff +10
sibgHigh.height = zoff
sibgHigh.rotx = xrot

sibgHigh1.x = xoff
sibgHigh1.y = yoff +10
sibgHigh1.height = zoff
sibgHigh1.rotx = xrot


SIBGHigh2.x = xoff
SIBGHigh2.y = yoff +10
SIBGHigh2.height = zoff
SIBGHigh2.rotx = xrot


'bgGrill.x = xoff
'bgGrill.y = yoff
'bgGrill.height = zoff
'bgGrill.rotx = xrot

SIBGFrame.x = xoff +5
SIBGFrame.y = yoff
SIBGFrame.height = zoff - 5
SIBGFrame.rotx = xrot


'POS 0

SIBGL1Pos0.x = xoff
SIBGL1Pos0.y = yoff +L0DEPTH
SIBGL1Pos0.height = zoff
SIBGL1Pos0.rotx = xrot

SIBGL2Pos0.x = xoff
SIBGL2Pos0.y = yoff +L1DEPTH
SIBGL2Pos0.height = zoff +L2OFF
SIBGL2Pos0.rotx = xrot

SIBGL3Pos0.x = xoff
SIBGL3Pos0.y = yoff +L2DEPTH
SIBGL3Pos0.height = zoff +L3OFF
SIBGL3Pos0.rotx = xrot

SIBGL4Pos0.x = xoff
SIBGL4Pos0.y = yoff +L3DEPTH
SIBGL4Pos0.height = zoff +L4OFF
SIBGL4Pos0.rotx = xrot

'Pos 1
SIBGL1Pos1.x = xoff
SIBGL1Pos1.y = yoff +L0DEPTH
SIBGL1Pos1.height = zoff
SIBGL1Pos1.rotx = xrot


SIBGL2Pos1.x = xoff
SIBGL2Pos1.y = yoff +L1DEPTH
SIBGL2Pos1.height = zoff +L2OFF
SIBGL2Pos1.rotx = xrot

SIBGL3Pos1.x = xoff
SIBGL3Pos1.y = yoff +L2DEPTH
SIBGL3Pos1.height = zoff +L3OFF
SIBGL3Pos1.rotx = xrot

SIBGL4Pos1.x = xoff
SIBGL4Pos1.y = yoff +L3DEPTH
SIBGL4Pos1.height = zoff +L4OFF
SIBGL4Pos1.rotx = xrot

' Pos 1H

SIBGL1Pos1h.x = xoff
SIBGL1Pos1h.y = yoff +L0DEPTH
SIBGL1Pos1h.height = zoff
SIBGL1Pos1h.rotx = xrot

SIBGL2Pos1h.x = xoff
SIBGL2Pos1h.y = yoff +L1DEPTH
SIBGL2Pos1h.height = zoff +L2OFF
SIBGL2Pos1h.rotx = xrot

SIBGL3Pos1h.x = xoff
SIBGL3Pos1h.y = yoff +L2DEPTH
SIBGL3Pos1h.height = zoff +L3OFF
SIBGL3Pos1h.rotx = xrot

SIBGL4Pos1h.x = xoff
SIBGL4Pos1h.y = yoff +L3DEPTH
SIBGL4Pos1h.height = zoff +L4OFF
SIBGL4Pos1h.rotx = xrot



'Pos 2
SIBGL1Pos2.x = xoff
SIBGL1Pos2.y = yoff +L0DEPTH
SIBGL1Pos2.height = zoff
SIBGL1Pos2.rotx = xrot


SIBGL2Pos2.x = xoff
SIBGL2Pos2.y = yoff +L1DEPTH
SIBGL2Pos2.height = zoff +L2OFF
SIBGL2Pos2.rotx = xrot

SIBGL3Pos2.x = xoff
SIBGL3Pos2.y = yoff +L2DEPTH
SIBGL3Pos2.height = zoff +L3OFF
SIBGL3Pos2.rotx = xrot

SIBGL4Pos2.x = xoff
SIBGL4Pos2.y = yoff +L3DEPTH
SIBGL4Pos2.height = zoff +L4OFF
SIBGL4Pos2.rotx = xrot

' Pos 2H

SIBGL1Pos2h.x = xoff
SIBGL1Pos2h.y = yoff +L0DEPTH
SIBGL1Pos2h.height = zoff
SIBGL1Pos1h.rotx = xrot

SIBGL2Pos2h.x = xoff
SIBGL2Pos2h.y = yoff +L1DEPTH
SIBGL2Pos2h.height = zoff +L2OFF
SIBGL2Pos2h.rotx = xrot

SIBGL3Pos2h.x = xoff
SIBGL3Pos2h.y = yoff +L2DEPTH
SIBGL3Pos2h.height = zoff +L3OFF
SIBGL3Pos2h.rotx = xrot

SIBGL4Pos2h.x = xoff
SIBGL4Pos2h.y = yoff +L3DEPTH
SIBGL4Pos2h.height = zoff +L4OFF
SIBGL4Pos2h.rotx = xrot


'Pos 3
SIBGL1Pos3.x = xoff
SIBGL1Pos3.y = yoff +L0DEPTH
SIBGL1Pos3.height = zoff
SIBGL1Pos3.rotx = xrot


SIBGL2Pos3.x = xoff
SIBGL2Pos3.y = yoff +L1DEPTH
SIBGL2Pos3.height = zoff +L2OFF
SIBGL2Pos3.rotx = xrot

SIBGL3Pos3.x = xoff
SIBGL3Pos3.y = yoff +L2DEPTH
SIBGL3Pos3.height = zoff +L3OFF
SIBGL3Pos3.rotx = xrot

SIBGL4Pos3.x = xoff
SIBGL4Pos3.y = yoff +L3DEPTH
SIBGL4Pos3.height = zoff +L4OFF
SIBGL4Pos3.rotx = xrot

' Pos 3H

SIBGL1Pos3h.x = xoff
SIBGL1Pos3h.y = yoff +L0DEPTH
SIBGL1Pos3h.height = zoff
SIBGL1Pos3h.rotx = xrot

SIBGL2Pos3h.x = xoff
SIBGL2Pos3h.y = yoff +L1DEPTH
SIBGL2Pos3h.height = zoff +L2OFF
SIBGL2Pos3h.rotx = xrot

SIBGL3Pos3h.x = xoff
SIBGL3Pos3h.y = yoff +L2DEPTH
SIBGL3Pos3h.height = zoff +L3OFF
SIBGL3Pos3h.rotx = xrot

SIBGL4Pos3h.x = xoff
SIBGL4Pos3h.y = yoff +L3DEPTH
SIBGL4Pos3h.height = zoff +L4OFF
SIBGL4Pos3h.rotx = xrot

'Pos 4
SIBGL1Pos4.x = xoff
SIBGL1Pos4.y = yoff +L0DEPTH
SIBGL1Pos4.height = zoff
SIBGL1Pos4.rotx = xrot


SIBGL2Pos4.x = xoff
SIBGL2Pos4.y = yoff +L1DEPTH
SIBGL2Pos4.height = zoff +L2OFF
SIBGL2Pos4.rotx = xrot

SIBGL3Pos4.x = xoff
SIBGL3Pos4.y = yoff +L2DEPTH
SIBGL3Pos4.height = zoff +L3OFF
SIBGL3Pos4.rotx = xrot

SIBGL4Pos4.x = xoff
SIBGL4Pos4.y = yoff +L3DEPTH
SIBGL4Pos4.height = zoff +L4OFF
SIBGL4Pos4.rotx = xrot

' Pos 4H

SIBGL1Pos4h.x = xoff
SIBGL1Pos4h.y = yoff +L0DEPTH
SIBGL1Pos4h.height = zoff
SIBGL1Pos4h.rotx = xrot

SIBGL2Pos4h.x = xoff
SIBGL2Pos4h.y = yoff +L1DEPTH
SIBGL2Pos4h.height = zoff +L2OFF
SIBGL2Pos4h.rotx = xrot

SIBGL3Pos4h.x = xoff
SIBGL3Pos4h.y = yoff +L2DEPTH
SIBGL3Pos4h.height = zoff +L3OFF
SIBGL3Pos4h.rotx = xrot

SIBGL4Pos4h.x = xoff
SIBGL4Pos4h.y = yoff +L3DEPTH
SIBGL4Pos4h.height = zoff +L4OFF
SIBGL4Pos4h.rotx = xrot


center_graphix()
center_digits()

StartLampTimer

end sub


Dim BGArr
BGArr=Array (Flasher7,Flasher8,Flasher19,Flasher20,Flasher21,Flasher22,Flasher23,Flasher24,Flasher25,_
Flasher26, Flasher9,Flasher10,Flasher11,Flasher12,Flasher13,Flasher14,Flasher15,Flasher16,Flasher17,_
Flasher18)


Sub center_graphix()
Dim xx,yy,yfact,xfact,xobj
zscale = 0.00000001

xcen =(1305 /2) - (16 / 2)
ycen = (1090 /2 ) + (80/2)
yoff = yoff -2

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
	'xobj.visible =1 ' for testing
	Next
end sub


Sub center_digits()

Dim iii,xx,yy,yfact,xfact, objdx

zscale = 0.0000001
xcen =(1305 /2) - (16 / 2)
ycen = (1090 /2 ) + (80/2)

yfact =0 'y fudge factor (ycen was wrong so fix)
xfact =0


for iii =0 to 31
	For Each objdx In LEDDISP(iii)

	'if obj NOT n then

	xx =objdx.x

	objdx.x = (xoff -xcen) + xx +xfact
	yy = objdx.y ' get the yoffset before it is changed
	objdx.y =yoff

		If(yy < 0.) then
		yy = yy * -1
		end if

	objdx.height =( zoff - ycen) + yy - (yy * (zscale)) + yfact

	objdx.rotx = xrot

	objdx.visible = 0
	'end if
	Next
	'MsgBox "Count: " & iii

	Next
end sub


'Digital LED Display

Dim LEDDISP(32)

LEDDISP(0) = Array(LED1x0,LED1x1,LED1x2,LED1x3,LED1x4,LED1x5,LED1x6)', dummy, LED1x8)
LEDDISP(1) = Array(LED2x0,LED2x1,LED2x2,LED2x3,LED2x4,LED2x5,LED2x6)', dummy, LED2x8)
LEDDISP(2) = Array(LED3x0,LED3x1,LED3x2,LED3x3,LED3x4,LED3x5,LED3x6)', dummy, LED3x8)
LEDDISP(3) = Array(LED4x0,LED4x1,LED4x2,LED4x3,LED4x4,LED4x5,LED4x6)', dummy, LED4x8)
LEDDISP(4) = Array(LED5x0,LED5x1,LED5x2,LED5x3,LED5x4,LED5x5,LED5x6)', dummy, LED5x8)
LEDDISP(5) = Array(LED6x0,LED6x1,LED6x2,LED6x3,LED6x4,LED6x5,LED6x6)', dummy, LED6x8)
LEDDISP(6) = Array(LED7x0,LED7x1,LED7x2,LED7x3,LED7x4,LED7x5,LED7x6)', dummy, LED7x8)


LEDDISP(7) = Array(LED8x0,LED8x1,LED8x2,LED8x3,LED8x4,LED8x5,LED8x6)', dummy, LED8x8)
LEDDISP(8) = Array(LED9x0,LED9x1,LED9x2,LED9x3,LED9x4,LED9x5,LED9x6)', dummy, LED9x8)
LEDDISP(9) = Array(LED10x0,LED10x1,LED10x2,LED10x3,LED10x4,LED10x5,LED10x6)', dummy, LED10x8)
LEDDISP(10) = Array(LED11x0,LED11x1,LED11x2,LED11x3,LED11x4,LED11x5,LED11x6)', dummy, LED11x8)
LEDDISP(11) = Array(LED12x0,LED12x1,LED12x2,LED12x3,LED12x4,LED12x5,LED12x6)', dummy, LED12x8)
LEDDISP(12) = Array(LED13x0,LED13x1,LED13x2,LED13x3,LED13x4,LED13x5,LED13x6)', dummy, LED13x8)
LEDDISP(13) = Array(LED14x0,LED14x1,LED14x2,LED14x3,LED14x4,LED14x5,LED14x6)', dummy, LED14x8)

LEDDISP(14) = Array(LED1x000,LED1x001,LED1x002,LED1x003,LED1x004,LED1x005,LED1x006)', dummy, LED1x008)
LEDDISP(15) = Array(LED1x100,LED1x101,LED1x102,LED1x103,LED1x104,LED1x105,LED1x106)', dummy, LED1x108)
LEDDISP(16) = Array(LED1x200,LED1x201,LED1x202,LED1x203,LED1x204,LED1x205,LED1x206)', dummy, LED1x208)
LEDDISP(17) = Array(LED1x300,LED1x301,LED1x302,LED1x303,LED1x304,LED1x305,LED1x306)', dummy, LED1x308)
LEDDISP(18) = Array(LED1x400,LED1x401,LED1x402,LED1x403,LED1x404,LED1x405,LED1x406)', dummy, LED1x408)
LEDDISP(19) = Array(LED1x500,LED1x501,LED1x502,LED1x503,LED1x504,LED1x505,LED1x506)', dummy, LED1x508)
LEDDISP(20) = Array(LED1x600,LED1x601,LED1x602,LED1x603,LED1x604,LED1x605,LED1x606)', dummy, LED1x608)


LEDDISP(21) = Array(LED2x000,LED2x001,LED2x002,LED2x003,LED2x004,LED2x005,LED2x006)', dummy, LED2x008)
LEDDISP(22) = Array(LED2x100,LED2x101,LED2x102,LED2x103,LED2x104,LED2x105,LED2x106)', dummy, LED2x108)
LEDDISP(23) = Array(LED2x200,LED2x201,LED2x202,LED2x203,LED2x204,LED2x205,LED2x206)', dummy, LED2x208)
LEDDISP(24) = Array(LED2x300,LED2x301,LED2x302,LED2x303,LED2x304,LED2x305,LED2x306)', dummy, LED2x308)
LEDDISP(25) = Array(LED2x400,LED2x401,LED2x402,LED2x403,LED2x404,LED2x405,LED2x406)', dummy, LED2x408)
LEDDISP(26) = Array(LED2x500,LED2x501,LED2x502,LED2x503,LED2x504,LED2x505,LED2x506)', dummy, LED2x508)
LEDDISP(27) = Array(LED2x600,LED2x601,LED2x602,LED2x603,LED2x604,LED2x605,LED2x606)', dummy, LED2x608)

LEDDISP(28) = Array(LEDax300,LEDax301,LEDax302,LEDax303,LEDax304,LEDax305,LEDax306)', dummy, LEDax308)
LEDDISP(29) = Array(LEDbx400,LEDbx401,LEDbx402,LEDbx403,LEDbx404,LEDbx405,LEDbx406)', dummy, LEDbx408)

LEDDISP(30) = Array(LEDcx500,LEDcx501,LEDcx502,LEDcx503,LEDcx504,LEDcx505,LEDcx506)', dummy, LEDcx508)
LEDDISP(31) = Array(LEDdx600,LEDdx601,LEDdx602,LEDdx603,LEDdx604,LEDdx605,LEDdx606)', dummy, LEDdx608)


Sub DisplayTimer_Timer
if table1.showFSS = true then
	Dim ChgLED,ii,num,chg,stat,obj
	ChgLed = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
	If Not IsEmpty(ChgLED) Then
		'If DesktopMode = True Then
		For ii = 0 To UBound(chgLED)
			num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
			if (num < 32) then
				For Each obj In LEDDISP(num)
					If chg And 1 Then obj.visible = stat And 1
					chg = chg\2 : stat = stat\2
				Next
			else

			end if
		next
		'end if
	end if
end if
End Sub



Sub Table1_KeyDown(ByVal keycode)
	if keycode = LeftFlipperKey  then FastFlips.FlipL True :  FastFlips.FlipUL True
	if keycode = RightFlipperKey then FastFlips.FlipR True :  FastFlips.FlipUR True
	If keycode = LeftTiltKey Then Nudge 90, 3 : PlaySound SoundFX("fx_nudge",0)
	If keycode = RightTiltKey Then Nudge 270, 3 : PlaySound SoundFX("fx_nudge",0)
	If keycode = CenterTiltKey Then Nudge 0, 3 : PlaySound SoundFX("fx_nudge",0)
	If keycode = PlungerKey Then PlaySoundAtVol "fx_PlungerPull", Plunger, 1 :Plunger.Pullback
	'If vpmKeyDown(keycode) Then Exit Sub
	vpmKeyDown(keycode)
End Sub

Sub Table1_KeyUp(ByVal keycode)
	if keycode = LeftFlipperKey then FastFlips.FlipL False :  FastFlips.FlipUL False
	if keycode = RightFlipperKey then FastFlips.FlipR False :  FastFlips.FlipUR False
	If keycode = PlungerKey Then PLaySoundAtVol "fx_plunger", Plunger, 1 :Plunger.Fire
	'If vpmKeyUp(keycode) Then Exit Sub
	 vpmKeyUp(keycode)
End Sub




' BUMPERS

Sub BumperLeft_Hit(): vpmTimer.PulseSwitch (40), 0, "" : PlaySoundAtVol SoundFX("fx_bumper1",DOFContactors) , ActiveBall, VolBump: End Sub
Sub BumperCenter_Hit(): vpmTimer.PulseSwitch (38), 0, "" : PlaySoundAtVol SoundFX("fx_bumper2",DOFContactors) , ActiveBall, VolBump: End Sub
Sub BumperRight_Hit(): vpmTimer.PulseSwitch (39), 0, "" : PlaySoundAtVol SoundFX("fx_bumper3",DOFContactors) , ActiveBall, VolBump: End Sub



Sub Bumper1_Hit(): vpmTimer.PulseSwitch (5), 0, "" : PlaySoundAtVol SoundFX("rubber_hit_1",DOFContactors) , ActiveBall, VolBump: End Sub
Sub Bumper2_Hit(): vpmTimer.PulseSwitch (5), 0, "" : PlaySoundAtVol SoundFX("rubber_hit_2",DOFContactors) , ActiveBall, VolBump: End Sub
Sub Bumper3_Hit(): vpmTimer.PulseSwitch (5), 0, "" : PlaySoundAtVol SoundFX("rubber_hit_3",DOFContactors) , ActiveBall, VolBump: End Sub
Sub Bumper4_Hit(): vpmTimer.PulseSwitch (5), 0, "" : PlaySoundAtVol SoundFX("rubber_hit_1",DOFContactors) , ActiveBall, VolBump: End Sub


'Sub BumperCenter_Hit(): vpmTimer.PulseSwitch (38), 0, "" : PlaySound SoundFX("fx_bumper2",DOFContactors) : End Sub
'Sub BumperRight_Hit(): vpmTimer.PulseSwitch (39), 0, "" : PlaySound SoundFX("fx_bumper3",DOFContactors) : End Sub



'DROP TARGETS

Sub sw1_Dropped:dtBank.hit 1 :End Sub
Sub sw2_Dropped:dtBank.hit 2 :End Sub
Sub sw3_Dropped:dtBank.hit 3 :End Sub

Sub sw34_Dropped: dtSingle.hit 1 :End Sub


'TARGETS

Sub sw15_Hit:vpmTimer.PulseSw(15): playsoundAtVol"target", ActiveBall, VolTarg: End Sub
Sub sw17_Hit:vpmTimer.PulseSw(17): playsoundAtVol"target", ActiveBall, VolTarg: End Sub
Sub sw18_Hit:vpmTimer.PulseSw(18): playsoundAtVol"target", ActiveBall, VolTarg: End Sub
Sub sw19_Hit:vpmTimer.PulseSw(19): playsoundAtVol"target", ActiveBall, VolTarg: End Sub
Sub sw20_Hit:vpmTimer.PulseSw(20): playsoundAtVol"target", ActiveBall, VolTarg: End Sub
Sub sw21_Hit:vpmTimer.PulseSw(21): playsoundAtVol"target", ActiveBall, VolTarg: End Sub


' right trough
Sub Sw31_Hit:Controller.Switch(31) = 1: playsoundAtVol"rollover", ActiveBall, VolRol:End Sub
Sub Sw31_Unhit:Controller.Switch(31) = 0:End Sub

'
Sub Sw22_Hit:Controller.Switch(22) = 1: playsoundAtVol"rollover", ActiveBall, VolRol:End Sub
Sub Sw22_Unhit:Controller.Switch(22) = 0:End Sub
'
Sub Sw23_Hit:Controller.Switch(23) = 1: playsoundAtVol"rollover", ActiveBall, VolRol:End Sub
Sub Sw23_Unhit:Controller.Switch(23) = 0:End Sub
'
Sub Sw24_Hit:Controller.Switch(24) = 1: playsoundAtVol"rollover", ActiveBall, VolRol:End Sub
Sub Sw24_Unhit:Controller.Switch(24) = 0:End Sub



' left outer outlane drain
Sub Sw30_Hit:Controller.Switch(30) = 1: playsoundAtVol"rollover", ActiveBall, VolRol:End Sub
Sub Sw30_Unhit:Controller.Switch(30) = 0:End Sub
' left center inlane
Sub Sw29_Hit:Controller.Switch(29) = 1: playsoundAtVol"rollover", ActiveBall, VolRol:End Sub
Sub Sw29_Unhit:Controller.Switch(29) = 0:End Sub
' left inner inlane
Sub Sw28_Hit:Controller.Switch(28) = 1: playsoundAtVol"rollover", ActiveBall, VolRol:End Sub
Sub Sw28_Unhit:Controller.Switch(28) = 0:End Sub



' right outer outlane drain
Sub Sw25_Hit:Controller.Switch(25) = 1: playsoundAtVol"rollover", ActiveBall, VolRol:End Sub
Sub Sw25_Unhit:Controller.Switch(25) = 0:End Sub
' right center inlane
Sub Sw26_Hit:Controller.Switch(26) = 1: playsoundAtVol"rollover", ActiveBall, VolRol:End Sub
Sub Sw26_Unhit:Controller.Switch(26) = 0:End Sub
' right inner inlane
Sub Sw27_Hit:Controller.Switch(27) = 1: playsoundAtVol"rollover", ActiveBall, VolRol:End Sub
Sub Sw27_Unhit:Controller.Switch(27) = 0:End Sub

Dim nsw32: nsw32 =0
Dim nsw32h: nsw32h =0

Sub Sw32_Hit:

	if nsw32h=0 then
	nsw32=1
	elseif nsw32=1 then
	nsw32=0
	end if


	if nsw32h=1 then
	Controller.Switch(32) = 1
	Sw32.timerenabled = 1
	end If

playsoundAtVol"rollover", ActiveBall, 1
End Sub


Sub Sw32_Timer()

	if nsw32 = 5 then
	nsw32h =0
	nsw32 =0
	Me.timerenabled = 0
	Controller.Switch(32) = 0
	exit sub
	end If
nsw32 = nsw32 +1
End Sub


Sub Sw32a_Hit:

	if nsw32=0 then
	nsw32h=1
	elseif nsw32h=1 then
	nsw32h=0

	end if

	if nsw32=1 then
	Controller.Switch(32) = 1
	Sw32a.timerenabled = 1
	end If

playsoundAtVol"rollover", ActiveBall, 1
End Sub

Sub Sw32a_Timer()
	if nsw32h = 5 then
	nsw32h =0
	nsw32 =0
	Me.timerenabled = 0
	Controller.Switch(32) = 0
	exit sub
	end If
nsw32h = nsw32h +1
End Sub



' GATES

Sub Gate4_Hit()
gamestart = 1
end sub

' SPINNERS
Sub Sw4_Spin:vpmTimer.PulseSw(4) : playsoundAtVol"fx_spinner" , sw4, VolSpin: End Sub

'Drain
Sub Drain_Hit()
gamestart = 0
bsTrough.addball me
PlaySoundAtVol "drain", drain, 1
End Sub


'SLINGSHOTS

Sub LMidBankSlingShot_Slingshot
PlaySoundAtVol SoundFX("left_slingshot",DOFContactors), ActiveBall, 1
	vpmTimer.PulseSw (5)
end sub

Sub LUpSlingShot_Slingshot
PlaySoundAtVol SoundFX("left_slingshot",DOFContactors), ActiveBall, 1
	vpmTimer.PulseSw (5)
end sub


Sub LCenSlingShot_Slingshot
	PlaySoundAtVol SoundFX("left_slingshot",DOFContactors), ActiveBall, 1
	vpmTimer.PulseSw (5)
end sub


Sub RUpSlingShot_Slingshot
PlaySoundAtVol SoundFX("right_slingshot",DOFContactors), ActiveBall, 1
	vpmTimer.PulseSw (5)
end sub


Sub RCenSlingShot_Slingshot
	PlaySoundAtVol SoundFX("right_slingshot",DOFContactors), ActiveBall, 1
	vpmTimer.PulseSw (5)
end sub

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot

    PlaySoundAtVol SoundFX("right_slingshot",DOFContactors), sling1, 1
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
	vpmTimer.PulseSw(36)
	RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0:
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot

    PlaySoundAtVol SoundFX("left_slingshot",DOFContactors), sling2, 1
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
	vpmTimer.PulseSw(37)
	LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0:
    End Select
    LStep = LStep + 1
End Sub

'*****GI Lights On
dim xx

For each xx in GI:xx.State = 1: Next

' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX
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

'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

Function RndNum(min, max)
    RndNum = Int(Rnd() * (max-min + 1) ) + min ' Sets a random number between min and max
End Function

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / table1.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / table1.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
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

'*****************************************
'   rothbauerw's Manual Ball Control
'*****************************************

Dim BCup, BCdown, BCleft, BCright
Dim ControlBallInPlay, ControlActiveBall
Dim BCvel, BCyveloffset, BCboostmulti, BCboost

BCboost = 1				'Do Not Change - default setting
BCvel = 4				'Controls the speed of the ball movement
BCyveloffset = -0.01 	'Offsets the force of gravity to keep the ball from drifting vertically on the table, should be negative
BCboostmulti = 3		'Boost multiplier to ball veloctiy (toggled with the B key)

ControlBallInPlay = false

Sub StartBallControl_Hit()
	Set ControlActiveBall = ActiveBall
	ControlBallInPlay = true
End Sub

Sub StopBallControl_Hit()
	ControlBallInPlay = false
End Sub

Sub BallControlTimer_Timer()
	If EnableBallControl and ControlBallInPlay then
		If BCright = 1 Then
			ControlActiveBall.velx =  BCvel*BCboost
		ElseIf BCleft = 1 Then
			ControlActiveBall.velx = -BCvel*BCboost
		Else
			ControlActiveBall.velx = 0
		End If

		If BCup = 1 Then
			ControlActiveBall.vely = -BCvel*BCboost
		ElseIf BCdown = 1 Then
			ControlActiveBall.vely =  BCvel*BCboost
		Else
			ControlActiveBall.vely = bcyveloffset
		End If
	End If
End Sub


'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 5 ' total number of balls
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

'************************************
' What you need to add to your table
'************************************

' a timer called RollingTimer. With a fast interval, like 10
' one collision sound, in this script is called fx_collide
' as many sound files as max number of balls, with names ending with 0, 1, 2, 3, etc
' for ex. as used in this script: fx_ballrolling0, fx_ballrolling1, fx_ballrolling2, fx_ballrolling3, etc


'******************************************
' Explanation of the rolling sound routine
'******************************************

' sounds are played based on the ball speed and position

' the routine checks first for deleted balls and stops the rolling sound.

' The For loop goes through all the balls on the table and checks for the ball speed and
' if the ball is on the table (height lower than 30) then then it plays the sound
' otherwise the sound is stopped, like when the ball has stopped or is on a ramp or flying.

' The sound is played using the VOL, AUDIOPAN, AUDIOFADE and PITCH functions, so the volume and pitch of the sound
' will change according to the ball speed, and the AUDIOPAN & AUDIOFADE functions will change the stereo position
' according to the position of the ball on the table.


'**************************************
' Explanation of the collision routine
'**************************************

' The collision is built in VP.
' You only need to add a Sub OnBallBallCollision(ball1, ball2, velocity) and when two balls collide they
' will call this routine. What you add in the sub is up to you. As an example is a simple Playsound with volume and paning
' depending of the speed of the collision.


Sub Pins_Hit (idx)
	PlaySound "pinhit_low", 0, Vol(ActiveBall)*VolPi, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
	PlaySound "target", 0, Vol(ActiveBall)*VolTarg, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
	PlaySound "metalhit_thin", 0, Vol(ActiveBall)*VolMetal, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
	PlaySound "metalhit_medium", 0, Vol(ActiveBall)*VolMetal, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
	PlaySound "metalhit2", 0, Vol(ActiveBall)*VolMetal, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
	PlaySound "gate4", 0, Vol(ActiveBall)*VolGates, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinner_Spin
	PlaySoundAtVol "fx_spinner", Spinner, VolSpin
End Sub

Sub Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub Posts_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 16 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolPo, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 16 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub

Sub LeftFlipper_Collide(parm)
 	RandomSoundFlipper()
End Sub

Sub RightFlipper_Collide(parm)
 	RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub

'***************************************************
'  JP's Fading Lamps & Flashers version 9 for VP921
'   Based on PD's Fading Lights
' SetLamp 0 is Off
' SetLamp 1 is On
' FadingLevel(x) = fading state
' LampState(x) = light state
' Includes the flash element (needs own timer)
' Flashers can be used as lights too
'***************************************************

Dim LampState(200), FadingLevel(200)
Dim FlashState(200), FlashLevel(200)
Dim FlashSpeedUp, FlashSpeedDown

FlashInit()
FlasherTimer.Interval = 10 'flash fading speed
FlasherTimer.Enabled = 1

' Lamp & Flasher Timers

Sub StartLampTimer
	AllLampsOff()
	LampTimer.Interval = 30 'lamp fading speed
	LampTimer.Enabled = 1
End Sub

Sub LampTimer_Timer()
    Dim chgLamp, num, chg, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp) Then
        For ii = 0 To UBound(chgLamp)
            LampState(chgLamp(ii, 0) ) = chgLamp(ii, 1)
            FadingLevel(chgLamp(ii, 0) ) = chgLamp(ii, 1) + 4
			FlashState(chgLamp(ii, 0) ) = chgLamp(ii, 1)
        Next
    End If

    UpdateLamps
'	CheckDropShadows
'	UpdateMultipleLamps

End Sub

Dim gamestart: gamestart = 0
Dim Pos : Pos =0

Sub UpdateLamps

'NFadeL 8, L8

	if table1.ShowFSS = true then
	If gamestart = 0 Then
	SIBGL1Pos1.intensityscale = 1.0
	SIBGL2Pos1.intensityscale = 1.1
	SIBGL3Pos1.intensityscale = 1.2
	SIBGL4Pos1.intensityscale = 1.3

	SIBGL1Pos1h.intensityscale = 1.0
	SIBGL2Pos1h.intensityscale = 1.1
	SIBGL3Pos1h.intensityscale = 1.2
	SIBGL4Pos1h.intensityscale = 1.3


	SIBGL1Pos2.intensityscale = 1.0
	SIBGL2Pos2.intensityscale = 1.1
	SIBGL3Pos2.intensityscale = 1.2
	SIBGL4Pos2.intensityscale = 1.3

	SIBGL1Pos2h.intensityscale = 1.0
	SIBGL2Pos2h.intensityscale = 1.1
	SIBGL3Pos2h.intensityscale = 1.2
	SIBGL4Pos2h.intensityscale = 1.3

	SIBGL1Pos3.intensityscale = 1.0
	SIBGL2Pos3.intensityscale = 1.1
	SIBGL3Pos3.intensityscale = 1.2
	SIBGL4Pos3.intensityscale = 1.3

	SIBGL1Pos3h.intensityscale = 1.0
	SIBGL2Pos3h.intensityscale = 1.1
	SIBGL3Pos3h.intensityscale = 1.2
	SIBGL4Pos3h.intensityscale = 1.3

	SIBGL1Pos4.intensityscale = 1.0
	SIBGL2Pos4.intensityscale = 1.1
	SIBGL3Pos4.intensityscale = 1.2
	SIBGL4Pos4.intensityscale = 1.3

	SIBGL1Pos4h.intensityscale = 1.0
	SIBGL2Pos4h.intensityscale = 1.1
	SIBGL3Pos4h.intensityscale = 1.2
	SIBGL4Pos4h.intensityscale = 1.3


	SIBGL4Pos0.intensityscale = 1.0
	SIBGL3Pos0.intensityscale = 1.5
	SIBGL2Pos0.intensityscale = 2.0
	SIBGL1Pos0.intensityscale = 2.5

	SIBGL4Pos0.visible = 1
	SIBGL3Pos0.visible = 1
	SIBGL2Pos0.visible = 1
	SIBGL1Pos0.visible = 1


		if LampState(113) = 1 and Pos =0 then
		SIBGL1Pos1.visible = 1
		SIBGL2Pos1.visible = 1
		SIBGL3Pos1.visible = 1
		SIBGL4Pos1.visible = 1
		SIBGL1Pos1h.visible = 1
		SIBGL2Pos1h.visible = 1
		SIBGL3Pos1h.visible = 1
		SIBGL4Pos1h.visible = 1


	SIBGL1Pos4.visible = 0
	SIBGL2Pos4.visible = 0
	SIBGL3Pos4.visible = 0
	SIBGL4Pos4.visible = 0
	SIBGL1Pos4h.visible = 0
	SIBGL2Pos4h.visible = 0
	SIBGL3Pos4h.visible = 0
	SIBGL4Pos4h.visible = 0

		Pos = 1
		elseif LampState(84) = 1  and Pos =1 then
		SIBGL1Pos2.visible = 1
		SIBGL2Pos2.visible = 1
		SIBGL3Pos2.visible = 1
		SIBGL4Pos2.visible = 1
		SIBGL1Pos2h.visible = 1
		SIBGL2Pos2h.visible = 1
		SIBGL3Pos2h.visible = 1
		SIBGL4Pos2h.visible = 1

		SIBGL1Pos1.visible = 0
	SIBGL2Pos1.visible = 0
	SIBGL3Pos1.visible = 0
	SIBGL4Pos1.visible = 0
	SIBGL1Pos1h.visible = 0
	SIBGL2Pos1h.visible = 0
	SIBGL3Pos1h.visible = 0
	SIBGL4Pos1h.visible = 0
		Pos = 2
		elseif LampState(116) = 1 and Pos = 2  then
		SIBGL1Pos3.visible =1
		SIBGL2Pos3.visible = 1
		SIBGL3Pos3.visible = 1
		SIBGL4Pos3.visible = 1
		SIBGL1Pos3h.visible = 1
		SIBGL2Pos3h.visible = 1
		SIBGL3Pos3h.visible = 1
		SIBGL4Pos3h.visible = 1

	SIBGL1Pos2.visible = 0
	SIBGL2Pos2.visible = 0
	SIBGL3Pos2.visible = 0
	SIBGL4Pos2.visible = 0
	SIBGL1Pos2h.visible = 0
	SIBGL2Pos2h.visible = 0
	SIBGL3Pos2h.visible = 0
	SIBGL4Pos2h.visible = 0
		Pos = 3
		elseif LampState(81) = 1  and Pos = 3 then
		SIBGL1Pos4.visible = 1
		SIBGL2Pos4.visible = 1
		SIBGL3Pos4.visible = 1
		SIBGL4Pos4.visible = 1
		SIBGL1Pos4h.visible =1
		SIBGL2Pos4h.visible = 1
		SIBGL3Pos4h.visible = 1
		SIBGL4Pos4h.visible = 1

		SIBGL1Pos3.visible = 0
	SIBGL2Pos3.visible = 0
	SIBGL3Pos3.visible =0
	SIBGL4Pos3.visible = 0
	SIBGL1Pos3h.visible = 0
	SIBGL2Pos3h.visible = 0
	SIBGL3Pos3h.visible = 0
	SIBGL4Pos3h.visible = 0
		Pos =0
		end if

	Else
	SIBGL1Pos1.intensityscale = 1.0
	SIBGL2Pos1.intensityscale = 1.1
	SIBGL3Pos1.intensityscale = 1.2
	SIBGL4Pos1.intensityscale = 1.3

	SIBGL1Pos1h.intensityscale = 1.0
	SIBGL2Pos1h.intensityscale = 1.1
	SIBGL3Pos1h.intensityscale = 1.2
	SIBGL4Pos1h.intensityscale = 1.3


	SIBGL1Pos2.intensityscale = 1.0
	SIBGL2Pos2.intensityscale = 1.1
	SIBGL3Pos2.intensityscale = 1.2
	SIBGL4Pos2.intensityscale = 1.3

	SIBGL1Pos2h.intensityscale = 1.0
	SIBGL2Pos2h.intensityscale = 1.1
	SIBGL3Pos2h.intensityscale = 1.2
	SIBGL4Pos2h.intensityscale = 1.3

	SIBGL1Pos3.intensityscale = 1.0
	SIBGL2Pos3.intensityscale = 1.1
	SIBGL3Pos3.intensityscale = 1.2
	SIBGL4Pos3.intensityscale = 1.3

	SIBGL1Pos3h.intensityscale = 1.0
	SIBGL2Pos3h.intensityscale = 1.1
	SIBGL3Pos3h.intensityscale = 1.2
	SIBGL4Pos3h.intensityscale = 1.3

	SIBGL1Pos4.intensityscale = 1.0
	SIBGL2Pos4.intensityscale = 1.1
	SIBGL3Pos4.intensityscale = 1.2
	SIBGL4Pos4.intensityscale = 1.3

	SIBGL1Pos4h.intensityscale = 1.0
	SIBGL2Pos4h.intensityscale = 1.1
	SIBGL3Pos4h.intensityscale = 1.2
	SIBGL4Pos4h.intensityscale = 1.3


	SIBGL4Pos0.intensityscale = 1.0
	SIBGL3Pos0.intensityscale = 1.5
	SIBGL2Pos0.intensityscale = 2.0
	SIBGL1Pos0.intensityscale = 2.5



	SIBGL1Pos1.intensityscale = 1.0
	SIBGL1Pos1h.intensityscale = 1.0
	SIBGL1Pos2.intensityscale = 1.0
	SIBGL1Pos2h.intensityscale = 1.0
	SIBGL1Pos3.intensityscale = 1.0
	SIBGL2Pos3h.intensityscale = 1.0
	SIBGL1Pos4.intensityscale = 1.0
	SIBGL4Pos4h.intensityscale = 1.0

	SIBGL4Pos0.visible = 1
	SIBGL3Pos0.visible = 1
	SIBGL2Pos0.visible = 1
	SIBGL1Pos0.visible = 1


	SIBGL1Pos1.visible = LampState(113)
	SIBGL2Pos1.visible = LampState(113)
	SIBGL3Pos1.visible = LampState(113)
	SIBGL4Pos1.visible = LampState(113)
	SIBGL1Pos1h.visible = LampState(113)
	SIBGL2Pos1h.visible = LampState(113)
	SIBGL3Pos1h.visible = LampState(113)
	SIBGL4Pos1h.visible = LampState(113)

	SIBGL1Pos2.visible = LampState(84)
	SIBGL2Pos2.visible = LampState(84)
	SIBGL3Pos2.visible = LampState(84)
	SIBGL4Pos2.visible = LampState(84)
	SIBGL1Pos2h.visible = LampState(84)
	SIBGL2Pos2h.visible = LampState(84)
	SIBGL3Pos2h.visible = LampState(84)
	SIBGL4Pos2h.visible = LampState(84)

	SIBGL1Pos3.visible = LampState(116)
	SIBGL2Pos3.visible = LampState(116)
	SIBGL3Pos3.visible = LampState(116)
	SIBGL4Pos3.visible = LampState(116)
	SIBGL1Pos3h.visible = LampState(116)
	SIBGL2Pos3h.visible = LampState(116)
	SIBGL3Pos3h.visible = LampState(116)
	SIBGL4Pos3h.visible = LampState(116)

	SIBGL1Pos4.visible = LampState(81)
	SIBGL2Pos4.visible = LampState(81)
	SIBGL3Pos4.visible = LampState(81)
	SIBGL4Pos4.visible = LampState(81)
	SIBGL1Pos4h.visible = LampState(81)
	SIBGL2Pos4h.visible = LampState(81)
	SIBGL3Pos4h.visible = LampState(81)
	SIBGL4Pos4h.visible = LampState(81)

	end if
	end if

NFadeL 1, L1
NFadeL 2, L2
NFadeL 3, L3
NFadeLm 4, L4
NFadeL 4, L4a
NFadeL 5, L5
NFadeL 6, L6
NFadeL 7, L7
NFadeL 8, L8
NFadeL 9, L9
NFadeL 10, L10
'NFadeL 11, L11
NFadeL 12, L12
'NFadeL 13, L13
NFadeL 14, L14
NFadeL 15, L15
'NFadeL 16, L16
NFadeL 17, L17
NFadeL 18, L18
NFadeL 19, L19
NFadeLm 20, L20
NFadeL 20, L20a
NFadeL 21, L21
NFadeL 22, L22
NFadeL 23, L23
NFadeL 24, L24
NFadeL 25, L25
NFadeL 26, L26
'NFadeL 27, L27
NFadeL 28, L28
'NFadeL 29, L29
NFadeL 30, L30
'NFadeL 31, L31
'NFadeL 32, L32
NFadeL 33, L33
NFadeL 34, L34
NFadeL 35, L35
NFadeLm 36, L36
NFadeL 36, L36a
NFadeL 37, L37
NFadeL 38, L38
NFadeL 39, L39
NFadeL 40, L40
NFadeL 41, L41
NFadeL 42, L42
NFadeL 43, L43
NFadeL 44, L44
'NFadeL 45, L45
NFadeL 46, L46
NFadeL 47, L47
'NFadeL 48, L48
NFadeL 49, L49
NFadeL 50, L50
NFadeL 51, L51
NFadeL 52, L52
NFadeL 53, L53
NFadeL 54, L54
NFadeL 55, L55
NFadeL 56, L56
NFadeL 57, L57
NFadeL 58, L58
'NFadeL 59, L59
NFadeL 60, L60
'NFadeL 61, L61
NFadeL 62, L62
NFadeL 63, L63
'NFadeL 64, L64
'NFadeL 65, L65
'NFadeL 66, L66
'NFadeL 67, L67
'NFadeL 68, L68
'NFadeL 69, L69
'NFadeL 70, L70

End Sub

'Sindbad: You can use this instead of FadeLN
' call it this way: FadeLight lampnumber, light, Array
Sub FadeLight(nr, Light, group)
    Select Case FadingLevel(nr)
        Case 2:Light.image = group(3):FadingLevel(nr) = 0 'Off
        Case 3:Light.image = group(2):FadingLevel(nr) = 2 'fading...
        Case 4:Light.image = group(1):FadingLevel(nr) = 3 'fading...
        Case 5:Light.image = group(0):FadingLevel(nr) = 1 'ON
    End Select
End Sub


Sub NFadeL(nr, a)
    Select Case FadingLevel(nr)
        Case 4:a.state = 0:FadingLevel(nr) = 0
        Case 5:a.State = 1:FadingLevel(nr) = 1
    End Select
End Sub

Sub NFadeLm(nr, a)
    Select Case FadingLevel(nr)
        Case 4:a.state = 0
        Case 5:a.State = 1
    End Select
End Sub



'cyberpez FadeFlash can be used to swap images on ramps or flashers


Sub FadeFlash(nr, light, group)
	Select Case FadingLevel(nr)
		Case 2:light.imageA = group(3):FadingLevel(nr) = 0
		Case 3:light.imageA = group(2):FadingLevel(nr) = 2
		Case 4:light.imageA = group(1):FadingLevel(nr) = 3
		Case 5:light.imageA = group(0):FadingLevel(nr) = 1
	End Select
End Sub

Sub FadeFlashm(nr, light, group)
	Select Case FadingLevel(nr)
		Case 2:light.imageA = group(3):
		Case 3:light.imageA = group(2):
		Case 4:light.imageA = group(1):
		Case 5:light.imageA = group(0):
	End Select
End Sub


'trxture swap
dim itemw, itemp

Sub FadeMaterialW(nr, itemw, group)
    Select Case FadingLevel(nr)
        Case 4:itemw.TopMaterial = group(1):itemw.SideMaterial = group(1)
        Case 5:itemw.TopMaterial = group(0):itemw.SideMaterial = group(0)
    End Select
End Sub


Sub FadeMaterialP(nr, itemp, group)
    Select Case FadingLevel(nr)
        Case 4:itemp.Material = group(1)
        Case 5:itemp.Material = group(0)
    End Select
End Sub






' div lamp subs

Sub AllLampsOff()
    Dim x
    For x = 0 to 200
        LampState(x) = 0
        FadingLevel(x) = 4
    Next

UpdateLamps:UpdateLamps:Updatelamps
End Sub

Sub SetLamp(nr, value)
    If value <> LampState(nr) Then
        LampState(nr) = abs(value)
        FadingLevel(nr) = abs(value) + 4
    End If
End Sub




' flasher subs




Sub FlasherTimer_Timer()

'Flashm 111, Flasher10
'Flash 111, Flasher11

End Sub


Sub FlashInit
    Dim i
    For i = 0 to 200
        FlashState(i) = 0
        FlashLevel(i) = 0
    Next

    FlashSpeedUp = 30   ' fast speed when turning on the flasher
    FlashSpeedDown = 10 ' slow speed when turning off the flasher, gives a smooth fading
    ' you could also change the default images for each flasher or leave it as in the editor
    ' for example
    ' Flasher1.Image = "fr"
    AllFlashOff()
End Sub

Sub AllFlashOff
    Dim i
    For i = 0 to 200
        FlashState(i) = 0
    Next
End Sub

Sub SetFlash(nr, stat)
    FlashState(nr) = ABS(stat)
End Sub


' Flasher objects
' Uses own faster timer

Sub Flash(nr, object)
    Select Case FlashState(nr)
        Case 0 'off
            FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown
            If FlashLevel(nr) < 0 Then
                FlashLevel(nr) = 0
                FlashState(nr) = -1 'completely off
            End if
            Object.opacity = FlashLevel(nr)
        Case 1 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp
            If FlashLevel(nr) > 250 Then
                FlashLevel(nr) = 250
                FlashState(nr) = -2 'completely on
            End if
            Object.opacity = FlashLevel(nr)
    End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it doesn't change the flashstate
    Select Case FlashState(nr)
        Case 0         'off
            Object.opacity = FlashLevel(nr)
        Case 1         ' on
            Object.opacity = FlashLevel(nr)
    End Select
End Sub

Sub FlashVal(nr, object, value)
    Select Case FlashState(nr)
        Case 0 'off
            FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown
            If FlashLevel(nr) < 0 Then
                FlashLevel(nr) = 0
                FlashState(nr) = -1 'completely off
            End if
            Object.opacity = FlashLevel(nr)
        Case 1 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp
            If FlashLevel(nr) > value Then
                FlashLevel(nr) = value
                FlashState(nr) = -2 'completely on
            End if
            Object.opacity = FlashLevel(nr)
    End Select
End Sub

Sub FlashValm(nr, object, value) 'multiple flashers, it doesn't change the flashstate
    Select Case FlashState(nr)
        Case 0 'off
            Object.opacity = FlashLevel(nr)
        Case 1 ' on
            Object.opacity = FlashLevel(nr)
    End Select
End Sub



Sub FadeLn(nr, Light, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:Light.Offimage = d:FadingLevel(nr) = 0 'Off
        Case 3:Light.Offimage = c:FadingLevel(nr) = 2 'fading...
        Case 4:Light.Offimage = b:FadingLevel(nr) = 3 'fading...
        Case 5:Light.Offimage = a:FadingLevel(nr) = 1 'ON
    End Select
End Sub

Sub FadeLnm(nr, Light, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:Light.Offimage = d
        Case 3:Light.Offimage = c
        Case 4:Light.Offimage = b
        Case 5:Light.Offimage = a
    End Select
End Sub

Sub LMapn(nr, Light, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:Light.Onimage = d:FadingLevel(nr) = 0 'Off
        Case 3:Light.ONimage = c:FadingLevel(nr) = 2 'fading...
        Case 4:Light.ONimage = b:FadingLevel(nr) = 3 'fading...
        Case 5:Light.ONimage = a:FadingLevel(nr) = 1 'ON
    End Select
End Sub

Sub LMapnm(nr, Light, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:Light.ONimage = d
        Case 3:Light.ONimage = c
        Case 4:Light.ONimage = b
        Case 5:Light.ONimage = a
    End Select
End Sub
' Walls

Sub FadeW(nr, a, b, c)
    Select Case FadingLevel(nr)
        Case 2:c.IsDropped = 1:FadingLevel(nr) = 0                 'Off
        Case 3:b.IsDropped = 1:c.IsDropped = 0:FadingLevel(nr) = 2 'fading...
        Case 4:a.IsDropped = 1:b.IsDropped = 0:FadingLevel(nr) = 3 'fading...
        Case 5:a.IsDropped = 0:FadingLevel(nr) = 1                 'ON
    End Select
End Sub

Sub FadeWm(nr, a, b, c)
    Select Case FadingLevel(nr)
        Case 2:c.IsDropped = 1
        Case 3:b.IsDropped = 1:c.IsDropped = 0
        Case 4:a.IsDropped = 1:b.IsDropped = 0
        Case 5:a.IsDropped = 0
    End Select
End Sub

Sub NFadeW(nr, a)
    Select Case FadingLevel(nr)
        Case 4:a.IsDropped = 1:FadingLevel(nr) = 0
        Case 5:a.IsDropped = 0:FadingLevel(nr) = 1
    End Select
End Sub

Sub NFadeWm(nr, a)
    Select Case FadingLevel(nr)
        Case 4:a.IsDropped = 1
        Case 5:a.IsDropped = 0
    End Select
End Sub

Sub NFadeWi(nr, a)
    Select Case FadingLevel(nr)
        Case 4:a.IsDropped = 0:FadingLevel(nr) = 0
        Case 5:a.IsDropped = 1:FadingLevel(nr) = 1
    End Select
End Sub

Sub NFadeWim(nr, a)
    Select Case FadingLevel(nr)
        Case 4:a.IsDropped = 0
        Case 5:a.IsDropped = 1
    End Select
End Sub

'Lights

Sub FadeL(nr, a, b)
    Select Case FadingLevel(nr)
        Case 2:b.state = 0:FadingLevel(nr) = 0
        Case 3:b.state = 1:FadingLevel(nr) = 2
        Case 4:a.state = 0:FadingLevel(nr) = 3
        Case 5:a.state = 1:FadingLevel(nr) = 1
    End Select
End Sub

Sub FadeLm(nr, a, b)
    Select Case FadingLevel(nr)
        Case 2:b.state = 0
        Case 3:b.state = 1
        Case 4:a.state = 0
        Case 5:a.state = 1
    End Select
End Sub

'Sub NFadeL(nr, a)
'    Select Case FadingLevel(nr)
'        Case 4:a.state = 0:FadingLevel(nr) = 0
'        Case 5:a.State = 1:FadingLevel(nr) = 1
'    End Select
'End Sub
'
'Sub NFadeLm(nr, a)
'    Select Case FadingLevel(nr)
'        Case 4:a.state = 0
'        Case 5:a.State = 1
'    End Select
'End Sub

Sub LMap(nr, a, b, c) 'can be used with normal/olod style lights too
    Select Case FadingLevel(nr)
        Case 2:c.state = 0:FadingLevel(nr) = 0
        Case 3:b.state = 0:c.state = 1:FadingLevel(nr) = 2
        Case 4:a.state = 0:b.state = 1:FadingLevel(nr) = 3
        Case 5:b.state = 0:c.state = 0:a.state = 1:FadingLevel(nr) = 1
    End Select
End Sub

Sub LMapm(nr, a, b, c)
    Select Case FadingLevel(nr)
        Case 2:c.state = 0
        Case 3:b.state = 0:c.state = 1
        Case 4:a.state = 0:b.state = 1
        Case 5:b.state = 0:c.state = 0:a.state = 1
    End Select
End Sub

'Reels

Sub FadeR(nr, a)
    Select Case FadingLevel(nr)
        Case 2:a.SetValue 3:FadingLevel(nr) = 0
        Case 3:a.SetValue 2:FadingLevel(nr) = 2
        Case 4:a.SetValue 1:FadingLevel(nr) = 3
        Case 5:a.SetValue 1:FadingLevel(nr) = 1
    End Select
End Sub

Sub FadeRm(nr, a)
    Select Case FadingLevel(nr)
        Case 2:a.SetValue 3
        Case 3:a.SetValue 2
        Case 4:a.SetValue 1
        Case 5:a.SetValue 1
    End Select
End Sub

'Texts

Sub NFadeT(nr, a, b)
    Select Case FadingLevel(nr)
        Case 4:a.Text = "":FadingLevel(nr) = 0
        Case 5:a.Text = b:FadingLevel(nr) = 1
    End Select
End Sub

Sub NFadeTm(nr, a, b)
    Select Case FadingLevel(nr)
        Case 4:a.Text = ""
        Case 5:a.Text = b
    End Select
End Sub

' Flash a light, not controlled by the rom

Sub FlashL(nr, a, b)
    Select Case FadingLevel(nr)
        Case 1:b.state = 0:FadingLevel(nr) = 0
        Case 2:b.state = 1:FadingLevel(nr) = 1
        Case 3:a.state = 0:FadingLevel(nr) = 2
        Case 4:a.state = 1:FadingLevel(nr) = 3
        Case 5:b.state = 1:FadingLevel(nr) = 4
    End Select
End Sub

' Light acting as a flash. C is the light number to be restored

Sub MFadeL(nr, a, b, c)
    Select Case FadingLevel(nr)
        Case 2:b.state = 0:FadingLevel(nr) = 0:SetLamp c, FadingLevel(c)
        Case 3:b.state = 1:FadingLevel(nr) = 2
        Case 4:a.state = 0:FadingLevel(nr) = 3
        Case 5:a.state = 1:FadingLevel(nr) = 1
    End Select
End Sub

Sub MFadeLm(nr, a, b, c)
    Select Case FadingLevel(nr)
        Case 2:b.state = 0:SetLamp c, FadingLevel(c)
        Case 3:b.state = 1
        Case 4:a.state = 0
        Case 5:a.state = 1
    End Select
End Sub

'Alpha Ramps used as fading lights
'ramp is the name of the ramp
'a,b,c,d are the images used for on...off
'r is the refresh light

Sub FadeAR(nr, ramp, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:ramp.image = d:FadingLevel(nr) = 0 'Off
        Case 3:ramp.image = c:FadingLevel(nr) = 2 'fading...
        Case 4:ramp.image = b:FadingLevel(nr) = 3 'fading...
        Case 5:ramp.image = a:FadingLevel(nr) = 1 'ON
    End Select
End Sub

Sub FadeARm(nr, ramp, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:ramp.image = d
        Case 3:ramp.image = c
        Case 4:ramp.image = b
        Case 5:ramp.image = a
    End Select
End Sub

Sub FlashFO(nr, ramp, a, b, c)                                   'used for reflections when there is no off ramp
    Select Case FadingLevel(nr)
        Case 2:ramp.IsVisible = 0:FadingLevel(nr) = 0                'Off
        Case 3:ramp.image = c:FadingLevel(nr) = 2                'fading...
        Case 4:ramp.image = b:FadingLevel(nr) = 3                'fading...
        Case 5:ramp.image = a:ramp.IsVisible = 1:FadingLevel(nr) = 1 'ON
    End Select
End Sub

Sub FlashAR(nr, ramp, a, b, c)                                   'used for reflections when there is no off ramp
    Select Case FadingLevel(nr)
        Case 2:ramp.alpha = 0:FadingLevel(nr) = 0                'Off
        Case 3:ramp.image = c:FadingLevel(nr) = 2                'fading...
        Case 4:ramp.image = b:FadingLevel(nr) = 3                'fading...
        Case 5:ramp.image = a:ramp.alpha = 1:FadingLevel(nr) = 1 'ON
    End Select
End Sub

Sub FlashARm(nr, ramp, a, b, c)
    Select Case FadingLevel(nr)
        Case 2:ramp.alpha = 0
        Case 3:ramp.image = c
        Case 4:ramp.image = b
        Case 5:ramp.image = a:ramp.alpha = 1
    End Select
End Sub

Sub NFadeAR(nr, ramp, a, b)
    Select Case FadingLevel(nr)
        Case 4:ramp.image = b:FadingLevel(nr) = 0 'off
        Case 5:ramp.image = a:FadingLevel(nr) = 1 'on
    End Select
End Sub

Sub NFadeARm(nr, ramp, a, b)
    Select Case FadingLevel(nr)
        Case 4:ramp.image = b
        Case 5:ramp.image = a
    End Select
End Sub

Sub MNFadeAR(nr, ramp, a, b, c)
    Select Case FadingLevel(nr)
        Case 4:ramp.image = b:FadingLevel(nr) = 0:SetLamp c, FadingLevel(c) 'off
        Case 5:ramp.image = a:FadingLevel(nr) = 1                           'on
    End Select
End Sub

Sub MNFadeARm(nr, ramp, a, b, c)
    Select Case FadingLevel(nr)
        Case 4:ramp.image = b:SetLamp c, FadingLevel(c) 'off
        Case 5:ramp.image = a                           'on
    End Select
End Sub

' Flashers using PRIMITIVES
' pri is the name of the primitive
' a,b,c,d are the images used for on...off

Sub FadePri(nr, pri, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:pri.image = d:FadingLevel(nr) = 0 'Off
        Case 3:pri.image = c:FadingLevel(nr) = 2 'fading...
        Case 4:pri.image = b:FadingLevel(nr) = 3 'fading...
        Case 5:pri.image = a:FadingLevel(nr) = 1 'ON
    End Select
End Sub

Sub FadePri3m(nr, pri, group)
    Select Case FadingLevel(nr)
        Case 3:pri.image = group(2)
        Case 4:pri.image = group(1)
        Case 5:pri.image = group(0)
    End Select
End Sub

Sub FadePri3(nr, pri, group)
    Select Case FadingLevel(nr)
        Case 3:pri.image = group(2):FadingLevel(nr) = 0 'Off
        Case 4:pri.image = group(1):FadingLevel(nr) = 3 'fading...
        Case 5:pri.image = group(0):FadingLevel(nr) = 1 'ON
    End Select
End Sub



Sub FadePri2(nr, pri, group)
    Select Case FadingLevel(nr)
        Case 4:pri.image = group(1) 'off
        Case 5:pri.image = group(0) 'ON
    End Select
End Sub

Sub FadePri2m(nr, pri, group)
    Select Case FadingLevel(nr)
        Case 4:pri.image = group(1):FadingLevel(nr) = 0 'off
        Case 5:pri.image = group(0):FadingLevel(nr) = 1 'ON
    End Select
End Sub




Sub FadePri4m(nr, pri, group)
    Select Case FadingLevel(nr)
		Case 2:pri.image = group(3) 'Off
        Case 3:pri.image = group(2) 'Fading...
        Case 4:pri.image = group(1) 'Fading...
        Case 5:pri.image = group(0) 'ON
    End Select
End Sub

Sub FadePri4(nr, pri, group)
    Select Case FadingLevel(nr)
        Case 2:pri.image = group(3):FadingLevel(nr) = 0 'Off
        Case 3:pri.image = group(2):FadingLevel(nr) = 2 'Fading...
        Case 4:pri.image = group(1):FadingLevel(nr) = 3 'Fading...
        Case 5:pri.image = group(0):FadingLevel(nr) = 1 'ON
    End Select
End Sub












Sub FadePriC(nr, pri, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:For each xx in pri:xx.image = d:Next:FadingLevel(nr) = 0 'Off
        Case 3:For each xx in pri:xx.image = c:Next:FadingLevel(nr) = 2 'fading...
        Case 4:For each xx in pri:xx.image = b:Next:FadingLevel(nr) = 3 'fading...
        Case 5:For each xx in pri:xx.image = a:Next:FadingLevel(nr) = 1 'ON
    End Select
End Sub

Sub FadePrih(nr, pri, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:pri.image = d:SetFlash nr, 0:FadingLevel(nr) = 0 'Off
        Case 3:pri.image = c:FadingLevel(nr) = 2 'fading...
        Case 4:pri.image = b:FadingLevel(nr) = 3 'fading...
        Case 5:pri.image = a:SetFlash nr, 1:FadingLevel(nr) = 1 'ON
    End Select
End Sub

Sub FadePrim(nr, pri, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:pri.image = d
        Case 3:pri.image = c
        Case 4:pri.image = b
        Case 5:pri.image = a
    End Select
End Sub

Sub NFadePri(nr, pri, a, b)
    Select Case FadingLevel(nr)
        Case 4:pri.image = b:FadingLevel(nr) = 0 'off
        Case 5:pri.image = a:FadingLevel(nr) = 1 'on
    End Select
End Sub

Sub NFadePrim(nr, pri, a, b)
    Select Case FadingLevel(nr)
        Case 4:pri.image = b
        Case 5:pri.image = a
    End Select
End Sub

'Fade a collection of lights

Sub FadeLCo(nr, a, b) 'fading collection of lights
    Dim obj
    Select Case FadingLevel(nr)
        Case 2:vpmSolToggleObj b, Nothing, 0, 0:FadingLevel(nr) = 0
        Case 3:vpmSolToggleObj b, Nothing, 0, 1:FadingLevel(nr) = 2
        Case 4:vpmSolToggleObj a, Nothing, 0, 0:FadingLevel(nr) = 3
        Case 5:vpmSolToggleObj a, Nothing, 0, 1:FadingLevel(nr) = 1
    End Select
End Sub

'only for this table

' Alternate two lights made with walls

Sub AlternateW(nr, a, b)
    Select Case LampState(nr)
        Case 4:SetLamp a, 0:SetLamp b, 0:LampState(nr) = 0   'both off
        Case 5:SetLamp a, 1:SetLamp b, 0:LampState(nr) = 6   'On - Off
        Case 6:LampState(nr) = 7                             'wait
        Case 7:LampState(nr) = 8                             'wait
        Case 8:LampState(nr) = 9                             'wait
        Case 9:LampState(nr) = 10                            'wait
        Case 10:SetLamp a, 0:SetLamp b, 1:LampState(nr) = 11 'Off - On
        Case 11:LampState(nr) = 12                           'wait
        Case 12:LampState(nr) = 13                           'wait
        Case 13:LampState(nr) = 14                           'wait
        Case 14:LampState(nr) = 15                           'wait
        Case 15:LampState(nr) = 5
    End Select
End Sub


'********Options********
Sub editDips
Dim vpmDips : Set vpmDips = New cvpmDips
  	With vpmDips
	.AddForm 700,400,"DIP switch setup"

	.AddChk 7,2,180,Array("Match feature",&H08000000)'dip 28
  	.AddChk 7,21,180,Array("Credits displayed",&H04000000)'dip 27
	.AddChk	205,2,180,Array("Flipper button blast noise",&H00000080)'dip 8
  	.AddChk 205,21,180,Array("Background thumping sound",&H00004000)'dip 15
  	.AddFrame 2,40,190,"Maximum credits",&H03000000,Array("10 credits",0,"15 credits",&H01000000,"25 credits",&H02000000,"free play (40 credits)",&H03000000)'dip 25&26
  	.AddFrame 2,116,190,"Sound features",&H30000000,Array("chime effects",0,"noises and no background",&H10000000,"noise effects",&H20000000,"noises and background",&H30000000)'dip 29&30
  	.AddFrame 2,192,190,"High score feature award",&H00000060,Array("no award",&H00000020,"extra ball",&H00000040,"replay",&H00000060)'dip 6&7
  	.AddFrame 2,253,190,"High score to date award",&H00200000,Array("no award",0,"3 credits",&H00200000)'dip 22
  	.AddFrame 2,299,190,"Score version",&H00100000,Array("6 digit scoring",0,"7 digit scoring",&H00100000)'dip 21
  	.AddFrame 205,40,190,"Balls per game",&H40000000,Array("3 balls",0,"5 balls",&H40000000)'dip 31
  	.AddFrame 205,92,190,"Red invader lite",&H00400000,Array("will come back on",0,"will not come back on",&H00400000)'dip 23
  	.AddFrame 205,144,190,"Bumpers score",&H00002000,Array("100 points",0,"1000 points",&H00002000)'dip 14
  	.AddFrame 205,192,190,"Left and right extra ball arrow",&H80000000,Array("1 lite comes on then alternates",0,"both lites come on",&H80000000)'dip 32
  	.AddFrame 205,253,190,"Small flipper feed lanes lite",32768,Array("separate",0,"tied together",32768)'dip 16
  	.AddFrame 205,299,190,"Clone chamber value",&H00800000,Array("does not step up value",0,"steps up value",&H00800000)'dip 24
	.AddLabel 50,350,300,20,"After hitting OK, press F3 to reset game with new settings."

	'.AddChk 7,2,180,Array("Match feature",&H08000000)'dip 28
  	'.AddChk 7,21,180,Array("Credits displayed",&H04000000)'dip 27
  	'.AddChk	205,2,180,Array("Flipper button blast noise",&H00000080)'dip 8
  	'.AddChk 205,21,180,Array("Background thumping sound",&H00004000)'dip 15
  	'.AddFrame 2,40,190,"Maximum credits",&H03000000,Array("10 credits",0,"15 credits",&H01000000,"25 credits",&H02000000,"40 credits",&H03000000)'dip 25&26
  	'.AddFrame 2,116,190,"Sound features",&H30000000,Array("chime effects",0,"noises and no background",&H10000000,"noise effects",&H20000000,"noises and background",&H30000000)'dip 29&30
  	'.AddFrame 2,194,190,"High score feature award",&H00000060,Array("no award",&H00000020,"extra ball",&H00000040,"replay",&H00000060)'dip 6&7
  	'.AddFrame 2,258,190,"High score to date award",&H00300000,Array("no award",0,"1 credit",&H00100000,"2 credits",&H00200000,"3 credits",&H00300000)'dip 21&22
  	'.AddFrame 205,40,190,"Balls per game",&H40000000,Array("3 balls",0,"5 balls",&H40000000)'dip 31
  	'.AddFrame 205,90,190,"Red invader lite",&H00400000,Array("will come back on",0,"will not come back on",&H00400000)'dip 23
  	'.AddFrame 205,139,190,"Bumpers score",&H00002000,Array("100 points",0,"1000 points",&H00002000)'dip 14
  	'.AddFrame 205,188,190,"Left and right extra ball arrow",&H80000000,Array("1 lite comes on then alternates",0,"both lites come on",&H80000000)'dip 32
  	'.AddFrame 205,237,190,"Small flipper feed lanes lite",32768,Array("separate",0,"tied together",32768)'dip 16
  	'.AddFrame 205,286,190,"Clone chamber value",&H00800000,Array("does not step up value",0,"steps up value",&H00800000)'dip 24

	.ViewDips

	End With
 End Sub
 Set vpmShowDips = GetRef("editDips")

' ***************************************************************************
'                   cFastFlips by nFozzy
' ***************************************************************************

'cFastFlips by nFozzy
'Bypasses pinmame callback for faster and more responsive flippers
'Version 1.1 beta2 (More proper behaviour, extra safety against script errors)

'Flipper / game-on Solenoid # reference
'Atari: Sol16
'Astro:  ?
'Bally Early 80's: Sol19
'Bally late 80's (Blackwater 100, etc): Sol19
'Game Plan: Sol16
'Gottlieb System 1: Sol17
'Gottlieb System 80: No dedicated flipper solenoid? GI circuit Sol10?
'Gottlieb System 3: Sol32
'Playmatic: Sol8
'Spinball: Sol25
'Stern (80's): Sol19
'Taito: ?
'Williams System 3, 4, 6: Sol23
'Williams System 7: Sol25
'Williams System 9: Sol23
'Williams System 11: Sol23
'Bally / Williams WPC 90', 92', WPC Security: Sol31
'Data East (and Sega pre-whitestar): Sol23
'Zaccaria: ???

'********************Setup*******************:

'....somewhere outside of any subs....
'dim FastFlips

'....table init....
'Set FastFlips = new cFastFlips
'with FastFlips
'	.CallBackL = "SolLflipper"	'Point these to flipper subs
'	.CallBackR = "SolRflipper"	'...
''	.CallBackUL = "SolULflipper"'...(upper flippers, if needed)
''	.CallBackUR = "SolURflipper"'...
'	.TiltObjects = True 'Optional, if True calls vpmnudge.solgameon automatically. IF YOU GET A LINE 1 ERROR, DISABLE THIS! (or setup vpmNudge.TiltObj!)
''	.InitDelay "FastFlips", 100			'Optional, if > 0 adds some compensation for solenoid jitter (occasional problem on Bram Stoker's Dracula)
''	.DebugOn = False		'Debug, always-on flippers. Call FastFlips.DebugOn True or False in debugger to enable/disable.
'end with

'...keydown section... commenting out upper flippers is not necessary as of 1.1
'If KeyCode = LeftFlipperKey then FastFlips.FlipL True :  FastFlips.FlipUL True
'If KeyCode = RightFlipperKey then FastFlips.FlipR True :  FastFlips.FlipUR True
'(Do not use Exit Sub, this script does not handle switch handling at all!)

'...keyUp section...
'If KeyCode = LeftFlipperKey then FastFlips.FlipL False :  FastFlips.FlipUL False
'If KeyCode = RightFlipperKey then FastFlips.FlipR False :  FastFlips.FlipUR False

'...Solenoid...
'SolCallBack(31) = "FastFlips.TiltSol"
'//////for a reference of solenoid numbers, see top /////


'One last note - Because this script is super simple it will call flipper return a lot.
'It might be a good idea to add extra conditional logic to your flipper return sounds so they don't play every time the game on solenoid turns off
'Example:
'Instead of
		'LeftFlipper.RotateToStart
		'playsound SoundFX("FlipperDown",DOFFlippers), 0, 1, 0.01	'return
'Add Extra conditional logic:
		'LeftFlipper.RotateToStart
		'if LeftFlipper.CurrentAngle = LeftFlipper.StartAngle then
		'	playsound SoundFX("FlipperDown",DOFFlippers), 0, 1, 0.01	'return
		'end if
'That's it]
'*************************************************
Function NullFunction(aEnabled):End Function	'1 argument null function placeholder
Class cFastFlips
	Public TiltObjects, DebugOn, hi
	Private SubL, SubUL, SubR, SubUR, FlippersEnabled, Delay, LagCompensation, Name, FlipState(3)

	Private Sub Class_Initialize()
		Delay = 0 : FlippersEnabled = False : DebugOn = False : LagCompensation = False
		Set SubL = GetRef("NullFunction"): Set SubR = GetRef("NullFunction") : Set SubUL = GetRef("NullFunction"): Set SubUR = GetRef("NullFunction")
	End Sub

	'set callbacks
	Public Property Let CallBackL(aInput)  : Set SubL  = GetRef(aInput) : Decouple sLLFlipper, aInput: End Property
	Public Property Let CallBackUL(aInput) : Set SubUL = GetRef(aInput) : End Property
	Public Property Let CallBackR(aInput)  : Set SubR  = GetRef(aInput) : Decouple sLRFlipper, aInput:  End Property
	Public Property Let CallBackUR(aInput) : Set SubUR = GetRef(aInput) : End Property
	Public Sub InitDelay(aName, aDelay) : Name = aName : delay = aDelay : End Sub	'Create Delay
	'Automatically decouple flipper solcallback script lines (only if both are pointing to the same sub) thanks gtxjoe
	Private Sub Decouple(aSolType, aInput)  : If StrComp(SolCallback(aSolType),aInput,1) = 0 then SolCallback(aSolType) = Empty End If : End Sub

	'call callbacks
	Public Sub FlipL(aEnabled)
		FlipState(0) = aEnabled	'track flipper button states: the game-on sol flips immediately if the button is held down (1.1)
		If not FlippersEnabled and not DebugOn then Exit Sub
		subL aEnabled
	End Sub

	Public Sub FlipR(aEnabled)
		FlipState(1) = aEnabled
		If not FlippersEnabled and not DebugOn then Exit Sub
		subR aEnabled
	End Sub

	Public Sub FlipUL(aEnabled)
		FlipState(2) = aEnabled
		If not FlippersEnabled and not DebugOn then Exit Sub
		subUL aEnabled
	End Sub

	Public Sub FlipUR(aEnabled)
		FlipState(3) = aEnabled
		If not FlippersEnabled and not DebugOn then Exit Sub
		subUR aEnabled
	End Sub

	Public Sub TiltSol(aEnabled)	'Handle solenoid / Delay (if delayinit)
		If delay > 0 and not aEnabled then 	'handle delay
			vpmtimer.addtimer Delay, Name & ".FireDelay" & "'"
			LagCompensation = True
		else
			If Delay > 0 then LagCompensation = False
			EnableFlippers(aEnabled)
		end If
	End Sub

	Sub FireDelay() : If LagCompensation then EnableFlippers False End If : End Sub

	Private Sub EnableFlippers(aEnabled)
		If aEnabled then SubL FlipState(0) : SubR FlipState(1) : subUL FlipState(2) : subUR FlipState(3)
		FlippersEnabled = aEnabled
		If TiltObjects then vpmnudge.solgameon aEnabled
		If Not aEnabled then
			subL False
			subR False
			If not IsEmpty(subUL) then subUL False
			If not IsEmpty(subUR) then subUR False
		End If
	End Sub

End Class

' EOF

' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


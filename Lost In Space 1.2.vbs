Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="lostspc",UseSolenoids=2,UseLamps=0,UseGI=0, SCoin="coin"

' Thalamus 2020 April : Improved directional sounds
' !! NOTE : Table not verified yet !!

' Options
' Volume devided by - lower gets higher sound

Const VolDiv = 5000    ' Lower number, louder ballrolling/collition sound
Const VolCol = 10      ' Ball collition divider ( voldiv/volcol )

' The rest of the values are multipliers
'
'  .5 = lower volume
' 1.5 = higher volume

Const VolBump   = 2    ' Bumpers volume.
Const VolFlip   = 1    ' Flipper volume.
Const VolSpin   = 3    ' Spinners volume.

LoadVPM "01530000","sega.vbs",3.1

Dim DesktopMode: DesktopMode = Table1.ShowDT

If DesktopMode = True Then 'Show Desktop components
  Ramp16.visible=1
  Ramp15.visible=1
  Primitive13.visible=1
Else
  Ramp16.visible=0
  Ramp15.visible=0
  Primitive13.visible=0
End if

'*************************************************************
'Solenoid Call backs
'*************************************************************

SolCallback(1)=		"SolTrough"
SolCallback(2)=		"Auto_Plunger"
SolCallback(3)=		"VukTopPop"
SolCallback(4)=		"VukBotPop"
'SolCallback(5)= 	'Top Left Magnet
SolCallback(6)=		"VukLeftPop"
'SolCallback(7)= 	'Top Right Magnet
'SolCallback(8)='Ticket Dispenser
'SolCallback(9)="vpmSolSound ""Jet3"","
'SolCallback(10)="vpmSolSound ""Jet3"","
'SolCallback(11)="vpmSolSound ""Jet3"","
'SolCallback(12)="vpmSolSound ""Jet3"","
'SolCallback(13)="vpmSolSound ""Jet3"","
SolCallback(14)=	"SolDiscMagnet"
SolCallback(17)=	"SolSpinWheelsMotor"
'SolCallback(18)="vpmSolSound ""Sling"","
'SolCallback(19)="vpmSolSound ""Sling"","
SolCallback(20)=	"SolRobot" 'Robot (Shaker)
'21-23  UK Only
'SolCallback(24)= 'Coin Meter
SolCallback(25)=	"SetLamp 125,"'25'F1 Red X4
SolCallback(26)=	"SetLamp 126,"'26'F2 Yellow X4
SolCallback(27)=	"SetLamp 127,"'27'F3 Green X4
SolCallback(28)=	"SetLamp 128,"'28'F4 Warning X4
SolCallback(29)=	"SetLamp 129,"'29'F5 Warning X3
SolCallback(30)=	"SetLamp 130,"'30 F6 Pops X3
SolCallback(31)=	"SetLamp 131,"'31 F7 Pops X2
SolCallback(32)=	"SetLamp 132,"'32 F8 Ramp X2

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors), LeftFlipper, VolFlip:LeftFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors), LeftFlipper, VolFlip:LeftFlipper.RotateToStart
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors), RightFlipper, VolFlip:RightFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors), RightFlipper, VolFlip:RightFlipper.RotateToStart
     End If
End Sub
'**********************************************************************************************************

'Solenoid Controlled toys
'**********************************************************************************************************

Sub SolTrough(Enabled)
	If Enabled Then
		If bsTrough.Balls Then
			vpmTimer.PulseSw 15
			bsTrough.ExitSol_On
		End If
	End If
End Sub

Sub Auto_Plunger(Enabled) 'plunger
	If Enabled Then
		PlungerIM.AutoFire
	End If
End Sub

Sub SolDiscMagnet(Enabled)
	'debug.print "Magnet: " & Enabled
	Magnetdisc.MagnetOn=Enabled
	if not Enabled Then
		Magnetdisc.GrabCenter = False
	end if
End Sub

Sub SolRobot(Enabled)
	If Enabled Then
		PrimRobot.TransZ = 15
		PrimRobot1.TransZ = 15
		playsoundAtVol SoundFX("Popper",DOFContactors), PrimRobot1, 1
	Else
		PrimRobot.TransZ = 0
		PrimRobot1.TransZ = 0
		'PlaySound "fx_relay"
	End If
End Sub

set GICallback = GetRef("UpdateGI")
Sub UpdateGI(no, Enabled)
	If Enabled Then
		dim xx
		For each xx in GI:xx.State = 1: Next
        PlaySound "fx_relay"
	Else
		For each xx in GI:xx.State = 0: Next
        PlaySound "fx_relay"
	End If
End Sub

'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************

Dim bsTrough, TopLMag, TopRMag
Dim TSPINA, Magnetdisc

Sub Table1_Init
	vpmInit Me
	On Error Resume Next
		With Controller
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
		.SplashInfoLine = ""&chr(13)&"You Suck"
		.HandleMechanics=0
		.HandleKeyboard=0
		.ShowDMDOnly=1
		.ShowFrame=0
		.ShowTitle=0
		.hidden = 0
        .Games(cGameName).Settings.Value("sound")=1
		'.PuPHide = 1
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With
     On Error Goto 0

    PinMAMETimer.Interval=PinMAMEInterval:
	PinMAMETimer.Enabled=1

	vpmNudge.TiltSwitch=56:
	vpmNudge.Sensitivity= 2
	vpmNudge.TiltObj=Array(S36,S37,S38,S39,S40,LeftSlingshot,RightSlingshot)

	Set bsTrough=New cvpmBallStack
		bsTrough.InitSw 0,14,13,12,11,0,0,0
		bsTrough.InitKick BallRelease,95,10
		bsTrough.Balls=4
		bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)

	Set TopLMag=New cvpmMagnet
		TopLMag.InitMagnet TopLeftMag,200
		TopLMag.Solenoid=5
		TopLMag.GrabCenter=1
		TopLMag.CreateEvents"TopLMag"

	Set TopRMag=New cvpmMagnet
		TopRMag.InitMagnet TopRightMag,50
		TopRMag.Solenoid=7
		TopRMag.GrabCenter=1
		TopRMag.CreateEvents"TopRMag"

	Set TSPINA = New cvpmTurntable
		TSPINA.InitTurntable SpinTrigger, 35
		TSPINA.SpinDown = 10
		TSPINA.CreateEvents "TSPINA"

	Set Magnetdisc = New cvpmMagnet
		Magnetdisc.InitMagnet iman, 22
		Magnetdisc.strength = 8
		Magnetdisc.GrabCenter = false
		Magnetdisc.MagnetOn = False
		Magnetdisc.CreateEvents "Magnetdisc"

End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)
	'If KeyCode=KeyUpperLeft Then Controller.Switch(1)=1 'UK Only
	'If KeyCode=KeyUpperRight Then Controller.Switch(8)=1 'UK Only
	If KeyCode=PlungerKey Then Controller.Switch(53)=1

	If KeyDownHandler(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
	'If KeyCode=KeyUpperLeft Then Controller.Switch(1)=0 'UK Only
	'If KeyCode=KeyUpperRight Then Controller.Switch(8)=0 'UK Only
	If KeyCode=PlungerKey Then Controller.Switch(53)=0

	If KeyUpHandler(keycode) Then Exit Sub
End Sub

	' Impulse Plunger
	dim plungerIM

	Const IMPowerSetting = 40 ' Plunger Power
	Const IMTime = 0.6        ' Time in seconds for Full Plunge
	Set plungerIM = New cvpmImpulseP
	With plungerIM
		.InitImpulseP swPlunger, IMPowerSetting, IMTime
		.Random 0.3
		.Switch 16
		.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
		.CreateEvents "plungerIM"
	End With

'**********************************************************************************************************

 ' Drain hole and kickers
Sub Drain_Hit:bsTrough.addball me : playsoundAtVol"drain" , Drain, 1: End Sub

 '***********************************
 'Left Raising VUK sw46
 '***********************************
 'Variables used for VUK
 Dim Lraiseballsw, Lraiseball

 Sub LeftVUK_Hit()
	playsoundAtVol "popper_ball", LeftVUK, 1
	StopSound "subway"
 	LeftVUK.Enabled=FALSE
	Controller.switch (46) = True
 End Sub

 Sub VukLeftPop(enabled)
	if(enabled and Controller.switch (46)) then
		playsoundAtVol SoundFX("Popper",DOFContactors), LeftVUK, 1
		LeftVUK.DestroyBall
 		Set Lraiseball= LeftVUK.CreateBall
 		Lraiseballsw = True
 		LeftVUKraiseballtimer.Enabled = True 'Added by Rascal
		LeftVUK.Enabled=TRUE
 		Controller.switch (46) = False
	else
		'PlaySound "Popper"
	end if
End Sub

 Sub LeftVUKraiseballtimer_Timer()
 	If Lraiseballsw = True then
 		Lraiseball.z = Lraiseball.z + 10
 		If Lraiseball.z > 50 then
 			'msgbox ("Over")
 			LeftVUK.Kick 160, 15
 			Set Lraiseball= Nothing
 			LeftVUKraiseballtimer.Enabled = False
 			Lraiseballsw = False
			'PlaySound "Wire Ramp"
 		End If
 	End If
 End Sub

 '***********************************
 'Top Raising VUK sw47
 '***********************************
 'Variables used for VUK
 Dim raiseballsw, raiseball

 Sub TopVUK_Hit()
	playsoundAtVol "popper_ball", TopVUK, 1
 	TopVUK.Enabled=FALSE
	Controller.switch (47) = True
 End Sub

 Sub VukTopPop(enabled)
	if(enabled and Controller.switch (47)) then
		playsoundAtVol SoundFX("Popper",DOFContactors), TopVUK, 1
		TopVUK.DestroyBall
 		Set raiseball = TopVUK.CreateBall
 		raiseballsw = True
 		TopVukraiseballtimer.Enabled = True 'Added by Rascal
		TopVUK.Enabled=TRUE
 		Controller.switch (47) = False
	else
		'PlaySound "Popper"
	end if
End Sub

 Sub TopVukraiseballtimer_Timer()
 	If raiseballsw = True then
 		raiseball.z = raiseball.z + 10
 		If raiseball.z > 120 then
 			'msgbox ("Over")
 			TopVUK.Kick 295, 15
 			Set raiseball = Nothing
 			TopVukraiseballtimer.Enabled = False
 			raiseballsw = False
			PlaySoundAtVol "Wire Ramp", TopVUK, 1
 		End If
 	End If
 End Sub

 '***********************************
 'Bottom Raising VUK sw48 'Layer 9 need a Prim PF
 '***********************************
 Dim braiseballsw, braiseball

 Sub BottomVuk_Hit()
	playsoundAtVol"popper_ball", BottomVUK, 1
	StopSound "subway"
 	BottomVUK.Enabled=FALSE
	Controller.switch (48) = True
 End Sub

 Sub VukBotPop(enabled)
	if(enabled and Controller.switch (48)) then
		playsoundAtVol SoundFX("Popper",DOFContactors), BottomVUK, 1
		BottomVUK.DestroyBall
 		Set braiseball = BottomVUK.CreateBall
 		braiseballsw = True
 		BottomVukraiseballtimer.Enabled = True 'Added by Rascal
		BottomVUK.Enabled=TRUE
 		Controller.switch (48) = False
	else
		'PlaySound "Popper"
	end if
End Sub

 Sub BottomVukraiseballtimer_Timer()
 	If braiseballsw = True then
 		braiseball.z = braiseball.z + 10
 		If braiseball.z > 90 then
 			'msgbox ("Over")
 			BottomVUK.Kick 295, 20
 			Set braiseball = Nothing
 			BottomVukraiseballtimer.Enabled = False
 			braiseballsw = False
			PlaySoundAtVol "Wire Ramp", BottomVUK, 1
 		End If
 	End If
 End Sub
 '***********************************
 '***********************************

'Subways
Sub Sw41_Hit:Controller.Switch(41)=1 :  playsoundAtVol "subway", ActiveBall, 1 : End Sub
Sub Sw41_unHit:Controller.Switch(41)=0:End Sub
Sub Sw45_Hit:Controller.Switch(45)=1 :  playsoundAtVol "subway" , ActiveBall, 1: End Sub
Sub Sw45_unHit:Controller.Switch(45)=0:End Sub

'Wire Triggers
Sub sw16_Hit : : playsoundAtVol "rollover" , ActiveBall, 1: End Sub 'Coded to impulse plunger
'Sub sw16_unHit : Controller.Switch (16)=0:End Sub
Sub S30_Hit:Controller.Switch(30)=1 : playsoundAtVol "rollover" , ActiveBall, 1: End Sub
Sub S30_unHit:Controller.Switch(30)=0:End Sub
Sub S31_Hit:Controller.Switch(31)=1 : playsoundAtVol "rollover" , ActiveBall, 1: End Sub
Sub S31_unHit:Controller.Switch(31)=0:End Sub
Sub S32_Hit:Controller.Switch(32)=1 : playsoundAtVol "rollover" , ActiveBall, 1: End Sub
Sub S32_unHit:Controller.Switch(32)=0:End Sub
Sub S51_Hit:Controller.Switch(51)=1 : playsoundAtVol "rollover" , ActiveBall, 1: End Sub
Sub S51_unHit:Controller.Switch(51)=0:End Sub
Sub S52_Hit:Controller.Switch(52)=1 : playsoundAtVol "rollover" , ActiveBall, 1: End Sub
Sub S52_unHit:Controller.Switch(52)=0:End Sub
Sub S57_Hit:Controller.Switch(57)=1 : playsoundAtVol "rollover" , ActiveBall, 1: End Sub
Sub S57_unHit:Controller.Switch(57)=0:End Sub
Sub S58_Hit:Controller.Switch(58)=1 : playsoundAtVol "rollover" , ActiveBall, 1: End Sub
Sub S58_unHit:Controller.Switch(58)=0:End Sub
Sub S60_Hit:Controller.Switch(60)=1 : playsoundAtVol "rollover" , ActiveBall, 1: End Sub
Sub S60_unHit:Controller.Switch(60)=0:End Sub
Sub S61_Hit:Controller.Switch(61)=1 : playsoundAtVol "rollover" , ActiveBall, 1: End Sub
Sub S61_unHit:Controller.Switch(61)=0:End Sub

'Stand Up Targets
Sub S17_Hit:vpmTimer.PulseSw 17:End Sub
Sub S18_Hit:vpmTimer.PulseSw 18:End Sub
Sub S19_Hit:vpmTimer.PulseSw 19:End Sub
Sub S25_Hit:vpmTimer.PulseSw 25:End Sub
Sub S26_Hit:vpmTimer.PulseSw 26:End Sub
Sub S27_Hit:vpmTimer.PulseSw 27:End Sub
Sub S33_Hit:vpmTimer.PulseSw 33:End Sub
Sub S34_Hit:vpmTimer.PulseSw 34:End Sub
Sub S35_Hit:vpmTimer.PulseSw 35:End Sub

'Ramp Gate Triggers
Sub S20_Hit:vpmTimer.PulseSw 20:End Sub
Sub S21_Hit:vpmTimer.PulseSw 21:End Sub

'Bumpers
Sub S36_Hit:vpmTimer.PulseSw 36 : playsoundAtVol SoundFX("fx_bumper1",DOFContactors), ActiveBall, VolBump: End Sub
Sub S37_Hit:vpmTimer.PulseSw 37 : playsoundAtVol SoundFX("fx_bumper1",DOFContactors), ActiveBall, VolBump: End Sub
Sub S38_Hit:vpmTimer.PulseSw 38 : playsoundAtVol SoundFX("fx_bumper1",DOFContactors), ActiveBall, VolBump: End Sub
Sub S39_Hit:vpmTimer.PulseSw 39 : playsoundAtVol SoundFX("fx_bumper1",DOFContactors), ActiveBall, VolBump: End Sub
Sub S40_Hit:vpmTimer.PulseSw 40 : playsoundAtVol SoundFX("fx_bumper1",DOFContactors), ActiveBall, VolBump: End Sub

'Spinners
Sub S49_Spin:vpmTimer.PulseSw 49 : playsoundAtVol "fx_spinner" , S49, VolSpin: End Sub
Sub S50_Spin:vpmTimer.PulseSw 50 : playsoundAtVol "fx_spinner" , S50, VolSpin: End Sub


'Generic Sounds
Sub Trigger1_Hit:StopSound "Wire Ramp":PlaySoundAtVol "fx_ballrampdrop", ActiveBall, 1:End Sub
Sub Trigger2_Hit:PlaySoundAtVol "fx_ballrampdrop", ActiveBall, 1:End Sub
Sub Trigger3_Hit:StopSound "Wire Ramp":PlaySoundAtVol "fx_ballrampdrop", ActiveBall, 1:End Sub


'********************  Spinning Discs Animation Timer ****************************
Dim SpinnerMotorOff, SpinnerStep, ss

Sub SolSpinWheelsMotor(enabled)
	If enabled Then
		TSPINA.MotorOn = True
		SpinnerStep = 10
		SpinnerMotorOff = False
		SpinnerTimer.Interval = 10
		SpinnerTimer.enabled = True
	Else
		SpinnerMotorOff = True
		TSPINA.MotorOn = False
	end If
End Sub

Sub SpinnerTimer_Timer()
	If Not(SpinnerMotorOff) Then
		spina.ObjRotZ  = ss
		ss = ss + SpinnerStep
	Else
		if SpinnerStep < 0 Then
			SpinnerTimer.enabled = False
		Else
		'slow the rate of spin by decreasing rotation step
			SpinnerStep = SpinnerStep - 0.05

			spina.ObjRotZ  = ss
			ss = ss + SpinnerStep
		End If
	End If
	if ss > 360 then ss = ss - 360
End Sub


'*************************************************************************
'*************************************************************************


'***************************************************
'       JP's VP10 Fading Lamps & Flashers
'       Based on PD's Fading Light System
' SetLamp 0 is Off
' SetLamp 1 is On
' fading for non opacity objects is 4 steps
'***************************************************

Dim LampState(200), FadingLevel(200)
Dim FlashSpeedUp(200), FlashSpeedDown(200), FlashMin(200), FlashMax(200), FlashLevel(200)

InitLamps()             ' turn off the lights and flashers and reset them to the default parameters
LampTimer.Interval = 5 'lamp fading speed
LampTimer.Enabled = 1

' Lamp & Flasher Timers

Sub LampTimer_Timer()
    Dim chgLamp, num, chg, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp) Then
        For ii = 0 To UBound(chgLamp)
            LampState(chgLamp(ii, 0) ) = chgLamp(ii, 1)       'keep the real state in an array
            FadingLevel(chgLamp(ii, 0) ) = chgLamp(ii, 1) + 4 'actual fading step

	   'Special Handling
	   'If chgLamp(ii,0) = 2 Then solTrough chgLamp(ii,1)
	   'If chgLamp(ii,0) = 4 Then PFGI chgLamp(ii,1)

        Next
    End If
    UpdateLamps
End Sub


Sub UpdateLamps()
NFadeL 1, L1
NFadeL 2, L2
NFadeL 3, L3
NFadeL 4, L4
NFadeL 5, L5
NFadeL 6, L6
NFadeL 7, L7
NFadeL 8, L8
NFadeL 9, L9
NFadeL 10, L10
NFadeL 11, L11
NFadeL 12, L12
NFadeL 13, L13
NFadeL 14, L14
NFadeL 15, L15
'NFadeL 16, L16 'Launch Button
NFadeL 17, L17
NFadeL 18, L18
NFadeL 19, L19

NFadeL 25, L25
NFadeL 26, L26
NFadeL 27, L27
NFadeL 28, L28

NFadeL 33, L33
NFadeL 34, L34
NFadeL 35, L35
NFadeLm 36, L36a 'Bumper
NFadeL 36, L36b
NFadeLm 37, L37a 'Bumper
NFadeL 37, L37b
NFadeLm 38, L38a 'Bumper
NFadeL 38, L38b
NFadeLm 39, L39a 'Bumper
NFadeL 39, L39b
NFadeLm 40, L40a 'Bumper
NFadeL 40, L40b
NFadeL 41, L41

NFadeL 43, L43
NFadeL 44, L44
NFadeL 45, L45
NFadeL 46, L46
NFadeL 47, L47
NFadeL 48, L48
NFadeL 49, L49

NFadeL 51, L51
NFadeL 52, L52
NFadeL 53, L53
NFadeL 54, L54
NFadeL 55, L55
NFadeL 56, L56
NFadeL 57, L57
NFadeL 58, L58

NFadeL 60, L60
NFadeL 61, L61
NFadeL 62, L62
NFadeL 63, L63
NFadeL 64, L64

'Solenoid Controlled

NFadeLm 125, F1A
NFadeLm 125, F1B
NFadeLm 125, F1C
NFadeL 125, F1D

NFadeLm 126, F2A
NFadeLm 126, F2B
NFadeLm 126, F2C
NFadeL 126, F2D

NFadeLm 127, F3A
NFadeLm 127, F3B
NFadeLm 127, F3C
NFadeL 127, F3D

NFadelm 128, F4a
NFadelm 128, F4b
NFadelm 128, F4c
NFadel 128, F4d

NFadelm 129, F5a
NFadelm 129, F5b
NFadel 129, F5c

NFadelm 130, F6a
NFadelm 130, F6b
NFadel 130, F6c

NFadelm 131, F7a
NFadel 131, F7b

NFadelm 132, F8a
NFadel 132, F8b

End Sub




' div lamp subs

Sub InitLamps()
    Dim x
    For x = 0 to 200
        LampState(x) = 0        ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4      ' used to track the fading state
        FlashSpeedUp(x) = 0.4   ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.2 ' slower speed when turning off the flasher
        FlashMax(x) = 1         ' the maximum value when on, usually 1
        FlashMin(x) = 0         ' the minimum value when off, usually 0
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
End Sub

Sub Flashm(nr, object) 'multiple flashers, it just sets the flashlevel
    Object.IntensityScale = FlashLevel(nr)
End Sub

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
	vpmTimer.PulseSw 62
    PlaySoundAtVol SoundFX("right_slingshot",DOFContactors), sling1, 1
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.rotx = 20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.rotx = 10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.rotx = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
	vpmTimer.PulseSw 59
    PlaySoundAtVol SoundFX("left_slingshot",DOFContactors), sling2, 1
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.rotx = 20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.rotx = 10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.rotx = 0:LeftSlingShot.TimerEnabled = 0:
    End Select
    LStep = LStep + 1
End Sub

' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX, Rothbauerw and Herweh
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

' set position as table object and Vol + RndPitch manually

Sub PlaySoundAtVolPitch(sound, tableobj, Vol, RndPitch)
  PlaySound sound, 1, Vol, AudioPan(tableobj), RndPitch, 0, 0, 1, AudioFade(tableobj)
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

Sub PlaySoundAtBallAbsVol(sound, VolMult)
  PlaySound sound, 0, VolMult, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub

' requires rampbump1 to 7 in Sound Manager

Sub RandomBump(voladj, freq)
  Dim BumpSnd:BumpSnd= "rampbump" & CStr(Int(Rnd*7)+1)
  PlaySound BumpSnd, 0, Vol(ActiveBall)*voladj, AudioPan(ActiveBall), 0, freq, 0, 1, AudioFade(ActiveBall)
End Sub

' set position as bumperX and Vol manually. Allows rapid repetition/overlaying sound

Sub PlaySoundAtBumperVol(sound, tableobj, Vol)
  PlaySound sound, 1, Vol, AudioPan(tableobj), 0,0,1, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBOTBallZ(sound, BOT)
  PlaySound sound, 0, ABS(BOT.velz)/17, Pan(BOT), 0, Pitch(BOT), 1, 0, AudioFade(BOT)
End Sub

' play a looping sound at a location with volume
Sub PlayLoopSoundAtVol(sound, tableobj, Vol)
  PlaySound sound, -1, Vol, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
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

Function BallRollVol(ball) ' Calculates the Volume of the sound based on the ball speed
  BallRollVol = Csng(BallVel(ball) ^2 / (80000 - (79900 * Log(RollVol) / Log(100))))
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

'********************************************************************
'      JP's VP10 Rolling Sounds (+rothbauerw's Dropping Sounds)
'********************************************************************

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

    For b = 0 to UBound(BOT)
        ' play the rolling sound for each ball
        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If

        ' play ball drop sounds
        If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
            PlaySound "fx_ball_drop" & b, 0, ABS(BOT(b).velz)/17, AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
        End If
    Next
End Sub

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
	PlaySound("fx_collide"), 0, Csng(velocity) ^2 / (VolDiv/VolCol), AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub


'*****************************************
'	ninuzzu's	FLIPPER SHADOWS
'*****************************************

sub FlipperTimer_Timer()
	FlipperLSh.RotZ = LeftFlipper.currentangle
	FlipperRSh.RotZ = RightFlipper.currentangle
End Sub

'*****************************************
'	ninuzzu's	BALL SHADOW
'*****************************************
Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5)

Sub BallShadowUpdate_timer()
    Dim BOT, b
    BOT = GetBalls
    ' hide shadow of deleted balls
    If UBound(BOT)<(tnob-1) Then
        For b = (UBound(BOT) + 1) to (tnob-1)
            BallShadow(b).visible = 0
        Next
    End If
    ' exit the Sub if no balls on the table
    If UBound(BOT) = -1 Then Exit Sub
    ' render the shadow for each ball
    For b = 0 to UBound(BOT)
        If BOT(b).X < Table1.Width/2 Then
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) + 6
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) - 6
        End If
        ballShadow(b).Y = BOT(b).Y + 12
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
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
	PlaySound "pinhit_low", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
	PlaySound "target", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
	PlaySound "metalhit_thin", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
	PlaySound "metalhit_medium", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
	PlaySound "metalhit2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
	PlaySound "gate4", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinner_Spin
	PlaySound "fx_spinner", 0, .25, AudioPan(Spinner), 0.25, 0, 0, 1, AudioFade(Spinner)
End Sub

Sub Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub Posts_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 16 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 16 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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
		Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub


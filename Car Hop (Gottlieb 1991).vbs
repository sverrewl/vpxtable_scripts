Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="carhop",UseSolenoids=1,UseLamps=0,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff", SCoin="coin"

LoadVPM "01120100", "gts3.vbs", 3.26

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
'**********************************************************************************************************
'SolCallback(1)  = "vpmSolSound ""lSling"","
'SolCallback(2)  = "vpmSolSound ""lSling"","
SolCallback(3)  = "bsRKicker.SolOut"
SolCallback(4)  = "bsLKicker.SolOut"
SolCallback(5)  = "dtLBank.SolDropUp" 'Drop Targets
SolCallback(6)  = "dtRBank.SolDropUp" 'Drop Targets
SolCallback(7)  = "dtTBank.SolDropUp" 'Drop Targets
SolCallback(9)  = "SolDrop1"
SolCallback(10) = "SolDrop2"
SolCallback(11) = "SolDrop3"
SolCallback(12) = "SolDrop4"
SolCallback(13) = "SolDrop5"
SolCallback(14) = "SolTop"
SolCallback(15) = "SetLamp 171," 'Bottom Left 67 Flasher
SolCallBack(16) = "SetLamp 172," 'Bottom Right 67 Flasher
SolCallback(17) = "SetLamp 173," 'Top 67 Flasher
SolCallback(18) = "SetLamp 176," 'Bullseye 67 Flasher
SolCallback(19) = "SetLamp 174," 'Right Drop 67 Flasher
SolCallback(20) = "SetLamp 175," 'Left Drop 67 Flasher

'SolCallback(23)="vpmFlasher Array(BlueCar1,BlueCar2),"
'SolCallback(24)="vpmFlasher Array(RedCar1,RedCar2),"
'SolCallback(25) = "vpmFlasher SunFlasher,"
SolCallback(29) = "bsTrough.SolOut" 
SolCallback(30)=  "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
'SolCallBack(31)=
SolCallback(32)= "vpmNudge.SolGameOn"

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFX("fx_Flipperup",DOFContactors):LeftFlipper.RotateToEnd
     Else
         PlaySound SoundFX("fx_Flipperdown",DOFContactors):LeftFlipper.RotateToStart
     End If
  End Sub
  
Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFX("fx_Flipperup",DOFContactors):RightFlipper.RotateToEnd:RightFlipper1.RotateToEnd
     Else
         PlaySound SoundFX("fx_Flipperdown",DOFContactors):RightFlipper.RotateToStart:RightFlipper1.RotateToStart
     End If
End Sub
'**********************************************************************************************************

'Solenoid Controlled toys
'**********************************************************************************************************

	Sub SolTop(Enabled)
		If Enabled Then
			dtTBank.Hit 1
		End If
	End Sub

	Sub SolDrop1(Enabled)
		If Enabled Then
			dtLBank.Hit 1
			dtRBank.Hit 5
		End If
	End Sub

	Sub SolDrop2(Enabled)
		If Enabled Then
			dtLBank.Hit 2
			dtRBank.Hit 4
		End If
	End Sub

	Sub SolDrop3(Enabled)
		If Enabled Then
			dtLBank.Hit 3
			dtRBank.Hit 3
		End If
	End Sub

	Sub SolDrop4(Enabled)
		If Enabled Then
			dtLBank.Hit 4
			dtRBank.Hit 2
		End If
	End Sub

	Sub SolDrop5(Enabled)
		If Enabled Then
			dtLBank.Hit 5
			dtRBank.Hit 1
		End If
	End Sub

'*****GI Lights On
dim xx
For each xx in GI:xx.State = 1: Next

'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************

Dim bsTrough, bsLKicker, bsRKicker, dtLBank, dtRBank, dtTBank, mhole, mhole1

Sub Table1_Init
	vpmInit Me
	On Error Resume Next
		With Controller
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
		.SplashInfoLine = "Car Hop Gottlieb/Premier 1990"&chr(13)&"You Suck"
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

	PinMAMETimer.Interval = PinMAMEInterval 
	PinMAMETimer.Enabled = 1
	vpmNudge.TiltSwitch = 151
	vpmNudge.Sensitivity = 2
	vpmNudge.TiltObj = Array(RightSlingShot, LeftSlingShot)

     Set bsTrough = New cvpmBallStack
         '.InitSw 0, 0, 0, 0, 0, 0, 0, 0			'set trough switches
         bsTrough.InitNoTrough BallRelease, 25, 80, 6
         bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
         bsTrough.Balls = 1	

     Set bsRKicker = New cvpmBallStack
         bsRKicker.InitSaucer sw40, 40, 230, 10
         bsRKicker.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
         bsRKicker.KickAngleVar = 2

     Set bsLKicker = New cvpmBallStack
         bsLKicker.InitSaucer sw41, 41, 140, 10
         bsLKicker.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
         bsLKicker.KickAngleVar = 2

     set dtLBank = new cvpmdroptarget
         dtLBank.initdrop array(sw16, sw26, sw36, sw46, sw56), array(16, 26, 36, 46, 56)
         dtLBank.initsnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

     set dtRBank = new cvpmdroptarget
         dtRBank.initdrop array(sw17, sw27, sw37, sw47, sw57), array(17, 27, 37, 47, 57)
         dtRBank.initsnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

	 set dtTBank = new cvpmdroptarget
         dtTBank.initdrop sw15, 15
         dtTBank.initsnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

     Set mHole = New cvpmMagnet
     With mHole
         .InitMagnet Umagnet2, 4
         '.GrabCenter = 0
         .MagnetOn = 1
         .CreateEvents "mHole"
     End With

     Set mHole1 = New cvpmMagnet
     With mHole1
         .InitMagnet Umagnet1, 3
         '.GrabCenter = 0
         .MagnetOn = 1
         .CreateEvents "mHole1"
     End With

End Sub



'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal keycode)

	If keycode= PlungerKey Then Plunger.Pullback:playsound"plungerpull"

	If keycode= StartGameKey Then Controller.Switch(3)= 1

	If KeyCode= LeftFlipperKey Then 
		Controller.Switch(4)=1
		Controller.Switch(6)=1
	End If

	If KeyCode= RightFlipperkey Then
		Controller.Switch(5)=1
		Controller.Switch(7)=1
	End If

	If vpmKeyDown(KeyCode) Then Exit Sub

End Sub

Sub Table1_KeyUp(ByVal keycode)

	If keycode= PlungerKey Then Plunger.Fire:playsound "plunger"

	If keycode= StartGameKey Then Controller.Switch(3)= 0

	If KeyCode= LeftFlipperKey Then 
		Controller.Switch(4)=0
		Controller.Switch(6)=0
	End If

	If KeyCode= RightFlipperkey Then
		Controller.Switch(5)=0
		Controller.Switch(7)=0
	End If

	If vpmKeyUp(KeyCode) Then Exit Sub

End Sub
'**********************************************************************************************************

 ' Drain hole and kickers
Sub Drain_Hit:bsTrough.addball me : playsound"drain" : End Sub
Sub sw40_Hit:bsRKicker.AddBall 0 : playsound "popper_ball": End Sub
Sub sw41_Hit:bsLKicker.Addball 0 : playsound "popper_ball": End Sub


'Drop Targets
 Sub Sw16_Dropped:dtLBank.Hit 1 :End Sub  
 Sub Sw26_Dropped:dtLBank.Hit 2 :End Sub  
 Sub Sw36_Dropped:dtLBank.Hit 3 :End Sub
 Sub Sw46_Dropped:dtLBank.Hit 4 :End Sub
 Sub Sw56_Dropped:dtLBank.Hit 5 :End Sub

 Sub Sw17_Dropped:dtRBank.Hit 1 :End Sub  
 Sub Sw27_Dropped:dtRBank.Hit 2 :End Sub  
 Sub Sw37_Dropped:dtRBank.Hit 3 :End Sub
 Sub Sw47_Dropped:dtRBank.Hit 4 :End Sub
 Sub Sw57_Dropped:dtRBank.Hit 5 :End Sub

 Sub Sw15_Dropped:dtTBank.Hit 1 :End Sub

 'Stand Up Targets
Sub sw12_hit:vpmTimer.pulseSw 12 : End Sub 
Sub sw22_hit:vpmTimer.pulseSw 22 : End Sub
Sub sw32_hit:vpmTimer.pulseSw 32 : End Sub

'Bullseye
 Sub sw33_Hit:vpmTimer.PulseSw 33:sw33.TimerEnabled = 1:sw33p.TransX = -12 : playsound"target" : End Sub
 Sub sw33_unHit:sw33p.TransX = 0:End Sub
 Sub sw43_Hit:vpmTimer.PulseSw 43:sw43.TimerEnabled = 1:sw33p.TransX = -12 : playsound"target" : End Sub
 Sub sw43_unHit:sw33p.TransX = 0:End Sub

'Spinners
Sub sw13_Spin:vpmTimer.PulseSw 13 : playsound"fx_spinner" : End Sub

'Wire Triggers
	Sub sw14_Hit:Controller.Switch(14) = 1 : playsound"rollover" : End Sub 
	Sub sw14_UnHit:Controller.Switch(14) = 0:End Sub
	Sub sw30_Hit:Controller.Switch(30) = 1 : playsound"rollover" : End Sub 
	Sub sw30_UnHit:Controller.Switch(30) = 0:End Sub
	Sub sw31_Hit:Controller.Switch(31) = 1 : playsound"rollover" : End Sub 
	Sub sw31_UnHit:Controller.Switch(31) = 0:End Sub
	Sub sw34_Hit:Controller.Switch(34) = 1 : playsound"rollover" : End Sub 
	Sub sw34_UnHit:Controller.Switch(34) = 0:End Sub
	Sub sw42_Hit:Controller.Switch(42) = 1 : playsound"rollover" : End Sub 
	Sub sw42_UnHit:Controller.Switch(42) = 0:End Sub
	Sub sw44_Hit:Controller.Switch(44) = 1 : playsound"rollover" : End Sub 
	Sub sw44_UnHit:Controller.Switch(44) = 0:End Sub
	Sub sw55_Hit:Controller.Switch(55) = 1 : playsound"rollover" : End Sub 
	Sub sw55_UnHit:Controller.Switch(55) = 0:End Sub

'Star Triggers
	Sub sw20_Hit:Controller.Switch(20) = 1 : playsound"rollover" : End Sub 
	Sub sw20_UnHit:Controller.Switch(20) = 0:End Sub
	Sub sw21_Hit:Controller.Switch(21) = 1 : playsound"rollover" : End Sub 
	Sub sw21_UnHit:Controller.Switch(21) = 0:End Sub
	Sub sw23_Hit:Controller.Switch(23) = 1 : playsound"rollover" : End Sub 
	Sub sw23_UnHit:Controller.Switch(23) = 0:End Sub
	Sub sw24_Hit:Controller.Switch(24) = 1 : playsound"rollover" : End Sub 
	Sub sw24_UnHit:Controller.Switch(24) = 0:End Sub

 'Scoring Rubber
Sub sw45a_hit:vpmTimer.pulseSw 45 : playsound"flip_hit_3" : End Sub 
Sub sw45b_hit:vpmTimer.pulseSw 45 : playsound"flip_hit_3" : End Sub 

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
        Next
    End If
    UpdateLamps
End Sub

Sub UpdateLamps
	NFadeL 0, L0
	NFadeL 1, L1
	NFadeL 2, L2
	NFadeL 3, L3
	NFadeL 4, L4
	NFadeL 5, L5
	NFadeL 6, L6
	NFadeL 7, L7
	NFadeL 10, L10
	NFadeL 11, L11
	NFadeL 12, L12
	NFadeL 13, L13
	NFadeL 14, L14
	NFadeL 15, L15
	NFadeL 16, L16
	NFadeL 17, L17
	NFadeL 20, L20
	NFadeL 21, L21
	NFadeL 22, L22
	NFadeL 23, L23
	NFadeL 24, L24
	NFadeL 25, L25
	NFadeL 26, L26
	NFadeL 27, L27
	NFadeL 30, L30
	NFadeL 31, L31
	NFadeL 32, L32
	NFadeL 33, L33
	NFadeL 34, L34
	NFadeL 35, L35
	NFadeL 36, L36
	NFadeL 37, L37
	NFadeL 40, L40
	NFadeL 41, L41
	NFadeL 42, L42
	NFadeL 43, L43
	NFadeL 44, L44
	NFadeL 45, L45
	NFadeL 46, L46
	NFadeL 47, L47
	NFadeL 50, L50
	NFadeL 51, L51
	NFadeL 52, L52
	NFadeL 53, L53
	NFadeL 54, L54
	NFadeL 55, L55
	NFadeL 60, L60
	NFadeL 61, L61
	NFadeL 62, L62
	NFadeL 63, L63
	NFadeL 64, L64
	NFadeL 65, L65
	NFadeL 70, L70
	NFadeL 71, L71
	NFadeL 72, L72
	NFadeL 73, L73
	NFadeL 74, L74
	NFadeL 75, L75
	NFadeL 76, L76
	NFadeL 80, L80
	NFadeL 81, L81
	NFadeL 82, L82
	NFadeL 83, L83
	NFadeL 84, L84
	NFadeL 85, L85
	NFadeL 86, L86
	NFadeL 87, L87

'Solenoid Controlled Lights

	NFadeLm 171, s171
	NFadeL 171, s171a

	NFadeLm 172, s172
	NFadeL 172, s172a

	NFadeLm 173, s173
	NFadeLm 173, s173a
	NFadeLm 173, s173b
	NfadeL 173, s173c

	NFadeLm 174, s174
	NFadeLm 174, s174a
	NFadeLm 174, s174b
	NFadeLm 174, s174c

	NFadeLm 175, s175
	NFadeLm 175, s175a
	NFadeLm 175, s175b
	NFadeLm 175, s175c

	NFadeLm 176, s176
	NFadeLm 176, s176a
	NFadeLm 176, s176c
	NFadeL 176, s176d
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


'**********************************************************************************************************
'Digital Display
'**********************************************************************************************************
 Dim Digits(40)
 Digits(0)=Array(a00, a05, a0c, a0d, a08, a01, a06, a0f, a02, a03, a04, a07, a0b, a0a, a09, a0e)
 Digits(1)=Array(a10, a15, a1c, a1d, a18, a11, a16, a1f, a12, a13, a14, a17, a1b, a1a, a19, a1e)
 Digits(2)=Array(a20, a25, a2c, a2d, a28, a21, a26, a2f, a22, a23, a24, a27, a2b, a2a, a29, a2e)
 Digits(3)=Array(a30, a35, a3c, a3d, a38, a31, a36, a3f, a32, a33, a34, a37, a3b, a3a, a39, a3e)
 Digits(4)=Array(a40, a45, a4c, a4d, a48, a41, a46, a4f, a42, a43, a44, a47, a4b, a4a, a49, a4e)
 Digits(5)=Array(a50, a55, a5c, a5d, a58, a51, a56, a5f, a52, a53, a54, a57, a5b, a5a, a59, a5e)
 Digits(6)=Array(a60, a65, a6c, a6d, a68, a61, a66, a6f, a62, a63, a64, a67, a6b, a6a, a69, a6e)
 Digits(7)=Array(a70, a75, a7c, a7d, a78, a71, a76, a7f, a72, a73, a74, a77, a7b, a7a, a79, a7e)
 Digits(8)=Array(a80, a85, a8c, a8d, a88, a81, a86, a8f, a82, a83, a84, a87, a8b, a8a, a89, a8e)
 Digits(9)=Array(a90, a95, a9c, a9d, a98, a91, a96, a9f, a92, a93, a94, a97, a9b, a9a, a99, a9e)
 Digits(10)=Array(aa0, aa5, aac, aad, aa8, aa1, aa6, aaf, aa2, aa3, aa4, aa7, aab, aaa, aa9, aae)
 Digits(11)=Array(ab0, ab5, abc, abd, ab8, ab1, ab6, abf, ab2, ab3, ab4, ab7, abb, aba, ab9, abe)
 Digits(12)=Array(ac0, ac5, acc, acd, ac8, ac1, ac6, acf, ac2, ac3, ac4, ac7, acb, aca, ac9, ace)
 Digits(13)=Array(ad0, ad5, adc, add, ad8, ad1, ad6, adf, ad2, ad3, ad4, ad7, adb, ada, ad9, ade)
 Digits(14)=Array(ae0, ae5, aec, aed, ae8, ae1, ae6, aef, ae2, ae3, ae4, ae7, aeb, aea, ae9, aee)
 Digits(15)=Array(af0, af5, afc, afd, af8, af1, af6, aff, af2, af3, af4, af7, afb, afa, af9, afe)
 Digits(16)=Array(b00, b05, b0c, b0d, b08, b01, b06, b0f, b02, b03, b04, b07, b0b, b0a, b09, b0e)
 Digits(17)=Array(b10, b15, b1c, b1d, b18, b11, b16, b1f, b12, b13, b14, b17, b1b, b1a, b19, b1e)
 Digits(18)=Array(b20, b25, b2c, b2d, b28, b21, b26, b2f, b22, b23, b24, b27, b2b, b2a, b29, b2e)
 Digits(19)=Array(b30, b35, b3c, b3d, b38, b31, b36, b3f, b32, b33, b34, b37, b3b, b3a, b39, b3e)

 Digits(20)=Array(b40, b45, b4c, b4d, b48, b41, b46, b4f, b42, b43, b44, b47, b4b, b4a, b49, b4e)
 Digits(21)=Array(b50, b55, b5c, b5d, b58, b51, b56, b5f, b52, b53, b54, b57, b5b, b5a, b59, b5e)
 Digits(22)=Array(b60, b65, b6c, b6d, b68, b61, b66, b6f, b62, b63, b64, b67, b6b, b6a, b69, b6e)
 Digits(23)=Array(b70, b75, b7c, b7d, b78, b71, b76, b7f, b72, b73, b74, b77, b7b, b7a, b79, b7e)
 Digits(24)=Array(b80, b85, b8c, b8d, b88, b81, b86, b8f, b82, b83, b84, b87, b8b, b8a, b89, b8e)
 Digits(25)=Array(b90, b95, b9c, b9d, b98, b91, b96, b9f, b92, b93, b94, b97, b9b, b9a, b99, b9e)
 Digits(26)=Array(ba0, ba5, bac, bad, ba8, ba1, ba6, baf, ba2, ba3, ba4, ba7, bab, baa, ba9, bae)
 Digits(27)=Array(bb0, bb5, bbc, bbd, bb8, bb1, bb6, bbf, bb2, bb3, bb4, bb7, bbb, bba, bb9, bbe)
 Digits(28)=Array(bc0, bc5, bcc, bcd, bc8, bc1, bc6, bcf, bc2, bc3, bc4, bc7, bcb, bca, bc9, bce)
 Digits(29)=Array(bd0, bd5, bdc, bdd, bd8, bd1, bd6, bdf, bd2, bd3, bd4, bd7, bdb, bda, bd9, bde)
 Digits(30)=Array(be0, be5, bec, bed, be8, be1, be6, bef, be2, be3, be4, be7, beb, bea, be9, bee)
 Digits(31)=Array(bf0, bf5, bfc, bfd, bf8, bf1, bf6, bff, bf2, bf3, bf4, bf7, bfb, bfa, bf9, bfe)
 Digits(32)=Array(ac18, ac16, acc1, acd1, ac19, ac17, acf1, ac11, ac13, ac12, ac14, acb1, aca1, ac10, ac15)
 Digits(33)=Array(ad18, ad16, adc1, add1, ad19, ad17, adf1, ad11, ad13, ad12, ad14, adb1, ada1, ad10, ad15)
 Digits(34)=Array(ae18, ae16, aec1, aed1, ae19, ae17, aef1, ae11, ae13, ae12, ae14, aeb1, aea1, ae10, ae15)
 Digits(35)=Array(af18, af16, afc1, afd1, af19, af17, aff1, af11, af13, af12, af14, afb1, afa1, af10, af15)
 Digits(36)=Array(b9, b7, b0c1, b0d1, b100, b8, b0f1, b2, b4, b3, b5, b0b1, b0a1, b1, b6)
 Digits(37)=Array(b109, b107, b1c1, b1d1, b110, b108, b1f1, b102, b104, b103, b105, b1b1, b1a1, b101, b106)
 Digits(38)=Array(b119, b117, b2c1, b2d1, b120, b118, b2f1, b112, b114, b113, b115, b2b1, b2a1, b111, b116)
 Digits(39)=Array(b129, b127, b3c1, b3d1, b130, b128, b3f1, b3a1, b124, b123, b125, b3b1, b3a1, b121, b126)


 Sub DisplayTimer_Timer
    Dim ChgLED, ii, jj, num, chg, stat, obj, b, x
    ChgLED=Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
		If DesktopMode = True Then
       For ii=0 To UBound(chgLED)
          num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)
			if (num < 40) then
              For Each obj In Digits(num)
                   If chg And 1 Then obj.State=stat And 1
                   chg=chg\2 : stat=stat\2
                  Next
			Else
			       end if
        Next
	   end if
    End If
 End Sub
'**********************************************************************************************************
'**********************************************************************************************************

'*********************************************************************
'                 Start of VPX Functions
'*********************************************************************\

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
	vpmTimer.PulseSw 11
    PlaySound SoundFX("right_slingshot",DOFContactors), 0,1, 0.05,0.05 '0,1, AudioPan(RightSlingShot), 0.05,0,0,1,AudioFade(RightSlingShot)
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
	vpmTimer.PulseSw 10
    PlaySound SoundFX("left_slingshot",DOFContactors), 0,1, -0.05,0.05 '0,1, AudioPan(LeftSlingShot), 0.05,0,0,1,AudioFade(LeftSlingShot)
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub


'*********************************************************************
'                 Positional Sound Playback Functions
'*********************************************************************

' Play a sound, depending on the X,Y position of the table element (especially cool for surround speaker setups, otherwise stereo panning only)
' parameters (defaults): loopcount (1), volume (1), randompitch (0), pitch (0), useexisting (0), restart (1))
' Note that this will not work (currently) for walls/slingshots as these do not feature a simple, single X,Y position
Sub PlayXYSound(soundname, tableobj, loopcount, volume, randompitch, pitch, useexisting, restart)
	PlaySound soundname, loopcount, volume, AudioPan(tableobj), randompitch, pitch, useexisting, restart, AudioFade(tableobj)
End Sub

' Similar subroutines that are less complicated to use (e.g. simply use standard parameters for the PlaySound call)
Sub PlaySoundAt(soundname, tableobj)
    PlaySound soundname, 1, 1, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname)
    PlaySoundAt soundname, ActiveBall
End Sub


'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

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

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function


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
        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
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
	PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub


'*****************************************
'	ninuzzu's	FLIPPER SHADOWS
'*****************************************

sub FlipperTimer_Timer()
	FlipperLSh.RotZ = LeftFlipper.currentangle
	FlipperRSh.RotZ = RightFlipper.currentangle
	FlipperT3L.Roty = LeftFlipper.currentangle +240
	FlipperT3R.Roty = RightFlipper.currentangle +120
	FlipperT3R1.Roty = RightFlipper1.currentangle +120
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

Option Explicit
Randomize

' Thalamus 2018-07-20
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
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
Const VolRH     = 1    ' Rubber hits volume.
Const VolPo     = 1    ' Rubber posts volume.
Const VolPi     = 1    ' Rubber pins volume.
Const VolTarg   = 1    ' Targets volume.
Const VolSpin   = 1.5  ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01120100", "gts3.vbs", 3.02
Dim DesktopMode: DesktopMode = Table1.ShowDT

Const cGameName="deadweap",UseSolenoids=2,UseLamps=1,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff", SCoin="coin"

'*************************************************************
'Solenoid Call backs
'**********************************************************************************************************
'Solenoids increase by 1 because first solenoid is Zero
SolCallback(7)    = "SolKickback"
SolCallback(8)    = "dtDrop.SolDropUp"
SolCallback(9)    = "bsL2Kicker.SolOut"
SolCallback(10)   = "bsL3Kicker.SolOut"

SolCallback(26)   = "GIController"     'lightbox relay / GI controller
SolCallback(28)	  = "bsTrough.SolOut"
SolCallback(29)	  = "bsTrough.SolIn"
SolCallback(30)   = "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"

'playfield flashers
SolCallback(11)   = "vpmFlasher array(Flasher11,Flasher11a),"  '"Multi-Ball"
SolCallback(12)   = "vpmFlasher Flasher12,"  '"Special"
SolCallback(13)   = "vpmFlasher array(Flasher13,Flasher13a),"  '"Gas Station"
SolCallback(14)   = "vpmFlasher array(Flasher14,Flasher14a),"  '"Alley"
SolCallback(15)   = "vpmFlasher Flasher15,"  '"Burger Grill"
SolCallback(16)   = "vpmFlasher Flasher16,"  '"Bank"
SolCallback(17)   = "vpmFlasher array(Flasher17,Flasher17a),"  '"Mall"
SolCallback(18)   = "vpmFlasher array(Flasher18,Flasher18a),"  '"High School"

'backboard flashers
'SolCallback(19)   =
'SolCallback(20)   =  '
'SolCallback(21)   =  '
'SolCallback(22)   =
'SolCallback(23)   =
'SolCallback(24)   =
'SolCallback(25)   =

Sub SolKickBack(enabled)
    If enabled Then
       Plunger1.Fire
       PlaySoundAtVol SoundFX("Popper",DOFContactors), Plunger1, 1
    Else
       Plunger1.PullBack
    End If
End Sub


SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors), LeftFlipper, VolFlip:LeftFlipper.RotateToEnd:LeftFlipper1.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors),LeftFlipper, VolFlip:LeftFlipper.RotateToStart:LeftFlipper1.RotateToStart
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors),RightFlipper, VolFlip:RightFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors), RightFlipper, VolFlip:RightFlipper.RotateToStart
     End If
End Sub
'**********************************************************************************************************

'Solenoid Controlled toys
'**********************************************************************************************************

Sub FlipperTimer_Timer
	FlipperT1.roty = LeftFlipper.currentangle - 120
	FlipperT2.roty = LeftFlipper1.currentangle - 160
	FlipperT3.roty = RightFlipper.currentangle + 120
End Sub

'GI Lights always on  turn off if solenoid is triggered
Sub GIController(Enabled)
	If Enabled Then
		dim xx
		For each xx in GI:xx.State = 0: Next
	Else
		For each xx in GI:xx.State = 1: Next
	End if
 End Sub

'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************

Dim bsTrough,bsL2Kicker,bsL3Kicker,dtDrop

Sub Table1_Init
	vpmInit Me
	On Error Resume Next
		With Controller
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
		.SplashInfoLine = "Deadly Weapon (Gottlieb)"&chr(13)&"You Suck"
		.HandleMechanics=0
		.HandleKeyboard=0
		.ShowDMDOnly=1
		.ShowFrame=0
		.ShowTitle=0
        .hidden = 1
		If Err Then MsgBox Err.Description
	End With
	On Error Goto 0
		Controller.SolMask(0)=0
      vpmTimer.AddTimer 2000,"Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the timer to renable all the solenoids after 2 seconds
		Controller.Run
	If Err Then MsgBox Err.Description
	On Error Goto 0


	PinMAMETimer.Interval=PinMAMEInterval
	PinMAMETimer.Enabled=1
	' Nudging
	vpmNudge.TiltSwitch=151
	vpmNudge.Sensitivity=5
	vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3,Bumper4,RightSlingShot)

	Set bsTrough=new cvpmBallStack
	bsTrough.InitSw 16,26,0,0,0,0,0,0
	bsTrough.InitKick BallRelease,80,6
	bsTrough.InitExitSnd SoundFX("Ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
	bsTrough.Balls=4

	set bsL2Kicker = new cvpmBallStack
	bsL2Kicker.InitSaucer kicker1,20,180,5
	bsL2Kicker.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)

	set bsL3Kicker = new cvpmBallStack
	bsL3Kicker.InitSaucer kicker2,21, 200, 12
	bsL3Kicker.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)

	set dtDrop = new cvpmDropTarget
	dtDrop.InitDrop Array(sw24,sw34,sw44), Array(24,34,44)
	dtDrop.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

	plunger1.pullback
End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************
Sub Table1_KeyDown(ByVal keycode)
    If keycode=LeftFlipperKey Then Controller.Switch(6)=1
	If keycode=RightFlipperkey Then	Controller.Switch(7)=1
	iF KEYCODE = StartGameKey THEN CONTROLLER.SWITCH(3)=1
	If keycode = KeySlamDoorHit Then Controller.Switch(152) = 1
	If keycode = KeyUpperLeft Then controller.switch(4)=1
	If keycode = KeyUpperRight Then controller.switch(5)=1

	If KeyDownHandler(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Pullback:playsoundAtVol"plungerpull", Plunger, 1
End Sub

Sub Table1_KeyUp(ByVal keycode)
    If keycode=LeftFlipperKey Then Controller.Switch(6)=0
	If keycode=RightFlipperkey Then	Controller.Switch(7)=0
	iF KEYCODE =StartGameKey THEN CONTROLLER.SWITCH(3)=0
	If keycode = KeySlamDoorHit Then Controller.Switch(152) = 0
	If keycode = KeyUpperLeft Then controller.switch(4)=0
	If keycode = KeyUpperRight Then controller.switch(5)=0

	If KeyUpHandler(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Fire:PlaySoundAtVol"plunger", Plunger, 1
End Sub

'**********************************************************************************************************
' Drain hole
Sub Drain_Hit:playsoundAtVol"drain", drain, 1:bsTrough.addball me:End Sub
Sub Kicker1_Hit():bsL2Kicker.AddBall 0:End Sub
Sub Kicker2_Hit():bsL3Kicker.AddBall 0:End Sub

'Drop Targets
 Sub Sw24_Hit:dtDrop.Hit 1 :End Sub
 Sub Sw34_Hit:dtDrop.Hit 2 :End Sub
 Sub Sw44_Hit:dtDrop.Hit 3 :End Sub

'Bumpers
Sub Bumper1_Hit : vpmTimer.PulseSwitch 10,0,0 : playsoundAtVol SoundFX("fx_bumper1",DOFContactors), Bumper1, VolBump: End Sub
Sub Bumper2_Hit : vpmTimer.PulseSwitch 11,0,0 : playsoundAtVol SoundFX("fx_bumper1",DOFContactors), Bumper2, VolBump: End Sub
Sub Bumper3_Hit : vpmTimer.PulseSwitch 12,0,0 : playsoundAtVol SoundFX("fx_bumper1",DOFContactors), Bumper3, VolBump: End Sub
Sub Bumper4_Hit : vpmTimer.PulseSwitch 13,0,0 : playsoundAtVol SoundFX("fx_bumper1",DOFContactors), Bumper4, VolBump: End Sub

'Wire Triggers
Sub sw22_Hit		: Controller.Switch(22) = 1 : playsoundAtVol"rollover", ActiveBall, VolRol : End Sub
Sub sw22_UnHit  	: Controller.Switch(22) = 0 : End Sub
Sub sw23_Hit		: Controller.Switch(23) = 1 : playsoundAtVol"rollover", ActiveBall, VolRol : End Sub
Sub sw23_UnHit  	: Controller.Switch(23) = 0 : End Sub
Sub sw25_Hit		: Controller.Switch(25) = 1 : playsoundAtVol"rollover", ActiveBall, VolRol : End Sub
Sub sw25_UnHit	    : Controller.Switch(25) = 0 : End Sub
Sub sw32_Hit		: Controller.Switch(32) = 1 : playsoundAtVol"rollover", ActiveBall, VolRol : End Sub
Sub sw32_UnHit  	: Controller.Switch(32) = 0 : End Sub
Sub sw33_Hit		: Controller.Switch(33) = 1 : playsoundAtVol"rollover", ActiveBall, VolRol : End Sub
Sub sw33_UnHit  	: Controller.Switch(33) = 0 : End Sub
Sub sw35_Hit		: Controller.Switch(35) = 1 : playsoundAtVol"rollover", ActiveBall, VolRol : End Sub
Sub sw35_UnHit  	: Controller.Switch(35) = 0 : End Sub
Sub sw40_Hit		: Controller.Switch(40) = 1 : playsoundAtVol"rollover", ActiveBall, VolRol : End Sub
Sub sw40_UnHit  	: Controller.Switch(40) = 0 : End Sub
Sub sw41_Hit		: Controller.Switch(41) = 1 : playsoundAtVol"rollover", ActiveBall, VolRol : End Sub
Sub sw41_UnHit  	: Controller.Switch(41) = 0 : End Sub
Sub sw42_Hit		: Controller.Switch(42) = 1 : playsoundAtVol"rollover", ActiveBall, VolRol : End Sub
Sub sw42_UnHit  	: Controller.Switch(42) = 0 : End Sub
Sub sw43_Hit		: Controller.Switch(43) = 1 : playsoundAtVol"rollover", ActiveBall, VolRol : End Sub
Sub sw43_UnHit  	: Controller.Switch(43) = 0 : End Sub
Sub sw45_Hit		: Controller.Switch(45) = 1 : playsoundAtVol"rollover", ActiveBall, VolRol : End Sub
Sub sw45_UnHit  	: Controller.Switch(45) = 0 : End Sub

'spinners
Sub Spinner1_Spin(): vpmTimer.PulseSwitch(30),0,"" : playsoundAtVol"fx_spinner", Spinner1, VolSpin : End Sub
Sub Spinner2_Spin(): vpmTimer.PulseSwitch(31),0,"" : playsoundAtVol"fx_spinner", Spinner2, VolSpin : End Sub

'StandUp Targets
Sub sw17_hit(): vpmTimer.PulseSwitch(17),0,"": End Sub
Sub sw27_hit(): vpmTimer.PulseSwitch(27),0,"": End Sub
Sub sw37_hit(): vpmTimer.PulseSwitch(37),0,"": End Sub
Sub sw47_hit(): vpmTimer.PulseSwitch(47),0,"": End Sub
Sub sw57_hit(): vpmTimer.PulseSwitch(57),0,"": End Sub

'Square StandUp Target
Sub sw14_hit(): vpmTimer.PulseSwitch(14),0,"": End Sub

'**********************************************************************************************************

     Lights(1)=array(Light1,Light1a) 'Top Left Bumper
     Set Lights(2)=Light2
     Set Lights(3)=Light3
     Set Lights(4)=Light4
     Set Lights(5)=Light5
     Set Lights(6)=Light6
     Set Lights(7)=Light7
     Set Lights(9)=Light9
     Set Lights(10)=Light10
     Set Lights(11)=Light11
     Set Lights(12)=Light12
     Set Lights(13)=Light13
     Set Lights(14)=Light14
     Set Lights(15)=Light15
     Set Lights(16)=Light16
     Set Lights(17)=Light17
     Set Lights(20)=Light20
     Set Lights(21)=Light21
     Set Lights(22)=Light22
     Set Lights(23)=Light23
     Set Lights(24)=Light24
     Set Lights(25)=Light25
     Set Lights(26)=Light26
     Set Lights(27)=Light27
     Set Lights(30)=Light30
     Set Lights(31)=Light31
     Set Lights(32)=Light32
     Set Lights(33)=Light33
     Set Lights(34)=Light34
     Set Lights(35)=Light35
     Set Lights(36)=Light36
     Set Lights(37)=Light37
     Set Lights(40)=Light40
     Set Lights(41)=Light41
     Set Lights(42)=Light42
     Set Lights(43)=Light43
     Set Lights(44)=Light44
     Set Lights(45)=Light45
     Set Lights(46)=Light46
     Set Lights(47)=Light47
     Set Lights(50)=Light50
     Set Lights(51)=Light51
     Set Lights(52)=Light52
     Set Lights(53)=Light53
     Set Lights(54)=Light54
     Set Lights(55)=Light55
     Set Lights(56)=Light56
     Set Lights(57)=Light57
     Set Lights(60)=Light60
     Set Lights(61)=Light61
     Set Lights(62)=Light62
     Set Lights(63)=Light63
     Set Lights(64)=Light64
     Set Lights(65)=Light65
     Set Lights(66)=Light66
     Set Lights(67)=Light67
	 Lights(70)=array(Light70,Light70a) 'Top Right Bumper
     Set Lights(71)=Light71
     Set Lights(72)=Light72
     Set Lights(73)=Light73
     Set Lights(74)=Light74
     Lights(75)=array(Light75,Light75a) 'Center Bumper
     Set Lights(76)=Light76
     Set Lights(77)=Light77
     Set Lights(80)=Light80
     Set Lights(81)=Light81
     Set Lights(82)=Light82
     Set Lights(83)=Light83
     Set Lights(90)=Light90
     Set Lights(91)=Light91
     Set Lights(92)=Light92
     Set Lights(93)=Light93
     Set Lights(94)=Light94
     Set Lights(95)=Light95
     Set Lights(96)=Light96
     Set Lights(97)=Light97
     Set Lights(100)=Light100
     Set Lights(101)=Light101
     Set Lights(102)=Light102
     Set Lights(103)=Light103
     Set Lights(104)=Light104
     Set Lights(105)=Light105
     Set Lights(106)=Light106


'**********************************************************************************************************
'Digital Display
'**********************************************************************************************************
 Dim Digits(48)
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


 Digits(40) = Array(LED10,LED11,LED12,LED13,LED14,LED15,LED16)
 Digits(41) = Array(LED20,LED21,LED22,LED23,LED24,LED25,LED26)
 Digits(42) = Array(LED30,LED31,LED32,LED33,LED34,LED35,LED36)
 Digits(43) = Array(LED40,LED41,LED42,LED43,LED44,LED45,LED46)
 Digits(44) = Array(LED50,LED51,LED52,LED53,LED54,LED55,LED56)
 Digits(45) = Array(LED60,LED61,LED62,LED63,LED64,LED65,LED66)
 Digits(46) = Array(LED80,LED81,LED82,LED83,LED84,LED85,LED86)
 Digits(47) = Array(LED90,LED91,LED92,LED93,LED94,LED95,LED96)

 Sub DisplayTimer_Timer
    Dim ChgLED, ii, jj, num, chg, stat, obj, b, x
    ChgLED=Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
		If DesktopMode = True Then
       For ii=0 To UBound(chgLED)
          num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)
			if (num < 48) then
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

' ********************************************************************************************************************************
' ********************************************************************************************************************************
   'Start of VPX  Call Back Functions
' ********************************************************************************************************************************
' ********************************************************************************************************************************


'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep

Sub RightSlingShot_Slingshot
	vpmTimer.PulseSwitch 15,0,0
    PlaySoundAtVol SoundFX("right_slingshot",DOFContactors), sling1, 1
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0:
    End Select
    RStep = RStep + 1
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

' The sound is played using the VOL, PAN and PITCH functions, so the volume and pitch of the sound
' will change according to the ball speed, and the PAN function will change the stereo position according
' to the position of the ball on the table.


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

Sub PlaySoundAtVol(sound, tableobj, Vol)
  PlaySound sound, 1, Vol, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
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
  Vol = Csng(BallVel(ball) ^2 / VolDiv)
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

' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


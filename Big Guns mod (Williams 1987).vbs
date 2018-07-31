Option Explicit
Randomize

' Thalamus 2018-07-19
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' No special SSF tweaks yet.

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="bguns_l8",UseSolenoids=2,UseLamps=1,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff",SCoin="coin"

LoadVPM "00990300", "S11.VBS", 3.36
Dim DesktopMode: DesktopMode = Table1.ShowDT

'Solenoid Call backs
'**********************************************************************************************************

SolCallback(1)="bsTrough.SolIn"
SolCallback(2)="bsTrough.SolOut"
SolCallback(3)="SolTowerPopper"
SolCallback(4)="bsLeftEject.SolOut"
SolCallback(5)="bsRightEject.SolOut"
SolCallback(6)="bsLeftCannon.SolOut"
SolCallback(7)="bsRightCannon.SolOut"
SolCallback(8)="vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(13)="dtr.SolDropUp"
SolCallback(14)="FF"
SolCallback(17)="KingChamber"'Backbox Kicker (King's Chamber)

SolCallback(18) = "SolAPlunger"
SolCallBack(19) = "SolFlipperDiverter"
SolCallback(20)="dtl.SolDropUp"

	SolCallback(sLRFlipper) = "SolRFlipper"
	SolCallback(sLLFlipper) = "SolLFlipper"

	Sub SolLFlipper(Enabled)
		 If Enabled Then
			 PlaySound SoundFX("fx_Flipperup",DOFContactors):LeftFlipper.RotateToEnd:LeftFlipper1.RotateToEnd
		 Else
			 PlaySound SoundFX("fx_Flipperdown",DOFContactors):LeftFlipper.RotateToStart:LeftFlipper1.RotateToStart
		 End If
	  End Sub

	Sub SolRFlipper(Enabled)
		 If Enabled Then
			 PlaySound SoundFX("fx_Flipperup",DOFContactors):RightFlipper.RotateToEnd:RightFlipper1.RotateToEnd
		 Else
			 PlaySound SoundFX("fx_Flipperdown",DOFContactors):RightFlipper.RotateToStart:RightFlipper1.RotateToStart
		 End If
	End Sub

'SolCallback(12)=			'A/C Select

'Flashers and GI controlled by solenoids
SolCallback(9)= "LGI"	'Left GI Relay
SolCallback(10)= "RGI"	'Right GI Relay

'Solenoid enabled Flashers in VPX its just a regular light
SolCallback(15)= "vpmFlasher Flasher15,"'Playfield Invincible
SolCallback(25)= "vpmFlasher solenoid25,"'Playfield Back Panel
SolCallback(26)= "vpmFlasher Array(Flasher26,Flasher26b),"  'Lower Left / Queen Backglass
SolCallback(27)= "vpmFlasher Flasher27,"  'Lower Right
SolCallback(29)= "vpmFlasher Flasher29,"  'Left Cannon
SolCallback(30)= "vpmFlasher Flasher30,"  'Right Cannon
SolCallback(31)= "vpmFlasher Flasher31,"'Left Troll Playfield
SolCallback(32)= "vpmFlasher Flasher32,"'Right Troll Playfield

'Backglass ONLY
'SolCallback(11)=			'Backbox GI Relay
'SolCallback(16)=           'Backbox Top Flasher
SolCallback(28)= "vpmFlasher Flasher28b,"	'Mini Playfield
'SolCallback(29)=
'SolCallback(30)=
'SolCallback(31)=
'SolCallback(32)=
 '**********************************************************************************************************

'Solenoid Controlled toys
'**********************************************************************************************************

Sub SolFlipperDiverter(enabled)
	If Enabled Then
		 PlaySound SoundFX("fx_Flipperup",DOFContactors):Flipper3.RotateToEnd
	Else
		 PlaySound SoundFX("fx_Flipperdown",DOFContactors):Flipper3.RotateToStart
	End If
End Sub

Sub SolAPlunger(enabled)
    If enabled Then
       Plunger1.Fire
       PlaySound SoundFX("Popper",DOFContactors)
    Else
       Plunger1.PullBack
    End If
End Sub

Sub SolTowerPopper(Enabled)
	If Enabled Then
		bsPopper.ExitSol_On
		Kicker5.DestroyBall
	End If
End Sub

'Poppup Saver
Sub FF(Enabled)
	If Enabled Then
		If ForceField.IsDropped Then
			Controller.Switch(15)=0
			ForceField.IsDropped=0
			L6a.visible=1
			CenterPost.transz = 23
			Playsound SoundFX("Centerpost_Up",DOFContactors)
		Else
			ForceField.IsDropped=1
			Controller.Switch(15)=1
			L6a.visible=0
			CenterPost.transz = 0
			Playsound SoundFX("Centerpost_Down",DOFContactors)
		End If
	End If
End Sub

Sub LGI(Enabled)
	If Enabled Then
	'*****GI Lights Off
		dim xx
		For each xx in GIL:xx.State = 0: Next
	Else
		For each xx in GIL:xx.State = 1: Next
	End if
 End Sub

Sub RGI(Enabled)
	If Enabled Then
	'*****GI Lights Off
		dim xxx
		For each xxx in GIR:xxx.State = 0: Next
	Else
		For each xxx in GIR:xxx.State = 1: Next
	End if
 End Sub

Sub FlipperTimer_Timer
	diverter.roty = Flipper3.currentangle
End Sub


'Initiate Table
'**********************************************************************************************************


Dim bsTrough,bsLeftCannon,bsRightCannon,bsLeftEject,bsRightEject,dtL,dtR,bsPopper

Sub Table1_Init
	vpmInit Me
	On Error Resume Next
		With Controller
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
		.SplashInfoLine = "Big Guns (Williams)"&chr(13)&"You Suck"
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

	vpmNudge.TiltSwitch = swTilt
	vpmNudge.Sensitivity = 5
	vpmNudge.TiltObj=Array(LeftSlingshot,RightSlingshot)


	set bsTrough = new cvpmBallStack
	bsTrough.InitSw 10,11,12,13,0,0,0,0
	bsTrough.InitKick BallRelease,50,7
	bsTrough.InitExitSnd SoundFX("BallRelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
	bsTrough.Balls = 3

	set bsLeftCannon = new cvpmBallStack
	bsLeftCannon.InitSw 0,37,0,0,0,0,0,0
	bsLeftCannon.InitKick Kicker7,32,20
	bsLeftCannon.InitExitSnd SoundFX("popper",DOFContactors), SoundFX("Solenoid",DOFContactors)

	set bsRightCannon = new cvpmBallStack
	bsRightCannon.InitSw 0,38,0,0,0,0,0,0
	bsRightCannon.InitKick Kicker8,-35,20
	bsRightCannon.InitExitSnd SoundFX("popper",DOFContactors), SoundFX("Solenoid",DOFContactors)

	set bsLeftEject = new cvpmBallStack
	bsLeftEject.InitSaucer Kicker3,35,107,8
	bsLeftEject.InitExitSnd SoundFX("popper",DOFContactors), SoundFX("Solenoid",DOFContactors)

	set bsRightEject = new cvpmBallStack
	bsRightEject.InitSaucer Kicker4,36,267,8
	bsRightEject.InitExitSnd SoundFX("popper",DOFContactors), SoundFX("Solenoid",DOFContactors)

	Set dtL = New cvpmDropTarget
	dtL.InitDrop Array(sw17,sw18,sw19),Array(17,18,19)
	dtL.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

	Set dtR = New cvpmDropTarget
	dtR.InitDrop Array(sw20,sw21,sw22),Array(20,21,22)
	dtR.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

	Set bsPopper = new cvpmBallStack
	bsPopper.InitSw 0,57,0,0,0,0,0,0
	bsPopper.InitKick Kicker9,323,15
	bsPopper.InitExitSnd SoundFX("popper",DOFContactors), SoundFX("Solenoid",DOFContactors)

Kicker6.CreateBall.Image="Red"
Plunger1.Pullback
ForceField.IsDropped=1
Controller.Switch(15)=1
L6a.visible=0
End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************
Sub Table1_KeyUp(ByVal keycode)
	If keycode=leftflipperkey then controller.switch(52)=0
	If keycode=rightflipperkey then controller.switch(53)=0
	If KeyUpHandler(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Fire:PlaySound"plunger"
End Sub

Sub Table1_KeyDown(ByVal keycode)
	If keycode=leftflipperkey then controller.switch(52)=1
	If keycode=rightflipperkey then controller.switch(53)=1
	If KeyDownHandler(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Pullback:playsound"plungerpull"
End Sub

'**********************************************************************************************************
'kickers
Sub Drain_Hit:playsound"drain":bsTrough.addball me:End Sub

Sub Kicker1_Hit:bsLeftCannon.AddBall Me:End Sub
Sub Kicker2_Hit:bsRightCannon.AddBall Me:End Sub
Sub Kicker3_Hit:bsLeftEject.AddBall 0:End Sub
Sub Kicker4_Hit:bsRightEject.AddBall 0:End Sub
Sub Kicker5_Hit:bsPopper.AddBall 0:End Sub

Sub Kicker7_Hit:bsLeftCannon.AddBall Me:End Sub
Sub Kicker8_Hit:bsRightCannon.AddBall Me:End Sub

Sub Kicker10_Hit:bsLeftCannon.AddBall Me:End Sub
Sub Kicker11_Hit:bsrighteject.addball Me:End Sub


'Stand Up targets
Sub sw25_Hit:vpmTimer.PulseSw 25 :End Sub
Sub sw26_Hit:vpmTimer.PulseSw 26 :End Sub
Sub sw27_Hit:vpmTimer.PulseSw 27 :End Sub
Sub sw28_Hit:vpmTimer.PulseSw 28 :End Sub
Sub sw29_Hit:vpmTimer.PulseSw 29 :End Sub
Sub sw30_Hit:vpmTimer.PulseSw 30 :End Sub
Sub sw31_Hit:vpmTimer.PulseSw 31 :End Sub
Sub sw32_Hit:vpmTimer.PulseSw 32 :End Sub
Sub sw33_Hit:vpmTimer.PulseSw 33 :End Sub
Sub sw34_Hit:vpmTimer.PulseSw 34 :End Sub
Sub sw39_Hit:vpmTimer.PulseSw 39 :End Sub
Sub sw40_Hit:vpmTimer.PulseSw 40 :End Sub
Sub sw58_Hit:vpmTimer.PulseSw 58 :End Sub

'Drop Targets
 Sub Sw17_Hit: dtL.Hit 1 : End Sub
 Sub Sw18_Hit: dtL.Hit 2 : End Sub
 Sub Sw19_Hit: dtL.Hit 3 : End Sub

 Sub Sw20_Hit: dtR.Hit 1 : End Sub
 Sub Sw21_Hit: dtR.Hit 2 : End Sub
 Sub Sw22_Hit: dtR.Hit 3 : End Sub


'Wire Triggers
Sub sw23_Hit:Controller.Switch(23)=1 :playsound"rollover" : End Sub
Sub sw23_unHit:Controller.Switch(23)=0:End Sub
Sub sw24_Hit:Controller.Switch(24)=1 :playsound"rollover" : End Sub
Sub sw24_unHit:Controller.Switch(24)=0:End Sub
Sub sw60_Hit:Controller.Switch(60)=1 :playsound"rollover" : End Sub
Sub sw60_unHit:Controller.Switch(60)=0:End Sub
Sub sw61_Hit:Controller.Switch(61)=1 :playsound"rollover" : End Sub
Sub sw61_unHit:Controller.Switch(61)=0:End Sub
Sub sw54_Hit:Controller.Switch(54)=1 :playsound"rollover" : End Sub
Sub sw54_unHit:Controller.Switch(54)=0:End Sub
Sub sw55_Hit:Controller.Switch(55)=1 :playsound"rollover" : End Sub
Sub sw55_unHit:Controller.Switch(55)=0:End Sub
Sub sw56_Hit:Controller.Switch(56)=1 :playsound"rollover" : End Sub
Sub sw56_unHit:Controller.Switch(56)=0:End Sub
Sub sw59_Hit:Controller.Switch(59)=1 :playsound"rollover" : End Sub
Sub sw59_unHit:Controller.Switch(59)=0:End Sub

'Generic ramp sounds
Sub Trigger1_Hit: playsound"Wire Ramp" : End Sub
Sub Trigger2_Hit: playsound"Ball Drop" : End Sub
Sub Trigger3_Hit: playsound"Wire Ramp" : End Sub
Sub Trigger9_Hit: playsound"Ball Drop" : End Sub
Sub Trigger10_Hit: playsound"Ball Drop" : End Sub
Sub Trigger11_Hit: playsound"Wire Ramp" : End Sub
'**********************************************************************************************************

'Backbox playfield
Dim R'Random Kick variable force
Randomize'reset random number seed

Sub KingChamber(Enabled)
	If Enabled Then
		R=RND(1)*60
		Kicker6.Kick 0,70+R
		PlaySound SoundFX("Popper",DOFContactors)
	End If
End Sub

Sub Trigger4_Hit:Controller.Switch(41)=1:End Sub
Sub Trigger4_unHit:Controller.Switch(41)=0:End Sub
Sub Trigger5_Hit:Controller.Switch(42)=1:End Sub
Sub Trigger5_unHit:Controller.Switch(42)=0:End Sub
Sub Trigger6_Hit:Controller.Switch(43)=1:End Sub
Sub Trigger6_unHit:Controller.Switch(43)=0:End Sub
Sub Trigger7_Hit:Controller.Switch(44)=1:End Sub
Sub Trigger7_unHit:Controller.Switch(44)=0:End Sub
Sub Trigger8_Hit:Controller.Switch(45)=1:End Sub
Sub Trigger8_unHit:Controller.Switch(45)=0:End Sub


'**********************************************************************************************************
'Map lights to an array
'**********************************************************************************************************

Lights(6)=array(L6,L6a)
Set Lights(7)=L7
Set Lights(8)=L8
Set Lights(10)=L10
Set Lights(11)=L11
Set Lights(12)=L12
Set Lights(13)=L13
Set Lights(14)=L14
Set Lights(15)=L15
Set Lights(16)=L16
Set Lights(17)=L17
Set Lights(18)=L18
Set Lights(19)=L19
Set Lights(20)=L20
Set Lights(21)=L21
Set Lights(22)=L22
Set Lights(23)=L23
Set Lights(24)=L24
Set Lights(25)=L25
Set Lights(26)=L26
Set Lights(27)=L27
Set Lights(28)=L28
Set Lights(29)=L29
Set Lights(30)=L30
Set Lights(31)=L31
Set Lights(32)=L32
Set Lights(33)=L33
Set Lights(34)=L34
Set Lights(35)=L35
Set Lights(36)=L36
Set Lights(37)=L37
Set Lights(38)=L38
Set Lights(39)=L39
Set Lights(40)=L40
Set Lights(41)=L41
Set Lights(42)=L42
Set Lights(43)=L43
Set Lights(44)=L44
Set Lights(45)=L45
Set Lights(46)=L46
Set Lights(47)=L47
Set Lights(48)=L48
Set Lights(54)=L54
Set Lights(55)=L55
Set Lights(56)=L56
Set Lights(62)=L62
Set Lights(63)=L63
Set Lights(64)=L64

'BackGlass
Set Lights(9)=Light9
Set Lights(1)=Light1
Set Lights(2)=Light2
Set Lights(3)=Light3
Set Lights(4)=Light4
Set Lights(5)=Light5

'Flahser Controlled by lights
Sub LSample_Timer()
	Flasher25.visible = solenoid25.state 'backwall light
End Sub
'**********************************************************************************************************
'Digital Display
'**********************************************************************************************************

 Dim Digits(28)
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

Digits(14) = Array(LED10,LED11,LED12,LED13,LED14,LED15,LED16)
Digits(15) = Array(LED20,LED21,LED22,LED23,LED24,LED25,LED26)
Digits(16) = Array(LED30,LED31,LED32,LED33,LED34,LED35,LED36)
Digits(17) = Array(LED40,LED41,LED42,LED43,LED44,LED45,LED46)
Digits(18) = Array(LED50,LED51,LED52,LED53,LED54,LED55,LED56)
Digits(19) = Array(LED60,LED61,LED62,LED63,LED64,LED65,LED66)
Digits(20) = Array(LED70,LED71,LED72,LED73,LED74,LED75,LED76)
Digits(21) = Array(LED80,LED81,LED82,LED83,LED84,LED85,LED86)
Digits(22) = Array(LED90,LED91,LED92,LED93,LED94,LED95,LED96)
Digits(23) = Array(LED100,LED101,LED102,LED103,LED104,LED105,LED106)
Digits(24) = Array(LED110,LED111,LED112,LED113,LED114,LED115,LED116)
Digits(25) = Array(LED120,LED121,LED122,LED123,LED124,LED125,LED126)
Digits(26) = Array(LED130,LED131,LED132,LED133,LED134,LED135,LED136)
Digits(27) = Array(LED140,LED141,LED142,LED143,LED144,LED145,LED146)

Sub DisplayTimer_Timer
	Dim ChgLED,ii,num,chg,stat,obj
	ChgLed = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
If Not IsEmpty(ChgLED) Then
		If DesktopMode = True Then
		For ii = 0 To UBound(chgLED)
			num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
			if (num < 28) then
				For Each obj In Digits(num)
					If chg And 1 Then obj.State = stat And 1
					chg = chg\2 : stat = stat\2
				Next
			else
				'if char(stat) > "" then msg(num) = char(stat)
			end if
		next
		end if
end if
End Sub

'**********************************************************************************************************
'**********************************************************************************************************
'**********************************************************************************************************
'VPX Functions
'**********************************************************************************************************
'**********************************************************************************************************
'**********************************************************************************************************

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
	vpmTimer.PulseSw 64
    PlaySound SoundFX("right_slingshot",DOFContactors), 0, 1, 0.05, 0.05
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

Sub LeftSlingShot_Slingshot
	vpmTimer.PulseSw 63
    PlaySound SoundFX("left_slingshot",DOFContactors),0,1,-0.05,0.05
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0:
    End Select
    LStep = LStep + 1
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
	PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
	PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
	PlaySound "metalhit_thin", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
	PlaySound "metalhit_medium", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
	PlaySound "metalhit2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
	PlaySound "gate4", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinner_Spin
	PlaySound "fx_spinner",0,.25,0,0.25
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

Sub RandomSoundFlipper()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub

Sub GraphicsTimer_Timer()
	batleft.objrotz = LeftFlipper.CurrentAngle + 1
	batright.objrotz = RightFlipper.CurrentAngle - 1
	batleft1.objrotz = LeftFlipper1.CurrentAngle + 1
	batright1.objrotz = RightFlipper1.CurrentAngle - 1
End Sub

' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX
' PlaySound sound, 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
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
  PlaySound sound, 1, Vol, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
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

Function AudioFade(ball) ' Can this be together with the above function ?
  Dim tmp
  tmp = ball.y * 2 / Table1.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
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
      If BallVel(BOT(b) ) > 1 Then
        rolling(b) = True
        if BOT(b).z < 30 Then ' Ball on playfield
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        Else ' Ball on raised ramp
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*.5, Pan(BOT(b) ), 0, Pitch(BOT(b) )+50000, 1, 0, AudioFade(BOT(b) )
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
  If Table1.VersionMinor > 3 OR Table1.VersionMajor > 10 Then
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
  Else
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0
  End if
End Sub


' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


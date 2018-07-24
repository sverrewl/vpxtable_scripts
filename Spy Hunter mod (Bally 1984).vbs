Option Explicit
Randomize

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' No special SSF tweaks yet.

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="spyhuntr",UseSolenoids=1,UseLamps=1,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff",SCoin="coin"

LoadVPM "01000100","BALLY.VBS",3.1
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
SolCallBack(10)="bsSaucer.SolOut"
SolCallback(11)="dtL.SolDropUp Not"
SolCallback(12)="vpmFlasher Flasher12,"   'Bumper sound turned to flasher
SolCallback(14)="bsTrough.SolOut"
SolCallback(15)=  "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(17)="SGToggle"
SolCallback(19)="vpmNudge.SolGameOn"

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

'**********************************************************************************************************

'Solenoid Controlled toys
'**********************************************************************************************************
Sub SGToggle(Enabled) 'Routine to handle the gate light accurately - not in light matrix
vpmSolDiverter Flipper3,True,Enabled
If GateLight.State=LightStateOn Then
	GateLight.State=LightStateOff

	Else
	GateLight.State=LightStateOn
End If
End Sub

'Primitive Flipper Code
Sub FlipperTimer_Timer
	Primitive9.roty = Flipper3.currentangle +90
	Primitive84.roty = LeftFlipper1.currentangle +220
	Primitive85.roty = RightFlipper1.currentangle +135
End Sub
'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************
Dim bsTrough,bsSaucer,dtL,cbCaptive

Sub Table1_Init

	On Error Resume Next
		With Controller
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
		.SplashInfoLine = "Spy Hunter (Bally)"&chr(13)&"You Suck"
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

	vpmNudge.TiltSwitch=15
	vpmNudge.Sensitivity=2
	vpmNudge.TiltObj=Array(Bumper1,LeftSlingshot,RightSlingshot)

	Set bsTrough=New cvpmBallStack
    bsTrough.InitNoTrough BallRelease,8,110,4
    bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)

	Set bsSaucer=New cvpmBallStack
	bsSaucer.InitSaucer Kicker2,5,199+rnd(1)*3,4+rnd(1)*3
	bsSaucer.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)

	Set dtL=New cvpmDropTarget
	dtL.InitDrop Array(sw17,sw18,sw19,sw20),Array(17,18,19,20)
	dtL.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

	Captive1.CreateBall
	Captive1.Kick 60,6
End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)
	If KeyCode=RightFlipperKey Then Controller.Switch(12)=1
	If KeyDownHandler(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Pullback:playsound"plungerpull"
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
	If KeyCode=RightFlipperKey Then Controller.Switch(12)=0
	If KeyUpHandler(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Fire:PlaySound"plunger"
End Sub

'**********************************************************************************************************

'Drain and Kicker
Sub Drain_Hit:playsound"drain":bsTrough.addball me:End Sub
Sub Kicker2_Hit:bsSaucer.AddBall 0 : playsound SoundFX("popper_ball",DOFContactors): End Sub

'Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 14 : playsound SoundFX("fx_bumper1",DOFContactors): End Sub

'Drop Targets
 Sub sw17_Dropped:dtL.Hit 1 :End Sub
 Sub sw18_Dropped:dtL.Hit 2 :End Sub
 Sub sw19_Dropped:dtL.Hit 3 :End Sub
 Sub sw20_Dropped:dtL.Hit 4 :End Sub

'Stand up targets
Sub Sw21_Hit:vpmTimer.PulseSw 21:End Sub
Sub Sw22_Hit:vpmTimer.PulseSw 22:End Sub
Sub Sw23_Hit:vpmTimer.PulseSw 23:End Sub
Sub Sw24_Hit:vpmTimer.PulseSw 24:End Sub

Sub Sw26_Hit:vpmTimer.PulseSw 26:End Sub
Sub Sw27_Hit:vpmTimer.PulseSw 27:End Sub

Sub Sw28_Hit:vpmTimer.PulseSw 28:End Sub
Sub Sw29_Hit:vpmTimer.PulseSw 29:End Sub
Sub Sw30_Hit:vpmTimer.PulseSw 30:End Sub

Sub sw31_Hit:vpmTimer.PulseSw 21:End Sub
Sub sw32_Hit:vpmTimer.PulseSw 32:End Sub

Sub sw37_Hit:vpmTimer.PulseSw 37:End Sub

'Spinner
Sub sw7_Spin:vpmTimer.PulseSw 7 : playsound"fx_spinner" : End Sub

'Wire Tirggers
Sub SW25_Hit:Controller.Switch(25)=1 : playsound"rollover" : End Sub
Sub SW25_unHit:Controller.Switch(25)=0:End Sub
Sub SW25a_Hit:Controller.Switch(25)=1 : playsound"rollover" : End Sub
Sub SW25a_unHit:Controller.Switch(25)=0:End Sub
Sub SW33_Hit:Controller.Switch(33)=1 : playsound"rollover" : End Sub
Sub SW33_unHit:Controller.Switch(33)=0:End Sub


'Weapon Triggers
Sub sw2_Hit:vpmTimer.PulseSw 2 : playsound"rollover" : End Sub
Sub sw3_Hit:vpmTimer.PulseSw 3 : playsound"rollover" : End Sub
Sub sw4_Hit:vpmTimer.PulseSw 4 : playsound"rollover" : End Sub

Sub sw38_Hit:vpmTimer.PulseSw 38 : playsound"rollover" : End Sub
Sub sw39_Hit:vpmTimer.PulseSw 39 : playsound"rollover" : End Sub
Sub sw40_Hit:vpmTimer.PulseSw 40 : playsound"rollover" : End Sub

Sub sw34_Hit:vpmTimer.PulseSw 34 : playsound"rollover" : End Sub
Sub sw35_Hit:vpmTimer.PulseSw 35 : playsound"rollover" : End Sub
Sub sw36_Hit:vpmTimer.PulseSw 36 : playsound"rollover" : End Sub
'**********************************************************************************************************

'Map lights to an array
'**********************************************************************************************************
Set Lights(1)= Light1
Set Lights(2)= Light2
Set Lights(3)= Light3
Set Lights(4)= Light4
Set Lights(5)= Light5
Set Lights(6)= Light6
Set Lights(7)= Light7
Set Lights(8)= Light8
Set Lights(9)= Light9
Set Lights(10)= Light10
Set Lights(11)= Light11
Set Lights(12)= Light12
Set Lights(14)= Light14
Set Lights(15)= Light15
Set Lights(16)= Light16
Set Lights(17)= Light17
Set Lights(18)= Light18
Set Lights(19)= Light19
Set Lights(20)= Light20
Set Lights(21)= Light21
Set Lights(22)= Light22
Set Lights(23)= Light23
Set Lights(24)= Light24
Set Lights(25)= Light25
Set Lights(26)= Light26
Set Lights(28)= Light28
Set Lights(30)= Light30
Set Lights(31)= Light31
Set Lights(32)= Light32
Set Lights(33)= Light33
Set Lights(34)= Light34
Set Lights(35)= Light35
Set Lights(36)= Light36
Set Lights(37)= Light37
Set Lights(38)= Light38
Set Lights(39)= Light39
Set Lights(40)= Light40
Set Lights(41)= Light41
Set Lights(42)= Light42
Set Lights(43)= Light43
Set Lights(44)= Light44
Set Lights(46)= Light46
Set Lights(47)= Light47
Set Lights(49)= Light49
Set Lights(50)= Light50
Set Lights(51)= Light51
Set Lights(52)= Light52
Set Lights(53)= Light53
Set Lights(54)= Light54
Set Lights(55)= Light55
Set Lights(56)= Light56
Set Lights(57)= Light57
Set Lights(59)= Light59
Set Lights(60)= Light60
Set Lights(62)= Light62
Set Lights(63)= Light63
Set Lights(65)= Light65
Set Lights(66)= Light66
Set Lights(67)= Light67
Set Lights(81)= Light81
Set Lights(82)= Light82
Set Lights(83)= Light83
Set Lights(97)= Light97
Set Lights(98)= Light98
Set Lights(99)= Light99

Set Lights(113)= Light113
Set Lights(114)= Light114
Set Lights(115)= Light115

'BackGlass Lights
'Set Lights(13)= Light13
'Set Lights(27)= Light27
'Set Lights(29)= Light29
'Set Lights(45)= Light45
'Set Lights(61)= Light61


'**********************************************************************************************************
'Digital Display
'**********************************************************************************************************
Dim Digits(32)

' 1st Player
Digits(0) = Array(LED10,LED11,LED12,LED13,LED14,LED15,LED16)
Digits(1) = Array(LED20,LED21,LED22,LED23,LED24,LED25,LED26)
Digits(2) = Array(LED30,LED31,LED32,LED33,LED34,LED35,LED36)
Digits(3) = Array(LED40,LED41,LED42,LED43,LED44,LED45,LED46)
Digits(4) = Array(LED50,LED51,LED52,LED53,LED54,LED55,LED56)
Digits(5) = Array(LED60,LED61,LED62,LED63,LED64,LED65,LED66)
Digits(6) = Array(LED70,LED71,LED72,LED73,LED74,LED75,LED76)

' 2nd Player
Digits(7) = Array(LED80,LED81,LED82,LED83,LED84,LED85,LED86)
Digits(8) = Array(LED90,LED91,LED92,LED93,LED94,LED95,LED96)
Digits(9) = Array(LED100,LED101,LED102,LED103,LED104,LED105,LED106)
Digits(10) = Array(LED110,LED111,LED112,LED113,LED114,LED115,LED116)
Digits(11) = Array(LED120,LED121,LED122,LED123,LED124,LED125,LED126)
Digits(12) = Array(LED130,LED131,LED132,LED133,LED134,LED135,LED136)
Digits(13) = Array(LED140,LED141,LED142,LED143,LED144,LED145,LED146)

' 3rd Player
Digits(14) = Array(LED150,LED151,LED152,LED153,LED154,LED155,LED156)
Digits(15) = Array(LED160,LED161,LED162,LED163,LED164,LED165,LED166)
Digits(16) = Array(LED170,LED171,LED172,LED173,LED174,LED175,LED176)
Digits(17) = Array(LED180,LED181,LED182,LED183,LED184,LED185,LED186)
Digits(18) = Array(LED190,LED191,LED192,LED193,LED194,LED195,LED196)
Digits(19) = Array(LED200,LED201,LED202,LED203,LED204,LED205,LED206)
Digits(20) = Array(LED210,LED211,LED212,LED213,LED214,LED215,LED216)

' 4th Player
Digits(21) = Array(LED220,LED221,LED222,LED223,LED224,LED225,LED226)
Digits(22) = Array(LED230,LED231,LED232,LED233,LED234,LED235,LED236)
Digits(23) = Array(LED240,LED241,LED242,LED243,LED244,LED245,LED246)
Digits(24) = Array(LED250,LED251,LED252,LED253,LED254,LED255,LED256)
Digits(25) = Array(LED260,LED261,LED262,LED263,LED264,LED265,LED266)
Digits(26) = Array(LED270,LED271,LED272,LED273,LED274,LED275,LED276)
Digits(27) = Array(LED280,LED281,LED282,LED283,LED284,LED285,LED286)

' Credits
Digits(28) = Array(LED4,LED2,LED6,LED7,LED5,LED1,LED3)
Digits(29) = Array(LED18,LED9,LED27,LED28,LED19,LED8,LED17)
' Balls
Digits(30) = Array(LED39,LED37,LED48,LED49,LED47,LED29,LED38)
Digits(31) = Array(LED67,LED58,LED69,LED77,LED68,LED57,LED59)

Sub DisplayTimer_Timer
	Dim ChgLED,ii,num,chg,stat,obj
	ChgLed = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
If Not IsEmpty(ChgLED) Then
		If DesktopMode = True Then
		For ii = 0 To UBound(chgLED)
			num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
			if (num < 32) then
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
'Bally Spy Hunter
'added by Inkochnito
Sub editDips
	Dim vpmDips:Set vpmDips=New cvpmDips
	With vpmDips
		.AddForm 700,400,"Spy Hunter - DIP switches"
		.AddChk 7,10,180,Array("Match feature",&H08000000)'dip 28
		.AddChk 205,10,115,Array("Credits display",&H04000000)'dip 27
		.AddFrame 2,30,190,"Maximum credits",&H03000000,Array("10 credits",0,"15 credits",&H01000000,"25 credits",&H02000000,"40 credits",&H03000000)'dip 25&26
		.AddFrame 2,106,190,"Weapon values adjustment",&H00000020,Array("weapon values are reset",0,"weapon values are recalled",&H00000020)'dip 6
		.AddFrame 2,152,190,"Completing Weapons first time",&H00000040,Array("lights 90K light",0,"scores extra ball",&H00000040)'dip 7
		.AddFrame 2,198,190,"Weapons values special limit",&H00000080,Array("only 1 special",0,"more than 1 special",&H00000080)'dip 8
		.AddFrame 2,248,190,"Game style adjustment",&H00002000,Array("bonus must be qualified",0,"spinner lights must match",&H00002000)'dip 14
		.AddFrame 2,298,190,"Drop target special adjustment",49152,Array ("special lights at 70K",0,"special lights at 60K",&H00004000,"special lights at 50K",32768,"special lights at 40K",49152)'dip 15&16
		.AddFrame 205,30,190,"Balls per game",&HC0000000,Array ("2 balls",&HC0000000,"3 balls",0,"4 balls",&H80000000,"5 balls",&H40000000)'dip 31&32
		.AddFrame 205,106,190,"Double/triple playfield values",&H00100000,Array("Missile arrow every 5th target",0,"Missile arrow every 4th target",&H00100000)'dip 21
		.AddFrame 205,152,190,"Bonus advance adjustment",&H00200000,Array("no advance",0,"drop targets advance bonus value",&H00200000)'dip 22
		.AddFrame 205,198,190,"Bonus saucer qualified",&H00400000,Array("drop targets do not qualify bonus",0,"drop targets qualify bonus",&H00400000)'dip 23
		.AddFrame 205,248,190,"Drop target values adjustment",&H00800000,Array("values reset",0,"values in memory",&H00800000)'dip 24
		'.AddFrame 205,298,190,"not used",&H10000000,Array("off",0,"on",&H10000000)'dip 29
		.AddFrame 205,298,190,"Spinner lights adjustment",&H20000000,Array("harder",0,"easy",&H20000000)'dip 30
		.AddLabel 50,400,300,20,"Set selftest position 17,18 and 19 to 03 for the best gameplay."
		.AddLabel 50,420,300,20,"After hitting OK, press F3 to reset game with new settings."
		.ViewDips
	End With
End Sub
Set vpmShowDips=GetRef("editDips")

'**********************************************************************************************************
'**********************************************************************************************************

'*****GI Lights On
dim xx
For each xx in GI:xx.State = 1: Next

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
	vpmTimer.PulseSw 1
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
	vpmTimer.PulseSw 13
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


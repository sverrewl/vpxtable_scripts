Option Explicit
Randomize

' Thalamus 2018-07-20
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' No special SSF tweaks yet.


On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="dvlsdre",UseSolenoids=2,UseLamps=1,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff", SCoin="coin"

LoadVPM "01560000","SYS80.VBS",3.2
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

SolCallback(1) = "DTT.SolDropUp"
SolCallback(2) = "bsKicker.SolOut"
SolCallback(3) = "bsHole.SolOut"
SolCallback(4) = "SolBallSaver"
SolCallback(5) = "DTL.SolDropUp"
SolCallback(6) = "DTR.SolDropUp"
SolCallback(8) = "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(9)  = "bsTrough.SolIn"
SolCallback(10) = "Sol10"

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

Dim BallSaverActive,FlipperActive
Sub SolBallSaver(Enabled)
	if enabled then BallSaverActive = True
End Sub

Sub Sol10(Enabled)
	FlipperActive = enabled
	if not Flipperactive then
		LeftFlipper.RotateToStart
		RightFlipper.RotateToStart
		RightFlipper1.RotateToStart
	end if
End Sub


'*****GI Lights On
dim xx
For each xx in GI:xx.State = 1: Next

'Primitive Flipper
Sub FlipperTimer_Timer
	FlipperT1.roty = LeftFlipper.currentangle  + 237
	FlipperT5.roty = RightFlipper.currentangle + 120
	FlipperT2.roty = RightFlipper1.currentangle + 180
End Sub

'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************

Dim bsTrough, bsKicker, bsHole, DTT, DTL, DTR

Sub Table1_Init
	vpmInit Me
	On Error Resume Next
		With Controller
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
		.SplashInfoLine = "Devil's Dare"&chr(13)&"You Suck"
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

    vpmNudge.TiltSwitch=57
    vpmNudge.Sensitivity=5
	vpmNudge.TiltObj = Array(Bumper1,Bumper2,Bumper3,LeftslingShot,RightslingShot) 
    
    Set bsTrough=New cvpmBallStack
		bsTrough.InitSw 55,50,99,51,0,0,0,0
        bsTrough.InitKick BallRelease,90,5
        bsTrough.InitExitSnd  SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
        bsTrough.Balls=3

    Set bsKicker=New cvpmBallStack       
        bsKicker.InitSaucer Kicker,5,0,22
        bsKicker.InitExitSnd  SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)

    Set bsHole=New cvpmBallStack       
        bsHole.InitSaucer Hole,6,180,5
        bsHole.InitExitSnd  SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)

	set DTT = new cvpmDropTarget
		DTT.InitDrop Array(sw4,sw14,sw24), Array(4,14,24)
		DTT.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

	set DTL = new cvpmDropTarget
		DTL.InitDrop Array(sw0,sw10,sw20,sw30,sw40), Array(0,10,20,30,40)
		DTL.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

	set DTR = new cvpmDropTarget
		DTR.InitDrop Array(sw1,sw11,sw21,sw31,sw41), Array(1,11,21,31,41)
		DTR.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

	BallSaverActive = False
	BallSaveKicker.pullback

End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)
	If KeyDownHandler(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Pullback:playsound"plungerpull"
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
	If KeyUpHandler(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Fire:PlaySound"plunger"
End Sub

'**********************************************************************************************************

 ' Drain hole
Sub Drain_Hit : playsound"drain" : bsTrough.addball me : End Sub
Sub Kicker_Hit : bsKicker.addball me : playsound "popper_ball" : End Sub
Sub Hole_Hit : bsHole.addball me : playsound "popper_ball" : End Sub

'Bumpers
Sub Bumper1_Hit : vpmTimer.PulseSw(34) : playsound SoundFX("fx_bumper1",DOFContactors): End Sub
Sub Bumper2_Hit : vpmTimer.PulseSw(34) : playsound SoundFX("fx_bumper1",DOFContactors): End Sub
Sub Bumper3_Hit : vpmTimer.PulseSw(34) : playsound SoundFX("fx_bumper1",DOFContactors): End Sub

'Drop Targets
Sub sw0_Dropped:DTL.Hit 1:End Sub
Sub sw10_Dropped:DTL.Hit 2:End Sub
Sub sw20_Dropped:DTL.Hit 3:End Sub
Sub sw30_Dropped:DTL.Hit 4:End Sub
Sub sw40_Dropped:DTL.Hit 5:End Sub

Sub sw1_Dropped:DTR.Hit 1:End Sub
Sub sw11_Dropped:DTR.Hit 2:End Sub
Sub sw21_Dropped:DTR.Hit 3:End Sub
Sub sw31_Dropped:DTR.Hit 4:End Sub
Sub sw41_Dropped:DTR.Hit 5:End Sub

Sub sw4_Dropped:DTT.Hit 1:End Sub
Sub sw14_Dropped:DTT.Hit 2:End Sub
Sub sw24_Dropped:DTT.Hit 3:End Sub

'Stand Up Targets
Sub sw2_Hit : vpmTimer.PulseSw(2) : End Sub
Sub sw12_Hit : vpmTimer.PulseSw(12) : End Sub
Sub sw22_Hit : vpmTimer.PulseSw(22) : End Sub
Sub sw32_Hit : vpmTimer.PulseSw(32) : End Sub

Sub sw45_Hit : vpmTimer.PulseSw(45) : End Sub
Sub sw46_Hit : vpmTimer.PulseSw(46) : End Sub

Sub sw3_Hit : vpmTimer.PulseSw(3): End Sub
Sub sw13_Hit : vpmTimer.PulseSw(13) : End Sub
Sub sw23_Hit : vpmTimer.PulseSw(23) : End Sub
Sub sw33_Hit : vpmTimer.PulseSw(33) : End Sub

Sub sw25_Hit : vpmTimer.PulseSw(25) : End Sub
Sub sw25a_Hit : vpmTimer.PulseSw(25) : End Sub

'Wire Triggers
Sub sw36_Hit:Controller.Switch(36)=1 : playsound"rollover" : End Sub 
Sub sw36_unHit:Controller.Switch(36)=0:End Sub
Sub sw54_Hit:Controller.Switch(54)=1 : playsound"rollover" : End Sub 
Sub sw54_unHit:Controller.Switch(54)=0:End Sub
Sub sw53_Hit:Controller.Switch(53)=1 : playsound"rollover" : End Sub 
Sub sw53_unHit:Controller.Switch(53)=0:End Sub
Sub sw42_Hit:Controller.Switch(42)=1 : playsound"rollover" : End Sub 
Sub sw42_unHit:Controller.Switch(42)=0:End Sub
Sub sw43_Hit:Controller.Switch(43)=1 : playsound"rollover" : End Sub 
Sub sw43_unHit:Controller.Switch(43)=0:End Sub
Sub sw44_Hit:Controller.Switch(44)=1 : playsound"rollover" : End Sub 
Sub sw44_unHit:Controller.Switch(44)=0:End Sub

'Gate Trigger
Sub sw16_Hit : vpmTimer.PulseSw(16) : End Sub

'Spinners
Sub sw26_Spin:vpmTimer.PulseSw 26 : playsound"fx_spinner" : End Sub
Sub sw15_Spin:vpmTimer.PulseSw 15 : playsound"fx_spinner" : End Sub

'Scoring Rubbers
Sub sw52a_Slingshot:vpmTimer.PulseSw 52 : playsound"rubber_hit_3" : End Sub
Sub sw52b_Slingshot:vpmTimer.PulseSw 52 : playsound"rubber_hit_3" : End Sub
Sub sw52c_Slingshot:vpmTimer.PulseSw 52 : playsound"rubber_hit_3" : End Sub
Sub sw35a_Slingshot:vpmTimer.PulseSw 35 : playsound"rubber_hit_3" : End Sub
Sub sw35b_Slingshot:vpmTimer.PulseSw 35 : playsound"rubber_hit_3" : End Sub

'**********************************************************************************************************
' Map Lights to Array
'**********************************************************************************************************

set lights(3) = lamp3		'shoot again pf
set lights(4) = lamp4
set lights(5) = lamp5
set lights(6) = lamp6
set lights(7) = lamp7
set lights(12) = l12
set lights(15) = lamp15
lights(16) = array(L16a,L16b,L16c) 'Bumper Lights
set lights(17) = lamp17
set lights(18) = lamp18
set lights(19) = lamp19
set lights(20) = lamp20
set lights(21) = lamp21
set lights(22) = lamp22
set lights(23) = lamp23
set lights(24) = lamp24
set lights(25) = lamp25
set lights(26) = lamp26
set lights(27) = lamp27
set lights(28) = lamp28
set lights(29) = lamp29
set lights(30) = lamp30
set lights(31) = lamp31
set lights(32) = lamp32
set lights(33) = lamp33
set lights(34) = lamp34
set lights(35) = lamp35
set lights(36) = lamp36
set lights(37) = lamp37
set lights(38) = lamp38
set lights(39) = lamp39
lights(40) = array(lamp40a,lamp40b)
set lights(41) = lamp41
set lights(42) = lamp42
set lights(43) = lamp43
set lights(44) = lamp44
set lights(45) = lamp45
set lights(46) = lamp46
set lights(47) = lamp47

'BackGlass
'set lights(0) = l0		'ballinplay
'set lights(1) = l1		'tilt
'set lights(2) = l2	'shoot again lb = pf
'set lights(8) = l8		'??? lb
'set lights(9) = l9		'multi-mode
'set lights(10) = l10	'highscore
'set lights(11) = l11	'gameover
'set lights(12) = l12	'ball release
'set lights(13) = l13	'multi-bonus
'set lights(14) = l14	'???


'Lamp Callback for Ball Release and Ballsaver
Sub LampTimer_Timer
	If (L12.State <> 0) and (not BallRelease.timerenabled) Then
		BallRelease.timerenabled = True
		bsTrough.ExitSol_On
	End if

	if BallSaverActive then
		BallSaveA.state = Lightstateon
		BallSaveB.state = Lightstateon	
		xLBallSaver.state = Lightstateon	
	else	
		BallSaveA.state = Lightstateoff
		BallSaveB.state = Lightstateoff
		xLBallSaver.state = Lightstateoff
	end if



	xL50.state = controller.switch(50)
	xL51.state = controller.switch(51)
	xL55.state = controller.switch(55)

	if controller.switch(51) then				'deactivate on last Ball drain
		BallSaverActive = False
	end if	
End sub

Sub BallSaveTimer_Timer
	BallSaveTimer.enabled = False
	BallSaverActive = False
End Sub

Sub BallRelease_timer
	BallRelease.timerenabled = False
End Sub

'**********************************************************************************************************
' Backglass Light Displays
'**********************************************************************************************************

Dim Digits(44)
Digits(0)=Array(a00,a01,a02,a03,a04,a05,a06,n,a08)
Digits(1)=Array(a10,a11,a12,a13,a14,a15,a16,n,a18)
Digits(2)=Array(a20,a21,a22,a23,a24,a25,a26,n,a28)
Digits(3)=Array(a30,a31,a32,a33,a34,a35,a36,n,a38)
Digits(4)=Array(a40,a41,a42,a43,a44,a45,a46,n,a48)
Digits(5)=Array(a50,a51,a52,a53,a54,a55,a56,n,a58)
Digits(6)=Array(b00,b01,b02,b03,b04,b05,b06,n,b08)

Digits(7)=Array(b10,b11,b12,b13,b14,b15,b16,n,b18)
Digits(8)=Array(b20,b21,b22,b23,b24,b25,b26,n,b28)
Digits(9)=Array(b30,b31,b32,b33,b34,b35,b36,n,b38)
Digits(10)=Array(b40,b41,b42,b43,b44,b45,b46,n,b48)
Digits(11)=Array(b50,b51,b52,b53,b54,b55,b56,n,b58)
Digits(12)=Array(c00,c01,c02,c03,c04,c05,c06,n,c08)
Digits(13)=Array(c10,c11,c12,c13,c14,c15,c16,n,c18)

Digits(14)=Array(c20,c21,c22,c23,c24,c25,c26,n,c28)
Digits(15)=Array(c30,c31,c32,c33,c34,c35,c36,n,c38)
Digits(16)=Array(c40,c41,c42,c43,c44,c45,c46,n,c48)
Digits(17)=Array(c50,c51,c52,c53,c54,c55,c56,n,c58)
Digits(18)=Array(d00,d01,d02,d03,d04,d05,d06,n,d08)
Digits(19)=Array(d10,d11,d12,d13,d14,d15,d16,n,d18)
Digits(20)=Array(d20,d21,d22,d23,d24,d25,d26,n,d28)

Digits(21)=Array(d30,d31,d32,d33,d34,d35,d36,n,d38)
Digits(22)=Array(d40,d41,d42,d43,d44,d45,d46,n,d48)
Digits(23)=Array(d50,d51,d52,d53,d54,d55,d56,n,d58)
Digits(24)=Array(e00,e01,e02,e03,e04,e05,e06,n,e08)
Digits(25)=Array(e10,e11,e12,e13,e14,e15,e16,n,e18)
Digits(26)=Array(f00,f01,f02,f03,f04,f05,f06,n,f08)
Digits(27)=Array(f10,f11,f12,f13,f14,f15,f16,n,f18)

'credit -- Ball In Play
Digits(28) = Array(e2,e3,e7,e4,e5,e1,e6,n,e23)
Digits(29) = Array(e9,e17,e22,e19,e20,e8,e21,n,e24)
Digits(30) = Array(f2,f3,f7,f4,f5,f1,f6,n,f23)
Digits(31) = Array(f9,f17,f22,f19,f20,f8,f21,n,f24)


Digits(32)=Array(d2,d3,d7,d4,d5,d1,d6,n,d62)
Digits(33)=Array(d9,d17,d37,d19,d27,d8,d29,n,d63)
Digits(34)=Array(d47,d49,d61,d57,d59,d39,d60,n,d64)
Digits(35)=Array(e26,e27,e31,e28,e29,e25,e30,n,e39)
Digits(36)=Array(e33,e34,e38,e35,e36,e32,e37,n,e40)
Digits(37)=Array(f26,f27,f31,f28,f29,f25,f30,n,f39)


Digits(38)=Array(f33,f34,f38,f35,f36,f32,f37,n,f40)
Digits(39)=Array(d82,d83,d87,d84,d85,d81,d86,n,d85)
Digits(40)=Array(d89,d90,d94,d91,d92,d88,d93,n,d96)
Digits(41)=Array(e58,e59,e63,e60,e61,e57,e62,n,e71)
Digits(42)=Array(e65,e66,e70,e67,e68,e64,e69,n,e72)
Digits(43)=Array(f50,f51,f55,f52,f53,f49,f54,n,f56)


Sub DisplayTimer_Timer
	Dim ChgLED,ii,num,chg,stat,obj
	ChgLed = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
If Not IsEmpty(ChgLED) Then
		If DesktopMode = True Then
		For ii = 0 To UBound(chgLED)
			num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
			if (num < 44) then
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
'Gottlieb Devils Dare
'added by Inkochnito
Sub editDips
	Dim vpmDips : Set vpmDips = New cvpmDips
	With vpmDips
		.AddForm 700,400,"Devils Dare - DIP switches"
		.AddChk 7,10,100,Array("Match feature",&H02000000)'dip 26
		.AddChkExtra 110,10,100,Array("Speech",&H0020)'SS-board dip 6
		.AddChkExtra 220,10,110,Array("Background sound",&H0010)'SS-board dip 5
		.AddFrame 2,30,190,"Maximum credits",49152,Array("8 credits",0,"10 credits",32768,"15 credits",&H00004000,"25 credits",49152)'dip 15&16
		.AddFrame 2,106,190,"Coin chute 1 and 2 control",&H00002000,Array("seperate",0,"same",&H00002000)'dip 14
		.AddFrame 2,152,190,"Playfield special",&H00200000,Array("replay",0,"extra ball",&H00200000)'dip 22
		.AddFrame 2,198,190,"3rd coin chute credits control",&H20000000,Array("no effect",0,"add 9",&H20000000)'dip 30
		.AddFrame 2,244,190,"High score to date awards",&H00C00000,Array("not displayed and no award",0,"displayed and no award",&H00800000,"displayed and 2 credits",&H00400000,"displayed and 3 credits",&H00C00000)'dip 23&24
		.AddFrameExtra 205,30,190,"Attract Sound",&H000C,Array("off",0,"every 10 seconds",&H0004,"every 2 minutes",&H0008,"every 4 minutes",&H000C)'sounddip 3&4
		.AddFrame 205,106,190,"Balls per game",&H01000000,Array("5 balls",0,"3 balls",&H01000000)'dip 25
		.AddFrame 205,152,190,"Replay limit",&H04000000,Array("no limit",0,"one per ball",&H04000000)'dip 27
		.AddFrame 205,198,190,"Novelty mode",&H08000000,Array("normal game mode",0,"50,000 points per special/extra ball",&H08000000)'dip 28
		.AddFrame 205,244,190,"Game mode",&H10000000,Array("replay",0,"extra ball",&H10000000)'dip 29
		.AddLabel 50,325,300,20,"After hitting OK, press F3 to reset game with new settings."
	End With
	Dim extra
	extra = Controller.Dip(4) + Controller.Dip(5)*256
	extra = vpmDips.ViewDipsExtra(extra)
	Controller.Dip(4) = extra And 255
	Controller.Dip(5) = (extra And 65280)\256 And 255
End Sub
Set vpmShowDips = GetRef("editDips")


' *********************************************************************
' *********************************************************************

					'Start of VPX call back Functions

' *********************************************************************
' *********************************************************************


'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
	vpmTimer.PulseSw 52
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
	vpmTimer.PulseSw 52
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

' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


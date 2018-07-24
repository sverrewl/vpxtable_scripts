Option Explicit
Randomize

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' No special SSF tweaks yet.
' , AudioFade(ActiveBall)

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="totem",UseSolenoids=1,UseLamps=0,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff",SCoin="coin"

LoadVPM "01001100", "gts1.vbs", 3.02
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
SolCallback(1)= "bsTrough.SolOut"
SolCallback(2)=  "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(6)= "bsSaucer.SolOut"
'SolCallBack(7)= "SolVReset"
SolCallback(8)= "dtDropL.SolDropUp"

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
         PlaySound SoundFX("fx_Flipperup",DOFContactors):RightFlipper.RotateToEnd
     Else
         PlaySound SoundFX("fx_Flipperdown",DOFContactors):RightFlipper.RotateToStart
     End If
End Sub
'**********************************************************************************************************

'Solenoid Controlled toys
'**********************************************************************************************************

'Primitive Flipper Code
Sub FlipperTimer_Timer
	FlipperT1.roty = LeftFlipper.currentangle  + 237
	FlipperT5.roty = RightFlipper.currentangle + 122
End Sub


'*****GI Lights On
dim xx
For each xx in GI:xx.State = 1: Next

'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************
Dim bsTrough, bsSaucer, dtDropL

Sub Table1_Init
	vpmInit Me
	On Error Resume Next
		With Controller
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
		.SplashInfoLine = "Totem (Gotlieb)"&chr(13)&"You Suck"
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

	vpmNudge.TiltSwitch = 4
	vpmNudge.Sensitivity = 5
	vpmNudge.TiltObj = Array(Bumper1,Bumper2,LeftSlingshot,RightSlingshot)

 	Set bsTrough = New cvpmBallStack
		bsTrough.initnotrough ballrelease, 66, 55, 6
		bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)

  	Set dtDropL=New cvpmDropTarget'
		dtDropL.InitDrop Array(sw20,sw21,sw24,sw30,sw31,sw34),Array(20,21,24,30,31,34)
		dtDropL.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

	Set bsSaucer=New cvpmBallstack
		bsSaucer.InitSaucer sw41,41,234+rnd(1)*3,5+rnd(1)*3
		bsSaucer.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)

End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)
	If KeyDownHandler(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Pullback:playsound"plungerpull"
	If keycode=AddCreditKey then playsound "coin": vpmTimer.pulseSW (swCoin1): end if
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
	If KeyUpHandler(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Fire:PlaySound"plunger"
End Sub

'**********************************************************************************************************

'Switches
'**********************************************************************************************************
'**********************************************************************************************************
'Drain hole and kicker
Sub Drain_Hit:playsound"drain":bsTrough.addball me:End Sub
Sub sw41_Hit():bsSaucer.Addball 0 : playsound "popper_ball": End Sub

'Bumpers
Sub Bumper1_Hit : vpmTimer.PulseSw(14) : playsound SoundFX("fx_bumper1",DOFContactors): End Sub
Sub Bumper2_Hit : vpmTimer.PulseSw(14) : playsound SoundFX("fx_bumper1",DOFContactors): End Sub

'Stand Up Targets
 Sub sw62a_Hit():vpmtimer.pulsesw 62:End Sub
 Sub sw62b_Hit():vpmtimer.pulsesw 62:End Sub
 Sub sw62c_Hit():vpmtimer.pulsesw 62:End Sub
 Sub sw64_Hit():vpmtimer.pulsesw 64:End Sub

'Drop Targets
Sub sw20_Dropped:dtDropL.Hit 1:End Sub
Sub sw21_Dropped:dtDropL.Hit 2:End Sub
Sub sw24_Dropped:dtDropL.Hit 3:End Sub
Sub sw30_Dropped:dtDropL.Hit 4:End Sub
Sub sw31_Dropped:dtDropL.Hit 5:End Sub
Sub sw34_Dropped:dtDropL.Hit 6:End Sub

'Star Triggers
Sub sw40_Hit():Controller.Switch(40)=1 : playsound"rollover" :End Sub
Sub sw40_UnHit:Controller.Switch(40)=0:End Sub
Sub sw50_Hit():Controller.Switch(50)=1 : playsound"rollover" : End Sub
Sub sw50_UnHit:Controller.Switch(50)=0:End Sub
Sub sw40a_Hit():Controller.Switch(40)=1 : playsound"rollover":End Sub
Sub sw40a_UnHit:Controller.Switch(40)=0:End Sub

'Scoring Gate
Sub Gate1_Hit():Controller.Switch(61)=1:End Sub
Sub Gate1_UnHit:Controller.Switch(61)=0:End Sub

'Scoring Rubbers
Sub sw60a_Slingshot():vpmtimer.pulsesw 60: playsound"slingshot" : End Sub
Sub sw60b_Slingshot():vpmtimer.pulsesw 60: playsound"slingshot" : End Sub

'Wire Triggers
 Sub sw10_Hit():Controller.Switch(10)=1 : playsound"rollover" : End Sub
 Sub sw10_UnHit:Controller.Switch(10)=0:End Sub
 Sub sw11_Hit():Controller.Switch(11)=1 : playsound"rollover" : End Sub
 Sub sw11_UnHit:Controller.Switch(11)=0:End Sub
 Sub sw51_Hit():Controller.Switch(51)=1 : playsound"rollover":End Sub
 Sub sw51_UnHit:Controller.Switch(51)=0:End Sub
 Sub sw51a_Hit():Controller.Switch(51)=1 : playsound"rollover":End Sub
 Sub sw51a_UnHit:Controller.Switch(51)=0:End Sub
Sub sw70_Hit():Controller.Switch(70)=1 : playsound"rollover" : End Sub
Sub sw70_UnHit:Controller.Switch(70)=0:End Sub
Sub sw71_Hit():Controller.Switch(71)=1 : playsound"rollover" : End Sub
Sub sw71_UnHit:Controller.Switch(71)=0:End Sub
Sub sw72_Hit():Controller.Switch(72)=1 : playsound"rollover" : End Sub
Sub sw72_UnHit:Controller.Switch(72)=0:End Sub
Sub sw74_Hit():Controller.Switch(74)=1 : playsound"rollover" : End Sub
Sub sw74_UnHit:Controller.Switch(74)=0:End Sub

'Moving Target
 Sub sw42_Hit:vpmTimer.PulseSw 42:PlaySound "target"
 PrimTarget1.visible = 0:PrimTarget2.visible = 1:PrimTarget3.visible = 0:PrimTarget4.visible = 0:
 End Sub

 Sub sw52_Hit:vpmTimer.PulseSw 52
 PrimTarget1.visible = 0:PrimTarget2.visible = 0:PrimTarget3.visible = 1:PrimTarget4.visible = 0:
 End Sub

 Sub sw51b_Hit:vpmTimer.PulseSw 51
 PrimTarget1.visible = 0:PrimTarget2.visible = 0:PrimTarget3.visible = 0:PrimTarget4.visible = 1:
 End Sub

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
If DesktopMode = True Then
	FadeIReel 1, EMReelGO
end if
If DesktopMode = True Then
	FadeReel 2, EMReelTilt
end if
If DesktopMode = True Then
	FadeReel 3, EMReelHTD
end if
	NFadeLm 4, l4
If DesktopMode = True Then
	FadeReel 4, EMReelSPSA
end if
	NFadeL 5, l5
	NFadeL 6, l6
	NFadeL 7, l7
	NFadeL 8, l8
	NFadeL 9, l9
	NFadeL 10, l10
	NFadeL 11, l11
	NFadeL 12, l12
	NFadeL 13, l13
	NFadeL 14, l14
	NFadeL 15, l15
	NFadeL 16, l16
	NFadeL 17, l17
	NFadeL 18, l18
	'NFadeL 19, l19
	'NFadeL 20, l20
	NFadeL 21, l21
	NFadeL 22, l22
	NFadeL 23, l23
	NFadeL 24, l24
	NFadeL 25, l25
	NFadeL 26, l26
	NFadeL 27, l27
	NFadeL 28, l28
	NFadeL 29, l29
	NFadeL 30, l30
	NFadeL 31, l31
	NFadeL 32, l32
	NFadeL 33, l33
	NFadeL 34, l34
	NFadeL 35, l35
	NFadeL 36, l36


	'FadeReel lBallInPlay, EMReelBIP
	'FadeReel lNumberToMatch, EMReelNTM
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

 'Reels
Sub FadeReel(nr, reel)
    Select Case FadingLevel(nr)
        Case 2:FadingLevel(nr) = 0
        Case 3:FadingLevel(nr) = 2
        Case 4:reel.Visible = 0:FadingLevel(nr) = 3
        Case 5:reel.Visible = 1:FadingLevel(nr) = 1
    End Select
End Sub

 'Inverted Reels
Sub FadeIReel(nr, reel)
    Select Case FadingLevel(nr)
        Case 2:FadingLevel(nr) = 0
        Case 3:FadingLevel(nr) = 2
        Case 4:reel.Visible = 1:FadingLevel(nr) = 3
        Case 5:reel.Visible = 0:FadingLevel(nr) = 1
    End Select
End Sub

'**********************************************************************************************************
 '**********************************************************************************************************
' Backglass Light Displays (7 digit 7 segment displays)
Dim Digits(28)
Digits(0)=Array(a00,a01,a02,a03,a04,a05,a06,n,a08)
Digits(1)=Array(a10,a11,a12,a13,a14,a15,a16,n,a18)
Digits(2)=Array(a20,a21,a22,a23,a24,a25,a26,n,a28)
Digits(3)=Array(a30,a31,a32,a33,a34,a35,a36,n,a38)
Digits(4)=Array(a40,a41,a42,a43,a44,a45,a46,n,a48)
Digits(5)=Array(a50,a51,a52,a53,a54,a55,a56,n,a58)


Digits(6)=Array(b20,b21,b22,b23,b24,b25,b26,n,b28)
Digits(7)=Array(b30,b31,b32,b33,b34,b35,b36,n,b38)
Digits(8)=Array(b40,b41,b42,b43,b44,b45,b46,n,b48)
Digits(9)=Array(b50,b51,b52,b53,b54,b55,b56,n,b58)
Digits(10)=Array(c00,c01,c02,c03,c04,c05,c06,n,c08)
Digits(11)=Array(c10,c11,c12,c13,c14,c15,c16,n,c18)

Digits(12)=Array(c20,c21,c22,c23,c24,c25,c26,n,c28)
Digits(13)=Array(c30,c31,c32,c33,c34,c35,c36,n,c38)
Digits(14)=Array(c40,c41,c42,c43,c44,c45,c46,n,c48)
Digits(15)=Array(c50,c51,c52,c53,c54,c55,c56,n,c58)
Digits(16)=Array(d00,d01,d02,d03,d04,d05,d06,n,d08)
Digits(17)=Array(d10,d11,d12,d13,d14,d15,d16,n,d18)

Digits(18)=Array(d40,d41,d42,d43,d44,d45,d46,n,d48)
Digits(19)=Array(d50,d51,d52,d53,d54,d55,d56,n,d58)
Digits(20)=Array(e00,e01,e02,e03,e04,e05,e06,n,e08)
Digits(21)=Array(e10,e11,e12,e13,e14,e15,e16,n,e18)
Digits(22)=Array(f00,f01,f02,f03,f04,f05,f06,n,f08)
Digits(23)=Array(f10,f11,f12,f13,f14,f15,f16,n,f18)

'credit -- Ball In Play
Digits(24) = Array(e2,e3,e7,e4,e5,e1,e6,n,e23)
Digits(25) = Array(e9,e17,e22,e19,e20,e8,e21,n,e24)
Digits(26) = Array(f2,f3,f7,f4,f5,f1,f6,n,f23)
Digits(27) = Array(f9,f17,f22,f19,f20,f8,f21,n,f24)


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
			end if
		next
		end if
end if
End Sub

'**********************************************************************************************************
'**********************************************************************************************************
Sub editDips
   Dim vpmDips : Set vpmDips = New cvpmDips
   With vpmDips
  .AddForm 700,400,"System 1 - DIP switches"
  .AddFrame 0,0,190,"Coin chutecontrol",&H00040000,Array("seperate",0,"same",&H00040000)'dip 19
  .AddFrame 0,46,190,"Game mode",&H00000400,Array("extraball",0,"replay",&H00000400)'dip 11
  .AddFrame 0,92,190,"High game to date awards",&H00200000,Array("noaward",0,"3 replays",&H00200000)'dip 22
  .AddFrame 0,138,190,"Balls per game",&H00000100,Array("5 balls",0,"3balls",&H00000100)'dip 9
  .AddFrame 0,184,190,"Tilt effect",&H00000800,Array("game over",0,"ball in play only",&H00000800)'dip 12
  .AddFrame 205,0,190,"Maximum credits",&H00030000,Array("5 credits",0,"8 credits",&H00020000,"10 credits",&H00010000,"15 credits",&H00030000)'dip 17&18
  .AddFrame 205,76,190,"Soundsettings",&H80000000,Array("sounds",0,"tones",&H80000000)'dip 32
  .AddFrame 205,122,190,"Attract tune",&H10000000,Array("no attract tune",0,"attract tune played every 6 minutes",&H10000000)'dip 29
  .AddChk 205,175,190,Array("Match feature",&H00000200)'dip 10
  .AddChk 205,190,190,Array("Credits displayed",&H00001000)'dip 13
  .AddChk 205,205,190,Array("Play credit button tune",&H00002000)'dip 14
  .AddChk 205,220,190,Array("Play tones when scoring",&H00080000)'dip 20
  .AddChk 205,235,190,Array("Play coin switch tune",&H00400000)'dip 23
  .AddChk 205,250,190,Array("High game to date displayed",&H00100000)'dip 21
  .AddLabel 50,280,300,20,"After hitting OK, press F3 to reset game with new settings."
  .ViewDips
  End With
 End Sub
 Set vpmShowDips = GetRef("editDips")

'**********************************************************************************************************
'**********************************************************************************************************
'VPX
'**********************************************************************************************************
'**********************************************************************************************************


'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
	vpmTimer.PulseSw 60
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
	vpmTimer.PulseSw 60
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

'Pop Bumpers

Sub Bumper1_Hit
	vpmTimer.PulseSw 50
	DOF 101,2
	playsound SoundFX("fx_bumper4",DOFContactors)
	B1L1.State = 1:B1L2. State = 1
	me.timerenabled=1
End Sub

Sub Bumper1_timer
	BumperTimerRing1.Enabled=0
	if bumperring1.transz>-36 then 	BumperRing1.transz=BumperRing1.transz-4
	if BumperRing1.transz=-36 then
		BumperTimerRing1.enabled=1
		me.timerenabled=0
	end if
End Sub

Sub BumperTimerRing1_timer
	if bumperring1.transz<0 then BumperRing1.transz=BumperRing1.transz+4
	If BumperRing1.transz=0 then
		B1L1.State = 0:B1L2. State = 0
		BumperTimerRing1.enabled=0
	end if
End sub

Sub Bumper2_Hit
	vpmTimer.PulseSw 50
	DOF 102,2
	playsound SoundFX("fx_bumper3",DOFContactors)
	B2L1.State = 1:B2L2. State = 1
	me.timerenabled=1
End Sub

Sub Bumper2_timer
	BumperTimerRing2.enabled=0
	if BumperRing2.transz>-36 then 	BumperRing2.transz=BumperRing2.transz-4
	if BumperRing2.transz=-36 then
		BumperTimerRing2.enabled=1
		me.timerenabled=0
	end if
End Sub

Sub BumperTimerRing2_timer
	if BumperRing2.transz<0 then BumperRing2.transz=BumperRing2.transz+4
	If BumperRing2.transz=0 then
		B2L1.State = 0:B2L2. State = 0
		BumperTimerRing2.enabled=0
	End If
End sub

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


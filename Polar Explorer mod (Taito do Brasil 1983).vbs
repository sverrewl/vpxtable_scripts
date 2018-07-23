Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="polar",UseSolenoids=1,UseLamps=1,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff",SFlipperOn="",SFlipperOff="", SCoin="coin"

LoadVPM "01200100","Taito.VBS",3.1
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

SolCallback(1)		= "bsTrough.SolOut"
SolCallback(2)		= "BandE.SolDropUp"
SolCallback(3)		= "BandD.SolDropUp"
SolCallback(4)		= "BandS.SolDropUp"
SolCallback(5)		= "bsSaucer.SolOut"
SolCallback(6)		= "TopFlipperRelay"
SolCallback(18)		= "vpmNudge.SolGameOn"

SolCallback(sLLFlipper) = "SolLeftFlipper"	 			'Sol2
SolCallback(sLRFlipper) = "SolRightFlipper"	 			'Sol3

Dim TopFlipperActive
Sub TopFlipperRelay(Enabled)
	if enabled then
		TopFlipperActive = True
	else
		TopFlipperActive = False
		LeftFlipper1.Rotatetostart
		RightFlipper1.Rotatetostart
	end if
End Sub

Sub SolLeftFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFX("fx_Flipperup",DOFContactors):LeftFlipper.RotateToEnd
         if TopFlipperActive then LeftFlipper1.RotateToEnd
     Else
         PlaySound SoundFX("fx_Flipperdown",DOFContactors):LeftFlipper.RotateToStart:LeftFlipper1.RotateToStart
     End If
  End Sub
  
Sub SolRightFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFX("fx_Flipperup",DOFContactors):RightFlipper.RotateToEnd
	     if TopFlipperActive then RightFlipper1.RotateToEnd
     Else
         PlaySound SoundFX("fx_Flipperdown",DOFContactors):RightFlipper.RotateToStart:RightFlipper1.RotateToStart
     End If
End Sub

'**********************************************************************************************************

'Solenoid Controlled toys
'**********************************************************************************************************

'*****GI Lights On
dim xx
For each xx in GI:xx.State = 1: Next

'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************

Dim bsTrough, bsSaucer, BandD, BandE, BandS

Sub Table1_Init
	vpmInit Me
	On Error Resume Next
		With Controller
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
		.SplashInfoLine = "Memory Lane -- Stern, 1978"&chr(13)&"You Suck"
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

	PinMAMETimer.Interval=PinMAMEInterval
	PinMAMETimer.Enabled=1

	vpmNudge.TiltSwitch=30
	vpmNudge.Sensitivity=5
	vpmNudge.TiltObj=Array(LeftSlingshot,RightSlingshot,Bumper1,Bumper2)

	Set bsTrough=New cvpmBallStack
		bsTrough.InitSw 0,1,0,0,0,0,0,0
		bsTrough.InitKick BallRelease,80,20
		bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
		bsTrough.Balls=1

	Set bsSaucer=New cvpmBallStack
		bsSaucer.InitSaucer sw2,2,0+Int(rnd(1))*3,15+Int(rnd(1))*5
		bsSaucer.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
		bsSaucer.KickForceVar=2

	Set BandD=New cvpmDropTarget
		BandD.InitDrop Array(sw12,sw22,sw32),Array(12,22,32)
		BandD.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

	Set BandE=New cvpmDropTarget
		BandE.InitDrop Array(sw11,sw21,sw31,sw41),Array(11,21,31,41)
		BandE.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

	Set BandS=New cvpmDropTarget
		BandS.InitDrop Array(sw4,sw14,sw24,sw34,sw44),Array(4,14,24,34,44)
		BandS.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

	Controller.Switch(60)=1
	Controller.Switch(70)=1

	TopFlipperActive = False

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

 ' Drain hole and kickers
Sub Drain_Hit:bsTrough.addball me : playsound"drain" : End Sub
Sub sw2_Hit:bsSaucer.AddBall 0 : playsound "popper_ball": End Sub

'Drop Targets
 Sub Sw12_Dropped:BandD.Hit 1 :End Sub  
 Sub Sw22_Dropped:BandD.Hit 2 :End Sub  
 Sub Sw32_Dropped:BandD.Hit 3 :End Sub

 Sub Sw11_Dropped:BandE.Hit 1 :End Sub  
 Sub Sw21_Dropped:BandE.Hit 2 :End Sub  
 Sub Sw31_Dropped:BandE.Hit 3 :End Sub
 Sub Sw41_Dropped:BandE.Hit 4 :End Sub

 Sub Sw4_Dropped:BandS.Hit 1 :End Sub  
 Sub Sw14_Dropped:BandS.Hit 2 :End Sub  
 Sub Sw24_Dropped:BandS.Hit 3 :End Sub
 Sub Sw34_Dropped:BandS.Hit 4 :End Sub
 Sub Sw44_Dropped:BandS.Hit 5 :End Sub

 'Stand Up Targets
Sub sw5_hit:vpmTimer.pulseSw 5 : End Sub
Sub sw15_hit:vpmTimer.pulseSw 15 : End Sub 
Sub sw25_hit:vpmTimer.pulseSw 25 : End Sub
Sub sw35_hit:vpmTimer.pulseSw 35 : End Sub
Sub sw45_hit:vpmTimer.pulseSw 45 : End Sub 
Sub sw63_hit:vpmTimer.pulseSw 63 : End Sub
Sub sw73_hit:vpmTimer.pulseSw 73 : End Sub 

'Wire Triggers
Sub SW42_Hit:Controller.Switch(42)=1 : playsound"rollover" : End Sub 
Sub SW42_unHit:Controller.Switch(42)=0:End Sub
Sub SW52_Hit:Controller.Switch(52)=1 : playsound"rollover" : End Sub 
Sub SW52_unHit:Controller.Switch(52)=0:End Sub
Sub SW62_Hit:Controller.Switch(62)=1 : playsound"rollover" : End Sub 
Sub SW62_unHit:Controller.Switch(62)=0:End Sub
Sub SW72_Hit:Controller.Switch(72)=1 : playsound"rollover" : End Sub 
Sub SW72_unHit:Controller.Switch(72)=0:End Sub

'Star Triggers
Sub SW3_Hit:Controller.Switch(3)=1 : playsound"rollover" : End Sub 
Sub SW3_unHit:Controller.Switch(3)=0:End Sub
Sub SW13_Hit:Controller.Switch(13)=1 : playsound"rollover" : End Sub 
Sub SW13_unHit:Controller.Switch(13)=0:End Sub
Sub SW23_Hit:Controller.Switch(23)=1 : playsound"rollover" : End Sub 
Sub SW23_unHit:Controller.Switch(23)=0:End Sub
Sub SW33_Hit:Controller.Switch(33)=1 : playsound"rollover" : End Sub 
Sub SW33_unHit:Controller.Switch(33)=0:End Sub
Sub SW43_Hit:Controller.Switch(43)=1 : playsound"rollover" : End Sub 
Sub SW43_unHit:Controller.Switch(43)=0:End Sub

'Gate Trigger
Sub SW53_Hit:vpmTimer.pulseSw 53 :: End Sub 


'Bumpers
Sub Bumper1_Hit : vpmTimer.PulseSw(54) : playsound SoundFX("fx_bumper1",DOFContactors): End Sub
Sub Bumper2_Hit : vpmTimer.PulseSw(64) : playsound SoundFX("fx_bumper1",DOFContactors): End Sub

 'Scoring Rubber
Sub sw55_hit:vpmTimer.pulseSw 55 : playsound"flip_hit_3" : End Sub 

'Spinners
Sub sw74_Spin:vpmTimer.PulseSw 74 : playsound"fx_spinner" : End Sub

Sub Trigger3_Hit : TopFlipperActive = True : End Sub
Sub Trigger3_unHit : TopFlipperActive = False : End Sub

Sub Trigger2_Hit : playsound"Wire Ramp" : End Sub 

'**********************************************************************************************************
'animated wire ramp
'**********************************************************************************************************

Dim WireRamp_Dir

Sub UPFWireRamp_Enter_Hit
	WireRamp_Dir = -1
	UPFWireRampTimer.enabled = False
	UPFWireRampTimer.enabled = True
End Sub

Sub UPFWireRamp_Exit_Hit
	WireRamp_Dir = 1
	UPFWireRampTimer.enabled = False
	UPFWireRampTimer.enabled = True
	playsound"fx_ballrampdrop" 
End Sub

Sub UPFWireRampTimer_Timer
	P_WireRamp.ObjRotX = P_WireRamp.ObjRotX + WireRamp_Dir
	if P_WireRamp.ObjRotX <= -15 then
		UPFWireRampTimer.enabled = False
		WireRamp_Dir = 0
	end if
	if P_WireRamp.ObjRotX >= 5 then 
		UPFWireRampTimer.enabled = False
		P_WireRamp.ObjRotX = 5
		WireRamp_Dir = 0		
	end if
End Sub

'**********************************************************************************************************
'Map lights to an array
'**********************************************************************************************************


Set Lights(0)=L0
Set Lights(1)=L1
Set Lights(2)=L2
Set Lights(3)=L3

Set Lights(10)=L10
Set Lights(11)=L11
Set Lights(12)=L12

Set Lights(20)=L20
Set Lights(21)=L21
Set Lights(22)=L22

Set Lights(30)=L30
Set Lights(31)=L31
Set Lights(32)=L32

Set Lights(40)=L40
Set Lights(41)=L41
Set Lights(42)=L42

Set Lights(50)=L50
Set Lights(51)=L51
Set Lights(52)=L52

Set Lights(60)=L60
Set Lights(61)=L61
Set Lights(62)=L62  

Set Lights(70)=L70
Set Lights(71)=L71
Set Lights(72)=L72

Set Lights(80)=L80
Set Lights(81)=L81
Set Lights(82)=L82
Set Lights(83)=L83

Set Lights(89)=L89
Set Lights(90)=L90
Set Lights(91)=L91
Set Lights(92)=L92
Set Lights(93)=L93

Set Lights(99)=L99

Set Lights(100)=L100
Set Lights(101)=L101
Set Lights(102)=L102
Set Lights(103)=L103
Set Lights(109)=L109
Set Lights(110)=L110
Set Lights(111)=L111
Set Lights(112)=L112
Set Lights(113)=L113
Set Lights(119)=L119
Set Lights(120)=L120
Set Lights(121)=L121
Set Lights(122)=L122
Set Lights(123)=L123

Set Lights(129)=L129
Set Lights(130)=L130
Set Lights(131)=L131
Set Lights(132)=L132
Set Lights(133)=L133

Set Lights(153)=L153
Set Lights(163)=L163


'**********************************************************************************************************
'Digital Display
'**********************************************************************************************************

Dim Digits(28)
' 1st Player
Digits(0) = Array(LED10,LED11,LED12,LED13,LED14,LED15,LED16)
Digits(1) = Array(LED20,LED21,LED22,LED23,LED24,LED25,LED26)
Digits(2) = Array(LED30,LED31,LED32,LED33,LED34,LED35,LED36)
Digits(3) = Array(LED40,LED41,LED42,LED43,LED44,LED45,LED46)
Digits(4) = Array(LED50,LED51,LED52,LED53,LED54,LED55,LED56)
Digits(5) = Array(LED60,LED61,LED62,LED63,LED64,LED65,LED66)

' 2nd Player
Digits(6) = Array(LED80,LED81,LED82,LED83,LED84,LED85,LED86)
Digits(7) = Array(LED90,LED91,LED92,LED93,LED94,LED95,LED96)
Digits(8) = Array(LED100,LED101,LED102,LED103,LED104,LED105,LED106)
Digits(9) = Array(LED110,LED111,LED112,LED113,LED114,LED115,LED116)
Digits(10) = Array(LED120,LED121,LED122,LED123,LED124,LED125,LED126)
Digits(11) = Array(LED130,LED131,LED132,LED133,LED134,LED135,LED136)

' 3rd Player
Digits(12) = Array(LED150,LED151,LED152,LED153,LED154,LED155,LED156)
Digits(13) = Array(LED160,LED161,LED162,LED163,LED164,LED165,LED166)
Digits(14) = Array(LED170,LED171,LED172,LED173,LED174,LED175,LED176)
Digits(15) = Array(LED180,LED181,LED182,LED183,LED184,LED185,LED186)
Digits(16) = Array(LED190,LED191,LED192,LED193,LED194,LED195,LED196)
Digits(17) = Array(LED200,LED201,LED202,LED203,LED204,LED205,LED206)

' 4th Player
Digits(18) = Array(LED220,LED221,LED222,LED223,LED224,LED225,LED226)
Digits(19) = Array(LED230,LED231,LED232,LED233,LED234,LED235,LED236)
Digits(20) = Array(LED240,LED241,LED242,LED243,LED244,LED245,LED246)
Digits(21) = Array(LED250,LED251,LED252,LED253,LED254,LED255,LED256)
Digits(22) = Array(LED260,LED261,LED262,LED263,LED264,LED265,LED266)
Digits(23) = Array(LED270,LED271,LED272,LED273,LED274,LED275,LED276)

' Credits
Digits(24) = Array(LED4,LED2,LED6,LED7,LED5,LED1,LED3)
Digits(25) = Array(LED18,LED9,LED27,LED28,LED19,LED8,LED17)
' Balls
Digits(26) = Array(LED39,LED37,LED48,LED49,LED47,LED29,LED38)
Digits(27) = Array(LED67,LED58,LED69,LED77,LED68,LED57,LED59)

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
'	Start of VPX functions
'**********************************************************************************************************
'**********************************************************************************************************

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
	vpmTimer.PulseSw 61
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
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0:
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
	vpmTimer.PulseSw 51
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

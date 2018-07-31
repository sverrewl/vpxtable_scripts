Option Explicit
Randomize

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' No special SSF tweaks yet.

'******************* Options *********************
' Blackout Mode
Const BlackoutMode = 0		'0=Authentic mode,  1=Fantasy mode (GI remains off until ball drain)
'*************************************************

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="blkou_l1",UseSolenoids=2,UseLamps=1,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff",SCoin="coin"

LoadVPM "01320000", "S6.VBS", 3.21
Dim DesktopMode: DesktopMode = Table1.ShowDT

If DesktopMode = True Then 'Show Desktop components
Ramp16.visible=1
Ramp15.visible=1
Ramp10.visible=1

Else
Ramp16.visible=0
Ramp15.visible=0
Ramp10.visible=0

End if

'*************************************************************

'Solenoid Call backs
'**********************************************************************************************************
SolCallback(1)      = "bsTrough.SolOut"
SolCallback(4)      = "dtCDrop.SolDropUp"
SolCallback(5)      = "dtTDrop.SolDropUp"
SolCallback(6)      = "bsRightEject.SolOut"
' SolCallback(9)	= "VpmSolSound""SolChime1"
' SolCallback(10)	= "VpmSolSound""SolChime2"
' SolCallback(11)	= "VpmSolSound""SolChime3"
' SolCallback(12)	= "VpmSolSound""SolChime4"
' SolCallback(13)	= "VpmSolSound""SolChime5"
SolCallback(15)     = "SolSpecial"

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
Dim BlackoutFlag
Sub SolSpecial(Enabled)
	If Enabled Then
	'*****GI Lights Off
		dim xx
		For each xx in GI:xx.State = 0: Next
		If BlackoutMode = 1 Then BlackoutFlag = 1
        PlaySound "fx_relay"
	Else
		For each xx in GI:xx.State = 1: Next
        PlaySound "fx_relay"
	End if
 End Sub

 '**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************
 Dim bsTrough,bsRightEject,dtCDrop,dtTDrop
 
Sub Table1_Init
	vpmInit Me
	On Error Resume Next
		With Controller
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
         .SplashInfoLine = "Blackout (Williams 1980)" & vbNewLine & "You Suck"
		.HandleMechanics=0
		.HandleKeyboard=0
		.ShowDMDOnly=1
		.ShowFrame=0
		.ShowTitle=0
        .hidden = 1
		.Games(cGameName).Settings.Value("sound")=1
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
    vpmNudge.TiltSwitch = 1
	vpmNudge.Sensitivity = 5
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingshot)

	Set bsTrough = New cvpmBallStack
		bsTrough.Initsw 0,9,0,0,0,0,0,0
		bsTrough.InitKick BallRelease,120,3
        bsTrough.InitExitSnd  SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
		bsTrough.IsTrough = True 
		bsTrough.Balls = 1

	Set bsRightEject = New cvpmBallStack
		bsRightEject.InitSaucer sw40, 40, 220, 8
		bsRightEject.KickForceVar = 2
		bsRightEject.KickAngleVar = 2
        bsRightEject.InitExitSnd  SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)

  	Set dtCDrop=New cvpmDropTarget
		dtCDrop.InitDrop Array(sw25,sw26,sw27),Array(25,26,27)
		dtCDrop.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)
		dtCDrop.AllDownSw = 28

  	Set dtTDrop=New cvpmDropTarget
		dtTDrop.InitDrop Array(sw33,sw34,sw35),Array(33,34,35)
		dtTDrop.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)
		dtTDrop.AllDownSw = 36


  End Sub
 '**********************************************************************************************************
'Plunger code
'**********************************************************************************************************
 Sub Table1_KeyDown(ByVal keycode)
	If vpmKeyDown(KeyCode) Then Exit Sub
	If keycode=PlungerKey Then Plunger.Pullback:playsound"plungerpull"
End Sub

Sub Table1_KeyUp(ByVal keycode)
	If vpmKeyUp(KeyCode) Then Exit Sub
	If keycode=PlungerKey Then Plunger.Fire:playsound"plunger"
End Sub
'**********************************************************************************************************
 ' Switches & Targets
 '*******************
Sub sw40_Hit:bsRightEject.AddBall Me : playsound "popper_ball": End Sub
Sub sw40_UnHit: 
	dim xx
	If BlackoutFlag = 1 Then 'Turn GI off until ball drain
		For each xx in GI:xx.State = 0: Next
		PlaySound "fx_relay"
	End If
End Sub

Sub Drain_Hit
	dim xx
	playsound"drain"
	bsTrough.addball me
	If BlackoutFlag = 1 Then	'Turn GI back on
		BlackoutFlag = 0
		For each xx in GI:xx.State = 1: Next
		PlaySound "fx_relay"
	End If
End Sub

'scoring rubbers
Sub sw19_Slingshot:vpmTimer.PulseSw 19 : playsound"slingshot" : End Sub 
Sub sw38_Slingshot:vpmTimer.PulseSw 38 : playsound"slingshot" : End Sub  
 
'bumpers
Sub Bumper1_Hit : vpmTimer.PulseSw(22) : playsound SoundFX("fx_bumper1",DOFContactors): End Sub
Sub Bumper2_Hit : vpmTimer.PulseSw(21) : playsound SoundFX("fx_bumper1",DOFContactors): End Sub
Sub Bumper3_Hit : vpmTimer.PulseSw(23) : playsound SoundFX("fx_bumper1",DOFContactors): End Sub

'Rollover wire Triggers
 Sub sw10_Hit():Controller.Switch(10) = 1 : playsound"rollover" : End Sub
 Sub sw10_UnHit():Controller.Switch(10) = 0:End Sub
 Sub sw45_Hit():Controller.Switch(45) = 1 : playsound"rollover" : End Sub
 Sub sw45_UnHit():Controller.Switch(45) = 0:End Sub
 Sub sw42_Hit():Controller.Switch(42) = 1 : playsound"rollover" : End Sub
 Sub sw42_UnHit():Controller.Switch(42) = 0:End Sub
 Sub sw41_Hit():Controller.Switch(41) = 1 : playsound"rollover" : End Sub
 Sub sw41_UnHit():Controller.Switch(41) = 0:End Sub
 Sub sw29_Hit():Controller.Switch(29) = 1 : playsound"rollover" : End Sub
 Sub sw29_UnHit():Controller.Switch(29) = 0:End Sub
 Sub sw30_Hit():Controller.Switch(30) = 1 : playsound"rollover" : End Sub
 Sub sw30_UnHit():Controller.Switch(30) = 0:End Sub
 Sub sw31_Hit():Controller.Switch(31) = 1 : playsound"rollover" : End Sub
 Sub sw31_UnHit():Controller.Switch(31) = 0:End Sub
 
'Stand Up targets
Sub sw11_Hit:vpmTimer.PulseSw 11 : End Sub
Sub sw12_Hit:vpmTimer.PulseSw 12 : End Sub
Sub sw13_Hit:vpmTimer.PulseSw 13 : End Sub
Sub sw14_Hit:vpmTimer.PulseSw 14 : End Sub
Sub sw15_Hit:vpmTimer.PulseSw 15 : End Sub
Sub sw20_Hit:vpmTimer.PulseSw 20 : End Sub

'Spinners
  Sub Spinner1_Spin():vpmTimer.pulsesw 37 : playsound"fx_spinner" : End Sub
  Sub Spinner2_Spin():vpmTimer.pulsesw 39 : playsound"fx_spinner" : End Sub
  Sub Spinner3_Spin():vpmTimer.pulsesw 18 : playsound"fx_spinner" : End Sub

'Drop Targets
 Sub Sw25_Dropped: dtCDrop.Hit 1:  End Sub 
 Sub Sw26_Dropped: dtCDrop.Hit 2:  End Sub 
 Sub Sw27_Dropped: dtCDrop.Hit 3:  End Sub 

 Sub Sw33_Dropped: dtTDrop.Hit 1:  End Sub 
 Sub Sw34_Dropped: dtTDrop.Hit 2:  End Sub 
 Sub Sw35_Dropped: dtTDrop.Hit 3:  End Sub 



'**********************************************************************************************************
 
'Map lights to an array
'**********************************************************************************************************
Set Lights(1)=l1
Set Lights(2)=l2
Set Lights(3)=l3
Set Lights(4)=l4
Set Lights(5)=l5
Set Lights(6)=l6
Set Lights(7)=l7
Set Lights(8)=l8
Set Lights(9)=l9
Set Lights(10)=l10
Set Lights(12)=l12
Set Lights(13)=l13
Set Lights(14)=l14
Set Lights(15)=l15
Set Lights(16)=l16
Set Lights(17)=l17
Set Lights(18)=l18
Set Lights(19)=l19
Set Lights(20)=l20
Set Lights(21)=l21
Set Lights(22)=l22
Set Lights(23)=l23
Set Lights(24)=l24
Set Lights(25)=l25
Set Lights(26)=l26
Set Lights(27)=l27
Set Lights(28)=l28
Set Lights(29)=l29
Set Lights(30)=l30
Set Lights(31)=l31
Set Lights(32)=l32
Set Lights(33)=l33
Set Lights(34)=l34
Set Lights(35)=l35
Set Lights(36)=l36
Set Lights(37)=l37
Set Lights(38)=l38
Set Lights(39)=l39
Set Lights(40)=l40
Set Lights(41)=l41
Set Lights(42)=l42
Set Lights(43)=l43
Set Lights(44)=l44
Set Lights(45)=l45
Set Lights(46)=l46
Set Lights(47)=l47
Set Lights(48)=l48
Set Lights(49)=l49

'BackGlass
Set Lights(50)=l50
Set Lights(51)=l51
Set Lights(52)=l52
Set Lights(53)=l53
'Set Lights(54)=l54 'Match
Set Lights(55)=l55 ' Ball In Play
Set Lights(57)=l57
Set Lights(58)=l58
Set Lights(59)=l59
Set Lights(60)=l60
Set Lights(61)=l61 'TILT
Set Lights(62)=l62 'Game Over
Set Lights(63)=l63 'Shoot Again
Set Lights(64)=l64 'High Score


'if Full Screen turn off Backglas lights
If DesktopMode = True Then
dim xxx
For each xxx in BL:xxx.Visible = 1: Next
else
For each xxx in BL:xxx.Visible = 0: Next
End if

' ********************************************
' Light Displays (Segmented Displays)
' ********************************************

Dim Digits(28)

' 1st Player
Digits(0) = Array(LED10,LED11,LED12,LED13,LED14,LED15,LED16)
Digits(1) = Array(LED20,LED21,LED22,LED23,LED24,LED25,LED26)
Digits(2) = Array(LED30,LED31,LED32,LED33,LED34,LED35,LED36)
Digits(3) = Array(LED40,LED41,LED42,LED43,LED44,LED45,LED46)
Digits(4) = Array(LED50,LED51,LED52,LED53,LED54,LED55,LED56)
Digits(5) = Array(LED60,LED61,LED62,LED63,LED64,LED65,LED66)

' 2nd Player
Digits(6) = Array(LED70,LED71,LED72,LED73,LED74,LED75,LED76)
Digits(7) = Array(LED80,LED81,LED82,LED83,LED84,LED85,LED86)
Digits(8) = Array(LED90,LED91,LED92,LED93,LED94,LED95,LED96)
Digits(9) = Array(LED100,LED101,LED102,LED103,LED104,LED105,LED106)
Digits(10) = Array(LED110,LED111,LED112,LED113,LED114,LED115,LED116)
Digits(11) = Array(LED120,LED121,LED122,LED123,LED124,LED125,LED126)

' 3rd Player
Digits(12) = Array(LED130,LED131,LED132,LED133,LED134,LED135,LED136)
Digits(13) = Array(LED140,LED141,LED142,LED143,LED144,LED145,LED146)
Digits(14) = Array(LED150,LED151,LED152,LED153,LED154,LED155,LED156)
Digits(15) = Array(LED160,LED161,LED162,LED163,LED164,LED165,LED166)
Digits(16) = Array(LED170,LED171,LED172,LED173,LED174,LED175,LED176)
Digits(17) = Array(LED180,LED181,LED182,LED183,LED184,LED185,LED186)

' 4th Player
Digits(18) = Array(LED190,LED191,LED192,LED193,LED194,LED195,LED196)
Digits(19) = Array(LED200,LED201,LED202,LED203,LED204,LED205,LED206)
Digits(20) = Array(LED210,LED211,LED212,LED213,LED214,LED215,LED216)
Digits(21) = Array(LED220,LED221,LED222,LED223,LED224,LED225,LED226)
Digits(22) = Array(LED230,LED231,LED232,LED233,LED234,LED235,LED236)
Digits(23) = Array(LED240,LED241,LED242,LED243,LED244,LED245,LED246)

' Credits
Digits(24) = Array(LED250,LED251,LED252,LED253,LED254,LED255,LED256)
Digits(25) = Array(LED260,LED261,LED262,LED263,LED264,LED265,LED266)

' Balls
Digits(26) = Array(LED270,LED271,LED272,LED273,LED274,LED275,LED276)
Digits(27) = Array(LED280,LED281,LED282,LED283,LED284,LED285,LED286)


Sub DisplayTimer_Timer
	Dim ChgLED,ii,num,chg,stat,obj
	ChgLED = Controller.ChangedLEDs(&H0000003f, &Hffffffff)
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
'**********************************************************************************************************



'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
	vpmTimer.PulseSw 43
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
	vpmTimer.PulseSw 44
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
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0:
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

' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


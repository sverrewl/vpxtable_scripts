Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="slbmanib",UseSolenoids=1,UseLamps=1,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff",SFlipperOn="fx_Flipperup",SFlipperOff="fx_Flipperdown"
Const SCoin="coin"

LoadVPM "01130100", "Bally.VBS", 3.21
Dim DesktopMode: DesktopMode = Table1.ShowDT

'Solenoid Call backs
'**********************************************************************************************************
 SolCallback(6)		= "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
 SolCallback(7)     = "bsTrough.SolOut"			
 SolCallback(13)    = "UpKicker"
 SolCallback(14)    = "KickAndDown"
 
SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFx("fx_Flipperup",DOFContactors):LeftFlipper.RotateToEnd
     Else
         PlaySound SoundFx("fx_Flipperdown",DOFContactors):LeftFlipper.RotateToStart
     End If
  End Sub
  
Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFx("fx_Flipperup",DOFContactors):RightFlipper.RotateToEnd
     Else
         PlaySound SoundFx("fx_Flipperdown",DOFContactors):RightFlipper.RotateToStart
     End If
End Sub

'fantasy bumper flashers
 SolCallback(8) ="vpmFlasher array(Flasher8,Flasher8a),"
 SolCallback(9) ="vpmFlasher array(Flasher9,Flasher9a),"
 SolCallback(10)="vpmFlasher array(Flasher10,Flasher10a),"
'**********************************************************************************************************

'Solenoid Controlled toys
'**********************************************************************************************************
 dim NewlySet : NewlySet = False
 Sub KickAndDown(enabled)
 	If enabled Then
 		If NewlySet = True Then : KickIM.AutoFire : End If
 		KickPost.IsDropped = True
 		NewlySet = False
 	End If
 End Sub
 
 Sub UpKicker(enabled)
 	If enabled Then
 		KickPost.IsDropped = False
 		NewlySet = True
 	End If
 End Sub

'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************
Dim bsTrough, KickIM

Sub Table1_Init
	vpmInit Me
	On Error Resume Next
		With Controller
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
		.SplashInfoLine = "Pharaoh (Williams)"&chr(13)&"You Suck"
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

	vpmNudge.TiltSwitch  = 7
	vpmNudge.Sensitivity = 5
	vpmNudge.Tiltobj = Array(LeftSlingshot,RightSlingshot,Bumper1,Bumper2,Bumper3)

	Set bsTrough = New cvpmBallStack
	bsTrough.Initsw 0,8,0,0,0,0,0,0
	bsTrough.InitKick BallRelease, 55, 4
	bsTrough.InitExitSnd SoundFX("BallRelease",DOFContactors),SoundFX("solenoid",DOFContactors)
	bsTrough.Balls = 1

	Set KickIM = New cvpmImpulseP
	KickIM.InitImpulseP KickIt, 30, 0
	KickIM.Random 1
 	KickIM.Switch 1
	KickIM.InitExitSnd SoundFX("BallRelease",DOFContactors),SoundFX("solenoid",DOFContactors)
	KickIM.CreateEvents "KickIM"

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
' Switch handling
'**********************************************************************************************************
'**********************************************************************************************************

 ' Drain hole
Sub Drain_Hit:playsound"drain":bsTrough.addball me:End Sub

'Wire Triggers
Sub sw3_Hit:Controller.Switch(3)=1 : playsound"rollover" : End Sub 
Sub sw3_unHit:Controller.Switch(3)=0:End Sub
Sub sw4_Hit:Controller.Switch(4)=1 : playsound"rollover" : End Sub 
Sub sw4_unHit:Controller.Switch(4)=0:End Sub
Sub sw5_Hit:Controller.Switch(5)=1 : playsound"rollover" : End Sub 
Sub sw5_unHit:Controller.Switch(5)=0:End Sub
Sub sw26_Hit:Controller.Switch(26)=1 : playsound"rollover" : End Sub 
Sub sw26_unHit:Controller.Switch(26)=0:End Sub
Sub sw27_Hit:Controller.Switch(27)=1 : playsound"rollover" : End Sub 
Sub sw27_unHit:Controller.Switch(27)=0:End Sub
Sub sw29_Hit:Controller.Switch(29)=1 : playsound"rollover" : End Sub 
Sub sw29_unHit:Controller.Switch(29)=0:End Sub
Sub sw30_Hit:Controller.Switch(30)=1 : playsound"rollover" : End Sub 
Sub sw30_unHit:Controller.Switch(30)=0:End Sub

'Star Triggers
Sub sw2_Hit:Controller.Switch(2)=1 : playsound"rollover" : End Sub 
Sub sw2_unHit:Controller.Switch(2)=0:End Sub
Sub sw2a_Hit:Controller.Switch(2)=1 : playsound"rollover" : End Sub 
Sub sw2a_unHit:Controller.Switch(2)=0:End Sub

'scoring rubbers
Sub sw34e_Hit():vpmtimer.pulsesw 34 : playsound SoundFX("slingshot",DOFContactors): End Sub
Sub sw34f_Hit():vpmtimer.pulsesw 34 : playsound SoundFX("slingshot",DOFContactors): End Sub
Sub sw34g_Hit():vpmtimer.pulsesw 34 : playsound SoundFX("slingshot",DOFContactors): End Sub
Sub sw34h_Hit():vpmtimer.pulsesw 34 : playsound SoundFX("slingshot",DOFContactors): End Sub

'Spinners
Sub Spinner_Spin:vpmTimer.PulseSw 25 : playsound"fx_spinner" : End Sub
Sub Spinner1_Spin:vpmTimer.PulseSw 33 : playsound"fx_spinner" : End Sub

'Stand Up Targets
Sub sw17_Hit():vpmtimer.pulsesw 17: End Sub 
Sub sw18_Hit():vpmtimer.PulseSw 18: End Sub 
Sub sw19_Hit():vpmtimer.PulseSw 19: End Sub 
Sub sw20_Hit():vpmtimer.PulseSw 20: End Sub 
Sub sw21_Hit():vpmtimer.PulseSw 21: End Sub 
Sub sw22_Hit():vpmtimer.PulseSw 22: End Sub 
Sub sw23_Hit():vpmtimer.PulseSw 23: End Sub 
Sub sw24_Hit():vpmtimer.PulseSw 24: End Sub 
Sub sw28_Hit():vpmtimer.PulseSw 28: End Sub 
Sub sw31_Hit():vpmtimer.PulseSw 31: End Sub 
Sub sw32_Hit():vpmtimer.PulseSw 32: End Sub 


'Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw(38) : playsound SoundFX("fx_bumper1",DOFContactors): End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw(39) : playsound SoundFX("fx_bumper1",DOFContactors): End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw(40) : playsound SoundFX("fx_bumper1",DOFContactors): End Sub

Sub sw34a_Hit:vpmtimer.pulsesw(34) : End Sub
Sub sw34b_Hit:vpmtimer.pulsesw(34) : End Sub
Sub sw34c_Hit:vpmtimer.pulsesw(34) : End Sub
Sub sw34d_Hit:vpmtimer.pulsesw(34) : End Sub



Set Lights(1)=Light1
Set Lights(2)=Light2
Set Lights(3)=Light3
Set Lights(4)=Light4
Set Lights(5)=Light5
Set Lights(6)=Light6
Set Lights(7)=Light7
Set Lights(8)=Light8
Set Lights(9)=Light9
Set Lights(10)=Light10
Set Lights(11)=Light11
Set Lights(17)=Light17
Set Lights(18)=Light18
Set Lights(19)=Light19
Set Lights(20)=Light20
Set Lights(21)=Light21
Set Lights(22)=Light22
Set Lights(23)=Light23
Set Lights(24)=Light24
Lights(25)=array(Light25,Light25a)
Set Lights(26)=Light26
Set Lights(33)=Light33
Set Lights(34)=Light34
Set Lights(35)=Light35
Set Lights(36)=Light36
Set Lights(37)=Light37
Set Lights(38)=Light38
Set Lights(39)=Light39
Set Lights(40)=Light40
Set Lights(41)=Light41
Set Lights(42)=Light42
Set Lights(43)=Light43
Set Lights(44)=Light44
Set Lights(49)=Light49
Set Lights(50)=Light50
Set Lights(51)=Light51
Lights(52)=array(Light52,Light52a)
Set Lights(53)=Light53
Set Lights(54)=Light54
Set Lights(55)=Light55
Set Lights(56)=Light56
Lights(57)=array(Light57,Light57a)
Set Lights(58)=Light58
Set Lights(59)=Light59
Set Lights(60)=Light60
'Backglass
'Set Lights(11)=	' Shoot Again
'Set Lights(13)=	' Ball In Play
'Set Lights(27)=	' Match
'Set Lights(29)=	' High Score To Date
'Set Lights(45)=	' Game Over
'Set Lights(61)=	' Tilt
 'Silverball on Backglass
'Set Lights(14)=	' "S"
'Set Lights(30)=	' "I"
'Set Lights(46)=	' "L"
'Set Lights(62)=	' "V"
'Set Lights(15)=	' "E"
'Set Lights(31)=	' "R"
'Set Lights(47)=	' "B"
'Set Lights(63)=	' "A"
'Set Lights(12)=	' "L"
'Set Lights(28)=	' "L"

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
			end if
		next
		end if
end if
End Sub


'**********************************************************************************************************
'**********************************************************************************************************

'Bally Silverball Mania 7 digits
'Added by Inkochnito
Sub editDips
	Dim vpmDips:Set vpmDips=New cvpmDips
	With vpmDips
		.AddForm 415,400,"Silverball Mania 7 digits - DIP switches"
		.AddFrame 2,0,190,"Maximum credits",&H03000000,Array("10 credits",0,"15 credits",&H01000000,"25 credits",&H02000000,"free play (40 credits)",&H03000000)'dip 25&26
		.AddFrame 2,76,190,"Sound features",&H30000000,Array("chime effects",0,"no background noises",&H10000000,"noise effects",&H20000000,"background noises",&H30000000)'dip 29&30
		.AddFrame 2,152,190,"High score to date",&H00200000,Array("no award",0,"3 credits",&H00200000)'dip 22
		.AddFrame 2,200,190,"Score version",&H00100000,Array("6 digit scoring",0,"7 digit scoring",&H00100000)'dip 21
		.AddFrame 2,248,190,"Special/extra ball modes",&H00000060,Array("points",0,"extra ball",&H00000040,"replay/extra ball",&H00000060)'dip6&7
		.AddFrame 210,0,190,"Balls per game",&H40000000,Array("3 balls",0,"5 balls",&H40000000)'dip 31
		.AddFrame 210,46,190,"Carryover award",&H00000080,Array("1 credit",0,"3 credits",&H00000080)'dip 8
		.AddFrame 210,92,190,"Carryover advance",&H00004000,Array("advance on SBM special",&H00004000,"advance on kicker special",0)'dip 15
		.AddFrame 210,138,190,"Extra ball lites",&H80000000,Array("together with 5X bonus",&H80000000,"after 5X bonus",0)'dip32
		.AddFrame 210,184,190,"Kicker special lites",32768,Array("together with SBM special",32768,"after awarding SBM special",0)'dip 16
		.AddFrame 210,230,190,"Center hoop advances",&H00800000,Array("1 letter",0,"2 letters",&H00800000)'dip24
		.AddChk 210,280,190,Array("Match feature",&H08000000)'dip 28
		.AddChk 210,295,190,Array("Credits displayed",&H04000000)'dip 27
		.AddChk 210,310,200,Array("Silverball (backglass) carryover feature",&H00400000)'dip 23
		.AddLabel 50,330,340,20,"After hitting OK, press F3 to reset game with new settings."
		.ViewDips
	End With
End Sub
Set vpmShowDips=GetRef("editDips")


'**********************************************************************************************************
'**********************************************************************************************************
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
	vpmTimer.PulseSw 36
    PlaySound SoundFX("left_slingshot",DOFContactors), 0, 1, 0.05, 0.05
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
	vpmTimer.PulseSw 37
    PlaySound SoundFX("right_slingshot",DOFContactors),0,1,-0.05,0.05
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



' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 2000)
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
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0
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
	PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0
End Sub

'*****************************************
'			FLIPPER SHADOWS
'*****************************************

sub FlipperTimer_Timer()
	FlipperLSh.RotZ = LeftFlipper.currentangle
	FlipperRSh.RotZ = RightFlipper.currentangle

End Sub

'*****************************************
'			BALL SHADOW
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
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) + 10
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) - 10
        End If
        ballShadow(b).Y = BOT(b).Y + 20
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
	PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
End Sub

Sub Targets_Hit (idx)
	PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
End Sub

Sub Metals_Thin_Hit (idx)
	PlaySound "metalhit_thin", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub Metals_Medium_Hit (idx)
	PlaySound "metalhit_medium", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub Metals2_Hit (idx)
	PlaySound "metalhit2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub Gates_Hit (idx)
	PlaySound "gate4", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub Spinner_Spin
	PlaySound "fx_spinner",0,.25,0,0.25
End Sub

Sub Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then 
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub Posts_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 16 then 
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End if
	If finalspeed >= 6 AND finalspeed <= 16 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
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
		Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End Select
End Sub

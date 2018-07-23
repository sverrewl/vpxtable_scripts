Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="magicfp",UseSolenoids=1,UseLamps=1,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff", SCoin="coin"

LoadVPM "01560000","STERN.VBS",3.1
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

SolCallback(12)	= "dtL.SolDropUp" 'Drop Targets
SolCallback(13) = "bsLeftKicker.SolOut" 				'Sol12 Left Saucer --> 13 ok
SolCallback(6) = "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(14)	= "SolRaiseDropTargets"					'Sol14 Drop Target
SolCallback(15) = "bsTrough.SolOut"                    	'Sol15 Ball Release ok 
SolCallback(16) = "vpmNudge.SolGameOn"					'Sol15 --> 16 ok
SolCallback(19)	=  "Sol19"								'Sol19 Lock Out Coin Chute ok

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

Dim Sol19enabled
Sub Sol19(enabled)
	Sol19enabled = enabled
end sub

Sub SolRaiseDropTargets(enabled)
	If enabled Then
		dtL.DropSol_On
		Controller.Switch(28)=0
		Controller.Switch(32)=0
	end if
end sub


'*****GI Lights On
dim xx
For each xx in GI:xx.State = 1: Next


'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************

Dim bsTrough, dtL, bsLeftKicker, SLMagnet

Sub Table1_Init
	vpmInit Me
	On Error Resume Next
		With Controller
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
		.SplashInfoLine = "Magic (Stern)"&chr(13)&"You Suck"
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
	PinMAMETimer.Enabled = true

    vpmNudge.TiltSwitch=7
    vpmNudge.Sensitivity=2
	vpmNudge.TiltObj = Array(Bumper1,Bumper2,Bumper3,LeftslingShot,RightslingShot)

    Set bsTrough=New cvpmBallStack
        bsTrough.InitSw 0,33,0,0,0,0,0,0
        bsTrough.InitKick BallRelease,90,5
        bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
        bsTrough.Balls=1

    Set bsLeftKicker=New cvpmBallStack       
        bsLeftKicker.InitSaucer LeftKicker,26,70,15 
        bsLeftKicker.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)

   ' Saucer Magnet Left
    Set SLMagnet = New cvpmMagnet
        SLMagnet.InitMagnet SaucerLeftMagnet, 9
        SLMagnet.GrabCenter = 0
        SLMagnet.MagnetOn = 1
        SLMagnet.CreateEvents "SLMagnet"

	set dtL = new cvpmDropTarget
		dtL.InitDrop Array(sw29,sw30,sw31), Array(29,30,31)
		dtL.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)
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
Sub Drain_Hit:playsound"drain":bsTrough.addball me:End Sub
Sub LeftKicker_Hit:bsLeftKicker.AddBall 0 : playsound"popper_ball" : End Sub

'Spinners
Sub sw4_Spin:vpmTimer.PulseSw 4 : playsound"fx_spinner" : End Sub
Sub sw5_Spin:vpmTimer.PulseSw 5 : playsound"fx_spinner" : End Sub

'Star Triggers
sub sw10_hit: Controller.Switch(10)=1 : playsound"rollover" : End Sub 
sub sw10_unhit:Controller.Switch(10)=0:End Sub
sub sw18_hit: Controller.Switch(18)=1 : playsound"rollover" : End Sub 
sub sw18_unhit:Controller.Switch(18)=0:End Sub
sub sw27_hit: Controller.Switch(27)=1 : playsound"rollover" : End Sub 
sub sw27_unhit:Controller.Switch(27)=0:End Sub
sub sw40_hit: Controller.Switch(40)=1 : playsound"rollover" : End Sub 
sub sw40_unhit:Controller.Switch(40)=0:End Sub

'Bumpers
sub Bumper1_hit : vpmTimer.PulseSw(12) : playsound SoundFX("fx_bumper1",DOFContactors): End Sub
sub Bumper2_hit : vpmTimer.PulseSw(16) : playsound SoundFX("fx_bumper1",DOFContactors): End Sub
sub Bumper3_hit : vpmTimer.PulseSw(15) : playsound SoundFX("fx_bumper1",DOFContactors): End Sub

'Stamd Up Targets
sub sw21_hit : vpmTimer.pulseSw 21: DOF 102, DOFPulse : End Sub
sub sw22_hit : vpmTimer.pulseSw 22: DOF 101, DOFPulse : End Sub
sub sw23_hit : vpmTimer.pulseSw 23: DOF 102, DOFPulse : End Sub
sub sw24_hit : vpmTimer.pulseSw 24: DOF 101, DOFPulse : End Sub

'Scoring Rubber
Sub sw27a_hit:vpmTimer.pulseSw 27 : playsound"flip_hit_3" : End Sub 
Sub sw27b_hit:vpmTimer.pulseSw 27 : playsound"flip_hit_3" : End Sub 

'Drop Targets
Sub Sw29_Dropped:dtL.Hit 1  : Controller.Switch(28)=1: End Sub  
Sub Sw30_Dropped:dtL.Hit 2 :End Sub  
Sub Sw31_Dropped:dtL.Hit 3  : Controller.Switch(32)=1: End Sub 

'Wire Triggers
Sub sw17_hit:Controller.Switch(17)=1 : playsound"rollover" : End Sub 
Sub sw17_unhit:Controller.Switch(17)=0:End Sub
Sub sw19_hit:Controller.Switch(19)=1 : playsound"rollover" : DOF 105, DOFOn : End Sub
Sub sw19_unhit:Controller.Switch(19)=0: DOF 105, DOFOff : End Sub
Sub sw19a_hit:Controller.Switch(19)=1 : playsound"rollover" : DOF 106, DOFOn : End Sub
Sub sw19a_unhit:Controller.Switch(19)=0: DOF 106, DOFOff : End Sub
Sub sw39_hit:Controller.Switch(39)=1 : playsound"rollover" : End Sub 
Sub sw39_unhit:Controller.Switch(39)=0:End Sub
Sub sw20_hit:Controller.Switch(20)=1 : playsound"rollover" : End Sub 
Sub sw20_unhit:Controller.Switch(20)=0:End Sub
Sub sw22a_hit:Controller.Switch(22)=1 : playsound"rollover" : DOF 104, DOFOn : End Sub 
Sub sw22a_unhit:Controller.Switch(22)=0: DOF 104, DOFOff : End Sub
Sub sw23a_hit:Controller.Switch(23)=1 : playsound"rollover" : DOF 103, DOFOn : End Sub 
Sub sw23a_unhit:Controller.Switch(23)=0 : DOF 103, DOFOff : End Sub


Set Lights(1)=Light1
Set Lights(2)=Light2
Set Lights(3)=Light3
Set Lights(4)=Light4
Set Lights(5)=Light5
Lights(10)=array(Light10b,Light10a)
Set Lights(11)=Light11
Set Lights(12)=Light12
Set Lights(17)=Light17
Set Lights(18)=Light18
Set Lights(19)=Light19
Set Lights(20)=Light20
Set Lights(21)=Light21
Lights(26)=array(Light26b,Light26a)
Set Lights(28)=Light28
Set Lights(33)=Light33
Set Lights(34)=Light34
Lights(36)=array(Light36b,Light36a)
Set Lights(37)=Light37
Lights(40)=array(Light40,Light40b,Light40a)
Set Lights(43)=Light43
Set Lights(44)=Light44
Set Lights(49)=Light49
Set Lights(50)=Light50
Set Lights(51)=Light51
Set Lights(52)=Light52
Lights(53)=array(Light53a,Light53b)
Set Lights(55)=Light55
Set Lights(60)=Light60


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

'Stern Magic
'added by Inkochnito
Sub editDips
	Dim vpmDips : Set vpmDips = New cvpmDips
	With vpmDips
		.AddForm 700,400,"Magic - DIP switches"
		.AddChk 2,10,180,Array("Match feature",&H00100000)'dip 21
		.AddChk 205,10,115,Array("Credits display",&H00080000)'dip 20
		.AddFrame 2,30,190,"Maximum credits",&H00070000,Array("5 credits",0,"10 credits",&H00010000,"15 credits",&H00020000,"20 credits",&H00030000,"25 credits",&H00040000,"30 credits",&H00050000,"35 credits",&H00060000,"40 credits",&H00070000)'dip 17&18&19
		.AddFrame 2,160,190,"High game to date",49152,Array("points",0,"1 free game",&H00004000,"2 free games",32768,"3 free games",49152)'dip 15&16
		.AddFrame 2,235,190,"Special award",&HC0000000,Array("100,000 points",0,"free ball",&H40000000,"free game",&H80000000,"free ball and free game",&HC0000000)'dip 31&32
		.AddFrame 2,310,190,"High score feature",&H00000020,Array("extra ball",0,"replay",&H00000020)'dip 6
		.AddFrame 205,30,190,"Balls per game",&H00000040,Array("3 balls",0,"5 balls",&H00000040)'dip 7
		.AddFrame 205,76,190,"Extra ball alternation",&H04000000,Array("no alternation",0,"Inner lane alternating with outlane",&H04000000)'dip 27
		.AddFrame 205,122,190,"Target special limit",&H01000000,Array ("no limit",0,"one replay per ball",&H01000000)'dip 25
		.AddFrame 205,169,190,"Outlane special limit",&H02000000,Array("no limit",0,"one replay per ball",&H02000000)'dip 26
		.AddFrame 205,216,190,"Outlane special lite on after",&H00800000,Array("second target reset",0,"third target reset",&H00800000)'dip 24
		.AddFrame 205,263,190,"Extra ball",&H00400000,Array("no extra ball (bypass)",0,"award extra ball",&H00400000)'dip 23
		.AddFrame 205,310,190,"Melody option",&H00000080,Array("2 tones only",0,"full melody",&H00000080)'dip 8
		.AddLabel 50,382,300,20,"After hitting OK, press F3 to reset game with new settings."
		.ViewDips
	End With
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
	vpmTimer.pulseSw 13
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
	vpmTimer.pulseSw 14
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

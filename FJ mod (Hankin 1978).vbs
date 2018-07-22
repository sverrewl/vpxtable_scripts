Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="fjholden",UseSolenoids=1,UseLamps=1,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff",SCoin="coin" 

LoadVPM "01120100","Hankin.VBS",3.02
Dim DesktopMode: DesktopMode = Table1.ShowDT
'*************************************************************

'Solenoid Call backs
'**********************************************************************************************************
SolCallback(1)="vpmSolSound SoundFX(""knocker"",DOFKnocker),"
SolCallback(2)="bsTopHole.SolOut"
SolCallback(3)="bsLeftHole.SolOut"
SolCallback(4)="bsRightHole.SolOut"
SolCallback(7)="bsTrough.SolOut"
SolCallback(19)="vpmNudge.SolGameOn"

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

'Fantasy Bumper Lights
SolCallback(9)="vpmFlasher array(Flasher9,Flasher9a),"  'Top Bumper
SolCallback(10)="vpmFlasher array(Flasher10,Flasher10a)," 'Right Bumper
SolCallback(11)="vpmFlasher array(Flasher11,Flasher11a)," 'Left Bumper

'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************
Dim bsTrough, bsRightHole, bsLeftHole, bsTopHole

Sub Table1_Init
	vpmInit Me
	On Error Resume Next
		With Controller
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
		.SplashInfoLine = "FJ (Hankin 1978)"&chr(13)&"You Suck"
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
	
	vpmNudge.TiltSwitch=2 
	vpmNudge.Sensitivity=1
	vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingShot)

	Set bsTrough=New cvpmBallStack ' Trough handler
    bsTrough.InitSw 0,40,0,0,0,0,0,0
	bsTrough.InitKick ballrelease,110,5
	bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors),SoundFX("solenoid",DOFContactors)
	bsTrough.Balls=1
	
	Set bsTopHole=New cvpmBallStack
	bsTopHole.InitSaucer Kicker3,32,179+rnd(1)*3,5+rnd(1)*4
	bsTopHole.InitExitSnd SoundFX("popper",DOFContactors),SoundFX("solenoid",DOFContactors)
	
	Set bsLeftHole=New cvpmBallStack
	bsLeftHole.InitSaucer Kicker2,38,89+rnd(1)*2,7+rnd(1)*3
	bsLeftHole.InitExitSnd SoundFX("popper",DOFContactors),SoundFX("solenoid",DOFContactors)

	Set bsRightHole=New cvpmBallStack
	bsRightHole.InitSaucer Kicker1,39,299+rnd(1)*2,7+rnd(1)*3
	bsRightHole.InitExitSnd SoundFX("popper",DOFContactors),SoundFX("solenoid",DOFContactors)
	
End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)
	If KeyDownHandler(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Pullback:playsound"plungerpull"
	If keycode=AddCreditKey then vpmTimer.pulseSW (swCoin1)
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
	If KeyUpHandler(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Fire:PlaySound"plunger"
End Sub

'**********************************************************************************************************

 ' Drain hole
Sub Drain_Hit:playsound"drain":bsTrough.addball me:End Sub
Sub Kicker1_Hit:bsRightHole.AddBall 0:End Sub
Sub Kicker2_Hit:bsLeftHole.AddBall 0:End Sub
Sub Kicker3_Hit:bsTopHole.AddBall 0:End Sub

'Bumpers
Sub Bumper1_Hit : vpmTimer.PulseSw(35) : playsound SoundFX("fx_bumper1",DOFContactors): End Sub
Sub Bumper2_Hit : vpmTimer.PulseSw(33) : playsound SoundFX("fx_bumper1",DOFContactors): End Sub
Sub Bumper3_Hit : vpmTimer.PulseSw(34) : playsound SoundFX("fx_bumper1",DOFContactors): End Sub

'scoring rubbers
Sub sw11a_Slingshot:vpmTimer.PulseSw 11 : playsound SoundFX("slingshot",DOFContactors) : End Sub 
Sub sw11b_Slingshot:vpmTimer.PulseSw 11 : playsound SoundFX("slingshot",DOFContactors) : End Sub 


'Star Triggers
Sub sw12_Hit:Controller.Switch(12)=1 : playsound"rollover" : End Sub 
Sub sw12_unHit:Controller.Switch(12)=0:End Sub
Sub sw13_Hit:Controller.Switch(13)=1 : playsound"rollover" : End Sub 
Sub sw13_unHit:Controller.Switch(13)=0:End Sub
Sub sw14_Hit:Controller.Switch(14)=1 : playsound"rollover" : End Sub 
Sub sw14_unHit:Controller.Switch(14)=0:End Sub
Sub sw15_Hit:Controller.Switch(15)=1 : playsound"rollover" : End Sub 
Sub sw15_unHit:Controller.Switch(15)=0:End Sub
Sub sw17_Hit:Controller.Switch(17)=1 : playsound"rollover" : End Sub 
Sub sw17_unHit:Controller.Switch(17)=0:End Sub
Sub sw18_Hit:Controller.Switch(18)=1 : playsound"rollover" : End Sub 
Sub sw18_unHit:Controller.Switch(18)=0:End Sub

'Wire Triggers
Sub sw19_Hit:Controller.Switch(19)=1 : playsound"rollover" : End Sub 
Sub sw19_unHit:Controller.Switch(19)=0:End Sub
Sub sw20_Hit:Controller.Switch(20)=1 : playsound"rollover" : End Sub 
Sub sw20_unHit:Controller.Switch(20)=0:End Sub
Sub sw21_Hit:Controller.Switch(21)=1 : playsound"rollover" : End Sub 
Sub sw21_unHit:Controller.Switch(21)=0:End Sub
Sub sw28_Hit:Controller.Switch(28)=1 : playsound"rollover" : End Sub 
Sub sw28_unHit:Controller.Switch(28)=0:End Sub
Sub sw29_Hit:Controller.Switch(29)=1 : playsound"rollover" : End Sub 
Sub sw29_unHit:Controller.Switch(29)=0:End Sub
Sub sw30_Hit:Controller.Switch(30)=1 : playsound"rollover" : End Sub 
Sub sw30_unHit:Controller.Switch(30)=0:End Sub

'Spinners
Sub Spinner_Spin:vpmTimer.PulseSw(25) : playsound"fx_spinner" : End Sub
Sub Spinner1_Spin:vpmTimer.PulseSw(26) : playsound"fx_spinner" : End Sub


'Stand Up Targets
Sub sw22_Hit:vpmTimer.PulseSw(22):End Sub
Sub sw23_Hit:vpmTimer.PulseSw(23):End Sub
Sub sw24_Hit:vpmTimer.PulseSw(24):End Sub

Sub sw27_Hit:vpmTimer.PulseSw(27):End Sub

Sub sw31_Hit:vpmTimer.PulseSw(31):End Sub


Set Lights(1)=L1
Set Lights(2)=L2
Set Lights(3)=L3
Set Lights(4)=L4
Set Lights(5)=L5
Set Lights(6)=L6
Set Lights(8)=L8
Set Lights(9)=L9
Set Lights(10)= L10
Set Lights(11)= L11
Set Lights(12)= L12
Set Lights(13)= L13
Set Lights(17)= L17
Set Lights(18)= L18
Set Lights(19)= L19
Set Lights(20)= L20
Set Lights(22)= L22
Set Lights(23)= L23
Set Lights(24)= L24
Set Lights(25)= L25
Set Lights(26)= L26
Set Lights(28)= L28
Set Lights(33)= L33
Set Lights(34)= L34
Set Lights(35)= L35
Set Lights(38)= L38
Set Lights(39)= L39
Set Lights(40)= L40
Set Lights(41)= L41
Set Lights(42)= L42
Set Lights(43)= L43
Set Lights(44)= L44
Set Lights(49)= L49
Set Lights(51)=L51
Set Lights(52)=L52
Set Lights(54)=L54
Set Lights(55)=L55
Set Lights(56)=L56
Set Lights(57)=L57
Set Lights(58)=L58
Set Lights(59)=L59
Set Lights(60)=L60

'Backglass Light Displays (7 digit 7 segment displays)
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



'Hankin FJ Holden
'Added by Inkochnito
Sub editDips
	Dim vpmDips:Set vpmDips=New cvpmDips
	With vpmDips
		.AddForm 700,400,"FJ Holden - DIP switches"
		.AddFrame 0,0,180,"Credits per coins",&H00000007,Array("1/1",&H00,"1/2",&H01,"1/3",&H02,"1/4",&H03,"1/1, 3/2",&H04,"1/2, 2/3",&H05,"2/3, 4/5",&H06,"2/5",&H07)'dips 1+2+3
		.AddFrame 0,137,180,"High game to date award",&H00000060,Array("No award",0,"1 free game",&H00000020,"2 free games",&H00000040,"3 free games",&H00000060)'dip 6+7
		.AddFrame 200,0,180,"Maximum credits",&H00030000,Array("5",0,"10",&H00010000,"15",&H00020000,"20",&H00030000)'dip 17+18
		.AddFrame 200,76,180,"High score award",&H00000008,Array("Free game",0,"Extra ball",&H00000008)'dip 4
		.AddFrame 200,122,180,"Eject hole bonus award",&H00000100,Array("Awards 1 bonus advance",0,"Awards 2 bonus advances",&H00000100)'dip 9
		.AddFrame 200,168,180,"Triple target lamp cycle speed",&H00001000,Array("Fast",0,"Slow",&H00001000)'dip 13
		.AddFrame 200,214,180,"Extra ball award after",&H00002000,Array("completing row 3",0,"completing row 2",&H00002000)'dip 14
		.AddFrame 200,260,180,"Pop bumper award",&H00004000,Array("3X",0,"2X",&H00004000)'dip 15
		.AddFrame 0,214,180,"Balls per game",32768,Array("3 balls",0,"5 balls",32768)'dip 16
		.AddFrame 0,260,180,"Free game sound",&H00200000,Array("Special tune",0,"Knocker",&H00200000)'dip 22
		.AddChk 0,310,180,Array("Triple target status memory",&H00000200)'dip 10
		.AddChk 0,325,180,Array("Background sound",&H00000400)'dip 11
		.AddChk 0,340,180,Array("Match feature",&H00000010)'dip 5
		.AddChk 200,310,180,Array("Game Over tune",&H00000080)'dip 8
		.AddChk 200,325,180,Array("Coin counter reset enable",&H00400000)'dip 23
		.AddChk 200,340,180,Array("Self-test Time-out feature enable",&H00800000)'dip 24
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
	vpmTimer.PulseSw(36)
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
	vpmTimer.PulseSw(37)
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

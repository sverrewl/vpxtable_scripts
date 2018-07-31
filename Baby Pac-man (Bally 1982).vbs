Option Explicit
Randomize

' Thalamus 2018-07-19
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' No special SSF tweaks yet.

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01560000", "BALLY.VBS", 3.26
Dim DesktopMode: DesktopMode = Table1.ShowDT

Const cGameName="babypac",UseSolenoids=1,UseLamps=1,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff",SCoin="coin"

'if Desktop hide Backpannel primitive
If DesktopMode = True Then
  Primitive3.Visible = 0
  Ramp16.visible=1
  Ramp15.visible=1
  Ramp19.visible=1
  Ramp20.visible=1
else
  Primitive3.Visible = 1
  Ramp16.visible=0
  Ramp15.visible=0
  Ramp19.visible=0
  Ramp20.visible=0
End if

'*************************************************************
'Solenoid Call backs
'**********************************************************************************************************
 'Targets handled by class
 '***************************Solenoids
   SolCallback(1)  = "bsOutSaucer.SolOut"
   SolCallback(2)  = "dtC.SolDropUp"
   SolCallback(3)  = "dtC.SolHit 1,"
   SolCallback(5)  = "dtC.SolHit 3,"
   SolCallback(7)  = "dtC.SolHit 5,"
   SolCallback(8)  = "solFlipOn"
   SolCallback(9)  = "bsLSaucer.SolOut"
   SolCallback(10) = "bsRSaucer.SolOut"
  '*****Flipper Subs
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
 		 Controller.Switch(1) = 1
     Else
         PlaySound SoundFX("fx_Flipperdown",DOFContactors):RightFlipper.RotateToStart
 		 Controller.Switch(1) = 0
     End If
End Sub

 'Solenoid Controlled toys
'**********************************************************************************************************

  Sub solFlipOn(Enabled)  'Flipper Relay kills flippers when not enabled
 	If enabled then
 		FOn = 1
 	Else
 		FOn = 0
 		LeftFlipper.RotateToStart
 		RightFlipper.RotateToStart
 	End if
  End Sub

'*****GI Lights On
dim xx
For each xx in GI:xx.State = 1: Next

'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************
  Dim bsOutSaucer, bsRSaucer, bsLSaucer, dtC, GameOn, FOn
  GameOn=0

Sub Table1_Init
	vpmInit Me
	On Error Resume Next
		With Controller
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
		.SplashInfoLine = "Baby Pacman Bally 1982"&chr(13)&"You Suck"
		.HandleMechanics=0
		.HandleKeyboard=0
		.ShowDMDOnly=1
		.ShowFrame=0
		.ShowTitle=0
        .hidden = 0
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With
     On Error Goto 0

	PinMAMETimer.Interval=PinMAMEInterval
	PinMAMETimer.Enabled=1

    vpmNudge.TiltSwitch = 15
    vpmNudge.Sensitivity = 4
    vpmNudge.TiltObj = Array(LeftSlingShot, RightSlingShot)

  	Set bsOutSaucer = New cvpmBallStack
  		bsOutSaucer.InitSaucer  sw30,30,250,40
 		bsOutSaucer.Balls = 1
        bsOutSaucer.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)

   Set bsRSaucer = New cvpmBallStack
       bsRSaucer.InitSaucer sw31, 31, 210, 10
       bsRSaucer.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)

   Set bsLSaucer = New cvpmBallStack
       bsLSaucer.InitSaucer sw32, 32, 150, 10
       bsLSaucer.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)

  Set dtC = New cvpmDropTarget
      dtC.InitDrop Array(sw29, sw28, sw27, sw26, sw25), Array(29, 28, 27, 26, 25)
      dtC.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

 '*****Create Loop Captive Ball
  CBKick.CreateBall
  CBKick.Kick 180, 1

 '*****Create outhole ball
  sw30_Hit()

  '*****Flipper Init
 	FOn = 0
   End Sub
'**********************************************************************************************************


   Sub BabyPacman_Paused:Controller.Pause = 1:End Sub
   Sub BabyPacman_unPaused:Controller.Pause = 0:End Sub


'Plunger code
'**********************************************************************************************************
   Sub Table1_KeyDown(ByVal Keycode)
 	'****Jotstick based on ellenoff script
 	 If Keycode = keyJoyUp then:Controller.Switch(-0)=1
 	 If Keycode = keyJoyDown then:Controller.Switch(-1)=1
 	 If Keycode = keyJoyLeft then:Controller.Switch(-2)=1
 	 If Keycode = keyJoyRight then:Controller.Switch(-3)=1
 	 If Keycode = keyInsertCoin1 then:vpmTimer.PulseSw 10
 	 If Keycode = keyInsertCoin2 then:vpmTimer.PulseSw 10
 	 If Keycode = keyInsertCoin3 then:vpmTimer.PulseSw 10
 	 If Keycode = keyInsertCoin4 then:vpmTimer.PulseSw 10
  	 If vpmKeyDown(keycode) Then Exit Sub
 	 If Keycode = keyReset then
  	end if
   End Sub

   Sub Table1_KeyUp(ByVal Keycode)
 	 If Keycode = keyJoyUp then:Controller.Switch(-0)=0
 	 If Keycode = keyJoyDown then:Controller.Switch(-1)=0
 	 If Keycode = keyJoyLeft then:Controller.Switch(-2)=0
 	 If Keycode = keyJoyRight then:Controller.Switch(-3)=0
       If vpmKeyUp(keycode) Then Exit Sub
   End Sub
'**********************************************************************************************************
'**********************************************************************************************************


 ' Drain hole
Sub sw30_Hit:playsound"drain":bsOutSaucer.addball 0:sw30.Enabled = False:End Sub
Sub sw32_Hit:bsLSaucer.AddBall 0:End Sub
Sub sw31_Hit:bsRSaucer.AddBall 0:End Sub

  '*****Rollovers
 Sub sw21_Hit:Controller.Switch(21) = 1 : playsound"rollover" : End Sub
 Sub sw21_UnHit:Controller.Switch(21) = 0:End Sub
 Sub sw22_Hit:Controller.Switch(22) = 1 : playsound"rollover" : End Sub
 Sub sw22_UnHit:Controller.Switch(22) = 0:End Sub
 Sub sw23_Hit:Controller.Switch(23) = 1 : playsound"rollover" : End Sub
 Sub sw23_UnHit:Controller.Switch(23) = 0:End Sub
 Sub sw24_Hit:Controller.Switch(24) = 1 : playsound"rollover" : End Sub
 Sub sw24_UnHit:Controller.Switch(24) = 0:End Sub
 'Loop Switches
  Sub sw20_Hit:Controller.Switch(20) = 1 : BallSide=1 : playsound"rollover" : End Sub
 Sub sw20_UnHit:Controller.Switch(20) = 0:End Sub
 Sub sw17_Hit:Controller.Switch(17) = 1 : BallSide=2 : playsound"rollover" : End Sub
 Sub sw17_UnHit:Controller.Switch(17) = 0:End Sub

  '*****Gates & Spinner
   Sub sw7_Spin:vpmTimer.PulseSw 7 : playsound"fx_spinner" : End Sub
   Sub sw8_Spin:vpmTimer.PulseSw 8 : playsound"fx_spinner" : End Sub

'Drop Targets
 Sub Sw29_Dropped:dtC.Hit 1 :End Sub
 Sub Sw28_Dropped:dtC.Hit 2 :End Sub
 Sub Sw27_Dropped:dtC.Hit 3 :End Sub
 Sub Sw26_Dropped:dtC.Hit 4 :End Sub
 Sub Sw25_Dropped:dtC.Hit 5 :End Sub

'**********************************************************************************************************

  	Set Lights(1)= l1
  	Set Lights(2)= l2
  	Set Lights(3)= l3
  	Set Lights(4)= l4
  	Set Lights(5)= l5
  	Set Lights(6)= l6
  	Set Lights(7)= l7
  	Set Lights(8)= l8
  	Set Lights(9)= l9
  	Set Lights(10)= l10
  	Set Lights(11)= l11
  	Set Lights(12)= l12
  	Set Lights(13)= l13

	Set Lights(14)= l14

  	Set Lights(17)= l17
  	Set Lights(18)= l18
  	Set Lights(19)= l19
  	Set Lights(20)= l20
  	Set Lights(21)= l21
  	Set Lights(22)= l22
  	Set Lights(23)= l23
  	Set Lights(24)= l24
  	Set Lights(25)= l25
  	Set Lights(26)= l26
  	Set Lights(27)= l27
  	Set Lights(28)= l28
  	Set Lights(29)= l29

  	Set Lights(30)= l30

  	Set Lights(32)= l32
  	Set Lights(33)= l33
  	Set Lights(34)= l34
  	Set Lights(35)= l35
  	Set Lights(36)= l36
  	Set Lights(37)= l37
  	Set Lights(38)= l38
  	Set Lights(39)= l39
  	Set Lights(40)= l40
  	Set Lights(41)= l41
  	Set Lights(42)= l42
  	Set Lights(43)= l43
  	Set Lights(44)= l44
  	Set Lights(45)= l45
  	Set Lights(46)= l46
  	Set Lights(49)= l49
  	Set Lights(50)= l50
  	Set Lights(51)= l51
  	Set Lights(52)= l52
  	Set Lights(53)= l53
  	Set Lights(54)= l54
  	Set Lights(55)= l55
  	Set Lights(56)= l56
  	Set Lights(57)= l57
  	Set Lights(58)= l58
  	Set Lights(59)= l59
  	Set Lights(60)= l60
  	Set Lights(61)= l61
	Set Lights(62)= l62

'added by Inkochnito
Sub editDips
	Dim vpmDips : Set vpmDips = New cvpmDips
	With vpmDips
		.AddForm 700,400,"Baby Pacman - DIP switches"
		.AddFrame 2,0,190,"Any center lites on will",&H00000020,Array("be reset for next Pac-Man",0,"come back on for next Pac-Man",&H00000020)'dip 6
		.AddFrame 2,46,190,"Pac maze cherry",&H00000040,Array("no cherry at start of maze",0,"cherry will show on maze",&H00000040)'dip 7
		.AddFrame 2,92,190,"Pac maze side tunnels",&H00000080,Array("gates are closed at start of maze",0,"gates are open at start of maze",&H00000080)'dip 8
		.AddFrame 2,138,190,"After 3 ball ejects and no score",&H00002000,Array("the maze will lite",0,"ball is still in play",&H00002000)'dip 14
		.AddFrame 2,184,190,"Any energizers on will",49152,Array ("be reset for next Pac-Man",0,"come back on for next Pac-Man",&H00004000)'dip 15
		.AddFrame 205,0,190,"Pac-men per game",&HC0000000,Array ("2 Pac-men",&HC0000000,"3 Pac-men",0,"4 Pac-men",&H80000000,"5 Pac-men",&H40000000)'dip 31&32
		.AddFrame 205,78,190,"Maze special awarded after",&H00600000,Array("completing 3 mazes",0,"completing 4 mazes",&H00200000,"completing 5 mazes",&H00600000)'dip 22&23
		.AddFrame 205,138,190,"Playfield/video (trouble shooting only)",32768,Array("playfield will operate with video",0,"playfield will operate only",32768)'dip 16
		.AddFrame 205,184,190,"Any center arrows on will",&H00800000,Array("go out for next ball in outhole",0,"stay on for next ball in outhole",&H00800000)'dip 24
		.AddChk 2,240,120,Array("Credits displayed",&H04000000)'dip 27
		.AddChk 205,240,120,Array("Free play",&H20000000)'dip 30
		.AddLabel 20,260,380,20,"Set selftest 'High score mode' and 'Special mode' to 03 for the best gameplay."
		.AddLabel 50,280,300,20,"After hitting OK, press F3 to reset game with new settings."
		.ViewDips
	End With
End Sub
Set vpmShowDips = GetRef("editDips")


'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
vpmTimer.PulseSw 5
    PlaySound SoundFXDOF("right_slingshot",102,DOFPulse,DOFContactors), 0, 1, 0.05, 0.05
    'RSling.Visible = 0
    'RSling1.Visible = 1
    'sling1.TransZ = -20
    'RStep = 0
    'RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0:
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
vpmTimer.PulseSw 5
    PlaySound SoundFXDOF("left_slingshot",101,DOFPulse,DOFContactors),0,1,-0.05,0.05
    'LSling.Visible = 0
    'LSling1.Visible = 1
    'sling2.TransZ = -20
    'LStep = 0
    'LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0:
    End Select
    LStep = LStep + 1
End Sub

Dim BallSide

Sub DropFix_Timer
If BallSide=1 Then sw29.isdropped = 0 End If
If BallSide=2 Then sw25.isdropped = 0 End If
DropFix.Enabled = 1
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

Sub Table1_Exit
	If B2SOn Then Controller.Stop
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


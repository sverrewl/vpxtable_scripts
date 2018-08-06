'############################################################################################
'############################################################################################
'#######                                                                             ########
'#######          Zira                                                               ########
'#######          (Playmatic 1980)                                                   ########
'#######                                                                             ########
'############################################################################################
'############################################################################################
' Version 1.0 FS mfuegemann 2017
'
' Thanks to:
' Akiles for providing the plastics and playfield images
'
' Version 1.1:
' - remove Williams Flipper decal
' - adjust star trigger animation speed to avoid stuck triggers
' - changed the Upper Slingshot switch to #39
' - adjusted friction values


Option Explicit
'First, try to load the Controller.vbs (DOF), which helps controlling additional hardware like lights, gears, knockers, bells and chimes (to increase realism)
'This table uses DOF via the 'SoundFX' calls that are inserted in some of the PlaySound commands, which will then fire an additional event, instead of just playing a sample/sound effect
On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the Controller.vbs file in order to run this table (installed with the VPX package in the scripts folder)"
On Error Goto 0
'If using Visual PinMAME (VPM), place the ROM/game name in the constant below,
'both for VPM, and DOF to load the right DOF config from the Configtool, whether it's a VPM or an Original table

Const cGameName = "zira"
LoadVPM "01560000","play2.VBS",3.2

Const UseSolenoids=1,UseLamps=1,SSolenoidOn="SolOn",SSolenoidOff="SolOff",SFlipperOn="FlipperUp",SFlipperOff="FlipperDown",SCoin="coin3"
Dim KeySelfTestValue

'-----------------------------------
'------  Global Cofigurations ------
'-----------------------------------
BallSize = 25  					'Ball radius
Const DimGI = 2					'Dim GI intensity, 10 is the base value
Const VolumeMultiplier = 1.5   	'adjusts table sound volume
Const FreePlay = True			'Insert coin on GameStart if True
KeySelfTestValue = LeftMagnaSave		'defines the settings key, select with StartGameKey, cycle through to 20 to exit menu (A-key = 30, if You want another keyboard key)



'--------------------------------------------
'------  Solenoid Assignment Playmatic ------
'--------------------------------------------

SolCallback(1)="bsRightHole.SolOut"				'OK
SolCallback(2)="DropTargetBank2.soldropup"		'OK
'SolCallback(3)="Sol3"
SolCallback(4)="DropTargetBank1.soldropup"		'OK
SolCallback(5)="bsTrough.SolOut"				'OK
SolCallback(6)="vpmSolSound SoundFX(""Knocker"",DOFKnocker),"		'OK	
SolCallback(7)="Sol7"							'OK Captive Ball Post
'SolCallback(8)="Sol8"							'Game On - not stable

'SolCallback(sLLFlipper)="vpmSolFlipper LeftFlipper,ULeftFlipper,"
'SolCallback(sLRFlipper)="vpmSolFlipper RightFlipper,Nothing,"

'GameOn
Sub Sol_GameOn(enabled)
	Flipperactive = enabled
	VpmNudge.SolGameOn(enabled)
	if not Flipperactive then
		ULeftFlipper.Rotatetostart
		LeftFlipper.Rotatetostart
		RightFlipper.Rotatetostart
	end if
End Sub

Sub Sol7(enabled)
	if enabled then
		cPost.isdropped = true
	else
		cPost.timerenabled = True
	end if
End Sub

Sub cPost_Timer
	cPost.timerenabled = False
	cPost.isdropped = False
End Sub


If Zira.ShowDT = false then

End If

'--------------------------
'------  Table Init  ------
'--------------------------
Dim bsTrough,obj,bsRightHole,cCaptive,DropTargetBank1,DropTargetBank2,Flipperactive

Sub Zira_Init
	vpminit me
	
'	Flipperactive = True

	CaptiveKicker1.createBall
	CaptiveKicker1.kick 180,2
	CaptiveKicker2.createBall
	CaptiveKicker2.kick 180,2
	CaptiveKicker3.createBall
	CaptiveKicker3.kick 180,2
	CaptiveKicker4.createBall
	CaptiveKicker4.kick 180,2


    Controller.GameName=cGameName
    Controller.SplashInfoLine="Zira" & vbNewLine & "created by mfuegemann"
    Controller.HandleKeyboard=False
    Controller.ShowTitle=0
    Controller.ShowFrame=0
    Controller.ShowDMDOnly=1
	'Controller.Hidden = 1			'enable to hide DMD if You use a B2S backglass

'    'DMD position for 3 Monitor Setup
'    'Controller.Games(cGameName).Settings.Value("dmd_pos_x")=3850		'set this to 0 if You cannot find the DMD
'    'Controller.Games(cGameName).Settings.Value("dmd_pos_y")=300		'set this to 0 if You cannot find the DMD
'    'Controller.Games(cGameName).Settings.Value("dmd_width")=505
'    'Controller.Games(cGameName).Settings.Value("dmd_height")=155
'    'Controller.Games(cGameName).Settings.Value("rol")=0				
'
'	'Controller.Games(cGameName).Settings.Value("ddraw") = 0             'set to 0 if You have problems with DMD showing or table stutter
		   
    Controller.HandleMechanics=0

	Controller.SolMask(0)=0
	vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'"

    Controller.Run 
    If Err Then MsgBox Err.Description
    On Error Goto 0

    PinMAMETimer.Interval=PinMAMEInterval
	PinMAMETimer.Enabled = true

    vpmNudge.TiltSwitch=7
    vpmNudge.Sensitivity=5
	vpmNudge.TiltObj = Array(LeftFlipper,ULeftFlipper,RightFlipper) 

    vpmMapLights AllLights

    Set bsTrough=New cvpmBallStack
        bsTrough.InitSw 0,5,0,0,0,0,0,0 		'0.1 = Switch 0
        bsTrough.InitKick BallRelease,90,5
        bsTrough.InitExitSnd SoundFX("BallRel",DOFContactors),SoundFX("Solenoid",DOFContactors)
        bsTrough.Balls=1

	set DropTargetBank1 = new cvpmDropTarget
		DropTargetBank1.InitDrop Array(DT21,DT22,DT23,DT24,DT25,DT26,DT27), Array(21,22,23,24,25,26,27)
		DropTargetBank1.InitSnd SoundFX("Targetdrop1",DOFContactors),SoundFX("TargetBankreset1",DOFContactors)
		DropTargetBank1.CreateEvents "DropTargetBank1"

	set DropTargetBank2 = new cvpmDropTarget
		DropTargetBank2.InitDrop Array(DT35,DT36,DT37,DT38), Array(35,36,37,38)
		DropTargetBank2.InitSnd SoundFX("Targetdrop1",DOFContactors),SoundFX("TargetBankreset1",DOFContactors)
		DropTargetBank2.CreateEvents "DropTargetBank2"

	Set bsRightHole = New cvpmSaucer
		bsRightHole.InitSounds SoundFX("kicker_enter_center",DOFContactors),SoundFX("Popper",DOFContactors),SoundFX("Solenoid",DOFContactors)
		bsRightHole.initkicker RHole,44,185,10,0
		bsRightHole.InitExitVariance 5,2
	

	For each obj in GI
		obj.intensity = obj.intensity + DimGI
	Next
End Sub

Sub Zira_Exit()
	Controller.Pause = False
	Controller.Stop
End Sub


'------------------------------
'------  Trough Handler  ------
'------------------------------
Sub Drain_Hit()
	PlaySound "drain",0,1,AudioPan(Drain),0.25,0,0,1,AudioFade(Drain)
	bsTrough.AddBall Me
	Sol_GameOn False
End Sub

'-------------------------------
'------  Keybord Handler  ------
'-------------------------------

Sub Zira_KeyDown(ByVal keycode)
	If keycode = PlungerKey Then
		Plunger.PullBack
		PlaySound "plungerpull",0,1,AudioPan(Plunger),0.25,0,0,1,AudioFade(Plunger)
	End If

	If keycode = LeftFlipperKey Then
		controller.switch(84) = 1
		if Flipperactive then
			LeftFlipper.RotateToEnd
			ULeftFlipper.RotateToEnd
			PlaySound SoundFX("fx_flipperup",DOFFlippers), 0, .67, AudioPan(LeftFlipper), 0.05,0,0,1,AudioFade(LeftFlipper)
		end If
	End If

	If keycode = RightFlipperKey Then
		controller.switch(82) = 1
		if Flipperactive then
			RightFlipper.RotateToEnd
			PlaySound SoundFX("fx_flipperup",DOFFlippers), 0, .67, AudioPan(RightFlipper), 0.05,0,0,1,AudioFade(RightFlipper)
		end If
	End If

	if (Keycode = StartGamekey) and FreePlay then
		vpmtimer.pulsesw 1
	end If

	if Keycode = KeySelfTestValue Then
		vpmtimer.pulsesw 6
	end If

	If vpmKeyDown(KeyCode) Then Exit Sub
End Sub

Sub Zira_KeyUp(ByVal keycode)
	If keycode = PlungerKey Then
		Plunger.Fire
		PlaySound "plunger",0,1,AudioPan(Plunger),0.25,0,0,1,AudioFade(Plunger)
	End If

	If keycode = LeftFlipperKey Then
		controller.switch(84) = 0
		LeftFlipper.RotateToStart
		ULeftFlipper.Rotatetostart
		if Flipperactive then
			PlaySound SoundFX("fx_flipperdown",DOFFlippers), 0, 1, AudioPan(LeftFlipper), 0.05,0,0,1,AudioFade(LeftFlipper)
		end If
	End If

	If keycode = RightFlipperKey Then
		controller.switch(82) = 0
		RightFlipper.RotateToStart
		if Flipperactive then
			PlaySound SoundFX("fx_flipperdown",DOFFlippers), 0, 1, AudioPan(RightFlipper), 0.05,0,0,1,AudioFade(RightFlipper)
		end If
	End If

	If vpmKeyUp(KeyCode) Then Exit Sub
End Sub


'------------------------------
'------  Switch Handler  ------
'------------------------------
Sub Trigger1_Hit:Sol_GameOn True:End Sub

Sub RHole_Hit:bsRightHole.addball 0:End Sub

Sub Bumper1_Hit:vpmTimer.PulseSw 45:PlaySound SoundFX("Jet2",DOFContactors),0,1,-0.6,0.05:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 46:PlaySound SoundFX("Jet1",DOFContactors),0,1,-0.5,0.05:End Sub


sub SW11_hit:Controller.Switch(11)=1:End Sub
sub SW11_unhit:Controller.Switch(11)=0:End Sub
sub SW12_hit:Controller.Switch(12)=1:End Sub
sub SW12_unhit:Controller.Switch(12)=0:End Sub
sub SW13_hit:Controller.Switch(13)=1:End Sub
sub SW13_unhit:Controller.Switch(13)=0:End Sub
sub SW15_hit:Controller.Switch(15)=1:End Sub
sub SW15_unhit:Controller.Switch(15)=0:End Sub
sub SW16_hit:Controller.Switch(16)=1:End Sub
sub SW16_unhit:Controller.Switch(16)=0:End Sub
sub SW17_hit:Controller.Switch(17)=1:End Sub
sub SW17_unhit:Controller.Switch(17)=0:End Sub
sub SW18_hit:Controller.Switch(18)=1:End Sub
sub SW18_unhit:Controller.Switch(18)=0:End Sub

sub SW28_hit:Controller.Switch(28)=1:End Sub
sub SW28_unhit:Controller.Switch(28)=0:End Sub
sub SW31_hit:Controller.Switch(31)=1:End Sub
sub SW31_unhit:Controller.Switch(31)=0:End Sub
sub SW34_hit:Controller.Switch(34)=1:End Sub
sub SW34_unhit:Controller.Switch(34)=0:End Sub

sub SW47_hit:Controller.Switch(47)=1:End Sub
sub SW47_unhit:Controller.Switch(47)=0:End Sub
sub SW48_hit:Controller.Switch(48)=1:End Sub
sub SW48_unhit:Controller.Switch(48)=0:End Sub

sub T32a_hit:vpmTimer.PulseSw 32:End Sub
sub T32b_hit:vpmTimer.PulseSw 32:End Sub

Sub Spinner14_spin:vpmTimer.PulseSw 14:PlaySound "fx_spinner", 0, 0.25,-0.8,0.5:End Sub

Sub SW42a_Hit:vpmTimer.PulseSw 42:End Sub
Sub SW42b_Hit:vpmTimer.PulseSw 42:End Sub
Sub SW42c_Hit:vpmTimer.PulseSw 42:End Sub
Sub SW43a_Hit:vpmTimer.PulseSw 43:End Sub
Sub SW43b_Hit:vpmTimer.PulseSw 43:End Sub
Sub SW43c_Hit:vpmTimer.PulseSw 43:End Sub

Sub LampTimer_Timer
	B1.state = abs(Controller.switch(37) and Controller.switch(38))  	'Red Bumper
	B1a.state = abs(Controller.switch(37) and Controller.switch(38))
	B1b.state = abs(Controller.switch(37) and Controller.switch(38))
	LightA.state = abs(Controller.switch(37) and Controller.switch(38))		'Upper Rollover

	B2.state = abs(Controller.switch(35) and Controller.switch(36))	'Blue Bumper
	B2a.state = abs(Controller.switch(35) and Controller.switch(36))
	B2b.state = abs(Controller.switch(35) and Controller.switch(36))
	LightB.state = abs(Controller.switch(35) and Controller.switch(36))		'Lower Rollover

	LightC.state = abs(Controller.switch(24))	'SW34 captive lane
	LightD.state = abs(Controller.switch(24))	'SW14 spinner
End Sub


Dim CBall
Sub CaptiveInit_Hit
	Set Cball = ActiveBall
	CaptiveCenter.Timerenabled = True
End Sub

Sub CaptiveCenter_Timer
	if cball.y > CaptiveCenter.y then
		cball.y = CaptiveCenter.y
	end If
End Sub


Dim GateSpeed
Sub ReleaseGateOpen_Hit
	if not ReleaseGateOpen.Timerenabled then
		GateSpeed = 1.2
		ReleaseGateOpen.Timerenabled = True
	end if
End Sub

Sub ReleaseGateOpen_Timer
	P_ReleaseGate.rotz = P_ReleaseGate.rotz + GateSpeed
	if P_ReleaseGate.rotz > 21 then
		GateSpeed = -1.2
	end if
	if P_ReleaseGate.rotz <= 0 then
		P_ReleaseGate.rotz = 0
		ReleaseGateOpen.Timerenabled = False
	end if
end Sub


'############################################################################################
'############################################################################################
'*****GI Lights On
dim xx

For each xx in GI:xx.State = 1: Next

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
	vpmTimer.PulseSw 41
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
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub UpperSlingShot_Slingshot
	vpmTimer.PulseSw 39
    PlaySound SoundFX("right_slingshot",DOFContactors), 0,1, 0.05,0.05 '0,1, AudioPan(UpperSlingShot), 0.05,0,0,1,AudioFade(UpperSlingShot)
    USling.Visible = 0
    USling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    UpperSlingShot.TimerEnabled = 1
End Sub

Sub UpperSlingShot_Timer
    Select Case LStep
        Case 3:USLing1.Visible = 0:USLing2.Visible = 1:sling2.TransZ = -10
        Case 4:USLing2.Visible = 0:USLing.Visible = 1:sling2.TransZ = 0:UpperSlingShot.TimerEnabled = 0
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

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "Zira" is the name of the table
	Dim tmp
    tmp = tableobj.y * 2 / Zira.height-1
    If tmp > 0 Then
		AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "Zira" is the name of the table
    Dim tmp
    tmp = tableobj.x * 2 / Zira.width-1
    If tmp > 0 Then
        AudioPan = Csng(tmp ^10)
    Else
        AudioPan = Csng(-((- tmp) ^10) )
    End If
End Function

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 2000) * VolumeMultiplier
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
	FlipperULSh.RotZ = ULeftFlipper.currentangle
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
        If BOT(b).X < Zira.Width/2 Then
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Zira.Width/2))/7)) + 6
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Zira.Width/2))/7)) - 6
        End If
        ballShadow(b).Y = BOT(b).Y + 12
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
		' no shadow into plunger lane
		If (BOT(b).X > 1030) and (BOT(b).Y > 720) Then
            BallShadow(b).visible = 0
        Else
            BallShadow(b).visible = 1
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
	PlaySound "target", 0, 0.5, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
	PlaySound "metalhit_thin", 0, 0.2, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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
	PlaySound "fx_spinner", 0, 1, AudioPan(Spinner), 0.25, 0, 0, 1, AudioFade(Spinner)
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

Sub ULeftFlipper_Collide(parm)
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


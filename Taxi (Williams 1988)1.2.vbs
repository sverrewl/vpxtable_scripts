'############################################################################################
'############################################################################################
'#######                                                                             ########
'#######          Taxi                                                               ########
'#######          (Williams 1988)                                                    ########
'#######                                                                             ########
'############################################################################################
'############################################################################################

Option Explicit

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' No special SSF tweaks yet.

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

'-----------------------------------
'------  Global Cofigurations ------
'-----------------------------------
Const ColorMod = 1 '0= Regular lighting 1= Color Mod lighting


Const cGameName = "taxi_l4"
Const UseSolenoids = 1
Const UseLamps = 0
Const UseSync = 1

' Standard Sounds
Const SSolenoidOn = ""
Const SSolenoidOff = ""
Const SFlipperOn = ""
Const SFlipperOff = ""
Const SCoin = "Coin"

LoadVPM "01560000", "S11.VBS", 3.26

Dim DesktopMode: DesktopMode = Taxi.ShowDT

If DesktopMode = True Then 'Show Desktop components
	Ramp16.visible=1
	Ramp15.visible=1
	SideWood.visible=1
	SideWood1.visible=0
	BackdropCover.visible=1
	BackdropCover1.visible=0
Else
	Ramp16.visible=0
	Ramp15.visible=0
	SideWood.visible=0
	SideWood1.visible=1
	BackdropCover.visible=0
	BackdropCover1.visible=1
End if


'-----------------------------------
'------  Solenoid Assignment  ------
'-----------------------------------
SolCallback(1) = "bsTrough.SolIn"
SolCallback(2) = "bsTrough.SolOut"
SolCallback(3) = "SolCatapult"
SolCallback(4) = "CenterDTBank.SolDropUp"
SolCallback(5) = "bsTopHole.SolOut"
SolCallback(6) = "RightDTBank.SolDropUp"
SolCallback(7) = "SolSpinoutKicker"
SolCallback(8) = "bsRightHole.SolOut"
Solcallback(9) = "TopGateLeft.open ="
SolCallback(10) = "Sol10" 'Insert Gen Illum Relay
SolCallback(11) = "Sol11" 'Playfield Gen Illumination
'SolCallback(12) = 'A/C Select Relay
SolCallback(13) = "vpmSolSound SoundFX(""ringing_bell"",DOFBell),"
SolCallback(14) = "vpmSolSound SoundFX(""knocker"",DOFKnocker),"
'SolCallback(15) = 'JACKPOT Flasher - handled in LampTimer
SolCallback(16) = "SetLamp 116,"	'JOYRIDE Flasher - handled in LampTimer too
'SolCallback(17) = 'Bumper
'SolCallback(18) = 'Sling
'SolCallback(19) = 'Bumper
'SolCallback(20) = 'Sling
'SolCallback(21) = 'Bumper
SolCallback(23) = "SolGameOn"
SolCallback(25) = "SetLamp 125,"   'Pinbot Insert - handled in LampTimer
SolCallback(26) = "SetLamp 126,"   'Drac Insert   - handled in LampTimer
SolCallback(27) = "SetLamp 127,"   'Lola Insert   - handled in LampTimer
SolCallback(28) = "SetLamp 128,"   'Santa Insert  - handled in LampTimer
SolCallback(29) = "SetLamp 129,"   'Gorby Insert  - handled in LampTimer
SolCallback(30) = "SetLamp 130,"	'Left Ramp Flasher
SolCallback(31) = "SetLamp 131,"	'Right Ramp Flasher
SolCallback(32) = "SetLamp 132,"	'Spinout Flasher

Dim FlipperActive
FlipperActive = False
Sub SolGameOn(enabled)
	FlipperActive = enabled
	VpmNudge.SolGameOn(enabled)
	if not FlipperActive then
		RightFlipper.RotateToStart
		LeftFlipper.RotateToStart
	end if
End Sub

'Catapult
Sub SolCatapult(enabled)
	if enabled then
		FireCatapult
		if Controller.Switch(35) then
			bsCatapult.ExitSol_On
		end If
	end if
End Sub

Dim CatapultDir
Sub FireCatapult
	catapultLaunchKicker.Timerenabled = False
	CatapultDir = 1.5
	catapultLaunchKicker.Timerenabled = True
End Sub

Sub catapultLaunchKicker_Timer
	P_Catapult.rotx = P_Catapult.rotx + CatapultDir
	if P_Catapult.rotx > 90 then
		P_Catapult.rotx = 90
		CatapultDir = -0.5
	end if
	if P_Catapult.rotx < 0 then
		catapultLaunchKicker.Timerenabled = False
		P_Catapult.rotx = 0
		CatapultDir = 0
	end if
end Sub

'Spinout Kicker
Sub SolSpinoutKicker(enabled)
	if enabled then
		bsSpinOutKicker.ExitSol_On
		P_Spinout.TransZ = -60
		SpinOutKicker.Timerenabled = True
	end if
End Sub

Sub SpinOutKicker_Timer
	P_Spinout.TransZ = P_Spinout.TransZ + 1
	if P_Spinout.TransZ >= 0 then
		SpinOutKicker.Timerenabled = False
		P_Spinout.TransZ = 0
	end if
end Sub

' GI
Sub Sol10(enabled)
	SetLamp 99,enabled	'#99 used for GI Inserts
End Sub

dim obj
Sub Sol11(enabled)		'inverse action, GI is Off if Solenoid is enabled
	if ColorMod then
		for Each obj in GI
			obj.state = not enabled
		next
	else
		for Each obj in GI_mono
			obj.state = not enabled
		next
	end if
	if enabled Then
		if ColorMod then
			Primitive_Ramp2.image = "NewRenderMapB_final1b_BLUE_OFF"
		else
			Primitive_Ramp2.image = "NewRenderMapB_final1b_OFF"
		end If
		Flasher1.visible = False
		Flasher2.visible = False
	Else
		if ColorMod then
			Primitive_Ramp2.image = "NewRenderMapB_final1b_BLUE"
		else
			Primitive_Ramp2.image = "NewRenderMapB_final1"
		end If
		Flasher1.visible = True
		Flasher2.visible = True
	end If
End Sub

'--------------------------
'------  Table Init  ------
'--------------------------
Dim bsTrough,bsTopEject,bsRightHole,bsTopHole,bsCatapult,bsSpinoutKicker,RightDTBank,CenterDTBank
Sub Taxi_Init()
	Flasher1.visible = True
	Flasher2.visible = True
	if ColorMod then
		Primitive_Ramp2.image = "NewRenderMapB_final1b_BLUE"
		for Each obj in GI
			obj.state = True
		next
		for Each obj in GI_mono
			obj.state = False
		next
		F16.state = 0
		F30.state = 0
		F31.state = 0
		F32.state = 0
	else
		Primitive_Ramp2.image = "NewRenderMapB_final1"
		for Each obj in GI
			obj.state = False
		next
		for Each obj in GI_mono
			obj.state = True
		next
		F16c.state = 0
		F30c.state = 0
		F31c.state = 0
		F32c.state = 0
	end If

	Dim DesktopMode: DesktopMode = Taxi.ShowDT
	If DesktopMode = True Then 'Show Desktop components
		for Each obj in Backboxlights
			obj.state = Lightstateon
		next
		LEDTimer.enabled = True
	Else
		for Each obj in Backboxlights
			obj.state = Lightstateoff
		next
		LEDTimer.enabled = False
	End if

	vpmInit Me
    With Controller
		.GameName = cGameName
        'If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Taxi"
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = 0
 		.Dip(0) = &H00


	    'DMD position for 3 Monitor Setup
		Controller.Games(cGameName).Settings.Value("dmd_pos_x")=0
		Controller.Games(cGameName).Settings.Value("dmd_pos_y")=0
		'Controller.Games(cGameName).Settings.Value("dmd_width")=505
		'Controller.Games(cGameName).Settings.Value("dmd_height")=155
		'Controller.Games(cGameName).Settings.Value("rol")=0

        On Error Resume Next
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
        On Error Goto 0
	End With

    On Error Goto 0

	' Nudging
	vpmNudge.TiltSwitch = 1
	vpmNudge.Sensitivity = 1
	vpmNudge.TiltObj = Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingshot)

    ' Trough handler
    Set bsTrough = New cvpmBallStack
    bsTrough.InitSw 10,11,12,0,0,0,0,0
	bsTrough.InitKick BallRelease,60,8
	bsTrough.InitEntrySnd SoundFX("BallRelease",DOFContactors),SoundFX("Solon",DOFContactors)
	bsTrough.InitExitSnd SoundFX("BallRelease",DOFContactors), SoundFX("Solon",DOFContactors)
	bsTrough.Balls=2

 	'Right Hole
    set bsRightHole = new cvpmSaucer
	bsRightHole.InitKicker RightHole,36,180,8,0
	bsRightHole.InitExitVariance 5,2
	bsRightHole.InitSounds SoundFX("kicker_enter_center",DOFContactors),SoundFX("solon",DOFContactors),SoundFX("popper_ball",DOFContactors)

	'Top Hole - Joy Ride
    set bsTopHole = new cvpmSaucer
	bsTopHole.InitKicker TopHole,13,160,8,0
	bsTopHole.InitExitVariance 2, 2
	bsTopHole.InitSounds SoundFX("kicker_enter_center",DOFContactors),SoundFX("solon",DOFContactors),SoundFX("popper_ball",DOFContactors)

	'Catapult
    Set bsCatapult = New cvpmSaucer
    bsCatapult.InitKicker CatapultLaunchKicker,35,0,40,0
    bsCatapult.InitSounds SoundFX("kicker_enter_center",DOFContactors),SoundFX("solon",DOFContactors),SoundFX("popper_ball",DOFContactors)

	'SpinOut Kicker
    Set bsSpinoutKicker = New cvpmSaucer
    bsSpinoutKicker.InitKicker SpinoutKicker,43,0,35,0
	bsSpinoutKicker.InitExitVariance 5,2
	bsSpinoutKicker.InitSounds SoundFX("soloff",DOFContactors),SoundFX("solon",DOFContactors),SoundFX("popper_ball",DOFContactors)

    Set RightDTBank = New cvpmDropTarget
    RightDTBank.InitDrop Array(DT30,DT31,DT32),Array(30,31,32)
    RightDTBank.InitSnd SoundFX("droptarget",DOFDropTargets),SoundFX("fx_resetdrop",DOFContactors)
    RightDTBank.CreateEvents "RightDTBank"

    Set CenterDTBank = New cvpmDropTarget
    CenterDTBank.InitDrop Array(DT27,DT28,DT29),Array(27,28,29)
    CenterDTBank.InitSnd SoundFX("droptarget",DOFDropTargets),SoundFX("fx_resetdrop",DOFContactors)
    CenterDTBank.CreateEvents "CenterDTBank"

    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

	vpmtimer.PulseSW 24
End Sub



'------------------------------
'------  Switch Handler  ------
'------------------------------
Sub SpinoutKicker_Hit:bsSpinOutKicker.addball Me:End Sub
Sub RightHole_Hit:bsRightHole.addball Me:End Sub
Sub TopHole_Hit:bsTopHole.addball Me:End Sub
Sub CatapultKicker_Hit:bsCatapult.addball Me:End Sub

Sub Bumper1_Hit:vpmTimer.PulseSw 21:Playsound SoundFX("fx_bumper3",DOFContactors),0,1,0,0.1:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 17:Playsound SoundFX("fx_bumper2",DOFContactors),0,1,-0.1,0.1:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 19:Playsound SoundFX("fx_bumper4",DOFContactors),0,1,0.1,0.1:End Sub

Sub SW14_Hit:Controller.Switch(14) = True:End Sub
Sub SW14_Unhit:Controller.Switch(14) = False:End Sub
Sub SW15_Hit:Controller.Switch(15) = True:End Sub
Sub SW15_Unhit:Controller.Switch(15) = False:End Sub
Sub SW16_Hit:Controller.Switch(16) = True:End Sub
Sub SW16_Unhit:Controller.Switch(16) = False:End Sub

Sub SW22_Hit:Controller.Switch(22) = True:End Sub
Sub SW22_Unhit:Controller.Switch(22) = False:End Sub
Sub SW23_Hit:Controller.Switch(23) = True:End Sub
Sub SW23_Unhit:Controller.Switch(23) = False:End Sub
Sub Target24_Hit:vpmtimer.PulseSW 24:End Sub
Sub SW25_Hit:vpmtimer.PulseSW 25:End Sub
Sub SW26_Hit:vpmtimer.PulseSW 26:End Sub

Sub SW33_Hit:Controller.Switch(33) = True:Primitive_SwitchArm33.objrotz = -15:End Sub
Sub SW33_Unhit:Controller.Switch(33) = False:SW33.Timerenabled = True:End Sub
Sub SW34_Hit:Controller.Switch(34) = True:Primitive_SwitchArm34.objrotz = 18:End Sub
Sub SW34_Unhit:Controller.Switch(34) = False:SW34.Timerenabled = True:End Sub
Sub SW37_Hit:Controller.Switch(37) = True:End Sub
Sub SW37_Unhit:Controller.Switch(37) = False:End Sub
Sub SW38_Hit:Controller.Switch(38) = True:End Sub
Sub SW38_Unhit:Controller.Switch(38) = False:End Sub
Sub SW39_Hit:Controller.Switch(39) = True:End Sub
Sub SW39_Unhit:Controller.Switch(39) = False:End Sub
Sub SW40_Hit:Controller.Switch(40) = True:End Sub
Sub SW40_Unhit:Controller.Switch(40) = False:End Sub
'SW43 handled via saucer
Sub SW44_Hit:vpmtimer.PulseSW 44:playsound "fx_sensor" End Sub


Sub SW33_Timer
	Primitive_SwitchArm33.objrotz = Primitive_SwitchArm33.objrotz + 3
	if Primitive_SwitchArm33.objrotz >= 0 then
		SW33.Timerenabled = False
		Primitive_SwitchArm33.objrotz = 0
	end If
End Sub
Sub SW34_Timer
	Primitive_SwitchArm34.objrotz = Primitive_SwitchArm34.objrotz - 3
	if Primitive_SwitchArm34.objrotz <= 0 then
		SW34.Timerenabled = False
		Primitive_SwitchArm34.objrotz = 0
	end If
End Sub

'###########################################
Sub SpinoutHelper_Hit
	if ActiveBall.Velx > 9 then
		ActiveBall.Velx = 1.05 * ActiveBall.Velx
	end if
End Sub

Sub RampHelper1_Hit
	ActiveBall.velZ = 0
End Sub
Sub RampHelper2_Hit
	ActiveBall.velZ = 0
End Sub


Sub Taxi_KeyDown(ByVal keycode)
	If keycode = PlungerKey Then
		Plunger.PullBack
		PlaySound "plungerpull",0,1,0.25,0.25
	End If

	If keycode = LeftFlipperKey Then
		if FlipperActive then
			LeftFlipper.RotateToEnd
			PlaySound SoundFX("fx_flipperup",DOFFlippers), 0, .67, -0.05, 0.05
		end if
	End If

	If keycode = RightFlipperKey Then
		if FlipperActive then
			RightFlipper.RotateToEnd
			PlaySound SoundFX("fx_flipperup",DOFFlippers), 0, .67, 0.05, 0.05
		end if
	End If

	If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub Taxi_KeyUp(ByVal keycode)

	If keycode = PlungerKey Then
		Plunger.Fire
		PlaySound "plunger",0,1,0.25,0.25
	End If

	If keycode = LeftFlipperKey Then
		LeftFlipper.RotateToStart
		if FlipperActive then
			PlaySound SoundFX("fx_flipperdown",DOFFlippers), 0, 1, -0.05, 0.05
		end if
	End If

	If keycode = RightFlipperKey Then
		RightFlipper.RotateToStart
		if FlipperActive then
			PlaySound SoundFX("fx_flipperdown",DOFFlippers), 0, 1, 0.05, 0.05
		end if
	End If

	If vpmKeyUp(keycode) Then Exit Sub
End Sub

Sub Drain_Hit()
	PlaySound "drain",0,1,0,0.25
	bsTrough.AddBall Me
	'Drain.DestroyBall
End Sub

'-----------------------------------
' Flipper Primitives
'-----------------------------------
sub FlipperTimer_Timer()
	pleftFlipper.rotz=leftFlipper.CurrentAngle
	prightFlipper.rotz=rightFlipper.CurrentAngle

	Primitive_LGateWire.rotx = LGate.currentangle
	Primitive_RGateWire.rotx = RGate.currentangle
end sub



'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
	vpmtimer.pulsesw 20
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
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
	vpmtimer.pulsesw 18
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
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

'*****************************************
'	Ball Shadow
'*****************************************

Dim BallShadow
BallShadow = Array (BallShadow1, BallShadow2, BallShadow3, BallShadow4, BallShadow5, BallShadow6)

Sub BallShadowUpdate()
    Dim BOT, b, shadowZ
    BOT = GetBalls

	' render the shadow for each ball
    For b = 0 to UBound(BOT)
		BallShadow(b).X = BOT(b).X
		BallShadow(b).Y = BOT(b).Y + 10
		If BOT(b).Z > 0 and BOT(b).Z < 30 Then
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
	PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
	PlaySound SoundFX("target",DOFTargets), 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
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

Sub DTCheckTimer_Timer
	'Drop Target GI Lights
	if DT27.isdropped then
		Light22a.state = Light22.state
		Light26a.state = Light26.state
	Else
		Light22a.state = Lightstateoff
		Light26a.state = Lightstateoff
	end If
	if DT28.isdropped then
		Light22b.state = Light22.state
		Light26b.state = Light26.state
	Else
		Light22b.state = Lightstateoff
		Light26b.state = Lightstateoff
	end If
	if DT29.isdropped then
		Light22c.state = Light22.state
		Light26c.state = Light26.state
	Else
		Light22c.state = Lightstateoff
		Light26c.state = Lightstateoff
	end If

	if DT30.isdropped then
		Light27a.state = Light27.state
		Light28a.state = Light28.state
	Else
		Light27a.state = Lightstateoff
		Light28a.state = Lightstateoff
	end If
	if DT31.isdropped then
		Light27b.state = Light27.state
		Light28b.state = Light28.state
	Else
		Light27b.state = Lightstateoff
		Light28b.state = Lightstateoff
	end If
	if DT32.isdropped then
		Light27c.state = Light27.state
		Light28c.state = Light28.state
	Else
		Light27c.state = Lightstateoff
		Light28c.state = Lightstateoff
	end If
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
Dim ccount

InitLamps()             ' turn off the lights and flashers and reset them to the default parameters
LampTimer.Interval = 10 'lamp fading speed
LampTimer.Enabled = 1

Dim chgLamp, num, chg, ii
Sub LampTimer_Timer()
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp) Then
        For ii = 0 To UBound(chgLamp)
            LampState(chgLamp(ii, 0) ) = chgLamp(ii, 1)       'keep the real state in an array
			FadingLevel(chgLamp(ii, 0) ) = chgLamp(ii, 1) + 4 'actual fading step
		Next
    End If

	'Multi-LampStates
	LampState(43) = controller.Lamp(43) or controller.solenoid(16)	'Joyride
	LampState(44) = controller.Lamp(44) or controller.solenoid(15)	'Jackpot
	LampState(25) = controller.Lamp(25) or controller.solenoid(25) or controller.solenoid(10)	'Pinbot Insert
	LampState(26) = controller.Lamp(26) or controller.solenoid(26) or controller.solenoid(10)	'Drac Insert
	LampState(27) = controller.Lamp(27) or controller.solenoid(27) or controller.solenoid(10)	'Lola Insert
	LampState(28) = controller.Lamp(28) or controller.solenoid(28) or controller.solenoid(10)	'Santa Insert
	LampState(29) = controller.Lamp(29) or controller.solenoid(29) or controller.solenoid(10)	'Gorby Insert
    UpdateLamps
End Sub

Sub InitLamps()
    Dim x
    For x = 0 to 200
        LampState(x) = 0         ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4       ' used to track the fading state
        FlashSpeedUp(x) = 0.3    ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.15 ' slower speed when turning off the flasher
        FlashMax(x) = 1          ' the maximum value when on, usually 1
        FlashMin(x) = 0          ' the minimum value when off, usually 0
        FlashLevel(x) = 0        ' the intensity of the flashers, usually from 0 to 1
    Next
End Sub

Sub UpdateLamps
	If DesktopMode = True Then
		'Joy Ride
		NFadeLm 1, L1
		NFadeL 1, L1a
		NFadeLm 2, L2
		NFadeL 2, L2a
		NFadeLm 3, L3
		NFadeL 3, L3a
		NFadeLm 4, L4
		NFadeL 4, L4a
		NFadeLm 5, L5
		NFadeL 5, L5a

		'Backbox
		NFadeL 6, L6
		NFadeL 7, L7
		NFadeL 8, L8
		NFadeL 6, L6
		NFadeL 49, L49
		NFadeL 50, L50
		NFadeL 51, L51
		NFadeL 52, L52
		NFadeL 53, L53
		NFadeL 54, L54
		NFadeL 55, L55
		NFadeL 56, L56
		NFadeL 64, L64
	end If

	'Inserts
	NFadeLm 9, L9
	Flash 9, L9a
	NFadeLm 10, L10
	Flash 10, L10a
	NFadeLm 11, L11
	Flash 11, L11a
	NFadeLm 12, L12
	Flash 12, L12a
	NFadeLm 13, L13
	Flash 13, L13a
	NFadeLm 14, L14
	NFadeLm 14, L14b
	Flash 14, L14a
	NFadeLm 15, L15
	NFadeLm 15, L15b
	Flash 15, L15a
	NFadeLm 16, L16b
	NFadeLm 16, L16
	Flash 16, L16a
	NFadeLm 17, L17
	Flash 17, L17a
	NFadeLm 18, L18
	Flash 18, L18a
	NFadeLm 19, L19
	Flash 19, L19a
	NFadeLm 20, L20
	Flash 20, L20a
	NFadeLm 21, L21
	Flash 21, L21a
	NFadeLm 22, L22
	Flash 22, L22a
	NFadeLm 23, L23b
	NFadeLm 23, L23
	Flash 23, L23a
	NFadeLm 24, L24
	Flash 24, L24a
	NFadeLm 25, L25
	Flash 25, L25a
	NFadeLm 26, L26
	Flash 26, L26a
	NFadeLm 27, L27
	Flash 27, L27a
	NFadeLm 28, L28
	Flash 28, L28a
	NFadeLm 29, L29
	Flash 29, L29a
	NFadeLm 30, L30
	Flash 30, L30a
	NFadeLm 31, L31
	Flash 31, L31a
	NFadeLm 32, L32b
	NFadeLm 32, L32
	Flash 32, L32a
	NFadeLm 33, L33
	Flash 33, L33a
	NFadeLm 34, L34
	Flash 34, L34a
	NFadeLm 35, L35
	Flash 35, L35a
	NFadeLm 36, L36
	Flash 36, L36a
	NFadeLm 37, L37
	Flash 37, L37a
	NFadeLm 38, L38
	NFadeLm 38, L38b
	NFadeLm 38, L38c
	Flash 38, L38a
	NFadeLm 39, L39
	Flash 39, L39a
	NFadeLm 40, L40
	NFadeLm 40, L40b
	NFadeLm 40, L40c
	Flash 40, L40a
	NFadeLm 41, L41
	Flash 41, L41a
	NFadeLm 42, L42
	Flash 42, L42a
	NFadeLm 43, L43
	Flash 43, L43a
	NFadeLm 44, L44
	Flash 44, L44a
	NFadeLm 45, L45
	Flash 45, L45a
	NFadeLm 46, L46
	Flash 46, L46a
	NFadeLm 47, L47
	Flash 47, L47a
	NFadeLm 48, L48
	Flash 48, L48a

	NFadeL 57, L57
	NFadeL 58, L58
	NFadeL 59, L59
	NFadeL 60, L60
	NFadeL 61, L61
	NFadeL 62, L62
	NFadeL 63, L63

	'GI Flashers
	NFadeLm 25, F25
	NFadeLm 26, F26
	NFadeLm 27, F27
	NFadeLm 28, F28
	NFadeLm 29, F29

	'Flashers
	if ColorMod then
		NFadeL 116, F16c
		NFadeL 130, F30c
		NFadeL 131, F31c
		NFadeL 132, F32c
	else
		NFadeL 116, F16
		NFadeL 130, F30
		NFadeL 131, F31
		NFadeL 132, F32
	end if
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

Sub NFadeLmb(nr, object) ' used for multiple lights with blinking
    Select Case FadingLevel(nr)
        Case 4:object.state = 0
        Case 5:object.state = 2
    End Select
End Sub

Sub MFadeL(nr, object, state)
    Select Case FadingLevel(nr)
        Case 4:object.state = 0:If LampState(state) =1 Then SetLamp state, 1
        Case 5:object.state = 1
    End Select
End Sub


'Lights, Ramps & Primitives used as 4 step fading lights
'a,b,c,d are the images used from on to off

Sub FadeObj(nr, object, a, b, c, d)
	Select Case FadingLevel(nr)
		Case 2:object.image = d:FadingLevel(nr) = 0 'Off
		Case 3:object.image = c:FadingLevel(nr) = 2 'fading...
		Case 4:object.image = b:FadingLevel(nr) = 3 'fading...
		Case 5:object.image = a:FadingLevel(nr) = 1 'ON
 End Select
End Sub

Sub FadeObjm(nr, object, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:object.image = d
        Case 3:object.image = c
        Case 4:object.image = b
        Case 5:object.image = d
    End Select
End Sub

Sub NFadeObj(nr, object, a, b)
    Select Case FadingLevel(nr)
        Case 4:object.image = b:FadingLevel(nr) = 0 'off
        Case 5:object.image = a:FadingLevel(nr) = 1 'on
    End Select
End Sub

Sub NFadeObjm(nr, object, a, b)
    Select Case FadingLevel(nr)
        Case 4:object.image = b
        Case 5:object.image = a
    End Select
End Sub

' Flasher objects

Sub Flash(nr, object)
    Select Case FadingLevel(nr)
        Case 4 'off
            FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown(nr)
            If FlashLevel(nr) < FlashMin(nr) Then
                FlashLevel(nr) = FlashMin(nr)
                FadingLevel(nr) = 0 'completely off
            End if
            Object.IntensityScale = FlashLevel(nr)
        Case 5 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
            If FlashLevel(nr) > FlashMax(nr) Then
                FlashLevel(nr) = FlashMax(nr)
                FadingLevel(nr) = 1 'completely on
            End if
            Object.IntensityScale = FlashLevel(nr)
    End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it just sets the flashlevel
    Object.IntensityScale = FlashLevel(nr)
End Sub

'Ramp Helpers
Sub LHelp_Hit():Playsound "fx_ball_drop", 0, 1, -0.05, 0.05:end Sub
Sub RHelp_Hit():Playsound "fx_ball_drop", 0, 1, 0.05, 0.05:end Sub
Sub SpinOutHelp_Hit():Playsound "fx_ball_drop", 0, 1, 0.05, 0.05:end Sub
Sub SpinOutHelp1_Hit():Playsound "fx_ball_drop", 0, 1, 0.05, 0.05:end Sub

'Wire ramp sounds
Sub WireRampSound_Hit():Playsound "WireRamp", 0, 1, 0, 0.35:end Sub
Sub WireRampSound1_Hit():Playsound "WireRamp", 0, 1, 0, 0.35:end Sub
Sub WireRampSound2_Hit():Playsound "WireRamp", 0, 1, 0, 0.35:end Sub
Sub WireRampSound3_Hit():Playsound "WireRamp", 0, 1, 0, 0.35:end Sub


'Plastic ramp sounds
Sub RRHit0_Hit:PlaySound "fx_rr1", 0, 1, pan(ActiveBall):End Sub
Sub RRHit1_Hit:PlaySound "fx_rr2", 0, 1, pan(ActiveBall):End Sub
Sub RRHit2_Hit:PlaySound "fx_rr3", 0, 1, pan(ActiveBall):End Sub
Sub RRHit3_Hit:PlaySound "fx_rr4", 0, 1, pan(ActiveBall):End Sub
Sub RRHit4_Hit:PlaySound "fx_rr1", 0, 1, pan(ActiveBall):End Sub
Sub RRHit5_Hit:PlaySound "fx_rr2", 0, 1, pan(ActiveBall):End Sub
Sub RRHit6_Hit:PlaySound "fx_rr3", 0, 1, pan(ActiveBall):End Sub
Sub RRHit7_Hit:PlaySound "fx_lr4", 0, 1, pan(ActiveBall):End Sub
Sub RRHit8_Hit:PlaySound "fx_lr5", 0, 1, pan(ActiveBall):End Sub
Sub RRHit9_Hit:PlaySound "fx_lr5", 0, 1, pan(ActiveBall):End Sub
Sub RRHit10_Hit:PlaySound "fx_lr1", 0, 1, pan(ActiveBall):End Sub
Sub RRHit11_Hit:PlaySound "fx_lr2", 0, 1, pan(ActiveBall):End Sub
Sub RRHit12_Hit:PlaySound "fx_lr4", 0, 1, pan(ActiveBall):End Sub
Sub RRHit13_Hit:PlaySound "fx_lr5", 0, 1, pan(ActiveBall):End Sub

'--------------------------------------
'------  Destruk's Display Code  ------
'--------------------------------------
'16 - 14
'7  - 7
'16 - 7

Dim Digits(53)
Const Offset=32
Digits(Offset+0)=Array(LED1,LED2,LED3,LED4,LED5,LED6,LED7,LED8)'ok
Digits(Offset+1)=Array(LED9,LED10,LED11,LED12,LED13,LED14,LED15)'ok
Digits(Offset+2)=Array(LED16,LED17,LED18,LED19,LED20,LED21,LED22)'ok
Digits(Offset+3)=Array(LED23,LED24,LED25,LED26,LED27,LED28,LED29,LED30)'ok
Digits(Offset+4)=Array(LED31,LED32,LED33,LED34,LED35,LED36,LED37)'ok
Digits(Offset+5)=Array(LED38,LED39,LED40,LED41,LED42,LED43,LED44)'ok
Digits(Offset+6)=Array(LED45,LED46,LED47,LED48,LED49,LED50,LED51)'ok

Sub LEDTimer_Timer
	Dim ChgLED,ii,num,chg,stat,obj
	ChgLED = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
	If Not IsEmpty(ChgLED) Then
		For ii = 0 To UBound(chgLED)
			num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
			if (num >= Offset+0) and (num <= Offset+6) then
				For Each obj In Digits(num)
					If chg And 1 Then obj.State = stat And 1
					chg = chg\2 : stat = stat\2
				Next
			end if
		Next
	End If
End Sub


'Additional Sounds

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 30, -0.1, 0.15
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 30, 0.1, 0.15
End Sub

Sub RSling_hit: PlaySound "fx_rubber", 0, (Vol(ActiveBall)) /5, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub LSling_hit: PlaySound "fx_rubber", 0, (Vol(ActiveBall)) /5, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
'Sub Target24_hit: PlaySound SoundFX("target",DOFTargets) End Sub



'******************************
' Diverse Collection Hit Sounds
'******************************
'Borrowed from Uncle Willy's Bad Cats VPX TablesDirectory

Sub aRubber_Posts_Hit(idx):PlaySound "fx_rubber", 0, (Vol(ActiveBall)) /6, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aRubber_Pins_Hit(idx):PlaySound "fx_postrubber", 0, (Vol(ActiveBall)) /8, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aPlastics_Hit(idx): PlaySound SoundFX("fx_droptarget",DOFDropTargets), 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub Rubber15_hit: Playsound "fx_postrubber", 0, (Vol(ActiveBall)) /5, pan(ActiveBall), 0, Pitch(ActiveBall) End Sub

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
	BallShadowUpdate

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


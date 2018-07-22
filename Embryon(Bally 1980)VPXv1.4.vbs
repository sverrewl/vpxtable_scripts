'==============================================================================================='
'																								' 			  	 
'        	                 		      EMBRYON	     			                			'
'          	              		  		Bally (1980)            	 		                   	'
'		  				  http://www.ipdb.org/machine.cgi?id=783								'
'																								'
' 	  		 	 	  	Created by ICPjuggla, OldSkoolGamer and Herweh							'
'																								' 			  	 
'==============================================================================================='

' Thalamus 2018-07-20
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' No special SSF tweaks yet.

Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

' ***********************************************************************************************
' OPTIONS
' ***********************************************************************************************

' DMD rotation
Const cDMDRotation 				= 1					' 0 or 1 for a DMD rotation of 90°

' VPinMAME ROM name
Const cGameName 				= "embryon"			' enter string of valid ROM 
													' "embryond" is the 7-digit ROM

' flasher and GI on or off option
Const Flashers_ON				= 1					' 0 or 1 to disable or enable flasher
Const GI_ON						= 1					' 0 or 1 to disable or enable GI

' lights update interval and fading interval
Const Lights_Refresh_Interval	= 35 				' interval in milliseconds

' ball size
Const BallSize 					= 50				' sets the ball size

' ***********************************************************************************************
' OPTIONS END
' ***********************************************************************************************

' ===============================================================================================
' some general constants and variables
' ===============================================================================================
	
Const UseSolenoids 		= 1
Const UseLamps 			= False
Const UseGI 			= False
Const UseSync 			= False
Const HandleMech 		= False

Const SSolenoidOn 		= "SolOn"
Const SSolenoidOff 		= "SolOff"
Const SCoin 			= "Coin"
Const SKnocker 			= "Knocker"

Dim I, x, obj, bsTrough, plungerIM, dtCenter, dtLeft, dtRight, bsSaucer


' ===============================================================================================
' load game controller
' ===============================================================================================

LoadVPM "01130100", "Bally.VBS", 3.21

' ===============================================================================================
' solenoids
' ===============================================================================================

SolCallback(1)			= "dtRight.SolDropUp"
SolCallback(2)			= "dtLeft.SolDropUp"
SolCallback(3)			= "dtCenter.SolDropUp"
SolCallback(4)			= "SolSaucerEject"
SolCallback(5)      	= "SolTroughOut"
SolCallback(6)			= "SolKnocker"

SolCallback(20) 		= "SolRLFlipper"

SolCallback(sLLFlipper) = "SolLFlipper"
SolCallback(sLRFlipper) = "SolRFlipper"


' =====================================================================================================
' table events
' =====================================================================================================

Sub Table1_Init
	' table initialization
	vpmInit Me

	' do some controller settings
	On Error Resume Next
	With Controller
		.GameName 								= cGameName
		.SplashInfoLine							= "Embryon, Bally, 1980"
		.HandleKeyboard 						= False
		.HandleMechanics 						= False
		.ShowTitle								= False
		.ShowDMDOnly							= True
		.ShowFrame								= False
		.ShowTitle 								= False
		If B2SOn Then 
			.Hidden 							= True
		End If
		.Games(cGameName).Settings.Value("rol")	= cDMDRotation
		.Run GetPlayerHWnd
		If Err Then MsgBox Err.Description
	End With
	On Error Goto 0
		
	' basic pinmame timer
	PinMAMETimer.Interval	= PinMAMEInterval
	PinMAMETimer.Enabled	= True

	' nudging
	vpmNudge.TiltSwitch		= 7
	vpmNudge.Sensitivity	= 3
	vpmNudge.TiltObj 		= Array(Bumper1,Bumper2,Bumper3,Bumper4,LeftSlingshot,RightSlingshot)
 
	' ball stack for trough
	Set bsTrough = New cvpmBallStack
		bsTrough.InitSw 0, 8, 5, 0, 0, 0, 0, 0
		bsTrough.InitKick BallRelease, 75, 4
		bsTrough.InitExitSnd SoundFX("BallRelease",DOFContactors), SoundFX("SolOn",DOFContactors)
		bsTrough.Balls = 2

	' create some captive balls
	k28.CreateBall
	k3433.CreateBall
	k34.CreateBall
	k33.CreateBall
	k17.CreateBall
	
	' drop targets
	Set dtCenter = New cvpmDropTarget
		dtCenter.InitDrop Array(t20,t21,t22),Array(20,21,22)
		dtCenter.InitSnd SoundFX("drop_fall_c",DOFContactors),SoundFX("drop_reset_c",DOFContactors)
		dtCenter.CreateEvents "dtCenter"
	Set dtLeft = New cvpmDropTarget
		dtLeft.InitDrop Array(t25,t26,t27),Array(25,26,27)
		dtLeft.InitSnd SoundFX("drop_fall_l",DOFContactors),SoundFX("drop_reset_l",DOFContactors)
		dtLeft.CreateEvents "dtLeft"
	Set dtRight = New cvpmDropTarget
		dtRight.InitDrop Array(t17),Array(17)
		dtRight.InitSnd SoundFX("drop_fall_r",DOFContactors),SoundFX("drop_reset_r",DOFContactors)
		dtRight.CreateEvents "dtRight"

	' saucer
	Set bsSaucer = New cvpmBallStack
		bsSaucer.InitSaucer sw4, 4, 270, 7
		bsSaucer.InitExitSnd SoundFX("Ball_Bounce_Low",DOFContactors), SoundFX("SolOn",DOFContactors)


End Sub
 
Sub Table1_Exit()
	Controller.Stop
End Sub

Sub Table1_Paused
	Controller.Pause = True
End Sub
Sub Table1_UnPaused
	Controller.Pause = False
End Sub

Sub Table1_KeyDown(ByVal keycode)
	If vpmKeyDown(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Pullback
	If keycode = LeftTiltKey Then LeftNudge 80, 1, 20
    If keycode = RightTiltKey Then RightNudge 280, 1, 20
    If keycode = CenterTiltKey Then CenterNudge 0, 1, 25
	If keycode = RightFlipperKey Then Controller.Switch(3) = False
End Sub

Sub Table1_KeyUp(ByVal keycode)
	If vpmKeyUp(KeyCode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Fire
	If keycode = RightFlipperKey Then Controller.Switch(3) = True
	'If keycode = LeftFlipperKey Then LeftFlipper.RotateToStart
End Sub


' ===============================================================================================
' ball draining and release plus auto plunger
' ===============================================================================================

Sub Drain_Hit()
	ClearBallID
	PlaySound "Drain"
	bsTrough.AddBall Me
'	plunger.PullBack
End Sub

Sub SolTroughOut(Enabled)
	If Enabled Then
		If bsTrough.Balls > 0 Then
			SolGI False
			ManualRightTargetUp
			bsTrough.ExitSol_On
		End If
	End If
End Sub

Sub SolLFlipper(Enabled)
     If Enabled Then
		PlaySound SoundFX("flipperup",DOFContactors)
		LeftFlipper.RotateToEnd
     Else
		PlaySound SoundFX("flipperdown",DOFContactors)
		LeftFlipper.RotateToStart
     End If
 End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
		 PlaySound SoundFX("flipperup",DOFContactors)
		 RightFlipper.RotateToEnd
		 RightUFlipper.RotateToEnd
     Else
		 PlaySound SoundFX("flipperdown",DOFContactors)
		 RightFlipper.RotateToStart
		 RightUFlipper.RotateToStart
    End If
 End Sub 

Sub SolRLFlipper(Enabled)
     If Enabled Then
		 PlaySound SoundFX("flipperup",DOFContactors)
		 RightLFlipper.RotateToEnd
     Else
		 PlaySound SoundFX("flipperdown",DOFContactors)
		 RightLFlipper.RotateToStart
    End If
 End Sub  

dim MotorCallback
Set MotorCallback = GetRef("GameTimer")
Sub GameTimer
    UpdateFlipperLogos
End Sub

Sub UpdateFlipperLogo_Timer
	LFLogo.RotY  = LeftFlipper.CurrentAngle - 90
	RFLogo.RotY  = RightFlipper.CurrentAngle + 90
	RUFLogo.RotY = RightUFlipper.CurrentAngle + 90
	RLFLogo.RotY = RightLFlipper.CurrentAngle - 90

	Dim bulb
	for each bulb in GILights
	bulb.State = Light80.State
	Next

	Dim lamp
	for each lamp in MidFlash
	lamp.State = l11.State
	Next
End Sub 

' ===============================================================================================
' slingshots events and slingshot animation scripting from JPSalas
' ===============================================================================================

'Sub LeftSlingshot_Slingshot()
'	LeftSling1.IsDropped = False
'	If SlingShot_Sound_ON = 1 Then PlaySound "left_slingshot"
'	vpmTimer.PulseSw 36
'End Sub
'
'
'Sub RightSlingshot_Slingshot
'	RightSling1.IsDropped = False
'	If SlingShot_Sound_ON = 1 Then PlaySound "right_slingshot"
'	vpmTimer.PulseSw 35
'End Sub

Dim LeftSlingshot,RightSlingshot,Bumper1,Bumper2,Bumper3

 Sub LeftSlingShot_Slingshot
	Leftsling = True
 	PlaySound SoundFX("left_slingshot",DOFContactors)
	vpmTimer.PulseSw 36
  End Sub

Dim Leftsling:Leftsling = False

Sub LS_Timer()
	If Leftsling = True and Left1.ObjRotZ < -13 then Left1.ObjRotZ = Left1.ObjRotZ + 2
	If Leftsling = False and Left1.ObjRotZ > -26 then Left1.ObjRotZ = Left1.ObjRotZ - 2
	If Left1.ObjRotZ >= -13 then Leftsling = False
	If Leftsling = True and Left2.ObjRotZ > -219 then Left2.ObjRotZ = Left2.ObjRotZ - 2
	If Leftsling = False and Left2.ObjRotZ < -206 then Left2.ObjRotZ = Left2.ObjRotZ + 2
	If Left2.ObjRotZ <= -219 then Leftsling = False
	If Leftsling = True and Left3.TransZ > -23 then Left3.TransZ = Left3.TransZ - 4
	If Leftsling = False and Left3.TransZ < -0 then Left3.TransZ = Left3.TransZ + 4
	If Left3.TransZ <= -23 then Leftsling = False
End Sub

 Sub RightSlingShot_Slingshot
	Rightsling = True
 	PlaySound SoundFX("right_slingshot",DOFContactors)
	vpmTimer.PulseSw 35
  End Sub

 Dim Rightsling:Rightsling = False

Sub RS_Timer()
	If Rightsling = True and Right1.ObjRotZ > 13 then Right1.ObjRotZ = Right1.ObjRotZ - 2
	If Rightsling = False and Right1.ObjRotZ < 26 then Right1.ObjRotZ = Right1.ObjRotZ + 2
	If Right1.ObjRotZ <= 13 then Rightsling = False
	If Rightsling = True and Right2.ObjRotZ < 219 then Right2.ObjRotZ = Right2.ObjRotZ + 2
	If Rightsling = False and Right2.ObjRotZ > 206 then Right2.ObjRotZ = Right2.ObjRotZ - 2
	If Right2.ObjRotZ >= 219 then Rightsling = False
	If Rightsling = True and Right3.TransZ > -23 then Right3.TransZ = Right3.TransZ - 4
	If Rightsling = False and Right3.TransZ < -0 then Right3.TransZ = Right3.TransZ + 4
	If Right3.TransZ <= -23 then Rightsling = False
End Sub



' ===============================================================================================
' bumpers and rings
' ===============================================================================================

Sub Bumper1_Hit()
	PlayBumperSoundAndSetSwitch 40
End Sub

Sub Bumper2_Hit()
	PlayBumperSoundAndSetSwitch 39
End Sub

Sub Bumper3_Hit()
	PlayBumperSoundAndSetSwitch 38
End Sub

Sub Bumper4_Hit()
	PlayBumperSoundAndSetSwitch 37
End Sub

Sub PlayBumperSoundAndSetSwitch(nSwitchID)
		Select Case Int(Rnd*3)+1
			Case 1 : PlaySound SoundFX("bumper_1",DOFContactors)
			Case 2 : PlaySound SoundFX("bumper_2",DOFContactors)
			Case 3 : PlaySound SoundFX("bumper_3",DOFContactors)
		End Select
	vpmTimer.PulseSw nSwitchID
End Sub


' ===============================================================================================
' knocker
' ===============================================================================================

Sub SolKnocker(Enabled)
	If Enabled Then
		PlaySound SoundFX("knocker",DOFKnocker)
	End If
End Sub


' ===============================================================================================
' spinner
' ===============================================================================================

Sub sw18_Spin()
	vpmTimer.PulseSw 18
	Playsound "Gate2_low"
End Sub


' ===============================================================================================
' saucer
' ===============================================================================================

Sub sw4_Hit()
	bsSaucer.AddBall 0
	Playsound "ball_bounce_low"
End Sub
Sub SolSaucerEject(Enabled)
	If Enabled Then
		If bsSaucer.Balls > 0 Then
			bsSaucer.InitSaucer sw4, 4, 265+Rnd()*10, 6+Rnd()*8
			bsSaucer.ExitSol_On
		End If
	End If
End Sub


' ===============================================================================================
' captive balls
' ===============================================================================================

Const damping2 = 0.75
Const damping1 = 0.95

Sub trig3433_Hit()
	If ActiveBall.VelY < -2 Then
		PlayBallCollideSound ActiveBall.VelY
		Dim velX, velY
		velX = ActiveBall.VelX * damping2
		velY = Abs(ActiveBall.VelY) * damping2
		If velX < 0 Then
			k34.Kick 340, Abs(velX) * 0.2 + velY * 0.8
			k33.Kick 20, Abs(velX) * 0.1 + velY * 0.4
		Else
			k33.Kick 20, Abs(velX) * 0.2 + velY * 0.8
			k34.Kick 340, Abs(velX) * 0.1 + velY * 0.4
		End If
	End If
End Sub

Sub trig28_Hit()
	If t25.IsDropped And ActiveBall.VelY < -2 Then
		PlayBallCollideSound ActiveBall.VelY
		Dim velX, velY
		velX = Abs(ActiveBall.VelX) * damping1
		velY = Abs(ActiveBall.VelY) * damping1
		k28.Kick 325, velX * 0.2 + velY * 0.8
		trig28.TimerInterval = 200
		trig28.TimerEnabled  = True
		k28.Enabled 		 = False
	End If
End Sub
Sub trig28_Timer()
	trig28.TimerEnabled = False
	k28.Enabled 		= True	
End Sub

Sub trig17_Hit()
	If t17.IsDropped And ActiveBall.VelY < -2 Then
		PlayBallCollideSound ActiveBall.VelY
		Dim velX, velY
		velX = Abs(ActiveBall.VelX) * damping1
		velY = Abs(ActiveBall.VelY) * damping1
		k17.Kick 35, velX * 0.2 + velY * 0.8
		trig17.TimerInterval = 200
		trig17.TimerEnabled  = True
		k17.Enabled 		 = False
	End If
End Sub
Sub trig17_Timer()
	trig17.TimerEnabled = False
	k17.Enabled 		= True
End Sub

Sub PlayBallCollideSound(VelY)
	Dim soundNo
	soundNo = Int(Abs(VelY) / 4) - 1
	If soundNo < 0 Then soundNo = 0
	If soundNo > 9 Then soundNo = 9
	PlaySound "collide" & soundNo
End Sub


' ===============================================================================================
' drop down targets patch
' ===============================================================================================

Sub ManualRightTargetUp()
	If t17.IsDropped Then
		dtRight.SolDropUp True
	End If
End Sub


' ===============================================================================================
' stand up targets
' ===============================================================================================

InitStandUps

Sub InitStandUps()
'	t29a.IsDropped = True
'	t29b.IsDropped = True
'	t30a.IsDropped = True
'	t30b.IsDropped = True
'	t31a.IsDropped = True
'	t31b.IsDropped = True
'	t32a.IsDropped = True
'	t32b.IsDropped = True
'	t28a.IsDropped = True
'	t28b.IsDropped = True
'	t34a.IsDropped = True
'	t34b.IsDropped = True
'	t33a.IsDropped = True
'	t33b.IsDropped = True
'	t17a.IsDropped = True
'	t17b.IsDropped = True
End Sub

' stand-ups
Sub t29_Hit()
	HitStandup 29, t29
End Sub
Sub t29_Timer()
	MoveStandup t29, t29p
End Sub
Sub t30_Hit()
	HitStandup 30, t30
End Sub
Sub t30_Timer()
	MoveStandup t30, t30p
End Sub
Sub t31_Hit()
	HitStandup 31, t31
End Sub
Sub t31_Timer()
	MoveStandup t31, t31p
End Sub
Sub t32_Hit()
	HitStandup 32, t32
End Sub
Sub t32_Timer()
	MoveStandup t32, t32p
End Sub

' captived stand-ups
Sub sw28_Hit()
	If BallSpeed > 5 Then
		HitStandup 28, t28
	End If
End Sub
Sub t28_Timer()
	MoveStandup t28, t28p
End Sub
Sub sw34_Hit()
	If BallSpeed > 5 Then
		HitStandup 34, t34
	End If
End Sub
Sub t34_Timer()
	MoveStandup t34, t34p
End Sub
Sub sw33_Hit()
	If BallSpeed > 5 Then
		HitStandup 33, t33
	End If
End Sub
Sub t33_Timer()
	MoveStandup t33, t33p
End Sub

Sub t17x_Hit()
	If BallSpeed > 5 Then
		HitStandup 17, t17x
		ManualRightTargetUp
	End If
End Sub
Sub t17x_Timer()
	MoveStandup t17x, t1
End Sub


Sub HitStandup(no, t)
	vpmTimer.PulseSw no
	t.TimerInterval = 10
	t.TimerEnabled  = True
End Sub

Sub MoveStandup(t, tp)
	Select Case t.TimerInterval
		Case 10 : tp.TransX = -5
		Case 11 : tp.TransX = 0
	End Select
	t.TimerInterval = t.TimerInterval + 1
End Sub


' ===============================================================================================
' switches
' ===============================================================================================

InitSwitches

Sub InitSwitches()
'	sw23a.IsDropped = True
'	sw19a.IsDropped = True
'	sw24a.IsDropped = True
'	sw15a.IsDropped = True
'	sw13a.IsDropped = True
'	sw12a.IsDropped = True
'	sw14a.IsDropped = True
End Sub

' top lanes
Sub sw23_Hit()
	SwitchHit 23, sw23a
End Sub
Sub sw23_Unhit()
	SwitchUnhit 23, sw23a
End Sub
Sub sw19_Hit()
	SwitchHit 19, sw19a
End Sub
Sub sw19_Unhit()
	SwitchUnhit 19, sw19a
End Sub
Sub sw24_Hit()
	SwitchHit 24, sw24a
End Sub
Sub sw24_Unhit()
	SwitchUnhit 24, sw24a
End Sub

' return lanes
Sub sw15_Hit()
	SwitchHit 15, sw15a
End Sub
Sub sw15_Unhit()
	SwitchUnhit 15, sw15a
End Sub
Sub sw13_Hit()
	SwitchHit 13, sw13a
End Sub
Sub sw13_Unhit()
	SwitchUnhit 13, sw13a
End Sub

' outlanes
Sub sw12_Hit()
	SwitchHit 12, sw12a
End Sub
Sub sw12_Unhit()
	SwitchUnhit 12, sw12a
End Sub
Sub sw14_Hit()
	SwitchHit 14, sw14a
End Sub
Sub sw14_Unhit()
	SwitchUnhit 14, sw14a
End Sub

' 10 point rubber switches
Sub sw1a_Hit()
	vpmTimer.PulseSw 1
End Sub
Sub sw1b_Hit()
	vpmTimer.PulseSw 1
End Sub
Sub sw1c_Hit()
	vpmTimer.PulseSw 1
End Sub
Sub sw1d_Hit()
	vpmTimer.PulseSw 1
End Sub
Sub sw1e_Hit()
	vpmTimer.PulseSw 1
End Sub
Sub sw1f_Hit()
	vpmTimer.PulseSw 1
End Sub

' star trigger
Sub sw2a_Hit()
	StarHit 2, sw2ap
End Sub
Sub sw2a_Unhit()
	StarUnhit 2, sw2ap
End Sub
Sub sw2b_Hit()
	StarHit 2, sw2bp
End Sub
Sub sw2b_Unhit()
	StarUnhit 2, sw2bp
End Sub
Sub sw2c_Hit()
	StarHit 2, sw2cp
End Sub
Sub sw2c_Unhit()
	StarUnhit 2, sw2cp
End Sub

Sub ShooterLane_Hit()
	If B2SOn Then Controller.B2SSetData 101,1
End Sub
Sub ShooterLane_Unhit()
	If B2SOn Then Controller.B2SSetData 101,0
End Sub


Sub SwitchHit(switchno, switch)
	Controller.Switch(switchno) = True
	If Not switch Is Nothing Then
		switch.TransZ = 15
	End If
	PlaySound "Sensor"
End Sub
Sub SwitchUnhit(switchno, switch)
	Controller.Switch(switchno) = False
	If Not switch Is Nothing Then
		switch.TransZ = 0
	End If
End Sub

Sub StarHit(switchno, switch)
	Controller.Switch(switchno) = True
	If Not switch Is Nothing Then
		switch.TransZ = -5
	End If
	PlaySound "Sensor"
End Sub
Sub StarUnhit(switchno, switch)
	Controller.Switch(switchno) = False
	If Not switch Is Nothing Then
		switch.TransZ = 0
	End If
End Sub
' ===============================================================================================
' GI and flashers
' ===============================================================================================

Sub SolGI(IsGIOff)
	SetLamp 150, IIf(IsGIOff,0,1)
End Sub


' ===============================================================================================
' some general methods
' ===============================================================================================

Function BallSpeed()
	On Error Resume Next
	BallSpeed = SQR(ActiveBall.VelX ^ 2 + ActiveBall.VelY ^ 2)
End Function

Function Sign(number)
	If number < 0 Then
		Sign = -1
	ElseIf number > 0 Then
		Sign = 1
	Else
		Sign = 0
	End If
End Function

Function PlaySoundEffect(minspeed, lowsound, lowspeed, sound, highspeed, highsound, velX, velY)
	Dim currentspeed
	currentspeed = BallSpeed
	If currentspeed >= minspeed Then
		If (velX = 0 Or Sign(ActiveBall.VelX) = Sign(velX)) And (velY = 0 Or Sign(ActiveBall.VelY) = Sign(velY)) Then
			If highspeed > 0 And currentspeed > highspeed Then
				PlaySound highsound
			ElseIf currentspeed < lowspeed Then
				PlaySound lowsound
			Else
				PlaySound sound
			End If
		End If
	End If
End Function

Function IIF(ifcond, iftrue, iffalse)
	If ifcond Then
		IIF = iftrue
	Else
		IIF = iffalse
	End If
End Function


Dim LampState(200), FadingLevel(200), FadingState(200)
Dim FlashState(200), FlashLevel(200)
Dim FlashSpeedUp, FlashSpeedDown
'Dim x

AllLampsOff()
LampTimer.Interval = 40 'lamp fading speed
LampTimer.Enabled = 1
'
FlashInit()
'FlasherTimer.Interval = 10 'flash fading speed
'FlasherTimer.Enabled = 1

Sub LampTimer_Timer()
    Dim chgLamp, num, chg, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp) Then
        For ii = 0 To UBound(chgLamp)
            LampState(chgLamp(ii, 0) ) = chgLamp(ii, 1)
            FadingLevel(chgLamp(ii, 0) ) = chgLamp(ii, 1) + 4
			FlashState(chgLamp(ii, 0) ) = chgLamp(ii, 1)
        Next
    End If

    UpdateLamps
	If Table1.ShowDT = True then UpdateLEDs:End If
End Sub

Sub FlashInit
    Dim i
    For i = 0 to 200
        FlashState(i) = 0
        FlashLevel(i) = 0
    Next

    FlashSpeedUp = 50   ' fast speed when turning on the flasher
    FlashSpeedDown = 10 ' slow speed when turning off the flasher, gives a smooth fading
    AllFlashOff()
End Sub

Sub AllFlashOff
    Dim i
    For i = 0 to 200
        FlashState(i) = 0
    Next
End Sub
 
Sub UpdateLamps
	' standard lamps
NFadeL 1, l1
NFadeL 2, l2
NFadeL 3, l3
NFadeL 4, l4
NFadeL 5, l5
NFadeL 6, l6
NFadeL 7, l7
NFadeL 8, l8
NFadeL 9, l9
NFadeLm 10, l10r
NFadeL 10, l10
NFadeL 17, l17
NFadeL 18, l18
NFadeL 19, l19
NFadeL 20, l20
NFadeL 21, l21
NFadeL 22, l22
NFadeL 23, l23
NFadeL 24, l24
NFadeL 25, l25
NFadeLm 26, l26r
NFadeL 26, l26
NFadeL 33, l33
NFadeL 34, l34
NFadeL 35, l35
NFadeL 36, l36
NFadeL 37, l37
NFadeL 38, l38
NFadeL 39, l39
NFadeL 40, l40
NFadeL 41, l41
NFadeLm 42, l42r
NFadeL 42, l42
NFadeL 43, l43
NFadeL 49, l49
NFadeL 50, l50
NFadeL 51, l51
NFadeL 53, l53
NFadeLm 54, l54r
NFadeL 54, l54
NFadeL 55, l55
NFadeL 56, l11
NFadeL 57, l57
NFadeLm 58, l58r
NFadeL 58, l58
NFadeL 65, l65
NFadeL 66, l66
NFadeL 67, l67
NFadeL 68, l68
NFadeL 69, l69
NFadeL 71, l71
NFadeL 81, l81
NFadeL 82, l82
NFadeL 83, l83
NFadeL 84, l84
NFadeL 85, l85
NFadeL 97, l97
NFadeL 98, l98
NFadeL 99, l99
NFadeL 100, l100
NFadeL 101, l101
NFadeL 102, l102
NFadeL 113, l113
NFadeL 114, l114
NFadeL 115, l115
NFadeL 116, l116
NFadeL 118, l118
NFadeLm 119, lxx
NFadeLm 119, l119

NFadeL 59, l59 'left apron light
NFadeL 103, l103 'right apron light
NFadeL 119, l119f
		
	' refresh flag
	Dim refresh
	refresh = False
	If Flashers_ON = 1 Then
		Light58.State = 1
		Light77.State = 1
		Light78.State = 1
	Else
		Light58.State = 0
		Light77.State = 0
		Light78.State = 0
	End If
	If GI_On = 1 Then
		NFadeL 12, l12 'bumper4
		NFadeL 28, l28 'bumper3
		NFadeL 44, l44 'bumper2
		NFadeL 60, l60 'bumper1

		NFadeLm 87, l87a 'alpha lane guides
		NFadeL 87, l87b 'alpha lane guides
	Else
		NFadeL 12, l12 'bumper4
		NFadeL 28, l28 'bumper3
		NFadeL 44, l44 'bumper2
		NFadeL 60, l60 'bumper1
	End If
	If refresh Then
		fRefresh.State = Abs(fRefresh.State - 1)
	End If

	' GI
	If GI_ON = 1 Then
'		SetBumperLights
		NFadeL 150, Light80
	End If
End Sub

'Sub SetBumperLights()
'	If LampState(150) = 4 Or LampState(150) = 5 Then
'		For Each obj In Array(Bumper4,Bumper5,Bumper6,Bumper7,Bumper8,Bumper9,Bumper10, _
'							  Bumper11,Bumper12,Bumper13,Bumper14,Bumper15,Bumper16,Bumper17,Bumper18,Bumper19,Bumper20, _
'							  Bumper21,Bumper22,Bumper23,Bumper24,Bumper25,Bumper26,Bumper27,Bumper28,Bumper29,Bumper30, _
'							  Bumper31,Bumper32,Bumper33,Bumper34,Bumper35,Bumper36,Bumper37,Bumper38,Bumper39,Bumper40, _
'							  Bumper41,Bumper42,Bumper43,Bumper44,Bumper45)
'			obj.State = IIf(LampState(150)=4, 0, 1)
'		Next
'	End If
'End Sub

Sub FadePrim(nr, pri, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:pri.image = d:FadingLevel(nr) = 0
        Case 3:pri.image = c:FadingLevel(nr) = 1
        Case 4:pri.image = b:FadingLevel(nr) = 2
        Case 5:pri.image = a:FadingLevel(nr) = 3
    End Select
End Sub

Sub NFadePrim(nr, pri, a, b)
    Select Case FadingLevel(nr)
        Case 4:pri.image = b
        Case 5:pri.image = a
    End Select
End Sub

Sub FadePri(nr, pri, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:pri.image = d:FadingLevel(nr) = 0 'Off
        Case 3:pri.image = c:FadingLevel(nr) = 2 'fading...
        Case 4:pri.image = b:FadingLevel(nr) = 3 'fading...
        Case 5:pri.image = a:FadingLevel(nr) = 1 'ON
    End Select
End Sub

''Lights

Sub NFadeL(nr, a)
    Select Case FadingLevel(nr)
        Case 4:a.state = 0:FadingLevel(nr) = 0
        Case 5:a.State = 1:FadingLevel(nr) = 1
    End Select
End Sub

Sub NFadeLm(nr, a)
    Select Case FadingLevel(nr)
        Case 4:a.state = 0
        Case 5:a.State = 1
    End Select
End Sub

' Flasher objects
' Uses own faster timer

Sub Flash(nr, object)
    Select Case FlashState(nr)
        Case 0 'off
            FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown
            If FlashLevel(nr) < 0 Then
                FlashLevel(nr) = 0
                FlashState(nr) = -1 'completely off
            End if
            Object.opacity = FlashLevel(nr)
        Case 1 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp
            If FlashLevel(nr) > 1000 Then
                FlashLevel(nr) = 1000
                FlashState(nr) = -2 'completely on
            End if
            Object.opacity = FlashLevel(nr)
    End Select
End Sub

 Sub AllLampsOff():For x = 1 to 200:LampState(x) = 4:FadingLevel(x) = 4:Next:UpdateLamps:UpdateLamps:Updatelamps:End Sub
 
Sub SetLamp(nr, value)
    If value = 0 AND LampState(nr) = 0 Then Exit Sub
    If value = 1 AND LampState(nr) = 1 Then Exit Sub
    LampState(nr) = abs(value) + 4
FadingLevel(nr ) = abs(value) + 4: FadingState(nr ) = abs(value) + 4
End Sub

Sub SetFlash(nr, stat)
    FlashState(nr) = ABS(stat)
End Sub

' Use CreateBallID(kickername) to manually create a ball with a BallID
' Can also be used on nonVPM tables (EM or Custom)

Sub CreateBallID(aKicker)
	For cnt = 1 to ubound(ballStatus)				' Loop through all possible ball IDs
		If ballStatus(cnt) = 0 Then					' If ball ID is available...
			Set CurrentBall(cnt) = aKicker.createball		' Set ball object with the first available ID
			CurrentBall(cnt).uservalue = cnt				' Assign the ball's uservalue to it's new ID
			ballStatus(cnt) = 1						' Mark this ball status active
			ballStatus(0) = ballStatus(0)+1 		' Increment ballStatus(0), the number of active balls
			If B2BOn > 0 Then						' If B2BOn is 0, it overrides auto-turn on collision detection
													' If more than one ball active, start collision detection process
				If ballStatus(0) > 1 and XYdata.enabled = False Then XYdata.enabled = True
			End If
			Exit For								' New ball ID assigned, exit loop
		End If
   	Next 
End Sub

'Call this sub from every kicker that destroys a ball, before the ball is destroyed.
	
Sub ClearBallID
  	On Error Resume Next							' Error handling for debugging purposes
   	iball = ActiveBall.uservalue					' Get the ball ID to be cleared
   	If Err Then Msgbox Err.description & vbCrLf & iball
    ballStatus(iBall) = 0 							' Clear the ball status
   	ballStatus(0) = ballStatus(0)-1 				' Subtract 1 ball from the # of balls in play
   	On Error Goto 0
End Sub

Sub ClearBallIDByNo(iBall)
  	On Error Resume Next							' Error handling for debugging purposes
   	ballStatus(iBall) = 0 							' Clear the ball status
   	ballStatus(0) = ballStatus(0)-1 				' Subtract 1 ball from the # of balls in play
   	On Error Goto 0
End Sub

Sub PrimT_Timer
	If t17.isdropped = true then t17f.rotatetoend:t17p.reflectionenabled = 0:
	if t17.isdropped = false then t17f.rotatetostart:t17p.reflectionenabled = 1
	t17p.transy = t17f.currentangle
	If t20.isdropped = true then t20f.rotatetoend:t20p.reflectionenabled = 0
	if t20.isdropped = false then t20f.rotatetostart:t20p.reflectionenabled = 1
	t20p.transy = t20f.currentangle
	If t21.isdropped = true then t21f.rotatetoend:t21p.reflectionenabled = 0
	if t21.isdropped = false then t21f.rotatetostart:t21p.reflectionenabled = 1
	t21p.transy = t21f.currentangle
	If t22.isdropped = true then t22f.rotatetoend:t22p.reflectionenabled = 0
	if t22.isdropped = false then t22f.rotatetostart:t22p.reflectionenabled = 1
	t22p.transy = t22f.currentangle
	If t25.isdropped = true then t25f.rotatetoend:t25p.reflectionenabled = 0
	if t25.isdropped = false then t25f.rotatetostart:t25p.reflectionenabled = 1
	t25p.transy = t25f.currentangle
	If t26.isdropped = true then t26f.rotatetoend:t26p.reflectionenabled = 0
	if t26.isdropped = false then t26f.rotatetostart:t26p.reflectionenabled = 1
	t26p.transy = t26f.currentangle
	If t27.isdropped = true then t27f.rotatetoend:t27p.reflectionenabled = 0
	if t27.isdropped = false then t27f.rotatetostart:t27p.reflectionenabled = 1
	t27p.transy = t27f.currentangle
End Sub

'************************************
' What you need to add to your table
'************************************

' a timer called CollisionTimer. With a fast interval, like from 1 to 10
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

' The Double For loop: This is a double cycle used to check the collision between a ball and the other ones.
' If you look at the parameters of both cycles, you’ll notice they are designed to avoid checking 
' collision twice. For example, I will never check collision between ball 2 and ball 1, 
' because I already checked collision between ball 1 and 2. So, if we have 4 balls, 
' the collision checks will be: ball 1 with 2, 1 with 3, 1 with 4, 2 with 3, 2 with 4 and 3 with 4.

' Sum first the radius of both balls, and if the height between them is higher then do not calculate anything more,
' since the balls are on different heights so they can't collide.

' The next 3 lines calculates distance between xth and yth balls with the Pytagorean theorem,

' The first "If": Checking if the distance between the two balls is less than the sum of the radius of both balls, 
' and both balls are not already colliding.

' Why are we checking if balls are already in collision? 
' Because we do not want the sound repeting when two balls are resting closed to each other.

' Set the collision property of both balls to True, and we assign the number of the ball colliding

' Play the collide sound of your choice using the VOL, PAN and PITCH functions

' Last lines: If the distance between 2 balls is more than the radius of a ball,
' then there is no collision and then set the collision property of the ball to False (-1).


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

' ===============================================================================================
' LED's display
' Based on JP's code and he based on Eala's routine
' ===============================================================================================

Dim Digits(27)
 
Digits(0)  = Array(a_00, a_01, a_02, a_03, a_04, a_05, a_06)
Digits(1)  = Array(a_10, a_11, a_12, a_13, a_14, a_15, a_16)
Digits(2)  = Array(a_20, a_21, a_22, a_23, a_24, a_25, a_26)
Digits(3)  = Array(a_30, a_31, a_32, a_33, a_34, a_35, a_36, a_37)
Digits(4)  = Array(a_40, a_41, a_42, a_43, a_44, a_45, a_46)
Digits(5)  = Array(a_50, a_51, a_52, a_53, a_54, a_55, a_56)

Digits(6)  = Array(b_00, b_01, b_02, b_03, b_04, b_05, b_06)
Digits(7)  = Array(b_10, b_11, b_12, b_13, b_14, b_15, b_16)
Digits(8)  = Array(b_20, b_21, b_22, b_23, b_24, b_25, b_26)
Digits(9)  = Array(b_30, b_31, b_32, b_33, b_34, b_35, b_36, b_37)
Digits(10) = Array(b_40, b_41, b_42, b_43, b_44, b_45, b_46)
Digits(11) = Array(b_50, b_51, b_52, b_53, b_54, b_55, b_56)

Digits(12) = Array(c_00, c_01, c_02, c_03, c_04, c_05, c_06, c_07)
Digits(13) = Array(c_10, c_11, c_12, c_13, c_14, c_15, c_16)
Digits(14) = Array(c_20, c_21, c_22, c_23, c_24, c_25, c_26)
Digits(15) = Array(c_30, c_31, c_32, c_33, c_34, c_35, c_36, c_37)
Digits(16) = Array(c_40, c_41, c_42, c_43, c_44, c_45, c_46)
Digits(17) = Array(c_50, c_51, c_52, c_53, c_54, c_55, c_56)

Digits(18) = Array(d_00, d_01, d_02, d_03, d_04, d_05, d_06, d_07)
Digits(19) = Array(d_10, d_11, d_12, d_13, d_14, d_15, d_16)
Digits(20) = Array(d_20, d_21, d_22, d_23, d_24, d_25, d_26)
Digits(21) = Array(d_30, d_31, d_32, d_33, d_34, d_35, d_36, d_37)
Digits(22) = Array(d_40, d_41, d_42, d_43, d_44, d_45, d_46)
Digits(23) = Array(d_50, d_51, d_52, d_53, d_54, d_55, d_56)

Digits(24) = Array(e_00, e_01, e_02, e_03, e_04, e_05, e_06)
Digits(25) = Array(e_10, e_11, e_12, e_13, e_14, e_15, e_16)

Digits(26) = Array(f_00, f_01, f_02, f_03, f_04, f_05, f_06)
Digits(27) = Array(f_10, f_11, f_12, f_13, f_14, f_15, f_16)

Sub UpdateLeds
	Dim chgLED, ii, num, chg, stat, digit
	chgLED = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
	If Not IsEmpty(chgLED) Then
		For ii = 0 To UBound(chgLED)
			num  = chgLED(ii, 0)
			chg  = chgLED(ii, 1)
			stat = chgLED(ii, 2)
			For Each digit In Digits(num)
				If (chg And 1) Then digit.State = (stat And 1)
				chg  = chg \ 2
				stat = stat \ 2
			Next
		Next
	End If
End Sub

'*************************************
'Bally Embryon
'*************************************
Sub editDips
    Dim vpmDips:Set vpmDips = New cvpmDips
    With vpmDips
        .AddForm 500, 600, "Embryon - DIP switches"
        .AddFrame 120, 0, 120, "Maximum credits", &H03000000, Array("10 credits", 0, "15 credits", &H01000000, "25 credits", &H02000000, "40 credits", &H03000000)                    																		'dip 25&26
		.AddFrame 120, 75, 120, "Balls per game", &H40000000, Array("2 balls", &H40000000+&H80000000, "3 balls", 0, "4 balls", &H80000000,  "5 balls", &H40000000)                                        		                                            'dip 31&32	                          
		.AddFrame 2, 0, 100,   "Play/Credits Chute 1", &H00000002+&H00000008, Array("1/1Coin", 0, "1/2Coin", &H00000003, "1/2Coin", &H3, "2/5Coin",&H00000009+&H0000006+&H00000010, "14/1Coin",&H00000002+&H00000008) 										'dip 1-5
		.AddFrame 2, 90, 100, "Play/Credits Chute 2", &H00080000+&H00060000+&H00100000, Array("1/1Coin", 0, "1/1Coin", &H00010000, "2/1Coin", &H00020000, "3/1Coin", &H00030000, "4/1Coin", &H00040000, "14/1Coin",&H00080000+&H00060000+&H00100000) 		'dip 17-20											
		.AddFrame 2, 195, 100, "Play/Credits Chute 3", &H00001000, Array("1/1Coin", 0, "1/2Coin", &H00000300, "1/2Coin", &H300, "2/5Coin", &H00000900+&H0000600+&H00000100, "14/1Coin", &H00000200+&H00000800)				'dip 9-13
        .AddChk 120, 155, 263,  Array("Attract Mode Voice (Activate Embryon)", &H20000000) 																																			    'dip 30
		.AddChk 120, 175, 263,  Array("Memory for Bonus and Specials", &H00000040) 																																		    'dip 7
		.AddChk 120, 195, 263,  Array("Memory for Multipliers", &H00000080) 																																				'dip 8
		.AddChk 120, 215, 263,  Array("Memory for lit Lanes 1&2", &H00400000) 																																				'dip 23
		.AddChk 120, 235, 263,  Array("Memory for Flipsave Feature", &H00800000) 																																			'dip 24
		.AddChk 120, 255, 263,  Array("Memory for Top Center Lane Feature", &H00002000) 																																	'dip 14
		.AddChk 120, 275, 263,  Array("Memory for Left Side Captive Ball Feature", 32768)																																	'dip 16
		.AddChk 120, 295, 263,  Array("Memory for Right Side Captive Ball Feature", &H00004000) 																															'dip 15
		.AddChk 120, 315, 263,  Array("Memory for lit Center Ouside Targets", &H00200000)																																	'dip 22
		.AddChk 120, 335, 263,  Array("Hitting 3 top drop targets once to Collect/ off Twice ", &H00000020) 																												'dip 6
		.AddChk 120, 355, 263,  Array("Hitting Center left/right target will light 2 lites", &H00100000)																													'dip 21
		.AddChk 120, 375, 263,  Array("Replays Earned will be Collected/ off only 1 per play", &H10000000)																												    'dip 29
		.AddChk 2, 290, 115,  Array("Match feature", &H08000000)                                                                                                                               							    'dip 28
		.AddChk 2, 310, 115,  Array("Credits displayed", &H04000000) 																																						'dip 27																																																	
        .AddLabel 50, 395, 350, 20, "After hitting OK, press F3 to reset game with new settings."
        .ViewDips
    End With
End Sub
Set vpmShowDips = GetRef("editDips")

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
'    JP's VP10 Collision & Rolling Sounds
'*****************************************

Const tnob = 8 ' total number of balls
ReDim rolling(tnob)
ReDim collision(tnob)
Initcollision

Sub Initcollision
    Dim i
    For i = 0 to tnob
        collision(i) = -1
        rolling(i) = False
    Next
End Sub

Sub CollisionTimer_Timer()
    Dim BOT, B, B1, B2, dx, dy, dz, distance, radii
    BOT = GetBalls

    ' rolling

	For B = UBound(BOT) +1 to tnob
        rolling(b) = False
        StopSound("fx_ballrolling" & b)
	Next

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

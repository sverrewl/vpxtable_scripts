'No Good Gofers by Bodydump
'Models by Dark
Option Explicit
Randomize

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' No special SSF tweaks yet.
' This is a JP table. He often uses walls as switches so I need to be careful of using PlaySoundAt
' , AudioFade(ActiveBall)


On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const BallSize = 50

'***Apron Style***
' 0 = Standard apron
' 1 = Custom apron
apronstyle = 0

'**************************
'Desktop/Fullscreen Changes
'**************************

If Table.ShowDT = true then
	Ramp15.visible = True
	Ramp16.visible = True
	Ramp17.visible = True
	Light47.Intensity = 0
else
	Ramp15.visible = False
	Ramp16.visible = False
	Ramp17.visible = False
end if


   LoadVPM"01520000","WPC.VBS",3.1



'********************
'Standard definitions
'********************

    Dim xx, apronstyle
    Dim Bump1, Bump2, Bump3, RGoferIsUp, RightRampUp, LGoferIsUp, LeftRampUp
	Dim bsTrough, bsLeftEject, bsPuttOutPopper, bsJetPopper, bsUpperRightEject, mTT, cbCaptive
	LGoferIsUp = 0
	LeftRampUp = 0
	RGoferIsUp = 0
	RightRampUp = 0
	Const cGameName = "ngg_13"
	Const UseSolenoids = 1
	Const UseLamps = 0
	Const UseGI = 0
	Const UseSync = 0
	Const HandleMechs = 1

	Set GICallback = GetRef("UpdateGIon")
	Set GICallback2 = GetRef("UpdateGI")

	Const SSolenoidOn="fx_Solenoid"
	Const SSolenoidOff="fx_solenoidoff"
	Const sCoin="fx_Coin"

'***********
' Solenoids
' the commented solenoids are not in used in this script
'***********

		SolCallback(1)="vpmSolAutoPlunger Plunger1,3,"					'Autofire
		SolCallback(2)="vpmSolAutoPlunger KickBack,3,"					'KickBack
		SolCallback(3)="SolPuttOut"										'Clubhouse Kicker
		SolCallback(4)="SolLeftGoferUP"									'Left Gofer Up
		SolCallback(5)="SolRightGoferUP"								'Right Gofer Up
		SolCallback(6)="bsJetPopper.SolOut"								'Jet Popper
		SolCallback(7)="bsLeftEject.SolOut"								'Left Eject (Sand Trap)
		SolCallback(8)="bsUpperRightEject.SolOut"						'Upper Right Eject
		SolCallback(9)="SolBallRelease"									'Trough Eject
'		'SolCallback(10)="vpmSolSound ""Sling"","						'Left Slingshot
'		'SolCallback(11)="vpmSolSound ""Sling"","						'Right Slingshot
'		'SolCallback(12)="vpmSolSound ""Jet3"","						'Top Jet Bumper
'		'SolCallback(13)="vpmSolSound ""Jet3"","						'Middle Jet Bumper
'		'SolCallback(14)="vpmSolSound ""Jet3"","						'Bottom Jet Bumper
		SolCallback(15)="SolLeftGoferDown"								'Left Gofer Down
		SolCallback(16)="SolRightGoferDown"								'Right Gofer Down
		SolCallback(17)="SetLamp 170,"									'Jet Flasher
		SolCallback(18)="SetLamp 178,"									'Lower Left Flasher
		SolCallback(19)="SetLamp 180,"									'Left Spinner Flasher
		SolCallback(20)="SetLamp 122,"									'Right Spinner Flasher
		SolCallback(21)="SetLamp 179,"									'Lower Right Flasher
		SolCallback(24)="SolSubway"										'Underground Pass
		SolCallback(25)="SetLamp 131,"									'Sand Trap Flasher
		SolCallback(26)="SetLamp 190,"									'Wheel Flasher
		SolCallback(27)="SolLeftRampDown"								'Left Ramp Down
		SolCallback(28)="SolRightRampDown"								'Right Ramp Down
		SolCallback(35)="SolSlamRamp"									'Ball Launch Ramp
		SolCallback(51)="SetLamp 177,"									'Upper Right Flasher 1
		SolCallback(52)="SetLamp 175,"									'Upper Right Flasher 2
		SolCallback(53)="SetLamp 173,"									'Upper Right Flasher 3
		SolCallback(56)="SetLamp 176,"									'Upper Left Flasher 1
		SolCallback(57)="SetLamp 174,"									'Upper Left Flasher 2
		SolCallback(58)="SetLamp 172,"									'Upper Left Flasher 3
'		SolCallback(sLRFlipper)="vpmSolFlipper RightFlipper,Nothing,"
'		SolCallback(sLLFlipper)="vpmSolFlipper LeftFlipper,Nothing,"
'		SolCallback(sURFlipper)="vpmSolFlipper UpperFlipper,Nothing,"
		SolCallback(54)="SetLamp 145,"
		SolCallback(55)="SetLamp 146,"
'
	Sub SolBallRelease(Enabled)
			If Enabled Then
				If bsTrough.Balls Then
					vpmTimer.PulseSw 31
				End If
			bsTrough.ExitSol_On
			End If
	End Sub

	Sub SolPuttOut(Enabled)
		If Enabled Then
			If bsPuttOutPopper.Balls Then
				bsPuttOutPopper.InitKick sw44a,170,12
				PlaySound"scoopexit"
			End If
			bsPuttOutPopper.ExitSol_On
		End If
	End Sub
	Sub SolSubway(Enabled)
		If Enabled Then
			If bsPuttOutPopper.Balls Then
				bsPuttOutPopper.InitKick subwaystart,180,3
				bsPuttOutPopper.ExitSol_On
			End If
		End If
	End Sub

	Sub SolLeftGoferUP(Enabled)
		If Enabled Then
			LrampTimer.Enabled = 0
			If LGoferIsUp = 0 Then
				Controller.Switch(41)=0
				Controller.Switch(47)=0
				LGofer.IsDropped=0
				LGoferUp.Enabled = 1
				LGoferDown.Enabled = 0
				Lramp.collidable=0
				Lramp1.collidable=1
				LeftRampUp = 1
			Else
				Controller.Switch(41)=0
				Controller.Switch(47)=0
				LGoferUp.Enabled = 1
				LeftGoferBlock.collidable = 1
				LGPos = 6
			End If
		End If
	End Sub

	Sub SolRightGoferUP(Enabled)
		If Enabled Then
			RrampTimer.Enabled = 0
			If RGoferIsUp = 0 Then
				Controller.Switch(42)=0
				Controller.Switch(48)=0
				RGofer.IsDropped=0
				RGoferUp.Enabled = 1
				RGoferDown.Enabled = 0
				Rramp.collidable=0
				Rramp1.collidable=1
				RightRampUp = 1
			Else
				Controller.Switch(42)=0
				Controller.Switch(48)=0
				RGoferUp.Enabled = 1
				RGPos = 6
			End If
		End If
	End Sub

	Sub SolLeftGoferDown(Enabled)
		If Enabled Then
			Controller.Switch(41)=1
			LGoferDown.Enabled = 1
			LGoferUp.Enabled = 0
			LGofer.IsDropped=1
		End If
	End Sub

	Sub SolRightGoferDown(Enabled)
		If Enabled Then
			Controller.Switch(42)=1
			RGoferDown.Enabled = 1
			RGoferUp.Enabled = 0
			RGofer.IsDropped = 1
		End If
	End Sub

	Sub SolLeftRampDown(Enabled)
		If Enabled Then
			RightRampUp = 0
			LrampDir = 1
			Lramp.collidable=1
			Lramp1.collidable=0
			PlaySound "gate"
			LrampTimer.Enabled = 1

		Else

		End If
	End Sub

	Sub SolRightRampDown(Enabled)
		If Enabled Then
			RightRampUp = 0
			RrampDir = 1
			Rramp.collidable=1
			Rramp1.collidable=0
			PlaySound "gate"
			RrampTimer.Enabled = 1

		Else

		End If
	End Sub

	Sub SolSlamRamp(Enabled)
		If Enabled Then
			For each xx in aslamramp:xx.IsDropped=0:Next
			For each xx in aslamrampup:xx.IsDropped=1:Next
			jumpramptrigger.enabled = 1
			jumpramp.Collidable=1
			slamrampup.Collidable = 0
			Flipper1.RotateToEnd
			rampshadow.visible = 0
			SlamRampDownTimer.Enabled = 1
		Else

		End If
	End Sub

Sub SlamRampDownTimer_Timer()
		SlamRampUpTimer.Enabled = 1
		SlamRampDownTimer.Enabled = 0
End Sub
Sub SlamRampUpTimer_Timer()
		For each xx in aslamramp:xx.IsDropped=1:Next
		For each xx in aslamrampup:xx.IsDropped=9:Next
		jumpramp.Collidable=0
		slamrampup.Collidable = 1
		jumpramptrigger.enabled = 0
		Flipper1.RotateToStart
		rampshadow.visible = 1
		SlamRampUpTimer.Enabled = 0
End Sub

'************
' Table init.
'************

  Sub Table_Init
	vpmInit me
	With Controller
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
		.SplashInfoLine="No Good Gofers" & vbNewLine & "VP Table by Bodydump"
		.HandleKeyboard=0
		.ShowTitle=0
		.ShowDMDOnly=1
		.ShowFrame=0
		.HandleMechanics=1
		.Hidden=1
		On Error Resume Next
		.Run GetPlayerHWnd
		If Err Then MsgBox Err.Description
		On Error Goto 0
	End With
	'Nudging
    	vpmNudge.TiltSwitch=14
    	vpmNudge.Sensitivity=1
    	vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingshot)

	'Trough
		Set bsTrough=New cvpmBallStack
		With bsTrough
			.InitSw 0,32,33,34,35,36,37,0
			.InitKick BallRelease,105,5
			.InitEntrySnd "fx_Solenoid", "fx_Solenoid"
			.InitExitSnd SoundFX("fx_ballrel",DOFContactors), SoundFX("fx_Solenoid",DOFContactors)
			.Balls=6
		End With

	'Sand Trap Eject
		Set bsLeftEject=New cvpmBallStack
		With bsLeftEject
			.InitSaucer LeftEject,78,78,20
			.InitExitSnd SoundFX("fx_Popper",DOFContactors), SoundFX("fx_Solenoid",DOFContactors)
		End With

	'Putt Out Popper
		Set bsPuttOutPopper=New cvpmBallStack
		With bsPuttOutPopper
			.InitSw 0,44,0,0,0,0,0,0
			.InitKick sw44a,170,12
		End With

	'Jet Popper
		Set bsJetPopper=New cvpmBallStack
		With bsJetPopper
			.InitSw 0,38,25,0,0,0,0,0
			.InitKick JetPopper,180,10
			.InitExitSnd SoundFX("fx_Popper",DOFContactors), SoundFX("fx_Solenoid",DOFContactors)
		End With

	'Upper Right Eject
		Set bsUpperRightEject=New cvpmBallStack
		With bsUpperRightEject
			.InitSw 0,46,0,0,0,0,0,0
			.InitKick upperrightkicker,120,5
			.InitExitSnd SoundFX("fx_Popper",DOFContactors), SoundFX("fx_Solenoid",DOFContactors)
		End With

	'Turntable
		Set mTT=New cvpmTurnTable
		With mTT
			.InitTurnTable TurnTable,40
			.SpinUp=40
			.SpinDown=40
			.CreateEvents "mTT"
		End With

	'Captive Ball
		Set cbCaptive=New cvpmCaptiveBall
		With cbCaptive
			.InitCaptive Captive,CWALL,Array(Captive1,Captive2),350
			.NailedBalls=1
			.Start
			.ForceTrans=0.9
			.MinForce=4
			.CreateEvents "cbCaptive"
		End With
		Captive1.CreateBall
	If apronstyle = 0 Then
		apron.IsDropped = 1
		apron1.IsDropped = 0
	Else
		apron.IsDropped = 0
		apron1.IsDropped = 1
	End If

	If nightday<5 then
	For each xx in extraambient:xx.intensityscale = 1:Next
	End If
	If nightday>5 then
	For each xx in extraambient:xx.intensity = xx.intensity*.9:Next
	End If
	If nightday>10 then
	For each xx in extraambient:xx.intensity = xx.intensity*.7:Next
	For each xx in GIleft:xx.intensity = xx.intensity*.9:Next
	For each xx in GIright:xx.intensity = xx.intensity*.9:Next
	End If
	If nightday>20 then
	For each xx in extraambient:xx.intensity = xx.intensity*.6:Next
	For each xx in GIleft:xx.intensity = xx.intensity*.8:Next
	For each xx in GIright:xx.intensity = xx.intensity*.8:Next
	End If
	If nightday>30 then
	For each xx in extraambient:xx.intensity = xx.intensity*.4:Next
	For each xx in GIleft:xx.intensity = xx.intensity*.7:Next
	For each xx in GIright:xx.intensity = xx.intensity*.7:Next
	End If



'      '**Main Timer init
           PinMAMETimer.Enabled = 1

	Plunger1.PullBack
	Kickback.PullBack
	Controller.Switch(22)=1 'close coin door
	Controller.Switch(24)=0 'always closed
	Controller.Switch(41)=1 'drop left gofer
	Controller.Switch(42)=1 'drop right gofer
	Controller.Switch(47)=1 'left ramp down
	Controller.Switch(48)=1 'right ramp down
	LGofer.IsDropped = 1
	RGofer.IsDropped = 1
	Rramp1.collidable=0
	Lramp1.collidable=0

End Sub

		Sub Table_Paused:Controller.Pause = 1:End Sub
		Sub Table_unPaused:Controller.Pause = 0:End Sub

'**********
' Keys
'**********

		Sub table_KeyDown(ByVal Keycode)
			If keycode = PlungerKey Then Plunger.Pullback:PlaySound SoundFX("fx_plungerpull",DOFContactors)
			If keycode = LeftTiltKey Then vpmNudge.DoNudge 90, 3:PlaySound SoundFX("fx_nudge",0), 0, 1, -0.1, 0.25
			If keycode = RightTiltKey Then vpmNudge.DoNudge 270, 3:PlaySound SoundFX("fx_nudge",0), 0, 1, 0.1, 0.25
			If keycode = CenterTiltKey Then vpmNudge.DoNudge 0, 4:PlaySound SoundFX("fx_nudge",0), 0, 1, 0, 0.25
			If vpmKeyDown(keycode) Then Exit Sub
		End Sub

		Sub table_KeyUp(ByVal Keycode)
			If keycode = PlungerKey Then Plunger.Fire:PlaySound SoundFX("fx_plunger",DOFContactors)
			If vpmKeyUp(keycode) Then Exit Sub
		End Sub


'********************
'    Flippers
'********************

		SolCallback(sLRFlipper) = "SolRFlipper"
		SolCallback(sLLFlipper) = "SolLFlipper"

		Sub SolLFlipper(Enabled)
			If Enabled Then
				PlaySound SoundFX("fx_flipperup",DOFContactors), 0, 1, -0.1, 0.05
				LeftFlipper.RotateToEnd
			Else
				PlaySound SoundFX("fx_flipperdown",DOFContactors), 0, 1, -0.1, 0.05
				LeftFlipper.RotateToStart
			End If
		End Sub

		Sub SolRFlipper(Enabled)
			If Enabled Then
				PlaySound SoundFX("fx_flipperup",DOFContactors), 0, 1, 0.1, 0.05
				RightFlipper.RotateToEnd
				UpperFlipper.RotateToEnd
			Else
				PlaySound SoundFX("fx_flipperdown",DOFContactors), 0, 1, 0.1, 0.05
				RightFlipper.RotateToStart
				UpperFlipper.RotateToStart
			End If
		End Sub

'****************************
' Drain holes, vuks & saucers
'****************************
		Sub Drain_Hit:Playsound "fx_drain":bsTrough.AddBall Me:End Sub
		Sub LeftEject_Hit:PlaySound "fx_kicker_enter", 0, 1, -0.05, 0.05:bsLeftEject.AddBall 0:End Sub  		'Sand Trap Eject
		Sub sw44a_Hit:PlaySound "ball_bounce", 0, 1, 0.05, 0.05:bsPuttOutPopper.AddBall Me:End Sub			'Putt Out Popper
		Sub sw44b_Hit:PlaySound "ball_bounce", 0, 1, 0.05, 0.05:bsPuttOutPopper.AddBall Me:End Sub			'Putt Out Popper
		Sub sw44c_Hit:PlaySound "ball_bounce", 0, 1, 0.05, 0.05:bsPuttOutPopper.AddBall Me:End Sub			'Putt Out Popper
		Sub sw44d_Hit:PlaySound "ball_bounce", 0, 1, 0.05, 0.05:bsPuttOutPopper.AddBall Me:End Sub			'Putt Out Popper
		Sub sw44e_Hit:PlaySound "ball_bounce", 0, 1, 0.05, 0.05:bsPuttOutPopper.AddBall Me:End Sub			'Putt Out Popper
		Sub sw44f_Hit:PlaySound "ball_bounce", 0, 1, 0.05, 0.05:bsPuttOutPopper.AddBall Me:End Sub			'Putt Out Popper
		Sub sw44g_Hit:PlaySound "ball_bounce", 0, 1, 0.05, 0.05:bsPuttOutPopper.AddBall Me:End Sub			'Putt Out Popper
		Sub sw44h_Hit:PlaySound "ball_bounce", 0, 1, 0.05, 0.05:bsPuttOutPopper.AddBall Me:End Sub			'Putt Out Popper
		Sub sw44i_Hit:PlaySound "ball_bounce", 0, 1, 0.05, 0.05:bsPuttOutPopper.AddBall Me:End Sub			'Putt Out Popper
		Sub subwayend_Hit:bsJetPopper.AddBall Me:End Sub								'Underground Pass
		Sub holeinone_Hit:PlaySound "ball_bounce", 0, 1, 0.05, 0.05:Me.DestroyBall:vpmTimer.PulseSwitch(68),160,"HandleTopHole":End Sub
		Sub behindleftgofer_Hit:PlaySound "fx_kicker_enter", 0, 1, -0.05, 0.05:Me.DestroyBall:vpmTimer.PulseSwitch(67),100,"HandlePuttOut":End Sub
		Sub rightpopperjam_Hit:PlaySound "fx_kicker_enter", 0, 1, 0.05, 0.05:Me.DestroyBall:vpmTimer.PulseSwitch(45),160,"HandleRightGofer":End Sub
		Sub HandleTopHole(swNo):vpmTimer.PulseSwitch(67),100,"HandlePuttOut":End Sub
		Sub HandlePuttOut(swNo):bsPuttOutPopper.AddBall 0:End Sub
		Sub HandleRightGofer(swNo):bsUpperRightEject.AddBall 0:End Sub

'***************
'  Slingshots
'***************
'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
		Dim RStep, Lstep

		Sub RightSlingShot_Slingshot
			PlaySound SoundFX("right_slingshot",DOFContactors), 0, 1, 0.05, 0.05
			vpmTimer.PulseSw 52
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
			PlaySound SoundFX("left_slingshot",DOFContactors), 0, 1, -0.05, 0.05
			vpmTimer.PulseSw 51
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

'***************
'   Bumpers
'***************
		Sub Bumper1_Hit:vpmTimer.PulseSw 54:PlaySound SoundFX("fx_bumper",DOFContactors), 0, 1, -0.03, 0.05:End Sub
		Sub Bumper2_Hit:vpmTimer.PulseSw 53:PlaySound SoundFX("fx_bumper",DOFContactors), 0, 1, -0.05, 0.05:End Sub
		Sub Bumper3_Hit:vpmTimer.PulseSw 55:PlaySound SoundFX("fx_bumper",DOFContactors), 0, 1, -0.05, 0.05:End Sub


'************
' Spinners
'************
		Sub Spinner1_Spin():vpmTimer.PulseSwitch(61),0,"":PlaySound "fx_spinner", 0, 1, -0.05, 0.05: End Sub
		Sub Spinner2_Spin():vpmTimer.PulseSwitch(62),0,"":PlaySound "fx_spinner", 0, 1, 0.05, 0.05: End Sub

'*********************
' Switches & Rollovers
'*********************
		Sub sw18_Hit:Controller.Switch(18)=1:PlaySound "fx_sensor":End Sub
		Sub sw18_UnHit:Controller.Switch(18)=0:End Sub
		Sub sw16_Hit:Controller.Switch(16)=1:PlaySound "fx_sensor":End Sub
		Sub sw16_UnHit:Controller.Switch(16)=0:End Sub
		Sub sw17_Hit:Controller.Switch(17)=1:PlaySound "fx_sensor":End Sub
		Sub sw17_UnHit:Controller.Switch(17)=0:End Sub
		Sub sw26_Hit:Controller.Switch(26)=1:PlaySound "fx_sensor":End Sub
		Sub sw26_UnHit:Controller.Switch(26)=0:End Sub
		Sub sw27_Hit:Controller.Switch(27)=1:PlaySound "fx_sensor":End Sub
		Sub sw27_UnHit:Controller.Switch(27)=0:End Sub
		Sub sw28_Hit:Controller.Switch(28)=1:PlaySound "fx_sensor":End Sub
		Sub sw28_UnHit:Controller.Switch(28)=0:End Sub
		Sub sw71_Hit:Controller.Switch(71)=1:PlaySound "fx_sensor":End Sub
		Sub sw71_UnHit:Controller.Switch(71)=0:End Sub
		Sub sw72_Hit:Controller.Switch(72)=1:PlaySound "fx_sensor":End Sub
		Sub sw72_UnHit:Controller.Switch(72)=0:End Sub
		Sub sw86_Hit:Controller.Switch(86)=1:PlaySound "fx_sensor":End Sub
		Sub sw86_UnHit:Controller.Switch(86)=0:End Sub

'***************
'  Targets
'***************

		Sub sw23_Hit
			vpmTimer.PulseSw 23
			PlaySound SoundFX("fx_target",DOFContactors), 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
		End Sub

		Sub sw56_Hit
			vpmTimer.PulseSw 56
			PlaySound SoundFX("fx_target",DOFContactors), 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
		End Sub

		Sub sw57_Hit
			vpmTimer.PulseSw 57
			PlaySound SoundFX("fx_target",DOFContactors), 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
		End Sub

		Sub sw58_Hit
			vpmTimer.PulseSw 58
			PlaySound SoundFX("fx_target",DOFContactors), 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
		End Sub

		Sub sw77_Hit
			vpmTimer.PulseSw 77
			PlaySound SoundFX("fx_target",DOFContactors), 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
		End Sub

		Sub sw81_Hit
			vpmTimer.PulseSw 81
			PlaySound SoundFX("fx_target",DOFContactors), 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
		End Sub

		Sub sw82_Hit
			vpmTimer.PulseSw 82
			PlaySound SoundFX("fx_target",DOFContactors), 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
		End Sub

		Sub sw83_Hit
			vpmTimer.PulseSw 83
			PlaySound SoundFX("fx_target",DOFContactors), 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
		End Sub

		Sub sw84_Hit
			vpmTimer.PulseSw 84
			PlaySound SoundFX("fx_target",DOFContactors), 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
		End Sub

		Sub sw85_Hit
			vpmTimer.PulseSw 85
			PlaySound SoundFX("fx_target",DOFContactors), 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
		End Sub

'*********************
' Random Switches
'*********************
		Sub Gate4_Hit():vpmTimer.PulseSw 12:End Sub
		Sub sw15_Hit():vpmTimer.PulseSw 15:End Sub
		Sub cart_Hit():PlaySound "plastichit":vpmTimer.PulseSw 74:GolfCart.Y = 218:WheelTargetPosition.Y = 217:Me.TimerEnabled = 1:End Sub			'Golf Cart
		Sub cart_Timer:GolfCart.Y = 215:WheelTargetPosition.Y = 211:Me.TimerEnabled = 0:End Sub
		Sub sw73_Hit():vpmTimer.PulseSw 73:End Sub
		Sub LGofer_Hit():PlaySound "plastichit":vpmTimer.PulseSw 65:LeftGoferBlock.Y = 441:Me.TimerEnabled = 1:End Sub
		Sub LGofer_Timer:LeftGoferBlock.Y = 449:Me.TimerEnabled = 0:End Sub
		Sub RGofer_Hit():PlaySound "plastichit":vpmTimer.PulseSw 75:RightGoferBlock.Y = 516.625:Me.TimerEnabled = 1:End Sub
		Sub RGofer_Timer:RightGoferBlock.Y = 524.625:Me.TimerEnabled = 0:End Sub

'****************
'Gofer Animations
'****************

Dim LGPos, LGDir, RGPos, RGPos2, RGDir, LrampPos, LrampDir, RrampPos, RrampDir
LGPos=22:RGPos=22:LrampPos=22:RrampPos=22
LGDir=0:RGDir=0:LrampDir=1:RrampDir=1
LGofer.TimerEnabled=1:RGofer.TimerEnabled=1:LrampTimer.Enabled=1

	  Sub LGoferUp_Timer()
			  Select Case LGPos
					Case 0: LeftGoferBlock.z=-20
								LGoferIsUp = 1
								LGoferUp.Enabled = 0
								Lramp.HeightBottom=110
					Case 1: LeftGoferBlock.z=-25:Lramp.HeightBottom=104
					Case 2: LeftGoferBlock.z=-31:Lramp.HeightBottom=102
					Case 3: LeftGoferBlock.z=-37:Lramp.HeightBottom=100
					Case 4: LeftGoferBlock.z=-31:Lramp.HeightBottom=102
					Case 5: LeftGoferBlock.z=-25:Lramp.HeightBottom=104
					Case 6: LeftGoferBlock.z=-20:Lramp.HeightBottom=110
					Case 7: LeftGoferBlock.z=-25:Lramp.HeightBottom=104
					Case 8: LeftGoferBlock.z=-31:Lramp.HeightBottom=98
					Case 9: LeftGoferBlock.z=-37:Lramp.HeightBottom=91
					Case 10: LeftGoferBlock.z=-43:Lramp.HeightBottom=84
					Case 11: LeftGoferBlock.z=-49:Lramp.HeightBottom=77
					Case 12: LeftGoferBlock.z=-55:Lramp.HeightBottom=70
					Case 13: LeftGoferBlock.z=-61:Lramp.HeightBottom=63
					Case 14: LeftGoferBlock.z=-67:Lramp.HeightBottom=56
					Case 15: LeftGoferBlock.z=-73:Lramp.HeightBottom=49
					Case 16: LeftGoferBlock.z=-79:Lramp.HeightBottom=42
					Case 17: LeftGoferBlock.z=-85:Lramp.HeightBottom=35
					Case 18: LeftGoferBlock.z=-91:Lramp.HeightBottom=28
					Case 19: LeftGoferBlock.z=-97:Lramp.HeightBottom=21
					Case 20: LeftGoferBlock.z=-103:Lramp.HeightBottom=14
					Case 21: LeftGoferBlock.z=-109:Lramp.HeightBottom=7
					Case 22: LeftGoferBlock.z=-115:Lramp.HeightBottom=0
			  End Select
					If LGpos>0 then LGpos=LGpos-1
	  End Sub

	  Sub LGoferDown_Timer()
	  Select Case LGPos
			Case 0: LeftGoferBlock.z=-19
			Case 1: LeftGoferBlock.z=-25
			Case 2: LeftGoferBlock.z=-31
			Case 3: LeftGoferBlock.z=-37
			Case 4: LeftGoferBlock.z=-43
			Case 5: LeftGoferBlock.z=-49
			Case 6: LeftGoferBlock.z=-55
			Case 7: LeftGoferBlock.z=-61
			Case 8: LeftGoferBlock.z=-67
			Case 9: LeftGoferBlock.z=-73
			Case 10: LeftGoferBlock.z=-79
			Case 11: LeftGoferBlock.z=-85
			Case 12: LeftGoferBlock.z=-91
			Case 13: LeftGoferBlock.z=-97
			Case 14: LeftGoferBlock.z=-103
			Case 15: LeftGoferBlock.z=-109
			Case 16: LeftGoferBlock.z=-115
						LGoferDown.Enabled = 0
						LGoferIsUp = 0
	  End Select
			If LGpos<16 then LGpos=LGpos+1
	  End Sub

	  Sub RGoferUp_Timer()
			  Select Case RGPos
					Case 0: RightGoferBlock.z=-20
								RGoferIsUp = 1
								RGoferUp.Enabled = 0
								Rramp.HeightBottom=110
					Case 1: RightGoferBlock.z=-25:Rramp.HeightBottom=104
					Case 2: RightGoferBlock.z=-31:Rramp.HeightBottom=102
					Case 3: RightGoferBlock.z=-37:Rramp.HeightBottom=100
					Case 4: RightGoferBlock.z=-31:Rramp.HeightBottom=102
					Case 5: RightGoferBlock.z=-25:Rramp.HeightBottom=104
					Case 6: RightGoferBlock.z=-20:Rramp.HeightBottom=110
					Case 7: RightGoferBlock.z=-25:Rramp.HeightBottom=104
					Case 8: RightGoferBlock.z=-31:Rramp.HeightBottom=98
					Case 9: RightGoferBlock.z=-37:Rramp.HeightBottom=91
					Case 10: RightGoferBlock.z=-43:Rramp.HeightBottom=84
					Case 11: RightGoferBlock.z=-49:Rramp.HeightBottom=77
					Case 12: RightGoferBlock.z=-55:Rramp.HeightBottom=70
					Case 13: RightGoferBlock.z=-61:Rramp.HeightBottom=63
					Case 14: RightGoferBlock.z=-67:Rramp.HeightBottom=56
					Case 15: RightGoferBlock.z=-73:Rramp.HeightBottom=49
					Case 16: RightGoferBlock.z=-79:Rramp.HeightBottom=42
					Case 17: RightGoferBlock.z=-85:Rramp.HeightBottom=35
					Case 18: RightGoferBlock.z=-91:Rramp.HeightBottom=28
					Case 19: RightGoferBlock.z=-97:Rramp.HeightBottom=21
					Case 20: RightGoferBlock.z=-103:Rramp.HeightBottom=14
					Case 21: RightGoferBlock.z=-109:Rramp.HeightBottom=7
					Case 22: RightGoferBlock.z=-115:Rramp.HeightBottom=0
			  End Select
					If RGpos>0 then RGpos=RGpos-1
	  End Sub

	  Sub RGoferDown_Timer()
	  Select Case RGPos
			Case 0: RightGoferBlock.z=-19
			Case 1: RightGoferBlock.z=-25
			Case 2: RightGoferBlock.z=-31
			Case 3: RightGoferBlock.z=-37
			Case 4: RightGoferBlock.z=-43
			Case 5: RightGoferBlock.z=-49
			Case 6: RightGoferBlock.z=-55
			Case 7: RightGoferBlock.z=-61
			Case 8: RightGoferBlock.z=-67
			Case 9: RightGoferBlock.z=-73
			Case 10: RightGoferBlock.z=-79
			Case 11: RightGoferBlock.z=-85
			Case 12: RightGoferBlock.z=-91
			Case 13: RightGoferBlock.z=-97
			Case 14: RightGoferBlock.z=-103
			Case 15: RightGoferBlock.z=-109
			Case 16: RightGoferBlock.z=-115
						RGoferDown.Enabled = 0
						RGoferIsUp = 0
	  End Select
			If RGpos<16 then RGpos=RGpos+1
	  End Sub

Sub LrampTimer_Timer()
	If LGoferIsUp = 0 Then
	  Select Case LrampPos
					Case 0: Lramp.HeightBottom=110
					Case 1: Lramp.HeightBottom=104
					Case 2: Lramp.HeightBottom=102
					Case 3: Lramp.HeightBottom=100
					Case 4: Lramp.HeightBottom=102
					Case 5: Lramp.HeightBottom=104
					Case 6: Lramp.HeightBottom=110
					Case 7: Lramp.HeightBottom=104
					Case 8: Lramp.HeightBottom=98
					Case 9: Lramp.HeightBottom=91
					Case 10: Lramp.HeightBottom=84
					Case 11: Lramp.HeightBottom=77
					Case 12: Lramp.HeightBottom=70
					Case 13: Lramp.HeightBottom=63
					Case 14: Lramp.HeightBottom=56
					Case 15: Lramp.HeightBottom=49
					Case 16: Lramp.HeightBottom=42
					Case 17: Lramp.HeightBottom=35
					Case 18: Lramp.HeightBottom=28
					Case 19: Lramp.HeightBottom=21
					Case 20: Lramp.HeightBottom=14
					Case 21: Lramp.HeightBottom=7
					Case 22: Lramp.HeightBottom=0:Controller.Switch(47)=1:LeftRampUp = 0:LrampTimer.Enabled = 0
			If Lramppos<22 then Lramppos=Lramppos+1
		End Select
		End If
	  End Sub


Sub RrampTimer_Timer()
	If RGoferIsUp = 0 Then
	  Select Case RrampPos
					Case 0: Rramp.HeightBottom=110
					Case 1: Rramp.HeightBottom=104
					Case 2: Rramp.HeightBottom=102
					Case 3: Rramp.HeightBottom=100
					Case 4: Rramp.HeightBottom=102
					Case 5: Rramp.HeightBottom=104
					Case 6: Rramp.HeightBottom=110
					Case 7: Rramp.HeightBottom=104
					Case 8: Rramp.HeightBottom=98
					Case 9: Rramp.HeightBottom=91
					Case 10: Rramp.HeightBottom=84
					Case 11: Rramp.HeightBottom=77
					Case 12: Rramp.HeightBottom=70
					Case 13: Rramp.HeightBottom=63
					Case 14: Rramp.HeightBottom=56
					Case 15: Rramp.HeightBottom=49
					Case 16: Rramp.HeightBottom=42
					Case 17: Rramp.HeightBottom=35
					Case 18: Rramp.HeightBottom=28
					Case 19: Rramp.HeightBottom=21
					Case 20: Rramp.HeightBottom=14
					Case 21: Rramp.HeightBottom=7
					Case 22: Rramp.HeightBottom=0:Controller.Switch(48)=1:RightRampUp = 0:RrampTimer.Enabled = 0
			If Rramppos<22 then Rramppos=Rramppos+1
		End Select
		End If
	  End Sub

'**************
'Turntable Subs
'**************

		Set MotorCallback=GetRef("SRPRoutine")
		Dim WheelNewPos,WheelOldPos
		WheelOldPos=0

		Sub SRPRoutine
			WheelNewPos=ABS(INT(Controller.GetMech(0)/4))
			If WheelNewPos<>WheelOldPos Then
				'disc.roty = DiskArray(WheelNewPos)
				If WheelNewPos>WheelOldPos And WheelNewPos<15 Then:mTT.SolMotorState False,True:TTtimer.Enabled=0:BlurDir = 0:Blur.Enabled = 1:End If
				If WheelNewPos<WheelOldPos And WheelNewPos>0 Then:mTT.SolMotorState True,True:TTtimer.Enabled=0:BlurDir = 0:Blur.Enabled = 1:End If
			End If

			disc.roty = 360 - (Controller.GetMech(0) /64 * 360) + 12

			WheelOldPos=WheelNewPos
		End Sub
'		Sub SRPRoutine
'			WheelNewPos=ABS(INT(Controller.GetMech(0)/4))
'			If WheelNewPos<>WheelOldPos Then
'				disc.roty = DiskArray(WheelNewPos)
'				If WheelNewPos>WheelOldPos And WheelNewPos<15 Then:mTT.SolMotorState False,True:TTtimer.Enabled=0:BlurDir = 0:Blur.Enabled = 1:End If
'				If WheelNewPos<WheelOldPos And WheelNewPos>0 Then:mTT.SolMotorState True,True:TTtimer.Enabled=0:BlurDir = 0:Blur.Enabled = 1:End If
'			End If
'			WheelOldPos=WheelNewPos
'		End Sub

		Sub TTtimer_Timer
			mTT.SolMotorState False,False
			mTT.SolMotorState True,False
			BlurDir = 1
			Blur.Enabled = 1
			TTtimer.Enabled=0
		End Sub
	Dim DiskArray, BlurDir, Blurlevel
		DiskArray=Array("12","350","335","305","282","258","235","215","190","168","146","125","100","80","55","32")
		Blurlevel = 0
	Sub Blur_Timer()
		Select Case Blurlevel
			Case 0: disc.image = "NGG_Map"
				If BlurDir=1 Then
					Blur.Enabled = 0
					TTtimer.Enabled=1
				End If
			Case 1: disc.image = "NGG_Map_Mblur1"
			Case 2: disc.image = "NGG_Map_Mblur2"
			Case 3: disc.image = "NGG_Map_Mblur3"
			Case 4: disc.image = "NGG_Map_Mblur4"
			Case 5: disc.image = "NGG_Map_Mblur5"
			Case 6: disc.image = "NGG_Map_Mblur6"
				If BlurDir = 0 Then
					Blur.Enabled = 0
					TTtimer.Enabled=1
				End If
		End Select
		If BlurDir = 1 Then
			If Blurlevel>0 then Blurlevel = Blurlevel - 1
		Else
			If Blurlevel<6 Then Blurlevel = Blurlevel + 1
		End If
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
LampTimer.Interval = 10 'lamp fading speed
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
	NFadeL 11, L11
	NFadeL 12, L12
	NFadeL 13, L13
	NFadeL 14, L14
	NFadeL 15, L15
	NFadeL 16, L16
	NFadeL 17, L17
	NFadeL 18, L18
	NFadeL 21, L21
	NFadeL 22, L22
	NFadeL 23, L23
	NFadeL 24, L24
	NFadeL 25, L25
	NFadeL 26, L26
	NFadeL 27, L27
	NFadeL 28, L28
	NFadeL 31, L31
	NFadeL 32, L32
	NFadeL 33, L33
	NFadeL 34, L34
	NFadeL 35, L35
	NFadeL 36, L36
	NFadeL 37, L37
	NFadeL 38, L38
	NFadeL 41, L41
	NFadeL 42, L42
	NFadeL 43, L43
	NFadeL 44, L44
	NFadeL 45, L45
	NFadeL 46, L46
	NFadeL 47, L47
	NFadeL 48, L48
	NFadeL 51, L51
	NFadeL 52, L52
	NFadeL 53, L53
	NFadeL 54, L54
	NFadeLm 55, L55b
	NFadeL 55, L55
	NFadeL 56, L56
	NFadeL 57, L57
	NFadeL 58, L58
	NFadeL 61, L61
	NFadeL 62, L62
	NFadeL 63, L63
	NFadeL 64, L64
	NFadeL 65, L65
	NFadeL 66, L66
	NFadeL 67, L67
	NFadeL 68, L68
	NFadeLm 71, L71b
	NFadeL 71, L71
	NFadeL 72, L72
	NFadeL 73, L73
	NFadeL 74, L74
	NFadeL 75, L75
	NFadeL 76, L76
	NFadeL 77, L77
	NFadeLm 78, bump3light1
	NFadeL 78, bump3light
	NFadeL 81, L81
	NFadeL 82, L82
	NFadeL 83, L83
	NFadeL 84, L84
	NFadeL 85, L85
	NFadeLm 86, bump2light1
	NFadeL 86, bump2light
	NFadeLm 87, bump1light1
	NFadeL 87, bump1light
'	FadeL 88, L88, L88a	'start button
	NFadeL 122, L122
	NFadeLm 131, L131b
	NFadeLm 131, L131a
	NFadeL 131, L131
	'NFadeL 145, L145
	'NFadeL 146, L146
	NFadeARmCombi 145, 146, upperplayfield, "upperplayfield", "upperplayfieldright", "upperplayfieldleft", "upperplayfieldboth", Combifor145and146
	NFadeLm 170, L170b
	NFadeLm 170, L170a
	NFadeL 170, L170
	NFadeLm 172, L172a
	NFadeLm 172, L172b
	NFadeL 172, L172
	NFadeLm 173, L173a
	NFadeLm 173, L173b
	NFadeL 173, L173
	NFadeLm 174, L174a
	NFadeLm 174, L174b
	NFadeL 174, L174
	NFadeLm 175, L175a
	NFadeLm 175, L175b
	NFadeL 175, L175
	NFadeLm 176, L176a
	NFadeLm 176, L176b
	NFadeL 176, L176
	NFadeLm 177, L177a
	NFadeLm 177, L177b
	NFadeL 177, L177
	NFadeLm 178, L178a
	NFadeLm 178, L178b
	NFadeL 178, L178
	NFadeLm 179, L179a
	NFadeLm 179, L179b
	NFadeL 179, L179
	NFadeLm 180, L180a
	NFadeLm 180, L180b
	NFadeLm 180, L180c
	NFadeL 180, L180
	NFadeLm 190, L190b
	NFadeLm 190, L190a
	NFadeL 190, L190
End Sub
'
'Sub FlasherTimer_Timer()
'		Flash 11, h11
'		Flash 12, h12
'		Flash 13, h13
'		Flash 14, h14
'		Flash 15, h15
'		Flash 16, h16
'		Flash 17, h17
'		Flash 18, h18
'		Flash 21, h21
'		Flash 22, h22
'		Flash 23, h23
'		Flash 24, h24
'		Flash 25, h25
'		Flash 26, h26
'		Flash 27, h27
'		Flash 28, h28
'		Flash 31, h31
'		Flash 32, h32
'		Flash 33, h33
'		Flash 34, h34
'		Flash 35, h35
'		Flash 36, h36
'		Flash 37, h37
'		Flash 38, h38
'		Flash 41, h41
'		Flash 42, h42
'		Flash 43, h43
'		Flash 44, h44
'		Flash 45, h45
'		Flash 46, h46
'		Flash 47, h47
'		Flash 48, h48
'		Flash 51, h51
'		Flash 52, h52
'		Flash 53, h53
'		Flash 54, h54
'		Flashm 55, h55a
'		Flash 55, h55
'		Flash 56, h56
'		Flash 57, h57
'		Flash 58, h58
'		Flash 61, h61
'		Flash 62, h62
'		Flash 63, h63
'		Flash 64, h64
'		Flash 65, h65
'		Flash 66, h66
'		Flash 68, h68
'		Flashm 71, h71a
'		Flash 71, h71
'		Flash 72, h72
'		Flash 73, h73
'		Flash 74, h74
'		Flash 75, h75
'		Flash 77, h77
'		Flash 78, h78
'		Flash 81, h81
'		Flash 82, h82
'		Flash 83, h83
'		Flash 84, h84
'		Flash 85, h85
'		Flash 86, h86
'		Flash 87, h87
'		Flashm 120, Flasher19
'		Flash 120, Flasher1
'		Flashm 121, Flasher20
'		Flash 121, Flasher2
'		Flashm 122, Flasher21
'		Flashm 122, Flasher4
'		Flash 122, Flasher3
'		Flash 123, Flasher5
'		Flashm 124, Flasher15
'		Flash 124, Flasher6
'		Flashm 125, Flasher18
'		Flash 125, Flasher7
'		Flash 126, Flasher8
'		Flash 127, Flasher9
'		Flashm 128, Flasher17
'		Flash 128, Flasher10
'		Flashm 129, Flasher16
'		Flash 129, Flasher11
'		Flash 130, Flasher12
'		Flashm 131, Flasher14
'		Flash 131, Flasher13
' End Sub
'

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

'Lights, Ramps & Primitives used as 4 step fading lights
'a,b,c,d are the images used from on to off

Sub FadeObj(nr, object, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 4:object.image = b:FadingLevel(nr) = 6                   'fading to off...
        Case 5:object.image = a:FadingLevel(nr) = 1                   'ON
        Case 6, 7.8:FadingLevel(nr) = FadingLevel(nr) + 1             'wait
        Case 9:object.image = c:FadingLevel(nr) = FadingLevel(nr) + 1 'fading...
        Case 10, 11, 12:FadingLevel(nr) = FadingLevel(nr) + 1         'wait
        Case 13:object.image = d:FadingLevel(nr) = 0                  'Off
    End Select
End Sub

Sub FadeObjm(nr, object, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 4:object.image = b
        Case 5:object.image = a
        Case 9:object.image = c
        Case 13:object.image = d
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

' Desktop Objects: Reels & texts (you may also use lights on the desktop)

' Reels
Dim NewCombiResult, Combifor145and146
Sub NFadeARmCombi(nr1, nr2, ramp, aOff, a1, a2, aBoth, OldCombiResult)
	NewCombiResult=LampState(nr1)+(2*LampState(nr2))
	If OldCombiResult=NewCombiResult Then Exit Sub
	Select Case NewCombiResult
		Case 0:ramp.image=aOff
		Case 1:ramp.image=a1
		Case 2:ramp.image=a2
		Case 3:ramp.image=aBoth
	End Select
	OldCombiResult=NewCombiResult
End Sub

Sub FadeR(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.SetValue 1:FadingLevel(nr) = 6                   'fading to off...
        Case 5:object.SetValue 0:FadingLevel(nr) = 1                   'ON
        Case 6, 7.8:FadingLevel(nr) = FadingLevel(nr) + 1              'wait
        Case 9:object.SetValue 2:FadingLevel(nr) = FadingLevel(nr) + 1 'fading...
        Case 10, 11, 12:FadingLevel(nr) = FadingLevel(nr) + 1          'wait
        Case 13:object.SetValue 3:FadingLevel(nr) = 0                  'Off
    End Select
End Sub

Sub FadeRm(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.SetValue 1
        Case 5:object.SetValue 0
        Case 9:object.SetValue 2
        Case 3:object.SetValue 3
    End Select
End Sub

'Texts

Sub NFadeT(nr, object, message)
    Select Case FadingLevel(nr)
        Case 4:object.Text = "":FadingLevel(nr) = 0
        Case 5:object.Text = message:FadingLevel(nr) = 1
    End Select
End Sub

Sub NFadeTm(nr, object, b)
    Select Case FadingLevel(nr)
        Case 4:object.Text = ""
        Case 5:object.Text = message
    End Select
End Sub

'***********
' Update GI
'***********
Dim gistep, Giswitch
Giswitch = 0

Sub UpdateGIOn(no, Enabled)
	Select Case no
		Case 0 'GIright
			If Enabled Then
				For each xx in GIright:xx.State = 1: Next
			Else
				For each xx in GIright:xx.state = 0: Next
			End If
		Case 1
			If Enabled Then
				For Each xx in GIleft:xx.state = 1: Next
			else
				For Each xx in GIleft:xx.state = 0: Next
			End If
		Case 2
			If Enabled Then
				For Each xx in GIother:xx.state = 1: Next
			Else
				For Each xx in GIother:xx.state = 0: Next
			End If
	End Select
End Sub


Sub UpdateGI(no, step)
    Dim ii
    If step = 0 then exit sub 'only values from 1 to 8 are visible and reliable. 0 is not reliable and 7 & 8 are the same so...
    gistep = (step-1) / 7
    Select Case no
        Case 0
          For each ii in GIleft
				ii.IntensityScale = gistep
			Next
        Case 1
            For each ii in GIright
                ii.IntensityScale = gistep
            Next
        Case 2 ' also the bumpers er GI
            For each ii in GIother
                ii.IntensityScale = gistep
            Next
			Select case gistep
				Case 0
					GolfCart.image = "GolfCartMap"
				Case 1
					GolfCart.image = "GolfCartMap_ON1"
				Case 2
					GolfCart.image = "GolfCartMap_ON1"
				Case 3
					GolfCart.image = "GolfCartMap_ON2"
				Case 4
					GolfCart.image = "GolfCartMap_ON3"
				Case 5
					GolfCart.image = "GolfCartMap_ON4"
				Case 6
					GolfCart.image = "GolfCartMap_ON4"
				Case 7
					GolfCart.image = "GolfCartMap_ON5"
			End Select
    End Select
'    ' change the intensity of the flasher depending on the gi to compensate for the gi lights being off
'    For ii = 0 to 200
'        FlashMax(ii) = 6 - gistep * 3 ' the maximum value of the flashers
'    Next
End Sub


'''Ramp Helpers

'Sub centerramphelper_Hit():ActiveBall.Vely=ActiveBall.Vely*1.3:End Sub
'Sub centerramphelper2_Hit():Activeball.Velx=ActiveBall.Velx-20:End Sub
'Sub upperhelper_Hit():ActiveBall.Velx=ActiveBall.Velx*1.2:End Sub
'Sub upperhelper2_Hit():ActiveBall.Vely=ActiveBall.Vely*1.3:End Sub
'Sub jumpramptrigger_Hit():ActiveBall.Vely=ActiveBall.Vely*1.3:End Sub
'Sub upperjump_Hit():ActiveBall.Vely=ActiveBall.Vely*1.5:End Sub
'Sub upperjump_UnHit():ActiveBall.Velz=ActiveBall.Velz=0:End Sub


Sub Game_timer()
	slamrampP.z = Flipper1.CurrentAngle
	If RightRampUp = 0 Then RrampTimer.Enabled = 1
	If LeftRampUp = 0 Then LRampTimer.Enabled = 1
End Sub

Sub platformdrop_Unhit()
	If ActiveBall.Vely>0 then PlaySound "ball_bounce"
End Sub

Sub rightrampdrop_UnHit():PlaySound "ball_bounce":End Sub
Sub leftrampdrop_UnHit():PlaySound "ball_bounce":End Sub

Sub jumpramptrigger_Hit():PlaySound "metalhit_medium":End Sub

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

Sub UpperFlipper_Collide(parm)
 	RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub

Sub upperjump_UnHit():ActiveBall.Velz=ActiveBall.Velz=0:End Sub

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

Const tnob = 8 ' total number of balls in this table is 4, but always use a higher number here bacuse of the timing
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


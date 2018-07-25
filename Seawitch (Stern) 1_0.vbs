' Seawitch / IPD No. 43 / Stern May, 1980 / 4 Players
' http://www.ipdb.org/machine.cgi?id=2089
' VPX table by Flash62
' Big thanks to JPSalas as I used his Stern tables as a template.Oh, and thanks for all the great tables.
' Thanks to Loafer for the custom apron from his VP9 version of Seawitch.
' I was unable to make a decent one, so it worked.

' DOF commands by arngrim

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' No special SSF tweaks yet.


Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

	Const BallSize = 50

	LoadVPM "01560000", "stern.VBS", 3.26


	Dim VarHidden
	If Table1.ShowDT = true then
		VarHidden = 1
	else
		VarHidden = 0
		For each x in aReels
			x.Visible = 0
		Next
		lrail.Visible = 0
		rrail.Visible = 0
	end if


	Dim bsTrough, dtTBank, dtMBank, dtLBank
	Dim x, i, j, k 'used in loops
	Dim SpringMotion 'spring animation counter
	Dim LockSpring 'if sw25 (Right side rolloever) is hit, prevent ball from going up through spring

	Const cGameName = "seawitch"

	Const UseSolenoids = 2
	Const UseLamps = 0
	Const UseGI = 0
	Const UseSync = 0
	Const HandleMech = 0

	' Standard Sounds
	Const SSolenoidOn = "fx_Solenoid"
	Const SSolenoidOff = ""
	Const SCoin = "fx_coin"

	'************
	' Table init.
	'************

	Sub Table1_Init
		With Controller
			.GameName = cGameName
			.SplashInfoLine = "Seawitch, Stern 1980" & vbNewLine & "VPX table by Flash62"
			.HandleMechanics = 0
			.HandleKeyboard = 0
			.ShowDMDOnly = 1
			.ShowFrame = 0
			.ShowTitle = 0
			.Hidden = VarHidden
			If Err Then MsgBox Err.Description
		End With

		On Error Goto 0
		Controller.SolMask(0) = 0
		vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the timer to renable all the solenoids after 2 seconds
		Controller.Run


		' Nudging
		vpmNudge.TiltSwitch = 7
		vpmNudge.Sensitivity = 0.1
		vpmNudge.TiltObj = Array(LBumper, TBumper, RBumper, LeftSlingshot, RightSlingshot)


		' Lower Drop targets
		set dtLBank = new cvpmdroptarget
		With dtLBank
			.InitDrop Array(sw38, sw39, sw40), Array(38, 39, 40)
			'.Initsnd SoundFX("fx_droptarget",DOFDropTargets), SoundFX("fx_resetdrop",DOFContactors)
			.Initsnd "fx_droptarget", "fx_resetdrop"
			'.CreateEvents "dtLBank"
		End With


		' Mid Drop targets
		set dtMBank = new cvpmdroptarget
		With dtMBank
			.InitDrop Array(sw21, sw22, sw23, sw24), Array(21, 22, 23, 24)
			'.Initsnd SoundFX("fx_droptarget",DOFDropTargets), SoundFX("fx_resetdrop",DOFContactors)
			.Initsnd "fx_droptarget", "fx_resetdrop"
			'.CreateEvents "dtMBank"
		End With

		' Top Drop targets
		set dtTBank = new cvpmdroptarget
		With dtTBank
			.InitDrop Array(sw29, sw30, sw31, sw32), Array(29, 30, 31, 32)
			'.Initsnd SoundFX("fx_droptarget",DOFDropTargets), SoundFX("fx_resetdrop",DOFContactors)
			.Initsnd "fx_droptarget", "fx_resetdrop"
			'.CreateEvents "dtTBank"
		End With


		' Trough
		Set bsTrough = New cvpmBallStack
			bsTrough.InitNoTrough BallRelease,33,115,3
			bsTrough.InitExitSnd "fx_ballrel", "fx_Solenoid"
			bsTrough.Balls = 1


		' Main Timer init
		PinMAMETimer.Interval = PinMAMEInterval
		PinMAMETimer.Enabled = 1

		l18.State = 1
		lrro.state = 1
		WallSpring.IsDropped = 0 'wall for curved spring above plunger. Set to up to start. Only will drop when trigger below it is hit (I.E. ball is launched)
'		RampSpr1.Visible = 1'animation for spring. Only show the first one, closed position
'		RampSpr2.Visible = 0
'		RampSpr3.Visible = 0
'		RampSpr4.Visible = 0
'		RampSpr5.Visible = 0

		WallSpring1.Visible = 1'animation for spring. Only show the first one, closed position
		WallSpring2.Visible = 0
		WallSpring3.Visible = 0
		WallSpring4.Visible = 0
		WallSpring5.Visible = 0

		LockSpring = 0 'not set, regular play, can launch ball through the spring

'		RampSpr1.HasWallImage = True'animation for spring. Only show the first one, closed position
'		RampSpr2.HasWallImage = False
'		RampSpr3.HasWallImage = False
'		RampSpr4.HasWallImage = False
'		RampSpr5.HasWallImage = False
	End Sub

	Sub Table1_Paused:Controller.Pause = 1:End Sub
	Sub Table1_unPaused:Controller.Pause = 0:End Sub

	'**********
	' Keys
	'**********

	Sub table1_KeyDown(ByVal Keycode)
		If keycode = LeftTiltKey Then Nudge 270, 5:PlaySound SoundFX("fx_nudge",0), 0, 1, -0.1, 0.25
		If keycode = RightTiltKey Then Nudge 90, 5:PlaySound SoundFX("fx_nudge",0), 0, 1, 0.1, 0.25
		If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound SoundFX("fx_nudge",0), 0, 1, 0, 0.25
		If keycode = PlungerKey Then PlaySound "fx_PlungerPull", 0, 1, 0.1, 0.25:Plunger.Pullback
		If keycode = 64 then editDips ' F6
		If vpmKeyDown(keycode) Then Exit Sub
	End Sub

	Sub table1_KeyUp(ByVal Keycode)
		If vpmKeyUp(keycode) Then Exit Sub
		If keycode = PlungerKey Then PlaySound "fx_plunger", 0, 1, 0.1, 0.25:Plunger.Fire
	End Sub

	'*********
	' Spring wall
	'*********


	Sub TriggerSpring_Hit

		If LockSpring = 0 Then 'only drop it if the lock is off (triggered by right side rollover)
			WallSpring.IsDropped = 1 'drop the wall to let the ball through
		End If
	End Sub


	Sub TriggerSpring_UnHit
		If LockSpring = 0 Then 'only drop it if the lock is off (triggered by right side rollover)
			SpringMotion = 0 'reset the process
			TriggerSpring.TimerEnabled = 1 'fire the timer off
		End If
	End Sub


	Sub TriggerSpring_Timer
		SpringMotion = SpringMotion + 1 'advance the counter

		Select Case SpringMotion
			Case 1
				WallSpring1.Visible = 0
				WallSpring2.Visible = 1
			Case 2
				WallSpring2.Visible = 0
				WallSpring3.Visible = 1
			Case 3
				WallSpring3.Visible = 0
				WallSpring4.Visible = 1
				WallSpring.IsDropped = 0 'put the Wall back up
			Case 4
				WallSpring4.Visible = 0
				WallSpring5.Visible = 1
			Case 5
				WallSpring5.Visible = 0
				WallSpring4.Visible = 1
			Case 6
				WallSpring4.Visible = 0
				WallSpring3.Visible = 1
			Case 7
				WallSpring3.Visible = 0
				WallSpring2.Visible = 1
			Case 8
				WallSpring2.Visible = 0
				WallSpring1.Visible = 1
				TriggerSpring.TimerEnabled = 0
		End Select
	End Sub



	'*********
	' Switches
	'*********

	Dim LStep, RStep

	Sub LeftSlingShot_Slingshot
		PlaySound SoundFX("fx_slingshot",DOFContactors), 0, 1, -0.05, 0.05
		LeftSling4.Visible = 1
		Lemk.RotX = 26
		LStep = 0
		'vpmTimer.PulseSw 16
		vpmTimer.PulseSw 12
		LeftSlingShot.TimerEnabled = 1
	End Sub

	Sub LeftSlingShot_Timer
		Select Case LStep
			Case 1:LeftSLing4.Visible = 0:LeftSLing3.Visible = 1:Lemk.RotX = 14
			Case 2:LeftSLing3.Visible = 0:LeftSLing2.Visible = 1:Lemk.RotX = 2
			Case 3:LeftSLing2.Visible = 0:Lemk.RotX = -10:LeftSlingShot.TimerEnabled = 0
		End Select

		LStep = LStep + 1
	End Sub


	Sub RightSlingShot_Slingshot
		PlaySound SoundFX("fx_slingshot",DOFContactors), 0, 1, 0.05, 0.05
		RightSling4.Visible = 1
		Remk.RotX = 26
		RStep = 0
		'vpmTimer.PulseSw 15
		vpmTimer.PulseSw 13
		RightSlingShot.TimerEnabled = 1
	End Sub

	Sub RightSlingShot_Timer
		Select Case RStep
			Case 1:RightSLing4.Visible = 0:RightSLing3.Visible = 1:Remk.RotX = 14
			Case 2:RightSLing3.Visible = 0:RightSLing2.Visible = 1:Remk.RotX = 2
			Case 3:RightSLing2.Visible = 0:Remk.RotX = -10:RightSlingShot.TimerEnabled = 0
		End Select

		RStep = RStep + 1
	End Sub

	' Bumpers
	'******
	Sub LBumper_Hit:vpmTimer.PulseSw 10:PlaySound SoundFX("fx_bumper",DOFContactors), 0, 1, 0.15, 0.15:End Sub
	Sub RBumper_Hit:vpmTimer.PulseSw 14:PlaySound SoundFX("fx_bumper",DOFContactors), 0, 1, 0.15, 0.15:End Sub
	Sub TBumper_Hit:vpmTimer.PulseSw 9:PlaySound SoundFX("fx_bumper",DOFContactors), 0, 1, 0.15, 0.15:End Sub


	' Drain & holes
	Sub Drain_Hit:Playsound "fx_drain":bsTrough.AddBall Me:End Sub



	' Rollovers
	'******
	'Lanes
	'Left Out
	Sub sw34_Hit:Controller.Switch(34) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
	Sub sw34_UnHit:Controller.Switch(34) = 0:End Sub

	'Left In
	Sub sw36_Hit:Controller.Switch(36) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
	Sub sw36_UnHit:Controller.Switch(36) = 0:End Sub

	'Right In
	Sub sw37_Hit:Controller.Switch(37) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
	Sub sw37_UnHit:Controller.Switch(37) = 0:End Sub

	'Right Out
	Sub sw35_Hit:Controller.Switch(35) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
	Sub sw35_UnHit:Controller.Switch(35) = 0:End Sub



	'Upper Star Rollover
	Sub sw4_Hit
		Controller.Switch(4) = 1
		l18.State = 0
		PlaySound "fx_sensor", 0, 1, pan(ActiveBall)
	End Sub

	Sub sw4_UnHit
		Controller.Switch(4) = 0
		l18.State = 1
		'vpmTimer.AddTimer 250, "l18.State=1"
	End Sub


	'Mid Right Star Rollover
 	Sub sw25_Hit() 'sw25
 		Controller.Switch(25) = 1
		lrro.State=0
		PlaySound "fx_sensor", 0, 1, pan(ActiveBall)
		LockSpring = 1 'lock the spring
		sw25.TimerEnabled = 1
	End Sub

	Sub sw25_UnHit
		Controller.Switch(25) = 0
		lrro.State=1
	End Sub

	Sub sw25_Timer 'trigger's timer set to 1000 - 1 second
		LockSpring = 0 'unlock the spring wall
		sw25.TimerEnabled = 0
	End Sub




	' Droptargets
	'******
	Sub sw38_Dropped:dtLBank.hit 1:PlaySound SoundFX("fx_droptarget",DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
	Sub sw39_Dropped:dtLBank.hit 2:PlaySound SoundFX("fx_droptarget",DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
	Sub sw40_Dropped:dtLBank.hit 3:PlaySound SoundFX("fx_droptarget",DOFDropTargets), 0, 1, pan(ActiveBall):End Sub

	Sub sw21_Dropped:dtMBank.hit 1:PlaySound SoundFX("fx_droptarget",DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
	Sub sw22_Dropped:dtMBank.hit 2:PlaySound SoundFX("fx_droptarget",DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
	Sub sw23_Dropped:dtMBank.hit 3:PlaySound SoundFX("fx_droptarget",DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
	Sub sw24_Dropped:dtMBank.hit 4:PlaySound SoundFX("fx_droptarget",DOFDropTargets), 0, 1, pan(ActiveBall):End Sub

	Sub sw29_Dropped:dtTBank.hit 1:PlaySound SoundFX("fx_droptarget",DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
	Sub sw30_Dropped:dtTBank.hit 2:PlaySound SoundFX("fx_droptarget",DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
	Sub sw31_Dropped:dtTBank.hit 3:PlaySound SoundFX("fx_droptarget",DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
	Sub sw32_Dropped:dtTBank.hit 4:PlaySound SoundFX("fx_droptarget",DOFDropTargets), 0, 1, pan(ActiveBall):End Sub

	'Spinner
	'******
	Sub sw5_Spin
		vpmTimer.PulseSw 5
		Playsound "fx_spinner", 0, 1, 0, -0.1
	End Sub


	' Targets
	'******
	Sub sw17_Hit:vpmTimer.PulseSw 17:PlaySound SoundFX("fx_target",DOFTargets), 0, 1, pan(ActiveBall):End Sub

	Sub sw18_Hit:vpmTimer.PulseSw 18:PlaySound SoundFX("fx_target",DOFTargets), 0, 1, pan(ActiveBall):End Sub

	Sub sw19_Hit:vpmTimer.PulseSw 19:PlaySound SoundFX("fx_target",DOFTargets), 0, 1, pan(ActiveBall):End Sub

	Sub sw20_Hit:vpmTimer.PulseSw 20:PlaySound SoundFX("fx_target",DOFTargets), 0, 1, pan(ActiveBall):End Sub


	' Gates
	Sub Gate1_Hit():PlaySound "gate":End Sub


	'*********
	'Solenoids
	'*********

	' Solenoids from the manual, HEY, THEY DON'T WORK.
	'Solcallback(3)  = "vpmsolsound SoundFX(""fx_knocker"",DOFKnocker),"
	'SolCallback(4)  = "dtMbank.SolDropUp"
	'SolCallback(8)  = "dtLBank.SolDropUp"
	'SolCallback(13) = "dtTbank.SolDropUp"
	'Solcallback(14) = "bsTrough.SolOut"
	'SolCallback(19) = "vpmNudge.SolGameOn"

'Ok, gonna grab them from the VP9 version
	SolCallback(8)  = "dtLBank.SolDropUp"
	SolCallback(7)  = "dtMBank.SolDropUp"
	SolCallback(9)  = "dtTbank.SolDropUp"
	SolCallback(10) = "bsTrough.SolOut"
	SolCallback(6)  = "vpmsolsound SoundFX(""fx_knocker"",DOFKnocker),"


	'
	'********************
	'     Flippers
	'********************

	'******
	SolCallback(sLRFlipper) = "SolRFlipper"
	SolCallback(sLLFlipper) = "SolLFlipper"

	Sub SolLFlipper(Enabled)
		If Enabled Then
			PlaySound SoundFX("fx_flipperup",DOFFlippers), 0, 1, -0.1, 0.25
			LeftFlipper.RotateToEnd
			LeftFlipper2.RotateToEnd
		Else
			PlaySound SoundFX("fx_flipperdown",DOFFlippers), 0, 1, -0.1, 0.25
			LeftFlipper.RotateToStart
			LeftFlipper2.RotateToStart
		End If
	End Sub

	Sub SolRFlipper(Enabled)
		If Enabled Then
			PlaySound SoundFX("fx_flipperup",DOFFlippers), 0, 1, 0.1, 0.25
			RightFlipper.RotateToEnd
			RightFlipper2.RotateToEnd
		Else
			PlaySound SoundFX("fx_flipperdown",DOFFlippers), 0, 1, 0.1, 0.25
			RightFlipper.RotateToStart
			RightFlipper2.RotateToStart
		End If
	End Sub

	Sub LeftFlipper_Collide(parm)
		PlaySound "fx_rubber_flipper", 0, parm / 10, -0.1, 0.25
	End Sub

	Sub LeftFlipper2_Collide(parm)
		PlaySound "fx_rubber_flipper", 0, parm / 10, -0.1, 0.25
	End Sub

	Sub RightFlipper_Collide(parm)
		PlaySound "fx_rubber_flipper", 0, parm / 10, 0.1, 0.25
	End Sub

	Sub RightFlipper2_Collide(parm)
		PlaySound "fx_rubber_flipper", 0, parm / 10, 0.1, 0.25
	End Sub

	'**************
	' Extra sounds
	'**************


'************************************
'          LEDs Display
'     Based on Scapino's LEDs
'************************************

Dim Digits(32)
Dim Patterns(11)
Dim Patterns2(11)

Patterns(0) = 0     'empty
Patterns(1) = 63    '0
Patterns(2) = 6     '1
Patterns(3) = 91    '2
Patterns(4) = 79    '3
Patterns(5) = 102   '4
Patterns(6) = 109   '5
Patterns(7) = 125   '6
Patterns(8) = 7     '7
Patterns(9) = 127   '8
Patterns(10) = 111  '9

Patterns2(0) = 128  'empty
Patterns2(1) = 191  '0
Patterns2(2) = 134  '1
Patterns2(3) = 219  '2
Patterns2(4) = 207  '3
Patterns2(5) = 230  '4
Patterns2(6) = 237  '5
Patterns2(7) = 253  '6
Patterns2(8) = 135  '7
Patterns2(9) = 255  '8
Patterns2(10) = 239 '9

'Assign 7-digit output to reels
Set Digits(0) = a0
Set Digits(1) = a1
Set Digits(2) = a2
Set Digits(3) = a3
Set Digits(4) = a4
Set Digits(5) = a5
Set Digits(6) = a6

Set Digits(7) = b0
Set Digits(8) = b1
Set Digits(9) = b2
Set Digits(10) = b3
Set Digits(11) = b4
Set Digits(12) = b5
Set Digits(13) = b6

Set Digits(14) = c0
Set Digits(15) = c1
Set Digits(16) = c2
Set Digits(17) = c3
Set Digits(18) = c4
Set Digits(19) = c5
Set Digits(20) = c6

Set Digits(21) = d0
Set Digits(22) = d1
Set Digits(23) = d2
Set Digits(24) = d3
Set Digits(25) = d4
Set Digits(26) = d5
Set Digits(27) = d6

Set Digits(28) = e0
Set Digits(29) = e1
Set Digits(30) = e2
Set Digits(31) = e3

Sub UpdateLeds
    On Error Resume Next
    Dim ChgLED, ii, jj, chg, stat
    ChgLED = Controller.ChangedLEDs(&HFF, &HFFFF)
    If Not IsEmpty(ChgLED)Then
        For ii = 0 To UBound(ChgLED)
            chg = chgLED(ii, 1):stat = chgLED(ii, 2)
            For jj = 0 to 10
                If stat = Patterns(jj)OR stat = Patterns2(jj)then Digits(chgLED(ii, 0)).SetValue jj
            Next
        Next
    End IF
End Sub


	Dim OldGiState
	OldGiState = -1 'start witht he Gi off

	Sub GiON
		For each x in aGiLights
			x.State = 1
		Next
	End Sub

	Sub GiOFF
		For each x in aGiLights
			x.State = 0
		Next
	End Sub

	Sub GiEffect
		For each x in aGiLights
			x.Duration 2, 3000, 1
		Next
	End Sub

	Sub GIUpdate
		Dim tmp, obj
		tmp = Getballs
		If UBound(tmp) <> OldGiState Then
			OldGiState = Ubound(tmp)
			If UBound(tmp) = -1 Then
				GiOff
			Else
				GiOn
			End If
		End If
	End Sub
	'*********

	'***************************************************
	'       JP's VP10 Fading Lamps & Flashers
	'       Based on PD's Fading Light System
	' SetLamp 0 is Off
	' SetLamp 1 is On
	' fading for non opacity objects is 4 steps
	'***************************************************

	Dim LampState(200), FadingLevel(200)
	Dim FlashSpeedUp(200), FlashSpeedDown(200), FlashMin(200), FlashMax(200), FlashLevel(200), FlashRepeat(200)

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

		UpdateLeds
		UpdateLamps
		GIUpdate
		RollingUpdate
	End Sub



	Sub UpdateLamps
		NFadeL 1, l1
		NFadeL 2, l2
		NFadeL 3, l3
		NFadeL2 4, l4, l4a
		NFadeL 5, l5
		NFadeL 6, l6
		NFadeL 7, l7
		NFadeL 8, l8
		NFadeL 9, l9

		NFadeL 10, l10
		NFadeL 11, l11
		NFadeL 12, l12
		'13
		NFadeL 14, l14
		NFadeL 15, l15
		'16
		NFadeL 17, l17
		NFadeL 18, l18
		NFadeL 19, l19

		NFadeL 20, l20
		NFadeL 21, l21
		NFadeL 22, l22
		NFadeL 23, l23
		NFadeL 24, l24
		NFadeL 25, l25
		NFadeL 26, l26
		NFadeL 27, l27
		NFadeL 28, l28
		'29

		NFadeL 30, l30
		NFadeL 31, l31
		'32
		NFadeL 33, l33
		NFadeL 34, l34
		NFadeL 35, l35
		NFadeL 36, l36
		NFadeL 37, l37
		NFadeL 38, l38
		NFadeL 39, l39

		NFadeL 40, l40
	'    NFadeL 41, l41
		NFadeL 42, l42
		NFadeL 43, l43
		NFadeL 44, l44
		'45
		NFadeL2 46, l46, lrro '???
		NFadeL 47, l47
		'48
		NFadeL 49, l49

		NFadeL 50, l50
	'    NFadeL 51, l51
		NFadeL 52, l52
		NFadeL 53, l53
		NFadeL 54, l54
		NFadeL 55, l55
		NFadeL 56, l56
	'    NFadeL 57, l57
		NFadeL 58, l58
		NFadeL 59, l59

		NFadeL 60, l60
		'61
		NFadeL 62, l62

	' backdrop lights
		If VarHidden Then
			NFadeT 13, l13, "Highscore"
			NFadeT 29, l29, "Ball in Play"
			NFadeT 45, l45, "Game Over"
			NFadeT 61, l61, "Tilt"
			NFadeT 63, l63, "Match"
		End If
	End Sub

	' div lamp subs

	Sub InitLamps()
		Dim x
		For x = 0 to 200
			LampState(x) = 0        ' current light state, independent of the fading level. 0 is off and 1 is on
			FadingLevel(x) = 4      ' used to track the fading state
			FlashSpeedUp(x) = 0.2   ' faster speed when turning on the flasher
			FlashSpeedDown(x) = 0.1 ' slower speed when turning off the flasher
			FlashMax(x) = 1         ' the maximum value when on, usually 1
			FlashMin(x) = 0         ' the minimum value when off, usually 0
			FlashLevel(x) = 0       ' the intensity of the flashers, usually from 0 to 1
			FlashRepeat(x) = 20     ' how many times the flash repeats
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

	Sub NFadeL2(nr, object1, object2)
		Select Case FadingLevel(nr)
			Case 4:object1.state = 0:object2.state = 0:FadingLevel(nr) = 0
			Case 5:object1.state = 1:object2.state = 1:FadingLevel(nr) = 1
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
			Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1            'wait
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
				If FlashLevel(nr) <FlashMin(nr) Then
					FlashLevel(nr) = FlashMin(nr)
					FadingLevel(nr) = 0 'completely off
				End if
				Object.IntensityScale = FlashLevel(nr)
			Case 5 ' on
				FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
				If FlashLevel(nr)> FlashMax(nr) Then
					FlashLevel(nr) = FlashMax(nr)
					FadingLevel(nr) = 1 'completely on
				End if
				Object.IntensityScale = FlashLevel(nr)
		End Select
	End Sub

	Sub Flashm(nr, object) 'multiple flashers, it doesn't change anything, it just follows the main flasher
		Select Case FadingLevel(nr)
			Case 4, 5
				Object.IntensityScale = FlashLevel(nr)
		End Select
	End Sub

	Sub FlashBlink(nr, object)
		Select Case FadingLevel(nr)
			Case 4 'off
				FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown(nr)
				If FlashLevel(nr) <FlashMin(nr) Then
					FlashLevel(nr) = FlashMin(nr)
					FadingLevel(nr) = 0 'completely off
				End if
				Object.IntensityScale = FlashLevel(nr)
				If FadingLevel(nr) = 0 AND FlashRepeat(nr) Then 'repeat the flash
					FlashRepeat(nr) = FlashRepeat(nr) -1
					If FlashRepeat(nr) Then FadingLevel(nr) = 5
				End If
			Case 5 ' on
				FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
				If FlashLevel(nr)> FlashMax(nr) Then
					FlashLevel(nr) = FlashMax(nr)
					FadingLevel(nr) = 1 'completely on
				End if
				Object.IntensityScale = FlashLevel(nr)
				If FadingLevel(nr) = 1 AND FlashRepeat(nr) Then FadingLevel(nr) = 4
		End Select
	End Sub

	' Desktop Objects: Reels & texts (you may also use lights on the desktop)

	' Reels

	Sub FadeR(nr, object)
		Select Case FadingLevel(nr)
			Case 4:object.SetValue 1:FadingLevel(nr) = 6                   'fading to off...
			Case 5:object.SetValue 0:FadingLevel(nr) = 1                   'ON
			Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1             'wait
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

	'******************************
	' Diverse Collection Hit Sounds
	'******************************

	Sub aMetals_Hit(idx):PlaySound "fx_MetalHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
	Sub aRubber_Bands_Hit(idx):PlaySound "fx_rubber_band", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
	Sub aRubber_Posts_Hit(idx):PlaySound "fx_rubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
	Sub aRubber_Pins_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
	Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
	Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
	Sub aWoods_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub

	'*******************************
	' Dipswitches from the old table
	'*******************************


Sub editDips
	Dim vpmDips : Set vpmDips = New cvpmDips
	With vpmDips
		.AddForm 900,400,"Seawitch - DIP switches"


		.AddChk 7,10,120,Array("Match feature",&H00100000)'dip 21

		.AddChk 130,10,120,Array("Credits display",&H00080000)'dip 20

		.AddChk 260,10,120,Array("Background Sounds",&H00002000)'dip 14

		.AddFrame 2,30,190,"Maximum credits",&H00060000,Array("10 credits",0,"15 credits",&H00020000,"25 credits",&H00040000,"40 credits",&H00060000)'dip 18 & 19
		.AddFrame 2,106,190,"# of Specials",&H00200000,Array("One per Ball",0,"One per Game",&H00200000)'dip 22
		.AddFrame 2,152,190,"Alternate Special Light",&H00010000,Array("Alternate",0,"Always On",&H00010000)'dip 17
		.AddFrame 2,198,190,"Special On & Bonus Reset to 3X",&H00000080,Array("Bonus @ 7X",0,"Bonus @ 6X or 7X",&H00000080)'dip 8
		.AddFrame 2,244,190,"Special Award",&HC0000000,Array("No Award",0,"Extra Ball",&H80000000, "100,000 Points",&H40000000,"Replay",&H80000000)'dip 31 & 32


		.AddFrame 205,30,190,"High Game To Date Award",&H0000C000,Array("Novelty",0,"1 Free Game",&H00004000,"2 Free Games",&H00008000,"3 Free Games",&H0000C000)'dip 15 & 16
		.AddFrame 205,106,190,"Balls per game",&H00000040,Array("3",0,"5",&H00000040)'dip 7
		.AddFrame 205,152,190,"High Score Award",&H00000020,Array("Extra Ball",0,"Replay",&H00000020)'dip 6
		.AddFrame 205,198,190,"If Highscore set to Extra Ball",&H30000000,Array("Only One",0,"3",&10000000,"5",&H30000000)'dip 29 & 30
		.AddFrame 205,264,190,"Extra Ball Lights",&H00C00000,Array("Completely Off",0,"Alternate",&H00800000,"Both On/Off",&H00C00000)'dip 23 & 24


'		.AddLabel 50,300,350,20,"Set selftest position 16,17,18 and 19 to 03 for the best gameplay."
		.AddLabel 50,420,300,20,"After hitting OK, press F3 to reset game with new settings."
		.ViewDips
	End With
End Sub
Set vpmShowDips = GetRef("editDips")

'	Dim saveDips
'	saveDips = Array(191, 255, 7, 63)
'
'	Sub DipSwitchEditor()
'		dim vpmDips, i, settings(3)
'
'		'Save the settings I don't have code for
'		for i = 0 to 3
'			settings(i) = Controller.dip(i) and saveDips(i)
'		next
'
'		on error resume next
'		set vpmDips = new cvpmDips
'		with vpmDips
'			'Title
'			'.AddForm 315, 250, "Seawitch DIP Switch Settings"
'			.AddForm 400, 600, "Seawitch DIP Switch Settings"
'
''sw6
'
'			'#1 - Balls per game - sw7
'			.AddFrame 2, 10, 140, "Balls per game", &H00000040, Array("3", 0, "5", &H00000040)
'
'			'#2 Special = 31 & 32 ???
'			.AddFrame 160, 10, 140, "Special scores", &HC0000000,  _
'				Array("No Award", 0, "100,000", &H40000000, _
'				"Extra Ball", &H80000000, "Replay", &HC0000000)
'
'			'#3 Speciials per ball - sw22 (says specials per game)
'			.AddFrame 2, 90, 140, "# of Specials", &H00200000, _
'				Array("One per Ball", 0, "One per Game", &H00200000)
'
'			'sw17
'			.AddFrame 160, 90, 140, "Alternate Special Lite", &H00010000,               _
'				Array("Alternate", 0, "Always On", &H00010000)
''
''			'#5Greatest Lights special sw24
''			.AddFrame 2, 170, 140, "GREATEST lights Special", &H00800000, _
''				Array("2nd time", 0, "1st time", &H00800000)
'
'			'sw6
'			.AddFrame 2, 200, 140, "High Score", &H00000020, _
'				Array("Extra Ball", 0, "Replay", &H00000020)
'
'
'			'sw 29
'			.AddFrame 2, 280, 140, "Highscore # of Extra Balls", &H10000000, _
'				Array("Only One", 0, "3 or 5", &H10000000)
'
'			'sw30
'			.AddFrame 160, 280, 140, "Highscore Extra Balls 3 or 5", &H20000000,               _
'				Array("3", 0, "5", &H20000000)
'
'			'#6 Display credit - sw20 (Hex = 14)
'			.AddChk 7, 500, 148, Array("Credit display", &H00080000)
'
'			'#7 Display match - sw21
'			.AddChk 160, 500, 148, Array("Match", &H00100000)
'
'			'buttons
'			.AddLabel 7, 530, 300, 20, "After hitting OK, press F3 to reset game with new settings."
'			.ViewDips
'		end with
'		if Err then DipSwitchDisplayError
'		on error goto 0
'
'		'Restore non-coded settings
'		for i = 0 to 3
'			Controller.dip(i) = settings(i) or((255-saveDips(i) ) and Controller.dip(i) )
'		next
'	End Sub
'
'	Sub DipSwitchDisplayError()
'		MsgBox "Can't display dip switch editor." & vbCRLF & vbCRLF &           _
'			"Be sure you have wshLtWtForm.ocx loaded and registered" & vbCRLF & _
'			"and vbs scripts 3.02 or higher." & vbCRLF & vbCRLF &               _
'			"You may want to hit F6 to edit the switches manually."
'	End Sub

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

	Const tnob = 5 ' total number of balls in this table is 4, but always use a higher number here because of the timing
	ReDim rolling(tnob)
	InitRolling

	Sub InitRolling
		Dim i
		For i = 0 to tnob
			rolling(i) = False
		Next
	End Sub

	Sub RollingUpdate()
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


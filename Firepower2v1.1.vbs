Option Explicit
Randomize

' Thalamus 2018-07-23
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' No special SSF tweaks yet.

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const BallsSize = 50

Const cGameName="fpwr2_l2",UseSolenoids=1,UseLamps=1,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff", SCoin="coin"

LoadVPM "01560000", "S7.VBS", 3.26
Dim DeskTopEnabled: DeskTopEnabled = FirePower2.showDT


Dim trTrough


If DeskTopEnabled = True Then 'Show Desktop components
Ramp16.visible=1
Ramp15.visible=1
Primitive13.visible=1
Else
Ramp16.visible=0
Ramp15.visible=0
Primitive13.visible=0
End if


'================================================================
' Define Easy Names for Solenoid Numbers
'================================================================
Const cOutHole = 1
Const cBallRampThrower = 2
Const cEjectHole = 3
Const cOrbitFlasher = 4
Const cGIRelay = 11
Const cBell = 15
Const cCoinLockout = 16
Const cLeftSlingShot = 17
Const cRightSlingShot = 18
Const cLowerLeftBumper = 19
Const cUpperRightBumper = 20
Const cUpperLeftBumper = 21
Const cLowerRightBumper = 22
'================================================================
' Define Easy Names for Switch Numbers
'================================================================
Const sPlumbTilt = 1
Const sBallRollTilt = 2
Const sCreditButton = 3
Const sRightCoin = 4
Const sCenterCoin = 5
Const sLeftCoin = 6
Const sSlamTilt = 7
Const sHighScoreReset = 8
Const sDrain = 9
Const sRightBallRamp = 10
Const sLeftBallRamp = 11
Const sBallShooter = 12
Const sLaneChange = 13
Const sRampInRollUnder = 14
Const sRampOutRollUnder = 15
Const sCenterLeftStandup = 16
Const sStandup = 17
Const sSpinner = 18
Const sOrbitInRollUnder = 19
Const sUpperLeftStandup = 20
Const sUpperRightStandup = 22
Const sEjectHole = 23
Const sReleaseTarget = 24
Const sFtgt = 25
Const sItgt = 26
Const sRtgt = 27
Const sEtgt = 28
Const sPtgt = 29
Const sOtgt = 30
Const sWtgt = 31
Const sEtgt2 = 32
Const sRtgt2 = 33
Const sA = 34
Const sB = 35
Const sC = 36
Const sD = 37
Const sLeftOutlane = 38
Const sRightOutlane = 39
Const sOrbitOutGate = 40
Const sLowerLeftBumper = 41
Const sUpperRightBumper = 42
Const sUpperLeftBumper = 43
Const sLowerRightBumper = 44
Const sLeftReturnLane = 45
Const sRightReturnLane = 46
Const sLeftSlingshot = 47
Const sRightSlingshot = 48
Const sPlayfieldTilt = 49


'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************

Sub FirePower2_Init
	vpmInit Me
	On Error Resume Next
		With Controller
		.GameName = cGameName
         If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Firepower II (Williams 1983)" & vbNewLine & "Created for VPX by Walamab"
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
	PinMAMETimer.Interval	= PinMAMEInterval
	PinMAMETimer.Enabled	= True
	vpmMapLights AllLights

' Setup Trough
Set trTrough = new cvpmTrough
trTrough.CreateEvents "trTrough", Array(Drain, BallRelease)
trTrough.balls = 2
trTrough.size = 2
'trTrough.EntrySw = swDrain
trTrough.InitEntrySounds "drain", "drain", "drain"
trTrough.InitExitSounds "ballrelease", "ballrelease"
'trTrough.addsw 2, 9
trTrough.addsw 1, 11
trTrough.addsw 0, 10
trTrough.initexit BallRelease, 90, 10
trTrough.StackExitBalls = 1
trTrough.MaxBallsPerKick = 1
trTrough.Reset

'----------------------------------------------
' Setup Desktop elements if DeskTopEnabled = True
If DeskTopEnabled Then
	DisplayTimer7.enabled = True
	P1D1.visible = 1
	P1D2.visible = 1
	P1D3.visible = 1
	P1D4.visible = 1
	P1D5.visible = 1
	P1D6.visible = 1
	P1D7.visible = 1
	P2D1.visible = 1
	P2D2.visible = 1
	P2D3.visible = 1
	P2D4.visible = 1
	P2D5.visible = 1
	P2D6.visible = 1
	P2D7.visible = 1
	P3D1.visible = 1
	P3D2.visible = 1
	P3D3.visible = 1
	P3D4.visible = 1
	P3D5.visible = 1
	P3D6.visible = 1
	P3D7.visible = 1
	P4D1.visible = 1
	P4D2.visible = 1
	P4D3.visible = 1
	P4D4.visible = 1
	P4D5.visible = 1
	P4D6.visible = 1
	P4D7.visible = 1
	BaD1.visible = 1
	BaD2.visible = 1
	CrD1.visible = 1
	CrD2.visible = 1

Else
	DisplayTimer7.enabled = False
	P1D1.visible = 0
	P1D2.visible = 0
	P1D3.visible = 0
	P1D4.visible = 0
	P1D5.visible = 0
	P1D6.visible = 0
	P1D7.visible = 0
	P2D1.visible = 0
	P2D2.visible = 0
	P2D3.visible = 0
	P2D4.visible = 0
	P2D5.visible = 0
	P2D6.visible = 0
	P2D7.visible = 0
	P3D1.visible = 0
	P3D2.visible = 0
	P3D3.visible = 0
	P3D4.visible = 0
	P3D5.visible = 0
	P3D6.visible = 0
	P3D7.visible = 0
	P4D1.visible = 0
	P4D2.visible = 0
	P4D3.visible = 0
	P4D4.visible = 0
	P4D5.visible = 0
	P4D6.visible = 0
	P4D7.visible = 0
	BaD1.visible = 0
	BaD2.visible = 0
	CrD1.visible = 0
	CrD2.visible = 0
end If

End Sub


Sub FirePower2_KeyDown(ByVal keycode)

	If KeyDownHandler(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Pullback
	If keycode =  RightFlipperKey Then:Controller.Switch(57) = 1:End If
	If keycode = LeftFlipperKey Then:Controller.Switch(58) = 1:End If
	'If keycode = CoinIn Then PlaySound "Coin"

End Sub

Sub FirePower2_KeyUp(ByVal keycode)

	If KeyUpHandler(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Fire:PlaySound"plunger"
	If keycode = RightFlipperKey Then:Controller.Switch(57) = 0:End If
	If keycode = LeftFlipperKey Then:Controller.Switch(58) = 0:End If

End Sub

'==================================================================
'Setup Solenoids
'==================================================================
SolCallback(cOutHole) = "trTrough.SolIn"
SolCallback(cBallRampThrower) = "trTrough.SolOut"
SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"
SolCallback(cEjectHole) = "SolEjectHole"
SolCallback(cOrbitFlasher) = "FLOrbitFlasher"
SolCallback(cBell) = "SolBell"
SolCallback(cGIRelay) = "SolGIRelay"

Sub SolBell(enabled)
	If enabled Then
		Playsound "Ding_01", 0, 1, -0.05, 0.05
	End If
End Sub

Dim lt
Sub SolGIRelay(enabled)
	If enabled Then
		For each lt In GI
			lt.state = lightstateoff
		Next
	Else
		For each lt In GI
			lt.state = lightstateOn
		Next
	End If
End Sub

sub FLOrbitFlasher(enabled)
	if enabled Then
	light62.state = LightStateOn
	Else
	light62.state = lightstateoff
	End If
End Sub

'=========================================================================
' Flipper Solenoid Handlers
'=========================================================================
sub SolRFlipper(enabled)
	if enabled Then
		PlaySound "FPFlipperUp", 0, .67, -0.05, 0.05
		RightFlipper.RotateToEnd

	Else
		PlaySound "FPFlipperDown", 0, 1, -0.05, 0.05
		RightFlipper.RotateToStart
	end If
end Sub


sub SolLFlipper(enabled)
	if enabled Then
		PlaySound "FPFlipperUp", 0, .67, -0.05, 0.05
		LeftFlipper.RotateToEnd
	Else
		PlaySound "FPFlipperD", 0, 1, -0.05, 0.05
		LeftFlipper.RotateToStart
	end If
end Sub
'===========================================================================
' Other Solenoid Handlers
'============================================================
sub SolEjectHole(enabled)
	if enabled then PlaySound "FPBallRelease", 0, .67,-0.05,0.05
	Kicker1.Kick -90, 15
end Sub

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
	vpmTimer.PulseSw sRightSlingshot
    PlaySound "right_slingshot", 0, 1, 0.05, 0.05
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
	vpmTimer.PulseSw sLeftSlingshot
    PlaySound "left_slingshot",0,1,-0.05,0.05
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
	PlaySound "FPtarget", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
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


'======================================================================================
' Mesh gate animations based on overlated VPX gates with visible = 0
'======================================================================================
dim gate2angle
Sub Gate2_Timer():gate2angle = 180 - gate2.CurrentAngle:LeftRampGate.RotX = -gate2angle:End Sub

dim gate3angle
Sub Gate3_Timer():gate3angle = 180 - gate3.CurrentAngle:RightRampGate.RotX = -gate3angle:End Sub

dim gate4angle
Sub Gate4_Timer():gate4angle = 180 - gate4.CurrentAngle:SpinnerGate.RotX = gate4angle:End Sub

dim gate1angle
Sub Gate1_Timer():gate1angle = 180 - gate1.CurrentAngle:ShooterLaneGate.RotX = gate1angle:End Sub
'======================================================================================

Sub sw12_Hit():Controller.Switch(sBallShooter) = 1:End Sub

Sub sw12_Unhit():Controller.Switch(sBallShooter) = 0:End Sub

Sub sw34_Hit():vpmTimer.PulseSw sA:End Sub

Sub sw35_Hit():vpmTimer.PulseSw sB:End Sub

Sub sw36_Hit():vpmTimer.PulseSw sC:End Sub

Sub sw37_Hit():vpmTimer.PulseSw sD:End Sub

Sub Kicker1_Hit():Controller.Switch(sEjectHole) = 1:End Sub

Sub Kicker1_Unhit():Controller.Switch(sEjectHole) = 0:End Sub

Sub sw38_Hit():vpmTimer.PulseSw sRightOutlane:End Sub

Sub sw45_Hit():vpmTimer.PulseSw sLeftReturnLane:End Sub

Sub sw46_Hit():vpmTimer.PulseSw sRightReturnLane:End Sub

Sub sw39_Hit():vpmTimer.PulseSw sRightOutlane:End Sub

Sub Gate2_Hit()
	vpmTimer.PulseSw sRampInRollUnder
	Playsound "FPGate", 0, .67,-0.05,0.05
End Sub

Sub targetF25_Hit():vpmTimer.PulseSw sFtgt:End Sub

Sub targetI26_Hit():vpmTimer.PulseSw sItgt:End Sub

Sub targetR27_Hit():vpmTimer.PulseSw sRtgt:End Sub

Sub targetE2_Hit():vpmTimer.PulseSw sEtgt:End Sub

Sub targetP29_Hit():vpmTimer.PulseSw sPtgt:End Sub

Sub targetO30_Hit():vpmTimer.PulseSw sOtgt:End Sub

Sub targetW31_Hit():vpmTimer.PulseSw sWtgt:End Sub

Sub targetE232_Hit():vpmTimer.PulseSw sEtgt2:End Sub

Sub targetR233_Hit():vpmTimer.PulseSw sRtgt2:End Sub

Sub targetCenter24_Hit():vpmTimer.PulseSw sReleaseTarget:End Sub

Sub Gate1_Hit():Playsound "FPGate", 0, .67,-0.05,0.05:End Sub

Sub Trigger1_Hit()
	Wall32.IsDropped = 1
	OrbitGate.RotY = 10
	vpmTimer.PulseSw sOrbitOutGate
End Sub

Sub Trigger2_Hit()
	if Wall32.IsDropped Then
		Wall32.IsDropped = False
		OrbitGate.Roty = 0
	End If
End Sub

Sub Gate3_Hit():Playsound "FPGate", 0, .67,-0.05,0.05:vpmtimer.PulseSw sRampOutRollUnder:End Sub

Sub Gate6_Hit():Playsound "FPGate", 0, .67,-0.05,0.05:End Sub

Sub Gate5_Hit():Playsound "FPGate", 0, .67,-0.05,0.05:End Sub

Sub Gate4_Hit():Playsound "FPGate", 0, .67,-0.05,0.05:vpmTimer.PulseSw sOrbitInRollUnder:End Sub

Sub Spinner1_Spin():Playsound "FPSpinner", 0, .67,-0.05,0.05:vpmTimer.PulseSw sSpinner:End Sub

Sub RSling19_Hit():vpmTimer.PulseSw sUpperLeftStandup:End Sub

Sub RSling27_Hit():vpmTimer.PulseSw sUpperRightStandup:End Sub

Sub RSling15_Hit():vpmTimer.PulseSw sCenterLeftStandup:End Sub

Sub RSling17_Hit():vpmTimer.PulseSw sStandup:End Sub

Sub Bumper1_Hit():vpmTimer.PulseSw sUpperLeftBumper:Playsound "FPBumper", 0, .67,-0.05,0.05:End Sub

Sub Bumper2_Hit():vpmTimer.PulseSw sUpperRightBumper:Playsound "FPBumper", 0, .67,-0.05,0.05:End Sub

Sub Bumper4_Hit():vpmTimer.PulseSw sLowerLeftBumper:Playsound "FPBumper", 0, .67,-0.05,0.05:End Sub

Sub Bumper3_Hit():vpmtimer.pulsesw sLowerRightBumper:Playsound "FPBumper", 0, .67,-0.05,0.05:End Sub

Sub Trigger3_Hit()
	Playsound "Wire Ramp", 0, 1, -0.05, 0.05
End Sub

Sub RSling18_Hit():End Sub

Sub Rubbers_Dropped(Index)

End Sub

Sub Trigger4_Hit()
	Playsound "Wire Ramp", 0, 1, -0.05, 0.05
End Sub

 '=========================================================
'                    LED Handling
'=========================================================
'Modified version of Scapino's LED code for Fathom
'and borrowed from Uncle Willy's VP9 Firepower
'
Dim SixDigitOutput(32)
Dim SevenDigitOutput(32)
Dim DisplayPatterns(11)
Dim DigStorage(32)


'Binary/Hex Pattern Recognition Array
DisplayPatterns(0) = 0		'0000000 Blank
DisplayPatterns(1) = 63		'0111111 zero
DisplayPatterns(2) = 6		'0000110 one
DisplayPatterns(3) = 91		'1011011 two
DisplayPatterns(4) = 79		'1001111 three
DisplayPatterns(5) = 102	'1100110 four
DisplayPatterns(6) = 109	'1101101 five
DisplayPatterns(7) = 125	'1111101 six
DisplayPatterns(8) = 7		'0000111 seven
DisplayPatterns(9) = 127	'1111111 eight
DisplayPatterns(10)= 111	'1101111 nine

'Assign 7-digit output to reels
Set SevenDigitOutput(0)  = P1D7
Set SevenDigitOutput(1)  = P1D6
Set SevenDigitOutput(2)  = P1D5
Set SevenDigitOutput(3)  = P1D4
Set SevenDigitOutput(4)  = P1D3
Set SevenDigitOutput(5)  = P1D2
Set SevenDigitOutput(6)  = P1D1

Set SevenDigitOutput(7)  = P2D7
Set SevenDigitOutput(8)  = P2D6
Set SevenDigitOutput(9)  = P2D5
Set SevenDigitOutput(10) = P2D4
Set SevenDigitOutput(11) = P2D3
Set SevenDigitOutput(12) = P2D2
Set SevenDigitOutput(13) = P2D1

Set SevenDigitOutput(14) = P3D7
Set SevenDigitOutput(15) = P3D6
Set SevenDigitOutput(16) = P3D5
Set SevenDigitOutput(17) = P3D4
Set SevenDigitOutput(18) = P3D3
Set SevenDigitOutput(19) = P3D2
Set SevenDigitOutput(20) = P3D1

Set SevenDigitOutput(21) = P4D7
Set SevenDigitOutput(22) = P4D6
Set SevenDigitOutput(23) = P4D5
Set SevenDigitOutput(24) = P4D4
Set SevenDigitOutput(25) = P4D3
Set SevenDigitOutput(26) = P4D2
Set SevenDigitOutput(27) = P4D1

Set SevenDigitOutput(28) = CrD2
Set SevenDigitOutput(29) = CrD1
Set SevenDigitOutput(30) = BaD2
Set SevenDigitOutput(31) = BaD1

Sub DisplayTimer7_Timer ' 7-Digit output
	On Error Resume Next
	Dim ChgLED,ii,chg,stat,obj,TempCount,temptext,adj

	ChgLED = Controller.ChangedLEDs(&HFF, &HFFFF) 'hex of binary (display 111111, or first 6 digits)

	If Not IsEmpty(ChgLED) Then
		For ii = 0 To UBound(ChgLED)
			chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
			For TempCount = 0 to 10
				If stat = DisplayPatterns(TempCount) then
					If LedStatus = 2 Then SevenDigitOutput(chgLED(ii, 0)).SetValue(TempCount)
					DigStorage(chgLED(ii, 0)) = TempCount
				End If
				If stat = (DisplayPatterns(TempCount) + 128) then
					If LedStatus = 2 Then SevenDigitOutput(chgLED(ii, 0)).SetValue(TempCount)
					DigStorage(chgLED(ii, 0)) = TempCount
				End If
			Next
		Next
	End IF
End Sub


'*DOF method for non rom controller tables by Arngrim****************
'*******Use DOF 1**, 1 to activate a ledwiz output*******************
'*******Use DOF 1**, 0 to deactivate a ledwiz output*****************
'*******Use DOF 1**, 2 to pulse a ledwiz output**********************
Sub DOF(dofevent, dofstate)
	If B2SOn=True Then
		If dofstate = 2 Then
			Controller.B2SSetData dofevent, 1:Controller.B2SSetData dofevent, 0
		Else
			Controller.B2SSetData dofevent, dofstate
		End If
	End If
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
Sub FirePower2_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


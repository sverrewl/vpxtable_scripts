' Robo-War (Gottlieb 1988)
' =======================================================
' VP9 Version by Kevin Lee Drum
' B2B Collision thanks to jimmyfingers
' Thanks to the VP8 authors TAB, Destruk, MNPG
' Much was also inspired by JPSalas and unclewilly

' VPX Version By Dozer - January 2017.

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' No special SSF tweaks yet.

Option Explicit

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Dim DesktopMode:DesktopMode = Table1.ShowDT

'GRAPHICS OPTIONS
Const BallFlare =1 'Render a lense flare effect above the ball when certain flashers are active.
Const Playfield_Dimples = 1 ' Render a texture over the playfield to give it some depth.
'------------------------------------

LoadVPM "01210000", "sys80.vbs", 3.1

' Variables
' ==================================================================
Const cGameName = "robowars"
Const UseSolenoids = 2
Const UseLamps = 0
Const UseGI = 0
Const UseSync = 0
Const HandleMech = 0
Dim bsTop, bsTrough, dtAlpha1, dtAlpha2, dtAlpha3, dtBeta,gion
Dim QRelay, OldQRelay, TRelay, OldTRelay, SRelay,OldSRelay, ARelay, OldARelay, BallRel, OldBallRel, BallLaunched, StargateRelay, OldStargateRelay
Dim AuxLightsStep, BlinkGIStep, LeftSlingStep, RightSlingStep, UpOrDown, PlungeWidth, PlungeStep, PlungeRamps, PDirection
' Standard Sounds
' ==================================================================
' All the sounds in this table are tweaked versions of sounds from the VPForums sound library.
Const SSolenoidOn = "Solenoid"
Const SSolenoidOff = ""
Const SFlipperOn = "FlipperUp"
Const SFlipperOff = "FlipperDown"
Const SCoin = "Coin"

' Table Init
' ==================================================================
Sub Table1_Init
	vpmInit Me 'new
	With Controller
		.GameName = cGameName
		.Games(cGameName).Settings.Value("rol") = 0
		.Games(cGameName).Settings.Value("dmd_red") = 0
		.Games(cGameName).Settings.Value("dmd_green") = 223
		.Games(cGameName).Settings.Value("dmd_blue") = 223
		.SplashInfoLine = "Robo-War (Gottlieb 1988)" & vbNewLine & "VPM table by Kevin Lee Drum"
		.HandleMechanics = 0
		.HandleKeyboard = 0
		.ShowTitle = 0
		.ShowDMDOnly = 1
		.ShowFrame = 0
		If DesktopMode AND B2SOn then
			.Hidden = 0
		Else
			If B2SOn then
				.Hidden = 1
			Else
				.Hidden = 0
			End If
        End If
		If Err Then MsgBox Err.Description
		On Error Resume Next
	End With

	On Error Goto 0
	Controller.SolMask(0) = 0
	vpmTimer.AddTimer 2000, "Controller.SolMask(0) = &Hffffffff'" 'ignore all solenoids - then add the Timer to renable all the solenoids after 2 seconds
	Controller.Run GetPlayerHWnd

	' Timers
	PinMAMETimer.Interval = PinMAMEInterval
	PinMAMETimer.Enabled = 1

	' Nudging
	vpmNudge.TiltSwitch = 57
	vpmNudge.Sensitivity = 1
	vpmNudge.TiltObj = Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingshot)

	' Ball Stacks
	Set bsTop = New cvpmBallStack
	With bsTop
		.InitSaucer Kicker,76,140,5
		.InitExitSnd SoundFX("Popper_Ball",DOFContactors), SoundFX(SSolenoidOn,DOFContactors)
	End With

	Set bsTrough = New cvpmBallStack
	With bsTrough
		.InitNoTrough BallRelease,66,45,10
		.InitSw 66,0,56,0,0,0,0,0
		.InitKick BallRelease,45,10
		.InitExitSnd SoundFX("BallRelease",DOFContactors), SoundFX(SSolenoidOn,DOFContactors)
		'.BallImage = "BallDark"
		.Balls = 2
	End With

	' Drop Targets
	Set dtAlpha1 = New cvpmDropTarget
	dtAlpha1.InitDrop Alpha1,43
	dtAlpha1.InitSnd SoundFX("TargetDrop",DOFDropTargets), SoundFX(SSolenoidOn,DOFDropTargets)
	dtAlpha1.CreateEvents "dtAlpha1"

	Set dtAlpha2 = New cvpmDropTarget
	dtAlpha2.InitDrop Array(Alpha2A,Alpha2B), Array(42,52)
	dtAlpha2.InitSnd SoundFX("TargetDrop",DOFDropTargets), SoundFX(SSolenoidOn,DOFDropTargets)
	dtAlpha2.CreateEvents "dtAlpha2"

	Set dtAlpha3 = New cvpmDropTarget
	dtAlpha3.InitDrop Array(Alpha3A,Alpha3B,Alpha3C), Array(41,51,61)
	dtAlpha3.InitSnd SoundFX("TargetDrop",DOFDropTargets), SoundFX(SSolenoidOn,DOFDropTargets)
	dtAlpha3.CreateEvents "dtAlpha3"

	Set dtBeta = New cvpmDropTarget
	dtBeta.InitDrop Array(BetaB,BetaE,BetaT,BetaA), Array(40,50,60,70)
	dtBeta.InitSnd SoundFX("TargetDrop",DOFDropTargets), SoundFX(SSolenoidOn,DOFDropTargets)
	dtBeta.CreateEvents "dtBeta"

	' GI Option
	' DOF 150,1
twos = 0
gion = 1
End Sub

If Table1.ShowDT = true then

Ramp16.visible = 1
Ramp15.visible = 1
 else

Ramp16.visible = 0
Ramp15.visible = 0
 End If

If Playfield_Dimples = 1 Then
LTFDim.Visible = 1
RTFDim.Visible = 1
PFDim.Visible = 1
else
LTFDim.Visible = 0
RTFDim.Visible = 0
PFDim.Visible = 0
End If

' Keys, etc.
' ==================================================================
Sub Table1_Paused:Controller.Pause = True:End Sub
Sub Table1_unPaused:Controller.Pause = False:End Sub
Sub Table1_KeyDown(ByVal keycode)
	If KeyCode = RightFlipperKey Then Controller.Switch(46)=1
    'If keycode = 3 Then Msgbox Activeball.Y
	If KeyCode = PlungerKey Then
		Plunger.Pullback
		'PDirection = 1:Plunge.Interval = 40:Plunge.Enabled = 1 ' Plunger pull-back animation
	End If
	If vpmKeyDown(KeyCode) Then Exit Sub
	If KeyCode = KeyRules then Rules
End Sub
Sub Table1_KeyUp(ByVal keycode)
	If KeyCode = RightFlipperKey Then Controller.Switch(46)=0
	If KeyCode = PlungerKey Then
		Plunger.Fire:PlaySound "Plunger"
		'PDirection = -1:Plunge.Interval = 6:Plunge.Enabled = 1 ' Plunger release animation
	End If
	If vpmKeyUp(KeyCode) Then Exit Sub
End Sub

' Switches
' ==================================================================
' Slingshots

Dim RStep, Lstep

Sub RightSlingShot_Slingshot
    PlaySound SoundFXDOF("right_slingshot",102,DOFPulse,DOFContactors), 0, 1, 0.05, 0.05
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
    RightSlingShot.TimerInterval = 10
	vpmTimer.PulseSw 33
    DOF 111,DOFPulse
    RSS.opacity = 60
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
    RSS.opacity = 0
End Sub

Sub LeftSlingShot_Slingshot
    PlaySound SoundFXDOF("left_slingshot",101,DOFPulse,DOFContactors),0,1,-0.05,0.05
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
    LeftSlingShot.TimerInterval = 10
	vpmTimer.PulseSw 33
    LSS.opacity = 60
    DOF 110, DOFPulse
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
    LSS.opacity = 0
End Sub

' Bumpers
Sub Bumper1_Hit:PlaySound SoundFX("Bumper",DOFContactors):vpmTimer.PulseSw 53:End Sub
Sub Bumper2_Hit:PlaySound SoundFX("Bumper",DOFContactors):vpmTimer.PulseSw 63:End Sub
Sub Bumper3_Hit:PlaySound SoundFX("Bumper",DOFContactors):vpmTimer.PulseSw 73:End Sub
' Drain
Sub Drain_Hit:PlaySound "Drain":bsTrough.AddBall Me:End Sub
' Rollovers
Sub LaneR_Hit:Controller.Switch(44) = 1:				End Sub
Sub LaneR_UnHit:Controller.Switch(44) = 0:				End Sub
Sub LaneO_Hit:Controller.Switch(54) = 1:				End Sub
Sub LaneO_UnHit:Controller.Switch(54) = 0:				End Sub
Sub LaneB_Hit:Controller.Switch(64) = 1:				End Sub
Sub LaneB_UnHit:Controller.Switch(64) = 0:				End Sub
Sub LaneO2_Hit:Controller.Switch(74) = 1:				End Sub
Sub LaneO2_UnHit:Controller.Switch(74) = 0:				End Sub
Sub LeftOutlane_Hit:Controller.Switch(45) = 1:			End Sub
Sub LeftOutlane_UnHit:Controller.Switch(45) = 0:		End Sub
Sub LeftInlane_Hit:Controller.Switch(55) = 1:			End Sub
Sub LeftInlane_UnHit:Controller.Switch(55) = 0:			End Sub
Sub RightInlane_Hit:Controller.Switch(65) = 1:			End Sub
Sub RightInlane_UnHit:Controller.Switch(65) = 0:		End Sub
Sub RightOutlane_Hit:Controller.Switch(75) = 1:			End Sub
Sub RightOutlane_UnHit:Controller.Switch(75) = 0:		End Sub
Sub RightLane_Hit:Controller.Switch(35) = 1:			End Sub
Sub RightLane_UnHit:Controller.Switch(35) = 0:			End Sub
Sub ShooterLane_Hit:Controller.Switch(36) = 1:			End Sub
Sub ShooterLane_UnHit:Controller.Switch(36) = 0:    	End Sub
Sub TopRightRollover_Hit:Controller.Switch(31) = 1:		End Sub
Sub TopRightRollover_UnHit:Controller.Switch(31) = 0:	End Sub
Sub StargateRollover_Hit:Controller.Switch(71) = 1:		End Sub
Sub StargateRollover_UnHit:Controller.Switch(71) = 0:	End Sub
' Spinner
Sub Spinner_Spin:vpmTimer.PulseSw 30:Playsound "fx_spinner":End Sub
' Drop Targets
Sub Alpha1_Hit:dtAlpha1.hit 1:End Sub
Sub Alpha2A_Hit:dtAlpha2.hit 1:End Sub
Sub Alpha2B_Hit:dtAlpha2.hit 2:End Sub
Sub Alpha3A_Hit:dtAlpha3.hit 1:End Sub
Sub Alpha3B_Hit:dtAlpha3.hit 2:End Sub
Sub Alpha3C_Hit:dtAlpha3.hit 3:End Sub
Sub BetaB_Hit:dtBeta.hit 1:End Sub
Sub BetaE_Hit:dtBeta.hit 2:End Sub
Sub BetaT_Hit:dtBeta.hit 3:End Sub
Sub BetaA_Hit:dtBeta.hit 4:End Sub
' Top Saucer
Sub Kicker_Hit:bsTop.AddBall 0:PlaySound "kicker_enter":End Sub
Sub Kicker_Unhit:DOF 113, DOFPulse:End Sub
Sub Kicker1_Hit():PlaySound "Ball_Bounce":End Sub
Sub Trigger1_hit():PlaySound "Launch":End Sub
' Targets
Sub StargateTarget_Hit:vpmTimer.PulseSw 72:PlaySound "Bounce":End Sub
Sub LeftTarget_Hit:vpmTimer.PulseSw 62:PlaySound "Bounce":End Sub
' Trapped Ball Untrapper
'Sub Untrapper1_Hit:Me.DestroyBall:Untrapper2.Enabled=1:Untrapper2.CreateBall:Untrapper2.Kick 90,10:Controller.B2SSetData 114,1:Controller.B2SSetData 114,0:End Sub
' Ball Image Changers
'Sub BallChanger_Hit:ActiveBall.Image = "Ball":End Sub
'Sub BallChangerRed_Hit:ActiveBall.Image = "BallRed":End Sub
'Sub BallChangerRed_UnHit:ActiveBall.Image = "Ball":End Sub
' Sound Effects Only
'Sub Bounce_Hit(parm):PlaySound "Bounce":End Sub
'Sub LeftFlipper_Collide(parm):PlaySound "Bounce":End Sub
'Sub RightFlipper_Collide(parm):PlaySound "Bounce":End Sub
'Sub GateTopLeft_Hit:PlaySound "Gate":End Sub
'Sub GateTopRight_Hit:PlaySound "Gate":End Sub
'Sub GateRightFlipper_Hit:PlaySound "Gate":End Sub
'Sub BallRollSound_Hit:PlaySound "BallRoll3":End Sub
'Sub DrainSound_Hit:PlaySound "BallRoll1":End Sub
' Ball Launched (Used for GI Control)
'Sub GateLane_Hit:BallLaunched = 1:End Sub

' Solenoids
' ==================================================================
SolCallback(1) = "SolOne"										' 1 Reset Alpha I / Top Left Flasher 3
SolCallback(2) = "SolTwo"										' 2 Reset Alpha II / Top Left Flasher 2
SolCallback(3) = "SolThree"										' 3 Top Kicker / Stargate Flashers
SolCallback(4) = "SolFour"		' 4 Right Under-Plastic Flasher
SolCallback(5) = "SolFive"										' 5 Reset Alpha III / Top Left Flasher 1
SolCallback(6) = "SolSix"										' 6 Reset Beta / Top Right Flasher
SolCallback(7) = "SolSeven"		' 7 Left Under-Plastic Flasher
SolCallback(8) = "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(9) = "bsTrough.SolIn"
'SolCallback(sLLFlipper) = "vpmSolFlipper LeftFlipper,nothing,"
'SolCallback(sLRFlipper) = "vpmSolFlipper RightFlipper,nothing,"

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
	If Enabled Then
		PlaySound SoundFX("FlipperUp",DOFFlippers):LeftFlipper.RotateToEnd
	Else
		PlaySound SoundFX("FlipperDown",DOFFlippers):LeftFlipper.RotateToStart
	End If
End Sub

Sub SolRFlipper(Enabled)
	If Enabled Then
		PlaySound SoundFX("FlipperUp",DOFFlippers):RightFlipper.RotateToEnd
	Else
		PlaySound SoundFX("FlipperDown",DOFFlippers):RightFlipper.RotateToStart
	End If
End Sub

Dim ltxx
Sub SolOne(Enabled)
	If Enabled Then
		If SRelay = 0 Then
			dtAlpha1.DropSol_On
            DOF 115, DOFPulse
		Else
            For each ltxx in LTFlash:ltxx.state=1:Next
			SetLamp 163,1
		End If
	Else
		'If SRelay = 1 Then
        SetLamp 163,0
        For each ltxx in LTFlash:ltxx.state=0:Next
	'End If
End If
End Sub

Dim lmxx, twos
Sub SolTwo(Enabled)
	If Enabled Then
		If SRelay = 0 Then
			dtAlpha2.DropSol_On
            DOF 116, DOFPulse
		Else
			SetLamp 162,1:twos = 1
            For each lmxx in LMFlash:lmxx.state=1:Next
            If LampState(1) = 1 Then
            LSL.state = 1
            End If
		End If
	Else
		'If SRelay = 1 Then
        SetLamp 162,0:twos = twos = 0
        For each lmxx in LMFlash:lmxx.state=0:Next
        LSL.State = 0
	    'End If
End If
End Sub

Sub SolThree(Enabled)
	If Enabled Then
		If SRelay = 0 Then
			bsTop.ExitSol_On
		Else
			SGP.state = 1
            SGP1.state = 1
            SGP2.state = 1
		End If
	Else
            SGP.state = 0
            SGP1.state = 0
            SGP2.state = 0
		If SRelay = 0 Then
		End If
	End If
End Sub

Sub SolFour(Enabled)
	If Enabled Then
			RPF.state = 1
            RPF1.state = 1
            RPF2.state = 1
		Else
            RPF.state = 0
            RPF1.state = 0
            RPF2.state = 0
	End If
End Sub

Sub SolSeven(Enabled)
	If Enabled Then
			LPF.state = 1
            LPF1.state = 1
            LPF2.state = 1
		Else
            LPF.state = 0
            LPF1.state = 0
            LPF2.state = 0
End If
End Sub

Dim lbxx

Sub SolFive(Enabled)
	If Enabled Then
		If SRelay = 0 Then
			dtAlpha3.DropSol_On
            DOF 117, DOFPulse
		Else
             For each lbxx in LBFlash:lbxx.state=1:Next
			 SetLamp 161,1
		End If
	Else
		'If SRelay = 1 Then
        SetLamp 161,0
        For each lbxx in LBFlash:lbxx.state=0:Next
	'End If
End If
End Sub

Dim rtxx

Sub SolSix(Enabled)
	If Enabled Then
		If SRelay = 0 Then
			dtBeta.DropSol_On
            DOF 118, DOFPulse
		Else
			SetLamp 164,1:twos = 1
            For each rtxx in RTFlash:rtxx.state=1:Next
            If LampState(1) = 1 Then
            RSL.state = 1
            End If
		End If
	Else
		'If SRelay = 1 Then
        SetLamp 164,0:twos = 0
        For each rtxx in RTFlash:rtxx.state=0:Next
        RSL.State = 0
	End If
End Sub

' Relays/Lamp Events
' ==================================================================
OldQRelay = 0:OldTRelay = 0:OldARelay = 0:OldBallRel = 0:OldSRelay = 1:OldStargateRelay = 1:UpOrDown = 1
Set LampCallback = GetRef("UpdateRelays")
Sub UpdateRelays
	SRelay = LampState(12) ' Used for solenoid multiplexing.
	If SRelay <> OldSRelay Then OldSRelay = SRelay

	'If GIAlwaysOn = 0 Then
		QRelay = LampState(0)
		If QRelay <> OldQRelay Then
			If QRelay = 0 Then
				'SetLamp 101, 0 ' Turn off the GI when not playing ('Q' Relay is off).
			End If
			'BallLaunched = 0
			OldQRelay = QRelay
		'End If

		TRelay = LampState(1)
		If TRelay <> OldTRelay Then
			If QRelay = 1 Then
				If TRelay = 0 Then
					If SRelay = 1 Then
						SetLamp 101, 1 ' Turn on the GI when the 'T' Relay is off during play.
					End If
				Else
                        SetLamp 101, 0
				End If
			End If
			OldTRelay = TRelay
		End If
	End If

	ARelay = LampState(13)
	If ARelay <> OldARelay Then
		If ARelay = 1 Then
			AuxLights.Enabled = 1 ' Turn on aux light sequence when 'A' Relay is on.
		Else
			AuxLights.Enabled = 0
			Dim x, y
			For x = 71 To 100 ' Turn off all aux lights when 'A' Relay is off.
				SetLamp x, 0
			Next
		End If
	End If
	OldARelay = ARelay

	BallRel = LampState(2)
	If BallRel <> OldBallRel Then
		If BallRel = 0 Then
			If bsTrough.Balls Then bsTrough.ExitSol_On ' Release ball when relay 2 is on.
		End If
	End If
	OldBallRel = BallRel

	StargateRelay = LampState(14)
	If StargateRelay <> OldStargateRelay Then
		If StargateRelay = 0 Then ' Ramp is down when relay 14 is off.
			UpOrDown = -1
			If StargateRamp.HeightBottom <> 0 Then RampMove.Enabled = 1
			StargateInvisible.Collidable = 1
			Untrapper1.Enabled = 1
		End If
		If StargateRelay = 1 Then ' Ramp is up when relay 14 is on.
			UpOrDown = 1
			If StargateRamp.HeightBottom <> 60 Then PlaySound SoundFX("RampUp",DOFcontactors)
			If StargateRamp.HeightBottom <> 60 Then RampMove.Enabled = 1
			StargateInvisible.Collidable = 0
		End If
	OldStargateRelay = StargateRelay
	End If

End Sub

' Auxiliary Lights Timer (Playfield "Robo-Units" and Rear Chaser Lights)
' ==================================================================
' Thanks to Destruk for the lamp schematic!
AuxLightsStep = 0
Sub AuxLights_Timer
	AuxLightsStep = AuxLightsStep + 1
	Select Case AuxLightsStep
		Case 1
			SetLamp 100,0:SetLamp 89,0:SetLamp 90,0
			SetLamp 91,1:SetLamp 71,1:SetLamp 72,1
DOF 100,0
DOF 89,0
DOF 90,0
DOF 91,2
DOF 71,2
DOF 72,2
		Case 2
			SetLamp 91,0:SetLamp 71,0:SetLamp 72,0
			SetLamp 92,1:SetLamp 73,1:SetLamp 74,1
DOF 91,0
DOF 71,0
DOF 72,0
DOF 92,2
DOF 73,2
DOF 74,2
		Case 3
			SetLamp 92,0:SetLamp 73,0:SetLamp 74,0
			SetLamp 93,1:SetLamp 75,1:SetLamp 76,1
DOF 92,0
DOF 73,0
DOF 74,0
DOF 93,2
DOF 75,2
DOF 76,2
		Case 4
			SetLamp 93,0:SetLamp 75,0:SetLamp 76,0
			SetLamp 94,1:SetLamp 77,1:SetLamp 78,1
DOF 93,0
DOF 75,0
DOF 76,0
DOF 94,2
DOF 77,2
DOF 78,2
		Case 5
			SetLamp 94,0:SetLamp 77,0:SetLamp 78,0
			SetLamp 95,1:SetLamp 79,1:SetLamp 80,1
DOF 94,0
DOF 77,0
DOF 78,0
DOF 95,2
DOF 79,2
DOF 80,2
		Case 6
			SetLamp 95,0:SetLamp 79,0:SetLamp 80,0
			SetLamp 96,1:SetLamp 81,1:SetLamp 82,1
DOF 95,0
DOF 79,0
DOF 80,0
DOF 96,2
DOF 81,2
DOF 82,2
		Case 7
			SetLamp 96,0:SetLamp 81,0:SetLamp 82,0
			SetLamp 97,1:SetLamp 83,1:SetLamp 84,1
DOF 96,0
DOF 81,0
DOF 82,0
DOF 97,2
DOF 83,2
DOF 84,2
		Case 8
			SetLamp 97,0:SetLamp 83,0:SetLamp 84,0
			SetLamp 98,1:SetLamp 85,1:SetLamp 86,1
DOF 97,0
DOF 83,0
DOF 84,0
DOF 98,2
DOF 85,2
DOF 86,2
		Case 9
			SetLamp 98,0:SetLamp 85,0:SetLamp 86,0
			SetLamp 99,1:SetLamp 87,1:SetLamp 88,1
DOF 98,0
DOF 85,0
DOF 86,0
DOF 99,2
DOF 87,2
DOF 88,2
		Case 10
			SetLamp 99,0:SetLamp 87,0:SetLamp 88,0
			SetLamp 100,1:SetLamp 89,1:SetLamp 90,1
DOF 99,0
DOF 87,0
DOF 88,0
DOF 100,2
DOF 89,2
DOF 90,2
	End Select
	If AuxLightsStep = 10 Then AuxLightsStep = 0
'DOF 93,0
End Sub

Set MotorCallback = GetRef("GameTimer")

Sub GameTimer
    'RollingSound
End Sub

'*****************************************
'	Ball Shadow
'*****************************************

Dim BallShadow
BallShadow = Array (BallShadow1, BallShadow2)

Sub BallShadowUpdate()
    Dim BOT, b, shadowZ
    BOT = GetBalls

	' render the shadow for each ball
    For b = 0 to UBound(BOT)
		BallShadow(b).X = BOT(b).X
		BallShadow(b).Y = BOT(b).Y + 20
		'If BOT(b).Z > 90 and BOT(b).Z < 120 Then
		'	BallShadow(b).visible = 1
		'Else
		'	BallShadow(b).visible = 0
		'End If
	Next
End Sub

Dim XBallShadow
XBallShadow = Array (XBallShadow1, XBallShadow2)



Sub XBallShadowUpdate()
    Dim XBOT, c, XshadowZ
    XBOT = GetBalls

	' render the shadow for each ball
    For c = 0 to UBound(XBOT)
		XBallShadow(c).X = XBOT(c).X
		XBallShadow(c).Y = XBOT(c).Y - 10
        If ballflare = 1 AND twos = 1 AND XBOT(c).VelY > 0 AND XBOT(c).Y < 1800 Then
        XBallShadow(c).visible = 1
		Else
		XBallShadow(c).visible = 0
		End If
        Shad_Rot.Interval = NOT XBOT(c).VelY
	Next
End Sub

' Ramp Movement Timer
' ==================================================================
Sub RampMove_Timer
	Dim x:x = 20 * UpOrDown
	StargateRamp.HeightBottom = StargateRamp.HeightBottom + x
	StargateHelper.HeightBottom = StargateHelper.HeightBottom + x
	StargateHelper.HeightTop = StargateHelper.HeightTop + x
	RampRefresh.State = 1:RampRefresh.State = 0
	If StargateRamp.HeightBottom = 0 Or StargateRamp.HeightBottom = 60 Then RampMove.Enabled = 0
End Sub

' Rules
' ==================================================================
' Based on Inkochnito's Reproduction Card and JP's script
Dim Msg(20)
Sub Rules()
	Msg(0) = "HOW TO PLAY" &Chr(10)
	Msg(1) = "ROBO-WAR" &Chr(10) &Chr(10)
	Msg(2) = ""
	Msg(3) = "SPECIAL: ADD A LETTER TO R-O-B-O-W-A-R BY COMPLETING EITHER THE"
	Msg(4) = "TOP ROLLOVERS (R-O-B-O), OR BY HITTING THE STROBING"
	Msg(5) = "DROP TARGETS (B-E-T-A). COMPLETING (R-O-B-O-W-A-R)"
	Msg(6) = "LIGHTS A SPECIAL."
	Msg(7) = ""
	Msg(8) = "EXTRA BALL: COMPLETING THE ALPHA DROP TARGET SEQUENCE LIGHTS"
	Msg(9) = "AN EXTRA BALL."
	Msg(10) = ""
	Msg(11) = "MULTIPLIER: ADVANCE MULTIPLIER ON VARIOUS PLAYFIELD TARGETS"
	Msg(12) = "WHEN LIT. SCORE 10,000 TIMES MULTIPLIER FOR EACH"
	Msg(13) = "LETTER AWARDED IN (R-O-B-O-W-A-R) AT THE END OF A BALL"
	Msg(14) = "IN PLAY."
	Msg(15) = ""
	Msg(16) = "MULTI-BONUS: SCORE 5000 TIMES MULTIPLIER FOR EACH DROP TARGET"
	Msg(17) = "HIT DURING MULTI-BALL PLAY. SCORE MULTI-BONUS VALUE IN"
	Msg(18) = "HOLE DURING MULTI-BALL PLAY AND AFTER LAST BALL"
	Msg(19) = "IN PLAY."
	Msg(20) = ""
	Dim x
	For x = 1 To 20
		Msg(0) = Msg(0) + Msg(X) &Chr(13)
    Next
	MsgBox Msg(0), , "         Instructions and Rule Card"
End Sub

' Dip Switches
' ==================================================================
' Gottlieb Robo-War
' originally added by Inkochnito
' Updated Switches 8, 31, and 32
Sub editDips
	Dim vpmDips : Set vpmDips = New cvpmDips
	With vpmDips
		.AddForm  700,400,"Robo-War - DIP switches"
		.AddFrame 2,4,190,"Maximum Credits",49152,Array("8",0,"10",32768,"15",&H00004000,"20",49152)'dip 15&16
		.AddFrame 2,80,190,"Coin Chute Left and Right Control",&H00002000,Array("Separate",0,"Same",&H00002000)'dip 14
		.AddFrame 2,126,190,"Playfield Special",&H00200000,Array("Special",0,"Extra Ball",&H00200000)'dip 22
		.AddFrame 2,172,190,"Highest Games to Date Control",&H00000020,Array("No Effect",0,"Reset High Games #2-#5 on Power Off",&H00000020)'dip 6
		.AddFrame 2, 218, 190, "Auto-Percentage Control", &H00000080, Array("Disabled (Normal High Score Mode)", 0, "Enabled", &H00000080)'dip 8
        .AddFrame 2, 264, 190, "Alpha Drop Bank Sequence", &H40000000, Array("Also Award ROBOWAR Letter", 0, "Light Extra Ball Only", &H40000000)'dip 31
        .AddFrame 2, 310, 190, "Number of Active ADV X Targets", &H80000000, Array("More", 0, "Less", &H80000000)'dip 32
		.AddFrame 205,4,190,"Highest Game to Date Awards",&H00C00000,Array("None (Not Displayed)",0,"None",&H00800000,"2 Replay",&H00400000,"3 Replay",&H00C00000)'dip 23&24
		.AddFrame 205,80,190,"Balls/Game",&H01000000,Array("5",0,"3",&H01000000)'dip 25
		.AddFrame 205,126,190,"Replay Limit",&H04000000,Array("No Limit",0,"1",&H04000000)'dip 27
		.AddFrame 205,172,190,"Novelty",&H08000000,Array("Normal",0,"Score 500,000 in Place of Extra Ball and Special",&H08000000)'dip 28
		.AddFrame 205,218,190,"Game Mode",&H10000000,Array("Replay",0,"Extra Ball",&H10000000)'dip 29
		.AddFrame 205,264,190,"3rd Coin Chute Credit Control",&H20000000,Array("No Effect",0,"Add 9",&H20000000)'dip 30
		.AddChk 205,316,180,Array("Match",&H02000000)'dip 26
		.AddChk 205,331,190,Array("Attract Mode Sound",&H00000040)'dip 7
		.AddLabel 50,360,300,20,"After hitting OK, press F3 to reset game with new settings."
		.ViewDips
	End With
End Sub
Set vpmShowDips = GetRef("editDips")

Sub Primflip_Timer()
FlipperT1.roty = LeftFlipper.currentangle  + 245
FlipperT5.roty = RightFlipper.currentangle + 116
MiddleWheel.ObjRotz = LeftFlipper.currentangle - 90
MiddleWheel1.ObjRotz = RightFlipper.currentangle - 100
BallShadowUpdate
XBallShadowUpdate
End Sub

'****************************************
'  JP's Fading Lamps v5 VP9 Fading only
'      Based on PD's Fading Lights
' SetLamp 0 is Off
' SetLamp 1 is On
' LampState(x) current state
'****************************************

Dim LampState(200), FadingLevel(200)
Dim FlashSpeedUp(200), FlashSpeedDown(200), FlashMin(200), FlashMax(200), FlashLevel(200)

InitLamps()
LampTimer.Interval = 10
LampTimer.Enabled = 1


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

Sub InitLamps()
    Dim x
    For x = 0 to 200
        LampState(x) = 0        ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4      ' used to track the fading state
        FlashSpeedUp(x) = 0.2   ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.1 ' slower speed when turning off the flasher
        FlashMax(x) = 1         ' the maximum value when on, usually 1
        FlashMin(x) = 0.05         ' the minimum value when off, usually 0
        FlashLevel(x) = 0       ' the intensity of the flashers, usually from 0 to 1
    Next
End Sub

Sub UpdateLamps

	' I had to create lights for each of the relays, or they wouldn't be 100% reliable:
	XNFadeL 0, l0 ' ['Q' Relay] Game Over Relay
	XNFadeL 1, l1 ' ['T' Relay] Tilt Relay
	XNFadeL 2, l2 ' Ball Release
	' NFadeL 4, l4 ' Sound 16
	'XNFadeL 12, l12 ' ['S' Relay] Solenoid Multiplexer
    L12.State = LampState(12)
	XNFadeL 13, l13 ' ['A' Relay] Robo-Unit and Rear Chaser Light Sequences
	XNFadeL 14, l14 ' Stargate Ramp / Launch Attack

	' These lamps will fade slower during gameplay than they do during the attract sequence.
	If LampState(0) = 1 Then
		XNFadeL 3, l3 ' Fight Again
		XNFadeL 15, l15 ' 1x
		XNFadeL 16, l16 ' 2x
		XNFadeL 17, l17 ' 4x
       XNFadeL 18, l18
		XNFadeL 31, l31 ' Base
		XNFadeL 32, l32 ' III (Alpha)
		XNFadeL 33, l33 ' II (Alpha)
		XNFadeL 34, l34 ' I (Alpha)
		XNFadeL 35, l35 ' Special
		XNFadeL 36, l36 ' Special
		XNFadeL 37, l37 ' Special
		XNFadeL 38, l38 ' Special
		XNFadeL 39, l39 ' Adv X
		XNFadeL 40, l40 ' Adv X
		XNFadeL 41, l41 ' Adv X
		XNFadeL 42, l42 ' Adv X
		XNFadeL 43, l43 ' Battle Station
	Else
		XNFadeL 3, l3' Fight Again
		XNFadeL 15, l15 ' 1x
		XNFadeL 16, l16 ' 2x
		XNFadeL 17, l17 ' 4x
		XNFadeL 18, l18
		XNFadeL 31, l31 ' Base
		XNFadeLm 32, l32 ' III (Alpha)
        XNFadeL 32, l32sc ' III (Alpha)
		XNFadeL 33, l33 ' II (Alpha)
		XNFadeL 34, l34 ' I (Alpha)
		XNFadeLm 35, l35 ' Special
        XNFadeL 35, l35sc ' Special
		XNFadeL 36, l36 ' Special
		XNFadeLm 37, l37 ' Special
        XNFadeL 37, l37sc ' Special
		XNFadeLm 38, l38 ' Special
        XNFadeL 38, L38sc
		XNFadeL 39, l39 ' Adv X
		XNFadeL 40, l40 ' Adv X
		XNFadeL 41, l41 ' Adv X
		XNFadeL 42, l42 ' Adv X
		XNFadeL 43, l43 ' Battle Station
	End If

	' The rest are always the same speed.
    XNFadeLm 5, l5
    XNFadeL 5, l5a
    XNFadeLm 6, l6
    XNFadeL 6, l6a
	XNFadeLm 7, l7	' (ro)B(owar)
    XNFadeL 7, l7a
	XNFadeLm 8, l8	' (rob)O(war)
    XNFadeL 8, l8a
	XNFadeLm 9, l9	' (robo)W(ar)
	XNFadeL 9, l9a
    XNFadeLm 10, l10	' (robow)A(r)
    XNFadeL 10, l10a
	XNFadeLm 11, l11	' (robowa)R
    XNFadeL 11, l11a
	XNFadeLm 19, l19	' Extra Ball 1
    XNFadeL 19, l19x	' Extra Ball 1
	XNFadeLm 20, l20	' Extra Ball 2
    XNFadeL 20, l20x	' Extra Ball 2
	XNFadeLm 21, l21	' Extra Ball 3
	XNFadeL 21, l21x	' Extra Ball 2
    XNFadeLm 22, l22	' Extra Ball 4
    XNFadeL 22, l22x	' Extra Ball 2
	XNFadeL 23, l23	' R(obo)
	XNFadeL 24, l24	' (r)O(bo)
	XNFadeL 25, l25	' (ro)B(o)
	XNFadeL 26, l26	' (rob)O
	XNFadeL 27, l27	' B(eta)
	XNFadeL 28, l28	' (b)E(ta)
	XNFadeL 29, l29	' (be)T(a)
	XNFadeL 30, l30	' (bet)A
	XNFadeL 44, l44	' Power Surge

	XNFadeLm 45, X45
    XNFadeLm 45, X45a
    XNFadeLm 45, X45b
    XNFadeLm 45, X45c
    XNFadeLm 45, X45d
    XNFadeLm 45, X45r
    XNFadeLm 45, X45sc
    XNFadeL 45, X45sc1

    XNFadeLm 46, X46
    XNFadeLm 46, X46a
    XNFadeLm 46, X46b
    XNFadeLm 46, X46c
    XNFadeLm 46, X46d
    XNFadeLm 46, X46e
    XNFadeLm 46, X46r
    XNFadeLm 46, X46sc
    XNFadeL 46, X46sc1

    XNFadeLm 47, X47
    XNFadeLm 47, X47a
    XNFadeLm 47, X47b
	XNFadeLm 47, X47c
    XNFadeL 47, X47d

    SetLamp 147,LampState(47)
    Flashm 147, f47
    Flash 147, f47a

    SetLamp 146,LampState(46)
    Flashm 146, f46
    Flashm 146, f46a
    Flash 146, f46b

    SetLamp 145,LampState(45)
    Flash 145, f45

    Flashm 164, RTF
    Flashm 164, RTF1
    Flash 164, RTFDim

    Flash 161, LF61

    Flashm 162, LF62
    Flash 162, LTFDim
    Flash 163, LF63

    	' Flashers
	'FadeL 161, l61, l61z ' Top Left Flasher 1
	'FadeL 162, l62, l62z ' Top Left Flasher 2
	'FadeL 163, l63, l63z ' Top Left Flasher 3
	'FadeL 164, l64, l64z ' Top Right Flasher
	'FlashAR 61, l61a, l61b, l61c, FRefresh1, 200, 200 ' Top Left Flasher 1
	'FlashAR 62, l62a, l62b, l62c, FRefresh1, 200, 200 ' Top Left Flasher 2
	'FlashAR 63, l63a, l63b, l63c, FRefresh1, 200, 200 ' Top Left Flasher 3
	'FlashAR 64, l64a, l64b, l64c, FRefresh2, 200, 200 ' Top Right Flasher
	'FlashAR 65, l65a, l65b, l65c, FRefresh3, 200, 200 ' Right Plastic Flasher
	'FlashAR 66, l66a, l66b, l66c, FRefresh4, 200, 200  ' Left Plastic Flasher
	' FlashAR 67 ' Stargate Flasher 1 (Didn't Use)
	' FlashAR 68 ' Stargate Flasher 2 (Didn't Use)

	' Auxiliary Lights
	XNFadeL 71, l71
	XNFadeL 72, l72
	XNFadeL 73, l73
	XNFadeL 74, l74
	XNFadeL 75, l75
	XNFadeL 76, l76
	XNFadeL 77, l77
	XNFadeL 78, l78
	XNFadeL 79, l79
	XNFadeL 80, l80
	XNFadeL 81, l81
	XNFadeL 82, l82
	XNFadeL 83, l83
	XNFadeL 84, l84
	XNFadeL 85, l85
	XNFadeL 86, l86
	XNFadeL 87, l87
	XNFadeL 88, l88
	XNFadeL 89, l89
	XNFadeL 90, l90

XNFadeLm 91, l91a ' Rear Chaser Light
XNFadeLm 91, l91b ' Rear Chaser Light
XNFadeL 91, l91c ' Rear Chaser Light

SetLamp 191,LampState(91)
Flash 191, F91

XNFadeLm 100, l100a ' Rear Chaser Light
XNFadeLm 100, l100b ' Rear Chaser Light
XNFadeL 100, l100c ' Rear Chaser Light

SetLamp 200,LampState(100)
Flash 200, F100

XNFadeLm 92, l92a ' Rear Chaser Light
XNFadeLm 92, l92b ' Rear Chaser Light
XNFadeL 92, l92c ' Rear Chaser Light

SetLamp 192,LampState(92)
Flash 192, F92

XNFadeLm 99, l99a ' Rear Chaser Light
XNFadeLm 99, l99b ' Rear Chaser Light
XNFadeL 99, l99c ' Rear Chaser Light

SetLamp 199,LampState(99)
Flash 199, F99

XNFadeLm 93, l93a ' Rear Chaser Light
XNFadeLm 93, l93b ' Rear Chaser Light
XNFadeL 93, l93c ' Rear Chaser Light

SetLamp 193,LampState(93)
Flash 193, F93

XNFadeLm 98, l98a ' Rear Chaser Light
XNFadeLm 98, l98b ' Rear Chaser Light
XNFadeL 98, l98c ' Rear Chaser Light

SetLamp 198,LampState(98)
Flash 198, F98

XNFadeLm 94, l94a ' Rear Chaser Light
XNFadeLm 94, l94b ' Rear Chaser Light
XNFadeL 94, l94c ' Rear Chaser Light

SetLamp 194,LampState(94)
Flash 194, F94

XNFadeLm 97, l97a ' Rear Chaser Light
XNFadeLm 97, l97b ' Rear Chaser Light
XNFadeL 97, l97c ' Rear Chaser Light

SetLamp 197,LampState(97)
Flash 197, F97

XNFadeLm 95, l95a ' Rear Chaser Light
XNFadeLm 95, l95b ' Rear Chaser Light
XNFadeL 95, l95c ' Rear Chaser Light

SetLamp 195, LampState(95)
Flash 195, F95

XNFadeLm 96, l96a ' Rear Chaser Light
XNFadeLm 96, l96b ' Rear Chaser Light
XNFadeL 96, l96c ' Rear Chaser Light

SetLamp 196,LampState(96)
Flash 196, F96

Dim gixx

If LampState(1) = 1 Then
For each gixx in GI:gixx.state = 0:next
'DOF 150,0
else
For each gixx in GI:gixx.state = 1:next
'DOF 150,1
End If

If LampState(1) = 1 AND gion = 1 Then
PlaySound "fx_relay"
DOF 103,0
gion = 0
End If

If LampState(1) = 0 AND gion = 0 Then
PlaySound "fx_relay"
DOF 103,1
gion = 1
End If

If LampState(1) = 1 Then
Table1.ColorGradeImage = "ColorGrade_off"
else
Table1.ColorGradeImage = "ColorGrade_on"
End If

if Alpha1.isdropped AND LampState(1) < 1 Then
FL1.opacity = 0
else
FL1.opacity = 55
End If

if Alpha2A.isdropped AND LampState(1) < 1 Then
FL2.opacity = 0
else
FL2.opacity = 55
End If

if Alpha2B.isdropped AND LampState(1) < 1 Then
FL3.opacity = 0
else
FL3.opacity = 55
End If

if Alpha3A.isdropped AND LampState(1) < 1 Then
FL4.opacity = 0
else
FL4.opacity = 55
End If

if Alpha3B.isdropped AND LampState(1) < 1 Then
FL5.opacity = 0
else
FL5.opacity = 55
End If

if Alpha3C.isdropped AND LampState(1) < 1 Then
FL6.opacity = 0
else
FL6.opacity = 55
End If

if betab.isdropped AND LampState(1) < 1 Then
FLb.opacity = 0
else
FLb.opacity = 55
End If

if betae.isdropped AND LampState(1) < 1 Then
FLe.opacity = 0
else
FLe.opacity = 55
End If

if betat.isdropped AND LampState(1) < 1 Then
FLt.opacity = 0
else
FLt.opacity = 55
End If

if betaa.isdropped AND LampState(1) < 1 Then
FLa.opacity = 0
else
FLa.opacity = 55
End If

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

'Walls

Sub FadeW(nr, a, b, c)
    Select Case LampState(nr)
        Case 2:c.IsDropped = 1:LampState(nr) = 0                 'Off
        Case 3:b.IsDropped = 1:c.IsDropped = 0:LampState(nr) = 2 'fading...
        Case 4:a.IsDropped = 1:b.IsDropped = 0:LampState(nr) = 3 'fading...
        Case 5:c.IsDropped = 1:b.IsDropped = 0:LampState(nr) = 6 'ON
        Case 6:b.IsDropped = 1:a.IsDropped = 0:LampState(nr) = 1 'ON
    End Select
End Sub

Sub FadeWm(nr, a, b, c)
    Select Case LampState(nr)
        Case 2:c.IsDropped = 1
        Case 3:b.IsDropped = 1:c.IsDropped = 0
        Case 4:a.IsDropped = 1:b.IsDropped = 0
        Case 5:c.IsDropped = 1:b.IsDropped = 0
        Case 6:b.IsDropped = 1:a.IsDropped = 0
    End Select
End Sub

Sub NFadeW(nr, a)
    Select Case LampState(nr)
        Case 4:a.IsDropped = 1:LampState(nr) = 0
        Case 5:a.IsDropped = 0:LampState(nr) = 1
    End Select
End Sub

Sub NFadeWm(nr, a)
    Select Case LampState(nr)
        Case 4:a.IsDropped = 1
        Case 5:a.IsDropped = 0
    End Select
End Sub

Sub NFadeWi(nr, a)
    Select Case LampState(nr)
        Case 4:a.IsDropped = 0:LampState(nr) = 0
        Case 5:a.IsDropped = 1:LampState(nr) = 1
    End Select
End Sub

Sub NFadeWim(nr, a)
    Select Case LampState(nr)
        Case 4:a.IsDropped = 0
        Case 5:a.IsDropped = 1
    End Select
End Sub

'Lights

Sub XNFadeL(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.state = 0:FadingLevel(nr) = 0
        Case 5:object.state = 1:FadingLevel(nr) = 1
    End Select
End Sub

Sub XNFadeLm(nr, object) ' used for multiple lights
    Select Case FadingLevel(nr)
        Case 4:object.state = 0
        Case 5:object.state = 1
    End Select
End Sub

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
        '   Object.IntensityScale = 1
End Sub

Sub Flashm(nr, object) 'multiple flashers, it just sets the flashlevel
    Object.IntensityScale = FlashLevel(nr)
         ' Object.IntensityScale = 1
End Sub

Sub FadeL(nr, a, b)
    Select Case LampState(nr)
        Case 2:b.state = 0:LampState(nr) = 0
        Case 3:b.state = 1:LampState(nr) = 2
        Case 4:a.state = 0:LampState(nr) = 3
        Case 5:b.state = 1:LampState(nr) = 6
        Case 6:a.state = 1:LampState(nr) = 1
    End Select
End Sub

Sub FastL(nr, a, b)
    Select Case LampState(nr)
        Case 2:b.state = 0:LampState(nr) = 0
        Case 3:b.state = 1:LampState(nr) = 2
        Case 4:a.state = 0:LampState(nr) = 3
        Case 5:a.state = 1:LampState(nr) = 1
    End Select
End Sub

Sub FadeLm(nr, a, b)
    Select Case LampState(nr)
        Case 2:b.state = 0
        Case 3:b.state = 1
        Case 4:a.state = 0
        Case 5:b.state = 1
        Case 6:a.state = 1
    End Select
End Sub

Sub NFadeL(nr, a)
    Select Case LampState(nr)
        Case 4:a.state = 0:LampState(nr) = 0
        Case 5:a.State = 1:LampState(nr) = 1
    End Select
End Sub

Sub NFadeLm(nr, a)
    Select Case LampState(nr)
        Case 4:a.state = 0
        Case 5:a.State = 1
    End Select
End Sub

Sub FadeOldL(nr, a, b, c)
    Select Case LampState(nr)
        Case 2:c.state = 0:LampState(nr) = 0
        Case 3:b.state = 0:c.state = 1:LampState(nr) = 2
        Case 4:a.state = 0:b.state = 1:LampState(nr) = 3
        Case 5:a.state = 0:c.state = 0:b.state = 1:LampState(nr) = 6
        Case 6:b.state = 0:a.state = 1:LampState(nr) = 1
    End Select
End Sub

Sub FadeOldLm(nr, a, b, c)
    Select Case LampState(nr)
        Case 2:c.state = 0
        Case 3:b.state = 0:c.state = 1
        Case 4:a.state = 0:b.state = 1
        Case 5:a.state = 0:c.state = 0:b.state = 1
        Case 6:b.state = 0:a.state = 1
    End Select
End Sub

'Reels

Sub FadeR(nr, a)
    Select Case LampState(nr)
        Case 2:a.SetValue 3:LampState(nr) = 0
        Case 3:a.SetValue 2:LampState(nr) = 2
        Case 4:a.SetValue 1:LampState(nr) = 3
        Case 5:a.SetValue 1:LampState(nr) = 6
        Case 6:a.SetValue 0:LampState(nr) = 1
    End Select
End Sub

Sub FadeRm(nr, a)
    Select Case LampState(nr)
        Case 2:a.SetValue 3
        Case 3:a.SetValue 2
        Case 4:a.SetValue 1
        Case 5:a.SetValue 1
        Case 6:a.SetValue 0
    End Select
End Sub

'Texts

Sub NFadeT(nr, a, b)
    Select Case LampState(nr)
        Case 4:a.Text = "":LampState(nr) = 0
        Case 5:a.Text = b:LampState(nr) = 1
    End Select
End Sub

Sub NFadeTm(nr, a, b)
    Select Case LampState(nr)
        Case 4:a.Text = ""
        Case 5:a.Text = b
    End Select
End Sub

' Flash a light, not controlled by the rom

Sub FlashL(nr, a, b)
    Select Case LampState(nr)
        Case 1:b.state = 0:LampState(nr) = 0
        Case 2:b.state = 1:LampState(nr) = 1
        Case 3:a.state = 0:LampState(nr) = 2
        Case 4:a.state = 1:LampState(nr) = 3
        Case 5:b.state = 1:LampState(nr) = 4
    End Select
End Sub

' Light acting as a flash. C is the light number to be restored

Sub MFadeL(nr, a, b, c)
    Select Case LampState(nr)
        Case 2:b.state = 0:LampState(nr) = 0
            If LampState(c) = 1 Then SetLamp c, 1
        Case 3:b.state = 1:LampState(nr) = 2
        Case 4:a.state = 0:LampState(nr) = 3
        Case 5:a.state = 1:LampState(nr) = 1
    End Select
End Sub

' Added in version 5 : lights made with alpha ramps
' a, b, c and d are the ramps from on to off
' r is the refresh light
' wt is the top width of the ramp
' wb is the bottom width of the ramp

Sub FadeAR(nr, a, b, c, d, r, wt, wb)
    Select Case LampState(nr)
        Case 2:c.WidthBottom = 0:c.WidthTop = 0
            d.WidthBottom = wb:d.WidthTop = wt
            r.State = ABS(r.state -1)
            LampState(nr) = 0 'Off
        Case 3:b.WidthBottom = 0:b.WidthTop = 0
            c.WidthBottom = wb:c.WidthTop = wt
            r.State = ABS(r.state -1)
            LampState(nr) = 2 'fading...
        Case 4:a.WidthBottom = 0:a.WidthTop = 0
            b.WidthBottom = wb:b.WidthTop = wt
            r.State = ABS(r.state -1)
            LampState(nr) = 3 'fading...
        Case 5:d.WidthBottom = 0:d.WidthTop = 0
            b.WidthBottom = wb:b.WidthTop = wt
            r.State = ABS(r.state -1)
            LampState(nr) = 6 ' 1/2 ON
        Case 6:b.WidthBottom = 0:b.WidthTop = 0
            a.WidthBottom = wb:a.WidthTop = wt
            r.State = ABS(r.state -1)
            LampState(nr) = 1 'ON
    End Select
End Sub

Sub FadeARm(nr, a, b, c, d, r, wt, wb)
    Select Case LampState(nr)
        Case 2:c.WidthBottom = 0:c.WidthTop = 0
            d.WidthBottom = wb:d.WidthTop = wt
            r.State = ABS(r.state -1)
        Case 3:b.WidthBottom = 0:b.WidthTop = 0
            c.WidthBottom = wb:c.WidthTop = wt
            r.State = ABS(r.state -1)
        Case 4:a.WidthBottom = 0:a.WidthTop = 0
            b.WidthBottom = wb:b.WidthTop = wt
            r.State = ABS(r.state -1)
        Case 5:d.WidthBottom = 0:d.WidthTop = 0
            b.WidthBottom = wb:b.WidthTop = wt
            r.State = ABS(r.state -1)
        Case 6:b.WidthBottom = 0:b.WidthTop = 0
            a.WidthBottom = wb:a.WidthTop = wt
            r.State = ABS(r.state -1)
    End Select
End Sub

Sub FlashAR(nr, a, b, c, r, wt, wb) 'used for reflections when the off is transparent - no ramp
    Select Case LampState(nr)
        Case 2:c.WidthBottom = 0:c.WidthTop = 0
            r.State = ABS(r.state -1)
            LampState(nr) = 0 'Off
        Case 3:b.WidthBottom = 0:b.WidthTop = 0
            c.WidthBottom = wb:c.WidthTop = wt
            r.State = ABS(r.state -1)
            LampState(nr) = 2 'fading...
        Case 4:a.WidthBottom = 0:a.WidthTop = 0
            b.WidthBottom = wb:b.WidthTop = wt
            r.State = ABS(r.state -1)
            LampState(nr) = 3 'fading...
        Case 5:b.WidthBottom = wb:b.WidthTop = wt
            r.State = ABS(r.state -1)
            LampState(nr) = 6 ' 1/2 ON
        Case 6:b.WidthBottom = 0:b.WidthTop = 0
            a.WidthBottom = wb:a.WidthTop = wt
            r.State = ABS(r.state -1)
            LampState(nr) = 1 'ON
    End Select
End Sub

Sub FlashARm(nr, a, b, c, r, wt, wb) 'used for reflections when the off is transparent - no ramp
    Select Case LampState(nr)
        Case 2:c.WidthBottom = 0:c.WidthTop = 0
            r.State = ABS(r.state -1)
        Case 3:b.WidthBottom = 0:b.WidthTop = 0
            c.WidthBottom = wb:c.WidthTop = wt
            r.State = ABS(r.state -1)
        Case 4:a.WidthBottom = 0:a.WidthTop = 0
            b.WidthBottom = wb:b.WidthTop = wt
            r.State = ABS(r.state -1)
        Case 5:b.WidthBottom = wb:b.WidthTop = wt
            r.State = ABS(r.state -1)
        Case 6:b.WidthBottom = 0:b.WidthTop = 0
            a.WidthBottom = wb:a.WidthTop = wt
            r.State = ABS(r.state -1)
    End Select
End Sub

Sub NFadeAR(nr, a, b, r, wt, wb) ' a is the ramp on, b if the ramp off
    Select Case LampState(nr)
        Case 4:a.WidthBottom = 0:a.WidthTop = 0
            b.WidthBottom = wb:b.WidthTop = wt
            r.State = ABS(r.state -1)
            LampState(nr) = 0 'off
        Case 5:a.WidthBottom = wb:a.WidthTop = wt
            b.WidthBottom = 0:b.WidthTop = 0
            r.State = ABS(r.state -1)
            LampState(nr) = 1 'on
    End Select
End Sub

Sub NFadeARm(nr, a, b, r, wt, wb) ' a is the ramp on, b if the ramp off
    Select Case LampState(nr)
        Case 4:a.WidthBottom = 0:a.WidthTop = 0
            b.WidthBottom = wb:b.WidthTop = wt
            r.State = ABS(r.state -1)
        Case 5:a.WidthBottom = wb:a.WidthTop = wt
            b.WidthBottom = 0:b.WidthTop = 0
            r.State = ABS(r.state -1)
    End Select
End Sub

'--------- ADDED by JF
Sub BallRelease_UnHit(): DOF 112,2:End Sub
Sub Untrapper2_UnHit(): Untrapper2.Enabled = 0 : End Sub
'--------- END ADDED by JF

' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Sub Pins_Hit (idx)
	PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
	PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub TargetBankWalls_Hit (idx)
	PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
	PlaySound "fx_metalhit", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
	PlaySound "metalhit_medium", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
	PlaySound "fx_metalclank", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
	PlaySound "gate", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Rolls_Hit (idx)
	PlaySound "fx_sensor", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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

Sub Shad_Rot_Timer()
XBallShadow1.ObjRotZ = XBallShadow1.ObjRotZ + 1
XBallShadow2.ObjRotZ = XBallShadow2.ObjRotZ + 1
End Sub

Sub Table1_Exit
Controller.Stop
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

Const tnob = 8 ' total number of balls
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
        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 AND BOT(b).z > 0 Then
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


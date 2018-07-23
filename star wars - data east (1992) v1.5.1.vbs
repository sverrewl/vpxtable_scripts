
'Star Wars by Data East (1992) [VPX]

'Authors.

'Dids666 & DJRobX 


'----------------------------------------

'(Contributors)

'Neo: Plastic Scans.

'Nfozzy: Lamp Fading Script

'Flupper: Flipper Primitives

'Dark: Apron Primitive

'Ninuzzu: Ball/Flipper Shadow Script.

'Draifet (See Below)

'Alistaircg (See Below)

'----------------------------------------

'Changelog.

'New 4K Playfield (Scans by Alistaircg)

'SSF Updated

'Missing switches added

'R2D2 drophole Added

'Physics update (by Draifet)

'DT View DMD surround (by Draifet)

'Lots of other little bits ;)

'----------------------------------------


'(Release Media)

'Hauntfreaks: New B2S Backglass

'Alistaircg/Wildman: B2s Backglass

'Bambi Plattfuss: Topper Video


Option Explicit
Randomize

Const PostIt = 0
Const BB8Ball = 1
Const DeathStarBolt = 0

Dim Ballsize,BallMass
Ballsize = 50
Ballmass = 1.3

Dim DesktopMode:DesktopMode = Table1.ShowDT
Dim UseVPMDMD:UseVPMDMD = DesktopMode

Const FlippersAlwaysOn = 0 'Enable Flippers for testing
' ===============================================================================================
' Load game controller
' ===============================================================================================
 
On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0
 
LoadVPM "01550000", "DE.VBS", 3.26
 
' ===============================================================================================
' General constants and variables
' ===============================================================================================
Const UseSolenoids = 2
Const UseLamps = 0
Const UseGI = 1
Const UseSync = 1 'set it to 1 if the table runs too fast
Const HandleMech = 1
 
 
' Standard Sounds
Const SSolenoidOn = "fx_Solenoid"
Const SSolenoidOff = ""
Const SFlipperOn    = "fx_Flipperup"
Const SFlipperOff   = "fx_Flipperdown"
Const SCoin = "fx_Coin"
 
'Dim bsTrough, dtbank1, dtbank2, bsLeftSaucer, bsRightSaucer, x
'Const cGameName = "stwr_a14" ' Patched
Const cGameName = "stwr_104" ' Patched




'Dim bsLowerEject,bsUpperEject
Dim gameRun
Dim bsTrough,mechDeathStar,mechDeathStarDoor,mechR2D2,R2UpSol,plungerIM, bsLSaucer, bsRSaucer, dtBank
Dim R2Height:R2Height = 0
Dim DSBalls:DSBalls = 0
Dim LastMotorMove:LastMotorMove = 0


' Lights
 
 vpmMapLights AllLights

Dim NullFader : set NullFader = new NullFadingObject
Dim FadeLights : Set FadeLights = New LampFader


InitLamps
LampTimer.Interval = 1
LampTimer.Enabled = 1
 
 
 
' ===============================================================================================
' Init routines
' ===============================================================================================
 
Sub table1_Init
    vpmInit me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Star Wars - Data East" & vbNewLine & "VPX table by Dids666 and DJRobX"
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
		.Hidden = DesktopMode
        .Games(cGameName).Settings.Value("rol") = 0 '1= rotated display, 0= normal
        '.SetDisplayPosition 0,0, GetPlayerHWnd 'restore dmd window position
        On Error Resume Next
       ' Controller.SolMask(0) = 1
        'vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the Timer to renable all the solenoids after 2 seconds
        Controller.Run GetPlayerHWnd
        On Error Goto 0
    End With

	if DesktopMode Then	
		CenterRamp.image = "Ramp Desktop View"
	end if
		
    if CBool(DeathStarBolt) Then
		DeathStarBoltPrim.visible = True
	end if 
	if not CBool(PostIt) Then
		PostItPrim.Visible = False
	end if 
	if BB8Ball Then
		Table1.BallFrontDecal = "bb8-texture6"
		Table1.BallDecalMode = True
		Table1.BallImage = "Ball_Final"
		Table1.DefaultBulbIntensityScale = 0.1
	end if 
   	Set bsTrough=New cvpmTrough
 	with bsTrough
		.Size=4 
		.InitSwitches Array(11,13,12,10)
		.InitExit BallRelease,90,7
		.InitEntrySounds "fx_drain", SoundFX("fx_solenoid",DOFContactors), SoundFX("fx_solenoid",DOFContactors)
		.InitExitSounds  SoundFX("fx_solenoid",DOFContactors), SoundFX("fx_ballrel",DOFContactors)
		.Balls=3
		.CreateEvents "bsTrough", Drain
 	end with

	'Auto Plunger
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP sw14, 65, 0.5
        .Switch 14
        .Random 0.05
        .CreateEvents "plungerIM"
    End With

	' Vuk
'	Set bsVUK = New cvpmSaucer
'	With bsVUK
'		.InitKicker sw36, 36, 0, 45, 1.56
'	    .InitExitVariance 1, 1
'		.InitSounds "fx_kickerenter", SoundFX(SSolenoidOn,DOFContactors), SoundFX("fx_solenoid",DOFContactors)
'		.CreateEvents "bsVUK", sw36
'	End With

	'Left Scoop
    Set bsLSaucer = new cvpmSaucer
    With bsLSaucer
        .InitKicker LScoopEject, 37, 0, 45, 1.56
		.InitExitVariance 0, 2
      '  .InitSounds "fx_kickerenter", SoundFX(SSolenoidOn,DOFContactors), SoundFX("fx_solenoid",DOFContactors)
        .CreateEvents "bsLSaucer", LScoopEject
    End With

'Left Scoop
    Set bsRSaucer = new cvpmSaucer
    With bsRSaucer
        .InitKicker RScoopEject, 33, 0, 45, 1.56
		.InitExitVariance 0, 2
        .InitSounds "fx_kickerenter", SoundFX(SSolenoidOn,DOFContactors), SoundFX("fx_solenoid",DOFContactors)
        .CreateEvents "bsRSaucer", RScoopEject
    End With

	set dtBank = new cvpmdroptarget
     With dtBank
         .initdrop array(sw32, sw31, sw30), array(32, 31, 30)
     End With


     ' Nudging
     vpmNudge.TiltSwitch = 1
     vpmNudge.Sensitivity = 1
     vpmNudge.TiltObj = Array(sw25, sw26, sw27, sw39, sw40)

	' Death Star
     Set mechDeathStarDoor = new cvpmMyMech
     With mechDeathStarDoor
         .Mtype = vpmMechOneSol + vpmMechReverse
         .Sol1 = 12
         .Length = 220
         .Steps = 72
         .Acc = 30
         .Ret = 0
         .AddSw 41, 0, 0
         .AddSw 42, 71, 71
         .Callback = GetRef("UpdateDeathStarDoor")
         .Start
     End With

	Set mechDeathStar = new cvpmMyMech
	With mechDeathStar
		.Sol1 = 15
		.MType = vpmMechLinear + vpmMechCircle + vpmMechOneSol
		.Acc = 60 : .Ret = 2
		.Length = 220
		.Steps = 360
		.Callback = GetRef("UpdateDeathStar")
		.Start
	End With

	GILights 1
	Kickback.PullBack
	InitBallShadow
	

	'SetLamp 130,1
End Sub

Sub r2In_Hit()
	r2In.DestroyBall
	vpmTimer.PulseSw 35
	AddVUK
	'vpmTimer.PulseSwitch 35, 100, "bsVUK.AddBall 0'"
End Sub

Sub dsIn_Hit()
	dsIn.DestroyBall
	vpmTimer.PulseSw 34
	AddVUK
	'vpmTimer.PulseSwitch 34, 100, "bsVUK.AddBall 0'"
End Sub

Sub sw36hole_Hit()
	PlaySoundAt "fx_hole1",ActiveBall
	sw36hole.DestroyBall
	AddVUK
End Sub


Sub AddVUK
	vpmTimer.AddTimer 200, "dsBalls = dsBalls + 1:Controller.Switch(36) = 1'"
End Sub 

Sub SolLeftPopper(Enabled)
	if Enabled and DSBalls > 0 then 
		sw36hole.Enabled = 0 
		sw36.CreateSizedBallWithMass Ballsize/2, BallMass
		sw36.Kick 0, 45, 1.56
		PlaySoundAt "fx_popper", sw36
		DSBalls = DSBalls - 1
		if DSBalls = 0 Then
			Controller.Switch(36) = 0
		end if 
	End If
End Sub


Sub UpdateDeathStar(pos, speed, prevpos)
	DeathStar.RotY = pos
    if CBool(DeathStarBolt) Then
		DeathStarBoltPrim.RotY = pos
	end if 
	UpdateR2Head
	if speed > .1 Then
		if LastMotorMove = 0 then
			PlaySound"motor", -1, .08, 0, 0, 0, 1, 0, AudioFade(dsIn)
			LastMotorMove = timer + 1
		end if
	Else	
		LastMotorMove = 0
		StopSound "motor"
	end if 
End Sub

dim LastDoorMove:LastDoorMove =0

Sub UpdateDeathStarDoor(pos, speed, prevpos)
	'debug.print pos
	DSDoor.Z = -pos + 12
	if DSDoor.Z <= -52 then 
		sw40.Collidable = 0
	Else	
		sw40.Collidable = 1
	end if 
	if speed > .1 Then
		if LastDoorMove = 0 Then	
			LastDoorMove = Timer + 1
			PlaySound"fx_motor", -1, .08, Pan(dsIn), 0, 0, 1, 0, AudioFade(dsIn)
		end if
	Else	
		LastDoorMove = 0
		StopSound "fx_motor"
	end if 
End Sub


Dim R2HeadDir:R2HeadDir =1
Const R2HeadMin = -70
Const R2HeadMax = 70

Sub UpdateR2Head
	dim R2HeadSpeed:R2HeadSpeed = (R2HeadMax - R2Head.RotY)/20
	dim R2HeadSpeedB:R2HeadSpeedB = (R2Head.RotY-R2HeadMin+1)/10
	if R2HeadSpeedB < R2HeadSpeed then R2HeadSpeed = R2HeadSpeedB
	If R2HeadSpeed > 1 Then R2HeadSpeed = 1

	R2Head.RotY = R2Head.RotY + (R2HeadDir*R2HeadSpeed)
	If R2Head.RotY> (R2HeadMax-1) Then R2HeadDir = -1
	If R2Head.RotY < R2HeadMin Then R2HeadDir = 1
	R2Head2.RotY = R2Head.RotY
End Sub

Const R2Top = 15
Const R2SpeedDown = 1.8
Const R2SpeedUp = 2.5

Sub UpdateR2Light(R2Light)
	R2Head2.BlendDisableLighting = R2Light
	MaterialColor "Plastic Blue Transp", RGB(R2Light * 25, R2Light * 30 + 50, R2Light * 50 + 204)
End Sub 

Sub UpdateR2
	if R2UpSol<>0 and R2Height < R2Top then 
		R2Height = R2Height + R2SpeedUp
		if R2Height > R2Top then R2Height = R2Top
	end if 
	if R2UpSol=0 and R2Height > 0 then 
		R2Height = R2Height - R2SpeedDown
		if R2Height < 0 then R2Height = 0
	end if
	R2Head.Z = (R2Top-R2Height)+50
	R2Head2.Z = R2Head.Z
	R2Body.Z = R2Head.Z
End Sub


Sub Table1_KeyDown(ByVal keycode)
	If keycode = LeftTiltKey Then Nudge 90, 2
	If keycode = RightTiltKey Then Nudge 270, 2
	If keycode = CenterTiltKey Then	Nudge 0, 2
	if keycode = LeftMagnaSave then Controller.Switch(51) = 0
	if keycode = RightMagnaSave then Controller.Switch(51) = 1'R2UpSol=1


	If vpmKeyDown(keycode) Then Exit Sub
	If keycode = PlungerKey Then 
		Controller.Switch(50) = 1
		if not CBool(PostIt) Then Controller.Switch(51)=1 
	end if 
'	If FlippersAlwaysOn =1 Then
'		If keycode = LeftFlipperKey Then LeftFlipper.RotateToEnd: PlaySound SoundFX("fx_flipperup",DOFContactors), 0, .67, -0.05, 0.05
'		If keycode = RightFlipperKey Then RightFlipper.RotateToEnd: PlaySound SoundFX("fx_flipperup",DOFContactors), 0, .67, 0.05, 0.05
'	End If
End Sub

Sub Table1_KeyUp(ByVal keycode)
	If keycode = PlungerKey Then 
		Controller.Switch(50) = 0
		if not CBool(PostIt) Then Controller.Switch(51)=0
	end if 	'if keycode = RightMagnaSave then Controller.Switch(51) = 1'R2UpSol=0
	If FlippersAlwaysOn =1 Then
		If keycode = LeftFlipperKey Then LeftFlipper.RotateToStart: PlaySound SoundFX("fx_flipperdown",DOFContactors), 0, 1, -0.05, 0.05
		If keycode = RightFlipperKey Then RightFlipper.RotateToStart: PlaySound SoundFX("fx_flipperdown",DOFContactors), 0, 1, 0.05, 0.05
	End If

	If vpmKeyUp(keycode) Then Exit Sub

End Sub

Sub GILights (enabled)
	Dim light
	For each light in GI:light.State = Enabled: Next
	For each light in SLINGLIGHTS:light.State = Enabled: Next
	For each light in FLIPPERLIGHTS:light.State = Enabled: Next
	For each light in UPPERPLAYFIELDLIGHTS:light.State = Enabled: Next
End Sub




'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
	vpmTimer.PulseSw 44
    PlaySound SoundFX("fx_slingshot",DOFContactors), 0, 1, 0.05, 0.05
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
	vpmTimer.PulseSw 43
    PlaySound SoundFX("fx_slingshot",DOFContactors),0,1,-0.05,0.05
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


Sub SolLeftEject(enabled)
	If Enabled Then
		bsLSaucer.ExitSol_On
		sw37.Enabled=0
		PlaySoundAt SoundFX("fx_kicker",DOFContactors), sw37
		vpmTimer.AddTimer 200, "sw37.Enabled=1 '"
	End If
End Sub

Sub SolRightEject(enabled)
	If Enabled Then
		bsRSaucer.ExitSol_On
		sw33.Enabled=0
		PlaySoundAt SoundFX("fx_kicker",DOFContactors), sw33
		vpmTimer.AddTimer 200, "sw33.Enabled=1 '"
	End If
End Sub

'************************************************
' Solenoids
'************************************************
SolCallback(1) = 	"bsTrough.SolIn"
SolCallback(2) = 	"bsTrough.SolOut"
SolCallback(3) = 	"SolAutoPlungerIM"
SolCallback(4) = 	"SolLeftEject"
SolCallback(5) = 	"SolLeftPopper" 
SolCallback(6) =    "PlaySoundAt SoundFX(""fx_droptargetreset"",DOFDropTargets),sw32:dtBank.SolDropUp"
SolCallback(7) = 	"SolRightEject"
SolCallback(9) = 	"SolR2"
SolCallback(11) = 	"GILights Not "
SolCallback(16) = 	"SolKickback" 
SolCallback(17) = 	"SetLamp 117,"
SolCallback(18) = 	"SetLamp 118,"
SolCallback(19) = 	"SetLamp 119,"
SolCallback(20) = 	"SetLamp 120,"

SolCallback(25) = "SetLamp 125,"
SolCallBack(26) = "SetLamp 126,"
SolCallBack(27) = "SetLamp 127,"
SolCallBack(28) = "SetLamp 128,"
SolCallBack(29) = "SetLamp 129,"
SolCallBack(30) = "SetLamp 130,"
SolCallBack(31) = "SetLamp 131,"
SolCallback(32) = "SetLamp 132,"


SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

sub SolKickback(enabled)
	PlaySoundAt SoundFX("fx_autoplunger", DOFContactors), Kickback
	vpmSolAutoPlunger Kickback,10,enabled
	if enabled then
		KickbackAnim.RotateToEnd
	Else		
		KickbackAnim.RotateToStart
	end if 
end sub 

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySound SoundFX("fx_flipperup",DOFContactors), 0, 1, -0.1, 0.25
        LeftFlipper.RotateToEnd
    Else
        PlaySound SoundFX("fx_flipperdown",DOFContactors), 0, 1, -0.1, 0.25
        LeftFlipper.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySound SoundFX("fx_flipperup",DOFContactors), 0, 1, 0.1, 0.25
        RightFlipper.RotateToEnd
   Else
        PlaySound SoundFX("fx_flipperdown",DOFContactors), 0, 1, 0.1, 0.25
        RightFlipper.RotateToStart
    End If
End Sub

Sub SolKnocker(Enabled)
	If Enabled Then PlaySound SoundFX("Knocker",DOFKnocker)
End Sub

Sub SolAutoPlungerIM(Enabled)
	If Enabled Then
		PlaySoundAt SoundFX("fx_autoplunger",DOFContactors), sw14
		AutoplungerAnim.RotateToEnd
		PlungerIM.AutoFire
	Else	
		AutoplungerAnim.RotateToStart
	End If
End Sub

sub SolR2(enabled)
	R2UpSol=enabled
	if enabled then PlaySoundAt SoundFX("fx_r2", DOFContactors), sw38
end sub 
		

'************************************************
' Switches
'************************************************

Sub sw17_hit: vpmTimer.PulseSw 17 : PlaySoundAt "fx_sensor",ActiveBall:End Sub

Sub sw19_hit: vpmTimer.PulseSw 19 : PlaySoundAt "fx_gate",ActiveBall:End Sub

'Sub sw20_hit: Controller.Switch(20) = 1 : PlaySoundAt "fx_gate",ActiveBall:End Sub
'Sub sw20_unhit:Controller.Switch(20) = 0: End Sub

Sub sw20_hit: vpmTimer.PulseSw 20 : PlaySoundAt "fx_gate",ActiveBall:End Sub

Sub sw21_hit: Controller.Switch(21) = 1 : PlaySoundAt "fx_gate",ActiveBall:End Sub
Sub sw21_unhit:Controller.Switch(21) = 0: End Sub

Sub sw22_hit: Controller.Switch(22) = 1 : PlaySoundAt "fx_gate",ActiveBall:End Sub
Sub sw22_unhit:Controller.Switch(22) = 0: End Sub

Sub sw23_hit: Controller.Switch(23) = 1 : PlaySoundAt "fx_gate",ActiveBall:End Sub
Sub sw23_unhit:Controller.Switch(23) = 0: End Sub

Sub sw24_hit: Controller.Switch(24) = 1 : PlaySoundAt "fx_gate",ActiveBall:End Sub
Sub sw24_unhit:Controller.Switch(24) = 0: End Sub

Sub sw25_hit: vpmtimer.pulsesw 25 : PlaySoundAt "fx_target",ActiveBall:End Sub
Sub sw26_hit: vpmtimer.pulsesw 26 : PlaySoundAt "fx_target",ActiveBall:End Sub
Sub sw27_hit: vpmtimer.pulsesw 27 : PlaySoundAt "fx_target",ActiveBall:End Sub
Sub sw28_hit: vpmtimer.pulsesw 28 : PlaySoundAt "fx_target",ActiveBall:End Sub
Sub sw29_hit: vpmtimer.pulsesw 29 : PlaySoundAt "fx_target",ActiveBall:End Sub

Sub dsHole_hit: PlaySoundAt "fx_hole1",ActiveBall:End Sub
Sub r2Hole_hit: PlaySoundAt "fx_hole2",ActiveBall:End Sub
Sub sw37_hit:PlaySoundAt "fx_kicker_enter",sw37:End Sub
Sub sw33_hit:PlaySoundAt "fx_kicker_enter",sw33:End Sub

Sub sw30_Hit:dtBank.Hit 3: PlaySoundAt SoundFX("fx_droptarget",DOFDropTargets),ActiveBall: End Sub
Sub sw31_Hit:dtBank.Hit 2: PlaySoundAt SoundFX("fx_droptarget",DOFDropTargets),ActiveBall: End Sub
Sub sw32_Hit:dtBank.Hit 1: PlaySoundAt SoundFX("fx_droptarget",DOFDropTargets),ActiveBall: End Sub


Sub sw38_hit: Controller.Switch(38) = 1 : PlaySoundAt "fx_sensor",ActiveBall:End Sub
Sub sw38_unhit:Controller.Switch(38) = 0: End Sub

Sub sw39_hit: vpmTimer.AddTimer 1400, "vpmTimer.PulseSw 39'":PlaySoundAt "fx_sensor",ActiveBall:End Sub

Sub sw40_hit: vpmtimer.PulseSw 40 : PlaySoundAt "fx_target",ActiveBall:End Sub

' ******
' Bumpers
' ******


Sub sw45_hit:vpmtimer.pulsesw 45:PlaySoundAt SoundFX("fx_bumper1",DOFContactors), ActiveBall:End Sub
Sub sw46_hit:vpmtimer.pulsesw 46:PlaySoundAt SoundFX("fx_bumper2",DOFContactors), ActiveBall:End Sub
Sub sw47_hit:vpmtimer.pulsesw 47:PlaySoundAt SoundFX("fx_bumper3",DOFContactors), ActiveBall:End Sub
Sub sw48_hit:vpmtimer.pulsesw 48:PlaySoundAt SoundFX("fx_bumper3",DOFContactors), ActiveBall:End Sub


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

function AudioFade(ball)
	Dim tmp
    tmp = ball.y * 2 / Table1.height-1
    If tmp > 0 Then
		AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Function RndNum(min,max)
 RndNum = Int(Rnd()*(max-min+1))+min     ' Sets a random number between min and max
End Function

'Set position as table object (Use object or light but NOT wall) and Vol to 1
Sub PlaySoundAt(sound, tableobj)
	PlaySound sound, 1, 1, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'Set all as per ball position & speed.
Sub PlaySoundAtBall(sound)
	PlaySound sound, 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub

'Set position as table object and Vol manually.
Sub PlaySoundAtVol(sound, tableobj, Vol)
	PlaySound sound, 1, Vol, Pan(tableobj), 0, 0, 0, 1, AudioFade(tableobj)
End Sub

Sub PlayLoopSoundAtVol(sound, tableobj, Vol)
	PlaySound sound, -1, Vol, Pan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

'Set all as per ball position & speed, but Vol Multiplier may be used eg; PlaySoundAtBallVol "sound",3
Sub PlaySoundAtBallVol(sound, VolMult)
	PlaySound sound, 0, Vol(ActiveBall) * VolMult, Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub

'Set position as bumperX and Vol manually.
Sub PlaySoundAtBumperVol(sound, tableobj, Vol)
	PlaySound sound, 1, Vol, Pan(tableobj), 0,0,1, 1, AudioFade(tableobj)
End Sub

Dim NextOrbitHit:NextOrbitHit = 0
Sub PlasticRampBumps_Hit(idx)
	if BallVel(ActiveBall) > .3 and Timer > NextOrbitHit then
		RandomBump 10, -20000
		' Schedule the next possible sound time.  This prevents it from rapid-firing noises too much. 
		' Lowering these numbers allow more closely-spaced clunks.
		NextOrbitHit = Timer + .1 + (Rnd * .2)
	end if 
End Sub

Sub PlasticBumps_Hit(idx)
	PlaySoundAtBall "fx_plastichit"
End Sub

Sub MetalWallBumps_Hit(idx)
	if BallVel(ActiveBall) > .3 and Timer > NextOrbitHit then
		RandomBump 3, 20000 'Increased pitch to simulate metal wall
		' Schedule the next possible sound time.  This prevents it from rapid-firing noises too much. 
		' Lowering these numbers allow more closely-spaced clunks.
		NextOrbitHit = Timer + .2 + (Rnd * .2)
	end if 
End Sub

Sub WireRampBumps_Hit(idx)
	if BallVel(ActiveBall) > .3 and Timer > NextOrbitHit then
		dim BumpSnd:BumpSnd= "wirerampbump" & CStr(Int(Rnd*5)+1)
		PlaySound BumpSnd, 0, Vol(ActiveBall), Pan(ActiveBall), 0, 30000, 0, 1, AudioFade(ActiveBall)
		NextOrbitHit = Timer + .2 + (Rnd * .2)
	end if 
End Sub

Sub RandomBump(voladj, freq)
	dim BumpSnd:BumpSnd= "rampbump" & CStr(Int(Rnd*7)+1)
	PlaySound BumpSnd, 0, Vol(ActiveBall)*voladj, Pan(ActiveBall), 0, freq, 0, 1, AudioFade(ActiveBall)
End Sub

' Requires rampbump1 to 7 in Sound Manager
Sub RandomBump(voladj, freq)
	dim BumpSnd:BumpSnd= "rampbump" & CStr(Int(Rnd*7)+1)
	PlaySound BumpSnd, 0, Vol(ActiveBall)*voladj, Pan(ActiveBall), 0, freq, 0, 1, AudioFade(ActiveBall)
End Sub


Sub LeftRampEnd_Hit(): sw36hole.Enabled = 1 :BumpSTOPWire(): vpmTimer.AddTimer 100, "PlaySoundAt ""fx_balldrop"",LeftRampEnd'":End Sub
Sub RightRampEnd_Hit(): BumpSTOPWire(): vpmTimer.AddTimer 100, "PlaySoundAt ""fx_balldrop"",RightRampEnd'":End Sub

Sub REnter_Hit():If ActiveBall.VelY < 0 Then PlaySoundAtVol "fx_metalrampenter", REnter, 0.5:End If:End Sub			'
Sub REnter_UnHit():If ActiveBall.VelY > 0 Then StopSound "fx_metalramprenter":End If:End Sub		

' Stop Bump Sounds
Sub BumpSTOPMetal ()
dim i:for i=1 to 7:StopSound "RampBump" & i:next
NextOrbitHit = Timer + 1
End Sub

Sub BumpSTOPWire ()
dim i:for i=1 to 4:StopSound "WireRampBump" & i:next
NextOrbitHit = Timer + 1
End Sub
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

Sub RollingUpdate()
    Dim BOT, b, ballpitch
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
            If BOT(b).z < 30 Then
                ballpitch = Pitch(BOT(b) )
            Else
                ballpitch = Pitch(BOT(b) ) * 100
            End If
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, ballpitch, 1, 0, AudioFade(BOT(b))
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
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 1, AudioFade(ball1)
End Sub

'******************************
' Diverse Collection Hit Sounds
'******************************

' Ramp Soundss
Sub RHelp1_Hit()
    StopSound "fx_metalrolling"
    PlaySound "fx_ballrampdrop", 0, 1, pan(ActiveBall), 0, 0, 0, 1, AudioFade(ActiveBall)
End Sub

Sub RHelp2_Hit()
    StopSound "fx_metalrolling"
    PlaySound "fx_ballrampdrop", 0, 1, pan(ActiveBall), 0, 0, 0, 1, AudioFade(ActiveBall)
End Sub

Sub LeftFlipper_Collide(parm)
    RandomSoundFlipper
End Sub

Sub Rightflipper_Collide(parm)
    RandomSoundFlipper
End Sub

Sub RubberPosts_Hit(idx)
	RandomSoundRubber()
End Sub

Sub Rubbers_Hit(idx)
	RandomSoundRubber()
End Sub

Sub RandomSoundRubber()
	Select Case RndNum(1,3)
		Case 1 : PlaySoundAtBallVol "fx_rubber_hit_1",10
		Case 2 : PlaySoundAtBallVol "fx_rubber_hit_2",10
		Case 3 : PlaySoundAtBallVol "fx_rubber_hit_3",10
	End Select
End Sub

Sub RandomSoundFlipper()
	Select Case RndNum(1,3)
		Case 1 : PlaySoundAtBallVol "flip_hit_1", 20
		Case 2 : PlaySoundAtBallVol "flip_hit_2", 20
		Case 3 : PlaySoundAtBallVol "flip_hit_3", 20
	End Select
End Sub

'*************************************************************
' Lamp routines
' SetLamp xx, 0 is Off. SetLamp xx,1 is On
'*************************************************************

dim BumperCaps:BumperCaps = array(Bumper1Cap, Bumper2Cap, Bumper3Cap, Bumper4Cap)
dim BumperRings:BumperRings = array(Bumper1Ring, Bumper2Ring, Bumper3Ring, Bumper4Ring)
dim BumperScrewAs:BumperScrewAs = array(Bumper1CapScrew1, Bumper2CapScrew1, Bumper3CapScrew1, Bumper4CapScrew1)
dim BumperScrewBs:BumperScrewBs = array(Bumper1CapScrew2, Bumper2CapScrew2, Bumper3CapScrew2, Bumper4CapScrew2)

Sub Pop(Num, Lvl)
	dim offset:offset = cInt(Lvl * 10)
		
	BumperCaps(num).z = 80 - offset
	BumperRings(num).z = 10 - offset
	BumperScrewAs(num).z = 100 - offset
	BumperScrewBs(num).z = 100 - offset
end sub 

sub Jabba(Lvl)
	DomeLeft2.BlendDisableLighting = lvl / 2
	MaterialColor "JabbaDome", RGB(128 + Lvl * 128, 128 + Lvl * 128, 128)
	l48.IntensityScale = lvl
end sub

sub DropTargetLights(Lvl)
	lvl = lvl ^ 1.3 * .3
	sw30.BlendDisableLighting = lvl
	sw31.BlendDisableLighting = lvl
	sw32.BlendDisableLighting = lvl
end sub


Sub InitLamps()
	
	'Adjust fading speeds (1 / full MS fading time)
	dim x 
	for x = 0 to 140 : FadeLights.FadeSpeedUp(x) = 1/80 : FadeLights.FadeSpeedDown(x) = 1/100 : next
	for x = 117 to 120 : FadeLights.FadeSpeedUp(x) = 1/10 : FadeLights.FadeSpeedDown(x) = 1/10 : next
	FadeLights.FadeSpeedUp(6) = 1/45 : FadeLights.FadeSpeedDown(6) = 1/60
	FadeLights.FadeSpeedUp(48) = 1/25 : FadeLights.FadeSpeedDown(48) = 1/40
	FadeLights.FadeSpeedUp(57) = 1/25 : FadeLights.FadeSpeedDown(57) = 1/40
	
	FadeLights.obj(127) = array(f127c)	
	FadeLights.obj(128) = array(f128a, f128b, f128c, f128d, f128e, f128f, f128g, f128h, f128i, f128j, f128k, f128l)	
	FadeLights.obj(132) = array(f132a, f132b, F132c, F132d, F132e, F132f, F132g, F132h, F132i)
	FadeLights.obj(125) = array(f125c, f125d, f125e, f125f, f125g, f125h, f125i)
    FadeLights.obj(130) = array(f130c)
	FadeLights.obj(126) = array(f126a, f126b, f126c)
	FadeLights.Callback(6) = "DropTargetLights "
	FadeLights.Callback(48) = "Jabba "
	FadeLights.Callback(57) = "UpdateR2Light" 
	FadeLights.Callback(117) = "Pop 0,"
	FadeLights.Callback(118) = "Pop 1,"
	FadeLights.Callback(119) = "Pop 2,"
	FadeLights.Callback(120) = "Pop 3,"

	'Lamp Assignments
end Sub

Sub SetLamp(idx, state)
	Dim tmp
	' do same as internal VPM lighting routine... but here so we can also
	' deal with flashers.
	If IsArray(Lights(idx)) Then
		For Each tmp In Lights(idx) : tmp.State =state: Next
	Elseif not IsEmpty(Lights(idx)) then 
		Lights(idx).State = state
	End If
	FadeLights.state(idx) = state
End Sub


Sub LampTimer_Timer()
	FadeLights.Update1
End Sub

Sub FrameTimer_Timer()
    Dim chgLamp, num, chg, ii, idx
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp) Then
        For ii = 0 To UBound(chgLamp)
			idx = chgLamp(ii, 0)
			SetLamp idx, ChgLamp(ii,1)
        Next
    End If
 	UpdateR2	
	UpperGatePrim.RotX = UpperGate.CurrentAngle
	KickbackPrim.RotX = KickbackAnim.CurrentAngle
	AutoplungerPrim.RotX = AutoplungerAnim.CurrentAngle
	LeftFlipperPrim.RotY = LeftFlipper.CurrentAngle+150
	RightFlipperPrim.RotY = RightFlipper.CurrentAngle+150-90-30
	FlipperLSh.RotZ = LeftFlipper.currentangle
	FlipperRSh.RotZ = RightFlipper.currentangle	
	FadeLights.Update
	BallShadowUpdate
	RollingUpdate
End Sub
 

' *********************************************************************
'						BALL SHADOW
' *********************************************************************
ReDim BallShadow(tnob-1)
InitBallShadow

Sub InitBallShadow
	Dim i: For i=0 to tnob-1
		ExecuteGlobal "Set BallShadow(" & i & ")=BallShadow" & (i+1) & ":"
	Next
End Sub

Sub BallShadowUpdate()
    Dim BOT, b
    BOT = GetBalls
	Dim nBot:nBot = UBound(BOT)
	' hide shadow of deleted balls
	If nBot<(tnob-1) Then
		For b = (UBound(BOT) + 1) to (tnob-1)
			BallShadow(b).visible = 0
		Next
	End If
	' exit the Sub if no balls on the table
    If nBot = -1 Then Exit Sub
	if nBot >= tnob then nBot = tnob-1
	' render the shadow for each ball
    For b = 0 to nBot
		If BOT(b).X < Table1.Width/2 Then
			BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/10)) + 10
		Else
			BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/10)) - 10
		End If
	    ballShadow(b).Y = BOT(b).Y + 20
		If BOT(b).Z > 20 Then
			BallShadow(b).visible = 1
		Else
			BallShadow(b).visible = 0
		End If
	Next
End Sub


' *** NFozzy's lamp fade routines *** 


Class NullFadingObject : Public Property Let IntensityScale(input) : : End Property : End Class	'todo do better

Class LampFader
	Public FadeSpeedDown(140), FadeSpeedUp(140)
	Private Lock(140), Loaded(140), OnOff(140)
	Public UseFunction
	Private cFilter
	Private UseCallback(140), cCallback(140)
	Public Lvl(140), Obj(140)

	Sub Class_Initialize()
		dim x : for x = 0 to uBound(OnOff) 	'Set up fade speeds
			if FadeSpeedDown(x) <= 0 then FadeSpeedDown(x) = 1/100	'fade speed down
			if FadeSpeedUp(x) <= 0 then FadeSpeedUp(x) = 1/80'Fade speed up
			UseFunction = False
			lvl(x) = 0
			OnOff(x) = False
			Lock(x) = True : Loaded(x) = False
		Next

		for x = 0 to uBound(OnOff) 		'clear out empty obj
			if IsEmpty(obj(x) ) then Set Obj(x) = NullFader' : Loaded(x) = True
		Next
	End Sub

	Public Property Get Locked(idx) : Locked = Lock(idx) : End Property		'debug.print Lampz.Locked(100)	'debug
	Public Property Get state(idx) : state = OnOff(idx) : end Property
	Public Property Let Filter(String) : Set cFilter = GetRef(String) : UseFunction = True : End Property
	Public Property Let Callback(idx, String) : cCallback(idx) = String : UseCallBack(idx) = True : End Property

	Public Property Let state(ByVal idx, input) 'Major update path
		input = cBool(input)
		if OnOff(idx) = Input then : Exit Property : End If	'discard redundant updates
		OnOff(idx) = input
		Lock(idx) = False 
		Loaded(idx) = False
	End Property

	Public Sub TurnOnStates()	'If obj contains any light objects, set their states to 1 (Fading is our job!)
		dim debugstr
		dim idx : for idx = 0 to uBound(obj)
			if IsArray(obj(idx)) then 
				dim x, tmp : tmp = obj(idx) 'set tmp to array in order to access it
				for x = 0 to uBound(tmp)
					if typename(tmp(x)) = "Light" then DisableState tmp(x) : debugstr = debugstr & tmp(x).name & " state'd" & vbnewline
					
				Next
			Else
				if typename(obj(idx)) = "Light" then DisableState obj(idx) : debugstr = debugstr & obj(idx).name & " state'd (not array)" & vbnewline
				
			end if
		Next
		debug.print debugstr
	End Sub
	Private Sub DisableState(ByRef aObj) : aObj.FadeSpeedUp = 1000 : aObj.State = 1 : End Sub	'turn state to 1

	Public Sub Update1()	 'Handle all boolean numeric fading. If done fading, Lock(True). Update on a '1' interval Timer!
		dim x : for x = 0 to uBound(OnOff)
			if not Lock(x) then 'and not Loaded(x) then
				if OnOff(x) then 'Fade Up
					Lvl(x) = Lvl(x) + FadeSpeedUp(x)
					if Lvl(x) > 1 then Lvl(x) = 1 : Lock(x) = True
				elseif Not OnOff(x) then 'fade down
					Lvl(x) = Lvl(x) - FadeSpeedDown(x)
					if Lvl(x) < 0 then Lvl(x) = 0 : Lock(x) = True
				end if
			end if
		Next
	End Sub


	Public Sub Update()	'Handle object updates. Update on a -1 Timer! If done fading, loaded(x) = True
		dim x,xx : for x = 0 to uBound(OnOff)
			if not Loaded(x) then
				if IsArray(obj(x) ) Then	'if array
					If UseFunction then 
						for each xx in obj(x) : xx.IntensityScale = cFilter(Lvl(x)) : Next
					Else
						for each xx in obj(x) : xx.IntensityScale = Lvl(x) : Next
					End If
				else						'if single lamp or flasher
					If UseFunction then 
						obj(x).Intensityscale = cFilter(Lvl(x))
					Else
						obj(x).Intensityscale = Lvl(x)
					End If
				end if
				' Sleazy hack for regional decimal point problem
				If UseCallBack(x) then execute cCallback(x) & " CSng(" & CInt(10000 * Lvl(x)) & " / 10000)"	'Callback
				If Lock(x) Then
					if Lvl(x) = 1 or Lvl(x) = 0 then Loaded(x) = True	'finished fading
				end if
			end if
		Next
	End Sub
End Class


' **** Replacement mech class that moves smoothly.   Carbon copy of the cvpmMech
Class cvpmMyMech
	Public Sol1, Sol2, MType, Length, Steps, Acc, Ret
	Private mMechNo, mNextSw, mSw(), mLastPos, mLastSpeed, mCallback

	Private Sub Class_Initialize
		ReDim mSw(10)
		gNextMechNo = gNextMechNo + 1 : mMechNo = gNextMechNo : mNextSw = 0 : mLastPos = 0 : mLastSpeed = 0
		MType = 0 : Length = 0 : Steps = 0 : Acc = 0 : Ret = 0 : vpmTimer.addResetObj Me
	End Sub

	Public Sub AddSw(aSwNo, aStart, aEnd)
		mSw(mNextSw) = Array(aSwNo, aStart, aEnd, 0)
		mNextSw = mNextSw + 1
	End Sub

	Public Sub AddPulseSwNew(aSwNo, aInterval, aStart, aEnd)
		If Controller.Version >= "01200000" Then
			mSw(mNextSw) = Array(aSwNo, aStart, aEnd, aInterval)
		Else
			mSw(mNextSw) = Array(aSwNo, -aInterval, aEnd - aStart + 1, 0)
		End If
		mNextSw = mNextSw + 1
	End Sub

	Public Sub Start
		Dim sw, ii
		With Controller
			.Mech(1) = Sol1 : .Mech(2) = Sol2 : .Mech(3) = Length
			.Mech(4) = Steps : .Mech(5) = MType : .Mech(6) = Acc : .Mech(7) = Ret
			ii = 10
			For Each sw In mSw
				If IsArray(sw) Then
					.Mech(ii) = sw(0) : .Mech(ii+1) = sw(1)
					.Mech(ii+2) = sw(2) : .Mech(ii+3) = sw(3)
					ii = ii + 10
				End If
			Next
			.Mech(0) = mMechNo
		End With
		If IsObject(mCallback) Then mCallBack 0, 0, 0 : mLastPos = 0 : vpmTimer.EnableUpdate Me, True, True  ' <------- All for this.
	End Sub

	Public Property Get Position : Position = Controller.GetMech(mMechNo) : End Property
	Public Property Get Speed    : Speed = Controller.GetMech(-mMechNo)   : End Property
	Public Property Let Callback(aCallBack) : Set mCallback = aCallBack : End Property

	Public Sub Update
		Dim currPos, speed
		currPos = Controller.GetMech(mMechNo)
		speed = Controller.GetMech(-mMechNo)
		If currPos < 0 Or (mLastPos = currPos And mLastSpeed = speed) Then Exit Sub
		mCallBack currPos, speed, mLastPos : mLastPos = currPos : mLastSpeed = speed
	End Sub

	Public Sub Reset : Start : End Sub
	
End Class





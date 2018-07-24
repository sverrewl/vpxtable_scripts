'______       ______           _                 _   _ _       _____             _ _            _    ______
'|  _  \      |  _  \         | |        ___    | | | (_)     |  ___|           | | |          | |   | ___ \ 
'| | | |_ __  | | | |_   _  __| | ___   ( _ )   | |_| |_ ___  | |____  _____ ___| | | ___ _ __ | |_  | |_/ /__ _ _   _
'| | | | '__| | | | | | | |/ _` |/ _ \  / _ \/\ |  _  | / __| |  __\ \/ / __/ _ \ | |/ _ \ '_ \| __| |    // _` | | | |
'| |/ /| |_   | |/ /| |_| | (_| |  __/ | (_>  < | | | | \__ \ | |___>  < (_|  __/ | |  __/ | | | |_  | |\ \ (_| | |_| |
'|___/ |_(_)  |___/  \__,_|\__,_|\___|  \___/\/ \_| |_/_|___/ \____/_/\_\___\___|_|_|\___|_| |_|\__| \_| \_\__,_|\__, |
'                                                                                                                 __/ |
'                                                                                                                |___/
'                                   ______       _ _         __   _____  _____  _____
'                                   | ___ \     | | |       /  | |  _  ||  _  ||  _  |
'                                   | |_/ / __ _| | |_   _  `| | | |_| || |_| || |/' |
'                                   | ___ \/ _` | | | | | |  | | \____ |\____ ||  /| |
'                                   | |_/ / (_| | | | |_| | _| |_.___/ /.___/ /\ |_/ /
'                                   \____/ \__,_|_|_|\__, | \___/\____/ \____/  \___/
'                                                     __/ |
'                                                    |___/                               '
'
'
'Dr. Dude (Midway 1990) Version 2.0 for VP10
'Early development by "gtxJoe" and completed by "wrd1972"
'Thanks to gtxJoe for his  assistance and allowing me to complete this table...which by the way, is my very first VP table...EVER!
'Extra thanks to "ninuzzu" for the invaluable guidance he provided me when I started authoring.
'Also extra thanks to Cyberpez for the many small 3D jobs and script tweaks.
'Additional scripting assistance by "Rothbauerw"
'Flasher Domes, pop-bumper caps, side-rail image by "Flupper"
'Clear ramp by "Bord"
'White flipper prims by "Zany"
'Playfield image rework by "ClarkKent"
'Desktop table scoring displays by "32Assassin"
'Table lighting by "wrd1972"
'HMLF table physics by "wrd1972"
'Flipper physics by "wrd1972" and "Rothbauerw"
'DOF by "Arngrim"
'Mixmaster model and coiled wires 3D models by "nfozzy"
'Totally awesome new light fading scripting by "nFozzy"
'Color-Mod scripting by cyberpez
'PMD feedback and surround sound mods by "Rusty Cardores"
'Many thanks to countles others for in the VPF community for helping me learn the table creation process.

' Thalamus 2018-07-24
' Table has already its own "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' No special SSF tweaks yet.

Option Explicit
Randomize
On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

'***************************************************************************************************************************************************************
'***************************************************************************************************************************************************************
' _____     _     _        ___________ _   _                   _   _
'|_   _|   | |   | |      |  _  | ___ \ | (_)                 | | | |
'  | | __ _| |__ | | ___  | | | | |_/ / |_ _  ___  _ __  ___  | |_| | ___ _ __ ___
'  | |/ _` | '_ \| |/ _ \ | | | |  __/| __| |/ _ \| '_ \/ __| |  _  |/ _ \ '__/ _ \ 
'  | | (_| | |_) | |  __/ \ \_/ / |   | |_| | (_) | | | \__ \ | | | |  __/ | |  __/
'  \_/\__,_|_.__/|_|\___|  \___/\_|    \__|_|\___/|_| |_|___/ \_| |_/\___|_|  \___|
'
'
'INTRO MUSIC
'   Add Intro Music Snippet =0
'   No Intro Music Snippet = 1
'Change the value below to set option
IntroMusic = 0


'APRON COLOR MOD
'   White Apron = 0
'	Blue Apron = 1
'   Green Apron = 2
'   Custom Apron & Walls = 3
'   Random = 4
'Change the value below to set option
Aproncolor = 1


'FLIPPER BAT STYLE MOD
'   Yellow Bats = 0
'   White Bats = 1
'   Random = 2
'Change the value below to set option
Flipperstyle = 0


'CHEATER DRAIN POST MOD
'   Disable Cheater Post = 0
'   Enable Cheater Post = 1
'Change the value below to set option
CheaterPost = 0


'EXCELLENT RAY BEAM MOD
'   No Raybeam = 0
'   Show Raybeam = 1
'Change the value below to set option
Raybeam = 1


'GI COLOR MOD
'   White Bulbs = 0
'   Blue Bulbs = 1
'   Purple Bulbs = 2
'   Ice-Blue Bulbs = 3
'   Random = 4
'Change the value below to set option
GIColorMod = 3


'POP-BUMPER COLOR MOD
'   White Bulbs = 0
'   Blue Bulbs = 1
'   Purple Bulbs = 2
'   Ice-Blue Bulbs = 3
'   Random = 4
'Change the value below to set option
BumperColor =4


'Remove Side Rails
'   Show Side Rails = 0
'   Hide Side Rails = 1
'Change the value below to set option
Rails = 0

'***************************************************************************************************************************************************************
'***************************************************************************************************************************************************************



Dim DesktopMode: DesktopMode = Table1.ShowDT
If DesktopMode = True Then 'Show Desktop components
Else
End if


' DMD Rotation
Const cDMDRotation 				= -1 			'-1 for No change, 0 - DMD rotation of 0?, 1 - DMD rotation of 90?
Const cGameName 				= "dd_l2"		'ROM name
Const ballsize = 50
Const ballmass = 1.7
Const DynamicFlipperFriction = True
Const DynamicFlipperFrictionResting = 0.4
Const DynamicFlipperFrictionActive = 1.0


LoadVPM "00990300", "s11.VBS",3.10
SetLocale(1033)
'********************
'Standard definitions
'********************
Const UseSolenoids = 2
Const UseLamps = 0
Const UseSync = 1
Const HandleMech = 0

'Standard Sounds
Const SSolenoidOn = "fx_Solenoid"
Const SSolenoidOff = ""
Const SFlipperOn = ""
Const SFlipperOff = ""
Const SCoin = "fx_coin"


'********************
'Table Init
'********************

Dim bsTrough, DTRBank1, ttcentre, mMagnet

Sub Table1_Init
	SetOptions
	vpmInit Me
	With Controller
        .GameName = cGameName
        .SplashInfoLine = "Dr. Dude (Bally 1990)"
		.HandleKeyboard = 0
		.ShowTitle = 0
		.ShowDMDOnly = 1
		.ShowFrame = 0
		.HandleMechanics = 1
		.Hidden = 1

		if cDMDRotation >= 0 Then .Games(cGameName).Settings.Value("rol") = cDMDRotation

		On Error Resume Next
		.Run GetPlayerHWnd
		If Err Then MsgBox Err.Description
	End With

    On Error Goto 0


	'Nudging
	vpmNudge.TiltSwitch=1
	vpmNudge.Sensitivity=5
	vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingshot)

   'Trough
	Set bsTrough=New cvpmBallStack
        with bsTrough
			.InitSw 0,11,12,13,0,0,0,0
			.InitKick BallRelease,0,12
			.Balls=3
			.InitExitSnd SoundFX("fx_ballrel",DOFContactors) , SoundFX("fx_Solenoid",DOFContactors)
		End With


 	'DropTargets
	Set dtrbank1 = New cvpmDropTarget
		dtrbank1.InitDrop Array(sw21, sw22, sw23, sw24), Array(21,22,23,24)
		dtrbank1.InitSnd SoundFX("Droptarget",DOFContactors),SoundFX("Droptargetreset",DOFContactors)


	' Disc
 	Set ttcentre = new cvpmTurnTable
	ttcentre.InitTurnTable trMixMaster,240
	ttcentre.CreateEvents "ttcentre"
	ttcentre.SpinUp = 240
	ttcentre.SpinDown = 240

	' Magnet
	Set mMagnet = New cvpmMagnet
		With mMagnet
			.InitMagnet magnetTrigger, 10
			.CreateEvents "mMagnet"
			.Solenoid = 13
			.GrabCenter = False
		End With

    Dim obj
    For Each obj In colLampPoles:obj.IsDropped = 1:Next
	colLampPoles(2).isdropped = 0

    '**Main Timer init
    PinMAMETimer.Enabled = 1

    PrevGameOver = 0
	StartLevel


	If GIColorModType = 0 then
	'White
    GIWallWhite. visible = 1
    GIWallBlue. visible = 0
    GIWallPurple. visible = 0
    GIWallIceBlue. visible = 0

    GIWallWhite2. visible = 1
    GIWallBlue2. visible = 0
    GIWallPurple2. visible = 0
    GIWallIceBlue2. visible = 0
End If


	If BumperColorType = 0 then
	'WhiteBumperColor
    GIWallWhite3. visible = 1
    GIWallBlue3. visible = 0
    GIWallPurple3. visible = 0
    GIWallIceBlue3. visible = 0
End If


	If GIColorModType = 1 then
	'Blue
    GIWallWhite. visible = 0
    GIWallBlue. visible = 1
    GIWallPurple. visible = 0
    GIWallIceBlue. visible = 0

    GIWallWhite2. visible = 0
    GIWallBlue2. visible = 1
    GIWallPurple2. visible = 0
    GIWallIceBlue2. visible = 0
End If


	If BumperColorType = 1 then
	'Blue
    GIWallWhite3. visible = 0
    GIWallBlue3. visible = 1
    GIWallPurple3. visible = 0
    GIWallIceBlue3. visible = 0

End If


	If GIColorModType = 2 then
	'Purple
    GIWallWhite. visible = 0
    GIWallBlue. visible = 0
    GIWallPurple. visible = 1
    GIWallIceBlue. visible = 0

    GIWallWhite2. visible = 0
    GIWallBlue2. visible = 0
    GIWallPurple2. visible = 1
    GIWallIceBlue2. visible = 0
End If

	If BumperColorType = 2 then
	'Purple
    GIWallWhite3. visible = 0
    GIWallBlue3. visible = 0
    GIWallPurple3. visible = 1
    GIWallIceBlue3. visible = 0
End If


	If GIColorModType = 3 then
	'IceBlue
    GIWallWhite. visible = 0
    GIWallBlue. visible = 0
    GIWallPurple. visible = 0
    GIWallIceBlue. visible = 1

    GIWallWhite2. visible = 0
    GIWallBlue2. visible = 0
    GIWallPurple2. visible = 0
    GIWallIceBlue2. visible = 1
End If


	If BumperColorType = 3 then
	'Ice Blue
    GIWallWhite3. visible = 0
    GIWallBlue3. visible = 0
    GIWallPurple3. visible = 0
    GIWallIceBlue3. visible = 1
End If


End Sub


Sub B2SCommand(nr, state)
	If B2SOn Then
		Controller.B2SSetData nr, state
	End If
End Sub

Sub Table_Paused:Controller.Pause = 1:End Sub
Sub Table_unPaused:Controller.Pause = 0:End Sub
Sub Table1_Exit():Controller.Stop:End Sub


'********************
'		KEYS
'********************

Sub Table1_KeyDown(ByVal keycode)
    If keycode = PlungerKey Then Plunger.PullBack:PlaySoundAt "plungerpull",Plunger
    If keycode = LeftFlipperKey Then Controller.Switch(58) = 1
    If keycode = RightFlipperKey Then Controller.Switch(57) = 1
    If keycode = LeftTiltKey Then Nudge 90,2: Playsound SoundFX("fx_nudge",0)
    If keycode = RightTiltKey Then Nudge 270,2: Playsound SoundFX("fx_nudge",0)
    If keycode = CenterTiltKey Then Nudge 0,3: Playsound SoundFX("fx_nudge",0)

'CP test Key

	If keycode = 22 then  ''''''''''''''''''''U Key used for testing



		NudgeStrangthY = 13

		ShakeXMax = 9
		ShakeXMin = -3
		ShakeXDir = -1
		ShakeYMax = 1
		ShakeYMin = -1
		ShakeYDir = -1
		ShakeX.Enabled = True
		ShakeY.Enabled = True

	End If

'   '************************   Start Ball Control 1/3
'        if keycode = 46 then                ' C Key
'            If contball = 1 Then
'                contball = 0
'            Else
'                contball = 1
'            End If
'        End If
'        if keycode = 48 then                'B Key
'            If bcboost = 1 Then
'                bcboost = bcboostmulti
'            Else
'                bcboost = 1
'            End If
'        End If
'        if keycode = 203 then bcleft = 1        ' Left Arrow
'        if keycode = 200 then bcup = 1          ' Up Arrow
'        if keycode = 208 then bcdown = 1        ' Down Arrow
'        if keycode = 205 then bcright = 1       ' Right Arrow
'    '************************   End Ball Control 1/3

    If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal keycode)
    If keycode = PlungerKey Then Plunger.Fire:PlaySoundAt "plunger",Plunger
    If keycode = LeftFlipperKey Then Controller.Switch(58) = 0
    If keycode = RightFlipperKey Then Controller.Switch(57) = 0

'    '************************   Start Ball Control 2/3
'    if keycode = 203 then bcleft = 0        ' Left Arrow
'    if keycode = 200 then bcup = 0          ' Up Arrow
'    if keycode = 208 then bcdown = 0        ' Down Arrow
'    if keycode = 205 then bcright = 0       ' Right Arrow
'    '************************   End Ball Control 2/3

    If vpmKeyUp(keycode) Then Exit Sub
End Sub

''************************   Start Ball Control 3/3
'Sub StartControl_Hit()
'    Set ControlBall = ActiveBall
'    contballinplay = true
'End Sub
'
'Sub StopControl_Hit()
'    contballinplay = false
'End Sub
'
'Dim bcup, bcdown, bcleft, bcright, contball, contballinplay, ControlBall, bcboost
'Dim bcvel, bcyveloffset, bcboostmulti
'
'bcboost = 1     'Do Not Change - default setting
'bcvel = 4       'Controls the speed of the ball movement
'bcyveloffset = -0.01    'Offsets the force of gravity to keep the ball from drifting vertically on the table, should be negative
'bcboostmulti = 3    'Boost multiplier to ball veloctiy (toggled with the B key)
'
'Sub BallControl_Timer()
'    If Contball and ContBallInPlay then
'        If bcright = 1 Then
'            ControlBall.velx = bcvel*bcboost
'        ElseIf bcleft = 1 Then
'            ControlBall.velx = - bcvel*bcboost
'        Else
'            ControlBall.velx=0
'        End If
'
'        If bcup = 1 Then
'            ControlBall.vely = -bcvel*bcboost
'        ElseIf bcdown = 1 Then
'            ControlBall.vely = bcvel*bcboost
'        Else
'            ControlBall.vely= bcyveloffset
'        End If
'    End If
'End Sub
''************************   End Ball Control 3/3


'*********** SOLENOIDS ************

SolCallBack(1)	= "bsTrough.SolIn"
SolCallBack(2)	= "bsTrough.SolOut"
SolCallBack(3)	= "solLeftKicker"
SolCallBack(4)	= "solRightKicker"
SolCallBack(6)	= "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallBack(7)  = "dtRBank1.SolDropUp"
'SolCallBack(10) = "solGI"
SolCallBack(10) = "SetGI"	'setlamp 110
SolCallBack(14)	= "solBigGuy"
SolCallBack(15) = "SetLamp 115,"				'Big Guy Flasher
SolCallBack(16) = "solMixer"					'Mixer Motor
SolCallBack(25) = "SetLamp 125,"				'Mixer Heart Flasher
SolCallBack(26) = "SetLamp 126,"				'Mixer Gab Flasher
SolCallBack(27) = "SetLamp 127,"				'Mixer Magnet Flasher
SolCallBack(28) = "SetLamp 128,"				'Magnet Flasher
SolCallBack(29) = "SetLamp 129,"				'Gab Flasher
SolCallBack(30) = "SetLamp 130,"				'Heart Flasher
SolCallBack(31) = "SetLamp 131,"				'Drop Targets Flasher
SolCallBack(32) = "SetLamp 132,"				'Raygun Flasher
SolCallback(sLLFlipper) = "SolLFlipper"
SolCallback(sLRFlipper) = "SolRFlipper"
'SolCallback(23) = "FastFlips.TiltSol"  'handled by core.vbs now


Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFX("fx_flipperup",DOFContactors),LeftFlipper,2
		if DynamicFlipperFriction Then LeftFlipper.Friction = DynamicFlipperFrictionActive
        LeftFlipper.RotateToEnd
    Else
        PlaySoundAt SoundFX("fx_flipperdown",DOFContactors),LeftFlipper
		if DynamicFlipperFriction Then LeftFlipper.Friction = DynamicFlipperFrictionActive
        LeftFlipper.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFX("fx_flipperup",DOFContactors),RightFlipper,2
		if DynamicFlipperFriction Then RightFlipper.Friction = DynamicFlipperFrictionActive
        RightFlipper.RotateToEnd
    Else
        PlaySoundAt SoundFX("fx_flipperdown",DOFContactors),RightFlipper
		if DynamicFlipperFriction Then LeftFlipper.Friction = DynamicFlipperFrictionActive
        RightFlipper.RotateToStart
    End If
End Sub


Sub solLeftKicker (enabled)
	If (enabled and Controller.Switch(51)) Then
		PlaysoundAtVol SoundFX("fx_vuk_exit2",DOFContactors), sw51, .8
		Controller.Switch(51) = 0
		sw51.timerenabled = 1
		sw51.kick 0,40, 3.14/2
	End If
End Sub


'@@@@@@@@@@@@@@@@@@@@@@@@@
' Level
'@@@@@@@@@@@@@@@@@@@@@@@@@

'Dim ShakeXMin, ShakeXMax, ShakeXDir, ShakeXSpeed, ShakeXDamper
'Dim ShakeYMin, ShakeYMax, ShakeYDir, ShakeYSpeed, ShakeYDamper

Dim lBall1, lBall2
Sub StartLevel()

	kLevel.Enabled = 1
	Set lBall1 = kLevel.CreateSizedBallWithMass(4, .005)

	kLevel.kick 0, 0
	kLevel.Enabled = 0

	kLevel1.Enabled = 1
	Set lBall2 = kLevel1.CreateSizedBallWithMass(4, .005)

	kLevel1.kick 0, 0
	kLevel1.Enabled = 0

End Sub

Dim NudgeStrangthY, NudgeStrangthX

Sub Level_Timer()


'	fBubble.y = lBall1.y - 60  'Moves bubble in level

	NudgeStrangthX = ((kLevel.y - lBall1.y) / 2)
	NudgeStrangthY = ((kLevel1.x - lBall2.x) / 2)


	NudgeStrangthY = Round(NudgeStrangthY)
	If NudgeStrangthY = 1 then NudgeStrangthY = 0 End If

'	NudgeStrangthX = ((kLevel1.x - lBall2.x) / 2)
	NudgeStrangthX = Round(NudgeStrangthX)
	If NudgeStrangthX = 1 then NudgeStrangthX = 0 End If

'	TextBox3.Text = NudgeStrangthY


	If NudgeStrangthY > 1 then
		If ShakeYMax > NudgeStrangthY then

		Else

			ShakeYMax = NudgeStrangthY
			ShakeYMin = ShakeYMax * -1
			ShakeYDir = 1
			ShakeY.Enabled = True
		End If

	Else

		If NudgeStrangthY < -1 then
			If ShakeYMin < NudgeStrangthY then

			Else

				ShakeYMin = NudgeStrangthY
				ShakeYMax = ShakeYMin * -1
				ShakeYDir = -1
				ShakeY.Enabled = True
			End If
		End If
	End If


	If NudgeStrangthX > 1 then

		If ShakeXMax > NudgeStrangthX Then

		Else
			If NudgeStrangthX < 1 then NudgeStrangthX = 0 End If
			ShakeXMax = NudgeStrangthX
			ShakeXMin = ShakeXMax * -1
			ShakeXDir = 1
			ShakeX.Enabled = True
		End If

	Else

		If NudgeStrangthX < -1 then
			If ShakeXMin < NudgeStrangthX Then

			Else
				If NudgeStrangthX > -1 then NudgeStrangthX = 0 End If
				ShakeXMin = NudgeStrangthX
				ShakeXMax = ShakeXMin * -1
				ShakeXDir = -1
				ShakeX.Enabled = True
			End If
		End If
	End If



End Sub


'@@@@@@@@@@@@@@@@@@@@@@@@@
' End Level
'@@@@@@@@@@@@@@@@@@@@@@@@@

'^^^^^^^^^^^^^^^^^^^^^^^^^
' The Shakes
'^^^^^^^^^^^^^^^^^^^^^^^^^

Dim UfoLedPos, cBall, ufoalternate


Sub SolUfoShake(Enabled)
    If Enabled Then


		ShakeXMax = 9
		ShakeXMin = -3
		ShakeXDir = -1
		ShakeYMax = 1
		ShakeYMin = -1
		ShakeYDir = -1
		ShakeX.Enabled = True
		ShakeY.Enabled = True

	Else


    End If
End Sub


Dim ShakeXMin, ShakeXMax, ShakeXDir, ShakeXSpeed, ShakeXDamper
Dim ShakeYMin, ShakeYMax, ShakeYDir, ShakeYSpeed, ShakeYDamper


Sub ShakeX_timer()


Select Case ShakeXMax

	Case 1:ShakeXSpeed = 1
	Case 2:ShakeXSpeed = 1.1
	Case 3:ShakeXSpeed = 1.2
	Case 4:ShakeXSpeed = 1.3
	Case 5:ShakeXSpeed = 1.4
	Case 6:ShakeXSpeed = 1.5
	Case 7:ShakeXSpeed = 1.6
	Case 8:ShakeXSpeed = 1.7
	Case 9:ShakeXSpeed = 1.8
	Case 10:ShakeXSpeed = 1.9
	Case 11:ShakeXSpeed = 2
	Case 12:ShakeXSpeed = 2.1
	Case 13:ShakeXSpeed = 2.2
	Case 14:ShakeXSpeed = 2.3
	Case 15:ShakeXSpeed = 2.4
	Case 16:ShakeXSpeed = 2.5
	Case 17:ShakeXSpeed = 2.6
	Case 18:ShakeXSpeed = 2.7
	Case 19:ShakeXSpeed = 2.8
	Case 20:ShakeXSpeed = 2.9
	Case 21:ShakeXSpeed = 3

End Select


If ShakeXMax > 11 then ShakeXDamper = 1 End If
If ShakeXMax < 10 then ShakeXDamper = .5 End If
PrimGuy.ObjRotx = PrimGuy.ObjRotx + ShakeXDir * ShakeXSpeed


	If PrimGuy.Objrotx <= ShakeXMin Then

		ShakeXDir = 1
		ShakeXMax = ShakeXMax - ShakeXDamper
	End If

	If PrimGuy.Objrotx >= ShakeXMax Then

		ShakeXDir = -1
		ShakeXMin = ShakeXMin + ShakeXDamper
	End If


If 	ShakeXMax = 0 or ShakeXMin = 0 Then
 ShakeX.Enabled = false
End If

End Sub


Sub ShakeY_timer()


Select Case ShakeYMax

	Case 1:ShakeYSpeed = 1
	Case 2:ShakeYSpeed = 1.1
	Case 3:ShakeYSpeed = 1.2
	Case 4:ShakeYSpeed = 1.3
	Case 5:ShakeYSpeed = 1.4
	Case 6:ShakeYSpeed = 1.5
	Case 7:ShakeYSpeed = 1.6
	Case 8:ShakeYSpeed = 1.7
	Case 9:ShakeYSpeed = 1.8
	Case 10:ShakeYSpeed = 1.9
	Case 11:ShakeYSpeed = 2
	Case 12:ShakeYSpeed = 2.1
	Case 13:ShakeYSpeed = 2.2
	Case 14:ShakeYSpeed = 2.3
	Case 15:ShakeYSpeed = 2.4
	Case 16:ShakeYSpeed = 2.5
	Case 17:ShakeYSpeed = 2.6
	Case 18:ShakeYSpeed = 2.7
	Case 19:ShakeYSpeed = 2.8
	Case 20:ShakeYSpeed = 2.9
	Case 21:ShakeYSpeed = 3

End Select


If ShakeYMax > 11 then ShakeYDamper = 1 End If

If ShakeYMax < 10 then ShakeYDamper = .5 End If

PrimGuy.ObjRotY = PrimGuy.ObjRotY + ShakeYDir * ShakeYSpeed


	If PrimGuy.ObjrotY <= ShakeYMin Then

		ShakeYDir = 1
		ShakeYMax = ShakeYMax - ShakeYDamper
	End If

	If PrimGuy.ObjrotY >= ShakeYMax Then

		ShakeYDir = -1
		ShakeYMin = ShakeYMin + ShakeYDamper
	End If

'End If


If 	ShakeYMax = 0 or ShakeYMin = 0 Then
 ShakeY.Enabled = false
End If

End Sub


'ShipDampon = 1


'^^^^^^^^^^^^^^^^^^^^^^^^^
' End The Shakes
'^^^^^^^^^^^^^^^^^^^^^^^^^

Dim sw51step

Sub sw51_timer()	' I have a kickass keyframe animation script for this kind of thing if you want it -nf
	Select Case sw51step
		Case 0:Leftupkicker.TransY = 10
		Case 1:Leftupkicker.TransY = 20
		Case 2:Leftupkicker.TransY = 30
		Case 3:'pUpKicker.TransY = 30
		Case 4:
		Case 5:Leftupkicker.TransY = 25
		Case 6:Leftupkicker.TransY = 20
		Case 7:Leftupkicker.TransY = 15
		Case 8:Leftupkicker.TransY = 10
		Case 9:Leftupkicker.TransY = 5
		Case 10:Leftupkicker.TransY = 0:sw51.timerEnabled = 0:sw51step = 0
	End Select
	sw51step = sw51step + 1
End Sub


Sub solRightKicker (enabled)
	If (enabled) Then
		PlaysoundAtVol SoundFX("fx_vuk_exit2",DOFContactors), sw32, .8
		Controller.Switch(32) = 0
		sw32.timerenabled = 1
		sw32.kick 0,70, 3.14/2
	End If
End Sub

Dim sw32step

Sub sw32_timer()
	Select Case sw32step
		Case 0:Rightupkicker.TransY = 10
		Case 1:Rightupkicker.TransY = 20
		Case 2:Rightupkicker.TransY = 30
		Case 3:'pUpKicker.TransY = 30
		Case 4:
		Case 5:Rightupkicker.TransY = 25
		Case 6:Rightupkicker.TransY = 20
		Case 7:Rightupkicker.TransY = 15
		Case 8:Rightupkicker.TransY = 10
		Case 9:Rightupkicker.TransY = 5
		Case 10:Rightupkicker.TransY = 0:sw32.timerEnabled = 0:sw32step = 0
	End Select
	sw32step = sw32step + 1
End Sub


Sub solBigGuy (enabled)
	If Enabled Then
		PlaySoundAtVol SoundFX("BigGuyShake",DOFShaker),Light79,4
'		PlaySound "ShakerPulse",0,1 'Uncomment for use in cabs with PMD Audio Shaker. Change "1" to decimal to decrease effect.
		PrimGuyHit
	End If
End Sub
Sub PrimGuyTimer_Timer: PrimGuyMove: End Sub
Dim GuyCnt
Const GuyMoveMax = 20
Sub PrimGuyHit
	GuyCnt = 0 					'Reset count
	PrimGuyTimer.Interval = 20 	'Set timer interval
	PrimGuyTimer.Enabled = 1 	'Enable timer
End Sub
Sub	PrimGuyMove
	Select Case GuyCnt
		Case 0: 	primGuy.RotX = GuyMoveMax * .25
		Case 1: 	primGuy.RotX = GuyMoveMax * .50
		Case 2: 	primGuy.RotX = GuyMoveMax * .75
		Case 3: 	primGuy.RotX = GuyMoveMax
		Case 4: 	primGuy.RotX = GuyMoveMax * .25
		Case 5: 	primGuy.RotX = GuyMoveMax * .50
		Case 6: 	primGuy.RotX = GuyMoveMax * .75
		Case 7: 	primGuy.RotX = 0:PrimGuyShake
		Case else: 	PrimGuyTimer.Enabled = 0
	End Select
	GuyCnt = GuyCnt + 1
End Sub


'***Mixmaster motor***
Sub SolMixer(enabled)
if enabled then
ttcentre.MotorOn = True
ttTimer.Enabled = True
Playsound SoundFX("fx_motor",DOFGear), -1, .5, Pan(trMixMaster), 0, 0, 1, 0,AudioFade(trMixMaster)
else
ttcentre.MotorOn = False
ttTimer.Enabled = False
StopSound "fx_motor"
end if
end sub


'***Mixmaster drop posts***
ttTimer.interval = 10
Dim lampLastPos
Sub ttTimer_Timer
    PostMM.RotZ=(PostMM.RotZ + 10) mod 360
	lampLastPos = Int(SpinningDisc.ObjRotz / 10 + .5)
    colLampPoles(lampLastPos).IsDropped = True
	SpinningDisc.ObjRotZ = (SpinningDisc.ObjRotz + 10) mod 360
	lampLastPos = Int(SpinningDisc.ObjRotz / 10 + .5)
    colLampPoles(lampLastPos).IsDropped = False
End Sub


Sub Wall100_Hit
	Wall100Cnt = Wall100Cnt + 1
	If Wall100Cnt >= Wall100Max Then Wall100.IsDropped = True
	'debug.print Wall100Cnt & " , " & Wall100Max
End Sub

Sub Drain_Hit()
	vpmTimer.PulseSw 10
	bsTrough.AddBall Me
	PlaySoundAtVol "drain",Drain,.5
End Sub

'****************
'  POP BUMPERS
'****************

Dim Bumper1Cnt,Bumper2Cnt,Bumper3Cnt
Const BumperMoveMax = 20
Sub Bumper1_Hit
        vpmTimer.PulseSw 52
        PlaysoundAtVol SoundFX("fx_bumper1",DOFContactors),Bumper1,.4
        Bumper1Hit
End Sub
Sub Bumper1_Timer: Bumper1Move: End Sub
Sub Bumper1Hit
    Bumper1Cnt = 0                  'Reset count
    Bumper1.TimerInterval = 20  'Set timer interval
    Bumper1.TimerEnabled = 1    'Enable timer
End Sub

Sub Bumper1Move
    Select Case Bumper1Cnt
        Case 0:     Bumper1Cap.TransY = -BumperMoveMax * .25 :   Bumper1CapS1.TransZ = BumperMoveMax * .25 :Bumper1CapS2.TransZ = BumperMoveMax * .25 :   BR1.TransY = -BumperMoveMax * .25 : BumperSkirt1.RotZ = -BumperMoveMax * .10
        Case 1:     Bumper1Cap.TransY = -BumperMoveMax * .50 :   Bumper1CapS1.TransZ = BumperMoveMax * .50 :Bumper1CapS2.TransZ = BumperMoveMax * .50 :   BR1.TransY = -BumperMoveMax * .50 : BumperSkirt1.RotZ = -BumperMoveMax * .20
        Case 2:     Bumper1Cap.TransY = -BumperMoveMax * .75 :   Bumper1CapS1.TransZ = BumperMoveMax * .75 :Bumper1CapS2.TransZ = BumperMoveMax * .75 :   BR1.TransY = -BumperMoveMax * .75 : BumperSkirt1.RotZ = -BumperMoveMax * .30
        Case 3:     Bumper1Cap.TransY = -BumperMoveMax        :  Bumper1CapS1.TransZ = BumperMoveMax :Bumper1CapS2.TransZ = BumperMoveMax :  br1.TransY = -BumperMoveMax :  BumperSkirt1.RotZ = -BumperMoveMax
        Case 4:     Bumper1Cap.TransY = -BumperMoveMax * .25 :   Bumper1CapS1.TransZ = BumperMoveMax * .25 :Bumper1CapS2.TransZ = BumperMoveMax * .25 :   BR1.TransY = -BumperMoveMax * .25 : BumperSkirt1.RotZ = BumperMoveMax * .10
        Case 5:     Bumper1Cap.TransY = -BumperMoveMax * .50 :   Bumper1CapS1.TransZ = BumperMoveMax * .50 :Bumper1CapS2.TransZ = BumperMoveMax * .50 :   BR1.TransY = -BumperMoveMax * .50 : BumperSkirt1.RotZ = BumperMoveMax * .20
        Case 6:     Bumper1Cap.TransY = -BumperMoveMax * .75 :   Bumper1CapS1.TransZ = BumperMoveMax * .75 :Bumper1CapS2.TransZ = BumperMoveMax * .75 :   BR1.TransY = -BumperMoveMax * .75 : BumperSkirt1.RotZ = BumperMoveMax * .30
        Case 7:     Bumper1Cap.TransY = 0 : Bumper1CapS1.TransZ = 0 :Bumper1CapS2.TransZ = 0 : br1.TransY = 0: BumperSkirt1.RotZ = 0
        Case else:  Bumper1.TimerEnabled = 0
    End Select
    Bumper1Cnt = Bumper1Cnt + 1
End Sub


Sub Bumper2_Hit
	vpmTimer.PulseSw 53
	PlaysoundAtVol SoundFX("fx_bumper2",DOFContactors), Bumper2,.4
	Bumper2Hit
End Sub
Sub Bumper2_Timer: Bumper2Move: End Sub
Sub Bumper2Hit
    Bumper2Cnt = 0                  'Reset count
    Bumper2.TimerInterval = 20  'Set timer interval
    Bumper2.TimerEnabled = 1    'Enable timer
End Sub

Sub Bumper2Move
    Select Case Bumper2Cnt
        Case 0:     Bumper2Cap.TransY = -BumperMoveMax * .25 :   Bumper2CapS1.TransZ = BumperMoveMax * .25 :Bumper2CapS2.TransZ = BumperMoveMax * .25 :   BR2.TransY = -BumperMoveMax * .25 : BumperSkirt2.RotZ = -BumperMoveMax * .10
        Case 1:     Bumper2Cap.TransY = -BumperMoveMax * .50 :   Bumper2CapS1.TransZ = BumperMoveMax * .50 :Bumper2CapS2.TransZ = BumperMoveMax * .50 :   BR2.TransY = -BumperMoveMax * .50 : BumperSkirt2.RotZ = -BumperMoveMax * .20
        Case 2:     Bumper2Cap.TransY = -BumperMoveMax * .75 :   Bumper2CapS1.TransZ = BumperMoveMax * .75 :Bumper2CapS2.TransZ = BumperMoveMax * .75 :   BR2.TransY = -BumperMoveMax * .75 : BumperSkirt2.RotZ = -BumperMoveMax * .30
        Case 3:     Bumper2Cap.TransY = -BumperMoveMax        :  Bumper2CapS1.TransZ = BumperMoveMax :Bumper2CapS2.TransZ = BumperMoveMax :  br2.TransY = -BumperMoveMax :  BumperSkirt2.RotZ = -BumperMoveMax
        Case 4:     Bumper2Cap.TransY = -BumperMoveMax * .25 :   Bumper2CapS1.TransZ = BumperMoveMax * .25 :Bumper2CapS2.TransZ = BumperMoveMax * .25 :   BR2.TransY = -BumperMoveMax * .25 : BumperSkirt2.RotZ = BumperMoveMax * .10
        Case 5:     Bumper2Cap.TransY = -BumperMoveMax * .50 :   Bumper2CapS1.TransZ = BumperMoveMax * .50 :Bumper2CapS2.TransZ = BumperMoveMax * .50 :   BR2.TransY = -BumperMoveMax * .50 : BumperSkirt2.RotZ = BumperMoveMax * .20
        Case 6:     Bumper2Cap.TransY = -BumperMoveMax * .75 :   Bumper2CapS1.TransZ = BumperMoveMax * .75 :Bumper2CapS2.TransZ = BumperMoveMax * .75 :   BR2.TransY = -BumperMoveMax * .75 : BumperSkirt2.RotZ = BumperMoveMax * .30
        Case 7:     Bumper2Cap.TransY = 0 : Bumper2CapS1.TransZ = 0 :Bumper2CapS2.TransZ = 0 : br2.TransY = 0: BumperSkirt2.RotZ = 0
        Case else:  Bumper2.TimerEnabled = 0
    End Select
    Bumper2Cnt = Bumper2Cnt + 1
End Sub


Sub Bumper3_Hit
		vpmTimer.PulseSw 54
		PlaysoundAtVol SoundFX("fx_bumper3",DOFContactors), Bumper3, .4
		Bumper3Hit
End Sub
Sub Bumper3_Timer: Bumper3Move: End Sub
Sub Bumper3Hit
    Bumper3Cnt = 0                  'Reset count
    Bumper3.TimerInterval = 20  'Set timer interval
    Bumper3.TimerEnabled = 1    'Enable timer
End Sub

Sub Bumper3Move
    Select Case Bumper3Cnt
        Case 0:     Bumper3Cap.TransY = -BumperMoveMax * .25 :   Bumper3CapS1.TransZ = BumperMoveMax * .25 :Bumper3CapS2.TransZ = BumperMoveMax * .25 :   BR3.TransY = -BumperMoveMax * .25 : BumperSkirt3.RotZ = -BumperMoveMax * .10
        Case 1:     Bumper3Cap.TransY = -BumperMoveMax * .50 :   Bumper3CapS1.TransZ = BumperMoveMax * .50 :Bumper3CapS2.TransZ = BumperMoveMax * .50 :   BR3.TransY = -BumperMoveMax * .50 : BumperSkirt3.RotZ = -BumperMoveMax * .20
        Case 2:     Bumper3Cap.TransY = -BumperMoveMax * .75 :   Bumper3CapS1.TransZ = BumperMoveMax * .75 :Bumper3CapS2.TransZ = BumperMoveMax * .75 :   BR3.TransY = -BumperMoveMax * .75 : BumperSkirt3.RotZ = -BumperMoveMax * .30
        Case 3:     Bumper3Cap.TransY = -BumperMoveMax        :  Bumper3CapS1.TransZ = BumperMoveMax :Bumper3CapS2.TransZ = BumperMoveMax :  br3.TransY = -BumperMoveMax :  BumperSkirt3.RotZ = -BumperMoveMax
        Case 4:     Bumper3Cap.TransY = -BumperMoveMax * .25 :   Bumper3CapS1.TransZ = BumperMoveMax * .25 :Bumper3CapS2.TransZ = BumperMoveMax * .25 :   BR3.TransY = -BumperMoveMax * .25 : BumperSkirt3.RotZ = BumperMoveMax * .10
        Case 5:     Bumper3Cap.TransY = -BumperMoveMax * .50 :   Bumper3CapS1.TransZ = BumperMoveMax * .50 :Bumper3CapS2.TransZ = BumperMoveMax * .50 :   BR3.TransY = -BumperMoveMax * .50 : BumperSkirt3.RotZ = BumperMoveMax * .20
        Case 6:     Bumper3Cap.TransY = -BumperMoveMax * .75 :   Bumper3CapS1.TransZ = BumperMoveMax * .75 :Bumper3CapS2.TransZ = BumperMoveMax * .75 :   BR3.TransY = -BumperMoveMax * .75 : BumperSkirt3.RotZ = BumperMoveMax * .30
        Case 7:     Bumper3Cap.TransY = 0 : Bumper3CapS1.TransZ = 0 :Bumper3CapS2.TransZ = 0 : br3.TransY = 0: BumperSkirt3.RotZ = 0
        Case else:  Bumper3.TimerEnabled = 0
    End Select
    Bumper3Cnt = Bumper3Cnt + 1
End Sub


'**********Sling Shots and Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub LeftSlingShot_Slingshot
    PlaySoundAt SoundFX("fx_left_slingshot",DOFContactors),SLING2
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -25
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
	vpmTimer.PulseSw 55
	'gi1.State = 0:Gi2.State = 0
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 1:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 2:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0:'gi1.State = 1:Gi2.State = 1
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
    PlaySoundAt SoundFX("fx_right_slingshot",DOFContactors),SLING1
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -25
    RStep = 0
    RightSlingShot.TimerEnabled = 1
	vpmTimer.PulseSw 56
	'gi1.State = 0:Gi2.State = 0
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 1:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 2:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0:'gi1.State = 1:Gi2.State = 1
    End Select
    RStep = RStep + 1
End Sub


Sub wall83_Hit:vpmTimer.PulseSw 100:rubberback1.visible = 0::rubberback1a.visible = 1:wall83.timerenabled = 1:End Sub
Sub wall83_timer:rubberback1.visible = 1::rubberback1a.visible = 0: wall83.timerenabled= 0:End Sub

Sub Wall2_Hit:vpmTimer.PulseSw 100:MMrubberl.visible = 0::MMrubberla.visible = 1:Wall2.timerenabled = 1:End Sub
Sub Wall2_timer:MMrubberl.visible = 1::MMrubberla.visible = 0: Wall2.timerenabled= 0:End Sub

Sub Wall3_Hit:vpmTimer.PulseSw 100:MMrubberR.visible = 0::MMrubberRa.visible = 1:Wall3.timerenabled = 1:End Sub
Sub Wall3_timer:MMrubberR.visible = 1::MMrubberRa.visible = 0: Wall3.timerenabled= 0:End Sub


Sub Wall233_Hit:vpmTimer.PulseSw 100:Rubber_white_1.visible = 0::Rubber_white_1a.visible = 1:Wall233.timerenabled = 1:End Sub
Sub Wall233_timer:Rubber_white_1.visible = 1::Rubber_white_1a.visible = 0: Wall233.timerenabled= 0:End Sub

Sub Wall101_Hit:vpmTimer.PulseSw 100:Rubber_white_3.visible = 0::Rubber_white_3a.visible = 1:Wall101.timerenabled = 1: vpmTimer.PulseSw 39: End Sub
Sub Wall101_timer:Rubber_white_3.visible = 1::Rubber_white_3a.visible = 0: Wall101.timerenabled= 0:End Sub

Sub Wall32_Hit:vpmTimer.PulseSw 100:Rubber_white_15.visible = 0::Rubber_white_17.visible = 1:Wall32.timerenabled = 1:End Sub
Sub Wall32_timer:Rubber_white_15.visible = 1::Rubber_white_17.visible = 0: Wall32.timerenabled= 0:End Sub


' ************************
'      RealTime Updates
' ************************

'Set MotorCallback = GetRef("GameTimer")

Sub GameTimer_Timer
    'UpdateMechs
	RollingSoundsUpdate
    UpdateGatesSpinners
End Sub


Sub Flipperstimer_Timer
    flipperL.rotz = LeftFlipper.CurrentAngle
	batleftshadow.objrotz = LeftFlipper.CurrentAngle

    flipperR.RotZ = RightFlipper.CurrentAngle
	batrightshadow.objrotz = RightFlipper.CurrentAngle

End Sub


Sub ramphelper_Hit()
	ActiveBall.vely = Activeball.vely*1.2
End Sub

'
''***********************************************************************************
''Primitve Flippers
''***********************************************************************************
'
'Sub FlippersTimer_Timer()
'	flipperbatright1.RotAndTra8 = RightFlipper1.CurrentAngle + 90
'
'
'    Flipperbatleft.RotAndTra8 = LeftFlipper.CurrentAngle - 90
'	batleftshadow.objrotz = LeftFlipper.CurrentAngle
'
'    Flipperbatright.RotAndTra8 = RightFlipper.CurrentAngle + 90
'	batrightshadow.objrotz = RightFlipper.CurrentAngle
'End Sub


Dim primCnt(100), primDir(100), primBmprDir(6)
'****************************************************************************
'***** Primitive Standup Target Animation
'****************************************************************************
'USAGE: 	Sub sw1_Hit: 	PrimStandupTgtHit  1, Sw1, PrimSw1: End Sub
'USAGE: 	Sub Sw1_Timer: 	PrimStandupTgtMove 1, Sw1, PrimSw1: End Sub

Const StandupTgtMovementDir = "TransX"
Const StandupTgtMovementMax = 6

Sub PrimStandupTgtHit (swnum, wallName, primName)
	PlaySoundAtBallVol SoundFx("fx_target",DOFContactors),1
	vpmTimer.PulseSw swnum
	primCnt(swnum) = 0 									'Reset count
	wallName.TimerInterval = 20 	'Set timer interval
	wallName.TimerEnabled = 1 	'Enable timer
	'Debug.print "Hit"
End Sub

Sub	PrimStandupTgtMove (swnum, wallName, primName)
	Select Case StandupTgtMovementDir
		Case "TransX":
			Select Case primCnt(swnum)
				Case 0: 	primName.TransX = -StandupTgtMovementMax * .5
				Case 1: 	primName.TransX = -StandupTgtMovementMax
				Case 2: 	primName.TransX = -StandupTgtMovementMax * .5
				Case 3: 	primName.TransX = 0
				Case else: 	wallName.TimerEnabled = 0
			End Select
		Case "TransY":
			Select Case primCnt(swnum)
				Case 0: 	primName.TransY = -StandupTgtMovementMax * .5
				Case 1: 	primName.TransY = -StandupTgtMovementMax
				Case 2: 	primName.TransY = -StandupTgtMovementMax * .5
				Case 3: 	primName.TransY = 0
				Case else: 	wallName.TimerEnabled = 0
			End Select
		Case "TransZ":
			Select Case primCnt(swnum)
				Case 0: 	primName.TransZ = -StandupTgtMovementMax * .5
				Case 1: 	primName.TransZ = -StandupTgtMovementMax
				Case 2: 	primName.TransZ = -StandupTgtMovementMax * .5
				Case 3: 	primName.TransZ = 0
				Case else: 	wallName.TimerEnabled = 0
			End Select
	End Select
	primCnt(swnum) = primCnt(swnum) + 1
End Sub


' ************************
'      Gates
' ************************
Sub Gate1_Hit():PlaySoundAtBall "fx_gate":End Sub
Sub Gate5_Hit():PlaySoundAt "fx_gate",Gate5:End Sub
Sub Gate9_Hit():PlaySoundAt "fx_gate",Gate9:End Sub


'Sub GameTimer_Timer()
'    UpdateGatesSpinners
'End Sub

Dim Pi, GateSpeed
Pi = Round(4*Atn(1),6)
GateSpeed = 0.5

Dim Gate1Open,Gate1Angle:Gate1Open=0:Gate1Angle=0
Sub Gate1_Hit():Gate1Open=1:Gate1Angle=0:End Sub

Dim Gate2Open,Gate2Angle:Gate2Open=0:Gate2Angle=0
Sub Gate2_Hit():Gate2Open=1:Gate2Angle=0:End Sub

Dim Gate9Open,Gate9Angle:Gate9Open=0:Gate9Angle=0
Sub Gate9_Hit():Gate9Open=1:Gate9Angle=0: End Sub

Sub UpdateGatesSpinners

    If Gate2Open Then
        If Gate2Angle < Gate2.currentangle Then:Gate2Angle=Gate2.currentangle:End If
        If Gate2Angle > 5 and Gate2.currentangle < 5 Then:Gate2Open=0:End If
        If Gate2Angle > 70 Then
            Gate2P.RotX = -90
        Else
            Gate2P.RotX = -(Gate2Angle+20)
            Gate2Angle=Gate2Angle - GateSpeed
        End If
    Else
        if Gate2Angle > 0 Then
            Gate2Angle = Gate2Angle - GateSpeed
        Else
            Gate2Angle = 0
        End If
        Gate2P.RotX = -(Gate2Angle + 20)
    End If


    If Gate1Open Then
        If Gate1Angle < Gate1.currentangle Then:Gate1Angle=Gate1.currentangle:End If
        If Gate1Angle > 5 and Gate1.currentangle < 5 Then:Gate1Open=0:End If
        If Gate1Angle > 70 Then
            Gate1P.Rotx = -90
        Else
            Gate1P.Rotx = -(Gate1Angle+20)
            Gate1Angle=Gate1Angle - GateSpeed
        End If
    Else
        if Gate1Angle > 0 Then
            Gate1Angle = Gate1Angle - GateSpeed
        Else
            Gate1Angle = 0
        End If
        Gate1P.Rotx = -(Gate1Angle + 0)
    End If



 If Gate9Open Then
        If Gate9Angle < Gate9.currentangle Then:Gate9Angle=Gate9.currentangle:End If
        If Gate9Angle > 5 and Gate9.currentangle < 5 Then:Gate9Open=0:End If
        If Gate9Angle > 70 Then
            Gate9P.Rotx = 0
        Else
            Gate9P.Rotx = -(Gate9Angle+20)
            Gate9Angle=Gate9Angle - GateSpeed
        End If
    Else
        if Gate9Angle > 0 Then
            Gate9Angle = Gate9Angle - GateSpeed
        Else
            Gate9Angle = 0
        End If
        Gate9P.Rotx = -(Gate9Angle + 270)
    End If

End Sub



' ************************
'      Switches
' ************************

sub sw9_hit:   controller.switch(9)=1: PlaySoundAt "fx_rollover",sw9: StopSound "intro":  end sub
sub sw9_unhit: controller.switch(9)=0
    If ActiveBall.VelY < 0 Then 'on the way up
    End If
End Sub
Sub sw14_Hit:vpmTimer.PulseSw(14):PlaysoundAtVol SoundFX("fx_target",DOFTargets),sw14,2:Me.TimerEnabled = 1: End Sub
'Sub Standup27_Hit:vpmTimer.pulseSw 27:Playsound SoundFX("fx_target",DOFTargets):Me.TimerEnabled = 1: End Sub

Sub sw15_Hit:Controller.Switch(15) = 1
        PlaySoundAtBall "fx_loosemetalplate"
End Sub

Sub sw15_Unhit:Controller.Switch(15) = 0:End Sub
sub sw16_hit:   controller.switch(16)=1: playsoundat "fx_rollover",Sw16: end sub
sub sw16_unhit: controller.switch(16)=0: end sub
sub sw17_hit:   controller.switch(17)=1: playsoundat "fx_rollover",Sw17: end sub
sub sw17_unhit: controller.switch(17)=0: end sub
sub sw18_hit:   controller.switch(18)=1: playsoundat "fx_rollover",Sw18: end sub
sub sw18_unhit: controller.switch(18)=0: end sub
sub sw19_hit:   controller.switch(19)=1: playsoundat "fx_rollover",Sw19: end sub
sub sw19_unhit: controller.switch(19)=0: end sub
sub sw20_hit:   controller.switch(20)=1: playsoundat "fx_rollover",Sw20: end sub
sub sw20_unhit: controller.switch(20)=0: end sub
sub sw21_hit: dtrbank1.hit 1: PlaysoundAt "fx_droptarget",sw21: end sub
sub sw22_hit: dtrbank1.hit 2: PlaysoundAt "fx_droptarget",sw22: end sub
sub sw23_hit: dtrbank1.hit 3: PlaysoundAt "fx_droptarget",sw23: end sub
sub sw24_hit: dtrbank1.hit 4: PlaysoundAt "fx_droptarget",sw24: end sub
Sub sw25_Hit:vpmTimer.PulseSw(25):PlaySoundAtVol SoundFX("fx_target",DOFTargets),sw25,2:Me.TimerEnabled = 1: End Sub
Sub sw26_Hit:vpmTimer.PulseSw(26):PlaySoundAtVol SoundFX("fx_target",DOFTargets),sw26,2:Me.TimerEnabled = 1: End Sub
Sub sw27_Hit:vpmTimer.PulseSw(27):PlaySoundAtVol SoundFX("fx_target",DOFTargets),sw27,2:Me.TimerEnabled = 1: End Sub
Sub sw28_Hit:vpmTimer.PulseSw(28):PlaySoundAtVol SoundFX("fx_target",DOFTargets),sw28,2:Me.TimerEnabled = 1: End Sub
Sub sw29_Hit:vpmTimer.PulseSw(29):PlaySoundAtVol SoundFX("fx_target",DOFTargets),sw29,2:Me.TimerEnabled = 1: End Sub
Sub sw30_Hit:vpmTimer.PulseSw(30):PlaySoundAtVol SoundFX("fx_target",DOFTargets),sw30,2:Me.TimerEnabled = 1: End Sub
Sub sw31_Hit:vpmTimer.PulseSw(31):PlaySoundAtVol SoundFX("fx_target",DOFTargets),sw31,2:Me.TimerEnabled = 1: End Sub
Sub sw32_Hit:   Controller.Switch(32) = 1:PlaysoundAt "fx_saucer_enter",sw32: Sw32.Enabled = 0:End Sub


Sub Sw32Trigger_Hit
	'debug.print activeball.vely
	If activeball.vely > -10 Then
		Sw32.Enabled = 1
	Else
		Sw32.Enabled = 0
	End If
End Sub


Sub sw33_Hit:vpmTimer.PulseSw(33):PlaySound SoundFX("fx_target",DOFTargets),1, Vol(ActiveBall)+2, Pan(sw33), 0,0,1, 1, AudioFade(sw33):Me.TimerEnabled = 1: End Sub
Sub sw34_Hit:vpmTimer.PulseSw(34):PlaySound SoundFX("fx_target",DOFTargets),1, Vol(ActiveBall)+2, Pan(sw34), 0,0,1, 1, AudioFade(sw34):Me.TimerEnabled = 1: End Sub
Sub sw35_Hit:vpmTimer.PulseSw(35):PlaySound SoundFX("fx_target",DOFTargets),1, Vol(ActiveBall)+2, Pan(sw35), 0,0,1, 1, AudioFade(sw35):Me.TimerEnabled = 1: End Sub
Sub sw36_Hit:vpmTimer.PulseSw(36):PlaySound SoundFX("fx_target",DOFTargets),1, Vol(ActiveBall)+2, Pan(sw36), 0,0,1, 1, AudioFade(sw36):Me.TimerEnabled = 1: End Sub
Sub sw37_Hit:vpmTimer.PulseSw(37):PlaySound SoundFX("fx_target",DOFTargets),1, Vol(ActiveBall)+2, Pan(sw37), 0,0,1, 1, AudioFade(sw37):Me.TimerEnabled = 1: End Sub
Sub sw38_Hit:vpmTimer.PulseSw(38):PlaySound SoundFX("fx_target",DOFTargets),1, Vol(ActiveBall)+2, Pan(sw38), 0,0,1, 1, AudioFade(sw38):Me.TimerEnabled = 1: End Sub
'sub sw39_hit:   vpmTimer.PulseSw 39: debug.print "39":end sub
Sub sw41_Hit:vpmTimer.PulseSw(41):PlaySound SoundFX("fx_target",DOFTargets),1, Vol(ActiveBall)+2, Pan(sw41), 0,0,1, 1, AudioFade(sw41):Me.TimerEnabled = 1: End Sub
Sub sw42_Hit:vpmTimer.PulseSw(42):PlaySound SoundFX("fx_target",DOFTargets),1, Vol(ActiveBall)+2, Pan(sw42), 0,0,1, 1, AudioFade(sw42):Me.TimerEnabled = 1: End Sub
Sub sw43_Hit:vpmTimer.PulseSw(43):PlaySound SoundFX("fx_target",DOFTargets),1, Vol(ActiveBall)+2, Pan(sw43), 0,0,1, 1, AudioFade(sw43):Me.TimerEnabled = 1: End Sub
sub sw46_hit:   vpmTimer.PulseSw 46: end sub
sub sw47_hit:   vpmTimer.PulseSw 47: end sub
sub sw48_hit:   vpmTimer.PulseSw 48: end sub
Sub sw49_Hit:vpmTimer.PulseSw(49):PlaySoundAtVol SoundFX("fx_target",DOFTargets),sw49,2:Me.TimerEnabled = 1: End Sub
Sub sw50_Hit:vpmTimer.PulseSw(50):PlaySoundAtVol SoundFX("fx_target",DOFTargets),sw50,2:Me.TimerEnabled = 1: End Sub
Sub sw51_Hit:   Controller.Switch(51) = 1:PlaysoundAt "fx_saucer_enter",sw51: Sw51.Enabled = 0:End Sub

Sub Sw51Trigger_Hit
	'debug.print activeball.vely
	If activeball.vely > -10 Then
		Sw51.Enabled = 1
	Else
		Sw51.Enabled = 0
	End If
End Sub

sub sw59_hit:   controller.switch(59)=1: playsoundat "fx_rollover",ActiveBall: end sub
sub sw59_unhit: controller.switch(59)=0: end sub

'giambient.x = 476
'giambient.y = 1073.311

'***************************************************
'GI collection Lamp/Flasher sorting and GiOFF scaling init
'***************************************************

redim GILamps(99) : redim GIFlashers(99)	'Splits GI collection into these two new arrays

SortGI GILamps, GIFlashers, GI

dim TestString, TestStringAll 	'debug strings
Sub SortGI(ByRef aLight,aFlasher, GImixed) 'different method using Arrays instead of scripting dictionary objects
	dim x, CountMe: CountMe = 0
	for x = 0 to (GImixed.Count-1)
		if TypeName(GImixed(x) ) = "Light" Then
			Set aLight(CountMe) = GImixed(x)
			TestString = TestString & "assigned " & GImixed(x).Name & " to aLight(" & CountMe & ")"	& vbnewline 'debug
			CountMe = CountMe+1
			redim Preserve aLight(CountMe)
		end if
	Next
	CountMe = 0
	for x = 0 to (GImixed.Count-1)	'(note: this sub assumes there ARE flashers in the collection!)
		if TypeName(GImixed(x) ) = "Flasher" Then
			Set aFlasher(CountMe) = GImixed(x)
			TestString = TestString & "assigned " & GImixed(x).Name & " to aFlasher(" & CountMe & ")" & vbnewline	'debug
			CountMe = CountMe+1
			redim Preserve aFlasher(CountMe)
		end if
	Next
	redim Preserve aLight(uBound(aLight)-1)	'final trim of the arrays
	redim Preserve aFlasher(uBound(aFlasher)-1)
	'TestSTR(0) = TestSTR(0) & "ubound aLight: " & uBound(aLight) & " uBound aFlashers:" & uBound(aFlasher)	'debug
	'Debug.Print TestString
End Sub


'These arrays contain the following info of all non-GI lights (collected from GetElements via SortLamps sub)
Redim LightsA(999)' Object references
Redim LightsB(999)'	Opacity / Intensity
Redim LightsC(999)'	Fade Up (Light objects)
Redim LightsD(999)'	Fade Down(Light Objects)

SortLamps GI, Display
Sub SortLamps(ByVal aGI, aExclude)	'Sorts remaining light and flashers objects (EXCLUDES those in the GI collection)
	dim Counter,x,xx,skipme : skipme = False:Counter = 0 : TestStringAll = "Test String 2"
	for each x in GetElements	'now we're cooking
		'if TypeName(x) = "IDecal" then Continue For 'Decals don't have names. Evil imo D:
		if TypeName(x) = "Light" or TypeName(x) = "Flasher" Then
			SkipMe = False
			for each xx in aGI 'Find duplicates and Skip them
				if x.Name = xx.Name then
					TestStringAll = TestStringAll & x.Name & "found in GI collection, Disregarding & Continuing..." & vbnewline 'debug
					SkipMe = True'Continue For
				End If
			next
			for each xx in aExclude 'Exclude collection
				if x.Name = xx.Name then
					TestStringAll = TestStringAll & x.Name & "found in exclude collection, Disregarding & Continuing..." & vbnewline 'debug
					SkipMe = True'Continue For
				End If
			next
			If Not SkipMe Then
				On Error Resume Next
				'LightsA(Counter) = x.name	'name
				Set LightsA(Counter) = x	'ref
				LightsB(Counter) = x.Opacity
				LightsB(Counter) = x.Intensity
				LightsC(Counter) = x.FadeSpeedUp
				LightsD(Counter) = x.FadeSpeedDown
				On Error Goto 0
				Counter = Counter + 1
				redim Preserve LightsA(Counter)
				redim Preserve LightsB(Counter)
				redim Preserve LightsC(Counter)
				redim Preserve LightsD(Counter)
			End If
		End If
	next
	redim Preserve LightsA(uBound(LightsA)-1)	'final trim of the arrays
	redim Preserve LightsB(uBound(LightsB)-1)
	redim Preserve LightsC(uBound(LightsC)-1)
	redim Preserve LightsD(uBound(LightsD)-1)

	TestStringAll = TestStringAll & "Ubound LightsA = " & UBound(LightsA)	'Debug
	'debug.print TestTwo
End Sub


'Lamps NF

'Image swap arrays...
Dim BumperCapNormal: BumperCapNormal = Array("bumpertopunlit", "bumpertop33", "bumpertop66", "bumpertoplit")
Dim BumperCapIceBlue: BumperCapIceBlue = Array("bumpertopunlit", "BumperCapBlueIce_33", "BumperCapBlueIce_66", "BumperCapBlueIceLit")
Dim BumperCapPurple: BumperCapPurple = Array("bumpertopunlit", "BumperCapPurple_33", "BumperCapPurple_66", "BumperCapPurpleLit")
Dim BumperCapBlue: BumperCapBlue = Array("bumpertopunlit", "BumperCapBlue_33", "BumperCapBlue_66", "BumperCapBlueLit")
Dim BumperCapGreen: BumperCapGreen = Array("bumpertopunlit", "bumpertopGreen_33", "bumpertopGreen_66", "bumpertopGreenLit")
Dim BumperCapRed: BumperCapRed = Array("bumpertopunlit", "bumpertopRed_33", "bumpertopRed_66", "bumpertopRedLit")
Dim TextureArray1: TextureArray1 = Array("bumpertoplit", "bumpertopunlit")

Dim LeftRampArray: LeftRampArray = Array("LeftWireRamp_textureOff", "LeftWireRamp_texture33", "LeftWireRamp_texture33", "LeftWireRamp_texture33")
Dim RightRampArray: RightRampArray = Array("RightWireRamp_textureOff", "RightWireRamp_texture33", "RightWireRamp_texture66", "RightWireRamp_textureOn")

Dim GreenDomeArray: GreenDomeArray = Array("DomeGreenOff", "DomeGreenOn", "DomeGreenOn", "DomeGreenOn")
Dim RedDomeArray: RedDomeArray = Array("DomeRedOff", "DomeRedOn", "DomeRedOn", "DomeRedOn")
Dim YellowDomeArray: YellowDomeArray = Array("DomeYellowOff", "DomeYellow33", "DomeYellow66", "DomeYellow100")


Dim NullFader : set NullFader = new NullFadingObject
Dim Lampz : Set Lampz = New LampFader


InitLamps
LampTimer.Interval = 1
LampTimer.Enabled = 1

' Lamp & Flasher Updates
' LampFader object (Lampz) updates on two timers: Logic on 1, Game updates on -1
Sub LampTimer_Timer()
	dim x, chglamp
	chglamp = Controller.ChangedLamps
	If Not IsEmpty(chglamp) Then
		For x = 0 To UBound(chglamp) 			'nmbr = chglamp(x, 0), state = chglamp(x, 1)
			Lampz.state(chglamp(x, 0)) = chglamp(x, 1)
		next
	End If
	Lampz.Update1	'update (fading logic only)
End Sub

'dim FrameTime, InitFrameTime : InitFrameTime = 0
Wall78.TimerInterval = -1
Wall78.TimerEnabled = True
Sub Wall78_Timer()	'Stealing this random wall's timer for -1 updates
	'FrameTime = gametime - InitFrameTime : InitFrameTime = gametime	'Count frametime. Unused atm?
	Lampz.Update 'updates on frametime (Object updates only)
End Sub


Sub InitLamps()
	'Filtering (comment out to disable)
	Lampz.Filter = "LampFilter"	'Puts all lamp intensityscale output (no callbacks) through this function before updating

	'Adjust fading speeds (1 / full MS fading time)
	dim x : for x = 0 to 140 : Lampz.FadeSpeedUp(x) = 1/80 : Lampz.FadeSpeedDown(x) = 1/100 : next

	Lampz.FadeSpeedUp(110) = 1/64 'GI

	'Lamp Assignments
	Set Lampz.obj(1) = F1
	Set Lampz.obj(2) = F2
	Set Lampz.obj(3) = F3
	Set Lampz.obj(4) = F4
	Set Lampz.obj(5) = F5
	Set Lampz.obj(6) = L6
	Set Lampz.obj(7) = L7
	Set Lampz.obj(8) = L8
	Set Lampz.obj(9) = F9
	Set Lampz.obj(10) = F10
	Set Lampz.obj(11) = F11
	Set Lampz.obj(12) = F12
	Set Lampz.obj(13) = L13
	Lampz.Callback(13) = "PlayMusic"

	Set Lampz.obj(14) = L14
	Set Lampz.obj(15) = L15
	Set Lampz.obj(16) = L16

	Lampz.obj(17) = array(RayLight17a, RayLight17a1, RayLight17c, RayLight17c1, Raylight17w)
	Lampz.obj(18) = array(RayLight18a, RayLight18a1, RayLight18c, RayLight18c1)
	Lampz.obj(19) = array(RayLight19a, RayLight19a1, RayLight19c, RayLight19c1, Raylight19w)
	Lampz.obj(20) = array(RayLight20a, RayLight20a1, RayLight20c, RayLight20c1)
	Lampz.obj(21) = array(RayLight21a, RayLight21a1, RayLight21c, RayLight21c1, Raylight21w)

	Set Lampz.Obj(22) = l22
	Set Lampz.Obj(23) = l23
	Set Lampz.Obj(24) = l24
	Set Lampz.Obj(25) = l25
	Set Lampz.Obj(26) = l26
	Set Lampz.Obj(27) = l27
	Set Lampz.Obj(28) = l28
	Set Lampz.Obj(29) = l29
	Set Lampz.Obj(30) = l30
	Set Lampz.Obj(31) = l31
	Set Lampz.Obj(32) = l32

	Lampz.Obj(33) = Array(L133, L133a, L133b, MMBloom_Red)
	Lampz.Callback(33) = "FadePrim4 RedDome, RedDomeArray, "	'Fadepri4m RedDome, RedDomeArray

	Lampz.Obj(34) = Array(L134, L134a, MMBloom_Green)
	Lampz.Callback(34) = "FadePrim4 GreenDome, GreenDomeArray, "	'FadePri4m 34, GreenDome, GreenDomeArray

	Lampz.Obj(35) = Array(L135, L135a, L135b, MMBloom_Yellow)
	Lampz.Callback(35) = "UpdateYellowDome"	'Prim_WRampR1, RightRampArray AND 'YellowDome, YellowDomeArray

	Lampz.Obj(36) = Array(L36, F36, F36a)
	'NFadeObjmGreenBulbSwap 36, Magnet_bulb, l36, "none", "none"
	Lampz.obj(37) = Array(L37, L37a, F37, F37a)
	'NFadeObjmRedBulbSwap 37, Heart_Bulb, l37, "none", "none"
	Lampz.Obj(38) = Array(L38, F38)
	'NFadeObjmYellowBulbSwap 38, GOG_Bulb, l38, "none", "none"

	Set Lampz.Obj(39) = l39
	Lampz.Obj(40) = array(l40a, l40b)
	Set Lampz.Obj(41) = l41
	Set Lampz.Obj(42) = l42
	Set Lampz.Obj(43) = l43
	Set Lampz.Obj(44) = l44
	Set Lampz.Obj(45) = l45
	Set Lampz.Obj(46) = l46
	Set Lampz.Obj(47) = l47
	Set Lampz.Obj(48) = l48
	Set Lampz.Obj(49) = l49
	Set Lampz.Obj(50) = l50
	Set Lampz.Obj(51) = l51
	Set Lampz.Obj(52) = l52
	Set Lampz.Obj(53) = l53
	Set Lampz.Obj(54) = l54
	Set Lampz.Obj(55) = l55
	Set Lampz.Obj(56) = l56
	Lampz.Obj(57) = Array(L57, F57, F57a)
	'NFadeObjmYellowBulbSwap 57, Jackpot_bulb, l57, "none", "none"

	Set Lampz.Obj(58) = l58
	Set Lampz.Obj(59) = l59
	Set Lampz.Obj(60) = l60
	Set Lampz.Obj(61) = l61
	Set Lampz.Obj(62) = l62
	Set Lampz.Obj(63) = l63
	Set Lampz.Obj(64) = l64


	'GI Assignments
	Lampz.Obj(110) = ColToArray(GI)
	lampz.Callback(110) = "GIupdates"

	'Flasher Assignments
	Lampz.Obj(115) = Array(f101) 'Big Guy Flasher

	Lampz.obj(125) = Array(F33, MMBloom, flasher2, Flasher4)	'Mixer Heart Flasher
	Lampz.obj(126) = Array(F35, F35b, flasher3, Flasher5)		'Mixer Gab Flasher
	Lampz.obj(127) = Array(F34, F34a, flasher6, Flasher7)		'Mixer Magnet Flasher
	Lampz.obj(128) = Array(L128)					'Magnet Flasher

	Lampz.obj(129) = Array(l29b, l29t, GOGBloom)	'Gab Flasher
	'NFadeObjmBulbSwap 129, GOGModel, F129, "GOGTexOn", "GOGTexOff", "GOGmodel"

	Lampz.obj(130) = Array(F130)					'Heart Flasher
	Lampz.obj(131) = Array(L131, L131a)					'Drop Targets Flasher
	Lampz.obj(132) = Array(Ray132B, Ray132C, RayBloom, RayBloom1, F200, F200a, F200b, F200c)	'Raygun Flasher
	Lampz.Callback(132) = "FadePrim4 Prim_WRampL, LeftRampArray, "

	'Turn on GI to Start
	Lampz.state(110) = 1

	'Turn off all lamps on startup
	lampz.TurnOnStates	'Set any lamps state to 1. (Object handles fading!)
	lampz.update

End Sub

'Lamp Filter
Function LampFilter(aLvl)
	LampFilter = aLvl^1.6	'exponential curve?
End Function

'Callback procedures - these should only call when they have to

Sub PlayMusic(aLvl)
    'Intro Music
	If aLvl > 0 and PrevGameOver = 0 Then
		If IntroMusic = 0 Then
			PlaySound "intro"
			PrevGameOver = 1
		End If
	else
'		PrevGameOver = 0
	End If
End Sub

Sub UpdateYellowDome(ByVal aLvl)	'Special for yellow dome, double primitive image swap
	FadePrim4 YellowDome, YellowDomeArray, aLvl
	FadePrim4 Prim_WRampR1, RightRampArray, aLvl
End Sub

Sub FadePrim4(pri, group, ByVal aLvl)	'cp's script
	if Lampz.UseFunction then aLvl = LampFilter(aLvl)	'Callbacks don't get this filter automatically
    Select case FlashLevelToIndex(aLvl, 4)
		Case 1:pri.image = group(0) 'Off
		Case 2:pri.image = group(1) 'Fading...
		Case 3:pri.image = group(2) 'Fading...
        Case 4:pri.image = group(3) 'ON
    End Select
End Sub

function FlashLevelToIndex(Input, MaxSize)
	FlashLevelToIndex = cInt(Input * (MaxSize-1)+.5)+1
end function

Sub FadeDisableLighting1(aObject, ByVal aLvl)
	if Lampz.UseFunction then aLvl = LampFilter(aLvl)	'Callbacks don't get this filter automatically
	aObject.BlendDisableLighting = aLvl * 0.2
End Sub


Sub FadeMaterialP(itemp, group, ByVal aLvl)	'cp's script
	Select Case aLvl
		case 0 : itemp.Material = group(1)
		case 1 : itemp.Material = group(0)
	end select
End Sub

Dim GIoffMult : GIoffMult = 3 'Multiplies all non-GI opacity when the GI is off
Sub GIupdates(ByVal aLvl)
	if Lampz.UseFunction then aLvl = LampFilter(aLvl)	'Callbacks don't get this filter automatically
	dim a : a = Array(Bumper1cap, Bumper2cap, Bumper3Cap)
	dim idx : for idx = 0 to uBound(a)
		FadeDisableLighting1 a(idx), aLvl
		FadeMaterialP a(idx), TextureArray1, aLvl
	Next

'	dim bmprcolor : bmprcolor = Array(BumperCapNormal, BumperCapBlue, BumperCapPurple, BumperCapIceBlue)	'0->3)
'	if BumperColorType <= 3 then
'		for idx = 0 to uBound(a)
'			FadePrim4 a(idx), bmprcolor(BumperColorType), aLvl
'		next
'	Else
'		Select Case BumperColorType
'			Case 4
'				FadePrim4 Bumper1Cap, BumperCapIceBlue, aLvl
'				FadePrim4 Bumper2Cap, BumperCapPurple, aLvl
'				FadePrim4 Bumper3Cap, BumperCapBlue, aLvl
'			Case 6
'				FadePrim4 Bumper1Cap, BumperCapIceBlue, aLvl
'				FadePrim4 Bumper2Cap, BumperCapRed, aLvl
'				FadePrim4 Bumper3Cap, BumperCapGreen, aLvl
'		End Select
'	end If

	''Lut fading
	'dim LutName, LutCount, GoLut
	'''FadeLUT 100, "ColorGradeBOP_", 7
	'LutName = "ColorGradeBOP_"
	'LutCount = 7
	'GoLut = cInt(LutCount * aLvl	)+1	'+1 because no 0 with these luts
	'GoLut = LutName & GoLut
	'if Table1.ColorGradeImage <> GoLut then Table1.ColorGradeImage = GoLut : 	tb.text = golut


	'Fade lamps up when GI is off
	dim GIscale
	GiScale = (GIoffMult-1) * (ABS(aLvl-1 )  ) + 1	'invert
	dim x : for x = 0 to uBound(LightsA)
		On Error Resume Next
		LightsA(x).Opacity = LightsB(x) * GIscale
		LightsA(x).Intensity = LightsB(x) * GIscale
		'LightsA(x).FadeSpeedUp = LightsC(x) * GIscale
		'LightsA(x).FadeSpeedDown = LightsD(x) * GIscale
		On Error Goto 0
	Next


End Sub


'Helper function
Function ColtoArray(aDict)	'converts a collection to an indexed array. Indexes will come out random probably.
	redim a(999)
	dim count : count = 0
	dim x  : for each x in aDict : set a(Count) = x : count = count + 1 : Next
	redim preserve a(count-1) : ColtoArray = a
End Function

Sub GIOn  : SetGI False: End Sub 'These are just debug commands now
Sub GIOff : SetGI True : End Sub



'*******************************
'Intermediate Solenoid Procedures (Setlamp, etc)
'********************************
'Solenoid pipeline looks like this:
'Pinmame Controller -> UseSolenoids -> Solcallback -> intermediate subs (here) -> Lampz fading object -> object updates / more callbacks

'Lamps, for reference:
'Pinmame Controller -> UpdateLamps sub -> Lampz Fading Object -> Object Updates / callbacks

Sub SetLamp(aNr, aOn)
	Lampz.state(aNr) = abs(aOn)
End Sub

Sub SetGI(aOFF)	'Inverted, Solenoid cuts GI circuit on this era of game
	select case aOFF
		Case True  	 'GI off
			PlaysoundAtVol "fx_relay_off",Light66,2
			SetLamp 110, 0
		Case False 	   'GI on
			PlaysoundAtVol "fx_relay_on",Light66,5
			SetLamp 110, 1
	End Select
End Sub


' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed, decrease the default 2000 to hear a louder rolling ball sound
    Vol = Csng(BallVel(ball) ^2 / 1000)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / table1.width-1
    If tmp> 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
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


Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

'*****************************************
'    JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 3 ' total number of balls
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub RollingSoundsUpdate()
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
						PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*.8, Pan(BOT(b) ), 0, Pitch(BOT(b) )+30000, 1, 0, AudioFade(BOT(b) )
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
	PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 100, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

'**************************************************************************
'                 Positional Sound Playback Functions by DJRobX
'**************************************************************************

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
		PlaySound sound, 1, Vol, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub


'Set all as per ball position & speed, but Vol Multiplier may be used eg; PlaySoundAtBallVol "sound",3

Sub PlaySoundAtBallVol(sound, VolMult)
		PlaySound sound, 0, Vol(ActiveBall) * VolMult, Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub


'Set position as bumperX and Vol manually. Allows rapid repetition/overlaying sound

Sub PlaySoundAtBumperVol(sound, tableobj, Vol)
		PlaySound sound, 1, Vol, Pan(tableobj), 0,0,1, 1, AudioFade(tableobj)
End Sub

'*****************************************
'    JimmyFingers VP10 Sound Routines
'*****************************************


Sub BallHitSound(dummy):PlaySound "fx_balldrop",0,2,Pan(RWireEnd),0,0,0,0,AudioFade(RWireEnd):End Sub
Sub BallHitSound2(dummy):PlaySound "fx_balldrop",0,2,Pan(LWireEnd),0,0,0,0,AudioFade(LWireEnd):End Sub
Sub BallHitSound3(dummy):PlaySound "fx_balldrop",0,2,Pan(MMdrop),0,0,0,0,AudioFade(MMdrop):End Sub

Sub MMdrop_Hit
		vpmtimer.addtimer 150, "BallHitSound3"
End Sub

Sub LWireEnd_Hit
		vpmtimer.addtimer 150,"BallHitSound2"
End Sub

Sub RWireEnd_Hit
		vpmtimer.addtimer 150,"BallHitSound"
End Sub


Sub Wall33_Hit: PlaySoundAtBallVol "metalguidebump2",2:End Sub
Sub Wall34_Hit: PlaySoundAtBallVol "metalguidebump2",2:End Sub


'***Random Rubber Hit Sounds***
Sub RandomSoundRubber()
    Select Case Int(Rnd * 3) + 1
        Case 1:PlaySound "fx_rubber_hit_1", 0, 20*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
        Case 2:PlaySound "fx_rubber_hit_2", 0, 20*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
        Case 3:PlaySound "fx_rubber_hit_3", 0, 20*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    End Select
End Sub


Sub RubbersBandsLargeRings_Hit(idx)
	If Table1.VersionMinor > 3 OR Table1.VersionMajor > 10 Then
		RandomSoundRubber()
	Else
		PlaySound "fx_rubber_hit_3", 0, 3*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1
    End If
End Sub


Sub RubbersSmallRings_Hit(idx)
	If Table1.VersionMinor > 3 OR Table1.VersionMajor > 10 Then
		RandomSoundRubber()
	Else
		PlaySound "fx_rubber_hit_3", 0, 3*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1
    End If
End Sub



Sub RubbersWalls_Hit(idx)
	If Table1.VersionMinor > 3 OR Table1.VersionMajor > 10 Then
		RandomSoundRubber()
	Else
		PlaySound "fx_rubber_hit_3", 0, 3*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1
    End If
End Sub


Sub RubbersMM_Hit(idx)
	If Table1.VersionMinor > 3 OR Table1.VersionMajor > 10 Then
		RandomSoundRubber()
	Else
		PlaySound "fx_rubber_hit_3", 0, 3*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1
    End If
End Sub


'***Random Flipper Hit Sounds***
Sub LeftFlipper_Collide(parm)
 	RandomSoundFlipper()
End Sub

Sub RightFlipper_Collide(parm)
 	RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "fx_flip_hit1", 1, 10*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 2 : PlaySound "fx_flip_hit2", 1, 10*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 3 : PlaySound "fx_flip_hit3", 1, 10*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End Select
End Sub

'***************** Random Real Time Sounds ****************************


Dim NextOrbitHit:NextOrbitHit = 0

Sub WireRampBumps_Hit(idx)
	if BallVel(ActiveBall) > .3 and Timer > NextOrbitHit then
		RandomBump3 1, Pitch(ActiveBall)+5
		' Schedule the next possible sound time.  This prevents it from rapid-firing noises too much.
		' Lowering these numbers allow more closely-spaced clunks.
		NextOrbitHit = Timer + .4 + (Rnd * .2)
	end if
End Sub


Sub PlasticRampBumps_Hit(idx)
	if BallVel(ActiveBall) > .3 and Timer > NextOrbitHit then
		RandomBump 2, Pitch(ActiveBall)
		' Schedule the next possible sound time.  This prevents it from rapid-firing noises too much.
		' Lowering these numbers allow more closely-spaced clunks.
		NextOrbitHit = Timer + .05 + (Rnd * .2)
	end if
End Sub


Sub MetalWallBumps_Hit(idx)
	if BallVel(ActiveBall) > .3 and Timer > NextOrbitHit then
		RandomBump 3, 20000 'Increased pitch to simulate metal wall
		' Schedule the next possible sound time.  This prevents it from rapid-firing noises too much.
		' Lowering these numbers allow more closely-spaced clunks.
		NextOrbitHit = Timer + .2 + (Rnd * .2)
	end if
End Sub


' Requires rampbump1 to 7 in Sound Manager
Sub RandomBump(voladj, freq)
	dim BumpSnd:BumpSnd= "rampbump" & CStr(Int(Rnd*7)+1)
		PlaySound BumpSnd, 0, Vol(ActiveBall)*voladj, Pan(ActiveBall), 0, freq, 0, 1, AudioFade(ActiveBall)
End Sub


' Requires metalguidebump1 to 2 in Sound Manager
Sub RandomBump2(voladj, freq)
	dim BumpSnd:BumpSnd= "metalguidebump" & CStr(Int(Rnd*2)+1)
		PlaySound BumpSnd, 0, Vol(ActiveBall)*voladj, Pan(ActiveBall), 0, freq, 0, 1, AudioFade(ActiveBall)
End Sub


' Requires WireRampBump1 to 5 in Sound Manager
Sub RandomBump3(voladj, freq)
	dim BumpSnd:BumpSnd= "WireRampBump" & CStr(Int(Rnd*5)+1)
		PlaySound BumpSnd, 0, Vol(ActiveBall)*voladj, Pan(ActiveBall), 0, freq, 0, 1, AudioFade(ActiveBall)
End Sub


' Stop Bump Sounds
Sub BumpSTOPplastic_Hit ()
dim i:for i=1 to 4:StopSound "RampBump" & i:next
NextOrbitHit = Timer + 1
End Sub

Sub BumpSTOPwireR_Hit ()
dim i:for i=1 to 4:StopSound "WireRampBump" & i:next
NextOrbitHit = Timer + 1
End Sub

Sub BumpSTOPwireL_Hit ()
dim i:for i=1 to 4:StopSound "WireRampBump" & i:next
NextOrbitHit = Timer + 1
End Sub

'Generic ramp sounds
Sub Wall78_Hit:Playsound "FX_metalhit",0,Vol(ActiveBall),Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub
Sub Leftlane1_Hit:Playsound "FX_metalhit",0,Vol(ActiveBall),Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub
Sub Leftlane2_Hit:Playsound "FX_metalhit",0,Vol(ActiveBall),Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub
Sub Wall69_Hit:Playsound "FX_metalhit",0,Vol(ActiveBall),Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub


'****Prim Guy Shake
Const PrimGuyShakeXMax=3
Const PrimGuyShakeXStep=.3
PrimGuyShakeTimer.Interval = 40
Dim PrimGuyShakeDirection, PrimGuyShakeXOffset

Sub PrimGuyShake
	PrimGuyShakeXOffset = PrimGuyShakeXMax
	PrimGuy.ObjRotX = PrimGuyShakeXOffset
	PrimGuyShakeDirection = 1
	PrimGuyShakeTimer.Enabled = 1
End Sub

sub PrimGuyShakeTimer_Timer
	PrimGuyShakeDirection = -1*PrimGuyShakeDirection	'Change Direction
	PrimGuyShakeXOffset = PrimGuyShakeXOffset - PrimGuyShakeXStep		'Calc New Offset
	If PrimGuyShakeXOffset > 0 Then 'Keep Shaking
		PrimGuy.ObjRotX = PrimGuyShakeDirection * PrimGuyShakeXOffset
	Else	'Time to stop shaking
		PrimGuy.ObjRotX  = 0
		PrimGuyShakeTimer.Enabled = 0
	End If
End Sub


'**********************************************************************************************************
'**********************************************************************************************************
'***Options***
'**********************************************************************************************************
'**********************************************************************************************************
Dim BumperColor, Prevgameover, IntroMusic, xxGIColor, aproncolor, flipperstyle
Dim cheaterpost, raybeam, MMflashers, GIColorModType, GIColorMod, BumperColorType, Rails

Sub SetOptions()

If GIColorMod = 4 Then
	GIColorModType = Int(Rnd*4)
Else
	GIColorModType = GIColorMod
End If

If GIColorModType = 0 then
	for each xxGIColor in GIMain
		xxGIColor.Color=White
		xxGIColor.ColorFull=WhiteFull
		xxGIColor.Intensity = WhiteI
		next
	for each xxGIColor in GIPlastic
		xxGIColor.Color=WhitePlastic
		xxGIColor.ColorFull=WhitePlasticFull
		xxGIColor.Intensity = WhitePlasticI
		next
    for each xxGIColor in GIBulbs
		xxGIColor.Color=WhiteBulbs
		xxGIColor.ColorFull=WhiteBulbsFull
		xxGIColor.Intensity = WhiteBulbsI
		next
End If



If GIColorModType = 1 then
	for each xxGIColor in GIMain
		xxGIColor.Color=Blue
		xxGIColor.ColorFull=BlueFull
		xxGIColor.Intensity = BlueI
		next
	for each xxGIColor in GIPlastic
		xxGIColor.Color=BluePlastic
		xxGIColor.ColorFull=BluePlasticFull
		xxGIColor.Intensity = BluePlasticI
		next
    for each xxGIColor in GIBulbs
		xxGIColor.Color=BlueBulbs
		xxGIColor.ColorFull=BlueBulbsFull
		xxGIColor.Intensity = BlueBulbsI
		next
End If

If GIColorModType = 2 then
	for each xxGIColor in GIMain
		xxGIColor.Color=Purple
		xxGIColor.ColorFull=PurpleFull
		xxGIColor.Intensity = PurpleI
		next
	for each xxGIColor in GIPlastic
		xxGIColor.Color=PurplePlastic
		xxGIColor.ColorFull=PurplePlasticFull
		xxGIColor.Intensity = PurplePlasticI
		next
    for each xxGIColor in GIBulbs
		xxGIColor.Color=PurpleBulbs
		xxGIColor.ColorFull=PurpleBulbsFull
		xxGIColor.Intensity = PurpleBulbsI
		next
End If

If GIColorModType = 3 then
	for each xxGIColor in GIMain
		xxGIColor.Color=IceBlue
		xxGIColor.ColorFull=IceBlueFull
		xxGIColor.Intensity = IceBlueI
		next
	for each xxGIColor in GIPlastic
		xxGIColor.Color=IceBluePlastic
		xxGIColor.ColorFull=IceBluePlasticFull
		xxGIColor.Intensity = IceBluePlasticI
		next
    for each xxGIColor in GIBulbs
		xxGIColor.Color=IceBlueBulbs
		xxGIColor.ColorFull=IceBlueBulbsFull
		xxGIColor.Intensity = IceBlueBulbsI
		next

	End If

	If Aproncolor = 4 then Aproncolor = Int(Rnd*4) End If
	Select Case Aproncolor
		Case 0 :
			pApron.image = "apron_texture_White":
			pApronOverlay.Visible = False:
			pCustomWall.visible = false:
			'pSidewall_DT.visible = false
		Case 1 :
			pApron.image = "apron_texture_Blue":
			pApronOverlay.Visible = False:
			pCustomWall.visible = false:
			'pSidewall_DT.visible = false
		Case 2 :
			pApron.image = "apron_texture_Green":
			pApronOverlay.Visible = False:
			pCustomWall.visible = false:
			'pSidewall_DT.visible = false
	    Case 3 :
			pApronOverlay.Visible = True:
			If DesktopMode = True Then
				'pSidewall_DT.visible = True
				pCustomWall.visible = true
			else
				'pSidewall_DT.visible = true
				pCustomWall.visible = true
			End If
	End Select

	if flipperstyle = 2 then flipperstyle = int(rnd*2) end if
	select case flipperstyle
		case 0: flipperl.visible=false: flipperr.visible=False: RightFlipper.visible=True: LeftFlipper.visible=true
		case 1: flipperl.visible=true: flipperr.visible=True: RightFlipper.visible=False: LeftFlipper.visible=false
	end select


    If cheaterpost = 1 then
		cpost.visible = 1
		cRubberRubber.collidable = 1
		cRubberRubber.visible = 1
	Else
		cpost.visible = 0
		cRubberRubber.collidable = 0
		cRubberRubber.visible = 0
	End If

	If Rails =1 Then
		Leftrail.visible = 0
		Rightrail.visible = 0
	Else
		Leftrail.visible = 1
		Rightrail.visible = 1
	End if

	If Raybeam = 1 then
	     f200.visible = 1
	     f200a.visible = 1
	Else
	     f200.visible = 0
	     f200a.visible = 0
	End If

If BumperColor = 4 Then
	BumperColorType = Int(Rnd*4)
Else
	BumperColorType = BumperColor
End If



If BumperColorType = 0 then
	for each xxGIColor in GIBumper
		xxGIColor.Color=WhiteBumper
		xxGIColor.ColorFull=WhiteBumperFull
		xxGIColor.Intensity = WhiteBumperI
		next
	for each xxGIColor in SkillShotLights
		xxGIColor.Color= rgb(255,255,180)
		xxGIColor.opacity= 400
	next
	popbumperbloom.intensity = 3
	Bumper1Cap.image = "BumperTopLit"
	Bumper1Cap.BlendDisableLighting = 0.05
	Bumper1Cap.material = "bumpertopunlit"
	Bumper2Cap.image = "BumperTopLit"
	Bumper2Cap.BlendDisableLighting = 0.05
	Bumper2Cap.material = "bumpertopunlit"
	Bumper3Cap.image = "BumperTopLit"
	Bumper3Cap.BlendDisableLighting = 0.05
	Bumper3Cap.material = "bumpertopunlit"
End If


If BumperColorType = 1 then
	for each xxGIColor in GIBumper
		xxGIColor.Color=BlueBumper
		xxGIColor.ColorFull=BlueBumperFull
		xxGIColor.Intensity = BlueBumperI
	next
	for each xxGIColor in SkillShotLights
		xxGIColor.Color= rgb(255,255,150)
		xxGIColor.opacity= 900
	next
	popbumperbloom.intensity = 7
	Bumper1Cap.image = "BumperCapBlueLit"
	Bumper1Cap.BlendDisableLighting = 0.05
	Bumper1Cap.material = "bumpertopunlit"
	Bumper2Cap.image = "BumperCapBlueLit"
	Bumper2Cap.BlendDisableLighting = 0.05
	Bumper2Cap.material = "bumpertopunlit"
	Bumper3Cap.image = "BumperCapBlueLit"
	Bumper3Cap.BlendDisableLighting = 0.05
	Bumper3Cap.material = "bumpertopunlit"
End If



If BumperColorType = 2 then
	for each xxGIColor in GIBumper
		xxGIColor.Color=PurpleBumper
		xxGIColor.ColorFull=PurpleBumperFull
		xxGIColor.Intensity = PurpleBumperI
		next
	for each xxGIColor in SkillShotLights
		xxGIColor.Color= rgb(255,255,150)
		xxGIColor.opacity= 400
	next
	popbumperbloom.intensity = 7
	Bumper1Cap.image = "BumperCapPurpleLit"
	Bumper1Cap.BlendDisableLighting = 0.05
	Bumper1Cap.material = "bumpertopunlit"
	Bumper2Cap.image = "BumperCapPurpleLit"
	Bumper2Cap.BlendDisableLighting = 0.05
	Bumper2Cap.material = "bumpertopunlit"
	Bumper3Cap.image = "BumperCapPurpleLit"
	Bumper3Cap.BlendDisableLighting = 0.05
	Bumper3Cap.material = "bumpertopunlit"
End If
If BumperColorType = 3 then
	for each xxGIColor in GIBumper
		xxGIColor.Color=IceBlueBumper
		xxGIColor.ColorFull=IceBlueBumperFull
		xxGIColor.Intensity = IceBlueBumperI
		next
	for each xxGIColor in SkillShotLights
		xxGIColor.Color= rgb(255,255,150)
		xxGIColor.opacity= 500
	next
	popbumperbloom.intensity = 4
	Bumper1Cap.image = "BumperCapBlueIceLit"
	Bumper1Cap.BlendDisableLighting = 0.05
	Bumper1Cap.material = "bumpertopunlit"
	Bumper2Cap.image = "BumperCapBlueIceLit"
	Bumper2Cap.BlendDisableLighting = 0.05
	Bumper2Cap.material = "bumpertopunlit"
	Bumper3Cap.image = "BumperCapBlueIceLit"
	Bumper3Cap.BlendDisableLighting = 0.05
	Bumper3Cap.material = "bumpertopunlit"
End If

End Sub

'*****************************************************************************************************************************************************
'*****************************************************************************************************************************************************



'**********************************************************************************************************
'Digital Display
'**********************************************************************************************************

Dim Digits(32)
 Digits(0)=Array(a00, a05, a0c, a0d, a08, a01, a06, a0f, a02, a03, a04, a07, a0b, a0a, a09, a0e)
 Digits(1)=Array(a10, a15, a1c, a1d, a18, a11, a16, a1f, a12, a13, a14, a17, a1b, a1a, a19, a1e)
 Digits(2)=Array(a20, a25, a2c, a2d, a28, a21, a26, a2f, a22, a23, a24, a27, a2b, a2a, a29, a2e)
 Digits(3)=Array(a30, a35, a3c, a3d, a38, a31, a36, a3f, a32, a33, a34, a37, a3b, a3a, a39, a3e)
 Digits(4)=Array(a40, a45, a4c, a4d, a48, a41, a46, a4f, a42, a43, a44, a47, a4b, a4a, a49, a4e)
 Digits(5)=Array(a50, a55, a5c, a5d, a58, a51, a56, a5f, a52, a53, a54, a57, a5b, a5a, a59, a5e)
 Digits(6)=Array(a60, a65, a6c, a6d, a68, a61, a66, a6f, a62, a63, a64, a67, a6b, a6a, a69, a6e)
 Digits(7)=Array(a70, a75, a7c, a7d, a78, a71, a76, a7f, a72, a73, a74, a77, a7b, a7a, a79, a7e)
 Digits(8)=Array(a80, a85, a8c, a8d, a88, a81, a86, a8f, a82, a83, a84, a87, a8b, a8a, a89, a8e)
 Digits(9)=Array(a90, a95, a9c, a9d, a98, a91, a96, a9f, a92, a93, a94, a97, a9b, a9a, a99, a9e)
 Digits(10)=Array(aa0, aa5, aac, aad, aa8, aa1, aa6, aaf, aa2, aa3, aa4, aa7, aab, aaa, aa9, aae)
 Digits(11)=Array(ab0, ab5, abc, abd, ab8, ab1, ab6, abf, ab2, ab3, ab4, ab7, abb, aba, ab9, abe)
 Digits(12)=Array(ac0, ac5, acc, acd, ac8, ac1, ac6, acf, ac2, ac3, ac4, ac7, acb, aca, ac9, ace)
 Digits(13)=Array(ad0, ad5, adc, add, ad8, ad1, ad6, adf, ad2, ad3, ad4, ad7, adb, ada, ad9, ade)
 Digits(14)=Array(ae0, ae5, aec, aed, ae8, ae1, ae6, aef, ae2, ae3, ae4, ae7, aeb, aea, ae9, aee)
 Digits(15)=Array(af0, af5, afc, afd, af8, af1, af6, aff, af2, af3, af4, af7, afb, afa, af9, afe)
 Digits(16)=Array(b00, b05, b0c, b0d, b08, b01, b06, b0f, b02, b03, b04, b07, b0b, b0a, b09, b0e)
 Digits(17)=Array(b10, b15, b1c, b1d, b18, b11, b16, b1f, b12, b13, b14, b17, b1b, b1a, b19, b1e)
 Digits(18)=Array(b20, b25, b2c, b2d, b28, b21, b26, b2f, b22, b23, b24, b27, b2b, b2a, b29, b2e)
 Digits(19)=Array(b30, b35, b3c, b3d, b38, b31, b36, b3f, b32, b33, b34, b37, b3b, b3a, b39, b3e)
 Digits(20)=Array(b40, b45, b4c, b4d, b48, b41, b46, b4f, b42, b43, b44, b47, b4b, b4a, b49, b4e)
 Digits(21)=Array(b50, b55, b5c, b5d, b58, b51, b56, b5f, b52, b53, b54, b57, b5b, b5a, b59, b5e)
 Digits(22)=Array(b60, b65, b6c, b6d, b68, b61, b66, b6f, b62, b63, b64, b67, b6b, b6a, b69, b6e)
 Digits(23)=Array(b70, b75, b7c, b7d, b78, b71, b76, b7f, b72, b73, b74, b77, b7b, b7a, b79, b7e)
 Digits(24)=Array(b80, b85, b8c, b8d, b88, b81, b86, b8f, b82, b83, b84, b87, b8b, b8a, b89, b8e)
 Digits(25)=Array(b90, b95, b9c, b9d, b98, b91, b96, b9f, b92, b93, b94, b97, b9b, b9a, b99, b9e)
 Digits(26)=Array(ba0, ba5, bac, bad, ba8, ba1, ba6, baf, ba2, ba3, ba4, ba7, bab, baa, ba9, bae)
 Digits(27)=Array(bb0, bb5, bbc, bbd, bb8, bb1, bb6, bbf, bb2, bb3, bb4, bb7, bbb, bba, bb9, bbe)
 Digits(28)=Array(bc0, bc5, bcc, bcd, bc8, bc1, bc6, bcf, bc2, bc3, bc4, bc7, bcb, bca, bc9, bce)
 Digits(29)=Array(bd0, bd5, bdc, bdd, bd8, bd1, bd6, bdf, bd2, bd3, bd4, bd7, bdb, bda, bd9, bde)
 Digits(30)=Array(be0, be5, bec, bed, be8, be1, be6, bef, be2, be3, be4, be7, beb, bea, be9, bee)
 Digits(31)=Array(bf0, bf5, bfc, bfd, bf8, bf1, bf6, bff, bf2, bf3, bf4, bf7, bfb, bfa, bf9, bfe)

 Sub DisplayTimer_Timer
    Dim ChgLED, ii, jj, num, chg, stat, obj, b, x
    ChgLED=Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
		If DesktopMode = True Then
       For ii=0 To UBound(chgLED)
          num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)
			if (num < 32) then
              For Each obj In Digits(num)
                   If chg And 1 Then obj.State=stat And 1
                   chg=chg\2 : stat=stat\2
                  Next
			Else
			       end if
        Next
	   end if
    End If
 End Sub


'**********************************************************************************************************
'**********************************************************************************************************
'***RGB OPTIONS COLOR ADJUST***
'**********************************************************************************************************
'**********************************************************************************************************


Dim White, WhiteFull, WhiteI, WhitePlastic, WhitePlasticFull, WhitePlasticI, WhiteBumper, WhiteBumperFull, WhiteBumperI, WhiteBulbs, WhiteBulbsFull, WhiteBulbsI
WhiteFull = rgb(255,255,255)
White = rgb(255,255,180)
WhiteI = 1
WhitePlasticFull = rgb(255,255,255)
WhitePlastic = rgb(255,255,180)
WhitePlasticI = 10
WhiteBumperFull = rgb(255,255,180)
WhiteBumper = rgb(255,255,255)
WhiteBumperI = 2
WhiteBulbsFull = rgb(255,255,255)
WhiteBulbs = rgb(255,255,180)
WhiteBulbsI = 20

Dim Blue, BlueFull, BlueI, BluePlastic, BluePlasticFull, BluePlasticI, BlueBumper, BlueBumperFull, BlueBumperI, BlueBulbs, BlueBulbsFull, BlueBulbsI
BlueFull = rgb(75,75,255)
Blue = rgb(75,75,255)
BlueI = 2
BluePlasticFull = rgb(75,75,255)
BluePlastic = rgb(75,75,255)
BluePlasticI = 14
BlueBumperFull = rgb(0,0,255)
BlueBumper = rgb(0,0,255)
BlueBumperI = 1
BlueBulbsFull = rgb(75,75,255)
BlueBulbs = rgb(75,75,255)
BlueBulbsI = 20

Dim Purple, PurpleFull, PurpleI, PurplePlastic, PurplePlasticFull, PurplePlasticI, PurpleBumper, PurpleBumperFull, PurpleBumperI, PurpleBulbs, PurpleBulbsFull, PurpleBulbsI
Purple = rgb(125,60,125)
PurpleI = 5
PurplePlasticFull = rgb(125,0,125)
PurplePlastic = rgb(125,0,125)
PurplePlasticI = 20
PurpleBumperFull = rgb(125,0,125)
PurpleBumper = rgb(125,0,125)
PurpleBumperI = 3
PurpleBulbsFull = rgb(125,0,125)
PurpleBulbs = rgb(125,0,125)
PurpleBulbsI = 20

Dim IceBlue, IceBlueFull, IceBlueI, IceBluePlastic, IceBluePlasticFull, IceBluePlasticI, IceBlueBumper, IceBlueBumperFull, IceBlueBumperI, IceBlueBulbs, IceBlueBulbsFull, IceBlueBulbsI
IceBlueFull = rgb(0,255,255)
IceBlue = rgb(0,255,255)
IceBlueI = 1
IceBluePlasticFull = rgb(0,255,255)
IceBluePlastic = rgb(0,255,255)
IceBluePlasticI = 15
IceBlueBumperFull = rgb(0,255,255)
IceBlueBumper = rgb(0,255,255)
IceBlueBumperI = 1
IceBlueBulbsFull = rgb(0,255,255)
IceBlueBulbs = rgb(0,255,255)
IceBlueBulbsI = 10


'====================
'Class jungle nf
'=============

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
				If UseCallBack(x) then execute cCallback(x) & " " & (Lvl(x))	'Callback
				If Lock(x) Then
					if Lvl(x) = 1 or Lvl(x) = 0 then Loaded(x) = True	'finished fading
				end if
			end if
		Next
	End Sub
End Class

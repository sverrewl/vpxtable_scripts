' __    __  __ __  ____  ______    ___      __    __   ____  ______    ___  ____
'|  |__|  ||  |  ||    ||      |  /  _]    |  |__|  | /    ||      |  /  _]|    \ 
'|  |  |  ||  |  | |  | |      | /  [_     |  |  |  ||  o  ||      | /  [_ |  D  )
'|  |  |  ||  _  | |  | |_|  |_||    _]    |  |  |  ||     ||_|  |_||    _]|    /
'|  '  '  ||  |  | |  |   |  |  |   [_     |  '  '  ||  _  |  |  |  |   [_ |    \ 
' \      / |  |  | |  |   |  |  |     |     \      / |  |  |  |  |  |     ||  .  |
'  \_/\_/  |__|__||____|  |__|  |_____|      \_/\_/  |__|__|  |__|  |_____||__|\_|


' Williams White Water / IPD No. 2768 / January, 1993 / 4 Players
' made for VPX by Flupper
' http://www.ipdb.org/machine.cgi?id=2768
' Thanks to JPSalas and PacDude's earlier versions
' Many thanks also to:
' rock primitives & base textures, bigfoot, transparent targets by Dark
' playfield texture and plastic images by Clark Kent
' texture redraws by Lobotomy
' Physics tuning by wrd1972 and Clark Kent
' reference images of darkened table by darquayle
' many mechanical sounds by Knorr
' sound tuning by DjRobX
'
' Notes (from experience during development):
' If Bigfoot head gets out of sync or VPM crashes, delete NVRAM

' Thalamus 2018-07-24
' Table has already "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' No special SSF tweaks yet.

Option Explicit : Randomize
Dim BallShadow, ForceSiderailsFS, GlobalSoundLevel, HitTheGlass, DisableGI, DesktopVPXDMD
Dim FlasherTimerInterval, debugballs, NoSideWallRelfections

' *** Enable / Disable ball shadow ********************************************************
ballshadow = true

' *** relative sound level for all mechanical sounds **************************************
' *** 1: if you have separate sound system for mech sounds, 2/3/etc when you do not *******
GlobalSoundLevel = 2

' *** synchronize flasher frequency with screen refresh rate (ms) *************************
' *** 17 for 60 fps, 20 for 100 fps, 13 for 75 fps ****************************************
' *** you can use -1 for a fps locked 60 fps **********************************************
FlasherTimerInterval = 17

' *** Show siderails in fullscreen mode: True = show siderails, False = do not show *******
ForceSiderailsFS = False

' *** Special effect: let the ball hit the glass at the top of the waveramp ***************
HitTheGlass = True

' *** Special effect: Play with the global illumination off *******************************
DisableGI = False

' **** enables manual ball control with C key (enable/disable control) ********************
' **** and B key (speed boost) and arrow keys *********************************************
debugballs = False

' **** Use built-in VPX DMD instead of relying on VPM one *********************************
' **** looks nicer, required for exclusive fullscreen *************************************
DesktopVPXDMD = False

' *** disable side wall reflections *******************************************************
NoSideWallRelfections = False


'******************************************************************************************
'* END OF TABLE OPTIONS *******************************************************************
'******************************************************************************************

' release notes

' version 1.0 : initial version

' explanation of standard constants (info by Jpsalas and others on VpForums):
' These constants are for the vpinmame emulation, and they tell vpinmame what it is supposed
' to emulate.
' UseSolenoids=1 means the vpinmame will use the solenoids, so in the script there are calls
' 				 for a solenoid to do different things (like reset droptargets, kick a ball, etc)
' UseLamps=0 	 means the vpinmame won't bother updating the lights, but done in script
' UseSync=0      (unclear) but probably is to enable or disable the sync in the vpinmame window
' 				 (or dmd). So it syncs with the screen or not.
' HandleMech=0   means vpinmame won't handle  special animations, they will have to be done
'				 manually in Scripts
' UseGI=1        If 1 and used together with "Set GiCallback2 = GetRef("UpdateGI")" where
'				 UpdateGI is the sub routine that sets the Global Illumination lights
' 				 only the Williams wpc tables have Gi circuitry support (so you can use GICallback)
'				 for other tables a solenoid has to be used
' SFlipperOn     - Flipper activate sound
' SFlipperOff    - Flipper deactivate sound
' SSolenoidOn    - Solenoid activate sound
' SSolenoidOff   - Solenoid deactivate sound
' SCoin          - Coin Sound
' UseVPMModSol   When True this allows the ROM to control the intensity level of modulated solenoids
'				 instead of just on/off.
' UseVPMDMD      - Enable VPX rendering of DMD

' *** Global constants ***
Const UseSolenoids = 2
Const UseLamps = 0
Const UseSync = 0
Const HandleMech = 0
Const UseGI=1
Const cCredits="White Water, Williams 1993"
Const SSolenoidOn = "Solenoid"
Const SSolenoidOff = "SolenoidOff"
Const SFlipperOn = "FlipperUp"
Const SFlipperOff = "FlipperDown"
Const SCoin = "coin"
Const UseVPMModSol=1

' *** Dynamic flipper friction Mod ***
Const DynamicFlipperFriction = True
Const DynamicFlipperFrictionResting = 0.6
Const DynamicFlipperFrictionActive = 1.0

' *** Global variables ***
Dim relaylevel : relaylevel = 0.5 * GlobalSoundLevel ' Sound level of relay clicking
Dim metalvolume : metalvolume = 0 ' is zero until balls have been created
Dim FlippersEnabled	 ' Used to enable/disable flippers based on tilt status
Dim UseVPMDMD: UseVPMDMD = (whitewater.ShowDT AND DesktopVPXDMD)

' *** Start VPM ***

if Version < 10400 then msgbox "This table requires Visual Pinball 10.4 beta or newer!" & vbnewline & "Your version: " & Version/1000

On Error Resume Next
	ExecuteGlobal GetTextFile("controller.vbs")
	If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01560000", "WPC.VBS", 3.50
Set GiCallback2 = GetRef("UpdateGI")

Sub whitewater_Init
	Dim cGameName
	Dim Light, Prim, obj
	vpmInit Me
	With Controller
		cGameName = "ww_l5"
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
		.SplashInfoLine = "Whitewater - Williams 1993"
		.HandleMechanics=0
		.HandleKeyboard=0
		.ShowDMDOnly=1
		.ShowFrame=0
		.ShowTitle=0
		.hidden = UseVPMDMD
		'.hidden = not whitewater.showdt
		 On Error Resume Next
		 .Run GetPlayerHWnd
		 If Err Then MsgBox Err.Description
		 On Error Goto 0
	 End With

	' old script from vp9: What's this for?
	'Controller.SolMask(0) = 0
	'vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the Timer to renable all the solenoids after 2 seconds
	'Controller.DIP(0) = &H00
	'Controller.Run GetPlayerHWnd
	'Controller.Switch(22) = 1 'close coin door
	'Controller.Switch(24) = 1 'and keep it close

	' Nudging
	vpmNudge.TiltSwitch = 14
	vpmNudge.Sensitivity = 1
	vpmNudge.TiltObj = Array(LeftSlingshot, RightSlingshot)

	PinMAMETimer.Interval = PinMAMEInterval
	PinMAMETimer.Enabled = 1
	LampTimer.Enabled = 1
	GILowerFade(8) : GIMiddleFade(8) : GIUpperFade(8)
	CheckMaxBalls 'Allow balls to be created at table start up
	BigFoot_Init
	InitLights(Insertlights)
	If whitewater.ShowDT or ForceSiderailsFS then
		If NoSideWallRelfections Then
			Primitive65.visible = 0 : Primitive61.visible = 0 : Primitive111.visible = 1 : Primitive112.visible = 1 : primitive113.visible = 1 : primitive114.visible = 0
		Else
			Primitive65.visible = 1 : Primitive61.visible = 0 : Primitive111.visible = 1 : Primitive112.visible = 1 : primitive113.visible = 0 : primitive114.visible = 0
		End If
	else
		If NoSideWallRelfections Then
			Primitive65.visible = 0 : Primitive61.visible = 0 : Primitive111.visible = 0 : Primitive112.visible = 0 : primitive113.visible = 1 : primitive114.visible = 1
		Else
			Primitive65.visible = 0 : Primitive61.visible = 1 : Primitive111.visible = 0 : Primitive112.visible = 0 : primitive113.visible = 0 : primitive114.visible = 0
		End If
	End If

	For Each obj In IndirectLights
		obj.FalloffPower = obj.FalloffPower * 2
		obj.IntensityScale = 3
		obj.FadeSpeedUp = obj.FadeSpeedUp * 3
		obj.FadeSpeedDown = obj.FadeSpeedDown * 3
	Next

	' adjust timerinterval to user settting
	Flasherlight24.TimerInterval = FlasherTimerInterval
	Flasherlight23.TimerInterval = FlasherTimerInterval
	FlasherFlash22.TimerInterval = FlasherTimerInterval
	Flasherlight21.TimerInterval = FlasherTimerInterval
	Flasherlight20.TimerInterval = FlasherTimerInterval
	FlasherFlash19.TimerInterval = FlasherTimerInterval
	FlasherFlash18.TimerInterval = FlasherTimerInterval
	Flasherlight17.TimerInterval = FlasherTimerInterval

End Sub

'**********
' Keys
'**********

' PlaySound "name",loopcount,volume,pan,randompitch,pitch,useexisting,restart,fade  - Y position added in VPX 10.4.
' pitch can be positive or negative and directly adds onto the standard sample frequency

Sub whitewater_KeyDown(ByVal Keycode)
	If keycode = PlungerKey Then PlaySound "fx_PlungerPull", 0, 1, 0.1, 0.25,0,0,1,1:Plunger.Pullback  ' SSF fade towards front of cab
	If keycode = LeftTiltKey Then Nudge 90, 3 : PlaySound SoundFX("fx_nudge",0)
	If keycode = RightTiltKey Then Nudge 270, 3 : PlaySound SoundFX("fx_nudge",0)
	If keycode = CenterTiltKey Then Nudge 0, 3 : PlaySound SoundFX("fx_nudge",0)
	If vpmKeyDown(keycode) Then Exit Sub
	If keycode = KeyRules Then Rules
	if keycode = LeftFlipperKey and FlippersEnabled Then SolLFlipper(True)
 	if keycode = RightFlipperKey and FlippersEnabled Then SolRflipper(True)
	' *** manual Ball Control ***
	If debugballs Then
		if keycode = 46 then If contball = 1 Then contball = 0 : BallControl.Enabled = False Else contball = 1 : BallControl.Enabled = True: End If : End If ' C Key
		if keycode = 48 then If bcboost = 1 Then bcboost = bcboostmulti Else bcboost = 1 : End If : End If 'B Key
		if keycode = 203 then bcleft = 1        ' Left Arrow
		if keycode = 200 then bcup = 1          ' Up Arrow
		if keycode = 208 then bcdown = 1        ' Down Arrow
		if keycode = 205 then bcright = 1       ' Right Arrow
	End If
End Sub

Sub whitewater_KeyUp(ByVal Keycode)
	if keycode = LeftFlipperKey and FlippersEnabled Then SolLFlipper(False)
	if keycode = RightFlipperKey and FlippersEnabled Then SolRflipper(False)
	If vpmKeyUp(keycode) Then Exit Sub
	If keycode = PlungerKey Then PLaySound "fx_plunger", 0, 1, 0.1, 0.25,0,0,1,1:Plunger.Fire  ' SSF fade towards front of cab
	' *** manual Ball Control ***
	If debugballs Then
		if keycode = 203 then bcleft = 0        ' Left Arrow
		if keycode = 200 then bcup = 0          ' Up Arrow
		if keycode = 208 then bcdown = 0        ' Down Arrow
		if keycode = 205 then bcright = 0       ' Right Arrow
	End If
End Sub

' *** manual Ball Control ***
Sub StartControl_Hit() : Set ControlBall = ActiveBall : contballinplay = true : End Sub
Sub StopControl_Hit() : contballinplay = false : End Sub

Dim bcup, bcdown, bcleft, bcright, contball, contballinplay, ControlBall, bcboost
Dim bcvel, bcyveloffset, bcboostmulti

bcboost = 1     		'Do Not Change - default setting
bcvel = 4       		'Controls the speed of the ball movement
bcyveloffset = -0.01    'Offsets the force of gravity to keep the ball from drifting vertically on the table, should be negative
bcboostmulti = 3    	'Boost multiplier to ball veloctiy (toggled with the B key)

Sub BallControl_Timer()
    If Contball and ContBallInPlay then
        If bcright = 1 Then ControlBall.velx = bcvel*bcboost Else If bcleft = 1 Then ControlBall.velx = - bcvel*bcboost Else ControlBall.velx=0 : End If : End If
        If bcup = 1 Then ControlBall.vely = -bcvel*bcboost Else If bcdown = 1 Then ControlBall.vely = bcvel*bcboost Else ControlBall.vely= bcyveloffset : End If : End If
    End If
End Sub

' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 10000)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / whitewater.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 200
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function


' *********************************************************************
'            Supporting Surround Sound Feedback (SSF) Functions
' *********************************************************************

Function AudioFade(ball) 'Calculates front-rear fade based on Y position on the table.
	Dim tmp
    tmp = ball.y * 2 / whitewater.height-1
    If tmp > 0 Then
		AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

'Set position as table object (Use object or light but NOT wall) and Vol to 1
Sub PlaySoundAt(sound, tableobj)
	PlaySound sound, 1, 1, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'Set position as table object and Vol manually.
Sub PlaySoundAtVol(sound, tableobj, Vol)
	PlaySound sound, 1, Vol, Pan(tableobj), 0, 0, 0, 1, AudioFade(tableobj)
End Sub

'Set position as table object and Vol + RndPitch manually
Sub PlaySoundAtVolPitch(sound, tableobj, Vol, RndPitch)
	PlaySound sound, 1, Vol, Pan(tableobj), RndPitch, 0, 0, 1, AudioFade(tableobj)
End Sub

'Set all as per ball position & speed, but Vol Multiplier may be used eg; PlaySoundAtBallVol "sound",3
Sub PlaySoundAtBallVol(sound, VolMult)
	PlaySound sound, 0, Vol(ActiveBall) * VolMult, Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub

' Play a looping sound at a location with volume
Sub PlayLoopSoundAtVol(sound, tableobj, Vol)
	PlaySound sound, -1, Vol, Pan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

' Requires rampbump1 to 7 in Sound Manager
Sub RandomBump(voladj, freq)
	dim BumpSnd:BumpSnd= "rampbump" & CStr(Int(Rnd*7)+1)
	PlaySound BumpSnd, 0, Vol(ActiveBall)*voladj, Pan(ActiveBall), 0, freq, 0, 1, AudioFade(ActiveBall)
End Sub

'*********
' Switches
'*********

Sub RollOverSound() : PlaySoundAtVolPitch  "rollover", ActiveBall, GlobalSoundLevel * 0.02, .25 : End Sub
Sub Targetsound() : PlaySoundAtVolPitch  "target", ActiveBall, GlobalSoundLevel * 2, .25 : End Sub
Sub RubberSleevesPins_Hit(idx) : PlaySound "post5", 0, Vol(ActiveBall)*10*GlobalSoundLevel , pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub RubberBandsRings_Hit(idx):PlaySound "rubber", 0, Vol(ActiveBall)*10*GlobalSoundLevel , pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub Wood_Hit(idx) : PlaySoundAt "WoodHit", ActiveBall : End Sub

Dim NextMetalHit:NextMetalHit = 0
Sub MetalWalls_Hit(idx)
	if BallVel(ActiveBall) > .3 and Timer > NextMetalHit then
		dim BumpSnd:BumpSnd= "metalhit" & CStr(Int(Rnd*3)+1)
		PlaySound BumpSnd, 0, Vol(ActiveBall)*metalvolume, Pan(ActiveBall), 0.5, 0, 0, 1, AudioFade(ActiveBall)
		' Schedule the next possible sound time.  This prevents it from rapid-firing noises too much.
		' Lowering these numbers allow more closely-spaced clunks.
		NextMetalHit = Timer + .1 + (Rnd * .2)
	end if
End Sub

Dim NextOrbitHit:NextOrbitHit = 0
Sub ramps_hit(idx)
	if BallVel(ActiveBall) > .3 and Timer > NextOrbitHit then
		RandomBump 15, -20000
		' Schedule the next possible sound time.  This prevents it from rapid-firing noises too much.
		' Lowering these numbers allow more closely-spaced clunks.
		NextOrbitHit = Timer + .1 + (Rnd * .2)
	end if
End Sub

Sub waverampdrop1_Hit: PlaySoundAt "fx_ballrampdrop", ActiveBall : End Sub
Sub waverampdrop2_Hit: PlaySoundAt "fx_ballrampdrop", ActiveBall : End Sub
Sub glasshit_Hit: If HitTheGlass and ActiveBall.VelY > 0 Then PlaySoundAt "glass_hit", l21b5 : Playsound "quake" : End If : End Sub

Sub Bumper1_Hit: vpmTimer.PulseSw 16 : PlaySoundAt SoundFX("BumperLeft",DOFContactors), Bumper1: End Sub
Sub Bumper2_Hit: vpmTimer.PulseSw 17 : PlaySoundAt SoundFX("BumperRight",DOFContactors), Bumper2: End Sub
Sub Bumper3_Hit: vpmTimer.PulseSw 18 : PlaySoundAt SoundFX("BumperBottom",DOFContactors), Bumper3: End Sub

Sub sw25_Hit:Controller.Switch(25) = 1:RollOverSound():End Sub
Sub sw25_Unhit:Controller.Switch(25) = 0:End Sub

Sub sw26_Hit:Controller.Switch(26) = 1:RollOverSound():End Sub
Sub sw26_Unhit:Controller.Switch(26) = 0:End Sub

Sub sw27_Hit:Controller.Switch(27) = 1:RollOverSound():End Sub
Sub sw27_Unhit:Controller.Switch(27) = 0:End Sub

Sub sw28_Hit:Controller.Switch(28) = 1:RollOverSound():End Sub
Sub sw28_Unhit:Controller.Switch(28) = 0:End Sub

Sub sw29_Hit:Controller.Switch(29) = 1:RollOverSound():End Sub
Sub sw29_Unhit:Controller.Switch(29) = 0:End Sub

Sub switch31_Hit : If Primitive_Target8.ObjRotY = 0 Then Controller.Switch(31) = 1 : Primitive_TargetParts2.ObjRotY = -1 : Primitive_Target8.ObjRotY = -1 : Me.TimerEnabled = 1 : Targetsound() : End If : End Sub
Sub switch31_timer : Me.TimerEnabled = 0 : Primitive_Target8.ObjRotY = 0 : Primitive_TargetParts2.ObjRotY = 0 : Controller.Switch(31) = 0 : end sub

Sub switch32_Hit : If Primitive_Target2.ObjRotY = 0 Then Controller.Switch(32) = 1 : Primitive_TargetParts1.ObjRotY = -1 : Primitive_Target2.ObjRotY = -1 : Me.TimerEnabled = 1 : Targetsound() : End If : End Sub
Sub switch32_timer : Me.TimerEnabled = 0 : Primitive_Target2.ObjRotY = 0 : Primitive_TargetParts1.ObjRotY = 0 : Controller.Switch(32) = 0 : end sub

Sub switch33_Hit : If Primitive_Target3.ObjRotY = 0 Then Controller.Switch(33) = 1 : Primitive_TargetParts3.ObjRotY = -1 : Primitive_Target3.ObjRotY = -1 : Me.TimerEnabled = 1 : Targetsound() : End If : End Sub
Sub switch33_timer : Me.TimerEnabled = 0 : Primitive_Target3.ObjRotY = 0 : Primitive_TargetParts3.ObjRotY = 0 : Controller.Switch(33) = 0 : end sub

Sub switch34_Hit : If Primitive_Target4.ObjRotY = 0 Then Controller.Switch(34) = 1 : Primitive_TargetParts4.ObjRotY = -1 : Primitive_Target4.ObjRotY = -1 : Me.TimerEnabled = 1 : Targetsound() : End If : End Sub
Sub switch34_timer : Me.TimerEnabled = 0 : Primitive_Target4.ObjRotY = 0 : Primitive_TargetParts4.ObjRotY = 0 : Controller.Switch(34) = 0 : end sub

Sub switch35_Hit : If Primitive_Target5.ObjRotY = 0 Then Controller.Switch(35) = 1 : Primitive_TargetParts5.ObjRotY = -1 : Primitive_Target5.ObjRotY = -1 : Me.TimerEnabled = 1 : Targetsound() : End If : End Sub
Sub switch35_timer : Me.TimerEnabled = 0 : Primitive_Target5.ObjRotY = 0 : Primitive_TargetParts5.ObjRotY = 0 : Controller.Switch(35) = 0 : end sub

Sub switch36_Hit : If Primitive13.RotZ = 0 Then Controller.Switch(36) = 1 : Primitive13.RotZ = -4 : Me.TimerEnabled = 1 : Targetsound() : End If : End Sub
Sub switch36_timer : Me.TimerEnabled = 0 : Primitive13.RotZ = 0 : Controller.Switch(36) = 0 : end sub

Sub switch37_Hit : If Primitive12.RotZ = 0 Then Controller.Switch(37) = 1 : Primitive12.RotZ = -4 : Me.TimerEnabled = 1 : Targetsound() : End If : End Sub
Sub switch37_timer : Me.TimerEnabled = 0 : Primitive12.RotZ = 0 : Controller.Switch(37) = 0 : end sub

Sub switch38_Hit : If Primitive11.RotZ = 0 Then Controller.Switch(38) = 1 : Primitive11.RotZ = -4 : Me.TimerEnabled = 1 : Targetsound() : End If : End Sub
Sub switch38_timer : Me.TimerEnabled = 0 : Primitive11.RotZ = 0 : Controller.Switch(38) = 0 : end sub

Sub switch41_Hit : If Primitive_Target7.ObjRotX = 0 Then Controller.Switch(41) = 1 : Primitive_TargetParts8.ObjRotX = 1 : Primitive_Target7.ObjRotX = 1 : Me.TimerEnabled = 1 : Targetsound() : End If : End Sub
Sub switch41_timer : Me.TimerEnabled = 0 : Primitive_Target7.ObjRotX = 0 : Primitive_TargetParts8.ObjRotX = 0 : Controller.Switch(41) = 0 : end sub

Sub switch42_Hit : If Primitive_Target6.ObjRotX = 0 Then Controller.Switch(42) = 1 : Primitive_TargetParts7.ObjRotX = 1 : Primitive_Target6.ObjRotX = 1 : Me.TimerEnabled = 1 : Targetsound() : End If : End Sub
Sub switch42_timer : Me.TimerEnabled = 0 : Primitive_Target6.ObjRotX = 0 : Primitive_TargetParts7.ObjRotX = 0 : Controller.Switch(42) = 0 : end sub

Sub SW43_Hit:Controller.Switch(43) = 1:RollOverSound():End Sub
Sub SW43_Unhit:Controller.Switch(43) = 0:End Sub

Sub sw44_Hit:Controller.Switch(44) = 1:RollOverSound():End Sub
Sub sw44_Unhit:Controller.Switch(44) = 0:End Sub

Sub sw45_Hit:Controller.Switch(45) = 1:RollOverSound():End Sub
Sub sw45_Unhit:Controller.Switch(45) = 0:End Sub

Sub sw46_Hit:Controller.Switch(46) = 1:PlaySoundAt "gate4", ActiveBall:End Sub
Sub sw46_Unhit:Controller.Switch(46) = 0:End Sub

Sub sw47_Hit:Controller.Switch(47) = 1:PlaySoundAt "gate4", ActiveBall:End Sub
Sub sw47_Unhit:Controller.Switch(47) = 0:End Sub

Sub sw48_Hit:Controller.Switch(48) = 1  End Sub
Sub sw48_Unhit:Controller.Switch(48) = 0 : PlaySoundAt "gate4", ActiveBall : End Sub

Sub swPlunger_Hit:Controller.Switch(53) = 1:End Sub
Sub swPlunger_UnHit:Controller.Switch(53) = 0:End Sub

Sub sw54_Hit:vpmTimer.PulseSw 54:RubberBand2.visible = 0::RubberBand2a.visible = 1:sw54.timerenabled = 1:End Sub
Sub sw54_timer:RubberBand2.visible = 1::RubberBand2a.visible = 0: sw54.timerenabled= 0:End Sub

Sub sw55_Hit:vpmTimer.PulseSw 55:PlaySoundAt "rubber", ActiveBall:End Sub
Sub switch56_Hit : If Primitive_Target1.ObjRotY = 0 Then Controller.Switch(56) = 1 : Primitive_TargetParts6.ObjRotY = 1 : Primitive_TargetParts6.ObjRotX = 1 : Primitive_Target1.ObjRotX = 1: Primitive_Target1.ObjRotY = 1 : Me.TimerEnabled = 1 : Targetsound() : End If : End Sub
Sub switch56_timer : Me.TimerEnabled = 0 : Primitive_Target1.ObjRotY = 0 : Primitive_TargetParts6.ObjRotX = 0 : Primitive_Target1.ObjRotX = 0 : Primitive_TargetParts6.ObjRotY = 0 : Controller.Switch(56) = 0 : end sub

Sub sw57_Hit:Controller.Switch(57) = 1:End Sub
Sub sw57_Unhit:Controller.Switch(57) = 0:End Sub

Sub sw58_Hit:Controller.Switch(58) = 1:RollOverSound():End Sub
Sub sw58_Unhit:Controller.Switch(58) = 0:End Sub

Dim popped : popped = true
Sub sw61_hit() : Controller.Switch(61) = 1 : popped = true : PlaySoundAt "BallFallInGun", ActiveBall : End Sub
Sub KickPopper(Enabled) : sw61.Kick 307,50,0 : Controller.Switch(61) = 0  : If popped Then Playsound "vuk" : popped = false : End If : End Sub

Sub FallThrough2_Hit:PlaySound "hole_enter": vpmTimer.PulseSw 62 : End Sub

' *** three ball lock ***
Dim vukked : vukked = true
Sub sw65_Hit() : Controller.Switch(65) = 1 : ActiveBall.VelX = 2 : PlaySoundAt "kicker_enter", ActiveBall : End Sub
Sub sw65_unHit() : Controller.Switch(65) = 0 : End Sub
Sub sw64_Hit() : Controller.Switch(64) = 1 : multiballwall65.isDropped = false : ActiveBall.VelX = 2 : End Sub
Sub sw64_unHit() : Controller.Switch(64) = 0 : vpmTimer.AddTimer 500, "multiballwall65.isDropped = true'" : End Sub
Sub sw63_hit() : Controller.Switch(63) = 1 : multiballwall64.isDropped = false : vukked = true : End Sub

Sub KickBallUp(Enabled) : sw63.Kick 0,60,1.50 : If vukked Then PlaySoundAt "vuk", sw63 : vpmTimer.AddTimer 300, "ballrampdropsound sw63'" : vukked = false : end if : Controller.Switch(63) = 0 : vpmTimer.AddTimer 500, "multiballwall64.isDropped = true'" : End Sub
Sub ballrampdropsound(tableobj): PlaySoundAt "fx_ballrampdrop", tableobj : End Sub

Sub sw66_Hit:Controller.Switch(66) = 1:End Sub
Sub sw66_Unhit:Controller.Switch(66) = 0:End Sub

Sub sw68_Hit:Controller.Switch(68) = 1:End Sub
Sub sw68_Unhit:Controller.Switch(68) = 0:End Sub

Sub sw71_Hit:Controller.Switch(71) = 1:RollOverSound():End Sub
Sub sw71_Unhit:Controller.Switch(71) = 0: end sub

Sub switch73_Hit : If Primitive15.RotZ = 0 Then Controller.Switch(73) = 1 : Primitive15.RotZ = -4 : Me.TimerEnabled = 1 : Targetsound() : End If : End Sub
Sub switch73_timer : Me.TimerEnabled = 0 : Primitive15.RotZ = 0 : Controller.Switch(73) = 0 : end sub

Sub switch74_Hit : If Primitive14.RotZ = 0 Then Controller.Switch(74) = 1 : Primitive14.RotZ = -4 : Me.TimerEnabled = 1 : Targetsound() : End If : End Sub
Sub switch74_timer : Me.TimerEnabled = 0 : Primitive14.RotZ = 0 : Controller.Switch(74) = 0 : end sub

Sub sw75_Hit:Controller.Switch(75) = 1:RollOverSound() : End Sub
Sub sw75_Unhit:Controller.Switch(75) = 0:End Sub

Sub sw75balldrop_Hit : PlaySoundAt "fx_ballrampdrop", ActiveBall: End Sub
Sub sw27balldrop_Hit : PlaySoundAt "fx_ballrampdrop", ActiveBall : End Sub
Sub wireguidedrop_Hit : PlaySoundAt "fx_ballrampdrop", ActiveBall:  end sub
Sub plungerfallback_hit : If ActiveBall.VelY > 0 Then PlaySoundAtVol "rubber",ActiveBall,ActiveBall.VelY / 50 : End If : End Sub

Sub TiltSol(Enabled) : FlippersEnabled = Enabled : SolLFlipper(False) : SolRFlipper(False) : End Sub

' *****************************************
'********** Sling Shot Animations *********
' *****************************************

Dim RStep, Lstep

Sub RightSlingShot_Slingshot
	vpmTimer.PulseSw 52 : PlaySoundAtVolPitch SoundFX("rightslingshot",DOFContactors), GI5,1,0.05
    RSling.Visible = 0 : RSling1.Visible = 1 : sling1.TransZ = -20 : RStep = 0 : RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 0:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 2:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0:
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
	vpmTimer.PulseSw 51 : PlaySoundAtVolPitch SoundFX("leftslingshot",DOFContactors),GI2, 1,0.05
    LSling.Visible = 0 : LSling1.Visible = 1 : sling2.TransZ = -20 : LStep = 0 : LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 0:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 2:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0:
    End Select
    LStep = LStep + 1
End Sub

'*******************************  Solenoids **************************
'*** flipper solenoids are disabled for faster response in KeyDown ***
'*********************************************************************

SolCallback(1) = "kisort"    				' Outhole
SolCallback(2) = "KickBallToLane"   		' Ball Release
SolCallback(3) = "KickPopper"  				' Whirlpool Popper
SolCallback(4) = "KickBallUp"    			' Lockup Popper
SolCallback(5) = "SolKick"          		' Kickback
SolCallback(6) = "SolDiv"           		' Ramp Diverter
SolCallback(7) = "vpmSolSound ""Knocker"","
SolCallback(25) = "SolBigFootDrive"  		' Bigfoot Drive
SolCallback(26) = "SolBigFootEnable" 		' Bigfoot Enable
SolCallback(31)="TiltSol" 					' 31 for WPC
SolModCallback(17) = "Flasherset17" 		' Bigfoot Body
SolModCallback(18) = "Flasherset18" 		' Right Mountains
SolModCallback(19) = "Flasherset19" 		' Left Mountains
SolModCallback(20) = "Flasherset20" 		' Upper Left Playfield
SolModCallback(21) = "Flasherset21" 		' Insanity Falls
SolModCallback(22) = "Flasherset22" 		' Whirlpool Popper
SolModCallback(23) = "Flasherset23" 		' Whirlpool Enter
SolModCallback(24) = "Flasherset24" 		' Bigfoot cave

Sub SolKick(Enabled) : Kickback.Enabled = Enabled : End Sub
Sub Kickback_Hit() : PlaySoundAt "kickback", ActiveBall : Kickback.kick -0, 50 + (rnd * 8) : Kickback.Enabled = False : End Sub


'******************
'Big Foot Animation
'******************

Sub SolDiv(Enabled)
	Diverter.IsDropped = Not Enabled
	If Enabled Then up = 1 : DiverterAnimationTimer.enabled = True : Playsound "TopDiverterOn" Else  up = -1 : DiverterAnimationTimer.enabled = True : Playsound "TopDiverterOn" : End If
End Sub

dim frameCount,up : frameCount=1 : up=0

sub DiverterAnimationTimer_Timer
	frameCount = frameCount + up
	If framecount > 4 or framecount < 1 Then up = 0 : me.enabled = False : end if
	Primitive_BigFootArmFur.ShowFrame frameCount : Primitive_BigFootDiverter.ShowFrame frameCount
end sub

Dim BigDir, BigCount, BigNewPos, BigOldPos : BigDir = 1:BigOldPos = 0:BigNewPos = 0

Sub BigFoot_Init : BigOldPos = 8 : controller.switch(86) = 1 : controller.switch(87) = 0 : End Sub
Sub SolBigFootDrive(Enabled) : 	BigTimer.Enabled = enabled : End Sub
Sub SolBigFootEnable(Enabled) : BigDir = ABS(Enabled) : End Sub

Sub BigTimer_Timer()
	If BigDir = 1 Then BigNewPos = BigNewPos + 1 Else BigNewPos = BigNewPos - 1 End If
	If BigNewPos <0 Then BigNewPos = 95
	If BigNewPos> 95 Then BigNewPos = 0
	Primitive_BigFoot.ObjRotZ = BigNewPos * 3.75 'value is due to 96 steps for 360 degrees rotation
	Primitive_BigFootWig.ObjRotZ = Primitive_BigFoot.ObjRotZ : Primitive_BigFootWig2.ObjRotZ = Primitive_BigFoot.ObjRotZ
	If BigNewPos = 24 Then:controller.switch(86) = 1:controller.switch(87) = 0:End If  ' Left  (Diverter)
	If BigNewPos = 45 Then:controller.switch(86) = 0:controller.switch(87) = 1:End If ' Up    (Up)
	If BigNewPos = 72 Then:controller.switch(86) = 0:controller.switch(87) = 0:End If ' Right (Unknown)
	If BigNewPos = 7 Then:controller.switch(86) = 1:controller.switch(87) = 1:End If  ' Down  (Player)
	BigOldPos = BigNewPos
End Sub

'**************
' Flipper Subs
'**************

Sub SolLFlipper(Enabled)
 If Enabled Then
		LeftFlipper.RotateToEnd : PlaySound SoundFX("FlipperL",DOFFlippers), 0, 1, -0.1, 0.05,0,0,1,1  ' SSF fade all the way towards front of cab
        if DynamicFlipperFriction Then LeftFlipper.Friction = DynamicFlipperFrictionActive
    Else
		LeftFlipper.RotateToStart : PlaySound SoundFX("FlipperDown",DOFFlippers), 0, 1, -0.1, 0.05,0,0,1,1 ' SSF fade all the way towards front of cab
        if DynamicFlipperFriction Then LeftFlipper.Friction = DynamicFlipperFrictionResting
    End If
End Sub

Sub SolRFlipper(Enabled)
  If Enabled Then
		RightFlipper.RotateToEnd() : URightFlipper.RotateToEnd() : PlaySound SoundFX("FlipperR",DOFFlippers), 0, 1, 0.1, 0.05,0,0,1,1 ' SSF fade all the way towards front of cab
		if DynamicFlipperFriction Then RightFlipper.Friction = DynamicFlipperFrictionActive : URightFlipper.Friction = DynamicFlipperFrictionActive
    Else
		RightFlipper.RotateToStart() : URightFlipper.RotateToStart(): PlaySound SoundFX("FlipperDown",DOFFlippers), 0, 1, 0.1, 0.05,0,0,1,1 ' SSF fade all the way towards front of cab
		if DynamicFlipperFriction Then RightFlipper.Friction = DynamicFlipperFrictionResting : URightFlipper.Friction = DynamicFlipperFrictionResting
    End If
End Sub

Sub OnBallBallCollision(ball1, ball2, velocity) : PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1) : End Sub

Sub LeftFlipper_Collide(parm) : PlaySound "flip_hit_1", 0, GlobalSoundLevel * parm / 50, -0.1, 0.25,0,0,1,1  : End Sub
Sub RightFlipper_Collide(parm) : PlaySound "flip_hit_1", 0, GlobalSoundLevel * parm / 50, 0.1, 0.25,0,0,1,1  : End Sub
Sub URightFlipper_Collide(parm) : PlaySound "flip_hit_1", 0, GlobalSoundLevel * parm / 50, 0.1, 0.25,0,0,1,1  : End Sub


' *****************************************************
' *** Ball Shadow code / Primitive Flipper Update *****
' *****************************************************

Dim BallShadowArray : BallShadowArray = Array (BallShadow1, BallShadow2, BallShadow3)

Sub GraphicsTimer_Timer()

	' *** move ball shadows ***
	If BallShadow then
		Dim BOT, b : BOT = GetBalls
		For b = 0 to UBound(BOT)
			BallShadowArray(b).X = BOT(b).X + (BOT(b).X - whitewater.Width/2)/8 : BallShadowArray(b).Y = BOT(b).Y + 30 : BallShadowArray(b).Z = BOT(b).Z - 24
		Next
	End If

	' *** move primitive bats ***
	batleft.objrotz = LeftFlipper.CurrentAngle + 1 : batleftshadow.objrotz = batleft.objrotz
	batright.objrotz = RightFlipper.CurrentAngle - 1 : batrightshadow.objrotz  = batright.objrotz
	batrightupper.objrotz = URightFlipper.CurrentAngle - 1 : batrightshadowupper.objrotz  = batrightupper.objrotz
End Sub

' *** Enhance the ball shine for balls on upper deck or plastic ramps ***

Dim LightArray : LightArray = Array (Light16, Light18, Light21)

Sub BallReflections_timer()
	Dim BOT, vx, vy, mx, b : BOT = Getballs
	For b = 0 to UBound(BOT)
		vx =BOT(b).VelX: vy = BOT(b).VelY
		If abs(vx) + abs(vy) > 8 and BOT(b).Z > 40 Then
			mx = (abs(vx) + abs(vy)) / 8 : If mx > 2 Then mx = 2 : End If
			LightArray(b).intensityscale = mx : LightArray(b).BulbHaloHeight = BOT(b).Z + 26
			LightArray(b).X = BOT(b).X + rnd * 100 + 20 * vx: LightArray(b).Y = BOT(b).Y + rnd * 100 + 20 * vy
		else
			LightArray(b).intensityscale = 0.03
		end if
	Next
End Sub

' *****************************************
' *** insert lights, whirlpool lights *****
' *****************************************

Dim PFLights(200,3), PFLightsCount(200)

Sub InitLights(aColl)
	Dim obj, idx
	For Each obj In aColl
		idx = obj.TimerInterval
		Set PFLights(idx, PFLightsCount(idx)) = obj
		PFLightsCount(idx) = PFLightsCount(idx) + 1
	Next
End Sub

Sub LampTimer_Timer()
	Dim chgLamp, num, chg, ii, nr
    chgLamp = Controller.ChangedLamps
	If Not IsEmpty(chgLamp) Then
        For ii = 0 To UBound(chgLamp)
			select case chglamp(ii,0)
				case 12 : PFLights(chglamp(ii,0),0).state = chglamp(ii,1) : PFLights(chglamp(ii,0),1).state = chglamp(ii,1) : Flasher12.visible = chglamp(ii,1)
				case 13 : PFLights(chglamp(ii,0),0).state = chglamp(ii,1) : PFLights(chglamp(ii,0),1).state = chglamp(ii,1) : Flasher13.visible = chglamp(ii,1)
				case 17 : PFLights(chglamp(ii,0),0).state = chglamp(ii,1) : If chglamp(ii,1) = 1 Then upf_yellow_light = 8	Else upf_yellow_light = 7 End If
				case 52 : PFLights(chglamp(ii,0),0).state = chglamp(ii,1) : PFLights(chglamp(ii,0),1).state = chglamp(ii,1) : Flasher52.visible = chglamp(ii,1)
				case 55 : PFLights(chglamp(ii,0),0).state = chglamp(ii,1) : PFLights(chglamp(ii,0),1).state = chglamp(ii,1) : If chglamp(ii,1) = 1 Then upf_red_light = 8	Else upf_red_light = 7 End If
				case 71,72,73,74,75,76 : PFLights(chglamp(ii,0),0).state = chglamp(ii,1) : If chglamp(ii,1) = 1 Then whirlpoollight(chglamp(ii,0) - 71) = 8 Else whirlpoollight(chglamp(ii,0) - 71) = 7 : End If
				case else For nr = 1 to PFLightsCount(chglamp(ii,0)) : PFLights(chglamp(ii,0),nr - 1).state = chglamp(ii,1) : Next
			end select
        Next
	End If
	Rollingupdate()
End Sub

dim whirlpoollight(6), upf_red_light, upf_yellow_light

Sub ImageLights_Timer()
	dim obj, idx
	For Each obj in whirlpoolbulb
		idx = obj.DepthBias - 71
		If whirlpoollight(idx) > 0 Then
			If whirlpoollight(idx) = 8 Then
				obj.image = "simplelight7" : obj.blenddisablelighting = 1
			Else
				whirlpoollight(idx) = whirlpoollight(idx) - 1 : obj.image = "simplelight" & whirlpoollight(idx) : obj.blenddisablelighting = whirlpoollight(idx) / 7 + 0.3
			End If
		End If
	Next
	If upf_red_light > 0 Then
		If upf_red_light = 8 Then
			Primitive100.image = "simplelight7" : Primitive100.blenddisablelighting = 1
		Else
			upf_red_light = upf_red_light - 1 : Primitive100.image = "simplelight" & upf_red_light : Primitive100.blenddisablelighting = upf_red_light / 7 + 0.3
		End If
	End If
	If upf_yellow_light > 0 Then
		If upf_yellow_light = 8 Then
			Primitive99.image = "simplelightyellow7" : Primitive99.blenddisablelighting = 2
		Else
			upf_yellow_light = upf_yellow_light - 1 : Primitive99.image = "simplelightyellow" & upf_yellow_light : Primitive99.blenddisablelighting = upf_yellow_light/ 20
		End If
	End If
End Sub

' *************************************************************
'      Based on JP's VP10 Rolling Sounds
' *************************************************************

Const tnob = 10 ' total number of balls
ReDim rolling(tnob)
InitRolling

Sub InitRolling : Dim i : For i = 0 to tnob : rolling(i) = False : Next : End Sub

Sub RollingUpdate()
    Dim BOT, b : BOT = GetBalls
    For b = 0 to UBound(BOT)
        If BallVel(BOT(b) ) > 1 and BOT(b).Y < 1800 Then
			If BOT(b).z < 60 Then
				rolling(b) = True : StopSound("plasticrolling" & b) : PlaySound("fx_ballrolling" & b), -1, GlobalSoundLevel * Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b))
			Else
				rolling(b) = True : StopSound("fx_ballrolling" & b) : PlaySound("plasticrolling" & b), -1, GlobalSoundLevel * Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b))
			End If
		Else
			If rolling(b) = True Then StopSound("fx_ballrolling" & b) : StopSound("plasticrolling" & b)  : rolling(b) = False : End If
		End If
	Next
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''  Ball Trough system ''''''''''''''''''''''''''
'''''''''''''''''''''by cyberpez''''''''''''''''''''''''''''''''
''''''''''''''''based off of EalaDubhSidhe's''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim BallCount, cBall1, cBall2, cBall3, BallSize, BallMass, MaxBalls : BallSize = 25 : BallMass = 1.7 : MaxBalls = 3

Sub CheckMaxBalls() : BallCount = MaxBalls : TroughWall1.isDropped = true : TroughWall2.isDropped = true : multiballwall65.isDropped = true : multiballwall64.isDropped = true : End Sub

Sub CreatBalls_timer()
	If BallCount > 0 then
		If BallCount = 3 Then Set cBall1 = drain.CreateSizedBallWithMass(BallSize, BallMass) : End If
		If BallCount = 2 Then Set cBall2 = drain.CreateSizedBallWithMass(BallSize, BallMass) : End If
		If BallCount = 1 Then Set cBall3 = drain.CreateSizedBallWithMass(BallSize, BallMass) : End If
		Drain.kick 70,30
		BallCount = BallCount - 1
	Else CreatBalls.enabled = false : metalvolume = 1 : End If
End Sub

Sub ballrelease_hit() : Controller.Switch(76)=1 : TroughWall1.isDropped = false : End Sub
Sub sw77_Hit() : Controller.Switch(77)=1 : TroughWall2.isDropped = false : End Sub
Sub sw77_unHit() : Controller.Switch(77)=0 : TroughWall2.isDropped = true : End Sub
Sub sw78_Hit() : Controller.Switch(78)=1 : End Sub
Sub sw78_unHit() : Controller.Switch(78)=0 : End Sub
sub kisort(enabled) : Drain.Kick 70,30 : controller.switch(15) = 0 : end sub
Sub Drain_hit() : PlaySoundAt "drain", drain : controller.switch(15) = 1 : End Sub

Dim DontKickAnyMoreBalls,DKTMstep

Sub KickBallToLane(Enabled)
	If DontKickAnyMoreBalls = 0 then
		PlaySoundAt SoundFX("ballrel",DOFContactors),BallRelease
		ballrelease.Kick 60,10
		TroughWall1.isDropped = true
		Controller.Switch(76)=0
		DontKickAnyMoreBalls = 1
		DKTMstep = 1
		DontKickToMany.enabled = true
	End If
End Sub

Sub DontKickToMany_timer()
	Select Case DKTMstep
		Case 1:
		Case 2:
		Case 3: DontKickAnyMoreBalls = 0:DontKickToMany.Enabled = False: DontKickAnyMoreBalls = 0
	End Select
	DKTMstep = DKTMstep + 1
End Sub

' ***************************************************************************************************************************
' ***             										Update GI           											*****
' ***     	Only WPC games have special GI circuit, Callback function must take two arguments							*****
' *** 		StringNo: GIString that has changed state (0-4) 															*****
' ***		for White Water: 0 = upper pf, 1 = middle pf, 2 = lower pf, 3 = backbox sky, 4 = back boat					*****
' *** 		Status: New status of GI string (0-8)																		*****
' ***************************************************************************************************************************

Sub UpdateGI(no, step)
	IF step = 1 or DisableGI Then step = 0 : end if
	Select Case no
		case 0 : GIUpperFade(step)
		case 1 : GIMiddleFade(step)
		case 2 : GILowerFade(step)
	End Select
End Sub

Sub GILowerFade(step)
	Dim obj
	For Each obj In GILower : Obj.IntensityScale = step/8 : Next
	Primitive39.blenddisablelighting = step / 16
	Primitive40.blenddisablelighting = step / 16
	Primitive5temp.material = "rampsGI" & step
	Flasherbase22.BlendDisableLighting =  step/8
End Sub

Sub GIMiddleFade(step)
	Dim obj
	For Each obj In GImiddle : Obj.IntensityScale = step/8 : Next
	Flasher7.opacity = step * 1500 : Flasher8.opacity = step * 1000 : Flasher9.opacity = step * 135 : Flasher10.opacity = step * 500
	Bulb17.image = "simplelightwhite" & step
	Rock2_Boulder_garden.material = "rockGI"  & step
	Rock1_Lower_popbumper.material = "rockGI"  & step
	Rock3_Rightpopbumper.material = "rockGI"  & step
	If step = 0 Then
		Rock2_Boulder_garden.image = "Bouldergarden_LPCompleteMap"
		Rock1_Lower_popbumper.image = "LowPopRock_LPMap"
		Rock3_Rightpopbumper.image = "RightPop-LPCompleteMap"
		Primitive57.image = "shooterrampdark"
	else
		Rock2_Boulder_garden.image = "bouldergarden1"
		Rock1_Lower_popbumper.image = "lowpop1"
		Rock3_Rightpopbumper.image = "rightpop1"
		Primitive57.image = "shooterramp"
	End If
	Primitive5temp7.material = "rampsGI" & step : Primitive5temp2.material = "rampsGI" & step : Primitive_Target3.blenddisablelighting = step/2
	Primitive_Target4.blenddisablelighting = step/2 : Primitive_Target5.blenddisablelighting = step/2 : Primitive_Target8.blenddisablelighting = step/2
	Primitive_Target2.blenddisablelighting = step/2
	Primitive57.blenddisablelighting = 0.2 + 0.2 * step / 8
End Sub

Dim GIUp ' needed for using a 2 VPX flasher objects for GI as well as an actual flasher

Sub GIUpperFade(step)
	Dim obj
	GIUp = step
	For Each obj In GIupper : Obj.IntensityScale = step/8 : Next
	For Each obj In GIupperbulbs : Obj.image = "simplelightwhite" & step : Next
	If Flashlevel20 < 0.01 Then Flasher5.opacity = step * 1000 : Flasher6.opacity = step * 75 : end if
	Flasher1.opacity = step * 250 : Flasher2.opacity = step * 150
	Primitive5temp5.material = "rampsGI" & step : Primitive5temp3.material = "rampsGI" & step : Primitive5temp8.material = "rampsGI" & step
End Sub

' *****************************************
' ***           Flasher subs          *****
' *****************************************

Dim FlashLevel17, FlashLevel18, FlashLevel19, FlashLevel20, FlashLevel21, FlashLevel22, FlashLevel23, FlashLevel24
FlasherLight18.state = 0 : FlasherLight19.state = 0 : FlasherLight22.state = 0 : FlasherLight17.state = 0 : FlasherLight17b.state = 0
FlasherLight20.state = 0 : FlasherLight21.state = 0 : FlasherLight23.state = 0 : FlasherLight24.state = 0

Sub FlasherClick(oldvalue, newvalue) : If oldvalue < 0 and newvalue > 0.2 Then PlaySound "fx_relay_on",0,0.1: End If : End Sub

' ********* Bigfoot body flasher **********
Sub Flasherset17(value) : FlasherClick FlashLevel17, value : If value < 160 Then value = 160 : End If : If value > Flashlevel17 * 255 Then FlashLevel17 = value / 255 : Flasherlight17_Timer : End If : End Sub
Sub Flasherlight17_Timer()
	dim flashx3 : flashx3 = FlashLevel17^2
	Flasherlight17.IntensityScale = 50 * flashx3 : Flasherlight17b.IntensityScale = 10 * flashx3 : FlasherFlash17.opacity = 5000 * flashx3 : FlashLevel17 = FlashLevel17 * 0.8 - 0.01
	If not Flasherlight17.TimerEnabled Then Flasherlight17.state = 1 : Flasherlight17b.state = 1 : FlasherFlash17.visible = 1 : Flasherlight17.TimerEnabled = True : End If
	If FlashLevel17 < 0 Then Flasherlight17.TimerEnabled = False : FlasherFlash17.visible = 0 : Flasherlight17.state = 0 : Flasherlight17b.state = 0 : End If
End Sub

' ********* Right Mountain flasher ********
Sub Flasherset18(value) : FlasherClick FlashLevel18, value : If value < 160 Then value = 160 : End If : If value > Flashlevel18 * 255 Then FlashLevel18 = value / 255 : FlasherFlash18_Timer : End If : End Sub
Sub FlasherFlash18_Timer()
	dim flashx3, matdim : flashx3 = FlashLevel18^3
	Flasherflash18.opacity = 3000 * flashx3^0.8 : Flasherlit18.BlendDisableLighting = 4 * FlashLevel18^0.5 : Flasherbase18.BlendDisableLighting = flashx3
	Flasherlight18.IntensityScale = 2 * flashx3 : Flasher19.opacity = 300000 * Flashx3 : Flasher18.opacity = 100000 * Flashx3
	matdim = Round(10 * FlashLevel18) : Flasherlit18.material = "domelit" & matdim : Rock4_Back_wall2.material = "domelit" & matdim
	FlashLevel18 = FlashLevel18 * 0.8 - 0.01
	If not Flasherflash18.TimerEnabled Then Flasherflash18.visible = 1 : Flasher18.visible = 1 : Flasher19.visible = 1 : Flasherlit18.visible = 1 : Rock4_Back_wall2.visible = 1 : FlasherLight18.state = 1 : Flasherflash18.TimerEnabled = True : End If
	If FlashLevel18 < 0 Then 				Flasherflash18.visible = 0 : Flasher18.visible = 0 : Flasher19.visible = 0 : Flasherlit18.visible = 0 : Rock4_Back_wall2.visible = 0 : FlasherLight18.state = 0 : Flasherflash18.TimerEnabled = False : End If
End Sub

' ******* Left Mountain flasher ************
Sub Flasherset19(value) : FlasherClick FlashLevel19, value : If value < 160 Then value = 160 : End If : If value > Flashlevel19 * 255 Then FlashLevel19 = value / 255 : FlasherFlash19_Timer : End If : End Sub
Sub FlasherFlash19_Timer()
	dim flashx3, matdim
	flashx3 = FlashLevel19^3
	Flasherflash19.opacity = 3500 * flashx3^0.8 : Flasherlit19.BlendDisableLighting = 4 * FlashLevel19^0.5 : Flasherbase19.BlendDisableLighting =  flashx3
	Flasherlight19.IntensityScale = 2 * flashx3 : Flasher15.opacity = 300000 * Flashx3 : Flasher17.opacity = 100000 * Flashx3
	matdim = Round(10 * FlashLevel19) : Flasherlit19.material = "domelit" & matdim : Rock4_Back_wall4.material = "domelit" & matdim
	FlashLevel19 = FlashLevel19 * 0.8 - 0.01
	If not Flasherflash19.TimerEnabled Then Flasherflash19.visible = 1 : Flasher15.visible = 1 : Flasher17.visible = 1 : Flasherlit19.visible = 1 : Rock4_Back_wall4.visible = 1 : FlasherLight19.state = 1 : Flasherflash19.TimerEnabled = True : End If
	If FlashLevel19 < 0 Then 				Flasherflash19.visible = 0 : Flasher15.visible = 0 : Flasher17.visible = 0 : Flasherlit19.visible = 0 : Rock4_Back_wall4.visible = 0 : FlasherLight19.state = 0 : Flasherflash19.TimerEnabled = False : End If
End Sub

' ****** Upper left playfield flasher ******
Sub Flasherset20(value) : FlasherClick FlashLevel20, value : If value < 160 Then value = 160 : End If : If value > Flashlevel20 * 255 Then FlashLevel20 = value / 255 : Flasherlight20_Timer : End If : End Sub
Sub Flasherlight20_Timer()
	dim flashx3 : flashx3 = FlashLevel20^3
	Flasherlight20.IntensityScale = 5 * flashx3: Flasher5.opacity = flashx3 * 100000 + GIUp * 1000 : Flasher6.opacity = flashx3 * 100000 + GIUp * 75
	FlashLevel20 = FlashLevel20 * 0.8 - 0.01
	If not Flasherlight20.TimerEnabled Then Flasherlight20.state = 1 : Flasherlight20.TimerEnabled = True : End If
	If FlashLevel20 < 0 Then 				Flasherlight20.state = 0 : Flasherlight20.TimerEnabled = False : End If
End Sub

' ******** Insanity Falls flasher **********
Sub Flasherset21(value) : FlasherClick FlashLevel21, value : If value < 160 Then value = 160 : End If : If value > Flashlevel21 * 255 Then FlashLevel21 = value / 255 : Flasherlight21_Timer : End If : End Sub
Sub Flasherlight21_Timer()
	dim flashx3 : flashx3 = FlashLevel21^3
	FlasherFlash21.opacity = 1000000 * flashx3 : Flasherlight21.IntensityScale = 20 * flashx3 : Flasher16.opacity = 2000 * FlashLevel21 :
	FlashLevel21 = FlashLevel21 * 0.8 - 0.01
	If not Flasherlight21.TimerEnabled Then FlasherFlash21.visible = 1 : Flasher16.visible = 1 : Flasherlight21.state = 1 : Flasherlight21.TimerEnabled = True : End If
	If FlashLevel21 < 0 Then 				FlasherFlash21.visible = 0 : Flasher16.visible = 0 : Flasherlight21.state = 0 : Flasherlight21.TimerEnabled = False : End If
End Sub

' ******* Whirlpool popper flasher *********
Sub Flasherset22(value) : FlasherClick FlashLevel22, value : If value < 160 Then value = 160 : End If : If value > Flashlevel22 * 255 Then FlashLevel22 = value / 255 : FlasherFlash22_Timer : End If : End Sub
Sub FlasherFlash22_Timer()
	dim flashx3, matdim : flashx3 = FlashLevel22^3
	Flasherflash22.opacity = 2000 * flashx3^0.6 :  Flasherlit22.BlendDisableLighting = 2 * FlashLevel22^0.5
	Flasherlight22.IntensityScale = flashx3 : Flasher11.opacity = 200000 * Flashx3 : Flasher14.opacity = 30000 * Flashx3 : Flasher20.opacity = 2000 * FlashLevel22
	matdim = Round(10 * FlashLevel22) : Flasherlit22.material = "domelit" & matdim : Rock6_lost_mine_flash.material = "domelit" & matdim
	FlashLevel22 = FlashLevel22 * 0.85 - 0.01
	If not Flasherflash22.TimerEnabled Then Flasherflash22.visible = 1 : Flasher11.visible = 1 : Flasher14.visible = 1 : Flasher20.visible = 1 : Flasherlit22.visible = 1 : Rock6_lost_mine_flash.visible = 1 : Flasherlight22.state = 1 : Flasherflash22.TimerEnabled = True : End If
	If FlashLevel22 < 0 Then Flasherlit22.visible = 0 : Flasherflash22.TimerEnabled = False : Flasherflash22.visible = 0 :	Rock6_lost_mine_flash.visible = 0 : Flasher20.visible = 0 : Flasher11.visible = 0 : Flasher14.visible = 0 : Flasherlight22.state =  0 : End If
End Sub

' ******* Enter Whirlpool flasher ***********
Sub Flasherset23(value) : FlasherClick FlashLevel23, value : If value < 160 Then value = 160 : End If : If value > Flashlevel23 * 255 Then FlashLevel23 = value / 255 : Flasherlight23_Timer : End If : End Sub
Sub Flasherlight23_Timer()
	Flasherlight23.IntensityScale = 5 * FlashLevel23^3 : Primitive77.BlendDisableLighting = 1000 * FlashLevel23^3 : FlashLevel23 = FlashLevel23 * 0.8 - 0.01
	If not Flasherlight23.TimerEnabled Then Flasherlight23.state = 1 : Flasherlight23.TimerEnabled = True : End If
	If FlashLevel23 < 0 Then 				Flasherlight23.state = 0 : Flasherlight23.TimerEnabled = False : End If
End Sub

' ********** Bigfoot cave flasher ***********
Sub Flasherset24(value) : FlasherClick FlashLevel24, value : If value > Flashlevel24 * 255 Then FlashLevel24 = value / 255 : Flasherlight24_Timer : End If : End Sub
Sub Flasherlight24_Timer()
	dim flashx3 : flashx3 = FlashLevel24^3^0.8 : Flasherlight24.IntensityScale = 5 * flashx3: FlasherFlash24.opacity = 50000 * flashx3 : FlashLevel24 = FlashLevel24 * 0.8 - 0.01
	If not Flasherlight24.TimerEnabled Then FlasherFlash24.visible = 1 : Flasherlight24.state = 1 : Flasherlight24.TimerEnabled = True : End If
	If FlashLevel24 < 0 Then 				FlasherFlash24.visible = 0 : Flasherlight24.state = 0 : Flasherlight24.TimerEnabled = False: End If
End Sub

' Thalamus : Exit in a clean and proper way
Sub whitewater_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


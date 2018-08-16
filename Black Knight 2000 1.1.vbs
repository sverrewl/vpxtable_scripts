'*****************************************************************************************
'****************************** Williams, Black Knight 2000, 1989 ************************
'*****************************************************************************************

' Black Knight 2000 by Flupper
' Based on VP8 version of Lio
' upper and lower playfield redraw by Tomasaco
' Parts of the script/table taken from Totan, AFM, Dirty Harry
' Script review / DOF changes / some soundfx by Ninuzzu
' Several plastics photos from Johngreve
' So big thanks to Lio, Ninuzzu, Tomasaco, Johngreve, JPSalas, Dozer, Knorr

' Thalamus 2018-07-19
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' Thallamus 2018-08-16 : Improved directional positions

Option Explicit
Randomize

Const VolDiv = 2000

Const VolBump   = 2    ' Bumpers multiplier.
Const VolRol    = 1    ' Rollovers volume multiplier.
Const VolGates  = 1    ' Gates volume multiplier.
Const VolMetals = 1    ' Metals volume multiplier.
Const VolRB     = 1    ' Rubber bands multiplier.
Const VolRH     = 1    ' Rubber hits multiplier.
Const VolRPo    = 1    ' Rubber posts multiplier.
Const VolRPi    = 1    ' Rubber pins multiplier.
Const VolPlast  = 1    ' Plastics multiplier.
Const VolTarg   = 1    ' Targets multiplier.
Const VolWood   = 1    ' Woods multiplier.

Const VolSpin   = 1.5  ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.

Dim LightningBats, ShowRansom, GlowAmount, InsertBrightness, AmbienceCategory
Dim GlowBall, ChooseBall, CustomBulbIntensity(10), red(10), green(10), Blue(10)
Dim CustomBallImage(10), CustomBallLogoMode(10), CustomBallDecal(10), CustomBallGlow(10)
Dim ForceSiderailsFS

'******************************************************************************************
'* TABLE OPTIONS **************************************************************************
'******************************************************************************************

'*** Lightning bolt bats activate: true = normal bats, false = Lightning flipper bats *****

LightningBats = True

'*** Show Ransom: true = show ransom plate, false = do not show ransom plate **************

ShowRansom = True

' *** Ambience Category: on = true, off = false, only works in DeskTop mode ***************

AmbienceCategory = False

' *** Show siderails in fullscreen mode: True = show siderails, False = do not show *******

ForceSiderailsFS = False

'********************************** Insert Lighting settings ******************************
'The 2 settings below together with Night->Day cycle setting change the lighting looks.
'Night time: Night->day cycle completely left, GlowAmount = 3.0, InsertBrightness = 2.0
'Day time:   Night->day cycle 2 notches from left, GlowAmount = 0.5, InsertBrightness = 0.5
'Mixed:      Night->day cycle 1 notch from left, GlowAmount = 1.0, InsertBrightness = 0.5
'******************************************************************************************

'1.0 = default, useful range 0.1 - 3

GlowAmount = 1.0

'0.75 = default, useful range 0.25 - 2

InsertBrightness = 0.75

'********************************** Ball settings *****************************************
' 0 = normal ball
' 1 = white GlowBall
' 2 = magma GlowBall
' 3 = blue GlowBall
' 4 = another normal ball
' 5 = earth
'******************************************************************************************

ChooseBall = 0

' default Ball
CustomBallGlow(0) = 		False
CustomBallImage(0) = 		"TTMMball"
CustomBallLogoMode(0) = 	False
CustomBallDecal(0) = 		"scratches"
CustomBulbIntensity(0) = 	10
Red(0) = 0 : Green(0)	= 0 : Blue(0) = 0

' white GlowBall
CustomBallGlow(1) = 		True
CustomBallImage(1) = 		"white"
CustomBallLogoMode(1) = 	True
CustomBallDecal(1) = 		""
CustomBulbIntensity(1) = 	0
Red(1) = 255 : Green(1)	= 255 : Blue(1) = 255

' Magma GlowBall
CustomBallGlow(2) = 		True
CustomBallImage(2) = 		"ballblack"
CustomBallLogoMode(2) = 	True
CustomBallDecal(2) = 		"ballmagma"
CustomBulbIntensity(2) = 	0
Red(2) = 255 : Green(2)	= 180 : Blue(2) = 100

' Blue ball
CustomBallGlow(3) = 		True
CustomBallImage(3) = 		"blueball2"
CustomBallLogoMode(3) = 	False
CustomBallDecal(3) = 		""
CustomBulbIntensity(3) = 	0
Red(3) = 30 : Green(3)	= 40 : Blue(3) = 200

' Pin ball
CustomBallGlow(4) = 		False
CustomBallImage(4) = 		"pinball"
CustomBallLogoMode(4) = 	False
CustomBallDecal(4) = 		"scratch2"
CustomBulbIntensity(4) = 	10
Red(4) = 0 : Green(4)	= 0 : Blue(4) = 0

' Earth
CustomBallGlow(5) = 		True
CustomBallImage(5) = 		"ballblack"
CustomBallLogoMode(5) = 	True
CustomBallDecal(5) = 		"earth"
CustomBulbIntensity(5) = 	0
Red(5) = 100 : Green(5)	= 100 : Blue(5) = 100


'******************************************************************************************
'******************************************************************************************
'******************************************************************************************

' release notes

' version 1.0 : initial version

' version 1.1
' SliderPoint's new plunger lane (no more magic balls!)
' Ninuzzu's fix for sounds for rolling on metal and balldrop
' nFozzy's physics changes (rubbers, kickers, FlipperTricks code, ...)
' extensive testing by JohnGreve
' small visual fixes
' added Glowball code for upper wireramps and different switches in beginning of script
' added FS siderails switch
' fixed ball entering trough when falling back down plunger lane
' fixed potential ball hang before left ramp to upf

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const BallSize = 50

LoadVPM "00990300", "S11.VBS", 3.10

'********************
'Standard definitions
'********************

Const UseSolenoids = 2
Const UseLamps = 0
Const UseSync = 0
Const HandleMech = 0
Const cGameName = "bk2k_l4"
Const Wirerampsound = "metalrolling2"
Const sCoin = "fx_coin"

Dim bsTrough,bsREject,bsBallPopper,dtKNI,dtGHT, Lock, Mech3Bank, MagnaSave

Set MotorCallback = GetRef("RollingUpdate") 'realtime updates - rolling sound

'********************
'Light definitions
'********************

Dim LightType(65), LightObjectFirst(65), LightObjectSecond(65)
Dim GlowLow(10), GlowMed(10), GlowHigh(10)

' *** LightType 0 = normal
' *** LightType 1 = MagnaSave
' *** LightType 2 = Ransom

LightType(1)  = 2 : Set LightObjectFirst(1)  = FlasherR										  '*R*ANSOM (Backbox)
LightType(2)  = 2 : Set LightObjectFirst(2)  = FlasherA 		  							  'R*A*NSOM (Backbox)
LightType(3)  = 2 : Set LightObjectFirst(3)  = FlasherN 	      							  'RA*N*SOM (Backbox)
LightType(4)  = 2 : Set LightObjectFirst(4)  = FlasherS 	      							  'RAN*S*OM (Backbox)
LightType(6)  = 2 : Set LightObjectFirst(6)  = FlasherO 		  							  'RANS*O*M (Backbox)
LightType(7)  = 2 : Set LightObjectFirst(7)  = FlasherM 		  							  'RANSO*M* (Backbox)
LightType(19) = 1																			  'Magna Save
LightType(5)  = 0 : Set LightObjectFirst(5)  = light5b  : Set LightObjectSecond(5)  = Light5  'Bolt Circle Center
LightType(8)  = 0 : Set LightObjectFirst(8)  = Light8b  : Set LightObjectSecond(8)  = Light8  'LAST CHANCE (L. Outlane)
LightType(9)  = 0 : Set LightObjectFirst(9)  = Light9b  : Set LightObjectSecond(9)  = Light9  'U-Turn Bolt (Right)
LightType(10) = 0 : Set LightObjectFirst(10) = Light10b : Set LightObjectSecond(10) = Light10 'Spin Bolt (Ball Popper)
LightType(11) = 0 : Set LightObjectFirst(11) = Light11b : Set LightObjectSecond(11) = Light11 'Lock Bolt (R. Eject)
LightType(12) = 0 : Set LightObjectFirst(12) = Light12b : Set LightObjectSecond(12) = Light12 '*B* in BLACK
LightType(13) = 0 : Set LightObjectFirst(13) = Light13b : Set LightObjectSecond(13) = Light13 '*L* in BLACK
LightType(14) = 0 : Set LightObjectFirst(14) = Light14b : Set LightObjectSecond(14) = Light14 '*A* in BLACK
LightType(15) = 0 : Set LightObjectFirst(15) = Light15b : Set LightObjectSecond(15) = Light15 '*C* in BLACK
LightType(16) = 0 : Set LightObjectFirst(16) = Light16b : Set LightObjectSecond(16) = Light16 '*K* in BLACK
LightType(17) = 0 : Set LightObjectFirst(17) = Light17b : Set LightObjectSecond(17) = Light17 'Extra Ball Bolt
LightType(18) = 0 : Set LightObjectFirst(18) = Light18b : Set LightObjectSecond(18) = Light18 'Red Bolt (UPF right)
LightType(20) = 0 : Set LightObjectFirst(20) = Light20b : Set LightObjectSecond(20) = Light20 'Blue Bolt (UPF left)
LightType(21) = 0 : Set LightObjectFirst(21) = Light21b : Set LightObjectSecond(21) = Light21 'Last Chance (R. Outlane)
LightType(22) = 0 : Set LightObjectFirst(22) = Light22b : Set LightObjectSecond(22) = Light22 'U-Turn Bolt (Left)
LightType(23) = 0 : Set LightObjectFirst(23) = light23b : Set LightObjectSecond(23) = Light23 'KICKBACK (L. Outlane)
LightType(24) = 0 : Set LightObjectFirst(24) = Light24b : Set LightObjectSecond(24) = Light24 'Shoot Again
LightType(25) = 0 : Set LightObjectFirst(25) = Light25b : Set LightObjectSecond(25) = Light25 'W-Lane (Top UPF)
LightType(26) = 0 : Set LightObjectFirst(26) = Light26b : Set LightObjectSecond(26) = Light26 'I-Lane (Top UPF)
LightType(27) = 0 : Set LightObjectFirst(27) = Light27b : Set LightObjectSecond(27) = Light27 'N-Lane (Top UPF)
LightType(28) = 0 : Set LightObjectFirst(28) = Light28b : Set LightObjectSecond(28) = Light28 'W-Lane (Btm UPF)
LightType(29) = 0 : Set LightObjectFirst(29) = Light29b : Set LightObjectSecond(29) = Light29 'A-Lane (Btm UPF)
LightType(30) = 0 : Set LightObjectFirst(30) = Light30b : Set LightObjectSecond(30) = Light30 'R-Lane (Btm UPF)
LightType(31) = 0 : Set LightObjectFirst(31) = Light31b : Set LightObjectSecond(31) = Light31 'JACKPOT Bolt
LightType(32) = 0 : Set LightObjectFirst(32) = Light32b : Set LightObjectSecond(32) = Light32 'ADVANCE RANSOM Bolt
LightType(33) = 0 : Set LightObjectFirst(33) = Light33b : Set LightObjectSecond(33) = Light33 'Motor Target Bolt 3 (L UPF)
LightType(34) = 0 : Set LightObjectFirst(34) = Light34b : Set LightObjectSecond(34) = Light34 'Motor Target Bolt 2 (L UPF)
LightType(35) = 0 : Set LightObjectFirst(35) = Light35b : Set LightObjectSecond(35) = Light35 'Motor Target Bolt 1 (L UPF)
LightType(36) = 0 : Set LightObjectFirst(36) = Light36b : Set LightObjectSecond(36) = light36 '2X (Bonus Lights LPF)
LightType(37) = 0 : Set LightObjectFirst(37) = Light37b : Set LightObjectSecond(37) = Light37 '3X (Bonus Lights LPF)
LightType(38) = 0 : Set LightObjectFirst(38) = Light38b : Set LightObjectSecond(38) = Light38 'BONUS HOLD (Bonus Lights LPF)
LightType(39) = 0 : Set LightObjectFirst(39) = Light39b : Set LightObjectSecond(39) = Light39 '4X (Bonus Lights LPF)
LightType(40) = 0 : Set LightObjectFirst(40) = Light40b : Set LightObjectSecond(40) = Light40 '5X (Bonus Lights LPF)
LightType(41) = 0 : Set LightObjectFirst(41) = Light41b : Set LightObjectSecond(41) = Light41 '*G* in KNIGHT (R 3-Bank Dr Tgt)
LightType(42) = 0 : Set LightObjectFirst(42) = Light42b : Set LightObjectSecond(42) = Light42 '*H* in KNIGHT (R 3-Bank Dr Tgt)
LightType(43) = 0 : Set LightObjectFirst(43) = Light43b : Set LightObjectSecond(43) = Light43 '*T* in KNIGHT (R 3-Bank Dr Tgt)
LightType(44) = 0 : Set LightObjectFirst(44) = Light44b : Set LightObjectSecond(44) = Light44 '*K* in KNIGHT (L 3-Bank Dr Tgt)
LightType(45) = 0 : Set LightObjectFirst(45) = Light45b : Set LightObjectSecond(45) = Light45 '*N* in KNIGHT (L 3-Bank Dr Tgt)
LightType(46) = 0 : Set LightObjectFirst(46) = Light46b : Set LightObjectSecond(46) = Light46 '*I* in KNIGHT (L 3-Bank Dr Tgt)
LightType(47) = 0 : Set LightObjectFirst(47) = Light47b : Set LightObjectSecond(47) = Light47 'SKYWAY Bolt
LightType(48) = 0 : Set LightObjectFirst(48) = Light48b : Set LightObjectSecond(48) = Light48 'HURRY-UP Bolt
LightType(49) = 0 : Set LightObjectFirst(49) = Light49b : Set LightObjectSecond(49) = Light49 'EXTRA BALL (Bolt Circle)
LightType(50) = 0 : Set LightObjectFirst(50) = Light50b : Set LightObjectSecond(50) = Light50 '50,000 (Bolt Circle)
LightType(51) = 0 : Set LightObjectFirst(51) = Light51b : Set LightObjectSecond(51) = Light51 'MAGNA SAVE (Bolt Circle)
LightType(52) = 0 : Set LightObjectFirst(52) = Light52b : Set LightObjectSecond(52) = Light52 '10,000 (Bolt Circle)
LightType(53) = 0 : Set LightObjectFirst(53) = Light53b : Set LightObjectSecond(53) = Light53 'MULTI-BALL (Bolt Circle)
LightType(54) = 0 : Set LightObjectFirst(54) = Light54b : Set LightObjectSecond(54) = Light54 '100,000 (Bolt Circle)
LightType(55) = 0 : Set LightObjectFirst(55) = Light55b : Set LightObjectSecond(55) = Light55 'RANSOM (Bolt Circle)
LightType(56) = 0 : Set LightObjectFirst(56) = Light56b : Set LightObjectSecond(56) = Light56 '200,000 (Bolt Circle)
LightType(57) = 0 : Set LightObjectFirst(57) = Light57b : Set LightObjectSecond(57) = Light57 'SPECIAL (Bolt Circle)
LightType(58) = 0 : Set LightObjectFirst(58) = Light58b : Set LightObjectSecond(58) = Light58 '20,000 (Bolt Circle)
LightType(59) = 0 : Set LightObjectFirst(59) = Light59b : Set LightObjectSecond(59) = Light59 'KICKBACK (Bolt Circle)
LightType(60) = 0 : Set LightObjectFirst(60) = Light60b : Set LightObjectSecond(60) = Light60 '150,000 (Bolt Circle)
LightType(61) = 0 : Set LightObjectFirst(61) = Light61b : Set LightObjectSecond(61) = Light61 'DRAWBRIDGE (Bolt Circle)
LightType(62) = 0 : Set LightObjectFirst(62) = Light62b : Set LightObjectSecond(62) = Light62 '75,000 (Bolt Circle)
LightType(63) = 0 : Set LightObjectFirst(63) = Light63b : Set LightObjectSecond(63) = Light63 'HURRY-UP (Bolt Circle)
LightType(64) = 0 : Set LightObjectFirst(64) = Light64b : Set LightObjectSecond(64) = Light64 '250,000 (Bolt Circle)

RightFlipper.TimerInterval = 16
LeftFlipper.TimerInterval = 16

Set GlowLow(0) = Glowball1low : Set GlowMed(0) = Glowball1med : Set GlowHigh(0) = Glowball1high
Set GlowLow(1) = Glowball2low : Set GlowMed(1) = Glowball2med : Set GlowHigh(1) = Glowball2high
Set GlowLow(2) = Glowball3low : Set GlowMed(2) = Glowball3med : Set GlowHigh(2) = Glowball3high


'************
' Table init.
'************

Sub BK2K_Init
	Dim Light

	vpmInit Me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Black Knight 2000 - Williams 1989" & vbNewLine & "VPX by Flupper"
        .Games(cGameName).Settings.Value("rol") = 0 'set it to 1 to rotate the DMD to the left
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = 0
        On Error Resume Next
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
        On Error Goto 0
        .Switch(22) = 1 'close coin door
        .Switch(24) = 1 'and keep it close
    End With

	'*** option settings ***

	If AmbienceCategory and BK2K.ShowDT Then Flasher1.Visible = 1 Else Flasher1.Visible = 0 End If
	If LightningBats Then LeftFlipper.image = "flipperaoleft" : RightFlipper.image = "flipperaoright" : URightFlipper.image = "flipperaoright" End If
	If not ShowRansom Then ransomplate.visible = 0 : FlasherR.visible = 0 : FlasherA.visible = 0 : FlasherN.visible = 0 : FlasherS.visible = 0 : FlasherO.visible = 0 : FlasherM.visible = 0 end If
	For Each Light in GlowLights : Light.IntensityScale = GlowAmount: Light.FadeSpeedUp = Light.Intensity /  10 : Light.FadeSpeedDown = Light.FadeSpeedUp : Next
	For Each Light in InsertLights : Light.IntensityScale = InsertBrightness : Light.FadeSpeedUp = Light.Intensity / 10 : Light.FadeSpeedDown = Light.FadeSpeedUp : Next
	If BK2K.ShowDT or ForceSiderailsFS then wallleftFS.isdropped = 1 : wallrightFS.isdropped = 1 Else lockbar.Visible = 0 End If
	ChangeBall(ChooseBall)

    ' ### Nudging ###
    vpmNudge.TiltSwitch = 1 'was 14 nFozzy
    vpmNudge.Sensitivity = 1
	vpmNudge.TiltObj = Array(Bumper1,Bumper2,Bumper3,LSling,RSling)

	' ### Main Timer init ###
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

	' autoplunger pullback
    LockPlunger.PullBack

    ' Trough
    Set bsTrough = New cvpmTrough
    With bsTrough
        .size = 3
        .initSwitches Array(swTrough1,swTrough2,swTrough3)
        .Initexit BallRelease, 90, 4
        .InitEntrySounds "fx_drain", "", ""
        .InitExitSounds SoundFX("fx_Solenoid",DOFContactors), SoundFX("fx_ballrel",DOFContactors)
        .Balls = 3
		.EntrySw = 10
    End With

	'Right Eject - Knights Challenge (Lower Right)
	Set bsREject = New cvpmSaucer
	With bsREject
		'.InitKicker REject, swREject, 180, 2.2, 0
		.InitKicker REject, swREject, 250, 40, 0
		.InitSounds "SolIn", SoundFX("solenoid",DOFContactors), SoundFX("solenoid",DOFContactors)
	End With

 	'Ball Popper - Spin
	Set bsBallPopper = New cvpmSaucer
	With bsBallPopper
		.InitKicker BallPopper1, swBallPopper, 220, 60, 1.56 'was 45, now 60 based on JohnGreve's comments
		.InitSounds "SolIn", SoundFX("solenoid",DOFContactors), Wirerampsound
		.CreateEvents "bsBallPopper", BallPopper1
	End With

	' KNI-Targets (Left Bank)
	Set dtKNI = New cvpmDropTarget
	With dtKNI
        .InitDrop Array(sw4, sw5, sw6), Array(44, 45, 46)
        .initsnd SoundFX("fx_droptarget", DOFContactors), SoundFX("fx_resetdrop", DOFContactors)
        .CreateEvents "dtKNI"
    End With

	' GHT-Targets (Right Bank)
	Set dtGHT = New cvpmDropTarget
	With dtGHT
        .InitDrop Array(sw1, sw2, sw3), Array(41, 42, 43)
        .initsnd SoundFX("fx_droptarget", DOFContactors), SoundFX("fx_resetdrop", DOFContactors)
        .CreateEvents "dtGHT"
    End With

	'ball lock
	Set Lock = New cvpmVLock
	With Lock
		.InitVLock Array(Lock1,Lock2,Lock3), Array(LockKicker1,LockKicker2,LockKicker3), Array(swLock1,swLock2,swLock3)
		.InitSnd Wirerampsound, SoundFX("solenoid",DOFContactors)
		.ExitDir = 0
		.ExitForce = 20
		.KickForceVar = 1
		.CreateEvents "Lock"
	End With

	'3 Targets Bank (copied from AFM)
    Set Mech3Bank = new cvpmMech
    With Mech3Bank
        .Sol1 = sMTargets
        .Mtype = vpmMechLinear + vpmMechReverse + vpmMechOneSol
        .Length = 60
        .Steps = 50
        .AddSw swMTargetsUp, 0, 0
        .AddSw swMTargetsDn, 50, 50
        .Callback = GetRef("Update3Bank")
        .Start
    End With

	'Setup magnets
	Set MagnaSave = New cvpmMagnet
	With MagnaSave
		.InitMagnet MSave, 12
		.Solenoid = sMagnaSave
		.GrabCenter = False 	'Try to set to true it if you like...i prefer it off as that´s closer to the real bk2k magnet
	End With
End Sub

' keyboard handlers
Sub BK2K_KeyDown(ByVal keycode)
	If keycode = LeftTiltKey Then Nudge 90, 3 : PlaySound SoundFX("fx_nudge",0)
	If keycode = RightTiltKey Then Nudge 270, 3 : PlaySound SoundFX("fx_nudge",0)
	If keycode = CenterTiltKey Then Nudge 0, 3 : PlaySound SoundFX("fx_nudge",0)
	If keycode = RightMagnaSave Then:Controller.Switch(59) = 1:End If
	If keycode = PlungerKey Then PlaySoundAtVol "fx_PlungerPull", Plunger, 1:Plunger.Pullback
	If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub BK2K_KeyUp(ByVal keycode)
	If keycode = PlungerKey Then PLaySoundAtVol "fx_plunger", plunger, 1:Plunger.Fire
	If keycode = RightMagnaSave Then:Controller.Switch(59) = False:End If
	If vpmKeyUp(keycode) Then Exit Sub
End Sub

Sub BK2K_Paused: Controller.Pause = 1:End Sub
Sub BK2K_UnPaused: Controller.Pause = 0:End Sub
Sub BK2K_Exit(): Controller.Stop:End Sub

'******************* Switches ***********************

Const swOuthole	   = 10 ' Outhole
Const swTrough1	   = 11 ' Ball Trough 1 (left)
Const swTrough2	   = 12 ' Ball Trough 2 (middle)
Const swTrough3    = 13 ' Ball Trough 3 (right)
Const swSwitch14   = 14 ' not used
Const swSwitch15   = 15 ' not used
Const swMTargetsUp = 16 ' UP Motor Targets
Const swJet1       = 17 ' Left Jet Bumper
Const swLSling     = 18 ' Left Slingshot
Const swJet2       = 19 ' Right Jet Bumper
Const swRSling     = 20 ' Left Slingshot
Const swJet3       = 21 ' Bottom Jet Bumper
Const swLoop1      = 22 ' Left UPF Loop End
Const swLoop2      = 23 ' Right UPF Loop End
Const swMTargetsDn = 24 ' Down Motor Targets
Const swWin1       = 25 ' W-Lane in WIN
Const swWin2       = 26 ' I-Lane in WIN
Const swWin3       = 27 ' N-Lane in WIN
Const swWar1       = 28 ' W-Lane in WAR
Const swWar2       = 29 ' W-Lane in WAR
Const swWar3       = 30 ' W-Lane in WAR
Const swURampEntry = 31 ' Upper Ramp Entry
Const swLRampExit  = 32 ' Lower Ramp Exit (skyway)
Const swBridge1    = 33 ' Motor Targets (Upper Target)
Const swBridge2    = 34 ' Motor Targets (Middle Target)
Const swBridge3    = 35 ' Motor Targets (Lower Target)
Const swLock1      = 36 ' UPF Lock Switch (Lower)
Const swLock2      = 37 ' UPF Lock Switch (Middle)
Const swLock3      = 38 ' UPF Lock Switch (Upper)
Const swLOutlane   = 39 ' Left Outlane
Const swREject     = 40 ' Right Eject (Challenge Hole)
Const swKNIGHT4    = 41 ' *G* Target in KNIGHT (Right Dt Bank - left)
Const swKNIGHT5    = 42 ' *H* Target in KNIGHT (Right Dt Bank - middle)
Const swKNIGHT6    = 43 ' *T* Target in KNIGHT (Right Dt Bank - right)
Const swKNIGHT1    = 44 ' *K* Target in KNIGHT (Left Dt Bank - lower)
Const swKNIGHT2    = 45 ' *N* Target in KNIGHT (Left Dt Bank - middle)
Const swKNIGHT3    = 46 ' *I* Target in KNIGHT (Left Dt Bank - upper)
Const swBallPopper = 47 ' Ball Popper
Const swROutlane   = 48 ' Right Outlane
Const swUTurn1     = 49 ' U-Turn 1 (Left Entrance lower switch)
Const swUTurn2     = 50 ' U-Turn 2 (Left Entrance upper switch)
Const swUTurn3     = 51 ' U-Turn 3 (Right Entrance upper switch)
Const swUTurn4     = 52 ' U-Turn 4 (Right Entrance lower switch)
Const swShooter    = 53 ' Ball Shooter
Const swRRetLane   = 54 ' Right Return Lane
Const swLRetLane   = 55 ' Left Return Lane
Const swSwitch55   = 56 ' not used
Const swRLaneChg   = 57 ' Flipper Right
Const swLLaneChg   = 58 ' Flipper Left
Const swMagnaSave  = 59 ' Magna Save
Const swSwitch60   = 60 ' not used
Const swSwitch61   = 61 ' not used
Const swSwitch62   = 62 ' not used
Const swSwitch63   = 63 ' not used
Const swSwitch64   = 64 ' not used


' -------------------------------------
' Solenoids
' -------------------------------------

'Solenoids A-SIDE
Const sOuthole      = 1 ' Outhole Kicker
Const sBallRelease  = 2 ' Ball Release
Const sLDropBank    = 3 ' Left 3-Bank Drop Target Reset
Const sRDropBank    = 4 ' Right 3-Bank Drop Target Reset
'Const sSolenoid5    = 5 ' not used
Const sBallPopper   = 6 ' Ball Popper
Const sLockPlunger  = 7 ' UPF Lockup Kickback
Const sREject       = 8 ' Right Eject (Challenge hole)

'Solenoids C-SIDE
Const sUPFRedBoltFlash  = 25 '1c (UPF Big Red Bolt Flasher)
Const sUPFBlueBoltFlash = 26 '2c (UPF Big Blue Bolt Flasher)
Const sKnightHeadFlash  = 27 '3c (Bolt Circle Center Flasher)
Const sFlipLaneFlash    = 28 '4c (Flipper Lane Flasher)
Const sDTFlash          = 29 '5c (Drop Target Flasher)
Const sSkyWayFlash      = 30 '6c (LPF Ramp Flasher)
Const sREjectFlash      = 31 '7c (Right Eject/Upr Flipper Flashers)
Const sLockFlash        = 32 '8c (UPF Lockup Kickback Flasher)

'And The Rest ;)
Const sIGIRelay     = 9  ' (Insert Board GI-Relay)
Const sUPFGIRelay   = 10 ' (UPF GI-Relay)
Const sLPFGIRelay   = 11 ' (LPF GI-Relay)
Const sACselect     = 12 ' (A/C Select Relay)
Const sKickback     = 13 ' (Kickback - Left Outlane)
Const sKnocker      = 14 ' (Knocker)
Const sMagnaSave    = 15 ' (Magna Save Driver)
Const sMTargets     = 16 ' (Motor Targets[UPF] Relay)
Const sLJet         = 17 ' (Left Jet Bumper)
Const sLSling       = 18 ' (Left Slingshot)
Const sRJet         = 19 ' (Right Jet Bumper)
Const sRSling       = 20 ' (Right SlingShot)
Const sBJet         = 21 ' (Bottom Jet Bumper)
'Const sSolenoid     = 22 ' not used

'Solenoid CallBacks (A-Side)
SolCallback(sOuthole)      = "bsTrough.SolIn"
SolCallback(sBallRelease)  = "bsTrough.SolOut"
SolCallback(sLDropBank)    = "dtKNI.SolDropUp"
SolCallback(sRDropBank)    = "dtGHT.SolDropUp"
SolCallback(sBallPopper)   = "SolBallPopper"
'SolCallback(sLockPlunger)  = "SolLockPlunger" 'not used due to vlock
SolCallback(sLockPlunger)  = "Lock.SolExit" 'vlock
SolCallback(sREject)       = "SolREject"

'Solenoid CallBacks (C-Side)
SolCallback(sUPFRedBoltFlash)  = "SolRedBoltFlash"
SolCallback(sUPFBlueBoltFlash) = "SolBlueBoltFlash"
SolCallback(sKnightHeadFlash)  = "SolKnightHeadFlash"
SolCallback(sFlipLaneFlash)    = "SolFlipLaneFlash"
SolCallback(sDTFlash)          = "SolDTFlash"
SolCallback(sSkyWayFlash)      = "SolSkyWayFlash"
SolCallback(sREjectFlash)      = "SolREjectFlash"
SolCallback(sLockFlash)        = "SolLockFlash"

'Solenoid CallBacks (The Rest)
SolCallback(sIGIRelay)   = ""
SolCallback(sUPFGIRelay) = "SolUPFGIRelay"
SolCallback(sLPFGIRelay) = "SolLPFGIRelay"
SolCallback(sACselect)   = "SolACSelect"
SolCallback(sKickback)   = "SolKickBack"
SolCallback(sKnocker)    = "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(sMTargets)   = "SolMTargets"
SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------

'SWITCH HANDLING (Lower PF)

'Shooter Lane
Sub Shooter_Hit()       : Controller.Switch(swShooter) = 1 : End Sub
Sub Shooter_UnHit()     : Controller.Switch(swShooter) = 0 : End Sub
'In/Out Lanes
Sub LOutLane_Hit()	    : Controller.Switch(swLOutLane) = 1 : End Sub
Sub LOutLane_UnHit()	: Controller.Switch(swLOutLane) = 0 : End Sub
Sub LRetLane_Hit()		: Controller.Switch(swLRetLane) = 1 : End Sub
Sub LRetLane_UnHit()	: Controller.Switch(swLRetLane) = 0 : End Sub
Sub ROutLane_Hit()	    : Controller.Switch(swROutLane) = 1 : End Sub
Sub ROutLane_UnHit()	: Controller.Switch(swROutLane) = 0 : End Sub
Sub RRetLane_Hit()		: Controller.Switch(swRRetLane) = 1 : End Sub
Sub RRetLane_UnHit()	: Controller.Switch(swRRetLane) = 0 : End Sub
'Skyway
Sub Skyway_Hit()        : Controller.Switch(swLRampExit) = 1 : End Sub
Sub Skyway_UnHit()      : Controller.Switch(swLRampExit) = 0 : End Sub
'U-Turn
Sub UTurn1_Hit()        : Controller.Switch(swUTurn1) = 1 : If ActiveBall.VelY < -10 Then ActiveBall.VelY = -10 End If : End Sub
Sub UTurn1_UnHit()      : Controller.Switch(swUTurn1) = 0 : End Sub
Sub UTurn2_Hit()        : Controller.Switch(swUTurn2) = 1 : End Sub
Sub UTurn2_UnHit()      : Controller.Switch(swUTurn2) = 0 : End Sub
Sub UTurn3_Hit()        : Controller.Switch(swUTurn3) = 1 : If ActiveBall.VelY > 10 Then ActiveBall.VelY = 10 End If : End Sub
Sub UTurn3_UnHit()      : Controller.Switch(swUTurn3) = 0 : End Sub
Sub UTurn4_Hit()        : Controller.Switch(swUTurn4) = 1 : End Sub
Sub UTurn4_UnHit()      : Controller.Switch(swUTurn4) = 0 : End Sub
Sub REject_Hit()	    : bsREject.AddBall Me : End Sub

'SWITCH HANDLING (Upper PF)
'UPF Wire Ramp (Lock)
Sub URampEntry_Hit()    : Controller.Switch(swURampEntry) = 1 : PlaySoundAt "Wirerampsound", URampEntry  : End Sub
Sub URampEntry_UnHit()  : Controller.Switch(swURampEntry) = 0 : End Sub
Sub URampExit_Hit(): If ActiveBall.VelY > 0 Then StopSound Wirerampsound : vpmtimer.addtimer 500, "BalldropSound'" : End If : End Sub

'W-I-N Lanes
Sub Win1_Hit()          : Controller.Switch(swWin1) = 1 : End Sub
Sub Win1_UnHit()        : Controller.Switch(swWin1) = 0 : End Sub
Sub Win2_Hit()          : Controller.Switch(swWin2) = 1 : End Sub
Sub Win2_UnHit()        : Controller.Switch(swWin2) = 0 : End Sub
Sub Win3_Hit()          : Controller.Switch(swWin3) = 1 : End Sub
Sub Win3_UnHit()        : Controller.Switch(swWin3) = 0 : End Sub
'W-A-R Lanes
Sub War1_Hit()          : Controller.Switch(swWar1) = 1 : End Sub
Sub War1_UnHit()        : Controller.Switch(swWar1) = 0 : End Sub
Sub War2_Hit()          : Controller.Switch(swWar2) = 1 : End Sub
Sub War2_UnHit()        : Controller.Switch(swWar2) = 0 : End Sub
Sub War3_Hit()          : Controller.Switch(swWar3) = 1 : End Sub
Sub War3_UnHit()        : Controller.Switch(swWar3) = 0 : End Sub
'Loop
Sub Loop1_Hit()         : Controller.Switch(swLoop1) = 1 : End Sub
Sub Loop1_UnHit()       : Controller.Switch(swLoop1) = 0 : End Sub
Sub Loop2_Hit()         : Controller.Switch(swLoop2) = 1 : If ActiveBall.VelY > 10 Then ActiveBall.VelY = 5 End If:End Sub
Sub Loop2_UnHit()       : Controller.Switch(swLoop2) = 0 : End Sub
'Drawbridge Targets
Sub DBTrgt1_Slingshot() : vpmTimer.PulseSwitch (swBridge1), 0, "" : PlaysoundAtVol SoundFX("fx_target",DOFContactors),backbank,VolTarg: End Sub
Sub DBTrgt2_Slingshot() : vpmTimer.PulseSwitch (swBridge2), 0, "" : PlaysoundAtVol SoundFX("fx_target",DOFContactors),backbank,VolTarg: End Sub
Sub DBTrgt3_Slingshot() : vpmTimer.PulseSwitch (swBridge3), 0, "" : PlaysoundAtVol SoundFX("fx_target",DOFContactors),backbank,VolTarg: End Sub
'Bumpers
Sub Bumper1_Hit()		: vpmTimer.PulseSwitch (swJet1), 0, "" : PlaySoundAtVol SoundFX("Jet1",DOFContactors), Bumper1, VolBump : End Sub
Sub Bumper2_Hit()		: vpmTimer.PulseSwitch (swJet2), 0, "" : PlaySoundAtVol SoundFX("Jet2",DOFContactors), Bumper2, VolBump : End Sub
Sub Bumper3_Hit()		: vpmTimer.PulseSwitch (swJet3), 0, "" : PlaySoundAtVol SoundFX("Jet1",DOFContactors), Bumper3, VolBump : End Sub
'Drain
Sub Outhole_Hit()       : vpmTimer.PulseSwitch(swOutHole), 100, "HandleOutHole" : End Sub
Sub HandleOutHole(swNo) : bsTrough.AddBall Outhole : End Sub

'****************************************
'    JP Flippers with nFozzy additions
'****************************************

dim returnspeed, lfstep, rfstep
returnspeed = leftflipper.return
lfstep = 1
rfstep = 1

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFX("fx_flipperup",DOFContactors), LeftFlipper, VolFlip
        LeftFlipper.RotateToEnd
    Else
        PlaySoundAtVol SoundFX("fx_flipperdown",DOFContactors), LeftFlipper, VolFlip
        LeftFlipper.RotateToStart
		LeftFlipper.TimerEnabled = 1
        LeftFlipper.return = returnspeed * 0.5
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFX("fx_flipperup",DOFContactors), RightFlipper, VolFlip
        RightFlipper.RotateToEnd
		URightFlipper.RotateToEnd
    Else
        PlaySoundAtVol SoundFX("fx_flipperdown",DOFContactors), RightFlipper, VolFlip
        RightFlipper.RotateToStart
		URightFlipper.RotateToStart
		rightflipper.TimerEnabled = 1
        rightflipper.return = returnspeed * 0.5
		URightFlipper.return = returnspeed * 0.5
    End If
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, -0.1, 0.25
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, 0.1, 0.25
End Sub

sub leftflipper_timer()
	select case lfstep
		Case 1: leftflipper.return = returnspeed * 0.6 :lfstep = lfstep + 1
		Case 2: leftflipper.return = returnspeed * 0.7 :lfstep = lfstep + 1
		Case 3: leftflipper.return = returnspeed * 0.8 :lfstep = lfstep + 1
		Case 4: leftflipper.return = returnspeed * 0.9 :lfstep = lfstep + 1
		Case 5: leftflipper.return = returnspeed * 1.0 :lfstep = lfstep + 1
		Case 6: leftflipper.timerenabled = 0 : lfstep = 1
	end select
end sub

sub rightflipper_timer()
	select case rfstep
		Case 1: rightflipper.return = returnspeed * 0.6 : URightFlipper.return = returnspeed * 0.6 : rfstep = rfstep + 1
		Case 2: rightflipper.return = returnspeed * 0.7 : URightFlipper.return = returnspeed * 0.7 : rfstep = rfstep + 1
		Case 3: rightflipper.return = returnspeed * 0.8 : URightFlipper.return = returnspeed * 0.8 : rfstep = rfstep + 1
		Case 4: rightflipper.return = returnspeed * 0.9 : URightFlipper.return = returnspeed * 0.9 : rfstep = rfstep + 1
		Case 5: rightflipper.return = returnspeed * 1.0 : URightFlipper.return = returnspeed * 1.0 : rfstep = rfstep + 1
		Case 6: rightflipper.timerenabled = 0 : rfstep = 1
	end select
end sub

'******************************************************************************************
' Sling Shot Animations: Rstep and Lstep  are the variables that increment the animation
'******************************************************************************************

Dim RStep, Lstep

Sub RSling_Slingshot
	vpmTimer.PulseSwitch (swRSling), 0, ""
  	PlaySoundAtVol SoundFx("SlingshotRight",DOFContactors), sling1, 1
    RSling3.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RSling.TimerEnabled = 1
	'gi1.State = 0:Gi2.State = 0
End Sub

Sub RSling_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSling3.Visible = 1:sling1.TransZ = 0:RSling.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LSling_Slingshot
	vpmTimer.PulseSwitch (swLSling), 0, ""
    PlaySoundAtVol SoundFx("SlingshotLeft",DOFContactors), sling2, 1
    LSling3.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LSling.TimerEnabled = 1
End Sub

Sub LSling_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSling3.Visible = 1:sling2.TransZ = 0:LSling.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------

'Solenoid Callback handlers

Sub SolACSelect(enabled)
If Enabled Then
	PlaySound "fx_relay_on"
	StopSound "fx_relay_off"
Else
	PlaySound "fx_relay_off"
	StopSound "fx_relay_on"
End If
End Sub

Sub SolBallRelease(enabled)
	if enabled then
		if bsTrough.Balls then vpmTimer.PulseSwitch(swTroughEject),0,""
		bsTrough.ExitSol_On
	End if
End Sub

Sub SolREject(enabled)
	if enabled then
		bsREject.SolOut Enabled
	end if
end sub

Sub SolBallPopper(enabled)
	if enabled then
		bsBallPopper.SolOut Enabled
	end if
end sub

Sub SolLockPlunger(enabled) 'currently not used - using vlock instead :)
    if Enabled then
		LockPlunger.Fire
		PlaySoundAt SoundFX("autoplung",DOFContactors), plunger
	else
		LockPlunger.PullBack
	end If
End Sub

Sub SolKickBack(enabled)
    if Enabled then
		Playsound SoundFX("solenoid",DOFContactors)
		KickbackTimer.Interval = 500
    	KickbackTimer.Enabled = True
    	Kickback.Enabled = True
	End If
End Sub

'"FlipLaneFlashers" (LPF)
Sub SolFlipLaneFlash(enabled)
	if enabled then
		GI14.state = LightStateOn : GI14b.state = LightStateOn : GI15.state = LightStateOn : GI15b.state = LightStateOn
	Else
		GI14.state = LightStateOff : GI14b.state = LightStateOff : GI15.state = LightStateOff : GI15b.state = LightStateOff
	End If
End Sub

Sub SolDTFlash(enabled)
	if enabled then
		GI17.state = LightStateOn : GI17b.state = LightStateOn
		PegPlasticT18.DisableLighting = 1 : PegPlasticT17.DisableLighting = 1 : PegPlasticT16.DisableLighting = 1  : PegPlasticT15.DisableLighting = 1
	Else
		GI17.state = LightStateOff : GI17b.state = LightStateOff
		PegPlasticT18.DisableLighting = 0 : PegPlasticT17.DisableLighting = 0 : PegPlasticT16.DisableLighting = 0  : PegPlasticT15.DisableLighting = 0
	End If
End Sub


'Light-Flashers
Sub SolREjectFlash(enabled) : if enabled then GI24b.state = LightStateOn Else GI24b.state = LightStateOff End If : End Sub
Sub SolBlueBoltFlash(enabled) : if enabled then	Light20.State=LightStateOn:BlueBoltFlash.State=LightStateOn else Light20.State=LightStateOff: BlueBoltFlash.State=LightStateOff end if : End Sub
Sub SolRedBoltFlash(enabled) : if enabled then Light18.State=LightStateOn: RedBoltFlash.State=LightStateOn else Light18.State=LightStateOff: RedBoltFlash.State=LightStateOff end if : End Sub
Sub SolKnightHeadFlash(enabled): if enabled then KnightHeadFlash.State=LightStateOn	else KnightHeadFlash.State=LightStateOff end if: End Sub
Sub SolSkyWayFlash(enabled) : if enabled then SkyWayFlash.State=LightStateOn else SkyWayFlash.State=LightStateOff end if : End Sub
Sub SolLockFlash(enabled) : if enabled then	LockFlash.State=LightStateOn else LockFlash.State=LightStateOff end if : End Sub

'---------------------------------------
'      MagnaSave!
'---------------------------------------

Sub MSave_Hit(): MagnaSave.AddBall ActiveBall:End Sub
Sub MSave_UnHit(): MagnaSave.RemoveBall ActiveBall:End Sub

'******************************
'Motor Bank Up Down (from AFM)
'******************************

Sub Update3Bank(currpos, currspeed, lastpos)
    Dim zpos
    zpos = 135 - currpos
    If currpos <> lastpos Then BackBank.Z = zpos - 10 End If

    If currpos >= 49 Then
        DBTrgt1.Isdropped = 1 : DBTrgt2.Isdropped = 1 : DBTrgt3.Isdropped = 1
    End If

    If currpos = 0 Then
        DBTrgt1.Isdropped = 0 : DBTrgt2.Isdropped = 0 : DBTrgt3.Isdropped = 0
    End If

End Sub

Sub SolMTargets(enabled)
    If enabled Then
		PlaySound SoundFX("bk2k_drawbridge",DOFGear), -1
	Else
		StopSound "bk2k_drawbridge"
	End If
End Sub

' ### General Illumination-UPF ###
Sub SolUPFGIRelay(enabled)
    Dim Prim, Light
    If enabled Then
		PlaySound "fx_relay_off"
		StopSound "fx_relay_on"
		For Each Light in LightsGIupf : Light.State = LightStateOff : Next
		For Each Prim in primitivesGIupf : Prim.DisableLighting = 0 : Next
		For Each Prim in primitivesGIupfpegs : Prim.image = "peg" : Prim.material = "peg3": Next
		Primitive98.material = "RubberOff" : Primitive99.material = "RubberOff" : Primitive100.material = "RubberOff" : Primitive101.material = "RubberOff"
	Else
		PlaySound "fx_relay_on"
		StopSound "fx_relay_off"
		For Each Light in LightsGIupf : Light.State = LightStateOn : Next
		For Each Prim in primitivesGIupfpegs : Prim.image = "peglight" : Prim.material = "peg3light": Next
		For Each Prim in primitivesGIupf : Prim.DisableLighting = 1 : Next
		Primitive98.material = "RubberOn" : Primitive99.material = "RubberOn" : Primitive100.material = "RubberOn" : Primitive101.material = "RubberOn"
    End If
End Sub

' ### General Illumination-LPF ###
Sub SolLPFGIRelay(enabled)
	Dim Prim, Light
	If enabled Then
		PlaySound "fx_relay_off"
		StopSound "fx_relay_on"
		For Each Light in LightsGIlpf : Light.State = LightStateOff : Next
		For Each Prim in primitivesGIlpf : Prim.DisableLighting = 0 : Next
		For Each Prim in primitivesGIlpfpegs : Prim.image = "peg" : Prim.material = "peg3": Next
	Else
		PlaySound "fx_relay_on"
		StopSound "fx_relay_off"
		For Each Light in LightsGIlpf : Light.State = LightStateOn : Next
		For Each Prim in primitivesGIlpf : Prim.DisableLighting = 1 : Next
		For Each Prim in primitivesGIlpfpegs : Prim.image = "peglight" : Prim.material = "peg3light": Next
	End If
End Sub

'-----------------------------------------------------------------------------------
'WorkArounds
'-----------------------------------------------------------------------------------

Sub KickbackTimer_Timer
	' Turn off kickback after a certain period
	KickbackTimer.Enabled = False
	Kickback.Enabled = False
End Sub

Sub Kickback_Hit()
	' Kick ball back up, with random unpredictable force (like real life)
	Kickback.Kick 0, (33 + ((Rnd - 0.5) * 3)), 15
	'change the 33 to something else to raise or lower the kickback strength
	Kickback.Enabled = False
End Sub

Sub NoJump_Hit() : ActiveBall.VelZ = 0 : End Sub

sub MagnetHelper_unhit
	if MagnaSave.MagnetOn = 1 then	'if magnet is on...
		activeball.vely = activeball.vely * -0.25	'reverse direction of the ball
		activeball.velx = activeball.velx * -0.25
	end if
End Sub


'***************************************************
'    Lamps & Flashers
'***************************************************

LampTimer.Enabled = 1

Sub LampTimer_Timer()
    Dim chgLamp, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp) Then
        For ii = 0 To UBound(chgLamp)
			Select Case LightType(chgLamp(ii, 0))
				Case 0
					If LightObjectFirst(chgLamp(ii, 0)).state <> chgLamp(ii, 1) Then
						LightObjectFirst(chgLamp(ii, 0)).state = chgLamp(ii, 1) : LightObjectSecond(chgLamp(ii, 0)).state = chgLamp(ii, 1)
					End If
				Case 1 :
						If Light19.State <> chgLamp(ii, 1) Then
						Select Case chgLamp(ii, 1)
							Case 0
								Light19.state = 0:Light19b.state = 0:bulbyellow.Disablelighting = 0
								PegPlasticT23.Disablelighting = 0 : PegPlasticT23.image = "peg" : PegPlasticT23.material = "peg3"
								PegPlasticT24.Disablelighting = 0 : PegPlasticT24.image = "peg" : PegPlasticT24.material = "peg3"
							Case 1
								Light19.state = 1:Light19b.state = 1:bulbyellow.Disablelighting = 1
								PegPlasticT23.Disablelighting = 1 : PegPlasticT23.image = "peglight" : PegPlasticT23.material = "peg3light"
								PegPlasticT24.Disablelighting = 1 : PegPlasticT24.image = "peglight" : PegPlasticT24.material = "peg3light"
						 End Select
						End If
				Case 2 : Select Case chgLamp(ii, 1)
							Case 0: LightObjectFirst(chgLamp(ii, 0)).IntensityScale = 0.5
							Case 1: LightObjectFirst(chgLamp(ii, 0)).IntensityScale = 20.0
						 End Select
			End Select
        Next
    End If
End Sub

'******************************
' Change Ball appearance
'******************************

Sub ChangeBall(ballnr)
	Dim BOT, ii, col
	Bk2k.BallDecalMode = CustomBallLogoMode(ballnr)
	Bk2k.BallFrontDecal = CustomBallDecal(ballnr)
	Bk2k.DefaultBulbIntensityScale = CustomBulbIntensity(ballnr)
	Bk2k.BallImage = CustomBallImage(ballnr)
	GlowBall = CustomBallGlow(ballnr)
	For ii = 0 to 2
		col = RGB(red(ballnr), green(ballnr), Blue(ballnr))
		GlowLow(ii).color = col : GlowMed(ii).color = col : GlowHigh(ii).color = col
		GlowLow(ii).colorfull = col : GlowMed(ii).colorfull = col : GlowHigh(ii).colorfull = col
	Next
End Sub

'******************************
' Diverse Collection Hit Sounds
'******************************

' Rubber sounds
Sub Rubbers_Hit(idx):PlaySound "fx_rubber", 0, Vol(ActiveBall)*VolRB, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub

' Ramp Sounds
Sub Trigger1_Hit() : If ActiveBall.VelY > 0 Then PlaySoundAt "Wirerampsound", Trigger1 : End If : End Sub
Sub Trigger2_Hit() : StopSound Wirerampsound : vpmtimer.addtimer 250, "BalldropSound'" :End Sub
Sub Trigger3_Hit() : StopSound Wirerampsound : vpmtimer.addtimer 250, "BalldropSound'" :End Sub
Sub Trigger4_Hit() : StopSound Wirerampsound : vpmtimer.addtimer 250, "BalldropSound'" :End Sub
Sub Trigger5_Hit() : PlaySoundAt "Wirerampsound", Trigger5 : End Sub
Sub BalldropSound: Playsound "fx_ball_drop" : End Sub

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

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "BK2K" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / BK2K.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "BK2K" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / BK2K.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "BK2K" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / BK2K.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioFade(ball) ' Can this be together with the above function ?
  Dim tmp
  tmp = ball.y * 2 / BK2K.height-1
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

'*************************************************************
'      JP's VP10 Rolling Sounds with Flupper GlowBall addition
'*************************************************************

Const anglecompensate = 15
Const tnob = 10 ' total number of balls
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
		If GlowBall and b <3 Then GlowLow(b).state = 0 : GlowMed(b).state = 0 : GlowHigh(b).state = 0 End If
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
		If GlowBall Then
			If BOT(b).z > 160 Then
				If GlowHigh(b).state = 0 Then GlowLow(b).state = 0 : GlowMed(b).state = 0 : GlowHigh(b).state = 1 : End If
				GlowHigh(b).x = BOT(b).x : GlowHigh(b).y = BOT(b).y + anglecompensate
			Else
				If BOT(b).z > 50 Then
					If GlowMed(b).state = 0 Then GlowLow(b).state = 0 : GlowHigh(b).state = 0 : GlowMed(b).state = 1 : End If
					GlowMed(b).x = BOT(b).x : GlowMed(b).y = BOT(b).y + anglecompensate
				Else
					If GlowLow(b).state = 0 Then GlowMed(b).state = 0 : GlowHigh(b).state = 0 : GlowLow(b).state = 1 : End If
					GlowLow(b).x = BOT(b).x : GlowLow(b).y = BOT(b).y + anglecompensate
				End If
			End if
		End If
    Next
End Sub

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
  If BK2K.VersionMinor > 3 OR BK2K.VersionMajor > 10 Then
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
  Else
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0
  End if
End Sub

' Thalamus : Exit in a clean and proper way
Sub BK2K_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


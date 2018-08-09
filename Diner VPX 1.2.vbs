'*****************************************************************************************
'*****************************************************************************************
'****************************** Diner, Williams 1990 *************************************
'*****************************************************************************************

' Diner by Flupper
' Based on VP9 version by Tamoore which was based on Pac-Dude's version
' Playfield redraw by Bodydump
' Plastics redraws by Ben Logan
' Several plastics photos from George
' standup target by Dark
' many sounds from Knorr's sound package
' testing by Clark Kent and Ben Logan
' Shadow ball code by Ninuzzu

' Thalamus 2018-07-20
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' Thalamus 2018-08-09 : Improved directional sounds

Option Explicit
Randomize

Dim GlowAmountDay, InsertBrightnessDay, GlowAmountCutoutDay, InsertBrightnessCutoutDay
Dim GlowAmountNight, InsertBrightnessNight, GlowAmountCutoutNight, InsertBrightnessCutoutNight
Dim GlowBall, ChooseBall, CustomBulbIntensity(10), red(10), green(10), Blue(10)
Dim CustomBallImage(10), CustomBallLogoMode(10), CustomBallDecal(10), CustomBallGlow(10)
Dim ForceSiderailsFS, ChooseBats, LightsDemo, ShowClock, ContrastSetting, BallShadow
Dim GlobalSoundLevel, AlternateDroptargets

'******************************************************************************************
'* TABLE OPTIONS **************************************************************************
'******************************************************************************************

' *** 0=default flipper, 1=primitive flipper, 2=glow green, 3=glow blue, 4=glow orange ****
ChooseBats = 1

' *** Show siderails in fullscreen mode: True = show siderails, False = do not show *******
ForceSiderailsFS = False

' *** Show clock on back wall *************************************************************
ShowClock = False

' *** Contrast level, possible values are 0 - 7, can be done in game with magnasave keys **
' *** 0: bright, good for desktop view, daytime settings in insert lighting below *********
' *** 1: same as 0, but with nighttime settings in insert lighting below ******************
' *** 2: darker, better contrast, daytime settings in insert lighting below ***************
' *** 3: same as 2, but with nighttime settings in insert lighting below ******************
' *** etc for 4-7; default is 3 ***********************************************************
ContrastSetting = 3

' *** BallShadow : Ninnuzu's ball shadow addon ********************************************
BallShadow = True

' *** relative sound level for all mechanical sounds **************************************
' *** 1: if you have separate sound system for mech sounds, 2/3/etc when you do not *******
GlobalSoundLevel = 3

' *** Other decals for the droptargets ****************************************************
AlternateDroptargets = False

' *** Insert Lighting settings ************************************************************
' *** The settings below together with ContrastSetting determines how the lighting looks **
' *** for all values: 1.0 = default, useful range 0.1 - 5 *********************************
GlowAmountDay = 0.05
InsertBrightnessDay = 0.8
GlowAmountCutoutDay = 0
InsertBrightnessCutoutDay = 0.6
GlowAmountNight = 0.5
InsertBrightnessNight = 0.6
GlowAmountCutoutNight = 0.25
InsertBrightnessCutoutNight = 0.75

' *** Ball settings ***********************************************************************
' *** 0 = normal ball
' *** 1 = white GlowBall
' *** 2 = magma GlowBall
' *** 3 = blue GlowBall
' *** 4 = HDR ball
' *** 5 = earth
' *** 6 = green glowball (matching green GlowBat)
' *** 7 = blue glowball (matching blue GlowBat)
' *** 8 = orange glowball (matching orange GlowBat)
' *** 9 = shiny Ball

ChooseBall = 4

' default Ball
CustomBallGlow(0) = 		False
CustomBallImage(0) = 		"TTMMball"
CustomBallLogoMode(0) = 	False
CustomBallDecal(0) = 		"scratches"
CustomBulbIntensity(0) = 	0.01
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

' HDR ball
CustomBallGlow(4) = 		False
CustomBallImage(4) = 		"ball_HDR"
CustomBallLogoMode(4) = 	False
CustomBallDecal(4) = 		"JPBall-Scratches"
CustomBulbIntensity(4) = 	0.01
Red(4) = 0 : Green(4)	= 0 : Blue(4) = 0

' Earth
CustomBallGlow(5) = 		True
CustomBallImage(5) = 		"ballblack"
CustomBallLogoMode(5) = 	True
CustomBallDecal(5) = 		"earth"
CustomBulbIntensity(5) = 	0
Red(5) = 100 : Green(5)	= 100 : Blue(5) = 100

' green GlowBall
CustomBallGlow(6) = 		True
CustomBallImage(6) = 		"glowball green"
CustomBallLogoMode(6) = 	True
CustomBallDecal(6) = 		""
CustomBulbIntensity(6) = 	0
Red(6) = 100 : Green(6)	= 255 : Blue(6) = 100

' blue GlowBall
CustomBallGlow(7) = 		True
CustomBallImage(7) = 		"glowball blue"
CustomBallLogoMode(7) = 	True
CustomBallDecal(7) = 		""
CustomBulbIntensity(7) = 	0
Red(7) = 100 : Green(7)	= 100 : Blue(7) = 255

' orange GlowBall
CustomBallGlow(8) = 		True
CustomBallImage(8) = 		"glowball orange"
CustomBallLogoMode(8) = 	True
CustomBallDecal(8) = 		""
CustomBulbIntensity(8) = 	0
Red(8) = 255 : Green(8)	= 255 : Blue(8) = 100

' shiny Ball
CustomBallGlow(9) = 		False
CustomBallImage(9) = 		"pinball3"
CustomBallLogoMode(9) = 	False
CustomBallDecal(9) = 		"JPBall-Scratches"
CustomBulbIntensity(9) = 	0.01
Red(9) = 0 : Green(9)	= 0 : Blue(9) = 0

' *** Developer switch: Show all lights, no gameplay **************************************
LightsDemo = False

'******************************************************************************************
'* END OF TABLE OPTIONS *******************************************************************
'******************************************************************************************

' release notes

' version 1.0 : initial version
' version 1.1 : merged 2 lights-layers, improving performance by 8%
'				added flipperbats shadows for the primitive Bats
'				added flasher fade effect (afterglow on the flasher primitive)
'				several DOF related changes proposed by Arngrim
'				slingshot rom sound fix
'				reduced some texture sizes without visual impact
'				some small visual fixes
' version 1.2 : fix for GI not on at first game
'				fix for inserts showing only black
'				potential fix for GI not on (issue similar to black inserts)

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

' *** Global constants ***
Const UseSolenoids = 2
Const UseLamps = 0
Const UseSync = 0
Const HandleMech = 0
Const UseGI=0
Const cGameName="diner_l4"
Const sCoin = "fx_coin"
Const cCredits="Diner, Williams 1990"

' *** Global variables ***
Dim FlippersEnabled		' Used to enable/disable flippers based on tilt status
Dim Leftflash, Rightflash ' Used to detect which flasher is flashing
Dim CupCounter 'Used for influencing speed in the cup
Dim LRAindex, LRA(10) ' angle of lockramp
Dim DVindex, DV(10), DVdir ' angle of diverter
Dim OptionOpacity 'opacity of the option primive
OptionOpacity = 8
Dim relaylevel ' Sound level op relay clicking
relaylevel = 0.5 * GlobalSoundLevel
LRA(0) = -9 : LRA(1) = -6: LRA(2) = -1: LRA(3) = 3: LRA(4) = 5: LRA(5) = 2: LRA(6) = 0
DV(0) = 179 : DV(1) = 180 : DV(2) = 181 : DV(3) = 182 : DV(4) = 184 : DV(5) = 186 : DV(6) = 201
LRAindex = 0 : DVindex = 0 : DVdir = 0

' *** game specific constants ***
Const swOuthole=9,swUpDownRamp=10,swTrough1=11,swTrough2=12,swTrough3=13,swShooterLane=14,swSubPlfdShooter1=15
Const swSubPlfdShooter2=16,swCup=17,swGrillBonus=18,swE=19,swA=20,swT=21,swHotdog=22,swBurger=23,swChili=24
Const swRightRampEntry=27,swRightRampExit=28,swCupEntry=29,swRootBeer=30,swFries=31,swIcedTea=32
Const swLeftRampExit=36,swLeftOutlane=37,swLeftReturnLane=38,swRightReturnLane=39,swRightOutlane=40
Const swUpperLeftEject=49,swLowerLeftEject=50,swLeftJetBumper=51
Const swRightJetBumper=52,swLowerJetBumper=53,swBRKicker=54,swBLKicker=55,swSpinner=56,swFlipR=57,swFlipL=58,swClockWheel=59
Const sOuthole=1,sRampDown=2,sC3BDTReset=3,sRampUp=4,sUpperLeftEject=5,sSubPlfdShooter=6,sKnocker=7,sLowerLeftEject=8
Const sRightRampFlasher=9,sBackBoxRelay=10,sLeftRampFlasher=11,sL3BDTReset=13,sClockWheel=15
Const sACselectrelay=12, sLeftJetBumper=17,sRightJetBumper=19,sLowerJetBumper=21,sShooterLaneFeeder=22,sDineTF=32
Const sLSling=18,sRSling=20,swLSling=55,swRSling=54

' *** mechanic handlers ***
Dim dtright, dtleft, bsTrough, bsLowKicker, bsUpperEject, bsSubShooter, WheelMech, bsSubWay

' *** creating lights arrays for playfield inserts ***

' *** LightType 0 = normal insert lighting with two lights
' *** LightType 1 = one object only for visible on/off
' *** LightType 2 = EAT multiple Lights
' *** LightType 3 = cutout flashers (pepe, babs, boris, haji, buck)
' *** LightType 4 = not connected

Dim LightType(65), LightObjectFirst(65), LightObjectSecond(65), LightObjectThird(65), Glowing(10), LightState(65)

LightType(1)  = 1 : Set LightObjectFirst(1)  = register20k  								  		' 20k register
LightType(2)  = 1 : Set LightObjectFirst(2)  = register40k  								  		' 40k register
LightType(3)  = 1 : Set LightObjectFirst(3)  = register60k  								  		' 60k register
LightType(4)  = 1 : Set LightObjectFirst(4)  = register80k  								  		' 80k register
LightType(5)  = 1 : Set LightObjectFirst(5)  = register100k  								  		' 100k register
LightType(6)  = 0 : Set LightObjectFirst(6)  = Light6b  : Set LightObjectSecond(6)  = Light6  		' left adv dine time
LightType(7)  = 0 : Set LightObjectFirst(7)  = Light7b  : Set LightObjectSecond(7)  = Light7  		' right adv dine time
LightType(8)  = 0 : Set LightObjectFirst(8)  = Light8b  : Set LightObjectSecond(8)  = Light8  		' extra ball right
LightType(9)  = 0 : Set LightObjectFirst(9)  = Light9b  : Set LightObjectSecond(9)  = Light9  		' serve again
LightType(10)  = 0 : Set LightObjectFirst(10)  = Light10b  : Set LightObjectSecond(10)  = Light10  	' left ramp scores
LightType(11)  = 0 : Set LightObjectFirst(11)  = Light11b  : Set LightObjectSecond(11)  = Light11  	' right ramp scores
LightType(12)  = 0 : Set LightObjectFirst(12)  = Light12b  : Set LightObjectSecond(12)  = Light12  	' Lock
LightType(13)  = 0 : Set LightObjectFirst(13)  = Light13b  : Set LightObjectSecond(13)  = Light13  	' Release
LightType(14)  = 0 : Set LightObjectFirst(14)  = Light14b  : Set LightObjectSecond(14)  = Light14  	' Rush1
LightType(15)  = 0 : Set LightObjectFirst(15)  = Light15b  : Set LightObjectSecond(15)  = Light15  	' Rush2
LightType(16)  = 0 : Set LightObjectFirst(16)  = Light16b  : Set LightObjectSecond(16)  = Light16  	' Spinner
LightType(17)  = 0 : Set LightObjectFirst(17)  = Light17b  : Set LightObjectSecond(17)  = Light17  	' D
LightType(18)  = 0 : Set LightObjectFirst(18)  = Light18b  : Set LightObjectSecond(18)  = Light18  	' I
LightType(19)  = 0 : Set LightObjectFirst(19)  = Light19b  : Set LightObjectSecond(19)  = Light19  	' N
LightType(20)  = 0 : Set LightObjectFirst(20)  = Light20b  : Set LightObjectSecond(20)  = Light20  	' E
LightType(21)  = 0 : Set LightObjectFirst(21)  = Light21b  : Set LightObjectSecond(21)  = Light21  	' R
LightType(22)  = 2 : Set LightObjectFirst(22)  = Light22b  : Set LightObjectSecond(22)  = Light22 : Set LightObjectThird(22)  = circleE  	' E
LightType(23)  = 2 : Set LightObjectFirst(23)  = Light23b  : Set LightObjectSecond(23)  = Light23 : Set LightObjectThird(23)  = circleA 	' A
LightType(24)  = 2 : Set LightObjectFirst(24)  = Light24b  : Set LightObjectSecond(24)  = Light24 : Set LightObjectThird(24)  = circleT 	' T
LightType(25)  = 1 : Set LightObjectFirst(25)  = juke150k  										   	' 150k jukebox
LightType(26)  = 1 : Set LightObjectFirst(26)  = juke100k   									   	' 100k jukebox
LightType(27)  = 1 : Set LightObjectFirst(27)  = juke75k 										   	' 75k jukebox
LightType(28)  = 1 : Set LightObjectFirst(28)  = juke50k 										   	' 50k jukebox
LightType(29)  = 1 : Set LightObjectFirst(29)  = juke25k 										   	' 25k jukebox
LightType(30)  = 0 : Set LightObjectFirst(30)  = Light30b  : Set LightObjectSecond(30)  = Light30  	' hotdog
LightType(31)  = 0 : Set LightObjectFirst(31)  = Light31b  : Set LightObjectSecond(31)  = Light31  	' Burger
LightType(32)  = 0 : Set LightObjectFirst(32)  = Light32b  : Set LightObjectSecond(32)  = Light32  	' soup
LightType(33)  = 3 : Set LightObjectFirst(33)  = Light33b  : Set LightObjectSecond(33)  = Light33  	' Haji
LightType(34)  = 3 : Set LightObjectFirst(34)  = Light34b  : Set LightObjectSecond(34)  = Light34  	' Babs
LightType(35)  = 3 : Set LightObjectFirst(35)  = Light35b  : Set LightObjectSecond(35)  = Light35  	' Boris
LightType(36)  = 3 : Set LightObjectFirst(36)  = Light36b  : Set LightObjectSecond(36)  = Light36  	' Pepe
LightType(37)  = 3 : Set LightObjectFirst(37)  = Light37b  : Set LightObjectSecond(37)  = Light37  	' Buck
LightType(38)  = 0 : Set LightObjectFirst(38)  = Light38b  : Set LightObjectSecond(38)  = Light38  	' Beer
LightType(39)  = 0 : Set LightObjectFirst(39)  = Light39b  : Set LightObjectSecond(39)  = Light39  	' Fries
LightType(40)  = 0 : Set LightObjectFirst(40)  = Light40b  : Set LightObjectSecond(40)  = Light40  	' Lemonade
LightType(41)  = 0 : Set LightObjectFirst(41)  = Light41b  : Set LightObjectSecond(41)  = Light41  	' 100k
LightType(42)  = 0 : Set LightObjectFirst(42)  = Light42b  : Set LightObjectSecond(42)  = Light42  	' 150k
LightType(43)  = 0 : Set LightObjectFirst(43)  = Light43b  : Set LightObjectSecond(43)  = Light43  	' 250k
LightType(44)  = 0 : Set LightObjectFirst(44)  = Light44b  : Set LightObjectSecond(44)  = Light44  	' 1M
LightType(45)  = 0 : Set LightObjectFirst(45)  = Light45b  : Set LightObjectSecond(45)  = Light45  	' extra ball
LightType(46)  = 0 : Set LightObjectFirst(46)  = Light46b  : Set LightObjectSecond(46)  = Light46  	' food
LightType(47)  = 0 : Set LightObjectFirst(47)  = Light47b  : Set LightObjectSecond(47)  = Light47  	' cup scores
LightType(48)  = 0 : Set LightObjectFirst(48)  = Light48b  : Set LightObjectSecond(48)  = Light48  	' extra ball left
LightType(49)  = 4 : 																				' 49 - 60 are 1-12 on clock lights
LightType(50)  = 4 :
LightType(51)  = 4 :
LightType(52)  = 4 :
LightType(53)  = 4 :
LightType(54)  = 4 :
LightType(55)  = 4 :
LightType(56)  = 4 :
LightType(57)  = 4 :
LightType(58)  = 4 :
LightType(59)  = 4 :
LightType(60)  = 4 :
LightType(61)  = 0 : Set LightObjectFirst(61)  = Light61b  : Set LightObjectSecond(61)  = Light61  	' TOP 5
LightType(62)  = 0 : Set LightObjectFirst(62)  = Light62b  : Set LightObjectSecond(62)  = Light62  	' special
LightType(63)  = 0 : Set LightObjectFirst(63)  = Light63b  : Set LightObjectSecond(63)  = Light63  	' dinetime
LightType(64)  = 0 : Set LightObjectFirst(64)  = Light64b  : Set LightObjectSecond(64)  = Light64  	' food
LightType(65)  = 4 :

' *** prepare the variable with references to three lights for glow ball ***

Set Glowing(0) = Glowball1 : Set Glowing(1) = Glowball2 : Set Glowing(2) = Glowball3

Dim stat
Sub B2SSTAT
	If B2SOn Then
		Controller.B2SSetData 100+stat, 0
		stat = INT(rnd*6)+1
		Controller.B2SSetData 100+stat, 1
	End If
End Sub

' *** Start VPM ***

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0
LoadVPM "01560000", "S11.VBS", 3.26

' MotorCallback	 Function called after every update of solenoids and lamps (i.e. as often as
'				 possible). Main purpose is to update table based on playfield motors. It can also
'				 be used to override the standard solenoid callback handler
Set MotorCallback = GetRef("RollingUpdate") 'realtime updates for rolling sound

'*** solenoid definitions ***

SolCallback(23)="TiltSol"
SolCallback(sOuthole)="bsTrough.SolIn"
SolCallback(sShooterLaneFeeder)="bsTrough.SolOut"
SolCallback(sKnocker) = "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(sLeftJetBumper)="vpmSolSound SoundFX(""bumperleft"",DOFContactors),"
SolCallback(sRightJetBumper)="vpmSolSound SoundFX(""bumperright"",DOFContactors),"
SolCallback(sLowerJetBumper)="vpmSolSound SoundFX(""bumpermiddle"",DOFContactors),"
SolCallback(sLowerLeftEject) = "SolLowKicker"
SolCallback(sUpperLeftEject) = "SolUpperEject"
SolCallback(sRightRampFlasher) 	= "SolFlash9"
SolCallback(sLeftRampFlasher)   = "SolFlash11"
SolCallback(sSubPlfdShooter) = "SolSubPlfdShooter"
SolCallback(25) = "SolHajiF"
SolCallback(26) = "SolBabsF"
SolCallback(27) = "SolBorisF"
SolCallback(28) = "SolPepeF"
SolCallback(29) = "SolBuckF"
SolCallback(30) = "SolCupF"
SolCallback(sBackBoxRelay) = "SolGIRelay"
SolCallback(sACselectrelay)   = "SolACSelect"
SolCallback(sKnocker)    = "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(sC3BDTReset) = "dtright.SolDropUp"
SolCallback(sL3BDTReset) = "dtleft.SolDropUp"
SolCallback(sRampUp)="SolRampUp"
SolCallback(sRampDown)="SolRampDown"
SolCallback(14)="Sol14Diverter"

'*** flipper solenoids disabled for faster response in KeyDown ***
'SolCallback(sLRFlipper) = "SolRFlipper"
'SolCallback(sLLFlipper) = "SolLFlipper"

' *** table init ***

Sub Diner_Init
	Dim Light, Prim
	vpmInit Me
	If LightsDemo Then
		For Each Light in GlowLights : Light.state = 1 : Next
		For Each Light in InsertLights : Light.state = 1 : Next
		For Each Light in glowlightscutout : Light.state = 1 : Next
		For Each Light in insertlightscutout : Light.state = 1 : Next
		For Each Light in LightsGI : Light.State = LightStateOn : Next
		For Each Prim in primitivesGI : Prim.DisableLighting = 1 : Next
		For Each Prim in primitivesGIpegs : Prim.image = "peglight" : Prim.material = "peg3light": Next
		'SolFlash11(True)
		'SolFlash9(true)
		'solCupF(True)
	Else
		With Controller
			' *** explanation of controller properties
			' .HandleKeyboard	if set to 1, VPinMAME will process the keyboard. Standard MAME
			'					keys can be used in VPinMAME output window. Also the internal
			'					ball simulator is enabled. If set to 0, the scripting language
			'					will process the keyboard.
			' ShowTitle			if set to 1, Show title bar on VPinMAME output window (to move
			'					it around)
			' ShowDMDOnly		if set to 1, does not show VPinMAME status matrices.
			' ShowFrame			if set to 1, Enables the window border, only works if ShowTitle
			'					is set to 0
			' SplashInfoLine	Game credits to display in startup splash screen.
			' Hidden			If set to 1, hides the vpinmame display. This is useful for
			'					applications that want to do the display drawing on their own.
			' HandleMechanics	If the game has a PinMAME simulator it can be used to simulate
			'					hardware "toys".
			.GameName = cGameName
			If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
			.SplashInfoLine = "Diner - Williams 1990" & vbNewLine & "VPX by Flupper"
			.Games(cGameName).Settings.Value("rol") = 0 'set it to 1 to rotate the DMD to the left
			.Games(cGameName).Settings.Value("ror") = 0
			.HandleKeyboard = 0
			.ShowTitle = 0
			.ShowDMDOnly = 1
			.ShowFrame = 0
			.HandleMechanics = 0
			.Hidden = 1
			On Error Resume Next
			.Run GetPlayerHWnd
			If Err Then MsgBox Err.Description
			On Error Goto 0
		End With
	End if

	'*** Nudging ***
    vpmNudge.TiltSwitch = 1
    vpmNudge.Sensitivity = 1
	vpmNudge.TiltObj = Array(LeftJetBumper,RightJetBumper,LowerJetBumper,LSling,RSling)

	' *** Main Timer init ***
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

	'*** option settings ***
	If ContrastSetting=0 or ContrastSetting=2 or ContrastSetting=4 or ContrastSetting=6 Then ChangeGlow(True) Else ChangeGlow(False) End If
	ChangeBats(ChooseBats)
	ChangeBall(ChooseBall)
	If Diner.ShowDT or ForceSiderailsFS then sidewallFS.visible = 0 : sidewallDT.visible = 1 Else lockbar.Visible = 0 : sidewallDT.visible = 0 : sidewallFS.visible = 1 : End If
	If ShowClock Then clock.visible = 1 : clockwijzer.visible = 1 : end If
	If GlowBall or BallShadow Then GraphicsTimer.enabled = True End If
	IF not AlternateDroptargets Then
		hotdog.image = "DropTargetyellow" : burger.image = "DropTargetyellow" : chili.image = "DropTargetyellow"
		icedtea.image = "DropTargetred" : rootbeer.image = "DropTargetred" : fries.image = "DropTargetred"
	End If

	' *** Desktop specific changes for better 3D view ***
	If Diner.ShowDT then
		Flasher4light.BulbHaloHeight = 250
		rampsDT.visible = 1
		rampsFS.visible = 0
	Else
		rampsDT.visible = 0
		rampsFS.visible = 1
	end If

	' *** mechanic for clock hand ***
	Set WheelMech = New cvpmMech
	With WheelMech :
		.MType = vpmMechStepSol + vpmMechCircle + vpmMechLinear + vpmMechFast
		.Sol1 = 15 : .Sol2 = 16 : .Length = 200 : .Steps = 60
		.CallBack = GetRef("UpdateClockHand")
		.AddSw 59, 0, 0
		.Start
	End With

    ' *** Trough ***
    Set bsTrough = New cvpmTrough
    With bsTrough
		.size = 3
		.initSwitches Array(swTrough1,swTrough2,swTrough3)
        .Initexit BallRelease, 90, 4
        .InitEntrySounds "drain", "", ""
        .InitExitSounds SoundFX("fx_Solenoid",DOFContactors), SoundFX("fx_Ballrel",DOFContactors)
        .Balls = 3
		.EntrySw = 10
    End With

 	' *** Left droptargets Bank ***
	Set dtleft = New cvpmDropTarget
	With dtleft
        .InitDrop Array(rootbeer, fries, icedtea), Array(30, 31, 32)
        .initsnd SoundFX("fx_droptarget", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
        .CreateEvents "dtleft"
    End With

	' *** Right droptargets Bank ***
	Set dtright = New cvpmDropTarget
	With dtright
        .InitDrop Array(hotdog, burger, chili), Array(22, 23, 24)
        .initsnd SoundFX("fx_droptarget", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
        .CreateEvents "dtright"
    End With

' *** lock ramp saucer ***
	Set bsLowKicker = New cvpmSaucer
	With bsLowKicker
		.InitKicker LowKicker, swLowerLeftEject, 45, 6, 0
		.InitSounds "kicker_enter", SoundFX("popper_ball",DOFContactors), SoundFX("popper_ball",DOFContactors)
	End With

	' *** upper left saucer ***
	Set bsUpperEject = New cvpmSaucer
	With bsUpperEject
		.InitKicker UpperEject, swUpperLeftEject, 90, 6 , 0
		.InitExitVariance 0,2
		.InitSounds "kicker_enter", SoundFX("popper_ball",DOFContactors), SoundFX("popper_ball",DOFContactors)
	End With

	' *** subway kicker ***
	Set bsSubWay = New cvpmTrough
    With bsSubWay
		.size = 2
		.initSwitches Array(swSubPlfdShooter1,swSubPlfdShooter2)
        .Initexit SubWaypopper, 0, 100
        .InitEntrySounds "WireRamp_Hit", "", ""
        .InitExitSounds SoundFX("fx_Solenoid",DOFContactors), SoundFX("BallRelease",DOFContactors)
        .Balls = 0
		.EntrySw = 0
		.CreateEvents "bsSubWay", SubWaypopper
    End With

	' *** workaround: on some setups, the GI is not by default on ***
	SolGIRelay(False)

End Sub

'*** change insert glow appearance ***

Sub ChangeGlow(day)
	Dim Light
	If day Then
		For Each Light in GlowLights : Light.IntensityScale = GlowAmountDay: Light.FadeSpeedUp = Light.Intensity * GlowAmountDay / 2: Light.FadeSpeedDown = Light.FadeSpeedUp / 25: Next
		For Each Light in InsertLights : Light.IntensityScale = InsertBrightnessDay : Light.FadeSpeedUp = Light.Intensity * InsertBrightnessDay / 2 : Light.FadeSpeedDown = Light.FadeSpeedUp / 25 : Next
		For Each Light in glowlightscutout : Light.IntensityScale = GlowAmountCutoutDay : Light.FadeSpeedUp = Light.Intensity /  2 : Light.FadeSpeedDown = Light.FadeSpeedUp / 25 : Next
		For Each Light in insertlightscutout : Light.IntensityScale = InsertBrightnessCutoutDay : Light.FadeSpeedUp = Light.Intensity / 2 : Light.FadeSpeedDown = Light.FadeSpeedUp / 25: Next
		For Each Light in eaters: Light.FadeSpeedUp = 5000 : Light.FadeSpeedDown = 5000 : Next
	Else
		For Each Light in GlowLights : Light.IntensityScale = GlowAmountNight: Light.FadeSpeedUp = Light.Intensity * GlowAmountNight / 2: Light.FadeSpeedDown = Light.FadeSpeedUp / 25: Next
		For Each Light in InsertLights : Light.IntensityScale = InsertBrightnessNight : Light.FadeSpeedUp = Light.Intensity * InsertBrightnessNight / 2 : Light.FadeSpeedDown = Light.FadeSpeedUp / 25 : Next
		For Each Light in glowlightscutout : Light.IntensityScale = GlowAmountCutoutNight : Light.FadeSpeedUp = Light.Intensity /  2 : Light.FadeSpeedDown = Light.FadeSpeedUp / 25 : Next
		For Each Light in insertlightscutout : Light.IntensityScale = InsertBrightnessCutoutNight : Light.FadeSpeedUp = Light.Intensity / 2 : Light.FadeSpeedDown = Light.FadeSpeedUp / 25: Next
		For Each Light in eaters: Light.FadeSpeedUp = 5000 : Light.FadeSpeedDown = 5000 : Next
	End If
End Sub

'*** change bat appearance ***

Sub ChangeBats(Bats)
	Select Case Bats
		Case 0
			glowbatleft.visible = 0 : glowbatright.visible = 0 : GlowBatLightLeft.visible = 0 : GlowBatLightRight.visible = 0
			batleft.visible = 0 : batright.visible = 0 : LeftFlipper.visible = 1 : RightFlipper.visible = 1
			batleftshadow.visible = 0 : batrightshadow.visible = 0 : GraphicsTimer.enabled = False
		Case 1
			glowbatleft.visible = 0 : glowbatright.visible = 0 : GlowBatLightLeft.visible = 0 : GlowBatLightRight.visible = 0
			batleft.visible = 1 : batright.visible = 1 : LeftFlipper.visible = 0 : RightFlipper.visible = 0
			batleftshadow.visible = 1 : batrightshadow.visible = 1 : GraphicsTimer.enabled = True
		Case 2
			glowbatleft.visible = 1 : glowbatright.visible = 1 : GlowBatLightLeft.visible = 1 : GlowBatLightRight.visible = 1
			batleft.visible = 0 : batright.visible = 0 : LeftFlipper.visible = 0 : RightFlipper.visible = 0
			glowbatleft.image = "glowbat green" : glowbatright.image = "glowbat green"
			GlowBatLightLeft.color = RGB(0,255,0) : GlowBatLightRight.color = RGB(0,255,0)
			batleftshadow.visible = 0 : batrightshadow.visible = 0 : GraphicsTimer.enabled = True
		Case 3
			glowbatleft.visible = 1 : glowbatright.visible = 1 : GlowBatLightLeft.visible = 1 : GlowBatLightRight.visible = 1
			batleft.visible = 0 : batright.visible = 0 : LeftFlipper.visible = 0 : RightFlipper.visible = 0
			glowbatleft.image = "glowbat blue" : glowbatright.image = "glowbat blue"
			GlowBatLightLeft.color = RGB(0,0,255) : GlowBatLightRight.color = RGB(0,0,255)
			batleftshadow.visible = 0 : batrightshadow.visible = 0 : GraphicsTimer.enabled = True
		Case 4
			glowbatleft.visible = 1 : glowbatright.visible = 1 : GlowBatLightLeft.visible = 1 : GlowBatLightRight.visible = 1
			batleft.visible = 0 : batright.visible = 0 : LeftFlipper.visible = 0 : RightFlipper.visible = 0
			glowbatleft.image = "glowbat orange" : glowbatright.image = "glowbat orange"
			GlowBatLightLeft.color = RGB(255,0,0) : GlowBatLightRight.color = RGB(255,0,0)
			batleftshadow.visible = 0 : batrightshadow.visible = 0 : GraphicsTimer.enabled = True
	End Select
End Sub

'*** change ball appearance ***

Sub ChangeBall(ballnr)
	Dim BOT, ii, col
	Diner.BallDecalMode = CustomBallLogoMode(ballnr)
	Diner.BallFrontDecal = CustomBallDecal(ballnr)
	Diner.DefaultBulbIntensityScale = CustomBulbIntensity(ballnr)
	Diner.BallImage = CustomBallImage(ballnr)
	GlowBall = CustomBallGlow(ballnr)
	For ii = 0 to 2
		col = RGB(red(ballnr), green(ballnr), Blue(ballnr))
		Glowing(ii).color = col : Glowing(ii).colorfull = col
	Next
End Sub

' *** various hit event handling ***

Sub ShooterLane_Hit():B2SSTAT: Controller.Switch(swShooterLane)=1:End Sub
Sub ShooterLane_UnHit():Controller.Switch(swShooterLane)=0:End Sub
Sub LeftJetBumper_Hit():B2SSTAT: vpmTimer.PulseSwitch (swLeftJetBumper), 0, "" : PlaySoundAt SoundFX("Jet1",DOFContactors),LeftJetBumper : End Sub
Sub RightJetBumper_Hit():B2SSTAT: vpmTimer.PulseSwitch (swRightJetBumper), 0, "" : PlaySoundAt SoundFX("Jet2",DOFContactors),RightJetBumper : End Sub
Sub LowerJetBumper_Hit():B2SSTAT: vpmTimer.PulseSwitch (swLowerJetBumper), 0, "" : PlaySoundAt SoundFX("Jet1",DOFContactors),LowerJetBumper : End Sub
Sub LeftReturnLane_Hit():Controller.Switch(swLeftReturnLane)=1:End Sub
Sub LeftReturnLane_UnHit():Controller.Switch(swLeftReturnLane)=0:End Sub
Sub RightReturnLane_Hit():Controller.Switch(swRightReturnLane)=1:End Sub
Sub RightReturnLane_UnHit():Controller.Switch(swRightReturnLane)=0:End Sub
Sub LeftOutlane_Hit():Controller.Switch(swLeftOutLane)=1:End Sub
Sub LeftOutlane_UnHit():Controller.Switch(swLeftOutLane)=0:End Sub
Sub RightOutlane_Hit():Controller.Switch(swRightOutlane)=1:End Sub
Sub RightOutlane_UnHit():Controller.Switch(swRightOutlane)=0:End Sub
Sub E_Hit():B2SSTAT: Controller.Switch(swE)=1:End Sub
Sub E_UnHit():Controller.Switch(swE)=0:End Sub
Sub A_Hit():B2SSTAT: Controller.Switch(swA)=1:End Sub
Sub A_UnHit():Controller.Switch(swA)=0:End Sub
Sub T_Hit():B2SSTAT: Controller.Switch(swT)=1:End Sub
Sub T_UnHit():Controller.Switch(swT)=0:End Sub
Sub grillbonus_Hit():VpmTimer.PulseSw swGrillBonus:PlaySound "target":End Sub
Sub Spinner_Spin():vpmTimer.PulseSwitch(swSpinner),0,"":End Sub
Sub RightRampEntry_Hit() : Controller.Switch(swRightRampEntry)=1:End Sub
Sub RightRampEntry_UnHit():Controller.Switch(swRightRampEntry)=0:End Sub
Sub LeftRampExit_Hit():Controller.Switch(swLeftRampExit)=1:End Sub
Sub LeftRampExit_UnHit():Controller.Switch(swLeftRampExit)=0:End Sub
Sub RightRampExit_Hit():Controller.Switch(swRightRampExit)=1:End Sub
Sub RightRampExit_UnHit():Controller.Switch(swRightRampExit)=0:End Sub
Sub UpperEject_Hit() : bsUpperEject.AddBall Me : End Sub
Sub LowKicker_Hit() : bsLowKicker.AddBall Me : End Sub
Sub subwayenter_Hit() : PlaySound "NeutralZone2" : End Sub
Sub Outhole_Hit() : vpmTimer.PulseSwitch(swOutHole), 100, "HandleOutHole" : End Sub
Sub HandleOutHole(swNo) : bsTrough.AddBall Outhole : End Sub
Sub TiltSol(Enabled) : FlippersEnabled = Enabled : SolLFlipper(False) : SolRFlipper(False) : End Sub
Sub CupExit_Hit() : Playsound "fx_ball_drop" : CupExit.Enabled = 0 : End Sub
Sub RightRampDrop_Hit:StopSound "metalrolling": Playsound "fx_ball_drop":End Sub
Sub LeftRampDrop_Hit:Stopsound "plasticrolling": Playsound "fx_ball_drop": ActiveBall.VelY = 0 : End Sub
Sub UpdateClockHand(aNewPos, aSpeed, aLastPos) : clockwijzer.ObjrotY = anewpos * 6 : DOF 101, DOFPulse : end Sub
Sub SolUpperEject(enabled) : if enabled then bsUpperEject.SolOut Enabled : end if : end sub 'Solout: Fire the exit kicker and ejects ball if one is present.
Sub SolLowKicker(enabled) : if enabled then bsLowKicker.SolOut Enabled : end if : end sub 'Solout: Fire the exit kicker and ejects ball if one is present.

'*** Cup handling ***

dim CupEntrySpeed

Sub CupEntry_Hit()
	vpmTimer.PulseSw swCupEntry
	Stopsound "plasticrolling"
	CupExit.Enabled = 1 : CupCounter = 30 : cupblocker.collidable = False
	ActiveBall.VelZ = 0
	CupEntrySpeed =  sqr(ActiveBall.VelX ^2 + ActiveBall.VelY ^2) * 1.25

End Sub

Sub Cup_Hit() : Cupspeeder : vpmTimer.PulseSw swCup : cupblocker.collidable = True : End Sub
Sub Cup2_Hit() : Cupspeeder : end sub
Sub Cup_UnHit() : Cupspeeder : end sub
Sub Cup2_UnHit() : Cupspeeder : end sub

Sub Cupspeeder()
	If abs(ActiveBall.VelY) > 5 And abs(ActiveBall.VelX) > 5 Then
		Dim BallSpeed, BallSpeedReduction
		BallSpeed = sqr(ActiveBall.VelX ^2 + ActiveBall.VelY ^2)
		BallSpeedReduction = (CupEntrySpeed / BallSpeed) * 30 / ( sqr(2) * CupCounter )
		ActiveBall.VelY =  ActiveBall.VelY * BallSpeedReduction
		ActiveBall.VelX =  ActiveBall.VelX * BallSpeedReduction
		cupcounter = cupcounter + 1
	End If
End Sub

'*** Lockramp handling ***

Sub SolRampDown(Enabled)
	If ramp1.collidable Then
		Controller.Switch(swUpDownRamp)=1
	Else
		ramp1.collidable = True : LRAindex = 6 : lockramptimer.enabled = True
	End If
End Sub

Sub SolRampUp(Enabled)
	If Ramp1.collidable Then
		ramp1.collidable = False : lockramptimer.enabled = True : LRAindex = 0
	Else
		Controller.Switch(swUpDownRamp)=0
	End If
End Sub

Sub lockramptimer_timer
	If ramp1.collidable Then
		lockramp.ObjRotX = LRA(LRAindex) : LRAindex = LRAindex - 1
		If LRAindex < 0 then LRAindex = 0 : lockramptimer.enabled = False : Controller.Switch(swUpDownRamp)=1 : playsound"TrapDoorLow" : end If
	Else
		lockramp.ObjRotX = LRA(LRAindex) : LRAindex = LRAindex + 1
		If LRAindex > 6 then LRAindex = 6 : lockramptimer.enabled = False : Controller.Switch(swUpDownRamp)=0 : playsound"TrapDoorHigh" : end If
	End If
end Sub

'*** diverter handling ***

Sub Sol14Diverter(Enabled)
	if Enabled then
		Sol14Closed.collidable = False : DVdir = 1 : divertertimer.enabled = True : playsound SoundFX("TopDiverterOn",DOFContactors)
	else
		Sol14Closed.collidable = True : Dvdir = -1 : divertertimer.enabled = True : playsound SoundFX("TopDiverterOff",DOFContactors)
	end if
End Sub

Sub divertertimer_timer
	If DVdir = 0 Then
		divertertimer.enabled = False
	Else
		DVindex = DVindex + DVdir
		If DVindex < 0 Then DVindex = 0 : DVdir = 0 : end If
		IF DVindex > 6 Then Dvindex = 6:  DVdir = 0 : end If
		divider179to201.RotY = DV(DVindex)
	End If
end Sub

' *** drophole sub playfield shooter ***

Sub SolSubPlfdShooter(enabled)
    If enabled and bsSubWay.balls > 0 Then
		bsSubWay.ExitSol_On : rightflap.RotX = 80 : rightflaptimer.enabled = True
	End If
End Sub

Sub rightflaptimer_timer : rightflap.RotX = 90 : rightflaptimer.enabled = False : end Sub

' *** character flasher handling ***

Sub solHajiF(enabled) :  FlashCutout 33,enabled : End Sub
Sub solBabsF(enabled) :  FlashCutout 34,enabled : End Sub
Sub solBorisF(enabled) : FlashCutout 35,enabled : End Sub
Sub solPepeF(enabled) :  FlashCutout 36,enabled : End Sub
Sub solBuckF(enabled) :  FlashCutout 37,enabled : End Sub

Sub FlashCutout(nr,enabled)
	Dim GlowCutout, InsertCutout
	If ContrastSetting=0 or ContrastSetting=2 or ContrastSetting=4 or ContrastSetting=6 Then
		GlowCutout = GlowAmountCutoutDay : InsertCutout = InsertBrightnessCutoutDay
	Else
		GlowCutout = GlowAmountCutoutNight : InsertCutout = InsertBrightnessCutoutNight
	End If
	if enabled then
		LightObjectFirst(nr).IntensityScale = GlowCutout * 4 : LightObjectFirst(nr).state = LightStateOn
		LightObjectSecond(nr).IntensityScale = InsertCutout * 12 : LightObjectSecond(nr).state = LightStateOn
	Else
		LightObjectFirst(nr).IntensityScale = GlowCutout : LightObjectFirst(nr).state = LightState(nr)
		LightObjectSecond(nr).IntensityScale = InsertCutout : LightObjectSecond(nr).state = LightState(nr)
	End If
End Sub

' *** cup Flasher ***

Sub solCupF(enabled)
    if enabled then
		PlaySound "fx_relay_on",0,relaylevel : StopSound "fx_relay_off"
		Flasher5light.state = LightStateOn : Flasher5lightsec.state = LightStateOn
	Else
		PlaySound "fx_relay_off",0,relaylevel : StopSound "fx_relay_on"
		Flasher5light.state = LightStateOff : Flasher5lightsec.state = LightStateOff
	End If
End Sub

' *** Rightrampflasher ***

Sub SolFlash9(Enabled)
	Rightflash = enabled
	FlashTimerRight.enabled = False
	if enabled then
		PlaySound "fx_relay_on",0,relaylevel : StopSound "fx_relay_off": Flasher1.disablelighting = True : Flasher2.disablelighting = True
		Flasher1light.IntensityScale = 1 : Flasher2light.IntensityScale = 1 : Flasher1lightsec.IntensityScale = 1 : Flasher2lightsec.IntensityScale = 1
		Flasher1light.state = LightStateOn : Flasher2light.state = LightStateOn : Flasher1lightsec.state = LightStateOn : Flasher2lightsec.state = LightStateOn
		Flasher1.material = "flasherlit" : Flasher1.image = "domelighton" :Flasher2.material = "flasherlit" : Flasher2.image = "domelighton"
		If Leftflash Then
			sidewallFS.image = "sidewallsFSbothflash" : backwall.image = "backwallbothflash" : sidewallDT.image = "sidewallsFSbothflash"
			sidewallFS.material = "sidewallsbothflash" : sidewallDT.material = "sidewallsbothflash"
			ColorGradeFlash()
		Else
			sidewallFS.image = "sidewallsFSrightflash" : backwall.image = "backwallrightflash" : sidewallDT.image = "sidewallsFSrightflash"
			sidewallFS.material = "sidewallsrightflash" : sidewallDT.material = "sidewallsrightflash"
			ColorGradeFlash()
		End If
	Else
		PlaySound "fx_relay_off",0,relaylevel : StopSound "fx_relay_on"
		Flasher1light.IntensityScale = 0.5 : Flasher2light.IntensityScale = 0.5 : Flasher1lightsec.IntensityScale = 0.5 : Flasher2lightsec.IntensityScale = 0.5
		Flasher1light.state = LightStateOff : Flasher2light.state = LightStateOff : Flasher1lightsec.state = LightStateOff : Flasher2lightsec.state = LightStateOff
		'Flasher1.material = "flasherunlit" : Flasher1.image = "domelightoff" : Flasher2.material = "flasherunlit" : Flasher2.image = "domelightoff"
		'Flasher1.disablelighting = False : Flasher2.disablelighting = False
		Flasher1.material = "flasherunlit" : Flasher2.material = "flasherunlit" : FlashStepRight = 0 : FlashTimerRight.enabled = True
		If Leftflash Then
			sidewallFS.image = "sidewallsFSleftflash" : backwall.image = "backwallleftflash" : sidewallDT.image = "sidewallsFSleftflash"
			sidewallFS.material = "sidewallsleftflash" : sidewallDT.material = "sidewallsleftflash"
			ColorGradeFlash()
		Else
			sidewallFS.image = "sidewallsFS" : backwall.image = "backwallnoflash" : sidewallDT.image = "sidewallsFS"
			sidewallFS.material = "sidewallsnoflash" : sidewallDT.material = "sidewallsnoflash"
			ColorGradeFlash()
		End If

	End If
End sub

Dim FlashStepRight

Sub FlashTimerRight_timer()
	Flasher1.image = "flasherfade" & FlashStepRight
	Flasher2.image = "flasherfade" & FlashStepRight
	FlashStepRight = FlashStepRight + 1
	If FlashStepRight = 6 Then
		Flasher1.image = "domelightoff" : Flasher2.image = "domelightoff"
		Flasher1.disablelighting = False : Flasher2.disablelighting = False
		FlashTimerRight.enabled = False
	End If
End Sub

' *** Leftrampflasher ***

Sub SolFlash11(Enabled)
	Leftflash = enabled
	FlashTimerLeft.enabled = False
	if enabled then
		PlaySound "fx_relay_on",0,relaylevel : StopSound "fx_relay_off": Flasher3.disablelighting = True : Flasher4.disablelighting = True
		Flasher3light.IntensityScale = 1 : Flasher4light.IntensityScale = 1 : Flasher3lightsec.IntensityScale = 1 : Flasher4lightsec.IntensityScale = 1
		Flasher3light.state = LightStateOn : Flasher4light.state = LightStateOn : Flasher3lightsec.state = LightStateOn : Flasher4lightsec.state = LightStateOn
		Flasher3.material = "flasherlit" : Flasher3.image = "domelighton" :Flasher4.material = "flasherlit" : Flasher4.image = "domelighton"
		If Rightflash Then
			sidewallFS.image = "sidewallsFSbothflash" : backwall.image = "backwallbothflash" : sidewallDT.image = "sidewallsFSbothflash"
			sidewallFS.material = "sidewallsbothflash" : sidewallDT.material = "sidewallsbothflash"
			ColorGradeFlash()
		Else
			sidewallFS.image = "sidewallsFSleftflash" : backwall.image = "backwallleftflash" : sidewallDT.image = "sidewallsFSleftflash"
			sidewallFS.material = "sidewallsleftflash" : sidewallDT.material = "sidewallsleftflash"
			ColorGradeFlash()
		End If
	Else
		PlaySound "fx_relay_off",0,relaylevel : StopSound "fx_relay_on": 'Flasher3.disablelighting = False : Flasher4.disablelighting = False
		Flasher3light.IntensityScale = 0.5 : Flasher4light.IntensityScale = 0.5 : Flasher3lightsec.IntensityScale = 0.5 : Flasher4lightsec.IntensityScale = 0.5
		Flasher3light.state = LightStateOff : Flasher4light.state = LightStateOff : Flasher3lightsec.state = LightStateOff : Flasher4lightsec.state = LightStateOff
		'Flasher3.material = "flasherunlit" : Flasher3.image = "domelightoff" : Flasher4.material = "flasherunlit" : Flasher4.image = "domeshadowoff"
		Flasher3.material = "flasherunlit" : Flasher4.material = "flasherunlit" : FlashStepLeft = 0 : FlashTimerLeft.enabled = True
		If Rightflash Then
			sidewallFS.image = "sidewallsFSrightflash" : backwall.image = "backwallrightflash" : sidewallDT.image = "sidewallsFSrightflash"
			sidewallFS.material = "sidewallsrightflash" : sidewallDT.material = "sidewallsrightflash"
			ColorGradeFlash()
		Else
			sidewallFS.image = "sidewallsFS" : backwall.image = "backwallnoflash" : sidewallDT.image = "sidewallsFS"
			sidewallFS.material = "sidewallsnoflash" : sidewallDT.material = "sidewallsnoflash"
			ColorGradeFlash()
		End If
	End If
End sub

Dim FlashStepLeft

Sub FlashTimerLeft_timer()
	Flasher3.image = "flasherfade" & FlashStepLeft
	Flasher4.image = "flasherfade" & FlashStepLeft
	FlashStepLeft = FlashStepLeft + 1
	If FlashStepLeft = 6 Then
		Flasher3.image = "domelightoff" : Flasher4.image = "domeshadowoff"
		Flasher3.disablelighting = False : Flasher4.disablelighting = False
		FlashTimerLeft.enabled = False
	End If
End Sub

' *** flashes the whole screen by changing colorgrade LUT ***

Sub ColorGradeFlash()
	Dim lutlevel, ContrastLut
	Lutlevel = 0
	If Leftflash Then Lutlevel = lutlevel + 1 End If
	If Rightflash Then Lutlevel = lutlevel + 1 End If
	If ContrastSetting=0 or ContrastSetting=2 or ContrastSetting=4 or ContrastSetting=6 Then
		ContrastLut = ContrastSetting / 2
	Else
		ContrastLut = ContrastSetting / 2 - 0.5
	End If
	diner.ColorGradeImage = "LUT" & ContrastLut & "_" & lutlevel
End Sub

' *** keyboard handlers ***

Sub Diner_KeyDown(ByVal keycode)
	if keycode = LeftFlipperKey and FlippersEnabled Then B2SSTAT: SolLFlipper(True)
	if keycode = RightFlipperKey and FlippersEnabled Then B2SSTAT: SolRflipper(True)
	If keycode = LeftTiltKey Then Nudge 90, 3 : PlaySound SoundFX("fx_nudge",0)
	If keycode = RightTiltKey Then Nudge 270, 3 : PlaySound SoundFX("fx_nudge",0)
	If keycode = CenterTiltKey Then Nudge 0, 3 : PlaySound SoundFX("fx_nudge",0)
	If keycode = RightMagnaSave Then
		ContrastSetting = ContrastSetting + 1
		If ContrastSetting > 7 Then ContrastSetting = 7 End If
		If ContrastSetting=0 or ContrastSetting=2 or ContrastSetting=4 or ContrastSetting=6 Then ChangeGlow(True) Else ChangeGlow(False) End If
		ColorGradeFlash()
		OptionPrim.image = "contrastsetting" & ContrastSetting : OptionPrim.visible = 1 : OptionOpacity = 0 : OptionTimer.enabled = True
	End If
	If keycode = LeftMagnaSave Then
		ContrastSetting = ContrastSetting - 1
		If ContrastSetting < 0 Then ContrastSetting = 0 End If
		If ContrastSetting=0 or ContrastSetting=2 or ContrastSetting=4 or ContrastSetting=6 Then ChangeGlow(True) Else ChangeGlow(False) End If
		ColorGradeFlash()
		OptionPrim.image = "contrastsetting" & ContrastSetting : OptionPrim.visible = 1 : OptionOpacity = 0 : OptionTimer.enabled = True
	End If
	If keycode = PlungerKey Then PlaySoundAt "fx_PlungerPull", Plunger:Plunger.Pullback
	If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub Diner_KeyUp(ByVal keycode)
	if keycode = LeftFlipperKey and FlippersEnabled Then SolLFlipper(False)
	if keycode = RightFlipperKey and FlippersEnabled Then SolRflipper(False)
	If keycode = PlungerKey Then B2SSTAT: PlaySoundAt "fx_plunger", Plunger:Plunger.Fire
	If vpmKeyUp(keycode) Then Exit Sub
End Sub

Sub Diner_Paused: Controller.Pause = 1:End Sub
Sub Diner_UnPaused: Controller.Pause = 0:End Sub
Sub Diner_Exit(): Controller.Stop:End Sub



Sub OptionTimer_timer()
	If OptionOpacity = 8 Then
		OptionPrim.visible = 0
		OptionTimer.enabled = False
	Else
		OptionPrim.material = "PrimOption" & OptionOpacity
		OptionOpacity = OptionOpacity + 1
	End If
End Sub


'*** Flippers ***
' PlaySound "name",loopcount,volume,pan,randompitch,pitch,useexisting,restart
' pitch can be positive or negative and directly adds onto the standard sample frequency

Sub SolLFlipper(Enabled)
    If Enabled Then
		LeftFlipper.RotateToEnd : PlaySoundAt SoundFX("FlipperL",DOFFlippers), LeftFlipper
    Else
		LeftFlipper.RotateToStart : PlaySoundAt SoundFX("FlipperDown",DOFFlippers), LeftFlipper
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
		RightFlipper.RotateToEnd : PlaySoundAt SoundFX("FlipperR",DOFFlippers), RightFlipper
    Else
		RightFlipper.RotateToStart : PlaySoundAt SoundFX("FlipperDown",DOFFlippers), RightFlipper
    End If
End Sub

'******************************************************************************************
' Sling Shot Animations: Rstep and Lstep  are the variables that increment the animation
'******************************************************************************************

Dim RStep, Lstep

Sub RSling_Slingshot
	vpmTimer.PulseSwitch (swRSling), 0, "" : PlaySound SoundFx("SlingshotRight",DOFContactors), 0, 1, 0.05, 0.05
    RSling3.Visible = 0 : RSling1.Visible = 1 : sling1.TransZ = -20 : RStep = 0 : RSling.TimerEnabled = 1
End Sub

Sub RSling_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSling3.Visible = 1:sling1.TransZ = 0:RSling.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LSling_Slingshot
	vpmTimer.PulseSwitch (swLSling), 0, "" : PlaySound SoundFx("SlingshotLeft",DOFContactors), 0, 1, -0.05, 0.05
    LSling3.Visible = 0 : LSling1.Visible = 1 : sling2.TransZ = -20 : LStep = 0 : LSling.TimerEnabled = 1
End Sub

Sub LSling_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSling3.Visible = 1:sling2.TransZ = 0:LSling.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub SolBallRelease(enabled)
	if enabled then
		if bsTrough.Balls then vpmTimer.PulseSwitch(swTroughEject),0,""
		bsTrough.ExitSol_On
	End if
End Sub

' *** General Illumination ***

Sub SolGIRelay(enabled)
	Dim Prim, Light
	If enabled Then
		PlaySound "fx_relay_off",0,relaylevel : StopSound "fx_relay_on"
		For Each Light in LightsGI: Light.State = LightStateOff : Next
		For Each Prim in primitivesGI : Prim.DisableLighting = 0 : Next
		For Each Prim in primitivesGIpegs : Prim.image = "peg" : Prim.material = "peg3": Next
		If Diner.ShowDT then rampsDT.image = "rampsoffDT" Else rampsFS.image = "rampsoffFS"	End If
		cupprim.image = "cupoff"
	Else
		PlaySound "fx_relay_on",0,relaylevel : StopSound "fx_relay_off"
		For Each Light in LightsGI : Light.State = LightStateOn : Next
		For Each Prim in primitivesGI : Prim.DisableLighting = 1 : Next
		For Each Prim in primitivesGIpegs : Prim.image = "peglight" : Prim.material = "peg3light": Next
		If Diner.ShowDT then rampsDT.image = "rampsDT" Else	rampsFS.image = "rampsFS" End If
		cupprim.image = "cup"
	End If
End Sub

'*** Lighting system (mainly inserts) ***

Sub LampTimer_Timer()
    Dim chgLamp, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp) Then
        For ii = 0 To UBound(chgLamp)
			Select Case LightType(chgLamp(ii, 0))
				Case 0
					' *** LightType 0 = normal insert lighting with two lights
					If LightObjectFirst(chgLamp(ii, 0)).state <> chgLamp(ii, 1) Then
						LightObjectFirst(chgLamp(ii, 0)).state = chgLamp(ii, 1) : LightObjectSecond(chgLamp(ii, 0)).state = chgLamp(ii, 1)
					End If
				Case 1
					' *** LightType 1 = one object only for visible on/off
					If LightObjectFirst(chgLamp(ii, 0)).visible <> chgLamp(ii, 1) Then
						LightObjectFirst(chgLamp(ii, 0)).visible = chgLamp(ii, 1)
					End If
				Case 2
					' *** LightType 2 = EAT multiple Lights
					If LightObjectFirst(chgLamp(ii, 0)).state <> chgLamp(ii, 1) Then
						LightObjectFirst(chgLamp(ii, 0)).state = chgLamp(ii, 1) : LightObjectSecond(chgLamp(ii, 0)).state = chgLamp(ii, 1)
						LightObjectThird(chgLamp(ii, 0)).visible = chgLamp(ii, 1)
					End If
				Case 3
					' *** LightType 3 = cutout flashers (pepe, babs, boris, haji, buck)
					If LightObjectFirst(chgLamp(ii, 0)).state <> chgLamp(ii, 1) Then
						LightObjectFirst(chgLamp(ii, 0)).state = chgLamp(ii, 1) : LightObjectSecond(chgLamp(ii, 0)).state = chgLamp(ii, 1)
					End If
					LightState(chgLamp(ii, 0)) = chgLamp(ii, 1)
			End Select
        Next
    End If
End Sub

' *** Collection Hit Sounds ***

Sub Rubbers_Hit(idx):PlaySound "rubber11", 0, Vol(ActiveBall)*10*GlobalSoundLevel , pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub Posts_Hit(idx):PlaySound "post5", 0, Vol(ActiveBall)*10*GlobalSoundLevel , pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub Wood_Hit(idx) : PlaySound "WoodHit" : End Sub
Sub Metal_Hit(idx) : PlaySound "MetalHit2" : End Sub
Sub BalldropSound: PlaySound "fx_ball_drop" : End Sub
Sub Gates_hit(idx) : Playsound "gate4",0,3*GlobalSoundLevel  : End Sub
Sub OnBallBallCollision(ball1, ball2, velocity) : PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ActiveBall) : End Sub
Sub Trigger1_Hit() : Stopsound "plasticrolling" : PlaySound "metalrolling", -1,GlobalSoundLevel * 0.3, 0, 0, 1, 0, 0, AudioFade(ActiveBall) : End Sub
Sub LeftFlipper_Collide(parm) : PlaySound "flip_hit_1", 0, GlobalSoundLevel * parm / 50, -0.1, 0.25, AudioFade(ActiveBall) : End Sub
Sub RightFlipper_Collide(parm) : PlaySound "flip_hit_1", 0, GlobalSoundLevel * parm / 50, 0.1, 0.25, AudioFade(ActiveBall) : End Sub
Sub SolACSelect(enabled) : If Enabled Then PlaySound "fx_relay_on",0,GlobalSoundLevel  * relaylevel : StopSound "fx_relay_off" Else PlaySound "fx_relay_off",0, GlobalSoundLevel * relaylevel : StopSound "fx_relay_on": End If : End Sub
Sub RRentrysound_hit() : If Activeball.VelY < 0 Then Playsound "plasticrolling",-1, GlobalSoundLevel * 0.05 Else Stopsound "plasticrolling" : End If : End Sub
Sub RRentrysound_UnHit() : If Activeball.VelY > 0 Then Stopsound "plasticrolling" : End If : End Sub
Sub LRentrysound_hit() : If Activeball.VelY < 0 and Ramp1.collidable Then Playsound "plasticrolling",-1, GlobalSoundLevel * 0.05 Else	Stopsound "plasticrolling" : End If : End Sub
Sub LRentrysound_UnHit() : If Activeball.VelY > 0 Then Stopsound "plasticrolling" : End If : End Sub

' *** Ball Shadow code / Glow Ball code / Primitive Flipper Update ***

Dim BallShadowArray
BallShadowArray = Array (BallShadow1, BallShadow2, BallShadow3)
Const anglecompensate = 15

Sub GraphicsTimer_Timer()
	Dim BOT, b
    BOT = GetBalls

	' switch off glowlight for removed Balls
	IF GlowBall Then
		For b = UBound(BOT) + 1 to 2
			If GlowBall and Glowing(b).state = 1 Then Glowing(b).state = 0 End If
		Next
	End If

    For b = 0 to UBound(BOT)
		' *** move ball shadow for max 3 balls ***
		If BallShadow and b < 3 Then
			If BOT(b).X < Diner.Width/2 Then
				BallShadowArray(b).X = ((BOT(b).X) - (50/6) + ((BOT(b).X - (Diner.Width/2))/7)) + 10
			Else
				BallShadowArray(b).X = ((BOT(b).X) + (50/6) + ((BOT(b).X - (Diner.Width/2))/7)) - 10
			End If
			BallShadowArray(b).Y = BOT(b).Y + 20 : BallShadowArray(b).Z = 1
			If BOT(b).Z > 20 Then BallShadowArray(b).visible = 1 Else BallShadowArray(b).visible = 0 End If
		End If
		' *** move glowball light for max 3 balls ***
		If GlowBall and b < 3 Then
			If Glowing(b).state = 0 Then Glowing(b).state = 1 end if
			Glowing(b).BulbHaloHeight = BOT(b).z + 51
			Glowing(b).x = BOT(b).x : Glowing(b).y = BOT(b).y + anglecompensate
		End If
	Next
	If ChooseBats = 1 Then
		' *** move primitive bats ***
		batleft.objrotz = LeftFlipper.CurrentAngle + 1
		batleftshadow.objrotz = batleft.objrotz
		batright.objrotz = RightFlipper.CurrentAngle - 1
		batrightshadow.objrotz  = batright.objrotz
	Else
		If ChooseBats > 1 Then
			' *** move glowbats ***
			GlowBatLightLeft.y = 1720 - 121 + LeftFlipper.CurrentAngle
			glowbatleft.objrotz = LeftFlipper.CurrentAngle
			GlowBatLightRight.y =1720 - 121 - RightFlipper.CurrentAngle
			glowbatright.objrotz = RightFlipper.CurrentAngle
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

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "Diner" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / Diner.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "Diner" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / Diner.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "Diner" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / Diner.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioFade(ball) ' Can this be together with the above function ?
  Dim tmp
  tmp = ball.y * 2 / Diner.height-1
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

' *************************************************************
'      JP's VP10 Rolling Sounds
' *************************************************************

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
		If GlowBall and b <3 Then Glowing(b).state = 0 End If
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


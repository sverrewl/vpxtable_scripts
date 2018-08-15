'############################################################################################
'############################################################################################
'#######                                                                             ########
'#######          Junk Yard                                                          ########
'#######          (Williams 1996)                                                    ########
'#######                                                                             ########
'############################################################################################
'############################################################################################
' Version 1.3 FS mfuegemann 2016
'
' Thanks to:
' 3D primitives and images provided by Fuzzel (Toilet, Dog/House, Fridge, Bus, Flippers)
' 3D primitives and images provided by Dark (Toilet Ramp, Bus Ramp, Car Switch Assembly)
' Hauntfreaks for the lighting review and texture cleanups
'
' Version 1.1
' - corrected ramp end drop sound
' - corrected toilet ramp cover DepthBias setting to show ramp beneath
' - delete unused images
' - DOF sound review by Arngrim
' - Flipper and Wall friction settings reviewed by ClarkKent
' - removed Wire Triggers from collection to get correct Hit events
'
' Version 1.2
' - corrected backdrop flasher and lower left flasher RotX in DT mode
' - corrected SideWood material for correct DT backdrop display
' - adjustable sound level for ball rolling sound
' - option to increase BallSize (must be placed before LoadVPM)
' - fixed the plunger animation for mech plungers
'
' Version 1.3
' - corrected error on table start if B2S backglass was active

' Thalamus 2018-07-23
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' Thalamus 2018-08-15 : Improved directional sounds

option Explicit

' !! NOTE : Table not verified yet !!

Const cSingleLFlip = 0
Const cSingleRFlip = 0

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

'-----------------------------------
' Configuration
'-----------------------------------
cGameName = "jy_12"
Const DimFlashers = -0.4		'set between -1 and 0 to dim or brighten the Flashers (minus is darker)
Const OutLanePosts = 1			'1=Easy, 2=Medium, 3=Hard
Const ApronShopping = True		'Show WindowShopping ontop of the apron
const RollingSoundFactor = 23	'set sound level factor here for Ball Rolling Sound, 1=default level
Const BallSize = 51				'default 50, 52 plays well, above 55 the ball will get stuck

' Standard Sounds and Settings
Const SSolenoidOn="Solon",SSolenoidOff="Soloff",SFlipperOn="",SFlipperOff=""
Const UseSolenoids=2,UseLamps=True

InitWreckerBall						'must be called before LoadVPM becaus of B2s caused delay on trigger code

LoadVPM "01560000","WPC.VBS",3.2

'-----------------------------------
' Solenoids
'-----------------------------------

SolCallBack(1)	= "AutoPlunger"
SolCallBack(2)	= "bsFridgePopper.SolOut"
SolCallBack(3)	= "SolPowerCrane"
SolCallBack(5)	= "ScoopDown"
SolCallBack(6)	= "BusDiverter"
SolCallBack(7)	= "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallBack(9)	= "SolTrough"
SolCallBack(10)	= "vpmSolSound SoundFX(""Slingshot"",DOFContactors),"
SolCallBack(11)	= "vpmSolSound SoundFX(""Slingshot"",DOFContactors),"
SolCallBack(15)	= "SolHoldCrane"
SolCallBack(16)	= "SpikeBark"						     'Move Dog
SolCallBack(17)	= "SolFlash17"               'Dog Face Flasher
'SolCallBack(18)	= "SolWindowShopFlasher"
SolCallBack(19)	= "vpmFlasher Sol19,"		     'Autofire Flasher
SolCallBack(20)	= "SolFlash20"					     'Left red flasher
SolCallBack(21)	= "ScoopUp"
SolCallBack(22)	= "SolFlash22"						   'Under Crane Flasher
SolCallBack(23)	= "SolFlash23"						   'Back Left Flasher
SolCallBack(24)	= "SolFlash24"						   'Back Right Flasher
SolCallBack(25)	= "SolFlash25"						   'Shooter Flasher
SolCallBack(26)	= "SolFlash26"						   'Scoop Flasher
SolCallBack(27)	= "SolFlash27"						   'Dog House Flasher
SolCallBack(28)	= "SolFlash28"						   'Car Flashers (2)
SolCallback(sLLFlipper) = "SolLFlipper"
SolCallback(sLRFlipper) = "SolRFlipper"

Set GICallBack = GetRef("UpdateGI")

Sub Autoplunger(enabled)
	if enabled then
		if controller.switch(18) then
      PlaySoundAt SoundFX("plunger",DOFContactors), Plunger
			Auto_Plunger.Pullback
			Auto_Plunger.Fire
		end if
	Else
		Auto_Plunger.Pullback
	end If
End Sub

Sub SolTrough(Enabled)
	If Enabled then
		if not TestWall.isdropped Then
			TestWall.isdropped = True
		end If
		bsTrough.ExitSol_On
		vpmTimer.PulseSw 31
    End If
End Sub

Sub BusDiverter(Enabled)
	If Enabled then
    PlaySound SoundFX("Solon",DOFContactors)
		Sol6.IsDropped = False
    PlaySound SoundFX("Soloff",DOFContactors)
		Sol6.IsDropped = True
	End If
End Sub

Sub ScoopDown(Enabled)
	If Enabled Then
    PlaySoundAt SoundFX("Solon",DOFContactors), Primitive19
		Controller.Switch(72) = True
		P_Fork1.RotX = 17
		P_Fork2.RotX = 17
		P_Fork1.collidable = False
		P_Fork2.collidable = False
	End If
End Sub

Sub ScoopUp(Enabled)
	If Enabled Then
    PlaySoundAt SoundFX("Solon",DOFContactors), Primitive19
		Controller.Switch(72) = False
		P_Fork1.RotX = 0
		P_Fork2.RotX = 0
		P_Fork1.collidable = True
		P_Fork2.collidable = True
	End If
End Sub

dim spikeOut:spikeOut=1
Sub SpikeBark(Enabled)
	SpikeTimer.Enabled=Enabled
	if Enabled=0 then
		spike.transy=0
   	    spikeOut=1
	end if
End Sub

Sub SpikeTimer_Timer()
    PlaySound SoundFX("motor",DOFGear)
	if spikeOut=1 then
		spike.transy=spike.transy-8
		if spike.transy<-50 then
			spike.transy=-50
			spikeOut=0
		end if
	else
		spike.transy=spike.transy+8
		if spike.transy>0 then
			spike.transy=0
			spikeOut=1
		end if
	end if
end sub

Sub SolFlash17(Enabled)
	if enabled then
		setflash 4,1
	else
		setflash 4,0
	end if
End Sub

Sub SolFlash20(Enabled)
	if enabled then
		P_FridgeFlasher.image = "dome2_0_red"
		setflash 0,1
	else
		P_FridgeFlasher.image = "dome2_0_red_dark"
		setflash 0,0
	end if
End Sub

Sub SolFlash22(Enabled)
    if enabled then
		CraneFlasherLight.state = Lightstateon
		setflash 6,1
	else
		CraneFlasherLight.state = Lightstateoff
		setflash 6,0
	end if
End Sub

Sub SolFlash23(Enabled)
	if enabled then
		P_BackLeftFlasher.image = "dome2_0_clear"
		setflash 2,1
	else
		P_BackLeftFlasher.image = "dome2_0_clear_dark"
		setflash 2,0
	end if
End Sub

Sub SolFlash24(Enabled)
	if enabled then
		P_BackRightFlasher.image = "dome2_0_clear"
		setflash 3,1
	else
		P_BackRightFlasher.image = "dome2_0_clear_dark"
		setflash 3,0
	end if
End Sub

Sub SolFlash25(Enabled)
	if enabled then
		setflash 5,1
	else
		setflash 5,0
	end if
End Sub

Sub SolFlash26(Enabled)
	if enabled then
		P_DogScoopFlasher.image = "dome2_0_red"
		setflash 1,1
	else
		P_DogScoopFlasher.image = "dome2_0_red_dark"
		setflash 1,0
	end if
End Sub

Sub SolFlash27(Enabled)
	if enabled then
		setflash 7,1
	else
		setflash 7,0
	end if
End Sub

Sub SolFlash28(Enabled)
	if enabled then
		setflash 8,1
	else
		setflash 8,0
	end if
End Sub


'-----------------------------------
' Table Init
'-----------------------------------

Dim bsTrough,bsFridgePopper,cGameName,RightDrain,WindowUp

Sub JunkYard_Init()
	if JunkYard.ShowDT then
		Flasher1.Rotx = -30
		Flasher1.Roty = 50
		Flasher1.height = 240
		Flasher1a.Rotx = -40
		Flasher1a.height = 220
		Flasher1b.Rotx = -40
		Flasher1b.height = 225
		Flasher2.Rotx = -40
		Flasher2.height = 225
		Flasher2a.Rotx = -40
		Flasher2a.height = 220
		Flasher2b.Rotx = -40
		Flasher2b.height = 250
		Flasher3.Rotx = -20
		Flasher3a.Rotx = -20
		Flasher4.Rotx = -20
		Flasher4a.Rotx = -20
	Else
		Ramp9.widthbottom = 0
		Ramp9.widthtop = 0
		Ramp15.widthbottom = 0
		Ramp15.widthtop = 0
		Ramp13.widthbottom = 0
		Ramp13.widthtop = 0
		Ramp16.widthbottom = 0
		Ramp16.widthtop = 0
	End If

	if Apronshopping Then
		WindowUp = True
		P_Window.TransZ = 3
		P_Window1.TransZ = 0
	end If

	select case OutLanePosts
		case 2: 	'Medium Position
			P_RightRingHard.visible = False
			P_RightPostHard.visible = False
			RightPostHard.isdropped = True
			P_RightRingMedium.visible = True
			P_RightPostMedium.visible = True
			RightPostMedium.isdropped = False
			P_RightRingEasy.visible = False
			P_RightPostEasy.visible = False
			RightPostEasy.isdropped = True

			P_LeftRingHard.visible = False
			P_LeftPostHard.visible = False
			LeftPostHard.isdropped = True
			P_LeftRingMedium.visible = True
			P_LeftPostMedium.visible = True
			LeftPostMedium.isdropped = False
			P_LeftRingEasy.visible = False
			P_LeftPostEasy.visible = False
			LeftPostEasy.isdropped = True
		case 3: 	'Hard Position
			P_RightRingHard.visible = True
			P_RightPostHard.visible = True
			RightPostHard.isdropped = False
			P_RightRingMedium.visible = False
			P_RightPostMedium.visible = False
			RightPostMedium.isdropped = True
			P_RightRingEasy.visible = False
			P_RightPostEasy.visible = False
			RightPostEasy.isdropped = True

			P_LeftRingHard.visible = True
			P_LeftPostHard.visible = True
			LeftPostHard.isdropped = False
			P_LeftRingMedium.visible = False
			P_LeftPostMedium.visible = False
			LeftPostMedium.isdropped = True
			P_LeftRingEasy.visible = False
			P_LeftPostEasy.visible = False
			LeftPostEasy.isdropped = True
		case Else 	'Easy Position
			P_RightRingHard.visible = False
			P_RightPostHard.visible = False
			RightPostHard.isdropped = True
			P_RightRingMedium.visible = False
			P_RightPostMedium.visible = False
			RightPostMedium.isdropped = True
			P_RightRingEasy.visible = True
			P_RightPostEasy.visible = True
			RightPostEasy.isdropped = False

			P_LeftRingHard.visible = False
			P_LeftPostHard.visible = False
			LeftPostHard.isdropped = True
			P_LeftRingMedium.visible = False
			P_LeftPostMedium.visible = False
			LeftPostMedium.isdropped = True
			P_LeftRingEasy.visible = True
			P_LeftPostEasy.visible = True
			LeftPostEasy.isdropped = False
	end select

	vpminit me

	On Error Resume Next

    Controller.GameName=cGameName
    Controller.Games(cGameName).Settings.Value("samples")=0
	Controller.SplashInfoLine	= "Junk Yard, Williams 1996" & vbNewLine & "Created by mfuegemann"
	Controller.ShowDMDOnly		= True
    Controller.HandleKeyboard	= False
	Controller.ShowTitle		= False
    Controller.ShowFrame 		= False
	Controller.HandleMechanics 	= False

    ' DMD position for 3 Monitor Setup
    'Controller.Games(cGameName).Settings.Value("dmd_pos_x")=3850
    'Controller.Games(cGameName).Settings.Value("dmd_pos_y")=300
    'Controller.Games(cGameName).Settings.Value("dmd_width")=505
    'Controller.Games(cGameName).Settings.Value("dmd_height")=155
    'Controller.Games(cGameName).Settings.Value("rol")=0
	'Controller.Games(cGameName).Settings.Value("ddraw") = 1             'set to 0 if You have problems with DMD showing or table stutter

	Controller.Run
	If Err Then MsgBox Err.Description
	On Error Goto 0

	PinMAMETimer.Interval=PinMAMEInterval

	vpmNudge.TiltSwitch = 14
	vpmNudge.Sensitivity = 4

	vpmMapLights AllLights

	'Bus diverter starts down
	Sol6.IsDropped = True

	'Scoop Init (Up)
	controller.switch(72) = False
	P_Fork1.RotX = 0
	P_Fork2.RotX = 0
	P_Fork1.collidable = True
	P_Fork2.collidable = True

	'Crane Init (Down)
	controller.switch(28) = True

	'Fridge Popper
	Set bsFridgePopper = New cvpmBallStack
	bsFridgePopper.InitSw 0,37,36,43,0,0,0,0
	bsFridgePopper.InitKick Sol2, 180, 6
	bsFridgePopper.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solon",DOFContactors)
	bsFridgePopper.Balls = 1

	'Trough
	Set bsTrough = New cvpmBallStack
	bsTrough.InitSw 0,32,33,34,35,0,0,0
	bsTrough.InitKick BallRelease, 95, 5
	bsTrough.InitEntrySnd SoundFX("BallRelease",DOFContactors), "Solon"
	bsTrough.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solon",DOFContactors)
	bsTrough.Balls = 3

	Auto_Plunger.Pullback
End Sub

Sub JunkYard_Exit
	Controller.Stop
End Sub


'-----------------------------------
' Keyboard Handlers
'-----------------------------------

Sub JunkYard_KeyDown(ByVal keycode)
	If keycode = PlungerKey Then Plunger.Pullback
	If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub JunkYard_KeyUp(ByVal keycode)
	If keycode = PlungerKey Then Plunger.Fire
	If vpmKeyUp(keycode) Then Exit Sub
End Sub

'-----------------------------------
'UnderTrough Handling
'-----------------------------------

' Trough Handler
Sub Drain_Hit()
  PlaySoundAt "Drain5", Drain
	bsTrough.AddBall Me
End Sub

' Past Spinner
Sub LaunchHole_Hit
  PlaySoundAt "metalhit2", LaunchHole
  PlaySoundAt "Drain5", LaunchHole
  Me.DestroyBall
  LaunchEntry.createBall
  LaunchEntry.kick 220,5
End Sub

' Sewer
Sub Sewer_Hit
  PlaySoundAt "metalhit2", Sewer
  PlaySoundAt "Drain5", Sewer
  Me.DestroyBall
  SewerEntry.createBall
  SewerEntry.kick 260,5
End Sub

' Dog Entry
Sub DogHole_Hit
  PlaySoundAt "Drain5", DogHole
  Me.DestroyBall
  DogEntry.createBall
  DogEntry.kick 220,5
End Sub

' Crane Hole
Sub CraneHole_Hit
	playsoundAt "Drain5", CraneHole
	Me.DestroyBall
	CraneEntry.createBall
	CraneEntry.kick 180,5
End Sub

' Subway Handler
Sub FridgeEnter_Hit
	bsFridgePopper.Addball me
End Sub


'-----------------------------------
' Switch Handling
'-----------------------------------
Sub Switch12_Hit:vpmTimer.PulseSw 12:End Sub
Sub Switch12a_Hit:vpmTimer.PulseSw 12:End Sub
Sub Switch115_Spin:vpmTimer.PulseSw 115:PlaySoundAt "fx_spinner", Switch115:End Sub
Sub Switch11_Hit:vpmTimer.PulseSw 11:End Sub
Sub switch15_Hit:vpmTimer.PulseSw 15:End Sub
Sub switch16_Hit:Controller.Switch(16) = true:End Sub
Sub switch16_UnHit:Controller.Switch(16) = false:End Sub
Sub switch17_Hit:Controller.Switch(17) = true:End Sub
Sub switch17_UnHit:Controller.Switch(17) = false:End Sub
Sub switch18_Hit:Controller.Switch(18) = True:End Sub
Sub switch18_UnHit:Controller.Switch(18) = False:End Sub
Sub switch26_Hit:Controller.Switch(26) = true:End Sub
Sub switch26_UnHit:Controller.Switch(26) = false:End Sub
Sub switch27_Hit:Controller.Switch(27) = True:End Sub
Sub switch27_UnHit:Controller.Switch(27) = False:End Sub
Sub switch38_Hit:vpmTimer.PulseSw 38:End Sub
Sub switch41_Hit:vpmTimer.PulseSw 41:End Sub
Sub switch42_Hit:vpmTimer.PulseSw 42:End Sub
Sub switch44_Hit:vpmTimer.PulseSw 44:End Sub
Sub switch45_Hit:Controller.Switch(45) = True:Primitive_SwitchArm.objrotx = 10:End Sub
Sub switch45_UnHit:Controller.Switch(45) = False:Switch45.timerenabled = True:End Sub
Sub switch56_Hit:vpmTimer.PulseSw 56:End Sub
Sub switch57_Hit:vpmTimer.PulseSw 57:End Sub
Sub switch58_Hit:vpmTimer.PulseSw 58:End Sub
Sub switch61_Hit:vpmTimer.PulseSw 61:End Sub
Sub switch62_Hit:vpmTimer.PulseSw 62:End Sub
Sub switch63_Hit:vpmTimer.PulseSw 63:End Sub
Sub switch64_Hit:vpmTimer.PulseSw 64:End Sub
Sub switch65_Hit:vpmTimer.PulseSw 65:End Sub
Sub switch66_Hit:vpmTimer.PulseSw 66:End Sub
Sub switch67_Hit:Controller.Switch(67) = true : End Sub
Sub switch67_UnHit:Controller.Switch(67) = false : End Sub
Sub switch68_Hit:Controller.Switch(68) = True :End Sub
Sub switch68_UnHit:Controller.Switch(68) = False : End Sub
Sub switch71_Hit:Controller.Switch(71) = True : End Sub
Sub switch71_UnHit:Controller.Switch(71) = False : End Sub
Sub switch73_Hit:vpmTimer.PulseSw 73:End Sub
Sub switch74_Hit:Controller.Switch(74) = True:End Sub
Sub switch74_UnHit:Controller.Switch(74) = False:End Sub
Sub switch76_Hit:vpmTimer.PulseSw 76:End Sub
Sub switch77_Hit:vpmTimer.PulseSw 77:End Sub
Sub switch78_Hit:vpmTimer.PulseSw 78:End Sub

Sub Switch45_Timer
	Switch45.timerenabled = False
	Primitive_SwitchArm.objrotx = -20
End Sub

Sub ScoopMade_Hit:PlaySoundAt "WireRamp1", ScoopMade:End Sub
Sub EnterWireRamp_Hit:PlaySoundAt "WireRamp1", EnterWireRamp:End Sub
Sub ToiletBowlSwitch_Hit:ActiveBall.VelY = ActiveBall.VelY * 0.6:End Sub

Sub BusRampEnd_UnHit:PlaySoundAt "fx_collide", BusRampEnd:End Sub
Sub FridgeRampEnd_Hit
	FridgeRampEnd.timerenabled = True
End Sub
Sub FridgeRampEnd1_Hit
	if FridgeRampEnd.timerenabled Then
		FridgeRampEnd.timerenabled = False
		playsound "fx_collide",0,0.1,-0.15,0.25
	end If
End Sub

Sub RightRampEnd_UnHit
	RightRampEnd.timerenabled = True
End Sub
Sub RightRampEnd1_Hit
	if RightRampEnd.timerenabled Then
		RightRampEnd.timerenabled = False
    PlaySoundAt "fx_collide", RightRampEnd1
	end If
End Sub


'-----------------------------------
' Wrecking Ball Code
'-----------------------------------
Dim Wrecker,Wrecker2,bottomlimit,XBallX,YBallX,ScaleFactor,LowerBottomLimit,UpperBottomLimit,Wreckerballsize
Dim HoldCrane,WreckBallCenterX,WreckBallCenterY
LowerBottomLimit = 30
UpperBottomLimit = 120
bottomlimit = LowerBottomLimit
HoldCrane = False
WreckBallCenterX = WreckerCenterTrigger.X
WreckBallCenterY = WreckerCenterTrigger.Y
ScaleFactor = 0.65    'length of wrecker ball chain

Sub InitWreckerBall
	WreckerBallSize = Ballsize
	set Wrecker = WreckBallKicker1.createsizedball(Wreckerballsize / 2)
	WreckBallKicker1.kick 0,0
	set Wrecker2 = WreckBallKicker.createball
	Wrecker2.visible = False
	WreckBallKicker.kick 0,0
	WreckBallKicker.timerEnabled = true
End Sub

Dim ORotX,ORotY,TransZValue
const P1Z=62
const P2Z=72
const P3Z=82
const P4Z=92
const P5Z=102
const P6Z=112
const P7Z=122
const P8Z=132
const P9Z=142
const P10Z=152
const ChainOrigin=162
const CraneUp=100

Sub WreckBallKicker_timer()
	if Wrecker2.z > 350 Then
		Wrecker2.z = 350
	end If

	Wrecker2.Velx = Wrecker2.Velx + Wrecker.velx * 1.7
	Wrecker2.Vely = Wrecker2.Vely + Wrecker.vely * 1.7
	Wrecker2.Velz = Wrecker2.Velz + Wrecker.velz * 1.3

    Wrecker.velx = 0
    Wrecker.vely = 0
    Wrecker.velZ = 0

	XBallX = (Wrecker2.X - WreckerCenterTrigger1.X)
	YBallX = (Wrecker2.Y - WreckerCenterTrigger1.Y)

    Wrecker.X = WreckBallCenterX + XBallX
    Wrecker.Y = WreckBallCenterY + YBallX
    Wrecker.Z = Wrecker2.Z + BottomLimit - 30

	if abs(XballX) > 10 then
		if abs(XballX) > 30 then
			if XballX > 0 then
				PCraneArm.RotZ = 2.7
			end if
			if XballX < 0 then
				PCraneArm.RotZ = 3.3
			end if
		else
			if abs(XballX) > 20 then
				if XballX > 0 then
					PCraneArm.RotZ = 2.8
				end if
				if XballX < 0 then
					PCraneArm.RotZ = 3.2
				end if
			else
				if XballX > 0 then
					PCraneArm.RotZ = 2.9
				end if
				if XballX < 0 then
					PCraneArm.RotZ = 3.1
				end if
			end if
		end if
	else
		PCraneArm.RotZ = 3
	end if
	PWBallCylinder.X = Wrecker.X
	PWBallCylinder.Y = Wrecker.Y
	PWBallCylinder.Z = Wrecker.Z
	PWBallCylinder.TransZ = 25

	PChain1.X = WreckBallCenterX + (XBallX * ScaleFactor *.87)
	PChain1.Y = WreckBallCenterY + (YBallX * ScaleFactor *.87)
	PChain2.X = WreckBallCenterX + (XBallX * ScaleFactor *.79)
	PChain2.Y = WreckBallCenterY + (YBallX * ScaleFactor *.79)
	PChain3.X = WreckBallCenterX + (XBallX * ScaleFactor *.71)
	PChain3.Y = WreckBallCenterY + (YBallX * ScaleFactor *.71)
	PChain4.X = WreckBallCenterX + (XBallX * ScaleFactor *.63)
	PChain4.Y = WreckBallCenterY + (YBallX * ScaleFactor *.63)
	PChain5.X = WreckBallCenterX + (XBallX * ScaleFactor *.55)
	PChain5.Y = WreckBallCenterY + (YBallX * ScaleFactor *.55)
	PChain6.X = WreckBallCenterX + (XBallX * ScaleFactor *.47)
	PChain6.Y = WreckBallCenterY + (YBallX * ScaleFactor *.47)
	PChain7.X = WreckBallCenterX + (XBallX * ScaleFactor *.39)
	PChain7.Y = WreckBallCenterY + (YBallX * ScaleFactor *.39)
	PChain8.X = WreckBallCenterX + (XBallX * ScaleFactor *.31)
	PChain8.Y = WreckBallCenterY + (YBallX * ScaleFactor *.31)
	PChain9.X = WreckBallCenterX + (XBallX * ScaleFactor *.23)
	PChain9.Y = WreckBallCenterY + (YBallX * ScaleFactor *.23)
	PChain10.X = WreckBallCenterX + (XBallX * ScaleFactor *.15)
	PChain10.Y = WreckBallCenterY + (YBallX * ScaleFactor *.15)

	OrotX = YBallX * 0.85 					'reduce MaxRotation to 55 degrees
	OrotY = -XBallX * 0.85
	PWBallCylinder.RotX = ORotX * 0.8
	PWBallCylinder.RotY = ORotY * 0.8
	PChain1.RotX = ORotX * 1.3	 			'add some distortion to chain angle for each link
	PChain1.RotY = ORotY * 1.3
	PChain2.RotX = ORotX * 1.2
	PChain2.RotY = ORotY * 1.2
	PChain3.RotX = ORotX * 1.1
	PChain3.RotY = ORotY * 1.1
	PChain4.RotX = ORotX * 1.05
	PChain4.RotY = ORotY * 1.05
	PChain5.RotX = ORotX
	PChain5.RotY = ORotY
	PChain6.RotX = ORotX * 0.95
	PChain6.RotY = ORotY * 0.95
	PChain7.RotX = ORotX * 0.9
	PChain7.RotY = ORotY * 0.9
	PChain8.RotX = ORotX * 0.85
	PChain8.RotY = ORotY * 0.85
	PChain9.RotX = ORotX * 0.7
	PChain9.RotY = ORotY * 0.7
	PChain10.RotX = ORotX * 0.7
	PChain10.RotY = ORotY * 0.7

	TransZValue = (ChainOrigin + (bottomlimit-lowerbottomlimit) * Craneup / (upperbottomlimit-lowerbottomlimit) - Wrecker.Z)/10
	PChain1.Z = Wrecker.Z + 25 + 0.7*TransZValue
	PChain2.Z = Wrecker.Z + 25 + 1.8*TransZValue
	PChain3.Z = Wrecker.Z + 25 + 2.9*TransZValue
	PChain4.Z = Wrecker.Z + 25 + 4*TransZValue
	PChain5.Z = Wrecker.Z + 25 + 5*TransZValue
	PChain6.Z = Wrecker.Z + 25 + 6*TransZValue
	PChain7.Z = Wrecker.Z + 25 + 7*TransZValue
	PChain8.Z = Wrecker.Z + 25 + 8*TransZValue
	PChain9.Z = Wrecker.Z + 25 + 9*TransZValue
	PChain10.Z = Wrecker.Z + 25 + 10*TransZValue
End Sub

Dim DirX, DirY
const CraneUpAngle = -1
const CraneDownAngle = -11

Sub SolPowerCrane(enabled)
  PlaySoundAt SoundFX("Crane",DOFShaker), WreckBallKicker1
	If  enabled then
		CraneUpTimer.enabled = True
		CraneDownTimer.enabled = False
		DirX = Wrecker.velx
		DirY = Wrecker.vely
		if DirX <> 0 then
			DirX = DirX/ABS(DirX)
		else
			DirX = -1
		end if
		if DirY <> 0 then
			DirY = DirY/ABS(DirY)
		else
			DirY = 1
		end if

		Wrecker.velx = Wrecker.velx + (2.5 * DirX)		' If crane is pulled up, increase current movement
		Wrecker.vely = Wrecker.vely + (2.5 * DirY)

	Else
		If HoldCrane = False Then
			CraneUpTimer.enabled = False
			CraneDownTimer.enabled = True
		end if
	End If
End Sub

Sub SolHoldCrane(enabled)
	PlaySoundAt "Crane", PCraneArm
	If enabled then
		HoldCrane = True
		CraneUpTimer.enabled = True
		CraneDownTimer.enabled = False
		DirX = Wrecker.velx
		DirY = Wrecker.vely
		if DirX <> 0 then
			DirX = DirX/ABS(DirX)
		else
			DirX = 1
		end if
		if DirY <> 0 then
			DirY = DirY/ABS(DirY)
		else
			DirY = -1
		end if

		Wrecker.velx = Wrecker.velx + 2.5 * DirX		' If crane is pulled up, increase current movement
		Wrecker.vely = Wrecker.vely + 2.5 * DirY
	Else
		HoldCrane = False
		CraneUpTimer.enabled = False
		CraneDownTimer.enabled = True
	End If
End Sub

Sub CraneUpTimer_Timer
	PCraneArm.RotX = PCraneArm.RotX + 1
	PCraneArm.TransZ = PCraneArm.TransZ + 1
	BottomLimit = BottomLimit + 9
	if PCraneArm.RotX >= CraneUpAngle then
		CraneUpTimer.enabled = false
		PCraneArm.RotX = CraneUpAngle
		PCraneArm.TransZ = 10
		BottomLimit = UpperBottomLimit
		controller.switch(28) = False
	end if
End Sub

Sub CraneDownTimer_Timer
	PCraneArm.RotX = PCraneArm.RotX - 1
	PCraneArm.TransZ = PCraneArm.TransZ - 1
	BottomLimit = BottomLimit - 9
	if PCraneArm.RotX <= CraneDownAngle then
		CraneDownTimer.enabled = false
		PCraneArm.RotX = CraneDownAngle
		PCraneArm.TransZ = 0
		BottomLimit = LowerBottomLimit
		controller.switch(28) = true
	end if
End Sub

Sub SWCar1_Hit:vpmTimer.PulseSw 46:MoveCar1:End Sub
Sub SWCar2_Hit:vpmTimer.PulseSw 47:MoveCar2:End Sub
Sub SWCar3_Hit:vpmTimer.PulseSw 48:MoveCar3:End Sub
Sub SWCar4_Hit:vpmTimer.PulseSw 53:MoveCar4:End Sub
Sub SWCar5_Hit:vpmTimer.PulseSw 54:MoveCar5:End Sub

Sub MoveCar1
	P_Car1.TransZ = 5
	P_Car1.ObjRotX = -3
	P_Car1.ObjRotY = 6
	SWCar1.Timerenabled = False
	SWCar1.Timerenabled = True
End Sub
Sub SWCar1_Timer
	SWCar1.Timerenabled = False
	P_Car1.TransZ = 0
	P_Car1.ObjRotX = 0
	P_Car1.ObjRotY = 0
End Sub

Sub MoveCar2
	P_Car2.TransZ = 5
	P_Car2.ObjRotX = -6
	P_Car2.ObjRotY = 3
	SWCar2.Timerenabled = False
	SWCar2.Timerenabled = True
End Sub
Sub SWCar2_Timer
	SWCar2.Timerenabled = False
	P_Car2.TransZ = 0
	P_Car2.ObjRotX = 0
	P_Car2.ObjRotY = 0
End Sub

Sub MoveCar3
	P_Car3.TransZ = 5
	P_Car3.ObjRotX = -7
	SWCar3.Timerenabled = False
	SWCar3.Timerenabled = True
End Sub
Sub SWCar3_Timer
	SWCar3.Timerenabled = False
	P_Car3.TransZ = 0
	P_Car3.ObjRotX = 0
End Sub

Sub MoveCar4
	P_Car4.TransZ = 5
	P_Car4.ObjRotX = -6
	P_Car4.ObjRotY = -4
	SWCar4.Timerenabled = False
	SWCar4.Timerenabled = True
End Sub
Sub SWCar4_Timer
	SWCar4.Timerenabled = False
	P_Car4.TransZ = 0
	P_Car4.ObjRotX = 0
	P_Car4.ObjRotY = 0
End Sub

Sub MoveCar5
	P_Car5.TransZ = 5
	P_Car5.ObjRotX = -3
	P_Car5.ObjRotY = -6
	SWCar5.Timerenabled = False
	SWCar5.Timerenabled = True
End Sub
Sub SWCar5_Timer
	SWCar5.Timerenabled = False
	P_Car5.TransZ = 0
	P_Car5.ObjRotX = 0
	P_Car5.ObjRotY = 0
End Sub

'-----------------------------------
' Flipper Primitives
'-----------------------------------
sub FlipperTimer_Timer()
	pleftFlipper.rotz=leftFlipper.CurrentAngle
	prightFlipper.rotz=rightFlipper.CurrentAngle

	P_Spinner.rotx = -Switch115.currentangle
	P_Spinnerrod.rotx = -Switch115.currentangle
end sub

Sub SolLFlipper(Enabled)
    If Enabled Then
PlaySoundAt SoundFX("fx_flipperup",DOFContactors), LeftFlipper
		 LeftFlipper.RotateToEnd
    Else
PlaySoundAt SoundFX("fx_flipperdown",DOFContactors), LeftFlipper
		 LeftFlipper.RotateToStart
	End If
End Sub

Sub SolRFlipper(Enabled)
	If Enabled Then
PlaySoundAt SoundFX("fx_flipperup",DOFContactors), RightFlipper
		 RightFlipper.RotateToEnd
    Else
PlaySoundAt SoundFX("fx_flipperdown",DOFContactors), RightFlipper
		 RightFlipper.RotateToStart
    End If
End Sub

Sub LampTimer_Timer
	if Controller.Lamp(81) and WindowUp Then
		F_81.intensityscale = 1
	Else
		F_81.intensityscale = 0
	End If
	if Controller.Lamp(82) and WindowUp  Then
		F_82.intensityscale = 1
	Else
		F_82.intensityscale = 0
	End If
	if Controller.Lamp(83) and WindowUp  Then
		F_83.intensityscale = 1
	Else
		F_83.intensityscale = 0
	End If
	if Controller.Lamp(84) and WindowUp  Then
		F_84.intensityscale = 1
	Else
		F_84.intensityscale = 0
	End If
	if Controller.Lamp(85) and WindowUp  Then
		F_85.intensityscale = 1
	Else
		F_85.intensityscale = 0
	End If
End Sub

'-----------------------------------
' GI
'-----------------------------------
dim obj
Sub UpdateGI(GINo,Status)
	select case GINo
		case 0: if status then
					for each obj in GIString1
						obj.state = lightstateon
					next
				else
					for each obj in GIString1
						obj.state = lightstateoff
					next
				end if
		case 1: if status then
					for each obj in GIString2
						obj.state = lightstateon
					next
				else
					for each obj in GIString2
						obj.state = lightstateoff
					next
				end if

		case 2: if status then
					'BackGlass "Junk Yard" logo On
				else
					'BackGlass "Junk Yard" logo Off
				end if
	end select
End Sub



'########################################################################


'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
	vpmTimer.PulseSwitch 52, 0, 0
    PlaySoundAt SoundFX("left_slingshot",DOFContactors), sling1
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
	vpmTimer.PulseSwitch 51, 0, 0
    PlaySoundAt SoundFX("right_slingshot",DOFContactors), sling2
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
	PlaySound "metalhit_thin", 0, 0.2, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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
  PlaySoundAt "fx_spinner", Spinner
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


'---------------------------------------
'------  JP's Flasher Fading Sub  ------
'---------------------------------------

Dim Flashers
Flashers = Array(Flasher1,Flasher2,Flasher3,Flasher4,Flasher17,Flasher25,Flasher22,Flasher27,Flasher28,Flasher1a,Flasher2a,Flasher3a,Flasher4a,Flasher28a,Flasher1b,Flasher2b,Flasher17a)

Dim FlashMaxAlpha
Dim FlashState(200), FlashLevel(200)
Dim FlashSpeedUp, FlashSpeedDown

FlashInit()
FlasherTimer.Interval = 5
FlasherTimer.Enabled = 1

Sub FlashInit
    Dim i
    For i = 0 to 200
        FlashState(i) = 0
        FlashLevel(i) = 0
    Next

    FlashSpeedUp = 0.18    'fast speed when turning on the flasher
    FlashSpeedDown = 0.04  'slow speed when turning off the flasher, gives a smooth fading
    ' you could also change the default images for each flasher or leave it as in the editor
    ' for example
    ' Flasher1.Image = "fr"

	'added by mfuegemann to apply Dim settings
	FlashMaxAlpha = 1
	if DimFlashers < 0 then
		FlashMaxAlpha = FlashMaxAlpha + DimFlashers
		if FlashMaxAlpha < 0 then
			FlashMaxAlpha = 0
		end if
	end if

    AllFlashOff()
End Sub

Sub AllFlashOff
    Dim i
    For i = 0 to 200
        FlashState(i) = 0
    Next
End Sub

Sub FlasherTimer_Timer()
	flashm 0,Flasher1a
	flashm 0,Flasher1b
	flash 0, Flasher1
	flashm 1,Flasher2a
	flashm 1,Flasher2b
	flash 1, Flasher2
	flashm 2,Flasher3a
	flash 2, Flasher3
	flashm 3,Flasher4a
	flash 3, Flasher4
	Flashm 4, Flasher17a
	flash 4, Flasher17
	flash 5, Flasher25
	flash 6, Flasher22
	flash 7, Flasher27
	flashm 8,Flasher28a
	flash 8, Flasher28
End Sub

Sub SetFlash(nr, stat)
    FlashState(nr) = ABS(stat)
End Sub

Sub Flash(nr, object)
    Select Case FlashState(nr)
        Case 0 'off
            FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown
            If FlashLevel(nr) < 0 Then
                FlashLevel(nr) = 0
                FlashState(nr) = -1 'completely off
            End if
            Object.intensityscale = FlashLevel(nr)
        Case 1 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp
            If FlashLevel(nr) > FlashMaxAlpha Then				'1 original JP code
                FlashLevel(nr) = FlashMaxAlpha					'1 original JP code
                FlashState(nr) = -2 'completely on
            End if
            Object.intensityscale = FlashLevel(nr)
    End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it doesn't change the flashstate
	Object.intensityscale = FlashLevel(nr)
'    Select Case FlashState(nr)
'        Case 0         'off
'            Object.intensityscale = FlashLevel(nr)
'        Case 1         ' on
'            Object.intensityscale = FlashLevel(nr)
'    End Select
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

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "JunkYard" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / JunkYard.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "JunkYard" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / JunkYard.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "JunkYard" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / JunkYard.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioFade(ball) ' Can this be together with the above function ?
  Dim tmp
  tmp = ball.y * 2 / JunkYard.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
  Vol = Csng(BallVel(ball) ^2 / 2000)
  Vol = Vol * RollingSoundFactor
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

Const tnob = 6 ' total number of balls		(JY: 4 + Wrecker + Wrecker2)
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
  If JunkYard.VersionMinor > 3 OR JunkYard.VersionMajor > 10 Then
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
  Else
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0
  End if
End Sub


' Thalamus : Exit in a clean and proper way
Sub JunkYard_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


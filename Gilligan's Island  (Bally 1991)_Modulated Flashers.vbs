
'#########################################################################
Option Explicit
Randomize


' Thalamus 2018-07-23
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' No special SSF tweaks yet.

Const UseVPMModSol = 1

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName = "gi_l9"


Const GISetColor = False		'False = Normal GI, True = Color LEDs

for each obj in LightSetNormal
	obj.visible = not GISetColor
Next
for each obj in LightSetColor
	obj.visible = GISetColor
Next


Dim DesktopMode: DesktopMode = Gilligan.ShowDT

If DesktopMode = True Then 'Show Desktop components
Ramp16.visible=1
Ramp12.visible=1
SideWood.visible=1
SideWood1.visible=0
Else
Ramp16.visible=0
Ramp12.visible=0
SideWood.visible=0
SideWood1.visible=1
End if

Const UseSolenoids = 2
Const UseLamps = 0
Const UseSync = 1

' Standard Sounds
Const SSolenoidOn = ""
Const SSolenoidOff = ""
Const SFlipperOn = ""
Const SFlipperOff = ""
Const SCoin = "Coin"

LoadVPM "01560000", "WPC.VBS", 3.26


SolCallback(1) = "SolLeftAutoPlunger"			'Lagoon Kicker (Left Lock)
'SolCallback(2) = 'Island Lock
SolCallback(3) = "bsTrough.SolIn"				'Outhole
SolCallback(4) = "SolVUKPopper"					'VUK Popper
'5	RightSlingShot
'6	LeftSlingShot
SolCallback(7) = "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"	'Knocker
SolCallback(8) = "SolKickBack"					'sKickBack
'SolCallback(9) = IslandMotor handled by Mech
SolCallback(10) = "bsTrough.SolOut"				'BallRelease
SolCallback(11) = "SolHoldLock"					'Hold Lock
SolModCallback(12) = "Sol12"						'Island Light
'13	Bumper
'14 Bumper
'15 Bumper
SolCallback(16) = "SolTopKicker"				'Top Kicker
SolModCallback(17) = "Sol17"						'Head 1 Flasher
SolModCallback(18) = "SetModLamp 118,"				'Island Flasher
SolModCallback(19) = "SetModLamp 119,"				'Left Bank Flasher
SolModCallback(20) = "SetModLamp 120,"				'Left Lane Flasher
SolModCallback(21) = "SetModLamp 121,"				'Right Lane Flasher
SolModCallback(22) = "Sol22"						'Head 2 Flasher
SolCallback(23) = "SolRampUp"					'RampUp
SolCallback(24) = "SolRampDown"					'RampDown
SolModCallback(25) = "SetModLamp 125,"				'Right Bank Flasher
'SolCallback(26) = "vpmFlasher xxx,"			'Treasure Flasher
'SolCallback(27) = "vpmFlasher xxx,"			'Title Flasher - Backbox?
'SolCallback(28) = "vpmFlasher xxx,"			'Professor Flasher - Backbox?

'SolCallback(sLLFlipper)="vpmSolFlipper LeftFlipper,Nothing,"		'deactivated for better flipper response
'SolCallback(sLRFlipper)="vpmSolFlipper RightFlipper,Nothing,"
SolCallback(31) = "SolGameOn"					'GameOn

Set GICallBack = GetRef("UpdateGI")

Dim FlipperActive
FlipperActive = False
Sub SolGameOn(enabled)
	FlipperActive = enabled
	VpmNudge.SolGameOn(enabled)
	if not FlipperActive then
		RightFlipper.RotateToStart
		LeftFlipper.RotateToStart
	end if
End Sub

Sub Sol12(level)
	SetModLamp 112,level
	UpdateKona
End Sub

Sub Sol17(level)
	SetModLamp 117,level
	UpdateKona
End Sub

Sub Sol22(level)
	SetModLamp 122,level
	UpdateKona
End Sub

' GI
dim obj
Sub UpdateGI(GINo,Status)
	select case GINo
		'Left Inserts
		case 0: if status then
					if GISetColor Then
						for each obj in ColorGIString1
							obj.state = lightstateon
						next
					else
						for each obj in GIString1
							obj.state = lightstateon
						next
					end If
				else
					if GISetColor Then
						for each obj in ColorGIString1
							obj.state = lightstateoff
						next
					Else
						for each obj in GIString1
							obj.state = lightstateoff
						next
					End If
				end if
		'Bottom Playfield
		case 1: if status then
					if GISetColor Then
						for each obj in ColorGIString2
							obj.state = lightstateon
						next
					Else
						for each obj in GIString2
							obj.state = lightstateon
						next
					End If
				else
					if GISetColor Then
						for each obj in ColorGIString2
							obj.state = lightstateoff
						next
					Else
						for each obj in GIString2
							obj.state = lightstateoff
						next
					end If
				end if
		'Middle Playfield
		case 2: if status then
					if GISetColor Then
						for each obj in ColorGIString3
							obj.state = lightstateon
						next
					Else
						for each obj in GIString3
							obj.state = lightstateon
						next
					end If
				else
					if GISetColor Then
						for each obj in ColorGIString3
							obj.state = lightstateoff
						next
					Else
						for each obj in GIString3
							obj.state = lightstateoff
						next
					end If
				end if
		'Right Inserts
		case 3: if status then
					if GISetColor Then
						for each obj in ColorGIString4
							obj.state = lightstateon
						next
					Else
						for each obj in GIString4
							obj.state = lightstateon
						next
					end If
				else
					if GISetColor Then
						for each obj in ColorGIString4
							obj.state = lightstateoff
						next
					Else
						for each obj in GIString4
							obj.state = lightstateoff
						next
					end If
				end if
		'Top Playfield
		case 4: if status then
					if GISetColor Then
						for each obj in ColorGIString5
							obj.state = lightstateon
						next
					Else
						for each obj in GIString5
							obj.state = lightstateon
						next
					end If
					Primitive_PlasticRamp.image = "PlasticRampMapON"
					Primitive_RampHexPost1.image = "ParrotPost"
					Primitive_PlasticsCollection3.image = "PlasticsCollection3"
					Primitive_ProfessorPlastic.image = "ProfessorPlastic-Map"
					Primitive_bumpercap3.image = "BumperMap_ON"
					Primitive_bumpercap2.image = "BumperMap_ON"
					Primitive_bumpercap1.image = "BumperMap_ON"
				else
					if GISetColor Then
						for each obj in ColorGIString5
							obj.state = lightstateoff
						next
					Else
						for each obj in GIString5
							obj.state = lightstateoff
						next
					end If
					Primitive_PlasticRamp.image = "PlasticRampMap"
					Primitive_RampHexPost1.image = "ParrotPost_off"
					Primitive_PlasticsCollection3.image = "PlasticsCollection3_off"
					Primitive_ProfessorPlastic.image = "ProfessorPlasticOFF"
					Primitive_bumpercap3.image = "BumperMap_OFF"
					Primitive_bumpercap2.image = "BumperMap_OFF"
					Primitive_bumpercap1.image = "BumperMap_OFF"
				end if
	end select
End Sub

'Top Kicker
Sub SolTopKicker(enabled)
	bsTopEject.ExitSol_On
	Primitive_BallEjectArm.RotX = 50
	TopKicker.Timerenabled = True
End Sub

Sub TopKicker_Timer
	Primitive_BallEjectArm.RotX = Primitive_BallEjectArm.RotX - 5
	if Primitive_BallEjectArm.RotX <= 90 Then
		TopKicker.Timerenabled = False
		Primitive_BallEjectArm.RotX = 90
	End If
End Sub

'Kickback
Sub SolKickBack(enabled)
	if enabled Then
		LeftOutlane.timerenabled = False
		LeftOutlane.timerenabled = True
		KickerGate.collidable = True
	Else
		KickBack.Pullback
	end If
End Sub

Sub KickBackTrigger_Timer
	KickBackTrigger.timerenabled = False
	KickBack.Pullback
End Sub

Sub LeftOutLane_Timer
	LeftOutlane.timerenabled = False
	KickerGate.collidable = False
End Sub

Sub KickBackTrigger_Hit
	if LeftOutlane.timerenabled Then
		KickBack.Fire
		PlaySound SoundFX("popper_ball",DOFContactors)
		KickBackTrigger.timerenabled = True
	end If
End Sub

'Top Left Autoplunger
Sub SolLeftAutoPlunger(enabled)
	if enabled Then
		LeftAutoPlunger.Fire
		PlaySound SoundFX("popper_ball",DOFContactors),0,1,-0.18,0.25
	Else
		LeftAutoPlunger.Pullback
	End If
End Sub

'Ramp
Sub SolRampDown(Enabled)
	If Enabled  Then
		RampDir = 1
		RampBridgeTimer.enabled = True
		Controller.Switch(62)=1
		playsound SoundFX("Soloff",DOFContactors),0,1,-0.08,0.25
	End If
End Sub

Sub SolRampUp(Enabled)
	If Enabled then
		RampDir = -1
		RampBridgeTimer.enabled = True
		Controller.Switch(62)=0
		playsound SoundFX("Soloff",DOFContactors),0,1,-0.08,0.25
	End If
End Sub

' Vuk Popper code from JPSalas
Dim VUKBall, VUKBallZpos, AnimStarted
AnimStarted = 0

Sub SolVUKPopper(Enabled)
	If AnimStarted Then Exit Sub
	If Enabled Then
		If bsUpperEject.Balls Then
			AnimStarted = 1
			Set VUKBall = SeltzerHole.Createball
			VUKBallZpos = -50
			VUK.TimerInterval = 2
			VUK.TimerEnabled = 1
		End If
	End If
End Sub

Sub VUK_Timer
	VUKBall.Z = VUKBallZpos
	VUKBallZpos = VUKBallZpos + 5
	If VUKBallZpos> 155 Then
		VUK.TimerEnabled = 0
		SeltzerHole.DestroyBall
		bsUpperEject.ExitSol_On
		AnimStarted = 0
	End If
End Sub

Dim bsTrough,bsTopEject,bsLowKicker,bsUpperEject,mIsland
Sub Gilligan_Init
	'set initial textures
	Primitive_PlasticRamp.image = "PlasticRampMap"
	Primitive_RampHexPost1.image = "ParrotPost_off"
	Primitive_PlasticsCollection3.image = "PlasticsCollection3_off"
	Primitive_ProfessorPlastic.image = "ProfessorPlasticOFF"
	Primitive_bumpercap3.image = "BumperMap_OFF"
	Primitive_bumpercap2.image = "BumperMap_OFF"
	Primitive_bumpercap1.image = "BumperMap_OFF"
	Primitive_Kona.image = "KonaMap"
	P_BridgeRamp.image = "MechRamp_Map2"
	Primitive_PalmPostP.image = "PalmTreePlastic(sm)Map"
	Primitive_KonaFlasher.image = "dome4_red"

	vpmInit Me
    With Controller
		.GameName = cGameName
        'If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Gilligan's Island"
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = 0
 		.Dip(0) = &H00
        On Error Resume Next
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
        On Error Goto 0
        .Switch(22) = 1 'close coin door
        .Switch(24) = 1 'and keep it close
	End With

	'Nudging
    vpmNudge.TiltSwitch = 14
    vpmNudge.Sensitivity = 1
	vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingshot)

    'Trough
    Set bsTrough=new cvpmBallStack
    bsTrough.InitSw 18,17,16,0,0,0,0,0
	bsTrough.InitKick BallRelease,60,8
	bsTrough.InitEntrySnd SoundFX("BallRelease",DOFContactors),SoundFX("Solon",DOFContactors)
	bsTrough.InitExitSnd SoundFX("BallRelease",DOFContactors), SoundFX("Solon",DOFContactors)
	bsTrough.Balls=2

 	'Top kicker
    set bsTopEject = new cvpmSaucer
	bsTopEject.InitKicker TopKicker,67, 173, 8, 0
	bsTopEject.InitExitVariance 5, 2
	bsTopEject.InitSounds SoundFX("kicker_enter_center",DOFContactors),SoundFX("solon",DOFContactors),SoundFX("popper_ball",DOFContactors)

   'seltzerhole
    Set bsUpperEject = new cvpmBallStack
    bsUpperEject.InitSw 0,66,0,0,0,0,0,0
    bsUpperEject.InitKick VUK,130,6
    bsUpperEject.InitEntrySnd SoundFX("kicker_enter_center",DOFContactors), SoundFX("kicker_enter_center",DOFContactors)
    bsUpperEject.InitExitSnd SoundFX("popper_ball",DOFContactors), SoundFX("popper_ball",DOFContactors)
    bsUpperEject.Balls=0
    bsUpperEject.KickBalls=2

	Set mIsland=New cvpmmech
	With mIsland
		.mtype=vpmMechOneSol+vpmMechCircle+vpmMechLinear'onedirsol
		.sol1=9
		'Sw77 is the opto - It's mostly off
		'There is one gap for each stop - and a long gap for the home position
		'The switch would be 'on' during the gaps and off otherwise
		'I think each position is just after a pulse of the opto - since the home is a long pulse
		'so it must stop when the switch turns back off

		.length=120 '2 sec to find the new Position
		.steps=34
		.AddSw 77,1,6 ' First on gap is 1 position before key 1'
		.AddSw 77,8,13 ' Second on gap is 1 position before key 2'
		.AddSw 77,15,20 ' Third on gap is 1 position before key 3'
		.AddSw 77,22,27 ' Fourth on gap is 1 position before key 4'
		.AddSw 77,29,33 ' This is the 'long' gap on for several positions before key 0'
		.callback=getref("UpdateIsland")
		.start
    End With

    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1
	LeftAutoPlunger.Pullback
	KickBack.Pullback
end sub

Sub Gilligan_Paused:Controller.Pause = 1:End Sub
Sub Gilligan_unPaused:Controller.Pause = 0:End Sub
Sub Gilligan_Exit:Controller.Stop:End Sub

Sub drain_Hit()
    playsound "Drain"
	bsTrough.AddBall Me
End Sub

Sub Gilligan_KeyDown(ByVal keycode)
	If keycode = PlungerKey Then
		Plunger.PullBack
		PlaySound "plungerpull",0,1,0.25,0.25
	End If

	If keycode = LeftFlipperKey Then
		if FlipperActive then
			LeftFlipper.RotateToEnd
			PlaySound SoundFX("fx_flipperup",DOFFlippers), 0, .67, -0.05, 0.05
		end if
	End If

	If keycode = RightFlipperKey Then
		if FlipperActive then
			RightFlipper.RotateToEnd
			PlaySound SoundFX("fx_flipperup",DOFFlippers), 0, .67, 0.05, 0.05
		end if
	End If

'	If keycode = LeftTiltKey Then
'		Nudge 90, 2
'	End If
'
'	If keycode = RightTiltKey Then
'		Nudge 270, 2
'	End If
'
'	If keycode = CenterTiltKey Then
'		Nudge 0, 2
'	End If

	If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub Gilligan_KeyUp(ByVal keycode)

	If keycode = PlungerKey Then
		Plunger.Fire
		PlaySound "plunger",0,1,0.25,0.25
	End If

	If keycode = LeftFlipperKey Then
		LeftFlipper.RotateToStart
		if FlipperActive then
			PlaySound SoundFX("fx_flipperdown",DOFFlippers), 0, 1, -0.05, 0.05
		end if
	End If

	If keycode = RightFlipperKey Then
		RightFlipper.RotateToStart
		if FlipperActive then
			PlaySound SoundFX("fx_flipperdown",DOFFlippers), 0, 1, 0.05, 0.05
		end if
	End If

	If vpmKeyUp(keycode) Then Exit Sub
End Sub

'#######################################################################
'Switches

Sub EnterKonaHole_Hit
	EnterKonaHole.DestroyBall
	KonaKicker.CreateBall
	KonaKicker.kick 140,1
End Sub


Sub TopKicker_Hit:bsTopEject.Addball me:ActivateJungle:End Sub
Sub KonaHole_Hit:bsUpperEject.AddBall Me:End Sub
Sub SeltzerHole_Hit:bsUpperEject.AddBall Me:ActivateJungle:End Sub

Sub LeftOutlane_Hit:Controller.Switch(25) = True:End Sub
Sub LeftOutlane_Unhit:Controller.Switch(25) = False:End Sub
Sub LeftInlane_Hit:Controller.Switch(26) = True:End Sub
Sub LeftInlane_Unhit:Controller.Switch(26) = False:End Sub
Sub RightInlane_Hit:Controller.Switch(27) = True:End Sub
Sub RightInlane_Unhit:Controller.Switch(27) = False:End Sub
Sub RightOutlane_Hit:Controller.Switch(28) = True:End Sub
Sub RightOutlane_Unhit:Controller.Switch(28) = False:End Sub

Sub SW31_Hit:Controller.Switch(31) = True:End Sub
Sub SW31_Unhit:Controller.Switch(31) = False:End Sub
Sub SW33_Hit:Controller.Switch(33) = True:End Sub
Sub SW33_Unhit:Controller.Switch(33) = False:End Sub

Sub SW61a_Hit:Controller.Switch(61) = True:End Sub
Sub SW61a_Unhit:Controller.Switch(61) = False:End Sub
Sub SW61b_Hit:Controller.Switch(61) = True:End Sub
Sub SW61b_Unhit:Controller.Switch(61) = False:End Sub
Sub SW63_Hit:Controller.Switch(63) = True:End Sub
Sub SW63_Unhit:Controller.Switch(63) = False:End Sub
Sub SW64_Hit:Controller.Switch(64) = True:End Sub
Sub SW64_Unhit:Controller.Switch(64) = False:End Sub
Sub SW65_Hit:Controller.Switch(65) = True:End Sub
Sub SW65_Unhit:Controller.Switch(65) = False:End Sub
Sub SW68_Hit:Controller.Switch(68) = True:End Sub
Sub SW68_Unhit:Controller.Switch(68) = False:End Sub

Sub SW78_Hit:Controller.Switch(78) = True:End Sub
Sub SW78_Unhit:Controller.Switch(78) = False:End Sub
Sub SW75_Hit:Controller.Switch(75) = True:End Sub
Sub SW75_Unhit:Controller.Switch(75) = False:End Sub

Sub SW83_Hit:Controller.Switch(83) = True:End Sub
Sub SW83_Unhit:Controller.Switch(83) = False:End Sub
Sub SW84_Hit:Controller.Switch(84) = True:End Sub
Sub SW84_Unhit:Controller.Switch(84) = False:End Sub

Sub SW71_Hit:vpmTimer.PulseSw 71:End Sub
Sub SW72_Hit:vpmTimer.PulseSw 72:End Sub
Sub SW73_Hit:Controller.Switch(73) = True:Primitive_Switcharm73.rotx = 27:End Sub
Sub SW73_Unhit:Controller.Switch(73) = False:SW73.Timerenabled = True:End Sub
Sub SW74_Hit:Controller.Switch(74) = True:Primitive_Switcharm74.rotx = 27:End Sub
Sub SW74_Unhit:Controller.Switch(74) = False:SW74.Timerenabled = True:End Sub

Sub SW73_Timer
	Primitive_Switcharm73.rotx = Primitive_Switcharm73.rotx - 1
	if Primitive_Switcharm73.rotx <= 0 Then
		SW73.Timerenabled = False
		Primitive_Switcharm73.rotx = 0
	End If
End Sub

Sub SW74_Timer
	Primitive_Switcharm74.rotx = Primitive_Switcharm74.rotx - 1
	if Primitive_Switcharm74.rotx <= 0 Then
		SW74.Timerenabled = False
		Primitive_Switcharm74.rotx = 0
	End If
End Sub

'Rubbers
Sub SW32_Hit:vpmTimer.PulseSw 32:End Sub
Sub SW58_Hit:vpmTimer.PulseSw 58:End Sub



Sub Target35_Hit:vpmTimer.PulseSw 35:MoveTarget35:End Sub
Sub Target36_Hit:vpmTimer.PulseSw 36:MoveTarget36:End Sub
Sub Target37_Hit:vpmTimer.PulseSw 37:MoveTarget37:End Sub
Sub Target38_Hit:vpmTimer.PulseSw 38:MoveTarget38:End Sub

Sub Target46_Hit:vpmTimer.PulseSw 46:MoveTarget46:End Sub
Sub Target47_Hit:vpmTimer.PulseSw 47:MoveTarget47:End Sub
Sub Target48_Hit:vpmTimer.PulseSw 48:MoveTarget48:End Sub

Sub Target51_Hit:vpmTimer.PulseSw 51:MoveTarget51:End Sub
Sub Target52_Hit:vpmTimer.PulseSw 52:MoveTarget52:End Sub
Sub Target53_Hit:vpmTimer.PulseSw 53:MoveTarget53:End Sub
Sub Target54_Hit:vpmTimer.PulseSw 54:MoveTarget54:End Sub
Sub Target55_Hit:vpmTimer.PulseSw 55:MoveTarget55:End Sub
Sub Target56_Hit:vpmTimer.PulseSw 56:MoveTarget56:End Sub
Sub Target57_Hit:vpmTimer.PulseSw 57:MoveTarget57:End Sub

'#######################################################################
' Jungle Ramp

Sub ToggleRamp
	if not RampBridgeTimer.enabled Then
		RampDir = -RampDir
		RampBridgeTimer.enabled = True
	end If
end Sub

Dim RampDir
RampDir = -1

Const RampSpeed= 0.5
JungleRampDown.collidable = False

Sub RampBridgeTimer_Timer
	Primitive_MechArm.ObjRotX = Primitive_MechArm.ObjRotX + (RampDir * RampSpeed)
	Primitive_MechArm.ObjRotY = abs(Primitive_MechArm.ObjRotX)/10

	if Primitive_MechArm.ObjRotX >= -30 then
		JungleRampDown.collidable = True
		JungleRampUp.collidable = False
	Else
		JungleRampUp.collidable = True
		JungleRampDown.collidable = False
	End If

	if Primitive_MechArm.ObjRotX >= 4 then
		RampBridgeTimer.enabled = False
		Primitive_MechArm.ObjRotX = 4
	end if
	if Primitive_MechArm.ObjRotX <= -45 then
		RampBridgeTimer.enabled = False
		Primitive_MechArm.ObjRotX = -45
	end if
	P_BridgeRamp.Rotx = RampRot(Primitive_MechArm.ObjRotX)
End Sub

Dim ArmAngle
Const Pi=3.141592654
'P_BridgeRamp.Rotx = -RampRot(Primitive_MechArm.ObjRotX)

Function RampRot(ArmRotPar)
	ArmAngle = 122 + ArmRotPar
	RampRot = (1-Sin(ArmAngle/180*Pi))*100
	if RampRot > 14.9 Then
		RampRot = 14.9
	end If
End Function

'#############################################################################
'Jungle Turntable

'** Update Island
Dim islpos,IslandRotation

Sub UpdateIsland(anewpos,aspeed,alastpos)
	if JungleActive Then
		islpos=anewpos
		Select Case islpos
			Case 33,34,0,1:   IslandRotation = 357 	'Ramp_1
			Case 6,7,8:  IslandRotation = 283		'Ramp_2
			Case 13,14,15: IslandRotation = 211		'Ramp_3
			Case 20,21,22: IslandRotation = 140		'Ramp_4
			Case 27,28,29: IslandRotation = 67		'Ramp_5
		End Select
		TurntableGate.collidable = True
		TurnTableTimer.enabled = True
		TurnTableSoundTimer.enabled = True
	end If
End Sub

'Avoid turning wheel at start of table
Dim JungleActive
JungleActive = False

Sub ActivateJungle
	JungleActive = True
End Sub

'Init Paths
ToggleTurnTablePaths

Sub ToggleTurnTablePaths
	JungleRamp_1.collidable = (Primitive_TurnTable.ObjRotZ > 346) Or (Primitive_TurnTable.ObjRotZ < 8) 		'357
	JungleRamp_2.collidable = (Primitive_TurnTable.ObjRotZ > 272) and (Primitive_TurnTable.ObjRotZ < 294)	'283
	JungleRamp_3.collidable = (Primitive_TurnTable.ObjRotZ > 200) and (Primitive_TurnTable.ObjRotZ < 222)	'211
	JungleRamp_4.collidable = (Primitive_TurnTable.ObjRotZ > 129) and (Primitive_TurnTable.ObjRotZ < 151)	'140
	JungleRamp_5.collidable = (Primitive_TurnTable.ObjRotZ > 56) and (Primitive_TurnTable.ObjRotZ < 78)		' 67
End Sub

Sub TurnTableTimer_Timer
	Primitive_TurnTable.objRotZ = Primitive_TurnTable.ObjRotZ - 1
	if Primitive_TurnTable.objRotZ <= 0 Then
		Primitive_TurnTable.objRotZ = 360
	End If
	Primitive_PalmTreePlastics.objRotZ = Primitive_TurnTable.ObjRotZ
	Primitive_TurntableScrews.objRotZ = Primitive_TurnTable.ObjRotZ
    Primitive_TurnTableCover.ObjRotZ = Primitive_TurnTable.ObjRotZ
	Primitive_TopDecals.objRotZ = Primitive_TurnTable.ObjRotZ
	if Primitive_TurnTable.objRotZ = IslandRotation Then
		TurnTableTimer.enabled = False
		TurnTableSoundTimer.enabled = False
		TurntableGate.collidable = False
	end If
	ToggleTurnTablePaths
End Sub

Sub TurnTableSoundTimer_Timer
	playsound SoundFX("motor1",DOFGear),0,1,-0.05,0.1
End Sub

Sub SolHoldLock (enabled)
	if enabled then
		Controller.switch(76)=0
	else
		Controller.switch(76)=1
	end if
end sub

'#########################################################################
' Helper Functions

'Jungle Rescue - reactivate a ball that may become trapped under the turntable
Sub RescueKicker1_Hit
	RescueKicker1.destroyball
	RescueKicker2.createball
	RescueKicker2.kick 180,1
End Sub

Sub RampHelper_Hit
	ActiveBall.VelZ = 0
End Sub

'Gates
Sub GateTimer_Timer
	Primitive_VUKGateWire.RotX = -VUKGate.Currentangle
	Primitive_GateWire.RotX = -Gate4.Currentangle
End Sub

Sub MoveTarget35
	Primitive_TargetParts35.TransZ = 5
	Primitive_Target35.TransZ = 5
	Target35.Timerenabled = False
	Target35.Timerenabled = True
End Sub
Sub	Target35_Timer
	Target35.Timerenabled = False
	Primitive_TargetParts35.TransZ = 0
	Primitive_Target35.TransZ = 0
End Sub

Sub MoveTarget36
	Primitive_TargetParts36.TransZ = 5
	Primitive_Target36.TransZ = 5
	Target36.Timerenabled = False
	Target36.Timerenabled = True
End Sub
Sub	Target36_Timer
	Target36.Timerenabled = False
	Primitive_TargetParts36.TransZ = 0
	Primitive_Target36.TransZ = 0
End Sub

Sub MoveTarget37
	Primitive_TargetParts37.TransZ = 5
	Primitive_Target37.TransZ = 5
	Target37.Timerenabled = False
	Target37.Timerenabled = True
End Sub
Sub	Target37_Timer
	Target37.Timerenabled = False
	Primitive_TargetParts37.TransZ = 0
	Primitive_Target37.TransZ = 0
End Sub

Sub MoveTarget38
	Primitive_TargetParts38.TransZ = 5
	Primitive_Target38.TransZ = 5
	Target38.Timerenabled = False
	Target38.Timerenabled = True
End Sub
Sub	Target38_Timer
	Target38.Timerenabled = False
	Primitive_TargetParts38.TransZ = 0
	Primitive_Target38.TransZ = 0
End Sub

Sub MoveTarget46
	Primitive_Target46.TransZ = 5
	Target46.Timerenabled = False
	Target46.Timerenabled = True
End Sub
Sub	Target46_Timer
	Target46.Timerenabled = False
	Primitive_Target46.TransZ = 0
End Sub

Sub MoveTarget47
	Primitive_Target47.TransZ = 5
	Target47.Timerenabled = False
	Target47.Timerenabled = True
End Sub
Sub	Target47_Timer
	Target47.Timerenabled = False
	Primitive_Target47.TransZ = 0
End Sub

Sub MoveTarget48
	Primitive_Target48.TransZ = 5
	Target48.Timerenabled = False
	Target48.Timerenabled = True
End Sub
Sub	Target48_Timer
	Target48.Timerenabled = False
	Primitive_Target48.TransZ = 0
End Sub

Sub MoveTarget51
	Primitive_TargetParts51.TransZ = 5
	Primitive_Target51.TransZ = 5
	Target51.Timerenabled = False
	Target51.Timerenabled = True
End Sub
Sub	Target51_Timer
	Target51.Timerenabled = False
	Primitive_TargetParts51.TransZ = 0
	Primitive_Target51.TransZ = 0
End Sub

Sub MoveTarget52
	Primitive_TargetParts52.TransZ = 5
	Primitive_Target52.TransZ = 5
	Target52.Timerenabled = False
	Target52.Timerenabled = True
End Sub
Sub	Target52_Timer
	Target52.Timerenabled = False
	Primitive_TargetParts52.TransZ = 0
	Primitive_Target52.TransZ = 0
End Sub

Sub MoveTarget53
	Primitive_TargetParts53.TransZ = 5
	Primitive_Target53.TransZ = 5
	Target53.Timerenabled = False
	Target53.Timerenabled = True
End Sub
Sub	Target53_Timer
	Target53.Timerenabled = False
	Primitive_TargetParts53.TransZ = 0
	Primitive_Target53.TransZ = 0
End Sub

Sub MoveTarget54
	Primitive_TargetParts54.TransZ = 5
	Primitive_Target54.TransZ = 5
	Target54.Timerenabled = False
	Target54.Timerenabled = True
End Sub
Sub	Target54_Timer
	Target54.Timerenabled = False
	Primitive_TargetParts54.TransZ = 0
	Primitive_Target54.TransZ = 0
End Sub

Sub MoveTarget55
	Primitive_TargetParts55.TransZ = 5
	Primitive_Target55.TransZ = 5
	Target55.Timerenabled = False
	Target55.Timerenabled = True
End Sub
Sub	Target55_Timer
	Target55.Timerenabled = False
	Primitive_TargetParts55.TransZ = 0
	Primitive_Target55.TransZ = 0
End Sub

Sub MoveTarget56
	Primitive_TargetParts56.TransZ = 5
	Primitive_Target56.TransZ = 5
	Target56.Timerenabled = False
	Target56.Timerenabled = True
End Sub
Sub	Target56_Timer
	Target56.Timerenabled = False
	Primitive_TargetParts56.TransZ = 0
	Primitive_Target56.TransZ = 0
End Sub

Sub MoveTarget57
	Primitive_Target57.TransX = 5
	Target57.Timerenabled = False
	Target57.Timerenabled = True
End Sub
Sub	Target57_Timer
	Target57.Timerenabled = False
	Primitive_Target57.TransX = 0
End Sub

'Bumper animation
Dim BumperDir1,BumperDir2,BumperDir3,BumperSpeed
BumperSpeed = 1.75 * Bumper1.Ringspeed
Const BumperLowerLimit = -30

Sub Bumper1_Hit
	vpmTimer.PulseSw 41
	Playsound SoundFX("fx_bumper3",DOFContactors),0,1,0.08,0.1
	MoveBumperCap1
End Sub
Sub Bumper2_Hit
	vpmTimer.PulseSw 42
	Playsound SoundFX("fx_bumper2",DOFContactors),0,1,0.15,0.1
	MoveBumperCap2
End Sub
Sub Bumper3_Hit
	vpmTimer.PulseSw 43
	Playsound SoundFX("fx_bumper4",DOFContactors),0,1,0.15,0.1
	MoveBumperCap3
End Sub

Sub MoveBumperCap1
	BumperDir1 = -BumperSpeed
	Bumper1.Timerenabled = True
End Sub
Sub Bumper1_Timer
	Primitive_bumpercap1.TransY = Primitive_bumpercap1.TransY + BumperDir1
	Primitive_bumpercapSCREWS1.TransY = Primitive_bumpercapSCREWS1.TransY + BumperDir1
	if Primitive_bumpercap1.TransY >= 0 Then
		Bumper1.Timerenabled = False
		Primitive_bumpercap1.TransY = 0
		Primitive_bumpercapSCREWS1.TransY = 0
	end if
	if Primitive_bumpercap1.TransY <= BumperLowerLimit Then
		BumperDir1 = BumperSpeed
	end if
End Sub

Sub MoveBumperCap2
	BumperDir2 = -BumperSpeed
	Bumper2.Timerenabled = True
End Sub
Sub Bumper2_Timer
	Primitive_bumpercap2.TransY = Primitive_bumpercap2.TransY + BumperDir2
	Primitive_bumpercapSCREWS2.TransY = Primitive_bumpercapSCREWS2.TransY + BumperDir2
	if Primitive_bumpercap2.TransY >= 0 Then
		Bumper2.Timerenabled = False
		Primitive_bumpercap2.TransY = 0
		Primitive_bumpercapSCREWS2.TransY = 0
	end if
	if Primitive_bumpercap2.TransY <= BumperLowerLimit Then
		BumperDir2 = BumperSpeed
	end if
End Sub

Sub MoveBumperCap3
	BumperDir3 = -BumperSpeed
	Bumper3.Timerenabled = True
End Sub
Sub Bumper3_Timer
	Primitive_bumpercap3.TransY = Primitive_bumpercap3.TransY + BumperDir3
	Primitive_bumpercapSCREWS3.TransY = Primitive_bumpercapSCREWS3.TransY + BumperDir3
	if Primitive_bumpercap3.TransY >= 0 Then
		Bumper3.Timerenabled = False
		Primitive_bumpercap3.TransY = 0
		Primitive_bumpercapSCREWS3.TransY = 0
	end if
	if Primitive_bumpercap3.TransY <= BumperLowerLimit Then
		BumperDir3 = BumperSpeed
	end if
End Sub


'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim LStep, RStep

Sub LeftSlingShot_Slingshot
PlaySound SoundFX("left_slingshot",DOFContactors), 0, 0.3, -0.1, 0.25
vpmTimer.PulseSw 44
	LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
    LeftSlingShot.TimerInterval = 10
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
PlaySound SoundFX("right_slingshot",DOFContactors), 0, 0.3, 0.1, 0.25
vpmTimer.PulseSw 45
	RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
    RightSlingShot.TimerInterval = 10
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

'Ramp Helpers
Sub LHelp_Hit():Playsound "fx_ball_drop", 0, 1, -0.05, 0.05:end Sub
Sub RHelp_Hit():Playsound "fx_ball_drop", 0, 1, 0.05, 0.05:end Sub

'Wire ramp sounds
Sub WireRampSound_Hit():Playsound "WireRampR", 0, 1, 0, 0.35:end Sub
Sub WireRampSound1_Hit():Playsound "WireRampR", 0, 1, 0, 0.35:end Sub
Sub WireRampSound2_Hit():Playsound "WireRampL", 0, 1, 0, 0.35:end Sub
Sub WireRampSound3_Hit():Playsound "WireRampR", 0, 1, 0, 0.35:end Sub

'Plastic ramp sounds
Sub RRHit0_Hit:PlaySound "fx_rr1", 0, 1, pan(ActiveBall):End Sub
Sub RRHit1_Hit:PlaySound "fx_rr2", 0, 1, pan(ActiveBall):End Sub
Sub RRHit2_Hit:PlaySound "fx_rr3", 0, 1, pan(ActiveBall):End Sub
Sub RRHit3_Hit:PlaySound "fx_rr4", 0, 1, pan(ActiveBall):End Sub
Sub RRHit4_Hit:PlaySound "fx_rr5", 0, 1, pan(ActiveBall):End Sub
Sub RRHit5_Hit:PlaySound "fx_rr6", 0, 1, pan(ActiveBall):End Sub
Sub RRHit6_Hit:PlaySound "fx_rr2", 0, 1, pan(ActiveBall):End Sub
Sub RRHit7_Hit:PlaySound "fx_lr1", 0, 1, pan(ActiveBall):End Sub



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
	PlaySound SoundFX("target",DOFTargets), 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
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

'Kona Texture Swap - called only if necessary
Dim KonaState
Sub UpdateKona
	KonaState = 0
	if Controller.solenoid(12) then Konastate = Konastate + 1	'Spot
	if Controller.solenoid(17) then Konastate = Konastate + 10	'Red Dome
	if Controller.solenoid(22) then Konastate = Konastate + 100	'Head2

	Select Case Konastate
		case 0:	  	Primitive_Kona.image = "KonaMap"			'all OFF
		case 1:	  	Primitive_Kona.image = "KonaMap_SPOTL_ON1"	'Spot only
		case 10:	Primitive_Kona.image = "KonaMap_FLSH_ON1"	'Red Dome only
		case 11:	Primitive_Kona.image = "KonaMap_FLSH_ON3"	'Spot+Red Dome
		case 100:	Primitive_Kona.image = "KonaMap_FLSH2_ON1"	'Head2 only
		case 101:	Primitive_Kona.image = "KonaMap_FLSH2_ON3"	'Spot+Head2
		case 110:	Primitive_Kona.image = "KonaMap_FLSHx2_ON2"	'Red Dome+Head2
		case 111:	Primitive_Kona.image = "KonaMap_FLSHx2_ON3"	'all ON
	End Select
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
Dim ccount

InitLamps()             ' turn off the lights and flashers and reset them to the default parameters
LampTimer.Interval = 10 'lamp fading speed
LampTimer.Enabled = 1

Dim chgLamp, num, chg, ii, JungleRampChanged
Sub LampTimer_Timer()
	JungleRampChanged = False
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp) Then
        For ii = 0 To UBound(chgLamp)
            LampState(chgLamp(ii, 0) ) = chgLamp(ii, 1)       'keep the real state in an array
			FadingLevel(chgLamp(ii, 0) ) = chgLamp(ii, 1) + 4 'actual fading step

			'change Jungle Ramp only if necessary
			if chgLamp(ii, CHGNO) = 67 or chgLamp(ii, CHGNO) = 68 Then
				JungleRampChanged = True
			End If
		Next
    End If
    UpdateLamps

	'Gilligan texture swap light handler
	if JungleRampChanged Then
		if Controller.Lamp(67) or Controller.Lamp(68) Then
			P_BridgeRamp.image = "mechramp_map-on"
			Primitive_PalmPostP.image = "PalmTreePlastic(sm)MapON"
		Else
			P_BridgeRamp.image = "MechRamp_Map2"
			Primitive_PalmPostP.image = "PalmTreePlastic(sm)Map"
		End If
	End If
End Sub

Sub InitLamps()
    Dim x
    For x = 0 to 200
        LampState(x) = 0         ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4       ' used to track the fading state
        FlashSpeedUp(x) = 0.3    ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.15 ' slower speed when turning off the flasher
        FlashMax(x) = 1          ' the maximum value when on, usually 1
        FlashMin(x) = 0          ' the minimum value when off, usually 0
        FlashLevel(x) = 0        ' the intensity of the flashers, usually from 0 to 1
    Next
End Sub

Sub UpdateLamps
	'Inserts
	  Flash 11, Lamp11a
	NFadeLm 11, Lamp11b
	NFadeLm 11,  Lamp11
	NFadeLm 11, Lamp11c
	NFadeLm 11, Lamp11d
	NFadeLm 11, Lamp11e
	  Flash 12, Lamp12a
	NFadeLm 12, Lamp12b
	NFadeLm 12,  Lamp12
	NFadeLm 12, Lamp12c
	NFadeLm 12, Lamp12d
	NFadeLm 12, Lamp12e
	  Flash 13, Lamp13a
	NFadeLm 13, Lamp13b
	NFadeLm 13,  Lamp13
	NFadeLm 13, Lamp13c
	NFadeLm 13, Lamp13d
	  Flash 14, Lamp14a
	NFadeLm 14, Lamp14b
	NFadeLm 14,  Lamp14
	NFadeLm 14, Lamp14c
	NFadeLm 14, Lamp14d
	  Flash 15, Lamp15a
	NFadeLm 15, Lamp15b
	NFadeLm 15,  Lamp15
	  Flash 16, Lamp16a
	NFadeLm 16,  Lamp16
	NFadeLm 16, Lamp16b
	NFadeLm 16, Lamp16c
	NFadeLm 16, Lamp16d
	NFadeLm 16, Lamp16e
	NFadeLm 16, Lamp16f
	NFadeLm 16, Lamp16g
	  Flash 17, Lamp17g
	NFadeLm 17, Lamp17a
	NFadeLm 17, Lamp17b
	NFadeLm 17, Lamp17c
	NFadeLm 17, Lamp17d
	NFadeLm 17, Lamp17e
	NFadeLm 17, Lamp17f
	NFadeLm 17,  Lamp17
	  Flash 18, Lamp18a
	NFadeLm 18, Lamp18b
	NFadeLm 18, Lamp18c
	NFadeLm 18,  Lamp18

	  Flash 21, Lamp21a
	NFadeLm 21,  Lamp21
	  Flash 22, Lamp22a
	NFadeLm 22,  Lamp22
	  Flash 23, Lamp23a
	NFadeLm 23,  Lamp23
      Flash 24, Lamp24a
	NFadeLm 24,  Lamp24
      Flash 25, Lamp25a
	NFadeLm 25,  Lamp25
      Flash 26, Lamp26a
	NFadeLm 26,  Lamp26
	NFadeLm 26, Lamp26b
	NFadeLm 26, Lamp26c
      Flash 27, Lamp27a
	NFadeLm 27, Lamp27b
	NFadeLm 27,  Lamp27
      Flash 28, Lamp28a
	NFadeLm 28, Lamp28b
	NFadeLm 28,  Lamp28

      Flash 31, Lamp31a
	NFadeLm 31,  Lamp31
      Flash 32, Lamp32a
	NFadeLm 32,  Lamp32
      Flash 33, Lamp33a
	NFadeLm 33,  Lamp33
      Flash 34, Lamp34a
	NFadeLm 34,  Lamp34
      Flash 35, Lamp35a
	NFadeLm 35,  Lamp35
      Flash 36, Lamp36a
	NFadeLm 36,  Lamp36
	  Flash 37, Lamp37a
	NFadeLm 37,  Lamp37
	NFadeLm 37, Lamp37b
	  Flash 38, Lamp38a
	NFadeLm 38,  Lamp38

      Flash 41, Lamp41a
	NFadeLm 41,  Lamp41
      Flash 42, Lamp42a
	NFadeLm 42,  Lamp42
      Flash 43, Lamp43a
	NFadeLm 43,  Lamp43
      Flash 44, Lamp44a
	NFadeLm 44,  Lamp44
      Flash 45, Lamp45a
	NFadeLm 45,  Lamp45
      Flash 46, Lamp46a
	NFadeLm 46,  Lamp46
      Flash 47, Lamp47a
	NFadeLm 47,  Lamp47
     Flash 48, Lamp48a
	NFadeLm 48,  Lamp48

      Flash 51, Lamp51a
	NFadeLm 51,  Lamp51
      Flash 52, Lamp52a
	NFadeLm 52,  Lamp52
      Flash 53, Lamp53a
	NFadeLm 53,  Lamp53
      Flash 54, Lamp54a
	NFadeLm 54,  Lamp54
      Flash 55, Lamp55a
	NFadeLm 55,  Lamp55
      Flash 56, Lamp56a
	NFadeLm 56,  Lamp56
      Flash 57, Lamp57a
	NFadeLm 57, Lamp57b
	NFadeLm 57,  Lamp57
      Flash 58, Lamp58a
	NFadeLm 58, Lamp58b
	NFadeLm 58,  Lamp58

      Flash 61, Lamp61a
	NFadeLm 61, Lamp61b
	NFadeLm 61,  Lamp61
	NFadeLm 61, Lamp61c
	NFadeLm 61, Lamp61d
      Flash 62, Lamp62a
	NFadeLm 62, Lamp62b
	NFadeLm 62,  Lamp62
	NFadeLm 62, Lamp62c
	NFadeLm 62, Lamp62d
      Flash 63, Lamp63a
	NFadeLm 63, Lamp63b
	NFadeLm 63,  Lamp63
	NFadeLm 63, Lamp63c
      Flash 64, Lamp64a
	NFadeLm 64,  Lamp64
      Flash 65, Lamp65a
	NFadeLm 65,  Lamp65
      Flash 66, Lamp66a
	NFadeLm 66,  Lamp66
      Flash 67, Lamp67a
	NFadeLm 67,  Lamp67
      Flash 68, Lamp68a
	NFadeLm 68,  Lamp68
	  Flash 71, Lamp71a
	NFadeLm 71, Lamp71b
	NFadeLm 71,  Lamp71
      Flash 72, Lamp72a
	NFadeLm 72, Lamp72b
	NFadeLm 72,  Lamp72
      Flash 73, Lamp73a
	NFadeLm 73, Lamp73b
	NFadeLm 73,  Lamp73
      Flash 74, Lamp74a
	NFadeLm 74, Lamp74b
	NFadeLm 74,  Lamp74
      Flash 76, Lamp76a
	NFadeLm 76,  Lamp76
	NFadeLm 76, Lamp76b
      Flash 77, Lamp77a
	NFadeLm 77,  Lamp77
	NFadeLm 77, Lamp77b

'	'Flashers

	NFadeLMod 112, F112
	NFadeObjModm 117, Primitive_KonaFlasher, "dome4_red_lit", "dome4_red"	'call the Object update before the light update because of the fast status change
	NFadeLMod 117, F117
	NFadeLMod 117, F117a
	NFadeLMod 117, F117b
	NFadeLMod 118, F118
	NFadeLMod 119, F119
	NFadeLMod 120, F120
	NFadeLMod 121, F121
	NFadeLMod 122, F122
	NfadeLMod 122, F122b
	NFadeLMod 122, F122c
	NFadeLMod 122, F122d
	NFadeLMod 122, F122e
	NFadeLMod 122, F122f
	NfadeLMod 125, F125
 End Sub

Sub SetModLamp(nr, value)
	Debug.Print "SetModLamp " & nr & ": " & value
	FadingLevel(nr) = value
End Sub

Sub NFadeLMod(nr, object)
	if FadingLevel(nr) > 0 Then
		object.IntensityScale = FadingLevel(nr) / 255
		object.state = 1
	Else
		object.state = 0
	end if
End Sub

Sub NFadeObjModm(nr, object, a, b)
    if  FadingLevel(nr) = 0 then
        object.image = b
	Else
        object.image = a
    End If
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

Sub NFadeLmb(nr, object) ' used for multiple lights with blinking
    Select Case FadingLevel(nr)
        Case 4:object.state = 0
        Case 5:object.state = 2
    End Select
End Sub

'Lights, Ramps & Primitives used as 4 step fading lights
'a,b,c,d are the images used from on to off

Sub FadeObj(nr, object, a, b, c, d)
	Select Case FadingLevel(nr)
		Case 2:object.image = d:FadingLevel(nr) = 0 'Off
		Case 3:object.image = c:FadingLevel(nr) = 2 'fading...
		Case 4:object.image = b:FadingLevel(nr) = 3 'fading...
		Case 5:object.image = a:FadingLevel(nr) = 1 'ON
 End Select
End Sub

Sub FadeObjm(nr, object, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:object.image = d
        Case 3:object.image = c
        Case 4:object.image = b
        Case 5:object.image = d
    End Select
End Sub

Sub NFadeObj(nr, object, a, b)
    Select Case FadingLevel(nr)
        Case 4:object.image = b:FadingLevel(nr) = 0 'off
        Case 5:object.image = a:FadingLevel(nr) = 1 'on
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


Sub Lamp44a_Init()

End Sub

' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function Vol2(ball1, ball2) ' Calculates the Volume of the sound based on the speed of two balls
    Vol2 = (Vol(ball1) + Vol(ball2) ) / 2
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "Gilligan" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / Gilligan.width-1
    If tmp> 0 Then
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

'*****************************************
'    JP's VP10 Collision & Rolling Sounds
'*****************************************

Const tnob = 8 ' total number of balls
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

    ' rolling

	For B = UBound(BOT) +1 to tnob
        rolling(b) = False
        'StopSound("fx_ballrolling" & b)
		StopSound("fx_ballrolling" & b & "_loud")
	Next

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


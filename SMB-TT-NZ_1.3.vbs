'*************************************
 'Super Mario Brothers - IPDB No. 2579
 '© Gottlieb 1992
 'VPX by ninuzzu/Tom Tower
 'Script by edizzle/ninuzzu
' Thanks to zany for the domes and flippers models
' Thanks to VPDev Team for the freaking amazing VPX
 '************************************
	Option Explicit
	Randomize

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' No special SSF tweaks yet.
' , AudioFade(ActiveBall)

Dim Ballsize,BallMass
BallSize = 51
BallMass = (Ballsize^3)/125000

Dim DesktopMode:DesktopMode = Table1.ShowDT
Dim UseVPMDMD:UseVPMDMD = DesktopMode

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01560000", "gts3.VBS", 3.26

'********************
'Standard definitions
'********************

 Const UseSolenoids = 2
 Const UseLamps = 0
 Const UseSync = 0

'Standard Sounds
 Const SSolenoidOn = "Solenoid"
 Const SSolenoidOff = ""
 Const SCoin = "fx_coin"

'******************************************************
' 					TABLE INIT
'******************************************************

Const cGameName= "smb3"
Dim dtDrop

Sub Table1_Init
    With Controller
        .GameName = cGameName
        .SplashInfoLine = "Super Mario Bros. (Gottlieb 1994)"
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = DesktopMode
        On Error Resume Next
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
        On Error Goto 0
    End With

    '************  Main Timer init  ********************

	PinMAMETimer.Interval = PinMAMEInterval
	PinMAMETimer.Enabled = 1

    '************  Nudging   **************************

	vpmNudge.TiltSwitch = 151
	vpmNudge.Sensitivity = 3
	vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingShot, RightSlingShot)

    '************  Lower Droptarget
	Set dtDrop = new cvpmDropTarget
	With dtDrop
		.InitDrop Array(sw30), Array(30)
		.Initsnd  SoundFX("droptarget",DOFContactors),  SoundFX("resetdrop",DOFContactors)
	End With

    '************  Trough	**************************
	Slot1.CreateSizedballWithMass Ballsize/2,Ballmass
	Slot2.CreateSizedballWithMass Ballsize/2,Ballmass
	Slot3.CreateSizedballWithMass Ballsize/2,Ballmass
	Controller.Switch(34) = 1

    '************  Misc Stuff  ******************
	sw30.IsDropped=1
 	Lkick.collidable=0
	Rkick.collidable=0
	LeftRail.visible = DesktopMode
	RightRail.visible = DesktopMode

Select Case DesktopMode
	Case 0
		Apron.material = "Apron"
	Case 1
		Apron.material = "ApronDT"
End Select
End Sub

'******************************************************
' 						KEYS
'******************************************************

Sub Table1_KeyDown(ByVal keycode)
	If keycode = StartGameKey Then Controller.Switch(4)= 1
'		If BIP = 0 then Controller.Switch(4)= 1
'		If BIP = 1 AND Controller.Switch(32) Then Controller.Switch(4) = 1
'	End If
	If keycode = LeftFlipperKey Then Controller.Switch(45) = 1
	If keycode = RightFlipperKey Then Controller.Switch(47) = 1
	If keycode = LeftTiltKey Then Nudge 90, 3:PlaySound SoundFX("fx_nudge", 0)
    If keycode = RightTiltKey Then Nudge 270, 3:PlaySound SoundFX("fx_nudge", 0)
    If keycode = CenterTiltKey Then Nudge 0, 3:PlaySound SoundFX("fx_nudge", 0)
	If keycode = PlungerKey Then Controller.Switch(5)=1
	If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal keycode)
	If keycode = StartGameKey Then Controller.Switch(4) = 0
	If keycode = LeftFlipperKey Then Controller.Switch(45) = 0
	If keycode = RightFlipperKey Then Controller.Switch(47) = 0
	If keycode = PlungerKey Then Controller.Switch(5)=0
		If vpmKeyUp(keycode) Then Exit Sub
End Sub

Sub Table1_Paused: Controller.Pause = 1:End Sub
Sub Table1_UnPaused: Controller.Pause = 0:End Sub
Sub Table1_Exit(): Controller.Stop:End Sub

'******************************************************
'					SOLENOIDS
'******************************************************

'	SolCallback(1) = ""                           		'1 - Left Pop Bumper
'	SolCallback(2) = ""                           		'2 - Top Pop Bumper
'	SolCallback(3) = ""                           		'3 - Bottom Pop Bumper
'	SolCallback(4) = ""									'4 - Left Slingshot
'	SolCallback(5) = ""									'5 - Right Slingshot
	SolCallback(6) = "dtDrop.SolDropUp"					'6 - Shield Reset
	SolCallback(7) = "dtDrop.SolDropDown"				'7 - Shield Trip
	SolCallback(8) = "Ballsave"							'8 - Kicksave
	SolCallback(9) = "Shooter"							'9 - Shooter
   	SolCallback(10) = "LeftKicker"						'10 - Left upkicker
	SolCallback(11) = "RightKicker"						'11 - Right Upkicker
	SolCallback(12) = "UpperKicker"						'12	- Upper upkicker
	SolCallback(13) = "SetLamp 153,"					'13 - Mario speaking Flasher
	SolCallback(14) = "SetLamp 154,"					'14 - Mario speaking Flasher
	SolCallback(15) = "SetLamp 155,"					'15 - castle/princess ??????
	SolCallback(16) = "SetLamp 156,"					'16 - luigi and yoshi????????
	SolCallback(17) = "Setlamp 157,"					'17 - Left Ramp Flash
	SolCallback(18) = "SetLamp 158,"					'18 - Right Ramp Flash
	SolCallback(19) = "SetLamp 159,"					'19 - Mario Cave Flash (x2)
	SolCallback(20) = "SetLamp 160,"					'20 - Castle Flash (x2)
	SolCallback(21) = "SetLamp 161,"					'21 - Bottom Left Flash
   	SolCallback(22) = "SetLamp 162,"					'22 - Upper Left Flash
	SolCallback(23) = "SetLamp 163,"					'23 - Upper Center Flash (x2)
	SolCallback(24) = "SetLamp 164,"					'24 - Upper Right Flash (x2)
	SolCallback(25) = "CastleRotate"					'25 - Motor Relay
	SolCallback(26) = "GIRelay"							'26 - LightBox Relay
	SolCallback(28) = "ReleaseBall"						'28 - Ball Release
	SolCallback(29) = "SolOuthole"						'29 - Outhole
	SolCallback(30) = "SolKnocker"						'30 - Knocker
	SolCallback(31) = "TiltRelay"						'31 - Tilt
	SolCallback(32) = "vpmNudge.SolGameOn"				'32 - Game Over Relay

   SolCallback(sLRFlipper) = "SolRFlipper"
   SolCallback(sLLFlipper) = "SolLFlipper"

'******************************************************
'			TROUGH BASED ON FOZZY'S
'******************************************************

Sub Slot3_Hit():Controller.Switch(34) = 1:UpdateTrough:End Sub
Sub Slot3_UnHit():Controller.Switch(34) = 0:UpdateTrough:End Sub
Sub Slot2_UnHit():UpdateTrough:End Sub
Sub Slot1_UnHit():UpdateTrough:End Sub

Sub UpdateTrough()
	UpdateTroughTimer.Interval = 300
	UpdateTroughTimer.Enabled = 1
End Sub

Sub UpdateTroughTimer_Timer()
	If Slot1.BallCntOver = 0 Then Slot2.kick 60, 9
	If Slot2.BallCntOver = 0 Then Slot3.kick 60, 9
	Me.Enabled = 0
End Sub

'******************************************************
'				DRAIN & RELEASE
'******************************************************

Dim BIP : BIP = 0

Sub Drain_Hit()
	PlaySound "fx_drain"
	UpdateTrough
	BIP = BIP - 1
	Controller.Switch(33) = 1
End Sub

Sub Drain_UnHit()
	Controller.Switch(33) = 0
End Sub

Sub SolOuthole(enabled)
	If enabled Then
		Drain.kick 60,20
		PlaySound SoundFX(SSolenoidOn,DOFContactors)
	End If
End Sub

Sub ReleaseBall(enabled)
	If enabled Then
		PlaySound SoundFX("fx_ballrel",DOFContactors)
		Slot1.kick 60, 9
		UpdateTrough
		BIP = BIP + 1
	End If
End Sub

'******************************************************
'					LEFT KICKER
'******************************************************

Sub LeftKicker(enabled)
If enabled Then
sw16a.kick 0,35,1.56
Controller.Switch(16) = 0
vpmTimer.AddTimer 500, "sw16.enabled=1'"
vpmTimer.AddTimer 600, "Lkick.collidable=0'"
Playsound SoundFX("scoopexit",DOFContactors)
End If
End Sub

Sub sw16a_hit()
PlaySound "fx_kicker_catch"
Controller.Switch(16) = 1
sw16.enabled=0
Lkick.collidable=1
End Sub

Sub sw16_hit: PlaySound "kicker_enter": End Sub

'******************************************************
'					RIGHT KICKER
'******************************************************

Sub RightKicker(enabled)
If enabled Then
sw26a.kick 0,52,1.56
Controller.Switch(26) = 0
vpmTimer.AddTimer 500, "sw26.enabled=1'"
vpmTimer.AddTimer 600, "Rkick.collidable=0'"
Playsound SoundFX("scoopexit",DOFContactors)
End If
End Sub

sub sw26a_hit
StopSound "fx_subway"
PlaySound "fx_kicker_catch"
Controller.Switch(26) = 1
sw26.enabled= 0
Rkick.collidable=1
End sub

Sub sw26_hit: PlaySound "kicker_enter": End Sub
Sub sw27_hit: vpmTimer.PulseSw 27:PlaySound "kicker_enter": End Sub
Sub TrSubway_hit: PlaySound "fx_subway" : End sub

'******************************************************
'					TOP KICKER
'******************************************************

Sub UpperKicker(enabled)
If enabled Then
sw36.kick 0,35,1.56
Controller.Switch(36) = 0
Playsound SoundFX("vuk_exit",DOFContactors)
End If
End Sub

Sub sw36_hit: PlaySound "vuk_enter" : Controller.Switch(36) = 1 : End sub

'******************************************************
'			BALLSAVER (based on PacDude's)
'******************************************************

Sub BallSave(enabled) : If enabled Then PlaySound SoundFX(SSolenoidOn,DOFContactors) :BallSaver.enabled = True:Else BallSaver.enabled = False:End If: End Sub
Sub BallSaver_Hit() : PlaySound SoundFX(SSolenoidOn,DOFContactors) : BallSaver.kick 0,35+(rnd*20)  : BallSaver.enabled = False : End Sub

'******************************************************
'					AUTOPLUNGER
'******************************************************

Sub Shooter(Enabled)
    If Enabled Then
        Plunger.Fire
		PlaySound SoundFX("fx_plunger2",DOFContactors)
    End If
End Sub

'******************************************************
'					KNOCKER
'******************************************************

Sub SolKnocker(enabled)
	If enabled Then
		PlaySound SoundFX("knocker",DOFKnocker)
	End If
End Sub

'******************************************************
'				NFOZZY'S FLIPPERS
'******************************************************

Dim returnspeed, lfstep, rfstep
returnspeed = LeftFlipper.return
lfstep = 1
rfstep = 1

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySound SoundFX("fx_flipper1",DOFContactors), 0, 1, -0.1, 0.25
        LeftFlipper.RotateToEnd
        Flipper1.RotateToEnd
    Else
        PlaySound SoundFX("fx_flipperdown",DOFContactors), 0, 1, -0.1, 0.25
        LeftFlipper.RotateToStart
        Flipper1.RotateToStart
        LeftFlipper.TimerEnabled = 1
        LeftFlipper.TimerInterval = 16
        LeftFlipper.return = returnspeed * 0.5
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySound SoundFX("fx_flipper2",DOFContactors), 0, 1, 0.1, 0.25
        RightFlipper.RotateToEnd
    Else
        PlaySound SoundFX("fx_flipperdown",DOFContactors), 0, 1, 0.1, 0.25
        RightFlipper.RotateToStart
        RightFlipper.TimerEnabled = 1
        RightFlipper.TimerInterval = 16
        RightFlipper.return = returnspeed * 0.5
    End If
End Sub

Sub LeftFlipper_timer()
	select case lfstep
		Case 1: LeftFlipper.return = returnspeed * 0.6 :lfstep = lfstep + 1
		Case 2: LeftFlipper.return = returnspeed * 0.7 :lfstep = lfstep + 1
		Case 3: LeftFlipper.return = returnspeed * 0.8 :lfstep = lfstep + 1
		Case 4: LeftFlipper.return = returnspeed * 0.9 :lfstep = lfstep + 1
		Case 5: LeftFlipper.return = returnspeed * 1 :lfstep = lfstep + 1
		Case 6: LeftFlipper.timerenabled = 0 : lfstep = 1
	end select
end sub

Sub RightFlipper_timer()
	select case rfstep
		Case 1: RightFlipper.return = returnspeed * 0.6 :rfstep = rfstep + 1
		Case 2: RightFlipper.return = returnspeed * 0.7 :rfstep = rfstep + 1
		Case 3: RightFlipper.return = returnspeed * 0.8 :rfstep = rfstep + 1
		Case 4: RightFlipper.return = returnspeed * 0.9 :rfstep = rfstep + 1
		Case 5: RightFlipper.return = returnspeed * 1 :rfstep = rfstep + 1
		Case 6: RightFlipper.timerenabled = 0 : rfstep = 1
	end select
end sub

'******************************************************
'					BUMPERS
'******************************************************

Dim dirRing1:dirRing1 = -1
Dim dirRing2:dirRing2 = -1
Dim dirRing3:dirRing3 = -1

Sub Bumper1_Hit:PlaySound SoundFX("fx_bumper1",DOFContactors):Me.TimerEnabled = 1:vpmTimer.PulseSw 10:Me.TimerInterval = 20:End Sub
Sub Bumper2_Hit:PlaySound SoundFX("fx_bumper2",DOFContactors):Me.TimerEnabled = 1:vpmTimer.PulseSw 12:Me.TimerInterval = 20:End Sub
Sub Bumper3_Hit:PlaySound SoundFX("fx_bumper3",DOFContactors):Me.TimerEnabled = 1:vpmTimer.PulseSw 11:Me.TimerInterval = 20:End Sub

Sub Bumper1_timer()
	BRing1.Z = BRing1.Z + (5 * dirRing1)
	If BRing1.Z <= -10 Then dirRing1 = 1
	If BRing1.Z >= 25 Then
		dirRing1 = -1
		BRing1.Z = 25
		Me.TimerEnabled = 0
	End If
End Sub

Sub Bumper2_timer()
	BRing2.Z = BRing2.Z + (5 * dirRing2)
	If BRing2.Z <= -10 Then dirRing2 = 1
	If BRing2.Z >= 25 Then
		dirRing2 = -1
		BRing2.Z = 25
		Me.TimerEnabled = 0
	End If
End Sub

Sub Bumper3_timer()
	BRing3.Z = BRing3.Z + (5 * dirRing3)
	If BRing3.Z <= -10 Then dirRing3 = 1
	If BRing3.Z >= 25 Then
		dirRing3 = -1
		BRing3.Z = 25
		Me.TimerEnabled = 0
	End If
End Sub

'******************************************************
'				SLINGSHOTS
'******************************************************

Dim RStep, Lstep

Sub LeftSlingShot_Slingshot
	PlaySound SoundFX("SlingshotLeft",DOFContactors), 0, 1, -0.05, 0.05
	vpmTimer.PulseSw 13
    LSling.Visible = 0
    LSling1.Visible = 1
    sling1.TransZ = -20
    LStep = 0
    Me.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling1.TransZ = 0:Me.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
	PlaySound SoundFX("SlingshotRight",DOFContactors), 0, 1, 0.05, 0.05
	vpmTimer.PulseSw 14
    RSling.Visible = 0
    RSling1.Visible = 10
    sling2.TransZ = -2
    RStep = 0
    Me.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling2.TransZ = 0:Me.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

'******************************************************
'				SWITCHES
'******************************************************

'*******	Targets		******************
Sub sw20_Hit:vpmTimer.PulseSw 20:End Sub
Sub sw21_Hit:vpmTimer.PulseSw 21:End Sub
Sub sw21a_Hit:vpmTimer.PulseSw 21:End Sub
Sub sw22_Hit:vpmTimer.PulseSw 22:End Sub
Sub sw23_Hit:vpmTimer.PulseSw 23:End Sub
Sub sw24_Hit:vpmTimer.PulseSw 24:End Sub
Sub sw30_hit:DTdrop.Hit 1:End Sub

'*******	Opto Switches		******************
Sub sw17_Hit:Controller.Switch(17) = 1:End Sub
Sub sw17_UnHit:Controller.Switch(17) = 0:End Sub

'*******	Rollover Switches	******************

Sub sw32_Hit:Controller.Switch(32) = 1:PlaySound "rollover":End Sub
Sub sw32_UnHit:Controller.Switch(32) = 0:End Sub
Sub sw40_Hit:Controller.Switch(40) = 1:PlaySound "rollover":End Sub
Sub sw40_UnHit:Controller.Switch(40) = 0:End Sub
Sub sw41_Hit:Controller.Switch(41) = 1:PlaySound "rollover":End Sub
Sub sw41_UnHit:Controller.Switch(41) = 0:End Sub
Sub sw42_Hit:Controller.Switch(42) = 1:PlaySound "rollover":End Sub
Sub sw42_UnHit:Controller.Switch(42) = 0:End Sub
Sub sw43_Hit:Controller.Switch(43) = 1:PlaySound "rollover":End Sub
Sub sw43_UnHit:Controller.Switch(43) = 0:End Sub
Sub sw44_Hit:Controller.Switch(44) = 1:PlaySound "rollover":End Sub
Sub sw44_UnHit:Controller.Switch(44) = 0:End Sub
Sub sw46_Hit:Controller.Switch(46) = 1:PlaySound "rollover":End Sub
Sub sw46_UnHit:Controller.Switch(46) = 0:End Sub

'******************************************************
'				Castle Animation
'******************************************************

Dim discAngle:discAngle=0
Dim stepAngle, stopRotation

Sub CastleRotate(enabled)
	If Enabled Then
		CastleTimer.Enabled = 1
		stepAngle=15
		stopRotation=0
	Else
		stopRotation=1
	End If
End Sub

Sub CastleTimer_Timer()
	discAngle = discAngle + stepAngle
	If discAngle >= 360 Then
		discAngle = discAngle - 360
	End If
	Prim_Castle.RotZ = 360 - discAngle
	If stopRotation Then
		stepAngle = stepAngle - 0.1
		If stepAngle <= 0 Then
			CastleTimer.Enabled 	= 0
		End If
	End If
	Prim_CastleDecal.RotZ = Prim_Castle.RotZ
	Prim_CastleBolts.RotZ = -Prim_Castle.RotZ
End Sub

'******************************************************
'       		RealTime Updates
'******************************************************

Set MotorCallback = GetRef("GameTimer")

Sub GameTimer
    UpdateMechs
	RollingSoundUpdate
	BallShadowUpdate
End Sub

Sub UpdateMechs
	Prim_LeftFlipper.RotY=LeftFlipper.currentangle-90
	Prim_RightFlipper.RotY=RightFlipper.currentangle-90
	Prim_Flipper1.RotY=Flipper1.currentangle-90
	FlipperLSh.RotZ = LeftFlipper.currentangle
	FlipperRSh.RotZ = RightFlipper.currentangle
	FlipperL1Sh.RotZ = Flipper1.currentangle
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
    UpdateLamps
End Sub

Sub InitLamps()
    Dim x
    For x = 0 to 200
        LampState(x) = 0         ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4       ' used to track the fading state
        FlashSpeedUp(x) = 0.5    ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.35 ' slower speed when turning off the flasher
        FlashMax(x) = 1          ' the maximum value when on, usually 1
        FlashMin(x) = 0          ' the minimum value when off, usually 0
        FlashLevel(x) = 0        ' the intensity of the flashers, usually from 0 to 1
    Next
End Sub

Sub UpdateLamps
  	NFadeL 0, l0
  	NFadeL 1, l1
  	NFadeL 2, l2
  	NFadeL 3, l3
  	NFadeL 4, l4
  	NFadeL 5, l5
  	NFadeL 6, l6
  	NFadeL 7, l7

   	NFadeL 10, l10
  	NFadeL 11, l11
  	NFadeL 12, l12
  	NFadeL 13, l13
  	NFadeL 14, l14
  	NFadeL 15, l15
  	NFadeL 16, l16
  	NFadeL 17, l17

  	NFadeL 20, l20
  	NFadeL 21, l21
  	NFadeL 22, l22
  	NFadeL 23, l23
  	NFadeL 24, l24
  	NFadeL 25, l25
  	NFadeL 26, l26
  	NFadeL 27, l27

  	NFadeL 30, l30
  	NFadeL 31, l31
  	NFadeL 32, l32
  	NFadeL 33, l33
  	NFadeL 34, l34
  	NFadeL 35, l35
  	NFadeL 36, l36
  	NFadeL 37, l37

 	NFadeL 40, l40
 	NFadeL 41, l41
  	NFadeL 42, l42
  	NFadeL 43, l43
  	NFadeL 44, l44
   	NFadeLm 45, l45
   	NFadeLm 45, l45a
   	NFadeL 45, l45b
  	NFadeLm 46, l46
  	NFadeLm 46, l46a
  	NFadeL 46, l46b
  	NFadeLm 47, l47
  	NFadeLm 47, l47a
  	NFadeL 47, l47b

  	NFadeL 50, l50
  	NFadeL 51, l51
  	NFadeL 52, l52
  	NFadeL 53, l53
  	NFadeL 54, l54
  	NFadeL 55, l55
  	NFadeL 56, l56
  	NFadeL 57, l57

  	NFadeL 60, l60
  	NFadeL 61, l61
  	NFadeL 62, l62
  	NFadeL 63, l63
  	NFadeL 64, l64
  	NFadeL 65, l65
  	NFadeL 66, l66
  	NFadeL 67, l67

  	NFadeL 70, l70
  	NFadeL 71, l71
  	NFadeL 72, l72
  	NFadeL 73, l73
  	NFadeL 74, l74
  	NFadeL 75, l75
  	NFadeL 76, l76
  	NFadeL 77, l77

  	NFadeL 81, l81
  	NFadeL 82, l82
  	NFadeL 83, l83
  	NFadeL 84, l84
  	NFadeL 85, l85
  	NFadeL 86, l86
  	NFadeL 87, l87

  	NFadeL 90, l90
  	NFadeL 91, l91
  	NFadeL 92, l92
  	NFadeL 93, l93
  	NFadeL 94, l94
  	NFadeL 95, l95
  	NFadeL 96, l96
  	NFadeL 97, l97

'Mario speaking Flasher (x2)
NFadeLm 153, f13
Flash 153, f13a
NFadeL 154, f14
' castle/princess ??????
NFadeLm 155, f15
Flash 155, f15a
'luigi and yoshi????????
NFadeLm 156, f16
NFadeLm 156, f16a
Flashm 156, f16b
Flash 156, f16c
'Left Ramp Flasher
NFadeLm 157, f17
NFadeLm 157, f17a
NFadeLm 157, f17b
Flash 157,f17c
'Right Ramp Flasher
NFadeLm  158, f18
NFadeLm  158, f18a
Flashm  158, f18b
Flash  158, f18c
'Mario's Cave PF Flashers (2)
NFadeL 159, f19
'Castle Flashers
NFadeLm 160, f20
NFadeLm 160, f20a
NFadeLm 160, f20b
Flashm 160, f20c
Flash 160, f20d
'Left Kicker Flasher
NFadeL 161, f21
'Upper Left Flasher
NFadeLm 162, f22
NFadeLm 162, f22a
Flashm 162, f22b
Flash 162, f22c
'Upper Center Flashers (2)
NFadeLm 163, f23
NFadeLm 163, f23a
NFadeLm 163, f23b
NFadeLm 163, f23e
NFadeLm 163, f23f
Flashm 163, f23c
Flashm 163, f23d
Flash 163, f23g
'Upper Right Flashers (2)
NFadeLm 164, f24
NFadeLm 164, f24a
NFadeLm 164, f24b
NFadeLm 164, f24e
Flashm 164, f24c
Flashm 164, f24d
Flash 164, f24f
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

'******************************************************
' 			GENERAL ILLUMINATION
'******************************************************

Dim xx

Sub GIRelay(enabled)
	If Enabled Then
 		for each xx in GI:xx.State=0:Next
 		for each xx in GIbulbs:xx.intensityscale=0:Next
		Table1.ColorGradeImage = "ColorGrade_1"
		Playsound "fx_relay_off"
	Else
 		for each xx in GI:xx.State=1:Next
 		for each xx in GIbulbs:xx.intensityscale=1:Next
		Table1.ColorGradeImage = "ColorGrade_8"
		Playsound "fx_relay_on"
	End If
End Sub

Sub TiltRelay(enabled)
	If Enabled Then
 		for each xx in GI:xx.State=0:Next
 		for each xx in GIbulbs:xx.intensityscale=0:Next
		Table1.ColorGradeImage = "ColorGrade_1"
		Playsound "fx_relay_off"
	Else
 		for each xx in GI:xx.State=1:Next
 		for each xx in GIbulbs:xx.intensityscale=1:Next
		Table1.ColorGradeImage = "ColorGrade_8"
		Playsound "fx_relay_on"
	End If
End Sub

'*****************************************
'	Ball Shadow
'*****************************************

Dim BallShadow
BallShadow = Array (BallShadow1, BallShadow2, BallShadow3)

Sub BallShadowUpdate()
    Dim BOT, b, shadowZ
    BOT = GetBalls

	' render the shadow for each ball
    For b = 0 to UBound(BOT)
		If BOT(b).X < Table1.Width/2 Then
			BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) + 10
		Else
			BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) - 10
		End If

		If BOT(b).Z < 10 and BOT(b).Z>90 then BallShadow(b).Z = 1 else BallShadow(b).Z=BOT(b).z-20

			BallShadow(b).Y = BOT(b).Y + 40
		If BOT(b).Z > 20 Then
			BallShadow(b).visible = 1
		Else
			BallShadow(b).visible = 0
		End If
	Next
End Sub

'******************************************************
' 			JP's VP10 Ball Collision Sound
'******************************************************

Sub OnBallBallCollision(ball1, ball2, velocity)
	PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0
End Sub

'******************************************************
' 				JF's Sound Routines
'******************************************************

Sub Pins_Hit (idx)
	PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
	PlaySound SoundFX("target",DOFContactors)
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

Sub Flipper1_Collide(parm)
 	RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub

'******************************************************
' 				Balldrop & Ramp Sound
'******************************************************

Sub BallHitSound(dummy):PlaySound "fx_ball_bounce":End Sub
Sub LRampEnter_Hit():Playsound "ramp_enter":End Sub
Sub RRampEnter_Hit():Playsound "ramp_enter":End Sub
Sub RRampEnter1_Hit():Playsound "ramp_enter":End Sub

Sub LMRampRolling_Hit():Playsound "WireRamp":End Sub
Sub RRampRolling_Hit():Playsound "ramp_rolling":End Sub

Sub CenterDrop_Hit():If activeball.z > 60  Then vpmTimer.AddTimer 200, "BallHitSound":End If:End Sub
Sub LMRamp_Hit():StopSound "WireRamp":vpmTimer.AddTimer 150, "BallHitSound":End Sub
Sub RRamp_Hit():StopSound "ramp_rolling":vpmTimer.AddTimer 150, "BallHitSound":End Sub

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

'******************************************************
'      		JP's VP10 Rolling Sounds
'******************************************************

Const tnob = 3 ' total number of balls
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub RollingSoundUpdate()
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


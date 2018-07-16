Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const Ballsize = 55
Const BallMass = 1.6

LoadVPM "01120100", "DE.VBS", 3.36

'********************
'Standard definitions
'********************

Const UseSolenoids = 2
Const UseLamps = 0
Const UseSync = 1
Const HandleMech = 0

' Standard Sounds
Const SSolenoidOn = "fx_Solenoid"
Const SSolenoidOff = "fx_solenoidoff"
Const SCoin = "fx_Coin"

'Solenoids
SolCallback(1)  = "bsTrough.SolIn"									'6-ball lockout								(1L)
SolCallback(2)  = "bsTrough.SolOut"									'ball eject									(2L)
SolCallback(3)  = "bsLScoop.SolOut"									'Left Scoop                          		(3L)
SolCallback(4)  = "SolAutoPlungerIM"								'Autolaunch                             	(4L)
'NOT USED																										(5L)
SolCallback(6)	= "bsRScoop.SolOut"									'Right Scoop		                    	(6L)
'NOT USED																										(7L)
SolCallback(8)  = "Solknocker"										'Knocker							  		(8L)
SolCallback(9)  = "SetLamp 109,"        							'FlashLamp x4 (2 backbox + 2 pf)			(09)
'SolCallback(10) = ""												'L/R Relay 									(10)
SolCallback(11) = "SolGi"											'GI Relay                             		(11)
SolCallback(12) = "SetLamp 112,"                           			'FlashLamp x4 (1 pf + 3 backbox)			(12)
SolCallback(13) = "SetLamp 113,"									'FlashLamp x4 (2 pf + 2 backbox)        	(13)
SolCallBack(14) = "SetLamp 114,"                          			'FlashLamp x4 (1 pf + 3 backbox)			(14)
'SolCallBack(15) = ""												'Ticket Dispenser							(15)
'SolCallBack(16) = ""												'Bar Motor									(16)
'SolCallBack(17) = ""												'Left Bumper								(17)
'SolCallBack(18) = ""												'Center Bumper								(18)
'SolCallBack(19) = ""			 									'Right Bumper								(19)
'SolCallBack(20) = ""												'Left Slingshot								(20)
'SolCallBack(21) = ""												'Right Slingshot							(21)
SolCallback(22) = "SolDiv"                           				'Ramp Diverter                      		(22)
'NOT USED																										(23)
'NOT USED																										(24)
SolCallback(25) = "SetLamp 125,"									'Flashlamp X4 (3 pf + backbox        		(1R)
SolCallback(26) = "SetLamp 126,"									'Flashlamp X4 (1 pf + 2 ramp + 1 backbox)	(2R)
SolCallback(27) = "SetLamp 127,"									'Flashlamp X4 (2 backbox + 2 pf)        	(3R)
SolCallback(28) = "SetLamp 128,"									'Flashlamp X4 (2 backbox + 2 pf) 			(4R)
SolCallback(29) = "SetLamp 129,"									'Flashlamp X4 (4 pf)						(5R)
SolCallback(30) = "SetLamp 130,"									'Flashlamp X4 (3 pf + 1 backbox)			(6R)
SolCallback(31) = "SetLamp 131,"									'Flashlamp X4 (3 pf + 1 backbox)			(7R)
SolCallBack(32) = "SetLamp 132,"     								'Flashlamp X4 (2 bat + 2 backbox)			(8R)

SolCallback(46) = "SolRFlipper"                         			'Right Flipper
SolCallback(48) = "SolLFlipper"                         			'Left Flipper

'************
' Table init.
'************

Const cGameName = "btmn_106"

Dim bsTrough, plungerIM, bsLScoop, bsRScoop, mBar

Sub Table1_Init
    vpmInit Me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Batman, Data East 1991" & vbNewLine & "VPX table by Javier v1.0"
		.HandleKeyboard = 0
		.ShowTitle = 0
		.ShowDMDOnly = 1
		.ShowFrame = 0
		.HandleMechanics = 0
		.Hidden = 0
		.Games(cGameName).Settings.Value("sound") = 1
		On Error Resume Next
		.Run GetPlayerHWnd
		If Err Then MsgBox Err.Description
	End With
    On Error Goto 0

    ' Nudging
    vpmNudge.TiltSwitch = 1
	vpmNudge.Sensitivity = 2 
    vpmNudge.tiltobj = Array(LeftSlingShot,RightSlingShot,Bumper1B,Bumper2B,Bumper3B)

	PinMAMETimer.Interval = PinMAMEInterval
	PinMAMETimer.Enabled = 1

	' Trough
     Set bsTrough = new cvpmTrough 
     With bsTrough
		.Size = 3
		.InitSwitches Array (13,12,11)
		.EntrySw = 10
		.InitExit BallRelease, 90, 6
		.InitEntrySounds "fx_drain", SoundFX(SSolenoidOn,DOFContactors), SoundFX(SSolenoidOn,DOFContactors)
		.InitExitSounds  SoundFX(SSolenoidOn,DOFContactors), SoundFX("fx_ballrel",DOFContactors)
		.Balls = 3
		.CreateEvents "bsTrough", Drain
     End With

     ' Scoop Left
	Set bsLScoop = New cvpmSaucer
	With bsLScoop
	    .InitKicker Sw39b, 39,50, 30, 1.56
		.InitSounds "scoopenter", SoundFX(SSolenoidOn,DOFContactors), SoundFX("salidadebolaL",DOFContactors)
		.CreateEvents "bsLScoop", sw39b
    End With

     ' Scoop Right
	Set bsRScoop = New cvpmTrough
	With bsRScoop
		.Size = 2
		.InitSwitches Array (52,53)
		.InitExit Sw52, 192, 25
		.InitEntrySounds "fx_chapa", SoundFX(SSolenoidOn,DOFContactors), SoundFX(SSolenoidOn,DOFContactors)
		.InitExitSounds  SoundFX(SSolenoidOn,DOFContactors), SoundFX("salidadebolaR",DOFContactors)
		.Balls = 0
    End With

    ' Impulse Plunger
    Const IMPowerSetting = 55
    Const IMTime = 0.6
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swplunger, IMPowerSetting, IMTime
        .Random 0.3
        .switch 14
        .InitExitSnd SoundFX("bumper_retro",DOFContactors), SoundFX("fx_target",DOFContactors)
        .CreateEvents "plungerIM"
    End With

	' Bar Motor
     set mBar = new cvpmMech
     with mBar
         .MType = vpmMechOneSol + vpmMechReverse + vpmMechNonLinear 
         .Sol1 = 16
         .Length = 130
         .Steps = 50
		 .addsw 50,0,0
		 .addsw 51,47,49
		 .acc=0
		 .ret=0
         .Callback = GetRef("UpdateBar")
         .Start
     End with

      Controller.Switch(50) = 0 
      Controller.Switch(51) = 1

	  RampDiverter.Isdropped = 1

   If Table1.ShowDT = False then
       Ramp26.visible = 0
       Ramp32.visible = 0
       Ramp33.visible = 0
       Ramp34.visible = 0
       Ramp35.visible = 0
       Ramp36.visible = 0
   End If

End Sub

'******************
'Keys Up and Down
'*****************

Sub Table1_KeyDown(ByVal Keycode)
	If keycode = PlungerKey Then Plunger.Pullback : PlaySoundAt "plungerpull",Plunger
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal Keycode)
If keycode = PlungerKey Then Plunger.Fire : PlaySoundAt "plunger",Plunger
    If vpmKeyUp(keycode) Then Exit Sub
End Sub

Sub Table1_Paused : Controller.Pause = True : End Sub
Sub Table1_unPaused : Controller.Pause = False : End Sub
Sub Table1_Exit() : Controller.Pause = False : Controller.Stop() : End Sub

'*****************
'Solenoids
'*****************

'AutoPlunger
Sub SolAutoPlungerIM(Enabled)
    If Enabled Then
        PlungerIM.AutoFire
    End If
End Sub

'Knocker
Sub SolKnocker(enabled)
If enabled then Playsound SoundFX("fx_knocker",DOFKnocker),0,2,0,0,0,1,0,-.8
End Sub

'CavTargets
Dim CPos
 Sub UpdateBar(acurrpos, aspeed, alastpos)
     CPos = acurrpos
     StopSound"motor": PlaySoundAt "motor",FlasherLight8
     Sw49a.TransY = 50 - CPos
     Sw49b.TransY = 50 - CPos
     If CPos => 45 Then Sw49.isdropped = 1
     If CPos =< 25 Then Sw49.isdropped = 0     
 End Sub

'Ramp Diverter
Sub SolDiv(enabled)
	If enabled Then
		RampDiverter.Isdropped = 0
        PlaySoundAt "fx_diverter",Trigger2
 		RampDiverter2.Isdropped = 1 
        Primitive_RampDiverter.RotY = 12
	Else
		RampDiverter.Isdropped = 1 
        StopSound "fx_Ramp"
        PlaySoundAt "fx_diverter",Trigger2
 		RampDiverter2.Isdropped = 0
        Primitive_RampDiverter.RotY = 0
	End If
End Sub

' Flippers
Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("FlipperUpLeft", DOFFlippers),LFLIPsound
        LeftFlipper.RotateToEnd
    Else
        PlaySoundAt SoundFX("fx_Flipperdown", DOFFlippers),LFLIPsound
        LeftFlipper.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("FlipperUpRight", DOFFlippers),RFLIPsound
        RightFlipper.RotateToEnd
    Else
        PlaySoundAt SoundFX("fx_Flipperdown", DOFFlippers),RFLIPsound
        RightFlipper.RotateToStart
    End If
End Sub

'*****************
' Switches
'*****************

' Slingshots
Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    PlaySoundAt SoundFX("left_slingshot", DOFContactors),Lemk
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 47
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 1:LeftSLing4.Visible = 0:LeftSLing3.Visible = 1:Lemk.RotX = 14
        Case 2:LeftSLing3.Visible = 0:LeftSLing2.Visible = 1:Lemk.RotX = 2
        Case 3:LeftSLing2.Visible = 0:Lemk.RotX = -10:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
    PlaySoundAt SoundFX("right_slingshot", DOFContactors),Remk
    RightSling4.Visible = 1
    Remk.RotX = 26
    RStep = 0
    vpmTimer.PulseSw 48
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 1:RightSLing4.Visible = 0:RightSLing3.Visible = 1:Remk.RotX = 14
        Case 2:RightSLing3.Visible = 0:RightSLing2.Visible = 1:Remk.RotX = 2
        Case 3:RightSLing2.Visible = 0:Remk.RotX = -10:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

' Bumpers
Sub Bumper1B_Hit:vpmTimer.PulseSw 54:PlaySoundAtBumperVol SoundFX("fx_bumper1",DOFContactors),Bumper1B,1:End Sub
Sub Bumper2B_Hit:vpmTimer.PulseSw 56:PlaySoundAtBumperVol SoundFX("fx_bumper3",DOFContactors),Bumper2B,1:End Sub
Sub Bumper3B_Hit:vpmTimer.PulseSw 55:PlaySoundAtBumperVol SoundFX("fx_bumper2",DOFContactors),Bumper3B,1:End Sub

'Top Lanes
Sub Sw17_Hit():PlaySoundAtVol "fx_sensor",sw17,2:Controller.Switch(17)=1: End Sub
Sub Sw17_UnHit():Controller.Switch(17)=0: End Sub

Sub Sw18_Hit():PlaySoundAtVol "fx_sensor",sw18,2:Controller.Switch(18)=1: End Sub
Sub Sw18_UnHit():Controller.Switch(18)=0: End Sub

Sub Sw19_Hit():PlaySoundAtVol "fx_sensor",Sw19,2:Controller.Switch(19)=1: End Sub
Sub Sw19_UnHit():Controller.Switch(19)=0: End Sub


Sub Sw21_Hit():PlaySoundAtVol "fx_sensor",sw21,2:Controller.Switch(21)=1: End Sub
Sub Sw21_UnHit():Controller.Switch(21)=0: End Sub

Sub Sw22_Hit():PlaySoundAtVol "fx_sensor",sw22,2:Controller.Switch(22)=1: End Sub
Sub Sw22_UnHit():Controller.Switch(22)=0: End Sub

Sub Sw23_Hit():PlaySoundAtVol "fx_sensor",sw23,2:Controller.Switch(23)=1: End Sub
Sub Sw23_UnHit():Controller.Switch(23)=0: End Sub

Sub Sw24_Hit():PlaySoundAtVol "fx_sensor",sw24,2:Controller.Switch(24)=1: End Sub
Sub Sw24_UnHit():Controller.Switch(24)=0: End Sub

'Center Ramp
Sub Sw28_Hit():PlaySoundAtVol "fx_sensor",sw28,2:Controller.Switch(28)=1: End Sub
Sub Sw28_UnHit():Controller.Switch(28)=0: End Sub

'Center Ramp Exit
Sub Sw29_Hit()
	If ActiveBall.VelY > 0 Then
		PlaySoundAt"plastic_ramp",sw29
	Else
		StopSound"plastic_ramp"
	End If
   Controller.Switch(29)=1
 End Sub
Sub Sw29_UnHit():Controller.Switch(29)=0: End Sub

' Left Banks Targets
Sub Sw33_Hit:vpmTimer.PulseSw 33 :MoveTarget33 :PlaySoundAtBallVol SoundFX("target",DOFContactors),3:End Sub
Sub Sw34_Hit:vpmTimer.PulseSw 34 :MoveTarget34 :PlaySoundAtBallVol SoundFX("target",DOFContactors),3:End Sub
Sub Sw35_Hit:vpmTimer.PulseSw 35 :MoveTarget35 :PlaySoundAtBallVol SoundFX("target",DOFContactors),3:End Sub

Sub MoveTarget33
	Sw33a.TransZ = 5
	Sw33b.TransZ = 5
	Sw33.Timerenabled = False
	Sw33.Timerenabled = True
End Sub
Sub	Sw33_Timer
	Sw33.Timerenabled = False
	Sw33a.TransZ = 0
	Sw33b.TransZ = 0
End Sub

Sub MoveTarget34
	Sw34a.TransZ = 5
	Sw34b.TransZ = 5
	Sw34.Timerenabled = False
	Sw34.Timerenabled = True
End Sub
Sub	Sw34_Timer
	Sw34.Timerenabled = False
	Sw34a.TransZ = 0
	Sw34b.TransZ = 0
End Sub

Sub MoveTarget35
	Sw35a.TransZ = 5
	Sw35b.TransZ = 5
	Sw35.Timerenabled = False
	Sw35.Timerenabled = True
End Sub
Sub	Sw35_Timer
	Sw35.Timerenabled = False
	Sw35a.TransZ = 0
	Sw35b.TransZ = 0
End Sub

'Joker Eyes and Mouth
Sub Sw36_Hit: PlaySoundAtVol "kicker_hit",Sw36,.7: vpmTimer.PulseSw (36) : Sw39.enabled = 0: End Sub
Sub Sw37_Hit: PlaySoundAtVol "kicker_hit",Sw37,7: vpmTimer.PulseSw (37) : Sw39.enabled = 0: End Sub
Sub Sw38_Hit: PlaySoundAtVol "kicker_hit",Sw38,7: vpmTimer.PulseSw (38) : Sw39.enabled = 0: End Sub

'Left Scoop
Sub Sw39_hit: Sw39.enabled = 0 : End Sub
Sub Sw39b_UnHit: vpmtimer.addtimer 200, "Sw39.enabled = 1 '" :End Sub

' Right Banks Targets
Sub Sw41_Hit:vpmTimer.PulseSw 41 :MoveTarget41 :PlaySoundAtBallVol SoundFX("target",DOFContactors),3:End Sub
Sub Sw42_Hit:vpmTimer.PulseSw 42 :MoveTarget42 :PlaySoundAtBallVol SoundFX("target",DOFContactors),3:End Sub
Sub Sw43_Hit:vpmTimer.PulseSw 43 :MoveTarget43 :PlaySoundAtBallVol SoundFX("target",DOFContactors),3:End Sub

Sub MoveTarget41
	Sw41a.TransZ = 5
	Sw41b.TransZ = 5
	Sw41.Timerenabled = False
	Sw41.Timerenabled = True
End Sub
Sub	Sw41_Timer
	Sw41.Timerenabled = False
	Sw41a.TransZ = 0
	Sw41b.TransZ = 0
End Sub

Sub MoveTarget42
	Sw42a.TransZ = 5
	Sw42b.TransZ = 5
	Sw42.Timerenabled = False
	Sw42.Timerenabled = True
End Sub
Sub	Sw42_Timer
	Sw42.Timerenabled = False
	Sw42a.TransZ = 0
	Sw42b.TransZ = 0
End Sub

Sub MoveTarget43
	Sw43a.TransZ = 5
	Sw43b.TransZ = 5
	Sw43.Timerenabled = False
	Sw43.Timerenabled = True
End Sub
Sub	Sw43_Timer
	Sw43.Timerenabled = False
	Sw43a.TransZ = 0
	Sw43b.TransZ = 0
End Sub

'Bat Bar StandUp
Sub Sw49_Hit
    vpmTimer.PulseSw 49
    PlaySoundAt "fx_chapa",gi16
End Sub

Sub Sw52_Hit() : PlaySoundAt "trough",Sw52: Me.DestroyBall:vpmTimer.Addtimer 1000, "bsRScoop.AddBall" : End Sub

'***************************************************
'	General Illumination
'***************************************************
Dim x

Sub SolGi(enabled)
  If enabled Then
     GiOFF
	PlaySoundAtVol "fx_relay_off",l18f,2
	Table1.ColorGradeImage = "ColorGradeLUT256x16_ConSatDark"
   Else
     GiON
	PlaySoundAtVol "fx_relay_on",l18f,2
	Table1.ColorGradeImage = "ColorGradeLUT256x16_ConSat"
 End If
End Sub

Sub GiON
    For each x in aGiLights:x.State = 1:Next
    gi29.IntensityScale=1 : gi30.IntensityScale=1
    Primitive_PlasticRamp.Image = "BatRampMapResized"
End Sub

Sub GiOFF
    For each x in aGiLights:x.State = 0:Next
    gi29.IntensityScale=0 : gi30.IntensityScale=0
    Primitive_PlasticRamp.Image = "BatmanRampMap_resizedOFF"
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
	nFadeL 1, l1
	nFadeL 2, l2
    nFadeL 3, l3
    nFadeL 4, l4
    nFadeL 5, l5
    nFadeL 6, l6
    nFadeL 7, l7
    nFadeL 8, l8
    nFadeL 9, l9
    nFadeL 10, l10
 
    nFadeL 11, l11
    nFadeL 12, l12
    nFadeL 13, l13
    Flash 14, l14
    nFadeL 15, l15
    nFadeL 16, l16
    NFadeLm 17, l17
    Flash 17, L17f
    NFadeLm 18, l18
    Flash 18, L18f
    NFadeLm 19, l19
    Flash 19, L19f
    nFadeL 20, l20
 
    nFadeL 21, l21
    nFadeL 22, l22
    NFadeLm 23, l23
    Flash 23, l23f
    nFadeL 24, l24

	nFadeLm 25, FlasherLight25a
	Flashm 25, FlasherL25
	Flash 25, FlasherLight25d

	nFadeLm 26, FlasherLight26a
	Flashm 26, FlasherL26
	Flash 26, FlasherLight26d

	nFadeLm 27, FlasherLight27a
	Flashm 27, FlasherL27
	Flash 27, FlasherLight27d

	nFadeLm 28, FlasherLight28a
	Flashm 28, FlasherL28
	Flash 28, FlasherLight28d

	nFadeLm 29, FlasherLight29a
	Flashm 29, FlasherL29
	Flash 29, FlasherLight29d

    nFadeL 30, l30
 
    nFadeL 31, l31
    nFadeL 32, l32
    nFadeL 33, l33
    nFadeL 34, l34
    nFadeL 35, l35

	nFadeLm 36, l36b
	Flashm 36, l36
	Flash 36, l36Fb

	nFadeLm 37, l37b
	Flash 37, l37

	nFadeLm 38, L38
	Flashm 38, l38f
	Flashm 38, FlasherL38
	Flash 38, FlasherL38r

    nFadeL 39, l39
    nFadeL 40, l40
 
    nFadeL 41, l41
    nFadeL 42, l42
    nFadeL 43, l43

	nFadeLm 44, BumperL_Flasher
	nFadeLm 44, BumperL_Flasher_a
	Flash 44, BumperL_Flasher1

	nFadeLm 45, BumperB_Flasher
	nFadeLm 45, BumperB_Flasher_a
	Flash 45, BumperB_Flasher1

	nFadeLm 46, BumperR_Flasher
	nFadeLm 46, BumperR_Flasher_a
	Flash 46, BumperR_Flasher1

    NFadeLm 47, l47
    Flash 47, l47f  
    nFadeL 48, l48

	nFadeLm 49, l49
	Flashm 49, FlasherL49
	Flash 49, l49f

	Flash 55, L55
    NFadeLm 56, l56
    Flash 56, BatCave3
    nFadeL 57, l57
    nFadeL 58, l58
    nFadeL 59, l59
    nFadeL 60, l60
    nFadeL 61, l61
    nFadeL 62, l62
    NFadeLm 63, l63
    Flash 63, BatCave1
    NFadeLm 64, l64 
    Flash 64, BatCave2

	'Flashers
	NFadeObjm 109, Primitive53, "BatCave_bothON", "BatCave_CompleteMap"
	nFadeLm 109, FlasherL9
	nFadeLm 109, FlasherL9b
	Flashm 109, Flasher9
	Flash 109, Flasher9a
	Flash 112, Flasher12
	Flash 113, FlasherL38
	Flash 114, Flasher14

	NFadeLm 125, FlasherL1b
	Flash 125, Flasher1b

	NFadeObjm 126, Primitive53, "BatCave_bothON", "BatCave_CompleteMap"
	NFadeLm 126, FlasherLight2
	NFadeLm 126, FlasherLight2b
	Flash 126, Flasher2b

	NFadeL 127, FlasherLight3
	NFadeL 128, FlasherLight4

	NFadeLm 129, FlasherL1
	NFadeLm 129, FlasherL1a
	NFadeLm 129, FlasherL2
	NFadeLm 129, FlasherL2a
	NFadeLm 129, FlasherL3
	NFadeLm 129, FlasherL3a
	NFadeLm 129, FlasherL4
	NFadeLm 129, FlasherL4a
	Flashm 129, Flasher1
	Flashm 129, FlasherL1c
	Flashm 129, Flasher2
	Flashm 129, FlasherL2c
    Flashm 129, Flasher3
    Flashm 129, FlasherL3c
    Flashm 129, Flasher4
    Flash 129, FlasherL4c

    NFadeLm 130, Flasher6
    NFadeLm 130, Flasher6a
    NFadeLm 130, Flasher6b
    NFadeL 130, Flasher6c

	NFadeLm 131, FlasherLight7a
	NFadeLm 131, FlasherLight7b
	NFadeL 131, FlasherLight7c

	NFadeObjm 132, Primitive_museum, "MuseumMap_on", "MuseumMap_off"
	nFadeLm 132, FlasherLight8
	Flashm 132, Flasher8
	Flash 132, Flasher8a

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

' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 400)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / Table1.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10))
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
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2)))
End Function

'****************************************
' Real Time updates
'****************************************

Set MotorCallBack = GetRef("GameTimer")

Sub GameTimer
    RollingUpdate
	BallShadowUpdate
    LFLogo.RotY = LeftFlipper.CurrentAngle
    RFlogo.RotY = RightFlipper.CurrentAngle
End Sub

'*********** ROLLING SOUND *********************************
Const tnob = 3  ' total number of balls
ReDim rolling(tnob)
InitRolling

Sub InitRolling:Dim i:For i=0 to (tnob-1):rolling(i) = False:Next:End Sub

Sub RollingUpdate()
    Dim BOT, b, ballpitch
    BOT = GetBalls
	' stop the sound of deleted balls
	If UBound(BOT)<(tnob -1) Then
		For b = (UBound(BOT) + 1) to (tnob-1)
			rolling(b) = False
			StopSound("fx_ballrolling" & b)
		Next
	End If
	' exit the Sub if no balls on the table
    If UBound(BOT) = -1 Then Exit Sub
       ' play the rolling sound for each ball
    For b = 0 to UBound(BOT)
        If BallVel(BOT(b) ) > 1 Then
			rolling(b) = True
			if BOT(b).z < 30 Then ' Ball on playfield
						PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
			Else ' Ball on raised ramp
						PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*.9, Pan(BOT(b) ), 0, Pitch(BOT(b) )+30000, 1, 0, AudioFade(BOT(b) )
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
		PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1) 
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


'Set position as bumperX and Vol manually.

Sub PlaySoundAtBumperVol(sound, tableobj, Vol)
		PlaySound sound, 1, Vol, Pan(tableobj), 0,0,1, 1, AudioFade(tableobj)
End Sub


'*****************************************
'	ninuzzu's	FLIPPER SHADOWS
'*****************************************

sub FlipperTimer_Timer()
	FlipperLSh.RotZ = LeftFlipper.currentangle
	FlipperRSh.RotZ = RightFlipper.currentangle

End Sub

'*********** BALL SHADOW *********************************
Dim BallShadow
BallShadow = Array (BallShadow1, BallShadow2, BallShadow3)

Sub BallShadowUpdate()
    Dim BOT, b
    BOT = GetBalls
	' hide shadow of deleted balls
	If UBound(BOT)<(tnob-1) Then
		For b = (UBound(BOT) + 1) to (tnob-1)
			BallShadow(b).visible = 0
		Next
	End If
	' exit the Sub if no balls on the table
    If UBound(BOT) = -1 Then Exit Sub
	' render the shadow for each ball
    For b = 0 to UBound(BOT)
		If BOT(b).X < Table1.Width/2 Then
			BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) + 10
		Else
			BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) - 10
		End If
	    ballShadow(b).Y = BOT(b).Y + 20

		If BOT(b).Z > 20 Then
			BallShadow(b).visible = 1
		Else
			BallShadow(b).visible = 0
		End If
	Next
End Sub

'******************************
'		 Sound FX
'******************************

Sub aMetals_Hit(idx):PlaySoundAtBall "metalhit2":End Sub
Sub aGates_Hit(idx):PlaySoundAtBallVol "fx_Gate",1:End Sub
Sub aWoods_Hit(idx):PlaySoundAtBall "fx_woodhit":End Sub

Sub KickerHelper_Hit:PlaySoundAt "ball_drop2",kickerHelper :End Sub
Sub UpLaneKickers1_Hit():vpmtimer.addtimer 150, "UpKickers '" End Sub
Sub UpLaneKickers2_Hit():vpmtimer.addtimer 150, "UpKickers '" End Sub
Sub UpLaneKickers3_Hit():vpmtimer.addtimer 150, "UpKickers '" End Sub
Sub RampDrop_Hit():vpmtimer.addtimer 150, "UpKickers2 '" End Sub

Sub UpKickers()
    PlaySound "ball_drop",0,1,0,0,0,1,0,-.8
End Sub

Sub UpKickers2()
    PlaySound "ball_drop",0,1,.15,0,0,1,0,-.6
End Sub

'************************************
'  Random Sounds
'************************************


Sub aRubber_Bands_Hit(idx)
    PlaySoundAtBallVol "rubber_hit_" & Int(Rnd*3)+1, 1
End Sub

Sub aRubber_Posts_Hit(idx)
    PlaySoundAtBallVol "rubber_hit_" & Int(Rnd*3)+1, 1
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySoundAtBallVol "flip_hit_" & Int(Rnd*3)+1, 1
End Sub

Sub RightFlipper_Collide(parm)
    PlaySoundAtBallVol "flip_hit_" & Int(Rnd*3)+1, 1
End Sub

Dim NextOrbitHit:NextOrbitHit = 0 

Sub WireRampBumps_Hit(idx)
	if BallVel(ActiveBall) > .3 and Timer > NextOrbitHit then
		RandomBump3 .5, Pitch(ActiveBall)+5
		' Schedule the next possible sound time.  This prevents it from rapid-firing noises too much. 
		' Lowering these numbers allow more closely-spaced clunks.
		NextOrbitHit = Timer + .2 + (Rnd * .2)
	end if 
End Sub

Sub PlasticRampBumps_Hit(idx)
	if BallVel(ActiveBall) > .3 and Timer > NextOrbitHit then
		RandomBump 2, Pitch(ActiveBall)
		' Schedule the next possible sound time.  This prevents it from rapid-firing noises too much. 
		' Lowering these numbers allow more closely-spaced clunks.
		NextOrbitHit = Timer + .1 + (Rnd * .2)
	end if 
End Sub

Sub PlasticRampBumps2_Hit(idx)
	if BallVel(ActiveBall) > .3 and Timer > NextOrbitHit then
		RandomBump 4, Pitch(ActiveBall)
		' Schedule the next possible sound time.  This prevents it from rapid-firing noises too much. 
		' Lowering these numbers allow more closely-spaced clunks.
		NextOrbitHit = Timer + .1 + (Rnd * .2)
	end if 
End Sub


Sub MetalGuideBumps_Hit(idx)
	if BallVel(ActiveBall) > .3 and Timer > NextOrbitHit then
		RandomBump2 2, Pitch(ActiveBall)
		' Schedule the next possible sound time.  This prevents it from rapid-firing noises too much. 
		' Lowering these numbers allow more closely-spaced clunks.
		NextOrbitHit = Timer + .2 + (Rnd * .2)
	end if 
End Sub

Sub MetalWallBumps_Hit(idx)
	if BallVel(ActiveBall) > .3 and Timer > NextOrbitHit then
		RandomBump 2, 20000 'Increased pitch to simulate metal wall
		' Schedule the next possible sound time.  This prevents it from rapid-firing noises too much. 
		' Lowering these numbers allow more closely-spaced clunks.
		NextOrbitHit = Timer + .2 + (Rnd * .2)
	end if 
End Sub


'' Requires rampbump1 to 7 in Sound Manager
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
Sub BumpSTOPplastic1_Hit ()
dim i:for i=1 to 7:StopSound "rampbump" & i:next
NextOrbitHit = Timer + 1
End Sub


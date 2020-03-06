'Title Fight VP10 table by bodydump
Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

' Thalamus 2020 February : Improved directional sounds
' !! NOTE : Table not verified yet !!
' Please note, there are a few samples that was referred in the table, but, didn't have corresponding samples.

' Options
' Volume devided by - lower gets higher sound

Const VolDiv = 2000    ' Lower number, louder ballrolling/collition sound
Const VolCol = 10      ' Ball collition divider ( voldiv/volcol )

Const BallSize = 50

If Table1.ShowDT = true then
	Wall28.IsDropped = False
	Wall30.IsDropped = False
	Ramp15.visible = True
	Ramp16.visible = True
	Ramp17.visible = True
	Boxers.visible = True
	For each xx in Digitss:xx.visible = 1: Next
else
	Wall28.IsDropped = True
	Wall30.IsDropped = True
	Ramp15.visible = False
	Ramp16.visible = False
	Ramp17.visible = False
	Boxers.visible = False
	For each xx in Digitss:xx.visible = 0: Next
end if

LoadVPM "01120100", "gts3.vbs", 3.26

'********************
'Standard definitions
'********************

Const UseSolenoids = 1
Const UseLamps = 0
Const UseSync = 1
Const UseGI = 0
Const HandleMech = 0

' Standard Sounds
Const SSolenoidOn = "fx_Solenoid"
Const SSolenoidOff = "fx_solenoidoff"
Const SCoin = "fx_Coin"

Dim bsTrough, dt5bank, bsKickerHole, dt3Bank, cbLeft, cbRight
GIoff
'************
' Table init.
'************
Dim VarHidden

Const cGameName = "tfight" 'Title Fight


Sub Table1_Init
    vpmInit Me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Title Fight" & vbNewLine & "VP10 table by bodydump"
        .Games(cGameName).Settings.Value("sound") = 1: 'ensure the sound is on
        .Games(cGameName).Settings.Value("rol") = 0
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = VarHidden
        On Error Resume Next
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
        On Error Goto 0
    End With

    ' Nudging
    vpmNudge.TiltSwitch = 151                                                 'plumb tilt
    vpmNudge.Sensitivity = 1
    vpmNudge.TiltObj = Array(Bumper1, LeftSlingShot) 'bumpers & slingshots

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 46, 56, 0, 0, 0, 0, 0, 0
        .InitKick BallRelease, 90, 8
        .InitEntrySnd "fx_Solenoid", "fx_Solenoid"
        .InitExitSnd SoundFX("fx_ballrel",DOFContactors), SoundFX("fx_Solenoid",DOFContactors)
        .Balls = 3
    End With

	'Kicker
		Set bsKickerHole = New cvpmBallStack
	With bsKickerHole
		.InitSaucer Kicker1,36,210,7
		.InitEntrySnd "kicker_enter", "SolOn"
		.InitExitSnd SoundFX("fx_Popper",DOFContactors), SoundFX("fx_Solenoid",DOFContactors)
		.KickAngleVar = 2
	End With

	'Captive Ball
'    Set cbLeft = New cvpmCaptiveBall
'    With cbLeft
'        .InitCaptive CaptiveTrigger1, CaptiveWall, Array(Captive1, Captive1a), 355
'        .NailedBalls = 1
'        .ForceTrans = .9
'        .MinForce = 3.5
'        .CreateEvents "cbLeft"
'        .Start
'    End With
    Captive1.CreateBall
	Captive1.kick 180, 4
	Captive1.enabled = 0

'    Set cbRight = New cvpmCaptiveBall
'    With cbRight
'        .InitCaptive CaptiveTrigger2, CaptiveWall2, Array(Captive2, Captive2a), 0
'        .NailedBalls = 1
'        .ForceTrans = .9
'        .MinForce = 3.5
'        .CreateEvents "cbRight"
'        .Start
'    End With
    Captive2.CreateBall
	Captive2.kick 180, 4
	Captive2.enabled = 0

	'DropTargets
	     Set dt5Bank = new cvpmdroptarget
     With dt5Bank
         .initdrop array(DT17, DT27, DT37, DT47, DT57), array(17, 27, 37, 47, 57)		'set wall and switch arrays
         .InitSnd SoundFX("fx_droptarget",DOFContactors), SoundFX("fx_resetdrop",DOFContactors)
     End With

	     Set dt3Bank = new cvpmdroptarget
     With dt3Bank
         .initdrop array (Array(DT15, DT15a), Array(DT25, DT25a), Array(DT35, DT35a)), array(15, 25, 35)		'set wall and switch arrays
         .InitSnd SoundFX("fx_droptarget",DOFContactors), SoundFX("fx_resetdrop",DOFContactors)
     End With

	GIon
	BIP = 0

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1


	If nightday<5 then
		For each xx in GI:xx.intensityscale = 1:Next
		For each xx in controlledgi:xx.intensityscale = 1:Next
		For each xx in ambientgi:xx.intensityscale = 2: Next
		For each xx in controlledambientgi:xx.intensityscale = 2:Next
		For each xx in inserts:xx.intensityscale = 1.2: Next
	End If
	If nightday>5 then
		For each xx in GI:xx.intensityscale = 1:Next
		For each xx in controlledgi:xx.intensityscale = 1:Next
		For each xx in ambientgi:xx.intensityscale = 1: Next
		For each xx in controlledambientgi:xx.intensityscale = 1:Next
		For each xx in inserts:xx.intensityscale = 1: Next
	End If
	If nightday>10 then
		For each xx in GI:xx.intensityscale = .8:Next
		For each xx in controlledgi:xx.intensityscale = .8:Next
		For each xx in ambientgi:xx.intensityscale = .9: Next
		For each xx in controlledambientgi:xx.intensityscale = .9:Next
	End If
	If nightday>40 then
		For each xx in GI:xx.intensityscale = .5:Next
		For each xx in controlledgi:xx.intensityscale = .5:Next
		For each xx in ambientgi:xx.intensityscale = .7: Next
		For each xx in controlledambientgi:xx.intensityscale = .7:Next
		For each xx in inserts:xx.intensityscale = .8: Next
	End If
End Sub

'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
    If keycode = LeftFlipperKey Then
		 Controller.Switch(4) = 1
	End If
	If keycode = RightFlipperkey Then
		 Controller.Switch(5) = 1
	End If
    If keycode = PlungerKey Then Plunger.Pullback
    If keycode = LeftTiltKey Then vpmNudge.DoNudge 90, 3:PlaySound SoundFX("fx_nudge",0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then vpmNudge.DoNudge 270, 3:PlaySound SoundFX("fx_nudge",0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then vpmNudge.DoNudge 0, 4:PlaySound SoundFX("fx_nudge",0), 0, 1, 0, 0.25
	If keycode = StartGameKey Then Controller.Switch(3) = 1
    If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If keycode = LeftFlipperKey Then
		 Controller.Switch(6) = 0
		 Controller.Switch(4) = 0
	End If
	If keycode = RightFlipperkey Then
		 Controller.Switch(7) = 0
		 Controller.Switch(5) = 0
	End If
    If keycode = PlungerKey Then Plunger.Fire:PlaySoundAtVol SoundFX("SolOn",DOFContactors), Plunger, 1
	If keycode = StartGameKey Then Controller.Switch(3) = 0
    If vpmKeyUp(keycode) Then Exit Sub
End Sub

'***********
' Solenoids
' the commented solenoids are not in used in this script
'***********
	'SolCallback(1)  = "vpmSolSound ""Jet"","    	'pop bumper
	'SolCallback(2)  = "vpmSolSound ""rSling"","	'rightslingshot
	SolCallback(3)  = "dt5Drop"						'5 bank reset
	SolCallback(4)  = "dt3Drop"						'3 Bank Reset
	SolCallback(5)  = "bsKickerHole.SolOut"			'Hole Kicker
	'SolCallback(6)  = Unused
	SolCallback(7)  = "SetLamp 170,"				'Bottom Left Flasher
	SolCallback(8)  = "SetLamp 184,"				'Middle Left Flasher
	SolCallback(9)  = "SetLamp 181,"				'Top Left Flasher
	SolCallback(10) = "SetLamp 182,"				'Top Center Flasher
	SolCallback(11) = "SetLamp 179,"				'Captive Ball Flasher
	SolCallback(12) = "SetLamp 183,"				'Top Right Flasher
	SolCallback(13) = "SetLamp 178,"				'Middle Right Flasher
	SolCallback(14) = "SetLamp 180,"				'5 Bank Flasher
	SolCallback(15) = "SetLamp 171,"				'Shooter Lane Flasher
	SolCallBack(16) = "SetLamp 172,"				'Bottom Right Flasher
	SolCallback(17) = "SolBoxerLeft" 				'Left Boxer
	SolCallback(18) = "SolBoxerRight" 				'Right Boxer
	'SolCallback(19) = "SetLamp 174," 				'Backbox Flasher
	'SolCallback(20) = "SetLamp 175," 				'Backbox Flasher
	'SolCallback(21) = "SetLamp 175," 				'Backbox Flasher
	'SolCallback(22) = "SetLamp 175," 				'Backbox Flasher
	'SolCallback(23) = "SetLamp 175,"				'Backbox Flasher
	'SolCallback(24) = "SetLamp 175,"				'Backbox Flasher
	SolCallback(25) = "vpmSolSound ""boxing_bell"","		'Boxing Bell
	SolCallback(26) = "vpmSolSound ""SolOn"","		'GI Relay A
	SolCallback(27) = "vpmSolSound ""SolOn"","
	SolCallback(28) = "bsTrough.SolOut"				'Ball Release
	SolCallback(29) = "bsTrough.SolIn" 				'Outhole
	SolCallback(30) = "vpmSolSound SoundFX(""fx_knocker"",DOFKnocker),"		'Knocker
	'SolCallBack(31) = "vpmSolSound ""SolOn"","
	'SolCallback(32) = "vpmNudge.SolGameOn"
	'SolCallback(sLLFlipper)="vpmSolFlipper LeftFlipper,Nothing,"
	'SolCallback(sLRFlipper)="vpmSolFlipper RightFlipper,Flipper1,"

    Sub dt3Drop(Enabled)
		If Enabled Then
		delaytimer.enabled = 1
		End if
    End Sub
	Sub delaytimer_timer
		dt3Bank.DropSol_On
		DT1525.IsDropped = 0
		DT2535.IsDropped = 0
		delaytimer.enabled = 0
	End Sub
    Sub dt5Drop(Enabled)
		If Enabled Then
			dt5Bank.DropSol_On
			DT1727.IsDropped = 0
			DT2737.IsDropped = 0
			DT3747.IsDropped = 0
			DT4757.IsDropped = 0
		End if
    End Sub
	Sub SolBoxerLeft(enabled)
		if enabled then
			Boxers.setvalue(2)
		else
			Boxers.setvalue(0)
		end if
	end sub

	Sub SolBoxerRight(enabled)
		if enabled then
			Boxers.setvalue(1)
		else
			Boxers.setvalue(0)
		end if
	end sub
'***************
'  Slingshots
'***************
'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub LeftSlingShot_Slingshot
	vpmTimer.PulseSw(11)
    PlaySoundAtVol SoundFX("left_slingshot",DOFContactors), sling2, 1
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

Sub RubberTemp7_Hit():RubberTemp11.visible = 0:RubberTemp10.visible = 1:me.TimerEnabled = 1:End Sub
Sub RubberTemp7_Timer:RubberTemp11.visible = 1:RubberTemp10.visible = 0:me.TimerEnabled = 0:End Sub
'***************
'   Bumpers
'***************

Sub Bumper1_Hit:vpmTimer.PulseSw 10:PlaySoundAtVol SoundFX("fx_bumper",DOFContactors), ActiveBall, 1:End Sub


'*********************
' Switches & Rollovers
'*********************

Sub sw20a_Hit:Controller.Switch(20) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1 :End Sub
Sub sw20a_unHit:Controller.Switch(20) = 0:End Sub

Sub sw21_Hit:Controller.Switch(21) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1 :End Sub
Sub sw21_unHit:Controller.Switch(21) = 0:End Sub

Sub sw30_Hit:Controller.Switch(30) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1 :End Sub
Sub sw30_unHit:Controller.Switch(30) = 0:End Sub

Sub sw31_Hit:Controller.Switch(31) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1 :End Sub
Sub sw31_unHit:Controller.Switch(31) = 0:End Sub

Sub sw32_Hit:Controller.Switch(32) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1 :End Sub
Sub sw32_unHit:Controller.Switch(32) = 0:End Sub

Sub sw40_Hit:Controller.Switch(40) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1 :End Sub
Sub sw40_unHit:Controller.Switch(40) = 0:End Sub

Sub sw41_Hit:Controller.Switch(41) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1 :End Sub
Sub sw41_unHit:Controller.Switch(41) = 0:End Sub

Sub sw42_Hit:Controller.Switch(42) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1 :End Sub
Sub sw42_unHit:Controller.Switch(42) = 0:End Sub

Sub sw52_Hit:Controller.Switch(52) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1 :End Sub
Sub sw52_unHit:Controller.Switch(52) = 0:End Sub

Sub sw50_Hit:Controller.Switch(50) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1 :End Sub
Sub sw50_unHit:Controller.Switch(50) = 0:End Sub

'****************************
' Drain holes, vuks & saucers
'****************************

Sub Drain_Hit():BIP=BIP-1:PlaysoundAtVol "fx_drain", Drain, 1: Me.TimerEnabled = 1:End Sub
Sub Drain_Timer:bsTrough.AddBall Me:Me.TimerEnabled = 0:End Sub
Sub Kicker1_Hit():PlaySoundAtVol "fx_kicker_enter", Kicker1, 1 :bsKickerHole.Addball me:End Sub

'***************
'  Targets
'***************

Sub sw12_Hit
    vpmTimer.PulseSw 12
    PlaySound SoundFX("fx_target",DOFContactors), 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub sw13_Hit
    vpmTimer.PulseSw 13
    PlaySound SoundFX("fx_target",DOFContactors), 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub sw14_Hit
    vpmTimer.PulseSw 14
    PlaySound SoundFX("fx_target",DOFContactors), 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub sw22_Hit
    vpmTimer.PulseSw 22
    PlaySound SoundFX("fx_target",DOFContactors), 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub sw23_Hit
    vpmTimer.PulseSw 23
    PlaySound SoundFX("fx_target",DOFContactors), 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub sw24_Hit
    vpmTimer.PulseSw 24
    PlaySound SoundFX("fx_target",DOFContactors), 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub sw33_Hit
    vpmTimer.PulseSw 33
    PlaySound SoundFX("fx_target",DOFContactors), 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub sw34_Hit
    vpmTimer.PulseSw 34
    PlaySound SoundFX("fx_target",DOFContactors), 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub sw43_Hit
    vpmTimer.PulseSw 43
    PlaySound SoundFX("fx_target",DOFContactors), 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub sw44_Hit
    vpmTimer.PulseSw 44
    PlaySound SoundFX("fx_target",DOFContactors), 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub sw45_Hit
    vpmTimer.PulseSw 45
    PlaySound SoundFX("fx_target",DOFContactors), 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub sw53_Hit
    vpmTimer.PulseSw 53
    PlaySound SoundFX("fx_target",DOFContactors), 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub sw54_Hit
    vpmTimer.PulseSw 54
    PlaySound SoundFX("fx_target",DOFContactors), 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub sw55_Hit
    vpmTimer.PulseSw 55
    PlaySound SoundFX("fx_target",DOFContactors), 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

'***************
' Droptargets
'***************

Sub DT15_Hit() dt3Bank.Hit 1:DT1525.IsDropped = 1:End Sub
Sub DT25_Hit() dt3Bank.Hit 2:DT1525.IsDropped = 1:DT2535.IsDropped = 1:End Sub
Sub DT35_Hit() dt3Bank.Hit 3:DT2535.IsDropped = 1:End Sub
Sub DT17_Hit() dt5Bank.Hit 1:DT1727.IsDropped = 1:End Sub
Sub DT27_Hit() dt5Bank.Hit 2:DT1727.IsDropped = 1:DT2737.IsDropped = 1:End Sub
Sub DT37_Hit() dt5Bank.Hit 3:DT2737.IsDropped = 1:DT3747.IsDropped = 1:End Sub
Sub DT47_Hit() dt5Bank.Hit 4:DT3747.IsDropped = 1:DT4757.IsDropped = 1:End Sub
Sub DT57_Hit() dt5Bank.Hit 5:DT4757.IsDropped = 1:End Sub
Sub DT1525_Hit() dt3bank.Hit 1:dt3bank.Hit 2:DT1525.IsDropped = 1:End Sub
Sub DT2535_Hit() dt3bank.Hit 2:dt3bank.Hit 3:DT2535.IsDropped = 1:End Sub
Sub DT1727_Hit() dt5bank.Hit 1:dt5bank.Hit 2:DT1727.IsDropped = 1:End Sub
Sub DT2737_Hit() dt5bank.Hit 2:dt5bank.Hit 3:DT2737.IsDropped = 1:End Sub
Sub DT3747_Hit() dt5bank.Hit 3:dt5bank.Hit 4:DT3747.IsDropped = 1:End Sub
Sub DT4757_Hit() dt5bank.Hit 4:dt5bank.Hit 5:DT4757.IsDropped = 1:End Sub
'************
' Spinners
'************

'Sub Spinner_Spin:vpmTimer.PulseSw(25):PlaySound "fx_spinner", 0, 1, 0.05, 0.05:End Sub


'********************
'    Flippers
'********************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFX("fx_flipperup",DOFContactors), LeftFlipper, 1
        LeftFlipper.RotateToEnd
		LeftFlipper1.RotateToEnd
    Else
        PlaySoundAtVol SoundFX("fx_flipperdown",DOFContactors), LeftFlipper, 1
        LeftFlipper.RotateToStart
		LeftFlipper1.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFX("fx_flipperup",DOFContactors), RightFlipper, 1
        RightFlipper.RotateToEnd
		UpperFlipper.RotateToEnd
		RightFlipper1.RotateToEnd
    Else
        PlaySoundAtVol SoundFX("fx_flipperdown",DOFContactors), RightFlipper, 1
        RightFlipper.RotateToStart
		UpperFlipper.RotateToStart
		RightFlipper1.RotateToStart
    End If
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySoundAtBallVol "fx_rubber_flipper", parm / 10
End Sub

Sub RightFlipper_Collide(parm)
    PlaySoundAtBallVol "fx_rubber_flipper", parm / 10
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

Sub UpdateLamps
	NFadeL 0, L0
	NFadeL 1, L1
	NFadeL 2, L2
	NFadeL 3, L3
	NFadeL 4, L4
	NFadeL 5, L5
	NFadeL 6, L6
	NFadeL 7, L7
	NFadeL 10, L10
	NFadeL 11, L11
	NFadeL 13, L13
	NFadeL 14, L14
	NFadeL 20, L20
	NFadeL 21, L21
	NFadeL 22, L22
	NFadeL 23, L23
	NFadeL 24, L24
	NFadeL 25, L25
	NFadeL 26, L26
	NFadeL 27, L27
	NFadeL 30, L30
	NFadeL 31, L31
	NFadeL 32, L32
	NFadeL 33, L33
	NFadeL 34, L34
	NFadeLm 35, L35a
	NFadeLm 35, L35bulb
	NFadeL 35, L35
	NFadeLm 36, L36a
	NFadeLm 36, L36b
	NFadeLm 36, L36bulb
	NFadeL 36, L36
	NFadeL 37, L37
	NFadeL 40, L40
	NFadeL 41, L41
	NFadeL 42, L42
	NFadeL 46, L46
	NFadeL 47, L47
	NFadeL 50, L50
	NFadeL 51, L51
	NFadeL 52, L52
	NFadeL 53, L53
	NFadeL 54, L54
	NFadeL 55, L55
	NFadeL 56, L56
	NFadeL 57, L57
	NFadeL 60, L60
	NFadeL 63, L63
	NFadeL 64, L64
	NFadeL 65, L65
	NFadeL 66, L66
	NFadeL 67, L67
	NFadeLm 70, L70a
	NFadeLm 70, L70bulb
	NFadeL 70, L70
	NFadeLm 71, L71a
	NFadeLm 71, L71bulb
	NFadeL 71, L71
	NFadeLm 72, L72a
	NFadeLm 72, L72bulb
	NFadeL 72, L72
	NFadeLm 73, L73a
	NFadeLm 73, L73bulb
	NFadeL 73, L73
	NFadeLm 74, L74a
	NFadeLm 74, L74b
	NFadeLm 74, L74bulb
	NFadeL 74, L74
	NFadeLm 75, L75a
	NFadeLm 75, L75b
	NFadeLm 75, L75bulb
	NFadeL 75, L75
	NFadeLm 76, L76a
	NFadeLm 76, L76b
	NFadeL 76, L76
	NFadeLm 80, L80a
	NFadeLm 80, L80b
	NFadeL 80, L80
	NFadeLm 81, L81a
	NFadeLm 81, L81b
	NFadeL 81, L81
	NFadeLm 82, L82a
'	NFadeLm 82, L82b
	NFadeL 82, L82
	NFadeLm 83, L83a
'	NFadeLm 83, L83b
	NFadeL 83, L83
	NFadeL 84, L84
	NFadeL 85, L85

	NFadeLm 170, L170b
	NFadeLm 170, L170a
	NFadeL 170, L170
	NFadeLm 171, L171b
	NFadeLm 171, L171a
	NFadeL 171,	 L171
	NFadeLm 172, L172b
	NFadeLm 172, L172a
	NFadeL 172, L172
	NFadeLm 178, L178b
	NFadeLm 178, L178a
	NFadeL 178, L178
	NFadeLm 179, L179e
	NFadeLm 179, L179d
	NFadeLm 179, L179c
	NFadeLm 179, L179b
	NFadeLm 179, L179a
	NFadeL 179, L179
	NFadeLm 180, L180b
	NFadeLm 180, L180a
	NFadeL 180, L180
	NFadeLm 181, L181b
	NFadeLm 181, L181a
	NFadeL 181, L181
	NFadeLm 182, L182b
	NFadeLm 182, L182a
	NFadeL 182, L182
	NFadeLm 183, L183b
	NFadeLm 183, L183a
	NFadeL 183, L183
	NFadeLm 184, L184b
	NFadeLm 184, L184a
	NFadeL 184, L184
    'Flashers

End Sub

' div lamp subs

Sub InitLamps()
    Dim x
    For x = 0 to 200
        LampState(x) = 0        ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4      ' used to track the fading state
        FlashSpeedUp(x) = 0.4   ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.2 ' slower speed when turning off the flasher
        FlashMax(x) = 1         ' the maximum value when on, usually 1
        FlashMin(x) = 0         ' the minimum value when off, usually 0
        FlashLevel(x) = 0       ' the intensity of the flashers, usually from 0 to 1
    Next
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
        Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1             'wait
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

' Desktop Objects: Reels & texts (you may also use lights on the desktop)

' Reels

Sub FadeR(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.SetValue 1:FadingLevel(nr) = 6                   'fading to off...
        Case 5:object.SetValue 0:FadingLevel(nr) = 1                   'ON
        Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1              'wait
        Case 9:object.SetValue 2:FadingLevel(nr) = FadingLevel(nr) + 1 'fading...
        Case 10, 11, 12:FadingLevel(nr) = FadingLevel(nr) + 1          'wait
        Case 13:object.SetValue 3:FadingLevel(nr) = 0                  'Off
    End Select
End Sub

Sub FadeRm(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.SetValue 1
        Case 5:object.SetValue 0
        Case 9:object.SetValue 2
        Case 3:object.SetValue 3
    End Select
End Sub

'Texts

Sub NFadeT(nr, object, message)
    Select Case FadingLevel(nr)
        Case 4:object.Text = "":FadingLevel(nr) = 0
        Case 5:object.Text = message:FadingLevel(nr) = 1
    End Select
End Sub

Sub NFadeTm(nr, object, b)
    Select Case FadingLevel(nr)
        Case 4:object.Text = ""
        Case 5:object.Text = message
    End Select
End Sub

' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX and Rothbauerw
' PlaySound sound, 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
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

Sub PlaySoundAtVol(sound, tableobj, Volume)
  PlaySound sound, 1, Volume, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'Set all as per ball position & speed, but Vol Multiplier may be used eg; PlaySoundAtBallVol "sound",3

Sub PlaySoundAtBallVol(sound, VolMult)
  PlaySound sound, 0, Vol(ActiveBall) * VolMult, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub

'Set position as bumperX and Vol manually.

Sub PlaySoundAtBumperVol(sound, tableobj, Vol)
  PlaySound sound, 1, Vol, AudioPan(tableobj), 0,0,1, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBOTBallZ(sound, BOT)
    PlaySound sound, 0, ABS(BOT.velz)/17, Pan(BOT), 0, Pitch(BOT), 1, 0, AudioFade(BOT)
End Sub

' play a looping sound at a location with volume
Sub PlayLoopSoundAtVol(sound, tableobj, Vol)
	PlaySound sound, -1, Vol, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

Function RndNum(min, max)
    RndNum = Int(Rnd() * (max-min + 1) ) + min ' Sets a random number between min and max
End Function

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = tableobj.y * 2 / table1.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = tableobj.x * 2 / table1.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = ball.x * 2 / table1.width-1
  If tmp > 0 Then
    Pan = Csng(tmp ^10)
  Else
    Pan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
  Vol = Csng(BallVel(ball) ^2 / VolDiv)
End Function

Function BallRollVol(ball) ' Calculates the Volume of the sound based on the ball speed
   BallRollVol = Csng(BallVel(ball) ^2 / (80000 - (79900 * Log(RollVol) / Log(100))))
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
  Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
  BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

Function BallVelZ(ball) 'Calculates the ball speed in the -Z
    BallVelZ = INT((ball.VelZ) * -1 )
End Function

Function VolZ(ball) ' Calculates the Volume of the sound based on the ball speed in the Z
    VolZ = Csng(BallVelZ(ball) ^2 / 200)*1.2
End Function

'*** Determines if a Points (px,py) is inside a 4 point polygon A-D in Clockwise/CCW order

Function InRect(px,py,ax,ay,bx,by,cx,cy,dx,dy)
  Dim AB, BC, CD, DA
  AB = (bx*py) - (by*px) - (ax*py) + (ay*px) + (ax*by) - (ay*bx)
  BC = (cx*py) - (cy*px) - (bx*py) + (by*px) + (bx*cy) - (by*cx)
  CD = (dx*py) - (dy*px) - (cx*py) + (cy*px) + (cx*dy) - (cy*dx)
  DA = (ax*py) - (ay*px) - (dx*py) + (dy*px) + (dx*ay) - (dy*ax)

  If (AB <= 0 AND BC <=0 AND CD <= 0 AND DA <= 0) Or (AB >= 0 AND BC >=0 AND CD >= 0 AND DA >= 0) Then
    InRect = True
  Else
    InRect = False
  End If
End Function


'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 17 ' total number of balls
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
    For b = 0 to UBound(BOT)
      If BallVel(BOT(b) ) > 1 Then
        rolling(b) = True
        if BOT(b).z < 30 Then ' Ball on playfield
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), AudioPan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        Else ' Ball on raised ramp
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*.5, AudioPan(BOT(b) ), 0, Pitch(BOT(b) )+50000, 1, 0, AudioFade(BOT(b) )
        End If
      Else
        If rolling(b) = True Then
          StopSound("fx_ballrolling" & b)
          rolling(b) = False
        End If
      End If
    If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
      PlaySoundAtBOTBallZ "fx_ball_drop" & b, BOT(b)
    End If
    Next
End Sub

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / (VolDiv/VolCol), AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

'************************************
' What you need to add to your table
'************************************

' a timer called RollingTimer. With a fast interval, like 10
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

' The collision is built in VP.
' You only need to add a Sub OnBallBallCollision(ball1, ball2, velocity) and when two balls collide they
' will call this routine. What you add in the sub is up to you. As an example is a simple Playsound with volume and paning
' depending of the speed of the collision.



'***Sounds***

Sub Pins_Hit (idx)
	PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Wood_Hit (idx)
	PlaySound "fx_woodhit", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
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

Sub Gate3_Hit ()
	PlaySoundAtVol "gate4", ActiveBall, 1
End Sub

Sub Gate2_Hit ()
	PlaySoundAtVol "gate", ActiveBall, 1
End Sub

Sub Gate1_Hit ()
	PlaySoundAtVol "gate", ActiveBall, 1
End Sub

Sub physicalrubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub physicalposts_Hit(idx)
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
Sub LeftFlipper1_Collide(parm)
 	RandomSoundFlipper()
End Sub
Sub RightFlipper_Collide(parm)
 	RandomSoundFlipper()
End Sub
Sub RightFlipper1_Collide(parm)
 	RandomSoundFlipper()
End Sub
Sub UpperFlipper_Collide(parm)
 	RandomSoundFlipper()
End Sub
Sub RandomSoundFlipper()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub

'*****GI Lights On
dim xx, BIP
Sub GIon()
	For each xx in GI:xx.State = 1: Next
	For each xx in ambientgi:xx.State = 1: Next
	For each xx in gibulbs:xx.State = 1: Next
'	table1.ColorGradeImage = "ColorGradeLUT256x16_ConSat"
End Sub
Sub GIoff()
	For each xx in GI:xx.State = 0: Next
	For each xx in inserts:xx.intensityscale = 2: Next
	For each xx in ambientgi:xx.State = 0: Next
	For each xx in gibulbs:xx.State = 0: Next
'	table1.ColorGradeImage = "ColorGradeLUT256x16_ConSat3"
End Sub
Sub Trigger1_Hit():BIP=BIP+1:GIon:End Sub
Sub boxingring_Hit()
	If BIP>1 Then
		GIon
	Else
		GIoff
	End If
End Sub
Sub kickertrigger_Hit()
	If BIP>1 Then
		GIon
	Else
		GIoff
	End If
End Sub
Sub kickertrigger_UnHit():Me.TimerEnabled = 1:End Sub
Sub kickertrigger_Timer:GIon:Me.TimerEnabled = 0:End Sub
'Loop GI sequence
Dim loops, loopa, loopb
Sub loopstop_Hit():loops = 0:loopa= 0:loopb = 0:End Sub
Sub Loop1_Hit():loopa = 1:End Sub
Sub Loop2_Hit()
	If loops=0 Then
		If loopa = 1 and loopb = 1 Then
			loops=1
		End If
		If loopa=1 and loopb = 0 Then
			loopb = 1
		End If
	Else
		loops=loops+1
	End If
End Sub
Sub Looptest_Hit()
	If loops=1 Then
		GIoff
	Else
		If loopa = 1 and loopb = 0 Then
			loopa = 0
		End If
	End If
End Sub

'Flipper Prims
Sub FlipperTimer_Timer
	Primitive147.roty = LeftFlipper.currentangle - 120
	Primitive149.roty = UpperFlipper.currentangle - 215
	Primitive148.roty = RightFlipper.currentangle + 120
End Sub




'
'VpinMame Broken Display from TAB
Dim Digits(47)
Digits(0)=Array(J1,J2,J3,J4,J5,J6,J7,J8,J9,J10,J11,J12,J13,J14,J15,J16)
Digits(1)=Array(J17,J18,J19,J20,J21,J22,J23,J24,J25,J26,J27,J28,J29,J30,J31,J32)
Digits(2)=Array(J33,J34,J35,J36,J37,J38,J39,J40,J41,J42,J43,J44,J45,J46,J47,J48)
Digits(3)=Array(J49,J50,J51,J52,J53,J54,J55,J56,J57,J58,J59,J60,J61,J62,J63,J64)
Digits(4)=Array(J65,J66,J67,J68,J69,J70,J71,J72,J73,J74,J75,J76,J77,J78,J79,J80)
Digits(5)=Array(J81,J82,J83,J84,J85,J86,J87,J88,J89,J90,J91,J92,J93,J94,J95,J96)
Digits(6)=Array(J97,J98,J99,J100,J101,J102,J103,J104,J105,J106,J107,J108,J109,J110,J111,J112)
Digits(7)=Array(J113,J114,J115,J116,J117,J118,J119,J120,J121,J122,J123,J124,J125,J126,J127,J128)
Digits(8)=Array(J129,J130,J131,J132,J133,J134,J135,J136,J137,J138,J139,J140,J141,J142,J143,J144)
Digits(9)=Array(J145,J146,J147,J148,J149,J150,J151,J152,J153,J154,J155,J156,J157,J158,J159,J160)
Digits(10)=Array(J161,J162,J163,J164,J165,J166,J167,J168,J169,J170,J171,J172,J173,J174,J175,J176)
Digits(11)=Array(J177,J178,J179,J180,J181,J182,J183,J184,J185,L186,L187,L188,L189,L190,L191,L192)
Digits(12)=Array(L193,L194,L195,L196,L197,L198,L199,L200,L201,L202,L203,L204,L205,L206,L207,L208)
Digits(13)=Array(L209,L210,L211,L212,L213,L214,L215,L216,L217,L218,L219,L220,L221,L222,L223,L224)
Digits(14)=Array(L225,L226,L227,L228,L229,L230,L231,L232,L233,L234,L235,L236,L237,L238,L239,L240)
Digits(15)=Array(L241,L242,L243,L244,L245,L246,L247,L248,L249,L250,L251,L252,L253,L254,L255,L256)
Digits(16)=Array(L257,L258,L259,L260,L261,L262,L263,L264,L265,L266,L267,L268,L269,L270,L271,L272)
Digits(17)=Array(L273,L274,L275,L276,L277,L278,L279,L280,L281,L282,L283,L284,L285,L286,L287,L288)
Digits(18)=Array(L289,L290,L291,L292,L293,L294,L295,L296,L297,L298,L299,L300,L301,L302,L303,L304)
Digits(19)=Array(L305,L306,L307,L308,L309,L310,L311,L312,L313,L314,L315,L316,L317,L318,L319,L320)
Digits(20)=Array(R1,R2,R3,R4,R5,R6,R7,R8,R9,R10,R11,R12,R13,R14,R15,R16)
Digits(21)=Array(R17,R18,R19,R20,R21,R22,R23,R24,R25,R26,R27,R28,R29,R30,R31,R32)
Digits(22)=Array(R33,R34,R35,R36,R37,R38,R39,R40,R41,R42,R43,R44,R45,R46,R47,R48)
Digits(23)=Array(R49,R50,R51,R52,R53,R54,R55,R56,R57,R58,R59,R60,R61,R62,R63,R64)
Digits(24)=Array(R65,R66,R67,R68,R69,R70,R71,R72,R73,R74,R75,R76,R77,R78,R79,R80)
Digits(25)=Array(R81,R82,R83,R84,R85,R86,R87,R88,R89,R90,R91,R92,R93,R94,R95,R96)
Digits(26)=Array(R97,R98,R99,R100,R101,R102,R103,R104,R105,R106,R107,R108,R109,R110,R111,R112)
Digits(27)=Array(R113,R114,R115,R116,R117,R118,R119,R120,R121,R122,R123,R124,R125,R126,R127,R128)
Digits(28)=Array(R129,R130,R131,R132,R133,R134,R135,R136,R137,R138,R139,R140,R141,R142,R143,R144)
Digits(29)=Array(R145,R146,R147,R148,R149,R150,R151,R152,R153,R154,R155,R156,R157,R158,R159,R160)
Digits(30)=Array(R161,R162,R163,R164,R165,R166,R167,R168,R169,R170,R171,R172,R173,R174,R175,R176)
Digits(31)=Array(R177,R178,R179,R180,R181,R182,R183,R184,R185,R186,R187,R188,R189,R190,R191,R192)
Digits(32)=Array(R193,R194,R195,R196,R197,R198,R199,R200,R201,R202,R203,R204,R205,R206,R207,R208)
Digits(33)=Array(R209,R210,R211,R212,R213,R214,R215,R216,R217,R218,R219,R220,R221,R222,R223,R224)
Digits(34)=Array(R225,R226,R227,R228,R229,R230,R231,R232,R233,R234,R235,R236,R237,R238,R239,R240)
Digits(35)=Array(R241,R242,R243,R244,R245,R246,R247,R248,R249,R250,R251,R252,R253,R254,R255,R256)
Digits(36)=Array(R257,R258,R259,R260,R261,R262,R263,R264,R265,R266,R267,R268,R269,R270,R271,R272)
Digits(37)=Array(R273,R274,R275,R276,R277,R278,R279,R280,R281,R282,R283,R284,R285,R286,R287,R288)
Digits(38)=Array(R289,R290,R291,R292,R293,R294,R295,R296,R297,R298,R299,R300,R301,R302,R303,R304)
Digits(39)=Array(R305,R306,R307,R308,R309,R310,R311,R312,R313,R314,R315,R316,R317,R318,R319,R320)
Digits(40)=Array(LED1,LED2,LED3,LED4,LED5,LED6,LED7)
Digits(41)=Array(LED8,LED9,LED10,LED11,LED12,LED13,LED14)
Digits(42)=Array(LED15,LED16,LED17,LED18,LED19,LED20,LED21)
Digits(43)=Array(LED22,LED23,LED24,LED25,LED26,LED27,LED28)
Digits(44)=Array(LED29,LED30,LED31,LED32,LED33,LED34,LED35)
Digits(45)=Array(LED36,LED37,LED38,LED39,LED40,LED41,LED42)
Digits(46)=Array(LED43,LED44,LED45,LED46,LED47,LED48,LED49)
Digits(47)=Array(LED50,LED51,LED52,LED53,LED54,LED55,LED56)

Sub DisplayTimer_Timer
	Dim ChgLED,ii,num,chg,stat,obj
	ChgLED = Controller.ChangedLEDs(&HFFFFFFFF, &HFFFFFFFF) 'hex of binary (display 111111, or first 6 digits)
	If Not IsEmpty(ChgLED) Then
		For ii = 0 To UBound(chgLED)
			num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
			For Each obj In Digits(num)
				If chg And 1 Then obj.State = stat And 1
				chg = chg\2 : stat = stat\2
			Next
		next
	end if
End Sub

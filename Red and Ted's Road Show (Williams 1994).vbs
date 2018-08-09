'*********************************************************************
'1994 Red and Ted´s RoadShow by Williams Pinball
'*********************************************************************
'*********************************************************************
'Concept by:		Pat Lawlor
'Design by:			Pat Lawlor, Dwight Sullivan, Ted Estes
'Art by:			John Youssi
'Dots/Animation:	Scott Slomiany, Eugene Geer, Adam Rhine
'Mechanics by:		John Krutsch
'Music by:			Chris Granner
'Sound by:			Chris Granner
'Software by:		Dwight Sullivan, Ted Estes
'*********************************************************************
'*********************************************************************
'recreated for Visual Pinball by Knorr and Clark Kent
'*********************************************************************
'*********************************************************************
'Special thanks to Fuzzel and Toxie
'FlasherCaps and Spotlights made by Zany and Dark
'BumperCaps made by MJR
'Special thanks to Arngrim for DOF testing
'thanks to kiwi and gigagula for testing the Table1
'*********************************************************************
'*********************************************************************
'V1.1
'DOF changes by Arngrim
'V1.0
'First Release For VP10.2

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2, was a problem - fix by DjRobX applied
' Thalamus 2018-08-09 : Improved directional sounds


Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01560000", "wpc.VBS", 3.36

' Options
' Volume devided by - lower gets higher sound

Const VolDiv = 2000

'********************
' Flares
'********************

Flares = 1			'1 = 0n, 0 = 0ff

'********************
'Standard definitions
'********************

Const cGameName = "rs_l6"
Const UseSolenoids = 2
' Thal : Added because of useSolenoids=2
Const cSingleLFlip = 0
Const cSingleRFlip = 0
Const UseLamps = 0
Const SSolenoidOn = "SolOn"
Const SSolenoidOff = "SolOff"
Const SFlipperOn = "FlipperUp"
Const SFlipperOff = "FlipperDown"
Const SCoin = "Coin4"


Dim bsTrough, StartCity, bsLockKickout, bsLock, bsRed, BullDozerMech, TedJawMech, RedJawMech, PPL, aBall, aBall77a, aBall77b, aBall77c, aBall77d,  Flares

'Set LampCallback = GetRef("UpdateMultipleLamps")
Set MotorCallback = GetRef("RealTimeUpdates")
Set GiCallback2 = GetRef("UpdateGI")
Set GiCallback = GetRef("UpdateGi2")
BSize = 25.5
BMass = 1

'************
' Table init
'************

Sub table1_Init
    vpmInit Me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
        .Games(cGameName).Settings.Value("rol") = 0
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = 0
        '.SetDisplayPosition 0, 0, GetPlayerHWnd 'uncomment this line if you don't see the DMD
        On Error Resume Next
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
        On Error Goto 0
        .Switch(22) = 1 'close coin door
        .Switch(24) = 1 'and keep it close
        PinMAMETimer.Interval = PinMAMEInterval
        PinMAMETimer.Enabled = true
        vpmNudge.TiltSwitch = 14
        vpmNudge.Sensitivity = 2
    End With

    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0, 42, 43, 44, 45, 0, 0, 0
        .InitKick BallRelease, 80, 10
        .InitExitSnd SoundFX("BallRelease", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .Balls = 4
        .IsTrough = 1
    End With

    Set StartCity = New cvpmBallStack
    With StartCity
        .InitSw 0, 78, 0, 0, 0, 0, 0, 0
        .InitKick sw78, 184, 24 '19
        .KickBalls = 2
        .InitExitSnd SoundFX("StartCity_SolOut",DOFContactors), SoundFX("solenoid",DOFContactors)
    End With

    Set bsLockKickout = new cvpmBallStack
    With bsLockKickout
        .InitSw 0, 54, 0, 0, 0, 0, 0, 0
        .InitKick Locker, 178, 0.001
        .InitExitSnd SoundFX("SafeHouseKick",DOFContactors), SoundFX("solenoid",DOFContactors)
		.KickForceVar = 0
        .KickBalls = 1
    End With

    Set bsLock = new cvpmBallStack
    With bsLock
        .InitSw 0, 52, 53, 0, 0, 0, 0, 0
        .KickBalls = 1
    End With

 ' Red Kickout
    Set bsRed = new cvpmBallStack
    With bsRed
        .InitSw 0, 25, 0, 0, 0, 0, 0, 0
        .InitKick sw25, 225, 6
        .KickForceVar = 3
        .KickBalls = 1
        .InitExitSnd SoundFX("SafeHouseKick",DOFContactors), SoundFX("solenoid",DOFContactors)
    End With

    Set BulldozerMech = New cvpmMech
    With BulldozerMech
        .MType = vpmMechOneSol + vpmMechReverse + vpmMechNonLinear + vpmMechFast
        .Sol1 = 23
        .Length = DozerSpeed
        .Steps = 11
        .AddSw 12, 0, 0
        .AddSw 15, 10, 11
        .Callback = GetRef("UpdateDozer")
        .Start
    End With

	Set TedJawMech = New cvpmMech
	With TedJawMech
		.MType = vpmMechOneDirSol + vpmMechStopEnd + vpmMechNonLinear + vpmMechFast
		.Sol1 = 20
		.Sol2 = 19
		.length = jawspeed
		.steps = 18
		.callback = getRef("UpdateJawTed")
		.start
	End With

	Set RedJawMech = New cvpmMech
	With RedJawMech
		.MType = vpmMechOneDirSol + vpmMechStopEnd + vpmMechNonLinear + vpmMechFast
		.Sol1 = 17
		.Sol2 = 18
		.length = jawspeed
		.steps = 18
		.callback = getRef("UpdateJawRed")
		.start
	End With

	diverterleftramp1.isDropped = 1

If Flares = 1 Then
		BlastZoneFlasher.visible = 1
		LeftFlasher.visible = 1
		RightRampFlasher.visible = 1
		LeftRampFlasher.visible = 1
	End if


    If table1.ShowDT = False then
        LeftSideRail1.WidthTop = 0
        LeftSideRail1.WidthBottom = 0
        RightSideRail1.WidthTop = 0
        RightSideRail1.WidthBottom = 0
		Lockbar.Visible = False
        Korpus.Size_Y = 3
		Korpus.Z = 60
		GiFlasher3.Height = 200
		GiFlasher4.Height = 200
		GiFlasher.Height = 220
		GiFlasher1.Height = 220
		GiFlasher2.Height = 220
		GiFlasher5.Height = 220
		GiFlasher6.Height = 220
		GiFlasher7.Height = 220
		GiFlasher8.Height = 220
'		f85d.Height = 330
'		f43b.Height = 270
		Else
		GiFlasher3.Height = 100
		GiFlasher4.Height = 100
		GiFlasher.Height = 120
		GiFlasher1.Height = 120
		GiFlasher2.Height = 120
		GiFlasher5.Height = 120
		GiFlasher6.Height = 120
		GiFlasher7.Height = 120
		GiFlasher8.Height = 120
    End if
End Sub

Sub table1_Exit
	Controller.Stop
End Sub

'*********
'Solenoids
'*********

SolCallback(1) = "SolRelease"
SolCallback(2) = "LowerLeftDiverter"
SolCallback(3) = "LockupPin"
SolCallback(4) = "UpperLeftDiverter"
SolCallback(5) = "UpperRightDiverter"
SolCallback(6) = "StartCity.SolOut"
SolCallback(7) = "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallBack(8) = "bsLockKickOut.SolOut"
SolCallBack(9) = "TedEyesLeft"
SolCallBack(10) = "TedLidsDown"
SolCallBack(11) = "TedLidsUp"
SolCallBack(12) = "TedEyesRight"
SolCallBack(13) = "RedLidsDown"
SolCallBack(14) = "RedEyesLeft"
SolCallBack(15) = "RedLidsUp"
SolCallBack(16) = "RedEyesRight"
SolCallBack(17) = "RedMotorOn"
'SolCallBack(18) = "RedMotorDir"
'SolCallBack(19) = "TedMotorDir"
SolCallBack(20) = "TedMotorOn"
SolCallBack(23) = "BullDozerMotor"
SolCallBack(24) = "bsRed.SolOut"
'SolCallBack(28) = "ShakerMotor"
'
SolCallBack(51) = "SetLamp 137," ' Little Flipper Flasher
SolCallBack(52) = "SetLamp 138," ' Left Ramp Flasher
SolCallBack(53) = "SetLamp 139," ' Back White Flashers
SolCallBack(54) = "SetLamp 140," ' Back Yellow Flashers
SolCallBack(55) = "SetLamp 141," ' Back Red Flashers
SolCallBack(56) = "SetLamp 142," ' Blasting Zone Flashers
SolCallBack(57) = "SetLamp 143," ' Right Ramp Flasher
SolCallBack(58) = "SetLamp 144," ' Jets Flasher


Sub SolRelease(Enabled)
    If Enabled And bsTrough.Balls> 0 Then
        vpmTimer.PulseSw 41
        bsTrough.ExitSol_On
    End If
End Sub

Sub Drain_Hit
    PlaySoundAt "Balltruhe", Drain
    bsTrough.AddBall Me
End Sub

Sub bsLockAddBall()
	bsLock.AddBall 1
End Sub

Sub Locker_Hit()
	PlaySoundAt "Lock_Hit2", Locker
	bsLock.AddBall Me
End Sub

Sub LockUpPin(enabled)
    If enabled then
        If bsLock.Balls> 0 Then
            bsLock.SolOut 1
            PlaySound "SaveHouseKick" ' TODO : improve positional
            bsLockKickout.AddBall 1
        Else
            PlaySound "LockUpPin"
        End If
    End If
End Sub

'TedsMouth
Sub sw11_Hit()
	vpmTimer.PulseSw 11
	PlaySound "Ted_Hit2"
	vpmTimer.AddTimer 1200, "bsLockAddBall'"
	Me.DestroyBall
End Sub

Sub sw48_Hit()
	vpmTimer.PulseSw 48
End Sub

'Dozer
Sub sw47_Hit()
	vpmTimer.PulseSw 47
End Sub

Sub dozerwall_Hit()
	shakedozer
	PlaySound "flip_hit_3"
End Sub

Sub sw37a_Hit()
	vpmTimer.PulseSw 37
End Sub


'RedsMouth
Sub sw25_Hit()
	PlaySound "SafeHouseHit", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	bsRed.AddBall Me
End Sub



'BlastZone
Sub sw77a_Hit()
	vpmTimer.PulseSw 77
	Set aBall77a = ActiveBall:Me.TimerEnabled = 1
	RandomSoundBlastZone
	vpmTimer.AddTimer 1000, "StartCityAddBall'"
End Sub

Sub sw77a_Timer()
    Do While aBall77a.Z > 0
        aBall77a.Z = aBall77a.Z -5
        Exit Sub
    Loop
    Me.DestroyBall
    Me.TimerEnabled = 0
End Sub

Sub sw77b_Hit()
	vpmTimer.PulseSw 77
	Set aBall77b = ActiveBall:Me.TimerEnabled = 1
	RandomSoundBlastZone
	vpmTimer.AddTimer 1000, "StartCityAddBall'"
End Sub

Sub sw77b_Timer()
    Do While aBall77b.Z > 0
        aBall77b.Z = aBall77b.Z -5
        Exit Sub
    Loop
    Me.DestroyBall
    Me.TimerEnabled = 0
End Sub

Sub sw77c_Hit()
	vpmTimer.PulseSw 77
	Set aBall77c = ActiveBall:Me.TimerEnabled = 1
	PlaySound "BlastZone_HitR", 0, 0.5, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	vpmTimer.AddTimer 1000, "StartCityAddBall'"
End Sub

Sub sw77c_Timer()
    Do While aBall77c.Z > 0
        aBall77c.Z = aBall77c.Z -5
        Exit Sub
    Loop
    Me.DestroyBall
    Me.TimerEnabled = 0
End Sub

Sub sw77d_Hit()
	vpmTimer.PulseSw 77
	Set aBall77d = ActiveBall:Me.TimerEnabled = 1
	PlaySound "BlastZone_HitR", 0, 0.5, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	vpmTimer.AddTimer 1000, "StartCityAddBall'"
End Sub

Sub sw77d_Timer()
    Do While aBall77d.Z > 0
        aBall77d.Z = aBall77d.Z -5
        Exit Sub
    Loop
    Me.DestroyBall
    Me.TimerEnabled = 0
End Sub


Sub StartCityAddBall()
	StartCity.AddBall 1
End Sub



'diverter

Sub UpperRightDiverter(Enabled)
	If Enabled then
		diverterr1.visible = 1
		diverterr2.visible = 0
		diverterrightramp.isDropped = 1
		PlaySound "DiverterRight_Open"
	Else
		diverterr1.visible = 0
		diverterr2.visible = 1
		diverterrightramp.isDropped = 0
		PlaySound "DiverterLeft_Close"
	End if
End Sub

Sub UpperLeftDiverter(Enabled)
	If Enabled then
		diverterl1.visible = 1
		diverterl2.visible = 0
		diverterleftramp1.isDropped = 0
		diverterleftramp2.isDropped = 1
		PlaySound "DiverterRight_Open"
	Else
		diverterl1.visible = 0
		diverterl2.visible = 1
		diverterleftramp1.isDropped = 1
		diverterleftramp2.isDropped = 0
		PlaySound "DiverterLeft_Close"
	End if
End Sub

Sub LowerLeftDiverter(enabled)
	If Enabled Then
		diverter1.rotatetoend
		diverter2.rotatetoend
		PlaySound "DiverterRight_Open"
	Else
		diverter1.rotatetostart
		diverter2.rotatetostart
		PlaySound "DiverterLeft_Close"
	End If
End Sub


'*********
' Bumper
'*********

Sub LeftJetBumper_hit:vpmTimer.pulseSw 63:Playsound SoundFX("BumperLeft", DOFContactors), 0, 1, 0.15, 0.15, AudioFade(LeftJetBumper):Me.TimerEnabled = 1:End Sub
Sub LeftJetBumper_Timer:Me.Timerenabled = 0:End Sub

Sub RightJetBumper_hit:vpmTimer.pulseSw 65:Playsound SoundFX("BumperRight", DOFContactors), 0, 1, 0.15, 0.15, AudioFade(RightJetBumper):Me.TimerEnabled = 1:End Sub
Sub RightJetBumper_Timer:Me.Timerenabled = 0:End Sub

Sub TopJetBumper_hit:vpmTimer.pulseSw 64:Playsound SoundFX("BumperMiddle", DOFContactors), 0, 1, 0.15, 0.15, AudioFade(TopJetBumper):Me.TimerEnabled = 1:End Sub
Sub TopJetBumper_Timer:Me.Timerenabled = 0:End Sub


'*********
' Switches
'*********

Sub sw16_Hit:Controller.Switch(16) = 1:sw16wire.RotX = 15:PlaySoundAt "metalhit_thin",sw16:End Sub
Sub sw16_UnHit:Controller.Switch(16) = 0:sw16wire.RotX = 0:End Sub

Sub sw17_Hit:Controller.Switch(17) = 1:sw17wire.RotX = 15:PlaySoundAt "metalhit_thin",sw17:End Sub
Sub sw17_UnHit:Controller.Switch(17) = 0:sw17wire.RotX = 0:End Sub

Sub sw18_Hit:Controller.Switch(18) = 1:sw18wire.RotX = 15:PlaySoundAt "metalhit_thin",sw18:End Sub
Sub sw18_UnHit:Controller.Switch(18) = 0:sw18wire.RotX = 0:End Sub

Sub sw26_Hit:Controller.Switch(26) = 1:sw26wire.RotX = 15:PlaySoundAt "metalhit_thin",sw26:End Sub
Sub sw26_UnHit:Controller.Switch(26) = 0:sw26wire.RotX = 0:End Sub

Sub sw27_Hit:Controller.Switch(27) = 1:sw27wire.RotX = 15:PlaySoundAt "metalhit_thin",sw27:End Sub
Sub sw27_UnHit:Controller.Switch(27) = 0:sw27wire.RotX = 0:End Sub

Sub sw31_Hit:Controller.Switch(31) = 1:sw31wire.RotX = 15:PlaySoundAt "metalhit_thin",sw31:End Sub
Sub sw31_UnHit:Controller.Switch(31) = 0:sw31wire.RotX = 0:End Sub

Sub sw32_Hit:Controller.Switch(32) = 1:sw32wire.RotX = 15:PlaySoundAt "metalhit_thin",sw32:End Sub
Sub sw32_UnHit:Controller.Switch(32) = 0:sw32wire.RotX = 0:End Sub

Sub sw33_Hit:Controller.Switch(33) = 1:sw33wire.RotX = 15:PlaySoundAt "metalhit_thin",sw33:End Sub
Sub sw33_UnHit:Controller.Switch(33) = 0:sw33wire.RotX = 0:End Sub

Sub sw38_Hit:Controller.Switch(38) = 1:sw38wire.RotX = 15:PlaySound "metalhit_thin":End Sub
Sub sw38_UnHit:Controller.Switch(38) = 0:sw38wire.RotX = 0:End Sub

Sub sw46_Hit:Controller.Switch(46) = 1:sw46wire.RotX = 15:PlaySoundAt "metalhit_thin",sw46:End Sub
Sub sw46_UnHit:Controller.Switch(46) = 0:sw46wire.RotX = 0:End Sub

Sub sw55_Hit:Controller.Switch(55) = 1:sw55wire.RotX = 15:PlaySoundAt "metalhit_thin",sw55:End Sub
Sub sw55_UnHit:Controller.Switch(55) = 0:sw55wire.RotX = 0:End Sub

Sub sw58_Hit:Controller.Switch(58) = 1:PPL = True:PlungerL.MechPlunger = 1:PlungerR.MechPlunger = 0:sw58wire.RotX = 15:PlaySoundAt "metalhit_thin",sw58:StopSound "WireRamp":End Sub
Sub sw58_UnHit:Controller.Switch(58) = 0:PPL = False:PlungerL.MechPlunger = 0:PlungerR.MechPlunger = 1:sw58wire.RotX = 0:End Sub

Sub sw85_Hit:Controller.Switch(85) = 1:sw85wire.RotX = 15:PlaySoundAt "metalhit_thin",sw85:End Sub
Sub sw85_UnHit:Controller.Switch(85) = 0:sw85wire.RotX = 0:End Sub

Sub sw86_Hit:Controller.Switch(86) = 1:sw86wire.RotX = 15:PlaySoundAt "metalhit_thin",sw86:End Sub
Sub sw86_UnHit:Controller.Switch(86) = 0:sw86wire.RotX = 0:End Sub

Sub sw73_Hit:Controller.Switch(73) = 1:PlaySoundAt "metalhit_thin",sw73:vpmTimer.AddTimer 100, "BallDropSound'":StopSound "wireramp_right":End Sub
Sub sw73_UnHit:Controller.Switch(73) = 0:StopSound "wireramp":End Sub

Sub sw74_Hit:Controller.Switch(74) = 1:PlaySoundAt "metalhit_thin",sw74:vpmTimer.AddTimer 100, "WireRampHitS'":PlaySound "wireramp_right", 1, 0.3, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0:End Sub
Sub sw74_UnHit:Controller.Switch(74) = 0:End Sub

Sub sw75_Hit:Controller.Switch(75) = 1:PlaySoundAt "metalhit_thin",sw75:vpmTimer.AddTimer 100, "WireRampHitS'":End Sub
Sub sw75_UnHit:Controller.Switch(75) = 0::End Sub

Sub sw76_Hit:Controller.Switch(76) = 1:PlaySoundAt "metalhit_thin",sw76:End Sub
Sub sw76_UnHit:Controller.Switch(76) = 0:End Sub

Sub Wall26_Hit: PlaySound "rubber_hit_1", 0, 0.5, -0.2:End Sub

Sub sw51spinner_Spin():vpmTimer.PulseSw 51:PlaySoundAt "fx_spinner",sw51spinner:End Sub

Sub WireRampHitS()
	PlaySound "wireramp_stop", 0, 0.3, -0.5
End Sub

'Targets

Sub T28a_Hit:vpmTimer.PulseSw 28:PlaySound SoundFX("target",DOFTargets), 0, 1, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(T28a):End Sub
Sub T28b_Hit:vpmTimer.PulseSw 28:PlaySound SoundFX("target",DOFTargets), 0, 1, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(T28b):End Sub
Sub T28c_Hit:vpmTimer.PulseSw 28:PlaySound SoundFX("target",DOFTargets), 0, 1, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(T28c):End Sub

Sub T34a_Hit:vpmTimer.PulseSw 34:PlaySound SoundFX("target",DOFTargets), 0, 1, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(T34a):End Sub
Sub T34b_Hit:vpmTimer.PulseSw 34:PlaySound SoundFX("target",DOFTargets), 0, 1, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(T34b):End Sub
Sub T34c_Hit:vpmTimer.PulseSw 34:PlaySound SoundFX("target",DOFTargets), 0, 1, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(T34c):End Sub

Sub T35_Hit:vpmTimer.PulseSw 35:PlaySound SoundFX("target",DOFTargets), 0, 1, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(T35):End Sub
Sub T36_Hit:vpmTimer.PulseSw 36:PlaySound SoundFX("target",DOFTargets), 0, 1, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(T36):End Sub

Sub T81_Hit:vpmTimer.PulseSw 81:PlaySound SoundFX("target",DOFTargets), 0, 1, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(T81):End Sub
Sub T82_Hit:vpmTimer.PulseSw 82: T82P.RotX = T82P.RotX +5:Playsound SoundFX("target", DOFContactors), 0, 1, pan(ActiveBall), 0, Pitch(ActiveBall):Me.TimerEnabled = 1:End Sub
Sub T82_Timer:T82P.RotX = T82P.RotX -5:Me.TimerEnabled = 0:End Sub
Sub T83_Hit:vpmTimer.PulseSw 83:PlaySound SoundFX("target",DOFTargets), 0, 1, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(T83):End Sub
Sub T84_Hit:vpmTimer.PulseSw 84:PlaySound SoundFX("target",DOFTargets), 0, 1, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(T84):End Sub

'Ramps

Sub sw71_Hit:Controller.Switch(71) = 1:End Sub
Sub sw71_UnHit:Controller.Switch(71) = 0:End Sub
Sub sw72_Hit:Controller.Switch(72) = 1:WireSwitchRightRamp.RotX = 110:PlaySound "metalhit_thin":End Sub
Sub sw72_UnHit:Controller.Switch(72) = 0:WireSwitchRightRamp.RotX = 90:End Sub

Sub sw57_Hit:Controller.Switch(57) = 1:End Sub
Sub sw57_UnHit:Controller.Switch(57) = 0:End Sub
Sub sw56_Hit:Controller.Switch(56) = 1:WireSwitchLeftRamp.RotX = 110:PlaySound "metalhit_thin":End Sub
Sub sw56_UnHit:Controller.Switch(56) = 0:WireSwitchLeftRamp.RotX = 90:End Sub

'KickerAnimation

Sub sw78_Hit()
    PlaySound "StartCity_Hit2",  0, 1, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Set aBall = ActiveBall:Me.TimerEnabled = 1
    vpmTimer.AddTimer 1000, "StartCityAddBall2'"
End Sub

Sub sw78_Timer()
    Do While aBall.Z > 0
        aBall.Z = aBall.Z -5
        Exit Sub
    Loop
    Me.DestroyBall
    Me.TimerEnabled = 0
End Sub

Sub StartCityAddBall2()
	StartCity.AddBall 0
End Sub

'***************
' ToysAnimation
'***************

'Bulldozer

'Sub UpdateDozer(aNewPos,aSpeed,aLastPos)
'	TedBar.RotX = 90 +BullDozermech.position
'	TedBarPl.RotX = 90 +BullDozermech.position
'	dozerscrews.RotX = 90 +BullDozermech.position
'	If TedBar.RotX > 95 then dozerwall.isDropped = 1: Else dozerwall.isDropped = 0
'End Sub

Sub UpdateDozer(aNewPos,aSpeed,aLastPos)
	DozerRotate.Enabled = 1
End Sub


Dim DozerSpeed, DozerStep


'normal Mech
'DozerSpeed = 48
'DozerStep = 2.22*DozerSpeed
'Sub DozerRotate_Timer()
'	If bulldozermech.position +90 > TedBar.RotX then
'	TedBar.RotX = TedBar.RotX +(11/DozerStep)
'	End if
'	If bulldozermech.position +90 < TedBar.RotX then
'	TedBar.RotX = TedBar.RotX -(11/DozerStep)
'	End if
'	TedBarPl.RotX = TedBar.RotX
'	dozerscrews.RotX = TedBar.RotX
'	If TedBar.RotX > 95 then dozerwall.isDropped = 1: Else dozerwall.isDropped = 0
'End Sub

'fast Mech
DozerSpeed = 192
DozerStep = 0.4166666666667*DozerSpeed
Sub DozerRotate_Timer()
	If bulldozermech.position +90 > TedBar.RotX then
	TedBar.RotX = TedBar.RotX +(11/DozerStep)
	End if
	If bulldozermech.position +90 < TedBar.RotX then
	TedBar.RotX = TedBar.RotX -(11/DozerStep)
	End if
	TedBarPl.RotX = TedBar.RotX
	dozerscrews.RotX = TedBar.RotX
	If TedBar.RotX > 95 then dozerwall.isDropped = 1: Else dozerwall.isDropped = 0
End Sub


'Shake Dozer

Dim DozerPos

Sub ShakeDozer
    DozerPos = 3 '8
    DozerTimer.Enabled = 1
End Sub

Sub DozerTimer_Timer
    TedBar.RotAndTra4 = DozerPos
    TedBarPl.RotAndTra4 = DozerPos
    If DozerPos = 0 Then DozerTimer.Enabled = False:Exit Sub
    If DozerPos < 0 Then
        DozerPos = ABS(DozerPos) - 1
    Else
        DozerPos = - DozerPos + 1
    End If
End Sub

'Sound
Sub BullDozerMotor(Enabled)
	If enabled then
		PlaySound SoundFX("BulldozerMotor_Up",DofGear)
	Else
		StopSound SoundFX("BulldozerMotor_Up",DofGear)
	End If
End Sub


'TED
'
'Sub UpdateJawTed(aNewPos,aSpeed,aLastPos)
'	TedJaw.RotX = 73 +TedJawMech.position
'	If TedJaw.RotX < 75 then sw48.isDropped = 1: Else sw48.isDropped = 0
'End Sub

Sub UpdateJawTed(aNewPos,aSpeed,aLastPos)
'	TedJaw.RotX = 73 +TedJawMech.position
	TedJawTimer.Enabled = 1
	If aNewPos < 17 Then DOF 201, DOFPulse
End Sub


Dim JawSpeed, JawStep
JawSpeed = 90
JawStep = 0.4166666666667*jawspeed


Sub TedJawTimer_Timer()
	If TedJawMech.position +73 > TedJaw.RotX then
	TedJaw.RotX = TedJaw.RotX +(18/JawStep)
	End if
	If TedJawMech.position +73 < TedJaw.RotX then
	TedJaw.RotX = TedJaw.RotX -(18/JawStep)
	End if
	If TedJaw.RotX < 75 then sw48.isDropped = 1:Else sw48.isDropped = 0
'	If TedJawMech.position +73 = TedJaw1.RotX then PlaySound "Knocker"
End Sub

Sub TedLidsDown(Enabled)
	If Enabled then
		TedEyelid.RotX = 20
		PlaySound "TedLids_Down"
	End if
End Sub


Sub TedLidsUp(Enabled)
	If Enabled then
		 If TedEyelid.RotX = 90 then TedEyelid.RotX = 105
		PlaySound "TedLids_Up"
		Else
		TedEyelid.RotX = 90
	End if
End Sub

Sub TedEyesRight(Enabled)
	If Enabled then
		TedEyeRight.RotY = -25
		TedEyeLeft.RotY = -25
		PlaySound "TedEyes_Right"
	Else
		TedEyeRight.RotY = 0
		TedEyeLeft.RotY = 0
	End if
End Sub

Sub TedEyesLeft(Enabled)
	If Enabled then
		TedEyeRight.RotY = 25
		TedEyeLeft.RotY = 25
		PlaySound "TedEyes_Left"
	Else
		TedEyeRight.RotY = 0
		TedEyeLeft.RotY = 0
	End if
End Sub

'Sound
Sub TedMotorOn(enabled)
	If enabled then
		PlaySound SoundFX("BulldozerMotor_Down3",DofGear)
	Else
		StopSound SoundFX("BulldozerMotor_Down3",DofGear)
	End if
End Sub


'RED

Sub UpdateJawRed(aNewPos,aSpeed,aLastPos)
'	If RedJaw.RotX > 105 then sw37.isDropped = 1: Else sw37.isDropped = 0
	RedJawTimer.Enabled = 1
	If aNewPos > 0 Then DOF 201, DOFPulse
End Sub



Sub RedJawTimer_Timer()
	If RedJawMech.position +90 > RedJaw.RotX then
	RedJaw.RotX = RedJaw.RotX +(18/JawStep)
	End if
	If RedJawMech.position +90 < RedJaw.RotX then
	RedJaw.RotX = RedJaw.RotX -(18/JawStep)
	End if
	If RedJaw.RotX > 105 then sw37.isDropped = 1:Else sw37.isDropped = 0
End Sub


Sub RedLidsDown(Enabled)
	If Enabled then
		RedEyelid.RotX = 40
		PlaySound "RedLids_Down"
	End if
End Sub

Sub RedLidsUp(Enabled)
	If Enabled then
		If RedEyelid.RotX = 95 then RedEyelid.RotX = 110
		PlaySound "RedLids_Up"
		Else
		RedEyelid.RotX = 95
	End if
End Sub

Sub RedEyesRight(Enabled)
	If Enabled then
		RedEyeRight.RotY = -17
		RedEyeLeft.RotY = -17
		PlaySound "RedEyes_Right"
	Else
		RedEyeRight.RotY = 0
		RedEyeLeft.RotY = 0
	End if
End Sub

Sub RedEyesLeft(Enabled)
	If Enabled then
		RedEyeRight.RotY = 17
		RedEyeLeft.RotY = 17
		PlaySound "RedEyes_Left"
	Else
		RedEyeRight.RotY = 0
		RedEyeLeft.RotY = 0
	End if
End Sub

'Sound
Sub RedMotorOn(enabled)
	If enabled then
		PlaySound SoundFX("BulldozerMotor_Down2",DofGear)
	Else
		StopSound SoundFX("BulldozerMotor_Down2",DofGear)
	End if
End Sub

'************
' SlingShots
'************


Dim RStep, Lstep

Sub SlingShotRight_Slingshot
    PlaySound SoundFX("SlingshotRight", DOFContactors), 0, 1, 0.05, 0.05
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    SlingShotRight.TimerEnabled = 1
    vpmTimer.PulseSw 62
End Sub

Sub SlingShotRight_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:SlingShotRight.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub SlingShotLeft_Slingshot
    PlaySound SoundFX("SlingshotLeft", DOFContactors), 0, 1, -0.05, 0.05
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    SlingShotLeft.TimerEnabled = 1
    vpmTimer.PulseSw 61
End Sub

Sub SlingShotLeft_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:SlingShotLeft.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
    If KeyCode = MechanicalTilt Then
        vpmTimer.PulseSw vpmNudge.TiltSwitch
        Exit Sub
    End if
	If keycode = keyFront Then Controller.Switch(23) = True
	If keycode = PlungerKey And PPL = False Then PlungerR.Pullback:PlaysoundAt "plungerpull", PlungerR
	If keycode = PlungerKey And PPL = True Then PlungerL.Pullback:PlaysoundAt "plungerpull", PlungerL
	If keycode = LeftMagnaSave Then PlungerL.Pullback:PlaysoundAt "plungerpull", PlungerL
	If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub table1_KeyUp(ByVal Keycode)
	If keycode = keyFront Then Controller.Switch(23) = False
    If keycode = PlungerKey And PPL = False Then PlungerR.Fire: PlaySoundAt "Plunger", PlungerR
    If keycode = PlungerKey And PPL = True Then PlungerL.Fire: PlaySoundAt "Plunger", PlungerL
	If keycode = LeftMagnaSave Then PlungerL.Fire: PlaySoundAt "Plunger", PlungerL
    If vpmKeyUp(keycode) Then Exit Sub
End Sub


'**************
' Flipper Subs
'**************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"
'SolCallback(34) = "SolLFlipper3" ' Upper Left Flipper
'SolCallback(36) = "SolLFlipper2" ' Middle Left Flipper

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("FlipperUpLeft", DOFFlippers),FlipperL:FlipperL.RotateToEnd
    Else
        PlaySoundAt SoundFX("FlipperDown", DOFFlippers),FlipperL:FlipperL.RotateToStart
    End If
    SolLFlipper3 Enabled
	SolLFlipper2 Enabled
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("FlipperUpRight", DOFFlippers),FlipperR:FlipperR.RotateToEnd
    Else
        PlaySoundAt SoundFX("FlipperDown", DOFFlippers),FlipperR:FlipperR.RotateToStart
    End If
End Sub

Sub SolLFlipper3(Enabled)
    If Enabled Then
        PlaySound SoundFX("", DOFFlippers):FlipperL3.RotateToEnd
    Else
        PlaySound SoundFX("", DOFFlippers):FlipperL3.RotateToStart
    End If
End Sub

Sub SolLFlipper2(Enabled)
    If Enabled Then
        PlaySound SoundFX("", DOFContactors):FlipperL2.RotateToEnd
    Else
        PlaySound SoundFX("", DOFContactors):FlipperL2.RotateToStart
    End If
End Sub




dim returnspeed, lfstep, rfstep
returnspeed = FlipperL.return
lfstep = 1
rfstep = 1

sub FlipperL_timer()
	select case lfstep
		Case 1: FlipperL.return = returnspeed * 0.6 :lfstep = lfstep + 1
		Case 2: FlipperL.return = returnspeed * 0.7 :lfstep = lfstep + 1
		Case 3: FlipperL.return = returnspeed * 0.8 :lfstep = lfstep + 1
		Case 4: FlipperL.return = returnspeed * 0.9 :lfstep = lfstep + 1
		Case 5: FlipperL.return = returnspeed * 1 :lfstep = lfstep + 1
		Case 6: FlipperL.timerenabled = 0 : lfstep = 1
	end select
end sub

sub FlipperR_timer()
	select case rfstep
		Case 1: FlipperR.return = returnspeed * 0.6 :rfstep = rfstep + 1
		Case 2: FlipperR.return = returnspeed * 0.7 :rfstep = rfstep + 1
		Case 3: FlipperR.return = returnspeed * 0.8 :rfstep = rfstep + 1
		Case 4: FlipperR.return = returnspeed * 0.9 :rfstep = rfstep + 1
		Case 5: FlipperR.return = returnspeed * 1 :rfstep = rfstep + 1
		Case 6: FlipperR.timerenabled = 0 : rfstep = 1
	end select
end sub

sub FlipperL2_timer()
	select case rfstep
		Case 1: FlipperL2.return = returnspeed * 0.6 :rfstep = rfstep + 1
		Case 2: FlipperL2.return = returnspeed * 0.7 :rfstep = rfstep + 1
		Case 3: FlipperL2.return = returnspeed * 0.8 :rfstep = rfstep + 1
		Case 4: FlipperL2.return = returnspeed * 0.9 :rfstep = rfstep + 1
		Case 5: FlipperL2.return = returnspeed * 1 :rfstep = rfstep + 1
		Case 6: FlipperL2.timerenabled = 0 : rfstep = 1
	end select
end sub

sub FlipperL3_timer()
	select case rfstep
		Case 1: FlipperL3.return = returnspeed * 0.6 :rfstep = rfstep + 1
		Case 2: FlipperL3.return = returnspeed * 0.7 :rfstep = rfstep + 1
		Case 3: FlipperL3.return = returnspeed * 0.8 :rfstep = rfstep + 1
		Case 4: FlipperL3.return = returnspeed * 0.9 :rfstep = rfstep + 1
		Case 5: FlipperL3.return = returnspeed * 1 :rfstep = rfstep + 1
		Case 6: FlipperL3.timerenabled = 0 : rfstep = 1
	end select
end sub



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
        FlashSpeedDown(x) = 0.1 ' slower speed when turning off the flasher
        FlashMax(x) = 1          ' the maximum value when on, usually 1
        FlashMin(x) = 0          ' the minimum value when off, usually 0
        FlashLevel(x) = 0        ' the intensity of the flashers, usually from 0 to 1
    Next
End Sub

 Sub UpdateLamps
	'Inserts
	NFadeL 11,  l11
	NFadeL 12,  l12
	NFadeL 13,  l13
	NFadeL 14,  l14
	NFadeL 15,  l15
	NFadeL 16,  l16
	NFadeL 17,  l17
	NFadeL 18,  l18

	NFadeL 21,  l21
	NFadeL 22,  l22
	NFadeL 23,  l23
    NFadeL 24,  l24
    NFadeL 25,  l25
    NFadeL 26,  l26
    NFadeL 27,  l27
    NFadeL 28,  l28

    NFadeL 31,  l31
    NFadeL 32,  l32
    NFadeL 33,  l33
    NFadeL 34,  l34
    NFadeL 35,  l35
    NFadeL 36,  l36
    NFadeL 37,  l37
    NFadeL 38,  l38

    NFadeL 41,  l41
    NFadeL 42,  l42
'    NFadeL 43,  l43
    NFadeL 44,  l44
    NFadeL 45,  l45
    NFadeL 46,  l46
    NFadeL 47,  l47
    NFadeL 48,  l48

    NFadeL 51,  l51
    NFadeL 52,  l52
    NFadeL 53,  l53
    NFadeL 54,  l54
    NFadeL 55,  l55
    NFadeL 56,  l56
    NFadeL 57,  l57
    NFadeL 58,  l58

    NFadeL 61,  l61
    NFadeL 62,  l62
    NFadeL 63,  l63
    NFadeL 64,  l64
    NFadeL 65,  l65
    NFadeL 66,  l66
    NFadeL 67,  l67
    NFadeL 68,  l68

    NFadeL 71,  l71
    NFadeL 72,  l72
    NFadeL 73,  l73
    NFadeL 74,  l74
    NFadeL 75,  l75
    NFadeL 76,  l76
    NFadeL 78,  l78



	'Bulbs
	Flash 43, f43
	Flashm 43, f43a
	Flashm 43, f43b
	Flashm 43, f43c
	NFadeLm 43, l43
	NFadeObjm 43, bc43, "bulbcover1_yellowOn", "bulbcover1_yellow"
	Flash 77, f77
	Flashm 77, f77a
	Flashm 77, f77b
	NFadeObjm 77, bc77, "bulbcover1_yellowOn", "bulbcover1_yellow"
	Flash 81,f81
	Flashm 81, f81a
	Flashm 81, f81b
	Flashm 81, f81c
	NFadeObjm 81, bc81, "bulbcover1_yellowOn", "bulbcover1_yellow"
	Flash 82, f82
	Flashm 82, f82a
	Flashm 82, f82b
	Flashm 82, f82c
	NFadeObjm 82, bc82, "bulbcover1_yellowOn", "bulbcover1_yellow"
	Flash 83,f83
	Flashm 83,f83a
	Flashm 83,f83b
	Flashm 83,f83c
	NFadeObjm 83, bc83, "bulbcover1_redOn", "bulbcover1_red"
	Flash 85,f85
	Flashm 85, f85a
	Flashm 85, f85b
	Flashm 85, f85c
	Flashm 85, f85d
	NFadeObjm 85, bc85, "bulbcover1_yellowOn", "bulbcover1_yellow"


	'RampBulbs
	NFadeLmb 84, l84l
	NFadeLmb 84, l84r
	NFadeLmb 86, l86l
	NFadeLmb 86, l86r



	'Flashers
	Flash 137, f137
	Flashm 137, leftflasher
	NFadeLm 137, l137
	NFadeLm 137, l137a

	Flash 138, f138
	Flashm 138, leftrampflasher
	NFadeLm 138, l138
	NFadeLm 138, l138a

	Flash 139, f139
	Flashm 139, f239
'	Flashm 139, backboardflasherwhite
	NFadeLm 139, l139
	NFadeLm 139, l139a
	NFadeLm 139, l139b
	NFadeLm 139, l239
	NFadeLm 139, lbackw

	Flash 140, f140
	Flashm 140, f240
'	Flashm 140, BackBoardFlasherYellow
	NFadeLm 140, l140
	NFadeLm 140, l140a
	NFadeLm 140, l140b
	NFadeLm 140, l240
	NFadeLm 140, lbacky

	Flash 141, f141
	Flashm 141, f241
'	Flashm 141, BackBoardFlasherRed
	NFadeLm 141, l141
	NFadeLm 141, l141a
	NFadeLm 141, l141b
	NFadeLm 141, l241
	NFadeLm 141, lbackr

	Flash 142, f142
	Flashm 142, BlastZoneFlasher
	NFadeLm 142, l142
	NFadeLm 142, l142a

	Flash 143, f143
	Flashm 143, f143T
	Flashm 143, rightrampflasher
	NFadeLm 143, l143
	NFadeLm 143, l143a
	NFadeLm 143, l143b

	NFadeLm 144, l144 'JetsFlasher
	NFadeLm 144, l144a
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


Sub FlasherTimer1_Timer()
	If l84l.State = 2 then
		f84l.visible = 1
		f84r.visible = 0
		Else
		f84l.visible = 0
		f84r.visible = 0
		End if
	If l84l.State = 2 then FlasherTimer2.Enabled = 1:FlasherTimer1.Enabled = 0:Else FlasherTimer1.Enabled = 1
End Sub

Sub FlasherTimer2_Timer()
	If l84l.State = 2 then
		f84l.visible = 0
		f84r.visible = 1
		Else
		f84l.visible = 0
		f84r.visible = 0
	End if
	If l84l.State = 2 then
	Flashertimer1.Enabled = 1
	FlasherTimer2.Enabled = 0
	End if
End Sub

Sub FlasherTimer3_Timer()
	If l86l.State = 2 then
		f86l.visible = 1
		f86r.visible = 0
		f86r2.visible = 0
		Else
		f86l.visible = 0
		f86r.visible = 0
		f86r2.visible = 0
	End if
	If l86l.State = 2 then FlasherTimer4.Enabled = 1:FlasherTimer3.Enabled = 0:Else FlasherTimer3.Enabled = 1
End Sub

Sub FlasherTimer4_Timer()
	If l86l.State = 2 then
		f86l.visible = 0
		f86r.visible = 1
		f86r2.visible = 1
		Else
		f86l.visible = 0
	End if
	If l86l.State = 2 then 
	FlasherTimer3.Enabled = 1
	FlasherTimer4.Enabled = 0
	End if
End Sub


'***********
' Update GI
'***********

Dim gistep, xx, obj
   gistep = 1 / 8


Sub UpdateGI(no, step)
    If step = 0 or step = 7 then exit sub
    Select Case no
        Case 3
            For each xx in RightPlayfield:xx.IntensityScale = gistep*step:next
			Table1.ColorGradeImage = "grade_" & step
        Case 4
            For each xx in LeftPlayfield:xx.IntensityScale = gistep*step:next
			Table1.ColorGradeImage = "grade_" & step
    End Select
End Sub

Sub UpdateGI2(GiNo,Status)
	Select Case GiNo
		Case 0
			If status then
				For each obj in PlayfieldInsert1: obj.state = 1:next
			Else
				For each obj in PlayfieldInsert1: obj.state = 0:Next
			End if
		Case 1
			If status then
				For each obj in PlayfieldInsert2: obj.state = 1:next
			Else
				for each obj in PlayfieldInsert2: obj.state = 0:Next
			End if
		Case 2
			If Status Then
				For each obj in PlayfieldInsert3: obj.state = 1:next
			Else
				for each obj in PlayfieldInsert3: obj.state = 0:Next
			End If
	End Select
End Sub



'****************************
'     Realtime Updates
' called by the MotorCallBack
'****************************

Sub RealTimeUpdates
    'flippers
    FlipperLP.RotY = FlipperL.CurrentAngle
    FlipperRP.RotY = FlipperR.CurrentAngle
    FlipperL2P.RotY = FlipperL2.CurrentAngle
    FlipperL3P.RotY = FlipperL3.CurrentAngle
    Diverter1P.RotY = Diverter1.CurrentAngle
    Diverter2P.RotY = Diverter2.CurrentAngle
    ' rolling sound
    RollingSoundUpdate
    'Gates
	GatePlungerLaneP.RotX = GatePlungerLane.CurrentAngle +90
End Sub

'*******************
' WireRampDropSounds
'*******************
'PlaySound "name",loopcount,volume,pan,randompitch,pitch,useexisting,restart
Dim SoundBall


Sub RWireSStart_Hit()
	Set SoundBall = ActiveBall
	PlaySound "wireramp_right", 1, 0.3, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	StopSound "plasticrolling"
End Sub

Sub RWireSStop_Hit()
	StopSound "wireramp_right"
	PlaySound "wireramp_stop", 0, 0.3, 0.2
	vpmTimer.AddTimer 200, "BallDropSound'"
End Sub

Sub LWireSStart1_Hit()
	Set SoundBall = ActiveBall
	PlaySound "wireramp_right", -1, 0.3, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	StopSound "plasticrolling"
End Sub

Sub LWireSStart2_Hit()
	Set SoundBall = ActiveBall
	PlaySound "wireramp_right", 1, 0.3, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	StopSound "plasticrolling"
End Sub

Sub LRampSStart_Hit()
	Set SoundBall = ActiveBall
	PlaySound "plasticrolling", 1, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub RRampSStart_Hit()
	Set SoundBall = ActiveBall
	PlaySound "plasticrolling", 1, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub RRampSStop_Hit()
	StopSound "plasticrolling"
End Sub

Sub RRampSStop1_Hit()
	StopSound "plasticrolling"
End Sub


Sub LRampSStop_Hit
	StopSound "plasticrolling"
End Sub

Sub LWireSStop_Hit()
	StopSound "wireramp_right"
	PlaySound "wireramp_stop", 0, 0.3, -0.2
	vpmTimer.AddTimer 200, "BallDropSound'"
End Sub

Sub FRWireSStart_Hit()
	Set SoundBall = ActiveBall
	PlaySound "wireramp_right", 1, 0.3, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub FRWireSStart1_Hit()
	StopSound "wireramp_right"
End Sub

Sub BallDropSound()
	PlaySound "BallDrop"
End Sub

'******************************
' Diverse Collection Hit Sounds
'******************************

Sub Metals_Medium_Hit(idx)
    PlaySound "metalhit_medium", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub MetalBallGuide_Hit(idx)
    PlaySound "metalhit_medium", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit(idx)
    PlaySound "gate", 0, 0.9, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinner_Hit(idx)
    PlaySound "gate4", 0, 0.7, 0, 0.25, AudioFade(ActiveBall)
End Sub

Sub Rubbers_Hit(idx)
    dim finalspeed
    finalspeed = SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
    If finalspeed > 20 then
        PlaySound "rubber_hit_4", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    End if
    If finalspeed >= 6 AND finalspeed <= 20 then
        RandomSoundRubber()
    End If
End Sub

Sub RandomSoundRubber()
    Select Case Int(Rnd * 3) + 1
        Case 1:PlaySound "rubber_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
        Case 2:PlaySound "rubber_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
        Case 3:PlaySound "rubber_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    End Select
End Sub

Sub Wall38_Hit()
	PlaySound "rubber_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub FlipperL_Collide(parm)
    RandomSoundFlipper()
End Sub

Sub FlipperL2_Collide(parm)
    RandomSoundFlipper()
End Sub
Sub FlipperL3_Collide(parm)
    RandomSoundFlipper()
End Sub


Sub FlipperR_Collide(parm)
    RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
    Select Case Int(Rnd * 3) + 1
        Case 1:PlaySound "flip_hit_1", 0, 1, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
        Case 2:PlaySound "flip_hit_2", 0, 1, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
        Case 3:PlaySound "flip_hit_3", 0, 1, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    End Select
End Sub

Sub RandomSoundBlastZone()
    Select Case Int(Rnd * 3) + 1
        Case 1:PlaySound "BlastZone_Hit1", 0, 1, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
        Case 2:PlaySound "BlastZone_Hit2", 0, 1, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
        Case 3:PlaySound "BlastZone_Hit3", 0, 1, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    End Select
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
  Vol = Csng(BallVel(ball) ^2 / VolDiv)
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

Const tnob = 9 ' total number of balls
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = FALSE
    Next
End Sub

Sub RollingSoundUpdate()
    Dim BOT, b
    BOT = GetBalls

    ' stop the sound of deleted balls
    For b = UBound(BOT) + 1 to tnob
        rolling(b) = FALSE
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


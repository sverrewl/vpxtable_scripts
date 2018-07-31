'**********************
'  Xenon Bally(1980)
' VPX table by jpsalas
' modified by HauntFreaks, Bord
' version 1.2.3
'**********************

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' No special SSF tweaks yet.
' This is a JP table. He often uses walls as switches so I need to be careful of using PlaySoundAt

Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01550000", "Bally.vbs", 3.26

Dim bsTrough, bsLeftHole, bsTopHole, bsOuthole, dtRBank, FastFlips
Const cGameName = "Xenon"

Const UseSolenoids = 1
Const UseLamps = 0
Const UseGI = 0
Const UseSync = 0
Const HandleMech = 0

'Standard Sounds
Const SSolenoidOn = "fx_Solenoid"
Const SSolenoidOff = ""
Const SCoin = "fx_Coin"

'************
' Table init.
'************

Sub Table1_Init
    vpmInit me

    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Xenon, Bally 1980" & vbNewLine & "VPX table by JPSalas v.1.0.3"
        .Games(cGameName).Settings.Value("rol") = 0 '1= rotated display, 0= normal
        .HandleMechanics = 0
        .HandleKeyboard = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .ShowTitle = 0
        .Hidden = 0
        If Err Then MsgBox Err.Description
    End With
    On Error Goto 0
    Controller.SolMask(0) = 0
    vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the timer to renable all the solenoids after 2 seconds
    Controller.Run

	Set FastFlips = new cFastFlips
	with FastFlips
	.CallBackL = "SolLflipper"  'Point these to flipper subs
	.CallBackR = "SolRflipper"  '...
	'  .CallBackUL = "SolULflipper"'...(upper flippers, if needed)
	'  .CallBackUR = "SolURflipper"'...
	.TiltObjects = True 'Optional, if True calls vpmnudge.solgameon automatically. IF YOU GET A LINE 1 ERROR, DISABLE THIS! (or setup vpmNudge.TiltObj!)
	'  .InitDelay "FastFlips", 100         'Optional, if > 0 adds some compensation for solenoid jitter (occasional problem on Bram Stoker's Dracula)
	'  .DebugOn = False        'Debug, always-on flippers. Call FastFlips.DebugOn True or False in debugger to enable/disable.
	end with

    'Nudging
    vpmNudge.TiltSwitch = 7
    vpmNudge.Sensitivity = 2
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, Bumper4, LeftSlingshot, RightSlingShot)

    '************  Trough	**************************
	sw2.CreateBall
	sw28.CreateBall
	Controller.Switch(2) = 1
	Controller.Switch(28) = 1

    'Left hole
    Set bsLeftHole = New cvpmBallStack
    With bsLeftHole
        .InitSaucer sw33, 33, 100, 10
        .KickForceVar = 2
        .KickAngleVar = 2
        .InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_kicker", DOFContactors)
        .InitAddSnd SoundFX("fx_kicker_enter", DOFContactors)
    End With

    'Right hole
    Set bsTopHole = New cvpmBallStack
    With bsTopHole
        .InitSaucer sw34, 34, 165, 10
        .KickForceVar = 2
        .KickAngleVar = 2
        .InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_kicker", DOFContactors)
        .InitAddSnd SoundFX("fx_kicker_enter", DOFContactors)
        .CreateEvents "bsTopHole", sw34
    End With

    'Droptargets
    set dtRBank = new cvpmdroptarget
    With dtRBank
        .InitDrop Array(sw21,sw22,sw23,sw24), Array(21, 22, 23, 24)
        .Initsnd SoundFX("fx_droptarget", DOFContactors), SoundFX("fx_resetdrop", DOFContactors)
        '.CreateEvents "dtRBank" 'this doesn't work with the new droptargets
    End With

    'Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

	' Remove the cabinet rails if in FS mode
If Table1.ShowDT = False then
    Ramp6.Visible = False
    Ramp9.Visible = False
    Ramp12.Visible = False
    Ramp4.Visible = False
    Ramp13.Visible = False
    Ramp10.Visible = False
	backwallmetalbrackets_prim.Visible = False
	backwallmetal_prim.Visible = False
	backwallmetalscrews_prim.Visible = False
	backwallwood_prim.Visible = False
End If
End Sub

Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub

'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If keycode = PlungerKey Then PlaySound "fx_PlungerPull", 0, 1, 0.1, 0.05:Plunger.Pullback
	If KeyCode = LeftFlipperKey then FastFlips.FlipL True :  FastFlips.FlipUL True
	If KeyCode = RightFlipperKey then FastFlips.FlipR True :  FastFlips.FlipUR True
    If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If keycode = PlungerKey Then PlaySound "fx_plunger", 0, 1, 0.1, 0.05:Plunger.Fire
	If KeyCode = LeftFlipperKey then FastFlips.FlipL False :  FastFlips.FlipUL False
	If KeyCode = RightFlipperKey then FastFlips.FlipR False :  FastFlips.FlipUR False
    If vpmKeyUp(keycode) Then Exit Sub
End Sub

'*********
' Switches
'*********

'Slings & Rubbers
Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    PlaySound SoundFX("fx_slingshot", DOFContactors), 0, 1, -0.05, 0.05
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 36
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
    PlaySound SoundFX("fx_slingshot", DOFContactors), 0, 1, 0.05, 0.05
    RightSling4.Visible = 1
    Remk.RotX = 26
    RStep = 0
    vpmTimer.PulseSw 35
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

'Rubbers

Sub sw26_Hit():PlaySound "fx_Rubber", 0, 1, -0.1, 0.15::vpmTimer.PulseSw 26:End Sub
Sub sw26a_Hit():PlaySound "fx_Rubber", 0, 1, 0.1, 0.15::vpmTimer.PulseSw 26:End Sub

'Spinner
Sub sw5_Spin():vpmTimer.pulsesw 5:PlaySound "fx_spinner":End Sub

' Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 40:PlaySound SoundFX("fx_bumper", DOFContactors), 0, 1, -0.1, 0.15:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 39:PlaySound SoundFX("fx_bumper", DOFContactors), 0, 1, 0.1, 0.15:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 38:PlaySound SoundFX("fx_bumper", DOFContactors), 0, 1, 0, 0.15:End Sub
Sub Bumper4_Hit:vpmTimer.PulseSw 37:PlaySound SoundFX("fx_bumper", DOFContactors), 0, 1, 0, 0.15:End Sub

'Rollover & Ramp Switches
Sub sw1_Hit:Controller.Switch(1) = 1:PlaySound "fx_sensor", 0, 1, -0.1, 0.15:StopSound "wireramp":End Sub
Sub sw1_UnHit:Controller.Switch(1) = 0:End Sub

Sub sw17_Hit:w17.IsDropped=1:Controller.Switch(17) = 1:PlaySound "fx_sensor", 0, 1, 0, 0.15:End Sub
Sub sw17_UnHit:w17.IsDropped=0:Controller.Switch(17) = 0:End Sub

Sub sw18_Hit:w18.IsDropped=1:Controller.Switch(18) = 1:PlaySound "fx_sensor", 0, 1, 0, 0.15:End Sub
Sub sw18_UnHit:w18.IsDropped=0:Controller.Switch(18) = 0:End Sub

Sub sw19_Hit:w19.IsDropped=1:Controller.Switch(19) = 1:PlaySound "fx_sensor", 0, 1, 0, 0.15:End Sub
Sub sw19_UnHit:w19.IsDropped=0:Controller.Switch(19) = 0:End Sub

Sub sw20_Hit:w20.IsDropped=1:Controller.Switch(20) = 1:PlaySound "fx_sensor", 0, 1, 0, 0.15:End Sub
Sub sw20_UnHit:w20.IsDropped=0:Controller.Switch(20) = 0:End Sub

Sub sw25_Hit:Controller.Switch(25) = 1:PlaySound "fx_sensor", 0, 1, 0.1, 0.15:End Sub
Sub sw25_UnHit:Controller.Switch(25) = 0:End Sub

Sub sw27_Hit:Controller.Switch(27) = 1:PlaySound "fx_sensor", 0, 1, 0.1, 0.15:End Sub
Sub sw27_UnHit:Controller.Switch(27) = 0:End Sub

Sub sw29_Hit:Controller.Switch(29) = 1:PlaySound "fx_sensor", 0, 1, -0.1, 0.15:End Sub
Sub sw29_UnHit:Controller.Switch(20) = 0:End Sub

Sub sw30_Hit:Controller.Switch(30) = 1:PlaySound "fx_sensor", 0, 1, -0.1, 0.15:End Sub
Sub sw30_UnHit:Controller.Switch(30) = 0:End Sub

Sub sw31_Hit:Controller.Switch(31) = 1:PlaySound "fx_sensor", 0, 1, 0.1, 0.15:End Sub
Sub sw31_UnHit:Controller.Switch(31) = 0:End Sub

Sub sw32_Hit:Controller.Switch(32) = 1:PlaySound "fx_sensor", 0, 1, 0.1, 0.15:End Sub
Sub sw32_UnHit:Controller.Switch(32) = 0:End Sub

'Standup Targets
Sub sw3_Hit:vpmTimer.PulseSw 3:PlaySound SoundFX("fx_target", DOFContactors), 0, 1, -0.1, 0.15:End Sub
Sub sw4_Hit:vpmTimer.PulseSw 4:PlaySound SoundFX("fx_target", DOFContactors), 0, 1, -0.1, 0.15:End Sub

'Droptargets VPX
Sub sw21_Hit:PlaySound SoundFX("fx_droptarget", DOFContactors), 0, 1, 0.1, 0.15:End Sub 'hit event only for the sound
Sub sw21_Dropped:dtRBank.hit 1 : D1L1.state=1 : End Sub

Sub sw22_Hit:PlaySound SoundFX("fx_droptarget", DOFContactors), 0, 1, 0.1, 0.15:End Sub 'hit event only for the sound
Sub sw22_Dropped:dtRBank.hit 2 : D2L1.state=1 : D2L2.state=1 : End Sub

Sub sw23_Hit:PlaySound SoundFX("fx_droptarget", DOFContactors), 0, 1, 0.1, 0.15:End Sub 'hit event only for the sound
Sub sw23_Dropped:dtRBank.hit 3 : D3L1.state=1 : D3L2.state=1 : End Sub

Sub sw24_Hit:PlaySound SoundFX("fx_droptarget", DOFContactors), 0, 1, 0.1, 0.15:End Sub 'hit event only for the sound
Sub sw24_Dropped:dtRBank.hit 4 : D4L1.state=1 : D4L2.state=1 : End Sub

' Drain & holes
Sub sw33_Hit():bsLeftHole.AddBall 0:End Sub

'******************************************************
'			TROUGH BASED ON NFOZZY'S
'******************************************************

Sub sw28_Hit():Controller.Switch(28) = 1:UpdateTrough:End Sub
Sub sw28_UnHit():Controller.Switch(28) = 0:UpdateTrough:End Sub
Sub sw2_Hit():Controller.Switch(2) = 1:UpdateTrough:End Sub
Sub sw2_UnHit():Controller.Switch(2) = 0:UpdateTrough:End Sub

Sub UpdateTrough()
	UpdateTroughTimer.Interval = 300
	UpdateTroughTimer.Enabled = 1
End Sub

Sub UpdateTroughTimer_Timer()
	If sw2.BallCntOver = 0 Then sw28.kick 30, 9
	Me.Enabled = 0
End Sub

'******************************************************
'				DRAIN & RELEASE
'******************************************************

Sub sw8_Hit()
	PlaySound "fx_drain"
	UpdateTrough
	Controller.Switch(8) = 1
End Sub

Sub sw8_UnHit()
	Controller.Switch(8) = 0
End Sub

Sub SolOuthole(enabled)
	debug.print "Outhole " & enabled
	If enabled Then
		sw8.kick 30,20
		PlaySound SoundFX("fx_Solenoid",DOFContactors)
	End If
End Sub

Sub ReleaseBall(enabled)
	debug.print "Release " & enabled
	If enabled Then
		PlaySound SoundFX("fx_ballrel",DOFContactors)
		sw2.kick 60, 7
		UpdateTrough
	End If
End Sub



'***********
' Solenoids
'***********
' from pacdudes script

SolCallback(1) = "SolDropUp"
SolCallback(2) = "dtRBank.SolHit 4,"
SolCallback(3) = "dtRBank.SolHit 3,"
SolCallback(4) = "dtRBank.SolHit 2,"
SolCallback(5) = "dtRBank.SolHit 1,"
SolCallback(6) = "vpmsolsound SoundFX(""fx_knocker"",DOFKnocker),"
SolCallback(7) = "SolOuthole"
SolCallback(8) = "ReleaseBall"
SolCallback(9) = "SolLeftOut"
SolCallback(17) = "SolTopOut"
SolCallback(19) = "FastFlips.TiltSol"
'SolCallback(19) = "ACRelay"
'SolCallback(20) = "PowerOn"

'Took a look at Xenon. While Sol18 kind of works, correct one should be Sol19.
'
'The tilt solenoid is assigned to vpmNudge.SolGameOn, but FastFlips.TiltObjects = True it will call that automatically. So you can safely replace the existing script.

''Solenoid Subs
'Sub ACRelay(enabled)
'    vpmNudge.SolGameOn enabled
'End Sub

'**************
' Flipper Subs
'**************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySound SoundFX("fx_flipperup", DOFContactors), 0, 1, -0.1, 0.25
        LeftFlipper.RotateToEnd
    Else
        PlaySound SoundFX("fx_flipperdown", DOFContactors), 0, 1, -0.1, 0.25
        LeftFlipper.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySound SoundFX("fx_flipperup", DOFContactors), 0, 1, 0.1, 0.25
        RightFlipper.RotateToEnd
    Else
        PlaySound SoundFX("fx_flipperdown", DOFContactors), 0, 1, 0.1, 0.25
        RightFlipper.RotateToStart
    End If
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, -0.1, 0.15
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, 0.1, 0.15
End Sub

'***** kicker animation

Dim LeftKick, TopKick

Sub SolLeftOut(enabled)
	If enabled Then
		bslefthole.ExitSol_On
		LeftKick = 0
		kickarmleft_prim.ObjRotY = 12
		LeftKickTimer.Enabled = 1
	End If
End Sub

Sub LeftKickTimer_Timer
    Select Case LeftKick
        Case 1:kickarmleft_prim.ObjRotY = 50
        Case 2:kickarmleft_prim.ObjRotY = 50
        Case 3:kickarmleft_prim.ObjRotY = 50
        Case 4:kickarmleft_prim.ObjRotY = 50
        Case 5:kickarmleft_prim.ObjRotY = 50
        Case 6:kickarmleft_prim.ObjRotY = 50
        Case 7:kickarmleft_prim.ObjRotY = 50
        Case 8:kickarmleft_prim.ObjRotY = 50
        Case 9:kickarmleft_prim.ObjRotY = 50
        Case 10:kickarmleft_prim.ObjRotY = 50
        Case 11:kickarmleft_prim.ObjRotY = 24
        Case 12:kickarmleft_prim.ObjRotY = 12
        Case 13:kickarmleft_prim.ObjRotY = 0:LeftKickTimer.Enabled = 0
    End Select
    LeftKick = LeftKick + 1
End Sub

Sub SoltopOut(enabled)
	If enabled Then
		bstophole.ExitSol_On
		topKick = 0
		kickarmtop_prim.ObjRotX = -12
		topKickTimer.Enabled = 1
	End If
End Sub

Sub topKickTimer_Timer
    Select Case topKick
        Case 1:kickarmtop_prim.ObjRotX = -50
        Case 2:kickarmtop_prim.ObjRotX = -50
        Case 3:kickarmtop_prim.ObjRotX = -50
        Case 4:kickarmtop_prim.ObjRotX = -50
        Case 5:kickarmtop_prim.ObjRotX = -50
        Case 6:kickarmtop_prim.ObjRotX = -50
        Case 7:kickarmtop_prim.ObjRotX = -50
        Case 8:kickarmtop_prim.ObjRotX = -50
        Case 9:kickarmtop_prim.ObjRotX = -50
        Case 10:kickarmtop_prim.ObjRotX = -50
        Case 11:kickarmtop_prim.ObjRotX = -24
        Case 12:kickarmtop_prim.ObjRotX = -12
        Case 13:kickarmtop_prim.ObjRotX = 0:topKickTimer.Enabled = 0
    End Select
    topKick = topKick + 1
End Sub

'*****Drop lights off
dim xx
For each xx in DTLights: xx.state=0:Next

Sub SolDropUp(enabled)
	dim xx
	if enabled Then
	dtRBank.SolDropUp enabled
	For each xx in DTLights: xx.state=0:Next
	end if
End Sub

'******************************************************
'        JP's VP10 Fading Lamps & Flashers
'  very reduced, mostly for rom activated flashers
' if you need to turn a light on or off then use:
'	LightState(lightnumber) = 0 or 1
'        Based on PD's Fading Light System
'******************************************************

Dim LightState(200), FlashSpeedUp(200), FlashSpeedDown(200), FlashMin(200), FlashMax(200), FlashLevel(200)

InitFlashers() ' turn off the lights and flashers and reset them to the default parameters

LampTimer.Interval = 50 'lamp fading speed
LampTimer.Enabled = 1

Sub LampTimer_timer()
    Dim chgLamp, x
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp) Then
        For x = 0 To UBound(chgLamp)
            LightState(chgLamp(x, 0) ) = chgLamp(x, 1) 'light state as set by the rom
        Next
    End If
    ' Lights & Flashers
    LightX 2, l2
    LightX 3, l3
    LightX 4, l4
    LightXm 5, l5b
    LightXm 5, l5c
    LightX 5, l5
    LightX 6, l6
    LightX 7, l7
    LightX 8, l8
    LightX 9, l9
    Flashm 10, Diode1
    Flash 10, Diode2
    LightX 11, l11
    LightX 12, l12
    Flashm 14, Diode9
    Flash 14, Diode10
    LightX 15, l15
    LightX 17, l17
    LightX 18, l18
    LightX 19, l19
    LightX 20, l20
    LightXm 21, l21b
    LightXm 21, l21c
    LightX 21, l21
    LightX 22, l22
    LightX 23, l23
    LightX 24, l24
    LightX 25, l25
    LightXm 26, l26a
    LightX 26, l26
    LightX 28, l28
    Flashm 30, Diode7
    Flash 30, Diode8
    LightX 31, l31
    LightX 33, l33
    LightX 34, l34
    LightX 35, l35
    LightX 36, l36
    LightXm 37, l37b
    LightXm 37, l37c
    LightX 37, l37
    LightX 38, l38
    LightX 39, l39
    LightX 40, l40
    LightX 41, l41
    LightX 42, l42
    LightX 44, l44
    Flashm 46, Diode5
    Flash 46, Diode6
    LightX 47, l47
    LightX 49, l49
    LightX 50, l50
    LightX 51, l51
    LightXm 52, l52b
    LightX 52, l52
    LightXm 53, l53b
    LightXm 53, l53c
    LightX 53, l53
    LightX 54, l54
    LightX 55, l55
    LightX 56, l56
    LightX 57, l57
    LightX 58, l58
    LightX 59, l59
    LightX 60, l60
    Flashm 62, Diode3
    Flash 62, Diode4
    LightX 63, l63
    LightXm 69, l69a
    LightX 69, l69
	LightX 101, l101
    Flashm 199, f199
	Flash 199, f199a
End Sub

' div lamp subs

Sub InitFlashers()
    Dim x
    For x = 0 to 200
        LightState(x) = 0     	 ' light state: 0=off, 1=on, -1=no change (on or off)
        FlashSpeedUp(x) = 0.5    ' Fade Speed Up
        FlashSpeedDown(x) = 0.25 ' Fade Speed Down
        FlashMax(x) = 1          ' the maximum intensity when on, usually 1
        FlashMin(x) = 0          ' the minimum intensity when off, usually 0
        FlashLevel(x) = 0        ' the intensity/fading of the flashers
    Next
End Sub

' VPX Lights, just turn them on or off

Sub LightX(nr, object)
    Select Case LightState(nr)
        Case 0, 1:object.state = LightState(nr):LightState(nr) = -1
    End Select
End Sub

Sub LightXm(nr, object) 'multiple lights
    Select Case LightState(nr)
        Case 0, 1:object.state = LightState(nr)
    End Select
End Sub

' VPX Flashers, changes the intensity

Sub Flash(nr, object)
    Select Case LightState(nr)
        Case 0 'off
            FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown(nr)
            If FlashLevel(nr) < FlashMin(nr) Then
                FlashLevel(nr) = FlashMin(nr)
                LightState(nr) = -1 'completely off, so stop the fading loop
            End if
            Object.IntensityScale = FlashLevel(nr)
        Case 1 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
            If FlashLevel(nr) > FlashMax(nr) Then
                FlashLevel(nr) = FlashMax(nr)
                LightState(nr) = -1 'completely on, so stop the fading loop
            End if
            Object.IntensityScale = FlashLevel(nr)
    End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it just sets the intensity
    Object.IntensityScale = FlashLevel(nr)
End Sub

'*****************************************
'			FLIPPER SHADOWS
'*****************************************

sub FlipperTimer_Timer()
	FlipperLSh.RotZ = LeftFlipper.currentangle
	FlipperRSh.RotZ = RightFlipper.currentangle

End Sub

'*****************************************
'			BALL SHADOW
'*****************************************
Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5)

Sub BallShadowUpdate_timer()
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
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) + 13
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) - 13
        End If
        ballShadow(b).Y = BOT(b).Y + 10
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
End Sub

'******************
' RealTime Updates
'******************

Set MotorCallback = GetRef("RealTimeUpdates")

Sub RealTimeUpdates
    RollingUpdate
    GIUpdate
End Sub

' Gi turns off when no balls are in play - just for fun JP :)
Dim OldGiState
OldGiState = -1 'start witht he Gi off

Sub GIUpdate
    Dim tmp, obj
    tmp = Getballs
    If UBound(tmp) <> OldGiState Then
        OldGiState = Ubound(tmp)
        If UBound(tmp) = -1 Then
            For each obj in aGiLights:obj.State = 0:Next
			LightState(199) = 0
        Else
            For each obj in aGiLights:obj.State = 1:Next
			LightState(199) = 1
        End If
    End If
End Sub
'******************************
' Diverse Collection Hit Sounds
'******************************

Sub aMetal_Wires_Hit(idx):PlaySound "fx_metalhit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Bands_Hit(idx):PlaySound "fx_rubber_band", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Posts_Hit(idx):PlaySound "fx_rubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Pins_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aWoods_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub

Sub Trigger2_Hit
	If ActiveBall.VelY < 0 then
		PlaySound"fx_metalrolling"
	End If
End Sub
Sub Trigger2_UnHit
	If ActiveBall.VelY > 0 Then
		StopSound"fx_metalrolling"
	End If
End Sub

Sub Trigger3_Hit: PlaySound"fx_rampbump1":End Sub
Sub Trigger4_Hit: PlaySound"fx_rampbump2":Stopsound"fx_metalrolling":End Sub
Sub Trigger5_Hit: PlaySound"wireramp":End Sub

'Bally Xenon
'added by Inkochnito
Sub editDips
    Dim vpmDips:Set vpmDips = New cvpmDips
    With vpmDips
        .AddForm 700, 400, "Xenon - DIP switches"
        .AddChk 7, 10, 180, Array("Match feature", &H08000000)                                                                                                                                                       'dip 28
        .AddChk 205, 10, 115, Array("Credits display", &H04000000)                                                                                                                                                   'dip 27
        .AddFrame 2, 30, 190, "Maximum credits", &H03000000, Array("10 credits", 0, "15 credits", &H01000000, "25 credits", &H02000000, "40 credits", &H03000000)                                                    'dip 25&26
        .AddFrame 2, 106, 190, "Drop target 2X lite adjustment", &H00000020, Array("2X is off at start game", 0, "2X is on at start game", &H00000020)                                                               'dip 6
        .AddFrame 2, 152, 190, "Drop target tube exit value", &H00000040, Array("exit value does not step", 0, "exit value steps up 1", &H00000040)                                                                  'dip 7
        .AddFrame 2, 198, 190, "Drop target special lite", &H00000080, Array("lite steps to 25000", 0, "lite stays lit", &H00000080)                                                                                 'dip 8
        .AddFrame 2, 248, 190, "Outlanes and flipper feed lanes", &H00002000, Array("1 lite comes on then alternates", 0, "both lites come on", &H00002000)                                                          'dip 14
        .AddFrame 2, 298, 190, "Top saucer scoring and Xenon lite", 49152, Array("scores 5,000 and no lite advances", 0, "scores 10,000 and 1 lite advance", &H00004000, "scores 10,000 and 2 lite advances", 49152) 'dip 15&16
        .AddFrame 205, 30, 190, "Balls per game", &HC0000000, Array("2 balls", &HC0000000, "3 balls", 0, "4 balls", &H80000000, "5 balls", &H40000000)                                                               'dip 31&32
        .AddFrame 205, 106, 190, "Side saucer mota special lite", &H00100000, Array("special resets with the score", 0, "special will alternate", &H00100000)                                                        'dip 21
        .AddFrame 205, 152, 190, "Side saucer mota score lites", &H00200000, Array("any lite will reset to 5,000", 0, "any lite will come on for next ball", &H00200000)                                             'dip 22
        .AddFrame 205, 198, 190, "Side saucer mota lite advance", &H00400000, Array("mota lites advance 1 at a time", 0, "mota lites advance 2 times", &H00400000)                                                   'dip 23
        .AddFrame 205, 248, 190, "Side saucer mota 50K, 90K lite", &H00800000, Array("lites step to 50K only", 0, "lites step to 90K", &H00800000)                                                                   'dip 24
        .AddFrame 205, 298, 190, "Game over attract", &H10000000, Array("no voice", 0, "voice says: try me again", &H10000000)                                                                                       'dip 29
        .AddFrame 205, 348, 190, "Top saucer first 2X lites adjust", &H20000000, Array("not in memory", 0, "in memory", &H20000000)                                                                                  'dip 30
        .AddLabel 50, 400, 300, 20, "Set selftest position 17,18 and 19 to 03 for the best gameplay."
        .AddLabel 50, 420, 300, 20, "After hitting OK, press F3 to reset game with new settings."
        .ViewDips
    End With
End Sub
Set vpmShowDips = GetRef("editDips")

Sub table1_Exit:Controller.Stop:End Sub

'cFastFlips by nFozzy
'Bypasses pinmame callback for faster and more responsive flippers
'Version 1.1 beta2 (More proper behaviour, extra safety against script errors)

'Flipper / game-on Solenoid # reference
'Atari: Sol16
'Astro:  ?
'Bally Early 80's: Sol19
'Bally late 80's (Blackwater 100, etc): Sol19
'Game Plan: Sol16
'Gottlieb System 1: Sol17
'Gottlieb System 80: No dedicated flipper solenoid? GI circuit Sol10?
'Gottlieb System 3: Sol32
'Playmatic: Sol8
'Spinball: Sol25
'Stern (80's): Sol19
'Taito: ?
'Williams System 3, 4, 6: Sol23
'Williams System 7: Sol25
'Williams System 9: Sol23
'Williams System 11: Sol23
'Bally / Williams WPC 90', 92', WPC Security: Sol31
'Data East (and Sega pre-whitestar): Sol23
'Zaccaria: ???

'********************Setup*******************:

'....somewhere outside of any subs....
'dim FastFlips

'....table init....
'Set FastFlips = new cFastFlips
'with FastFlips
'   .CallBackL = "SolLflipper"  'Point these to flipper subs
'   .CallBackR = "SolRflipper"  '...
''  .CallBackUL = "SolULflipper"'...(upper flippers, if needed)
''  .CallBackUR = "SolURflipper"'...
'   .TiltObjects = True 'Optional, if True calls vpmnudge.solgameon automatically. IF YOU GET A LINE 1 ERROR, DISABLE THIS! (or setup vpmNudge.TiltObj!)
''  .InitDelay "FastFlips", 100         'Optional, if > 0 adds some compensation for solenoid jitter (occasional problem on Bram Stoker's Dracula)
''  .DebugOn = False        'Debug, always-on flippers. Call FastFlips.DebugOn True or False in debugger to enable/disable.
'end with

'...keydown section... commenting out upper flippers is not necessary as of 1.1
'If KeyCode = LeftFlipperKey then FastFlips.FlipL True :  FastFlips.FlipUL True
'If KeyCode = RightFlipperKey then FastFlips.FlipR True :  FastFlips.FlipUR True
'(Do not use Exit Sub, this script does not handle switch handling at all!)

'...keyUp section...
'If KeyCode = LeftFlipperKey then FastFlips.FlipL False :  FastFlips.FlipUL False
'If KeyCode = RightFlipperKey then FastFlips.FlipR False :  FastFlips.FlipUR False

'...Solenoid...
'SolCallBack(31) = "FastFlips.TiltSol"
'//////for a reference of solenoid numbers, see top /////


'One last note - Because this script is super simple it will call flipper return a lot.
'It might be a good idea to add extra conditional logic to your flipper return sounds so they don't play every time the game on solenoid turns off
'Example:
'Instead of
        'LeftFlipper.RotateToStart
        'playsound SoundFX("FlipperDown",DOFFlippers), 0, 1, 0.01   'return
'Add Extra conditional logic:
        'LeftFlipper.RotateToStart
        'if LeftFlipper.CurrentAngle = LeftFlipper.StartAngle then
        '   playsound SoundFX("FlipperDown",DOFFlippers), 0, 1, 0.01    'return
        'end if
'That's it]
'*************************************************
Function NullFunction(aEnabled):End Function    '1 argument null function placeholder
Class cFastFlips
    Public TiltObjects, DebugOn, hi
    Private SubL, SubUL, SubR, SubUR, FlippersEnabled, Delay, LagCompensation, Name, FlipState(3)

    Private Sub Class_Initialize()
        Delay = 0 : FlippersEnabled = False : DebugOn = False : LagCompensation = False
        Set SubL = GetRef("NullFunction"): Set SubR = GetRef("NullFunction") : Set SubUL = GetRef("NullFunction"): Set SubUR = GetRef("NullFunction")
    End Sub

    'set callbacks
    Public Property Let CallBackL(aInput)  : Set SubL  = GetRef(aInput) : Decouple sLLFlipper, aInput: End Property
    Public Property Let CallBackUL(aInput) : Set SubUL = GetRef(aInput) : End Property
    Public Property Let CallBackR(aInput)  : Set SubR  = GetRef(aInput) : Decouple sLRFlipper, aInput:  End Property
    Public Property Let CallBackUR(aInput) : Set SubUR = GetRef(aInput) : End Property
    Public Sub InitDelay(aName, aDelay) : Name = aName : delay = aDelay : End Sub   'Create Delay
    'Automatically decouple flipper solcallback script lines (only if both are pointing to the same sub) thanks gtxjoe
    Private Sub Decouple(aSolType, aInput)  : If StrComp(SolCallback(aSolType),aInput,1) = 0 then SolCallback(aSolType) = Empty End If : End Sub

    'call callbacks
    Public Sub FlipL(aEnabled)
        FlipState(0) = aEnabled 'track flipper button states: the game-on sol flips immediately if the button is held down (1.1)
        If not FlippersEnabled and not DebugOn then Exit Sub
        subL aEnabled
    End Sub

    Public Sub FlipR(aEnabled)
        FlipState(1) = aEnabled
        If not FlippersEnabled and not DebugOn then Exit Sub
        subR aEnabled
    End Sub

    Public Sub FlipUL(aEnabled)
        FlipState(2) = aEnabled
        If not FlippersEnabled and not DebugOn then Exit Sub
        subUL aEnabled
    End Sub

    Public Sub FlipUR(aEnabled)
        FlipState(3) = aEnabled
        If not FlippersEnabled and not DebugOn then Exit Sub
        subUR aEnabled
    End Sub

    Public Sub TiltSol(aEnabled)    'Handle solenoid / Delay (if delayinit)
        If delay > 0 and not aEnabled then  'handle delay
            vpmtimer.addtimer Delay, Name & ".FireDelay" & "'"
            LagCompensation = True
        else
            If Delay > 0 then LagCompensation = False
            EnableFlippers(aEnabled)
        end If
    End Sub

    Sub FireDelay() : If LagCompensation then EnableFlippers False End If : End Sub

    Private Sub EnableFlippers(aEnabled)
        If aEnabled then SubL FlipState(0) : SubR FlipState(1) : subUL FlipState(2) : subUR FlipState(3)
        FlippersEnabled = aEnabled
        If TiltObjects then vpmnudge.solgameon aEnabled
        If Not aEnabled then
            subL False
            subR False
            If not IsEmpty(subUL) then subUL False
            If not IsEmpty(subUR) then subUR False
        End If
    End Sub

End Class

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

'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 5 ' total number of balls in this table is 4, but always use a higher number here because of the timing
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub RollingUpdate()
    Dim BOT, b, ballpitcht
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


' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


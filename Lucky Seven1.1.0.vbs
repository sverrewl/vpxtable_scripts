'*******************************************************************************************************
'
'									 Lucky Seven Williams 1978 v1.0.0
'								 http://www.ipdb.org/machine.cgi?id=1491
'
'											Created by Kiwi
'
'*******************************************************************************************************

Option Explicit
Randomize

' Thalamus 2018-07-23
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' No special SSF tweaks yet.

'*******************************************************************************************************

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

'********************************************** OPTIONS ************************************************

Dim DmdHidden

 If Table1.ShowDT = True Then
	DmdHidden = 1
Else
	DmdHidden = 0		'Put 1 if you want to have DMD hidden in FS mode
End If

 If B2SOn = True Then DmdHidden = 1

Const cGameName   = "lucky_l1"	'ROM

'********************************************** Volume Chimes

Const VolChimes = 0.5

'********************************************** Rails Hidden/Visible in FS mode : 0 hidden , 1 visible

Const RailsVisible = 0

'********************************************** Ball

'Const BallSize    = 50

'Const BallMass    = 1			'Mass=(50^3)/125000 ,(BallSize^3)/125000

'******************************************** OPTIONS END **********************************************

LoadVPM "01120100", "s4.vbs", 3.36

Dim bsTrough, dtBank, PinPlay

Const UseSolenoids = 2
Const UseLamps = 0
Const UseGI = 0
Const UseSync = 0
Const HandleMech = 0

' Standard Sounds
Const SSolenoidOn = "Solenoid"
Const SSolenoidOff = ""
Const SFlipperOn = ""
Const SFlipperOff = ""
Const SCoin = "coin3"

'************
' Table init.
'************

Sub Table1_Init
	With Controller
		.GameName = cGameName
		.SplashInfoLine = "Lucky Seven Williams 1978" & vbNewLine & "VPX table by Kiwi 1.0.0"
		.HandleMechanics = 0
		.HandleKeyboard = 0
		.ShowDMDOnly = 1
		.ShowFrame = 0
		.ShowTitle = 0
		.Hidden = DmdHidden
'		.Games(cGameName).Settings.Value("dmd_pos_x")=0
'		.Games(cGameName).Settings.Value("dmd_pos_y")=0
'		.Games(cGameName).Settings.Value("dmd_width")=400
'		.Games(cGameName).Settings.Value("dmd_height")=92
'		.Games(cGameName).Settings.Value("rol") = 0
'		.Games(cGameName).Settings.Value("sound") = 1
		If Err Then MsgBox Err.Description
	End With
	On Error Goto 0
	Controller.SolMask(0) = 0
	vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the timer to renable all the solenoids after 2 seconds
	Controller.Run

	' Nudging
	vpmNudge.TiltSwitch = 31
	vpmNudge.Sensitivity = 1
	vpmNudge.TiltObj = Array(Bumper1, Bumper2, LeftSlingShot, RightSlingShot)

	' Trough
	Set bsTrough = New cvpmBallStack
	With bsTrough
		.InitSw 0, 23, 0, 0, 0, 0, 0, 0
		.InitKick BallRelease, 90, 5
		.InitEntrySnd "Solenoid", "Solenoid"
		.InitExitSnd SoundFX("popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
		.Balls = 1
	End With

	' Drop targets
	set dtBank = new cvpmdroptarget
	With dtBank
		.InitDrop Array(sw15, sw16), Array(15, 16)
'		.Initsnd  SoundFX("DROPTARG",DOFContactors), SoundFX("resetdrop",DOFContactors)
		.AllDownSw = 17
'		.CreateEvents "dtBank"
	End With

	' Main Timer init
	PinMAMETimer.Interval = PinMAMEInterval
	PinMAMETimer.Enabled = 1

' Init Reels

'	Controller.Switch(27) = 1
	Controller.Switch(50) = 1
'	Controller.Switch(41) = 1


	If ShowDT=True Then

		da1.visible=1
		da2.visible=1
		da3.visible=1
		da4.visible=1
		da5.visible=1
		da6.visible=1
		db1.visible=1
		db2.visible=1
		db3.visible=1
		db4.visible=1
		db5.visible=1
		db6.visible=1
		dc1.visible=1
		dc2.visible=1
		dc3.visible=1
		dc4.visible=1
		dc5.visible=1
		dc6.visible=1
		dd1.visible=1
		dd2.visible=1
		dd3.visible=1
		dd4.visible=1
		dd5.visible=1
		dd6.visible=1
		de1.visible=1
		de2.visible=1
		de3.visible=1
		de4.visible=1

		l50.visible=1
		l51.visible=1
		l52.visible=1
		l53.visible=1
		l57.visible=1
		l58.visible=1
		l59.visible=1
		l60.visible=1

		L54.visible=1
		L55.visible=1
		L61.visible=1
		L62.visible=1
		L63.visible=1
		L64.visible=1

		RailSX.visible=1
		RailDX.visible=1
	Else
		RailSX.visible=RailsVisible
		RailDX.visible=RailsVisible
		BulbB1D.TransmissionScale=0
		BulbB2D.TransmissionScale=0
'		l50.visible=0
'		l51.visible=0
'		l52.visible=0
'		l53.visible=0
'		l57.visible=0
'		l58.visible=0
'		l59.visible=0
'		l60.visible=0

	End If

End Sub
Sub Table1_Exit():Controller.Stop:End Sub
Sub Table1_Paused:Controller.Pause = 1:End Sub
Sub Table1_unPaused:Controller.Pause = 0:End Sub

'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
	If KeyCode = LeftFlipperKey And PinPlay=1 Then LeftFlipper.RotateToEnd
	If KeyCode = RightFlipperKey And PinPlay=1 Then RightFlipper.RotateToEnd
	If keycode = LeftTiltKey Then Nudge 90, 4:PlaySound SoundFX("nudge_left",0)
	If keycode = RightTiltKey Then Nudge 270, 4:PlaySound SoundFX("nudge_right",0)
	If keycode = CenterTiltKey Then Nudge 0, 5:PlaySound SoundFX("nudge_forward",0)
	If vpmKeyDown(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Pullback
End Sub

Sub table1_KeyUp(ByVal Keycode)
	If KeyCode = LeftFlipperKey And PinPlay=1 Then LeftFlipper.RotateToStart
	If KeyCode = RightFlipperKey And PinPlay=1 Then RightFlipper.RotateToStart
	If vpmKeyUp(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Fire:PlaySound "plunger2"
End Sub

'*********
' Switches
'*********

' Slings

Dim LStep, RStep

Sub LeftSlingshot_Slingshot:vpmTimer.PulseSw 24:LeftSling.Visible=1:LStep=0:EmKickerT1L.TransX = -22:Me.TimerEnabled=1:PlaySound SoundFX("Slingshot",DOFContactors):End Sub
Sub LeftSlingshot_Timer
	Select Case LStep
		Case 0:LeftSling.Visible = 1:EmKickerT1L.TransX = -22
		Case 1: 'pause
		Case 2:LeftSling.Visible = 0 :LeftSling1.Visible = 1:EmKickerT1L.TransX = -17
		Case 3:LeftSling1.Visible = 0:LeftSling2.Visible = 1:EmKickerT1L.TransX = -11
		Case 4:LeftSling2.Visible = 0:Me.TimerEnabled = 0:EmKickerT1L.TransX = 0
	End Select
	LStep = LStep + 1
End Sub

Sub RightSlingshot_Slingshot:vpmTimer.PulseSw 36:RightSling.Visible=1:RStep=0:EmKickerT1R.TransX = -22:Me.TimerEnabled=1:PlaySound SoundFX("Slingshot",DOFContactors):End Sub
Sub RightSlingshot_Timer
	Select Case RStep
		Case 0:RightSling.Visible = 1:EmKickerT1R.TransX = -22
		Case 1: 'pause
		Case 2:RightSling.Visible = 0 :RightSling1.Visible = 1:EmKickerT1R.TransX = -17
		Case 3:RightSling1.Visible = 0:RightSling2.Visible = 1:EmKickerT1R.TransX = -11
		Case 4:RightSling2.Visible = 0:Me.TimerEnabled = 0:EmKickerT1R.TransX = 0
	End Select
	RStep = RStep + 1
End Sub

'Rubbers
Sub sw37_Hit:vpmTimer.PulseSw 37:PlaySound "rubber1", 0, Vol(ActiveBall), pan(ActiveBall), 0.25:End Sub
Sub sw38_Hit:vpmTimer.PulseSw 38:PlaySound "rubber1", 0, Vol(ActiveBall), pan(ActiveBall), 0.25:End Sub

' Bumpers

Sub Bumper1_Hit:Me.TimerEnabled=1:vpmTimer.PulseSw 34:Ring1.Z = -30:PlaySound SoundFX("jet1",DOFContactors),0,1,-0.1,0:End Sub
Sub Bumper1_Timer()
	Ring1.Z = Ring1.Z +2
 If Ring1.Z = 0 Then:Me.TimerEnabled = 0
End Sub

Sub Bumper2_Hit:Me.TimerEnabled=1:vpmTimer.PulseSw 33:Ring2.Z = -30:PlaySound SoundFX("jet1",DOFContactors),0,1,0,0:End Sub
Sub Bumper2_Timer()
	Ring2.Z = Ring2.Z +2
 If Ring2.Z = 0 Then:Me.TimerEnabled = 0
End Sub

' Spinner
Sub sw14_Spin:vpmTimer.PulseSw 14:PlaySound "spinner",0,1,0.1,0:End Sub

' Kickers
Sub Drain_Hit:bsTrough.AddBall Me:PlaySound "drain1a":End Sub

' Rollovers
Sub sw10_Hit:  Controller.Switch(10) = 1:PlaySound "sensor":Psw10.Z=-30:End Sub
Sub sw10_UnHit:Controller.Switch(10) = 0:Psw10.Z=-8:End Sub
Sub sw11_Hit:  Controller.Switch(11) = 1:PlaySound "sensor":Psw11.Z=-30:End Sub
Sub sw11_UnHit:Controller.Switch(11) = 0:Psw11.Z=-8:End Sub
Sub sw12_Hit:  Controller.Switch(12) = 1:PlaySound "sensor":Psw12.Z=-30:End Sub
Sub sw12_UnHit:Controller.Switch(12) = 0:Psw12.Z=-8:End Sub
Sub sw13_Hit:  Controller.Switch(13) = 1:PlaySound "sensor":Psw13.Z=-30:End Sub
Sub sw13_UnHit:Controller.Switch(13) = 0:Psw13.Z=-8:End Sub
Sub sw19_Hit:  Controller.Switch(19) = 1:PlaySound "sensor":Psw19.Z=-30:End Sub
Sub sw19_UnHit:Controller.Switch(19) = 0:Psw19.Z=-8:End Sub
Sub sw20_Hit:  Controller.Switch(20) = 1:PlaySound "sensor":Psw20.Z=-30:End Sub
Sub sw20_UnHit:Controller.Switch(20) = 0:Psw20.Z=-8:End Sub
Sub sw21_Hit:  Controller.Switch(21) = 1:PlaySound "sensor":Psw21.Z=-30:End Sub
Sub sw21_UnHit:Controller.Switch(21) = 0:Psw21.Z=-8:End Sub
Sub sw22_Hit:  Controller.Switch(22) = 1:PlaySound "sensor":Psw22.Z=-30:End Sub
Sub sw22_UnHit:Controller.Switch(22) = 0:Psw22.Z=-8:End Sub

' Targets
Sub sw9_Hit:vpmTimer.PulseSw 9:Sw9a.TransY=-5:Me.TimerEnabled=1:PlaySound SoundFX("target",DOFContactors),0,1,-0.2,0:End Sub
Sub Sw9_Timer():Sw9a.TransY=0:Me.TimerEnabled=0:End Sub
Sub sw18_Hit:vpmTimer.PulseSw 18:Sw18a.TransY=-5:Me.TimerEnabled=1:PlaySound SoundFX("target",DOFContactors),0,1,-0.2,0:End Sub
Sub Sw18_Timer():Sw18a.TransY=0:Me.TimerEnabled=0:End Sub
Sub sw35_Hit:vpmTimer.PulseSw 35:Sw35a.TransY=-5:Me.TimerEnabled=1:PlaySound SoundFX("target",DOFContactors):End Sub
Sub Sw35_Timer():Sw35a.TransY=0:Me.TimerEnabled=0:End Sub
Sub sw39_Hit:vpmTimer.PulseSw 39:Sw39a.TransY=-5:Me.TimerEnabled=1:PlaySound SoundFX("target",DOFContactors):End Sub
Sub Sw39_Timer():Sw39a.TransY=0:Me.TimerEnabled=0:End Sub
Sub sw40_Hit:vpmTimer.PulseSw 40:Sw40a.TransY=-5:Me.TimerEnabled=1:PlaySound SoundFX("target",DOFContactors):End Sub
Sub Sw40_Timer():Sw40a.TransY=0:Me.TimerEnabled=0:End Sub

' Drop Targets
Sub sw15_Hit:PlaySound SoundFX("DROPTARG",DOFContactors),0,1,-0.2,0:End Sub
Sub sw15_Dropped:dtBank.hit 1:Bulb9DT.State=1:End Sub
Sub sw16_Hit:PlaySound SoundFX("DROPTARG",DOFContactors),0,1,0.2,0:End Sub
Sub sw16_Dropped:dtBank.hit 2:Bulb10DT.State=1:End Sub

Sub Gate_Hit:PlaySound "Gate51":End Sub

'*********
' Reels
'*********

' Left Reel

' 0	sw29 Ferro di Cavallo
' 1 sw30 Diamante
' 2 sw27 Trifoglio
' 3 sw28 Sette
' 4 sw27 Trifoglio
' 5 sw29 Ferro di Cavallo
' 6 sw30 Diamante
' 7 sw27 Trifoglio
' 8 sw28 Sette
' 9 sw27 Trifoglio

' Center Reel

' 0	sw29 Ferro di Cavallo
' 1 sw30 Diamante
' 2 sw29 Ferro di Cavallo
' 3 sw28 Sette
' 4 sw50 Trifoglio
' 5 sw29 Ferro di Cavallo
' 6 sw30 Diamante
' 7 sw29 Ferro di Cavallo
' 8 sw28 Sette
' 9 sw50 Trifoglio

' Right Reel

' 0	sw29 Ferro di Cavallo
' 1 sw30 Diamante
' 2 sw41 Trifoglio
' 3 sw28 Sette
' 4 sw41 Trifoglio
' 5 sw29 Ferro di Cavallo
' 6 sw30 Diamante
' 7 sw41 Trifoglio
' 8 sw28 Sette
' 9 sw41 Trifoglio

'*********
'Solenoids
'*********
SolCallback(1) = "LeftReel"
SolCallback(2) = "CenterReel"
SolCallback(3) = "RightReel"
SolCallback(4) = "SolReset 1,"
SolCallback(5) = "SolReset 2,"
SolCallback(6) = "bsTrough.SolOut"
Solcallback(9) = "Chime10Sound"
Solcallback(10) = "Chime100Sound"
Solcallback(11) = "Chime1000Sound"
Solcallback(12) = "Chime10000Sound"
'Solcallback(13) = "vpmSolSound SoundFX(""NoiseDrum"",DOFKnocker),"
Solcallback(14) = "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
Solcallback(15) = "BuzzerSound"
Solcallback(23) = "SolRun"

Dim LR28sw, LR29sw, LR30sw, LRPos, LRAngl
	LRPos = 0
	LRAngl = 0
'	LR28sw=0:CR28sw=0:RR28sw=0

Sub LeftReel(Enabled)
	If Enabled Then
	Select Case LRPos
		Case 0:LR28sw=0:LR29sw=1:LR30sw=0:Controller.Switch(27) = 0	'36
		Case 1:LR28sw=0:LR29sw=0:LR30sw=1:Controller.Switch(27) = 0	'72
		Case 2:LR28sw=0:LR29sw=0:LR30sw=0:Controller.Switch(27) = 1	'108
		Case 3:LR28sw=1:LR29sw=0:LR30sw=0:Controller.Switch(27) = 0	'144
		Case 4:LR28sw=0:LR29sw=0:LR30sw=0:Controller.Switch(27) = 1	'180
		Case 5:LR28sw=0:LR29sw=1:LR30sw=0:Controller.Switch(27) = 0	'216
		Case 6:LR28sw=0:LR29sw=0:LR30sw=1:Controller.Switch(27) = 0	'252
		Case 7:LR28sw=0:LR29sw=0:LR30sw=0:Controller.Switch(27) = 1	'288
		Case 8:LR28sw=1:LR29sw=0:LR30sw=0:Controller.Switch(27) = 0	'324
		Case 9:LR28sw=0:LR29sw=0:LR30sw=0:Controller.Switch(27) = 1	'0
	End Select
	LRPos = LRPos + 1
 If LRPos > 9 Then
	LRPos = 0
End If
End If
	LRTimer.Enabled=1
	PlaySound SoundFX("cluper",DOFContactors)
End Sub

Sub LRTimer_Timer()
	LRAngl = LRAngl + 6
	ReelLeft.ObjRotX = LRAngl
	If ReelLeft.ObjRotX = 360 Then:LRTimer.Enabled=0:LRAngl=0:End If
	If ReelLeft.ObjRotX =  36 Then:LRTimer.Enabled=0:End If
	If ReelLeft.ObjRotX =  72 Then:LRTimer.Enabled=0:End If
	If ReelLeft.ObjRotX = 108 Then:LRTimer.Enabled=0:End If
	If ReelLeft.ObjRotX = 144 Then:LRTimer.Enabled=0:End If
	If ReelLeft.ObjRotX = 180 Then:LRTimer.Enabled=0:End If
	If ReelLeft.ObjRotX = 216 Then:LRTimer.Enabled=0:End If
	If ReelLeft.ObjRotX = 252 Then:LRTimer.Enabled=0:End If
	If ReelLeft.ObjRotX = 288 Then:LRTimer.Enabled=0:End If
	If ReelLeft.ObjRotX = 324 Then:LRTimer.Enabled=0:End If
	ReelTimer.Enabled=1
End Sub

Dim CR28sw, CR29sw, CR30sw, CRPos, CRAngl
	CRPos = 0
	CRAngl = 0
'	LR29sw=0:CR29sw=0:RR29sw=0

Sub CenterReel(Enabled)
	If Enabled Then
	Select Case CRPos
		Case 0:CR28sw=0:CR29sw=1:CR30sw=0:Controller.Switch(50) = 0	'36
		Case 1:CR28sw=0:CR29sw=0:CR30sw=1:Controller.Switch(50) = 0	'72
		Case 2:CR28sw=0:CR29sw=1:CR30sw=0:Controller.Switch(50) = 0	'108
		Case 3:CR28sw=1:CR29sw=0:CR30sw=0:Controller.Switch(50) = 0	'144
		Case 4:CR28sw=0:CR29sw=0:CR30sw=0:Controller.Switch(50) = 1	'180
		Case 5:CR28sw=0:CR29sw=1:CR30sw=0:Controller.Switch(50) = 0	'216
		Case 6:CR28sw=0:CR29sw=0:CR30sw=1:Controller.Switch(50) = 0	'252
		Case 7:CR28sw=0:CR29sw=1:CR30sw=0:Controller.Switch(50) = 0	'288
		Case 8:CR28sw=1:CR29sw=0:CR30sw=0:Controller.Switch(50) = 0	'324
		Case 9:CR28sw=0:CR29sw=0:CR30sw=0:Controller.Switch(50) = 1	'0
	End Select
	CRPos = CRPos + 1
 If CRPos > 9 Then
	CRPos = 0
End If
End If
	CRTimer.Enabled=1
	PlaySound SoundFX("cluper",DOFContactors)
End Sub

Sub CRTimer_Timer()
	CRAngl = CRAngl + 6
	ReelCenter.ObjRotX = CRAngl
	If ReelCenter.ObjRotX = 360 Then:CRTimer.Enabled=0:CRAngl=0:End If
	If ReelCenter.ObjRotX =  36 Then:CRTimer.Enabled=0:End If
	If ReelCenter.ObjRotX =  72 Then:CRTimer.Enabled=0:End If
	If ReelCenter.ObjRotX = 108 Then:CRTimer.Enabled=0:End If
	If ReelCenter.ObjRotX = 144 Then:CRTimer.Enabled=0:End If
	If ReelCenter.ObjRotX = 180 Then:CRTimer.Enabled=0:End If
	If ReelCenter.ObjRotX = 216 Then:CRTimer.Enabled=0:End If
	If ReelCenter.ObjRotX = 252 Then:CRTimer.Enabled=0:End If
	If ReelCenter.ObjRotX = 288 Then:CRTimer.Enabled=0:End If
	If ReelCenter.ObjRotX = 324 Then:CRTimer.Enabled=0:End If
	ReelTimer.Enabled=1
End Sub

Dim RR28sw, RR29sw, RR30sw, RRPos, RRAngl
	RRPos = 0
	RRAngl = 0
'	LR30sw=0:CR30sw=0:RR30sw=0

Sub RightReel(Enabled)
	If Enabled Then
	Select Case RRPos
		Case 0:RR28sw=0:RR29sw=1:RR30sw=0:Controller.Switch(41) = 0	'36
		Case 1:RR28sw=0:RR29sw=0:RR30sw=1:Controller.Switch(41) = 0	'72
		Case 2:RR28sw=0:RR29sw=0:RR30sw=0:Controller.Switch(41) = 1	'108
		Case 3:RR28sw=1:RR29sw=0:RR30sw=0:Controller.Switch(41) = 0	'144
		Case 4:RR28sw=0:RR29sw=0:RR30sw=0:Controller.Switch(41) = 1	'180
		Case 5:RR28sw=0:RR29sw=1:RR30sw=0:Controller.Switch(41) = 0	'216
		Case 6:RR28sw=0:RR29sw=0:RR30sw=1:Controller.Switch(41) = 0	'252
		Case 7:RR28sw=0:RR29sw=0:RR30sw=0:Controller.Switch(41) = 1	'288
		Case 8:RR28sw=1:RR29sw=0:RR30sw=0:Controller.Switch(41) = 0	'324
		Case 9:RR28sw=0:RR29sw=0:RR30sw=0:Controller.Switch(41) = 1	'0
	End Select
	RRPos = RRPos + 1
 If RRPos > 9 Then
	RRPos = 0
End If
End If
	RRTimer.Enabled=1
	PlaySound SoundFX("cluper",DOFContactors)
End Sub

Sub RRTimer_Timer()
	RRAngl = RRAngl + 6
	ReelRight.ObjRotX = RRAngl
	If ReelRight.ObjRotX = 360 Then:RRTimer.Enabled=0:RRAngl=0:End If
	If ReelRight.ObjRotX =  36 Then:RRTimer.Enabled=0:End If
	If ReelRight.ObjRotX =  72 Then:RRTimer.Enabled=0:End If
	If ReelRight.ObjRotX = 108 Then:RRTimer.Enabled=0:End If
	If ReelRight.ObjRotX = 144 Then:RRTimer.Enabled=0:End If
	If ReelRight.ObjRotX = 180 Then:RRTimer.Enabled=0:End If
	If ReelRight.ObjRotX = 216 Then:RRTimer.Enabled=0:End If
	If ReelRight.ObjRotX = 252 Then:RRTimer.Enabled=0:End If
	If ReelRight.ObjRotX = 288 Then:RRTimer.Enabled=0:End If
	If ReelRight.ObjRotX = 324 Then:RRTimer.Enabled=0:End If
	ReelTimer.Enabled=1
End Sub

Sub ReelTimer_Timer
	If LR28sw + CR28sw + RR28sw = 3 Then
	Controller.Switch(28) = 1:Controller.Switch(29) = 0:Controller.Switch(30) = 0
End If
	If LR29sw + CR29sw + RR29sw = 3 Then
	Controller.Switch(28) = 0:Controller.Switch(29) = 1:Controller.Switch(30) = 0
End If
	If LR30sw + CR30sw + RR30sw = 3 Then
	Controller.Switch(28) = 0:Controller.Switch(29) = 0:Controller.Switch(30) = 1
End If
	If LR28sw + CR28sw + RR28sw < 3 And LR29sw + CR29sw + RR29sw < 3 And LR30sw + CR30sw + RR30sw < 3 Then
	Controller.Switch(28) = False:Controller.Switch(29) = False:Controller.Switch(30) = False
End If
	ReelTimer.Enabled=0
End Sub

Sub SolReset(No ,Enabled)
	If Enabled Then
		Controller.Switch(17) = False
			dtBank.SolUnHit 1, True
			dtBank.SolUnHit 2, True
	PlaySound SoundFX("DTResetB",DOFContactors)
	Bulb9DT.TimerEnabled=1
	Bulb10DT.TimerEnabled=1
	End If
End Sub

Sub SolRun(Enabled)
	vpmNudge.SolGameOn Enabled
 If Enabled Then
	PinPlay=1
	LeftSlingShot.Disabled=0
	RightSlingShot.Disabled=0
Else
	PinPlay=0
	LeftFlipper.RotateToStart
	RightFlipper.RotateToStart
	LeftSlingShot.Disabled=1
	RightSlingShot.Disabled=1
End If
End Sub

Sub Bulb9DT_Timer()
	Bulb9DT.State=0
	Bulb9DT.TimerEnabled=0
End Sub

Sub Bulb10DT_Timer()
	Bulb10DT.State=0
	Bulb10DT.TimerEnabled=0
End Sub

' add additional (optional) parameters to PlaySound to increase/decrease the frequency,
' apply all the settings to an already playing sample and choose if to restart this sample from the beginning or not
' PlaySound "name",loopcount,volume,pan,randompitch,pitch,useexisting,restart
' pan ranges from -1.0 (left) over 0.0 (both) to 1.0 (right)
' randompitch ranges from 0.0 (no randomization) to 1.0 (vary between half speed to double speed)
' pitch can be positive or negative and directly adds onto the standard sample frequency
' useexisting is 0 or 1 (if no existing/playing copy of the sound is found, then a new one is created)
' restart is 0 or 1 (only useful if useexisting is 1)

Sub Chime10Sound(Enabled)
	If Enabled Then
	PlaySound SoundFX("NoNota1a",DOFChimes),0,VolChimes,0,0,0,0,0
	End If
End Sub

Sub Chime100Sound(Enabled)
	If Enabled Then
	PlaySound SoundFX("NoNota3a",DOFChimes),0,VolChimes,0,0,0,0,0
	End If
End Sub

Sub Chime1000Sound(Enabled)
	If Enabled Then
	PlaySound SoundFX("NoNota",DOFChimes),0,VolChimes,0,0,0,0,0
	End If
End Sub

Sub Chime10000Sound(Enabled)
	If Enabled Then
	PlaySound SoundFX("NoNota2a",DOFChimes),0,VolChimes,0,0,-1,0,0
	End If
End Sub

Sub BuzzerSound(Enabled)
	If Enabled Then
	PlaySound "Buzzer",0,VolChimes,0,0,-1,0,0
	End If
End Sub

'**************
' Flipper Subs
'**************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
	If Enabled Then
		PlaySound SoundFX("flipperup_left",DOFContactors)':LeftFlipper.RotateToEnd
	Else
		PlaySound SoundFX("flipperdown_left",DOFContactors)':LeftFlipper.RotateToStart
	End If
End Sub

Sub SolRFlipper(Enabled)
	If Enabled Then
		PlaySound SoundFX("flipperup_right",DOFContactors)':RightFlipper.RotateToEnd
	Else
		PlaySound SoundFX("flipperdown_right",DOFContactors)':RightFlipper.RotateToStart
	End If
End Sub

Sub LeftFlipper_Collide(parm)
	PlaySound "rubber_flipper", 0, Vol(ActiveBall), pan(ActiveBall)
End Sub

Sub RightFlipper_Collide(parm)
	PlaySound "rubber_flipper", 0, Vol(ActiveBall), pan(ActiveBall)
End Sub

'******************
' RealTime Updates
'******************
Set MotorCallback = GetRef("GameTimer")

Sub GameTimer
	UpdateGates
End Sub

Dim PI
PI = Round(4 * Atn(1), 6) '3.1415926535897932384626433832795

Sub UpdateGates
	GateA.RotX = Gate.CurrentAngle*0.6
	LogoL.RotZ = LeftFlipper.CurrentAngle
	LogoR.RotZ = RightFlipper.CurrentAngle
	pSpinnerRod.TransX = sin( (sw14.CurrentAngle+180) * (2*PI/360)) * 8
	pSpinnerRod.TransZ = sin( (sw14.CurrentAngle- 90) * (2*PI/360)) * 8
	pSpinnerRod.RotY = sin( (sw14.CurrentAngle-180) * (2*PI/360)) * 6
End Sub

'**********************************
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

	UpdateLeds
    UpdateLamps
End Sub

Sub UpdateLamps
'	FadeL 0, l0, "On", "F66", "F33", "Off"
	NFadeLm 25, l25
	NFadeL 25, l25a
	NFadeLm 26, l26
	NFadeL 26, l26a
	NFadeLm 27, l27
	NFadeL 27, l27a
	NFadeLm 28, l28
	NFadeL 28, l28a
	NFadeL 29, l29
	NFadeL 30, l30
	NFadeL 31, l31
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
	NFadeL 45, l45
	NFadeL 46, l46
	NFadeL 47, l47
	NFadeL 48, l48
	NFadeL 49, l49
	NFadeL 56, l56

'Backdrop

	FadeR 50, l50
	FadeR 51, l51
	FadeR 52, l52
	FadeR 53, l53
	FadeR 54, l54
	FadeR 55, l55
	FadeR 57, l57
	FadeR 58, l58
	FadeR 59, l59
	FadeR 60, l60
	FadeR 61, l61
	FadeR 62, l62
	FadeR 63, l63
	FadeR 64, l64

End Sub

' div lamp subs

Sub InitLamps()
    Dim x
    For x = 0 to 200
        LampState(x) = 0        ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4      ' used to track the fading state
        FlashSpeedUp(x) = 0.5   ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.07	' slower speed when turning off the flasher
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

' Lights: old method, using 4 images

Sub FadeL(nr, light, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 4:light.image = b:FadingLevel(nr) = 6                   'fading to off...
        Case 5:light.image = a:light.State = 1:FadingLevel(nr) = 1   'ON
        Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1           'wait
        Case 9:light.image = c:FadingLevel(nr) = FadingLevel(nr) + 1 'fading...
        Case 10, 11, 12:FadingLevel(nr) = FadingLevel(nr) + 1        'wait
        Case 13:light.image = d:Light.State = 0:FadingLevel(nr) = 0  'Off
    End Select
End Sub

Sub FadeLm(nr, light, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 4:light.image = b
        Case 5:light.image = a
        Case 9:light.image = c
        Case 13:light.image = d
    End Select
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

' Ramps & Primitives used as 4 step fading lights
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
            Object.IntensityScale = Lumen*FlashLevel(nr)
        Case 5 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
            If FlashLevel(nr) > FlashMax(nr) Then
                FlashLevel(nr) = FlashMax(nr)
                FadingLevel(nr) = 1 'completely on
            End if
            Object.IntensityScale = Lumen*FlashLevel(nr)
    End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it just sets the flashlevel
    Object.IntensityScale = Lumen*FlashLevel(nr)
End Sub

' Desktop Objects: Reels & texts (you may also use lights on the desktop)

' Reels

Sub FadeR(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.SetValue 1:FadingLevel(nr) = 6                   'fading to off...
        Case 5:object.SetValue 0:FadingLevel(nr) = 1                   'ON
        Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1             'wait
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

' *********************************************************************
' 					Wall, rubber and metal hit sounds
' *********************************************************************

Sub Rubbers_Hit(idx):PlaySound "rubber1", 0, Vol(ActiveBall), pan(ActiveBall), 0.25:End Sub

'************************************
'          LEDs Display
'************************************

Dim Digits(28)

Set Digits(0) = da1
Set Digits(1) = da2
Set Digits(2) = da3
Set Digits(3) = da4
Set Digits(4) = da5
Set Digits(5) = da6

Set Digits(6) = db1
Set Digits(7) = db2
Set Digits(8) = db3
Set Digits(9) = db4
Set Digits(10) = db5
Set Digits(11) = db6

Set Digits(12) = dc1
Set Digits(13) = dc2
Set Digits(14) = dc3
Set Digits(15) = dc4
Set Digits(16) = dc5
Set Digits(17) = dc6

Set Digits(18) = dd1
Set Digits(19) = dd2
Set Digits(20) = dd3
Set Digits(21) = dd4
Set Digits(22) = dd5
Set Digits(23) = dd6

Set Digits(24) = de1
Set Digits(25) = de2
Set Digits(26) = de3
Set Digits(27) = de4

Sub UPdateLEDs
	On Error Resume Next
	Dim ChgLED, ii, jj, chg, stat
	ChgLED = Controller.ChangedLEDs(&H0000003f, &Hffffffff)
	If Not IsEmpty(ChgLED) Then
		For ii = 0 To UBound(ChgLED)
			chg = chgLED(ii, 1):stat = chgLED(ii, 2)

			Select Case stat
				Case 0:Digits(chgLED(ii, 0) ).SetValue 0    'empty
				Case 63:Digits(chgLED(ii, 0) ).SetValue 1   '0
				Case 6:Digits(chgLED(ii, 0) ).SetValue 2    '1
				Case 91:Digits(chgLED(ii, 0) ).SetValue 3   '2
				Case 79:Digits(chgLED(ii, 0) ).SetValue 4   '3
				Case 102:Digits(chgLED(ii, 0) ).SetValue 5  '4
				Case 109:Digits(chgLED(ii, 0) ).SetValue 6  '5
				Case 124:Digits(chgLED(ii, 0) ).SetValue 7  '6
				Case 125:Digits(chgLED(ii, 0) ).SetValue 7  '6
				Case 252:Digits(chgLED(ii, 0) ).SetValue 7  '6
				Case 7:Digits(chgLED(ii, 0) ).SetValue 8    '7
				Case 127:Digits(chgLED(ii, 0) ).SetValue 9  '8
				Case 103:Digits(chgLED(ii, 0) ).SetValue 10 '9
				Case 111:Digits(chgLED(ii, 0) ).SetValue 10 '9
				Case 231:Digits(chgLED(ii, 0) ).SetValue 10 '9
				Case 128:Digits(chgLED(ii, 0) ).SetValue 0  'empty
				Case 191:Digits(chgLED(ii, 0) ).SetValue 1  '0
				Case 832:Digits(chgLED(ii, 0) ).SetValue 2  '1
				Case 896:Digits(chgLED(ii, 0) ).SetValue 2  '1
				Case 768:Digits(chgLED(ii, 0) ).SetValue 2  '1
				Case 134:Digits(chgLED(ii, 0) ).SetValue 2  '1
				Case 219:Digits(chgLED(ii, 0) ).SetValue 3  '2
				Case 207:Digits(chgLED(ii, 0) ).SetValue 4  '3
				Case 230:Digits(chgLED(ii, 0) ).SetValue 5  '4
				Case 237:Digits(chgLED(ii, 0) ).SetValue 6  '5
				Case 253:Digits(chgLED(ii, 0) ).SetValue 7  '6
				Case 135:Digits(chgLED(ii, 0) ).SetValue 8  '7
				Case 255:Digits(chgLED(ii, 0) ).SetValue 9  '8
				Case 239:Digits(chgLED(ii, 0) ).SetValue 10 '9
			End Select
		Next
	End IF
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

Const tnob = 2 ' total number of balls
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

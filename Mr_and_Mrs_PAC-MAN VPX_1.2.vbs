'__________________________________________________
'
'
'	Mr. & Mrs. PAC-MAN
'	Bally 1982
'	Version VPX_1.2
'	Shockman 2/3/17
'
'          		---------------------
'       				 --    F6 Dip Switches    --
'          	    ---------------------
'		use to turn on 'balls in play', 'match', and table settings
'__________________________________________________
'
' Thalamus 2018-07-24
' This table doesn't have "Positional Sound Playback Functions" or "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2

Option Explicit

LoadVPM "01500000","bally.vbs",3.1
Sub LoadVPM(VPMver,VBSfile,VBSver)
	On Error Resume Next
		If ScriptEngineMajorVersion<5 Then MsgBox"VB Script Engine 5.0 or higher required"
		ExecuteGlobal GetTextFile(VBSFile)
		If Err Then MsgBox"Unable to open "&VBSfile&". Ensure that it is in the same folder as this table."&vbNewLine&Err.Description:Err.Clear
		Set Controller=CreateObject("B2S.Server")
		If Err Then MsgBox"Can't Load VPinMAME."&vbNewLine&Err.Description
		If VPMver>"" Then If Controller.Version<VPMver Or Err Then MsgBox"VPinMAME ver "&VPMver&" required.":Err.Clear
		If VPinMAMEDriverVer<VBSver Or Err Then MsgBox VBSFile&" ver "&VBSver&" or higher required."
	On Error Goto 0
End Sub

'----------------------------------------------------
' Bind events to solenoids
' Last argument will always be "enabled" (True, False)
'----------------------------------------------------
SolCallback(sBallRelease) = "bsTrough.SolOut"
SolCallback(sKnocker)     = "vpmSolSound ""Knocker"","
SolCallback(sLSaucerL)	  = "bsEject.SolOut"
SolCallback(sLSaucerR)	  = "bsEject.SolOutAlt"
SolCallback(sRSaucer)	  = "bsPacMan.SolOut"
SolCallback(sLSling)      = "vpmSolSound ""Bumper"","
SolCallback(sRSling)      = "vpmSolSound ""Bumper"","
SolCallback(sLJet)        = "vpmSolSound ""Bumper"","
SolCallback(sRJet)        = "vpmSolSound ""Bumper"","
SolCallback(sTargetResetL)= "soldtLDropReset"
SolCallback(sTargetResetR)= "soldtRDropReset"
SolCallback(sTargetReset3)= "soldt3DropReset"
SolCallback(sLReset1)     = "soldtLDrop1"
SolCallback(sLReset2)     = "soldtLDrop2"
SolCallback(sLReset3)     = "soldtLDrop3"
SolCallback(sLReset4)     = "soldtLDrop4"
SolCallback(sRReset1)     = "soldtRDrop1"
SolCallback(sRReset2)     = "soldtRDrop2"
SolCallback(sRReset3)     = "soldtRDrop3"
SolCallback(sRReset4)     = "soldtRDrop4"
SolCallback(sEnable)      = "vpmNudge.SolGameOn"
SolCallback(sGate)		  = "vpmSolDiverter SaveGate,""wire"","
SolCallback(sCLo)		  = "vpmSolDiverter SaveGate,""wire"","

SolCallback(sLLFlipper)   = "vpmSolFlipper LeftFlipper,  Nothing,"
SolCallback(sLRFlipper)   = "vpmSolFlipper RightFlipper, UpperFlipper,"


'--------------------------------
' Init the table, Start VPinMAME
'--------------------------------


Dim bsTrough, dtLDrop, dt3Drop, dtRDrop, bsPacMan, bsEject

ExtraKeyHelp = "Left Ctrl" & vbTab & "(unknown cabinet switch)"

Sub MrMsPacMan_Init

	MrMSPacMan.YieldTime = 1

	Set PlungerIM = New cvpmImpulseP
	PlungerIM.InitImpulseP IMPlunger,66,1.6
	PlungerIM.CreateEvents "plungerIM"
	'PlungerIM.InitExitSnd "", "pinhit"

	Set LampCallback = GetRef("UpdateLamps")
	Set MotorCallback = GetRef("UpdateSolenoids")
	With Controller
		.GameName = cGameName
		.SplashInfoLine = cCredits
		.HandleKeyboard = False

'-----------------------------------------------------
		.hidden=true
		.ShowTitle = False
		.ShowDMDOnly = false
'-----------------------------------------------------

		On Error Resume Next
			.Run
			If Err Then MsgBox Err.Description
		On Error Goto 0
	End With
	' Main Timer init
	PinMAMETimer.Interval = PinMAMEInterval
	PinMAMETimer.Enabled = True
	' Nudging
	vpmNudge.TiltSwitch = 15
	vpmNudge.Sensitivity = .015
	vpmNudge.TiltObj = Array(Bumper1,Bumper2,LeftSlingshot,RightSlingshot)
	' Left Drop targets
	Set dtLDrop = New cvpmDropTarget
	dtLDrop.InitDrop Array(Ghost13,Ghost14,Ghost15,Ghost16),Array(1,2,3,4)
	dtLDrop.InitSnd "",""
		' Right Drop targets
	Set dtRDrop = New cvpmDropTarget
	dtRDrop.InitDrop Array(Ghost17,Ghost18,Ghost19,Ghost20),Array(25,26,27,28)
	dtRDrop.InitSnd "",""
		' Top 3 droptargets
	Set dt3Drop = New cvpmDropTarget
	dt3Drop.InitDrop Array(TopT10,TopT11,TopT12),Array(22,23,24)
	dt3Drop.InitSnd "",""
	' Trough (or really just an outhole)
	Set bsTrough = New cvpmBallStack
	bsTrough.InitNoTrough Feed,5,90,5
	bsTrough.Balls = 1
	' PacMan saucer
	Set bsPacMan = New cvpmBallStack
	bsPacMan.InitSaucer PacManSaucer,8,160,8
	bsPacMan.InitExitSnd "Saucer","Saucer"
	' Top eject saucer (two kickouts)
	Set bsEject = New cvpmBallStack
	bsEject.initSaucer EjectHole,7,260,9
	bsEject.InitAltKick 45,17.4
	bsEject.InitExitSnd "Saucer","Saucer"

	GIt.Enabled=True

	sblmatch.isdropped=True
	sblhighscore.isdropped=False

	Wall302.timerenabled=False
	LightsOff
	Wall302.timerenabled=True

	'Turn off Upper digit display for full screen
	Dim dd
	For Each dd in Upper_Digits
		If ShowDT=False Then
			dd.Visible=0
		Else
			dd.Visible=1
		End If
	Next
End Sub

'==================================================================
' Game Specific code starts here
'==================================================================
Const cGameName     = "m_mpac"   ' PinMAME short name
Const cCredits      = "Mr. & Ms. PacMan Table, Shockman, Script by Gaston - improved by WPCmame, Dip settings menu added by Inkochnito,  btribble-Graphics"
Const UseSolenoids  = 2
Const UseLamps      = False
Const UseGI         = True
' Standard Sounds
Const SSolenoidOn   = ""
Const SSolenoidOff  = ""
Const SFlipperOn    = "FlipperUp"
Const SFlipperOff   = "FlipperDown"
Const SCoin         = "Coin"
bumper1.timerenabled=false
bumper2.timerenabled=false

Sub GIt_Timer
	If ShowDT=True Then
		Light16.state=1
		Light32.state=1
	End If
		LightsOn
	GIt.Enabled=False
End Sub

Dim slider, sxx
slider = (0 + MrMsPacMan.nightday)/4
For Each sxx In GI: sxx.intensity = sxx.intensity - slider: Next

Dim Hit,PlungerIM,domaze,tr,trp,LOn
Dim SoundActive,BallVel,Flying,SoundBall,OnRamp,WireRamp,Panning,BallXPos
'*************************************************************
'Toggle DOF sounds on/off based on cController value
'*************************************************************
Dim ToggleMechSounds
Function SoundFX (sound)
    If cNewController= 4 And ToggleMechSounds = 0 Then
        SoundFX = ""
    Else
        SoundFX = sound
    End If
End Function

Const SoundPanning = 1										'1 enables sound panning for VPinball v9.16 and above
Const TableWidth = 1000                         	        'used for sound panning, enter the value specified in table options "Table Width"
Const PanningFactor = 0.7									'panning factor 0..1, 1 = left speaker off, if ball is rightmost and vice versa, 0 = no panning

Const SoundVel = 4											'sound is played if ball velocity is above this threshold

'Ball-assignment
Sub SoundTrigger1_Hit										'to be placed in the plunger lane, add similiar Triggers to all spots balls will be created (VUKs, holes etc.)
	Set SoundBall = Activeball
	SoundTimer.enabled = 1
	Flying = 0
	OnRamp = 0
	WireRamp = 0
End Sub

'Sound deactivation if ball is destroyed
Sub SoundDrain_Hit()
	SoundTimer.enabled = 0									'this can be added to the normal Drain_Hit procedure, add similiar Triggers to all spots balls will be destroyed (VUKs, holes etc.)
	StopAllSounds                                           'make sure continuus loops are ended
End Sub

Sub SoundTimer_Timer
	If soundPanning = 1 Then
		BallXPos = SoundBall.X
		If BallXPos < 0 Then BallXPos = 0					'boundary check
		If BallXPos > TableWidth Then BallXPos = Tablewidth
		Panning = ((2 * BallXPos / TableWidth) - 1) * PanningFactor		'scale the X position of the ball to the interval [-1,1]
	End If

	If SoundBall.Z > 30 Then	        					'Ball radius is 24, so the ball is definitely above the playfield - must be adjusted, if playfield level is not 0
		If Flying = 0 Then StopAllSounds
		Flying = 1
		If OnRamp = 0 Then Exit Sub							'no sound if "in the air"
	Else
		If Flying = 1 Then 									'Ball is returning to the playfield
			If OnRamp = 0 Then 									'no sound if the ball returns down a ramp
				StopAllSounds
				If soundPanning = 1 Then
					PlaySound "BallCollision2",1,1,Panning,0
				Else
					PlaySound "BallCollision2"
				End If
			End If
		Flying = 0
		OnRamp = 0
		WireRamp = 0
		End If
	End If

	BallVel = SQR((SoundBall.velx^2) + (SoundBall.vely^2))  'calculating the speed vector within the X/Y plane only once per timer call

	If SoundActive = 0 Then									'play new sounds only after the current sound has ended
		If BallVel > SoundVel Then							'if velocity exceeds the threshold a sound will be played
			SoundActive = 1
			If OnRamp = 1 Then
				If WireRamp = 1 Then								'Ball is on a wire ramp
					If soundPanning = 1 Then
						PlaySound "WireRamp1",1,1,Panning,0.25
					Else
						PlaySound "WireRamp1"
					End If
				Else												'Ball is on a normal ramp
					If soundPanning = 1 Then
						PlaySound "Ballroll3",1,1,Panning,0.25
					Else
						PlaySound "Ballroll3"
					End If
				End If
			Else												'Ball is on the playfield, sound will be played as a loop
				If soundPanning = 1 Then
					PlaySound "roll1",-1,1,Panning,0.35
				Else
					PlaySound "roll1",-1,1
				End If
			End If
		End If
		Exit Sub
	Else
		If BallVel < SoundVel Then							'Sound is active and ball velocity drops below the threshold --> stop any sound active
			StopAllSounds										'sound stopped so a new sound can be started
		End If
	End If
End Sub

Sub StopAllSounds()
	SoundActive = 0
	StopSound "WireRamp1"
	StopSound "Ballroll3"
	StopSound "Ballroll2"
	StopSound "roll1"
End Sub

Hit=0:domaze=0

'-------------------
' keyboard handlers
'-------------------
Sub MrMsPacMan_KeyUp(ByVal keycode)
	If keycode = 29 Then Controller.Switch(31) = False
	If vpmKeyUp(keycode) Then
		Exit Sub
	End If
	If keycode = PlungerKey Then
		Plunger1.Fire
		PlungerIM.Fire
		playsound "pinhit"
	End If
End Sub

Sub MrMsPacMan_KeyDown(ByVal keycode)
	If keycode = LeftTiltKey Then
		playsound "wood2"
		Nudge 45, 3
	End If
	If keycode = RightTiltKey Then
		playsound "wood2"
		Nudge 315, 3
	End If
	If keycode = CenterTiltKey Then
		playsound "wood2"
		Nudge 0, 3
	End If
	If keycode = 29 Then Controller.Switch(31) = True
	If vpmKeyDown(keycode) Then
		If keycode=2 Then
			If sblgameover.isdropped=False And light13.state=1 Then
				lightsoff
				wall214.timerenabled=True
				keycode=0
			End If
		End If
		Exit Sub
	End If
 	If keycode = StartGameKey Then
		playsound "coin"
	End If
	If keycode = PlungerKey Then
		Plunger1.Pullback
		PlungerIM.Pullback
		IMPlunger.timerenabled=True
	End If
End Sub

Sub UpdateSolenoids
	Dim Changed, Count, funcName, ii, sel, solNo
	Changed = Controller.ChangedSolenoids
	If Not IsEmpty(Changed) Then
		sel = Controller.Lamp(87)
		Count = UBound(Changed, 1)
		For ii = 0 To Count
			solNo = Changed(ii, CHGNO)
			If SolNo >= 9 And SolNo <= 15 And sel Then solNo = solNo + 24 '9->33 etc
			funcName = SolCallback(solNo)
			If funcName <> "" Then Execute funcName & " CBool(" & Changed(ii, CHGSTATE) &")"
		Next
	End If
End Sub

Dim Digits(37)
Digits(0)=Array(D1,D2,D3,D4,D5,D6,D7,D267)
Digits(1)=Array(D8,D9,D10,D11,D12,D13,D14)
Digits(2)=Array(D15,D16,D17,D18,D19,D20,D21)
Digits(3)=Array(D22,D23,D24,D25,D26,D27,D28,D268)
Digits(4)=Array(D29,D30,D31,D32,D33,D34,D35)
Digits(5)=Array(D36,D37,D38,D39,D40,D41,D42)
Digits(6)=Array(D43,D44,D45,D46,D47,D48,D49)
Digits(7)=Array(D50,D51,D52,D53,D54,D55,D56,D269)
Digits(8)=Array(D57,D58,D59,D60,D61,D62,D63)
Digits(9)=Array(D64,D65,D66,D67,D68,D69,D70)
Digits(10)=Array(D71,D72,D73,D74,D75,D76,D77,D270)
Digits(11)=Array(D78,D79,D80,D81,D82,D83,D84)
Digits(12)=Array(D85,D86,D87,D88,D89,D90,D91)
Digits(13)=Array(D92,D93,D94,D95,D96,D97,D98)
Digits(14)=Array(D99,D100,D101,D102,D103,D104,D105,D271)
Digits(15)=Array(D106,D107,D108,D109,D110,D111,D112)
Digits(16)=Array(D113,D114,D115,D116,D117,D118,D119)
Digits(17)=Array(D120,D121,D122,D123,D124,D125,D126,D272)
Digits(18)=Array(D127,D128,D129,D130,D131,D132,D133)
Digits(19)=Array(D134,D135,D136,D137,D138,D139,D140)
Digits(20)=Array(D141,D142,D143,D144,D145,D146,D147)
Digits(21)=Array(D148,D149,D150,D151,D152,D153,D154,D273)
Digits(22)=Array(D155,D156,D157,D158,D159,D160,D161)
Digits(23)=Array(D162,D163,D164,D165,D166,D167,D168)
Digits(24)=Array(D169,D170,D171,D172,D173,D174,D175,D274)
Digits(25)=Array(D176,D177,D178,D179,D180,D181,D182)
Digits(26)=Array(D183,D184,D185,D186,D187,D188,D189)
Digits(27)=Array(D190,D191,D192,D193,D194,D195,D196)
Digits(28)=Array(D197,D198,D199,D200,D201,D202,D203)
Digits(29)=Array(D204,D205,D206,D207,D208,D209,D210)
Digits(30)=Array(D211,D212,D213,D214,D215,D216,D217)
Digits(31)=Array(D218,D219,D220,D221,D222,D223,D224)
Digits(32)=Array(D225,D226,D227,D228,D229,D230,D231)
Digits(33)=Array(D232,D233,D234,D235,D236,D237,D238)
Digits(34)=Array(D239,D240,D241,D242,D243,D244,D245)
Digits(35)=Array(D246,D247,D248,D249,D250,D251,D252)
Digits(36)=Array(D253,D254,D255,D256,D257,D258,D259)
Digits(37)=Array(D260,D261,D262,D263,D264,D265,D266)

Sub Light42a_Timer
	domaze=0
	LightsOff
	Light42a.TimerEnabled=False
End Sub

Sub wall221_timer
	lightsoff
	wall419.timerenabled=True
	wall221.timerenabled=False
End Sub

Sub wall419_timer
	'lightson
	wall419.timerenabled=False
End Sub

Sub DisplayTimer_Timer
	Dim ChgLED,ii,num,chg,stat,obj
	ChgLED=Controller.ChangedLEDs(&Hffffffff,&Hffffffff)
	If Not IsEmpty(ChgLED) Then
		For ii=0 To UBound(chgLED)
			num=chgLED(ii,0):chg=chgLED(ii,1):stat=chgLED(ii,2)
			For Each obj In Digits(num)
				If chg And 1 Then obj.State=stat And 1
				chg=chg\2:stat=stat\2
			Next
		Next
	End If

	If light42.state=1 Or light26.state=1 Then
		domaze=1
		wall87.timerenabled=True
	End If
	If light71.state=1 Then
		light71r.state=1
	Else
		light71r.state=0
	End If
	If light42.state=1 Then
		light42r.state=1
	Else
		light42r.state=0
	End If
	If light45.state=1 Then
		sblgameover.isdropped=False
	Else
		sblgameover.isdropped=True
	End If
	If light29.state=1 Then
		sblhighscore.isdropped=False
	Else
		sblhighscore.isdropped=True
	End If
	If light59.state=1 Then
		sblballinplay.isdropped=False
	Else
		sblballinplay.isdropped=True
	End If
	If light27.state=1 Then
		sblmatch.isdropped=False
	Else
		sblmatch.isdropped=True
	End If
	If light61.state=1 Then
		sbltilt.isdropped=False
		lightsoff
	Else
		sbltilt.isdropped=True
	End If

End Sub


Dim Matrix(128)

Sub UpdateLamps
	Dim ii, Lamp1, Lamp2, Lamp3, stat1, stat2, chgLamp
	ChgLamp = Controller.ChangedLamps
	If IsEmpty(ChgLamp) Then Exit Sub
	On Error Resume Next
	For ii = 0 To UBound(ChgLamp)
		Lights(ChgLamp(ii, CHGNO)).State = ChgLamp(ii, CHGSTATE)
	Next
	On Error Goto 0
	Light8a.state = Light8.state
	Light9a.state = Light9.state
	Light10a.state = Light10.state
	Light25a.state = Light25.state
	Light41a.state = Light41.state
	Light42a.state = Light42.state
	Light57a.state = Light57.state
	Light71a.state = Light71.state
	For ii = 0 To UBound(ChgLamp)
		If IsObject(Matrix(ChgLamp(ii,CHGNO))) Then
			lamp1 = chgLamp(ii, CHGNO) And 63 : stat1 = Controller.Lamp(lamp1)
			lamp2 = lamp1 + 64 : stat2 = Controller.Lamp(lamp2)
			lamp3 = lamp2 + 8
			If stat1 Then
				If stat2 Then
					Matrix(lamp2).state = 0 : Matrix(lamp1).state = 0 : Matrix(lamp3).State = 1
				Else
					Matrix(lamp3).State = 0 : Matrix(lamp2).state = 0 : Matrix(lamp1).state = 1
				End If
			Else
				If stat2 Then
					Matrix(lamp3).State = 0 : Matrix(lamp1).state = 0 : Matrix(lamp2).state = 1
				Else
					Matrix(lamp3).State = 0 : Matrix(lamp2).state = 0 : Matrix(lamp1).state = 1 : Matrix(lamp1).state = 0
				End If
			End If
		End If
	Next
End Sub

' Kickers
Sub Drain_Hit
	hit=0
	playsound "drain"
	bsTrough.AddBall Me
End Sub

Sub feed_UnHit
	LightsOn
End Sub

Sub PacmanSaucer_Hit
	bsPacMan.AddBall 0
	Hit=1
 	PlaySound"KickerOn"
	if domaze=1 then
		LightsOff
	end if
End Sub

Sub EjectHole_Hit
	Hit=1
 	PlaySound"KickerOn"
	bsEject.AddBall 0
	if domaze=1 then
		LightsOff
	end if
End Sub

Sub PacmanSaucer_UnHit
	Hit=0
	domaze=0
		LightsOn
End Sub

Sub EjectHole_UnHit
	Hit=0
	domaze=0
	LightsOn
End Sub

' Drop targets
'************************************************************************
'***** Primitive Drop Target Animation
'************************************************************************

Sub TopT_timer
	If LOn=1 Then
		If TopT10.isdropped=False Then light10p.state=0 Else light10p.state=1
		If TopT11.isdropped=False Then light11p.state=0 Else light11p.state=1
		If TopT12.isdropped=False Then light12p.state=0 Else light12p.state=1
		If Ghost13.isdropped=False Then light13p.state=0 Else light13p.state=1
		If Ghost14.isdropped=False Then light14p.state=0 Else light14p.state=1
		If Ghost15.isdropped=False Then light15p.state=0 Else light15p.state=1
		If Ghost16.isdropped=False Then light16p.state=0 Else light16p.state=1
		If Ghost17.isdropped=False Then light17p.state=0 Else light17p.state=1
		If Ghost18.isdropped=False Then light18p.state=0 Else light18p.state=1
		If Ghost19.isdropped=False Then light19p.state=0 Else light19p.state=1
		If Ghost20.isdropped=False Then light20p.state=0 Else light20p.state=1
	End If
End Sub

Const DropTgtMovementDir = "TransY"
Const DropTgtMovementMax = 49
Const DropTgtMovementMax2 = 48
Dim primCnt(100), primDir(100)

Sub Ghost13_Hit:   PrimDropTgtDown dtLDrop, 1, 1, Ghost13: End Sub
Sub Ghost13_Timer: PrimDropTgtMove 1, Ghost13, Ghost13p: End Sub
Sub Ghost14_Hit:   PrimDropTgtDown dtLDrop, 2, 2, Ghost14: End Sub
Sub Ghost14_Timer: PrimDropTgtMove 2, Ghost14, Ghost14p: End Sub
Sub Ghost15_Hit:   PrimDropTgtDown dtLDrop, 3, 3, Ghost15: End Sub
Sub Ghost15_Timer: PrimDropTgtMove 3, Ghost15, Ghost15p: End Sub
Sub Ghost16_Hit:   PrimDropTgtDown dtLDrop, 4, 4, Ghost16: End Sub
Sub Ghost16_Timer: PrimDropTgtMove 4, Ghost16, Ghost16p: End Sub

Sub Ghost17_Hit:   PrimDropTgtDown dtRDrop, 1, 25, Ghost17: End Sub
Sub Ghost17_Timer: PrimDropTgtMove 25, Ghost17, Ghost17p: End Sub
Sub Ghost18_Hit:   PrimDropTgtDown dtRDrop, 2, 26, Ghost18: End Sub
Sub Ghost18_Timer: PrimDropTgtMove 26, Ghost18, Ghost18p: End Sub
Sub Ghost19_Hit:   PrimDropTgtDown dtRDrop, 3, 27, Ghost19: End Sub
Sub Ghost19_Timer: PrimDropTgtMove 27, Ghost19, Ghost19p: End Sub
Sub Ghost20_Hit:   PrimDropTgtDown dtRDrop, 4, 28, Ghost20: End Sub
Sub Ghost20_Timer: PrimDropTgtMove 28, Ghost20, Ghost20p: End Sub

Sub TopT10_Hit:   PrimDropTgtDown dt3Drop, 1, 22, TopT10: End Sub
Sub TopT10_Timer: PrimDropTgtMove2 22, TopT10, TopT10p: End Sub
Sub TopT11_Hit:   PrimDropTgtDown dt3Drop, 2, 23, TopT11: End Sub
Sub TopT11_Timer: PrimDropTgtMove2 23, TopT11, TopT11p: End Sub
Sub TopT12_Hit:   PrimDropTgtDown dt3Drop, 3, 24, TopT12: End Sub
Sub TopT12_Timer: PrimDropTgtMove2 24, TopT12, TopT12p: End Sub

Sub soldtLDropReset (enabled)
	If enabled Then
		PrimDropTgtUp dtLDrop, 1, 1, Ghost13, 1
		PrimDropTgtUp dtLDrop, 2, 2, Ghost14, 0
		PrimDropTgtUp dtLDrop, 3, 3, Ghost15, 0
		PrimDropTgtUp dtLDrop, 4, 4, Ghost16, 0
		PlaySound "TargetReset"
	End If
End Sub

Sub soldtRDropReset (enabled)
	If enabled Then
		PrimDropTgtUp dtRDrop, 1, 25, Ghost17, 1
		PrimDropTgtUp dtRDrop, 2, 26, Ghost18, 0
		PrimDropTgtUp dtRDrop, 3, 27, Ghost19, 0
		PrimDropTgtUp dtRDrop, 4, 28, Ghost20, 0
		PlaySound "TargetReset"
	End If
End Sub

Sub soldt3DropReset (enabled)
	If enabled Then
		PrimDropTgtUp dt3Drop, 1, 22, TopT10, 1
		PrimDropTgtUp dt3Drop, 2, 23, TopT11, 0
		PrimDropTgtUp dt3Drop, 3, 24, TopT12, 0
		PlaySound "TargetReset"
	End If
End Sub

Sub soldtLDrop1 (enabled): PrimDropTgtDown dtLDrop, 1, 1, Ghost13: PlaySound "DropTarget":End Sub
Sub soldtLDrop2 (enabled): PrimDropTgtDown dtLDrop, 2, 2, Ghost14: PlaySound "DropTarget":End Sub
Sub soldtLDrop3 (enabled): PrimDropTgtDown dtLDrop, 3, 3, Ghost15: PlaySound "DropTarget":End Sub
Sub soldtLDrop4 (enabled): PrimDropTgtDown dtLDrop, 4, 4, Ghost16: PlaySound "DropTarget":End Sub
Sub soldtRDrop1 (enabled): PrimDropTgtDown dtRDrop, 1, 25, Ghost17: PlaySound "DropTarget":End Sub
Sub soldtRDrop2 (enabled): PrimDropTgtDown dtRDrop, 2, 26, Ghost18: PlaySound "DropTarget":End Sub
Sub soldtRDrop3 (enabled): PrimDropTgtDown dtRDrop, 3, 27, Ghost19: PlaySound "DropTarget":End Sub
Sub soldtRDrop4 (enabled): PrimDropTgtDown dtRDrop, 4, 28, Ghost20: PlaySound "DropTarget":End Sub

Sub PrimDropTgtDown (targetbankname, targetbanknum, swnum, wallName)
	PrimDropTgtAnimate swnum, wallName, 0
	targetbankname.Hit targetbanknum
End Sub

Sub PrimDropTgtUp  (targetbankname, targetbanknum, swnum, wallName, resetvpmbank)
	PrimDropTgtAnimate swnum, wallName, 1
	If resetvpmbank = 1 Then targetbankname.DropSol_On
End Sub

Sub PrimDropTgtMove (swNum, wallName, primName) 'Customize direction as needed
			If primDir(swNum) = 1 Then 'Up
				Select Case primCnt(swNum)
					Case 0: 	primName.TransY = -DropTgtMovementMax * .75
					Case 1: 	primName.TransY = -DropTgtMovementMax * .25
					Case 2,3,4: primName.TransY = 10
					Case 5: 	primName.TransY = 0
					Case Else: 	wallName.TimerEnabled = 0
					PlaySound "TargetReset"
				End Select
			Else 'Down
				Select Case primCnt(swNum)
					Case 0: primName.TransY = -DropTgtMovementMax * .25
					Case 1: primName.TransY = -DropTgtMovementMax * .5
					Case 2: primName.TransY = -DropTgtMovementMax * .75
					Case 3: primName.TransY = -DropTgtMovementMax
					Case Else: wallName.TimerEnabled = 0
					PlaySound "DropTarget"
				End Select
			End If
	primCnt(swnum) = primCnt(swnum) + 1
End Sub

Sub PrimDropTgtMove2 (swNum, wallName, primName) 'Customize direction as needed
			If primDir(swNum) = 1 Then 'Up
				Select Case primCnt(swNum)
					Case 0: 	primName.TransY = -DropTgtMovementMax2 * .75
					Case 1: 	primName.TransY = -DropTgtMovementMax2 * .25
					Case 2,3,4: primName.TransY = 10
					Case 5: 	primName.TransY = 0
					Case else: 	wallName.TimerEnabled = 0
					PlaySound "TargetReset"
				End Select
			Else 'Down
				Select Case primCnt(swNum)
					Case 0: primName.TransY = -DropTgtMovementMax2 * .25
					Case 1: primName.TransY = -DropTgtMovementMax2 * .5
					Case 2: primName.TransY = -DropTgtMovementMax2 * .75
					Case 3: primName.TransY = -DropTgtMovementMax2
					Case Else: wallName.TimerEnabled = 0
					PlaySound "DropTarget"
				End Select
			End If
	primCnt(swnum) = primCnt(swnum) + 1
End Sub

Sub PrimDropTgtAnimate  (swnum, wallName, dir)
	primCnt(swnum) = 0
	primDir(swnum) = dir
	wallName.TimerInterval = 16
	wallName.TimerEnabled = 1
End Sub


' Other targets
Sub Wall22_Hit :Standup4.transx=6:wall22.timerenabled=true: vpmTimer.PulseSwitch 34, 0, 0 : playsound "wood" : End Sub
Sub Wall24_Hit :Standup5.transx=6:wall24.timerenabled=true: vpmTimer.PulseSwitch 35, 0, 0 : playsound "wood" : End Sub
Sub Wall23_Hit :Standup6.transx=6:wall23.timerenabled=true: vpmTimer.PulseSwitch 36, 0, 0 : playsound "wood" : End Sub
Sub Wall25_Hit : Standup1.transx=6:wall25.timerenabled=true:vpmTimer.PulseSwitch 38, 0, 0 : playsound "wood" : End Sub
Sub Wall27_Hit : Standup2.transx=6:wall27.timerenabled=true:vpmTimer.PulseSwitch 39, 0, 0 : playsound "wood" : End Sub
Sub Wall26_Hit : Standup3.transx=6:wall26.timerenabled=true:vpmTimer.PulseSwitch 40, 0, 0 : playsound "wood" : End Sub
Sub ssling1_Hit : vpmTimer.PulseSwitch 29, 0, 0 : End Sub
Sub ssling2_Hit : vpmTimer.PulseSwitch 29, 0, 0 : End Sub
' triggers
Sub LeftInlane_Hit     :  Playsound "sensor" : Controller.Switch(12) = True : End Sub
Sub LeftInlane_UnHit   : Controller.Switch(12) = False :End Sub
Sub RightOutlane_Hit   :  Playsound "sensor" : Controller.Switch(13) = True : End Sub
Sub RightOutlane_UnHit : Controller.Switch(13) = False : End Sub
Sub RightInlane_Hit    : Playsound "sensor" :  Controller.Switch(12) = True :End Sub
Sub RightInlane_UnHit  : Controller.Switch(12) = False :End Sub
Sub LeftOutlane_Hit    : Playsound "sensor" : Controller.Switch(14) = True :  End Sub
Sub LeftOutlane_UnHit  : Controller.Switch(14) = False : End Sub
Sub Trigger1_Hit   : Playsound "sensor" :  Controller.Switch(21) = True : End Sub
Sub Trigger1_UnHit : Controller.Switch(21) = False : End Sub
Sub Wall22_Timer
	Standup4.transx=0
	wall22.timerenabled=false
end sub
Sub Wall24_Timer
	Standup5.transx=0
	wall24.timerenabled=false
end sub
Sub Wall23_Timer
	Standup6.transx=0
	wall23.timerenabled=false
End Sub
Sub Wall25_Timer
	Standup1.transx=0
	wall25.timerenabled=false
end sub
Sub Wall27_Timer
	Standup2.transx=0
	wall27.timerenabled=false
end sub
Sub Wall26_Timer
	Standup3.transx=0
	wall26.timerenabled=false
end sub

' Bumpers/slingshots etc
Sub Bumper2_Hit : vpmTimer.PulseSwitch 17, 0, 0 : End Sub
Sub Bumper1_Hit : vpmTimer.PulseSwitch 18, 0, 0 :End Sub


'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingshot_Slingshot
	vpmTimer.PulseSwitch 20, 0, 0
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingshot.TimerEnabled = 1
End Sub

Sub RightSlingshot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingshot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingshot_Slingshot
    'PlaySound "right_slingshot",0,1,-0.05,0.05
	vpmTimer.PulseSwitch 19, 0, 0
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingshot.TimerEnabled = 1
End Sub

Sub LeftSlingshot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingshot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub SaveGate_Timer
	SaveGatep.RotY=SaveGate.currentangle+90
End Sub

'-------------------------------------
' Map lights into array
'-------------------------------------
' Matrix Yellow
Set Matrix(1)  = Light1
Set Matrix(2)  = Light2
Set Matrix(3)  = Light3
Set Matrix(4)  = Light4
Set Matrix(5)  = Light5
Set Matrix(6)  = Light6

Set Matrix(17) = Light17
Set Matrix(18) = Light18
Set Matrix(19) = Light19
Set Matrix(20) = Light20
Set Matrix(21) = Light21
Set Matrix(22) = Light22

Set Matrix(33) = Light33
Set Matrix(34) = Light34
Set Matrix(35) = Light35
Set Matrix(36) = Light36
Set Matrix(37) = Light37
Set Matrix(38) = Light38

Set Matrix(49) = Light49
Set Matrix(50) = Light50
Set Matrix(51) = Light51
Set Matrix(52) = Light52
Set Matrix(53) = Light53
Set Matrix(54) = Light54
Set Matrix(55) = Light55

' Matrix Red
Set Matrix(65) = Light65
Set Matrix(66) = Light66
Set Matrix(67) = Light67
Set Matrix(68) = Light68
Set Matrix(69) = Light69
Set Matrix(70) = Light70

Set Matrix(81) = Light81
Set Matrix(82) = Light82
Set Matrix(83) = Light83
Set Matrix(84) = Light84
Set Matrix(85) = Light85
Set Matrix(86) = Light86

Set Matrix(97) = Light97
Set Matrix(98) = Light98
Set Matrix(99) = Light99
Set Matrix(100) = Light100
Set Matrix(101) = Light101
Set Matrix(102) = Light102

Set Matrix(113) = Light113
Set Matrix(114) = Light114
Set Matrix(115) = Light115
Set Matrix(116) = Light116
Set Matrix(117) = Light117
Set Matrix(118) = Light118
Set Matrix(119) = Light119

' Matrix orange
Set Matrix(73) = Light73
Set Matrix(74) = Light74
Set Matrix(75) = Light75
Set Matrix(76) = Light76
Set Matrix(77) = Light77
Set Matrix(78) = Light78

Set Matrix(89) = Light89
Set Matrix(90) = Light90
Set Matrix(91) = Light91
Set Matrix(92) = Light92
Set Matrix(93) = Light93
Set Matrix(94) = Light94

Set Matrix(105) = Light105
Set Matrix(106) = Light106
Set Matrix(107) = Light107
Set Matrix(108) = Light108
Set Matrix(109) = Light109
Set Matrix(110) = Light110

Set Matrix(121) = Light121
Set Matrix(122) = Light122
Set Matrix(123) = Light123
Set Matrix(124) = Light124
Set Matrix(125) = Light125
Set Matrix(126) = Light126
Set Matrix(127) = Light127
' Normal lights
Set Lights(7)  = Light7
Set Lights(8)  = Light8
Set Lights(9)  = Light9
Set Lights(10) = Light10
Set Lights(11) = Light11
Set Lights(12) = Light12
Set Lights(13) = Light59
Set Lights(14) = Light14
Set Lights(15) = Light15
Set Lights(23) = Light23
Set Lights(24) = Light24
Set Lights(25) = Light25
Set Lights(26) = Light26
Set Lights(27) = Light27
Set Lights(28) = Light28
Set Lights(29) = Light29
Set Lights(30) = Light30
Set Lights(31) = Light31
Set Lights(39) = Light39
Set Lights(40) = Light40
Set Lights(41) = Light41
Set Lights(42) = Light42
Set Lights(43) = Light43
Set Lights(44) = Light44
Set Lights(45) = Light45
Set Lights(46) = Light46
Set Lights(47) = Light47
Set Lights(56) = Light56
Set Lights(57) = Light57
Set Lights(58) = Light58
Set Lights(59) = Light13
Set Lights(60) = Light60
Set Lights(61) = Light61
Set Lights(62) = Light62
Set Lights(63) = Light63
Set Lights(71) = Light71
Set Lights(103) = Light103

' Constant declarations
' Solenoids
Const sLSaucerL		=  1 '1
Const sLSaucerR		=  2 '2
Const sRSaucer		=  3 '3
Const sKnocker		=  6 '4
Const sBallRelease	=  7 '5
Const sLJet			=  9 '6
Const sRJet			= 10 '7
Const sLSling		= 11 '8
Const sRSling		= 12 '9
Const sTargetResetL = 13 '10
Const sTargetResetR	= 14 '11
Const sTargetReset3	= 15 '12
Const sLReset1	    =  8 '13

Const sLReset2	    = 33 ' 9 Mux (14)
Const sLReset3      = 34 '10 Mux (15)
Const sLReset4      = 35 '11 Mux (16)
Const sRReset1      = 36 '12 Mux (17)
Const sRReset2      = 37 '13 Mux (18)
Const sRReset3      = 38 '14 Mux (19)
Const sRReset4      = 39 '15 Mux (20)

Const sGate			= 17 '21
Const sCLo			= 18 '22
Const sEnable		= 19 '23


'**************************************************************************************************************
'*****GI Lights On

Sub LightsOn()
	dim xx,ml,ref,sl,dx
	LOn=1
	For each ml in Maze_yellow:ml.intensity=50: Next
	For each xx in GI:xx.State = 1: Next
	For each ref in Reflections:ref.IsDropped = 1: Next

	If TopT10.isdropped=true then light10p.state=1
	If TopT11.isdropped=true then light11p.state=1
	If TopT12.isdropped=true then light12p.state=1

	If Ghost13.isdropped=true then light13p.state=1
	If Ghost14.isdropped=true then light14p.state=1
	If Ghost15.isdropped=true then light15p.state=1
	If Ghost16.isdropped=true then light16p.state=1
	If Ghost17.isdropped=true then light17p.state=1
	If Ghost18.isdropped=true then light18p.state=1
	If Ghost19.isdropped=true then light19p.state=1
	If Ghost20.isdropped=true then light20p.state=1
End Sub


'**************************************************************************************************************
'*****GI Lights Off

Sub LightsOff()
	dim xx,ml,ref,sl
	LOn=0
	For each ml in Maze_yellow:ml.intensity=150: Next
	For each xx in GI:xx.State = 0: Next
	For each ref in Reflections:ref.IsDropped = 0: Next
	For each sl in Shadow_lights:sl.State = 0: Next

	Light10p.state=0
	Light11p.state=0
	Light12p.state=0
	Light13p.state=0
	Light14p.state=0
	Light15p.state=0
	Light16p.state=0
	Light17p.state=0
	Light18p.state=0
	Light19p.state=0
	Light20p.state=0
End Sub


'**************************************************************************************************************
'*****Gates

 'Left Gate

Sub GateL_Hit
	playsound"wire"
	vpmTimer.PulseSwitch 30, 50, 0
End Sub

Sub GateL_Timer()
	GateLP.RotZ = GateL.CurrentAngle/3
End Sub


 'Right Gate

Sub GateR_Hit
	playsound"wire"
	vpmTimer.PulseSwitch 30, 50, 0
End Sub

Sub GateR_Timer()
	GateRP.RotZ = GateR.CurrentAngle/3
End Sub

 'Right Gate

Sub Gate1_Hit
 playsound "wire"
 End Sub


 'Top Gate

Sub Gate_Hit
 playsound "wire"
End Sub

Sub Gate_Timer()
	GateTP.RotZ = Gate.CurrentAngle/3
End Sub

Sub Trigger6_hit
	'domaze=0
	If Light88.State=1 Then
		lightson
	End If
	playsound "plunger"
End Sub

Sub wall229_hit
	playsound "pins"
end sub


'added by Inkochnito
Sub editDips
	Dim vpmDips : Set vpmDips = New cvpmDips
	With vpmDips
		.AddForm  700,400,"Mr. & Mrs. PAC-MAN - DIP switches - by Inkochnito"
		.AddChk   7,10,180,Array("Match feature", &H08000000)'dip 28
		.AddChk   205,10,115,Array("Credits display", &H04000000)'dip 27
		.AddFrame   2, 30,  190,"Maximum credits",&H03000000,Array("10 credits",0,"15 credits", &H01000000,"25 credits", &H02000000,"40 credits", &H03000000)'dip 25&26
		.AddFrame   2, 106, 190,"Ball in Maze-saucer kickout",&H00000020,Array("kicks out before all moves are used",0,"does not kick till all moves are used",&H00000020)'dip 6
		.AddFrame   2, 152, 190,"Drop target 20000 yellow arrow",&H00000040,Array("20000 Arrow starts off",0,"20000 Arrow starts on",&H00000040)'dip 7
		.AddFrame   2, 198, 190,"Game over attract",&H20000000,Array("Off",0,"On",&H20000000)'dip 30
		.AddFrame	2, 248, 190,"Pacman lites at start of the game",&H00002000,Array("3 Pacman lites",0,"4 Pacman lites",&H00002000)'dip 14
		.AddFrame	2, 298, 190,"Pacman aggressive lite",49152,Array ("goes out after all moves are used",0,"goes out when monster dies",&H00004000,"stays on entire ball",49152)'dip 15&16
		.AddFrame	205, 30, 190,"Balls per game *",&HC0000000,Array ("2 balls",&HC0000000,"3 balls",0,"4 balls",&H80000000,"5 balls",&H40000000)'dip 31&32
		.AddFrame   205, 106, 190,"Time to beat",&H00700000,Array("10",0,"15",&H00100000,"20",&H00200000,"25",&H00300000,"30",&H00400000,"35",&H00500000,"40",&H00600000,"45",&H00700000)'dip 21&22&23
		.AddFrame   205, 248, 190,"Extra ball lite on after",&H00800000,Array("going into saucer 10 times",0,"going into saucer 5 times",&H00800000)'dip 24
		.AddFrame   205, 298, 190,"Number of replays per game",&H10000000,Array("Only 1 replay per player per game",0,"All replays earned will be collected",&H10000000)'dip 29
		.AddLabel	30, 390, 350, 20,"After hitting OK, press F3 to reset game with new settings."
		.ViewDips
	End With
End Sub
Set vpmShowDips = GetRef("editDips")

Sub Rubber_Hit(idx):PlaySound "Rubber2":End Sub

Sub LeftFlipper_Collide(parm):PlaySound "Rubber2":End Sub
Sub RightFlipper_Collide(parm):PlaySound "Rubber2":End Sub
Sub UpperFlipper_Collide(parm):PlaySound "Rubber2":End Sub

Sub Wall214_Timer()
lightson
	wall214.timerenabled=false
End Sub

Sub Wall87_Timer()
domaze=0
wall87.timerenabled=false
End Sub

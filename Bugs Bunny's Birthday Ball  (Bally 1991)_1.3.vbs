'______                  ______                         _
'| ___ \                 | ___ \                       ( )
'| |_/ /_   _  __ _ ___  | |_/ /_   _ _ __  _ __  _   _|/ ___
'| ___ \ | | |/ _` / __| | ___ \ | | | '_ \| '_ \| | | | / __|
'| |_/ / |_| | (_| \__ \ | |_/ / |_| | | | | | | | |_| | \__ \
'\____/ \__,_|\__, |___/ \____/ \__,_|_| |_|_| |_|\__, | |___/
'              __/ |                               __/ |
'             |___/                               |___/
'______ _      _   _         _              ______       _ _
'| ___ (_)    | | | |       | |             | ___ \     | | |
'| |_/ /_ _ __| |_| |__   __| | __ _ _   _  | |_/ / __ _| | |
'| ___ \ | '__| __| '_ \ / _` |/ _` | | | | | ___ \/ _` | | |
'| |_/ / | |  | |_| | | | (_| | (_| | |_| | | |_/ / (_| | | |
'\____/|_|_|   \__|_| |_|\__,_|\__,_|\__, | \____/ \__,_|_|_|
'                                     __/ |
'                                    |___/
'______       _ _          __  __   _____  _____  __   __
'| ___ \     | | |        / / /  | |  _  ||  _  |/  |  \ \ 
'| |_/ / __ _| | |_   _  | |  `| | | |_| || |_| |`| |   | |
'| ___ \/ _` | | | | | | | |   | | \____ |\____ | | |   | |
'| |_/ / (_| | | | |_| | | |  _| |_.___/ /.___/ /_| |_  | |
'\____/ \__,_|_|_|\__, | | |  \___/\____/ \____/ \___/  | |
'                  __/ |  \_\                          /_/
'                 |___/
'
'Bugs Bunny's Birthday Ball (Bally 1991) Revision 1.3 for VP10
'Intial table design and scripting by "Bodydump" and completed by "wrd1972"
'Plastics primtives by "Cyberpez"
'Additional scripting by "cyberpez" and "Rothbauerw"
'Clear ramp primitives and sidewalls by "Flupper"
'Lighting by "wrd1972"
'Flashers by "wrd1972"
'GI flashing by "clarkkent"
'HMLF physics by "wrd1972"
'Flipper physics by "wrd1972" and "Rothbauerw"
'Looney Tunes lighting by "nfozzy"
'Domes and Henhouse primitives by "Zany"
'DT table scoring reels and DT backdrop by "32 Assassin"
'DOF and Controller by "Arngrim"
'Very special thanks to Bodydump for allowing me to complete his table
'Thanks to Alan for resolving the Captive Ball issues.
'Many thanks to countless others for in the VPF community for helping me learn the table creation process.

'Note: Due to the unique nature of this table and its lower playfield, playfield reflections are not possible.

Option Explicit
Randomize

' Thalamus 2018-07-23
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' No special SSF tweaks yet.

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="bbnny_l2",UseSolenoids=1,UseLamps=0,UseGI=0,SSolenoidOn="Solenoid",SSolenoidOff="SolOff", SCoin="Coin"
Const ballsize = 50
Const ballmass = 1.7

LoadVPM "01560000", "S11.VBS", 3.26
Dim ShowFSrails, DesktopMode: DesktopMode = Table1.ShowDT

If DesktopMode = True Then 'Show Desktop components
Railleft.visible = 1
Railright.visible = 1
Lockdownbar. visible =1
Sidewalls_DT. visible =1
Sidewalls_FS. visible =0
Ramp1_DT. visible = 1
Ramp2_DT. visible = 1
Ramp3_DT. visible = 1
Ramp1_FS. visible = 0
Ramp2_FS. visible = 0
Ramp3_FS. visible = 0



Else
Railleft.visible = 0
Railright.visible = 0
Lockdownbar. visible =0
Sidewalls_DT. visible =0
Sidewalls_FS. visible =1
Ramp1_DT. visible = 0
Ramp2_DT. visible = 0
Ramp3_DT. visible = 0
Ramp1_FS. visible = 1
Ramp2_FS. visible = 1
Ramp3_FS. visible = 1
End if


'******************************************************************************************************************************************************
'Solenoid Call backs
'******************************************************************************************************************************************************
	SolCallback(1) = "bsTrough.SolIn"
	SolCallback(2) = "bsTrough.SolOut"
	SolCallBack(5) = "bsHoleKick.SolOut"
	SolCallBack(6) = "dtBombBank.SolDropUp" 'Drop Targets
	SolCallBack(7) = "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
	SolCallBack(13) = "TopKick"  'ball launcher
	SolCallback(14) = "SolKickBack"  'Left outlane kickout

'GI
	SolCallback(10) = "UpdateGI" 'PF GI relay
	'SolCallback(11) =  'Insert Backbox GI relay

'Flashers
	'SolCallBack(8) =  					'Right BackPannel Flash
	SolCallBack(9)  = "Flashlooney" 	'Looney Relay
	SolCallBack(16) = "Flashtunes" 	'Tunes relay
	SolCallback(25) = "SetLamp 125," 	'Left Ramp flasher
	SolCallback(26) = "SetLamp 126,"    'Left Stand up Flasher
	SolCallBack(27) = "SetLamp 127,"  	'Millions Light
	SolCallback(28) = "SetLamp 128," 	'Taz Flasher
	SolCallback(29) = "SetLamp 129," 	'Right Standup flasher
	SolCallBack(30) = "SetLamp 130," 	'Bugs light
	SolCallBack(31) = "SetLamp 131,"    'Back Left
	SolCallBack(32) = "SetLamp 132," 	'Back Right

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFX("fx_FlipperupL",DOFContactors):LeftFlipper.RotateToEnd
     Else
         PlaySound SoundFX("fx_Flipperdown",DOFContactors):LeftFlipper.RotateToStart
     End If
End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFX("fx_FlipperupR",DOFContactors):RightFlipper.RotateToEnd:RightFlipper1.RotateToEnd
     Else
         PlaySound SoundFX("fx_Flipperdown",DOFContactors):RightFlipper.RotateToStart:RightFlipper1.RotateToStart
     End If
End Sub

dim defaultEOS,EOSAngle,EOSTorque
defaulteos = leftflipper.eostorque
EOSAngle = 3
EOSTorque = .9

'Primitve Flippers
Sub FlippersTimer_Timer()
	flipperbatright1.RotAndTra8 = RightFlipper1.CurrentAngle + 90
    Flipperbatleft.RotAndTra8 = LeftFlipper.CurrentAngle - 90
    Flipperbatright.RotAndTra8 = RightFlipper.CurrentAngle + 90

	If LeftFlipper.CurrentAngle < LeftFlipper.EndAngle + EOSAngle Then
		LeftFlipper.eostorque = EOSTorque
	Else
		LeftFlipper.eostorque = defaultEOS
	End If

	If RightFlipper.CurrentAngle > RightFlipper.EndAngle - EOSAngle Then
		RightFlipper.eostorque = EOSTorque
	Else
		RightFlipper.eostorque = defaultEOS
	End If

	If RightFlipper1.CurrentAngle < RightFlipper1.EndAngle + EOSAngle Then
		RightFlipper1.eostorque = EOSTorque
	Else
		RightFlipper1.eostorque = defaultEOS
	End If

End Sub



'**********************************************************************************************************

'Solenoid Controlled toys
'**********************************************************************************************************

Sub SolKickBack(Enabled)
	If Enabled Then
		PlungerIM.AutoFire
	End If
End Sub

Sub TopKick(Enabled)
	If Enabled Then
		Kicker1.Enabled = True
	Else
		Kicker1.Enabled = False
	End If
End Sub


'Flipper Fix (Script by Uncle Willy)
dim GameOnFF
GameOnFF = 0

Sub FFTimer_Timer()
	If bsTrough.Balls > 0 or bumper1b.Force = 0 Then
		GameOnFF = 0
	else
		GameOnFF = 1
	end if
End Sub



'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************
	Dim bsTrough, bsTopKick, bsHoleKick, dtBombBank, xx, PlungerIM

  Sub Table1_Init
	vpmInit Me
	On Error Resume Next
		With Controller
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
		.SplashInfoLine = "Bugs Bunny Birthday Ball"&chr(13)&""
		.HandleMechanics=0
		.HandleKeyboard=0
		.ShowDMDOnly=1
		.ShowFrame=0
		.ShowTitle=0
        .hidden = 1
      On Error Resume Next

  	  Controller.Run
  	If Err Then MsgBox Err.Description
      On Error Goto 0
      End With

	PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

	'Nudging
    vpmNudge.TiltSwitch= 1
    vpmNudge.Sensitivity=5
    vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingshot)

  '**Trough
      Set bsTrough = New cvpmBallStack
      With bsTrough
          .InitSw 10, 11, 12, 0, 0, 0, 0, 0
          .InitKick BallRelease, 00,12
          .InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("solon",DOFContactors)
          .Balls = 2
      End With

	Set bsHoleKick = New cvpmSaucer
		bsHoleKick.InitKicker Kicker3,16,180,12,0.5
		bsHoleKick.InitSounds "",SoundFX("solon",DOFContactors),SoundFX("Kicker_Release",DOFContactors)

 	'DropTargets
	Set dtBombBank = New cvpmDropTarget
		dtBombBank.InitDrop Array(sw20, sw21, sw22), Array(20,21,22)
		dtBombBank.InitSnd SoundFX("Droptarget",DOFContactors),SoundFX("DTReset",DOFContactors)

'Auto kickback
    ' Impulse Plunger used for kickback
    Const IMPowerSetting = 41 'Plunger Power
    Const IMTime = 0.3        ' Time in seconds for Full Plunge
    Set plungerIM = New cvpmImpulseP
        plungerIM.InitImpulseP sw51, IMPowerSetting, IMTime
        plungerIM.Random 0.3
        plungerIM.switch 51
        plungerIM.InitExitSnd SoundFX("plunger2",DOFContactors), SoundFX("plunger",DOFContactors)
        plungerIM.CreateEvents "plungerIM"

  'Captive Ball
	Kicker2.CreateBall
	Kicker2.Kick 0,3
	Kicker2.enabled = 0

  'GI Init
	GIOn
 End Sub




sub tb_timer	'debug
	me.text = LeftFlipper.EOStorque & " ... " & RightFlipper.EosTorque & vbnewline & RightFlipper1.EosTorque
'			& (LeftFlipper.Strength * LeftFlipper EosTorque) & " ... " & (RightFlipper.Strength * RightFlipper.EosTorque)
End Sub
'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************
Sub Destroyer_Hit:Me.destroyball:End Sub'debug
dim t1, t2 : t1 = 180 : t2 = 18

Sub Table1_KeyDown(ByVal KeyCode)
'	if keycode = 31 then Kicker4.CreateSizedBallWithMass Ballsize/2, BallMass : Kicker4.Kick t1, t2
	If KeyDownHandler(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Pullback:playsound"plungerpull"

	'* Test Kicker
	If keycode = 37 Then TestKick ' K
	If keycode = 19 Then return_to_test ' R return ball to kicker
	If keycode = 46 Then create_testball ' C create ball ball in test kicker
	If keycode = 205 Then TKickAngle = TKickAngle + 3:fKickDirection.Visible=1:fKickDirection.RotZ=TKickAngle'+90 ' right arrow
	If keycode = 203 Then TKickAngle = TKickAngle - 3:fKickDirection.Visible=1:fKickDirection.RotZ=TKickAngle'+90 'left arrow
	If keycode = 200 Then TKickPower = TKickPower + 2:debug.print "TKickPower: "&TKickPower ' up arrow
	If keycode = 208 Then TKickPower = TKickPower - 2:debug.print "TKickPower: "&TKickPower ' down arrow

End Sub

Sub Table1_KeyUp(ByVal KeyCode)
	If KeyUpHandler(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Fire:PlaySound"plunger"

	'* Test Kicker
	If keycode = 205 Then fKickDirection.Visible=0 ' right arrow
	If keycode = 203 Then fKickDirection.Visible=0 'left arrow

End Sub

'******************************************************
'					Test Kicker
'******************************************************

Dim TKickAngle, TKickPower, TKickBall
TKickAngle = 0
TKickPower = 10

Sub testkick()
	test.kick TKickAngle,TKickPower
End Sub

Sub create_testball():Set TKickBall = test.CreateBall:End Sub
Sub test_hit():Set TKickBall=ActiveBall:End Sub
Sub return_to_test():TKickBall.velx=0:TKickBall.vely=0:TKickBall.x=test.x:TKickBall.y=test.y-50:test.timerenabled=0:End Sub

'**********************************************************************************************************

'Drains and Kickers
Sub Drain_Hit:Playsound "Drain2":bsTrough.AddBall Me:End Sub
Sub Kicker3_hit:PlaySound "scoopenter":bsHoleKick.AddBall me:End Sub

' LaserKick
Dim TKStep
Sub Kicker1_Hit:PlaySound "kicker_enter_center":Me.Kick 185,50:PlaySound SoundFX("Kicker_Release",DOFContactors):TKStep = 0:pTopKick.TransY = 50:Me.TimerEnabled = true:End Sub
Sub Kicker1_timer()
	Select Case TKStep
		Case 0:pTopKick.TransY = 35
		Case 1:pTopKick.TransY = 18
		Case 2:pTopKick.TransY = 0:me.TimerEnabled = false
	End Select
	TKStep = TKStep + 1
End Sub

'Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 54 : playsound SoundFX("fx_bumper1",DOFContactors): End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 52 : playsound SoundFX("fx_bumper2",DOFContactors): End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 53 : playsound SoundFX("fx_bumper3",DOFContactors): End Sub



'Wire Triggers
Sub sw14_Hit:Controller.Switch(14) = 1:PlaySound "rollover":End Sub
Sub sw14_Unhit:Controller.Switch(14) = 0:End Sub
sub sw20_hit: dtBombBank.hit 1: Playsound "droptarget": end sub
sub sw21_hit: dtBombBank.hit 2: Playsound "droptarget": end sub
sub sw22_hit: dtBombBank.hit 3: Playsound "droptarget": end sub
Sub sw36_Hit:Controller.Switch(36) = 1:PlaySound "rollover":End Sub
Sub sw36_Unhit:Controller.Switch(36) = 0:End Sub
Sub sw37_Hit:Controller.Switch(37) = 1:PlaySound "rollover":End Sub
Sub sw37_Unhit:Controller.Switch(37) = 0:End Sub
Sub sw38_Hit:Controller.Switch(38) = 1:PlaySound "rollover":End Sub
Sub sw38_Unhit:Controller.Switch(38) = 0:End Sub
Sub sw39_Hit:Controller.Switch(39) = 1:PlaySound "rollover":End Sub
Sub sw39_Unhit:Controller.Switch(39) = 0:End Sub
Sub sw40_Hit:Controller.Switch(40) = 1:PlaySound "rollover":End Sub
Sub sw40_Unhit:Controller.Switch(40) = 0:End Sub
Sub sw46_Hit:Controller.Switch(46) = 1:PlaySound "rollover":End Sub
Sub sw46_Unhit:Controller.Switch(46) = 0:End Sub
Sub sw47_Hit:Controller.Switch(47) = 1:PlaySound "rollover":End Sub
Sub sw47_Unhit:Controller.Switch(47) = 0:End Sub
Sub sw48_Hit:Controller.Switch(48) = 1:PlaySound "rollover":End Sub
Sub sw48_Unhit:Controller.Switch(48) = 0:End Sub
Sub sw51_Hit:Controller.Switch(15) = 1:PlaySound "rollover":End Sub
Sub sw51_Unhit:Controller.Switch(15) = 0:End Sub

'Stand-Ups
Sub sw17_Hit:vpmTimer.PulseSw 17:End Sub
Sub sw19_Hit:vpmTimer.PulseSw 19:End Sub
Sub sw24_Hit:vpmTimer.PulseSw 24:End Sub
Sub sw25_Hit:vpmTimer.PulseSw 25:End Sub
Sub sw26_Hit:vpmTimer.PulseSw 26:End Sub
Sub sw27_Hit:vpmTimer.PulseSw 27:End Sub
Sub sw28_Hit:vpmTimer.PulseSw 28:End Sub
Sub sw29_Hit:vpmTimer.PulseSw 29:End Sub
Sub sw30_Hit:vpmTimer.PulseSw 30:End Sub
Sub sw31_Hit:vpmTimer.PulseSw 31:End Sub
Sub sw32_Hit:vpmTimer.PulseSw 32:End Sub
Sub sw33_Hit:vpmTimer.PulseSw 33:End Sub
Sub sw34_Hit:vpmTimer.PulseSw 34:End Sub
Sub sw35_Hit:vpmTimer.PulseSw 35:End Sub

'Upper PF Stand Up Targets
Sub sw44_Hit:vpmTimer.PulseSw 44:End Sub
Sub sw45_Hit:vpmTimer.PulseSw 45:End Sub
Sub sw60_Hit:vpmTimer.PulseSw 60:End Sub
Sub sw61_Hit:vpmTimer.PulseSw 61:End Sub
Sub sw62_Hit:vpmTimer.PulseSw 62:End Sub

'Scoring Rubber
''Sub sw50_Hit:vpmTimer.PulseSw 50:End Sub

'Hidden Ramp Trigger
Sub sw42_Hit:Controller.Switch(42) = 1:PlaySound "rollover":End Sub
Sub sw42_UnHit:Controller.Switch(42) = 0:End Sub

'Gate Triggers
Sub sw15_Hit: vpmTimer.PulseSw(15) : End Sub
Sub sw41_Hit: vpmTimer.PulseSw(41) : End Sub

'Spinners
Sub Spinner1_Spin:vpmTimer.PulseSw 23 : playsound"fx_spinner" : End Sub

'TopKick  -cp
Sub Sw18_Hit()Controller.Switch(18)=1:sw18p.TransY = -5:End Sub
Sub Sw18_UnHit()Controller.Switch(18)=0:sw18p.TransY = 0:End Sub

'Generic ramp sounds
Sub Trigger1_Hit: playsound"ball drop" : End Sub
Sub Trigger5_Hit: playsound"ball drop" : End Sub
Sub Trigger3_Hit: playsound"ball drop" : End Sub
Sub Trigger2_Hit: playsound"ball drop" : End Sub
Sub Ramp1_Hit:Playsound "metalhit_medium":End Sub
Sub Wall112_Hit:Playsound "metalhit_medium":End Sub
Sub Primitive29_Hit:Playsound "Drop ball":End Sub


'**************************************************************
'**************************************************************
'Spiral Ramp helper and Brakes
'**************************************************************
'**************************************************************

Sub ramphelper_Hit()
	ActiveBall.velx = Activeball.velx*1.4
End Sub



'**********************************************************************************************************
'Update GI Lights
'**********************************************************************************************************

'GI On
Sub GIOff
	For each xx in GI_lighting:xx.State = 0: Next
	sw17.image = "targetgreenoff"
	sw25.image = "targetwhiteoff"
	sw26.image = "targetyellowoff"
	sw27.image = "targetorangedark"
	sw28.image = "targetredoff"
	sw29.image = "targetgreenoff"
	sw30.image = "targetblueoff"
	sw31.image = "targetgreenoff"
	sw32.image = "targetredoff"
	sw33.image = "targetorangedark"
	sw34.image = "targetyellowoff"
	sw35.image = "targetwhiteoff"
	sw44.image = "targetgreenoff"
	sw45.image = "targetgreenoff"
	sw60.image = "targetgreenoff"
	sw61.image = "targetgreenoff"
	sw62.image = "targetgreenoff"
	looney_dome.image = "Looney_dome_off1"
	tunes_dome.image = "Tunes_dome_off1"
	bugssign.image = "bb_top_plasticoff"
    pHenHouse.Image = "bb_hen_house_off"

End Sub

Sub GIOn
	For each xx in GI_lighting:xx.State = 1: Next

	Tunes_dome.image = "bb_plastic_LLon"
	Looney_dome.image = "bb_plastic_UL"
	bugssign.image = "bb_top_plastic"
	pHenHouse.Image = "bb_hen_house_on"

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
LampTimer.Interval = 40 'lamp fading speed
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

	nFadeL 1, L1
	nFadeLm 2, Light48
	nFadeLm 2, Light48a
	nFadeL 2, Light45
	NFadeLm 3, Light49
	NFadeLm 3, Light49a
	NFadeL 3, Light46
	NFadeLm 4, Light50
	NFadeLm 4, Light50a
	NFadeL 4, Light36
	Flash 5, Flasher6
	Flash 6, Flasher5
	Flash 7, Flasher4
	'Flash 8, Flasher3
	Flash 9, Flasher9
	Flash 10, Flasher10
	Flash 11, Flasher11
	Flash 12, Flasher12
	Flash 13, Flasher13
	nFadeL 14, L14
	nFadeL 15, L15
	nFadeL 16, L16
	Flash 17, F17
	Flash 18, F18
	Flash 19, F19
	nFadeL 20, L20
	nFadeL 21, L21
	nFadeL 22, L22
	nFadeL 23, L23
	nFadeL 24, L24

'Tunes plastic dome
	Flashm 25, L25
	Flashm 25, L25_bloom
	Flash 25, Tunes_T
	Flashm 26, L26
	Flashm 26, L26_bloom
	Flash 26, Tunes_U
	Flashm 27, L27
	Flashm 27, L27_bloom
	Flash 27, Tunes_N
	Flashm 28, L28
	Flashm 28, L28_bloom
	Flash 28, Tunes_E
	Flashm 29, L29
	Flashm 29, L29_bloom
	Flash 29, Tunes_s















Looney_L.rotx = -37.36	'Looney L
Looney_L.roty = 77.8
Looney_L.rotz = -52.64

Looney_L.x = 285.45
Looney_L.height = 292
Looney_L.y = 1177

Looney_O.rotx = -37.36	'Looney O
Looney_O.roty = 77.8
Looney_O.rotz = -52.64

Looney_O.x = 285.45
Looney_O.height = 292
Looney_O.y = 1177

Looney_2ndO.rotx = -37.36	'Looney 2NDO
Looney_2ndO.roty = 77.8
Looney_2ndO.rotz = -52.64

Looney_2ndO.x = 285.45
Looney_2ndO.height = 292
Looney_2ndO.y = 1177



Looney_N.x = 387		'Looney N
Looney_N.y = 770
Looney_N.height = 292

Looney_N.rotx = 0
Looney_N.rotY = 80
Looney_N.rotz = -90

Looney_E.x = 387		'Looney E
Looney_E.y = 770
Looney_E.height = 292

Looney_E.rotx = 0
Looney_E.rotY = 80
Looney_E.rotz = -90

Looney_Y.x = 387		'Looney Y
Looney_Y.y = 770
Looney_Y.height = 292

Looney_Y.rotx = 0
Looney_Y.rotY = 80
Looney_Y.rotz = -90



Tunes_T.x = 789.01	'Tunes_T
Tunes_T.y = 1233
Tunes_T.height = 287

Tunes_T.rotx = 0
Tunes_T.rotY = -79.921
Tunes_T.rotz = 90

Tunes_U.x = 789.01	'Tunes_U
Tunes_U.y = 1233
Tunes_U.height = 287

Tunes_U.rotx = 0
Tunes_U.rotY = -79.921
Tunes_U.rotz = 90

Tunes_N.x = 789.01	'Tunes_N
Tunes_N.y = 1233
Tunes_N.height = 287

Tunes_N.rotx = 0
Tunes_N.rotY = -79.921
Tunes_N.rotz = 90

Tunes_E.x = 789.01	'Tunes_E
Tunes_E.y = 1233
Tunes_E.height = 287

Tunes_E.rotx = 0
Tunes_E.rotY = -79.921
Tunes_E.rotz = 90

Tunes_S.x = 789.01	'Tunes_S
Tunes_S.y = 1233
Tunes_S.height = 287

Tunes_S.rotx = 0
Tunes_S.rotY = -79.921
Tunes_S.rotz = 90








'Looney plastic dome
	Flashm 30, L30
	Flashm 30, L30_bloom
	Flash 30, Looney_L
	Flashm 31, L31
	Flashm 31, L31_bloom
	Flash 31, Looney_O
	Flashm 32, L32
	Flashm 32, L32_bloom
	Flash 32, Looney_2ndO
	Flashm 33, L33
	Flashm 33, L33_bloom
	Flash 33, Looney_N
	Flashm 34, L34
	Flashm 34, L34_bloom
	Flash 34, Looney_E
	Flashm 35, L35
	Flashm 35, L35_bloom
	Flash 35, Looney_Y


	nFadeL 36, L36
	nFadeL 37, L37
	nFadeL 38, L38
	nFadeL 39, L39
	nFadeL 40, L40

	nFadeLm 52, L52
	nFadeLm 52, Light57
	nFadeL 52, Light58


	nFadeLm 41, L41
	nFadeLm 41, light53
	nFadeLm 41, light54
	nFadeLm 41, light59
	nFadeL 41, light60
	nFadeLm 42, L42
	nFadeLm 42, Light55
	nFadeL 42, Light56

	nFadeL 43, L43
	nFadeL 44, L44
	nFadeL 45, L45
	nFadeL 46, L46
	nFadeL 47, L47
	nFadeL 48, L48
	nFadeL 49, L49
	nFadeL 50, L50
	nFadeL 51, L51
	'nFadeL 52, L52
	Flashm 53, F53a
	Flash 53, F53b
	Flash 54, F54
	nFadeL 63, L63
	Flash 64, F64a


'Solenoid Controlled Lights and Flahsers
'Soloenoid 09
	nFadeLm 109, Looney0
	nFadeLm 109, Looney1
	nFadeLm 109, Looney2
	nFadeLm 109, Looney3
	nFadeLm 109, Looney4
	nFadeL 109, Looney5
'Soloenoid 16

	nFadeLm 116, Tunes0
	nFadeLm 116, Tunes1
	nFadeLm 116, Tunes2
	nFadeLm 116, Tunes3
	nFadeL 116, Tunes4
'Soloenoid 25 Sylvester Flasher

	Flashm 125, Flasher25c
	Flashm 125, Flasher25b
	Flashm 125, Flasher25a
	Flash 125, Flasher25
'Soloenoid 26

	Flashm 126, Flasher26
	Flashm 126, Flasher26a
	Flash 126, Flasher26b

'Soloenoid27
nFadeL 127, L199

'Soloenoid 28 Taz Flasher
	Flashm 128, Flasher28
	Flash 128, Flasher28a

'Soloenoid 09 Bomb Flasher
	Flashm 129, Flasher29
	Flashm 129, Flasher29a
	Flash 129, Flasher29b

'Soloenoid 30
'    Flash 130, L200
'Soloenoid 31
	Flash 131, Flasher7C
'Soloenoid 32
	Flash 132, Flasher8C

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



'**********************************************************************************************************
'Digital Display
'**********************************************************************************************************
 Dim Digits(32)
 Digits(0)=Array(a00, a05, a0c, a0d, a08, a01, a06, a0f, a02, a03, a04, a07, a0b, a0a, a09, a0e)
 Digits(1)=Array(a10, a15, a1c, a1d, a18, a11, a16, a1f, a12, a13, a14, a17, a1b, a1a, a19, a1e)
 Digits(2)=Array(a20, a25, a2c, a2d, a28, a21, a26, a2f, a22, a23, a24, a27, a2b, a2a, a29, a2e)
 Digits(3)=Array(a30, a35, a3c, a3d, a38, a31, a36, a3f, a32, a33, a34, a37, a3b, a3a, a39, a3e)
 Digits(4)=Array(a40, a45, a4c, a4d, a48, a41, a46, a4f, a42, a43, a44, a47, a4b, a4a, a49, a4e)
 Digits(5)=Array(a50, a55, a5c, a5d, a58, a51, a56, a5f, a52, a53, a54, a57, a5b, a5a, a59, a5e)
 Digits(6)=Array(a60, a65, a6c, a6d, a68, a61, a66, a6f, a62, a63, a64, a67, a6b, a6a, a69, a6e)
 Digits(7)=Array(a70, a75, a7c, a7d, a78, a71, a76, a7f, a72, a73, a74, a77, a7b, a7a, a79, a7e)
 Digits(8)=Array(a80, a85, a8c, a8d, a88, a81, a86, a8f, a82, a83, a84, a87, a8b, a8a, a89, a8e)
 Digits(9)=Array(a90, a95, a9c, a9d, a98, a91, a96, a9f, a92, a93, a94, a97, a9b, a9a, a99, a9e)
 Digits(10)=Array(aa0, aa5, aac, aad, aa8, aa1, aa6, aaf, aa2, aa3, aa4, aa7, aab, aaa, aa9, aae)
 Digits(11)=Array(ab0, ab5, abc, abd, ab8, ab1, ab6, abf, ab2, ab3, ab4, ab7, abb, aba, ab9, abe)
 Digits(12)=Array(ac0, ac5, acc, acd, ac8, ac1, ac6, acf, ac2, ac3, ac4, ac7, acb, aca, ac9, ace)
 Digits(13)=Array(ad0, ad5, adc, add, ad8, ad1, ad6, adf, ad2, ad3, ad4, ad7, adb, ada, ad9, ade)
 Digits(14)=Array(ae0, ae5, aec, aed, ae8, ae1, ae6, aef, ae2, ae3, ae4, ae7, aeb, aea, ae9, aee)
 Digits(15)=Array(af0, af5, afc, afd, af8, af1, af6, aff, af2, af3, af4, af7, afb, afa, af9, afe)

 Digits(16)=Array(b00, b05, b0c, b0d, b08, b01, b06, b0f, b02, b03, b04, b07, b0b, b0a, b09, b0e)
 Digits(17)=Array(b10, b15, b1c, b1d, b18, b11, b16, b1f, b12, b13, b14, b17, b1b, b1a, b19, b1e)
 Digits(18)=Array(b20, b25, b2c, b2d, b28, b21, b26, b2f, b22, b23, b24, b27, b2b, b2a, b29, b2e)
 Digits(19)=Array(b30, b35, b3c, b3d, b38, b31, b36, b3f, b32, b33, b34, b37, b3b, b3a, b39, b3e)
 Digits(20)=Array(b40, b45, b4c, b4d, b48, b41, b46, b4f, b42, b43, b44, b47, b4b, b4a, b49, b4e)
 Digits(21)=Array(b50, b55, b5c, b5d, b58, b51, b56, b5f, b52, b53, b54, b57, b5b, b5a, b59, b5e)
 Digits(22)=Array(b60, b65, b6c, b6d, b68, b61, b66, b6f, b62, b63, b64, b67, b6b, b6a, b69, b6e)
 Digits(23)=Array(b70, b75, b7c, b7d, b78, b71, b76, b7f, b72, b73, b74, b77, b7b, b7a, b79, b7e)
 Digits(24)=Array(b80, b85, b8c, b8d, b88, b81, b86, b8f, b82, b83, b84, b87, b8b, b8a, b89, b8e)
 Digits(25)=Array(b90, b95, b9c, b9d, b98, b91, b96, b9f, b92, b93, b94, b97, b9b, b9a, b99, b9e)
 Digits(26)=Array(ba0, ba5, bac, bad, ba8, ba1, ba6, baf, ba2, ba3, ba4, ba7, bab, baa, ba9, bae)
 Digits(27)=Array(bb0, bb5, bbc, bbd, bb8, bb1, bb6, bbf, bb2, bb3, bb4, bb7, bbb, bba, bb9, bbe)
 Digits(28)=Array(bc0, bc5, bcc, bcd, bc8, bc1, bc6, bcf, bc2, bc3, bc4, bc7, bcb, bca, bc9, bce)
 Digits(29)=Array(bd0, bd5, bdc, bdd, bd8, bd1, bd6, bdf, bd2, bd3, bd4, bd7, bdb, bda, bd9, bde)
 Digits(30)=Array(be0, be5, bec, bed, be8, be1, be6, bef, be2, be3, be4, be7, beb, bea, be9, bee)
 Digits(31)=Array(bf0, bf5, bfc, bfd, bf8, bf1, bf6, bff, bf2, bf3, bf4, bf7, bfb, bfa, bf9, bfe)

 Sub DisplayTimer_Timer
    Dim ChgLED, ii, jj, num, chg, stat, obj, b, x
    ChgLED=Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
		If DesktopMode = True Then
       For ii=0 To UBound(chgLED)
          num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)
			if (num < 32) then
              For Each obj In Digits(num)
                   If chg And 1 Then obj.State=stat And 1
                   chg=chg\2 : stat=stat\2
                  Next
			Else
			       end if
        Next
	   end if
    End If
 End Sub

'******************************************************************************************************************************************************
'******************************************************************************************************************************************************
'VPX  call pack functions
'******************************************************************************************************************************************************
'******************************************************************************************************************************************************


'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep, UStep

Sub RightSlingShot_Slingshot
	vpmTimer.PulseSw 55
    PlaySound SoundFX("right_slingshot",DOFContactors), 0, 1, 0.05, 0.05
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 0:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 2:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0:
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
	vpmTimer.PulseSw 56
    PlaySound SoundFX("left_slingshot",DOFContactors),0,1,-0.05,0.05
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 0:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 2:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0:
    End Select
    LStep = LStep + 1
End Sub




Sub UpperSlingShot_Slingshot
	vpmTimer.PulseSw 49
    PlaySound SoundFX("left_slingshot",DOFContactors),0,1,-0.05,0.05
    USling.Visible = 0
    USling1.Visible = 1
    sling3.TransZ = -20
    UStep = 0
    UpperSlingShot.TimerEnabled = 1
End Sub

Sub UpperSlingShot_Timer
    Select Case UStep
        Case 0:USLing1.Visible = 0:USLing2.Visible = 1:sling3.TransZ = -10
        Case 2:USLing2.Visible = 0:USLing.Visible = 1:sling3.TransZ = 0:UpperSlingShot.TimerEnabled = 0:
    End Select
    UStep = UStep + 1
End Sub

'***LPF sling animations***
Sub wall66_Hit:vpmTimer.PulseSw 100:LPF_sling.visible = 0::LPF_slinga.visible = 1:wall66.timerenabled = 1:End Sub
Sub wall66_timer:LPF_sling.visible = 1::LPF_slinga.visible = 0: wall66.timerenabled= 0:End Sub

Sub wall72_Hit:vpmTimer.PulseSw 100:LPF_sling1.visible = 0::LPF_sling1a.visible = 1:wall72.timerenabled = 1:End Sub
Sub wall72_timer:LPF_sling1.visible = 1::LPF_sling1a.visible = 0: wall72.timerenabled= 0:End Sub

'***Rear wall rubber animation***
Sub sw50_Hit:vpmTimer.PulseSw 50:vpmTimer.PulseSw 100:rubber29.visible = 0::rubber29a.visible = 1:sw50.timerenabled = 1:End Sub
Sub sw50_timer:rubber29.visible = 1::rubber29a.visible = 0: sw50.timerenabled= 0:End Sub

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


'*******************************
'"GI"
'********************************
'Sub UpdateGI(enabled)
'	If enabled then
'		GIOff
'	else
'		GIOn
'	End If
'End Sub


'*******************************
'Flash GI"
'********************************

DIM ee
 Sub updateGI (enabled)
    If enabled Then
For each ee in GI_lighting:ee.State = 0: Next
    GIOff
	Playsound "fx_relay_off"
	Table1.ColorGradeImage = "ColorGrade_off"
    Else
For each ee in GI_lighting:ee.State = 1: Next
	GIOn
    Playsound "fx_relay_on"
	Table1.ColorGradeImage = "ColorGrade_on"
    End If
End Sub


'*******************************
'Flash "looney"
'********************************

DIM ff
 Sub Flashlooney (enabled)
    If enabled Then
For each ff in Flashlooney_col:ff.State = 1: Next
    Playsound "fx_relay_off"
    Else
For each ff in Flashlooney_col:ff.State = 0: Next

    Playsound "fx_relay_off"
    End If
End Sub



'*******************************
'Flash "tunes"
'********************************

DIM gg
 Sub Flashtunes (enabled)
    If enabled Then
For each gg in Flashtunes_col:gg.State = 1: Next
    Playsound "fx_relay_off"
    Else
For each gg in Flashtunes_col:gg.State = 0: Next

    Playsound "fx_relay_off"
    End If
End Sub

Sub Leftrampstart_hit():PlaySound "fx_rlenter":End Sub
Sub Mrampstart_hit():PlaySound "fx_rlenter":End Sub
Sub Mrampstart1_hit():PlaySound "fx_rlenter":End Sub
Sub Mrampstart2_hit():PlaySound "Plastic ramp":End Sub





Sub Shooter_ramp_start_Hit()
If ActiveBall.VelY < 0 Then Playsound "wireramp"
End Sub

Sub Shooter_ramp_end_Hit()
	 StopSound "wireramp"
 End Sub


'***MISC global. sounds

Sub RubbersRubbersWalls_Hit(idx):PlaySound "fx_rubber_hit_3", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub RubbersBandsLargeRings_Hit(idx):PlaySound "fx_rubber_hit_3", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub RubbersSnallRings_Hit(idx):PlaySound "fx_rubber_hit_3", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub Metals_Hit(idx):PlaySound "fx_metalhit2", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub SpotTargets_Hit(idx):PlaySound "target", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub LeftFlipper_Collide(parm) : PlaySound "flip_hit_3", 0, .5, 0.25 : End Sub
Sub RightFlipper_Collide(parm) : PlaySound "flip_hit_3", 0, .5, 0.25 : End Sub
Sub RightFlipper1_Collide(parm) : PlaySound "flip_hit_3", 0, .5, 0.25 : End Sub
Sub Gates_Hit(idx):PlaySound "gate", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub Primitive22_Hit(idx):PlaySound "fx_rubber_hit_3", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub

Sub Spinner_Spin
	PlaySound "fx_spinner",0,.25,0,0.25
End Sub

'Sub RandomSoundRubber()
'	Select Case Int(Rnd*3)+1
'		Case 1 : PlaySound "fx_rubber1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
'		Case 2 : PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
'		Case 3 : PlaySound "fx_rubber3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
'	End Select
'End Sub

'***Flipper hit sounds
'Sub LeftFlipper_Collide(parm)
' 	RandomSoundFlipper()
'End Sub
'
'Sub RightFlipper_Collide(parm)
' 	RandomSoundFlipper()
'End Sub
'
'Sub RandomSoundFlipper()
'	Select Case Int(Rnd*3)+1
'		Case 1 : PlaySound "fx_flip_hit1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
'		Case 2 : PlaySound "fx_flip_hit2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
'		Case 3 : PlaySound "fx_flip_hit3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
'	End Select
'End Sub

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

Const tnob = 9 ' total number of balls
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

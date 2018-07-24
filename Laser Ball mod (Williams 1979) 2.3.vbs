' Laser Ball Williams 1979
' Allknowing2012 - PF from vpforums resources
' Plastics from my physical machine :-)
' IPDB for switches and VP9 table
' Script segments from various tables and authors

' Thalamus 2018-07-23
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' No special SSF tweaks yet.

Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

Const cGameName = "lzbal_l2"

Dim xx
Dim dtl,dtR,bsSaucer,bsSaucer1,bsSaucer2,bsTrough

' Thalamus, heavier ball please

Const BallMass = 1.6

LoadVPM "01560000", "S6.VBS", 3.36
Dim DesktopMode: DesktopMode = Table.ShowDT

Const UseSolenoids=1,UseLamps=1,UseSync=1,UseGI=0
Const SSolenoidOn="Solenoid",SSolenoidOff="SolenoidOff",SFlipperOn="FlipperUp",SFlipperOff="FlipperDown",SCoin="coin3"

Const sBallRelease=1
Const sSaucer2=2
Const sDTLReset=3          ' solenoid for LASER
Const sSaucer=5            ' solenoid for Top Kicker
Const sRDTReset=7          ' solenoid for BALL
Const sSaucer1=8
Const sKnocker=14
Const sbuzzer=15
Const sCoinLockout=16
Const sEnable=23

SolCallback(sBallRelease)="bsTrough.SolOut"
SolCallback(sKnocker)="vpmSolSound SoundFX(""knocker"",DOFKnocker),"
SolCallback(sDTLReset)="SolRaiseDropL"  'dtL.SolDropUp"  ' Add a delay
SolCallback(sRDTReset)="SolRaiseDropR"  'dtR.SolDropUp"  ' Add a delay
SolCallback(sSaucer)="bsSaucer.SolOut"
SolCallback(sSaucer1)="bsSaucer1.SolOut"
SolCallback(sSaucer2)="bsSaucer2.SolOut"

SolCallback(sLRFlipper)="SolRFlipper"
SolCallback(sLLFlipper)="SolLFlipper"

'Fantasy Bumper Lights
solCallback(17)=  "vpmFlasher array(Flasher1,Flasher2),"
SolCallback(18)= "vpmFlasher array(Flasher3,Flasher4),"
SolCallback(19)= "vpmFlasher array(Flasher5,Flasher6),"

'SolCallback(sEnable)="vpmNudge.SolGameOn"
SolCallback(sEnable)="GameOn"

Sub GameOn(enabled)
  vpmNudge.SolGameOn(enabled)
  If Enabled Then
    GIOn
  Else
    GIOff
  End If
End Sub

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFx("fx_flipperupLeft",DOFContactors):LeftFlipper.RotateToEnd
     Else
         PlaySound SoundFx("fx_flipperdown",DOFContactors):LeftFlipper.RotateToStart
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFx("fx_flipperupRight",DOFContactors):RightFlipper.RotateToEnd:RightFlipperUpper.RotateToEnd
     Else
         PlaySound SoundFx("fx_flipperdown",DOFContactors):RightFlipper.RotateToStart:RightFlipperUpper.RotateToStart
     End If
End Sub

Sub SolRaiseDropL(enabled)
   If enabled Then
      TargetResetL.interval=500
      TargetResetL.enabled=True
   End if
End Sub

Sub TargetResetL_Timer()
  TargetResetL.enabled=False
  dtL.DropSol_On
End Sub

Sub SolRaiseDropR(enabled)
   If enabled Then
      TargetResetR.interval=500
      TargetResetR.enabled=True
   End if
End Sub

Sub TargetResetR_Timer()
  TargetResetR.enabled=False
  dtR.DropSol_On
End Sub


Sub Table_Init
	vpmInit Me
	On Error Resume Next
	With Controller
		.GameName=cGameName
		.SplashInfoLine="Laser Ball - Williams 1979"
		.HandleKeyboard=0
		.ShowTitle=0
		.ShowDMDOnly=1
		.ShowFrame=0
		.Hidden=1
		.Run
		'On Error Goto 0
	End With

	PinMAMETimer.Interval=PinMAMEInterval
	PinMAMETimer.Enabled=True

	vpmNudge.TiltSwitch=swTilt
	vpmNudge.Sensitivity=5
	vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3,LeftSlingShot,RightSlingShot)

	Set bsTrough=New cvpmBallStack
	bsTrough.InitNoTrough BallRelease,9, 45, 7
	bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors),SoundFX("drain",DOFContactors)

	Set dtL=new cvpmDropTarget
	dtL.InitDrop Array(target1,target2,target3,target4,target5),Array(19,20,21,22,23)
	dtL.InitSnd SoundFX("flapclos",DOFContactors),SoundFX("flapopen",DOFContactors)
	dtL.AllDownSw=24

	Set dtR=new cvpmDropTarget
	dtR.InitDrop Array(target6,target7,target8,target9),Array(34,35,36,37)
	dtR.InitSnd SoundFX("target",DOFContactors), SoundFX("target",DOFContactors)
	dtR.AllDownSw=38

	Set bsSaucer=New cvpmBallStack
	bsSaucer.InitSaucer Kicker1,32,205,14
	bsSaucer.InitExitSnd SoundFX("popper_ball",DOFContactors), SoundFX("popper_ball",DOFContactors)
	bsSaucer.KickForceVar=5

	Set bsSaucer1=New cvpmBallStack
	bsSaucer1.InitSaucer Kicker2,41,0,25
	bsSaucer1.InitExitSnd SoundFX("popper_ball",DOFContactors), SoundFX("popper_ball",DOFContactors)
	bsSaucer1.KickForceVar=5

	Set bsSaucer2=New cvpmBallStack
	bsSaucer2.InitSaucer Kicker3,10,0,25
	bsSaucer2.InitExitSnd SoundFX("popper_ball",DOFContactors), SoundFX("popper_ball",DOFContactors)
	bsSaucer2.KickForceVar=8.5

    GIOff
End Sub

Sub table_Paused:Controller.Pause = 1:End Sub
Sub table_unPaused:Controller.Pause = 0:End Sub

Sub table_KeyDown(ByVal keycode)
    If KeyDownHandler(keycode) Then Exit Sub
    If keycode = PlungerKey Then Plunger.Pullback:playsound "plungerpull"
End Sub

Sub table_KeyUp(ByVal keycode)
	If KeyUpHandler(keycode) Then Exit Sub
    If keycode = PlungerKey Then Plunger.Fire:PlaySound "Plunger"
End Sub

Sub GIOn
	dim bulb
	for each bulb in GILights
	bulb.state = LightStateOn
	next
	Flasher7.visible=1
End Sub

Sub GIOff
	dim bulb
	for each bulb in GILights
	bulb.state = LightStateOff
	next
	Flasher7.visible=0
End Sub

Set LampCallback=GetRef("UpdateMultipleLamps")

Sub ballreleased_hit:GIOn():End Sub

Sub Drain_Hit:bsTrough.AddBall Me:GIOff:End Sub		'switch9 Trough

Sub Kicker3_Hit:bsSaucer2.AddBall 0:End Sub			'switch10

Sub sw11_hit:Controller.Switch(11)=1:End Sub		'switch11
Sub sw11_unHit:Controller.Switch(11)=0:End Sub		'switch11
Sub sw12_Hit:Controller.Switch(12)=1:End Sub		'switch12
Sub sw12_unHit:Controller.Switch(12)=0:End Sub		'switch12

Sub sw14_Hit:vpmTimer.PulseSw(14):End Sub			'switch14

Sub sw16_Hit:vpmTimer.PulseSw(16):End Sub			'switch16
Sub Spinner1_Spin:vpmTimer.PulseSw(17):PlaySound "fx_spinner",0,.25,0,0.25:End Sub				'switch17 Left Spinner


Sub sw18_Hit:vpmTimer.PulseSw(18):End Sub						'switch18

Sub target1_Hit:dtL.Hit 1:End Sub							'switch19
Sub target2_Hit:dtL.Hit 2:End Sub							'switch20
Sub target3_Hit:dtL.Hit 3:End Sub							'switch21
Sub target4_Hit:dtL.Hit 4:End Sub							'switch22
Sub target5_Hit:dtL.Hit 5:End Sub							'switch23
															'switch24 used to reset LASER bank
Sub sw25_Hit:vpmTimer.PulseSw(25):End Sub					'switch25
Sub sw26_Hit:vpmTimer.PulseSw(26):End Sub					'switch26

Sub sw27_Hit:Controller.Switch(27)=1:End Sub		 		'switch27
Sub sw27_unHit:Controller.Switch(27)=0:End Sub			'switch27
Sub sw28_Hit:Controller.Switch(28)=1:End Sub			'switch28
Sub sw28_unHit:Controller.Switch(28)=0:End Sub			'switch28
Sub sw29_Hit:Controller.Switch(29)=1:End Sub			'switch29
Sub sw29_unHit:Controller.Switch(29)=0:End Sub			'switch29
Sub sw30_Hit:Controller.Switch(30)=1:End Sub			'switch30
Sub sw30_unHit:Controller.Switch(30)=0:End Sub			'switch30
Sub sw31_Hit:Controller.Switch(31)=1:End Sub			'switch31
Sub sw31_unHit:Controller.Switch(31)=0:End Sub			'switch31

Sub Kicker1_Hit:bsSaucer.AddBall 0:End Sub					'switch32

Sub sw33_Hit:vpmTimer.PulseSw(33):End Sub					'switch33

Sub target6_Hit:dtR.Hit 1:End Sub							'switch34
Sub target7_Hit:dtR.Hit 2:End Sub							'switch35
Sub target8_Hit:dtR.Hit 3:End Sub							'switch36
Sub target9_Hit:dtR.Hit 4:End Sub							'switch37

Sub Spinner2_Spin:vpmTimer.PulseSw(40):PlaySound "fx_spinner",0,.25,0,0.25:End Sub	'switch40 Right Spinner

Sub Kicker2_Hit:bsSaucer1.AddBall 0:End Sub					'switch41

Sub sw42_Hit:Controller.Switch(42)=1:End Sub		'switch42
Sub sw42_unHit:Controller.Switch(42)=0:End Sub		'switch42
Sub sw43_Hit:Controller.Switch(43)=1:End Sub		'switch43
Sub sw43_unHit:Controller.Switch(43)=0:End Sub		'switch43

Dim bump1,bump2,bump3

Sub Bumper1_Hit:vpmTimer.PulseSw 45:bump1 = 1:Me.TimerEnabled = 1:End Sub
Sub Bumper1_Timer()
	Select Case bump1
        Case 1:Ring1.Z = -30:bump1 = 2
        Case 2:Ring1.Z = -20:bump1 = 3
        Case 3:Ring1.Z = -10:bump1 = 4
        Case 4:Ring1.Z = 0:Me.TimerEnabled = 0
	End Select
End Sub

Sub Bumper2_Hit:vpmTimer.PulseSw 46:bump2 = 1:Me.TimerEnabled = 1:End Sub
Sub Bumper2_Timer()
	Select Case bump2
        Case 1:Ring2.Z = -30:bump2 = 2
        Case 2:Ring2.Z = -20:bump2 = 3
        Case 3:Ring2.Z = -10:bump2 = 4
        Case 4:Ring2.Z = 0:Me.TimerEnabled = 0
	End Select
End Sub

Sub Bumper3_Hit:vpmTimer.PulseSw 47:bump3 = 1:Me.TimerEnabled = 1:End Sub
Sub Bumper3_Timer()
	Select Case bump3
        Case 1:Ring3.Z = -30:bump3 = 2
        Case 2:Ring3.Z = -20:bump3 = 3
        Case 3:Ring3.Z = -10:bump3 = 4
        Case 4:Ring3.Z = 0:Me.TimerEnabled = 0
	End Select
End Sub


Sub sw50_Hit :Controller.Switch(50)=1:End Sub
Sub sw50_unHit :Controller.Switch(50)=0:End Sub
Sub sw51_Hit:Controller.Switch(51)=1:End Sub
Sub sw51_Unhit:Controller.Switch(51)=0:End Sub
Sub sw52_Hit :Controller.Switch(52)=1:End Sub
Sub sw52_Unhit:Controller.Switch(52)=0:End Sub
Sub sw53_Hit :Controller.Switch(53)=1:End Sub
Sub sw53_Unhit:Controller.Switch(53)=0:End Sub
Sub sw54_Hit :Controller.Switch(54)=1:End Sub
Sub sw54_Unhit:Controller.Switch(54)=0:End Sub
Sub sw55_Hit :Controller.Switch(55)=1:End Sub
Sub sw55_Unhit:Controller.Switch(55)=0:End Sub
Sub sw56_Hit :Controller.Switch(56)=1:End Sub
Sub sw56_Unhit:Controller.Switch(56)=0:End Sub
Sub sw57_Hit :Controller.Switch(57)=1:End Sub
Sub sw57_Unhit:Controller.Switch(57)=0:End Sub
Sub sw58_Hit :Controller.Switch(58)=1:End Sub
Sub sw58_Unhit:Controller.Switch(58)=0:End Sub
Sub sw59_Hit :Controller.Switch(59)=1:End Sub
Sub sw59_Unhit:Controller.Switch(59)=0:End Sub

 Dim dBall, dZpos

'***********************************************
'***********************************************
					'Switches
'***********************************************
'***********************************************
Sub UpdateMultipleLamps
	L24b.State=L24.State
	L9b.State=L9.State
End Sub

Set Lights(1)=L1
Set Lights(2)=L2
Set Lights(3)=L3
Set Lights(4)=L4
Set Lights(5)=L5
Set Lights(6)=L6
Set Lights(7)=L7
Set Lights(8)=L8
Set Lights(9)=L9
Set Lights(10)=L10
Set Lights(11)=L11
Set Lights(12)=L12
Set Lights(13)=L13
Set Lights(14)=L14
Set Lights(15)=L15
Set Lights(16)=L16
Set Lights(17)=L17
Set Lights(18)=L18
Set Lights(19)=L19
Set Lights(20)=L20
Set Lights(21)=L21
Set Lights(22)=L22
Set Lights(23)=L23
Set Lights(24)=L24
Set Lights(25)=L25
Set Lights(26)=L26
Set Lights(27)=L27
Set Lights(28)=L28
Set Lights(29)=L29
Set Lights(30)=L30
Set Lights(31)=L31
Set Lights(32)=L32
Set Lights(33)=L33
Set Lights(34)=L34
Set Lights(35)=L35
Set Lights(36)=L36
Set Lights(37)=L37
Set Lights(38)=L38
Set Lights(39)=L39
Set Lights(40)=L40
Set Lights(41)=L41
Set Lights(42)=L42
Set Lights(43)=L43
Set Lights(44)=L44
Set Lights(45)=L45
Set Lights(46)=L46
Set Lights(47)=L47
Set Lights(48)=L48
Set Lights(49)=L49
Set Lights(56)=L56

'**********************************************************************************************************
'Digital Display
'**********************************************************************************************************

Dim Digits(28)

' 1st Player
Digits(0) = Array(LED10,LED11,LED12,LED13,LED14,LED15,LED16)
Digits(1) = Array(LED20,LED21,LED22,LED23,LED24,LED25,LED26)
Digits(2) = Array(LED30,LED31,LED32,LED33,LED34,LED35,LED36)
Digits(3) = Array(LED40,LED41,LED42,LED43,LED44,LED45,LED46)
Digits(4) = Array(LED50,LED51,LED52,LED53,LED54,LED55,LED56)
Digits(5) = Array(LED60,LED61,LED62,LED63,LED64,LED65,LED66)

' 2nd Player
Digits(6) = Array(LED70,LED71,LED72,LED73,LED74,LED75,LED76)
Digits(7) = Array(LED80,LED81,LED82,LED83,LED84,LED85,LED86)
Digits(8) = Array(LED90,LED91,LED92,LED93,LED94,LED95,LED96)
Digits(9) = Array(LED100,LED101,LED102,LED103,LED104,LED105,LED106)
Digits(10) = Array(LED110,LED111,LED112,LED113,LED114,LED115,LED116)
Digits(11) = Array(LED120,LED121,LED122,LED123,LED124,LED125,LED126)

' 3rd Player
Digits(12) = Array(LED130,LED131,LED132,LED133,LED134,LED135,LED136)
Digits(13) = Array(LED140,LED141,LED142,LED143,LED144,LED145,LED146)
Digits(14) = Array(LED150,LED151,LED152,LED153,LED154,LED155,LED156)
Digits(15) = Array(LED160,LED161,LED162,LED163,LED164,LED165,LED166)
Digits(16) = Array(LED170,LED171,LED172,LED173,LED174,LED175,LED176)
Digits(17) = Array(LED180,LED181,LED182,LED183,LED184,LED185,LED186)

' 4th Player
Digits(18) = Array(LED190,LED191,LED192,LED193,LED194,LED195,LED196)
Digits(19) = Array(LED200,LED201,LED202,LED203,LED204,LED205,LED206)
Digits(20) = Array(LED210,LED211,LED212,LED213,LED214,LED215,LED216)
Digits(21) = Array(LED220,LED221,LED222,LED223,LED224,LED225,LED226)
Digits(22) = Array(LED230,LED231,LED232,LED233,LED234,LED235,LED236)
Digits(23) = Array(LED240,LED241,LED242,LED243,LED244,LED245,LED246)

' Credits
Digits(24) = Array(LED250,LED251,LED252,LED253,LED254,LED255,LED256)
Digits(25) = Array(LED260,LED261,LED262,LED263,LED264,LED265,LED266)

' Balls
Digits(26) = Array(LED270,LED271,LED272,LED273,LED274,LED275,LED276)
Digits(27) = Array(LED280,LED281,LED282,LED283,LED284,LED285,LED286)

Sub DisplayTimer_Timer
	Dim ChgLED,ii,num,chg,stat,obj
	ChgLED = Controller.ChangedLEDs(&H0000003f, &Hffffffff)
	If Not IsEmpty(ChgLED) Then
		If DesktopMode = True Then
		For ii = 0 To UBound(chgLED)
			num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
			if (num < 28) then
				For Each obj In Digits(num)
					If chg And 1 Then obj.State = stat And 1
					chg = chg\2 : stat = stat\2
				Next
			else

			end if
		next
		end if
	end if
End Sub


' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Sub Pins_Hit (idx)
   debug.prnit "pin"
	PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
	PlaySound SoundFX("target",DOFContactors), 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub TargetBankWalls_Hit (idx)
	PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
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

Sub Rubbers_Hit(idx)
   debug.print "rubber"
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
    debug.print "Posts"
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 16 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 16 then
 		RandomSoundRubber()
 	End If
End Sub

Sub Rubbers_Hit(idx)
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub

Sub Bumpers_Hit(idx)
	Select Case Int(Rnd*4)+1
		Case 1 : PlaySound SoundFx("fx_bumper1",DOFContactors), 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound SoundFx("fx_bumper2",DOFContactors), 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound SoundFx("fx_bumper3",DOFContactors), 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 4 : PlaySound SoundFx("fx_bumper4",DOFContactors), 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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


'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep


Sub RightSlingShot_Slingshot
    PlaySound Soundfx("right_slingshot",DOFContactors), 0, 1, 0.05, 0.05
    vpmTimer.PulseSw 44
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 1: if B2SOn Then Controller.B2SSetData 73,0
                if B2SOn Then Controller.B2SSetData 71,0
                if B2SOn Then Controller.B2SSetData 72,0
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
               if B2SOn Then Controller.B2SSetData 73,1
               if B2SOn Then Controller.B2SSetData 71,1
               if B2SOn Then Controller.B2SSetData 72,1
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
    PlaySound Soundfx("left_slingshot",DOFContactors), 0, 1, 0.05, 0.05
    vpmTimer.PulseSw 13
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 1: if B2SOn Then Controller.B2SSetData 73,0
                if B2SOn Then Controller.B2SSetData 71,0
                if B2SOn Then Controller.B2SSetData 72,0
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
               if B2SOn Then Controller.B2SSetData 73,1
               if B2SOn Then Controller.B2SSetData 71,1
               if B2SOn Then Controller.B2SSetData 72,1
    End Select
    LStep = LStep + 1
End Sub

Dim LKStep, RKStep
Sub Trigger2_Hit
  LKStep=0:LKTimer.enabled=True
End Sub

Sub LKTimer_Timer
  Select Case LKStep
    Case 1: LK.TransZ=-5
    Case 2: LK.TransZ=-15
    Case 3: LK.TransZ=-25
    Case 4: LK.TransZ=0:LKTimer.enabled=False
  End Select
  LKStep=LKStep+1
End Sub

Sub Trigger1_Hit
  RKStep=0:RKTimer.enabled=True
  debug.print "trigger1 hit"
End Sub

Sub RKTimer_Timer
  Select Case RKStep
    Case 1: RK.TransZ=-5
    Case 2: RK.TransZ=-15
    Case 3: RK.TransZ=-25
    Case 4: RK.TransZ=0:RKTimer.enabled=False:debug.print "Done"
  End Select
  RKStep=RKStep+1
End Sub


Sub GraphicsTimer_Timer()
	batleft.objrotz = LeftFlipper.CurrentAngle + 1
	batright.objrotz = RightFlipper.CurrentAngle - 1
    'batleft1.objrotz = LeftFlipper.CurrentAngle + 1
	batright1.objrotz = RightFlipperUpper.CurrentAngle - 1
End Sub


'*****************************************
'	ninuzzu's	BALL SHADOW
'*****************************************
Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2)

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
        If BOT(b).X < Table.Width/2 Then
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table.Width/2))/7)) + 6
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table.Width/2))/7)) - 6
        End If
        ballShadow(b).Y = BOT(b).Y + 12
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
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

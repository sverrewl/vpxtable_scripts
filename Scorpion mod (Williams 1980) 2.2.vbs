' Scorpion Williams 1980
' Allknowing2012 - PF, plastics from internet
' IPDB for switches and VP9 table
' Script segments from various tables and authors
' 1.0b - hide the backglass lights when no int desktop mode

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' No special SSF tweaks yet.

Option Explicit
Randomize

If Table.ShowDT = False then
  light64.visible=False
  light57.visible=False
  light58.visible=False
  light59.visible=False
  light60.visible=False
  light50.visible=False
  light51.visible=False
  light52.visible=False
  light53.visible=False
End If

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

Const cGameName = "scrpn_l1"

Dim xx
Dim dtl,dtC1,dtC2,dtR,bsSaucer,bsSaucer2,bsTrough

' Thalamus, heavier ball please

Const BallMass = 1.5

LoadVPM "01560000", "S6.VBS", 3.36
Dim DesktopMode: DesktopMode = Table.ShowDT

Const UseSolenoids=2,UseLamps=1,UseSync=1,UseGI=0
Const SSolenoidOn="SolOn",SSolenoidOff="SolOff",SFlipperOn="FlipperUp",SFlipperOff="FlipperDown",SCoin="coin3"


Const sOutHole=1
Const sC1DTReset=2
Const sC2DTReset=3
Const sLDTReset=4
Const sRDTReset=5         ' solenoid for Top Right Targets
Const sBallRelease=6
Const sKicker1=7
Const sKicker3=8

Const sKnocker=14
Const sbuzzer=15
Const sCoinLockout=16
Const sEnable=23

SolCallback(sOutHole)="bsTrough.SolIn"
SolCallback(sKnocker)="vpmSolSound SoundFX(""knocker"",DOFKnocker),"
SolCallback(sLDTReset)="SolRaiseDropL"  'dtL.SolDropUp"  ' Add a delay
SolCallback(sRDTReset)="SolRaiseDropR"  'dtR.SolDropUp"  ' Add a delay
SolCallback(sBallRelease)="bsTrough.SolOut"
SolCallback(sC1DTReset)="SolRaiseDropC1"  'Add a delay
SolCallback(sC2DTReset)="SolRaiseDropC2"  'Add a delay
SolCallback(sKicker1)="bsSaucer.SolOut"
SolCallback(sKicker3)="bsSaucer2.SolOut"

SolCallback(sLRFlipper)="SolRFlipper"
SolCallback(sLLFlipper)="SolLFlipper"

'Fantasy Bumper Lights
solCallback(17)= "vpmFlasher array(Flasher1,Flasher2),"
SolCallback(18)= "vpmFlasher array(Flasher3,Flasher4),"
SolCallback(19)= "vpmFlasher array(Flasher5,Flasher6),"

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
         PlaySound SoundFx("fx_flipperupLeft",DOFContactors):LeftFlipper.RotateToEnd:LeftFlipperUpper.RotateToEnd
     Else
         PlaySound SoundFx("fx_flipperdown",DOFContactors):LeftFlipper.RotateToStart::LeftFlipperUpper.RotateToStart
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
   debug.print "SolRaiseDropL"
   If enabled Then
      TargetResetL.interval=500
      TargetResetL.enabled=True
   End if
End Sub

Sub TargetResetL_Timer()
  debug.print "Raise Left Targets"
  TargetResetL.enabled=False
  dtL.DropSol_On
End Sub

Sub SolRaiseDropR(enabled)
   debug.print "SolRaiseDropR"
   If enabled Then
      TargetResetR.interval=500
      TargetResetR.enabled=True
   End if
End Sub

Sub TargetResetR_Timer()
  debug.print "Raise Right Targets"
  TargetResetR.enabled=False
  dtR.DropSol_On
End Sub

Sub SolRaiseDropC1(enabled)
   If enabled Then
      TargetResetC1.interval=500
      TargetResetC1.enabled=True
   End if
End Sub

Sub TargetResetC1_Timer()
  TargetResetC1.enabled=False
  dtC1.DropSol_On
End Sub

Sub SolRaiseDropC2(enabled)
   If enabled Then
      TargetResetC2.interval=500
      TargetResetC2.enabled=True
   End if
End Sub

Sub TargetResetC2_Timer()
  TargetResetC2.enabled=False
  dtC2.DropSol_On
End Sub


Sub Table_Init
	vpmInit Me
	On Error Resume Next
	With Controller
		.GameName=cGameName
		.SplashInfoLine="Scorpion - Williams 1980"
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
	bsTrough.InitSw 0,9,10,0,0,0,0,0
    bsTrough.InitKick BallRelease,45,6
	bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors),SoundFX("drain",DOFContactors)
    bsTrough.Balls=2

	Set dtL=new cvpmDropTarget
	dtL.InitDrop Array(target1,target2,target3),Array(25,26,27)
	dtL.InitSnd SoundFX("target",DOFContactors),SoundFX("target",DOFContactors)

	Set dtR=new cvpmDropTarget
	dtR.InitDrop Array(target4,target5,target6),Array(28,29,30)
	dtR.InitSnd SoundFX("target",DOFContactors), SoundFX("target",DOFContactors)
	dtR.AllDownSw=31
    dtR.LinkedTo=dtL
    dtL.LinkedTo=dtR

	Set dtC1=new cvpmDropTarget
	dtC1.InitDrop Array(target7,target8),Array(33,34)
	dtC1.InitSnd SoundFX("target",DOFContactors), SoundFX("target",DOFContactors)


	Set dtC2=new cvpmDropTarget
	dtC2.InitDrop Array(target9,target10,target11),Array(35,36,37)
	dtC2.InitSnd SoundFX("target",DOFContactors), SoundFX("target",DOFContactors)
	dtC2.AllDownSw=38
    dtC2.LinkedTo=dtC1
    dtC1.LinkedTo=dtC2

	Set bsSaucer=New cvpmBallStack
	bsSaucer.InitSaucer Kicker1,11,165,4
	bsSaucer.InitExitSnd SoundFX("popper_ball",DOFContactors), SoundFX("popper_ball",DOFContactors)
	bsSaucer.KickForceVar=5

	Set bsSaucer2=New cvpmBallStack
	bsSaucer2.InitSaucer Kicker3,12,205,4
	bsSaucer2.InitExitSnd SoundFX("popper_ball",DOFContactors), SoundFX("popper_ball",DOFContactors)
	bsSaucer2.KickForceVar=8.5

    vpmMapLights InsertLights

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

'Set LampCallback=GetRef("UpdateMultipleLamps")

Sub ballreleased_hit:GIOn():End Sub

'Sub Drain_Hit:bsTrough.AddBall Me:GIOff:End Sub	'switch9 Trough
Sub Drain_Hit:bsTrough.AddBall Me:End Sub			'switch9 Thal, don't turn off GI it might be a multiball

Sub Kicker1_Hit:bsSaucer.AddBall 0:End Sub
Sub Kicker3_Hit:bsSaucer2.AddBall 0:End Sub			'switch12

Sub sw41_Hit:vpmTimer.PulseSw(41):End Sub			'switch41
Sub sw42_Hit:vpmTimer.PulseSw(42):End Sub			'switch42

Sub Spinner1_Spin:vpmTimer.PulseSw(32):PlaySound "fx_spinner",0,.25,0,0.25:End Sub				'switch32 Left Spinner

Sub target1_Hit:dtL.Hit 1:End Sub							'switch25
Sub target2_Hit:dtL.Hit 2:End Sub							'switch26
Sub target3_Hit:dtL.Hit 3:End Sub							'switch27

Sub target4_Hit:dtR.Hit 1:End Sub							'switch28
Sub target5_Hit:dtR.Hit 2:End Sub							'switch29
Sub target6_Hit:dtR.Hit 3:End Sub							'switch30

Sub target7_Hit:dtC1.Hit 1:End Sub							'switch33
Sub target8_Hit:dtC1.Hit 2:End Sub							'switch34

Sub target9_Hit:dtC2.Hit 1:End Sub							'switch35
Sub target10_Hit:dtC2.Hit 2:End Sub							'switch36
Sub target11_Hit:dtC2.Hit 3:End Sub							'switch37

'***********************************************
'***********************************************
					'Switches
'***********************************************
'***********************************************


Sub sw14_Hit:Controller.Switch(14)=1:End Sub
Sub sw14_unHit:Controller.Switch(14)=0:End Sub
Sub sw18_Hit:Controller.Switch(18)=1:End Sub
Sub sw18_unHit:Controller.Switch(18)=0:End Sub
Sub sw19_Hit:Controller.Switch(19)=1:End Sub
Sub sw19_unHit:Controller.Switch(19)=0:End Sub

Sub sw20_Hit:Controller.Switch(20)=1:End Sub
Sub sw20_unHit:Controller.Switch(20)=0:End Sub
Sub sw21_Hit:Controller.Switch(21)=1:End Sub
Sub sw21_unHit:Controller.Switch(21)=0:End Sub
Sub sw15_Hit:Controller.Switch(15)=1:End Sub
Sub sw15_unHit:Controller.Switch(15)=0:End Sub

Sub sw46_Hit:Controller.Switch(46)=1:End Sub
Sub sw46_unHit:Controller.Switch(46)=0:End Sub
Sub sw47_Hit:Controller.Switch(47)=1:End Sub
Sub sw47_unHit:Controller.Switch(47)=0:End Sub
Sub sw48_Hit:Controller.Switch(48)=1:End Sub
Sub sw48_unHit:Controller.Switch(48)=0:End Sub

Sub sw49_Hit:Controller.Switch(49)=1:End Sub
Sub sw49_unHit:Controller.Switch(49)=0:End Sub

Sub sw43_Hit:vpmTimer.PulseSw(43):End Sub
Sub sw44_Hit:vpmTimer.PulseSw(44):End Sub
Sub sw45_Hit:vpmTimer.PulseSw(45):End Sub


Dim bump1,bump2,bump3

Sub Bumper1_Hit:vpmTimer.PulseSw 24:bump1 = 1:Me.TimerEnabled = 1:End Sub
Sub Bumper1_Timer()
	Select Case bump1
        Case 1:Ring1.Z = -30:bump1 = 2
        Case 2:Ring1.Z = -20:bump1 = 3
        Case 3:Ring1.Z = -10:bump1 = 4
        Case 4:Ring1.Z = 0:Me.TimerEnabled = 0
	End Select
End Sub

Sub Bumper2_Hit:vpmTimer.PulseSw 22:bump2 = 1:Me.TimerEnabled = 1:End Sub
Sub Bumper2_Timer()
	Select Case bump2
        Case 1:Ring2.Z = -30:bump2 = 2
        Case 2:Ring2.Z = -20:bump2 = 3
        Case 3:Ring2.Z = -10:bump2 = 4
        Case 4:Ring2.Z = 0:Me.TimerEnabled = 0
	End Select
End Sub

Sub Bumper3_Hit:vpmTimer.PulseSw 23:bump3 = 1:Me.TimerEnabled = 1:End Sub
Sub Bumper3_Timer()
	Select Case bump3
        Case 1:Ring3.Z = -30:bump3 = 2
        Case 2:Ring3.Z = -20:bump3 = 3
        Case 3:Ring3.Z = -10:bump3 = 4
        Case 4:Ring3.Z = 0:Me.TimerEnabled = 0
	End Select
End Sub


Sub sw50_Hit :vpmTimer.PulseSw 50:End Sub
Sub sw51_Hit :vpmTimer.PulseSw 51:End Sub
Sub sw52_Hit :vpmTimer.PulseSw 52:End Sub
Sub sw53_Hit :vpmTimer.PulseSw 53:End Sub
Sub sw54_Hit :vpmTimer.PulseSw 54:End Sub
Sub sw55_Hit :vpmTimer.PulseSw 55:End Sub
Sub sw56_Hit :vpmTimer.PulseSw 56:End Sub
Sub sw57_Hit :vpmTimer.PulseSw 57:End Sub

 Dim dBall, dZpos

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
    debug.print "metals_medium"
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


Sub LeftSlingShot_Slingshot
    PlaySound Soundfx("left_slingshot",DOFContactors), 0, 1, 0.05, 0.05
    vpmTimer.PulseSw 39
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

Sub RightSlingShot_Slingshot
    PlaySound Soundfx("right_slingshot",DOFContactors), 0, 1, 0.05, 0.05
    vpmTimer.PulseSw 40
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub GraphicsTimer_Timer()
	batleft.objrotz = LeftFlipper.CurrentAngle + 1
	batright.objrotz = RightFlipper.CurrentAngle - 1
    batleft1.objrotz = LeftFlipperUpper.CurrentAngle + 1
	batright1.objrotz = RightFlipperUpper.CurrentAngle - 1
End Sub

'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 3 ' total number of balls
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
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


' Thalamus : Exit in a clean and proper way
Sub Table_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


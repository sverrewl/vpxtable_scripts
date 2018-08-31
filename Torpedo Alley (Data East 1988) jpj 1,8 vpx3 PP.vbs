Option Explicit
Randomize

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' Script uses non standard ball rolling
' No special SSF tweaks yet.
' Wob 2018-08-08
' Added vpmInit Me to table init

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const BallSize = 50
Public check,ca9,ca90,ca901,ca0,ca1,ca01,ca91

'Load the core.vbs for supporting subs and functions

LoadVPM "01500000","DE.VBS",3.10

Const cGameName="torp_e21",UseSolenoids=2,UseLamps=1,UseSync=0,UseGI=0
Const SSolenoidOn="solon",SSolenoidOff="soloff",SFlipperOn="FlipperUp",SFlipperOff="FlipperDown",SCoin="coin3"

Const sBallRelease=16	'ok
Const sSolBRelay=10'Not Used
Const sGI=11'Not Used
Const sLSKicker=25		'ok
Const sLeftKick=26		'ok
Const sUpKicker=27		'ok
Const sCenterKicker=28	'ok
Const s3BankReset=29	'ok
Const sKnocker=30		'ok
Const sOutHole=31		'ok
Const sSinkingShip=32

'SolCallback(3)="vpmFlasher Flasher3L,"
'SolCallback(8)="vpmFlasher Flash8,"
SolCallback(sBallRelease)="bsTrough.SolOut"				'Sol16
SolCallback(sLSKicker)="vpmSolAutoPlunger Plunger1,1,"	'Sol25
SolCallback(SGI)="GIUpdate"
SolCallback(sLeftKick)="bsLeftLock.SolOut"				'Sol26
SolCallback(sUpKicker)="VUKKick"						'Sol27
SolCallback(sCenterKicker)="bsRightLock.SolOut"			'Sol28
'SolCallback(s3BankReset)="dtT.SolDropUp"				'Sol29
SolCallBack(s3BankReset) = "ResetDrops"					'Sol29 method changed JPJ
SolCallback(sKnocker)="vpmSolSound SoundFX(""knocker"",DOFKnocker),"		'Sol30
SolCallback(sOutHole)="bsTrough.SolIn"					'Sol31
SolCallback(sSinkingShip)="SinkIt"

'FLASHERS
Const sDestroyHotdog=1
Const sTorpedoHotdog=2
Const sFlagshipHotdog=3
Const sAircraftHotdog=4
Const sSpecialHotdog=5
Const sCruiserHotdog=6
Const sScope=7
Const sInsert=8
Const sLeftPair=9
Const sCenterPair=14
Const sRightPair=15

SolCallback(sDestroyHotdog)="DESTHD"				'Sol1
SolCallback(sTorpedoHotdog)="TorpHD"				'Sol2
SolCallback(sFlagshipHotdog)="FlagHD"				'Sol3
SolCallback(sAircraftHotdog)="ACHD"					'Sol4
SolCallback(sSpecialHotdog)="SpecialHD"				'Sol5
SolCallback(sCruiserHotdog)="CRUISEHD"				'Sol6
SolCallback(sScope)="ScopeHD"						'Sol7
'SolCallback(sInsert)="vpmFlasher ,"				'Sol8
SolCallback(sLeftPair)="SolLeft"					'Sol9
SolCallback(sCenterPair)="SolCentre"				'Sol14
SolCallback(sRightPair)="SolRight"					'Sol15

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
		 If Enabled Then
			 PlaySound SoundFX("FlipperUp",DOFFlippers):LeftFlipper.RotateToEnd
		 Else
			 PlaySound SoundFX("FlipperDown",DOFFlippers):LeftFlipper.RotateToStart
		 End If
	  End Sub

	Sub SolRFlipper(Enabled)
		 If Enabled Then
			 PlaySound SoundFX("FlipperUp",DOFFlippers):RightFlipper.RotateToEnd:RightFlipper1.RotateToEnd
		 Else
			 PlaySound SoundFX("FlipperDown",DOFFlippers):RightFlipper.RotateToStart:RightFlipper1.RotateToStart
		 End If
	End Sub

Dim DTBank
Dim bsTrough,dtT,bsLeftLock,bsTopLock,bsRightLock
Dim obj
Dim shpos
Dim subpos
dim sw29l 'to know if droptarget on or off to get them ok after flashing GI lights JPJ
dim sw30l
dim sw31l

check = 1

If Table1.ShowDT = False then
    Scoretext.Visible = false
	Ramp15.Visible = False
	Ramp16.Visible = False
End If

Sub Table1_Init
vpmInit Me
On Error Resume Next
		With Controller
			.GameName=cGameName
			If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
			.SplashInfoLine = "Torpedo Alley - Data East" & vbnewline & "Table by Destruk/TAB"
			.HandleMechanics=0
			.HandleKeyboard=0
			.ShowDMDOnly=1
			.ShowFrame=0
			.ShowTitle=0
			.DIP(0)=&H00
			.Run
 			.Hidden=1
			If Err Then MsgBox Err.Description
		End With
	On Error Goto 0

	vpmNudge.TiltSwitch=1
	vpmNudge.Sensitivity=5
	vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingshot)

	Set bsTrough=New cvpmBallStack
	bsTrough.InitSw 10,13,12,11,0,0,0,0
	bsTrough.InitKick BallRelease,90,6
	bsTrough.InitEntrySnd "solenoid", "solenoid"
	bsTrough.InitExitSnd SoundFX("ballrel",DOFContactors), SoundFX("solenoid",DOFContactors)
	bsTrough.Balls=3

	Set dtT=New cvpmDropTarget
	dtT.InitDrop Array(SW29,SW30,SW31),Array(29,30,31)
	dtT.InitSnd SoundFX("target_drop",DOFDropTargets),SoundFX("reset_drop",DOFContactors)

	Set bsLeftLock=New cvpmBallStack
	bsLeftLock.InitSw 0,42,41,0,0,0,0,0
	bsLeftLock.InitKick Kicker2,45,15
	bsLeftLock.InitExitSnd SoundFX("popper",DOFContactors),SoundFX("solenoid",DOFContactors)
	'bsLeftLock.KickBalls=2

	Set bsTopLock=New cvpmBallStack
	bsTopLock.InitSw 0,43,0,0,0,0,0,0
	bsTopLock.InitKick VUK,45,10
	bsTopLock.InitExitSnd SoundFX("popper",DOFContactors),SoundFX("solenoid",DOFContactors)

	Set bsRightLock=New cvpmBallStack
	bsRightLock.InitSw 0,45,44,0,0,0,0,0
	bsRightLock.InitKick Kicker1,45,35
	bsRightLock.InitExitSnd SoundFX("popper",DOFContactors),SoundFX("solenoid",DOFContactors)
	bsRightLock.KickBalls=2

    SW37A.isdropped = 1
    SW56A.isdropped = 1
    Plunger1.Pullback
	Sink5.isdropped = 1
	Sink4.isdropped = 1
	Sink3.isdropped = 1
	Sink2.isdropped = 1
	Sink1.isdropped = 0
    'Shpos = 1
    'SinkShip.enabled = 1
End Sub

Sub Table1_KeyDown(ByVal keycode)


	If keycode = PlungerKey Then
		Plunger.PullBack
		PlaySound "plungerpull",0,1,0.25,0.25
	End If

	If keycode = LeftTiltKey Then
		SbmLR = -1
		Nudge 90, 2
	End If

	If keycode = RightTiltKey Then
		SbmLR = 1
		Nudge 270, 2
	End If

	If keycode = CenterTiltKey Then
		Nudge 0, 2
	End If

    If KeyCode=LeftFlipperKey Then Controller.Switch(15)=1
	If KeyCode=RightFlipperKey Then Controller.Switch(16)=1
	If vpmKeyDown(KeyCode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal keycode)

	If keycode = PlungerKey Then
		Plunger.Fire
		PlaySound "plunger",0,1,0.25,0.25
	End If
	If KeyCode=LeftFlipperKey Then Controller.Switch(15)=0
	If KeyCode=RightFlipperKey Then Controller.Switch(16)=0
	If vpmKeyUp(KeyCode) Then Exit Sub
End Sub

Sub Drain_Hit:bsTrough.AddBall Me:PlaySound "drain":End Sub


Sub Plunger_Init()
	'PlaySound "ballrelease",0,0.5,0.5,0.25
	'Plunger.CreateBall
	'BallRelease.CreateBall
	'BallRelease.Kick 90, 8
End Sub


Sub Bumper1_Hit
    vpmTimer.PulseSw 50
	PlaySound SoundFX("fx_bumper4",DOFContactors)
	'B1L1.State = 1:B1L2. State = 1
	Me.TimerEnabled = 1
	bumper3Flash.state = 1
End Sub

Sub Bumper1_Timer
	'B1L1.State = 0:B1L2. State = 0
	Me.Timerenabled = 0
	bumper3Flash.state = 0
End Sub

Sub Bumper2_Hit
    vpmTimer.PulseSw 51
	PlaySound SoundFX("fx_bumper4",DOFContactors)
	'B2L1.State = 1:B2L2. State = 1
	Me.TimerEnabled = 1
	bumper2Flash.state = 1
End Sub

Sub Bumper2_Timer
	'B2L1.State = 0:B2L2. State = 0
	Me.Timerenabled = 0
	bumper2Flash.state = 0
End Sub

Sub Bumper3_Hit
    vpmTimer.PulseSw 49
	PlaySound SoundFX("fx_bumper4",DOFContactors)
	'B3L1.State = 1:B3L2. State = 1
	Me.TimerEnabled = 1
	bumper1Flash.state = 1
End Sub

Sub Bumper3_Timer
	'B3L1.State = 0:B3L2. State = 0
	Me.Timerenabled = 0
	bumper1Flash.state = 0
End Sub

Sub SolCentre(Enabled)
If enabled Then
SCenter.state = 1
LC.State = 1
else
SCenter.state = 0
LC.State = 0
End If
End Sub

Sub SolLeft(Enabled)
If enabled Then
SLeft.state = 1
LL.State = 1
else
SLeft.state = 0
LL.State = 0
End If
End Sub

Sub SolRight(Enabled)
If enabled Then
SRight.state = 1
LR.State = 1
else
SRight.state = 0
LR.State = 0
End If
End Sub

Sub SpecialHD(Enabled)
If enabled Then
F30.Visible = 1
else
F30.Visible = 0
End If
End Sub
'*************************************************************************
Sub ScopeHD(Enabled)
If enabled Then
FSCOPE01.state=1
FSCOPE02.state=1
FSCOPE03.state=1
FSCOPE04.state=1
FSCOPE05.state=1
FSCOPE06.state=1
'FScope.Visible = 1
else
FSCOPE01.state=0
FSCOPE02.state=0
FSCOPE03.state=0
FSCOPE04.state=0
FSCOPE05.state=0
FSCOPE06.state=0
'FScope.Visible = 0
End If
End Sub

Sub ACHD(Enabled)
If enabled Then
F22.Visible = 1
else
F22.Visible = 0
End If
End Sub

Sub TorpHD(Enabled)
If enabled Then
F8.Visible = 1
else
F8.Visible = 0
End If
End Sub

Sub FlagHD(Enabled)
If enabled Then
F45.Visible = 1
else
F45.Visible = 0
End If
End Sub

Sub DestHD(Enabled)
If enabled Then
F36.Visible = 1
else
F36.Visible = 0
End If
End Sub

Sub CruiseHD(Enabled)
If enabled Then
F40.Visible = 1
else
F40.Visible = 0
End If
End Sub

Sub RLKICK_hit()
me.destroyball
BsRightLock.AddBall Me
PlaySound "Scoopenter"
End Sub

Sub LLKICK_hit()
me.destroyball
BsLeftLock.addball Me
PlaySound "Scoopenter"
End Sub

Sub VUK_hit()
'me.destroyball
hasbeenhit = 1
Vtimer.enabled = 1
PlaySound "Scoopenter"
End Sub

Sub Vtimer_timer()
Controller.Switch(43) = 1
me.enabled = 0
End Sub


'****Targets

Sub SW14_hit():vpmTimer.PulseSw 14:End Sub
Sub SW54_hit():vpmTimer.PulseSw 54:End Sub
Sub SW55_hit():vpmTimer.PulseSw 55:End Sub
Sub SW17_hit():Controller.Switch(17) = 1:End Sub
Sub SW17_Unhit():Controller.Switch(17) = 0:End Sub
Sub SW18_hit():vpmTimer.PulseSw 18:End Sub
Sub SW26_hit():vpmTimer.PulseSw 26:End Sub
Sub SW27_hit():vpmTimer.PulseSw 27:End Sub
Sub SW28_hit():vpmTimer.PulseSw 28:End Sub
Sub SW25_Spin():vpmTimer.PulseSw 25:PlaySound "fx_spinner":End Sub
Sub SW32_Spin():vpmTimer.PulseSw 32:PlaySound "fx_spinner":End Sub

'Added jpj
Sub SW23_hit():vpmTimer.PulseSw 18:SubDown1.Enabled = 1:End Sub
Sub SW36_hit():vpmTimer.PulseSw 26:SubDown2.Enabled = 1:End Sub



'************************************************************************

 ' Drop Targets jpj variation (changing lights) - Base from cyberpez code (back to the future)


	dim sw29Dir, sw30Dir, sw31Dir
	dim sw29Pos, sw30Pos, sw31Pos

	sw29Dir = 1:sw30Dir = 1:sw30Dir = 1
	sw29Pos = 0:sw30Pos = 0:sw30Pos = 0

  'Targets Init
	sw29a.TimerEnabled = 1:sw30a.TimerEnabled = 1:sw31a.TimerEnabled = 1





  Sub sw29_Hit:DTBank.Hit 1:sw29Dir = 0:sw29l=0:sw29a.TimerEnabled = 1:check=0:End Sub
  Sub sw30_Hit:DTBank.Hit 2:sw30Dir = 0:sw30l=0:sw30a.TimerEnabled = 1:check=0:End Sub
  Sub sw31_Hit:DTBank.Hit 3:sw31Dir = 0:sw31l=0:sw31a.TimerEnabled = 1:check=0:End Sub

   	Set DTBank = New cvpmDropTarget
   	  With DTBank
   		.InitDrop Array(Array(sw29,sw29a),Array(sw30,sw30a),Array(sw31,sw31a)), Array(29,30,31)
		.InitSnd SoundFX("target_drop",DOFDropTargets),SoundFX("reset_drop",DOFContactors)
       End With

Sub lightchoice
	if sw29Dir = 0 and sw30Dir = 1 and sw31Dir = 1 then
		Lightcvup1a9.state = 1:Lightcvup3aa9.state=1
		Lightcvup1a01.state = 0:Lightcvup3aa01.state=0
		Lightcvup3a1.state = 0:Lightcvup1aa1.state=0
		Lightcvup1a0.state = 0:Lightcvup3aa0.state=0
		Lightcvup1a91.state = 0:Lightcvup3aa91.state=0
		Lightcvup1a901.state = 0:Lightcvup3aa901.state=0
		Lightcvup1a90.state = 0:Lightcvup3aa90.state=0
		Lightcvup4.state=0:Lightcvup2.state=0'default light behind Targets
		Check=0
	end if
	if sw29Dir = 0 and sw30Dir = 0 and sw31Dir = 1 then
		Lightcvup1a90.state = 1:Lightcvup3aa90.state=1
		Lightcvup1a01.state = 0:Lightcvup3aa01.state=0
		Lightcvup3a1.state = 0:Lightcvup1aa1.state=0
		Lightcvup1a0.state = 0:Lightcvup3aa0.state=0
		Lightcvup1a91.state = 0:Lightcvup3aa91.state=0
		Lightcvup1a901.state = 0:Lightcvup3aa901.state=0
		Lightcvup1a9.state = 0:Lightcvup3aa9.state=0
		Lightcvup4.state=0:Lightcvup2.state=0'default light behind Targets
		Check=0
	end if
	if sw29Dir = 0 and sw30Dir = 0 and sw31Dir = 0 then
		Lightcvup1a901.state = 1:Lightcvup3aa901.state=1
		Lightcvup1a01.state = 0:Lightcvup3aa01.state=0
		Lightcvup3a1.state = 0:Lightcvup1aa1.state=0
		Lightcvup1a0.state = 0:Lightcvup3aa0.state=0
		Lightcvup1a91.state = 0:Lightcvup3aa91.state=0
		Lightcvup1a90.state = 0:Lightcvup3aa90.state=0
		Lightcvup1a9.state = 0:Lightcvup3aa9.state=0
		Lightcvup4.state=0:Lightcvup2.state=0'default light behind Targets
		Check=0
	end if
	if sw29Dir = 0 and sw30Dir = 1 and sw31Dir = 0 then
		Lightcvup1a91.state = 1:Lightcvup3aa91.state=1
		Lightcvup1a01.state = 0:Lightcvup3aa01.state=0
		Lightcvup3a1.state = 0:Lightcvup1aa1.state=0
		Lightcvup1a0.state = 0:Lightcvup3aa0.state=0
		Lightcvup1a901.state = 0:Lightcvup3aa901.state=0
		Lightcvup1a90.state = 0:Lightcvup3aa90.state=0
		Lightcvup1a9.state = 0:Lightcvup3aa9.state=0
		Lightcvup4.state=0:Lightcvup2.state=0'default light behind Targets
		Check=0
	end if
		if sw29Dir = 1 and sw30Dir = 0 and sw31Dir = 1 then
		Lightcvup1a0.state = 1:Lightcvup3aa0.state=1
		Lightcvup1a01.state = 0:Lightcvup3aa01.state=0
		Lightcvup3a1.state = 0:Lightcvup1aa1.state=0
		Lightcvup1a91.state = 0:Lightcvup3aa91.state=0
		Lightcvup1a901.state = 0:Lightcvup3aa901.state=0
		Lightcvup1a90.state = 0:Lightcvup3aa90.state=0
		Lightcvup1a9.state = 0:Lightcvup3aa9.state=0
		Lightcvup4.state=0:Lightcvup2.state=0'default light behind Targets
		Check=0
	end if
		if sw29Dir = 1 and sw30Dir = 1 and sw31Dir = 0 then
		Lightcvup3a1.state = 1:Lightcvup1aa1.state=1
		Lightcvup1a01.state = 0:Lightcvup3aa01.state=0
		Lightcvup1a0.state = 0:Lightcvup3aa0.state=0
		Lightcvup1a91.state = 0:Lightcvup3aa91.state=0
		Lightcvup1a901.state = 0:Lightcvup3aa901.state=0
		Lightcvup1a90.state = 0:Lightcvup3aa90.state=0
		Lightcvup1a9.state = 0:Lightcvup3aa9.state=0
		Lightcvup4.state=0:Lightcvup2.state=0'default light behind Targets
		Check=0
	end if
		if sw29Dir = 1 and sw30Dir = 0 and sw31Dir = 0 then
		Lightcvup1a01.state = 1:Lightcvup3aa01.state=1
		Lightcvup3a1.state = 0:Lightcvup1aa1.state=0
		Lightcvup1a0.state = 0:Lightcvup3aa0.state=0
		Lightcvup1a91.state = 0:Lightcvup3aa91.state=0
		Lightcvup1a901.state = 0:Lightcvup3aa901.state=0
		Lightcvup1a90.state = 0:Lightcvup3aa90.state=0
		Lightcvup1a9.state = 0:Lightcvup3aa9.state=0
		Lightcvup4.state=0:Lightcvup2.state=0'default light behind Targets
		Check=0
	end if
		if sw29Dir = 1 and sw30Dir = 1 and sw31Dir = 1 then
		Lightcvup1a01.state = 0:Lightcvup3aa01.state=0
		Lightcvup3a1.state = 0:Lightcvup1aa1.state=0
		Lightcvup1a0.state = 0:Lightcvup3aa0.state=0
		Lightcvup1a91.state = 0:Lightcvup3aa91.state=0
		Lightcvup1a901.state = 0:Lightcvup3aa901.state=0
		Lightcvup1a90.state = 0:Lightcvup3aa90.state=0
		Lightcvup1a9.state = 0:Lightcvup3aa9.state=0
		Lightcvup4.state=1:Lightcvup2.state=1'default light behind Targets
		check=1
	end if

if check=1 then Lightcvup4.state=1:Lightcvup2.state=1:end If
if check=0 then Lightcvup4.state=0:Lightcvup2.state=0:end If


End Sub

 Sub sw29a_Timer()

  Select Case sw29Pos
        Case 0: sw29P.z=0
				If sw29Dir = 1 then
					sw29l = 1
					sw29a.TimerEnabled = 0
					lightchoice
				else
					sw29Dir = 0
					check = 0
					Lightcvup4.state=0:Lightcvup2.state=0
					sw29a.TimerEnabled = 1
					lightchoice
				end if
        Case 1: sw29P.z=0
        Case 2: sw29P.z=-4
        Case 3: sw29P.z=-8
        Case 4: sw29P.z=-12
        Case 5: sw29P.z=-16
        Case 6: sw29P.z=-20
        Case 7: sw29P.z=-24
        Case 8: sw29P.z=-28
        Case 9: sw29P.z=-32
        Case 10: sw29P.z=-36
        Case 11: sw29P.z=-40
        Case 12: sw29P.z=-44
        Case 13: sw29P.z=-48:sw29P.ReflectionEnabled = true
        Case 14: sw29P.z=-52:sw29P.ReflectionEnabled = false
				 If sw29Dir = 1 then
				 else
					sw29a.TimerEnabled = 0
			     end if


End Select
	If sw29Dir = 1 then
		If sw29pos>0 then sw29pos=sw29pos-1
	else
		If sw29pos<14 then sw29pos=sw29pos+1
	end if
  End Sub





 Sub sw30a_Timer()
  Select Case sw30Pos
        Case 0: sw30P.z=0
				 If sw30Dir = 1 then
					sw30l = 1
					sw30a.TimerEnabled = 0
					lightchoice
				 else
					sw30Dir = 0
					check = 0
					Lightcvup4.state=0:Lightcvup2.state=0
					sw30a.TimerEnabled = 1
					lightchoice
			     end if
        Case 1: sw30P.z=0
        Case 2: sw30P.z=-4
        Case 3: sw30P.z=-8
        Case 4: sw30P.z=-12
        Case 5: sw30P.z=-16
        Case 6: sw30P.z=-20
        Case 7: sw30P.z=-24
        Case 8: sw30P.z=-28
        Case 9: sw30P.z=-32
        Case 10: sw30P.z=-36
        Case 11: sw30P.z=-40
        Case 12: sw30P.z=-44
        Case 13: sw30P.z=-48:sw30P.ReflectionEnabled = true
        Case 14: sw30P.z=-52:sw30P.ReflectionEnabled = false
				 If sw30Dir = 1 then
				 else
					sw30a.TimerEnabled = 0
			     end if


End Select
	If sw30Dir = 1 then
		If sw30pos>0 then sw30pos=sw30pos-1
	else
		If sw30pos<14 then sw30pos=sw30pos+1
	end if
  End Sub


 Sub sw31a_Timer()
  Select Case sw31Pos
        Case 0: sw31P.z=0
				 If sw31Dir = 1 then
					sw31l = 1
					sw31a.TimerEnabled = 0
					lightchoice
				 else
					sw31Dir = 0
					check = 0
					Lightcvup4.state=0:Lightcvup2.state=0
					sw31a.TimerEnabled = 1
					lightchoice
			     end if
        Case 1: sw31P.z=0
        Case 2: sw31P.z=-4
        Case 3: sw31P.z=-8
        Case 4: sw31P.z=-12
        Case 5: sw31P.z=-16
        Case 6: sw31P.z=-20
        Case 7: sw31P.z=-24
        Case 8: sw31P.z=-28
        Case 9: sw31P.z=-32
        Case 10: sw31P.z=-36
        Case 11: sw31P.z=-40
        Case 12: sw31P.z=-44
        Case 13: sw31P.z=-48:sw31P.ReflectionEnabled = true
        Case 14: sw31P.z=-52:sw31P.ReflectionEnabled = false
				 If sw31Dir = 1 then
				 else
					sw31a.TimerEnabled = 0
			     end if



End Select
	If sw31Dir = 1 then
		If sw31pos>0 then sw31pos=sw31pos-1
	else
		If sw31pos<14 then sw31pos=sw31pos+1
	end if
  End Sub



'DT Subs
   Sub ResetDrops(Enabled)
		If Enabled Then
			sw29Dir = 1:sw30Dir = 1:sw31Dir = 1
			sw29l=1:sw30l=1:sw31l=1
			check=1
			sw29a.TimerEnabled = 1:sw30a.TimerEnabled = 1:sw31a.TimerEnabled = 1
			DTBank.DropSol_On
		Lightcvup1a01.state = 0:Lightcvup3aa01.state=0
		Lightcvup3a1.state = 0:Lightcvup1aa1.state=0
		Lightcvup1a0.state = 0:Lightcvup3aa0.state=0
		Lightcvup1a91.state = 0:Lightcvup3aa91.state=0
		Lightcvup1a901.state = 0:Lightcvup3aa901.state=0
		Lightcvup1a90.state = 0:Lightcvup3aa90.state=0
		Lightcvup1a9.state = 0:Lightcvup3aa9.state=0
		Lightcvup4.state=1:Lightcvup2.state=1
		End if
   End Sub



'=========================================================================
Sub SW46_Hit():vpmTimer.PulseSw 46:End Sub

Sub SW24_hit():sw24p.transx = -10:Me.TimerEnabled = 1:vpmTimer.PulseSw 24:End Sub
Sub SW24_Timer():sw24p.transx = 0:Me.TimerEnabled = 0:End Sub

Sub SW33_hit():sw33p.transx = -10:Me.TimerEnabled = 1:vpmTimer.PulseSw 33:End Sub
Sub SW33_Timer():sw33p.transx = 0:Me.TimerEnabled = 0:End Sub

Sub SW34_hit():sw34p.transx = -10:Me.TimerEnabled = 1:vpmTimer.PulseSw 34:End Sub
Sub SW34_Timer():sw34p.transx = 0:Me.TimerEnabled = 0:End Sub

Sub SW35_hit():sw35p.transx = -10:Me.TimerEnabled = 1:vpmTimer.PulseSw 35:End Sub
Sub SW35_Timer():sw35p.transx = 0:Me.TimerEnabled = 0:End Sub

Sub SW20_hit():sw20p.transx = -10:Me.TimerEnabled = 1:vpmTimer.PulseSw 20:End Sub
Sub SW20_Timer():sw20p.transx = 0:Me.TimerEnabled = 0:End Sub

Sub SW21_hit():sw21p.transx = -10:Me.TimerEnabled = 1:vpmTimer.PulseSw 21:End Sub
Sub SW21_Timer():sw21p.transx = 0:Me.TimerEnabled = 0:End Sub

Sub SW22_hit():sw22p.transx = -10:Me.TimerEnabled = 1:vpmTimer.PulseSw 22:End Sub
Sub SW22_Timer():sw22p.transx = 0:Me.TimerEnabled = 0:End Sub

Sub SW38_hit():sw38p.transx = -10:Me.TimerEnabled = 1:vpmTimer.PulseSw 38:End Sub
Sub SW38_Timer():sw38p.transx = 0:Me.TimerEnabled = 0:End Sub

Sub SW39_hit():sw39p.transx = -10:Me.TimerEnabled = 1:vpmTimer.PulseSw 39:End Sub
Sub SW39_Timer():sw39p.transx = 0:Me.TimerEnabled = 0:End Sub

Sub SW40_hit():sw40p.transx = -10:Me.TimerEnabled = 1:vpmTimer.PulseSw 40:End Sub
Sub SW40_Timer():sw40p.transx = 0:Me.TimerEnabled = 0:End Sub

Sub SW37_Hit():vpmTimer.PulseSw 37:SW37.isdropped = 1:SW37a.isdropped = 0:me.timerenabled = 1:End Sub
Sub SW37_Timer():SW37.isdropped = 0:SW37a.isdropped = 1:me.timerenabled = 0:End Sub

Sub SW56_Hit():vpmTimer.PulseSw 56:SubDown.Enabled = 1:SW56.isdropped = 1:SW56a.isdropped = 0:me.timerenabled = 1:End Sub
Sub SW56_Timer():SW56.isdropped = 0:SW56a.isdropped = 1:me.timerenabled = 0:End Sub

'*****GI Lights On
dim xx
Sub GIUpdate(enabled)

	if sw29Dir = 0 and sw30Dir = 1 and sw31Dir = 1 then
				If enabled Then
					For each xx in GI9:xx.State = 0:Next
				else
					For each xx in GI9:xx.State = 1:Next
				End If
	end if
	if sw29Dir = 0 and sw30Dir = 0 and sw31Dir = 1 then
				If enabled Then
					For each xx in GI90:xx.State = 0:Next
				else
					For each xx in GI90:xx.State = 1:Next
				End If
	end if
	if sw29Dir = 0 and sw30Dir = 0 and sw31Dir = 0 then
				If enabled Then
					For each xx in GI901:xx.State = 0:Next
				else
					For each xx in GI901:xx.State = 1:Next
				End If
	end if
	if sw29Dir = 0 and sw30Dir = 1 and sw31Dir = 0 then
				If enabled Then
					For each xx in GI91:xx.State = 0:Next
				else
					For each xx in GI91:xx.State = 1:Next
				End If
	end if
		if sw29Dir = 1 and sw30Dir = 0 and sw31Dir = 1 then
				If enabled Then
					For each xx in GI0:xx.State = 0:Next
				else
					For each xx in GI0:xx.State = 1:Next
				End If
	end if
		if sw29Dir = 1 and sw30Dir = 1 and sw31Dir = 0 then
				If enabled Then
					For each xx in GI1:xx.State = 0:Next
				else
					For each xx in GI1:xx.State = 1:Next
				End If
	end if
		if sw29Dir = 1 and sw30Dir = 0 and sw31Dir = 0 then
				If enabled Then
					For each xx in GI01:xx.State = 0:Next
				else
					For each xx in GI01:xx.State = 1:Next
				End If
	end if
		if sw29Dir = 1 and sw30Dir = 1 and sw31Dir = 1 then
				If enabled Then
					For each xx in GIempty:xx.State = 0:Next
				else
					For each xx in GIempty:xx.State = 1:Next
				End If
		end if
LightLeftSlingshot.state =0
LightRightSlingshot.state =0

End Sub





'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
    vpmTimer.PulseSw 53
    PlaySound SoundFX("right_slingshot",DOFContactors), 0, 1, 0.05, 0.05
	LightRightSlingshot.state =0
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
	'gi1.State = 0:Gi2.State = 0
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
	LightRightSlingshot.state =1
if RSLing1.Visible = 0 then LightRightSlingshot.state =0:end If
End Sub

Sub LeftSlingShot_Slingshot
    vpmTimer.PulseSw 52
    PlaySound SoundFX("left_slingshot",DOFContactors),0,1,-0.05,0.05
	LightLeftSlingshot.state =0
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
	'gi3.State = 0:Gi4.State = 0
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
	LightLeftSlingshot.state =1
if LSLing1.Visible = 0 then LightLeftSlingshot.state =0:end If
End Sub

Set MotorCallback=GetRef("UpdateMultipleLamps")

Sub UpdateMultipleLamps
    Llogo3dJPJ.ObjRotZ = LeftFlipper.CurrentAngle -121
    Rlogo3dJPJ.ObjRotZ = RightFlipper.CurrentAngle +121

If Controller.Switch(43) = True Then
Light1.State = 1
else
Light1.State = 0
End If

End Sub

Lights(1)=Array(L1,L1a)		'yellow lights

Lights(2)=Array(L2,L2z)
Lights(3)=Array(L3,L3z)
Lights(4)=Array(L4,L4z)
Lights(5)=Array(L5,L5z)
Lights(6)=Array(L6,L6z)
Lights(7)=Array(L7,L7z)
Lights(8)=Array(L8,L8z,L8r)
Lights(9)=Array(L9,L9z)
Lights(10)=Array(L10,L10z)
Lights(11)=Array(L11,L11z)
Lights(12)=Array(L12,L12z)
Lights(13)=Array(L13,L13z)
Lights(14)=Array(L14,L14z)
Lights(15)=Array(L15,L15z)

Set Lights(16)=L16'Set Lights(16)=L16 TEST

Lights(17)=Array(L17,L17z)
Lights(18)=Array(L18,L18z)
Lights(19)=Array(L19,L19z)
Lights(20)=Array(L20,L20z)
Lights(21)=Array(L21,L21z)
Lights(22)=Array(L22,L22z)

Lights(23)=Array(L23,L23A)	'yellow lights

Set Lights(24)=L24'Set Lights(24)=L24 TEST

Lights(25)=Array(L25,L25z)
Lights(26)=Array(L26,L26z)
Lights(27)=Array(L27,L27z)
Lights(28)=Array(L28,L28z)
Lights(29)=Array(L29,L29z)
Lights(30)=Array(L30,L30z)
Lights(31)=Array(L31,L31zz) 'binoculars
Lights(32)=Array(L32,L32zz) 'binoculars
Lights(33)=Array(L33,L33z,L33r)
Lights(34)=Array(L34,L34z)
Lights(35)=Array(L35,L35z)
Lights(36)=Array(L36,L36z)
Lights(37)=Array(L37,L37z)
Lights(38)=Array(L38,L38z)
Lights(39)=Array(L39,L39z)
Lights(40)=Array(L40,L40z)
Lights(41)=Array(L41,L41z)
Lights(42)=Array(L42,L42z)
Lights(43)=Array(L43,L43z)
Lights(44)=Array(L44,L44z)
Lights(45)=Array(L45,L45z)
Lights(46)=Array(L46,L46z)
Lights(47)=Array(L47,L47z)
Lights(48)=Array(L48,L48z,L48r)
Lights(49)=Array(L49,L49z)
Lights(50)=Array(L50,L50z)
Lights(51)=Array(L51,L51z)
Lights(52)=Array(L52,L52z)
Lights(53)=Array(L53,L53z)
Lights(54)=Array(L54,L54z)
Lights(55)=Array(L55,L55z)
Lights(56)=Array(L56,L56z)
Lights(57)=Array(L57,L57z)
Lights(58)=Array(L58,L58z)
Lights(59)=Array(L59,L59z)
Lights(60)=Array(L60,L60z)
Lights(61)=Array(L61,L61z)
Lights(62)=Array(L62,L62z)
Lights(63)=Array(L63,L63z,L63r)
set Lights(64)=L64			'yellow light

if check=1 then Lightcvup4.state=1:Lightcvup2.state=1:end If
if check=0 then Lightcvup4.state=0:Lightcvup2.state=0:end If


' *********************************************************************
'                      Yellow reflects over Rocks JPJ
' *********************************************************************

lightrocks.enabled = 1

Sub LightRocks_Timer()
select case L1.state
	case 0:RockLeft.Image = "RocherLoff"
	case 1:RockLeft.Image = "RocherLon"
end Select
select case L23.state or L64.state
	case 0:RockRight.Image = "RocherRoff"
	case 1:RockRight.Image = "RocherRon"
end Select
select case L23A.state
	case 0:RockCenter.Image = "RocherCoff"
	case 1:RockCenter.Image = "RocherCon"
end Select
End Sub

Function Vol2(ball1, ball2) ' Calculates the Volume of the sound based on the speed of two balls
    Vol2 = (Vol(ball1) + Vol(ball2) ) / 2
End Function

'************************************
' What you need to add to your table
'************************************

' a timer called CollisionTimer. With a fast interval, like from 1 to 10
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

' The Double For loop: This is a double cycle used to check the collision between a ball and the other ones.
' If you look at the parameters of both cycles, you’ll notice they are designed to avoid checking
' collision twice. For example, I will never check collision between ball 2 and ball 1,
' because I already checked collision between ball 1 and 2. So, if we have 4 balls,
' the collision checks will be: ball 1 with 2, 1 with 3, 1 with 4, 2 with 3, 2 with 4 and 3 with 4.

' Sum first the radius of both balls, and if the height between them is higher then do not calculate anything more,
' since the balls are on different heights so they can't collide.

' The next 3 lines calculates distance between xth and yth balls with the Pytagorean theorem,

' The first "If": Checking if the distance between the two balls is less than the sum of the radius of both balls,
' and both balls are not already colliding.

' Why are we checking if balls are already in collision?
' Because we do not want the sound repeting when two balls are resting closed to each other.

' Set the collision property of both balls to True, and we assign the number of the ball colliding

' Play the collide sound of your choice using the VOL, PAN and PITCH functions

' Last lines: If the distance between 2 balls is more than the radius of a ball,
' then there is no collision and then set the collision property of the ball to False (-1).


Sub Pins_Hit (idx)
	PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
	PlaySound "target"
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

Sub Rollovers_Hit (idx)
	PlaySound "rollover", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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

Sub RandomSoundFlipper()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub

Sub MRIN_Hit()
me.destroyball
mballout.enabled = 1
End Sub

Sub mballout_timer()
mrout.createball
mrout.kick 95, 10
PlaySound "ball_bounce"
me.enabled = 0
End Sub

Sub trigger1_hit()
ActiveBall.VelY = ActiveBall.VelY * 2
End Sub

Dim raiseballsw, raiseball, hasbeenhit

 Sub VUKKick(Enabled)
	if(enabled) AND hasbeenhit = 1 then
 		VUK.destroyball
        PlaySound "Kicker_enter_center"
		 'bsRTPop.balls = bsRTPop.balls  -1
		'VUK.DestroyBall
 		Set raiseball = VUK.CreateBall
 		raiseballsw = True
 		Vukraiseballtimer.Enabled = True
		'VUK.Enabled=TRUE
	end if
End Sub

 Sub Vukraiseballtimer_Timer()
 	If raiseballsw = True then
 		raiseball.z = raiseball.z + 10
 		If raiseball.z > 120 then
 			VUK.Kick 130, 10
			PlaySound SoundFX("popper",DOFContactors)
 			Set raiseball = Nothing
 			Vukraiseballtimer.Enabled = False
 			raiseballsw = False
		Controller.Switch(43) = 0
        hasbeenhit = 0
 		End If
 	End If
 End Sub

Sub Sinkit(enabled)
If enabled Then
shpos = 1
SinkShip.enabled = 1
end If


End Sub

Sub SinkShip_Timer()
Select Case shpos
Case 1:shpos = 2
TorpA1.state=1:TorpB5.state=1
Case 2:shpos = 3
TorpA1.state=0:TorpA2.state=1:TorpB1.state=0:TorpB2.state=1:
Case 3:shpos = 4
TorpA2.state=0:TorpA3.state=1:TorpB2.state=0:TorpB3.state=1:
Case 4:shpos = 5
TorpA3.state=0:TorpA4.state=1:TorpB3.state=0:TorpB4.state=1:
Case 5:shpos = 6
TorpA4.state=0:TorpA5.state=1:TorpB4.state=0:TorpB5.state=1:
Case 6:Sink1.isdropped = 1:Sink2.isdropped = 0:shpos = 7
TorpA5.state=0:TorpB5.state=0
Case 7:Sink2.isdropped = 1:Sink3.isdropped = 0:shpos = 8
Case 8:Sink3.isdropped = 1:Sink4.isdropped = 0:shpos = 9
Case 9:Sink4.isdropped = 1:Sink5.isdropped = 0:shpos = 10
Case 10:Sink5.isdropped = 1:Sink1.isdropped = 0:me.enabled = 0
End Select
End Sub


'***********
'   Submarine Left Right Nudge routine
'**********
Dim Rot
Dim SbmRotaYY
Dim SbmLR

SbmLR = 0
'Rot = 0
SbmRotaYY = 90

SubLR.enabled = 1


Sub SubLR_Timer()
:debug.print "on est dans le SubLR_Timer"
	if SbmLR= -1 then Rot = 1:debug.print "OK pour le nudge"end if
	if SbmLR= 1 then Rot = 10:end if
	UpdateSbmLR

  End Sub


  Sub UpDateSbmLR
	select case Rot
'case 0:end Select
case 1:SbmRotaYY = 82:Rot = 2
Submarine3.RotZ = SbmRotaYY
case 2:Rot = 3:SbmRotaYY = 78:debug.print SbmRotaYY
Submarine3.RotZ = SbmRotaYY
case 3:Rot = 4:SbmRotaYY = 74:
Submarine3.RotZ = SbmRotaYY
case 4:Rot = 5:SbmRotaYY = 80:debug.print SbmRotaYY
Submarine3.RotZ = SbmRotaYY
case 5:Rot = 6:SbmRotaYY = 86:
Submarine3.RotZ = SbmRotaYY
case 6:Rot = 7:SbmRotaYY = 92:debug.print SbmRotaYY
Submarine3.RotZ = SbmRotaYY
case 7:Rot = 8:SbmRotaYY = 95:
Submarine3.RotZ = SbmRotaYY
case 8:Rot = 9:SbmRotaYY = 92:debug.print SbmRotaYY
Submarine3.RotZ = SbmRotaYY
case 9:SbmRotaYY = 90:Rot = 0:debug.print SbmRotaYY
Submarine3.RotZ = SbmRotaYY
case 10:Rot = 11:SbmRotaYY = 98:UpDateSbmLRAction
case 11:Rot = 12:SbmRotaYY = 102:UpDateSbmLRAction
case 12:Rot = 13:SbmRotaYY = 106:UpDateSbmLRAction
case 13:Rot = 14:SbmRotaYY = 100:UpDateSbmLRAction
case 14:Rot = 15:SbmRotaYY = 94:UpDateSbmLRAction
case 15:Rot = 16:SbmRotaYY = 88:UpDateSbmLRAction
case 16:Rot = 17:SbmRotaYY = 85:UpDateSbmLRAction
case 17:Rot = 18:SbmRotaYY = 88:UpDateSbmLRAction
case 18:SbmRotaYY = 90:UpDateSbmLRAction:Rot = 0
	End select

  End Sub

Sub UpDateSbmLRAction
	Submarine3.RotZ = SbmRotaYY
	SubMarine2.RotZ = SbmRotaYY
	Submarine1.RotZ = SbmRotaYY
debug.print SbmRotaYY
End Sub

'***********
'   Submarine3 routine
'**********
Dim SbmDir
Dim SbmTransX, SbmIntensity
SbmDir = -1
SbmTransX = 0
SbmIntensity = 0


  Sub SubDown_Timer()
	if SbmDir= -1 then SbmTransX = SbmTransX - 3:SbmIntensity = SbmIntensity + 1:end if
	if SbmDir= 1 then SbmTransX = SbmTransX + 3:SbmIntensity = SbmIntensity - 1:end If
	UpdateSbm
	If SbmTransX >= 0 then
		SbmTransX = 0:SubDown.Enabled = 0: SbmDir=-1
	end if
	If SbmTransX <= -39 then
		SbmTransX = -39:SubDown.Enabled = 0: SbmDir=1
	end if
  End Sub


  Sub UpDateSbm
	Submarine3.TransX = SbmTransX
	LightSubmarine3.intensity() = SbmIntensity

  End Sub


'***********
'   Submarine2 routine
'**********
Dim SbmDir2, SbmIntensity2
Dim SbmTransX2
SbmDir2 = -1
SbmTransX2 = 0
SbmIntensity2 = 0


  Sub SubDown2_Timer()
	if SbmDir2= -1 then SbmTransX2 = SbmTransX2 - 3:SbmIntensity2 = SbmIntensity2 + 1:end if
	if SbmDir2= 1 then SbmTransX2 = SbmTransX2 + 3:SbmIntensity2 = SbmIntensity2 - 1:end If
	UpdateSbm2
	If SbmTransX2 >= 0 then
		SbmTransX2 = 0:SubDown2.Enabled = 0: SbmDir2=-1
	end if
	If SbmTransX2 <= -39 then
		SbmTransX2 = -39:SubDown2.Enabled = 0: SbmDir2=1
	end if
  End Sub


  Sub UpDateSbm2
	Submarine2.TransX = SbmTransX2
	LightSubmarine2.intensity() = SbmIntensity2

  End Sub


'***********
'   Submarine1 routine
'**********
Dim SbmDir1, SbmIntensity1
Dim SbmTransX1
SbmDir1 = -1
SbmTransX1 = 0
SbmIntensity1 = 0


  Sub SubDown1_Timer()
	if SbmDir1= -1 then SbmTransX1 = SbmTransX1 - 3:SbmIntensity1 = SbmIntensity1 + 1:end if
	if SbmDir1= 1 then SbmTransX1 = SbmTransX1 + 3:SbmIntensity1 = SbmIntensity1 - 1:end If
	UpdateSbm1
	If SbmTransX1 >= 0 then
		SbmTransX1 = 0:SubDown1.Enabled = 0: SbmDir1=-1
	end if
	If SbmTransX1 <= -39 then
		SbmTransX1 = -39:SubDown1.Enabled = 0: SbmDir1=1
	end if
  End Sub


  Sub UpDateSbm1
	Submarine1.TransX = SbmTransX1
	LightSubmarine1.intensity() = SbmIntensity1

  End Sub




Sub DOF(dofevent, dofstate)
	'If cController = 3 Then
		If dofstate = 2 Then
			Controller.B2SSetData dofevent, 1:Controller.B2SSetData dofevent, 0
		Else
			Controller.B2SSetData dofevent, dofstate
		End If
	'End If
End Sub

'-------------------------------

Sub Table1_Exit
Controller.Stop
End Sub

'************************  Shake Scripting  *************************
'* Code adapted from Cirqus Voltaire from Dozer, Ninuzzu and Nfozzy *
'* Big thank you for this fabulous method                           *
'********************************************************************
 Dim mMagnet, cBall, pMod, rmmod

 Set mMagnet = new cvpmMagnet
 With mMagnet
	.InitMagnet WobbleMagnet, 1.5
	.Size = 100
	.CreateEvents mMagnet
	.MagnetOn = True
 End With
 WobbleInit

 Sub RMShake
	cball.velx = cball.velx + rmball.velx*pMod
	cball.vely = cball.vely + rmball.vely*pMod
 End Sub

Sub RMShake2
	cball.velx = cball.velx + activeball.velx*.05
	cball.vely = cball.vely + activeball.vely*.05
 End Sub

Sub RM_Make_Hit()
RMShake2
End Sub

'Includes stripped down version of my reverse slope scripting for a single ball
 Dim ngrav, ngravmod, pslope, nslope, slopemod
 Sub WobbleInit
	pslope = Table1.SlopeMin +((Table1.SlopeMax - Table1.SlopeMin) * Table1.GlobalDifficulty)
	nslope = pslope
	slopemod = pslope + nslope
	ngravmod = 60/aWobbleTimer.interval
	ngrav = slopemod * .0905 * Table1.Gravity / ngravmod
	pMod = .15					'percentage of hit power transfered to captive wobble ball
	Set cBall = ckicker.createball:cball.image = "blank":ckicker.Kick 0,0:mMagnet.addball cball
	aWobbleTimer.enabled = 1
 End Sub

 Sub aWobbleTimer_Timer
	'BallShake.Enabled = RMBallInMagnet
	cBall.Vely = cBall.VelY-ngrav					'modifier for slope reversal/cancellation
	rmmod = (SubMarine3.z+265.5)/265*.6				'.4 is a 40% modifier for ratio of ball movement to head movement
	SubMarine3.rotx = 90+(ckicker.y - cball.y)*rmmod
	SubMarine3.rotz = 90 + (cball.x - ckicker.x)*rmmod
	SubMarine2.rotx = 90+(ckicker.y - cball.y)*rmmod
	SubMarine2.rotz = 90 + (cball.x - ckicker.x)*rmmod
	SubMarine1.rotx = 90+(ckicker.y - cball.y)*rmmod
	SubMarine1.rotz = 90 + (cball.x - ckicker.x)*rmmod
 End Sub

 'Sub BallShake_Timer
'	If Not IsEmpty(RMMagBall) Then
'		RMMagBall.y = RMMagnetkicker.y - dsin(SubMarine3.rotx)*265.5
'		RMMagBall.x = RMMagnetkicker.x + dsin(SubMarine3.roty)*265.5
'	End If
 'End Sub

'*************  End Shake Scripting  ****************

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
'    JP's VP10 Collision & Rolling Sounds
'*****************************************

Const tnob = 6 ' total number of balls
ReDim rolling(tnob)
ReDim collision(tnob)
Initcollision

Sub Initcollision
    Dim i
    For i = 0 to tnob
        collision(i) = -1
        rolling(i) = False
    Next
End Sub

Sub CollisionTimer_Timer()
    Dim BOT, B, B1, B2, dx, dy, dz, distance, radii
    BOT = GetBalls

    ' rolling

	For B = UBound(BOT) +1 to tnob
        rolling(b) = False
        StopSound("fx_ballrolling" & b)
	Next

    If UBound(BOT) = -1 Then Exit Sub

    For B = 0 to UBound(BOT)
        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*15, Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0
        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If
    Next

    'collision

    If UBound(BOT) < 1 Then Exit Sub

    For B1 = 0 to UBound(BOT)
        For B2 = B1 + 1 to UBound(BOT)
            dz = INT(ABS((BOT(b1).z - BOT(b2).z) ) )
            radii = BOT(b1).radius + BOT(b2).radius
			If dz <= radii Then

            dx = INT(ABS((BOT(b1).x - BOT(b2).x) ) )
            dy = INT(ABS((BOT(b1).y - BOT(b2).y) ) )
            distance = INT(SQR(dx ^2 + dy ^2) )

            If distance <= radii AND (collision(b1) = -1 OR collision(b2) = -1) Then
                collision(b1) = b2
                collision(b2) = b1
                PlaySound("fx_collide"), 0, Vol2(BOT(b1), BOT(b2)), Pan(BOT(b1)), 0, Pitch(BOT(b1)), 0, 0
            Else
                If distance > (radii + 10)  Then
                    If collision(b1) = b2 Then collision(b1) = -1
                    If collision(b2) = b1 Then collision(b2) = -1
                End If
            End If
			End If
        Next
    Next
End Sub


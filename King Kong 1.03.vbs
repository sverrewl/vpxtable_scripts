Option Explicit
Randomize

Const cGameName="kiko_a10"


On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

LoadVPM "01560000", "de.VBS", 3.10

'Variables
Dim xx
Dim Bump1, Bump2, Bump3
Dim bsTrough
Dim bsMissileKicker
Dim bsRadarEject
Dim vLock
Dim bsVuk

Const UseSolenoids = 1
Const UseLamps = 0
Const UseSync = 1
Const HandleMech = 0 

'Standard Sounds
Const SSolenoidOn = "Solenoid"
Const SSolenoidOff = ""
Const SCoin = "Coin_In"

'Table Init
Sub Table1_Init
	vpmInit Me

	With Controller
		.GameName = cGameName
		.SplashInfoLine = "King Kong, Data East 1990"
		.HandleMechanics = 0
		.HandleKeyboard = 0
		.ShowDMDOnly = 1
		.ShowFrame = 0
		.ShowTitle = 0
		.Hidden = 0
	End With
	On Error Resume Next

	Controller.Run
	If Err Then MsgBox Err.Description
	On Error Goto 0

	'Nudging
	vpmNudge.TiltSwitch=6
	vpmNudge.Sensitivity=1
	vpmNudge.TiltObj=Array(Bumper1b,Bumper2b,Bumper3b,LeftSlingshot,RightSlingshot)

	'**Trough
	Set bsTrough = New cvpmBallStack
	With bsTrough
		.InitSw 0, 13, 12, 11, 0, 0, 0, 0
		.InitKick BallRelease, 90, 4
		.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
		.Balls = 3
	End With

	'**Kicker
	Set bsMissileKicker = New cvpmBallStack
	bsMissileKicker.InitSaucer missilekicker, 30, 90, 35
	bsMissileKicker.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("solenoid",DOFContactors)

	'**RadarEject
	Set bsRadarEject = New cvpmBallStack
	bsRadarEject.InitSaucer radareject, 45, 90, 3
	bsRadarEject.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)

	' Visible Lock
	Set vLock = New cvpmVLock2
	With vLock
		.InitVLock Array(sw38, sw39, sw40), Array(sw38k, sw39k, sw40k), Array(38, 39, 40)
		.InitSnd SoundFX("Solenoid",DOFContactors), SoundFX("Solenoid",DOFContactors)
		.ExitDir = 194
		.ExitForce = .1
		.CreateEvents "vLock"
	End With
'
'    ' Vuk
'    Set bsVuk = New cvpmBallStack
'    With bsVuk
'        .InitSw 0, 29, 0, 0, 0, 0, 0, 0
'        .InitKick sw29a, 160, 16
'        .InitExitSnd "scoopexit", "Solenoid"
'    End With




	'**Main Timer init
	PinMAMETimer.Enabled = 1

End Sub

dim speed
speed = 42
'*****Keys
Sub Table1_KeyDown(ByVal keycode)



	If Keycode = LeftFlipperKey then 
		SolLFlipper true
	End If
	If Keycode = RightFlipperKey then 
		SolRFlipper true
	End If
	If keycode = PlungerKey Then PlungerPull
	If keycode = LeftTiltKey Then PlaySound SoundFX("Nudge_Left",0): 'LeftNudge 80, 1, 20
	If keycode = RightTiltKey Then PlaySound SoundFX("Nudge_Right",0): 'RightNudge 280, 1, 20
	If keycode = CenterTiltKey Then PlaySound SoundFX("Nudge_Forward",0): 'CenterNudge 0, 1, 25
	If vpmKeyDown(keycode) Then Exit Sub 
End Sub

Sub Table1_KeyUp(ByVal keycode)
	If vpmKeyUp(keycode) Then Exit Sub
 	If Keycode = LeftFlipperKey then 
		SolLFlipper false
	End If
 	If Keycode = RightFlipperKey then 
		SolRFlipper False
	End If
    If keycode = PlungerKey Then PlungerRelease
End Sub

Sub PlungerPull
	Plunger.Pullback
	PlaySound "fx_plungerpull"
'	kicker2.createball: kicker2.kick -10, speed, 0
'	if speed = 30 then 
'		speed = 40
'	elseif speed = 40 then 
'		speed = 50
'	else 
'		speed = 60
'	end if
End Sub

Sub PlungerRelease
	Plunger.Fire
	PlaySound "fx_plunger2"
End Sub

Sub Table1_Paused:Controller.Pause = 1:End Sub
Sub Table1_unPaused:Controller.Pause = 0:End Sub
Sub Table1_exit():Controller.Stop:	End sub

'*****************************************
' Dropwall Init
''*****************************************


'*******************************
'' Rollovers
''*******************************
Sub sw14_Hit	: Controller.Switch(14) = 1: Playsound "rollover": End Sub 'Shooter Lane
Sub sw14_UnHit	: Controller.Switch(14) = 0: End Sub
Sub sw17_Hit:     Controller.Switch(17) = 1: Playsound "rollover": End Sub  'Left Outlane
Sub sw17_UnHit:   Controller.Switch(17) = 0: End Sub
Sub sw18_Hit:     Controller.Switch(18) = 1: Playsound "rollover": End Sub  'Left Return
Sub sw18_UnHit:   Controller.Switch(18) = 0: End Sub
Sub sw19_Hit:     Controller.Switch(19) = 1: Playsound "rollover": End Sub  'Right Outlane
Sub sw19_UnHit:   Controller.Switch(19) = 0: End Sub
Sub sw20_Hit:     Controller.Switch(20) = 1: Playsound "rollover": End Sub  'Right Return
Sub sw20_UnHit:   Controller.Switch(20) = 0: End Sub
Sub sw23_Hit:     Controller.Switch(23) = 1: Playsound "rollover": End Sub  'Right Return
Sub sw23_UnHit:   Controller.Switch(23) = 0: End Sub
Sub sw25_Hit   : Controller.Switch(25) = 1: Playsound "rollover": End Sub 'A
Sub sw25_UnHit : Controller.Switch(25) = 0: End Sub
Sub sw26_Hit   : Controller.Switch(26) = 1: Playsound "rollover": End Sub	'P
Sub sw26_UnHit : Controller.Switch(26) = 0: End Sub
Sub sw27_Hit   : Controller.Switch(27) = 1: Playsound "rollover": End Sub	'E
Sub sw27_UnHit : Controller.Switch(27) = 0: End Sub
Sub sw31_Hit:   Controller.Switch(31) = 1: Playsound "rollover": End Sub  'Loop Left
Sub sw31_UnHit: Controller.Switch(31) = 0: End Sub
Sub sw32_Hit:   Controller.Switch(32) = 1: Playsound "rollover": End Sub  'Loop Right
Sub sw32_UnHit: Controller.Switch(32) = 0: End Sub

'Sub sw17_Hit:   PrimRolloverHit   17, PrimSw17: End Sub  'Left Outlane
'Sub sw17_UnHit: PrimRolloverUnHit 17, PrimSw17: End Sub
'Sub sw18_Hit:   PrimRolloverHit   18, PrimSw18: End Sub  'Left Return
'Sub sw18_UnHit: PrimRolloverUnHit 18, PrimSw18: End Sub
'Sub sw19_Hit:   PrimRolloverHit   19, PrimSw19: End Sub  'Right Outlane
'Sub sw19_UnHit: PrimRolloverUnHit 19, PrimSw19: End Sub
'Sub sw20_Hit:   PrimRolloverHit   20, PrimSw20: End Sub  'Right Return
'Sub sw20_UnHit: PrimRolloverUnHit 20, PrimSw20: End Sub
'Sub sw23_Hit:   PrimRolloverHit   23, PrimSw23: End Sub
'Sub sw23_UnHit: PrimRolloverUnHit 23, PrimSw23: End Sub
'Sub sw25_Hit:   PrimRolloverHit   25, PrimSw25: End Sub 'A
'Sub sw25_UnHit: PrimRolloverUnHit 25, PrimSw25: End Sub
'Sub sw26_Hit:   PrimRolloverHit   26, PrimSw26: End Sub	'P
'Sub sw26_UnHit: PrimRolloverUnHit 26, PrimSw26: End Sub
'Sub sw27_Hit:   PrimRolloverHit   27, PrimSw27: End Sub	'E
'Sub sw27_UnHit: PrimRolloverUnHit 27, PrimSw27: End Sub

Sub Gate3_Hit:Controller.Switch(28) = 1::End Sub  		'APE lane ramp (using gate instead of switch)
Sub Gate3_UnHit:Controller.Switch(28) = 0:End Sub
'Sub sw31_Hit:   PrimRolloverHit   31, PrimSw31: End Sub  'Loop Left
'Sub sw31_UnHit: PrimRolloverUnHit 31, PrimSw31: End Sub
'Sub sw32_Hit:   PrimRolloverHit   32, PrimSw32: End Sub  'Loop Right
'Sub sw32_UnHit: PrimRolloverUnHit 32, PrimSw32: End Sub
'
''Targets
Sub sw41_Hit:   PrimStandupTgtHit   41, Sw41, PrimSw41: End Sub 'multi-ball standups
Sub sw41_Timer: PrimStandupTgtMove  41, Sw41, PrimSw41: End Sub
Sub sw42_Hit:   PrimStandupTgtHit   42, Sw42, PrimSw42: End Sub
Sub sw42_Timer: PrimStandupTgtMove  42, Sw42, PrimSw42: End Sub
Sub sw43_Hit:   PrimStandupTgtHit   43, Sw43, PrimSw43: End Sub
Sub sw43_Timer: PrimStandupTgtMove  43, Sw43, PrimSw43: End Sub


''Cubs, Bears, Bulls
Sub sw49_Hit:   PrimStandupTgtHit   49, Sw49, PrimSw49: End Sub
Sub sw49_Timer: PrimStandupTgtMove  49, Sw49, PrimSw49: End Sub
Sub sw50_Hit:   PrimStandupTgtHit   50, Sw50, PrimSw50: End Sub
Sub sw50_Timer: PrimStandupTgtMove  50, Sw50, PrimSw50: End Sub
Sub sw51_Hit:   PrimStandupTgtHit   51, Sw51, PrimSw51: End Sub
Sub sw51_Timer: PrimStandupTgtMove  51, Sw51, PrimSw51: End Sub
'
''Tower
Sub sw33_Hit:   PrimStandupTgtHit   33, Sw33, PrimSw33: End Sub
Sub sw33_Timer: PrimStandupTgtMove  33, Sw33, PrimSw33: End Sub
Sub sw34_Hit:   PrimStandupTgtHit   34, Sw34, PrimSw34: End Sub
Sub sw34_Timer: PrimStandupTgtMove  34, Sw34, PrimSw34: End Sub
Sub sw35_Hit:   PrimStandupTgtHit   35, Sw35, PrimSw35: End Sub
Sub sw35_Timer: PrimStandupTgtMove  35, Sw35, PrimSw35: End Sub
Sub sw36_Hit:   PrimStandupTgtHit   36, Sw36, PrimSw36: End Sub
Sub sw36_Timer: PrimStandupTgtMove  36, Sw36, PrimSw36: End Sub
Sub sw37_Hit:   PrimStandupTgtHit   37, Sw37, PrimSw37: End Sub
Sub sw37_Timer: PrimStandupTgtMove  37, Sw37, PrimSw37: End Sub

Sub Bumper1b_Hit
	vpmTimer.PulseSw 46 'Left
	Playsound SoundFX("fx_bumper2",DOFContactors)
	lbumper1.state = 1
	Me.TimerEnabled = 1
End Sub
Sub Bumper1b_Timer
	lbumper1.state = 0
	Me.TimerEnabled = 0
End Sub

Sub Bumper2b_Hit
	vpmTimer.PulseSw 47 'Center
	Playsound SoundFX("fx_bumper3",DOFContactors)
	lbumper2.state = 1
	Me.TimerEnabled = 1
End Sub
Sub Bumper2b_Timer
	lbumper2.state = 0
	Me.TimerEnabled = 0
End Sub
Sub Bumper3b_Hit
	vpmTimer.PulseSw 48 'Right
	Playsound SoundFX("fx_bumper4",DOFContactors)
	lbumper3.state = 1
	Me.TimerEnabled = 1
End Sub
Sub Bumper3b_Timer
	lbumper3.state = 0
	Me.TimerEnabled = 0
End Sub

''*************BumperLights Subs 
'  'Sub B1On
'	 	'For each xx in Bumper1On:xx.IsDropped = 0:Next
'		'Bumper1Cap.State = 1:Bumper1a.State=1
'  'End Sub
'
'  'Sub B2On
'	 	'For each xx in Bumper2On:xx.IsDropped = 0:Next
'		'Bumper2Cap.State = 1:Bumper2a.State=1
'  'End Sub
'
'  'Sub B3On
'	 	'For each xx in Bumper3On:xx.IsDropped = 0:Next
'		'Bumper3Cap.State = 1:Bumper3a.State=1
'  'End Sub
'
'  'Sub B1OFF
'	 	'For each xx in Bumper1On:xx.IsDropped = 1:Next
'		'Bumper1Cap.State = 0:Bumper1a.State=0
'  'End Sub
'
'  'Sub B2OFF
'	 	'For each xx in Bumper2On:xx.IsDropped = 1:Next
'		'Bumper2Cap.State = 0:Bumper2a.State=0
'  'End Sub
'
'  'Sub B3OFF
'	 	'For each xx in Bumper3On:xx.IsDropped = 1:Next
'		'Bumper3Cap.State = 0:Bumper3a.State=0
'  'End Sub
''*************BumperLights Subs
'
''BUMPER 1
''******************************************
'Sub Bumper1a_Hit()
'	PlaySound"Bumper17"
'	vpmTimer.PulseSw 46 'Left
'	b1cap.Image = "bumper-cap-t2lit"
'	Bumper1a.TimerEnabled = True
'	b1ringctr = 0:b1ringdir = -6
'End Sub
'
'Dim b1ringctr,b1ringdir
'Sub Bumper1a_Timer()
'	If b1ringctr <= -40 then b1ringdir = 4
'	b1ringctr = b1ringctr + b1ringdir
'	b1ring.z = b1ringctr
'	If b1ringctr >= 0 then Bumper1a.TimerEnabled = False:b1cap.Image = "bumper-cap-t2"
'End Sub
'
''BUMPER 2
''******************************************
'Sub Bumper2a_Hit()
'	PlaySound"Bumper17"
'	vpmTimer.PulseSw 48	'Right
'	b2cap.Image = "bumper-cap-t2lit"
'	Bumper2a.TimerEnabled = True
'	b2ringctr = 0:b2ringdir = -6
'End Sub
'
'Dim b2ringctr,b2ringdir
'Sub Bumper2a_Timer()
'	If b2ringctr <= -40 then b2ringdir = 4
'	b2ringctr = b2ringctr + b2ringdir
'	b2ring.z = b2ringctr
'	If b2ringctr >= 0 then Bumper2a.TimerEnabled = False:b2cap.Image = "bumper-cap-t2"
'End Sub
'
''BUMPER 3
''******************************************
'Sub Bumper3a_Hit()
'	PlaySound"Bumper17"
'	vpmTimer.PulseSw 47 'Center
'	b3cap.Image = "bumper-cap-t2lit"
'	Bumper3a.TimerEnabled = True
'	b3ringctr = 0:b3ringdir = -6
'End Sub
'
'Dim b3ringctr,b3ringdir
'Sub Bumper3a_Timer()
'	If b3ringctr <= -40 then b3ringdir = 4
'	b3ringctr = b3ringctr + b3ringdir
'	b3ring.z = b3ringctr
'	If b3ringctr >= 0 then Bumper3a.TimerEnabled = False:b3cap.Image = "bumper-cap-t2"
'End Sub
' 


'Solenoids
SolCallBack(16) = "bsMissileKicker.SolOut"
SolCallBack(25) = "bsTrough.SolIn"
SolCallBack(26) = "bsTrough.SolOut"
SolCallBack(27) = "vLock.SolExit"
SolCallBack(28) = "bsRadarEject.SolOut"
'SolCallBack(29) = "bsVuk.SolOut"  'VUK VERSION
'SolCallBack(30) = "dtDrop.SolDropUp"  'TARGET BANK NOT USED IN THIS VERSION
SolCallBack(32) = "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"    'Need to confirm
SolCallBack(1) = "SetLamp 101," 'King Shooter
SolCallBack(2) = "SetLamp 102," 'King Shooter
SolCallback(3) = "SetLamp 103," 'Ape Ramp BG Skull
SolCallback(4) = "SetLamp 104," 'BG NDSRS Shooter
SolCallback(5) = "SetLamp 105," 'BG Tower 25k
SolCallback(6) = "SetLamp 106," 'BG Tower 50k
SolCallback(7) = "SetLamp 107," 'BG Tower 100k
SolCallback(8) = "SetLamp 108," 'BG Tower Extra Ball
SolCallback(9) = "SetLamp 109," 'BG Tower Million
'SolCallback(10) = NOT USED
SolCallback(11) = "solGI " ' General Illum
SolCallback(12) = "solLamp112" 'BG Girl Missile
SolCallback(13) = "SetLamp 113," 'BG Kong Explode
SolCallback(14) = "SetLamp 114," 'Tower Flasher
SolCallback(15) = "SetLamp 115," 'Skull Island
'SolCallback(22) = NOT USED
'SolCallback(23) = NOT USED
'SolCallback(24) = NOT USED

Sub solGI (enabled)
	if (enabled) Then
		SetLamp 111, 0
	else
		SetLamp 111, 1
	End If
End Sub

Sub solLamp112 (enabled)
	SetLamp 112, enabled
	SetLamp 199, enabled
End Sub
'***********************************************
   'Flipper Subs
    'Based on jp's dropwall flippers
  
   'Flipper Subs
       'SolCallback(sLRFlipper) = "SolRFlipper"
       'SolCallback(sLLFlipper) = "SolLFlipper"

   
       Sub SolLFlipper(Enabled)
           If Enabled Then
				   PlaySound SoundFX("LFlip",DOFContactors):LeftFlipper.RotateToEnd
		   Else
				   PlaySound SoundFX("LFlipD",DOFContactors)
				LeftFlipper.RotateToStart
           End If
       End Sub
     
       Sub SolRFlipper(Enabled)
           If Enabled Then
                   PlaySound SoundFX("LFlip",DOFContactors):RightFlipper.RotateToEnd:RightFlipper2.RotateToEnd
		   Else
				   PlaySound SoundFX("LFlipD",DOFContactors) 
				RightFlipper.RotateToStart:RightFlipper2.RotateToStart
           End If
       End Sub

 'Drains and Kickers
   Sub Drain_Hit():PlaySound "Drain":vpmTimer.PulseSw 10:bsTrough.Addball Me:End Sub
   Sub BallRelease_UnHit():End Sub

'Missile Kicker Lock
Sub MissileKicker_Hit():PlaySound "kicker_enter_center": bsMissileKicker.AddBall 0:End Sub

'Radar Eject Lock
Sub RadarEject_Hit():PlaySound "kicker_enter_center": bsRadarEject.AddBall 0:End Sub

'Cage Tunnel
dim theBall
Sub CageKicker_Hit():set theball = activeball:activeball.z = -100:PlaySound "Drain":CageReturnTimer.Enabled = 1:End Sub
Sub CageReturnTimer_Timer(): : theball.X = 376: theball.Y = 10.25: theball.Z = 80: cagekicker.kick 180, 1:CageReturnTimer.Enabled = 0:Playsound "Drop_Left": End Sub

'Ape Upper PF
'Sub ApeKicker1_Hit():PlaySound "Drain":ApeKickerTimer.Enabled = 1:End Sub
'Sub ApeKickerTimer_Timer(): : ApeMoveBall:ApeKicker2.kick 180, 1: ApeKickerTimer.Enabled = 0: End Sub
'Sub ApeMoveBall()::ApeKicker1.DestroyBall:ApeKicker2.CreateBall:ApeReturnTrigger.enabled = 1:End Sub
'Sub ApeReturnTrigger_Hit() :   PlaySound "Drop_left":ApeReturnTrigger.enabled = 0: End sub   
Sub ApeReturnTrigger_Hit() :  PlaySound "Drop_right": End sub   
'
'VUK Loc
Sub sw29_Hit():	vpmTimer.PulseSw 29	:End Sub


'***Slings and rubbers
  ' Slings
 Dim LStep, RStep
 
 Sub LeftSlingShot_Slingshot
' 	For each xx in LHammerA:xx.IsDropped = 0:Next
 	PlaySound SoundFX("slingshot",DOFContactors):LStep = 0:Me.TimerEnabled = 1
	vpmTimer.PulseSw 21
  End Sub
 
 Sub LeftSlingShot_Timer
     Select Case LStep
         Case 0:lkongout.enabled=1
         Case 1: 'pause
'         Case 2:For each xx in LHammerA:xx.IsDropped = 1::Next
' 				For each xx in LHammerB:xx.IsDropped = 0:Next
'         Case 3:For each xx in LHammerB:xx.IsDropped = 1:Next
' 				For each xx in LHammerC:xx.IsDropped = 0:Next
'         Case 4:For each xx in LHammerC:xx.IsDropped = 1:Next
 		 Case 5:Me.TimerEnabled = 0
     End Select
     LStep = LStep + 1
 End Sub
' 
'
'
'
'
'
'
Sub RightSlingShot_Slingshot
' 	For each xx in RHammerA:xx.IsDropped = 0:Next
 	PlaySound SoundFX("slingshot",DOFContactors):RStep = 0:Me.TimerEnabled = 1
	vpmTimer.PulseSw 22
End Sub
 
 Sub RightSlingShot_Timer
     Select Case RStep
         Case 0:rkongout.enabled=1
'         Case 1: 'pause
'         Case 2:For each xx in RHammerA:xx.IsDropped = 1:Next
' 				For each xx in RHammerB:xx.IsDropped = 0:Next
'         Case 3:For each xx in RHammerB:xx.IsDropped = 1:Next
' 				For each xx in RHammerC:xx.IsDropped = 0:Next
'         Case 4:For each xx in RHammerC:xx.IsDropped = 1:Next
 		Case 5: Me.TimerEnabled = 0
     End Select
     RStep = RStep + 1
 End Sub



  
   'Bumpers
'      Sub Bumper1b_Hit::PlaySound "bumper":bump1 = 1:Me.TimerEnabled = 1:End Sub
'     
'       Sub Bumper1b_Timer()
'           Select Case bump1
'               Case 1:Ring1a.IsDropped = 0:Ring1c.IsDropped = 1:bump1 = 2
'               Case 2:Ring1b.IsDropped = 0:Ring1a.IsDropped = 1:bump1 = 3
'               Case 3:Ring1c.IsDropped = 0:Ring1b.IsDropped = 1:bump1 = 4
'               Case 4:Ring1c.IsDropped = 0:Me.TimerEnabled = 0
'           End Select
'       End Sub
' 
'      Sub Bumper2b_Hit::PlaySound "bumper":bump2 = 1:Me.TimerEnabled = 1:End Sub
'     
'       Sub Bumper2b_Timer()
'           Select Case bump2
'               Case 1:Ring2a.IsDropped = 0:Ring2c.IsDropped = 1:bump2 = 2
'               Case 2:Ring2b.IsDropped = 0:Ring2a.IsDropped = 1:bump2 = 3
'               Case 3:Ring2c.IsDropped = 0:Ring2b.IsDropped = 1:bump2 = 4
'               Case 4:Ring2c.IsDropped = 0:Me.TimerEnabled = 0
'           End Select
'       End Sub
' 
'      Sub Bumper3b_Hit::PlaySound "bumper":bump3 = 1:Me.TimerEnabled = 1:End Sub
'     
'       Sub Bumper3b_Timer()
'           Select Case bump3
'               Case 1:Ring3a.IsDropped = 0:Ring3c.IsDropped = 1:bump3 = 2
'2               Case 2:Ring3b.IsDropped = 0:Ring3a.IsDropped = 1:bump3 = 3
'               Case 3:Ring3c.IsDropped = 0:Ring3b.IsDropped = 1:bump3 = 4
'               Case 4:Ring3c.IsDropped = 0:Me.TimerEnabled = 0
'           End Select
'       End Sub
 

 
'Sounds
'Rubber Sounds
dim speedx
dim speedy
dim finalspeed

Sub Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 14 then 
		PlaySound "bump"
	End if
	If finalspeed >= 4 AND finalspeed <= 14 then
 		RandomSoundRubber()
 	End If
	If finalspeed < 4 AND finalspeed > 1 then
 		RandomSoundRubberLowVolume()
 	End If
End sub

Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "rubber_hit_1"
		Case 2 : PlaySound "rubber_hit_2"
		Case 3 : PlaySound "rubber_hit_3"
	End Select
End Sub

Sub RandomSoundRubberLowVolume()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "rubber_hit_1_low"
		Case 2 : PlaySound "rubber_hit_2_low"
		Case 3 : PlaySound "rubber_hit_3_low"
	End Select
End Sub

Sub LeftFlipper_Collide(parm)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 4 then 
		RandomSoundFlipper()
	Else 
 		RandomSoundFlipperLowVolume()
 	End If
End Sub

Sub LeftFlipper1_Collide(parm)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 4 then 
		RandomSoundFlipper()
	Else 
 		RandomSoundFlipperLowVolume()
 	End If
End Sub

Sub RightFlipper_Collide(parm)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 4 then 
		RandomSoundFlipper()
	Else 
 		RandomSoundFlipperLowVolume()
 	End If
End Sub

Sub RandomSoundFlipper()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "flip_hit_1"
		Case 2 : PlaySound "flip_hit_2"
		Case 3 : PlaySound "flip_hit_3"
	End Select
End Sub

Sub RandomSoundFlipperLowVolume()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "flip_hit_1_low"
		Case 2 : PlaySound "flip_hit_2_low"
		Case 3 : PlaySound "flip_hit_3_low"
	End Select
End Sub

Sub UpdateDoor
	Primitive7.rotx = (gate5.currentAngle * 35/90)
	'Primitive8.rotx = (gate5.currentAngle * 60/90)+90
    LLogo.ObjRotZ = LeftFlipper.CurrentAngle -90
    Rlogo.ObjRotZ = RightFlipper.CurrentAngle -90
    Rlogo1.ObjRotZ = RightFlipper2.CurrentAngle -90

End Sub

'dim speedx
'dim speedy
'dim finalspeed
'Sub Rubbers_Hit(IDX)
'	finalspeed=SQR(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
'	if finalspeed > 11 then PlaySound "rubber" else PlaySound "rubberFlipper":end if
'End Sub
Sub Gates_Hit(IDX):PlaySound"Gate4":End Sub  
'Sub LeftFlipper_Collide(parm)
'	PlaySound "fx_rubber_flipper2"
'End Sub
'Sub RightFlipper_Collide(parm)
'  PlaySound "fx_rubber_flipper"
'End Sub
'Sub RightFlipper2_Collide(parm)
'  PlaySound "fx_rubber_flipper2"
'End Sub
Sub RailTrigger_Hit():  PlaySound "Rail": End Sub
Sub LRailDropTrigger_Hit():  PlaySound "Drop_Right" :  End Sub
Sub RRailDropTrigger_Hit():  PlaySound "Drop_Right" :  End Sub
'Sub MissileDropTrigger_Hit(): PlaySound "Drop_Left" :  End Sub
Sub ShortDropTrigger2_Hit(): PlaySound "Drop_Left" :  End Sub
Sub Bullseyes_Hit(IDX):PlaySound"target":End Sub  
Sub Metals_Hit(IDX):PlaySound"metalhit":End Sub 

 
'****************************************
'  JP's Fading Lamps 3.4 VP9 Fading only
'      Based on PD's Fading Lights
' SetLamp 0 is Off
' SetLamp 1 is On
' LampState(x) current state
'****************************************
 
Dim LampState(200), FadingLevel(200), FadingState(200)
Dim FlashState(200), FlashLevel(200)
Dim FlashSpeedUp, FlashSpeedDown
Dim x
AllLampsOff()
LampTimer.Interval = 30
LampTimer.Enabled = 1

Sub LampTimer_Timer()
    Dim chgLamp, num, chg, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp) Then
        For ii = 0 To UBound(chgLamp)
            LampState(chgLamp(ii, 0) ) = chgLamp(ii, 1) + 4
            FadingLevel(chgLamp(ii, 0) ) = chgLamp(ii, 1) + 4
			FlashState(chgLamp(ii, 0) ) = chgLamp(ii, 1)
        Next
    End If
    UpdateLamps
	'UpdateLeds

	UpdateDoor
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
	nFadeL 14, l14
	nFadeL 15, l15
	nFadeL 16, l16
	nFadeL 17, l17
	nFadeL 18, l18
	nFadeL 19, l19
	nFadeL 20, l20
	nFadeLm 21, l21x
	nFadeL 21, l21
	nFadeLm 22, l22x
	nFadeL 22, l22
	nFadeLm 23, l23x
	nFadeL 23, l23
	nFadeLm 24, l24x
	nFadeL 24, l24
	nFadeL 25, l25
	nFadeL 26, l26
	nFadeL 27, l27
	nFadeL 28, l28
	nFadeL 29, l29
	nFadeL 30, l30
	nFadeL 31, l31
	nFadeL 32, l32
	nFadeL 33, l33
	nFadeL 34, l34
	nFadeL 35, l35
	nFadeL 36, l36
	nFadeL 37, l37
	nFadeL 38, l38
	nFadeL 39, l39
	nFadeL 40, l40
	nFadeL 41, l41
	nFadeL 42, l42
	nFadeL 43, l43
	nFadeL 44, l44
	nFadeL 45, l45
	nFadeL 46, l46
	nFadeL 47, l47
	nFadeLm 48, l48
	nFadeL  48, l48_a
	'imgFadeL 49, l49, "kk_full", "kk_b", "kk_a", "kk_off"  'NOT USED
	'imgFadeL 50, l50, "kk_full", "kk_b", "kk_a", "kk_off"  'NOT USED
	'imgFadeL 51, l51, "kk_full", "kk_b", "kk_a", "kk_off"  'NOT USED
	'imgFadeL 52, l52, "kk_full", "kk_b", "kk_a", "kk_off"  'NOT USED
	'imgFadeL 53, l53, "kk_full", "kk_b", "kk_a", "kk_off"  'NOT USED
	'imgFadeL 54, l54, "kk_full", "kk_b", "kk_a", "kk_off"  'NOT USED
	'imgFadeL 55, l55, "kk_full", "kk_b", "kk_a", "kk_off"  'NOT USED
	nFadeLm 56, l56
	nFadeL  56, l56_a
'''	Reflection 57, l57r  						
'''	Reflection 58, l58r  						
'''	Reflection 59, l59r 						
'''	Reflection 60, l60r  						
'''	Reflection 61, l61r  						
	nFadeL 62, l62
	nFadeL 63, l63
	nFadeL 64, l64

	'Solenoid flasher inserts
	nFadeL  101, l101
	nFadeL  102, l102
	nFadeL  103, l103
	'imgFadeL  104, l104, "kk_full", "kk_b", "kk_a", "kk_off"
	nFadeL 112, l112 'BG Girl Missile
	nFadeL 113, l113 'BG Kong Explode
	nFadeL 115, l115 'BG Kong Explode


'	'Triggered by Solenoid	
'	'FadeL 101, templight101, templight101 'King Shooter (Playfield lighting behind Shooter lane flasher)
'	FlashAR 101, f101, "wf_on", "wf_a", "wf_b", ARRefresh
'	
''	FadeL 102, l102, l102a 'King Shooter  
''	FadeL 103, l103, l103a 'Ape Ramp BG Skull
''	FadeL 104, templight104, templight104  'BG NDSRS Shooter (Flasher by Shooter lane)
'	FlashAR 104, f104, "yfmed_on", "yfmed_a", "yfmed_b", ARRefresh
''	FadeL 105, l105, l105a 'BG Tower 25k
''	FadeL 106, l106, l106a 'BG Tower 50k
''	FadeL 107, l107, l107a 'BG Tower 100k
''	FadeL 108, l108, l108a 'BG Tower Extra Ball
''	FadeL 109, l109, l109a 'BG Tower Million
'	FlashAR 112, f112, "wf_on", "wf_a", "wf_b", ARRefresh 'Tower Flasher  (Flasher by Tower targets)
'	'Strobe 113,l113_a      'BG Kong Explode
'	FlashARm 113, f113_a, "wf_on", "wf_a", "wf_b", ARRefresh
'	FlashARm 113, f113_b, "wf_on", "wf_a", "wf_b", ARRefresh
''	FadeL 114, templight114, templight114 'Tower Flasher  (Flasher by Tower targets)
'	FlashAR 114, f114, "yfmed_on", "yfmed_a", "yfmed_b", ARRefresh 'Tower Flasher  (Flasher by Tower targets)
''	FadeL 115, templight115, templight115 'Skull Island   (Skull Island - Top Right Corner)
'	FlashARm 115, f115, "wf_on", "wf_a", "wf_b", ARRefresh 'Skull Island   (Skull Island - Top Right Corner)
'	'FadeL 115, templight116, templight116 
'	FlashAR 115, f116, "yfmed_on", "yfmed_a", "yfmed_b", ARRefresh'Missile Kicker (Top Left Flasher)
End Sub

Sub Reflection(nr, ramp) 'used for reflections when there is no off ramp
	Select Case LampState(nr)
		Case 2:ramp.alpha = 0:ramp.TriggerSingleUpdate:LampState(nr) = 0                  'Off
		Case 3:ramp.alpha = 0.3:ramp.TriggerSingleUpdate:LampState(nr) = 2 'fading...
		Case 4:ramp.alpha = 0.6:ramp.TriggerSingleUpdate:LampState(nr) = 3 'fading...
		Case 5:ramp.alpha = 1:ramp.TriggerSingleUpdate:LampState(nr) = 1   'ON
	End Select
End Sub

Sub Strobe(nr,ramp)
	Select Case LampState(nr)
		Case 4:ramp.alpha = 0
		Case 5:ramp.alpha = 1
	End Select
End Sub


Sub FlashARm(nr, ramp, imga, imgb, imgc, r) 'used for reflections when the off is transparent - no ramp
	Select Case LampState(nr)
		Case 2:ramp.alpha = 0
			r.State = ABS(r.state -1)
		Case 3:ramp.image = imgc
			r.State = ABS(r.state -1)
		Case 4:ramp.image = imgb
			r.State = ABS(r.state -1)
		Case 5:ramp.alpha = 1
			ramp.image = imgb
			r.State = ABS(r.state -1)
		Case 6:ramp.image = imga
			r.State = ABS(r.state -1)
	End Select
End Sub


Sub FlashAR(nr, ramp, imga, imgb, imgc, r) 'used for reflections when the off is transparent - no ramp
	Select Case LampState(nr)
		Case 2:ramp.alpha = 0
			r.State = ABS(r.state -1)
			LampState(nr) = 0 'Off
		Case 3:ramp.image = imgc
			r.State = ABS(r.state -1)
			LampState(nr) = 2 'fading...
		Case 4:ramp.image = imgb
			r.State = ABS(r.state -1)
			LampState(nr) = 3 'fading...
		Case 5:ramp.alpha = 1
			ramp.image = imgb
			r.State = ABS(r.state -1)
			LampState(nr) = 6 ' 1/2 ON
		Case 6:ramp.image = imga
			r.State = ABS(r.state -1)
			LampState(nr) = 1 'ON
	End Select
End Sub

 Sub AllLampsOff():For x = 1 to 200:LampState(x) = 4:Next:UpdateLamps:UpdateLamps:Updatelamps:End Sub
 
 Sub SetLamp(nr, value)
	'DEBUG.PRINT timer & " Set:" & nr & ", " & value
	LampState(nr) = abs(value) + 4
	FadingLevel(nr) = abs(value) + 4
	FlashState(nr) = abs(value)
End Sub

Sub FadeAR(nr, a, b, c, d, r, wt, wb)
	Select Case LampState(nr)
		Case 2:c.WidthBottom = 0:c.WidthTop = 0:d.WidthBottom = wb:d.WidthTop = wt:LampState(nr) = 0 'Off
		Case 3:b.WidthBottom = 0:b.WidthTop = 0:c.WidthBottom = wb:c.WidthTop = wt:LampState(nr) = 2 'fading...
		Case 4:a.WidthBottom = 0:a.WidthTop = 0:b.WidthBottom = wb:b.WidthTop = wt:LampState(nr) = 3 'fading...
		Case 5:d.WidthBottom = 0:d.WidthTop = 0:b.WidthBottom = wb:b.WidthTop = wt:LampState(nr) = 6 ' 1/2 ON
		Case 6:b.WidthBottom = 0:b.WidthTop = 0:a.WidthBottom = wb:a.WidthTop = wt:LampState(nr) = 1 'ON
	End Select

	r.State = ABS(r.state -1)                                                                        'refresh
End Sub

 Sub FadeAR3(nr, a, b, c, r, wt, wb)
     Select Case LampState(nr)
         Case 2:c.WidthBottom = 0:c.WidthTop = 0:LampState(nr) = 0                 'Off
         Case 3:b.WidthBottom = 0:b.WidthTop = 0:c.WidthBottom = wb:c.WidthTop = wt:LampState(nr) = 2 'fading...
         Case 4:a.WidthBottom = 0:a.WidthTop = 0:b.WidthBottom = wb:b.WidthTop = wt:LampState(nr) = 3 'fading...
         Case 5:b.WidthBottom = 0:b.WidthTop = 0:c.WidthBottom = wb:c.WidthTop = wt:LampState(nr) = 6 'turning ON
         Case 6:c.WidthBottom = 0:c.WidthTop = 0:a.WidthBottom = wb:a.WidthTop = wt:LampState(nr) = 1 'ON
     End Select

	r.State = ABS(r.state -1)
 End Sub
 
 Sub FadeW(nr, a, b, c)
     Select Case LampState(nr)
         Case 2:c.IsDropped = 1:LampState(nr) = 0                 'Off
         Case 3:b.IsDropped = 1:c.IsDropped = 0:LampState(nr) = 2 'fading...
         Case 4:a.IsDropped = 1:b.IsDropped = 0:LampState(nr) = 3 'fading...
         Case 5:b.IsDropped = 1:c.IsDropped = 0:LampState(nr) = 6 'turning ON
         Case 6:c.IsDropped = 1:a.IsDropped = 0:LampState(nr) = 1 'ON
     End Select
 End Sub

 Sub FadeLGIB(nr, a)
     Select Case LampState(nr)
         Case 2:LampState(nr) = 0                 'Off
         Case 3:For each xx in a:xx.State = 0:Next:LampState(nr) = 2 'fading...
         Case 4:LampState(nr) = 3 'fading...
         Case 5:For each xx in a:xx.State = 1:Next
				LampState(nr) = 6 'turning ON
         Case 6:LampState(nr) = 1 'ON
     End Select
 End Sub

 Sub FadeLGIBm (nr, a)
     Select Case LampState(nr)
         Case 0:For each xx in a:xx.State = 0:Next 'fading...
         Case 1:For each xx in a:xx.State = 1:Next
     End Select
 End Sub
 
  Sub FadeGICW(nr, a, b, c, d, e, f)
     Select Case LampState(nr)
         Case 2:c.IsDropped = 1:For each xx in f:xx.IsDropped=1:Next:LampState(nr) = 0                 'Off
         Case 3:b.IsDropped = 1:c.IsDropped = 0:LampState(nr) = 2 'fading...
         Case 4:a.IsDropped = 1:b.IsDropped = 0:LampState(nr) = 3 'fading...
         Case 5:b.IsDropped = 1:c.IsDropped = 0
 				If GIState=1 then
 					For each xx in d:xx.IsDropped=0:Next
 				else
 					For each xx in e:xx.IsDropped=0:Next
 				end if
 				LampState(nr) = 6 'turning ON
         Case 6:c.IsDropped = 1:a.IsDropped = 0:LampState(nr) = 1 'ON
     End Select
 End Sub
 
  Sub FadeCW(nr, a, b, c, d)
     Select Case LampState(nr)
         Case 2:c.IsDropped = 1:For each xx in d:xx.IsDropped=1:Next:LampState(nr) = 0                 'Off
         Case 3:b.IsDropped = 1:c.IsDropped = 0:LampState(nr) = 2 'fading...
         Case 4:a.IsDropped = 1:b.IsDropped = 0:LampState(nr) = 3 'fading...
         Case 5:b.IsDropped = 1:c.IsDropped = 0:For each xx in d:xx.IsDropped=0:Next:LampState(nr) = 6 'turning ON
         Case 6:c.IsDropped = 1:a.IsDropped = 0:LampState(nr) = 1 'ON
     End Select
 End Sub
 
  Sub NFadeCP(a, b, c, d, e, f, g, h, l, m)
 	 for each xx in CPLL: xx.IsDropped=1:next
     If (LampState(a)=1 and LampState(b)=1 and LampState(c)=1 and LampState(d)=1) then h.IsDropped = 0
     If (LampState(a)=1 and LampState(b)=1 and LampState(c)=1 and LampState(d)<>1) then g.IsDropped = 0
 	 If (LampState(a)=1 and LampState(b)=1 and LampState(c)<>1 and LampState(d)<>1) then f.IsDropped = 0
      If (LampState(a)=1 and LampState(b)<>1 and LampState(c)<>1 and LampState(d)<>1) then e.IsDropped = 0
 	 If (LampState(a)<>1 and LampState(b)<>1 and LampState(c)>=1 and LampState(d)<>1) then l.IsDropped = 0
 	 If (LampState(a)<>1 and LampState(b)<>1 and LampState(c)<>1 and LampState(d)=1) then m.IsDropped = 0
 End Sub
 
  Sub FadeGW(nr, a, b, c, d, e, f) 'Fade a group of walls for double flashers
     Select Case LampState(nr)
         Case 2:c.IsDropped = 1:f.IsDropped = 1:LampState(nr) = 0                 'Off
         Case 3:b.IsDropped = 1:c.IsDropped = 0:e.IsDropped = 1:f.IsDropped = 0:LampState(nr) = 2 'fading...
         Case 4:a.IsDropped = 1:b.IsDropped = 0:d.IsDropped = 1:e.IsDropped = 0:LampState(nr) = 3 'fading...
         Case 5:b.IsDropped = 1:c.IsDropped = 0:e.IsDropped = 1:f.IsDropped = 0:LampState(nr) = 6 'turning ON
         Case 6:c.IsDropped = 1:a.IsDropped = 0:f.IsDropped = 1:d.IsDropped = 0:LampState(nr) = 1 'ON
     End Select
 End Sub
 
   Sub FadeCGW(nr, a, b, c, d, e, f, g, h, i) 'Fade a group of walls for double flashers
     Select Case LampState(nr)
         Case 2:c.IsDropped = 1:f.IsDropped = 1:For each xx in i:xx.IsDropped=1:Next:LampState(nr) = 0                 'Off
         Case 3:b.IsDropped = 1:c.IsDropped = 0:e.IsDropped = 1:f.IsDropped = 0:LampState(nr) = 2 'fading...
         Case 4:a.IsDropped = 1:b.IsDropped = 0:d.IsDropped = 1:e.IsDropped = 0:LampState(nr) = 3 'fading...
         Case 5:b.IsDropped = 1:c.IsDropped = 0:e.IsDropped = 1:f.IsDropped = 0
 				If GIState=1 then
 					For each xx in g:xx.IsDropped=0:Next
 				else
 					For each xx in h:xx.IsDropped=0:Next
 				end if
 				LampState(nr) = 6 'turning ON
         Case 6:c.IsDropped = 1:a.IsDropped = 0:f.IsDropped = 1:d.IsDropped = 0:LampState(nr) = 1 'ON
     End Select
 End Sub
 
   Sub FadeGWm(nr, a, b, c, d, e, f) 'Fade a group of walls for double flashers
     Select Case LampState(nr)
         Case 2:c.IsDropped = 1:f.IsDropped = 1:LampState(nr) = 0                 'Off
         Case 3:b.IsDropped = 1:c.IsDropped = 0:e.IsDropped = 1:f.IsDropped = 0
         Case 4:a.IsDropped = 1:b.IsDropped = 0:d.IsDropped = 1:e.IsDropped = 0
         Case 5:b.IsDropped = 1:c.IsDropped = 0:e.IsDropped = 1:f.IsDropped = 0
         Case 6:c.IsDropped = 1:a.IsDropped = 0:f.IsDropped = 1:d.IsDropped = 0
     End Select
 End Sub
 
 Sub FadeWm(nr, a, b, c)
     Select Case LampState(nr)
         Case 2:c.IsDropped = 1
         Case 3:b.IsDropped = 1:c.IsDropped = 0
         Case 4:a.IsDropped = 1:b.IsDropped = 0
         Case 5:b.IsDropped = 1:c.IsDropped = 0
         Case 6:c.IsDropped = 1:a.IsDropped = 0
     End Select
 End Sub
 
 Sub NFadeW(nr, a)
     Select Case LampState(nr)
         Case 4:a.IsDropped = 1:LampState(nr) = 0
         Case 5:a.IsDropped = 0:LampState(nr) = 1
     End Select
 End Sub

 Sub NFadeWCo(nr, a)
     Select Case LampState(nr)
         Case 4:For each xx in a:xx.IsDropped = 1:Next:LampState(nr) = 0
         Case 5:For each xx in a:xx.IsDropped = 0:Next:LampState(nr) = 1
     End Select
 End Sub

 Sub FadeBump1(nr)
     Select Case LampState(nr)
         Case 4:B1Off:LampState(nr) = 0
         Case 5:B1On:LampState(nr) = 1
     End Select
 End Sub

 Sub FadeBump2(nr)
     Select Case LampState(nr)
         Case 4:B2Off:LampState(nr) = 0
         Case 5:B2On:LampState(nr) = 1
     End Select
 End Sub

 Sub FadeBump3(nr)
     Select Case LampState(nr)
         Case 4:B3Off:LampState(nr) = 0
         Case 5:B3On:LampState(nr) = 1
     End Select
 End Sub
 
 Sub NFadeWm(nr, a)
     Select Case LampState(nr)
         Case 4:a.IsDropped = 1
         Case 5:a.IsDropped = 0
     End Select
 End Sub
 
 Sub NFadeWi(nr, a)
     Select Case LampState(nr)
         Case 5:a.IsDropped = 1:LampState(nr) = 0
         Case 4:a.IsDropped = 0:LampState(nr) = 1
     End Select
 End Sub
 
 Sub FadeL(nr, a, b)
     Select Case LampState(nr)
         Case 2:b.state = 0:LampState(nr) = 0
         Case 3:b.state = 1:LampState(nr) = 2
         Case 4:a.state = 0:LampState(nr) = 3
         Case 5:b.state = 1:LampState(nr) = 6
         Case 6:a.state = 1:LampState(nr) = 1
     End Select
 End Sub
 
 Sub FadeLm(nr, a, b)
     Select Case LampState(nr)
         Case 2:b.state = 0
         Case 3:b.state = 1
         Case 4:a.state = 0
         Case 5:b.state = 1
         Case 6:a.state = 1
     End Select
 End Sub

Sub imgFadeL(nr, Light, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:Light.Offimage = d:FadingLevel(nr) = 0 'Off
        Case 3:Light.Offimage = c:FadingLevel(nr) = 2 'fading...
        Case 4:Light.Offimage = b:FadingLevel(nr) = 3
        Case 5:Light.Offimage = a:FadingLevel(nr) = 1
    End Select
End Sub

Sub imgFadeLm(nr, Light, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:Light.Offimage = d: 'Off
        Case 3:Light.Offimage = c: 'fading...
        Case 4:Light.Offimage = b:
        Case 5:Light.Offimage = a:
    End Select
End Sub
  Sub FadeGL(nr, a, b, c, d)
     Select Case LampState(nr)
         Case 2:b.state = 0:d.state = 0:LampState(nr) = 0
         Case 3:b.state = 1:d.state = 1:LampState(nr) = 2
         Case 4:a.state = 0:c.state = 0:LampState(nr) = 3
         Case 5:b.state = 1:d.state = 1:LampState(nr) = 6
         Case 6:a.state = 1:c.state = 1:LampState(nr) = 1
     End Select
 End Sub
 

 
 Sub NFadeL(nr, a)
     Select Case LampState(nr)
         Case 4:a.state = 0:LampState(nr) = 0
         Case 5:a.State = 1:LampState(nr) = 1
     End Select
 End Sub
 
 Sub NFadeLm(nr, a)
     Select Case LampState(nr)
         Case 4:a.state = 0
         Case 5:a.State = 1
     End Select
 End Sub
 
 Sub FadeR(nr, a)
     Select Case LampState(nr)
         Case 2:a.SetValue 3:LampState(nr) = 0
         Case 3:a.SetValue 2:LampState(nr) = 2
         Case 4:a.SetValue 1:LampState(nr) = 3
         Case 5:a.SetValue 1:LampState(nr) = 6
         Case 6:a.SetValue 0:LampState(nr) = 1
     End Select
 End Sub
 
  Sub FadeRG(nr, a, b)
     Select Case LampState(nr)
         Case 2:a.SetValue 3:b.SetValue 3:LampState(nr) = 0
         Case 3:a.SetValue 2:b.SetValue 2:LampState(nr) = 2
         Case 4:a.SetValue 1:b.SetValue 1:LampState(nr) = 3
         Case 5:a.SetValue 1:b.SetValue 1:LampState(nr) = 6
         Case 6:a.SetValue 0:b.SetValue 0:LampState(nr) = 1
     End Select
 End Sub
 
 Sub FadeRm(nr, a)
     Select Case LampState(nr)
         Case 2:a.SetValue 3
         Case 3:a.SetValue 2
         Case 4:a.SetValue 1
         Case 5:a.SetValue 1
         Case 6:a.SetValue 0
     End Select
 End Sub
 
 Sub NFadeT(nr, a, b)
     Select Case LampState(nr)
         Case 4:a.Text = "":LampState(nr) = 0
         Case 5:a.Text = b:LampState(nr) = 1
     End Select
 End Sub
 
 Sub NFadeTm(nr, a, b)
     Select Case LampState(nr)
         Case 4:a.Text = ""
         Case 5:a.Text = b
     End Select
 End Sub
 
 Sub NFadeWi(nr, a)
     Select Case LampState(nr)
         Case 4:a.IsDropped = 0:LampState(nr) = 0
         Case 5:a.IsDropped = 1:LampState(nr) = 1
     End Select
 End Sub
 
 Sub NFadeWim(nr, a)
     Select Case LampState(nr)
         Case 4:a.IsDropped = 0
         Case 5:a.IsDropped = 1
     End Select
 End Sub
 
 Sub FadeLCo(nr, a, b) 'fading collection of lights
     Dim obj
     Select Case LampState(nr)
         Case 2:vpmSolToggleObj b, Nothing, 0, 0:LampState(nr) = 0
         Case 3:vpmSolToggleObj b, Nothing, 0, 1:LampState(nr) = 2
         Case 4:vpmSolToggleObj a, Nothing, 0, 0:LampState(nr) = 3
         Case 5:vpmSolToggleObj b, Nothing, 0, 1:LampState(nr) = 6
         Case 6:vpmSolToggleObj a, Nothing, 0, 1:LampState(nr) = 1
     End Select
 End Sub
 
  Sub NFadeLCo(nr, a) 'fading collection of lights
     Dim obj
     Select Case LampState(nr)
         Case 4:vpmSolToggleObj a, Nothing, 0, 0:LampState(nr) = 0
         Case 5:vpmSolToggleObj a, Nothing, 1, 1:LampState(nr) = 1
     End Select
 End Sub
 
 Sub FlashL(nr, a, b) ' simple light flash, not controlled by the rom
     Select Case LampState(nr)
         Case 2:b.state = 0:LampState(nr) = 0
         Case 3:b.state = 1:LampState(nr) = 2
         Case 4:a.state = 0:LampState(nr) = 3
         Case 5:a.state = 1:LampState(nr) = 4
     End Select
 End Sub
 
 Sub MFadeL(nr, a, b, c) 'Light acting as a flash. C is the light number to be restored
     Select Case LampState(nr)
         Case 2:b.state = 0:LampState(nr) = 0
             If LampState(c) = 1 Then SetLamp c, 1
         Case 3:b.state = 1:LampState(nr) = 2
         Case 4:a.state = 0:LampState(nr) = 3
         Case 5:a.state = 1:LampState(nr) = 1
     End Select
 End Sub
 
 Sub NFadeB(nr, a, b, c, d) 'New Bally Bumpers: a and b are the off state, c and d and on state, no fading
     Select Case LampState(nr)
         Case 4:a.IsDropped = 0:b.IsDropped = 0:c.IsDropped = 1:d.IsDropped = 1:LampState(nr) = 0
         Case 5:a.IsDropped = 1:b.IsDropped = 1:c.IsDropped = 0:d.IsDropped = 0:LampState(nr) = 1
     End Select
 End Sub
 
 Sub NFadeBm(nr, a, b, c, d)
     Select Case LampState(nr)
         Case 4:a.IsDropped = 0:b.IsDropped = 0:c.IsDropped = 1:d.IsDropped = 1
         Case 5:a.IsDropped = 1:b.IsDropped = 1:c.IsDropped = 0:d.IsDropped = 0
     End Select
 End Sub
 
 ' only for this table
 
 Sub ThumperW(nr, a, b) 'nr = light number, a = dropwall array 16 walls, b = bumper animation state
     Dim ii
     Select Case LampState(nr)
         Case 2:For Each ii in a:ii.IsDropped = 1:Next:a(b).IsDropped = 0:LampState(nr) = 0      'Off
         Case 3:For Each ii in a:ii.IsDropped = 1:Next:a(4 + b).IsDropped = 0:LampState(nr) = 2  'fading to Off
         Case 4:For Each ii in a:ii.IsDropped = 1:Next:a(8 + b).IsDropped = 0:LampState(nr) = 3  'fading to Off
         Case 5:For Each ii in a:ii.IsDropped = 1:Next:a(12 + b).IsDropped = 0:LampState(nr) = 1 'On
     End Select
 End Sub
 
 Sub ThumperWm(nr, a, b) 'used to update the state with the bumper animation
     Dim ii
     Select Case LampState(nr)
         Case 0:For Each ii in a:ii.IsDropped = 1:Next:a(b).IsDropped = 0
         Case 1:For Each ii in a:ii.IsDropped = 1:Next:a(12 + b).IsDropped = 0
         Case 2:For Each ii in a:ii.IsDropped = 1:Next:a(b).IsDropped = 0
         Case 3:For Each ii in a:ii.IsDropped = 1:Next:a(4 + b).IsDropped = 0
         Case 4:For Each ii in a:ii.IsDropped = 1:Next:a(8 + b).IsDropped = 0
         Case 5:For Each ii in a:ii.IsDropped = 1:Next:a(12 + b).IsDropped = 0
 
     End Select
 End Sub
 
 
 
   '*************************************
  '          Nudge System
  ' JP's based on Noah's nudgetest table
  '*************************************
  
  Dim LeftNudgeEffect, RightNudgeEffect, NudgeEffect
  
  Sub LeftNudge(angle, strength, delay)
      vpmNudge.DoNudge angle, (strength * (delay-LeftNudgeEffect) / delay) + RightNudgeEffect / delay
      LeftNudgeEffect = delay
      RightNudgeEffect = 0
      RightNudgeTimer.Enabled = 0
      LeftNudgeTimer.Interval = delay
      LeftNudgeTimer.Enabled = 1
  End Sub
  
  Sub RightNudge(angle, strength, delay)
      vpmNudge.DoNudge angle, (strength * (delay-RightNudgeEffect) / delay) + LeftNudgeEffect / delay
      RightNudgeEffect = delay
      LeftNudgeEffect = 0
      LeftNudgeTimer.Enabled = 0
      RightNudgeTimer.Interval = delay
      RightNudgeTimer.Enabled = 1
  End Sub
  
  Sub CenterNudge(angle, strength, delay)
      vpmNudge.DoNudge angle, strength * (delay-NudgeEffect) / delay
      NudgeEffect = delay
      NudgeTimer.Interval = delay
      NudgeTimer.Enabled = 1
  End Sub
  
  Sub LeftNudgeTimer_Timer()
      LeftNudgeEffect = LeftNudgeEffect-1
      If LeftNudgeEffect = 0 then LeftNudgeTimer.Enabled = 0
  End Sub
  
  Sub RightNudgeTimer_Timer()
      RightNudgeEffect = RightNudgeEffect-1
      If RightNudgeEffect = 0 then RightNudgeTimer.Enabled = 0
  End Sub
  
  Sub NudgeTimer_Timer()
      NudgeEffect = NudgeEffect-1
      If NudgeEffect = 0 then NudgeTimer.Enabled = 0
  End Sub
 
'   '****************************************
'  '  rascal's Ball Rolling Script
'  '****************************************
'  Dim VeloY(3),VeloX(3),tapglass,rolling(3),b
' Sub RollTimer_Timer()       
' B=B+1
'  If B>3 Then B=1
'  If BallStatus(b)=0 Then Exit Sub
'
' VeloY(b) = Cint(CurrentBall(b).VelY) + 50       
' VeloX(b) = Cint(CurrentBall(b).VelX) + 50       
' If ((VeloY(b) < 40 or VeloY(b) > 60) or (VeloX(b) < 40 or VeloX(b) > 60)) Then
' 	If rolling(b) = True then 
' 		Exit Sub
' 	Else
' 		rolling(b) = True
' 		PlaySound "roll"&b
' 	End If     
' Else
' 	If rolling(b)=True Then
' 	StopSound "roll"&b
' 	PlaySound "rollstop"  
' 	rolling(b) = False
' 	End If
'  
'  End If
'    
'' If Cint(CurrentBall(b).Z) > 99 then PlaySound "tap": CurrentBall(b).Z = 99       
' 'If Cint(CurrentBall(b).Z) > 62 then TapTimer.Enabled = True: CurrentBall(b).Z = 62  
' 
' End Sub
' 
' Sub TapTimer_Timer()       
' PlaySound "ballbounce"       
' TapTimer.Enabled = False  
' End Sub
' 
' '======================================================================================
' ' Many thanks go to Randy Davis and to all the multitudes of people who have 
' ' contributed to VP over the years, keeping it alive!!!  Enjoy, Steely & PK
' '======================================================================================
' 
' Dim tnopb, nosf
' '
' tnopb = 5 	' <<<<< SET to the "Total Number Of Possible Balls" in play at any one time
' nosf = 9	' <<<<< SET to the "Number Of Sound Files" used / B2B collision volume levels
' 
' Dim currentball(5), ballStatus(5)
' Dim iball, cnt, coff, errMessage
' 
' XYdata.interval = 1 			' Timer interval starts at 1 for the highest ball data sample rate
' coff = False				' Collision off set to false
' 
' For cnt = 0 to ubound(ballStatus)	' Initialize/clear all ball stats, 1 = active, 0 = non-existant
' 	ballStatus(cnt) = 0
' Next
' 
' '======================================================
' ' <<<<<<<<<<<<<< Ball Identification >>>>>>>>>>>>>>
' '======================================================
' ' Call this sub from every kicker(or plunger) that creates a ball.
' Sub NewBallID 						' Assign new ball object and give it ID for tracking
' 	For cnt = 1 to ubound(ballStatus)		' Loop through all possible ball IDs
'    	If ballStatus(cnt) = 0 Then			' If ball ID is available...
'    	Set currentball(cnt) = ActiveBall			' Set ball object with the first available ID
'    	currentball(cnt).uservalue = cnt			' Assign the ball's uservalue to it's new ID
'    	ballStatus(cnt) = 1				' Mark this ball status active
'    	ballStatus(0) = ballStatus(0)+1 		' Increment ballStatus(0), the number of active balls
' 	If coff = False Then				' If collision off, overrides auto-turn on collision detection
' 							' If more than one ball active, start collision detection process
' 	If ballStatus(0) > 1 and XYdata.enabled = False Then XYdata.enabled = True
' 	End If
' 	Exit For					' New ball ID assigned, exit loop
'    	End If
'    	Next 
' '  	Debugger 					' For demo only, display stats
' End Sub
' 
' ' Call this sub from every kicker that destroys a ball, before the ball is destroyed.
' Sub ClearBallID
'   	On Error Resume Next				' Error handling for debugging purposes
'    	iball = ActiveBall.uservalue			' Get the ball ID to be cleared
'    	currentball(iball).UserValue = 0 			' Clear the ball ID
'    	If Err Then Msgbox Err.description & vbCrLf & iball
'     	ballStatus(iBall) = 0 				' Clear the ball status
'    	ballStatus(0) = ballStatus(0)-1 		' Subtract 1 ball from the # of balls in play
'    	On Error Goto 0
' End Sub
' 
' '=====================================================
' ' <<<<<<<<<<<<<<<<< XYdata_Timer >>>>>>>>>>>>>>>>>
' '=====================================================
' ' Ball data collection and B2B Collision detection.
' ReDim baX(tnopb,4), baY(tnopb,4), bVx(tnopb,4), bVy(tnopb,4), TotalVel(tnopb,4)
' Dim cForce, bDistance, xyTime, cFactor, id, id2, id3, B1, B2
' 
' Sub XYdata_Timer()
' 	' xyTime... Timers will not loop or start over 'til it's code is finished executing. To maximize
' 	' performance, at the end of this timer, if the timer's interval is shorter than the individual
' 	' computer can handle this timer's interval will increment by 1 millisecond. 
'     xyTime = Timer+(XYdata.interval*.001)	' xyTime is the system timer plus the current interval time
' 	' Ball Data... When a collision occurs a ball's velocity is often less than it's velocity before the
' 	' collision, if not zero. So the ball data is sampled and saved for four timer cycles. 
'    	If id2 >= 4 Then id2 = 0						' Loop four times and start over
'    	id2 = id2+1								' Increment the ball sampler ID
'    	For id = 1 to ubound(ballStatus)					' Loop once for each possible ball
'   	If ballStatus(id) = 1 Then						' If ball is active...
'    		baX(id,id2) = round(currentball(id).x,2)				' Sample x-coord
'    		baY(id,id2) = round(currentball(id).y,2)				' Sample y-coord
'    		bVx(id,id2) = round(currentball(id).velx,2)				' Sample x-velocity
'    		bVy(id,id2) = round(currentball(id).vely,2)				' Sample y-velocity
'    		TotalVel(id,id2) = (bVx(id,id2)^2+bVy(id,id2)^2) 		' Calculate total velocity
'  		If TotalVel(id,id2) > TotalVel(0,0) Then TotalVel(0,0) = int(TotalVel(id,id2))
'    	End If
'    	Next
' 	' Collision Detection Loop - check all possible ball combinations for a collision.
' 	' bDistance automatically sets the distance between two colliding balls. Zero milimeters between
' 	' balls would be perfect, but because of timing issues with ball velocity, fast-traveling balls
' 	' prevent a low setting from always working, so bDistance becomes more of a sensitivity setting,
' 	' which is automated with calculations using the balls' velocities.
' 	' Ball x/y-coords plus the bDistance determines B2B proximity and triggers a collision. 
' 	id3 = id2 : B2 = 2 : B1 = 1						' Set up the counters for looping
' 	Do
' 	If ballStatus(B1) = 1 and ballStatus(B2) = 1 Then			' If both balls are active...
' 		bDistance = int((TotalVel(B1,id3)+TotalVel(B2,id3))^1.04)
' 		If ((baX(B1,id3)-baX(B2,id3))^2+(baY(B1,id3)-baY(B2,id3))^2)<2800+bDistance Then collide B1,B2 : Exit Sub		
' 		End If
' 		B1 = B1+1							' Increment ball1
' 		If B1 >= ballStatus(0) Then Exit Do				' Exit loop if all ball combinations checked	
' 		If B1 >= B2 then B1 = 1:B2 = B2+1				' If ball1 >= reset ball1 and increment ball2
' 	Loop
' 	
'  	If ballStatus(0) <= 1 Then XYdata.enabled = False 			' Turn off timer if one ball or less
' 
' 	If XYdata.interval >= 40 Then coff = True : XYdata.enabled = False	' Auto-shut off
' 	If Timer > xyTime * 3 Then coff = True : XYdata.enabled = False		' Auto-shut off
'    	If Timer > xyTime Then XYdata.interval = XYdata.interval+1		' Increment interval if needed
' End Sub
' 
' '=========================================================
' ' <<<<<<<<<<< Collide(ball id1, ball id2) >>>>>>>>>>>
' '=========================================================
' 'Calculate the collision force and play sound accordingly.
' Dim cTime, cb1,cb2, avgBallx, cAngle, bAngle1, bAngle2
' 
' Sub Collide(cb1,cb2) 	
' ' The Collision Factor(cFactor) uses the maximum total ball velocity and automates the cForce calculation, maximizing the
' ' use of all sound files/volume levels. So all the available B2B sound levels are automatically used by adjusting to a
' ' player's style and the table's characteristics.
'  	If TotalVel(0,0)/1.8 > cFactor Then cFactor = int(TotalVel(0,0)/1.8)
' ' The following six lines limit repeated collisions if the balls are close together for any period of time
'   	avgBallx = (bvX(cb2,1)+bvX(cb2,2)+bvX(cb2,3)+bvX(cb2,4))/4
'   	If avgBallx < bvX(cb2,id2)+.1 and avgBallx > bvX(cb2,id2)-.1 Then
'  	If ABS(TotalVel(cb1,id2)-TotalVel(cb2,id2)) < .000005 Then Exit Sub
'  	End If
'   	If Timer < cTime Then Exit Sub
'   	cTime = Timer+.1				' Limits collisions to .1 seconds apart
' ' GetAngle(x-value, y-value, the angle name) calculates any x/y-coords or x/y-velocities and returns named angle in radians
'  	GetAngle baX(cb1,id3)-baX(cb2,id3), baY(cb1,id3)-baY(cb2,id3),cAngle	' Collision angle via x/y-coordinates
' 	id3 = id3 - 1 : If id3 = 0 Then id3 = 4		' Step back one xyData sampling for a good velocity reading
'  	GetAngle bVx(cb1,id3), bVy(cb1,id3), bAngle1	' ball 1 travel direction, via velocity
'  	GetAngle bVx(cb2,id3), bVy(cb2,id3), bAngle2	' ball 2 travel direction, via velocity
' ' The main cForce formula, calculating the strength of a collision
' 	cForce = Cint((abs(TotalVel(cb1,id3)*Cos(cAngle-bAngle1))+abs(TotalVel(cb2,id3)*Cos(cAngle-bAngle2))))
'     	If cForce < 4 Then Exit Sub			' Another collision limiter
'    	cForce = Cint((cForce)/(cFactor/nosf))		' Divides up cForce for the proper sound selection.
'   	If cForce > nosf-1 Then cForce = nosf-1		' First sound file 0(zero) minus one from number of sound files
'    	PlaySound("collide" & cForce)			' Combines "collide" with the calculated sound level and play sound
' End Sub
' 
' '=================================================
' ' <<<<<<<< GetAngle(X, Y, Anglename) >>>>>>>>
' '=================================================
' ' A repeated function which takes any set of coordinates or velocities and calculates an angle in radians.
' Dim Xin,Yin,rAngle,Radit,wAngle,Pi
' Pi = Round(4*Atn(1),6)					'3.1415926535897932384626433832795
' 
' Sub GetAngle(Xin, Yin, wAngle)						
'  	If Sgn(Xin) = 0 Then
'  		If Sgn(Yin) = 1 Then rAngle = 3 * Pi/2 Else rAngle = Pi/2
'  		If Sgn(Yin) = 0 Then rAngle = 0
'  	Else
'  		rAngle = atn(-Yin/Xin)			' Calculates angle in radians before quadrant data
'  	End If
'  	If sgn(Xin) = -1 Then Radit = Pi Else Radit = 0
'  	If sgn(Xin) = 1 and sgn(Yin) = 1 Then Radit = 2 * Pi
'  	wAngle = round((Radit + rAngle),4)		' Calculates angle in radians with quadrant data
' 	'"wAngle = round((180/Pi) * (Radit + rAngle),4)" ' Will convert radian measurements to degrees - to be used in future
' End Sub
' 


'      [ raising ramp animation]
 Dim lkpos	'Animation position 
  lkpos = 0:	'Set Starting positions
 
 sub lkong_anim(enabled)
 if enabled then
 	lkongout.Enabled=1 'Start animation

 else
    lkongin.Enabled=1

    
 end if
 End Sub
 





  Sub lkongout_Timer()
 	 Select Case lkPos	
 			Case 1:lskong.x=175
 			Case 2:lskong.x=185
 			Case 3:lskong.x=195
 			Case 4:lskong.x=205
 			Case 5:lskong.x=215
            Case 6:lskong.x=225:lkongout.Enabled=0:lkongin.Enabled=1:

 
End Select

 	If lkpos<6 then lkPos=lkpos+1
  End Sub
 
  Sub lkongin_Timer()
 	 Select Case lkPos	


 			Case 1:lskong.x=175:lkongin.Enabled=0
 			Case 2:lskong.x=185
 			Case 3:lskong.x=195
 			Case 4:lskong.x=205
 			Case 5:lskong.x=215
            Case 6:lskong.x=225
 	End Select
 	If lkpos>0 Then lkPos=lkpos-1
  End Sub
'**************************************************************************************************************
 Dim rkpos	'Animation position 
  rkpos = 0:	'Set Starting positions
 
 sub rkong_anim(enabled)
 if enabled then
 	rkongout.Enabled=1 'Start animation

 else
    rkongin.Enabled=1

    
 end if
 End Sub
 





  Sub rkongout_Timer()
 	 Select Case rkPos	
 			Case 1:rskong.x=680
 			Case 2:rskong.x=670
 			Case 3:rskong.x=660
 			Case 4:rskong.x=650
 			Case 5:rskong.x=640
            Case 6:rskong.x=630:rkongout.Enabled=0:rkongin.Enabled=1:

 
End Select

 	If rkpos<6 then rkPos=rkpos+1
  End Sub
 
  Sub rkongin_Timer()
 	 Select Case rkPos	


 			Case 1:rskong.x=680:rkongin.Enabled=0
 			Case 2:rskong.x=670
 			Case 3:rskong.x=660
 			Case 4:rskong.x=650
 			Case 5:rskong.x=640
            Case 6:rskong.x=630
 	End Select
 	If rkpos>0 Then rkPos=rkpos-1
  End Sub


  
  '***********************
'   Visible Locks - from JPSalas BTTF table
' Adapted to this table
' based on the core.vbs
'***********************
Sub Sound2_Hit: vlockReturnTrigger.enabled = 1:End Sub
Sub vlockReturnTrigger_Hit() :  PlaySound "Drop_left":vlockReturnTrigger.enabled = 0: end sub 


Class cvpmVLock2
    Private mTrig, mKick, mSw(), mSize, mBalls, mGateOpen, mRealForce, mBallSnd, mNoBallSnd
    Public ExitDir, ExitForce, KickForceVar

    Private Sub Class_Initialize
        mBalls = 0:ExitDir = 0:ExitForce = 0:KickForceVar = 0:mGateOpen = False
        vpmTimer.addResetObj Me
    End Sub

    Public Sub InitVLock(aTrig, aKick, aSw)
        Dim ii
        mSize = vpmSetArray(mTrig, aTrig)
        If vpmSetArray(mKick, aKick) <> mSize Then MsgBox "cvpmVLock: Unmatched kick+trig":Exit Sub
        On Error Resume Next
        ReDim mSw(mSize)
        If IsArray(aSw) Then
            For ii = 0 To UBound(aSw):mSw(ii) = aSw(ii):Next
        ElseIf aSw = 0 Or Err Then
            For ii = 0 To mSize:mSw(ii) = mTrig(ii).TimerInterval:Next
        Else
            mSw(0) = aSw
        End If
    End Sub

    Public Sub InitSnd(aBall, aNoBall):mBallSnd = aBall:mNoBallSnd = aNoBall:End Sub
    Public Sub CreateEvents(aName)
        Dim ii
        If Not vpmCheckEvent(aName, Me) Then Exit Sub
        For ii = 0 To mSize
            vpmBuildEvent mTrig(ii), "Hit", aName & ".TrigHit ActiveBall," & ii + 1

            vpmBuildEvent mTrig(ii), "Unhit", aName & ".TrigUnhit ActiveBall," & ii + 1

            vpmBuildEvent mKick(ii), "Hit", aName & ".KickHit " & ii + 1
        Next
    End Sub

    Public Sub SolExit(aEnabled)
        Dim ii
        mGateOpen = aEnabled
        If Not aEnabled Then Exit Sub
        If mBalls> 0 Then PlaySound mBallSnd:Else PlaySound mNoBallSnd:Exit Sub
        For ii = 0 To mBalls-1
            mKick(ii).Enabled = False:If mSw(ii) Then Controller.Switch(mSw(ii) ) = False
        Next
        '		If ExitForce > 0 Then ' Up
        '			mRealForce = ExitForce + (Rnd - 0.5)*KickForceVar : mKick(mBalls-1).Kick ExitDir, mRealForce
        '		Else ' Down
        mRealForce = ExitForce + (Rnd - 0.5) * KickForceVar:mKick(0).Kick ExitDir, mRealForce
    '		End If
    End Sub

    Public Sub Reset
        Dim ii:If mBalls = 0 Then Exit Sub
        For ii = 0 To mBalls-1
            If mSw(ii) Then Controller.Switch(mSw(ii) ) = True
        Next
    End Sub

    Public Property Get Balls:Balls = mBalls:End Property

    Public Property Let Balls(aBalls)
        Dim ii:mBalls = aBalls
        For ii = 0 To mSize
            If ii >= aBalls Then
                mKick(ii).DestroyBall:If mSw(ii) Then Controller.Switch(mSw(ii) ) = False
                Else
                    vpmCreateBall mKick(ii):If mSw(ii) Then Controller.Switch(mSw(ii) ) = True
            End If
        Next
    End Property

    Public Sub TrigHit(aBall, aNo)
        aNo = aNo - 1:If mSw(aNo) Then Controller.Switch(mSw(aNo) ) = True
        If aBall.VelY <-1 Then Exit Sub ' Allow small upwards speed
        If aNo = mSize Then mBalls = mBalls + 1
        If mBalls> aNo Then mKick(aNo).Enabled = Not mGateOpen
    End Sub

    Public Sub TrigUnhit(aBall, aNo)
        aNo = aNo - 1:If mSw(aNo) Then Controller.Switch(mSw(aNo) ) = False
        If aBall.VelY> -1 Then
            If aNo = 0 Then mBalls = mBalls - 1
            If aNo <mSize Then mKick(aNo + 1).Kick 0, 0
            Else
                If aNo = mSize Then mBalls = mBalls - 1
                If aNo> 0 Then mKick(aNo-1).Kick ExitDir, mRealForce
        End If
    End Sub

    Public Sub KickHit(aNo):mKick(aNo-1).Enabled = False:End Sub
End Class

' Flasher objects
' Uses own faster timer
FlashInit()
FlasherTimer.Interval = 10 'flash fading speed
FlasherTimer.Enabled = 1
Sub FlashInit
    Dim i
    For i = 0 to 200
        FlashState(i) = 0
        FlashLevel(i) = 0
    Next

    FlashSpeedUp = 50   ' fast speed when turning on the flasher
    FlashSpeedDown = 10 ' slow speed when turning off the flasher, gives a smooth fading

	SetLamp 111, 1

End Sub

Sub AllFlashOff
    Dim i
    For i = 0 to 200
        FlashState(i) = 0
    Next
End Sub

Sub Flash(nr, object)
    Select Case FlashState(nr)
        Case 0 'off
            FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown
            If FlashLevel(nr) < 0 Then
                FlashLevel(nr) = 0
                FlashState(nr) = -1 'completely off
            End if
            Object.opacity = FlashLevel(nr)
        Case 1 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp
            If FlashLevel(nr) > 255 Then
                FlashLevel(nr) = 255
                FlashState(nr) = -2 'completely on
            End if
            Object.opacity = FlashLevel(nr)
    End Select
End Sub

Sub Flashx(nr, object) 'For multiple flashers
            Object.opacity = FlashLevel(nr)
End Sub

Sub FlashVal(nr, object, value)
    Select Case FlashState(nr)
        Case 0 'off
            FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown
            If FlashLevel(nr) < 0 Then
                FlashLevel(nr) = 0
                FlashState(nr) = -1 'completely off
            End if
            Object.opacity = FlashLevel(nr)
        Case 1 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp
            If FlashLevel(nr) > value Then
                FlashLevel(nr) = value
                FlashState(nr) = -2 'completely on
            End if
            Object.opacity = FlashLevel(nr)
    End Select
End Sub
Sub FlashValm(nr, object, value)
	dim ftemp
    Select Case FlashState(nr)
        Case 0 'off
            ftemp = FlashLevel(nr) - FlashSpeedDown
            If ftemp < 0 Then
                ftemp = 0
            End if
            Object.opacity = ftemp
        Case 1 ' on
            ftemp = FlashLevel(nr) + FlashSpeedUp
            If ftemp > value Then
                ftemp = value
            End if
            Object.opacity = ftemp
    End Select
End Sub

Sub FlasherTimer_Timer()
Dim obj
	FlashVal  57,  f57, 255
	FlashVal  58,  f58, 255
	FlashVal  59,  f59, 255
	FlashVal  60,  f60, 255
	FlashVal  61,  f61, 255
	FlashVal 101, f101, 200
	FlashVal 102, f102, 200
	FlashVal 103, f103, 200
	FlashVal 104, f104, 255
	FlashVal 112, f112, 255
	FlashValm 113, f113_b, 255
	FlashVal 113, f113_a, 255
	FlashVal 199, F112X, 255 'set when 112 is set
	FlashVal 114, f114, 255
	FlashVal 115, f115, 255
	Flashx   115, F115X

'	FlashVal 111, aGILamps(0), 125
'	For each obj in aGILamps
'		Flashx 111, obj
'	Next
'	For each obj in aGILampRays
'		Flashx 111, obj
'	Next
End Sub


'****************************************************************************
'*****
'***** Primitive Animation subroutines					(gtxjoe v1.1)
'*****
'****************************************************************************
Const WallPrefix 		= "Sw" 'Change this based on your naming convention
Const PrimitivePrefix 	= "PrimSw"'Change this based on your naming convention
Const PrimitiveBumperPrefix = "BumperRing" 'Change this based on your naming convention
Dim primCnt(100), primDir(100), primBmprDir(6)


'****************************************************************************
'***** Primitive Rollover Animation
'USAGE: 	Sub sw18_Hit	: PrimRolloverHit   18, PrimSw18: End Sub 
'USAGE: 	Sub sw18_UnHit	: PrimRolloverUnHit 18, PrimSw18: End Sub
'****************************************************************************
Const RolloverMovementDir = "TransY" 
Const RolloverMovementInit = 0
Const RolloverMovementMax = -14	 

Sub PrimRolloverHit (swnum, primName): PrimRollover swnum, primName, 1: End Sub
Sub PrimRolloverUnHit (swnum, primName): PrimRollover swnum, primName, 0: End Sub

Sub PrimRollover (swnum, primName, isdropped)
	If isdropped = 1 Then
		PlaySound "rollover"
		Controller.Switch(swnum) = 1
		PrimMove primName, RolloverMovementDir, RolloverMovementMax
	Else
		Controller.Switch(swnum) = 0
		PrimMove primName, RolloverMovementDir, RolloverMovementInit
	End If
End Sub

'=== Generic Primitive Routines
Sub PrimMove (primName, primDir, val)
	Select Case primDir
		Case "TransX":  primName.TransX = val
		Case "TransY":  primName.TransY = val  
		Case "TransZ":  primName.TransZ = val  
	End Select
End Sub
'****************************************************************************
'***** Primitive StarRollover Animation
'USAGE: 	Sub sw17_Hit	: PrimStarRolloverHit   17, Sw17, PrimSw17: End Sub 
'USAGE: 	Sub sw17_UnHit	: PrimStarRolloverUnHit 17, Sw17, PrimSw17: End Sub

' or if delay needed use this
'USAGE: 	Sub sw17_Hit	: PrimStarRolloverHit   17, Sw17, PrimSw17: End Sub 
'USAGE: 	Sub sw17_UnHit	: PrimStarRolloverUnHitWithDelay 17, Sw17, PrimSw17, 200: End Sub
'USAGE:		Sub sw17_Timer	: PrimStarRolloverTimeout Sw17, PrimSw17: End Sub
'****************************************************************************
Const StarRolloverMovementDir = "TransY" 	'Specify direction
Const StarRolloverMovementInit = 0	 		'Specify starting position, usually 0
Const StarRolloverMovementMax = 6	 		'Amount to move when hit

Sub PrimStarRolloverHit (swnum, triggerName, primName): PrimStarRollover swnum, triggerName, primName, 1, 0: End Sub
Sub PrimStarRolloverUnHit (swnum, triggerName, primName): PrimStarRollover swnum, triggerName, primName, 0, 0: End Sub
Sub PrimStarRolloverUnHitWithDelay (swnum, triggerName, primName, delay): PrimStarRollover swnum, triggerName, primName, 0, delay: End Sub
Sub PrimStarRolloverTimeout (triggerName, primName)
	PrimMove primName, StarRolloverMovementDir, StarRolloverMovementInit
	triggerName.TimerEnabled = 0
End Sub

Sub PrimStarRollover (swNum, triggerName, primName, isDropped, upTimerDelay)
	PrimStarRolloverCustom swNum, triggerName, primName, isDropped, StarRolloverMovementDir, StarRolloverMovementInit, StarRolloverMovementMax, upTimerDelay
End Sub

Sub PrimStarRolloverCustom (swNum, triggerName, primName, isDropped, primDirection, startingPos, endingPos, upTimerDelay)
	If isdropped = 1 Then
		PlaySound "rollover"
		Controller.Switch(swnum) = 1
		PrimMove primName, primDirection, endingPos
	Else
		Controller.Switch(swnum)= 0
		If upTimerDelay = 0 Then
			PrimMove primName, primDirection, startingPos
		Else 'use timer before returning trigger to normal position
			triggerName.TimerInterval = upTimerDelay
			triggerName.TimerEnabled = 1
		End If
	End If
End Sub

Sub PrimStarRolloverCustomTimeout (triggerName, primName, startingPos)
	PrimMove primName, StarRolloverMovementDir, startingPos
	triggerName.TimerEnabled = 0
End Sub

'****************************************************************************
'***** Primitive Standup Target Animation
'****************************************************************************
'USAGE: 	Sub sw1_Hit: 	PrimStandupTgtHit  1, Sw1, PrimSw1: End Sub
'USAGE: 	Sub Sw1_Timer: 	PrimStandupTgtMove 1, Sw1, PrimSw1: End Sub

Const StandupTgtMovementDir = "TransZ" 
Const StandupTgtMovementMax = -6	 

Sub PrimStandupTgtHit (swnum, wallName, primName) 	
	PlaySound SoundFx("target", 1)
	vpmTimer.PulseSw swnum	
	primCnt(swnum) = 0 									'Reset count
	wallName.TimerInterval = 20 	'Set timer interval
	wallName.TimerEnabled = 1 	'Enable timer
End Sub

Sub	PrimStandupTgtMove (swnum, wallName, primName)
	Select Case StandupTgtMovementDir
		Case "TransX":
			Select Case primCnt(swnum)
				Case 0: 	primName.TransX = -StandupTgtMovementMax * .5
				Case 1: 	primName.TransX = -StandupTgtMovementMax
				Case 2: 	primName.TransX = -StandupTgtMovementMax * .5
				Case 3: 	primName.TransX = 0
				Case else: 	wallName.TimerEnabled = 0
			End Select
		Case "TransY":
			Select Case primCnt(swnum)
				Case 0: 	primName.TransY = -StandupTgtMovementMax * .5
				Case 1: 	primName.TransY = -StandupTgtMovementMax
				Case 2: 	primName.TransY = -StandupTgtMovementMax * .5
				Case 3: 	primName.TransY = 0
				Case else: 	wallName.TimerEnabled = 0
			End Select
		Case "TransZ":
			Select Case primCnt(swnum)
				Case 0: 	primName.TransZ = -StandupTgtMovementMax * .5
				Case 1: 	primName.TransZ = -StandupTgtMovementMax
				Case 2: 	primName.TransZ = -StandupTgtMovementMax * .5
				Case 3: 	primName.TransZ = 0
				Case else: 	wallName.TimerEnabled = 0
			End Select			
	End Select
	primCnt(swnum) = primCnt(swnum) + 1
End Sub 


'************************************************************************
'***** Primitive Drop Target Animation
'************************************************************************
'USAGE:  Sub Sw13_Hit: PrimDropTgtDown RBank, 3, 13, Sw13: End Sub 
'USAGE:  Sub Sw13_Timer: PrimDropTgtMove 13, Sw13, PrimSw13: End Sub
'USAGE:  Sub solRBankReset (enabled): If enabled Then PrimDropTgtUp RBank, 1, 13, Sw13, 1: PrimDropTgtUp RBank, 2, 12, Sw12, 0: PrimDropTgtUp RBank, 3, 11, Sw11, 0: End If: End Sub

Const DropTgtMovementDir = "TransY" 
Const DropTgtMovementMax = 50	 
	
Sub PrimDropTgtDown (targetbankname, targetbanknum, swnum, wallName)
	PrimDropTgtAnimate swnum, wallName, 0
	targetbankname.Hit targetbanknum

End Sub
Sub PrimDropTgtUp  (targetbankname, targetbanknum, swnum, wallName, resetvpmbank)
	PrimDropTgtAnimate swnum, wallName, 1
	If resetvpmbank = 1 Then targetbankname.SolDropUp True
End Sub
Sub PrimDropTgtSolDown  (targetbankname, targetbanknum, swnum, wallName, resetvpmbank)
	PrimDropTgtAnimate swnum, wallName, 0
	If resetvpmbank = 1 Then targetbankname.SolDropDown True
End Sub

Sub PrimDropTgtMove (swNum, wallName, primName) 'Customize direction as needed
	Select Case DropTgtMovementDir
		Case "TransX":
			If primDir(swNum) = 1 Then 'Up
				Select Case primCnt(swNum)
					Case 0: 	primName.TransX = -DropTgtMovementMax * .75
					Case 1: 	primName.TransX = -DropTgtMovementMax * .25
					Case 2,3,4: primName.TransX = 10
					Case 5: 	primName.TransX = 0
					Case else: 	wallName.TimerEnabled = 0
				End Select
			Else 'Down
				Select Case primCnt(swNum)
					Case 0: primName.TransX = -DropTgtMovementMax * .25
					Case 1: primName.TransX = -DropTgtMovementMax * .5
					Case 2: primName.TransX = -DropTgtMovementMax * .75
					Case 3: primName.TransX = -DropTgtMovementMax
					Case else: wallName.TimerEnabled = 0
				End Select
			End If
		Case "TransY":
			If primDir(swNum) = 1 Then 'Up
				Select Case primCnt(swNum)
					Case 0: 	primName.TransY = -DropTgtMovementMax * .75
					Case 1: 	primName.TransY = -DropTgtMovementMax * .25
					Case 2,3,4: primName.TransY = 10
					Case 5: 	primName.TransY = 0
					Case else: 	wallName.TimerEnabled = 0
				End Select
			Else 'Down
				Select Case primCnt(swNum)
					Case 0: primName.TransY = -DropTgtMovementMax * .25
					Case 1: primName.TransY = -DropTgtMovementMax * .5
					Case 2: primName.TransY = -DropTgtMovementMax * .75
					Case 3: primName.TransY = -DropTgtMovementMax
					Case else: wallName.TimerEnabled = 0
				End Select
			End If
		Case "TransZ":
			If primDir(swNum) = 1 Then 'Up
				Select Case primCnt(swNum)
					Case 0: 	primName.TransZ = -DropTgtMovementMax * .75
					Case 1: 	primName.TransZ = -DropTgtMovementMax * .25
					Case 2,3,4: primName.TransZ = 10
					Case 5: 	primName.TransZ = 0
					Case else: 	wallName.TimerEnabled = 0
				End Select
			Else 'Down
				Select Case primCnt(swNum)
					Case 0: primName.TransZ = -DropTgtMovementMax * .25
					Case 1: primName.TransZ = -DropTgtMovementMax * .5
					Case 2: primName.TransZ = -DropTgtMovementMax * .75
					Case 3: primName.TransZ = -DropTgtMovementMax
					Case else: wallName.TimerEnabled = 0
				End Select
			End If			
	End Select
	primCnt(swnum) = primCnt(swnum) + 1
End Sub

Sub PrimDropTgtAnimate  (swnum, wallName, dir)
	primCnt(swnum) = 0
	primDir(swnum) = dir
	wallName.TimerInterval = 10
	wallName.TimerEnabled = 1			
End Sub


'************************************************************************
'***** Primitive Bumper Animation 
'****************************************************************************
'USAGE:  Sub Bumper1_Hit(): PrimBumperHit 1, 40: End Sub
'USAGE:  Sub Bumper2_Hit(): PrimBumperHit 2, 39: End Sub

Const PrimBumperRingDir = "TransY" 
Const PrimBumperRingMovementMax  = 50
Const PrimBumperRingMovementStep = 10
Const PrimBumperTimeInterval     = 10

Sub PrimBumperHit (ringnum, swnum)
	vpmTimer.PulseSw swnum
	If Bumper1.Disabled = False Then
		PrimBmprDir(ringnum) = 1
		Execute "Bumper" & ringnum & "Timer.Interval = " & PrimBumperTimeInterval
		Execute "Bumper" & ringnum & "Timer.Enabled = 1" 		'Enable timer
		Select Case Int(Rnd*3)+1
			Case 1 : PlaySound Soundfx("bumper_1",1)
			Case 2 : PlaySound Soundfx("bumper_2",1)
			Case Else : PlaySound Soundfx("bumper_3",1)
		End Select
	End If
End Sub

Sub Bumper1Timer_Timer: PrimBumperMove(1): End Sub
Sub Bumper2Timer_Timer: PrimBumperMove(2): End Sub
Sub Bumper3Timer_Timer: PrimBumperMove(3): End Sub
Sub Bumper4Timer_Timer: PrimBumperMove(4): End Sub
Sub Bumper5Timer_Timer: PrimBumperMove(5): End Sub

Sub PrimBumperMove (ringnum)
	Dim tempZ
	Execute "tempZ = " & PrimitiveBumperPrefix & ringnum & "." & PrimBumperRingDir
	If PrimBmprDir(ringnum) = 1 and tempZ > -PrimBumperRingMovementMax then 'Go down
		tempZ = tempZ - PrimBumperRingMovementStep
	ElseIf PrimBmprDir(ringnum) = 0 and tempZ > 0 Then 'Stop timer
		tempZ = 0
		Execute "Bumper" & ringnum & "Timer.Enabled = 0"
	Else 'Go up
		PrimBmprDir(ringnum) = 0
		tempZ = tempZ + PrimBumperRingMovementStep
	End If
	Execute PrimitiveBumperPrefix & ringnum & "." & PrimBumperRingDir & "=" & tempZ
	'debug.print bumperring1.transy & "," & bumperring2.transy
End Sub

''*************************************************************
''Toggle DOF sounds on/off based on cController value
''*************************************************************
'Dim ToggleMechSounds
'Function SoundFX (sound)
'    If cController = 4 and ToggleMechSounds = 0 Then
'        SoundFX = ""
'    Else
'        SoundFX = sound
'    End If
'End Function

dim incup
sub trigger1_hit
	activeball.vely = activeball.vely  * 2
	activeball.velz = activeball.velz * 2
'	activeball.vely = activeball.vely  * (1.3+incup)
'	activeball.velz = activeball.velz * (1.3+incup)
'	debug.print incup & ": " & activeball.vely & ":" & activeball.velx
'	incup = incup + .1
end sub

sub trigger3_hit
	activeball.vely = activeball.vely  * 1.1
	activeball.velz = activeball.velz * 1.1
'	activeball.vely = activeball.vely  * (1.3+incup)
'	activeball.velz = activeball.velz * (1.3+incup)
	debug.print incup & ":1 " & activeball.vely & ":" & activeball.velx
'	incup = incup + .1
end sub
sub trigger4_hit
	activeball.vely = activeball.vely  * 1.1
	activeball.velz = activeball.velz * 1.1
'	activeball.vely = activeball.vely  * (1.3+incup)
'	activeball.velz = activeball.velz * (1.3+incup)
	debug.print incup & ":2 " & activeball.vely & ":" & activeball.velx
'	incup = incup + .1
end sub
sub trigger5_hit
	activeball.vely = activeball.vely  * 1.06
	activeball.velz = activeball.velz * 1.06
'	activeball.vely = activeball.vely  * (1.3+incup)
'	activeball.velz = activeball.velz * (1.3+incup)
	debug.print incup & ":3 " & activeball.vely & ":" & activeball.velx
'	incup = incup + .1
end sub
sub trigger6_hit
	activeball.vely = activeball.vely  * 1.02
	activeball.velz = activeball.velz * 1.02
'	activeball.vely = activeball.vely  * (1.3+incup)
'	activeball.velz = activeball.velz * (1.3+incup)
	debug.print incup & ":4 " & activeball.vely & ":" & activeball.velx
'	incup = incup + .1
end sub
sub trigger2_hit
	trigger2.destroyball
end sub



'Sub Table1_KeyDown(ByVal keycode)
'
'	If keycode = PlungerKey Then
'		Plunger.PullBack
'		PlaySound "plungerpull",0,1,0.25,0.25
'	End If
'
'	If keycode = LeftFlipperKey Then
'		LeftFlipper.RotateToEnd
'		PlaySound "fx_flipperup", 0, .67, -0.05, 0.05
'	End If
'    
'	If keycode = RightFlipperKey Then
'		RightFlipper.RotateToEnd
'		PlaySound "fx_flipperup", 0, .67, 0.05, 0.05
'	End If
'    
'	If keycode = LeftTiltKey Then
'		Nudge 90, 2
'	End If
'    
'	If keycode = RightTiltKey Then
'		Nudge 270, 2
'	End If
'    
'	If keycode = CenterTiltKey Then
'		Nudge 0, 2
'	End If
'    
'End Sub
'
'Sub Table1_KeyUp(ByVal keycode)
'
'	If keycode = PlungerKey Then
'		Plunger.Fire
'		PlaySound "plunger",0,1,0.25,0.25
'	End If
'    
'	If keycode = LeftFlipperKey Then
'		LeftFlipper.RotateToStart
'		PlaySound "fx_flipperdown", 0, 1, -0.05, 0.05
'	End If
'    
'	If keycode = RightFlipperKey Then
'		RightFlipper.RotateToStart
'		PlaySound "fx_flipperdown", 0, 1, 0.05, 0.05
'	End If
'
'End Sub
'
'Sub Drain_Hit()
'	PlaySound "drain",0,1,0,0.25
'	Drain.DestroyBall
'	BIP = BIP - 1
'	If BIP = 0 then
'		'Plunger.CreateBall
'		BallRelease.CreateBall
'		BallRelease.Kick 90, 8
'		PlaySound "ballrelease",0,1,0,0.25
'		BIP = BIP + 1
'	End If
'End Sub

'Dim BIP
'BIP = 0
'
'Sub LeftSlingshot_Slingshot()
'	PlaySound "fx_bumper4",0,1,-0.15,0.25
'End Sub
'
'Sub Plunger_Init()
'	PlaySound "ballrelease",0,0.5,0.5,0.25
'	'Plunger.CreateBall
'	BallRelease.CreateBall
'	BallRelease.Kick 90, 8
'	BIP = BIP +1
'End Sub
'
'Sub Gate_Hit
'	Kicker1.Kick 190, 10
'End Sub
'
'Sub Bumper1_Hit
'	PlaySound "fx_bumper4"
'	B1L1.State = 1:B1L2. State = 1
'	Me.TimerEnabled = 1
'End Sub
'
'Sub Bumper1_Timer
'	B1L1.State = 0:B1L2. State = 0
'	Me.Timerenabled = 0
'End Sub
'
'Sub Bumper2_Hit
'	PlaySound "fx_bumper4"
'	B2L1.State = 1:B2L2. State = 1
'	Me.TimerEnabled = 1
'End Sub
'
'Sub Bumper2_Timer
'	B2L1.State = 0:B2L2. State = 0
'	Me.Timerenabled = 0
'End Sub	
'
'Sub Bumper3_Hit
'	PlaySound "fx_bumper4"
'	B3L1.State = 1:B3L2. State = 1
'	Me.TimerEnabled = 1
'End Sub
'
'Sub Bumper3_Timer
'	B3L1.State = 0:B3L2. State = 0
'	Me.Timerenabled = 0
'End Sub
'
'Sub Bumper4_Hit
'	PlaySound "fx_bumper4"
'	B4L1.State = 1:B4L2. State = 1
'	Me.TimerEnabled = 1
'End Sub
'
'Sub Bumper4_Timer
'	B4L1.State = 0:B4L2. State = 0
'	Me.Timerenabled = 0
'End Sub
'
'Sub Bumper5_Hit
'	PlaySound "fx_bumper4"
'	B5L1.State = 1:B5L2. State = 1
'	Me.TimerEnabled = 1
'End Sub
'
'Sub Bumper5_Timer
'	B5L1.State = 0:B5L2. State = 0
'	Me.Timerenabled = 0
'End Sub
'
'Sub sw9_Hit
'	If L9.State = 1 then 
'		L9.State  = 0
'	else 
'		L9.State = 1
'	end if
'End Sub
'
'Sub sw8_Hit
'	If L8.State = 1 then 
'		L8.State  = 0
'	else 
'		L8.State = 1
'	end if
'End Sub
'
'Sub sw7_Hit
'	If L7.State = 1 then 
'		L7.State  = 0
'	else 
'		L7.State = 1
'	end if
'End Sub
'
'Sub sw6_Hit
'	If L6.State = 1 then 
'		L6.State  = 0
'	else 
'		L6.State = 1
'	end if
'End Sub
'
'
''****Targets
'Sub sw1_Hit
'	If L1.State = 1 then 
'		L1.State  = 0
'	else 
'		L1.State = 1
'	end if
'	sw1p.transx = -10
'	Me.TimerEnabled = 1
'End Sub
'
'Sub sw1_Timer
'	sw1p.transx = 0
'	Me.TimerEnabled = 0
'End Sub
'
'Sub sw2_Hit
'	If L2.State = 1 then 
'		L2.State  = 0
'	else 
'		L2.State = 1
'	end if
'	sw2p.transx = -10
'	Me.TimerEnabled = 1
'End Sub
'
'Sub sw2_Timer
'	sw2p.transx = 0
'	Me.TimerEnabled = 0
'End Sub
'
'Sub sw3_Hit
'	If L3.State = 1 then 
'		L3.State  = 0
'	else 
'		L3.State = 1
'	end if
'	sw3p.transx = -10
'	Me.TimerEnabled = 1
'End Sub
'
'Sub sw3_Timer
'	sw3p.transx = 0
'	Me.TimerEnabled = 0
'End Sub
'
'Sub sw11_Hit
'	If L11.State = 1 then 
'		L11.State  = 0
'	else 
'		L11.State = 1
'	end if
'	sw11p.transx = -10
'	Me.TimerEnabled = 1
'End Sub
'
'Sub sw11_Timer
'	sw11p.transx = 0
'	Me.TimerEnabled = 0
'End Sub
'
'Sub sw12_Hit
'	If L12.State = 1 then 
'		L12.State  = 0
'	else 
'		L12.State = 1
'	end if
'	sw12p.transx = -10
'	Me.TimerEnabled = 1
'End Sub
'
'Sub sw12_Timer
'	sw12p.transx = 0
'	Me.TimerEnabled = 0
'End Sub
'
'Sub sw13_Hit
'	If L13.State = 1 then 
'		L13.State  = 0
'	else 
'		L13.State = 1
'	end if
'	sw13p.transx = -10
'	Me.TimerEnabled = 1
'End Sub
'
'Sub sw13_Timer
'	sw13p.transx = 0
'	Me.TimerEnabled = 0
'End Sub
'
'
'Sub Kicker1_Hit
'	PlaySound "kicker_enter_center", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
'	'Plunger.CreateBall
'	PlaySound "ballrelease",0,0.5,0.5,0.25
'	BallRelease.CreateBall
'	BallRelease.Kick 90, 8
'	BIP = BIP +1
'End Sub
'
'Sub Kicker1_UnHit
'	PlaySound "popper_ball",0,.75,0,0.25
'End Sub
'
''*****GI Lights On
'dim xx
'
'For each xx in GI:xx.State = 1: Next
'
''**********Sling Shot Animations
'' Rstep and Lstep  are the variables that increment the animation
''****************
'Dim RStep, Lstep
'
'Sub RightSlingShot_Slingshot
'    PlaySound "left_slingshot", 0, 1, 0.05, 0.05
'    RSling.Visible = 0
'    RSling1.Visible = 1
'    sling1.TransZ = -20
'    RStep = 0
'    RightSlingShot.TimerEnabled = 1
'	gi1.State = 0:Gi2.State = 0
'End Sub
'
'Sub RightSlingShot_Timer
'    Select Case RStep
'        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
'        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0:gi1.State = 1:Gi2.State = 1
'    End Select
'    RStep = RStep + 1
'End Sub
'
'Sub LeftSlingShot_Slingshot
'    PlaySound "right_slingshot",0,1,-0.05,0.05
'    LSling.Visible = 0
'    LSling1.Visible = 1
'    sling2.TransZ = -20
'    LStep = 0
'    LeftSlingShot.TimerEnabled = 1
'	gi3.State = 0:Gi4.State = 0
'End Sub
'
'Sub LeftSlingShot_Timer
'    Select Case LStep
'        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
'        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0:gi3.State = 1:Gi4.State = 1
'    End Select
'    LStep = LStep + 1
'End Sub



' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function Vol2(ball1, ball2) ' Calculates the Volume of the sound based on the speed of two balls
    Vol2 = (Vol(ball1) + Vol(ball2) ) / 2
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / table1.width-1
    If tmp> 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
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

Const tnob = 8 ' total number of balls
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
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0
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
	PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
End Sub

Sub Targets_Hit (idx)
	PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
End Sub

Sub Metals_Thin_Hit (idx)
	PlaySound "metalhit_thin", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub Metals_Medium_Hit (idx)
	PlaySound "metalhit_medium", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub Metals2_Hit (idx)
	PlaySound "metalhit2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

'Sub Gates_Hit (idx)
'	PlaySound "gate4", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
'End Sub

Sub Spinner_Spin
	PlaySound "fx_spinner",0,.25,0,0.25
End Sub

Sub Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then 
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub Posts_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 16 then 
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End if
	If finalspeed >= 6 AND finalspeed <= 16 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
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
		Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End Select
End Sub


'set lights(3)=l3
'set lights(4)=l4
'set lights(5)=l5
'set lights(6)=l6
'set lights(7)=l7
'set lights(8)=l8
'set lights(9)=l9
'set lights(10)=l10
'set lights(11)=l11
'set lights(12)=l12
'set lights(13)=l13
'set lights(14)=l14
'set lights(15)=l15
'set lights(16)=l16
'set lights(17)=L17
'set lights(18)=L18
'set lights(19)=L19

Sub BalldropR_Hit
	Playsound "Drop_Right"
	Activeball.vely=0
End Sub

Sub BalldropR1_Hit
	Playsound "Drop_Right"
	Activeball.vely=0
End Sub

Sub BalldropL_Hit
	Playsound "Drop_Left"
	Activeball.vely=0

End Sub
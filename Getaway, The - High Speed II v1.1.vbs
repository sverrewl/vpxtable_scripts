'*******************************************************************************
'   _____ ______ _______  __          __ __     __ _    _  _____ _____ _____
'  / ____|  ____|__   __|/\ \        / /\\ \   / /| |  | |/ ____|_   _|_   _|
' | |  __| |__     | |  /  \ \  /\  / /  \\ \_/ (_) |__| | (___   | |   | |
' | | |_ |  __|    | | / /\ \ \/  \/ / /\ \\   /  |  __  |\___ \  | |   | |
' | |__| | |____   | |/ ____ \  /\  / ____ \| |  _| |  | |____) |_| |_ _| |_
'  \_____|______|  |_/_/    \_\/  \/_/    \_\_| (_)_|  |_|_____/|_____|_____|
'
'*******************************************************************************
'
'The Getaway: High Speed II - Williams 1992
'
'Credits:
'
'Uses cFastFlips by nFozzy
'
'Build with VPX 10.4
'
'Enjoy and remember, winners don't drive drunk.
'
'Big thanks for help to flupper1
'
'flupper1 - supercharger, sc ramps, bumper caps, EE, flipper prims and shadows, playfield shadows
'32assassin - stripped VP9 table, script edit
'ganjafarmer - 2d graphics, physics, lighting, some simple prims
'nFozzy - fast flips script
'bassgeige - backglass
'
'Visit www.vpforums.org for more great pinball recreations
'
'ROM: http://www.vpforums.org/index.php?app=downloads&showfile=1330
'WHEEL: http://www.vpforums.org/index.php?app=downloads&showfile=6610
'BACKGLASS: http://www.vpforums.org/index.php?app=downloads&showfile=11010
'
'V1.0 - First release
'V1.01 - improved physics (thanks nFozzy)
'
'WIP topic: http://www.vpforums.org/index.php?showtopic=38595
'
'Enjoy

' Thalamus 2018-07-23
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' No special SSF tweaks yet.

Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const BallSize = 51
Const BallMass = 1.2

Const cGameName="gw_l5", UseSolenoids=2, UseLamps=0, UseGI=0, SSolenoidOn="SolOn", SSolenoidOff="SolOff", SCoin="coin"

LoadVPM "01560000", "WPC.VBS", 3.46

Dim DesktopMode: DesktopMode = Table1.ShowDT

Dim ramptrig:ramptrig=0

dim FastFlips

Dim Wbl: Wbl=1 ' wobble the supercharger on each loop

If Wbl=0 Then
	Trigger10.enabled=0
End If

If DesktopMode = True Then 'Show Desktop components
Ramp16.visible=1
Ramp15.visible=1
Primitive13.visible=1
refL.visible=0
refR.visible=0
f121c.visible=0
f123c.visible=0
f122b.visible=0
f117b.visible=0
f66b.visible=0
f67b.visible=0
Else
Ramp16.visible=0
Ramp15.visible=0
Primitive13.visible=0
refL.visible=1
refR.visible=1
End if

'*************************************************************
'Solenoid Call backs
'**********************************************************************************************************
SolCallback(2)  = "SolRampUp"
SolCallback(3)  = "SolRampDown"
SolCallback(4)	= "LockPost"
SolCallback(7)  =   "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(8)  = "SolKickback"              'Kickback
SolCallback(9)  = "bsEjectHole.SolOut"
SolCallback(10) = "SuperchargerDiverter"
SolCallback(11) = "bsTrough.SolOut"
SolCallback(12) = "SolPlunger"   'AutoPlunger
SolCallback(16) = "bsTrough.SolIn"

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"
SolCallback(sURFlipper) = "SolURFlipper"

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFX("fx_Flipperup",DOFContactors):LeftFlipper.RotateToEnd
     Else
         PlaySound SoundFX("fx_Flipperdown",DOFContactors):LeftFlipper.RotateToStart
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFX("fx_Flipperup",DOFContactors):RightFlipper.RotateToEnd
     Else
         PlaySound SoundFX("fx_Flipperdown",DOFContactors):RightFlipper.RotateToStart
     End If
End Sub

 Sub SolURFlipper(Enabled)
	 If Enabled Then
		 'PlaySound SoundFX("fx_Flipperup",DOFContactors)
		RightFlipper1.RotateToEnd
	 Else
		 PlaySound SoundFX("fx_Flipperdown",DOFContactors):RightFlipper1.RotateToStart
	 End If
 End Sub


SolCallback(17) = "SetLamp 117," 'Right Middle PF
SolCallback(18) = "SetLamp 118," ' Motor toy lights
SolCallback(19) = "SetLamp 119," 'Right Slingshot
SolCallback(20) = "SetLamp 120," 'PF
SolCallback(21) = "SetLamp 121," 'left Middle PF
SolCallback(22) = "SetLamp 122," 'left Middle PF
SolCallback(23) = "SetLamp 123," 'Right Middle PF
SolCallback(24) = "SetLamp 124," 'Left Slingshot
SolCallback(27) = "Siren"
SolCallBack(31) = "FastFlips.TiltSol"
'**********************************************************************************************************

'Solenoid Controlled toys
'**********************************************************************************************************

Sub SolRampUp(Enabled)
    If Enabled Then
        Controller.Switch(54) = False
        RampaMovil.Collidable = False
        playsound SoundFX("fx_diverterUp",DOFContactors)
		div_ramp.ObjRotY=15
		ramptrig=0
    End If
End Sub

Sub SolRampDown(Enabled)
    If Enabled Then
        Controller.Switch(54) = true
        RampaMovil.Collidable = True
        playsound SoundFX("fx_diverterUp",DOFContactors)
		div_ramp.ObjRotY=0
		ramptrig=1
    End If
End sub

Sub SuperchargerDiverter(enabled)
	If Enabled Then
		Wall29.collidable=True
		sc_div.ObjRotZ=0
		PlaySound SoundFX("fx_DiverterUp",DOFContactors)
	Else
		Wall29.collidable=False
		sc_div.ObjRotZ=-43
		PlaySound SoundFX("sc_diverter",DOFContactors)
	End If
End Sub

'plunger
Sub SolPlunger(Enabled)
	If Enabled Then
		plungerIM.AutoFire
	End If
End Sub

'kickback
Sub SolKickback(Enabled)
    If enabled Then
       Plunger1.Fire
       PlaySound SoundFX("Popper",DOFContactors)
    Else
       Plunger1.PullBack
    End If
End Sub

'kicker ball bounce

'hit to enable wall lock
Sub e_kick_Hit()
	ek.enabled=1
	PlaySound "metalhit_medium"
	PlaySound "fx_rr6"
End Sub

'lock ball between walls (80ms)
Sub ek_Timer
	l_kick.collidable=1
	kk.enabled=1
	ek.enabled=0
End Sub

'destroy walls and enable real kicker (400ms)
Sub kk_timer
	sw77.enabled=1
	l_kick.collidable=0
	kk.enabled=0
End Sub

'disable real kicker after release
Sub k_off_Hit()
	sw77.enabled=0
End Sub


'Visible Lock, original PD

  dim postdown : postdown = false

  Sub LockPost(enabled)
	visibleLock.Solexit enabled
  	If enabled then
  		If postdown = false Then : PosteArriba.IsDropped=1 : postlock.z=-10 : Playsound "sc_diverter" : vpmTimer.AddTimer 400, "RaisePost" : End If
  		postdown = true
  	End If
  End Sub

 Sub RaisePost(aSw) : If postdown = true Then : PosteArriba.IsDropped = 0 : postlock.z=45 : PlaySound "sc_diverter" : End If : postdown = false : End Sub

'**************************
'           GI
'**************************

Set GiCallback2 = GetRef("UpdateGI")
Sub UpdateGI(no, Enabled)
	If Enabled Then
		dim xx
		For each xx in GI:xx.State = 1: Next
		For each xx in Plastics:xx.BlendDisableLighting = .41: Next
		batleft.BlendDisableLighting = .11:batright.BlendDisableLighting = .11:batright1.BlendDisableLighting = .11
		Primitive35.BlendDisableLighting = .11:Primitive36.BlendDisableLighting = .11 'sling cover
		Primitive11.BlendDisableLighting = .11:Primitive12.BlendDisableLighting = .11
		l66.BlendDisableLighting = .11:l67.BlendDisableLighting = .11
		Wall26.BlendDisableLighting = .15 'ramp decal
		Primitive45.BlendDisableLighting = .21 : Primitive46.BlendDisableLighting = .31' drive 65
		donut.BlendDisableLighting = .11 ' donut heaven
		dome1.BlendDisableLighting = .15:dome2.BlendDisableLighting = .15:dome3.BlendDisableLighting = .15 ' traffic lights
		supercharger_p.BlendDisableLighting = .20:superramp_p.BlendDisableLighting = 1 ' supercharger and ramp
		refL.opacity=7:refR.opacity=7
        PlaySound "fx_relay"
		DOF 101, DOFOn
	Else
		For each xx in GI:xx.State = 0: Next
		For each xx in Plastics:xx.BlendDisableLighting =0: Next
		batleft.BlendDisableLighting =0:batright.BlendDisableLighting =0:batright1.BlendDisableLighting =0
		Primitive35.BlendDisableLighting =0:Primitive36.BlendDisableLighting =0 'sling cover
		Primitive11.BlendDisableLighting = 0:Primitive12.BlendDisableLighting = 0
		l66.BlendDisableLighting =0:l67.BlendDisableLighting =0
		Wall26.BlendDisableLighting = 0 'ramp decal
		Primitive45.BlendDisableLighting = 0 : Primitive46.BlendDisableLighting = 0 ' drive 65
		donut.BlendDisableLighting = 0 ' donut heaven
		dome1.BlendDisableLighting = 0:dome2.BlendDisableLighting = 0:dome3.BlendDisableLighting = 0 ' traffic lights
		supercharger_p.BlendDisableLighting = 0:superramp_p.BlendDisableLighting = 0 ' supercharger and ramp
		refL.opacity=0:refR.opacity=0
        'PlaySound "fx_relay"
		DOF 101, DOFOff
	End If
End Sub

Sub Siren(enabled)
   	If enabled Then

	Else

	End If
End Sub


'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************
Dim bsTrough, bsEjectHole, visibleLock
Dim bsSwitch74, bsSwitch75, bsSwitch76

Sub Table1_Init
	vpmInit Me
	On Error Resume Next
		With Controller
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
		.SplashInfoLine = "Gateway High Speed II Williams"&chr(13)&"You Suck"
		.HandleMechanics=0
		.HandleKeyboard=0
		.ShowDMDOnly=1
		.ShowFrame=0
		.ShowTitle=0
        .hidden = 0
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With
     On Error Goto 0

	PinMAMETimer.Interval=PinMAMEInterval
	PinMAMETimer.Enabled=1

	Set FastFlips = new cFastFlips
	with FastFlips
	.CallBackL = "SolLflipper"  'Point these to flipper subs
	.CallBackR = "SolRflipper"  '...
	'  .CallBackUL = "SolULflipper"'...(upper flippers, if needed)
	.CallBackUR = "SolURflipper"'...
	.TiltObjects = True 'Optional, if True calls vpmnudge.solgameon automatically. IF YOU GET A LINE 1 ERROR, DISABLE THIS! (or setup vpmNudge.TiltObj!)
	'  .InitDelay "FastFlips", 100         'Optional, if > 0 adds some compensation for solenoid jitter (occasional problem on Bram Stoker's Dracula)
	'  .DebugOn = False        'Debug, always-on flippers. Call FastFlips.DebugOn True or False in debugger to enable/disable.
	end with

    vpmNudge.TiltSwitch = 14
    vpmNudge.Sensitivity = 2
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingshot)

 	Set bsTrough=New cvpmBallStack
		bsTrough.InitSw 55,58,57,56,0,0,0,0
		bsTrough.InitKick BallRelease,85,7
		bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
		bsTrough.Balls=3

 	Set bsEjectHole=New cvpmBallStack
 		bsEjectHole.InitSaucer sw77,77,100,10
 		bsEjectHole.InitExitSnd SoundFX("fx_saucer_exit",DOFContactors), SoundFX("Solenoid",DOFContactors)

	Set visibleLock = New cvpmVLock
		visibleLock.InitVLock Array(sw76,sw75, sw74),Array(k76, k75, k74), Array(76,75,74)
		visibleLock.InitSnd SoundFX("sc_diverter",DOFContactors), SoundFX("Solenoid",DOFContactors)
		visibleLock.ExitDir = 180
		visibleLock.ExitForce = 0
		visibleLock.createevents "visibleLock"

   	Controller.Switch(22) = True  ' Coin Door Closed
  	Controller.Switch(24) = True  ' Always Closed Switch

  	Set bsSwitch74 = New cvpmBallStack : bsSwitch74.Initsw 0,0,0,0,0,0,0,0
  	Set bsSwitch75 = New cvpmBallStack : bsSwitch75.Initsw 0,0,0,0,0,0,0,0
  	Set bsSwitch76 = New cvpmBallStack : bsSwitch76.Initsw 0,0,0,0,0,0,0,0

	Plunger1.PullBack

End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)
	If KeyDownHandler(keycode) Then Exit Sub
	If keycode = PlungerKey Then Controller.Switch(34) = True
	if KeyCode = LeftTiltKey Then Nudge 90, 4
	if KeyCode = RightTiltKey Then Nudge 270, 4
	if KeyCode = CenterTiltKey Then Nudge 0, 12
	If keycode = LeftMagnaSave Then:Controller.Switch(33) = 1:End If
	If KeyCode = LeftFlipperKey then FastFlips.FlipL True :  FastFlips.FlipUL True
	If KeyCode = RightFlipperKey then FastFlips.FlipR True :  FastFlips.FlipUR True
'(Do not use Exit Sub, this script does not handle switch handling at all!)
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
	If KeyUpHandler(keycode) Then Exit Sub
	If keycode = PlungerKey Then Controller.Switch(34) = False
	If keycode = LeftMagnaSave Then:Controller.Switch(33) = False:End If
	If KeyCode = LeftFlipperKey then FastFlips.FlipL False :  FastFlips.FlipUL False
	If KeyCode = RightFlipperKey then FastFlips.FlipR False :  FastFlips.FlipUR False
End Sub

Dim plungerIM
	Set plungerIM = New cvpmImpulseP
	With plungerIM
		.InitImpulseP ShooterLane, 46, 0.1
		.Random 0.2
		.Switch 78
		.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
		.CreateEvents "plungerIM"
	End With


'**********************************************************************************************************
 ' Drain hole and kickers
Sub Drain_Hit:bsTrough.addball me : playsound"drain" : End Sub
Sub sw77_Hit(): bsEjectHole.Addball me : playsound "kicker_enter_center": End Sub

'Bumpers
Sub Bumper1_Hit : vpmTimer.PulseSw(61) : playsound SoundFX("RightJet",DOFContactors): End Sub
Sub Bumper2_Hit : vpmTimer.PulseSw(62) : playsound SoundFX("LeftJet",DOFContactors): End Sub
Sub Bumper3_Hit : vpmTimer.PulseSw(63) : playsound SoundFX("BottomJet",DOFContactors): End Sub

'Wire Triggers
Sub SW15_Hit:Controller.Switch(15)=1 : playsound"rollover" : End Sub
Sub SW15_unHit:Controller.Switch(15)=0:End Sub
Sub SW16_Hit:Controller.Switch(16)=1 : playsound"rollover" : End Sub
Sub SW16_unHit:Controller.Switch(16)=0:End Sub
Sub SW17_Hit:Controller.Switch(17)=1 : playsound"rollover" : End Sub
Sub SW17_unHit:Controller.Switch(17)=0:End Sub
Sub SW18_Hit:Controller.Switch(18)=1 : playsound"rollover" : End Sub
Sub SW18_unHit:Controller.Switch(18)=0:End Sub
Sub SW25_Hit:Controller.Switch(25)=1 : playsound"rollover" : End Sub
Sub SW25_unHit:Controller.Switch(25)=0:End Sub
Sub SW26_Hit:Controller.Switch(26)=1 : playsound"rollover" : End Sub
Sub SW26_unHit:Controller.Switch(26)=0:End Sub
Sub SW27_Hit:Controller.Switch(27)=1 : playsound"rollover" : End Sub
Sub SW27_unHit:Controller.Switch(27)=0:End Sub
Sub SW28_Hit:Controller.Switch(28)=1 : playsound"rollover" : End Sub
Sub SW28_unHit:Controller.Switch(28)=0:End Sub
Sub SW71_Hit:Controller.Switch(71)=1 : playsound"rollover" : End Sub
Sub SW71_unHit:Controller.Switch(71)=0:End Sub
Sub SW72_Hit:Controller.Switch(72)=1 : playsound"rollover" : End Sub
Sub SW72_unHit:Controller.Switch(72)=0:End Sub
Sub SW73_Hit:Controller.Switch(73)=1 : playsound"rollover" : End Sub
Sub SW73_unHit:Controller.Switch(73)=0:End Sub

Sub lock74_Hit:stopsound "fx_metalrolling" : playsound"fx_collide" : End Sub

 'Stand Up Targets
 Sub sw36_Hit: vpmTimer.pulseSw 36:End Sub
 Sub sw37_Hit: vpmTimer.pulseSw 37:End Sub
 Sub sw38_Hit: vpmTimer.pulseSw 38:End Sub
 Sub sw41_Hit: vpmTimer.pulseSw 41:End Sub
 Sub sw42_Hit: vpmTimer.pulseSw 42:End Sub
 Sub sw43_Hit: vpmTimer.pulseSw 43:End Sub
 Sub sw51_Hit: vpmTimer.pulseSw 51:End Sub
 Sub sw52_Hit: vpmTimer.pulseSw 52:End Sub
 Sub sw53_Hit: vpmTimer.pulseSw 53:End Sub
 Sub sw44_Hit: vpmTimer.pulseSw 44:End Sub
 Sub sw45_Hit: vpmTimer.pulseSw 45:End Sub
 Sub sw46_Hit: vpmTimer.pulseSw 46:End Sub
 Sub sw86_Hit: vpmTimer.pulseSw 86:End Sub
 Sub sw87_Hit: vpmTimer.pulseSw 87:End Sub
 Sub sw88_Hit: vpmTimer.pulseSw 88:End Sub

 'Ramp Triggers
Sub SW65_Hit:Controller.Switch(65)=1 : playsound"fx_metalrolling" : End Sub
Sub SW65_unHit:Controller.Switch(65)=0:End Sub
Sub SW67_Hit:Controller.Switch(67)=1 : playsound"rollover" : End Sub
Sub SW67_unHit:Controller.Switch(67)=0:End Sub
Sub SW84_Hit:Controller.Switch(84)=1 : playsound"rollover" : End Sub
Sub SW84_unHit:Controller.Switch(84)=0:End Sub

 'supercharger

Sub SW81_Hit
	Controller.Switch(81)=true
	'activeball.velY=2
	activeball.velX=activeball.velX+11
End Sub

Sub SW81_unHit
	Controller.Switch(81)=false
End Sub

Sub SW82_Hit
	Controller.Switch(82)=true
	'activeball.velY=3
	activeball.velX=activeball.velX+15
End Sub

Sub SW82_unHit
	Controller.Switch(82)=false
End Sub

Sub SW83_Hit
	Controller.Switch(83)=true
	'activeball.velY=5
	activeball.velX=activeball.velX+15
End Sub

Sub SW83_unHit
	Controller.Switch(83)=false
End Sub

Sub SW85_Hit
	vpmTimer.pulseSw 85
	'Controller.Switch(85)=true
	playsound"sc_loop2"
End Sub

Sub SW85_unHit
	'Controller.Switch(85)=false
End Sub

Sub Trigger11_Hit()
	Wall30.collidable=1
End Sub

Sub Wall30_hit()
	Wall30.collidable=0
	playSound "gate"
	drop.enabled=0
	drop.enabled=1
End Sub

Sub Trigger12_Hit()
	Wall31.collidable=1
End Sub

Sub Wall31_hit()
	Wall31.collidable=0
	playSound "gate"
	drop.enabled=0
	drop.enabled=1
End Sub

Sub Trigger9_Hit()
	Primitive29.transz=1
	wobble1.enabled=1
End Sub

Sub Trigger10_Hit()
	supercharger_p.transx=-2
	primitive41.transx=-2
	dome118.transx=-2
	dome121.transx=-2
	dome123.transx=-2
	bolt12.transx=2
	bolt13.transx=2
	wobble1.enabled=1
End Sub

Sub wobble1_Timer
	Primitive29.transz=0
	supercharger_p.transx=0
	primitive41.transx=0
	dome118.transx=0
	dome121.transx=0
	dome123.transx=0
	bolt12.transx=0
	bolt13.transx=0
	wobble1.enabled=0
End Sub

Sub Trigger1_Hit():activeball.velY=activeball.velY * 0.8:playSound "fx_metalrolling":End Sub

Sub Trigger2_Hit():stopSound "fx_shortmetal" : End Sub

Sub Trigger3_Hit():stopSound "fx_metalrolling" : End Sub

Sub Trigger4_Hit():playSound "fx_rr7":End Sub

Sub Trigger5_Hit():playSound "fx_rrenter":End Sub

Sub Trigger6_Hit()
	If ramptrig=1 Then
		playSound "fx_rrenter"
	End If
End Sub

Sub Trigger7_Hit():playSound "fx_shortmetal":End Sub

Sub Trigger8_Hit():playSound "fx_rr7":End Sub

Sub Gate1_Hit():playSound "gate":End Sub

Sub Gate5_Hit():playSound "gate":End Sub

Sub BallReleaseGate_Hit():playSound "gate":End Sub

Sub drop_Timer
	playSound "fx_ballrampdrop"
	drop.enabled=0
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
LampTimer.Interval = 5 'lamp fading speed
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
     NFadeLm 11, l11
	NFadeL 11, l11a
     NFadeLm 12, l12
	NFadeL 12, l12a
     NFadeLm 13, l13
	NFadeL 13, l13a
     NFadeLm 14, l14
	NFadeL 14, l14a
     NFadeLm 15, l15
	NFadeL 15, l15a
     NFadeLm 16,l16
     NFadeLm 16, l116
	NFadeL 16, l116a
     NFadeLm 17, l17
	NFadeL 17, l17a
     NFadeLm 18,l118
     NFadeLm 18, l18
	NFadeLm 18, l18a
	NFadeL 18, l118a

      NFadeLm 21, l21
	NFadeLm 21, l21a
      NFadeLm 22, l22
	NFadeLm 22, l22a
      NFadeLm 23, l23
	NFadeL 23, l23a
      NFadeLm 24, l24
	NFadeLm 24, l24a
      NFadeLm 25, l25
	NFadeLm 25, l25a
      NFadeLm 26, l26
	NFadeL 26, l26a
      NFadeLm 27, l27
	NFadeL 27, l27a
      NFadeLm 28, l28
	NFadeL 28, l28a

     NFadeL 31, l31
     NFadeL 32, l32
     NFadeLm 33, l33
	NFadeL 33, l33a
     NFadeL 34, l34
     NFadeLm 35, l35
     NFadeL 35, l135
     NFadeL 36, l36
     NFadeLm 37, l37
	NFadeL 37, l37a
     NFadeLm 38, l38
	NFadeL 38, l38a

     NFadeLm 41, l41
	NFadeL 41, l41a
     NFadeLm 42, l42
	NFadeL 42, l42a
     NFadeLm 43, l43
	NFadeL 43, l43a
     NFadeLm 44, l44
	NFadeL 44, l44a
     NFadeLm 45, l45
	NFadeL 45, l45a
     NFadeL 46, l46
     NFadeL 47, l47
     NFadeLm 48, l48
	NFadeL 48, l48a

     NFadeL 51, l51
     NFadeLm 52, l52
	NFadeLm 52, l52a
     NFadeLm 53, l53
	NFadeLm 53, l53a
     NFadeL 54, l54
     NFadeL 55, l55
     NFadeL 56, l56
     NFadeL 57, l57
     NFadeLm 58, l58
	 NFadeL 58, l58a

     NFadeL 61, l61
     NFadeL 62, l62
     NFadeLm 63, l63
     NFadeL 63, l163
     NFadeLm 64, l64
     NFadeL 64, l164
     NFadeLm 65, l65
     NFadeL 65, l165
	 NFadeLm 66, l66a
     NFadeObjm 66, l66, "bulbcover1_yellowOn", "bulbcover1_yellow"   'Ramp Entrence Yellow  LED
     Flashm 66, f66
	 Flash 66, f66b
	 NFadeLm 67, l67a
     NFadeObjm 67, l67, "bulbcover1_redOn", "bulbcover1_red"       'Ramp Entrence Red    LED
     Flashm 67, f67
	 Flash 67, f67b

     NFadeLm 71, l71
	NFadeL 71, l71a
     NFadeLm 72, l72
	NFadeL 72, l72a

	 NFadeObjm 73, dome1, "light3on", "light3off"
	 Flash 73, f73 	'Traffic Light Primitive Green

	 NFadeObjm 74, dome2, "light2on", "light2off"
	 Flash 74, f74 'Traffic Light Primitive Yellow

	 NFadeObjm 75, dome3, "light1on", "light1off"
	 Flash 75, f75 'Traffic Light Primitive Red

     NFadeLm 76, l76
	NFadeL 76, l76a
     NFadeLm 77, l77
	NFadeL 77, l77a
     NFadeLm 78, l78
	NFadeL 78, l78a

	 NFadeLm 81, l81
	NFadeL 81, l81a
	 NFadeLm 82, l82
	NFadeLm 82, l82a
	 NFadeL 83, l83
	 NFadeL 84, l84
 	 NFadeLm 85, l85
	NFadeL 85, l85a
 	 NFadeLm 86, l86
	NFadeL 86, l86a
 	 NFadeLm 87, l87
	NFadeL 87, l87a
 	 NFadeLm 88, l88
	NFadeL 88, l88a

'Solenoid Controlled Lights and Flashers

 	 NFadeLm 117, f117
	 Flash 117, f117b

	 NFadeLm 118, f118b
     Flash 118, f118a
	 NFadeObjm 118, dome118, "domeon", "domeoff"   'supercharger dome 1

 	 NFadeL 119, f119

 	 NFadeLm 120, f120
	 NFadeLm 120, f120a

  	 NFadeLm 121, f121
	 NFadeLm 121, f121b
     Flashm 121, f121a
	 Flash 121, f121c
	 NFadeObjm 121, dome121, "domeon", "domeoff"   'supercharger dome 2

 	 NFadeLm 122, f122
	 Flash 122, f122b

 	 NFadeLm 123, f123
	 NFadeLm 123, f123b
     Flashm 123, f123a
	 Flash 123, f123c
	 NFadeObjm 123, dome123, "domeon", "domeoff"   'supercharger dome 3

 	 NFadeL 124, f124

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

 'Reels
Sub FadeReel(nr, reel)
    Select Case FadingLevel(nr)
        Case 2:FadingLevel(nr) = 0
        Case 3:FadingLevel(nr) = 2
        Case 4:reel.Visible = 0:FadingLevel(nr) = 3
        Case 5:reel.Visible = 1:FadingLevel(nr) = 1
    End Select
End Sub

 'Inverted Reels
Sub FadeIReel(nr, reel)
    Select Case FadingLevel(nr)
        Case 2:FadingLevel(nr) = 0
        Case 3:FadingLevel(nr) = 2
        Case 4:reel.Visible = 1:FadingLevel(nr) = 3
        Case 5:reel.Visible = 0:FadingLevel(nr) = 1
    End Select
End Sub

'**********************************************************************************************************
'**********************************************************************************************************
'	Start of VPX functions
'**********************************************************************************************************
'**********************************************************************************************************

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
	vpmTimer.pulseSw 32
    PlaySound SoundFX("right_slingshot",DOFContactors), 0, 1, 0.05, 0.05
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

Sub LeftSlingShot_Slingshot
	vpmTimer.pulseSw 31
    PlaySound SoundFX("left_slingshot",DOFContactors),0,1,-0.05,0.05
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


Sub Pins_Hit (idx)
	PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
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

Sub Spinner_Spin
	PlaySound "fx_spinner",0,.25,0,0.25
End Sub

Sub Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then
		PlaySound "rubber1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
		bounce.enabled=1
 	End If
End Sub

Sub Posts_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 16 then
		PlaySound "rubber1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 16 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "rubber1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "post4", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub

Sub LeftFlipper_Collide(parm)
 	RandomSoundFlipper()
End Sub

Sub RightFlipper_Collide(parm)
 	RandomSoundFlipper()
End Sub

Sub RightFlipper1_Collide(parm)
 	RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub

Sub bounce_timer()
		PlaySound "fx_bounce"
		bounce.enabled=0
End Sub

Sub GraphicsTimer_Timer()
		batleft.objrotz = LeftFlipper.CurrentAngle + 1
		batleftshadow.objrotz = batleft.objrotz
		batright.objrotz = RightFlipper.CurrentAngle - 1
		batrightshadow.objrotz  = batright.objrotz
		batright1.objrotz = RightFlipper1.CurrentAngle - 1
		batrightshadow1.objrotz  = batright1.objrotz
End Sub

'	Ball shadows

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
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) + 5
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) - 5
        End If
        ballShadow(b).Y = BOT(b).Y + 10
        If BOT(b).Z > 10 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
End Sub

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

Const tnob = 5 ' total number of balls
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

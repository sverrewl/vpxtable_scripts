'........'''''.      ..'',,,,,,,,,;;;;;;;;;;;;;;;;;;;;;;;;;;,..                ..';:cc::;,..           .;;;;,                   .''''''.'
',:'........,dc    ,loc;;;,,,,,,,,,,,;;;;;;;;;;;;;;;;;;;;;;;cod1'           'codlc:,,,,;:loxxc'      ..xx;;;dd.                 cx,....,o'
',c.        .x:  ,ol'...                                      .,oo'      .:ol,.           ..':xkl.   ..kk. ..ckl.               cO.    'o'
';l.        .kc cx:..     .......................................;d,   .cd:.        ...       .'oO1. .'kk.   .'lkc.             lO.    'd'
'.;,,,,,,,,;;c;'ko..   ;ddllllllllooooooooooodddddddddddddddddddxxdo. .dc.     .,:oodddxdl;.    .;kO,.'Ok.    ..'dk:.           lO.    'd'
'              ,x: .   .kd............................................ol.     ,dl,......'ckOc.....,OO''OO.      ..;xx'          l0.    .d'
'              ,kc .   .ko..   ....,kdlclllxd'.:kdooodddddddddddddxx':x...  .lo............1Kx.....;0d'OO.        ..cko. .......o0.    ,d'
'              ,kc .   .kd..  .....:Kd.....O0,.oK;...............:kc.dc.....ck'.............1Xc.....k0,OO.          .,Ox,xc'''',':.    ,x'
'              ,k: .   .kd..    ...:0d.....O0,.oK;.............;dd,..d:.....lO..............cKc.....x0,0O.    .......;0x;0:            ,x'
'              ,k: .    kd..      .;0d.....O0,.,kd,.....':::cloc'....co.    ,ko............,OO.....,0x'0O.  ..ox:,,'';:''dx;           ,x,
'              ,x:      ko..       ;0d.....O0'...ckd;.....'oxc........o,     .ox;.........1Ox......xK;,OO.  ..oO,       ..;xd.         ,x,
'              ,x:      xo..       ;0o.   .k0'.....:xo,.....'ld;......'o'.    .'colc::cldd1'    ..xK:.'OO.  ..oO,          .cxc.       ,x'
'              ,x;      xo..       ;Oo.   .k0'.     .;dd;.    .:o;.    .lc.       .......      .;Ox' .'OO.  ..lO,            .cd;      ,x'
'              ,x;      xl .       ,Oo.   .k0'.       .,od;,    .:dc.    'cc'.              .'lkk;   .'Ok.  ..lO,              .ld;    ,d'
'              ,d;      xl .       ,kl.   .xO..          'dd;.    .:o:.    .;ll;'........':okxc.     .'kk.  ..lk,               .'oo.  ,d,
'              .c;;''',;,          .oo::cccl:.             'ldigitalllc.      ..,;:ccllllc:'.         .:c:::
'
'
'
'***** TRON LEGACY LE *****

'***** PinUp Player mod code has been removed, so this table can work with the May 2018 Tron Legacy PuP-Packs *****

'***** SSF V1.1 *****

'We thank the previous friends of TRON
'(Apologies if we have missed anyone)

' ICPjuggla, freneticamnesic: Original VPX Table (V1.3f)

' 85Vette: Original VP9 table

' Rom: Original FP table

' TerryRed: PinUp Player original mod removed (to work with Pinup proper May 2018 onward) Table is a standard VPX table now
'			Ball Controller Mod added.

' Dozer: Fixed recognizer and disc turntable movement (not all light mods moved to this version)

' RustyCardores: Surround sound mod, new sounds added (where there were none)

' Hannibal: Lighting and graphical improvements

' Draifet: Physics and graphical improvements
 
' HauntFreaks: Graphical and material improvements

'****** THIS VERSION (2.2) 4K VPX 10.5+ ****

' DJRobX: Updated physics and code to bring table inline with VPX 10.4 routines. 
'		  ROM-controlled GI and PWM flasher support.   Merging changes between
'         existing tables. Fastflips hardcoded.. 

' G5k:	Playfield, plastics, ramps and other graphical improvments, new arcade primitive, modified ramp primitives.
'		Lighting, material and physics adjustments and general trial and error adjustments.
'
' Hauntfreaks: Adjusted POV


Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

' ***************************
' User options
' ***************************

Const LensFlares = True ' Hannibal lens flares.  Very large textures, may cause lag 
Const GlowRamps = 10   ' Boost neon tube glow (0 = off, 10 = default 100 = very bright)
Const Sidewall = True    ' Dozer's sidewall reflections of neon tubes (new artwork created)
Const SpecialApron = True  ' False = Stern Cards, True = Lovely Tron lady cards
Const GIBleedOpacity = 25 ' Blue cast over rear of table.   100= max, 0 = off
Const WhiteGI =  True ' GI LED Color: Blue = 0, Cool White = 1 



Const UseVPMModSol = True
Dim DesktopMode: DesktopMode = Table.ShowDT
Dim UseVPMDMD: UseVPMDMD = False
If DesktopMode = True Then 'Show Desktop components
	Ramp16.visible=1
	Ramp15.visible=1
	UseVPMDMD = True   
'Reflect2.X = 920
'Reflect2.Y = 875.8458
'Reflect2.Height = 145
'Reflect1.X = 8
'Reflect1.Y = 890.5281
'Reflect1.Height = 180
'Reflect3.Height = 200
Else
	Ramp16.visible=0
	Ramp15.visible=0
End if

gibleed.opacity = GIBleedOpacity
if GlowRamps >0 Then
	for each bulb in RampGlow1
		bulb.IntensityScale=GlowRamps / 10
	Next
	for each bulb in RampGlow2
		bulb.IntensityScale=GlowRamps / 10 
	Next
end if 
if WhiteGI Then
	dim bulb
	for each bulb in GI
	bulb.color = RGB(255,255,255)
	bulb.colorfull = RGB(255,255,255)
'	bulb.IntensityScale=.8
	bulb.TransmissionScale=.3
	next
	' All blue flippers
	LFLogo.Image="tron_flipper_blue"
	LFLogoUp.Image="tron_flipper_blue"
	GI_9.IntensityScale = .5
	GI_9.BulbModulateVsAdd = 1
	GI_10.IntensityScale = .5
	GI_16.IntensityScale = .5
	
end if 

LoadVPM "01560000", "sam.VBS", 3.10

'********************
'Standard definitions
'********************

Const cGameName = "trn_174h"

Const UseSolenoids = 1
Const UseLamps = 0
Const UseSync = 1
Const HandleMech = 0 
Const UseGI=1

'Standard Sounds
Const SSolenoidOn = "Solenoid"
Const SSolenoidOff = ""
Const SCoin = "CoinIn"

 '************
' Table init.
'************

Dim xx
Dim Bump1, Bump2, Bump3, Mech3bank,bsTrough,bsRHole,DTBank4,turntable,ttDisc1
Dim PlungerIM
Dim lighthanmovepos(2)
Dim wechsel, Drehmerker, Bildschirmaktiv, Arcadetimer1, Arcadetimer2
DIM Discdir
Discdir = 40


Sub Table_Init
	vpmInit Me
	UpPost.Isdropped=true
	With Controller
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
		.SplashInfoLine = "Tron Legacy LE 4k Edition (Stern 2011)"
		.HandleKeyboard = 0
		.ShowTitle = 0
		.ShowDMDOnly = 1
		.ShowFrame = 0
		.HandleMechanics = 1
		.Hidden = UseVPMDMD  
        .Games(cGameName).Settings.Value("sound") = 1
		On Error Resume Next
		.Run GetPlayerHWnd
		If Err Then MsgBox Err.Description
	End With

    On Error Goto 0


       '**Trough
    Set bsTrough = New cvpmBallStack
    bsTrough.InitSw 0, 21, 20, 19, 18, 0, 0, 0
    bsTrough.InitKick BallRelease, 90, 8
    bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
    bsTrough.Balls = 4

	'***Right Hole bsRHole
     Set bsRHole = New cvpmBallStack
     With bsRHole
         .InitSw 0, 11, 0, 0, 0, 0, 0, 0
         .InitKick sw11, 198, 30
         .KickZ = 0.4
         .InitExitSnd SoundFX("popper_ball",DOFContactors), SoundFX("Solenoid",DOFContactors)
         .KickForceVar = 2
     End With


 	'DropTargets
   	Set DTBank4 = New cvpmDropTarget  
   	  With DTBank4
   		.InitDrop Array(sw04,sw03,sw02,sw01),Array(4,3,2,1)
        .Initsnd SoundFX("fx_droptarget",DOFContactors), SoundFX("fx_droptargetup",DOFContactors)
       End With

      '**Main Timer init
	PinMAMETimer.Interval = PinMAMEInterval
	PinMAMETimer.Enabled = 1
 
	'Nudging
    	vpmNudge.TiltSwitch=-7
    	vpmNudge.Sensitivity=3    	
		vpmNudge.TiltObj=Array(Bumper1b,Bumper2b,Bumper3b,LeftSlingshot,RightSlingshot)

     ' Impulse Plunger
    Const IMPowerSetting = 150
    Const IMTime = 0.7
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swplunger, IMPowerSetting, IMTime
        .Random 0.3
        .InitExitSnd "plunger2", "plunger"
        .CreateEvents "plungerIM"
    End With

	TBPos=28:TBTimer.Enabled=0:TBDown=1:Controller.Switch(52) = 1:Controller.Switch(53) = 0

'   Spinning Disk

	Set ttDisc1 = New myTurnTable
		ttDisc1.InitTurnTable Disc1Trigger, 8
		ttDisc1.SpinCW = False
		ttDisc1.CreateEvents "ttDisc1"

	'vpmMapLights Collection1

	if SpecialApron = true then Apron.Image = "apron-tronspecial"

' Fast flips

	vpmFlipsSam.FlipperSolNumber(0) = 15
	vpmFlipsSam.FlipperSolNumber(1) = 16
	InitVpmFlipsSAM
	vpmflipssam.tiltsol True	'make sure this is on for SS flippers
End Sub

Sub Table_Paused:Controller.Pause = 1:End Sub
Sub Table_unPaused:Controller.Pause = 0:End Sub

 
'*****Keys
Sub Table_KeyDown(ByVal keycode)
	' Fast flips - once integrated into sam.vbs this won't be necessary
 	If Keycode = LeftFlipperKey then 
		vpmFlipsSam.FlipL true
	End If
 	If Keycode = RightFlipperKey then 
		vpmFlipsSam.FlipR true
	End If
	' end fast flips
    If keycode = PlungerKey Then PlaySoundAt "PlungerPull", Plunger:Plunger.Pullback
	If Keycode = LeftTiltKey Then Nudge 90, 8: PlaySound "fx_nudge_left"
	If Keycode = RightTiltKey Then Nudge 270, 8: PlaySound "fx_nudge_right"
	If Keycode = CenterTiltKey Then Nudge 0, 8: PlaySound "fx_nudge_forward"
    If vpmKeyDown(keycode) Then Exit Sub 
End Sub

Sub Table_KeyUp(ByVal keycode)
	' Fast flips - once integrated into sam.vbs this won't be necessary
 	If Keycode = LeftFlipperKey then 
		vpmFlipsSam.FlipL false
	End If
 	If Keycode = RightFlipperKey then 
		vpmFlipsSam.FlipR false
	End If
	' end fast flips
	If vpmKeyUp(keycode) Then Exit Sub
	If Keycode = StartGameKey Then Controller.Switch(16) = 0
    If keycode = PlungerKey Then PlaySoundAt "plunger", Plunger:Plunger.Fire
End Sub

   'Solenoids
SolCallback(1) = "solTrough"
SolCallback(2) = "solAutofire"
SolCallback(3) = "DTBank4.SolDropUp"
SolCallback(4) = "bsRHole.SolOut"
SolCallback(5)="SolDiscMotor"' spinning disk
SolCallback(6) = "TBMove"
SolCallback(7) = "orbitpost"
'SolCallback(8) = "shaker"
SolModCallback(9) = "SetLampMod 139,"
SolModCallback(10) = "SetLampMod 140,"
SolModCallback(11) = "SetLampMod 141,"
'SolCallback(12) = "upperleftflipper"
'SolCallback(13) = "leftslingshot"
'SolCallback(14) = "rightslingshot"
SolCallback(15) = "SolLFlipper"
SolCallback(16) = "SolRFlipper"

'Flashers
SolModCallback(17) = "SetLampMod 117,"    'flash zen
SolModCallback(18) = "SetLampMod 118,"'flash videogame
SolModCallback(25) = "setlampMod 119,"    'flash right domes x2
SolModCallback(20) = "setLampMod 120,"  'LE apron left
SolModCallback(21) = "setlampmod 121,"   'LE apron right
'SolCallback(22) = "discdirrelay"  'LE disc direction relay
SolCallback(23) = "recogrelay"    'LE recognizer

SolModCallback(19) = "setlampmod 125,"'flash left domes
SolModCallback(26) = "SetLampMod 126,"'flash disc left 
SolModCallback(27) = "SetLampMod 127,"'flash disc right
SolModCallback(28) = "SetLampMod 128,"'flash backpanel x2
SolModCallback(29) = "SetLampMod 129,"'flash recognizer
SolModCallback(30) = "SetLampMod 130,"'disc motor relay
SolModCallback(31) = "SetLampMod 131,"'flash red disc left x2
SolModCallback(32) = "SetLampMod 132,"'LE flash red disc x2


Dim XLocation,XDir,T(4),ZRot
Dim recogdir
XDir=1
XLocation=-30
ZRot=1


Sub RecognizerTimer_Timer
	If recognizer.rotz <= -18 Then
		Controller.Switch(56) = 1
	Else
		Controller.Switch(56) = 0
	End If
	If recognizer.rotz => 18 Then
		Controller.Switch(54) = 1
	Else
		Controller.Switch(54) = 0
	End If
	If recognizer.rotz <=0 AND recognizer.rotz > -1 OR recognizer.rotz => 0 AND recognizer.rotz <= 1 Then
		Controller.Switch(55) = 1
	Else
		Controller.Switch(55) = 0
	End If 

	Select Case recogdir
	Case 1:
		If recognizer.rotz <= -20 Then
			recogdir = 2
		End If
		recognizer.rotz = recognizer.rotz - 0.1
	Case 2:
		If recognizer.rotz => 20 Then
			recogdir = 1
		End If
		recognizer.rotz = recognizer.rotz + 0.1
	End Select
End Sub


Sub recogrelay(Enabled)
	If Enabled Then 
		recogdir = 1:RecognizerTimer.enabled=1
	Else
		RecognizerTimer.enabled=0
	End If
End Sub

Sub solTrough(Enabled)
	If Enabled Then
		bsTrough.ExitSol_On
		vpmTimer.PulseSw 22
	End If
 End Sub



 '******
 'Auto Plunger add by Hanibal
 '******

Dim AP
Dim Zeit

Sub SolAutofire(Enabled)
	If Enabled Then
		AP = True
'		PlungerIM.AutoFire
      PlaySoundAt "plunger", Plunger
	End If
End Sub

Sub PlungerPTimer_Timer()
	if AP = True and Zeit < 10 then Plunger.visible = 0 :Plunger1.visible = 1 : PlungerPTimer1.Enabled = 1	 :Zeit = Zeit +1 ':Test1.state = 1
	if AP = False and Zeit > 0 then Plunger1.Fire : Zeit = 0 ':Test1.state = 0
	if Zeit >= 10 then AP = False : 	PlaySoundAt "ShooterLane", Plunger
End Sub


Sub PlungerPTimer1_Timer()
	Plunger.visible = 1 :Plunger1.visible = 0
	PlungerPTimer1.Enabled = 0
End Sub


Sub Sol3bankmotor(Enabled)
	 	If Enabled then
 		RiseBank
		DropBank
	end if
End Sub


Sub orbitpost(Enabled)
	If Enabled Then
		UpPost.Isdropped=false
	Else
		UpPost.Isdropped=true
	End If
 End Sub



'Switches

Sub sw01_Hit:DTBank4.Hit 4:End Sub
Sub sw02_Hit:DTBank4.Hit 3:End Sub
Sub sw03_Hit:DTBank4.Hit 2:End Sub
Sub sw04_Hit:DTBank4.Hit 1:End Sub
Sub sw7_Hit
	Me.TimerEnabled = 1
	sw7p.TransX = -2
	vpmTimer.PulseSw 7
	PlaySoundAt SoundFX("fx_target",DOFContactors),sw7p
End Sub

Sub sw7_Timer:Me.TimerEnabled = 0:sw7p.TransX = 0:End Sub
Sub sw8_Hit
	Me.TimerEnabled = 1
	sw8p.TransX = -2
	vpmTimer.PulseSw 8
	PlaySoundat SoundFX("fx_target",DOFContactors),sw8p
End Sub
Sub sw8_Timer:Me.TimerEnabled = 0:sw8p.TransX = 0:End Sub
Sub sw12_Hit:Controller.Switch(12) = 1:PlaySoundAt "rollover",sw12:End Sub
Sub sw12_UnHit:Controller.Switch(12) = 0:End Sub
Sub sw13_Hit
	Me.TimerEnabled = 1
	sw13p.TransX = -2
	vpmTimer.PulseSw 13
	PlaySoundAt SoundFX("fx_target",DOFContactors),sw13p
End Sub
Sub sw13_Timer:Me.TimerEnabled = 0:sw13p.TransX = 0:End Sub
Sub sw14_Hit
	Controller.Switch(14) = 1
	PlaySoundAt "rollover",sw14
End Sub
Sub sw14_UnHit:Controller.Switch(14) = 0:End Sub
'Sub sw23:End Sub
Sub sw24_Hit
	Controller.Switch(24) = 1
	PlaySoundAt "rollover",sw24
End Sub
Sub sw24_UnHit:Controller.Switch(24) = 0:End Sub
Sub sw25_Hit
	Controller.Switch(25) = 1
	PlaySoundAt "rollover",sw25
End Sub
Sub sw25_UnHit:Controller.Switch(25) = 0:End Sub
Sub sw28_Hit
	Controller.Switch(28) = 1
	PlaySoundAt "rollover",sw28
End Sub
Sub sw28_UnHit:Controller.Switch(28) = 0:End Sub
Sub sw29_Hit
	Controller.Switch(29) = 1
	PlaySoundAt "rollover",sw29
End Sub
Sub sw29_UnHit:Controller.Switch(29) = 0:End Sub
Sub sw34_Hit
	Controller.Switch(34) = 1
	PlaySoundAt "Gate",Sw34
	'LeftCount = LeftCount + 1
End Sub
Sub sw34_UnHit:Controller.Switch(34) = 0:End Sub
Sub sw35_Hit:Controller.Switch(35) = 1:PlaySoundAt "Gate",sw35:End Sub
Sub sw35_UnHit:Controller.Switch(35) = 0:End Sub
Sub sw36_Spin:vpmTimer.PulseSw 36::PlaySoundAtVol"fx_spinner",sw36,.2:End Sub
Sub sw37_Hit 
	Controller.Switch(37) = 1
	PlaySoundAt "Gate",sw37
	'RightCount = RightCount + 1
End Sub
Sub sw37_UnHit:Controller.Switch(37) = 0:End Sub
Sub sw38_Hit:Controller.Switch(38) = 1:PlaySoundAt "Gate",sw38:End Sub
Sub sw38_UnHit:Controller.Switch(38) = 0:End Sub
Sub sw39_Hit
	Controller.Switch(39) = 1
	PlaySoundAt "rollover",sw39
End Sub
Sub sw39_UnHit:Controller.Switch(39) = 0:End Sub
Sub sw41_Hit
	Controller.Switch(41) = 1
End Sub
Sub sw41_UnHit:Controller.Switch(41) = 0:End Sub
Sub sw43_Hit
	Controller.Switch(43) = 1
	PlaySoundAt "rollover",sw41
End Sub
Sub sw43_UnHit:Controller.Switch(43) = 0:End Sub
Sub sw44_Spin
	vpmTimer.PulseSw 44
	PlaySoundAtVol"fx_spinner",l49,.2
End Sub
Sub sw46_Hit
	Controller.Switch(46) = 1
	PlaySoundAt "rollover",sw46
End Sub
Sub sw46_UnHit:Controller.Switch(46) = 0:End Sub
Sub sw48_Hit
	Me.TimerEnabled = 1
	sw48p.TransX = -2
	vpmTimer.PulseSw 48
	PlaySoundAt SoundFX("fx_target",DOFContactors),sw48p
End Sub
Sub sw48_Timer:Me.TimerEnabled = 0:sw48p.TransX = 0:End Sub
Sub Laneexit_Hit: Ramphelfer.Enabled= True: End Sub


Sub Ramphelfer_Timer
Playsound "DROP_RIGHT"
Ramphelfer.Enabled=False
End Sub


'Arcade Scoop
 Dim aBall, aZpos
 Dim bBall, bZpos

 Sub sw11_Hit
     Set bBall = ActiveBall
     PlaySoundAt "VUKEnter", sw11
     bZpos = 45
     Me.TimerInterval = 2
     Me.TimerEnabled = 1
 End Sub
 
 Sub sw11_Timer
     bBall.Z = bZpos
     bZpos = bZpos-4
     If bZpos <-30 Then
         Me.TimerEnabled = 0
         Me.DestroyBall
         bsRHole.AddBall Me
     End If
 End Sub
 
' ===============================================================================================
' spinning discs (New) Taken from Whirlwind written by Herweh
' ===============================================================================================

Dim discAngle, stepAngle, stopDiscs, discsAreRunning

Sub discdirrelay(Enabled)
	If Enabled Then
		Discdir = -40.0
	Else
		Discdir = 40.0
	End If
End Sub



InitDiscs()

Sub InitDiscs()
	discAngle 			= 0
	discsAreRunning		= False
End Sub

Sub SolDiscMotor(Enabled)
    ttDisc1.MotorOn = Enabled
	If Enabled Then
		stepAngle			= Discdir '40.0
		discsAreRunning		= True
		stopDiscs			= False
		DiscsTimer.Interval = 10
		DiscsTimer.Enabled 	= True
		Drehtimer.Enabled  = True
        playsound "spindisc", -1, 0.02, 0, 0, 5, 1, 0
	Else
		stopDiscs			= True
		discsAreRunning		= True
		Drehtimer.Enabled = False
		Drehlicht1.state = 0: Drehlicht2.state = 0 : Drehlicht3.state = 0
        stopsound "spindisc"
	End If
End Sub

Sub DiscsTimer_Timer()
	' calc angle
	discAngle = discAngle + stepAngle
	If discAngle >= 360 Then
		discAngle = discAngle - 360
	End If
	If discAngle <= 0 Then
		discAngle = discAngle + 360
	End If

	' rotate discs
	Disc1.RotAndTra2 = 360 - discAngle

	If stopDiscs AND Discdir > 0 Then
		stepAngle = stepAngle -0.1
		If stepAngle <= 0 Then
			DiscsTimer.Enabled 	= False
		End If
	End If

	If stopDiscs AND Discdir < 0 Then
		stepAngle = stepAngle + 0.1
		If stepAngle >= 0 Then
			DiscsTimer.Enabled 	= False
		End If
	End If
End Sub


Class myTurnTable
	Private mX, mY, mSize, mMotorOn, mDir, mBalls, mTrigger
	Public MaxSpeed, SpinDown, Speed

	Private Sub Class_Initialize
		mMotorOn = False : Speed = 0 : mDir = 1 : SpinDown = 15
		Set mBalls = New cvpmDictionary
	End Sub

	Public Sub InitTurntable(aTrigger, aMaxSpeed)
		mX = aTrigger.X : mY = aTrigger.Y : mSize = aTrigger.Radius
		MaxSpeed = aMaxSpeed : Set mTrigger = aTrigger
	End Sub

	Public Sub CreateEvents(aName)
		If vpmCheckEvent(aName, Me) Then
			vpmBuildEvent mTrigger, "Hit", aName & ".AddBall ActiveBall"
			vpmBuildEvent mTrigger, "UnHit", aName & ".RemoveBall ActiveBall"
			vpmBuildEvent mTrigger, "Timer", aName & ".Update"
		End If
	End Sub

	Public Sub SolMotorState(aCW, aEnabled)
		mMotorOn = aEnabled
		If aEnabled Then If aCW Then mDir = 1 Else mDir = -1
		NeedUpdate = True
	End Sub
	Public Property Let MotorOn(aEnabled)
		mMotorOn = aEnabled
		NeedUpdate = (mBalls.Count > 0) Or (SpinDown > 0)
	End Property
	Public Property Get MotorOn
		MotorOn = mMotorOn
	End Property

	Public Sub AddBall(aBall)
		On Error Resume Next
		mBalls.Add aBall,0
		NeedUpdate = True
	End Sub
	Public Sub RemoveBall(aBall)
		On Error Resume Next
		mBalls.Remove aBall
		NeedUpdate = (mBalls.Count > 0) Or (SpinDown > 0)
	End Sub

	Public Property Let SpinCW(aCW)
		If aCW Then mDir = 1 Else mDir = -1
		NeedUpdate = True
	End Property
	Public Property Get SpinCW
		SpinCW = (mDir = 1)
	End Property

	Public Sub Update
		If mMotorOn Then
			Speed = MaxSpeed
			NeedUpdate = mBalls.Count
		Else
			Speed = Speed - SpinDown*MaxSpeed/3000 '100
			If Speed < 0 Then 
				Speed = 0
				'msgbox "off"
				NeedUpdate = mBalls.Count
			End If
		End If
		If Speed > 0 Then
			Dim obj
			On Error Resume Next
			For Each obj In mBalls.Keys
				If obj.X < 0 Or Err Then RemoveBall obj Else AffectBall obj
			Next
			On Error Goto 0
		End If
	End Sub

	Public Sub AffectBall(aBall)
		Dim dX, dY, dist
		dX = aBall.X - mX : dY = aBall.Y - mY : dist = Sqr(dX*dX + dY*dY)
		If dist > mSize Or dist < 1 Or Speed = 0 Then Exit Sub
		aBall.VelX = aBall.VelX - (dY * mDir * Speed / 1000)
		aBall.VelY = aBall.VelY + (dX * mDir * Speed / 1000)
	End Sub

	Private Property Let NeedUpdate(aEnabled)
		If mTrigger.TimerEnabled <> aEnabled Then
			mTrigger.TimerInterval = 10
			mTrigger.TimerEnabled = aEnabled
		End If
	End Property
End Class

'*****************************************************************************************
'*******************   Arcade Bildwechsel             ************************************
'*****************************************************************************************


Sub Arcadetimer2_Timer()
	IF Bildschirmaktiv = True Then
		Select Case Int(Rnd*9)+1
			Case 1 : Monitor.image = "Arcadeframe1"
			Case 2 : Monitor.image = "Arcadeframe2"
			Case 3 : Monitor.image = "Arcadeframe3"
			Case 4 : Monitor.image = "Arcadeframe4" 
			Case 5 : Monitor.image = "Arcadeframe5"
			Case 6 : Monitor.image = "Arcadeframe6"
			Case 7 : Monitor.image = "Arcadeframe7"
			Case 8 : Monitor.image = "Arcadeframe8"
			Case 9 : Monitor.image = "Arcadeframe9"
		End Select
	End If
End Sub


Sub Arcadetimer1_Timer()
	IF Bildschirmaktiv = True Then
		Monitorflash.opacity = (RND * 200)
		Monitorlicht.intensity = 10 + (RND * 5)
	Else
		Monitorflash.opacity = 0
	End If
End Sub

'*****************************************************************************************
'*******************   Drehtimer  (Spinning disc lights)            ************************************
'*****************************************************************************************

Sub Drehtimer_Timer
	Drehmerker = Drehmerker +1
	Select Case Drehmerker
		Case 1 : Drehlicht1.state = 1: Drehlicht2.state = 0 : Drehlicht3.state = 0
		Case 2 : Drehlicht1.state = 0: Drehlicht2.state = 1 : Drehlicht3.state = 0
		Case 3 : Drehlicht1.state = 0: Drehlicht2.state = 0 : Drehlicht3.state = 1 : Drehmerker = 0
	End Select
End Sub

  
Sub SolLFlipper(Enabled)
     If Enabled Then
		 PlaySoundAt SoundFX("FlipperUpLeft",DOFContactors),LeftFlipper
		 LeftFlipper.RotateToEnd:LeftFlipper1.RotateToEnd
     Else
		 PlaySoundAt SoundFX("FlipperDown",DOFContactors),LeftFlipper
		 LeftFlipper.RotateToStart:LeftFlipper1.RotateToStart
     End If
 End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
		 PlaySoundAt SoundFX("FlipperUpRight",DOFContactors),RightFlipper
		 RightFlipper.RotateToEnd
     Else
		 PlaySoundAt SoundFX("FlipperDown",DOFContactors),RightFlipper
		 RightFlipper.RotateToStart
     End If
 End Sub   


 'Drains and Kickers
Dim BallCount:BallCount = 0

Sub Drain_Hit()
	PlaySound "Drain"
	BallCount = BallCount - 1
	bsTrough.AddBall Me
	If BallCount = 0 then Bildschirmaktiv = FALSE : Monitor.image = "Arcadeframe0" 
End Sub

Sub BallRelease_UnHit()
	BallCount = BallCount + 1
	Bildschirmaktiv = True
End Sub



'***Slings and rubbers

 Dim LStep, RStep

Sub LeftSlingShot_Slingshot
	PlaySoundAt SoundFX("left_slingshot",DOFContactors),SLING2
	vpmTimer.PulseSw 26
	LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
    LeftSlingShot.TimerInterval = 10
    Sling_linkslicht.state = True
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0: Sling_linkslicht.state = False
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
	PlaySoundAt SoundFX("right_slingshot",DOFContactors),SLING1
	vpmTimer.PulseSw 27
	RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
    RightSlingShot.TimerInterval = 10
	Sling_rechtslicht.state = True
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0: Sling_rechtslicht.state = False
    End Select
    RStep = RStep + 1
End Sub

   'Bumpers
Sub Bumper1b_Hit
	vpmTimer.PulseSw 31
	PlaySoundAtBumperVol SoundFX("fx_bumper1",DOFContactors),Bumper1b,1
End Sub
     
 
Sub Bumper2b_Hit
	vpmTimer.PulseSw 30
	PlaySoundAtBumperVol SoundFX("fx_bumper1",DOFContactors),Bumper2b,1
End Sub

Sub Bumper3b_Hit
	vpmTimer.PulseSw 32
	PlaySoundAtBumperVol SoundFX("fx_bumper1",DOFContactors),Bumper3b,1
End Sub
  

Dim LampState(200), FadingLevel(200)
Dim FlashSpeedUp, FlashSpeedDown
Dim x

InitLamps()
'

'' Lamp & Flasher Timers


Sub SetLampMod(nr, value)
    If value > 0 Then
		LampState(nr) = 1
		FadingLevel(nr) = value  
	Else
		LampState(nr) = 0
	End If
End Sub

Sub LampMod(nr, object)
	dim nstate
	if FadingLevel(nr) > 0 then 
		nstate = 1
	Else	
		nstate = 0
	end if
	If TypeName(object) = "Light" Then
		Object.IntensityScale = FadingLevel(nr)/128		
		Object.State = nstate
	End If
	If TypeName(object) = "Flasher" Then
		Object.IntensityScale = FadingLevel(nr)/128
		Object.visible = nstate
	End If
	If TypeName(object) = "Primitive" Then
		Object.DisableLighting = nstate
	End If
End Sub

Sub LampTimer_Timer()
    Dim chgLamp, num, chg, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp) Then
        For ii = 0 To UBound(chgLamp)
            LampState(chgLamp(ii, 0) ) = chgLamp(ii, 1)
        Next
    End If

    UpdateLamps
	BallShadowUpdate
	UpdateFlipperLogo
End Sub

'dim testx:testx = 4
'dim testy:testy = 120
'dim testl:testl = 128

Sub UpdateLamps()

' ** TEST** Temporary performance testning code 
'testx = testx -1
'if testl= 116 then
'	if testx > 4 then 
'		gion 
'	else 
'		gioff
'	end if 
'else
'if testx > 15 then setlampmod testl, 128 else setlampmod testl,0
'end if 
'if testx = 0 then testx = 30
'TESTy = TESTy -1
'if TESTy = 0 then TESTl=TESTl+1:debug.print TESTl:TESTy= 120:if TESTl=142 then TESTl=116

'' ** END TEST 

MaterialColor "Linkerstring", RGB(LampState(103), LampState(102), LampState(101))
MaterialColor "Rechterstring", RGB(LampState(106), LampState(105), LampState(104))






'****************** Hanibal Glowing Lights **********************************

if GlowRamps > 0 then

SetRGBLamp Rampenlicht1, Lampstate(106),Lampstate(105),Lampstate(104)
SetRGBLamp Rampenlicht1a1, Lampstate(106),Lampstate(105),Lampstate(104)
SetRGBLamp Rampenlicht1a2, Lampstate(106),Lampstate(105),Lampstate(104)
SetRGBLamp Rampenlicht1a3, Lampstate(106),Lampstate(105),Lampstate(104)
SetRGBLamp Rampenlicht1a4, Lampstate(106),Lampstate(105),Lampstate(104)
SetRGBLamp Rampenlicht1a5, Lampstate(106),Lampstate(105),Lampstate(104)
SetRGBLamp Rampenlicht1a6, Lampstate(106),Lampstate(105),Lampstate(104)
SetRGBLamp Rampenlicht1a7, Lampstate(106),Lampstate(105),Lampstate(104)
SetRGBLamp Rampenlicht1a8, Lampstate(106),Lampstate(105),Lampstate(104)
SetRGBLamp Rampenlicht1a9, Lampstate(106),Lampstate(105),Lampstate(104)
SetRGBLamp Rampenlicht1a10, Lampstate(106),Lampstate(105),Lampstate(104)
SetRGBLamp Rampenlicht1a11, Lampstate(106),Lampstate(105),Lampstate(104)
SetRGBLamp Rampenlicht1a12, Lampstate(106),Lampstate(105),Lampstate(104)
SetRGBLamp Rampenlicht1a13, Lampstate(106),Lampstate(105),Lampstate(104)
SetRGBLamp Rampenlicht1a14, Lampstate(106),Lampstate(105),Lampstate(104)
SetRGBLamp Rampenlicht1a15, Lampstate(106),Lampstate(105),Lampstate(104)
SetRGBLamp Rampenlicht1a16, Lampstate(106),Lampstate(105),Lampstate(104)
SetRGBLamp Rampenlicht1a17, Lampstate(106),Lampstate(105),Lampstate(104)
SetRGBLamp Rampenlicht1a18, Lampstate(106),Lampstate(105),Lampstate(104)
SetRGBLamp Rampenlicht1a19, Lampstate(106),Lampstate(105),Lampstate(104)
SetRGBLamp Rampenlicht1a20, Lampstate(106),Lampstate(105),Lampstate(104)
SetRGBLamp Rampenlicht1a21, Lampstate(106),Lampstate(105),Lampstate(104)
SetRGBLamp Rampenlicht1a22, Lampstate(106),Lampstate(105),Lampstate(104)




SetRGBLamp Rampenlicht2, Lampstate(103),Lampstate(102),Lampstate(101)
SetRGBLamp Rampenlicht2a1, Lampstate(103),Lampstate(102),Lampstate(101)
SetRGBLamp Rampenlicht2a2, Lampstate(103),Lampstate(102),Lampstate(101)
SetRGBLamp Rampenlicht2a3, Lampstate(103),Lampstate(102),Lampstate(101)
SetRGBLamp Rampenlicht2a4, Lampstate(103),Lampstate(102),Lampstate(101)
SetRGBLamp Rampenlicht2a5, Lampstate(103),Lampstate(102),Lampstate(101)
SetRGBLamp Rampenlicht2a6, Lampstate(103),Lampstate(102),Lampstate(101)
SetRGBLamp Rampenlicht2a7, Lampstate(103),Lampstate(102),Lampstate(101)
SetRGBLamp Rampenlicht2a8, Lampstate(103),Lampstate(102),Lampstate(101)
SetRGBLamp Rampenlicht2a9, Lampstate(103),Lampstate(102),Lampstate(101)
SetRGBLamp Rampenlicht2a10, Lampstate(103),Lampstate(102),Lampstate(101)
SetRGBLamp Rampenlicht2a11, Lampstate(103),Lampstate(102),Lampstate(101)
SetRGBLamp Rampenlicht2a12, Lampstate(103),Lampstate(102),Lampstate(101)
SetRGBLamp Rampenlicht2a13, Lampstate(103),Lampstate(102),Lampstate(101)
SetRGBLamp Rampenlicht2a14, Lampstate(103),Lampstate(102),Lampstate(101)
SetRGBLamp Rampenlicht2a15, Lampstate(103),Lampstate(102),Lampstate(101)
SetRGBLamp Rampenlicht2a16, Lampstate(103),Lampstate(102),Lampstate(101)
SetRGBLamp Rampenlicht2a17, Lampstate(103),Lampstate(102),Lampstate(101)
SetRGBLamp Rampenlicht2a18, Lampstate(103),Lampstate(102),Lampstate(101)
SetRGBLamp Rampenlicht2a19, Lampstate(103),Lampstate(102),Lampstate(101)
SetRGBLamp Rampenlicht2a20, Lampstate(103),Lampstate(102),Lampstate(101)
SetRGBLamp Rampenlicht2a21, Lampstate(103),Lampstate(102),Lampstate(101)
SetRGBLamp Rampenlicht2a22, Lampstate(103),Lampstate(102),Lampstate(101)
SetRGBLamp Rampenlicht2a23, Lampstate(103),Lampstate(102),Lampstate(101)
SetRGBLamp Rampenlicht2a24, Lampstate(103),Lampstate(102),Lampstate(101)
SetRGBLamp Rampenlicht2a25, Lampstate(103),Lampstate(102),Lampstate(101)
SetRGBLamp Rampenlicht2a26, Lampstate(103),Lampstate(102),Lampstate(101)
SetRGBLamp Rampenlicht2a27, Lampstate(103),Lampstate(102),Lampstate(101)
SetRGBLamp Rampenlicht2a28, Lampstate(103),Lampstate(102),Lampstate(101)

end if 

if Sidewall then 
Reflect1.Color = RGB(Lampstate(106),Lampstate(105),Lampstate(104))
Reflect2.Color = RGB(Lampstate(103),Lampstate(102),Lampstate(101))
Reflect3.Color = RGB(Lampstate(103),Lampstate(102),Lampstate(101))
end if 

NFadeLm 1, l1
NFadeLm 1, l1a
NFadeLm 2, l2
NFadeLm 2, l2a
NFadeLm 3, l3
NFadeLm 3, l3a
NFadeLm 4, l4
NFadeLm 4, l4a
NFadeLm 5, l5
NFadeLm 5, l5a
NFadeLm 6, l6
NFadeLm 6, l6a
NFadeLm 7, l7
NFadeLm 7, l7a
NFadeLm 8, l8
NFadeLm 8, l8a
NFadeLm 9, l9
NFadeLm 9, l9a
NFadeLm 10, l10
NFadeLm 10, l10a
NFadeLm 11, l11
NFadeLm 11, l11a
NFadeLm 12, l12
NFadeLm 12, l12a
NFadeLm 13, l13
NFadeLm 13, l13a
NFadeLm 14, l14
NFadeLm 14, l14a
NFadeLm 15, l15
NFadeLm 15, l15a
NFadeLm 16, l16
NFadeLm 16, l16a
NFadeLm 17, l17
NFadeLm 17, l17a
NFadeLm 18, l18
NFadeLm 18, l18a
NFadeLm 19, l19
NFadeLm 19, l19a
NFadeLm 20, l20
NFadeLm 20, l20a
NFadeLm 21, l21
NFadeLm 21, l21a
NFadeLm 22, l22
NFadeLm 22, l22a
NFadeLm 23, l23
NFadeLm 23, l23a
NFadeLm 24, l24
NFadeLm 24, l24a
NFadeLm 25, l25
NFadeLm 25, l25a
NFadeLm 26, l26
NFadeLm 26, l26a

NFadeLm 27, l27
NFadeLm 27, l27a
NFadeLm 28, l28
NFadeLm 28, l28d
NFadeLm 29, l29
NFadeLm 29, l29d
NFadeLm 30, l30
NFadeLm 30, l30a
NFadeLm 31, l31
NFadeLm 31, l31a
NFadeLm 32, l32
NFadeLm 32, l32a
NFadeLm 33, l33
NFadeLm 33, l33a
NFadeLm 34, l34
NFadeLm 34, l34a
NFadeLm 35, l35
NFadeLm 35, l35a
NFadeLm 36, l36
NFadeLm 36, l36a
NFadeLm 37, l37
NFadeLm 37, l37a
NFadeLm 38, l38
NFadeLm 38, l38a
NFadeLm 39, l39
NFadeLm 39, l39a
NFadeLm 40, l40
NFadeLm 40, l40a
NFadeLm 42, l42

NFadeLm 43, l43
NFadeLm 43, l43a
NFadeLm 45, l45
NFadeLm 45, l45a

NFadeLm 46, l46
NFadeLm 47, l47
NFadeLm 48, l48

NFadeLm 49, l49
NFadeLm 49, l49a
NFadeLm 50, l50
NFadeLm 50, l50a
NFadeLm 51, l51
NFadeLm 51, l51a

NFadeLm 52, l52
NFadeLm 52, l52a

NFadeLm 53, l53
NFadeLm 53, l53a

NFadeLm 54, l54

NFadeLm 55, l55
NFadeLm 55, l55a
NFadeLm 56, l56
NFadeLm 56, l56a
NFadeLm 57, l57
NFadeLm 57, l57a
NFadeLm 58, l58
NFadeLm 58, l58a
NFadeLm 59, l59
NFadeLm 59, l59a
NFadeLm 60, l60
NFadeLm 60, l60a
NFadeLm 61, l61
NFadeLm 61, l61a
NFadeLm 62, l62
NFadeLm 62, l62a
NFadeLm 63, l63
NFadeLm 63, l63a
NFadeLm 64, l64
NFadeLm 64, l64a


'Flashers

' Handle gradual fade outs and limit max intensity
FlashModSol 117
FlashModSol 118
FlashModSol 119
FlashModSol 120
FlashModSol 121
FlashModSol 125
FlashModSol 126
FlashModSol 127
FlashModSol 128 
FlashModSol 129
FlashModSol 131
FlashModSol 132
FlashModSol 139 
FlashModSol 140
FlashModSol 141

' Now show new states

LampMod 117, F117
LampMod 118, Flasher7
LampMod 118, Monitorlicht1
LampMod 118, Flasher7a


LampMod 119, Flasher1
LampMod 119, Flasher1a1
LampMod 119, Flasher1a2
LampMod 119, Flasher2
LampMod 119, Flasher2a1
LampMod 119, Flasher2a2

LampMod 120, Linkerflasher
LampMod 121, Lanelight 
LampMod 121, Rechterflasher

LampMod 125, Flasher5
LampMod 125, Flasher5a
LampMod 125, Flasher6
LampMod 125, Flasher6a

LampMod 126, F126

LampMod 127,F127 

'LampMod 128, Flasher3
LampMod 128, Flasher3c
LampMod 128, Flasher3d
LampMod 128, Flasher3a
'LampMod 128, Flasher4
LampMod 128, Flasher4c
LampMod 128, Flasher4d
LampMod 128, Flasher4a 

LampMod 129, f129
LampMod 129, f129a

LampMod 131, f131a
LampMod 131, f131b

LampMod 132, f132a
LampMod 132, f132b

LampMod 139, f9
LampMod 139, f9a

LampMod 140, f10
LampMod 140, f10a

LampMod 141, f11
LampMod 141, f11a

' Hannibal's lens flares

If LensFlares Then
	LampMod 117, Flasher117
	LampMod 118, Flasher7b
	LampMod 119, Flasher1b
    LampMod 125, Flasher6b
	LampMod 128, Flasher4b
    LampMod 129, Flasher129
	LampMod 131, Flasher131
	LampMod 139, f9b	
    LampMod 140, f10b
	LampMod 141, f11b
end if 


End Sub

''Lights

Sub NFadeLm(nr, a)
    a.state = LampState(nr)
End Sub

' Flasher objects
Sub FlashModSol(nr)
    Select Case LampState(nr)
        Case 0 'off
			' If lamp has been turned off harshly, gradually fade out
            FadingLevel(nr) = FadingLevel(nr) - FlashSpeedDown
            If FadingLevel(nr) < 0 Then
                FadingLevel(nr) = 0
            End if           
        Case 1 			
            '  Cap FlashLevel at 128 here.
            If FadingLevel(nr) > 128 Then
                FadingLevel(nr) = 128
            End if
    End Select
End Sub

Sub InitLamps()
	For x = 1 to 200
		LampState(x) = 0
		FadingLevel(x) = 0
    Next
    FlashSpeedUp = 10   ' fast speed when turning on the flasher
    FlashSpeedDown = 20 ' slow speed when turning off the flasher, gives a smooth fading
	UpdateLamps
	LampTimer.Interval = 16 'lamp fading speed
	LampTimer.Enabled = 1
End Sub
 
Sub SetRGBLamp(Lamp, R, G, B)   
     'dim IntensityValue 
	 'IntensityValue = .2126 * R + .7152 * G + .0722 * B
	 Lamp.Color = RGB(R, G, B)
	 Lamp.ColorFull = RGB(R, G, B)
	 'Lamp.Intensity = IntensityValue *100 / 255
	 Lamp.State = 1
End Sub

Sub SetLamp(nr, value)
    If value <> LampState(nr) Then
        LampState(nr) = abs(value)
		if abs(value) = 1 then 
			FadingLevel(nr) = 128
		end if 
    End If
End Sub

Sub SetFlash(nr, stat)
    FlashState(nr) = ABS(stat)
End Sub


'*********************************************************************
'* TARGETBANK TARGETS Taken from AFM written by Groni ****************
'*********************************************************************

Sub SW49_Hit
vpmTimer.PulseSw 49
SW49P.X=442.9411
SW49P.Y=449.8546
MotorBank.Y = 439.95 + (3*RND)
MotorBank.X = 448.6911 -1 + (2*RND)
Me.TimerEnabled = 1
PlaySoundAt SoundFX("fx_target",DOFContactors),f129a
Flasherswitch49.state =1
End Sub

Sub SW49_Timer:SW49P.X=442.6875:SW49P.Y=453.3662:MotorBank.Y = 439.95:MotorBank.X = 448.6911:Me.TimerEnabled = 0:Flasherswitch49.state =0:End Sub

Sub SW50_Hit
vpmTimer.PulseSw 50
SW50P.X=448.6911
SW50P.Y=449.8546
MotorBank.Y = 439.95 + (3*RND)
MotorBank.X = 448.6911 -1 + (2*RND)
Me.TimerEnabled = 1
PlaySoundAt SoundFX("fx_target",DOFContactors),f129a
'If Ballresting = True Then
 'DPBall.VelY = ActiveBall.VelY * 3
'End If
Flasherswitch50.state =1
End Sub

Sub SW50_Timer:SW50P.X=448.4375:SW50P.Y=453.3662:MotorBank.Y = 439.95:MotorBank.X = 448.6911:Me.TimerEnabled = 0:Flasherswitch50.state =0:End Sub

Sub SW51_Hit
vpmTimer.PulseSw 51
SW51P.X=454.0661
SW51P.Y=449.8546
MotorBank.Y = 439.95 + (3*RND)
MotorBank.X = 448.6911 -1 + (2*RND)
Me.TimerEnabled = 1
PlaySoundAt SoundFX("fx_target",DOFContactors),f129a
'If Ballresting = True Then
 'DPBall.VelY = ActiveBall.VelY * 3
'End If
Flasherswitch51.state =1
End Sub

Sub SW51_Timer:SW51P.X=453.8125:SW51P.Y=453.3662:MotorBank.Y = 439.95:MotorBank.X = 448.6911:Me.TimerEnabled = 0:Flasherswitch51.state =0:End Sub


'*********************************************************************
'* TARGETBANK MOVEMENT Taken from AFM written by Groni ***************
'*********************************************************************

Dim TBPos, TBDown, TBdif

Sub TBMove (enabled)
if enabled then
TBTimer.Enabled=1

PlaySound SoundFX("TargetBank",DOFContactors)

End If
End Sub

Sub TBTimer_Timer()	


IF TBPos = 0 Then

MotorBank.Y = 439.95
SW49P.Y=451.6104
SW50P.Y=451.6104
SW51P.Y=451.6104
MotorBank.X = 448.6911
SW49P.X=442.9411
SW50P.X=448.6911
SW51P.X=454.0661
End If

IF TBPos = 29 Then

MotorBank.Y = 439.95
SW49P.Y=451.6104
SW50P.Y=451.6104
SW51P.Y=451.6104
MotorBank.X = 448.6911
SW49P.X=442.9411
SW50P.X=448.6911
SW51P.X=454.0661
End If


Select Case TBPos
Case 0: MotorBank.Z=-20:SW49P.Z=-20:SW50P.Z=-20:SW51P.Z=-20:TBPos=0:TBDown=0:TBTimer.Enabled=0:Controller.Switch(52) = 0:Controller.Switch(53) = 1::SW49.isdropped=0:SW50.isdropped=0:SW51.isdropped=0:DPWall.isdropped=0:DPWall1.isdropped=1
Case 1: MotorBank.Z=-22:SW49P.Z=-22:SW50P.Z=-22:SW51P.Z=-22
Case 2: MotorBank.Z=-24:SW49P.Z=-24:SW50P.Z=-24:SW51P.Z=-24:Controller.Switch(53) = 0
Case 3: MotorBank.Z=-26:SW49P.Z=-26:SW50P.Z=-26:SW51P.Z=-26
Case 4: MotorBank.Z=-28:SW49P.Z=-28:SW50P.Z=-28:SW51P.Z=-28
Case 5: MotorBank.Z=-30:SW49P.Z=-30:SW50P.Z=-30:SW51P.Z=-30
Case 6: MotorBank.Z=-32:SW49P.Z=-32:SW50P.Z=-32:SW51P.Z=-32
Case 7: MotorBank.Z=-34:SW49P.Z=-34:SW50P.Z=-34:SW51P.Z=-34
Case 8: MotorBank.Z=-36:SW49P.Z=-36:SW50P.Z=-36:SW51P.Z=-36
Case 9: MotorBank.Z=-38:SW49P.Z=-38:SW50P.Z=-38:SW51P.Z=-38
Case 10: MotorBank.Z=-40:SW49P.Z=-40:SW50P.Z=-40:SW51P.Z=-40
Case 11: MotorBank.Z=-42:SW49P.Z=-42:SW50P.Z=-42:SW51P.Z=-42
Case 12: MotorBank.Z=-44:SW49P.Z=-44:SW50P.Z=-44:SW51P.Z=-44:
Case 13: MotorBank.Z=-46:SW49P.Z=-46:SW50P.Z=-46:SW51P.Z=-46:
Case 14: MotorBank.Z=-48:SW49P.Z=-48:SW50P.Z=-48:SW51P.Z=-48
Case 15: MotorBank.Z=-50:SW49P.Z=-50:SW50P.Z=-50:SW51P.Z=-50
Case 16: MotorBank.Z=-52:SW49P.Z=-52:SW50P.Z=-52:SW51P.Z=-52
Case 17: MotorBank.Z=-54:SW49P.Z=-54:SW50P.Z=-54:SW51P.Z=-54
Case 18: MotorBank.Z=-56:SW49P.Z=-56:SW50P.Z=-56:SW51P.Z=-56
Case 19: MotorBank.Z=-58:SW49P.Z=-58:SW50P.Z=-58:SW51P.Z=-58
Case 20: MotorBank.Z=-60:SW49P.Z=-60:SW50P.Z=-60:SW51P.Z=-60
Case 21: MotorBank.Z=-62:SW49P.Z=-62:SW50P.Z=-62:SW51P.Z=-62
Case 22: MotorBank.Z=-64:SW49P.Z=-64:SW50P.Z=-64:SW51P.Z=-64
Case 23: MotorBank.Z=-66:SW49P.Z=-66:SW50P.Z=-66:SW51P.Z=-66
Case 24: MotorBank.Z=-68:SW49P.Z=-68:SW50P.Z=-68:SW51P.Z=-68
Case 25: MotorBank.Z=-70:SW49P.Z=-70:SW50P.Z=-70:SW51P.Z=-70
Case 26: MotorBank.Z=-72:SW49P.Z=-72:SW50P.Z=-72:SW51P.Z=-72:Controller.Switch(52) = 0
Case 27: MotorBank.Z=-74:SW49P.Z=-74:SW50P.Z=-74:SW51P.Z=-74
Case 28: MotorBank.Z=-76:SW49P.Z=-76:SW50P.Z=-76:SW51P.Z=-76:SW49.isdropped=1:SW50.isdropped=1:SW51.isdropped=1:DPWALL.isdropped=1
Case 29: TBTimer.Enabled=0:TBDown=1:Controller.Switch(52) = 1:Controller.Switch(53) = 0
End Select

If TBDown=0 then TBPos=TBPos+1 
If TBDown=1 then TBPos=TBPos-1

TBdif = (-1 + (2*RND))/((TBPos/5)+1)

MotorBank.Y = MotorBank.Y + TBdif
SW49P.Y=SW49P.Y + TBdif
SW50P.Y=SW50P.Y + TBdif
SW51P.Y=SW51P.Y + TBdif
MotorBank.X = MotorBank.X + TBdif
SW49P.X=SW49P.X + TBdif
SW50P.X=SW50P.X + TBdif
SW51P.X=SW51P.X + TBdif


End Sub




	


Sub ShooterLane_Hit()
	Controller.Switch(23)=1
    Lanelight1.state = 1

End Sub

Sub ShooterLane_Unhit()
	Controller.Switch(23)=0
    Lanelight1.state = 0
End Sub

Dim frame, FinalFrame  'ArcadeTimer
FinalFrame = 126 'number of frames - 1
frame = 0

 Sub ArcadeTimer_Timer()       
	Arcade(frame).isdropped = True
	frame = frame + 1
	If frame = FinalFrame Then frame=0
	Arcade(frame).isdropped = False
 End Sub

Sub Trigger1_hit
	PlaySound "DROP_LEFT"
 End Sub

 Sub Trigger2_hit
	PlaySound "DROP_RIGHT"
 End Sub 


Sub Table_exit()
	Controller.Pause = False
	Controller.Stop
End Sub


Sub RLS_Timer()
              RampGate1.RotZ = -(Spinner4.currentangle)
              RampGate2.RotZ = -(Spinner1.currentangle)
              RampGate3.RotZ = -(Spinner3.currentangle)
              RampGate4.RotZ = -(Spinner2.currentangle)
              SpinnerT4.RotZ = -(sw44.currentangle)
              SpinnerT1.RotZ = -(sw36.currentangle)
End Sub
  
'primitive flippers!
Sub UpdateFlipperLogo
    LFLogo.RotY = LeftFlipper.CurrentAngle
    RFlogo.RotY = RightFlipper.CurrentAngle
	LFLogoUP.RotY = LeftFlipper1.CurrentAngle
End Sub

'******DROP TARGET PRIMITIVES******
Dim sw1up, sw2up, sw3up, sw4up
Dim PrimT

Sub PrimT_Timer
	if sw01.IsDropped = True then sw1up = False else sw1up = True
	if sw02.IsDropped = True then sw2up = False else sw2up = True
	if sw03.IsDropped = True then sw3up = False else sw3up = True
	if sw04.IsDropped = True then sw4up = False else sw4up = True
End Sub


Sub sw1T_Timer()
	If sw1up = True and sw1p.z < 0 then sw1p.z = sw1p.z + 3
	If sw1up = False and sw1p.z > -45 then sw1p.z = sw1p.z - 3
	If sw1p.z >= -45 then sw1up = False
End Sub

Sub sw2T_Timer()
	If sw2up = True and sw2p.z < 0 then sw2p.z = sw2p.z + 3
	If sw2up = False and sw2p.z > -45 then sw2p.z = sw2p.z - 3
	If sw2p.z >= -45 then sw2up = False
End Sub

Sub sw3T_Timer()
	If sw3up = True and sw3p.z < 0 then sw3p.z = sw3p.z + 3
	If sw3up = False and sw3p.z > -45 then sw3p.z = sw3p.z - 3
	If sw3p.z >= -45 then sw3up = False
End Sub

Sub sw4T_Timer()
	If sw4up = True and sw4p.z < 0 then sw4p.z = sw4p.z + 3
	If sw4up = False and sw4p.z > -45 then sw4p.z = sw4p.z - 3
	If sw4p.z >= -45 then sw4up = False
End Sub

Sub GIOn
	dim bulb
	for each bulb in GI
	bulb.state = 1
	next
	if GiBleedOpacity > 0 then gibleed.visible = 1
End Sub

Sub GIOff
	dim bulb
	for each bulb in GI
	bulb.state = 0
	next
	gibleed.visible = 0
End Sub

 'Sub RightSlingShot_Timer:Me.TimerEnabled = 0:End Sub
 
' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Sub Pins_Hit (idx)
	PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
	PlaySound "fx_target", 0, 1, Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub TargetBankWalls_Hit (idx)
	PlaySound "fx_target", 0, 1, Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
	PlaySound "metalhit_thin", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
	PlaySound "metalhit_medium", 0, Vol(ActiveBall)+2, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
	PlaySound "metalhit2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
	'PlaySound "gate4", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub


Sub RampDrop_Hit ' Launch ramp ball drop over small lip
	PlaySound "BallDrop", 0, Vol(ActiveBall)*.2, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub UpPost_Hit ' Launch ramp post
	PlaySound "metalhit2", 0, Vol(ActiveBall)+2, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub swplunger1_Hit ' Launch ramp switch
	PlaySoundAt "rollover", swplunger1
End Sub

Sub Rubbers_Hit(idx)
    PlaySoundAtBallVol "rubber_hit_" & Int(Rnd*3)+1, 1.6
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySoundAtBallVol "flip_hit_" & Int(Rnd*3)+1, 1.5
End Sub

Sub LeftFlipper1_Collide(parm)
    PlaySoundAtBallVol "flip_hit_" & Int(Rnd*3)+1, 1.5
End Sub

Sub RightFlipper_Collide(parm)
    PlaySoundAtBallVol "flip_hit_" & Int(Rnd*3)+1, 1.5
End Sub

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 200)
End Function

function AudioFade(ball)
    Dim tmp
    tmp = ball.y * 2 / Table.height-1
    If tmp > 0 Then
        AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / table.width-1
    If tmp > 0 Then
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
						PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*.4, Pan(BOT(b) ), 0, Pitch(BOT(b) )+30000, 1, 0, AudioFade(BOT(b) )
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
	PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 80, Pan(ball1), 0, Pitch(ball1), 0, 0
End Sub

Sub leftdrop_hit
	PlaySoundAt "BallDrop",leftdrop
	BumpStop
End Sub

Sub rightdrop_hit
	PlaySoundAt "BallDrop",rightdrop
	BumpStop
End Sub

'**************************************************************************
'                 Positional Sound Playback Functions by DJRobX
'**************************************************************************

'Set position as table object (Use object or light but NOT wall) and Vol to 1

Sub PlaySoundAt(sound, tableobj)
		PlaySound sound, 1, 1, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub


'Set all as per ball position & speed.

Sub PlaySoundAtBall(sound)
		PlaySound sound, 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
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

' Plastic Ramp Sounds

Dim NextOrbitHit:NextOrbitHit = 0 

Sub PlasticRampBumps_Hit(idx)
	if BallVel(ActiveBall) > .3 and Timer > NextOrbitHit then
		RandomBump 2, Pitch(ActiveBall)
		' Schedule the next possible sound time.  This prevents it from rapid-firing noises too much. 
		' Lowering these numbers allow more closely-spaced clunks.
		NextOrbitHit = Timer + .1 + (Rnd * .2)
	end if 
End Sub

' Requires rampbump1 to 7 in Sound Manager
Sub RandomBump(voladj, freq)
	dim BumpSnd:BumpSnd= "rampbump" & CStr(Int(Rnd*7)+1)
		PlaySound BumpSnd, 0, Vol(ActiveBall)*voladj, Pan(ActiveBall), 0, freq, 0, 1, AudioFade(ActiveBall)
End Sub

Sub BumpStop()
dim i:for i=1 to 4:StopSound "rampbump" & i:next
NextOrbitHit = Timer + 1
End Sub



Sub Hanbewegung_timer

 '************************************
 '***** Hanibal special moving Light1
 '************************************


dim anglehan, dirhan, anglezhan, dirzhan, angleXYoffsethan, lengthhanoffset, hanmoveX, hanmoveY


    lighthanmovepos(0) = recognizer.x + recognizer.transx
	lighthanmovepos(1) = recognizer.y +30 
	lighthanmovepos(2) = recognizer.z +200  
	anglehan = - recognizer.transx
	anglezhan = -recognizer.rotz
    lengthhanoffset = 200
    Bewegungslicht.x = lighthanmovepos(0)
	Bewegungslicht.y = lighthanmovepos(1)


    'debug.print anglehan & ":" &  anglezhan & ": " &  Bewegungslicht.x &  ":" &  Bewegungslicht.y




'		recognizer.transx=XLocation
'		recognizer.rotz=zrot

End Sub


 Sub Augen_Timer

' ******Hanibals Random Lights Script

Lightcycle1.Intensity = (5+(1*Rnd))
Lightcycle2.Intensity =Lightcycle1.Intensity

Ship1.Intensity = (2+(1*Rnd))
Ship2.Intensity = (2+(1*Rnd))


Lightcycle3.Intensity = (5+(2*Rnd))
Lightcycle4.Intensity =	Lightcycle3.Intensity 
Lightcycle5.Intensity =	Lightcycle3.Intensity 

Schild.Intensity = (35+(10*Rnd))


Flasher1a1.Intensity = (150+(30*Rnd))
Flasher2a1.Intensity = Flasher1a1.Intensity 
'Flasher3a.Intensity = (150+(30*Rnd))
'Flasher4a.Intensity = Flasher3a.Intensity 
Flasher5a.Intensity = (60+(30*Rnd))
Flasher6a.Intensity = Flasher5a.Intensity 

Flasher7.Intensity = (15+(3*Rnd))
Flasher7a.Intensity =  (Flasher7.Intensity  *3)

Lanelight1.Intensity = (100+(10*Rnd))
Lanelight.Intensity = Lanelight1.Intensity

Linkerflasher.Intensity = (50+(10*Rnd))
Rechterflasher.Intensity = Linkerflasher.Intensity

l28a.Intensity = (50+(2*Rnd))
l28b.Intensity = l28a.Intensity
l28c.Intensity = l28a.Intensity
l29a.Intensity = l28a.Intensity
l29b.Intensity = l28a.Intensity
l29c.Intensity = l28a.Intensity
Kartenlichtlinks.Intensity = (10+(1*Rnd))
Kartenlichtlinks1.Intensity = Kartenlichtlinks.Intensity


l46.Intensity = (2+(10*Rnd))
l47.Intensity = (2+(10*Rnd))
l48.Intensity = (2+(7*Rnd))


 End Sub



 '************************************
 '***** HNew GI Controller
 '************************************

'*************************
' GI - needs new vpinmame
'*************************

Set GICallback = GetRef("GIUpdate")

Sub GIUpdate(no, Enabled)
	If enabled Then
		GIOn
    Else 
		GIOff
	end if 
End Sub


'*********** BALL SHADOW *********************************
Dim BallShadow:BallShadow = Array (BallShadow1, BallShadow2, BallShadow3,Ballshadow4,Ballshadow5,Ballshadow6,Ballshadow7,Ballshadow8,Ballshadow9,Ballshadow10)

Sub BallShadowUpdate()
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
			BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table.Width/2))/7)) '+ 10
		Else
			BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table.Width/2))/7)) '- 10
		End If
	    ballShadow(b).Y = BOT(b).Y + 10
 
		If BOT(b).Z > 20 Then
			BallShadow(b).visible = 1
		Else
			BallShadow(b).visible = 0
		End If
	Next
End Sub





''keydown
'if KeyCode = LeftFlipperKey then vpmFlipsSam.FlipL True : vpmFlipsSam.FlipUL True
'if KeyCode = RightFlipperKey then vpmFlipsSam.FlipR True : vpmFlipsSam.FlipUR True

''KeyUp
'if KeyCode = LeftFlipperKey then vpmFlipsSam.FlipL False : vpmFlipsSam.FlipUL False
'if KeyCode = RightFlipperKey then vpmFlipsSam.FlipR False : vpmFlipsSam.FlipUR False




dim vpmFlipsSAM : set vpmFlipsSAM = New cvpmFlipsSAM : vpmFlipsSAM.Name = "vpmFlipsSAM"

'*************************************************
Sub InitVpmFlipsSAM()
	vpmFlipsSAM.CallBackL = SolCallback(vpmflipsSAM.FlipperSolNumber(0))          'Lower Flippers
	vpmFlipsSAM.CallBackR = SolCallback(vpmFlipsSAM.FlipperSolNumber(1))
	On Error Resume Next
		if cSingleLflip or Err then vpmFlipsSAM.CallbackUL=SolCallback(vpmFlipsSAM.FlipperSolNumber(2))
		err.clear
		if cSingleRflip or Err then vpmFlipsSAM.CallbackUR=SolCallback(vpmFlipsSAM.FlipperSolNumber(3))
	On Error Goto 0
	'msgbox "~Debug-Active Flipper subs~" & vbnewline & vpmFlipsSAM.SubL &vbnewline& vpmFlipsSAM.SubR &vbnewline& vpmFlipsSAM.SubUL &vbnewline& vpmFlipsSAM.SubUR' & 

End Sub



'New Command - 
'vpmflipsSam.RomControl=True/False 		-	 True for rom controlled flippers, False for FastFlips (Assumes flippers are On)


'Debug Stuff ~~~~~~~~~~~~~
'Sub TestFF_Timer()	'Testing switching back and forth between fastflips and rom controlled flips on a timer 
'	me.interval = 500'RndNum(100, 500)
'	vpmFlipsSAM.RomControl = Not vpmFlipsSAM.RomControl
'End Sub
'
'tb.timerinterval = -1 : tb.timerenabled=1	'debug textbox
'Sub tb_Timer()
'	tb.text = vpmFlipsSAM.SolState(0) & " " & vpmFlipsSAM.ButtonState(0) & " " & vbnewline & vpmFlipsSAM.romcontrol
'End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~


Class cvpmFlipsSAM	'test fastflips with support for both Rom and Game-On Solenoid flipping
	Public TiltObjects, DebugOn, Name, Delay
	Public SubL, SubUL, SubR, SubUR, FlippersEnabled,  LagCompensation, ButtonState(3), Sol	'set private
	Public RomMode,	SolState(3)'set private
	Public FlipperSolNumber(3)	'0=left 1=right 2=Uleft 3=URight

	Private Sub Class_Initialize()
		dim idx :for idx = 0 to 3 :ButtonState(idx)=0:SolState(idx)=0: Next : Delay=0: FlippersEnabled=1: DebugOn=0 : LagCompensation=0 : Sol=0 : TiltObjects=1
		SubL = "NullFunction": SubR = "NullFunction" : SubUL = "NullFunction": SubUR = "NullFunction"
		RomMode=True :FlipperSolNumber(0)=sLLFlipper :FlipperSolNumber(1)=sLRFlipper :FlipperSolNumber(2)=sULFlipper :FlipperSolNumber(3)=sURFlipper
		SolCallback(33)="vpmFlipsSAM.RomControl = not "
	End Sub

	'set callbacks
	Public Property Let CallBackL(aInput) : if Not IsEmpty(aInput) then SubL  = aInput :SolCallback(FlipperSolNumber(0)) = name & ".RomFlip(0)=":end if :End Property	'execute
	Public Property Let CallBackR(aInput) : if Not IsEmpty(aInput) then SubR  = aInput :SolCallback(FlipperSolNumber(1)) = name & ".RomFlip(1)=":end if :End Property
	Public Property Let CallBackUL(aInput): if Not IsEmpty(aInput) then SubUL = aInput :SolCallback(FlipperSolNumber(2)) = name & ".RomFlip(2)=":end if :End Property	'this should no op if aInput is empty
	Public Property Let CallBackUR(aInput): if Not IsEmpty(aInput) then SubUR = aInput :SolCallback(FlipperSolNumber(3)) = name & ".RomFlip(3)=":end if :End Property
	
	Public Property Let RomFlip(idx, ByVal aEnabled)
		aEnabled = abs(aEnabled)
		SolState(idx) = aEnabled
		If Not RomMode then Exit Property
		Select Case idx
			Case 0 : execute subL & " " & aEnabled
			Case 1 : execute subR & " " & aEnabled
			Case 2 : execute subUL &" " & aEnabled
			Case 3 : execute subUR &" " & aEnabled
		End Select
	End property
	
	Public Property Let RomControl(aEnabled) 		'todo improve choreography
		'MsgBox "Rom Control " & CStr(aEnabled)
		RomMode = aEnabled
		If aEnabled then 					'Switch to ROM solenoid states or button states
			Execute SubL &" "& SolState(0)
			Execute SubR &" "& SolState(1)
			Execute SubUL &" "& SolState(2)
			Execute SubUR &" "& SolState(3)
		Else
			Execute SubL &" "& ButtonState(0)
			Execute SubR &" "& ButtonState(1)
			Execute SubUL &" "& ButtonState(2)
			Execute SubUR &" "& ButtonState(3)
		End If
	End Property
	Public Property Get RomControl : RomControl = RomMode : End Property

	public DebugTestKeys, DebugTestInit	'orphaned (stripped out the debug stuff)

	Public Property Let Solenoid(aInput) : if not IsEmpty(aInput) then Sol = aInput : end if : End Property	'set solenoid
	Public Property Get Solenoid : Solenoid = sol : End Property
	
	'call callbacks
	Public Sub FlipL(ByVal aEnabled)
		aEnabled = abs(aEnabled) 'True / False is not region safe with execute. Convert to 1 or 0 instead.
		DebugTestKeys = 1
		ButtonState(0) = aEnabled	'track flipper button states: the game-on sol flips immediately if the button is held down (1.1)
		If FlippersEnabled and Not Romcontrol or DebugOn then execute subL & " " & aEnabled end If
	End Sub

	Public Sub FlipR(ByVal aEnabled)
		aEnabled = abs(aEnabled) : ButtonState(1) = aEnabled : DebugTestKeys = 1
		If FlippersEnabled and Not Romcontrol or DebugOn then execute subR & " " & aEnabled end If
	End Sub

	Public Sub FlipUL(ByVal aEnabled)
		aEnabled = abs(aEnabled)  : ButtonState(2) = aEnabled
		If FlippersEnabled and Not Romcontrol or DebugOn then execute subUL & " " & aEnabled end If
	End Sub	

	Public Sub FlipUR(ByVal aEnabled)
		aEnabled = abs(aEnabled)  : ButtonState(3) = aEnabled
		If FlippersEnabled and Not Romcontrol or DebugOn then execute subUR & " " & aEnabled end If
	End Sub	
	
	Public Sub TiltSol(aEnabled)	'Handle solenoid / Delay (if delayinit)
		If delay > 0 and not aEnabled then 	'handle delay
			vpmtimer.addtimer Delay, Name & ".FireDelay" & "'"
			LagCompensation = 1
		else
			If Delay > 0 then LagCompensation = 0
			EnableFlippers(aEnabled)
		end If
	End Sub
	
	Sub FireDelay() : If LagCompensation then EnableFlippers 0 End If : End Sub
	
	Public Sub EnableFlippers(aEnabled)	'private
		If aEnabled then execute SubL &" "& ButtonState(0) :execute SubR &" "& ButtonState(1) :execute subUL &" "& ButtonState(2): execute subUR &" "& ButtonState(3)':end if
		FlippersEnabled = aEnabled
		If TiltObjects then vpmnudge.solgameon aEnabled
		If Not aEnabled then
			execute subL & " " & 0 : execute subR & " " & 0
			execute subUL & " " & 0 : execute subUR & " " & 0
		End If
	End Sub
End Class
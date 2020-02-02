
'                                                                                                ,***
'                                                        .,,*,,,.                      .                 /,
'                                .*(((*.            .,,,,.,,..,,.,,,,.            ./((((       .,.,      **.     *,,.
'                                .*(#(((((,.     ,,,,,,,,,,/##/,,,,,,,,,.     .*((((((((       .//(/*.  .**,  .*((/*
'                                 *%((/(((/(((//*,,,,,,*(%&(,,#&&(*,,,,,,//((((/((((((%#           .,**,**,/****
'                                  *((/((#(((((((#(((#%%/*  .,*/(#%%(#(((#(((((#((((((,               ,**/***,
'                                  *(((((#(((/(((#((((((((,/(((((####(%(((((#(((((((,             .***/***,**,
'                               .,,/(#((#####((#(,((((((%((((#%%%#(((/*((((#(#((((#(/*,.     .,,(*,   .*/,   .//,,.
'                           .,*,,,,*(###((((#(##(.((((#(((((((((*(##/(((%(((##((((###(,.,,,*,  .//,      **.     */*
'                        .,*,,,.,,,((###(##((###(.((###((#(###(*,(#####(%(######((####/*,,.,,,*,.        /,
'                   .,,*,,,,,,*(&&%#(###((((####(*((###((#(####((&(###(#((####((((###(#%&%#*,..,,,,,,.  ,,,,
'               ..,**,,..,*/%/..*(######/#############(#%(##########(#(######/######(( .*#(,,,,,,,*,,,.
'           .,,,.,,,,*/%&%/,     ./&%%&&&%%%%%%%%&&%%%%%%%.*%%%##%&&%%%&&%%%%%%%&&&&&&      ./&**,,,,.,,,..
'       .,,,,*/////(%&&/*,                    */////////,    ,(((/*/(//*    *///***********.,*****(&&%(**,,,,,,,,.
'    ,,,,,,,,(((((((#####(((,    .*//((((/*.  ((#######(,   ./(######((*    ((###########((*((#######(((((/*,,,,.,,,.
' .*,.,,,,*#&%##############(,*((##/(((####(((##(####((%,   ./#(####(##/    (#(####((((####(#(############(#&&%/,,,,.,,.
'.,,,,*/%&%* (#(#####(#(##(((#(((*,..,***,/(###%#####((.      *((###(/.     .*(####(%%(((##(%%#####(######((, ,(&%(,*,,,
',, .*(,       ((((((/,((////(//,.. ..,,,,***(##(#(((((       ,((((#(*       .((((#(#(((%#(((/(((((((/((((((/     .(,,.*
',,,,,,        (((((((*((((((%//*.....     .,*(##((((((    ,  ,((((((*       .(((((((((((%%%**(((((((*((((((/      ,,,,*
',/,,.,*,.     ((//((((((///#(//#,.  .*(##(///((#((((((   *(((/((((((*   *((,,(((((((((((..  *((//(((((///((*   ,,,,..,(
' /%(,,,,.,,,,,(/////////((%#///#//((((////(#*//((///(( ,/(/(##((///(*  /(((#%((///(%((((,((.*((//////*/*/(#/,,,..,,,*%,
'  ,(&%(*,,,..*(//*//(/**/(%#/**/*,,,,,...,//*/(#/***((((/*/#%((/,**((((/**/#/(/***(##((((((((#/**/(//*/((%(,.,,,*/%.
'       *%&%(,,(//*//(/***/(%#/*,,,.    *(%%%###%*,,,,,,***(##/******,,,,*/#%(/,,,,/((##(/*/#((/**/((/**/(/**(%,.
'          .(&&%//*//%(**,,*(&%(/,,,,,/%%%%%%%%%%%####%%%&&&*(&&&&&%%####(((((/*,,,,,***,,,/#/(/,**(#(/*//%&%/,
'             .(/***/##/*,,,*/%%(/(#%%%%&@@&%%%(,,,..         .,,,...,,*/.(%&&%#(///**,,*(%(/,,,*(##/**/((
'             .(/*,,*/#(/,,,*/(*,*/&%%%%#/,.,%%%#      **,     (%%%%, */*/(###*  ,%%%%%%%%%%(//,,,,,*((*,,/(*
'            *(/,,,,,*#%(/(#%&(*,*#%%%%*.   ,%&&/.   *#%#%*   ,%%%%/ #%%%%%%%#%*#%%%%%%%#%%/#&&%((/**(#/,,*/(*.
'            /(*,,*/(#%*/,  .*%&&%%%%%*.,,/#%%%%%%(..%%%%#%( .%%%&/./%%%%/#%%%%#%%%*%%%%&%*    ./%%%%(/*,,//.
'            /(((#&%(,          #%%%**#%%%%%%%%#. /%%%&%%%%%%%%#,/%%%%/#%%%%&%%%%%%%%%%#*          .,%%%((//.
'            (%%/,             .%%%%(/%&%%%%#%%%%#//%&&(,/%%%%%%%%%%%%%%*#%%##%%&&&%%%%%/.               ./(%#,
'                              ,%###. /%%%&&&%%%%%%%%%%###%%%#%%%%%%%%(,.**(&@&(,..... .
'                              *%%#%. ,%%%%%%%%%%%%%(/(##%#/,,******.***#&*.
'                              .%%%%#(,*/(%%%%%%%//#&%#/*,,,..,,,,*(%&(*
'                               *%%%%%%%%%###%%%%*   .*%&&(****(%/.
'                                ,(%%%%%%%%%%#((*        ./%&&%/,
'
' High Roller Casino
' Stern 2001
'
' Table by DJRobX & Dids666
' Plastics renderings and dynamic drop shadow technique by Bord
'
' Special thanks to:
'
' Stern for creating this fun table.   Go put some quarters in a real machine soon!
' Knorr for his Stern Apron prim.
' Destruk, TAB, bmiki75 for the previous versions of this table that were used as a reference for this one
' VPX and VPinmame dev teams
' NFozzy, Rothbauerw for their excellent physics tutorial and fading lights classes.
' Ninuzzu, JPSalas and others who have contributed script and other bits that make this work!
' Nailbuster for his autoQA script
' G5k, Nailbuster, TerryRed, Flupper, Freezy and all of the other inspirational contributors that keep this hobby interesting
' RandR, Dazz, and Noah Fentz for hosting our community
'
' About this Table
'
' Straight up, the resources we had weren't great.  There's no pretty 4k playfield scan, no crisp plastic scans.   Photoshop and upscalers
' can only take it so far.  If better resources become available we will update the table!
'
' Vpinmame 3.3 rev 4895 or later is REQUIRED!   This table has a 4-solenoid stepper motor that cannot be read correctly without
' this code update.
'
' Known issues:
'
' 1) The ROM complains a lot about "possibly broken" switches.  The switches that the diagnostic report identifies are very random.
' 2) The game may hang onto a ball at the pop entrance under the roulette wheel.   I've caught this a couple times and confirmed
'    that VPinMame thinks Switch 32 is active but the test menu shows it is not.   So this must be either a bug in the game code or a bug in the emulation, probably the latter.
'    Once the game does a ball hunt it will let go of it, so it's not a showstopper.
' 3) The roulette wheel is fairly complex for VPX.  In multiball, if it dumps 3+ balls onto it at once, some of that dogpiling activity
'    may slow down some systems.   MaxBallsRoulette will delay autoplunges to try to mitigate this.
'
' This table is for personal use only, and should not be included with a commercial product.
'
' This table will be free to mod without permission after Septeember 1,2020.
' Before then, please reach out to DJRobX or Dids666 before distributing mods of this table.




Option Explicit
Randomize

Const cGameName = "hirolcas"

Dim AutoQA, MaxBallsRoulette, LUTmeUP, DisableLUTSelector

MaxBallsRoulette = 2        'Limit number of balls that can be active in the roulette wheel at once.   Reduce to 1 if you are getting multiball stutter.
AutoQA= False               'Enable AI player
LUTMeUp = 2  				'0 = No LUT.  1 = More vibrant color, 2 = Photoshop autocurve based on screeenshot, 3 = Autocurve with more saturation
DisableLUTSelector = 0  	' Disables the ability to change LUT option with magna saves in game when set to 1


On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

Dim DesktopMode:DesktopMode = Table1.ShowDT
Dim UseVPMDMD:UseVPMDMD = True

LoadVPM "02000000", "Sega.VBS", 3.43

Const UseSolenoids = 2
Const UseLamps = False
Const UseSync = False
Const MaxLut = 3

' Standard Sounds
Const SSolenoidOn = ""
Const SSolenoidOff = ""
Const SCoin = "fx_coin"

Dim NullFader : set NullFader = new NullFadingObject
Dim FadeLights : Set FadeLights = New LampFader
Dim bsTrough, dtDrop,bsVLock,mMagnet,mRoulette,mRoulMag,bsSlot,mSlotRamp
Dim BallInRampArea: BallInRampArea = False

NoUpperLeftFlipper
NoUpperRightFlipper

Sub Table1_Init
	if not DesktopMode then ScoreText.Y = -500
    if table1.VersionMinor < 6 AND table1.VersionMajor = 10 then MsgBox "This table requires VPX 10.6, you have " & table1.VersionMajor & "." & table.VersionMinor
	if VPinMAMEDriverVer < 3.57 then MsgBox "This table requires core.vbs 3.57 or higher, which is included with VPX 10.6.  You have " & VPinMAMEDriverVer & ". Be sure scripts folder is up to date, and that there are no old .vbs files in your table folder."

     With Controller
         .GameName = cGameName
         If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
         .SplashInfoLine = ""
         .HandleKeyboard = 0
         .ShowTitle = 0
         .ShowDMDOnly = 1
         .ShowFrame = 0
         .HandleMechanics = False
         .Hidden = UseVPMDMD
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With

	if Controller.Version < "03030000" then MsgBox "This table requires vpinmame 3.3 rev 4895 or later. Yours reoprts " & Controller.Version & ".  Please get the latest SAMBuild from VPUniverse"

	AutoPlunger.Pullback

 	Set dtDrop=New cvpmDropTarget
	dtDrop.InitDrop Array(sw17,sw18,sw19,sw20),Array(17,18,19,20)

   	Set bsTrough=New cvpmBallStack
 	with bsTrough
		.InitSw 0,14,13,12,11,0,0,0
		.InitKick BallRelease,90,7
		bsTrough.InitExitSnd Soundfx("fx_ballrel",DOFContactors), Soundfx("fx_solenoid",DOFContactors)
		.Balls=4
 	end with

	Set bsVLock=New cvpmVLock
	bsVLock.InitVLock Array(LockTr1,LockTr2,LockTr3),Array(LockKick1,LockKick2,LockKick3),Array(25,26,27)
	bsVLock.CreateEvents "bsVLock"


     ' Nudging
     vpmNudge.TiltSwitch = 56
     vpmNudge.Sensitivity = 5
'     vpmNudge.TiltObj = Array(sw25, sw26, sw27, sw39, sw40)


	Set mMagnet=New cvpmMagnet
	mMagnet.InitMagnet OrbitMagnet,100
	mMagnet.Solenoid=13
	mMagnet.GrabCenter=1
'	mMagnet.CreateEvents "mMagnet"


	vpmTimer.InitTimer RF1, True
	RouletteFloor.TimerInterval = 7
'
	Set mSlotRamp=New cvpmMyMech
	' In VPX 10.7+ should use "vpmMechFourStepSol" which happens to be the same as setting both StepSol and TwoDirSol
	' mSlotRamp.MType=vpmMechFourStepSol+vpmMechStopEnd+vpmMechLinear+vpmMechFast
	mSlotRamp.MType=vpmMechStepSol+vpmMechTwoDirSol+vpmMechStopEnd+vpmMechLinear+vpmMechFast
	mSlotRamp.Sol1=26
	mSlotRamp.Length=300'150
	mSlotRamp.Steps=160
	mSlotRamp.AddSw 52,159,160
	mSlotRamp.AddSw 38,159,160
	mSlotRamp.Callback=GetRef("UpdateSlot")
	mSlotRamp.Start

	InitLamps
'	vpmtimer.addtimer 500, "PulseTimer.Interval=20 '"
	LUTBox.Visible = 0
	SetLUT

End Sub

Sub SetLUT
	Select Case LUTmeUP
		Case 0:table1.ColorGradeImage = 0
		Case Else: table1.ColorGradeImage = "LUT-" & CStr(LUTmeUP)
	end Select
end sub

Sub LUTBox_Timer
	LUTBox.TimerEnabled = 0
	LUTBox.Visible = 0
End Sub

Sub ShowLUT
	LUTBox.visible = 1
	LUTBox.text = "LUTmeUP: " & CStr(LUTmeUP)
	LUTBox.TimerEnabled = 1
End Sub

Sub Table1_KeyDown(ByVal keycode)
   If keycode = RightMagnaSave Then
		if DisableLUTSelector = 0 then
			LUTmeUP = LUTMeUp + 1
			if LutMeUp > MaxLut then LUTmeUP = 0
			SetLUT
			ShowLUT
		end if
	end if
	If keycode = LeftMagnaSave Then
          if DisableLUTSelector = 0 then
			LUTmeUP = LUTMeUp - 1
			if LutMeUp < 0 then LUTmeUP = MaxLut
			SetLUT
			ShowLUT
		end if
	end if
	If keycode = LeftTiltKey Then Nudge 90, 2
	If keycode = RightTiltKey Then Nudge 270, 2
	If keycode = CenterTiltKey Then	Nudge 0, 2

	If vpmKeyDown(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.PullBack: PlaySound "fx_plungerpull",0,1,0.25,0.25: 	End If
End Sub

Sub Table1_KeyUp(ByVal keycode)
	If keycode = PlungerKey Then Plunger.Fire: PlaySound "fx_plunger",0,1,0.25,0.25
	If vpmKeyUp(keycode) Then Exit Sub

End Sub

Sub Table1_Paused:Controller.Pause = 1:End Sub
Sub Table1_UnPaused:Controller.Pause = 0:End Sub
Sub Table1_Exit:Controller.Stop:End Sub


Sub ShooterLane_Hit
End Sub

Sub BallRelease_UnHit
End Sub


Sub Drain_Hit()
	PlaySound "fx_drain",0,1,0,0.25
	bsTrough.AddBall Me
End Sub


Set GICallback = GetRef("GIUpdate")

Sub GIUpdate(no, Enabled)
	SetLamp 99, Enabled
'	Dim light
'	For each light in GI:light.State = Enabled: Next
End Sub


Sub HandleFlip(obj, angle)
	angle = (angle mod 360)
	dim ea:ea = obj.EndAngle
	if CInt(ea) <> CInt(angle) then
		obj.StartAngle = ea
		obj.EndAngle = angle
		obj.RotateToEnd
	end if
end sub

sub RF1_Collide(parm):RouletteFlippersCollide():End Sub
sub RF2_Collide(parm):RouletteFlippersCollide():End Sub
sub RF3_Collide(parm):RouletteFlippersCollide():End Sub
sub RF4_Collide(parm):RouletteFlippersCollide():End Sub
sub RF6_Collide(parm):RouletteFlippersCollide():End Sub
sub RF7_Collide(parm):RouletteFlippersCollide():End Sub
sub RF8_Collide(parm):RouletteFlippersCollide():End Sub


Sub RouletteFlippersCollide()
	debug.print "FlipColl"
	RandomSoundFlipper()
	Dim BV:BV =  BallVel(ActiveBall)

	' VPX doesn't seem to do elasticity for the flipper tops.  Simulate some random bounce by diverting some velocity to Z up if the ball is moving fast.
	if (Rnd < .8 and BV > 3) then
		debug.print "Spring"
		Dim amt: amt = Rnd
		ActiveBall.VelZ = (amt * BV) + 10
		if ActiveBall.VelZ > 20 then ActiveBall.VelZ = 20

		ActiveBall.VelX = ActiveBall.VelX * (1 - amt)
		ActiveBall.VelY = ActiveBall.VelY * (1 - amt)

	end if
End Sub


Sub RouletteFloor_Hit
	StopBallSpin
	'PlaySoundAtBall "fx_metalhit"
End Sub


Sub RouletteHole_Hit
	' ** Silly hack - i'm not sure if this is a ROM bug but the table may hold the ball in multiball up in the entry lock
    ' while another one is going round in the roulette wheen, without
	' releasing the ball, which doesn't leave room for another ball to fall.  So if there is a ball waiting in the raised post and another
    ' is coming down, force release to make room for this ball on the next loop.
	if Controller.Switch(32)<>0 and FadeLights.state(122)=0 then
		SolPopEntry(True)
		vpmtimer.addtimer 500, "SolPopEntry(False) '"
	end if

'	Debug.Print "RouletteHole: " & CStr(Round(BallVel(ActiveBall) , 2)) & " X " & ActiveBall.x & " Y " & ActiveBall.y & " Z " & ActiveBall.z
	StopBallSpin
	if BallVel(ActiveBall) > 1 then
		' We get crazy ball spin action if the ball hits this the wrong way, calm it down.   This doesn't
	    ' fully fix the problem, either, so I added a sketchy-AF one way gate
		ActiveBall.VelX = 0
		ActiveBall.VelY = 2
		ActiveBall.VelZ = .1
	  	'Debug.Print "RouletteHole STOP: " & CStr(Round(BallVel(ActiveBall) , 2)) & " X " & ActiveBall.x & " Y " & ActiveBall.y & " Z " & ActiveBall.z
		PlaySound "fx_balldrop", 0, 1, AudioPan(ActiveBall), 0, 1, 1, 0, AudioFade(ActiveBall)
	end if


	'	ActiveBall.Z = 40
   ' vpmTimer.PulseSw 32

End Sub

dim wp: wp = 0

Sub RouletteFloor_Timer
	wp = (wp + .6)
	if wp > 360 then wp = wp - 360
	dim aNewPos:aNewPos = 360-wp
	Roulette.RotY = aNewPos -14
	rsflasher.RotZ = aNewPos
    Primitive3.RotY = aNewPos
	HandleFlip RF1, aNewPos + 106
	HandleFlip RF2, aNewPos + 106 + 45
	HandleFlip RF3, aNewPos + 106 + 90
	HandleFlip RF4, aNewPos + 106 + 90 + 45
	HandleFlip RF5, aNewPos + 286
	HandleFlip RF6, aNewPos + 286 + 45
	HandleFlip RF7, aNewPos + 286 + 90
	HandleFlip RF8, aNewPos + 286 + 90 + 45
    Select Case ( ((aNewPos + 333) MOD 360) \ 45 )
	Case 0:Controller.Switch(41)=1:Controller.Switch(42)=0:Controller.Switch(43)=0:Controller.Switch(44)=0'13
	Case 1:Controller.Switch(41)=1:Controller.Switch(42)=0:Controller.Switch(43)=0:Controller.Switch(44)=1'0
	Case 2:Controller.Switch(41)=0:Controller.Switch(42)=0:Controller.Switch(43)=0:Controller.Switch(44)=1'18
	Case 3:Controller.Switch(41)=0:Controller.Switch(42)=0:Controller.Switch(43)=1:Controller.Switch(44)=1'26
	Case 4:Controller.Switch(41)=0:Controller.Switch(42)=0:Controller.Switch(43)=1:Controller.Switch(44)=0'7
	Case 5:Controller.Switch(41)=0:Controller.Switch(42)=1:Controller.Switch(43)=1:Controller.Switch(44)=0'00
	Case 6:Controller.Switch(41)=0:Controller.Switch(42)=1:Controller.Switch(43)=0:Controller.Switch(44)=0'11
	Case 7:Controller.Switch(41)=1:Controller.Switch(42)=1:Controller.Switch(43)=0:Controller.Switch(44)=0'21
	End Select
End Sub

Dim MechPos:MechPos = 0

Sub UpdateSlot(NewPos, aSpeed, LastPos)
	SlotRampPrim.RotX = (NewPos * 20 / 160) - 21
	MechPos = NewPos
End Sub

sub SyncSlotRamp
	if not BallInRampArea then
		if MechPos < 10 then
			SlotRampUp.Collidable = True
			SlotBlock.Collidable = True
		elseif MechPos < 150 then
			SlotRampUp.Collidable = false
			SlotRampMid.Collidable = True
			SlotBlock.Collidable = True
		Else
			SlotRampUp.Collidable = False
			SlotRampMid.Collidable = False
			SlotBlock.Collidable = False
		end if
	end if
End Sub

sub StopBallSpin
	ActiveBall.AngMomZ = 0
	ActiveBall.AngMomX = 0
	ActiveBall.AngMomY = 0
end sub

sub OrbitMagnet_Hit
	debug.print "magnethit"
	StopBallSpin
	mMagnet.AddBall ActiveBall

end sub

sub OrbitMagnet_UnHit
	debug.print "magnetunhit"
	mMagnet.RemoveBall ActiveBall
end sub



'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep, RStepB, LStepB

Sub sw59_Slingshot
	vpmTimer.PulseSw 59
	PlaySoundAt SoundFX("fx_slingshot", DOFContactors), sw58 :    LSling.Visible = 0
    LSling3.Visible = 1
    Lslingm1.TransZ = -20
    LStep = 0
    sw59.TimerEnabled = 1
End Sub			'

Sub sw59_Timer
    Select Case LStep
        Case 3:LSLing3.Visible = 0:LSLing2.Visible = 1:lslingm1.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:LSlingm1.TransZ = 0:me.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub sw62_Slingshot
	vpmTimer.PulseSw 62:PlaySoundAt SoundFX("fx_slingshot", DOFContactors), sw61
	RSling.Visible = 0
    RSling3.Visible = 1
    Rslingm1.TransZ = -20
    RStep = 0
    sw62.TimerEnabled = 1
End Sub

Sub sw62_Timer
    Select Case RStep
        Case 3:RSling3.Visible = 0:RSLing2.Visible = 1:rslingm1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:rslingm1.TransZ = 0:me.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub


Sub sw47_Slingshot
	vpmTimer.PulseSw 47:PlaySoundAt SoundFX("fx_slingshot", DOFContactors), L130a
	LSlingB.Visible = 0
    LSlingB3.Visible = 1
    LslingBM.TransZ = -20
    LStepB = 0
    sw47.TimerEnabled = 1
End Sub

Sub sw47_Timer
	Select Case LStepB
        Case 3:LSLingB3.Visible = 0:LSLingB2.Visible = 1:lslingbm.TransZ = -10
        Case 4:LSLingB2.Visible = 0:LSLingB.Visible = 1:lslingbm.TransZ = 0:me.TimerEnabled = 0
    End Select
    LStepB = LStepB + 1
End Sub



Sub sw48_Slingshot
	vpmTimer.PulseSw 48:PlaySoundAt SoundFX("fx_slingshot", DOFContactors), LockKick3
	RSlingB.Visible = 0
    RSlingB3.Visible = 1
    RslingBM.TransZ = -20
    RStepB = 0
    sw48.TimerEnabled = 1
End Sub

Sub sw48_Timer
	Select Case RStepB
        Case 3:RSLingB3.Visible = 0:RSLingB2.Visible = 1:rslingbm.TransZ = -10
        Case 4:RSLingB2.Visible = 0:RSLingB.Visible = 1:rslingbm.TransZ = 0:me.TimerEnabled = 0
    End Select
    RStepB = RStepB + 1
End Sub

Sub LockOut_Hit
	StopBallSpin
End Sub


'************************************************
' Solenoids
'************************************************

SolCallback(1) = 	"vpmTimer.PulseSw 15:bsTrough.SolOut"
SolCallback(2) =    "SolRequestAutoPlunger"
SolCallback(3)		= "SolDropReset"
SolCallback(5)		= "SetLamp 105,"  ' Slot Arm
SolCallback(7)      = "SolRouletteMotor"
SolCallback(8)		= "SetLamp 108,"
SolCallback(14)	    = "PlaySoundAt ""fx_plastichit"", sw39:SetLamp 114,"
SolCallback(13)     = "debug.print ""magnet"" '"
'SolCallBack(17)		= "vpmSolSound ""sling"","
'SolCallBack(18)		= "vpmSolSound ""sling"","
'SolCallBack(19)		= "vpmSolSound ""sling"","
'SolCallBack(20)		= "vpmSolSound ""sling"","
SolCallBack(21) 	= "SetLamp 121,"
SolCallBack(22)		= "SolPopEntry"
SolCallBack(23)		= "SolLockRelease"
SolCallback(24)	= "vpmSolSound ""Knocker"","
SolCallBack(25)		=  "SolLockDiverter"
SolCallback(26)     = "SetLamp 126,"
SolCallBack(30) 	= "SetLamp 130,"
SolCallBack(31)	    = "SetLamp 131,"
SolCallBack(32)  	= "SetLamp 132,"

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Dim AutoPlungeRequest:AutoPlungeRequest = 0

sub SolRequestAutoPlunger(Enabled)
	if Enabled then AutoPlungeRequest = 1
end sub

sub DoAutoPlunge
	vpmSolAutoPlungeS AutoPlunger, SoundFX(SSolenoidOn, DOFContactors), 8, True
	vpmtimer.addtimer 500,"vpmSolAutoPlungeS AutoPlunger, """", 8, False '"
	AutoPlungeRequest = 0
end sub

Sub SolRampMotor(Lvl)
    ' NOTE: This ramp is primarily driven by a mech, we are also using the fading lamp routine to drive the motor sounds (wind up, wind down).
	if Lvl > 0 then
		PlaySound SoundFX("highpitchmotor", DOFGear), -1, lvl * .02, AudioPan(sw35), 0, (1-lvl) * -10000, 1, 0, AudioFade(sw35)
	Else
		StopSound "highpitchmotor"
	end if
End Sub

Sub SolRouletteMotor(Enabled)
	SetLamp 107,Enabled
	RouletteFloor.TimerEnabled = Enabled
End Sub

Sub SolRouletteMotorSound(Lvl)
	if Lvl > 0 then
		PlaySound SoundFX("machinemotor", DOFGear), -1, lvl * .0005, AudioPan(sw53), 0, (1-lvl) * -10000, 1, 0, AudioFade(sw53)
	Else
		StopSound "machinemotor"
	end if
End Sub


Sub SolSlotArm(Lvl)
	SlotArm.RotZ = 90- ( 90 * Lvl)
End Sub

Sub SolRampDiverter(Lvl)
	CenterDiverter.RotZ = 26 * Lvl
	if Lvl > .5 then
		RampCenterLeft.Collidable = False
		RampCenterRight.Collidable = True
	Else
		RampCenterLeft.Collidable = True
		RampCenterRight.Collidable = False
	End if
End Sub

Sub SolPopEntry(Enabled)
	SetLamp 122, Enabled
	PopEntryPost.IsDropped=Enabled
	if Enabled then
		PlaySoundAt "fx_solenoid", sw32
	Else
		PlaySoundAt "fx_solenoidoff", sw32
	End if
End Sub

Sub SolPopEntryPin(lvl)
	popentrypostp.Z = -50 - lvl * 50
End Sub

Sub SolLockDiverter(Enabled)
	if Enabled then
		PlaySoundAt "fx_solenoid", l33
	Else
		PlaySoundAt "fx_solenoidoff", l33
	End if
	DiverterWall.IsDropped = not Enabled
End Sub

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySound SoundFX("fx_flipperup",DOFContactors), 0, 1, -0.1, 0.25
        LeftFlipper.RotateToEnd
    Else
        PlaySound SoundFX("fx_flipperdown",DOFContactors), 0, 1, -0.1, 0.25
        LeftFlipper.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySound SoundFX("fx_flipperup",DOFContactors), 0, 1, 0.1, 0.25
        RightFlipper.RotateToEnd
   Else
        PlaySound SoundFX("fx_flipperdown",DOFContactors), 0, 1, 0.1, 0.25
        RightFlipper.RotateToStart
    End If
End Sub

Sub SolKnocker(Enabled)
	If Enabled Then PlaySound SoundFX("Knocker",DOFKnocker)
End Sub

Sub SolDropReset(Enabled)
	If Enabled Then
		dtDrop.DropSol_On
		PlaySoundAt "fx_resetdrop", L15
		Controller.Switch(21)=0
		Controller.Switch(22)=0
		Controller.Switch(23)=0
	End If
End Sub

Sub SolLockRelease(Enabled)
	SetLamp 123, Enabled
	if Enabled Then
		PlaySoundAt "fx_solenoid", LockKick1
		bsVLock.SolExit Enabled
	Else
		bsVLock.SolExit False
	End if
End sub

Sub SolLockReleasePin(lvl)
	balllock.Z = 0 - lvl * 50
End Sub



'************************************************
' Switches
'************************************************


Sub sw9_Hit:PlaySoundAtBall "fx_target":vpmTimer.PulseSw 9:End Sub

Sub sw10_Hit:PlaySoundAtBall "fx_target":vpmTimer.PulseSw 10:End Sub

Sub sw16_Hit
	Controller.Switch(16)=1:
    if AutoQA=false Then Exit Sub
	vpmtimer.addtimer 100,"Table1_KeyDown(PlungerKey) '"
    vpmtimer.addtimer 1000+RND(400),"Table1_KeyUp(PlungerKey) '"
End Sub

Sub sw16_unHit:Controller.Switch(16)=0:End Sub

Sub sw17_Hit:dtDrop.Hit 1:PlaySoundAtBall "fx_target":End Sub
Sub sw18_Hit:dtDrop.Hit 2:PlaySoundAtBall "fx_target":End Sub
Sub sw19_Hit:dtDrop.Hit 3:PlaySoundAtBall "fx_target":End Sub
Sub sw20_Hit:dtDrop.Hit 4:PlaySoundAtBall "fx_target":End Sub

Sub sw21_Hit:vpmTimer.PulseSw 21:PlaySoundAtBall "fx_target":End Sub
Sub sw22_Hit:vpmTimer.PulseSw 22:PlaySoundAtBall "fx_target":End Sub
Sub sw23_Hit:vpmTimer.PulseSw 23:PlaySoundAtBall "fx_target":End Sub

Sub Spinner_Spin:vpmTimer.PulseSw 24:PlaySoundAt "fx_spinner", Spinner:End Sub

Sub sw28_Hit:vpmTimer.PulseSw 28:PlaySoundAtBall "fx_target":End Sub

Sub sw29_Hit:vpmTimer.PulseSw 29:PlaySoundAtBall "fx_target":End Sub

Sub sw30_Hit:vpmTimer.PulseSw 30:PlaySoundAtBall "fx_target":End Sub

Sub sw31_Hit:vpmTimer.PulseSw 31:PlaySoundAtBall "fx_target":End Sub

Sub sw32_Hit:Controller.Switch(32)=1:End Sub
Sub sw32_unHit:Controller.Switch(32)=0:End Sub

Sub SlotBack_Hit:vpmTimer.PulseSw 33: End Sub				'
Sub SlotBlock_Hit:vpmTimer.PulseSw 34:PlaySoundAtBall "fx_metalhit2": End Sub
Sub SlotLegsP_Hit:PlaySoundAtBall "fx_metalhit2": End Sub

Sub sw35_Hit
	Controller.Switch(35)=1
	if ActiveBall.VelZ < -1 then PlaySoundAtBall "fx_balldrop"
End Sub

Sub sw35_unHit:Controller.Switch(35)=0:End Sub

Sub sw36_Hit:vpmTimer.PulseSw 36:End Sub				'36
Sub sw37_Hit:Controller.Switch(37)=1:End Sub				'37 Center Left RampEnter
Sub sw37_unHit:Controller.Switch(37)=0:End Sub
Sub sw39_Hit:Controller.Switch(39)=1:End Sub				'39 Left Orbit Gate
Sub sw39_unHit:Controller.Switch(39)=0:End Sub
Sub sw40_Hit:vpmTimer.pulseSw 40:End Sub				'40 Left Ramp Exit
Sub Bumper1_Hit:vpmTimer.PulseSw 49:PlaySoundAtBumper SoundFX("fx_bumper1", DOFContactors), Bumper1:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 50:PlaySoundAtBumper SoundFX("fx_bumper2", DOFContactors), Bumper2:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 51:PlaySoundAtBumper SoundFX("fx_bumper3", DOFContactors), Bumper3:End Sub
Sub sw53_Hit:vpmTimer.PulseSw 53:End Sub				'53 Roulette Spin

'Outlane_Inlane
Sub sw60_Hit: PlaySoundAtBall "fx_sensor": Controller.Switch(60)=1:End Sub
Sub sw60_UnHit: Controller.Switch(60)=0:End Sub
Sub sw61_Hit:	Controller.Switch(61) = 1:	PlaySoundAtBall "fx_sensor":End Sub
Sub sw61_UnHit:	Controller.Switch(61) = 0: End Sub
Sub sw57_Hit:  Controller.Switch(57) = 1: 	PlaySoundAtBall "fx_sensor": End Sub
Sub sw57_UnHit: Controller.Switch(57) = 0: End Sub
Sub sw58_Hit: Controller.Switch(58) = 1: 	PlaySoundAtBall "fx_sensor": End Sub
Sub sw58_UnHit:	Controller.Switch(58) = 0: End Sub

Sub sw45_Hit:	Controller.Switch(45) = 1: PlaySoundAtBall "fx_sensor": End Sub
Sub sw45_UnHit: Controller.Switch(45) = 0: End Sub

Sub sw46_Hit(): debug.print "sw46": vpmTimer.PulseSw 46: PlaySoundAtBall "fx_sensor":End Sub


Sub LeftRampStep_UnHit(): PlaySoundAtBall "fx_balldrop": End Sub


' Cheesy but effective hacks.   Move ball behind slot if it ends up trapped near it.

Sub BallEscape1_Hit
	ActiveBall.x = L132a.x
	ActiveBall.y = L132a.y
	ActiveBall.z = 50
End Sub

Sub BallEscape2_Hit: BallEscape1_Hit: End Sub
Sub BallEscape3_Hit: BallEscape1_Hit: End Sub


' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

' Similar subroutines that are less complicated to use (e.g. simply use standard parameters for the PlaySound call)
Sub PlaySoundAt(soundname, tableobj)
    PlaySound soundname, 1, 1, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBallVol(soundname, voladd)
	PlaySound soundname, 1, Vol(ActiveBall) + voladd, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtBumper(sound, tableobj)
	PlaySound sound, 1, 1, AudioPan(tableobj), 0,0,1, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname)
    PlaySoundAt soundname, ActiveBall
End Sub

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


Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

Sub RandomSoundMetal()
    PlaySoundAtBallVol "fx_metal_hit_" & Int(Rnd*3)+1, .2
End Sub

Dim NextOrbitHit:NextOrbitHit = 0

Sub RandomSoundFlipper()
	if Timer > NextOrbitHit then
		NextOrbitHit = Timer + .1 + (Rnd * .2)
		PlaySoundAtBallVol "fx_flip_hit_" & Int(Rnd*3)+1, .2
	end if
End Sub

Sub LeftFlipper_Collide(parm)
    RandomSoundFlipper()
End Sub

Sub RightFlipper_Collide(parm)
    RandomSoundFlipper()
End Sub

Sub Rubbers_Hit(idx)
    PlaySoundAtBallVol "fx_rubber_hit_" & Int(Rnd*3)+1, .2
End Sub

Sub Posts_Hit(idx)
    PlaySoundAtBallVol "fx_rubber_hit_" & Int(Rnd*3)+1, .2
End Sub


Sub PlasticRamps_Hit(idx)
	if BallVel(ActiveBall) > .3 and Timer > NextOrbitHit then
		RandomBump 2, Pitch(ActiveBall)
		' Schedule the next possible sound time.  This prevents it from rapid-firing noises too much.
		' Lowering these numbers allow more closely-spaced clunks.
		NextOrbitHit = Timer + .1 + (Rnd * .2)
	end if
End Sub

Sub MetalRamps_Hit(idx)
	if BallVel(ActiveBall) > .3 and Timer > NextOrbitHit then
		RandomBump 2, 20000 'Increased pitch to simulate metal wall
		' Schedule the next possible sound time.  This prevents it from rapid-firing noises too much.
		' Lowering these numbers allow more closely-spaced clunks.
		NextOrbitHit = Timer + .2 + (Rnd * .2)
	end if
End Sub


'' Requires rampbump1 to 7 in Sound Manager
Sub RandomBump(voladj, freq)
	dim BumpSnd:BumpSnd= "rampbump" & CStr(Int(Rnd*7)+1)
		PlaySound BumpSnd, 0, Vol(ActiveBall)+voladj, AudioPan(ActiveBall), 0, freq, 0, 1, AudioFade(ActiveBall)
End Sub



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
    Dim BOT, b, ballpitch, BallInRoulette
    BOT = GetBalls

    ' stop the sound of deleted balls
    For b = UBound(BOT) + 1 to tnob
        rolling(b) = False
        StopSound("fx_ballrolling" & b)
    Next

    ' exit the sub if no balls on the table
    If UBound(BOT) = -1 Then Exit Sub

	' Track whether any ball is in the slot ramp area.   Delay physics movements until balls are away from there,
    ' to prevent balls from getting stuck inside.
	BallInRampArea = False
	BallInRoulette = 0

    ' play the rolling sound for each ball
    For b = 0 to UBound(BOT)
		if BOT(b).x > 153 and BOT(B).y > 421 and BOT(b).x < 350 and BOT(b).y < 765 Then
			BallInRampArea = True
		end if
		if BOT(b).z> 70 and BOT(B).x >  569 and BOT(B).y <  374 then
			BallInRoulette = BallInRoulette + 1
		end if
		if BOT(B).z < 0 then BallEscape1_Hit ' Ball falling into space..

        If BallVel(BOT(b) ) > 1 Then
            If BOT(b).z < 30 Then
                ballpitch = Pitch(BOT(b) )
            Else
                ballpitch = Pitch(BOT(b) ) * 100
				' SLEAZE!
				if BOT(b).z > 150 and BOT(b).VelZ > 0 then BOT(b).VelZ = 0
            End If
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), AudioPan(BOT(b) ), 0, ballpitch, 1, 0, AudioFade(BOT(b))
			' play ball drop sounds
			If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
				PlaySound "fx_balldrop", 0, ABS(BOT(b).velz)/17, AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
			End If

        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If
    Next
	if AutoPlungeRequest and BallInRoulette < MaxBallsRoulette then DoAutoPlunge
	SyncSlotRamp
End Sub

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

'******************************
' Diverse Collection Hit Sounds
'******************************

Sub aBumpers_Hit (idx): PlaySound SoundFX("fx_bumper", DOFContactors), 0, 1, Audiopan(ActiveBall), 0, 0, 0, 0, AudioFade(ActiveBall): End Sub
Sub aRollovers_Hit(idx):PlaySound "fx_sensor", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRollovers_Hit(idx):PlaySound "fx_sensor", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aMetals_Hit(idx):PlaySound "fx_MetalHit", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Bands_Hit(idx):PlaySound "fx_rubber", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubbers_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
	Sub aRubber_Posts_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
	Sub aRubber_Pins_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
	Sub aYellowPins_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aWoods_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub

' Ramp Sounds
Sub RHelp1_Hit()
    StopSound "fx_metalrolling"
    PlaySound "fx_ballrampdrop", 0, 1, pan(ActiveBall)
End Sub

Sub RHelp2_Hit()
    StopSound "fx_metalrolling"
    PlaySound "fx_ballrampdrop", 0, 1, pan(ActiveBall)
End Sub

'*************************************************************
' Lamp routines
' SetLamp xx, 0 is Off. SetLamp xx,1 is On
'*************************************************************
Dim LampState(200), FadingLevel(200), FadingState(200)
Dim FlashState(200), FlashLevel(200)
Dim FlashSpeedUp, FlashSpeedDown

LampTimer.Interval = -1
LampTimer.Enabled = 1

Dim MiniDMDMap(15,7)

Sub GIAdjust(lvl)
	if (lvl > .5) then
		plastics.image="plastics1"
		plastics.blenddisablelighting = (lvl - .5) * 2
	Else
		plastics.image="plastics1off"
		plastics.blenddisablelighting = lvl * 2 + .5
	end if

    if (lvl > .2) then
		sidewalls.image="Side_Walls_GI_On"
		sidewalls.blenddisablelighting = lvl
	Else
		sidewalls.image="Side_Walls_GI_Off"
		sidewalls.blenddisablelighting = lvl * 2
	end if

	if (lvl > 0) then
		smallsidewalls.image="metal_walls_GI_OnTex"
		smallsidewalls.blenddisablelighting = lvl
	Else
		smallsidewalls.image="metal_walls_GI_OffTex"
		smallsidewalls.blenddisablelighting = lvl * 6
	end if
end sub

Sub FlasherP(obj, lvl)
	dim lvlstep:lvlstep = CInt(lvl * 5)
	obj.image="Flasher_Dome_Stage_" & lvlstep
	obj.blenddisablelighting = lvl
end sub


Sub InitLamps

	'Adjust fading speeds (1 / full MS fading time)
	dim x
	for x = 0 to 140 : FadeLights.FadeSpeedUp(x) = 1/10 : FadeLights.FadeSpeedDown(x) = 1/25 : next
	for x = 117 to 120 : FadeLights.FadeSpeedUp(x) = 1/5 : FadeLights.FadeSpeedDown(x) = 1/5 : next

	'Lamp Assignments
	dim lightobj
	For each lightobj in TableLights
		' Asumptions: Light is named "Lxxy" where x is the lamp number, and y is optionally a,b,c for multiple on the same id
		dim arr, id, lname
		lname = mid(lightobj.Name,2)
		if not IsNumeric(Right(lname, 1)) then lname = Left(lname, len(lname)-1)
		id = cInt(lname)
	    if TypeName(FadeLights.Obj(id)) = "NullFadingObject" then
			arr = array(lightobj)
		Else
			arr = FadeLights.Obj(id)
			ReDim Preserve arr(UBound(arr) + 1)
			set arr(UBound(arr)) = lightobj
		end if
		FadeLights.Obj(id) = arr
	next

	' Flasher extras
	FadeLights.Callback(132) = "FlasherP L132p,"
	FadeLights.Callback(130) = "FlasherP L130p,"
	FadeLights.Callback(121) = "FlasherP L121p,"
	FadeLights.Callback(108) = "FlasherP L108p,"


	dim giarr
	For each lightobj in GI
	    if IsEmpty(giarr) then
			giarr = array(lightobj)
		Else
			ReDim Preserve giarr(UBound(giarr) + 1)
			set giarr(UBound(giarr)) = lightobj
		end if
	next

	FadeLights.Callback(99) = "GIAdjust"
	FadeLights.Obj(99) = GiArr

    if DesktopMode Then
		For each lightobj in BGLights
			lightobj.state = 1
			Set FadeLights.Obj(lightobj.TimerInterval) = lightobj
		next
	end if

	dim led

	'Solenoid assignments

	FadeLights.Callback(105) = "SolSlotArm"

	FadeLights.Callback(107) = "SolRouletteMotorSound"
	FadeLights.FadeSpeedUp(107) = 1/5 : FadeLights.FadeSpeedDown(107) = 1/5

	FadeLights.Callback(126) = "SolRampMotor"
	FadeLights.FadeSpeedUp(126) = 1/5 : FadeLights.FadeSpeedDown(126) = 1/5

	FadeLights.Callback(114) = "SolRampDiverter"

	FadeLights.Callback(122) = "SolPopEntryPin"
	FadeLights.FadeSpeedUp(122) = 1/5 : FadeLights.FadeSpeedDown(122) = 1/5

	FadeLights.Callback(123) = "SolLockReleasePin"
	FadeLights.FadeSpeedUp(123) = 1/5 : FadeLights.FadeSpeedDown(123) = 1/5


	' Drop target shadows

	for x = 91 to 94 : FadeLights.FadeSpeedUp(x) = 1/20 : FadeLights.FadeSpeedDown(x) = 1/20 : next

	Set FadeLights.Obj(91) = drop1
	Set FadeLights.Obj(92) = drop2
	Set FadeLights.Obj(93) = drop3
	Set FadeLights.Obj(94) = drop4

	'MiniDMD setup
	For Each led in MiniDMD
		Set MiniDMDMap(CInt(Mid(led.name, 2))-1, (asc("g") - asc(led.name)))= led
	next

End Sub

Sub OneMsTimer_Timer()
	FadeLights.Update1
End Sub

Sub LampTimer_Timer()
    Dim chgLamp, num, chg, ii
	chgLamp = Controller.ChangedLamps
	If Not IsEmpty(chgLamp) Then
		For ii = 0 To UBound(chgLamp)
			 SetLamp ChgLamp(ii,0), ChgLamp(ii,1)
		Next
	End If

	Dim ChgLED,jj,stat,obj
	ChgLED=Controller.ChangedLEDs(&H00000000, &Hffffffff)
	If Not IsEmpty(ChgLED) Then
		For ii = 0 To UBound(chgLED)
			num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
			for jj = 0 to 7
				If chg And 1 Then
					if Stat and 1 then
						MiniDMDMap(num, jj).Material = "DMDOn"
						MiniDMDMap(num, jj).BlendDisableLighting = 1
					Else
						MiniDMDMap(num, jj).Material = "DMDOff"
						MiniDMDMap(num, jj).BlendDisableLighting = 0
					end if
				end if
				chg = chg\4:stat = stat\4
			Next
		Next
	End If
	LeftFlipperP.ObjRotZ = LeftFlipper.CurrentAngle
	RightFlipperP.ObjRotZ = RightFlipper.CurrentAngle
	Flipperlsh.RotZ = LeftFlipper.CurrentAngle
	Flipperrsh.RotZ = RightFlipper.CurrentAngle
	BallShadowUpdate
	UpdateTargetShadows
	FadeLights.Update
End Sub

Sub UpdateTargetShadows
	' If the upper GI is off, shadows should be off.
	if  FadeLights.state(99) = 0 then
		SetLamp 91, 0
		SetLamp 92, 0
		SetLamp 93, 0
		SetLamp 94, 0
	Else
		SetLamp 91, 1+Controller.Switch(17)
		SetLamp 92, 1+Controller.Switch(18)
		SetLamp 93, 1+Controller.Switch(19)
		SetLamp 94, 1+Controller.Switch(19)
	end if
end sub

' *********************************************************************
'						BALL SHADOW
' *********************************************************************
ReDim BallShadow(tnob-1)
InitBallShadow

Sub InitBallShadow
	Dim i: For i=0 to tnob-1
		ExecuteGlobal "Set BallShadow(" & i & ")=BallShadow" & (i+1) & ":"
	Next
End Sub

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
		' stuck ball hack!  ' 889.1185 168.7713 24.99723
		if BOT(b).x > 888 and BOT(b).y < 169 and BOT(b).z < 26 then BOT(b).x = 500
		If BOT(b).X < Table1.Width/2 Then
			BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/10)) + 10
		Else
			BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/10)) - 10
		End If
	    ballShadow(b).Y = BOT(b).Y + 20
		If BOT(b).Z > 20 Then
			BallShadow(b).visible = 1
		Else
			BallShadow(b).visible = 0
		End If
	Next
End Sub


Sub SetLamp(nr, value)
	' If the lamp state is not changing, just exit.
	if FadeLights.state(nr) = value then exit sub

    FadeLights.state(nr) = value
End Sub


Class NullFadingObject : Public Property Let IntensityScale(input) : : End Property : End Class	'todo do better

Class LampFader
	Public FadeSpeedDown(140), FadeSpeedUp(140)
	Private Lock(140), Loaded(140), OnOff(140)
	Public UseFunction
	Private cFilter
	Private UseCallback(140), cCallback(140)
	Public Lvl(140), Obj(140)

	Sub Class_Initialize()
		dim x : for x = 0 to uBound(OnOff) 	'Set up fade speeds
			if FadeSpeedDown(x) <= 0 then FadeSpeedDown(x) = 1/100	'fade speed down
			if FadeSpeedUp(x) <= 0 then FadeSpeedUp(x) = 1/80'Fade speed up
			UseFunction = False
			lvl(x) = 0
			OnOff(x) = False
			Lock(x) = True : Loaded(x) = False
		Next

		for x = 0 to uBound(OnOff) 		'clear out empty obj
			if IsEmpty(obj(x) ) then Set Obj(x) = NullFader' : Loaded(x) = True
		Next
	End Sub

	Public Property Get Locked(idx) : Locked = Lock(idx) : End Property		'debug.print Lampz.Locked(100)	'debug
	Public Property Get state(idx) : state = OnOff(idx) : end Property
	Public Property Let Filter(String) : Set cFilter = GetRef(String) : UseFunction = True : End Property
	Public Property Let Callback(idx, String) : cCallback(idx) = String : UseCallBack(idx) = True : End Property

	Public Property Let state(ByVal idx, input) 'Major update path
		input = cBool(input)
		if OnOff(idx) = Input then : Exit Property : End If	'discard redundant updates
		OnOff(idx) = input
		Lock(idx) = False
		Loaded(idx) = False
	End Property

	Public Sub TurnOnStates()	'If obj contains any light objects, set their states to 1 (Fading is our job!)
		dim debugstr
		dim idx : for idx = 0 to uBound(obj)
			if IsArray(obj(idx)) then
				dim x, tmp : tmp = obj(idx) 'set tmp to array in order to access it
				for x = 0 to uBound(tmp)
					if typename(tmp(x)) = "Light" then DisableState tmp(x) : debugstr = debugstr & tmp(x).name & " state'd" & vbnewline

				Next
			Else
				if typename(obj(idx)) = "Light" then DisableState obj(idx) : debugstr = debugstr & obj(idx).name & " state'd (not array)" & vbnewline

			end if
		Next
	End Sub
	Private Sub DisableState(ByRef aObj) : aObj.FadeSpeedUp = 1000 : aObj.State = 1 : End Sub	'turn state to 1

	Public Sub Update1()	 'Handle all boolean numeric fading. If done fading, Lock(True). Update on a '1' interval Timer!
		dim x : for x = 0 to uBound(OnOff)
			if not Lock(x) then 'and not Loaded(x) then
				if OnOff(x) then 'Fade Up
					Lvl(x) = Lvl(x) + FadeSpeedUp(x)
					if Lvl(x) > 1 then Lvl(x) = 1 : Lock(x) = True
				elseif Not OnOff(x) then 'fade down
					Lvl(x) = Lvl(x) - FadeSpeedDown(x)
					if Lvl(x) < 0 then Lvl(x) = 0 : Lock(x) = True
				end if
			end if
		Next
	End Sub


	Public Sub Update()	'Handle object updates. Update on a -1 Timer! If done fading, loaded(x) = True
		dim x,xx : for x = 0 to uBound(OnOff)
			if not Loaded(x) then
				if IsArray(obj(x) ) Then	'if array
					If UseFunction then
						for each xx in obj(x) : xx.IntensityScale = cFilter(Lvl(x)) : Next
					Else
						for each xx in obj(x) : xx.IntensityScale = Lvl(x) : Next
					End If
				else						'if single lamp or flasher
					If UseFunction then
						obj(x).Intensityscale = cFilter(Lvl(x))
					Else
						obj(x).Intensityscale = Lvl(x)
					End If
				end if
				' Sleazy hack for regional decimal point problem
				If UseCallBack(x) then execute cCallback(x) & " CSng(" & CInt(10000 * Lvl(x)) & " / 10000)"	'Callback
				If Lock(x) Then
					if Lvl(x) = 1 or Lvl(x) = 0 then Loaded(x) = True	'finished fading
				end if
			end if
		Next
	End Sub
End Class




Class cvpmMyMech
	Public Sol1, Sol2, MType, Length, Steps, Acc, Ret
	Private mMechNo, mNextSw, mSw(), mLastPos, mLastSpeed, mCallback

	Private Sub Class_Initialize
		ReDim mSw(10)
		gNextMechNo = gNextMechNo + 1 : mMechNo = gNextMechNo : mNextSw = 0 : mLastPos = 0 : mLastSpeed = 0
		MType = 0 : Length = 0 : Steps = 0 : Acc = 0 : Ret = 0 : vpmTimer.addResetObj Me
	End Sub

	Public Sub AddSw(aSwNo, aStart, aEnd)
		mSw(mNextSw) = Array(aSwNo, aStart, aEnd, 0)
		mNextSw = mNextSw + 1
	End Sub

	Public Sub AddPulseSwNew(aSwNo, aInterval, aStart, aEnd)
		If Controller.Version >= "01200000" Then
			mSw(mNextSw) = Array(aSwNo, aStart, aEnd, aInterval)
		Else
			mSw(mNextSw) = Array(aSwNo, -aInterval, aEnd - aStart + 1, 0)
		End If
		mNextSw = mNextSw + 1
	End Sub

	Public Sub Start
		Dim sw, ii
		With Controller
			.Mech(1) = Sol1 : .Mech(2) = Sol2 : .Mech(3) = Length
			.Mech(4) = Steps : .Mech(5) = MType : .Mech(6) = Acc : .Mech(7) = Ret
			ii = 10
			For Each sw In mSw
				If IsArray(sw) Then
					.Mech(ii) = sw(0) : .Mech(ii+1) = sw(1)
					.Mech(ii+2) = sw(2) : .Mech(ii+3) = sw(3)
					ii = ii + 10
				End If
			Next
			.Mech(0) = mMechNo
		End With
		If IsObject(mCallback) Then mCallBack 0, 0, 0 : mLastPos = 0 : vpmTimer.EnableUpdate Me, True, True  ' <------- All for this.
	End Sub

	Public Property Get Position : Position = Controller.GetMech(mMechNo) : End Property
	Public Property Get Speed    : Speed = Controller.GetMech(-mMechNo)   : End Property
	Public Property Let Callback(aCallBack) : Set mCallback = aCallBack : End Property

	Public Sub Update
		Dim currPos, speed
		currPos = Controller.GetMech(mMechNo)
		speed = Controller.GetMech(-mMechNo)
		If currPos < 0 Or (mLastPos = currPos And mLastSpeed = speed) Then Exit Sub
		mCallBack currPos, speed, mLastPos : mLastPos = currPos : mLastSpeed = speed
	End Sub

	Public Sub Reset : Start : End Sub

End Class

'*****************************
' AUTO TESTING
' by:NailBuster
' Global variable "AutoQA" below will switch all this on/off during testing.
'
'*****************************
' NailBusters AutoQA Code and triggers..
' this to do for ROM based:  timeout on keydown.  if 30 seconds, then assume game is over and you add coins/start game key.
' add a timer called AutoQAStartGame.  you can run every 10000 interval.

Dim QACoinStartSec:QACoinStartSec=60   'timeout seconds for AutoCoinStartSec
Dim QANumberOfCoins:QANumberOfCoins=4 'number of coins to add for each start
Dim QASecondsDiff

Dim QALastFlipperTime:QALastFlipperTime=Now()
Dim AutoFlipperLeft:AutoFlipperLeft=false
Dim AutoFlipperRight:AutoFlipperRight=false



Sub AutoQAStartGame_Timer()                 'this is a timeout when sitting in attract with no flipper presses for 60 seconds, then add coins and start game.
 if AutoQA=false Then Exit Sub

 QASecondsDiff = DateDiff("s",QALastFlipperTime,NOW())

 if QASecondsDiff>QACoinStartSec Then

    'simulate quarters and start game keys
    Dim fx : fx=0
    Dim keydelay : keydelay=100
	Do While fx<QANumberOfCoins
		vpmtimer.addtimer keydelay,"Table1_KeyDown(keyInsertCoin1) '"
        vpmtimer.addtimer keydelay+200,"Table1_KeyUp(keyInsertCoin1) '"
        keydelay=keydelay+500
		fx=fx+1
	Loop
    vpmtimer.addtimer keydelay,"Table1_KeyDown(StartGameKey) '"
    vpmtimer.addtimer keydelay+200,"Table1_KeyUp(StartGameKey) '"
    QALastFlipperTime=Now()
	AutoFlipperLeft=false
    AutoFlipperRight=false
 End if

 if QASecondsDiff>30 Then   'safety of stuck up flipers.
   AutoFlipperLeft=false
   AutoFlipperRight=false
  End if
End Sub



Sub FlipperUP(which)  'which=1 left 2 right
QALastFlipperTime=Now()
if which=1 Then
   Table1_KeyDown(LeftFlipperKey)
   vpmtimer.addtimer 200+Rnd(200),"Table1_KeyUP(LeftFlipperKey):AutoFlipperLeft=false  '"
Else
   Table1_KeyDown(RightFlipperKey)
   vpmtimer.addtimer 200+Rnd(200),"Table1_KeyUP(RightFlipperKey):AutoFlipperRight=false  '"
end If

End Sub


Sub AutoQATriggerLeft_Hit()
	if AutoQA And AutoFlipperLeft=false then vpmtimer.addtimer 20+Rnd(50),"FlipperUP(1) '"
    AutoFlipperLeft=true
End Sub

Sub AutoQATriggerRight_Hit()
	if AutoQA and AutoFlipperRight=false then vpmtimer.addtimer 20+Rnd(20),"FlipperUP(2) '"
    AutoFlipperRight=true
End Sub


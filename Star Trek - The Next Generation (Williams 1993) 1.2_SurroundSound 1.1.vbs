'*********************************************************************************************************************************************************************
'*********************************************************************************************************************************************************************
'*********************************************************************************************************************************************************************
'*********************************************************************************************************************************************************************
'                                                     #####,  #######,  ,#####  ####### ;#####, ####### ######' ##, ###
'                                                    ######  :#######   ######  ####### ###### ,######, ###### ,## ###
'                                                    ###       .##         ### ;##  ##`        ##' ,## ###     ##'###
'                                                    ###:      ###      ;; ##  ##+;###   ;;    ##;;##' ##      #####
'                                                     ####     ##      ## `## .######   +##   ####### '###### #####.
'                                                       ##,   ###     ##. ### ##+,###   ##.   ##,,##, ######. ##`##
'                                                       ##.   ##`    ;#;  ##. ##  ##+  :##   '##  ## .##     '## ##`
'                                                 ,#######   :##    .####### ###  ##   ##:   ##. '## ######' ##. ##+
'                                                 #######`   ##,    ######## ##` ###  .##   ,##  ##. ###### ,##  ,##
'                                                `#####:    .##    ########,;##  ##`  ###   ##; ,## ####### ##;   ##
'
'            ###### :##  .## ######'     '######: ######' `##  ###; #######,     ####### :###### #######  ###### `######`  ,#####  #######, ##. ####### '######:
'           `###### ##:  ##+ ######      #######, ######   ## #### :#######     ######## ######,`####### +###### #######   ###### :####### :## :####### #######,
'                  .##   ## ###              .## ###       #####:    .##                .##     ###  ### ##.     ##  ###      ###   .##    ##: ##;  ##'.##  .##
'            :;;   ###;;### ##          ;;:  ### ##        '###:     ###       :;; ;;;; ###     ##   ##`,##     ###;'##    ;; ##    ###   .## `##   ## ##'  ###
'            ##,   #######`'######      ##   ## '######    ###:      ##        ##. ###. ###### +##  ;## ######: ######;   ## `##    ##    ### ###  ### ##   ##
'           .##   ###,,+## ######.     ###  ### ######.   ####      ###       ,## .### ####### ##.  ##,`###### :##,:##   ##. ###   ###    ##  ##   ## ###  ###
'           ##+   ##.  ##,.##          ##   ##`.##       #####:     ##`       ##;  ##; ##`    :##  .## ###     ##: :##  ;#;  ##.   ##`   ### ###  '## ##   ##`
'           ##   :##  .## ######'     '##  :## ######' ####.'##    :##       `####### :###### ##;  ##' ###### `##  ##: .#######   :##    ##` #######.+##  :##
'          ###   ##:  ##' ######      ##.  ##: ###### ;###.  ##    ##,       ######## ######,.##   ## +###### ### `##  ########   ##,   ;## ,####### ##.  ##:
'          ##`  .##   ## #######     ,##  .## ####### ###    ##   .##        '#####' .###### ###  ### ######` ##  ### ########,  .##    ##:  ###### ,##  .##
'*********************************************************************************************************************************************************************
'*********************************************************************************************************************************************************************
'*********************************************************************************************************************************************************************
'*********************************************************************************************************************************************************************

'*********************************************************************
'*********************************************************************
'1993 Star Trek - The Next Generation by Williams Pinball
'*********************************************************************
'*********************************************************************
'Concept by:		Steve Ritchie
'Design by:			Steve Ritchie, Dwight Sullivan, Greg Freres
'Art by:			Greg Freres
'Dots/Animation by:	Scott Slomiany, Eugene Geer
'Mechanics by:		Carl Biagi
'Music by:			Dan Forden
'Sound by:			Dan Forden
'Software by:		Dwight Sullivan, Matt Coriale
'*********************************************************************
'*********************************************************************
'recreated for Visual Pinball by Knorr and Clark Kent
'*********************************************************************
'*********************************************************************
'V1.2				Added more lights
'					updated physics
'					updated ballshadow and ballrolling script from Ninuzzu --- Thanks!
'					reduced polycount
'					added new sounds
'V1.1				Quick Fix for the SpiralRamp
'V1.0				First Release for Visual Pinball 10.2

Option Explicit
Randomize

' Thalamus 2018-07-24
' Table has its own "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' No special SSF tweaks yet.

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Dim DesktopMode
DesktopMode = Table1.ShowDT
Dim UseVPMDMD
UseVPMDMD = DesktopMode
'UseVPMDMD = 1
Const UseVPMModSol = 1

'**********
'***Mods***
'**********

'1 for ON, 0 for OFF

InlanesMod = 1				'with this mod the inlanes are about 1cm/0,39inch longer
BorgMod = 1					'changes the Original Model to a Custom Model

RGBMod = 0					'RGB
	RGBBumpers = 1
	RGBArrows = 0

LaserMod = 0				'Activate Lasers on the Cannons
	LaserColor = 1			'1 = RED, 2 = GREEN, 3 = BLUE
	LaserType = 1			'1 = Fast, turns off after the ball is shoot, 2 = Slow, turns off after the cannon is back initial_pos


'********************
'Standard definitions
'********************

LoadVPM "01560000", "wpc.VBS", 3.36

'***ROM***ROM***ROM***ROM***
Const cGameName = "sttng_l7"
'***************************

Const UseSolenoids = 2
Const UseLamps = 0
Const HandleMech = 0
Const SSolenoidOn = "SolOn"
Const SSolenoidOff = "SolOff"
Const SFlipperOn = "FlipperUp"
Const SFlipperOff = "FlipperDown"
Const SCoin = "Coin4"
Dim bsTrough,TopDrop, LeftCannonMech, RightCannonMech, BorgLock
Dim KB, CP, UnderDiverterTop1, UnderDiverterBottom1
Dim InlanesMod, BorgMod, FlipperMod, LaserMod, LaserColor, LaserType, RGBMod, RGBBumpers, RGBArrows


Set MotorCallback = GetRef("RealTimeUpdates")
Set GiCallback2 = GetRef("UpdateGI")
BSize = 25.5
'BMass = ((BSize*2)^3)/125000
BMass = 1.5


'************
' Table Init
'************

Sub table1_Init
    vpmInit Me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
        .Games(cGameName).Settings.Value("rol") = 0
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = DesktopMode
        '.SetDisplayPosition 0, 0, GetPlayerHWnd 'uncomment this line if you don't see the DMD
        On Error Resume Next
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
        On Error Goto 0
        .Switch(22) = 1 'close coin door
        .Switch(24) = 0 'and keep it close
        PinMAMETimer.Interval = PinMAMEInterval
        PinMAMETimer.Enabled = true
        vpmNudge.TiltSwitch = 14
        vpmNudge.Sensitivity = 2
		vpmNudge.TiltObj = Array(LeftJetBumper, RightJetBumper, BottomJetBumper, SlingShotLeft, SlingShotRight)
    End With

    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0, 66, 65, 64, 63, 62, 61, 0
        .InitKick BallRelease, 80, 10
        .InitExitSnd SoundFX("BallRelease", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .Balls = 6
    End With

     Set TopDrop = New cvpmDropTarget
     With TopDrop
        .InitDrop sw57, 57
		.InitSnd SoundFX("Dropdown",DOFDropTargets), SoundFX("DropUp",DOFDropTargets)
     End With

     Set BorgLock = New cvpmBallStack
     With BorgLock
         .InitSw 0, 31, 0, 0, 0, 0, 0, 0
         .InitKick BorgKicker, 180, 24
         .InitExitSnd SoundFX("popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
'         .Balls = 0
     End With

'**********************
'***DesktopMode Init***
'**********************

If table1.ShowDt = False then
	Korpus.visible = 0
	Korpus1.visible = 1
	LeftSideRail.visible = 0
	RightSideRail.visible = 0
	Else
	Korpus.visible = 1
	Korpus1.visible = 0
	LeftSideRail.visible = 1
	RightSideRail.visible = 1
End if

'****************************
'***Switches/Diverter Init***
'****************************


Controller.Switch(127) = 1
Controller.Switch(122) = 1
Controller.Switch(125) = 1
Controller.Switch(126) = 1
DiverterFRG.isDropped = 1
DiverterFLG.isDropped = 0
sw57.isDropped = 1
sw41dr.isDropped = 1
sw35dr.isDropped = 1
sw36dr.isDropped = 1
sw37dr.isDropped = 1
sw37dr1.isDropped = 0
sw42dr.isDropped = 1
sw43dr.isDropped = 1
sw47dr.isDropped = 1

'**************
'***Mods Init***
'**************

If InlanesMod = 1 then
	leftinlanewall.collidable = 0
	rightinlanewall.collidable = 0
	inlanemetal.visible = 0
	leftinlanewallmod.collidable = 1
	rightinlanewallmod.collidable = 1
	inlanemetalmod.visible = 1
	Else
	leftinlanewall.collidable = 1
	rightinlanewall.collidable = 1
	inlanemetal.visible = 1
	leftinlanewallmod.collidable = 0
	rightinlanewallmod.collidable = 0
	inlanemetalmod.visible = 0
	End if


If BorgMod = 1 then
	borgshipcustom.visible = 1
	borgshipcustomledoff.visible = 1
	borgshipcustomwiresplines.visible = 1
	deltarampendborgcustom.visible = 1

	l78a.Intensity = 30
	l78b.Intensity = 30
	l78c.Intensity = 6
	l78d.Intensity = 30
	l78e.Intensity = 10
	l78a1.State = 1
	l78b1.State = 1
	l78c1.State = 1
	l78d1.State = 1
	l78e1.State = 1

	borgshiporiginal.visible = 0
	borgledoriginal.visible = 0
	deltarampendborg.visible = 0
	l78borga.Intensity = 0
	l78borgb.Intensity = 0
	l78borgc.Intensity = 0
	l78borgd.Intensity = 0
	l78borge.Intensity = 0
	l78borga1.State = 0
	l78borgb1.State = 0
	l78borgc1.State = 0
	l78borgd1.State = 0
	l78borge1.State = 0
	Else
	borgshipcustom.visible = 0
	borgshipcustomledoff.visible = 0
	borgshipcustomwiresplines.visible = 0
	deltarampendborgcustom.visible = 0

	l78a.Intensity = 0
	l78b.Intensity = 0
	l78c.Intensity = 0
	l78d.Intensity = 0
	l78e.Intensity = 0
	l78a1.State = 0
	l78b1.State = 0
	l78c1.State = 0
	l78d1.State = 0
	l78e1.State = 0

	borgshiporiginal.visible = 1
	borgledoriginal.visible = 1
	deltarampendborg.visible = 1
	l78borga.Intensity = 15
	l78borgb.Intensity = 50
	l78borgc.Intensity = 60
	l78borgd.Intensity = 30
	l78borge.Intensity = 50
	l78borga1.State = 1
	l78borgb1.State = 1
	l78borgc1.State = 1
	l78borgd1.State = 1
	l78borge1.State = 1
	End if

If RGBMod = 1 then RGBTimer.Enabled = 1
	If RGBBumpers = 1 And RGBMod = 1 then
		BumperCap1.Image = "bumpero_weis"
		BumperCap2.Image = "bumpero_weis"
		BumperCap3.Image = "bumpero_weis_half"
	End if


If LaserColor = 1 then LaserL.Material = "LaserRed"
If LaserColor = 1 then LaserL1.Material = "LaserRed1"
If LaserColor = 1 then LaserR.Material = "LaserRed"
If LaserColor = 1 then LaserR1.Material = "LaserRed1"

If LaserColor = 2 then LaserL.Material = "LaserGreen"
If LaserColor = 2 then LaserL1.Material = "LaserGreen1"
If LaserColor = 2 then LaserR.Material = "LaserGreen"
If LaserColor = 2 then LaserR1.Material = "LaserGreen1"

If LaserColor = 3 then LaserL.Material = "LaserBlue"
If LaserColor = 3 then LaserL1.Material = "LaserBlue1"
If LaserColor = 3 then LaserR.Material = "LaserBlue"
If LaserColor = 3 then LaserR1.Material = "LaserBlue1"
End Sub

'*********
'Solenoids
'*********

SolCallback(1) = "LeftCannonKicker"
SolCallback(2) = "RightCannonKicker"
SolCallback(3) = "UnderLeftGun"
SolCallback(4) = "UnderRightGun"
SolCallback(5) = "LeftLock"
SolCallback(6) = "AutoPlunge"
SolCallback(7) = "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallBack(8) = "Kickback"
SolCallBack(11) = "SolRelease"
SolCallBack(15) = "TopDiverter"
SolCallBack(16) = "BorgLock.SolOut"
SolCallback(17) = "LeftCannonMotor"
SolCallback(18) = "RightCannonMotor"
SolCallBack(51) = "UnderDiverterTop"
SolCallBack(52) = "UnderDiverterBottom"
SolCallBack(53) = "TopDrop.SolDropUp"
SolCallBack(54) = "TopDrop.SolDropDown"

'***Flasher***
SolModCallBack(20) = "Flash120"			'JetsFlasher
SolModCallBack(21) = "Flash121"			'RightPopperFlasher
SolModCallBack(22) = "Flash122"			'MiddleRampFlasher
SolModCallBack(23) = "Flash123"			'ShieldsFlasher
SolCallBack(24) = "SetLamp 124,"		'AutofireFlasher
SolModCallBack(25) = "Flash125"			'ExitUnGndFlasher
SolModCallBack(26) = "Flash126"			'RightBorgFlasher
SolModCallBack(27) = "Flash127"			'LeftBorgFlasher
SolModCallBack(28) = "Flash128" 		'CenterBorgFlasher
SolModCallBack(55) = "Flash141"			'RomulanFlasher
SolModCallBack(56) = "Flash142"		'RightRampFlasher


Sub SolRelease(Enabled)
    If Enabled And bsTrough.Balls> 0 Then
        vpmTimer.PulseSw 67
        bsTrough.ExitSol_On
    End If
End Sub

Sub Drain_Hit
    PlaySoundAtVol "Drain",Drain,.7
    bsTrough.AddBall Me
End Sub


Sub AutoPlunge(Enabled)
	if enabled then
		AutoPlunger.Kick 0, 145
		PlaySoundAt SoundFX("PlungerCatapult",DOFContactors),Catapult
		CP = True
		Else
		Controller.Switch(68) = 0
		CP = False
	End if
End Sub
Sub AutoPlunger_Hit:Controller.Switch(68) = 1:End Sub


Sub TopDiverter(enabled)
	If Enabled Then
		diverter.rotatetoend
		topdiverterwall.isdropped = 0
		PlaySoundAt SoundFX("TopdiverterOn",DOFContactors),Gi23
	Else
		diverter.rotatetostart
		topdiverterwall.isdropped = 1
		PlaySoundAt SoundFX("TopdiverterOff",DOFContactors),Gi23
	End If
End Sub

Sub UnderDiverterTop(enabled)
	If enabled Then
		DiverterFRG.isDropped = 0
		sw37dr1.isDropped = 1
	Else
		DiverterFRG.isDropped = 1
		sw37dr1.isDropped = 0
	End If
End Sub

Sub UnderDiverterBottom(enabled)
	If enabled Then
		DiverterFLG.isDropped = 1
	Else
		DiverterFLG.isDropped = 0
	End If
End Sub


Sub BorgKicker1_Hit()
	borglock.Addball Me
End Sub

Sub KickBack(Enabled)
	If enabled Then
		KickbackPlunger.Fire
		KB = True
'		KickBackTimer.Enabled = 1
		PlaysoundAt SoundFX("Kickback", DOFContactors),KickBackPlunger
		Else
	End if
End Sub

Sub LeftCannonMotor(enabled)
	If enabled Then
		CML = True
		CannonLTimer.Enabled = 1
		PlaySound SoundFX("LGunMotor",DOFGear), 0, 1, Pan(l52), 0, -50000,0,0,AudioFade(l52)
	Else
		CannonLTimer.Enabled = 0
		StopSound "LGunMotor"
	End If
End Sub

Sub RightCannonMotor(enabled)
	If enabled Then
		CMR = True
		CannonRTimer.Enabled = 1
		PlaySound SoundFX("RGunMotor",DOFGear), 0, 1, Pan(l82), 0, -50000,0,0,AudioFade(l82)
	Else
		CannonRTimer.Enabled = 0
		StopSound "RGunMotor"
	End If
End Sub

Sub UnderLeftGun(Enabled)
	If enabled then
	sw36.kick 161, 50, 1.4
	PlaySoundAt SoundFX("LGunPopper",DOFContactors),l52
	Controller.Switch(36) = 0
	Else
'	Controller.Switch(36) = 0
	vpmtimer.addtimer 1000, "sw36dr.isDropped = 1'"
	End if
End Sub

Sub UnderRightGun(Enabled)
	If enabled then
	sw37.kick 199, 50, 1.4
	PlaySoundAt SoundFX("RGunPopper",DOFContactors),l82
	Controller.Switch(37) = 0
	Else
'	Controller.Switch(37) = 0
	vpmtimer.addtimer 1000, "sw37dr.isDropped = 1'"
	End if
End Sub

Sub LeftLock(Enabled)
	If enabled then
	sw41.kick 255, 50, 1.4
	PlaySoundAt SoundFX("LeftPopper",DOFContactors),l52
	PlaySound "LeftPopperOut", 0, 0.4, -0.5
	Controller.Switch(41) = 0
	Else
	vpmtimer.addtimer 250, "sw41dr.isDropped = 1'"
	End if
End Sub


'*********
' Bumper
'*********

Sub LeftJetBumper_hit:vpmTimer.pulseSw 71:PlaySoundAtBumperVol SoundFX("BumperLeft", DOFContactors),LeftJetBumper,2:Me.TimerEnabled = 1:End Sub
Sub LeftJetBumper_Timer:Me.Timerenabled = 0:End Sub

Sub RightJetBumper_hit:vpmTimer.pulseSw 72:PlaysoundatBumperVol SoundFX("BumperRight", DOFContactors),RightJetBumper,2:Me.TimerEnabled = 1:End Sub
Sub RightJetBumper_Timer:Me.Timerenabled = 0:End Sub

Sub BottomJetBumper_hit:vpmTimer.pulseSw 73:PlaysoundAtBumperVol SoundFX("BumperMiddle", DOFContactors),BottomJetBumper,2:Me.TimerEnabled = 1:End Sub
Sub BottomJetBumper_Timer:Me.Timerenabled = 0:End Sub


'*********
' Switches
'*********

Sub sw15_Hit:Controller.Switch(15) = 1:PlaySoundAtBall "rollover":End Sub
Sub sw15_UnHit:Controller.Switch(15) = 0:End Sub

Sub sw16_Hit:Controller.Switch(16) = 1:End Sub
Sub sw16_UnHit:Controller.Switch(16) = 0:End Sub

Sub sw17_Hit:Controller.Switch(17) = 1:End Sub
Sub sw17_UnHit:Controller.Switch(17) = 0:End Sub

Sub sw18_Hit:Controller.Switch(18) = 1:PlaySoundAtBall "rollover":End Sub
Sub sw18_UnHit:Controller.Switch(18) = 0:End Sub

Sub sw117spinner_Spin():vpmTimer.PulseSw 117:PlaySound "fx_spinner",0,1,-.3,0,0,1,0,-.65:End Sub

Sub sw44_Hit:Controller.Switch(44) = 1:End Sub
Sub sw44_UnHit:Controller.Switch(44) = 0:End Sub
Sub sw48_Hit:Controller.Switch(48) = 1:End Sub
Sub sw48_UnHit:Controller.Switch(48) = 0:End Sub

Sub sw57_Hit:PlaySoundAtBallVol "target",1:TopDrop.hit 1:End Sub

Sub sw58_Hit:Controller.Switch(58) = 1:PlaySoundAtBall "rollover":End Sub
Sub sw58_UnHit:Controller.Switch(58) = 0:End Sub

Sub sw76_Hit:Controller.Switch(76) = 1:PlaySoundAtBall "rollover":End Sub
Sub sw76_UnHit:Controller.Switch(76) = 0:End Sub

Sub sw77_Hit:Controller.Switch(77) = 1:PlaySoundAtBall "rollover":End Sub
Sub sw77_UnHit:Controller.Switch(77) = 0:End Sub

Sub sw78_Hit:Controller.Switch(78) = 1:PlaySoundAtBall "rollover":End Sub
Sub sw78_UnHit:Controller.Switch(78) = 0:End Sub


Sub sw25_Hit:vpmTimer.PulseSw 25: PlaySoundAtBallVol "gate4",0.1: End Sub 'right Ramp
Sub sw87_Hit:vpmTimer.PulseSw 87: PlaySoundAtBallVol "gate4",0.1: End Sub

Sub sw88_Hit:vpmTimer.PulseSw 88: PlaySoundAtBallVol "gate4",0.1: End Sub 'left Ramp
Sub sw83_Hit:vpmTimer.PulseSw 83: sw83wire.RotX = 110:Me.TimerEnabled = 1:End Sub
Sub sw83_Timer: sw83wire.RotX = 90: Me.TimerEnabled = 0: End Sub

Sub sw23_Hit:vpmTimer.PulseSw 23: sw23wire.RotX = 110:Me.TimerEnabled = 1:End Sub 'middle Ramp
Sub sw23_Timer: sw23wire.RotX = 90: Me.TimerEnabled = 0: End Sub

Sub BallCatcher1_Hit:If ActiveBall.VelY <= -15 then ActiveBall.VelY = -15:End if:End Sub

'GunSwitches

Sub sw32_Hit:Controller.Switch(32) = 1:End Sub
'Sub sw32_UnHit:Controller.Switch(32) = 0:End Sub
Sub sw36_Hit:Controller.Switch(36) = 1:Controller.Switch(32) = 0:sw36dr.isDropped = 0:End Sub

Sub sw33_Hit:Controller.Switch(33) = 1:End Sub
'Sub sw33_UnHit:Controller.Switch(33) = 0:End Sub
Sub sw37_Hit:Controller.Switch(37) = 1:Controller.Switch(33) = 0:sw37dr.isDropped = 0:End Sub

'LeftLock

Sub sw41_Hit:Controller.Switch(41) = 1:sw41dr.isDropped = 0:End Sub

Sub sw35_Hit:Controller.Switch(35) = 1:sw35dr.isDropped = 0:End Sub
Sub sw35_UnHit:Controller.Switch(35) = 0:sw35dr.isDropped = 1:End Sub

Sub sw42_Hit:Controller.Switch(42) = 1:sw42dr.isDropped = 0:End Sub
Sub sw42_UnHit:Controller.Switch(42) = 0:sw42dr.isDropped = 1:End Sub

Sub sw43_Hit:Controller.Switch(43) = 1:sw43dr.isDropped = 0:End Sub
Sub sw43_UnHit:Controller.Switch(43) = 0:sw43dr.isDropped = 1:End Sub

'***Holes***

Sub sw27_Hit: vpmTimer.PulseSw 27:T27P.RotX = 100:T27P2.RotX = 100:PlaySoundAtBallVol SoundFX("target",DOFTargets),1:T27.TimerEnabled = 1:PlaySoundAt "NeutralZone1",sw27:End Sub
Sub sw45_Hit: SoundNeutralZone:vpmTimer.PulseSw 45:End Sub
'Sub sw45t_Hit: Controller.Switch(45) = 1:End Sub
'Sub sw45t_UnHit: Controller.Switch(45) = 0:End Sub
Sub sw46t_Hit: Controller.Switch(46) = 1:PlaySoundAt "CommandD",sw46:PlaySoundAt "NeutralZone1",sw46:End Sub
Sub sw46t_UnHit: Controller.Switch(46) = 0:End Sub
Sub sw47t_Hit: Controller.Switch(47) = 1:sw47dr.isDropped = 0:End Sub
Sub sw47_Hit:PlaySoundAt "StartMission1",sw47:End Sub
Sub sw47t_UnHit: Controller.Switch(47) = 0:sw47dr.isDropped = 1:End Sub


'***Targets***
'T27 only pulse with sw27(Hole) to prevent double hits
Sub T26_Hit:vpmTimer.PulseSw 26:PlaySoundAtBallVol SoundFX("target",DOFTargets),1.2:End Sub
'Sub T27_Hit:T27P.RotX = 100:T27P2.RotX = 100:PlaySound "target":T27.TimerEnabled = 1:End Sub
Sub T27_Timer:T27P.RotX = 90:T27P2.RotX = 90:End Sub
Sub T28_Hit:vpmTimer.PulseSw 28:PlaySoundAtBallVol SoundFX("target",DOFTargets),1.2:End Sub

Sub T51_Hit:vpmTimer.PulseSw 51:PlaySoundAtBallVol SoundFX("target",DOFTargets),1.2:End Sub
Sub T52_Hit:vpmTimer.PulseSw 52:PlaySoundAtBallVol SoundFX("target",DOFTargets),1.2:End Sub
Sub T53_Hit:vpmTimer.PulseSw 53:PlaySoundAtBallVol SoundFX("target",DOFTargets),1.2:End Sub

Sub T54_Hit:vpmTimer.PulseSw 54:PlaySoundAtBallVol SoundFX("target",DOFTargets),1.2:End Sub
Sub T55_Hit:vpmTimer.PulseSw 55:PlaySoundAtBallVol SoundFX("target",DOFTargets),1.2:End Sub
Sub T56_Hit:vpmTimer.PulseSw 56:PlaySoundAtBallVol SoundFX("target",DOFTargets),1.2:End Sub

Sub T81_Hit:vpmTimer.PulseSw 81:PlaySoundAtBallVol SoundFX("target",DOFTargets),1.2:End Sub
Sub T82_Hit:vpmTimer.PulseSw 82:PlaySoundAtBallVol SoundFX("target",DOFTargets),1.2:End Sub

Sub T84_Hit:vpmTimer.PulseSw 84:PlaySoundAtBallVol SoundFX("target",DOFTargets),1.2:End Sub

Sub T85_Hit:vpmTimer.PulseSw 85:PlaySoundAtBallVol SoundFX("target",DOFTargets),1.2:End Sub
Sub T86_Hit:vpmTimer.PulseSw 86:PlaySoundAtBallVol SoundFX("target",DOFTargets),1.2:End Sub


'KickerAnimation


'***************
' ToysAnimation
'***************


'***Cannons***

Dim CannonSpeed, CannonSteps, CannonAnimationSteps, CannonLBall, CannonRBall, RadiusL, RadiusR, CannonAngleL, CannonAngleR
Dim RadiusRL, RadiusLL, CML, CMR, CPL, CPR
Dim ForceL, ForceR, LaserLON, LaserRON

CannonSpeed = 600 'Timer with 1 = 1000/Second
CannonSteps = 82

Const Pi = 3.14159265358979
RadiusL = CannonBaseL.Y - (kicker1.Y-6)
RadiusR = CannonBaseR.Y - (kicker2.Y-6)
RadiusLL = CannonBaseL.Y - l52.Y
RadiusRL = CannonBaseR.Y - l82.Y


Sub Kicker1_Hit()
'	me.destroyball
'	Set CannonLBall = kicker1.createball
	Set CannonLBall = ActiveBall
	Controller.Switch(38) = 1
	me.Enabled = False
	If LaserMod = 1 then
		LaserL.Size_Z = 0.3
		LaserL1.Size_Z = 0.3
		LaserL.visible = 1
		LaserL1.visible = 1
		LaserZTimer.Enabled = 1
		LaserLON = 1
	End if
End Sub

Sub Kicker2_Hit()
'	me.destroyball
'	Set CannonRBall = kicker2.createball
	Set CannonRBall = ActiveBall
	Controller.Switch(34) = 1
	me.Enabled = False
	If LaserMod = 1 then
		LaserR.Size_Z = 0.5
		LaserR1.Size_Z = 0.5
		LaserR.visible = 1
		LaserR1.visible = 1
		LaserZTimer1.Enabled = 1
		LaserRON = 1
	End if
End Sub




Sub CannonLTimer_Timer()
		If CML = True and CannonBaseL.ObjRotZ < 64 Then
		CannonBaseL.ObjRotZ = CannonBaseL.ObjRotZ +(CannonSteps*2)/CannonSpeed
		End if
		If CML = False and CannonBaseL.ObjRotZ > -19 then
		CannonBaseL.ObjRotZ = CannonBaseL.ObjRotZ -(CannonSteps*2)/CannonSpeed
		End if
		If CannonBaseL.ObjRotZ >= 64 then CML = False
		If CannonBaseL.ObjRotZ >= -20 And CannonBaseL.ObjRotZ <= -17 then Controller.Switch(127) = 1: Else Controller.Switch(127) = 0
		If CannonBaseL.ObjRotZ >= -20 And CannonBaseL.ObjRotZ <= 9 then Controller.Switch(122) = 1:ForceL = 25: Else Controller.Switch(122) = 0: ForceL = 50

		If Not IsEmpty(CannonLBall) Then
		CannonLBall.X= CannonBaseL.X - RadiusL*sin(-CannonBaseL.ObjRotZ*(Pi/180))
		CannonLBall.Y= CannonBaseL.Y - RadiusL*cos(-CannonBaseL.ObjRotZ*(Pi/180))
		CannonLBall.Z=105 + BSize
		CannonAngleL = CannonBaseL.ObjRotZ
		End If
		CannonL.ObjRotZ = CannonBaseL.ObjRotZ
		CannonDomeL.ObjRotZ = CannonBaseL.ObjRotZ
		CannonPinL.ObjRotZ = CannonBaseL.ObjRotZ
		l52.X = CannonBaseL.X - RadiusRL*sin(-CannonBaseL.ObjRotZ*(Pi/180))
		l52.Y = CannonBaseL.Y - RadiusRL*cos(-CannonBaseL.ObjRotZ*(Pi/180))
		LaserL.ObjRotZ = CannonBaseL.ObjRotZ
		LaserL1.ObjRotZ = CannonBaseL.ObjRotZ
End Sub



Sub CannonRTimer_Timer()
		If CMR = True and CannonBaseR.ObjRotZ < 64 Then
		CannonBaseR.ObjRotZ = CannonBaseR.ObjRotZ +(CannonSteps*2)/CannonSpeed
		End if
		If CMR = False and CannonBaseR.ObjRotZ > -19 then
		CannonBaseR.ObjRotZ = CannonBaseR.ObjRotZ -(CannonSteps*2)/CannonSpeed
		End if
		If CannonBaseR.ObjRotZ >= 64 then CMR = False
		If CannonBaseR.ObjRotZ >= -20 And CannonBaseR.ObjRotZ <= -17 then Controller.Switch(125) = 1: Else Controller.Switch(125) = 0
		If CannonBaseR.ObjRotZ >= -20 And CannonBaseR.ObjRotZ <= 9 then Controller.Switch(126) = 1:ForceR = 25: Else Controller.Switch(126) = 0: ForceR = 50

		If Not IsEmpty(CannonRBall) Then
		CannonRBall.X= CannonBaseR.X - RadiusR*sin(+CannonBaseR.ObjRotZ*(Pi/180))
		CannonRBall.Y= CannonBaseR.Y - RadiusR*cos(+CannonBaseR.ObjRotZ*(Pi/180))
		CannonRBall.Z=105 + BSize
		CannonAngleR = CannonBaseR.ObjRotZ- (CannonBaseR.ObjRotZ*2)
		End If
		CannonR.ObjRotZ = CannonBaseR.ObjRotZ
		CannonDomeR.ObjRotZ = CannonBaseR.ObjRotZ
		CannonPinR.ObjRotZ = CannonBaseR.ObjRotZ
		l82.X = CannonBaseR.X - RadiusRL*sin(+CannonBaseR.ObjRotZ*(Pi/180))
		l82.Y = CannonBaseR.Y - RadiusRL*cos(+CannonBaseR.ObjRotZ*(Pi/180))
		LaserR.ObjRotZ = CannonBaseR.ObjRotZ
		LaserR1.ObjRotZ = CannonBaseR.ObjRotZ
End Sub

'***CannonShoot***

Sub LeftCannonKicker(Enabled)
	If Enabled then
		CPL = True
		If Not IsEmpty(CannonLBall) Then
		Kicker1.kick CannonAngleL, ForceL
		PlaysoundAt SoundFX("LGunKicker", DOFContactors),l52
		Controller.Switch(38) = 0
		CannonLBall = Empty
		Kicker1.Enabled = True
		LaserLON = 0
		If LaserType = 1 then
			LaserL.Visible = 0
			LaserL1.Visible = 0
			LaserZTimer.Enabled = 0
		End if
		End if
	End if
End Sub


Sub RightCannonKicker(Enabled)
	If Enabled then
		CPR = True
		If Not IsEmpty(CannonRBall) Then
		Kicker2.kick CannonAngleR, ForceR
		PlaysoundAt SoundFX("RGunKicker", DOFContactors),l82
		Controller.Switch(34) = 0
		CannonRBall = Empty
		Kicker2.Enabled = True
		LaserRON = 0
		If LaserType = 1 then
			LaserR.Visible = 0
			LaserR1.Visible = 0
			LaserZTimer1.Enabled = 0
		End if
		End if
	End if
End Sub


'***LaserZTimer***

Sub LaserZTimer_Timer()
	If CannonBaseL.ObjRotZ >= -20 And CannonBaseL.ObjRotZ <= -17 then LaserL.Size_Z = 0.3: LaserL1.Size_Z = 0.3
	If CannonBaseL.ObjRotZ >= -16.99 And CannonBaseL.ObjRotZ <= -8 then LaserL.Size_Z = 0.25: LaserL1.Size_Z = 0.25
	If CannonBaseL.ObjRotZ >= -7.99 And CannonBaseL.ObjRotZ <= -5 then LaserL.Size_Z = 0.3: LaserL1.Size_Z = 0.3
	If CannonBaseL.ObjRotZ >= -4.99 And CannonBaseL.ObjRotZ <= -1 then LaserL.Size_Z = 0.67: LaserL1.Size_Z = 0.67
	If CannonBaseL.ObjRotZ >= 0.99 And CannonBaseL.ObjRotZ <= 5.99 then LaserL.Size_Z = 1: LaserL1.Size_Z = 1

	If BorgMod = 0 then
		If CannonBaseL.ObjRotZ >= 6 And CannonBaseL.ObjRotZ <= 21.99 then LaserL.Size_Z = 0.75: LaserL1.Size_Z = 0.75
	End if
	If BorgMod = 1 then
		If CannonBaseL.ObjRotZ >= 6 And CannonBaseL.ObjRotZ <= 21.99 then LaserL.Size_Z = 1: LaserL1.Size_Z = 1
	End if

	If CannonBaseL.ObjRotZ >= 22 And CannonBaseL.ObjRotZ <= 56.99 then LaserL.Size_Z = 1: LaserL1.Size_Z = 1
	If CannonBaseL.ObjRotZ >= 57 And CannonBaseL.ObjRotZ <= 62.99 then LaserL.Size_Z = 0.48: LaserL1.Size_Z = 0.48
	If CannonBaseL.ObjRotZ >= 63 then LaserL.Size_Z = 1:LaserL1.Size_Z = 1

	If CannonBaseL.ObjRotZ <= -19 And LaserLON = 0 Then
		LaserL.Visible = 0
		LaserL1.Visible = 0
		LaserZTimer.Enabled = 0
	End if
End Sub

Sub LaserZTimer1_Timer()
	If CannonBaseR.ObjRotZ >= -20 And CannonBaseR.ObjRotZ <= -18 then LaserR.Size_Z = 0.5: LaserR1.Size_Z = 0.5
	If CannonBaseR.ObjRotZ >= -17.99 And CannonBaseR.ObjRotZ <= -8 then LaserR.Size_Z = 0.235: LaserR1.Size_Z = 0.235
	If CannonBaseR.ObjRotZ >= -7.99 And CannonBaseR.ObjRotZ <= 8.99 then LaserR.Size_Z = 1: LaserR1.Size_Z = 1

	If BorgMod = 0 then
		If CannonBaseR.ObjRotZ >= 9 And CannonBaseR.ObjRotZ <= 22.99 then LaserR.Size_Z = 0.76: LaserR1.Size_Z = 0.76
	End if
	If BorgMod = 1 then
		If CannonBaseR.ObjRotZ >= 9 And CannonBaseR.ObjRotZ <= 22.99 then LaserR.Size_Z = 1: LaserR1.Size_Z = 1
	End if

	If CannonBaseR.ObjRotZ >= 23 And CannonBaseR.ObjRotZ <= 28.99 then LaserR.Size_Z = 1: LaserR1.Size_Z = 1
	If CannonBaseR.ObjRotZ >= 29 And CannonBaseR.ObjRotZ <= 32.99 then LaserR.Size_Z = 0.785: LaserR1.Size_Z = 0.785
	If CannonBaseR.ObjRotZ >= 33 And CannonBaseR.ObjRotZ <= 50.99 then LaserR.Size_Z = 1: LaserR1.Size_Z = 1
	If CannonBaseR.ObjRotZ >= 51 And CannonBaseR.ObjRotZ <= 57.99 then LaserR.Size_Z = 0.51: LaserR1.Size_Z = 0.51
	If CannonBaseR.ObjRotZ >= 58 And CannonBaseR.ObjRotZ <= 62.99 then LaserR.Size_Z = 0.49: LaserR1.Size_Z = 0.49
	If CannonBaseR.ObjRotZ >= 63 then LaserR.Size_Z = 1: LaserR1.Size_Z = 1

	If CannonBaseR.ObjRotZ <= -18 And LaserRON = 0 Then
		LaserR.Visible = 0
		LaserR1.Visible = 0
		LaserZTimer1.Enabled = 0
	End if
End Sub

'***CannonPin Animation***

Sub CannonPinLTimer
	If CPL = True and CannonPinL.TransZ < 40 then CannonPinL.TransZ = CannonPinL.TransZ + 5
	If CPL = False and CannonPinL.TransZ > 0 then CannonPinL.TransZ = CannonPinL.TransZ - 5
	If CannonPinL.TransZ >= 35 then CPL = False
End Sub


Sub CannonPinRTimer
	If CPR = True and CannonPinR.TransZ < 40 then CannonPinR.TransZ = CannonPinR.TransZ + 5
	If CPR = False and CannonPinR.TransZ > 0 then CannonPinR.TransZ = CannonPinR.TransZ - 5
	If CannonPinR.TransZ >= 35 then CPR = False
End Sub

'***Kickback Animation***


Sub KickBackTimer()
	if KB = True and KickBackP.TransZ < 60 then KickBackP.TransZ = KickBackP.TransZ +5
	if KB = False and KickBackP.TransZ > 0 then KickBackP.TransZ = KickBackP.TransZ -5
	if KickBackP.TransZ >= 60 then KB = False
End Sub

'***Catapult Animation***


Sub CatapultTimer()
	if CP = True and Catapult.RotX < 110 then Catapult.RotX = Catapult.RotX +10
	if CP = False and Catapult.RotX > 90 then Catapult.RotX = Catapult.RotX -10
'	if Catapult.RotX >= 107 then CP = False
End Sub

'************
' SlingShots
'************


Dim RStep, Lstep

Sub SlingShotRight_Slingshot
    PlaySoundAt SoundFX("SlingshotRight", DOFContactors),SLING1
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    SlingShotRight.TimerEnabled = 1
    vpmTimer.PulseSw 74
End Sub

Sub SlingShotRight_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:SlingShotRight.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub SlingShotLeft_Slingshot
    PlaySoundAt SoundFX("SlingshotLeft", DOFContactors),SLING2
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    SlingShotLeft.TimerEnabled = 1
    vpmTimer.PulseSw 75
End Sub

Sub SlingShotLeft_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:SlingShotLeft.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
    If KeyCode = MechanicalTilt Then
        vpmTimer.PulseSw vpmNudge.TiltSwitch
        Exit Sub
    End if
    If keycode = PlungerKey Then Controller.Switch(12) = 1
	If keycode = keyFront Then Controller.Switch(11) = 1
	If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If keycode = PlungerKey Then Controller.Switch(12) = 0
	If keycode = keyFront Then Controller.Switch(11) = 0
	If vpmKeyUp(keycode) Then Exit Sub
End Sub

Sub table1_Paused:Controller.Pause = True:End Sub
Sub table1_unPaused:Controller.Pause = False:End Sub
Sub table1_exit():Controller.Pause = False:Controller.Stop:End Sub



'**************
' Flipper Subs
'**************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"


Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("FlipperUpLeft", DOFFlippers),FlipperL:FlipperL.RotateToEnd
    Else
        PlaySoundAt SoundFX("FlipperDown", DOFFlippers),FlipperL:FlipperL.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("FlipperUpRightBoth", DOFFlippers),FlipperR:FlipperR.RotateToEnd:FlipperR1.RotateToEnd
    Else
        PlaySoundAt SoundFX("FlipperDown", DOFFlippers),FlipperR:FlipperR.RotateToStart:FlipperR1.RotateToStart
    End If
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
Dim ccount

InitLamps()             ' turn off the lights and flashers and reset them to the default parameters
LampTimer.Interval = 10 'lamp fading speed
LampTimer.Enabled = 1

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



Sub InitLamps()
    Dim x
    For x = 0 to 200
        LampState(x) = 0         ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4       ' used to track the fading state
        FlashSpeedUp(x) = 0.5    ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.1 ' slower speed when turning off the flasher
        FlashMax(x) = 1          ' the maximum value when on, usually 1
        FlashMin(x) = 0          ' the minimum value when off, usually 0
        FlashLevel(x) = 0        ' the intensity of the flashers, usually from 0 to 1
    Next
End Sub

 Sub UpdateLamps
	'Inserts
	Flash 11, f11
	NFadeLm 11,  l11
	NFadeLm 11,  l11b
	Flash 12, f12
	NFadeLm 12,  l12
	NFadeLm 12,  l12b
	NFadeLm 13,  l13
	NFadeLm 13,  l13b
	NFadeLm 14,  l14
	NFadeLm 14,  l14b
	Flash 15, f15
	NFadeLm 15,  l15
	NFadeLm 15,  l15b
	NFadeLm 16,  l16
	NFadeLm 16,  l16b
	NFadeLm 17,  l17
	NFadeLm 17,  l17b
	NFadeLm 18,  l18
	NFadeLm 18,  l18b

	NFadeLm 21,  l21
	NFadeLm 21,  l21b
	NFadeLm 22,  l22
	NFadeLm 22,  l22b
	NFadeLm 23,  l23
	NFadeLm 23,  l23b
	Flash 24, f24
	NFadeLm 24,  l24
	NFadeLm 24,  l24b
	Flash 25, f25
	NFadeLm 25,  l25
	NFadeLm 25,  l25b
	Flash 26, f26
	Flashm 26, f26s
	Flashm 26, f26s1
	NFadeObjm 26, l26blue, "bulbcover_blueOn", "bulbcover_blue"
	NFadeLm 26,  l26
	NFadeLm 26,  l26b
	NFadeLm 27,  l27
	NFadeLm 27,  l27b
	Flash 28, f28
	NFadeLm 28,  l28
	NFadeLm 28,  l28b

	NFadeLm 31,  l31
	NFadeLm 31,  l31b
	NFadeLm 32,  l32
	NFadeLm 32,  l32b
	NFadeLm 33,  l33
	NFadeLm 33,  l33b
	Flash 34, f34
	NFadeLm 34,  l34
	NFadeLm 34,  l34b
	NFadeLm 35,  l35
	NFadeLm 35,  l35b
	NFadeLm 36,  l36
	NFadeLm 36,  l36b
	NFadeLm 37,  l37
	NFadeLm 37,  l37b
	Flash 38, f38
	NFadeLm 38,  l38
	NFadeLm 38,  l38b

	Flash 41, f41
	NFadeLm 41,  l41
	NFadeLm 41,  l41b
	NFadeLm 42,  l42
	NFadeLm 42,  l42b
 	NFadeLm 43,  l43
 	NFadeLm 43,  l43b
	NFadeLm 44,  l44
	NFadeLm 44,  l44b
	Flash 45, f45
	NFadeLm 45,  l45
	NFadeLm 45,  l45b
	NFadeLm 46,  l46
	NFadeLm 46,  l46b
	NFadeLm 47,  l47
	NFadeLm 47,  l47b
	Flash 48, f48
	NFadeLm 48,  l48
	NFadeLm 48,  l48b

	NFadeLm 51,  l51
	NFadeLm 51,  l51b
	Flash 52, f52
	Flashm 52, f52s
	Flashm 52, f52s1
	NFadeLm 52,  l52
	Flash 53, f53
	Flashm 53, f53s
	NFadeObjm 53, l53yellow, "bulbcover_yellowOn", "bulbcover_yellow"
	NFadeLm 54,  l54
	NFadeLm 54,  l54b
	NFadeLm 55,  l55
	NFadeLm 55,  l55b
	NFadeLm 56,  l56
	NFadeLm 56,  l56b
	NFadeLm 57,  l57
	NFadeLm 57,  l57b
	NFadeLm 58,  l58
	NFadeLm 58,  l58b

	NFadeLm 61,  l61
	NFadeLm 61,  l61b
	NFadeLm 62,  l62
	NFadeLm 62,  l62b
	NFadeLm 63,  l63
	NFadeLm 63,  l63b
	NFadeLm 64,  l64
	NFadeLm 64,  l64b
	NFadeLm 65,  l65
	NFadeLm 65,  l65b
	NFadeLm 66,  l66
	NFadeLm 66,  l66b
	Flash 67, f67
	NFadeLm 67,  l67
	NFadeLm 67,  l67b
	Flash 68, f68
	NFadeLm 68,  l68
	NFadeLm 68,  l68b

	NFadeLm 71,  l71
	NFadeLm 71,  l71b
	NFadeLm 72,  l72
	NFadeLm 72,  l72b
	NFadeLm 73,  l73
	NFadeLm 73,  l73b
	NFadeLm 74,  l74
	NFadeLm 74,  l74b
	NFadeLm 75,  l75
	NFadeLm 75,  l75b
	NFadeLm 76,  l76
	NFadeLm 76,  l76b
	NFadeLm 77,  l77
	NFadeLm 77,  l77b
	NFadeLm 78,  l78a
	NFadeLm 78,  l78b
	NFadeLm 78,  l78c
	NFadeLm 78,  l78d
	NFadeLm 78,  l78e

	NFadeLm 78,  l78borga
	NFadeLm 78,  l78borgb
	NFadeLm 78,  l78borgc
	NFadeLm 78,  l78borgd
	NFadeLm 78,  l78borge

	NFadeLm 81,  l81
	NFadeLm 81,  l81b
	Flash 82,  f82
	Flashm 82,  f82s
	Flashm 82,  f82s1
	NFadeLm 82,  l82
	Flash 84, f84
	NFadeLm 84,  l84
	NFadeLm 84,  l84b
	Flash 85, f85
	Flashm 85, f85s
	NFadeObjm 85, l85green, "bulbcover_greenOn", "bulbcover_green"
	Flash 86, f86
	Flashm 86, f86s
	Flashm 86, f86s1
	NFadeObjm 86, l86red, "bulbcover_redOn", "bulbcover_red"
	NFadeLm 124, l124
	NFadeLm 124, l124b
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

Sub NFadeLmb(nr, object) ' used for multiple lights with blinking
    Select Case FadingLevel(nr)
        Case 4:object.state = 0
        Case 5:object.state = 2
    End Select
End Sub

'Lights, Ramps & Primitives used as 4 step fading lights
'a,b,c,d are the images used from on to off

Sub FadeObj(nr, object, a, b, c, d)
	Select Case FadingLevel(nr)
		Case 2:object.image = d:FadingLevel(nr) = 0 'Off
		Case 3:object.image = c:FadingLevel(nr) = 2 'fading...
		Case 4:object.image = b:FadingLevel(nr) = 3 'fading...
		Case 5:object.image = a:FadingLevel(nr) = 1 'ON
 End Select
End Sub

Sub FadeObjm(nr, object, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:object.image = d
        Case 3:object.image = c
        Case 4:object.image = b
        Case 5:object.image = d
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


'***SolModFlasher***

Sub Flash120(Enabled) 'JetsFlasher
	l120.Intensity = enabled / 10
	If enabled then
	l83.State = 1
	l83.Intensity = enabled / 10
	Else
	l83.Intensity = 20
	l83.State = 0
	End If
End Sub

Sub Flash121(Enabled) 'RightPopper
	l121.Intensity = enabled / 10
	f121.Opacity = enabled *2
End Sub

Sub Flash122(Enabled) 'MiddleRamp
	f122.Opacity = enabled *2
	f122b.Opacity = enabled
	f122s.Opacity = enabled * 0.8
End Sub

Sub Flash123(Enabled) 'ShieldFlasher
	f123.Opacity = enabled * 10
	f123a.Opacity = enabled * 3
	ShieldGiBig7.Intensity = enabled / 2
	ShieldGiBig8.Intensity = enabled / 2
	ShieldGiBig9.Intensity = enabled / 2
	ShieldGiBig10.Intensity = enabled / 2
	ShieldGiBig11.Intensity = enabled / 2
	ShieldGiBig12.Intensity = enabled / 2
End Sub


Sub Flash125(Enabled) 'LeftPopper
	l125.Intensity = enabled / 10
	f125.Opacity = enabled *2
'	if enabled Then
'	flasher1.opacity = 2000
'	flasher2.opacity = 2000
'	Else
'	flasher1.opacity = 800
'	flasher2.opacity = 800
'	End if
End Sub

Sub Flash126(Enabled) 'RightBorgFlasher
	f126.Opacity = enabled *12
	f126s.Opacity = enabled /2
	l78d1.Intensity = enabled
	l78e1.Intensity = enabled
	l78borgb1.Intensity = enabled
	l78borga1.Intensity = enabled
End Sub

Sub Flash127(Enabled) 'LeftBorgFlasher
	f127.Opacity = enabled * 8
	f127s.Opacity = enabled /2
	l78b1.Intensity = enabled
	l78c1.Intensity = enabled
	l78borgd1.Intensity = enabled
	l78borge1.Intensity = enabled
End Sub

Sub Flash128(Enabled) 'CenterBorgFlasher
	f128.Opacity = enabled * 8
	f128s.Opacity = enabled * 0.8
	GiBigGreen.Opacity = enabled *2
	l78a1.Intensity = enabled *2
	l78borgc1.Intensity = enabled *2
End Sub

Sub Flash141(Enabled) 'RomulanFlasher
	f141.Opacity = enabled *50
	f141s.Opacity = enabled
	f141s1.Opacity = enabled
	l141.Intensity = enabled / 5
	l141a.Intensity = enabled / 10
	GiBigGreen.Opacity = enabled *2
End Sub

Sub Flash142(Enabled) 'RightRampFlasher
	If Enabled Then
		FlasherCapRed.Image = "dome4redOn"
	Else
		FlasherCapRed.Image = "dome4red"
	End if
	GiBigRed.Opacity = enabled *2
	GiBigRed1.Opacity = enabled *2
	f142.Opacity = enabled *12
	l142.Intensity = enabled
	f142s.Opacity = enabled
	f142s1.Opacity = enabled
End Sub

'BorgLampTimer - turns OFF the blue lamps while the green flasher are ON

'Sub BorgLampTimer_Timer()
Sub BorgLampTimer
	If BorgMod = 1 then
		If l78a1.Intensity >= 10 Then l78a.Intensity = 0:Else:l78a.Intensity = 30:End if
		If l78b1.Intensity >= 10 Then l78b.Intensity = 0:Else:l78b.Intensity = 30:End if
		If l78c1.Intensity >= 10 Then l78c.Intensity = 0:Else:l78c.Intensity = 6:End if
		If l78d1.Intensity >= 10 Then l78d.Intensity = 0:Else:l78d.Intensity = 30:End if
		If l78e1.Intensity >= 10 Then l78e.Intensity = 0:Else:l78e.Intensity = 10:End if
	End if
	If BorgMod = 0 then
		If l78borga1.Intensity >= 10 Then l78borga.Intensity = 0:Else:l78borga.Intensity = 15:End if
		If l78borgb1.Intensity >= 10 Then l78borgb.Intensity = 0:Else:l78borgb.Intensity = 50:End if
		If l78borgc1.Intensity >= 10 Then l78borgc.Intensity = 0:Else:l78borgc.Intensity = 60:End if
		If l78borgd1.Intensity >= 10 Then l78borgd.Intensity = 0:Else:l78borgd.Intensity = 30:End if
		If l78borge1.Intensity >= 10 Then l78borge.Intensity = 0:Else:l78borge.Intensity = 50:End if
	End if
End Sub

'**************
'***RGB Mode***
'**************

Dim RGBStep, RGBFactor, Red, Green, Blue

RGBStep = 0
RGBFactor = 1
Red = 255
Green = 0
Blue = 0

Sub RGBTimer_timer 'rainbow light color changing
    Select Case RGBStep
        Case 0 'Green
            Green = Green + RGBFactor
            If Green > 255 then
                Green = 255
                RGBStep = 1
            End If
        Case 1 'Red
            Red = Red - RGBFactor
            If Red < 0 then
                Red = 0
                RGBStep = 2
            End If
        Case 2 'Blue
            Blue = Blue + RGBFactor
            If Blue > 255 then
                Blue = 255
                RGBStep = 3
            End If
        Case 3 'Green
            Green = Green - RGBFactor
            If Green < 0 then
                Green = 0
                RGBStep = 4
            End If
        Case 4 'Red
            Red = Red + RGBFactor
            If Red > 255 then
                Red = 255
                RGBStep = 5
            End If
        Case 5 'Blue
            Blue = Blue - RGBFactor
            If Blue < 0 then
                Blue = 0
                RGBStep = 0
            End If
    End Select
    'Light1.color = RGB(Red\10, Green\10, Blue\10)
'    light1.colorfull = RGB(Red, Green, Blue)

	If RGBBumpers = 1 then
		lbumperr.colorfull = RGB(Red, Green, Blue)
		lbumperr1.colorfull = RGB(Red, Green, Blue)
		lbumperr2.colorfull = RGB(Red, Green, Blue)
		lbumperr3.colorfull = RGB(Red, Green, Blue)
		lbumperr4.colorfull = RGB(Red, Green, Blue)
		lbumperr5.colorfull = RGB(Red, Green, Blue)
	End if
	If RGBArrows = 1 then
		l42.colorfull = RGB(Red, Green, Blue)
		l42b.colorfull = RGB(Red, Green, Blue)
		l46.colorfull = RGB(Red, Green, Blue)
		l46b.colorfull = RGB(Red, Green, Blue)
		l54.colorfull = RGB(Red, Green, Blue)
		l54b.colorfull = RGB(Red, Green, Blue)
		l61.colorfull = RGB(Red, Green, Blue)
		l61b.colorfull = RGB(Red, Green, Blue)
		l64.colorfull = RGB(Red, Green, Blue)
		l64b.colorfull = RGB(Red, Green, Blue)
		l71.colorfull = RGB(Red, Green, Blue)
		l71b.colorfull = RGB(Red, Green, Blue)
		l76.colorfull = RGB(Red, Green, Blue)
		l76b.colorfull = RGB(Red, Green, Blue)
	End if
	'Gi
'	gi4.colorfull = RGB(Red, Green, Blue)
'	gis7.colorfull = RGB(Red, Green, Blue)
'	gi8.colorfull = RGB(Red, Green, Blue)
'	gis8.colorfull = RGB(Red, Green, Blue)

End Sub


'***********
' Update GI
'***********

Dim gistep, xx, obj
   gistep = 1 / 8

Sub UpdateGI(no, step)
	If step > 1 Then
		DOF 200, DOFOn
	Else
		DOF 200, DOFOff
	End If
    Select Case no
        Case 0
            For each xx in St1Shields:xx.IntensityScale = gistep*step:next
'        Case 1
'            For each xx in St2Gi1:xx.IntensityScale = gistep*step:next
'			Table1.ColorGradeImage = "grade_" & step
'        Case 2
'            For each xx in St3Gi2:xx.IntensityScale = gistep*step:next
'			Table1.ColorGradeImage = "grade_" & step
        Case 3
            For each xx in St4PFGI:xx.IntensityScale = gistep*step:next
			Table1.ColorGradeImage = "grade_" & step
			For each xx in St4PFGI: if xx.IntensityScale = 0 then Table1.ColorGradeImage = "grade_1":End if:next
        Case 4
            For each xx in St5ReLa:xx.IntensityScale = gistep*step:next
			Table1.ColorGradeImage = "grade_" & step
			For each xx in St5ReLa: if xx.IntensityScale = 0 then Table1.ColorGradeImage = "grade_1":End if:next
    End Select
End Sub

'****************************
'     Realtime Updates
' called by the MotorCallBack
'****************************

Sub RealTimeUpdates
    'flippers,gates
    FlipperLP.RotY = FlipperL.CurrentAngle
    FlipperLShadow.RotZ = FlipperL.CurrentAngle
    FlipperRP.RotY = FlipperR.CurrentAngle
    FlipperRShadow.RotZ = FlipperR.CurrentAngle
	FlipperRP1.RotY = FlipperR1.CurrentAngle
    FlipperR1Shadow.RotZ = FlipperR1.CurrentAngle
	diverterp.RotY = diverter.CurrentAngle
	sw25spinnerp.RotX = sw25spinner.CurrentAngle +90
	sw87spinnerp.RotX = sw87spinner.CurrentAngle +90
	sw88spinnerp.RotX = sw88spinner.CurrentAngle +90

    ' other stuff
    RollingSoundUpdate
	BallShadowUpdate
	BorgLampTimer
	CatapultTimer
	KickbackTimer
	CannonPinRTimer
	CannonPinLTimer
End Sub

'*******************
' WireRampDropSounds
'*******************

Dim SoundBall

Sub LeftPopperTriggerStart_Hit()
	Set SoundBall = ActiveBall
	vpmTimer.AddTimer 300, "WireHit'"
	'PlaySound "WireRamp", -1, 0.02, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub LeftWireTriggerStart_Hit()
	StopSound "Plastic_Ramp"
	Set SoundBall = ActiveBall
	'PlaySound "WireRamp", -1, 0.01, -.3, 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub LeftWireTriggerEnd_Hit()
	StopSound "WireRamp"
	PlaySound "metalhit2", 0, 1, -0.4, 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	vpmTimer.AddTimer 200, "BallDropSoundLeft'"
End Sub

Sub RightWireTriggerStart_Hit()
	StopSound "Plastic_Ramp"
	Set SoundBall = ActiveBall
	'PlaySound "WireRamp", -1, 0.01, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub RightWireTriggerEnd_Hit()
	StopSound "WireRamp"
	PlaySound "metalhit2", 0, 1, 0.4, 0, 0, 1, 0, .8
	vpmTimer.AddTimer 200, "BallDropSoundRight2'"
End Sub

Sub MiddleWireTriggerStart_Hit()
	StopSound "Plastic_Ramp"
	Set SoundBall = ActiveBall
	'PlaySound "WireRamp", -1, 0.01, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub MiddleWireTriggerEnd_Hit()
	StopSound "WireRamp"
	PlaySound "wireramp_hit", 0, 0.1, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub PlungerWireTriggerStart_Hit()
	Set SoundBall = ActiveBall
	'PlaySound "WireRamp", -1, 0.01, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub PlungerWireTriggerEnd_Hit()
	StopSound "WireRamp"
	vpmTimer.AddTimer 100, "BallDropSoundRight'"
End Sub

Sub BorgTrigger_Hit()
	vpmTimer.AddTimer 100, "BallDropSoundCenter'"
End Sub

Sub BallDropSoundCenter()
	PlaySound "BallDrop2",0,.6,0,0,0,0,0,-.8
End Sub

Sub BallDropSoundLeft()
	PlaySound "BallDrop", 0, 0.6, -.4,0,0,0,0,.8
End Sub

Sub BallDropSoundRight()
	PlaySound "BallDrop", 0, 1,.2, 0, 0, 1, 0, -.8
End Sub

Sub BallDropSoundRight2()
	PlaySound "BallDrop", 0, 0.6,.4, 0, 0, 1, 0, .8
End Sub

Sub WireHit()
	PlaySound "WireRamp_Hit", 0, 0.1, -.2, 0, 0, 1, 0, -.6
End Sub


'***plasticrolling***

'Sub RampBetaP_Hit()
'	Set SoundBall = ActiveBall
'	PlaySound "plasticrolling", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
'End Sub
'
'Sub RampAlphaP_Hit()
'	Set SoundBall = ActiveBall
'	PlaySound "plasticrolling", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
'End Sub
'
'Sub RampDeltaP_Hit()
'	Set SoundBall = ActiveBall
'	PlaySound "plasticrolling", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
'End Sub

' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 300)
End Function

Function Vol2(ball1, ball2) ' Calculates the Volume of the sound based on the speed of two balls
    Vol2 = (Vol(ball1) + Vol(ball2) ) / 2
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

function AudioFade(ball)
	Dim tmp
    tmp = ball.y * 2 / Table1.height-1
    If tmp > 0 Then
		AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

'*****************************************
'    ROLLING SOUND - thanks Ninuzzo
'*****************************************

Const tnob = 6						' total number of balls : 6 (trough) + 1 (Captive Ball)
Const fakeballs = 0					' number of balls created on table start (rolling sound will be skipped)
ReDim rolling(tnob)

Sub InitRolling:Dim i:For i=0 to (tnob-1):rolling(i) = False:Next:End Sub

Sub RollingSoundUpdate()
    Dim BOT, b
    BOT = GetBalls
	' stop the sound of deleted balls
	If UBound(BOT)<(tnob - 1) Then
		For b = (UBound(BOT) + 1) to (tnob-1)
			rolling(b) = False
			StopSound("fx_ballrolling" & b)
		Next
	End If
	' exit the Sub if no balls on the table
    If UBound(BOT) = fakeballs-1 Then Exit Sub
       ' play the rolling sound for each ball
    For b = 0 to UBound(BOT)
        If BallVel(BOT(b) ) > 1 Then
			rolling(b) = True
			if BOT(b).z < 30 Then ' Ball on playfield
						PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
			Else ' Ball on raised ramp
						PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*.4, Pan(BOT(b) ), 0, Pitch(BOT(b) )+50000, 1, 0, AudioFade(BOT(b) )
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
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ActiveBall)
End Sub


'******************************
' Diverse Collection Hit Sounds
'******************************


Sub Metals_Medium_Hit(idx)
	StopSound "Plastic_Ramp"
    PlaySound "metalhit_medium", 0, Vol(ActiveBall) /4, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub MetalBallGuide_Hit(idx)
    PlaySound "metalhit_medium", 0, 1, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit(idx)
    PlaySound "gate", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Rubbers_Hit(idx)
    PlaySoundAtBallVol "rubber_hit_" & Int(Rnd*3)+1, 1
End Sub

Sub Posts_Hit(idx)
    PlaySoundAtBallVol "post_" & Int(Rnd*4)+1, 1
End Sub

Sub FlipperL_Collide(parm)
    PlaySoundAtBallVol "flip_hit_" & Int(Rnd*3)+1, 3
End Sub

Sub FlipperR_Collide(parm)
    PlaySoundAtBallVol "flip_hit_" & Int(Rnd*3)+1, 3
End Sub

Sub FlipperR1_Collide(parm)
    PlaySoundAtBallVol "flip_hit_" & Int(Rnd*3)+1, 2
End Sub

Sub SoundNeutralZone()
    Select Case Int(Rnd * 3) + 1
        Case 1:PlaySound "NeutralZone1",0,1,-.2,0,0,0,0,-.7
        Case 2:PlaySound "NeutralZone2",0,1,-.2,0,0,0,0,-.7
        Case 3:PlaySound "Lock",0,1,-.2,0,0,0,0,-.7
    End Select
End Sub



'*****************************************
'	Ball Shadow by Ninuzzo
'*****************************************

Dim BallShadow
BallShadow = Array (BallShadow1, BallShadow2, BallShadow3,Ballshadow4,Ballshadow5,Ballshadow6)

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
		If BOT(b).X < Table1.Width/2 Then
			BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) + 10
		Else
			BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) - 10
		End If
	    ballShadow(b).Y = BOT(b).Y + 20
		If BOT(b).Z > 20 Then
			BallShadow(b).visible = 1
		Else
			BallShadow(b).visible = 0
		End If
	Next

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


Dim NextOrbitHit:NextOrbitHit = 0

Sub WireRampBumps_Hit(idx)
	if BallVel(ActiveBall) > .3 and Timer > NextOrbitHit then
		RandomBump3 .5, Pitch(ActiveBall)+5
		' Schedule the next possible sound time.  This prevents it from rapid-firing noises too much.
		' Lowering these numbers allow more closely-spaced clunks.
		NextOrbitHit = Timer + .5 + (Rnd * .2)
	end if
End Sub

Sub WireLaunchRampBumps_Hit(idx)
	if BallVel(ActiveBall) > .3 and Timer > NextOrbitHit then
		RandomBump .5, Pitch(ActiveBall)+5
		' Schedule the next possible sound time.  This prevents it from rapid-firing noises too much.
		' Lowering these numbers allow more closely-spaced clunks.
		NextOrbitHit = Timer + .2 + (Rnd * .2)
	end if
End Sub

Sub PlasticRampBumps_Hit(idx)
	if BallVel(ActiveBall) > .3 and Timer > NextOrbitHit then
		RandomBump 1, -20000
		' Schedule the next possible sound time.  This prevents it from rapid-firing noises too much.
		' Lowering these numbers allow more closely-spaced clunks.
		NextOrbitHit = Timer + .1 + (Rnd * .2)
	end if
End Sub


'Sub MetalGuideBumps_Hit(idx)
'	if BallVel(ActiveBall) > .3 and Timer > NextOrbitHit then
'		RandomBump2 2, Pitch(ActiveBall)
'		' Schedule the next possible sound time.  This prevents it from rapid-firing noises too much.
'		' Lowering these numbers allow more closely-spaced clunks.
'		NextOrbitHit = Timer + .2 + (Rnd * .2)
'	end if
'End Sub

'Sub MetalWallBumps_Hit(idx)
'	if BallVel(ActiveBall) > .3 and Timer > NextOrbitHit then
'		RandomBump 2, 20000 'Increased pitch to simulate metal wall
'		' Schedule the next possible sound time.  This prevents it from rapid-firing noises too much.
'		' Lowering these numbers allow more closely-spaced clunks.
'		NextOrbitHit = Timer + .2 + (Rnd * .2)
'	end if
'End Sub


' Requires rampbump1 to 7 in Sound Manager
Sub RandomBump(voladj, freq)
	dim BumpSnd:BumpSnd= "rampbump" & CStr(Int(Rnd*7)+1)
		PlaySound BumpSnd, 0, Vol(ActiveBall)*voladj, Pan(ActiveBall), 0, freq, 0, 1, AudioFade(ActiveBall)
End Sub

'' Requires metalguidebump1 to 2 in Sound Manager
'Sub RandomBump2(voladj, freq)
'	dim BumpSnd:BumpSnd= "metalguidebump" & CStr(Int(Rnd*2)+1)
'		PlaySound BumpSnd, 0, Vol(ActiveBall)*voladj, Pan(ActiveBall), 0, freq, 0, 1, AudioFade(ActiveBall)
'End Sub

' Requires WireRampBump1 to 5 in Sound Manager
Sub RandomBump3(voladj, freq)
	dim BumpSnd:BumpSnd= "WireRampBump" & CStr(Int(Rnd*5)+1)
		PlaySound BumpSnd, 0, Vol(ActiveBall)*voladj, Pan(ActiveBall), 0, freq, 0, 1, AudioFade(ActiveBall)
End Sub



' Stop Bump Sounds
Sub BumpSTOP1_Hit ()
dim i:for i=1 to 4:StopSound "WireRampBump" & i:next
NextOrbitHit = Timer + 1
End Sub

Sub BumpSTOP2_Hit ()
dim i:for i=1 to 4:StopSound "WireRampBump" & i:next
NextOrbitHit = Timer + 1
End Sub

Sub BumpSTOP3_Hit ()
dim i:for i=1 to 4:StopSound "WireRampBump" & i:next
NextOrbitHit = Timer + 1
End Sub

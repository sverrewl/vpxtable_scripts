' Medieval Madness - IPDB No. 4032
' © Williams 1997
' VPX recreation by ninuzzu & Tom Tower
' Thanks to all the authors (JPSalas,Dozer,Pinball Ken, Jamin, Macho, Joker, PacDude) who made this table before.
' Thanks to Clark Kent for the pics and the advices
' Thanks to zany for the domes and bumpers
' Thanks to knorr for some sound effects I borrowed from his tables
' Thanks to VPDev Team for the freaking amazing VPX


' Thalamus 2018-07-24
' Table has its own "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2

Option Explicit
Randomize

Dim Ballsize,BallMass
BallSize = 50
BallMass = (Ballsize^3)/125000

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Dim DesktopMode:DesktopMode = Table1.ShowDT
Dim UseVPMDMD:UseVPMDMD = DesktopMode
Const UseVPMModSol = 1

LoadVPM "01560000", "WPC.VBS", 3.50

'********************
'Standard definitions
'********************

dim bsTrough, bsCat, bsMe, bsMo, x

'Const cGameName = "mm_10" 'Williams official rom
'Const cGameName = "mm_109" 'free play only
'Const cGameName="mm_109b" 'unofficial
Const cGameName="mm_109c" 'unofficial profanity rom

Const UseSolenoids = 2
Const UseLamps = 1
Const UseGI = 1
Const UseSync = 0
Const HandleMech = 1

' Standard Sounds
Const SSolenoidOn = "fx_Solenoid"
Const SSolenoidOff = ""
Const SFlipperOn = ""
Const SFlipperOff = ""
Const SCoin = "fx_Coin"

'************************************************************************
'						 INIT TABLE
'************************************************************************

Sub Table1_Init
    vpmInit Me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Medieval Madness (Williams 1997)"
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 1
        .Hidden = DesktopMode
        On Error Resume Next
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
    End With

    On Error Goto 0

    ' Init switches
    Controller.Switch(22) = 1 'close coin door
    Controller.Switch(24) = 0 'always closed

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

    ' Nudging
    vpmNudge.TiltSwitch = 14
    vpmNudge.Sensitivity = 1
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingshot)

    ' Trough
    Set bsTrough = New cvpmTrough
    With bsTrough
        .size = 4
        .initSwitches Array(32, 33, 34, 35)
        .Initexit BallRelease, 960, 6
        .InitEntrySounds "fx_drain", SoundFX("fx_Solenoid",DOFContactors), SoundFX("fx_Solenoid",DOFContactors)
        .InitExitSounds SoundFX("fx_Solenoid",DOFContactors), SoundFX("fx_ballrel",DOFContactors)
        .Balls = 4
    End With

    ' Merlin
    Set bsMe = New cvpmSaucer
    With bsMe
        .InitKicker sw28, 28, 110, 15, 0
        .InitExitVariance 1, 1
        .InitSounds "fx_kicker_enter", SoundFX("fx_Solenoid",DOFContactors), SoundFX("fx_saucer_exit",DOFContactors)
        .InitExitVariance 1, 1 
End With

    ' Catapult
    Set bsCat = New cvpmTrough
    With bsCat
		.size = 1
		.initSwitches Array(38)
		.Initexit sw38a, 0, 45
		.InitExitSounds SoundFX("fx_Solenoid",DOFContactors), SoundFx("fx_scoop_out",DOFContactors)
    End With

    ' Moat
    Set bsMo = New cvpmTrough
    With bsMo
        .size = 4
        .initSwitches Array(36)
        .Initexit sw36k, 280, 2
        .InitExitSounds SoundFX("fx_Solenoid",DOFContactors), SoundFx("fx_popper",DOFContactors)
        .InitExitVariance 1, 1
    End With

	SolCastle 0
	SolMod17 0
	SolMod18 0
	SolMod19 0
	SolMod20 0
	SolMod21 0
	SolMod22 0
	SolMod23 0
	SolMod24 0
	SolMod25 0
	LockPost.IsDropped=1
	LockPostP.IsDropped=1
	TrollP1X.IsDropped = 1
	sw45.IsDropped = 1
	TrollP2X.IsDropped = 1
	sw46.IsDropped = 1
	BW1.isdropped = 1
	BW2.isdropped = 1
	InitOptions
	InitLamps
	UpdateGI 0,0
	UpdateGI 1,0
	UpdateGI 2,0
End Sub

'************************************************************************
'							KEYS
'************************************************************************

Sub Table1_KeyDown(ByVal Keycode)
    If keycode = StartGameKey Then Controller.Switch(13) = 1
    If keycode = PlungerKey Then Controller.Switch(11) = 1
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound SoundFX("fx_nudge",0)
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound SoundFX("fx_nudge",0)
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound SoundFX("fx_nudge",0)
    If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal Keycode)
    If keycode = StartGameKey Then Controller.Switch(13) = 0
    If keycode = PlungerKey Then Controller.Switch(11) = 0
    If vpmKeyUp(keycode) Then Exit Sub
End Sub

Sub Table1_Paused:Controller.Pause = True:End Sub
Sub Table1_unPaused:Controller.Pause = False:End Sub
Sub Table1_exit():Controller.Pause = False:Controller.Stop:End Sub

'************************************************************************
'						 SOLENOIDS
'************************************************************************

SolCallback(1) = "Auto_Plunger"							'AutoPlunger
SolCallback(2) = "SolBallRelease"						'Trough Eject
SolCallback(3) = "bsMo.SolOut"							'Left Popper
SolCallback(4) = "SolCastle"							'Castle	Towers					
SolCallback(5) = "SolCastlegatePow"						'Castle Gate Power
SolCallback(6) = "SolCastlegateHold"					'Castle Gate Hold
SolCallback(7) = "vpmSolSound SoundFX(""fx_Knocker"",DOFKnocker),"
SolCallback(8) = "SolCatapult"							'Catapult
SolCallback(9) = "bsMe.SolOut"							'Right Eject
'SolCallback(10)=""										'Lsling		
'SolCallback(11)=""										'Rsling	
'SolCallback(12)=""										'LBump
'SolCallback(13)=""										'TBump	
'SolCallback(14)=""										'RBump
SolCallback(15)= "SolTowerDivPow"						'Tower Diverter Power	
SolCallback(16)= "SolTowerDivHold"						'Tower Diverter Hold			
SolModcallback(17) = "SolMod17"							'Left Side Low Flasher + Insert Panel
SolModcallback(18) = "SolMod18"							'Left Ramp Flasher + Insert Panel
SolModcallback(19) = "SolMod19"							'Left Side High Flasher + Insert Panel
SolModcallback(20) = "SolMod20"							'Right Side High Flasher + Insert Panel
SolModcallback(21) = "SolMod21"							'Right Ramp Flasher + Insert Panel
SolModcallback(22) = "SolMod22"							'Castle Right Side Flasher + Backpanel
SolModcallback(23) = "SolMod23"							'Right Side Low Flashers
SolModcallback(24) = "SolMod24"							'Moat Flashers
SolModcallback(25) = "SolMod25"							'Castle Left Side Flashers + BackPanel
Solcallback(26) = "SolTowerLock"						'Tower Lock Post
SolCallback(27) = "gate3.open ="						'Right Gate
Solcallback(28) = "gate2.open ="						'Left Gate

Solcallback(33) = "SolLeftTrollPow"						'Left Troll Power
Solcallback(34) = "SolLeftTrollHold"					'Left Troll Hold
Solcallback(35) = "SolRightTrollPow"					'Right Troll Power
Solcallback(36) = "SolRightTrollHold"					'Right Troll Hold
Solcallback(37) = "SolDrawBridge"						'Drawbridge Motor

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

'******************************************************
'				NFOZZY'S FLIPPERS
'******************************************************

Dim returnspeed, lfstep, rfstep
returnspeed = LeftFlipper.return
lfstep = 1
rfstep = 1

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySound SoundFX("fx_FlipperUp",DOFFlippers), 0, 1, -0.2, 0.25
        LeftFlipper.RotateToEnd
    Else
        PlaySound SoundFX("fx_FlipperDown",DOFFlippers), 0, 1, -0.2, 0.25
        LeftFlipper.RotateToStart
        LeftFlipper.TimerEnabled = 1
        LeftFlipper.TimerInterval = 16
        LeftFlipper.return = returnspeed * 0.5
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySound SoundFX("fx_FlipperUp",DOFFlippers), 0, 1, 0.2, 0.25
        RightFlipper.RotateToEnd
    Else
        PlaySound SoundFX("fx_FlipperDown",DOFFlippers), 0, 1, 0.2, 0.25
        RightFlipper.RotateToStart
        RightFlipper.TimerEnabled = 1
        RightFlipper.TimerInterval = 16
        RightFlipper.return = returnspeed * 0.5
    End If
End Sub

Sub LeftFlipper_timer()
	select case lfstep
		Case 1: LeftFlipper.return = returnspeed * 0.6 :lfstep = lfstep + 1
		Case 2: LeftFlipper.return = returnspeed * 0.7 :lfstep = lfstep + 1
		Case 3: LeftFlipper.return = returnspeed * 0.8 :lfstep = lfstep + 1
		Case 4: LeftFlipper.return = returnspeed * 0.9 :lfstep = lfstep + 1
		Case 5: LeftFlipper.return = returnspeed * 1 :lfstep = lfstep + 1
		Case 6: LeftFlipper.timerenabled = 0 : lfstep = 1
	end select
end sub

Sub RightFlipper_timer()
	select case rfstep
		Case 1: RightFlipper.return = returnspeed * 0.6 :rfstep = rfstep + 1
		Case 2: RightFlipper.return = returnspeed * 0.7 :rfstep = rfstep + 1
		Case 3: RightFlipper.return = returnspeed * 0.8 :rfstep = rfstep + 1
		Case 4: RightFlipper.return = returnspeed * 0.9 :rfstep = rfstep + 1
		Case 5: RightFlipper.return = returnspeed * 1 :rfstep = rfstep + 1
		Case 6: RightFlipper.timerenabled = 0 : rfstep = 1
	end select
end sub









'************************************************************************
'						 AUTOPLUNGER 
'************************************************************************

Sub Auto_Plunger(Enabled)
    If Enabled Then
       Plunger.Fire
       PlaySoundAt SoundFX("fx_AutoPlunger",DOFContactors),sw18
	End If
End Sub

'************************************************************************
'						 DRAIN AND RELEASE
'************************************************************************

Sub SolBallRelease(Enabled)
    If Enabled Then
        If bsTrough.Balls Then
            vpmTimer.PulseSw 31
        End If
        bsTrough.ExitSol_On
    End If
End Sub

Sub Drain_Hit:bsTrough.AddBall Me:End Sub


'************************************************************************
'       CASTLE TOWERS
'************************************************************************

Dim vel,per,brake,cnt4on,cnt4off
vel=0:per=0:brake=0

Sub SolCastle(Enabled)
If Enabled Then
	PlaySoundAt SoundFX("fx_AutoPlunger",DOFContactors),sw37
	brake=0
	cnt4off=0
	explosion.enabled=1
	Else
	cnt4on=0
	StopSound "fx_AutoPlunger"
End If
End Sub

Sub explosion_timer()
	If Controller.Solenoid(4) then cnt4on=cnt4on+1
	If NOT Controller.Solenoid(4) then cnt4off=cnt4off+1
	If Controller.Solenoid(4) and cnt4on>15 then 
		LUtower.roty=(-20)*0.75
		LDtower.roty=(-70)*0.75
		URtower.roty=(40)*0.75
		URtower.rotx=(-40)*0.75
		Ctower.roty=(-20)*0.75
		Ctower.rotx=(-20)*0.75
		brake=0.09
		vel=0.9
    End If

	If brake<1 then  vel=vel+0.1:brake=brake+0.01
	If LUtower.roty>0 then per=0:vel=0
	per=(cos(vel)-brake)
	LUtower.roty=LUtower.roty-(per*2)
	LDtower.roty=LDtower.roty-(per*7)
	URtower.roty=URtower.roty+(per*4)
	URtower.rotx=URtower.rotx-(per*4)
	Ctower.roty=Ctower.roty-(per*2)
	Ctower.rotx=Ctower.rotx-(per*2)
If NOT Controller.Solenoid(4) AND cnt4off>100 Then 
	LUtower.roty=0
	LDtower.roty=0
	URtower.roty=0
	URtower.rotx=0
	Ctower.roty=0
	Ctower.rotx=0
	Me.Enabled=0
End If
End Sub

'************************************************************************
'						 CASTLE GATE 
'************************************************************************

Dim IronGateDir

Sub SolCastlegatePow(Enabled)
If Enabled Then
CastleGateTimer.Enabled = 1
IronGateDir = 1
PlaySoundAt SoundFX("fx_gateUp",DOFContactors),sw37
TwrShake
End If
End Sub

Sub SolCastlegateHold(Enabled)
	If Enabled Then        
		DOF 102, DOFPulse
	End If
	If NOT Enabled AND IronGateDir = 1 Then
		IronGateDir = -1
		CastleGateTimer.Enabled = 1
		PlaySoundAt SoundFX("fx_gateDown",DOFContactors),sw37
		DOF 102, DOFPulse
    End If
End Sub

Sub CastleGateTimer_Timer
	DOF 101, DOFOn
	gate.Z = gate.Z + IronGateDir
        If gate.Z <= 15 Then
			door2.IsDropped = 0
			DOF 101, DOFOff
            Me.Enabled = 0
        End If
        If gate.Z >= 115 Then
			door2.IsDropped = 1
			DOF 101, DOFOff
            Me.Enabled = 0
        End If
End Sub

'************************************************************************
'							Catapult 
'************************************************************************

Dim catdir

Sub SolCatapult(enabled)
If enabled Then
bsCat.ExitSol_On
sw38.destroyball
catdir = 1
sw38.TimerEnabled = 1
sw38.TimerInterval = 1
End If
End Sub

Sub sw38_timer()
Prim_Catapult.Rotx = Prim_Catapult.Rotx + catdir
If Prim_Catapult.Rotx >= 97 AND catdir=1 Then catdir = -1
If Prim_Catapult.Rotx <= 7 Then Me.TimerEnabled = 0
End Sub

Sub sw38_Hit:PlaySoundAt "fx_kicker_enter",sw38:bsCat.AddBall Me: sw38.createSizedballWithMass Ballsize/2,BallMass : End Sub

'************************************************************************
'						 TOWER DIVERTER
'************************************************************************

Dim DiverterDir

Sub SolTowerDivPow(Enabled)
If Enabled Then
Diverter.IsDropped=1
Diverter.TimerEnabled = 1
DiverterDir = 1
PlaySoundAt SoundFX("fx_DiverterUp",DOFContactors),DiverterP
End If
End Sub

Sub SolTowerDivHold(Enabled)
	If NOT Enabled AND DiverterDir = 1 Then
		Diverter.IsDropped=0
		DiverterDir = -1
		Diverter.TimerEnabled = 1
		PlaySoundAt SoundFX("fx_DiverterDown",DOFContactors),DiverterP
    End If
End Sub

Sub Diverter_Timer
	DiverterP.Z = DiverterP.Z - 5*DiverterDir
        If DiverterP.Z <= 87 Then Me.TimerEnabled = 0
        If DiverterP.Z >= 137 Then Me.TimerEnabled = 0
End Sub

'************************************************************************
'						 DRAW BRIDGE 
'************************************************************************

Dim dbpos

Sub SolDrawBridge(enabled)
If Enabled AND Controller.GetMech(0)/16 <= 15 then dbridge.enabled = 1 : dbpos = 1:PlaySound SoundFX("Bridge_Move", DOFGear), -1, 0.1, 0, 0, 0, 1, 0:DOF 104, DOFOn
If Enabled AND Controller.GetMech(0)/16 > 15 then dbridge.enabled = 1 : dbpos = 2:PlaySound SoundFX("Bridge_Move", DOFGear), -1, 0.1, 0, 0, 0, 1, 0:DOF 104, DOFOn
End Sub

Sub dbridge_timer()
Select Case dbpos
	Case 1:			'bridge is going down
		drawbridgep.RotX = drawbridgep.Rotx - 1
		DBdecal.rotx=DBdecal.rotx-1
		braket.rotx=braket.rotx-1
		If drawbridgep.RotX <= -90 Then
			DOF 104, DOFOff
			drawbridgep.RotX= -90
			DBdecal.rotx=-90
			braket.rotx=-90
			XXX.enabled = 0
			XXX1.enabled = 0
			Door1.isdropped = 1
			BridgeRamp.collidable = 1
			Ramp22.collidable = 0
			BW1.isdropped = 0
			BW2.isdropped = 0
			Me.Enabled = 0
			StopSound "Bridge_Move"
			PlaysoundAt SoundFX("Bridge_Stop", 0), sw37
		End If

	Case 2:			'bridge is going up
		drawbridgep.RotX = drawbridgep.Rotx + 1
		DBdecal.rotx=DBdecal.rotx+1
		braket.rotx=braket.rotx+1
		If drawbridgep.RotX >= 0 Then
			DOF 104, DOFOff
			drawbridgep.RotX= 0
			DBdecal.rotx=0
			braket.rotx=0
			XXX.enabled = 1
			XXX1.enabled = 1
			Door1.isdropped = 0
			BridgeRamp.collidable = 0
			Ramp22.collidable = 1
			BW1.isdropped = 1
			BW2.isdropped = 1
			Me.Enabled = 0
			StopSound "Bridge_Move"
			PlaysoundAt SoundFX("Bridge_Stop", 0), sw37
		End If
End Select
End Sub


'************************************************************************
'				Drawbridge and Castle Door Shake
'************************************************************************

Sub Door1_Hit()
RandomSoundMetal
If Controller.Switch(56) Then 'solo se sta su
DOF 102, DOFPulse
doorshake.Enabled = 1
TwrShake
End If
End Sub

dim doors,doorsh,doorbrake

sub doorshake_timer()
       if doorbrake<3 then  
        doorbrake=doorbrake+0.1
        doors=doors+0.5
        doorsh=sin(doors)*(3-(doorbrake))
		if gateshake=0 then drawbridgep.RotX = drawbridgep.Rotx +doorsh
		if gateshake=0 then DBdecal.rotx=DBdecal.rotx+doorsh


        if gateshake=1 then gate.transz=doorsh
       end if
'		braket.rotx=braket.rotx+50
       if doorbrake>3  then 
				if gateshake=0 then drawbridgep.RotX = 0
				if gateshake=0 then DBdecal.rotx=0

        if gateshake=1 then gate.transz=0

        doors=0:doorsh=0:doorbrake=0:gateshake=0:Me.Enabled=0 :exit sub
       end if
end sub

dim gateshake

Sub Door2_Hit()
gateshake=1
RandomSoundMetal
doorshake.enabled=1
DOF 102, DOFPulse
Controller.Switch(37) = 1
vpmtimer.addtimer 200, "Controller.Switch(37) = 0'"
TwrShake
End Sub

'************************************************************************
'						TOWERS SHAKE
'************************************************************************

dim vel2,per2,brake2
per2=0:vel2=0:brake2=0

Sub TwrShake
    towersshake.enabled = 1
End Sub

Sub towersshake_timer()

        if brake2 < 1 then vel2=vel2+0.1:brake2=brake2+0.01
        if LUtower.roty>0 then per2=0:vel2=0 
        per2=cos(vel2)-brake2
        
	LUtower.roty=LUtower.roty-(per2*0.2)
	LDtower.roty=LDtower.roty-(per2*0.2)
	URtower.roty=URtower.roty+(per2*0.2)
	URtower.rotx=URtower.rotx-(per2*0.2)
	Ctower.roty=Ctower.roty-(per2*0.2)
	Ctower.rotx=Ctower.rotx-(per2*0.2)

   if brake2>0.9 then me.enabled=0 : per2=0:vel2=0:brake2=0:LUtower.roty=0 :LDtower.roty=0 : URtower.roty=0: Ctower.roty=0:Ctower.rotx=0 : URtower.rotx=0

End Sub

'************************************************************************
' 							Tower Lock
'************************************************************************

Sub SolTowerLock(Enabled)
	StopSound "fx_Postup":PlaySoundAt SoundFX("fx_Postup",DOFContactors),sw58
	LockPost.IsDropped=NOT Enabled
	LockPostP.IsDropped=NOT Enabled
End Sub

Sub sw58_Hit: Controller.Switch(58)=1:End Sub
Sub sw58_UnHit: Controller.Switch(58)=0:vpmTimer.AddTimer 200, "BallDropSound":End Sub

'************************************************************************
'					Slingshots Animation
'************************************************************************

Dim LStep, RStep

Sub LeftSlingShot_Slingshot
	PlaySoundAt SoundFX("LeftSlingshot",DOFContactors),sling1
	vpmTimer.PulseSw 51
	LSling.Visible = 0
	LSling1.Visible = 1
	sling1.TransZ = -20
	LStep = 0
	LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling1.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub


Sub RightSlingShot_Slingshot
	PlaySoundAt SoundFX("RightSlingshot",DOFContactors),SLING2
	vpmTimer.PulseSw 52
	RSling.Visible = 0
	RSling1.Visible = 1
	sling2.TransZ = -20
	RStep = 0
	RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling2.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

'************************************************************************
'						Bumpers Animation
'************************************************************************

Dim dirRing1:dirRing1 = -1
Dim dirRing2:dirRing2 = -1
Dim dirRing3:dirRing3 = -1

Sub Bumper1_Hit:vpmTimer.PulseSw 53:PlaySoundAt SoundFX("LeftBumper_Hit",DOFContactors),Bumper1:Me.TimerEnabled = 1:me.timerinterval = 10:End Sub

Sub Bumper2_Hit:vpmTimer.PulseSw 54:PlaySoundAt SoundFX("RightBumper_Hit",DOFContactors),Bumper2:Me.TimerEnabled = 1:me.timerinterval = 10:End Sub

Sub Bumper3_Hit:vpmTimer.PulseSw 55:PlaySoundAt SoundFX("TopBumper_Hit",DOFContactors),Bumper3:Me.TimerEnabled = 1:me.timerinterval = 10:End Sub

Sub Bumper1_timer()
	BumperRing1.Z = BumperRing1.Z + (5 * dirRing1)
	If BumperRing1.Z <= -35 Then dirRing1 = 1
	If BumperRing1.Z >= 0 Then
		dirRing1 = -1
		BumperRing1.Z = 0
		Me.TimerEnabled = 0
	End If
End Sub

Sub Bumper2_timer()
	BumperRing2.Z = BumperRing2.Z + (5 * dirRing2)
	If BumperRing2.Z <= -35 Then dirRing2 = 1
	If BumperRing2.Z >= 0 Then
		dirRing2 = -1
		BumperRing2.Z = 0
		Me.TimerEnabled = 0
	End If
End Sub

Sub Bumper3_timer()
	BumperRing3.Z = BumperRing3.Z + (5 * dirRing3)
	If BumperRing3.Z <= -35 Then dirRing3 = 1
	If BumperRing3.Z >= 0 Then
		dirRing3 = -1
		BumperRing3.Z = 0
		Me.TimerEnabled = 0
	End If
End Sub

'************************************************************************
'								Trolls 
'************************************************************************

Sub SolLeftTrollPow(Enabled)
    If Enabled then
		PlaySound SoundFX("fx_TrollUp",DOFContactors), 0, 1, -.01, 0, 0, 0, 1
        ltroll.Z = -8
        liftL.Z = ltroll.Z
        TrollP1X.IsDropped = 0
        sw45.IsDropped = 0
        Controller.Switch(74) = 1
		LTP.Enabled=1
    End If
	If NOT Enabled AND ltroll.Z = -8 Then
		TrollP1X.TimerEnabled=1
		TrollP1X.TimerInterval=10
	End If
End Sub

Sub SolLeftTrollHold(Enabled)
	If NOT Enabled AND ltroll.Z = -8 Then
		ltroll.Z = -104
		liftL.Z = ltroll.z
		TrollP1X.IsDropped = 1
		sw45.IsDropped = 1
		Controller.Switch(74) = 0
		PlaySound SoundFX("fx_TrollDown",DOFContactors), 0, 1, -.01, 0, 0, 0, 1
		LTP.Enabled=0
    End If
End Sub

Sub SolRightTrollPow(Enabled)
    If Enabled then
		PlaySound SoundFX("fx_TrollUp",DOFContactors), 0, 1, .01, 0, 0, 0, 1
        rtroll.Z = -8
        liftR.Z = rtroll.Z
        TrollP2X.IsDropped = 0
        sw46.IsDropped = 0
        Controller.Switch(75) = 1
		RTP.Enabled=1
    End If
	If NOT Enabled AND rtroll.Z = -8 Then
		TrollP2X.TimerEnabled=1
		TrollP2X.TimerInterval=10
    End If
End Sub

Sub SolRightTrollHold(Enabled)
	If NOT Enabled AND rtroll.Z = -8 Then
        rtroll.Z = -104
        liftR.Z = rtroll.Z
        TrollP2X.IsDropped = 1
        sw46.IsDropped = 1
        Controller.Switch(75) = 0
		PlaySound SoundFX("fx_TrollDown",DOFContactors), 0, 1, .01, 0, 0, 0, 1
		RTP.Enabled=0
    End If
End Sub

dim lshake,rshake
lshake=0:rshake=0

Sub TrollP1X_timer
	lshake=lshake+1
    If lshake<12 then 
		LiftL.transz=1 + Sin(lshake)
		ltroll.transz=1 + Sin(lshake)
	Else
		me.TimerEnabled=0
		lshake=0
		LiftL.transz=0
	End If
End Sub

Sub TrollP2X_timer
	rshake=rshake+1
    if rshake<12 then 
		LiftR.transz=1 + Sin(rshake)
		rtroll.transz=1 + Sin(rshake)
	Else 
		me.TimerEnabled=0
		rshake=0
		LiftR.transz=0
	End If
End Sub

'Shake Trolls when hit

dim sbou,sbou2

Sub sw45_Hit():PlaySound "fx_woodhit",0,5:vpmTimer.PulseSw 45:Me.TimerEnabled = 1:sbou=ActiveBall.vely/4 :End Sub
Sub sw46_Hit():PlaySound "fx_woodhit",0,5:vpmTimer.PulseSw 46:Me.TimerEnabled = 1:sbou2=ActiveBall.vely/4 :End Sub

dim bou,brakeTl,perc
perc=1
Sub sw45_Timer()

    bou=bou+0.3:brakeTl=brakeTl+0.2
    if (perc-(brakeTl*(perc/6)))<0 then Me.TimerEnabled = 0 :bou=0 :brakeTl=0 :perc=5
    ltroll.rotx=sin(bou)*(perc-(brakeTl*(perc/6)))
    ltroll.roty=sin((bou)*sbou)*(perc-(brakeTl*(perc/6)))

    LiftL.rotx=sin(bou)*(perc-(brakeTl*(perc/6)))
    LiftL.roty=sin((bou)*sbou)*(perc-(brakeTl*(perc/6)))

end sub 

dim bou2,brakeTr,perc2
perc2=1
Sub sw46_Timer()

    bou2=bou2+0.3:brakeTr=brakeTr+0.2
    if (perc2-(brakeTr*(perc2/6)))<0 then Me.TimerEnabled = 0 :bou2=0 :brakeTr=0 :perc2=5
    rtroll.rotx=sin(bou2)*(perc2-(brakeTr*(perc2/10)))
    rtroll.roty=sin((bou2)*sbou2)*(perc2-(brakeTr*(perc2/6)))

    liftR.rotx=sin(bou2)*(perc2-(brakeTr*(perc2/10)))
    liftR.roty=sin((bou2)*sbou2)*(perc2-(brakeTr*(perc2/6)))

end sub

Sub LTP_Hit
	activeball.z = activeball.z + 50
	PlaySoundAt "fx_BallDrop",LTP
End Sub

Sub RTP_Hit
	activeball.z = activeball.z + 50
	PlaySoundAt "fx_BallDrop",RTP
End Sub

'************************************************************************
'					Switches
'************************************************************************
' Eject holes

Sub sw28_Hit:bsMe.AddBall Me:End Sub
Sub sw36_Hit:PlaySoundAtBall "fx_Moat_enter":bsMo.AddBall Me:End Sub

' Lanes
Sub sw66_Hit:Controller.Switch(66) = 1:PlaySoundAt "fx_sensor",sw66:End Sub
Sub sw66_UnHit:Controller.Switch(66) = 0:End Sub

Sub sw67_Hit:Controller.Switch(67) = 1:PlaySoundAt "fx_sensor",sw67:End Sub
Sub sw67_UnHit:Controller.Switch(67) = 0:End Sub

Sub sw18_Hit:Controller.Switch(18) = 1:PlaySoundAt "fx_sensor",sw18:End Sub
Sub sw18_UnHit:Controller.Switch(18) = 0:End Sub

Sub sw16_Hit:Controller.Switch(16) = 1:PlaySoundAt "fx_sensor",sw16:End Sub
Sub sw16_UnHit:Controller.Switch(16) = 0:End Sub

Sub sw26_Hit
Controller.Switch(26) = 1
If ActiveBall.VelY => 3.5 Then
ActiveBall.VelY = 3.5
End If
If ActiveBall.VelY <= -10 Then
ActiveBall.VelY = -6
End If
PlaySoundAt "fx_sensor",sw26
End Sub

Sub sw26_UnHit:Controller.Switch(26) = 0:End Sub

Sub sw17_Hit
Controller.Switch(17) = 1
If ActiveBall.VelY => 3.5 Then
ActiveBall.VelY = 3.5
End If
If ActiveBall.VelY <= -10 Then
ActiveBall.VelY = -6
End If
PlaySoundAt "fx_sensor",sw17
End Sub

Sub sw17_UnHit:Controller.Switch(17) = 0:End Sub

Sub sw27_Hit:Controller.Switch(27) = 1:PlaySoundAt "fx_sensor",sw27:End Sub
Sub sw27_UnHit:Controller.Switch(27) = 0:End Sub

Sub sw65_Hit:Controller.Switch(65) = 1:PlaySoundAt "fx_sensor",sw65:End Sub
Sub sw65_UnHit:Controller.Switch(65) = 0:End Sub

Sub sw47_Hit:Controller.Switch(47) = 1:PlaySoundAt "fx_sensor",sw47:End Sub
Sub sw47_UnHit:Controller.Switch(47) = 0:End Sub

Sub sw48_Hit:Controller.Switch(48) = 1:PlaySoundAt "fx_sensor",sw48:End Sub
Sub sw48_UnHit:Controller.Switch(48) = 0:End Sub

Sub sw68_Hit:Controller.Switch(68) = 1:PlaySoundAt "fx_sensor",sw68:End Sub
Sub sw68_UnHit:Controller.Switch(68) = 0:End Sub

' Targets
Sub sw12_Hit:vpmTimer.PulseSw 12:PlaySoundAt SoundFX("fx_target",DOFTargets),sw12:End Sub
Sub sw15_Hit:vpmTimer.PulseSw 15:PlaySoundAt SoundFX("fx_target",DOFTargets),sw15:End Sub
Sub sw25_Hit:vpmTimer.PulseSw 25:PlaySoundAt SoundFX("fx_target",DOFTargets),sw25:End Sub
Sub sw71_Hit:vpmTimer.PulseSw 71:PlaySoundAt SoundFX("fx_target",DOFTargets),sw71:End Sub
Sub sw72_Hit:vpmTimer.PulseSw 72:PlaySoundAt SoundFX("fx_target",DOFTargets),sw72:End Sub
Sub sw73_Hit:vpmTimer.PulseSw 73:PlaySoundAt SoundFX("fx_target",DOFTargets),sw73:End Sub

' Triggers on the ramps
Sub sw61_Hit:Controller.Switch(61) = 1:End Sub
Sub sw61_Unhit:Controller.Switch(61) = 0:End Sub

Sub sw62_Hit:Controller.Switch(62) = 1:sw62flip.rotatetoend:PlaySoundAt "fx_gate4", sw62:End Sub
Sub sw62_Unhit:Controller.Switch(62) = 0:sw62flip.rotatetostart:End Sub

Sub sw63_Hit:Controller.Switch(63) = 1:End Sub
Sub sw63_Unhit:Controller.Switch(63) = 0:End Sub

Sub sw64_Hit:Controller.Switch(64) = 1:sw64flip.rotatetoend:PlaySoundAt "fx_gate4", sw64:End Sub
Sub sw64_Unhit:Controller.Switch(64) = 0:sw64flip.rotatetostart:End Sub

' Gates
Sub Gate1_Hit():PlaySoundAt "fx_gate",Gate1:End Sub
Sub Gate2_Hit():PlaySoundAt "fx_gate",Gate2:End Sub
Sub Gate3_Hit():PlaySoundAt "fx_gate",Gate3:End Sub
Sub Gate9_Hit():PlaySoundAt "fx_gate",Gate9:End Sub
Sub Gate6_Hit():PlaySoundAt "fx_gate",Gate6:End Sub

' Opto Switches
Sub sw41_hit
vpmtimer.pulsesw 41
End sub

Sub sw37_Hit():Controller.Switch(37) = 1
vpmtimer.addtimer 100, "Controller.Switch(37) = 0'"
End Sub

' Lock_Door
dim rotat,brakew

Sub GateT_Hit
PlaySoundAt "fx_gate", GateT
rotat=0
brakew=1
me.TimerEnabled=0
me.TimerEnabled=1
me.TimerInterval=7
End Sub

Sub GateT_timer()
	If brakew <7 then rotat=rotat+0.07:brakew=brakew+0.003 else Lock_Door.rotX=0:me.TimerEnabled = 0
	Lock_Door.rotX=(cos(rotat+90)*100)/brakew^3
End Sub

Sub sw44_Hit():vpmtimer.PulseSw 44:PlaySoundAt "fx_gate",sw44: End Sub

'************************************************************************
'						Modulated Flashers
'************************************************************************

Dim xxst, xxet, xxnt, xxtw, xxto, xxtt, xxt3, xxmf, xxtf

Sub SolMod17(value)
If value > 0 Then
For each xxst in BL_Flash:xxst.State = 1:next
Else
For each xxst in BL_Flash:xxst.State = 0:next
End If
For each xxst in BL_Flash:xxst.IntensityScale = value/160:next
F_lt1.color=RGB(255,0,0)			'troll reflection
F_lt2.color=RGB(255,0,0)			'troll reflection
f17b.IntensityScale=value/160		'ambient reflection
f17c.IntensityScale=value/160		'dome lit
f17d.IntensityScale=value/160		'dome lit
If ltroll.Z > -50 Then
F_lt1.IntensityScale=value/160
F_lt2.IntensityScale=value/160
Else
F_lt1.IntensityScale=0
F_lt2.IntensityScale=0
End If
End Sub

Sub SolMod18(value)
If value > 0 Then
For each xxet in LR_Flash:xxet.State = 1:next
Else
For each xxet in LR_Flash:xxet.State = 0:next
End If
For each xxet in LR_Flash:xxet.IntensityScale = value/160:next
F_lt1.color=RGB(255,200,145)			'troll reflection
F_lt2.color=RGB(255,200,145)			'troll reflection
f18.intensityscale=value/160			'castle reflection
F18a.intensityscale=value/160			'merlin decal
If ltroll.Z > -50 Then
F_lt1.IntensityScale=value/160
F_lt2.IntensityScale=value/160
Else
F_lt1.IntensityScale=0
F_lt2.IntensityScale=0
End If
End Sub

Sub SolMod19(value)
If value > 0 Then
For each xxnt in TL_Flash:xxnt.State = 1:next
Else
For each xxnt in TL_Flash:xxnt.State = 0:next
End If
For each xxnt in TL_Flash:xxnt.IntensityScale = value/160:next
For each xxnt in Castle_LF:xxnt.IntensityScale = value/160:next
End Sub

Sub SolMod20(value)
If value > 0 Then
For each xxtw in TR_Flash:xxtw.State = 1:next
Else
For each xxtw in TR_Flash:xxtw.State = 0:next
End If
For each xxtw in TR_Flash:xxtw.IntensityScale = value/160:next
For each xxtw in Castle_RF:xxtw.IntensityScale = value/160:next
End Sub

Sub SolMod21(value)
If value > 0 Then
For each xxto in RR_Flash:xxto.State = 1:next
Else
For each xxto in RR_Flash:xxto.State = 0:next
End If
For each xxto in RR_Flash:xxto.IntensityScale = value/160:next
F_rt1.color=RGB(255,200,145)			'troll reflection
F_rt2.color=RGB(255,200,145)			'troll reflection
f21.intensityscale=value/160			'castle reflection
f21a.intensityscale=value/160			'merlin decal
f21b.intensityscale=value/160			'dragon
f21c.intensityscale=value/160			'dragon
f21d.intensityscale=value/160			'dragon
If rtroll.Z > -50 Then
F_rt1.IntensityScale=value/160
F_rt2.IntensityScale=value/160
Else
F_rt1.IntensityScale=0
F_rt2.IntensityScale=0
End If
End Sub

Sub SolMod22(value)
If value > 0 Then
For each xxtt in CR_Flash:xxtt.State = 1:next
Else
For each xxtt in CR_Flash:xxtt.State = 0:next
End If
For each xxtt in CR_Flash:xxtt.IntensityScale = value/160:next
F22.IntensityScale=value/160			'castle reflection
F22a.IntensityScale=value/160			'dome lit
F22b.IntensityScale=value/160			'dome lit
End Sub

Sub SolMod23(value)
If value > 0 Then
For each xxt3 in BR_Flash:xxt3.State = 1:next
Else
For each xxt3 in BR_Flash:xxt3.State = 0:next
End If
For each xxt3 in BR_Flash:xxt3.IntensityScale=value/160:next
F_rt1.color=RGB(255,0,0)				'troll reflection
F_rt2.color=RGB(255,0,0)				'troll reflection
f23.IntensityScale=value/160			'dragon
f23a.IntensityScale=value/160			'dome lit
f23b.IntensityScale=value/160			'dome lit
f23c.IntensityScale=value/160			'dome lit
If rtroll.Z > -50 Then
F_rt1.IntensityScale=value/160
F_rt2.IntensityScale=value/160
Else
F_rt1.IntensityScale=0
F_rt2.IntensityScale=0
End If
End Sub

Sub SolMod24(value)
If value > 0 Then
For each xxmf in Moat_Flash:xxmf.State = 1:next
Else
For each xxmf in Moat_Flash:xxmf.State = 0:next
End If
For each xxmf in Moat_Flash:xxmf.IntensityScale=value/160:next
F24.IntensityScale=value/160			'castle reflection
F24a.IntensityScale=value/160			'merlin decal
End Sub

Sub SolMod25(value)
If value > 0 Then
For each xxtf in CL_Flash:xxtf.State = 1:next
Else
For each xxtf in CL_Flash:xxtf.State = 0:next
End If
For each xxtf in CL_Flash:xxtf.IntensityScale=value/160:next
F25.IntensityScale=value/160			'castle reflection
F25a.IntensityScale=value/160			'dome lit
F25b.IntensityScale=value/160			'dome lit
End Sub

'**************
' Inserts
'**************

Sub InitLamps
On Error Resume Next
Dim i
For i=0 To 127: Execute "Set Lights(" & i & ")  = L" & i: Next
Lights(14)=Array(L14,L14a)
Lights(64)=Array(L64,L64a)
Lights(78)=Array(L78,L78a)
End Sub

Sub SynchFlasherObj
L11a.IntensityScale=L11.state
L12a.IntensityScale=L12.state
L13a.IntensityScale=L13.state
L15a.IntensityScale=L15.state
L55a.IntensityScale=L55.state
L56a.IntensityScale=L56.state
L57a.IntensityScale=L57.state
L58a.IntensityScale=L58.state
L81a.IntensityScale=L81.state
L82a.IntensityScale=L82.state
L83a.IntensityScale=L83.state
L84a.IntensityScale=L84.state
End Sub

'**************
' 8-step GI
'**************

Set GiCallback2 = GetRef("UpdateGI")

Dim gistep,xx
gistep = 1/8

Sub UpdateGI(no, step)
Select Case no  
Case 0 'bottom
	If step = 0 Then
		For each xx in GIB:xx.State = 0:Next
	Else
		For each xx in GIB:xx.State = 1:Next
	End If
	For each xx in GIB:xx.IntensityScale = gistep * step:next
Case 1 'middle
	If step = 0 Then
		DOF 103, DOFOff
		For each xx in GIM:xx.State = 0:Next
	Else
		DOF 103, DOFOn
		For each xx in GIM:xx.State = 1:Next
	End If
	For each xx in GIM:xx.IntensityScale = gistep * step:next
Case 2 'top
	If step = 0 Then
		For each xx in GIT:xx.State = 0:Next
		For each xx in bump1:xx.State = 0:Next
		For each xx in bump2:xx.State = 0:Next
		For each xx in bump3:xx.State = 0:Next
	Else
		For each xx in GIT:xx.State = 1:Next
		For each xx in bump1:xx.State = 1:Next
		For each xx in bump2:xx.State = 1:Next
		For each xx in bump3:xx.State = 1:Next
	End If
	If step>4 then
		Prim_Spot1.image= "spot_map (black version)on"
		Prim_Spot2.image= "spot_map (black version)on"
	Else
		Prim_Spot1.image= "spot_map (black version)"
		Prim_Spot2.image= "spot_map (black version)"
	End If
For each xx in CastleGI:xx.IntensityScale = gistep * step:next
For each xx in GIT:xx.IntensityScale = gistep * step:next
For each xx in bump1:xx.IntensityScale = gistep * step:next
For each xx in bump2:xx.IntensityScale = gistep * step:next
For each xx in bump3:xx.IntensityScale = gistep * step:next
End Select
End Sub

'******************
' RealTime Updates
'******************

Set MotorCallback = GetRef("GameTimer")

Sub GameTimer
UpdateMechs
RollingSoundUpdate
BallShadowUpdate
SynchFlasherObj
UpdateMechs
End Sub

Sub UpdateMechs
Gate2P.RotX=Gate2.currentangle
Gate3P.RotX=Gate3.currentangle
WireGateRSmall.RotX=Gate9.currentangle
WireGateR.RotX=Spinner1.currentangle
WireGateLR.RotX=Spinner2.currentangle
WireGateRR.RotX=Spinner3.currentangle
WireGateMerlin.RotX=Gate6.currentangle
WireGateCatapult.RotY=Gate1.currentangle
sw62p.rotY = sw62flip.currentangle
sw64p.rotY = sw64flip.currentangle
	FlipperL.RotZ = LeftFlipper.CurrentAngle
	FlipperR.RotZ = RightFlipper.CurrentAngle
End Sub

'**************
' Ramp Sounds
'**************

Sub BallDropSound(dummy):PlaySoundAt "fx_BallDrop",sw41:End Sub
Sub BallDropSoundL(dummy):PlaySoundAt "fx_BallDrop",LHD:End Sub
Sub BallDropSoundR(dummy):PlaySoundAt "fx_BallDrop",RHD:End Sub
Sub BallDropSoundTopL(dummy):PlaySoundAt "fx_BallDrop",Trigger8:End Sub

Sub ShooterEnd_Hit
If ActiveBall.Z > 30  Then			'ball is flying
vpmTimer.AddTimer 100, "BallDropSound"
End If
End Sub

Sub LHP1_Hit()
If ActiveBall.velY < 0  Then		'ball is going up
PlaySoundAt "fx_rrenterL",LHP1
Else
StopSound "fx_rrenter"
End If
End Sub

Sub RHP1_Hit()
If ActiveBall.velY < 0  Then		'ball is going up
PlaySoundAt "fx_rrenter",RHP1
Else
StopSound "fx_rrenter"
End If
End Sub

Sub LHP2_Hit()
PlaySoundAt "fx_lr2",LHP2
End Sub

Sub RHP2_Hit()
PlaySoundAt "fx_lr2",RHP2
End Sub

Sub LHM_Hit()
PlaySoundAt "fx_metalrolling",LTP
StopSound "fx_rrenter"
End Sub

Sub RHM_Hit()
PlaySoundAt "fx_metalrolling",RHM
StopSound "fx_rrenter"
End Sub

Sub CHM_Hit()
PlaySoundAt "fx_metalrolling",CHM
End Sub

Sub LHD_Hit()
	StopSound "fx_metalrolling"
	vpmtimer.addtimer 200, "BallDropSoundL"
End Sub

Sub RHD_Hit()
	StopSound "fx_metalrolling"
	vpmtimer.addtimer 200, "BallDropSoundR"
End Sub

Sub Trigger8_Hit(): StopSound "fx_metalrolling":vpmtimer.addtimer 250, "BallDropSoundTopL":End Sub

Sub Trigger4_Hit():RandomSoundMetal:End Sub

Sub LHelp_Hit: ActiveBall.VelY = ActiveBall.VelY * 0.8: End Sub

' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 /200)
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

'*********************************************************************
'                 Positional Sound Playback Functions
'*********************************************************************

Sub PlaySoundAt(sound, tableobj)
	If Table1.VersionMinor > 3 OR Table1.VersionMajor > 10 Then
		PlaySound sound, 1, 1, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
	Else
		PlaySound sound, 1, 1, Pan(tableobj)
	End If
End Sub

Sub PlaySoundAtBall(sound)
	If Table1.VersionMinor > 3 OR Table1.VersionMajor > 10 Then
		PlaySound sound, 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
	Else
		PlaySound sound, 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1
	End If
End Sub

Sub PlaySoundAtVol(sound, tableobj, Vol)
	If Table1.VersionMinor > 3 OR Table1.VersionMajor > 10 Then
		PlaySound sound, 1, Vol, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
	Else
		PlaySound sound, 1, Vol, Pan(tableobj)
	End If
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

Sub RollingSoundUpdate()
    Dim BOT, b
    BOT = GetBalls

	' stop the sound of deleted balls
    For b = UBound(BOT) + 1 to tnob
        rolling(b) = False
        StopSound("fx_ballrolling" & b)
    Next

	' play the rolling sound for each ball
    For b = 0 to UBound(BOT)
        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
            rolling(b) = True
        If Table1.VersionMinor > 3 OR Table1.VersionMajor > 10 Then
				PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
			Else		
				PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0
            End If
            Else
                If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If
    Next
End Sub

'*****************************************
'	Ball Shadow
'*****************************************

Dim BallShadow
BallShadow = Array (BallShadow1, BallShadow2, BallShadow3, BallShadow4, BallShadow5)


Sub BallShadowUpdate()
    Dim BOT, b
    BOT = GetBalls

	' render the shadow for each ball
    For b = 0 to UBound(BOT)
		If BOT(b).X < Table1.Width/2 Then
			BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) + 10
		Else
			BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) - 10
		End If

			BallShadow(b).Y = BOT(b).Y + 10
			BallShadow(b).Z = 1
		If BOT(b).Z > 20 Then
			BallShadow(b).visible = 1
		Else
			BallShadow(b).visible = 0
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

'**********************
' Other Sound FX
'**********************

Sub MoatK_Hit(idx)
	RandomSoundHole()
End Sub

Sub Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then 
		PlaySoundAtBall "fx_rubber"
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub Posts_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 16 then 
		PlaySoundAtBall "fx_rubber"
	End if
	If finalspeed >= 6 AND finalspeed <= 16 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySoundAtBall "fx_rubber_hit_1"
		Case 2 : PlaySoundAtBall "fx_rubber_hit_2"
		Case 3 : PlaySoundAtBall "fx_rubber_hit_3"
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
		Case 1 : PlaySoundAtBall "fx_flip_hit_1"
		Case 2 : PlaySoundAtBall "fx_flip_hit_2"
		Case 3 : PlaySoundAtBall "fx_flip_hit_3"
	End Select
End Sub

Sub RandomSoundHole()
	Select Case Int(Rnd*4)+1
		Case 1 : PlaySoundAtBall "fx_Hole1"
		Case 2 : PlaySoundAtBall "fx_Hole2"
		Case 3 : PlaySoundAtBall "fx_Hole3"
		Case 4 : PlaySoundAtBall "fx_Hole4"
	End Select
End Sub

Sub RandomSoundMetal()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySoundAtBall "fx_metal_hit_1"
		Case 2 : PlaySoundAtBall "fx_metal_hit_2"
		Case 3 : PlaySoundAtBall "fx_metal_hit_3"
	End Select
End Sub
'**********************OPTIONS***************************

Sub InitOptions
Ramp15.visible = DesktopMode
Ramp16.visible = DesktopMode
End Sub

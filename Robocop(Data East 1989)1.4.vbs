'==============================================================================================='
'                                               																								'
'        	                 		       Robocop	                       			                			'
'          	              		        Data East (1989)                     	 		                '
'		  				  	  www.ipdb.org/machine.cgi?id=1976                                            '
'																							                                                 	'
' 	  		 	 	  	Created by ICPjuggla, Talantyyr, Dozer316 and dark 		                        '
'																								                                                '
'==============================================================================================='
Option Explicit
Randomize


' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' Wob 2018-08-08
' Added vpmInit Me to table init and both cSingleLFlip and cSingleRFlip
' Thalamus 2018-11-01 : Improved directional sounds
' !! NOTE : Table not verified yet !!

' Options
' Volume devided by - lower gets higher sound

Const VolDiv = 2000    ' Lower number, louder ballrolling/collition sound
Const VolCol = 10      ' Ball collition divider ( voldiv/volcol )

' The rest of the values are multipliers
'
'  .5 = lower volume
' 1.5 = higher volume

Const VolBump   = 2    ' Bumpers volume.
Const VolRol    = 1    ' Rollovers volume.
Const VolGates  = 1    ' Gates volume.
Const VolMetal  = 1    ' Metals volume.
Const VolRB     = 1    ' Rubber bands volume.
Const VolRH     = 1    ' Rubber hits volume.
Const VolPo     = 1    ' Rubber posts volume.
Const VolPi     = 1    ' Rubber pins volume.
Const VolPlast  = 1    ' Plastics volume.
Const VolTarg   = 1    ' Targets volume.
Const VolWood   = 1    ' Woods volume.
Const VolKick   = 1    ' Kicker volume.
Const VolSpin   = 1.5  ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Dim DesktopMode:DesktopMode = Table1.ShowDT

 LoadVPM "01560000", "DE.VBS", 3.46

 dim bsTrough, bsLE, bsTP, bsRI, bsBP, cbCaptive

 Const cGameName="robo_a34"

 Const UseSolenoids=2
 ' Wob: Added for Fast Flips (No upper Flippers)
 Const cSingleLFlip = 0
 Const cSingleRFlip = 0
 Const UseLamps=1
 Const UseGI=0
 Const UseSync=0
 Const HandleMech=0

 ' Standard Sounds
 Const SSolenoidOn="Solenoid"
 Const SSolenoidOff=""
 Const SFlipperOn="FlipperUp"
 Const SFlipperOff="FlipperDown"
 Const SCoin="Coin"

 '************
 ' Table init.
 '************

 Sub Table1_Init
    vpmInit Me
    With Controller
       .GameName=cGameName
       If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
       .SplashInfoLine="Robocop (Data East 1990)"
       .HandleKeyboard=0
       .ShowTitle=0
       .ShowDMDOnly=1
       .ShowFrame=0
       .HandleMechanics=0
        If DesktopMode then
       .Hidden = 1
		else
        If B2SOn then
       .Hidden = 1
        else
       .Hidden = 0
		End If
        End If
       '.SetDisplayPosition 1600, 1600, GetPlayerHWnd
       On Error Resume Next
       .Run GetPlayerHWnd
       If Err Then MsgBox Err.Description
    End With

    On Error Goto 0

    ' Nudging
    vpmNudge.TiltSwitch=swTilt
    vpmNudge.Sensitivity=3
    vpmNudge.TiltObj=Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingshot)

     ' Trough
    Set bsTrough=New cvpmBallStack
    bsTrough.InitSw 10, 13, 12, 11, 0, 0, 0, 0
    bsTrough.InitKick BallRelease, 80, 6
    'bsTrough.InitEntrySnd "Solenoid", "Solenoid"
    bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("fx_Solenoid",DOFContactors)
    bsTrough.Balls=3

    ' Left Eject
    Set bsLE=New cvpmBallStack
    bsLE.InitSaucer sw28, 28, 0, 22
    bsLE.InitExitSnd SoundFX("popper_ball",DOFContactors), SoundFX("popper_ball",DOFContactors)

    ' Top Eject
    Set bsTP=New cvpmBallStack
    bsTP.InitSaucer sw29, 29, 115, 4
    bsTP.InitExitSnd SoundFX("popper_ball",DOFContactors), SoundFX("popper_ball",DOFContactors)
    bsTP.KickAngleVar=1
    bsTP.KickForceVar=1

    ' Right Eject
    Set bsRI=New cvpmBallStack
    bsRI.InitSaucer sw30, 30, 225, 4
    bsRI.InitExitSnd SoundFX("popper_ball",DOFContactors), SoundFX("popper_ball",DOFContactors)
    bsRI.KickAngleVar=1
    bsRI.KickForceVar=1

    ' Ball Popper (VUK)
    Set bsBP=New cvpmBallStack
    bsBP.InitSw 0, 45, 0, 0, 0, 0, 0, 0
    bsBP.InitKick sw45a, 140, 10
    bsBP.InitExitSnd SoundFX("fx_solenoid",DOFContactors), SoundFX("fx_solenoid",DOFContactors)


    ' Main Timer init
    PinMAMETimer.Interval=PinMAMEInterval
    PinMAMETimer.Enabled=1
	Leds.Enabled=1

    ' Init Kickback
    KickBack.Pullback


	CapBall1.createball
	CapBall1.kick 180,1

    For each xx in GI:xx.State = 1: Next
    LFLogo.Image = "flipper-l2"
	RFLogo.Image = "flipper-r2"

 End Sub

 Sub Table1_Paused : Controller.Pause=True : End Sub
 Sub Table1_unPaused : Controller.Pause=False : End Sub

Dim jdxx

 If Table1.ShowDT = true then
 For each jdxx in BDLights:jdxx.IntensityScale = 7:next
Ramp19.visible = 1
Ramp20.visible = 1
Primitive221.visible = 1
 else
 For each jdxx in BDLights:jdxx.IntensityScale = 0:next
Ramp19.visible = 0
Ramp20.visible = 0
Primitive221.visible = 0
 End If
 '**********
 ' Keys
 '**********

 Sub Table1_KeyDown(ByVal Keycode)
    If keycode=RightFlipperKey Then Controller.Switch(15)=1
    If keycode=LeftFlipperKey Then Controller.Switch(16)=1
    If vpmKeyDown(keycode)Then Exit Sub
    If keycode=PlungerKey Then Plunger.Pullback : PlaySoundAtVol "plungerpull", Plunger, 1
 End Sub

 Sub Table1_KeyUp(ByVal Keycode)
    If keycode=RightFlipperKey Then Controller.Switch(15)=0
    If keycode=LeftFlipperKey Then Controller.Switch(16)=0
    If vpmKeyUp(keycode)Then Exit Sub
    If keycode=PlungerKey Then Plunger.Fire : PlaySoundAtVol "plunger", Plunger, 1
 End Sub

'Solenoids

 '*********
 'Solenoids
 '*********
 SolCallback(1)="Flash62"
 SolCallback(2)="Flash61"
 SolCallback(3)="Flash60"
 SolCallback(4)="Flash59"
 SolCallback(5)="Flash58"
 SolCallback(6)="Flash57"
 SolCallback(7)="Sol7"
 SolCallback(8)="Sol8"
 SolCallback(9)="Flash9"
 Solcallback(11)="SolGIUpdate"
 SolCallback(12)="Flash12"
 SolCallback(13)="Flash13"
 SolCallback(14)="Sol14"
 SolCallback(15)="Sol15"
 SolCallback(16)="SolKickback"
 SolCallback(25)="bsTrough.SolIn"
 SolCallback(26)="bsTrough.SolOut"
 SolCallback(27)="bsLE.SolOut"
 SolCallback(28)="BsTP.SolOut"
 SolCallback(29)="bsRI.SolOut"
 SolCallback(30)="SolBallPopper"

'**************
' Flipper Subs
'**************

SolCallback(sLLFlipper) = "SolLFlipper"
SolCallback(sLRFlipper) = "SolRFlipper"



Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFX("fx_flipperup",DOFContactors), LeftFlipper, VolFlip
        LeftFlipper.RotateToEnd
    Else
        PlaySoundAtVol SoundFX("fx_flipperdown",DOFContactors), LeftFlipper, VolFlip
        LeftFlipper.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFX("fx_flipperup",DOFContactors), RightFlipper, VolFlip
        RightFlipper.RotateToEnd
    Else
        PlaySoundAtVol SoundFX("fx_flipperdown",DOFContactors), RightFlipper, VolFlip
        RightFlipper.RotateToStart
    End If
End Sub

Sub SolBallPopper(Enabled)
If Enabled Then
 If bsBP.Balls Then
 sw45.DestroyBall
 bsBP.ExitSol_On
 End If
End If
End Sub

Sub SolKickBack(enabled)
    If enabled Then
       KickBack.Fire
       PlaySoundAtVol SoundFX("fx_solenoid",DOFContactors), KickBack, 1
    Else
       KickBack.PullBack
    End If
End Sub

Sub Sol7(enabled)
If enabled Then
F7.State = 1
F7a.State = 1
else
F7.State = 0
F7a.State = 0
End If
End Sub

Sub Sol8(enabled)
If enabled Then
F8.State = 1
else
F8.State = 0
End If
End Sub

Sub Sol14(enabled)
If enabled Then
F14.State = 1
F14a.State = 1
F14b.State = 1
else
F14.State = 0
F14a.State = 0
F14b.State = 0
End If
End Sub

Sub Sol15(enabled)
If enabled Then
F15.State = 1
F15a.State = 1
F15b.State = 1
else
F15.State = 0
F15a.State = 0
F15b.State = 0
End If
End Sub

Sub Flash62(enabled)
If enabled Then
F62.State = 1
else
F62.State = 0
End If
End Sub

Sub Flash61(enabled)
If enabled Then
F61.State = 1
else
F61.State = 0
End If
End Sub

Sub Flash60(enabled)
If enabled Then
F60.State = 1
else
F60.State = 0
End If
End Sub

Sub Flash59(enabled)
If enabled Then
F59.State = 1
else
F59.State = 0
End If
End Sub

Sub Flash58(enabled)
If enabled Then
F58.State = 1
else
F58.State = 0
End If
End Sub

Sub Flash57(enabled)
If enabled Then
F57.State = 1
else
F57.State = 0
End If
End Sub

Sub Flash9(enabled)
If enabled Then
F9.State = 1
F9a.State = 1
F9b.State = 1
else
F9.State = 0
F9a.State = 0
F9b.State = 0
End If
End Sub

Sub Flash12(enabled)
If enabled Then
F12.State = 1
F12a.State = 1
F12b.State = 1
else
F12.State = 0
F12a.State = 0
F12b.State = 0
End If
End Sub

Sub Flash13(enabled)
If enabled Then
F13.State = 1
F13a.State = 1
F13b.State = 1
else
F13.State = 0
F13a.State = 0
F13b.State = 0
End If
End Sub

'*****GI Lights On
dim xx
Sub SolGIUpdate(enabled)
If enabled Then
For each xx in GI:xx.State = 0: Next
LFLogo.Image = "flipper-l2_dark"
RFLogo.Image = "flipper-r2_dark"
else
For each xx in GI:xx.State = 1: Next
LFLogo.Image = "flipper-l2"
RFLogo.Image = "flipper-r2"
End If
End Sub



'VPM Light Subs - VP10 Built in Fading Routines are used.

Lights(24)=Array(L24,L24a)
Set Lights(1) = L1
Set Lights(2) = L2
Set Lights(3) = L3
Set Lights(4) = L4
Set Lights(5) = L5
Set Lights(6) = L6
Set Lights(7) = L7
Set Lights(8) = L8
Set Lights(9) = L9
Set Lights(10) = L10
Set Lights(11) = L11
Set Lights(12) = L12
Set Lights(13) = L13
Set Lights(14) = L14
Set Lights(15) = L15
Set Lights(16) = L16
Set Lights(17) = L17
Set Lights(18) = L18
Set Lights(19) = L19
Set Lights(20) = L20
Set Lights(21) = L21
Set Lights(22) = L22
Set Lights(23) = L23
Set Lights(25) = L25
Set Lights(26) = L26
Set Lights(27) = L27
Set Lights(28) = L28
Set Lights(29) = L29
Set Lights(30) = L30
Set Lights(31) = L31
Set Lights(32) = L32
Set Lights(33) = L33
Set Lights(34) = L34
Set Lights(35) = L35
Set Lights(36) = L36
Set Lights(37) = L37
Set Lights(38) = L38
Set Lights(39) = L39
Set Lights(40) = L40
Set Lights(41) = L41
Set Lights(42) = L42
Set Lights(43) = L43
Set Lights(45) = L45
Set Lights(46) = L46
Set Lights(47) = L47
Set Lights(48) = L48
Set Lights(49) = L49
Set Lights(50) = L50
Set Lights(51) = L51

Set Lights(52) = Light44
Set Lights(53) = Light46
Set Lights(54) = Light48
Set Lights(55) = Light50

Set Lights(57) = L57
Set Lights(58) = L58
Set Lights(59) = L59
Set Lights(60) = L60
Set Lights(61) = L61
Set Lights(62) = L62
Set Lights(63) = L63
Set Lights(64) = L64

Sub LSample_Timer()
F43.visible = L43.state
F42.visible = L42.state
F41.visible = L41.state
End Sub

'*****

Sub UpdateFlipperLogo_Timer
    LFLogo.RotY = LeftFlipper.CurrentAngle
    RFlogo.RotY = RightFlipper.CurrentAngle
    'RFlogo1.RotY = RightFlipper2.CurrentAngle
    LLogo.ObjRotZ = LeftFlipper.CurrentAngle -90
    Rlogo.ObjRotZ = RightFlipper.CurrentAngle -90
    'Rlogo1.ObjRotZ = RightFlipper2.CurrentAngle -90
End Sub

Sub Drain_Hit()
	PlaySoundAtVol "drain", Drain, 1
    bsTrough.AddBall Me
End Sub

Sub Bumper1_Hit
	PlaySoundAtVol SoundFX("fx_bumper4",DOFContactors), Bumper1, VolBump
	vpmTimer.PulseSw(46)
	'B1L1.State = 1:B1L2. State = 1
	Me.TimerEnabled = 1
End Sub

Sub Bumper1_Timer
	'B1L1.State = 0:B1L2. State = 0
	Me.Timerenabled = 0
End Sub

Sub Bumper2_Hit
    vpmTimer.PulseSw(47)
	PlaySoundAtVol SoundFX("fx_bumper4",DOFContactors), Bumper2, VolBump
	'B2L1.State = 1:B2L2. State = 1
	Me.TimerEnabled = 1
End Sub

Sub Bumper2_Timer
	'B2L1.State = 0:B2L2. State = 0
	Me.Timerenabled = 0
End Sub

Sub Bumper3_Hit
    vpmTimer.PulseSw(46)
	PlaySoundAtVol SoundFX("fx_bumper4",DOFContactors), Bumper3, VolBump
	'B3L1.State = 1:B3L2. State = 1
	Me.TimerEnabled = 1
End Sub

Sub Bumper3_Timer
	'B3L1.State = 0:B3L2. State = 0
	Me.Timerenabled = 0
End Sub


Const WallPrefix 		= "T" 'Change this based on your naming convention
Const PrimitivePrefix 	= "PrimT"'Change this based on your naming convention
Const PrimitiveBumperPrefix = "BumperRing" 'Change this based on your naming convention
Dim primCnt(100), primDir(100), primBmprDir(100)

'****************************************************************************
'***** Primitive Standup Target Animation
'****************************************************************************
'USAGE: 	Sub sw1_Hit: 	PrimStandupTgtHit  1, Sw1, PrimSw1: End Sub
'USAGE: 	Sub Sw1_Timer: 	PrimStandupTgtMove 1, Sw1, PrimSw1: End Sub

Const StandupTgtMovementDir = "TransX"
Const StandupTgtMovementMax = 6

Sub PrimStandupTgtHit (swnum, wallName, primName)
	PlaySoundAtVol SoundFX("target",DOFContactors), ActiveBall, 1
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


Sub sw14_Hit()
Controller.Switch(14) = 1
End Sub

Sub sw14_UnHit()
Controller.Switch(14) = 0
End Sub

Sub sw17_Hit : vpmTimer.PulseSw(17) : End Sub
Sub sw18_Hit : vpmTimer.PulseSw(18) : End Sub
Sub sw19_Hit : vpmTimer.PulseSw(19) : End Sub
Sub sw20_Hit : vpmTimer.PulseSw(20) : End Sub

'Sub sw23_Hit : vpmTimer.PulseSw(23) : End Sub
Sub t23_Hit: 	PrimStandupTgtHit 23, T23, PrimT23: End Sub
Sub t23_Timer: PrimStandupTgtMove 23, T23, PrimT23: End Sub

Sub sw25_Hit : vpmTimer.PulseSw(25) : End Sub
Sub sw26_Hit : vpmTimer.PulseSw(26) : End Sub
Sub sw27_Hit : vpmTimer.PulseSw(27) : End Sub
Sub sw32_Hit : vpmTimer.PulseSw(32) : End Sub

'Sub sw33_Hit : vpmTimer.PulseSw(33) : End Sub
'Sub sw34_Hit : vpmTimer.PulseSw(34) : End Sub
'Sub sw35_Hit : vpmTimer.PulseSw(35) : End Sub
'Sub sw36_Hit : vpmTimer.PulseSw(36) : End Sub

Sub t33_Hit: 	PrimStandupTgtHit 33, T33, PrimT33: End Sub
Sub t33_Timer: PrimStandupTgtMove 33, T33, PrimT33: End Sub
Sub t34_Hit: 	PrimStandupTgtHit 34, T34, PrimT34: End Sub
Sub t34_Timer: PrimStandupTgtMove 34, T34, PrimT34: End Sub
Sub t35_Hit: 	PrimStandupTgtHit 35, T35, PrimT35: End Sub
Sub t35_Timer: PrimStandupTgtMove 35, T35, PrimT35: End Sub
Sub t36_Hit: 	PrimStandupTgtHit 36, T36, PrimT36: End Sub
Sub t36_Timer: PrimStandupTgtMove 36, T36, PrimT36: End Sub

'Sub sw41_Hit : vpmTimer.PulseSw(41) : End Sub
'Sub sw42_Hit : vpmTimer.PulseSw(42) : End Sub
'Sub sw43_Hit : vpmTimer.PulseSw(43) : End Sub

Sub t41_Hit: 	PrimStandupTgtHit 41, T41, PrimT41: End Sub
Sub t41_Timer: PrimStandupTgtMove 41, T41, PrimT41: End Sub
Sub t42_Hit: 	PrimStandupTgtHit 42, T42, PrimT42: End Sub
Sub t42_Timer: PrimStandupTgtMove 42, T42, PrimT42: End Sub
Sub t43_Hit: 	PrimStandupTgtHit 43, T43, PrimT43: End Sub
Sub t43_Timer: PrimStandupTgtMove 43, T43, PrimT43: End Sub

Sub sw44_Spin() : vpmtimer.pulsesw(44) : PlaySoundAtVol "fx_spinner",sw44,VolSpin: End Sub

 Sub sw28_Hit : bsLE.AddBall Me : End Sub
 Sub sw29_Hit : bsTP.AddBall Me : End Sub
 Sub sw30_Hit : bsRI.AddBall Me : End Sub

Sub sw31_Hit()
vpmTimer.PulseSw(31)
PlaySoundAtVol "Gate4", ActiveBall, VolGates
gpos2 = 1
GatePrim2.enabled = 1
End Sub

 Sub sw45_Hit : bsBP.AddBall Me : End Sub

 Sub LR1_Hit()
	 LR1.DestroyBall
     LR2.CreateBall
     LR2.Kick 0, 10
 End Sub

Sub CapBall1_Unhit()
	me.enabled = 0
End Sub

Sub SRD_Hit()
  PlaySoundAtVol "Ball_Bounce", ActiveBall, 1
End Sub

Sub LRD_Hit()
PlaySoundAtVol "Ball_Bounce", ActiveBall, 1
End Sub

Sub PGT_Hit()
PlaySoundAtVol "Gate4", ActiveBall, VolGates
If ActiveBall.VelY < - 1 Then
gpos = 1
GatePrim.enabled = 1
End If
If ActiveBall.VelY > 1 Then
gpos = 3
GatePrim.enabled = 1
End If
End Sub

Dim gpos
Sub GatePrim_Timer()
Select Case gpos
Case 1:
If Primitive226.RotX => 90 Then
gpos = 2
End If
Primitive226.RotX = Primitive226.RotX + 1
Case 2:
If Primitive226.RotX <= 0 Then
Me.enabled = 0
End If
Primitive226.RotX = Primitive226.RotX - 1
Case 3:
If Primitive226.RotX <= -90 Then
gpos = 4
End If
Primitive226.RotX = Primitive226.RotX - 1
Case 4:
If Primitive226.RotX => 0 Then
Me.enabled = 0
End If
Primitive226.RotX = Primitive226.RotX + 1
End Select
End Sub

Dim gpos2
Sub GatePrim2_Timer()
Select Case gpos2
Case 1:
If Primitive225.RotX => 90 Then
gpos2 = 2
End If
Primitive225.RotX = Primitive225.RotX + 1
Case 2:
If Primitive225.RotX <= 0 Then
Me.enabled = 0
End If
Primitive225.RotX = Primitive225.RotX - 1
End Select
End Sub

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
    vpmTimer.PulseSw(22)
    PlaySoundAtVol SoundFX("left_slingshot",DOFContactors), sling1, 1
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
    RightSlingShot.TimerInterval = 10
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
    vpmTimer.PulseSw(21)
    PlaySoundAtVol SoundFX("right_slingshot",DOFContactors), sling2, 1
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
	LeftSlingShot.TimerInterval = 10
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub Pins_Hit (idx)
	PlaySound "pinhit_low", 0, Vol(ActiveBall)*VolPi, Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
	PlaySound "target", 0, Vol(ActiveBall)*VolTarg, Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
	PlaySound "metalhit_thin", 0, Vol(ActiveBall)*VolMetal, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
	PlaySoundAtVol "metalhit_medium", 0, Vol(ActiveBall)*VolMetal, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
	PlaySound "metalhit2", 0, Vol(ActiveBall)*VolMetal, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
	PlaySound "gate4", 0, Vol(ActiveBall)*VolGates, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Rollover_Hit (idx)
	PlaySound "rollover", 0, Vol(ActiveBall)*VolRol, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Kickers_Hit (idx)
	PlaySound "kicker_enter_center", 0, Vol(ActiveBall)*VolKick, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub Posts_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 16 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolPo, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 16 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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
		Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub

 ' LED display
 ' Based on the Eala's rutine

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

 Sub Leds_Timer
    Dim ChgLED, ii, jj, num, chg, stat, obj, b, x
    ChgLED=Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
       For ii=0 To UBound(chgLED)
          num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)
          For Each obj In Digits(num)
             If chg And 1 Then obj.State=stat And 1
             chg=chg\2 : stat=stat\2
          Next
       Next
    End If
 End Sub

Sub Table1_exit()
	Controller.Pause = False
	Controller.Stop
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

Sub PlaySoundAtVol(sound, tableobj, Volum)
  PlaySound sound, 1, Volum, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
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
  Vol = Csng(BallVel(ball) ^2 / VolDiv)
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

Const tnob = 15 ' total number of balls
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
    If ball1.id = 666 OR ball2.id = 666 Then
      PlaySound("rubber_hit_3"), 0, Csng(velocity) ^2 / (VolDiv/VolCol), Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(Ball1)
    else
      PlaySound("fx_collide"), 0, Csng(velocity) ^2 / (VolDiv/VolCol), Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
    End If
End Sub


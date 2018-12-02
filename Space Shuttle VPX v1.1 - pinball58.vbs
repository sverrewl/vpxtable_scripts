'*********************************************************************************************************************
'SPACE SHUTTLE
'Williams 1984
'version 1.1
'VPX SS recreation by pinball58
'3D ramp,shuttle toy texture modifications and ambient light tuning by Tom Tower
'Thanks to Sepeteus for the shuttle toy borrowed from his FP table
'*********************************************************************************************************************
Option Explicit
Randomize


' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
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
Const VolRH     = 1    ' Rubber hits volume.
Const VolTarg   = 1    ' Targets volume.
Const VolSpin   = 1.5  ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.


On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Dim bsTrough,bsLeftHole,bsRightHole,dtBank,dtDrop,ShuttleLightMod,FallingAstronautMod,Astronaut,TDT,HiddenValue,CabinetMode

LoadVPM "01560000", "S11.vbs", 3.36

'*********************************************** TABLE OPTIONS *******************************************************
'---------------------------------------------------------------------------------------------------------------------
'*********************************************************************************************************************


ShuttleLightMod = 1         '1 = enables the light inside the Shuttle                               0 = disabled

FallingAstronautMod = 1     '1 = enables astronaut's spin when the ball falls in the outlanes       0 = disabled

CabinetMode = 1             '1 = hides pinball's side and rails in cabinet mode                     0 = visibles


'*********************************************************************************************************************
'---------------------------------------------------------------------------------------------------------------------
'*********************************************************************************************************************

'*********** Desktop/Cabinet settings ************************

If SpaceShuttle.ShowDT = true Then
	HiddenValue = 0
Else
	HiddenValue = 1
	If CabinetMode = 1 Then
	Ramp16.visible=0:Ramp15.visible=0:Primitive52.visible=0:Primitive12.visible=0
	Primitive87.visible=0:Primitive89.visible=0:Primitive88.visible=0:Primitive90.visible=0
End If
End If

'*************************************************************

'*********** Standard definitions ****************

 Const UseSolenoids = 2
 Const UseLamps = 0
 Const UseSync = 0
 Const HandleMech = 0

'Standard Sounds
 Const SSolenoidOn = "solenoid"
 Const SSolenoidOff = ""
 Const SFlipperOn = ""
 Const SFlipperOff = ""
 Const SCoin = "CoinIn"

Const cGameName = "sshtl_l7"

'*************************************************

'************ Space Shuttle Init *****************

Sub SpaceShuttle_Init
	vpmInit me
	With Controller
		.GameName = cGameName
        .SplashInfoLine = "Space Shuttle - Williams 1984"
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = HiddenValue
        On Error Resume Next
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
        On Error Goto 0
		SolGI(0)
		Wall119.IsDropped=true : Wall118.IsDropped=false : CPwall.IsDropped=true
		Astronaut=0
	    End With

'Nudging
	vpmNudge.TiltSwitch = swTilt
	vpmNudge.Sensitivity = 3
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingShot, RightSlingShot)

'Trough
Set bsTrough=New cvpmBallStack
	With bsTrough
		.InitSw 9,12,11,10,0,0,0,0
		.InitKick BallRelease,85,7
		.InitExitSnd SoundFX("ballrelease",DOFContactors),SoundFX("",DOFContactors)
		.Balls=3
 	End With

'Left Kicker
Set bsLeftHole=New cvpmBallStack
 	With bsLeftHole
 		.InitSaucer sw16,16,190,8
 		.InitExitSnd SoundFX("",DOFContactors),SoundFX("",DOFContactors)
 	End With

'Right Kicker
 Set bsRightHole=New cvpmBallStack
 	With bsRightHole
 		.InitSaucer sw24,24,170,8
 		.InitExitSnd SoundFX("",DOFContactors),SoundFX("",DOFContactors)
 	End With

'3 Bank Target
Set dtbank = new cvpmdroptarget
     With dtbank
         .initdrop array(sw33, sw34, sw35), array(33, 34, 35)
         .initsnd SoundFX("droptarget",DOFContactors),SoundFX("resetdrop",DOFContactors)
     End With

End Sub

'*************************************************

'********* Solenoids ************

SolCallback(1) = "bsTrough.SolIn"     'Outhole
SolCallback(2) = "bsTrough.SolOut"    'BallRelease
SolCallback(3) = "bsLeftHole.SolOut"  'Left Kicker
SolCallback(4) = "bsRightHole.SolOut" 'Right Kicker
SolCallback(5) = "TdtUP"              'T Drop Target

Sub TdtUP(enabled)
	If enabled Then
	Controller.Switch(20)=0
	PlaySound SoundFX("resetdrop",DOFContactors)
	sw20t.IsDropped=false
	TDT=0
	updateTDT.enabled=1
	End If
End Sub

SolCallback(6) = "dtbank.SolDropUp"   '3 Bank Target
SolCallback(7) = "UPpost"             'Up Post
SolCallback(8) = "DOWNpost"           'Down Post
SolCallback(9) = "Space"              'Shuttle Flash Lamp
SolCallback(10) = "Shuttle"           'Space Flash Lamp

Sub Space(enabled)
	If enabled Then
	FlashSP.visible=1:FlashSP1.visible=1:FlashSP2.visible=1:Light76.state=1:Light77.state=1:Light78.state=1:Light93.state=1:l1.state=1
	If ShuttleLightMod=1 Then
	ShuttleLightON
	End If
	Else
	FlashSP.visible=0:FlashSP1.visible=0:FlashSP2.visible=0:Light76.state=0:Light77.state=0:Light78.state=0:Light93.state=0:l1.state=0
	If ShuttleLightMod=1 Then
	ShuttleLightOFF
	End If
	End If
End Sub

Sub Shuttle(enabled)
	If enabled Then
	FlashSH.visible=1:FlashSH1.visible=1:FlashSH2.visible=1:Light79.state=1:Light80.state=1:Light81.state=1:Light94.state=1:l2.state=1
	Else
	FlashSH.visible=0:FlashSH1.visible=0:FlashSH2.visible=0:Light79.state=0:Light80.state=0:Light81.state=0:Light94.state=0:l2.state=0
	End If
End Sub

SolCallback(11) = "SolGI"

Sub SolGI(enabled)
    If enabled Then
	For each lights in GI:lights.state=0:Next
	For each lights in GI2:lights.state=0:Next
	SpaceShuttle.ColorGradeImage="ColorGradeDark"
	usal.visible=0:usal1.visible=0:usal2.visible=0:usal3.visible=0
	Primitive105.image="p1_text":Primitive97.image="p2_text"
    Else
    For each lights in GI:lights.state=1:Next
	For each lights in GI2:lights.state=1:Next
	SpaceShuttle.ColorGradeImage="ColorGradeBright"
	usal.visible=1:usal1.visible=1:usal2.visible=1:usal3.visible=1
	Primitive105.image="p1_textON":Primitive97.image="p2_textON"
    End If
End Sub

'SolCallback(12) = ""                  Not Used
SolCallback(13) = "RLdiv"             'Diverter Gate
'SolCallback(14) = ""                  Insert Flash Lamp
SolCallback(15) = "RingingBell"       'Bell

Sub RingingBell(enabled)
	If enabled Then
	PlaySound "ringing_bell"
	Else
	StopSound "ringing_bell"
	End If
End Sub

'SolCallback(16) = ""                  Coin LockOut
'SolCallback(17) = ""                  Left SlingShot
'SolCallback(18) = ""                  Right SlingShot
'SolCallback(19) = ""                  Left Jet Bumper
'SolCallback(20) = ""                  Bottom Jet Bumper
'SolCallback(21) = ""                  Right Jet Bumper
'SolCallback(22) = ""                  Not Used

Sub RLdiv(enabled)
	If enabled Then
	Primitive11.RotY=7 : Wall119.IsDropped=false : Wall118.IsDropped=true
	PlaySoundAtVol SoundFX("barrier_click",DOFContactors), Primitive11,1
	Else
	Primitive11.RotY=52 : Wall119.IsDropped=true : Wall118.IsDropped=false
	PlaySoundAtVol SoundFX("barrier_click",DOFContactors), Primitive11,1
	End If
End Sub

Sub UPpost(enabled)
  If CenterPost.TransZ=0 Then PlaySoundAtVol SoundFX("centerpost",DOFContactors),CenterPost,1:End If
	If enabled Then
	CPwall.IsDropped=false : CenterPost.TransZ=25
	End If
End Sub

Sub DOWNpost(enabled)
  If CenterPost.TransZ=25 Then PlaySoundAtVol SoundFX("centerpost",DOFContactors),CenterPost,1:End If
	If enabled Then
	CPwall.IsDropped=true : CenterPost.TransZ=0
	End If
End Sub

'Drain
 Sub Drain_Hit():PlaysoundAtVol "drain",drain,1:bsTrough.AddBall Me:End Sub

'********************************

'********* Flippers *************

SolCallback(sLLFlipper) = "SolLFlipper"
SolCallback(sLRFlipper) = "SolRFlipper"

Sub SolLFlipper(Enabled)
     If Enabled Then
		 PlaySoundAtVol SoundFX("FlipperUp",DOFContactors), LeftFlipper, VolFlip
		 LeftFlipper.RotateToEnd
     Else
		 PlaySoundAtVol SoundFX("FlipperDown",DOFContactors), LeftFlipper, VolFlip
		 LeftFlipper.RotateToStart
     End If
 End Sub

Sub SolRFlipper(Enabled)
	If enabled Then
		 PlaySoundAtVol SoundFX("FlipperUp",DOFContactors), RightFlipper, VolFlip
		 RightFlipper.RotateToEnd
     Else
		 PlaySoundAtVol SoundFX("FlipperDown",DOFContactors), RightFlipper, VolFlip
		 RightFlipper.RotateToStart
     End If
 End Sub

'*******************************

'************** Keys *************************

Sub SpaceShuttle_KeyDown(ByVal keycode)
If keycode = PlungerKey Then
		Plunger.PullBack
		PlaySoundAtVol "plungerpull", Plunger, 1
	End If
If keycode = LeftTiltKey Then PlaySound SoundFX("fx_nudge",0)
If keycode = RightTiltKey Then PlaySound SoundFX("fx_nudge",0)
If keycode = CenterTiltKey Then PlaySound SoundFX("fx_nudge",0)
If vpmKeyDown(keycode) Then Exit Sub

End Sub

Sub SpaceShuttle_KeyUp(ByVal keycode)
If keycode = PlungerKey Then
		Plunger.Fire
		PlaySoundAtVol "plunger", Plunger, 1
	End If
If vpmKeyUp(keycode) Then Exit Sub
End Sub

'*********************************************

'********** Switches ***********

Sub sw29_Hit():VPMTimer.PulseSw 29:End Sub 'U Lane
Sub sw30_Hit():VPMTimer.PulseSw 30:End Sub 'S Lane
Sub sw31_Hit():VPMTimer.PulseSw 31:End Sub 'A Lane
Sub sw16_Hit():PlaySoundAtVol "kicker_enter_center", ActiveBall, 1: bsLeftHole.Addball me:Controller.Switch(16) = 1:End Sub 'Left Kicker
Sub sw24_Hit():PlaySoundAtVol "kicker_enter_center",ActiveBall, 1: bsRightHole.Addball me:Controller.Switch(24) = 1:End Sub 'Right Kicker
Sub sw45_Hit():VPMTimer.PulseSw 45:End Sub 'Center Ramp Lower Switch
Sub sw32_Hit():VPMTimer.PulseSw 32:End Sub 'Center Ramp Upper Switch
Sub LeftOutLane_Hit():VPMTimer.PulseSw 28:If FallingAstronautMod=1 And Astronaut=0 Then:RotateAstronaut:End If:End Sub 'Left OutLane
Sub LeftInLane_Hit():VPMTimer.PulseSw 14:End Sub 'Left InLane
Sub RightInLane_Hit():VPMTimer.PulseSw 15:End Sub 'Right InLane
Sub RightOutLane_Hit():VPMTimer.PulseSw 13:If FallingAstronautMod=1 And Astronaut=0 And Wall118.IsDropped=true Then:RotateAstronaut:End If:End Sub 'Right OutLane
Sub sw37_Spin():VPMTimer.PulseSw 37:End Sub 'Spinner
Sub sw38_Hit():VPMTimer.PulseSw 38:PlaySoundAtVol SoundFX("target",DOFContactors),ActiveBall,VolTarg:End Sub 'Ramp Target
Sub sw36_Hit():Controller.Switch(36)=1:End Sub 'Plunger Switch
Sub sw36_unHit():Controller.Switch(36)=0:End Sub 'Plunger Switch
Sub sw17_Hit():VPMTimer.PulseSw 17:PlaySoundAtVol SoundFX("target",DOFContactors),ActiveBall, VolTarg:End Sub 'S Target
Sub sw18_Hit():VPMTimer.PulseSw 18:PlaySoundAtVol SoundFX("target",DOFContactors),ActiveBall, VolTarg:End Sub 'H Target
Sub sw19_Hit():VPMTimer.PulseSw 19:PlaySoundAtVol SoundFX("target",DOFContactors),ActiveBall, VolTarg:End Sub 'U Target
Sub sw20_Hit():Controller.Switch(20)=1:PlaySoundAtVol "target",ActiveBall, VolTarg:PlaySoundAtVol SoundFX("droptarget",DOFContactors),ActiveBall,VolTarg:sw20t.IsDropped=true:TDT=1:updateTDT.enabled=1:End Sub 'T Drop Target
Sub sw21_Hit():VPMTimer.PulseSw 21:PlaySoundAtVol SoundFX("target",DOFContactors),ActiveBall,VolTarg:End Sub 'T Target
Sub sw22_Hit():VPMTimer.PulseSw 22:PlaySoundAtVol SoundFX("target",DOFContactors),ActiveBall,VolTarg:End Sub 'L Target
Sub sw23_Hit():VPMTimer.PulseSw 23:PlaySoundAtVol SoundFX("target",DOFContactors),ActiveBall,VolTarg:End Sub 'E Target
Sub sw33_dropped():dtbank.Hit 1:End Sub 'Bank Target 1
Sub sw34_dropped():dtbank.Hit 2:End Sub 'Bank Target 2
Sub sw35_dropped():dtbank.Hit 3:End Sub 'Bank Target 3

'*******************************

'********** Slingshots ***************

Dim Lstep,RStep

Sub LeftSlingShot_Slingshot
	VPMTimer.PulseSw 39
    PlaySoundAtVol SoundFX("LSling",DOFContactors),sling2,1
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -28
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -15
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1

End Sub

Sub RightSlingShot_Slingshot
	VPMTimer.PulseSw 40
    PlaySoundAtVol SoundFX("RSling",DOFContactors),sling1,1
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -28
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -15
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

'*******************************************

'********** Bumpers **********************************************

Dim dirRing : dirRing = -1
Dim dirRing2 : dirRing2 = -1
Dim dirRing3 : dirRing3 = -1

Sub Bumper1_Hit
VPMTimer.PulseSw 25
Me.TimerEnabled = 1
	Dim BumpSound
	BumpSound = Int(rnd*4)+1
	Select Case BumpSound
	Case 1: PlaySoundAtVol SoundFX("fx_bumper1",DOFContactors), Bumper1, VolBump
	Case 2: PlaySoundAtVol SoundFX("fx_bumper2",DOFContactors), Bumper1, VolBump
	Case 3: PlaySoundAtVol SoundFX("fx_bumper3",DOFContactors), Bumper1, VolBump
	Case 4: PlaySoundAtVol SoundFX("fx_bumper4",DOFContactors), Bumper1, VolBump
	End Select
End Sub

Sub Bumper1_timer()
	'If Tilt = false Then
	BumperRing1.Z = BumperRing1.Z + (5 * dirRing)
	If BumperRing1.Z <= 0 Then dirRing = 1
	If BumperRing1.Z >= 40 Then
		dirRing = -1
		BumperRing1.Z = 40
		Me.TimerEnabled = 0
	End If
	'End If
End Sub

Sub Bumper2_Hit
VPMTimer.PulseSw 26
Me.TimerEnabled = 1
	Dim BumpSound
	BumpSound = Int(rnd*4)+1
	Select Case BumpSound
	Case 1: PlaySoundAtVol SoundFX("fx_bumper1",DOFContactors), Bumper2, VolBump
	Case 2: PlaySoundAtVol SoundFX("fx_bumper2",DOFContactors), Bumper2, VolBump
	Case 3: PlaySoundAtVol SoundFX("fx_bumper3",DOFContactors), Bumper2, VolBump
	Case 4: PlaySoundAtVol SoundFX("fx_bumper4",DOFContactors), Bumper2, VolBump
	End Select
End Sub

Sub Bumper2_timer()
	BumperRing2.Z = BumperRing2.Z + (5 * dirRing2)
	If BumperRing2.Z <= 0 Then dirRing2 = 1
	If BumperRing2.Z >= 40 Then
		dirRing2 = -1
		BumperRing2.Z = 40
		Me.TimerEnabled = 0
	End If
End Sub

Sub Bumper3_Hit
VPMTimer.PulseSw 27
Me.TimerEnabled = 1
	Dim BumpSound
	BumpSound = Int(rnd*4)+1
	Select Case BumpSound
	Case 1: PlaySoundAtVol SoundFX("fx_bumper1",DOFContactors), Bumper3, VolBump
	Case 2: PlaySoundAtVol SoundFX("fx_bumper2",DOFContactors), Bumper3, VolBump
	Case 3: PlaySoundAtVol SoundFX("fx_bumper3",DOFContactors), Bumper3, VolBump
	Case 4: PlaySoundAtVol SoundFX("fx_bumper4",DOFContactors), Bumper3, VolBump
	End Select
End Sub

Sub Bumper3_timer()
	BumperRing3.Z = BumperRing3.Z + (5 * dirRing3)
	If BumperRing3.Z <= 0 Then dirRing3 = 1
	If BumperRing3.Z >= 40 Then
		dirRing3 = -1
		BumperRing3.Z = 40
		Me.TimerEnabled = 0
	End If
End Sub

'***************************************************************

'*********** Standup Switches ****************

Sub Wall122_Hit
	VPMTimer.PulseSw 42
	PlaySound "target" , 1,0.3 ' TODO
	PlaySound "rubber_hit_3"
	Rubber2.visible=0
	Rubber19.visible=1
	RubberTimer.enabled=1
End Sub

Sub RubberTimer_timer()
	Rubber2.visible=1
	Rubber19.visible=0
	RubberTimer.enabled=0
End Sub

Sub Wall123_Hit
	VPMTimer.PulseSw 43
	PlaySound "target" , 1,0.3
	PlaySound "rubber_hit_3"
	Rubber17.visible=0
	Rubber20.visible=1
	RubberTimer1.enabled=1
End Sub

Sub RubberTimer1_timer()
	Rubber17.visible=1
	Rubber20.visible=0
	RubberTimer1.enabled=0
End Sub

Sub Wall124_Hit
	VPMTimer.PulseSw 44
	PlaySound "target" , 1,0.3
	PlaySound "rubber_hit_3"
	Rubber16.visible=0
	Rubber21.visible=1
	RubberTimer2.enabled=1
End Sub

Sub RubberTimer2_timer()
	Rubber16.visible=1
	Rubber21.visible=0
	RubberTimer2.enabled=0
End Sub

Sub Wall125_Hit
	VPMTimer.PulseSw 46
	PlaySound "target" , 1,0.3
	PlaySound "rubber_hit_3"
	Rubber5.visible=0
	Rubber23.visible=1
	RubberTimer3.enabled=1
End Sub

Sub RubberTimer3_timer()
	Rubber5.visible=1
	Rubber23.visible=0
	RubberTimer3.enabled=0
End Sub

Sub Wall126_Hit
	VPMTimer.PulseSw 47
	PlaySound "target" , 1,0.3
	PlaySound "rubber_hit_3"
	Rubber4.visible=0
	Rubber22.visible=1
	RubberTimer4.enabled=1
End Sub

Sub RubberTimer4_timer()
	Rubber4.visible=1
	Rubber22.visible=0
	RubberTimer4.enabled=0
End Sub

Sub Wall127_Hit
	VPMTimer.PulseSw 48
	PlaySound "target" , 1,0.3
	PlaySound "rubber_hit_3"
	Rubber1.visible=0
	Rubber24.visible=1
	RubberTimer5.enabled=1
End Sub

Sub RubberTimer5_timer()
	Rubber1.visible=1
	Rubber24.visible=0
	RubberTimer5.enabled=0
End Sub

'************************************

'********************** InGame Updates ***************************

Sub UpdateWG_timer()
	GWL.RotZ=Gate.CurrentAngle
	s_wiregate.RotX=sw32.CurrentAngle-180
	Primitive50.RotY=Primitive96.RotY
End Sub

Sub updateTDT_timer()
	If TDT=0 Then
	sw20bw.ISDropped=false
	sw20.IsDropped=false
	End If
	If TDT=1 Then
	sw20bw.ISDropped=true
	sw20.IsDropped=true
	End If
	updateTDT.enabled=0
End Sub

'*****************************************************************

'*********** Shuttle Light Mod ************

Sub ShuttleLightON:shuttleprim.image="shuttle_textmapON":shuttlelight.state=1:End Sub
Sub ShuttleLightOFF::shuttleprim.image="shuttle_textmap":shuttlelight.state=0:End Sub

'******************************************

'*********** Falling Astronaut Mod ************************

Dim RotVel

Sub RotateAstronaut()
	RotateAstronautTimer.Enabled = 1
	RotVel=12
	Astronaut=1
End Sub

Sub RotateAstronautTimer_Timer()
	Primitive96.RotY =Primitive96.RotY + RotVel
	If Primitive96.RotY >= 3*360 Then RotVel = RotVel - 0.1
	If RotVel <= 0 Then RotateAstronautTimer.Enabled = 0 : Primitive96.RotY = Primitive96.RotY : Astronaut=0
End Sub

'*******************************************************

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
LampTimer.Interval = 5  'lamp fading speed
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
    'NFadeL 1, l1
    'NFadeL 2, l2
    'NFadeL 3, l3
    'NFadeL 4, l4
    'NFadeL 5, l5
    'NFadeL 6, l6
    NFadeLm 7, l7:CPL
	FadeObj 7, CenterPost, "Centerpost4", "Centerpost3", "Centerpost2", "Centerpost"
    NFadeL 8, l8
    NFadeL 9, l9
    NFadeL 10, l10
    NFadeL 11, l11
    NFadeL 12, l12
    NFadeL 13, l13
    NFadeL 14, l14
    NFadeLm 15, l15
	NFadeL 15, l15a
    NFadeL 16, l16
    NFadeL 17, l17
    NFadeL 18, l18
    NFadeL 19, l19
    NFadeL 20, l20
    NFadeL 21, l21
    NFadeL 22, l22
    NFadeL 23, l23
    NFadeL 24, l24
    NFadeLm 25, l25:BCL1
	NFadeLm 25, l25a
	FadeObj 25, Bumpercap1, "bumpercap4", "bumpercap3", "bumpercap2", "bumpercap"
    NFadeLm 26, l26:BCL2
	NFadeLm 26, l26a
	FadeObj 26, Bumpercap2, "bumpercap4", "bumpercap3", "bumpercap2", "bumpercap"
    NFadeLm 27, l27:BCL3
	NFadeLm 27, l27a
	FadeObj 27, Bumpercap3, "bumpercap4", "bumpercap3", "bumpercap2", "bumpercap"
    NFadeL 28, l28
    NFadeL 29, l29
    NFadeL 30, l30
    NFadeL 31, l31
    NFadeL 32, l32
    NFadeL 33, l33
    NFadeL 34, l34
    NFadeL 35, l35
    NFadeL 36, l36
    NFadeL 37, l37
    'NFadeL 38, l38
    'NFadeL 39, l39
    'NFadeL 40, l40
    NFadeLm 41, l41
	NFadeL 41, l41a
    NFadeL 42, l42
    NFadeL 43, l43
    NFadeL 44, l44
    NFadeL 45, l45
    NFadeL 46, l46
    NFadeL 47, l47
    NFadeL 48, l48
    NFadeL 49, l49
    NFadeL 50, l50
    NFadeL 51, l51
    NFadeL 52, l52
    NFadeL 53, l53
    NFadeL 54, l54
    NFadeL 55, l55
    NFadeL 56, l56
    NFadeL 57, l57
    NFadeL 58, l58
    NFadeL 59, l59
    NFadeL 60, l60
    NFadeL 61, l61
    NFadeL 62, l62
    NFadeL 63, l63
    NFadeL 64, l64

'Solenoid Controlled lamps

End Sub

'Center Post illumination
Sub CPL()
If l7.state=1 Then
CPLight.visible=1
Else
CPLight.visible=0
End If
CPLight.Height=CenterPost.TransZ+1
End Sub

'Bumpers illumination
Sub BCL1()
If l25.state=1 Then
bclight1.visible=1:bclight1a.visible=1
Else
bclight1.visible=0:bclight1a.visible=0
End If
End Sub

Sub BCL2()
If l26.state=1 Then
bclight2.visible=1:bclight2a.visible=1
Else
bclight2.visible=0:bclight2a.visible=0
End If
End Sub

Sub BCL3()
If l27.state=1 Then
bclight3.visible=1:bclight3a.visible=1
Else
bclight3.visible=0:bclight3a.visible=0
End If
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

Sub Spinner_Spin
	PlaySoundAtVol "fx_spinner", Spinner, VolSpin
End Sub

Sub Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then
		PlaySound "rubber", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End Select
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySound "rubber", 0, parm / 10, -0.1, 0.15
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "rubber", 0, parm / 10, 0.1, 0.15
End Sub

'******** Others Table Sounds *********

Sub Trigger1_Hit:me.timerenabled=1:End Sub

Sub Trigger1_timer()
	Dim HopSound
	HopSound = Int(rnd*2)+1
	Select Case HopSound
	Case 1: PlaySoundAtVol "ballhop", PegPlasticT30, 1
	Case 2: PlaySoundAtVol "ballhoptwice", PegPlasticT30, 1
	End Select
	me.timerenabled=0
End Sub

Sub CPWall_Hit:PlaySound "post_hit":End Sub

Sub sw16_UnHit:PlaySound "popper_ball":End Sub

Sub sw24_UnHit:PlaySound "popper_ball":End Sub

Sub Wall48_Hit:PlaySound "rubber":End Sub

Sub Wall54_Hit:PlaySound "rubber":End Sub

Sub Wall50_Hit:PlaySound "flip_hit_3":End Sub

Sub Wall56_Hit:PlaySound "flip_hit_3":End Sub

Sub Wall68_Hit:PlaySound "flip_hit_3":End Sub

Sub Wall74_Hit:PlaySound "flip_hit_3":End Sub

Sub Wall71_Hit:PlaySound "flip_hit_3":End Sub

Sub Wall128_Hit:PlaySound "metalhit_medium":End Sub

Sub Wall27_Hit:PlaySound "metalhit_medium":End Sub

Sub Wall120_Hit:PlaySound "metalhit_thin":End Sub

Sub Wall121_Hit:PlaySound "metalhit_thin":End Sub

Sub Wall88_Hit:PlaySound "metalhit_thin":End Sub

Sub Wall15_Hit:PlaySound "metalhit_thin":End Sub

Sub Wall21_Hit:PlaySound "metalhit_thin":End Sub

Sub Wall228_Hit:PlaySound "metalhit_thin":End Sub

Sub Wall11_Hit:PlaySound "metalhit_thin":End Sub

Sub Wall79_Hit:PlaySound "metalhit_thin":End Sub

Sub Wall13_Hit:PlaySound "metalhit_thin":End Sub

Sub Wall119_Hit:PlaySound "metalhit_thin":End Sub

Sub Wall118_Hit:PlaySound "metalhit_thin":End Sub

Sub Wall57_Hit:PlaySound "metalhit_medium":End Sub

Sub Wall62_Hit:PlaySound "metalhit_medium":End Sub

Sub Wall82_Hit:PlaySound "metalhit_thin":End Sub

Sub Wall83_Hit:PlaySound "metalhit_thin":End Sub

Sub Wall116_Hit:PlaySound "metalhit_thin":End Sub

Sub Wall117_Hit:PlaySound "metalhit_thin":End Sub

Sub Wall86_Hit:PlaySound "metalhit_thin":End Sub

Sub Wall34_Hit:PlaySound "metalhit_medium":End Sub

Sub Gate_Hit:PlaySound "gate":End Sub

Sub Wall129_Hit:PlaySound "rubber",1,0.6:End Sub

Sub Wall130_Hit:PlaySound "rubber",1,0.6:End Sub

Sub Trigger2_Hit:PlaySound "enter_ramp":End Sub

Sub Trigger4_Hit:PlaySound "ramp_down":End Sub

Sub Trigger3_Hit:PlaySound "ramp":End Sub

Sub Trigger6_Hit:PlaySound "hop2":End Sub

Sub Trigger7_Hit:PlaySound "hop2":End Sub

Sub Trigger8_Hit:PlaySound "ramp_down":End Sub

Sub Trigger9_Hit:PlaySound "enter_ramp":End Sub

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

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "SpaceShuttle" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / SpaceShuttle.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "SpaceShuttle" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / SpaceShuttle.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "SpaceShuttle" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / SpaceShuttle.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioFade(ball) ' Can this be together with the above function ?
  Dim tmp
  tmp = ball.y * 2 / SpaceShuttle.height-1
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
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / (VolDiv/VolCol), Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub


' Thalamus : Exit in a clean and proper way
Sub SpaceShuttle_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


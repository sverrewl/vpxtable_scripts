Option Explicit
Randomize

' Thalamus 2018-07-19
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Improved directional positions
' Changed UseSolenoids=1 to 2

' Options
' Volume devided by - lower gets higher sound

Const VolDiv = 2000

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const Ballsize = 50
Const BallMass = 1.65

Const cGameName="austin",UseSolenoids=2,UseLamps=0,UseGI=0,SSolenoidOn="fx_SolOn",SSolenoidOff="", SCoin="fx_coin"

LoadVPM "01200000", "SEGA.VBS", 3.10
Dim DesktopMode: DesktopMode = Table1.ShowDT

'**********************************************************************************************************
'Solenoid Call backs
'**********************************************************************************************************
SolCallBack(1)   = "SolTrough"
SolCallBack(2)   = "SolAutoPlungerIM"
SolCallback(3)	 = "bsScoop.SolOut"
SolCallback(4)	 = "SolLaserBeam"											'Laser Beam Eject
SolCallback(7)	 = "SolAustinDance"											'Dancing Austin
SolCallback(8)	 = "bsInodoro.SolOut" 										'Toilet Post
SolCallback(20)	= "SolTimeMachine"											'Time Machine Motor Relay Board
SolCallback(22)	= "SolDiv"													'Laser Beam Diverter
SolCallback(23) = "vpmSolGate Gate, SoundFX(SSolenoidOn,DOFContactors),"
SolCallBack(24) = "vpmSolSound SoundFX(""fx_knocker"",DOFKnocker),"					'Knocker
SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"


SolCallBack(5)   = "SetLamp 105," 								'Time Travel Ramp
SolCallBack(6)   = "SolFlasher6" 								'Right Green Dome
SolCallBack(12)  = "SolFlasher12" 								'Left Blue Dome
SolCallBack(13)  = "SolFlasher13" 								'Right Red Dome
SolCallBack(25)  = "SetLamp 125," 								'PF light insert
SolCallBack(26)  = "SetLamp 126," 								'PF light insert
SolCallBack(27)  = "SolFlasher27" 								'Spot Light
SolCallBack(28)  = "SetLamp 128," 								'PF light insert
SolCallBack(29)  = "SetLamp 129," 								'PF light insert
SolCallBack(30)  = "SetLamp 130," 								'PF light insert
SolCallBack(31)  = "SetLamp 131," 								'PF light insert
SolCallBack(32)  = "SolFlasher32" 								'Left yellow Dome

'************************************************************************
'						 INIT TABLE
'************************************************************************

Dim bsTrough, bsInodoro, bsScoop, mMagnet, mLaserGun, mEvil, plungerIM

Sub Table1_Init
	vpmInit Me
	On Error Resume Next
		With Controller
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
		.SplashInfoLine = "Austin Powers -- Stern, 2001"&chr(13)&"You Suck"
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

	PinMAMETimer.Interval=PinMAMEInterval:
	PinMAMETimer.Enabled=1

	vpmNudge.TiltSwitch=56
	vpmNudge.Sensitivity=5
	vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingshot)

    ' Trough & Ball Release
    Set bsTrough = New cvpmtrough
	With bsTrough
		.Size = 4
		.InitSwitches Array (14, 13, 12, 11)
        .InitExit BallRelease, 90, 10
		.InitEntrySounds "fx_drain", SoundFX(SSolenoidOn,DOFContactors), SoundFX(SSolenoidOn,DOFContactors)
		.InitExitSounds  SoundFX(SSolenoidOn,DOFContactors), SoundFX("fx_ballrel",DOFContactors)
		.Balls = 4
		.CreateEvents "bsTrough", Drain
	End With

    ' Impulse Plunger
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swplunger, 55, 0.6
        .Random 0.3
        .switch 16
        .InitExitSnd SoundFX("fx_AutoPlunger",DOFContactors), SoundFX(SSolenoidOn,DOFContactors)
        .CreateEvents "plungerIM"
    End With

    Set bsInodoro = new cvpmSaucer
	With bsInodoro
        .InitKicker Sw21, 21, 240, 8, 0
		.InitExitVariance 5, 3
		.InitSounds "popper_ball", SoundFX(SSolenoidOn,DOFContactors), SoundFX("popper",DOFContactors)
        .CreateEvents "bsInodoro", sw21
	End With

	Set bsScoop = New cvpmSaucer
	With bsScoop
	    .InitKicker Sw46b, 46, 190, 100, 1.5
		.InitSounds "popper_ball", SoundFX(SSolenoidOn,DOFContactors), SoundFX("popper",DOFContactors)
		.CreateEvents "bsScoop", sw46b
    End With

	Set mMagnet= New cvpmMagnet
 	With mMagnet
		.InitMagnet Magnet, 100
		.GrabCenter = 1
 		.solenoid= 14
		.CreateEvents "mMagnet"
	End With

     ' LaserGun
	Set mLaserGun = new cvpmMech
	With mLaserGun
         .MType = vpmMechOneSol + vpmMechReverse + vpmMechLinear + vpmMechFast
         .Sol1 = 21
         .Length = 750
         .Steps = 75
         .AddSw 43, 0, 0
         .AddSw 44, 40, 74
         .Callback = GetRef("UpdateLserGun")
         .Start
	End With

    'DREvil "By JPSalas"
    Set mEvil = New cvpmMech
	With mEvil
        .MType=vpmMechOneSol+vpmMechReverse+vpmMechLinear
        .Sol1=19
        .Length=130
        .Steps=50
        .AddSw 23,0,1
        .AddSw 22,48,49
        .Callback=GetRef("UpdateEvil")
        .Start
	End With

    ' Misc. Initialisation
	RampDiverter.IsDropped=0:RampDiverter2.IsDropped=1
	sw24.Isdropped = True
	Dim x:x = controller.getmech(3)
	UpdateLserGun x, 0, x
	Primitive13.visible=DesktopMode

End Sub


'******************
'Keys Up and Down
'*****************
Sub Table1_KeyUp(ByVal Keycode)
	If Keycode=LeftMagnaSave Or KeyCode=RightMagnaSave Then Controller.Switch(53)=0
    If keycode = plungerkey then Plunger.Fire: StopSound "fx_plungerpull":Playsound "fx_plunger"
    If vpmKeyUp(keycode) Then Exit Sub
End Sub

Sub Table1_KeyDown(ByVal Keycode)
	If Keycode=LeftMagnaSave Or KeyCode=RightMagnaSave Then Controller.Switch(53)=1
    If keycode = PlungerKey Then Plunger.Pullback: Playsound "fx_plungerpull"
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub Table1_Paused:Controller.Pause = 1:End Sub
Sub Table1_UnPaused:Controller.Pause = 0:End Sub
Sub Table1_Exit:Controller.Stop:End Sub

'************************************************************************
'Solenoid Subs
'************************************************************************
'Trough
Sub SolTrough(Enabled)
    If Enabled Then
        bsTrough.ExitSol_On
        If bsTrough.Balls Then vpmTimer.PulseSw 15
    End If
End Sub

'Autoplunger
Sub SolAutoPlungerIM(Enabled)
    If Enabled Then
        PlungerIM.AutoFire
    End If
End Sub

'Diverter
Sub SolDiv(Enabled)
	If enabled Then
		RampDiverter.Isdropped = 1
        PlaySoundAt SoundFX("DiverterOn",DOFContactors) , ActiveBall
 		RampDiverter2.Isdropped = 0
        Primitive_RampDiverter.RotY = 25
	Else
		RampDiverter.Isdropped = 0
        PlaySoundAt SoundFX("DiverterOff",DOFContactors) , ActiveBall
 		RampDiverter2.Isdropped = 1
        Primitive_RampDiverter.RotY = 0
	End If
End Sub

'Maquina del Tiempo
Sub SolTimeMachine(Enabled)
    If Enabled Then
       TimeMachineTimer.Enabled=1
      Else
       TimeMachineTimer.Enabled=0
  End If
End Sub

Sub TimeMachineTimer_Timer()
    StopSound"Motor": PlaySoundAt "Motor", TimeMachineP
    TimeMachineP.RotX  = TimeMachineP.RotX + 15
    Me.Enabled = 1
End Sub


'Flashers
Sub SolFlasher6(enabled)
 If enabled Then
    Flasher6.opacity = 200
    Flasher6m.opacity = 100
    FlasherL6.state = 1
    Flasher6b.opacity = 100
  Else
    Flasher6.opacity = 0
    Flasher6m.opacity = 0
    FlasherL6.state = 0
    Flasher6b.opacity = 0
 End If
End Sub


Sub SolFlasher12(enabled)
 If enabled Then
    Flasher12.opacity = 200
    Flasher12m.opacity = 100
    FlasherL12.state = 1
    Flasher12b.opacity = 100
  Else
    Flasher12.opacity = 0
    Flasher12m.opacity = 0
    FlasherL12.state = 0
    Flasher12b.opacity = 0
 End If
End Sub


Sub SolFlasher13(enabled)
 If enabled Then
    Flasher13.opacity = 200
    Flasher13m.opacity = 100
    FlasherL13.state = 1
    Flasher13b.opacity = 100
  Else
    Flasher13.opacity = 0
    Flasher13m.opacity = 0
    FlasherL13.state = 0
    Flasher13b.opacity = 0
 End If
End Sub

Sub SolFlasher27(enabled)
 If enabled Then
    Flasher27.opacity = 200
    Flasher27m.opacity = 100
    FlasherLight27.state = 1
    F127a.image = "bulbOn"
  Else
    Flasher27.opacity = 0
    Flasher27m.opacity = 0
    FlasherLight27.state = 0
    F127a.image = "bulb"
 End If
End Sub

Sub SolFlasher32(enabled)
 If enabled Then
    Flasher32.opacity = 200
    Flasher32m.opacity = 100
    FlasherL32.state = 1
    Flasher32b.opacity = 100
  Else
    Flasher32.opacity = 0
    Flasher32m.opacity = 0
    FlasherL32.state = 0
    Flasher32b.opacity = 0
 End If
End Sub

' Flippers
Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("fx_flipperUp", DOFFlippers), LeftFlipper
        LeftFlipper.RotateToEnd
    Else
        PlaySoundAt SoundFX("fx_flipperDown", DOFFlippers), LeftFlipper
        LeftFlipper.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("fx_flipperUp", DOFFlippers), RightFlipper
        RightFlipper.RotateToEnd
    Else
        PlaySoundAt SoundFX("fx_flipperDown", DOFFlippers), RightFlipper
        RightFlipper.RotateToStart
    End If
End Sub

'*****************
' Switches
'*****************

' Slingshots
Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    PlaySoundAt SoundFX("LeftSlingShot", DOFContactors), PegPlasticT4
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 59
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 1:LeftSLing4.Visible = 0:LeftSLing3.Visible = 1:Lemk.RotX = 14
        Case 2:LeftSLing3.Visible = 0:LeftSLing2.Visible = 1:Lemk.RotX = 2
        Case 3:LeftSLing2.Visible = 0:Lemk.RotX = -10:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
    PlaySoundAt SoundFX("RightSlingShot", DOFContactors), PegPlasticT18
    RightSling4.Visible = 1
    Remk.RotX = 26
    RStep = 0
    vpmTimer.PulseSw 62
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 1:RightSLing4.Visible = 0:RightSLing3.Visible = 1:Remk.RotX = 14
        Case 2:RightSLing3.Visible = 0:RightSLing2.Visible = 1:Remk.RotX = 2
        Case 3:RightSLing2.Visible = 0:Remk.RotX = -10:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub


'Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 49 : playsoundAt SoundFX("LeftBumper_Hit",DOFContactors), Bumper1 : End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 50 : playsoundAt SoundFX("RightBumper_Hit",DOFContactors), Bumper2 : End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 51 : playsoundAt SoundFX("TopBumper_Hit",DOFContactors), Bumper3 : End Sub

'Stand Up Targets
' Left Banks Targets
Sub Sw25_Hit:vpmTimer.PulseSw 25 :MoveTarget25 :PlaySoundAt SoundFX("fx_bumper",DOFContactors), ActiveBall:End Sub
Sub Sw26_Hit:vpmTimer.PulseSw 26 :MoveTarget26 :PlaySoundAt SoundFX("fx_bumper",DOFContactors), ActiveBall:End Sub
Sub Sw27_Hit:vpmTimer.PulseSw 27 :MoveTarget27 :PlaySoundAt SoundFX("fx_bumper",DOFContactors), ActiveBall:End Sub
Sub Sw28_Hit:vpmTimer.PulseSw 28 :MoveTarget27 :PlaySoundAt SoundFX("fx_bumper",DOFContactors), ActiveBall:End Sub
Sub Sw29_Hit:vpmTimer.PulseSw 29 :MoveTarget29 :PlaySoundAt SoundFX("fx_bumper",DOFContactors), ActiveBall:End Sub
Sub Sw30_Hit:vpmTimer.PulseSw 30 :MoveTarget30 :PlaySoundAt SoundFX("fx_bumper",DOFContactors), ActiveBall:End Sub

Sub MoveTarget25
	Sw25a.TransZ = 5
	Sw25b.TransZ = 5
	Sw25.Timerenabled = False
	Sw25.Timerenabled = True
End Sub
Sub	Sw25_Timer
	Sw25.Timerenabled = False
	Sw25a.TransZ = 0
	Sw25b.TransZ = 0
End Sub

Sub MoveTarget26
	Sw26a.TransZ = 5
	Sw26b.TransZ = 5
	Sw26.Timerenabled = False
	Sw26.Timerenabled = True
End Sub
Sub	Sw26_Timer
	Sw26.Timerenabled = False
	Sw26a.TransZ = 0
	Sw26b.TransZ = 0
End Sub

Sub MoveTarget27
	Sw27a.TransZ = 5
	Sw27b.TransZ = 5
	Sw27.Timerenabled = False
	Sw27.Timerenabled = True
End Sub
Sub	Sw27_Timer
	Sw27.Timerenabled = False
	Sw27a.TransZ = 0
	Sw27b.TransZ = 0
End Sub

Sub MoveTarget28
	Sw28a.TransZ = 5
	Sw28b.TransZ = 5
	Sw28.Timerenabled = False
	Sw28.Timerenabled = True
End Sub
Sub	Sw28_Timer
	Sw28.Timerenabled = False
	Sw28a.TransZ = 0
	Sw28b.TransZ = 0
End Sub


Sub MoveTarget29
	Sw29a.TransZ = 5
	Sw29b.TransZ = 5
	Sw29.Timerenabled = False
	Sw29.Timerenabled = True
End Sub
Sub	Sw29_Timer
	Sw29.Timerenabled = False
	Sw29a.TransZ = 0
	Sw29b.TransZ = 0
End Sub

Sub MoveTarget30
	Sw30a.TransZ = 5
	Sw30b.TransZ = 5
	Sw30.Timerenabled = False
	Sw30.Timerenabled = True
End Sub
Sub	Sw30_Timer
	Sw30.Timerenabled = False
	Sw30a.TransZ = 0
	Sw30b.TransZ = 0
End Sub


Sub sw31_Hit:vpmTimer.PulseSw 31:End Sub
Sub sw32_Hit:vpmTimer.PulseSw 32:End Sub

'Wire Triggers
Sub sw16_Hit:Controller.Switch(16)=1 : playsoundAt "rollover", ActiveBall : End Sub
Sub sw16_unHit:Controller.Switch(16)=0:End Sub
Sub sw33_Hit:Controller.Switch(33)=1 : playsoundAt "rollover", ActiveBall : End Sub
Sub sw33_unHit:Controller.Switch(33)=0:End Sub
Sub sw39_Hit:Controller.Switch(39)=1 : playsoundAt "rollover", ActiveBall : End Sub
Sub sw39_unHit:Controller.Switch(39)=0:End Sub
Sub sw47_Hit:Controller.Switch(47)=1 : playsoundAt "rollover", ActiveBall : End Sub
Sub sw47_unHit:Controller.Switch(47)=0:End Sub
Sub sw48_Hit:Controller.Switch(48)=1 : playsoundAt "rollover", ActiveBall : End Sub
Sub sw48_unHit:Controller.Switch(48)=0:End Sub
Sub sw57_Hit:Controller.Switch(57)=1 : playsoundAt "rollover", ActiveBall : End Sub
Sub sw57_unHit:Controller.Switch(57)=0:End Sub
Sub sw58_Hit:Controller.Switch(58)=1 : playsoundAt "rollover", ActiveBall : End Sub
Sub sw58_unHit:Controller.Switch(58)=0:End Sub
Sub sw60_Hit:Controller.Switch(60)=1 : playsoundAt "rollover", ActiveBall : End Sub
Sub sw60_unHit:Controller.Switch(60)=0:End Sub
Sub sw61_Hit:Controller.Switch(61)=1 : playsoundAt "rollover", ActiveBall : End Sub
Sub sw61_unHit:Controller.Switch(61)=0:End Sub


'Ramp Gate Triggers
Sub sw18_hit:vpmTimer.pulseSw 18 : playsoundAt "rollover", ActiveBall : End Sub
Sub sw35_hit:vpmTimer.pulseSw 35 : playsoundAt "fx_chapa", ActiveBall : End Sub

Sub sw37_Hit:Controller.Switch(37)=1 : playsoundAt "fx_gate4", ActiveBall : End Sub
Sub sw37_unHit:Controller.Switch(37)=0:End Sub


Sub sw38_hit:vpmTimer.pulseSw 38 : playsoundAt "fx_gate4", ActiveBall : End Sub

Sub sw41_Hit:Controller.Switch(41)=1 End Sub
Sub sw41_unHit:Controller.Switch(41)=0:End Sub

Sub sw42_hit:vpmTimer.pulseSw 42 :playsoundAt "fx_gate4", ActiveBall: End Sub
Sub sw42_Unhit:vpmTimer.pulseSw 42 : End Sub

'Scoop
Sub Sw46_hit: Sw46.enabled = 0 : End Sub
Sub Sw46b_UnHit: vpmtimer.addtimer 500, "Sw46.enabled = 1 '" :End Sub

'**********************************************************************************************************
'LaserBeam
'**********************************************************************************************************

Const Pi=3.1415926535
Dim BallInGun, GPos, BallInGunRadius
BallInGunRadius = SQR((PCannon.X - Sw45.X)^2 + (PCannon.Y - Sw45.Y)^2)

 Sub SolLaserBeam(Enabled)
	If Enabled Then
		PlaySoundAt SoundFX("CannoShot", DOFContactors), sw45
		If NOT IsEmpty(BallInGun) Then
			Sw45.kick 300 + GPos, 50
			Controller.switch(45) = 0
			BallInGun = Empty
		End If
     End If
End Sub

 Sub UpdateLserGun(acurrpos, aspeed, alastpos)
	GPos = acurrpos
	StopSound"Motor": PlaySoundAt "Motor", PCannon
	PCannon.RotY = GPos
	If Not IsEmpty(BallInGun) Then
		BallInGun.X = PCannon.X - BallInGunRadius * Cos((GPos+30)*Pi/180)
		BallInGun.Y = PCannon.Y - BallInGunRadius * Sin((GPos+30)*Pi/180)
	End If
 End Sub

Sub Sw45_hit(): StopSound "fx_ramp_enter3": PlaySoundAt "popper_ball", sw45: controller.switch(45) = 1:Set BallInGun = ActiveBall:End Sub

'**********************************************************************************************************
'Dancing Austin
'**********************************************************************************************************

Dim cBall
DanceInit

Sub DanceInit
    Set cBall = ckicker.createball
    ckicker.Kick 0, 0
End Sub


Sub SolAustinDance(Enabled)
    If Enabled Then
        AustinDance
    End If
End Sub

Sub AustinDance
    cball.velx = 47
End Sub

'**********************************************************************************************************
'Minimi Spiner "By JPSalas"
'**********************************************************************************************************
Dim SpinPos, SpinDir
SpinPos = 0
SpinDir = 0

Sub Sw34_Spin():playsound "fx_spinner" :End Sub

Sub Sw34Anim_Hit()
    SpinDir = -10 * (ActiveBall.VelY / 2)
    SpinTimer.Enabled = 1
End Sub

Sub Spintimer_Timer()
    SpinPos = SpinPos + int(SpinDir)
    if SpinPos>360 Then
       SpinPos = SpinPos - 360
       vpmTimer.PulseSw 34
    elseif SpinPos < 0 Then
       SpinPos = SpinPos + 360
       vpmTimer.PulseSw 34
    end if
    SpinDir = SpinDir * .95
    if abs(SpinDir) <= 10 then ' Lost momentum, start return to upgright position
        If SpinPos <= 180 Then
            SpinDir =  -((SpinPos * 19) / 180)-1
        else
             SpinDir = ((360-SpinPos) * 19 / 180)+1
        end if
        if SpinPos <=1 or SpinPos >= 359 then SpinPos = 0:Me.Enabled =0
    end if
    MiniMiP.RotX = SpinPos
End Sub

'**********************************************************************************************************
'Dr Evil "By JPSalas"
'**********************************************************************************************************
Dim Sw24Pos
Sub Sw24_Hit():Sw24pos=0:vpmTimer.PulseSw 24:Me.TimerInterval=5:Me.TimerEnabled=1: End Sub

 Sub Sw24_Timer()
 	Select Case Sw24pos
        Case 0: Sw24P.y=657: PlaySoundAt "fx_chapa", sw24P
        Case 1: Sw24P.y=647
        Case 2: Sw24P.y=637
        Case 3: Sw24P.y=627
        Case 4: Sw24P.y=637
        Case 5: Sw24P.y=647
        Case 6: Sw24P.y=657
        Case 7: Sw24P.y=647
        Case 8: Sw24P.y=637
        Case 9: Sw24P.y=627
        Case 10: Sw24P.y=637
        Case 11: Sw24P.y=647
        Case 12: Sw24P.y=657 :Me.TimerEnabled = 0
     End Select
        Sw24pos = Sw24pos + 1
  End Sub

Sub UpdateEvil(aNewPos,aSpeed,aLastPos)
    Sw24P.TransY = aNewPos * 3
    StopSound"Motor": PlaySoundAt "Motor", sw24p
    If aNewPos > 3 then
        sw24.IsDropped = 0
        'DR_Evil.IsDropped = 0
    Else
        sw24.Isdropped = 1
        'DR_Evil.IsDropped = 1
    End If
End Sub

'*************************************************************************************
'		GENERAL ILLUMINATION
'*************************************************************************************
set GICallback = GetRef("UpdateGI")

Sub UpdateGI(no, Enabled)
	If Enabled Then
		dim xx
		For each xx in GILights:xx.State = 1: Next
		For each xx in GILightsBackwalls:xx.visible=1: Next  'BackWall Flashers
        PlaySoundAt "fx_relay_on", Primitive6
        Table1.ColorGradeImage = "ColorGradeLUT256x16_ConSat"
	Else
		For each xx in GILights:xx.State = 0: Next
		For each xx in GILightsBackwalls:xx.visible=0: Next
        PlaySoundAt "fx_relay_off", Primitive6
	    Table1.ColorGradeImage = "ColorGradeLUT256x16_ConSatDark"
	End If
End Sub

'***************************************************
'       JP's VP10 Fading Lamps & Flashers
'       Based on PD's Fading Light System
' SetLamp 0 is Off
' SetLamp 1 is On
' fading for non opacity objects is 4 steps
'***************************************************

Dim bulb
Dim LampState(200), FadingLevel(200)
Dim FlashSpeedUp(200), FlashSpeedDown(200), FlashMin(200), FlashMax(200), FlashLevel(200)

InitLamps()             ' turn off the lights and flashers and reset them to the default parameters
LampTimer.Interval = 10 'lamp fading speed
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

Sub InitLamps()
    Dim x
    For x = 0 to 200
        LampState(x) = 0         ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4       ' used to track the fading state
        FlashSpeedUp(x) = 0.5    ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.35 ' slower speed when turning off the flasher
        FlashMax(x) = 1          ' the maximum value when on, usually 1
        FlashMin(x) = 0          ' the minimum value when off, usually 0
        FlashLevel(x) = 0        ' the intensity of the flashers, usually from 0 to 1
    Next
End Sub

Sub UpdateLamps
	NFadeL 1, Light1
	NFadeL 2, Light2
	NFadeL 3, Light3
	NFadeL 4, Light4
	NFadeL 5, Light5
	NFadeL 6, Light6
	NFadeL 7, Light7
	NFadeL 8, Light8
	NFadeL 9, Light9
	NFadeL 10, Light10
	NFadeL 11, Light11
	NFadeL 12, Light12

	nFadeLm 13, BumperL_Flasher
	nFadeLm 13, BumperL_Flasher_a
	Flash 13, BumperL_Flasher1

	nFadeLm 14, BumperR_Flasher
	nFadeLm 14, BumperR_Flasher_a
	Flash 14, BumperR_Flasher1

	nFadeLm 15, BumperT_Flasher
	nFadeLm 15, BumperT_Flasher_a
	Flash 15, BumperT_Flasher1

	NFadeL 16, Light16
	NFadeL 17, Light17
	NFadeL 18, Light18
	NFadeL 19, Light19
	NFadeL 20, Light20
	NFadeL 21, Light21
	NFadeL 22, Light22
	NFadeL 23, Light23
	NFadeL 24, Light24
 	NFadeL 25, Light25
	NFadeL 26, Light26
	NFadeL 27, Light27
	NFadeL 28, Light28
	NFadeL 29, Light29
	NFadeL 30, Light30
	NFadeL 31, Light31
	NFadeL 32, Light32
	NFadeL 33, Light33
	NFadeL 34, Light34
	NFadeL 35, Light35
	NFadeL 36, Light36
	NFadeL 37, Light37
	NFadeL 38, Light38
	NFadeL 39, Light39
	NFadeL 40, Light40
	NFadeL 41, Light41
	NFadeL 42, Light42
	NFadeL 43, Light43
	NFadeL 44, Light44
	NFadeL 45, Light45
	NFadeL 46, Light46
	NFadeL 47, Light47
	NFadeL 48, Light48
	NFadeL 49, Light49
	NFadeL 50, Light50
	NFadeL 51, Light51
	NFadeL 52, Light52
    NFadeObjm 53, l53, "FireButtonOn", "FireButton"
    NFadeL 53, F53 'Fire Button
	NFadeL 54, Light54 'DR Evil
	NFadeL 57, Light57
	NFadeL 58, Light58
	NFadeL 59, Light59
	NFadeL 60, Light60
	NFadeL 61, Light61
	NFadeL 62, Light62
	NFadeL 63, Light63
	NFadeL 64, Light64
	NFadeL 65, Light65
	NFadeL 66, Light66
	NFadeL 67, Light67
	NFadeL 68, Light68
	NFadeL 73, Light73
	NFadeL 74, Light74
	NFadeL 75, Light75
	NFadeL 76, Light76

'Solenoid Controlled Lamps
	NFadeLm 105, F105
	NFadeL 105, F105a
	NFadeL 125, FlasherLight25
	NFadeL 126, FlasherLight26
	NFadeL 128, FlasherLight28
	NFadeL 129, FlasherLight29
	NFadeL 130, FlasherLight30
	NFadeL 131, FlasherLight31

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
        Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1            'wait
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
        Case 4:object.image = b:FadingLevel(nr) = 0
        Case 5:object.image = a:FadingLevel(nr) = 1
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

' RGB Leds

Sub RGBLED (object,red,green,blue)
	If TypeName(object) = "Light" Then
		object.color = RGB(0,0,0)
		object.colorfull = RGB(2.5*red,2.5*green,2.5*blue)
		object.state=1
	ElseIf TypeName(object) = "Flasher" Then
		object.color = RGB(2.5*red,2.5*green,2.5*blue)
		object.IntensityScale = 1
	End If
End Sub

' Modulated Flasher and Lights objects

Sub SetLampMod(nr, value)
    If value > 0 Then
		LampState(nr) = 1
	Else
		LampState(nr) = 0
	End If
	FadingLevel(nr) = value
End Sub

Sub LampMod(nr, object)
	Object.IntensityScale = FadingLevel(nr)/128
	If TypeName(object) = "Light" Then
		Object.State = LampState(nr)
	End If
	If TypeName(object) = "Flasher" Then
		Object.visible = LampState(nr)
	End If
End Sub

'****************************************
' Real Time updates
'****************************************

Set MotorCallBack = GetRef("GameTimer")

Sub GameTimer
    RollingUpdate
	BallShadowUpdate
	LFLogo.RotY = LeftFlipper.CurrentAngle
	RFlogo.RotY = RightFlipper.CurrentAngle
	AustinDancerP.RotY = -(Spinner1.currentangle)
End Sub


'*********** BALL SHADOW *********************************
Dim BallShadow
BallShadow = Array (BallShadow1, BallShadow2, BallShadow3,BallShadow4,BallShadow5)

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

'******************************
'		 Sound FX
'******************************

Sub aMetals_Hit(idx):PlaySound "fx_MetalHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Bands_Hit(idx):PlaySound "fx_rubber_band", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Posts_Hit(idx):PlaySound "fx_rubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aWoods_Hit(idx):PlaySound "fx_woodhit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub


Sub Trigger2_Hit:PlaySoundAt "fx_chapa", Trigger2 End Sub

'Flipper rubber sounds

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, -0.1, 0.15
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, 0.1, 0.15
End Sub

' Ball Collision Sound
Sub OnBallBallCollision(ball1, ball2, velocity)
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / VolDiv, Pan(ball1), 0, Pitch(ball1), 0, 0
End Sub

Sub LREnter_Hit():If ActiveBall.VelY < 0 Then PlaySoundAt "fx_ramp_enter1",LREnter:End If:End Sub			'ball is going up
Sub LREnter_UnHit():If ActiveBall.VelY > 0 Then StopSound "fx_ramp_enter1":End If:End Sub		'ball is going down
Sub LREnter1_Hit():PlaySoundAt "fx_ramp_enter2",LREnter1:End Sub
Sub LREnter2_Hit():PlaySoundAt "fx_ramp_enter3",LREnter2:End Sub

Sub LRExit_Hit():ActiveBall.VelY=1:StopSound "fx_ramp_enter3":Me.TimerInterval=200:Me.TimerEnabled=1:End Sub
Sub LRExit_timer():Me.TimerEnabled=0:RandomDropSound:End Sub

Sub RREnter_Hit():If ActiveBall.VelY < 0 Then PlaySoundAt "fx_ramp_enter1", RRenter:End If:End Sub			'ball is going up
Sub RREnter_UnHit():If ActiveBall.VelY > 0 Then StopSound "fx_ramp_enter1":End If:End Sub		'ball is going down
Sub RREnter1_Hit():PlaySoundAt "fx_ramp_enter2",RRenter1:End Sub
Sub RREnter2_Hit():PlaySoundAt "fx_ramp_enter3",RRenter2:End Sub

Sub RRExit_Hit():StopSound "fx_ramp_enter3":Me.TimerInterval=200:Me.TimerEnabled=1:End Sub
Sub RRExit_timer():Me.TimerEnabled=0:RandomDropSound:End Sub

Sub TriggerLaunchBall_Hit(): vpmtimer.addtimer 350, "RandomDropSound '" End Sub

Sub RandomDropSound()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "fx_BallDrop1"
		Case 2 : PlaySound "fx_BallDrop2"
		Case 3 : PlaySound "fx_BallDrop3"
	End Select
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
  Vol = Csng(BallVel(ball) ^2 / VolDiv)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
  Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
  BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

'*********** ROLLING SOUND *********************************
Const tnob = 5						' total number of balls : 4 (trough) + 1 (fake ball used to animate Austin toy)
Const fakeballs = 1					' number of balls created on table start (rolling sound will be skipped)
ReDim rolling(tnob)
InitRolling

Sub InitRolling:Dim i:For i=0 to (tnob-1):rolling(i) = False:Next:End Sub

Sub RollingUpdate()
    Dim BOT, b, ballpitch
    BOT = GetBalls
	' stop the sound of deleted balls
	If UBound(BOT)<(tnob -1) Then
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
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


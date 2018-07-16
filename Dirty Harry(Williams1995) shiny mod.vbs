 '*********************************************************************
 '1995 Williams DIRTY HARRY
 '*********************************************************************
 '*********************************************************************
 'Pinball Machine designed by Barry Oursler
 '*********************************************************************
 '*********************************************************************
 'recreated for Visual Pinball by Knorr
 '*********************************************************************
 '*********************************************************************
 'I would like to give my sincere thanks to
 'Mfuegemann, Freneticamnesic, Toxie and the VPdevs, Clark Kent and
 'Gigalula for always being so friendly, helpful and motivating while
 'building this table.
 '*********************************************************************


 'V1.7
		'Added controller.vbs
		'Bug Fixes

 'V1.6
		'First Release For VP10.0.0
		'CrimeWave Gunfix (only 1 Ball can be in the Gun)
		'reduced gun model
		'reduced warehouse model (again)
		'reduced png for smaller file size
		'BallSize is now 52

 'V1.5
		'Fixed Multiball
		'Changed Primitves Rampentry Metals

 'V1.4
		'Added Global Light for Flasher
		'Added Plunger Animation
		'Cleaned up RightRamp Primitive
		'Added Dropwall to Warehouse so only one Ball can be in
		'new Slingshot Plastics Primitives and added Walls for Collision

 'V1.3
		'Updated Ball Rolling/Collision Script
		'small changes with primitives

 'V1.2
		'Fixed CrimeWave Multiball
		'Added Global Lightning (thanks for helping Fren)
		'Added missing Lights for BumperCap and BankRobber
		'added images for ON/OFF effects
		'Changed Sound for Plunger

 'V1.1
		'Reduced meshes in the warehouse model (almost the half)
		'improved lightning for environment (thanks to Fren)
		'reduced lightning for the inserts
		'reduced size of the playfield.jpg and the warehouse
		'minor changes with flashers

 'V1.0
		'First Release For VP10 Beta

Option Explicit
 Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

 LoadVPM "01560000", "WPC.VBS", 3.36
 
 '********************
 'Standard definitions
 '********************
 
  Const cGameName = "dh_lx2"
 Const UseSolenoids = 1
 Const UseLamps = 1
 Const SSolenoidOn = "SolOn"
 Const SSolenoidOff = "SolOff" 
 Const SFlipperOn = "FlipperUp"
 Const SFlipperOff = "FlipperDown"
 Const SCoin = "Coin5"
 
 Set GiCallback2 = GetRef("UpdateGI")
 BSize = 26

 ' Standard Sounds
' Const SSolenoidOn = "Solenoid"
' Const SSolenoidOff = ""
' Const SFlipperOn = "FlipperUp"
' Const SFlipperOff = "FlipperDown"
' Const SCoin = "Coin"

Dim bsTrough, BallInGun, bsSafeHouse, LeftPopper, WareHousePopper, GunPopper, RightMagnet
 
 '************
 ' Table init.
 '************
 
 Sub Table1_Init
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
         .Hidden = 0
         '.SetDisplayPosition 0, 0, GetPlayerHWnd 'uncomment this line if you don't see the DMD
          On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
         .Switch(22) = 1 'close coin door
         .Switch(24) = 1 'and keep it close
			PinMAMETimer.Interval = PinMAMEInterval
			PinMAMETimer.Enabled = true
			vpmNudge.TiltSwitch = 14
			vpmNudge.Sensitivity = 2
     End With


Set bsTrough = New cvpmBallStack
     With bsTrough
         .InitSw 0, 32, 33, 34, 35, 0, 0, 0
         .InitKick BallRelease, 90, 10
		 .InitExitSnd SoundFX("BallRelease",DOFContactors), SoundFX("fx_Solenoid",DOFContactors)
         .Balls = 4
         .IsTrough = 1
     End With


Set bsSafeHouse = New cvpmBallStack
	bsSafeHouse.InitSaucer sw73, 73, 167, 22
	bsSafeHouse.InitExitSnd SoundFX("SafeHouseKick",DOFContactors), SoundFX("fx_Solenoid",DOFContactors)
	bsSafeHouse.KickForceVar = 2
	bsSafeHouse.KickAngleVar = 0.8


Set LeftPopper = New cvpmBallStack
	With LeftPopper
		.InitSw 0, 47, 0, 0, 0, 0, 0, 0
		.InitKick sw47, 180, 15
		.InitExitSnd SoundFX("HeadquarterKick",DOFContactors), "fx_Solenoid"
	End With


Set WarehousePopper = New cvpmBallStack
	With WarehousePopper
		.InitSw 0, 46, 0, 0, 0, 0, 0, 0
		.InitKick sw46, 2, 10
		.KickZ = 1
		.InitExitSnd SoundFX("WareHouseKick",DOFContactors), SoundFX("fx_Solenoid",DOFContactors)
		.KickBalls = 1
	End With


Set GunPopper = New cvpmBallStack
	With GunPopper
		.InitSw 0, 45, 0, 0, 0, 0, 0, 0
		.InitKick sw45, 105, 7
		.InitExitSnd SoundFX("GunPopper",DOFContactors), SoundFX("fx_Solenoid",DOFContactors)
	End With


Set RightMagnet = New cvpmMagnet
     With RightMagnet
         .InitMagnet RMagnet, 70
'         .Solenoid = 35
         .GrabCenter = 1
         .CreateEvents "RightMagnet"
     End With

 
 DiverterOn.isDropped = 1
 DiverterOn2.isDropped = 1
 DiverterOff.isDropped = 0
 Warehousedw.isDropped = 1
 If table1.ShowDT = False then
	Ramp16.WidthTop = 0
	Ramp16.WidthBottom = 0
	Ramp15.WidthTop = 0
	Ramp15.WidthBottom = 0
	Korpus.Size_Y = 0
 End if

End Sub




 '******
 'Trough
 '******

Sub SolRelease(Enabled)
	If Enabled Then
		If bsTrough.Balls = 4 Then vpmTimer.PulseSw 31
		If bsTrough.Balls > 0 Then bsTrough.ExitSol_On
     End If
 End Sub

 Sub Drain_Hit
	PlaySound "Balltruhe",0,0.5,0
	bsTrough.AddBall Me
End Sub


 '*********
 'Safehouse
 '*********

Sub sw73_Hit
    PlaySound "SafeHouseHit"
    bsSafeHouse.AddBall Me
End Sub


 '**********
 'LeftPopper
 '**********


Dim aBall


Sub HQHole_Hit
    PlaySound "HeadquarterHit", 0, 1, pan(ActiveBall)
    Set aBall = ActiveBall:Me.TimerEnabled = 1

    LeftPopper.AddBall 1
End Sub

Sub HQHole_Timer
    Do While aBall.Z > 0
        aBall.Z = aBall.Z -5
        Exit Sub
    Loop
    Me.DestroyBall
    Me.TimerEnabled = 0
End Sub

 '*********
 'Warehouse
 '*********

Sub WarehouseEntry_Hit 'Warehouse
	PlaySound "WareHouseHit"
	WarehousePopper.AddBall Me
	Warehousedw.isDropped = 0
End Sub

Sub Warehousedwtrigger_Hit
	Warehousedw.isDropped = 1
End Sub

 '*********
 'GunPopper
 '*********



Sub TrapDoorKicker_Hit
	PlaySound "HeadquarterHit"
	GunPopper.AddBall Me
End Sub

 '************
 'TrapDoorRamp
 '************

Sub TrapDoorLow(Enabled)
    PlaySound "TrapDoorHigh"
    If Enabled then
        TrapDoorP.RotX = TrapDoorP.RotX + 25
		TrapDoorKicker.Enabled = True
    Else
		 TrapDoorP.RotX = TrapDoorP.RotX - 25
		 PlaySound "TrapDoorLow"
	     TrapDoorKicker.Enabled = False
    End If
End Sub


 '*********
 'MagnumGun
 '*********


Sub SolGunLaunch(Enabled)
     If Enabled AND BallInGun then
		 vpmCreateBall GunKick
         GunKick.kick GPos, 50
         PlaySound "GunShot"
         controller.switch(3) = 0
         BallInGun = 0
         BallP.Visible = False
	Else
		Controller.switch(44) = 0
'		sw44.Enabled = True
		vpmTimer.AddTimer 200, "sw44.Enabled = True'"
     End If
 End Sub



Sub SolGunMotor(Enabled)
    If Enabled Then
       PlaySound SoundFX("GunMotor",DOFGear)
       GDir = -1
       UpdateGun.Enabled=1 
       Controller.switch(77) = 1
     Else
       UpdateGun.Enabled=0
       Controller.switch(77) = 0
	   StopSound "GunMotor"
  End If
 
End Sub

 
Dim GPos, GDir
GPos = -50
GDir = -50
 Sub updategun_Timer()
'    StopSound"": PlaySound ""
    GPos = GPos + GDir
    If GPos <= -98 Then GDir = 1
    If GPos >= -4 Then Controller.switch(76) = 1 Else Controller.switch(76) = 0
    If GPos >= -2 Then GDir = -1
    MagnumGun.RotY = GPos
 End Sub

Sub sw44_hit()
	sw44.Enabled = False
	PlaySound "BallFallInGun"
	StopSound "WireRamp"
	RightWireStart2.Enabled = True
	Controller.switch(44) = 1
	me.DestroyBall
	BallInGun = 1
	BallP.Visible = True
End Sub


 '******
 'Magnet
 '******


Sub SolMagnetOn(Enabled)
	If Enabled then 
		RightMagnet.MagnetOn = True
	Else
		RightMagnet.MagnetOn = False
	End if
End Sub

 '*************
 'RightLoopGate
 '*************

Sub RightLoopGate(Enabled)
	If Enabled then
    Playsound "gate"
    GateR.open = True
	Else
	GateR.open = False
	End if
End sub


 '******
 'Plunger
 '******


Dim AP

Sub AutoPlunge(Enabled)
	if enabled then
		AP = True
		Kicker1.Kick 0,45
		PlaySound SoundFX("Plunger",DOFContactors)
	End if
End Sub

	Sub PlungerPTimer_Timer()
	if AP = True and PlungerP.TransZ < 45 then PlungerP.TransZ = PlungerP.TransZ +10
	if AP = False and PlungerP.TransZ > 0 then PlungerP.TransZ = PlungerP.TransZ -10
	if PlungerP.TransZ >= 45 then AP = False
End Sub


 


 '*********
 'Solenoids
 '*********
 
 SolCallback(1) = "SolRelease"
 SolCallback(2) = "AutoPlunge"
 SolCallback(3) = "SolGunLaunch"
 SolCallback(4) = "WarehousePopper.SolOut"
 SolCallback(5) = "GunPopper.SolOut"
 SolCallback(7) = "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
' SolCallback(8) = "TrapDoorHigh"
 SolCallback(14) = "LeftPopper.SolOut"
 SolCallBack(15) = "vpmSolDiverter diverterR,""DiverterRight"","
 SolCallback(16) = "TrapDoorLow"
 SolCallBack(20) = "SolGunMotor"
 SolCallback(26) = "bsSafeHouse.SolOut"
 SolCallback(27) = "SolDiverterHold"
 SolCallBack(28) = "RightLoopGate"
 SolCallBack(35) = "SolMagnetOn"



 '*********
 'Flasher
 '*********

 SolCallback(17) = "Multi117"
 SolCallback(18) = "Multi118"
 SolCallback(19) = "Multi119"
 SolCallback(21) = "Multi121"
 SolCallback(22) = "Multi122"
 SolCallback(23) = "Multi123"
 SolCallback(24) = "Multi124"


 '**********
 ' Keys
 '**********

 Sub table1_KeyDown(ByVal Keycode)
	If KeyCode=MechanicalTilt Then
		vpmTimer.PulseSw vpmNudge.TiltSwitch
		Exit Sub
	End if

    If keycode = PlungerKey Then Controller.Switch(11) = 1
    If keycode = keyFront Then Controller.Switch(23) = 1
    If vpmKeyDown(keycode) Then Exit Sub
 End Sub
 
 Sub table1_KeyUp(ByVal Keycode)
     If vpmKeyUp(keycode) Then Exit Sub
     If keycode = PlungerKey Then Controller.Switch(11) = 0
     If keycode = keyFront Then Controller.Switch(23) = 0
 End Sub




 '**************
 ' Flipper Subs
 '**************
 
 SolCallback(sLRFlipper) = "SolRFlipper"
 SolCallback(sLLFlipper) = "SolLFlipper"
 
 Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFX("FlipperUpLeft",DOFContactors):LeftFlipper.RotateToEnd
     Else
         PlaySound SoundFX("FlipperDown",DOFContactors):LeftFlipper.RotateToStart
     End If
 End Sub
 
 Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFX("FlipperUpRightBoth",DOFContactors):RightFlipper.RotateToEnd:RightFlipper1.RotateToEnd
     Else
         PlaySound SoundFX("FlipperDown",DOFContactors):RightFlipper.RotateToStart:RightFlipper1.RotateToStart
     End If
 End Sub


Sub FlipperTimer_Timer
LeftFlipperP.RotY = LeftFlipper.CurrentAngle
RightFlipperP1.RotY = RightFlipper1.CurrentAngle
RightFlipperP.RotY = RightFlipper.CurrentAngle
End Sub
 
 
 '*********
 ' Switches
 '*********


Sub sw15_Hit:Controller.Switch(15) = 1:sw15wire.RotX=15:PlaySound "metalhit_thin":End Sub 'shooterlane'
Sub sw15_UnHit:Controller.Switch(15) = 0:sw15wire.RotX=0:End Sub
Sub sw16_Hit:Controller.Switch(16) = 1:sw16wire.RotX=15:PlaySound "metalhit_thin":End Sub 'right outlane'
Sub sw16_UnHit:Controller.Switch(16) = 0:sw16wire.RotX=0:End Sub
Sub sw17_Hit:Controller.Switch(17) = 1:sw17wire.RotX=15:PlaySound "metalhit_thin":End Sub 'right inlane'
Sub sw17_UnHit:Controller.Switch(17) = 0:sw17wire.RotX=0:End Sub
Sub sw26_Hit:Controller.Switch(26) = 1:sw26wire.RotX=15:PlaySound "metalhit_thin":End Sub 'left outlane'
Sub sw26_UnHit:Controller.Switch(26) = 0:sw26wire.RotX=0:End Sub
Sub sw25_Hit:Controller.Switch(25) = 1:sw25wire.RotX=15:PlaySound "metalhit_thin":End Sub 'left inlane'
Sub sw25_UnHit:Controller.Switch(25) = 0:sw25wire.RotX=0:End Sub
Sub sw71_Hit:Controller.Switch(71) = 1:sw71wire.RotX=15:PlaySound "metalhit_thin":End Sub 'left loop'
Sub sw71_UnHit:Controller.Switch(71) = 0:sw71wire.RotX=0:End Sub
Sub sw66_Hit:Controller.Switch(66) = 1:sw66wire.RotX=15:PlaySound "metalhit_thin":End Sub 'left rollover (bumper)'
Sub sw66_UnHit:Controller.Switch(66) = 0:sw66wire.RotX=0:End Sub
Sub sw67_Hit:Controller.Switch(67) = 1:sw67wire.RotX=15:PlaySound "metalhit_thin":End Sub 'middle rollover (bumper)'
Sub sw67_UnHit:Controller.Switch(67) = 0:sw67wire.RotX=0:End Sub
Sub sw68_Hit:Controller.Switch(68) = 1:sw68wire.RotX=15:PlaySound "metalhit_thin":End Sub 'right rollover (bumper)'
Sub sw68_UnHit:Controller.Switch(68) = 0:sw68wire.RotX=0:End Sub
Sub sw42_Hit:Controller.Switch(42) = 1:End Sub				   'right loop'
Sub sw42_UnHit:Controller.Switch(42) = 0: End Sub


Sub BallDrop1_Hit(): StopSound "WireRamp": End Sub
Sub BallDrop2_Hit(): PlaySound "BallDrop": End Sub
Sub BallDrop3_Hit(): PlaySound "BallDrop": End Sub
Sub ZHelper_Hit(): ActiveBall.VelZ = ActiveBall.VelZ -5: End Sub

 '*********
 ' Ramps
 '*********


Sub sw41_Hit:vpmTimer.pulseSw 41: End Sub
Sub sw43_Hit:vpmTimer.pulseSw 43: End Sub
Sub sw51_Hit:vpmTimer.pulseSw 51: End Sub
Sub sw38_Hit:vpmTimer.pulseSw 38: End Sub


	'RampSounds

Dim SoundBall


Sub MiddleWireStart_Hit
 Set SoundBall = Activeball 'Ball-assignment
 playsound "WireRamp",-1,0.6,0,0.35                    '-1 = looping, Panning = -1 bis 1, je nach Position
End Sub


Sub LeftWireStart_Hit
 Set SoundBall = Activeball 'Ball-assignment
 playsound "WireRamp",-1,0.6,0,0.35                    '-1 = looping, Panning = -1 bis 1, je nach Position
End Sub


Sub RightWireStart1_Hit
 RightWireStart2.Enabled = False
 Set SoundBall = Activeball 'Ball-assignment
 playsound "WireRamp",-1,0.6,0,0.35                    '-1 = looping, Panning = -1 bis 1, je nach Position
End Sub

Sub RightWireStart2_Hit
 Set SoundBall = Activeball 'Ball-assignment
 playsound "WireRamp",-1,0.6,0,0.35                    '-1 = looping, Panning = -1 bis 1, je nach Position
End Sub


 '***************
 ' StandupTargets
 '***************

Sub Standup27_Hit:vpmTimer.pulseSw 27:Standup27p.RotY=Standup27p.RotY -3:Playsound SoundFX("target",DOFContactors):Me.TimerEnabled = 1: End Sub
Sub Standup27_Timer:Standup27p.RotY=Standup27p.RotY +3:Me.TimerEnabled = 0: End Sub
Sub Standup28_Hit:vpmTimer.pulseSw 28:Standup28p.RotY=Standup28p.RotY -3:Playsound SoundFX("target",DOFContactors):Me.TimerEnabled = 1: End Sub
Sub Standup28_Timer:Standup28p.RotY=Standup28p.RotY +3:Me.TimerEnabled = 0: End Sub
Sub Standup58_Hit:vpmTimer.pulseSw 58:Standup58p.RotY=Standup58p.RotY -3:Playsound SoundFX("target",DOFContactors):Me.TimerEnabled = 1: End Sub
Sub Standup58_Timer:Standup58p.RotY=Standup58p.RotY +3:Me.TimerEnabled = 0: End Sub
Sub Standup57_Hit:vpmTimer.pulseSw 57:Standup57p.RotY=Standup57p.RotY -3:Playsound SoundFX("target",DOFContactors):Me.TimerEnabled = 1: End Sub
Sub Standup57_Timer:Standup57p.RotY=Standup57p.RotY +3:Me.TimerEnabled = 0: End Sub
Sub Standup56_Hit:vpmTimer.pulseSw 56:Standup56p.RotY=Standup56p.RotY -3:Playsound SoundFX("target",DOFContactors):Me.TimerEnabled = 1: End Sub
Sub Standup56_Timer:Standup56p.RotY=Standup56p.RotY +3:Me.TimerEnabled = 0: End Sub
Sub Standup54_Hit:vpmTimer.pulseSw 54:Standup54p.RotX=Standup54p.RotX +3:Playsound SoundFX("target",DOFContactors):Me.TimerEnabled = 1: End Sub
Sub Standup54_Timer:Standup54p.RotX=Standup54p.RotX -3:Me.TimerEnabled = 0: End Sub
Sub Standup55_Hit:vpmTimer.pulseSw 55:Standup55p.RotY=Standup55p.RotY -3:Playsound SoundFX("target",DOFContactors):Me.TimerEnabled = 1: End Sub
Sub Standup55_Timer:Standup55p.RotY=Standup55p.RotY +3:Me.TimerEnabled = 0: End Sub
Sub Standup18_Hit:vpmTimer.pulseSw 18:Standup18p.RotY=Standup18p.RotY -3:Playsound SoundFX("target",DOFContactors):Me.TimerEnabled = 1: End Sub
Sub Standup18_Timer:Standup18p.RotY=Standup18p.RotY +3:Me.TimerEnabled = 0: End Sub


 '*********
 ' Bumper
 '*********

Sub Bumper63_hit:vpmTimer.pulseSw 63:Playsound SoundFX("BumperLeft",DOFContactors):Me.TimerEnabled = 1: End Sub
Sub Bumper63_Timer:Me.Timerenabled = 0: End Sub

Sub Bumper64_hit:vpmTimer.pulseSw 64:Playsound SoundFX("BumperMiddle",DOFContactors):Me.TimerEnabled = 1: End Sub
Sub Bumper64_Timer:Me.Timerenabled = 0: End Sub

Sub Bumper65_hit:vpmTimer.pulseSw 65:Playsound SoundFX("BumperRight",DOFContactors):Me.TimerEnabled = 1: End Sub
Sub Bumper65_Timer:Me.Timerenabled = 0: End Sub


 

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
    PlaySound SoundFX ("SlingshotLeft",DOFContactors), 0, 1, 0.05, 0.05
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
	vpmTimer.PulseSw 62
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
    PlaySound SoundFX("SlingshotRight",DOFContactors),0,1,-0.05,0.05
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
	vpmTimer.PulseSw 61
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub



 '********
 'Diverter
 '********


Sub DiverterPtimer_timer()
    DiverterP.RotY = diverterR.CurrentAngle
End Sub

Sub SolDiverterHold(Enabled)
	DiverterOFF.IsDropped = Enabled:DiverterOn.IsDropped = Not Enabled
	DiverterOn2.IsDropped = Not Enabled
	If Enabled then Playsound "DiverterLeft": End if
End Sub

 '********
 'RampGate
 '********

Sub GateTimer_Timer
	SpinnerP.RotX = Spinner1.currentangle +95
End Sub



 '*********
 'Update GI
 '*********



Dim xx
Dim gistep
gistep = 1 / 8

Sub UpdateGI(no, step)
    If step = 0 OR step = 7 then exit sub
    Select Case no

		'Bottom String
        Case 4
            For each xx in GIString1:xx.IntensityScale = gistep * step:next
					if step = 1 then Table1.ColorGradeImage = "-70"
					if step = 2 then Table1.ColorGradeImage = "-60"
					if step = 3 then Table1.ColorGradeImage = "-50"
					if step = 4 then Table1.ColorGradeImage = "-40"
					if step = 5 then Table1.ColorGradeImage = "-30"
					if step = 6 then Table1.ColorGradeImage = "-20"
					if step = 7 then Table1.ColorGradeImage = "-10"
					if step = 8 then Table1.ColorGradeImage = "ColorGradeLUT256x16_ConSat"


		'Left String
		Case 1
            For each xx in GIString2:xx.IntensityScale = gistep * step:next


		'Right String
		Case 0
            For each xx in GIString3:xx.IntensityScale = gistep * step:next

    End Select
End Sub


'*****************************************
 '  JP's Fading Lamps 3.4 VP9 Fading only
 '      Based on PD's Fading Lights
 ' SetLamp 0 is Off
 ' SetLamp 1 is On
 ' LampState(x) current state
 '****************************************
 


 Dim LampState(200), FadingLevel(200), FadingState(200)
 Dim FlashState(200), FlashLevel(1000)
 Dim FlashSpeedUp, FlashSpeedDown
 Dim x
 AllLampsOff()
 LampTimer.Interval = 1
 LampTimer.Enabled = 1
FlashInit()
FlasherTimer.Interval = 10 'flash fading speed
FlasherTimer.Enabled = 1
 
Sub LampTimer_Timer()
    Dim chgLamp, num, chg, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp) Then
        For ii = 0 To UBound(chgLamp)
            LampState(chgLamp(ii, 0) ) = chgLamp(ii, 1)
            FadingLevel(chgLamp(ii, 0) ) = chgLamp(ii, 1) + 4
			FlashState(chgLamp(ii, 0) ) = chgLamp(ii, 1)
        Next
    End If

    UpdateLamps
End Sub


Sub FlashInit
    Dim i
    For i = 0 to 200
        FlashState(i) = 0
        FlashLevel(i) = 0
    Next

    FlashSpeedUp = 50   ' fast speed when turning on the flasher
    FlashSpeedDown = 10 ' slow speed when turning off the flasher, gives a smooth fading
    AllFlashOff()
End Sub

Sub AllFlashOff
    Dim i
    For i = 0 to 200
        FlashState(i) = 0
    Next
End Sub
 
  Sub UpdateLamps()
'		NFadeLm 1, l1b
'		NFadeL 1, l1a
'		NFadeLm 2, l2b
'		NFadeL 2, l2a
'		NFadeLm 3, l3b
'		NFadeL 3, l3a
'		NFadeLm 4, l4b
'		NFadeL 4, l4a
'		NFadeLm 5, l5b
'		NFadeL 5, l5a
'		NFadeLm 6, l6b
'		NFadeL 6, l6a
'		NFadeL 7, l7
'		NFadeL 8, l8
'		NFadeL 9, l9
'		NFadeL 10, l10
'		NFadeL 11, l11
		NFadeL 14, l14
		NFadeL 15, l15
		NFadeL 16, l16
		NFadeL 17, l17
		NFadeL 18, l18
		NFadeL 31, l31
		NFadeL 32, l32
		NFadeL 33, l33
		NFadeL 34, l34
		NFadeL 35, l35
		NFadeL 36, l36
		NFadeL 77, l77
		NFadeL 86, l86
		NFadeL 82, l82
		NFadeL 83, l83
		NFadeL 58, l58
		NFadeL 57, l57
		NFadeL 37, l37
		NFadeL 66, l66
		NFadeL 68, l68
		NFadeL 43, l43
		NFadeL 52, l52
		NFadeL 53, l53
		NFadeL 22, l22
		NFadeL 78, l78
		NFadeL 81, l81
		NFadeL 65, l65
		NFadeL 67, l67
		NFadeL 63, l63
		NFadeL 54, l54
		NFadeL 62, l62
		NFadeL 27, l27
		NFadeL 24, l24
		NFadeL 44, l44
		NFadeL 45, l45
		NFadeL 46, l46
		NFadeL 26, l26
		NFadeL 25, l25
		NFadeL 23, l23
		NFadeL 61, l61
		NFadeL 64, l64
		NFadeL 42, l42
		NFadeL 41, l41
		NFadeL 28, l28
		NFadeL 21, l21
		NFadeL 55, l55
		NFadeL 56, l56
		NFadeL 51, l51
		NFadeL 11, l11
		NFadeL 12, l12
		NFadeL 13, l13
		NFadeL 84, l84
		NFadeL 85, l85
		NFadeL 47, l47
		NFadeL 48, l48
		NFadeL 38, l38
If l48.state = 1 then bulbyellow.image = "bulbcover1_yellowOn": l84a.state = 1: else bulbyellow.image = "bulbcover1_yellow": l84a.state = 0
If l47.state = 1 then bulbred.image = "bulbcover1_redOn": else bulbred.image = "bulbcover1_red"
If l85.state = 1 then domesmall.image = "domesmallredOn": else domesmall.image = "domesmallred"
If l84.state = 1 then domesmall1.image = "domesmallredOn": else domesmall1.image = "domesmallred"
If l38.state = 1 then Bankrobber.image = "bankrobbermesh1On": else Bankrobber.image = "bankrobbermesh1"
End Sub


Sub SetFlash(nr, stat)
    FlashState(nr) = ABS(stat)
End Sub


Sub Multi117(Enabled)
	If Enabled Then
		SetFlash 117, 1
		l117a.State = 1
		l117b.state = 1
		l117c.state = 1
		l117d.state = 1
		l117e.state = 1
		Dome1.Image = "dome3_orange_On"
		domesmall2.Image = "domesmallredOn"
	Else
		SetFlash 117, 0
		l117a.State = 0
		l117b.state = 0
		l117c.state = 0
		l117d.state = 0
		l117e.state = 0
		Dome1.Image = "dome3_orange"
		domesmall2.Image = "domesmallred"
	End If
End Sub

Sub Multi118(Enabled)
	If Enabled Then
		SetFlash 118, 1
		l118a.State = 1
		l118b.state = 1
		l118c.state = 1
		l118d.state = 1
		l118e.state = 1
		l118f.state = 1
	Else
		SetFlash 118, 0
		l118a.State = 0
		l118b.state = 0
		l118c.state = 0
		l118d.state = 0
		l118e.state = 0
		l118f.state = 0
	End If
End Sub



Sub Multi119(Enabled)
	If Enabled Then
		SetFlash 119, 1
		l119a.State = 1
		l119b.state = 1
		l119c.state = 1
		l119d.state = 1
		l119e.state = 1
		l119f.state = 1
	Else
		SetFlash 119, 0
		l119a.State = 0
		l119b.state = 0
		l119c.state = 0
		l119d.state = 0
		l119e.state = 0
		l119f.state = 0
	End If
End Sub


Sub Multi122(Enabled)
	If Enabled Then
		SetFlash 122, 1
		l122a.State = 1
		l122b.state = 1
	Else
		SetFlash 122, 0
		l122a.State = 0
		l122b.state = 0
	End If
End Sub


Sub Multi123(Enabled)
	If Enabled Then
		SetFlash 123, 1
		l123a.State = 1
		l123b.state = 1
		l123c.state = 1
		l123d.state = 1
		l123ab.state = 1
		l123ab1.state = 1
		Dome5.Image = "dome3_blue_On"
		Dome3.Image = "dome3_clear_On"
		GIWhite1.State = 1
	Else
		SetFlash 123, 0
		l123a.State = 0
		l123b.state = 0
		l123c.state = 0
		l123d.state = 0
		l123ab.state = 0
		l123ab1.state = 0
		Dome5.Image = "dome3_blue"
		Dome3.Image = "dome3_clear"
		GIWhite1.State = 0
	End If
End Sub

Sub Multi124(Enabled)
	If Enabled Then
		SetFlash 124, 1
		l124a.State = 1
		l124b.state = 1
		l124ab.state = 1
		l124ab1.state = 1
		l124c.state = 1
		l124d.state = 1
		Dome2.Image = "dome3_clear_On"
		Dome4.Image = "dome3_blue_On"
		GIWhite1.State = 1
	Else
		SetFlash 124, 0
		l124a.State = 0
		l124b.state = 0
		l124c.state = 0
		l124d.state = 0
		l124ab.State = 0
		l124ab1.state = 0
		Dome2.Image = "dome3_clear"
		Dome4.Image = "dome3_blue"
		GIWhite1.State = 0
	End If
End Sub

Sub Multi121(Enabled)
	If Enabled Then
		SetFlash 121, 1
		l121a.State = 1
		l121b.state = 1
		l121c.state = 1
	Else
		SetFlash 121, 0
		l121a.State = 0
		l121b.state = 0
		l121c.state = 0
	End If
End Sub

 
 Sub AllLampsOff():For x = 1 to 200:LampState(x) = 4:FadingLevel(x) = 4:Next:UpdateLamps:UpdateLamps:Updatelamps:End Sub
 
' Sub SetLamp(nr, value):LampState(nr) = abs(value) + 4:End Sub

Sub SetLamp(nr, value)
    If value = 0 AND LampState(nr) = 0 Then Exit Sub
    If value = 1 AND LampState(nr) = 1 Then Exit Sub
    LampState(nr) = abs(value) + 4
FadingLevel(nr ) = abs(value) + 4: FadingState(nr ) = abs(value) + 4
End Sub

Sub NFadeL(nr, a)
    Select Case FadingLevel(nr)
        Case 4:a.state = 0:FadingLevel(nr) = 0
        Case 5:a.State = 1:FadingLevel(nr) = 1
    End Select
End Sub

Sub NFadeLm(nr, a)
    Select Case FadingLevel(nr)
        Case 4:a.state = 0
        Case 5:a.State = 1
    End Select
End Sub

Sub FadeL(nr, a, b)
    Select Case FadingLevel(nr)
        Case 2:b.state = 0:FadingLevel(nr) = 0
        Case 3:b.state = 1:FadingLevel(nr) = 2
        Case 4:a.state = 0:FadingLevel(nr) = 3
        Case 5:a.state = 1:FadingLevel(nr) = 1
    End Select
End Sub



Sub FlashAR(nr, ramp, a, b, c, r)                                                          'used for reflections when there is no off ramp
    Select Case FadingState(nr)
        Case 2:ramp.opacity = 0:r.State = ABS(r.state -1):FadingState(nr) = 0                'Off
        Case 3:ramp.image = c:r.State = ABS(r.state -1):FadingState(nr) = 2                'fading...
        Case 4:ramp.image = b:r.State = ABS(r.state -1):FadingState(nr) = 3                'fading...
        Case 5:ramp.image = a:ramp.alpha = 1:r.State = ABS(r.state -1):FadingState(nr) = 1 'ON
    End Select
End Sub


Sub FlashARm(nr, ramp, a, b, c, r)
    Select Case FadingState(nr)
        Case 2:ramp.opacity = 0:r.State = ABS(r.state -1)
        Case 3:ramp.image = c:r.State = ABS(r.state -1)
        Case 4:ramp.image = b:r.State = ABS(r.state -1)
        Case 5:ramp.image = a:ramp.alpha = 1:r.State = ABS(r.state -1)
    End Select
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

Sub Flashm(nr, object) 'multiple flashers, it doesn't change the flashstate
    Select Case FlashState(nr)
        Case 0         'off
            Object.opacity = FlashLevel(nr)
        Case 1         ' on
            Object.opacity = FlashLevel(nr)
    End Select
End Sub






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

Const tnob = 9 ' total number of balls
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = FALSE
    Next
End Sub

Sub CollisionTimer_Timer()
    Dim BOT, b
    BOT = GetBalls

    ' stop the sound of deleted balls
    For b = UBound(BOT) + 1 to tnob
        rolling(b) = FALSE
        StopSound("fx_ballrolling" & b)
    Next

    ' exit the sub if no balls on the table
    If UBound(BOT) = -1 Then Exit Sub

    ' play the rolling sound for each ball
    For b = 0 to UBound(BOT)
        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
            rolling(b) = TRUE
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0
        Else
            If rolling(b) = TRUE Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = FALSE
            End If
        End If
    Next
End Sub

'**********************
' Ball Collision Sound
'**********************


Sub OnBallBallCollision(ball1, ball2, velocity)
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0
End Sub


'******************************
' Diverse Collection Hit Sounds
'******************************


Sub Pins_Hit (idx)
	PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
End Sub

Sub Targets_Hit (idx)
	PlaySound SoundFX("target",DOFContactors), 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
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

Sub Gates_Hit (idx)
	PlaySound "gate4", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

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
		Case 1 : PlaySound "flip_hit_1", 0, 1, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 2 : PlaySound "flip_hit_2", 0, 1, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 3 : PlaySound "flip_hit_3", 0, 1, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End Select
End Sub

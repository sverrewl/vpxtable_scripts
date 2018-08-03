'***************************************************************************
' New RingBall Shot - Thanks to SliderPoint for helping me correct the issue.
'****************************************************************************
' To use the Rom for Sound
' remove line 117 --> ' Set MotorCallback = GetRef("TrackSounds")
' change the 0 for 1 in the line 124 --> .Games(cGameName).Settings.Value("sound") = 1

 Option Explicit
    Randomize


  LoadVPM "02000000", "sam.VBS", 3.15

 '******************* Options *********************
' DMD/Backglass Controller Setting
Const cController = 1		'0=Use value defined in cController.txt, 1=VPinMAME, 2=UVP server, 3=B2S server, 4=B2S with DOF (disable VP mech sounds)
'*************************************************

Dim cNewController
Sub LoadVPM(VPMver, VBSfile, VBSver)
	Dim FileObj, ControllerFile, TextStr

	On Error Resume Next
	If ScriptEngineMajorVersion < 5 Then MsgBox "VB Script Engine 5.0 or higher required"
	ExecuteGlobal GetTextFile(VBSfile)
	If Err Then MsgBox "Unable to open " & VBSfile & ". Ensure that it is in the same folder as this table. " & vbNewLine & Err.Description

	cNewController = 1
	If cController = 0 then
		Set FileObj=CreateObject("Scripting.FileSystemObject")
		If Not FileObj.FolderExists(UserDirectory) then
			Msgbox "Visual Pinball\User directory does not exist. Defaulting to vPinMame"
		ElseIf Not FileObj.FileExists(UserDirectory & "cController.txt") then
			Set ControllerFile=FileObj.CreateTextFile(UserDirectory & "cController.txt",True)
			ControllerFile.WriteLine 1: ControllerFile.Close
		Else
			Set ControllerFile=FileObj.GetFile(UserDirectory & "cController.txt")
			Set TextStr=ControllerFile.OpenAsTextStream(1,0)
			If (TextStr.AtEndOfStream=True) then
				Set ControllerFile=FileObj.CreateTextFile(UserDirectory & "cController.txt",True)
				ControllerFile.WriteLine 1: ControllerFile.Close
			Else
				cNewController=Textstr.ReadLine: TextStr.Close
			End If
		End If
	Else
		cNewController = cController
	End If

	Select Case cNewController
		Case 1
			Set Controller = CreateObject("VPinMAME.Controller")
			If Err Then MsgBox "Can't Load VPinMAME." & vbNewLine & Err.Description
			If VPMver>"" Then If Controller.Version < VPMver Or Err Then MsgBox "VPinMAME ver " & VPMver & " required."
			If VPinMAMEDriverVer < VBSver Or Err Then MsgBox VBSFile & " ver " & VBSver & " or higher required."
		Case 2
			Set Controller = CreateObject("UltraVP.BackglassServ")
		Case 3,4
			Set Controller = CreateObject("B2S.Server")
	End Select
	On Error Goto 0
End Sub

'*************************************************************
'Toggle DOF sounds on/off based on cController value
'*************************************************************
Dim ToggleMechSounds
Function SoundFX (sound)
    If cNewController= 4 and ToggleMechSounds = 0 Then
        SoundFX = ""
    Else
        SoundFX = sound
    End If
End Function

Sub DOF(dofevent, dofstate)
	If cNewController>2 Then
		If dofstate = 2 Then
			Controller.B2SSetData dofevent, 1:Controller.B2SSetData dofevent, 0
		Else
			Controller.B2SSetData dofevent, dofstate
		End If
	End If
End Sub

'******************************
'***START DONOR TABLE SCRIPT***
'******************************


'********************
'Standard definitions
'********************

Const MagnetControlsRingEntrance = 0	'1 - Magnet will control Ring entrance, 0 - Ball can enter Ring as long as ball has enough velocity

Const cGameName = "lotr"
Const UseSolenoids = 15
Const UseLamps = 0
Const UseGI = 0
Const UseSync = 1
Const HandleMech = 0

Const SSolenoidOn = "SolOn"
Const SSolenoidOff = "SolOff"
Const SFlipperOn = "FlipperUp"
Const SFlipperOff = "FlipperDown"
Const SCoin = "Coin"

Const rVol = 1.0	'0.0 -> 1.0 Loudness of rubber impact
Const fVol = 0.5	'0.0 -> 1.0 Loudness of flipper collisions
Const sVol = 0.1	'0.0 -> 1.0 Loudness of switches

'************
' Table init.
'************

Dim bsTrough, bsL, bsR, bsTR, bsTL, vlLock, bsRing
Dim bump1, bump2, bump3, plungerIM, TowerDir, TowerStep, DiverterDir, DiverterPos
Dim Controller   ' VPinMAME Controller Object
Dim vpmTimer     ' Timer Object
Dim vpmNudge     ' Nudge handler Object

Set MotorCallback = GetRef("TrackSounds") 'disable this line if you are using the rom for sound

Sub Table1_Init
	vpmInit Me
	With Controller
		.GameName = cGameName
		.Games(cGameName).Settings.Value("sound") = 0 'turn off the ROM sound and use samples instead
		If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
		.SplashInfoLine = "Lord of the Rings (Stern 2003)"
		.HandleKeyboard = 0
		.ShowTitle = 0
		.ShowDMDOnly = 1
		.ShowFrame = 0
		.Hidden = 0
		On Error Resume Next
		.Run GetPlayerHWnd
		If Err Then MsgBox Err.Description

		On Error Goto 0
	End With

'*****************
' Nudging
'*****************
    vpmNudge.TiltSwitch = 56
    vpmNudge.Sensitivity = 1
    vpmNudge.TiltObj = Array(lbumper, bbumper, rbumper, LeftSlingshot, RightSlingshot)

'*****************
' Trough
'*****************
    Set bsTrough = New cvpmBallStack
    bsTrough.InitSw 0, 14, 13, 12, 11, 0, 0, 0
    bsTrough.InitKick BallRelease, 90, 10
    bsTrough.InitEntrySnd "fx_Solenoid", "fx_Solenoid"
    bsTrough.InitExitSnd "fx_Ballrel", "fx_Solenoid"
    bsTrough.Balls = 4

'*****************
' Left vuk
'*****************
    Set bsL = New cvpmBallStack
    bsL.InitSaucer sw9, 9, 270, 35
    bsL.KickZ = 1.3
    bsL.InitExitSnd "fx_popper", "fx_Solenoid"

'*****************
    ' Right vuk
'*****************
    Set bsR = New cvpmBallStack
    bsR.InitSaucer sw30, 30, 75, 80
    bsR.KickZ = 1.5
    bsR.InitExitSnd "fx_popper", "fx_Solenoid"

'******************
' Top Right Saucer
'******************
    Set bsTR = New cvpmBallStack
    bsTR.InitSaucer sw46, 46, 270, 4
    bsTR.KickForceVar = 2
    bsTR.InitExitSnd "fx_popper_ball", "fx_popper"

'******************
' Top Left vuk
'******************
    Set bsTL = New cvpmBallStack
    bsTL.InitSw 0, 41, 0, 0, 0, 0, 0, 0
    bsTL.InitKick sw41, 270, 4
    bsTL.InitExitSnd "fx_popper", "fx_Solenoid"
    bsTL.KickForceVar = 2

'*****************************************
' Visible Lock - implements post ball lock
'*****************************************
'    Set vlLock = New cvpmVLock
'    vlLock.InitVLock Array(sw19, sw18, sw17), Array(k17,k18,k19), Array(19, 18, 17)
'    vlLock.InitSnd "sensor", "sensor"
'    vlLock.CreateEvents "vlLock"
'
	k17.Enabled=False
	k18.Enabled=False
	k19.Enabled=True

If MagnetControlsRingEntrance = 0 Then ringblock.isDropped = 1



'*********************
' Main Timer init
'*********************
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1


 End Sub

'**********************************
'**********    Div Init   *********
'**********************************
    Ring1a.IsDropped = 1:Ring2a.IsDropped = 1:Ring3a.IsDropped = 1
    Ring1b.IsDropped = 1:Ring2b.IsDropped = 1:Ring3b.IsDropped = 1
    Ring1c.IsDropped = 1:Ring2c.IsDropped = 1:Ring3c.IsDropped = 1
    Ring1d.IsDropped = 1:Ring2d.IsDropped = 1:Ring3d.IsDropped = 1
	DiverterPos = 0
    TowerStep = 1
    OrbitPin.IsDropped = 1


Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub
Sub Table_exit ()

End Sub

'********************************************
'**********   Keys   ************************
'********************************************


Sub Table1_KeyDown(ByVal KeyCode)

	If keycode = StartGameKey Then
	Controller.Switch(54) = 1
	End If
	If Keycode = AddCreditKey Then
	Controller.Switch(6) = 1
	End If
	If KeyCode = PlungerKey Then Plunger.Pullback
	If KeyCode = RightFlipperKey then Controller.Switch(swLRFlip) = True: vpmFlips.FlipR True: Exit Sub
	if KeyCode = LeftFlipperKey  then Controller.Switch(swLLFlip) = True :vpmFlips.FlipL True: Exit Sub
	If KeyDownHandler(KeyCode) Then Exit Sub

End Sub

Sub Table1_KeyUp(ByVal KeyCode)

	If keycode = StartGameKey Then
	Controller.Switch(54) = 0
	Exit Sub
End If
	If Keycode = AddCreditKey Then
	Controller.Switch(6) = 0
	Exit Sub
End If
	If KeyCode = PlungerKey Then Plunger.Fire:Playsound "fx_plunger"
	If KeyCode = RightFlipperKey then Controller.Switch(swLRFlip) = False :vpmFlips.FlipR False: Exit Sub
	if KeyCode = LeftFlipperKey  then Controller.Switch(swLLFlip) = False :vpmFlips.FlipL False: Exit Sub

	If KeyUpHandler(KeyCode) Then Exit Sub
	End Sub

''******************************************************************************
''*******             solenoid callbacks                ************************
''******************************************************************************


SolCallBack(1) = "SolRelease"
SolCallBack(2) = "SolAutofire"
SolCallback(3) = "bsL.SolOut"
SolCallback(4) = "bsTL.SolOut"
SolCallback(5) = "bsR.SolOut"
'SolCallback(6) = "SolRM"
SolCallback(7) = "SolTower"
SolCallback(8) = "SolDiv"
SolCallback(13) = "SolOrbit"
SolCallback(14) = "SetLamp 114,"
SolCallback(sLLFlipper) = "SolLFlipper"
SolCallback(sLRFlipper) = "SolRFlipper"
SolCallBack(19) = "bsTR.SolOut"
SolCallBack(21) = "SolLockRelease"
SolCallBack(22) = "SolBalrog"
SolCallBack(23) = "SetLamp 123,"
SolCallBack(24) = "vpmsolsound ""knocker"","
SolCallBack(25) = "SetLamp 125,"
SolCallBack(26) = "SetLamp 126,"
SolCallBack(27) = "SetLamp 127,"
SolCallBack(29) = "SetLamp 129,"
SolCallBack(30) = "SetLamp 130,"
SolCallBack(31) = "SetLamp 19,"
'SolCallBack(32) = "SolFlash32"


Sub SolFlash32(Enabled)
    If Enabled Then
        BalrogFlash = 12
    Else
        BalrogFlash = 0
   End If
   UpdateBalrog
End Sub

Sub SolLockRelease(enabled)
    vlLock.SolExit enabled
End Sub

Sub SolLockRelease(Enabled)
	If Enabled Then
 		k19.enabled = False:k19.kick 180,.01: Playsound "fx_solenoidon"
	Else
		k19.enabled = True: Playsound "fx_solenoidoff"
	End If
End Sub

Sub SolRelease(Enabled)
    If Enabled And bsTrough.Balls > 0 Then
        bsTrough.ExitSol_On
        vpmTimer.PulseSw 15
    End If
End Sub

Sub SolOrbit(Enabled):
	OrbitPin.IsDropped = Enabled
	If Enabled Then
		Playsound "fx_solenoidon"
	Else
		Playsound "fx_solenoidoff"
	End If
End Sub

''********************************************************************************************
''**************               Plunger             *******************************************
''********************************************************************************************

Sub swPlunger_Hit
	BallinPlunger = 1
End Sub

Sub swPlunger_UnHit
	BallinPlunger = 0
End Sub

Sub SolAutofire(Enabled)
	If Enabled Then
		PlungerIM.AutoFire
	End If
End Sub

    Const IMPowerSetting = 60 ' Plunger Power
    Const IMTime = 0.6        ' Time in seconds for Full Plunge
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swplunger, IMPowerSetting, IMTime
        .Random 0.3
        .InitExitSnd "fx_plunger2", "fx_plunger"
        .CreateEvents "plungerIM"
    End With


''************************************************************************************
''*****************       SLING Subs                      ****************************
''************************************************************************************

 Sub LeftSlingShot_Slingshot
    vpmTimer.PulseSw 59
 	PlaySound SoundFX("fx_slingshot2")
	Me.TimerEnabled = 1
  End Sub

 Sub LeftSlingShot_Timer

 End Sub

 Sub RightSlingShot_Slingshot
    vpmTimer.PulseSw 62
 	PlaySound SoundFX("fx_slingshot2")
	Me.TimerEnabled = 1
  End Sub

 Sub RightSlingShot_Timer

 End Sub

''************************************************************************************
''*****************               Bumpers                 ****************************
''************************************************************************************

      Sub LBumper_Hit
      vpmTimer.PulseSw 49
      PlaySound SoundFX("fx_bumper1")
    	End Sub


      Sub BBumper_Hit
      vpmTimer.PulseSw 51
      PlaySound SoundFX("fx_bumper1")
       End Sub

      Sub RBumper_Hit
      vpmTimer.PulseSw 50
      PlaySound SoundFX("fx_bumper1")
       End Sub

'*********************************************************************
'************         Drain Holes, Vuks, Saucers       ***************
'*********************************************************************

Dim BallCount:BallCount = 0
   Sub Drain_Hit():PlaySound "fx_Drain"
	BallCount = BallCount - 1
	bsTrough.AddBall Me
	If BallCount = 0 then GIOff
   End Sub
   Sub BallRelease_UnHit()
		BallCount = BallCount + 1
		GIOn
	End Sub

Sub sw9_Hit:Playsound "fx_kicker_enter":bsL.AddBall 0:End Sub
Sub sw30_Hit:Playsound "fx_kicker_enter":bsR.AddBall 0: sw30.Enabled = False:End Sub

Sub sw30Trigger_Hit
	If ActiveBall.VelY < -40 Then
		sw30.Enabled = False
	Else
		sw30.Enabled = True
	End If
End Sub

Sub sw41a_Hit():Playsound "fx_hole_enter"
	'ClearBallID
	bsTL.AddBall Me
	if BallCount > 0 Then
	GIOff
	End If
	End Sub
Sub sw41b_Hit():Playsound "fx_hole_enter"
	'ClearBallID
	bsTL.AddBall Me
	if BallCount > 0 Then
	GIOff
	End If
	End Sub
Sub sw41c_Hit():Playsound "fx_hole_enter"
	'ClearBallID
	bsTL.AddBall Me
	if BallCount > 0 Then
	GIOff
	End If
	End Sub
Sub sw41d_Hit():Playsound "fx_hole_enter"
	'ClearBallID
	bsTL.AddBall Me
	if BallCount > 0 Then
	GIOff
	End If
	End Sub
Sub sw46_Hit():Playsound "fx_kicker_enter"
	bsTR.AddBall 0
	End Sub

Sub PathGI_Hit():GIOn:End Sub
''************************************************************************************
''*****************               Flippers                ****************************
''************************************************************************************


SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"


Sub SolLFlipper(Enabled)
	Dim tmp
     If Enabled Then
		 PlaySound "fx_flipperup1"
		 LeftFlipper.RotateToEnd
     Else
		 PlaySound "fx_Flipperup2"
		 LeftFlipper.RotateToStart
     End If
 End Sub

Sub SolRFlipper(Enabled)
	Dim tmp
     If Enabled Then
		 PlaySound "fx_flipperup1"
		 RightFlipper.RotateToEnd
     Else
		 PlaySound "fx_Flipperup2"
		 RightFlipper.RotateToStart
     End If
 End Sub


'******************************************************************
'******************      Rollovers & Ramp Switches        *********
'******************************************************************

Sub sw57_Hit:Controller.Switch(57) = 1:Playsound "fx_sensor":End Sub
Sub sw57_UnHit:Controller.Switch(57) = 0:End Sub

Sub sw58_Hit:Controller.Switch(58) = 1:Playsound "fx_sensor":End Sub
Sub sw58_UnHit:Controller.Switch(58) = 0:End Sub

Sub sw60_Hit:Controller.Switch(60) = 1:Playsound "fx_sensor":End Sub
Sub sw60_Unhit:Controller.Switch(60) = 0:End Sub

Sub sw61_Hit:Controller.Switch(61) = 1:Playsound "fx_sensor":End Sub
Sub sw61_Unhit:Controller.Switch(61) = 0:End Sub

Sub sw20_Hit:Controller.Switch(20) = 1:Playsound "fx_sensor":End Sub
Sub sw20_UnHit:Controller.Switch(20) = 0:End Sub

Sub sw21_Hit:Controller.Switch(21) = 1:Playsound "fx_sensor":End Sub
Sub sw21_Unhit:Controller.Switch(21) = 0:End Sub

Sub sw43_Hit:Controller.Switch(43) = 1:Playsound "fx_sensor":End Sub
Sub sw43_Unhit:Controller.Switch(43) = 0:End Sub

Sub sw44_Hit:Controller.Switch(44) = 1:Playsound "fx_sensor":End Sub
Sub sw44_Unhit:Controller.Switch(44) = 0:End Sub

Sub sw45_Hit:Controller.Switch(45) = 1:Playsound "fx_sensor":End Sub
Sub sw45_Unhit:Controller.Switch(45) = 0:End Sub

Sub sw37_Hit:Controller.Switch(37) = 1:Playsound "fx_sensor":End Sub
Sub sw37_Unhit:Controller.Switch(37) = 0:End Sub

Sub sw38_Hit:Controller.Switch(38) = 1:Playsound "fx_sensor":End Sub
Sub sw38_Unhit:Controller.Switch(38) = 0:End Sub

Sub sw39_Hit:Controller.Switch(39) = 1:Playsound "fx_sensor":End Sub
Sub sw39_Unhit:Controller.Switch(39) = 0:End Sub

Sub sw33_Hit:Controller.Switch(33) = 1:Playsound "fx_sensor":End Sub
Sub sw33_Unhit:Controller.Switch(33) = 0:End Sub

Sub sw34_Hit:Controller.Switch(34) = 1:Playsound "fx_sensor":End Sub
Sub sw34_Unhit:Controller.Switch(34) = 0:End Sub

Sub sw35_Hit:Controller.Switch(35) = 1:Playsound "fx_sensor":End Sub
Sub sw35_Unhit:Controller.Switch(35) = 0:End Sub

Sub sw36_Hit:Controller.Switch(36) = 1:Playsound "fx_sensor":End Sub
Sub sw36_Unhit:Controller.Switch(36) = 0:End Sub

Sub sw40_Hit:Controller.Switch(40) = 1:Playsound "fx_sensor":End Sub
Sub sw40_Unhit:Controller.Switch(40) = 0:End Sub

Sub sw42_Hit:Controller.Switch(42) = 1:Playsound "fx_sensor":End Sub
Sub sw42_Unhit:Controller.Switch(42) = 0:End Sub

Sub sw24_Hit:Controller.Switch(24) = 1:Playsound "fx_sensor":End Sub
Sub sw24_Unhit:Controller.Switch(24) = 0:End Sub

Sub sw22_Hit:Controller.Switch(22) = 1:Playsound "fx_sensor":End Sub
Sub sw22_Unhit:Controller.Switch(22) = 0:End Sub

Sub sw48_Hit:Controller.Switch(48) = 1:Playsound "fx_sensor":End Sub
Sub sw48_Unhit:Controller.Switch(48) = 0:End Sub 'ActiveBall.VelX = 3:End Sub

Sub sw16_Hit:Controller.Switch(16) = 1:Playsound "fx_sensor":End Sub
Sub sw16_Unhit:Controller.Switch(16) = 0:End Sub

'*******************************************************************
'*****************        Visible lock         ********************
'*******************************************************************

Sub sw17_hit()  :Controller.Switch(17) = true:  end sub
Sub sw17_unhit():Controller.Switch(17) = false: end sub
Sub sw18_hit()  :Controller.Switch(18) = true:  k17.enabled = True:end sub
Sub sw18_unhit():Controller.Switch(18) = false: k17.kick 180, .01:k17.enabled = False: end sub
Sub sw19_hit()  :Controller.Switch(19) = true:  k18.enabled = True:end sub
Sub sw19_unhit():Controller.Switch(19) = false: k18.kick 180, .01:k18.enabled = False:end sub


Sub sw25_Hit
    Controller.Switch(25) = 1
    If ActiveBall.VelY < -20 Then ActiveBall.VelY = -20
End Sub

Sub sw25_Unhit:Controller.Switch(25) = 0:End Sub


'*******************************************************************
' *************************       Targets        *******************
'*******************************************************************
Sub sw10_Hit:vpmTimer.PulseSw 10:PlaySound "fx_target":End Sub
Sub sw23_Hit:vpmTimer.PulseSw 23:PlaySound "fx_target":End Sub
Sub sw28_Hit:BalrogHit 2:vpmTimer.PulseSw 28:PlaySound "fx_metalhit":End Sub
Sub sw29_Hit:vpmTimer.PulseSw 29:PlaySound "fx_target":End Sub
Sub sw53_Hit:vpmTimer.PulseSw 53:PlaySound "fx_target":End Sub

Sub sw52_Spin():vpmTimer.PulseSw 52:Playsound "fx_spinner":End Sub

'***********************************************************************
'************************        Ramps helpers     *********************
'***********************************************************************

'Sub RHelp1_Hit()
'    ActiveBall.VelZ = -2
'    ActiveBall.VelY = 0
'    ActiveBall.VelX = 0
'    StopSound "fx_metallrolling"
'    PlaySound "fx_ballhit"
'End Sub
'
'Sub RHelp2_Hit()
'    ActiveBall.VelZ = -2
'    ActiveBall.VelY = 0
'    ActiveBall.VelX = 0
'    StopSound "fx_metallrolling"
'    PlaySound "fx_ballhit"
'End Sub
'
'Sub RHelp3_Hit
'    ActiveBall.VelZ = -0.1
'    If ActiveBall.VelY > -8 Then ActiveBall.VelY = -8
'End Sub
'
Sub RHelp4_Hit()
    ActiveBall.VelZ = -2
    ActiveBall.VelY = 0
    ActiveBall.VelX = 0
    PlaySound "fx_ballhit"
End Sub
'
'Sub RHelp5_Hit
'    If ActiveBall.VelY < -20 Then ActiveBall.VelY = -20
'    If ActiveBall.VelZ > 2 Then ActiveBall.VelZ = 2
'End Sub
'
Sub RHelp6_Hit
    PlaySound "rail",0,.05
'    If ActiveBall.VelY > 20 Then ActiveBall.VelY = 20
'    If ActiveBall.VelY < 12 Then ActiveBall.VelY = 12
'    ActiveBall.VelX = 0
End Sub

'*******************************************************************
'*****************        Ring Magnet        ***********************
'*******************************************************************


Dim RingBall:RingBall = 0
Dim RingShot:RingShot = 0
Dim mRingMagnet

	Set mRingMagnet = New cvpmMagnet
 	With mRingMagnet
		.InitMagnet sw47a, 30
		.GrabCenter = True
 		.solenoid = 6						'Ring Magnet
		.CreateEvents "mRingMagnet"
	End With

Sub sw47_Hit:Controller.Switch(47) = 1:PlaySound "fx_kicker_enter"
'    If RingBall Then
'        sw47a.DestroyBall
'		sw48.CreateBall
'		sw48.kick 90, 3, 0
'		PlaySound "fx_popper"
'        RingBall = 0
'		Ringshot = 1
'	Else
'	If Ringshot Then
'		sw47.DestroyBall
'		sw48.CreateBall
'		sw48.kick 90, 3, 0
'		PlaySound "fx_popper"
'		RingBall = 0
'		Ringshot = 0
'	End If
'	End If
End Sub

Sub sw47_UnHit:Controller.Switch(47) = 0:End Sub

'Sub SolRM(Enabled)
'	    mRingMagnet.MagnetOn = 1
''	debug.print timer & "magnet sol " & enabled
'    If Enabled Then
'        sw47a.Enabled = 1
'    Else
'        sw47a.Enabled = 1
'		If RingBall Then
'            sw47a.DestroyBall
'			sw48.CreateBall
'			sw48.kick 90, 3, 0
'			PlaySound "fx_popper"
'			Ringball = 0
'			Ringshot = 1
'		Else
'		If Ringshot Then
'			sw47.DestroyBall
'			sw48.CreateBall
'			sw48.kick 90, 3, 0
'			PlaySound "fx_popper"
'			RingBall = 0
'			RingShot = 0
'			End If
'		End If
'	End If
'End Sub

'Sub sw47_Hit:RingBall = 1:Ringshot=0:PlaySound "fx_kicker_enter":End Sub

'************************
' Tower animation
'************************

Sub SolTower(Enabled)
    If Enabled Then
        TowerDir = 1
    Else
        TowerDir = -1
    End If

    TowerAnim.Enabled = 0
    If TowerStep < 1 Then TowerStep = 1
    If TowerStep > 20 Then TowerStep = 20

    TowerAnim.Enabled = 1
End Sub
'
Sub TowerAnim_Timer()
    Select Case TowerStep
        Case 0:Primitive46.RotX=90:TowerAnim.Enabled = 0
        Case 1:Primitive46.RotX=88
        Case 2:Primitive46.RotX=86
        Case 3:Primitive46.RotX=84
        Case 4:Primitive46.RotX=82
        Case 5:Primitive46.RotX=80
        Case 6:Primitive46.RotX=78
        Case 7:Primitive46.RotX=76
        Case 8:Primitive46.RotX=74
		Case 9:Primitive46.RotX=72
		Case 10:Primitive46.RotX=70
		Case 11:Primitive46.RotX=72
		Case 12:Primitive46.RotX=74
		Case 13:Primitive46.RotX=76
		Case 14:Primitive46.RotX=78
		Case 15:Primitive46.RotX=80
		Case 16:Primitive46.RotX=82
		Case 17:Primitive46.RotX=84
		Case 18:Primitive46.RotX=86
		Case 19:Primitive46.RotX=88
		Case 20:Primitive46.RotX=90
        Case 21:TowerAnim.Enabled = 0
    End Select
'
    TowerStep = TowerStep + TowerDir
End Sub

'************************
' Diverter animation
'************************

Sub SolDiv(Enabled)
    If Enabled Then
        DiverterDir = 1
    Else
        DiverterDir = -1
    End If

	Diverter.Enabled = 0
    If DiverterPos < 1 Then DiverterPos = 1
    If DiverterPos > 4 Then DiverterPos = 4

    Diverter.Enabled = 1
End Sub

Sub Diverter_Timer()
    Select Case DiverterPos
        Case 0:Diverter.Enabled = 0
        Case 1:Div1.ObjRotZ = -44
        Case 2:Div1.ObjRotZ = -52
        Case 3:Div1.ObjRotZ = -60
        Case 4:Div1.ObjRotZ = -68
        Case 5:Diverter.Enabled = 0

    End Select

    DiverterPos = DiverterPos + DiverterDir
End Sub

''***************************************************************
''*******************    Balrog     *****************************
''***************************************************************
'
Dim BalrogPos, BalrogDir, BalrogFlash
BalrogDir = 0:BalrogPos = 0:BalrogFlash = 0
sw28.IsDropped = 1

Sub SolBalrog(Enabled)
    If Enabled Then
        If BalrogDir = 0 Then
            Controller.Switch(32) = 0
            Controller.Switch(31) = 1
            BalrogClose.Enabled = 0
            BalrogOpen.Enabled = 1
            BalrogOpen_Timer
            BalrogDir = 1
        Else
            Controller.Switch(31) = 0
            Controller.Switch(32) = 1
            BalrogOpen.Enabled = 0
            BalrogClose.Enabled = 1
            BalrogClose_Timer
            BalrogDir = 0
        End If
    End If
End Sub

Sub BalrogOpen_Timer()
    UpdateBalrog
    BalrogPos = BalrogPos + 1

    If BalrogPos > 45 Then
        BalrogPos = 45
        BalrogOpen.Enabled = 0
    End If
End Sub

Sub BalrogClose_Timer()
    UpdateBalrog
    BalrogPos = BalrogPos - 1

    If BalrogPos < 0 Then
        BalrogPos = 0
        BalrogClose.Enabled = 0
    End If
End Sub

Sub UpdateBalrog
    Select Case BalrogPos
		Case 0:Balrog.ObjRotZ=0:sw28.IsDropped = 0
		Case 1:Balrog.ObjRotZ=-2:sw28.IsDropped = 0
		Case 2:Balrog.ObjRotZ=-4:sw28.IsDropped = 1
		Case 3:Balrog.ObjRotZ=-6:sw28.IsDropped = 1
		Case 4:Balrog.ObjRotZ=-8:sw28.IsDropped = 1
		Case 5:Balrog.ObjRotZ=-10:sw28.IsDropped = 1
		Case 6:Balrog.ObjRotZ=-12:sw28.IsDropped = 1
		Case 7:Balrog.ObjRotZ=-14:sw28.IsDropped = 1
		Case 8:Balrog.ObjRotZ=-16:sw28.IsDropped = 1
		Case 9:Balrog.ObjRotZ=-18:sw28.IsDropped = 1
		Case 10:Balrog.ObjRotZ=-20:sw28.IsDropped = 1
		Case 11:Balrog.ObjRotZ=-22:sw28.IsDropped = 1
		Case 12:Balrog.ObjRotZ=-24:sw28.IsDropped = 1
		Case 13:Balrog.ObjRotZ=-26:sw28.IsDropped = 1
		Case 14:Balrog.ObjRotZ=-28:sw28.IsDropped = 1
		Case 15:Balrog.ObjRotZ=-30:sw28.IsDropped = 1
		Case 16:Balrog.ObjRotZ=-32:sw28.IsDropped = 1
		Case 17:Balrog.ObjRotZ=-34:sw28.IsDropped = 1
		Case 18:Balrog.ObjRotZ=-36:sw28.IsDropped = 1
		Case 19:Balrog.ObjRotZ=-38:sw28.IsDropped = 1
		Case 20:Balrog.ObjRotZ=-40:sw28.IsDropped = 1
		Case 21:Balrog.ObjRotZ=-42:sw28.IsDropped = 1
		Case 22:Balrog.ObjRotZ=-44:sw28.IsDropped = 1
		Case 23:Balrog.ObjRotZ=-46:sw28.IsDropped = 1
		Case 24:Balrog.ObjRotZ=-48:sw28.IsDropped = 1
		Case 25:Balrog.ObjRotZ=-50:sw28.IsDropped = 1
		Case 26:Balrog.ObjRotZ=-52:sw28.IsDropped = 1
		Case 27:Balrog.ObjRotZ=-54:sw28.IsDropped = 1
		Case 28:Balrog.ObjRotZ=-56:sw28.IsDropped = 1
		Case 29:Balrog.ObjRotZ=-58:sw28.IsDropped = 1
		Case 30:Balrog.ObjRotZ=-60:sw28.IsDropped = 1
		Case 31:Balrog.ObjRotZ=-62:sw28.IsDropped = 1
		Case 32:Balrog.ObjRotZ=-64:sw28.IsDropped = 1
		Case 33:Balrog.ObjRotZ=-66:sw28.IsDropped = 1
		Case 34:Balrog.ObjRotZ=-68:sw28.IsDropped = 1
		Case 35:Balrog.ObjRotZ=-70:sw28.IsDropped = 1
		Case 36:Balrog.ObjRotZ=-72:sw28.IsDropped = 1
		Case 37:Balrog.ObjRotZ=-74:sw28.IsDropped = 1
		Case 38:Balrog.ObjRotZ=-76:sw28.IsDropped = 1
		Case 39:Balrog.ObjRotZ=-78:sw28.IsDropped = 1
		Case 40:Balrog.ObjRotZ=-80:sw28.IsDropped = 1
		Case 41:Balrog.ObjRotZ=-82:sw28.IsDropped = 1
		Case 42:Balrog.ObjRotZ=-84:sw28.IsDropped = 1
		Case 43:Balrog.ObjRotZ=-86:sw28.IsDropped = 1
		Case 44:Balrog.ObjRotZ=-88:sw28.IsDropped = 1
		Case 45:Balrog.ObjRotZ=-90:sw28.IsDropped = 1
	End Select
End Sub

dim BalrogBall,zz
set BalrogBall = kicker1.createball
kicker1.kick 0,0
Dim Mag1
	Set mag1= New cvpmMagnet
 	With mag1
		.InitMagnet BalrogHitTrigger, 25
		.GrabCenter = False
		.magnetOn = True
	End With

Sub BalrogHit (Vmove)
	dim yy
	BalrogBall.vely = Vmove / 2
	yy = (rnd(1) - 0.5) * Vmove
	BalrogBall.velx = yy
	BalrogHitTimer.Enabled = 1:zz=0
	'Playsound "solenoidon"
End Sub


Sub BalrogHitTimer_Timer
	Balrog.ObjRotX = (BalrogBall.x - BalrogHitTrigger.x) * 1
	zz = zz + 1
	if zz = 100 then BalrogHitTimer.Enabled = 0
End Sub
'*******************************************************************
'****************        LAMPS AND FLASHERS          ***************
'*******************************************************************

Dim LampState(200), FadingLevel(200), FadingState(200)
Dim FlashState(200), FlashLevel(200), FlashMin(200), FlashMax(200)
Dim FlashSpeedUp, FlashSpeedDown
Dim x

AllLampsOff()
LampTimer.Interval = 40
LampTimer.Enabled = 1

FlashInIt()
FlasherTimer.Interval = 10
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

Sub UpdateLamps
    NFadeL 1, l1
    NFadeL 2, l2
    NFadeL 3, l3
    NFadeL 4, l4
    NFadeL 5, l5
    NFadeL 6, l6
    NFadeL 7, l7
    NFadeL 8, l8
    NFadeL 9, l9
    NFadeL 10, l10
    NFadeL 11, l11
    NFadeL 12, l12
    NFadeL 13, l13
    NFadeL 14, l14
    NFadeL 15, l15
    NFadeL 16, l16
    NFadeL 17, l17
    NFadeL 18, l18
    NFadeL 19, l19
    NFadeL 20, l20
    NFadeL 21, l21
    NFadeL 22, l22
    NFadeL 23, l23
    NFadeL 24, l24
    NFadeL 25, l25
    NFadeL 26, l26
    NFadeL 27, l27
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
    NFadeL 38, l38
    NFadeL 39, l39
    NFadeL 40, l40
    NFadeL 41, l41
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
	NFadeL 65, l65
    NFadeL 66, l66
    NFadeL 67, l67
    NFadeL 68, l68
    NFadeL 69, l69
    NFadeL 70, l70
    NFadeL 71, l71
    NFadeL 72, l72
    NFadeL 73, l73
    NFadeL 74, l74
    NFadeL 75, l75
    NFadeL 76, l76
    NFadeL 77, l77
    NFadeL 78, l78
    NFadeL 81, l98
    NFadeL 82, l97
    NFadeL 83, l96
    NFadeL 84, l95
    NFadeL 85, l94
    NFadeL 86, l93
    NFadeL 87, l92
    NFadeL 88, l91
    NFadeL 89, l99
    NFadeL 90, l86
    NFadeL 91, l87
    NFadeL 92, l88
    NFadeL 93, l89
    NFadeL 94, l90
    NFadeL 95, l84
    NFadeL 96, l85
    NFadeL 97, l82
    NFadeL 98, l81
    NFadeL 99, l83

	NFadeL 130, f130

    NFadeLm 125, f25a
    NFadeLm 125, f25b
    NFadeLm 125, f25c
	NFadeLm	130, F130c
'    FadeWm 125, f25a, f25aa, f25ab
'    FadeWm 125, f25b, f25ba, f25bb
'    FadeW 125, f25c, f25ca, f25cb
    NFadeLm 126, f26
	NFadeLm 126, F26a
'    FadeARm 123, f23b, "fy_on", "fy_a", "fy_b", "fy", dummy
    NFadeLM 123, f23
    NFadeLm 114, f14
    NFadeLm 127, f27
    NFadeLm 129, f29
End Sub

Sub FadePrim(nr, pri, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:pri.image = d:FadingLevel(nr) = 0
        Case 3:pri.image = c:FadingLevel(nr) = 1
        Case 4:pri.image = b:FadingLevel(nr) = 2
        Case 5:pri.image = a:FadingLevel(nr) = 3
    End Select
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

' Flasher objects
' Uses own faster timer

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
            If FlashLevel(nr) > 1000 Then
                FlashLevel(nr) = 1000
                FlashState(nr) = -2 'completely on
            End if
            Object.opacity = FlashLevel(nr)
    End Select
End Sub

 Sub AllLampsOff():For x = 1 to 200:LampState(x) = 4:FadingLevel(x) = 4:Next:UpdateLamps:UpdateLamps:Updatelamps:End Sub


Sub SetLamp(nr, value)
    If value = 0 AND LampState(nr) = 0 Then Exit Sub
    If value = 1 AND LampState(nr) = 1 Then Exit Sub
    LampState(nr) = abs(value) + 4
FadingLevel(nr ) = abs(value) + 4: FadingState(nr ) = abs(value) + 4
End Sub

'
''*********************************************************************
''******************      FLASHER OBJECTS     *************************
''*********************************************************************

Sub SetFlash(nr, stat)
    FlashState(nr) = ABS(stat)
End Sub

Sub FlasherTimer_Timer ()
End Sub

Sub GIOn
	dim bulb
	for each bulb in Collection1
	bulb.state = 1
	next
End Sub

Sub GIOff
	dim bulb
	for each bulb in Collection1
	bulb.state = 0
	next
End Sub

'********************************
' Sound Subs from Destruk's table
'********************************

Dim Playing
Playing = 0

'Music & Sound Stuff
Sub TrackSounds
    Dim NewSounds, ii, Snd
    NewSounds = Controller.NewSoundCommands
    If Not IsEmpty(NewSounds) Then
        For ii = 0 To UBound(NewSounds)
            Snd = NewSounds(ii, 0)
            If Snd = 254 Then Playing = 3 'FE
            If Snd = 253 Then Playing = 2 'FD
            If Snd = 252 Then Playing = 1 'FC
            If Snd <> 38 And Snd <> 255 And Snd <> 1 And Snd <> 0 And Snd <> 16 Then SoundCommand(Snd)
        Next
    End If
End Sub

Sub SoundCommand(Cmd)
    Dim SndName
    If Playing = 3 Then SndName = "FE"
    If Playing = 2 Then SndName = "FD"
    If Playing = 1 Then SndName = "FC"
    If Playing = 2 And Cmd < 40 And Cmd <> 5 And Cmd <> 31 Then 'Ignore FD1F - unknown command, FD5 handled after FD4 is played
        MusicCommand(Cmd)
    Else
        Dim FinalSnd
        FinalSnd = HEX(Cmd)
        SndName = SndName & FinalSnd
        If SndName = "FD5" Then StopSound "FD4"

        PlaySound SndName
    End If
End Sub

Dim LastMus
LastMus = " "

Sub MusicCommand(Cmd)
    SevTimer.Enabled = 0
    If Len(LastMus) > 0 Or Cmd = 0 Then StopSound LastMus

    Dim FinalMus
    FinalMus = Hex(Cmd)
    LastMus = "FD" & FinalMus
    If Cmd = 7 Then
        LastMus = "FD07A"
        SevTimer.Enabled = 1
        PlaySound "FD07A"
        Exit Sub
    End If

    PlaySound LastMus, -1
End Sub

Sub SevTimer_Timer
    LastMus = "FD07B"
    PlaySound LastMus, -1
    SevTimer.Enabled = 0
End Sub

''******************************
''****END DONOR TABLE SCRIPT****
''******************************

Sub UpdateFlipperLogo_Timer
    LFLogo.RotY = LeftFlipper.CurrentAngle
    RFlogo.RotY = RightFlipper.CurrentAngle
End Sub

' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Sub Pins_Hit (idx)
	PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
End Sub

Sub Targets_Hit (idx)
	PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
End Sub

Sub TargetBankWalls_Hit (idx)
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
		Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End Select
End Sub


'Sub LRRail_Hit:PlaySound "fx_metalrolling", 0, 150, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0:End Sub
'
'Sub RLRail_Hit:PlaySound "fx_metalrolling", 0, 150, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0:End Sub

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 1)
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
        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 AND BOT(b).z > 0 Then
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0
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
	PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0
End Sub

Dim LeftCount:LeftCount = 0
Sub leftdrop_hit
	If LeftCount = 1 then
		playsound "BallDrop"
	End If
	LeftCount = 0
End Sub

Dim RightCount:RightCount = 0
Sub rightdrop_hit
	If RightCount = 1 then
		playsound "BallDrop"
	End If
	RightCount = 0
End Sub


Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="vector",UseSolenoids=0,UseLamps=0,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff", SCoin="coin"

LoadVPM "01130100", "Bally.VBS", 3.21
Dim DesktopMode: DesktopMode = Table1.ShowDT

If DesktopMode = True Then 'Show Desktop components
Ramp16.visible=1
Ramp15.visible=1
Primitive13.visible=1
Else
Ramp16.visible=0
Ramp15.visible=0
Primitive13.visible=1
End if

'*************************************************************
'Solenoid Call backs
'**********************************************************************************************************

SolCallback(1)      		= "dtU.SolDropUp"					        	' Top Bank       Target Reset    (1)
SolCallback(2)      		= "dtL.SolDropUp"					        	' Bot Bank Low   Target Reset    (7)
SolCallback(4)      		= "dtLL.SolDropUp"					        	' Top Bank Upper Target Reset    (6)
SolCallback(3)				= "DoSaucerL"								' Saucer (Lower Left)       Release
SolCallback(5)				= "DoSaucerR"								' Saucer (Lower Right)      Release
SolCallback(6)		        = "vpmSolSound SoundFX(""Knocker"",DOFKnocker)," 		        ' Knocker     (4)
SolCallback(7)      		= "bsTrough.SolOut"						    ' Outhole Kicker (5)  (Ball Release)
SolCallback(9)				= "DoUpperSaucer3Up"						' Saucer (2nd Floor Top)    Release
SolCallback(10)				= "DoUpperSaucer3Down"						' Saucer (2nd Floor Top)    Release
SolCallback(11)				= "DoUpperSaucer2"							' Saucer (2nd Floor Mid)    Release
SolCallback(12)				= "DoUpperSaucer1"							' Saucer (2nd Floor Bottom) Release
SolCallback(18)     		= ""								        ' Coin Box (lockout) ()
SolCallback(19)     		= "RelayAC"									' K1 Relay (Flipper Enable)()

' These are reassigned in the Solenoid Timer Handler via Lamp 63 as a Selection Switch in VPM
SolCallback(31)				= "Drop1"									' Lower Top Drop Target Down 1
SolCallback(32)				= "Drop2"									' Lower Top Drop Target Down 2
SolCallback(33)				= "Drop3"									' Lower Top Drop Target Down 3
SolCallback(34)				= "Drop4"									' Lower Bot Drop Target Down 1
SolCallback(35)				= "Drop5"									' Lower Bot Drop Target Down 2
SolCallback(36)				= "Drop6"									' Lower Bot Drop Target Down 3

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySoundAt SoundFX("Flipperup",DOFContactors), LeftFlipper:LeftFlipper.RotateToEnd:LeftFlipper1.RotateToEnd
     Else
         PlaySoundAt SoundFX("Flipperdown",DOFContactors), LeftFlipper:LeftFlipper.RotateToStart:LeftFlipper1.RotateToStart
     End If
  End Sub
  
Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySoundAt SoundFX("Flipperup",DOFContactors), RightFlipper:RightFlipper.RotateToEnd:RightFlipper1.RotateToEnd
     Else
         PlaySoundAt SoundFX("Flipperdown",DOFContactors), RightFlipper:RightFlipper.RotateToStart:RightFlipper1.RotateToStart
     End If
End Sub

'**********************************************************************************************************

 Sub FlipperTimer_Timer
	FlipperT1.roty = LeftFlipper1.currentangle  + 0
	FlipperT2.roty = RightFlipper1.currentangle + 0
	FlipperLSh.RotZ = LeftFlipper.currentangle
	FlipperRSh.RotZ = RightFlipper.currentangle
	FlipperLSh1.RotZ = LeftFlipper1.currentangle
	FlipperRSh1.RotZ = RightFlipper1.currentangle
	Prim2.Rotz = Gate2.Currentangle 
	Prim3.Rotz = Gate3.Currentangle 
	Prim4.Rotz = Gate4.Currentangle 
	Prim5.Rotz = Gate5.Currentangle
	Prim6.Rotz = Gate6.Currentangle 
	Prim14.Rotz = sw14.Currentangle  
End Sub

 Sub Drop1(enabled) : If enabled then : sw31_Dropped : DOF dDropUpperLeft, 2 : end if : End Sub
 Sub Drop2(enabled) : If enabled then : sw30_Dropped : DOF dDropUpperMiddle, 2 : end if : End Sub
 Sub Drop3(enabled) : If enabled then : sw29_Dropped : DOF dDropUpperRight, 2 : end if : End Sub
 Sub Drop4(enabled) : If enabled then : sw28_Dropped : DOF dDropLowerLeft, 2 : end if : End Sub
 Sub Drop5(enabled) : If enabled then : sw27_Dropped : DOF dDropLowerMiddle, 2 : end if : End Sub
 Sub Drop6(enabled) : If enabled then : sw26_Dropped : DOF dDropLowerRight, 2 : end if : End Sub

 Sub Solenoid_Timer()
	Dim Changed, Count, funcName, ii, sel, solNo
	Changed = Controller.ChangedSolenoids
	If Not IsEmpty(Changed) Then
		sel = Controller.Lamp(63)
		Count = UBound(Changed, 1)
		For ii = 0 To Count
			solNo = Changed(ii, CHGNO)
			If SolNo >= 7 And SolNo <= 12 And sel Then solNo = solNo +24
			funcName = SolCallback(solNo)
			If funcName <> "" Then Execute funcName & " CBool(" & Changed(ii, CHGSTATE) &")"
		Next
	End If
End Sub

 'Handle Solenoid Events
Sub SolTimer_Timer()
	Dim ChgSol, tmp, ii, CBoard, solnum
	ChgSol  = Controller.ChangedSolenoids
	If Not IsEmpty(ChgSol) Then
 	CBoard = Controller.Lamp(63)
		For ii = 0 To UBound(ChgSol)
 			solnum = ChgSol(ii, 0)
 			If solnum <= 12 and CBoard Then solnum = solnum + 24
 			tmp = Solcallback(solnum)
			If tmp <> "" Then Execute tmp & vpmTrueFalse(ChgSol(ii, 1)+1)
		Next
	End If
End Sub

' Tie In Nudge to AC Relay
Sub RelayAC(enabled)
	vpmNudge.SolGameOn enabled  		
End Sub

'**********************************************************************************************************
'Initiate Table
'**********************************************************************************************************

Dim bsTrough, bsSaucerL, SLMagnet, bsSaucerR, SRMagnet, bsSaucer1, bsSaucer2, bsSaucer3, DTL, DTLL, DTU

Sub Table1_Init
	vpmInit Me
	On Error Resume Next
		With Controller
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
		.SplashInfoLine = "Vector (Bally)"&chr(13)&"You Suck"
		.HandleMechanics=0
		.HandleKeyboard=0
		.ShowDMDOnly=1
		.ShowFrame=0
		.ShowTitle=0
        .hidden = 1
		If Err Then MsgBox Err.Description
	End With
	On Error Goto 0
		Controller.SolMask(0)=0
      vpmTimer.AddTimer 2000,"Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the timer to renable all the solenoids after 2 seconds
		Controller.Run
	If Err Then MsgBox Err.Description
	On Error Goto 0

	PinMAMETimer.Interval=PinMAMEInterval
	PinMAMETimer.Enabled=1

	vpmNudge.TiltSwitch  = swTiltMe
	vpmNudge.Sensitivity = 5
	vpmNudge.Tiltobj = Array(LeftSlingshot,RightSlingshot,Bumper1) ' Slings and Pop Bumpers

 	' Trough
	Set bsTrough = New cvpmBallStack
		bsTrough.Initsw 0,3,2,1,0,0,0,0
 		bsTrough.InitKick BallRelease,90,25
 		bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
		bsTrough.Balls = 3
 		bsTrough.IsTrough = True

  	' Saucer (Bottom Left)
	Set bsSaucerL = New cvpmBallStack
		bsSaucerL.InitSaucer SaucerL,swLeftBottomSaucer,78,16

 		bsSaucerL.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)

     ' Saucer Magnet Left
    Set SLMagnet = New cvpmMagnet
        SLMagnet.InitMagnet SaucerLeftMagnet, 9
        SLMagnet.GrabCenter = 0
        SLMagnet.MagnetOn = 1
        SLMagnet.CreateEvents "SLMagnet"

  	' Saucer (Bottom Right)
	Set bsSaucerR = New cvpmBallStack
		bsSaucerR.InitSaucer SaucerR,swRightBottomSaucer,-87,16
 		bsSaucerR.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)

     ' Saucer Magnet Right
    Set SRMagnet = New cvpmMagnet
        SRMagnet.InitMagnet SaucerRightMagnet, 9
        SRMagnet.GrabCenter = 0
        SRMagnet.MagnetOn = 1
        SRMagnet.CreateEvents "SRMagnet"

  	' Saucer (2nd Floor Bottom)
	Set bsSaucer1 = New cvpmBallStack
		bsSaucer1.InitSaucer Saucer1,swUpperSaucer1,178,1.30
 		bsSaucer1.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)

  	' Saucer (2nd Floor Middle)
	Set bsSaucer2 = New cvpmBallStack
		bsSaucer2.InitSaucer Saucer2,swUpperSaucer2,180,4
 		bsSaucer2.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)

  	' Saucer (2nd Floor Top)
	Set bsSaucer3 = New cvpmBallStack
		bsSaucer3.InitSaucer Saucer3,swUpperSaucer3,340,18
 		bsSaucer3.InitAltKick 160,10
 		bsSaucer3.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)

  	Set dtL=New cvpmDropTarget
		dtL.InitDrop Array(sw26,sw27,sw28),Array(swLowerDropTargetRight,swLowerDropTargetMiddle,swLowerDropTargetLeft)
		dtL.InitSnd SoundFX("DTDrop1",DOFContactors),SoundFX("DTReset",DOFContactors)

 	Set dtLL=New cvpmDropTarget
		dtLL.InitDrop Array(sw29,sw30,sw31),Array(swUpperDropTargetRight,swUpperDropTargetMiddle,swUpperDropTargetLeft)
		dtLL.InitSnd SoundFX("DTDropFAKE",DOFContactors),SoundFX("DTReset1",DOFContactors)

 	Set dtU=New cvpmDropTarget
		dtU.InitDrop Array(sw46,sw47,sw48),Array(swTopDropTargetZ,swTopDropTargetY,swTopDropTargetX)
		dtU.InitSnd SoundFX("DTDropFAKE",DOFContactors),SoundFX("DTReset2",DOFContactors)

 End Sub
 
'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)
	If KeyDownHandler(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Pullback:PlaySoundAt "plungerpull", Plunger
 	If keycode = RightMagnaSave Then : controller.switch(8) = True ' VectorScan Reset Switch (Page Up Key)
 	If keycode = 38 Then LampTest
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
	If KeyUpHandler(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Fire:PlaySoundAt "plunger", Plunger
	If keycode = RightMagnaSave Then : controller.switch(8) = False ' VectorScan Reset Switch (Page Up Key)
End Sub


'**********************************************************************************************************

 ' Drain and Kickers
SaucerRWU.IsDropped = 1
SaucerLWU.IsDropped = 1
UpperSaucer1WU.IsDropped = 1
UpperSaucer2WU.IsDropped = 1
UpperSaucer3DWU.IsDropped = 1
UpperSaucer3UWU.IsDropped = 1

Sub Drain_Hit:PlaySoundAt "drainLong", Drain:bsTrough.addball me:End Sub

Sub Saucer1_Hit:bsSaucer1.addball me :Stopsound "fx_BallRolling": PlaySoundAtVol "KickerEnter", Saucer1, 0.5 : End Sub  
Sub Saucer1_UnHit: PlaySoundAt "PopperBall", Saucer1: End Sub

Sub Saucer2_Hit:bsSaucer2.addball me :Stopsound "fx_BallRolling": PlaySoundAtVol "KickerEnter", Saucer2, 0.5 : End Sub 
Sub Saucer2_UnHit: PlaySoundAt "PopperBall", Saucer2: End Sub

Sub Saucer3_Hit:bsSaucer3.addball me :Stopsound "fx_BallRolling": PlaySoundAtVol "KickerEnter", Saucer3, 0.5 : End Sub 
Sub Saucer3_UnHit: PlaySoundAt "PopperBall", Saucer3: End Sub

Sub SaucerL_Hit:bsSaucerL.addball me :Stopsound "fx_BallRolling": PlaySoundAtVol "KickerEnter", SaucerL, 0.5: End Sub 
Sub SaucerL_UnHit: PlaySoundAt "PopperBall", SaucerL: End Sub

Sub SaucerR_Hit:bsSaucerR.addball me :Stopsound "fx_BallRolling": PlaySoundAtVol "KickerEnter", SaucerR, 0.5: End Sub 
Sub SaucerR_UnHit: PlaySoundAt "PopperBall", SaucerR: End Sub

Sub DoSaucerR(enabled)
	If enabled Then 
		bsSaucerR.ExitSol_On
		SaucerRWU.IsDropped = 0
		SaucerRWU.TimerEnabled = 1
 	End if
End Sub
Sub SaucerRWU_Timer:SaucerRWU.IsDropped = 1:Me.TimerEnabled = 0:End Sub

Sub DoSaucerL(enabled)
	If enabled Then 
		bsSaucerL.ExitSol_On
		SaucerLWU.IsDropped = 0
		SaucerLWU.TimerEnabled = 1
 	End if
End Sub
Sub SaucerLWU_Timer:SaucerLWU.IsDropped = 1:Me.TimerEnabled = 0:End Sub

Sub DoUpperSaucer1(enabled)
	If enabled Then 
		bsSaucer1.ExitSol_On
		UpperSaucer1WU.IsDropped = 0
		UpperSaucer1WU.TimerEnabled = 1
		DOF dUpperSaucer1, 2
 	End if
End Sub
Sub UpperSaucer1WU_Timer:UpperSaucer1WU.IsDropped = 1:Me.TimerEnabled = 0:End Sub

Sub DoUpperSaucer2(enabled)
	If enabled Then 
		bsSaucer2.ExitSol_On
		UpperSaucer2WU.IsDropped = 0
		UpperSaucer2WU.TimerEnabled = 1
		DOF dUpperSaucer2, 2
 	End if
End Sub
Sub UpperSaucer2WU_Timer:UpperSaucer2WU.IsDropped = 1:Me.TimerEnabled = 0:End Sub

Sub DoUpperSaucer3Down(enabled)
	If enabled Then 
		bsSaucer3.ExitAltSol_On
		UpperSaucer3DWU.IsDropped = 0
		UpperSaucer3DWU.TimerEnabled = 1
		DOF dUpperSaucer3Down, 2
 	End if
End Sub
Sub UpperSaucer3DWU_Timer:UpperSaucer3DWU.IsDropped = 1:Me.TimerEnabled = 0:End Sub

Sub DoUpperSaucer3Up(enabled)
	If enabled Then 
		bsSaucer3.ExitSol_On
		UpperSaucer3UWU.IsDropped = 0
		UpperSaucer3UWU.TimerEnabled = 1
		DOF dUpperSaucer3Up, 2
 	End if
End Sub
Sub UpperSaucer3UWU_Timer:UpperSaucer3UWU.IsDropped = 1:Me.TimerEnabled = 0:End Sub

'Gate Trigger
Sub sw14_Hit()   	  : vpmTimer.PulseSw(swVectorScanGate) :PlaySound "GateWire", 0, Vol(ActiveBall), AudioPan(sw14), 0, Pitch(ActiveBall), 1, 0, AudioFade(sw14): End Sub  

'Wire Triggers
Sub sw12_Hit()   	  : controller.switch (swShooterLane)=1 : PlaySoundAtBall "Sensor" : End Sub  
Sub sw12_unHit() 	  : controller.switch (swShooterLane)=0:End Sub
Sub sw13_Hit()		  : controller.switch (swTopRightLane)=1 : PlaySoundAtBall "Sensor" : End Sub 
Sub sw13_unHit()	  : controller.switch (swTopRightLane)=0:End Sub
Sub sw33_Hit()   	  : controller.switch (swRightOutLane)=1 : PlaySoundAtBall "Sensor" : End Sub
Sub sw33_unHit() 	  : controller.switch (swRightOutLane)=0:End Sub
Sub sw34_Hit()   	  : controller.switch (swRightReturnLane)=1 : PlaySoundAtBall "Sensor" : End Sub 
Sub sw34_unHit() 	  : controller.switch (swRightReturnLane)=0:End Sub
Sub sw35_Hit()   	  : controller.switch (swLeftOutLane)=1 : PlaySoundAtBall "Sensor" : End Sub 
Sub sw35_unHit() 	  : controller.switch (swLeftOutLane)=0:End Sub
Sub sw36_Hit()   	  : controller.switch (swLeftReturnLane)=1 : PlaySoundAtBall "Sensor" : End Sub 
Sub sw36_unHit() 	  : controller.switch (swLeftReturnLane)=0:End Sub

'Star Triggers
Sub sw37_Hit()   	  : controller.switch (swAdvanceBonusRollOver)=1 : PlaySoundAtBall "Sensor" : End Sub 
Sub sw37_unHit() 	  : controller.switch (swAdvanceBonusRollOver)=0:End Sub

'Bumpers
Sub Bumper1_Hit : vpmTimer.PulseSw(swBumper) : PlaySoundAt SoundFX("bumper1",DOFContactors), Bumper1: End Sub

 'Stand Up Targets
 Sub sw21_Hit : vpmTimer.PulseSw(swTargetE) : End Sub 
 Sub sw22_Hit : vpmTimer.PulseSw(swTargetP): End Sub 
 Sub sw23_Hit : vpmTimer.PulseSw(swTargetY) : End Sub 
 Sub sw24_Hit : vpmTimer.PulseSw(swTargetH) : End Sub 

'Drop Targets
 Sub Sw26_Dropped:dtL.Hit 1 :End Sub  
 Sub Sw27_Dropped:dtL.Hit 2 :End Sub  
 Sub Sw28_Dropped:dtL.Hit 3 :End Sub

'Drop Targets
 Sub Sw29_Dropped:dtLL.Hit 1 :End Sub  
 Sub Sw30_Dropped:dtLL.Hit 2 :End Sub  
 Sub Sw31_Dropped:dtLL.Hit 3 :End Sub

'Drop Targets
 Sub Sw46_Dropped:dtU.Hit 1 :End Sub  
 Sub Sw47_Dropped:dtU.Hit 2 :End Sub  
 Sub Sw48_Dropped:dtU.Hit 3 :End Sub


'***************************************************
'       JP's VP10 Fading Lamps & Flashers
'       Based on PD's Fading Light System
' SetLamp 0 is Off
' SetLamp 1 is On
' fading for non opacity objects is 4 steps
'***************************************************

Dim LampTst, LampState(200), FadingLevel(200)
Dim FlashSpeedUp(200), FlashSpeedDown(200), FlashMin(200), FlashMax(200), FlashLevel(200)

InitLamps()             ' turn off the lights and flashers and reset them to the default parameters
LampTimer.Interval = 5 'lamp fading speed
LampTimer.Enabled = 1
LampTst = 0

' Lamp & Flasher Timers

Sub LampTest
	Dim x
	If LampTst = 1 Then
		LampTst = 0
		LampTimer.Enabled = 1
	Else
		LampTimer.Enabled = 0
		For x = 0 to 200
			LampState(x) = 0       ' current light state, independent of the fading level. 0 is off and 1 is on
			FadingLevel(x) = 5      ' used to track the fading state
			FlashLevel(x) = 0       ' the intensity of the flashers, usually from 0 to 1
		Next
		LampTst = 1
	End If
End Sub

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
NFadeL 2, L2
NFadeL 3, L3
NFadeL 4, L4
NFadeL 6, L6
NFadeL 7, L7
NFadeL 8, L8
NFadeL 9, L9
NFadeL 10, L10
NFadeLm 12, L12a				' Pop Bumper Plastic
NFadeLm 12, L12b
NFadeLm 12, L12c
NFadeL 14, L14
NFadeL 15, L15
NFadeL 18, L18
NFadeL 19, L19
NFadeL 20, L20
NFadeL 21, L21
NFadeL 22, L22
NFadeL 23, L23
NFadeL 24, L24
NFadeL 25, L25
NFadeL 26, L26
NFadeL 28, L28
NFadeL 30, L30
NFadeL 31, L31
NFadeLm 33, L33a1
NFadeLm 33, L33b1
NFadeLm 33, L33c1
NFadeLm 33, L33a2
NFadeLm 33, L33b2
NFadeLm 33, L33c2
NFadeLm 33, L33a3
NFadeLm 33, L33b3
NFadeLm 33, L33c3
NFadeLm 33, L33a4
NFadeLm 33, L33b4
NFadeLm 33, L33c4
NFadeLm 33, L33a5
NFadeLm 33, L33b5
NFadeLm 33, L33c5
NFadeLm 33, L33a6
NFadeLm 33, L33b6
NFadeLm 33, L33a7
NFadeLm 33, L33c7
NFadeL 34, L34
NFadeL 35, L35
NFadeL 36, L36
NFadeL 37, L37
NFadeL 38, L38
NFadeL 39, L39
NFadeL 40, L40
NFadeL 41, L41
NFadeL 42, L42
NFadeL 43, L43
NFadeL 44, L44
NFadeL 46, L46
NFadeL 47, L47
NFadeLm 49, L49rot
NFadeLm 49, L49a1   'UPPER PF GI???????????????
NFadeLm 49, L49b1
NFadeLm 49, L49c1
NFadeLm 49, L49a2   
NFadeLm 49, L49b2
NFadeLm 49, L49c2
Flash 49, L49f2
NFadeLm 49, L49a3   
NFadeLm 49, L49b3
NFadeLm 49, L49c3
Flash 49, L49f3
NFadeLm 49, L49a4   
NFadeLm 49, L49b4
NFadeLm 49, L49c4
NFadeLm 49, L49a5
NFadeLm 49, L49c5
'NFadeObjm 49, l49g, "Plastics2B_off", "Plastics2B_off"
NFadeLm 49, L49a6
NFadeLm 49, L49c6
NFadeL 50, L50
NFadeL 51, L51
NFadeL 52, L52
NFadeL 53, L53
NFadeL 54, L54
NFadeL 55, L55
NFadeL 56, L56
NFadeL 57, L57
NFadeL 58, L58
NFadeL 59, L59
NFadeL 60, L60
NFadeL 62, L62
NFadeLm 65, l65
NFadeLm 65, l65a
NFadeLm 65, l65b
Flash 65, F65
NFadeLm 66, L66
NFadeLm 66, l66a
Flash 66, L66f
NFadeLm 67, L67
NFadeLm 67, l67a
NFadeLm 67, l67b
Flash 67, F67
NFadeLm 81, l81
NFadeLm 81, l81a
NFadeLm 81, l81b
Flash 81, F81
NFadeLm 82, L82
NFadeLm 82, l82b
Flash 82, L82f
NFadeL 83, L83				' Plunger Flasher
NFadeLm 97, l97
NFadeLm 97, l97a
NFadeLm 97, l97b
Flash 97, F97
NFadeLm 98, L98
NFadeLm 98, l98a
NFadeLm 98, l98b
Flash 98, L98f
Flashm 98, L98f1
NFadeL 99, L99
NFadeLm 113, l113
NFadeLm 113, l113a
NFadeLm 113, l113b
Flash 113, F113
NFadeLm 114, l114
NFadeLm 114, l114a
NFadeLm 114, l114b
Flash 114, l114f
Flashm 114, l114f1
NFadeL 115, L115
If FadingLevel(49) > 3 Then FadingLevel(149) = FadingLevel(49) 
FlashOpacity 149, Flasher1, 0.9, 0.9

'Backglass Lights
'FadeReel 11, L11					' Shoot Again
'FadeReel 13, L13					' Ball In Play
'FadeReel 27, L27					' Match
'FadeReel 29, L29					' High Score
'FadeReel 45, L45					' Game Over
'FadeReel 61, L61					' Tilt

 'Infinity Lights
'NFadeL 68, L68				'1 top    & right
'NFadeL 70, L70			    '1 left   & bottom
'NFadeL 84, L84				'2 top    & right
'NFadeL 86, L86				'2 left   & bottom
'NFadeL 100, L100			'3 top    & right
'NFadeL 102, L102			'3 left   & bottom
'NFadeL 116, L116			'4 top    & right
'NFadeL 118, L118			'4 left   & bottom
'NFadeL 69, L69			    '5 bottom & right
'NFadeL 71, L71				'5 top    & left
'NFadeL 85, L85				'6 bottom & right
'NFadeL 87, L87				'6 top    & left
'NFadeL 101, L101			'7 bottom & right
'NFadeL 103, L103			'7 top    & left
'NFadeL 117, L117			'8 bottom & right
'NFadeL 119, L119			'8 top    & left

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
	FlashMax(149) = 40
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

Sub FlashOpacity(nr, object, stepup, stepdown) 
    Select Case FadingLevel(nr)
        Case 4 'off
            FlashLevel(nr) = FlashLevel(nr) - stepdown
            If FlashLevel(nr) < 0 Then
                FlashLevel(nr) = 0
                FadingLevel(nr) = 0 'completely off
            End if
            Object.Opacity = FlashLevel(nr)
        Case 5 ' on
            FlashLevel(nr) = FlashLevel(nr) + stepup
            If FlashLevel(nr) > FlashMax(nr) Then
                FlashLevel(nr) = FlashMax(nr)
                FadingLevel(nr) = 1 'completely on
            End if
            Object.Opacity = FlashLevel(nr)
    End Select
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
'Digital Display
'**********************************************************************************************************

Dim Digits(40)
Digits(0)=Array(b00,b01,b02,b03,b04,b05,b06)
Digits(1)=Array(b10,b11,b12,b13,b14,b15,b16)
Digits(2)=Array(b20,b21,b22,b23,b24,b25,b26)
Digits(3)=Array(b30,b31,b32,b33,b34,b35,b36)
Digits(4)=Array(b40,b41,b42,b43,b44,b45,b46)
Digits(5)=Array(b50,b51,b52,b53,b54,b55,b56)
Digits(6)=Array(b60,b61,b62,b63,b64,b65,b66)
Digits(7)=Array(b70,b71,b72,b73,b74,b75,b76)
Digits(8)=Array(b80,b81,b82,b83,b84,b85,b86)
Digits(9)=Array(b90,b91,b92,b93,b94,b95,b96)
Digits(10)=Array(ba0,ba1,ba2,ba3,ba4,ba5,ba6)
Digits(11)=Array(bb0,bb1,bb2,bb3,bb4,bb5,bb6)
Digits(12)=Array(bc0,bc1,bc2,bc3,bc4,bc5,bc6)
Digits(13)=Array(bd0,bd1,bd2,bd3,bd4,bd5,bd6)
Digits(14)=Array(be0,be1,be2,be3,be4,be5,be6)
Digits(15)=Array(bf0,bf1,bf2,bf3,bf4,bf5,bf6)
Digits(16)=Array(bg0,bg1,bg2,bg3,bg4,bg5,bg6)
Digits(17)=Array(bh0,bh1,bh2,bh3,bh4,bh5,bh6)
Digits(18)=Array(bi0,bi1,bi2,bi3,bi4,bi5,bi6)
Digits(19)=Array(bj0,bj1,bj2,bj3,bj4,bj5,bj6)
Digits(20)=Array(bk0,bk1,bk2,bk3,bk4,bk5,bk6)
Digits(21)=Array(bl0,bl1,bl2,bl3,bl4,bl5,bl6)
Digits(22)=Array(bm0,bm1,bm2,bm3,bm4,bm5,bm6)
Digits(23)=Array(bn0,bn1,bn2,bn3,bn4,bn5,bn6)
Digits(24)=Array(bo0,bo1,bo2,bo3,bo4,bo5,bo6)
Digits(25)=Array(bp0,bp1,bp2,bp3,bp4,bp5,bp6)
Digits(26)=Array(bq0,bq1,bq2,bq3,bq4,bq5,bq6)
Digits(27)=Array(br0,br1,br2,br3,br4,br5,br6)
Digits(28)=Array(bs0,bs1,bs2,bs3,bs4,bs5,bs6)
Digits(29)=Array(bt0,bt1,bt2,bt3,bt4,bt5,bt6)
Digits(30)=Array(bu0,bu1,bu2,bu3,bu4,bu5,bu6)
Digits(31)=Array(bv0,bv1,bv2,bv3,bv4,bv5,bv6)
Digits(32)=Array(a00,a01,a02,a03,a04,a05,a06)
Digits(33)=Array(a10,a11,a12,a13,a14,a15,a16)
Digits(34)=Array(a20,a21,a22,a23,a24,a25,a26)
Digits(35)=Array(a30,a31,a32,a33,a34,a35,a36)
Digits(36)=Array(a40,a41,a42,a43,a44,a45,a46)
Digits(37)=Array(a50,a51,a52,a53,a54,a55,a56)

Sub DisplayTimer_Timer
	Dim ChgLED,ii,num,chg,stat,obj
	ChgLed = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
If Not IsEmpty(ChgLED) Then
		For ii = 0 To UBound(chgLED)
			num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
			if (num < 38) then
				For Each obj In Digits(num)
					If chg And 1 Then obj.State = stat And 1 
					chg = chg\2 : stat = stat\2
				Next
			else
			end if
		next
end if
End Sub

'if Full Screen turn off Backglas Components but keep PF LED displays on
If DesktopMode = True Then
dim xxx
For each xxx in BG:xxx.Visible = 1: Next
else
For each xxx in BG:xxx.Visible = 0: Next
End if


'Bally Vector
'added by Inkochnito
Sub editDips
	Dim vpmDips : Set vpmDips = New cvpmDips
	With vpmDips
		.AddForm 700,400,"Vector - DIP switches"
		.AddChk 7,10,180,Array("Match feature", &H08000000)'dip 28
		.AddChk 205,10,115,Array("Credits display", &H04000000)'dip 27
		.AddFrame 2,30,190,"Maximum credits",&H03000000,Array("10 credits",0,"15 credits", &H01000000,"25 credits", &H02000000,"40 credits", &H03000000)'dip 25&26
		.AddFrame 2,106,190,"Making 3 upper left targets will",&H00000020,Array("reset 3 lower targets",0,"not reset 3 lower targets",&H00000020)'dip 6
		.AddFrame 2,152,190,"Multipiers lite will",&H00000040,Array("be off then alternate",0,"stay on till bonus lites are made",&H00000040)'dip 7
		.AddFrame 2,198,190,"X-Y-Z drop targets special",&H00000080,Array("25K - special - 25K",0,"25K - special keeps alternating",&H00000080)'dip 8
		.AddFrame 2,248,190,"Vectorscan to date readout",&H00002000,Array("can only be decreased manually",0,"after 8 games decrease by 20K",&H00002000)'dip 14
		.AddFrame 2,298,190,"With Vectorscan capture lite on",&H00004000,Array("3 lower targets will reset",0,"3 lower targets will go back down",&H00004000)'dip 15
		.AddFrame 2,348,190,"With Vectorscan capture lite off",32768,Array("targets down will reset",0,"any target down will go back down",32768)'dip 16
		.AddFrame 205,30,190,"Balls per game",&HC0000000,Array ("2 balls",&HC0000000,"3 balls",0,"4 balls",&H80000000,"5 balls",&H40000000)'dip 31&32
		.AddFrame 205,106,190,"Vectorscan bonus score readout",&H00100000,Array("readout will reset to 1",0,"readout will come back on",&H00100000)'dip 21
		.AddFrame 205,152,190,"Attract sound",&H00200000,Array("no voice",0,"voice says 'I am P.A.C. play analysis",&H00200000)'dip 22
		.AddFrame 205,198,190,"Vectorscan scoring adjust",&H00400000,Array("scores only when capture lite is on",0,"scores when capture lite is on or off",&H00400000)'dip 23
		.AddFrame 205,248,190,"Lite special and extra ball when",&H00800000,Array("hitting vector speed 800 or over",0,"hitting vector speed 750 or over",&H00800000)'dip 24
		.AddFrame 205,298,190,"Replay limit",&H10000000,Array("no limit",0,"1 replay per game",&H10000000)'dip 29
		.AddFrame 205,348,190,"H-Y-P-E target bonus adjust",&H20000000,Array("advances bonus only when lit",0,"advances bonus every time",&H20000000)'dip 30
		.AddLabel 50,400,350,20,"Set selftest position 16,17,18 and 19 to 03 for the best gameplay."
		.AddLabel 50,420,300,20,"After hitting OK, press F3 to reset game with new settings."
		.ViewDips
	End With
End Sub
Set vpmShowDips = GetRef("editDips")

' *********************************************************************
' *********************************************************************

					'Start of VPX call back Functions

' *********************************************************************
' *********************************************************************

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
	vpmtimer.pulsesw swSlingRight
    PlaySoundAt SoundFX("SlingshotRight",DOFContactors), sling1
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0:
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
	vpmtimer.pulsesw 39
    PlaySoundAt SoundFX("SlingshotLeft",DOFContactors), sling2
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0:
    End Select
    LStep = LStep + 1
End Sub






'*****************************************
'	ninuzzu's	BALL SHADOW
'*****************************************
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
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/40) + ((BOT(b).X - (Table1.Width/2))/40)) + 1
		Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/40) + ((BOT(b).X - (Table1.Width/2))/40)) - 1
        End If
        ballShadow(b).Y = BOT(b).Y + 30
		BallShadow(b).Z = BOT(b).Z + 1
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
End Sub





'***********************************************************************************
'****				        	Switch reference      				  		  	****
'***********************************************************************************
Const swBallStack1					= 1
Const swBallStack2					= 2
Const swOuthole						= 3
Const swRightBottomSaucer			= 4
Const swLeftBottomSaucer			= 5
Const swShooterLane					= 12
Const swTopRightLane				= 13
Const swVectorScanGate				= 14
Const swTiltMe						= 15
Const swUpperSaucer1				= 17
Const swUpperSaucer2				= 18
Const swUpperSaucer3				= 19
Const swTargetE						= 21
Const swTargetP						= 22
Const swTargetY						= 23
Const swTargetH						= 24
Const swLowerDropTargetRight		= 26
Const swLowerDropTargetMiddle		= 27
Const swLowerDropTargetLeft			= 28
Const swUpperDropTargetRight		= 29
Const swUpperDropTargetMiddle		= 30
Const swUpperDropTargetLeft			= 31
Const swRightOutLane				= 33
Const swRightReturnLane				= 34
Const swLeftOutLane					= 35
Const swLeftReturnLane				= 36
Const swAdvanceBonusRollOver		= 37
Const swSlingRight					= 38
Const swSlingLeft					= 39
Const swBumper						= 40
Const swTopDropTargetZ				= 46
Const swTopDropTargetY				= 47
Const swTopDropTargetX				= 48

'***********************************************************************************
'****				        		DOF reference      				    		****
'***********************************************************************************
Const dShooterLane					= 200
Const dBallRelease					= 201
Const dUpperSaucer3Up				= 202
Const dUpperSaucer3Down				= 203
Const dUpperSaucer2					= 204
Const dUpperSaucer1					= 205
Const dDropUpperLeft				= 206
Const dDropUpperMiddle				= 207
Const dDropUpperRight				= 208
Const dDropLowerLeft				= 209
Const dDropLowerMiddle				= 210
Const dDropLowerRight				= 211









'*********************************************************************
'                 Positional Sound Playback Functions
'*********************************************************************

' Play a sound, depending on the X,Y position of the table element (especially cool for surround speaker setups, otherwise stereo panning only)
' parameters (defaults): loopcount (1), volume (1), randompitch (0), pitch (0), useexisting (0), restart (1))
' Note that this will not work (currently) for walls/slingshots as these do not feature a simple, single X,Y position

Sub PlayXYSound(soundname, tableobj, loopcount, volume, randompitch, pitch, useexisting, restart)
	PlaySound soundname, loopcount, volume, AudioPan(tableobj), randompitch, pitch, useexisting, restart, AudioFade(tableobj)
End Sub

' Similar subroutines that are less complicated to use (e.g. simply use standard parameters for the PlaySound call)
Sub PlaySoundAt(soundname, tableobj)
    PlaySound soundname, 1, 1, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname)
    PlaySoundAt soundname, ActiveBall
End Sub

'Set position as table object and Vol manually.

Sub PlaySoundAtVol(sound, tableobj, Vol)
		PlaySound sound, 1, Vol, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
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
        If BallVel(BOT(b) ) > 1 and not BOT(b).z Then
			rolling(b) = True
			if BOT(b).z < 30 Then ' Ball on playfield
						PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), audioPan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
			Else ' Ball on raised ramp
						PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*.5, audioPan(BOT(b) ), 0, Pitch(BOT(b) )+30000, 1, 0, AudioFade(BOT(b) )
				End If
		Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If
    Next

	on error resume next ' In case VP is too old..
		' Kill ball spin
	if mLockMagnet.MagnetOn then
		dim rampballs:rampballs = mLockMagnet.Balls
		dim obj
		for each obj in rampballs
			obj.AngMomZ= 0
			obj.AngVelZ= 0
			obj.AngMomY= 0
			obj.AngVelY= 0
			obj.AngMomX= 0
			obj.AngVelX= 0
			obj.velx = 0
			obj.vely = 0
			obj.velz = 0
		next 
	end if
	on error goto 0
End Sub

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
	PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
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

' The sound is played using the VOL, AUDIOPAN, AUDIOFADE and PITCH functions, so the volume and pitch of the sound
' will change according to the ball speed, and the AUDIOPAN & AUDIOFADE functions will change the stereo position
' according to the position of the ball on the table.


'**************************************
' Explanation of the collision routine
'**************************************

' The collision is built in VP.
' You only need to add a Sub OnBallBallCollision(ball1, ball2, velocity) and when two balls collide they 
' will call this routine. What you add in the sub is up to you. As an example is a simple Playsound with volume and paning
' depending of the speed of the collision.


Sub Pins_Hit (idx)
	PlaySound "fx_PinHit", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
	PlaySound "target", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Switches_Hit (idx)
	PlaySound "Sensor", 0, 1, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub DropTargets_Hit (idx)
	PlaySound "DTDrop", 0, 1, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
	PlaySound "metalhit_thin", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
	PlaySound "metalhit_medium", 0, Vol(ActiveBall)/2, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
	PlaySound "metalhit2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
	PlaySound "GateWire", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Wall_Hit(idx)
	RandomGlassHit
End Sub

Sub Spinner_Spin
	PlaySound "fx_spinner", 0, .25, AudioPan(Spinner), 0.25, 0, 0, 1, AudioFade(Spinner)
End Sub

Sub Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then 
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub Posts_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 16 then 
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 16 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub

Sub RandomSoundHole()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "Hole1", 0, 1, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "Hole2", 0, 1, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "Hole4", 0, 1, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub

Sub RandomGlassHit()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "GlassHit1", 0, 25, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "GlassHit2", 0, 25, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "GlassHit3", 0, 25, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub


Sub LeftFlipper_Collide(parm)
 	RandomSoundFlipper()
End Sub

Sub RightFlipper_Collide(parm)
 	RandomSoundFlipper()
End Sub

Sub LeftFlipper1_Collide(parm)
 	RandomSoundFlipper()
End Sub

Sub RightFlipper1_Collide(parm)
 	RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub



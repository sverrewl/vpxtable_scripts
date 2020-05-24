'************************************************************************************************
'*****************************		Airborne by Capcom (1996)	*********************************
'************************************************************************************************
'		Based on the VP8 table by TAB and Destruk, and the VP9 table by Bodydump
'		VPX conversion by robbo43 and DCrosby
'		Thanks go to:
'			the original table authors
'			DCrosby for the ramps, graphical updates and general support
'			those people who have contributed scripting techniques (JPSalas, ninuzzu, uncle willy)
'			those people that contribute high quality resources (Dark, nFozzy, Zany, Flupper)
'		My final thanks go the authors of VPX and VPinMame for making this all possible.
'
'		This has been quite a start/stop build and taken longer than it should, so apologies If
'		I have forgotten to credit anyone else. Please drop me a line if I've missed you off the
'		credits and I'll add you.
'************************************************************************************************

'**************************************************************************************************
'*********************************** Table Notes / Calls for Help *********************************
'	I'm not sure how far capcom emulation has come since the VP9 table was release, so I'll put
'	out a call for help
'
'	1.	I've tried many different methods to make the left drop ramp behave properly using rom
'		commands however the best behaviours are those workarounds developed by Destruk / Bodydump
'		I'm not sure whether these are 100% accurate, so please contribute any improvements
'
'	2.	I'm not sure whether locked balls should be kicked out if you select eject co-pilot. The
'		table currently does not do this and I do not know if this is correct behavour. I tried to
'		script this but couldn't get it to work properly.
'
'	3.	The table may still have many legacy commands / functions, please feel free to update and
'		contribute - I've reached the limit of my knowledge/skill!!
'
'**************************************************************************************************


Option Explicit
 Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Dim DesktopMode:DesktopMode = Table1.ShowDT
Dim UseVPMDMD:UseVPMDMD = DesktopMode
Const UseVPMModSol = 0

LoadVPM "01560000", "capcom.VBS", 3.26

Const UseSolenoids=2,UseLamps=0,UseSync=1,SSolenoidOn="SolOn",SSolenoidOff="SolOff",SFlipperOn="",SFlipperOff="",SCoin="Coin3"

Set LampCallback=GetRef("UpdateMultipleLamps")

SolCallback(1)="bsTrough.SolIn"
SolCallback(2)="bsTrough.SolOut"
SolCallback(3)="vpmSolSound ""Knocker"","
SolCallback(4)="" 'sling
SolCallback(5)="" 'sling
SolCallback(6)="bsBottomSaucer.SolOut"
SolCallback(7)="SolLockDoor"
SolCallback(8)="bsLeftSaucer.SolOut"
SolCallback(11)="dtSHOW"
SolCallback(12)="BallSaveGate"
SolCallback(13)="" 'Bumper
SolCallback(14)="" 'Bumper
SolCallback(15)="" 'Bumper
SolCallback(16)="dtAIR"
SolCallback(17)="SolLDiverter"
SolCallback(19)="EnableRamp"
SolCallback(29)="SolBackKicker"
SolCallback(30)="bsVLock.SolExit"
SolCallback(31)="SolRDiverter"
SolCallback(32)="bsRightSaucer.SolOut"
SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"
SolCallback(51)="GameOn"
Dim bump1, bump2, bump3
Dim rampstate

'DropTarget Solenoids
   Sub dtAIR(Enabled)
		If Enabled Then
			DropADir = 1:DropIDir = 1:DropRDir = 1
			DropAa.TimerEnabled = 1:DropIa.TimerEnabled = 1:DropRa.TimerEnabled = 1
			dtR.DropSol_On
		End if
   End Sub
   Sub dtSHOW(Enabled)
		If Enabled Then
			DropSDir = 1:DropHDir = 1:DropODir = 1:DropWDir = 1
			DropSa.TimerEnabled = 1:DropHa.TimerEnabled = 1:DropOa.TimerEnabled = 1:DropWa.TimerEnabled = 1
			dtL.DropSol_On
		End if
   End Sub

'***********	Flipper Primitives 	('0=yellow/black, 1=yellow/red)
Const FlippersType = 0
Select Case FlippersType
		Case 0
			CapcomLFlip.visible=1:CapcomRFlip.visible=1
			CapcomLFlip.image="star_flip_left":CapcomRFlip.image="star_flip_right"

		Case 1
			CapcomLFlip.visible=1:CapcomRFlip.visible=1
			CapcomLFlip.image="star_flip_leftRED":CapcomRFlip.image="star_flip_rightRED"
End Select

Sub FlipperTimer_Timer()
CapcomLFlip.RotY = LeftFlipper.CurrentAngle
CapcomRFlip.RotY = RightFlipper.CurrentAngle
FlipperLSh.RotZ = LeftFlipper.CurrentAngle
FlipperRSh.RotZ = RightFlipper.CurrentAngle
End Sub

'******Capcom Emulation Fixes (scripting by Destruk)
	'Right Diverter requires HOLD timer
		Sub SolRDiverter(Enabled)
			If Enabled Then
				RightDiverter.IsDropped=1
				RightDiverter2.IsDropped=0
				'RightDiverter.TimerEnabled=0
				'RightDiverter.TimerEnabled=1
			End If
		End Sub

		Sub RightDiverter_Timer
			RightDiverter.TimerEnabled=0
			If RightDiverter.IsDropped Then
				RightDiverter2.IsDropped=1
				RightDiverter.IsDropped=0
			End If
		End Sub

		Sub Trigger23_Hit
			If ActiveBall.VelY<0 Then
				If RightDiverter.IsDropped Then
					RightDiverter.TimerEnabled=0
					RightDiverter2.IsDropped=1
					RightDiverter.IsDropped=0
				End If
			End If
		End Sub

		Sub Trigger24_Hit
			If ActiveBall.VelX<0 Then
				If RightDiverter.IsDropped Then
					RightDiverter.TimerEnabled=0
					RightDiverter2.IsDropped=1
					RightDiverter.IsDropped=0
				End If
			End If
		End Sub

		Sub Trigger1_Hit
			Controller.Switch(69)=1
			PlaySoundAtVol "Sensor", ActiveBall, 1
			If ActiveBall.VelY>0 Then
				If RightDiverter.IsDropped Then
					RightDiverter.TimerEnabled=0
					RightDiverter2.IsDropped=1
					RightDiverter.IsDropped=0
				End If
			End If
		End Sub
		Sub Trigger1_unHit:Controller.Switch(69)=0:End Sub

		Sub Trigger2_Hit
			Controller.Switch(70)=1
			PlaySoundAtVol "Sensor", ActiveBall, 1
			If ActiveBall.VelY>0 Then
				If RightDiverter.IsDropped Then
					RightDiverter.TimerEnabled=0
					RightDiverter2.IsDropped=1
					RightDiverter.IsDropped=0
				End If
			End If
		End Sub
		Sub Trigger2_unHit:Controller.Switch(70)=0:End Sub

		Sub Trigger3_Hit
			Controller.Switch(71)=1
			PlaySoundAtVol "Sensor", ActiveBall, 1
			If ActiveBall.VelY>0 Then
				If RightDiverter.IsDropped Then
					RightDiverter.TimerEnabled=0
					RightDiverter2.IsDropped=1
					RightDiverter.IsDropped=0
				End If
			End If
		End Sub
		Sub Trigger3_unHit:Controller.Switch(71)=0:End Sub

	'Left Diverter requires HOLD timer
		Sub SolLDiverter(Enabled)
			If Enabled Then
				LeftDiverter.RotateToEnd
				LeftDiverter.TimerEnabled=0
				LeftDiverter.TimerEnabled=1
			End If
		End Sub

		Sub LeftDiverter_Timer
			LeftDiverter.TimerEnabled=0
			LeftDiverter.RotateToStart
		End Sub

		Sub Trigger11_Hit:Controller.Switch(64)=1:End Sub
		Sub Trigger11_unHit:Controller.Switch(64)=0:End Sub


	'Back Kicker requires HOLD timer
		Sub SolBackKicker(Enabled)
			ModifyBall.Enabled=1
			BKTTimer.Enabled=1
		End Sub

		Sub BKTTimer_Timer
			BKTTimer.Enabled=0
			InTheZone=0
			ModifyBall.Enabled=0
		End Sub


	'Ball Save Gate - Gate is open based on Ball Save light
		Sub BallSaveGate(Enabled)
			If Enabled Then
				Flipper1.RotateToEnd
				SaveTimer.Enabled=0
				SaveTimer.Enabled=1 'in case machine activates solenoid during ball search, enable the close timer
			End If
		End Sub
		Sub Trigger22_Hit:Flipper1.RotateToStart:SaveTimer.Enabled=0:End Sub
		Sub SaveTimer_Timer:Flipper1.RotateToStart:SaveTimer.Enabled=0:End Sub

	'Lower Ball Lock/Kickback on 1st ball code
		Sub SolLockDoor(Enabled)
			If Enabled Then	LockDoor.IsDropped=1
		End Sub

		Sub LeftOutlane_Hit
			If ActiveBall.VelY<0 Then LockDoor.IsDropped=0
			Controller.Switch(36)=1
			PlaySoundAtVol "Sensor", ActiveBall, 1
		End Sub

		Sub LeftOutlane_unHit:Controller.Switch(36)=0:End Sub

	'Left Drop Ramp Stuff (Modified Destruk script)
		Dim undertheramp
		undertheramp=0
		rampstate=0
		Sub EnableRamp(Enabled)
			If Enabled AND rampstate=0 Then
				RRampDir = 1
				UpdateRamp.Enabled = True
				rampstate=1
			End If
		End Sub

		Dim RRampDir, RRAmpCurrPos, RRamp, rampdown
		RRampCurrPos = 0 ' down
		RRampDir = -1     '1 is up -1 is down dir

		Sub UpdateRamp_Timer
			RRampCurrPos = RRampCurrPos + RRampDir
			If RRampCurrPos> 52 Then
				RRampCurrPos = 52
				PlaySoundAtVol "DTLReset", dropramp1, 1
				UpdateRamp.Enabled = 0
				dropramp1.Collidable= 0
			End If
			If RRampCurrPos <0 Then
				RRampCurrPos = 0
				PlaySoundAtVol "left_flipper_down", dropramp1, 1
				dropramp1.Collidable= 1
				UpdateRamp.Enabled = 0
			End If
			dropramp1.HeightBottom = RRampCurrPos
		End Sub

		Dim SA,OSA,RD,ORD
		OSA=0:ORD=False

		Sub UpdateMultipleLamps
			RD=Controller.Lamp(48)
			If RD<>ORD Then
				ORD=RD
				If RD Then
					If undertheramp=0 AND rampstate=1 then
					RRampDir = -1
					RampTimer.Enabled = True
					End If
				End If
			End If

			SA=Controller.Lamp(72)
			If SA<>OSA Then
				SaveTimer.Enabled=0
				SaveTimer.Enabled=1
			OSA=SA
			End If
		End Sub

		Sub UnderRamp_Hit
			If ActiveBall.Vely>0 Then
				'PlaySound "popper"
				RRampDir = 1
				UpdateRamp.Enabled = True
				undertheramp = 1
				rampstate = 1
			End If
			If ActiveBall.Vely<0 Then
				undertheramp = 0
			End If
		End Sub
		Sub underrampexit_Hit
			If ActiveBall.Vely>10 Then
				If rampstate=1 Then
					RampTimer.Enabled = 1
				End If
			End If
			If ActiveBall.Vely<10 Then
				undertheramp = 1
			End If
		End Sub

		Sub RampTimer_Timer '****NOTE: Kicker2_Hit also calls the RampTimer to drop the ramp if the top right lock is hit ***
			undertheramp = 0
			RRampDir = -1
			UpdateRamp.Enabled = True
			rampstate=0
			RampTimer.Enabled = 0
		End Sub

	'Flipper Fix (Script by Uncle Willy)
		dim GameOnFF
		GameOnFF = 0
		Sub GameOn(Enabled)
			If Enabled Then
				GameOnFF = 1
			Else
				GameOnFF = 0
				RampTimer.Enabled = 1
			End If
		End Sub

'FLASHERS
SolCallback(20)= "SetLamp 170," 'Left Ramp Flasher (Playfield) And Backbox Flasher 1
SolCallback(26)= "SetLamp 176," 'Right Rear Flasher (Playfield) And Backbox 8
SolCallback(27)= "SetLamp 178," 'Left Back Flasher Playfield ONLY

Sub Sol22(Enabled)
End Sub

Sub Sol24(Enabled)
End Sub

Dim bsTrough,bsLeftSaucer,bsRightSaucer,dtL,dtR,bsVLock,bsBottomSaucer,PlungerIM, x

Sub Table1_Init
	On Error Resume Next
		With Controller
			.GameName="abv106"
			.SplashInfoLine="Airborne, CAPCOM" & vbnewline & "Table by robbbo43 and DCrosby"
			.HandleMechanics=0
			.HandleKeyboard=0
			.ShowDMDOnly=1
			.ShowFrame=0
			.ShowTitle=0
			.Run
		End With
	On Error Goto 0

	vpmNudge.TiltSwitch=10
	vpmNudge.Sensitivity=5
	vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingshot)
	PinMAMETimer.Interval=PinMAMEInterval:PinMAMETimer.Enabled=1
	vpmCreateEvents AllSwitches
	Set bsTrough=New cvpmBallStack
		bsTrough.InitSw 28,29,30,31,0,0,0,0
		bsTrough.InitKick BallRelease,105,6
		bsTrough.Balls=3
		bsTrough.InitExitSnd "ballrelease","ballrelease"

	Set dtL = New cvpmDropTarget
   	  With dtL
   		.InitDrop Array(array(DropS,DropSa),array(DropH,DropHa),array(DropO,DropOa),array(DropW,DropWa)), Array(61,60,59,58)
        .Initsnd "DTL", "DTResetL"
       End With

  	Set dtR = New cvpmDropTarget
   	  With dtR
   		.InitDrop Array(array(DropA,DropAa),array(DropI,DropIa),array(DropR,DropRa)), Array(53,52,51)
        .Initsnd "DTR", "DTResetR"
       End With

	Set bsLeftSaucer=New cvpmBallStack
		bsLeftSaucer.InitSaucer Kicker1,57,138,20
		'bsLeftSaucer.InitExitSnd "scoopexit","SolOn"
		bsLeftSaucer.InitExitSnd "popper_ball", "popper_ball"

	Set bsRightSaucer=New cvpmBallStack
		bsRightSaucer.InitSaucer Kicker2,56,280,15
		'bsRightSaucer.InitExitSnd "scoopexit","SolOn"
		bsRightSaucer.InitExitSnd "popper_ball", "popper_ball"

	Set bsVLock=New cvpmVLock
		bsVLock.InitVLock Array(Trigger5,Trigger6),Array(Kicker3,Kicker4),Array(47,46)
		bsVLock.CreateEvents "bsVLock"
        bsVLock.InitSnd "popper", "popper"

	Set bsBottomSaucer=New cvpmBallStack
		bsBottomSaucer.InitSaucer Kicker5,39,0,50
		'bsBottomSaucer.InitExitSnd "Popper","SolOn"
        bsBottomSaucer.InitExitSnd "popper_ball", "popper_ball"

'Right Diverter Init
		RightDiverter2.IsDropped=1
'Left Drop Ramp Init
		dropramp1.collidable=true
		dropramp1.visible=true
'Slingshot Init
		LeftSling.IsDropped=1:LeftSling2.IsDropped=1:LeftSling3.IsDropped=1
		LHammer.IsDropped=1:LHammer2.IsDropped=1:LHammer3.IsDropped=1
		RightSling.IsDropped=1:RightSling2.IsDropped=1:RightSling3.IsDropped=1
		RHammer.IsDropped=1:RHammer2.IsDropped=1:RHammer3.IsDropped=1
End Sub

'Keys
	Sub Table1_KeyDown(ByVal KeyCode)
		If KeyCode=LeftFlipperKey Then
			Controller.Switch(5)=1
			Controller.Switch(25)=1
		End If
		If KeyCode=RightFlipperKey Then
			Controller.Switch(6)=1
			Controller.Switch(26)=1
		End If
		If keycode = PlungerKey Then PlaysoundAtVol "PlungerPull", Plunger, 1:Plunger.Pullback
		If KeyCode = 3 Then Controller.Switch(11) = True
		If vpmKeyDown(KeyCode) Then Exit Sub
	End Sub

	Sub Table1_KeyUp(ByVal KeyCode)
		If KeyCode=LeftFlipperKey Then
			Controller.Switch(5)=0
			Controller.Switch(25)=0
		End If
		If KeyCode=RightFlipperKey Then
			Controller.Switch(6)=0
			Controller.Switch(26)=0
		End If
		If keycode = PlungerKey Then StopSound "PlungerPull":Plunger.Fire
		If KeyCode = 3 Then Controller.Switch(11) = False
		'If KeyCode=KeyFront Then Controller.Switch(11)=0
		If vpmKeyUp(KeyCode) Then Exit Sub
	End Sub

'Switches

	Sub Drain_Hit:PlaySoundAtVol "fx_drain", drain, 1:bsTrough.AddBall Me:End Sub
	Sub BallRelease_UnHit():End Sub
	Sub sw27_Hit:Controller.Switch(27)=1:PlaySoundAtVol "sensor", ActiveBall, 1:End Sub
	Sub sw27_unHit:Controller.Switch(27)=0:End Sub
	Sub Kicker1_Hit:PlaySoundAtVol "kicker_enter", kicker1, 1:bsLeftSaucer.AddBall 0:End Sub
	Sub Kicker2_Hit:PlaySoundAtVol "kicker_enter", kicker2, 1:RightDiverter.TimerEnabled=0:RightDiverter.IsDropped=0:RightDiverter2.IsDropped=1:bsRightSaucer.AddBall 0:RampTimer.Enabled=1: End Sub
	Sub RampRight_Hit:Controller.Switch(55)=1:End Sub
	Sub RampRight_UnHit:Controller.Switch(55)=0:End Sub

	Sub sw41_Hit():vpmTimer.PulseSw 41:PlaySoundAtVol "target", ActiveBall, 1:ME.TimerEnabled=1:End Sub
	Sub sw41_Timer():ME.TimerEnabled=0:End Sub
	Sub sw42_Hit():vpmTimer.PulseSw 42:PlaySoundAtVol "target", ActiveBall, 1:ME.TimerEnabled=1:End Sub
	Sub sw42_Timer():ME.TimerEnabled=0:End Sub
	Sub sw43_Hit():vpmTimer.PulseSw 43:PlaySoundAtVol "target", ActiveBall, 1:ME.TimerEnabled=1:End Sub
	Sub sw43_Timer():ME.TimerEnabled=0:End Sub
	Sub sw49_Hit():vpmTimer.PulseSw 49:PlaySoundAtVol "target", ActiveBall, 1:ME.TimerEnabled=1:End Sub
	Sub sw49_Timer():ME.TimerEnabled=0:End Sub
	Sub sw50_Hit():vpmTimer.PulseSw 50:PlaySoundAtVol "target", ActiveBall, 1:ME.TimerEnabled=1:End Sub
	Sub sw50_Timer():ME.TimerEnabled=0:End Sub
	Sub sw62_Hit():vpmTimer.PulseSw 62:PlaySoundAtVol "target", ActiveBall, 1:ME.TimerEnabled=1:End Sub
	Sub sw62_Timer():ME.TimerEnabled=0:End Sub
	Sub sw65_Hit():vpmTimer.PulseSw 65:PlaySoundAtVol "target", ActiveBall, 1:ME.TimerEnabled=1:End Sub
	Sub sw65_Timer():ME.TimerEnabled=0:End Sub
	Sub sw66_Hit():vpmTimer.PulseSw 66: PlaySoundAtVol "target", ActiveBall, 1:ME.TimerEnabled=1:End Sub
	Sub sw66_Timer():ME.TimerEnabled=0:End Sub

	'Air Drop
	Sub DropAa_Hit:dtr.Hit 1:DropADir = 0:DropAa.TimerEnabled = 1:End Sub
	Sub DropIa_Hit:dtr.Hit 2:DropIDir = 0:DropIa.TimerEnabled = 1:End Sub
	Sub DropRa_Hit:dtr.Hit 3:DropRDir = 0:DropRa.TimerEnabled = 1:End Sub
	Sub DropSa_Hit:dtl.Hit 1:DropSDir = 0:DropSa.TimerEnabled = 1:End Sub
	Sub DropHa_Hit:dtl.Hit 2:DropHDir = 0:DropHa.TimerEnabled = 1:End Sub
	Sub DropOa_Hit:dtl.Hit 3:DropODir = 0:DropOa.TimerEnabled = 1:End Sub
	Sub DropWa_Hit:dtl.Hit 4:DropWDir = 0:DropWa.TimerEnabled = 1:End Sub


	Sub LeftInlane_Hit:Controller.Switch(35)=1:PlaySoundAtVol "Sensor", ActiveBall, 1:End Sub
	Sub LeftInlane_UnHit:Controller.Switch(35)=0:End Sub
	Sub RightInlane_Hit:Controller.Switch(37)=1:PlaySoundAtVol "Sensor", ActiveBall, 1:End Sub
	Sub RightInlane_UnHit:Controller.Switch(37)=0:End Sub
	Sub RightOutlane_Hit:Controller.Switch(38)=1:PlaySoundAtVol "Sensor", ActiveBall, 1:End Sub
	Sub RightOutlane_UnHit:Controller.Switch(38)=0:End Sub

	Sub Kicker5_Hit:PlaySoundAtVol "kicker_enter_left", Kicker5, 1:LockDoor.IsDropped=0:bsBottomSaucer.AddBall 0:End Sub

	Dim Speed1
	Sub Kicker6_Hit:Speed1=SQR(ActiveBall.VelY*ActiveBall.VelY)/3:vpmTimer.PulseSw 73:Kicker6.DestroyBall:PlaySoundAtVol "rail", Kicker6, 1:Kicker8.TimerEnabled=1:End Sub
	Kicker8.TimerInterval=150
	Sub Kicker8_Timer ()
		Kicker8.CreateBall: kicker8.Kick 180,Speed1
		Kicker8.TimerEnabled=0
	End Sub
	Dim Speed2
	Sub Kicker7_Hit:Speed2=SQR(ActiveBall.VelY*ActiveBall.VelY)/3:vpmTimer.PulseSw 74:Kicker7.DestroyBall:PlaySoundAtVol "rail", Kicker7, 1:Kicker9.TimerEnabled=1:End Sub
	Kicker9.TimerInterval=150
	Sub Kicker9_Timer ()
		Kicker9.CreateBall: Kicker9.Kick 180,Speed2
		Kicker9.TimerEnabled=0
	End Sub

	Sub Trigger12_Hit:ActiveBall.VelX=0:ActiveBall.VelY=3:End Sub
	Sub Trigger13_Hit:ActiveBall.VelX=0:ActiveBall.VelY=3:End Sub
	Sub Trigger14_Hit:ActiveBall.VelX=0:ActiveBall.VelY=3:End Sub
	Sub Trigger15_Hit:ActiveBall.VelX=0:ActiveBall.VelY=3:End Sub
	Sub Trigger16_Hit:ActiveBall.VelX=0:ActiveBall.VelY=3:End Sub
	Sub Trigger17_Hit:ActiveBall.VelX=0:ActiveBall.VelY=3:End Sub
	Sub Trigger20_Hit:ActiveBall.VelX=0:ActiveBall.VelY=1:End Sub
	Sub Trigger18_Hit:ActiveBall.VelX=1.5:ActiveBall.VelY=1.5:End Sub
	Sub Trigger19_Hit:ActiveBall.VelX=-1.5:ActiveBall.VelY=1.5:End Sub
'Slingshots
	'sling animation scripting from JPSalas
 	Dim LStep, RStep

 Sub LeftSlingShot_Slingshot
 	PlaySoundAtBallVol "LSling", 1:vpmTimer.PulseSw 33:LStep = 0:Me.TimerEnabled = 1
  End Sub
	'Sub LeftSlingShot_Slingshot:PlaySound "LSling":LeftSling.IsDropped = 0:vpmTimer.PulseSw 33:LStep = 0:Me.TimerEnabled = 1:End Sub
	Sub LeftSlingShot_Timer
     Select Case LStep
         Case 0:LeftSLing.IsDropped = 0:LHammer.IsDropped = 0
         Case 1: 'pause
         Case 2:LeftSLing.IsDropped = 1:LHammer.IsDropped=1:LeftSLing2.IsDropped = 0:LHammer2.IsDropped = 0
         Case 3:LeftSLing2.IsDropped = 1:LHammer2.IsDropped = 1:LeftSLing3.IsDropped = 0:LHammer3.IsDropped = 0
         Case 4:LeftSLing3.IsDropped = 1:LHammer3.IsDropped = 1:Me.TimerEnabled = 0
     End Select
      LStep = LStep + 1
	End Sub
 Sub RightSlingShot_Slingshot
 	PlaySoundAtBallVol "RSling", 1:vpmTimer.PulseSw 34:RStep = 0:Me.TimerEnabled = 1
  End Sub
	'Sub RightSlingShot_Slingshot:PlaySound "RSling":RightSling.IsDropped = 0:vpmTimer.PulseSw 34:RStep = 0:Me.TimerEnabled = 1:End Sub
	Sub RightSlingShot_Timer
     Select Case RStep
         Case 0:RightSLing.IsDropped = 0:RHammer.IsDropped = 0
         Case 1: 'pause
         Case 2:RightSLing.IsDropped = 1:RHammer.IsDropped=1:RightSLing2.IsDropped = 0:RHammer2.IsDropped = 0
         Case 3:RightSLing2.IsDropped = 1:RHammer2.IsDropped = 1:RightSLing3.IsDropped = 0:RHammer3.IsDropped = 0
         Case 4:RightSLing3.IsDropped = 1:RHammer3.IsDropped = 1:Me.TimerEnabled = 0
     End Select
      RStep = RStep + 1
	End Sub


'Bumpers
	Sub Bumper1_Hit:RandomSoundBumper:vpmTimer.PulseSw 78:Bumper1Cover.image = "Bumper_On" :bump1 = 1:Me.TimerEnabled = 1:End Sub
	Sub Bumper1_Timer:Bumper1Cover.image = "Bumper_off":Me.TimerEnabled = 0:End Sub
	Sub Bumper2_Hit:RandomSoundBumper:vpmTimer.PulseSw 79:Bumper2Cover.image = "Bumper_On" :bump2 = 1:Me.TimerEnabled = 1:End Sub
	Sub Bumper2_Timer:Bumper2Cover.image = "Bumper_off":Me.TimerEnabled = 0:End Sub
	Sub Bumper3_Hit:RandomSoundBumper:vpmTimer.PulseSw 80:Bumper3Cover.image = "Bumper_On" :bump3 = 1:Me.TimerEnabled = 1:End Sub
	Sub Bumper3_Timer:Bumper3Cover.image = "Bumper_off":Me.TimerEnabled = 0:End Sub

	Sub RandomSoundBumper()
		Select Case Int(Rnd*3)+1
			Case 1 : PlaySoundAtVol "bumper_1", ActiveBall, 1
			Case 2 : PlaySoundAtVol "bumper_2", ActiveBall, 1
			Case 3 : PlaySoundAtVol "bumper_3", ActiveBall, 1
		End Select
	End Sub

'Upper Jet Ball Release direction (Script by Destruk)
	Dim RM
	Sub Trigger21_Hit
	RM=RM+1
	If RM>1 Then RM=0
	If RM=0 Then
		ActiveBall.VelX=1
	Else
		ActiveBall.VelX=-1
	End If
	End Sub

'Injector Scripting (by Destruk)
	Dim BKTBall,InTheZone
	InTheZone=0

	Sub BKT_Hit
		Set BKTBall=ActiveBall
		InTheZone=1
	End Sub

	Sub BKT_unHit
		InTheZone=0

		'Injection_Bar_Fwd.IsDropped=1
		Dim x
		For each x in Injection_Bar_Fwd
			x.visible=0
			x.collidable=0
		Next
		'Injection_Bar_Back.isDropped=0
		Dim y
		For each y in Injection_Bar_Back
			y.collidable=1
			y.visible=1
		Next
	End Sub

	Sub ModifyBall_Timer
		If InTheZone=1 Then
			BKTBall.VelY=BKTBall.VelY+5
				Dim x
				For each x in Injection_Bar_Fwd
					x.collidable=1
					x.visible=1
				Next
				Dim y
				For each y in Injection_Bar_Back
					y.collidable=0
					y.visible=0
				Next
			If BKTBall.VelX<-2 Then BKTBall.VelX=BKTBall.VelX+5
		End If
	End Sub


'****************************************
'  JP's Fading Lamps Routine
'      Based on PD's Fading Lights
' SetLamp 0 is Off
' SetLamp 1 is On
' LampState(x) current state
'****************************************

Dim LampState(200), FadingLevel(200)
Dim FlashSpeedUp(200), FlashSpeedDown(200), FlashMin(200), FlashMax(200), FlashLevel(200), ModulationLevel(200)

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

Sub UpdateLamps
	NFadeL 5, L5
	NFadeL 6, L6
	NFadeL 7, L7
	NFadeLm 8, L8a
	NFadeLm 8, L8b
	NFadeL 8, L8
	NFadeL 9, L9
	NFadeLm 11, L11a
	NFadeLm 11, L11b
	NFadeLm 11, L11c
    NFadeL 11, L11
	NFadeLm 12, L12a
	NFadeLm 12, L12b
	NFadeLm 12, L12c
	NFadeLm 12, L12d
	NFadeLm 12, L12e
	NFadeL 12, L12
	NFadeL 13, L13
	NFadeL 14, L14
	NFadeL 15, L15
	NFadeLm 16, L16
	NFadeL 16, L16a
	NFadeL 17, L17
	NFadeL 18, L18
	NFadeL 19, L19
	NFadeL 20, L20
	NFadeL 21, L21
	NFadeL 22, L22
	NFadeLm 23, L23a
	NFadeLm 23, L23b
	NFadeLm 23, L23c
	NFadeL 23, L23
	NFadeLm 24, L24a
	NFadeLm 24, L24b
	NFadeLm 24, L24c
	NFadeLm 24, L24d
	NFadeLm 24, L24e
	NFadeL 24, L24
	NFadeLm 25, L25a
	NFadeLm 25, L25b
	NFadeL 25, L25
	NFadeLm 26, L26a
	NFadeLm 26, L26b
	NFadeL 26, L26
	NFadeLm 27, L27a
	NFadeLm 27, L27b
	NFadeL 27, L27
	NFadeLm 28, L28a
	NFadeLm 28, L28b
	NFadeL 28, L28
	NFadeLm 29, L29a
	NFadeLm 29, L29b
	NFadeLm 29, L29c
	NFadeLm 29, L29d
	NFadeLm 29, L29e
	NFadeL 29, L29
	NFadeLm 30, L30
	NFadeL 30, L30
	NFadeLm 31, L31a
	NFadeL 31, L31
	NFadeLm 32, L32a
	NFadeL 32, L32
	NFadeL 33, L33
	NFadeL 34, L34
	NFadeLm 35, L35a
	NFadeLm 35, L35b
	NFadeLm 35, L35c
	NFadeLm 35, L35d
	NFadeLm 35, L35e
	NFadeL 35, L35
	NFadeL 36, L36
	NFadeL 37, L37
	NFadeL 38, L38
	NFadeL 39, L39
	NFadeL 40, L40
	NFadeL 41, L41
	NFadeL 42, L42
	NFadeL 43, L43
	NFadeLm 44, L44a
	NFadeLm 44, L44b
	NFadeL 44, L44
	NFadeLm 45, L45a
	NFadeLm 45, L45b
	NFadeL 45, L45
	NFadeLm 46, L46a
	NFadeLm 46, L46b
	NFadeLm 46, L46c
	NFadeLm 46, L46d
	NFadeLm 46, L46e
	NFadeL 46, L46
	NFadeLm 47, L47a
	NFadeLm 47, L47b
	NFadeL 47, L47
	NFadeL 48, L48
	NFadeL 49, L49
	NFadeL 50, L50
	NFadeL 51, L51
	NFadeL 52, L52
	NFadeL 53, L53
	NFadeL 54, L54
	NFadeLm 55, L55a
	NFadeLm 55, L55b
	NFadeL 55, L55
	NFadeLm 56, L56a
	NFadeLm 56, L56b
	NFadeL 56, L56
	NFadeL 57, L57
	NFadeL 58, L58
	NFadeL 59, L59
	NFadeL 60, L60
	NFadeL 61, L61
	NFadeL 62, L62
	NFadeL 63, L63
	NFadeL 64, L64
	NFadeL 65, L65
	NFadeL 66, L66
	NFadeL 67, L67
	NFadeL 68, L68
	NFadeL 69, L69
	NFadeLm 70, L70a
	NFadeLm 70, L70b
	NFadeL 70, L70
	NFadeL 71, L71
	NFadeL 72, L72
	NFadeL 73, L73
	NFadeL 74, L74
	NFadeL 75, L75
	NFadeL 76, L76
	NFadeLm 77, L77a
	NFadeLm 77, L77b
	NFadeL 77, L77
	NFadeLm 78, L78a
	NFadeLm 78, L78b
	NFadeL 78, L78
	NFadeLm 79, L79a
	NFadeLm 79, L79b
	NFadeL 79, L79
	NFadeLm 80, L80a
	NFadeL 80, L80
	NFadeL 81, L81
	NFadeL 82, L82
	NFadeL 83, L83
	NFadeL 84, L84
	'NFadeLm 85, L91
	NFadeLm 85, L85a
	NFadeLm 85, L85b
	NFadeLm 85, L85c
	NFadeL 85, L85
	NFadeLm 86, L86a
	NFadeLm 86, L86b
	NFadeL 86, L86
	NFadeLm 87, L87a
	NFadeLm 87, L87b
	NFadeLm 87, L87c
	NFadeLm 87, L87d
	NFadeL 87, L87
	NFadeLm 89, LBumper1a
	NFadeL 89, LBumper1
	NFadeLm 90, LBumper3a
	NFadeL 90, LBumper3
	NFadeLm 91, LBumper2a
	NFadeLm 91, LBumper2
	NFadeLm 91, L91a
	NFadeLm 91, L91b
	NFadeL 91, L91
	'NFadeLm 93, rightcenterplastic
	'NFadeLm 93, L105
	NFadeLm 92, L92a
	NFadeLm 92, L92b
	NFadeL 92, L92
	NFadeL 94, L93
	NFadeL 95, L94
	NFadeL 96, L95
	NFadeL 97, L96
	NFadeL 98, L97
	NFadeL 99, L98
	NFadeLm 100, L99a
	NFadeL 100, L99
	NFadeL 101, L100
	NFadeL 102, L101
	NFadeL 103, L102
	NFadeLm 104, L103a
	NFadeLm 104, L103b
	NFadeLm 104, L103c
	NFadeLm 104, L103d
	NFadeLm 104, L103e
	NFadeL 104, L103
	'NFadeL 106, L105
	FlashAR 178, InjectorLights, "InjectorPanel_On", "InjectorPanel_Half","InjectorPanel_Off"
	Flash 170, F20
	Flash 176, F26
	Flashm 178, F27a
	Flash 178, F27
End Sub

Sub InitLamps()
    Dim x
    For x = 0 to 200
        LampState(x) = 0         ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4       ' used to track the fading state
        FlashSpeedUp(x) = 0.2    ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.05 ' slower speed when turning off the flasher
        FlashMax(x) = 1          ' the maximum value when on, usually 1
        FlashMin(x) = 0          ' the minimum value when off, usually 0
        FlashLevel(x) = 0        ' the intensity of the flashers, usually from 0 to 1
		ModulationLevel(x) = 0		 ' the starting modulation level, from 0 to 255
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

'Lights

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

Sub SetLampMod(nr, value)
    If value > 0 Then
		LampState(nr) = 1
	Else
		LampState(nr) = 0
	End If
	ModulationLevel(nr) = value
End Sub

Sub LampMod(nr, object)
	Object.IntensityScale = ModulationLevel(nr)/128
End Sub

Sub FlashAR(nr, ramp, a, b, c)
	Select Case LampState(nr)
		Case 2:LampState(nr) = 0 'Off
		Case 3:ramp.image = c:LampState(nr) = 2 'fading...
		Case 4:ramp.image = b:LampState(nr) = 3 'fading...
		Case 5:ramp.image = a:LampState(nr) = 1 'ON
	End Select
End Sub

Sub FlashARm(nr, ramp, a, b, c)
	Select Case LampState(nr)
		Case 2:LampState(nr) = 0 'Off
		Case 3:ramp.image = c:LampState(nr) = 2 'fading...
		Case 4:ramp.image = b:LampState(nr) = 3 'fading...
		Case 5:ramp.image = a:LampState(nr) = 1 'ON
	End Select
End Sub

' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX, Rothbauerw, Thalamus and Herweh
' PlaySound sound, 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
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

' set position as table object and Vol + RndPitch manually

Sub PlaySoundAtVolPitch(sound, tableobj, Vol, RndPitch)
  PlaySound sound, 1, Vol, AudioPan(tableobj), RndPitch, 0, 0, 1, AudioFade(tableobj)
End Sub

'Set all as per ball position & speed.

Sub PlaySoundAtBall(soundname)
  PlaySoundAt soundname, ActiveBall
End Sub

'Set position as table object and Vol manually.

Sub PlaySoundAtVol(sound, tableobj, Volume)
  PlaySound sound, 1, Volume, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'Set all as per ball position & speed, but Vol Multiplier may be used eg; PlaySoundAtBallVol "sound",3

Sub PlaySoundAtBallVol(sound, VolMult)
  PlaySound sound, 0, Vol(ActiveBall) * VolMult, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtBallAbsVol(sound, VolMult)
  PlaySound sound, 0, VolMult, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub

' requires rampbump1 to 7 in Sound Manager

Sub RandomBump(voladj, freq)
  Dim BumpSnd:BumpSnd= "rampbump" & CStr(Int(Rnd*7)+1)
  PlaySound BumpSnd, 0, Vol(ActiveBall)*voladj, AudioPan(ActiveBall), 0, freq, 0, 1, AudioFade(ActiveBall)
End Sub

' set position as bumperX and Vol manually. Allows rapid repetition/overlaying sound

Sub PlaySoundAtBumperVol(sound, tableobj, Vol)
  PlaySound sound, 1, Vol, AudioPan(tableobj), 0,0,1, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBOTBallZ(sound, BOT)
  PlaySound sound, 0, ABS(BOT.velz)/17, Pan(BOT), 0, Pitch(BOT), 1, 0, AudioFade(BOT)
End Sub

' play a looping sound at a location with volume
Sub PlayLoopSoundAtVol(sound, tableobj, Vol)
  PlaySound sound, -1, Vol, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

Function RndNum(min, max)
  RndNum = Int(Rnd() * (max-min + 1) ) + min ' Sets a random number between min and max
End Function

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = tableobj.y * 2 / table1.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = tableobj.x * 2 / table1.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = ball.x * 2 / table1.width-1
  If tmp > 0 Then
    Pan = Csng(tmp ^10)
  Else
    Pan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
  Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function VolMulti(ball,Multiplier) ' Calculates the Volume of the sound based on the ball speed
  VolMulti = Csng(BallVel(ball) ^2 / 150 ) * Multiplier
End Function

Function DVolMulti(ball,Multiplier) ' Calculates the Volume of the sound based on the ball speed
  DVolMulti = Csng(BallVel(ball) ^2 / 150 ) * Multiplier
  debug.print DVolMulti
End Function

Function BallRollVol(ball) ' Calculates the Volume of the sound based on the ball speed
  BallRollVol = Csng(BallVel(ball) ^2 / (80000 - (79900 * Log(RollVol) / Log(100))))
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
  Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
  BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

Function BallVelZ(ball) 'Calculates the ball speed in the -Z
  BallVelZ = INT((ball.VelZ) * -1 )
End Function

Function VolZ(ball) ' Calculates the Volume of the sound based on the ball speed in the Z
  VolZ = Csng(BallVelZ(ball) ^2 / 200)*1.2
End Function

'*** Determines if a Points (px,py) is inside a 4 point polygon A-D in Clockwise/CCW order

Function InRect(px,py,ax,ay,bx,by,cx,cy,dx,dy)
  Dim AB, BC, CD, DA
  AB = (bx*py) - (by*px) - (ax*py) + (ay*px) + (ax*by) - (ay*bx)
  BC = (cx*py) - (cy*px) - (bx*py) + (by*px) + (bx*cy) - (by*cx)
  CD = (dx*py) - (dy*px) - (cx*py) + (cy*px) + (cx*dy) - (cy*dx)
  DA = (ax*py) - (ay*px) - (dx*py) + (dy*px) + (dx*ay) - (dy*ax)

  If (AB <= 0 AND BC <=0 AND CD <= 0 AND DA <= 0) Or (AB >= 0 AND BC >=0 AND CD >= 0 AND DA >= 0) Then
    InRect = True
  Else
    InRect = False
  End If
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
        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
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
	PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

'Sounds

Sub RailTrigger1_Hit

	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  	if finalspeed > 2 AND finalspeed < 8 Then
		'PlaySound "rail_low"
		PlaySoundAtVol "rail_low_slower", ActiveBall, 1
	elseif finalspeed >= 8 Then
		PlaySoundAtVol "rail", ActiveBall, 1
	End if
End Sub
Sub RailTrigger2_Hit

	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  	if finalspeed > 2 AND finalspeed < 8 Then
		'PlaySound "rail_low"
		PlaySoundAtVol "rail_low_slower", ActiveBall, 1
	elseif finalspeed >= 8 Then
		PlaySoundAtVol "rail", ActiveBall, 1
	End if
End Sub
Sub RailTrigger3_Hit
	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  	if finalspeed > 2 AND finalspeed < 8 Then
		'PlaySound "rail_low"
		PlaySoundAtVol "rail_low_slower", ActiveBall, 1
	elseif finalspeed >= 8 Then
		PlaySoundAtVol "rail", ActiveBall, 1
	End if
End Sub
Sub RailTrigger4_Hit
		dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  	if finalspeed > 2 AND finalspeed < 8 Then
		'PlaySound "rail_low"
		PlaySoundAtVol "rail_low_slower", ActiveBall, 1
	elseif finalspeed >= 8 Then
		PlaySoundAtVol "rail", ActiveBall, 1
	End if
End Sub
Sub RailTrigger5_Hit
		dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  	if finalspeed > 2 AND finalspeed < 8 Then
		'PlaySound "rail_low"
		PlaySoundAtVol "rail_low_slower", ActiveBall, 1
	elseif finalspeed >= 8 Then
		PlaySoundAtVol "rail", ActiveBall, 1
	End if
End Sub
Sub RailTrigger4a_Hit
	StopSound "rail"
	StopSound "rail_low_slower"
	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  	if finalspeed > 2 AND finalspeed < 8 Then
		'PlaySound "rail_low"
		PlaySoundAtVol "rail_low_slower", ActiveBall, 1
	elseif finalspeed >= 8 Then
		PlaySoundAtVol "rail", ActiveBall, 1
	End if
End Sub
Sub RailTrigger5a_Hit
	StopSound "rail"
	StopSound "rail_low_slower"
	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  	if finalspeed > 2 AND finalspeed < 8 Then
		'PlaySound "rail_low"
		PlaySoundAtVol "rail_low_slower", ActiveBall, 1
	elseif finalspeed >= 8 Then
		PlaySoundAtVol "rail", ActiveBall, 1
	End if
	PlaySoundAtVol "rampdrop", ActiveBall, 1
End Sub
Sub RailTrigger6a_Hit
	StopSound "rail"
	StopSound "rail_low_slower"
	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  	if finalspeed > 2 AND finalspeed < 8 Then
		'PlaySound "rail_low"
		PlaySoundAtVol "rail_low_slower", ActiveBall, 1
	elseif finalspeed >= 8 Then
		PlaySoundAtVol "rail", ActiveBall, 1
	End if
	PlaySoundAtVol "rampdrop", ActiveBall, 1
End Sub
Sub RailTrigger4b_Hit
	StopSound "rail"
	StopSound "rail_low_slower"
	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  	if finalspeed > 2 AND finalspeed < 8 Then
		'PlaySound "rail_low"
		PlaySoundAtVol "rail_low_slower", ActiveBall, 1
	elseif finalspeed >= 8 Then
		PlaySoundAtVol "rail", ActiveBall, 1
	End if
	PlaySoundAtVol "rampdrop", ActiveBall, 1
End Sub
Sub RailTrigger4c_Hit
	StopSound "rail"
	StopSound "rail_low_slower"
	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  	if finalspeed > 2 AND finalspeed < 8 Then
		'PlaySound "rail_low"
		PlaySoundAtVol "rail_low_slower", ActiveBall, 1
	elseif finalspeed >= 8 Then
		PlaySoundAtVol "rail", ActiveBall, 1
	End if
	PlaySoundAtVol "rampdrop", ActiveBall, 1
End Sub
Sub RailTrigger1a_Hit
	StopSound "rail"
	StopSound "rail_low_slower"
	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  	if finalspeed > 2 AND finalspeed < 8 Then
		'PlaySound "rail_low"
		PlaySoundAtVol "rail_low_slower", ActiveBall, 1
	elseif finalspeed >= 8 Then
		PlaySoundAtVol "rail", ActiveBall, 1
	End if
	PlaySoundAtVol "rampdrop", ActiveBall, 1
End Sub
Sub RailTrigger1b_Hit
	StopSound "rail"
	StopSound "rail_low_slower"
	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  	if finalspeed > 2 AND finalspeed < 8 Then
		'PlaySound "rail_low"
		PlaySoundAtVol "rail_low_slower", ActiveBall, 1
	elseif finalspeed >= 8 Then
		PlaySoundAtVol "rail", ActiveBall, 1
	End if
	PlaySoundAtVol "rampdrop", ActiveBall, 1
End Sub
Sub RailTrigger1end_Hit
	StopSound "rail"
	StopSound "rail_low_slower"
	PlaySoundAtVol "ball_bounce", ActiveBall, 1
End Sub
Sub RailTrigger4end_Hit
	StopSound "rail"
	StopSound "rail_low_slower"
	PlaySoundAtVol "ball_bounce", ActiveBall, 1
End Sub


'********************
'    JP Flippers
'********************

Sub SolLFlipper(Enabled)
    If GameOnFF = 0 Then
        LeftFlipper.RotateToStart
	Else
		If Enabled Then
        PlaySound "left_flipper_up", 0, 1, -0.1, 0.15
        LeftFlipper.RotateToEnd
		Else
        PlaySound "left_flipper_down", 0, 1, -0.1, 0.15
        LeftFlipper.RotateToStart
		End If
	End If
End Sub

Sub SolRFlipper(Enabled)
    If GameOnFF = 0 Then
        RightFlipper.RotateToStart
	Else
		If Enabled Then
        PlaySound "right_flipper_up", 0, 1, 0.1, 0.15
        RightFlipper.RotateToEnd
		Else
        PlaySound "right_flipper_down", 0, 1, 0.1, 0.15
        RightFlipper.RotateToStart
		End If
    End If
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySoundAtBallVol "fx_rubber_flipper", parm
End Sub

Sub RightFlipper_Collide(parm)
    PlaySoundAtBallVol "fx_rubber_flipper", parm
End Sub

Sub RightFlipper1_Collide(parm)
    PlaySoundAtBallVol "fx_rubber_flipper", parm
End Sub

Sub arubberposts_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 14 then
		PlaySoundAtVol "bump", ActiveBall, 1
	End if
	If finalspeed >= 4 AND finalspeed <= 14 then
 		RandomSoundRubber()
 	End If
	If finalspeed < 4 AND finalspeed > 1 then
 		RandomSoundRubberLowVolume()
 	End If

End sub

Sub arubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 16 then
		PlaySoundAtVol "bump", ActiveBall, 1
	End if
	If finalspeed >= 8 AND finalspeed <= 16 then
 		RandomSoundRubber()
 	End If

End sub

Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySoundAtVol "rubber_hit_1", ActiveBall, 1
		Case 2 : PlaySoundAtVol "rubber_hit_2", ActiveBall, 1
		Case 3 : PlaySoundAtVol "rubber_hit_3", ActiveBall, 1
	End Select
End Sub

Sub RandomSoundRubberLowVolume()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySoundAtVol "rubber_hit_1_low", ActiveBall, 1
		Case 2 : PlaySoundAtVol "rubber_hit_2_low", ActiveBall, 1
		Case 3 : PlaySoundAtVol "rubber_hit_3_low", ActiveBall, 1
	End Select
End Sub

dim MetalsSoundActive
dim WoodsSoundActive

'Metals Rails
Sub ametalrails_Hit(idx)
	If MetalsSoundActive=True Then Exit Sub
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  	If finalspeed > 2 AND finalspeed < 6  AND ABS(activeball.vely) > 2 Then
		PlaySoundAtVol "metalhit_thin_low", ActiveBall, 1
		MetalsSoundActive=True
		MetalSoundTimer.Enabled=1
	ElseIf finalspeed >= 6 AND ABS(activeball.vely) > 2 Then
		PlaySoundAtVol "metalhit_thin", ActiveBall, 1
		MetalsSoundActive=True
		MetalSoundTimer.Enabled=1
	End If
End sub

'Metals2 (sides)
Sub ametalwalls_Hit(idx)
	If MetalsSoundActive=True Then Exit Sub
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  	If finalspeed > 1 AND finalspeed < 6  AND ABS(activeball.vely) > 1 Then
		PlaySoundAtVol "metalhit2_low", ActiveBall, 1
		MetalsSoundActive=True
		MetalSoundTimer.Enabled=1
	ElseIf finalspeed >= 6 AND ABS(activeball.vely) > 2 Then
		PlaySoundAtVol "metalhit2", ActiveBall, 1
		MetalsSoundActive=True
		MetalSoundTimer.Enabled=1
	End If
End sub

Sub MetalSoundTimer_Timer()
	MetalsSoundActive=False
	Me.Enabled=0
End Sub

'Pins
Sub apins_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  	if finalspeed > 6 AND finalspeed < 16 then
		PlaySoundAtVol "pinhit_low", ActiveBall, 1
	elseif finalspeed >= 16 then
		PlaySoundAtVol "pinhit", ActiveBall, 1
	End if

End sub



'Gate Sounds
Sub agates_Hit(idx):PlaySoundAtVol "gate", ActiveBall, 1: End Sub

'*****************************************
'	ninuzzu's	BALL SHADOW
'*****************************************
Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5) 'Ball Shadow Objects for Muliball

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


'new primitave drop target

'AIR SHOW Drop Animations
  dim DropADir, DropIDir, DropRDir, DropSDir, DropHDir, DropODir, DropWDir
  dim DropAPos, DropIPos, DropRPos, DropSPos, DropHPos, DropOPos, DropWPos

  DropADir = 1:DropIDir = 1:DropRDir = 1:DropSDir = 1:DropHDir = 1:DropODir = 1:DropWDir = 1
  DropAPos = 0:DropIDir = 0:DropRPos = 0:DropSPos = 0:DropHPos = 0:DropOPos = 0:DropWPos = 0

  'Targets Init
  DropAa.TimerEnabled = 1:DropIa.TimerEnabled = 1:DropRa.TimerEnabled = 1:DropSa.TimerEnabled = 1:DropHa.TimerEnabled = 1:DropOa.TimerEnabled = 1:DropWa.TimerEnabled = 1


'Animations
  Sub DropAa_Timer()
  Select Case DropAPos
        Case 0: DropAp.z=25
				 If DropADir = 1 then
					DropAa.TimerEnabled = 0
				 else
			     end if
        Case 1: DropAp.z=21
        Case 2: DropAp.z=17
        Case 3: DropAp.z=13
        Case 4: DropAp.z=9
        Case 5: DropAp.z=5
        Case 6: DropAp.z=1
        Case 7: DropAp.z=0
        Case 8: DropAp.z=-1
        Case 9: DropAp.z=-5
        Case 10: DropAp.z=-9
        Case 11: DropAp.z=-13
        Case 12: DropAp.z=-17
        Case 13: DropAp.z=-21
        Case 14: DropAp.z=-24
				 If DropADir = 1 then
				 else
					DropAa.TimerEnabled = 0
			     end if


End Select
	If DropADir = 1 then
		If DropApos>0 then DropApos=DropApos-1
	else
		If DropApos<14 then DropApos=DropApos+1
	end if
  End Sub

  Sub DropIa_Timer()
  Select Case DropIPos
        Case 0: DropIp.z=25
				 If DropIDir = 1 then
					DropIa.TimerEnabled = 0
				 else
			     end if
        Case 1: DropIp.z=21
        Case 2: DropIp.z=17
        Case 3: DropIp.z=13
        Case 4: DropIp.z=9
        Case 5: DropIp.z=5
        Case 6: DropIp.z=1
        Case 7: DropIp.z=0
        Case 8: DropIp.z=-1
        Case 9: DropIp.z=-5
        Case 10: DropIp.z=-9
        Case 11: DropIp.z=-13
        Case 12: DropIp.z=-17
        Case 13: DropIp.z=-21
        Case 14: DropIp.z=-24
				 If DropIDir = 1 then
				 else
					DropIa.TimerEnabled = 0
			     end if


End Select
	If DropIDir = 1 then
		If DropIpos>0 then DropIpos=DropIpos-1
	else
		If DropIpos<14 then DropIpos=DropIpos+1
	end if
  End Sub

  Sub DropRa_Timer()
  Select Case DropRPos
        Case 0: DropRp.z=25
				 If DropRDir = 1 then
					DropRa.TimerEnabled = 0
				 else
			     end if
        Case 1: DropRp.z=21
        Case 2: DropRp.z=17
        Case 3: DropRp.z=13
        Case 4: DropRp.z=9
        Case 5: DropRp.z=5
        Case 6: DropRp.z=1
        Case 7: DropRp.z=0
        Case 8: DropRp.z=-1
        Case 9: DropRp.z=-5
        Case 10: DropRp.z=-9
        Case 11: DropRp.z=-13
        Case 12: DropRp.z=-17
        Case 13: DropRp.z=-21
        Case 14: DropRp.z=-24
				 If DropRDir = 1 then
				 else
					DropRa.TimerEnabled = 0
			     end if


End Select
	If DropRDir = 1 then
		If DropRpos>0 then DropRpos=DropRpos-1
	else
		If DropRpos<14 then DropRpos=DropRpos+1
	end if
  End Sub

  Sub DropSa_Timer()
  Select Case DropSPos
        Case 0: DropSp.z=25
				 If DropSDir = 1 then
					DropSa.TimerEnabled = 0
				 else
			     end if
        Case 1: DropSp.z=21
        Case 2: DropSp.z=17
        Case 3: DropSp.z=13
        Case 4: DropSp.z=9
        Case 5: DropSp.z=5
        Case 6: DropSp.z=1
        Case 7: DropSp.z=0
        Case 8: DropSp.z=-1
        Case 9: DropSp.z=-5
        Case 10: DropSp.z=-9
        Case 11: DropSp.z=-13
        Case 12: DropSp.z=-17
        Case 13: DropSp.z=-21
        Case 14: DropSp.z=-24
				 If DropSDir = 1 then
				 else
					DropSa.TimerEnabled = 0
			     end if


End Select
	If DropSDir = 1 then
		If DropSpos>0 then DropSpos=DropSpos-1
	else
		If DropSpos<14 then DropSpos=DropSpos+1
	end if
  End Sub

  Sub DropHa_Timer()
  Select Case DropHPos
        Case 0: DropHp.z=25
				 If DropHDir = 1 then
					DropHa.TimerEnabled = 0
				 else
			     end if
        Case 1: DropHp.z=21
        Case 2: DropHp.z=17
        Case 3: DropHp.z=13
        Case 4: DropHp.z=9
        Case 5: DropHp.z=5
        Case 6: DropHp.z=1
        Case 7: DropHp.z=0
        Case 8: DropHp.z=-1
        Case 9: DropHp.z=-5
        Case 10: DropHp.z=-9
        Case 11: DropHp.z=-13
        Case 12: DropHp.z=-17
        Case 13: DropHp.z=-21
        Case 14: DropHp.z=-24
				 If DropHDir = 1 then
				 else
					DropHa.TimerEnabled = 0
			     end if


End Select
	If DropHDir = 1 then
		If DropHpos>0 then DropHpos=DropHpos-1
	else
		If DropHpos<14 then DropHpos=DropHpos+1
	end if
  End Sub

  Sub DropOa_Timer()
  Select Case DropOPos
        Case 0: DropOp.z=25
				 If DropODir = 1 then
					DropOa.TimerEnabled = 0
				 else
			     end if
        Case 1: DropOp.z=21
        Case 2: DropOp.z=17
        Case 3: DropOp.z=13
        Case 4: DropOp.z=9
        Case 5: DropOp.z=5
        Case 6: DropOp.z=1
        Case 7: DropOp.z=0
        Case 8: DropOp.z=-1
        Case 9: DropOp.z=-5
        Case 10: DropOp.z=-9
        Case 11: DropOp.z=-13
        Case 12: DropOp.z=-17
        Case 13: DropOp.z=-21
        Case 14: DropOp.z=-24
				 If DropODir = 1 then
				 else
					DropOa.TimerEnabled = 0
			     end if


End Select
	If DropODir = 1 then
		If DropOpos>0 then DropOpos=DropOpos-1
	else
		If DropOpos<14 then DropOpos=DropOpos+1
	end if
  End Sub

  Sub DropWa_Timer()
  Select Case DropWPos
        Case 0: DropWp.z=25
				 If DropWDir = 1 then
					DropWa.TimerEnabled = 0
				 else
			     end if
        Case 1: DropWp.z=21
        Case 2: DropWp.z=17
        Case 3: DropWp.z=13
        Case 4: DropWp.z=9
        Case 5: DropWp.z=5
        Case 6: DropWp.z=1
        Case 7: DropWp.z=0
        Case 8: DropWp.z=-1
        Case 9: DropWp.z=-5
        Case 10: DropWp.z=-9
        Case 11: DropWp.z=-13
        Case 12: DropWp.z=-17
        Case 13: DropWp.z=-21
        Case 14: DropWp.z=-24
				 If DropWDir = 1 then
				 else
					DropWa.TimerEnabled = 0
			     end if


End Select
	If DropWDir = 1 then
		If DropWpos>0 then DropWpos=DropWpos-1
	else
		If DropWpos<14 then DropWpos=DropWpos+1
	end if
  End Sub

'Loop Ramp Speed Checks
	Sub leftloopTrigger_Hit ()
		Dim RandomVelX, RandomVelY, NewVelZ
		RandomVelX=INT(RND*3)
		RandomVelY=INT(RND*3)
		If ABS(ActiveBall.VelY) > (30+RandomVelY) Then
			Kicker6.Enabled=True
			LoopTimer.Enabled=True
		End If
	End Sub
	Sub rightloopTrigger_Hit ()
		Dim RandomVelX, RandomVelY, NewVelZ
		RandomVelX=INT(RND*3)
		RandomVelY=INT(RND*3)
		If ABS(ActiveBall.VelY) > (30+RandomVelY) Then
			Kicker7.Enabled=True
			LoopTimer.Enabled=True
		End If
	End Sub
	Sub LoopTimer_Timer ()
		Kicker7.Enabled=False
		Kicker6.Enabled=False
		Me.Enabled=False
	End Sub

Sub Table1_Exit()
Controller.Stop
End Sub

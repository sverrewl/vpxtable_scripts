Option Explicit
Randomize

' Thalamus 2018-07-20
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' No special SSF tweaks yet.
 
On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

Const cSingleLFlip = 0
Const cSingleRFlip = 0

Const cGameName="Centaur",UseSolenoids=2,UseLamps=0,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff",SFlipperOn="fx_Flipperup",SFlipperOff="fx_Flipperdown"
Const SCoin="coin"


LoadVPM "01000100", "Bally.VBS", 2.24

Dim DesktopMode: DesktopMode = Table1.ShowDT

Dim BallShadows: Ballshadows=1  		'******************set to 1 to turn on Ball shadows
Dim FlipperShadows: FlipperShadows=1  '***********set to 1 to turn on Flipper shadows

If DesktopMode = True Then 'Show Desktop components
Ramp16.visible=1
Ramp15.visible=1
Else
Ramp16.visible=0
Ramp15.visible=0
End if

Dim EnableBallControl
EnableBallControl = false 'Change to true to enable manual ball control (or press C in-game) via the arrow keys and B (boost movement) keys

'Solenoid Call backs
'**********************************************************************************************************
SolCallback(1)	=  "dtM.soldropup"
SolCallback(2)	= "Drop1"
SolCallback(3)	= "Drop2"
SolCallback(4)	= "Drop3"
SolCallback(5)	= "Drop4"
SolCallback(6)  = "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(7)	= "SRelease"
SolCallback(8)	=  "dtLLdrop"
SolCallback(9)	=  "dtLRdrop"
SolCallback(14)	= "OrbsOut"
SolCallback(15)	= "BRelease"
SolCallback(19)	= "vpmNudge.SolGameOn"
'SolCallback(20)	= "Rmag"
'SolCallback(sLLFlipper)="vpmSolFlipper LeftFlipper,nothing,"
'SolCallback(sLRFlipper)="vpmSolFlipper RightFlipper,nothing,"

	SolCallback(sLRFlipper) = "SolRFlipper"
	SolCallback(sLLFlipper) = "SolLFlipper"

	Sub SolLFlipper(Enabled)
		 If Enabled Then
			 PlaySound SoundFX("fx_Flipperup",DOFContactors):LeftFlipper.RotateToEnd
		 Else
			 PlaySound SoundFX("fx_Flipperdown",DOFContactors):LeftFlipper.RotateToStart
		 End If
	  End Sub
	  
	Sub SolRFlipper(Enabled)
		 If Enabled Then
			 PlaySound SoundFX("fx_Flipperup",DOFContactors):RightFlipper.RotateToEnd
		 Else
			 PlaySound SoundFX("fx_Flipperdown",DOFContactors):RightFlipper.RotateToStart
		 End If
	End Sub

'Solenoid Controlled toys
'**********************************************************************************************************		
Sub Drop1(enabled) : If enabled Then: PlaySound SoundFX("DTReset",DOFDropTargets) :sw28.IsDropped = 1 : controller.switch(28) = 1 : drop1shadow.image="blank" : End If : End Sub
Sub Drop2(enabled) : If enabled Then: PlaySound SoundFX("DTReset",DOFDropTargets) :sw27.IsDropped = 1 : controller.switch(27) = 1 : drop2shadow.image="blank" : End If : End Sub
Sub Drop3(enabled) : If enabled Then: PlaySound SoundFX("DTReset",DOFDropTargets) :sw26.IsDropped = 1 : controller.switch(26) = 1 : drop3shadow.image="blank" : End If : End Sub
Sub Drop4(enabled) : If enabled Then: PlaySound SoundFX("DTReset",DOFDropTargets) :sw25.IsDropped = 1 : controller.switch(25) = 1 : drop4shadow.image="blank" : End If : End Sub

Sub SRelease(enabled)
 If enabled Then
	If bsTrough.balls > 0 Then
		If bsTroughO.balls = 4 Then
			bsTrough.SolOut 1
			BallRelease.Createball
			BallRelease.Kick 150,7
 Else
			bsTrough.SolOut 1
			bsTroughO.AddBall 1
		End If
  End If	
 End If	
End Sub
 
Sub BRelease(enabled)
	If enabled and bsTroughO.balls > 0 Then 
		bsTroughO.SolOut 1
		bsTroughE.Addball 1
	End If
End Sub
 
 Sub OrbsOut(enabled)
	If enabled Then
		If bsTroughE.balls > 0 Then
			vpmtimer.addtimer 800,"OrbEject"
		End If
	End If
End Sub
 
 'Orb Ejection and End of Trough Switch
Sub OrbEject(aSw)
	If bsTroughE.balls > 0 Then
		bsTroughE.SolOut 1
		PlaySound "wireramp" ', -1, 0.1, 0.5
		Release2.CreateBall
		Release2.Kick 0,(57+(5*Rnd))
		vpmtimer.pulsesw 33 ' End of Trough Switch
      end If
End Sub 

Sub DTLLdrop(enabled)
	if enabled then
		dtLL.SolDropUp enabled
	bank1shadow.Image="bank2"
	bank2shadow.Image="bank3"
	bank3shadow.Image="bank4"
	end if
End Sub

Sub DTLRdrop(enabled)
	if enabled then
		dtLR.SolDropUp enabled
	drop1shadow.Image="drop1"
	drop2shadow.Image="drop2"
	drop3shadow.Image="drop3"
	drop4shadow.Image="drop4"
	end if
End Sub

Sub Trigger1_hit
	me.timerenabled=1
End sub

Sub Trigger1_timer
	multiguard_prim.objrotx=-8
	me.timerenabled=0
End sub

Sub Trigger2_hit
	stopsound "wireramp"
	trigger1.timerenabled=0
	me.timerenabled=1
end sub

Sub Trigger2_timer
	multiguard_prim.objrotx=0
	me.timerenabled=0
End sub

'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************
 
Dim bsTrough, bsTroughE, bsTroughO, dtLR, dtM, dtLL, RMag
 
Sub Table1_Init
	vpmInit Me
'	On Error Resume Next
	With Controller
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
		.SplashInfoLine = "Centaur (Bally 1981)"
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
	vpmNudge.TiltSwitch = 15
	vpmNudge.Sensitivity = 3
	vpmNudge.TiltObj = Array(Bumper1,Bumper2,LeftSlingshot,RightSlingshot)
 
 	Set bsTrough = New cvpmBallStack
		bsTrough.Initsw 0,8,0,0,0,0,0,0
	    bsTrough.InitExitSnd SoundFX("BallRelease",DOFContactors),SoundFX("solenoid",DOFContactors)
		bsTrough.Balls = 1
		bsTrough.CreateEvents "bsTrough", Drain 

	Set bsTroughE = New cvpmBallStack
		bsTroughE.Initsw 0,0,0,0,0,0,0,0
	    bsTroughE.InitExitSnd SoundFX("BallRelease",DOFContactors),SoundFX("solenoid",DOFContactors)

	Set bsTroughO = New cvpmBallStack
		bsTroughO.Initsw 0,17,0,0,1,2,0,0
		bsTroughO.IsTrough = 1
		bsTroughO.Balls = 4

 	Set dtLR=New cvpmDropTarget
		dtLR.InitDrop Array(sw25,sw26,sw27,sw28),Array(25,26,27,28)
		dtLR.InitSnd SoundFX("DTDrop",DOFDropTargets),SoundFX("DTReset",DOFContactors)

 	Set dtM=New cvpmDropTarget
		dtM.InitDrop Array(sw29,sw30,sw31,sw32),Array(29,30,31,32)
		dtM.InitSnd SoundFX("DTDrop",DOFDropTargets),SoundFX("DTReset",DOFContactors)

  	Set dtLL=New cvpmDropTarget
		dtLL.InitDrop Array(sw41,sw42,sw43,sw44),Array(41,42,43,44)
		dtLL.InitSnd SoundFX("DTDrop",DOFDropTargets),SoundFX("DTReset",DOFContactors)

 	Set RMag = New cvpmMagnet
		with Rmag
	    .initMagnet MagTrigger, 18
		.GrabCenter=False
'	    .X = 836
'	    .Y = 318
'        .Size = 60
		.solenoid = 20
		.CreateEvents "Rmag"
		end with

'    Set mTLMag= New cvpmMagnet
'    With mTLMag
'        .InitMagnet Magnet1, 16  
'        .GrabCenter = False
'        .solenoid=51   ' LE
''       .solenoid=32   ' Pro
'        .CreateEvents "mTLMag"
'    End With

	CapKicker.CreateBall
	CapKicker.Kick 180,2
End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)
'if keycode=rightmagnasave then					'***********used for testing power orb ramp
'    PlaySound "wireramp" ', -1, 0.1, 0.5
'  	Release2.CreateBall
'	Release2.Kick 0,(57+(5*Rnd))
'end if
	if keycode = RightFlipperKey then Controller.Switch(21) = false
	if keycode = LeftFlipperKey then Controller.Switch(21) = false
	If KeyDownHandler(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Pullback:playsound"plungerpull"

   ' Manual Ball Control
	If keycode = 46 Then	 				' C Key
		If EnableBallControl = 1 Then
			EnableBallControl = 0
		Else
			EnableBallControl = 1
		End If
	End If
    If EnableBallControl = 1 Then
		If keycode = 48 Then 				' B Key
			If BCboost = 1 Then
				BCboost = BCboostmulti
			Else
				BCboost = 1
			End If
		End If
		If keycode = 203 Then BCleft = 1	' Left Arrow
		If keycode = 200 Then BCup = 1		' Up Arrow
		If keycode = 208 Then BCdown = 1	' Down Arrow
		If keycode = 205 Then BCright = 1	' Right Arrow
	End If

End Sub

Sub Table1_KeyUp(ByVal KeyCode)
	if keycode = RightFlipperKey then Controller.Switch(21) = true
	if keycode = LeftFlipperKey then Controller.Switch(21) = true
	If KeyUpHandler(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Fire:PlaySound"plunger"

    'Manual Ball Control
	If EnableBallControl = 1 Then
		If keycode = 203 Then BCleft = 0	' Left Arrow
		If keycode = 200 Then BCup = 0		' Up Arrow
		If keycode = 208 Then BCdown = 0	' Down Arrow
		If keycode = 205 Then BCright = 0	' Right Arrow
	End If

End Sub
'**********************************************************************************************************
'**********************************************************************************************************

 'switches
'**********************************************************************************************************
'Multi Ball kicker 
 Sub sw33_Hit   : Controller.Switch(33) = 1: End Sub
 Sub sw33_UnHit : Controller.Switch(33) = 0 : End Sub

 ' Drain kickers
 Sub Drain_Hit:playsound"drain":End Sub

'Drop Targets
 Sub Sw25_Dropped:dtLR.Hit 1 : drop4shadow.image="blank" : End Sub  
 Sub Sw26_Dropped:dtLR.Hit 2 : drop3shadow.image="blank" : End Sub  
 Sub Sw27_Dropped:dtLR.Hit 3 : drop2shadow.image="blank" : End Sub
 Sub Sw28_Dropped:dtLR.Hit 4 : drop1shadow.image="blank" : End Sub

 Sub Sw29_Dropped:dtM.Hit 1 :End Sub
 Sub Sw30_Dropped:dtM.Hit 2 :End Sub
 Sub Sw31_Dropped:dtM.Hit 3 :End Sub 
 Sub Sw32_Dropped:dtM.Hit 4 :End Sub

 Sub Sw41_Dropped:dtLL.Hit 1 : bank1shadow.image="blank" : End Sub  
 Sub Sw42_Dropped:dtLL.Hit 2 : bank2shadow.image="blank" : End Sub  
 Sub Sw43_Dropped:dtLL.Hit 3 : bank3shadow.image="blank" : End Sub  
 Sub Sw44_Dropped:dtLL.Hit 4 : bank4shadow.image="blank" : End Sub  

'Stand Up Targets
 Sub Tsw12a_Hit: PrimStandupTgtHit 12, Tsw12a, PrimT12a: DOF 102, DOFPulse: End Sub
 Sub Tsw12a_Timer: PrimStandupTgtMove 12, Tsw12a, PrimT12a: End Sub
' Sub sw12a_Hit:vpmTimer.PulseSW 12:DOF 102, DOFPulse:End Sub
 Sub sw12b_Hit:vpmTimer.PulseSW 12:DOF 103, DOFPulse:End Sub
 Sub sw12c_Hit:vpmTimer.PulseSW 12:DOF 103, DOFPulse:End Sub
 Sub sw19_Hit:vpmTimer.PulseSW 19:End Sub
 Sub sw20_Hit:vpmTimer.PulseSW 20:End Sub
 Sub sw22_Hit:vpmTimer.PulseSW 22:End Sub
 Sub sw24_Hit:vpmTimer.PulseSW 24:End Sub

'Bumpers
  Sub Bumper1_Hit:vpmTimer.PulseSw 39 : playsound SoundFX("fx_bumper1",DOFContactors): End Sub
  Sub Bumper2_Hit:vpmTimer.PulseSw 40 : playsound SoundFX("fx_bumper1",DOFContactors): End Sub

 'Star Rollovers
 Sub sw12_Hit   : Controller.Switch(12) = 1 : playsound"rollover" : DOF 101, DOFOn : End Sub
 Sub sw12_UnHit : Controller.Switch(12) = 0 : DOF 101, DOFOff : End Sub
 Sub sw18_Hit   : Controller.Switch(18) = 1 : playsound"rollover" : End Sub
 Sub sw18_UnHit : Controller.Switch(18) = 0 : End Sub
 Sub sw18a_Hit  : Controller.Switch(18) = 1 : playsound"rollover" : End Sub
 Sub sw18a_UnHit: Controller.Switch(18) = 0 : End Sub

 'Wire Triggers
 Sub sw3_Hit   : Controller.Switch(3) = 1 : playsound"rollover" : End Sub
 Sub sw3_UnHit : Controller.Switch(3) = 0 : End Sub
 Sub sw4_Hit   : Controller.Switch(4) = 1 : playsound"rollover" : End Sub
 Sub sw4_UnHit : Controller.Switch(4) = 0 : End Sub
 Sub sw5_Hit   : Controller.Switch(5) = 1 : playsound"rollover" : End Sub
 Sub sw5_UnHit : Controller.Switch(5) = 0 : End Sub
 Sub sw45_Hit  : Controller.Switch(45) = 1 : playsound"rollover" : End Sub
 Sub sw45_UnHit: Controller.Switch(45) = 0 : End Sub
 Sub sw46_Hit  : Controller.Switch(46) = 1 : playsound"rollover" : End Sub
 Sub sw46_UnHit: Controller.Switch(46) = 0 : End Sub
 Sub sw47_Hit  : Controller.Switch(47) = 1 : playsound"rollover" : End Sub
 Sub sw47_UnHit: Controller.Switch(47) = 0 : End Sub
 Sub sw48_Hit  : Controller.Switch(48) = 1 : playsound"rollover" : End Sub
 Sub sw48_UnHit: Controller.Switch(48) = 0 : End Sub
 
'GI Lights single Lamp source to control a collection
 Dim N1,O1, N2,O2, N3,O3, GION
 N1=0:O1=0:N2=0:O2=0:N3=0:O3=0:
 Set LampCallback=GetRef("UpdateMultipleLamps")
 Sub UpdateMultipleLamps
 	N1=Controller.Lamp(1) 'upper PF
  	N2=Controller.Lamp(66) 'Right lower PF
 	N3=Controller.Lamp(82) 'Left lower PF
 	If N1<>O1 Then
 		If N1 Then 
			dim xx
			GION=1
			For each xx in GI1:xx.State = 1: Next
			metalright_prm.image="rightmetal"
			plasticsedgesGI1_prim.image="edgegi1"
			outerback_prim.image="outerback"
			woodlrightupper_prim.image="woodtopright"
			reargatebracket_prim.image="backgate"
			bumperring1_prim.image="bumperring1GION"
			bumperring2_prim.image="bumperring2GION"
			brackets_prim.image="gatebrackets"
			leftgatebracket_prim.image="leftgate"
			If l115.state=1 Then
				metalouter_prm.image="outermetalLAIRON"
			Else
				metalouter_prm.image="outermetal"
			End If
			If l100.state=0 Then
				metalrear_prm.image="backmetalSTAROFF"
			Else
				metalrear_prm.image="backmetal"
			End If
		Else
			For each xx in GI1:xx.State = 0: Next
			GION=0
			metalright_prm.image="rightmetalGIOFF"
			plasticsedgesGI1_prim.image="edgegi1GIOFF"
			outerback_prim.image="outerbackGIOFF"
			woodlrightupper_prim.image="woodtoprightGIOFF"
			reargatebracket_prim.image="backgateGIOFF"
			bumperring1_prim.image="bumperring1GIOFF"
			bumperring2_prim.image="bumperring2GIOFF"
			brackets_prim.image="gatebracketsGIOFF"
			leftgatebracket_prim.image="leftgateGIOFF"
			If l115.state=1 Then
				metalouter_prm.image="outermetalGIOFFLAIRON"
			Else
				metalouter_prm.image="outermetalGIOFF"
			End If
			If l100.state=0 Then
				metalrear_prm.image="backmetalGIOFFSTAROFF"
			Else
				metalrear_prm.image="backmetalGIOFF"
			End If
		End If
 	O1=N1
	End If
 	If N2<>O2 Then
 		If N2 Then 
			dim mm
			For each mm in GI2:mm.State = 1: Next
			plasticsedgesGI2_prim.image="edgegi2"
			outerright_prim.image="outerright"
			woodlrightlower_prim.image="rightwood"
			plasticsedgeslowerGI2_prim.image="edgelowerGI2"
'			flasher1.visible=1
'			flasher2.visible=0
		Else
			For each mm in GI2:mm.State = 0: Next
			plasticsedgesGI2_prim.image="edgegi2GIOFF"
			outerright_prim.image="outerrightGIOFF"
			woodlrightlower_prim.image="rightwoodGIOFF"
			plasticsedgeslowerGI2_prim.image="edgelowerGI2GIOFF"
'			flasher1.visible=0
'			flasher2.visible=1
		End If
 	O2=N2
	End If
 	If N3<>O3 Then
 		If N3 Then 
			dim nn
			For each nn in GI3:nn.State = 1: Next
			plasticsedgesGI3_prim.image="edgegi3"
			woodleftlow_prim.image="leftwoodlow"
			outerleft_prim.image="outerleft"
			plasticsedgeslowerGI3_prim.image="edgelowerGI3"
		Else
			For each nn in GI3:nn.State = 0: Next
			plasticsedgesGI3_prim.image="edgegi3GIOFF"
			woodleftlow_prim.image="leftwoodlowGIOFF"
			outerleft_prim.image="outerleftGIOFF"
			plasticsedgeslowerGI3_prim.image="edgelowerGI3"
		End If
 	O3=N3
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

InitLamps()             ' turn off the lights and flashers and reset them to the default parameters
LampTimer.Interval = 5 'lamp fading speed
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
	NFadeLm 2, l2
   	NFadeLm 2, l2a
	NFadeLm 2, l2z
	NFadeLm 3, l3
   	NFadeLm 3, l3a
	NFadeLm 3, l3z
	NFadeLm 4, l4
   	NFadeLm 4, l4a
	NFadeLm 4, l4z
	NFadeLm 5, l5
   	NFadeLm 5, l5a
	NFadeLm 5, l5z
	NFadeLm 6, l6
   	NFadeLm 6, l6a
   	NFadeLm 6, l6b
	NFadeLm 6, l6z
	NFadeLm 7, l7
   	NFadeLm 7, l7a
	NFadeLm 7, l7z
	NFadeLm 8, l8
   	NFadeLm 8, l8a
	NFadeLm 8, l8z
	NFadeLm 9, l9
   	NFadeLm 9, l9a
	NFadeLm 9, l9z
   	NFadeLm 10, l10
   	NFadeLm 10, l10a
   	NFadeLm 10, l10z
   	NFadeLm 12, l12
   	NFadeLm 12, l12a
   	NFadeLm 12, l12z
   	NFadeLm 14, l14
   	NFadeLm 14, l14a
   	NFadeLm 14, l14z
   	NFadeLm 15, l15
   	NFadeLm 15, l15a
   	NFadeLm 15, l15b
   	NFadeLm 15, l15z
   	NFadeLm 17, l17
   	NFadeLm 17, l17a
   	NFadeLm 17, l17y
   	NFadeLm 17, l17z
   	NFadeLm 18, l18
   	NFadeLm 18, l18a
   	NFadeLm 18, l18z
   	NFadeLm 19, l19
   	NFadeLm 19, l19a
   	NFadeLm 19, l19z
   	NFadeLm 20, l20
   	NFadeLm 20, l20a
   	NFadeLm 20, l20z
   	NFadeLm 21, l21
   	NFadeLm 21, l21a
   	NFadeLm 21, l21z
   	NFadeLm 22, l22
   	NFadeLm 22, l22a
   	NFadeLm 22, l22z
   	NFadeLm 23, l23
   	NFadeLm 23, l23a
   	NFadeLm 23, l23z
   	NFadeLm 24, l24
   	NFadeLm 24, l24a
   	NFadeLm 24, l24z
   	NFadeLm 25, l25
   	NFadeLm 25, l25a
   	NFadeLm 25, l25z
   	NFadeLm 26, l26
   	NFadeLm 26, l26a
   	NFadeLm 26, l26z
   	NFadeLm 28, l28
   	NFadeLm 28, l28a
   	NFadeLm 28, l28z
   	NFadeLm 30, l30
   	NFadeLm 30, l30a
   	NFadeLm 30, l30z
   	NFadeLm 31, l31
   	NFadeLm 31, l31a
   	NFadeLm 31, l31b
   	NFadeLm 31, l31z
   	NFadeLm 33, l33
   	NFadeLm 33, l33a
   	NFadeLm 33, l33z
   	NFadeLm 34, l34
   	NFadeLm 34, l34a
   	NFadeLm 34, l34z
   	NFadeLm 35, l35
   	NFadeLm 35, l35a
   	NFadeLm 35, l35z
	NFadeLm 36, l36
	NFadeLm 36, l36a
	NFadeLm 36, l36b
	NFadeLm 36, l36c
	NFadeLm 36, l36z
   	NFadeLm 37, l37
   	NFadeLm 37, l37a
   	NFadeLm 37, l37z
   	NFadeLm 38, l38
   	NFadeLm 38, l38a
   	NFadeLm 38, l38b
   	NFadeLm 38, l38z
   	NFadeLm 39, l39
   	NFadeLm 39, l39a
   	NFadeLm 39, l39z
   	NFadeLm 40, l40
   	NFadeLm 40, l40a
   	NFadeLm 40, l40z
   	NFadeLm 41, l41
   	NFadeLm 41, l41a
   	NFadeLm 41, l41z
   	NFadeLm 42, l42
   	NFadeLm 42, l42a
   	NFadeLm 42, l42z
   	NFadeLm 44, l44
   	NFadeLm 44, l44a
   	NFadeLm 44, l44z
   	NFadeLm 46, l46
   	NFadeLm 46, l46a
   	NFadeLm 46, l46z
   	NFadeLm 47, l47
   	NFadeLm 47, l47a
   	NFadeLm 47, l47b
   	NFadeLm 47, l47z
   	NFadeLm 49, l49
   	NFadeLm 49, l49a
   	NFadeLm 49, l49z
   	NFadeLm 50, l50
   	NFadeLm 50, l50a
   	NFadeLm 50, l50z
   	NFadeLm 51, l51
   	NFadeLm 51, l51a
   	NFadeLm 51, l51z
	NFadeLm 52, l52
	NFadeLm 52, l52a
	NFadeLm 52, l52b
	NFadeLm 52, l52z
   	NFadeLm 53, l53
   	NFadeLm 53, l53a
   	NFadeLm 53, l53z
   	NFadeLm 54, l54
   	NFadeLm 54, l54a
   	NFadeLm 54, l54z
   	NFadeLm 55, l55
   	NFadeLm 55, l55a
   	NFadeLm 55, l55z
   	NFadeLm 56, l56
   	NFadeLm 56, l56a
   	NFadeLm 56, l56z
	NFadeLm 57, l57
	NFadeLm 57, l57a
	NFadeLm 57, l57z
   	NFadeLm 58, l58
   	NFadeLm 58, l58a
   	NFadeLm 58, l58z
   	NFadeLm 59, l59
   	NFadeLm 60, l60
   	NFadeLm 60, l60a
   	NFadeLm 60, l60z
   	NFadeLm 62, l62
   	NFadeLm 62, l62a
   	NFadeLm 62, l62z
   	NFadeLm 63, l63
   	NFadeLm 63, l63a
	NFadeLm 63, l63b
   	NFadeLm 63, l63z
   	NFadeLm 65, l65
   	NFadeLm 65, l65a
   	NFadeLm 65, l65z
	Flash 66, Flasher1
	SetLamp 166, 1-LampState(66)
	Flash 166, Flasher2
	Nfadelm 67, l67
	Nfadelm 67, l67a
	Nfadelm 67, l67b
	Nfadelm 67, l67c
	If l67.state=1 Then
		If l83.state=0 Then
			If l99.state=0 Then
				If sw41.isdropped=0 Then
					bank1shadow.image="bank1l1"
					bank2shadow.image="bank2l1"
					bank3shadow.image="bank3l1"
				Else
					If sw42.isdropped=0 Then
						bank2shadow.image="bank2l1"
						bank3shadow.image="bank3l1"
					Else
						If sw43.isdropped=0 Then
							bank3shadow.image="bank3l1"
						Else
							bank3shadow.image="blank"
						End If
						bank2shadow.image="blank"
					End If
					bank1shadow.image="blank"
				End If
				plasticsedgesLair_prim.image="edgeslairL67"
				bankshadow.image="bankL1"
				bank4shadow.image="blank"
			Else
				If sw42.isdropped=0 Then
					bank2shadow.image="bank2"
					bank3shadow.image="bank3"
				Else
					If sw43.isdropped=0 Then
						bank3shadow.image="bank3"
					Else
						bank3shadow.image="blank"
					End If
					bank2shadow.image="blank"
				End If
				bank1shadow.image="blank"
				bank4shadow.image="bank4"
				bankshadow.image="bankON"
				plasticsedgesLair_prim.image="edgeslair"
			End If
		End If
	End If
   	NFadeLm 68, l68
   	NFadeLm 68, l68a
   	NFadeLm 68, l68z
   	NFadeLm 81, l81
   	NFadeLm 81, l81a
   	NFadeLm 81, l81z
	Nfadelm 83, l83
	Nfadelm 83, l83a
	Nfadelm 83, l83b
	Nfadelm 83, l83c
	If l83.state=1 Then
		If l67.state=0 Then
			If l115.state=0 Then
				If sw41.isdropped=0 Then
					bank1shadow.image="bank1l2"
					bank2shadow.image="bank2l2"
					bank3shadow.image="bank3l2"
				Else
					If sw42.isdropped=0 Then
						bank2shadow.image="bank2l2"
						bank3shadow.image="bank3l2"
					Else
						If sw43.isdropped=0 Then
							bank3shadow.image="bank3l2"
						Else
							bank3shadow.image="blank"
						End If
						bank2shadow.image="blank"
					End If
					bank1shadow.image="blank"
				End If
				plasticsedgesLair_prim.image="edgeslairL83"
				bankshadow.image="bankL2"
				bank4shadow.image="blank"
			Else
				If sw42.isdropped=0 Then
					bank2shadow.image="bank2"
					bank3shadow.image="bank3"
				Else
					If sw43.isdropped=0 Then
						bank3shadow.image="bank3"
					Else
						bank3shadow.image="blank"
					End If
					bank2shadow.image="blank"
				End If
				bank4shadow.image="bank4"
				bankshadow.image="bankON"
				plasticsedgesLair_prim.image="edgeslair"
			End If
		End If
		bank1shadow.image="blank"
	End If
   	NFadeLm 84, l84
   	NFadeLm 84, l84a
   	NFadeLm 84, l84z
	If l84.state=1 Then
		FadeObjm 84, metalleft_prm, "leftmetal", "leftmetal", "leftmetal", "leftmetalstaroff"
	Else
		metalleft_prm.image="leftmetalstaroff"
	End if
	NFadeLm 97, l97
   	NFadeLm 97, l97a
   	NFadeLm 97, l97z
	NFadeLM 98, L98
			If l98.state=1 Then
				If	bumperring2_prim.image="bumperring2GION" Then
					bumpercap2_prim.image="bumpercap2ONGION"
				Else
					bumpercap2_prim.image="bumpercap2ONGIOFF"
				End If
			End If
			If l98.state=0 Then
				If	bumperring2_prim.image="bumperring2GION" Then
					bumpercap2_prim.image="bumpercap2OFFGION"
				Else
					bumpercap2_prim.image="bumpercap2OFFGIOFF"
				End If
			End If
	Nfadelm 99, l99
	Nfadelm 99, l99a
	Nfadelm 99, l99b
	Nfadelm 99, l99c
	If l99.state=1 Then
		If l83.state=0 Then
			If l67.state=0 Then
				If sw42.isdropped=0 Then
					bank2shadow.image="bank2l3"
					bank3shadow.image="bank3l3"
				Else
					If sw43.isdropped=0 Then
						bank3shadow.image="bank3l3"
					Else
						bank3shadow.image="blank"
					End If
					bank2shadow.image="blank"
					bank1shadow.image="blank"
				End If
				plasticsedgesLair_prim.image="edgeslairL99"
				bankshadow.image="bankL3"
				bank1shadow.image="blank"
				bank4shadow.image="blank"
			Else
				If sw42.isdropped=0 Then
					bank2shadow.image="bank2"
					bank3shadow.image="bank3"
				Else
					If sw43.isdropped=0 Then
						bank3shadow.image="bank3"
					Else
						bank3shadow.image="blank"
					End If
					bank2shadow.image="blank"
				End If
				bankshadow.image="bankON"
				bank4shadow.image="bank4"
				plasticsedgesLair_prim.image="edgeslair"
			End If
		End If
		bank1shadow.image="blank"
	End If
	NFadeLm 100, l100
   	NFadeLm 100, l100a
	NFadeLm 100, l100z
	NFadeLM 114, L114
			If l114.state=1 Then
				If	bumperring1_prim.image="bumperring1GION" Then
					bumpercap1_prim.image="bumpercap1ONGION"
				Else
					bumpercap1_prim.image="bumpercap1ONGIOFF"
				End If
			End If
			If l114.state=0 Then
				If	bumperring1_prim.image="bumperring1GION" Then
					bumpercap1_prim.image="bumpercap1OFFGION"
				Else
					bumpercap1_prim.image="bumpercap1OFFGIOFF"
				End If
			End If

	Nfadelm 115, l115
	Nfadelm 115, l115a
	Nfadelm 115, l115b
	Nfadelm 115, l115c
	If l115.state=1 Then
		If GION=0 then metalouter_prm.image="outermetalGIOFFLairon"
		If l99.state=0 Then
			If l83.state=0 Then
				If sw43.isdropped=0 Then
					bank3shadow.image="bank3l3"
					bank4shadow.image="bank4l4"
				Else
					If sw44.isdropped=0 Then
						bank4shadow.image="bank4l4"
					Else
						bank4shadow.image="blank"
					End If
					bank3shadow.image="blank"
				End If
				plasticsedgesLair_prim.image="edgeslairL115"
				bankshadow.image="bankL4"
			Else
				If sw43.isdropped=0 Then
					bank3shadow.image="bank3"
					bank4shadow.image="bank4"
				Else
					If sw44.isdropped=0 Then
						bank4shadow.image="bank4"
					Else
						bank4shadow.image="blank"
					End If
					bank3shadow.image="blank"
				End If
				bankshadow.image="bankON"
				bank4shadow.image="bank4"
				plasticsedgesLair_prim.image="edgeslair"
			End If
		Else
		bank1shadow.image="blank"
		End If
	Else
	If GION=0 then metalouter_prm.image="outermetalGIOFF"
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

Sub FadePrim(nr, pri, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:pri.image = d:FadingLevel(nr) = 0
        Case 3:pri.image = c:FadingLevel(nr) = 2
        Case 4:pri.image = b:FadingLevel(nr) = 3
        Case 5:pri.image = a:FadingLevel(nr) = 1
    End Select
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

 '****************************************
'BackGlass Lights
'Set Lights(43)=Light43 'Warning
'Set Lights(29)=Light29 'High Score
'Set Lights(63)=Light63 'TILT
'Set Lights(45)=Light45 'GameOver
'Set Lights(11)=Light11 'shoot Again
'Set Lights(13)=Light13 'ball in play
'Set Lights(27)=Light27 'match

'**********************************************************************************************************
'Digital Display
'**********************************************************************************************************
Dim Digits(32)

' 1st Player
Digits(0) = Array(LED10,LED11,LED12,LED13,LED14,LED15,LED16)
Digits(1) = Array(LED20,LED21,LED22,LED23,LED24,LED25,LED26)
Digits(2) = Array(LED30,LED31,LED32,LED33,LED34,LED35,LED36)
Digits(3) = Array(LED40,LED41,LED42,LED43,LED44,LED45,LED46)
Digits(4) = Array(LED50,LED51,LED52,LED53,LED54,LED55,LED56)
Digits(5) = Array(LED60,LED61,LED62,LED63,LED64,LED65,LED66)
Digits(6) = Array(LED70,LED71,LED72,LED73,LED74,LED75,LED76)

' 2nd Player
Digits(7) = Array(LED80,LED81,LED82,LED83,LED84,LED85,LED86)
Digits(8) = Array(LED90,LED91,LED92,LED93,LED94,LED95,LED96)
Digits(9) = Array(LED100,LED101,LED102,LED103,LED104,LED105,LED106)
Digits(10) = Array(LED110,LED111,LED112,LED113,LED114,LED115,LED116)
Digits(11) = Array(LED120,LED121,LED122,LED123,LED124,LED125,LED126)
Digits(12) = Array(LED130,LED131,LED132,LED133,LED134,LED135,LED136)
Digits(13) = Array(LED140,LED141,LED142,LED143,LED144,LED145,LED146)

' 3rd Player
Digits(14) = Array(LED150,LED151,LED152,LED153,LED154,LED155,LED156)
Digits(15) = Array(LED160,LED161,LED162,LED163,LED164,LED165,LED166)
Digits(16) = Array(LED170,LED171,LED172,LED173,LED174,LED175,LED176)
Digits(17) = Array(LED180,LED181,LED182,LED183,LED184,LED185,LED186)
Digits(18) = Array(LED190,LED191,LED192,LED193,LED194,LED195,LED196)
Digits(19) = Array(LED200,LED201,LED202,LED203,LED204,LED205,LED206)
Digits(20) = Array(LED210,LED211,LED212,LED213,LED214,LED215,LED216)

' 4th Player
Digits(21) = Array(LED220,LED221,LED222,LED223,LED224,LED225,LED226)
Digits(22) = Array(LED230,LED231,LED232,LED233,LED234,LED235,LED236)
Digits(23) = Array(LED240,LED241,LED242,LED243,LED244,LED245,LED246)
Digits(24) = Array(LED250,LED251,LED252,LED253,LED254,LED255,LED256)
Digits(25) = Array(LED260,LED261,LED262,LED263,LED264,LED265,LED266)
Digits(26) = Array(LED270,LED271,LED272,LED273,LED274,LED275,LED276)
Digits(27) = Array(LED280,LED281,LED282,LED283,LED284,LED285,LED286)

' Credits
Digits(28) = Array(LED4,LED2,LED6,LED7,LED5,LED1,LED3)
Digits(29) = Array(LED18,LED9,LED27,LED28,LED19,LED8,LED17)
' Balls
Digits(30) = Array(LED39,LED37,LED48,LED49,LED47,LED29,LED38)
Digits(31) = Array(LED67,LED58,LED69,LED77,LED68,LED57,LED59)

Sub DisplayTimer_Timer
	Dim ChgLED,ii,num,chg,stat,obj
	ChgLed = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
If Not IsEmpty(ChgLED) Then
		If not B2SOn then
		For ii = 0 To UBound(chgLED)
			num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
			if (num < 32) then
				For Each obj In Digits(num)
					If chg And 1 Then obj.State = stat And 1 
					chg = chg\2 : stat = stat\2
				Next
			else

			end if
		next
		end if
end if
End Sub


'**********************************************************************************************************
'**********************************************************************************************************
'Bally Centaur
'added by Gaston
'corrected by Inkochnito
Sub editDips
	Dim vpmDips : Set vpmDips = New cvpmDips
	With vpmDips
		.AddForm 315,220,"Centaur - DIP switches"
		.AddFrame 2,5,88,"Balls in play",&HC0000000,Array("2 balls",&HC0000000,"3 balls",0,"4 balls",&H80000000,"5 balls",&H40000000)'dip 31&32
		.AddFrame 2,85,88,"Tilt settings",&H00400000, Array("Ball Tilt",&H00400000,"Game Tilt",0)'dip 23
		.AddFrame 2,137,88,"Max. credits",&H03000000,Array("10 credits",0,"15 credits",&H01000000,"25 credits",&H02000000,"40 credits",&H03000000)'dip 25&26
		.AddFrame 104,5,88,"Guardian release",&H000F8F1F,Array("Each time",&H58001,"Once per ball",&H50001)'dip 16
		.AddFrame 104,52,88,"Drop targets",&H00000080,Array("Store dropped",&H00000080,"Reset always",0)'dip 8
		.AddFrame 104,99,88,"Drop tgt. bonus",&H00002000,Array("Stored",&H00002000,"Not stored",0)'dip 14
		.AddChk 105,150,95,Array("Store orbs",&H00004000)'dip 15
		.AddChk 105,166,95,Array("Store multiplier",&H00000020)'dip 6
	    .AddChk 105,182,95,Array("Store bonus",&H00000040)'dip 7
		.AddChk 105,198,200,Array("Fires ball every 15 min. at Game Over",&H20000000)'dip 30
		.AddFrame 206,5,88,"Last ball release",&H00100000,Array("On",&H00100000,"Off",0)'dip 21
		.AddFrame 206,52,88,"Right release",&H00800000,Array("Stays on",&H00800000,"Off",0)'dip 24
		.AddFrame 206,99,88,"Initial orbs lit",&H00200000,Array("One",&H00200000,"None",0)'dip 22
		.AddChk 207,150,95,Array("Match feature",&H08000000)'dip 28
		.AddChk 207,166,95,Array("Credits display",&H04000000)'dip 27
		.AddChk 207,182,100,Array("Unlimited replays",&H10000000)'dip 29
		.AddLabel 2,220,308,20,"Set selftest position 16,17,18 and 19 to 03 for the best gameplay."
		.AddLabel 2,240,300,20,"After hitting OK, press F3 to reset game with new settings."
		.ViewDips
	End With
End Sub
Set vpmShowDips = GetRef("editDips")
'**********************************************************************************************************
'**********************************************************************************************************
'**********************************************************************************************************
'**********************************************************************************************************


'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep, LAstep, LBstep, LB2step, LCstep, LC2step, RCstep, RAStep, RA2Step, RBStep, RB2Step

Sub RightSlingShot_Slingshot
	vpmTimer.PulseSw 37
    PlaySound SoundFX("left_slingshot",DOFContactors), 0, 1, 0.05, 0.05
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
	vpmTimer.PulseSw 38
    PlaySound SoundFX("right_slingshot",DOFContactors),0,1,-0.05,0.05
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

Sub WallLA_Hit
    RubberLA.Visible = 0
    RubberLA1.Visible = 1
    LAstep = 0
    WallLA.TimerEnabled = 1
End Sub
 
Sub WallLA_Timer
    Select Case LAstep
        Case 3:RubberLA1.Visible = 0:RubberLA2.Visible = 1
        Case 4:RubberLA2.Visible = 0:RubberLA.Visible = 1:WallLA.TimerEnabled = 0:
    End Select
    LAstep = LAstep + 1
End Sub
'
Sub sw34a_Hit
    vpmTimer.PulseSw 34
	playsound"target"
    RubberLB.Visible = 0
    RubberLB1.Visible = 1
    LBstep = 0
    sw34a.TimerEnabled = 1
End Sub
 
Sub sw34a_Timer
    Select Case LBstep
        Case 3:RubberLB1.Visible = 0:RubberLB2.Visible = 1
        Case 4:RubberLB2.Visible = 0:RubberLB.Visible = 1:sw34a.TimerEnabled = 0:
    End Select
    LBstep = LBstep + 1
End Sub

Sub sw34b_Hit
    vpmTimer.PulseSw 34
	playsound SoundFX("target",DOFTargets)
    RubberLB.Visible = 0
    RubberLB3.Visible = 1
    LB2step = 0
    sw34b.TimerEnabled = 1
End Sub
 
Sub sw34b_Timer
    Select Case LB2step
        Case 3:RubberLB3.Visible = 0:RubberLB4.Visible = 1
        Case 4:RubberLB4.Visible = 0:RubberLB.Visible = 1:sw34b.TimerEnabled = 0:
    End Select
    LB2step = LB2step + 1
End Sub

Sub WallLC_Hit
    RubberLC.Visible = 0
    RubberLC1.Visible = 1
    LCstep = 0
    WallLC.TimerEnabled = 1
End Sub
 
Sub WallLC_Timer
    Select Case LCstep
        Case 3:RubberLC1.Visible = 0:RubberLC2.Visible = 1
        Case 4:RubberLC2.Visible = 0:RubberLC.Visible = 1:WallLC.TimerEnabled = 0:
    End Select
    LCstep = LCstep + 1
End Sub

Sub Wall1RA_Hit
    RubberRA.Visible = 0
    RubberRA1.Visible = 1
    RAstep = 0
    Wall1RA.TimerEnabled = 1
End Sub
 
Sub Wall1RA_Timer
    Select Case RAstep
        Case 3:RubberRA1.Visible = 0:RubberRA2.Visible = 1
        Case 4:RubberRA2.Visible = 0:RubberRA.Visible = 1:Wall1RA.TimerEnabled = 0:
    End Select
    RAstep = RAstep + 1
End Sub

Sub Wall2RA_Hit
    vpmTimer.PulseSw 34
	playsound SoundFX("target",DOFTargets)
    RubberRA.Visible = 0
    RubberRA2.Visible = 1
    RA2step = 0
    Wall2RA.TimerEnabled = 1
End Sub
 
Sub Wall2RA_Timer
    Select Case RAstep
        Case 3:RubberRA2.Visible = 0:RubberRA1.Visible = 1
        Case 4:RubberRA1.Visible = 0:RubberRA.Visible = 1:Wall2RA.TimerEnabled = 0:
    End Select
    RA2step = RA2step + 1
End Sub

Sub sw34c_Hit
    vpmTimer.PulseSw 34
	playsound SoundFX("target",DOFTargets)
    RubberRB.Visible = 0
    RubberRB1.Visible = 1
    RBstep = 0
    sw34c.TimerEnabled = 1
End Sub
 
Sub sw34c_Timer
    Select Case RBstep
        Case 3:RubberRB1.Visible = 0:RubberRB2.Visible = 1
        Case 4:RubberRB2.Visible = 0:RubberRB.Visible = 1:sw34c.TimerEnabled = 0:
    End Select
    RBstep = RBstep + 1
End Sub

Sub sw34d_Hit
    vpmTimer.PulseSw 34
	playsound SoundFX("target",DOFTargets)
    RubberRB.Visible = 0
    RubberRB2.Visible = 1
    RB2step = 0
    sw34d.TimerEnabled = 1
End Sub
 
Sub sw34d_Timer
    Select Case RB2step
        Case 3:RubberRB2.Visible = 0:RubberRB1.Visible = 1
        Case 4:RubberRB1.Visible = 0:RubberRB.Visible = 1:sw34d.TimerEnabled = 0:
    End Select
    RB2step = RB2step + 1
End Sub

Sub WallRC_Hit
    RubberRC.Visible = 0
    RubberRC1.Visible = 1
    RCstep = 0
    WallRC.TimerEnabled = 1
End Sub
 
Sub WallRC_Timer
    Select Case RCstep
        Case 3:RubberRC1.Visible = 0:RubberRC2.Visible = 1
        Case 4:RubberRC2.Visible = 0:RubberRC.Visible = 1:WallRC.TimerEnabled = 0:
    End Select
    RCstep = RCstep + 1
End Sub

Const WallPrefix 		= "T" 'Change this based on your naming convention
Const PrimitivePrefix 	= "PrimT"'Change this based on your naming convention
Const PrimitiveBumperPrefix = "BumperRing" 'Change this based on your naming convention
Dim primCnt(120), primDir(120), primBmprDir(120)

'****************************************************************************
'***** Primitive Standup Target Animation
'****************************************************************************
'USAGE: 	Sub sw1_Hit: 	PrimStandupTgtHit  1, Sw1, PrimSw1: End Sub
'USAGE: 	Sub Sw1_Timer: 	PrimStandupTgtMove 1, Sw1, PrimSw1: End Sub

Const StandupTgtMovementDir = "TransY" 
Const StandupTgtMovementMax = 6	 

Sub PrimStandupTgtHit (swnum, wallName, primName) 	
	PlaySound SoundFX("target",DOFContactors)
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

'*****************************************
'   rothbauerw's Manual Ball Control
'*****************************************

Dim BCup, BCdown, BCleft, BCright
Dim ControlBallInPlay, ControlActiveBall
Dim BCvel, BCyveloffset, BCboostmulti, BCboost

BCboost = 1				'Do Not Change - default setting
BCvel = 4				'Controls the speed of the ball movement
BCyveloffset = -0.01 	'Offsets the force of gravity to keep the ball from drifting vertically on the table, should be negative
BCboostmulti = 3		'Boost multiplier to ball veloctiy (toggled with the B key) 

ControlBallInPlay = false

Sub StartBallControl_Hit()
	Set ControlActiveBall = ActiveBall
	ControlBallInPlay = true
End Sub

Sub StopBallControl_Hit()
	ControlBallInPlay = false
End Sub	

Sub BallControlTimer_Timer()
	If EnableBallControl and ControlBallInPlay then
		If BCright = 1 Then
			ControlActiveBall.velx =  BCvel*BCboost
		ElseIf BCleft = 1 Then
			ControlActiveBall.velx = -BCvel*BCboost
		Else
			ControlActiveBall.velx = 0
		End If

		If BCup = 1 Then
			ControlActiveBall.vely = -BCvel*BCboost
		ElseIf BCdown = 1 Then
			ControlActiveBall.vely =  BCvel*BCboost
		Else
			ControlActiveBall.vely = bcyveloffset
		End If
	End If
End Sub

if flippershadows=1 then 
	FlipperLSh.visible=1
	FlipperRSh.visible=1
else
	FlipperLSh.visible=0
	FlipperRSh.visible=0
end if

'*****************************************
'			FLIPPER SHADOWS
'*****************************************

sub FlipperTimer_Timer()
	FlipperLSh.RotZ = LeftFlipper.currentangle
	FlipperRSh.RotZ = RightFlipper.currentangle
    flipperleft_prim.rotz = leftflipper.currentangle
    flipperright_prim.rotz = rightflipper.currentangle
	plungegateleft_prim.RotX = Gate9.CurrentAngle + 60
	plungegateright_prim.RotX = Gate8.CurrentAngle + 60
End Sub

if ballshadows=1 then
	BallShadowUpdate.enabled=1
else
	BallShadowUpdate.enabled=0
end if

'*****************************************
'			BALL SHADOW
'*****************************************
Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5,BallShadow6,BallShadow7,BallShadow8,BallShadow9,BallShadow10)

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
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/16) + ((BOT(b).X - (Table1.Width/2))/17))' + 13
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/16) + ((BOT(b).X - (Table1.Width/2))/17))' - 13
        End If
        ballShadow(b).Y = BOT(b).Y + 10
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
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


Sub Pins_Hit (idx)
	PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
	PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
	PlaySound "metalhit_thin", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
	PlaySound "metalhit_medium", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
	PlaySound "metalhit2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
	PlaySound "gate4", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinner_Spin
	PlaySound "fx_spinner",0,.25,0,0.25
End Sub

Sub Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then 
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub Posts_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 16 then 
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 16 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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
		Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub

Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub
 
Sub table1_Exit:Controller.Stop:End Sub

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

Const tnob = 10 ' total number of balls
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


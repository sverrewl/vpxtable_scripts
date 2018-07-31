Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="torch",UseSolenoids=1,UseLamps=1,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff",SFlipperOn="fx_Flipperup",SFlipperOff="fx_Flipperdown"
Const SCoin="coin",cCredits=""

LoadVPM"01150000","GTS1.VBS",3.22

'**********************************************************
'********   	OPTIONS		*******************************
'**********************************************************

Dim BallShadows: Ballshadows=1  		'******************	set to 1 to turn on Ball shadows
Dim FlipperShadows: FlipperShadows=1  	'***********		set to 1 to turn on Flipper shadows
Dim ChimesorTones: ChimesorTones=1		'**************		set to 0 for chimes, 1 for tones (machine came with electronic sounds but a 3 chime set can easily be swapped in


'Solenoid Call backs
'**********************************************************************************************************
SolCallback(1)="DrainKick"
SolCallback(2)="vpmSolSound SoundFX(""Knock"",DOFKnocker),"
if ChimesorTones=1 then 
	SolCallback(3)="vpmSolSound SoundFX(""10tone"", DOFChimes),"
	SolCallback(4)="vpmSolSound SoundFX(""100tone"", DOFChimes),"
	SolCallback(5)="vpmSolSound SoundFX(""1000tone"", DOFChimes),"
  else
	SolCallback(3)="vpmSolSound SoundFX(""10a"", DOFChimes),"
	SolCallback(4)="vpmSolSound SoundFX(""100a"", DOFChimes),"
	SolCallback(5)="vpmSolSound SoundFX(""1000a"", DOFChimes),"
end if
SolCallback(6)="Lraised" 	
SolCallback(7)="Rraised" 	
SolCallback(8)="SolRoto"	

SolCallback(17)= "FastFlips.TiltSol"
SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub Rraised(enabled)
	if enabled then Rreset.enabled=True
End Sub

Sub Rreset_timer
	for each objekt in DTright
		objekt.isdropped=0
	next
	controller.switch(60) = 0
	controller.switch(61) = 0
	controller.switch(62) = 0
	controller.switch(63) = 0
	controller.switch(64) = 0
	PlaySoundAt SoundFX("DTreset", DOFContactors), sw64
	For each light in DTRightLights: light.state=0: Next
	me.enabled=false
End Sub

Sub Lraised(enabled)
	if enabled then Lreset.enabled=True
End Sub

Sub Lreset_timer
	for each objekt in dtLeft
		objekt.isdropped=0
	next
	controller.switch(70) = 0
	controller.switch(71) = 0
	controller.switch(72) = 0
	controller.switch(73) = 0
	controller.switch(74) = 0
	PlaySoundAt SoundFX("DTreset", DOFContactors), sw70
	For each light in DTLeftLights: light.state=0: Next
	me.enabled=false
End Sub

Sub DrainKick(enabled)
	If enabled Then
	Drain.kick 70, 12
	PlaySoundAt SoundFX("ballrelease",DOFContactors), Drain
	End If
end Sub

Sub Drain_Hit
	PlaySoundAt "drain", Drain
	controller.switch(66) = 1
End Sub

sub Drain_unhit
	controller.Switch(66) = 0
end sub

Sub SolLFlipper(Enabled)
     If Enabled Then
		PlaySoundAt SoundFX("fx_flipperup",DOFFlippers), LeftFlipper
		LeftFlipper.RotateToEnd
     Else
		PlaySoundAt SoundFX("fx_flipperdown",DOFFlippers), LeftFlipper
		LeftFlipper.RotateToStart
     End If
  End Sub
  
Sub SolRFlipper(Enabled)
     If Enabled Then
		PlaySoundAt SoundFX("fx_flipperup",DOFFlippers), RightFlipper
		RightFlipper.RotateToEnd
     Else
		PlaySoundAt SoundFX("fx_flipperdown",DOFFlippers), RightFlipper
		RightFlipper.RotateToStart
     End If
End Sub

'**********************************************************************************************************


'Primitive Flipper Code

Sub FlipperTimer_Timer

	lflip.roty = LeftFlipper.currentangle 
	rflip.roty = RightFlipper.currentangle 
	Pgate1.rotz = (Gate1.currentangle*.75)+25
	l33a.state = l33.state

	if FlipperShadows=1 then
		FlipperLSh.RotZ = LeftFlipper.currentangle
		FlipperRSh.RotZ = RightFlipper.currentangle
	end if

End Sub

'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************


Dim BPG, hsaward, n, plungerball, xx, Lstep, FastFlips, objekt

Sub Torch_Init
	vpmInit Me
	plungerball=0
	CurrRotoPos=1
	On Error Resume Next
		With Controller
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
		.SplashInfoLine="Close Encounters (Gottlieb 1978)"
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
	vpmNudge.TiltSwitch=4
	vpmNudge.Sensitivity=.5
	vpmNudge.TiltObj=Array(BumperL,BumperR,SlingL)
 
	Set FastFlips = new cFastFlips
	with FastFlips
		.CallBackL = "SolLflipper"	'Point these to flipper subs
		.CallBackR = "SolRflipper"	'...
'		.CallBackUL = "SolULflipper"'...(upper flippers, if needed)
'		.CallBackUR = "SolURflipper"'...
		.TiltObjects = True 'Optional, if True calls vpmnudge.solgameon automatically. IF YOU GET A LINE 1 ERROR, DISABLE THIS! (or setup vpmNudge.TiltObj!)
		.InitDelay "FastFlips", 100			'Optional, if > 0 adds some compensation for solenoid jitter (occasional problem on Bram Stoker's Dracula)
		.DebugOn = False		'Debug, always-on flippers. Call FastFlips.DebugOn True or False in debugger to enable/disable.
	end with

 	vpmMapLights AllLights

	FindDips		'find balls per game and high score reward

	If B2SOn Then 'Show Desktop components
		for each objekt in BackdropStuff: objekt.visible=0: next
	  Else
		for each objekt in BackdropStuff: objekt.visible=1: next
	End if

	if ballshadows=1 then
		BallShadowUpdate.enabled=1
	  else
		BallShadowUpdate.enabled=0
	end if

	if flippershadows=1 then 
		FlipperLSh.visible=1
		FlipperRSh.visible=1
	  else
		FlipperLSh.visible=0
		FlipperRSh.visible=0
	end if

	drain.createball
    startGame.enabled=true

End Sub


Sub Torch_Paused:Controller.Pause = 1:End Sub

Sub Torch_unPaused:Controller.Pause = 0:End Sub 

Sub Torch_Exit
	If b2son then controller.stop
End Sub

sub startGame_timer
	playsoundat "poweron", Plunger
	For each xx in GILights:xx.State = 1: Next		'*****GI Lights On
	For each xx in DTLeftLights: xx.state=0:Next
	For each xx in DTRightLights: xx.state=0:Next
	RotoFlasher.visible=1
	me.enabled=false
end sub


Sub Torch_KeyDown(ByVal keycode)
	If keycode = LeftFlipperKey Then FastFlips.FlipL True :  FastFlips.FlipUL True
	If keycode = RightFlipperKey Then FastFlips.FlipR True :  FastFlips.FlipUR True
	If vpmKeyDown(KeyCode) Then Exit Sub
	If keycode=AddCreditKey then PlaySoundAt "coinIn6", Drain: vpmTimer.pulseSW (swCoin1): end if
	If keycode=PlungerKey Then Plunger.Pullback:PlaySoundAt "plungerpull", Plunger

'************************   Start Ball Control 1/3
	if keycode = 46 then	 			' C Key
		If contball = 1 Then
			contball = 0
		Else
			contball = 1
		End If
	End If
	if keycode = 48 then 				'B Key
		If bcboost = 1 Then
			bcboost = bcboostmulti
		Else
			bcboost = 1
		End If
	End If
	if keycode = 203 then bcleft = 1		' Left Arrow
	if keycode = 200 then bcup = 1			' Up Arrow
	if keycode = 208 then bcdown = 1		' Down Arrow
	if keycode = 205 then bcright = 1		' Right Arrow
'************************   End Ball Control 1/3

End Sub

Sub Torch_KeyUp(ByVal keycode)
	If keycode = LeftFlipperKey Then FastFlips.FlipL False :  FastFlips.FlipUL False
   	If keycode = RightFlipperKey Then FastFlips.FlipR False :  FastFlips.FlipUR False
	
	If keycode = 61 then FindDips
	If vpmKeyUp(KeyCode) Then Exit Sub
	If keycode=PlungerKey Then 
		Plunger.Fire
		if plungerball=1 then
			PlaySoundAt "plungerreleaseball", Plunger
		  else
			PlaySoundAt "plungerreleasefree", Plunger
		end if
	end if

'************************   Start Ball Control 2/3
	if keycode = 203 then bcleft = 0		' Left Arrow
	if keycode = 200 then bcup = 0			' Up Arrow
	if keycode = 208 then bcdown = 0		' Down Arrow
	if keycode = 205 then bcright = 0		' Right Arrow
'************************   End Ball Control 2/3

End Sub

'************************   Start Ball Control 3/3
Sub StartControl_Hit()
	Set ControlBall = ActiveBall
	contballinplay = true
End Sub

Sub StopControl_Hit()
	contballinplay = false
End Sub	

Dim bcup, bcdown, bcleft, bcright, contball, contballinplay, ControlBall, bcboost
Dim bcvel, bcyveloffset, bcboostmulti

bcboost = 1		'Do Not Change - default setting
bcvel = 4		'Controls the speed of the ball movement
bcyveloffset = -0.01 	'Offsets the force of gravity to keep the ball from drifting vertically on the table, should be negative
bcboostmulti = 3	'Boost multiplier to ball veloctiy (toggled with the B key) 

Sub BallControl_Timer()
	If Contball and ContBallInPlay then
		If bcright = 1 Then
			ControlBall.velx = bcvel*bcboost
		ElseIf bcleft = 1 Then
			ControlBall.velx = - bcvel*bcboost
		Else
			ControlBall.velx=0
		End If

		If bcup = 1 Then
			ControlBall.vely = -bcvel*bcboost
		ElseIf bcdown = 1 Then
			ControlBall.vely = bcvel*bcboost
		Else
			ControlBall.vely= bcyveloffset
		End If
	End If
End Sub
'************************   End Ball Control 3/3

Sub Ballhome_hit
	plungerball=1
end sub

sub ballhome_unhit
	plungerball=0
end sub

'Gate

	  Sub sw21_Hit:vpmTimer.PulseSW 21:End Sub




'bumpers
Sub BumperL_Hit
	vpmTimer.PulseSw 11 
	PlaySoundAt SoundFXDOF("fx_bumper",101, DOFPulse, DOFContactors), BumperL
End Sub

Sub BumperR_Hit
	vpmTimer.PulseSw 11 
	PlaySoundAt SoundFXDOF("fx_bumper",102, DOFPulse, DOFContactors), BumperR
End Sub

'Hit Targets

Sub sw12_Hit:vpmTimer.PulseSw 12:End Sub


'Rollover wire Triggers

	  Sub sw10_Hit:Controller.Switch(10) = 1:End Sub
	  Sub sw10_UnHit:Controller.Switch(10) = 0:End Sub
	  Sub sw13_Hit:Controller.Switch(13) = 1:End Sub
	  Sub sw13_UnHit:Controller.Switch(13) = 0:End Sub
	  Sub sw14_Hit:Controller.Switch(14) = 1:End Sub
	  Sub sw14_UnHit:Controller.Switch(14) = 0:End Sub
	  Sub sw20_Hit:Controller.Switch(20) = 1:End Sub
	  Sub sw20_UnHit:Controller.Switch(20) = 0:End Sub
	  Sub sw23_Hit:Controller.Switch(23) = 1:End Sub
	  Sub sw23_UnHit:Controller.Switch(23) = 0:End Sub
	  Sub sw24_Hit:Controller.Switch(24) = 1:End Sub
	  Sub sw24_UnHit:Controller.Switch(24) = 0:End Sub
	  Sub sw32_Hit:
			Select Case (rotoSW(CurrRotoPos-1))
				Case 1:vpmtimer.PulseSw(34) '1
				Case 2:vpmtimer.PulseSw(32) '2
				Case 3:vpmtimer.PulseSw(31) '3
			End Select
	  End Sub

	  Sub sw52_Hit:
			Select Case (rotoSW(CurrRotoPos+1))
				Case 1:vpmtimer.PulseSw(54) '1
				Case 2:vpmtimer.PulseSw(52) '2
				Case 3:vpmtimer.PulseSw(51) '3
			End Select
	  End Sub


'Slingshot

Sub LeftSlingShot_Slingshot
	vpmTimer.PulseSw 40
    PlaySoundAt SoundFX("left_slingshot",DOFContactors), slingL
    LSling.Visible = 0
    LSling1.Visible = 1
	slingL.objroty = 15
    me.uservalue = 1
    me.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case me.uservalue
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:slingL.objroty = 7
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:slingL.objroty = 0:me.TimerEnabled = 0:
    End Select
    me.uservalue=me.uservalue + 1
End Sub
		 
Sub RightSlingShot_Slingshot
	vpmTimer.PulseSw 50
    PlaySoundAt SoundFX("right_slingshot",DOFContactors), slingR
    RSling.Visible = 0
    RSling1.Visible = 1
	slingR.objroty = -15
    me.uservalue = 1
    me.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case me.uservalue
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:slingR.objroty = -7
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:slingR.objroty = 0:me.TimerEnabled = 0:
    End Select
    me.uservalue = me.uservalue + 1
End Sub

'Scoring rubbers - ANIMATED!!!!!

	Sub sw30a_hit()
		vpmTimer.PulseSw 30
		R30a.visible=0
		R30a1.visible=1
		me.uservalue=1
		me.timerenabled=1
	End Sub

	sub sw30a_timer
		select case sw30a.uservalue
			Case 1: r30a1.visible=0: r30a.visible=1
			Case 2: r30a.visible=0: r30a2.visible=1
			Case 3: r30a2.visible=0: r30a.visible=1: me.timerenabled=0
		end Select
		me.uservalue=me.uservalue+1
	end sub

	Sub sw30b_hit()
		vpmTimer.PulseSw 30
		R30b.visible=0
		R30b1.visible=1
		me.uservalue=1
		me.timerenabled=1
	End Sub

	sub sw30b_timer
		select case sw30b.uservalue
			Case 1: r30b1.visible=0: r30b.visible=1
			Case 2: r30b.visible=0: r30b2.visible=1
			Case 3: r30b2.visible=0: r30b.visible=1: me.timerenabled=0
		end Select
		me.uservalue=me.uservalue+1
	end sub

	Sub sw30c_hit()
		vpmTimer.PulseSw 30
		R30c.visible=0
		R30c1.visible=1
		me.uservalue=1
		me.timerenabled=1
	End Sub

	sub sw30c_timer
		select case sw30c.uservalue
			Case 1: r30c1.visible=0: r30c.visible=1
			Case 2: r30c.visible=0: r30c2.visible=1
			Case 3: r30c2.visible=0: r30c.visible=1: me.timerenabled=0
		end Select
		me.uservalue=me.uservalue+1
	end sub

	Sub sw30d_hit()
		vpmTimer.PulseSw 30
		R30a.visible=0
		R30d1.visible=1
		me.uservalue=1
		me.timerenabled=1
	End Sub

	sub sw30d_timer
		select case sw30d.uservalue
			Case 1: r30d1.visible=0: r30a.visible=1
			Case 2: r30a.visible=0: r30d2.visible=1
			Case 3: r30d2.visible=0: r30a.visible=1: me.timerenabled=0
		end Select
		me.uservalue=me.uservalue+1
	end sub

	Sub sw30e_hit()
		vpmTimer.PulseSw 30
		R30e.visible=0
		R30e1.visible=1
		me.uservalue=1
		me.timerenabled=1
	End Sub

	sub sw30e_timer
		select case sw30e.uservalue
			Case 1: r30e1.visible=0: r30e.visible=1
			Case 2: r30e.visible=0: r30e2.visible=1
			Case 3: r30e2.visible=0: r30e.visible=1: me.timerenabled=0
		end Select
		me.uservalue=me.uservalue+1
	end sub

	Sub sw30f_hit()
		vpmTimer.PulseSw 30
		R30e.visible=0
		R30f1.visible=1
		me.uservalue=1
		me.timerenabled=1
	End Sub

	sub sw30f_timer
		select case sw30f.uservalue
			Case 1: r30f1.visible=0: r30e.visible=1
			Case 2: r30e.visible=0: r30f2.visible=1
			Case 3: r30f2.visible=0: r30e.visible=1: me.timerenabled=0
		end Select
		me.uservalue=me.uservalue+1
	end sub


'Drop Targets

	Sub sw60_dropped
		controller.switch(60) = 1
		Lsw60.state=1
	End Sub
	
	Sub sw61_dropped
		controller.switch(61) = 1
		Lsw61.state=1
		Lsw61a.state=1
	End Sub

	Sub sw62_dropped
		controller.switch(62) = 1
		Lsw62.state=1
		Lsw62a.state=1
	End Sub

	Sub sw63_dropped
		controller.switch(63) = 1
		Lsw63.state=1

	End Sub

	Sub sw64_dropped
		controller.switch(64) = 1
		Lsw64.state=1
	End Sub

	Sub sw70_dropped
		controller.switch(70) = 1
		Lsw70.state=1
	End Sub
	
	Sub sw71_dropped
		controller.switch(71) = 1
		Lsw71.state=1
		Lsw71a.state=1
	End Sub

	Sub sw72_dropped
		controller.switch(72) = 1
		Lsw72.state=1
		Lsw72a.state=1
	End Sub

	Sub sw73_dropped
		controller.switch(73) = 1
		Lsw73.state=1
		Lsw73a.state=1
		Lsw73b.state=1
	End Sub

	Sub sw74_dropped
		controller.switch(74) = 1
		Lsw74.state=1
		Lsw74a.state=1
	End Sub


' NEW ROTO CODE based on Pinuck in vp9 Close Encounters modified by BorgDog to fit primitive roto targets
' ========================

' -- roto switches:
' Left 1  = 34
' Left 2  = 32
' Left 3  = 31 
' Center 1  = 44
' Center 2  = 42
' Center 3  = 41 
' Right 1  = 54
' Right 2  = 52
' Right 3  = 51 


' -- rotowheel's 15 targets
' 1  3
' 2  2
' 3  1
' 4  3
' 5  2
' 6  1
' 7  3
' 8  2
' 9  1
' 10 3
' 11 2
' 12 1
' 13 3
' 14 2
' 15 1

Dim RotoTick, NewRotoPos, CurrRotoPos, RotoSpinning, rotoSW, rotoP
'Dim R1Value, R2Value, R3Value
rotoSW = array(1,3,2,1,3,2,1,3,2,1,3,2,1,3,2,1,2)
rotoP = array(n, Proto1, Proto2, Proto3, Proto4, Proto5, Proto6, Proto7, Proto8, Proto9, Proto10, Proto11, Proto12, Proto13, Proto14, Proto15)

Sub SolRoto(Enabled)
  if RotoTimer.enabled=0 then
	NewRotoPos = CurrRotoPos - 3 - Int(5 * Rnd) 'random but not overlapping old (from Scapino)
	If NewRotoPos<1 Then NewRotoPos=NewRotoPos+15
	RotoTimer.enabled=1 'spin it!
	DOF 111, DOFOn
	PlaySoundAt SoundFX("roto_9",DOFGear), TrotoCenter
  end if
End Sub

Sub RotoTimer_Timer() 'enable to spin roto
	dim roti
	RotoTick=RotoTick+1
	for roti=1 to 15
		rotoP(roti).roty=rotoP(roti).roty+3
	next
	if rotoP(NewRotoPos).roty=12 then
		me.enabled=0
		DOF 111, DOFOff
		stopsound "roto_9"
		for roti=1 to 15
			rotoP(roti).roty=rotoP(roti).roty-12
		next
		CurrRotoPos=NewRotoPos
	end if
	for roti=1 to 15
		if rotoP(roti).roty>359 then rotoP(roti).roty=rotoP(roti).roty-360
	next
End Sub



Sub TrotoLeft_Hit 'LEFT ROTO
	if CurrRotoPos=1 then
		rotoP(15).rotx=3
	  else
		rotoP(CurrRotoPos-1).rotx=3
	end if
	me.timerenabled=1
	If RotoTimer.enabled=1 Then Exit Sub
	Select Case (rotoSW(CurrRotoPos-1))
		Case 1:vpmtimer.PulseSw(34) '1
		Case 2:vpmtimer.PulseSw(32) '2
		Case 3:vpmtimer.PulseSw(31) '3
	End Select
End Sub 

Sub TrotoLeft_timer
	if CurrRotoPos=1 then
		rotoP(15).rotx=0
	  else
		rotoP(CurrRotoPos-1).rotx=0
	end if
end sub

Sub TrotoCenter_Hit 'CENTER ROTO
	rotoP(CurrRotoPos).rotx=3
	me.timerenabled=1
	If RotoTimer.enabled=1 Then Exit Sub
	Select Case (rotoSW(CurrRotoPos))
		Case 1:vpmtimer.PulseSw(44) '1
		Case 2:vpmtimer.PulseSw(42) '2
		Case 3:vpmtimer.PulseSw(41) '3
	End Select
End Sub 

Sub TrotoCenter_timer
	rotoP(CurrRotoPos).rotx=0
end sub

Sub TrotoRight_Hit 'RIGHT ROTO
	if CurrRotoPos=15 then
		rotoP(1).rotx=3
	  else
		rotoP(CurrRotoPos+1).rotx=3
	end if
	me.timerenabled=true
	If RotoTimer.enabled=1 Then Exit Sub
	Select Case (rotoSW(CurrRotoPos+1))
		Case 1:vpmtimer.PulseSw(54) '1
		Case 2:vpmtimer.PulseSw(52) '2
		Case 3:vpmtimer.PulseSw(51) '3
	End Select
End Sub 

Sub TrotoRight_timer
	if CurrRotoPos=15 then
		rotoP(1).rotx=0
	  else
		rotoP(CurrRotoPos+1).rotx=0
	end if
end sub

 
' END ROTO -----

 Dim N1,O1, Light
 N1=0:O1=0:
 Set LampCallback=GetRef("UpdateMultipleLamps")
 
Sub UpdateMultipleLamps

		N1=Controller.Lamp(1) 'Game Over triggers match and BIP
			If N1 then 
				EMReelBIP.setvalue 1
				EMReelNTM.setvalue 0
				EMReelGO.setvalue 0
			  else
				EMReelBIP.setvalue 0
				EMReelNTM.setvalue 1
				EMReelGO.setvalue 1
			end if

		N1=Controller.Lamp(2) 'Tilt
			If N1 then
				EMReelTilt.setvalue 1
			  else
				EMReelTilt.setvalue 0
			end if

		N1=Controller.Lamp(4) 'Shoot Again
			If N1 then
				EMReelSPSA.setvalue 1
			  else
				EMReelSPSA.setvalue 0
			end if

		N1=Controller.Lamp(3) 'HIGH SCORE TO DATE
			if N1 then
				EMReelHTD.setvalue 1
			  else
				EMReelHTD.setvalue 0
			end if

 End Sub

Dim Digits(32)

'Score displays

Digits(0)=Array(a1,a2,a3,a4,a5,a6,a7,n,a8)
Digits(1)=Array(a9,a10,a11,a12,a13,a14,a15,n,a16)
Digits(2)=Array(a17,a18,a19,a20,a21,a22,a23,n,a24)
Digits(3)=Array(a25,a26,a27,a28,a29,a30,a31,n,a32)
Digits(4)=Array(a33,a34,a35,a36,a37,a38,a39,n,a40)
Digits(5)=Array(a41,a42,a43,a44,a45,a46,a47,n,a48)
Digits(6)=Array(a49,a50,a51,a52,a53,a54,a55,n,a56)
Digits(7)=Array(a57,a58,a59,a60,a61,a62,a63,n,a64)
Digits(8)=Array(a65,a66,a67,a68,a69,a70,a71,n,a72)
Digits(9)=Array(a73,a74,a75,a76,a77,a78,a79,n,a80)
Digits(10)=Array(a81,a82,a83,a84,a85,a86,a87,n,a88)
Digits(11)=Array(a89,a90,a91,a92,a93,a94,a95,n,a96)
Digits(12)=Array(a97,a98,a99,a100,a101,a102,a103,n,a104)
Digits(13)=Array(a105,a106,a107,a108,a109,a110,a111,n,a112)
Digits(14)=Array(a113,a114,a115,a116,a117,a118,a119,n,a120)
Digits(15)=Array(a121,a122,a123,a124,a125,a126,a127,n,a128)
Digits(16)=Array(a129,a130,a131,a132,a133,a134,a135,n,a136)
Digits(17)=Array(a137,a138,a139,a140,a141,a142,a143,n,a144)
Digits(18)=Array(a145,a146,a147,a148,a149,a150,a151,n,a152)
Digits(19)=Array(a153,a154,a155,a156,a157,a158,a159,n,a160)
Digits(20)=Array(a161,a162,a163,a164,a165,a166,a167,n,a168)
Digits(21)=Array(a169,a170,a171,a172,a173,a174,a175,n,a176)
Digits(22)=Array(a177,a178,a179,a180,a181,a182,a183,n,a184)
Digits(23)=Array(a185,a186,a187,a188,a189,a190,a191,n,a192)

'Ball in Play and Credit displays

Digits(26)=Array(e00,e01,e02,e03,e04,e05,e06,n,e08)
Digits(27)=Array(e10,e11,e12,e13,e14,e15,e16,n,e18)
Digits(24)=Array(f00,f01,f02,f03,f04,f05,f06,n,f08)
Digits(25)=Array(f10,f11,f12,f13,f14,f15,f16,n,f18)



Sub DisplayTimer_Timer
	Dim ChgLED,ii,num,chg,stat,obj
	ChgLed = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
	If Not IsEmpty(ChgLED) Then
		If not b2son Then
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



'Finding an individual dip state based on scapino's Strikes and spares dip code - from unclewillys pinball pool
Dim TheDips(32)
  Sub FindDips
    Dim DipsNumber
	DipsNumber = Controller.Dip(1)
	TheDips(16) = Int(DipsNumber/128)
	If TheDips(16) = 1 then DipsNumber = DipsNumber - 128 end if
	TheDips(15) = Int(DipsNumber/64)
	If TheDips(15) = 1 then DipsNumber = DipsNumber - 64 end if
	TheDips(14) = Int(DipsNumber/32)
	If TheDips(14) = 1 then DipsNumber = DipsNumber - 32 end if
	TheDips(13) = Int(DipsNumber/16)
	If TheDips(13) = 1 then DipsNumber = DipsNumber - 16 end if
	TheDips(12) = Int(DipsNumber/8)
	If TheDips(12) = 1 then DipsNumber = DipsNumber - 8 end if
	TheDips(11) = Int(DipsNumber/4)
	If TheDips(11) = 1 then DipsNumber = DipsNumber - 4 end if
	TheDips(10) = Int(DipsNumber/2)
	If TheDips(10) = 1 then DipsNumber = DipsNumber - 2 end if
	TheDips(9) = Int(DipsNumber)
	DipsNumber = Controller.Dip(2)
	TheDips(24) = Int(DipsNumber/128)
	If TheDips(24) = 1 then DipsNumber = DipsNumber - 128 end if
	TheDips(23) = Int(DipsNumber/64)
	If TheDips(23) = 1 then DipsNumber = DipsNumber - 64 end if
	TheDips(22) = Int(DipsNumber/32)
	If TheDips(22) = 1 then DipsNumber = DipsNumber - 32 end if
	TheDips(21) = Int(DipsNumber/16)
	If TheDips(21) = 1 then DipsNumber = DipsNumber - 16 end if
	TheDips(20) = Int(DipsNumber/8)
	If TheDips(20) = 1 then DipsNumber = DipsNumber - 8 end if
	TheDips(19) = Int(DipsNumber/4)
	If TheDips(19) = 1 then DipsNumber = DipsNumber - 4 end if
	TheDips(18) = Int(DipsNumber/2)
	If TheDips(18) = 1 then DipsNumber = DipsNumber - 2 end if
	TheDips(17) = Int(DipsNumber)
	DipsTimer.Enabled=1
 End Sub


 Sub DipsTimer_Timer()
	hsaward = TheDips(22)
	BPG = TheDips(9)
	dim ebplay: ebplay= TheDips(11)
	If BPG = 1 then 
		if ebplay = 1 then
			instcard.image="InstCard3Balls"
		  else
			instcard.image="InstCard3BallsEB"
		end if
	  Else
		if ebplay = 1 then
			instcard.image="InstCard5Balls"
		  else
			instcard.image="InstCard5BallsEB"
		end if
	End if
	if ebplay = 1 then
		repcard.image="replaycard"&hsaward
	  else
		repcard.image="replaycardeb"&hsaward
	end if
	DipsTimer.enabled=0
 End Sub

 'Gottlieb System 1
 'added by Inkochnito
 Sub editDips
 	Dim vpmDips : Set vpmDips = New cvpmDips
 	With vpmDips
 		.AddForm 700,400,"System 1 - DIP switches"
 		.AddFrame 205,0,190,"Maximum credits",&H00030000,Array("5 credits",0,"8 credits",&H00020000,"10 credits",&H00010000,"15 credits",&H00030000)'dip 17&18
 		.AddFrame 0,0,190,"Coin chute control",&H00040000,Array("seperate",0,"same",&H00040000)'dip 19
 		.AddFrame 0,46,190,"Game mode",&H00000400,Array("extra ball",0,"replay",&H00000400)'dip 11
 		.AddFrame 0,92,190,"High game to date awards",&H00200000,Array("no award",0,"3 replays",&H00200000)'dip 22
 		.AddFrame 0,138,190,"Balls per game",&H00000100,Array("5 balls",0,"3 balls",&H00000100)'dip 9
 		.AddFrame 0,184,190,"Tilt effect",&H00000800,Array("game over",0,"ball in play only",&H00000800)'dip 12
 		.AddChk 205,80,190,Array("Match feature",&H00000200)'dip 10
 		.AddChk 205,95,190,Array("Credits displayed",&H00001000)'dip 13
 		.AddChk 205,110,190,Array("Play credit button tune",&H00002000)'dip 14
 		.AddChk 205,125,190,Array("Play tones when scoring",&H00080000)'dip 20
 		.AddChk 205,140,190,Array("Play coin switch tune",&H00400000)'dip 23
 		.AddChk 205,155,190,Array("High game to date displayed",&H00100000)'dip 21
 		.AddLabel 50,240,300,20,"After hitting OK, press F3 to reset game with new settings."
 		.ViewDips
 	End With
 End Sub
 Set vpmShowDips = GetRef("editDips")

'**********************************************************************************************************
'**********************************************************************************************************
'**********************************************************************************************************


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


'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
	Dim tmp
    tmp = tableobj.y * 2 / Torch.height-1
    If tmp > 0 Then
		AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = tableobj.x * 2 / Torch.width-1
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


'*****************************************
'			BALL SHADOW
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
        If BOT(b).X < Torch.Width/2 Then
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/16) + ((BOT(b).X - (Torch.Width/2))/17))' + 13
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/16) + ((BOT(b).X - (Torch.Width/2))/17))' - 13
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


Sub a_Triggers_Hit (idx)
	playsound "sensor", 0,1,AudioPan(ActiveBall),0,0,0,1,AudioFade(ActiveBall)
End sub

Sub a_Pins_Hit (idx)
	PlaySound "pinhit_low", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub a_Targets_Hit (idx)
	PlaySound "target", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub a_DropTargets_Hit (idx)
	PlaySound "DTDrop", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub a_Metals_Thin_Hit (idx)
	PlaySound "metalhit_thin", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub a_Metals_Medium_Hit (idx)
	PlaySound "metalhit_medium", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub a_Metals2_Hit (idx)
	PlaySound "metalhit2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub a_Gates_Hit (idx)
	PlaySound "gate4", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinner_Spin
	PlaySound "fx_spinner", 0, .25, AudioPan(Spinner), 0.25, 0, 0, 1, AudioFade(Spinner)
End Sub

Sub a_Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then 
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub a_Posts_Hit(idx)
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

Sub LeftFlipper_Collide(parm)
 	RandomSoundFlipper()
End Sub


Sub RightFlipper_Collide(parm)
 	RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub




'cFastFlips by nFozzy
'Bypasses pinmame callback for faster and more responsive flippers
'Version 1.1 beta2 (More proper behaviour, extra safety against script errors)

'Flipper / game-on Solenoid # reference
'Atari: Sol16
'Astro:  ?
'Bally Early 80's: Sol19
'Bally late 80's (Blackwater 100, etc): Sol19
'Game Plan: Sol16
'Gottlieb System 1: Sol17
'Gottlieb System 80: No dedicated flipper solenoid? GI circuit Sol10?
'Gottlieb System 3: Sol32
'Playmatic: Sol8
'Spinball: Sol25
'Stern (80's): Sol19
'Taito: ?
'Williams System 3, 4, 6: Sol23
'Williams System 7: Sol25
'Williams System 9: Sol23
'Williams System 11: Sol23
'Bally / Williams WPC 90', 92', WPC Security: Sol31
'Data East (and Sega pre-whitestar): Sol23
'Zaccaria: ???

'********************Setup*******************:

'....somewhere outside of any subs....
'dim FastFlips

'....table init....
'Set FastFlips = new cFastFlips
'with FastFlips
'	.CallBackL = "SolLflipper"	'Point these to flipper subs
'	.CallBackR = "SolRflipper"	'...
''	.CallBackUL = "SolULflipper"'...(upper flippers, if needed)
''	.CallBackUR = "SolURflipper"'...
'	.TiltObjects = True 'Optional, if True calls vpmnudge.solgameon automatically. IF YOU GET A LINE 1 ERROR, DISABLE THIS! (or setup vpmNudge.TiltObj!)
''	.InitDelay "FastFlips", 100			'Optional, if > 0 adds some compensation for solenoid jitter (occasional problem on Bram Stoker's Dracula)
''	.DebugOn = False		'Debug, always-on flippers. Call FastFlips.DebugOn True or False in debugger to enable/disable.
'end with

'...keydown section... commenting out upper flippers is not necessary as of 1.1
'If KeyCode = LeftFlipperKey then FastFlips.FlipL True :  FastFlips.FlipUL True
'If KeyCode = RightFlipperKey then FastFlips.FlipR True :  FastFlips.FlipUR True
'(Do not use Exit Sub, this script does not handle switch handling at all!)

'...keyUp section...
'If KeyCode = LeftFlipperKey then FastFlips.FlipL False :  FastFlips.FlipUL False
'If KeyCode = RightFlipperKey then FastFlips.FlipR False :  FastFlips.FlipUR False

'...Solenoid...
'SolCallBack(31) = "FastFlips.TiltSol"
'//////for a reference of solenoid numbers, see top /////


'One last note - Because this script is super simple it will call flipper return a lot.
'It might be a good idea to add extra conditional logic to your flipper return sounds so they don't play every time the game on solenoid turns off
'Example:
'Instead of
		'LeftFlipper.RotateToStart
		'playsound SoundFX("FlipperDown",DOFFlippers), 0, 1, 0.01	'return
'Add Extra conditional logic:
		'LeftFlipper.RotateToStart
		'if LeftFlipper.CurrentAngle = LeftFlipper.StartAngle then
		'	playsound SoundFX("FlipperDown",DOFFlippers), 0, 1, 0.01	'return
		'end if
'That's it]
'*************************************************

Function NullFunction(aEnabled):End Function	'1 argument null function placeholder
Class cFastFlips
	Public TiltObjects, DebugOn, hi
	Private SubL, SubUL, SubR, SubUR, FlippersEnabled, Delay, LagCompensation, Name, FlipState(3)
	
	Private Sub Class_Initialize()
		Delay = 0 : FlippersEnabled = False : DebugOn = False : LagCompensation = False
		Set SubL = GetRef("NullFunction"): Set SubR = GetRef("NullFunction") : Set SubUL = GetRef("NullFunction"): Set SubUR = GetRef("NullFunction")
	End Sub
	
	'set callbacks
	Public Property Let CallBackL(aInput)  : Set SubL  = GetRef(aInput) : Decouple sLLFlipper, aInput: End Property
	Public Property Let CallBackUL(aInput) : Set SubUL = GetRef(aInput) : End Property
	Public Property Let CallBackR(aInput)  : Set SubR  = GetRef(aInput) : Decouple sLRFlipper, aInput:  End Property
	Public Property Let CallBackUR(aInput) : Set SubUR = GetRef(aInput) : End Property
	Public Sub InitDelay(aName, aDelay) : Name = aName : delay = aDelay : End Sub	'Create Delay
	'Automatically decouple flipper solcallback script lines (only if both are pointing to the same sub) thanks gtxjoe
	Private Sub Decouple(aSolType, aInput)  : If StrComp(SolCallback(aSolType),aInput,1) = 0 then SolCallback(aSolType) = Empty End If : End Sub

	'call callbacks
	Public Sub FlipL(aEnabled)
		FlipState(0) = aEnabled	'track flipper button states: the game-on sol flips immediately if the button is held down (1.1)
		If not FlippersEnabled and not DebugOn then Exit Sub
		subL aEnabled
	End Sub

	Public Sub FlipR(aEnabled)
		FlipState(1) = aEnabled
		If not FlippersEnabled and not DebugOn then Exit Sub
		subR aEnabled
	End Sub

	Public Sub FlipUL(aEnabled)
		FlipState(2) = aEnabled
		If not FlippersEnabled and not DebugOn then Exit Sub
		subUL aEnabled
	End Sub	

	Public Sub FlipUR(aEnabled)
		FlipState(3) = aEnabled
		If not FlippersEnabled and not DebugOn then Exit Sub
		subUR aEnabled
	End Sub	
	
	Public Sub TiltSol(aEnabled)	'Handle solenoid / Delay (if delayinit)
		If delay > 0 and not aEnabled then 	'handle delay
			vpmtimer.addtimer Delay, Name & ".FireDelay" & "'"
			LagCompensation = True
		else
			If Delay > 0 then LagCompensation = False
			EnableFlippers(aEnabled)
		end If
	End Sub
	
	Sub FireDelay() : If LagCompensation then EnableFlippers False End If : End Sub
	
	Private Sub EnableFlippers(aEnabled)
		If aEnabled then SubL FlipState(0) : SubR FlipState(1) : subUL FlipState(2) : subUR FlipState(3)
		FlippersEnabled = aEnabled
		If TiltObjects then vpmnudge.solgameon aEnabled
		If Not aEnabled then
			subL False
			subR False
			If not IsEmpty(subUL) then subUL False
			If not IsEmpty(subUR) then subUR False
		End If		
	End Sub
	
End Class

	
' Thalamus : Exit in a clean and proper way
Sub Torch_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


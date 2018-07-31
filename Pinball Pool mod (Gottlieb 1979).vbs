Option Explicit
Randomize

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' No special SSF tweaks yet.

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="pinpool",UseSolenoids=2,UseLamps=1,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff",SFlipperOn="fx_Flipperup",SFlipperOff="fx_Flipperdown"
Const SCoin="coin",cCredits=""

LoadVPM"01150000","GTS1.VBS",3.22

Dim Step23a, Step23b, Step23c, Step23d, Step73a, Step73b, Step73c, Step73d


Dim DTshadow: DTshadow=0.5    'set drop target shadow level from 0 to 1, smaller number is more shadow

'Solenoid Call backs
'**********************************************************************************************************
SolCallback(1)="bsTrough.SolOut"
SolCallback(2)="vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(3)="vpmSolSound SoundFX(""10pts"",DOFChimes),"
SolCallback(4)="vpmSolSound SoundFX(""100pts"",DOFChimes),"
SolCallback(5)="vpmSolSound SoundFX(""1000pts"",DOFChimes),"
SolCallback(6)="kicker2exit"		'"bsSaucer2.SolOut"
SolCallback(7)="kicker1exit"
SolCallback(8)="Rraised"   '"dtRDrop.SolDropUp"
SolCallback(17)="vpmNudge.SolGameOn"

Sub Rraised(enabled)
	if enabled then Rreset.enabled=True
End Sub

Sub Rreset_timer
	dtRDrop.DropSol_On
	For each light in DTRightLights: light.intensityscale=DTshadow: Next
	Rreset.enabled=false
End Sub

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFx("fx_Flipperup",DOFContactors):LeftFlipper.RotateToEnd
     Else
         PlaySound SoundFx("fx_Flipperdown",DOFContactors):LeftFlipper.RotateToStart
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFx("fx_Flipperup",DOFContactors):RightFlipper.RotateToEnd
     Else
         PlaySound SoundFx("fx_Flipperdown",DOFContactors):RightFlipper.RotateToStart
     End If
End Sub

'**********************************************************************************************************


'Primitive Flipper Code
Sub FlipperTimer_Timer
	FlipperT1.roty = LeftFlipper.currentangle
	FlipperT5.roty = RightFlipper.currentangle
	Pgate.rotz = Gate3.currentangle+25
End Sub

'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************

Dim bsTrough,dtRDrop,dtLDrop,bsSaucer1,bsSaucer2
Dim BPG

Sub PinballPool_Init
	vpmInit Me
	On Error Resume Next
		With Controller
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
		.SplashInfoLine="Pinball Pool (Gottlieb)"
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
	vpmNudge.Sensitivity=3
	vpmNudge.TiltObj=Array(Bumper1,Bumper2,LeftSlingshot,RightSlingshot)

	Kicker3.CreateBall
	Kicker3.Kick 70,1
	Kicker3.Enabled=False

	Set bsTrough=New cvpmBallStack
	bsTrough.InitSw 0,66,0,0,0,0,0,0
	bsTrough.InitKick DRAIN,60,30
	bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors),SoundFX("ballrelease",DOFContactors)
 	bsTrough.Balls=1

  	Set dtLDrop=New cvpmDropTarget
	dtLDrop.InitDrop Array(sw10,sw11,sw13,sw14,sw20,sw21,sw24),Array(10,11,13,14,20,21,24)
	dtLDrop.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

  	Set dtRDrop=New cvpmDropTarget
	dtRDrop.InitDrop Array(sw60,sw61,sw63,sw64,sw70,sw71,sw74),Array(60,61,63,64,70,71,74)
	dtRDrop.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

	'find balls per game
	FindDips

	'*****GI Lights On
	dim xx

	For each xx in GI:xx.State = 1: Next
	For each xx in DTLeftLights: xx.state=1: xx.intensityscale=DTshadow:Next
	For each xx in DTRightLights: xx.state=1: xx.intensityscale=DTshadow:Next

	If b2son Then
		for each xx in backdropstuff
			xx.visible=0
		Next
	end if

End Sub

Sub PinballPool_Exit()
	If B2SOn Then Controller.stop
End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************
 Sub PinballPool_KeyDown(ByVal keycode)
	If vpmKeyDown(KeyCode) Then Exit Sub
	If keycode=AddCreditKey then vpmTimer.pulseSW (swCoin1)
	If keycode=PlungerKey Then Plunger.Pullback:playsound"plungerpull"
End Sub

Sub PinballPool_KeyUp(ByVal keycode)
	If vpmKeyUp(KeyCode) Then Exit Sub
	If keycode=PlungerKey Then Plunger.Fire:playsound"plunger"
End Sub


Sub Drain_Hit:playsound"drain":bsTrough.addball me:End Sub


'bumpers
Sub Bumper1_Hit
	vpmTimer.PulseSw 50
	DOF 103,2
	playsound SoundFX("fx_bumper1",DOFContactors)
	me.timerenabled=1
End Sub

Sub Bumper2_Hit
	vpmTimer.PulseSw 50
	DOF 104,2
	playsound SoundFX("fx_bumper1",DOFContactors)
	me.timerenabled=1
End Sub


'Star Triggers
Sub SW54_Hit:Controller.Switch(54)=1 : playsound"rollover" : End Sub
Sub SW54_unHit:Controller.Switch(54)=0:End Sub
Sub SW54a_Hit:Controller.Switch(54)=1 : playsound"rollover" : End Sub
Sub SW54a_unHit:Controller.Switch(54)=0:End Sub


'Rollover wire Tirggers
Sub SW34a_Hit:Controller.Switch(34)=1 : playsound"rollover" : End Sub
Sub SW34a_unHit:Controller.Switch(34)=0:End Sub
Sub SW30_Hit:Controller.Switch(30)=1 : playsound"rollover" : End Sub
Sub SW30_unHit:Controller.Switch(30)=0:End Sub
Sub SW32a_Hit:Controller.Switch(32)=1 : playsound"rollover" : End Sub
Sub SW32a_unHit:Controller.Switch(32)=0:End Sub

Sub SW53_Hit:Controller.Switch(53)=1 : playsound"rollover" : End Sub
Sub SW53_unHit:Controller.Switch(53)=0:End Sub
Sub SW34_Hit:Controller.Switch(34)=1 : playsound"rollover" : End Sub
Sub SW34_unHit:Controller.Switch(34)=0:End Sub
Sub SW32_Hit:Controller.Switch(32)=1 : playsound"rollover" : End Sub
Sub SW32_unHit:Controller.Switch(32)=0:End Sub
Sub SW53a_Hit:Controller.Switch(53)=1 : playsound"rollover" : End Sub
Sub SW53a_unHit:Controller.Switch(53)=0:End Sub

Sub SW40_Hit:Controller.Switch(40)=1 : playsound"rollover" : End Sub
Sub SW40_unHit:Controller.Switch(40)=0:End Sub

'Scoring rubbers
 Sub SW23a_Hit:vpmTimer.PulseSw 23: Step23a=1: R23a.visible=0: R23a1.visible=1: me.timerenabled=1: End Sub
 Sub SW23b_Hit:vpmTimer.PulseSw 23: Step23b=1: R23b.visible=0: R23b1.visible=1: me.timerenabled=1: End Sub
 Sub SW23c_Hit:vpmTimer.PulseSw 23: Step23c=1: R23c.visible=0: R23c1.visible=1: me.timerenabled=1: End Sub
 Sub SW23d_Hit:vpmTimer.PulseSw 23: Step23d=1: R23d.visible=0: R23d1.visible=1: me.timerenabled=1: End Sub

 Sub SW73a_Hit:vpmTimer.PulseSw 73: Step73a=1: R73a.visible=0: R73a1.visible=1: me.timerenabled=1: End Sub
 Sub SW73b_Hit:vpmTimer.PulseSw 73: Step73b=1: R73a.visible=0: R73b1.visible=1: me.timerenabled=1: End Sub
 Sub SW73c_Hit:vpmTimer.PulseSw 73: Step73c=1: R73c.visible=0: R73c1.visible=1: me.timerenabled=1: End Sub
 Sub SW73d_Hit:vpmTimer.PulseSw 73: Step73d=1: R73c.visible=0: R73d1.visible=1: me.timerenabled=1: End Sub

'Scoring rubbers animations

sub sw23a_timer
	select case Step23a
		Case 1: r23a1.visible=0: r23a.visible=1
		Case 2: r23a.visible=0: r23a2.visible=1
		Case 3: r23a2.visible=0: r23a.visible=1: me.timerenabled=0
	end Select
	step23a=step23a+1
end sub

sub sw23b_timer
	select case Step23b
		Case 1: r23b1.visible=0: r23b.visible=1
		Case 2: r23b.visible=0: r23b2.visible=1
		Case 3: r23b2.visible=0: r23b.visible=1: me.timerenabled=0
	end Select
	step23b=step23b+1
end sub

sub sw23c_timer
	select case Step23c
		Case 1: r23c1.visible=0: R23c.visible=1
		Case 2: r23c.visible=0: r23c2.visible=1
		Case 3: r23c2.visible=0: R23c.visible=1: me.timerenabled=0
	end Select
	Step23c=step23c+1
end sub

sub sw23d_timer
	select case Step23d
		Case 1: r23d1.visible=0: r23d.visible=1
		Case 2: r23d.visible=0: r23d2.visible=1
		Case 3: r23d2.visible=0: r23d.visible=1: me.timerenabled=0
	end Select
	Step23d=Step23d+1
end sub

sub sw73a_timer
	select case Step73a
		Case 1: r73a1.visible=0: r73a.visible=1
		Case 2: r73a.visible=0: r73a2.visible=1
		Case 3: r73a2.visible=0: r73a.visible=1: me.timerenabled=0
	end Select
	step73a=step73a+1
end sub

sub sw73b_timer
	select case Step73b
		Case 1: r73b1.visible=0: r73a.visible=1
		Case 2: r73a.visible=0: r73b2.visible=1
		Case 3: r73b2.visible=0: r73a.visible=1: me.timerenabled=0
	end Select
	step73b=step73b+1
end sub

sub sw73c_timer
	select case Step73c
		Case 1: r73c1.visible=0: R73c.visible=1
		Case 2: r73c.visible=0: r73c2.visible=1
		Case 3: r73c2.visible=0: R73c.visible=1: me.timerenabled=0
	end Select
	Step73c=step73c+1
end sub

sub sw73d_timer
	select case Step73d
		Case 1: r73d1.visible=0: r73c.visible=1
		Case 2: r73c.visible=0: r73d2.visible=1
		Case 3: r73d2.visible=0: r73c.visible=1: me.timerenabled=0
	end Select
	Step73d=Step73d+1
end sub

'Kickers
 Sub Kicker1_Hit
	controller.switch(42) = 1
 End Sub

Dim lkickstep

 Sub Kicker1exit(enabled)
  if enabled Then
	controller.switch(42) = 0
	Kicker1.kick 158,7
	playsound SoundFX("Solenoid",DOFContactors)
	PkickarmL.rotz=15
	lkickstep=1
	Kicker1.timerenabled=1
  end if
end Sub

sub Kicker1_timer
	Select Case lkickstep
	  Case 2:
		PkickarmL.rotz=0
		me.timerenabled=0
	End Select
	lkickstep=lkickstep+1
end Sub

 Sub Kicker2_Hit
	controller.switch(41) = 1
 End Sub

Dim rkickstep

 Sub Kicker2exit(enabled)
  if enabled then
	controller.switch(41) = 0
	Kicker2.kick 199,7
	playsound SoundFX("Solenoid",DOFContactors)
	PkickarmR.rotz=15
	rkickstep=1
	kicker2.timerenabled=1
  end if
end Sub

sub Kicker2_timer
	Select Case rkickstep
	  Case 2:
		PkickarmR.rotz=0
		me.timerenabled=0
	End Select
	rkickstep=rkickstep+1
end Sub

'Drop Targets
	Sub Sw10_Hit
		dtLDrop.Hit 1
		Ldt10.intensityscale=1
		If BPG = 1 then
			dtRDrop.Hit 7
			Ldt74.intensityscale=1
		end if
	End Sub

 Sub Sw11_Hit
	dtLDrop.Hit 2
	dtRDrop.Hit 6
	Ldt11.intensityscale=1
	Ldt71.intensityscale=1
 End Sub

Sub Sw13_Hit
	dtLDrop.Hit 3
	Ldt13.intensityscale=1
	If BPG = 1 then
		dtRDrop.Hit 5
		Ldt70.intensityscale=1
	End if
 End Sub

 Sub Sw14_Hit
	dtLDrop.Hit 4
	dtRDrop.Hit 4
	Ldt14.intensityscale=1
	Ldt14a.intensityscale=1
	Ldt64.intensityscale=1
	Ldt64a.intensityscale=1
 End Sub

 Sub Sw20_Hit
	dtLDrop.Hit 5
	Ldt20.intensityscale=1
	If BPG = 1 then
		dtRDrop.Hit 3
		Ldt63.intensityscale=1
	End If
 End Sub

 Sub Sw21_Hit
	Ldt21.intensityscale=1
	Ldt61.intensityscale=1
	dtLDrop.Hit 6
	dtRDrop.Hit 2
 End Sub

 Sub Sw24_Hit
	Ldt24.intensityscale=1
	dtLDrop.Hit 7
	If BPG = 1 then
		dtRDrop.Hit 1
		Ldt60.intensityscale=1
	end If
 End Sub

 Sub Sw60_Hit
	Ldt60.intensityscale=1
	dtRDrop.Hit 1
	If BPG=1 then
		dtLDrop.Hit 7
		Ldt24.intensityscale=1
	End If
 End Sub

 Sub Sw61_Hit
	dtRDrop.Hit 2
	dtLDrop.Hit 6
	Ldt61.intensityscale=1
	Ldt21.intensityscale=1
 End Sub

 Sub Sw63_Hit
	dtRDrop.Hit 3
	Ldt63.intensityscale=1
	If BPG=1 then
		dtLDrop.Hit 5
		Ldt20.intensityscale=1
	End If
 End Sub

 Sub Sw64_Hit
	dtRDrop.Hit 4
	dtLDrop.Hit 4
	Ldt64.intensityscale=1
	Ldt64a.intensityscale=1
	Ldt14.intensityscale=1
	Ldt14a.intensityscale=1
 End Sub

 Sub Sw70_Hit
	dtRDrop.Hit 5
	Ldt70.intensityscale=1
	If BPG=1 then
		dtLDrop.Hit 3
		Ldt13.intensityscale=1
	End if
 End Sub

 Sub Sw71_Hit
	dtRDrop.Hit 6
	dtLDrop.Hit 2
	Ldt71.intensityscale=1
	Ldt11.intensityscale=1
 End Sub

 Sub Sw74_Hit
	dtRDrop.Hit 7
	Ldt74.intensityscale=1
	If BPG=1 then
		dtLDrop.Hit 1
		Ldt10.intensityscale=1
	End If
 End Sub

'Lamp driven drop target reset
 Dim N1,O1, Light
 N1=0:O1=0:
 Set LampCallback=GetRef("UpdateMultipleLamps")

Sub UpdateMultipleLamps
 	N1=Controller.Lamp(17) 'Lamp driven Left Bank Reset
 	If N1<>O1 Then
 		If N1 Then
			dtLDrop.DropSol_On
			DOF 111,2
			for each light in DTLeftLights: light.intensityscale=.DTshadow: Next
		End If
 	O1=N1
	End If

	N1=Controller.Lamp(1) 'Game Over triggers match and BIP
	If N1 Then
		BIPreel.SetValue 1
		NTMreel.SetValue 0
		GOBox.text=""
	  Else
		BIPreel.SetValue 0
		NTMreel.SetValue 1
		GOBox.text="GAME OVER"
	End If

	N1=Controller.Lamp(2) 'Tilt
	If N1 Then
		TILTBox.text="TILT"
	  Else
		TILTBox.text=""
	End If

	N1=Controller.Lamp(4) 'Shoot Again
	If N1 Then
		ShootAgainBox.text="SHOOT AGAIN"
	  Else
		ShootAgainBox.text=""
	End If

	N1=Controller.Lamp(3) 'HIGH SCORE TO DATE
	If N1 Then
		HStoDateBox.text="HIGH SCORE TO DATE"
	  Else
		HStoDateBox.text=""
	End If
 End Sub



'**********************************************************************************************************

'Map lights to an array
'**********************************************************************************************************
'Set Lights(1) = l1 'Game Over Backbox
'Set Lights(2) = l2 'Tilt Backbox
'Set Lights(3) = l3 'High Score Backbox
Set Lights(4) = l4 ' Shoot Again Playfied and backglass
Set Lights(5) = l5
Set Lights(6) = l6
Set Lights(7) = l7
Set Lights(8) = l8
Set Lights(9) = l9
Set Lights(10) = l10
Set Lights(11) = l11
Set Lights(12) = l12
Set Lights(13) = l13
Set Lights(14) = l14
Set Lights(15) = l15
Set Lights(16) = l16
Set Lights(19) = l19
Set Lights(20) = l20
Set Lights(21) = l21
Set Lights(22) = l22
Set Lights(23) = l23
Set Lights(24) = l24
Set Lights(25) = l25
Set Lights(26) = l26
Set Lights(27) = l27
Set Lights(28) = l28
Set Lights(29) = l29
Set Lights(30) = l30
Set Lights(31) = l31
Set Lights(32) = l32
Set Lights(33) = l33
Set Lights(34) = l34
Set Lights(35) = L35
Set Lights(36) = l36


Dim Digits(32)
Digits(0)=Array(a00,a01,a02,a03,a04,a05,a06,n,a08)
Digits(1)=Array(a10,a11,a12,a13,a14,a15,a16,n,a18)
Digits(2)=Array(a20,a21,a22,a23,a24,a25,a26,n,a28)
Digits(3)=Array(a30,a31,a32,a33,a34,a35,a36,n,a38)
Digits(4)=Array(a40,a41,a42,a43,a44,a45,a46,n,a48)
Digits(5)=Array(a50,a51,a52,a53,a54,a55,a56,n,a58)
Digits(6)=Array(b00,b01,b02,b03,b04,b05,b06,n,b08)
Digits(7)=Array(b10,b11,b12,b13,b14,b15,b16,n,b18)
Digits(8)=Array(b20,b21,b22,b23,b24,b25,b26,n,b28)
Digits(9)=Array(b30,b31,b32,b33,b34,b35,b36,n,b38)
Digits(10)=Array(b40,b41,b42,b43,b44,b45,b46,n,b48)
Digits(11)=Array(b50,b51,b52,b53,b54,b55,b56,n,b58)
Digits(12)=Array(c00,c01,c02,c03,c04,c05,c06,n,c08)
Digits(13)=Array(c10,c11,c12,c13,c14,c15,c16,n,c18)
Digits(14)=Array(c20,c21,c22,c23,c24,c25,c26,n,c28)
Digits(15)=Array(c30,c31,c32,c33,c34,c35,c36,n,c38)
Digits(16)=Array(c40,c41,c42,c43,c44,c45,c46,n,c48)
Digits(17)=Array(c50,c51,c52,c53,c54,c55,c56,n,c58)
Digits(18)=Array(d00,d01,d02,d03,d04,d05,d06,n,d08)
Digits(19)=Array(d10,d11,d12,d13,d14,d15,d16,n,d18)
Digits(20)=Array(d20,d21,d22,d23,d24,d25,d26,n,d28)
Digits(21)=Array(d30,d31,d32,d33,d34,d35,d36,n,d38)
Digits(22)=Array(d40,d41,d42,d43,d44,d45,d46,n,d48)
Digits(23)=Array(d50,d51,d52,d53,d54,d55,d56,n,d58)
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
	DipsTimer.Enabled=1
 End Sub

 Sub DipsTimer_Timer()
	BPG = TheDips(9)
	If BPG = 1 then
		instcard.image="InstCard3Balls"
	  Else
		instcard.image="InstCard5Balls"
	End if
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



'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
	vpmTimer.PulseSw 23
    PlaySound SoundFXDOF("right_slingshot",102,DOFPulse,DOFContactors), 0, 1, 0.05, 0.05
    RSling.Visible = 0
    RSling1.Visible = 1
    slingR.objroty = -15
    RStep = 1
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:slingR.objroty = -7
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:slingR.objroty = 0::RightSlingShot.TimerEnabled = 0:
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
	vpmTimer.PulseSw 23
    PlaySound SoundFXDOF("left_slingshot",101,DOFPulse,DOFContactors),0,1,-0.05,0.05
    LSling.Visible = 0
    LSling1.Visible = 1
    slingL.objroty = 15
    LStep = 1
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:slingL.objroty = 7
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:slingL.objroty = 0:LeftSlingShot.TimerEnabled = 0:
    End Select
    LStep = LStep + 1
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
Sub PinballPool_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


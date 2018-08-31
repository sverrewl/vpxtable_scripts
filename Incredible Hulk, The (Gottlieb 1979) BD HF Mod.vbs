'****************************************************************************************
'  _____ _            ___                         _ _ _     _        _   _       _ _
' |_   _| |__   ___  |_ _|_ __   ___ _ __ ___  __| (_) |__ | | ___  | | | |_   _| | | __
'   | | | '_ \ / _ \  | || '_ \ / __| '__/ _ \/ _` | | '_ \| |/ _ \ | |_| | | | | | |/ /
'   | | | | | |  __/  | || | | | (__| | |  __/ (_| | | |_) | |  __/ |  _  | |_| | |   <
'   |_| |_| |_|\___| |___|_| |_|\___|_|  \___|\__,_|_|_.__/|_|\___| |_| |_|\__,_|_|_|\_\
'
'                      Incredible Hulk, The (Gottlieb 1979)
'****************************************************************************************
'Credits:
'32assassin (original table)
'BorgDog (mods, lighting, code)
'HauntFreaks (mods, lighting, primitives, graphics)

' Thalamus 2018-07-23
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' No special SSF tweaks yet.
' Wob 2018-08-08
' Added vpmInit Me to table init

Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="hulk",UseSolenoids=2,UseLamps=1,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff",SCoin="coin"

LoadVPM"01120100","GTS1.VBS",3.02
Dim DesktopMode: DesktopMode = HULK.ShowDT

Dim DTshadow: DTshadow=0.3    'set drop target shadow level from 0 to 1, smaller number is more shadow

'*************************************************************

'Solenoid Call backs
'**********************************************************************************************************
SolCallback(1)="bsTrough.SolOut"
SolCallback(2)="vpmSolSound ""Knocker"","
SolCallback(3)="vpmSolSound ""Chime1"","
SolCallback(4)="vpmSolSound ""Chime2"","
SolCallback(5)="vpmSolSound ""Chime3"","
SolCallback(6)=	"kicker1exit"				'"bsSaucer1.SolOut"
SolCallback(7)=	"kicker2exit"				'"bsSaucer2.SolOut"
SolCallback(8)="Rraised"   					'"dtDrop.SolDropUp"
SolCallback(17)="vpmNudge.SolGameOn"

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub Rraised(enabled)
	if enabled then Rreset.enabled=True
End Sub

Sub Rreset_timer
	Dim light
	dtDrop.DropSol_On
	For each light in DTLights: light.intensityscale=DTshadow: Next
	Rreset.enabled=false
End Sub

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySound"fx_Flipperup":LeftFlipper.RotateToEnd
     Else
         PlaySound "fx_Flipperdown":LeftFlipper.RotateToStart
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySound"fx_Flipperup":RightFlipper.RotateToEnd
     Else
         PlaySound "fx_Flipperdown":RightFlipper.RotateToStart
     End If
End Sub
'**********************************************************************************************************
'**********************************************************************************************************
'Primitive Flipper Code

Sub FlipperTimer_Timer
	dim PI:PI=3.1415926
	FlipperT1.roty = LeftFlipper.currentangle
	FlipperT5.roty = RightFlipper.currentangle
	Pgate.rotz=Gate2.currentangle+25
	SpinnerP.Rotz = sw24.CurrentAngle
	SpinnerRod.TransZ = sin( (sw24.CurrentAngle+180) * (2*PI/360)) * 5
	SpinnerRod.TransX = -1*(sin( (sw24.CurrentAngle- 90) * (2*PI/360)) * 5)
End Sub


'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************
Dim bsTrough,dtDrop,bsSaucer1,bsSaucer2

Sub HULK_Init
	vpmInit Me
	On Error Resume Next
	With Controller
		.GameName=cGameName
		If Err Then MsgBox"Can't start Game"&cGameName&vbNewLine&Err.Description:Exit Sub
		.SplashInfoLine="The Incredible Hulk Gottlieb"
		.HandleMechanics=0
		.ShowDMDOnly=1
		.ShowFrame=0
		.ShowTitle=0
 		.Hidden=1
		.Run
		If Err Then MsgBox Err.Description
	End With
	On Error Goto 0

	PinMAMETimer.Interval=PinMAMEInterval
	PinMAMETimer.Enabled=1
	vpmNudge.TiltSwitch=4
	vpmNudge.Sensitivity=3
    vpmNudge.TiltObj=Array(Bumper1,Bumper2)

	Set bsTrough=New cvpmBallStack
	bsTrough.InitSw 0,66,0,0,0,0,0,0
	bsTrough.InitKick BallRelease,90,6
	bsTrough.InitExitSnd"BallRelease","Solenoid"
 	bsTrough.Balls=1

 	Set dtDrop=New cvpmDropTarget
	dtDrop.InitDrop Array(sw34,sw32,sw31,sw22,sw21,sw12,sw11),Array(34,32,31,22,21,12,11)
	dtDrop.InitSnd"DTDrop","DTReset"

	'find balls per game
	FindDips

	dim xx

	For each xx in GI:xx.State = 1: Next
	For each xx in DTLights: xx.state=1: xx.intensityscale=DTshadow:Next

	If b2son Then
		for each xx in backdropstuff
			xx.visible=0
		Next
	end if

End Sub
'**********************************************************************************************************
'**********************************************************************************************************

'Plunger code
'**********************************************************************************************************
Sub HULK_KeyDown(ByVal keycode)
	If keycode = 61 then FindDips
	If vpmKeyDown(KeyCode) Then Exit Sub
	If keycode=AddCreditKey then playsound "coin": vpmTimer.pulseSW (swCoin1): end if
	If keycode=PlungerKey Then Plunger.Pullback:playsound"plungerpull"
End Sub

Sub HULK_KeyUp(ByVal keycode)
	If vpmKeyUp(KeyCode) Then Exit Sub
	If keycode=PlungerKey Then Plunger.Fire:playsound"plunger"
End Sub
'**********************************************************************************************************

'Switches
'**********************************************************************************************************
'**********************************************************************************************************

'Drain
Sub Drain_Hit:PlaySound "Drain":bsTrough.AddBall Me:End Sub
'Kickers

 Sub Kicker1_Hit
	controller.switch(41) = 1
 End Sub

Dim lkickstep

 Sub Kicker1exit(enabled)
  if enabled Then
	controller.switch(41) = 0
	Kicker1.kick -5,30
	playsound SoundFX("Solenoid",DOFContactors)
	PkickarmL.objrotx=15
	lkickstep=1
	Kicker1.timerenabled=1
  end if
end Sub

sub Kicker1_timer
	Select Case lkickstep
	  Case 3:
		PkickarmL.objrotx=7
	  Case 4:
		PkickarmL.objrotx=0
		me.timerenabled=0
	End Select
	lkickstep=lkickstep+1
end Sub

 Sub Kicker2_Hit
	controller.switch(42) = 1
 End Sub

Dim rkickstep

 Sub Kicker2exit(enabled)
  if enabled then
	controller.switch(42) = 0
	Kicker2.kick 0,30
	playsound SoundFX("Solenoid",DOFContactors)
	PkickarmR.objrotx=15
	rkickstep=1
	kicker2.timerenabled=1
  end if
end Sub

sub Kicker2_timer
	Select Case rkickstep
	  Case 3:
		PkickarmR.objrotx=7
	  Case 4:
		PkickarmR.objrotx=0
		me.timerenabled=0
	End Select
	rkickstep=rkickstep+1
end Sub

'Wire triggers
Sub sw41_Hit:Controller.Switch(44)=1 : playsound"rollover" : End Sub
Sub sw41_Unhit:Controller.Switch(44)=0:End Sub
Sub sw60_Hit:Controller.Switch(60)=1 : playsound"rollover" : End Sub
Sub sw60_Unhit:Controller.Switch(60)=0:End Sub
Sub sw61_Hit:Controller.Switch(61)=1 : playsound"rollover" : End Sub
Sub sw61_Unhit:Controller.Switch(61)=0:End Sub
Sub sw62_Hit:Controller.Switch(62)=1 : playsound"rollover" : End Sub
Sub sw62_UnHit:Controller.Switch(62)=0:End Sub
Sub sw64_Hit:Controller.Switch(64)=1 : playsound"rollover" : End Sub
Sub sw64_UnHit:Controller.Switch(64)=0:End Sub
Sub sw70_Hit:Controller.Switch(70)=1 : playsound"rollover" : End Sub
Sub sw70_UnHit:Controller.Switch(70)=0:End Sub
Sub sw71_Hit:Controller.Switch(71)=1 : playsound"rollover" : End Sub
Sub sw71_UnHit:Controller.Switch(71)=0:End Sub
Sub sw72_Hit:Controller.Switch(72)=1 : playsound"rollover" : End Sub
Sub sw72_UnHit:Controller.Switch(72)=0:End Sub
Sub sw74_Hit:Controller.Switch(74)=1 : playsound"rollover" : End Sub
Sub sw74_UnHit:Controller.Switch(74)=0:End Sub

'Standup Targets
Sub sw10_Hit:vpmTimer.PulseSw 10 : End Sub
Sub sw20_Hit:vpmTimer.PulseSw 20 : End Sub
Sub sw30_Hit:vpmTimer.PulseSw 30 : End Sub

'Drop Target
 Sub Sw34_Dropped: dtDrop.Hit 1: LDT34.intensityscale=1: End Sub
 Sub Sw32_Dropped: dtDrop.Hit 2: LDT32.intensityscale=1: End Sub
 Sub Sw31_Dropped: dtDrop.Hit 3: LDT31.intensityscale=1: LDT31a.intensityscale=1: End Sub
 Sub Sw22_Dropped: dtDrop.Hit 4: LDT22.intensityscale=1: LDT22a.intensityscale=1: End Sub
 Sub Sw21_Dropped: dtDrop.Hit 5: LDT21.intensityscale=1: End Sub
 Sub Sw11_Dropped: dtDrop.Hit 6: LDT11.intensityscale=1: End Sub
 Sub Sw12_Dropped: dtDrop.Hit 7: LDT12.intensityscale=1: End Sub

 'Spinner
Sub sw24_Spin:vpmTimer.PulseSw 24 : playsound"fx_spinner" : End Sub

'Star Rollover
Sub Trigger1_Hit:controller.switch(14)= 1 :playsound"rollover" : End Sub
Sub Trigger1_Unhit:controller.switch(14)= 0 :End Sub
Sub Trigger2_Hit:controller.switch(14)= 1: playsound"rollover" : End Sub
Sub Trigger2_UnHit:controller.switch(14)= 0 :End Sub
Sub Trigger3_Hit:controller.switch(14)= 1: playsound"rollover" : End Sub
Sub Trigger3_UnHit:controller.switch(14)= 0 :End Sub
Sub Trigger4_Hit:controller.switch(14)=1 : playsound"rollover" : End Sub
Sub Trigger4_UnHit:controller.switch(14)= 0 :End Sub

Dim Step1, Step2, Step3, Step4, Step5

'Scoring Rubbers (5)
 Sub sr1_Hit:vpmTimer.PulseSw 40: playsound"rubber_hit_1" : Step1=1: sr1R.visible=0: sr1R1.visible=1: me.timerenabled=1: End Sub
 Sub sr2_Hit:vpmTimer.PulseSw 40: playsound"rubber_hit_1" : Step2=1: sr2R.visible=0: sr2R1.visible=1: me.timerenabled=1: End Sub
 Sub sr3_Hit:vpmTimer.PulseSw 40: playsound"rubber_hit_1" : Step3=1: sr3R.visible=0: sr3R1.visible=1: me.timerenabled=1: End Sub
 Sub sr4_Hit:vpmTimer.PulseSw 40: playsound"rubber_hit_1" : Step4=1: sr4R.visible=0: sr4R1.visible=1: me.timerenabled=1: End Sub
 Sub sr5_Hit:vpmTimer.PulseSw 40: playsound"rubber_hit_1" : Step5=1: sr5R.visible=0: sr5R1.visible=1: me.timerenabled=1: End Sub

'Scoring rubbers animations

sub sr1_timer
	select case Step1
		Case 1: sr1r1.visible=0: sr1r.visible=1
		Case 2: sr1r.visible=0: sr1r2.visible=1
		Case 3: sr1r2.visible=0: sr1r.visible=1: me.timerenabled=0
	end Select
	Step1=Step1+1
end sub

sub sr2_timer
	select case Step2
		Case 1: sr2r1.visible=0: sr2r.visible=1
		Case 2: sr2r.visible=0: sr2r2.visible=1
		Case 3: sr2r2.visible=0: sr2r.visible=1: me.timerenabled=0
	end Select
	Step2=Step2+1
end sub

sub sr3_timer
	select case Step3
		Case 1: sr3r1.visible=0: sr3r.visible=1
		Case 2: sr3r.visible=0: sr3r2.visible=1
		Case 3: sr3r2.visible=0: sr3r.visible=1: me.timerenabled=0
	end Select
	Step3=Step3+1
end sub

sub sr4_timer
	select case Step4
		Case 1: sr4r1.visible=0: sr4r.visible=1
		Case 2: sr4r.visible=0: sr4r2.visible=1
		Case 3: sr4r2.visible=0: sr4r.visible=1: me.timerenabled=0
	end Select
	Step4=Step4+1
end sub

sub sr5_timer
	select case Step5
		Case 1: sr5r1.visible=0: sr5r.visible=1
		Case 2: sr5r.visible=0: sr5r2.visible=1
		Case 3: sr5r2.visible=0: sr5r.visible=1: me.timerenabled=0
	end Select
	Step5=Step5+1
end sub

'Pop Bumpers

Sub Bumper1_Hit
	vpmTimer.PulseSw 50
	DOF 101,2
	playsound SoundFX("fx_bumper1",DOFContactors)
	B1L1.State = 1:B1L2. State = 1
	me.timerenabled=1
End Sub

Sub Bumper1_timer
	BumperTimerRing1.Enabled=0
	if bumperring1.transz>-36 then 	BumperRing1.transz=BumperRing1.transz-4
	if BumperRing1.transz=-36 then
		BumperTimerRing1.enabled=1
		me.timerenabled=0
	end if
End Sub

Sub BumperTimerRing1_timer
	if bumperring1.transz<0 then BumperRing1.transz=BumperRing1.transz+4
	If BumperRing1.transz=0 then
		B1L1.State = 0:B1L2. State = 0
		BumperTimerRing1.enabled=0
	end if
End sub

Sub Bumper2_Hit
	vpmTimer.PulseSw 50
	DOF 102,2
	playsound SoundFX("fx_bumper1",DOFContactors)
	B2L1.State = 1:B2L2. State = 1
	me.timerenabled=1
End Sub

Sub Bumper2_timer
	BumperTimerRing2.enabled=0
	if BumperRing2.transz>-36 then 	BumperRing2.transz=BumperRing2.transz-4
	if BumperRing2.transz=-36 then
		BumperTimerRing2.enabled=1
		me.timerenabled=0
	end if
End Sub

Sub BumperTimerRing2_timer
	if BumperRing2.transz<0 then BumperRing2.transz=BumperRing2.transz+4
	If BumperRing2.transz=0 then
		B2L1.State = 0:B2L2. State = 0
		BumperTimerRing2.enabled=0
	End If
End sub

'Backdrop light controls
 Dim N1

 Set LampCallback=GetRef("UpdateMultipleLamps")

Sub UpdateMultipleLamps

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
'Set Lights(1)=Light1 'Game Over
'Set Lights(2)=Light2 'Tilt
'Set Lights(3)=Light3 'High Score
Set Lights(4)=Light4
Set Lights(5)=Light5
Set Lights(7)=Light7
Set Lights(8)=Light8
Set Lights(9)=Light9
Set Lights(10)=Light10
Set Lights(11)=Light11
Set Lights(12)=Light12
Set Lights(13)=Light13
Set Lights(14)=Light14
Set Lights(15)=Light15
Set Lights(16)=Light16
Set Lights(17)=Light17
Set Lights(18)=Light18
Set Lights(19)=Light19
Set Lights(20)=Light20
Set Lights(21)=Light21
Set Lights(22)=Light22
Set Lights(23)=Light23
Set Lights(24)=Light24
Set Lights(25)=Light25
Set Lights(26)=Light26
Set Lights(27)=Light27
Set Lights(28)=Light28
Set Lights(29)=Light29
Set Lights(30)=Light30
Set Lights(31)=Light31
Set Lights(32)=Light32
Set Lights(33)=Light33
Set Lights(34)=Light34
Set Lights(35)=Light35
Set Lights(36)=Light36

'**********************************************************************************************************

'backglass lamps
'**********************************************************************************************************
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
		If DesktopMode = True Then
		For ii = 0 To UBound(chgLED)
			num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
			if (num < 32) then
				For Each obj In Digits(num)
					If chg And 1 Then obj.State = stat And 1
					chg = chg\2 : stat = stat\2
				Next
			else
				'if char(stat) > "" then msg(num) = char(stat)
			end if
		next
		end if
end if
End Sub



'**********************************************************************************************************

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

Dim BPG

 Sub DipsTimer_Timer()
	BPG = TheDips(9)
	If BPG = 1 then
		instcard.image="InstCard3Balls"
	  Else
		instcard.image="InstCard5Balls"
	End if
	DipsTimer.enabled=0
 End Sub


Sub editDips
  	Dim vpmDips : Set vpmDips = New cvpmDips
  	With vpmDips
  		.AddForm 700,400,"System 1 (Multi-Mode sound) - DIP switches"
  		.AddFrame 0,0,190,"Coin chute control",&H00040000,Array("seperate",0,"same",&H00040000)'dip 19
  		.AddFrame 0,46,190,"Game mode",&H00000400,Array("extra ball",0,"replay",&H00000400)'dip 11
  		.AddFrame 0,92,190,"High game to date awards",&H00200000,Array("no award",0,"3 replays",&H00200000)'dip 22
  		.AddFrame 0,138,190,"Balls per game",&H00000100,Array("5 balls",0,"3 balls",&H00000100)'dip 9
  		.AddFrame 0,184,190,"Tilt effect",&H00000800,Array("game over",0,"ball in play only",&H00000800)'dip 12
  		.AddFrame 205,0,190,"Maximum credits",&H00030000,Array("5 credits",0,"8 credits",&H00020000,"10 credits",&H00010000,"15 credits",&H00030000)'dip 17&18
  		.AddFrame 205,76,190,"Sound settings",&H80000000,Array("sounds",0,"tones",&H80000000)'dip 32
  		.AddFrame 205,122,190,"Attract tune",&H10000000,Array("no attract tune",0,"attract tune played every 6 minutes",&H10000000)'dip 29
  		.AddChk 205,175,190,Array("Match feature",&H00000200)'dip 10
  		.AddChk 205,190,190,Array("Credits displayed",&H00001000)'dip 13
  		.AddChk 205,205,190,Array("Play credit button tune",&H00002000)'dip 14
  		.AddChk 205,220,190,Array("Play tones when scoring",&H00080000)'dip 20
  		.AddChk 205,235,190,Array("Play coin switch tune",&H00400000)'dip 23
  		.AddChk 205,250,190,Array("High game to date displayed",&H00100000)'dip 21
  		.AddLabel 50,280,300,20,"After hitting OK, press F3 to reset game with new settings."
  		.ViewDips
  	End With
 End Sub
 Set vpmShowDips = GetRef("editDips")

'****************************************************************************
'************************************************************************
'*****GI Lights On
dim xx
For each xx in GI:xx.State = 1: Next

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

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "HULK" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / HULK.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "HULK" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / HULK.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "HULK" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / HULK.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioFade(ball) ' Can this be together with the above function ?
  Dim tmp
  tmp = ball.y * 2 / HULK.height-1
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
  If HULK.VersionMinor > 3 OR HULK.VersionMajor > 10 Then
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
  Else
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0
  End if
End Sub

' Thalamus : Exit in a clean and proper way
Sub HULK_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


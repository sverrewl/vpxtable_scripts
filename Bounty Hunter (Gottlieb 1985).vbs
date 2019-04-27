Option Explicit

Const cGameName="bountyh",UseSolenoids=2,UseLamps=1,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff",SFlipperOn="FlipperDown",SFlipperOff="FlipperUp"
Const SCoin="coin3",cCredits=""

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01120100","sys80.vbs",3.02

Dim VarRol
If Table1.ShowDT = true then VarRol=0 Else VarRol=1

Set LampCallback=GetRef("UpdateMultipleLamps")

Sub UpdateMultipleLamps
End Sub

Sub Table1_KeyDown(ByVal keycode)
	If KeyCode=LeftFlipperKey Then Controller.Switch(6)=1
	If keyCode=RightFlipperKey Then Controller.Switch(16)=1
	If vpmKeyDown(KeyCode) Then Exit Sub
	If keycode=PlungerKey Then Plunger.Pullback
End Sub

Sub Table1_KeyUp(ByVal keycode)
	If KeyCode=LeftFlipperKey Then Controller.Switch(6)=0
	If keyCode=RightFlipperKey Then Controller.Switch(16)=0
	If vpmKeyUp(KeyCode) Then Exit Sub
	If keycode=PlungerKey Then Plunger.Fire:playsound "plunger"
End Sub

Const sDropDown=1
Const sKicker=2
Const sDropUp=5
Const sknocker=8
Const sOutHole=9

SolCallback(sKicker)="bsSaucer.SolOut"
SolCallback(sKnocker)="VpmSolSound SoundFX(""knocker"",DOFKnocker),"
SolCallback(sDropDown)="SolDown"
SolCallback(sDropUp)="SolDro"
SolCallback(sOutHole)="bsTrough.SolOut"
'SolCallback(sLLFlipper)="VpmSolFlipper LeftFlipper,nothing,"
'SolCallback(sLRFlipper)="VpmSolFlipper RightFlipper,Flipper1,"

 SolCallback(sLRFlipper) = "SolRFlipper"
 SolCallback(sLLFlipper) = "SolLFlipper"

	Sub SolLFlipper(Enabled)
		If Enabled Then
		PlaySound SoundFX("fx_flipperup",DOFFlippers)
		LeftFlipper.RotateToEnd
     Else
		 PlaySound SoundFX("fx_flipperdown",DOFFlippers)
		 LeftFlipper.RotateToStart
     End If
	End Sub

	Sub SolRFlipper(Enabled)
     If Enabled Then
		PlaySound SoundFX("fx_flipperup",DOFFlippers)
		 RightFlipper.RotateToEnd
		 Flipper1.RotateToEnd
     Else
		 PlaySound SoundFX("fx_flipperdown",DOFFlippers)
		 RightFlipper.RotateToStart
		 Flipper1.RotateToStart
     End If
	End Sub

Sub SolDro(Enabled)
	If Enabled Then
		Wall4.IsDropped=0
		Controller.Switch(74)=0
	End If
End Sub

Sub SolDown(Enabled)
	If Enabled Then
		Wall4.IsDropped=1
		Controller.Switch(74)=1
	End If
End Sub

Dim bsTrough,bsSaucer

Sub Table1_Init
	On Error Resume Next
	With Controller 
		.GameName=cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
		.SplashInfoLine="Bounty Hunter, Gottlieb/Premier 1985."
		.HandleMechanics=0
		.HandleKeyboard=0
		.ShowDMDOnly=1
		.ShowFrame=1
		.Games(cGameName).Settings.Value("rol") = VarRol
		.ShowTitle=0
		If Err Then MsgBox Err.Description
	End With
	On Error Goto 0
		Controller.SolMask(0)=0
      vpmTimer.AddTimer 2000,"Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the timer to renable all the solenoids after 2 seconds
		Controller.Run


'	Controller.Dip(0) = (0*1 + 0*2 + 0*4 + 0*8 + 0*16 + 0*32 + 0*64 + 1*128) '01-08
'	Controller.Dip(1) = (0*1 + 0*2 + 0*4 + 0*8 + 0*16 + 0*32 + 1*64 + 1*128) '09-16
'	Controller.Dip(2) = (0*1 + 0*2 + 0*4 + 0*8 + 0*16 + 1*32 + 1*64 + 1*128) '17-24
'	Controller.Dip(3) = (1*1 + 1*2 + 0*4 + 0*8 + 1*16 + 0*32 + 0*64 + 0*128) '25-32

'Switch 25 Number Of Balls
'    ON = 3
'   OFF = 5

	PinMAMETimer.Interval=PinMAMEInterval
	PinMAMETimer.Enabled=1
	
	vpmNudge.TiltSwitch=57
	vpmNudge.Sensitivity=1
	vpmNudge.TiltObj=Array(Bumper1,Bumper2,LeftSlingshot,RightSlingshot,Wall5,Wall6)
	
	Set bsTrough=New cvpmBallstack
	bsTrough.InitSw 0,67,0,0,0,0,0,0
	bsTrough.InitKick BallRelease,90,6
	bsTrough.InitExitSnd SoundFX("ballrel",DOFContactors),SoundFX("solon",DOFContactors)
	bsTrough.Balls=1
	
	Set bsSaucer=New cvpmBallStack
	bsSaucer.InitSaucer Kicker1,73,147,5
	bsSaucer.InitExitSnd SoundFX("popper",DOFContactors),SoundFX("popper",DOFContactors)

	Wall4.IsDropped=1
	Controller.Switch(74)=1
End Sub

Sub Trigger4_Hit:Controller.Switch(40)=1:End Sub				'switch 40
Sub Trigger4_unHit:Controller.switch(40)=0:End Sub				'switch 40
Sub Trigger5_Hit:Controller.Switch(41)=1:End Sub				'switch 41
Sub Trigger5_unHit:Controller.switch(41)=0:End Sub				'switch 41
Sub Trigger6_Hit:Controller.Switch(42)=1:End Sub				'switch 42
Sub Trigger6_unHit:Controller.switch(42)=0:End Sub				'switch 42
Sub Bumper1_Hit()Playsound SoundFX("bump",DOFContactors) :vpmTimer.PulseSw(43):Light1.State=1:Light1.TimerEnabled=True:DOF 103, DOFPulse:End Sub					'switch 43
Sub Light1_Timer:Light1.TimerEnabled=False:Light1.State=0:End Sub
Sub Bumper2_Hit()Playsound SoundFX("bump",DOFContactors) :vpmTimer.PulseSw(43):Light2.State=1:Light2.TimerEnabled=True:DOF 104, DOFPulse:End Sub					'switch 43
Sub Light2_Timer:Light2.TimerEnabled=False:Light2.State=0:End Sub
Sub Wall5_Slingshot:PlaySound SoundFX("sling",DOFContactors):VpmTimer.PulseSw 44:DOF 106, DOFPulse:End Sub			'switch 44
Sub Wall6_Slingshot:PlaySound SoundFX("sling",DOFContactors):VpmTimer.PulseSw 44:DOF 107, DOFPulse:End Sub			'switch 44
Sub LeftSlingShot_Slingshot:PlaySound SoundFX("sling",DOFContactors):VpmTimer.PulseSw 44:DOF 101, DOFPulse:End Sub 'switch 44
Sub RightSlingshot_Slingshot:PlaySound SoundFX("sling",DOFContactors):VpmTimer.PulseSw 44:DOF 102, DOFPulse:End Sub 'switch 44
Sub Wall1_Hit:VpmTimer.PulseSw 50:End Sub						'switch 50
Sub Wall7_Hit:vpmTimer.PulseSw 51:End Sub						'switch 51
Sub Trigger1_Hit:Controller.Switch(52)=1:End Sub				'switch 52
Sub Trigger1_unHit:Controller.switch(52)=0:End Sub				'switch 52
Sub RightInlane_Hit:Controller.Switch(53)=1:End Sub				'switch 53
Sub RightInlane_UnHit:Controller.Switch(53)=0:End Sub 			'switch 53
Sub LeftInlane_Hit:Controller.Switch(54)=1:End Sub				'switch 54
Sub LeftInlane_unHit:Controller.Switch(54)=0:End Sub			'switch 54
Sub Wall2_Hit:VpmTimer.PulseSw 60:End Sub						'switch 60
Sub Wall8_Hit:vpmTimer.PulseSw 61:End Sub						'switch 61
Sub Trigger2_Hit:Controller.Switch(62)=1:End Sub				'switch 62
Sub Trigger2_unHit:Controller.switch(62)=0:End Sub				'switch 62
Sub RightOutlane_Hit:Controller.Switch(63)=1:End Sub			'switch 63
Sub RightOutlane_UnHit:Controller.Switch(63)=0:End Sub			'switch 63
Sub LeftOutlane_Hit:Controller.Switch(64)=1:End Sub				'switch 64
Sub LetOutlane_unHit:Controller.Switch(64)=0:End Sub			'switch 64
Sub Drain_Hit()Playsound "drain" :bsTrough.AddBall Me:End Sub						'switch 67
Sub Wall3_Hit:VpmTimer.PulseSw 70:End Sub						'switch 70
Sub Wall9_Hit:vpmTimer.PulseSw 71:End Sub						'switch 71
Sub Trigger3_Hit:Controller.Switch(72)=1:End Sub				'switch 72
Sub Trigger3_unHit:Controller.switch(72)=0:End Sub				'switch 72
Sub Kicker1_Hit:bsSaucer.AddBall 0:End Sub						'switch 73
Sub Wall4_Hit:Wall4.IsDropped=1:Controller.Switch(74)=1:PlaySound"flapclos":End Sub'switch 74

set lights(3)=Light3
Set Lights(6)=Light6
Set Lights(7)=Light7
Set Lights(8)=Light8
Set Lights(9)=Light9
Set Lights(10)=Light10
Set lights(11)=light11
set lights(12)=light12
lights(13)=Array(light13,light13B)
set lights(14)=light14
set lights(15)=light15
set lights(16)=light16
lights(17)=Array(light17,light17B)
set lights(18)=light18
set lights(19)=light19
set lights(20)=light20
set lights(21)=light21
set lights(22)=light22
set lights(23)=light23
lights(24)=Array(light24,light24B)
set lights(25)=light25
lights(26)=Array(light26,light26B)
set lights(27)=light27
set lights(28)=light28
set lights(29)=light29
set lights(30)=light30
set lights(31)=light31
set lights(32)=light32
set lights(33)=light33
set lights(34)=light34
set lights(35)=light35
set lights(36)=light36
set lights(37)=light37
set lights(38)=light38
set lights(39)=light39
set lights(40)=light40
set lights(41)=light41
set lights(42)=light42
set lights(43)=light43
set lights(44)=light44
set lights(45)=light45
set lights(46)=light46
lights(47)=Array(light47, light48, light49, light50, light51, light52, light53, light54, light55, light56, light57, light58, light59, light60, light61, light62, light63)

Sub Gate1_Hit: PlaySound "Gate" : End Sub
Sub Gate_Hit: PlaySound "Gate" : End Sub
Sub Trigger7_Hit: PlaySound "PlungeBall" : End Sub

'Gottlieb System 80A Sound only (S)board
'added by Inkochnito
Sub editDips
	Dim vpmDips : Set vpmDips = New cvpmDips
	With vpmDips
		.AddForm 700,400,"System 80A with Sound only (S)board - DIP switches"
		.AddFrame 0,0,190,"Maximum credits",49152,Array("8 credits",0,"10 credits",32768,"15 credits",&H00004000,"25 credits",49152)'dip 15&16
		.AddFrame 0,76,190,"Coin chute 1 and 2 control",&H00002000,Array("seperate",0,"same",&H00002000)'dip 14
		.AddFrame 0,122,190,"3rd coin chute credits control",&H20000000,Array("no effect",0,"add 9",&H20000000)'dip 30
		.AddFrame 205,0,190,"High score to date awards",&H00C00000,Array("not displayed and no award",0,"displayed and no award",&H00800000,"displayed and 2 credits",&H00400000,"displayed and 3 credits",&H00C00000)'dip 23&24
		.AddFrame 205,76,190,"Game mode",&H10000000,Array("replay",0,"extra ball",&H10000000)'dip 29
		.AddFrame 205,122,190,"Playfield special",&H00200000,Array("replay",0,"extra ball",&H00200000)'dip 22
		.AddFrame 205,168,190,"Replay limit",&H04000000,Array("no limit",0,"one per game",&H04000000)'dip 27
		.AddFrame 205,214,190,"Balls per game",&H01000000,Array("5 balls",0,"3 balls",&H01000000)'dip 25
		.AddFrame 205,260,190,"Novelty mode",&H08000000,Array("normal game mode",0,"50K per special/extra ball",&H08000000)'dip 28
		.AddChk 0,170,190,Array("Match feature",&H02000000)'dip 26
		.AddChk 0,185,190,Array("Background sound",&H40000000)'dip 31
		.AddChk 0,200,190,Array("Dip 6 (spare)",&H00000020)'dip 6
		.AddChk 0,215,190,Array("Dip 7 (spare)",&H00000040)'dip 7
		.AddChk 0,230,190,Array("Dip 8 (spare)",&H00000080)'dip 8
		.AddChk 0,245,190,Array("Dip 32 (spare or game option)",&H80000000)'dip 32
		.AddLabel 50,310,300,20,"After hitting OK, press F3 to reset game with new settings."
		.ViewDips
	End With
End Sub
Set vpmShowDips = GetRef("editDips")
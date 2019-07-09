Option Explicit
Randomize

Const cGameName="icefever"

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01210000", "sys80.VBS", 3.1

' Options
' Volume devided by - lower gets higher sound

Const VolDiv = 2000    ' Lower number, louder ballrolling/collition sound
Const VolCol = 10      ' Ball collition divider ( voldiv/volcol )


Dim VarRol, VarHidden
VarHidden=1
If Table1.ShowDT = true then VarRol=0 Else VarRol=1

If B2SOn = true Then VarHidden=1

Const UseSolenoids=2,UseLamps=1,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff",SFlipperOn="FlipperUp",SFlipperOff="FlipperDown"
Const SCoin="coin3",cCredits=""


'**********************************************************
'********       OPTIONS     *******************************
'**********************************************************

Dim BallShadows: Ballshadows=1          	'******************set to 1 to turn on Ball shadows
Dim FlipperShadows: FlipperShadows=1  		'***********set to 1 to turn on Flipper shadows
Dim ROMSounds: ROMSounds=1					'**********set to 0 for no rom sounds, 1 to play rom sounds.. mostly used for testing
Dim CaptiveBallMass:  CaptiveBallMass=2		'******* play around with to make it harder to score goal with captive ball, higher number should be harder to move	
Dim UnicornFart: UnicornFart=1				'******* set to 1 for color GI, set to 0 for normal GI

Sub Table1_KeyDown(ByVal keycode)
	If vpmKeyDown(KeyCode) Then Exit Sub
	If keycode=PlungerKey Then Plunger.Pullback

End Sub

Sub Table1_KeyUp(ByVal keycode)
	If vpmKeyUp(KeyCode) Then Exit Sub
	If keycode=PlungerKey Then
		Plunger.Fire
		if ballhome.BallCntOver>0 then
			PlaySoundAt "plungerreleaseball", Plunger	'PLAY WHEN BALL IS HIT
		  else
			PlaySoundAt "plungerreleasefree", Plunger		'PLAY WHEN NO BALL TO PLUNGE
		end if
	end If


End Sub


Const sDrop=2
Const sPuck=5' backglass hockey puck
Const sknocker=8
Const sOutHole=9

SolCallback(sDrop)="dtL.SolDropUp"
SolCallback(sKnocker)="VpmSolSound SoundFX(""knocker"",DOFKnocker),"
SolCallback(sOutHole)="bsTrough.SolOut"
'SolCallback(sLLFlipper)="VpmSolFlipper LeftFlipper,nothing,"
'SolCallback(sLRFlipper)="VpmSolFlipper RightFlipper,nothing,"
SolCallback(sPuck)="AnimationStart"

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
	If Enabled Then
		PlaySound SoundFX("FlipperUp",DOFFlippers)
		LeftFlipper.RotateToEnd
     Else
		 PlaySound SoundFX("FlipperDown",DOFFlippers)
		 LeftFlipper.RotateToStart
     End If
End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
		PlaySound SoundFX("FlipperUp",DOFFlippers)
		 RightFlipper.RotateToEnd
     Else
		 PlaySound SoundFX("FlipperDown",DOFFlippers)
		 RightFlipper.RotateToStart
     End If
End Sub


Sub FlipperTimer_Timer
	Dim PI: PI=3.1415926

	LFlip.RotZ = LeftFlipper.CurrentAngle
	RFlip.RotZ = RightFlipper.CurrentAngle
	LflipRubber.RotZ = LeftFlipper.CurrentAngle
	RFlipRubber.RotZ = RightFlipper.CurrentAngle

	Pgate.Rotz = Gate1.CurrentAngle*0.7
	PrimGate2.rotx = -Gate2.currentangle*0.5
	PrimGate3.rotx = -Gate3.currentangle*0.5

	Dim SpinnerRadius: SpinnerRadius=7

	SpinnerRod.TransZ = (cos((spinner.CurrentAngle + 180) * (PI/180))+1) * SpinnerRadius
	SpinnerRod.TransY = (sin((spinner.CurrentAngle) * (PI/180))) * -SpinnerRadius

	L21.state=light21.state

	if FlipperShadows=1 then
		FlipperLSh.RotZ = LeftFlipper.currentangle
		FlipperRSh.RotZ = RightFlipper.currentangle
	end if
End Sub

Sub AnimationStart(Enabled)
	if Table1.showDT and not GoalTimer.enabled then 
		GoalTimer.uservalue=0
		GoalTimer.enabled=1
	end if
End Sub

Sub GoalTimer_timer
	if me.uservalue>0 then GoalLights(me.uservalue-1).state=0
	GoalLights(me.uservalue).state=1
	me.uservalue=me.uservalue+1
	if me.uservalue>20 Then
		GoalLights(20).state=0
		me.enabled=0
	end if
end Sub


Dim bsTrough,dtL, objekt, Light

Sub Table1_Init
	On Error Resume Next

	With Controller
		.GameName=cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
		.SplashInfoLine="Ice Fever, Gottlieb 1985" & vbNewLine & "by Tom, BorgDog, Bord, Thalamus and Arngrim" & vbNewLine & "VPX conversion by Rascal"
		.HandleMechanics=0
		.HandleKeyboard=0
		.ShowDMDOnly=1
		.ShowFrame=0
		.ShowTitle=0
		.Hidden = VarHidden
		.Games(cGameName).Settings.Value("rol")=VarRol
		If Err Then MsgBox Err.Description
		On Error Goto 0
	End With
'	On Error Goto 0
'		Controller.SolMask(0)=0
'      vpmTimer.AddTimer 2000,"Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the timer to renable all the solenoids after 2 seconds
		Controller.Run

	If B2SOn then
		for each objekt in backdropstuff : objekt.visible = 0 : next
	End If

	if ballshadows=1 then
		BallShadowUpdate.enabled=1
	  else
		BallShadowUpdate.enabled=0
	end if

	if UnicornFart=1 then 
		for each light in UnicornFarts:light.visible=1:Next
		for each light in GILights:light.visible=0: next
	  Else
		for each light in UnicornFarts:light.visible=0:Next
		for each light in GILights:light.visible=1: next
	end if

	if flippershadows=1 then
		FlipperLSh.visible=1
		FlipperRSh.visible=1
	  else
		FlipperLSh.visible=0
		FlipperRSh.visible=0
	end if

	Controller.SolMask(0)=0
	vpmTimer.AddTimer 2000,"Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the timer to renable all the solenoids after 2 s
	
	PinMAMETimer.Interval=PinMAMEInterval
	PinMAMETimer.Enabled=1

	vpmNudge.TiltSwitch=57
	vpmNudge.Sensitivity=1
	vpmNudge.TiltObj=Array(LeftSlingshot,RightSlingshot,Bumper1,Bumper2,Bumper3,Bumper4)

	Set bsTrough=New cvpmBallstack
	bsTrough.InitSw 0,67,0,0,0,0,0,0
	bsTrough.InitKick BallRelease,90,5
	bsTrough.InitExitSnd SoundFX("ballrel",DOFContactors),SoundFX("solon",DOFContactors)
	bsTrough.Balls=1

	Set dtL=New cvpmDropTarget
	dtL.InitDrop Array(Target5,Target6,Target7),Array(40,50,60)
	dtL.initsnd SoundFX("flapclos",DOFContactors),SoundFX("solenoid",DOFContactors)

	kicker1.CreateSizedBallWithMass 25, CaptiveBallMass
	kicker1.kick 0, 1

End Sub

Sub Trigger1_Hit:Controller.Switch(41)=1:End Sub
Sub Trigger1_unHit:Controller.switch(41)=0:End Sub
Sub LeftInlane_Hit:Controller.Switch(43)=1:End Sub
Sub LeftInlane_unHit:Controller.Switch(43)=0:End Sub
Sub Bumper1_Hit:vpmTimer.PulseSw 44:PlaySoundAt SoundFX("JetBumper",DOFContactors), ActiveBall:DOF 103, DOFPulse:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 44:PlaySoundAt SoundFX("JetBumper2",DOFContactors), ActiveBall:DOF 105, DOFPulse:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 44:PlaySoundAt SoundFX("JetBumper3",DOFContactors), ActiveBall:DOF 106, DOFPulse:End Sub
Sub Bumper4_Hit:vpmTimer.PulseSw 44:PlaySoundAt SoundFX("JetBumper4",DOFContactors), ActiveBall:DOF 104, DOFPulse:End Sub
Sub Trigger2_Hit:Controller.Switch(51)=1:End Sub
Sub Trigger2_unHit:Controller.switch(51)=0:End Sub
Sub RightInlane_Hit:Controller.Switch(53)=1:End Sub
Sub RightInlane_UnHit:Controller.Switch(53)=0:End Sub
Sub Trigger3_Hit:Controller.Switch(61)=1:End Sub
Sub Trigger3_unHit:Controller.switch(61)=0:End Sub
Sub LeftOutlane_Hit:Controller.Switch(63)=1:DOF 107, DOFOn:End Sub
Sub LeftOutlane_unHit:Controller.Switch(63)=0:DOF 107, DOFOff:End Sub
Sub RightOutlane_Hit:Controller.Switch(63)=1:DOF 108, DOFOn:End Sub
Sub RightOutlane_UnHit:Controller.Switch(63)=0:DOF 108, DOFOff:End Sub
Sub Drain_Hit()PlaysoundAt "drain", Drain:bsTrough.AddBall Me:End Sub 'switch 67
Sub Trigger4_Hit:Controller.Switch(71)=1:End Sub
Sub Trigger4_unHit:Controller.switch(71)=0:End Sub


Sub LeftSlingShot_Slingshot
	VpmTimer.PulseSw 73
	PlaySoundAt SoundFX("sling",DOFContactors), slingL
	DOF 101,DOFPulse
    LSling.Visible = 0
    LSling1.Visible = 1
	slingL.rotx = 20
    me.uservalue = 1
    Me.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case me.uservalue
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:slingL.rotx = 10
        Case 4:slingL.rotx = 0:LSLing2.Visible = 0:LSLing.Visible = 1:me.TimerEnabled = 0
    End Select
    me.uservalue = me.uservalue + 1
End Sub

Sub RightSlingShot_Slingshot
	VpmTimer.PulseSw 73
	PlaySoundAt SoundFX("sling2",DOFContactors), slingR
	DOF 102,DOFPulse
    RSling.Visible = 0
    RSling1.Visible = 1
	slingR.rotx = 20
    me.uservalue = 1
    Me.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case me.uservalue
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:slingR.rotx = 10
        Case 4:slingR.rotx = 0:RSLing2.Visible = 0:RSLing.Visible = 1:me.TimerEnabled = 0
    End Select
    me.uservalue = me.uservalue + 1
End Sub



Sub Spinner_Spin
	vpmTimer.PulseSw 74
	PlaySound "fx_spinner", 0, .25, AudioPan(Spinner), 0.25, 0, 0, 1, AudioFade(Spinner)
End Sub



'****Targets***

Sub Target1_hit:vpmTimer.PulseSw 62:PlaySoundAtBall SoundFX("fx_target",DOFTargets):End Sub
Sub Target2_hit:vpmTimer.PulseSw 52:PlaySoundAtBall SoundFX("fx_target",DOFTargets):End Sub
Sub Target3_hit:vpmTimer.PulseSw 42:PlaySoundAtBall SoundFX("fx_target",DOFTargets):End Sub
Sub Target4_hit:vpmTimer.PulseSw 72:PlaySoundAtBall SoundFX("fx_target",DOFTargets):End Sub
Sub Target5_dropped:vpmTimer.PulseSw 40:PlaySoundAtBall SoundFX("fx_target",DOFDropTargets):End Sub
Sub Target6_dropped:vpmTimer.PulseSw 50:PlaySoundAtBall SoundFX("fx_target",DOFDropTargets):End Sub
Sub Target7_dropped:vpmTimer.PulseSw 60:PlaySoundAtBall SoundFX("fx_target",DOFDropTargets):End Sub
Sub Target8_hit:vpmTimer.PulseSw 70:PlaySoundAtBall SoundFX("fx_target",DOFTargets):End Sub


'****Lights***

set lights(3)=Light3
set lights(4)=light4
set lights(5)=light5
set lights(6)=light6
set lights(7)=light7
set lights(12)=light12
set lights(13)=light13
set lights(14)=light14
set lights(15)=light15
set lights(16)=light16
set lights(17)=light17
set lights(18)=light18
set lights(19)=light19
set lights(20)=light20
set lights(21)=light21
set lights(22)=light22
set lights(23)=light23
set lights(24)=light24
set lights(25)=light25
set lights(26)=light26
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
set lights(47)=light47
set lights(48)=light48
set lights(49)=light49
set lights(50)=light50
set lights(51)=light51

'Gottlieb Ice Fever
'added by Inkochnito
Sub editDips
	Dim vpmDips : Set vpmDips = New cvpmDips
	With vpmDips
		.AddForm 700,400,"Ice Fever - DIP switches"
		.AddFrame 0,0,190,"Maximum credits",49152,Array("8 credits",0,"10 credits",32768,"15 credits",&H00004000,"20 credits",49152)'dip 15&16
		.AddFrame 0,76,190,"Coin chute 1 and 2 control",&H00002000,Array("seperate",0,"same",&H00002000)'dip 14
		.AddFrame 0,122,190,"3rd coin chute credits control",&H20000000,Array("no effect",0,"add 9",&H20000000)'dip 30
		.AddFrame 0,168,190,"Novelty mode",&H08000000,Array("normal game mode",0,"50K per special/extra ball",&H08000000)'dip 28
		.AddFrame 0,214,190,"Background sound volume",&H40000000,Array("half volume",0,"full volume",&H40000000)'dip 31
		.AddFrame 205,0,190,"High score to date awards",&H00C00000,Array("not displayed and no award",0,"displayed and no award",&H00800000,"displayed and 2 credits",&H00400000,"displayed and 3 credits",&H00C00000)'dip 23&24
		.AddFrame 205,76,190,"Game mode",&H10000000,Array("replay",0,"extra ball",&H10000000)'dip 29
		.AddFrame 205,122,190,"Playfield special",&H00200000,Array("replay",0,"extra ball",&H00200000)'dip 22
		.AddFrame 205,168,190,"Replay limit",&H04000000,Array("no limit",0,"one per game",&H04000000)'dip 27
		.AddFrame 205,214,190,"Balls per game",&H01000000,Array("5 balls",0,"3 balls",&H01000000)'dip 25
		.AddChk 205,265,190,Array("Match feature",&H02000000)'dip 26
		.AddChk 0,265,190,Array("Spare (game option)",&H80000000)'dip 32
		.AddLabel 50,290,300,20,"After hitting OK, press F3 to reset game with new settings."
		.ViewDips
	End With
End Sub
Set vpmShowDips = GetRef("editDips")

Sub DipsTimer_Timer()
	Dim TheDips(32)
	Dim BPG, hsaward1, hsaward2, ebplay
    Dim DipsNumber

	DipsNumber = Controller.Dip(3)

	TheDips(32) = Int(DipsNumber/128)
	If TheDips(32) = 1 then DipsNumber = DipsNumber - 128 end if
	TheDips(31) = Int(DipsNumber/64)
	If TheDips(31) = 1 then DipsNumber = DipsNumber - 64 end if
	TheDips(30) = Int(DipsNumber/32)
	If TheDips(30) = 1 then DipsNumber = DipsNumber - 32 end if
	TheDips(29) = Int(DipsNumber/16)
	If TheDips(29) = 1 then DipsNumber = DipsNumber - 16 end if
	TheDips(28) = Int(DipsNumber/8)
	If TheDips(28) = 1 then DipsNumber = DipsNumber - 8 end if
	TheDips(27) = Int(DipsNumber/4)
	If TheDips(27) = 1 then DipsNumber = DipsNumber - 4 end if
	TheDips(26) = Int(DipsNumber/2)
	If TheDips(26) = 1 then DipsNumber = DipsNumber - 2 end if
	TheDips(25) = Int(DipsNumber)
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

	hsaward1 = TheDips(23)
	hsaward2 = TheDips(24)
	ebplay = TheDips(29)


	BPG = TheDips(25)
	If BPG = 1 then 
		instcard.image="InstCard3Balls"
	  Else
		instcard.image="InstCard5Balls"
	End if
	if ebplay = 0 then
		if hsaward1 = 1 Then
			repcard.image="replaycard1"
		  Else
			repcard.image="replaycard0"
		end if
	  else
		repcard.image="replaycard0"
	end if
'	DipsTimer.enabled=0
 End Sub

Set LampCallback = GetRef("UpdateMultipleLamps")

Sub UpdateMultipleLamps

    If Controller.Lamp(11) Then		'Game Over triggers match and BIP
        GOBox.text="GAME OVER"
      Else
        GOBox.text=""
    End If
 

    If Controller.Lamp(1)  Then					'Tilt
        TILTBox.text="TILT"
      Else
        TILTBox.text=""
    End If
 

    If Controller.Lamp(10) Then				'HIGH SCORE TO DATE
        HStoDateBox.text="HIGH SCORE TO DATE"
      Else
        HStoDateBox.text=""
    End If

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
Digits(24)=Array(a193,a194,a195,a196,a197,a198,a199,n,a200)
Digits(25)=Array(a201,a202,a203,a204,a205,a206,a207,n,a208)
Digits(26)=Array(a209,a210,a211,a212,a213,a214,a215,n,a216)
Digits(27)=Array(a217,a218,a219,a220,a221,a222,a223,n,a224)
 
'Ball in Play and Credit displays
 
Digits(30)=Array(e00,e01,e02,e03,e04,e05,e06,n)
Digits(31)=Array(e10,e11,e12,e13,e14,e15,e16,n)
Digits(28)=Array(f00,f01,f02,f03,f04,f05,f06,n)
Digits(29)=Array(f10,f11,f12,f13,f14,f15,f16,n)
 
 
Sub DisplayTimer_Timer
    Dim ChgLED,ii,num,chg,stat, obj
    ChgLed = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED) Then
'        If not b2son Then
        For ii = 0 To UBound(chgLED)
            num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
            if (num < 32 ) then
                For Each obj In Digits(num)
                    If chg And 1 Then obj.State = stat And 1
                    chg = chg\2 : stat = stat\2
                Next
            else
 
            end if
        next
'        end if
    end if
End Sub

Sub Gate1_Hit: PlaySoundAtBall "Gate" : End Sub
Sub Gate2_Hit: PlaySoundAtBall "Gate" : End Sub


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

'**************************************************************************
'                 Additional Positional Sound Playback Functions by DJRobX
'**************************************************************************

Sub PlaySoundAtVol(sound, tableobj, Vol)
		PlaySound sound, 1, Vol, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'Set position at table object, vol, and loops manually.

Sub PlaySoundAtVolLoops(sound, tableobj, Vol, Loops)
		PlaySound sound, Loops, Vol, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub


'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
	Dim tmp
    tmp = tableobj.y * 2 / Table1.height-1
    If tmp > 0 Then
		AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = tableobj.x * 2 / Table1.width-1
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
	PlaySound("fx_collide"), 0, Csng(velocity) ^2 / (VolDiv/VolCol), AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

'*****************************************
'           BALL SHADOW by ninnuzu
'*****************************************
Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5)

Sub BallShadowUpdate_timer()
    Dim BOT, b
	Dim maxXoffset
	maxXoffset=13
    BOT = GetBalls

	' render the shadow for each ball
    For b = 0 to UBound(BOT)
		BallShadow(b).X = BOT(b).X-maxXoffset*(1-(Bot(b).X)/(Table1.Width/2))
		BallShadow(b).Y = BOT(b).Y + 10
		If BOT(b).Z > 0 and BOT(b).Z < 30 Then
			BallShadow(b).visible = 1
		Else
			BallShadow(b).visible = 0
		End If
	Next
End Sub

Sub a_Triggers_Hit (idx)
'debug.print "Trigger hit"
	playsound "fx_sensor", 0,1,AudioPan(ActiveBall),0,0,0,1,AudioFade(ActiveBall)
End sub

Sub a_Pins_Hit (idx)
'debug.print "Pins hit"
	PlaySound "pinhit_low", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub a_Targets_Hit (idx)
	PlaySound "target", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub a_DropTargets_Hit (idx)
	PlaySound "DTDrop", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub a_Metals_Thin_Hit (idx)
'debug.print "Metals Thin hit"
	PlaySound "metalhit_thin", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub a_Metals_Medium_Hit (idx)
	PlaySound "metalhit_medium", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub a_Metals2_Hit (idx)
'debug.print "Metals2 hit"
	PlaySound "metalhit2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub a_Gates_Hit (idx)
	PlaySound "gate4", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

'Sub Spinner_Spin
'	PlaySound "fx_spinner", 0, .25, AudioPan(Spinner), 0.25, 0, 0, 1, AudioFade(Spinner)
'End Sub

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

Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub

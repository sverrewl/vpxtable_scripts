' Night Rider Bally 1976
' Allknowing2012
' https://ia800304.us.archive.org/11/items/nightridermanual/Night_Rider_Manual.pdf
' http://stevekulpa.net/pinball/nightrider_sl.htm
'
' Thanks for tweaks from Arngrim & Hauntfreaks..
'
'1.00 - Original Release

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' No special SSF tweaks yet.
' This is a JP table. He often uses walls as switches so I need to be careful of using PlaySoundAt

Option Explicit
Randomize

If Table1.ShowDT = False then ' hide the backglass lights when in FS mode
  B1.visible=False
  B2.visible=False
  Light1.visible=False
  Light2.visible=False
  Light3.visible=False
  Light4b.visible=False
End If

Const cGameName="nightrdr"   ' Use nightrdb for Free Play
Const UseSolenoids=2,UseLamps=True,UseGI=0,UseSyn=1,SSolenoidOn="SolOn",SSolenoidOff="Soloff",SFlipperOn="FlipperUpLeft",SFlipperOff="FlipperDown"
Const SCoin="coin3",cCredits="Night Rider (Bally 1976)"

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

LoadVPM "01560000","Bally.vbs", 3.36
Dim DesktopMode: DesktopMode = Table1.ShowDT

SolCallback(1)="vpmSolSound SoundFX(""Chime10"",DOFChimes),"
SolCallback(2)="vpmSolSound SoundFX(""Chime100"",DOFChimes),"
SolCallback(3)="vpmSolSound SoundFX(""Chime1000"",DOFChimes),"
SolCallback(4)="vpmSolSound SoundFX(""ChimeExtra"",DOFChimes),"

SolCallback(6)="vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(7)="SolBallRelease"
SolCallback(8)="bsSaucer.SolOut"
SolCallback(9)="dtL_SolDropUp" ' dtL.SolDropUp
SolCallback(10)="dtR_SolDropUp" 'dtR.SolDropUp
' 11 left Bumper
' 13 right Bumper
' 14 bottom Bumper

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Dim dtL, dtR, bstrough, bsSaucer

Sub SolLFlipper(Enabled)
	If Enabled Then
			 PlaySound SoundFX("FlipperUp",DOFContactors):LeftFlipper.RotateToEnd
		 Else
			 PlaySound SoundFX("FlipperDown",DOFContactors):LeftFlipper.RotateToStart
	End If
End Sub

Sub SolRFlipper(Enabled)
	If Enabled Then
			 PlaySound SoundFX("FlipperUp",DOFContactors):RightFlipper.RotateToEnd
	Else
			 PlaySound SoundFX("FlipperDown",DOFContactors):RightFlipper.RotateToStart
	End If
End Sub

Const sEnable=18
SolCallback(sEnable)="GameOn"

Sub GameOn(enabled)
  vpmNudge.SolGameOn(enabled)
  If Enabled Then
    GIOn
  Else
    GIOff
  End If
End Sub


Sub solballrelease(enabled)
    bstrough.solexit ssolenoidon, ssolenoidon,enabled:playsound SoundFX("BallRelease",DOFContactors)
End sub

Sub Table1_Init
	On Error Resume Next
	With Controller
		.GameName=cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
		.SplashInfoLine=cCredits
		.HandleMechanics=0
		.HandleKeyboard=0
		.ShowDMDOnly=1
		.ShowFrame=0
		.Hidden=True
		.ShowTitle=0
	End With
	Controller.Run
	If Err Then MsgBox Err.Description
	On Error Goto 0

	PinMAMETimer.Interval=PinMAMEInterval
	PinMAMETimer.Enabled=1

	vpmNudge.TiltSwitch=swTilt
	vpmNudge.Sensitivity=5
	vpmNudge.TiltObj=Array(LeftSlingshot,RightSlingshot,Bumper1,Bumper2,Bumper3)
	vpmNudge.SolGameOn(True)

    set bstrough= new cvpmballstack
    bstrough.initnotrough ballrelease,8,80,5

	Set bsSaucer=New cvpmBallStack
	bsSaucer.InitSaucer Kicker1,22,175+Int(Rnd*10),8
	bsSaucer.InitExitSnd SoundFX("popper_ball",DOFContactors),SoundFX("popper_ball",DOFContactors)

	set dtL=New cvpmDropTarget
	dtL.InitDrop Array(target1,target2,target3,target4,target5),Array(1,2,3,4,5)
	dtL.initsnd SoundFX("DROP_LEFT",DOFContactors),SoundFX("DTReset",DOFContactors)

	set dtR=New cvpmDropTarget
	dtR.InitDrop Array(target6,target7,target8,target9,target10),Array(25,26,27,28,29)
	dtR.initsnd SoundFX("DROP_RIGHT",DOFContactors),SoundFX("DTReset",DOFContactors)

    vpmMapLights InsertLights
End Sub

Sub Table1_Exit
  If B2SOn then Controller.Stop
End Sub

Sub Table1_KeyDown(ByVal keycode)
	If KeyCode=PlungerKey Then Plunger.Pullback:PlaySound "PlungerPull"
	If vpmKeyDown(KeyCode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal keycode)
	If KeyCode=PlungerKey Then Plunger.Fire:PlaySound "Plunger"
	If vpmKeyUp(KeyCode) Then Exit Sub
End Sub

Sub dtL_SolDropUp(enabled)
  if enabled Then
    dtLTimer.interval=500
    dtLTimer.enabled=True
  end if
End Sub

Sub dtLTimer_Timer()
    dtLTimer.enabled=False
    dtL.DropSol_On
End Sub

Sub dtR_SolDropUp(enabled)
  if enabled Then
    dtRTimer.interval=500
    dtRTimer.enabled=True
  end if
End Sub

Sub dtRTimer_Timer()
    dtRTimer.enabled=False
    dtR.DropSol_On
End Sub

Sub SpinnerLeft_spin():PlaySound "fx_spinner":vpmtimer.pulsesw 21:end Sub
Sub SpinnerRight_spin():PlaySound "fx_spinner":vpmtimer.pulsesw 20:end Sub

' Left Targets
Sub target1_Hit
	dtL.Hit 1
End Sub
Sub target2_Hit
	dtL.Hit 2
	End Sub
Sub target3_Hit
	dtL.Hit 3
End Sub
Sub target4_Hit
    dtL.Hit 4
End Sub
Sub target5_Hit
    dtL.Hit 5
End Sub

' Right Targets
Sub target6_Hit
    dtR.Hit 1
End Sub
Sub target7_Hit
    dtR.Hit 2
End Sub
Sub Target8_Hit
    dtR.Hit 3
End Sub
Sub target9_Hit
    dtR.Hit 4
End Sub
Sub target10_Hit
    dtR.Hit 5
End Sub

'Circle Targets
Sub sw17_Hit:vpmTimer.PulseSw 17:End Sub
Sub sw17b_Hit:vpmTimer.PulseSw 17:End Sub
Sub sw30_Hit:vpmTimer.PulseSw 30:End Sub

' Rubber Walls
Sub sw18a_Hit:vpmTimer.PulseSw 18:End Sub
Sub sw18b_Hit:vpmTimer.PulseSw 18:End Sub
Sub sw19a_Hit:vpmTimer.PulseSw 19:End Sub
Sub sw19b_Hit:vpmTimer.PulseSw 19:End Sub

' Rollover Switches
Sub sw32_Hit:Controller.Switch(32)=1:End Sub
Sub sw32_unHit:Controller.Switch(32)=0:End Sub
Sub sw24_Hit:Controller.Switch(24)=1:End Sub
Sub sw24_unHit:Controller.Switch(24)=0:End Sub

Sub sw23_Hit:Controller.Switch(23)=1:End Sub
Sub sw23_unHit:Controller.Switch(23)=0:End Sub
Sub sw31_Hit:Controller.Switch(31)=1:End Sub
Sub sw31_unHit:Controller.Switch(31)=0:End Sub

Sub Drain_Hit():bsTrough.AddBall Me::End Sub

Sub Kicker1_hit():bsSaucer.AddBall 0:End Sub

Sub GIOn
	dim bulb
	for each bulb in GILights
	bulb.state = LightStateOn
	next
End Sub

Sub GIOff
	dim bulb
	for each bulb in GILights
	bulb.state = LightStateOff
	next
End Sub

'Set LampCallback=GetRef("UpdateMultipleLamps")

'Sub UpdateMultipleLamps
'End Sub
Sub FlipperTimer_Timer()
   GateLP.RotZ = ABS(GateL.currentangle)
   GateRP.RotZ = ABS(GateR.currentangle)
End Sub

Dim bump1,bump2,bump3

Sub Bumper1_Hit:vpmTimer.PulseSw 38:bump1 = 1:Me.TimerEnabled = 1:DOF 106, DOFPulse:End Sub
Sub Bumper1_Timer()
	Select Case bump1
        Case 1:Ring1.Z = -30:bump1 = 2
        Case 2:Ring1.Z = -20:bump1 = 3
        Case 3:Ring1.Z = -10:bump1 = 4
        Case 4:Ring1.Z = 0:Me.TimerEnabled = 0
	End Select
End Sub

Sub Bumper2_Hit:vpmTimer.PulseSw 40:bump2 = 1:Me.TimerEnabled = 1:DOF 103, DOFPulse:End Sub
Sub Bumper2_Timer()
	Select Case bump2
        Case 1:Ring2.Z = -30:bump2 = 2
        Case 2:Ring2.Z = -20:bump2 = 3
        Case 3:Ring2.Z = -10:bump2 = 4
        Case 4:Ring2.Z = 0:Me.TimerEnabled = 0
	End Select
End Sub

Sub Bumper3_Hit:vpmTimer.PulseSw 39:bump3 = 1:Me.TimerEnabled = 1:DOF 104, DOFPulse:End Sub
Sub Bumper3_Timer()
	Select Case bump3
        Case 1:Ring3.Z = -30:bump3 = 2
        Case 2:Ring3.Z = -20:bump3 = 3
        Case 3:Ring3.Z = -10:bump3 = 4
        Case 4:Ring3.Z = 0:Me.TimerEnabled = 0
	End Select
End Sub

' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Sub Pins_Hit (idx)
	PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
	PlaySound SoundFX("target",DOFContactors), 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub TargetBankWalls_Hit (idx)
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
 '  debug.print "rubber"
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
  '  debug.print "Posts"
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

Sub Bumpers_Hit(idx)
	Select Case Int(Rnd*4)+1
		Case 1 : PlaySound SoundFx("fx_bumper2",DOFContactors), 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound SoundFx("fx_bumper2",DOFContactors), 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound SoundFx("fx_bumper3",DOFContactors), 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 4 : PlaySound SoundFx("fx_bumper4",DOFContactors), 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep, Tstep

Sub RightSlingShot_Slingshot
    PlaySound SoundFX("slingshotright",DOFContactors), 0, 1, 0.05, 0.05
    vpmTimer.PulseSw 36
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
    PlaySound SoundFX("slingshotleft",DOFContactors), 0, 1, 0.05, 0.05
    vpmTimer.PulseSw 37
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub


'Digital LED Display

Dim Digits(28)
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
Digits(24)=Array(f00,f01,f02,f03,f04,f05,f06,n,f08)
Digits(25)=Array(f10,f11,f12,f13,f14,f15,f16,n,f18)

Digits(26)=Array(e00,e01,e02,e03,e04,e05,e06,n,e08)
Digits(27)=Array(e10,e11,e12,e13,e14,e15,e16,n,e18)




Sub DisplayTimer_Timer
	Dim ChgLED,ii,num,chg,stat,obj
	ChgLed = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
If Not IsEmpty(ChgLED) Then
		If DesktopMode = True Then
		For ii = 0 To UBound(chgLED)
			num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
			if (num < 28) then
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

'Bally Night Rider
'added by Inkochnito
Sub editDips
	Dim vpmDips : Set vpmDips = New cvpmDips
	With vpmDips
		.AddForm 700,400,"Night Rider - DIP switches"
		.AddFrame 2,2,190,"Maximum credits",&H00070000,Array("5 credits",0,"10 credits",&H00010000,"15 credits",&H00020000,"20 credits",&H00030000,"25 credits",&H00040000,"30 credits",&H00050000,"35 credits",&H00060000,"40 credits",&H00070000)'dip 17&18&19
		.AddFrame 2,140,190,"High game to date",&H00004000,Array("no award",0,"3 credits",&H00004000)'dip 15
		.AddChk 2,190,190,Array("Melody option",&H00000080)'dip 8
		.AddChk 2,205,115,Array("Credits display",&H00080000)'dip 20
		.AddChk 2,220,180,Array("Match feature",&H00100000)'dip 21
		.AddFrame 205,2,190,"Balls per game", 32768,Array("3 balls",0,"5 balls", 32768)'dip 16
		.AddFrame 205,48,190,"Novelty option",&H00800000,Array("normal game play",0,"points for extra ball or special",&H00800000)'dip 24
		.AddFrame 205,94,190,"Drop target award",&H40000000,Array("conservative",0,"liberal",&H40000000)'dip 31
		.AddFrame 205,140,190,"Lane extra ball lite",&H80000000,Array("constantly lit",0,"alternating",&H80000000)'dip 32
		.AddFrame 205,186,190,"High score feature",&H00000020,Array("extra ball",0,"replay",&H00000020)'dip 6
		.AddLabel 30,240,360,20,"Note: Spinners, Inlanes and Bumper lights alternate only in 5 Ball setting."
        .AddLabel 50,260,300,20,"After hitting OK, press F3 to reset game with new settings."
		.ViewDips
	End With
End Sub
Set vpmShowDips = GetRef("editDips")

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

Const tnob = 2 ' total number of balls
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    RollingTimer.interval=100
    RollingTimer.enabled=True
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


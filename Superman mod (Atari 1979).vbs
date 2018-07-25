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

Const cGameName="superman",UseSolenoids=2,UseLamps=1,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff",SFlipperOn="fx_Flipperup",SFlipperOff="fx_Flipperdown"
Const SCoin="coin"

LoadVPM "01120000", "ATARI.VBS", 3.0
Dim DesktopMode: DesktopMode = Table1.ShowDT

'Solenoid Call backs
'**********************************************************************************************************
SolCallback(12)	= "dtBank.SolDropUp"
SolCallback(13)	= "bsSaucer.SolOut"
SolCallback(15)	= "bsTrough.SolOut"
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

'Initiate Table
'**********************************************************************************************************

Dim bsTrough, bsSaucer, dtBank

Sub Table1_Init
	vpmInit Me
	On Error Resume Next
		With Controller
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
		.SplashInfoLine = "Superman (Atari 1979)"&chr(13)&"You Suck"
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

    vpmNudge.TiltSwitch = 48
    vpmNudge.Sensitivity = 5
	vpmNudge.TiltObj = Array(Bumper1,Bumper2,Bumper3,Bumper4,RightSlingshot,LeftSlingshot)

Set bsTrough = New cvpmBallStack
		bsTrough.InitSw 0,1,0,0,0,0,0,0
		bsTrough.InitKick BallRelease,90,5
		bsTrough.InitExitSnd SoundFX("BallRelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
		bsTrough.Balls = 1

	Set bsSaucer = New cvpmBallStack
		bsSaucer.InitSaucer Kicker1,43,176+rnd(1)*5, 11+rnd(1)*6
		bsSaucer.InitExitSnd SoundFX("Popper",DOFContactors),SoundFX("solenoid",DOFContactors)

	Set dtBank = New cvpmDropTarget
		dtBank.InitDrop Array(sw9,sw10,sw11,sw12,sw13),Array(9,10,11,12,13)
		dtBank.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)
End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)
	If KeyDownHandler(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Pullback:playsound"plungerpull"
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
	If KeyUpHandler(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Fire:PlaySound"plunger"
End Sub

'**********************************************************************************************************
'Switches
'**********************************************************************************************************

'Drain and Kickers
Sub Drain_Hit:playsound"drain":bsTrough.addball me:End Sub
Sub Kicker1_Hit:bsSaucer.AddBall 0:End Sub

'Drop Targets
 Sub Sw9_Hit: dtBank.Hit 1 :End Sub
 Sub Sw10_Hit:dtBank.Hit 2 :End Sub
 Sub Sw11_Hit:dtBank.Hit 3 :End Sub
 Sub Sw12_Hit:dtBank.Hit 4 :End Sub
 Sub Sw13_Hit:dtBank.Hit 5 :End Sub

'Bumpers
Sub Bumper1_Hit : vpmTimer.PulseSw(52) : playsound SoundFX("fx_bumper1",DOFContactors): End Sub
Sub Bumper2_Hit : vpmTimer.PulseSw(56) : playsound SoundFX("fx_bumper1",DOFContactors): End Sub
Sub Bumper3_Hit : vpmTimer.PulseSw(59) : playsound SoundFX("fx_bumper1",DOFContactors): End Sub
Sub Bumper4_Hit : vpmTimer.PulseSw(53) : playsound SoundFX("fx_bumper1",DOFContactors): End Sub

'Scoring Rubbers
Sub sw201_Hit()	: vpmTimer.PulseSwitch (20), 0, "" : playsound"slingshot": End Sub
Sub sw202_Hit()	: vpmTimer.PulseSwitch (20), 0, "" : playsound"slingshot": End Sub
Sub sw203_Hit()	: vpmTimer.PulseSwitch (20), 0, "" : playsound"slingshot": End Sub
Sub sw204_Hit()	: vpmTimer.PulseSwitch (20), 0, "" : playsound"slingshot": End Sub
Sub sw205_Hit()	: vpmTimer.PulseSwitch (20), 0, "" : playsound"slingshot": End Sub
Sub sw206_Hit()	: vpmTimer.PulseSwitch (20), 0, "" : playsound"slingshot": End Sub
Sub sw207_Hit()	: vpmTimer.PulseSwitch (20), 0, "" : playsound"slingshot": End Sub

'Spinners
Sub Spinner_Spin:vpmTimer.PulseSw 60 : playsound"fx_spinner" : End Sub
Sub Spinner1_Spin:vpmTimer.PulseSw 61 : playsound"fx_spinner" : End Sub

'Stand Up Targets
Sub sw32_Hit()		: vpmTimer.PulseSwitch (32), 0, "" : End Sub
Sub sw33_Hit()		: vpmTimer.PulseSwitch (33), 0, "" : End Sub
Sub sw34_Hit()		: vpmTimer.PulseSwitch (34), 0, "" : End Sub
Sub sw35_Hit()		: vpmTimer.PulseSwitch (35), 0, "" : End Sub
Sub sw36_Hit()		: vpmTimer.PulseSwitch (36), 0, "" : End Sub

'Trigger Wires
Sub sw16_Hit()		: Controller.Switch(16) = 1 : playsound"rollover" : End Sub
Sub sw16_UnHit()	: Controller.Switch(16) = 0 : End Sub
Sub sw17_Hit()		: Controller.Switch(17) = 1 : playsound"rollover" : End Sub
Sub sw17_UnHit()	: Controller.Switch(17) = 0 : End Sub
Sub sw18_Hit()		: Controller.Switch(18) = 1 : playsound"rollover" : End Sub
Sub sw18_UnHit()	: Controller.Switch(18) = 0 : End Sub
Sub sw19_Hit()		: Controller.Switch(19) = 1 : playsound"rollover" : End Sub
Sub sw19_UnHit()	: Controller.Switch(19) = 0 : End Sub
Sub sw27_Hit()		: Controller.Switch(27) = 1 : playsound"rollover" : End Sub
Sub sw27_UnHit()	: Controller.Switch(27) = 0 : End Sub
Sub sw28_Hit()		: Controller.Switch(28) = 1 : playsound"rollover" : End Sub
Sub sw28_UnHit()	: Controller.Switch(28) = 0 : End Sub
Sub sw29_Hit()		: Controller.Switch(29) = 1 : playsound"rollover" : End Sub
Sub sw29_UnHit()	: Controller.Switch(29) = 0 : End Sub
Sub sw40_Hit()		: Controller.Switch(40) = 1 : playsound"rollover" : End Sub
Sub sw40_UnHit()	: Controller.Switch(40) = 0 : End Sub
Sub sw41_Hit()		: Controller.Switch(41) = 1 : playsound"rollover" : End Sub
Sub sw41_UnHit()	: Controller.Switch(41) = 0 : End Sub

'Star Trigger
Sub sw20_Hit()		: Controller.Switch(20) = 1 : playsound"rollover" : End Sub
Sub sw20_UnHit()	: Controller.Switch(20) = 0 : End Sub
Sub sw20a_Hit()		: Controller.Switch(20) = 1 : playsound"rollover" : End Sub
Sub sw20a_UnHit()	: Controller.Switch(20) = 0 : End Sub
Sub sw42_Hit()		: Controller.Switch(42) = 1 : playsound"rollover" : End Sub
Sub sw42_UnHit()	: Controller.Switch(42) = 0 : End Sub



Set Lights(1) = Light1
Set Lights(2) = Light2
Set Lights(3) = Light3
Set Lights(4) = Light4
Set Lights(5) = Light5
Set Lights(6) = Light6
Set Lights(7) = Light7
Set Lights(8) = Light8
Set Lights(9) = Light9
Set Lights(10) = Light10
Set Lights(11) = Light11
Set Lights(12) = Light12
Set Lights(13) = Light13
Lights(14) = array(Light14,Light14a)
'Set Lights(15) = Light15
Set Lights(16) = Light16
Set Lights(17) = Light17
Set Lights(18) = Light18
Set Lights(19) = Light19
Set Lights(20) = Light20
Set Lights(21) = Light21
Lights(22) = array(Light22,Light22a)
Lights(23) = array(Light23,Light23a)
Lights(24) = array(Light24,Light24a)
Set Lights(25) = Light25
Set Lights(26) = Light26
Set Lights(27) = Light27
Set Lights(28) = Light28
Set Lights(29) = Light29
Set Lights(30) = Light30
Set Lights(31) = Light31
Set Lights(32) = Light32
Set Lights(33) = Light33
Set Lights(34) = Light34
Set Lights(35) = Light35
Set Lights(36) = Light36
Set Lights(37) = Light37
Set Lights(38) = Light38
Set Lights(39) = Light39
Set Lights(40) = Light40
Set Lights(41) = Light41
Set Lights(42) = Light42
Set Lights(43) = Light43
Set Lights(44) = Light44
Set Lights(45) = Light45
Set Lights(46) = Light46
Set Lights(47) = Light47
Set Lights(48) = Light48
Set Lights(49) = Light49
Set Lights(50) = Light50
Set Lights(51) = Light51
Set Lights(52) = Light52

'Backglass Lights
'Set Lights(53) = Light53
'Set Lights(54) = Light54
'Set Lights(55) = Light55
'Set Lights(56) = Light56
'Set Lights(57) = Light57
'Set Lights(58) = Light58
'Set Lights(59) = Light59
'Set Lights(60) = Light60
'Set Lights(61) = Light61
'Set Lights(62) = Light62
'Set Lights(63) = Light63
'Set Lights(64) = Light64

Dim Digits(27)

Digits(0)=Array(D1,D2,D3,D4,D5,D6,D7)
Digits(1)=Array(D11,D12,D13,D14,D15,D16,D17)
Digits(2)=Array(D21,D22,D23,D24,D25,D26,D27)
Digits(3)=Array(D31,D32,D33,D34,D35,D36,D37)
Digits(4)=Array(D41,D42,D43,D44,D45,D46,D47)
Digits(5)=Array(D51,D52,D53,D54,D55,D56,D57)
Digits(6)=Array(D61,D62,D63,D64,D65,D66,D67)
Digits(7)=Array(D71,D72,D73,D74,D75,D76,D77)
Digits(8)=Array(D81,D82,D83,D84,D85,D86,D87)
Digits(9)=Array(D91,D92,D93,D94,D95,D96,D97)
Digits(10)=Array(D101,D102,D103,D104,D105,D106,D107)
Digits(11)=Array(D111,D112,D113,D114,D115,D116,D117)
Digits(12)=Array(D121,D122,D123,D124,D125,D126,D127)
Digits(13)=Array(D131,D132,D133,D134,D135,D136,D137)
Digits(14)=Array(D141,D142,D143,D144,D145,D146,D147)
Digits(15)=Array(D151,D152,D153,D154,D155,D156,D157)
Digits(16)=Array(D161,D162,D163,D164,D165,D166,D167)
Digits(17)=Array(D171,D172,D173,D174,D175,D176,D177)
Digits(18)=Array(D181,D182,D183,D184,D185,D186,D187)
Digits(19)=Array(D191,D192,D193,D194,D195,D196,D197)
Digits(20)=Array(D201,D202,D203,D204,D205,D206,D207)
Digits(21)=Array(D211,D212,D213,D214,D215,D216,D217)
Digits(22)=Array(D221,D222,D223,D224,D225,D226,D227)
Digits(23)=Array(D231,D232,D233,D234,D235,D236,D237)
Digits(24)=Array(D241,D242,D243,D244,D245,D246,D247)
Digits(25)=Array(D251,D252,D253,D254,D255,D256,D257)
Digits(26)=Array(D261,D262,D263,D264,D265,D266,D267)
Digits(27)=Array(D271,D272,D273,D274,D275,D276,D277)

   Sub DisplayTimer_Timer
	Dim ChgLED,ii,jj,num,chg,stat,obj,b,x
	ChgLED=Controller.ChangedLEDs(&Hffffffff,&Hffffffff)
	If Not IsEmpty(ChgLED) Then
		If DesktopMode = True Then
		For ii=0 To UBound(chgLED)
			num=chgLED(ii,0):chg=chgLED(ii,1):stat=chgLED(ii,2)
			For Each obj In Digits(num)
				If chg And 1 Then obj.State=stat And 1
				chg=chg\2:stat=stat\2
			Next
		Next
		end if
	End If
End Sub

'Atari Superman
'added by Inkochnito
Sub editDips
	Dim vpmDips:Set vpmDips=New cvpmDips
	With vpmDips
		.AddForm 700,400,"Superman - DIP switches"
		.AddFrame 2,10,190,"Hi score to date award",&HC0000000,Array("1 credit",0,"2 credits",&HC0000000,"3 credits",&H40000000)'dip 31&32
		.AddFrame 2,73,190,"Extra ball award",&H00000300,Array("30,000 points",0,"20,000 points",&H00000200,"extra ball",&H00000300)'dip 9&10
		.AddFrame 2,136,190,"Special award",&H00000410,Array("50,000 points",0,"extra ball",&H00000010,"replay",&H00000410)'dip 5&11
		.AddFrame 210,10,190,"Maximum credits",&H00000007,Array("10 credits",&H00000006,"15 credits",&H00000005,"25 credits",&H00000003,"40 credits",0)'dip 1&2&3
		.AddFrame 210,86,190,"Balls per game",&H00000008,Array("5 balls",0,"3 balls",&H00000008)'dip 4
		.AddFrame 210,133,190,"SUPERMAN spellout award",&H00000800,Array("30,000 points",0,"50,000 points",&H00000800)'dip 12
		.AddChk 2,200,190,Array("Hi score displayed",&H00800000)'dip 24
		.AddChk 2,215,190,Array("Match feature",&H00000080)'dip 8
		.AddChk 210,185,190,Array("Free Play",&H00000020)'dip 6
		.AddChk 210,200,190,Array("Keep alive feature",32768)'dip 16
		.AddChk	210,215,190,Array("Million limit must be on",&H00400000)'dip 23
		.AddLabel 50,235,300,20,"After hitting OK, press F3 to reset game with new settings."
		.ViewDips
	End With
End Sub
Set vpmShowDips=GetRef("editDips")



'*****GI Lights On
dim xx
For each xx in GI:xx.State = 1: Next

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
	vpmTimer.PulseSwitch (57), 0, ""
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
	vpmTimer.PulseSwitch (58), 0, ""
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
  If Table1.VersionMinor > 3 OR Table1.VersionMajor > 10 Then
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
  Else
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0
  End if
End Sub


Option Explicit
Randomize

' Thalamus 2018-07-20
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' Thalamus 2018-11-01 : Improved directional sounds
' !! NOTE : Table not verified yet !!

' Options
' Volume devided by - lower gets higher sound

Const VolDiv = 2000    ' Lower number, louder ballrolling/collition sound
Const VolCol = 10      ' Ball collition divider ( voldiv/volcol )

' The rest of the values are multipliers
'
'  .5 = lower volume
' 1.5 = higher volume

Const VolBump   = 2    ' Bumpers volume.
Const VolRol    = 1    ' Rollovers volume.
Const VolGates  = 1    ' Gates volume.
Const VolMetal  = 1    ' Metals volume.
Const VolRB     = 1    ' Rubber bands volume.
Const VolRH     = 1    ' Rubber hits volume.
Const VolPo     = 1    ' Rubber posts volume.
Const VolPi     = 1    ' Rubber pins volume.
Const VolPlast  = 1    ' Plastics volume.
Const VolTarg   = 1    ' Targets volume.
Const VolWood   = 1    ' Woods volume.
Const VolKick   = 1    ' Kicker volume.
Const VolSpin   = 1.5  ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="eballchp",UseSolenoids=2,UseLamps=1,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff",SCoin="coin"

LoadVPM "02000000", "6803.VBS", 3.10
Dim DesktopMode: DesktopMode = Table1.ShowDT

If DesktopMode = True Then 'Show Desktop components
Ramp16.visible=1
Ramp15.visible=1
Primitive13.visible=1
Else
Ramp16.visible=0
Ramp15.visible=0
Primitive13.visible=0
End if

'*************************************************************

'Solenoid Call backs
'**********************************************************************************************************

SolCallback(1)  =	"dtDrop2.SolDropUp"
SolCallback(2)  =	"dtDrop3.SolDropUp"
SolCallback(3)  =	"dtDrop4.SolDropUp"
SolCallback(4)  =	"dtDrop1.SolDropUp"
SolCallback(5)  =	"dtDrop5.SolDropUp"
SolCallback(8)  =	"bsSaucer.SolOut"
SolCallback(14) =   "bsTrough.SolOut"
SolCallback(15) =   "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(19) =   "vpmNudge.SolGameOn"

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors), LeftFlipper, VolFlip:LeftFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors), LeftFlipper, VolFlip:LeftFlipper.RotateToStart
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors), RightFlipper, VolFlip:RightFlipper.RotateToEnd:RightFlipper1.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors), RightFlipper, VolFlip:RightFlipper.RotateToStart:RightFlipper1.RotateToStart
     End If
End Sub
'**********************************************************************************************************

'Primitive Gate
Sub FlipperTimer_Timer
	PrimGate2.Rotz = Gate7.Currentangle
 End Sub

'*****GI Lights On
dim xx
For each xx in GI:xx.State = 1: Next

'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************
Dim bsTrough,dtDrop1,dtDrop2,dtDrop3,dtDrop4,dtDrop5,bsSaucer

Sub Table1_Init
	vpmInit Me
	On Error Resume Next
		With Controller
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
		.SplashInfoLine="Eight Ball Champ"&chr(13)&"You Suck"
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

	vpmNudge.TiltSwitch=15
	vpmNudge.Sensitivity=1
	vpmNudge.TiltObj=Array(Bumper1,Bumper2,Rightslingshot,Leftslingshot)

	Set bsTrough = New cvpmBallStack
		bsTrough.InitSw 0,8,0,0,0,0,0,0
		bsTrough.InitKick BallRelease,55,10
		bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
		bsTrough.Balls = 1

	Set bsSaucer=New cvpmBallStack
		bsSaucer.InitSaucer sw5,5,259+rnd(1)*3,8+rnd(1)*5 '5,260,10
		bsSaucer.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)

	Set dtDrop1=New cvpmDropTarget
		dtDrop1.InitDrop Array(sw30),30
		dtDrop1.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

	Set dtDrop2=New cvpmDropTarget
		dtDrop2.InitDrop Array(sw25),25
		dtDrop2.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

	Set dtDrop3=New cvpmDropTarget
		dtDrop3.InitDrop Array(sw27),27
		dtDrop3.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

	Set dtDrop4=New cvpmDropTarget
		dtDrop4.InitDrop Array(sw28),28
		dtDrop4.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

	Set dtDrop5=New cvpmDropTarget
		dtDrop5.InitDrop Array(sw31),31
		dtDrop5.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************
Sub Table1_KeyDown(ByVal keycode)
	If vpmKeyDown(KeyCode) Then Exit Sub
	If keycode=PlungerKey Then Plunger.Pullback:playsoundAtVol"plungerpull", plunger, 1
End Sub

Sub Table1_KeyUp(ByVal keycode)
	If vpmKeyUp(KeyCode) Then Exit Sub
	If keycode=PlungerKey Then Plunger.Fire:playsoundAtVol"plunger", plunger, 1
End Sub
'***********************************************************

'SAUCER CODE
 Sub Drain_Hit:bsTrough.addball me : playsoundAtVol"drain",drain,1 : End Sub
 Sub sw5_Hit(): bsSaucer.Addball Me : playsoundAtVol"popper_ball", sw5, VolKick : End Sub

'DROP TARGET
 Sub Sw30_Dropped:  dtDrop1.Hit 1:  End Sub

 Sub Sw25_Dropped:  dtDrop2.Hit 1:  End Sub

 Sub Sw27_Dropped:  dtDrop3.Hit 1:  End Sub

 Sub Sw28_Dropped:  dtDrop4.Hit 1:  End Sub

 Sub Sw31_Dropped:  dtDrop5.Hit 1:  End Sub

'STAND UP TARGET
Sub sw33_Hit:vpmTimer.PulseSw 33 : End Sub
Sub sw34_Hit:vpmTimer.PulseSw 34 : End Sub
Sub sw35_Hit:vpmTimer.PulseSw 35 : End Sub
Sub sw36_Hit:vpmTimer.PulseSw 36 : End Sub
Sub sw37_Hit:vpmTimer.PulseSw 37 : End Sub

'Star Triggers
Sub sw12_Hit():Controller.Switch(12)=1 : playsoundAtVol"rollover", ActiveBall, 1 : End Sub
Sub sw12_UnHit():Controller.Switch(12)=0:End Sub
Sub sw13_Hit():Controller.Switch(13)=1 : playsoundAtVol"rollover", ActiveBall, 1 : End Sub
Sub sw13_UnHit():Controller.Switch(13)=0:End Sub
Sub sw29_Hit():Controller.Switch(29)=1 : playsoundAtVol"rollover", ActiveBall, 1 : End Sub
Sub sw29_UnHit:Controller.Switch(29)=0:End Sub

'SPINNER
Sub sw32_Spin:vpmTimer.PulseSw 32 : playsoundAtVol"fx_spinner", sw32, VolSpin : End Sub

'BUMPER
Sub Bumper1_Hit:vpmTimer.PulseSw 1 : playsoundAtVol SoundFX("fx_bumper1",DOFContactors), Bumper1, VolBump: End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 2 : playsoundAtVol SoundFX("fx_bumper1",DOFContactors), Bumper2, VolBump: End Sub

'Wire Triggers
Sub sw18_Hit():Controller.Switch(18)=1 : playsoundAtVol"rollover",sw18, VolRol: End Sub
Sub sw18_UnHit():Controller.Switch(18)=0:End Sub
Sub sw19_Hit():Controller.Switch(19)=1 : playsoundAtVol"rollover",sw19, VolRol : End Sub
Sub sw19_UnHit():Controller.Switch(19)=0:End Sub
Sub sw20_Hit():Controller.Switch(20)=1 : playsoundAtVol"rollover",sw20, VolRol : End Sub
Sub sw20_UnHit():Controller.Switch(20)=0:End Sub

Sub sw17_Hit():Controller.Switch(17)=1 : playsoundAtVol"rollover", sw17, VolRol : End Sub
Sub sw17_UnHit():Controller.Switch(17)=0:End Sub
Sub sw21_Hit():Controller.Switch(21)=1 : playsoundAtVol"rollover", sw21, VolRol : End Sub
Sub sw21_UnHit():Controller.Switch(21)=0:End Sub
Sub sw23_Hit():Controller.Switch(23)=1 : playsoundAtVol"rollover", sw23, VolRol : End Sub
Sub sw23_UnHit():Controller.Switch(23)=0:End Sub
Sub sw24_Hit():Controller.Switch(24)=1 : playsoundAtVol"rollover", sw24, VolRol : End Sub
Sub sw24_UnHit():Controller.Switch(24)=0:End Sub

'**********************************************************************************************************

'Map lights to an array
'**********************************************************************************************************
Set Lights(2)=L2
Set Lights(4)=L4
Set Lights(5)=L5
Set Lights(6)=L6
Set Lights(7)=L7
Set Lights(8)=L8
Set Lights(9)=L9
Set Lights(10)=L10
Set Lights(11)=L11
Set Lights(12)=L12
Set Lights(13)=L13
Set Lights(14)=L14
Set Lights(15)=L15
Set Lights(17)=L17
Set Lights(20)=L20
Set Lights(21)=L21
Set Lights(22)=L22
Set Lights(23)=L23
Set Lights(24)=L24
Set Lights(25)=L25
Set Lights(26)=L26
Set Lights(27)=L27
Set Lights(28)=L28
Set Lights(29)=L29
Set Lights(30)=L30
Set Lights(31)=L31
Set Lights(35)=L35
Set Lights(36)=L36
Set Lights(37)=L37
Set Lights(38)=L38
Set Lights(39)=L39
Set Lights(40)=L40
Set Lights(41)=L41
Set Lights(41)=L41
Set Lights(42)=L42
Set Lights(43)=L43
Set Lights(44)=L44
Set Lights(45)=L45
Set Lights(46)=L46
Set Lights(47)=L47
Set Lights(49)=L49
Set Lights(50)=L50
Set Lights(51)=L51
Set Lights(52)=L52
Set Lights(53)=L53
Set Lights(54)=L54
Set Lights(57)=L57
Set Lights(58)=L58
Set Lights(59)=L59
Lights(60)=Array(L60,L60_1)
Lights(61)=Array(L61,L61_1)
Set Lights(62)=L62
Set Lights(63)=L63
Set Lights(65)=L65
Set Lights(66)=L66
Set Lights(67)=L67
Set Lights(68)=L68
Set Lights(69)=L69
Set Lights(73)=L73
Set Lights(74)=L74
Set Lights(75)=L75
Lights(76)=array(L76,L76_1)
Set Lights(77)=L77
Set Lights(78)=L78
Set Lights(79)=L79
Set Lights(81)=L81
Set Lights(82)=L82
Set Lights(83)=L83
Set Lights(84)=L84
Set Lights(85)=L85
Set Lights(89)=L89
Set Lights(90)=L90
Lights(91)=array(L91,L91_1)
Lights(92)=array(L92,L92_1)
Set Lights(93)=L93
Set Lights(94)=L94
Set Lights(95)=L95

'BackGlass
'Set Lights(1)=L1 'Same Player Shoot Again
'Set Lights(3)=L3 ' Game Over
'Set Lights(18)=L18 'Ball In Play
'Set Lights(19)=L19 'Tilt
'Set Lights(33)=L33 'Match
'Set Lights(34)=L34 'High Score
'**********************************************************************************************************

'backglass lamps
'**********************************************************************************************************
 Dim LED(31)
LED(0)=Array(d111,d112,d113,d114,d115,d116,d117)
LED(1)=Array(d121,d122,d123,d124,d125,d126,d127)
LED(2)=Array(d131,d132,d133,d134,d135,d136,d137)
LED(3)=Array(d141,d142,d143,d144,d145,d146,d147)
LED(4)=Array(d151,d152,d153,d154,d155,d156,d157)
LED(5)=Array(d161,d162,d163,d164,d165,d166,d167)
LED(6)=Array(d171,d172,d173,d174,d175,d176,d177)

LED(7)=Array(d211,d212,d213,d214,d215,d216,d217)
LED(8)=Array(d221,d222,d223,d224,d225,d226,d227)
LED(9)=Array(d231,d232,d233,d234,d235,d236,d237)
LED(10)=Array(d241,d242,d243,d244,d245,d246,d247)
LED(11)=Array(d251,d252,d253,d254,d255,d256,d257)
LED(12)=Array(d261,d262,d263,d264,d265,d266,d267)
LED(13)=Array(d271,d272,d273,d274,d275,d276,d277)

LED(14)=Array(d311,d312,d313,d314,d315,d316,d317)
LED(15)=Array(d321,d322,d323,d324,d325,d326,d327)
LED(16)=Array(d331,d332,d333,d334,d335,d336,d337)
LED(17)=Array(d341,d342,d343,d344,d345,d346,d347)
LED(18)=Array(d351,d352,d353,d354,d355,d356,d357)
LED(19)=Array(d361,d362,d363,d364,d365,d366,d367)
LED(20)=Array(d371,d372,d373,d374,d375,d376,d377)

LED(21)=Array(d411,d412,d413,d414,d415,d416,d417)
LED(22)=Array(d421,d422,d423,d424,d425,d426,d427)
LED(23)=Array(d431,d432,d433,d434,d435,d436,d437)
LED(24)=Array(d441,d442,d443,d444,d445,d446,d447)
LED(25)=Array(d451,d452,d453,d454,d455,d456,d457)
LED(26)=Array(d461,d462,d463,d464,d465,d466,d467)
LED(27)=Array(d471,d472,d473,d474,d475,d476,d477)

LED(28)=Array(d511,d512,d513,d514,d515,d516,d517)
LED(29)=Array(d521,d522,d523,d524,d525,d526,d527)

LED(30)=Array(d611,d612,d613,d614,d615,d616,d617)
LED(31)=Array(d621,d622,d623,d624,d625,d626,d627)


Dim ChgLED, ii, num, chg, stat, obj



Sub DisplayTimer_Timer
	ChgLED = Controller.ChangedLEDs (&Hffffffff, &Hffffffff)
If Not IsEmpty (ChgLED) Then
		If DesktopMode = True Then
		For ii = 0 To UBound (chgLED)
		num = chgLED (ii, 0) : chg = chgLED (ii, 1) : stat = chgLED (ii, 2)
		For Each obj In LED (num)
		If chg And 1 Then obj.State = stat And 1
			chg = chg \ 2 : stat = stat \ 2
		Next
		Next
		end if
End If
End Sub
'**********************************************************************************************************
'**********************************************************************************************************
'**********************************************************************************************************



'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep, Cstep

Sub RightSlingShot_Slingshot
	vpmTimer.PulseSw 3
    PlaySoundAtVol SoundFX("right_slingshot",DOFContactors), sling1, 1
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
	vpmTimer.PulseSw 4
    PlaySoundAtVol SoundFX("left_slingshot",DOFContactors), sling2, 1
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

Sub CSlingShot_Slingshot
	vpmTimer.PulseSw 7
    PlaySoundAtVol SoundFX("left_slingshot",DOFContactors), sling3, 1
    L2Sling.Visible = 0
    L2Sling1.Visible = 1
    sling3.TransZ = -20
    Cstep = 0
    CSlingShot.TimerEnabled = 1
End Sub

Sub CSlingShot_Timer
    Select Case CStep
        Case 3:L2SLing1.Visible = 0:L2SLing2.Visible = 1:sling3.TransZ = -10
        Case 4:L2SLing2.Visible = 0:L2SLing.Visible = 1:sling3.TransZ = 0:CSlingShot.TimerEnabled = 0:
    End Select
    Cstep = Cstep + 1
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
	PlaySound "pinhit_low", 0, Vol(ActiveBall)*VolPi, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
	PlaySound "target", 0, Vol(ActiveBall)*VolTarg, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
	PlaySound "metalhit_thin", 0, Vol(ActiveBall)*VolMetal, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
	PlaySound "metalhit_medium", 0, Vol(ActiveBall)*VolMetal, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
	PlaySound "metalhit2", 0, Vol(ActiveBall)*VolMetal, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
	PlaySound "gate4", 0, Vol(ActiveBall)*VolGates, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinner_Spin
	PlaySound "fx_spinner", Spinner, VolSpin
End Sub

Sub Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub Posts_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 16 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolPo, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 16 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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
		Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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

Sub PlaySoundAtVol(sound, tableobj, Volume)
  PlaySound sound, 1, Volume, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
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

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
  Vol = Csng(BallVel(ball) ^2 / VolDiv)
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
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / (VolDiv/VolCol), Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


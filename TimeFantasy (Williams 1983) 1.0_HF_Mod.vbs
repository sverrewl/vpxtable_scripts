Option Explicit
Randomize

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' No special SSF tweaks yet.

Const BallSize = 50

Const BallMass = 1.1

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="tmfnt_l5",UseSolenoids=2,UseLamps=0,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff", SCoin="coin"

LoadVPM "00990200", "S7.VBS", 2.34

Dim DesktopMode: DesktopMode = TimeFantasy.ShowDT

If TimeFantasy.ShowDT Then
	LeftRail.visible=1
	RightRail.visible=1
	LeftRef.visible=0
	RightRef.visible=0
End If

SolCallback(1)="bsTrough.SolOut"
SolCallback(11)="Sol11"			'GI
SolCallback(15)="SolRing"		'Bell
SolCallback(25)="Sol25"			'GameOn

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFX("fx_Flipup",DOFContactors):LeftFlipper.RotateToEnd
     Else
         PlaySound SoundFX("fx_Flipdown",DOFContactors):LeftFlipper.RotateToStart
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFX("fx_Flipup",DOFContactors):RightFlipper.RotateToEnd
     Else
         PlaySound SoundFX("fx_Flipdown",DOFContactors):RightFlipper.RotateToStart
     End If
End Sub

'GI Relay

Sub Sol11(Enabled)
	If Enabled Then
		dim xx
		For each xx in GI:xx.State = 0: Next
		playsound"Fx_Relay"
		For each xx in Plastics:xx.disablelighting=0: Next
		nuke1.enabled=0:nuke2.enabled=0:nuke.intensity=0
		LeftRef.opacity=2:RightRef.opacity=2
	Else
		For each xx in GI:xx.State = 1: Next
		playsound"Fx_Relay"
		For each xx in Plastics:xx.disablelighting=0.5: Next
		LeftRef.opacity=19:RightRef.opacity=27
	End If
End Sub

Sub Sol25(enabled)

End Sub

Sub fx_Timer
	Flasher1.amount=Flasher1.amount + 1
	If Flasher1.amount>420 Then
		fx.enabled=0
		fx2.enabled=1
	End If
End Sub

Sub fx2_Timer
	Flasher1.amount=Flasher1.amount - 1
	If Flasher1.amount<1 Then
		fx.enabled=1
		fx2.enabled=0
	End If
End Sub

Sub SolRing(enabled)
	If enabled Then
		PlaySound SoundFX("ringing",DOFContactors)
		nuke1.enabled=1
	End If
End Sub

Sub nuke1_Timer
	nuke.intensity=nuke.intensity+1
	If nuke.intensity>25 Then
		nuke2.enabled=1
		nuke1.enabled=0
	End If
End Sub

Sub nuke2_Timer
	nuke.intensity=nuke.intensity-1
	If nuke.intensity=0 Then
		nuke1.enabled=1
		nuke2.enabled=0
	End If
End Sub

'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************
Dim bsTrough

Sub TimeFantasy_Init
	vpmInit Me
	On Error Resume Next
		With Controller
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
		.SplashInfoLine = "Time Fantasy (Williams)"&chr(13)&"Winners don't use drugs"
		.HandleMechanics=0
		.HandleKeyboard=0
		.ShowDMDOnly=1
		.ShowFrame=0
		.ShowTitle=0
        .hidden = 0
		If Err Then MsgBox Err.Description
	End With
	On Error Goto 0
		Controller.SolMask(0)=0
      vpmTimer.AddTimer 2000,"Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the timer to renable all the solenoids after 2 seconds
		Controller.Run
	If Err Then MsgBox Err.Description
	On Error Goto 0

	PlayRandomTrack()

	PinMAMETimer.Interval=PinMAMEInterval
	PinMAMETimer.Enabled=1

	vpmNudge.TiltSwitch = swTilt
	vpmNudge.Sensitivity = 5
	vpmNudge.TiltObj = Array(Bumper33,Bumper34,Bumper35,LeftSlingshot,RightSlingshot)

	Set bsTrough = New cvpmBallStack
        bsTrough.InitSw 0,39,0,0,0,0,0,0
        bsTrough.InitKick BallRelease,90,5
		bsTrough.InitExitSnd SoundFX("popper_ball",DOFContactors), SoundFX("Solenoid",DOFContactors)
		bsTrough.Balls = 1
End Sub

Sub PlayRandomTrack()
	Select Case Int(Rnd*2)+1
		Case 1 : PlayMusic "TimeFantasy/soundtrack01.mp3" : MusicLoop.Enabled=1
		Case 2 : PlayMusic "TimeFantasy/soundtrack01.mp3" : MusicLoop.Enabled=1
	End Select

End Sub

Sub MusicLoop_Timer
	PlayMusic "TimeFantasy/soundtrack01.mp3"
End sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub TimeFantasy_KeyDown(ByVal KeyCode)
	If KeyDownHandler(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Pullback:playsound"plungerpull"
    if keycode = RightFlipperKey then Controller.Switch(40) = true
End Sub

Sub TimeFantasy_KeyUp(ByVal KeyCode)
	If KeyUpHandler(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Fire:PlaySound"plunger"
    if keycode = RightFlipperKey then Controller.Switch(40) = false
End Sub

'------------------------------
'------  Trough Handler  ------
'------------------------------
Sub Drain_Hit()
	bsTrough.AddBall Me
	playsound "Drain5"
End Sub


'------------------------------
'------  Switch Handler  ------
'------------------------------

Sub Bumper33_Hit
	vpmTimer.PulseSw 33
	Playsound SoundFX("fx_bumper",DOFContactors)
	G1.state=1
	G1.timerenabled=1
	Green1.image="d_green_lit"
	Green1.disablelighting=0.5
	Backwall.sideimage="backwall_lit"
	Backwall.disablelighting=0.3
	Backwall.timerenabled=1
End Sub

Sub Bumper34_Hit
	vpmTimer.PulseSw 34
	Playsound SoundFX("fx_bumper",DOFContactors)
	G2.state=1
	G2.timerenabled=1
	Green2.image="d_green_lit"
	Green2.disablelighting=0.5
	Backwall.sideimage="backwall_lit"
	Backwall.disablelighting=0.3
	Backwall.timerenabled=1
End Sub

Sub Bumper35_Hit
	vpmTimer.PulseSw 35
	Playsound SoundFX("fx_bumper",DOFContactors)
	G3.state=1
	G3.timerenabled=1
	Green3.image="d_green_lit"
	Green3.disablelighting=0.5
	Backwall.sideimage="backwall_lit"
	Backwall.disablelighting=0.3
	Backwall.timerenabled=1
End Sub

Sub Backwall_Timer
	Backwall.sideimage="backwall"
	Backwall.disablelighting=0
End Sub

Sub G1_Timer
	G1.state=0
	Green1.image="d_green"
	Green1.disablelighting=0
	G1.timerenabled=0
End Sub

Sub G2_Timer
	G2.state=0
	Green2.image="d_green"
	Green2.disablelighting=0
	G2.timerenabled=0
End Sub

Sub G3_Timer
	G3.state=0
	Green3.image="d_green"
	Green3.disablelighting=0
	G3.timerenabled=0
End Sub

sub Trigger16_hit:Controller.Switch(16)=1:End Sub
sub Trigger16_unhit:Controller.Switch(16)=0:End Sub
sub Trigger17_hit:Controller.Switch(17)=1:End Sub
sub Trigger17_unhit:Controller.Switch(17)=0:End Sub
sub Trigger18_hit:Controller.Switch(18)=1:End Sub
sub Trigger18_unhit:Controller.Switch(18)=0:End Sub
sub Trigger19_hit:Controller.Switch(19)=1:End Sub
sub Trigger19_unhit:Controller.Switch(19)=0:End Sub
sub Trigger20_hit:Controller.Switch(20)=1:End Sub
sub Trigger20_unhit:Controller.Switch(20)=0:End Sub
sub Trigger21_hit:Controller.Switch(21)=1:End Sub
sub Trigger21_unhit:Controller.Switch(21)=0:End Sub
sub Trigger22_hit:Controller.Switch(22)=1:End Sub
sub Trigger22_unhit:Controller.Switch(22)=0:End Sub
sub Trigger23_hit:Controller.Switch(23)=1:End Sub
sub Trigger23_unhit:Controller.Switch(23)=0:End Sub
sub Trigger24_hit:Controller.Switch(24)=1:End Sub
sub Trigger24_unhit:Controller.Switch(24)=0:End Sub
sub Trigger25_hit:Controller.Switch(25)=1:End Sub
sub Trigger25_unhit:Controller.Switch(25)=0:End Sub
sub Trigger26_hit:Controller.Switch(26)=1:End Sub
sub Trigger26_unhit:Controller.Switch(26)=0:End Sub

sub Trigger42_hit:Controller.Switch(42)=1:End Sub
sub Trigger42_unhit:Controller.Switch(42)=0:End Sub

sub Target9_hit:MoveTarget9:vpmTimer.PulseSw 9:End Sub
sub Target10_hit:MoveTarget10:vpmTimer.PulseSw 10:End Sub
sub Target11_hit:MoveTarget11:vpmTimer.PulseSw 11:End Sub
sub Target12_hit:MoveTarget12:vpmTimer.PulseSw 12:End Sub
sub Target13_hit:MoveTarget13:vpmTimer.PulseSw 13:End Sub
sub Target14_hit:MoveTarget14:vpmTimer.PulseSw 14:End Sub
sub Target15_hit:MoveTarget15:vpmTimer.PulseSw 15:End Sub

sub Target36_hit:MoveTarget36:vpmTimer.PulseSw 36:End Sub


'Rubber hits
sub Rubber27_hit:vpmTimer.PulseSw 27:End Sub
sub Rubber28_hit:vpmTimer.PulseSw 28:End Sub
sub Rubber29_hit:vpmTimer.PulseSw 29:End Sub
sub Rubber30_hit:vpmTimer.PulseSw 30:End Sub
sub Rubber31_hit:vpmTimer.PulseSw 31:End Sub
sub Rubber32_hit:vpmTimer.PulseSw 32:End Sub

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

    NFadeL 7, l7
    NFadeL 8, l8
    NFadeL 9, l9
    NFadeL 10, l10
    NFadeL 11, l11
    NFadeL 12, l12
    NFadeL 13, l13
    NFadeL 14, l14
    NFadeL 15, l15
    NFadeL 16, l16
    NFadeL 17, l17
    NFadeL 18, l18
    NFadeL 19, l19
    NFadeL 20, l20
    NFadeL 21, l21
    NFadeL 22, l22
    NFadeL 23, l23
    NFadeL 24, l24
    NFadeL 25, l25
    NFadeL 26, l26
    NFadeL 27, l27
    NFadeL 28, l28
    NFadeL 29, l29
    NFadeL 30, l30
    NFadeL 31, l31
    NFadeL 32, l32
    NFadeL 33, l33
    NFadeL 34, l34
    NFadeL 35, l35
    NFadeL 36, l36
    NFadeL 37, l37
    NFadeL 38, l38
    NFadeL 39, l39
    NFadeL 40, l40
    NFadeL 41, l41
    NFadeL 42, l42
    NFadeL 43, l43
    NFadeL 44, l44
    NFadeL 45, l45
    NFadeL 46, l46
    NFadeL 47, l47
    NFadeL 48, l48
    NFadeL 49, l49
    NFadeL 50, l50
    NFadeL 51, l51
    NFadeL 52, l52
    NFadeL 53, l53
    NFadeL 54, l54
    NFadeL 55, l55
    NFadeL 56, l56
    NFadeL 57, l57
    NFadeL 58, l58
    NFadeL 59, l59
    NFadeL 60, l60
    NFadeL 61, l61
    NFadeL 62, l62
    NFadeL 63, l63

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

'--------------------------------
'------  Helper Functions  ------
'--------------------------------

'Gates
Sub Gate_Hit:PlaySound "Gate5",0,1,0.16,0.1:End Sub
Sub Gate3_Hit:PlaySound "Gate5",0,1,-0.16,0.1:End Sub


'Round Targets
Sub MoveTarget9
	P_Target9.Transy = 5
	Target9.Timerenabled = False
	Target9.Timerenabled = True
End Sub
Sub Target9_Timer
	Target9.timerenabled = False
	P_Target9.transy = 0
End Sub

Sub MoveTarget10
	P_Target10.Transy = 5
	Target10.Timerenabled = False
	Target10.Timerenabled = True
End Sub
Sub Target10_Timer
	Target10.timerenabled = False
	P_Target10.transy = 0
End Sub

Sub MoveTarget11
	P_Target11.Transy = 5
	Target11.Timerenabled = False
	Target11.Timerenabled = True
End Sub
Sub Target11_Timer
	Target11.timerenabled = False
	P_Target11.transy = 0
End Sub

Sub MoveTarget12
	P_Target12.Transy = 5
	Target12.Timerenabled = False
	Target12.Timerenabled = True
End Sub
Sub Target12_Timer
	Target12.timerenabled = False
	P_Target12.transy = 0
End Sub

Sub MoveTarget13
	P_Target13.Transy = 5
	Target13.Timerenabled = False
	Target13.Timerenabled = True
End Sub
Sub Target13_Timer
	Target13.timerenabled = False
	P_Target13.transy = 0
End Sub

Sub MoveTarget14
	P_Target14.Transy = 5
	Target14.Timerenabled = False
	Target14.Timerenabled = True
End Sub
Sub Target14_Timer
	Target14.timerenabled = False
	P_Target14.transy = 0
End Sub

Sub MoveTarget15
	P_Target15.Transy = 5
	Target15.Timerenabled = False
	Target15.Timerenabled = True
End Sub
Sub Target15_Timer
	Target15.timerenabled = False
	P_Target15.transy = 0
End Sub

Sub MoveTarget36
	P_Target36.Transy = 5
	Target36.Timerenabled = False
	Target36.Timerenabled = True
End Sub
Sub Target36_Timer
	Target36.timerenabled = False
	P_Target36.transy = 0
End Sub


'Flipper Primitives

Sub FlipperTimer_Timer

	LFPrim.ObjRotz = LeftFlipper.CurrentAngle
	RFPrim.ObjRotz= RightFlipper.CurrentAngle
	batrightshadow.ObjRotz = RightFlipper.CurrentAngle
	batleftshadow.ObjRotz = LeftFlipper.CurrentAngle

End Sub


'**********************************************************************************************************
'**********************************************************************************************************
'	Start of VPX functions
'**********************************************************************************************************
'**********************************************************************************************************

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
	vpmTimer.PulseSw 38
    PlaySound SoundFX("SlingshotRight",DOFContactors), 0, 1, 0.05, 0.05
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
	S2.state=1
	S2.TimerEnabled=1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
	vpmTimer.PulseSw 37
    PlaySound SoundFX("SlingshotLeft",DOFContactors),0,1,-0.05,0.05
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
	S1.state=1
	S1.TimerEnabled=1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub S1_Timer
	S1.state=0
	S1.timerenabled=0
End Sub

Sub S2_Timer
	S2.state=0
	S2.timerenabled=0
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
	PlaySound "metalhit_thin", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
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

'	Ball shadows

Dim BallShadow
BallShadow = Array (BallShadow1)

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
        If BOT(b).X < TimeFantasy.Width/2 Then
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (TimeFantasy.Width/2))/7)) + 5
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (TimeFantasy.Width/2))/7)) - 5
        End If
        ballShadow(b).Y = BOT(b).Y + 10
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
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

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "TimeFantasy" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / TimeFantasy.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "TimeFantasy" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / TimeFantasy.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "TimeFantasy" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / TimeFantasy.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioFade(ball) ' Can this be together with the above function ?
  Dim tmp
  tmp = ball.y * 2 / TimeFantasy.height-1
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

Const tnob = 1 ' total number of balls
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
  If TimeFantasy.VersionMinor > 3 OR TimeFantasy.VersionMajor > 10 Then
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
  Else
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0
  End if
End Sub


' Thalamus : Exit in a clean and proper way
Sub TimeFantasy_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


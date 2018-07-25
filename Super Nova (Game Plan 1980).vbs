Option Explicit
Randomize

' Thalamus 2018-07-24
' Table has its own "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' No special SSF tweaks yet.

Const FlippersAlwaysOn = 0 'Enable Flippers for testing
Const cGameName = "suprnova"
On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0
LoadVPM "01560000","GamePlan.vbs",3.36

Dim bsTrough,bsSaucer,bsSaucer2,wheel,oldvalue,newvalue,objekt,special
Const UseSolenoids=2,UseLamps=1,UseGI=1,UseSync=1,SCoin="fx_coin",SFlipperOn="fx_flipperup",SFlipperOff="fx_flipperdown"
Const SSolenoidOn = "fx_solenoid"
Const SSolenoidOff = "fx_solenoidoff"
Const sKicker1=15
Const sKicker2=10
Const sBallrelease=8

Sub Table1_Init
     With Controller
         .GameName = cGameName
         If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
         .SplashInfoLine = "Super Nova (GamePlan 1980)"
         .HandleKeyboard = 0
         .ShowTitle = 0
         .ShowDMDOnly = 1
         .ShowFrame = 0
         .HandleMechanics = False
         .Hidden = 0
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With

   	Set bsTrough=New cvpmBallStack
 	with bsTrough
		.InitSw 0,11,0,0,0,0,0,0
		.InitKick BallRelease,90,7
		bsTrough.InitExitSnd Soundfx("fx_ballrel",DOFContactors), Soundfx("fx_solenoid",DOFContactors)
		.Balls=1
 	end with

	Set bsSaucer=New cvpmBallStack
	with bsSaucer
		.InitSaucer Kicker1,19,150,15
		.KickAngleVar=2.5
		.InitExitSnd "fx_kicker","fx_kicker_enter"
	end with

	Set bsSaucer2=New cvpmBallStack
	with bsSaucer2
		.InitSaucer Kicker2,20,210,15
		.KickAngleVar=2.5
		.InitExitSnd "fx_kicker","fx_kicker_enter"
	end with

     vpmNudge.TiltSwitch = 8
     vpmNudge.Sensitivity = 1
     vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingshot)

	GILights 1
	special = 0

	PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1
	vpmMapLights aLights

	If Table1.ShowDT = False then
	for each objekt in backdropstuff:objekt.visible = False:Next
	End If

End Sub

Sub Table1_KeyDown(ByVal keycode)
	If keycode = LeftTiltKey Then Nudge 90, 2
	If keycode = RightTiltKey Then Nudge 270, 2
	If keycode = CenterTiltKey Then	Nudge 0, 2

	If vpmKeyDown(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.PullBack: PlaySound "fx_plungerpull",0,1,0.25,0.25: 	End If
	If FlippersAlwaysOn =1 Then
		If keycode = LeftFlipperKey Then LeftFlipper.RotateToEnd: PlaySound SoundFX("fx_flipperup",DOFContactors), 0, .67, -0.05, 0.05
		If keycode = RightFlipperKey Then RightFlipper.RotateToEnd: PlaySound SoundFX("fx_flipperup",DOFContactors), 0, .67, 0.05, 0.05
	End If
End Sub

Sub Table1_KeyUp(ByVal keycode)
	If keycode = PlungerKey Then Plunger.Fire: PlaySound "fx_plunger",0,1,0.25,0.25
	If FlippersAlwaysOn =1 Then
		If keycode = LeftFlipperKey Then LeftFlipper.RotateToStart: PlaySound SoundFX("fx_flipperdown",DOFContactors), 0, 1, -0.05, 0.05
		If keycode = RightFlipperKey Then RightFlipper.RotateToStart: PlaySound SoundFX("fx_flipperdown",DOFContactors), 0, 1, 0.05, 0.05
	End If

	If vpmKeyUp(keycode) Then Exit Sub

End Sub

Sub BallRelease_UnHit
	SpinWheel
End Sub


Sub GILights (enabled)
	Dim light
	For each light in GI:light.State = Enabled: Next
End Sub


'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
    PlaySound SoundFX("fx_slingshot",DOFContactors), 0, 1, 0.05, 0.05
	vpmTimer.PulseSw 24
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
    PlaySound SoundFX("fx_slingshot",DOFContactors),0,1,-0.05,0.05
	vpmTimer.PulseSw 15
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

'************************************************
' Solenoids
'************************************************
SolCallback(sBallrelease) = "bsTrough.SolOut"
SolCallback(sKicker1)="bsSaucer.SolOut"
SolCallback(sKicker2)="bsSaucer2.SolOut"
SolCallback(16)="vpmNudge.SolGameOn"
SolCallback(sLLFlipper)="vpmSolFlipper LeftFlipper, Nothing,"
SolCallback(sLRFlipper)="vpmSolFlipper RightFlipper, Nothing,"

'************************************************
' Switches
'************************************************
Sub Target1_Hit:vpmTimer.PulseSw 4:SpinWheel:End Sub
Sub Target2_Hit:vpmTimer.PulseSw 10:End Sub
Sub Target3_Hit:vpmTimer.PulseSw 4:SpinWheel:End Sub
Sub Target4_Hit:vpmTimer.PulseSw 12:End Sub
Sub Target5_Hit:vpmTimer.PulseSw 13:End Sub

Sub Target6_Hit:vpmTimer.PulseSw 17
If Special = 1 then SpinWheel
End Sub

Sub Wall15_Hit:vpmTimer.PulseSw 9:End Sub
Sub Wall16_Hit:vpmTimer.PulseSw 9:End Sub
Sub Wall17_Hit:vpmTimer.PulseSw 16:End Sub
Sub Wall18_Hit:vpmTimer.PulseSw 9:End Sub
Sub Wall19_Hit:vpmTimer.PulseSw 16:End Sub
Sub Wall20_Hit:vpmTimer.PulseSw 9:End Sub
Sub Wall21_Hit:vpmTimer.PulseSw 16:End Sub

Sub Drain_Hit:playsound "Drain":bsTrough.AddBall Me:End Sub

Sub LeftOutlane_Hit:Controller.Switch(18)=1:Me.TimerEnabled=1:End Sub
Sub LeftOutlane_Unhit:Controller.Switch(18)=0:End Sub
Sub LeftOutlane_Timer:Me.TimerEnabled=0:End Sub
Sub RightOutlane_Hit:Controller.Switch(18)=1:Me.TimerEnabled=1:End Sub
Sub RightOutlane_Unhit:Controller.Switch(18)=0:End Sub
Sub RightOutlane_Timer:Me.TimerEnabled=0:End Sub

Sub LeftInlane_Hit:Controller.Switch(16)=1:Me.TimerEnabled=1:End Sub
Sub LeftInlane_Unhit:Controller.Switch(16)=0:End Sub
Sub LeftInlane_Timer:Me.TimerEnabled=0:End Sub
Sub RightInlane_Hit:Controller.Switch(16)=1:Me.TimerEnabled=1:End Sub
Sub RightInlane_Unhit:Controller.Switch(16)=0:End Sub
Sub RightInlane_Timer:Me.TimerEnabled=0:End Sub

Sub Kicker1_Hit:bsSaucer.AddBall 0:End Sub
Sub Kicker2_Hit:bsSaucer2.AddBall 0:End Sub

Sub Bumper2_Hit:vpmTimer.PulseSw 33:RandomSoundBumper:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 21:RandomSoundBumper:End Sub
Sub Bumper1_Hit:vpmTimer.PulseSw 22:RandomSoundBumper:End Sub

Sub Spinner1_Spin:playsound "Spinner":vpmTimer.PulseSw 23:End Sub

Sub Trigger8_Hit:Controller.Switch(25)=1:End Sub
Sub Trigger8_unHit:Controller.Switch(25)=0:End Sub
Sub Trigger1_Hit:Controller.Switch(27)=1:End Sub
Sub Trigger1_unHit:Controller.Switch(27)=0:End Sub
Sub Trigger2_Hit:Controller.Switch(28)=1:End Sub
Sub Trigger2_unHit:Controller.Switch(28)=0:End Sub
Sub Trigger3_Hit:Controller.Switch(29)=1:End Sub
Sub Trigger3_unHit:Controller.Switch(29)=0:End Sub
Sub Trigger4_Hit:Controller.Switch(30)=1:End Sub
Sub Trigger4_unHit:Controller.Switch(30)=0:End Sub
Sub Trigger5_Hit:Controller.Switch(31)=1:End Sub
Sub Trigger5_unHit:Controller.Switch(31)=0:End Sub
Sub Trigger7_Hit:Controller.Switch(32)=1:End Sub
Sub Trigger7_unHit:Controller.Switch(32)=0:End Sub
Sub Trigger9_Hit:Controller.Switch(14)=1:End Sub
Sub Trigger9_unHit:Controller.Switch(14)=0:End Sub
Sub Trigger6_Hit:Controller.Switch(34)=1:End Sub
Sub Trigger6_unHit:Controller.Switch(34)=0:End Sub

Sub RandomSoundBumper()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "Bumper1", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "Bumper2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "Bumper3", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub

'*********************************************************************
'                Space Lab Wheel
'*********************************************************************

Sub SpinWheel
	Do While Wheel=OldValue
	Randomize
	Wheel=INT(RND*6+1)'randomize starting value of wheel
	Loop
	NewValue=Wheel+12
	playsound "wheelclick"
	WheelTimer.Enabled=1

End Sub

Sub WheelTimer_Timer
	NewValue=NewValue-1
	If NewValue>0 Then
	Select Case OldValue
		Case 1:Controller.Switch(35)=0: SkyLabWheel.ObjRotZ = 100:Special = 0
		Case 2:Controller.Switch(36)=0: SkyLabWheel.ObjRotZ = 130:Special = 1
		Case 3:Controller.Switch(37)=0: SkyLabWheel.ObjRotZ = 160:Special = 0
		Case 4:Controller.Switch(38)=0: SkyLabWheel.ObjRotZ = 190:Special = 0
		Case 5:Controller.Switch(39)=0: SkyLabWheel.ObjRotZ = 40:Special = 0
		Case 6:Controller.Switch(40)=0: SkyLabWheel.ObjRotZ = 70:Special = 0
	End Select
	OldValue=OldValue+1
	If OldValue>6 Then OldValue=1
	Select Case OldValue
		Case 1:Controller.Switch(35)=1: SkyLabWheel.ObjRotZ = 100:Special = 0
		Case 2:Controller.Switch(36)=1: SkyLabWheel.ObjRotZ = 130:Special = 1
		Case 3:Controller.Switch(37)=1: SkyLabWheel.ObjRotZ = 160:Special = 0
		Case 4:Controller.Switch(38)=1: SkyLabWheel.ObjRotZ = 190:Special = 0
		Case 5:Controller.Switch(39)=1: SkyLabWheel.ObjRotZ = 40:Special = 0
		Case 6:Controller.Switch(40)=1: SkyLabWheel.ObjRotZ = 70:Special = 0
	End Select
	Else
	WheelTimer.Enabled=0
	End If
End Sub
		'35 SpinLab Orion X3
		'36 SpinLab Special
		'37 SpinLab Extra Ball
		'38 SpinLab 50,000
		'39 SpinLab Ursa X3
		'40 SpinLab Comet 500


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

 Sub RollingTimer_Timer
    Dim BOT, b, ballpitch
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
            If BOT(b).z < 30 Then
                ballpitch = Pitch(BOT(b) )
            Else
                ballpitch = Pitch(BOT(b) ) * 100
            End If
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), AudioPan(BOT(b) ), 0, ballpitch, 1, 0
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
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 0, 0
End Sub

'*****************************************
'	ninuzzu's	FLIPPER SHADOWS (SPINNER)
'*****************************************

sub FlipperTimer_Timer()
	FlipperLSh.RotZ = LeftFlipper.currentangle
	FlipperRSh.RotZ = RightFlipper.currentangle
	SpinnerShadow.RotZ = Spinner1.currentangle
End Sub

'*****************************************
'	ninuzzu's	BALL SHADOW
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
        If BOT(b).X < Table1.Width/2 Then
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) + 6
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) - 6
        End If
        ballShadow(b).Y = BOT(b).Y + 12
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
End Sub

'******************************
' Diverse Collection Hit Sounds
'******************************

Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aMetals_Hit(idx):PlaySound "fx_MetalHit", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aRubber_Bands_Hit(idx):PlaySound "fx_rubber", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aRubber_Posts_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aWoods_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub

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

'************************************
'          LEDs Display
'************************************
Dim Digits(28)
Dim Patterns(11)
Dim Patterns2(11)

Patterns(0) = 0     'empty
Patterns(1) = 63    '0
Patterns(2) = 6     '1
Patterns(3) = 91    '2
Patterns(4) = 79    '3
Patterns(5) = 102   '4
Patterns(6) = 109   '5
Patterns(7) = 124   '6
Patterns(8) = 7     '7
Patterns(9) = 127   '8
Patterns(10) = 103  '9

Patterns2(0) = 128  'empty
Patterns2(1) = 191  '0
Patterns2(2) = 134  '1
Patterns2(3) = 219  '2
Patterns2(4) = 207  '3
Patterns2(5) = 230  '4
Patterns2(6) = 237  '5
Patterns2(7) = 252  '6
Patterns2(8) = 135  '7
Patterns2(9) = 255  '8
Patterns2(10) = 239 '9

Set Digits(0) = a0
Set Digits(1) = a1
Set Digits(2) = a2
Set Digits(3) = a3
Set Digits(4) = a4
Set Digits(5) = a5

Set Digits(6) = b0
Set Digits(7) = b1
Set Digits(8) = b2
Set Digits(9) = b3
Set Digits(10) = b4
Set Digits(11) = b5

Set Digits(12) = c0
Set Digits(13) = c1
Set Digits(14) = c2
Set Digits(15) = c3
Set Digits(16) = c4
Set Digits(17) = c5

Set Digits(18) = d0
Set Digits(19) = d1
Set Digits(20) = d2
Set Digits(21) = d3
Set Digits(22) = d4
Set Digits(23) = d5

Set Digits(24) = e0
Set Digits(25) = e1
Set Digits(26) = e2
Set Digits(27) = e3

Sub Leds_Timer()
     On Error Resume Next
    Dim ChgLED, ii, jj, chg, stat
    ChgLED = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
        For ii = 0 To UBound(ChgLED)
            chg = chgLED(ii, 1):stat = chgLED(ii, 2)
            For jj = 0 to 10
                If stat = Patterns(jj)OR stat = Patterns2(jj)then Digits(chgLED(ii, 0)).SetValue jj
            Next
        Next
    End IF
End Sub


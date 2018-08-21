Option Explicit

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

' Thalamus 2018-07-23
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' No special SSF tweaks yet.
' Wob 2018-08-08
' Added vpmInit Me to table init

LoadVPM "01300000","S6.VBS",3.1

'********************************************
'**     Game Specific Code Starts Here     **
'********************************************

Const UseSolenoids=2,UseLamps=1,UseSync=1
Const SSolenoidOn="solon",SSolenoidOff="soloff",SFlipperOn="FlipperUp",SFlipperOff="FlipperDown",SCoin="quarter"


'**************************************
'**     Bind Events To Solenoids     **
'**************************************

SolCallback(1)	= "bsTrough.SolOut"
SolCallBack(2)	= "bsLKicker.SolOut"
SolCallback(3)	= "GARtargets.SolDropUp"
SolCallback(4)	= "GORtargets.SolDropUp"
SolCallback(5)	= "mMagnet.MagnetOn="
SolCallback(6)  = "MagnetFlash"
SolCallback(14)	= "vpmSolSound SoundFX(""knocker"",DOFKnocker),"
'SolCallback(17)	= "vpmSolSound SoundFX(""bumper1"",DOFContactors),"
'SolCallback(18)	= "vpmSolSound SoundFX(""bumper2"",DOFContactors),"
'SolCallback(19)	= "vpmSolSound SoundFX(""bumper3"",DOFContactors),"
'SolCallback(20)	= "vpmSolSound SoundFX(""sling"",DOFContactors),"
'SolCallback(21)	= "vpmSolSound SoundFX(""sling"",DOFContactors),"
SolCallback(23) = "vpmNudge.SolGameOn"
SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
	If Enabled Then
		PlaySound SoundFX("fx_flipperup",DOFContactors):LeftFlipper.RotateToEnd
	Else
		PlaySound SoundFX("fx_flipperdown",DOFContactors):LeftFlipper.RotateToStart
	End If
End Sub

Sub SolRFlipper(Enabled)
	If Enabled Then
		PlaySound SoundFX("fx_flipperup",DOFContactors):RightFlipper.RotateToEnd
	Else
		PlaySound SoundFX("fx_flipperdown",DOFContactors):RightFlipper.RotateToStart
	End If
End Sub

Sub rubber_Hit(idx):PlaySound "rubber":End Sub


'*******************************
'**     Keyboard Handlers     **
'*******************************

Sub Gorgar_KeyDown(ByVal KeyCode)
	'If keycode = LeftFlipperKey Then
	'	PlaySound "fx_flipperup", 0, .67, -0.05, 0.05
	'End If

	'If keycode = RightFlipperKey Then
	'	PlaySound "fx_flipperup", 0, .67, 0.05, 0.05
	'End If

	If keycode = PlungerKey Then
		Plunger.PullBack
		PlaySound "plungerpull",0,1,0.25,0.25
	End If
    If vpmKeyDown(KeyCode) Then Exit Sub
End Sub


Sub Gorgar_KeyUp(ByVal KeyCode)
	'If keycode = LeftFlipperKey Then
	'	PlaySound "fx_flipperdown", 0, 1, -0.05, 0.05
	'End If

	'If keycode = RightFlipperKey Then
	'	PlaySound "fx_flipperdown", 0, 1, 0.05, 0.05
	'End If
	If keycode = PlungerKey Then
		Plunger.Fire
		PlaySound "plunger",0,1,0.25,0.25
	End If
    If vpmKeyUp(KeyCode) Then Exit Sub
End Sub


'********************************************
'**     Init The Table, Start VPinMAME     **
'********************************************

Dim GORtargets,GARtargets,bsLKicker,bsTrough,mtest,mMagnet,BallInPlay,plungerIM,x,ballOut


Sub Gorgar_Init
	vpmInit Me
	On Error Resume Next
	With Controller
	.GameName="grgar_l1"
	.SplashInfoLine="Gorgar v1.0 By drumndav AKA daveboy6"
	.HandleKeyboard=0
	.ShowTitle=0
	.ShowDMDOnly=1
	.ShowFrame=0
		If Gorgar.ShowDT = false then
			'Scoretext.Visible = false
			.Hidden = 1
		End If

		If Gorgar.ShowDT = true then
		'Scoretext.Visible = false
			.Hidden = 0
		End If
	End With
	Controller.Run
	If Err Then MsgBox Err.Description
		On Error Goto 0

	PinMAMETimer.Interval=PinMAMEInterval:PinMAMETimer.Enabled=1:vpmNudge.TiltSwitch=1:vpmNudge.Sensitivity=4
	vpmNudge.TiltObj=Array(LeftSling,RightSling,TopJet,LeftJet,BottomJet)

	PlaySound "gorgar_feel_my_power"

	Set bsTrough=New cvpmBallStack
		bsTrough.InitNoTrough BallRelease,9,85,7
		bsTrough.InitExitSnd SoundFX("Ballrelease",DOFContactors),SoundFX("solon",DOFContactors)


	Set GARtargets=New cvpmDropTarget
		GARtargets.InitDrop Array(Array(sw41),Array(sw42),Array(sw43)),Array(41,42,43)
		GARtargets.InitSnd SoundFX("droptargetdown",DOFContactors),SoundFX("droptargetreset",DOFContactors)
		GARtargets.AllDownSw=44

	Set GORtargets=New cvpmDropTarget
		GORtargets.InitDrop Array(Array(sw18),Array(sw19),Array(sw20)),Array(18,19,20)
		GORtargets.InitSnd SoundFX("droptargetdown",DOFContactors),SoundFX("droptargetreset",DOFContactors)
		GORtargets.AllDownSw=21

	Set bsLKicker=New cvpmBallStack
		bsLKicker.InitSaucer sw15,15,155,8
		bsLKicker.InitExitSnd SoundFX("popper_ball",DOFContactors),SoundFX("solon",DOFContactors)

	Set mMagnet=New cvpmMagnet
 		With mMagnet
 		.InitMagnet sw23, 26
		.GrabCenter=False
 		End With

end sub

'*****************************
'**     Switch Handling     **
'*****************************

Sub Drain_Hit : PlaySound "Drain" : bsTrough.AddBall Me : End Sub
Sub BallRelease_UnHit : Set BallInPlay = ActiveBall : End Sub
Sub LeftSling_Slingshot   : vpmTimer.PulseSw 12 : PlaySound SoundFX("left_slingshot",DOFContactors), 0, 1, 0.05, 0.05 : End Sub 'switch 12
Sub RightSling_Slingshot   : vpmTimer.PulseSw 36 : PlaySound SoundFX("right_slingshot",DOFContactors),0,1,-0.05,0.05 : End Sub 'switch 12
Sub sw13_Slingshot   : vpmTimer.PulseSwitch 13, 0, 0 : End Sub
Sub sw16_Slingshot   : vpmTimer.PulseSwitch 16, 0, 0 : End Sub
Sub sw17_Slingshot   : vpmTimer.PulseSwitch 17, 0, 0 : End Sub
Sub sw22_Slingshot   : vpmTimer.PulseSwitch 22, 0, 0 : End Sub
Sub sw24_Slingshot   : vpmTimer.PulseSwitch 24, 0, 0 : End Sub
Sub sw25_Slingshot   : vpmTimer.PulseSwitch 25, 0, 0 : End Sub
Sub sw40_Slingshot   : vpmTimer.PulseSwitch 40, 0, 0 : End Sub
Sub sw30_Spin : vpmTimer.PulseSwitch 30, 0, 0 :PlaySound "Spinner",0,.25,0,0.25: End Sub



Sub sw15_Hit
	bsLKicker.AddBall 0
End Sub 				'switch 15

Sub sw18_Hit:GORtargets.Hit 1:End Sub                         'switch 18
Sub sw19_Hit:GORtargets.Hit 2:End Sub                         'switch 19
Sub sw20_Hit:GORtargets.Hit 3:End Sub                         'switch 20
Sub sw23_Hit:Controller.Switch(23)=1:mMagnet.AddBall ActiveBall:End Sub
Sub sw23_UnHit:Controller.Switch(23)=0:mMagnet.RemoveBall ActiveBall:End Sub
Sub LeftJet_Hit : vpmTimer.PulseSwitch 37, 0, 0 : PlaySound SoundFX("fx_bumper3",DOFContactors): End Sub
Sub TopJet_Hit : vpmTimer.PulseSwitch 38, 0, 0 : PlaySound SoundFX("fx_bumper3",DOFContactors): End Sub
Sub BottomJet_Hit : vpmTimer.PulseSwitch 39, 0, 0 : PlaySound SoundFX("fx_bumper3",DOFContactors) :End Sub
Sub sw41_Hit:GARtargets.Hit 1:End Sub                        'switch 41
Sub sw42_Hit:GARtargets.Hit 2:End Sub                        'switch 42
Sub sw43_Hit:GARtargets.Hit 3:End Sub                        'switch 43

dim intensScale:intensScale=0
dim scaleStep:scaleStep=0.8
 Sub MagnetFlash(Enabled)
		If Enabled Then
			pitLamp.state=1
			pitLamp2.state=1
			pitLamp3.state=1
			intensScale = intensScale + scaleStep
			if intensScale>=1.0 Then intensScale=1.0
			pitLamp.IntensityScale = intensScale
			pitLamp2.IntensityScale = intensScale
			pitLamp3.IntensityScale = intensScale
		Else
			intensScale = intensScale - scaleStep
			if intensScale<0.0 then intensScale=0.0:pitLamp.state=0:pitLamp2.state=0:pitLamp3.state=0:end if
			pitLamp.IntensityScale = intensScale
			pitLamp2.IntensityScale = intensScale
			pitLamp3.IntensityScale = intensScale
		End If
 End Sub


'***********************************
'**     Map Lights Into Array     **
'** Set Unmapped Lamps To Nothing **
'***********************************

Set Lights(1)=Light1 'Same Player Shoots Again (Playfield)
Set Lights(2)=Light2 'Left Special
Set Lights(3)=Light3 'Right special
Set Lights(4)=Light4 '2X
Set Lights(5)=Light5 '3X
Set Lights(6)=Light6 'Star 1
Set Lights(7)=Light7 'Star 2
Set Lights(8)=Light8 '1,000 Bonus
Set Lights(9)=Light9 '2.000 Bonus
Set Lights(10)=Light10 '3,000 Bonus
Set Lights(11)=Light11 '4,000 Bonus
Set Lights(12)=Light12 '5,000 Bonus
Set Lights(13)=Light13 '6,000 Bonus
Set Lights(14)=Light14 '7,000 Bonus
Set Lights(15)=Light15 '8,000 Bonus
Set Lights(16)=Light16 '9,000 Bonus
'Light17 NOT USED
Set Lights(18)=Light18 '10,000 Bonus
Set Lights(19)=Light19 '20,000 Bonus
Set Lights(20)=Light20 'A
Set Lights(21)=Light21 'B
Set Lights(22)=Light22 'C
Set Lights(23)=Light23 'D
Set Lights(24)=Light24 'E
Set Lights(25)=Light25 '1 Target
Set Lights(26)=Light26 '2 Target
Set Lights(27)=Light27 '3 Target
Set Lights(28)=Light28 '4 Target
Set Lights(29)=Light29 '1 Target Arrow
Set Lights(30)=Light30 '2 Target Arrow
Set Lights(31)=Light31 '3 Target Arrow
Set Lights(32)=Light32 '4 Target Arrow
Set Lights(33)=Light33 'Magnet 5,000
Set Lights(34)=Light34 'Magnet 10,000
Set Lights(35)=Light35 'Magnet 20,000
Set Lights(36)=Light36 'Magnet 30,000
Set Lights(37)=Light37 'Magnet 50,000
Set Lights(38)=Light70  'Top Jet Bumper
Set Lights(39)=Light71 'Left Jet Bumper
Set Lights(40)=Light69 'Bottom Jet Bumper
Set Lights(41)=Light41 'GAR 5,000 When Lit
Set Lights(42)=Light42 'GOR
Set Lights(43)=Light43 'GAR
Set Lights(44)=Light44 'Eject Hole 10,000
Set Lights(45)=Light45 'Eject Hole 15,000
Set Lights(46)=Light46 'Eject Hole Extra Ball
'Light47 NOT USED
Set Lights(48)=Light48 'Spinner 1,000 When Lit
Set Lights(56)=Light88 'Credit Light



'**********************************
'**     Table-Specific Stuff     **
'**********************************

Sub sw31_Hit:vpmTimer.PulseSw 31:PlaySound SoundFX("droptargetR",DOFContactors):End Sub
Sub sw32_Hit:vpmTimer.PulseSw 32:PlaySound SoundFX("droptargetR",DOFContactors):End Sub
Sub sw33_Hit:vpmTimer.PulseSw 33:PlaySound SoundFX("droptargetR",DOFContactors):End Sub
Sub sw14_Hit:vpmTimer.PulseSw 14:PlaySound SoundFX("droptargetL",DOFContactors):End Sub


Sub sw10_Hit:   Controller.Switch(10)=1 : End Sub
Sub sw10_unHit: Controller.Switch(10)=0 : End Sub
Sub sw11_Hit:   Controller.Switch(11)=1 : End Sub
Sub sw11_unHit: Controller.Switch(11)=0 : End Sub
Sub sw34_Hit:   Controller.Switch(34)=1 : End Sub
Sub sw34_unHit: Controller.Switch(34)=0 : End Sub
Sub sw35_Hit:   Controller.Switch(35)=1 : End Sub
Sub sw35_unHit: Controller.Switch(35)=0 : End Sub
Sub sw24_Hit:   Controller.Switch(26)=1 : End Sub
Sub sw24_unHit: Controller.Switch(26)=0 : End Sub
Sub sw27_Hit:   Controller.Switch(27)=1 : End Sub
Sub sw27_unHit: Controller.Switch(27)=0 : End Sub
Sub sw28_Hit:   Controller.Switch(28)=1 : End Sub
Sub sw28_unHit: Controller.Switch(28)=0 : End Sub

'*****************************************
'			FLIPPER SHADOWS
'*****************************************

sub FlipperTimer_Timer()
	FlipperLSh.RotZ = LeftFlipper.currentangle
	FlipperRSh.RotZ = RightFlipper.currentangle

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
        If BOT(b).X < gorgar.Width/2 Then
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (gorgar.Width/2))/7)) + 5
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (gorgar.Width/2))/7)) - 5
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


Sub Pins_Hit (idx)
	PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
	PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Hit (idx)
	PlaySound "metalhit_medium", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "gorgar" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / gorgar.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "gorgar" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / gorgar.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "gorgar" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / gorgar.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioFade(ball) ' Can this be together with the above function ?
  Dim tmp
  tmp = ball.y * 2 / gorgar.height-1
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

	if sw43.isDropped=1 then
		dropLight1.state=1
	Else
		light52.state=1
		dropLight1.state=0
	end If
	if sw42.isDropped=1 then
		dropLight2.state=1
	Else
		dropLight2.state=0
	end If

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
  If gorgar.VersionMinor > 3 OR gorgar.VersionMajor > 10 Then
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
  Else
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0
  End if
End Sub


' Thalamus : Exit in a clean and proper way
Sub Gorgar_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


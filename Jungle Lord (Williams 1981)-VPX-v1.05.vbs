'Jungle Lord / IPD No. 1338 / February, 1981 / 4 Players
'VPX 1.03 by Tony The Wrench 2016
'Credits
'Plastics and Playfield from Francisco666
'Script and other elements from JPSalas' Black Knight VPX, and Kristian (VP9 Original Script)
'Additional Script from 32assassin
Option Explicit
Randomize


' Thalamus 2018-07-23
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

Const VolGates  = 1    ' Gates volume.
Const VolMetal  = 1    ' Metals volume.
Const VolRH     = 1    ' Rubber hits volume.
Const VolWood   = 1    ' Woods volume.
Const VolKick   = 1    ' Kicker volume.
Const VolFlip   = 1    ' Flipper volume.


On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01000100", "s7.vbs", 2.2

'**********************************************************************************************************
'Standard Definitions
'**********************************************************************************************************

Const UseSolenoids  = 2
Const UseLamps      = 1

Const SSolenoidOn   ="fx_solenoid"
Const SSolenoidOff  ="fx_solenoidoff"
Const SCoin			="fx_coin"

'**********************************************************************************************************
'Initiate Table
'**********************************************************************************************************

Dim bsTrough, LMAG, RMAG, RStep, Lstep, URStep, dtLL, dtLR, dtUP, LowerKicker, UpperKicker, MBLaunch

Const cGameName = "jngld_l2"

Sub Table1_Init
	vpmInit Me
		With Controller
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
		.SplashInfoLine = "Jungle Lord (Williams 1981)"
		.HandleMechanics=0
		.HandleKeyboard=0
		.ShowDMDOnly=1
		.ShowFrame=0
		.ShowTitle=0
        .Hidden = False
		If Err Then MsgBox Err.Description
        On Error Resume Next
        .SolMask(0) = 0
        vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the timer to renable all the solenoids after 2 seconds
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
        On Error Goto 0
    End With
End Sub

	'Main Timer
	PinMAMETimer.Interval=PinMAMEInterval
	PinMAMETimer.Enabled=1

    ' Nudging
    vpmNudge.TiltSwitch = 1
    vpmNudge.Sensitivity = 1
    vpmNudge.TiltObj = Array(LeftSlingshot, RightSlingshot, UpperRightSlingShot)

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 42, 9, 10, 0, 0, 0, 0, 0
        .InitKick BallRelease, 90, 7
        .InitEntrySnd "fx_Solenoid", "fx_Solenoid"
        .InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .Balls = 2
    End With

'**********************************************************************************************************
'Magnets
'**********************************************************************************************************

    ' Left Magnet
    Set LMAG = New cvpmMagnet
    With LMAG
        .InitMagnet Trigger50, 9
        .Solenoid = 21
        .CreateEvents "LMAG"
    End With

    ' Right Magnet
    Set RMAG = New cvpmMagnet
    With RMAG
        .InitMagnet Trigger49, 9
        .Solenoid = 22
        .CreateEvents "RMAG"
    End With

'**********************************************************************************************************
' Flippers
'**********************************************************************************************************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFX("fx_flipperup", DOFContactors), LeftFlipper, VolFlip
        PlaySoundAtVol SoundFX("fx_flipperup", DOFContactors), LeftFlipper2, VolFlip
        LeftFlipper.RotateToEnd
        LeftFlipper2.RotateToEnd
    Else
        PlaySoundAtVol SoundFX("fx_flipperdown", DOFContactors), LeftFlipper, VolFlip
        PlaySoundAtVol SoundFX("fx_flipperdown", DOFContactors), LeftFlipper2, VolFlip
        LeftFlipper.RotateToStart
        LeftFlipper2.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFX("fx_flipperup", DOFContactors), RightFlipper, VolFlip
        PlaySoundAtVol SoundFX("fx_flipperup", DOFContactors), RightFlipper2, VolFlip
        RightFlipper.RotateToEnd
        RightFlipper2.RotateToEnd
    Else
        PlaySoundAtVol SoundFX("fx_flipperdown", DOFContactors), RightFlipper, VolFlip
        PlaySoundAtVol SoundFX("fx_flipperdown", DOFContactors), RightFlipper2, VolFlip
        RightFlipper.RotateToStart
        RightFlipper2.RotateToStart
    End If
End Sub

'**********************************************************************************************************
'Keys
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)
    If KeyDownHandler(keycode) Then Exit Sub
    If keycode = PlungerKey Then:Plunger.Pullback:PlaySoundAtVol"fx_plungerpull", Plunger, 1
    If keycode = LeftMagnaSave Then:Controller.Switch(50) = 1:End If
    If keycode = RightMagnaSave Then:Controller.Switch(49) = 1:End If
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
    If KeyUpHandler(keycode) Then Exit Sub
    If keycode = PlungerKey Then:Plunger.Fire:PlaySoundAtVol"fx_plunger", Plunger, 1
    If keycode = LeftMagnaSave Then:Controller.Switch(50) = 0:End If
    If keycode = RightMagnaSave Then:Controller.Switch(49) = 0:End If
End Sub

'**********************************************************************************************************
'Lights
'**********************************************************************************************************

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
Set Lights(24)=Light24 'mb lights
Set Lights(25)=Light25
Set Lights(26)=Light26
Set Lights(27)=Light27
Set Lights(28)=Light28
Set Lights(29)=Light29
Set Lights(30)=Light30
Set Lights(31)=Light31
Set Lights(32)=Light32 'mb lights
Set Lights(33)=Light33
Set Lights(34)=Light34
Set Lights(35)=Light35
Set Lights(36)=Light36
Set Lights(37)=Light37
Set Lights(38)=Light38 'mb lights
Set Lights(39)=Light39
Set Lights(40)=Light40
Set Lights(41)=Light41
Set Lights(42)=Light42
Set Lights(43)=Light43
Set Lights(44)=Light44

Lights(45)=Array(Light45,Light45a)

Set Lights(46)=Light46
Set Lights(47)=Light47 'mb lights
Set Lights(48)=Light48 'mb lights
Set Lights(49)=Light49
Set Lights(50)=Light50
Set Lights(51)=Light51
Set Lights(52)=Light52
Set Lights(53)=Light53
Set Lights(54)=Light54
Set Lights(55)=Light55
Set Lights(56)=Light56
Set Lights(57)=Light57
Set Lights(58)=Light58
Set Lights(59)=Light59
Set Lights(60)=Light60
Set Lights(61)=Light61
Set Lights(62)=Light62
Set Lights(63)=Light63
Set Lights(64)=Light64

'**********************************************************************************************************
'GI Lighting
'**********************************************************************************************************

Dim xx, UpdateGI

Sub UpdateGITimer
	If UpdateGI = 0 then
		For each xx in GI:xx.State = 0: Next
		PlaySound "fx_relay"
	Else
		For each xx in GI:xx.State = 1: Next
		PlaySound "fx_relay"
	End If
End Sub

'**********************************************************************************************************
'SlingShot Animations
'**********************************************************************************************************

Sub RightSlingShot_Slingshot
    PlaySoundAtVol SoundFX("fx_slingshot",DOFContactors), sling1, 1
    RSling.Visible = 0
    RSling1.Visible = 1
    SLING1.TransZ = -20
    RStep = 0
	vpmtimer.pulsesw 28
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSling1.Visible = 0:RSling2.Visible = 1:SLING1.TransZ = -10
        Case 4:RSling2.Visible = 0:RSling.Visible = 1:SLING1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
    PlaySoundAtVol SoundFX("fx_slingshot",DOFContactors), sling2, 1
    LSling.Visible = 0
    LSling1.Visible = 1
    SLING2.TransZ = -20
    LStep = 0
	vpmtimer.pulsesw 12
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSling1.Visible = 0:LSling2.Visible = 1:SLING2.TransZ = -10
        Case 4:LSling2.Visible = 0:LSling.Visible = 1:SLING2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub UpperRightSlingShot_Slingshot
    PlaySoundAtVol SoundFX("fx_slingshot",DOFContactors), sling3, 1
    URSling.Visible = 0
    URSling1.Visible = 1
    SLING3.TransZ = -20
    URStep = 0
	vpmtimer.pulsesw 40
    UpperRightSlingShot.TimerEnabled = 1
End Sub

Sub UpperRightSlingShot_Timer
    Select Case URStep
        Case 3:URSling1.Visible = 0:URSling2.Visible = 1:SLING3.TransZ = -10
        Case 4:URSling2.Visible = 0:URSling.Visible = 1:SLING3.TransZ = 0:UpperRightSlingShot.TimerEnabled = 0
    End Select
    URStep = URStep + 1
End Sub

'**********************************************************************************************************
'Drop Targets
'**********************************************************************************************************

    ' Lower Left DropTargets
    Set dtLL = New cvpmDropTarget
    With dtLL
        .InitDrop Array(sw29, sw30, sw31), Array(29, 30, 31)
        .InitSnd SoundFX("fx_droptarget", DOFContactors), SoundFX("fx_resetdrop", DOFContactors)
        .CreateEvents "dtLL"
    End With

    ' Lower Right DropTargets
    Set dtLR = New cvpmDropTarget
    With dtLR
        .InitDrop Array(sw25, sw26, sw27), Array(25, 26, 27)
        .InitSnd SoundFX("fx_droptarget", DOFContactors), SoundFX("fx_resetdrop", DOFContactors)
        .CreateEvents "dtLR"
    End With

    ' Upper Playfield DropTargets
    Set dtUP = New cvpmDropTarget
    With dtUP
        .InitDrop Array(sw33, sw34, sw35, sw36, sw37), Array(33, 34, 35, 36, 37)
        .InitSnd SoundFX("fx_droptarget", DOFContactors), SoundFX("fx_resetdrop", DOFContactors)
        .CreateEvents "dtUP"
    End With

'**********************************************************************************************************
'Drop Target Timers
'**********************************************************************************************************

Sub dtUPreset(Enabled):Timer1.Enabled = true:End Sub
Sub Timer1_Timer()
    Timer1.Enabled = false
    Controller.Switch(33) = true
	Controller.Switch(34) = true
	Controller.Switch(35) = true
	Controller.Switch(36) = true
	Controller.Switch(37) = true
	sw33.isdropped = true
    sw34.isdropped = true
    sw35.isdropped = true
    sw36.isdropped = true
    sw37.isdropped = true
End Sub

Sub dtsw33(Enabled):sw33.isdropped = false:Controller.Switch (33)= false:End Sub
Sub dtsw34(Enabled):sw34.isdropped = false:Controller.Switch (34)= false:End Sub
Sub dtsw35(Enabled):sw35.isdropped = false:Controller.Switch (35)= false:End Sub
Sub dtsw36(Enabled):sw36.isdropped = false:Controller.Switch (36)= false:End Sub
Sub dtsw37(Enabled):sw37.isdropped = false:Controller.Switch (37)= false:End Sub

Sub dtLLreset(Enabled):Timer2.Enabled = true:End Sub
Sub Timer2_Timer()
    Timer2.Enabled = false
    Controller.Switch(29) = false
	Controller.Switch(30) = false
	Controller.Switch(31) = false
    sw29.isdropped = false
    sw30.isdropped = false
    sw31.isdropped = false
End Sub

'**********************************************************************************************************
' Stand Up Targets
'**********************************************************************************************************

Sub sw17_Hit:vpmTimer.PulseSw 17:End Sub
Sub sw18_Hit:vpmTimer.PulseSw 18:End Sub
Sub sw19_Hit:vpmTimer.PulseSw 19:End Sub

'**********************************************************************************************************
' Rollovers
'**********************************************************************************************************

Sub sw20_Hit:Controller.Switch(20) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw20_UnHit:Controller.Switch(20) = 0:End Sub

Sub sw21_Hit:Controller.Switch(21) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw21_UnHit:Controller.Switch(21) = 0:End Sub

Sub sw22_Hit:Controller.Switch(22) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw22_UnHit:Controller.Switch(22) = 0:End Sub

Sub sw23_Hit:Controller.Switch(23) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw23_UnHit:Controller.Switch(23) = 0:End Sub

Sub sw24_Hit:Controller.Switch(24) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw24_UnHit:Controller.Switch(24) = 0:End Sub

Sub sw32_Hit:Controller.Switch(32) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw32_UnHit:Controller.Switch(32) = 0:End Sub

Sub sw43_Hit:Controller.Switch(43) = 1:End Sub
Sub sw43_UnHit:Controller.Switch(43) = 0:End Sub

Sub sw44_Hit:Controller.Switch(44) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw44_UnHit:Controller.Switch(44) = 0:End Sub

Sub sw13_Hit:Controller.Switch(13) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw13_UnHit:Controller.Switch(13) = 0:End Sub

Sub sw14_Hit:Controller.Switch(14) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw14_UnHit:Controller.Switch(14) = 0:End Sub

Sub sw15_Hit:Controller.Switch(15) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw15_UnHit:Controller.Switch(15) = 0:End Sub

Sub sw16_Hit:Controller.Switch(16) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw16_UnHit:Controller.Switch(16) = 0:End Sub

'**********************************************************************************************************
'Kickers
'**********************************************************************************************************

	Kicker3.CreateSizedBall (16)

Sub Kicker1_Hit:LowerKicker.AddBall 0:UpdateGI = 0:UpdateGITimer:End Sub
Sub Kicker1_UnHit:UpdateGI = 1:UpdateGITimer:End Sub

   Set LowerKicker = New cvpmBallStack
   LowerKicker.InitSaucer Kicker1,39,306,11
   LowerKicker.KickAngleVar=2

Sub Kicker2_Hit:UpperKicker.AddBall 0:UpdateGI = 0:UpdateGITimer:End Sub
Sub Kicker2_UnHit:UpdateGI = 1:UpdateGITimer:End Sub

   Set UpperKicker = New cvpmBallStack
   UpperKicker.InitSaucer Kicker2,38,270,2
   UpperKicker.KickAngleVar=5

Sub MiniKicker(Enabled)
    MBLaunch = 6*rnd(1)
    MBLaunch = Int(MBLaunch)
	Kicker3.Kick 0, MBLaunch +12
    PlaysoundAtVol "fx_solenoid", kicker3, VolKick
End Sub

Sub UpperEnter_Hit
    UpperEnter.Destroyball
    UpperExit.CreateBall
    UpperExit.Kick 180, 7
End Sub

Sub Drain_Hit:PlaysoundAtVol "fx_drain", drain, 1:bsTrough.AddBall Me:End Sub

'**********************************************************************************************************
'Solenoids
'**********************************************************************************************************

SolCallback(1) = "bsTrough.SolIn"
SolCallback(2) = "bsTrough.SolOut"
SolCallback(4) = "dtLLreset"
Solcallback(5) = "dtLR.soldropup"
SolCallback(6) = "vpmSolSound""Buzz"","
SolCallback(7) = "LowerKicker.SolOut"
SolCallback(8) = "UpperKicker.SolOut"
SolCallback(9) = "dtsw33"
SolCallback(10) = "dtsw34"
SolCallback(11) = "dtsw35"
SolCallback(12) = "dtsw36"
SolCallback(13) = "dtsw37"
SolCallback(14) = "dtUPreset"
SolCallback(15) = "vpmSolSound""Bell"","
SolCallback(20) = "MiniKicker"

'**********************************************************************************************************
' Additional Sounds
'**********************************************************************************************************

Sub Metals_Hit (idx)
	PlaySound "metalhit_thin", 0, Vol(ActiveBall)*VolMetal, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
	PlaySound "fx_gate", 0, Vol(ActiveBall)*VolGates, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Woods_Hit(idx)
	PlaySound "fx_woodhit", 0, Vol(ActiveBall)*VolWood, pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then
		PlaySound "fx_rubber_band", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
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

Sub PlaySoundAtVol(sound, tableobj, Volum)
  PlaySound sound, 1, Volum, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
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
  Vol = Csng(BallVel(ball) ^2 / VolDiv)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
  Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
  BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

'**********************************************************************************************************
'      JP's VP10 Rolling Sounds
'**********************************************************************************************************

Const tnob = 5 ' total number of balls in this table is 2, but always use a higher number here because of the timing
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub RollingUpdate()
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


Option Explicit
Randomize

Const BallMass = 1.25
Const UseVPMModSol=1


' Thalamus 2018-07-19
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Added InitVpmFFlipsSAM
' Thalamus 2018-09-09 : Improved directional sounds

' !! NOTE : Table not verified yet !!

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
Const VolRH     = 1    ' Rubber hits volume.
Const VolPo     = 1    ' Rubber posts volume.
Const VolPi     = 1    ' Rubber pins volume.
Const VolTarg   = 1    ' Targets volume.
Const VolWood   = 1    ' Woods volume.
Const VolKick   = 1    ' Kicker volume.
Const VolSpin   = 1.5  ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.


On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="avs_170c",UseSolenoids=1,UseLamps=0,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff", SCoin="coin"

LoadVPM "01560000", "sam.VBS", 3.10
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
SolCallback(1)  = "solTrough"
SolCallback(2)  = "solAutofire"
SolCallback(3)  = "solHulkCCW"
SolCallback(4)  = "solHulkCW"
SolCallback(5)  = "bsHole.SolOut"
SolCallback(6)  = "dtBank.SolDropUp" 'Drop Targets
SolCallback(7)  = "OrbitControlGate"
SolCallback(12) = "solRampControlGate"
SolCallback(15) = "SolLFlipper"
SolCallback(16) = "SolRFlipper"
SolCallBack(17) = "solHulkArms" 'Hulk Arms
SolCallback(22) = "solLokiLockup"
SolCallBack(41)="solcheck 41,"
SolCallBack(42)="solcheck 42,"
SolCallBack(43)="solcheck 43,"
SolCallBack(44)="solcheck 44,"
SolCallBack(45)="solcheck 45,"
SolCallBack(46)="solcheck 46,"
SolCallBack(47)="solcheck 47,"
SolCallBack(48)="solcheck 48,"

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors),LeftFlipper, VolFlip:LeftFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors),LeftFlipper, VolFlip:LeftFlipper.RotateToStart
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors), RightFlipper, VolFlip:RightFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors), RightFlipeer, VolFlip:RightFlipper.RotateToStart
     End If
End Sub

'Solenoid Controlled Flashers
SolModCallback(18) = "SetLampMod 118," 'Left Flasher
SolModCallback(19) = "SetLampMod 119," 'Right Flasher
SolModCallback(20) = "SetLampMod 120," 'Sling Flashers
SolModCallback(21) = "SetLampMod 121," 'Hulk Flashers
SolModCallback(25) = "SetLampMod 125," 'Pop Bumper Flasher
SolModCallback(26) = "SetLampMod 126," 'Teseract Cube Flasher
SolModCallback(27) = "SetLampMod 127," 'Backwall Flasher
SolModCallback(28) = "SetLampMod 128," 'Backwall Flasher
SolModCallback(29) = "SetLampMod 129," 'Backwall Flasher
SolModCallback(30) = "SetLampMod 130," 'Backwall Flasher
SolModCallback(31) = "SetLampMod 131," 'Backwall Flasher
SolModCallback(32) = "SetLampMod 132," 'Backwall Flasher


'**********************************************************************************************************
'Solenoid Controlled toys
'**********************************************************************************************************

sub solcheck(value,enabled)
dim solx
    Select Case value
         Case 3: solx = "Hulk CCW"
         Case 4: solx = "Hulk CW"
         Case 5: solx = "Hulk Eject"
         Case 7: solx = "Orbit Control Gate"
         Case 8: solx = "Shaker"
         Case 12: solx = "Ramp Control Gate Left"
         Case 13: solx = "LeftSling"
         Case 14: solx = "RightSling"
         Case 17: solx = "Hulk Arms"
         Case 18: solx = "Flash Left"
         Case 19: solx = "Flash Right"
         Case 20: solx = "Flash Slings"
         Case 21: solx = "Flash Hulk"
         Case 23: solx = "Hulk Magnet"
         Case 24: solx = "Coin Meter"
         Case 25: solx = "Flash Pop Bumper"
         Case 27: solx = "Flash Back1"
         Case 28: solx = "Flash Back2"
         Case 29: solx = "Flash Back3"
         Case 30: solx = "Flash Back4"
         Case 31: solx = "Flash Back5"
         Case 32: solx = "Flash Back6"
    End Select

End Sub


Sub solTrough(Enabled)
	If Enabled Then
		bsTrough.ExitSol_On
		vpmTimer.PulseSw 22
	End If
End Sub

sub solHulkCCW (enabled)
	If enabled Then
		HulkPrim.Roty = HulkPrim.roty + 10
		HulkArm1.RotateToEnd
	Else
		HulkPrim.Roty = 0
		HulkArm1.RotateToStart
	End if
End Sub

sub solHulkCW (enabled)
	If enabled Then
		HulkPrim.Roty = HulkPrim.roty - 10
		HulkArm2.RotateToEnd
	Else
		HulkPrim.Roty = 0
		HulkArm2.RotateToStart
	End If
End Sub

Sub solHulkArms (enabled)
	If enabled Then
		HulkArm1.Enabled = 0
		HulkArm2.Enabled = 0
	Else
		HulkArm1.Enabled = 1
		HulkArm2.Enabled = 1
	End If
End Sub

Dim armsup:armsup=False
Sub solHulkArms (enabled)
	If enabled Then
		HulkArm1.Enabled = 0
		HulkArm2.Enabled = 0
		armsup = True
	Else
		HulkArm1.Enabled = 1
		HulkArm2.Enabled = 1
		armsup = False
	End If
End Sub


Sub solLokiLockup(Enabled)
	If Enabled Then
 		post49.isdropped = 1: playsoundAtVol "fx_woodhit2", post49, VolWood: 'DEBUG.PRINT "POST1 DOWN"
	Else
		post49.isdropped = 0: playsoundAtVol "fx_woodhit", post49, VolWood
	End If
End Sub


Sub solRampControlGate(Enabled)
	If Enabled Then
		RampControlGate.IsDropped = 0: PlaySoundAtVol "fx_diverter", RampControlGate, 1
	Else
		RampControlGate.IsDropped = 1: PlaySoundAtVol "fx_diverter", RampControlGate, 1
	End If
End Sub

Sub OrbitControlGate(enabled)
	If Enabled Then
		gate1.open=True
	Else
		gate1.open=False
	End If
End Sub

Sub solAutofire(Enabled)
	If Enabled Then
		PlungerIM.AutoFire
		playsoundAtVol SoundFX("Popper",DOFContactors), Plunger, 1
	End If
End Sub


'Stern-Sega GI
set GICallback = GetRef("UpdateGI")
Sub UpdateGI(no, Enabled)
			If Enabled Then
				Dim ig
				For each ig in GI:ig.State = 1:Next
				For each ig in GIbulbs:ig.State = 1:Next
				For each ig in GIBwallPrim:ig.image = "bulbcover1_yellowON":Next  'Backwall Primitive bulbs
				For each ig in GIBwallFLH:ig.visible=1:Next		'Backwall Flashers over primitive bulbs
				PlaySound "fx_relay"
			Else
				For each ig in GI:ig.State = 0:Next
				For each ig in GIbulbs:ig.State = 0:Next
				For each ig in GIBwallPrim:ig.image = "bulbcover1_yellow":Next
				For each ig in GIBwallFLH:ig.visible=0:Next
				PlaySound "fx_relay"
			End If
End Sub

'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************

Dim bsTrough, dtBank, bsHole, mHulkMag
Dim PlungerIM

Sub Table1_Init
	vpmInit Me
	On Error Resume Next
		With Controller
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
		.SplashInfoLine = "Avengers Pro Stern 2012"&chr(13)&"You Suck"
		.HandleMechanics=0
		.HandleKeyboard=0
		.ShowDMDOnly=1
		.ShowFrame=0
		.ShowTitle=0
        .hidden = 0
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With
     On Error Goto 0

	PinMAMETimer.Interval=PinMAMEInterval
	PinMAMETimer.Enabled = 1

    vpmNudge.TiltSwitch=-7
    vpmNudge.Sensitivity=7
    vpmNudge.TiltObj=Array(Bumper1b,Bumper2b,Bumper3b,LeftSlingshot,RightSlingshot)

    '**Trough
    Set bsTrough = New cvpmBallStack
		bsTrough.InitSw 0, 21, 20, 19, 18, 0, 0, 0
		bsTrough.InitKick BallRelease, 80	, 8
		bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
		bsTrough.Balls = 4

	Set dtBank = New cvpmDropTarget 'HULK
		dtBank.InitDrop Array(Sw52,Sw53,Sw54,Sw55),Array(52,53,54,55) 'Walls<->Switch No.
		dtBank.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

    'Hulk Eject
    Set bsHole = New cvpmSaucer
        bsHole.InitKicker HulkEject, 62, 25, 11, 2
       'bsHole.InitSounds "kicker_enter_left","fx_solenoid","popper_ball"

	Set mHulkMag= New cvpmMagnet
		mHulkMag.InitMagnet HulkMag, 16
		mHulkMag.GrabCenter = True
 		mHulkMag.solenoid=23
		mHulkMag.CreateEvents "mHulkMag"

	' Impulse Plunger
	Const IMPowerSetting = 55
	Const IMTime = 0.6
	Set plungerIM = New cvpmImpulseP
	With plungerIM
		.InitImpulseP swplunger, IMPowerSetting, IMTime
		.Random .5
		.InitExitSnd "fx_plunger2", "fx_plunger"
		.CreateEvents "plungerIM"
	End With

    Dim obj
    For Each obj In colLampPoles:obj.IsDropped = 1:Next
    For Each obj In colLampPoles2:obj.IsDropped = 1:Next
    lampSpinSpeed = 0:lampLastPos = -1:SpinTimer.Enabled = True ' force update

	' Loki Lock Init - Optos used so switch is opposite (1=no ball, 0=ball).  Default to 1
	Controller.Switch(51) = 1
	Controller.Switch(50) = 1
	Controller.Switch(49) = 1
	Post51.IsDropped = 1
	Post50.IsDropped = 1
	Post49.IsDropped = 0
	RampControlGate.IsDropped = 1
  InitVpmFFlipsSAM
End Sub


'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)
	If KeyDownHandler(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Pullback:playsoundAtVol"plungerpull",plunger,1
	if KeyCode = LeftTiltKey Then Nudge 90, 4
	if KeyCode = RightTiltKey Then Nudge 270, 4
	if KeyCode = CenterTiltKey Then Nudge 0, 4
	If Keycode = StartGameKey Then Controller.Switch(16) = 1
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
	If KeyUpHandler(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Fire:PlaySoundAtVol"plunger",plunger,1
	If Keycode = StartGameKey Then Controller.Switch(16) = 0
End Sub

'**********************************************************************************************************

 ' Drain hole and kickers
Sub Drain_Hit:bsTrough.addball me : playsoundAtVol"drain",drain,1 : End Sub
Sub HulkEject_Hit : bsHole.addball me : playsoundAtVol "popper_ball",HulkEject,VolKick: End Sub

'Stand Up Targets
Sub Sw1_Hit : vpmTimer.PulseSw 1 : End Sub
Sub Sw2_Hit : vpmTimer.PulseSw 2 : End Sub
Sub Sw3_Hit : vpmTimer.PulseSw 3 : End Sub
Sub Sw4_Hit : vpmTimer.PulseSw 4 : End Sub

Sub Sw7_Hit : vpmTimer.PulseSw 7 : End Sub

Sub Sw13_Hit : vpmTimer.PulseSw 13 : End Sub
Sub Sw14_Hit : vpmTimer.PulseSw 14 : End Sub

Sub Sw34_Hit : vpmTimer.PulseSw 34 : End Sub
Sub Sw35_Hit : vpmTimer.PulseSw 35 : End Sub
Sub Sw36_Hit : vpmTimer.PulseSw 36 : End Sub

Sub Sw63_Hit : vpmTimer.PulseSw 63 : End Sub

'Drop Targets
Sub sw52_Dropped : dtBank.Hit 1:End Sub
Sub sw53_Dropped : dtBank.Hit 2:End Sub
Sub sw54_Dropped : dtBank.Hit 3:End Sub
Sub sw55_Dropped : dtBank.Hit 4:End Sub

'Wire Triggers
Sub sw10_Hit:Controller.Switch(10) = 1 : playsoundAtVol"rollover", sw10, VolRol : End Sub
Sub sw10_UnHit:Controller.Switch(10) = 0:End Sub
Sub sw11_Hit:Controller.Switch(11) = 1 : playsoundAtVol"rollover",sw11, VolRol : End Sub
Sub sw11_UnHit:Controller.Switch(11) = 0:End Sub
Sub sw12_Hit:Controller.Switch(12) = 1 : playsoundAtVol"rollover",sw12, VolRol : End Sub
Sub sw12_UnHit:Controller.Switch(12) = 0:End Sub
Sub sw23_Hit:Controller.Switch(23) = 1:clockwise=0:End sub 'shooter lane
Sub sw23_UnHit:Controller.Switch(23) = 0:End Sub
Sub sw24_Hit:Controller.Switch(24) = 1 : playsoundAtVol"rollover",sw24, VolRol : End Sub
Sub sw24_UnHit:Controller.Switch(24) = 0:End Sub
Sub sw25_Hit:Controller.Switch(25) = 1 : playsoundAtVol"rollover",sw25, VolRol : End Sub
Sub sw25_UnHit:Controller.Switch(25) = 0:End Sub
Sub sw28_Hit:Controller.Switch(28) = 1 : playsoundAtVol"rollover",sw28, VolRol : End Sub
Sub sw28_UnHit:Controller.Switch(28) = 0:End Sub
Sub sw29_Hit:Controller.Switch(29) = 1 : playsoundAtVol"rollover",sw29, VolRol : End Sub
Sub sw29_UnHit:Controller.Switch(29) = 0:End Sub
Sub sw33_Hit:Controller.Switch(33) = 1 : playsoundAtVol"rollover",sw33, VolRol : End Sub
Sub sw33_UnHit:Controller.Switch(33) = 0:End Sub
Sub sw43_Hit:Controller.Switch(43) = 1 : playsoundAtVol"rollover",sw43, VolRol : End Sub
Sub sw43_UnHit:Controller.Switch(43) = 0:End Sub
Sub sw47_Hit:Controller.Switch(47) = 1 : playsoundAtVol"rollover",sw47, VolRol : End Sub
Sub sw47_UnHit:Controller.Switch(47) = 0:End Sub
Sub sw61_Hit:Controller.Switch(61) = 1 : playsoundAtVol"rollover",sw61, VolRol : End Sub
Sub sw61_UnHit:Controller.Switch(61) = 0:End Sub

'Bumpers
Sub Bumper1b_Hit : vpmTimer.PulseSw(31) : playsoundAtVol SoundFX("fx_bumper1",DOFContactors),Bumper1b,VolBump: End Sub
Sub Bumper2b_Hit : vpmTimer.PulseSw(30) : playsoundAtVol SoundFX("fx_bumper1",DOFContactors),Bumper2b,VolBump: End Sub
Sub Bumper3b_Hit : vpmTimer.PulseSw(32) : playsoundAtVol SoundFX("fx_bumper1",DOFContactors),Bumper3b,VolBump: End Sub

'Spinners
Sub sw44_Spin:vpmTimer.PulseSw 44 : playsoundAtVol"fx_spinner",sw44,VolSpin : End Sub

'Right Ramp
Dim clockwise
Sub sw48_Hit
	if clockwise = 1 then ' only register from left loop
		Controller.Switch(48) = 1
		clockwise = 0
	end if
End Sub
Sub sw48_UnHit:Controller.Switch(48) = 0:End Sub
Sub Sw48a_Hit: clockwise = 1: end sub

'Right Ramp ball Lock
Sub sw49_hit():Controller.Switch(49) = false:post50.isdropped = false:debug.print "POST2 UP":end sub
Sub sw49_unhit():Controller.Switch(49) = true:post50.isdropped = true:debug.print "POST2 DOWN":end sub
Sub sw50_hit():Controller.Switch(50) = false:post51.isdropped = false:debug.print "POST3 UP":end sub
Sub sw50_unhit():Controller.Switch(50) = true:post51.isdropped = true:debug.print "POST3 DOWN": end sub
Sub sw51_hit():Controller.Switch(51) = false::end sub 'Loki Locks
Sub sw51_unhit():Controller.Switch(51) = true:end sub

'Hidden Hulk
Sub sw57_Hit:Controller.Switch(57) = 1 : playsoundAtVol"rollover",sw57, VolRol: End Sub
Sub sw57_UnHit:Controller.Switch(57) = 0:End Sub

'Generic Sounds
Sub dropleft_Hit:PlaySoundAtVol "fx_ballrampdrop",dropleft,1:End Sub
Sub dropleft2_Hit:PlaySoundAtVol "fx_ballrampdrop",dropleft2,1:End Sub
Sub dropright_Hit:PlaySoundAtVol "fx_ballrampdrop",dropright,1:End Sub
Sub metalramp_Hit:PlaySoundAtVol "rail",metalramp,1:End Sub
Sub MetalRamp_UnHit : StopSound "rail": End Sub


Const TesseractSw1 = 45
Const TesseractSw2 = 46

'**********************************************************************************************************
'Tesserac Animation
'**********************************************************************************************************

Dim lampPosition, lampSpinSpeed, lampLastPos
Const cLampSpeedMult = 180             ' 180 - Affects speed transfer to object (deg/sec)
Const cLampFriction = 2.0              ' 2.0 - Friction coefficient (deg/sec/sec)
Const cLampMinSpeed = 16              ' 20 - Object stops at this speed (deg/sec)
Const cLampRadius = 72
Const cBallSpeedDampeningEffect = 0.45 ' 45 - The ball retains this fraction of its speed due to energy absorption by hitting the lamp.

' Draw lamp
Sub SpinTimer_Timer
    Dim curPos
    Dim oldLampSpeed:oldLampSpeed = lampSpinSpeed
    lampPosition = lampPosition + lampSpinSpeed * Me.Interval / 1000
    lampSpinSpeed = lampSpinSpeed * (1 - cLampFriction * Me.Interval / 1000)
    Do While lampPosition < 0
        lampPosition = lampPosition + 360
    Loop
    Do While lampPosition > 360
        lampPosition = lampPosition - 360
    Loop
    curPos = Int((lampPosition * colLampPoles.Count) / 360)

	cube.ObjRotZ = -lampPosition - 12'-curPos*10
	cubeb.ObjRotZ = -lampPosition - 12  '-curPos*10
	CubePosts.ObjRotZ = -lampPosition - 12
	tessbase.ObjRotZ =  -lampPosition - 12 '-curPos*10

    If curPos <> lampLastPos Then

		'LampPr.RotZ = 360 -lampPosition 'Not applicable
        If lampLastPos >= 0 Then ' not first time
            colLampPoles(lampLastPos).IsDropped = True
            colLampPoles2(lampLastPos).IsDropped = True

'			If cController = 3 Then Controller.B2SSetData 101, 1
        End If
'		If cController = 3 Then Controller.B2SSetData 101, 0

        On Error Resume Next
        colLampPoles(curPos).IsDropped = False
        If Err Then msgbox curPoles
        colLampPoles2(curPos).IsDropped = False

        If oldLampSpeed > 0 And lampLastPos > curPos Then
            'rev anticlockwise
            vpmTimer.PulseSw TesseractSW1
        ElseIf oldLampSpeed < 0 And lampLastPos < curPos Then
            'rev clockwise
            vpmTimer.PulseSw TesseractSW2
        End If
        lampLastPos = curPos
    End If
    If Abs(lampSpinSpeed) < cLampMinSpeed Then
        lampSpinSpeed = 0:Me.Enabled = False
    End If
End Sub

Sub colLampPoles_Hit(idx)

 	'play rubber sound
	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 14 then
		PlaySoundAtVol "bump", ActiveBall, 1
	End if
	If finalspeed >= 4 AND finalspeed <= 14 then
 		RandomSoundRubber()
 	End If
	If finalspeed < 4 AND finalspeed > 1 then
' 		RandomSoundRubberLowVolume()
 	End If


    Dim pi:pi = 3.14159265358979323846
    dim lampangle
    lampangle = NormAngle((idx) / Me.Count * 2 * pi + (pi / 17.6) + pi) ' added (pi / 17.6) because the first lamp post (18) is not quite at 0 angle, added another pi to make it standard angle

    dim mball, mlamp, rlamp, ilamp

    dim collisionangle
    dim ballspeedin, ballspeedout, lampspeedin, lampspeedout

    dim fudge:fudge = cLampSpeedMult / 2

    With ActiveBall
        collisionangle = GetCollisionAngle(idx, Me.Count, .X, -.Y) ' this is the angle from the center of ball to the center of post

        Set ballspeedout = new jVector
        ballspeedout.SetXY .VelX, -.VelY
        ballspeedout.ShiftAxes - collisionangle

        lampSpinSpeed = lampSpinSpeed + sqr(.VelX ^2 + .VelY ^2) * sin(collisionangle - lampangle) * fudge
        ballspeedout.SetXY ballspeedout.x, ballspeedout.y * cos(collisionangle - lampangle)

        ballspeedout.ShiftAxes collisionangle
		' we can give a more accurate ball return speed or let the normal VP physics give the ball speed
        '.VelX = ballspeedout.x * cBallSpeedDampeningEffect
        '.VelY = - ballspeedout.y * cBallSpeedDampeningEffect
    End With
    SpinTimer.Enabled = True
End Sub

Function GetCollisionAngle(idx, count, X, Y)
    Dim pi:pi = 3.14159265358979323846
    dim angle, postx, posty, dX, dY
    Dim ang
    angle = (idx) / count * 2 * pi + (pi / 17.6) ' added (pi / 17.6) because the first lamp post (18) is not quite at 0 angle
    postx = 510.13 - 60.25 * Cos(angle)             ' 551 and 825.25 are the actual coordinates of the center of the lamp
    posty = 802.58 + 60.25 * Sin(angle)          ' 60.25 is the radius of the circle with center at the center of the lamp and edge at the centers of all the lamp posts
    posty = -1 * posty
    Dim collisionV:Set collisionV = new jVector
    collisionV.SetXY postx - X, posty - Y
    GetCollisionAngle = collisionV.ang
End Function

Function NormAngle(angle)
    NormAngle = angle
    Dim pi:pi = 3.14159265358979323846
    Do While NormAngle > 2 * pi
        NormAngle = NormAngle - 2 * pi
    Loop
    Do While NormAngle < 0
        NormAngle = NormAngle + 2 * pi
    Loop
End Function

Class jVector
    Private m_mag, m_ang, pi

    Sub Class_Initialize
        m_mag = CDbl(0)
        m_ang = CDbl(0)
        pi = CDbl(3.14159265358979323846)
    End Sub

    Public Function add(anothervector)
        Dim tx, ty, theta
        If TypeName(anothervector) = "jVector" then
            Set add = new jVector
            add.SetXY x + anothervector.x, y + anothervector.y
        End If
    End Function

    Public Function multiply(scalar)
        Set multiply = new jVector
        multiply.SetXY x * scalar, y * scalar
    End Function

    Sub ShiftAxes(theta)
        ang = ang - theta
    end Sub

    Sub SetXY(tx, ty)

        if tx = 0 And ty = 0 Then
            ang = 0
         elseif tx = 0 And ty < 0 then
            ang = - pi / 180 ' -90 degrees
         elseif tx = 0 And ty > 0 then
            ang = pi / 180   ' 90 degrees
        else
            ang = atn(ty / tx)
            if tx < 0 then ang = ang + pi ' Add 180 deg if in quadrant 2 or 3
        End if

        mag = sqr(tx ^2 + ty ^2)
    End Sub

    Property Let mag(nmag)
        m_mag = nmag
    End Property

    Property Get mag
        mag = m_mag
    End Property

    Property Let ang(nang)
        m_ang = nang
        Do While m_ang > 2 * pi
            m_ang = m_ang - 2 * pi
        Loop
        Do While m_ang < 0
            m_ang = m_ang + 2 * pi
        Loop
    End Property

    Property Get ang
        Do While m_ang > 2 * pi
            m_ang = m_ang - 2 * pi
        Loop
        Do While m_ang < 0
            m_ang = m_ang + 2 * pi
        Loop
        ang = m_ang
    End Property

    Property Get x
        x = m_mag * cos(ang)
    End Property

    Property Get y
        y = m_mag * sin(ang)
    End Property

    Property Get dump
        dump = "vector "
        Select Case CInt(ang + pi / 8)
            case 0, 8:dump = dump & "->"
            case 1:dump = dump & "/'"
            case 2:dump = dump & "/\"
            case 3:dump = dump & "'\"
            case 4:dump = dump & "<-"
            case 5:dump = dump & ":/"
            case 6:dump = dump & "\/"
            case 7:dump = dump & "\:"
        End Select

        dump = dump & " mag:" & CLng(mag * 10) / 10 & ", ang:" & CLng(ang * 180 / pi) & ", x:" & CLng(x * 10) / 10 & ", y:" & CLng(y * 10) / 10
    End Property
End Class

'**********************************************************************************************************
'**********************************************************************************************************


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

Sub UpdateLamps()
NFadeL 3, l3
NFadeL 4, l4
NFadeL 5, l5
NFadeL 6, l6
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
NFadeLm 20, l20
NFadeL 20, l20x
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
'NadeL 40, l40
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
'NFadeL 56, l56
NFadeL 57, l57
'NFadeL 58, l58
'NFadeL 59, l59
NFadeL 60, l60 'bumpers
NFadeL 61, l61
NFadeL 62, l62
NFadeL 63, l63
NFadeL 65, l65
NFadeL 66, l66
NFadeL 67, l67
NFadeL 68, l68
NFadeL 69, l69
NFadeL 70, l70
NFadeL 78, l78
NFadeObjm 71, l71, "bulbcover1_greenOn", "bulbcover1_green"
Flash 71, f71
NFadeObjm 73, l73, "bulbcover1_greenOn", "bulbcover1_green"
Flash 73, f73
NFadeObjm 75, l75, "bulbcover1_greenOn", "bulbcover1_green"
Flash 75, f75


'Solenoid Controlled Lamps and Flahsers

NFadeObjm 118, Flasher18p, "dome2_0_redON", "dome2_0_red"
' NFadeL 118, Flasher18a
LampMod 118, Flasher18a

'NFadeObjm 119, Flasher19p, "DomeOn", "DomeOff"
'NFadeL 119, Flasher19a

NFadeObjm 120, Flasher20p1, "dome2_0_redON", "dome2_0_red"
NFadeObjm 120, Flasher20p2, "dome2_0_redON", "dome2_0_red"
' NFadeLm 120, Flasher20a
' NFadeL 120, Flasher20b
LampMod 120, Flasher20a
LampMod 120, Flasher20b

' NFadeLm 121, Spot21a
' NFadeL 121, Spot21b
LampMod 121, Spot21a
LampMod 121, Spot21b

' NFadeL 125, Flasher25
LampMod 125, Flasher25

' NFadeLm 126, Flasher26a
' NFadeLm 126, Flasher26b
' NFadeL 126, Flasher26c
LampMod 126, Flasher26a
LampMod 126, Flasher26b
LampMod 126, Flasher26c

NFadeObjm 127, Flasher27p, "dome2_0_blueOn", "dome2_0_blue"
' Flash 127, F27
LampMod 127, F27

NFadeObjm 128, Flasher28p, "dome2_0_purpleOn", "dome2_0_purple"
' Flash 128, F28
LampMod 128, F28

NFadeObjm 129, Flasher29p, "dome2_0_greenOn", "dome2_0_green"
' Flash 129, F29
LampMod 129, F29

NFadeObjm 130, Flasher30p, "dome2_0_redON", "dome2_0_red"
' Flash 130, F30
LampMod 130, F30

NFadeObjm 131, Flasher31p, "dome2_0_clearON", "dome2_0_clear"
' Flash 131, F31
LampMod 131, F31

NFadeObjm 132, Flasher32p, "dome2_0_yellowON", "dome2_0_yellow"
' Flash 132, F32
LampMod 132, F32

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
    vpmTimer.PulseSw 27
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
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
	vpmTimer.PulseSw 26
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
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
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
	PlaySound "pinhit_low", 0, Vol(ActiveBall)*VolPi, Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
	PlaySound "target", 0, Vol(ActiveBall)*VolTarg, Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
	PlaySound "metalhit_thin", 0, Vol(ActiveBall)*VolMetal, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
	PlaySound "metalhit_medium", 0, Vol(ActiveBall)*VolMetal, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
	PlaySound "metalhit2", 0, Vol(ActiveBall)*VolMetal, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
	PlaySound "gate4", 0, Vol(ActiveBall)*VolGates, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinner_Spin
	PlaySoundAtVol "fx_spinner", Spinner, VolSpin
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
		Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall)*VolFlip, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall)*VolFlip, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall)*VolFlip, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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

Sub SetLampMod(nr, value)
      If value > 0 Then
      LampState(nr) = 1
    Else
      LampState(nr) = 0
    End If
    FadingLevel(nr) = value
End Sub

Sub LampMod(nr, object)
    If TypeName(object) = "Light" Then
      Object.IntensityScale = FadingLevel(nr)/128
      Object.State = LampState(nr)
    End If
    If TypeName(object) = "Flasher" Then
      Object.IntensityScale = FadingLevel(nr)/128
      Object.visible = LampState(nr)
    End If
    If TypeName(object) = "Primitive" Then
      Object.DisableLighting = LampState(nr)
    End If
End Sub


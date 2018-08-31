Option Explicit
Randomize

' Thalamus 2018-07-23
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' No special SSF tweaks yet.


On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="goinnuts",UseSolenoids=2,UseLamps=1,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff",SFlipperOn="FlipperUp",SFlipperOff="FlipperDown"
Const SCoin="coin3",cCredits=""

LoadVPM "01120100","sys80.vbs",3.02

BMass=1.65

Dim DesktopMode: DesktopMode = GN.ShowDT

if DesktopMode=False Then
	leftrail.visible=0
	rightrail.visible=0
end If

Sub GN_KeyDown(ByVal keycode)
'	if  keycode=LeftFlipperKey then
'		LeftFlipper.RotateToEnd
'		LeftFlipper1.RotateToEnd
'	end if
'	if  keycode=RightFlipperKey then
'		RightFlipper.RotateToEnd
'		RightFlipper1.RotateToEnd
'	end if

	If vpmKeyDown(KeyCode) Then Exit Sub
End Sub

Sub GN_KeyUp(ByVal keycode)
'	if  keycode=LeftFlipperKey then
'		LeftFlipper.RotateToStart
'		LeftFlipper1.RotateToStart
'	end if
'	if  keycode=RightFlipperKey then
'		RightFlipper.RotateToStart
'		RightFlipper1.RotateToStart
'	end if
	If vpmKeyUp(KeyCode) Then Exit Sub
End Sub

Dim dtGreen,dtRed,dtYellow,dtBlue,dtWhite,bsTrough,cbCaptive

Sub GN_Init
	vpmInit Me
	On Error Resume Next
	With Controller
		.GameName=cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
		.SplashInfoLine="Goin' Nuts - Gottlieb, 1983" & vbNewLine & "Table Design: fuzzel"
		.HandleMechanics=0
		.HandleKeyboard=0
		.ShowDMDOnly=1
		.ShowFrame=0
		.ShowTitle=0
		.SolMask(0)=0
		.Hidden=0
		If Err Then MsgBox Err.Description
	End With
	On Error Goto 0
		vpmTimer.AddTimer 1000,"Controller.SolMask(0)=&Hffffffff'"
		Controller.Run

	'Dip switch settings, shown this way for clarity and ease of modification
'	Controller.Dip(0) = (0*1 + 1*2 + 0*4 + 0*8 + 0*16 + 0*32 + 1*64 + 0*128) '01-08
'	Controller.Dip(1) = (0*1 + 0*2 + 0*4 + 0*8 + 0*16 + 0*32 + 1*64 + 1*128) '09-16
'	Controller.Dip(2) = (0*1 + 1*2 + 0*4 + 1*8 + 0*16 + 0*32 + 1*64 + 1*128) '17-24
'	Controller.Dip(3) = (0*1 + 1*2 + 1*4 + 1*8 + 1*16 + 1*32 + 0*64 + 0*128) '25-32

	' Main Timer init
	PinMAMETimer.Interval=PinMAMEInterval
	PinMAMETimer.Enabled=1

	' Nudging
	vpmNudge.TiltSwitch=57
	vpmNudge.Sensitivity=5
	vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3,Bumper4,Bumper5,Bumper6,Bumper7,LeftSlingshot,RightSlingshot)

	'drop targets
	Set dtGreen=New cvpmDropTarget
	dtGreen.InitDrop Array(SW00,SW10,SW20),Array(0,10,20)
	dtGreen.initsnd SoundFX("flapclos",DOFDropTargets),SoundFX("flapopen",DOFContactors)

	Set dtRed=New cvpmDropTarget
	dtRed.InitDrop Array(SW01,SW11,SW21),Array(1,11,21)
	dtRed.initsnd SoundFX("flapclos",DOFDropTargets),SoundFX("flapopen",DOFContactors)

	Set dtBlue=New cvpmDropTarget
	dtBlue.InitDrop Array(SW02,SW12,SW22),Array(2,12,22)
	dtBlue.initsnd SoundFX("flapclos",DOFDropTargets),SoundFX("flapopen",DOFContactors)

	Set dtYellow=New cvpmDropTarget
	dtYellow.InitDrop Array(SW03,SW13,SW23),Array(3,13,23)
	dtYellow.initsnd SoundFX("flapclos",DOFDropTargets),SoundFX("flapopen",DOFContactors)

	Set dtWhite=New cvpmDropTarget
	dtWhite.InitDrop Array(SW04,SW14,SW24),Array(4,14,24)
	dtWhite.initsnd SoundFX("flapclos",DOFDropTargets),SoundFX("flapopen",DOFContactors)

' Trough handler
	Set bsTrough=new cvpmBallStack
	bsTrough.InitSw 0,54,0,0,0,0,0,0
	bsTrough.InitKick sw54,50,6
	bsTrough.InitExitSnd SoundFX("ballrel",DOFContactors),SoundFX("solon",DOFContactors)
	bsTrough.Balls=3

'    Set cbCaptive = New cvpmCaptiveBall
'    With cbCaptive
'        .InitCaptive CapTrigger, CapWall, Array(CapKicker1,CapKicker2), 320
'        .NailedBalls = 1
'        .ForceTrans = .9
'        .MinForce = 3.5
'        .CreateEvents "cbCaptive"
'        .Start
'    End With

	CapKicker.createball
	CapKicker.kick 120,6
End Sub

Sub GN_exit()
	Controller.Pause = False
	Controller.Stop
End Sub

Set LampCallback=GetRef("UpdateMultipleLamps")

dim CZ,NewCZ,DTZ,NewDTZ,BZ,NewBZ
CZ=0
DTZ=0
BZ=0

Sub UpdateMultipleLamps
	NewCZ=L12.State
	If NewCZ<>CZ Then
			If L12.State=1 And bsTrough.Balls>0 Then bsTrough.ExitSol_On
	End If
	CZ=NewCZ
	NewDTZ=L13.State
	If NewDTZ<>DTZ Then
			If L13.State=1 Then dtWhite.DropSol_On
	End If
	DTZ=NewDTZ
	NewBZ=L14.State
	If NewBZ<>BZ Then
			Controller.Switch(53)=0
			Kicker2.Kick 0,35
	End If
	BZ=NewBZ
End Sub

'constants
Const sdtGreen=1'ok
Const sdtRed=2'ok
Const sdtBlue=5'ok
Const sdtYellow=6'ok
Const sCoin1=3'ok
Const sCoin2=4'ok
Const sCoin3=7'ok
Const sknocker=8
Const sOutHole=9
Const swCaptiveBack=64

SolCallback(sdtGreen)="dtGreen.SolDropUp"
SolCallback(sdtRed)="dtRed.SolDropUp"
SolCallback(sdtYellow)="dtYellow.SolDropUp"
SolCallback(13)="dtWhite.SolDropUp"
SolCallback(sdtBlue)="dtBlue.SolDropUp"
SolCallback(sKnocker)="VpmSolSound SoundFX(""knocker"",DOFKnocker),"
SolCallback(sOutHole)="HandleTrough"

'SolCallback(sLLFlipper)="VpmSolFlipper LeftFlipper,nothing,"
'SolCallback(sLRFlipper)="SolRightF"

'Sub SolRightF(Enabled)
'	If Enabled Then
'			LeftFlipper1.RotateToEnd
'			VpmSolFlipper RightFlipper,RightFlipper1,True
'		Else
'			LeftFlipper1.RotateToStart
'			VpmSolFlipper RightFlipper,RightFlipper1,False
'	End If
'End Sub

' Flipper Subs
  SolCallback(sLRFlipper) = "SolRFlipper"
  SolCallback(sLLFlipper) = "SolLFlipper"

  Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFx("FlipperDown",DOFFlippers):LeftFlipper.RotateToEnd
     Else
         PlaySound SoundFx("FlipperUp",DOFFlippers):LeftFlipper.RotateToStart
     End If
  End Sub

  Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFx("FlipperDown",DOFFlippers):RightFlipper.RotateToEnd:RightFlipper1.RotateToEnd:LeftFlipper1.RotateToEnd
     Else
         PlaySound SoundFx("FlipperUp",DOFFlippers):RightFlipper.RotateToStart:RightFlipper1.RotateToStart:LeftFlipper1.RotateToStart
     End If
  End Sub

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, -0.1, 0.15
End Sub

Sub Rightflipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, 0.1, 0.15
End Sub

Sub rubbersCol_Hit(idx):PlaySound "fx_rubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub

'drop targets
Sub sw00_Hit:dtGreen.Hit 1:End Sub
Sub sw10_Hit:dtGreen.Hit 2:End Sub
Sub sw20_Hit:dtGreen.Hit 3:End Sub

Sub sw03_Hit:dtYellow.Hit 1:End Sub
Sub sw13_Hit:dtYellow.Hit 2:End Sub
Sub sw23_Hit:dtYellow.Hit 3:End Sub

Sub sw02_Hit:dtBlue.Hit 1:End Sub
Sub sw12_Hit:dtBlue.Hit 2:End Sub
Sub sw22_Hit:dtBlue.Hit 3:End Sub

Sub sw01_Hit:dtRed.Hit 1:End Sub
Sub sw11_Hit:dtRed.Hit 2:End Sub
Sub sw21_Hit:dtRed.Hit 3:End Sub

Sub sw04_Hit:dtWhite.Hit 1:End Sub
Sub sw14_Hit:dtWhite.Hit 2:End Sub
Sub sw24_Hit:dtWhite.Hit 3:End Sub

Sub Kicker2_Hit:Controller.Switch(53)=1:End Sub

'spot targets
Sub sw25a_Hit:VpmTimer.PulseSw 25:DOF 113, DOFPulse:End Sub ' #1 spot target
Sub sw25b_Hit:VpmTimer.PulseSw 25:DOF 113, DOFPulse:End Sub ' #3 spot target
Sub sw15a_Hit:VpmTimer.PulseSw 15:DOF 114, DOFPulse:End Sub ' lower rightspot target
Sub sw15b_Hit:VpmTimer.PulseSw 15:DOF 113, DOFPulse:End Sub ' #2 spot target
Sub sw06_Hit:VpmTimer.PulseSw 6:End Sub ' captive ball spot target

'rollovers/rollunders
Sub sw26_Hit:Controller.Switch(26)=1:End Sub
Sub sw26_UnHit:Controller.Switch(26)=0:End Sub
'Sub sw50_Hit:Controller.Switch(50)=1:End Sub
'Sub sw50_UnHit:Controller.Switch(50)=0:End Sub
Sub sw50_Hit:VpmTimer.PulseSw 50:End Sub
'Sub sw51_Hit:Controller.Switch(51)=1:End Sub
'Sub sw51_unHit:Controller.Switch(51)=0:End Sub
Sub sw51_Hit:VpmTimer.PulseSw 51:End Sub

'Bumpers/Slingshots
Sub Bumper1_Hit:VpmTimer.PulseSw 52:PlaySound SoundFX("jet1",DOFContactors):DOF 101, DOFPulse:End Sub
Sub Bumper2_Hit:VpmTimer.PulseSw 52:PlaySound SoundFX("jet1",DOFContactors):DOF 102, DOFPulse:End Sub
Sub Bumper3_Hit:VpmTimer.PulseSw 52:PlaySound SoundFX("jet1",DOFContactors):DOF 103, DOFPulse:End Sub
Sub Bumper4_Hit:VpmTimer.PulseSw 52:PlaySound SoundFX("jet1",DOFContactors):DOF 104, DOFPulse:End Sub
Sub Bumper5_Hit:VpmTimer.PulseSw 52:PlaySound SoundFX("jet1",DOFContactors):DOF 105, DOFPulse:End Sub
Sub Bumper6_Hit:VpmTimer.PulseSw 52:PlaySound SoundFX("jet1",DOFContactors):DOF 106, DOFPulse:End Sub
Sub Bumper7_Hit:VpmTimer.PulseSw 52:PlaySound SoundFX("jet1",DOFContactors):DOF 107, DOFPulse:End Sub

dim rSlingStep:rSlingStep=0
Sub RightSlingshot_Timer
	select case rSlingStep
		case 1: rsling1.visible=0:rsling2.visible=1
		case 2: rsling2.visible=0:rsling3.visible=1
		case 3: rsling3.visible=0:rsling0.visible=1:RightSlingshot.TimerEnabled=0
	end select
	rSlingStep = rSlingStep + 1
end sub

dim lSlingStep:lSlingStep=0
Sub LeftSlingShot_Timer
	select case lSlingStep
		case 1: lsling1.visible=0:lsling2.visible=1
		case 2: lsling2.visible=0:lsling3.visible=1
		case 3: lsling3.visible=0:lsling0.visible=1:LeftSlingshot.TimerEnabled=0
	end select
	lSlingStep = lSlingStep + 1
end sub


Sub LeftSlingshot_Slingshot:PlaySound SoundFX("rsling",DOFContactors):VpmTimer.PulseSw 16:DOF 108, 2:lSlingStep=0:LeftSlingshot.TimerEnabled=1:End Sub
Sub RightSlingshot_Slingshot:PlaySound SoundFX("rsling",DOFContactors):VpmTimer.PulseSw 16:DOF 109, 2:rSlingStep=0:RightSlingshot.TimerEnabled=1:End Sub
Sub upsling_Slingshot:PlaySound SoundFX("rsling",DOFContactors):VpmTimer.PulseSw 16:DOF 110, 2:End Sub
Sub msling_Slingshot:PlaySound SoundFX("rsling",DOFContactors):VpmTimer.PulseSw 16:DOF 111, 2:End Sub
Sub mrslingshot_Slingshot:PlaySound SoundFX("rsling",DOFContactors):VpmTimer.PulseSw 16:DOF 112, 2:End Sub
Sub mlslingshot_Slingshot:PlaySound SoundFX("rsling",DOFContactors):VpmTimer.PulseSw 16:DOF 112, 2:End Sub

sub UpdateFlippers
	pRightFlipper.rotZ = RightFlipper.currentAngle
	pLeftFlipper.rotZ = LeftFlipper.currentAngle
	pRightFlipper1.rotZ = RightFlipper1.currentAngle
	pLeftFlipper1.rotZ = LeftFlipper1.currentAngle
end Sub

Set Lights(3)=L3
Set Lights(12)=L12
Set Lights(13)=L13
Set Lights(14)=L14
Set Lights(15)=L15
Set Lights(16)=L16
Set Lights(17)=L17
Set Lights(18)=L18
Set Lights(19)=L19
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
Set Lights(32)=L32
Set Lights(33)=L33
Set Lights(34)=L34
Set Lights(35)=L35
Set Lights(36)=L36
Set Lights(37)=L37
Set Lights(38)=L38
Set Lights(39)=L39
Set Lights(40)=L40
Set Lights(41)=L41
Set Lights(42)=L42

Sub LightTimer_Timer
	if L15.State=1 then L50.State=1 else L50.State=0 end if
	if L19.State=1 then L51.State=1:L53.State=1 else L51.State=0:L53.State=0 end if
	if L40.State=1 then L52.State=1 else L52.State=0 end if
	if L42.State=1 then L54.State=1 else L54.State=0 end if
	if L24.State=1 then L55.State=1 else L55.State=0 end if
	if L30.State=1 then L56.State=1 else L56.State=0 end if
	UpdateFlippers
End Sub

Sub sw67_Hit:Controller.Switch(55)=1:End Sub
Sub HandleTrough(Enabled)
	If Enabled Then
	Controller.Switch(55)=0
	bsTrough.AddBall 0
	sw67.Destroyball
	End If
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

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "GN" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / GN.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "GN" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / GN.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "GN" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / GN.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioFade(ball) ' Can this be together with the above function ?
  Dim tmp
  tmp = ball.y * 2 / GN.height-1
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

Const tnob = 7 ' total number of balls in this table is 4, but always use a higher number here bacuse of the timing
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
  If GN.VersionMinor > 3 OR GN.VersionMajor > 10 Then
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
  Else
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0
  End if
End Sub

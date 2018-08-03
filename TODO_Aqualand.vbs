'
' Aqualand / IPD No. 3935 / 1986 / 4 Players
' Manufacturer: Juegos Populares, S.A., of Madrid, Spain (1986)
' http://ipdb.org/machine.cgi?id=3935
'
' VP8 version by Destruk
' VP10 version by MaX in July 2015
'
' \_  _/  /| |\/
'  \\//  / | |/\
'         /
'
' Thanks to akiles5000 for the pictures of the playfield, plastics, apron, the touched up backglass and everything else
'

Option Explicit
Randomize

'******************* Options *********************
' DMD/Backglass Controller Setting
Const cController = 0		'0=Use value defined in cController.txt, 1=VPinMAME, 2=UVP server, 3=B2S server, 4=B2S with DOF (disable VP mech sounds)
'*************************************************

Dim cNewController
Sub LoadVPM(VPMver, VBSfile, VBSver)
	Dim FileObj, ControllerFile, TextStr

	On Error Resume Next
	If ScriptEngineMajorVersion < 5 Then MsgBox "VB Script Engine 5.0 or higher required"
	ExecuteGlobal GetTextFile(VBSfile)
	If Err Then MsgBox "Unable to open " & VBSfile & ". Ensure that it is in the same folder as this table. " & vbNewLine & Err.Description

	cNewController = 1
	If cController = 0 then
		Set FileObj=CreateObject("Scripting.FileSystemObject")
		If Not FileObj.FolderExists(UserDirectory) then
			Msgbox "Visual Pinball\User directory does not exist. Defaulting to vPinMame"
		ElseIf Not FileObj.FileExists(UserDirectory & "cController.txt") then
			Set ControllerFile=FileObj.CreateTextFile(UserDirectory & "cController.txt",True)
			ControllerFile.WriteLine 1: ControllerFile.Close
		Else
			Set ControllerFile=FileObj.GetFile(UserDirectory & "cController.txt")
			Set TextStr=ControllerFile.OpenAsTextStream(1,0)
			If (TextStr.AtEndOfStream=True) then
				Set ControllerFile=FileObj.CreateTextFile(UserDirectory & "cController.txt",True)
				ControllerFile.WriteLine 1: ControllerFile.Close
			Else
				cNewController=Textstr.ReadLine: TextStr.Close
			End If
		End If
	Else
		cNewController = cController
	End If

	Select Case cNewController
		Case 1
			Set Controller = CreateObject("VPinMAME.Controller")
			If Err Then MsgBox "Can't Load VPinMAME." & vbNewLine & Err.Description
			If VPMver>"" Then If Controller.Version < VPMver Or Err Then MsgBox "VPinMAME ver " & VPMver & " required."
			If VPinMAMEDriverVer < VBSver Or Err Then MsgBox VBSFile & " ver " & VBSver & " or higher required."
		Case 2
			Set Controller = CreateObject("UltraVP.BackglassServ")
		Case 3,4
			Set Controller = CreateObject("B2S.Server")
	End Select
	On Error Goto 0
End Sub

'*************************************************************
'Toggle DOF sounds on/off based on cController value
'*************************************************************
Dim ToggleMechSounds
Function SoundFX (sound)
    If cNewController= 4 and ToggleMechSounds = 0 Then
        SoundFX = ""
    Else
        SoundFX = sound
    End If
End Function

Sub DOF(dofevent, dofstate)
	If cNewController>2 Then
		If dofstate = 2 Then
			Controller.B2SSetData dofevent, 1:Controller.B2SSetData dofevent, 0
		Else
			Controller.B2SSetData dofevent, dofstate
		End If
	End If
End Sub

LoadVPM "01520000", "juegos.vbs", 3.1

Const UseSolenoids = 2
Const cSingleLFlip = 0
Const cSingleRFlip = 0
Const UseLamps = 1
Const UseGI = 0
Const UseSync = 1 'set it to 1 if the table runs too fast
Const HandleMech = 0

' Standard Sounds
Const SSolenoidOn = "fx_Solenoid"
Const SSolenoidOff = "fx_SolenoidOff"
Const SFlipperOn = "fx_FlipperUp"
Const SFlipperOff = "fx_FlipperDown"
Const SCoin = "fx_coin"

Dim xx
If Aqualand.ShowDT = false then
		for each xx in aBackdrop: xx.visible = 0:next
	else
		for each xx in aBackdrop: xx.visible = 1:next
End If

SolCallback(8)="Flasher8.state="
SolCallback(9)="vpmSolSound SoundFX(""left_slingshot""),"
SolCallback(10)="vpmSolSound SoundFX(""right_slingshot""),"
SolCallback(11)="vpmSolSound SoundFX(""fx_bumper3""),"
SolCallback(12)="vpmSolSound SoundFX(""fx_bumper3""),"
SolCallback(13)="SolDropTargets"
SolCallback(14)="bsSaucer.SolOut"
SolCallback(15)="vpmSolSound SoundFX(""knocker""),"
SolCallback(16)="bsTrough.SolOut"
'SolCallback(sLLFlipper)="vpmSolFlipper LeftFlipper,Nothing,"
'SolCallback(sLRFlipper)="vpmSolFlipper RightFlipper,Nothing,"

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"


Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFx("FlipperUp"):LeftFlipper.RotateToEnd
     Else
         PlaySound SoundFx("FlipperDown"):LeftFlipper.RotateToStart
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFx("FlipperUp"):RightFlipper.RotateToEnd
     Else
         PlaySound SoundFx("FlipperDown"):RightFlipper.RotateToStart
     End If
End Sub



Dim bsTrough,dtDrop,bsSaucer

'************
' Table init.
'************

Sub Aqualand_Init

	Dim cGameName
	vpmInit Me
	With Controller
		cGameName = "aqualand"
		.GameName = cGameName
		.SplashInfoLine = ""
		.HandleMechanics = 0
		.HandleKeyboard = 0
		.ShowDMDOnly = 1
		.ShowFrame = 0
		.ShowTitle = 0
		.Hidden = 1
'        .Games(cGameName).Settings.Value("rol") = 0
'        .Games(cGameName).Settings.Value("sound") = 1
		If Err Then MsgBox Err.Description
	End With
	On Error Goto 0

	Controller.Run

	' Nudging
	vpmNudge.TiltSwitch=31:vpmNudge.Sensitivity=5:vpmNudge.TiltObj=Array(S13,S14,S15,S16)

	' Main Timer init
	PinMAMETimer.Interval = PinMAMEInterval
	PinMAMETimer.Enabled = 1

' dip switches

 'Coins 1 (4 on keyboard) = dip 0 switch 1/dip 0 switch 3
 'OFF OFF = 1:1   OFF ON =1:2
 'ON ON   = 2:1   ON OFF =1:3
 'Coins 2 (5 on keyboard) = dip 0 switch 5/dip 0 switch 7
 'OFF OFF = 1:1   OFF ON =1:2
 'ON ON   = 1:10  ON OFF =1:3

 'Extra Ball Award = dip 0 switch 4
 'Replay Level A = ON = 600.000 EB
 'Replay Level B = ON = 700.000 EB
 'Replay Level C = ON = 900.000 EB
 'Replay Level D = ON = 800.000 EB

 'Replay Levels = dip 1 switch 1/dip 1 switch 2
 'Replay Level A = OFF OFF=800.000/1.000.000/1.200.000
'Replay Level B = ON OFF=900.000/1.000.000/1.200.000
 'Replay Level C = ON ON =1.000.000/1.500.000/2.000.000
 'Replay Level D = OFF ON=1.000.000/1.000.000/1.400.000

 'BALLS = Controller.Dip 1 Switch 3 - off=3/on=5
'Ignore first replay level if under 1.000.000= dip 0 switch 6
 'ON=Ignore, OFF=Award Free Game
 'Only allow 1 score-based replay per game - to disable score based replays, enable 'ignore first replay level' feature
 'dip 0 switch 7
 'MATCH = Controller.Dip 1 Switch 8 - off=disabled/on=enabled

  'Coin Audits = Controller.Dip 2 Switch 3 - off=disabled/on=enabled
'Display Credit audit at end of each game = dip 2 switch 4
 'ON=Enabled, OFF=Disabled

'                    *           *     *     *      *      *      *
'Controller.Dip(0)=(0*1 + 0*2 + 0*4 + 0*8 + 0*16 + 0*32 + 0*64 + 0*128) '01-08
'                    *     *     *                                *
'Controller.Dip(1)=(0*1 + 0*2 + 0*4 + 0*8 + 0*16 + 0*32 + 0*64 + 0*128) '09-16
'                                *     *
'Controller.Dip(2)=(0*1 + 0*2 + 0*4 + 0*8 + 0*16 + 0*32 + 0*64 + 0*128) 'A-D


	Set bsTrough=New cvpmBallStack
	bsTrough.InitSw 0,25,0,0,0,0,0,0
	bsTrough.InitKick BallRelease,90,3
	bsTrough.InitExitSnd SoundFX("fx_BallRel"),SoundFX("fx_SolenoidOn")
	bsTrough.Balls=1

 	Set dtDrop=New cvpmDropTarget
	dtDrop.InitDrop Array(S8,S7,S6),Array(8,7,6)
	dtDrop.InitSnd SoundFX("fx_droptarget"),SoundFX("fx_resetdrop")

	Set bsSaucer=New cvpmBallStack
	bsSaucer.InitSaucer S12,12,283,15
	bsSaucer.InitExitSnd SoundFX("Popper_ball"),SoundFX("fx_SolenoidOn")

 	vpmMapLights Goodlights
 	vpmMapLights aLights

	VariTarget_Init

End Sub

Sub Aqualand_Paused:Controller.Pause = 1:End Sub
Sub Aqualand_unPaused:Controller.Pause = 0:End Sub
Sub Aqualand_exit:Controller.Pause = 0:Controller.stop:End Sub

'**********
' Keys
'**********

Sub Aqualand_KeyDown(ByVal Keycode)
'  	If KeyCode=LeftFlipperKey Then Controller.Switch(84)=1
'	If KeyCode=RightFlipperKey Then Controller.Switch(82)=1
	If KeyCode=PlungerKey Then Plunger.Pullback:playsound"plungerpull"
'	If keycode = LeftTiltKey Then LeftNudge 80, 2, 20:playsound SoundFX("nudge_left")
'	If keycode = RightTiltKey Then RightNudge 280, 2, 20:playsound SoundFX( "nudge_right")
'	If keycode = CenterTiltKey Then CenterNudge 0, 2, 25:playsound SoundFX("nudge_forward")
	If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub Aqualand_KeyUp(ByVal Keycode)
'	If KeyCode=LeftFlipperKey Then Controller.Switch(84)=0
'	If KeyCode=RightFlipperKey Then Controller.Switch(82)=0
	If KeyCode=PlungerKey Then PlaySound"plunger":Plunger.Fire
	If vpmKeyUp(keycode) Then Exit Sub
End Sub

'*****GI Lights On

For each xx in aGI:xx.State = 1: Next

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub S13_Slingshot
    PlaySound SoundFX("right_slingshot"), 0, 1, 0.05, 0.05
	vpmTimer.PulseSw 13
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    S13.TimerEnabled = 1
End Sub

Sub S13_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:S13.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub S14_Slingshot
    PlaySound SoundFX("left_slingshot"),0,1,-0.05,0.05
	vpmTimer.PulseSw 14
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    S14.TimerEnabled = 1
End Sub

Sub S14_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:S14.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function Vol2(ball1, ball2) ' Calculates the Volume of the sound based on the speed of two balls
    Vol2 = (Vol(ball1) + Vol(ball2) ) / 2
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "Aqualand" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / Aqualand.width-1
    If tmp> 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

'*****************************************
'    JP's VP10 Collision & Rolling Sounds
'*****************************************

Const tnob = 8 ' total number of balls
ReDim rolling(tnob)
ReDim collision(tnob)
Initcollision

Sub Initcollision
    Dim i
    For i = 0 to tnob
        collision(i) = -1
        rolling(i) = False
    Next
End Sub

Sub CollisionTimer_Timer()

    diverterP.RotZ = Flipper1.CurrentAngle

    Dim BOT, B, B1, B2, dx, dy, dz, distance, radii
    BOT = GetBalls

    ' rolling

	For B = UBound(BOT) +1 to tnob
        rolling(b) = False
        StopSound("fx_ballrolling" & b)
	Next

    If UBound(BOT) = -1 Then Exit Sub

    For B = 0 to UBound(BOT)
        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0
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
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0
End Sub

' sounds

Sub Pins_Hit (idx)
	PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
End Sub

Sub Targets_Hit (idx)
	PlaySound SoundFX("target"), 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
End Sub

Sub Metals_Thin_Hit (idx)
	PlaySound "metalhit_thin", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub Metals_Medium_Hit (idx)
	PlaySound "metalhit_medium", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub Metals2_Hit (idx)
	PlaySound "metalhit2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub Gates_Hit (idx)
	PlaySound "gate4", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub Spinner_Spin
	PlaySound "fx_spinner",0,.25,0,0.25
End Sub

Sub Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub Posts_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 16 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End if
	If finalspeed >= 6 AND finalspeed <= 16 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
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
		Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End Select
End Sub


'****Target
'Sub S11_Hit:vpmTimer.PulseSw 11:End Sub						'11
Sub S11_Hit
	vpmTimer.PulseSw 11
	s11p.transx = -10
	Me.TimerEnabled = 1
End Sub

Sub s11_Timer
	s11p.transx = 0
	Me.TimerEnabled = 0
End Sub

Sub S1_Hit:Controller.Switch(1)=1:End Sub					'1
Sub S1_unHit:Controller.Switch(1)=0:End Sub
Sub S2_Hit:Controller.Switch(2)=1:End Sub					'2
Sub S2_unHit:Controller.Switch(2)=0:End Sub
Sub S3_Hit:Controller.Switch(3)=1:End Sub					'3
Sub S3_unHit:Controller.Switch(3)=0:End Sub
Sub S4_Hit:Controller.Switch(4)=1:End Sub					'4
Sub S4_unHit:Controller.Switch(4)=0:End Sub
Sub S5_Hit:Controller.Switch(5)=1:End Sub					'5
Sub S5_unHit:Controller.Switch(5)=0:End Sub

' animated droptargets
Dim droptarget1pos,droptarget2pos,droptarget3pos
droptarget1pos=0:droptarget2pos=0:droptarget3pos=0

Sub S6_Hit:dtDrop.Hit 3:me.timerenabled=1:End Sub								'6
Sub S7_Hit:dtDrop.Hit 2:me.timerenabled=1:End Sub								'7
Sub S8_Hit:dtDrop.Hit 1:me.timerenabled=1:End Sub								'8

Sub SolDropTargets(enabled)
	if enabled then
		dtDrop.DropSol_On
		if droptarget1pos=48 then ldt1.enabled=1
		if droptarget2pos=48 then ldt2.enabled=1
		if droptarget3pos=48 then ldt3.enabled=1
	end if
End Sub

Sub S6_timer()
	droptarget3pos=droptarget3pos+4
	S6p.rotandtra5=0-droptarget3pos
	if droptarget3pos=48 then me.timerenabled=0
End Sub
Sub ldt3_timer()
	droptarget3pos=droptarget3pos-12
	s6p.rotandtra5=0-droptarget3pos
	if droptarget3pos=0 then me.enabled=0
End Sub
Sub S7_timer()
	droptarget2pos=droptarget2pos+4
	S7p.rotandtra5=0-droptarget2pos
	if droptarget2pos=48 then me.timerenabled=0
End Sub
Sub ldt2_timer()
	droptarget2pos=droptarget2pos-12
	s7p.rotandtra5=0-droptarget2pos
	if droptarget2pos=0 then me.enabled=0
End Sub
Sub S8_timer()
	droptarget1pos=droptarget1pos+4
	S8p.rotandtra5=0-droptarget1pos
	if droptarget1pos=48 then me.timerenabled=0
End Sub
Sub ldt1_timer()
	droptarget1pos=droptarget1pos-12
	s8p.rotandtra5=0-droptarget1pos
	if droptarget1pos=0 then me.enabled=0
End Sub

Sub S9_Hit:Controller.Switch(9)=1:End Sub					'9
Sub S9_unHit:Controller.Switch(9)=0:End Sub
Sub S10_Hit:Controller.Switch(10)=1:End Sub					'10
Sub S10_unHit:Controller.Switch(10)=0:End Sub
Sub S12_Hit:bsSaucer.AddBall 0:End Sub						'12
Sub S15_Hit:vpmTimer.PulseSw 15:End Sub						'15
Sub S16_Hit:vpmTimer.PulseSw 16:End Sub						'16
Sub S17_Spin:vpmTimer.PulseSw 17:End Sub					'17
Sub S18A_Hit:vpmTimer.PulseSw 18:End Sub					'18
Sub S18B_Hit:vpmTimer.PulseSw 18:End Sub
Sub S18C_Hit:vpmTimer.PulseSw 18:End Sub
Sub S25_Hit:bsTrough.AddBall Me:End Sub						'25
Sub S33_Hit:Controller.Switch(33)=1:End Sub					'33
Sub S33_unHit:Controller.Switch(33)=0:End Sub
Sub S34_Hit:Controller.Switch(34)=1:End Sub					'34
Sub S34_unHit:Controller.Switch(34)=0:End Sub
Sub S35_Hit:Controller.Switch(35)=1:End Sub					'35
Sub S35_unHit:Controller.Switch(35)=0:End Sub
Sub S36_Hit:Controller.Switch(36)=1:End Sub					'36
Sub S36_unHit:Controller.Switch(36)=0:End Sub
Sub S37_Hit:Controller.Switch(37)=1:End Sub					'37
Sub S37_unHit:Controller.Switch(37)=0:End Sub

Dim Old28,New28,Old97,New97
Old28=0:New28=0:Old97=0:New97=0

Set LampCallback=GetRef("UpdateMultipleLamps")

Sub UpdateMultipleLamps
	New28=G28.State
		If New28<>Old28 Then
			If New28=1 Then
				'VariReset
				VariTargetTimer.Enabled = True
			End If
		End If
	Old28=New28

	New97=G97.State
		If New97<>Old97 Then
			If New97=1 Then
				Flipper1.RotateToEnd
 			Else
 				Flipper1.RotateToStart
			End If
		End If
	Old97=New97

  	'not using these - they are correct but unreliable
	'New1=G1.State 'SOLENOID 9
	'	If New1<>Old1 Then
	'		If New1=1 Then PlaySound"sling"
	'	End If
	'Old1=New1
	'New2=G2.State 'SOLENOID 10
	'	If New2<>Old2 Then
	'		If New2=1 Then PlaySound"sling"
	'	End If
	'Old2=New2
	'New3=G3.State 'SOLENOID 11
	'	If New3<>Old3 Then
	'		If New3=1 Then PlaySound"jet3"
	'	End If
	'Old3=New3
	'New17=G17.State 'SOLENOID 12
	'	If New17<>Old17 Then
	'		If New17=1 Then PlaySound"jet3"
	'	End If
	'Old17=New17
	'New18=G18.State	'SOLENOID 13
	'	If New18<>Old18 Then
	'		If New18=1 Then dtDrop.DropSol_On
	'	End If
	'Old18=New18
	'New19=G19.State 'SOLENOID 14
	'	If New19<>Old19 Then
	'		If New19=1 Then
	'			bsSaucer.ExitSol_On
	'		End If
	'	End If
	'Old19=New19
	'New33=G33.State 'SOLENOID 15
	'	If New33<>Old33 Then
	'		If New33=1 Then PlaySound"knocker"
	'	End If
	'Old33=New33
	'New34=G34.State 'SOLENOID 16
	'	If New34<>Old34 Then
	'		If New34=1 Then
	'			If bsTrough.Balls Then bsTrough.ExitSol_On
	'		End If
	'	End If
	'Old34=New34
End Sub

Dim Digits(32)
Digits(0)=Array(Light1,Light2,Light3,Light4,Light5,Light6,Light7,Light8)
Digits(1)=Array(Light15,Light9,Light10,Light11,Light12,Light13,Light14)
Digits(2)=Array(Light22,Light16,Light17,Light18,Light19,Light20,Light21)
Digits(3)=Array(Light29,Light23,Light24,Light25,Light26,Light27,Light28,Light30)
Digits(4)=Array(Light37,Light31,Light32,Light33,Light34,Light35,Light36)
Digits(5)=Array(Light44,Light38,Light39,Light40,Light41,Light42,Light43)
Digits(6)=Array(Light51,Light45,Light46,Light47,Light48,Light49,Light50)
Digits(7)=Array(Light103,Light104,Light105,Light106,Light107,Light108,Light109,Light110)
Digits(8)=Array(Light117,Light111,Light112,Light113,Light114,Light115,Light116)
Digits(9)=Array(Light124,Light118,Light119,Light120,Light121,Light122,Light123)
Digits(10)=Array(Light131,Light125,Light126,Light127,Light128,Light129,Light130,Light132)
Digits(11)=Array(Light139,Light133,Light134,Light135,Light136,Light137,Light138)
Digits(12)=Array(Light146,Light140,Light141,Light142,Light143,Light144,Light145)
Digits(13)=Array(Light153,Light147,Light148,Light149,Light150,Light151,Light152)
Digits(14)=Array(Light52,Light53,Light54,Light55,Light56,Light57,Light58,Light59)
Digits(15)=Array(Light66,Light60,Light61,Light62,Light63,Light64,Light65)
Digits(16)=Array(Light73,Light67,Light68,Light69,Light70,Light71,Light72)
Digits(17)=Array(Light80,Light74,Light75,Light76,Light77,Light78,Light79,Light81)
Digits(18)=Array(Light88,Light82,Light83,Light84,Light85,Light86,Light87)
Digits(19)=Array(Light95,Light89,Light90,Light91,Light92,Light93,Light94)
Digits(20)=Array(Light102,Light96,Light97,Light98,Light99,Light100,Light101)
Digits(21)=Array(Light154,Light155,Light156,Light157,Light158,Light159,Light160,Light161)
Digits(22)=Array(Light168,Light162,Light163,Light164,Light165,Light166,Light167)
Digits(23)=Array(Light175,Light169,Light170,Light171,Light172,Light173,Light174)
Digits(24)=Array(Light182,Light176,Light177,Light178,Light179,Light180,Light181,Light183)
Digits(25)=Array(Light190,Light184,Light185,Light186,Light187,Light188,Light189)
Digits(26)=Array(Light197,Light191,Light192,Light193,Light194,Light195,Light196)
Digits(27)=Array(Light204,Light198,Light199,Light200,Light201,Light202,Light203)
Digits(28)=Array(Light211,Light205,Light206,Light207,Light208,Light209,Light210)
Digits(29)=Array(Light218,Light212,Light213,Light214,Light215,Light216,Light217)
Digits(30)=Array(Light225,Light219,Light220,Light221,Light222,Light223,Light224)
Digits(31)=Array(Light232,Light226,Light227,Light228,Light229,Light230,Light231)
Digits(32)=Array(Light239,Light233,Light234,Light235,Light236,Light237,Light238,Light240)

Sub DisplayTimer_Timer
	Dim ChgLED,ii,num,chg,stat,obj
	ChgLED=Controller.ChangedLEDs(&Hffffffff,&Hffffffff)
	If Not IsEmpty(ChgLED) Then
		For ii=0 To UBound(chgLED)
			num=chgLED(ii,0):chg=chgLED(ii,1):stat=chgLED(ii,2)
			For Each obj In Digits(num)
				If chg And 1 Then obj.State=stat And 1
				chg=chg\2:stat=stat\2
			Next
		Next
	End If
End Sub

 'Aqualand
Sub editDips
Dim vpmDips:Set vpmDips=New cvpmDips
With vpmDips
.AddForm 700,600,"Aqualand - DIP switches"
.AddFrame 0,0,190,"25 psts coin slot - credits",5,Array("2:1",5,"1:1",0,"1:2",4,"1:3",1)'vpm dip 1 and 3			'80
.AddFrame 220,0,190,"100 psts coin slot - credits",80,Array("1:1",0,"1:2",64,"1:3",16,"1:10",80)'vpm dip 5 and 7	'80
.AddFrame 0,80,190,"Balls per game",1024,Array("3 balls",0,"5 balls",1024)'vpm dip 11								'46
.AddFrame 220,80,150,"Match Feature",32768,Array("ON",32768,"OFF",0)'vpm dip 16										'46
.AddLabel 106,130,270,20,"Disable score replays by checking both of these"											'15
.AddChk 126,145,200,Array("Ignore first replay if under 1.000.000",32)'vpm dip 6									'15
.AddChk 126,160,200,Array("Only allow 1 score replay per game",64)'vpm dip 7										'15
.AddChk 21,185,190,Array("Enable Extra Ball via score award:",8)'vpm dip 4											'15
.AddLabel 160,200,280,20,"600.000"
.AddLabel 160,214,280,20,"700.000"
.AddLabel 160,228,280,20,"900.000"
.AddLabel 160,242,280,20,"800.000"
.AddFrame 220,186,190,"Score Replay Levels",768,Array("800.000 - 1.000,000 - 1.200.000",0,"900.000 - 1.000,000 - 1.200.000",256,"1.000.000 - 1.500.000 - 2.000.000",768,"1.000.000 - 1.000.000 - 1.400.000",512)'vpm dip 9 and 10	'80
.AddFrame 160,280,100,"Bookkeeping",524288,Array("bookkeeping off",0,"coin audits",262144,"play audits",524288)'vpm dip 19 and 20 	'75
.AddLabel 80,350,280,20,"After hitting OK, press F3 to reset game with new settings."
.ViewDips
End With
End Sub
Set vpmShowDips=GetRef("editDips")


'***********************************************************************************
'****				       		VariTarget Handling	     				    	****
'***********************************************************************************
Dim VariNewPos, VariOPos, VariSwitch, i
Dim VariSwitches: VariSwitches = Array(24, 23, 22, 21, 20, 19)
Dim VariPositions: VariPositions = Array(0, 4, 8, 12, 16, 20)
Const VT_Delay_Factor = 0.87									'used to slow down the ball when hitting the vari target
dim switchset

Sub vtn_Hit(vidx)
	if switchset=0 then
	VariNewPos=vidx
	If ((ActiveBall.VelY < 0) AND (vidx >= VariOPos)) Then
		ActiveBall.VelY = ActiveBall.VelY * VT_Delay_Factor
		Playsound SoundFX("fx_SolenoidOff"),0
		DOF 101, 2
		For i = 0 to 5:If vidx >= VariPositions(i) Then VariSwitch = VariSwitches(i):End If:Next
		VariOPos=vidx: VariTargetP.TransZ = (VariOPos * 10)
	End If
	If ActiveBall.VelY >= 0 Then
		controller.switch(variswitch)=1:switchset=1
	End If
	end if
End Sub

Sub VariTargetTimer_Timer
	If VariTargetP.TransZ > 0 Then
		VariTargetP.TransZ = VariTargetP.TransZ - 5
	Else
		VariOPos = 0: Me.Enabled = 0:controller.switch(variswitch)=0:switchset = 0':textbox1.text="WIEDER AUS"
	End If
End Sub

Sub VariTarget_Init
	VariNewPos=0: VariOPos=0: VariSwitch = 0
End Sub


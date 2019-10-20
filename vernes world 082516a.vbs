'table by freneticamnesic

' Thalamus 2018-07-24
' Table doesn't have "Positional Sound Playback Functions" or "Supporting Ball & Sound Functions"
' No special SSF tweaks yet.

Option Explicit
Randomize

' Thalamus 2018-11-01 : Improved directional sounds
' !! NOTE : Table not verified yet !!

' Options
' Volume devided by - lower gets higher sound

Const VolDiv = 5000    ' Lower number, louder ballrolling/collition sound
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

Dim UseVPMDMD, VarHidden
If table.ShowDT = true then
UseVPMDMD = true
VarHidden = 1
TextBox.visible = true
else
UseVPMDMD = false
VarHidden = 0
TextBox.visible = false
end if

	Const DB2SOn = False

    LoadVPM "01530000","spinball.vbs",3.10

     Sub LoadVPM(VPMver, VBSfile, VBSver)
       On Error Resume Next
       If ScriptEngineMajorVersion < 5 Then MsgBox "VB Script Engine 5.0 or higher required"
       ExecuteGlobal GetTextFile(VBSfile)
       If Err Then MsgBox "Unable to open " & VBSfile & ". Ensure that it is in the same folder as this table. " & vbNewLine & Err.Description
	   'If DB2SOn = True then
		Set Controller = CreateObject("B2S.Server")
	   'Else
		'Set Controller = CreateObject("VPinMAME.Controller")
	   'End If
       If Err Then MsgBox "Can't Load VPinMAME." & vbNewLine & Err.Description
       If VPMver > "" Then If Controller.Version < VPMver Or Err Then MsgBox "VPinMAME ver " & VPMver & " required."
       If VPinMAMEDriverVer < VBSver Or Err Then MsgBox VBSFile & " ver " & VBSver & " or higher required."
       On Error Goto 0
     End Sub



'********************
'Standard definitions
'********************

	Const cGameName = "vrnwrld"

     Const UseSolenoids = 1
     Const UseLamps = 0
     Const UseSync = 1
     Const HandleMech = 0

     'Standard Sounds
     Const SSolenoidOn = "Solenoid"
     Const SSolenoidOff = ""
     Const SCoin = "CoinIn"
	 Const SFlipperOn="FlipperUp"
	 Const SFlipperOff="FlipperDown"


'*DOF method for rom controller tables by Arngrim********************
'*************************use Tabletype rom or em********************
'*******Use DOF 1**, 1 to activate a ledwiz output*******************
'*******Use DOF 1**, 0 to deactivate a ledwiz output*****************
'*******Use DOF 1**, 2 to pulse a ledwiz output**********************
Sub DOF(dofevent, dofstate)
	If DB2SOn=True Then
		If dofstate = 2 Then
			Controller.B2SSetData dofevent, 1:Controller.B2SSetData dofevent, 0
		Else
			Controller.B2SSetData dofevent, dofstate
		End If
	End If
End Sub
'********************************************************************


'************
' Table init.
'************
   'Variables
    Dim xx
    Dim Bump1,Bump2,Bump3,Mech3bank,bsTrough,bsVUK,bsSVUK,visibleLock,bsTEject
	Dim dtRDrop
	Dim PlungerIM
	Dim bsBallKick
'
  Sub Table_Init
	Gate6.Collidable = False
	Gate7.Collidable = False
	volcanostop.isdropped = false
	volcanostop.isdropped = true
	pulpopost.isdropped = false

		Controller.SolMask(0)=0
		vpmTimer.AddTimer 3000,"Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the timer to renable all the solenoids after 3 seconds

	With Controller
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
		.SplashInfoLine = "Verne's World"
		.HandleKeyboard = 0
		.ShowTitle = 0
		.ShowDMDOnly = 1
		.ShowFrame = 0
		.HandleMechanics = 1
		if VarHidden = 1 then
		.Hidden = 1
		Else
		.Hidden = 0
		End If
		On Error Resume Next
		.Run GetPlayerHWnd
		If Err Then MsgBox Err.Description
	End With


    On Error Goto 0

'**Trough
	Set bsTrough=New cvpmBallStack
	bsTrough.InitSw 0,73,72,71,70,0,0,0
	bsTrough.InitKick BallRelease,62,8
	bsTrough.InitExitSnd"BallRel","SolOn"
	bsTrough.Balls=4

    Set bsVUK = New cvpmBallStack
    With bsVUK
        .InitSw 0, 67, 0, 0, 0, 0, 0, 0
        .InitKick catapult, 180, 220 'name, direction, force
        .InitExitSnd "scoopexit", "scoopexit"
		.InitAddSnd "scoopenter"
        .KickAngleVar = 0
        .KickBalls = 1
    End With

    Set bsSVUK = New cvpmBallStack
    With bsSVUK
        .InitSw 0, 50, 0, 0, 0, 0, 0, 0
        .InitKick volcanokicker, 0, 120
        .KickZ = 100
        .InitExitSnd "scoopexit", "scoopexit"
		.InitAddSnd "scoopenter"
        .KickBalls = 1
    End With

    	vpmNudge.TiltSwitch=-7
   	vpmNudge.Sensitivity=1
   	vpmNudge.TiltObj=Array(Bumper1b,Bumper2b,LeftSlingshot)


    ' Drop Targets
     Set dtRDrop = new cvpmDropTarget
     With dtRDrop
	      .Initdrop Array(sw41, sw42, sw43, sw44), Array(41, 42, 43, 44)
	      .InitSnd "DTResetR","DTR"
      End With

      '**Main Timer init
	PinMAMETimer.Interval = PinMAMEInterval
	PinMAMETimer.Enabled = 1

  End Sub


'***Keys
Sub Table_KeyDown(ByVal Keycode)
dim tt
tt = 0
If keycode = 17 then
	vpmTimer.PulseSw 52: debug.print "52=pulsed" &" lock.  Time(sec)=     " & timer 'w balloon lock
End if
If keycode = 18 then
	vpmTimer.PulseSw 77: debug.print "77=pulsed" &" loop.  Time(sec)=      " & timer 'e left loop
End if
If keycode = 19 then
	vpmTimer.PulseSw 45: debug.print "45=pulsed" &" lower.  Time(sec)=      " & timer 'e left loop
End if

If keycode = 20 then
	vpmTimer.PulseSw 85: debug.print "85=pulsed" &" control.  Time(sec)=      " & timer 'e left loop
'	if Controller.switch(85) = 0 then tt =1
'	Controller.switch(85) = tt: debug.print "85=" & tt &" control.  Time(sec)=   " & timer 't control
End if
If keycode = 21 then
	if Controller.switch(55) = 0 then tt =1
	Controller.switch(55) = tt: debug.print "55=" & tt &" middle.  Time(sec)=    " & timer 'y midde
End if
If keycode = 22 then
	if Controller.switch(65) = 0 then tt =1
	Controller.switch(65) = tt: debug.print "65=" & tt &" upper.  Time(sec)=     " & timer 'u upper
End if
If keycode = 23 then
	vpmTimer.PulseSw 57: debug.print "57=pulsed" &" ramp exit.  Time(sec)=      " & timer 'e left loop
'	if Controller.switch(57) = 0 then tt =1
'	Controller.switch(57) = tt: debug.print "57=" & tt &" ramp exit.  Time(sec)= " & timer 'i ramp exit
End if

' https://vpinball.com/forums/topic/vernes-world-spinball-1996/page/4/#post-103175

If keycode = RightMagnaSave then Controller.Switch(75) = 1

	If keycode=AddCreditKey then vpmTimer.pulseSW (swCoin1)
	if keycode = LeftMagnaSave then bstrough.addball 0: debug.print "drain"
 	If Keycode = LeftFlipperKey then
		'SolLFlipper true
		Controller.Switch(133)=1
	End If
 	If Keycode = RightFlipperKey then
		'SolRFlipper true
		Controller.Switch(131)=1
	End If
    If keycode = PlungerKey Then  Controller.Switch(86) = 1
'  	If keycode = LeftTiltKey Then LeftNudge 80, 1, 20:End If
'   If keycode = RightTiltKey Then RightNudge 280, 1, 20:End If
'   If keycode = CenterTiltKey Then CenterNudge 0, 1, 25 End If
   If vpmKeyDown(keycode) Then Exit Sub
End Sub





Sub Table_KeyUp(ByVal keycode)
'

	If vpmKeyUp(keycode) Then Exit Sub
 	If Keycode = LeftFlipperKey then
		Controller.Switch(133)=0
	End If
 	If Keycode = RightFlipperKey then
		Controller.Switch(131)=0
	End If
	'If Keycode = StartGameKey Then Controller.Switch(16) = 0
    If keycode = PlungerKey Then  Controller.Switch(86) = 0
End Sub


   'Solenoids guesses based on jolly park

'1 - on at start (attract)
'2 - on at start (attract)
'
'3- taca (turns on 5 & 25) 'unknown
'x4- salida bola
'x11- lanzador bola
'14-expulsor banda 'ejector band (slingshot)
'10-catapulta 'catapult

'x8-bancada diana 'right bank
'x16-bumper derecho 'right bumper
'x15-bumper izquierdo 'left bumper
'13-elevador volcan 'elevator volcano
'9-entrada volcan 'volcano entrance
'12-entrada globo 'balloon entry
'20-ataque pulpo 'octopus attack
'6-ret bolas pulpo 'return octopus balls lol
'23??? turns on and stays on when sw45 (balloon control) is hit


SolCallback(4) = "bsTrough.Solout" 'salida bola
SolCallback(11) = "solAutofire" 'autolaunch 'lanzador bola
SolCallback(8) = "SolRightBank" 'bancada diana
'SolCallback(sLLFlipper) = "vpmSolFlipper LeftFlipper,Nothing,"
'SolCallback(sLRFlipper) = "vpmSolFlipper RightFlipper,UpRightFlipper,"
SolCallback(sLLFlipper) = "SolLFlipper"
SolCallback(sLRFlipper) = "SolRFlipper"

	'base table solcallback
'SolCallback(2) = "solBallRelease" 'ballrelease
'SolCallback(4) = "BallKicker"
'SolCallback(5)= "bsSVUK.SolOut"
SolCallback(6)= "PulpoDiverter"
'SolCallback(7)= "SolLeftBank"
SolCallback(9)= "VolcEntrada"
SolCallback(10)="bsVUK.SolOut"
SolCallBack(12)="solCSDiverter"
SolCallback(13) = "bsSVUK.SolOut"
'SolCallback(16)="ShakeTimer.Enabled="
SolCallback(15) = "SolLeftPop" 'left bumper
SolCallback(16) = "SolRightPop" 'right bumper
'SolCallback(19) = "SolRightPop" 'right turbo bumper
SolCallback(20) = "SolPulpo" 'octopus ATTTTTTTACK
'SolCallback(21) = 'right slingshot
'SolCallback(22) = "vpmSolDiverter Flipper4,True," 'upper playfield right diverter
SolCallBack(23)="corkscrewMotor"


Sub SolRightBank(Enabled)
	If Enabled Then
		dtRDrop.DropSol_On
	End If
End Sub


Sub VolcEntrada(Enabled)
	If Enabled Then
		Gate6.collidable = true
		Gate7.collidable = true
		volcanostop.isdropped = true
		volcanostop2.isdropped = false
	Else
		Gate6.collidable = false
		Gate7.collidable = false
		volcanostop.isdropped = false
		volcanostop2.isdropped = true
	End If
End Sub


Sub PulpoDiverter(Enabled)
	If Enabled Then
		pulpopost.isdropped = true
	Else
		pulpopost.isdropped = false
	End If
End Sub

Sub Trigger1_Hit			'top entrance trigger
    Trigger1.DestroyBall	'top entrance destroy ball
	Kicker1.CreateBall		'below top entrance
	Kicker1.kick 160, 8		'below top entrance angle, force
End Sub

Sub catapult_Hit
    'Trigger2.DestroyBall
	'bsVUK.AddBall Me
	'VUK.kick 180, 12
	'Playsound "scoopexit"
End Sub



''Flashers

Sub FlashPops(Enabled)
	If Enabled Then
		SetLamp 126, 1
		SetLamp 134, 1
	Else
		SetLamp 126, 0
		SetLamp 134, 0
	End If
End Sub

Sub SolLeftPop(Enabled)
	If Enabled Then
		'SetLamp 60, 1
		'UpdateGI 4, 3
	Else
		'SetLamp 60, 0
		'UpdateGI 4, 8
	End If
End Sub

Sub SolRightPop(Enabled)
	If Enabled Then
		'UpdateGI 4, 3
	Else
		'SetLamp 61, 0
	End If
End Sub

Sub SolBottomPop(Enabled)
	If Enabled Then
		SetLamp 62, 1
	Else
		SetLamp 62, 0
	End If
End Sub


Sub SolSpinnerFlash(Enabled)
If slingflashers = False then
	If Enabled Then
		SetLamp 132, 1
	Else
		SetLamp 132, 0
	End If
Else
	If Enabled Then
		SetLamp 132, 1
	Else
		SetLamp 132, 0

	End if
End if
End Sub

Sub solTrough(Enabled)
	If Enabled Then
		bsTrough.ExitSol_On
		'vpmTimer.PulseSw 15
	End If
 End Sub

Sub SolBallRelease(Enabled):If Enabled Then:Controller.Switch(15)=0:BallRelease.Kick 62,8::PlaysoundAtVol"EnterKicker",BallRelease,1:End If:End Sub

Sub solAutofire(Enabled)
	If Enabled Then
		PlungerIM.AutoFire
	End If
 End Sub

Sub SolPulpo(Enabled)
	If Enabled Then
		pulpomove = true
		'Solenoid20.State= 1
	Else
		'Solenoid20.State= 0
	End If
End Sub


'primitive flippers!
dim MotorCallback
Set MotorCallback = GetRef("GameTimer")
Sub GameTimer
    UpdateFlipperLogos
End Sub

Sub UpdateFlipperLogo_Timer
    LFLogo.RotY = LeftFlipper.CurrentAngle +90
    RFlogo.RotY = RightFlipper.CurrentAngle +90
    RFlogo2.RotY = UpRightFlipper.CurrentAngle +90
End Sub


Sub SolLFlipper(Enabled)
     If Enabled Then
		 PlaySoundAtVol "flipperupleft", LeftFlipper, VolFlip
		 LeftFlipper.RotateToEnd:
     Else
		 LeftFlipper.RotateToStart:
		playsoundAtVol "flipperdown", LeftFlipper, VolFlip
     End If
 End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
		 PlaySoundAtVol "flipperupright", RightFlipper, VolFlip
		 PlaySoundAtVol "flipperupright", UpRightFlipper, VolFlip
		 RightFlipper.RotateToEnd
		 UpRightFlipper.RotateToEnd
     Else
		 RightFlipper.RotateToStart
		 UpRightFlipper.RotateToStart
		playsoundAtVol "flipperdown", RightFlipper, VolFlip
		playsoundAtVol "flipperdown", UpRightFlipper, VolFlip
     End If
 End Sub

   Sub Drain_Hit():PlaySoundAtVol "Drain",Drain,1:bsTrough.AddBall Me:End Sub
   Sub BallRelease_UnHit():End Sub

Sub RLS_Timer()
    sw60p.RotX = -(sw60s.currentangle) + 90
	sw92pr.z = sw92p.z
	sw93pr.z = sw93p.z
End Sub

pulpodropwall.isdropped = true
Dim pulpomove:pulpomove = False
Sub PulpoT_Timer()
	If pulpomove = True and pulpo.TransZ >= -51 then
		pulpo.TransZ = pulpo.TransZ - 1
	End If
	If pulpomove = False and pulpo.TransZ <= 0 then
		pulpo.TransZ = pulpo.TransZ + 1
	End If
	'If pulpo.TransZ = 0 then pulpomove = True
	If pulpo.TransZ = -51 then pulpomove = False
	pulpodoor.TransZ = pulpo.TransZ
	'corkscrew.ObjRotZ = corkscrew.ObjRotZ + 2
	If pulpodoor.TransZ < -5 then pulpodropwall.isdropped = false
	If pulpodoor.TransZ > -4 then pulpodropwall.isdropped = true
End Sub


'*************
' Targets
'*************

'***Drop Targets
Sub sw41_Hit:dtRDrop.Hit 1:End Sub
Sub sw42_Hit:dtRDrop.Hit 2:End Sub
Sub sw43_Hit:dtRDrop.Hit 3:End Sub
Sub sw44_Hit:dtRDrop.Hit 4:End Sub

'***Standup Targets
'Sub sw41_Hit:Controller.Switch(41)=1:PlaySound "rollover":End Sub
'Sub sw42_Hit:Controller.Switch(42)=1:PlaySound "rollover":End Sub
'Sub sw43_Hit:Controller.Switch(43)=1:PlaySound "rollover":End Sub
'Sub sw44_Hit:Controller.Switch(44)=1:PlaySound "rollover":End Sub
'Sub sw47_Hit:Controller.Switch(47)=1:PlaySound "rollover":End Sub
'Sub sw51_Hit:Controller.Switch(51)=1:PlaySound "rollover":End Sub
'Sub sw52_Hit:Controller.Switch(52)=1:PlaySound "rollover":End Sub
'Sub sw53_Hit:Controller.Switch(53)=1:PlaySound "rollover":End Sub
'Sub sw54_Hit:Controller.Switch(54)=1:PlaySound "rollover":End Sub
'Sub sw66_Hit:Controller.Switch(66)=1:PlaySound "rollover":End Sub
'Sub sw90_Hit:Controller.Switch(90)=1:PlaySound "rollover":End Sub
'Sub sw91_Hit:Controller.Switch(91)=1:PlaySound "rollover":End Sub

'Sub sw1_Hit  : vpmTimer.PulseSw 1:Me.TimerEnabled = 1:sw1p.TransX = -4: playsound "target": End Sub
'Sub sw1_Timer:Me.TimerEnabled = 0:sw1p.TransX = 0:End Sub
Sub sw47_Hit  : vpmTimer.PulseSw 47:Me.TimerEnabled = 1:sw47p.TransX = 4: playsoundAtVol "target", ActiveBall, VolTarg: End Sub
Sub sw47_Timer:Me.TimerEnabled = 0:sw47p.TransX = 0:End Sub
Sub sw51_Hit  : vpmTimer.PulseSw 51:Me.TimerEnabled = 1:sw51p.TransX = 4: playsoundAtVol "target", ActiveBall, VolTarg: End Sub
Sub sw51_Timer:Me.TimerEnabled = 0:sw51p.TransX = 0:End Sub
Sub sw52_Hit  : vpmTimer.PulseSw 52:Me.TimerEnabled = 1:sw52p.TransX = 4: playsoundAtVol "target", ActiveBall, VolTarg: End Sub
Sub sw52_Timer:Me.TimerEnabled = 0:sw52p.TransX = 0:End Sub
Sub sw53_Hit  : vpmTimer.PulseSw 53:Me.TimerEnabled = 1:sw53p.TransX = 4: playsoundAtVol "target", ActiveBall, VolTarg: End Sub
Sub sw53_Timer:Me.TimerEnabled = 0:sw53p.TransX = 0:End Sub
Sub sw54_Hit  : vpmTimer.PulseSw 54:Me.TimerEnabled = 1:sw54p.TransX = 4: playsoundAtVol "target", ActiveBall, VolTarg: End Sub
Sub sw54_Timer:Me.TimerEnabled = 0:sw54p.TransX = 0:End Sub
Sub sw66_Hit  : vpmTimer.PulseSw 66:Me.TimerEnabled = 1:sw66p.TransX = 4: playsoundAtVol "target", ActiveBall, VolTarg: End Sub
Sub sw66_Timer:Me.TimerEnabled = 0:sw66p.TransX = 0:End Sub
Sub sw90_Hit  : vpmTimer.PulseSw 90:Me.TimerEnabled = 1:sw90p.TransX = 4: playsoundAtVol "target", ActiveBall, VolTarg: End Sub
Sub sw90_Timer:Me.TimerEnabled = 0:sw90p.TransX = 0:End Sub
Sub sw91_Hit  : vpmTimer.PulseSw 91:Me.TimerEnabled = 1:sw91p.TransX = 4: playsoundAtVol "target", ActiveBall, VolTarg: End Sub
Sub sw91_Timer:Me.TimerEnabled = 0:sw91p.TransX = 0:End Sub

Sub sw50_Hit:bsSVUK.AddBall Me:End Sub

Sub sw60s_Spin:Controller.Switch(60) = 1:sw60s.TimerEnabled = 1:PlaySoundAtVol "Gate", sw60, 1:End Sub
Sub sw60s_Timer:Controller.Switch(60) = 0:sw60s.TimerEnabled = 0:End Sub

Sub sw60_Hit:Controller.Switch(60)=1:End Sub
Sub sw60_unHit:Controller.Switch(60)=0:End Sub

'***Rollover Targets
Sub sw94_Hit:Controller.Switch(94)=1:PlaySoundAtVol "rollover", ActiveBall, VolRol:End Sub
Sub sw94_unHit:Controller.Switch(94)=0:StopSound "rollover":End Sub
Sub sw61_Hit:Controller.Switch(61)=1:PlaySoundAtVol "rollover", ActiveBall, VolRol:End Sub
Sub sw61_unHit:Controller.Switch(61)=0:StopSound "rollover":End Sub
Sub sw62_Hit:Controller.Switch(62)=1:PlaySoundAtVol "rollover", ActiveBall, VolRol:End Sub
Sub sw62_unHit:Controller.Switch(62)=0:StopSound "rollover":End Sub
Sub sw63_Hit:Controller.Switch(63)=1:PlaySoundAtVol "rollover", ActiveBall, VolRol:End Sub
Sub sw63_unHit:Controller.Switch(63)=0:StopSound "rollover":End Sub
Sub sw94a_Hit:Controller.Switch(94)=1:PlaySoundAtVol "rollover", ActiveBall, VolRol:End Sub
Sub sw94a_unHit:Controller.Switch(94)=0:StopSound "rollover":End Sub
Sub sw64_Hit:Controller.Switch(64)=1:PlaySoundAtVol "rollover", ActiveBall, VolRol:End Sub
Sub sw64_unHit:Controller.Switch(64)=0:StopSound "rollover":End Sub
'Sub sw67_Hit:Controller.Switch(67)=1:PlaySound "rollover":End Sub
'Sub sw67_unHit:Controller.Switch(67)=0:StopSound "rollover":End Sub
Sub sw67_Hit:	sw67.DestroyBall: bsVUK.AddBall Me: End Sub
Sub sw57_Hit:
	'Controller.Switch(57)=1
	If DisableSw57 = 0 Then Controller.Switch(57)=1
	DisableSw57 = 0
	PlaySoundAtVol "rollover", ActiveBall, VolRol
End Sub
Sub sw57_unHit:Controller.Switch(57)=0:StopSound "rollover":End Sub
Sub sw74_Hit:Controller.Switch(74)=1:PlaySoundAtVol "rollover", ActiveBall, VolRol:End Sub
Sub sw74_unHit:Controller.Switch(74)=0:StopSound "rollover":End Sub
Sub sw77_Hit:Controller.Switch(77)=1:PlaySoundAtVol "rollover", ActiveBall, VolRol:End Sub
Sub sw77_unHit:Controller.Switch(77)=0:StopSound "rollover":End Sub
Sub sw76_Hit:Controller.Switch(76)=1:PlaySoundAtVol "rollover", ActiveBall, VolRol:End Sub
Sub sw76_unHit:Controller.Switch(76)=0:StopSound "rollover":End Sub
Sub sw80_Hit:Controller.Switch(80)=1:PlaySoundAtVol "rollover", ActiveBall, VolRol:End Sub
Sub sw80_unHit:Controller.Switch(80)=0:StopSound "rollover":End Sub
Sub sw81_Hit:Controller.Switch(81)=1:PlaySoundAtVol "rollover", ActiveBall, VolRol:End Sub
Sub sw81_unHit:Controller.Switch(81)=0:StopSound "rollover":End Sub
Sub sw82_Hit:Controller.Switch(82)=1:PlaySoundAtVol "rollover", ActiveBall, VolRol:End Sub
Sub sw82_unHit:Controller.Switch(82)=0:StopSound "rollover":End Sub

Sub sw83_Hit:Controller.Switch(83)=1:PlaySoundAtVol "rollover", ActiveBall, VolRol:switch83 = true:End Sub
Sub sw83_unHit:Controller.Switch(83)=0:StopSound "rollover":End Sub

Dim switch83:switch83 = False

Sub sw83t_Timer()
	If switch83 = True and sw83p.TransY > -45 then sw83p.TransY = sw83p.TransY - 5
	If switch83 = False and sw83p.TransY < 0 then sw83p.TransY = sw83p.TransY + 2.5
	If sw83p.TransY <= -45 then switch83 = False
End Sub

Sub sw84_Hit:Controller.Switch(84)=1:sw84.TimerEnabled = 1:End Sub 'volcano drop
Sub sw84_Timer:Controller.Switch(84)=0:sw84.TimerEnabled = 0:PlaySound "DROP_RIGHT":End Sub

Sub sw92_Hit:Controller.Switch(92)=1:PlaySoundAtVol "rollover", ActiveBall, VolRol:switch92 = true:End Sub
Sub sw92_unHit:Controller.Switch(92)=0:StopSound "rollover":End Sub
Sub sw93_Hit:Controller.Switch(93)=1:PlaySoundAtVol "rollover", ActiveBall, VolRol:switch93 = true:End Sub
Sub sw93_unHit:Controller.Switch(93)=0:StopSound "rollover":End Sub

Dim switch92:switch92 = False

Sub sw92t_Timer()
	If switch92 = True and sw92p.z > -30 then sw92p.z = sw92p.z - 5
	If switch92 = False and sw92p.z < -10 then sw92p.z = sw92p.z + 2
	If sw92p.z <= -30 then switch92 = False
End Sub

Dim switch93:switch93 = False

Sub sw93t_Timer()
	If switch93 = True and sw93p.z > -30 then sw93p.z = sw93p.z - 5
	If switch93 = False and sw93p.z < -10 then sw93p.z = sw93p.z + 2
	If sw93p.z <= -30 then switch93 = False
End Sub

Sub DropLeft_Hit:DropLeft.TimerEnabled = 1:End Sub
Sub DropLeft_Timer:DropLeft.TimerEnabled = 0:PlaySound "DROP_LEFT":End Sub

Sub DropRight_Hit:DropRight.TimerEnabled = 1:End Sub
Sub DropRight_Timer:DropRight.TimerEnabled = 0:PlaySound "DROP_RIGHT":End Sub

Sub DropRight2_Hit:DropRight2.TimerEnabled = 1:End Sub
Sub DropRight2_Timer:DropRight2.TimerEnabled = 0:PlaySound "DROP_RIGHT":End Sub

csDiverter.isDropped = True
Sub SolCSDiverter (enabled)
	csDiverter.isdropped = not enabled
End Sub

Dim csball, csball1, csball2, csball3, csball4, csballontop, csballReleased
Sub csKicker_Hit
	cskicker.destroyball
	if csball = 0 then csballontop = 1
	csball = csball + 1
	Controller.Switch (45) = 1
	Select Case csball
		Case 1: Set csball1=corkb1.createball: csball1.z = csLowerSwitchPos: csball1.x = 487: csball1.y = 173: MoveBalloon = 0
		Case 2: Set csball2=corkb2.createball: csball2.z = csLowerSwitchPos: csball2.x = 487: csball2.y = 173
		Case 3: Set csball3=corkb3.createball: csball3.z = csLowerSwitchPos: csball3.x = 487: csball3.y = 173
	End Select
End Sub

Sub corkscrewMotor (enabled)
	corkscrewMotorTimer.enabled = enabled
	'Solenoid23.State = enabled
	'debug.print "corkscrewmotor = " & enabled & " time(sec): " & timer
End Sub
Sub corkscrewMotorTimer_timer
	corkscrew.ObjRotZ = corkscrew.ObjRotZ + 12	'turn screw
	if csball >= 1 then csball1.z =  csball1.z + csStepsize: checkballoonswitches csball1, corkb1 'raise ball1
	if csball >= 2 then csball2.z =  csball2.z + csStepsize: checkballoonswitches csball2, corkb2 'raise ball2
	if csball >= 3 then csball3.z =  csball3.z + csStepsize: checkballoonswitches csball3, corkb3 'raise ball3
	if MoveBalloon = 1 then Balloon.transZ = BalloonBall.z - csMiddleSwitchPos	'Raise Balloon
	debug.Print Balloon.transZ

end sub

Const csLowerSwitchPos = 0
Const csMiddleSwitchPos = 100
Const csUpperSwitchPos = 200
Const csStepsize = 2.5
Dim DisableSw57, BalloonBall, MoveBalloon
Sub checkBalloonSwitches (ball, kickername)
	Select Case ball.z
		Case csLowerSwitchPos + 2*csStepsize:   Controller.Switch(45) = 0	'Clear Sw45
		Case 0:  Controller.Switch(85) = 1
		Case 5:  Controller.Switch(85) = 0
		Case csMiddleSwitchPos: 				Controller.Switch(55) = 1	'Ball at middle sw 55
		Case csMiddleSwitchPos + 2*csStepsize: 	Controller.Switch(55) = 0: Set BalloonBall = ball: If MoveBalloon = 0 Then MoveBalloon = 1
		Case csUpperSwitchPos:  				Controller.Switch(65) = 1	'Ball at upper sw 65
		Case csUpperSwitchPos + 2*csStepsize: 								'Ball is at top, kick ball onto ramp
			Controller.Switch(65) = 0
			kickername.kick 180,1
			MoveBalloonT.Enabled = 1: MoveBalloon = 2
			csballReleased = csballReleased + 1
			if csball = 3 and csballReleased = 1 then DisableSw57 = 1		'BUGFIX: if multiball and 1st ball released, disable the exit ramp switch once, so motor works
			if csBall = csballReleased then csball = 0:  csballReleased = 0	'if all balls released, reset ball counts
	End Select
End Sub

Dim MoveBalloonCnt
Sub MoveBalloonT_Timer
	MoveBalloonCnt = MoveBalloonCnt+1
	Select Case MoveBalloonCnt:
		Case 1: Balloon.transZ = 151
		Case 2: Balloon.transZ = 148
		Case 3: Balloon.transZ = 144
		Case 4: Balloon.transZ = 137
		Case 5: Balloon.transZ = 129
		Case 6: Balloon.transZ = 119
		Case 7: Balloon.transZ = 107
		Case 18 Balloon.transZ = 94
		Case 9: Balloon.transZ = 78
		Case 10: Balloon.transZ = 62
		Case 11: Balloon.transZ = 42
		Case 12: Balloon.transZ = 22
		Case 13: Balloon.transZ = 0: 	MoveBalloonT.Enabled = 0: MoveBalloonCnt = 0: If csball > 0 Then MoveBalloon = 1
	End Select
End Sub


 Sub LeftSlingShot_Slingshot
	Leftsling = True
 	PlaySoundAtBallVol "slingshotleft", 1
	vpmTimer.PulseSw 40
  End Sub


Dim Leftsling1:Leftsling1 = False

Dim Leftsling:Leftsling = False
Sub LS_Timer()
	If Leftsling = True and sling2.TransZ > -28 then sling2.TransZ = sling2.TransZ - 4
	If Leftsling = False and sling2.TransZ < -0 then sling2.TransZ = sling2.TransZ + 4
	If sling2.TransZ <= -28 then Leftsling = False
	sling1.TransX = sling2.TransZ
End Sub


     ' Impulse Plunger
    Const IMPowerSetting = 50
    Const IMTime = 0.1
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swplunger, IMPowerSetting, IMTime
		.Switch 74
        .Random .2
        .InitExitSnd "plunger2", "plunger"
        .CreateEvents "plungerIM"
    End With



'***Bumpers

      Sub Bumper1b_Hit:vpmTimer.PulseSw 46:PlaySoundAtVol "bumper", ActiveBall, VolBump:End Sub

      Sub Bumper2b_Hit:vpmTimer.PulseSw 56:PlaySoundAtVol "bumper", ActiveBall, VolBump:End Sub

 Dim LampState(200), FadingLevel(200), FadingState(200)
Dim FlashState(200), FlashLevel(200), FlashMin(200), FlashMax(200)
Dim FlashSpeedUp, FlashSpeedDown
Dim x

AllLampsOff()
LampTimer.Interval = 40 'lamp fading speed
LampTimer.Enabled = 1
'
FlashInit()
FlasherTimer.Interval = 10 'flash fading speed
FlasherTimer.Enabled = 1

Sub LampTimer_Timer()
    Dim chgLamp, num, chg, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp) Then
        For ii = 0 To UBound(chgLamp)
            LampState(chgLamp(ii, 0) ) = chgLamp(ii, 1)
            FadingLevel(chgLamp(ii, 0) ) = chgLamp(ii, 1) + 4
			FlashState(chgLamp(ii, 0) ) = chgLamp(ii, 1)
        Next
    End If

    UpdateLamps
End Sub

Sub FlashInit
    Dim i
    For i = 0 to 200
        FlashState(i) = 0
        FlashLevel(i) = 0
    Next

    FlashSpeedUp = 50   ' fast speed when turning on the flasher
    FlashSpeedDown = 10 ' slow speed when turning off the flasher, gives a smooth fading
    AllFlashOff()
End Sub

Sub AllFlashOff
    Dim i
    For i = 0 to 200
        FlashState(i) = 0
    Next
End Sub

 Sub UpdateLamps()
NFadeLm 1, l1
NFadeL 1, l1r
NFadeLm 2, l2
NFadeL 2, l2r
NFadeLm 3, l3
NFadeL 3, l3r
NFadeL 4, l4
NFadeL 5, l5
NFadeL 6, l6
NFadeL 7, l7
NFadeL 9, l9
NFadeL 10, l10
NFadeL 11, l11
NFadeL 12, l12
NFadeL 13, l13
NFadeL 14, l14
NFadeL 15, l15
NFadeL 17, l17
NFadeL 18, l18
NFadeLm 19, f19a
NFadeL 19, f19b
NFadeLm 20, f20a
NFadeL 20, f20b
NFadeL 21, l21
NFadeL 22, l22
NFadeL 23, l23
NFadeL 25, l25
NFadeL 26, l26
NFadeL 27, l27
NFadeL 28, l28
NFadeL 29, l29
NFadeL 30, l30
NFadeL 31, l31
NFadeL 33, l33
NFadeL 34, l34
NFadeL 35, l35
NFadeL 36, l36
NFadeL 37, l37
NFadeLm 38, l38
NFadeL 38, l38r
NFadeLm 39, l39
NFadeL 39, l39r
NFadeL 41, l41
NFadeL 42, l42
NFadeL 43, l43
NFadeL 44, l44
NFadeL 45, l45
NFadeL 46, l46
NFadeL 47, l47
NFadeL 49, l49
NFadeL 50, l50
NFadeL 51, l51
NFadeL 52, l52
NFadeL 53, l53
NFadeL 54, l54
NFadeLm 55, l55
NfadeL 55, l55r
NFadeL 57, l57
NFadeL 58, l58
NFadeL 59, l59
NFadeL 60, l60
NFadeL 61, l61
NFadeL 62, l62
NFadeL 65, l65
NFadeL 73, f73
NFadeL 74, f74
NFadeLm 77, f77a
NFadeL 77, f77b
NFadeLm 78, f78a
NFadeL 78, f78b
NFadeL 81, l81
NFadeLm 85, f85a
NFadeL 85, f85b
End Sub

Sub FadePrim(nr, pri, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:pri.image = d:FadingLevel(nr) = 0
        Case 3:pri.image = c:FadingLevel(nr) = 1
        Case 4:pri.image = b:FadingLevel(nr) = 2
        Case 5:pri.image = a:FadingLevel(nr) = 3
    End Select
End Sub

Sub NFadeL(nr, a)
    Select Case FadingLevel(nr)
        Case 4:a.state = 0:FadingLevel(nr) = 0
        Case 5:a.State = 1:FadingLevel(nr) = 1
    End Select
End Sub

Sub NFadeLm(nr, a)
    Select Case FadingLevel(nr)
        Case 4:a.state = 0
        Case 5:a.State = 1
    End Select
End Sub

Sub Flash(nr, object)
    Select Case FlashState(nr)
        Case 0 'off
            FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown
            If FlashLevel(nr) < 0 Then
                FlashLevel(nr) = 0
                FlashState(nr) = -1 'completely off
            End if
            Object.opacity = FlashLevel(nr)
        Case 1 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp
            If FlashLevel(nr) > 1000 Then
                FlashLevel(nr) = 1000
                FlashState(nr) = -2 'completely on
            End if
            Object.opacity = FlashLevel(nr)
    End Select
End Sub

 Sub AllLampsOff():For x = 1 to 200:LampState(x) = 4:FadingLevel(x) = 4:Next:UpdateLamps:UpdateLamps:Updatelamps:End Sub


Sub SetLamp(nr, value)
    If value = 0 AND LampState(nr) = 0 Then Exit Sub
    If value = 1 AND LampState(nr) = 1 Then Exit Sub
    LampState(nr) = abs(value) + 4
FadingLevel(nr ) = abs(value) + 4: FadingState(nr ) = abs(value) + 4
End Sub

Sub SetFlash(nr, stat)
    FlashState(nr) = ABS(stat)
End Sub

Sub FlasherTimer_Timer()

 End Sub


Sub GIFlashT_Timer 'basic rom controlled gi flash test
'	f9t.opacity = f9.opacity / 1.3
'	f9f.opacity = f9.opacity / 1.3
'	f10t.opacity = f10.opacity / 1.3
'	f10f.opacity = f10.opacity / 1.3
'	f74a.opacity = f74.opacity * 4
'	f85a.opacity = f85.opacity * 2
'	f85b.opacity = f85.opacity * 4
'	f85c.opacity = f85.opacity * 4
	Dim bulb
	for each bulb in GILights
	bulb.state = l65.state
	next
'	If LampState(65) = 1 then newlighttest.State = 1:End If
'	If LampState(65) = 0 then newlighttest.State = 0:End If
'	If LampState(65) = 1 then newlighttest1.State = 1:End If
'	If LampState(65) = 0 then newlighttest1.State = 0:End If
'	If LampState(65) = 1 then newlighttest2.State = 1:End If
'	If LampState(65) = 0 then newlighttest2.State = 0:End If
'	If LampState(65) = 1 then newlighttest3.State = 1:End If
'	If LampState(65) = 0 then newlighttest3.State = 0:End If
'	If LampState(65) = 1 then newlighttest4.State = 1:End If
'	If LampState(65) = 0 then newlighttest4.State = 0:End If
'	If LampState(65) = 1 then newlighttest5.State = 1:End If
'	If LampState(65) = 0 then newlighttest5.State = 0:End If
'	If LampState(65) = 1 then newlighttest6.State = 1:End If
'	If LampState(65) = 0 then newlighttest6.State = 0:End If
'	If LampState(65) = 1 then newlighttest7.State = 1:End If
'	If LampState(65) = 0 then newlighttest7.State = 0:End If
'	If LampState(65) = 1 then newlighttest8.State = 1:End If
'	If LampState(65) = 0 then newlighttest8.State = 0:End If
'	If LampState(65) = 1 then newlighttest9.State = 1:End If
'	If LampState(65) = 0 then newlighttest9.State = 0:End If
'	If LampState(65) = 1 then GIWhite.State = 1:End If
'	If LampState(65) = 0 then GIWhite.State = 0:End If
	If sw42.isdropped = true and LampState(65) = 1 then sw42l.State = 1
	If sw42.isdropped = false and LampState(65) = 1 then sw42l.State = 0
	If sw42.isdropped = true and LampState(65) = 0 then sw42l.State = 0
	If sw42.isdropped = false and LampState(65) = 0 then sw42l.State = 0
	'If LampState(65) = 0 then sw42l.State = 0
	If sw43.isdropped = true and LampState(65) = 1 then sw43l.State = 1
	If sw43.isdropped = false and LampState(65) = 1 then sw43l.State = 0
	If sw43.isdropped = true and LampState(65) = 0 then sw43l.State = 0
	If sw43.isdropped = false and LampState(65) = 0 then sw43l.State = 0
	'If LampState(65) = 0 then sw43l.State = 0
End Sub


'bumper cap primitive fade


Sub LightReflections_Timer
'	f19r.opacity = f19.opacity / 1.5
'	f20r.opacity = f20.opacity / 1.5
'	f1r.opacity = f1.opacity / 1.5
'	f2r.opacity = f2.opacity / 1.5
'	f3r.opacity = f3.opacity / 1.5
'	f53r.opacity = f53.opacity / 1.5
'	f38r.opacity = f38.opacity / 2.5
'	f39r.opacity = f39.opacity / 2.5
'	f55r.opacity = f55.opacity / 1.5
'	f77r.opacity = f77.opacity
'	f78r.opacity = f78.opacity
End Sub

Sub Table_exit()
	Controller.Pause = False
	Controller.Stop
End Sub

' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX
' PlaySound sound, 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
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
  PlaySound sound, 1, Volume, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'Set all as per ball position & speed, but Vol Multiplier may be used eg; PlaySoundAtBallVol "sound",3

Sub PlaySoundAtBallVol(sound, VolMult)
  PlaySound sound, 0, Vol(ActiveBall) * VolMult, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub

'Set position as bumperX and Vol manually.

Sub PlaySoundAtBumperVol(sound, tableobj, Vol)
  PlaySound sound, 1, Vol, AudioPan(tableobj), 0,0,1, 1, AudioFade(tableobj)
End Sub

'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

Function RndNum(min, max)
    RndNum = Int(Rnd() * (max-min + 1) ) + min ' Sets a random number between min and max
End Function

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "Table" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / Table.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "Table" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / Table.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "Table" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / Table.width-1
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
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), AudioPan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        Else ' Ball on raised ramp
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*.5, AudioPan(BOT(b) ), 0, Pitch(BOT(b) )+50000, 1, 0, AudioFade(BOT(b) )
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
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / (VolDiv/VolCol), AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub


Option Explicit
On Error Resume Next

' Thalamus 2019 March
' Made ready for improved directional sounds, but, there is very few samples in the table.
' Table lacks RollingTimer and BallRolling samples - needs to be added manually.
'
' !! NOTE : Table not verified yet !!

' Options
' Volume devided by - lower gets higher sound

Const VolDiv = 200    ' Lower number, louder ballrolling/collition sound
Const VolCol = 10      ' Ball collition divider ( voldiv/volcol )

Const VolGates  = 1    ' Gates volume.
Const VolMetal  = 1    ' Metals volume.
Const VolPi     = 1    ' Rubber pins volume.
Const VolTarg   = 1    ' Targets volume.
Const VolSpin   = 1.5  ' Spinners volume.
Const VolFlip   = .15  ' Flipper volume.
Const VolSling  = 1    ' Slingshoit volume.
Const VolWire   = 1.8  ' Wireramp volume.


ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01120100","gts3.vbs",3.10

Dim VarRol
If Table1.ShowDT = true then 
	VarRol=0 
	Wall8.IsDropped=False
	Wall9.IsDropped=False
Else 
	VarRol=1
	Wall8.IsDropped=True
	Wall9.IsDropped=True
End If

Const UseSolenoids=2,UseLamps=1,UseSync=1,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff",SFlipperOn="FlipperUp"
Const SFlipperOff="FlipperDown",SCoin="coin3",cGameName="barbwire"

Const swStartButton=4

'SolCallback(1)="vpmSolSound SoundFX(""Jet3"",DOFContactors),"
'SolCallback(2)="vpmSolSound SoundFX(""Sling"",DOFContactors),"
'SolCallback(3)="vpmSolSound SoundFX(""Sling"",DOFContactors),"
SolCallback(6)="bsLK.SolOut"
SolCallback(7)="bsRK.SolOut"
SolCallback(8)="vpmSolAutoPlungeS Plunger1, SoundFX(SSolenoidOn,DOFContactors), 2,"
'SolCallback(9)='Pull BallGate Diverter
SolCallback(10)="vpmSolDiverter BallGate,True,"'Hold Ballgate Diverter
'SolCallback(11)="vpmFlasher F10,"'Lens Unit Coil 1
SolCallback(11) = "setlampF10"
'SolCallback(12)="vpmFlasher F11,"'Lens Unit Coil 2
SolCallback(12) = "setlampF11"
'SolCallback(13)="vpmFlasher F12,"'Lens Unit Coil 3
SolCallback(13) = "setlampF12"
'SolCallback(14)="vpmFlasher F13,"
SolCallback(14) = "setlampF13"
'SolCallback(15)="vpmFlasher F14,"
SolCallback(15) = "setlampF14"
SolCallback(16)="MoveFatso"'Big Fatso Motor
'SolCallback(17)="vpmFlasher F16,"'Bottom Left Dome Flasher
SolCallback(17) = "setlampF16"
'SolCallback(18)="vpmFlasher F17,"'Captive Ball Flasher
SolCallback(18) = "setlampF17"
'SolCallback(19)="vpmFlasher F18,"'Top Left Upkicker Flasher
SolCallback(19) = "setlampF18"
'SolCallback(20)="vpmFlasher F19,"'Top Left Dome Flasher
SolCallback(20) = "setlampF19"
'SolCallback(21)="vpmFlasher F20,"'Left Ramp Flasher
SolCallback(21) = "setlampF20"
'SolCallback(22)="vpmFlasher F21,"'Big Fatso Flasher
'SolCallback(22) = "setlampF21"
'SolCallback(23)="vpmFlasher F22,"'Right Ramp Flasher
SolCallback(23) = "setlampF22"
'SolCallback(24)="vpmFlasher Array(F23,Bumper1),"'Top Right Dome Flasher
SolCallback(24) = "setlampF23"
'SolCallback(25)="vpmFlasher F24,"'Bottom Right Dome Flasher
SolCallback(25) = "setlampF24"
'SolCallback(26)='Lightbox Relay (A)
'SolCallback(27)='Ticket/Coin Meter
SolCallback(28)="bsTrough.SolOut"
SolCallback(29)="bsTrough.SolIn"
SolCallback(30)="vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(32)="vpmNudge.SolGameOn"
'SolCallback(sLLFlipper)="vpmSolFlipper LeftFlipper,Flipper1,"
'SolCallback(sLRFlipper)="vpmSolFlipper RightFlipper,Nothing,"

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
  If Enabled Then
    PlaySoundAtVol SoundFX("FlipperUp",DOFFlippers), LeftFlipper, VolFlip
    PlaySoundAtVol SoundFX("FlipperUp",DOFFlippers), Flipper1, VolFlip
    LeftFlipper.RotateToEnd
    Flipper1.RotateToEnd
  Else
    PlaySoundAtVol SoundFX("FlipperDown",DOFFlippers), LeftFlipper, VolFlip
    PlaySoundAtVol SoundFX("FlipperDown",DOFFlippers), Flipper1, VolFlip
    LeftFlipper.RotateToStart
    Flipper1.RotateToStart
  End If
End Sub

Sub SolRFlipper(Enabled)
  If Enabled Then
    PlaySoundAtVol SoundFX("FlipperUp",DOFFlippers), RightFlipper, VolFlip
    RightFlipper.RotateToEnd
  Else
    PlaySoundAtVol SoundFX("FlipperDown",DOFFlippers), RightFlipper, VolFlip
    RightFlipper.RotateToStart
  End If
End Sub

Dim bsTrough,bsLK,bsRK

Sub setlampF10(Enabled)
	If Enabled then F10.State=1 Else F10.State=0
End Sub

Sub setlampF11(Enabled)
	If Enabled then F11.State=1 Else F11.State=0
End Sub

Sub setlampF12(Enabled)
	If Enabled then F12.State=1 Else F12.State=0
End Sub

Sub setlampF13(Enabled)
	If Enabled then F13.State=1 Else F13.State=0
End Sub

Sub setlampF14(Enabled)
	If Enabled then F14.State=1 Else F14.State=0
End Sub

Sub setlampF16(Enabled)
	If Enabled then F16.State=1:F16a.visible=1 Else F16.State=0:F16a.visible=0
End Sub

Sub setlampF17(Enabled)
	If Enabled then F17.State=1:F17a.visible=1 Else F17.State=0:F17a.visible=0
End Sub

Sub setlampF18(Enabled)
	If Enabled then F18.State=1:F18a.visible=1 Else F18.State=0:F18a.visible=0
End Sub

Sub setlampF19(Enabled)
	If Enabled then F19.State=1:F19b.State=1:F19c.visible=1 Else F19.State=0:F19b.State=0:F19c.visible=0
End Sub

Sub setlampF10(Enabled)
	If Enabled then F10.State=1 Else F10.State=0
End Sub

Sub setlampF20(Enabled)
	If Enabled then F20.State=1:F20b.State=1:F20c.visible=1 Else F20.State=0:F20b.State=0:F20c.visible=0
End Sub

'Sub setlampF21(Enabled)
	'If Enabled then F21.State=1 Else F21.State=0
'End Sub

Sub setlampF22(Enabled)
	If Enabled then F22.State=1:F22b.State=1:F22c.visible=1 Else F22.State=0:F22b.State=0:F22c.visible=0
End Sub

Sub setlampF23(Enabled)
	If Enabled then F23.State=1:F23b.State=1:F23c.visible=1 Else F23.State=0:F23b.State=0:F23c.visible=0
End Sub

Sub setlampF24(Enabled)
	If Enabled then F24.State=1:F24c.visible=1 Else F24.State=0:F24c.visible=0
End Sub

Sub Table1_Init
Wall111.IsDropped=1
Kicker7.Enabled=0
	vpmInit Me
	Plunger1.PullBack
 	On Error Resume Next
 	With Controller
		.GameName=cGameName
		If Err Then MsgBox"Can't start Game"&cGameName&vbNewLine&Err.Description:Exit Sub
		.SplashInfoLine="Barb Wire - Gottlieb 1991"
		.HandleKeyboard=0
		.ShowTitle=0
		.ShowDMDOnly=1
		.HandleMechanics=0
		.ShowFrame=0
		.Games(cGameName).Settings.Value("rol") = VarRol
		.Run GetPlayerHwnd
		If Err Then MsgBox Err.Description
	End With
	On Error Goto 0
	PinMAMETimer.Interval=PinMAMEInterval
	PinMAMETimer.Enabled=1
	vpmNudge.TiltSwitch=151
	vpmNudge.Sensitivity=6
	vpmNudge.TiltObj=Array(Bumper1,Leftslingshot,Rightslingshot)

	vpmMapLights AllLights

	Set bsTrough=New cvpmBallStack
		bsTrough.InitSw 16,0,0,26,0,0,0,0
		bsTrough.InitKick BallRelease,120,2
		bsTrough.InitEntrySnd "SolOn","SolOn"
		bsTrough.InitExitSnd SoundFX("BallRel",DOFContactors),SoundFX("SolOn",DOFContactors)
		bsTrough.Balls=3

 	Set bsLK=New cvpmBallStack
		bsLK.InitSw 0,50,0,0,0,0,0,0
		bsLK.InitKick LUK,150,7
		bsLK.InitExitSnd SoundFX("Popper",DOFContactors),SoundFX("SolOn",DOFContactors)

	Set bsRK=New cvpmBallStack
		bsRK.InitSw 0,25,0,0,0,0,0,0
		bsRK.InitKick RUK,233,7
		bsRK.InitExitSnd SoundFX("Popper",DOFContactors),SoundFX("SolOn",DOFContactors)

	vpmCreateEvents AllSwitches
	Kicker1.CreateBall
	Kicker1.Kick 150,2
 	Kicker2.CreateBall
	Kicker2.Kick 180,1
End Sub

Sub Table1_KeyDown(ByVal KeyCode)
	If KeyCode=KeyFront Then Controller.Switch(1)=1'Buy-In
	If KeyCode=LeftFlipperKey Then Controller.Switch(42)=1
	If KeyCode=RightFlipperkey Then Controller.Switch(43)=1
  If vpmKeyDown(KeyCode) Then Exit Sub
	If KeyCode=PlungerKey Then Plunger.Pullback
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
	If KeyCode=KeyFront Then Controller.Switch(1)=0
	If KeyCode=LeftFlipperKey Then Controller.Switch(42)=0
	If KeyCode=RightFlipperkey Then	Controller.Switch(43)=0
	If vpmKeyUp(KeyCode) Then Exit Sub
	if KeyCode=PlungerKey Then Plunger.Fire:PlaySoundAtVol"Plunger", Plunger, 1
End Sub

Sub Bumper1_Hit:vpmTimer.PulseSw 10:PlaySoundAtBall SoundFX("Jet3",DOFContactors):End Sub
Sub LeftSlingShot_SlingShot:vpmTimer.PulseSw 11:PlaySoundAtBallVol SoundFX("SlingshotLeft",DOFContactors), VolSling:End Sub
Sub RightSlingShot_SlingShot:vpmTimer.PulseSw 12:PlaySoundAtBallVol SoundFX("SlingshotRight",DOFContactors), VolSling:End Sub
Sub Drain_Hit:bsTrough.Addball Me:End Sub
Sub RUK_Hit:bsRK.AddBall Me:End Sub
Sub LUK_Hit:bsLK.AddBall Me:End Sub
Sub Kicker4_Hit:Me.DestroyBall:vpmTimer.PulseSwitch(60),600,"ToRightUpkicker":End Sub
Sub ToRightUpkicker(swNo):bsRK.AddBall 0:End Sub
Sub Kicker3_Hit:Me.DestroyBall:vpmTimer.PulseSwitch(70),200,"ToLeftVuk":End Sub
Sub ToLeftVuk(swNo):bsLK.AddBall 0:End Sub

Dim FatsoPos,FatsoDir
FatsoPos=0:FatsoDir=1

Sub MoveFatso(Enabled)
	If Enabled Then
		If FatsoDir=-1 Then
			FatsoDir=1
		Else
			FatsoDir=-1
		End If
		FatsoTimer.Enabled=1
	Else
		FatsoTimer.Enabled=0
	End If
End Sub

Sub FatsoTimer_Timer
	Primitive2.TransZ=Primitive2.TransZ+FatsoDir
	If Primitive2.TransZ = -100 Then FatsoPos=1:FatsoTimer.Enabled=False
	If Primitive2.TransZ = 0 Then FatsoPos=0:FatsoTimer.Enabled=False
	Select Case FatsoPos
		Case 0:Controller.Switch(27)=1:Controller.Switch(17)=0:BigFatso.IsDropped=0
		Case 1:Controller.Switch(27)=0:Controller.Switch(17)=1:BigFatso.IsDropped=1
	End Select
End Sub

Sub BigFatso_Hit:vpmTimer.PulseSw 70:End Sub

Sub Kicker7_Hit:Me.DestroyBall:Kicker8.CreateBall:Kicker8.Kick 0,15:End Sub

Sub Trigger1_Hit : PlaySound "fx_metalrolling",-1, Vol(ActiveBall)*VolWire, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall) : End Sub
Sub Trigger2_Hit : StopSound "fx_metalrolling" : End Sub
Sub Trigger3_Hit : PlaySound "fx_metalrolling",-1, Vol(ActiveBall)*VolWire, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall) : End Sub
Sub Trigger4_Hit : StopSound "fx_metalrolling": End Sub


'WALLS
'Dim L(74)
'Set L(5)=L5
'Set L(6)=L6
'Set L(7)=L7
'Set L(11)=L11
'Set L(12)=L12
'Set L(13)=L13
'Set L(14)=L14
'Set L(15)=L15
'Set L(16)=L16
'Set L(17)=L17
'Set L(21)=L21
'Set L(22)=L22
'Set L(23)=L23
'Set L(24)=L24
'Set L(25)=L25
'Set L(26)=L26
'Set L(27)=L27
'Set L(32)=L32
'Set L(33)=L33
'Set L(34)=L34
'Set L(35)=L35
'Set L(36)=L36
'Set L(37)=L37
'Set L(42)=L42
'Set L(43)=L43
'Set L(44)=L44
'Set L(45)=L45
'Set L(46)=L46
'Set L(47)=L47
'Set L(50)=L50
'Set L(51)=L51
'Set L(52)=L52
'Set L(53)=L53
'Set L(54)=L54
'Set L(55)=L55
'Set L(56)=L56
'Set L(60)=L60
'Set L(61)=L61
'Set L(62)=L62
'Set L(63)=L63
'Set L(64)=L64
'Set L(65)=L65
'Set L(66)=L66
'Set L(67)=L67
'Set L(70)=L70
'Set L(71)=L71
'Set L(72)=L72
'Set L(73)=L73
'Set L(74)=L74

Set MotorCallback=GetRef("UpdateMultipleLamps")

Dim X

Sub UpdateMultipleLamps
dim chg,count,ii
chg=controller.changedlamps
count=ubound(chg)
On Error Resume Next
if count>0 then
For X=0 To count
	If chg(x,1) Then
		L(chg(x,0)).IsDropped=0
	else
		L(chg(x,0)).IsDropped=1
	end if
Next
Light0.State=ABS(Controller.Lamp(0))
L57.State=ABS(Controller.Lamp(57))
L75.State=ABS(Controller.Lamp(75))
L76.State=ABS(Controller.Lamp(76))
L77.State=ABS(Controller.Lamp(77))
End If
End Sub

Sub flippers_Timer()
	LeftFlipperP.objRotZ = LeftFlipper.CurrentAngle-90
	LeftFlipperP1.objRotZ = Flipper1.CurrentAngle-90
	RightFlipperP.objRotZ = RightFlipper.CurrentAngle-90
End Sub

' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX and Rothbauerw
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

Sub PlaySoundAtBOTBallZ(sound, BOT)
  PlaySound sound, 0, ABS(BOT.velz)/17, AudioPan(BOT), 0, Pitch(BOT), 1, 0, AudioFade(BOT)
End Sub

'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

Function RndNum(min, max)
  RndNum = Int(Rnd() * (max-min + 1) ) + min ' Sets a random number between min and max
End Function

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = tableobj.y * 2 / table1.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = tableobj.x * 2 / table1.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
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

Function BallVelZ(ball) 'Calculates the ball speed in the -Z
  BallVelZ = INT((ball.VelZ) * -1 )
End Function

Function VolZ(ball) ' Calculates the Volume of the sound based on the ball speed in the Z
  VolZ = Csng(BallVelZ(ball) ^2 / 200)*1.2
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
  If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
    PlaySoundAtBOTBallZ "fx_ball_drop" & b, BOT(b)
    'debug.print BOT(b).velz
  End If
  Next
End Sub

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
  PlaySound("fx_collide"), 0, Csng(velocity) ^2 / (VolDiv/VolCol), AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

' ************
' Collections
' ************

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
	PlaySoundAtVol "fx_spinner", Spinner, VolSpin
End Sub

' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub

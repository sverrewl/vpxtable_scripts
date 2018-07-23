Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const Ballsize = 52
Const BallMass = 1.5

LoadVPM "01000200", "DE.VBS", 3.38

      
'********************
'Standard definitions
'********************

Const UseSolenoids = 1
Const UseLamps = 0
Const UseSync = 1
Const HandleMech = 0

' Standard Sounds
Const SSolenoidOn = "fx_Solenoidon"
Const SSolenoidOff = "fx_solenoidoff"
Const SCoin = "fx_Coin"


'Solenoid
SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"
SolCallback(1)  = "bsTrough.SolIn"
SolCallback(2)  = "bsTrough.SolOut"
SolCallback(3)  = "bsLVuk.SolOut"
SolCallback(4)  = "SolKickback" 
SolCallback(5)  = "dtLDrop.SolDropUp"
SolCallback(6)  = "bsUpperEject.SolOut"
SolCallback(8)  = "Solknocker" 
SolCallback(11) = "SolGI"
SolCallBack(15)	= "SolAutoPlunge"
SolCallback(22)  = "SolRdiv"


'Flashers
'SolCallback(20) = "SetLamp 80,"  'Flash 20
'SolCallback(21) = "SetLamp 81,"  'Flash 21

SolCallback(25) = "SolFlash1B"    'Flash 1B
SolCallback(26) = "SetLamp 92,"    'Flash 2B
SolCallback(27) = "SolFlash3B"  'Flash 3B
SolCallback(28) = "SolFlash4B"  'Flash 4B
SolCallback(29) = "SetLamp 95,"  'Flash 5B
SolCallback(30) = "SolFlash6B"  'Flash 6B
SolCallback(31) = "SolFlash7B"  'Flash 7B
'SolCallback(32) = "SetLamp 98,"  'Flash 8B

Const cGameName = "Hook_500"


Dim bsTrough,bsUpperEject,bsCala,bsLVuk,SkillShotR,PlungerIM
DIM dtLDrop, x
Sub Table1_Init
    vpmInit Me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Hook Data East 1992" & vbNewLine & "VPX table by Javier v1.0"
		.HandleKeyboard = 0
		.ShowTitle = 0
		.ShowDMDOnly = 1
		.ShowFrame = 0
		.HandleMechanics = 0
		.Hidden = 0
		.Games(cGameName).Settings.Value("sound") = 1
		On Error Resume Next
		.Run GetPlayerHWnd
		If Err Then MsgBox Err.Description
	End With
    On Error Goto 0

    ' Nudging
    vpmNudge.TiltSwitch = 1
	vpmNudge.Sensitivity = 2 
    vpmNudge.tiltobj = Array(LeftSlingShot,RightSlingShot,Bumper1B,Bumper2B,Bumper3B)

	PinMAMETimer.Interval = PinMAMEInterval
	PinMAMETimer.Enabled = 1

	' Trough
     Set bsTrough = new cvpmTrough 
     With bsTrough
		.Size = 3
		.InitSwitches Array (13,12,11)
		.EntrySw = 10
		.InitExit BallRelease, 90, 6
		.InitEntrySounds "fx_drain", SoundFX(SSolenoidOn,DOFContactors), SoundFX(SSolenoidOn,DOFContactors)
		.InitExitSounds  SoundFX(SSolenoidOn,DOFContactors), SoundFX("fx_ballrel",DOFContactors)
		.Balls = 3
		.CreateEvents "bsTrough", Drain
     End With

    ' Left vuk
    Set bsLVuk = New cvpmSaucer
    With bsLVuk
	    .InitKicker Sw51, 51,110, 30, 1.45
		.InitSounds "fx_sensor", SoundFX(SSolenoidOn,DOFContactors), SoundFX("fx_vukout_LAH",DOFContactors)
		.CreateEvents "bsLVuk", Sw51
    End With 


     'PowerScoop
	Set bsUpperEject = New cvpmSaucer
	With bsUpperEject
	    .InitKicker Sw48b, 48,230, 30, 1.56
		.InitSounds "fx_sensor", SoundFX(SSolenoidOn,DOFContactors), SoundFX("salidadebola",DOFContactors)
		.CreateEvents "bsUpperEject", Sw48b
    End With 

     
'      Drop Targets
  	Set dtLDrop=New cvpmDropTarget
    With dtLDrop
	    .InitDrop Array(Sw25,Sw26,Sw27,Sw28), Array(25,26,27,28)
	    .InitSnd SoundFX("fx_droptarget",DOFDropTargets),SoundFX("fx_resetdrop",DOFContactors)
    End With



       Set SkillShotR = New cvpmBallStack
       With SkillShotR
            .InitSaucer Sw32,32, 0, 45
   		    .KickForceVar = 2
            .InitExitSnd "bumper_retro", "fx_sensor"
          End With


    'Diverter
    DivOpen.Isdropped = 1 
    DivClose.Isdropped = 0

End Sub

Dim BallsOnPlayfield: BallsOnPlayfield = 0 
Sub Trigger1_Hit()
    BallsOnPlayfield = BallsOnPlayfield - 1
    vpmtimer.addtimer 2000, "LEDStopTimer '"
End Sub

Sub LEDStopTimer
	If BallsOnPlayfield <= 0 Then LEDStop
End Sub



Sub BallRelease_Unhit
    BallsOnPlayfield = BallsOnPlayfield + 1
	LEDStart
End Sub



'Knocker
Sub SolKnocker(enabled)
If enabled then Playsound SoundFX("fx_knocker",DOFKnocker)
End Sub



' Kickback
   Sub SolKickback(enabled)
 	If Enabled then	
		SkillShotR.ExitSol_On
	end if
   End Sub


' LaserKick
Sub SolAutoPlunge(Enabled)
	If Enabled Then 
        Playsound SoundFX("bumper_retro",DOFContactors)
		LaserKick.Enabled=True
        LaserKickP1.TransY = 90
	Else
		LaserKick.Enabled=False
        vpmtimer.addtimer 500, "LaserKickRes '"
	End If
End Sub
Sub LaserKick_Hit: Me.Kick 0,30 End Sub

Sub LaserKickRes()
    LaserKickP1.TransY = 0
End Sub




' Ramp Diverter
Sub SolRdiv(Enabled)
    If Enabled Then 
       vpmtimer.addtimer 1450, "DiverterTimer.enabled = 1' "
        DivOpen.Isdropped = 0 
        DivClose.Isdropped = 1
        DivP.RotY = -38
        PlaySound "fx_diverter"
    End If
End Sub       

Sub DiverterTimer_Timer
    DiverterTimer.enabled = 0
    DivClose.Isdropped = 0
    DivOpen.Isdropped = 1
    DivP.RotY = 0
    PlaySound "fx_diverter"
End Sub

'Flashers

Sub SolFlash1B(enabled)
 If enabled Then
    Flasher25.state = 1
    Flasher25a.state = 1
    Flasher25b.state = 1
    Flasher25c.state = 1
    Flasher25d.state = 1
    Flasher25e.state = 1
    F25.opacity = 100
  Else
    Flasher25.state = 0
    Flasher25a.state = 0
    Flasher25b.state = 0
    Flasher25c.state = 0
    Flasher25d.state = 0
    Flasher25e.state = 0
    F25.opacity = 0
 End If
End Sub


Sub SolFlash3B(enabled)
 If enabled Then
    f3a.state = 1
    f3b.state = 1
  Else
    f3a.state = 0
    f3b.state = 0
 End If
End Sub

Sub SolFlash4B(enabled)
 If enabled Then
    f4a.state = 1
    f4b.state = 1
  Else
    f4a.state = 0
    f4b.state = 0
 End If
End Sub


Sub SolFlash6B(enabled)
 If enabled Then
    F6b2.opacity = 50
    f6b.state = 1
    f6c.state = 1
  Else
    F6b2.opacity = 0
    f6b.state = 0
    f6c.state = 0
 End If
End Sub

Sub SolFlash7B(enabled)
 If enabled Then
    F7c.opacity = 100
    F7.state = 1
    f7d.state = 1
    F31.state = 1
  Else
    F7c.opacity = 0
    f7d.state = 0
    F7.state = 0
    F31.state = 0
 End If
End Sub



'******************
'Keys Up and Down
'*****************

Sub Table1_KeyDown(ByVal Keycode)
    If keycode = PlungerKey Then Plunger.Pullback:PlaySound "fx_plungerpull"
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal Keycode)
    If keycode = plungerkey then Plunger.Fire :PlaySound "fx_plunger"
    If vpmKeyUp(keycode) Then Exit Sub
End Sub

Sub Table1_Paused : Controller.Pause = True : End Sub
Sub Table1_unPaused : Controller.Pause = False : End Sub
Sub Table1_Exit() : Controller.Pause = False : Controller.Stop() : End Sub



' Flippers
Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySound SoundFX("fx_flipperUp", DOFFlippers), 0, 1, -0.1, 0.15
        LeftFlipper.RotateToEnd
    Else
        PlaySound SoundFX("fx_Flipperdown", DOFFlippers), 0, 1, -0.1, 0.15
        LeftFlipper.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySound SoundFX("fx_flipperUp", DOFFlippers), 0, 1, 0.1, 0.15
        RightFlipper.RotateToEnd
    Else
        PlaySound SoundFX("fx_Flipperdown", DOFFlippers), 0, 1, 0.1, 0.15
        RightFlipper.RotateToStart
    End If
End Sub



'************************
'	General Illumination
'************************

Sub SolGi(enabled)
  If enabled Then
     GiOFF
	Playsound "fx_relay_off"
    Table1.ColorGradeImage = "ColorGradeLUT256x16_ConSatDark"
    Backdrop1.image = "BackBrop_offoff"
    skullP.image = "HookSkullMap_Dark"
   Else
     GiON
	Playsound "fx_relay_on"
	Table1.ColorGradeImage = "ColorGradeLUT256x16_ConSat"
    Backdrop1.image = "BackBrop_off"
    skullP.image = "HookSkullMap_On"
 End If
End Sub

Sub GiON
    For each x in aGiLights:x.State = 1:Next
    gi29.IntensityScale=1 : gi30.IntensityScale=1
    pfGi_On.opacity = 10
End Sub

Sub GiOFF
    For each x in aGiLights:x.State = 0:Next
    gi29.IntensityScale=0 : gi30.IntensityScale=0
    pfGi_On.opacity = 0
End Sub



' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 500)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / Table1.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10))
    End If
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2)))
End Function

'****************************************
' Real Time updates
'****************************************

Set MotorCallBack = GetRef("GameTimer")

Sub GameTimer
    RollingUpdate
	BallShadowUpdate
	LFLogo.RotY = LeftFlipper.CurrentAngle
	RFlogo.RotY = RightFlipper.CurrentAngle
End Sub

'*********** ROLLING SOUND *********************************
Const tnob = 5						' total number of balls : 4 (trough) + 1 (fake ball used to animate Austin toy)
Const fakeballs = 0					' number of balls created on table start (rolling sound will be skipped)
ReDim rolling(tnob)
InitRolling

Sub InitRolling:Dim i:For i=0 to (tnob-1):rolling(i) = False:Next:End Sub

Sub RollingUpdate()
    Dim BOT, b, ballpitch
    BOT = GetBalls
	' stop the sound of deleted balls
	If UBound(BOT)<(tnob -1) Then
		For b = (UBound(BOT) + 1) to (tnob-1)
			rolling(b) = False
			StopSound("fx_ballrolling" & b)
		Next
	End If
	' exit the Sub if no balls on the table
    If UBound(BOT) = fakeballs-1 Then Exit Sub
	' play the rolling sound for each ball
    For b = fakeballs to UBound(BOT)
        If BallVel(BOT(b) ) > 1 Then
           If BOT(b).z < 30 Then
                ballpitch = Pitch(BOT(b))
            Else
                ballpitch = Pitch(BOT(b)) * 100
            End If
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b)), Pan(BOT(b)), 0, ballpitch, 1, 0
        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If
    Next
End Sub

'*********** BALL SHADOW *********************************
Dim BallShadow
BallShadow = Array (BallShadow1, BallShadow2, BallShadow3,BallShadow4,BallShadow5)

Sub BallShadowUpdate()
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
			BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) + 10
		Else
			BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) - 10
		End If
	    ballShadow(b).Y = BOT(b).Y + 20

		If BOT(b).Z > 20 Then
			BallShadow(b).visible = 1
		Else
			BallShadow(b).visible = 0
		End If
	Next
End Sub




'******************************
'		 Sound FX
'******************************


Sub aMetals_Hit(idx):PlaySound "fx_MetalHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
'Sub aRubber_Bands_Hit(idx):PlaySound "fx_rubber_band", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
'Sub aRubber_Posts_Hit(idx):PlaySound "fx_rubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aRubber_Pins_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aWoods_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub


Function RndNum(min,max)
 RndNum = Int(Rnd()*(max-min+1))+min     ' Sets a random number between min and max
End Function

Sub RubberBands_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR((activeball.velx ^2) + (activeball.vely ^2))
 	If finalspeed > 20 then 
		PlaySound "fx_rubber", 0, 10*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RubbersPost_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR((activeball.velx ^2) + (activeball.vely ^2))
 	If finalspeed > 20 then 
		PlaySound "fx_rubber", 0, 10*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RubberPostPrimitive_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR((activeball.velx ^2) + (activeball.vely ^2))
 	If finalspeed > 20 then 
		PlaySound "fx_rubber", 0, 10*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RandomSoundRubber()
	Select Case RndNum(1,3)
		Case 1 : PlaySound "fx_rubber_hit_1", 0, 10*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 2 : PlaySound "fx_rubber_hit_2", 0, 10*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 3 : PlaySound "fx_rubber_hit_3", 0, 10*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End Select
End Sub


Sub TriggerWR_Hit()
    stopsound "metal"
    stopsound "fx_metalrolling"
End Sub

Sub RRFx_Hit: Playsound "comet_spinner" End Sub

 Sub LRHelp1_Hit()
	ActiveBall.VelY = ActiveBall.VelY/5
	ActiveBall.VelX = ActiveBall.VelX/5
 	me.timerinterval=150
 	me.timerenabled=true
	playSound "BallDrop"
End Sub


Sub RampFX1_hit:vpmtimer.addtimer 300, "BallDropFX' ": End Sub
Sub RampFX2_hit:vpmtimer.addtimer 300, "BallDropFX' ": End Sub
Sub RampFX3_hit:vpmtimer.addtimer 300, "BallDropFX' ": End Sub

'Flipper rubber sounds

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, -0.1, 0.15
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, 0.1, 0.15
End Sub

' Ball Collision Sound
Sub OnBallBallCollision(ball1, ball2, velocity)
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0
End Sub


Sub Oyos123_Hit(idx): vpmtimer.addtimer 800, "BallDropFX' " End Sub

Sub BallDropFX()
    playSound "BallDrop"
End Sub    


'*****************
' Switches
'*****************

' Slingshots
Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    PlaySound SoundFX("LeftSlingShot", DOFContactors), 0, 1, -0.05, 0.05
    LeftSling4.Visible = 1
    LeftSling1.Visible = 0
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 52
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 1:LeftSLing4.Visible = 0:LeftSLing3.Visible = 1:Lemk.RotX = 14
        Case 2:LeftSLing3.Visible = 0:LeftSLing2.Visible = 1:Lemk.RotX = 2
        Case 3:LeftSLing2.Visible = 0:LeftSling1.Visible = 1:Lemk.RotX = -10:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
    PlaySound SoundFX("RightSlingShot", DOFContactors), 0, 1, 0.05, 0.05
    RightSling4.Visible = 1
    RightSling1.Visible = 0
    Remk.RotX = 26
    RStep = 0
    vpmTimer.PulseSw 53
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 1:RightSLing4.Visible = 0:RightSLing3.Visible = 1:Remk.RotX = 14
        Case 2:RightSLing3.Visible = 0:RightSLing2.Visible = 1:Remk.RotX = 2
        Case 3:RightSLing2.Visible = 0:RightSling1.Visible = 1:Remk.RotX = -10:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

' Bumpers
Sub Bumper1B_Hit:vpmTimer.PulseSw 33:PlaySound SoundFX("fx_bumper1",DOFContactors), 0, 1, -0.1, 0.15:End Sub
Sub Bumper2B_Hit:vpmTimer.PulseSw 35:PlaySound SoundFX("fx_bumper3",DOFContactors), 0, 1, 0, 0.15:End Sub
Sub Bumper3B_Hit:vpmTimer.PulseSw 34:PlaySound SoundFX("fx_bumper2",DOFContactors), 0, 1, 0.1, 0.15:End Sub




Sub Sw14_Hit():Playsound "fx_sensor": Controller.Switch(14)=1: End Sub
Sub Sw14_UnHit():Controller.Switch(14)=0: End Sub

Sub Sw17_Hit():Playsound "fx_sensor": Controller.Switch(17)=1: End Sub
Sub Sw17_UnHit():Controller.Switch(17)=0: End Sub

Sub Sw18_Hit():Playsound "fx_sensor": Controller.Switch(18)=1: End Sub
Sub Sw18_UnHit():Controller.Switch(18)=0: End Sub

Sub Sw19_Hit():Playsound "fx_sensor": Controller.Switch(19)=1: End Sub
Sub Sw19_UnHit():Controller.Switch(19)=0: End Sub



Sub Sw20_Hit:vpmTimer.PulseSw 20 :MoveTarget20 :PlaySound SoundFX("fx_bumper",DOFContactors), 0, 1, 0.15, 0.15:End Sub
Sub MoveTarget20
	Sw20a.TransZ = 5
	Sw20b.TransZ = 5
	Sw20.Timerenabled = False
	Sw20.Timerenabled = True
End Sub
Sub	Sw20_Timer
	Sw20.Timerenabled = False
	Sw20a.TransZ = 0
	Sw20b.TransZ = 0
End Sub


Sub Sw21_Hit():Playsound "fx_sensor": Controller.Switch(21)=1: End Sub
Sub Sw21_UnHit():Controller.Switch(21)=0: End Sub

Sub Sw22_Hit():Playsound "fx_sensor": Controller.Switch(22)=1: End Sub
Sub Sw22_UnHit():Controller.Switch(22)=0: End Sub

Sub Sw23_Hit():Playsound "fx_sensor": Controller.Switch(23)=1: End Sub
Sub Sw23_UnHit():Controller.Switch(23)=0: End Sub

Sub Sw24_Hit():Playsound "fx_sensor": Controller.Switch(24)=1: End Sub
Sub Sw24_UnHit():Controller.Switch(24)=0: End Sub



'Drops Targets 
Sub sw25_dropped():dtLDrop.Hit 1:End Sub
Sub sw26_dropped():dtLDrop.Hit 2:End Sub
Sub sw27_dropped():dtLDrop.Hit 3:End Sub
Sub sw28_dropped():dtLDrop.Hit 4:End Sub



Sub Sw29_Hit():Playsound "fx_sensor": Controller.Switch(29)=1: End Sub
Sub Sw29_UnHit():Controller.Switch(29)=0: End Sub

Sub Sw30_Hit():Playsound "fx_sensor": Controller.Switch(30)=1: End Sub
Sub Sw30_UnHit():Controller.Switch(30)=0: End Sub

Sub Sw31_Hit():Playsound "fx_sensor": Controller.Switch(31)=1: End Sub
Sub Sw31_UnHit():Controller.Switch(31)=0: End Sub

Sub Sw32_Hit(): SkillShotR.AddBall 0:End Sub

Sub Sw32_UnHit
    LaserKickP.TransY = 90   
    vpmtimer.addtimer 500, "PlungerRes '"
End Sub

Sub PlungerRes()
    LaserKickP.TransY = 0
End Sub



 
Sub Sw33_Hit:vpmTimer.PulseSw 33 :MoveTarget33 :PlaySound SoundFX("fx_target",DOFContactors), 0, 1, 0.15, 0.15:End Sub
Sub Sw34_Hit:vpmTimer.PulseSw 34 :MoveTarget34 :PlaySound SoundFX("fx_target",DOFContactors), 0, 1, 0.15, 0.15:End Sub
Sub Sw35_Hit:vpmTimer.PulseSw 35 :MoveTarget35 :PlaySound SoundFX("fx_target",DOFContactors), 0, 1, 0.15, 0.15:End Sub
Sub Sw41_Hit:vpmTimer.PulseSw 41 :MoveTarget41 :PlaySound SoundFX("fx_target",DOFContactors), 0, 1, 0.15, 0.15:End Sub
Sub Sw42_Hit:vpmTimer.PulseSw 42 :MoveTarget42 :PlaySound SoundFX("fx_target",DOFContactors), 0, 1, 0.15, 0.15:End Sub
Sub Sw43_Hit:vpmTimer.PulseSw 43 :MoveTarget43 :PlaySound SoundFX("fx_target",DOFContactors), 0, 1, 0.15, 0.15:End Sub
Sub Sw44_Hit:vpmTimer.PulseSw 44 :MoveTarget44 :PlaySound SoundFX("fx_target",DOFContactors), 0, 1, 0.15, 0.15:End Sub
Sub Sw45_Hit:vpmTimer.PulseSw 45 :MoveTarget45 :PlaySound SoundFX("fx_target",DOFContactors), 0, 1, 0.15, 0.15:End Sub
Sub Sw46_Hit:vpmTimer.PulseSw 46 :MoveTarget46 :PlaySound SoundFX("fx_target",DOFContactors), 0, 1, 0.15, 0.15:End Sub

Sub MoveTarget33
	Sw33a.TransZ = 5
	Sw33b.TransZ = 5
	Sw33.Timerenabled = False
	Sw33.Timerenabled = True
End Sub
Sub	Sw33_Timer
	Sw33.Timerenabled = False
	Sw33a.TransZ = 0
	Sw33b.TransZ = 0
End Sub

Sub MoveTarget34
	Sw34a.TransZ = 5
	Sw34b.TransZ = 5
	Sw34.Timerenabled = False
	Sw34.Timerenabled = True
End Sub
Sub	Sw34_Timer
	Sw34.Timerenabled = False
	Sw34a.TransZ = 0
	Sw34b.TransZ = 0
End Sub

Sub MoveTarget35
	Sw35a.TransZ = 5
	Sw35b.TransZ = 5
	Sw35.Timerenabled = False
	Sw35.Timerenabled = True
End Sub
Sub	Sw35_Timer
	Sw35.Timerenabled = False
	Sw35a.TransZ = 0
	Sw35b.TransZ = 0
End Sub

Sub MoveTarget41
	Sw41a.TransZ = 5
	Sw41b.TransZ = 5
	Sw41.Timerenabled = False
	Sw41.Timerenabled = True
End Sub
Sub	Sw41_Timer
	Sw41.Timerenabled = False
	Sw41a.TransZ = 0
	Sw41b.TransZ = 0
End Sub


Sub MoveTarget42
	Sw42a.TransZ = 5
	Sw42b.TransZ = 5
	Sw42.Timerenabled = False
	Sw42.Timerenabled = True
End Sub
Sub	Sw42_Timer
	Sw42.Timerenabled = False
	Sw42a.TransZ = 0
	Sw42b.TransZ = 0
End Sub

Sub MoveTarget43
	Sw43a.TransZ = 5
	Sw43b.TransZ = 5
	Sw43.Timerenabled = False
	Sw43.Timerenabled = True
End Sub
Sub	Sw43_Timer
	Sw43.Timerenabled = False
	Sw43a.TransZ = 0
	Sw43b.TransZ = 0
End Sub


Sub MoveTarget44
	Sw44a.TransZ = 5
	Sw44b.TransZ = 5
	Sw44.Timerenabled = False
	Sw44.Timerenabled = True
End Sub
Sub	Sw44_Timer
	Sw44.Timerenabled = False
	Sw44a.TransZ = 0
	Sw44b.TransZ = 0
End Sub


Sub MoveTarget45
	Sw45a.TransZ = 5
	Sw45b.TransZ = 5
	Sw45.Timerenabled = False
	Sw45.Timerenabled = True
End Sub
Sub	Sw45_Timer
	Sw45.Timerenabled = False
	Sw45a.TransZ = 0
	Sw45b.TransZ = 0
End Sub

Sub MoveTarget46
	Sw46a.TransZ = 5
	Sw46b.TransZ = 5
	Sw46.Timerenabled = False
	Sw46.Timerenabled = True
End Sub
Sub	Sw46_Timer
	Sw46.Timerenabled = False
	Sw46a.TransZ = 0
	Sw46b.TransZ = 0
End Sub


' Left Ramp
Sub Sw36_Hit():Controller.Switch(36) = 1:End Sub 
Sub Sw36_UnHit()
	Controller.Switch(36) = 0
	If ActiveBall.VelY < 0 Then
		PlaySound"comet_ramp1"
	Else
		StopSound"comet_ramp1"
	End If
End Sub

Sub Sw37_Hit(): PlaySound"cyclone_ramp_enter": Controller.Switch(37)=1:LEDFlash: End Sub
Sub Sw37_UnHit(): Controller.Switch(37)=0: End Sub

Sub LRHelp_Hit():stopsound "comet_ramp1" :Playsound "fx_metalrolling": ActiveBall.VelY=ActiveBall.VelY/8: ActiveBall.VelX=ActiveBall.VelX/8: End Sub

' Right Ramp
Sub Sw40_Hit(): Controller.Switch(40)=1: ActiveBall.VelY=ActiveBall.VelY/5: ActiveBall.VelX=ActiveBall.VelX/5: End Sub
Sub Sw40_unHit():Controller.Switch(40)=0:End Sub

Sub Sw40b_Hit():Playsound "fx_metalrolling": End Sub







Sub Sw47_Hit(): PlaySound "fx_kicker-enter": vpmTimer.PulseSw 47: Sw48.enabled = 0: End Sub
Sub sw47b_Hit(): PlaySound "fx_plasticrolling": End Sub

Sub Sw48_hit:PlaySound "fx_kicker-enter": Sw48.enabled = 0 : End Sub
Sub Sw48b_UnHit: vpmtimer.addtimer 500, "Sw48.enabled = 1 '" :End Sub



Sub Sw49_Hit:vpmTimer.PulseSw 49:PlaySound "bump": End Sub

Sub Sw50_Hit():Playsound "fx_sensor": Controller.Switch(50)=1: End Sub
Sub Sw50_UnHit():Controller.Switch(50)=0: End Sub

Sub Sw51T_Hit():Playsound "fx_kicker-enter": End Sub
Sub Sw51T_UnHit():Playsound "fx_metalrolling": End Sub

Sub Sw51Ta_Hit():Stopsound "fx_metalrolling":vpmtimer.addtimer 400, "BallDropFX '": End Sub


'***************************************************
'       JP's VP10 Fading Lamps & Flashers
'       Based on PD's Fading Light System
' SetLamp 0 is Off
' SetLamp 1 is On
' fading for non opacity objects is 4 steps
'***************************************************
 
Dim bulb
Dim LampState(200), FadingLevel(200)
Dim FlashSpeedUp(200), FlashSpeedDown(200), FlashMin(200), FlashMax(200), FlashLevel(200)
 
InitLamps()             ' turn off the lights and flashers and reset them to the default parameters
LampTimer.Interval = 10 'lamp fading speed
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
 
Sub InitLamps()
    Dim x
    For x = 0 to 200
        LampState(x) = 0         ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4       ' used to track the fading state
        FlashSpeedUp(x) = 0.5    ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.35 ' slower speed when turning off the flasher
        FlashMax(x) = 1          ' the maximum value when on, usually 1
        FlashMin(x) = 0          ' the minimum value when off, usually 0
        FlashLevel(x) = 0        ' the intensity of the flashers, usually from 0 to 1
    Next
End Sub
 
Sub UpdateLamps

	NFadeL 1, Light1
	NFadeL 2, Light2
	NFadeL 3, Light3
	NFadeL 4, Light4
	NFadeL 5, Light5
	NFadeL 6, Light6
	NFadeL 7, Light7
	NFadeL 8, Light8
	NFadeL 9, Light9
	NFadeL 10, Light10
	NFadeL 11, Light11
	NFadeL 12, Light12
	NFadeL 13, Light12
	NFadeL 14, Light14
	NFadeL 15, Light15
	NFadeL 16, Light16

	NFadeL 20, Light20
	NFadeL 21, Light21
	NFadeL 22, Light22
	NFadeL 23, Light23
	NFadeL 24, Light24
	Flash 25, Light25
	Flash 26, Light26
	Flash 27, Light27
	Flash 28, Light28
	Flash 29, Light29
	Flash 30, Light30
	NFadeL 32, Light32
	NFadeL 33, Light33
	NFadeL 34, Light34
	NFadeL 35, Light35
	NFadeL 36, Light36
	NFadeL 37, Light37

    'Skull Eyes
    FadeObjm 39, EyeLP, "SkullEyesMap_ON3", "SkullEyesMap_ON2", "SkullEyesMap_ON1", "SkullEyesMap"
	Flash 39, EyeL
    FadeObjm 38, EyeRP, "SkullEyesMap_ON3", "SkullEyesMap_ON2", "SkullEyesMap_ON1", "SkullEyesMap"
	Flash 38, EyeR

	Flash 40, Light40
	NFadeL 41, Light41
	NFadeL 42, Light42
	NFadeL 43, Light43
	NFadeL 44, Light44
	NFadeL 45, Light45
	NFadeL 46, Light46
	NFadeL 47, Light47
	NFadeL 48, Light48
	NFadeL 49, Light49
	NFadeL 50, Light50
	NFadeL 51, Light51
	NFadeL 52, Light52
	Flash 53, Light53

	NFadeL 54, Light54
	NFadeL 55, Light55
	NFadeL 56, Light56

	NFadeL 57, Light57
	NFadeL 58, Light58
	NFadeL 59, Light59
	NFadeL 60, Light60
	NFadeL 61, Light61
	NFadeL 62, Light62
	NFadeL 63, Light63
	NFadeL 64, Light64

    'Flashers	
	NFadeL 92, f2b1
	NFadeLm 92, f2b2
    
 	NFadeL 95, f5b

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
        Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1            'wait
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
        Case 4:object.image = b:FadingLevel(nr) = 0
        Case 5:object.image = a:FadingLevel(nr) = 1
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
 
' RGB Leds
 
Sub RGBLED (object,red,green,blue)
    If TypeName(object) = "Light" Then
        object.color = RGB(0,0,0)
        object.colorfull = RGB(2.5*red,2.5*green,2.5*blue)
        object.state=1
    ElseIf TypeName(object) = "Flasher" Then
        object.color = RGB(2.5*red,2.5*green,2.5*blue)
        object.IntensityScale = 1
    End If
End Sub
 
' Modulated Flasher and Lights objects
 
Sub SetLampMod(nr, value)
    If value > 0 Then
        LampState(nr) = 1
    Else
        LampState(nr) = 0
    End If
    FadingLevel(nr) = value
End Sub
 
Sub LampMod(nr, object)
    Object.IntensityScale = FadingLevel(nr)/128
    If TypeName(object) = "Light" Then
        Object.State = LampState(nr)
    End If
    If TypeName(object) = "Flasher" Then
        Object.visible = LampState(nr)
    End If
End Sub
 
 









'**********
'LED Rope
'**********


Dim WindLamp1
Sub CheckWindCoasterLamps (Lamp, State)
	Select Case Lamp
		Case 1:
			If State = 1 then 'Blinking
				Light1.TimerEnabled=0: Light1.TimerInterval = 500: Light1.TimerEnabled=1: 
				WindLamp1=1
			Else
				Light1.TimerEnabled=0: Light1.TimerInterval = 500: Light1.TimerEnabled=1
			End If
	End Select
End Sub


Sub Light1_Timer
	Light1.TimerEnabled = 0: 
	WindLamp1 = 0 'Not blinking
End Sub 




Sub LEDOFF
	For each yy in aLedsLightsA: yy.state = 0: next
	For each yy in aLedsLightsB: yy.state = 0: next
	For each yy in aLedsLightsC: yy.state = 0: next
    Backdrop1.image = "BackBrop_offoff"
    SetLamp 195, 0
    SetLamp 196, 0
    SetLamp 197, 0
End Sub

Sub LEDON
	For each yy in aLedsLightsA: yy.state = 1: next
	For each yy in aLedsLightsB: yy.state = 1: next
	For each yy in aLedsLightsC: yy.state = 1: next
    Backdrop1.image = "BackBrop_off"
    SetLamp 195, 1
    SetLamp 196, 1
    SetLamp 197, 1
End Sub

Dim LEDcount,yy
Sub LEDCW
	LEDcount = (LEDcount + 1) mod 3
	Select Case LEDcount
		Case 0:
	        For each yy in aLedsLightsA: yy.state = 1: next
	        For each yy in aLedsLightsB: yy.state = 0: next
	        For each yy in aLedsLightsC: yy.state = 0: next
            Backdrop1.image = "BackBrop_3"
            SetLamp 195, 1
            SetLamp 196, 0
		Case 1:
	        For each yy in aLedsLightsA: yy.state = 0: next
	        For each yy in aLedsLightsB: yy.state = 1: next
	        For each yy in aLedsLightsC: yy.state = 0: next
            Backdrop1.image = "BackBrop_2"
            SetLamp 195, 0
            SetLamp 196, 1

		Case 2:
	        For each yy in aLedsLightsA: yy.state = 0: next
	        For each yy in aLedsLightsB: yy.state = 0: next
	        For each yy in aLedsLightsC: yy.state = 1: next
            Backdrop1.image = "BackBrop_1"
            SetLamp 195, 0
            SetLamp 196, 0

	End Select
End Sub
Sub LEDCCW
	LEDcount = (LEDcount - 1)
	If LEDcount < 0 then 
		LEDcount = (LEDcount+3) mod 3
	Else
		LEDcount = LEDcount mod 3
	End If

	Select Case LEDcount
		Case 0:
	        For each yy in aLedsLightsA: yy.state = 1: next
	        For each yy in aLedsLightsB: yy.state = 0: next
	        For each yy in aLedsLightsC: yy.state = 0: next
            Backdrop1.image = "BackBrop_3"
            SetLamp 195, 1
            SetLamp 196, 0

		Case 1:
	        For each yy in aLedsLightsA: yy.state = 0: next
	        For each yy in aLedsLightsB: yy.state = 1: next
	        For each yy in aLedsLightsC: yy.state = 0: next
            Backdrop1.image = "BackBrop_2"
            SetLamp 195, 0
            SetLamp 196, 1

		Case 2:
	        For each yy in aLedsLightsA: yy.state = 0: next
	        For each yy in aLedsLightsB: yy.state = 0: next
	        For each yy in aLedsLightsC: yy.state = 1: next
            Backdrop1.image = "BackBrop_1"
            SetLamp 195, 1
            SetLamp 196, 0
	End Select
End Sub

Sub LEDFlash
	LEDTimer.Enabled = 0
	LEDOFF
	LEDFlashTimer.Enabled = 1
End Sub

Sub LEDStop
	LEDTimer.Enabled = 0
	LEDOFF
End Sub

Sub LEDStart
	LEDTimer.Enabled = 1
End Sub

LEDTimer.Interval = 500
LEDTimer.Enabled = 0
Dim LEDTimerCount
Sub LEDTimer_Timer
	LEDTimerCount = LEDTimerCount + 1
	If LEDTimerCount < 20 or Light1BlinkingAndBallInPlay Then
		LEDCW
	Else
		LEDCCW
	End If
End Sub

LEDFlashTimer.Interval = 50
LEDFlashTimer.Enabled = 0
Dim LEDTimerFlashCount
Sub LEDFlashTimer_Timer
	LEDTimerFlashCount = LEDTimerFlashCount + 1
	If LEDTimerFlashCount < 31 Then
		If LEDTimerFlashCount Mod 2 = 0 Then
			LEDOFF
             SetLamp 197, 0
		Else
			LEDON
             SetLamp 197, 1
		End If
	Else
		LEDTimer.Enabled = 1
		LEDFlashTimer.Enabled = 0
		LEDTimerFlashCount = 0
		LEDTimerCount = 0
	End If
End Sub

Function Light1BlinkingAndBallInPlay
	If WindLamp1 = 1 and Controller.Switch(11) = 0 Then
		Light1BlinkingAndBallInPlay = 1
	End If
End Function

Option Explicit

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' No special SSF tweaks yet.


On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error GoTo 0

Dim UseVPMDMD: UseVPMDMD = Table1.ShowDT
Scoretext.visible = UseVPMDMD

LoadVPM "01120100", "gts3.VBS", 3.02


'#########################################
'Render the Mars Bulb light that sits on top of the real machine's
'bacbox over the top 1/4 of the playfield using a rotating primitive.
'#########################################

Const marson = 1

'######################################################

'==================================================================
' TODO:: Game Specific code starts here
'==================================================================

Const UseSolenoids = 2
Const UseLamps = True
Const UseSync = True
' Standard Sounds used by Driver help code
Const SSolenoidOn = "solon"
Const SSolenoidOff = "soloff"
Const SFlipperOn = "FlipperUp"
Const SFlipperOff = "FlipperDown"
Const SCoin = "coin3"    'Coin inserted

'switches
'switch 0-7 handled by pinmame
'8-9 not used
Const swLeftCoinShute = 0
Const swRightCoinShute = 1
Const swCenterCoinChute = 2
Const swCoinChute = 3
Const swStartButton = 4
Const swTournament = 5
Const swFrontDoor = 6

'magnet under heli
'picks up the ball from the lower elevator
'ball travels with the heli and droppes the ball in a ramp wich leads to the lower left flipper
'or ball can be dropped at 2 places (popbumper or hole just before the rampdrop area
Const sElectroMagnet = 13

Const sBallLiftMotor = 22
Const sHeliBladeMotor = 23
Const sHeliARMmotorRelay = 24
Const sHeliDirectionalRelay = 25

Const sLighBoxrelay = 26
Const sTicketCoinMeterEnabled = 27
Const sBallRelease = 28
Const sOutHole = 29
Const sKnocker = 30
Const stiltrelay = 31
Const sGameOverRelay = 32



'sol mapping

'SolCallback(sLLFlipper)		= "SolFlipper LeftFlipper,Nothing,"
'SolCallback(sLRFlipper) 	= "SolFlipper RightFlipper,RightFlipper1,"

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySound SoundFx("FlipperUp", DOFContactors): LeftFlipper.RotateToEnd
    Else
        PlaySound SoundFx("FlipperDown", DOFContactors): LeftFlipper.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySound SoundFx("FlipperUp", DOFContactors): RightFlipper.RotateToEnd : RightFlipper1.RotateToEnd
    Else
        PlaySound SoundFx("FlipperDown", DOFContactors): RightFlipper.RotateToStart : RightFlipper1.RotateToStart
    End If
End Sub

SolCallback(1) = "vpmSolSound ""jet1"","
SolCallback(2) = "vpmSolSound ""jet1"","
SolCallback(3) = "vpmSolSound ""Jet1"","
SolCallback(4) = "vpmSolSound ""lSling"","
SolCallback(5) = "vpmSolSound ""lSling"","
SolCallback(6) = "vpmSolSound ""lSling"","
SolCallback(7) = "vpmSolSound ""lSling"","
SolCallback(8) = "dtLDrop.SolDropUp"
SolCallback(9) = "BallLeftShooter"
'SolCallback(10) = "bsVUK.SolOut"
SolCallback(10) = "VUK_Out"
SolCallback(11) = "Tpost"
SolCallback(12) = "vlLock.SolExit"
SolCallback(13) = "ElectroMagnet"
SolCallback(14) = "Sol14"
SolCallBack(15) = "Sol15"
SolCallBack(16) = "Sol16"
SolCallBack(17) = "Sol17"
SolCallBack(18) = "Sol18"
SolCallBack(19) = "Sol19"
SolCallBack(20) = "Sol20"
SolCallback(21) = "MBulb"
SolCallback(22) = "BallLiftMotor" '							Do not use these - they are in mechanics handlers
SolCallback(23) = "Blades_Motor" '              					Do not use these - they are in mechanics handlers
SolCallback(24) = "HeliArm" '					Do not use these - they are in mechanics handlers
SolCallback(25) = "HeliArmRelay" '                   Do not use these - they are in mechanics handlers
'SolCallback(26)  = "GI_Update"
'SolCallback(27)  = "vpmSolSound ""solon"","
SolCallback(28) = "bsTrough.SolOut"
SolCallback(29) = "bsSewer.SolOut"
SolCallback(30) = "vpmSolSound SoundFX(""knocker"",DOFKnocker),"
SolCallback(31) = "SolTilt"
SolCallback(32) = "GameOn"


Dim gon, ton, xx

gon = 0
ton = 0

Sub GameOn(enabled)
    If enabled Then
        gon = 1
        xgon.state = 1
        playsound "SolOn"
    For Each xx In GI : xx.State = 1 : Next
    Else
        gon = 0
        xgon.state = 0
        For Each xx In GI : xx.State = 0 : Next
    End If
End Sub

Set LampCallback = GetRef("UpdateMultipleLamps")

Sub UpdateMultipleLamps()
    Light75a.State = Light75.State
    Light83a.State = Light83.State
    F124.State = Light124.State
    F125.State = Light125.State
Light54a.State = Light54.State
Light55a.State = Light55.State
Light56a.State = Light56.State
Light57a.State = Light57.State
Light80a.State = Light80.State
Light81a.State = Light81.State
Light82a.State = Light82.State
End Sub

Sub SolTilt(enabled)
    If enabled Then
        TTilt.enabled = 1
    End If
End Sub

Sub TTilt_Timer()
    If gon = 1 Then
        ton = 0
    Else
        ton = 1
        XTO.State = 1
        Me.enabled = 0
    End If
End Sub

Sub CaptiveTrigger1_Hit() : cbCaptive.TrigHit ActiveBall                      : End Sub
Sub CaptiveTrigger1_UnHit() : cbCaptive.TrigHit 0                               : End Sub
Sub CaptiveWall_Hit() : cbCaptive.BallHit ActiveBall : playsound "fx_collide" : End Sub
Sub Captive2_Hit() : cbCaptive.BallReturn Me                           : End Sub
Sub Captive3_Hit() : cbCaptive.BallReturn Me                           : End Sub

Const BuyInButton = 3

ExtraKeyHelp = KeyName(BuyInButton) & vbTab & "BuyInButton"

' -------------------------------------
' Lights Array
' -------------------------------------
' lights, map all lights on the playfield to the Lights array
On Error Resume Next
Dim i
For i = 0 To 127
Execute "Set Lights(" & i & ")  = Light" & i
Next

Dim bsTrough, bsSewer, cbCaptive
Dim bsVUK
Dim dtlDrop, k
Dim vllock

Sub Table1_Init()
    Controller.GameName = "rescu911"                ' Use specified ROM set.
    Controller.SplashInfoLine = "Rescue 911"
    Controller.HandleKeyboard = 0
    Controller.ShowTitle = 0
    Controller.ShowDMDOnly = 1
    Controller.ShowFrame = 0
    Controller.HandleMechanics = 0
	Controller.Hidden = UseVPMDMD
    Controller.Games("rescu911").Settings.Value("rol") = 0
    On Error Resume Next
    Controller.Run
    If Err Then MsgBox Err.Description
    On Error GoTo 0

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1
    vpmNudge.TiltSwitch = 151
    vpmNudge.Sensitivity = 5

	set bsTrough=New cvpmBallStack
     bsTrough.InitSw 0, 33, 0, 0, 0, 0, 0, 0
     bsTrough.InitKick BallRelease, 110, 8
     bsTrough.InitEntrySnd "solenoid", "solOn"
     bsTrough.InitExitSnd SoundFX("ballrel", DOFContactors), SoundFX("solOn", DOFContactors)
     bsTrough.Balls = 3

    Set bsSewer = New cvpmBallStack
	With bsSewer
        .InitSw 0, 23, 0, 0, 0, 0, 0, 0
        .InitKick KUT, 180, 5
        '.InitExitSnd SoundFX("BallRelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
    End With

    bsSewer.AddBall 0
    Controller.Switch(34) = 1
    PUBlock.isdropped = 1

	set bsVUK = New cvpmBallStack
	bsVUK.InitSw 0, 91, 0, 0, 0, 0, 0, 0
    bsVUK.InitKick Kicker_VUK, 139, 3
    bsVUK.InitEntrySnd "solenoid", "solenoid"
    bsVUK.InitExitSnd SoundFX("ballrel", DOFContactors), SoundFX("solenoid", DOFContactors)

   	set dtLDrop = New cvpmDropTarget
	dtlDrop.InitDrop Array(sw20, sw30), Array(20, 30)
    dtlDrop.InitSnd SoundFX("target_drop", DOFContactors), SoundFX("flapopen", DOFContactors)

	Set cbCaptive = New cvpmCaptiveBall : With cbCaptive
        .InitCaptive CaptiveTrigger1, CaptiveWall, Array(Captive1, Captive2, Captive3), 360
        .Start
    Captive1.CreateSizedBall(23).Image = "JPBall-Dark2"
    Captive3.CreateSizedBall(23).Image = "JPBall-Dark2"
    .NailedBalls = 2
    .ForceTrans = 0.2
    '.MinForce = 3.5
    End With


    'Visible Lock
    Set vlLock = New cvpmVLock
    With vllock
        .InitVLock Array(RL1, RL2), Array(RLK1, RLK2), Array(0, 0)
        .InitSnd SoundFX("ballrel", DOFContactors), SoundFX("solenoid", DOFContactors)
        .CreateEvents "vlLock"
    End With

    SW15a.isdropped = 1
    SW16a.isdropped = 1

End Sub

If Table1.ShowDT = True Then
ramp15.Visible = True
ramp16.Visible = True
Else
ramp15.Visible = False
ramp16.Visible = False
End If


'Dim cball1,cball2,cball3

'Set cball1 = CB1.createball:cball1.image = "JPBall-Dark2"
'CB1.Kick 180,1

'Sub CB1_Unhit
'CB1.enabled = 0
'End Sub

'Set cball2 = CB2.createball:cball2.image = "JPBall-Dark2"
'CB2.Kick 180,1

'Sub CB2_Unhit
'CB2.enabled = 0
'End Sub

'Set cball3 = CB3.createball:cball3.image = "JPBall-Dark2"
'CB3.Kick 180,1

'Sub CB3_Unhit
'CB3.enabled = 0
'End Sub

Sub Table1_KeyDown(ByVal keycode)
    If keycode = LeftFlipperKey Then Controller.Switch(82) = 1
    If keycode = RightFlipperkey Then Controller.Switch(83) = 1
    If vpmKeyDown(keycode) Then Exit Sub
    If keycode = PlungerKey Then Plunger.Pullback : PlaySound "plungerpull"
    If keycode = 16 Then plunger1.fire
    If keycode = 17 Then plunger1.pullback
    If keycode = 3 Then Kicker1.createball : Kicker1.kick 45, 1
End Sub

Sub Table1_KeyUp(ByVal keycode)
    If keycode = LeftFlipperKey Then Controller.Switch(82) = 0
    If keycode = RightFlipperkey Then Controller.Switch(83) = 0
    If vpmKeyUp(keycode) Then Exit Sub
    If keycode = PlungerKey Then Plunger.Fire : PlaySound "plunger"
End Sub

Sub Bumper1_Hit()
    PlaySound SoundFX("fx_bumper4", DOFContactors)
    'B1L1.State = 1:B1L2. State = 1
    vpmTimer.PulseSw 12
    Me.TimerEnabled = 1
End Sub

Sub Bumper1_Timer()
    'B1L1.State = 0:B1L2. State = 0
    Me.Timerenabled = 0
End Sub

Sub Bumper2_Hit()
    PlaySound SoundFX("fx_bumper4", DOFContactors)
    'B2L1.State = 1:B2L2. State = 1
    vpmTimer.PulseSw 10
    Me.TimerEnabled = 1
End Sub

Sub Bumper2_Timer()
    'B2L1.State = 0:B2L2. State = 0
    Me.Timerenabled = 0
End Sub

Sub Bumper3_Hit()
    PlaySound SoundFX("fx_bumper4", DOFContactors)
    'B3L1.State = 1:B3L2. State = 1
    vpmTimer.PulseSw 11
    Me.TimerEnabled = 1
End Sub

Sub Bumper3_Timer()
    'B3L1.State = 0:B3L2. State = 0
    Me.Timerenabled = 0
End Sub

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot()
    PlaySound SoundFX("left_slingshot", DOFContactors), 0, 1, 0.05, 0.05
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
    RightSlingShot.TimerInterval = 10
    vpmTimer.PulseSw 14
End Sub

Sub RightSlingShot_Timer()
    Select Case RStep
        Case 3 : RSLing1.Visible = 0 : RSLing2.Visible = 1 : sling1.TransZ = -10
        Case 4 : RSLing2.Visible = 0 : RSLing.Visible = 1 : sling1.TransZ = 0 : RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot()
    PlaySound SoundFX("right_slingshot", DOFContactors), 0, 1, -0.05, 0.05
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    Lstep = 0
    LeftSlingShot.TimerEnabled = 1
    LeftSlingShot.TimerInterval = 10
    'gi3.State = 0:Gi4.State = 0
    vpmTimer.PulseSw 13
End Sub

Sub LeftSlingShot_Timer()
    Select Case Lstep
        Case 3 : LSLing1.Visible = 0 : LSLing2.Visible = 1 : sling2.TransZ = -10
        Case 4 : LSLing2.Visible = 0 : LSLing.Visible = 1 : sling2.TransZ = 0 : LeftSlingShot.TimerEnabled = 0
    End Select
    Lstep = Lstep + 1
End Sub

Sub Pins_Hit(idx)
    PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit(idx)
    PlaySound SoundFX("target", DOFContactors), 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit(idx)
    PlaySound "metalhit_thin", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit(idx)
    PlaySound "metalhit_medium", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit(idx)
    PlaySound "metalhit2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit(idx)
    PlaySound "gate4", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Rolls_Hit(idx)
    PlaySound "rollover"
End Sub

Sub Rubbers_Hit(idx)
    Dim finalspeed
    finalspeed = SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
    If finalspeed > 20 Then
        PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    End If
    If finalspeed >= 6 And finalspeed <= 20 Then
        RandomSoundRubber()
    End If
End Sub

Sub Posts_Hit(idx)
    Dim finalspeed
    finalspeed = SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
    If finalspeed > 16 Then
        PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    End If
    If finalspeed >= 6 And finalspeed <= 16 Then
        RandomSoundRubber()
    End If
End Sub

Sub RandomSoundRubber()
    Select Case Int(Rnd * 3) + 1
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
    Select Case Int(Rnd * 3) + 1
        Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
        Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
        Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    End Select
End Sub

Dim x, y, a, r, bfx, bfy

a = 270 ' 4.68
r = 400

Dim hspeed : hspeed = 0.714 / 16.6
Dim hrotspeed : hrotspeed = 3.0 / 16.6
Dim bladespeed : bladespeed = 1
Dim hinitialrot : hinitialrot = 0
Dim hballr : hballr = 100
Dim cmechpos, helihome, heliaway
cmechpos = 1

Sub heli_animation_timer()

    Select Case cmechpos
        Case 1 ' Take off
			PUBlock.isdropped = 0
            If hinitialrot < 180 Then
				 hinitialrot = hinitialrot + hrotspeed
                 Chopper2.ObjRotz = Chopper2.ObjRotz + hrotspeed
				 ChopperS.ObjRotz = ChopperS.ObjRotz + hrotspeed
            End If
            If a > 270+178 Then
                cmechpos = 3
			Else
				a = a + hspeed
				x = Cos(a * 6.28318 / 360) * r + 190
				y = Sin(a * 6.28318 / 360) * r + 580
				PrimBlades.X = x
				PrimBlades.Y = y
                PrimBlades1.X = x
				PrimBlades1.Y = y
				Heli_Hub.ObjRotZ = Heli_Hub.ObjRotZ + hspeed
				Heli_Shaft.ObjRotz = Heli_Shaft.ObjRotz + hspeed
				Heli_Base.ObjRotz = Heli_Base.ObjRotz + hspeed
				Heli_Wire.ObjRotz = Heli_Wire.ObjRotz + hspeed
				Chopper2.ObjRotz = Chopper2.ObjRotz + hspeed
				ChopperS.ObjRotz = ChopperS.ObjRotz + hspeed
				Chopper2.X = x
				Chopper2.Y = y
                ChopperS.X = x
				ChopperS.Y = y
            End If
        Case 3 ' Waiting for pickup
            If mdirection = 0 Then
                cmechpos = 4
				Lower_Post.enabled = 1
				Heli_Kick.enabled = 1
            End If
        Case 4 ' Leaving with ball
            if hinitialrot > 0 then
				Chopper2.ObjRotz = Chopper2.ObjRotz - hrotspeed
                ChopperS.ObjRotz = ChopperS.ObjRotz - hrotspeed
				hinitialrot = hinitialrot - hrotspeed
			end if
            r = 405
            If a < 270 Then
                cmechpos = 6
			Else
				a = a - hspeed
				x = Cos(a * 6.28318 / 360) * r + 190
				y = Sin(a * 6.28318 / 360) * r + 580
				PrimBlades.X = x
				PrimBlades.Y = y
				PrimBlades1.X = x
				PrimBlades1.Y = y
				Heli_Hub.ObjRotZ = Heli_Hub.ObjRotZ - hspeed
				Heli_Shaft.ObjRotz = Heli_Shaft.ObjRotz - hspeed
				Heli_Base.ObjRotz = Heli_Base.ObjRotz - hspeed
				Heli_Wire.ObjRotz = Heli_Wire.ObjRotz - hspeed
				Chopper2.ObjRotz = Chopper2.ObjRotz - hspeed
                ChopperS.ObjRotz = ChopperS.ObjRotz - hspeed
				Chopper2.X = x
				Chopper2.Y = y
				ChopperS.X = x
				ChopperS.Y = y
				If b3active = 1 And hbout = 0 Then
					bfx = Cos((Chopper2.ObjRotz+100) * 6.28318 / 360) * hballr + x
					bfy = Sin((Chopper2.ObjRotz+100) * 6.28318 / 360) * hballr + y
					cball.X = bfx
					cball.Y = bfy
				End If
            End If
        Case 6 ' Finished.  Reset
            Chopper2.ObjRotZ = 80
            Chopper2.visible = 1
            ChopperS.ObjRotZ = 80
            ChopperS.visible = 1
            PrimBlades.X = 183.25
            PrimBlades.Y = 183.25
            PrimBlades1.X = 183.25
			PrimBlades1.Y = 183.25
            Heli_Shaft.ObjRotz = 0
            Heli_Hub.ObjRotz = 0
            Heli_Base.ObjRotz = 0
            Heli_Wire.ObjRotz = 0
			hinitialrot = 0
            cmechpos = 1
            a = 270 '4.68
            r = 400

            If emon = 1 Then
                PlaySound SoundFX("popper_ball", DOFContactors)
				PlaySound "subway"
				Heli_Raise.destroyball
                Heli_Raise.enabled = 1
                Heli_Kick.enabled = 0
                Kicker6.createball
                Kicker6.kick 180, 2
				emon = 0
            End If
            hbout = 1
            b3active = 1
			me.enabled = 0
    End Select
End Sub


Dim b3active, BOut
b3active = 1

Dim br, ba, by, bx

br = 100
ba = 3.38

Sub Heli_Opto_Watch_Timer()
    If Controller.Switch(80) = True Then
        L80.State = 1
    Else
        L80.State = 0
    End If

    If Controller.Switch(34) = True Then
        BLD.State = 1
    Else
        BLD.State = 0
    End If

    If Controller.Switch(24) = True Then
        BLU.State = 1
    Else
        BLU.State = 0
    End If

    If Heli_Base.ObjRotz >= 0 And Heli_Base.ObjRotz <= 2 Then
        Controller.Switch(100) = 1
        helihome = 1
        heliaway = 0
    Else
        Controller.Switch(100) = 0
    End If
    If Heli_Base.ObjRotz >= 178 And Heli_Base.ObjRotz <= 190 Then
        Controller.Switch(90) = 1
        heliaway = 1
        helihome = 0
    Else
        Controller.Switch(90) = 0
    End If
End Sub


'********************************************

Sub Lower_Post_Timer()
    If Ball_Raise.Z <= -170 Then
		drop_pause.enabled = 1
        Controller.Switch(34) = 1
        Me.enabled = 0
    End If
    If Ball_Raise.Z <= -110 Then
        Controller.Switch(24) = 0
    End If
    Ball_Raise.Z = Ball_Raise.Z - 10
End Sub

Sub drop_pause_timer()
	Ball_Raise.Z = -170
    PUBlock.isdropped = 1
	Heli_Kick.Enabled = 1
    Me.enabled = 0
End Sub

Dim cball

Sub Raise_Post_Timer()
    If cball.z >= 80 Then
        Me.enabled = 0
        Controller.Switch(24) = 1
        hbin = 0
        hbout = 0
    End If
    If cball.z >= 10 Then
        Controller.Switch(80) = 0
        Controller.Switch(34) = 0
    End If
    cball.z = cball.z + 0.31
    Ball_Raise.Z = Ball_Raise.Z + 0.31
End Sub

Sub Blades_Spin_Timer()
    PrimBlades.ObjRotz = SystemTime Mod 360
    PrimBlades1.ObjRotz = SystemTime Mod 360
End Sub


'********************************************

Sub HeliArm(enabled)
    If enabled Then
		heli_animation.enabled = 1
    End If
End Sub

Dim mdirection

Sub HeliArmRelay(enabled)
    If enabled Then
        mdirection = 0
        MF.State = 0
        MB.State = 1
    Else
        MB.State = 0
        MF.State = 1
        mdirection = 1
    End If
End Sub

Sub Blades_Motor(enabled)
    If enabled Then
        Blades_Spin.enabled = 1
    Else
        Blades_Spin.enabled = 0
    End If
End Sub

Sub BallLeftShooter(Enabled)
    If Enabled Then
        Heli_Kick.Kick 0, 28
            PlaySound SoundFX("popper_ball", DOFContactors)
            If hbin = 1 Then
            PlaySound SoundFX("popper_ball", DOFContactors)
            Heli_Raise.Kick 0, 28
            End If
        Controller.Switch(80) = 0
        hbin = 0
    End If
End Sub

Sub BallLiftMotor(Enabled)
    If Enabled And hbin = 1 Then
        Heli_Raise.DestroyBall
		Set cball = Heli_Raise.createball
		Raise_Post.enabled = 1
        Heli_Raise.enabled = 0
    End If
End Sub

Dim emon

'Sub ElectroMagnet(enabled)
'If enabled Then
'emon = 1
'Light101.State = 1
'else
'Light101.State = 0
'End If
'If NOT enabled AND Controller.Switch(100) = 0 AND emon = 1 Then
'b3active = 0
'Heli_Raise.Kick 0,0
'hbout = 1
'BOut = 1
'emon = 0
'Heli_Raise.enabled = 1
'Heli_Kick.enabled = 0
'End If
'End Sub

Sub ElectroMagnet(enabled)
    If enabled Then
        emon = 1
        Light101.State = 1
    Else
        If Controller.Switch(100) = 0 And emon = 1 Then
            Light101.State = 0
            b3active = 0
            Heli_Raise.Kick 0, 0
hbout = 1
            BOut = 1
            emon = 0
            Heli_Raise.enabled = 1
            Heli_Kick.enabled = 0
        End If
    End If
    'If NOT enabled AND Controller.Switch(100) = 0 AND emon = 1 Then
    'b3active = 0
    'Heli_Raise.Kick 0,0
    'hbout = 1
    'BOut = 1
    'emon = 0
    'Heli_Raise.enabled = 1
    'Heli_Kick.enabled = 0
    'End If
End Sub

Dim vball

Sub VUK_Out(enabled)
    If enabled Then
Set vball = Kicker_VE.createball
vraise.enabled = 1
        Controller.Switch(91) = 0
    End If
End Sub

Sub Vraise_Timer()
    If vball.Z >= 173 Then
        Kicker_VE.kick 0, 0
bsVUK.exitsol_on
        Me.enabled = 0
    End If
    vball.Z = vball.Z + 1
End Sub

Sub VDES_Hit()
    vball.destroyball
End Sub

Sub Flipper_Logo_Timer()
    UpdateFlipperLogos()
BallShadowUpdate
End Sub

Sub UpdateFlipperLogos()
    FLogoSx.RotAndTra2 = LeftFlipper.CurrentAngle
    FLogoDx.RotAndTra2 = RightFlipper.CurrentAngle
    FLogoDx1.RotAndTra2 = RightFlipper1.CurrentAngle
End Sub

Sub LockPost(enabled)
    If enabled Then
        PostLock.isdropped = 1
        PlaySound SoundFX("solon", DOFContactors)
Else
        PostLock.isdropped = 0
    End If
End Sub

Sub Tpost(enabled)
    If enabled Then
        TopPost.isdropped = 1
        PlaySound SoundFX("solon", DOFContactors)
Else
        TopPost.isdropped = 0
    End If
End Sub

Sub Sol14(enabled)
    If enabled Then
        F13.State = 1
        F13a.State = 1
        F13b.State = 1
        F13c.State = 1
    Else
        F13.State = 0
        F13a.State = 0
        F13b.State = 0
        F13c.State = 0
    End If
End Sub

Sub Sol15(enabled)
    If enabled Then
        F14.State = 1
        F14a.State = 1
        F14b.State = 1
        F14c.State = 1
    Else
        F14.State = 0
        F14a.State = 0
        F14b.State = 0
        F14c.State = 0
    End If
End Sub

Sub Sol16(enabled)
    If enabled Then
        F15.State = 1
        F15a.State = 1
    Else
        F15.State = 0
        F15a.State = 0
    End If
End Sub

Sub Sol17(enabled)
    If enabled Then
        F16.State = 1
        F16a.State = 1
    Else
        F16.State = 0
        F16a.State = 0
    End If
End Sub

Sub Sol18(enabled)
    If enabled Then
        F17a.State = 1
        F17.State = 1
    Else
        F17a.State = 0
        F17.State = 0
    End If
End Sub

Sub Sol19(enabled)
    If enabled Then
        F18.State = 1
    Else
        F18.State = 0
    End If
End Sub

Sub Sol20(enabled)
    If enabled Then
        F19.State = 1
    Else
        F19.State = 0
    End If
End Sub

'Switches

Sub SW15_Slingshot() : vpmTimer.PulseSw 15: SW15.IsDropped = 1 : SW15a.isdropped = 0 : PlaySound SoundFX("left_slingshot", DOFContactors): Me.timerenabled = 1 : End Sub
Sub SW15_Timer() : SW15.Isdropped = 0 : SW15a.isdropped = 1 : Me.timerenabled = 0 : End Sub

Sub SW16_Slingshot() : vpmTimer.PulseSw 16: SW16.IsDropped = 1 : SW16a.isdropped = 0 : PlaySound SoundFX("right_slingshot", DOFContactors): Me.timerenabled = 1 : End Sub
Sub SW16_Timer() : SW16.Isdropped = 0 : SW16a.isdropped = 1 : Me.timerenabled = 0 : End Sub

Sub t85_Hit() : PrimStandupTgtHit 85, T85, PrimT85: End Sub
Sub t85_Timer() : PrimStandupTgtMove 85, T85, PrimT85: End Sub

Sub t95_Hit() : PrimStandupTgtHit 95, T95, PrimT95: End Sub
Sub t95_Timer() : PrimStandupTgtMove 95, T95, PrimT95: End Sub

Sub t105_Hit() : PrimStandupTgtHit 105, T105, PrimT105: End Sub
Sub t105_Timer() : PrimStandupTgtMove 105, T105, PrimT105: End Sub

Sub t115_Hit() : PrimStandupTgtHit 115, T115, PrimT115: End Sub
Sub t115_Timer() : PrimStandupTgtMove 115, T115, PrimT115: End Sub

Sub t31_Hit() : PrimStandupTgtHit 31, T31, PrimT31: End Sub
Sub t31_Timer() : PrimStandupTgtMove 31, T31, PrimT31: End Sub

Sub t94_Hit() : PrimStandupTgtHit 94, T94, PrimT94: End Sub
Sub t94_Timer() : PrimStandupTgtMove 94, T94, PrimT94: End Sub

Sub t96_Hit() : PrimStandupTgtHit 96, T96, PrimT96: End Sub
Sub t96_Timer() : PrimStandupTgtMove 96, T96, PrimT96: End Sub

Sub t97_Hit() : PrimStandupTgtHit 97, T97, PrimT97: End Sub
Sub t97_Timer() : PrimStandupTgtMove 97, T97, PrimT97: End Sub

Sub t104_Hit() : PrimStandupTgtHit 104, T104, PrimT104: End Sub
Sub t104_Timer() : PrimStandupTgtMove 104, T104, PrimT104: End Sub

Sub t106_Hit() : PrimStandupTgtHit 106, T106, PrimT106: End Sub
Sub t106_Timer() : PrimStandupTgtMove 106, T106, PrimT106: End Sub

Sub t107_Hit() : PrimStandupTgtHit 107, T107, PrimT107: End Sub
Sub t107_Timer() : PrimStandupTgtMove 107, T107, PrimT107: End Sub

Sub t117_Hit() : PrimStandupTgtHit 117, T117, PrimT117: End Sub
Sub t117_Timer() : PrimStandupTgtMove 117, T117, PrimT117: End Sub

Sub t114_Hit() : PrimStandupTgtHit 114, T114, PrimT114: End Sub
Sub t114_Timer() : PrimStandupTgtMove 114, T114, PrimT114: End Sub

Sub t16_Hit() : PrimStandupTgtHit 16, T16, PrimT16: End Sub
Sub t16_Timer() : PrimStandupTgtMove 16, T16, PrimT16: End Sub

Sub SW111_Hit() : Controller.Switch(111) = 1 : End Sub
Sub SW111_UnHit() : Controller.Switch(111) = 0 : End Sub

Sub sw21_Hit() : Controller.Switch(21) = 1 : End Sub
Sub sw21_UnHit() : Controller.Switch(21) = 0 : End Sub

Sub sw22_Slingshot() : vpmTimer.PulseSw 22 : PlaySound SoundFX("right_slingshot", DOFContactors): End Sub
Sub sw32_Slingshot() : vpmTimer.PulseSw 32 : PlaySound SoundFX("left_slingshot", DOFContactors): End Sub
Sub swA2_Hit() : vpmTimer.PulseSw 102 : End Sub
Sub sw81_Hit() : vpmTimer.PulseSw 81 : PlaySound "ball_bounce": End Sub
Sub sw92_Hit() : vpmTimer.PulseSw 92 : End Sub
Sub sw93_Hit() : vpmTimer.PulseSw 93 : End Sub
Sub swA3_Hit() : vpmTimer.PulseSw 103 : End Sub
Sub sw113_Hit() : vpmTimer.PulseSw 113 : End Sub
Sub sw101_Hit() : Controller.Switch(101) = 1 : End Sub
Sub sw101_UnHit() : Controller.Switch(101) = 0 : End Sub

Sub LHD_Hit()
    PlaySound "ball_bounce"
End Sub

Sub RRD_Hit()
    PlaySound "ball_bounce"
End Sub

Sub TRD_Hit()
    PlaySound "ball_bounce"
End Sub

Sub RRD_Hit()
    PlaySound "ball_bounce"
End Sub

Sub LockExit_Hit()
    LockPost.IsDropped = 0
End Sub

Sub Kicker_VE_Hit()
    Playsound "Hole"
    bsVUK.AddBall Me
End Sub

Sub Lock_Drop_Hit()
    Playsound "Hole"
    bsVUK.AddBall Me
    vpmTimer.PulseSwitch(110),0,""
End Sub

Sub Heli_Mid_Drop_Hit()
    Me.destroyball
    'If helihome = 0 Then
    Playsound "Hole"
    bsVUK.AddBall Me
    vpmTimer.PulseSwitch(110),0,""
End Sub

Sub Trigger6_Hit()
    Me.destroyball
    PlaySound "popper_ball"
PlaySound "subway"
Kicker6.createball
    Kicker6.Kick 180, 2
End Sub

Sub drophalt_timer()
End Sub

Sub Outhole_Hit()
    If gon = 1 And ton = 1 Then
        bsSewer.AddBall Me
PlaySound "drain"
ton = 0
        XTO.State = 0
        Exit Sub
    End If
    If gon = 1 Then
        bsSewer.AddBall Me
PlaySound "drain"
Exit Sub
    End If
    If ton = 1 Then
        bsSewer.AddBall Me
PlaySound "drain"
ton = 0
        XTO.State = 0
        Exit Sub
    End If
    If gon = 0 Then
        bsTrough.AddBall Me
PlaySound "drain"
Exit Sub
    End If
End Sub

Sub KUTIN_hit() : bsTrough.Addball Me: End Sub

Sub sw20_Hit() : dtlDrop.Hit 1: sw20.IsDropped = 1 : End Sub
Sub sw30_Hit() : dtlDrop.Hit 2: sw30.Isdropped = 1 : End Sub

Sub Kicker3_Hit()
    Me.destroyball
    Kicker4.Createball
    Kicker4.kick 180, 1
End Sub

Sub SLT_Hit()
    PlaySound "gate"
Me.destroyball
    SLT1.Createball
    SLT1.Kick 180, 1
End Sub

Dim hbin, hbout

hbout = 1

Sub Heli_Raise_Hit()
    Controller.Switch(80) = 1
    PlaySound "kicker_enter"
hbin = 1
End Sub

Sub Heli_Kick_Hit()
    Controller.Switch(80) = 1
    PlaySound "kicker_enter"
End Sub

'Switch Handlers

Const WallPrefix = "T" 'Change this based on your naming convention
Const PrimitivePrefix = "PrimT" 'Change this based on your naming convention
Const PrimitiveBumperPrefix = "BumperRing" 'Change this based on your naming convention
Dim primCnt(120), primDir(120), primBmprDir(120)

'****************************************************************************
'***** Primitive Standup Target Animation
'****************************************************************************
'USAGE: 	Sub sw1_Hit: 	PrimStandupTgtHit  1, Sw1, PrimSw1: End Sub
'USAGE: 	Sub Sw1_Timer: 	PrimStandupTgtMove 1, Sw1, PrimSw1: End Sub

Const StandupTgtMovementDir = "TransX"
Const StandupTgtMovementMax = 6

Sub PrimStandupTgtHit(swnum, wallName, primName)
    PlaySound SoundFX("target", DOFContactors)
    vpmTimer.PulseSw swnum
    primCnt(swnum) = 0                                  'Reset count
    wallName.TimerInterval = 20     'Set timer interval
    wallName.TimerEnabled = 1   'Enable timer
End Sub

Sub PrimStandupTgtMove(swnum, wallName, primName)
    Select Case StandupTgtMovementDir
        Case "TransX"
            Select Case primCnt(swnum)
                Case 0 : primName.TransX = -StandupTgtMovementMax * 0.5
                Case 1 : primName.TransX = -StandupTgtMovementMax
                Case 2 : primName.TransX = -StandupTgtMovementMax * 0.5
                Case 3 : primName.TransX = 0
                Case Else : wallName.TimerEnabled = 0
            End Select
        Case "TransY"
            Select Case primCnt(swnum)
                Case 0 : primName.TransY = -StandupTgtMovementMax * 0.5
                Case 1 : primName.TransY = -StandupTgtMovementMax
                Case 2 : primName.TransY = -StandupTgtMovementMax * 0.5
                Case 3 : primName.TransY = 0
                Case Else : wallName.TimerEnabled = 0
            End Select
        Case "TransZ"
            Select Case primCnt(swnum)
                Case 0 : primName.TransZ = -StandupTgtMovementMax * 0.5
                Case 1 : primName.TransZ = -StandupTgtMovementMax
                Case 2 : primName.TransZ = -StandupTgtMovementMax * 0.5
                Case 3 : primName.TransZ = 0
                Case Else : wallName.TimerEnabled = 0
            End Select
    End Select
    primCnt(swnum) = primCnt(swnum) + 1
End Sub

Sub Mbulb(enabled)
    If enabled And marson = 1 Then
        Mars_Bulb.visible = 1
        SpinBulb.enabled = 1
    Else
        Mars_Bulb.visible = 0
        SpinBulb.enabled = 0
    End If
End Sub

Sub SpinBulb_Timer()
    Mars_Bulb.ObjRotZ = Mars_Bulb.ObjRotZ + 1
End Sub

Sub Table1_Exit()
    Controller.Stop
End Sub

Sub diag_timer()
    If Controller.Switch(100) = True Then
        Light102.State = 1
    Else
        Light102.State = 0
    End If
    If Controller.Switch(90) = True Then
        Light103.State = 1
    Else
        Light103.State = 0
    End If
End Sub

Sub Wall39_Unhit()
    PlaySound "ball_bounce"
End Sub

'*****************************************
'	Ball Shadow
'*****************************************

Dim BallShadow
BallShadow = Array (BallShadow1, BallShadow2, BallShadow3, BallShadow4, BallShadow5, BallShadow6, BallShadow7, BallShadow8)


Sub BallShadowUpdate()
    Dim BOT, b
    BOT = GetBalls

	' render the shadow for each ball
    For b = 0 to UBound(BOT)
		If BOT(b).X < table1.Width/2 Then
			BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (table1.Width/2))/7)) + 10
		Else
			BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (table1.Width/2))/7)) - 10
		End If

			BallShadow(b).Y = BOT(b).Y + 10
			BallShadow(b).Z = 1
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

Const tnob = 15 ' total number of balls
ReDim rolling(tnob)
InitRolling

Sub InitRolling()
    Dim i
    For i = 0 To tnob
        rolling(i) = False
    Next
End Sub

Sub RollingTimer_Timer()
    Dim BOT, b
    BOT = GetBalls

    ' stop the sound of deleted balls
    For b = UBound(BOT) + 1 To tnob
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
  If Table1.VersionMinor > 3 OR Table1.VersionMajor > 10 Then
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
  Else
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0
  End if
End Sub


' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


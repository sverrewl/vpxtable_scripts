'Flight 2000 (Stern 1980) for VPX by bord
'lock script courtesy of JP Salas

Option Explicit
Randomize

Const cGameName = "flight2k"  

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

LoadVPM "01550000", "stern.VBS", 3.26

Dim DesktopMode: DesktopMode = table1.ShowDT

'************************************************
'************************************************
'************************************************
'************************************************
'************************************************
Const UseSolenoids = True
Const UseLamps = True
Const UseGI = False
Const UseSync = False

' Standard Sounds
Const SSolenoidOn = "fx_solenoid"
Const SSolenoidOff = "fx_solenoidoff"
Const SCoin = "fx_coin"

Dim bsTrough, dtL, dtR, bump1, bump2, plLane1, plLane2, hiddenvalue

If DesktopMode = True Then 'Show Desktop components
CabinetRailLeft.visible=1
CabinetRailRight.visible=1
hiddenvalue=0
Else
CabinetRailLeft.visible=0
CabinetRailRight.visible=0
hiddenvalue=1
End if

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySound SoundFX("fx_flipperup",DOFContactors), 0, 1, -0.1, 0.25
        LeftFlipper.RotateToEnd
    Else
        PlaySound SoundFX("fx_flipperdown",DOFContactors), 0, 1, -0.1, 0.25
        LeftFlipper.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySound SoundFX("fx_flipperup",DOFContactors), 0, 1, 0.1, 0.25
        RightFlipper.RotateToEnd
   Else
        PlaySound SoundFX("fx_flipperdown",DOFContactors), 0, 1, 0.1, 0.25
        RightFlipper.RotateToStart
    End If
End Sub

Sub FlipperTimer_Timer
	'Add flipper, gate and spinner rotations here
	LFLogo1.RotY = LeftFlipper.CurrentAngle
	RFLogo1.RotY = RightFlipper.CurrentAngle
End Sub

Sub Table1_Init
     With Controller
         .GameName = cGameName
         If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
         .SplashInfoLine = "Flight 2000 (Stern 1980)"&chr(13)&"by bord"
         .HandleKeyboard = 0
         .ShowTitle = 0
         .ShowDMDOnly = 1
         .ShowFrame = 0
         .HandleMechanics = False
		 .Hidden = hiddenvalue
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With

    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1


'************************************************
' Solenoids
'************************************************

'SolCallback(1)
'SolCallback(2)
'SolCallback(3)
'SolCallback(4)
'SolCallback(5)
'SolCallback(6)
SolCallback(7) = "SolLeftTargetReset"
SolCallback(8) = "dtL.SolunHit 5,"
SolCallback(9) = "SolBallLock"
SolCallback(10) = "bsTrough.SolOut"
SolCallback(11) = "dtL.SolUnHit 4,"
SolCallback(12) = "dtL.SolUnHit 3,"
SolCallback(13) = "dtL.SolUnHit 1,"
SolCallback(14) = "dtL.SolUnHit 2,"
SolCallback(15) = "SolRightTargetReset"
'SolCallback(16)
SolCallback(17) = "SolLane2"
'SolCallback(18)
'SolCallback(19)
SolCallback(20) = "SolLane1"

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

'Trough
   	Set bsTrough=New cvpmBallStack
 	with bsTrough
		.InitSw 0,33,34,35,0,0,0,0
		.InitKick BallRelease,90,7
		.InitExitSnd Soundfx("fx_ballrel",DOFContactors), Soundfx("fx_solenoid",DOFContactors)
		.IsTrough = 1
		.Balls=3
 	end with

' Lane1 Impulse Plunger (by JP Salas)
    Const IMPowerSetting = 36 'Plunger Power
    Const IMTime = 0.6        ' Time in seconds for Full Plunge
    Set plLane1 = New cvpmImpulseP
    With plLane1
        .InitImpulseP sw15, IMPowerSetting, IMTime
        .Random 0.3
        .switch 15
        .InitExitSnd SoundFX("fx_popper", DOFContactors), SoundFX("fx_popper", DOFContactors)
        .CreateEvents "plLane1"
    End With

  ' Lane2 Impulse Plunger (by JP Salas)
    Const IMPowerSetting2 = 40 'Plunger Power
    Const IMTime2 = 0.6        ' Time in seconds for Full Plunge
    Set plLane2 = New cvpmImpulseP
    With plLane2
        .InitImpulseP sw16, IMPowerSetting2, IMTime2
        .Random 0.3
        .switch 16
        .InitExitSnd SoundFX("fx_popper", DOFContactors), SoundFX("fx_popper", DOFContactors)
        .CreateEvents "plLane2"
    End With

' Nudging
     vpmNudge.TiltSwitch = 7
     vpmNudge.Sensitivity = 3
     vpmNudge.TiltObj = Array(leftslingshot1, leftslingshot2, rightslingshot1, sw13, sw14)

 	Set dtL=New cvpmDropTarget
		dtL.InitDrop Array(sw28,sw29,sw30,sw31,sw32),Array(28,29,30,31,32)
		dtL.InitSnd SoundFX("fx2_droptarget",DOFContactors),SoundFX("fx2_DTReset",DOFContactors)

 	Set dtR=New cvpmDropTarget
		dtR.InitDrop Array(sw25,sw26,sw27),Array(25,26,27)
		dtR.InitSnd SoundFX("fx2_droptarget2",DOFContactors),SoundFX("fx2_DTReset",DOFContactors)


'*****Drop Lights Off
	dim xx

	For each xx in DTLeftLights: xx.state=0:Next
	For each xx in DTRightLights: xx.state=0:Next

GILights 1

 End Sub

Sub GILights (enabled)
	Dim light
	For each light in GI:light.State = Enabled: Next
End Sub

Sub Table1_KeyDown(ByVal keycode)
	If keycode = LeftTiltKey Then Nudge 90, 2
	If keycode = RightTiltKey Then Nudge 270, 2
	If keycode = CenterTiltKey Then	Nudge 0, 2

	If vpmKeyDown(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.PullBack: PlaySound "fx_plungerpull",0,1,0.25,0.25: 	End If
End Sub

Sub Table1_KeyUp(ByVal keycode)
	If keycode = PlungerKey Then Plunger.Fire: PlaySound "fx_plunger",0,1,0.25,0.25
	If vpmKeyUp(keycode) Then Exit Sub
End Sub

Sub ShooterLane_Hit
End Sub

Sub BallRelease_UnHit	
End Sub

Sub Drain_Hit()
	PlaySound "fx2_drain2",0,1,0,0.25 : bstrough.addball me
End Sub

 '*****************
 ' Lock Script Courtesy of JP Salas
 '*****************
 
Sub lock3_Hit:Playsound "fx_sensor":Controller.Switch(39) = 1:End Sub

Sub SolBallLock(Enabled)
    If Enabled Then
        Controller.Switch(39) = 0
        lock1a.IsDropped = 1:lock2a.IsDropped = 1:lock3a.IsDropped = 1
        PlaySound SoundFX("fx_popper",DOFContactors)
        lock1.kick 180, 4:lock2.kick 180, 4:lock3.kick 180, 4
    Else
        lock1a.IsDropped = 0:lock2a.IsDropped = 0:lock3a.IsDropped = 0
    End If
End Sub

'Drop Targets
 Sub sw28_Dropped:dtL.Hit 1 : End Sub
 Sub Sw29_Dropped:dtL.Hit 2 : End Sub  
 Sub Sw30_Dropped:dtL.Hit 3 : End Sub
 Sub Sw31_Dropped:dtL.Hit 4 : GI_Ldrop1.state=1 : End Sub
 Sub Sw32_Dropped:dtL.Hit 5 : GI_Ldrop2.state=1 : End Sub

Sub Sw25_Dropped:dtR.Hit 1 : GI_Rdrop1_1.state=1 : GI_Rdrop2_1.state=1 : End Sub
Sub Sw26_Dropped:dtR.Hit 2 : GI_Rdrop1_2.state=1 : GI_Rdrop2_2.state=1 : End Sub
Sub Sw27_Dropped:dtR.Hit 3 : GI_Rdrop2_3.state=1 : End Sub

Sub SolRightTargetReset(enabled)
	dim xx
	if enabled then
		dtR.SolDropUp enabled
		For each xx in DTRightLights: xx.state=0:Next
	end if
End Sub

Sub SolLeftTargetReset(enabled)
	dim xx
	if enabled then
		dtL.SolDropUp enabled
		For each xx in DTLeftLights: xx.state=0:Next
	end if
End Sub

'Bumpers

Sub sw13_Hit : vpmTimer.PulseSw 13 : playsound SoundFX("fx2_bumper_1",DOFContactors): End Sub
Sub sw14_Hit : vpmTimer.PulseSw 14 : playsound SoundFX("fx2_bumper_1",DOFContactors): End Sub

'Wire Triggers
Sub SW36_Hit:Controller.Switch(36)=1 : End Sub
Sub SW36_unHit:Controller.Switch(36)=0:End Sub
Sub SW17_Hit:Controller.Switch(17)=1 : End Sub
Sub SW17_unHit:Controller.Switch(17)=0:End Sub
Sub SW24_Hit:Controller.Switch(24)=1 : End Sub
Sub SW24_unHit:Controller.Switch(24)=0:End Sub
Sub SW23_Hit:Controller.Switch(23)=1 : End Sub
Sub SW23_unHit:Controller.Switch(23)=0:End Sub
Sub SW22_Hit:Controller.Switch(22)=1 : End Sub
Sub SW22_unHit:Controller.Switch(22)=0:End Sub
Sub SW37_Hit:Controller.Switch(37)=1 : End Sub
Sub SW37_unHit:Controller.Switch(37)=0:End Sub
Sub SW40_Hit:Controller.Switch(40)=1 : End Sub
Sub SW40_unHit:Controller.Switch(40)=0:End Sub	
Sub SW18_Hit:Controller.Switch(18)=1 : End Sub
Sub SW18_unHit:Controller.Switch(18)=0:End Sub
Sub SW19_Hit:Controller.Switch(19)=1 : End Sub
Sub SW19_unHit:Controller.Switch(19)=0:End Sub
Sub SW20_Hit:Controller.Switch(20)=1 : End Sub
Sub SW20_unHit:Controller.Switch(20)=0:End Sub
Sub SW21_Hit:Controller.Switch(21)=1 : End Sub
Sub SW21_unHit:Controller.Switch(21)=0:End Sub

'Targets
Sub sw17a_Hit:vpmTimer.PulseSw (17):End Sub
Sub sw17c_Hit:vpmTimer.PulseSw (17):End Sub
Sub sw38_Hit:vpmTimer.PulseSw (38):End Sub


'Spinners
Sub sw4_Spin : vpmTimer.PulseSw (4) :PlaySound "fx_spinner": End Sub
Sub sw5_Spin : vpmTimer.PulseSw (5) :PlaySound "fx_spinner": End Sub

Sub SolKnocker(Enabled)
	If Enabled Then PlaySound SoundFX("Knocker",DOFKnocker)
End Sub

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep1, RStep2, LStep1, RRStep, LUStep, LURStep, RLStep, RLaStep, RLa2Step, R2DropStep, RRRTStep, RRRUStep, RRRLStep

Sub RightSlingShot1_Slingshot
    PlaySound SoundFX("fx_slingshot",DOFContactors), 0, 1, 0.05, 0.05
	vpmtimer.PulseSw(12)
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -30
    RStep1 = 0
    RightSlingShot1.TimerEnabled = 1
End Sub

Sub RightSlingShot1_Timer
    Select Case RStep1
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing3.Visible = 1:sling1.TransZ = 0
        Case 5:RSLing3.Visible = 0:RSLing.Visible = 1:RightSlingShot1.TimerEnabled = 0 
    End Select
    RStep1 = RStep1 + 1
End Sub

Sub RightSlingShot2_Slingshot
    PlaySound SoundFX("fx_slingshot",DOFContactors), 0, 1, 0.05, 0.05
	vpmtimer.PulseSw(10)
    RUSling.Visible = 0
    RUSling1.Visible = 1
    sling4.TransZ = -30
    RStep2 = 0
    RightSlingShot2.TimerEnabled = 1
End Sub

Sub RightSlingShot2_Timer
    Select Case RStep2
        Case 3:RUSLing1.Visible = 0:RUSLing2.Visible = 1:sling4.TransZ = -10
        Case 4:RUSLing2.Visible = 0:RUSLing3.Visible = 1:sling4.TransZ = 0
        Case 5:RUSLing3.Visible = 0:RUSLing.Visible = 1:RightSlingShot2.TimerEnabled = 0 
    End Select
    RStep2 = RStep2 + 1
End Sub

Sub RRubberSlingShot_hit
    RRubber.Visible = 0
    RRubber1.Visible = 1
    RRStep = 0
    RRubberSlingShot.TimerEnabled = 1
End Sub

Sub RRubberSlingShot_Timer
    Select Case RRStep
        Case 3:RRubber1.Visible = 0:RRubber2.Visible = 1
        Case 4:RRubber2.Visible = 0:RRubber.Visible = 1:RRubberSlingShot.TimerEnabled = 0 
    End Select
    RRStep = RRStep + 1
End Sub

Sub LeftSlingShot1_Slingshot
    PlaySound SoundFX("fx_slingshot",DOFContactors), 0, 1, 0.05, 0.05
	vpmtimer.PulseSw(11)
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -36
    LStep1 = 0
    LeftSlingShot1.TimerEnabled = 1
End Sub

Sub LeftSlingShot1_Timer
    Select Case LStep1
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing3.Visible = 1:sling2.TransZ = 0
        Case 5:LSLing3.Visible = 0:LSLing.Visible = 1:LeftSlingShot1.TimerEnabled = 0 
    End Select
    LStep1 = LStep1 + 1
End Sub

Sub LeftSlingShot2_Slingshot
    PlaySound SoundFX("fx_slingshot",DOFContactors), 0, 1, 0.05, 0.05
	vpmtimer.PulseSw(9)
    LUSling.Visible = 0
    LUSling1.Visible = 1
    sling3.TransZ = -38
    LUStep = 0
    LeftSlingShot2.TimerEnabled = 1
End Sub

Sub LeftSlingShot2_Timer
    Select Case LUStep
        Case 3:LUSLing1.Visible = 0:LUSLing2.Visible = 1:sling3.TransZ = -10
        Case 4:LUSLing2.Visible = 0:LUSLing3.Visible = 1:sling3.TransZ = 0
        Case 5:LUSLing3.Visible = 0:LUSLing.Visible = 1:LeftSlingShot2.TimerEnabled = 0 
    End Select
    LUStep = LUStep + 1
End Sub

Sub LURubberSling_hit
    LUSling.Visible = 0
    LUSling4.Visible = 1
    LURStep = 0
    LURubberSling.TimerEnabled = 1
End Sub

Sub LURubberSling_Timer
    Select Case LURstep
        Case 3:LUSling4.Visible = 0:LUSling5.Visible = 1
        Case 4:LUSling5.Visible = 0:LUSling.Visible = 1:LURubberSling.TimerEnabled = 0 
    End Select
    LURStep = LURStep + 1
End Sub

Sub RubberLSling_hit
    RubberL.Visible = 0
    RubberL1.Visible = 1
    RLStep = 0
    RubberLSling.TimerEnabled = 1
End Sub

Sub RubberLSling_Timer
    Select Case RLstep
        Case 3:RubberL1.Visible = 0:RubberL2.Visible = 1
        Case 4:RubberL2.Visible = 0:RubberL.Visible = 1:RubberLSling.TimerEnabled = 0 
    End Select
    RLStep = RLStep + 1
End Sub

Sub RubberLaSling_hit
    RubberLa.Visible = 0
    RubberLa1.Visible = 1
    RLaStep = 0
    RubberLaSling.TimerEnabled = 1
End Sub

Sub RubberLaSling_Timer
    Select Case RLastep
        Case 3:RubberLa1.Visible = 0:RubberLa2.Visible = 1
        Case 4:RubberLa2.Visible = 0:RubberLa.Visible = 1:RubberLaSling.TimerEnabled = 0 
    End Select
    RLaStep = RLaStep + 1
End Sub

Sub RubberLaSling1_hit
    RubberLa.Visible = 0
    RubberLa3.Visible = 1
    RLa2Step = 0
    RubberLaSling1.TimerEnabled = 1
End Sub

Sub RubberLaSling1_Timer
    Select Case RLa2step
        Case 3:RubberLa3.Visible = 0:RubberLa4.Visible = 1
        Case 4:RubberLa4.Visible = 0:RubberLa.Visible = 1:RubberLaSling1.TimerEnabled = 0 
    End Select
    RLa2Step = RLa2Step + 1
End Sub

Sub Rubber2DropSling_hit
    Rubber2Drop.Visible = 0
    Rubber2Drop1.Visible = 1
    R2DropStep = 0
    Rubber2DropSling.TimerEnabled = 1
End Sub

Sub Rubber2DropSling_Timer
    Select Case R2Dropstep
        Case 3:Rubber2Drop1.Visible = 0:Rubber2Drop2.Visible = 1
		Case 4:Rubber2Drop2.Visible = 0:Rubber2Drop3.Visible = 1
        Case 5:Rubber2Drop3.Visible = 0:Rubber2Drop.Visible = 1:Rubber2DropSling.TimerEnabled = 0 
    End Select
    R2DropStep = R2DropStep + 1
End Sub

Sub RRRTSling_hit
    RRRubber.Visible = 0
    RRRubber1.Visible = 1
    RRRTStep = 0
    RRRTSling.TimerEnabled = 1
End Sub

Sub RRRTSling_Timer
    Select Case RRRTstep
        Case 3:RRRubber1.Visible = 0:RRRubber2.Visible = 1
        Case 4:RRRubber2.Visible = 0:RRRubber.Visible = 1:RRRTSling.TimerEnabled = 0 
    End Select
    RRRTStep = RRRTStep + 1
End Sub

Sub RRRUSling_hit
    RRRubber.Visible = 0
    RRRubber3.Visible = 1
    RRRUStep = 0
    RRRUSling.TimerEnabled = 1
End Sub

Sub RRRUSling_Timer
    Select Case RRRUstep
        Case 3:RRRubber3.Visible = 0:RRRubber4.Visible = 1
        Case 4:RRRubber4.Visible = 0:RRRubber.Visible = 1:RRRUSling.TimerEnabled = 0 
    End Select
    RRRUStep = RRRUStep + 1
End Sub

Sub RRRLSling_hit
    RRRubber.Visible = 0
    RRRubber5.Visible = 1
    RRRLStep = 0
    RRRLSling.TimerEnabled = 1
End Sub

Sub RRRLSling_Timer
    Select Case RRRLstep
        Case 3:RRRubber5.Visible = 0:RRRubber6.Visible = 1
        Case 4:RRRubber6.Visible = 0:RRRubber.Visible = 1:RRRLSling.TimerEnabled = 0 
    End Select
    RRRLStep = RRRLStep + 1
End Sub

'Kicker Animations (by JP Salas)
Dim sw15Step, sw16Step

Sub SolLane1(Enabled)
    If Enabled Then
        plLane1.AutoFire
        sw15Step = 0
        Remk2.RotX = 26
        sw15t.TimerEnabled = 1
    End If
End Sub

Sub sw15t_Timer
    Select Case sw15Step
        Case 1:Remk2.RotX = 14
        Case 2:Remk2.RotX = 2
        Case 3:Remk2.RotX = -10:sw15t.TimerEnabled = 0
    End Select

    sw15Step = sw15Step + 1
End Sub

Sub SolLane2(Enabled)
    If Enabled Then
        plLane2.AutoFire
        sw16Step = 0
        Remk1.RotX = 26
        sw16t.TimerEnabled = 1
    End If
End Sub

Sub sw16t_Timer
    Select Case sw16Step
        Case 1:Remk1.RotX = 14
        Case 2:Remk1.RotX = 2
        Case 3:Remk1.RotX = -10:sw16t.TimerEnabled = 0
    End Select

    sw16Step = sw16Step + 1
End Sub

' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 2000)
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

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 20 ' total number of balls
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

'******************************
' Diverse Collection Hit Sounds
'******************************

'Sub aBumpers_Hit (idx): PlaySound SoundFX("fx_bumper", DOFContactors), 0, 1, pan(ActiveBall): End Sub 
Sub aRollovers_Hit(idx):PlaySound "fx_sensor", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aMetals_Hit(idx):PlaySound "fx_MetalHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aRubber_Bands_Hit(idx):PlaySound "fx_rubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aRubbers_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aRubber_Posts_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aRubber_Pins_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aTargets_Hit(idx):PlaySound "fx2_target", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aWoods_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub

'-------------------------------------
' Map lights into array
' Set unmapped lamps to Nothing
'-------------------------------------

Set Lights(1)  = l1
Set Lights(2)  = l2
Set Lights(3)  = l3
Set Lights(4)  = l4
Set Lights(5)  = l5
Set Lights(6)  = l6
Set Lights(7)  = l7
Set Lights(8)  = l8
Set Lights(9)  = l9
Set Lights(10) = l10
Set Lights(11) = l11
Set Lights(12) = l12
'Set Lights(13) = l13	High Score to Date
Set Lights(14) = l14
Set Lights(15) = l15
'Set Lights(16) = l16
Set Lights(17) = l17
Set Lights(18) = l18
Set Lights(19) = l19
Set Lights(20) = l20
Set Lights(21) = l21
Set Lights(22) = l22
Set Lights(23) = l23
Set Lights(24) = l24
Set Lights(25) = l25
Set Lights(26) = l26
Set Lights(27) = l27
Set Lights(28) = l28
Set Lights(29) = l29
Set Lights(30) = l30
Set Lights(31) = l31
'Set Lights(32) = l32
Set Lights(33) = l33
Set Lights(34) = l34
Set Lights(35) = l35
Set Lights(36) = l36
Set Lights(37) = l37
Set Lights(38) = l38
Set Lights(39) = l39
Set Lights(40) = l40
Set Lights(41) = l41
Set Lights(42) = l42
Set Lights(43) = l43
Set Lights(44) = l44
'Set Lights(45) = l45	Game Over
Set Lights(46) = l46
'Set Lights(47) = l47
'Set Lights(48) = l48
Set Lights(49) = l49
Set Lights(50) = l50
Set Lights(51) = l51
Set Lights(52) = l52
Set Lights(53) = l53
Set Lights(54) = l54
Set Lights(55) = l55
Set Lights(56) = l56
Set Lights(57) = l57
Set Lights(58) = l58
Set Lights(59) = l59
Set Lights(60) = l60
'Set Lights(61) = l61	TILT
Set Lights(62) = l62
'63			Match
'Set Lights(64) = l64

'Stern Flight 2000
 'added by Inkochnito
 Sub editDips
     Dim vpmDips:Set vpmDips = New cvpmDips
     With vpmDips
         .AddForm 700, 400, "Flight 2000 - DIP switches"
         .AddChk 2, 10, 180, Array("Match feature", &H00100000)                                                                                                    'dip 21
         .AddChk 2, 25, 115, Array("Credits display", &H00080000)                                                                                                  'dip 20
         .AddFrame 2, 45, 190, "Maximum credits", &H00060000, Array("10 credits", 0, "15 credits", &H00020000, "25 credits", &H00040000, "40 credits", &H00060000) 'dip 18&19
         .AddFrame 2, 120, 190, "High game to date", 49152, Array("points", 0, "1 free game", &H00004000, "2 free games", 32768, "3 free games", 49152)            'dip 15&16
         .AddFrame 2, 195, 190, "High score feature", &H00000020, Array("extra ball", 0, "replay", &H00000020)                                                     'dip 6
         .AddFrame 2, 241, 190, "Spot 1 or 2 'F' lites", &H00001000, Array("spot one 'F'", 0, "spot two 'FF'", &H00001000)                                         'dip 13
         .AddFrame 2, 287, 190, "Right spinner lites", &H00200000, Array("start fresh", 0, "retain", &H00200000)                                                   'dip 22
         .AddChk 205, 10, 190, Array("Talking sound", &H00010000)                                                                                                  'dip 17
         .AddChk 205, 25, 190, Array("Background sound", &H00002000)                                                                                               'dip 14
         .AddFrame 205, 45, 190, "Special award", &HC0000000, Array("no award", 0, "100,000 points", &H40000000, "free ball", &H80000000, "free game", &HC0000000) 'dip 31&32
         .AddFrame 205, 120, 190, "Add a ball memory", &H00000090, Array("one ball", 0, "three balls", &H00000010, "five balls", &H00000090)                       'dip 5&8
         .AddFrame 205, 195, 190, "Balls per game", &H00000040, Array("3 balls", 0, "5 balls", &H00000040)                                                         'dip 7
         .AddFrame 205, 241, 190, "Special limit", &H10000000, Array("no limit", 0, "one replay per ball", &H10000000)                                             'dip 29
         .AddFrame 205, 287, 190, "Apollo 1 and 2 lites", &H00800000, Array("are left off at start of the game", 0, "are turned on at start game", &H00800000)     'dip24
         .AddFrame 100, 333, 190, "Bonus multiplier", &H20000000, Array("kept in memory", 0, "reset", &H20000000)                                                  'dip30
         .AddLabel 50, 385, 300, 20, "After hitting OK, press F3 to reset game with new settings."
         .ViewDips
     End With
 End Sub

Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub
Sub table1_Exit:Controller.Stop:End Sub
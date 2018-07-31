Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' No special SSF tweaks yet.


Const cGameName="smbmush",UseSolenoids=1,UseLamps=0,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff", SCoin="coin"

Const swStartButton=4   'Remap Start Button

LoadVPM "01210000","gts3.vbs",3.1
Dim DesktopMode: DesktopMode = Table1.ShowDT

If DesktopMode = True Then 'Show Desktop components
Ramp16.visible=1
Ramp15.visible=1
Primitive13.visible=1
Primitive22.Visible = 0
Else
Ramp16.visible=0
Ramp15.visible=0
Primitive13.visible=0
Primitive22.Visible = 1
End if

'*************************************************************
'Solenoid Call backs
'**********************************************************************************************************

SolCallback(5)= "SolTop"
SolCallback(6)= "bsBL.SolOut"
SolCallback(7)= "bsBR.SolOut"
SolCallback(8)= "SolRampPost"
SolCallback(9)= "dtL.SolDropUp"
SolCallback(10)= "dtR.SolDropUp"
SolCallback(11)= "dtL.SolHit 1,"
SolCallback(12)= "dtL.SolHit 2,"
SolCallback(13)= "dtL.SolHit 3,"
SolCallback(14)= "dtL.SolHit 4,"
SolCallback(15)= "dtL.SolHit 5,"
SolCallback(16)= "dtR.SolHit 1,"
SolCallback(17)= "dtR.SolHit 2,"
SolCallback(18)= "dtR.SolHit 3,"
SolCallback(19)= "dtR.SolHit 4,"
SolCallback(20)= "dtR.SolHit 5,"
SolCallback(21)= "SetLamp 150," 'BackWall Dome Flasher
SolCallback(22)= "SetLamp 160," 'BackWall Dome Flasher
SolCallback(23)= "SetLamp 170," 'BackWall Dome Flasher
SolCallback(26)= "SetLamp 180," 'BackGlass GI used to control backwall LEDs
'27 Ticket/Coin Meter Enable
SolCallback(28)= "bsTrough.SolOut"
SolCallback(29)= "bsTrough.SolIn"
SolCallback(30)= "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(32)= "vpmNudge.SolGameOn"

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFX("fx_Flipperup",DOFFlippers):LeftFlipper.RotateToEnd
     Else
         PlaySound SoundFX("fx_Flipperdown",DOFFlippers):LeftFlipper.RotateToStart
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFX("fx_Flipperup",DOFFlippers):RightFlipper.RotateToEnd
     Else
         PlaySound SoundFX("fx_Flipperdown",DOFFlippers):RightFlipper.RotateToStart
     End If
End Sub
'**********************************************************************************************************

'Solenoid Controlled toys
'**********************************************************************************************************

Sub SolTop(Enabled)
	If Enabled Then
		If bsTop.Balls Then
			sw24.DestroyBall
			bsTop.ExitSol_On
		End If
	End If
End Sub

Sub SolRampPost(Enabled)
	If Enabled Then
		RampPost.IsDropped=1
		playsound SoundFX("Popper",DOFContactors)
	Else
		RampPost.IsDropped=0
		playsound SoundFX("Popper",DOFContactors)
	End If
End Sub

Sub FlipperTimer_Timer
	LFLogo.roty = LeftFlipper.currentangle  + 0
	RFlogo.roty = RightFlipper.currentangle + 0
End Sub

'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************

Dim bsTrough, dtL, dtR, bsTop, bsBL, bsBR

Sub Table1_Exit
	Controller.Stop
End Sub

Sub Table1_Init
	vpmInit Me
	On Error Resume Next
		With Controller
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
		.SplashInfoLine = "Super Mario Mushroom World Gottlieb"&chr(13)&"You Suck"
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
	PinMAMETimer.Enabled=1

	vpmNudge.TiltSwitch=151
	vpmNudge.Sensitivity=3
	vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3,LeftSlingshot)

	Set bsTrough=New cvpmBallStack
		bsTrough.InitSw 33,0,23,0,0,0,0,0
		bsTrough.InitKick BallRelease,180,3
		bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
		bsTrough.Balls=1
		bsTrough.AddBall 0

	Set bsTop=New cvpmBallStack
		bsTop.InitSw 0,24,0,0,0,0,0,0
		bsTop.InitKick sw24a,117,3
		bsTop.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)

	Set bsBL=New cvpmBallStack
		bsBL.InitSw 0,14,0,0,0,0,0,0
		bsBL.InitKick sw14,115-rnd(1)*3,15
		bsBL.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)

	Set bsBR=New cvpmBallStack
		bsBR.InitSw 0,34,0,0,0,0,0,0
		bsBR.InitKick sw34,255+rnd(1)*3,15
		bsBR.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)

    Set dtL=New cvpmDropTarget
		dtL.InitDrop Array(sw20,sw5,sw15,sw25,sw35),Array (20,5,15,25,35)
		dtL.InitSnd SoundFX("DTDrop",DOFDropTargets),SoundFX("DTReset",DOFContactors)
		dtL.CreateEvents "dtL"

    Set dtR=New cvpmDropTarget
		dtR.InitDrop Array(sw30,sw6,sw16,sw26,sw36),Array (30,06,16,26,36)
		dtR.InitSnd SoundFX("DTDrop",DOFDropTargets),SoundFX("DTReset",DOFContactors)
		dtR.CreateEvents "dtR"
	vpmcreateevents AllSwitches
End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)
	If KeyDownHandler(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Pullback:playsound"plungerpull"
	If KeyCode=LeftFlipperKey Then Controller.Switch(81)=1
	If KeyCode=RightFlipperKey Then Controller.Switch(82)=1
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
	If KeyUpHandler(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Fire:PlaySound"plunger"
	If KeyCode=LeftFlipperKey Then Controller.Switch(81)=0
	If KeyCode=RightFlipperKey Then Controller.Switch(82)=0
End Sub

'**********************************************************************************************************
 ' Drain hole and kickers
Sub Drain_Hit:bsTrough.addball me : playsound"drain" : End Sub
Sub sw14_Hit:bsBL.AddBall Me : playsound "popper_ball": End Sub
Sub sw24_Hit:bsTop.AddBall 0 : playsound "popper_ball": End Sub
Sub sw34_Hit:bsBR.AddBall Me : playsound "popper_ball": End Sub

Sub sw80_Hit:Me.DestroyBall:vpmTimer.PulseSwitch 80,200,"HandleLeft" : playsound "popper_ball": End Sub
Sub HandleLeft(swNo):bsBL.AddBall 0:End Sub

Sub sw90_Hit:Me.DestroyBall:vpmTimer.PulseSwitch 90,200,"HandleRight" : playsound "popper_ball": End Sub
Sub HandleRight(swNo):bsBR.AddBall 0:End Sub

'Drop Targets
 Sub Sw20_Dropped:dtL.Hit 1 :End Sub
 Sub Sw5_Dropped: dtL.Hit 2 :End Sub
 Sub Sw15_Dropped:dtL.Hit 3 :End Sub
 Sub Sw25_Dropped:dtL.Hit 4 :End Sub
 Sub Sw35_Dropped:dtL.Hit 5 :End Sub

'Drop Targets
 Sub Sw36_Dropped:dtR.Hit 1 :End Sub
 Sub Sw26_Dropped:dtR.Hit 2 :End Sub
 Sub Sw16_Dropped:dtR.Hit 3 :End Sub
 Sub Sw6_Dropped: dtR.Hit 4 :End Sub
 Sub Sw30_Dropped:dtR.Hit 5 :End Sub

'Bumpers
Sub Bumper1_Hit : vpmTimer.PulseSw(10) : playsound SoundFX("fx_bumper1",DOFContactors): End Sub
Sub Bumper2_Hit : vpmTimer.PulseSw(11) : playsound SoundFX("fx_bumper1",DOFContactors): End Sub
Sub Bumper3_Hit : vpmTimer.PulseSw(12) : playsound SoundFX("fx_bumper1",DOFContactors): End Sub

'Wire Triggers
Sub SW22_Hit:Controller.Switch(22)=1 : playsound"rollover" : End Sub
Sub SW22_unHit:Controller.Switch(22)=0:End Sub

'Gate Triggers
Sub sw0_hit:vpmTimer.pulseSw 0 : End Sub

'Stand Up Target
Sub sw32_hit:vpmTimer.pulseSw 32:DOF 106, DOFPulse:End Sub

'Generic Soudns
Sub Trigger1_Hit: playsound"fx_ballrampdrop" : End Sub
Sub Trigger2_Hit: playsound"fx_ballrampdrop" : End Sub

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


Sub UpdateLamps
    NFadeObjm 1, l1, "bulbcover1_greenOn", "bulbcover1_green" 'Apron LED
	Flash 1, F1a 'Apron LED
    NFadeObjm 2, l2, "bulbcover1_greenOn", "bulbcover1_green" 'Apron LED
	Flash 2, F2a 'Apron LED
    NFadeObjm 3, l3, "bulbcover1_greenOn", "bulbcover1_green" 'Apron LED
	Flash 3, F3 'Apron LED
    NFadeObjm 4, l4, "bulbcover1_greenOn", "bulbcover1_green" 'Apron LED
	Flash 4, F4 'Apron LED
    NFadeObjm 5, l5, "bulbcover1_greenOn", "bulbcover1_green" 'Apron LED
	Flash 5, F5 'Apron LED
    NFadeObjm 6, l6, "bulbcover1_greenOn", "bulbcover1_green" 'Apron LED
	Flash 6, F6 'Apron LED
    NFadeObjm 7, l7, "bulbcover1_greenOn", "bulbcover1_green" 'Apron LED
	Flash 7, F7 'Apron LED

	NFadeLm 10, l10 'GI
	NFadeL 10, l10a 'GI
	NFadeLm 11, l11 'GI
	NFadeL 11, l11a 'GI

    NFadeObjm 14, l14, "bulbcover1_redOn", "bulbcover1_red"
	Flash 14, F14 'Mushroom Arch entrance
    NFadeObjm 15, l15, "bulbcover1_redOn", "bulbcover1_red"
    Flash 15, F15 'Mushroom Arch entrance
    NFadeObjm 16, l16, "bulbcover1_redOn", "bulbcover1_red"
    Flash 16, F16 'Mushroom Arch entrance
    NFadeObjm 17, l17, "bulbcover1_redOn", "bulbcover1_red"
    Flash 17, F17 'Mushroom Arch entrance

	NFadeL 20, l20 'Bumper 1
	NFadeL 21, l21 'Bumper 2
	NFadeL 22, l22 'Bumper 3

	NFadeL 23, l23

   NFadeObjm 24, l24, "bulbcover1_redOn", "bulbcover1_red"
   Flash 24, F24 'Mario L model LED
   NFadeObjm 25, l25, "bulbcover1_redOn", "bulbcover1_red"
   Flash 25, F25 'Mario L model LED
   NFadeObjm 26, l26, "bulbcover1_redOn", "bulbcover1_red"
   Flash 26, F26 'Mario L model LED
   NFadeObjm 27, l27, "bulbcover1_redOn", "bulbcover1_red"
   Flash 27, F27 'Mario L model LED

    NFadeL 33, l33
    NFadeL 34, l34
    NFadeL 35, l35
    NFadeL 36, l36
    NFadeL 37, l37

    NFadeL 43, l43
    NFadeL 44, l44
    NFadeL 45, l45
    NFadeL 46, l46
    NFadeL 47, l47

    NFadeL 53, l53
    NFadeL 54, l54
    NFadeL 55, l55
    NFadeL 56, l56
    NFadeL 57, l57

    INFadeLm 120, L120a 'Left Sling GI
    INFadeL 120, L120b 'Left Sling GI
    NFadeLm 121, F121  'Left Orange Dome
    NFadeL 121, F121a  'Left Orange Dome
    INFadeLm 122, L122a 'Left ramp GI
    INFadeLm 122, L122b 'Left ramp GI
    INFadeL 122, L122c 'Left ramp GI
    INFadeLm 123, L123a 'Left Drop Target GI
    INFadeL 123, L123b 'Left Drop Target GI
    INFadeLm 124, l124a
    INFadeLm 124, l124b
    INFadeL 124, l124c
    INFadeLm 125, l125a
    INFadeLm 125, l125b
    INFadeL 125, l125c
    NFadeLm 126, F126  'Right Orange Dome
    NFadeLm 126, F126a  'Right Orange Dome
	INFadeLm 127, l127a 'Right Slling GI
	INFadeLm 127, l127b 'Right Slling GI
	INFadeL 127, l127c 'Right Slling GI


'Solenoid Controlled Flashers
   NFadeObjm 150, L221, "dome2_0_orangeON", "dome2_0_orange"
	Flash 150, F221
   NFadeObjm 160, L222, "dome2_0_redON", "dome2_0_red"
	Flash 160, F222
   NFadeObjm 170, L223, "dome2_0_orangeON", "dome2_0_orange"
	Flash 170, F223

'BackWall LEDs
   NFadeObjm 180, LED1, "bulbcover1_orange", "bulbcover1_orangeOn"
   IFlashm 180, F180a
   NFadeObjm 180, LED2, "bulbcover1_red", "bulbcover1_redOn"
   IFlashm 180, F180b
   NFadeObjm 180, LED3, "bulbcover1_blue", "bulbcover1_blueOn"
   IFlashm 180, F180c
   NFadeObjm 180, LED4, "bulbcover1_red", "bulbcover1_redOn"
   IFlashm 180, F180d
   NFadeObjm 180, LED5, "bulbcover1_orange", "bulbcover1_orangeOn"
   IFlashm 180, F180e
   NFadeObjm 180, LED6, "bulbcover1_orange", "bulbcover1_orangeOn"
   IFlashm 180, F180f
   NFadeObjm 180, LED7, "bulbcover1_red", "bulbcover1_redOn"
   IFlashm 180, F180g
   NFadeObjm 180, LED8, "bulbcover1_blue", "bulbcover1_blueOn"
   IFlashm 180, F180h
   NFadeObjm 180, LED9, "bulbcover1_red", "bulbcover1_redOn"
   IFlashm 180, F180i
   NFadeObjm 180, LED10, "bulbcover1_orange", "bulbcover1_orangeOn"
   IFlash 180, F180j

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

'Inverted Lights always on but light turns them off

Sub INFadeL(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.state = 1:FadingLevel(nr) = 1
        Case 5:object.state = 0:FadingLevel(nr) = 0
    End Select
End Sub

Sub INFadeLm(nr, object) ' used for multiple lights
    Select Case FadingLevel(nr)
        Case 4:object.state = 1
        Case 5:object.state = 0
    End Select
End Sub

Sub IFlash(nr, object)
    Select Case FadingLevel(nr)
        Case 4 'on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
            If FlashLevel(nr) > FlashMax(nr) Then
                FlashLevel(nr) = FlashMax(nr)
                FadingLevel(nr) = 1 'completely on
            End if
            Object.IntensityScale = FlashLevel(nr)
        Case 5 ' off
            FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown(nr)
            If FlashLevel(nr) < FlashMin(nr) Then
                FlashLevel(nr) = FlashMin(nr)
                FadingLevel(nr) = 0 'completely off
            End if
            Object.IntensityScale = FlashLevel(nr)
    End Select
End Sub

Sub IFlashm(nr, object) 'multiple flashers, it just sets the flashlevel
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
	vpmTimer.PulseSw 13
	DOF 102, DOFPulse
    PlaySound SoundFX("right_slingshot",DOFContactors), 0, 1, 0.05, 0.05
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0:
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
	vpmTimer.PulseSw 13
	DOF 101, DOFPulse
    PlaySound SoundFX("left_slingshot",DOFContactors),0,1,-0.05,0.05
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0:
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
	PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
	PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
	PlaySound "metalhit_thin", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
	PlaySound "metalhit_medium", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
	PlaySound "metalhit2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
	PlaySound "gate4", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinner_Spin
	PlaySound "fx_spinner",0,.25,0,0.25
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


'Nitro Ground Shaker VP10 table by bodydump
Option Explicit
Randomize

' Thalamus 2018-07-24
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

Const VolBump   = 2    ' Bumpers volume.
Const VolGates  = 1    ' Gates volume.
Const VolMetal  = 1    ' Metals volume.
Const VolRH     = 1    ' Rubber hits volume.
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

Const BallSize = 50

Dim VarHidden, UseVPMColoredDMD
If Table1.ShowDT = true then
	Wall28.IsDropped = False
	Wall30.IsDropped = False
	Ramp15.visible = True
	Ramp16.visible = True
	Ramp17.visible = True
	For each xx in Digitss:xx.visible = 1: Next
	For each xx in bumplightsdt:xx.visible = 1: Next
	For each xx in bumplightsfs:xx.visible = 0: Next
	For each xx in dtgi:xx.visible = 1: Next
else
	Wall28.IsDropped = True
	Wall30.IsDropped = True
	Ramp15.visible = False
	Ramp16.visible = False
	Ramp17.visible = False
	For each xx in Digitss:xx.visible = 0: Next
	For each xx in bumplightsdt:xx.visible = 0: Next
	For each xx in bumplightsfs:xx.visible = 1: Next
	For each xx in dtgi:xx.visible = 0: Next
end if

LoadVPM "01120100", "BALLY.VBS", 3.02

'********************
'Standard definitions
'********************

Const UseSolenoids = 1
Const UseLamps = 0
Const UseSync = 0
Const UseGI = 0
Const HandleMech = 0

' Standard Sounds
Const SSolenoidOn = "fx_Solenoid"
Const SSolenoidOff = "fx_solenoidoff"
Const SCoin = "fx_Coin"

Dim bsTrough, bsSaucer, bsSaucer2, dtR

'************
' Table init.
'************


Const cGameName = "ngndshkr" 'Nitro GroundShaker


Sub Table1_Init
    vpmInit Me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Nitro GroundShaker" & vbNewLine & "VP10 table by bodydump"
        .Games(cGameName).Settings.Value("sound") = 1: 'ensure the sound is on
        .Games(cGameName).Settings.Value("rol") = 0
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = VarHidden
        On Error Resume Next
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
        On Error Goto 0
    End With

    ' Nudging
    vpmNudge.TiltSwitch = 7                                                  'plumb tilt
    vpmNudge.Sensitivity = 1
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingShot, RightSlingShot) 'bumpers & slingshots

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0, 8, 0, 0, 0, 0, 0, 0
        .InitKick BallRelease, 90, 8
        .InitEntrySnd "fx_Solenoid", "fx_Solenoid"
        .InitExitSnd SoundFX("fx_ballrel",DOFContactors), SoundFX("fx_Solenoid",DOFContactors)
        .Balls = 1
    End With

    ' Right Drop Targets
    Set dtR = New cvpmDropTarget
    With dtR
        .InitDrop Array(Array(sw12, sw12a, DT12), Array(sw13, sw13a, DT13), Array(sw14, sw14a, DT14), Array(sw15, sw15a, DT15)), Array(12, 13, 14, 15)
        .InitSnd SoundFX("fx_droptarget",DOFContactors), SoundFX("fx_resetdrop",DOFContactors)
    End With


	Set bsSaucer=New cvpmBallStack
	With bsSaucer
		.InitSaucer KickerUpper,24,185,15
		.InitExitSnd SoundFX("fx_Popper",DOFContactors), SoundFX("fx_Solenoid",DOFContactors)
		.KickAngleVar=2
	End With

	set bsSaucer2=New cvpmBallStack
	With bsSaucer2
		.InitSaucer KickerLower,32,182,20
		.InitExitSnd SoundFX("fx_Popper",DOFContactors), SoundFX("fx_Solenoid",DOFContactors)
		.KickAngleVar=1
	End With
    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1


	If nightday<5 then
		For each xx in GI:xx.intensityscale = 2:Next
		For each xx in inserts:xx.intensity = xx.intensity*3: Next
		l51a.intensityscale = 2
		l51.intensityscale = 2
		l51b.intensityscale = 2
	End If
	If nightday>5 then
		For each xx in GI:xx.intensityscale = 1.5:Next
		l51a.intensityscale = 1.5
		l51.intensityscale = 1.5
		l51b.intensityscale = 1.5
	End If
	If nightday>10 then
		For each xx in GI:xx.intensityscale = 1:Next
		l51a.intensityscale = 1
		l51.intensityscale = 1
		l51b.intensityscale = 1
	End If
	If nightday>20 then
		For each xx in GI:xx.intensityscale = .9:Next
		l51a.intensityscale = .9
		l51.intensityscale = .9
		l51b.intensityscale = .9
	End If
	If nightday>30 then
		For each xx in GI:xx.intensityscale = .8:Next
		For each xx in inserts:xx.intensity = xx.intensity*.8: Next
		l51a.intensityscale = .8
		l51.intensityscale = .8
		l51b.intensityscale = .8
	End If
End Sub

'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
    If keycode = PlungerKey Then Plunger.Pullback
    If keycode = LeftTiltKey Then vpmNudge.DoNudge 90, 3:PlaySound SoundFX("fx_nudge",0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then vpmNudge.DoNudge 270, 3:PlaySound SoundFX("fx_nudge",0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then vpmNudge.DoNudge 0, 4:PlaySound SoundFX("fx_nudge",0), 0, 1, 0, 0.25
    If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If keycode = PlungerKey Then Plunger.Fire:PlaySoundAtVol SoundFX("SolOn",DOFContactors), plunger, 1
    If vpmKeyUp(keycode) Then Exit Sub
End Sub

'***********
' Solenoids
' the commented solenoids are not in used in this script
'***********

SolCallBack(6) = "vpmSolSound SoundFX(""fx_knocker"",DOFKnocker),"
SolCallBack(7) = "bsTrough.SolOut"
SolCallBack(8) = "bsSaucer2.SolOut"
SolCallBack(9) = "bsSaucer.SolOut"
'SolCallBack(10)="" 'bumper
'SolCallBack(11)= "vpmSolSound ""lsling""," 			'left slingshot
'SolCallBack(12) = "" 'bumper
'SolCallBack(13) = "" 'bumper
'SolCallBack(14) = "vpmSolSound ""rsling""," 			'right slingshot
SolCallBack(15)= "dropreset"
'SolCallBack(18)= coinlockout
SolCallBack(19) = "vpmNudge.SolGameOn"

Sub dropreset(enabled)
	if enabled Then
		delaytimer.enabled = 1
	End If
End Sub
Sub delaytimer_timer
	dtR.DropSol_On
	delaytimer.enabled = 0
End Sub


'***************
'  Slingshots
'***************
'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
	vpmTimer.PulseSw(37)
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
	vpmTimer.PulseSw(38)
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


'***************
'   Bumpers
'***************

Sub Bumper1_Hit:vpmTimer.PulseSw 36:PlaySoundAtVol SoundFX("fx_bumper",DOFContactors),Bumper1, VolBump:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 40:PlaySoundAtVol SoundFX("fx_bumper",DOFContactors),Bumper2, VolBump:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 39:PlaySoundAtVol SoundFX("fx_bumper",DOFContactors),Bumper3, VolBump:End Sub

'*********************
' Switches & Rollovers
'*********************

Sub sw18_Hit:Controller.Switch(18) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw18_unHit:Controller.Switch(18) = 0:End Sub

Sub sw19_Hit:Controller.Switch(19) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw19_unHit:Controller.Switch(19) = 0:End Sub

Sub sw22_Hit:Controller.Switch(22) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw22_unHit:Controller.Switch(22) = 0:End Sub

Sub sw23_Hit:Controller.Switch(23) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw23_unHit:Controller.Switch(23) = 0:End Sub

Sub sw33_Hit:Controller.Switch(33) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw33_unHit:Controller.Switch(33) = 0:End Sub

Sub sw34_Hit:Controller.Switch(34) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw34_unHit:Controller.Switch(34) = 0:End Sub

Sub sw5a_Hit:vpmTimer.PulseSw 5:PlaySoundAtVol "fx_sensor", PegPlasticT72, 1:End Sub
Sub sw5b_Hit:vpmTimer.PulseSw 5:PlaySoundAtVol "fx_sensor", WhiteCap14, 1:End Sub
Sub sw5c_Hit:vpmTimer.PulseSw 5:PlaySoundAtVol "fx_sensor", WhiteCap6, 1:End Sub
'****************************
' Drain holes, vuks & saucers
'****************************

Sub Drain_Hit:PlaysoundAtVol "fx_drain", drain, 1:bsTrough.AddBall Me:End Sub
Sub KickerLower_Hit:PlaySoundAtVol "fx_kicker_enter", KickerLower, VolKick:bsSaucer2.AddBall 0:End Sub
Sub KickerUpper_Hit:PlaySoundAtVol "fx_kicker_enter", KickerUpper, VolKick:bsSaucer.AddBall 0:End Sub



'***************
'  Targets
'***************

Sub sw20_Hit
    vpmTimer.PulseSw 20
    PlaySoundAtVol SoundFX("fx_target",DOFContactors), ActiveBall, 1
End Sub

Sub sw21_Hit
    vpmTimer.PulseSw 21
    PlaySoundAtVol SoundFX("fx_target",DOFContactors), ActiveBall, 1
End Sub

Sub sw35_Hit
    vpmTimer.PulseSw 35
    PlaySoundAtVol SoundFX("fx_target",DOFContactors), ActiveBall, 1
End Sub

'***************
' Droptargets
'***************

Sub sw12_Hit:dtR.Hit 1:End Sub
Sub sw13_Hit:dtR.Hit 2:End Sub
Sub sw14_Hit:dtR.Hit 3:End Sub
Sub sw15_Hit:dtR.Hit 4:End Sub

'************
' Spinners
'************

Sub Spinner_Spin:vpmTimer.PulseSw(25):PlaySoundAtVol "fx_spinner", Spinner, VolSpin:End Sub


'********************
'    Flippers
'********************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFX("fx_flipperup",DOFContactors), LeftFlipper, VolFlip
        LeftFlipper.RotateToEnd
    Else
        PlaySoundAtVol SoundFX("fx_flipperdown",DOFContactors), LeftFlipper, VolFlip
        LeftFlipper.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFX("fx_flipperup",DOFContactors), RightFlipper, VolFlip
        RightFlipper.RotateToEnd
    Else
        PlaySoundAtVol SoundFX("fx_flipperdown",DOFContactors), RightFlipper, VolFlip
        RightFlipper.RotateToStart
    End If
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySoundAtVol "fx_rubber_flipper", 0, parm / 10, -0.1, 0.05
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, 0.1, 0.05
End Sub

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

Sub UpdateLamps
    NFadeL 1, l1
    NFadeL 2, l2
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
'    NFadeL 13, l13
    NFadeL 14, l14
'    NFadeLm 15, l15b
'    NFadeL 15, l15
'    NFadeL 16, l16
    NFadeL 17, l17
    NFadeL 18, l18
    NFadeL 19, l19
    NFadeL 20, l20
    NFadeL 21, l21
    NFadeL 22, l22
    NFadeL 23, l23
    NFadeL 24, l24
'    NFadeL 25, l25
    NFadeL 26, l26
'    NFadeL 27, l27
    NFadeL 28, l28
'    NFadeL 29, l29
    NFadeL 30, l30
'    NFadeL 31, l31
'    Flash 32, l32
    NFadeL 33, l33
    NFadeL 34, l34
    NFadeL 35, l35
    NFadeL 36, l36
    NFadeL 37, l37
    NFadeL 38, l38
    NFadeL 39, l39
    NFadeL 40, l40
    NFadeL 41, l41
    NFadeL 42, l42
    NFadeL 43, l43
    NFadeL 44, l44
    NFadeL 45, l45
'    NFadeL 46, l46
'    NFadeL 47, l47
'    NFadeL 48, l48
    NFadeL 49, l49
    NFadeL 50, l50
	NfadeLm 51, l51a
	NfadeLm 51, l51b
    NFadeL 51, l51
    NFadeL 52, l52
    NFadeL 53, l53
    NFadeL 54, l54
    NFadeL 55, l55
    NFadeL 56, l56
    NFadeL 57, l57
    NFadeL 58, l58
    NFadeL 59, l59
    NFadeL 60, l60
    NFadeL 61, l61
'    NFadeL 62, l62
'    NFadeL 63, l63
'    NFadeL 64, l64
'    NFadeL 65, l65
'    NFadeL 66, l66
'    NFadeL 67, l67
'    NFadeL 68, l68
'    NFadeL 69, l69
'    NFadeL 70, l70
'    NFadeL 71, l71
'    NFadeL 72, l72
'    NFadeL 73, l73
'    NFadeL 74, l74
'    NFadeL 75, l75
'    NFadeL 76, l76
'    NFadeL 77, l77
'    NFadeL 78, l78
'    NFadeL 79, l79 'Tournament Light
'    NFadeL 80, l80 'Start Button

    'Flashers

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

' Desktop Objects: Reels & texts (you may also use lights on the desktop)

' Reels

Sub FadeR(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.SetValue 1:FadingLevel(nr) = 6                   'fading to off...
        Case 5:object.SetValue 0:FadingLevel(nr) = 1                   'ON
        Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1              'wait
        Case 9:object.SetValue 2:FadingLevel(nr) = FadingLevel(nr) + 1 'fading...
        Case 10, 11, 12:FadingLevel(nr) = FadingLevel(nr) + 1          'wait
        Case 13:object.SetValue 3:FadingLevel(nr) = 0                  'Off
    End Select
End Sub

Sub FadeRm(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.SetValue 1
        Case 5:object.SetValue 0
        Case 9:object.SetValue 2
        Case 3:object.SetValue 3
    End Select
End Sub

'Texts

Sub NFadeT(nr, object, message)
    Select Case FadingLevel(nr)
        Case 4:object.Text = "":FadingLevel(nr) = 0
        Case 5:object.Text = message:FadingLevel(nr) = 1
    End Select
End Sub

Sub NFadeTm(nr, object, b)
    Select Case FadingLevel(nr)
        Case 4:object.Text = ""
        Case 5:object.Text = message
    End Select
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



'***Sounds***

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

Sub Woods_Hit (idx)
	PlaySound "fx_woodhit", 0, Vol(ActiveBall)*VolWood, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gate3_Hit ()
	PlaySoundAtVol "gate4", gate3, VolGates
End Sub

Sub Gate2_Hit ()
	PlaySoundAtVol "gate", gate2, VolGates
End Sub

Sub gateback_Hit ()
	PlaySoundAtVol "gate", gate1, VolGates
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

Sub Postsphysical_Hit(idx)
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

'*****GI Lights On
dim xx
Sub GIon()
	For each xx in GI:xx.State = 1: Next
	For each xx in inserts:xx.intensityscale = 1: Next
	table1.ColorGradeImage = "ColorGradeLUT256x16_ConSat"
End Sub
Sub GIoff()
	For each xx in GI:xx.State = 0: Next
	For each xx in inserts:xx.intensityscale = 2: Next
	table1.ColorGradeImage = "ColorGradeLUT256x16_ConSat3"
End Sub
Sub Trigger1_Hit()
	GIon
End Sub
Sub gametimer_timer
	If l45.state = 1 then GIoff
	If l61.state = 1 then GIoff
End Sub

'DIP-Switch Menu
'created by Inkochnito
Sub editDips
Dim vpmDips : Set vpmDips = New cvpmDips
With vpmDips
.AddForm 700,400,"Nitro Groundshaker - DIP switches"
.AddChk 7,10,180,Array("Match feature", &H08000000)'dip 28
.AddChk 205,10,115,Array("Credits display", &H04000000)'dip 27
.AddFrame 2, 30, 190,"Maximum credits",&H03000000,Array("10 credits",0,"15 credits", &H01000000,"25 credits", &H02000000,"40 credits", &H03000000)'dip 25&26
.AddFrame 2, 106, 190,"Sound features",&H30000000,Array("chime effects",0,"noises and no background",&H10000000,"noise effects",&H20000000,"noises and background",&H30000000)'dip 29&30
.AddFrame 2, 184, 190,"High score feature",&H00000060,Array("no award",0,"extra ball",&H00000040,"replay",&H00000060)'dip 6&7
.AddFrame 2, 248, 190,"Outlane bonus special adjust",&H00006000,Array("left and right bonus at 10K lites special",0,"left and right bonus at 15K lites special",&H00002000,"left and right bonus at 20K lites special",&H00006000)'dip 14&15
.AddFrame 2, 310, 190,"20K, 30K double bonus adjust",&H00400000,Array ("20K and 30K scores single",0,"20K and 30K scores double",&H00400000)'dip 23
.AddFrame 205, 30, 190,"High score to date",&H00300000,Array("no award",0,"1 credit",&H00100000,"2 credits",&H00200000,"3 credits",&H00300000)'dip 21&22
.AddFrame 205, 106, 190,"Balls per game",&H40000000,Array("3 balls",0,"5 balls",&H40000000)'dip 31
.AddFrame 205, 152, 190,"Outlane special",&H00000080,Array("alternating from left to right",0,"both lanes lite for special",&H00000080)'dip 8
.AddFrame 205, 198, 190,"Attract sound",&H00004000,Array("off",0,"on",&H00004000)'dip 24
.AddFrame 205, 248, 190,"'E' and 'F' rollover",&H80000000,Array("are not tied together",0,"are tied together",&H80000000)'dip 32
.AddFrame 205, 298, 190,"Top saucer double bonus light",32768,Array("lites every 3rd time",0,"lites every 2nd time",32768)'dip 16
.AddLabel 50, 360, 300, 20,"After hitting OK, press F3 to reset game with new settings."
.ViewDips
End With
End Sub
Set vpmShowDips = GetRef("editDips")

'************************************
'          LEDs Display
'Based on Scapino's 7 digit Reel LEDs
'************************************

Dim Digits(32)
Dim Patterns(11)
Dim Patterns2(11)

Patterns(0) = 0     'empty
Patterns(1) = 63    '0
Patterns(2) = 6     '1
Patterns(3) = 91    '2
Patterns(4) = 79    '3
Patterns(5) = 102   '4
Patterns(6) = 109   '5
Patterns(7) = 125   '6
Patterns(8) = 7     '7
Patterns(9) = 127   '8
Patterns(10) = 111  '9

Patterns2(0) = 128  'empty
Patterns2(1) = 191  '0
Patterns2(2) = 134  '1
Patterns2(3) = 219  '2
Patterns2(4) = 207  '3
Patterns2(5) = 230  '4
Patterns2(6) = 237  '5
Patterns2(7) = 253  '6
Patterns2(8) = 135  '7
Patterns2(9) = 255  '8
Patterns2(10) = 239 '9

Set Digits(0) = a0
Set Digits(1) = a1
Set Digits(2) = a2
Set Digits(3) = a3
Set Digits(4) = a4
Set Digits(5) = a5

Set Digits(6) = b0
Set Digits(7) = b1
Set Digits(8) = b2
Set Digits(9) = b3
Set Digits(10) = b4
Set Digits(11) = b5

Set Digits(12) = c0
Set Digits(13) = c1
Set Digits(14) = c2
Set Digits(15) = c3
Set Digits(16) = c4
Set Digits(17) = c5

Set Digits(18) = d0
Set Digits(19) = d1
Set Digits(20) = d2
Set Digits(21) = d3
Set Digits(22) = d4
Set Digits(23) = d5

Set Digits(24) = e0
Set Digits(25) = e1
Set Digits(26) = e2
Set Digits(27) = e3


Sub DisplayTimer_Timer
    On Error Resume Next
    Dim ChgLED, ii, jj, chg, stat
    ChgLED = Controller.ChangedLEDs(&HFF, &HFFFF)
    If Not IsEmpty(ChgLED) Then
        For ii = 0 To UBound(ChgLED)
            chg = chgLED(ii, 1):stat = chgLED(ii, 2)

            For jj = 0 to 10
                If stat = Patterns(jj) OR stat = Patterns2(jj) then Digits(chgLED(ii, 0) ).SetValue jj
            Next
        Next
    End IF
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


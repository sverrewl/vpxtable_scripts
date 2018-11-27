
Option Explicit
Randomize

' Thalamus 2018-07-23
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
Const VolRB     = 1    ' Rubber bands volume.
Const VolPo     = 1    ' Rubber posts volume.
Const VolPi     = 1    ' Rubber pins volume.
Const VolPlast  = 1    ' Plastics volume.
Const VolWood   = 1    ' Woods volume.
Const VolKick   = 1    ' Kicker volume.
Const VolFlip   = 1    ' Flipper volume.


On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01550000", "Bally.vbs", 3.42

Const UseSolenoids = 2
Const UseLamps = 0
Const UseGI = 1
Const UseSync = 0
Const HandleMech = 0

' Standard Sounds
Const SSolenoidOn = "fx_Solenoid"
Const SSolenoidOff = ""
Const SCoin = "fx_Coin"

Dim DesktopMode: DesktopMode = Fireball2.ShowDT

If DesktopMode = True Then 'Show Desktop components
Ramp16.visible=1
Ramp15.visible=1
Ramp17.visible=1
sidewall.visible=1
Else
Ramp16.visible=0
Ramp15.visible=0
Ramp17.visible=0
sidewall.visible=0
End if

SolCallback(1)="dtR.SolDropUp"
SolCallback(2)="dtL.SolDropUp"
SolCallback(3)="dtM.SolDropUp"
SolCallback(6)= "vpmSolSound SoundFX(""fx_Knocker"",DOFKnocker),"
SolCallback(7)= "bsTrough.SolOut"
SolCallBack(8)="bsLSaucer.SolOut"
SolCallBack(9)="bsRSaucer.SolOut"
SolCallback(10)="SolPostKicker"
SolCallback(11)="vpmSolSound ""jet3""," 'sol 3
SolCallback(12)="vpmSolSound ""jet3""," 'sol 4
SolCallback(13)="vpmSolSound ""jet3""," 'sol 5
SolCallback(14)="vpmSolSound ""sling"","
SolCallback(15)="vpmSolSound ""sling"","
SolCallback(17)="SolFireballEnable"
SolCallback(19) = "vpmNudge.SolGameOn"



'************
' Table init.
'************
Const cGameName = "fball_ii"

Dim bsTrough, bsLSaucer, bsRSaucer, dtL, dtR, dtM


Sub Fireball2_Init
    vpmInit Me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Fireball II Bally 1981" & vbNewLine & "VPX table by Javier v1.0"
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = 1
        .Games(cGameName).Settings.Value("rol") = 0 '1= rotated display, 0= normal
        '.SetDisplayPosition 0,0, GetPlayerHWnd 'restore dmd window position
        On Error Resume Next
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
        On Error Goto 0
    End With


    ' Nudging
    vpmNudge.TiltSwitch = swTilt
    vpmNudge.Sensitivity = 3
    vpmNudge.TiltObj = Array(Bumper1B, Bumper2B, Bumper3B, LeftSlingshot, RightSlingshot)

	PinMAMETimer.Interval=PinMAMEInterval
	PinMAMETimer.Enabled=1


    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0, 1, 2, 3, 0, 0, 0, 0
        .InitKick BallRelease, 90, 6
        .InitEntrySnd "fx_Solenoid", "fx_Solenoid"
        .InitExitSnd SoundFX("fx_ballrel",DOFContactors), SoundFX("fx_Solenoid",DOFContactors)
        .Balls = 3
    End With


    'Left droptargets
    Set dtL = New cvpmDropTarget
    With dtL
        .InitDrop Array(Sw32,Sw31,Sw30,Sw29),Array(32,31,30,29)
        .initsnd SoundFX("fx_droptarget", DOFContactors), SoundFX("fx_resetdrop", DOFContactors)
        .CreateEvents "dtL"
    End With

    'Right middle
    Set dtM = New cvpmDropTarget
    With dtM
        .InitDrop Array(Sw35, Sw34, Sw33), Array(35, 34, 34)
        .initsnd SoundFX("fx_droptarget", DOFContactors), SoundFX("fx_resetdrop", DOFContactors)
        .CreateEvents "dtM"
    End With

    'Right droptargets
    Set dtR = New cvpmDropTarget
    With dtR
        .InitDrop Array(Sw24,Sw23,Sw22,Sw21),Array(24,23,22,21)
        .initsnd SoundFX("fx_droptarget", DOFContactors), SoundFX("fx_resetdrop", DOFContactors)
        .CreateEvents "dtR"
    End With

    ' Left Saucer
    Set bsLSaucer = New cvpmBallStack
    With bsLSaucer
        .InitSaucer Sw5, 5, 65, 30
        .KickAngleVar = 2
        .KickForceVar = 1
        .InitExitSnd SoundFX("fx_kicker",DOFContactors), SoundFX("fx_kicker",DOFContactors)
    End With

    ' Right Saucer
    Set bsRSaucer = New cvpmBallStack
    With bsRSaucer
        .InitSaucer sw4, 4, 200, 20
        .KickAngleVar = 2
        .KickForceVar = 1
        .InitExitSnd SoundFX("fx_kicker",DOFContactors), SoundFX("fx_kicker",DOFContactors)
    End With

    KickerBottom.Createball
    KickerBottom.kick 0, 0

End Sub

Sub Fireball2_Paused:Controller.Pause = 1:End Sub
Sub Fireball2_unPaused:Controller.Pause = 0:End Sub


Sub SolPostKicker(enabled)
  If enabled Then
     rubberpost.IsDropped = 1
     PostKicker4.Visible = 1:PostKicker1.Visible = 0:PostKicker.TransZ = 20
     Plunger1.Fire
	 PlaySound"fx_popper"
   Else
     PostKicker4.Visible = 0:PostKicker1.Visible = 1:PostKicker.TransZ = 0
     rubberpost.IsDropped = 0
     Plunger1.Pullback
  End If
End Sub




Sub SolFireballEnable(Enabled)
	If Enabled Then
		KickerBottom1.Kick 0,15
		PlaySound"fx_kicker"
        FFireball.state = 1
     Else
        FFireball.state = 0
	End If
End Sub



'**********
' Keys
'**********

Sub Fireball2_KeyDown(ByVal Keycode)
    If KeyCode = LeftMagnaSave Then Controller.Switch(17) = 1
    If KeyCode = RightMagnaSave Then Controller.Switch(17) = 1
    If keycode = PlungerKey Then PlaySound "fx_PlungerPull", 0, 1, 0.1, 0.05:Plunger.Pullback
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub Fireball2_KeyUp(ByVal Keycode)
    If KeyCode = LeftMagnaSave Then Controller.Switch(17) = 0
    If KeyCode = RightMagnaSave Then Controller.Switch(17) = 0
    If keycode = PlungerKey Then PlaySoundAtVol "fx_plunger", Plunger, 1:Plunger.Fire
    If vpmKeyUp(keycode) Then Exit Sub
End Sub




'**************
' Flipper Subs
'**************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFX("fx_flipperup", DOFContactors), LeftFlipper, VolFlip
        LeftFlipper.RotateToEnd
        LeftFlipper1.RotateToEnd
    Else
        PlaySoundAtVol SoundFX("fx_flipperdown", DOFContactors), LeftFlipper, VolFlip
        LeftFlipper.RotateToStart
        LeftFlipper1.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFX("fx_flipperup", DOFContactors), RightFlipper, VolFlip
        RightFlipper.RotateToEnd
    Else
        PlaySoundAtVol SoundFX("fx_flipperdown", DOFContactors), RightFlipper, VolFlip
        RightFlipper.RotateToStart
    End If
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, -0.1, 0.15
End Sub

Sub Rightflipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, 0.1, 0.15
End Sub

Sub LeftFlipper1_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, -0.1, 0.15
End Sub


'*********
' Switches
'*********

' Slings & div switches

Dim LStep, LStep1, RStep

Sub LeftSlingShot_Slingshot
    PlaySoundAtVol SoundFX("fx_slingshot", DOFContactors), lemk, 1
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 37
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 1:LeftSLing4.Visible = 0:LeftSLing3.Visible = 1:Lemk.RotX = 14
        Case 2:LeftSLing3.Visible = 0:LeftSLing2.Visible = 1:Lemk.RotX = 2
        Case 3:LeftSLing2.Visible = 0:Lemk.RotX = -10:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub Sw25_Slingshot
    PlaySoundAtVol SoundFX("fx_slingshot", DOFContactors), Lemk1, 1
    LeftSling8.Visible = 1
    Lemk1.RotX = 26
    LStep1 = 0
    vpmTimer.PulseSw 25
    Sw25.TimerEnabled = 1
End Sub

Sub Sw25_Timer
    Select Case LStep1
        Case 1:LeftSLing8.Visible = 0:LeftSLing7.Visible = 1:Lemk1.RotX = 14
        Case 2:LeftSLing7.Visible = 0:LeftSLing6.Visible = 1:Lemk1.RotX = 2
        Case 3:LeftSLing6.Visible = 0:Lemk1.RotX = -10:Sw25.TimerEnabled = 0
    End Select
    LStep1 = LStep1 + 1
End Sub









Sub RightSlingShot_Slingshot
    PlaySoundAtVol SoundFX("fx_slingshot", DOFContactors),Remk, 1
    RightSling4.Visible = 1
    Remk.RotX = 26
    RStep = 0
    vpmTimer.PulseSw 36
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 1:RightSLing4.Visible = 0:RightSLing3.Visible = 1:Remk.RotX = 14
        Case 2:RightSLing3.Visible = 0:RightSLing2.Visible = 1:Remk.RotX = 2
        Case 3:RightSLing2.Visible = 0:Remk.RotX = -10:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

' Bumpers
Sub Bumper1B_Hit:vpmTimer.PulseSw 40:PlaySoundAtVol SoundFX("fx_bumper",DOFContactors),Bumper1B, VolBump:End Sub
Sub Bumper2B_Hit:vpmTimer.PulseSw 39:PlaySoundAtVol SoundFX("fx_bumper",DOFContactors),Bumper2B, VolBump:End Sub
Sub Bumper3B_Hit:vpmTimer.PulseSw 38:PlaySoundAtVol SoundFX("fx_bumper",DOFContactors),Bumper3B, VolBump:End Sub


' Drain & holes
Sub Drain_Hit:PlaysoundAtVol "fx_drain", drain, 1:bsTrough.AddBall Me:End Sub
Sub SwGi_Hit: BallsInPlay = BallsInPlay + 1:UpdateGI2(): End Sub
Sub Trigger1_hit: BallsInPlay = BallsInPlay - 1 : End Sub


Sub sw4_Hit:PlaySoundAtVol "fx_kicker_enter", sw4, VolKick:bsRSaucer.AddBall 0:End Sub
Sub sw5_Hit:PlaySoundAtVol "fx_kicker_enter", sw5, VolKick:bsLSaucer.AddBall 0:End Sub

Sub sw8_Hit:Controller.Switch(8) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw8_UnHit:Controller.Switch(8) = 0:End Sub

Sub sw12_Hit:Controller.Switch(12) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw12_UnHit:Controller.Switch(12) = 0:End Sub

Sub sw13_Hit:Controller.Switch(13) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw13_UnHit:Controller.Switch(13) = 0:End Sub

Sub sw14_Hit:Controller.Switch(14) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw14_UnHit:Controller.Switch(14) = 0:End Sub

Sub sw15_Hit:Controller.Switch(15) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw15_UnHit:Controller.Switch(15) = 0:End Sub

Sub sw19_Hit:Controller.Switch(19) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw19_UnHit:Controller.Switch(19) = 0:End Sub

Sub sw20_Hit:Controller.Switch(20) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw20_UnHit:Controller.Switch(20) = 0:End Sub


Sub sw26_Hit:vpmTimer.PulseSw 26:PlaySoundAtVol SoundFX("fx_target",DOFContactors),ActiveBall, 1:End Sub
Sub sw27_Hit:vpmTimer.PulseSw 27:PlaySoundAtVol SoundFX("fx_target",DOFContactors),ActiveBall, 1:End Sub
Sub sw28_Hit:vpmTimer.PulseSw 28:PlaySoundAtVol SoundFX("fx_target",DOFContactors),ActiveBall, 1:End Sub



'Drop-Targets
Sub sw32_dropped():dtL.Hit 1:End Sub
Sub sw31_dropped():dtL.Hit 2:End Sub
Sub sw30_dropped():dtL.Hit 3:End Sub
Sub sw29_dropped():dtL.Hit 4:End Sub

Sub sw35_dropped():dtM.Hit 1:End Sub
Sub sw34_dropped():dtM.Hit 2:End Sub
Sub sw33_dropped():dtM.Hit 3:End Sub

Sub sw24_dropped():dtR.Hit 1:End Sub
Sub sw23_dropped():dtR.Hit 2:End Sub
Sub sw22_dropped():dtR.Hit 3:End Sub
Sub sw21_dropped():dtR.Hit 4:End Sub





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
LampTimer.Interval = 20 'lamp fading speed
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
'    On Error Resume Next
	nFadeL 1, Light1
	nFadeL 2, Light2
	nFadeL 3, Light3
	nFadeL 4, Light4
	nFadeL 5, Light5
	nFadeL 6, Light6
	nFadeL 7, Light7
	nFadeL 8, Light8
	nFadeL 12, Light12
	nFadeLm 15, Light15
	NFadeLm 15, Light15a
	nFadeL 17, Light17
	nFadeL 18, Light18
	nFadeL 19, Light19
	nFadeL 20, Light20
	nFadeL 21, Light21
	nFadeL 22, Light22
	nFadeL 23, Light23
	nFadeL 24, Light24
	nFadeL 28, Light28
	nFadeLm 31, Light31
	nFadeLm 31, Light31a
	nFadeL 33, Light33
	nFadeL 34, Light34
	nFadeL 35, Light35
	nFadeL 36, Light36
	nFadeL 37, Light37
	nFadeL 38, Light38
	nFadeL 39, Light39
	nFadeL 40, Light40
	nFadeL 41, Light41
	nFadeL 42, Light42
	nFadeL 43, Light43
	nFadeL 44, Light44
	nFadeL 46, Light46
	nFadeLm 47, Light47
	nFadeLm 47, Light47a
	nFadeL 49, Light49
	nFadeL 50, Light50
	nFadeL 51, Light51
	nFadeL 52, Light52
	nFadeL 53, Light53
	nFadeL 54, Light54
	nFadeL 55, Light55
	nFadeL 56, Light56
	nFadeL 57, Light57
	nFadeL 58, Light58
	nFadeL 59, Light59
	nFadeL 60, Light60
	nFadeL 62, Light62
	nFadeL 65, Light65
	nFadeL 66, Light66
	nFadeL 67, Light67
	nFadeL 81, Light81
	nFadeL 82, Light82
	nFadeL 83, Light83
	nFadeL 97, Light97
	nFadeL 98, Light98
	nFadeL 99, Light99
	nFadeL 113, Light113
	nFadeL 114, Light114
	nFadeL 115, Light115

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


'******************
' RealTime Updates
'******************

Set MotorCallback = GetRef("RealTimeUpdates")

Sub RealTimeUpdates
    RollingUpdate
    UpdateGI2
End Sub

Sub GiON
    For each x in aGiLights
        x.State = 1
    Next
End Sub

Sub GiOFF
    For each x in aGiLights
        x.State = 0
    Next
End Sub

Dim obj,BallsInPlay, x
BallsInPlay = 0
Sub UpdateGI2()
  If BallsInPlay >= 1 Then GiON
  If BallsInPlay <= 0 Then GiOFF
End Sub


'******************************
' Diverse Collection Hit Sounds
'******************************

Sub aMetals_Hit(idx):PlaySound "fx_metalhit", 0, Vol(ActiveBall)*VolMetal, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Bands_Hit(idx):PlaySound "fx_rubber_band", 0, Vol(ActiveBall)*VolRB, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Posts_Hit(idx):PlaySound "fx_rubber", 0, Vol(ActiveBall)*VolPo, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Pins_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall)*VolPi, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall)*VolPlast, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aGates_Hit(idx):PlaySound "fx_gate", 0, Vol(ActiveBall)*VolGates, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aWoods_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall)*VolWood, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub



'Bally Fireball 2
'added by Inkochnito
Sub editDips
	Dim vpmDips : Set vpmDips = New cvpmDips
	With vpmDips
		.AddForm 700,400,"Fireball II - DIP switches"
		.AddChk 7,10,180,Array("Match feature",&H08000000)'dip 28
		.AddChk 205,10,115,Array("Credits display",&H04000000)'dip 27
		.AddFrame 2,30,190,"Maximum credits",&H03000000,Array("10 credits",0,"15 credits",&H01000000,"25 credits",&H02000000,"40 credits",&H03000000)'dip 25&26
		.AddFrame 2,106,190,"Outlane special lites",&H00000020,Array("after doomsday special is collected",0,"with doomsday special",&H00000020)'dip 6
		.AddFrame 2,152,190,"Making C or D lane",&H00000040,Array("only puts that lane out",0,"puts both lites out",&H00000040)'dip 7
		.AddFrame 2,198,190,"Doomsday special lite",&H00000080,Array("collect special after 2nd 39,000",0,"collect special after 1st 39,000",&H00000080)'dip 8
		.AddFrame 2,244,190,"Fireball bonus units limit",&H00002000,Array("12 units",0,"23 units",&H00002000)'dip 14
		.AddFrame 2,290,190,"Center 15,000 lite adjust",&H00004000,Array("15,000 is off at start of game",0,"15,000 is on at start of game",&H00004000)'dip 15
		.AddFrame 2,336,190,"Game over attract",32768,Array("no voice",0,"voice says: Fireball awaits you",32768)'dip 16
		.AddFrame 205,30,190,"Balls per game",&HC0000000,Array("2 balls",&HC0000000,"3 balls",0,"4 balls",&H80000000,"5 balls",&H40000000)'dip 31&32
		.AddFrame 205,106,190,"Fireball 2X,3X,4X,5X bonus lite",&H00100000,Array("will reset each ball",0,"in memory",&H00100000)'dip 21
		.AddFrame 205,152,190,"Odin and Wotan saucer arrow",&H00200000,Array("will reset each ball",0,"in memory",&H00200000)'dip 22
		.AddFrame 205,198,190,"Odin and Wotan3 target arrow lite",&H00400000,Array("will reset each ball",0,"in memory",&H00400000)'dip 23
		.AddFrame 205,244,190,"A-B-C-D land arrow lite",&H00800000,Array("will reset each ball",0,"in memory",&H00800000)'dip 24
		.AddFrame 205,290,190,"Replay limit",&H10000000,Array("1 replay per game",0,"unlimited replays",&H10000000)'dip 29
		.AddFrame 205,336,190,"Collect Fireball bonus in outhole adjust",&H20000000,Array("only regular bonus collected",0,"Fireball and regular bonus collected",&H20000000)'dip 30
		.AddLabel 50,400,350,20,"Set selftest position 16,17,18 and 19 to 03 for the best gameplay."
		.AddLabel 50,420,300,20,"After hitting OK, press F3 to reset game with new settings."
		.ViewDips
	End With
End Sub
Set vpmShowDips = GetRef("editDips")



'**********************************************************************************************************
'Digital Display
'**********************************************************************************************************
Dim Digits(32)

' 1st Player
Digits(0) = Array(LED10,LED11,LED12,LED13,LED14,LED15,LED16)
Digits(1) = Array(LED20,LED21,LED22,LED23,LED24,LED25,LED26)
Digits(2) = Array(LED30,LED31,LED32,LED33,LED34,LED35,LED36)
Digits(3) = Array(LED40,LED41,LED42,LED43,LED44,LED45,LED46)
Digits(4) = Array(LED50,LED51,LED52,LED53,LED54,LED55,LED56)
Digits(5) = Array(LED60,LED61,LED62,LED63,LED64,LED65,LED66)
Digits(6) = Array(LED70,LED71,LED72,LED73,LED74,LED75,LED76)

' 2nd Player
Digits(7) = Array(LED80,LED81,LED82,LED83,LED84,LED85,LED86)
Digits(8) = Array(LED90,LED91,LED92,LED93,LED94,LED95,LED96)
Digits(9) = Array(LED100,LED101,LED102,LED103,LED104,LED105,LED106)
Digits(10) = Array(LED110,LED111,LED112,LED113,LED114,LED115,LED116)
Digits(11) = Array(LED120,LED121,LED122,LED123,LED124,LED125,LED126)
Digits(12) = Array(LED130,LED131,LED132,LED133,LED134,LED135,LED136)
Digits(13) = Array(LED140,LED141,LED142,LED143,LED144,LED145,LED146)

' 3rd Player
Digits(14) = Array(LED150,LED151,LED152,LED153,LED154,LED155,LED156)
Digits(15) = Array(LED160,LED161,LED162,LED163,LED164,LED165,LED166)
Digits(16) = Array(LED170,LED171,LED172,LED173,LED174,LED175,LED176)
Digits(17) = Array(LED180,LED181,LED182,LED183,LED184,LED185,LED186)
Digits(18) = Array(LED190,LED191,LED192,LED193,LED194,LED195,LED196)
Digits(19) = Array(LED200,LED201,LED202,LED203,LED204,LED205,LED206)
Digits(20) = Array(LED210,LED211,LED212,LED213,LED214,LED215,LED216)

' 4th Player
Digits(21) = Array(LED220,LED221,LED222,LED223,LED224,LED225,LED226)
Digits(22) = Array(LED230,LED231,LED232,LED233,LED234,LED235,LED236)
Digits(23) = Array(LED240,LED241,LED242,LED243,LED244,LED245,LED246)
Digits(24) = Array(LED250,LED251,LED252,LED253,LED254,LED255,LED256)
Digits(25) = Array(LED260,LED261,LED262,LED263,LED264,LED265,LED266)
Digits(26) = Array(LED270,LED271,LED272,LED273,LED274,LED275,LED276)
Digits(27) = Array(LED280,LED281,LED282,LED283,LED284,LED285,LED286)

' Credits
Digits(28) = Array(LED4,LED2,LED6,LED7,LED5,LED1,LED3)
Digits(29) = Array(LED18,LED9,LED27,LED28,LED19,LED8,LED17)
' Balls
Digits(30) = Array(LED39,LED37,LED48,LED49,LED47,LED29,LED38)
Digits(31) = Array(LED67,LED58,LED69,LED77,LED68,LED57,LED59)

Sub DisplayTimer_Timer
	Dim ChgLED,ii,num,chg,stat,obj
	ChgLed = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
If Not IsEmpty(ChgLED) Then
		If DesktopMode = True Then
		For ii = 0 To UBound(chgLED)
			num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
			if (num < 32) then
				For Each obj In Digits(num)
					If chg And 1 Then obj.State = stat And 1
					chg = chg\2 : stat = stat\2
				Next
			else
			end if
		next
		end if
end if
End Sub


'**********************************************************************************************************
'**********************************************************************************************************

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

Sub PlaySoundAtVol(sound, tableobj, Volume)
  PlaySound sound, 1, Volume, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
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

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "Fireball2" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / Fireball2.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "Fireball2" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / Fireball2.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "Fireball2" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / Fireball2.width-1
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

Const tnob = 4 ' total number of balls in this table is 4, but always use a higher number here because of the timing
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub RollingUpdate()
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
Sub Fireball2_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


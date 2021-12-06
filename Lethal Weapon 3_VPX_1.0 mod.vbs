'gtxjoe changes - flipper sounds, flipper strength, friction. bumper strength, right loop exit metal wall

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
Const VolPo     = .8   ' Rubber posts volume.
Const VolPi     = 1    ' Rubber pins volume.
Const VolPlast  = 1    ' Plastics volume.
Const VolWood   = 1    ' Woods volume.
Const VolFlip   = 1    ' Flipper volume.


On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0


Const Ballsize = 51
Dim BallMass:BallMass=(BallSize^3)/125000

LoadVPM "01120100", "DE.VBS", 3.36

'********************
'Standard definitions
'********************

Const UseSolenoids = 2
Const UseLamps = 0
Const UseSync = 1
Const HandleMech = 0

' Standard Sounds
Const SSolenoidOn = "fx_Solenoidon"
Const SSolenoidOff = "fx_solenoidoff"
Const SCoin = "fx_Coin"


'Solenoids
SolCallback(1)  = "bsTrough.SolIn"
SolCallback(2)  = "bsTrough.SolOut"
SolCallback(4)  = "bsLScoop.SolOut"
SolCallback(5)  = "bsRScoop.SolOut"
SolCallback(6)  = "dtMDrop.SolDropUp"
SolCallback(7)  = "dtRDrop.SolDropUp"
SolCallback(9)  = "SolFlasher9"
SolCallback(10) = "SolGi"
SolCallback(11) = "SolFlasher11"
SolCallback(12) = "SolAutoPlungerIM"
SolCallback(15) = "bsVUK.SolOut"
SolCallback(16) = "SolFlasher16"
SolCallback(22) = "SolKickBack"
SolCallback(25) = "SolFlasher25"
SolCallback(26) = "SolFlasher26"
SolCallback(27) = "SolFlasher27"
SolCallback(28) = "SolFlasher28"
SolCallback(29) = "SolFlasher29"
SolCallback(30) = "SolFlasher30"
SolCallback(31) = "SolFlasher31"
SolCallback(32) = "SolFlasher32"


'************
' Table init.
'************

Const cGameName = "lw3_208"

Dim plungerIM, bsTrough, bsRScoop, bsLScoop, bsVuk, dtMDrop, dtRDrop, x

Sub Table1_Init
    vpmInit Me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Lethal Weapon 3" & vbNewLine & "VPX table by Javier v1.0"
    .HandleKeyboard = 0
    .ShowTitle = 0
    .ShowDMDOnly = 1
    .ShowFrame = 0
    .HandleMechanics = 1
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
    vpmNudge.tiltobj = Array(LeftSlingShot,RightSlingShot,LBumper,RBumper,BBumper)

  PinMAMETimer.Interval = PinMAMEInterval
  PinMAMETimer.Enabled = 1

  'Drain & BallRelease
     Set bsTrough=new cvpmBallStack
     With bsTrough
        .InitSw 10,13,12,11,0,0,0,0
        .InitKick BallRelease, 90, 6
        .InitEntrySnd "fx_Solenoid", "fx_Solenoid"
        .InitExitSnd SoundFX("fx_ballrel",DOFContactors), SoundFX("fx_Solenoid",DOFContactors)
        .Balls = 3
     End With


    ' Impulse Plunger
    Const IMPowerSetting = 43 'Plunger Power
    Const IMTime = 0.6        ' Time in seconds for Full Plunge
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swplunger, IMPowerSetting, IMTime
        .Random 0.3
        .switch 14
        .InitExitSnd SoundFX("bumper_retro",DOFContactors), SoundFX("fx_target",DOFContactors)
        .CreateEvents "plungerIM"
    End With

     ' Scoop Right
  Set bsRScoop = New cvpmBallStack
  With bsRScoop
      .InitSaucer Sw32, 32, 270, 17
        .KickZ = 0.33
        .InitExitSnd "salidadebola", SoundFX("fx_Solenoid",DOFContactors)
    End With


     ' Scoop Left
  Set bsLScoop = New cvpmBallStack
  With bsLScoop
       .InitSaucer Sw40, 40, 172, 17
        .KickZ = 0.33
        .InitExitSnd "salidadebola", SoundFX("fx_Solenoid",DOFContactors)
    End With


    ' Center Vuk
    Set bsVuk = New cvpmBallStack
    With bsVuk
        .InitSaucer sw31, 31, 0, 40
        .KickZ = 1.5
        .InitExitSnd "fx_vukout_LAH", SoundFX("fx_Solenoid",DOFContactors)
    End With


    Set dtMDrop=New cvpmDropTarget
    With dtMDrop
      .InitDrop Array(Sw25,Sw26,Sw27), Array(25,26,27)
      .InitSnd SoundFX("fx_target",DOFContactors),SoundFX("fx_resetdrop",DOFContactors)
    End With

    Set dtRDrop=New cvpmDropTarget
    With dtRDrop
      .InitDrop Array(Sw33,Sw34,Sw35), Array(33,34,35)
      .InitSnd SoundFX("fx_target",DOFContactors),SoundFX("fx_resetdrop",DOFContactors)
    End With


   If Table1.ShowDT = False then
       Ramp26.visible = 0
       Ramp32.visible = 0
       Ramp33.visible = 0
       Ramp34.visible = 0
       Ramp35.visible = 0
       Ramp36.visible = 0
   End If

End Sub


'*****************
'AutoPlunger
'*****************

Sub SolAutoPlungerIM(Enabled)
    If Enabled Then
        PlungerIM.AutoFire
    End If
End Sub

' KarateKid
Sub SolKickBack(Enabled)
  If Enabled Then
        PlaysoundAtVol "bumper_retro", LaserKickP1, 1
    LaserKick.Enabled=True
        LaserKickP1.TransY = 90
  Else
    LaserKick.Enabled=False
        vpmtimer.addtimer 500, "LaserKickRes '"
  End If
End Sub
Sub LaserKick_Hit: Me.Kick 0,35 End Sub

Sub LaserKickRes()
    LaserKickP1.TransY = 0
End Sub


Sub swPlunger_UnHit
    LaserKickP.TransY = 90
    vpmtimer.addtimer 500, "PlungerRest '"
End Sub

Sub PlungerRes()
    LaserKickP.TransY = 0
End Sub



'Flashers

Sub SolFlasher9(enabled)
 If enabled Then
    Flasher9.opacity = 100
    Flasher9a.state = 1
  Else
    Flasher9.opacity =  0
    Flasher9a.state = 0
 End If
End Sub


Sub SolFlasher11(enabled)
 If enabled Then
    Flasher11.opacity = 100
    Flasher11a.state = 1
  Else
    Flasher11.opacity = 0
    Flasher11a.state = 0
 End If
End Sub


Sub SolFlasher16(enabled)
 If enabled Then
    Flasher16.opacity = 100
    Flasher16a.state = 1
    LightFlasher16.state = 1
  Else
    Flasher16.opacity = 0
    Flasher16a.state = 0
    LightFlasher16.state = 0
 End If
End Sub


Sub SolFlasher25(enabled)
 If enabled Then
    Flasher25.state = 1
  Else
    Flasher25.state = 0
 End If
End Sub


Sub SolFlasher26(enabled)
 If enabled Then
    Flasher26.opacity = 100
  Else
    Flasher26.opacity = 0
 End If
End Sub


Sub SolFlasher27(enabled)
 If enabled Then
    Flasher27.opacity = 100
    FlasherLight27.state = 1
    Flasher27a.state = 1
  Else
    Flasher27.opacity = 0
    Flasher27a.state = 0
    FlasherLight27.state = 0
 End If
End Sub


Sub SolFlasher28(enabled)
 If enabled Then
    Flasher28.opacity = 100
    FlasherLight28.state = 1
    Flasher28a.state = 1
    LightFlasher28.state = 1
  Else
    Flasher28.opacity = 0
    Flasher28a.state = 0
    FlasherLight28.state = 0
    LightFlasher28.state = 0
 End If
End Sub


Sub SolFlasher29(enabled)
 If enabled Then
    Flasher29.opacity = 100
  Else
    Flasher29.opacity = 0
 End If
End Sub


Sub SolFlasher30(enabled)
 If enabled Then
    Flasher30.opacity = 100
  Else
    Flasher30.opacity = 0
 End If
End Sub


Sub SolFlasher32(enabled)
 If enabled Then
    Flasher32.opacity = 100
    Flasher32a.state = 1
  Else
    Flasher32.opacity = 0
    Flasher32a.state = 0
 End If
End Sub

Sub SolFlasher31(enabled)
 If enabled Then
    Flasher31.state = 1
  Else
    Flasher31.state = 0
 End If
End Sub


Sub LampFlasher()
 If LampState (42) = 1 Then
    FlasherLight27L.opacity = 100
    FlasherLight27La.state = 1
    FlasherLight27R.opacity = 100
    FlasherLight27Ra.state = 1
  Else
    FlasherLight27L.opacity = 0
    FlasherLight27La.state = 0
    FlasherLight27R.opacity = 0
    FlasherLight27Ra.state = 0
 End If


 If LampState (41) = 1 Then
    FlasherLight41.opacity = 100
    FlasherLight41a.state = 1
  Else
    FlasherLight41.opacity = 0
    FlasherLight41a.state = 0
 End If

 If LampState (42) = 1 Then
    FlasherLight42.opacity = 100
    FlasherLight42a.state = 1
  Else
    FlasherLight42.opacity = 0
    FlasherLight42a.state = 0
 End If

 If LampState (43) = 1 Then
    FlasherLight43.opacity = 100
    FlasherLight43a.state = 1
  Else
    FlasherLight43.opacity = 0
    FlasherLight43a.state = 0
 End If

 If LampState (44) = 1 Then
    FlasherLight44.opacity = 100
    FlasherLight44a.state = 1
  Else
    FlasherLight44.opacity = 0
    FlasherLight44a.state = 0
 End If

End Sub




'******************
'Keys Up and Down
'*****************

Sub Table1_KeyDown(ByVal Keycode)
    If keycode = plungerkey then controller.switch(9) = 1
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal Keycode)
    If keycode = plungerkey then controller.switch(9) = 0
    If vpmKeyUp(keycode) Then Exit Sub
End Sub


'**************
' Flipper Subs
'**************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)

    If Enabled Then
        PlaySoundAtVol SoundFX("flipperup", DOFContactors), LeftFlipper, VolFlip
        LeftFlipper.RotateToEnd
    Else
        PlaySoundAtVol SoundFX("flipperdown", DOFContactors), LeftFlipper, VolFlip
        LeftFlipper.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)

    If Enabled Then
        PlaySoundAtVol SoundFX("flipperup", DOFContactors), RightFlipper, VolFlip
        RightFlipper.RotateToEnd
    Else
        PlaySoundAtVol SoundFX("flipperdown", DOFContactors), RightFlipper, VolFlip
        RightFlipper.RotateToStart
    End If
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, -0.1, 0.15
End Sub

Sub Rightflipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, 0.1, 0.15
End Sub


' *********
' Switches
' *********

' Drain
Sub Drain_Hit():PlaySoundAtVol "fx_drain", drain ,1: BsTrough.AddBall Me:End Sub

' Slings & div switches

Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    PlaySoundAtVol SoundFX("fx_slingshot",DOFContactors), Lemk, 1
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 36
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

Sub RightSlingShot_Slingshot
    PlaySoundAtVol SoundFX("fx_slingshot",DOFContactors), Remk, 1
    RightSling4.Visible = 1
    Remk.RotX = 26
    RStep = 0
    vpmTimer.PulseSw 37
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

Sub Sw17_Hit:vpmTimer.PulseSw 17:PlaySoundAtVol "fx_plastichit", Activeball, 1:End Sub
Sub Sw18_Hit:vpmTimer.PulseSw 18:PlaySoundAtVol "fx_plastichit", Activeball, 1:End Sub
Sub Sw19_Hit:vpmTimer.PulseSw 19:PlaySoundAtVol "fx_plastichit", Activeball, 1:End Sub
Sub Sw20_Hit:vpmTimer.PulseSw 20:PlaySoundAtVol "fx_plastichit", Activeball, 1:End Sub



Sub Sw21_Hit():PlaysoundAtVol "fx_sensor", Activeball, 1:Controller.Switch(21)=1: End Sub
Sub Sw21_UnHit():Controller.Switch(21)=0: End Sub

Sub Sw22_Hit():PlaysoundAtVol "fx_sensor", Activeball, 1:Controller.Switch(22)=1: End Sub
Sub Sw22_UnHit():Controller.Switch(22)=0: End Sub

' Targets Center
Sub sw25_dropped():dtMDrop.Hit 1:End Sub
Sub sw26_dropped():dtMDrop.Hit 2:End Sub
Sub sw27_dropped():dtMDrop.Hit 3:End Sub

Sub Sw28_Hit():PlaysoundAtVol "fx_sensor", Activeball, 1:Controller.Switch(28)=1: End Sub
Sub Sw28_UnHit():Controller.Switch(28)=0: End Sub

Sub Sw29_Hit():PlaysoundAtVol "fx_sensor", Activeball, 1:Controller.Switch(29)=1: End Sub
Sub Sw29_UnHit():Controller.Switch(29)=0: End Sub

'Right Up Vuk
Sub Sw31_hit:PlaySoundAtVol "fx_kicker", Activeball, 1:bsVuk.AddBall 0:End Sub
Sub Sw31_hUNhit:PlaySoundAtVol "fx_rampR", Activeball, 1:End Sub

'Left Scoop
Sub Sw32_hit:PlaySoundAtVol "fx_kicker", Activeball, 1:bsRScoop.AddBall 0:End Sub

' Targets Right
Sub sw33_dropped():dtRDrop.Hit 1:End Sub
Sub sw34_dropped():dtRDrop.Hit 2:End Sub
Sub sw35_dropped():dtRDrop.Hit 3:End Sub

Sub Sw36_Hit():PlaysoundAtVol "fx_sensor", Activeball, 1:Controller.Switch(36)=1: End Sub
Sub Sw36_UnHit():Controller.Switch(36)=0: End Sub

Sub Sw37_Hit():PlaysoundAtVol "fx_sensor", Activeball, 1:Controller.Switch(37)=1: End Sub
Sub Sw37_UnHit():Controller.Switch(37)=0: End Sub

Sub Sw39_Hit:vpmTimer.PulseSw 39:PlaySoundAtVol "fx_plastichit", Activeball, 1:End Sub

'Right Scoop
Sub Sw40_hit:PlaySoundAtVol "fx_kicker", Activeball, 1:bsLScoop.AddBall 0:End Sub

'Top Lanes
Sub Sw41_Hit():PlaysoundAtVol "fx_sensor", Activeball, 1:Controller.Switch(41)=1: End Sub
Sub Sw41_UnHit():Controller.Switch(41)=0: End Sub

Sub Sw42_Hit():PlaysoundAtVol "fx_sensor", Activeball, 1:Controller.Switch(42)=1: End Sub
Sub Sw42_UnHit():Controller.Switch(42)=0: End Sub

Sub Sw43_Hit():PlaysoundAtVol "fx_sensor", Activeball, 1:Controller.Switch(43)=1: End Sub
Sub Sw43_UnHit():Controller.Switch(43)=0: End Sub



' Bumpers
Sub LBumper_Hit:vpmTimer.PulseSw 44:PlaySoundAtVol SoundFX("fx_bumper",DOFContactors), LBumper, VolBump:End Sub
Sub BBumper_Hit:vpmTimer.PulseSw 45:PlaySoundAtVol SoundFX("fx_bumper",DOFContactors), BBumper, VolBump:End Sub
Sub RBumper_Hit:vpmTimer.PulseSw 46:PlaySoundAtVol SoundFX("fx_bumper",DOFContactors), RBumper, VolBump:End Sub

'Spinner Left
Sub Sw47_Spin:vpmTimer.PulseSw 47:PlaySoundAtVol "fx_spinner" , sw47, 1: End sub
'Spinner Right
Sub Sw48_Spin:vpmTimer.PulseSw 48:PlaySoundAtVol "fx_spinner" , sw48, 1: End sub


'Center Ramp
Sub Sw49_Hit():PlaysoundAtVol "fx_sensor", Activeball, 1:Controller.Switch(49)=1: End Sub
Sub Sw49_UnHit():Controller.Switch(49)=0: End Sub
Sub Sw50_Hit():PlaysoundAtVol "fx_railShort", Activeball, 1:Controller.Switch(50)=1: End Sub
Sub Sw50_UnHit():Controller.Switch(50)=0: End Sub

'10 Point
Sub Sw52_Hit():PlaysoundAtVol "fx_sensor", Activeball, 1:Controller.Switch(52)=1: End Sub
Sub Sw52_UnHit():Controller.Switch(52)=0: End Sub

'Orbits
Sub Sw54_Hit():PlaysoundAtVol "fx_sensor", Activeball, 1:Controller.Switch(54)=1: End Sub
Sub Sw54_UnHit():Controller.Switch(54)=0: End Sub
Sub Sw55_Hit():PlaysoundAtVol "fx_sensor", Activeball, 1:Controller.Switch(55)=1: End Sub
Sub Sw55_UnHit():Controller.Switch(55)=0: End Sub


sub trigger1_hit():
vpmtimer.addtimer 200, "BallHitSound '"
activeball.vely=.1*activeball.vely
end sub
























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
    nFadeL 1, l1
    nFadeL 2, l2
    nFadeL 3, l3
    nFadeL 4, l4
    nFadeL 5, l5
    nFadeL 6, l6
    nFadeL 7, l7
    nFadeL 8, l8
    nFadeL 9, l9
    nFadeL 9, l9
    nFadeL 10, l10
    nFadeL 11, l11
    nFadeL 12, l12
    nFadeL 13, l13
    nFadeL 14, l14
    nFadeL 15, l15
    nFadeL 16, l16
    nFadeL 17, l17
    nFadeL 18, l18
    nFadeL 18, l18aa
    nFadeL 19, l19
    nFadeL 20, l20
    nFadeL 21, l21
    nFadeL 22, l22
    nFadeL 23, l23
    nFadeL 24, l24
    nFadeL 25, l25
    nFadeL 26, l26

    nFadeL 28, l28
    nFadeL 29, l29
    nFadeL 30, l30
    nFadeL 31, l31
    nFadeL 32, l32

    NFadeL 34, BumperB_Flasher
    NFadeLm 34, BumperB_Flasher_a
    nFadeL 35, l35
    nFadeL 36, l36
    nFadeL 37, l37
    nFadeL 38, l38
    nFadeL 39, l39
    nFadeL 40, l40
    nFadeL 45, l45
    nFadeL 46, l46
    nFadeL 47, l47
    nFadeL 48, l48
    nFadeL 49, l49
    nFadeL 50, l50
    nFadeL 51, l51
    nFadeL 52, l52
    nFadeL 53, l53
    nFadeL 54, l54
    nFadeL 55, l55
    nFadeL 56, l56
    nFadeL 57, l57
    nFadeL 58, l58
    nFadeL 59, l59
    NFadeL 60, BumperL_Flasher
    NFadeLm 60, BumperL_Flasher_a
    NFadeL 61, BumperR_Flasher
    NFadeLm 61, BumperR_Flasher_a
    nFadeL 62, l62
    nFadeL 63, l63
    nFadeL 64, l64

    'Flash And LampFlasher
    LampFlasher()

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

' Walls

Sub FadeWS(nr, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:a.IsDropped = 1:b.IsDropped = 1:c.IsDropped = 1:d.IsDropped = 0:FadingLevel(nr) = 0 'Off
        Case 3:a.IsDropped = 1:b.IsDropped = 0:c.IsDropped = 1:d.IsDropped = 1:FadingLevel(nr) = 2 'fading...
        Case 4:a.IsDropped = 1:b.IsDropped = 1:c.IsDropped = 0:d.IsDropped = 1:FadingLevel(nr) = 3 'fading...
        Case 5:a.IsDropped = 0:b.IsDropped = 1:c.IsDropped = 1:d.IsDropped = 1:FadingLevel(nr) = 1 'ON
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
object.color = RGB(0,0,0)
object.colorfull = RGB(2.5*red,2.5*green,2.5*blue)
object.state=1
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

Sub FlashMod(nr, object)
  Object.IntensityScale = FadingLevel(nr)/255
End Sub

Sub LampMod(nr, object)
Object.IntensityScale = FadingLevel(nr)/255
Object.State = LampState(nr)
End Sub






Sub SolGi(enabled)
  If enabled Then
     Playsound "fx_relay_on"
     GiON
   Else
     Playsound "fx_relay_off"
     GiOFF
 End If
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



'****************************************
' Real Time updatess using the GameTimer
'****************************************
'used for all the real time updates

Sub GameTimer_Timer
    RollingUpdate
    LFLogo.RotY = LeftFlipper.CurrentAngle
    RFLogo.RotY = RightFlipper.CurrentAngle
End Sub

'******************************
' Diverse Collection Hit Sounds
'******************************

Sub aMetals_Hit(idx):PlaySound "fx_MetalHit", 0, Vol(ActiveBall)*VolMetal, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Bands_Hit(idx):PlaySound "fx_rubber", 0, Vol(ActiveBall)*VolRB, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Posts_Hit(idx):PlaySound "fx_postrubber", 0, VolPo, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Pins_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall)*VolPi, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aYellowPins_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall)*VolPi, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall)*VolPlast, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, Vol(ActiveBall)*VolGates, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aWoods_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall)*VolWood, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aCaptiveWalls_Hit(idx):PlaySound "fx_collide", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub



' Ramp Soundss
Sub RHelp1_Hit()
activeball.vely=0
    StopSound "fx_railShort"
    vpmtimer.addtimer 200, "BallHitSound '"
End Sub

Sub RHelp2_Hit()
    PlaySound "balldrop", 0, 1, pan(ActiveBall)
End Sub


Sub Ramp_fx_Hit
'    Playsound "fx_railShort"
End Sub

Sub BallHitSound()
    StopSound "fx_railShort"
    PlaySound "balldrop"
End Sub

Sub Ramp_fx_Hit
    PlaysoundAtVol "fx_railShort", ActiveBall, 1
End Sub


Sub RampR_fx1_Hit
    StopSound "fx_railShort"
    PlaysoundAtVol "fx_railShort", ActiveBall, 1
End Sub
Sub RampR_fx2_Hit
    StopSound "fx_railShort"
    PlaysoundAtVol "fx_rampL", ActiveBall, 1
End Sub
Sub RampR_fx3_Hit
    StopSound "fx_railShort"
    PlaysoundAtVol "fx_rampL", ActiveBall, 1
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

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "Table1" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / Table1.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "Table1" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / Table1.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "Table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / Table1.width-1
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

Const tnob = 20 ' total number of balls
Const lob = 0   'number of locked balls
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub RollingUpdate()
    Dim BOT, b, ballpitch
    BOT = GetBalls

    ' stop the sound of deleted balls
    For b = UBound(BOT) + 1 to tnob
        rolling(b) = False
        StopSound("fx_ballrolling" & b)
    Next

    ' exit the sub if no balls on the table
    If UBound(BOT) = 3 Then Exit Sub 'there are always 4 balls on this table

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


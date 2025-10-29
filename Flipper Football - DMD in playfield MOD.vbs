Option Explicit
Randomize

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

Dim UseVPMDMD, RotDMD, HiddenDMD

Dim DesktopMode: DesktopMode = Table1.ShowDT
If DesktopMode = True Then 'Show Desktop components
Ramp16.visible=1
Ramp15.visible=1
Primitive13.visible=1
UseVPMDMD = False
RotDMD = 0: HiddenDMD = 0
Else
Ramp16.visible=0
Ramp15.visible=0
Primitive13.visible=1
UseVPMDMD = True
RotDMD = 1: HiddenDMD = 1
End if

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="ffv104",UseSolenoids=1,UseLamps=0,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff", SCoin="coin"

LoadVPM "00990300", "CAPCOM.VBS", 3.0

'*************************************************************
'Solenoid Call backs
'**********************************************************************************************************
SolCallback(1) = "SolTrough"
SolCallback(2) = "dtTDrop.SolDropDown"

SolCallback(9) = "SolLeftFlip"    ' Left flipper
SolCallback(10) = "SolRightFlip"  ' Right Flipper
SolCallback(11) = "SolLeftUFlip"  ' Upper left flipper
SolCallback(12) = "SolRightUFlip" ' upper right flipper

SolCallback(13) = "SolUL"
SolCallback(14) = "SolUR"
SolCallback(15) = "SolLOWL"
SolCallback(16) = "SolLOWR"

SolCallback(25) = "dtTDrop.SolDropUp"
SolCallback(26) = "dtLDrop.SolDropUp"
SolCallback(27) = "dtRDrop.SolDropUp"

SolCallback(28) = "SetLamp 28,"
SolCallback(29) = "SetLamp 29,"
SolCallback(30) = "SetLamp 30,"
SolCallback(31) = "SetLamp 31,"
SolCallback(32) = "SetLamp 32,"

'**************
' Flipper Subs
'**************

Dim FlippersEnabled
FlippersEnabled = False

'Enable flippers when ball is in play, done in the drain_unhit
'Disable when the ball is not in play, done in the drain_hit and the subwayhandler

Sub DisableFlippers_Timer
  DisableFlippers.Enabled = 0
  FlippersEnabled = False
End Sub

Sub ReEnableFlippers
  DisableFlippers.Enabled = 0
  DisableFlippers.Enabled = 1
End Sub

Sub LeftFlippersOn
  If FlippersEnabled Then
    PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors), LeftFlipper, VolFlip
    PlaySoundAtVol "fx_Flipperup", LeftFlipper1, VolFlip
    LeftFlipper.RotateToEnd
    LeftFlipper1.RotateToEnd
  End If
End Sub

Sub LeftFlippersOff
  PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors), LeftFlipper, VolFlip
  PlaySoundAtVol "fx_Flipperdown",LeftFlipper1, VolFlip
  LeftFlipper.RotateToStart
  LeftFlipper1.RotateToStart
End SUb

Sub RightFlippersOn
  If FlippersEnabled Then
    PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors), RightFlipper, VolFlip
    PlaySoundAtVol "fx_Flipperup", RightFlipper1, VolFlip
    RightFlipper.RotateToEnd
    RightFlipper1.RotateToEnd
  End If
End Sub

Sub RightFlippersOff
  PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors), RightFlipper, VolFlip
  PlaySoundAtVol "fx_Flipperdown", RightFlipper1, VolFlip
  RightFlipper.RotateToStart
  RightFlipper1.RotateToStart
End Sub

'**********************************************************************************************************

'Solenoid Controlled toys
'**********************************************************************************************************

Sub SolTrough(Enabled)
  If Enabled Then
    FlippersEnabled = True
    bsTrough.ExitSol_On
    ReEnableFlippers
  End If
End Sub

Sub SolLeftFlip(Enabled)
  If Enabled Then
    FlippersEnabled = True
    ReEnableFlippers
  End If
End Sub

Sub SolRightFlip(Enabled)
  If Enabled Then
    FlippersEnabled = True
    ReEnableFlippers
  End If
End Sub

Sub SolLeftUFlip(Enabled)
  If Enabled Then
    FlippersEnabled = True
    ReEnableFlippers
  End If
End Sub

Sub SolRightUFlip(Enabled)
  If Enabled Then
    FlippersEnabled = True
    ReEnableFlippers
  End If
End Sub



Sub SolUL(Enabled)
  If Enabled Then
    FlippersEnabled = True
    bsUpperLeft.ExitSol_On
    ReEnableFlippers
  End If
End Sub

Sub SolUR(Enabled)
  If Enabled Then
    FlippersEnabled = True
    bsUpperRight.ExitSol_On
    ReEnableFlippers
  End If
End Sub

Sub SolLOWL(Enabled)
  If Enabled Then
    FlippersEnabled = True
    bsLowerLeft.ExitSol_On
    ReEnableFlippers
  End If
End Sub

Sub SolLOWR(Enabled)
  If Enabled Then
    FlippersEnabled = True
    bsLowerRight.ExitSol_On
    ReEnableFlippers
  End If
End Sub

'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************

Dim bsTrough, bsLowerRight, bsUpperRight, bsLowerLeft, bsUpperLeft, dtTDrop, dtRDrop, dtLDrop

Sub Table1_Init
  vpmInit Me
  On Error Resume Next
    With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine = "Flipper Football (Capcom 1996)"&chr(13)&"You Suck"
    .HandleMechanics=0
    .HandleKeyboard=0
    .ShowDMDOnly=1
    .ShowFrame=0
    .ShowTitle=0
    .hidden = HiddenDMD
    .Games(cGameName).Settings.Value("rol") = RotDMD
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With
     On Error Goto 0

  PinMAMETimer.Interval=PinMAMEInterval
  PinMAMETimer.Enabled=1
  vpmNudge.TiltSwitch = 10
  vpmNudge.Sensitivity = 2
  vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingshot)

  Set bsTrough = New cvpmBallStack
    bsTrough.InitNotrough Drain, 35, 344, 25
    bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
    bsTrough.KickForceVar = 4
    bsTrough.KickAngleVar = 2

  Set bsUpperLeft = New cvpmBallStack
    bsUpperLeft.InitSaucer sw21, 21, 139, 12
    bsUpperLeft.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
    bsUpperLeft.KickForceVar = 3

  Set bsUpperRight = New cvpmBallStack
    bsUpperRight.InitSaucer sw30, 30, 224, 12
    bsUpperRight.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
    bsUpperRight.KickForceVar = 3

  Set bsLowerRight = New cvpmBallStack
    bsLowerRight.InitSaucer sw43, 43, 250, 15
    bsLowerRight.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
    bsLowerRight.KickForceVar = 3

  Set bsLowerLeft = New cvpmBallStack
    bsLowerLeft.InitSaucer sw68, 68, 120, 15
    bsLowerLeft.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
    bsLowerLeft.KickForceVar = 3

  set dtLDrop = new cvpmdroptarget
    dtLDrop.initdrop  Array(sw62,sw63,sw64), Array(62, 63, 64)
    dtLDrop.initsnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

  set dtRDrop = new cvpmdroptarget
    dtRDrop.initdrop  Array(sw54,sw55,sw56), Array(54, 55, 56)
    dtRDrop.initsnd SoundFX("DTDrop2",DOFContactors),SoundFX("DTReset2",DOFContactors)

  set dtTDrop = new cvpmdroptarget
    dtTDrop.initdrop Array(sw57,sw58,sw59,sw60,sw61), Array(57, 58, 59, 60, 61)
    dtTDrop.initsnd SoundFX("DTDrop3",DOFContactors),SoundFX("DTReset3",DOFContactors)


End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)
  if keycode = LeftFlipperKey then Controller.Switch(33) = 1:Controller.Switch(66) = 1:LeftFlippersOn
  if keycode = RightFlipperKey then Controller.Switch(34) = 1:Controller.Switch(49) = 1:RightFlippersOn
  If keycode = PlungerKey or keycode = LockBarKey Then  Controller.Switch(14) = 1
  If KeyDownHandler(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
  if keycode = LeftFlipperKey then Controller.Switch(33) = 0:Controller.Switch(66) = 0:LeftFlippersOff
  if keycode = RightFlipperKey then Controller.Switch(34) = 0:Controller.Switch(49) = 0:RightFlippersOff
  If keycode = PlungerKey or keycode = LockBarKey Then  Controller.Switch(14) = 0
  If KeyUpHandler(keycode) Then Exit Sub
End Sub

'**********************************************************************************************************

 ' Drain hole and kickers
Sub Drain_Hit:bsTrough.AddBall Me: playsoundAtVol"drain",drain,1 : End Sub
Sub sw43_Hit:bsLowerRight.AddBall 0 : playsoundAtVol "popper_ball", sw43, VolKick: End Sub
Sub sw68_Hit:bsLowerLeft.AddBall 0 : playsoundAtVol "popper_ball", sw68, VolKick: End Sub
Sub sw21_Hit:bsUpperLeft.AddBall 0 : playsoundAtVol "popper_ball", sw21, VolKick: End Sub
Sub sw30_Hit:bsUpperRight.AddBall 0 : playsoundAtVol "popper_ball", sw30, VolKick: End Sub

'Drop Targets
 Sub Sw62_Dropped:dtLDrop.Hit 1 :End Sub
 Sub Sw63_Dropped:dtLDrop.Hit 2 :End Sub
 Sub Sw64_Dropped:dtLDrop.Hit 3 :End Sub

 Sub Sw54_Dropped:dtRDrop.Hit 1 :End Sub
 Sub Sw55_Dropped:dtRDrop.Hit 2 :End Sub
 Sub Sw56_Dropped:dtRDrop.Hit 3 :End Sub

 Sub Sw57_Dropped:dtTDrop.Hit 1 :End Sub
 Sub Sw58_Dropped:dtTDrop.Hit 2 :End Sub
 Sub Sw59_Dropped:dtTDrop.Hit 3 :End Sub
 Sub Sw60_Dropped:dtTDrop.Hit 4 :End Sub
 Sub Sw61_Dropped:dtTDrop.Hit 5 :End Sub

'Bumpers
Sub Bumper1_Hit : vpmTimer.PulseSw(50) : playsoundAtVol SoundFX("fx_bumper1",DOFContactors), ActiveBall, VolBump: End Sub
Sub Bumper2_Hit : vpmTimer.PulseSw(51) : playsoundAtVol SoundFX("fx_bumper1",DOFContactors), ActiveBall, VolBump: End Sub
Sub Bumper3_Hit : vpmTimer.PulseSw(52) : playsoundAtVol SoundFX("fx_bumper1",DOFContactors), ActiveBall, VolBump: End Sub

' Trough Return
Sub sw17_Hit:Me.DestroyBall:vpmTimer.PulseSw 17:SubwayHandler:End Sub
Sub sw18_Hit:Me.DestroyBall:vpmTimer.PulseSw 18:SubwayHandler:End Sub
Sub sw19_Hit:Me.DestroyBall:vpmTimer.PulseSw 19:SubwayHandler:End Sub
Sub sw20_Hit:Me.DestroyBall:vpmTimer.PulseSw 20:SubwayHandler:End Sub
Sub sw25_Hit:Me.DestroyBall:vpmTimer.PulseSw 25:SubwayHandler:End Sub
Sub sw26_Hit:Me.DestroyBall:vpmTimer.PulseSw 26:SubwayHandler:End Sub
Sub sw27_Hit:Me.DestroyBall:vpmTimer.PulseSw 27:SubwayHandler:End Sub
Sub sw28_Hit:Me.DestroyBall:vpmTimer.PulseSw 28:SubwayHandler:End Sub

Sub SubwayHandler : playsound "subway": vpmTimer.PulseSwitch 36, 2500, "bsTrough.AddBall": End Sub

' Rollovers
Sub sw70_Hit : Controller.Switch(70) = 1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw70_UnHit : Controller.Switch(70) = 0 : End Sub
Sub sw45_Hit : Controller.Switch(45) = 1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw45_UnHit : Controller.Switch(45) = 0 : End Sub
Sub sw46_Hit : Controller.Switch(46) = 1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw46_UnHit : Controller.Switch(46) = 0 : End Sub
Sub sw71_Hit : Controller.Switch(71) = 1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw71_UnHit : Controller.Switch(71) = 0 : End Sub
Sub sw24_Hit : Controller.Switch(24) = 1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw24_UnHit : Controller.Switch(24) = 0 : End Sub
Sub sw65_Hit : Controller.Switch(65) = 1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw65_UnHit : Controller.Switch(65) = 0 : End Sub
Sub sw23_Hit : Controller.Switch(23) = 1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw23_UnHit : Controller.Switch(23) = 0 : End Sub
Sub sw22_Hit : Controller.Switch(22) = 1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw22_UnHit : Controller.Switch(22) = 0 : End Sub
Sub sw32_Hit : Controller.Switch(32) = 1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw32_UnHit : Controller.Switch(32) = 0 : End Sub
Sub sw31_Hit : Controller.Switch(31) = 1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw31_UnHit : Controller.Switch(31) = 0 : End Sub

 'Stand Up Targets
Sub sw72_Hit:vpmTimer.PulseSw 72:End Sub
Sub sw47_Hit:vpmTimer.PulseSw 47:End Sub
Sub sw53_Hit:vpmTimer.PulseSw 53:End Sub
Sub sw67_Hit:vpmTimer.PulseSw 67:End Sub
Sub sw29_Hit:vpmTimer.PulseSw 29:End Sub

'Spinners
Sub sw40_Spin:vpmTimer.PulseSw 40 : playsoundAtVol"fx_spinner" , sw40, VolSpin: End Sub
Sub sw48_Spin:vpmTimer.PulseSw 48 : playsoundAtVol"fx_spinner" , sw48, VolSpin: End Sub

 'Scoring Rubber
Sub sw69_Hit:vpmTimer.PulseSw 54 : playsoundAtVol"flip_hit_3" , ActiveBall, 1: End Sub
Sub sw44_Hit:vpmTimer.PulseSw 55 : playsoundAtVol"flip_hit_3" , ActiveBall, 1: End Sub

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

     'Special Handling
     'If chgLamp(ii,0) = 2 Then solTrough chgLamp(ii,1)
     'If chgLamp(ii,0) = 4 Then PFGI chgLamp(ii,1)

        Next
    End If
    UpdateLamps
End Sub

Sub UpdateLamps

'Lamp Matrix A
  'Unused

'Lamp Matrix B
  NFadeLm 65, l65a 'GI
  NFadeLm 65, l65b
  NFadeLm 65, l65c
  NFadeLm 65, l65d
  NFadeLm 65, l65e
  NFadeL 65, l65f
  NFadeLm 66, l66a 'GI
  NFadeLm 66, l66b
  NFadeLm 66, l66c
  NFadeL 66, l66d
  NFadeL 67, l67
  NFadeL 68, l68
  NFadeL 69, l69
  NFadeL 70, l70
  NFadeL 71, l71
  NFadeL 72, l72
  NFadeL 73, l73
  NFadeL 74, l74
  NFadeLm 75, l75a 'GI
  NFadeLm 75, l75b
  NFadeLm 75, l75c
  NFadeL 75, l75d
  NFadeLm 76, l76a 'GI
  NFadeLm 76, l76b
  NFadeLm 76, l76c
  NFadeL 76, l76d
  NFadeL 77, l77
  NFadeL 78, l78
  NFadeL 79, l79
  NFadeL 80, l80
  NFadeL 81, l81
  NFadeL 82, l82
  NFadeL 83, l83
  NFadeL 84, l84
  NFadeLm 85, l85a 'GI
  NFadeLm 85, l85b
  NFadeLm 85, l85c
  NFadeLm 85, l85d
  NFadeLm 85, l85e
  NFadeL 85, l85f
  NFadeLm 86, l86a 'GI
  NFadeLm 86, l86b
  NFadeLm 86, l86c
  NFadeL 86, l86d
  NFadeL 87, l87
  NFadeL 88, l88
  NFadeL 89, l89
  NFadeL 90, l90
  NFadeL 91, l91
  NFadeL 92, l92
  NFadeL 93, l93
  NFadeL 94, l94
  NFadeLm 95, l95a 'GI
  NFadeL 95, l95b
  NFadeLm 96, l96a 'GI
  NFadeLm 96, l96b
  NFadeLm 96, l96c
  NFadeL 96, l96d
  NFadeL 97, l97
  NFadeL 98, l98
  NFadeL 99, l99
  NFadeL 100, l100
  NFadeL 101, l101
  NFadeL 102, l102
  NFadeL 103, l103
  NFadeL 104, l104
  NFadeL 105, l105
  NFadeL 106, l106
  NFadeL 107, l107
  NFadeL 108, l108
  NFadeL 109, l109
  NFadeL 110, l110
  NFadeL 111, l111
  NFadeL 112, l112
  NFadeL 113, l113
  NFadeL 114, l114
  NFadeL 115, l115
  NFadeL 116, l116
  NFadeLm 117, l117a 'GI
  NFadeL 117, l117b
  NFadeLm 118, l118a 'GI
  NFadeLm 118, l118b
  NFadeLm 118, l118c
  NFadeL 118, l118d
  NFadeL 119, l119
  NFadeLm 120, l120a 'GI
  NFadeLm 120, l120b
  NFadeLm 120, l120c
  NFadeL 120, l120d
  NFadeL 121, l121
  NFadeL 122, l122
  NFadeL 123, l123
  NFadeL 124, l124
  NFadeL 125, l125
  'FadeL 126, l126
  NFadeLm 127, l127a 'GI
  NFadeLm 127, l127b
  NFadeLm 127, l127c
  NFadeLm 127, l127d
  NFadeL 127, l127a
  NFadeLm 128, l128a 'GI
  NFadeLm 128, l128b
  NFadeLm 128, l128c
  NFadeL 128, l128d

'Solenoid Controlled

  NFadeLm 28, f28a 'bumpers
  NFadeLm 28, f28b
  NFadeLm 28, f28c
  NFadeLm 28, f28d
  NFadeLm 28, f28e
  NFadeL 28, f28f

  NFadeLm 29, f29a
  NFadeL 29, f29b

  NFadeLm 30, f30a
  NFadeL 30, f30b

  NFadeL 31, f31

  NFadeL 32, f32


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


'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
  vpmTimer.PulseSw 42
    PlaySoundAtVol SoundFX("right_slingshot",DOFContactors), sling1, 1
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.rotx = 20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.rotx = 10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.rotx = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  vpmTimer.PulseSw 41
    PlaySoundAtVol SoundFX("left_slingshot",DOFContactors), sling2, 1
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.rotx = 20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.rotx = 10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.rotx = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
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


'*****************************************
' ninuzzu's FLIPPER SHADOWS
'*****************************************

sub FlipperTimer_Timer()
  FlipperLSh.RotZ = LeftFlipper.currentangle
  FlipperRSh.RotZ = RightFlipper.currentangle

End Sub

'*****************************************
' ninuzzu's BALL SHADOW
'*****************************************
Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5)

Sub BallShadowUpdate_timer()
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
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/21)) + 6
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/21)) - 6
        End If
        ballShadow(b).Y = BOT(b).Y + 4
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
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

' The sound is played using the VOL, AUDIOPAN, AUDIOFADE and PITCH functions, so the volume and pitch of the sound
' will change according to the ball speed, and the AUDIOPAN & AUDIOFADE functions will change the stereo position
' according to the position of the ball on the table.


'**************************************
' Explanation of the collision routine
'**************************************

' The collision is built in VP.
' You only need to add a Sub OnBallBallCollision(ball1, ball2, velocity) and when two balls collide they
' will call this routine. What you add in the sub is up to you. As an example is a simple Playsound with volume and paning
' depending of the speed of the collision.


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

Sub Rubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 20 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End Sub

Sub Posts_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolPo, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    RandomSoundRubber()
  End If
End Sub

Sub RandomSoundRubber()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  On Error Resume Next
  Controller.Pause = False
  Controller.Stop
End Sub


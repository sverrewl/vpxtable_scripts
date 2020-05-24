Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="godzilla",UseSolenoids=1,UseLamps=0,UseGI=0, SCoin="coin"

LoadVPM "01000000","SEGA.VBS",3.10

Dim DesktopMode: DesktopMode = Table1.ShowDT

If DesktopMode = True Then 'Show Desktop components
  Ramp16.visible=1
  Ramp15.visible=1
  Primitive13.visible=1
Else
  Ramp16.visible=0
  Ramp15.visible=0
  Primitive13.visible=0
End if

' Thalamus 2020 February : Improved directional sounds
' !! NOTE : Table not verified yet !!

' Options
' Volume devided by - lower gets higher sound

Const VolDiv = 1000    ' Lower number, louder ballrolling/collition sound
Const VolCol = 10      ' Ball collition divider ( voldiv/volcol )

' The rest of the values are multipliers
'
'  .5 = lower volume
' 1.5 = higher volume

Const VolBump   = 5    ' Bumpers volume.
Const VolRol    = 1    ' Rollovers volume.
Const VolGates  = 1    ' Gates volume.
Const VolMetal  = 1    ' Metals volume.
Const VolRH     = 1    ' Rubber hits volume.
Const VolPo     = 1    ' Rubber posts volume.
Const VolPi     = 1    ' Rubber pins volume.
Const VolTarg   = 1    ' Targets volume.
Const VolSpin   = 5    ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.
Const VolD      = 5    ' Diverter volume.

'*************************************************************
'Solenoid Call backs
'**********************************************************************************************************

 SolCallBack(1) = "SolRelease"
 SolCallBack(2) = "Auto_Plunger"
'SolCallback(3) = "solTopMagnet"
'SolCallback(4) = "solMiddleMagnet"
'SolCallback(5) = "solBottomMagnet"
SolCallback(6)  = "SolShake"
SolCallBack(7)="SetLamp 107, "
'SolCallback(14) = "solOrbitMagnet"
SolCallback(17) = "SolRampDiverter"
SolCallBack(18)="SetLamp 118, "
SolCallBack(19)="SetLamp 119, "
SolCallBack(20)="SetLamp 120, "
SolCallBack(25)="SetLamp 125, " 'F1
SolCallBack(26)="SetLamp 126, " 'F2
SolCallBack(27)="SetLamp 127, " 'F3
SolCallBack(28)="SetLamp 128, " 'F4
SolCallBack(29)="SetLamp 129, " 'F5
SolCallBack(30)="SetLamp 130, " 'F6
SolCallBack(31)="SetLamp 131, " 'F7
SolCallBack(32)="SetLamp 132, " 'F8

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors), LeftFlipper, VolFlip:LeftFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors), LeftFlipper, VolFlip:LeftFlipper.RotateToStart
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors), RightFlipper, VolFlip:RightFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors), RightFlipper, VolFlip:RightFlipper.RotateToStart
     End If
End Sub
'**********************************************************************************************************

'Solenoid Controlled toys
'**********************************************************************************************************

Sub SolRelease(Enabled)
  If Enabled Then
  bsTrough.ExitSol_On
  vpmTimer.PulseSw 15
  End If
End Sub

Sub Auto_Plunger(Enabled)
  If Enabled Then
  PlungerIM.AutoFire
  End If
End Sub

Sub SolRampDiverter(enabled)
  If enabled then
    Diverter1.IsDropped = 1
    Diverter2.IsDropped = 0
    PlaySoundAtVol "Diverter", Primitive22, VolD
  Else
    Diverter1.IsDropped = 0
    Diverter2.IsDropped = 1
    PlaySoundAtVol "Diverter", Primitive22, VolD
  End If
End Sub

Sub SolShake(enabled)
  If enabled Then
      ShakeTimer.Enabled = 1
    playsound SoundFX("Motor",DOFContactors)
  Else
      ShakeTimer.Enabled = 0
  End If
End Sub

Sub ShakeTimer_Timer()
  Nudge 0,1
  Nudge 90,1
  Nudge 180,1
  Nudge 270,1
End Sub

Sub solTopMagnet(enabled)
    If enabled Then
    mTopMagnet.MagnetOn = 1
  Else
    mTopMagnet.MagnetOn = 0
End If
End Sub

Sub solMiddleMagnet(enabled)
    If enabled Then
    mMiddleMagnet.MagnetOn = 1
  Else
    mMiddleMagnet.MagnetOn = 0
End If
End Sub

Sub solBottomMagnet(enabled)
    If enabled Then
    mBottomMagnet.MagnetOn = 1
  Else
    mBottomMagnet.MagnetOn = 0
End If
End Sub

Sub solOrbitMagnet(enabled)
    If enabled Then
    mOrbitMagnet.MagnetOn = 1
  Else
    mOrbitMagnet.MagnetOn = 0
End If
End Sub

set GICallback = GetRef("UpdateGI")
Sub UpdateGI(no, Enabled)
  If Enabled Then
    dim xx
    For each xx in GI:xx.State = 1: Next
        PlaySound "fx_relay"

  Else For each xx in GI:xx.State = 0: Next
        PlaySound "fx_relay"

  End If
End Sub

'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************

Dim bsTrough, mTopMagnet, mMiddleMagnet, mBottomMagnet, mOrbitMagnet

Sub Table1_Init
  vpmInit Me
  On Error Resume Next
    With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine = "Godzilla, Sega 1998"&chr(13)&"You Suck"
    .HandleMechanics=0
    .HandleKeyboard=0
    .ShowDMDOnly=1
    .ShowFrame=0
    .ShowTitle=0
    .hidden = 0
        '.Games(cGameName).Settings.Value("sound")=1
    '.PuPHide = 1
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With
     On Error Goto 0

  PinMAMETimer.Interval = PinMAMEInterval
  PinMAMETimer.Enabled = true
  vpmNudge.TiltSwitch = swTilt

  vpmNudge.Sensitivity = 2
  vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingshot)

  Set bsTrough = new cvpmBallStack
    bsTrough.InitSw 0,14,13,12,11,0,0,0
    bsTrough.Balls = 4
    bsTrough.InitKick BallRelease,45,8
        bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)

  Set mTopMagnet = New cvpmMagnet
    mTopMagnet.InitMagnet TopMagnet, 100
    mTopMagnet.Solenoid = 3
    mTopMagnet.GrabCenter = True
    mTopMagnet.CreateEvents "mTopMagnet"

  Set mMiddleMagnet = New cvpmMagnet
    mMiddleMagnet.InitMagnet MiddleMagnet, 2
    mMiddleMagnet.Solenoid = 4
    mMiddleMagnet.GrabCenter = False
    mMiddleMagnet.CreateEvents "mMiddleMagnet"

  Set mBottomMagnet = New cvpmMagnet
    mBottomMagnet.InitMagnet BottomMagnet, 2
    mBottomMagnet.Solenoid = 5
    mBottomMagnet.GrabCenter = False
    mBottomMagnet.CreateEvents "mBottomMagnet"

  Set mOrbitMagnet = New cvpmMagnet
    mOrbitMagnet.InitMagnet OrbitMagnet, 100
    mOrbitMagnet.Solenoid = 14
    mOrbitMagnet.GrabCenter = True
    mOrbitMagnet.CreateEvents "mOrbitMagnet"

    Kicker1.CreateBall
    Kicker1.Kick 0,1
    Kicker2.CreateBall
    Kicker2.Kick 0,1
    Kicker3.CreateBall
    Kicker3.Kick 0,1

  Diverter1.IsDropped = 0
  Diverter2.IsDropped = 1

End Sub


Sub Table1_Paused:Controller.Pause = 1:End Sub
Sub Table1_unPaused:Controller.Pause = 0:End Sub
sub Table1_Exit:Controller.Stop:end sub


'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)
  If KeyDownHandler(keycode) Then Exit Sub
  If KeyCode=PlungerKey Then Controller.Switch(53)=1
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
  If KeyUpHandler(keycode) Then Exit Sub
  If KeyCode=PlungerKey Then Controller.Switch(53)=0
End Sub

     ' Impulse Plunger
  Dim PlungerIM
    Const IMPowerSetting = 55
    Const IMTime = 0.6
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swplunger, IMPowerSetting, IMTime
        .Random 0.3
        .InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
        .CreateEvents "plungerIM"
    End With

'**********************************************************************************************************

 ' Drain hole and kickers
Sub Drain_Hit:bsTrough.addball me : playsoundAtVol "drain", Drain, 1 : End Sub

 'Stand Up Targets
Sub sw17_hit:vpmTimer.pulseSw 17 : End Sub
Sub sw18_hit:vpmTimer.pulseSw 18 : End Sub
Sub sw19_hit:vpmTimer.pulseSw 19 : End Sub

Sub sw21_hit:vpmTimer.pulseSw 21 : End Sub
Sub sw22_hit:vpmTimer.pulseSw 22 : End Sub
Sub sw23_hit:vpmTimer.pulseSw 23 : End Sub

Sub sw29_hit:vpmTimer.pulseSw 29 : End Sub
Sub sw30_hit:vpmTimer.pulseSw 30 : End Sub
Sub sw31_hit:vpmTimer.pulseSw 31 : End Sub
Sub sw32_hit:vpmTimer.pulseSw 32 : End Sub

Sub sw20_Hit:vpmTimer.PulseSw 20 : sw20p.TransX = -10 : sw20.TimerEnabled = 1 : playsound"Target" : End Sub
Sub sw20_Timer : sw20p.TransX = 0 : sw20.TimerEnabled = 0 : End Sub
Sub sw24_Hit:vpmTimer.PulseSw 24 : sw24p.TransX = -10 : sw24.TimerEnabled = 1 : playsound"Target" : End Sub
Sub sw24_Timer : sw24p.TransX = 0 : sw24.TimerEnabled = 0 : End Sub

'Wire Triggers
Sub sw16_Hit : Controller.Switch(16)=1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw16_Unhit : Controller.Switch(16)=0:End Sub
Sub sw41_Hit : Controller.Switch(41)=1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw41_Unhit : Controller.Switch(41)=0:End Sub
Sub sw42_Hit : Controller.Switch(42)=1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw42_Unhit : Controller.Switch(42)=0:End Sub
Sub sw43_Hit : Controller.Switch(43)=1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw43_Unhit : Controller.Switch(43)=0:End Sub
Sub sw44_Hit : Controller.Switch(44)=1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw44_Unhit : Controller.Switch(44)=0:End Sub
Sub sw47_Hit : Controller.Switch(47)=1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw47_Unhit : Controller.Switch(47)=0:End Sub
Sub sw48_Hit : Controller.Switch(48)=1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw48_Unhit : Controller.Switch(48)=0:End Sub
Sub sw57_Hit : Controller.Switch(57)=1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw57_Unhit : Controller.Switch(57)=0:End Sub
Sub sw58_Hit : Controller.Switch(58)=1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw58_Unhit : Controller.Switch(58)=0:End Sub
Sub sw60_Hit : Controller.Switch(60)=1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw60_Unhit : Controller.Switch(60)=0:End Sub
Sub sw61_Hit : Controller.Switch(61)=1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw61_Unhit : Controller.Switch(61)=0:End Sub

'Ramp Gate Triggers
Sub sw26_Hit:vpmTimer.PulseSw 26 : End Sub
Sub sw28_Hit:vpmTimer.PulseSw 28 : End Sub

' Gate Triggers
Sub sw45_Hit:vpmTimer.PulseSw 45 : End Sub

'Spinners
Sub sw46_Spin:vpmTimer.PulseSw 46 : playsoundAtVol"fx_spinner" , sw46, VolSpin: End Sub

'Bumpers
Sub Bumper1_Hit : vpmTimer.PulseSw(49) : playsoundAtVol SoundFX("fx_bumper1",DOFContactors), ActiveBall, VolBump: End Sub
Sub Bumper2_Hit : vpmTimer.PulseSw(50) : playsoundAtVol SoundFX("fx_bumper1",DOFContactors), ActiveBall, VolBump: End Sub
Sub Bumper3_Hit : vpmTimer.PulseSw(51) : playsoundAtVol SoundFX("fx_bumper1",DOFContactors), ActiveBall, VolBump: End Sub

'Generic Sounds
Sub Trigger1_Hit:PlaySoundAtVol "fx_ballrampdrop", Trigger1, 1:End Sub
Sub Trigger2_Hit:PlaySoundAtVol "fx_ballrampdrop", Trigger2, 1:End Sub
Sub Trigger3_Hit:PlaySoundAtVol "fx_ballrampdrop", Trigger3, 1:End Sub
Sub Trigger4_Hit:PlaySoundAtVol "fx_ballrampdrop", Trigger4, 1:End Sub

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
 NFadeL 1, l1
 NFadeL 2, l2
 NFadeL 3, l3
 NFadeL 4, l4
 NFadeLm 5, l5a
 NFadeL 5, l5
 NFadeL 6, l6
 NFadeL 7, l7
 'NFadeL 8, l8 'Launch Button
 NFadeL 9, l9
 NFadeL 10, l10
 NFadeL 11, l11
 NFadeL 12, l12
 NFadeL 13, l13
 NFadeL 14, l14
 NFadeL 15, l15
 NFadeL 16, l16
 NFadeL 17, l17
 NFadeL 18, l18
 NFadeL 19, l19
 NFadeL 20, l20
 NFadeL 21, l21
 NFadeL 22, l22
 NFadeL 23, l23
 NFadeL 24, l24
 NFadeL 25, l25
 NFadeL 26, l26
 NFadeL 27, l27
 NFadeLm 28, l28 'PF GI
 NFadeL 28, l28a
 NFadeL 29, l29
 NFadeL 30, l30
 NFadeL 31, l31
 NFadeLm 32, l32 'PF GI
 NFadeL 32, l32a
 NFadeL 33, l33
 NFadeL 34, l34
 NFadeL 35, l35
 NFadeL 36, l36
 NFadeL 37, l37
 NFadeL 38, l38
 NFadeL 39, l39

 NFadeL 41, l41
 NFadeL 42, l42
 NFadeL 43, l43
 NFadeL 44, l44
 NFadeLm 45, l45 'Bumper
 NFadeL 45, l45a
 NFadeLm 46, l46 'Bumper
 NFadeL 46, l46a
 NFadeLm 47, l47 'Bumper
 NFadeL 47, l47a

 NFadeL 49, l49
 NFadeL 50, l50
 NFadeL 51, l51
 NFadeL 52, l52
 NFadeL 53, l53
 NFadeL 54, l54
 NFadeL 55, l55
 NFadeL 56, l56

'Solenoid Controlled Lamps

 NFadeLm 107, S107
 NFadeLm 107, S107a
 NFadeL 107, S107b

 NFadeObjm 118, P118, "dome2_0_greenON", "dome2_0_green"
 NFadeL 118, S118

 NFadeObjm 119, P119, "dome2_0_greenON", "dome2_0_green"
 NFadeL 119, S119

 NFadeL 120, S120

 NFadeLm 125, S125
 NFadeLm 125, S125a
 NFadeLm 125, S125b
 NFadeL 125, S125c

 NFadeObjm 126, P126, "dome2_0_yellowON", "dome2_0_yellow"
 NFadeObjm 126, P126a, "dome2_0_yellowON", "dome2_0_yellow"
 NFadeLm 126, S126
 NFadeL 126, S126a

 NFadeL 127, S127

 NFadeL 128, S128

 NFadeL 129, S129

 NFadeL 130, S130

 NFadeLm 131, S131
 NFadeL 131, S131a

 NFadeObjm 132, P132, "dome2_0_redON", "dome2_0_red"
 NFadeLm 132, S132
 NFadeL 132, S132a

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
  vpmTimer.PulseSw 62
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
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0:
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  vpmTimer.PulseSw 59
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
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0:
    End Select
    LStep = LStep + 1
End Sub

' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX, Rothbauerw and Herweh
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

' set position as table object and Vol + RndPitch manually

Sub PlaySoundAtVolPitch(sound, tableobj, Vol, RndPitch)
  PlaySound sound, 1, Vol, AudioPan(tableobj), RndPitch, 0, 0, 1, AudioFade(tableobj)
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

Sub PlaySoundAtBallAbsVol(sound, VolMult)
  PlaySound sound, 0, VolMult, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub

' requires rampbump1 to 7 in Sound Manager

Sub RandomBump(voladj, freq)
  Dim BumpSnd:BumpSnd= "rampbump" & CStr(Int(Rnd*7)+1)
  PlaySound BumpSnd, 0, Vol(ActiveBall)*voladj, AudioPan(ActiveBall), 0, freq, 0, 1, AudioFade(ActiveBall)
End Sub

' set position as bumperX and Vol manually. Allows rapid repetition/overlaying sound

Sub PlaySoundAtBumperVol(sound, tableobj, Vol)
  PlaySound sound, 1, Vol, AudioPan(tableobj), 0,0,1, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBOTBallZ(sound, BOT)
  PlaySound sound, 0, ABS(BOT.velz)/17, Pan(BOT), 0, Pitch(BOT), 1, 0, AudioFade(BOT)
End Sub

' play a looping sound at a location with volume
Sub PlayLoopSoundAtVol(sound, tableobj, Vol)
  PlaySound sound, -1, Vol, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
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

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
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

Function BallRollVol(ball) ' Calculates the Volume of the sound based on the ball speed
  BallRollVol = Csng(BallVel(ball) ^2 / (80000 - (79900 * Log(RollVol) / Log(100))))
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

'*** Determines if a Points (px,py) is inside a 4 point polygon A-D in Clockwise/CCW order

Function InRect(px,py,ax,ay,bx,by,cx,cy,dx,dy)
  Dim AB, BC, CD, DA
  AB = (bx*py) - (by*px) - (ax*py) + (ay*px) + (ax*by) - (ay*bx)
  BC = (cx*py) - (cy*px) - (bx*py) + (by*px) + (bx*cy) - (by*cx)
  CD = (dx*py) - (dy*px) - (cx*py) + (cy*px) + (cx*dy) - (cy*dx)
  DA = (ax*py) - (ay*px) - (dx*py) + (dy*px) + (dx*ay) - (dy*ax)

  If (AB <= 0 AND BC <=0 AND CD <= 0 AND DA <= 0) Or (AB >= 0 AND BC >=0 AND CD >= 0 AND DA >= 0) Then
    InRect = True
  Else
    InRect = False
  End If
End Function


'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 7 ' total number of balls
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
    End If
    Next
End Sub

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
  PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

'*****************************************
' ninuzzu's FLIPPER SHADOWS
'*****************************************

sub FlipperTimer_Timer()
  FlipperLSh.RotZ = LeftFlipper.currentangle
  FlipperRSh.RotZ = RightFlipper.currentangle

    LFLogo.RotY = LeftFlipper.CurrentAngle
    RFlogo.RotY = RightFlipper.CurrentAngle

    sw28p.RotY = sw28.CurrentAngle

End Sub

'*****************************************
' ninuzzu's BALL SHADOW
'*****************************************
Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5,BallShadow6,BallShadow7)

Sub BallShadowUpdate_timer()
    Dim BOT, bd
    BOT = GetBalls
    ' hide shadow of deleted balls
    If UBound(BOT)<(tnob-1) Then
        For bd = (UBound(BOT) + 1) to (tnob-1)
            BallShadow(bd).visible = 0
       Next
    End If
    ' exit the Sub if no balls on the table
    If UBound(BOT) = -1 Then Exit Sub
    ' render the shadow for each ball
    For bd = 0 to UBound(BOT)
        If BOT(bd).X < Table1.Width/2 Then
          BallShadow(bd).X = ((BOT(bd).X) - (Ballsize/6) + ((BOT(bd).X - (Table1.Width/2))/7)) + 6
        Else
           BallShadow(bd).X = ((BOT(bd).X) + (Ballsize/6) + ((BOT(bd).X - (Table1.Width/2))/7)) - 6
        End If
        ballShadow(bd).Y = BOT(bd).Y + 12
       If BOT(bd).Z > 20 Then
           BallShadow(bd).visible = 1
        Else
           BallShadow(bd).visible = 0
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
  PlaySound "fx_spinner", 0, VolSpin, AudioPan(Spinner), 0.25, 0, 0, 1, AudioFade(Spinner)
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
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

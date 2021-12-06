Option Explicit
Randomize

' Thalamus 2018-11-01 : Improved directional sounds
' !! NOTE : Table not verified yet !!
' Changed rom to im_183ve for fastflips memory address hack

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
Const VolRH     = 1    ' Rubber hits volume.
Const VolPi     = 1    ' Rubber pins volume.
Const VolTarg   = 1    ' Targets volume.
Const VolSpin   = 1.5  ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

' Const cGameName="im_183",UseSolenoids=1,UseLamps=0,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff", SCoin="coin"
Const cGameName="im_183ve",UseSolenoids=1,UseLamps=0,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff", SCoin="coin"

LoadVPM "01560000", "sam.VBS", 3.10

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

'*************************************************************
'Solenoid Call backs
'**********************************************************************************************************

SolCallback(1) = "solTrough"
SolCallback(2) = "solAutofire"
'SolCallback(3) = "MongerMagnet"
'SolCallback(4) = whiplash magnet
SolCallback(5) = "bsSaucer.SolOut"
SolCallback(6) = "orbitpost"
SolCallback(12) = "ClanePost"
SolCallback(19) = "Solmonger"
SolCallBack(20) = "SetLamp 120," 'PF light
SolCallback(21) = "SetLamp 121,"
SolCallback(22) = "SetLamp 122," 'PF light
SolCallback(23) = "SetLamp 123," 'PF light
SolCallback(25) = "SetLamp 125," 'monger toy
SolCallback(26) = "SetLamp 126,"
SolCallback(27) = "SetLamp 127," 'warmachine toy
SolCallback(29) = "SetLamp 129," 'wiplash toy
'SolCallback(30) = "SetLamp 130,"
SolCallback(31) = "SetLamp 131,"
SolCallback(32) = "SetLamp 132,"

SolCallback(15) = "SolLFlipper"
SolCallback(16) = "SolRFlipper"

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

Sub solTrough(Enabled)
  If Enabled Then
    bsTrough.ExitSol_On
    vpmTimer.PulseSw 22
  End If
 End Sub

Sub solAutofire(Enabled)
  If Enabled Then
    PlungerIM.AutoFire
  End If
 End Sub

Sub ClanePost(Enabled)
  If Enabled Then
    ClaneUpPost.Isdropped=false
    playsound SoundFX("Diverter",DOFContactors) ' TODO
  Else
    ClaneUpPost.Isdropped=true

  End If
 End Sub

Sub orbitpost(Enabled)
  If Enabled Then
    UpPost.Isdropped=false
    playsound SoundFX("Diverter",DOFContactors) ' TODO
  Else
    UpPost.Isdropped=true

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

Dim bsTrough, bsSaucer, Mag1, Mag2

Sub Table1_Init
  vpmInit Me
  On Error Resume Next
    With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine = "IronMan (Stern 2010)"&chr(13)&"You Suck"
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

  PinMAMETimer.Interval = PinMAMEInterval
  PinMAMETimer.Enabled = 1
    vpmNudge.TiltSwitch=-7
    vpmNudge.Sensitivity=2
    vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingshot)

    InitVpmFFlipsSAM

    Set bsTrough = New cvpmBallStack
    bsTrough.InitSw 0, 21, 20, 19, 18, 0, 0, 0
    bsTrough.InitKick BallRelease, 90, 8
    bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
    bsTrough.Balls = 4

  Set bsSaucer = New cvpmBallStack
    bsSaucer.InitSaucer sw10, 10, 180, 20
    bsSaucer.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
    bsSaucer.KickForceVar = 2.5

  Set mag1= New cvpmMagnet
  With mag1
    .InitMagnet Magnet1, 22
    .GrabCenter = False
    .solenoid=3
    .CreateEvents "mag1"
  End With

  Set mag2= New cvpmMagnet
  With mag2
    .InitMagnet Magnet2, 22
    .GrabCenter = False
    .solenoid=4
    .CreateEvents "mag2"
  End With

 ClaneUpPost.Isdropped = 1
 UpPost.Isdropped = 1

sw5.IsDropped = 1
switchframe.IsDropped = 1

End Sub

 Sub Table_Paused:Controller.Pause = 1:End Sub
 Sub Table_unPaused:Controller.Pause = 0:End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)
  If KeyDownHandler(keycode) Then Exit Sub
  If keycode = PlungerKey Then Plunger.Pullback:playsoundAtVol"plungerpull", Plunger, 1

End Sub

Sub Table1_KeyUp(ByVal KeyCode)
  If KeyUpHandler(keycode) Then Exit Sub
  If keycode = PlungerKey Then Plunger.Fire:PlaySoundAtVol"plunger", Plunger, 1
  If Keycode = StartGameKey Then Controller.Switch(16) = 0
End Sub

  Dim PlungerIM
     ' Impulse Plunger
    Const IMPowerSetting = 50
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
Sub Drain_Hit:bsTrough.addball me : playsoundAtVol"drain" , Drain, 1: End Sub
Sub sw10_Hit:bsSaucer.AddBall 0 : playsoundAtVol "popper_ball", ActiveBall, 1: End Sub

'Wire Triggers
Sub sw7_Hit:Controller.Switch(7) = 1:PlaySoundAtVol "rollover", ActiveBall, VolRol:End Sub
Sub sw7_UnHit:Controller.Switch(7) = 0:End Sub
Sub sw9_Hit:Controller.Switch(9) = 1:PlaySoundAtVol "rollover", ActiveBall, VolRol:End Sub
Sub sw9_UnHit:Controller.Switch(9) = 0:End Sub
Sub sw23_Hit:Controller.Switch(23) = 1:PlaySoundAtVol "rollover", ActiveBall, VolRol:End Sub
Sub sw23_UnHit:Controller.Switch(23) = 0:End Sub
Sub sw24_Hit:Controller.Switch(24) = 1:PlaySoundAtVol "rollover", ActiveBall, VolRol:End Sub
Sub sw24_UnHit:Controller.Switch(24) = 0:End Sub
Sub sw25_Hit:Controller.Switch(25) = 1:PlaySoundAtVol "rollover", ActiveBall, VolRol:End Sub
Sub sw25_UnHit:Controller.Switch(25) = 0:End Sub
Sub sw28_Hit:Controller.Switch(28) = 1:PlaySoundAtVol "rollover", ActiveBall, VolRol:End Sub
Sub sw28_UnHit:Controller.Switch(28) = 0:End Sub
Sub sw29_Hit:Controller.Switch(29) = 1:PlaySoundAtVol "rollover", ActiveBall, VolRol:End Sub
Sub sw29_UnHit:Controller.Switch(29) = 0:End Sub
Sub sw38_Hit:Controller.Switch(38) = 1:PlaySoundAtVol "rollover", ActiveBall, VolRol:End Sub
Sub sw38_UnHit:Controller.Switch(38) = 0:End Sub
Sub sw39_Hit:Controller.Switch(39) = 1:PlaySoundAtVol "rollover", ActiveBall, VolRol:End Sub
Sub sw39_UnHit:Controller.Switch(39) = 0:End Sub

'Spinners
Sub sw11_Spin:vpmTimer.PulseSw 11 : playsoundAtVol"fx_spinner" , sw11, VolSpin: End Sub
Sub sw13_Spin:vpmTimer.PulseSw 13 : playsoundAtVol"fx_spinner" , sw13, VolSpin: End Sub
Sub sw14_Spin:vpmTimer.PulseSw 14 : playsoundAtVol"fx_spinner" , sw14, VolSpin: End Sub

'RAmp Gate Triggers
Sub sw12_Hit:vpmTimer.PulseSw 12:End Sub
Sub sw37_Hit:vpmTimer.PulseSw 37:End Sub
Sub sw43_Hit:vpmTimer.PulseSw 43:End Sub
Sub sw49_Hit:vpmTimer.PulseSw 49:End Sub

'Bumpers
Sub Bumper1_Hit : vpmTimer.PulseSw(31) : playsoundAtVol SoundFX("fx_bumper1",DOFContactors), ActiveBall, VolBump: End Sub
Sub Bumper2_Hit : vpmTimer.PulseSw(30) : playsoundAtVol SoundFX("fx_bumper1",DOFContactors), ActiveBall, VolBump: End Sub
Sub Bumper3_Hit : vpmTimer.PulseSw(32) : playsoundAtVol SoundFX("fx_bumper1",DOFContactors), ActiveBall, VolBump: End Sub

'Stand Up Targets
Sub sw33_Hit:vpmTimer.PulseSw 33:End Sub
Sub sw34_Hit:vpmTimer.PulseSw 34:End Sub
Sub sw35_Hit:vpmTimer.PulseSw 35:End Sub
Sub sw36_Hit:vpmTimer.PulseSw 36:End Sub
Sub sw40_Hit:vpmTimer.PulseSw 40:End Sub
Sub sw41_Hit:vpmTimer.PulseSw 41:End Sub
Sub sw42_Hit:vpmTimer.PulseSw 42:End Sub
Sub sw44_Hit:vpmTimer.PulseSw 44:End Sub
Sub sw45_Hit:vpmTimer.PulseSw 45:End Sub
Sub sw46_Hit:vpmTimer.PulseSw 46:End Sub
Sub sw47_Hit:vpmTimer.PulseSw 47:End Sub
Sub sw48_Hit:vpmTimer.PulseSw 48:End Sub
Sub sw50_Hit:vpmTimer.PulseSw 50:End Sub

'Generic Sounds
Sub Trigger1_hit:PlaySoundAtVol "fx_ballrampdrop", ActiveBall, 1:End Sub
Sub Trigger2_hit:PlaySoundAtVol "fx_ballrampdrop", ActiveBall, 1:End Sub
Sub Trigger3_hit:PlaySoundAtVol "fx_ballrampdrop", ActiveBall, 1:End Sub


  '***************************************************************************
 '    monger Animation
 '***************************************************************************

'1 monger down
'3 monger up
'4 monger left shoulder
'5 monger legs
Sub sw5_Hit:vpmTimer.PulseSw 5:PlaySoundAtVol "target", ActiveBall, 1:End Sub
'6 monger rt shoulder


Dim monger, mongerPos, mongerDir, mongerFlash
mongerDir = 0:mongerPos = 0:mongerFlash = 0

 Sub Solmonger(Enabled)
     If Enabled Then
         If mongerDir = 0 Then
             Controller.Switch(3) = 0
             Controller.Switch(1) = 1
             mongerClose.Enabled = 0
             mongerOpen.Enabled = 1
             mongerOpen_Timer
             mongerDir = 1
         Else
             Controller.Switch(1) = 0
             Controller.Switch(3) = 1
             mongerOpen.Enabled = 0
             mongerClose.Enabled = 1
             mongerClose_Timer
             mongerDir = 0
         End If
     End If
 End Sub

 Sub mongerOpen_Timer()
     Updatemonger
     mongerPos = mongerPos + 1

     If mongerPos> 41 Then
         mongerPos = 41
         mongerOpen.Enabled = 0
     End If
 End Sub

 Sub mongerClose_Timer()
     Updatemonger
     mongerPos = mongerPos - 1

     If mongerPos <0 Then
         mongerPos = 0
         mongerClose.Enabled = 0
     End If
 End Sub

 Sub Updatemonger
     Select Case mongerPos
Case 0:Primitive32.Z=300:Primitive31.Z=300::Primitive2.Z=90:mongerAnim.Enabled = 0:sw3.IsDropped = 0:switchframe.IsDropped = 0:sw5.IsDropped = 0' monger1.IsDropped = 0:monger2.IsDropped = 1:
Case 1:Primitive32.Z=292.5:Primitive31.Z=292.5:Primitive2.Z=82.5
Case 2:Primitive32.Z=285:Primitive31.Z=285:Primitive2.Z=75
Case 3:Primitive32.Z=277.5:Primitive31.Z=277.5:Primitive2.Z=67.5
Case 4:Primitive32.Z=270:Primitive31.Z=270:Primitive2.Z=60
Case 5:Primitive32.Z=262.5:Primitive31.Z=262.5:Primitive2.Z=52.5
Case 6:Primitive32.Z=255:Primitive31.Z=255:Primitive2.Z=45
Case 7:Primitive32.Z=247.5:Primitive31.Z=247.5:Primitive2.Z=37.5
Case 8:Primitive32.Z=240:Primitive31.Z=240:Primitive2.Z=30
Case 9:Primitive32.Z=232.5:Primitive31.Z=232.5:Primitive2.Z=22.5
Case 10:Primitive32.Z=225:Primitive31.Z=225:Primitive2.Z=15
Case 11:Primitive32.Z=217.5:Primitive31.Z=217.5:Primitive2.Z=7.5
Case 12:Primitive32.Z=210:Primitive31.Z=210:Primitive2.Z=0
Case 13:Primitive32.Z=202.5:Primitive31.Z=202.5:Primitive2.Z=-7.5
Case 14:Primitive32.Z=195:Primitive31.Z=195:Primitive2.Z=-15
Case 15:Primitive32.Z=187.5:Primitive31.Z=187.5:Primitive2.Z=-22.5
Case 16:Primitive32.Z=180:Primitive31.Z=180:Primitive2.Z=-30
Case 17:Primitive32.Z=172.5:Primitive31.Z=172.5:Primitive2.Z=-37.5
Case 18:Primitive32.Z=165:Primitive31.Z=165:Primitive2.Z=-45
Case 19:Primitive32.Z=157.5:Primitive31.Z=157.5:Primitive2.Z=-52.5
Case 20:Primitive32.Z=150:Primitive31.Z=150:Primitive2.Z=-60
Case 21:Primitive32.Z=142.5:Primitive31.Z=142.5:Primitive2.Z=-67.5
Case 22:Primitive32.Z=135:Primitive31.Z=135:Primitive2.Z=-75
Case 23:Primitive32.Z=127.5:Primitive31.Z=127.5:Primitive2.Z=-82.5
Case 24:Primitive32.Z=120:Primitive31.Z=120:Primitive2.Z=-90
Case 25:Primitive32.Z=112.5:Primitive31.Z=112.5:Primitive2.Z=-97.5
Case 26:Primitive32.Z=105:Primitive31.Z=105:Primitive2.Z=-105
Case 27:Primitive32.Z=97.5:Primitive31.Z=97.5:Primitive2.Z=-112.5
Case 28:Primitive32.Z=90:Primitive31.Z=90:Primitive2.Z=-120
Case 29:Primitive32.Z=82.5:Primitive31.Z=82.5:Primitive2.Z=-127.5
Case 30:Primitive32.Z=75:Primitive31.Z=75:Primitive2.Z=-135
Case 31:Primitive32.Z=67.5:Primitive31.Z=67.5:Primitive2.Z=-142.5
Case 32:Primitive32.Z=60:Primitive31.Z=60:Primitive2.Z=-150
Case 33:Primitive32.Z=52.5:Primitive31.Z=52.5:Primitive2.Z=-157.5
Case 34:Primitive32.Z=45:Primitive31.Z=45:Primitive2.Z=-165
Case 35:Primitive32.Z=37.5:Primitive31.Z=37.5:Primitive2.Z=-172.5
Case 36:Primitive32.Z=30:Primitive31.Z=30:Primitive2.Z=-180
Case 37:Primitive32.Z=22.5:Primitive31.Z=22.5:Primitive2.Z=-187.5
Case 38:Primitive32.Z=15:Primitive31.Z=15:Primitive2.Z=-195
Case 39:Primitive32.Z=7.5:Primitive31.Z=7.5:Primitive2.Z=-202.5
Case 40:Primitive32.Z=0:Primitive31.Z=0:Primitive2.Z=-210:mongerAnim.Enabled = 0:sw3.IsDropped = 1:switchframe.IsDropped = 1:sw5.IsDropped = 1
     End Select
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
NFadeL 28, l28
NFadeL 29, l29
NFadeL 30, l30
NFadeL 31, l31
NFadeL 32, l32
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
NFadeL 46, l46
NFadeL 47, l47
NFadeL 48, l48
NFadeL 49, l49
NFadeL 50, l50
NFadeL 51, l51
NFadeL 52, l52
NFadeL 53, l53
NFadeL 54, l54
NFadeL 55, l55
NFadeL 56, l56
NFadeL 57, l57
NFadeL 58, l58
NFadeL 59, l59
NFadeL 60, l60 'bumper2

NFadeL 61, l61 'bumper1

NFadeL 62, l62 'bumper3

NFadeL 63, l63


'Solenoid Controlled Lamps
NFadeL 120, S120

NFadeObjm 121, P121, "dome2_0_orangeON", "dome2_0_orange"
NFadeL 121, S121a

NFadeL 122, S122

NFadeL 123, S123

NFadeL 125, S125 'monger primitive
NFadeL 125, S125a

NFadeObjm 126, P126, "dome2_0_orangeON", "dome2_0_orange"
NFadeL 126, S126a

NFadeLm 127, S127
NFadeL 127, S127a

NFadeLm 129, S129
NFadeLm 129, S129a
NFadeLm 129, S129b
NFadeL 129, S129c

'NFadeLm 130, S130

NFadeObjm 131, P131, "dome2_0_redON", "dome2_0_red"
NFadeObjm 131, P131a, "dome2_0_redON", "dome2_0_red"
NFadeLm 131, S131a
NFadeL 131, S131b

NFadeObjm 132, P132, "dome2_0_redON", "dome2_0_red"
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



'*********************************************************************
'
'*********************************************************************


'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
  vpmTimer.PulseSw 27
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
  vpmTimer.PulseSw 26
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

    LFLogo.RotY = LeftFlipper.CurrentAngle
    RFlogo.RotY = RightFlipper.CurrentAngle

    sw12p.RotZ = -(sw12.currentangle)
    sw37p.RotZ = -(sw37.currentangle)
    sw43p.RotZ = -(sw43.currentangle)
    sw49p.RotZ = -(sw49.currentangle)

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
    PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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

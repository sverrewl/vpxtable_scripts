' Black Knight / IPD No. 310 / Williams, November, 1980 / 4 Players
' VPX 1.3.1 by JPSalas 2016
' Lights based on Destruk's script
' Thanks to Leo and Jolo for pictures and testing the table.
' And a special thanks to Classic Playfields for sending us the pictures of the playfields.

' Thalamus 2018-07-19
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' Thalamus 2018-08-10 : Improved directional sounds

Option Explicit
Randomize

' Options
' Volume devided by - lower gets higher sound

Const VolDiv = 2000

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01560000", "S7.VBS", 3.26

'********************
'Standard definitions
'********************

Const UseSolenoids = 2
Const UseLamps = 0
Const UseSync = 0
Const HandleMech = 0

' Standard Soundsfx_
Const SSolenoidOn = "fx_Solenoid"
Const SSolenoidOff = ""
Const SCoin = "fx_Coin"

Dim VarHidden
If Table1.ShowDT = true then
    VarHidden = 1
    For each x in aReels
        x.Visible = 1
    Next
else
    VarHidden = 0
    For each x in aReels
        x.Visible = 0
    Next
    lrail.Visible = 0
    rrail.Visible = 0
end if

if B2SOn = true then VarHidden = 1

Dim bsTrough, bsLock, bsSaucer, dtLL, dtLR, dtUL, dtUR, LMAG, RMAG
Dim x, i, j, k

'************
' Table init.
'************
Const cGameName = "bk_l4"

Sub Table1_Init
    vpmInit Me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Black Knight - by Williams 1980" & vbNewLine & "VPX table by JPSalas v1.3.1"
        .Games(cGameName).Settings.Value("rol") = 0 '1= rotated display, 0= normal
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = VarHidden
        On Error Resume Next
        .SolMask(0) = 0
        vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the timer to renable all the solenoids after 2 seconds
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
        On Error Goto 0
    End With

    ' Nudging
    vpmNudge.TiltSwitch = 46
    vpmNudge.Sensitivity = 1
    vpmNudge.TiltObj = Array(Bumper1, LeftSlingshot, RightSlingshot)

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 20, 17, 18, 19, 0, 0, 0, 0
        .InitKick BallRelease, 90, 4
        .InitEntrySnd "fx_Solenoid", "fx_Solenoid"
        .InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .Balls = 3
        .IsTrough = 1
    End With

    ' Lock
    Set bsLock = New cvpmBallStack
    With bsLock
        .InitSw 0, 41, 42, 43, 0, 0, 0, 0
        .InitKick LockOut, 194, 1
        .InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_kicker", DOFContactors)
    End With

    ' Lower Eject Hole
    Set bsSaucer = New cvpmBallStack
    With bsSaucer
        .InitSaucer sw24, 24, 160, 8
        .InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_kicker", DOFContactors)
    End With

    ' Lower Left droptargets
    Set dtLL = New cvpmDropTarget
    With dtLL
        .InitDrop Array(sw25, sw26, sw27), Array(25, 26, 27)
        .initsnd SoundFX("fx_droptarget", DOFContactors), SoundFX("fx_resetdrop", DOFContactors)
        .CreateEvents "dtLL"
    End With

    ' Lower Right droptargets
    Set dtLR = New cvpmDropTarget
    With dtLR
        .InitDrop Array(sw29, sw30, sw31), Array(29, 30, 31)
        .initsnd SoundFX("fx_droptarget", DOFContactors), SoundFX("fx_resetdrop", DOFContactors)
        .CreateEvents "dtLR"
    End With

    ' Upper Left droptargets
    Set dtUL = New cvpmDropTarget
    With dtUL
        .InitDrop Array(sw33, sw34, sw35), Array(33, 34, 35)
        .initsnd SoundFX("fx_droptarget", DOFContactors), SoundFX("fx_resetdrop", DOFContactors)
        .CreateEvents "dtUL"
    End With

    ' Upper right droptargets
    Set dtUR = New cvpmDropTarget
    With dtUR
        .InitDrop Array(sw37, sw38, sw39), Array(37, 38, 39)
        .initsnd SoundFX("fx_droptarget", DOFContactors), SoundFX("fx_resetdrop", DOFContactors)
        .CreateEvents "dtUR"
    End With

    ' Left Magnet
    Set LMAG = New cvpmMagnet
    With LMAG
        .InitMagnet MagnetL, 7
        .Solenoid = 10
        .CreateEvents "LMAG"
    End With

    ' Right Magnet
    Set RMAG = New cvpmMagnet
    With RMAG
        .InitMagnet MagnetR, 7
        .Solenoid = 9
        .CreateEvents "RMAG"
    End With

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

    SolGi 0
End Sub

Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub

'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
    If KeyCode = LeftMagnaSave Then Controller.Switch(10) = 1
    If KeyCode = RightMagnaSave Then Controller.Switch(9) = 1
    If keycode = PlungerKey Then PlaySoundAt "fx_PlungerPull", Plunger:Plunger.Pullback
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If KeyCode = LeftMagnaSave Then Controller.Switch(10) = 0
    If KeyCode = RightMagnaSave Then Controller.Switch(9) = 0
    If keycode = PlungerKey Then PlaySoundAt "fx_plunger", Plunger:Plunger.Fire
    If vpmKeyUp(keycode) Then Exit Sub
End Sub

'*********
' Switches
'*********

' Slings & div switches

Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    PlaySound SoundFX("fx_slingshot", DOFContactors), 0, 1, -0.05, 0.05
    ' PlaySoundAt SoundFX("fx_slingshot", DOFContactors), xxx ' TODO : Find a object near leftsling
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 21
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
    PlaySound SoundFX("fx_slingshot", DOFContactors), 0, 1, 0.05, 0.05
    ' PlaySoundAt SoundFX("fx_slingshot", DOFContactors), xxx ' TODO : Find a object near rightsling
    RightSling4.Visible = 1
    Remk.RotX = 26
    RStep = 0
    vpmTimer.PulseSw 22
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
Sub Bumper1_Hit:vpmTimer.PulseSw 36:PlaySoundAtBumperVol SoundFX("fx_bumper", DOFContactors), Bumper1, 2:End Sub

' Drain holes, vuks & saucers
Sub Drain_Hit:PlaysoundAt "fx_drain", Drain:bsTrough.AddBall Me:End Sub
Sub LockMech_Hit:PlaySoundAt "fx_kicker_enter", LockMech:bsLock.AddBall Me:End Sub
Sub sw24_Hit:PlaySoundAt "fx_metalhit", sw24:bsSaucer.AddBall 0:End Sub

' Rollovers & Ramp Switches
Sub sw11_Hit:Controller.Switch(11) = 1:PlaySoundAt "fx_sensor", sw11:End Sub
Sub sw11_UnHit:Controller.Switch(11) = 0:End Sub

Sub sw16_Hit:Controller.Switch(16) = 1:PlaySoundAt "fx_sensor", sw16:End Sub
Sub sw16_UnHit:Controller.Switch(16) = 0:End Sub

Sub sw15_Hit:Controller.Switch(15) = 1:PlaySoundAt "fx_sensor", sw15:End Sub
Sub sw15_UnHit:Controller.Switch(15) = 0:End Sub

Sub sw12_Hit:Controller.Switch(12) = 1:PlaySoundAt "fx_sensor", sw12:End Sub
Sub sw12_UnHit:Controller.Switch(12) = 0:End Sub

Sub sw23_Hit:Controller.Switch(23) = 1:PlaySoundAt "fx_sensor", sw23:End Sub
Sub sw23_UnHit:Controller.Switch(23) = 0:End Sub

Sub sw14_Hit:Controller.Switch(14) = 1:PlaySoundAt "fx_sensor", sw14:End Sub
Sub sw14_UnHit:Controller.Switch(14) = 0:End Sub

Sub sw44_Hit:Controller.Switch(44) = 1:PlaySoundAt "fx_sensor", sw44:End Sub
Sub sw44_UnHit:Controller.Switch(44) = 0:End Sub

Sub sw45_Hit:Controller.Switch(45) = 1:PlaySoundAt "fx_sensor", sw45:End Sub
Sub sw45_UnHit:Controller.Switch(45) = 0:End Sub

Sub sw13_Spin:vpmTimer.PulseSw 13:PlaySoundAt SoundFX("fx_spinner", DOFContactors), sw13:End Sub

Sub UpperEnter_Hit
    Dim a
    a = ABS(ActiveBall.VelY)
    UpperEnter.Destroyball
    UpperExit.CreateBall
    UpperExit.Kick 180, a
End Sub

'*********
'Solenoids
'*********

SolCallback(1) = "bsTrough.SolIn"
SolCallback(2) = "dtLL.SolDropUp"
SolCallback(3) = "dtLR.SolDropUp"
SolCallback(4) = "dtUL.SolDropUp"
SolCallback(5) = "dtUR.SolDropUp"
SolCallback(6) = "bsTrough.SolOut"
SolCallback(7) = "bsLock.SolOut"
SolCallback(8) = "bsSaucer.SolOut"
SolCallback(11) = "SolGi"
solcallback(15) = "vpmsolsound SoundFX(""fx_knocker"",DOFKnocker),"
SolCallback(23) = "vpmNudge.SolGameOn"

Sub SolGi(Enabled)
    If Enabled Then
        GiOFF
    Else
        GiON
    end if
end sub

'**************
' Flipper Subs
'**************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("fx_flipperup", DOFContactors), LeftFlipper
        LeftFlipper.RotateToEnd
        LeftFlipper1.RotateToEnd
    Else
        PlaySoundAt SoundFX("fx_flipperdown", DOFContactors), LeftFlipper
        LeftFlipper.RotateToStart
        LeftFlipper1.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("fx_flipperup", DOFContactors), RightFlipper
        RightFlipper.RotateToEnd
        RightFlipper1.RotateToEnd
    Else
        PlaySoundAt SoundFX("fx_flipperdown", DOFContactors), RightFlipper
        RightFlipper.RotateToStart
        RightFlipper1.RotateToStart
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

Sub Rightflipper1_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, 0.1, 0.15
End Sub

' Extra Lights on the backglass
'     1, "same player shoots again"
'     2, "ball in play"
'     3, "tilt"
'     4, "game over"
'     5, "match"
'     6, "highscore"
'     7, "credits"
'     8, "bonus ball time"

'***************************************************
'       JP's VP10 Fading Lamps & Flashers
'       Based on PD's Fading Light System
' SetLamp 0 is Off
' SetLamp 1 is On
' fading for non opacity objects is 4 steps
'***************************************************

Dim LampState(200), FadingLevel(200)
Dim FlashSpeedUp(200), FlashSpeedDown(200), FlashMin(200), FlashMax(200), FlashLevel(200), FlashRepeat(200)

InitLamps()             ' turn off the lights and flashers and reset them to the default parameters
LampTimer.Interval = 10 ' lamp fading speed
LampTimer.Enabled = 1

' Lamp & Flasher Timers

Sub LampTimer_Timer()
    Dim chgLamp, num, chg, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp)Then
        For ii = 0 To UBound(chgLamp)
            LampState(chgLamp(ii, 0)) = chgLamp(ii, 1)       'keep the real state in an array
            FadingLevel(chgLamp(ii, 0)) = chgLamp(ii, 1) + 4 'actual fading step
        Next
    End If
    If VarHidden Then
        UpdateLeds
    End If
    UpdateLamps
    GIUpdate
    RollingUpdate
End Sub

Sub UpdateLamps()

    'backdrop lights
    If VarHidden Then
     NFadeT 1, l1, "same player shoots again"
     NFadeT 2, l2, "ball in play"
     NFadeT 3, l3, "tilt"
     NFadeT 4, l4, "game over"
     NFadeT 5, l5, "match"
     NFadeT 6, l6, "highscore"
     NFadeTm 7, l7a, "credits"
     NFadeT 8, l8, "bonus ball time"
    End If

    NFadeL 7, l7
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
    NFadeLm 36, l36a
    NFadeL 36, l36
    NFadeL 37, l37
    NFadeL 38, l38
    NFadeL 39, l39
    NFadeL 40, l40
    NFadeL 41, l41
    NFadeL 42, l42
    'NFadeL 43, l43
    'NFadeL 44, l44
    'NFadeL 45, l45
    'NFadeL 46, l46
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
    NFadeL 60, l60
    NFadeL 61, l61
    NFadeL 62, l62
    NFadeL 63, l63
    NFadeL 64, l64
End Sub

' div lamp subs

Sub InitLamps()
    Dim x
    For x = 0 to 200
        LampState(x) = 0        ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4      ' used to track the fading state
        FlashSpeedUp(x) = 0.2   ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.1 ' slower speed when turning off the flasher
        FlashMax(x) = 1         ' the maximum value when on, usually 1
        FlashMin(x) = 0         ' the minimum value when off, usually 0
        FlashLevel(x) = 0       ' the intensity of the flashers, usually from 0 to 1
        FlashRepeat(x) = 20     ' how many times the flash repeats
    Next
End Sub

Sub AllLampsOff
    Dim x
    For x = 0 to 200
        SetLamp x, 0
    Next
End Sub

Sub SetLamp(nr, value)
    If value <> LampState(nr)Then
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
            FlashLevel(nr) = FlashLevel(nr)- FlashSpeedDown(nr)
            If FlashLevel(nr) < FlashMin(nr)Then
                FlashLevel(nr) = FlashMin(nr)
                FadingLevel(nr) = 0 'completely off
            End if
            Object.IntensityScale = FlashLevel(nr)
        Case 5 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
            If FlashLevel(nr) > FlashMax(nr)Then
                FlashLevel(nr) = FlashMax(nr)
                FadingLevel(nr) = 1 'completely on
            End if
            Object.IntensityScale = FlashLevel(nr)
    End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it doesn't change anything, it just follows the main flasher
    Select Case FadingLevel(nr)
        Case 4, 5
            Object.IntensityScale = FlashLevel(nr)
    End Select
End Sub

Sub FlashBlink(nr, object)
    Select Case FadingLevel(nr)
        Case 4 'off
            FlashLevel(nr) = FlashLevel(nr)- FlashSpeedDown(nr)
            If FlashLevel(nr) < FlashMin(nr)Then
                FlashLevel(nr) = FlashMin(nr)
                FadingLevel(nr) = 0 'completely off
            End if
            Object.IntensityScale = FlashLevel(nr)
            If FadingLevel(nr) = 0 AND FlashRepeat(nr)Then 'repeat the flash
                FlashRepeat(nr) = FlashRepeat(nr)-1
                If FlashRepeat(nr)Then FadingLevel(nr) = 5
            End If
        Case 5 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
            If FlashLevel(nr) > FlashMax(nr)Then
                FlashLevel(nr) = FlashMax(nr)
                FadingLevel(nr) = 1 'completely on
            End if
            Object.IntensityScale = FlashLevel(nr)
            If FadingLevel(nr) = 1 AND FlashRepeat(nr)Then FadingLevel(nr) = 4
    End Select
End Sub

' Desktop Objects: Reels & texts (you may also use lights on the desktop)

' Reels

Sub FadeR(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.SetValue 1:FadingLevel(nr) = 6                   'fading to off...
        Case 5:object.SetValue 0:FadingLevel(nr) = 1                   'ON
        Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1             'wait
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

Sub NFadeTm(nr, object, message)
    Select Case FadingLevel(nr)
        Case 4:object.Text = ""
        Case 5:object.Text = message
    End Select
End Sub


'********************
' General Ilumination
'********************

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

Dim OldGiState
OldGiState = -1 'start witht he Gi off

Sub GIUpdate
    Dim tmp, obj
    tmp = Getballs
    If UBound(tmp) <> OldGiState Then
        OldGiState = Ubound(tmp)
        If UBound(tmp) = -1 Then
            GiOff
        Else
            GiOn
        End If
    End If
End Sub

'******************************
' Diverse Collection Hit Sounds
'******************************

Sub aMetals_Hit(idx):PlaySound "fx_MetalHit2", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Bands_Hit(idx):PlaySound "fx_rubber_band", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Posts_Hit(idx):PlaySound "fx_rubber_post", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Pins_Hit(idx):PlaySound "fx_rubber_pin", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aWoods_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub

 '***************
 ' LED Handling
 '***************
 'Modified version of Scapino's LED code for Fathom
 '
 Dim SevenDigitOutput(32)
 Dim DisplayPatterns(11)

 'Binary/Hex Pattern Recognition Array
 DisplayPatterns(0) = 0    '0000000 Blank
 DisplayPatterns(1) = 63   '0111111 zero
 DisplayPatterns(2) = 6    '0000110 one
 DisplayPatterns(3) = 91   '1011011 two
 DisplayPatterns(4) = 79   '1001111 three
 DisplayPatterns(5) = 102  '1100110 four
 DisplayPatterns(6) = 109  '1101101 five
 DisplayPatterns(7) = 125  '1111101 six
 DisplayPatterns(8) = 7    '0000111 seven
 DisplayPatterns(9) = 127  '1111111 eight
 DisplayPatterns(10) = 111 '1101111 nine

 'Assign 7-digit output to reels
 Set SevenDigitOutput(0) = P3D1
 Set SevenDigitOutput(1) = P3D2
 Set SevenDigitOutput(2) = P3D3
 Set SevenDigitOutput(3) = P3D4
 Set SevenDigitOutput(4) = P3D5
 Set SevenDigitOutput(5) = P3D6
 Set SevenDigitOutput(6) = P3D7

 Set SevenDigitOutput(7) = P4D1
 Set SevenDigitOutput(8) = P4D2
 Set SevenDigitOutput(9) = P4D3
 Set SevenDigitOutput(10) = P4D4
 Set SevenDigitOutput(11) = P4D5
 Set SevenDigitOutput(12) = P4D6
 Set SevenDigitOutput(13) = P4D7

 Set SevenDigitOutput(14) = P1D1
 Set SevenDigitOutput(15) = P1D2
 Set SevenDigitOutput(16) = P1D3
 Set SevenDigitOutput(17) = P1D4
 Set SevenDigitOutput(18) = P1D5
 Set SevenDigitOutput(19) = P1D6
 Set SevenDigitOutput(20) = P1D7

 Set SevenDigitOutput(21) = P2D1
 Set SevenDigitOutput(22) = P2D2
 Set SevenDigitOutput(23) = P2D3
 Set SevenDigitOutput(24) = P2D4
 Set SevenDigitOutput(25) = P2D5
 Set SevenDigitOutput(26) = P2D6
 Set SevenDigitOutput(27) = P2D7

 Set SevenDigitOutput(28) = CrD2
 Set SevenDigitOutput(29) = CrD1
 Set SevenDigitOutput(30) = BaD2
 Set SevenDigitOutput(31) = BaD1

 Sub UpdateLeds ' 7-Digit output
     On Error Resume Next
     Dim ChgLED, ii, stat, TempCount

     ChgLED = Controller.ChangedLEDs(&HFF, &HFFFF) 'hex of binary (display 111111, or first 6 digits)

     If Not IsEmpty(ChgLED) Then
         For ii = 0 To UBound(ChgLED)
             stat = chgLED(ii, 2)

             For TempCount = 0 to 10
                 If stat = DisplayPatterns(TempCount) OR stat = (DisplayPatterns(TempCount) + 128) then
                     SevenDigitOutput(chgLED(ii, 0) ).SetValue(TempCount)
                 End If
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

Const tnob = 10 ' total number of balls
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


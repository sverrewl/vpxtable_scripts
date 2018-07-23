' Skateball - Bally 1980
' IPD No. 2170 / April 30, 1980 / 4 Players
' VPX - version by JPSalas 2017, version 1.0.0

Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01550000", "Bally.vbs", 3.26

Dim bsTrough, dtBankL, dtBankM, dtBankT, bsSaucer, x

Const cGameName = "skatebll"

Const UseSolenoids = 2
Const UseLamps = 0
Const UseGI = 0
Const UseSync = 0 'set it to 1 if the table runs too fast
Const HandleMech = 0

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

' Standard Sounds
Const SSolenoidOn = "fx_Solenoid"
Const SSolenoidOff = ""
Const SCoin = "fx_Coin"

'************
' Table init.
'************

Sub table1_Init
    vpmInit me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Future Spa - Bally 1979" & vbNewLine & "VPX table by JPSalas v.1.0.1"
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = VarHidden
        .Games(cGameName).Settings.Value("rol") = 0 '1= rotated display, 0= normal
        '.SetDisplayPosition 0,0, GetPlayerHWnd 'restore dmd window position
        On Error Resume Next
        Controller.SolMask(0) = 0
        vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the Timer to renable all the solenoids after 2 seconds
        Controller.Run GetPlayerHWnd
        On Error Goto 0
    End With

    ' Nudging
    vpmNudge.TiltSwitch = 7
    vpmNudge.Sensitivity = 3
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingshot)

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitNoTrough BallRelease, 8, 80, 6
        .InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    End With

    ' Drop targets
    set dtbankL = new cvpmdroptarget
    With dtbankL
        .initdrop array(sw24, sw23, sw22, sw21, sw20), array(24, 23, 22, 21, 20)
        .initsnd SoundFX("", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
    End With

    set dtbankM = new cvpmdroptarget
    With dtbankM
        .initdrop array(sw4, sw3, sw2), array(4, 3, 2)
        .initsnd SoundFX("", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
    End With

    set dtbankT = new cvpmdroptarget
    With dtbankT
        .initdrop array(sw28, sw27, sw26), array(28, 27, 26)
        .initsnd SoundFX("", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
    End With

    ' Saucer
    Set bsSaucer = New cvpmBallStack
    With bsSaucer
        .InitSaucer sw5, 5, 195, 17
        .InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .KickForceVar = 2
        .KickAngleVar = 3
    End With

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1
End Sub

Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub

'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If keycode = PlungerKey Then PlaySound "fx_PlungerPull", 0, 1, 0.1, 0.25:Plunger.Pullback
    If keycode = RightFlipperKey Then Controller.Switch(32) = True
    If vpmKeyDown(keycode)Then Exit Sub
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If vpmKeyUp(keycode)Then Exit Sub
    If keycode = PlungerKey Then PlaySound "fx_plunger", 0, 1, 0.1, 0.25:Plunger.Fire
    If keycode = RightFlipperKey Then Controller.Switch(32) = False
End Sub

'*********
' Switches
'*********

' Slings
Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    PlaySound SoundFX("fx_slingshot", DOFContactors), 0, 1, -0.05, 0.05
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

Sub RightSlingShot_Slingshot
    PlaySound SoundFX("fx_slingshot", DOFContactors), 0, 1, 0.05, 0.05
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

' Scoring rubbers

Sub sw17a_Hit:PlaySound "fx_Rubber", 0, 1, pan(ActiveBall):vpmTimer.PulseSw 17:End Sub
Sub sw17b_Hit:PlaySound "fx_Rubber", 0, 1, pan(ActiveBall):vpmTimer.PulseSw 17:End Sub
Sub sw17c_Hit:PlaySound "fx_Rubber", 0, 1, pan(ActiveBall):vpmTimer.PulseSw 17:End Sub
Sub sw17d_Hit:PlaySound "fx_Rubber", 0, 1, pan(ActiveBall):vpmTimer.PulseSw 17:End Sub
Sub sw17e_Hit:PlaySound "fx_Rubber", 0, 1, pan(ActiveBall):vpmTimer.PulseSw 17:End Sub
Sub sw17f_Hit:PlaySound "fx_Rubber", 0, 1, pan(ActiveBall):vpmTimer.PulseSw 17:End Sub
Sub sw17g_Hit:PlaySound "fx_Rubber", 0, 1, pan(ActiveBall):vpmTimer.PulseSw 17:End Sub

' Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 40:PlaySound SoundFX("fx_bumper", DOFContactors), 0, 1, 0, -0.05:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 38:PlaySound SoundFX("fx_bumper", DOFContactors), 0, 1, 0, 0.05:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 39:PlaySound SoundFX("fx_bumper", DOFContactors), 0, 1, 0, 0:End Sub

' Drain hole
Sub Drain_Hit:Playsound "fx_drain":bsTrough.AddBall Me:End Sub

' Saucer
Sub sw5_Hit():bsSaucer.AddBall 0:Playsound "fx_kicker_enter", 0, 1, 0, 0.1:End Sub

' Rollovers
Sub sw13_Hit:Controller.Switch(13) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw13_UnHit:Controller.Switch(13) = 0:End Sub

Sub sw14_Hit:Controller.Switch(14) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw14_UnHit:Controller.Switch(14) = 0:End Sub

Sub sw15_Hit:Controller.Switch(15) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw15_UnHit:Controller.Switch(15) = 0:End Sub

Sub sw18_Hit:Controller.Switch(18) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw18_UnHit:Controller.Switch(18) = 0:End Sub

Sub sw19_Hit:Controller.Switch(19) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw19_UnHit:Controller.Switch(19) = 0:End Sub

Sub sw30_Hit:Controller.Switch(30) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw30_UnHit:Controller.Switch(30) = 0:End Sub

Sub sw31_Hit:Controller.Switch(31) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw31_UnHit:Controller.Switch(31) = 0:End Sub

Sub sw34_Hit:Controller.Switch(34) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw34_UnHit:Controller.Switch(34) = 0:End Sub

Sub sw35_Hit:Controller.Switch(35) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw35_UnHit:Controller.Switch(35) = 0:End Sub

'Spinners
Sub sw33_Spin():vpmTimer.PulseSw 33:PlaySound "fx_spinner", 0, 1, -0.1:End Sub

' Droptargets
'Left
Sub sw24_Hit:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw23_Hit:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw22_Hit:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw21_Hit:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw20_Hit:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub

Sub sw24_Dropped:dtbankL.Hit 1:End Sub
Sub sw23_Dropped:dtbankL.Hit 2:End Sub
Sub sw22_Dropped:dtbankL.Hit 3:End Sub
Sub sw21_Dropped:dtbankL.Hit 4:End Sub
Sub sw20_Dropped:dtbankL.Hit 5:End Sub

'Middle
Sub sw4_Hit:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw3_Hit:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw2_Hit:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub

Sub sw4_Dropped:dtbankM.Hit 1:End Sub
Sub sw3_Dropped:dtbankM.Hit 2:End Sub
Sub sw2_Dropped:dtbankM.Hit 3:End Sub

'Top
Sub sw28_Hit:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw27_Hit:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw26_Hit:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub

Sub sw28_Dropped:dtbankT.Hit 1:End Sub
Sub sw27_Dropped:dtbankT.Hit 2:End Sub
Sub sw26_Dropped:dtbankT.Hit 3:End Sub

'*********
'Solenoids
'*********

SolCallback(6) = "vpmSolSound SoundFX(""fx_knocker"",DOFKnocker),"
SolCallback(25) = "vpmNudge.SolGameOn"
SolCallback(7) = "bsTrough.SolOut"
SolCallback(2) = "bsSaucer.solout"
SolCallback(9) = "dtbankT.SolDropUp"
SolCallback(10) = "dtbankM.SolDropUp"
SolCallback(8) = "dtbankL.SolDropUp"

Set MotorCallback = GetRef("UpdateFlipper")

Sub UpdateFlipper
    LeftFlipperP.RotZ = LeftFlipper1.CurrentAngle
End Sub

'**************
' Flipper Subs
'**************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySound SoundFX("fx_flipperup", DOFFlippers), 0, 1, -0.1, 0.1
        LeftFlipper.RotateToEnd
        LeftFlipper1.RotateToEnd
    Else
        PlaySound SoundFX("fx_flipperdown", DOFFlippers), 0, 1, -0.1, 0.1
        LeftFlipper.RotateToStart
        LeftFlipper1.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySound SoundFX("fx_flipperup", DOFFlippers), 0, 1, 0.1, 0.1
        RightFlipper.RotateToEnd
        RightFlipper1.RotateToEnd
    Else
        PlaySound SoundFX("fx_flipperdown", DOFFlippers), 0, 1, 0.1, 0.1
        RightFlipper.RotateToStart
        RightFlipper1.RotateToStart
    End If
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, -0.1, 0.25
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, 0.1, 0.25
End Sub

Sub LeftFlipper1_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, -0.1, 0.25
End Sub

Sub RightFlipper1_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, 0.1, 0.25
End Sub

'*****************
'   Gi Effects
'*****************

Dim OldGiState
OldGiState = -1 'start witht he Gi off

Sub GiON
    For each x in aGiLights
        GiEffect
    Next
End Sub

Sub GiOFF
    For each x in aGiLights
        x.State = 0
    Next
End Sub

Sub GiEffect
    For each x in aGiLights
        x.Duration 2, 1000, 1
    Next
End Sub

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
LampTimer.Interval = 10 'lamp fading speed
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
    NFadeL 1, l1
    NFadeL 2, l2
    NFadeL 3, l3
    NFadeL 4, l4
    NFadeL 5, l5
    NFadeL 6, l6
    NFadeL 7, l7
    NFadeL 8, l8
    NFadeLm 9, l9
    Flash 9, l9a
    NFadeL 10, l10
    NFadeT 11, l11, "Same Player Shoots Again"
    NFadeLm 12, l12
    Flash 12, l12a
    NFadeT 13, l13, "Ball In Play"
    NFadeL 14, l14
    NFadeL 15, l15
    'NFadeL 16, l16
    NFadeL 17, l17
    NFadeL 18, l18
    NFadeL 19, l19
    NFadeL 20, l20
    NFadeL 21, l21
    NFadeL 22, l22
    NFadeL 23, l23
    NFadeLm 24, l24
    Flash 24, l24a
    NFadeLm 25, Bumper1La
    NFadeL 25, Bumper1L
    NFadeL 26, l26
    NFadeT 27, l27, "Match"
    NFadeL 28, l28
    NFadeT 29, l29, "High Score to Date"
    NFadeL 30, l30
    NFadeL 31, l31
    'NFadeL 32, l32
    NFadeL 33, l33
    NFadeL 34, l34
    NFadeL 35, l35
    NFadeL 36, l36
    NFadeLm 37, l37
    Flash 37, l37a
    NFadeL 38, l38
    NFadeL 39, l39
    NFadeLm 40, l40
    Flash 40, l40a
    NFadeLm 41, Bumper3La
    NFadeL 41, Bumper3L
    NFadeL 42, l42
    NFadeL 43, l43
    NFadeLm 44, l44
    Flash 44, l44a
    NFadeTm 45, l45, "Game Over"
    NFadeL 46, l46
    NFadeL 47, l47
    'NFadeL 48, l48
    NFadeL 49, l49
    NFadeL 50, l50
    NFadeL 51, l51
    NFadeL 52, l52
    NFadeL 53, l53
    NFadeL 54, l54
    NFadeL 55, l55
    NFadeLm 56, l56
    Flash 56, l56a
    NFadeLm 57, Bumper2La
    NFadeL 57, Bumper2L
    NFadeL 58, l58
    NFadeL 59, l59
    NFadeL 60, l60
    NFadeT 61, l61, "TILT"
    NFadeL 62, l62
    NFadeL 63, l63
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

'************************************
'          LEDs Display
'     Based on Scapino's LEDs
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

'Assign 7-digit output to reels
Set Digits(0) = a0
Set Digits(1) = a1
Set Digits(2) = a2
Set Digits(3) = a3
Set Digits(4) = a4
Set Digits(5) = a5
Set Digits(6) = a6

Set Digits(7) = b0
Set Digits(8) = b1
Set Digits(9) = b2
Set Digits(10) = b3
Set Digits(11) = b4
Set Digits(12) = b5
Set Digits(13) = b6

Set Digits(14) = c0
Set Digits(15) = c1
Set Digits(16) = c2
Set Digits(17) = c3
Set Digits(18) = c4
Set Digits(19) = c5
Set Digits(20) = c6

Set Digits(21) = d0
Set Digits(22) = d1
Set Digits(23) = d2
Set Digits(24) = d3
Set Digits(25) = d4
Set Digits(26) = d5
Set Digits(27) = d6

Set Digits(28) = e0
Set Digits(29) = e1
Set Digits(30) = e2
Set Digits(31) = e3

Sub UpdateLeds
    On Error Resume Next
    Dim ChgLED, ii, jj, chg, stat
    ChgLED = Controller.ChangedLEDs(&HFF, &HFFFF)
    If Not IsEmpty(ChgLED)Then
        For ii = 0 To UBound(ChgLED)
            chg = chgLED(ii, 1):stat = chgLED(ii, 2)
            For jj = 0 to 10
                If stat = Patterns(jj)OR stat = Patterns2(jj)then Digits(chgLED(ii, 0)).SetValue jj
            Next
        Next
    End IF
End Sub

'******************************
' Diverse Collection Hit Sounds
'******************************

Sub aMetals_Hit(idx):PlaySound "fx_MetalHit2", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aRubber_Bands_Hit(idx):PlaySound "fx_rubber_band", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aRubber_Posts_Hit(idx):PlaySound "fx_rubber_post", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aRubber_Pins_Hit(idx):PlaySound "fx_rubber_pin", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aWoods_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub

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
        Pan = Csng(-((- tmp) ^10))
    End If
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2)))
End Function

Function AudioFade(ball) 'only on VPX 10.4 and newer
    Dim tmp
    tmp = ball.y * 2 / Table1.height-1
    If tmp > 0 Then
        AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10))
    End If
End Function

'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 5 ' total number of balls
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
    If UBound(BOT) = lob - 1 Then Exit Sub 'there no extra balls on this table

    ' play the rolling sound for each ball
    For b = lob to UBound(BOT)
        If BallVel(BOT(b)) > 1 Then
            If BOT(b).z < 30 Then
                ballpitch = Pitch(BOT(b))
            Else
                ballpitch = Pitch(BOT(b)) + 15000 'increase the pitch on a ramp or elevated surface
            End If
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b)), Pan(BOT(b)), 0, ballpitch, 1, 0
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

'*****************************************
'			FLIPPER SHADOWS
'*****************************************

sub FlipperTimer_Timer()
	FlipperLSh.RotZ = LeftFlipper.currentangle
	FlipperRSh.RotZ = RightFlipper.currentangle

End Sub

'*****************************************
'			BALL SHADOW
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
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) '+ 13
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) '- 13
        End If
        ballShadow(b).Y = BOT(b).Y + 10
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
End Sub

'Bally Skateball dips by Inkochnito
Sub editDips
    Dim vpmDips:Set vpmDips = New cvpmDips
    With vpmDips
        .AddForm 700, 400, "Skateball - DIP switches"
        .AddFrame 2, 0, 190, "Maximum credits", &H03000000, Array("10 credits", 0, "15 credits", &H01000000, "25 credits", &H02000000, "40 credits", &H03000000)                            'dip 25&26
        .AddFrame 2, 74, 190, "Hitting drop targets", &H00000040, Array("does not advance bonus", 0, "advences bonus 1 step", &H00000040)                                                   'dip 7
        .AddFrame 2, 120, 190, "50K && saucer extra ball lite", 32768, Array("50K then extra ball", 0, "50K && extra ball together", 32768)                                                 'dip 16
        .AddFrame 205, 0, 190, "Balls per game", &HC0000000, Array("2 balls", &HC0000000, "3 balls", 0, "4 balls", &H80000000, "5 balls", &H40000000)                                       'dip 31&32
        .AddFrame 205, 74, 190, "Saucer extra ball && outlane special", &H00800000, Array("first extra ball then outlane special", 0, "extra ball && outlane special together", &H00800000) 'dip 24
        .AddFrame 205, 120, 190, "Lighting A-B && S-K-A-T-E will", &H20000000, Array("not lite saucer extra ball", 0, "lite saucer extra ball", &H20000000)                                 'dip 30
        .AddChk 2, 170, 180, Array("Match feature", &H08000000)                                                                                                                             'dip 28
        .AddChk 2, 185, 115, Array("Credits displayed", &H04000000)                                                                                                                         'dip 27
        .AddChk 2, 200, 190, Array("Top drop targets memory", &H00000080)                                                                                                                   'dip 8
        .AddChk 205, 170, 190, Array("Center drop targets arrows memory", &H00002000)                                                                                                       'dip 14
        .AddChk 205, 185, 190, Array("Top drop targets arrows memory", &H00004000)                                                                                                          'dip 15
        .AddChk 205, 200, 190, Array("Bonus multiplier memory", &H00400000)                                                                                                                 'dip 23
        .AddLabel 50, 230, 320, 20, "Set selftest position 16,17,18 and 19 to 03 for the best gameplay."
        .AddLabel 50, 250, 320, 20, "Press 7 to enter selftest.  Press 7 repeatedly to get to 16,17,18,19"
        .AddLabel 50, 270, 300, 20, "Press Start button repeatedly to change to value 3. Press 7 to"
        .AddLabel 50, 290, 300, 20, "get to next setting.  Press 7 after setting 19 to exit selftest"
        .ViewDips
    End With
End Sub
Set vpmShowDips = GetRef("editDips")

BankTargetAnimation
Dim targetCntr, targetLoopCntr
Sub BankTargetAnimation  'Drop all 11 targets - one at a time during init
	sw20.timerinterval = 500: 	sw20.timerEnabled = 1: 	targetCntr = 0: 	targetLoopCntr = 1
End Sub

Sub Sw20_Timer()  'Add since there is time before vpinmame rom starts
	Select Case targetCntr
		Case 0: targetCntr = targetCntr+1:sw20.timerinterval = 400
		Case 1:	dtbankL.Hit(1): targetCntr = targetCntr+1:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, -.6
		Case 2:	dtbankL.Hit(2): targetCntr = targetCntr+1:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, -.6
		Case 3: dtbankL.Hit(3): targetCntr = targetCntr+1:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, -.6
		Case 4: dtbankL.Hit(4): targetCntr = targetCntr+1:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, -.6
		Case 5:	dtbankL.Hit(5): targetCntr = targetCntr+1:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, -.6
		Case 6: dtBankT.Hit(1): targetCntr = targetCntr+1:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, 0
		Case 7:	dtBankT.Hit(2): targetCntr = targetCntr+1:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, 0
		Case 8:	dtBankT.Hit(3): targetCntr = targetCntr+1:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, 0
		Case 9:	dtbankM.Hit(1): targetCntr = targetCntr+1:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, .6
		Case 10: dtbankM.Hit(2): targetCntr = targetCntr+1:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, .6
		Case 11: dtbankM.Hit(3): targetCntr = targetCntr+1 :PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, .6
		Case 12: targetCntr = targetCntr+1 
			If targetLoopCntr = 1 then 
				dtbankM.DropSol_On:dtbankT.DropSol_On:dtbankL.DropSol_On:sw20.timerinterval = 200: targetLoopCntr = 2: targetCntr = 1
			Else 
				sw20.timerEnabled = 0
			End If
	End Select
End Sub

'THINGS I CHANGED
'NEW MATERIALS
'ALL NEW TABLE SOUNDS IN SURROUND
'BALL SHADOW
'GOLD BALL
'WHITE ELVIS SUIT
'NEW LIGHTING

' Elvis - Stern 2004
' VPX 1.0.2 by JPSalas July 2017


' Thalamus - you should add a sw10 to the top of the plunger lane
' ref : https://vpinball.com/forums/topic/jps-elvis-mod/page/3/#post-201744

' Thalamus 2018-07-24
' Table has already its own "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' No special SSF tweaks yet.
' This is a JP table. He often uses walls as switches so I need to be careful of using PlaySoundAt

Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const BallSize = 50

' Language Roms
Const cGameName = "elvis" 'English

Dim VarHidden, UseVPMColoredDMD
If Table1.ShowDT = true then
    UseVPMColoredDMD = true
    VarHidden = 1
else
    UseVPMColoredDMD = False
    VarHidden = 0
    TextBox1.Visible = 0
    lrail.Visible = 0
    rrail.Visible = 0
end if

LoadVPM "01560000", "SEGA.VBS", 3.26

'********************
'Standard definitions
'********************

Const UseSolenoids = 2
Const UseLamps = 0
Const UseGI = 0
Const UseSync = 0
Const HandleMech = 0

' Standard Sounds
Const SSolenoidOn = "SolOn"
Const SSolenoidOff = "SolOff"
Const SCoin = "fx_Coin"

Dim bsTrough, mMag, mCenter, dtBankL, bsLock, bsUpper, bsJail, plungerIM, mElvis, x, i

'************
' Table init.
'************

Sub Table1_Init
    vpmInit Me
    With Controller
        .GameName = cGameName
        .Games(cGameName).Settings.Value("sound") = 1 'enable the rom sound
        .SplashInfoLine = "Elvis - Stern 2004" & vbNewLine & "VPX table by JPSalas v.1.0.2"
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .HandleKeyboard = 0
        .Hidden = VarHidden
        .Games(cGameName).Settings.Value("rol") = 0
        On Error Resume Next
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description

        On Error Goto 0
    End With

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

    vpmNudge.TiltSwitch = 56
    vpmNudge.Sensitivity = 1
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingshot)

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0, 14, 13, 12, 11, 0, 0, 0
        .InitKick BallRelease, 90, 4
        .InitExitSnd SoundFX("fx_ballrelease", DOFContactors), SoundFX("SolOn", DOFContactors)
        .Balls = 4
    End With

    ' Magnet
    Set mMag = New cvpmMagnet
    With mMag
        .InitMagnet Magnet, 60
        .Solenoid = 5
        .GrabCenter = 1
        .CreateEvents "mMag"
    End With

    ' Droptargets
    set dtBankL = new cvpmdroptarget
    With dtBankL
        .initdrop array(sw17, sw18, sw19, sw20, sw21), array(17, 18, 19, 20, 21)
        .initsnd SoundFX("fx_droptarget", DOFContactors), SoundFX("fx_DTReset", DOFContactors)
    End With

    ' Hotel Lock
    Set bsLock = New cvpmBallStack
    With bsLock
        .InitSw 0, 48, 0, 0, 0, 0, 0, 0
        .InitKick HLock, 180, 8
        .InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("SolOn", DOFContactors)
    End With

    ' Jail Lock
    Set bsJail = new cvpmBallStack
    With bsJail
        .InitSaucer sw34, 34, 186, 18
        .InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("SolOn", DOFContactors)
        .KickAngleVar = 1
        .KickForceVar = 1
    End With

    ' Upper Lock
    Set bsUpper = new cvpmBallStack
    With bsUpper
        .InitSaucer sw32, 32, 90, 10
        .InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("SolOn", DOFContactors)
        .KickAngleVar = 3
        .KickForceVar = 3
    End With

    ' Impulse Plunger
    Const IMPowerSetting = 60 ' Plunger Power
    Const IMTime = 0.6        ' Time in seconds for Full Plunge
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swplunger, IMPowerSetting, IMTime
        .Random 0.3
        .switch 16
        .InitExitSnd SoundFX("SlingshotLeft1", DOFContactors), SoundFX("SlingshotLeft1", DOFContactors)
        .CreateEvents "plungerIM"
    End With

    ' Elvis movement
    Set mElvis = New cvpmMech
    With mElvis
        .MType = vpmMechOneSol + vpmMechReverse + vpmMechLinear
        .Sol1 = 25
        .Length = 200
        .Steps = 200
        .AddSw 33, 0, 0
        .Callback = GetRef("UpdateElvis")
        .Start
    End With
    ' Initialize Elvis arms and legs
    ElvisArms 0
    ElvisLegs 0
End Sub

Sub Table1_Paused:Controller.Pause = 1:End Sub
Sub Table1_unPaused:Controller.Pause = 0:End Sub
Sub Table1_Exit:Controller.Stop:End Sub

'****
'Keys
'****

Sub Table1_KeyDown(ByVal keycode)
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If keycode = PlungerKey Then PlaySoundAt "PlungerPull", Plunger:Plunger.Pullback
    If KeyDownHandler(KeyCode) Then Exit Sub
    If keyUpperLeft Then Controller.Switch(55) = 1
    If keycode = KeyRules Then Rules
End Sub

Sub Table1_KeyUp(ByVal keycode)
    If KeyUpHandler(KeyCode) Then Exit Sub
    If keycode = PlungerKey Then PlaySoundAt "Plunger", Plunger:Plunger.Fire
    If keyUpperLeft Then Controller.Switch(55) = 0
End Sub

'**********************
'Elvis movement up/down
'**********************

Sub UpdateElvis(aNewPos, aSpeed, aLastPos)
    pStand.Transx = aNewPos: Playsound "MotorLong1", 0, 0.20, AudioPan(pStand), 0, 1, 1, 0, AudioFade(pStand)
    pLarm.Transx = aNewPos
    pLegs.Transx = aNewPos
    pRarm.Transx = aNewPos
    pBody.Transx = aNewPos
    pHead.Transx = aNewPos
End Sub

'********************
'     Flippers
'********************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("fx_flipperup", DOFFlippers), LeftFlipper
        LeftFlipper.RotateToEnd
        LeftFlipper1.RotateToEnd
    Else
        PlaySoundAt SoundFX("fx_flipperdown", DOFFlippers), LeftFlipper
        LeftFlipper.RotateToStart
        LeftFlipper1.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("fx_flipperup", DOFFlippers), RightFlipper
        RightFlipper.RotateToEnd
        RightFlipper1.RotateToEnd
    Else
        PlaySoundAt SoundFX("fx_flipperdown", DOFFlippers), RightFlipper
        RightFlipper.RotateToStart
        RightFlipper1.RotateToStart
    End If
End Sub



'*********
'Solenoids
'*********

SolCallBack(1) = "SolTrough"
SolCallBack(2) = "Auto_Plunger"
SolCallBack(3) = "dtBankL.SolDropUp"
SolCallBack(6) = "bsJail.SolOut"
SolCallBack(7) = "bsLock.SolOut"
SolCallBack(8) = "CGate.Open ="
SolCallBack(12) = "bsUpper.SolOut"
SolCallBack(19) = "SolHotelDoor"
SolCallBack(24) = "vpmsolsound SoundFX(""Knocker"",DOFKnocker),"

Sub SolTrough(Enabled)
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

Sub SolHotelDoor(Enabled)
    If Enabled Then
        fDoor.RotatetoEnd
        doorwall.IsDropped = 1
		PlaySoundAt "SolOn", door
    Else
        fDoor.RotatetoStart
        doorwall.IsDropped = 0
		PlaySoundAt "SolOff", door
    End If
End Sub

'*********
' Flashers
'*********
SolCallBack(20) = "SetLamp 100,"
SolCallback(21) = "SetLamp 101,"
SolCallBack(22) = "SetLamp 102,"
SolCallBack(23) = "SetLamp 103,"
SolCallBack(32) = "SetLamp 104,"
SolCallBack(31) = "SetLamp 105,"



'****************
' Elvis animation
'****************

SolCallBack(29) = "ElvisLegs"
SolCallBack(30) = "ElvisArms"

Sub ElvisLegs(Enabled)
    If Enabled Then
        fLegs.RotatetoStart
    Else
        fLegs.RotatetoEnd
    End If
End Sub

Sub ElvisArms(Enabled)
    If Enabled Then
        fArms.RotatetoStart
    Else
        fArms.RotatetoEnd
    End If
End Sub

' ************************************
' Switches, bumpers, lanes and targets
' ************************************

Sub Drain_Hit:PlaySoundAt "fx_drainLong", Drain:bsTrough.AddBall Me:End Sub
'Sub Drain_Hit:Me.destroyball:End Sub 'debug

Sub sw9_Hit:vpmTimer.PulseSw 9:End Sub

Sub sw10_Hit:Controller.Switch(10) = 1:End Sub
Sub sw10_unHit:Controller.Switch(10) = 0:End Sub
Sub sw17_Dropped:dtBankL.Hit 1:End Sub
Sub sw18_Dropped:dtBankL.Hit 2:End Sub
Sub sw19_Dropped:dtBankL.Hit 3:End Sub
Sub sw20_Dropped:dtBankL.Hit 4:End Sub
Sub sw21_Dropped:dtBankL.Hit 5:End Sub
Sub sw22_Hit:vpmTimer.PulseSw 22:End Sub
Sub sw23_Hit:vpmTimer.PulseSw 23:End Sub
Sub sw24_Hit:vpmTimer.PulseSw 24:End Sub
Sub sw25_Spin:PlaySound "fx_spinner", 0, .25, AudioPan(sw25), 0.25, 0, 0, 1, AudioFade(sw25):vpmTimer.PulseSw 25:End Sub
Sub sw26_Hit:Controller.Switch(26) = 1:End Sub
Sub sw26_unHit:Controller.Switch(26) = 0:End Sub
Sub sw27_Hit:Controller.Switch(27) = 1:End Sub
Sub sw27_unHit:Controller.Switch(27) = 0:End Sub
Sub sw28_Hit:Controller.Switch(28) = 1:End Sub
Sub sw28_unHit:Controller.Switch(28) = 0:End Sub
Sub sw30_Hit:Controller.Switch(30) = 1:End Sub
Sub sw30_unHit:Controller.Switch(30) = 0:End Sub
Sub sw31_Hit:Controller.Switch(31) = 1:End Sub
Sub sw31_unHit:Controller.Switch(31) = 0:End Sub
Sub sw32_Hit:PlaySoundAtVol "KickerEnter", sw32, 0.5:bsUpper.AddBall 0:End Sub
Sub sw32_unhit: PlaySoundAt "PopperBall", sw32: End Sub
Sub sw34_Hit:PlaySoundAtVol "KickerEnter", sw34, 0.5:bsJail.AddBall 0:End Sub
Sub sw34_unhit: PlaySoundAt "PopperBall", sw34: End Sub
Sub sw36_Hit:vpmTimer.PulseSw 36:End Sub
Sub sw37_Hit:vpmTimer.PulseSw 37:End Sub
Sub sw38_Hit:vpmTimer.PulseSw 38:End Sub
Sub sw39_Hit:vpmTimer.PulseSw 39:End Sub
Sub sw40_Hit:vpmTimer.PulseSw 40:End Sub
Sub sw41_Hit:Controller.Switch(41) = 1:End Sub
Sub sw41_UnHit:Controller.Switch(41) = 0:End Sub
Sub sw42_Hit:Controller.Switch(42) = 1:End Sub
Sub sw42_UnHit:Controller.Switch(42) = 0:End Sub
Sub sw43_Hit:Controller.Switch(43) = 1:End Sub
Sub sw43_UnHit:Controller.Switch(43) = 0:End Sub
Sub sw45_Hit:Controller.Switch(45) = 1:End Sub
Sub sw45_unHit:Controller.Switch(45) = 0:End Sub
Sub sw46_Hit:Controller.Switch(46) = 1:End Sub
Sub sw46_unHit:Controller.Switch(46) = 0:End Sub
Sub sw47_Hit:Controller.Switch(47) = 1:End Sub
Sub sw47_unHit:Controller.Switch(47) = 0:End Sub
Sub HLock_Hit:PlaySoundAtVol "KickerEnter", HLock, 0.5:bsLock.AddBall Me:End Sub
Sub HLock_unhit: PlaySoundAt "PopperBall", HLock: End Sub
Sub sw48_Hit:Controller.Switch(48) = 1:End Sub
Sub sw48_unHit:Controller.Switch(48) = 0:End Sub

'Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 49:PlaySoundAt SoundFX("fx_bumper1", DOFContactors), Bumper1:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 50:PlaySoundAt SoundFX("fx_bumper2", DOFContactors), Bumper2:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 51:PlaySoundAt SoundFX("fx_bumper3", DOFContactors), Bumper3:End Sub

Sub sw52_Hit:vpmTimer.PulseSw 52:End Sub
Sub sw53_Hit:vpmTimer.PulseSw 53:End Sub
Sub sw57_Hit::Controller.Switch(57) = 1:End Sub
Sub sw57_UnHit:Controller.Switch(57) = 0:End Sub
Sub sw58_Hit:Controller.Switch(58) = 1:End Sub
Sub sw58_UnHit:Controller.Switch(58) = 0:End Sub
Sub sw60_Hit:Controller.Switch(60) = 1:End Sub
Sub sw60_UnHit:Controller.Switch(60) = 0:End Sub
Sub sw61_Hit:Controller.Switch(61) = 1:End Sub
Sub sw61_UnHit:Controller.Switch(61) = 0:End Sub

Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    PlaySoundAt SoundFX("SlingshotLeft", DOFContactors), Lemk
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 59
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
    PlaySoundAt SoundFX("SlingshotRight", DOFContactors), Remk
    RightSling4.Visible = 1
    Remk.RotX = 26
    RStep = 0
    vpmTimer.PulseSw 62
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

Sub sw63_Hit:Controller.Switch(63) = 1:End Sub
Sub sw63_UnHit:Controller.Switch(63) = 0:PlaySoundAt "BallBounce", door:End Sub

Sub sw64_Hit
    vpmTimer.PulseSw 64
    str = INT(ABS(ActiveBall.Vely) ^3)
    debug.print str
    DogDir = 5 'upwards
    HoundDogTimer.Enabled = 1
    PlaySoundAt SoundFX("MetalHit_Medium", DOFTargets), sw64
	PlaySoundAt "SolOn", sw64
End Sub

' Animate Hound Dog
Dim str 'strength of the hit
Dim DogStep, DogDir
DogStep = 0
DogDir = 0

Sub HoundDogTimer_Timer()
    DogStep = DogStep + DogDir
    HoundDog.TransZ = DogStep
    If DogStep> 100 Then DogDir = -5
    If DogStep> str Then DogDir = -5
    If DogStep <5 Then HoundDogTimer.Enabled = 0
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
    NFadeLm 57, l57
    Flash 57, l57a
    NFadeLm 58, l58
    Flash 58, l58a
    NFadeLm 59, l59
    Flash 59, l59a
    Flash 60, l60
    Flash 61, l61
    Flash 62, l62
    NFadeL 63, l63
    NFadeL 64, l64
    Flash 65, l65
    Flash 66, l66
    Flash 67, l67
    Flash 68, l68
    Flash 69, l69
    Flashm 70, l70a
    Flash 70, l70
    Flashm 71, l71a
    Flash 71, l71
    Flashm 72, l72a
    Flash 72, l72
    Flash 73, l73
    Flash 74, l74
    Flash 75, l75
    Flash 76, l76
    Flash 77, l77
    Flash 78, l78
    'flashers
    Flashm 100, f20a
    Flashm 100, f20b
    Flashm 100, f20c
    Flash 100, f20
    Flash 101, f21
    Flash 102, f22
    NFadeLm 103, f23a
    NFadeL 103, f23
    Flashm 104, f32a
    Flashm 104, f32b
    Flashm 104, f32c
    Flash 104, f32
    Flashm 105, f31a
    Flash 105, f31

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
            If FlashLevel(nr) <FlashMin(nr) Then
                FlashLevel(nr) = FlashMin(nr)
                FadingLevel(nr) = 0 'completely off
            End if
            Object.IntensityScale = FlashLevel(nr)
        Case 5 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
            If FlashLevel(nr)> FlashMax(nr) Then
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
            FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown(nr)
            If FlashLevel(nr) <FlashMin(nr) Then
                FlashLevel(nr) = FlashMin(nr)
                FadingLevel(nr) = 0 'completely off
            End if
            Object.IntensityScale = FlashLevel(nr)
            If FadingLevel(nr) = 0 AND FlashRepeat(nr) Then 'repeat the flash
                FlashRepeat(nr) = FlashRepeat(nr) -1
                If FlashRepeat(nr) Then FadingLevel(nr) = 5
            End If
        Case 5 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
            If FlashLevel(nr)> FlashMax(nr) Then
                FlashLevel(nr) = FlashMax(nr)
                FadingLevel(nr) = 1 'completely on
            End if
            Object.IntensityScale = FlashLevel(nr)
            If FadingLevel(nr) = 1 AND FlashRepeat(nr) Then FadingLevel(nr) = 4
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

Sub NFadeTm(nr, object, b)
    Select Case FadingLevel(nr)
        Case 4:object.Text = ""
        Case 5:object.Text = message
    End Select
End Sub




'******************
' RealTime Updates
'******************

Set MotorCallback = GetRef("RealTimeUpdates")

Sub RealTimeUpdates
    RollingUpdate
    door.RotX = fdoor.CurrentAngle
    pRarm.objRoty = fArms.CurrentAngle
    pLarm.objRoty = fArms.CurrentAngle
    pLegs.objRotY = fLegs.CurrentAngle
End Sub

'*************************
' GI - needs new vpinmame
'*************************

Set GICallback = GetRef("GIUpdate")

Sub GIUpdate(no, Enabled)
    For each x in aGiLights
        x.State = ABS(Enabled)
    Next

    For each x in aGiFlashers
        x.Visible = ABS(Enabled)
    Next
End Sub

Sub Break_Hit
	ActiveBall.VelY = 0
	ActiveBall.VelX = 0
End Sub

'******
' Rules
'******
Dim Msg(20)
Sub Rules()
    Msg(0) = "Elvis - Stern 2004" &Chr(10) &Chr(10)
    Msg(1) = ""
    Msg(2) = "OBJECTIVE:Get to Graceland by lighting the following:"
    Msg(3) = "*FEATURED HITS COMPLETED (start all 5 song modes)"
    Msg(4) = "   Hound Dog (Shoot HOUND DOG Target)"
    Msg(5) = "   Blue Suede Shoes (Shoot CENTER LOOP with Upper Right Flipper)"
    Msg(6) = "   Heartbreak Hotel (Shoot balls into HEARTBREAK HOTEL on Upper Playfield)"
    Msg(7) = "   Jailhouse Rock (Shoot balls into the JAILHOUSE EJECT HOLE)"
    Msg(8) = "   All Shook Up (Shoot ALL-SHOOK shots)"
    Msg(9) = "*GIFTS FROM ELVIS COMPLETED ~ Shoot E-L-V-I-S Drop Targets to light GIFT"
    Msg(10) = " FROM ELVIS on the TOP EJECT HOLE"
    Msg(11) = "*TOP TEN COUNTDOWN COMPLETED ~ Shoot lit'music lights's to advance TOP"
    Msg(12) = " 10 COUNTDOWN."
    Msg(13) = "SKILL SHOT: Plunge ball in the WLVS Top Lanes or E-L-V-I-S Drop Targets"
    Msg(14) = "MYSTERY:Ball in the Pop Bumpers will change channels until all 3 TVs match."
    Msg(15) = "EXTRA BALL: Shoot Right Ramp ro light Extra Ball."
    Msg(16) = "TCB: Complete T-C-B to double all scoring."
    Msg(17) = "ENCORE: Spell E-N-C-O-R-E (letters lit in Back Panel) to earn"
    Msg(18) = "an Extra Multiball after the game."
    Msg(19) = ""
    Msg(20) = ""
    For X = 1 To 20
        Msg(0) = Msg(0) + Msg(X) &Chr(13)
    Next

    MsgBox Msg(0), , "         Instructions and Rule Card"
End Sub



'*****************************************
'	ninuzzu's	BALL SHADOW
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
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) + 6
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) - 6
        End If
        ballShadow(b).Y = BOT(b).Y + 12
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
End Sub



'*********************************************************************
'                 Positional Sound Playback Functions
'*********************************************************************

' Play a sound, depending on the X,Y position of the table element (especially cool for surround speaker setups, otherwise stereo panning only)
' parameters (defaults): loopcount (1), volume (1), randompitch (0), pitch (0), useexisting (0), restart (1))
' Note that this will not work (currently) for walls/slingshots as these do not feature a simple, single X,Y position

Sub PlayXYSound(soundname, tableobj, loopcount, volume, randompitch, pitch, useexisting, restart)
	PlaySound soundname, loopcount, volume, AudioPan(tableobj), randompitch, pitch, useexisting, restart, AudioFade(tableobj)
End Sub

' Similar subroutines that are less complicated to use (e.g. simply use standard parameters for the PlaySound call)
Sub PlaySoundAt(soundname, tableobj)
    PlaySound soundname, 1, 1, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname)
    PlaySoundAt soundname, ActiveBall
End Sub

Sub PlaySoundAtVol(sound, tableobj, Vol)
		PlaySound sound, 1, Vol, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
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
        If BallVel(BOT(b) ) > 1 and not BOT(b).z Then
			rolling(b) = True
			if BOT(b).z < 30 Then ' Ball on playfield
						PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), audioPan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
			Else ' Ball on raised ramp
						PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*.5, audioPan(BOT(b) ), 0, Pitch(BOT(b) )+30000, 1, 0, AudioFade(BOT(b) )
				End If
		Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If
    Next

	on error resume next ' In case VP is too old..
		' Kill ball spin
	if mLockMagnet.MagnetOn then
		dim rampballs:rampballs = mLockMagnet.Balls
		dim obj
		for each obj in rampballs
			obj.AngMomZ= 0
			obj.AngVelZ= 0
			obj.AngMomY= 0
			obj.AngVelY= 0
			obj.AngMomX= 0
			obj.AngVelX= 0
			obj.velx = 0
			obj.vely = 0
			obj.velz = 0
		next 
	end if
	on error goto 0
End Sub

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
	PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
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
	PlaySound "fx_PinHit", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
	PlaySound "target", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Switches_Hit (idx)
	PlaySound "fx_Sensor", 0, 1, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub DropTargets_Hit (idx)
	PlaySound "fx_DTDrop", 0, 1, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
	PlaySound "metalhit_thin", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
	PlaySound "metalhit_medium", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
	PlaySound "metalhit2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub aGates_Hit (idx)
	PlaySound "GateWire", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinner_Spin
	PlaySound "fx_spinner", 0, .25, AudioPan(Spinner), 0.25, 0, 0, 1, AudioFade(Spinner)
End Sub

Sub Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then 
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub Posts_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 16 then 
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 16 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub

Sub LeftFlipper_Collide(parm)
 	RandomSoundFlipper()
End Sub

Sub RightFlipper_Collide(parm)
 	RandomSoundFlipper()
End Sub

Sub LeftFlipper1_Collide(parm)
 	RandomSoundFlipper()
End Sub

Sub RightFlipper1_Collide(parm)
 	RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub



' Ramp Soundss & Help triggers

Sub sw31_Hit:PlaySound "Metalrolling", 0, Vol(ActiveBall), AudioPan(Primitive65), 0, Pitch(ActiveBall), 1, 0, AudioFade(Primitive65):End Sub

Sub sw28_Hit:PlaySound "Metalrolling", 0, Vol(ActiveBall), AudioPan(Primitive63), 0, Pitch(ActiveBall), 1, 0, AudioFade(Primitive63):End Sub

Sub REnd1_Hit()
    StopSound "Metalrolling"
    PlaySoundat "BallBounce", REnd1
End Sub

Sub REnd2_Hit()
    StopSound "Metalrolling"
    PlaySoundat "BallBounce", REnd2
End Sub

' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


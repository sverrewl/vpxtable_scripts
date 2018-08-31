' Lightning / IPD No. 1441 / March, 1981 / 4 Players
' http://www.ipdb.org/machine.cgi?id=1441
' VPX table by JPSalas 2017
' Script based on destruks script

Option Explicit
Randomize

' Thalamus 2018-07-23
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' No special SSF tweaks yet.
' This is a JP table. He often uses walls as switches so I need to be careful of using PlaySoundAt
' Wob 2018-08-09
' Added vpmInit Me to table init

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const BallSize = 50

LoadVPM "01530000", "stern.vbs", 3.10

Dim VarHidden
If Table1.ShowDT = true then
    VarHidden = 1
else
    VarHidden = 0
    For each x in aReels
        x.Visible = 0
    Next

    lrail.Visible = 0
    rrail.Visible = 0
end if

Dim bsTrough, bsSaucer1, bsSaucer2, dtL, dtT, dtC
Dim x, i, j, k

Const cGameName = "lightnin"
Const UseSolenoids = 2
Const UseLamps = 0
Const UseGI = 0
Const UseSync = 0
Const HandleMech = 0

' Standard Sounds
Const SSolenoidOn = "fx_Solenoid"
Const SSolenoidOff = ""
Const SCoin = "fx_coin"

'************
' Table init.
'************

Sub Table1_Init
	vpmInit Me
    With Controller
        .GameName = cGameName
        .SplashInfoLine = "Stern's Lightning, Stern 1981"
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = VarHidden
        If Err Then MsgBox Err.Description
    End With
    On Error Goto 0
    Controller.SolMask(0) = 0
    vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the timer to renable all the solenoids after 2 seconds
    Controller.Run

    ' Trough
    Set bsTrough = New cvpmBallstack
    bsTrough.InitSw 0, 33, 34, 35, 0, 0, 0, 0
    bsTrough.InitKick BallRelease, 90, 5
    bsTrough.InitExitSnd "fx_ballrel", "fx_Solenoid"
    bsTrough.Balls = 3

    ' Left Drop targets
    Set dtL = New cvpmDropTarget
    dtL.InitDrop Array(sw38, sw39, sw40), Array(38, 39, 40)
    dtL.InitSnd SoundFX("fx_droptarget", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)

    ' Center Drop targets
    Set dtC = New cvpmDropTarget
    dtC.InitDrop Array(sw19, sw20, sw21), Array(19, 20, 21)
    dtC.InitSnd SoundFX("fx_droptarget", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)

    ' Top Drop targets
    Set dtT = New cvpmDropTarget
    dtT.InitDrop Array(sw22, sw23, sw24), Array(22, 23, 24)
    dtT.InitSnd SoundFX("fx_droptarget", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)

    ' Saucers
    Set bsSaucer2 = New cvpmBallStack
    bsSaucer2.InitSaucer Kicker2, 36, 170, 14
    bsSaucer2.InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    bsSaucer2.KickAngleVar = 30

    Set bsSaucer1 = New cvpmBallStack
    bsSaucer1.InitSaucer Kicker1, 37, 330, 14
    bsSaucer1.InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    bsSaucer1.KickAngleVar = 30

    ' Nudging
    vpmNudge.TiltSwitch = 7
    vpmNudge.Sensitivity = 5
    vpmNudge.TiltObj = Array(LeftSlingshot, RightSlingshot, LeftSlingshot1, RightSlingshot1)

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1
End Sub

Sub Table1_Paused:Controller.Pause = 1:End Sub
Sub Table1_unPaused:Controller.Pause = 0:End Sub
Sub Table1_Exit:Controller.Stop:End Sub

'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If keycode = PlungerKey Then PlaySound "fx_PlungerPull", 0, 1, 0.1, 0.25:Plunger.Pullback
    If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If vpmKeyUp(keycode) Then Exit Sub
    If keycode = PlungerKey Then PlaySound "fx_plunger", 0, 1, 0.1, 0.25:Plunger.Fire
End Sub

'*********
' Switches
'*********

' Slings & Rubbers

Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    PlaySound SoundFX("fx_slingshot", DOFContactors), 0, 1, -0.05, 0.05
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 13
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
    vpmTimer.PulseSw 15
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

Dim LStep1, RStep1

Sub LeftSlingShot1_Slingshot
    PlaySound SoundFX("fx_slingshot", DOFContactors), 0, 1, -0.05, 0.05
    LeftSling8.Visible = 1
    Lemk1.RotX = 26
    LStep1 = 0
    vpmTimer.PulseSw 14
    LeftSlingShot1.TimerEnabled = 1
End Sub

Sub LeftSlingShot1_Timer
    Select Case LStep1
        Case 1:LeftSLing8.Visible = 0:LeftSLing7.Visible = 1:Lemk1.RotX = 14
        Case 2:LeftSLing7.Visible = 0:LeftSLing6.Visible = 1:Lemk1.RotX = 2
        Case 3:LeftSLing6.Visible = 0:Lemk1.RotX = -10:LeftSlingShot1.TimerEnabled = 0
    End Select

    LStep1 = LStep1 + 1
End Sub

Sub RightSlingShot1_Slingshot
    PlaySound SoundFX("fx_slingshot", DOFContactors), 0, 1, 0.05, 0.05
    RightSling8.Visible = 1
    Remk1.RotX = 26
    RStep1 = 0
    vpmTimer.PulseSw 16
    RightSlingShot1.TimerEnabled = 1
End Sub

Sub RightSlingShot1_Timer
    Select Case RStep1
        Case 1:RightSLing8.Visible = 0:RightSLing7.Visible = 1:Remk1.RotX = 14
        Case 2:RightSLing7.Visible = 0:RightSLing6.Visible = 1:Remk1.RotX = 2
        Case 3:RightSLing6.Visible = 0:Remk1.RotX = -10:RightSlingShot1.TimerEnabled = 0
    End Select

    RStep1 = RStep1 + 1
End Sub

' Drain & holes
Sub Drain_Hit:bsTrough.AddBall Me:PlaySound "fx_drain" End Sub
Sub Kicker1_Hit:bsSaucer1.AddBall 0:PlaySound "fx_kicker-enter":End Sub
Sub Kicker2_Hit:bsSaucer2.AddBall 0:PlaySound "fx_kicker-enter":End Sub

' Rollovers
Sub SW9_Hit:Controller.Switch(9) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub SW9_UnHit:Controller.Switch(9) = 0:End Sub
Sub SW10_Hit:Controller.Switch(10) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub SW10_UnHit:Controller.Switch(10) = 0:End Sub
Sub SW11_Hit:Controller.Switch(11) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub SW11_UnHit:Controller.Switch(11) = 0:End Sub
Sub SW12_Hit:Controller.Switch(12) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub SW12_UnHit:Controller.Switch(12) = 0:End Sub
Sub SW12a_Hit:Controller.Switch(12) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub SW12a_UnHit:Controller.Switch(12) = 0:End Sub
Sub SW17_Hit:Controller.Switch(17) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub SW17_UnHit:Controller.Switch(17) = 0:End Sub
Sub SW18_Hit:Controller.Switch(18) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub SW18_UnHit:Controller.Switch(18) = 0:End Sub

' Targets
Sub sw8_Hit:vpmTimer.PulseSw 8:PlaySound SoundFX("fx_target", DOFTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw25_Hit:vpmTimer.PulseSw 25:PlaySound SoundFX("fx_target", DOFTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw26_Hit:vpmTimer.PulseSw 26:PlaySound SoundFX("fx_target", DOFTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw27_Hit:vpmTimer.PulseSw 27:PlaySound SoundFX("fx_target", DOFTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw28_Hit:vpmTimer.PulseSw 28:PlaySound SoundFX("fx_target", DOFTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw29_Hit:vpmTimer.PulseSw 29:PlaySound SoundFX("fx_target", DOFTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw30_Hit:vpmTimer.PulseSw 30:PlaySound SoundFX("fx_target", DOFTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw31_Hit:vpmTimer.PulseSw 31:PlaySound SoundFX("fx_target", DOFTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw32_Hit:vpmTimer.PulseSw 32:PlaySound SoundFX("fx_target", DOFTargets), 0, 1, pan(ActiveBall):End Sub

' VPX Droptargets may use the HIT event to play the sound and the Dropped event to register the hit to vpinmama
Sub sw38_Hit:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw39_Hit:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw40_Hit:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw22_Hit:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw23_Hit:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw24_Hit:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw19_Hit:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw20_Hit:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw21_Hit:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub

Sub sw38_Dropped:dtL.Hit 1:End Sub
Sub sw39_Dropped:dtL.Hit 2:End Sub
Sub sw40_Dropped:dtL.Hit 3:End Sub
Sub sw22_Dropped:dtT.Hit 1:End Sub
Sub sw23_Dropped:dtT.Hit 2:End Sub
Sub sw24_Dropped:dtT.Hit 3:End Sub
Sub sw19_Dropped:dtC.Hit 1:End Sub
Sub sw20_Dropped:dtC.Hit 2:End Sub
Sub sw21_Dropped:dtC.Hit 3:End Sub

' Spinner
Sub Spinner1_Spin:vpmTimer.PulseSw 5:PlaySound "fx_spinner", 0, 1, 0.01:End Sub

'*********
'Solenoids
'*********

'SolCallback(1)="vpmSolSound ""sling"","
'SolCallback(2)="vpmSolSound ""sling"","
'SolCallback(3)="vpmSolSound ""sling"","
'SolCallback(4)="vpmSolSound ""sling"","
SolCallback(5) = "dtT.SolDropUp"
SolCallback(6) = "vpmsolsound SoundFX(""fx_knocker"",DOFKnocker),"

SolCallback(8) = "bsTrough.SolOut"
SolCallback(9) = "dtC.SolDropUp"
SolCallback(10) = "dtL.SolDropUp"
SolCallback(11) = "bsSaucer1.SolOut"
SolCallback(12) = "bsSaucer2.SolOut"

SolCallback(19) = "vpmNudge.SolGameOn"

Set LampCallback = GetRef("UpdateMultipleLamps")

'********************
'     Flippers
'********************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySound SoundFX("fx_flipperup", DOFFlippers), 0, 1, -0.1, 0.25
        LeftFlipper.RotateToEnd
        LeftFlipper1.RotateToEnd
    Else
        PlaySound SoundFX("fx_flipperdown", DOFFlippers), 0, 1, -0.1, 0.25
        LeftFlipper.RotateToStart
        LeftFlipper1.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySound SoundFX("fx_flipperup", DOFFlippers), 0, 1, 0.1, 0.25
        RightFlipper.RotateToEnd
        RightFlipper1.RotateToEnd
    Else
        PlaySound SoundFX("fx_flipperdown", DOFFlippers), 0, 1, 0.1, 0.25
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

'************************************
'          LEDs Display
'************************************

Dim Digits(34)
Digits(0) = Array(D331, D332, D333, D334, D335, D336, D337, D338)
Digits(1) = Array(D1, D2, D3, D4, D5, D6, D7)
Digits(2) = Array(D11, D12, D13, D14, D15, D16, D17)
Digits(3) = Array(D21, D22, D23, D24, D25, D26, D27, D28)
Digits(4) = Array(D31, D32, D33, D34, D35, D36, D37)
Digits(5) = Array(D41, D42, D43, D44, D45, D46, D47)
Digits(6) = Array(D51, D52, D53, D54, D55, D56, D57)
Digits(7) = Array(D341, D342, D343, D344, D345, D346, D347, D348)
Digits(8) = Array(D61, D62, D63, D64, D65, D66, D67)
Digits(9) = Array(D71, D72, D73, D74, D75, D76, D77)
Digits(10) = Array(D81, D82, D83, D84, D85, D86, D87, D88)
Digits(11) = Array(D91, D92, D93, D94, D95, D96, D97)
Digits(12) = Array(D101, D102, D103, D104, D105, D106, D107)
Digits(13) = Array(D111, D112, D113, D114, D115, D116, D117)
Digits(14) = Array(D351, D352, D353, D354, D355, D356, D357, D358)
Digits(15) = Array(D121, D122, D123, D124, D125, D126, D127)
Digits(16) = Array(D131, D132, D133, D134, D135, D136, D137)
Digits(17) = Array(D141, D142, D143, D144, D145, D146, D147, D148)
Digits(18) = Array(D151, D152, D153, D154, D155, D156, D157)
Digits(19) = Array(D161, D162, D163, D164, D165, D166, D167)
Digits(20) = Array(D171, D172, D173, D174, D175, D176, D177)
Digits(21) = Array(D361, D362, D363, D364, D365, D366, D367, D368)
Digits(22) = Array(D181, D182, D183, D184, D185, D186, D187)
Digits(23) = Array(D191, D192, D193, D194, D195, D196, D197)
Digits(24) = Array(D201, D202, D203, D204, D205, D206, D207, D208)
Digits(25) = Array(D211, D212, D213, D214, D215, D216, D217)
Digits(26) = Array(D221, D222, D223, D224, D225, D226, D227)
Digits(27) = Array(D231, D232, D233, D234, D235, D236, D237)
Digits(28) = Array(D241, D242, D243, D244, D245, D246, D247)
Digits(29) = Array(D251, D252, D253, D254, D255, D256, D257)
Digits(30) = Array(D261, D262, D263, D264, D265, D266, D267)
Digits(31) = Array(D271, D272, D273, D274, D275, D276, D277)
Digits(32) = Array(D281, D282, D283, D284, D285, D286, D287)
Digits(33) = Array(D291, D292, D293, D294, D295, D296, D297)

Sub UpdateMultipleLamps
    D281a.State = D281.State
    D282a.State = D282.State
    D286a.State = D286.State
    D291a.State = D291.State
    D292a.State = D292.State
    D296a.State = D296.State
End Sub

Sub UpdateLeds
    Dim ChgLED, ii, jj, num, chg, stat, obj, b, x
    ChgLED = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED) Then
        For ii = 0 To UBound(chgLED)
            num = chgLED(ii, 0):chg = chgLED(ii, 1):stat = chgLED(ii, 2)
            For Each obj In Digits(num)
                If chg And 1 Then obj.State = stat And 1
                chg = chg \ 2:stat = stat \ 2
            Next
        Next
    End If
End Sub

'***********
' GI lights
'***********

Dim OldGiState
OldGiState = -1 'start witht he Gi off

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

Sub GiEffect
    For each x in aGiLights
        x.Duration 2, 3000, 1
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
    If Not IsEmpty(chgLamp) Then
        For ii = 0 To UBound(chgLamp)
            LampState(chgLamp(ii, 0) ) = chgLamp(ii, 1)       'keep the real state in an array
            FadingLevel(chgLamp(ii, 0) ) = chgLamp(ii, 1) + 4 'actual fading step
        Next
    End If

    UpdateLeds
    UpdateLamps
    GIUpdate
    RollingUpdate
End Sub

Sub UpdateLamps

    NFadeL 4, l4

    NFadeL 6, l6
    NFadeL 7, l7
    NFadeL 8, l8
    NFadeL 9, l9
    NFadeL 10, l10
    NFadeL 11, l11
    NFadeL 12, l12
    NFadeL 14, l14

    NFadeL 20, l20

    NFadeL 22, l22
    NFadeL 23, l23
    NFadeL 24, l24
    NFadeL 25, l25
    NFadeL 26, l26
    NFadeLm 27, l27a
    NFadeL 27, l27
    NFadeL 28, l28
    NFadeLm 29, l29a
    NFadeL 29, l29
    NFadeL 30, l30

    NFadeL 36, l36

    NFadeL 38, l38
    NFadeL 39, l39
    NFadeL 40, l40
    NFadeL 41, l41
    NFadeL 42, l42
    NFadeLm 43, l43a
    NFadeL 43, l43
    NFadeL 44, l44
    NFadeL 46, l46
    NFadeL 47, l47

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
    NFadeL 62, l62

    ' backdrop lights
    If VarHidden Then
        NFadeT 13, l13, "Highscore"
        'NFadeT 29, l29, "Ball in Play"
        'NFadeT 45, l45, "Game Over"
        'NFadeT 61, l61, "Tilt"
        NFadeT 63, l63, "Match"
    End If
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

'******************************
' Diverse Collection Hit Sounds
'******************************

Sub aMetals_Hit(idx):PlaySound "fx_MetalHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Bands_Hit(idx):PlaySound "fx_rubber_band", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Posts_Hit(idx):PlaySound "fx_rubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Pins_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aWoods_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub REnd1_Hit:PlaySound "fx_balldrop":End Sub

'**********************
' Dipswitch menu script
'**********************

Sub editDips
    Dim vpmDips:Set vpmDips = New cvpmDips
    With vpmDips
        .AddForm 700, 400, "Lightning - DIP switches"
        .AddFrame 125, 122, 110, "Bonus timer award", &H00000010, Array("ON/Bonus Timer", 0, "OFF/HS feature", &H00000010)                                          'dip 5 reversed
        .AddFrame 125, 76, 110, "High score feature", &H00000020, Array("extra ball", 0, "free game", &H00000020)                                                   'dip 6 OK
        .AddFrame 2, 76, 110, "Balls per game", &H00000040, Array("3 balls", 0, "5 balls", &H00000040)                                                              'dip 7 OK
        .AddChk 248, 170, 110, Array("Background sound", &H00000080)                                                                                                'dip 8 OK
        .AddFrame 2, 122, 110, "Add-a-ball maximum", &H00001000, Array("3 balls", 0, "5 balls", &H00001000)                                                         'dip 13
        .AddChk 248, 185, 110, Array("Add-a-ball memory", &H00002000)                                                                                               'dip 14
        .AddFrame 125, 0, 110, "High game to date", 49152, Array("points", 0, "1 free game", &H00004000, "2 free games", 32768, "3 free games", 49152)              'dip 15&16 Wrong - Points=2freegames, 1game=None
        .AddFrame 371, 76, 110, "Green special limit", &H00010000, Array("1 per game", 0, "1 per ball", &H00010000)                                                 'dip 17 OK
        .AddFrame 2, 0, 110, "Maximum credits", &H00060000, Array("10 credits", 0, "15 credits", &H00020000, "25 credits", &H00040000, "40 credits", &H00060000)    'dip 18&19 wrong
        .AddChk 371, 185, 110, Array("Credits displayed", &H00080000)                                                                                               'dip 20 OK
        .AddChk 371, 170, 110, Array("Match feature", &H00100000)                                                                                                   'dip 21 OK
        .AddFrame 248, 122, 110, "Extra ball award", &H00200000, Array("100K points", 0, "extra ball", &H00200000)                                                  'dip 22
        .AddFrame 371, 122, 110, "Extra ball limit", &H00400000, Array("1 per game", 0, "1 per ball", &H00400000)                                                   'dip 23 OK
        .AddFrame 248, 76, 110, "Red special limit", &H00800000, Array("1 per game", 0, "1 per ball", &H00800000)                                                   'dip 24 OK
        .AddFrame 371, 0, 110, "Green special award", &H30000000, Array("no award", 0, "100K points", &H10000000, "free ball", &H20000000, "free game", &H30000000) 'dip 29&30
        .AddFrame 248, 0, 110, "Red special award", &HC0000000, Array("no award", 0, "100K points", &H40000000, "free ball", &H80000000, "free game", &HC0000000)   'dip 31&32
        .AddLabel 50, 240, 300, 20, "After hitting OK, press F3 to reset game with new settings."
        .ViewDips
    End With
End Sub
Set vpmShowDips = GetRef("editDips")

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

Const tnob = 5 ' total number of balls in this table is 4, but always use a higher number here because of the timing
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
  If Table1.VersionMinor > 3 OR Table1.VersionMajor > 10 Then
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
  Else
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0
  End if
End Sub


' Lightning / IPD No. 1441 / March, 1981 / 4 Players
' http://www.ipdb.org/machine.cgi?id=1441
' VPX8 table by JPSalas 2024, version 6.0.0
' Script based on destruks script

Option Explicit
Randomize

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
end if

if B2SOn = true then VarHidden = 1

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
    dtl.CreateEvents("dtL")

    ' Center Drop targets
    Set dtC = New cvpmDropTarget
    dtC.InitDrop Array(sw19, sw20, sw21), Array(19, 20, 21)
    dtC.InitSnd SoundFX("fx_droptarget", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
    dtC.CreateEvents("dtC")

    ' Top Drop targets
    Set dtT = New cvpmDropTarget
    dtT.InitDrop Array(sw22, sw23, sw24), Array(22, 23, 24)
    dtT.InitSnd SoundFX("fx_droptarget", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
    dtT.CreateEvents("dtT")

    ' Thalamus, more randomness pls
    ' Saucers
    Set bsSaucer2 = New cvpmBallStack
    bsSaucer2.InitSaucer Kicker2, 36, 170, 14
    bsSaucer2.InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    bsSaucer2.KickAngleVar = 30
    bsSaucer2.KickForceVar = 3
    bsSaucer2.KickAngleVar = 3

    Set bsSaucer1 = New cvpmBallStack
    bsSaucer1.InitSaucer Kicker1, 37, 330, 14
    bsSaucer1.InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    bsSaucer1.KickAngleVar = 30
    bsSaucer1.KickForceVar = 3
    bsSaucer1.KickAngleVar = 3

    ' Nudging
    vpmNudge.TiltSwitch = 7
    vpmNudge.Sensitivity = 5
    vpmNudge.TiltObj = Array(LeftSlingshot, RightSlingshot, LeftSlingshot1, RightSlingshot1)

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1
    Realtime.Enabled = 1
End Sub

Sub Table1_Paused:Controller.Pause = 1:End Sub
Sub Table1_unPaused:Controller.Pause = 0:End Sub
Sub Table1_Exit:Controller.Stop:End Sub

Sub Realtime_Timer
    GIUpdate
    RollingUpdate
End Sub

'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If vpmKeyDown(keycode) Then Exit Sub
    If keycode = PlungerKey Then PlaySoundAt "fx_PlungerPull", Plunger:Plunger.Pullback
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If vpmKeyUp(keycode) Then Exit Sub
    If keycode = PlungerKey Then PlaySoundAt "fx_plunger", Plunger:Plunger.Fire
End Sub

'*********
' Switches
'*********

' Slings & Rubbers

Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    PlaySoundAt SoundFX("fx_slingshot", DOFContactors), Lemk
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
    PlaySoundAt SoundFX("fx_slingshot", DOFContactors), Remk
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
    PlaySoundAt SoundFX("fx_slingshot", DOFContactors),Lemk1
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
    PlaySoundAt SoundFX("fx_slingshot", DOFContactors),Remk1
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
Sub Drain_Hit:bsTrough.AddBall Me:PlaySoundAt "fx_drain",Drain: End Sub
Sub Kicker1_Hit:bsSaucer1.AddBall 0:PlaySoundAt "fx_kicker-enter", Kicker1:End Sub
Sub Kicker2_Hit:bsSaucer2.AddBall 0:PlaySoundAt "fx_kicker-enter", Kicker2:End Sub

' Rollovers
Sub SW9_Hit:Controller.Switch(9) = 1:PlaySoundAt "fx_sensor",sw9:End Sub
Sub SW9_UnHit:Controller.Switch(9) = 0:End Sub
Sub SW10_Hit:Controller.Switch(10) = 1:PlaySoundAt "fx_sensor", sw10:End Sub
Sub SW10_UnHit:Controller.Switch(10) = 0:End Sub
Sub SW11_Hit:Controller.Switch(11) = 1:PlaySoundAt "fx_sensor", sw11:End Sub
Sub SW11_UnHit:Controller.Switch(11) = 0:End Sub
Sub SW12_Hit:Controller.Switch(12) = 1:PlaySoundAt "fx_sensor", sw12:End Sub
Sub SW12_UnHit:Controller.Switch(12) = 0:End Sub
Sub SW12a_Hit:Controller.Switch(12) = 1:PlaySoundAt "fx_sensor", sw12a:End Sub
Sub SW12a_UnHit:Controller.Switch(12) = 0:End Sub
Sub SW17_Hit:Controller.Switch(17) = 1:PlaySoundAt "fx_sensor", sw17:End Sub
Sub SW17_UnHit:Controller.Switch(17) = 0:End Sub
Sub SW18_Hit:Controller.Switch(18) = 1:PlaySoundAt "fx_sensor", sw18:End Sub
Sub SW18_UnHit:Controller.Switch(18) = 0:End Sub

' Targets
Sub sw8_Hit:vpmTimer.PulseSw 8:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw25_Hit:vpmTimer.PulseSw 25:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw26_Hit:vpmTimer.PulseSw 26:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw27_Hit:vpmTimer.PulseSw 27:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw28_Hit:vpmTimer.PulseSw 28:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw29_Hit:vpmTimer.PulseSw 29:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw30_Hit:vpmTimer.PulseSw 30:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw31_Hit:vpmTimer.PulseSw 31:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw32_Hit:vpmTimer.PulseSw 32:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

' Spinner
Sub Spinner1_Spin:vpmTimer.PulseSw 5:PlaySoundAt "fx_spinner",Spinner1:End Sub

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

'*******************
' Flipper Subs Rev3
'*******************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("fx_flipperup", DOFFlippers), LeftFlipper
        LeftFlipper.RotateToEnd
        LeftFlipper1.RotateToEnd
        LeftFlipperOn = 1
    Else
        PlaySoundAt SoundFX("fx_flipperdown", DOFFlippers), LeftFlipper
        LeftFlipper.RotateToStart
        LeftFlipper1.RotateToStart
        LeftFlipperOn = 0
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("fx_flipperup", DOFFlippers), RightFlipper
        RightFlipper.RotateToEnd
        RightFlipper1.RotateToEnd
        RightFlipperOn = 1
    Else
        PlaySoundAt SoundFX("fx_flipperdown", DOFFlippers), RightFlipper
        RightFlipper.RotateToStart
        RightFlipper1.RotateToStart
        RightFlipperOn = 0
    End If
End Sub

' flippers hit Sound

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, Vol(ActiveBall), pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, Vol(ActiveBall), pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub LeftFlipper1_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, Vol(ActiveBall), pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper1_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, Vol(ActiveBall), pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
End Sub

'*********************************************************
' Real Time Flipper adjustments - by JLouLouLou & JPSalas
'        (to enable flipper tricks)
'*********************************************************

Dim FlipperPower
Dim FlipperElasticity
Dim SOSTorque, SOSAngle
Dim FullStrokeEOS_Torque, LiveStrokeEOS_Torque
Dim LeftFlipperOn
Dim RightFlipperOn

Dim LLiveCatchTimer
Dim RLiveCatchTimer
Dim LiveCatchSensivity

FlipperPower = 3600
FlipperElasticity = 0.6
FullStrokeEOS_Torque = 0.6 ' EOS Torque when flipper hold up ( EOS Coil is fully charged. Ampere increase due to flipper can't move or when it pushed back when "On". EOS Coil have more power )
LiveStrokeEOS_Torque = 0.3 ' EOS Torque when flipper rotate to end ( When flipper move, EOS coil have less Ampere due to flipper can freely move. EOS Coil have less power )

LeftFlipper.EOSTorqueAngle = 10
RightFlipper.EOSTorqueAngle = 10

SOSTorque = 0.2
SOSAngle = 6

LiveCatchSensivity = 10

LLiveCatchTimer = 0
RLiveCatchTimer = 0

LeftFlipper.TimerInterval = 1
LeftFlipper.TimerEnabled = 1

Sub LeftFlipper_Timer 'flipper's tricks timer
'Start Of Stroke Flipper Stroke Routine : Start of Stroke for Tap pass and Tap shoot
    If LeftFlipper.CurrentAngle >= LeftFlipper.StartAngle - SOSAngle Then LeftFlipper.Strength = FlipperPower * SOSTorque else LeftFlipper.Strength = FlipperPower : End If

'End Of Stroke Routine : Livecatch and Emply/Full-Charged EOS
  If LeftFlipperOn = 1 Then
    If LeftFlipper.CurrentAngle = LeftFlipper.EndAngle then
      LeftFlipper.EOSTorque = FullStrokeEOS_Torque
      LLiveCatchTimer = LLiveCatchTimer + 1
      If LLiveCatchTimer < LiveCatchSensivity Then
        LeftFlipper.Elasticity = 0
      Else
        LeftFlipper.Elasticity = FlipperElasticity
        LLiveCatchTimer = LiveCatchSensivity
      End If
    End If
  Else
    LeftFlipper.Elasticity = FlipperElasticity
    LeftFlipper.EOSTorque = LiveStrokeEOS_Torque
    LLiveCatchTimer = 0
  End If


'Start Of Stroke Flipper Stroke Routine : Start of Stroke for Tap pass and Tap shoot
    If RightFlipper.CurrentAngle <= RightFlipper.StartAngle + SOSAngle Then RightFlipper.Strength = FlipperPower * SOSTorque else RightFlipper.Strength = FlipperPower : End If

'End Of Stroke Routine : Livecatch and Emply/Full-Charged EOS
  If RightFlipperOn = 1 Then
    If RightFlipper.CurrentAngle = RightFlipper.EndAngle Then
      RightFlipper.EOSTorque = FullStrokeEOS_Torque
      RLiveCatchTimer = RLiveCatchTimer + 1
      If RLiveCatchTimer < LiveCatchSensivity Then
        RightFlipper.Elasticity = 0
      Else
        RightFlipper.Elasticity = FlipperElasticity
        RLiveCatchTimer = LiveCatchSensivity
      End If
    End If
  Else
    RightFlipper.Elasticity = FlipperElasticity
    RightFlipper.EOSTorque = LiveStrokeEOS_Torque
    RLiveCatchTimer = 0
  End If
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

'**********************************************************
'     JP's Lamp Fading for VPX and Vpinmame v4.0
' FadingStep used for all kind of lamps
' FlashLevel used for modulated flashers
' LampState keep the real lamp state in a array
'**********************************************************

Dim LampState(200), FadingStep(200), FlashLevel(200)

InitLamps() ' turn off the lights and flashers and reset them to the default parameters

' vpinmame Lamp & Flasher Timers

Sub LampTimer_Timer()
    Dim chgLamp, num, chg, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp)Then
        For ii = 0 To UBound(chgLamp)
            LampState(chgLamp(ii, 0)) = chgLamp(ii, 1)       'keep the real state in an array
            FadingStep(chgLamp(ii, 0)) = chgLamp(ii, 1)
        Next
    End If

    UpdateLeds
    UpdateLamps
End Sub

Sub UpdateLamps

    Lamp 4, l4

    Lamp 6, l6
    Lamp 7, l7
    Lamp 8, l8
    Lamp 9, l9
    Lamp 10, l10
    Lamp 11, l11
    Lamp 12, l12
    Lamp 14, l14

    Lamp 20, l20

    Lamp 22, l22
    Lamp 23, l23
    Lamp 24, l24
    Lamp 25, l25
    Lamp 26, l26
    Lampm 27, l27a
    Lamp 27, l27
    Lamp 28, l28
    Lampm 29, l29a
    Lamp 29, l29
    Lamp 30, l30

    Lamp 36, l36

    Lamp 38, l38
    Lamp 39, l39
    Lamp 40, l40
    Lamp 41, l41
    Lamp 42, l42
    Lampm 43, l43a
    Lamp 43, l43
    Lamp 44, l44
    Lamp 46, l46
    Lamp 47, l47

    Lamp 51, l51
    Lamp 52, l52
    Lamp 53, l53
    Lamp 54, l54
    Lamp 55, l55
    Lamp 56, l56
    Lamp 57, l57
    Lamp 58, l58
    Lamp 59, l59
    Lamp 60, l60
    Lamp 62, l62

    ' backdrop lights
    If VarHidden Then
        Text 13, l13, "Highscore"
        'Text 29, l29, "Ball in Play"
        'Text 45, l45, "Game Over"
        'Text 61, l61, "Tilt"
        Text 63, l63, "Match"
    End If
End Sub

' div lamp subs

' Normal Lamp & Flasher subs

Sub InitLamps()
    Dim x
    LampTimer.Interval = 40
    LampTimer.Enabled = 1
    For x = 0 to 200
        FadingStep(x) = 0
        FlashLevel(x) = 0
    Next
End Sub

Sub SetLamp(nr, value) ' 0 is off, 1 is on
    FadingStep(nr) = abs(value)
End Sub

' Lights: used for VPX standard lights, the fading is handled by VPX itself, they are here to be able to make them work together with the flashers

Sub Lamp(nr, object)
    Select Case FadingStep(nr)
        Case 1:object.state = 1:FadingStep(nr) = -1
        Case 0:object.state = 0:FadingStep(nr) = -1
    End Select
End Sub

Sub Lampm(nr, object) ' used for multiple lights, it doesn't change the fading state
    Select Case FadingStep(nr)
        Case 1:object.state = 1
        Case 0:object.state = 0
    End Select
End Sub

' Flashers:  0 starts the fading until it is off

Sub Flash(nr, object)
    Dim tmp
    Select Case FadingStep(nr)
        Case 1:Object.IntensityScale = 1:FadingStep(nr) = -1
        Case 0
            tmp = Object.IntensityScale * 0.85 - 0.01
            If tmp > 0 Then
                Object.IntensityScale = tmp
            Else
                Object.IntensityScale = 0
                FadingStep(nr) = -1
            End If
    End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it doesn't change the fading state
    Dim tmp
    Select Case FadingStep(nr)
        Case 1:Object.IntensityScale = 1
        Case 0
            tmp = Object.IntensityScale * 0.85 - 0.01
            If tmp > 0 Then
                Object.IntensityScale = tmp
            Else
                Object.IntensityScale = 0
            End If
    End Select
End Sub

' Desktop Objects: Reels & texts

' Reels - 4 steps fading
Sub Reel(nr, object)
    Select Case FadingStep(nr)
        Case 1:object.SetValue 1:FadingStep(nr) = -1
        Case 0:object.SetValue 2:FadingStep(nr) = 2
        Case 2:object.SetValue 3:FadingStep(nr) = 3
        Case 3:object.SetValue 0:FadingStep(nr) = -1
    End Select
End Sub

Sub Reelm(nr, object)
    Select Case FadingStep(nr)
        Case 1:object.SetValue 1
        Case 0:object.SetValue 2
        Case 2:object.SetValue 3
        Case 3:object.SetValue 0
    End Select
End Sub

' Reels non fading
Sub NfReel(nr, object)
    Select Case FadingStep(nr)
        Case 1:object.SetValue 1:FadingStep(nr) = -1
        Case 0:object.SetValue 0:FadingStep(nr) = -1
    End Select
End Sub

Sub NfReelm(nr, object)
    Select Case FadingStep(nr)
        Case 1:object.SetValue 1
        Case 0:object.SetValue 0
    End Select
End Sub

'Texts

Sub Text(nr, object, message)
    Select Case FadingStep(nr)
        Case 1:object.Text = message:FadingStep(nr) = -1
        Case 0:object.Text = "":FadingStep(nr) = -1
    End Select
End Sub

Sub Textm(nr, object, message)
    Select Case FadingStep(nr)
        Case 1:object.Text = message
        Case 0:object.Text = ""
    End Select
End Sub

' Modulated Subs for the WPC tables

Sub SetModLamp(nr, level)
    FlashLevel(nr) = level / 150 'lights & flashers
End Sub

Sub LampMod(nr, object)          ' modulated lights used as flashers
    Object.IntensityScale = FlashLevel(nr)
    Object.State = 1             'in case it was off
End Sub

Sub FlashMod(nr, object)         'sets the flashlevel from the SolModCallback
    Object.IntensityScale = FlashLevel(nr)
End Sub

'Walls, flashers, ramps and Primitives used as 4 step fading images
'a,b,c,d are the images used from on to off

Sub FadeObj(nr, object, a, b, c, d)
    Select Case FadingStep(nr)
        Case 1:object.image = a:FadingStep(nr) = -1
        Case 0:object.image = b:FadingStep(nr) = 2
        Case 2:object.image = c:FadingStep(nr) = 3
        Case 3:object.image = d:FadingStep(nr) = -1
    End Select
End Sub

Sub FadeObjm(nr, object, a, b, c, d)
    Select Case FadingStep(nr)
        Case 1:object.image = a
        Case 0:object.image = b
        Case 2:object.image = c
        Case 3:object.image = d
    End Select
End Sub

Sub NFadeObj(nr, object, a, b)
    Select Case FadingStep(nr)
        Case 1:object.image = a:FadingStep(nr) = -1
        Case 0:object.image = b:FadingStep(nr) = -1
    End Select
End Sub

Sub NFadeObjm(nr, object, a, b)
    Select Case FadingStep(nr)
        Case 1:object.image = a
        Case 0:object.image = b
    End Select
End Sub

'*********************************
' Diverse Collection Hit Sounds
'*********************************

Sub aMetals_Hit(idx):PlaySoundAtBall "fx_MetalHit":End Sub
Sub aMetalWires_Hit(idx):PlaySoundAtBall "fx_MetalWire":End Sub
Sub aRubber_Bands_Hit(idx):PlaySoundAtBall "fx_rubber_band":End Sub
Sub aRubber_LongBands_Hit(idx):PlaySoundAtBall "fx_rubber_longband":End Sub
Sub aRubber_Posts_Hit(idx):PlaySoundAtBall "fx_rubber_post":End Sub
Sub aRubber_Pins_Hit(idx):PlaySoundAtBall "fx_rubber_pin":End Sub
Sub aRubber_Pegs_Hit(idx):PlaySoundAtBall "fx_rubber_peg":End Sub
Sub aPlastics_Hit(idx):PlaySoundAtBall "fx_PlasticHit":End Sub
Sub aGates_Hit(idx):PlaySoundAtBall "fx_Gate":End Sub
Sub aWoods_Hit(idx):PlaySoundAtBall "fx_Woodhit":End Sub
Sub aDroptargets_Hit(idx):PlaySound SoundFX("fx_droptarget",DOFDropTargets), 0, 1, Pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall):End Sub
Sub aTargets_Hit(idx):PlaySound SoundFX("fx_target",DOFTargets), 0, Vol(ActiveBall), pan(ActiveBall), 0.2, Pitch(ActiveBall) * 10, 0, 0, AudioFade(ActiveBall):End Sub
Sub aTriggers_Hit(idx): PlaySoundAt "fx_sensor", aTriggers(idx):End Sub

'***************************************************************
'             Supporting Ball & Sound Functions v4.0
'***************************************************************

Dim TableWidth, TableHeight

TableWidth = Table1.width
TableHeight = Table1.height

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / TableWidth-1
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
    BallVel = (SQR((ball.VelX ^2) + (ball.VelY ^2)))
End Function

Function AudioFade(ball) 'only on VPX 10.4 and newer
    Dim tmp
    tmp = ball.y * 2 / TableHeight-1
    If tmp > 0 Then
        AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10))
    End If
End Function

Sub PlaySoundAt(soundname, tableobj) 'play sound at X and Y position of an object, mostly bumpers, flippers and other fast objects
    PlaySound soundname, 0, 1, Pan(tableobj), 0.2, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname) ' play a sound at the ball position, like rubbers, targets, metals, plastics
    PlaySound soundname, 0, Vol(ActiveBall), pan(ActiveBall), 0.2, Pitch(ActiveBall) * 10, 0, 0, AudioFade(ActiveBall)
End Sub

Function RndNbr(n) 'returns a random number between 1 and n
    Randomize timer
    RndNbr = Int((n * Rnd) + 1)
End Function

'***********************************************
'   JP's VP10 Rolling Sounds + Ballshadow v4.0
'   uses a collection of shadows, aBallShadow
'***********************************************

Const tnob = 19   'total number of balls
Const lob = 0     'number of locked balls
Const maxvel = 34 'max ball velocity
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub RollingUpdate()
    Dim BOT, b, ballpitch, ballvol, speedfactorx, speedfactory
    BOT = GetBalls

    ' stop the sound of deleted balls
    For b = UBound(BOT) + 1 to tnob
        rolling(b) = False
        StopSound("fx_ballrolling" & b)
        aBallShadow(b).Y = 3000
    Next

    ' exit the sub if no balls on the table
    If UBound(BOT) = lob - 1 Then Exit Sub 'there no extra balls on this table

    ' play the rolling sound for each ball and draw the shadow
    For b = lob to UBound(BOT)
        aBallShadow(b).X = BOT(b).X
        aBallShadow(b).Y = BOT(b).Y
        aBallShadow(b).Height = BOT(b).Z -Ballsize/2

        If BallVel(BOT(b))> 1 Then
            If BOT(b).z <30 Then
                ballpitch = Pitch(BOT(b))
                ballvol = Vol(BOT(b))
            Else
                ballpitch = Pitch(BOT(b)) + 10000 'increase the pitch on a ramp
                ballvol = Vol(BOT(b)) * 2
            End If
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, ballvol, Pan(BOT(b)), 0, ballpitch, 1, 0, AudioFade(BOT(b))
        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If

        ' rothbauerw's Dropping Sounds
        If BOT(b).VelZ <-1 and BOT(b).z <55 and BOT(b).z> 27 Then 'height adjust for ball drop sounds
            PlaySound "fx_balldrop", 0, ABS(BOT(b).velz) / 17, Pan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
        End If

        ' jps ball speed & spin control
            BOT(b).AngMomZ = BOT(b).AngMomZ * 0.95
        If BOT(b).VelX AND BOT(b).VelY <> 0 Then
            speedfactorx = ABS(maxvel / BOT(b).VelX)
            speedfactory = ABS(maxvel / BOT(b).VelY)
            If speedfactorx <1 Then
                BOT(b).VelX = BOT(b).VelX * speedfactorx
                BOT(b).VelY = BOT(b).VelY * speedfactorx
            End If
            If speedfactory <1 Then
                BOT(b).VelX = BOT(b).VelX * speedfactory
                BOT(b).VelY = BOT(b).VelY * speedfactory
            End If
        End If
    Next
End Sub

'*****************************
' Ball 2 Ball Collision Sound
'*****************************

Sub OnBallBallCollision(ball1, ball2, velocity)
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

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

'*********************************
' Table Options F12 User Options
'*********************************
' Table1.Option arguments are:
' - option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional array of literal strings

Dim LUTImage

Sub Table1_OptionEvent(ByVal eventId)
    Dim x, y

    'LUT
    LutImage = Table1.Option("Select LUT", 0, 21, 1, 0, 0, Array("Normal 0", "Normal 1", "Normal 2", "Normal 3", "Normal 4", "Normal 5", "Normal 6", "Normal 7", "Normal 8", "Normal 9", "Normal 10", _
        "Warm 0", "Warm 1", "Warm 2", "Warm 3", "Warm 4", "Warm 5", "Warm 6", "Warm 7", "Warm 8", "Warm 9", "Warm 10") )
    UpdateLUT

    ' Cabinet rails
    x = Table1.Option("Cabinet Rails", 0, 1, 1, 1, 0, Array("Hide", "Show") )
    For each y in aRails:y.visible = x:next
End Sub

Sub UpdateLUT
    Select Case LutImage
        Case 0:table1.ColorGradeImage = "LUT0"
        Case 1:table1.ColorGradeImage = "LUT1"
        Case 2:table1.ColorGradeImage = "LUT2"
        Case 3:table1.ColorGradeImage = "LUT3"
        Case 4:table1.ColorGradeImage = "LUT4"
        Case 5:table1.ColorGradeImage = "LUT5"
        Case 6:table1.ColorGradeImage = "LUT6"
        Case 7:table1.ColorGradeImage = "LUT7"
        Case 8:table1.ColorGradeImage = "LUT8"
        Case 9:table1.ColorGradeImage = "LUT9"
        Case 10:table1.ColorGradeImage = "LUT10"
        Case 11:table1.ColorGradeImage = "LUT Warm 0"
        Case 12:table1.ColorGradeImage = "LUT Warm 1"
        Case 13:table1.ColorGradeImage = "LUT Warm 2"
        Case 14:table1.ColorGradeImage = "LUT Warm 3"
        Case 15:table1.ColorGradeImage = "LUT Warm 4"
        Case 16:table1.ColorGradeImage = "LUT Warm 5"
        Case 17:table1.ColorGradeImage = "LUT Warm 6"
        Case 18:table1.ColorGradeImage = "LUT Warm 7"
        Case 19:table1.ColorGradeImage = "LUT Warm 8"
        Case 20:table1.ColorGradeImage = "LUT Warm 9"
        Case 21:table1.ColorGradeImage = "LUT Warm 10"
    End Select
End Sub

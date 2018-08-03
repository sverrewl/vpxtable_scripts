' Medusa - Bally 1981
' http://www.ipdb.org/machine.cgi?id=1565
' Medusa / IPD No. 1565 / February 04, 1981 / 4 Players
' VPX version by JPSalas 2017, version 1.0.4
' Light numbers from the tables by Joe Entropy & RipleYYY and Pacdude.
' Uses the Left and Right Magna saves keys to activate the save post (shield)


' Thalamus 2018-08-03
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Fix from DJRobX is included
' No special SSF tweaks yet.
' This is a JP table. He often uses walls as switches so I need to be careful of using PlaySoundAt

Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const BallSize = 50

LoadVPM "01550000", "Bally.vbs", 3.26

Dim bsTrough, bsSaucer, dtRBank, dtTBank
Dim x

Const cGameName = "medusa"

Const UseSolenoids = 0
Const UseLamps = 0
Const UseGI = 0
Const UseSync = 0
Const HandleMech = 0

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

Set MotorCallback = GetRef("UpdateSolenoids")

' Standard Sounds
Const SSolenoidOn = "fx_Solenoid"
Const SSolenoidOff = ""
Const SCoin = "fx_coin"

'************
' Table init.
'************

Sub table1_Init
    vpmInit me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Medusa - Bally 1981" & vbNewLine & "VPX table by JPSalas v.1.0.4"
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = VarHidden
        '.Games(cGameName).Settings.Value("rol") = 0 '1= rotated display, 0= normal
        '.SetDisplayPosition 0,0, GetPlayerHWnd 'restore dmd window position
        On Error Resume Next
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
        On Error Goto 0
    End With

    ' Nudging
    vpmNudge.TiltSwitch = 15
    vpmNudge.Sensitivity = 1
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, Bumper4, LeftSlingshot, RightSlingshot, sw34)

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0, 8, 0, 0, 0, 0, 0, 0
        .InitKick BallRelease, 80, 6
        .InitAddSnd "fx_drain"
        .InitEntrySnd "fx_Solenoid", "fx_Solenoid"
        .InitExitSnd "fx_ballrel", "fx_Solenoid"
        .Balls = 1
        .IsTrough = True
    End With

    ' Saucer
    Set bsSaucer = New cvpmBallStack
    With bsSaucer
        .InitSaucer sw19, 19, 155, 8
        .InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .KickAngleVar = 2
        .KickForceVar = 2
    End With

    ' Right drop targets
    Set dtRBank = New cvpmDropTarget
    With dtRbank
        .initdrop array(sw1, sw2, sw3, sw4), array(1, 2, 3, 4)
        .initsnd SoundFX("fx_droptarget", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
    '.createEvents "dtRbank"
    End With

    'Top drop targets
    Set dtTBank = New cvpmDropTarget
    With dtTbank
        .initdrop array(sw48, sw47, sw46, sw45, sw44, sw43, sw42), Array(48, 47, 46, 45, 44, 43, 42)
        .initsnd SoundFX("fx_droptarget", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
    '.createEvents "dtTbank"
    End With

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

    ' init shield
    post2.IsDropped = 1
    post2rubber.visible = 0
    Post.Pullback

	' Manually init fast flips
	if not IsEmpty(SolCallback(sLLFlipper)) then vpmFlips.CallBackL = SolCallback(sLLFlipper)         
	if not IsEmpty(SolCallback(sLRFlipper)) then vpmFlips.CallBackR = SolCallback(sLRFlipper)
End Sub

Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub

'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
    If keycode = PlungerKey Then PlaySound "fx_PlungerPull", 0, 1, 0.1, 0.25:Plunger.Pullback
    If keycode = LeftMagnaSave OR keycode = RightMagnaSave Then Controller.Switch(17) = 1
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If vpmKeyDown(keycode)Then Exit Sub
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If keycode = LeftMagnaSave OR keycode = RightMagnaSave Then Controller.Switch(17) = 0
    If vpmKeyUp(keycode)Then Exit Sub
    If keycode = PlungerKey Then PlaySound "fx_plunger", 0, 1, 0.1, 0.25:Plunger.Fire
End Sub

'*********
' Switches
'*********

' Slings
Dim LStep, RStep, R2Step

Sub LeftSlingShot_Slingshot
    PlaySound SoundFX("fx_slingshot", DOFContactors), 0, 1, -0.05, 0.05
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
    PlaySound SoundFX("fx_slingshot", DOFContactors), 0, 1, 0.05, 0.05
    RightSling4.Visible = 1
    Remk.RotX = 26
    RStep = 0
    vpmTimer.PulseSw 35
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

Sub sw34_Slingshot
    PlaySound SoundFX("fx_slingshot", DOFContactors), 0, 1, 0.05, 0.05
    R2Sling4.Visible = 1
    Remk1.RotX = 26
    R2Step = 0
    vpmTimer.PulseSw 34
    sw34.TimerEnabled = 1
End Sub

Sub sw34_Timer
    Select Case R2Step
        Case 1:R2SLing4.Visible = 0:R2SLing3.Visible = 1:Remk1.RotX = 14
        Case 2:R2SLing3.Visible = 0:R2SLing2.Visible = 1:Remk1.RotX = 2
        Case 3:R2SLing2.Visible = 0:Remk1.RotX = -10:sw34.TimerEnabled = 0
    End Select
    R2Step = R2Step + 1
End Sub

' Rubbers
Sub sw18_Hit:vpmTimer.PulseSw 18:End Sub
Sub sw18a_Hit:vpmTimer.PulseSw 18:End Sub
Sub sw18b_Hit:vpmTimer.PulseSw 18:End Sub

' Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 37:PlaySound SoundFX("fx_bumper", DOFContactors), 0, 1, -0.15, 0.15:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 38:PlaySound SoundFX("fx_bumper", DOFContactors), 0, 1, -0.15, 0.15:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 39:PlaySound SoundFX("fx_bumper", DOFContactors), 0, 1, 0.15, 0.15:End Sub
Sub Bumper4_Hit:vpmTimer.PulseSw 40:PlaySound SoundFX("fx_bumper", DOFContactors), 0, 1, 0.15, 0.15:End Sub

' Drain holes
Sub Drain_Hit:bsTrough.AddBall Me:End Sub

'Saucer
Sub sw19_Hit:bsSaucer.AddBall 0:End Sub

' Rollovers
Sub sw32_Hit:Controller.Switch(32) = true:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw32_UnHit:Controller.Switch(32) = false:End Sub
Sub sw32a_Hit:Controller.Switch(32) = true:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw32a_UnHit:Controller.Switch(32) = false:End Sub
Sub sw20_Hit():controller.switch(20) = true:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw20_unHit():controller.switch(20) = false:End Sub
Sub sw21_Hit():controller.switch(21) = true:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw21_unHit():controller.switch(21) = false:End Sub
Sub sw22_Hit():controller.switch(22) = true:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw22_unHit():controller.switch(22) = false:End Sub
Sub sw23_Hit():controller.switch(23) = true:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw23_unHit():controller.switch(23) = false:End Sub
Sub sw24_Hit():controller.switch(24) = true:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw24_unHit():controller.switch(24) = false:End Sub
Sub sw25_Hit():controller.switch(25) = true:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw25_unHit():controller.switch(25) = false:End Sub
Sub sw25a_Hit():controller.switch(25) = true:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw25a_unHit():controller.switch(25) = false:End Sub
Sub sw26_Hit():controller.switch(26) = true:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw26_unHit():controller.switch(26) = false:End Sub
Sub sw28_Hit():controller.switch(28) = true:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw28_unHit():controller.switch(28) = false:End Sub
Sub sw31_Hit():SetLamp 150, 1:controller.switch(31) = true:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw31_unHit():SetLamp 150, 0:controller.switch(31) = false:End Sub
Sub sw31a_Hit():SetLamp 151, 1:controller.switch(31) = true:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw31a_unHit():SetLamp 151, 0:controller.switch(31) = false:End Sub

' Targets
Sub sw12_Hit:vpmTimer.PulseSw 12:PlaySound SoundFX("fx_target", DOFTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw13_Hit:vpmTimer.PulseSw 13:PlaySound SoundFX("fx_target", DOFTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw14_Hit:vpmTimer.PulseSw 14:PlaySound SoundFX("fx_target", DOFTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw29_Hit:vpmTimer.PulseSw 29:PlaySound SoundFX("fx_target", DOFTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw30_Hit:vpmTimer.PulseSw 30:PlaySound SoundFX("fx_target", DOFTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw27_Hit:vpmTimer.PulseSw 27:PlaySound SoundFX("fx_target", DOFTargets), 0, 1, pan(ActiveBall):End Sub

' Droptargets
Sub sw1_Dropped:dtRbank.Hit 1:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, 0.1:End Sub
Sub sw2_Dropped:dtRbank.Hit 2:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, 0.1:End Sub
Sub sw3_Dropped:dtRbank.Hit 3:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, 0.1:End Sub
Sub sw4_Dropped:dtRbank.Hit 4:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, 0.1:End Sub

Sub sw48_Dropped:dtTbank.Hit 1:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, -0.1:End Sub
Sub sw47_Dropped:dtTbank.Hit 2:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, -0.08:End Sub
Sub sw46_Dropped:dtTbank.Hit 3:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, -0.04:End Sub
Sub sw45_Dropped:dtTbank.Hit 4:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, 0:End Sub
Sub sw44_Dropped:dtTbank.Hit 5:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, 0.04:End Sub
Sub sw43_Dropped:dtTbank.Hit 6:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, 0.08:End Sub
Sub sw42_Dropped:dtTbank.Hit 7:PlaySound SoundFX("fx_droptarget", DOFDropTargets), 0, 1, 0.1:End Sub

'Spinner
Sub sw33_Spin:vpmTimer.PulseSw 33:PlaySound "spinner":End Sub

'*********
'Solenoids
'*********

Sub UpdateSolenoids
    Dim Changed, ii, solNo
    Changed = Controller.ChangedSolenoids
    If Not IsEmpty(Changed)Then
        For ii = 0 To UBound(Changed)
            solNo = Changed(ii, CHGNO)
            If Controller.Lamp(34)Then
                If SolNo = 1 Or SolNo = 2 Or SolNo = 3 Or SolNo = 4 Or Solno = 5 Or Solno = 6 Or Solno = 8 Then solNo = solNo + 24 '1->25 etc
            End If
			if solNo = GameOnSolenoid then vpmFlips.TiltSol cbool(Changed(ii, CHGSTATE))
            vpmDoSolCallback solNo, Changed(ii, CHGSTATE)
        Next
    End If
End Sub

SolCallback(1) = "vpmSolSound ""fx_knocker"","
SolCallback(2) = "dtRBank.SolDropUp"
SolCallback(3) = "dtTBank.SolDropUp"
SolCallback(7) = "SolShieldPost"
SolCallback(8) = "bsTrough.SolOut"
SolCallback(19) = "RelayAC"
SolCallback(25) = "SolZipOpen"
SolCallback(26) = "SolZipClose"
SolCallback(30) = "dtTBank.SolHit 1,"
SolCallback(6) = "dtTBank.SolHit 2,"
SolCallback(29) = "dtTBank.SolHit 3,"
SolCallback(5) = "dtTBank.SolHit 4,"
SolCallback(28) = "dtTBank.SolHit 5,"
SolCallback(4) = "dtTBank.SolHit 6,"
SolCallback(27) = "dtTBank.SolHit 7,"
SolCallback(32) = "bsSaucer.SolOut"

Sub SolZipOpen(enabled)
    PlaySound "SolenoidOn"
    LeftFlipper2.Visible = True:LeftFlipper2.Enabled = True
    RightFlipper2.Visible = True:RightFlipper2.Enabled = True
    LeftFlipper3.Visible = False:LeftFlipper3.Enabled = False
    RightFlipper3.Visible = False:RightFlipper3.Enabled = False
    'LeftHelp.isdropped = True
    'RightHelp.isdropped = True
    Controller.Switch(41) = 1
End Sub

Sub SolZipClose(enabled)
    PlaySound "SolenoidOff"
    LeftFlipper2.Visible = False:LeftFlipper2.Enabled = False
    RightFlipper2.Visible = False:RightFlipper2.Enabled = False
    LeftFlipper3.Visible = True:LeftFlipper3.Enabled = True
    RightFlipper3.Visible = true:RightFlipper3.Enabled = true
    'LeftHelp.isdropped = False
    'RightHelp.isdropped = False
    Controller.Switch(41) = 0
End Sub

Sub SolShieldPost(Enabled)
    If Enabled Then
        PlaySound "fx_autoplunger"
        post1.IsDropped = 1
        post1Rubber.Visible = 0
        post2.IsDropped = 0
        post2Rubber.Visible = 1
        Post.Fire
    Else
        post1.IsDropped = 0
        post1Rubber.Visible = 1
        post2.IsDropped = 1
        post2Rubber.Visible = 0
        Post.PullBack
    End If
End Sub

Sub RelayAC(Enabled)
    vpmNudge.SolGameOn Enabled
    If Enabled Then
        GiOn
        SetLamp 152, 1
        SetLamp 153, 1
    Else
        GiOff
        SetLamp 152, 0
        SetLamp 153, 0
    End If
End Sub

'**************
' Flipper Subs
'**************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySound SoundFX("fx_flipperup", DOFFlippers), 0, 1, -0.1, 0.25
        LeftFlipper.RotateToEnd:LeftFlipper2.RotateToEnd:LeftFlipper3.RotateToEnd:SetLamp 152, 0
    Else
        PlaySound SoundFX("fx_flipperdown", DOFFlippers), 0, 1, -0.1, 0.25
        LeftFlipper.RotateToStart:LeftFlipper2.RotateToStart:LeftFlipper3.RotateToStart:SetLamp 152, 1
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySound SoundFX("fx_flipperup", DOFFlippers), 0, 1, 0.1, 0.25
        RightFlipper.RotateToEnd:RightFlipper2.RotateToEnd:RightFlipper3.RotateToEnd:SetLamp 153, 0
    Else
        PlaySound SoundFX("fx_flipperdown", DOFFlippers), 0, 1, 0.1, 0.25
        RightFlipper.RotateToStart:RightFlipper2.RotateToStart:RightFlipper3.RotateToStart:SetLamp 153, 1
    End If
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, -0.1, 0.25
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, -0.1, 0.25
End Sub

Sub LeftFlipper2_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, -0.1, 0.25
End Sub

Sub RightFlipper2_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, -0.1, 0.25
End Sub

Sub LeftFlipper3_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, -0.1, 0.25
End Sub

Sub RightFlipper3_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, -0.1, 0.25
End Sub

'*****************
'   Gi Effects
'*****************

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
        x.Duration 2, 2000, 1
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
    UpdateLeds
    UpdateLamps
    GIUpdate
    RollingUpdate
End Sub

Sub UpdateLamps()
    NFadeL 1, l1
    NFadeL 2, l2
    NFadeL 3, l3
    NFadeLm 4, l4_1
    NFadeL 4, l4
    NFadeL 5, l5
    NFadeL 6, l6
    NFadeL 7, l7
    NFadeL 8, l8
    NFadeL 9, l9
    NFadeL 10, l10

    NFadeL 12, l12

    NFadeL 14, l14
    NFadeL 15, l15
    NFadeL 17, l17
    NFadeL 18, l18
    NFadeL 19, l19
    NFadeL 20, l20
    NFadeLm 21, l21_1
    NFadeL 21, l21
    NFadeL 22, l22
    NFadeL 23, l23
    NFadeL 24, l24
    NFadeL 25, l25
    NFadeL 26, l26

    NFadeL 28, l28

    NFadeL 30, l30
    NFadeL 31, l31
    NFadeL 33, l33
    NFadeL 35, l35
    NFadeL 36, l36
    NFadeLm 37, l37_1
    NFadeL 37, l37
    NFadeL 38, l38
    NFadeL 39, l39
    NFadeL 40, l40
    NFadeL 41, l41
    NFadeL 42, l42
    NFadeL 43, l43
    NFadeL 44, l44

    NFadeL 46, l46
    NFadeLm 47, l47a
    NFadeLm 47, l47b
    NFadeLm 47, l47c
    NFadeLm 47, l47d
    NFadeLm 47, l47e
    NFadeL 47, l47
    NFadeL 49, l49
    NFadeL 51, l51
    NFadeL 52, l52
    NFadeLm 53, l53_1
    NFadeL 53, l53
    NFadeL 54, l54
    NFadeL 55, l55
    NFadeL 56, l56
    NFadeL 57, l57
    NFadeL 58, l58
    NFadeL 59, l59

    NFadeL 62, l62
    NFadeL 63, l63
    NFadeLm 70, l70_1
    NFadeL 70, l70
    NFadeLm 71, l71_1
    NFadeL 71, l71
    NFadeLm 86, l86_1
    NFadeL 86, l86
    NFadeL 87, l87
    NFadeLm 102, l102_1
    NFadeL 102, l102
    NFadeL 103, l103
    NFadeLm 118, l118_1
    NFadeL 118, l118
    NFadeL 119, l119

    NFadeL 65, l65
    NFadeL 81, l81
    NFadeL 97, l97
    NFadeL 113, l113
    NFadeL 66, l66
    NFadeL 82, l82
    NFadeL 98, l98
    NFadeL 114, l114
    NFadeL 67, l67
    NFadeL 83, l83
    NFadeL 99, l99
    NFadeL 115, l115
    NFadeL 68, l68
    NFadeL 84, l84
    NFadeL 100, l100
    NFadeL 116, l116
    NFadeL 69, l69
    NFadeL 85, l85
    NFadeL 101, l101
    NFadeL 117, l117
    NFadeL 60, l60

    NFadeL 150, lleft
    NFadeL 151, lright
    NFadeLm 152, flipperlight1
    NFadeLm 152, flipperlight2
    NFadeLm 152, flipperlight3
    NFadeL 152, flipperlight4
    NFadeLm 153, flipperlight5
    NFadeLm 153, flipperlight6
    NFadeLm 153, flipperlight7
    NFadeL 153, flipperlight8
    ' backdrop lights
    If VarHidden Then
        NFadeT 11, l11, "SAME PLAYER SHOOTS AGAIN"
        NFadeT 13, l13, "BALL IN PLAY"
        NFadeT 27, l27, "MATCH"
        NFadeT 29, l29, "HI SCORE TO DATE"
        NFadeT 45, l45, "GAME OVER"
        NFadeT 61, l61, "TILT"
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

'************************************
'          LEDs Display
'     Based on Scapino's LEDs
'************************************

Dim Digits(38)
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

Digits(32) = Array(a00, a02, a05, a06, a04, a01, a03)
Digits(33) = Array(a10, a12, a15, a16, a14, a11, a13)
Digits(34) = Array(a20, a22, a25, a26, a24, a21, a23)
Digits(35) = Array(a30, a32, a35, a36, a34, a31, a33)
Digits(36) = Array(a40, a42, a45, a46, a44, a41, a43)
Digits(37) = Array(a50, a52, a55, a56, a54, a51, a53)

Sub UpdateLeds
    On Error Resume Next
    Dim ChgLED, ii, jj, chg, stat, num, obj
    ChgLED = Controller.ChangedLEDs(&H0000003f, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
        For ii = 0 To UBound(ChgLED)
            num = chgLED(ii, 0):chg = chgLED(ii, 1):stat = chgLED(ii, 2)
            If num < 32 Then
                For jj = 0 to 10
                    If stat = Patterns(jj)OR stat = Patterns2(jj)then Digits(chgLED(ii, 0)).SetValue jj
                Next
            Else
                For Each obj In Digits(num)
                    If chg And 1 Then obj.State = stat And 1
                    chg = chg \ 2:stat = stat \ 2
                Next
            End If
        Next
    End IF
End Sub

'Bally Medusa
'added by Inkochnito
Sub editDips
    Dim vpmDips:Set vpmDips = New cvpmDips
    With vpmDips
        .AddForm 700, 400, "Medusa - DIP switches"
        .AddChk 0, 5, 120, Array("Match feature", &H08000000)                                                                                                                                                                                                                 'dip 28
        .AddChk 0, 25, 120, Array("Credits displayed", &H04000000)                                                                                                                                                                                                            'dip 27
        .AddFrame 0, 44, 190, "Extra ball match number display adjust", &H00400000, Array("any number flashing will be reset", 0, "any number flashing will for next ball", &H00400000)                                                                                       'dip 23
        .AddFrame 0, 90, 190, "Top Olympus red lites", &H00000020, Array("step back 1 when target is hit", 0, "do not step back", &H00000020)                                                                                                                                 'dip 6
        .AddFrame 0, 136, 190, "Any advanced Colossus bonus lite on", &H00000040, Array("will be reset", 0, "will come back on for next ball", &H00000040)                                                                                                                    'dip 7
        .AddFrame 0, 182, 190, "Any lit left side 2 or 3 arrow", &H00004000, Array("will be reset", 0, "will come back on for next ball", &H00004000)                                                                                                                         'dip 15
        .AddFrame 0, 228, 190, "Medusa special lites with", 32768, Array("80K", 0, "40K and 80K", 32768)                                                                                                                                                                      'dip 16
        .AddFrame 0, 274, 190, "Collect Olympus bonus saucer lite", &H20000000, Array("will be reset", 0, "will come back on for next ball", &H20000000)                                                                                                                      'dip 30
        .AddFrame 0, 320, 395, "Movable flipper timer adjust", &H00000080, Array("closed flippers will open after 10 seconds until next target is hit", 0, "Hitting top targets adds 10 seconds each to keep flippers closed", &H00000080)                                    'dip 8
        .AddFrame 205, 0, 190, "Maximum credits", &H03000000, Array("10 credits", 0, "15 credits", &H01000000, "25 credits", &H02000000, "40 credits", &H03000000)                                                                                                            'dip 25&26
        .AddFrame 205, 76, 190, "Balls per game", &HC0000000, Array("2 balls", &HC0000000, "3 balls", 0, "4 balls", &H80000000, "5 balls", &H40000000)                                                                                                                        'dip 31&32
        .AddFrame 205, 152, 190, "Olympus bonus red lights", &H00100000, Array("1st 5-3, 2nd 5-2, 3dr 5-1, 4th 5-1", 0, "1st 5-3, 2nd 5-2, 3dr 5-2, 4th 5-2", &H00100000, "1st 5-3, 2nd 5-3, 3dr 5-2, 4th 5-2", &H00200000, "1st 5-3, 2nd 5-3, 3dr 5-3, 4th 5-3", &H00300000) 'dip 21&22
        .AddFrame 205, 228, 190, "Medusa bonus from 1 to 19 memory", &H00800000, Array("will be reset", 0, "will come back on for next ball", &H00800000)                                                                                                                     'dip 24
        .AddFrame 205, 274, 190, "Replay limit", &H10000000, Array("1 replay per game", 0, "unlimited replays", &H10000000)                                                                                                                                                   'dip 29
        .AddLabel 30, 390, 350, 15, "Set selftest position 16,17,18 and 19 to 03 for the best gameplay."
        .AddLabel 50, 410, 300, 15, "After hitting OK, press F3 to reset game with new settings."
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

Sub Table1_Exit():Controller.Stop:End Sub

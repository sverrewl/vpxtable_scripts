' Phoenix / IPD No. 1780 / August 25, 1978 / 4 Players
' VPX by jpsalas, graphics by halen, vpinmame script based on the table by Andre Needham, Kristian & PD¨
' DOF by arngrim

Option Explicit
Randomize

Const BallSize = 50
Const BallMass = 1

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Dim bsTrough, x, bAltSound, bBells

Const cGameName = "phnix_l1"

Dim VarHidden
If Table1.ShowDT = true Then
    VarHidden = 1
    For each x in aReels
        x.Visible = 1
    Next
Else
    VarHidden = 0
    For each x in aReels
        x.Visible = 0
    Next
End If

If B2SOn = true Then VarHidden = 1

LoadVPM "01210000", "S4.VBS", 3.1

'********************
'Standard definitions
'********************

Const UseSolenoids = 1
Const UseLamps = 0
Const UseGI = 0
Const UseSync = 0
Const HandleMech = 0

' Standard Sounds
Const SSolenoidOn = "fx_SolenoidOn"
Const SSolenoidOff = "fx_SolenoidOff"
Const SCoin = "fx_Coin"

'************
' Table init.
'************

Sub table1_Init
    vpmInit me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Phoenix - Williams 1978" & vbNewLine & "VPX8 table by JPSalas v6.0.0"
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = VarHidden
        .Games(cGameName).Settings.Value("rol") = 0   '1= rotated display, 0= normal
        .Games(cGameName).Settings.Value("sound") = 1 '1 enabled rom sound
        '.SetDisplayPosition 0,0,GetPlayerHWnd 'uncomment If you can't see the dmd
        On Error Resume Next
        Controller.SolMask(0) = 0
        vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - Then add the Timer to renable all the solenoids after 2 seconds
        Controller.Run GetPlayerHWnd
        On Error Goto 0
    End With

    ' Nudging
    vpmNudge.TiltSwitch = 1
    vpmNudge.Sensitivity = 3
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingshot)

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0, 37, 0, 0, 0, 0, 0, 0
        .InitKick BallRelease, 90, 4
        .InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .Balls = 1
    End With

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1
    RealTime.Enabled = 1

    ' Turn on Gi
    vpmtimer.addtimer 1500, "GiOn '"
    GIUpdateTimer.Enabled = 1
End Sub

'******************
' RealTime Updates
'******************

Sub RealTime_Timer
    RollingUpdate
End Sub

'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
    If keycode = LeftTiltKey Then Nudge 90, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 8:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If vpmKeyDown(keycode) Then Exit Sub
    If keycode = PlungerKey Then PlaySoundat "fx_PlungerPull", Plunger:Plunger.Pullback
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If vpmKeyUp(keycode) Then Exit Sub
    If keycode = PlungerKey Then PlaySoundAt "fx_plunger", Plunger:Plunger.Fire
End Sub

'*********
' Switches
'*********

' Slings
Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    PlaySoundAt SoundFX("fx_slingshot", DOFContactors), Lemk
    'DOF 101, DOFPulse
    LeftSling004.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 30
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 1:LeftSLing004.Visible = 0:LeftSLing003.Visible = 1:Lemk.RotX = 14
        Case 2:LeftSLing003.Visible = 0:LeftSLing002.Visible = 1:Lemk.RotX = 2
        Case 3:LeftSLing002.Visible = 0:Lemk.RotX = -20:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
    PlaySoundAt SoundFX("fx_slingshot", DOFContactors), Remk
    'DOF 102, DOFPulse
    RightSling004.Visible = 1
    Remk.RotX = 26
    RStep = 0
    vpmTimer.PulseSw 46
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 1:RightSLing004.Visible = 0:RightSLing003.Visible = 1:Remk.RotX = 14
        Case 2:RightSLing003.Visible = 0:RightSLing002.Visible = 1:Remk.RotX = 2
        Case 3:RightSLing002.Visible = 0:Remk.RotX = -20:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

' Scoring rubbers
Dim r1, r2, r3, r4

Sub sw20_Hit:vpmTimer.PulseSw 20:End Sub

Sub sw38_Hit:vpmTimer.PulseSw 38:End Sub

Sub sw39_Hit:vpmTimer.PulseSw 39:r1 = 1:sw39_Timer:End Sub
Sub sw39_Timer
    Select Case r1
        Case 1:Rubber001.Visible = 0:Rubber020.Visible = 1:sw39.TimerEnabled = 1
        Case 2:Rubber020.Visible = 0:Rubber023.Visible = 1
        Case 3:Rubber023.Visible = 0:Rubber001.Visible = 1:sw39.TimerEnabled = 0
    End Select
    r1 = r1 + 1
End Sub

Sub sw40_Hit:vpmTimer.PulseSw 40:r2 = 1:sw40_Timer:End Sub
Sub sw40_Timer
    Select Case r2
        Case 1:Rubber012.Visible = 0:Rubber022.Visible = 1:sw40.TimerEnabled = 1
        Case 2:Rubber022.Visible = 0:Rubber021.Visible = 1
        Case 3:Rubber021.Visible = 0:Rubber012.Visible = 1:sw40.TimerEnabled = 0
    End Select
    r2 = r2 + 1
End Sub

Sub sw22_Hit:vpmTimer.PulseSw 22:r3 = 1:sw22_Timer:End Sub
Sub sw22_Timer
    Select Case r3
        Case 1:Rubber011.Visible = 0:Rubber024.Visible = 1:sw22.TimerEnabled = 1
        Case 2:Rubber024.Visible = 0:Rubber025.Visible = 1
        Case 3:Rubber025.Visible = 0:Rubber011.Visible = 1:sw22.TimerEnabled = 0
    End Select
    r3 = r3 + 1
End Sub

Sub sw24_Hit:vpmTimer.PulseSw 24:r4 = 1:sw24_Timer:End Sub
Sub sw24_Timer
    Select Case r4
        Case 1:Rubber013.Visible = 0:Rubber026.Visible = 1:sw24.TimerEnabled = 1
        Case 2:Rubber026.Visible = 0:Rubber027.Visible = 1
        Case 3:Rubber027.Visible = 0:Rubber013.Visible = 1:sw24.TimerEnabled = 0
    End Select
    r4 = r4 + 1
End Sub

' Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 18:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper1:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 16:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper2:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 17:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper3:End Sub

' Drain & holes
Sub Drain_Hit:PlaysoundAt "fx_drain", Drain:bsTrough.AddBall Me:End Sub

' Rollovers
'inlanes - outlanes
Sub sw32_Hit:Controller.Switch(32) = 1:PlaySoundAt "fx_sensor", sw32:End Sub
Sub sw32_UnHit:Controller.Switch(32) = 0:End Sub

Sub sw31_Hit:Controller.Switch(31) = 1:PlaySoundAt "fx_sensor", sw31:End Sub
Sub sw31_UnHit:Controller.Switch(31) = 0:End Sub

Sub sw47_Hit:Controller.Switch(47) = 1:PlaySoundAt "fx_sensor", sw47:End Sub
Sub sw47_UnHit:Controller.Switch(47) = 0:End Sub

Sub sw48_Hit:Controller.Switch(48) = 1:PlaySoundAt "fx_sensor", sw48:End Sub
Sub sw48_UnHit:Controller.Switch(48) = 0:End Sub

'top
Sub sw9_Hit:Controller.Switch(9) = 1:PlaySoundAt "fx_sensor", sw9:End Sub
Sub sw9_UnHit:Controller.Switch(9) = 0:End Sub

Sub sw10_Hit:Controller.Switch(10) = 1:PlaySoundAt "fx_sensor", sw10:End Sub
Sub sw10_UnHit:Controller.Switch(10) = 0:End Sub

Sub sw11_Hit:Controller.Switch(11) = 1:PlaySoundAt "fx_sensor", sw11:End Sub
Sub sw11_UnHit:Controller.Switch(11) = 0:End Sub

Sub sw12_Hit:Controller.Switch(12) = 1:PlaySoundAt "fx_sensor", sw12:End Sub
Sub sw12_UnHit:Controller.Switch(12) = 0:End Sub

Sub sw13_Hit:Controller.Switch(13) = 1:PlaySoundAt "fx_sensor", sw13:End Sub
Sub sw13_UnHit:Controller.Switch(13) = 0:End Sub


'Spinners
Sub spinner_Spin():vpmTimer.PulseSw 15:PlaySoundAt "fx_spinner", spinner:End Sub

' Droptargets
Sub sw25_Dropped:Controller.Switch(25) = True:PlaySoundAt SoundFX("fx_droptarget", DOFDropTargets), sw25:CheckLeftDrop:End Sub
Sub sw26_Dropped:Controller.Switch(26) = True:PlaySoundAt SoundFX("fx_droptarget", DOFDropTargets), sw26:CheckLeftDrop:End Sub
Sub sw27_Dropped:Controller.Switch(27) = True:PlaySoundAt SoundFX("fx_droptarget", DOFDropTargets), sw27:CheckLeftDrop:End Sub
Sub sw28_Dropped:Controller.Switch(28) = True:PlaySoundAt SoundFX("fx_droptarget", DOFDropTargets), sw28:CheckLeftDrop:End Sub

Sub sw41_Dropped:Controller.Switch(41) = True:PlaySoundAt SoundFX("fx_droptarget", DOFDropTargets), sw41:CheckRightDrop:End Sub
Sub sw42_Dropped:Controller.Switch(42) = True:PlaySoundAt SoundFX("fx_droptarget", DOFDropTargets), sw42:CheckRightDrop:End Sub
Sub sw43_Dropped:Controller.Switch(43) = True:PlaySoundAt SoundFX("fx_droptarget", DOFDropTargets), sw43:CheckRightDrop:End Sub
Sub sw44_Dropped:Controller.Switch(44) = True:PlaySoundAt SoundFX("fx_droptarget", DOFDropTargets), sw44:CheckRightDrop:End Sub

Sub sw33_Dropped:Controller.Switch(33) = True:PlaySoundAt SoundFX("fx_droptarget", DOFDropTargets), sw33:End Sub
Sub sw35_Dropped:Controller.Switch(35) = True:PlaySoundAt SoundFX("fx_droptarget", DOFDropTargets), sw35:End Sub

Sub CheckLeftDrop
    If Controller.Switch(25) and Controller.Switch(26) and Controller.Switch(27) and Controller.Switch(28) Then
        Controller.Switch(29) = True
    End If
End Sub

Sub CheckRightDrop
    If Controller.Switch(41) and Controller.Switch(42) and Controller.Switch(43) and Controller.Switch(44) Then
        Controller.Switch(45) = True
    End If
End Sub

'Targets
Sub sw14_Hit:vpmTimer.PulseSw 14:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw19_Hit:vpmTimer.PulseSw 19:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw21_Hit:vpmTimer.PulseSw 21:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

'*********
'Solenoids
'*********
SolCallback(1) = "bsTrough.SolOut"
SolCallback(2) = "SolLCDrop"
SolCallback(3) = "SolLBDrop"
SolCallback(4) = "SolLTDrop"
SolCallback(5) = "SolRTDrop"
SolCallback(6) = "SolRBDrop"
SolCallback(7) = "SolRCDrop"

SolCallback(14) = "vpmSolSound SoundFX(""fx_Knocker"",DOFKnocker),"
SolCallback(23) = "vpmNudge.SolGameOn"

'solenoid handlers
Sub SolLCDrop(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("fx_droptarget_solenoid", DOFDropTargets), sw33
        Controller.Switch(33) = False
        sw33.IsDropped = False
    End If
End Sub

Sub SolLBDrop(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("fx_droptarget_solenoid", DOFDropTargets), sw26
        Controller.Switch(29) = False
        Controller.Switch(27) = False
        Controller.Switch(28) = False
        sw27.IsDropped = False
        sw28.IsDropped = False
    End If
End Sub

Sub SolLTDrop(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("fx_droptarget_solenoid", DOFDropTargets), sw26
        Controller.Switch(29) = False
        Controller.Switch(25) = False
        Controller.Switch(26) = False
        sw25.IsDropped = False
        sw26.IsDropped = False
    End If
End Sub

Sub SolRTDrop(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("fx_droptarget_solenoid", DOFDropTargets), sw42
        Controller.Switch(45) = False
        Controller.Switch(41) = False
        Controller.Switch(42) = False
        sw41.IsDropped = False
        sw42.IsDropped = False
    End If
End Sub

Sub SolRBDrop(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("fx_droptarget_solenoid", DOFDropTargets), sw43
        Controller.Switch(45) = False
        Controller.Switch(43) = False
        Controller.Switch(44) = False
        sw43.IsDropped = False
        sw44.IsDropped = False
    End If
End Sub

Sub SolRCDrop(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("fx_droptarget_solenoid", DOFDropTargets), sw35
        Controller.Switch(35) = False
        sw35.IsDropped = False
    End If
End Sub

Sub SolSoundAlt(enabled)
    'play once the intro sound after this solenoid is fired
    If Enabled Then
        bAltSound = True
    End If
End sub

Sub Sol10pt(enabled)
    If enabled Then
        If bAltSound Then
            PlaySound "start"
            bAltSound = False
        Else
            PlaySound SoundFX("10pts", DOFChimes)
        End If
    End If
End sub

'*******************
' Flipper Subs Rev3
'*******************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("fx_flipperup", DOFFlippers), LeftFlipper
        LeftFlipper.RotateToEnd
        LeftFlipperOn = 1
    Else
        PlaySoundAt SoundFX("fx_flipperdown", DOFFlippers), LeftFlipper
        LeftFlipper.RotateToStart
        LeftFlipperOn = 0
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("fx_flipperup", DOFFlippers), RightFlipper
        RightFlipper.RotateToEnd
        RightFlipperOn = 1
    Else
        PlaySoundAt SoundFX("fx_flipperdown", DOFFlippers), RightFlipper
        RightFlipper.RotateToStart
        RightFlipperOn = 0
    End If
End Sub

' flippers top animations

Sub LeftFlipper_Animate:LeftFlipperTop.RotZ = LeftFlipper.CurrentAngle:End Sub
Sub RightFlipper_Animate:RightFlipperTop.RotZ = RightFlipper.CurrentAngle:End Sub

' flippers hit Sound

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, Vol(ActiveBall), pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper_Collide(parm)
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
    If LeftFlipper.CurrentAngle >= LeftFlipper.StartAngle - SOSAngle Then LeftFlipper.Strength = FlipperPower * SOSTorque else LeftFlipper.Strength = FlipperPower:End If

    'End Of Stroke Routine : Livecatch and Emply/Full-Charged EOS
    If LeftFlipperOn = 1 Then
        If LeftFlipper.CurrentAngle = LeftFlipper.EndAngle then
            LeftFlipper.EOSTorque = FullStrokeEOS_Torque
            LLiveCatchTimer = LLiveCatchTimer + 1
            If LLiveCatchTimer <LiveCatchSensivity Then
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
    If RightFlipper.CurrentAngle <= RightFlipper.StartAngle + SOSAngle Then RightFlipper.Strength = FlipperPower * SOSTorque else RightFlipper.Strength = FlipperPower:End If

    'End Of Stroke Routine : Livecatch and Emply/Full-Charged EOS
    If RightFlipperOn = 1 Then
        If RightFlipper.CurrentAngle = RightFlipper.EndAngle Then
            RightFlipper.EOSTorque = FullStrokeEOS_Torque
            RLiveCatchTimer = RLiveCatchTimer + 1
            If RLiveCatchTimer <LiveCatchSensivity Then
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
    If Not IsEmpty(chgLamp) Then
        For ii = 0 To UBound(chgLamp)
            LampState(chgLamp(ii, 0) ) = chgLamp(ii, 1) 'keep the real state in an array
            FadingStep(chgLamp(ii, 0) ) = chgLamp(ii, 1)
        Next
    End If
    UpdateLeds
    UpdateLamps
End Sub

Sub UpdateLamps()
    'Lamp 1, li1
    'Lamp 2, li2
    'Lamp 3, li3
    'Lamp 4, li4
    'Lamp 5, li5
    'Lamp 6, li6
    'Lamp 7, li7
    'Lamp 8, li8
    'Lamp 9, li9
    Lamp 10, li10
    Lamp 11, li11
    Lamp 12, li12
    Lampm 13, li13a
    Lamp 13, li13
    'Lamp 14, li14
    'Lamp 15, li15
    'Lamp 16, li16
    Lamp 17, li17
    Lamp 18, li18
    Lamp 19, li19
    Lamp 20, li20
    Lamp 21, li21
    Lamp 22, li22
    Lamp 23, li23
    Lamp 24, li24
    Lamp 25, li25
    Lamp 26, li26
    Lamp 27, li27
    Lamp 28, li28
    Lamp 29, li29
    Lamp 30, li30
    Lamp 31, li31
    Lamp 32, li32
    Lamp 33, li33
    Lamp 34, li34
    Lamp 35, li35
    Lamp 36, li36
    Lamp 37, li37
    'Lamp 38, li38
    Lamp 39, li39
    Lamp 40, li40
    'Lamp 41, li41
    'Lamp 42, li42
    'Lamp 43, li43
    'Lamp 44, li44
    Lamp 45, li45
    Lamp 46, li46
    Lamp 47, li47
    Lamp 48, li48
    'Lamp 49, li49
    Lamp 50, li50 '1 can play
    Lamp 51, li51 '2 can play
    Lamp 52, li52 '3 can play
    Lamp 53, li53 '4 can play
    Lamp 54, li54 'match
    Lamp 55, li55 'ball in play
    Lamp 56, li56 'credits
    Lamp 57, li57 'player 1
    Lamp 58, li58 'player 2
    Lamp 59, li59 'player 3
    Lamp 60, li60 'player 4
    Lamp 61, li61 'tilt
    Lamp 62, li62 'game over
    Lamp 63, li63 'same player shoots (backbox)
    Lamp 64, li64 'highscore
End Sub

' div lamp subs

' Normal Lamp & Flasher subs

Sub InitLamps()
    Dim x
    LampTimer.Interval = 20
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
            If tmp> 0 Then
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
            If tmp> 0 Then
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

'***********
' GI lights
'***********

Dim oldGiState
oldGiState = 0

Sub GiOn 'enciEnde las luces GI
    Dim bulb
    PlaySound "fx_gion"
    For each bulb in aGiLights
        bulb.State = 1
    Next
End Sub

Sub GiOff 'apaga las luces GI
    Dim bulb
    PlaySound "fx_gioff"
    For each bulb in aGiLights
        bulb.State = 0
    Next
End Sub

Sub GIUpdateTimer_Timer
    Dim tmp, obj
    tmp = Getballs
    If UBound(tmp) <> OldGiState Then
        OldGiState = Ubound(tmp)
        If UBound(tmp) = -1 Then '-1 means no balls, 0 is the first captive ball, 1 is the second captive ball...)
            GiOff                ' turn off the gi if no active balls on the table, we could also have used the variable ballsonplayfield.
        Else
            Gion
        End If
    End If
End Sub

'=========================================================
'                    LED Handling
'=========================================================
'ModIfied version of Scapino's LED code for Fathom
'
Dim SixDigitOutput(32)
Dim DisplayPatterns(11)
Dim DigStorage(32)


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

'Assign 6-Digit output to reels
Set SixDigitOutput(0) = P1Digit6
Set SixDigitOutput(1) = P1Digit5
Set SixDigitOutput(2) = P1Digit4
Set SixDigitOutput(3) = P1Digit3
Set SixDigitOutput(4) = P1Digit2
Set SixDigitOutput(5) = P1Digit1

Set SixDigitOutput(6) = P2Digit6
Set SixDigitOutput(7) = P2Digit5
Set SixDigitOutput(8) = P2Digit4
Set SixDigitOutput(9) = P2Digit3
Set SixDigitOutput(10) = P2Digit2
Set SixDigitOutput(11) = P2Digit1

Set SixDigitOutput(12) = P3Digit6
Set SixDigitOutput(13) = P3Digit5
Set SixDigitOutput(14) = P3Digit4
Set SixDigitOutput(15) = P3Digit3
Set SixDigitOutput(16) = P3Digit2
Set SixDigitOutput(17) = P3Digit1

Set SixDigitOutput(18) = P4Digit6
Set SixDigitOutput(19) = P4Digit5
Set SixDigitOutput(20) = P4Digit4
Set SixDigitOutput(21) = P4Digit3
Set SixDigitOutput(22) = P4Digit2
Set SixDigitOutput(23) = P4Digit1

Set SixDigitOutput(24) = CrDigit2
Set SixDigitOutput(25) = CrDigit1
Set SixDigitOutput(26) = BaDigit2
Set SixDigitOutput(27) = BaDigit1

Sub UpdateLeds ' 6-Digit output
    On Error Resume Next
    Dim ChgLED, ii, chg, stat, obj, TempCount, temptext, adj

    ChgLED = Controller.ChangedLEDs(&HFF, &HFFFF) 'hex of binary (display 111111, or first 6 digits)

    If Not IsEmpty(ChgLED) Then
        For ii = 0 To UBound(ChgLED)
            chg = chgLED(ii, 1):stat = chgLED(ii, 2)
            For TempCount = 0 to 10
                If stat = DisplayPatterns(TempCount) Then
                    SixDigitOutput(chgLED(ii, 0) ).SetValue(TempCount)
                    DigStorage(chgLED(ii, 0) ) = TempCount
                End If
                If stat = (DisplayPatterns(TempCount) + 128) Then
                    SixDigitOutput(chgLED(ii, 0) ).SetValue(TempCount)
                    DigStorage(chgLED(ii, 0) ) = TempCount
                End If
            Next
        Next
    End If
End Sub


'************************************
' Diverse Collection Hit Sounds v3.0
'************************************

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
    If tmp> 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = (SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

Function AudioFade(ball) 'only on VPX 10.4 and newer
    Dim tmp
    tmp = ball.y * 2 / TableHeight-1
    If tmp> 0 Then
        AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
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
Const maxvel = 32 'max ball velocity
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
        aBallShadow(b).Height = BOT(b).Z - Ballsize / 2

        If BallVel(BOT(b) )> 1 Then
            If BOT(b).z <30 Then
                ballpitch = Pitch(BOT(b) )
                ballvol = Vol(BOT(b) )
            Else
                ballpitch = Pitch(BOT(b) ) + 50000 'increase the pitch on a ramp
                ballvol = Vol(BOT(b) ) * 3
            End If
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, ballvol, Pan(BOT(b) ), 0, ballpitch, 1, 0, AudioFade(BOT(b) )
        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If

        ' rothbauerw's Dropping Sounds
        If BOT(b).VelZ <-1 and BOT(b).z <55 and BOT(b).z> 27 Then 'height adjust for ball drop sounds
            PlaySound "fx_balldrop", 0, ABS(BOT(b).velz) / 17, Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
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

    ' Use Bells
    x = Table1.Option("Use Bells (needs a restart)", 0, 1, 1, 0, 0, Array("No", "Yes") )
    If x then bBells = True Else bBells = False

    If bBells Then
        SolCallback(9) = "Sol10pt"
        SolCallback(10) = "vpmSolSound SoundFX(""100pts"",DOFChimes),"
        SolCallback(11) = "vpmSolSound SoundFX(""1000pts"",DOFChimes),"
        SolCallback(12) = "vpmSolSound SoundFX(""10000pts"",DOFChimes),"
        SolCallback(13) = "SolSoundAlt"
    End If
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

' Thalamus : Exit in a clean and proper way
Sub Table1_exit
  Controller.Pause = False
  Controller.Stop
End Sub


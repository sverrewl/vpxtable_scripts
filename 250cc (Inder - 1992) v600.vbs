' 250 cc / IPD No. 4089 / Inder 1992 / 4 Players
' VPX8 table by JPSalas version 6.0.0

Option Explicit
Randomize

Const BallSize = 50
Const BallMass = 1

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01550000", "inder.vbs", 3.26

Dim bsTrough, dtBankL, dtBankR, dtBankT, TopBall, x

Const cGameName = "ind250cc"

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
end if

if B2SOn = true then VarHidden = 1

' Standard Sounds
Const SSolenoidOn = "fx_SolenoidOn"
Const SSolenoidOff = "fx_SolenoidOff"
Const SFlipperOn = ""
Const SFlipperOff = ""
Const SCoin = "fx_Coin"

'************
' Table init.
'************

Sub table1_Init
    vpmInit me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "250cc - Inder 1992" & vbNewLine & "VPX table by JPSalas v6.0.0"
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = VarHidden
        .Games(cGameName).Settings.Value("rol") = 0
        .Games(cGameName).Settings.Value("sound") = 1
        '.SetDisplayPosition 0,0,GetPlayerHWnd 'uncomment if you can't see the dmd
        On Error Resume Next
        Controller.SolMask(0) = 0
        vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the Timer to renable all the solenoids after 2 seconds
        Controller.Run GetPlayerHWnd
        On Error Goto 0
    End With

    ' Nudging
    vpmNudge.TiltSwitch = 53
    vpmNudge.Sensitivity = 5
    vpmNudge.TiltObj = Array(Bumper1, Bumper2)

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0, 91, 0, 0, 0, 0, 0, 0
        .InitKick BallRelease, 80, 6
        .InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .Balls = 1
    End With

    ' Drop targets
    Set dtBankL = New cvpmDropTarget
    With dtBankL
        .InitDrop array(sw54, sw64, sw74), array(74, 84, 94)
        .InitSnd SoundFX("fx_droptarget", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
        .CreateEvents "dtBankL"
    End With

    Set dtBankR = New cvpmDropTarget
    With dtBankR
        .InitDrop array(sw55, sw65, sw75), array(75, 85, 95)
        .InitSnd SoundFX("fx_droptarget", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
        .CreateEvents "dtBankR"
    End With

    Set dtBankT = New cvpmDropTarget
    With dtBankT
        .InitDrop array(sw47, sw46, sw45), array(67, 66, 65)
        .InitSnd SoundFX("fx_droptarget", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
        .CreateEvents "dtBankT"
    End With

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1
    RealTime.Enabled = 1

    ' Turn on Gi
    vpmtimer.addtimer 1500, "GiOn '"

    Set TopBall = TopBallC.CreateBall:TopBallC.Kick 180, 1
End Sub

'Fix for top droptargets not registering the hit

Sub sw47_Hit: vpmTimer.PulseSw 67:End Sub
Sub sw46_Hit: vpmTimer.PulseSw 66:End Sub
Sub sw45_Hit: vpmTimer.PulseSw 65:End Sub

Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub
Sub table1_exit:Controller.stop:End Sub

'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If vpmKeyDown(keycode)Then Exit Sub
    If keycode = PlungerKey Then PlaySoundat "fx_PlungerPull", Plunger:Plunger.Pullback
    If keycode = KeyRules Then Rules
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If vpmKeyUp(keycode)Then Exit Sub
    If keycode = PlungerKey Then PlaySoundAt "fx_plunger", Plunger:Plunger.Fire
End Sub

'*********
' Switches
'*********

' Scoring rubbers
Sub RubberBand9_Hit:vpmTimer.PulseSw 92:End Sub
Sub RubberBand7_Hit:vpmTimer.PulseSw 92:End Sub
Sub RubberBand6_Hit:vpmTimer.PulseSw 92:End Sub
Sub RubberBand2_Hit:vpmTimer.PulseSw 92:End Sub
Sub RubberBand3_Hit:vpmTimer.PulseSw 92:End Sub

' Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 96:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper1:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 86:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper2:End Sub

' Drain & holes
Sub Drain_Hit:PlaysoundAt "fx_drain", Drain:bsTrough.AddBall Me:End Sub
Sub sw43_Hit():vpmTimer.PulseSw 63:End Sub
Sub sw42_Hit():vpmTimer.PulseSw 62:End Sub
Sub sw41_Hit():vpmTimer.PulseSw 61:End Sub
Sub sw40_Hit():vpmTimer.PulseSw 60:End Sub
Sub Vuk1_Hit():x = ActiveBall.VelX:Vuk1.destroyball:Vuk2.CreateBall:Vuk2.Kick 90, x:End Sub

' Rollovers
Sub sw44_Hit:Controller.Switch(64) = 1:PlaySoundAt "fx_sensor", sw44:End Sub
Sub sw44_UnHit:Controller.Switch(64) = 0:End Sub

Sub sw73a_Hit:Controller.Switch(93) = 1:PlaySoundAt "fx_sensor", sw73a:End Sub
Sub sw73a_UnHit:Controller.Switch(93) = 0:End Sub

Sub sw73b_Hit:Controller.Switch(93) = 1:PlaySoundAt "fx_sensor", sw73b:End Sub
Sub sw73b_UnHit:Controller.Switch(93) = 0:End Sub

Sub sw70a_Hit:Controller.Switch(90) = 1:PlaySoundAt "fx_sensor", sw70a:DOF 104, DOFOn:End Sub
Sub sw70a_UnHit:Controller.Switch(90) = 0:DOF 104, DOFOff:End Sub

Sub sw70b_Hit:Controller.Switch(90) = 1:PlaySoundAt "fx_sensor", sw70b:DOF 103, DOFOn:End Sub
Sub sw70b_UnHit:Controller.Switch(90) = 0:DOF 103, DOFOff:End Sub

Sub sw70c_Hit:Controller.Switch(90) = 1:PlaySoundAt "fx_sensor", sw70c:End Sub
Sub sw70c_UnHit:Controller.Switch(90) = 0:End Sub

Sub sw32_Hit:Controller.Switch(52) = 1:PlaySoundAt "fx_sensor", sw32:End Sub
Sub sw32_UnHit:Controller.Switch(52) = 0:End Sub

Sub sw50_Hit:Controller.Switch(70) = 1:PlaySoundAt "fx_sensor", sw50:End Sub
Sub sw50_UnHit:Controller.Switch(70) = 0:End Sub

Sub sw51_Hit:Controller.Switch(71) = 1:PlaySoundAt "fx_sensor", sw51:End Sub
Sub sw51_UnHit:Controller.Switch(71) = 0:End Sub

Sub sw52_Hit:Controller.Switch(72) = 1:PlaySoundAt "fx_sensor", sw52:End Sub
Sub sw52_UnHit:Controller.Switch(72) = 0:End Sub

Sub sw60_Hit:Controller.Switch(80) = 1:PlaySoundAt "fx_sensor", sw60:End Sub
Sub sw60_UnHit:Controller.Switch(80) = 0:End Sub

Sub sw61_Hit:Controller.Switch(81) = 1:PlaySoundAt "fx_sensor", sw61:End Sub
Sub sw61_UnHit:Controller.Switch(81) = 0:End Sub

Sub sw62_Hit:Controller.Switch(82) = 1:PlaySoundAt "fx_sensor", sw62:End Sub
Sub sw62_UnHit:Controller.Switch(82) = 0:End Sub

'Targets
Sub sw53a_Hit:vpmTimer.PulseSw 73:PlaySoundAtBall SoundFXDOF("fx_target", 102, DOFPulse, DOFDropTargets):End Sub
Sub sw53b_Hit:vpmTimer.PulseSw 73:PlaySoundAtBall SoundFXDOF("fx_target", 101, DOFPulse, DOFDropTargets):End Sub
Sub sw63_Hit:vpmTimer.PulseSw 83:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub sw63a_Hit:vpmTimer.PulseSw 83:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub sw67_Hit:vpmTimer.PulseSw 87:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub sw77_Hit:vpmTimer.PulseSw 97:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub sw57_Hit:vpmTimer.PulseSw 77:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub sw56_Hit:vpmTimer.PulseSw 76:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub

'Top Ball
Dim TempBall
Sub Post_Hit:Set TempBall = ActiveBall:TopBall.VelX = - TempBall.VelX:TopBall.VelY = TempBall.VelY * 5:End Sub

'*********
'Solenoids
'*********
SolCallback(3) = "vpmSolSound SoundFX(""fx_knocker"",DOFKnocker),"
SolCallback(4) = "bsTrough.SolOut"
SolCallback(5) = "vpmNudge.SolGameOn"
SolCallback(8) = "dtBankT.SolDropUp"

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

Sub LeftFlipper_Animate:LeftFlipperTop.RotZ = LeftFlipper.CurrentAngle: End Sub
Sub RightFlipper_Animate: RightFlipperTop.RotZ = RightFlipper.CurrentAngle: End Sub

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
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

'*****************
'   Gi Effects
'*****************

Dim OldGiState
OldGiState = -1 'start witht he Gi off

Sub SolGi(enabled)
    If enabled Then
        GiOff
    Else
        GiOn
    End If
End Sub

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

Sub GiEffect(enabled)
    If enabled Then
        For each x in aGiLights
            x.Duration 2, 1000, 1
        Next
    End If
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

'************************************
' Game timer for real time updates
'(some tables may use it's own timer)
'************************************

Sub RealTime_Timer
    RollingUpdate
End Sub

Sub gatef_Animate: gatep.roty = gatef.CurrentAngle: End Sub

'**********************************************************
'     JP's Lamp Fading for VPX and Vpinmame v4.0
' FadingStep used for all kind of lamps
' FlashLevel used for modulated flashers
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

Sub UpdateLamps()
    Lamp 1, li1
    Lamp 2, li2
    Lamp 3, li3
    Lamp 6, li6 ' "Match"
    Lamp 7, li7 ' "Game Over"
    Flash 9, li9
    Flash 10, li10
    Flash 11, li11
    Flash 12, li12
    Lamp 17, li17
    Lamp 18, li18
    Lamp 19, li19
    Lamp 20, li20
    Lamp 21, li21
    Lamp 25, li25
    Lamp 26, li26
    Lamp 27, li27
    Lamp 28, li28
    Lamp 31, li31 ' "Ball in Play"
    Lamp 32, li32 ' "Handicap"
    Lamp 33, li33
    Lamp 43, li43
    Lamp 44, li44
    Lamp 45, li45
    Lamp 46, li46
    Lamp 47, li47
    Lamp 48, li48
    Lamp 49, li49
    Lamp 50, li50
    Lamp 51, li51
    Lamp 52, li52
    Lamp 53, li53
    Lamp 54, li54
    Lamp 55, li55
    Lamp 56, li56

    ' 1 is on, 0 is off
    If LampState(14) = 1 Then
        GateF.RotateToStart
    Else
        GateF.RotateToEnd
    End If
    If LampState(15) = 0 Then dtBankL.DropSol_On: LampState(15) = 1
    If LampState(16) = 0 Then dtBankR.DropSol_On: LampState(16) = 1
    'DisplayExtraBall
    If LampState(36) = 1 And LampState(37) = 1 Then Extra.SetValue 0: x1.State = 0: x2.State = 0: x3.State = 0
    If LampState(36) = 0 And LampState(37) = 1 Then Extra.SetValue 1: x1.State = 1: x2.State = 0: x3.State = 0
    If LampState(36) = 1 And LampState(37) = 0 Then Extra.SetValue 2: x1.State = 0: x2.State = 1: x3.State = 0
    If LampState(36) = 0 And LampState(37) = 0 Then Extra.SetValue 3: x1.State = 0: x2.State = 0: x3.State = 3
End Sub

' div lamp subs

' Normal Lamp & Flasher subs

Sub InitLamps()
    Dim x
    LampTimer.Interval = 10
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

'*********************
'LED's based on Eala's
'*********************
Dim Digits(32)
Digits(0) = Array(a00, a02, a05, a06, a04, a01, a03, a07)
Digits(1) = Array(a10, a12, a15, a16, a14, a11, a13)
Digits(2) = Array(a20, a22, a25, a26, a24, a21, a23)
Digits(3) = Array(a30, a32, a35, a36, a34, a31, a33, a37)
Digits(4) = Array(a40, a42, a45, a46, a44, a41, a43)
Digits(5) = Array(a50, a52, a55, a56, a54, a51, a53)
Digits(6) = Array(a60, a62, a65, a66, a64, a61, a63)

Digits(7) = Array(e00, e02, e05, e06, e04, e01, e03, e07)
Digits(8) = Array(e10, e12, e15, e16, e14, e11, e13)
Digits(9) = Array(e20, e22, e25, e26, e24, e21, e23)
Digits(10) = Array(e30, e32, e35, e36, e34, e31, e33, e37)
Digits(11) = Array(e40, e42, e45, e46, e44, e41, e43)
Digits(12) = Array(e50, e52, e55, e56, e54, e51, e53)
Digits(13) = Array(e60, e62, e65, e66, e64, e61, e63)

Digits(14) = Array(b00, b02, b05, b06, b04, b01, b03, b07)
Digits(15) = Array(b10, b12, b15, b16, b14, b11, b13)
Digits(16) = Array(b20, b22, b25, b26, b24, b21, b23)
Digits(17) = Array(b30, b32, b35, b36, b34, b31, b33, b37)
Digits(18) = Array(b40, b42, b45, b46, b44, b41, b43)
Digits(19) = Array(b50, b52, b55, b56, b54, b51, b53)
Digits(20) = Array(b60, b62, b65, b66, b64, b61, b63)

Digits(21) = Array(f00, f02, f05, f06, f04, f01, f03, f07)
Digits(22) = Array(f10, f12, f15, f16, f14, f11, f13)
Digits(23) = Array(f20, f22, f25, f26, f24, f21, f23)
Digits(24) = Array(f30, f32, f35, f36, f34, f31, f33, f37)
Digits(25) = Array(f40, f42, f45, f46, f44, f41, f43)
Digits(26) = Array(f50, f52, f55, f56, f54, f51, f53)
Digits(27) = Array(f60, f62, f65, f66, f64, f61, f63)

Digits(28) = Array(d00, d02, d05, d06, d04, d01, d03)
Digits(29) = Array(d10, d12, d15, d16, d14, d11, d13)

Digits(30) = Array(c00, c02, c05, c06, c04, c01, c03)
Digits(31) = Array(c10, c12, c15, c16, c14, c11, c13)

'********************
'Update LED's display
'********************

Sub UpdateLeds()
    Dim ChgLED, ii, num, chg, stat, obj
    ChgLED = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
        For ii = 0 To UBound(chgLED)
            num = chgLED(ii, 0):chg = chgLED(ii, 1):stat = chgLED(ii, 2)
            For Each obj In Digits(num)
                If chg And 1 Then obj.State = stat And 1
                chg = chg \ 2:stat = stat \ 2
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
'  includes random pitch in PlaySoundAt and PlaySoundAtBall
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
Const lob = 1     'number of locked balls
Const maxvel = 30 'max ball velocity
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
                ballpitch = Pitch(BOT(b)) + 50000 'increase the pitch on a ramp
                ballvol = Vol(BOT(b)) * 10
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

'***********************
' Ball Collision Sound
'***********************

Sub OnBallBallCollision(ball1, ball2, velocity)
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

'***************
' Rules
'***************
Dim Msg(20)
Sub Rules()
    Msg(0) = "250 cc - Inder 1992" &Chr(10)&Chr(10)
    Msg(1) = ""
    Msg(2) = "Side Droptargets: "
    Msg(3) = "   Reds: Lites Extra Ball on the upper right floor"
    Msg(4) = "   Yellows: Lites Extra Ball on captive ball"
    Msg(5) = "   Greens: Lites Special on the upper right floor"
    Msg(6) = "   Left side: Lites Extra Ball on the Outlane"
    Msg(7) = "   Right side: Open right outlane gate"
    Msg(8) = "   Both sides: Lites Special alternate on lower side targets"
    Msg(9) = ""
    Msg(10) = "Upper Droptargets"
    Msg(11) = "   1st target: Lites X5 on the left upper lane"
    Msg(12) = "   2nd target: Lites Extra ball on the left upper lane"
    Msg(13) = "   3rd target: Lites Special on the left upper lane"
    Msg(14) = ""
    Msg(15) = "Lanes 1-2-3"
    Msg(16) = "   Turning off the three 1-2-3 lights lites 500.000 points"
    Msg(17) = "   on the upper right floor and extra ball"
    Msg(18) = ""
    Msg(19) = ""
    Msg(20) = ""
    For X = 1 To 20
        Msg(0) = Msg(0) + Msg(X)&Chr(13)
    Next
    MsgBox Msg(0), , "         Instructions and Rule Card"
End Sub

'Inder 250CC
'added by Inkochnito & Themer
'The functionality of DIP SL3-6 is inverted compared to the documentation
Sub editDips
    Dim vpmDips:Set vpmDips = New cvpmDips
    With vpmDips
        .AddForm 700, 400, "250CC - DIP switches"
        .AddFrame 2, 0, 392, "Credits per coin", &H00000003, Array("1 coin - 1 credit && 1 coin - 4 credits", 0, "2 coins - 1 credit (4 coins - 3 credits) && 1 coin - 3 credits", &H00000003) 'SL1-3&4 (dip 2&1)
        .AddFrame 2, 92, 190, "Handicap value", &H00000300, Array("5,000,000 points", 0, "5,200,000 points", &H00000100, "5,400,000 points", &H00000200, "5,600,000 points", &H00000300)       'SL2-3&4 (dip 10&9)
        .AddFrame 205, 46, 190, "Balls per game", &H00000008, Array("3 balls", 0, "5 balls", &H00000008)                                                                                       'SL1-1 (dip 4)
        .AddFrame 205, 92, 190, "Replay threshold", &H00000030, Array("3,000,000 points", 0, "3,300,000 points", &H00000010, "3,500,000 points", &H00000020, "3,800,000 points", &H00000030)   'SL1-5&6 (dip 5&6)
        .AddFrame 2, 46, 190, "Allow extra ball", &H00200000, Array("no", &H00200000, "yes", 0)                                                                                                'SL3-6 (dip 22)
        .AddLabel 50, 180, 300, 20, "After hitting OK, press F3 to reset game with new settings."
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

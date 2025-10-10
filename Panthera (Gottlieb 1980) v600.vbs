' Panthera - Gottlieb 1980
' IPD No. 1745 / June, 1980 / 4 Players
' VPX8 - version by JPSalas 2025, version 6.0.0

Option Explicit
Randomize

Const cGameName = "panther7" '7 digits
'Const cGameName = "panthera" '6 digits

Const BallSize = 50
Const BallMass = 1

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01560000", "sys80.VBS", 3.37

'Variables
Dim bsTrough, dtL, dtR, dtT, bsLSaucer, x

Const UseSolenoids = 2
Const UseLamps = 0
Const UseGI = 0
Const UseSync = 0 'set it to 1 if the table runs too fast
Const HandleMech = 0

Dim VarHidden
If Table1.ShowDT = True then
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

if B2SOn = true then VarHidden = 1

' Standard Sounds
Const SSolenoidOn = "fx_Solenoidon"
Const SSolenoidOff = "fx_Solenoidoff"
Const SCoin = "fx_Coin"

'Table Init
Sub table1_Init
    vpmInit me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Panthera Gottlieb 1980" & vbNewLine & "VPX table by JPSalas v6.0.0"
        .Games(cGameName).Settings.Value("rol") = 0   '1= rotated display, 0= normal
        .Games(cGameName).Settings.Value("sound") = 1 '1 enabled rom sound
        '.SetDisplayPosition 0,0, GetPlayerHWnd 'restore dmd window position
        .Games(cGameName).Settings.Value("dmd_red") = 0
        .Games(cGameName).Settings.Value("dmd_green") = 223
        .Games(cGameName).Settings.Value("dmd_blue") = 223
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = VarHidden
        On Error Resume Next
        .SolMask(0) = 0
        vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - Then add the timer to renable all the solenoids after 2 seconds
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
        On Error Goto 0
    End With

    'Nudging
    vpmNudge.TiltSwitch = 57
    vpmNudge.Sensitivity = 4
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, LeftSlingshot, RightSlingShot, RightSlingShot2)

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0, 67, 0, 0, 0, 0, 0, 0
        .InitKick BallRelease, 90, 4
        .InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .Balls = 1
    End With

    ' Thalamus, more randomness pls
    'Left Saucer
    Set bsLSaucer = New cvpmBallStack
    With bsLSaucer
        .InitSaucer sw35, 35, 60, 9
        .InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .KickForceVar = 3
        .KickAngleVar = 3
    End With

    'Droptargets
    Set dtL = New cvpmDropTarget
    dtL.InitDrop Array(sw00, sw10, sw20, sw30), Array(00, 10, 20, 30)
    dtL.initsnd SoundFX("fx_droptarget", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
    dtL.CreateEvents("dtL")

    Set dtT = New cvpmDropTarget
    dtT.InitDrop Array(sw01, sw11, sw21, sw31), Array(01, 11, 21, 31)
    dtT.initsnd SoundFX("fx_droptarget", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
    dtT.CreateEvents("dtT")

    Set dtR = New cvpmDropTarget
    dtR.InitDrop Array(sw02, sw12, sw22, sw32), Array(02, 12, 22, 32)
    dtR.initsnd SoundFX("fx_droptarget", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
    dtR.CreateEvents("dtR")

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1
    Realtime.Enabled = 1
End Sub

' Realtime timer

Sub Realtime_Timer
    RollingUpdate
    GIUpdate
End Sub

'**********
' Keys
'**********

Sub Table1_KeyDown(ByVal KeyCode)
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If vpmKeyDown(keycode) Then Exit Sub
    If keycode = PlungerKey Then PlaySoundAt "fx_PlungerPull", Plunger:Plunger.Pullback
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
    If vpmKeyUp(keycode) Then Exit Sub
    If keycode = PlungerKey Then PlaySoundAt "fx_plunger", Plunger:Plunger.Fire
End Sub

' Slings
Dim LStep, RStep, RStep2

Sub LeftSlingShot_Slingshot
    PlaySoundAt SoundFX("fx_slingshot", DOFContactors), Lemk
  DOF 104, DOFPulse
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 34
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
    PlaySoundAt SoundFX("fx_slingshot", DOFContactors),Remk
  DOF 103, DOFPulse
    RightSling4.Visible = 1
    Remk.RotX = 26
    RStep = 0
    vpmTimer.PulseSw 34
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

Sub RightSlingShot2_Slingshot
    PlaySoundAt SoundFX("fx_slingshot", DOFContactors), Remk2
  DOF 105, DOFPulse
    RightSling8.Visible = 1
    Remk2.RotX = 26
    RStep2 = 0
    vpmTimer.PulseSw 34
    RightSlingShot2.TimerEnabled = 1
End Sub

Sub RightSlingShot2_Timer
    Select Case RStep2
        Case 1:RightSLing8.Visible = 0:RightSLing7.Visible = 1:Remk2.RotX = 14
        Case 2:RightSLing7.Visible = 0:RightSLing6.Visible = 1:Remk2.RotX = 2
        Case 3:RightSLing6.Visible = 0:Remk2.RotX = -10:RightSlingShot2.TimerEnabled = 0
    End Select

    RStep2 = RStep2 + 1
End Sub

' Bumpers

Sub Bumper1_Hit:vpmTimer.PulseSw 24:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper1:DOF 102, DOFPulse:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 24:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper2:DOF 101, DOFPulse:End Sub

' Rollovers
Sub sw03_Hit:Controller.Switch(03) = 1:PlaySoundAt "fx_sensor", sw03:End Sub
Sub sw03_UnHit:Controller.Switch(03) = 0:End Sub

Sub sw13_Hit:Controller.Switch(13) = 1:PlaySoundAt "fx_sensor", sw13:End Sub
Sub sw13_UnHit:Controller.Switch(13) = 0:End Sub

Sub sw23_Hit:Controller.Switch(23) = 1:PlaySoundAt "fx_sensor", sw23:End Sub
Sub sw23_UnHit:Controller.Switch(23) = 0:End Sub

Sub sw33_Hit:Controller.Switch(33) = 1:PlaySoundAt "fx_sensor", sw33:End Sub
Sub sw33_UnHit:Controller.Switch(33) = 0:End Sub

Sub sw05_Hit:Controller.Switch(05) = 1:PlaySoundAt "fx_sensor", sw05:End Sub
Sub sw05_UnHit:Controller.Switch(05) = 0:End Sub

Sub sw03b_Hit:Controller.Switch(03) = 1:PlaySoundAt "fx_sensor", sw03b:End Sub
Sub sw03b_UnHit:Controller.Switch(03) = 0:End Sub

Sub sw13b_Hit:Controller.Switch(13) = 1:PlaySoundAt "fx_sensor", sw13b:End Sub
Sub sw13b_UnHit:Controller.Switch(13) = 0:End Sub

Sub sw23b_Hit:Controller.Switch(23) = 1:PlaySoundAt "fx_sensor", sw23b:End Sub
Sub sw23b_UnHit:Controller.Switch(23) = 0:End Sub

Sub sw33b_Hit:Controller.Switch(33) = 1:PlaySoundAt "fx_sensor", sw33b:End Sub
Sub sw33b_UnHit:Controller.Switch(33) = 0:End Sub

Sub sw15_Hit:Controller.Switch(15) = 1:PlaySoundAt "fx_sensor", sw15:DOF 106, DOFOn:End Sub
Sub sw15_UnHit:Controller.Switch(15) = 0:DOF 106, DOFOff:End Sub

Sub sw15b_Hit:Controller.Switch(15) = 1:PlaySoundAt "fx_sensor", sw15b:DOF 107, DOFOn:End Sub
Sub sw15b_UnHit:Controller.Switch(15) = 0:DOF 107, DOFOff:End Sub

Sub sw25b_Hit:Controller.Switch(25) = 1:PlaySoundAt "fx_sensor", sw25b:DOF 106, DOFOn:End Sub
Sub sw25b_UnHit:Controller.Switch(25) = 0:DOF 106, DOFOff:End Sub

Sub sw14a_Hit:Controller.Switch(14) = 1:PlaySoundAt "fx_sensor", sw14a:lsw14a.Duration 2, 1000, 0:End Sub
Sub sw14a_UnHit:Controller.Switch(14) = 0:end sub

Sub sw14b_Hit:Controller.Switch(14) = 1:PlaySoundAt "fx_sensor", sw14b:lsw14b.Duration 2, 1000, 0:End Sub
Sub sw14b_UnHit:Controller.Switch(14) = 0:end sub

'Standup target

Sub sw04_Hit:vpmTimer.PulseSw 4:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub

'rubbers
Sub sw34a_Hit():vpmTimer.PulseSw 34:End Sub
Sub sw34b_Hit():vpmTimer.PulseSw 34:End Sub
Sub sw34c_Hit():vpmTimer.PulseSw 34:End Sub
Sub sw34d_Hit():vpmTimer.PulseSw 34:End Sub
Sub sw34e_Hit():vpmTimer.PulseSw 34:End Sub
Sub sw34f_Hit():vpmTimer.PulseSw 34:End Sub
Sub sw34g_Hit():vpmTimer.PulseSw 34:End Sub

'Spinner
Sub sw25_Spin():vpmTimer.PulseSw 25:PlaySoundAt "fx_spinner",sw25:DOF 108, DOFPulse:End Sub

' Drain & Holes
Sub Drain_Hit:bsTrough.AddBall Me:PlaySoundAt "fx_drain", Drain:End Sub
Sub sw35_Hit:bsLSaucer.AddBall 0:PlaysoundAt "fx_kicker_enter", sw35:End Sub

'****Solenoids

SolCallback(5) = "dtT.soldropup"
SolCallback(2) = "dtL.soldropup"
SolCallback(1) = "dtR.soldropup"
SolCallback(6) = "bsLSaucer.SolOut"
SolCallback(8) = "vpmSolSound SoundFX(""fx_knocker"",DOFKnocker),"
SolCallback(9) = "bsTrough.SolOut"

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

Sub LeftFlipper_Animate: LeftFlipper_Top.RotZ = LeftFlipper.CurrentAngle: End Sub
Sub RightFlipper_Animate: RightFlipper_Top.RotZ = RightFlipper.CurrentAngle: End Sub

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

Sub GiON
  PlaySound"fx_gion"
    For each x in aGiLights
        GiEffect
    Next
End Sub

Sub GiOFF
  PlaySound"fx_gioff"
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

Sub UpdateLamps()
    NfReel 1, l1' "TILT"
    NfReelm 3, l3a' "Shoot again"
    Lamp 3, l3
    Lamp 4, l4
    Lamp 5, l5
    Lamp 6, l6
    Lamp 7, l7
    'Lamp 8, l8
    NfReel 10, l10' "High Score to Date"
    NfReel 11, l11' "Game Over"
    ' Lamp 12, l12
    ' Lamp 13, l13
    ' Lamp 14, l14
    ' Lamp 15, l15
    ' Lamp 16, l16
    ' Lamp 17, l17
    ' Lamp 18, l18
    ' Lamp 19, l19
    ' Lamp 20, l20
    ' Lamp 21, l21
    ' Lamp 22, l22
    ' Lamp 23, l23
    Lamp 24, l24
    Lamp 25, l25
    Lamp 26, l26
    Lamp 27, l27
    Lamp 28, l28
    Lampm 29, l29b
    Lamp 29, l29
    Lampm 30, l30b
    Lamp 30, l30
    Lampm 31, l31b
    Lamp 31, l31
    Lamp 32, l32
    Lamp 33, l33
    Lamp 34, l34
    Lamp 35, l35
    Lamp 36, l36
    Lamp 37, l37
    Lamp 38, l38
    Lamp 39, l39
    Lamp 40, l40
    Lamp 41, l41
    Lamp 42, l42
    Lamp 43, l43
    Lampm 44, l44d
    Lampm 44, l44b
    Lamp 44, l44
    Lampm 45, l45d
    Lampm 45, l45b
    Lamp 45, l45
    Lampm 46, l46d
    Lampm 46, l46b
    Lamp 46, l46
    Lampm 47, l47d
    Lampm 47, l47b
    Lamp 47, l47
    Lampm 48, l48b
    Lamp 48, l48
    Lampm 49, l49b
    Lamp 49, l49
    Lampm 50, l50b
    Lamp 50, l50
    Lampm 51, l51b
    Lamp 51, l51
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

'************************************
'          LEDs Display
'     Based on Scapino's LEDs
'      Gottlieb led Patterns
'************************************

Dim Digits(32)
Dim Patterns(11) 'normal numbers
Dim Patterns2(11) 'numbers with a comma
Dim Patterns3(11) 'the credits and ball in play

Patterns(0) = 0     'empty
Patterns(1) = 63    '0
Patterns(2) = 768   '1
Patterns(3) = 91    '2
Patterns(4) = 79    '3
Patterns(5) = 102   '4
Patterns(6) = 109   '5
Patterns(7) = 124   '6
Patterns(8) = 7     '7
Patterns(9) = 127   '8
Patterns(10) = 103  '9

Patterns2(0) = 128  'empty
Patterns2(1) = 191  '0
Patterns2(2) = 896  '1
Patterns2(3) = 219  '2
Patterns2(4) = 207  '3
Patterns2(5) = 230  '4
Patterns2(6) = 237  '5
Patterns2(7) = 253  '6
Patterns2(8) = 135  '7
Patterns2(9) = 255  '8
Patterns2(10) = 231 '9

Patterns3(0) = 0     'empty
Patterns3(1) = 63    '0
Patterns3(2) = 6     '1
Patterns3(3) = 91    '2
Patterns3(4) = 79    '3
Patterns3(5) = 102   '4
Patterns3(6) = 109   '5
Patterns3(7) = 124   '6
Patterns3(8) = 7     '7
Patterns3(9) = 127   '8
Patterns3(10) = 103  '9

If cGameName = "panther7" AND Table1.ShowDT Then
'Assign 7-digit output to reels
        For each x in aReels6:x.Visible = 0:Next
        For each x in aReels7:x.Visible = 1:Next
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

ElseIF cGameName = "panthera" AND Table1.ShowDT Then
        'Assign 6-digit output to reels
        For each x in aReels6:x.Visible = 1:Next
        For each x in aReels7:x.Visible = 0:Next
        Set Digits(0) = a001
        Set Digits(1) = a002
        Set Digits(2) = a003
        Set Digits(3) = a004
        Set Digits(4) = a005
        Set Digits(5) = a006

        Set Digits(6) = b001
        Set Digits(7) = b002
        Set Digits(8) = b003
        Set Digits(9) = b004
        Set Digits(10) = b005
        Set Digits(11) = b006

        Set Digits(12) = c001
        Set Digits(13) = c002
        Set Digits(14) = c003
        Set Digits(15) = c004
        Set Digits(16) = c005
        Set Digits(17) = c006

        Set Digits(18) = d001
        Set Digits(19) = d002
        Set Digits(20) = d003
        Set Digits(21) = d004
        Set Digits(22) = d005
        Set Digits(23) = d006

        Set Digits(24) = e001
        Set Digits(25) = e002
        Set Digits(26) = e003
        Set Digits(27) = e004
    End If


Sub UpdateLeds
    On Error Resume Next
'    dim oldstat
    Dim ChgLED, ii, jj, chg, stat
    ChgLED = Controller.ChangedLEDs(&HFF, &HFFFF)
    If Not IsEmpty(ChgLED) Then
        For ii = 0 To UBound(ChgLED)
            chg = chgLED(ii, 1):stat = chgLED(ii, 2)
            For jj = 0 to 10
'if oldstat <> STAT then
'debug.print STAT
'oldstat = stat
'end if
                If(stat = Patterns(jj) ) OR (stat = Patterns2(jj) ) OR (stat = Patterns3(jj) ) then Digits(chgLED(ii, 0) ).SetValue jj
            Next
        Next
    End IF
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
                ballvol = Vol(BOT(b)) * 3
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

'Les transistors inutilisés qui servent pour les lampes alimentent la memoire des DropTargets
Dim N1, O1, N2, O2, N3, O3, N4, O4, N5, O5, N6, O6, N7, O7, N8, O8, N9, O9, N10, O10, N11, O11, N12, O12
N1 = 0:O1 = 0:N2 = 0:O2 = 0:N3 = 0:O3 = 0:N4 = 0:O4 = 0:N5 = 0:O5 = 0:N6 = 0:O6 = 0:N7 = 0:O7 = 0:N8 = 0:O8 = 0:N9 = 0:O9 = 0:N10 = 0:O10 = 0:N11 = 0:O11 = 0:N12 = 0:O12 = 0

Set LampCallback = GetRef("UpdateMultipleLamps")

Sub UpdateMultipleLamps
    '1ere bank
    N1 = Controller.Lamp(12)
    If N1 <> O1 Then
        If N1 Then sw00.isdropped = 1:PlaySound SoundFX("fx_droptarget_solenoid", DOFDropTargets), 0, 1, -0.1
        O1 = N1
    End If
    N2 = Controller.Lamp(13)
    If N2 <> O2 Then
        If N2 Then sw10.isdropped = 1:PlaySound SoundFX("fx_droptarget_solenoid", DOFDropTargets), 0, 1, -0.1
        O2 = N2
    End If
    N3 = Controller.Lamp(14)
    If N3 <> O3 Then
        If N3 Then sw20.isdropped = 1:PlaySound SoundFX("fx_droptarget_solenoid", DOFDropTargets), 0, 1, -0.1
        O3 = N3
    End If
    N4 = Controller.Lamp(15)
    If N4 <> O4 Then
        If N4 Then sw30.isdropped = 1:PlaySound SoundFX("fx_droptarget_solenoid", DOFDropTargets), 0, 1, -0.1
        O4 = N4
    End If

    '2eme bank
    N5 = Controller.Lamp(16)
    If N5 <> O5 Then
        If N5 Then sw01.isdropped = 1:PlaySound SoundFX("fx_droptarget_solenoid", DOFDropTargets), 0, 1, -0.05
        O5 = N5
    End If
    N6 = Controller.Lamp(17)
    If N6 <> O6 Then
        If N6 Then sw11.isdropped = 1:PlaySound SoundFX("fx_droptarget_solenoid", DOFDropTargets), 0, 1, -0.05
        O6 = N6
    End If
    N7 = Controller.Lamp(18)
    If N7 <> O7 Then
        If N7 Then sw21.isdropped = 1:PlaySound SoundFX("fx_droptarget_solenoid", DOFDropTargets), 0, 1, -0.05
        O7 = N7
    End If
    N8 = Controller.Lamp(19)
    If N8 <> O8 Then
        If N8 Then sw31.isdropped = 1:PlaySound SoundFX("fx_droptarget_solenoid", DOFDropTargets), 0, 1, -0.05
        O8 = N8
    End If

    '3eme bank
    N9 = Controller.Lamp(20)
    If N9 <> O9 Then
        If N9 Then sw02.isdropped = 1:PlaySound SoundFX("fx_droptarget_solenoid", DOFDropTargets), 0, 1, 0.1
        O9 = N9
    End If
    N10 = Controller.Lamp(21)
    If N10 <> O10 Then
        If N10 Then sw12.isdropped = 1:PlaySound SoundFX("fx_droptarget_solenoid", DOFDropTargets), 0, 1, 0.1
        O10 = N10
    End If
    N11 = Controller.Lamp(22)
    If N11 <> O11 Then
        If N11 Then sw22.isdropped = 1:PlaySound SoundFX("fx_droptarget_solenoid", DOFDropTargets), 0, 1, 0.1
        O11 = N11
    End If
    N12 = Controller.Lamp(23)
    If N12 <> O12 Then
        If N12 Then sw32.isdropped = 1:PlaySound SoundFX("fx_droptarget_solenoid", DOFDropTargets), 0, 1, 0.1
        O12 = N12
    End If
End Sub

'Gottlieb Panthera
'added by Inkochnito
'Added Coins chute by Mike da Spike

Sub editDips
    Dim vpmDips:Set vpmDips = New cvpmDips
    With vpmDips
        .AddForm 700, 400, "Panthera - DIP switches"
        .AddFrame 2,10,190,"Left Coin Chute (Coins/Credit)",&H0000000F,Array("2/1",&H00000009,"1/1",&H00000000,"1/2",&H00000008) 'Dip 1-4
        .AddFrame 2,72,190,"Right Coin Chute (Coins/Credit)",&H000000F0,Array("2/1",&H00000090,"1/1",&H00000000,"1/2",&H00000080) 'Dip 5-8
        .AddFrame 2,132,190,"Center Coin Chute (Coins/Credit)",&H00000F00,Array("2/1",&H00000900,"1/1",&H00000000,"1/2",&H00000800) 'Dip 9-12
        .AddFrame 2,194, 190, "Coin chute 1 and 2 control", &H00002000, Array("seperate", 0, "same", &H00002000)                                                                                                                   'dip 14
        .AddFrame 2,242, 190, "3rd coin chute credits control", &H00001000, Array("no effect", 0, "add 9", &H00001000)                                                                                                          'dip 13
        .AddFrame 207, 10, 190, "Maximum credits", 49152, Array("8 credits", 0, "10 credits", 32768, "15 credits", &H00004000, "25 credits", 49152)                                                                                  'dip 15&16
        .AddFrame 207, 86, 190, "Playfield special", &H00200000, Array("replay", 0, "extra ball", &H00200000)                                                                                                                       'dip 22
        .AddFrame 207, 132, 190, "Hole for special", &H80000000, Array("alternating", 0, "stays lit", &H80000000)                                                                                                                    'dip32
        .AddFrame 207, 178, 190, "High score to date awards", &H00C00000, Array("not displayed and no award", 0, "displayed and no award", &H00800000, "displayed and 2 replays", &H00400000, "displayed and 3 replays", &H00C00000) 'dip 23&24
        .AddChk 2, 300, 190, Array("Sound when scoring?", &H01000000)                                                                                                                                                              'dip 25
        .AddChk 2, 315, 190, Array("Replay button tune?", &H02000000)                                                                                                                                                              'dip 26
        .AddChk 2, 330, 190, Array("Coin switch tune?", &H04000000)                                                                                                                                                                'dip 27
        .AddChk 2, 345, 190, Array("Credits displayed?", &H08000000)                                                                                                                                                               'dip 28
        .AddChk 207, 300, 190, Array("Match feature", &H00020000)                                                                                                                                                                    'dip 18
        .AddChk 207, 315, 190, Array("Attract features", &H20000000)                                                                                                                                                                 'dip 30
        .AddChkExtra 207, 330, 190, Array("Background sound off", &H0100)                                                                                                                                                            'S-board dip 1
        .AddFrameExtra 412, 10, 190, "Attract tune", &H0200, Array("no attract tune", 0, "attract tune played every 6 minutes", &H0200)                                                                                            'S-board dip 2
        .AddFrame 412, 56, 190, "Balls per game", &H00010000, Array("5 balls", 0, "3 balls", &H00010000)                                                                                                                           'dip 17
        .AddFrame 412, 102, 190, "Replay limit", &H00040000, Array("no limit", 0, "one per ball", &H00040000)                                                                                                                      'dip 19
        .AddFrame 412, 148, 190, "Novelty", &H00080000, Array("normal game mode", 0, "50,000 points for special/extra ball", &H0080000)                                                                                            'dip 20
        .AddFrame 412, 194, 190, "Game mode", &H00100000, Array("replay", 0, "extra ball", &H00100000)                                                                                                                             'dip 21
        .AddFrame 412, 240, 190, "Tilt penalty", &H10000000, Array("game over", 0, "ball in play", &H10000000)                                                                                                                     'dip 29
        .AddFrame 412, 286, 190, "Extra ball target adjust", &H40000000, Array("alternating", 0, "stays lit", &H40000000)                                                                                                          'dip 31
        .AddLabel 150, 400, 300, 20, "After hitting OK, press F3 to reset game with new settings."
    End With
    Dim extra
    extra = Controller.Dip(4) + Controller.Dip(5) * 256
    extra = vpmDips.ViewDipsExtra(extra)
    Controller.Dip(4) = extra And 255
    Controller.Dip(5) = (extra And 65280) \ 256 And 255
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

' Thalamus : Exit in a clean and proper way
Sub Table1_exit
  Controller.Pause = False
  Controller.Stop
End Sub


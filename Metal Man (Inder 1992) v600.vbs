' Metal Man / IPD No. 4092 / INDER 1992 / 4 Players
' VPX8 table by jpsalas, version 6.0.0
' thanks to destruk for the vpinmame routines

Option Explicit
Randomize

Const BallSize = 50
Const BallMass = 1

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Dim bsTrough, dtDrop, bsCatapult, bsSubway, x

Const cGameName = "metalman"

Dim VarHidden
If Table1.ShowDT = true then
    For each x in aReels
        x.Visible = 1
    Next
    VarHidden = 1
Else
    For each x in aReels
        x.Visible = 0
    Next
    VarHidden = 0
End If

if B2SOn = true then VarHidden = 1

LoadVPM "01210000", "INDER.VBS", 3.1

'********************
'Standard definitions
'********************

Const UseSolenoids = 2
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
        .SplashInfoLine = "Metal Man .Inder 1992" & vbNewLine & "VPX table by JPSalas v6.0.0"
        .Games(cGameName).Settings.Value("sound") = 1
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = VarHidden
        .Games(cGameName).Settings.Value("rol") = 0
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
    vpmNudge.TiltObj = Array(Bumper1, LeftSlingshot, RightSlingshot)

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0, 91, 0, 0, 0, 0, 0, 0
        .InitKick BallRelease, 90, 4
        .InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .Balls = 1
    End With

    ' Drop targets
    set dtDrop = new cvpmdroptarget
    With dtDrop
        .initdrop array(sw51, sw52, sw53), array(71, 72, 73)
        .InitSnd SoundFX("fx_droptarget", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
        .CreateEvents "dtDrop"
    End With

    'Catapult
    Set bsCatapult = New cvpmBallStack
    With bsCatapult
        .InitSw 0, 80, 0, 0, 0, 0, 0, 0
        .InitKick sw60a, 0, 32
        .InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    End With

    ' Subway
    Set bsSubway = New cvpmBallStack
    With bsSubway
        .InitSw 0, 70, 0, 0, 0, 0, 0, 0
        .InitKick SubwayExit, 233, 20
        .InitExitSnd SoundFX("fx_popper", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .KickForceVar = 2
        .KickAngleVar = 2
    End With

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1
    RealTime.Enabled = 1

    RightSling002.IsDropped = 1:RightSling003.IsDropped = 1:RightSling004.IsDropped = 1
    kickback.Pullback
End Sub

'******************
' RealTime Updates
'******************

Sub RealTime_Timer
    RollingUpdate
End Sub

Sub Gate6_Animate: g1.rotx = 20 - Gate6.Currentangle / 4.5: End Sub
Sub Gate7_Animate: g2.rotx = 20 - Gate7.Currentangle / 4.5: End Sub
Sub Gate8_Animate: g3.rotx = 20 - Gate8.Currentangle / 4.5: End Sub

'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
    If keycode = LeftTiltKey Then Nudge 90, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 8:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If vpmKeyDown(keycode)Then Exit Sub
    If keycode = PlungerKey Then PlaySoundat "fx_PlungerPull", Plunger:Plunger.Pullback
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If vpmKeyUp(keycode)Then Exit Sub
    If keycode = PlungerKey Then PlaySoundAt "fx_plunger", Plunger:Plunger.Fire
End Sub

'*********
' Switches
'*********

' Slings
Dim RStep

Sub RightSlingShot_Slingshot
    PlaySoundAt SoundFX("fx_slingshot", DOFContactors), RSPOS
    'DOF 102, DOFPulse
    RightSling001.IsDropped = 1
    RightSling004.IsDropped = 0
    RStep = 0
    vpmTimer.PulseSw 66
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 1:RightSLing004.IsDropped = 1:RightSLing003.IsDropped = 0
        Case 2:RightSLing003.IsDropped = 1:RightSLing002.IsDropped = 0
        Case 3:RightSLing002.IsDropped = 1:RightSling001.IsDropped = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingshot_Slingshot:vpmTimer.PulseSw 86:PlaySoundAt SoundFX("fx_slingshot", DOFContactors), LSPOS:End Sub

' Scoring rubbers
Sub rsband003_Hit:vpmTimer.PulseSw 92:End Sub
Sub rsband005_Hit:vpmTimer.PulseSw 97:End Sub

' Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 96:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper1:End Sub

' Drain & holes
Sub Drain_Hit:PlaysoundAt "fx_drain", Drain:bsTrough.AddBall Me:End Sub
Sub sw60_Hit:PlaysoundAt "fx_kicker_enter", sw60:bsCatapult.AddBall Me:End Sub

' Rollovers
Sub sw40_Hit:Controller.Switch(60) = 1:PlaySoundAt "fx_sensor", sw40:End Sub
Sub sw40_UnHit:Controller.Switch(60) = 0:End Sub

Sub sw57_Hit:Controller.Switch(77) = 1:PlaySoundAt "fx_sensor", sw57:End Sub
Sub sw57_UnHit:Controller.Switch(77) = 0:End Sub

Sub sw43_Hit:Controller.Switch(63) = 1:PlaySoundAt "fx_sensor", sw43:End Sub
Sub sw43_UnHit:Controller.Switch(63) = 0:End Sub

Sub sw44_Hit:Controller.Switch(64) = 1:PlaySoundAt "fx_sensor", sw44:End Sub
Sub sw44_UnHit:Controller.Switch(64) = 0:End Sub

Sub sw45_Hit:Controller.Switch(65) = 1:PlaySoundAt "fx_sensor", sw45:End Sub
Sub sw45_UnHit:Controller.Switch(65) = 0:End Sub

Sub sw67_Hit:Controller.Switch(87) = 1:PlaySoundAt "fx_sensor", sw67:End Sub
Sub sw67_UnHit:Controller.Switch(87) = 0:End Sub

Sub sw32_Hit:Controller.Switch(52) = 1:PlaySoundAt "fx_sensor", sw32:End Sub
Sub sw32_UnHit:Controller.Switch(52) = 0:End Sub

Sub sw54_Hit:Controller.Switch(74) = 1:PlaySoundAt "fx_sensor", sw54:End Sub
Sub sw54_UnHit:Controller.Switch(74) = 0:End Sub

Sub sw55_Hit:Controller.Switch(75) = 1:PlaySoundAt "fx_sensor", sw55:End Sub
Sub sw55_UnHit:Controller.Switch(75) = 0:End Sub

Sub sw56_Hit:Controller.Switch(76) = 1:PlaySoundAt "fx_sensor", sw56:End Sub
Sub sw56_UnHit:Controller.Switch(76) = 0:End Sub

Sub sw70_Hit:Controller.Switch(90) = 1:PlaySoundAt "fx_sensor", sw70:End Sub
Sub sw70_UnHit:Controller.Switch(90) = 0:End Sub

Sub sw74_Hit:Controller.Switch(94) = 1:PlaySoundAt "fx_sensor", sw74:End Sub
Sub sw74_UnHit:Controller.Switch(94) = 0:End Sub

Sub sw61_Hit:Controller.Switch(81) = 1:PlaySoundAt "fx_sensor", sw61:End Sub
Sub sw61_UnHit:Controller.Switch(81) = 0:End Sub

Sub sw62_Hit:Controller.Switch(82) = 1:PlaySoundAt "fx_sensor", sw62:End Sub
Sub sw62_UnHit:Controller.Switch(82) = 0:End Sub

Sub sw63_Hit:Controller.Switch(83) = 1:PlaySoundAt "fx_sensor", sw63:End Sub
Sub sw63_UnHit:Controller.Switch(83) = 0:End Sub

'Targets
Sub sw41_Hit:vpmTimer.PulseSw 61:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw42_Hit:vpmTimer.PulseSw 62:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw47_Hit:vpmTimer.PulseSw 67:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw64_Hit:vpmTimer.PulseSw 84:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw65_Hit:vpmTimer.PulseSw 85:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

' Subway handling

Sub TopSubway_Hit 'Top Subway Entrance
    PlaySoundAt "fx_kicker_enter", TopSubway
    TopSubway.DestroyBall
    vpmTimer.AddTimer 500, "HandleTopSubway" 'Wait 500ms and then proceed
End Sub

Sub HandleTopSubway(swNo)
    vpmTimer.PulseSwitch(95), 500, "HandleMiddleSubway" 'Wait 500ms then add ball to SubwayExit Solenoid
End Sub

Sub MiddleSubway_Hit
    PlaySoundAt "fx_hole_enter", MiddleSubway
    MiddleSubway.DestroyBall
    vpmTimer.PulseSwitch(95), 500, "HandleMiddleSubway" 'Wait 500ms then add ball to SubwayExit Solenoid
End Sub

Sub HandleMiddleSubway(swNo):bsSubway.AddBall 0:End Sub

Sub SubwayExit2_Hit():PlaySoundAt "fx_hole_enter", SubwayExit:bsSubway.AddBall Me:End Sub

' VUK
Sub Vukdown_Hit
    Vukdown.DestroyBall
    Vukup.CreateBall
    Vukup.Kick 180, 8
End Sub

'*********
'Solenoids
'*********
'Manual Solenoids
'1=Taca (in cabinet) Knocker *5
'2=Ball Release *4
'3=Pop Bumper *6
'4=Top Left Slingshot *7
'5=Drop Target Reset *8
'6=Subway Exit *9
'7=Catapult *10
'8=Right Slingshot *11
'9=Kickback *12

SolCallback(4) = "bsTrough.SolOut"
SolCallback(3) = "vpmSolSound SoundFX(""fx_Knocker"",DOFKnocker),"
'SolCallback(6)="vpmSolSound ""bumper"","
'SolCallback(7)="vpmSolSound ""Slingshot"","
SolCallback(8) = "dtDrop.SolDropUp"
SolCallback(9) = "bsSubway.SolOut"
SolCallback(10) = "bsCatapult.SolOut"
'SolCallback(11)="vpmSolSound ""Slingshot"","
SolCallback(12) = "vpmSolAutoPlunger Kickback,10,"

'Flahsers
'SolCallback(17) = "SetLamp 117,"

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

' flippers top animations

Sub LeftFlipper_Animate: LeftFlipperTop.RotZ = LeftFlipper.CurrentAngle: End Sub
Sub RightFlipper_Animate: RightFlipperTop.RotZ = RightFlipper.CurrentAngle: End Sub

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
    Lamp 1, l1
    Lamp 2, l2
    Lamp 3, l3
    Lamp 4, l4
    Lamp 6, l6      'Backglass Match
    Lamp 7, l7      'Backglass Game Over
    Lampm 11, gi018 'right Gi
    Lampm 11, gi019
    Lampm 11, gi017
    Lampm 11, gi021
    Lampm 11, gi020
    Lampm 11, gi022
    Lampm 11, gi024
    Lampm 11, gi025
    Lampm 11, gi026
    Lampm 11, gi027
    Lamp 11, gi036
    Lampm 12, gi016 'left Gi
    Lampm 12, gi028
    Lampm 12, gi030
    Lampm 12, gi029
    Lampm 12, gi031
    Lampm 12, gi035
    Lampm 12, gi034
    Lampm 12, gi033
    Lamp 12, gi001
    Lamp 17, l17
    Lamp 18, l18
    Lamp 19, l19
    Lamp 20, l20
    Lamp 21, l21
    Lamp 22, l22
    Lamp 23, l23
    Lamp 25, l25
    Lamp 26, l26
    Lamp 27, l27
    Lamp 28, l28
    Lamp 29, l29
    Lamp 30, l30
    Lamp 31, l31 'Backglass Ball in Play
    Lamp 33, l33
    Lamp 34, l34
    Lamp 35, l35
    Lamp 36, l36 'Backglass Extra Ball 1
    Lamp 37, l37 'Backglass Extra Ball 2
    Lamp 38, l38
    Lamp 41, l41
    Lamp 42, l42
    Lamp 43, l43
    Lamp 44, l44
    Lamp 45, l45
    Lamp 46, l46
    Lampm 49, gi007 'top right gi
    Lampm 49, gi008
    Lampm 49, gi010
    Lampm 49, gi002
    Lampm 49, gi003
    Lampm 49, gi004
    Lampm 49, gi006
    Lamp 49, gi015
    Lampm 50, l50a 'bumper & Gi top left
    Lampm 50, gi013
    Lampm 50, gi014
    Lampm 50, gi012
    Lampm 50, gi011
    Lampm 50, gi005
    Lampm 50, gi009
    Lampm 50, gi023
    Lamp 50, l50
    Lamp 51, l51
    Lamp 52, l52
    Lamp 53, l53
    Lamp 54, l54
    Lamp 55, l55
    Lamp 56, l56
    Lamp 65, l65
    Lamp 66, l66
    Lamp 67, l67
    Lamp 68, l68
    Lamp 69, l69
    Lamp 70, l70
    Lamp 73, l73
    Lamp 74, l74
    Lamp 75, l75
    Lamp 76, l76
    Lamp 77, l77
    Lamp 78, l78
    Lamp 81, l81
    Lamp 82, l82
    Lamp 83, l83
    Lamp 84, l84
    Lamp 85, l85
    Lamp 89, l89
    Lamp 90, l90
    Lamp 91, l91
    Lamp 92, l92
    'flashers
    Lamp 109, l109
    Lamp 110, l110
    Lampm 111, l111a
    Lamp 111, l111
    Lampm 112, l112a
    Lamp 112, l112
    Lamp 150, l150
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
'    LEDs Display by TAB/destruk
'************************************

Dim Digits(32)
Digits(0) = Array(B1, B2, B3, B4, B5, B6, B7)
Digits(1) = Array(B8, B9, B10, B11, B12, B13, B14)
Digits(2) = Array(B15, B16, B17, B18, B19, B20, B21)
Digits(3) = Array(B22, B23, B24, B25, B26, B27, B28)
Digits(4) = Array(B29, B30, B31, B32, B33, B34, B35)
Digits(5) = Array(B36, B37, B38, B39, B40, B41, B42)
Digits(6) = Array(B43, B44, B45, B46, B47, B48, B49)
Digits(7) = Array(B50, B51, B52, B53, B54, B55, B56)
Digits(8) = Array(B57, B58, B59, B60, B61, B62, B63)
Digits(9) = Array(B64, B65, B66, B67, B68, B69, B70)
Digits(10) = Array(B71, B72, B73, B74, B75, B76, B77)
Digits(11) = Array(B78, B79, B80, B81, B82, B83, B84)
Digits(12) = Array(B85, B86, B87, B88, B89, B90, B91)
Digits(13) = Array(B92, B93, B94, B95, B96, B97, B98)
Digits(14) = Array(B99, B100, B101, B102, B103, B104, B105)
Digits(15) = Array(B106, B107, B108, B109, B110, B111, B112)
Digits(16) = Array(B113, B114, B115, B116, B117, B118, B119)
Digits(17) = Array(B120, B121, B122, B123, B124, B125, B126)
Digits(18) = Array(B127, B128, B129, B130, B131, B132, B133)
Digits(19) = Array(B134, B135, B136, B137, B138, B139, B140)
Digits(20) = Array(B141, B142, B143, B144, B145, B146, B147)
Digits(21) = Array(B148, B149, B150, B151, B152, B153, B154)
Digits(22) = Array(B155, B156, B157, B158, B159, B160, B161)
Digits(23) = Array(B162, B163, B164, B165, B166, B167, B168)
Digits(24) = Array(B169, B170, B171, B172, B173, B174, B175)
Digits(25) = Array(B176, B177, B178, B179, B180, B181, B182)
Digits(26) = Array(B183, B184, B185, B186, B187, B188, B189)
Digits(27) = Array(B190, B191, B192, B193, B194, B195, B196)
Digits(28) = Array(B197, B198, B199, B200, B201, B202, B203)
Digits(29) = Array(B204, B205, B206, B207, B208, B209, B210)
Digits(30) = Array(B211, B212, B213, B214, B215, B216, B217)
Digits(31) = Array(B218, B219, B220, B221, B222, B223, B224)
Digits(32) = Array(B225, B226, B227, B228, B229, B230, B231)

Sub UpdateLeds
    Dim ChgLED, ii, num, chg, stat, obj
    ChgLED = Controller.ChangedLEDs(&H0000003f, &Hffffffff)
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

'Inder 250CC
'added by Inkochnito
Sub editDips
    Dim vpmDips:Set vpmDips = New cvpmDips
    With vpmDips
        .AddForm 700, 400, "250CC - DIP switches"
        .AddFrame 2, 0, 392, "Credits per coin", &H00000003, Array("1 coin - 1 credit && 1 coin - 4 credits", 0, "2 coins - 1 credit (4 coins - 3 credits) && 1 coin - 3 credits", &H00000003) 'SL1-3&4 (dip 1&2)
        .AddFrame 2, 92, 190, "Handicap value", &H00000300, Array("5,000,000 points", 0, "5,200,000 points", &H00000100, "5,400,000 points", &H00000200, "5,600,000 points", &H00000300)       'SL2-4&3(dip 9&10)
        .AddFrame 205, 46, 190, "Balls per game", &H00000800, Array("3 balls", 0, "5 balls", &H00000800)                  'SL1-1 (dip 4)                                                                                     'SL1-1 (dip 4)
        .AddFrame 205, 92, 190, "Replay threshold", &H000000C0, Array("3,000,000 points", 0, "3,300,000 points", &H00000080, "3,500,000 points", &H00000040, "3,800,000 points", &H000000C0)   'SL1-6&5 (dip 7&8)
        .AddFrame 2, 46, 190, "Allow extra ball", &H00400000, Array("yes", 0, "no", &H00400000)                                                                                                'SL3-6 (dip 23)
        .AddLabel 50, 180, 300, 20, "After hitting OK, press F3 to reset game with new settings."
        .ViewDips
    End With
End Sub
Set vpmShowDips = GetRef("editDips")

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
Const maxvel = 45 'max ball velocity
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
                ballpitch = Pitch(BOT(b)) + 25000 'increase the pitch on a ramp
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

'***********************
' Ball Collision Sound
'***********************

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

